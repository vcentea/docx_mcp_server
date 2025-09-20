#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import sys
import os
from copy import deepcopy


def _collect_ids(node):
    """Collect all string IDs present in the given JSON structure."""
    ids = set()

    def walk(x):
        if isinstance(x, dict):
            ident = x.get("id")
            if isinstance(ident, str) and ident:
                ids.add(ident)
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for it in x:
                walk(it)

    walk(node)
    return ids


def _sanitize_and_assign_id(dst_node, existing_ids, id_counters):
    """Ensure an inserted node has a unique, non-conflicting ID.

    Behavior (simple and robust):
    - If no id provided, assign based on type with reserved insert prefixes:
        paragraph -> i-p-<n>, table -> i-t-<n>, default -> i-<n>
    - If an explicit id uses exported prefixes (p-/t-), rewrite to reserved insert prefixes
    - Guarantee uniqueness by incrementing a numeric suffix until free
    - Record assignments via id_counters per prefix to avoid excessive scanning
    """
    if not isinstance(dst_node, dict):
        return

    node_type = dst_node.get("type")
    if node_type == "paragraph":
        preferred_prefix = "i-p-"
    elif node_type == "table":
        preferred_prefix = "i-t-"
    else:
        preferred_prefix = "i-"

    provided = dst_node.get("id")
    # Rewrite reserved exported prefixes if provided
    if isinstance(provided, str) and provided:
        if provided.startswith("p-") or provided.startswith("t-"):
            print(f"Rewriting inserted id '{provided}' to reserved insert prefix.")
            provided = None

    # Choose a base candidate
    candidate = provided if (isinstance(provided, str) and provided) else None
    if not candidate or candidate in existing_ids:
        # Generate a unique candidate using counters
        idx = id_counters.get(preferred_prefix, 1)
        while True:
            cand = f"{preferred_prefix}{idx}"
            if cand not in existing_ids:
                candidate = cand
                id_counters[preferred_prefix] = idx + 1
                break
            idx += 1

    # Assign and record
    dst_node["id"] = candidate
    existing_ids.add(candidate)

def _first_run_rpr(node):
    """Return the first non-empty rPr found in a paragraph-like node's content."""
    if not isinstance(node, dict):
        return None
    for it in node.get("content", []) or []:
        if isinstance(it, dict) and it.get("type") == "run":
            rpr = it.get("rPr")
            if rpr:
                return rpr
    return None


def _ensure_paragraph_formatting(dst_node, src_node):
    """If dst_node (paragraph) lacks pPr/rPr, copy from src_node.

    - Copies pPr if missing or empty
    - For each run missing rPr, applies rPr from the first run in src_node (if present)
    """
    if not isinstance(dst_node, dict) or not isinstance(src_node, dict):
        return False
    if dst_node.get("type") != "paragraph" or src_node.get("type") != "paragraph":
        return False

    changed = False
    if not dst_node.get("pPr"):
        src_ppr = src_node.get("pPr")
        if src_ppr:
            dst_node["pPr"] = deepcopy(src_ppr)
            changed = True

    base_rpr = _first_run_rpr(src_node)
    if base_rpr:
        for it in dst_node.get("content", []) or []:
            if isinstance(it, dict) and it.get("type") == "run":
                if not it.get("rPr"):
                    it["rPr"] = deepcopy(base_rpr)
                    changed = True
    return changed

def find_element_and_parent(data, element_id):
    """
    Recursively searches for an element by its ID in the nested structure.
    Returns the element, its parent list, and its index in that list.
    """
    if isinstance(data, dict):
        # Check if the current dict is the target
        if data.get("id") == element_id:
            return data, None, None # Found at top level, no parent list

        for key, value in data.items():
            found, parent, index = find_element_and_parent(value, element_id)
            if found:
                # If the direct parent is a dictionary, it means the element is a value of some key.
                # We need to go up to find the list that contains this dictionary.
                # This logic assumes our target nodes are always in a list (like 'body' or 'content').
                return found, parent, index

    elif isinstance(data, list):
        for i, item in enumerate(data):
            if isinstance(item, dict) and item.get("id") == element_id:
                return item, data, i # Found it! Return the item, its list, and index.
            
            # Recurse into the item
            found, parent, index = find_element_and_parent(item, element_id)
            if found:
                return found, parent, index

    return None, None, None

def _set_deep_property(element, path_keys, value):
    """Set a nested property using dot notation keys."""
    current = element
    for key in path_keys[:-1]:
        if not isinstance(current, dict):
            return False
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    
    final_key = path_keys[-1]
    if isinstance(current, dict):
        current[final_key] = value
        return True
    return False


def edit_element_content(element, element_id, path, value):
    """Replace or modify content/properties of a document element."""
    if path is None or value is None:
        print(f"Skipping invalid 'replace' op: missing 'path' or 'value'. Element ID: {element_id}")
        return False

    # Support dotted paths (e.g., "pPr.jc" or "pPr.numPr")
    path_keys = path.split('.')
    
    success = _set_deep_property(element, path_keys, value)
    if success:
        print(f"Replaced '{path}' on element '{element_id}'.")
    else:
        print(f"Warning: Failed to set path '{path}' on element '{element_id}'.")
    
    return success


def delete_section(parent_list, index, element_id):
    """Remove an element/section from the document."""
    if parent_list is not None and index is not None:
        del parent_list[index]
        print(f"Deleted element '{element_id}'.")
        return True
    else:
        print(f"Warning: Cannot delete element '{element_id}' because its parent could not be determined.")
        return False


def add_element_after(parent_list, index, element_id, nodes, existing_ids, id_counters, data, neighbor_element):
    """Insert new elements after the specified element."""
    return _insert_elements(parent_list, index, element_id, nodes, existing_ids, id_counters, data, neighbor_element, after=True)


def add_element_before(parent_list, index, element_id, nodes, existing_ids, id_counters, data, neighbor_element):
    """Insert new elements before the specified element."""
    return _insert_elements(parent_list, index, element_id, nodes, existing_ids, id_counters, data, neighbor_element, after=False)


def _insert_elements(parent_list, index, element_id, nodes, existing_ids, id_counters, data, neighbor_element, after=True):
    """Internal function to handle element insertion logic."""
    if not isinstance(nodes, list):
        print(f"Skipping invalid insert op: 'nodes' must be a list. Element ID: {element_id}")
        return False

    if parent_list is None or index is None:
        print(f"Warning: Cannot insert elements near '{element_id}' because its parent list could not be determined.")
        return False

    # Insert with minimal auto-formatting nudges
    inserted_count = 0
    for i, raw_node in enumerate(nodes):
        node = deepcopy(raw_node)

        # Ensure safe, unique id for the inserted node
        _sanitize_and_assign_id(node, existing_ids, id_counters)

        # Optional explicit formatting reference: node["formatFrom"]
        fmt_from_id = node.get("formatFrom")
        fmt_from_node = None
        if fmt_from_id:
            cand, _, _ = find_element_and_parent(data, fmt_from_id)
            if cand:
                fmt_from_node = cand
            else:
                print(f"Warning: formatFrom id '{fmt_from_id}' not found; falling back to neighbor.")

        src_for_format = fmt_from_node or neighbor_element
        if src_for_format is not None:
            changed = _ensure_paragraph_formatting(node, src_for_format)
            if changed:
                src_id = src_for_format.get("id") if isinstance(src_for_format, dict) else None
                print(f"Applied formatting from '{src_id or 'neighbor'}' to inserted node '{node.get('id')}'.")

        if after:
            parent_list.insert(index + 1 + inserted_count, node)
        else:  # before
            parent_list.insert(index + inserted_count, node)
        inserted_count += 1

    position = "after" if after else "before"
    print(f"Inserted {inserted_count} node(s) {position} element '{element_id}'.")
    return True


def append_elements_at_end(data, nodes, existing_ids, id_counters):
    """Append elements at the end of the document body as fallback."""
    if not isinstance(nodes, list):
        print("Skipping invalid append op: 'nodes' must be a list.")
        return False
    
    # Get the document body (main content container)
    body = data.get("body", [])
    
    # Insert elements before the last element if it's sectionProps, otherwise append at very end
    insertion_index = len(body)
    if body and body[-1].get("type") == "sectionProps":
        insertion_index = len(body) - 1
        print("Inserting before sectionProps at end of document.")
    else:
        print("Appending at very end of document.")
    
    inserted_count = 0
    for raw_node in nodes:
        node = deepcopy(raw_node)
        
        # Ensure safe, unique id for the inserted node
        _sanitize_and_assign_id(node, existing_ids, id_counters)
        
        # Apply basic formatting for consistency (minimal formatting)
        if node.get("type") == "paragraph" and not node.get("pPr"):
            # Add basic paragraph formatting if none exists
            node["pPr"] = {}
        
        body.insert(insertion_index + inserted_count, node)
        inserted_count += 1
        print(f"Appended element '{node.get('id')}' at end of document.")
    
    print(f"Successfully appended {inserted_count} element(s) at end of document.")
    return True


def apply_patch(data, patch):
    """Applies a list of patch operations to the data using specific operation functions."""
    patched_data = deepcopy(data)
    # Track existing ids and simple counters for inserted-id generation
    existing_ids = _collect_ids(patched_data)
    id_counters = {}

    for op in patch.get("ops", []):
        op_type = op.get("op")
        element_id = op.get("id")

        if not op_type or not element_id:
            print(f"Skipping invalid operation: {op}")
            continue

        element, parent_list, index = find_element_and_parent(patched_data, element_id)

        if not element:
            # SMART FALLBACK: If reference element not found for insert operations,
            # fall back to appending at the end of the document body
            if op_type in ["insertAfter", "insertBefore"]:
                print(f"Reference element '{element_id}' not found. Falling back to append at end of document.")
                append_elements_at_end(patched_data, op.get("nodes"), existing_ids, id_counters)
            else:
                print(f"Warning: Element with ID '{element_id}' not found for operation '{op_type}'.")
            continue

        # Call appropriate operation function based on type
        if op_type == "replace":
            edit_element_content(element, element_id, op.get("path"), op.get("value"))
        
        elif op_type == "delete":
            delete_section(parent_list, index, element_id)
        
        elif op_type == "insertAfter":
            add_element_after(parent_list, index, element_id, op.get("nodes"), existing_ids, id_counters, patched_data, element)
        
        elif op_type == "insertBefore":
            add_element_before(parent_list, index, element_id, op.get("nodes"), existing_ids, id_counters, patched_data, element)
        
        else:
            print(f"Warning: Unknown operation type '{op_type}' for element '{element_id}'.")

    return patched_data


def main():
    if len(sys.argv) != 4:
        print("Usage: python patch_json.py <source.json> <patch.json> <output.json>", file=sys.stderr)
        sys.exit(2)

    source_path = sys.argv[1]
    patch_path = sys.argv[2]
    output_path = sys.argv[3]

    for path in [source_path, patch_path]:
        if not os.path.isfile(path):
            print(f"Error: File not found: {path}", file=sys.stderr)
            sys.exit(2)
            
    with open(source_path, 'r', encoding='utf-8') as f:
        source_data = json.load(f)

    with open(patch_path, 'r', encoding='utf-8') as f:
        patch_data = json.load(f)

    # Apply the patch
    output_data = apply_patch(source_data, patch_data)

    # Write the output
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f"\nSuccessfully applied patch and wrote to {output_path}")


if __name__ == "__main__":
    main()
