#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Enhanced MCP server exposing tools to:
 - Convert a .docx file to JSON (using docx2json_orchestrate.py)
 - Delete document elements by ID (single or multiple)
 - Add document elements (single or multiple, before/after/end)
 - Edit document element content and properties (single or multiple)
 - Support combined operations (edit + add + delete in single call)
 - Return updated document JSON with ID mappings after changes
 - Automatic versioning for all changes

Run: python mcp_server.py
This uses stdio transport for MCP clients.
"""

from __future__ import annotations

import json
import os
import tempfile
from typing import Any, Dict, Optional, Union, List, Literal
from collections import deque
from datetime import datetime
from pydantic import BaseModel, Field

from mcp.server.fastmcp import FastMCP
import contextlib
import io
import re

# Local modules - refactored docx manipulation system
import docx2json_orchestrate
import json2docx_structured as j2d
import patch_json as pjson


server = FastMCP("contract-docx-mcp")


def _ensure_abs(path: str) -> str:
    return os.path.abspath(os.path.expanduser(path))


def _strip_version_suffix(base_noext: str) -> str:
    """Strip a trailing .v<digits> suffix from a path without extension."""
    return re.sub(r"\.v\d+$", "", base_noext)


def _get_next_version_number(base_path: str) -> int:
    """Find the next available version number for a file."""
    dir_name = os.path.dirname(base_path)
    base_name = os.path.basename(base_path)
    
    if not os.path.exists(dir_name):
        return 1
        
    existing_versions = []
    for file in os.listdir(dir_name):
        if file.startswith(base_name + ".v") and not file.endswith(".tmp"):
            parts = file[len(base_name):].split('.')
            if len(parts) >= 2 and parts[1].startswith('v') and parts[1][1:].isdigit():
                version_num = int(parts[1][1:])
                existing_versions.append(version_num)
    
    return max(existing_versions) + 1 if existing_versions else 1


def _derive_versioned_path(base_path: str, extension: str, version: Optional[int] = None) -> str:
    """Generate a versioned path for any file type."""
    base_path = _ensure_abs(base_path)
    base, _ = os.path.splitext(base_path)
    
    if version is None:
        version = _get_next_version_number(base)
    
    base = _strip_version_suffix(base)
    return f"{base}.v{version}{extension}"


def _prepare_element_with_formatting(
    element: Union[str, Dict[str, Any]], 
    text_properties_ref: Optional[str] = None,
    available_properties: Optional[Dict[str, str]] = None
) -> Dict[str, Any]:
    """Convert string or element to properly formatted element with text properties."""
    
    # Handle simple string - wrap in paragraph with specified formatting
    if isinstance(element, str):
        # Determine which text properties to use
        if text_properties_ref:
            props_ref = text_properties_ref
        elif available_properties and "default_text_format" in available_properties:
            props_ref = "default_text_format"
        else:
            # Use the first available property or None
            props_ref = next(iter(available_properties.keys())) if available_properties else None
        
        run_content = {"type": "run", "text": element}
        if props_ref:
            run_content["textPropsRef"] = props_ref
            
        return {
            "type": "paragraph",
            "content": [run_content]
        }
    
    # Handle dictionary element - ensure runs have proper textPropsRef
    if isinstance(element, dict) and element.get("type") == "paragraph":
        # Apply text_properties_ref to runs that don't already have one
        if text_properties_ref and "content" in element:
            updated_content = []
            for item in element["content"]:
                if isinstance(item, dict) and item.get("type") == "run" and "textPropsRef" not in item:
                    item = item.copy()
                    item["textPropsRef"] = text_properties_ref
                updated_content.append(item)
            
            element = element.copy()
            element["content"] = updated_content
    
    return element


def _generate_id_mapping(original_json: Dict[str, Any], updated_json: Dict[str, Any]) -> Dict[str, str]:
    """Generate mapping of old IDs to new IDs after changes."""
    old_ids = set()
    new_ids = set()
    
    def collect_ids(node, id_set):
        if isinstance(node, dict):
            if "id" in node and isinstance(node["id"], str):
                id_set.add(node["id"])
            for value in node.values():
                collect_ids(value, id_set)
        elif isinstance(node, list):
            for item in node:
                collect_ids(item, id_set)
    
    collect_ids(original_json.get("body", []), old_ids)
    collect_ids(updated_json.get("body", []), new_ids)
    
    # Simple mapping - this could be enhanced with more sophisticated matching
    mapping = {}
    for old_id in old_ids:
        if old_id not in new_ids:
            mapping[old_id] = "DELETED"
    
    for new_id in new_ids:
        if new_id not in old_ids:
            mapping["NEW"] = mapping.get("NEW", [])
            if not isinstance(mapping["NEW"], list):
                mapping["NEW"] = [mapping["NEW"]]
            mapping["NEW"].append(new_id)
    
    return mapping


@server.tool(
    title="Get Document as JSON",
    description=(
        "Converts a DOCX document to a comprehensive structured JSON using the orchestrated system.\n\n"
        "This system provides enhanced document processing with:\n"
        "- Semantic text property identification and deduplication\n"
        "- Automatic ID assignment for paragraphs and tables\n" 
        "- Text formatting registry with meaningful names\n"
        "- Optimized run merging for cleaner structure\n\n"
        "ALWAYS call this before making changes to understand current document structure.\n"
        "Args: docx_path (str), output_json_path (optional), return_json (optional)."
    ),
)
def get_document_as_json(docx_path: str, output_json_path: Optional[str] = None, return_json: bool = True) -> Dict[str, Any]:
    """Convert a .docx file to our comprehensive structured JSON using the orchestrated system."""
    docx_path = _ensure_abs(docx_path)
    _log_tool_call("get_document_as_json", {"docx_path": docx_path, "output_json_path": output_json_path, "return_json": return_json})
    
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(f"DOCX not found: {docx_path}")
    if not docx_path.lower().endswith(".docx"):
        raise ValueError("Input must be a .docx file")

    if output_json_path is None:
        base, _ = os.path.splitext(docx_path)
        output_json_path = base + ".export.json"
    output_json_path = _ensure_abs(output_json_path)
    
    docx2json_orchestrate.build_docx_json(docx_path, output_json_path)
    
    if return_json:
        with open(output_json_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"json_path": output_json_path}


@server.tool(
    title="Get Document Text Properties",
    description=(
        "Extracts only the text properties registry from a DOCX document.\n\n"
        "Returns comprehensive text formatting information with semantic names.\n"
        "Much lighter than full document JSON when you only need formatting information.\n\n"
        "Args: docx_path (str), return_json (optional)."
    ),
)
def get_document_text_properties(docx_path: str, return_json: bool = True) -> Dict[str, Any]:
    """Extract only the text properties registry from a .docx file."""
    docx_path = _ensure_abs(docx_path)
    _log_tool_call("get_document_text_properties", {"docx_path": docx_path, "return_json": return_json})
    
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(f"DOCX not found: {docx_path}")
    if not docx_path.lower().endswith(".docx"):
        raise ValueError("Input must be a .docx file")

    full_json = get_document_as_json(docx_path, return_json=True)
    text_properties = full_json.get("textProperties", {})
    source_file = full_json.get("source_file")
    
    result = {
        "source_file": source_file,
        "textProperties": text_properties,
        "properties_count": len(text_properties)
    }
    
    if return_json:
        return result
    return {"properties_count": len(text_properties)}


@server.tool(
    title="Delete Elements",
    description=(
        "🗑️ DELETE DOCUMENT ELEMENTS BY ID (SINGLE OR MULTIPLE)\n\n"
        "Remove one or more document elements (paragraphs, tables, etc.) by their IDs.\n"
        "Supports bulk deletion for efficient document cleanup.\n\n"
        "📋 WORKFLOW:\n"
        "1. Call 'get_document_as_json' first to get element IDs\n"
        "2. Provide element_ids as string (single) or array (multiple)\n"
        "3. Choose response format: 'minimal', 'id_mapping', or 'full_document'\n\n"
        "⚠️ WARNING: This operation cannot be undone!\n\n"
        "Args: docx_path (str), element_ids (str|list), response_format (str), output_docx_path (optional)."
    ),
)
def delete_elements(
    docx_path: str,
    element_ids: Union[str, List[str]],
    response_format: Literal["minimal", "id_mapping", "full_document"] = "minimal",
    output_docx_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Delete one or more document elements by ID."""
    docx_path = _ensure_abs(docx_path)
    
    # Normalize to list
    if isinstance(element_ids, str):
        element_ids = [element_ids]
    
    _log_tool_call("delete_elements", {
        "docx_path": docx_path, 
        "element_ids": element_ids,
        "count": len(element_ids)
    })
    
    # Get current document JSON
    src_json = get_document_as_json(docx_path, return_json=True)
    
    # Create delete operations
    operations = []
    for element_id in element_ids:
        operations.append({
            "op": "delete",
            "id": element_id
        })
    
    patch_data = {"ops": operations}
    
    # Apply patch and get result
    result = _apply_patch_and_save(src_json, patch_data, docx_path, output_docx_path, response_format)
    
    result.update({
        "success": True,
        "deleted_element_ids": element_ids,
        "deleted_count": len(element_ids)
    })
    
    return result


@server.tool(
    title="Add Elements",
    description=(
        "➕ ADD NEW ELEMENTS TO DOCUMENT (SINGLE OR MULTIPLE)\n\n"
        "Add one or more new document elements at specified positions with proper text formatting.\n"
        "Supports bulk insertion for efficient document building.\n\n"
        "📋 WORKFLOW:\n"
        "1. Call 'get_document_as_json' to understand document structure\n"
        "2. Call 'get_document_text_properties' to see available formatting presets\n"
        "3. Create your new elements with proper structure\n"
        "4. Specify text_properties_ref for consistent formatting\n"
        "5. Choose position: 'after', 'before', or 'end'\n\n"
        "🔧 POSITION OPTIONS:\n"
        "• 'after': Insert after reference_element_id\n"
        "• 'before': Insert before reference_element_id  \n"
        "• 'end': Append at end of document (no reference needed)\n\n"
        "🎨 TEXT FORMATTING:\n"
        "• text_properties_ref: Use existing format from textProperties registry\n"
        "• If not specified, uses 'default_text_format' or inherits from reference element\n\n"
        "📝 ELEMENT STRUCTURE:\n"
        "Simple text: 'Hello world' (will be auto-wrapped in paragraph with specified formatting)\n"
        "Paragraph: {'type': 'paragraph', 'content': [{'type': 'run', 'text': 'New text'}]}\n"
        "Multiple: [element1, element2, ...] or ['text1', 'text2', ...]\n\n"
        "Args: docx_path (str), new_elements (dict|list|str), position (str), text_properties_ref (optional), reference_element_id (optional), response_format (str)."
    ),
)
def add_elements(
    docx_path: str,
    new_elements: Union[str, Dict[str, Any], List[Union[str, Dict[str, Any]]]],
    position: Literal["after", "before", "end"],
    text_properties_ref: Optional[str] = None,
    reference_element_id: Optional[str] = None,
    response_format: Literal["minimal", "id_mapping", "full_document"] = "minimal",
    output_docx_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Add one or more new elements to the document."""
    docx_path = _ensure_abs(docx_path)
    
    # Normalize to list
    if isinstance(new_elements, (str, dict)):
        new_elements = [new_elements]
    
    _log_tool_call("add_elements", {
        "docx_path": docx_path, 
        "position": position, 
        "reference_element_id": reference_element_id,
        "text_properties_ref": text_properties_ref,
        "element_count": len(new_elements)
    })
    
    if position in ["after", "before"] and not reference_element_id:
        return {"error": f"reference_element_id is required when position is '{position}'"}
    
    # Get current document JSON
    src_json = get_document_as_json(docx_path, return_json=True)
    available_properties = src_json.get("textProperties", {})
    
    # Prepare elements with proper formatting
    prepared_elements = []
    for element in new_elements:
        prepared_element = _prepare_element_with_formatting(
            element, text_properties_ref, available_properties
        )
        prepared_elements.append(prepared_element)
    
    # Create appropriate operation based on position
    if position == "end":
        patch_data = {
            "ops": [{
            "op": "insertAfter",
                "id": "_APPEND_AT_END_DUMMY_ID_",
                "nodes": prepared_elements
            }]
        }
    else:
        op = "insertAfter" if position == "after" else "insertBefore"
        patch_data = {
            "ops": [{
                "op": op,
            "id": reference_element_id,
                "nodes": prepared_elements
            }]
        }
    
    # Apply patch and get result
    result = _apply_patch_and_save(src_json, patch_data, docx_path, output_docx_path, response_format)
    
    result.update({
        "success": True,
        "added_position": position,
        "reference_element_id": reference_element_id,
        "added_count": len(new_elements)
    })
    
    return result


@server.tool(
    title="Edit Document (Combined Operations)",
    description=(
        "🖊️ PERFORM MULTIPLE DOCUMENT OPERATIONS IN SINGLE CALL\n\n"
        "Combine edits, additions, and deletions in one atomic operation.\n"
        "Perfect for complex document modifications with consistent results.\n\n"
        "📋 WORKFLOW:\n"
        "1. Call 'get_document_as_json' to get current structure and IDs\n"
        "2. Prepare your operations in the appropriate arrays\n"
        "3. Choose response format: 'minimal', 'id_mapping', or 'full_document'\n"
        "4. All operations are applied atomically\n\n"
        "🔧 OPERATION TYPES:\n"
        "• edits: [{'element_id': 'p-1', 'property_path': 'content', 'new_value': [...], 'text_properties_ref': 'format_name'}]\n"
        "• additions: [{'elements': [...], 'position': 'after', 'reference_id': 'p-2', 'text_properties_ref': 'format_name'}]\n"
        "• deletions: ['p-3', 'p-4']  # Just array of element IDs\n\n"
        "🎨 TEXT FORMATTING IN OPERATIONS:\n"
        "• For edits: text_properties_ref applies formatting when changing content\n"
        "• For additions: text_properties_ref applies to all new elements\n"
        "• If not specified, keeps existing formatting (edits) or uses default (additions)\n\n"
        "💡 EXECUTION ORDER: deletions → edits → additions\n\n"
        "Args: docx_path (str), edits (optional), additions (optional), deletions (optional), response_format (str)."
    ),
)
def edit_document(
    docx_path: str,
    edits: Optional[List[Dict[str, Any]]] = None,
    additions: Optional[List[Dict[str, Any]]] = None,
    deletions: Optional[List[str]] = None,
    response_format: Literal["minimal", "id_mapping", "full_document"] = "minimal",
    output_docx_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Perform multiple document operations in a single atomic call."""
    docx_path = _ensure_abs(docx_path)
    
    # Count operations
    edit_count = len(edits) if edits else 0
    add_count = len(additions) if additions else 0  
    delete_count = len(deletions) if deletions else 0
    
    _log_tool_call("edit_document", {
        "docx_path": docx_path,
        "edit_count": edit_count,
        "add_count": add_count, 
        "delete_count": delete_count
    })
    
    if edit_count + add_count + delete_count == 0:
        return {"error": "At least one operation (edits, additions, or deletions) must be provided"}
    
    # Get current document JSON
    src_json = get_document_as_json(docx_path, return_json=True)
    available_properties = src_json.get("textProperties", {})
    
    # Build operations array (order matters: deletes, edits, then adds)
    operations = []
    
    # 1. Deletions first
    if deletions:
        for element_id in deletions:
            operations.append({
                "op": "delete",
                "id": element_id
            })
    
    # 2. Edits second
    if edits:
        for edit in edits:
            new_value = edit["new_value"]
            text_props_ref = edit.get("text_properties_ref")
            
            # If editing content and text_properties_ref is specified, apply formatting
            if edit["property_path"] == "content" and text_props_ref:
                if isinstance(new_value, list):
                    # Apply text_properties_ref to runs that don't have it
                    formatted_content = []
                    for item in new_value:
                        if isinstance(item, dict) and item.get("type") == "run" and "textPropsRef" not in item:
                            item = item.copy()
                            item["textPropsRef"] = text_props_ref
                        formatted_content.append(item)
                    new_value = formatted_content
            
            operations.append({
                "op": "replace",
                "id": edit["element_id"],
                "path": edit["property_path"],
                "value": new_value
            })
    
    # 3. Additions last
    if additions:
        for addition in additions:
            position = addition["position"]
            elements = addition["elements"]
            reference_id = addition.get("reference_id")
            text_props_ref = addition.get("text_properties_ref")
            
            # Prepare elements with proper formatting
            if not isinstance(elements, list):
                elements = [elements]
            
            prepared_elements = []
            for element in elements:
                prepared_element = _prepare_element_with_formatting(
                    element, text_props_ref, available_properties
                )
                prepared_elements.append(prepared_element)
            
            if position == "end":
                operations.append({
                    "op": "insertAfter",
                    "id": "_APPEND_AT_END_DUMMY_ID_",
                    "nodes": prepared_elements
                })
            else:
                if not reference_id:
                    return {"error": f"reference_id is required for position '{position}'"}
                op = "insertAfter" if position == "after" else "insertBefore"
                operations.append({
                    "op": op,
                    "id": reference_id,
                    "nodes": prepared_elements
                })
    
    patch_data = {"ops": operations}
    
    # Apply patch and get result
    result = _apply_patch_and_save(src_json, patch_data, docx_path, output_docx_path, response_format)
    
    result.update({
        "success": True,
        "operations_applied": {
            "deletions": delete_count,
            "edits": edit_count,
            "additions": add_count,
            "total": len(operations)
        }
    })
    
    return result


@server.tool(
    title="Edit Element Content",
    description=(
        "🖊️ EDIT SINGLE ELEMENT CONTENT OR PROPERTIES\n\n"
        "Modify the content or properties of a single document element by ID.\n"
        "Perfect for quick text updates or formatting changes.\n\n"
        "📋 WORKFLOW:\n"
        "1. Call 'get_document_as_json' to get element IDs and current structure\n"
        "2. Optionally call 'get_document_text_properties' to see available formats\n"
        "3. Specify the element_id and property_path to modify\n"
        "4. Provide new_value and optionally text_properties_ref for formatting\n\n"
        "🔧 COMMON PROPERTY PATHS:\n"
        "• 'content': Update paragraph text content (array of runs)\n"
        "• 'pPr.jc': Change paragraph alignment ('left', 'center', 'right', 'justify')\n"
        "• 'pPr.styleId': Change paragraph style\n\n"
        "🎨 TEXT FORMATTING:\n"
        "• text_properties_ref: Apply specific format from textProperties registry\n"
        "• If not specified, keeps existing formatting when editing content\n"
        "• Only applies when property_path is 'content'\n\n"
        "💡 CONTENT EXAMPLES:\n"
        "Simple text: 'New text content' (auto-wrapped with formatting)\n"
        "Run array: [{'type': 'run', 'text': 'New text'}] (manual control)\n\n"
        "Args: docx_path (str), element_id (str), property_path (str), new_value (any), text_properties_ref (optional)."
    ),
)
def edit_element_content(
    docx_path: str,
    element_id: str,
    property_path: str,
    new_value: Any,
    text_properties_ref: Optional[str] = None,
    response_format: Literal["minimal", "id_mapping", "full_document"] = "minimal",
    output_docx_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Edit content or properties of a single document element."""
    docx_path = _ensure_abs(docx_path)
    _log_tool_call("edit_element_content", {
        "docx_path": docx_path, 
        "element_id": element_id, 
        "property_path": property_path,
        "text_properties_ref": text_properties_ref
    })
    
    # Get current document JSON
    src_json = get_document_as_json(docx_path, return_json=True)
    available_properties = src_json.get("textProperties", {})
    
    # Handle text formatting for content updates
    if property_path == "content" and text_properties_ref:
        # If new_value is a simple string, convert to properly formatted runs
        if isinstance(new_value, str):
            run_content = {"type": "run", "text": new_value, "textPropsRef": text_properties_ref}
            new_value = [run_content]
        elif isinstance(new_value, list):
            # Apply text_properties_ref to runs that don't have formatting
            formatted_content = []
            for item in new_value:
                if isinstance(item, dict) and item.get("type") == "run" and "textPropsRef" not in item:
                    item = item.copy()
                    item["textPropsRef"] = text_properties_ref
                formatted_content.append(item)
            new_value = formatted_content
    
    # Create replace operation
    patch_data = {
        "ops": [{
            "op": "replace",
            "id": element_id,
            "path": property_path,
            "value": new_value
        }]
    }
    
    # Apply patch
    result = _apply_patch_and_save(src_json, patch_data, docx_path, output_docx_path, response_format)
    
    result.update({
        "success": True,
        "edited_element_id": element_id,
        "property_path": property_path,
        "applied_text_properties_ref": text_properties_ref
    })
    
    return result


def _apply_patch_and_save(
    src_json: Dict[str, Any], 
    patch_data: Dict[str, Any], 
    docx_path: str, 
    output_docx_path: Optional[str] = None,
    response_format: str = "minimal"
) -> Dict[str, Any]:
    """Internal helper to apply patch, save to new DOCX file, and return appropriate response."""
    # Generate versioned output path
    base_noext, ext = os.path.splitext(docx_path)
    base_root = _strip_version_suffix(base_noext)
    version = _get_next_version_number(base_root)
    
    if output_docx_path:
        output_docx = _ensure_abs(output_docx_path)
    else:
        output_docx = _derive_versioned_path(base_root, ext, version)
    
    # Apply patch
    _buf = io.StringIO()
    with contextlib.redirect_stdout(_buf):
        patched_json = pjson.apply_patch(src_json, patch_data)
    
    # Create new DOCX from patched JSON
    source_docx_path = patched_json.get("source_file", docx_path)
    if not os.path.isfile(source_docx_path):
        source_docx_path = docx_path
    
    from docx import Document
    document = Document(source_docx_path)
    j2d.clear_body(document) 
    j2d.reconstruct_body(document, patched_json.get("body", []), patched_json.get("textProperties"))
    document.save(output_docx)
    
    # Prepare response based on format
    result = {
        "output_docx_path": output_docx,
        "version": version
    }
    
    if response_format == "minimal":
        # Just basic info
        pass
    elif response_format == "id_mapping":
        # Include ID mappings
        result["id_mapping"] = _generate_id_mapping(src_json, patched_json)
    elif response_format == "full_document":
        # Include full updated document JSON
        result["updated_document"] = patched_json
        result["id_mapping"] = _generate_id_mapping(src_json, patched_json)
    
    return result


# -------------------- Transport and configuration --------------------

def _load_env_file() -> None:
    """Load key=value pairs from a local env file if present."""
    candidates = ["mcp.env", ".env"]
    for name in candidates:
        path = os.path.join(os.getcwd(), name)
        if os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    for raw in f:
                        line = raw.strip()
                        if not line or line.startswith("#"):
                            continue
                        if "=" not in line:
                            continue
                        k, v = line.split("=", 1)
                        k = k.strip()
                        v = v.strip()
                        if k and (k not in os.environ or os.environ[k] == ""):
                            os.environ[k] = v
                break
            except Exception:
                pass


def _run_from_env() -> None:
    """Run the MCP server using transport selected by env."""
    mode = os.environ.get("MCP_TRANSPORT", "stdio").strip().lower()
    host = os.environ.get("MCP_HOST", "127.0.0.1")
    try:
        port = int(os.environ.get("MCP_PORT", "8765"))
    except ValueError:
        port = 8765
    path = os.environ.get("MCP_PATH", "/mcp")
    sse_path = os.environ.get("MCP_SSE_PATH")
    msg_path = os.environ.get("MCP_MESSAGE_PATH")

    try:
        settings = getattr(server, "settings", None)
        if settings is not None:
            settings.host = host
            settings.port = port
            
            if mode == "sse":
                if sse_path:
                    settings.sse_path = sse_path
                else:
                    settings.sse_path = "/sse"
                
                if msg_path:
                    settings.message_path = msg_path if msg_path.endswith('/') else (msg_path + '/')
                else:
                    sse_base = getattr(settings, "sse_path", "/sse")
                    sse_base = sse_base if sse_base.startswith('/') else ("/" + sse_base)
                    sse_base = sse_base.rstrip('/')
                    settings.message_path = f"{sse_base}/messages/"
    except Exception as e:
        print(f"Warning: Could not configure server settings: {e}")

    if mode in ("", "stdio", "std", "default"):
        server.run()
        return

    try:
        import asyncio
        if mode == "sse" and hasattr(server, "run_sse_async"):
            asyncio.run(server.run_sse_async(mount_path=path))
            return
        if mode in ("http", "https") and hasattr(server, "run_streamable_http_async"):
            try:
                if getattr(server, "settings", None) is not None:
                    server.settings.streamable_http_path = path
            except Exception:
                pass
            asyncio.run(server.run_streamable_http_async())
            return
        if mode in ("ws", "websocket") and hasattr(server, "run_websocket_async"):
            asyncio.run(server.run_websocket_async())
            return
    except Exception as e:
        print(f"Network transport '{mode}' failed: {e}. Falling back to stdio.")

    print(f"Transport '{mode}' not supported. Using stdio.")
    server.run()


def _apply_cli_overrides(argv: list[str]) -> None:
    """Parse simple CLI flags and override environment values."""
    import argparse
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--transport", dest="transport")
    parser.add_argument("--host", dest="host")
    parser.add_argument("--port", dest="port")
    parser.add_argument("--path", dest="path")
    parser.add_argument("--sse-path", dest="sse_path")
    parser.add_argument("--message-path", dest="message_path")

    args, _ = parser.parse_known_args(argv)

    def setenv(k: str, v: str | None):
        if v is not None and v != "":
            os.environ[k] = v

    setenv("MCP_TRANSPORT", args.transport)
    setenv("MCP_HOST", args.host)
    setenv("MCP_PORT", args.port)
    setenv("MCP_PATH", args.path)
    setenv("MCP_SSE_PATH", args.sse_path)
    setenv("MCP_MESSAGE_PATH", args.message_path)


# -------------------- Logging --------------------

TOOL_CALL_LOG: deque[Dict[str, Any]] = deque(maxlen=100)


def _log_tool_call(name: str, args: Dict[str, Any]) -> None:
    entry = {
        "time": datetime.utcnow().isoformat() + "Z",
        "tool": name,
        "args": args,
    }
    print(f"[TOOL] {name} args={args}")
    TOOL_CALL_LOG.append(entry)


if __name__ == "__main__":
    _load_env_file()
    import sys as _sys
    _apply_cli_overrides(_sys.argv[1:])
    _run_from_env()