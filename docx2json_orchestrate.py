#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import json
import tempfile
from copy import deepcopy
import re

from docx2json_runs_only import export_docx_to_json as export_runs_only


def _canon_props(d: dict) -> str:
    def normalize(obj):
        if isinstance(obj, dict):
            return {k: normalize(obj[k]) for k in sorted(obj.keys())}
        if isinstance(obj, list):
            return [normalize(x) for x in obj]
        return obj
    return json.dumps(normalize(d or {}), separators=(",", ":"))


def _gen_semantic_name(props: dict) -> str:
    parts = []
    rpr = props.get("text_full_properties", {}).get("rPr", {}) if "rPr" not in props else props.get("rPr", {})
    # style tag (paragraph)
    paragraphStyleName = props.get("text_full_properties", {}).get("paragraphStyleName") if "paragraphStyleName" not in props else props.get("paragraphStyleName")
    paragraphStyleId = props.get("text_full_properties", {}).get("paragraphStyleId") if "paragraphStyleId" not in props else props.get("paragraphStyleId")
    style_tag = None
    if isinstance(paragraphStyleName, str) and paragraphStyleName.strip():
        s = paragraphStyleName.strip().lower()
        m = re.search(r"heading\s*(\d+)", s)
        if m:
            style_tag = f"heading_{m.group(1)}"
        elif s == "title":
            style_tag = "title"
        else:
            style_tag = s.replace(" ", "_")
    elif isinstance(paragraphStyleId, str) and paragraphStyleId.strip():
        sid = paragraphStyleId.strip().lower()
        m = re.search(r"heading\s*(\d+)", sid)
        if m:
            style_tag = f"heading_{m.group(1)}"
        elif sid == "title":
            style_tag = "title"
        else:
            style_tag = sid.replace(" ", "_")
    if style_tag:
        parts.append(style_tag)
    # emphasis
    if rpr.get("bold"): parts.append("bold")
    if rpr.get("italic"): parts.append("italic")
    if rpr.get("underline"): parts.append("underline")
    # font
    font = props.get("text_full_properties", {}).get("fontName") if "fontName" not in props else props.get("fontName")
    if not font:
        font = rpr.get("rFonts", {}).get("ascii")
    if font:
        parts.append(str(font).lower().replace(" ", "_"))
    # size
    szhp = rpr.get("sizeHalfPoints")
    if isinstance(szhp, int):
        parts.append(f"{szhp//2}pt")
    # color / highlight
    color = rpr.get("color")
    if isinstance(color, str) and color:
        parts.append(f"color_{color.lower()}")
    hi = rpr.get("highlight")
    if isinstance(hi, str) and hi:
        parts.append(f"highlight_{hi.lower()}")
    base = ("_".join(parts) if parts else "default_text_format") + "_format"
    return base


def _merge_runs_sequence(items):
    out = []
    for it in items:
        if out and isinstance(it, dict) and it.get("type") == "run" and isinstance(out[-1], dict) and out[-1].get("type") == "run":
            prev = out[-1]
            if prev.get("textPropsRef") and prev.get("textPropsRef") == it.get("textPropsRef"):
                # Merge text
                if "text" in prev and "text" in it:
                    prev["text"] += it["text"]
                    continue
                # Merge chunks
                if "chunks" in prev and "chunks" in it:
                    prev["chunks"].extend(it["chunks"])
                    continue
                # Mixed: normalize to chunks
                def to_chunks(run):
                    if "chunks" in run:
                        return run["chunks"]
                    if "text" in run:
                        return [{"type": "text", "text": run["text"]}]
                    return []
                prev_chunks = to_chunks(prev)
                prev["chunks"] = prev_chunks + to_chunks(it)
                prev.pop("text", None)
                continue
        out.append(it)
    return out


def _assign_ids(blocks, pid_start=1, tid_start=1):
    pcount = pid_start
    tcount = tid_start
    def walk(blks):
        nonlocal pcount, tcount
        out = []
        for blk in blks:
            if not isinstance(blk, dict):
                out.append(blk)
                continue
            if blk.get("type") == "paragraph":
                nb = deepcopy(blk)
                nb["id"] = f"p-{pcount}"
                pcount += 1
                out.append(nb)
            elif blk.get("type") == "table":
                nb = deepcopy(blk)
                nb["id"] = f"t-{tcount}"
                tcount += 1
                # recurse into cells
                new_rows = []
                for row in nb.get("rows", []) or []:
                    new_cells = []
                    for cell in row.get("cells", []) or []:
                        nc = deepcopy(cell)
                        nc["content"] = walk(cell.get("content", []) or [])
                        new_cells.append(nc)
                    new_rows.append({"cells": new_cells})
                nb["rows"] = new_rows
                out.append(nb)
            else:
                out.append(blk)
        return out
    return walk(blocks)


def _annotate_with_refs(body, name_to_key):
    def annotate(blocks):
        out = []
        for blk in blocks:
            if not isinstance(blk, dict):
                out.append(blk)
                continue
            if blk.get("type") == "paragraph":
                nb = deepcopy(blk)
                new_content = []
                for item in blk.get("content", []) or []:
                    if isinstance(item, dict) and item.get("type") == "run":
                        run = deepcopy(item)
                        tfp = run.pop("text_full_properties", None)
                        if tfp:
                            tkey = _canon_props(tfp)
                            assigned = None
                            for nm, k in name_to_key.items():
                                if k == tkey:
                                    assigned = nm
                                    break
                            if assigned:
                                run["textPropsRef"] = assigned
                        new_content.append(run)
                    else:
                        if isinstance(item, dict) and item.get("type") == "hyperlink":
                            nh = deepcopy(item)
                            nh_runs = []
                            for r in item.get("runs", []) or []:
                                rr = deepcopy(r)
                                tfp = rr.pop("text_full_properties", None)
                                if tfp:
                                    tkey = _canon_props(tfp)
                                    assigned = None
                                    for nm, k in name_to_key.items():
                                        if k == tkey:
                                            assigned = nm
                                            break
                                    if assigned:
                                        rr["textPropsRef"] = assigned
                                nh_runs.append(rr)
                            nh["runs"] = nh_runs
                            new_content.append(nh)
                        else:
                            new_content.append(item)
                nb["content"] = _merge_runs_sequence(new_content)
                out.append(nb)
            elif blk.get("type") == "table":
                nb = deepcopy(blk)
                new_rows = []
                for row in blk.get("rows", []) or []:
                    new_cells = []
                    for cell in row.get("cells", []) or []:
                        nc = deepcopy(cell)
                        nc["content"] = annotate(cell.get("content", []) or [])
                        new_cells.append(nc)
                    new_rows.append({"cells": new_cells})
                nb["rows"] = new_rows
                out.append(nb)
            else:
                out.append(blk)
        return out
    return annotate(body)


def build_docx_json(input_docx: str, output_json: str):
    # Step 1: export runs-only JSON to a temp file, then load
    runs_only_path = output_json.replace(".json", "") + ".runs_only.json"
    export_runs_only(input_docx, runs_only_path)
    with open(runs_only_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    source_file = data.get("source_file")
    body = data.get("body", [])

    # Step 2: collect unique text_full_properties and assign semantic IDs
    key_to_props = {}
    def collect(blocks):
        for blk in blocks:
            if not isinstance(blk, dict):
                continue
            if blk.get("type") == "paragraph":
                for item in blk.get("content", []) or []:
                    if isinstance(item, dict) and item.get("type") == "run":
                        tfp = item.get("text_full_properties")
                        if tfp:
                            key = _canon_props(tfp)
                            key_to_props.setdefault(key, deepcopy(tfp))
                    elif isinstance(item, dict) and item.get("type") == "hyperlink":
                        for r in item.get("runs", []) or []:
                            tfp = r.get("text_full_properties")
                            if tfp:
                                key = _canon_props(tfp)
                                key_to_props.setdefault(key, deepcopy(tfp))
            elif blk.get("type") == "table":
                for row in blk.get("rows", []) or []:
                    for cell in row.get("cells", []) or []:
                        collect(cell.get("content", []) or [])
    collect(body)

    # Assign semantic IDs (include style info)
    registry = {}
    name_to_key = {}
    for key, props in key_to_props.items():
        base = _gen_semantic_name(props)
        name = base
        n = 2
        while name in name_to_key and name_to_key[name] != key:
            name = f"{base}_{n}"
            n += 1
        name_to_key[name] = key
        registry[name] = props

    # Step 3: annotate body with refs and compact adjacent runs
    annotated = _annotate_with_refs(body, name_to_key)

    # Step 4: assign unique IDs to paragraphs and tables
    final_body = _assign_ids(annotated, pid_start=1, tid_start=1)

    # Step 5: write final combined JSON
    out = {
        "source_file": source_file,
        "textProperties": registry,
        "body": final_body,
    }
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def main():
    if len(sys.argv) != 3:
        print("Usage: python docx2json_orchestrate.py <input.docx> <output.json>", file=sys.stderr)
        sys.exit(2)
    docx_path, out_json = sys.argv[1], sys.argv[2]
    if not os.path.isfile(docx_path):
        print(f"Input DOCX not found: {docx_path}", file=sys.stderr)
        sys.exit(2)
    build_docx_json(docx_path, out_json)
    print(f"Wrote {out_json}")


if __name__ == "__main__":
    main()


