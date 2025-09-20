#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX -> JSON (Runs-Only, Actual Formatting)

Goal: For each run in the document, emit ONLY the formatting that is actually
visible on that run in the rendered document, without leaking document defaults.

Rules used to compute the visible run properties (highest priority last):
  1) Paragraph style chain rPr (basedOn chain, base -> derived)
  2) Numbering level rPr (plus overrides), if the paragraph is in a list
  3) Paragraph-level rPr (direct in pPr)
  4) Character style chain rPr (basedOn chain, base -> derived)
  5) Run direct rPr

We DO NOT fall back to docDefaults when emitting output. If a property is only
provided by docDefaults and nowhere else, we leave it out (to avoid Calibri leaks).

Output JSON shape:
  {
    "source_file": <abs path>,
    "body": [
      { "type": "paragraph", "p": {optional minimal p props}, "content": [
          { "type": "run", "text": "...", "rPr": {visible props for this run} },
          { "type": "hyperlink", "target": "...", "runs": [ ... runs as above ... ] }
      ]},
      { "type": "table", rows: [ { cells: [ { content: [ paragraphs/tables... ] } ] } ] }
    ]
  }
"""

import sys, os, json, zipfile
from copy import deepcopy
from lxml import etree

NS = {
    "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
}

def qn(nskey, tag):
    return "{%s}%s" % (NS[nskey], tag)

def read_xml(z, path):
    with z.open(path) as f:
        return etree.parse(f)

def rels_path_for(part_path):
    d, fname = os.path.split(part_path)
    return os.path.join(d, "_rels", fname + ".rels")

def load_rels(z, part_path):
    rels = {}
    rp = rels_path_for(part_path)
    if rp in z.namelist():
        x = read_xml(z, rp)
        for rel in x.getroot().iterfind(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rId = rel.get("Id")
            rels[rId] = {
                "type": rel.get("Type"),
                "target": rel.get("Target"),
                "mode": rel.get("TargetMode", "Internal")
            }
    return rels

# ----------------- THEME / STYLES / NUMBERING -----------------

def parse_theme(z):
    theme = {"majorLatin": None, "minorLatin": None}
    path = "word/theme/theme1.xml"
    if path not in z.namelist():
        return theme
    try:
        x = read_xml(z, path)
        root = x.getroot()
        major = root.find(".//a:themeElements/a:fontScheme/a:majorFont/a:latin", namespaces=NS)
        minor = root.find(".//a:themeElements/a:fontScheme/a:minorFont/a:latin", namespaces=NS)
        if major is not None and major.get("typeface"):
            theme["majorLatin"] = major.get("typeface")
        if minor is not None and minor.get("typeface"):
            theme["minorLatin"] = minor.get("typeface")
    except Exception:
        pass
    return theme

def get_bool(el):
    if el is None: return None
    v = el.get(qn("w","val"))
    if v is None: return True
    # Handles val="false", val="0", val="off"
    return v.lower() not in ("false", "0", "off")

def extract_rPr(rPr):
    if rPr is None: return {}
    d = {}
    rs = rPr.find(qn("w","rStyle"))
    if rs is not None and rs.get(qn("w","val")):
        d["rStyle"] = rs.get(qn("w","val"))
    for key, tag in [
        ("bold","b"), ("italic","i"), ("strike","strike"), ("dstrike","dstrike"),
        ("caps","caps"), ("smallCaps","smallCaps"), ("emboss","emboss"),
        ("imprint","imprint"), ("outline","outline"), ("shadow","shadow"),
        ("vanish","vanish"), ("rtl","rtl"),
    ]:
        v = get_bool(rPr.find(qn("w",tag)))
        if v is not None: d[key] = v
    u = rPr.find(qn("w","u"))
    if u is not None and u.get(qn("w","val")) and u.get(qn("w","val")) != "none":
        d["underline"] = {"val": u.get(qn("w","val"))}
    c = rPr.find(qn("w","color"))
    if c is not None and c.get(qn("w","val")):
        d["color"] = c.get(qn("w","val"))
    hi = rPr.find(qn("w","highlight"))
    if hi is not None and hi.get(qn("w","val")):
        d["highlight"] = hi.get(qn("w","val"))
    rf = rPr.find(qn("w","rFonts"))
    if rf is not None:
        fonts = {}
        for fkey in ["ascii","hAnsi","eastAsia","cs","asciiTheme","hAnsiTheme","eastAsiaTheme","cstheme"]:
            val = rf.get(qn("w", fkey)) if fkey != "cstheme" else rf.get(qn("w","cstheme"))
            if val: fonts[fkey] = val
        if fonts: d["rFonts"] = fonts
    sz = rPr.find(qn("w","sz"))
    if sz is not None and sz.get(qn("w","val")):
        d["sizeHalfPoints"] = int(sz.get(qn("w","val")))
    szcs = rPr.find(qn("w","szCs"))
    if szcs is not None and szcs.get(qn("w","val")):
        d["sizeCsHalfPoints"] = int(szcs.get(qn("w","val")))
    va = rPr.find(qn("w","vertAlign"))
    if va is not None and va.get(qn("w","val")):
        d["vertAlign"] = va.get(qn("w","val"))
    return d

def extract_pPr(pPr):
    if pPr is None: return {}
    d = {}
    pstyle = pPr.find(qn("w","pStyle"))
    if pstyle is not None and pstyle.get(qn("w","val")):
        d["styleId"] = pstyle.get(qn("w","val"))
    numPr = pPr.find(qn("w","numPr"))
    if numPr is not None:
        n = {}
        ilvl = numPr.find(qn("w","ilvl"))
        numId = numPr.find(qn("w","numId"))
        if ilvl is not None and ilvl.get(qn("w","val")): n["ilvl"] = int(ilvl.get(qn("w","val")))
        if numId is not None and numId.get(qn("w","val")): n["numId"] = int(numId.get(qn("w","val")))
        if n: d["numPr"] = n
    rpr = extract_rPr(pPr.find(qn("w","rPr")))
    if rpr: d["rPr"] = rpr
    # Minimal extras helpful for layout
    jc = pPr.find(qn("w","jc"))
    if jc is not None and jc.get(qn("w","val")):
        d["jc"] = jc.get(qn("w","val"))
    return d

def parse_styles(z):
    styles = {"docDefaults": {"rPr": {}, "pPr": {}}, "styles": {}, "defaultParagraphStyleId": None, "defaultCharacterStyleId": None}
    path = "word/styles.xml"
    if path not in z.namelist():
        return styles
    x = read_xml(z, path)
    root = x.getroot()
    dd = root.find(qn("w","docDefaults"))
    if dd is not None:
        rdef = dd.find(".//w:rPrDefault/w:rPr", namespaces=NS)
        pdef = dd.find(".//w:pPrDefault/w:pPr", namespaces=NS)
        styles["docDefaults"]["rPr"] = extract_rPr(rdef)
        styles["docDefaults"]["pPr"] = extract_pPr(pdef)
    for s in root.findall(qn("w","style")):
        sid = s.get(qn("w","styleId"))
        stype = s.get(qn("w","type"))
        name = s.find(qn("w","name"))
        basedOn = s.find(qn("w","basedOn"))
        is_default = s.get(qn("w","default")) == "1"
        styles["styles"][sid] = {
            "styleId": sid,
            "type": stype,
            "name": name.get(qn("w","val")) if name is not None else "",
            "basedOn": basedOn.get(qn("w","val")) if basedOn is not None else "",
            "pPr": extract_pPr(s.find(qn("w","pPr"))),
            "rPr": extract_rPr(s.find(qn("w","rPr"))),
        }
        if is_default and stype == "paragraph" and not styles.get("defaultParagraphStyleId"):
            styles["defaultParagraphStyleId"] = sid
        if is_default and stype == "character" and not styles.get("defaultCharacterStyleId"):
            styles["defaultCharacterStyleId"] = sid
    return styles

def parse_numbering(z):
    out = {"abstractNums": [], "nums": []}
    path = "word/numbering.xml"
    if path not in z.namelist():
        return out
    x = read_xml(z, path)
    for an in x.getroot().findall(qn("w","abstractNum")):
        aid = int(an.get(qn("w","abstractNumId")))
        levels = []
        for lvl in an.findall(qn("w","lvl")):
            ilvl = int(lvl.get(qn("w","ilvl")))
            numFmt = (lvl.find(qn("w","numFmt")).get(qn("w","val"))
                      if lvl.find(qn("w","numFmt")) is not None else "")
            lvlText = (lvl.find(qn("w","lvlText")).get(qn("w","val"))
                       if lvl.find(qn("w","lvlText")) is not None else "")
            start = (int(lvl.find(qn("w","start")).get(qn("w","val")))
                     if lvl.find(qn("w","start")) is not None else None)
            pPr = extract_pPr(lvl.find(qn("w","pPr")))
            rPr = extract_rPr(lvl.find(qn("w","rPr")))
            levels.append({
                "ilvl": ilvl,
                "numFmt": numFmt,
                "lvlText": lvlText,
                "start": start,
                "pPr": pPr,
                "rPr": rPr
            })
        out["abstractNums"].append({"abstractNumId": aid, "levels": levels})
    for n in x.getroot().findall(qn("w","num")):
        nid = int(n.get(qn("w","numId")))
        absIdEl = n.find(qn("w","abstractNumId"))
        absId = int(absIdEl.get(qn("w","val"))) if absIdEl is not None else None
        overrides = []
        for olvl in n.findall(qn("w","lvlOverride")):
            ilvl = int(olvl.get(qn("w","ilvl")))
            lvl = olvl.find(qn("w","lvl"))
            pPr = extract_pPr(lvl.find(qn("w","pPr"))) if lvl is not None else {}
            rPr = extract_rPr(lvl.find(qn("w","rPr"))) if lvl is not None else {}
            startOverride = None
            so = olvl.find(qn("w","startOverride"))
            if so is not None and so.get(qn("w","val")):
                try:
                    startOverride = int(so.get(qn("w","val")))
                except Exception:
                    startOverride = None
            overrides.append({
                "ilvl": ilvl,
                "pPr": pPr,
                "rPr": rPr,
                "startOverride": startOverride
            })
        out["nums"].append({"numId": nid, "abstractNumId": absId, "overrides": overrides})
    return out

def resolve_style_chain(styles, style_id):
    pPr, rPr = {}, {}
    visited = set()
    chain = []
    cur = style_id
    while cur and cur not in visited and cur in styles["styles"]:
        visited.add(cur)
        chain.append(cur)
        cur = styles["styles"][cur].get("basedOn") or ""
    for sid in reversed(chain):  # base -> derived
        s = styles["styles"][sid]
        for k, v in s.get("pPr", {}).items():
            pPr[k] = deepcopy(v)
        for k, v in s.get("rPr", {}).items():
            rPr[k] = deepcopy(v)
    return pPr, rPr

def get_numbering_rpr(num_pr, numbering):
    if not isinstance(num_pr, dict):
        return {}
    num_id = num_pr.get("numId")
    ilvl = num_pr.get("ilvl", 0)
    if num_id is None:
        return {}
    num_def = None
    for n in numbering.get("nums", []):
        if n.get("numId") == num_id:
            num_def = n
            break
    if not num_def:
        return {}
    abs_def = None
    for an in numbering.get("abstractNums", []):
        if an.get("abstractNumId") == num_def.get("abstractNumId"):
            abs_def = an
            break
    if not abs_def:
        return {}
    lvl_def = None
    for lvl in abs_def.get("levels", []):
        if lvl.get("ilvl") == ilvl:
            lvl_def = lvl
            break
    lvl_r = deepcopy(lvl_def.get("rPr", {})) if lvl_def else {}
    # override
    for ov in num_def.get("overrides", []):
        if ov.get("ilvl") == ilvl:
            lvl_r.update(deepcopy(ov.get("rPr", {})))
            break
    return lvl_r

def get_numbering_info(num_pr, numbering):
    info = {"present": False}
    if not isinstance(num_pr, dict):
        return info
    num_id = num_pr.get("numId")
    ilvl = num_pr.get("ilvl", 0)
    if num_id is None:
        return info
    info.update({"present": True, "numId": num_id, "ilvl": ilvl})
    # find num/abstract
    num_def = None
    for n in numbering.get("nums", []):
        if n.get("numId") == num_id:
            num_def = n
            break
    if not num_def:
        return info
    abs_def = None
    for an in numbering.get("abstractNums", []):
        if an.get("abstractNumId") == num_def.get("abstractNumId"):
            abs_def = an
            break
    level = None
    if abs_def:
        for lvl in abs_def.get("levels", []):
            if lvl.get("ilvl") == ilvl:
                level = lvl
                break
    numFmt = (level or {}).get("numFmt")
    lvlText = (level or {}).get("lvlText")
    start = (level or {}).get("start")
    startOverride = None
    for ov in num_def.get("overrides", []):
        if ov.get("ilvl") == ilvl and ov.get("startOverride") is not None:
            startOverride = ov.get("startOverride")
            break
    info.update({
        "format": numFmt,
        "lvlText": lvlText,
        "start": start,
        "overrideStart": startOverride,
        "isMultiLevel": bool(ilvl and ilvl > 0),
    })
    # classify
    fmt = (numFmt or "").lower()
    if fmt == "bullet":
        info.update({"listKind": "bullet", "isBullet": True})
    elif fmt in ("decimal", "decimalZero", "ordinal"):
        info.update({"listKind": "number", "isBullet": False})
    elif fmt in ("lowerletter", "upperletter"):
        info.update({"listKind": "letter", "isBullet": False})
    elif fmt in ("lowerroman", "upperroman"):
        info.update({"listKind": "roman", "isBullet": False})
    else:
        info.update({"listKind": fmt or "unknown", "isBullet": False})
    # restart/continue
    if startOverride is not None:
        info["restart"] = True
        info["continue"] = False
    else:
        info["restart"] = False
        info["continue"] = True
    return info

def merge_rpr(base, override):
    out = deepcopy(base or {})
    for k, v in (override or {}).items():
        out[k] = deepcopy(v)
    return out

def resolve_theme_font(name_or_theme, theme):
    if not isinstance(name_or_theme, str):
        return None
    if name_or_theme.startswith("+"):
        low = name_or_theme.lower()
        if low.startswith("+major-latin") and theme.get("majorLatin"):
            return theme.get("majorLatin")
        if low.startswith("+minor-latin") and theme.get("minorLatin"):
            return theme.get("minorLatin")
    return name_or_theme

def pick_visible_font(run_rpr, char_rpr, para_rpr, pstyle_rpr, num_rpr, theme):
    # Order of picking: run -> character style -> paragraph style -> numbering -> paragraph-level rPr
    for src in (run_rpr, char_rpr, pstyle_rpr, num_rpr, para_rpr):
        rf = (src or {}).get("rFonts") or {}
        name = rf.get("ascii") or rf.get("hAnsi") or rf.get("asciiTheme") or rf.get("hAnsiTheme")
        name = resolve_theme_font(name, theme)
        if isinstance(name, str) and name.strip():
            return {"ascii": name, "hAnsi": name}
    return None

def pick_visible_size(run_rpr, char_rpr, para_rpr, pstyle_rpr, num_rpr):
    for src in (run_rpr, char_rpr, pstyle_rpr, num_rpr, para_rpr):
        sz = (src or {}).get("sizeHalfPoints") or (src or {}).get("sizeCsHalfPoints")
        if isinstance(sz, int):
            return sz
    return None

def pick_visible_alignment(declared_pPr, pstyle_p):
    jc = (declared_pPr or {}).get("jc")
    if isinstance(jc, str) and jc.strip():
        return jc
    jc2 = (pstyle_p or {}).get("jc")
    if isinstance(jc2, str) and jc2.strip():
        return jc2
    return None

def pick_visible_bool(prop, *rprs):
    for src in rprs:
        v = (src or {}).get(prop)
        if v is not None: # Will stop on True OR False
            return v
    return None # Inherit (effectively false)

def pick_visible_color(*rprs):
    for src in rprs:
        c = (src or {}).get("color")
        if isinstance(c, str) and c.strip() and c.upper() not in ("000000", "AUTO", "WINDOWTEXT"):
            return c
    return None

def build_visible_run_rpr(run_rpr, char_rpr, para_rpr, pstyle_rpr, num_rpr, theme):
    out = {}
    font = pick_visible_font(run_rpr, char_rpr, para_rpr, pstyle_rpr, num_rpr, theme)
    if font:
        out["rFonts"] = font
    sz = pick_visible_size(run_rpr, char_rpr, para_rpr, pstyle_rpr, num_rpr)
    if isinstance(sz, int):
        out["sizeHalfPoints"] = sz
    for bkey in ("bold", "italic", "strike"):
        val = pick_visible_bool(bkey, run_rpr, char_rpr, pstyle_rpr, num_rpr, para_rpr)
        if val:
            out[bkey] = True
        # If the run explicitly disables the property, include False to override style
        if isinstance(run_rpr, dict) and (bkey in (run_rpr or {})) and run_rpr.get(bkey) is False:
            out[bkey] = False
    u = (run_rpr or {}).get("underline") or (char_rpr or {}).get("underline") or (pstyle_rpr or {}).get("underline") or (num_rpr or {}).get("underline") or (para_rpr or {}).get("underline")
    if isinstance(u, dict) and u.get("val") and u.get("val") != "none":
        out["underline"] = {"val": u.get("val")}
    color = pick_visible_color(run_rpr, char_rpr, pstyle_rpr, num_rpr, para_rpr)
    if color:
        out["color"] = color
    hi = (run_rpr or {}).get("highlight") or (char_rpr or {}).get("highlight") or (pstyle_rpr or {}).get("highlight") or (num_rpr or {}).get("highlight") or (para_rpr or {}).get("highlight")
    if isinstance(hi, str) and hi.strip():
        out["highlight"] = hi
    va = (run_rpr or {}).get("vertAlign") or (char_rpr or {}).get("vertAlign") or (pstyle_rpr or {}).get("vertAlign") or (num_rpr or {}).get("vertAlign") or (para_rpr or {}).get("vertAlign")
    if isinstance(va, str) and va.strip():
        out["vertAlign"] = va
    return out

# ----------------- DOCUMENT PARSING -----------------

def text_from_t(t_el):
    return t_el.text if t_el is not None and t_el.text is not None else ""

def collect_run_chunks(run):
    chunks = []
    for child in run:
        if child.tag == qn("w","t"):
            chunks.append({"type":"text", "text": text_from_t(child)})
        elif child.tag == qn("w","tab"):
            chunks.append({"type":"tab"})
        elif child.tag == qn("w","br"):
            chunks.append({"type":"break", "breakType": child.get(qn("w","type")) or "textWrapping"})
        elif child.tag == qn("w","instrText"):
            chunks.append({"type":"fieldInstr", "text": child.text or ""})
        elif child.tag == qn("w","fldChar"):
            chunks.append({"type":"fldChar", "charType": child.get(qn("w","fldCharType")) or ""})
    return chunks

def consolidate_runs(content_list):
    out = []
    def canon_props(run_obj: dict) -> str:
        props = run_obj.get("text_full_properties") or {}
        try:
            return json.dumps(props, sort_keys=True, separators=(",", ":"))
        except Exception:
            return str(props)
    for item in content_list:
        if out and isinstance(item, dict) and item.get("type") == "run" and isinstance(out[-1], dict) and out[-1].get("type") == "run":
            prev = out[-1]
            if canon_props(prev) == canon_props(item):
                # Merge text
                if "text" in prev and "text" in item:
                    prev["text"] += item["text"]
                    continue
                # Merge chunks
                if "chunks" in prev and "chunks" in item:
                    prev["chunks"].extend(item["chunks"])
                    continue
                # Mixed: normalize to chunks
                def to_chunks(run):
                    if "chunks" in run:
                        return run["chunks"]
                    if "text" in run:
                        return [{"type": "text", "text": run["text"]}]
                    return []
                prev_chunks = to_chunks(prev)
                prev["chunks"] = prev_chunks + to_chunks(item)
                prev.pop("text", None)
                continue
        out.append(item)
    return out

def parse_paragraph(p, styles, numbering, theme):
    pPrEl = p.find(qn("w","pPr"))
    declared_pPr = extract_pPr(pPrEl)
    style_id = declared_pPr.get("styleId") or styles.get("defaultParagraphStyleId")
    pstyle_p, pstyle_r = resolve_style_chain(styles, style_id) if style_id else ({}, {})
    num_r = get_numbering_rpr(declared_pPr.get("numPr"), numbering)
    num_info = get_numbering_info(declared_pPr.get("numPr"), numbering)
    para_r = declared_pPr.get("rPr", {})
    # names
    paragraph_style_name = styles.get("styles", {}).get(style_id, {}).get("name") if style_id else None
    visible_alignment = pick_visible_alignment(declared_pPr, pstyle_p)

    content = []
    for child in p:
        if child.tag == qn("w","pPr"):
            continue
        if child.tag == qn("w","r"):
            rPr = extract_rPr(child.find(qn("w","rPr")))
            char_r = {}
            if rPr.get("rStyle"):
                _p, _r = resolve_style_chain(styles, rPr.get("rStyle"))
                char_r = _r
            vis = build_visible_run_rpr(rPr, char_r, para_r, pstyle_r, num_r, theme)
            chunks = collect_run_chunks(child)
            if all(c.get("type") == "text" for c in chunks):
                text = "".join(c.get("text", "") for c in chunks)
                tfp = {"rPr": vis}
                # augment with requested metadata
                if vis.get("rFonts"):
                    tfp["fontName"] = vis["rFonts"].get("ascii") or vis["rFonts"].get("hAnsi")
                if isinstance(vis.get("sizeHalfPoints"), int):
                    tfp["fontSizePt"] = vis["sizeHalfPoints"] / 2
                else:
                    # fallback to 11pt if no size resolved
                    tfp["fontSizePt"] = 11.0
                if style_id:
                    tfp["paragraphStyleId"] = style_id
                if paragraph_style_name:
                    tfp["paragraphStyleName"] = paragraph_style_name
                if rPr.get("rStyle"):
                    tfp["characterStyleId"] = rPr.get("rStyle")
                    tfp["characterStyleName"] = styles.get("styles", {}).get(rPr.get("rStyle"), {}).get("name")
                if num_info.get("present"):
                    tfp["numbering"] = num_info
                if visible_alignment:
                    tfp["alignment"] = visible_alignment
                run_obj = {"type":"run", "text": text, "text_full_properties": tfp}
                content.append(run_obj)
            else:
                tfp = {"rPr": vis}
                if vis.get("rFonts"):
                    tfp["fontName"] = vis["rFonts"].get("ascii") or vis["rFonts"].get("hAnsi")
                if isinstance(vis.get("sizeHalfPoints"), int):
                    tfp["fontSizePt"] = vis["sizeHalfPoints"] / 2
                else:
                    tfp["fontSizePt"] = 11.0
                if style_id:
                    tfp["paragraphStyleId"] = style_id
                if paragraph_style_name:
                    tfp["paragraphStyleName"] = paragraph_style_name
                if rPr.get("rStyle"):
                    tfp["characterStyleId"] = rPr.get("rStyle")
                    tfp["characterStyleName"] = styles.get("styles", {}).get(rPr.get("rStyle"), {}).get("name")
                if num_info.get("present"):
                    tfp["numbering"] = num_info
                if visible_alignment:
                    tfp["alignment"] = visible_alignment
                run_obj = {"type":"run", "chunks": chunks, "text_full_properties": tfp}
                content.append(run_obj)
        elif child.tag == qn("w","hyperlink"):
            runs = []
            for r in child.findall(qn("w","r")):
                rPr = extract_rPr(r.find(qn("w","rPr")))
                char_r = {}
                if rPr.get("rStyle"):
                    _p, _r = resolve_style_chain(styles, rPr.get("rStyle"))
                    char_r = _r
                vis = build_visible_run_rpr(rPr, char_r, para_r, pstyle_r, num_r, theme)
                chunks = collect_run_chunks(r)
                if all(c.get("type") == "text" for c in chunks):
                    text = "".join(c.get("text", "") for c in chunks)
                    tfp = {"rPr": vis}
                    if vis.get("rFonts"):
                        tfp["fontName"] = vis["rFonts"].get("ascii") or vis["rFonts"].get("hAnsi")
                    if isinstance(vis.get("sizeHalfPoints"), int):
                        tfp["fontSizePt"] = vis["sizeHalfPoints"] / 2
                    else:
                        tfp["fontSizePt"] = 11.0
                    if style_id:
                        tfp["paragraphStyleId"] = style_id
                    if paragraph_style_name:
                        tfp["paragraphStyleName"] = paragraph_style_name
                    if rPr.get("rStyle"):
                        tfp["characterStyleId"] = rPr.get("rStyle")
                        tfp["characterStyleName"] = styles.get("styles", {}).get(rPr.get("rStyle"), {}).get("name")
                    if num_info.get("present"):
                        tfp["numbering"] = num_info
                    if visible_alignment:
                        tfp["alignment"] = visible_alignment
                    run_obj = {"type":"run", "text": text, "text_full_properties": tfp}
                    runs.append(run_obj)
                else:
                    tfp = {"rPr": vis}
                    if vis.get("rFonts"):
                        tfp["fontName"] = vis["rFonts"].get("ascii") or vis["rFonts"].get("hAnsi")
                    if isinstance(vis.get("sizeHalfPoints"), int):
                        tfp["fontSizePt"] = vis["sizeHalfPoints"] / 2
                    else:
                        tfp["fontSizePt"] = 11.0
                    if style_id:
                        tfp["paragraphStyleId"] = style_id
                    if paragraph_style_name:
                        tfp["paragraphStyleName"] = paragraph_style_name
                    if rPr.get("rStyle"):
                        tfp["characterStyleId"] = rPr.get("rStyle")
                        tfp["characterStyleName"] = styles.get("styles", {}).get(rPr.get("rStyle"), {}).get("name")
                    if num_info.get("present"):
                        tfp["numbering"] = num_info
                    if visible_alignment:
                        tfp["alignment"] = visible_alignment
                    run_obj = {"type":"run", "chunks": chunks, "text_full_properties": tfp}
                    runs.append(run_obj)
            target = child.get(qn("r","id"))
            content.append({"type":"hyperlink", "target": target, "runs": consolidate_runs(runs)})

    return {"type":"paragraph", "p": {k: v for k, v in declared_pPr.items() if k in ("jc","numPr")}, "content": consolidate_runs(content)}

def parse_table(tbl, styles, numbering, theme):
    rows = []
    for tr in tbl.findall(qn("w","tr")):
        cells = []
        for tc in tr.findall(qn("w","tc")):
            tcPr = {}
            cp = tc.find(qn("w","tcPr"))
            if cp is not None:
                gridSpan = cp.find(qn("w","gridSpan"))
                if gridSpan is not None and gridSpan.get(qn("w","val")):
                    tcPr["gridSpan"] = int(gridSpan.get(qn("w","val")))
                vMerge = cp.find(qn("w","vMerge"))
                if vMerge is not None:
                    tcPr["vMerge"] = vMerge.get(qn("w","val")) or "continue"
            cell_content = []
            for ch in tc:
                if ch.tag == qn("w","p"):
                    cell_content.append(parse_paragraph(ch, styles, numbering, theme))
                elif ch.tag == qn("w","tbl"):
                    cell_content.append(parse_table(ch, styles, numbering, theme))
            cells.append({"content": cell_content, "tcPr": tcPr})
        rows.append({"cells": cells})
    return {"type":"table", "rows": rows}

def parse_document_part(z, part_path, styles, numbering, theme):
    x = read_xml(z, part_path)
    body = x.getroot().find(qn("w","body")) if x.getroot().tag == qn("w","document") else x.getroot()
    blocks = []
    if body is None:
        return blocks
    for el in body:
        if el.tag == qn("w","p"):
            blocks.append(parse_paragraph(el, styles, numbering, theme))
        elif el.tag == qn("w","tbl"):
            blocks.append(parse_table(el, styles, numbering, theme))
    return blocks

def export_docx_to_json(docx_path, json_path):
    with zipfile.ZipFile(docx_path, "r") as z:
        styles = parse_styles(z)
        numbering = parse_numbering(z)
        theme = parse_theme(z)
        blocks = parse_document_part(z, "word/document.xml", styles, numbering, theme)

    out = {"source_file": os.path.abspath(docx_path), "body": blocks}
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

def main():
    if len(sys.argv) != 3:
        print("Usage: python docx2json_runs_only.py <input.docx> <output.json>", file=sys.stderr)
        sys.exit(2)
    docx_path, json_path = sys.argv[1], sys.argv[2]
    if not os.path.isfile(docx_path):
        print(f"Input file not found: {docx_path}", file=sys.stderr)
        sys.exit(2)
    export_docx_to_json(docx_path, json_path)

if __name__ == "__main__":
    main()


