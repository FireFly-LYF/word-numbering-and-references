# -*- coding: utf-8 -*-
"""Renumber captions/equations and fix cross-references in a thesis chapter docx.

- Figures/tables -> 图N.x / 表N.x (Word SEQ captions)
- Equations -> (1).. continuous (Word SEQ Eq)
- Body citations keep/become REF fields
- Context phrases that depend on other chapters are neutralized
- Output is a NEW file; source is never overwritten

Edit CONFIG below for chapter-specific mappings before running.
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

# ---------------------------------------------------------------------------
# CONFIG — edit these for your chapter extraction / renumbering task
# ---------------------------------------------------------------------------

# SEQ identifier remaps: old SEQ name -> new SEQ name
SEQ_ID_MAP = [
    (r"SEQ\s+Fig4\b", "SEQ 图5."),
    (r"SEQ\s+图2\.(?=\s|\\|\*|$)", "SEQ 图5."),
    (r"SEQ\s+图4\.(?=\s|\\|\*|$)", "SEQ 图5."),
    (r"SEQ\s+Tbl4\b", "SEQ 表5."),
    (r"SEQ\s+表2\.(?=\s|\\|\*|$)", "SEQ 表5."),
    (r"SEQ\s+表4\.(?=\s|\\|\*|$)", "SEQ 表5."),
    (r"SEQ\s+Eq4\b", "SEQ Eq"),
]

# Target SEQ identifiers after normalize (used for counting / labels)
SEQ_FIG = "图5."
SEQ_TBL = "表5."
SEQ_EQ = "Eq"

# Caption prefix text in runs (图2. / 图4. -> 图5.)
FIG_PREFIX_OLD = ("图2.", "图4.", "图5.")
FIG_PREFIX_NEW = "图5."
TBL_PREFIX_OLD = ("表2.", "表4.", "表5.")
TBL_PREFIX_NEW = "表5."

# Heading renumber map (longest-first applied)
HEADING_MAP = [
    ("4.1.1", "5.2.1"),
    ("4.1.2", "5.2.2"),
    ("4.1", "5.2"),
    ("2.4.1", "5.3.1"),
    ("2.4.2", "5.3.2"),
    ("2.4", "5.3"),
    ("2.5.1", "5.4.1"),
    ("2.5.2", "5.4.2"),
    ("2.5.3", "5.4.3"),
    ("2.5", "5.4"),
]

# Context phrases to neutralize when extracting a standalone chapter
CONTEXT_FIXES = [
    ("上节中所介绍的基于伪正交波形的回波分离技术", "基于伪正交波形的回波分离技术"),
    ("上节中所介绍的空时编码技术", "前述空时编码技术"),
    ("上节中的空时编码技术", "前述空时编码技术"),
    ("根据上节的分析可知", "根据上述分析可知"),
    ("，这与2.3节中的分析相一致", ""),
    ("这与2.3节中的分析相一致。", ""),
    ("，相应解决方案将在后续章节中进行具体分析", ""),
    ("相应解决方案将在后续章节中进行具体分析。", ""),
    ("仿真实验参数如所示，", "仿真实验中，"),
    ("仿真实验参数如所示", "仿真实验中"),
]

# Plain equation numbers to convert into SEQ Eq captions, in order.
# Values are old display strings like "(2.21)".
PLAIN_EQ_ORDER = [f"(2.{i})" for i in range(21, 32)] + [f"(2.{i})" for i in range(45, 54)]

# Plain 式(2.xx) citations -> Eq_N bookmarks (N starts after existing Eq captions)
PLAIN_EQ_CITE_START = 36  # first bookmark Eq_36 for first PLAIN_EQ_ORDER item

# Leftover plain figure cites -> caption bookmark names
PLAIN_FIG_MAP = {
    "图2.22": "_Ref165159254",  # 成像结果中残留干扰能量分析
}

# Broken REF bookmark (missing caption) to drop
BROKEN_TABLE_REF_BM = "_Ref157517732"

# SEQ patterns to strip from mc:Fallback
FALLBACK_SEQ_RE = (
    r"SEQ\s+(图5\.|图2\.|图4\.|图|Fig4|表5\.|表2\.|表4\.|Tbl4|Eq|Eq4)(?=\s|\\|\*|$)"
)

# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=NSMAP))


def is_in_fallback(p) -> bool:
    """True if paragraph lives in mc:Fallback (duplicate VML caption)."""
    for anc in p.iterancestors():
        if etree.QName(anc).localname == "Fallback":
            return True
    return False


def iter_body_paras(root):
    for p in root.xpath("./w:body//w:p", namespaces=NSMAP):
        if is_in_fallback(p):
            continue
        yield p


def max_bookmark_id(root) -> int:
    ids = []
    for b in root.xpath(".//w:bookmarkStart", namespaces=NSMAP):
        try:
            ids.append(int(b.get(qn("id"))))
        except Exception:
            pass
    return max(ids) if ids else 1000


def field_instr_text(field_runs) -> str:
    return "".join(
        (el.text or "")
        for r in field_runs
        for el in r
        if el.tag == qn("instrText")
    )


def iter_paragraph_fields(p):
    """Yield (start_index, end_index_exclusive, runs_slice, instr, result_t_nodes)."""
    children = list(p)
    i = 0
    while i < len(children):
        child = children[i]
        if child.tag != qn("r"):
            i += 1
            continue
        fc = child.find(qn("fldChar"))
        if fc is None or fc.get(qn("fldCharType")) != "begin":
            i += 1
            continue
        j = i + 1
        instr_runs = []
        result_nodes = []
        phase = "instr"  # instr | result
        while j < len(children):
            c = children[j]
            if c.tag == qn("r"):
                fc2 = c.find(qn("fldChar"))
                if fc2 is not None:
                    typ = fc2.get(qn("fldCharType"))
                    if typ == "separate":
                        phase = "result"
                        j += 1
                        continue
                    if typ == "end":
                        j += 1
                        break
                if phase == "instr":
                    instr_runs.append(c)
                elif phase == "result":
                    for t in c.findall(qn("t")):
                        result_nodes.append(t)
            j += 1
        runs = children[i:j]
        instr = field_instr_text(instr_runs)
        yield i, j, runs, instr, result_nodes
        i = j


def _collect_plain_chars(p):
    """Collect characters from this paragraph's own runs (not nested txbx captions).

    Returns list of (ch, t_node, local_index, in_field_result).
    """
    chars = []
    depth = 0
    in_result = False

    def walk_run(r):
        nonlocal depth, in_result
        for el in r:
            if el.tag == qn("fldChar"):
                typ = el.get(qn("fldCharType"))
                if typ == "begin":
                    depth += 1
                    in_result = False
                elif typ == "separate":
                    in_result = True
                elif typ == "end":
                    depth = max(0, depth - 1)
                    in_result = False
            elif el.tag == qn("t") and el.text:
                for li, ch in enumerate(el.text):
                    chars.append((ch, el, li, depth > 0 and in_result))
            elif el.tag == qn("r"):
                walk_run(el)

    for child in p:
        if child.tag == qn("r"):
            walk_run(child)
        elif child.tag == qn("hyperlink"):
            for r in child.findall(qn("r")):
                walk_run(r)
    return chars


def replace_text_in_para(p, old: str, new: str) -> int:
    """Replace first plain-text occurrence of old (never touch field instr/result runs)."""
    chars = _collect_plain_chars(p)
    plain_idx = [i for i, c in enumerate(chars) if not c[3]]
    if not plain_idx:
        return 0
    joined = "".join(chars[i][0] for i in plain_idx)
    idx = joined.find(old)
    if idx < 0:
        return 0

    start_c = plain_idx[idx]
    end_c = plain_idx[idx + len(old) - 1]  # inclusive
    t_start, li_start = chars[start_c][1], chars[start_c][2]
    t_end, li_end = chars[end_c][1], chars[end_c][2]

    if t_start is t_end:
        text = t_start.text or ""
        t_start.text = text[:li_start] + new + text[li_end + 1 :]
        if t_start.text[:1] == " " or t_start.text[-1:] == " ":
            t_start.set(XML_SPACE, "preserve")
        return 1

    # Multi-node: left keeps prefix + new; middle cleared; end keeps suffix
    left = (t_start.text or "")[:li_start]
    right = (t_end.text or "")[li_end + 1 :]
    t_start.text = left + new
    if t_start.text[:1] == " " or t_start.text[-1:] == " ":
        t_start.set(XML_SPACE, "preserve")
    seen = set()
    for i in range(start_c, end_c + 1):
        t = chars[i][1]
        if t is t_start or t is t_end:
            continue
        if t not in seen:
            t.text = ""
            seen.add(t)
    t_end.text = right
    if right and (right[:1] == " " or right[-1:] == " "):
        t_end.set(XML_SPACE, "preserve")
    return 1


def make_ref_field_runs(bookmark: str, display: str, rPr=None):
    elems = []

    def with_rpr(r):
        if rPr is not None:
            r.insert(0, deepcopy(rPr))
        return r

    r = with_rpr(etree.Element(qn("r")))
    fc = etree.SubElement(r, qn("fldChar"))
    fc.set(qn("fldCharType"), "begin")
    elems.append(r)

    r = with_rpr(etree.Element(qn("r")))
    it = etree.SubElement(r, qn("instrText"))
    it.set(XML_SPACE, "preserve")
    it.text = f" REF {bookmark} \\h "
    elems.append(r)

    r = with_rpr(etree.Element(qn("r")))
    fc = etree.SubElement(r, qn("fldChar"))
    fc.set(qn("fldCharType"), "separate")
    elems.append(r)

    r = with_rpr(etree.Element(qn("r")))
    t = etree.SubElement(r, qn("t"))
    t.text = display
    elems.append(r)

    r = with_rpr(etree.Element(qn("r")))
    fc = etree.SubElement(r, qn("fldChar"))
    fc.set(qn("fldCharType"), "end")
    elems.append(r)
    return elems


def _replace_span_with_ref(p, start_idx: int, end_idx: int, bm: str, display: str) -> None:
    """Replace plain-char span [start_idx, end_idx) with a REF field at the same locus."""
    chars = _collect_plain_chars(p)
    if end_idx > len(chars) or start_idx < 0 or start_idx >= end_idx:
        return

    t_start, li_start = chars[start_idx][1], chars[start_idx][2]
    t_end, li_end = chars[end_idx - 1][1], chars[end_idx - 1][2]
    r_start = t_start.getparent()
    if r_start is None or r_start.tag != qn("r"):
        return
    rPr = r_start.find(qn("rPr"))
    rPr_copy = deepcopy(rPr) if rPr is not None else None
    parent = r_start.getparent()
    if parent is None:
        return

    left = (t_start.text or "")[:li_start]
    if t_start is t_end:
        right = (t_start.text or "")[li_end + 1 :]
        t_start.text = left
        if left[:1] == " " or (left and left[-1:] == " "):
            t_start.set(XML_SPACE, "preserve")
        insert_at = list(parent).index(r_start) + 1
        ref_runs = make_ref_field_runs(bm, display, rPr_copy)
        for offset, e in enumerate(ref_runs):
            parent.insert(insert_at + offset, e)
        if right:
            r_right = etree.Element(qn("r"))
            if rPr_copy is not None:
                r_right.append(deepcopy(rPr_copy))
            t = etree.SubElement(r_right, qn("t"))
            if right[:1] == " " or right[-1:] == " ":
                t.set(XML_SPACE, "preserve")
            t.text = right
            parent.insert(insert_at + len(ref_runs), r_right)
        return

    right = (t_end.text or "")[li_end + 1 :]
    t_start.text = left
    if left[:1] == " " or (left and left[-1:] == " "):
        t_start.set(XML_SPACE, "preserve")

    seen = set()
    for i in range(start_idx, end_idx):
        t = chars[i][1]
        if t is t_start or t is t_end:
            continue
        if t not in seen:
            t.text = ""
            seen.add(t)
    t_end.text = right
    if right and (right[:1] == " " or right[-1:] == " "):
        t_end.set(XML_SPACE, "preserve")

    insert_at = list(parent).index(r_start) + 1
    ref_runs = make_ref_field_runs(bm, display, rPr_copy)
    for offset, e in enumerate(ref_runs):
        parent.insert(insert_at + offset, e)


def cleanup_leading_caption_duplicates(root):
    """Neutralize duplicate captions inside mc:Fallback.

    Drawing captions often exist in both mc:Choice and mc:Fallback. Keeping SEQ
    in both makes Word/our counter assign two figure numbers to one figure.
    Strategy: remove SEQ fields + caption bookmarks from Fallback paragraphs,
    leave plain static text.
    """
    stripped_fields = 0
    cleared_bookmarks = 0
    for p in root.xpath("./w:body//w:p", namespaces=NSMAP):
        if not is_in_fallback(p):
            continue
        for bm in list(p.xpath("./w:bookmarkStart|./w:bookmarkEnd", namespaces=NSMAP)):
            name = bm.get(qn("name")) or ""
            if bm.tag == qn("bookmarkStart"):
                if name.startswith("_Toc"):
                    continue
                bid = bm.get(qn("id"))
                p.remove(bm)
                for be in list(p.xpath("./w:bookmarkEnd", namespaces=NSMAP)):
                    if be.get(qn("id")) == bid:
                        p.remove(be)
                cleared_bookmarks += 1
        fields = list(iter_paragraph_fields(p))
        seq_fields = [f for f in fields if re.search(FALLBACK_SEQ_RE, f[3])]
        for start_i, end_i, _runs, _instr, _res in reversed(seq_fields):
            children = list(p)
            result_text = "".join(t.text or "" for t in _res) or "1"
            insert_at = start_i
            for c in children[start_i:end_i]:
                p.remove(c)
            r = etree.Element(qn("r"))
            t = etree.SubElement(r, qn("t"))
            t.text = result_text
            p.insert(insert_at, r)
            stripped_fields += 1
    return stripped_fields, cleared_bookmarks


def normalize_seq_identifiers(root):
    """Unify SEQ identifiers and fix label prefixes in caption paras."""
    seq_changes = 0
    for instr in root.xpath(".//w:instrText", namespaces=NSMAP):
        orig = instr.text or ""
        s = orig
        for pat, repl in SEQ_ID_MAP:
            s = re.sub(pat, repl, s)
        # ensure target fig/tbl stay stable
        s = re.sub(rf"SEQ\s+{re.escape(SEQ_FIG)}(?=\s|\\|\*|$)", f"SEQ {SEQ_FIG}", s)
        if s != orig:
            instr.text = s
            instr.set(XML_SPACE, "preserve")
            seq_changes += 1

    prefix_changes = 0
    for p in iter_body_paras(root):
        has_fig = False
        has_tbl = False
        has_eq = False
        for _i, _j, _runs, instr, _res in iter_paragraph_fields(p):
            if re.search(rf"SEQ\s+{re.escape(SEQ_FIG)}(?=\s|\\|\*|$)", instr):
                has_fig = True
            if re.search(rf"SEQ\s+{re.escape(SEQ_TBL)}(?=\s|\\|\*|$)", instr):
                has_tbl = True
            if re.search(rf"SEQ\s+{re.escape(SEQ_EQ)}\b", instr):
                has_eq = True
        if has_fig:
            for t in p.xpath("./w:r/w:t", namespaces=NSMAP):
                if not t.text:
                    continue
                if t.text in FIG_PREFIX_OLD:
                    if t.text != FIG_PREFIX_NEW:
                        t.text = FIG_PREFIX_NEW
                        prefix_changes += 1
                elif re.match(r"^图[24]\.", t.text):
                    t.text = FIG_PREFIX_NEW + t.text[3:]
                    prefix_changes += 1
        if has_tbl:
            for t in p.xpath("./w:r/w:t", namespaces=NSMAP):
                if not t.text:
                    continue
                if t.text in TBL_PREFIX_OLD:
                    if t.text != TBL_PREFIX_NEW:
                        t.text = TBL_PREFIX_NEW
                        prefix_changes += 1
                elif re.match(r"^表[24]\.", t.text):
                    t.text = TBL_PREFIX_NEW + t.text[3:]
                    prefix_changes += 1
        if has_eq:
            for t in p.xpath("./w:r/w:t", namespaces=NSMAP):
                if t.text == "(4.":
                    t.text = "("
                    prefix_changes += 1
    return seq_changes, prefix_changes


def renumber_existing_eq_captions(root):
    """Set existing Eq SEQ caption results to 1..N and rename Eq_4_k bookmarks."""
    eq_index = 0
    renamed = []
    for p in iter_body_paras(root):
        fields = list(iter_paragraph_fields(p))
        eq_fields = [f for f in fields if re.search(rf"SEQ\s+{re.escape(SEQ_EQ)}\b", f[3])]
        if not eq_fields:
            continue
        text = para_text(p).strip()
        if len(text) > 12 or not re.search(r"\(\d", text):
            continue
        bm_names = [b.get(qn("name")) for b in p.xpath("./w:bookmarkStart", namespaces=NSMAP)]
        if any(re.fullmatch(r"Eq_(3[6-9]|[4-5]\d)", n or "") for n in bm_names):
            continue
        eq_index += 1
        _i, _j, _runs, _instr, result_nodes = eq_fields[0]
        if result_nodes:
            result_nodes[0].text = str(eq_index)
            for rn in result_nodes[1:]:
                rn.text = ""
        for bm in p.xpath("./w:bookmarkStart", namespaces=NSMAP):
            name = bm.get(qn("name")) or ""
            if re.fullmatch(r"Eq_4_\d+", name):
                new_name = f"Eq_{eq_index}"
                bm.set(qn("name"), new_name)
                renamed.append((name, new_name))
    return renamed, eq_index


def make_eq_number_elems(bookmark_name: str, bookmark_id: int, display_num: int):
    elems = []
    bm_s = etree.Element(qn("bookmarkStart"))
    bm_s.set(qn("id"), str(bookmark_id))
    bm_s.set(qn("name"), bookmark_name)
    elems.append(bm_s)

    def t_run(text):
        r = etree.Element(qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text
        return r

    def fld(typ):
        r = etree.Element(qn("r"))
        fc = etree.SubElement(r, qn("fldChar"))
        fc.set(qn("fldCharType"), typ)
        return r

    elems.append(t_run("("))
    elems.append(fld("begin"))
    r = etree.Element(qn("r"))
    it = etree.SubElement(r, qn("instrText"))
    it.set(XML_SPACE, "preserve")
    it.text = f" SEQ {SEQ_EQ} \\* ARABIC "
    elems.append(r)
    elems.append(fld("separate"))
    elems.append(t_run(str(display_num)))
    elems.append(fld("end"))
    elems.append(t_run(")"))
    bm_e = etree.Element(qn("bookmarkEnd"))
    bm_e.set(qn("id"), str(bookmark_id))
    elems.append(bm_e)
    return elems


def convert_plain_equations(root, start_num: int):
    """Convert plain equation paragraphs into SEQ Eq captions."""
    mapping = {}
    n = start_num
    for old in PLAIN_EQ_ORDER:
        mapping[old] = n
        n += 1

    next_id = max_bookmark_id(root) + 1
    converted = []
    for p in list(iter_body_paras(root)):
        text = para_text(p).strip()
        if text not in mapping:
            continue
        if any(re.search(rf"SEQ\s+{re.escape(SEQ_EQ)}\b", f[3]) for f in iter_paragraph_fields(p)):
            continue
        new_num = mapping[text]
        bm_name = f"Eq_{new_num}"
        pPr = p.find(qn("pPr"))
        for child in list(p):
            if child is not pPr:
                p.remove(child)
        for e in make_eq_number_elems(bm_name, next_id, new_num):
            p.append(e)
        next_id += 1
        converted.append((text, f"({new_num})", bm_name))
    return converted


def update_all_seq_results(root):
    """Walk document fields in order; assign continuous display numbers; build bookmark->label map."""
    counters = {SEQ_FIG: 0, SEQ_TBL: 0, SEQ_EQ: 0}
    bm_to_label = {}
    fig_list = []
    tbl_list = []
    eq_list = []

    for p in iter_body_paras(root):
        bm_names = [
            b.get(qn("name"))
            for b in p.xpath("./w:bookmarkStart", namespaces=NSMAP)
            if b.get(qn("name")) and not (b.get(qn("name")) or "").startswith("_Toc")
        ]
        for _i, _j, _runs, instr, result_nodes in iter_paragraph_fields(p):
            m = re.search(
                rf"SEQ\s+({re.escape(SEQ_FIG)}|{re.escape(SEQ_TBL)}|{re.escape(SEQ_EQ)})(?=\s|\\|\*|$)",
                instr,
            )
            if not m:
                continue
            key = m.group(1)
            counters[key] += 1
            num = counters[key]
            if result_nodes:
                result_nodes[0].text = str(num)
                for rn in result_nodes[1:]:
                    rn.text = ""
            if key == SEQ_FIG:
                label = f"{SEQ_FIG}{num}"
                fig_list.append((label, bm_names, para_text(p).strip()[:60]))
            elif key == SEQ_TBL:
                label = f"{SEQ_TBL}{num}"
                tbl_list.append((label, bm_names, para_text(p).strip()[:60]))
            else:
                label = f"({num})"
                eq_list.append((label, bm_names, para_text(p).strip()[:20]))
            for bm in bm_names:
                if bm not in bm_to_label:
                    bm_to_label[bm] = label

    return counters, bm_to_label, fig_list, tbl_list, eq_list


def replace_plain_eq_citations(root, bm_to_label):
    """Replace plain 式(2.xx) with REF to converted Eq bookmarks."""
    old_to_bm = {}
    n = PLAIN_EQ_CITE_START
    for old in PLAIN_EQ_ORDER:
        old_to_bm[old] = f"Eq_{n}"
        n += 1

    done = []
    changed = True
    while changed:
        changed = False
        for p in list(iter_body_paras(root)):
            chars = _collect_plain_chars(p)
            plain_idx = [i for i, c in enumerate(chars) if not c[3]]
            if not plain_idx:
                continue
            joined = "".join(chars[i][0] for i in plain_idx)
            for old, bm in old_to_bm.items():
                pos = 0
                while True:
                    idx = joined.find(old, pos)
                    if idx < 0:
                        break
                    look = joined[max(0, idx - 3) : idx]
                    if "式" not in look:
                        pos = idx + 1
                        continue
                    start_c = plain_idx[idx]
                    end_c = plain_idx[idx + len(old) - 1] + 1
                    disp = bm_to_label.get(bm, f"({bm.split('_')[-1]})")
                    _replace_span_with_ref(p, start_c, end_c, bm, disp)
                    done.append((old, disp, bm))
                    changed = True
                    break
                if changed:
                    break
            if changed:
                break
    return done


def convert_remaining_plain_fig_cites(root, bm_to_label):
    """Convert leftover plain 图x.xx (not yet REF) into caption cross-references."""
    done = []
    changed = True
    while changed:
        changed = False
        for p in list(iter_body_paras(root)):
            chars = _collect_plain_chars(p)
            plain_idx = [i for i, c in enumerate(chars) if not c[3]]
            if not plain_idx:
                continue
            joined = "".join(chars[i][0] for i in plain_idx)
            for old, bm in PLAIN_FIG_MAP.items():
                idx = joined.find(old)
                if idx < 0:
                    continue
                start_c = plain_idx[idx]
                end_c = plain_idx[idx + len(old) - 1] + 1
                disp = bm_to_label.get(bm, old)
                _replace_span_with_ref(p, start_c, end_c, bm, disp)
                done.append((old, disp, bm))
                changed = True
                break
            if changed:
                break
    return done


def flatten_missing_table_ref(root):
    """Broken REF to missing caption -> drop field; fix surrounding text."""
    count = 0
    bm = BROKEN_TABLE_REF_BM
    for p in list(iter_body_paras(root)):
        xml = etree.tostring(p, encoding="unicode")
        if bm not in xml and "表2.1" not in para_text(p):
            continue
        if bm in xml:
            children = list(p)
            new_children = []
            idx = 0
            while idx < len(children):
                child = children[idx]
                if child.tag == qn("r") and child.find(qn("fldChar")) is not None:
                    fc = child.find(qn("fldChar"))
                    if fc is not None and fc.get(qn("fldCharType")) == "begin":
                        j = idx + 1
                        instr = ""
                        while j < len(children):
                            if children[j].tag == qn("r"):
                                it = children[j].find(qn("instrText"))
                                if it is not None:
                                    instr += it.text or ""
                                fc2 = children[j].find(qn("fldChar"))
                                if fc2 is not None and fc2.get(qn("fldCharType")) == "end":
                                    j += 1
                                    break
                            j += 1
                        if bm in instr:
                            count += 1
                            idx = j
                            continue
                        new_children.extend(children[idx:j])
                        idx = j
                        continue
                new_children.append(child)
                idx += 1
            for c in list(p):
                p.remove(c)
            for c in new_children:
                p.append(c)
        if "表2.1" in para_text(p):
            if replace_text_in_para(p, "仿真实验参数如表2.1所示，", ""):
                count += 1
            elif replace_text_in_para(p, "如表2.1所示，", ""):
                count += 1
            elif replace_text_in_para(p, "表2.1", "表"):
                count += 1
    return count


def update_ref_results(root, bm_to_label):
    updated = []
    for p in iter_body_paras(root):
        for _i, _j, _runs, instr, result_nodes in iter_paragraph_fields(p):
            m = re.search(r"REF\s+(\S+)", instr)
            if not m or not result_nodes:
                continue
            bm = m.group(1)
            if bm not in bm_to_label:
                continue
            new_disp = bm_to_label[bm]
            old = "".join(t.text or "" for t in result_nodes)
            if (
                re.match(r"^(图|表)\d+\.\d+$", old)
                or re.match(r"^\(\d+\)$", old)
                or re.match(r"^(图|表)\d+\.\d+", old)
            ):
                result_nodes[0].text = new_disp
                for rn in result_nodes[1:]:
                    rn.text = ""
                updated.append((bm, old, new_disp))
    return updated


def renumber_headings(root):
    done = []
    for p in iter_body_paras(root):
        text = para_text(p).strip()
        if len(text) >= 50:
            continue
        for old, new in HEADING_MAP:
            if re.match(rf"^{re.escape(old)}(\s|$|\u00a0)", text) or text == old:
                if replace_text_in_para(p, old, new):
                    done.append((text, para_text(p).strip()))
                break
    return done


def fix_context(root):
    done = []
    for p in iter_body_paras(root):
        before = para_text(p)
        for old, new in CONTEXT_FIXES:
            if old in before and replace_text_in_para(p, old, new):
                done.append((old, new if new else "[删除]"))
                before = para_text(p)
    return done


def process(src: Path, dst: Path, report_path: Path | None, workdir: Path | None) -> None:
    if src.resolve() == dst.resolve():
        raise SystemExit("ERROR: --dst must differ from --src (never overwrite source)")

    own_workdir = workdir is None
    if workdir is None:
        workdir = Path(tempfile.mkdtemp(prefix="renumber_docx_"))
    else:
        if workdir.exists():
            shutil.rmtree(workdir)
        workdir.mkdir(parents=True)

    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)

        unzip = workdir / "unz"
        unzip.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(dst, "r") as z:
            z.extractall(unzip)

        doc_path = unzip / "word" / "document.xml"
        root = etree.parse(str(doc_path)).getroot()
        report = []

        stripped, cleared_bm = cleanup_leading_caption_duplicates(root)
        report.append(f"0) Fallback题注去SEQ: {stripped} 处, 清除重复书签 {cleared_bm}")

        seq_n, prefix_n = normalize_seq_identifiers(root)
        report.append(f"1) SEQ标识统一: {seq_n} 处; 题注前缀: {prefix_n} 处")

        renamed, eq4_count = renumber_existing_eq_captions(root)
        report.append(f"2) 原公式题注改为(1)..: {eq4_count} 个, 书签重命名 {len(renamed)}")

        converted = convert_plain_equations(root, start_num=eq4_count + 1)
        report.append(f"3) 纯文本公式改为SEQ题注: {len(converted)}")
        for a, b, c in converted:
            report.append(f"   {a} -> {b} ({c})")

        counters, bm_to_label, figs, tbls, eqs = update_all_seq_results(root)
        report.append(
            f"4) SEQ显示值重算: 图{counters[SEQ_FIG]} / 表{counters[SEQ_TBL]} / 公式{counters[SEQ_EQ]}"
        )
        report.append("   图题注顺序:")
        for label, bms, title in figs:
            report.append(f"   {label} {bms[:2]} {title}")
        report.append("   表题注顺序:")
        for label, bms, title in tbls:
            report.append(f"   {label} {bms[:2]} {title}")

        cites = replace_plain_eq_citations(root, bm_to_label)
        report.append(f"5) 正文公式交叉引用改为REF: {len(cites)}")
        for a, b, c in cites:
            report.append(f"   {a} -> {b} [{c}]")

        n_broken = flatten_missing_table_ref(root)
        report.append(f"6) 缺失表题注的引用处理: {n_broken}")

        _c2, bm_to_label, _f, _t, _e = update_all_seq_results(root)

        ref_upd = update_ref_results(root, bm_to_label)
        report.append(f"7) 已有REF显示值按题注更新: {len(ref_upd)}")

        plain_figs = convert_remaining_plain_fig_cites(root, bm_to_label)
        report.append(f"8) 残留纯文本图号改为REF: {len(plain_figs)}")
        for a, b, c in plain_figs:
            report.append(f"   {a} -> {b} [{c}]")

        heads = renumber_headings(root)
        report.append(f"9) 小节编号调整: {len(heads)}")
        for a, b in heads:
            report.append(f"   {a} -> {b}")

        ctx = fix_context(root)
        report.append(f"10) 上下文衔接修改: {len(ctx)}")
        for a, b in ctx:
            report.append(f"   {a} -> {b}")

        etree.ElementTree(root).write(
            str(doc_path), xml_declaration=True, encoding="UTF-8", standalone=True
        )

        if dst.exists():
            dst.unlink()
        with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for f in unzip.rglob("*"):
                if f.is_file():
                    z.write(f, f.relative_to(unzip).as_posix())

        text = "\n".join(report)
        if report_path:
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text(text, encoding="utf-8")
        print(text)
        print(f"\nSaved: {dst}")
    finally:
        if own_workdir and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)


def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Renumber thesis chapter captions/equations and fix REF cross-references."
    )
    p.add_argument(
        "--src",
        type=Path,
        required=True,
        help="Source docx (never overwritten). Example: thesis_chapter.docx",
    )
    p.add_argument(
        "--dst",
        type=Path,
        required=True,
        help="Output docx path (must differ from --src).",
    )
    p.add_argument(
        "--report",
        type=Path,
        default=None,
        help="Optional report text path. Example: renumber_report.txt",
    )
    p.add_argument(
        "--workdir",
        type=Path,
        default=None,
        help="Optional extract workdir (default: system temp).",
    )
    return p


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = build_argparser().parse_args(argv)
    if not args.src.is_file():
        raise SystemExit(f"Source not found: {args.src}")
    process(args.src, args.dst, args.report, args.workdir)


if __name__ == "__main__":
    main()
