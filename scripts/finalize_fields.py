# -*- coding: utf-8 -*-
"""Update Word fields via COM, resync REF display text, then strip Fallback SEQ.

Pipeline step 2 (after renumber_docx.py). Always followed by Fallback strip —
this script calls strip_fallback_seq by default; you can also run it separately.
"""
from __future__ import annotations

import argparse
import importlib.util
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}

# Caption SEQ patterns used when building bookmark -> label map
SEQ_FIG_RE = r"SEQ\s+图\d+\."
SEQ_TBL_RE = r"SEQ\s+表\d+\."
SEQ_EQ_RE = r"SEQ\s+Eq\b"
FIG_LABEL_RE = r"(图\d+\.\d+)"
TBL_LABEL_RE = r"(表\d+\.\d+)"
EQ_LABEL_RE = r"(\(\d+\))"


def qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=NSMAP))


def is_in_fallback(p) -> bool:
    for anc in p.iterancestors():
        if etree.QName(anc).localname == "Fallback":
            return True
    return False


def iter_paragraph_fields(p):
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
        instr_parts = []
        result_nodes = []
        phase = "instr"
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
                    it = c.find(qn("instrText"))
                    if it is not None:
                        instr_parts.append(it.text or "")
                elif phase == "result":
                    for t in c.findall(qn("t")):
                        result_nodes.append(t)
            j += 1
        yield i, j, "".join(instr_parts), result_nodes
        i = j


def word_update_fields(path: Path) -> None:
    import win32com.client as win32

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(path.resolve()), ReadOnly=False, AddToRecentFiles=False)
        doc.Fields.Update()
        for story in doc.StoryRanges:
            try:
                story.Fields.Update()
            except Exception:
                pass
        doc.Save()
        doc.Close(False)
    finally:
        word.Quit()


def resync_ref_from_captions(docx: Path, workdir: Path) -> int:
    if workdir.exists():
        shutil.rmtree(workdir)
    workdir.mkdir(parents=True)
    with zipfile.ZipFile(docx) as z:
        z.extractall(workdir)
    doc_path = workdir / "word" / "document.xml"
    root = etree.parse(str(doc_path)).getroot()

    bm_to_label = {}
    for p in root.xpath("./w:body//w:p", namespaces=NSMAP):
        if is_in_fallback(p):
            continue
        bms = [
            b.get(qn("name"))
            for b in p.xpath("./w:bookmarkStart", namespaces=NSMAP)
            if b.get(qn("name")) and not str(b.get(qn("name"))).startswith("_Toc")
        ]
        tx = para_text(p).strip()
        for _i, _j, instr, result_nodes in iter_paragraph_fields(p):
            if re.search(SEQ_FIG_RE, instr):
                m = re.match(FIG_LABEL_RE, tx)
                label = m.group(1) if m else None
            elif re.search(SEQ_TBL_RE, instr):
                m = re.match(TBL_LABEL_RE, tx)
                label = m.group(1) if m else None
            elif re.search(SEQ_EQ_RE, instr):
                m = re.match(EQ_LABEL_RE, tx)
                label = m.group(1) if m else None
            else:
                continue
            if not label:
                continue
            for bm in bms:
                if bm not in bm_to_label:
                    bm_to_label[bm] = label

    updated = 0
    for p in root.xpath("./w:body//w:p", namespaces=NSMAP):
        if is_in_fallback(p):
            continue
        for _i, _j, instr, result_nodes in iter_paragraph_fields(p):
            m = re.search(r"REF\s+(\S+)", instr)
            if not m or not result_nodes:
                continue
            bm = m.group(1)
            if bm not in bm_to_label:
                continue
            new_disp = bm_to_label[bm]
            old = "".join(t.text or "" for t in result_nodes)
            # Only rewrite pure label results
            if re.match(r"^(图|表)\d+\.\d+$", old) or re.match(r"^\(\d+\)$", old):
                if old != new_disp:
                    result_nodes[0].text = new_disp
                    for rn in result_nodes[1:]:
                        rn.text = ""
                    updated += 1

    etree.ElementTree(root).write(
        str(doc_path), xml_declaration=True, encoding="UTF-8", standalone=True
    )
    out = docx.with_name(docx.stem + "_tmp.zip.docx")
    if out.exists():
        out.unlink()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for f in workdir.rglob("*"):
            if f.is_file():
                z.write(f, f.relative_to(workdir).as_posix())
    shutil.move(str(out), str(docx))
    print("bookmark labels:", len(bm_to_label), "REF updated:", updated)
    return updated


def _load_strip_module():
    """Load sibling strip_fallback_seq.py for post-Word cleanup."""
    sibling = Path(__file__).resolve().parent / "strip_fallback_seq.py"
    if not sibling.is_file():
        return None
    spec = importlib.util.spec_from_file_location("strip_fallback_seq", sibling)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


def process(src: Path, dst: Path, workdir: Path | None, skip_strip: bool) -> None:
    if src.resolve() == dst.resolve():
        # Allow in-place finalize only when user explicitly sets same path —
        # still copy via temp to avoid partial COM writes on the source tree.
        pass

    own_tmp = workdir is None
    if workdir is None:
        workdir = Path(tempfile.mkdtemp(prefix="finalize_fields_"))
    else:
        workdir.mkdir(parents=True, exist_ok=True)

    tmpdoc = workdir / "working.docx"
    try:
        shutil.copy2(src, tmpdoc)
        print("Updating fields in Word...")
        word_update_fields(tmpdoc)
        print("Re-syncing REF displays...")
        resync_ref_from_captions(tmpdoc, workdir / "resync")

        if not skip_strip:
            mod = _load_strip_module()
            if mod is None:
                print(
                    "WARN: strip_fallback_seq.py not found beside this script; "
                    "run strip_fallback_seq.py after finalize (Word may reintroduce Fallback SEQ)."
                )
            else:
                print("Stripping Fallback SEQ (Word may have regenerated them)...")
                mod.strip_fallback_seq(tmpdoc, workdir / "strip_fb")
        else:
            print(
                "Skipped Fallback strip (--skip-strip). "
                "You MUST run strip_fallback_seq.py next."
            )

        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(tmpdoc, dst)
        print("Final saved:", dst)
    finally:
        if own_tmp and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)


def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Word Fields.Update + REF resync + optional Fallback SEQ strip."
    )
    p.add_argument(
        "--src",
        type=Path,
        required=True,
        help="Input docx from renumber_docx.py. Example: chapter_renumbered.docx",
    )
    p.add_argument(
        "--dst",
        type=Path,
        required=True,
        help="Output finalized docx. Example: chapter_final.docx",
    )
    p.add_argument(
        "--workdir",
        type=Path,
        default=None,
        help="Optional work directory (default: system temp).",
    )
    p.add_argument(
        "--skip-strip",
        action="store_true",
        help="Do not call strip_fallback_seq; you must run it yourself after.",
    )
    return p


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = build_argparser().parse_args(argv)
    if not args.src.is_file():
        raise SystemExit(f"Source not found: {args.src}")
    process(args.src, args.dst, args.workdir, args.skip_strip)


if __name__ == "__main__":
    main()
