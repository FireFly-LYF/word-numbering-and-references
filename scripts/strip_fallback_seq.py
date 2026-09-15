# -*- coding: utf-8 -*-
"""Strip SEQ fields from mc:Fallback captions (and duplicate bookmarks).

Must run AFTER Word Fields.Update — Word often regenerates Fallback SEQ,
which double-counts figure/table numbers.
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def is_in_fallback(p) -> bool:
    for a in p.iterancestors():
        if etree.QName(a).localname == "Fallback":
            return True
    return False


def iter_fields(p):
    children = list(p)
    i = 0
    while i < len(children):
        c = children[i]
        if c.tag != W + "r":
            i += 1
            continue
        fc = c.find(W + "fldChar")
        if fc is None or fc.get(W + "fldCharType") != "begin":
            i += 1
            continue
        j = i + 1
        instr = ""
        res = []
        phase = "instr"
        while j < len(children):
            c2 = children[j]
            if c2.tag == W + "r":
                fc2 = c2.find(W + "fldChar")
                if fc2 is not None:
                    typ = fc2.get(W + "fldCharType")
                    if typ == "separate":
                        phase = "result"
                        j += 1
                        continue
                    if typ == "end":
                        j += 1
                        break
                if phase == "instr":
                    it = c2.find(W + "instrText")
                    if it is not None:
                        instr += it.text or ""
                else:
                    res.extend(c2.findall(W + "t"))
            j += 1
        yield i, j, instr, res
        i = j


def strip_fallback_seq(docx: Path, workdir: Path | None = None) -> int:
    """In-place strip of Fallback SEQ inside docx. Returns count stripped."""
    own = workdir is None
    if workdir is None:
        workdir = Path(tempfile.mkdtemp(prefix="strip_fb_"))
    else:
        if workdir.exists():
            shutil.rmtree(workdir)
        workdir.mkdir(parents=True)

    try:
        with zipfile.ZipFile(docx) as z:
            z.extractall(workdir)
        root = etree.parse(str(workdir / "word" / "document.xml")).getroot()

        n = 0
        for p in root.xpath("./w:body//w:p", namespaces=NS):
            if not is_in_fallback(p):
                continue
            for bm in list(p.xpath("./w:bookmarkStart", namespaces=NS)):
                name = bm.get(W + "name") or ""
                if name.startswith("_Toc"):
                    continue
                bid = bm.get(W + "id")
                p.remove(bm)
                for be in list(p.xpath("./w:bookmarkEnd", namespaces=NS)):
                    if be.get(W + "id") == bid:
                        p.remove(be)
            for start, end, instr, res in reversed(list(iter_fields(p))):
                if not re.search(r"SEQ\s+", instr):
                    continue
                result_text = "".join(t.text or "" for t in res) or "1"
                children = list(p)
                for c in children[start:end]:
                    p.remove(c)
                r = etree.Element(W + "r")
                t = etree.SubElement(r, W + "t")
                t.text = result_text
                p.insert(start, r)
                n += 1

        print("stripped fallback SEQ:", n)
        etree.ElementTree(root).write(
            str(workdir / "word" / "document.xml"),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        out = docx.with_name(docx.stem + "_tmp.zip.docx")
        if out.exists():
            out.unlink()
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for f in workdir.rglob("*"):
                if f.is_file():
                    z.write(f, f.relative_to(workdir).as_posix())
        shutil.move(str(out), str(docx))
        print("updated", docx)
        return n
    finally:
        if own and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)


def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Strip SEQ from mc:Fallback captions after Word Fields.Update."
    )
    p.add_argument(
        "--src",
        type=Path,
        required=True,
        help="Docx to patch in place (or copy to --dst first). Example: chapter_final.docx",
    )
    p.add_argument(
        "--dst",
        type=Path,
        default=None,
        help="If set, copy --src to --dst then strip --dst (never overwrite --src).",
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

    target = args.src
    if args.dst is not None:
        if args.src.resolve() == args.dst.resolve():
            raise SystemExit("ERROR: --dst must differ from --src when both are set")
        args.dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(args.src, args.dst)
        target = args.dst

    strip_fallback_seq(target, args.workdir)


if __name__ == "__main__":
    main()
