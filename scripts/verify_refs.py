# -*- coding: utf-8 -*-
"""Verify leftover old numbers, empty REF results, and sample REF inline dumps."""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{%s}" % NS["w"]

# Default patterns for leftover chapter numbers (edit for your source chapters)
DEFAULT_OLD_NUM_RE = r"图2\.|表2\.|图4\.|表4\.|\(2\.|\(4\."
DEFAULT_CONTEXT_PHRASES = ["上节", "2.3节", "后续章节"]


def dump_inline(p, label: str) -> str:
    parts = []
    children = list(p)
    i = 0
    while i < len(children):
        c = children[i]
        if c.tag != W + "r":
            i += 1
            continue
        fc = c.find(W + "fldChar")
        if fc is not None and fc.get(W + "fldCharType") == "begin":
            j = i + 1
            instr = ""
            result = ""
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
                        for t in c2.findall(W + "t"):
                            result += t.text or ""
                j += 1
            bm = re.search(r"REF\s+(\S+)", instr)
            if bm or "REF" in instr:
                parts.append(f"{{REF:{bm.group(1) if bm else '?'}={result}}}")
            elif "SEQ" in instr:
                parts.append(f"{{SEQ={result}}}")
            else:
                parts.append(f"{{FLD:{instr.strip()[:40]}={result}}}")
            i = j
            continue
        ts = "".join(t.text or "" for t in c.findall(W + "t"))
        if ts:
            parts.append(ts)
        i += 1
    s = "".join(parts)
    print("====", label, "====")
    print(s[:350])
    print()
    return s


def count_empty_refs(paras) -> int:
    empty = 0
    for p in paras:
        children = list(p)
        i = 0
        while i < len(children):
            c = children[i]
            fc = c.find(W + "fldChar") if c.tag == W + "r" else None
            if c.tag == W + "r" and fc is not None and fc.get(W + "fldCharType") == "begin":
                j = i + 1
                instr = ""
                result = ""
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
                            for t in c2.findall(W + "t"):
                                result += t.text or ""
                    j += 1
                if "REF" in instr and result.strip() == "":
                    empty += 1
                    if empty <= 5:
                        print("EMPTY REF", instr.strip()[:60])
                i = j
            else:
                i += 1
    return empty


def verify(
    docx: Path,
    workdir: Path | None,
    old_num_re: str,
    context_phrases: list[str],
    sample_bookmarks: list[str],
    sample_para_indices: list[int],
) -> None:
    own = workdir is None
    if workdir is None:
        workdir = Path(tempfile.mkdtemp(prefix="verify_refs_"))
    else:
        if workdir.exists():
            shutil.rmtree(workdir)
        workdir.mkdir(parents=True)

    try:
        with zipfile.ZipFile(docx) as z:
            z.extractall(workdir)
        root = etree.parse(str(workdir / "word" / "document.xml")).getroot()
        paras = list(root.xpath("./w:body//w:p", namespaces=NS))

        dumped = set()
        for i, p in enumerate(paras):
            ins = "".join(x.text or "" for x in p.xpath(".//w:instrText", namespaces=NS))
            hit_bm = any(k in ins for k in sample_bookmarks) if sample_bookmarks else False
            hit_idx = i in sample_para_indices
            if hit_bm or hit_idx:
                dump_inline(p, f"p{i}")
                dumped.add(i)

        # If no samples specified, dump first few paras that contain REF
        if not dumped and not sample_bookmarks and not sample_para_indices:
            n = 0
            for i, p in enumerate(paras):
                ins = "".join(x.text or "" for x in p.xpath(".//w:instrText", namespaces=NS))
                if "REF" in ins:
                    dump_inline(p, f"p{i}")
                    n += 1
                    if n >= 3:
                        break

        body = "\n".join(
            "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=NS)) for p in paras
        )
        old_hit = bool(re.search(old_num_re, body)) if old_num_re else False
        ctx_hit = [s for s in context_phrases if s in body]
        print("old nums?", old_hit)
        if old_num_re:
            print("  pattern:", old_num_re)
        print("context leftovers?", bool(ctx_hit), ctx_hit if ctx_hit else "")

        empty = count_empty_refs(paras)
        print("empty REF count:", empty)

        # Fallback SEQ still present?
        fb_seq = 0
        for p in paras:
            in_fb = False
            for a in p.iterancestors():
                if etree.QName(a).localname == "Fallback":
                    in_fb = True
                    break
            if not in_fb:
                continue
            ins = "".join(x.text or "" for x in p.xpath(".//w:instrText", namespaces=NS))
            if re.search(r"SEQ\s+", ins):
                fb_seq += 1
        print("Fallback SEQ remaining:", fb_seq)
    finally:
        if own and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)


def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Verify leftover old numbers, empty REFs, and sample REF dumps."
    )
    p.add_argument(
        "--src",
        type=Path,
        required=True,
        help="Docx to verify. Example: chapter_final.docx",
    )
    p.add_argument(
        "--workdir",
        type=Path,
        default=None,
        help="Optional extract workdir (default: system temp).",
    )
    p.add_argument(
        "--old-num-re",
        default=DEFAULT_OLD_NUM_RE,
        help=f"Regex for leftover old numbers (default: {DEFAULT_OLD_NUM_RE})",
    )
    p.add_argument(
        "--context",
        default=",".join(DEFAULT_CONTEXT_PHRASES),
        help="Comma-separated context phrases that should be gone.",
    )
    p.add_argument(
        "--sample-bm",
        default="Eq_37,Eq_41,Eq_55",
        help="Comma-separated bookmark substrings to dump when present in instrText.",
    )
    p.add_argument(
        "--sample-para",
        default="",
        help="Comma-separated paragraph indices to dump (0-based body //w:p).",
    )
    return p


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = build_argparser().parse_args(argv)
    if not args.src.is_file():
        raise SystemExit(f"Source not found: {args.src}")

    bms = [x.strip() for x in args.sample_bm.split(",") if x.strip()]
    idxs = [int(x.strip()) for x in args.sample_para.split(",") if x.strip()]
    phrases = [x.strip() for x in args.context.split(",") if x.strip()]

    verify(args.src, args.workdir, args.old_num_re, phrases, bms, idxs)


if __name__ == "__main__":
    main()
