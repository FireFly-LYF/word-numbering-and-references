---
name: word-numbering-and-references
description: >-
  Renumbers Word docx figure/table/equation captions (SEQ fields), rebuilds
  cross-reference REF fields, strips duplicate mc:Fallback SEQ, and extracts
  standalone thesis chapters with consistent numbering. Also covers MathType
  formula layout: TAB+OLE+TAB right-aligned numbers, vertical centering of
  equation numbers (baseline + lower OLE — never textAlignment=center), and
  tab stops matched to page content width. Use when the user asks about 题注,
  交叉引用, 图表编号, 公式编号, 公式编号居中/右对齐, docx numbering, SEQ/REF
  fields, caption renumbering, or pulling a chapter out of a Chinese thesis
  Word file.
---

# Word numbering and cross-references

Private skill for **firefly-lyf**. Scripts live in `scripts/`. Hard pitfalls: read [PITFALLS.md](PITFALLS.md) before editing OOXML.

## When to use

- Renumber 图/表/公式 after extracting a standalone chapter from a thesis
- Fix broken or stale Word `REF` cross-references to caption bookmarks
- Stop double-counted figure numbers caused by `mc:Choice` / `mc:Fallback` dual captions
- Fix equation **number on the far right** and **vertically centered** vs tall MathType (no tables)
- Never overwrite the source docx — always write `--dst`

## Required order

```text
1. renumber_docx.py      # OOXML: SEQ normalize, captions, plain→REF, context fixes
2. finalize_fields.py    # Word COM Fields.Update + REF display resync (+ strip)
3. strip_fallback_seq.py # MUST after Word update (Word may regenerate Fallback SEQ)
4. verify_refs.py        # leftover old nums / empty REF / sample dumps
```

If `finalize_fields.py` runs without `--skip-strip`, step 3 is already done; still safe to re-run strip.

## Agent workflow (must follow)

1. **Read CONFIG** at the top of `scripts/renumber_docx.py` (`HEADING_MAP`, `CONTEXT_FIXES`, `PLAIN_FIG_MAP`, `SEQ_*`, `PLAIN_EQ_ORDER`). Edit for the target chapter — do not hardcode machine paths.
2. **Copy source → work dst**; never overwrite `--src`.
3. Run pipeline in order above. Prefer absolute paths via argparse.
4. After any `Fields.Update`, **re-strip Fallback SEQ** (see PITFALLS #1, #9).
5. When replacing body text or inserting REF: only touch **plain (non-field) characters**; never collapse all `w:t` into one run (PITFALLS #3–#5).
6. Cross-refs must be real `REF bookmark` fields with display matching caption format (`图5.1`, `式(37)` / `(37)`, etc.) — not plain number swaps (PITFALL #8).
7. Equation layout: keep `TAB|OLE|TAB|number` (PITFALL #11); vertical center via **baseline + lower OLE** — **never** `textAlignment=center` for MathType (PITFALL #12); tab stops = one center + one right at page content width (PITFALL #13).
8. Run `verify_refs.py`; fix leftovers before declaring done.
9. On failures, open [PITFALLS.md](PITFALLS.md) first — known broken patterns are documented there.

## Scripts

| Script | Role |
|--------|------|
| `scripts/renumber_docx.py` | Main OOXML pipeline; `--src` `--dst` `--report` |
| `scripts/finalize_fields.py` | Word COM update + REF resync; calls strip unless `--skip-strip` |
| `scripts/strip_fallback_seq.py` | Remove SEQ (+ duplicate bookmarks) from `mc:Fallback` |
| `scripts/verify_refs.py` | Sanity checks and sample REF inline dumps |

Install: `pip install -r requirements.txt` (needs desktop Word + pywin32 for finalize).

## Minimal commands

```bash
python scripts/renumber_docx.py --src chapter.docx --dst chapter_renumbered.docx --report report.txt
python scripts/finalize_fields.py --src chapter_renumbered.docx --dst chapter_final.docx
python scripts/strip_fallback_seq.py --src chapter_final.docx
python scripts/verify_refs.py --src chapter_final.docx
```

## Safe editing rules (short)

- Skip `mc:Fallback` paragraphs when counting SEQ / updating body text.
- Do not use naive `p.xpath(".//w:t")` on body paras that contain drawings (nested caption text leaks in).
- `replace_text_in_para` / `_replace_span_with_ref`: plain-only char index; split left/right around mid-run matches.
- Missing caption bookmarks: drop broken REF; neutralize awkward 「如所示」 wording — do not invent numbers.
- Context phrases (上节 / 后续章节 / x.x节): neutralize only chapter-glue wording, not technical content.
- Formula numbers: far-right via tabs (PITFALL #11/#13); vertical mid via baseline + OLE `w:position` lower (PITFALL #12) — no tables unless the user asks.
