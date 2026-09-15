# word-numbering-and-references

Word docx **caption SEQ** renumbering + **REF cross-references** for Chinese thesis chapter extraction.

中文：从学位论文整本中抽出独立章节时，重排图/表/公式题注编号，并修复正文交叉引用（Word 域），避免 mc:Fallback 双题注导致编号翻倍。

Private skill for **firefly-lyf**.

## Install

```bash
pip install -r requirements.txt
```

Needs desktop **Microsoft Word** + `pywin32` for `finalize_fields.py`.

## Pipeline order

1. `scripts/renumber_docx.py` — OOXML renumber (never overwrites `--src`)
2. `scripts/finalize_fields.py` — Word `Fields.Update` + REF display resync (+ strip Fallback SEQ)
3. `scripts/strip_fallback_seq.py` — strip Fallback SEQ again if Word regenerated them
4. `scripts/verify_refs.py` — leftover old numbers / empty REF / sample dumps

```bash
python scripts/renumber_docx.py --src chapter.docx --dst chapter_renumbered.docx --report report.txt
python scripts/finalize_fields.py --src chapter_renumbered.docx --dst chapter_final.docx
python scripts/strip_fallback_seq.py --src chapter_final.docx
python scripts/verify_refs.py --src chapter_final.docx
```

Edit `CONFIG` dicts at the top of `renumber_docx.py` before running (chapter maps, context phrase fixes, plain fig map).

See [SKILL.md](SKILL.md) for agent workflow and [PITFALLS.md](PITFALLS.md) for known OOXML failure modes.

## Formula number layout (MathType)

- Keep `TAB | OLE | TAB | (n)` — do not wipe the paragraph when converting numbers (PITFALL #11).
- **Vertical center:** use `textAlignment=baseline` and **lower the OLE** with `w:position`; do **not** use `textAlignment=center` (pins number to top) (PITFALL #12).
- **Horizontal right:** one center tab + one right tab at page content width; avoid dual center tabs (PITFALL #13).
- Prefer not to use tables for this layout unless the user explicitly asks.
