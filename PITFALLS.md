# PITFALLS — Word caption SEQ / REF renumbering

Real failures from a thesis chapter extraction pipeline. Read before changing OOXML edit logic.

## 1. mc:Choice / mc:Fallback dual captions

Drawing captions often exist **twice**: once under `mc:Choice` (modern drawing) and once under `mc:Fallback` (VML). Both can contain `SEQ` fields → figure/table counters **double-count** (one figure becomes 图5.1 and 图5.2).

**Fix:** Strip `SEQ` (and duplicate caption bookmarks) from Fallback; keep Choice only. Leave static plain digits in Fallback if needed for readability.

**Critical:** Re-strip **after** Word `Fields.Update` — Word may regenerate Fallback `SEQ`.

## 2. xpath `//w:p` includes nested drawing paragraphs

`./w:body//w:p` walks into textboxes / drawings. Counting or editing must **skip Fallback**. Do not naively concatenate `.//w:t` on a body paragraph that contains a drawing — nested caption text gets mixed into the body string.

Prefer: own-run walks (`_collect_plain_chars`) or top-level `./w:r/w:t` where appropriate; always `is_in_fallback(p)`.

## 3. `replace_text_in_para` collapsing all runs

Putting the full concatenated `w:t` into the first run and clearing the rest **destroys interleaved REF field results**. Symptoms: body loses `图x.x` display; empty/orphaned REF fields pile up at paragraph end.

**Fix:** Only edit **plain (non-field)** text. For multi-node spans, only clear/touch nodes inside the match span; never rewrite the whole paragraph as one run.

## 4. `_replace_span_with_ref` inserting at wrong locus

Inserting a REF **before the entire run** when the match is mid-run yields:

`REF + 满足式中` instead of `满足式 + REF + 中`.

**Fix:** Split left/right text around the match inside the run; insert REF **between** left and right.

## 5. Plain-char index pollution

Joining **all** characters including field results shifts search indices. Matches land on the wrong XML nodes.

**Fix:** Search/replace only on plain (non-field-result) characters; map joined indices back via `plain_idx`.

## 6. Broken REF to missing caption

Example: `表2.1` / bookmark `_Ref…` with **no caption** in the extracted chapter.

**Fix:** Drop the field; neutralize awkward leftover wording like 「如所示」 / 「仿真实验参数如所示」. Do not invent a fake table number.

## 7. Never overwrite source

Always write a new `--dst` docx. Keep the original for diff / rollback. COM update should also work on a copy.

## 8. Cross-refs must be Word REF fields

Do not “fix” citations by plain-text number swaps alone. Body cites must be `REF <bookmark> \h` pointing at caption bookmarks, with display format matching the original style (`图5.1`, `式(37)` / `(37)`, `表5.2`, …).

## 9. Fields.Update reintroduces Fallback SEQ

After `finalize_fields.py` / Word COM update, Fallback `SEQ` often returns.

**Always** run `strip_fallback_seq.py` after finalize (or use finalize without `--skip-strip`).

## 10. Context phrases vs technical text

When extracting a standalone chapter, neutralize glue phrases that assume other chapters:

- 上节 / 前述章节衔接
- 2.3节 / 后续章节 等跨章指向

Only change **context**, not technical claims. Do **not** casually rewrite figure/table numbers in prose without going through SEQ/REF + caption bookmarks.
