# Conversion Checklist

## Inputs

- Source PDF path
- Requested page range, using the PDF viewer's page numbering unless the user specifies printed page numbers
- Reference PPTX path, only when the user explicitly provides one for calibration. Normal runtime generation must use the skill's internal design contract, not a copied reference deck.
- Output PPTX path
- Whether to include problem-only slides, solution slides, answer summary slides, or all of them
- Before sharing or after changing the skill package, run `../scripts/self_check_skill.py --report <self-check-report.json>` and require `"passes": true`.
- For ordinary page-range conversion, run `../scripts/export_llm_extraction_packet.py --pdf <input.pdf> --pages <start-end> --output-dir <packet-dir>`, inspect the rendered page images, and write `<packet-dir>/llm-ir.json` as LLM-only IR.
- Build with `../scripts/build_concept_practice_deck_from_llm_ir.py --pdf <input.pdf> --ir <packet-dir>/llm-ir.json --output <output.pptx> --report-dir <report-dir>`. This is the default build runner; it loads LLM-authored `PracticeBlock` IR, applies `../assets/design/style-map.json`, injects built-in OOXML style parts, writes editable grouped slides with reveal animations, and runs the standard validation chain.
- The LLM-IR build runner writes `generation-report.json`, `generation-report-validation.json`, `editable-report.json`, `ooxml-style-report.json`, `pattern-report.json`, and `build-summary.json`; require `"passes": true` from `build-summary.json`.
- Do not use deterministic PDF text extraction or embedded page plans. If a page family needs special handling, express it in LLM IR and renderer/design-map rules.
- After generation, inspect the generator report's `quality_summary` and `slide_trace`. `quality_summary.warning_count` is not an automatic failure, but each warning must be checked against the produced deck before delivery. `quality_summary.normalizations` should explain intentional corrections, such as row-major ordering for two-column source extraction.

## PDF Extraction

1. Inspect the page range visually before bulk processing.
2. Search for headings such as `개념익히기`, `개념 익히기`, `개념 확인`, or a known book-specific variant.
3. Render target pages to images and have the LLM extract the IR directly from the images. Do not create conversion IR through deterministic PDF text parsing.
4. For scanned PDFs, use the same LLM-image workflow, then reconstruct the slide as editable text, math text, and shapes.
5. Use image crops only for atomic diagrams, graphs, QR-like icons, or illustrations; never use a full problem block crop as the slide body unless the user explicitly asks for image-only slides.
6. Split each `개념 익히기` block into prompt and subproblems before layout.
7. Record the source layout for each problem block: one-column vertical list, two-column grid, table, worked stack, number line, diagram layout, or mixed layout. Keep row baselines, left formula column, and right answer column relationships when reconstructing.
8. Preserve per-item source line structure in the inventory. If a subproblem spans several source lines, keep those lines with the item and render them as multiple editable math rows inside one reveal group. Do not split one-line reference-planned expressions merely because they contain several equals signs.
9. Mark whether each source problem contains actual square blanks and record each blank's semantic anchor, such as "after equals sign" or "exponent position." For every visible square blank, put exactly one blank token in item `raw_text`: `[ ]` for baseline blanks and `^[ ]` for exponent blanks. If `blank_count` is present, it must match the combined blank token count. If a blank contains red guide-answer text, keep the blank shape but remove the red text. Do not create square blank shapes for problems that do not have them in the PDF.
10. For number-line source graphics, record the geometry as a `display_segments` entry with `kind: "number_line"`. Use the formal semicolon spec: `min`, `max`, `point`, `direction=left|right`, `closed=true|false`, `ticks`, `labels`, and `blank=true`. Example: `{"kind":"number_line","text":"min=-4;max=1;point=-3;direction=right;closed=false;ticks=-3;labels=-3","line_index":0}`. For a blank student-marking axis, use `blank=true` and omit `point`, `direction`, and `closed`.
11. For number-line items, keep `raw_text` semantic, such as `(1) 수직선: 열린 점 -3, 오른쪽 영역` or `(5) x≥4`; do not encode the drawing as typed arrows or square blank tokens. Set `blank_count` to `0` unless the source also contains actual square blanks outside the number line.
12. When extraction is suspicious, re-open the page image and revise `llm-ir.json`. Do not switch to deterministic item extraction to resolve ambiguity.
13. Exclude `개념 다지기`, `개념 마무리`, and `학교 시험 준비하기` unless explicitly requested.
14. Keep a trace table in the report: problem id, source page, extraction source, source layout type, source lines, editable slide objects created, and slide number. Do not show source page numbers on slide canvas.
15. Confirm report slide traces include actual slide numbers, chunk indexes, item numbers, default-visible item numbers, and click-reveal item numbers.

## Design Contract

1. Read `design-contract.md` before changing the generator architecture.
2. Treat `../assets/design/style-map.json` as the runtime source of common slide coordinates, fonts, colors, and animation intent.
3. Do not reference local expected-output files or raw reference PPTX files from runtime scripts or skill instructions.
4. Add measured values to the design map when repeated source-layout families need tighter fidelity.

## Output Quality

- Do not mix unrelated design styles.
- Match the PDF source problem-body layout before applying reference slide item grids. If a source problem is a vertical one-column list, keep one item per row with source-like row spacing and the answer parentheses aligned in a right-side answer column.
- Hide source answer/example text in problem-only decks. Treat red filled-in values, red decimals, and red answer words such as `무한소수` as answers, not as problem text.
- Do not use full-page screenshots or large problem-block crops when the reference style expects editable text and grouped reveal objects.
- Do not include non-`개념익히기` exercises unless explicitly requested.
- Exclude `개념 다지기` and `개념 마무리` pages unless the user explicitly asks for them.
- Check that every slide has enough contrast and no overlapping text.
- Confirm that media files are embedded and not linked to missing local paths.
- Confirm problem slides contain editable text/shape content, not only a generated PNG/JPEG.
- Confirm on-slide reference-style reveal animations target groups and have no missing `p:spTgt` ids.
- Confirm on-slide square blanks are shape objects, not typed `□` characters.
- Confirm square blank shapes appear only where the source PDF had actual square blanks.
- Confirm `[ ]` blank shapes match the reference deck style: no-fill rounded rectangle, about `0.551in x 0.551in`, and `31750` EMU theme `bg2` outline with `lumMod=75000`, round cap, and round join. Baseline blanks inside stacked-fraction denominators should remain full-size; only explicit exponent blanks `^[ ]` should be compact.
- Confirm exponent-position blanks, such as `a^[ ]`, `b^[ ]`, or `(b³)^[ ]`, render as compact raised blanks with no visible caret. Confirm operator-following blanks, such as `×[ ]`, remain normal baseline blanks.
- When using `--visual-refine-loop`, confirm `visual-refinement-report.json` records iteration PNGs, diagnostics, pass/fail, and any adjustment values.
- Confirm the generator report has no quality warnings. If present, inspect the page images and correct `llm-ir.json` before delivery.
- Confirm square blank shapes sit at the same formula position as the PDF, especially blanks between `=`, `÷`, and the final `=` in conversion exercises.
- Confirm square blank shapes that contained red guide answers in the PDF are empty in the generated problem slide.
- Confirm red source answers or example answers are not visible unless the requested output includes answers or solution slides.
- Confirm source page numbers are not visible on slide canvas.
- Confirm no generated slide object sits fully or partially outside the slide canvas unless the user explicitly requested an intentional bleed object.
- Confirm prompt text uses `나눔스퀘어라운드 ExtraBold` when available, formulas use `BT수식M` when available, and fractions use horizontal fraction bars when the PDF shows stacked fractions.
- Confirm visible source marker shapes are represented as `display_segments` marker entries and rendered as editable shapes. In unit 1, the marker before `분모의 소인수:` should be a reference-style right arrow with `bg1` fill and `lumMod=50000`, not a typed arrow character or omitted text decoration.
- Confirm meaningful source horizontal gaps are represented in IR with `display_segments[].gap_after_in` and rendered as real object spacing, especially when a final `=` is separated from a far-right parenthesis answer blank.
- Confirm source number lines are represented as `display_segments` entries with `kind: "number_line"` and render as editable shapes: axis, arrowheads, tick marks, labels, optional interval band, optional endpoint dot, and optional interval arrow. Open endpoints must be unfilled circles; closed endpoints must be filled circles; blank axes should not show a point or interval band.
- Confirm number-line segment specs preserve the source interval direction and labels. For example, an open dot at `-3` shaded to the right must use `point=-3;direction=right;closed=false`, while an open dot at `5` shaded to the left must use `point=5;direction=left;closed=false`.
- Confirm every stacked fraction has enough width for the longest numerator/denominator and uses no-wrap text boxes; `19/10` must render as `19` over `10`, not as `1` over `9` over `0`.
- Confirm stacked fractions follow the reference PPTX geometry: numeric fractions use the measured digit-slot widths, multiplicative denominators such as `2²×3×5` stay in one denominator textbox with superscript runs, and denominator text/blank boxes sit close to the fraction bar. Denominator blanks with exponents, such as `11/([ ]²×[ ]²)` and `1/[ ]³`, must keep full-size blanks and place the exponent as a separate upper-right textbox.
- Confirm prompt/subproblem spacing, line spacing, and item spacing follow the source PDF layout where it affects readability, while retaining the reference deck's visual style and animation pattern.
- Confirm two-column worked-solution blocks render and reveal in visual row-major order, not raw PDF text extraction order.
- Confirm dense two-column worked-solution blocks with multiline items are split by visual row when needed, and continuation slides reset the row y-position instead of starting halfway down the slide.
- Confirm multiline worked-solution items are grouped as one reveal step containing all rows for that subproblem.
- Confirm the generator report has no unexplained quality warnings for missing item numbers, duplicate item numbers, unknown layout, empty prompts, missing source lines, or slide trace count mismatches.
- Confirm `validate_generation_report.py` passes, proving that requested pages are fully accounted for, included/excluded page sets do not overlap, slide traces cover every inventory item, chunk indexes are contiguous, and default/reveal item order is internally consistent.
- Render or visually inspect the final deck before delivery; fix issues and provide only the final corrected PPTX.
- Confirm fully off-slide helper/template objects from the reference were ignored during analysis and removed from generated slides.
- Confirm `self_check_skill.py` passes before committing or sharing the skill folder; it verifies required files, JSON validity, Python compile status, asset/style-part manifest hashes, and absence of removed runtime dependencies.
- Save the final deck as `.pptx`.
