# Design Contract

This skill should generate concept-practice PPTX decks from PDF source pages using a stable internal design contract rather than a runtime reference PPTX.

## Pipeline

1. Source PDF and requested pages are rendered to page images for LLM inspection.
2. The LLM writes `llm-ir.json`, including target `개념 익히기` blocks and excluded requested pages.
3. Each LLM-authored block is normalized to `PracticeBlock` IR:
   - source page
   - concept number and title
   - practice number
   - prompt
   - ordered subproblem items
   - source extraction lines for reporting/debugging
   - layout type, currently `vertical_list`, `two_column_grid`, `worked_stack`, or `unknown`
   - item inventory with problem number, expression text, square blank count, parenthesis blank flag, fraction/exponent flags, source PDF bounding box, row index, and column index
   - per-item square blanks represented as `[ ]` baseline tokens or `^[ ]` exponent tokens; when `blank_count` is present, it must agree with the combined blank token count
   - per-item `source_lines` preserving the PDF line structure for multiline rows
   - per-item `display_lines` derived from those source lines so separate PDF lines become separate editable PPTX text/math rows
   - per-item `display_segments` typed as `math`, `marker`, or `korean_label` so mixed rows can use the correct font and shapes per segment; segments may include `gap_after_in` when source-PDF visual spacing must be preserved
4. The renderer receives only the LLM-authored IR. Deterministic PyMuPDF item extraction is not part of this skill package.
5. The renderer applies `assets/design/style-map.json` to create 16:9 editable PPTX slides.
6. The math renderer converts common slash fractions and superscript text into editable text boxes, fraction bars, and blank shapes.
7. Style-only OOXML assets are injected after PPTX creation: theme, slide masters, slide layouts, presentation properties, view properties, and table styles. Generated title slides are redirected to the built-in title layout, and generated problem slides are redirected to the built-in problem layout while preserving actual slide content and slide order. Backgrounds and content panels come from those master/layout parts, not from duplicate slide-local overlay shapes.
8. Reveal animations are injected as group-targeted `p:set` timing nodes.
9. Validation checks editable text, embedded OOXML style parts, group animation, no visible page labels, no off-slide objects, and blank-shape requirements.
10. Report validation checks that requested pages are accounted for, included/excluded page sets are consistent, slide traces cover the normalized item inventory, pagination chunks are contiguous, and quality warnings are resolved or explicitly allowed.
11. The standard LLM-IR build runner writes all validation reports and fails unless generation-report, editable-deck, OOXML-style, and reference-pattern validation all pass.
12. Skill package self-check validates required files, JSON files, Python source compilation, bundled asset hashes, and absence of removed runtime dependencies before sharing.

## Runtime Design Source

- `assets/design/style-map.json` is the common runtime style source.
- `assets/concept-practice-style/` contains extracted reusable title/header images and style-only OOXML from prior reference analysis. These assets are general concept-practice style assets, not a page-range-specific runtime dependency.
- `assets/concept-practice-style/ooxml-style-parts/` is applied at runtime to generated PPTX packages. It must be validated against `style-parts-manifest.json` before sharing the skill.
- Generated slides must rely on the injected master/layout backgrounds and panels. Do not add full-canvas background rectangles or large problem-panel shapes directly to slide XML when the style parts already provide them.
- Raw reference PPTX files must not be required at runtime.
- Local expected-output folders are evaluation aids only and must not be referenced by skill instructions or runtime scripts.

## Extension Rules

- Add measured values to `assets/design/style-map.json` when a recurring layout family is found.
- Add fields to the LLM IR schema before adding renderer-specific hacks.
- Preserve source-visible square blanks as `[ ]` or `^[ ]` tokens in `raw_text`. Every visible square blank must map to exactly one IR blank token and one rendered blank shape.
- Use LLM-only extraction to classify ambiguous structure, such as exponent blanks, same-row fragments, or unreadable math normalization. Do not add deterministic PyMuPDF item extraction to create conversion IR.
- Preserve source PDF layout type first: one-column list, two-column grid, chained equality, table, diagram, or mixed layout.
- Preserve item-level source line structure before rendering. Multiline rows should keep `source_lines`, derive `display_lines`, and split those lines into typed `display_segments` before rendering. Math segments should stay editable math, and Korean label segments should use the same Korean font as prompt text. One-line embedded/reference-planned rows should remain inline even when they contain several equals signs.
- Prompt text should not be manually wrapped by character count or written with hard line breaks. Render prompts as one editable text string in the measured reference textbox width so PowerPoint performs the visible wrapping.
- Preserve visible source marker shapes with `display_segments` marker entries. In the unit 1 reference deck, the marker before `분모의 소인수:` is a `rightArrow` auto shape, about `0.400in x 0.292in`, theme `bg1` fill with `lumMod=50000`, no outline.
- Preserve meaningful source horizontal gaps with `display_segments[].gap_after_in`. Use this for rows where an answer parenthesis area is visually separated from the equation, such as `3/4=3÷4=` followed by a far-right `(        )`; do not encode that gap only as repeated spaces in `raw_text`.
- For two-column blocks, render and animate items in visual row-major order after extraction normalizes row and column indexes.
- For two-column blocks that have no source boxes and no explicit right-column items, derive grid positions from display order (`column = index % 2`, `row = index // 2`) instead of treating the default `column_index=0` as authoritative.
- For two-column worked blocks with multiline items, paginate by visual row to avoid overcrowding. Each continuation slide should normalize its local row index so the first row starts at the standard first-row y position.
- Reports must include the normalized block/item inventory so layout and math-extraction failures can be audited without opening the PPTX.
- Reports must include actual problem slide numbers, block chunk indexes, item numbers, default-visible item numbers, and click-reveal item numbers so animation order and pagination can be audited without opening the PPTX.
- Reports must include a `quality_summary` with source block/item counts, title/problem slide counts, source blank/fraction/exponent/multiline counts, intentional normalizations, and non-fatal warnings for suspicious extraction conditions.
- Math rendering should stay editable: use text boxes, rectangle bars, superscript runs, and blank shapes instead of equation screenshots.
- Reference `[ ]` blanks are full-size no-fill rounded rectangles, including blanks inside stacked-fraction denominators. Use the measured style-map values: about `0.551in x 0.551in`, line width `31750` EMU, theme `bg2` outline with `lumMod=75000`, round cap, and round join.
- When a denominator blank has a visible exponent, render it as a full-size denominator blank plus a separate small exponent textbox at the blank's upper-right. In the unit 1 reference deck, `[ ]²` and `[ ]³` place the exponent about `0.112in` above the blank top and about `0.023in` after the blank's right edge; the multiplication sign sits on the blank midline rather than on the numerator baseline.
- Stacked fraction geometry should follow the bundled reference PPTX measurements in `assets/design/style-map.json`: centered numerator/denominator text boxes, a horizontal bar, fixed digit-slot widths for numeric fractions, wider multiplicative denominator widths for `×` and exponent runs, and denominator boxes placed close to the fraction bar rather than symmetrically far below it.
- Exponent-position blanks must be written as `^[ ]` in LLM IR and render as smaller raised blank shapes with no visible caret. Baseline blanks after operators remain full-size `[ ]` blanks. Do not depend on `a[ ]` context inference.
- Visual refinement loop reports must include source PNGs, PPTX slide PNGs, comparison PNGs, diagnostics, pass/fail, and any design adjustments for each iteration.
- Inline numeric math tokens should use the style-map numeric width allowance (`math.inline_numeric_width_scale`, `math.inline_numeric_width_padding`, and `math.inline_numeric_min_w`) after the base character-width estimate. This keeps bold `BT수식M` decimal text such as `0.15` from exceeding its editable text box.
- Keep formulas editable. Use atomic images only for diagrams or figures that are not practical to redraw.
- If a generated deck differs from the expected visual pattern, update the design map and renderer rather than making the skill depend on a reference deck copy.
