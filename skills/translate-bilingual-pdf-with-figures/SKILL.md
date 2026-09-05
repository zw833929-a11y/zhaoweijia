---
name: translate-bilingual-pdf-with-figures
description: Translate full academic or scientific PDFs into paragraph-by-paragraph bilingual DOCX documents while preserving original figures, captions, references, equations, and reading order. Use for requests such as “全文翻译”, “逐段中英对照”, “把原图插入对应位置”, or a polished bilingual paper deliverable. Do not use for isolated terminology questions or summaries that do not require a document artifact.
---

# Academic PDF Bilingual Translation With Figures

Turn a scientific PDF into a verified bilingual Word document. Preserve the article's logic and evidence chain, not merely its visible words.

## Deliverable contract

Unless the user specifies otherwise, produce one `.docx` with:

- English source followed immediately by its Chinese translation for every content unit.
- The original title, authors, affiliations, abstract, keywords, main text, acknowledgements, declarations, figure legends, and references.
- Every figure available in the supplied PDF, inserted immediately before its bilingual legend.
- Original numbering for sections, figures, tables, equations, citations, and references.
- A translator's note for a demonstrable source inconsistency; never silently “correct” the paper.
- A short completion report containing quantitative QA results and any source limitations.

Never invent missing supplementary material, unreadable text, data, or images. State the omission explicitly.

## Required workflow

### 1. Lock the scope

Confirm internally:

- Source and target languages.
- Whether “每一段” means semantic paragraphs rather than PDF line fragments. Default to semantic paragraphs.
- Whether references stay source-only. Default to source-only, because translating bibliographic titles changes citation fidelity.
- Whether supplementary files were actually supplied.
- Desired output format. Default to `.docx` because bilingual layout and images are easier to review and edit.

Read the applicable PDF and document-creation instructions before handling the files. Save reusable artifacts durably when the environment supports it.

### 2. Inventory the PDF before translating

Run:

```bash
python3 scripts/extract_pdf_assets.py SOURCE.pdf --output-dir tmp/pdf-extract
```

Inspect the generated manifest, page renders, extracted text, text blocks, and embedded images. Record:

- Page count and whether the PDF is text-based, scanned, or mixed.
- Section order and apparent column layout.
- Counts of main figures, tables, equations, and numbered references.
- Whether figure legends are in the article body or appended.
- Whether supplementary figures named in the text exist in the supplied file.

Use OCR only for pages whose text layer is absent or unusable. Treat OCR output as a draft requiring visual verification.

### 3. Build a source-of-truth content model

Do not translate directly from raw `pdftotext` output. Normalize the paper into an ordered JSON model first; follow [references/model-schema.md](references/model-schema.md).

Model one semantic unit at a time:

- `heading`: title and section headings.
- `paragraph`: prose, list entries, declarations, and table notes.
- `figure`: image path plus source and translated legends.
- `equation`: equation text or a lossless rendered image.

Keep references in a dedicated array. Give every translatable unit a stable unique ID. Preserve the original English text verbatim except for layout artifacts such as line wrapping, discretionary hyphens, and ligature encoding.

Reconstruct reading order visually. For multi-column pages, coordinates and page screenshots outrank extraction order.

### 4. Map figures explicitly

Create a mapping table before assembly:

| Figure | Article location | Image file | Source legend | Status |
|---|---|---|---|---|
| Fig. 1 | after first in-text callout or at legend location | `figures/figure-001.png` | exact legend | verified/missing |

Rules:

- Prefer the complete compound figure over isolated panels or thumbnails.
- Verify panel labels and aspect ratio against the rendered PDF.
- Re-encode problematic JFIF/CMYK images to RGB PNG or JPEG if `python-docx` rejects them.
- Insert images inline, not as floating anchors.
- Keep a figure with its legend when pagination permits. Use an intentional page break for a full-page figure rather than allowing a lone caption fragment.

### 5. Translate scientifically, paragraph by paragraph

Translate from the normalized model, using the full article for context.

- Preserve claims, directionality, modality, negation, dosage, units, symbols, gene/protein names, and group labels.
- Keep abbreviations such as EVs, hUCMSCs, BECN1, PKH26, DiR, and MIF consistent. Define them where the source defines them; do not expand by guesswork.
- Distinguish methods from results and correlation from causation.
- Preserve citation markers with the sentence they support.
- Preserve `×`, `μ`, `°C`, superscripts, subscripts, and italic biological notation where feasible.
- Do not “improve” a surprising result. If text and figure visibly conflict, translate faithfully and add a clearly labeled translator's note.

After translation, compare each pair for omissions rather than checking Chinese fluency alone.

### 6. Build the DOCX from the model

Run:

```bash
python3 scripts/build_bilingual_docx.py \
  --model tmp/bilingual-model.json \
  --output output/article-bilingual.docx
```

The bundled builder is a reusable baseline, not a reason to skip article-specific formatting. Apply:

- Distinct source and translation styles.
- Readable CJK and Latin fonts with embedded-environment fallbacks.
- Consistent margins, spacing, page numbers, and restrained heading hierarchy.
- Source immediately followed by translation; never place all English before all Chinese.
- Original figure followed by English legend and Chinese legend.
- References as a compact source-only section unless the user asks otherwise.

### 7. Run programmatic QA

Run:

```bash
python3 scripts/audit_bilingual_docx.py \
  output/article-bilingual.docx \
  --model tmp/bilingual-model.json
```

The audit must pass, then perform article-specific checks:

- Every modeled content ID appears exactly once in the final document, directly or through its associated figure/legend.
- Every required translation field is nonempty.
- Source/translation pair counts match.
- Figure count and legend count match the verified source inventory.
- Images are inline; anchored/floating images count is zero.
- References are complete and numbered sequentially when the source uses numbered references.
- No `TODO`, placeholder, orphan heading, duplicate paragraph, or missing-image marker remains.
- Header/footer fields include current page and total pages when requested.

Normalize whitespace, Unicode dashes, superscript representations, and ligatures before exact-string coverage checks. A formatting transformation is not automatically a missing paragraph.

### 8. Render and inspect every page

Render the DOCX with the document workflow into a new empty directory for each revision. Inspect every rendered page, not a sample.

Check:

- No blank spill page, clipping, overlap, or orphaned heading.
- CJK glyphs render correctly; no tofu boxes or fallback corruption.
- Figures are legible and paired with the correct legends.
- Page breaks are intentional.
- Headers and footers do not collide with body text.
- The final page is complete.

If a defect appears, fix the model or builder, regenerate, rerun the programmatic audit, and rerender all pages.

## Acceptance standards

| Dimension | Pass standard | Reject when |
|---|---|---|
| Completeness | All in-scope semantic units represented | Any section, legend, declaration, or paragraph is absent |
| Pairing | Each source unit is immediately followed by its translation | English and Chinese are batched or misaligned |
| Fidelity | Numbers, units, directions, qualifiers, and citations agree | Translation strengthens, reverses, or invents a claim |
| Figure integrity | Every supplied main figure is correct, legible, inline, and correctly placed | Wrong panel, cropped image, missing legend, or floating anchor |
| Reference integrity | Count and numbering match source | Gaps, duplicates, or invented entries |
| Scientific typography | Symbols and super/subscripts remain intelligible | Broken units, corrupted glyphs, or ambiguous notation |
| Visual quality | Every rendered page passes inspection | Blank pages, clipping, tofu, overlap, or isolated fragments |
| Transparency | Source defects and missing supplements are disclosed | Silent correction or fabricated material |

## Proven failure modes and fixes

| Failure mode | Why it happens | Proven response |
|---|---|---|
| Raw text has broken words and paragraphs | PDF stores lines/glyphs, not semantic paragraphs | Rebuild paragraphs from coordinates and screenshots; remove only layout hyphens |
| Multi-column text is out of order | Extractor interleaves columns | Sort and group blocks by page geometry, then verify visually |
| Chinese renders as boxes | Office renderer lacks the named CJK font | Install/use an OFL CJK font and configure a font alias for the renderer |
| Cover creates a blank page | Both an explicit break and a heading/page setting force pagination | Use one pagination mechanism; rerender from a clean directory |
| Image is viewable but rejected by `python-docx` | JFIF/CMYK/color-space encoding incompatibility | Decode and save as RGB PNG/JPEG before insertion |
| Exact-string audit reports false omissions | DOCX splits text into runs for superscripts or symbols | Audit normalized paragraph text and accept equivalent caret/superscript forms |
| Old pages survive after regeneration | Render output directory was reused | Render each revision into a new empty directory |
| Source prose conflicts with a figure | The published article itself is inconsistent | Preserve the source and add a neutral translator's note |
| Supplementary figures are named but absent | Only the main article PDF was supplied | State that the supplement was unavailable; never fabricate it |
| Figure extraction yields many tiny assets | Compound figures are stored as several PDF objects | Compare assets with page render and reconstruct/export the complete figure |

## Reuse recipe

For the next similar task:

1. Copy the PDF into a task workspace.
2. Run `extract_pdf_assets.py` and inspect the inventory.
3. Create the ordered JSON model using the reference schema.
4. Fill translations and verified figure mappings.
5. Run `build_bilingual_docx.py`.
6. Run `audit_bilingual_docx.py` until it exits successfully.
7. Render into a fresh directory and inspect every page.
8. Deliver the document with counts, limitations, and a concise handoff note.

Recommended invocation:

> Use $translate-bilingual-pdf-with-figures to translate this scientific PDF paragraph by paragraph into Chinese, insert every original figure at the corresponding position, preserve references and scientific notation, and deliver a visually verified DOCX.

For a concrete completed-project record and handoff template, read [references/handoff.md](references/handoff.md).

## Resource map

- `scripts/extract_pdf_assets.py`: page/text/block/image inventory and extraction.
- `scripts/build_bilingual_docx.py`: model-driven bilingual DOCX baseline.
- `scripts/audit_bilingual_docx.py`: structural and content QA.
- `references/model-schema.md`: canonical intermediate model.
- `references/handoff.md`: completed-case lessons and reusable handoff format.
