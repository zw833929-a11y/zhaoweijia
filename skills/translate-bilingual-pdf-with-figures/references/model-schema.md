# Bilingual Article Model

Use one UTF-8 JSON file as the source of truth for ordering, translation coverage, figure placement, and document generation.

## Minimal example

```json
{
  "metadata": {
    "title_en": "Example article title",
    "title_zh": "示例文章标题",
    "authors": "A. Author, B. Author",
    "journal": "Example Journal",
    "doi": "10.0000/example",
    "source_file": "article.pdf",
    "source_pages": 12,
    "translator_note": ""
  },
  "items": [
    {
      "id": "h-abstract",
      "type": "heading",
      "level": 1,
      "en": "Abstract",
      "zh": "摘要"
    },
    {
      "id": "p-abstract-001",
      "type": "paragraph",
      "en": "The original English paragraph.",
      "zh": "对应的中文译文。"
    },
    {
      "id": "fig-1",
      "type": "figure",
      "number": "1",
      "image": "figures/figure-001.png",
      "caption_en": "Fig. 1. Original legend.",
      "caption_zh": "图1．对应的中文图注。",
      "page_break": false
    },
    {
      "id": "eq-1",
      "type": "equation",
      "en": "E = mc^2",
      "zh": "E = mc^2"
    }
  ],
  "references": [
    "[1] First reference.",
    "[2] Second reference."
  ]
}
```

## Field rules

### `metadata`

All fields are optional to the builder, but provide as many as the source supports. Never infer a DOI, journal, author, date, or affiliation that is not present or independently verified.

### `items`

The array order is the document order. Every item needs a stable, unique `id` and one of these types:

| Type | Required fields | Notes |
|---|---|---|
| `heading` | `id`, `type`, `en`, `zh` | `level` defaults to 1 and should be 1–3 |
| `paragraph` | `id`, `type`, `en`, `zh` | One semantic paragraph, not one extracted line |
| `figure` | `id`, `type`, `image`, `caption_en`, `caption_zh` | `number` and `page_break` are optional |
| `equation` | `id`, `type`, `en`, `zh` | Use lossless text when possible; otherwise use a figure item |

Paths in `image` are resolved relative to the JSON file. Absolute paths also work but make the model less portable.

### `references`

Store one complete reference per entry. Preserve the source's numbering and bibliographic language. If the PDF combines a numbered entry across multiple lines, join those lines before adding the entry.

## Normalization policy

Allowed source cleanup:

- Join lines inside the same paragraph.
- Remove a line-break hyphen only when the word is demonstrably continuous.
- Normalize nonsemantic whitespace.
- Convert broken PDF ligatures to the intended characters.
- Represent superscripts with Unicode or caret notation if lossless rich text is unavailable.

Not allowed:

- Rephrasing the source English.
- Silently fixing a scientific contradiction.
- Moving a citation to a different claim.
- Guessing an unreadable value or panel label.
- Creating a record for supplementary content that was not supplied.

## Pre-build validation checklist

- IDs are unique.
- Every non-figure item has nonempty `en` and `zh`.
- Every figure file exists and matches its figure number.
- Every figure has both legends.
- Items follow the verified article reading order.
- Reference count and numbering match the source.
- There are no `TODO`, `TBD`, dummy text, or unresolved image paths.
