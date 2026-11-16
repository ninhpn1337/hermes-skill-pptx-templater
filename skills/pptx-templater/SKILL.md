---
name: pptx-templater
description: Generate PowerPoint (.pptx) decks from any existing .pptx used as the template/theme, plus JSON/YAML data. Fill slide-master layouts (title, bullets, images, tables), or copy designed slides containing Jinja tags ({{ name }}, {% for %}). Can also turn an existing deck into a reusable template automatically.
version: 1.1.0
author: Based on kwlo/python-pptx-templater
license: MIT
platforms: [macos, linux, windows]
metadata:
  hermes:
    tags: [Productivity, PowerPoint, PPTX, Presentation, Template, Theme, Jinja, Report]
    related_skills: []
---

# PPTX Templater

Build decks that look exactly like an existing PowerPoint file. The template `.pptx`
supplies the design (theme, slide master, layouts, fonts, colours, backgrounds, logos);
a JSON/YAML file supplies the content. Based on
[kwlo/python-pptx-templater](https://github.com/kwlo/python-pptx-templater); the bundled
script reimplements it for current python-pptx/Jinja2 and adds layout mode and templatize.

Three tools in one script:

| Command | Use it when |
|---|---|
| `render` + **layout mode** | Template has a nice theme. Build new slides from its master layouts ("Title and Content", "Two Content", ...) and fill placeholders. No editing of the template needed. |
| `render` + **slide mode** | Template has fully designed slides with `{{ tags }}` in the text. Copy a slide N times and fill the tags. |
| `templatize` | User has a finished deck and wants to reuse it as a template. It replaces every text with `{{ variables }}` and writes the data JSON holding the original text. |

## When to Use

- "Make a deck / report / proposal using our company template (this .pptx)".
- The same slide design must repeat with different content (one slide per customer/item).
- "Reuse this presentation for next month / another client" -> `templatize`, edit JSON, `render`.
- Do NOT use to invent a brand-new visual design; the look always comes from the template file.

## Quick Reference

```bash
pip install python-pptx jinja2 pyyaml          # once
S=${HERMES_SKILL_DIR}/scripts/pptx_templater.py

python "$S" inspect template.pptx              # slides + variables, layouts + placeholders
python "$S" inspect template.pptx --json
python "$S" render template.pptx data.json out.pptx --strict
python "$S" templatize deck.pptx tpl.pptx data.json [--slides 0,2,5]
```

Data file (JSON or YAML); slide entries of both modes can be mixed, output follows list order:

```json
{
  "globals": {"company": "Acme", "quarter": "Q4"},
  "slides": [
    {"layout": "Title Slide",
     "placeholders": {"title": "{{ company }} report", "subtitle": "{{ quarter }}\n2026"}},
    {"layout": "Title and Content",
     "placeholders": {"title": "Highlights",
                      "body": ["Revenue +12%", ["EU +20%", "US +5%"], "Churn -2%"]}},
    {"layout": "Two Content",
     "placeholders": {"title": "Before / after", "body": ["old flow"], "body2": {"image": "chart.png"}}},
    {"layout": "Title and Content",
     "placeholders": {"title": "Numbers", "body": {"table": [["Region", "Sales"], ["EU", "1.2M"]]}}},
    {"slide": 0, "text": {"name": "Paul", "items": ["a", "b"]}}
  ]
}
```

**Layout mode** (`"layout"`: layout name, case/space-insensitive, or 0-based index):
- `placeholders` keys: alias `title`, `subtitle`, `body`/`content` (`body2`, `body3` = 2nd/3rd,
  ordered top-to-bottom, left-to-right), `picture`, `table`; or the exact placeholder name
  (`"Content Placeholder 2"`); or its idx (`"1"`, `"idx:1"`). `inspect` lists them.
- Values: string (`\n` = new paragraph) | list of bullets (nested list = indent level) |
  `{"image": "path.png"}` (relative to the data file) | `{"table": [[row], ...]}`.
  All strings are Jinja-rendered with `globals` + the slide's `text` + `slide_number`.
- Unfilled placeholders are removed (no "Click to add text"); keep them with
  `"keep_empty_placeholders": true`.

**Slide mode** (`"slide"` or original key `"layoutSlideNum"`: 0-based index of a template SLIDE):
- The slide is copied with everything on it (images, shapes, links, background) and Jinja tags
  in text boxes, placeholders, tables and groups are rendered with `globals` + `text`.
- Full Jinja: `{{ x | default("-") }}`, `{% if %}`, `{{ items | join(", ") }}`, filters.

Template slides are removed from the output unless `--keep-template-slides`.

## Procedure

1. Check deps: `python -c "import pptx, jinja2"`, else `pip install python-pptx jinja2 pyyaml`.
2. Get the template .pptx from the user. If there is none, generate the demo:
   `python ${HERMES_SKILL_DIR}/scripts/make_example_template.py demo.pptx`
   (pairs with `${HERMES_SKILL_DIR}/examples/example.json`).
3. `inspect` it. Decide per output slide:
   - the theme has a suitable layout -> layout mode;
   - a designed slide with `{{ }}` tags fits -> slide mode;
   - the user wants to reuse a finished deck -> `templatize` first, then edit the JSON values
     (and duplicate/reorder entries) and render with slide mode.
4. Write the data file using the exact layout names / placeholder keys / variables from `inspect`.
5. `render ... --strict` (fails loudly on missing variables, unknown layouts/placeholders,
   missing images; the message lists what is available).
6. Verify (below), then give the user the output path (`[[as_document]]` to send as a file).

## Pitfalls

- **Look comes from the template.** A plain default template gives a plain deck. If the user
  wants it nicer, ask for their branded .pptx or a theme file saved as .pptx.
- **Layout names differ per template and language** (e.g. Vietnamese Office: "Tiêu đề và
  Nội dung"). Always copy names from `inspect`, or use the layout index.
- **Placeholder types**: many themes use OBJECT (content) placeholders; `body` matches them.
  A picture placeholder crops the image to fill it; other placeholders fit the image inside.
- **Tag split across runs (slide mode)**: PowerPoint may split `{{ name }}` into several runs.
  The script merges that paragraph into its first run, so mixed formatting inside that
  paragraph becomes the first run's style. Retype the tag in one go to avoid it.
- **Smart quotes** inside tags are normalised automatically.
- **Block tags** (`{% if %}`, `{% for %}`) must open and close within one paragraph; a loop
  cannot create bullets. For bullets use layout mode lists.
- **templatize** keeps the first run's formatting of each paragraph; per-word styling
  (one bold word inside a sentence) is flattened. Charts, SmartArt and notes are not templatized.
- Charts are not filled with data; charts on copied slides are shared between copies.
- Speaker notes are not copied.
- Don't `pip install python-pptx-templater`: it pins jinja2 2.10 / python-pptx 0.6.18, which
  break on modern Python. Use the bundled script.

## Verification

- `render` prints `wrote OUT: N slide(s)`; N must equal the number of data entries.
- `python "$S" inspect out.pptx`: every slide shows `variables=[]` and the expected text;
  any leftover `{{` means a tag was not rendered.
- Optional visual check: `soffice --headless --convert-to pdf out.pptx`.
