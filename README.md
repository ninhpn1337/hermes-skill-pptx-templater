# Hermes skill: pptx-templater

A [Hermes Agent](https://hermes-agent.nousresearch.com/) skill that builds PowerPoint decks
from any existing `.pptx` (its theme, layouts and designed slides) plus JSON/YAML data.

- **Layout mode**: new slides from the template's slide-master layouts. Fills title, bullets
  (nested), images and tables.
- **Slide mode**: copies designed slides that contain Jinja tags (`{{ name }}`, `{% if %}`) and renders them.
- **templatize**: turns a finished deck into a template plus a data JSON.

Based on [kwlo/python-pptx-templater](https://github.com/kwlo/python-pptx-templater) (MIT).
Rewritten for current python-pptx / Jinja2.

## Install

```bash
hermes skills tap add ninhpn1337/hermes-skill-pptx-templater
hermes skills install pptx-templater
```

Or copy `skills/pptx-templater/` to `~/.hermes/skills/productivity/pptx-templater/`.

The skill needs `pip install python-pptx jinja2 pyyaml`.

## Use without Hermes

```bash
python skills/pptx-templater/scripts/pptx_templater.py inspect template.pptx
python skills/pptx-templater/scripts/pptx_templater.py render template.pptx data.json out.pptx --strict
python skills/pptx-templater/scripts/pptx_templater.py templatize deck.pptx tpl.pptx data.json
```

See [skills/pptx-templater/SKILL.md](skills/pptx-templater/SKILL.md) for the data format.
