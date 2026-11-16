#!/usr/bin/env python3
"""pptx_templater.py - render PowerPoint decks from an existing .pptx used as template.

Self-contained re-implementation and extension of kwlo/python-pptx-templater (MIT),
for current python-pptx (>=0.6.21) and Jinja2 (>=3).

Two ways to build each output slide (they can be mixed in one data file):

  * slide mode  - copy an existing slide of the template that contains Jinja tags
                  ({{ name }}, {% if %}, {% for %}) and render the tags.
  * layout mode - create a new slide from one of the template's slide-master layouts
                  (the theme) and fill its placeholders with text, bullets, images, tables.

Commands:
  python pptx_templater.py inspect    TEMPLATE.pptx [--json]
  python pptx_templater.py render     TEMPLATE.pptx DATA.json|yaml OUTPUT.pptx
                                      [--strict] [--keep-template-slides]
  python pptx_templater.py templatize DECK.pptx OUT_TEMPLATE.pptx OUT_DATA.json
                                      [--slides 0,2,5]

Data format (JSON or YAML):
  {
    "globals": {"company": "Acme"},                       # optional, every slide
    "slides": [
      {"slide": 0, "text": {"name": "Paul"}},              # slide mode ("layoutSlideNum" works too)
      {"layout": "Title and Content",                      # layout mode (name or 0-based index)
       "placeholders": {"title": "Hi {{ company }}",
                        "body": ["point 1", ["sub a", "sub b"], "point 2"]}}
    ]
  }
"""
from __future__ import annotations

import argparse
import copy
import json
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.oxml.ns import qn
except ImportError:  # pragma: no cover
    sys.exit("python-pptx is missing: pip install python-pptx jinja2 pyyaml")
try:
    import jinja2
    from jinja2 import meta
except ImportError:  # pragma: no cover
    sys.exit("jinja2 is missing: pip install python-pptx jinja2 pyyaml")

TAG_RE = re.compile(r"({{.*?}}|{%.*?%}|{#.*?#})", re.S)
SMART = str.maketrans({"“": '"', "”": '"', "‘": "'", "’": "'"})
R_ATTRS = (qn("r:embed"), qn("r:link"), qn("r:id"), qn("r:pict"))
SKIP_RELS = ("slideLayout", "notesSlide")

PH_ALIASES = {
    "title": {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.VERTICAL_TITLE},
    "subtitle": {PP_PLACEHOLDER.SUBTITLE},
    "body": {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.VERTICAL_BODY,
             PP_PLACEHOLDER.VERTICAL_OBJECT},
    "picture": {PP_PLACEHOLDER.PICTURE},
    "table": {PP_PLACEHOLDER.TABLE},
    "chart": {PP_PLACEHOLDER.CHART},
    "footer": {PP_PLACEHOLDER.FOOTER},
    "date": {PP_PLACEHOLDER.DATE},
    "slide_number": {PP_PLACEHOLDER.SLIDE_NUMBER},
}
PH_ALIASES["content"] = PH_ALIASES["body"]
PH_ALIASES["center_title"] = PH_ALIASES["title"]


# --------------------------------------------------------------------------- text
def _fix_quotes(text: str) -> str:
    return TAG_RE.sub(lambda m: m.group(0).translate(SMART), text)


def _balanced(text: str) -> bool:
    return all(text.count(a) == text.count(b) for a, b in (("{{", "}}"), ("{%", "%}"), ("{#", "#}")))


def _has_tag(text: str) -> bool:
    return any(t in text for t in ("{{", "{%", "{#"))


def _iter_text_frames(shapes, with_cells=False):
    """Yield (shape, text_frame[, (row, col)]) for text boxes, placeholders, tables, groups."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_text_frames(shape.shapes, with_cells)
            continue
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            yield (shape, shape.text_frame, None) if with_cells else (shape, shape.text_frame)
        if getattr(shape, "has_table", False) and shape.has_table:
            for r, row in enumerate(shape.table.rows):
                for c, cell in enumerate(row.cells):
                    yield (shape, cell.text_frame, (r, c)) if with_cells else (shape, cell.text_frame)


def _render_str(env, text, ctx):
    if not isinstance(text, str) or not _has_tag(text):
        return text
    return env.from_string(_fix_quotes(text)).render(**ctx)


def _render_paragraph(paragraph, env, ctx):
    runs = list(paragraph.runs)
    if not runs:
        return
    full = "".join(r.text for r in runs)
    if not _has_tag(full):
        return
    if all(_balanced(r.text) for r in runs):
        # every run is self-contained: render run by run, keeps per-run formatting
        for r in runs:
            if _has_tag(r.text):
                r.text = _render_str(env, r.text, ctx)
    else:
        # a tag is split across runs: merge into the first run (keeps its formatting)
        runs[0].text = _render_str(env, full, ctx)
        for r in runs[1:]:
            r._r.getparent().remove(r._r)


def render_slide(slide, env, ctx):
    for _shape, tf in _iter_text_frames(slide.shapes):
        for p in tf.paragraphs:
            _render_paragraph(p, env, ctx)


# -------------------------------------------------------------------- slide mode
def duplicate_slide(prs, index):
    """Append a copy of slide `index` (same layout, shapes and relationships)."""
    src = prs.slides[index]
    dst = prs.slides.add_slide(src.slide_layout)
    for shp in list(dst.shapes):  # drop the layout's empty placeholders
        shp._element.getparent().remove(shp._element)

    rid_map = {}
    for rid, rel in src.part.rels.items():
        if any(k in rel.reltype for k in SKIP_RELS):
            continue
        if rel.is_external:
            rid_map[rid] = dst.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            rid_map[rid] = dst.part.relate_to(rel.target_part, rel.reltype)

    src_tree, dst_tree = src.shapes._spTree, dst.shapes._spTree
    for el in src_tree.iterchildren():
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        new_el = copy.deepcopy(el)
        for node in new_el.iter():
            for attr in R_ATTRS:
                old = node.get(attr)
                if old in rid_map:
                    node.set(attr, rid_map[old])
        dst_tree.append(new_el)

    # keep the slide background if it is defined on the slide itself
    src_bg = src._element.cSld.find(qn("p:bg"))
    if src_bg is not None:
        dst._element.cSld.insert(0, copy.deepcopy(src_bg))
    return dst


def delete_slide(prs, index):
    sld_id_lst = prs.slides._sldIdLst
    sld_id = sld_id_lst[index]
    prs.part.drop_rel(sld_id.rId)
    sld_id_lst.remove(sld_id)


# ------------------------------------------------------------------- layout mode
def _norm(s) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())


def find_layout(prs, key):
    layouts = list(prs.slide_layouts)
    if isinstance(key, int) or (isinstance(key, str) and key.isdigit()):
        i = int(key)
        if not 0 <= i < len(layouts):
            raise ValueError(f"layout {i} does not exist (template has {len(layouts)} layouts, 0-based)")
        return layouts[i]
    for lay in layouts:
        if _norm(lay.name) == _norm(key):
            return lay
    names = ", ".join(repr(l.name) for l in layouts)
    raise ValueError(f"layout {key!r} not found. Available: {names}")


def _ph_type(ph):
    try:
        return ph.placeholder_format.type
    except Exception:  # pragma: no cover
        return None


def find_placeholder(slide, key, used):
    """Match by idx (int / 'idx:N'), exact shape name, or alias (title/subtitle/body/...).
    Aliases may carry an ordinal: body, body2, body3 = 1st/2nd/3rd body placeholder."""
    phs = [p for p in slide.placeholders if p.placeholder_format.idx not in used]
    k = str(key)
    m = re.fullmatch(r"(?:idx:)?(\d+)", k)
    if m:
        for p in phs:
            if p.placeholder_format.idx == int(m.group(1)):
                return p
    for p in phs:
        if _norm(p.name) == _norm(k):
            return p
    m = re.fullmatch(r"([a-z_]+?)(\d*)", k.lower())
    if m and m.group(1) in PH_ALIASES:
        types = PH_ALIASES[m.group(1)]
        cands = [p for p in slide.placeholders if _ph_type(p) in types]
        if m.group(1) == "body" and not cands:  # some layouts only have generic objects
            cands = [p for p in slide.placeholders if _ph_type(p) in PH_ALIASES["body"] | {None}]
        n = int(m.group(2) or 1) - 1
        cands.sort(key=lambda p: (p.top or 0, p.left or 0))
        if 0 <= n < len(cands) and cands[n].placeholder_format.idx not in used:
            return cands[n]
    avail = ", ".join(f"{p.placeholder_format.idx}:{p.name!r}({_ph_type(p)})" for p in slide.placeholders)
    raise ValueError(f"placeholder {key!r} not found on layout. Available: {avail}")


def _flatten_bullets(items, level=0):
    for it in items:
        if isinstance(it, (list, tuple)):
            yield from _flatten_bullets(it, level + 1)
        else:
            yield level, it


def _fill_text(tf, value, env, ctx):
    if isinstance(value, (list, tuple)):
        lines = list(_flatten_bullets(value))
    else:
        lines = [(0, line) for line in str(_render_str(env, value, ctx)).split("\n")]
    tf.clear()  # leaves one empty paragraph
    for i, (level, text) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = str(_render_str(env, text, ctx))
        p.level = min(level, 8)


def _replace_with(slide, ph, maker):
    left, top, width, height = ph.left, ph.top, ph.width, ph.height
    shape = maker(left, top, width, height)
    ph._element.getparent().remove(ph._element)
    return shape


def _fill_picture(slide, ph, path):
    if _ph_type(ph) == PP_PLACEHOLDER.PICTURE and hasattr(ph, "insert_picture"):
        return ph.insert_picture(str(path))  # crops to fill the placeholder

    def make(left, top, width, height):
        pic = slide.shapes.add_picture(str(path), left, top)
        ratio = min(width / pic.width, height / pic.height)
        pic.width, pic.height = int(pic.width * ratio), int(pic.height * ratio)
        pic.left = left + (width - pic.width) // 2
        pic.top = top + (height - pic.height) // 2
        return pic
    return _replace_with(slide, ph, make)


def _fill_table(slide, ph, rows, env, ctx):
    rows = [list(r) for r in rows]
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    if _ph_type(ph) == PP_PLACEHOLDER.TABLE and hasattr(ph, "insert_table"):
        table = ph.insert_table(n_rows, n_cols).table
    else:
        table = _replace_with(slide, ph, lambda l, t, w, h: slide.shapes.add_table(
            n_rows, n_cols, l, t, w, min(h, 370840 * n_rows))).table
    for r, row in enumerate(rows):
        for c in range(n_cols):
            val = row[c] if c < len(row) else ""
            table.cell(r, c).text = "" if val is None else str(_render_str(env, val, ctx))
    return table


def build_layout_slide(prs, spec, env, ctx, base_dir: Path, keep_empty=False):
    slide = prs.slides.add_slide(find_layout(prs, spec["layout"]))
    used = set()
    for key, value in (spec.get("placeholders") or {}).items():
        ph = find_placeholder(slide, key, used)
        used.add(ph.placeholder_format.idx)
        if isinstance(value, dict) and "image" in value:
            path = Path(_render_str(env, value["image"], ctx))
            path = path if path.is_absolute() else base_dir / path
            if not path.is_file():
                raise ValueError(f"image not found: {path}")
            _fill_picture(slide, ph, path)
        elif isinstance(value, dict) and "table" in value:
            _fill_table(slide, ph, value["table"], env, ctx)
        elif isinstance(value, dict) and "text" in value:
            _fill_text(ph.text_frame, value["text"], env, ctx)
        else:
            _fill_text(ph.text_frame, value, env, ctx)
    if not keep_empty:  # remove unfilled placeholders so no "Click to add text" remains
        for ph in list(slide.placeholders):
            if ph.placeholder_format.idx not in used:
                ph._element.getparent().remove(ph._element)
    return slide


# ---------------------------------------------------------------------------- data
def load_data(path: Path):
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() in (".yaml", ".yml"):
        try:
            import yaml
        except ImportError:
            sys.exit("YAML data needs pyyaml: pip install pyyaml")
        return yaml.safe_load(text)
    return json.loads(text)


def _slide_index(spec, n_template):
    for key in ("layoutSlideNum", "slide", "template_slide"):
        if key in spec:
            idx = int(spec[key])
            break
    else:
        raise ValueError(f"slide spec needs 'slide'/'layoutSlideNum' or 'layout': {spec!r}")
    if not 0 <= idx < n_template:
        raise ValueError(f"template slide {idx} does not exist (template has {n_template} slides, 0-based)")
    return idx


# ------------------------------------------------------------------------ commands
def cmd_render(args):
    data_path = Path(args.data)
    data = load_data(data_path)
    if not isinstance(data, dict) or not isinstance(data.get("slides"), list):
        sys.exit("data must be an object with a 'slides' list")
    env = jinja2.Environment(
        undefined=jinja2.StrictUndefined if args.strict else jinja2.Undefined,
        autoescape=False, keep_trailing_newline=True,
    )
    prs = Presentation(args.template)
    n_template = len(prs.slides)
    globals_ = data.get("globals") or {}

    for i, spec in enumerate(data["slides"]):
        ctx = {**globals_, **(spec.get("text") or {}), "slide_number": i + 1}
        try:
            if "layout" in spec:
                build_layout_slide(prs, spec, env, ctx, data_path.parent,
                                   keep_empty=spec.get("keep_empty_placeholders", False))
            else:
                render_slide(duplicate_slide(prs, _slide_index(spec, n_template)), env, ctx)
        except (jinja2.TemplateError, ValueError) as e:
            sys.exit(f"error in output slide {i + 1}: {e}")

    if not args.keep_template_slides:
        for _ in range(n_template):
            delete_slide(prs, 0)

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    prs.save(args.output)
    print(f"wrote {args.output}: {len(data['slides'])} slide(s)")


def cmd_inspect(args):
    env = jinja2.Environment()
    prs = Presentation(args.template)
    slides = []
    for i, slide in enumerate(prs.slides):
        item = {"index": i, "layout": slide.slide_layout.name, "shapes": [], "variables": set()}
        for shape, tf in _iter_text_frames(slide.shapes):
            text = tf.text
            if not text.strip():
                continue
            item["shapes"].append({"name": shape.name, "text": text})
            if _has_tag(text):
                try:
                    item["variables"] |= meta.find_undeclared_variables(env.parse(_fix_quotes(text)))
                except jinja2.TemplateSyntaxError as e:
                    item.setdefault("errors", []).append(f"{shape.name}: {e}")
        item["variables"] = sorted(item["variables"])
        slides.append(item)

    layouts = []
    for i, lay in enumerate(prs.slide_layouts):
        layouts.append({"index": i, "name": lay.name, "placeholders": [
            {"idx": p.placeholder_format.idx, "name": p.name,
             "type": str(_ph_type(p)).split(" ")[0].split(".")[-1]}
            for p in lay.placeholders
            if _ph_type(p) not in PH_ALIASES["footer"] | PH_ALIASES["date"] | PH_ALIASES["slide_number"]
        ]})

    if args.json:
        print(json.dumps({"slides": slides, "layouts": layouts}, ensure_ascii=False, indent=2))
        return
    print(f"{args.template}: {len(slides)} slide(s), {len(layouts)} layout(s), "
          f"size {prs.slide_width}x{prs.slide_height} EMU")
    print("\n== SLIDES (slide mode: {\"slide\": N}) ==")
    for item in slides:
        print(f"\n[slide {item['index']}] layout='{item['layout']}' variables={item['variables']}")
        for s in item["shapes"]:
            print(f"  - {s['name']}: {s['text']!r}")
        for err in item.get("errors", []):
            print(f"  ! syntax error {err}")
    print("\n== LAYOUTS (layout mode: {\"layout\": NAME}) ==")
    for lay in layouts:
        phs = ", ".join(f"{p['idx']}:{p['name']!r}[{p['type']}]" for p in lay["placeholders"])
        print(f"[layout {lay['index']}] {lay['name']!r}: {phs or '(no placeholders)'}")


def _slug(s: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")
    return s or "shape"


def cmd_templatize(args):
    prs = Presentation(args.deck)
    keep = None
    if args.slides:
        keep = {int(x) for x in args.slides.split(",") if x.strip() != ""}
    out_slides = []
    for i, slide in enumerate(prs.slides):
        if keep is not None and i not in keep:
            continue
        values, seen = {}, {}
        for shape, tf, cell in _iter_text_frames(slide.shapes, with_cells=True):
            paras = [p for p in tf.paragraphs if "".join(r.text for r in p.runs).strip()]
            if not paras:
                continue
            base = _slug(shape.name)
            if cell is not None:
                base += f"_r{cell[0]}c{cell[1]}"
            n = seen.get(base, 0)
            seen[base] = n + 1
            if n:
                base += f"_{n + 1}"
            for j, p in enumerate(paras):
                var = f"s{i}_{base}" + (f"_p{j + 1}" if len(paras) > 1 else "")
                runs = list(p.runs)
                values[var] = "".join(r.text for r in runs)
                runs[0].text = "{{ " + var + " }}"
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
        out_slides.append({"slide": i, "text": values})
    if keep is not None:  # drop slides that were not templatized
        for i in reversed(range(len(prs.slides))):
            if i not in keep:
                delete_slide(prs, i)
        for new_i, s in enumerate(out_slides):
            s["slide"] = new_i
    prs.save(args.out_template)
    Path(args.out_data).write_text(json.dumps({"slides": out_slides}, ensure_ascii=False, indent=2),
                                   encoding="utf-8")
    n_vars = sum(len(s["text"]) for s in out_slides)
    print(f"wrote {args.out_template} ({len(out_slides)} slide(s), {n_vars} variable(s)) "
          f"and {args.out_data}")


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    pi = sub.add_parser("inspect", help="list slides + variables, and layouts + placeholders")
    pi.add_argument("template")
    pi.add_argument("--json", action="store_true")
    pi.set_defaults(func=cmd_inspect)
    pr = sub.add_parser("render", help="render a deck from template + data")
    pr.add_argument("template")
    pr.add_argument("data")
    pr.add_argument("output")
    pr.add_argument("--strict", action="store_true", help="fail on undefined variables")
    pr.add_argument("--keep-template-slides", action="store_true",
                    help="keep the original template slides at the start of the deck")
    pr.set_defaults(func=cmd_render)
    pt = sub.add_parser("templatize", help="turn an existing deck into a template + data JSON")
    pt.add_argument("template", metavar="DECK")
    pt.add_argument("out_template")
    pt.add_argument("out_data")
    pt.add_argument("--slides", help="comma-separated 0-based slide indexes to keep (default: all)")
    pt.set_defaults(func=cmd_templatize, deck=None)
    args = p.parse_args(argv)
    if not Path(args.template).is_file():
        sys.exit(f"file not found: {args.template}")
    if args.cmd == "templatize":
        args.deck = args.template
    args.func(args)


if __name__ == "__main__":
    main()
