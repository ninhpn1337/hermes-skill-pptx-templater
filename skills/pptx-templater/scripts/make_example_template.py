#!/usr/bin/env python3
"""Create a small demo template with Jinja tags (matches examples/example.json).

Usage: python make_example_template.py OUTPUT.pptx
  slide 0: title "Hello {{ name }}"
  slide 1: title "Good bye {{ dog.name }}", body loops over {{ items }}
"""
import sys

from pptx import Presentation

out = sys.argv[1] if len(sys.argv) > 1 else "example_template.pptx"
prs = Presentation()

s0 = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide
s0.shapes.title.text = "Hello {{ name }}"
s0.placeholders[1].text = "{{ company | default('python-pptx-templater demo') }}"

s1 = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
s1.shapes.title.text = "Good bye {{ dog.name }}"
s1.placeholders[1].text = "Slide {{ slide_number }}{% if items is defined and items %}: {{ items | join(', ') }}{% endif %}"

prs.save(out)
print(f"wrote {out}")
