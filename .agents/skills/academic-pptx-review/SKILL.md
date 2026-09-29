---
name: academic-pptx-review
description: Analyze an existing academic PPTX before polishing it. Read slide objects and rendered pages, then report evidence-based comprehension problems or say no change is needed. Use for review and diagnosis, not editing.
---

# Academic PPTX review: read, then judge

Use this skill when the user asks to review an existing academic presentation. This stage produces analysis only. Do not edit the input PPTX.

## Work

1. Read the repository README and its 11 tips. Treat the tips as prompts for judgment, not automatic scores. A dense slide, a table, or a colorful slide can be appropriate if its task is clear.
2. Run `python analyze_pptx.py INPUT.pptx --out OUTPUT_DIR`. Inspect `inventory.json`, all PNGs in `renders/`, and the preliminary `diagnosis.json`. The preliminary report contains **cues**, not confirmed quality problems.
3. For every slide, compare the rendered page with its object list: title/body, text boxes and paragraphs, table cells, chart series/categories/values, pictures, shapes/groups, positions, notes, and hidden status. The slides are read only when visible elements are accounted for. Record any mismatch or failed render by slide.
4. Judge the slide's role in the talk and the effect on a listener. Prefer a few consequential findings. Record each confirmed issue with page, object path or region, observed evidence, audience impact, severity, confidence, related tip numbers, and a modification direction. Distinguish uncertain interpretations from confirmed problems. Say **无需修改** when the page communicates well enough.
5. Write a review JSON following [the usage guide](../../../docs/academic-pptx-review.md), binding it to the input and rendered PNG hashes. Re-run with `--review REVIEW.json`. Read the final Markdown report and verify every claim against its PNG and object path.

## Evidence boundaries

- Object data and visual render are separate sources. Notes, hidden shapes, embedded metadata, and source scripts are not slide-visible evidence.
- A native chart's cached values are usable when the extractor reads them. The apparent trend of a bitmap chart cannot supply exact numbers. Mark unreadable axes, citations, image text, SmartArt internals, or embedded objects as **未知** unless a visible review or OCR identifies them, with source and confidence.
- Do not use a test deck's positive/negative labels or its generating script as the diagnosis. Review the slide first; use paired examples afterward as a check for missed problems or overcriticism.
- If parsing or rendering fails, name the affected slides and narrow the conclusion to what was actually read. Never turn an unreviewed preliminary cue into a confirmed issue.
- This skill does not rewrite slides, validate the underlying paper's scientific truth, or claim PowerPoint-perfect rendering from LibreOffice output.
