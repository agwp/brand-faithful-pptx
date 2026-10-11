# Template Audit

Use this reference when a template is unfamiliar or the requested slide type is ambiguous.

## Audit steps

1. Run:

   ```bash
   python scripts/inspect_template.py template.pptx --format markdown
   ```

   If the inferred placeholder roles need template-specific corrections, create a reviewed JSON override file and run:

   ```bash
   python scripts/inspect_template.py template.pptx --format markdown \
     --role-overrides template-placeholder-roles.json
   ```

2. Review each text placeholder's inferred `role`, `class`, `basis`, and `review` value. Resolve every low-confidence or review-required entry before generation. Placeholder names and geometry are evidence, not proof that a field must be populated.

3. Identify layouts with:
   - A title placeholder.
   - A chart placeholder for chart-heavy slides.
   - Content placeholders with useful geometry when no chart placeholder exists.
   - Section divider, agenda, comparison, and appendix layouts if the task needs them.

4. Record the selected layout names, placeholder `idx` values, reviewed roles, and content classes in your working notes before generating slides.

## Placeholder role audit

The inspector reports likely roles for placeholders. These roles support review; they do not authorize content.

| Role | Intended content | Default class |
| --- | --- | --- |
| `title` | Slide or section title | `core_content` |
| `body_content` | Main body or multi-purpose content | `core_content` or `review` |
| `subtitle` | Supplied supporting title text | `optional_supporting` |
| `section_label` | Supplied section, category, eyebrow, or navigation label | `optional_supporting` |
| `caption` | Factual identification or explanation of an image | `optional_supporting` |
| `source` | Evidence note, source, citation, or credit | `optional_supporting` |
| `attribution` | Author, speaker, or quotation attribution | `optional_supporting` |
| `quote` | Supplied quotation or testimonial | `optional_supporting` |
| `callout` | Supported takeaway, key message, or implication | `optional_supporting` |
| `footer_or_admin` | Header, footer, date, slide number, or administrative field | `administrative` |
| `non_text` | Chart, picture, table, media, or other non-text content | `non_text` |
| `unknown` | Ambiguous text placeholder requiring review | `review` |

`core_content` means the placeholder is a likely destination for primary content. It never means that the placeholder must be filled. All generated text remains source-gated. `optional_supporting` placeholders stay empty unless the corresponding content is supplied, directly supported, functionally required, or explicitly requested.

### Template-specific role overrides

Use JSON so the override remains dependency-free and reviewable:

```json
{
  "version": 1,
  "overrides": {
    "12:3": {
      "layout_name": "Picture with Caption",
      "placeholder_name": "Text Placeholder 4",
      "content_role": "caption",
      "content_class": "optional_supporting",
      "note": "Use only for a supplied factual caption or source."
    }
  }
}
```

The key is `layout_index:placeholder_idx`. The optional layout and placeholder names are drift guards: the audit fails if a guarded name no longer matches or an override target no longer exists. The inspector also rejects unknown schema versions, fields, roles, and classes rather than silently applying them.

## Mapping content to layouts

Use the most specific layout that matches the requested content:

- Title slide: title/subtitle/date/client placeholders.
- Executive summary: title plus body/content placeholders.
- Chart slide: title plus chart placeholder, or title plus content placeholder geometry.
- Comparison: layouts with two content placeholders.
- Table-heavy slide: table placeholder or a wide content placeholder.

Content determines the layout; the layout must not determine or expand the content. Prefer the layout whose primary content regions match the available material. Do not choose a visually impressive layout if doing so requires a subtitle, caption, quote, section label, or takeaway that the source does not provide. Do not add freeform text boxes to compensate for an unsuitable layout when another approved layout fits.

An unused optional placeholder is not a visual defect or QA failure. Leave it empty, remove only the slide-level instance when safe, or select a more suitable approved layout. Do not alter a slide master or layout to hide an unused field. Preserve fixed logos, rules, slide numbers, and approved text already embedded on the master or layout, but do not create additional header or footer copy.

## Fallback when no chart placeholder exists

Use audited placeholder geometry:

```python
content = slide.placeholders[1]
graphic_frame = slide.shapes.add_chart(
    chart_type,
    content.left,
    content.top,
    content.width,
    content.height,
    chart_data,
)
chart = graphic_frame.chart
```

This still produces a native chart. The fallback only gives up placeholder insertion, not chart editability.

## Quality checks

- Confirm every slide was created from a template layout.
- Confirm titles are in title placeholders where possible.
- Confirm charts are graphic frames with `has_chart == True`.
- Confirm `verify_native_charts.py` reports embedded workbook relationships.
- Confirm no chart request was satisfied with a picture unless the user asked for a static image.
- Confirm every non-source text element is recorded as `transformed_source`, `functional_label`, or `generated_with_user_authorization`, with a source or authorization reference where appropriate.
- Confirm titles are literal, concise, and supported. Use a conclusion-led title only when the source supports the conclusion.
- Confirm captions provide factual identification, necessary context, evidence, or a source/credit. An image does not require a caption by default.
- Delete generated text that is redundant, vague, unsupported, decorative, or present only to fill space.
- Inspect freeform text boxes as well as placeholders. Empty optional placeholders pass QA.

Apply the same audit to user-approved fallback layouts. Additional fallback layouts or placeholders never authorize ornamental copy.
