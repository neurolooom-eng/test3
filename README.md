# AIDA leadership deck

A seven-slide, 16:9 HTML deck for Quality Engineers presenting AIDA to senior leadership. It exports to an editable PowerPoint file and also prints to PDF.

All content is **synthetic placeholder text and numbers**. Replace it with validated internal information before presenting, and keep confidential figures out of shared repositories.

## Quick start

| Task | How |
|---|---|
| Present | Open `index.html` in Chrome or Edge. It works offline from `file://`, with no build step. |
| Navigate | Use `→` / `Space` / `PageDown` to go forward and `←` / `PageUp` to go back. `Home` / `End` jump to the first or last slide. `?slide=3` in the URL opens a given slide. |
| Speaker notes | Press `N` or the **Notes** button. Notes show the presenter, the time guide, the demo steps and the handover line. |
| Light / dark | Use the **Theme** button, or add `?palette=dark` to the URL. |
| PowerPoint | Use the **Export PPTX** button. It exports the current palette and the current savings inputs. |
| PowerPoint (CLI) | Run `npm install` and then `npm run export:pptx`. The file is written to `dist/AIDA-leadership-deck.pptx`. Options: `--release minor --user experienced --palette dark --out path.pptx`. |
| PDF fallback | Use **Print / PDF** and choose *Save as PDF*. The page size is set to 13.333 × 7.5 in with one slide per page. Turn on *Background graphics* if your browser doesn't keep it on. |
| Tests | Run `npm test`. This tests the savings maths, divide-by-zero handling, the seven-slide structure, shape bounds and the wording guardrails. |

## Slides

| # | Owner | Slide | Time |
|---:|---|---|---:|
| 1 | Presenter 1 | AIDA: Evidence-led Quality Reviews | 0:30–0:45 |
| 2 | Presenter 1 | Use Case 1 — Evidence-based Release Traceability Review | 1:00 |
| 3 | Presenter 1 | Use Case 1 — From Formal Review to First Time Right (demo) | 2:30–3:00 |
| 4 | Presenter 2 | Use Case 2 — Source-driven Lifecycle Quality Assurance | 1:00 |
| 5 | Presenter 2 | Use Case 2 — Review Quality at the Source (demo) | 2:30–3:00 |
| 6 | Both | What AIDA Changes | 1:00–1:15 |
| 7 | Presenter 1 | Releasing QE Capacity for Deeper Technical Quality Review | 1:00–1:15 |

The handover line after slide 3 is in the speaker notes for slide 3 (`speakerNotes.handoverAfterSlide3`).

## Design

- **Headlines and visuals:** each slide has one short headline (8 words or fewer) and one visual. The visuals are an evidence chain with a red return loop, today/next lanes, a lifecycle track, a hub diagram, from→to shifts, and a large capacity number.
- **Word budget:** about 30 visible words per slide. `npm test` fails if any slide goes over 35 words, not counting footnotes or numbers.
- **Detail lives in the speaker notes:** problem statements, failure points, full checkpoint questions, demo steps, the handover line, the assumptions and the full disclaimer. Press `N` to show them.
- **Guardrail wording:** the caveat and the disclaimer stay on the slides as one-line footnotes. Current/future status is shown as small tags.
- **Assumptions on slide 7:** they are hidden behind the **Assumptions** button. Edit them there and the numbers recalculate. In the PPTX they are listed in the slide 7 speaker notes.
- **Colour:** navy and royal blue lead, and red is used only for gaps and risk. Change the colours in `theme.palettes`.

## Editing content: `config/deck.config.js`

Normal updates only touch this one file:

- **`brand`**: tool name, title, subtitle, footer and logo. Every string can use `{tool}`, so renaming AIDA happens in one place.
- **`systemLabels` / `systemLabelRoles`**: ARIS, Azure, Miro / Azure, repository, defects.
- **`presenters`**, **`agenda`**: presenter names, slide owners and time guides.
- **`useCaseOne`**: problem statement, review-cycle range, artefact labels, gap markers (`issueLabels`), target cycles and demo steps.
- **`useCaseTwo`**: lifecycle labels, checkpoints, `codeReviewStatus` (hidden on the slide by default) and demo steps.
- **`value`**: the three pillars, the caveat and the bottom-line statement.
- **`savings`**: release scenarios, user scenarios, adoption rate, labels, reinvestment items and `disclaimerLines`. Set `showIllustrativeTag: false` once real, validated values are in place.
- **`theme`**: light and dark palettes, HTML font stacks, PowerPoint font names, logo position and diagram line widths.

### Savings model (slide 7)

The inputs on slide 7 can be edited live. Changing the release type resets your edits to that scenario's defaults. The calculation is in `src/calc.js`:

```
eligibleDocuments             = eligibleDocumentsPerRelease × adoptionRate
currentQeHoursPerRelease      = eligibleDocuments × manualHoursPerDocument
aidaAssistedQeHoursPerRelease = eligibleDocuments × aidaHoursPerDocument
qeCapacityReleasedHours       = current − aidaAssisted
qeCapacityReleasedPercent     = released / current × 100    (shows "—" when current is 0 or blank)
annualQeCapacityReleasedHours = released × releasesPerYear  (shows "—" when blank)
```

`aidaHoursPerDocument` is set for each user scenario (new / typical / experienced) within each release scenario.

## Applying a corporate theme

1. **Colours and fonts**: edit `theme.palettes` and `theme.fonts`. The HTML deck (through CSS variables) and the PPTX use the same values.
2. **Logo**: set `brand.logoPlaceholder.imagePath` to a PNG or SVG next to `index.html`. For the browser export, serve the folder over HTTP (for example `npx serve`), because browsers may block reading local images from `file://`. The CLI export reads local paths directly.
3. **Inside PowerPoint**: every slide uses one of two slide masters (`AIDA_TITLE`, `AIDA_CONTENT`), and each title is a real title placeholder. You can apply an official theme with *Design → Themes* or by pasting slides into the corporate template with *Use destination theme*. All diagrams are native rectangles, ellipses, lines and text boxes, so they can be ungrouped, recoloured and edited.

## How it is built

```
config/deck.config.js   all content and theme (edit this)
src/calc.js             savings maths (pure functions, unit-tested)
src/slides.js           slide model: each slide is a list of shapes positioned in inches
src/render-html.js      draws the model as HTML; deck controls
src/pptx-builder.js     draws the same model as native PowerPoint shapes (PptxGenJS)
scripts/export-pptx.js  Node CLI export
vendor/pptxgenjs/       PptxGenJS 4.0.1 browser bundle (MIT), vendored so export works offline
```

Both the HTML deck and the PowerPoint file are drawn from the same shape list, so they stay in step.

## Content guardrails built into the defaults

- Current use and future adoption are labelled differently, by both tag text and border style (solid or dashed), so the difference never depends on colour alone.
- Use Case 2 does not claim code-review validation. `codeReviewStatus.showOnSlide` defaults to `false`, and the boundary is in the speaker notes.
- The deck says "full-population review" and "100% coverage of applicable configured checks", together with the explicit caveat that this is not 100% assurance. It never says "100% sampling".
- Review-cycle reduction is shown as a target or assumption (`[target]` placeholder). Savings are shown as *QE capacity released*, with the configurable disclaimer.

## Known limitations

- **Fonts:** the default is Segoe UI, which is standard on Windows. On machines without it, the browser falls back through the font stack. PowerPoint substitutes a font on its own, so line breaks can shift a little. Layouts leave room for this, but check the deck on the machine you will present from.
- **Minimum text size:** slides are built to be visual, so most of the explanation lives in the speaker notes. Headlines are 28 pt and the key statements on slides are 14–20 pt. Some labels are smaller: section pills and eyebrow labels (10.5–12 pt), the slide 7 assumption inputs (about 9.5–12 pt), and the disclaimer (8.5 pt).
- **Exported savings slide:** in the PPTX, slide 7 shows the numbers at export time, and the assumptions are listed in its speaker notes. To change them, edit the config (or the live panel) and export again.
