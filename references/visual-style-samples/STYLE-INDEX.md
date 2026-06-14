# Visual Style Sample Index

Use this index to understand which real sample images already exist and what they should be used for.

## 1. Thesis Figures

Directory:
- `references/visual-style-samples/figures`

Current sample files:
- `figure-er-diagram-sample-01.png`
- `figure-er-diagram-sample-01.svg`
- `figure-flowchart-vertical-sample-01.png`
- `figure-flowchart-vertical-sample-01.svg`
- `figure-layered-architecture-sample-20260419.jpg`
- `figure-system-structure-tree-sample-01.png`
- `figure-system-structure-tree-sample-01.svg`
- `figure-use-case-diagram-sample-01.png`
- `figure-use-case-diagram-sample-01.svg`
- `figure-sequence-diagram-sample-01.svg`

Default semantic usage:
- overall system architecture diagrams
- module or structure diagrams
- layered architecture diagrams with outer frame, inner grouped modules, and side pillars
- ER or entity-attribute style figures
- use-case or flow-style thesis figures
- vertical monochrome textbook flowcharts with rounded start/end terminators, centered decision diamonds, and simple `真 / 假` branch labels
- monochrome UML sequence diagrams with `alt` branches

Usage rule:
- When the user provides no stronger sample, compare newly generated thesis figures against these images first.
- If a figure type is more specific than these generic samples, prefer the user-provided sample or the current template.
- The stored flowchart SVG is the locked source-of-truth sample, and the matching PNG is the quick visual comparison artifact for top-down thesis flowcharts when no stronger newer sample exists.
- The layered architecture JPG is a user-provided screenshot reference and should be treated as the mandatory default comparison target for architecture-family thesis figures unless a newer stronger architecture sample overrides it.

## 2. Formula Style

Directory:
- `references/visual-style-samples/formulas`

Current sample files:
- `formula-layout-sample-01.png`
- `formula-layout-sample-01.svg`

Default semantic usage:
- formula line appearance
- formula numbering placement
- spacing around formulas

Usage rule:
- Use this as a visual fallback only when the current template does not clearly define formula appearance.
- If the accepted manuscript or template already uses chapter-based formula numbering such as `(2-1)`, preserve that numbering family instead of switching to generic global numbering.

## 3. Table Style

Directory:
- `references/visual-style-samples/tables`

Current sample files:
- `table-style-sample-03-20260419-page.png`

Default semantic usage:
- thesis table border style
- table caption position
- table density and alignment
- rendered acceptance evidence for the WPS built-in `second three-line table` preset

Usage rule:
- Prefer the template first.
- For thesis three-line tables under the active skill memory, the only preserved table-family authority is the WPS built-in `second three-line table` preset.
- `table-style-sample-03-20260419-page.png` is kept only as rendered acceptance evidence for that preset.
- Do not use older table screenshots or older stored SVG table samples as alternative three-line-table authorities.

## 4. TOC Style

Directory:
- `references/visual-style-samples/toc`

Current sample files:
- `toc-style-sample-01.png`
- `toc-style-sample-01.svg`

Default semantic usage:
- TOC indentation hierarchy
- dotted leaders
- right-aligned page numbers
- TOC spacing rhythm

Usage rule:
- If a TOC sample exists, treat it as a strong visual target during final thesis formatting review.

## 5. Mechanical CAD Style

Directory:
- `references/visual-style-samples/mechanical-cad`

Current sample files:
- `default-mechanical-cad-baseline-sgb62080t.md`
- `sgb62080t-v10-six-sheet-baseline/sgb62080t-v10-00-a0-overall-assembly.png`
- `sgb62080t-v10-six-sheet-baseline/sgb62080t-v10-01-a1-head-drive-assembly.png`
- `sgb62080t-v10-six-sheet-baseline/sgb62080t-v10-02-a1-middle-trough-assembly.png`
- `sgb62080t-v10-six-sheet-baseline/sgb62080t-v10-03-a1-tail-tension-assembly.png`
- `sgb62080t-v10-six-sheet-baseline/sgb62080t-v10-04-a2-scraper-part.png`
- `sgb62080t-v10-six-sheet-baseline/sgb62080t-v10-05-a2-sprocket-shaft-component.png`
- `sgb62080t-v10-six-sheet-baseline/rendered-review-v10.json`
- `sgb62080t-v10-six-sheet-baseline/rendered-review-manifest-v10.json`
- `sgb62080t-v10-six-sheet-baseline/cad-linework-fidelity-v10.json`
- `sgb62080t-v10-six-sheet-baseline/mechanical-drawing-package-v10-fixed.json`
- `sgb62080t-sheet-11-default-baseline.png`
- `sgb62080t-sheet-13-default-baseline.png`
- `sgb62080t-sheet-16-default-baseline.png`
- `sgb62080t-default-baseline-audit-v5.json`

Default semantic usage:
- fallback benchmark for graduation-level mechanical CAD drawing packages
- primary visual comparison target for the v10 six-sheet SGB620/80T package family
- visual comparison target for A0 total assembly, A1 head drive, A1 middle trough, A1 tail tensioning, A2 scraper part, and A2 sprocket shaft component sheets
- copied v10 audit baseline for density, rendered-review, lineweight/linetype fidelity, frame overflow, overlap, and package-workload expectations
- v5 three-sheet assets are secondary historical extension benchmarks only

Usage rule:
- When the current run has no stronger user-provided mechanical CAD package sample, this directory is the mandatory default mechanical CAD baseline.
- Use the stored v10 six-sheet PNG sheets together with the copied v10 audit JSON files and the baseline note; do not compare only border/title-block style.
- A stronger current-run user sample overrides this fallback baseline.
- This baseline controls both complexity and cleanliness expectations, including no-overlap, annotation margin clearance, local crowding, and manufacturing-view depth.
- For SGB620/80T-like conveyor tasks, the default expected sheet set is six formal sheets: A0 total assembly, A1 head drive assembly, A1 middle trough assembly, A1 tail tensioning assembly, A2 scraper part, and A2 sprocket shaft component. A locked six-sheet package may use `effective_min_dimensioned_dxf_files=6`; it must not be rejected merely because an older generic default asked for eight dimensioned DXF files.

## 6. Missing Sample Categories

The following categories are conceptually supported by the visual-style-samples system but do not currently contain confirmed real sample image files in this skill snapshot:

- `title-pages`
- `headers-footers`
- `body-text`

If future tasks depend heavily on those surfaces, add real sample images before treating them as style-controlled defaults.

## 7. Maintenance Rule

When new sample images are added:

- place them in the correct category directory
- update this index
- describe the intended semantic usage, not only the filename
- do not leave unidentified screenshots without a usage note
