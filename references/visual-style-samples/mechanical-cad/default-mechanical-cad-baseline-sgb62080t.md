# Default Mechanical CAD Baseline: SGB620/80T Six-Sheet Passed Package

Use this file when a future mechanical CAD task needs a fallback drawing baseline and the current run does not contain a stronger user-provided DWG/PDF package sample.

## Activation Rule

- This is the default mechanical CAD baseline only when the current run has no stronger user-provided mechanical CAD package.
- A stronger current-run user sample wins over this stored fallback baseline.
- Use this baseline as both a visual baseline and a complexity baseline. Do not reduce it to only title-block imitation, page size, or rough page occupancy.
- Do not allow the older three-sheet v5 stored sample or the earlier v10 six-sheet sample to be the default for SGB620/80T-like conveyor tasks after this v12 six-sheet baseline exists.

## Authoritative Stored Baseline Assets

Directory:
- `references/visual-style-samples/mechanical-cad`

Primary authoritative v12 six-sheet baseline directory:
- `references/visual-style-samples/mechanical-cad/sgb62080t-v12-six-sheet-baseline/`

Primary authoritative v12 baseline files:
- `default-mechanical-cad-baseline-sgb62080t.md`
- `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-00-a0-overall-assembly.png`
- `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-01-a1-head-drive-assembly.png`
- `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-02-a1-middle-trough-assembly.png`
- `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-03-a1-tail-tension-assembly.png`
- `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-04-a2-scraper-part.png`
- `sgb62080t-v12-six-sheet-baseline/sgb62080t-v12-05-a2-sprocket-shaft-component.png`
- `sgb62080t-v12-six-sheet-baseline/rendered-review-v12.json`
- `sgb62080t-v12-six-sheet-baseline/mechanical-drawing-package-v12-fixed.json`

Previous accepted v10 six-sheet files remain historical comparison evidence only:
- `references/visual-style-samples/mechanical-cad/sgb62080t-v10-six-sheet-baseline/`

Secondary historical v5 extension files:
- `sgb62080t-sheet-11-default-baseline.png`
- `sgb62080t-sheet-13-default-baseline.png`
- `sgb62080t-sheet-16-default-baseline.png`
- `sgb62080t-default-baseline-audit-v5.json`

Authoritative passed package recorded by v12 audit:
- path: `D:\项目\刮板输送机\最终交付\SGB620-80T刮板输送机图纸包_PDF-DWG-DXF_参考强化v12版.zip`
- SHA256: `6BEC6A564A9FC8FBA37D0098C292EAD3FA2F118587CFEB3CBDAE2EC2BEDEF84F`
- strict audit verdict: `pass`
- rendered overlap/overflow/ink/complexity verdicts: `pass`
- formal sheet workload: `3.0` A0-equivalent across six required sheets
- required dimensioned DXF count: `6` effective for this six-sheet formal package

## Benchmark Sheet Roles

- `SGB62080T-00`
  - role: A0 total assembly benchmark
  - sheet title: `SGB620-80T刮板输送机总装配图`
  - passed metrics: `pdf_drawing_objects=16379`, `dxf_DIMENSION=66`, `dxf_LINE=13598`, `dxf_HATCH=12148`, `dxf_TEXT=270`, `dxf_LEADER=29`
- `SGB62080T-01`
  - role: A1 head drive assembly benchmark
  - sheet title: `机头传动部装配图`
  - passed metrics: `pdf_drawing_objects=12232`, `dxf_DIMENSION=27`, `dxf_LINE=10271`, `dxf_HATCH=9507`, `dxf_TEXT=215`, `dxf_LEADER=29`
- `SGB62080T-02`
  - role: A1 middle trough assembly benchmark
  - sheet title: `中部槽装配图`
  - passed metrics: `pdf_drawing_objects=12486`, `dxf_DIMENSION=36`, `dxf_LINE=8026`, `dxf_HATCH=7207`, `dxf_TEXT=253`, `dxf_LEADER=43`
- `SGB62080T-03`
  - role: A1 tail tensioning assembly benchmark
  - sheet title: `机尾装紧部装配图`
  - passed metrics: `pdf_drawing_objects=9615`, `dxf_DIMENSION=29`, `dxf_LINE=7835`, `dxf_HATCH=7029`, `dxf_TEXT=210`, `dxf_LEADER=21`
- `SGB62080T-04`
  - role: A2 scraper part benchmark
  - sheet title: `刮板零件图`
  - passed metrics: `pdf_drawing_objects=5559`, `dxf_DIMENSION=31`, `dxf_LINE=4180`, `dxf_HATCH=3839`, `dxf_TEXT=184`, `dxf_LEADER=28`
- `SGB62080T-05`
  - role: A2 sprocket shaft component benchmark
  - sheet title: `链轮轴（连轮轴）组件零件图`
  - passed metrics: `pdf_drawing_objects=6255`, `dxf_DIMENSION=36`, `dxf_LINE=4218`, `dxf_HATCH=3662`, `dxf_TEXT=175`, `dxf_LEADER=14`

The historical v5 sheets remain usable only as additional density and visual-reference evidence. They must not replace the primary v10 six-sheet sheet-set requirement.

## Default Quality Expectations

- The fallback package must read like graduation-level manufacturing drawings, not like sparse schematic redraws.
- Use the current passed SGB620/80T v12 six-sheet package as the default complexity baseline for future mechanical CAD tasks: dense sheet-by-sheet structure, distributed detail, clear line family control, protected-zone clearance, and visible manufacturing intent are the default expectation, not an optional upgrade.
- Do not reduce future CAD work to outline-only linework, simplified skeleton views, or a single sparse sheet that only proves geometry exists.
- Density must be distributed across multiple sheets rather than concentrated into one overloaded strip or one isolated dense page.
- Title block, BOM tables, technical requirements, notes, and dimension bands must keep explicit clearance from each other and from drawing geometry.
- The drawing family should keep the same academic/manufacturing tone as the passed package: full borders, coordinate/title infrastructure, sectional/detail enrichment, hidden/center lines where needed, and no overlap between annotation clusters and frame tables.
- If the current task uses a user-provided CAD/PDF sample, the stronger sample wins; if it is not available, the stored SGB620/80T v12 package remains the default fallback benchmark.
- Rendered review is mandatory. Entity totals alone cannot clear the baseline.
- If one sheet is under-dense, add real structure inside existing view envelopes first: hidden lines, centerlines, section refinement, bolt rows, guide-slot lines, reducer internals, chain-run internals, bearing/seat details.
- When a sheet already contains notes, dimensions, title tables, or dense text blocks, preserve their clearance and move detail into the view geometry rather than crowding the annotation area.
- Do not clear a density shortfall by stuffing floating blocks into blank margins if that increases local crowding risk.
- For future SGB620/80T mechanical CAD output, the default expectation is a full graduation-project drawing package with clear sheet-to-sheet workload distribution, not a simplified export that only satisfies A0 area or file count.
- For SGB620/80T-like conveyor packages, the default formal sheet set is six drawings: A0 total assembly, A1 head drive assembly, A1 middle trough assembly, A1 tail tensioning assembly, A2 scraper part, and A2 sprocket shaft component. The wording `机尾装紧部` from user input maps to the correct engineering term `机尾张紧部`; the formal sheet title should use `张紧`.
- A six-sheet formal package must not be failed solely because a generic eight-sheet dimension-distribution default exists. The audit must clamp `min_dimensioned_dxf_files` to the available required sheet count when the locked manifest requires exactly six formal DXF sheets, and must record `effective_min_dimensioned_dxf_files=6` or the equivalent evidence.
- Do not use the older three-sheet v5 fallback or the earlier v10 package as the default for new SGB620/80T tasks after this v12 baseline exists. Those files are historical benchmarks, not the primary sheet-set baseline.

## Default Comparison Rule

- For future mechanical CAD tasks with no stronger user sample, compare against the v12 six-sheet stored PNG sheets and the copied v12 audit JSON files first.
- If a future task is narrower than a full package, choose the closest matching benchmark sheet family and still preserve the same rendered-review cleanliness standard.
- If a future run produces a stronger accepted mechanical CAD package through a canonical skill-maintenance pass, promote that package explicitly before replacing this fallback baseline.
