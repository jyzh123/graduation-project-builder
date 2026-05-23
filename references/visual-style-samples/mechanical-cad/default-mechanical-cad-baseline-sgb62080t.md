# Default Mechanical CAD Baseline: SGB620/80T Passed Package

Use this file when a future mechanical CAD task needs a fallback drawing baseline and the current run does not contain a stronger user-provided DWG/PDF package sample.

## Activation Rule

- This is the default mechanical CAD baseline only when the current run has no stronger user-provided mechanical CAD package.
- A stronger current-run user sample wins over this stored fallback baseline.
- Use this baseline as both a visual baseline and a complexity baseline. Do not reduce it to only title-block imitation, page size, or rough page occupancy.

## Authoritative Stored Baseline Assets

Directory:
- `references/visual-style-samples/mechanical-cad`

Authoritative files:
- `default-mechanical-cad-baseline-sgb62080t.md`
- `sgb62080t-sheet-11-default-baseline.png`
- `sgb62080t-sheet-13-default-baseline.png`
- `sgb62080t-sheet-16-default-baseline.png`
- `sgb62080t-default-baseline-audit-v5.json`

Authoritative passed package outside the skill:
- `D:\项目\刮板输送机\SGB620-80T刮板输送机图纸包_PDF-DWG-DXF.zip`
- SHA256: `C4BB72B964F0849FF05DF5F66E22331A3CC9FCC6876B1E3727C12C5F5D2292A5`

Accepted project audit evidence:
- `D:\项目\刮板输送机\.codex\graduation-project-builder\20260523-sgb620-cad-overlap-repair\evidence\strict-audit-detail-upgrade-v5.json`
- strict audit verdict: `pass`
- total A0-equivalent workload: `12.5`

## Benchmark Sheet Roles

- `SGB62080T-11`
  - role: installation and drive integration benchmark
  - passed metrics: `drawing_object_count=1556`, `dxf_entity_total=2752`, `accepted_min_guard_gap=30.0`
- `SGB62080T-13`
  - role: assembly and dimension-chain benchmark
  - passed metrics: `drawing_object_count=5037`, `dxf_entity_total=9643`, `accepted_min_guard_gap=11.0`
- `SGB62080T-16`
  - role: installation, maintenance, lubrication, and trial-run benchmark
  - passed metrics: `drawing_object_count=13618`, `dxf_entity_total=26887`, `accepted_min_guard_gap=27.0`

## Default Quality Expectations

- The fallback package must read like graduation-level manufacturing drawings, not like sparse schematic redraws.
- Density must be distributed across multiple sheets rather than concentrated into one overloaded strip or one isolated dense page.
- Title block, BOM tables, technical requirements, notes, and dimension bands must keep explicit clearance from each other and from drawing geometry.
- Rendered review is mandatory. Entity totals alone cannot clear the baseline.
- If one sheet is under-dense, add real structure inside existing view envelopes first: hidden lines, centerlines, section refinement, bolt rows, guide-slot lines, reducer internals, chain-run internals, bearing/seat details.
- Do not clear a density shortfall by stuffing floating blocks into blank margins if that increases local crowding risk.

## Default Comparison Rule

- For future mechanical CAD tasks with no stronger user sample, compare against these stored PNG sheets and the copied v5 audit JSON together.
- If a future task is narrower than a full package, choose the closest matching benchmark sheet family and still preserve the same rendered-review cleanliness standard.
- If a future run produces a stronger accepted mechanical CAD package through a canonical skill-maintenance pass, promote that package explicitly before replacing this fallback baseline.
