# Figure Plan Template

## Figure Inventory
- figure id:
- source comment id/path:
- source anchor text or existing caption:
- chapter:
- figure type:
- inferred figure family:
- semantic purpose:
- required source material:
- screenshot or drawn figure:
- source kind: runtime screenshot | code screenshot | data chart | structural figure
- task card path:
- status: pending | pass | fail | skipped-with-reason
- skip reason:

## Style Source
- user sample:
- template sample:
- template/accepted-sample figure checked:
- template figure sample baseline:
- template figure sample evidence path:
- no-template-figure-sample verdict:
- stored fallback sample:
- skill-internal fallback sample path:
- final chosen source of truth:
- sample-lock evidence path:

## Drawing Constraints
- background:
- text color:
- border/connector style:
- forbidden elements:
- caption rule:
- in-figure language rule:
- English-label exception rule:
- no-table-as-image rule:

## Asset Manifest
- figure asset manifest path:
- manifest entry id:
- draw.io source path:
- SVG export path:
- raster fallback path:
- runtime real route/page URL:
- runtime capture method:
- runtime readiness cue:
- runtime expected window/viewport size:
- runtime capture/window bbox:
- runtime actual image size:
- runtime full-window coverage ratio:
- runtime full-window capture verdict:
- final DOCX relationship evidence path:
- SVG-primary evidence path:
- in-figure language verdict:
- English-label exception reason:
- table-as-image check:
- real Word table requirement if tabular:

## Acceptance Plan
- pre-insertion checks:
- post-insertion checks:
- post-caption body style verdict:
- post-caption label-prefix leakage verdict:
- caption sibling contamination audit path:
- body donor paragraph evidence:
- rendered-page target:
- review evidence record paths:
- dense-zone crop evidence paths:
- final per-figure verdict evidence path:

## Per-Figure Closure Table
| figure id | family | source comment/anchor | task card path | manifest entry | pre evidence | post rendered evidence | final status | skip reason |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |

## Blocking Rules
- Do not replace the per-figure table with a single `figure set passed` summary.
- Every figure id in the active scope must have a task card, manifest entry, evidence path, final status, and skip reason when skipped.
- Runtime screenshot rows must prove full-window capture geometry; a partial top-left capture cannot be treated as a usable runtime screenshot.
- Structural figure rows for Chinese theses must record a passing in-figure language verdict; English labels require a literal-identifier exception reason.
- Rows-and-columns data must be represented as real Word tables, not as raster/SVG images with a figure caption.
- If any row has blank evidence or `pending`, thesis DOCX assembly and final acceptance are blocked.
