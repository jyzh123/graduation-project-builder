# Thesis Figure Task Template

## Figure Request Intake
- source comment id/path:
- source anchor text or caption:
- figure type:
- inferred figure family:
- chapter/location:
- semantic purpose:
- linked figure plan id:
- linked user-reported issue ledger id:
- source of truth for style:
- replacement or new figure:
- final DOCX path:
- final DOCX sha256:

## Style Lock
- background:
- text color:
- border/connector style:
- template/accepted-sample figure checked:
- template figure sample baseline:
- template figure sample evidence path:
- no-template-figure-sample verdict:
- skill-internal fallback sample path:
- final sample asset/template reference:
- in-figure language verdict:
- English-label exception reason:
- forbidden elements:
- figure family sample-lock verdict:
- sample-lock evidence path:

## Drawing Plan
1.
2.
3.

## Asset Contract
- source kind: runtime screenshot | code screenshot | data chart | structural figure | algorithm result
- draw.io source path:
- SVG export path:
- raster fallback path:
- asset manifest path:
- asset manifest entry id:
- structural geometry validation report path:
- source-scale shape bbox map path:
- inserted-scale geometry evidence path:
- inserted-scale collision evidence path:
- relation-attribute collision verdict:
- all-shape overlap verdict:
- source-to-inserted geometry verdict:
- final DOCX relationship evidence path:
- SVG-primary evidence path:
- table-as-image check:
- real Word table requirement if tabular:
- route-caption-asset map path:
- runtime real route/page URL:
- runtime capture method:
- runtime readiness cue:
- runtime expected window/viewport size:
- runtime capture/window bbox:
- runtime actual image size:
- runtime full-window coverage ratio:
- runtime full-window capture verdict:
- algorithm accepted result image path:
- algorithm source/input image path:
- algorithm generation or inference script path:
- algorithm model/output log path:
- algorithm existing result source:
- algorithm user-provided asset evidence:
- algorithm caption-to-asset mapping:
- algorithm authenticity verdict:
- dense-zone crop evidence paths:

## Review Log
### Pre-Insertion
- style compliance:
- geometry compliance:
- source-scale collision report:
- caption readiness:
- review evidence record paths:
- standalone export review evidence path:
- Chinese in-figure label check:

### Post-Insertion
- rendered-page visibility:
- caption pairing:
- surrounding layout:
- post-caption body style verdict:
- post-caption label-prefix leakage verdict:
- caption sibling contamination audit path:
- body donor paragraph evidence:
- review evidence record paths:
- touched-page rendered evidence path:

## Final Status
- pass / fail:
- skipped reason:
- next fix if failed:

## Blocking Rules
- Do not mark `pass` when source comment id/path is blank for comment-driven work.
- Do not mark `pass` when linked figure plan id, asset manifest path, pre-insertion evidence, post-insertion rendered evidence, or final DOCX relationship evidence is blank.
- Do not mark a runtime screenshot `pass` unless the manifest records real route/page URL, capture method, readiness cue, accepted screenshot path, caption-to-asset mapping, expected window or viewport size, actual image size, coverage ratio, and a passing full-window capture verdict.
- Do not mark an algorithm result `pass` unless the manifest records an accepted result image path, caption-to-asset mapping, source/provenance evidence, and a passing authenticity verdict.
- Do not mark YOLOv8, DBNet, CRNN, OCR, detection-box, preprocessing, or recognition-result figures `pass` when the asset is a mockup, placeholder, hand-drawn sample, or `示意图` / `样例图` without real provenance.
- Do not mark a structural figure `pass` unless draw.io source path, SVG export path, raster fallback path, SVG-primary evidence, and dense-zone crop evidence when applicable are present.
- Do not mark an ER structural figure `pass` unless structural geometry validation report path, source-scale shape bbox map path, inserted-scale geometry evidence path, inserted-scale collision evidence path, dense-zone crop evidence paths, relation-attribute collision verdict, all-shape overlap verdict, and source-to-inserted geometry verdict are present and passing.
- Do not mark a Chinese-thesis structural figure `pass` unless in-figure language verdict is `pass` and any remaining English label has a recorded literal-identifier exception reason.
- Do not use an image to masquerade as a table. If the figure content is mainly rows and columns, convert it to a real Word table object and use a table title instead of a figure caption.
- Use `skipped reason` only for a real not-applicable decision with evidence; never use it to bypass missing figure work.
