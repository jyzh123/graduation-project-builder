# Review Evidence Template

Use this template for evidence records referenced by `assets/final-acceptance-template.md`.

Each evidence path in the acceptance record should point to a file that follows this schema rather than directly to a screenshot, PDF, or raw log.

For thesis-format evidence, canonical protected surface ids, required fields, effective font-chain proof, and agent audit handoff fields are defined in `references/thesis/format-rules/protected-surface-evidence-contract.md`.

For protected abstract and TOC surfaces, `- target identifier:` and `- baseline surface id:` must both name the same exact protected surface id being proved, such as `zh_abstract_title`, `en_keyword_line`, `toc_title`, or `toc_page_number_column`. Generic ids such as `abstract`, `front matter`, `toc`, or `page` are invalid for these fields, and one evidence file must not be reused to prove multiple protected surfaces.

For protected abstract and TOC surfaces, the font fields are mandatory proof, not optional notes. `baseline effective font chain` and `actual effective font chain` must resolve direct run properties, character style, paragraph style, basedOn chain, docDefaults, theme major/minor mappings, and WPS/Word UI displayed font names before the evidence can pass.

For every thesis-format rendered evidence record, surface geometry fields are mandatory proof, not optional notes. The evidence must compare the rendered target region against the rendered template region with numeric baseline/actual measurements. Natural-language claims such as `template equals actual`, `matched`, or `looks correct` are not evidence.

For every template-owned thesis-format surface, surface paragraph-dialog and typography fields are mandatory proof, not optional notes. The evidence must compare template-vs-actual style binding, WPS/Word paragraph-dialog metrics, typography, spacing, line-spacing mode/value, indentation, tabs/leaders, keep/list/page-break behavior, and scale/compression verdict. A broad `baseline metrics` or `actual metrics` sentence is not enough.

For protected TOC surfaces, the TOC-specific visual geometry fields are additional mandatory proof. TOC title, entries, dotted leaders, and page-number column evidence must compare the rendered target TOC region against the rendered template TOC region at bounding-box, line-spacing, indentation, leader-density, page-number-column, and occupancy-rhythm level before the evidence can pass. These fields must contain numeric baseline/actual measurements; natural-language claims such as `template equals actual` or `looks matched` are not evidence.

For protected TOC surfaces, TOC level style binding remains mandatory and must be expanded into paragraph-dialog and typography fields. Evidence must show template-vs-actual values for TOC title and every used TOC level, including style id/name, font size, paragraph before/after spacing, line-spacing mode/value, indentation, tab stops, and leader type. A TOC that is style-bound but smaller, denser, or proportionally compressed still fails.

For protected TOC surfaces, the paragraph-and-typography proof must also include visible-run typography. The record must compare the locked template donor against the target for title text runs, each used TOC level's entry text runs, tab/leader runs, and page-number runs, including direct `w:rPr`, `w:rFonts` slots, theme font slots, size/sizeCs, and weight. A TOC whose paragraph style and line spacing match but whose visible runs lost the template direct font properties is a failed TOC.

For protected body heading surfaces, the record must prove every used heading level independently. Style names are not enough. Evidence must compare template-vs-actual direct run typography, effective font chain, `w:sz`/`w:szCs`, bold state including explicit off values, paragraph-dialog spacing/indent/list/outline values, body-format residue cleanup, and TOC/chapter-start synchronization.

For thesis structural figure evidence, source-scale and inserted-scale collision proof is mandatory when the figure is ER/dense structural. ER evidence must name the geometry validation report generated from draw.io source, the source-scale shape bounding-box map, inserted-scale dense-zone evidence, relation-attribute collision verdict, all-shape overlap verdict, and final source-to-inserted geometry verdict. A full-page screenshot or prose statement that the figure "looks clear" is not enough.

## Review Evidence Record

## Evidence Meta
- evidence type:
- task mode:
- protected-surface evidence contract path:
- reviewed output path:
- reviewed output sha256:
- evidence created after mutation?: yes
- artifact path:
- source review-artifact inventory path:
- source review-artifact inventory sha256:
- final review-artifact diff path:
- final review-artifact diff sha256:
- review comments/change marks preservation verdict:
- comments strip explicit user approval:
- source body-citation run inventory path:
- source body-citation run inventory sha256:
- final body-citation run diff path:
- final body-citation run diff sha256:
- body citation superscripts preservation verdict:
- citation audit final DOCX sha256:
- citation audit source-to-final run diff path:
- renderer executable path:
- render command:
- rendered PDF path:
- page image path list:
- measurement JSON path:
- measurement JSON sha256:
- measurement schema:
- measurement generator script:
- measurement generator command:
- measurement generated at UTC:
- evidence aggregator script:
- evidence aggregator is measurement producer?: no
- measurement provenance verdict:

## Target
- target surface:
- target identifier:
- canonical protected surface id:
- target pages or region:
- logical-to-physical page mapping method:
- sentinel text confirmed:

## Blast Radius
- blast-radius pages:
- neighboring surfaces checked:

## Baseline Comparison
- baseline source path and sha256:
- baseline surface id:
- protected surface owner lane:
- protected surface audit lane:
- baseline paragraph/run path:
- baseline metrics:
- actual paragraph/run path:
- actual metrics:
- surface style binding baseline/actual: template styleId= styleName= type= basedOn= directParagraphFormatting=; actual styleId= styleName= type= basedOn= directParagraphFormatting=
- surface WPS/Word paragraph-dialog metrics baseline/actual: template alignment= outline= list= before=pt after=pt lineMode= lineValue=pt left=pt right=pt firstLine=pt hanging=pt; actual alignment= outline= list= before=pt after=pt lineMode= lineValue=pt left=pt right=pt firstLine=pt hanging=pt
- surface typography baseline/actual: template font= size=pt weight= italic= underline= color=; actual font= size=pt weight= italic= underline= color=
- surface paragraph spacing baseline/actual: template before=pt after=pt; actual before=pt after=pt
- surface line-spacing mode/value baseline/actual: template mode= value=pt; actual mode= value=pt
- surface indentation chars/points baseline/actual: template left=pt right=pt firstLine=pt hanging=pt chars=; actual left=pt right=pt firstLine=pt hanging=pt chars=
- surface tab stop leader baseline/actual: template tab=pt leader=; actual tab=pt leader=
- surface keep/list/page-break baseline/actual: template keepNext= keepLines= widowControl= pageBreakBefore= outline= list=; actual keepNext= keepLines= widowControl= pageBreakBefore= outline= list=
- surface scale/compression verdict:
- surface paragraph-and-typography verdict:
- keyword run split extraction method:
- keyword label run text baseline/actual:
- keyword label run isolated baseline/actual:
- keyword content run count baseline/actual:
- keyword label bold/strong baseline/actual:
- keyword content bold baseline/actual:
- keyword separator baseline/actual:
- keyword run split verdict:
- Chinese abstract mixed-script extraction method:
- Chinese abstract Latin/digit run count baseline/actual:
- Chinese abstract Latin/digit font slots baseline/actual:
- Chinese abstract Latin/digit builder/default font rejection verdict:
- Chinese abstract mixed-script font-chain verdict:
- baseline effective font chain:
- actual effective font chain:
- effective font slots compared:
- theme/default font alias verdict:
- WPS/Word UI font display evidence:
- effective font-chain verdict:
- rendered region image path:
- template rendered region image path:
- actual rendered region image path:
- surface geometry comparison method:
- surface crop schema:
- surface crop generator:
- surface crop source page images baseline/actual:
- surface crop source image sha256 baseline/actual:
- surface crop source image size baseline/actual:
- surface crop fraction map baseline/actual:
- surface crop threshold baseline/actual:
- surface page index baseline/actual:
- surface crop bbox baseline/actual: template crop x= y= w= h=; actual crop x= y= w= h=
- surface content bbox baseline/actual: template content bbox x= y= w= h=; actual content bbox x= y= w= h=
- surface nonwhite ratio baseline/actual: template nonwhite_ratio=; actual nonwhite_ratio=
- surface blank crop verdict:
- surface binding method:
- surface bbox baseline/actual: template bbox x= y= w= h=; actual bbox x= y= w= h=
- surface position baseline/actual: template x= y=; actual x= y=
- surface size baseline/actual: template w= h=; actual w= h=
- surface line-height y-delta baseline/actual: template line=pt y-delta=; actual line=pt y-delta=
- surface spacing before/after baseline/actual: template before=pt after=pt; actual before=pt after=pt
- surface indentation/tab baseline/actual: template left=pt firstLine=pt tab=pt; actual left=pt firstLine=pt tab=pt
- surface page occupancy baseline/actual: template occupancy=%; actual occupancy=%
- surface geometry verdict:
- package baseline manifest path:
- package drift report path:
- package drift verdict:
- pre-mutation page map path:
- post-mutation page map path:
- whole-document pagination diff path:
- DOCX pagination structure schema:
- DOCX pagination structure generator:
- DOCX pagination structure evidence path:
- DOCX pagination structure verdict:
- section count baseline/actual:
- header/footer reference map baseline/actual:
- header/footer link-to-previous inferred map baseline/actual:
- section boundary map baseline/actual:
- section property map baseline/actual:
- page-number format/restart map baseline/actual:
- header/footer link-to-previous map baseline/actual:
- hard page-break / section-break map baseline/actual:
- fatal pagination topology differences:
- allowed content-growth pagination differences:
- all pagination topology differences:
- section count verdict:
- header/footer reference verdict:
- page-number restart verdict:
- field-refresh before/after state:
- TOC-to-heading page sync map:
- logical-to-physical page map:
- rendered page count baseline/actual:
- blank/near-empty page scan verdict:
- chapter opener page map:
- tail-block opener page map:
- tail-block references prior-block separation verdict:
- page-class occupancy rhythm verdict:
- whole-document pagination verdict:
- TOC template rendered page/region image path:
- TOC actual rendered page/region image path:
- TOC visual comparison method:
- TOC title bbox baseline/actual: template bbox x= y= w= h=; actual bbox x= y= w= h=
- TOC first-entry bbox baseline/actual: template first-entry bbox x= y= w= h=; actual first-entry bbox x= y= w= h=
- TOC row bbox map: template rows r1 x= y= w= h=, r2 x= y= w= h=; actual rows r1 x= y= w= h=, r2 x= y= w= h=
- TOC per-level left-indent x baseline/actual: template level1 x= level2 x= level3 x=; actual level1 x= level2 x= level3 x=
- TOC line-spacing y-delta baseline/actual: template y-delta= line=pt; actual y-delta= line=pt
- TOC dotted-leader start/end/density baseline/actual: template leader start_x= end_x= density=dots/cm; actual leader start_x= end_x= density=dots/cm
- TOC page-number x column baseline/actual: template page-number x=; actual page-number x=
- TOC row count per page baseline/actual: template page1 rows=; actual page1 rows=
- TOC title-to-first-entry gap baseline/actual: template gap=pt; actual gap=pt
- TOC page occupancy rhythm baseline/actual: template page1 rows= occupancy=%; actual page1 rows= occupancy=%
- TOC visual geometry verdict:
- TOC style binding baseline/actual: template title style= level1 style= level2 style= level3 style=; actual title style= level1 style= level2 style= level3 style=
- TOC WPS paragraph-dialog metrics baseline/actual: template style= outline= before=pt after=pt lineMode= lineValue=pt left=pt right=pt firstLine=pt hanging=pt tab=pt leader=; actual style= outline= before=pt after=pt lineMode= lineValue=pt left=pt right=pt firstLine=pt hanging=pt tab=pt leader=
- TOC title typography baseline/actual: template font= size=pt weight=; actual font= size=pt weight=
- TOC per-level typography baseline/actual: template L1 font= size=pt weight=; L2 font= size=pt weight=; L3 font= size=pt weight=; actual L1 font= size=pt weight=; L2 font= size=pt weight=; L3 font= size=pt weight=
- TOC per-level paragraph spacing baseline/actual: template L1 before=pt after=pt; L2 before=pt after=pt; L3 before=pt after=pt; actual L1 before=pt after=pt; L2 before=pt after=pt; L3 before=pt after=pt
- TOC per-level line-spacing mode/value baseline/actual: template L1 mode= value=pt; L2 mode= value=pt; L3 mode= value=pt; actual L1 mode= value=pt; L2 mode= value=pt; L3 mode= value=pt
- TOC per-level indentation chars/points baseline/actual: template L1 left=pt hanging=pt chars=; L2 left=pt hanging=pt chars=; L3 left=pt hanging=pt chars=; actual L1 left=pt hanging=pt chars=; L2 left=pt hanging=pt chars=; L3 left=pt hanging=pt chars=
- TOC per-level tab stop leader baseline/actual: template L1 tab=pt leader=; L2 tab=pt leader=; L3 tab=pt leader=; actual L1 tab=pt leader=; L2 tab=pt leader=; L3 tab=pt leader=
- TOC visible run typography baseline/actual: template title/text/tab/page-number visible run directRPr= eastAsia= ascii= hAnsi= cs= size=pt sizeCs=pt weight=; actual title/text/tab/page-number visible run directRPr= eastAsia= ascii= hAnsi= cs= size=pt sizeCs=pt weight=
- TOC per-level visible run typography baseline/actual: template L1 text visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L2 text visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L3 text visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; actual L1 text visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L2 text visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L3 text visible run directRPr= eastAsia= size=pt sizeCs=pt weight=
- TOC page-number run typography baseline/actual: template L1 page-number visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L2 page-number visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L3 page-number visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; actual L1 page-number visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L2 page-number visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L3 page-number visible run directRPr= eastAsia= size=pt sizeCs=pt weight=
- TOC tab/leader run typography baseline/actual: template L1 tab visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L2 tab visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L3 tab visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; actual L1 tab visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L2 tab visible run directRPr= eastAsia= size=pt sizeCs=pt weight=; L3 tab visible run directRPr= eastAsia= size=pt sizeCs=pt weight=
- TOC run typography verdict:
- TOC scale/compression verdict:
- TOC paragraph-and-typography verdict:
- TOC used-level inventory:
- TOC used-level evidence map:
- metric-by-metric comparison verdict:
- forbidden substitute evidence used?: no
- structural figure geometry validation report path:
- structural source-scale bbox map path:
- structural inserted-scale geometry evidence path:
- structural inserted-scale collision evidence path:
- structural dense-zone crop evidence paths:
- structural relation-attribute collision verdict:
- structural shape-overlap verdict:
- structural inserted-scale collision verdict:
- structural source-to-inserted geometry verdict:

## Checks
- checks performed:
- machine-vision verdict:
- TOC title baseline confirmed:
- TOC level formatting confirmed:
- TOC dotted-leader / right-tab confirmed:
- TOC page-number column confirmed:
- TOC restored from locked baseline confirmed:
- TOC page occupancy baseline confirmed:
- TOC title paragraph confirmed:
- TOC entries by level confirmed:
- TOC dotted leaders confirmed:
- TOC page-number column per entry confirmed:
- heading level baseline confirmed:
- heading direct-run typography confirmed:
- heading paragraph metrics confirmed:
- heading body-format residue cleared confirmed:
- heading TOC/chapter-start sync confirmed:
- table authority source confirmed:
- table manuscript binding confirmed:
- active table family confirmed:
- table-local structure clean confirmed:
- formula numbering same-line confirmed:
- formula numbering far-right confirmed:
- abstract surfaces confirmed:
- Chinese abstract title confirmed:
- Chinese abstract body confirmed:
- Chinese abstract mixed-script font chain confirmed:
- Chinese keyword line confirmed:
- Chinese keyword label/content run split confirmed:
- English abstract title confirmed:
- English abstract body confirmed:
- English keyword line confirmed:
- English keyword label/content run split confirmed:
- English abstract semantic parity confirmed:
- cover page-class baseline confirmed:
- cover identity-zone baseline confirmed:
- cover identity value-line baseline confirmed:
- declaration/title front matter baseline confirmed:
- declaration separated from cover confirmed:
- caption wording clean confirmed:
- caption baseline class confirmed:
- header/footer baseline confirmed:
- footer/page-number presentation confirmed:
- page-number structure confirmed:
- tail-block title baseline confirmed:
- tail-block opener fresh-page confirmed:
- tail-block separation from prior block confirmed:
- tail-block singular pagination owner confirmed:
- references title indentation confirmed:
- references entries indentation confirmed:
- acknowledgement title indentation confirmed:
- acknowledgement body indentation confirmed:
- end-matter rendered geometry confirmed:
- structural source-scale collision report confirmed:
- structural inserted-scale dense-zone review confirmed:
- result: pass | fail
- blocker:

## Notes
- summary:
