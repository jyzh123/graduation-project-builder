# User Feedback Persistence: Maintenance And Structure

Use this file for durable maintenance, file-structure, screenshot-asset, encoding, and rule-architecture corrections.

If the same thesis-scope complaint also includes bibliography repair or visible
reference label family issues, load
`references/user-feedback/citations-and-bibliography.md` as the sibling router
before acting.

## Enforcement Status

- Every numbered rule in this file is mandatory when this file is loaded for the current subtask.
- Apply these rules together with `references/user-feedback-persistence.md`.

## Table And Skill Structure Rules

### EXEC-MAINT-001 (legacy 40). Thesis Three-Line Table Variant Must Be Locked Explicitly (Mandatory)

- Do not hard-code one universal thesis three-line-table border family across all schools, templates, user screenshots, and editor presets.
- Before table generation or repair, lock the active table authority explicitly in this order:
- current user-identified built-in preset or sample
- current user-provided table screenshot or page image
- real local template table example
- accepted local manuscript sample
- stored fallback memory
- Do not assume that a thesis three-line table must always keep internal vertical separators.
- Do not assume that a thesis three-line table must always remove internal vertical separators.
- Treat `selected WPS table with pale internal guides` and `final printed border family` as different surfaces.
- If the user explicitly points to a WPS built-in preset such as the second three-line-table tile, that preset is the editor-style authority, but final acceptance still depends on the rendered unselected page or exported PDF.
- Table caption/title mode must be locked from the active donor before repair. A current user correction, approved sample, real template table, or locked table authority may require either a standalone caption paragraph outside the table or an in-table first merged title row; this maintenance rule must not override `references/thesis-table-style-memory.md` or a stronger current donor.
- If visible narrative prose such as `表7.1...`, explanatory sentences, or other body text appears inside a table cell, stop treating the issue as copy/style cleanup and repair the table-local structure first.
- Do not normalize, humanize, or rewrite prose that is still trapped inside a table cell; first move it back to the intended body or caption paragraph surface.

### EXEC-MAINT-002 (legacy 41). SKILL.md Must Stay As The Orchestrator, Not The Detail Dump (Mandatory)

- Do not keep expanding `SKILL.md` with every newly learned thesis-format edge case, checklist body, or visual micro-rule.
- `SKILL.md` should stay clean and act as the routing layer:
- mode split
- high-level workflow
- reference loading map
- gate and completion rules
- Move detailed rules into focused files under `references/` by topic, such as:
- thesis format SOP
- format class review
- table style memory
- figure rules
- user feedback persistence
- If a new rule is too detailed to help with first-step routing, it does not belong in `SKILL.md`.

### EXEC-MAINT-003 (legacy 42). Thesis Format Alignment Must Include Figure-Internal Style (Mandatory)

- During thesis format repair, do not treat figure formatting as only:
- image visibility
- image position
- caption position
- caption numbering
- The figure's internal style must also be aligned to the approved sample or template whenever the figure is not a raw runtime screenshot.
- Mandatory internal-style checks include:
- font family and visual weight inside the figure
- internal text-size hierarchy
- connector and border thickness
- fill behavior
- box padding and spacing rhythm
- overall visual density
- If the internal style is still off even though the figure is visible and the caption is correct, the figure block is still not complete.

### EXEC-MAINT-004 (legacy 43). Skill Markdown Files Must Use UTF-8 Without BOM (Mandatory)

- Skill rule files, reference files, checklist files, and related markdown assets must be stored as `UTF-8 without BOM`.
- Do not rely on the shell's default local code page or the editor's implicit save behavior for these files.
- When a rule file is created, rewritten, or normalized, explicitly preserve `UTF-8 without BOM` encoding.
- If a file shows mojibake risk or cross-environment encoding drift, normalize the file encoding before further rule edits.


### EXEC-MAINT-005 (legacy 44). File Role Index Must Distinguish Active From Backup And Archive (Mandatory)

- The skill directory must maintain a clear distinction between active files, backup files, archive files, and historical note files.
- Active maintenance must target only the active source-of-truth files unless the user explicitly requests archival cleanup.
- Backup and archive files must not be silently used as live rule sources when an active counterpart exists.
- If the skill directory structure is cleaned or expanded, the file-role index should be updated in the same turn.

### EXEC-MAINT-006 (legacy 45). System Chapters Must Have Authentic Full-Page Screenshot Assets (Mandatory)

- When a thesis includes system-description or implementation chapters that discuss real pages, views, or end-to-end flows, the project must possess authentic full-page system screenshots as evidence assets.
- This is an asset-availability requirement, not an automatic insertion requirement.
- Whether a given authentic screenshot is actually inserted into the thesis must still follow chapter semantics, template constraints, and current user direction.
- Cropped fragments, mockups, blank exports, or hand-drawn substitutes do not satisfy this requirement.

### EXEC-MAINT-007 (legacy 49). Never Use Fixed Word COM Paragraph Numbers For Thesis Body Or Reference Repair (Mandatory)

- Do not use fixed `Word COM` paragraph numbers as a write target for thesis body text, references, captions, or headings.
- `Word COM` paragraph collections include paragraphs inside table cells, text boxes, and other containers, so their numbering is not interchangeable with `python-docx` paragraph indexes.
- Do not mix `python-docx` paragraph indexes and `Word COM` paragraph indexes as if they refer to the same paragraph.
- For thesis body and bibliography repair, locate targets by verified text prefixes, style plus nearby context, bookmarks, or other structural anchors instead of fixed paragraph numbers.
- If a repair batch risks touching tables, first verify whether the candidate `Word COM` paragraph is inside a table before writing.

### EXEC-MAINT-008 (legacy 50). Figure Text Legibility Outranks Compact Layout (Mandatory)

- For thesis diagrams and other drawn figures, internal text must be judged by whether a reader can actually read it after the figure is inserted at thesis page scale.
- Do not compress layout, shrink labels, or preserve a crowded one-page composition at the cost of text legibility.
- If necessary, enlarge the figure, reduce density, wrap labels, or simplify the structure before acceptance.
- After insertion into the DOCX, run machine-vision inspection or an equivalent explicit readability check on the rendered page.
- If machine vision still cannot read the figure text clearly on the inserted page, the figure is not accepted even if the source image looked clear.

### EXEC-MAINT-009 (legacy 51). Focused References Are The Canonical Sink For Detailed Durable Rules (Mandatory)

- Keep `memory.md` as a short cross-project summary rather than a second full rulebook.
- Detailed durable corrections should be written into the most specific focused reference file that owns the topic.
- Use `references/user-feedback-persistence.md` when the correction is durable but not narrow enough for an existing more specific thesis or program reference.
- Do not dual-write detailed rules into both `memory.md` and a focused reference unless a short summary line in `memory.md` is genuinely necessary.

### EXEC-MAINT-010 (legacy 52). Split Overweight Rule Files Instead Of Letting Them Grow Indefinitely (Mandatory)

- If a focused markdown rule file becomes too heavy to maintain cleanly, split it by subtopic instead of continuing to append competing sections into one oversized file.
- Keep the parent file as the router or summary for that topic, and move dense subtopic rules into child markdown files with clear ownership.
- When a file split changes the active rule layout, update `SKILL.md`, `FILE-ROLE-INDEX.md`, and any validation logic in the same turn.

### EXEC-MAINT-011 (legacy 56). Figure Replacement Must Touch Only The Contiguous Image Block Immediately Above The Target Caption (Mandatory)

- When replacing an existing thesis figure, do not delete a broad nearby paragraph window around the figure number.
- Remove only the contiguous image-holder paragraphs immediately above the verified target caption.
- Stop deletion as soon as a non-empty non-image paragraph is encountered.
- Do not treat earlier body paragraphs that merely reference the figure as part of the replaceable image block.
- If no prior image-holder paragraph exists, insert a new dedicated image paragraph immediately before the caption instead of reusing a body paragraph.
- Do not chain a mutating insert plus `Paragraph.Next()` / `Paragraph.Previous()` traversal as the primary way to locate the new image paragraph, caption paragraph, or explanation paragraph in Word COM.
- After each insert, reacquire the intended local figure block by stable caption text, stable heading scope, or another verified structural sentinel instead of assuming the COM paragraph chain still points at the expected node.
- The image-holder paragraph must use a non-clipping layout configuration:
- do not use exact line spacing that can clip an inline image into a thin strip
- keep the image with its caption as one block during pagination
- If replacement logic can still move the wrong figure, delete the wrong paragraph, or clip the inserted image, the figure-repair automation is not safe enough for handoff.

### EXEC-MAINT-012 (legacy 60). Table Alignment Passes Must Preserve The Active Three-Line Table Border Strategy (Mandatory)

- When a thesis table is being centered or when cell content is being normalized, treat alignment repair as additive rather than as a replacement for the active border strategy.
- Do not leave a table in Word default full-grid form just because the current pass was focused on centering, pagination, or cell content alignment.
- Do not leave tinted header fills, colored cell shading, or other non-template background decoration in place just because the current pass was focused on borders or centering.
- If the active rule set or current user direction requires a three-line table, the centering pass must preserve or re-apply:
- top border
- header separator border
- bottom border
- inner vertical separators only when required by the locked active table rule
- outer left/right borders only when required by the locked active table rule
- body-row horizontal separators only when required by the locked active table rule
- A table pass that ends with correct alignment but wrong border family is still a failed table-format repair.

### EXEC-MAINT-013 (legacy 61). Table Caption, Header, And First Data Row Must Not Be Split Across Pages (Mandatory)

- Do not allow a table caption or table title to remain at the bottom of one page while the actual table body starts on the next page.
- Do not allow a table header row to remain isolated on one page while all data rows move to the next page.
- Minimum accepted table pagination unit:
- table caption paragraph
- header row
- first data row
- If that minimum unit does not fit in the remaining page area, move the full unit to the next page before handoff.
- For long tables that must continue across pages, preserve readable continuation behavior:
- keep the first page visually coherent
- keep the header visible with the following data
- avoid leaving only a thin tail row or a nearly empty page before the break
- If rendered pages still show a caption-only, header-only, or near-empty pre-break table artifact, the table-repair task is incomplete.

### EXEC-MAINT-014 (legacy 67). Existing Thesis Images Must Keep Original Media Relationships (Mandatory)

- When a thesis `.docx` already contains embedded figures, replacing the visible image content must preserve the original media relationship topology.
- Do not use a high-level image replacement path that can:
  - delete existing `word/media/*` parts
  - regenerate image relationship ids unnecessarily
  - rewrite image targets into malformed forms
  - reduce the media-part count after a nominal one-for-one replacement
- Required default behavior:
  - keep the original relationship id
  - keep the original `Target` path
  - replace only the media binary for the referenced image
  - change drawing size separately if required
- If a replacement attempt causes media count reduction or package-open failure in Word/WPS, treat that path as banned for future thesis runs until a relationship-preserving path is used instead.

### EXEC-MAINT-015 (legacy 70). Thesis Review Copy Must Be Singular And Explicitly Named (Mandatory)

- When a thesis workspace already contains many similarly named `.docx` drafts, do not assume the user will open the newest file by timestamp.
- For each review round, produce one clearly labeled review copy whose filename explicitly marks it as the current file to open.
- Do not describe a visual issue as fixed until that exact review-copy path has been checked against the expected visible heading, caption, or other sentinel text.
- If the user screenshot still shows an older title, caption, or comment state, first suspect wrong-file selection or stale preview before asserting that the edit logic itself failed.

### EXEC-MAINT-016 (legacy 71). Do Not Keep Rewriting A User-Opened Review DOCX In Place (Mandatory)

- If Word, WPS, or an embedded preview may already be holding the current review file open, do not keep mutating that same path in place during iterative QA.
- For late-stage thesis review passes, write the next pass into a new review-copy filename instead of racing the open file lock.
- Only promote or alias a newer review copy after the write pass is complete and the resulting file has been re-opened successfully by the document toolchain.
- Treat file-lock timeouts, partially applied edits, and stale preview cache as a review-copy management failure, not as proof that the content repair itself is complete.

### EXEC-MAINT-017 (legacy 72). Rule And Helper Text Files Must Be Saved As UTF-8 Without BOM (Mandatory)

- When writing project-local learnings, focused rule files, helper scripts, or skill-side persistence updates during graduation-project work, save them as UTF-8 without BOM.
- Treat encoding drift in these files as a real workflow failure, not as a harmless display issue.
- If a user explicitly asks for a correction to be written back into skill references, verify that the write path preserves UTF-8 without BOM in the same turn.

### EXEC-MAINT-018 (legacy 80). New Durable Thesis Corrections Must Be Merged Into Existing Focused Rule Files When Ownership Already Fits (Mandatory)

- When a new durable correction clearly belongs to an existing focused child file, append or merge it there instead of creating a new sibling file with overlapping scope.
- Preferred merge order:
  - expand the existing topic file
  - expand the router's rule-range reference
  - avoid introducing a new child file unless the current file would become structurally unclear
- Do not split one user correction across multiple files when one existing file already owns the topic cleanly.
- Avoid duplicate or near-duplicate rules that restate the same requirement in different child files.

### EXEC-MAINT-019 (legacy 81). Chinese-Heavy Helper Scripts And Local SQL Feeds Must Prefer UTF-8 No-BOM File Execution Over Shell-STDIN Injection (Mandatory)

- For Chinese-heavy helper scripts, SQL imports, or document-repair batches, do not assume shell inline text, here-strings, or stdin piping is encoding-safe enough.
- If the task has already shown BOM leakage, mojibake, SQL parser failure at line 1, or Python syntax failure from a hidden character, stop using inline shell text for that payload.
- Required fallback:
  - write the payload to a temporary UTF-8 without BOM file
  - execute the file explicitly with the chosen interpreter or client
  - verify that the temp file's encoding path is stable before continuing
- This rule applies especially to:
  - Python helper scripts
  - SQL import files
  - skill-side markdown/rule updates
  - any automation that carries Chinese thesis text through PowerShell

### EXEC-MAINT-020 (legacy 82). Thesis Runs That Depend On Real System Evidence Must Keep A Project-Local Route-To-Artifact Map (Mandatory)

- If the thesis content references real system pages, runtime screenshots, charts, or other evidence assets, keep one project-local mapping file that records:
  - thesis section or claim
  - real route or system page
  - screenshot or artifact path
  - whether the asset is inserted into the thesis or retained only as defense evidence
- This mapping file is a project-local execution artifact, not a global skill rule file.
- Use it to prevent drift between thesis wording, inserted screenshots, backup screenshot assets, and the actual running system.
- If a later thesis pass changes screenshots, routes, or page naming, update the mapping file in the same turn.

### EXEC-MAINT-021 (legacy 91). Once A Thesis Repair Run Starts Branching Into Many Draft Variants, Stop Patching Branches And Re-Collapse To One Canonical Source (Mandatory)

- If the workspace has already accumulated many repair outputs such as `v5`, `v8`, `final`, `fixed`, `toc-fixed`, `caption-fixed`, or similar variants, do not keep stacking new mutations on whichever file was touched last.
- Explicitly choose one canonical source manuscript, document that choice, and restart the repair flow from that single source.

### EXEC-MAINT-022 (legacy 92). Front-Matter Repair Must Rebuild Section Boundaries Explicitly Instead Of Reusing Stale Sample `sectPr` Topology (Mandatory)

- When a thesis build starts from a sample `.docx`, do not assume the sample's existing `sectPr` layout remains valid after body replacement, front-matter rewriting, or TOC reconstruction.
- If cover, abstracts, TOC, and main body are rebuilt or heavily edited, explicitly reconstruct the intended section boundaries for those zones rather than clearing content while leaving old sample `sectPr` nodes in place.
- Treat overlapping or drifting section ranges as a structural failure, even if the visible text still looks roughly correct.

### EXEC-MAINT-023 (legacy 93). Body Heading Baselines Must Come From Real Body Heading Instances, Not Abstract Titles Or Arbitrary Sample Offsets (Mandatory)

- For thesis heading repair or thesis generation, do not extract the final heading baseline from front-matter title paragraphs such as the Chinese abstract title, English abstract title, or TOC title.
- Do not rely on arbitrary sample paragraph numbers unless they have already been verified as:
  - the first real body chapter title
  - the first real body second-level heading
  - the first real body third-level heading when third-level headings are in scope
  - the first real body paragraph
- If body headings inherit spacing, alignment, numbering, or indentation from the wrong sample instance, treat the entire heading baseline as invalid and re-extract it from verified body paragraphs before continuing.
- If a script uses a paragraph such as `7.3.1` as the template source for rebuilding later `7.3.x` headings, verify first that the source paragraph is still a real heading-class instance rather than a polluted body paragraph.
- Treat uncontrolled draft branching as a workflow failure because it causes regressions, duplicate fixes, and contradictory formatting states.
- Final body-format audits must inspect every real post-TOC body heading by heading level, not only inherited style pass/fail. The audit must expose direct paragraph alignment, spacing, first-line/left indent, outline level, and direct run size for level 1, level 2, and level 3 headings.
- TOC/static leader entries and front-matter titles must be excluded from heading-level metrics; only actual body headings may satisfy or fail the body-heading gate.

### EXEC-MAINT-024 (legacy 94). If WPS Or Word COM Becomes Unstable, Narrow The Write Scope Instead Of Expanding The Repair Pass (Mandatory)

- When COM automation starts returning RPC failures, stale locks, or half-applied edits, do not react by making the next pass broader.
- Reduce the scope to one isolated surface such as:
  - TOC only
  - front matter only
  - bibliography only
  - captions only
- Reopen from a clean copy and verify that isolated surface before attempting another combined pass.
- Treat COM instability as a reason to tighten write scope, not to combine more repairs into the same risky run.

### EXEC-MAINT-025 (legacy 95). Reusable Thesis Repair Knowledge Must Be Summarized As Focused Operational Rules, Not As A Raw Session Log (Mandatory)

- When a thesis repair round produces many concrete lessons, do not persist them as an unstructured chronological diary.
- Convert the lessons into reusable operational rules grouped by topic such as TOC generation, heading demotion, citation rebuilding, bibliography order, and caption formatting.
- The stored knowledge should be directly actionable in future runs without requiring the reader to reconstruct the full history of one project.

### EXEC-MAINT-026 (legacy 96). WPS COM Page-Number Styling Must Fall Back To XML Or Narrow-Compatible Field Paths When `PageNumber.NumberStyle` Is Unsupported (Mandatory)

- Do not assume that Word COM and WPS COM expose identical page-number APIs.
- If the active Office COM layer does not support properties such as `PageNumber.NumberStyle`, do not keep retrying broad pagination passes against the same COM property path.
- Required fallback order:
  - insert or preserve the page-number field through the available COM range/page-number entry path

### EXEC-MAINT-027 (legacy 98). Planned Thesis Structural Figure Families Must Not Collapse During Delivery (Mandatory)

- If the thesis blueprint or figure plan names multiple structural figure families, do not treat completion of one family as closing the whole figure task.
- Each planned family such as architecture, ER, workflow, module tree, sequence, or use-case must be tracked separately as passed, failed, or explicitly skipped with a reason.
- A delivered manuscript that embeds only one surviving flowchart while the plan still requires other structural figures is incomplete.
- Do not hand off figure work by saying the figure set is complete unless every planned family has a recorded status and asset path.

### EXEC-MAINT-028 (legacy 99). Mermaid Quick Diagrams Cannot Stand In For Final Sample-Locked Thesis Structural Figures (Mandatory)

- When the active thesis figure family has a stored visual sample, draw.io source path, or draw.io-first requirement, Mermaid quick diagrams may be used only as planning artifacts.
- Do not treat Mermaid flowcharts, Mermaid `erDiagram` output, or similarly generic quick diagrams as final thesis structural figures unless a higher-precedence source explicitly approves that visual family.
- If a final structural figure still looks like a generic Mermaid diagram rather than the locked thesis sample family, the figure task has failed even if the semantics are roughly correct.
- For thesis ER diagrams, Mermaid crow's-foot `erDiagram` with table-style field lists does not satisfy a locked Chen-style ER sample unless the active sample explicitly requires crow's-foot notation.

### EXEC-MAINT-029 (legacy 100). Thesis Word Assembly Is Blocked Until Planned Structural Figures Pass Figure Review (Mandatory)

- Do not proceed to Word assembly, DOCX formatting, or thesis handoff while any planned structural figure is still missing, sample-unchecked, geometry-unchecked, or only present as a draft.
- Before assembly, every planned structural figure must have a source path, a figure family id, a pass/fail/skipped status, and a reason if skipped.
- If the figure plan lists four structural figures, assembly is blocked until all four figures have reached a resolved status.
- Completion of one flowchart never serves as a proxy for an ER diagram, architecture diagram, module tree, or other separately planned structural figure.

### EXEC-MAINT-030 (legacy 101). Thesis Structural Figures Must Use Draw.io As The Final Authoring Source (Mandatory)

- For thesis structural figures such as architecture diagrams, ER diagrams, business flowcharts, pricing flowcharts, use-case diagrams, sequence diagrams, module trees, and similar designed figures, the final authoring source must be a draw.io file.
- The preceding line is the default structural-figure path; the only supported exception is a current-user `CORE-FIGURE-019` material-only reuse lock.
- The final inserted image must come from a draw.io export that corresponds to that draw.io source unless material-only reuse is locked.
- Mermaid, ad hoc drawing scripts, quick sketch tools, and other non-draw.io paths may be used only for temporary planning and must not be accepted as final structural-figure assets.
- If a default-path structural figure has no draw.io source file, the figure task is incomplete even if a PNG or SVG already exists. If material-only reuse is locked, the figure task is incomplete unless the primary/supplemental material source and final embedded media SHA256 are bound.

### EXEC-MAINT-031 (legacy 102). Drawn Thesis Structural Figures Must Default To SVG-Primary DOCX Insertion, And Child Boxes Must Stay Inside Parent Frames (Mandatory)

- If a thesis structural figure is authored in draw.io and an SVG export exists, the default thesis insertion path must be SVG-primary rather than PNG-only.
- A PNG fallback may still be created by the renderer, but the final thesis package must not reduce the figure to a raster-only insertion path when an approved SVG source is available.
- For architecture, module-layer, and other parent-frame diagrams, every child box must stay fully inside its parent frame with visible padding on all sides.
- Before a draw.io-exported SVG is inserted into a thesis, remove any export fallback notice such as `Text is not SVG - cannot display` or other draw.io viewer-helper text that could render on the page.
- Treat any of the following as release-blocking figure failures:
  - a draw.io-backed thesis structural figure is inserted only as PNG
  - a child box extends beyond the boundary of its parent layer, lane, or frame
  - the rendered thesis page shows a draw.io SVG fallback notice or helper text inside the figure area
  - the final rendered thesis page shows a parent-frame diagram whose rightmost or bottom child box breaks the outer border
  - set numbering format and section start values through DOCX XML where needed
  - reopen and repaginate once after the XML change
- Treat a COM exception on page-number style properties as a compatibility limit, not as a reason to broaden the repair pass across unrelated surfaces.

### EXEC-MAINT-032 (legacy 103). Thesis Helper Scripts Must Lock One Exact Target DOCX Path Before Any Write (Mandatory)

- Do not let thesis helper scripts, batch mutations, or automation workers choose their target DOCX through timestamp freshness, wildcard matching, or similarly named temporary drafts.
- Before a script writes to a thesis manuscript, explicitly lock one canonical target path and one expected output path for that pass.
- If multiple review copies exist, choosing the target path is a required preflight step, not an implementation detail the script may guess later.
- If a script inserts or replaces figure blocks, also lock one exact local anchor scope for each block:
  - verified heading/subheading sentinel
  - verified local paragraph or caption sentinel
  - intended local block order such as `body paragraph -> image-holder -> caption -> explanation`
- Do not let a figure-insertion helper target a block through a broad prefix search such as `7.3` across the whole manuscript unless that match has first been narrowed to one verified local block in the current review copy.

### EXEC-MAINT-033 (legacy 104). One Protected Thesis Surface May Have Only One Active Write Owner Per Pass (Mandatory)

- For thesis repair, treat headings, TOC, chapter-start pagination, custom-layout result tables, captions, headers/footers, bibliography, and appendix structure as protected surfaces.
- Do not let a later generic script rewrite a surface already owned by a narrower custom repair script in the same pass.
- If two scripts would touch the same protected surface family, split the pass and verify the first script before running the second one.

### EXEC-MAINT-034 (legacy 105). Bulk Thesis Script Runs Must End With A Smoke Audit Before The Next Step (Mandatory)

- After any helper script or bulk DOCX mutation, do not continue directly into the next repair stage.
- First reopen the exact target DOCX and confirm at minimum:
  - the package still opens
  - the expected table count still matches
  - touched heading families still use the expected heading styles
  - TOC / bookmark-linked heading text did not silently drift
  - touched captions still remain attached to the intended figure or table
  - no touched body paragraph or table title/caption text has been pushed into a table cell
  - every touched local figure block still contains its own image object immediately next to the intended caption / explanation block rather than only the inserted text
- If that smoke audit fails, the run must stop and narrow scope before any later script runs.

### EXEC-MAINT-035 (legacy 106). Heading, TOC, And Chapter-Start Pagination Are One Linked Repair Surface (Mandatory)

- Do not treat heading-style repair, TOC repair, and chapter-start pagination as unrelated cosmetic tasks.
- When a thesis pass touches one of those surfaces, review all three together before acceptance.
- A run that restores heading styles but leaves TOC links stale or chapter openers drifting is still failing.

### EXEC-MAINT-036 (legacy 106A). Tail-Block First Pages Must Lock One Verified Opener Owner And Evidence Path (Mandatory)

- When a thesis pass touches `结论`, `参考文献`, `致谢`, `附录`, or another tail-block title page, do not let the opener ownership remain implicit.
- Explicitly lock:
  - the touched tail-block titles
  - one opener owner for each touched tail block
  - the rendered evidence path for the prior page and the opener page
- A run that records generic chapter-start checkpoints but leaves tail-block opener ownership or evidence implicit is incomplete.

### EXEC-MAINT-037 (legacy 107). Custom-Layout Thesis Result Tables Must Be Quarantined From Generic Restyle Passes (Mandatory)

- Once an empirical result table has been width-tuned or font-tuned for rendered readability, classify it as a custom-layout table.
- Do not allow later generic whole-document table-style passes to overwrite that table's width, font size, or readability tuning.
- If a generic table script must run later, it must explicitly skip those custom-layout tables or preserve their tuned state.

### EXEC-MAINT-038 (legacy 108). Table Text Offset Diagnosis Must Inspect Table And Cell Margins, Not Only Paragraph Indents (Mandatory)

- If thesis table text still looks visually indented after paragraph first-line indents have been reset, do not keep retrying paragraph-only fixes.
- Inspect and, when needed, reset:
  - table-level `tblCellMar`
  - cell-level `tcMar`
- Treat `paragraph indent fixed but cell margins still nonzero` as a real table-format failure.

### EXEC-MAINT-039 (legacy 109). Thesis PowerShell Helpers Must Not Hardcode Non-ASCII DOCX Font Attributes Through Unsafe Encodings (Mandatory)

- When a Windows PowerShell helper writes DOCX XML or Office COM properties that carry non-ASCII font-family names or visible Chinese text, do not rely on raw script-source literals alone.
- Accepted paths are:
  - clone formatting from a real template/sample paragraph and replace only text
  - pass the literal through a verified UTF-8-safe file or encoded-payload route
  - patch XML with a toolchain that preserves UTF-8 exactly
- After such a write, run a DOCX-internal font/encoding audit on the exact output file before rendered-page review.
- If the output contains corrupted font names such as mojibake in `w:rFonts`, treat the run as a tooling failure, not as a formatting success.

### EXEC-MAINT-040 (legacy 109B). Project-Local Thesis Repair Scripts Must Be Thin Wrappers, Not New Formatting Engines (Mandatory)

- Do not auto-generate a project-local PowerShell or Python script that re-implements thesis heading, abstract, figure, caption, table, bibliography, header/footer, or pagination rules outside the canonical skill bundle.
- If a local script is unavoidable, it must be a thin wrapper that only:
  - locks one exact target DOCX path
  - locks one narrow protected-surface family
  - delegates the actual mutation to canonical skill scripts/helpers
- A project-local script fails this rule if it contains its own broad paragraph-restyle matrix, its own figure/table formatting policy, or its own cross-surface rewrite logic.
- This check is a start-of-run preflight, not only a final acceptance check. For thesis lanes, run the canonical scanner on the real project root before generating or executing any project-local helper script.
- The scanner must cover common local thesis script locations such as `scripts`, `thesis_pipeline`, `generated_figures`, and other project-root subdirectories; scanning only `project/scripts` is not sufficient.
- Python DOCX rewrite signatures such as `paragraph.clear()`, `clear_paragraph()`, `set_paragraph_text()`, hardcoded `doc.paragraphs[n]`, `add_run()` after clearing, PIL-generated thesis figures, or local `add_picture()` insertion count as risky thick-script evidence unless the file is a thin wrapper delegating to canonical skill helpers.
- If risky thick local thesis repair scripts are already present in the project, classify the current thesis lane as contaminated instead of continuing on the current mutated final/review copy.
- A contaminated lane may only continue as audit-only long enough to extract evidence and write a restart plan; further thesis mutation must restart from a clean source manuscript under canonical helpers.

### EXEC-MAINT-041 (legacy 109C). Project-Local Thesis Repair Scripts Must Not Reclassify Captions, Table Titles, Image Holders, Or Table Cells As Body Text (Mandatory)

- A generated local thesis script must not set figure captions, table captions, continuation titles, image-holder paragraphs, or table-cell paragraphs to the body style family as a generic fallback.

### EXEC-MAINT-047. Local Thesis Adapters Are Data, Canonical Skill Scripts Are The Engine (Mandatory)

- General thesis-making scripts must be implemented in the canonical skill bundle, not regenerated inside each project or Codex workspace.
- Project-local thesis files may only carry template/project-specific adapter data, template profiles, content manifests, locked run manifests, or thin wrappers that call canonical skill scripts.
- A local adapter is allowed to record donor paragraph IDs, template fingerprints, surface labels, page-class expectations, project evidence paths, screenshot manifests, bibliography inputs, and locked output paths.
- A local adapter must not own generic DOCX assembly, font policy, heading policy, TOC policy, citation policy, table/figure policy, pagination policy, or final acceptance verdicts.
- Before use, every generated local adapter manifest must pass `scripts/validate_thesis_local_adapter.py`; if it fails, the run must extend the canonical skill path instead of broadening the local adapter.
- A run fails if a local script still applies one body style such as `正文` or `毕业论文正文格式` across those protected non-body surfaces.

### EXEC-MAINT-042 (legacy 109D). Project-Local Thesis Repair Scripts Must Not Use Fixed Exact Line Spacing For Image-Holder Paragraphs (Mandatory)

- Generated local scripts must not force image-holder paragraphs onto body-text exact line spacing such as `LineSpacingRule = Exact` with a fixed small line value.
- Image-holder paragraphs must stay image-safe blocks with single-line spacing and no clipping risk.
- If a local script formats image paragraphs with fixed exact line spacing, treat that script as blocked for thesis use.

### EXEC-MAINT-043 (legacy 109E). Comment-Driven Bibliography Repairs Must Prove Real Bibliography Mutation, Not Only Reinsert Or Preserve Comments (Mandatory)

- When a teacher/user comment says the bibliography is wrong, do not treat comment insertion, comment preservation, or comment relocation as evidence that the bibliography repair has been executed.
- The repair pass must prove real bibliography mutation on the exact target DOCX through bibliography baseline and numbering checks, not only through comment state.
- A pass fails if the bibliography comment remains but the bibliography entry baseline/numbering still has not been repaired.

### EXEC-MAINT-044 (legacy 109F). Project-Local Thesis Repair Scripts Must Not Guess Figure Assets Outside The Locked Route-Caption-Asset Map (Mandatory)

- A generated local script must not guess screenshot or figure asset paths from ad hoc local folders, timestamp freshness, or broad filename heuristics once the thesis figure lane has been classified.
- Required source of truth is the locked route-caption-asset map or another explicitly locked asset mapping artifact for the current run.
- If a local script hardcodes one project folder of images and bypasses that lock, treat the figure lane as unsafe even if some images happen to render.

### EXEC-MAINT-045 (legacy 110). Any Protected-Surface Text Mutation Must End With Baseline Replay And Same-Surface Review (Mandatory)

- For thesis repair, treat `content changed` and `format already safe` as separate states. A protected surface is not accepted merely because the new wording is correct.
- When a pass changes text on any protected surface, the required same-pass order is:
  1. identify the touched protected surface family
  2. lock the approved template / sample baseline instance for that exact family
  3. apply the text mutation
  4. replay the approved baseline onto the touched paragraph or run structure
  5. rerun the surface-specific review on the exact post-edit DOCX
- This rule applies at minimum to:
  - abstracts
  - keyword lines
  - citation-bearing body paragraphs
  - bibliography entries
  - bibliography-page title / opener surfaces
  - figure captions
  - table captions
  - code titles
  - header / footer text surfaces
- Do not accept a helper or script that only rewrites text and relies on document defaults, `Normal`, or prior direct formatting to "probably" keep the right layout.
- If no approved baseline instance can be locked for the touched surface, stop and record the repair as blocked on baseline evidence rather than silently approximating the style.

### EXEC-MAINT-046 (legacy 97). Runtime Screenshots And Code Screenshots Must Be Governed As Separate Figure Families (Mandatory)

- Do not apply the runtime-page screenshot rule set and the code-screenshot rule set as if they were the same artifact class.
- Runtime screenshots:
  - must come from the real running system
  - should preserve authentic page evidence assets, with Chrome full-page capture as the default browser path when applicable
  - must not be silently downgraded into `示意图` or `样例图`
- Code screenshots:
  - must come from real project code
  - must not be synthetic pseudocode panels or fake editor cards
  - should keep line numbers and enough code context to read the fragment
  - must follow the current crop rule from the user or approved sample, including code-pane-only crops when explicitly requested
- When the user reports that required algorithm code was not visible in the thesis, final acceptance must bind an algorithm-code visibility evidence record with the exact final DOCX SHA256, real source-code file path and SHA256, visible line range, visible code line count, and pass DOCX-binding and code-visibility verdicts. A dummy evidence file or pass-shaped prose cannot close this complaint.
- If the user changes the preferred code-screenshot crop during the run, treat that as the active override immediately instead of clinging to the earlier default capture framing.

### EXEC-MAINT-048. Skill Changes Must End With Unified Rule And Workflow Consolidation (Mandatory)

- After any `graduation-project-builder` skill edit, do not stop at patching the immediate file that failed.
- Run a consolidation pass that identifies the canonical owner file for the changed rule, then updates all required routing and enforcement surfaces in the same turn.
- The consolidation pass must check:
  - `SKILL.md` cross-cutting hard-gate wording when the rule changes global behavior
  - the focused owner file under `references/`
  - `references/rule-owner-map.json`
  - `FILE-ROLE-INDEX.md` when ownership, routing, or active-file roles change
  - affected templates under `assets/`
  - validators, generators, and selftests when the rule is mechanically enforceable
  - project/session state records when the change came from a user correction
- A skill-maintenance handoff is incomplete if the rule exists only as scattered notes, duplicate uncross-linked fragments, or script behavior without a documented owner.
- The acceptable final shape is one main entry path, one canonical rule owner, explicit routers/indexes, and validator or selftest evidence that the flow is still coherent.

### EXEC-MAINT-049. Skill Diagnosis Must Preserve Existing Rules And Name The Enforcement Gap (Mandatory)

- When auditing or repairing `graduation-project-builder`, first check whether the skill already has a rule for the reported behavior before calling the rule absent.
- A diagnosis must separate:
  - missing rule text
  - missing evidence fields
  - missing validator coverage
  - missing selftest coverage
  - stale workflow routing
- For TOC format, the existing per-level style-binding and baseline rules remain mandatory. If a TOC still passes with wrong font size, spacing, or a shrunken visual rhythm, classify the gap as missing paragraph-dialog / typography enforcement unless the owner files truly lack the style-binding rule.
- Do not weaken an existing rule while adding enforcement. Keep the old rule owner and add the missing template, validator, generator, or selftest coverage around it.

### EXEC-MAINT-050. Retired Thick Thesis Rewrite Scripts Must Stay Out Of Active Paths (Mandatory)

- When a skill audit finds an active canonical helper that formats protected thesis surfaces through fixed paragraph windows, fixed section counts, broad paragraph deletion, or stale sample `sectPr` replay, retire that helper from the active script tree instead of leaving it as a production or compatibility entry point.
- Retired thick thesis rewrite helpers must live only under `references/archive/` with their failure context, and active scripts must not import, shell-call, or list them as planned helper scripts.
- The bundle validator and selftest suite must fail if a retired helper such as `scripts/align_target_thesis.py`, `scripts/build_pass4_docx.py`, or `scripts/rebuild_complete_sample.py` reappears under active `scripts/`.
- New front-matter, TOC, figure, page-number, or body-build behavior must go through `scripts/build_canonical_thesis.py`, bounded surface helpers, template profiles, protected-surface evidence, and final validation rather than resurrecting a broad rewrite script.

### EXEC-MAINT-051. Whole-Document Pagination Evidence Must Come From Canonical DOCX Structure Parsing (Mandatory)

- A thesis acceptance record must not pass whole-document pagination from handwritten text such as `section/page/restart/link-to-previous passed`.
- The pagination evidence chain must include output from `scripts/inspect_docx_pagination_structure.py` with the canonical `graduation-project-builder.docx-pagination-structure.v1` schema, generator name, template and final DOCX SHA256 values, and parser-derived section/page-number/header-footer maps.
- The parser evidence must cover `sectPr`, `pgNumType`, section count, section boundary positions, header/footer references, inferred link-to-previous behavior, hard page-break/section-break maps, and footer PAGE-field containers.
- Rendered page counts, logical/physical page maps, TOC page sync, blank-page scans, chapter openers, and tail-block openers still require rendered evidence, but they cannot replace the DOCX structure parser provenance.
- Validators, generators, and selftests must reject whole-pagination JSON or final evidence that lacks this parser schema/provenance.

### EXEC-MAINT-052. Body Heading Levels Require Independent Protected-Surface Evidence (Mandatory)

- A thesis/article mutation may not close heading format as passed from `heading baseline summary`, `heading family summary`, style names, or a generic paragraph/typography audit alone.
- Final acceptance must include `body heading levels evidence path`, `body heading levels verdict`, level 1/2/3/4 heading verdicts, and a direct rPr/font/size/bold/spacing verdict.
- The high-risk thesis format matrix must include `body_heading_levels`; a missing row, generic evidence path, or reused catch-all evidence blocks handoff.
- The protected evidence record for `body_heading_levels` must confirm level baselines, direct-run typography, paragraph metrics, body-format residue cleanup, and TOC/chapter-start synchronization.
- The detector must compare heading style definitions and basedOn/effective font chains, not only the first visible run. Style-definition drift in Heading 1-4 must block delivery even when paragraphs still carry the expected style name.

### EXEC-MAINT-053. Skill Selftests And Thesis Helper Gates Must Use Executable Evidence (Mandatory)

- `scripts/selftest_skill_flow.py --suite fast` must stay a bounded lightweight suite for text gates, validator fixtures, and small DOCX fixtures. It must not launch Word, WPS, LibreOffice, or the full integration gate by default.
- `--suite integration` and `--suite all` must run integration coverage through one canonical `scripts/run_integration_gate.py --case all` path plus inventory parity, not by duplicating every integration wrapper case.
- Positive selftest fixtures must satisfy the current validator contracts. Do not use `.txt`, `.md`, `.html`, or other fake artifacts as substitutes for rendered page images, runtime screenshots, PDF evidence, figure assets, or PDF export evidence.
- For thesis generation, thesis revision, and format repair records, project-local helper preflight is a hard evidence field set. Acceptance and format-task records must carry the scanner report path, risk count, disposition, and canonical-source restart status before a final gate may pass.
- The project-local helper preflight report must be self-auditable: schema, project root, generated timestamp, scanner command including `--project-root`, scanner exit status, and risky script count must be present and consistent with the acceptance or format-task fields.
- If the project-local helper scanner finds thick thesis helper risk, the acceptable thesis-lane dispositions are audit-only, clean-source-restart-required, or clean-source-restart-completed. A pass handoff or repaired DOCX handoff is blocked until the run records a completed clean canonical restart, names the clean source path, and states that contaminated project-local helpers were not used.
- Template discovery and template-profile records must carry canonical generator, pre-mutation generation stage, generated timestamp, selected template path, and matching SHA256 fingerprint; a path-only template claim is not a lock.
- Active canonical `scripts/` files must not grow broad whole-paper DOCX rewrite behavior such as fixed paragraph indexes, `paragraph.clear()` bulk formatting, Word `Range.Style = Normal`, broad `word/document.xml` rewrites, or figure drawing plus DOCX insertion without a manifest route. Bounded canonical helpers need an explicit allowlist owner and surface scope.
- Runtime screenshot evidence must name the real route or page URL, capture method, readiness cue, accepted screenshot path, and caption-to-asset mapping.
- Runtime screenshot evidence is a required evidence class whenever chapter text, captions, or figure inventories claim real system screenshots.
- Runtime screenshot evidence must also prove full-window capture geometry: expected application window or viewport size, capture/window bbox, actual image width and height, coverage ratio, and a full-window capture verdict. A top-left client crop, partial-window bitmap, stale small thumbnail, or missing geometry verdict is failed evidence even when the image file exists.
- `source-preserved`, `preserved-existing`, or `no_image_mutation` only proves that the embedded media was not changed. It does not prove that a runtime screenshot is authentic. If the caption, figure family, nearby prose, or inventory says runtime/system/UI screenshot, the runtime evidence fields above remain mandatory even when the image is preserved from the source DOCX.
- Structural diagrams still require draw.io source, SVG primary export, and raster fallback instead of PIL, Mermaid, or hand-drawn PNG as the final source of truth unless the current user explicitly locks `CORE-FIGURE-019` material-only reuse.
- Table, front-matter, abstract, TOC, heading, reference, acknowledgement, and body-style pollution regressions must have validator or selftest coverage. A rule that only exists as prose is not enough for these repeat-failure surfaces.
- Front-matter and end-matter coverage matrices must be validated as surface-bound records with required rows and per-surface evidence, not accepted as generic path placeholders.
- Blank-page and near-empty-page evidence must come from rendered page-image ink metrics. A page-count-only statement cannot pass whole-document pagination.
- Blank-page and near-empty-page evidence must cover the whole rendered document, including body and tail pages; do not limit near-empty detection to front matter.
- TOC visual geometry evidence must come from rendered ink-pixel measurements of title, row boxes, leader/page-number columns, and occupancy rhythm. Fixed-proportion or synthetic TOC geometry cannot pass selftest, acceptance generation, or final validation.
- Page-class coverage matrices must bind each page class to its own target identifier, rendered page or region, and evidence path. Reusing one generic rendered evidence file across cover, abstract, TOC, body, figure, table, references, acknowledgement, and appendix rows is failed evidence unless the evidence file itself contains class-specific rows that the validator can inspect.
- Static TOC repair must not replace a whole one-run TOC row with a page number. A no-tab TOC row can be patched only when a trailing page-number segment is visibly separated by leaders or a wide whitespace gap.
- Front-matter artifact cleanup must preserve real cover/declaration/title tables. A table may be deleted only when it is an isolated instruction table with explicit delete or fill-instruction markers and no substantive thesis field values.
- Table-family formatting must operate on body tables with an immediately preceding standalone table-title paragraph. Missing adjacent title, title trapped inside a cell, figure caption trapped inside a cell, or body prose inside a cell is a structural failure, not a formatting target.
- Canonical thesis builder adapters must carry a 64-hex active template SHA256 fingerprint that matches the selected template path before dry-run validation or generation can pass.
- Hard-field measurement helpers must compute pass/fail verdicts from actual template-vs-target typography and geometry comparisons; they must not write `pass` only because a surface record exists.

### EXEC-MAINT-054. Audit-Lane Unavailability Must Become Explicit Sequential Fallback Evidence (Mandatory)

- If the requested multi-agent audit lane cannot be spawned because of thread limits, tooling unavailability, permission limits, or another runtime blocker, the run must not silently proceed as if multi-agent audit occurred.
- The active checklist, acceptance record, and handoff must record:
  - requested audit mode
  - attempted agent dispatch or reason dispatch was impossible
  - sequential audit fallback id or record path
  - exact scope reviewed by the sequential fallback
  - final audit verdict and unresolved gaps
- A validator, selftest, or final handoff that says `multi-agent passed` while only a single sequential review actually ran is a false pass.
- For skill-maintenance work, sequential fallback must still run the consolidation checklist from `EXEC-MAINT-048` and name the changed owner files, router exposure, rule-owner map entries, validators, selftests, and state records.

### EXEC-MAINT-055. Body-Style Repairs Must Trigger Cross-Surface Regression Freeze (Mandatory)

- Any thesis repair that changes or rewrites `Normal`, `正文`, docDefaults, theme fonts, TOC styles, table styles, caption styles, default paragraph spacing, or broad run/paragraph properties is a style-blast-radius repair.
- A style-blast-radius repair must create a pre-mutation protected-surface freeze and a post-mutation regression diff before final acceptance. The freeze must cover at least TOC title/entries/page-number runs, body heading levels, body text, table titles, table cells, reference title/entries, acknowledgement title/body, appendix title/body, headers, footers, and page numbers.
- The freeze and regression diff must be owned by a thesis mutation transaction record from `references/thesis/thesis-mutation-transaction.md`; detached freeze/diff snippets cannot pass the final gate.
- Final acceptance for a style-blast-radius repair must include a cross-surface regression evidence path, a pass cross-surface regression verdict, a pass TOC underline pollution verdict, and a pass table style regression verdict.
- A body-style repair cannot mark TOC or table evidence as `none`, `not-applicable`, sampled-only, or stale. TOC visible-run typography evidence and rendered table-family evidence are mandatory sibling checks whenever body defaults change.
- Selftests and validators must reject a body-style/Normal repair that passes body checks but lacks TOC underline and table-family regression closure.

### EXEC-MAINT-056. Format-Preservation Promises Require Chapter-Level Regression Closure (Mandatory)

- When the current user, teacher comment, task card, or agent wording promises `不要破坏格式`, `保持原格式`, `不改格式`, `preserve formatting`, or equivalent wording, treat that promise as a hard format-preservation contract.
- The contract applies to the edited target chapter, adjacent chapter openers, and non-target protected thesis surfaces. It is not satisfied by content correctness, style-name spot checks, exported PDF existence, or a single rendered page screenshot.
- Before a DOCX write under this promise, lock a chapter-level format baseline for every touched body chapter and nearby chapter boundary. The baseline must include heading styles, body paragraph style family, direct run/font metrics, paragraph spacing/indentation, figure/table/caption holders inside the chapter, chapter opener pagination owner, and rendered page evidence.
- After mutation, run the structured `chapter.format-preservation-contract` detector in `scripts/sample_self_check.py` against the exact final DOCX. A missing detector, missing evidence object, failed detector, damaged chapter body-style ratio, heading style drift, opener pagination loss, or non-target protected-surface drift blocks handoff.
- Final acceptance must include `format preservation promise verdict`, `chapter format preservation detector verdict`, `chapter format diff path`, `touched chapter rendered evidence paths`, and `non-target format preservation verdict`. These fields must be explicit, pass-shaped, and bound to the same final DOCX and SHA256 as the rest of the acceptance record.
- A local transaction that touches only keyword/front-matter run structure must not be forced to pass a body chapter detector solely because it freezes body or figure sibling surfaces. It may use a concrete `not-applicable-*` chapter detector verdict only when no body/chapter surface is targeted and no format-preservation promise is made; base transaction evidence and non-target protected-surface preservation still have to pass.

### EXEC-MAINT-057. Blocked Evidence Must Not Produce Pass-Shaped Handoff Fields (Mandatory)

- If a protected surface detector, inventory row, high-risk matrix row, transaction validator, figure contract, header contract, or rendered evidence verdict is failed, blocked, missing, stale, or needs manual review, every generated record for that lane must carry a blocked state.
- A generator must not write `audit_verdict: pass`, `handoff_status: pass`, `status: pass`, `blockers: none`, `known caveats: none`, or equivalent pass-shaped fields while the same record contains blocked protected-surface evidence.
- The final acceptance record, agent run manifest, format task card, and role task cards must agree on the blocked reason. A later validator failure is not enough if the generated handoff text still looks green to a human reader.
- Validators must reject any record where blocked protected-surface evidence coexists with pass-shaped handoff fields, even when `validation result` is already `fail`.
- Selftests must include at least one false-green fixture where the protected-surface verdict is blocked but handoff fields remain pass/none.

### EXEC-MAINT-058. Broad Style Replays Must Be Explicitly Scoped And Cannot Touch Front Matter By Default (Mandatory)

- A helper that replays template styles, `word/styles.xml`, `Normal`, docDefaults, body paragraph properties, or body run properties must require an explicit surface allowlist before it can run.
- Replacing `word/styles.xml` is a style-blast-radius action. It is forbidden in a default surface repair and must be requested through an explicit `styles` scope plus protected-surface freeze, rendered evidence, and post-mutation regression closure.
- Abstract, keyword, cover, TOC, header, footer, references, and acknowledgement run formatting must be preserved when the helper is not actively replacing text on that exact surface.
- A body-style normalization pass must not clear English abstract font slots, keyword label/content run separation, cover direct formatting, or other front-matter direct formatting as collateral work.
- If a helper cannot prove its write scope by changed ZIP parts and protected-surface evidence, its report verdict must be blocked or needs-protected-surface-review rather than pass.
- A run may not call itself `content-only`, `格式未破坏`, `same format`, or `format preserved` if it changed paragraph text in a body chapter and lacks chapter-level diff plus rendered evidence for that chapter. The correct status is blocked or audit-only until the format-preservation detector and transaction validator pass.
- When the user reports that a named chapter such as `第七章` has format damage, do not repair or validate only the visible screenshot area. The follow-up evidence must cover the whole named chapter range, its opener page, its first paragraph block, and sibling protected surfaces that could have been affected by the write.

### EXEC-MAINT-059. Review Artifacts And Citation Superscripts Must Not Be False-Passed (Mandatory)

- When a user reports that comments were deleted, tracked changes disappeared, citation superscripts were lost, or an audit lane passed a document despite those regressions, classify the incident as a whole-flow gate failure rather than a single detector miss.
- The corrective skill change must update all linked surfaces in one pass: `SKILL.md`, workflow routing, protected-surface contract, transaction record rules, agent audit rules, final acceptance templates, citation/report validators, and selftests or explicit validator coverage.
- The root cause must be recorded as one of:
  - missing protected surface
  - missing source-to-final evidence field
  - missing validator/selftest coverage
  - stale audit reuse
  - audit lane checked worker claims instead of independently comparing source and final DOCX
- A future thesis run cannot accept comment-driven revision evidence unless it proves both sides:
  - teacher/user comment content was implemented where requested
  - review comments, anchors, tracked changes, bookmarks, and citation superscript runs were preserved or explicitly disposed with user approval
- A handoff or audit card must fail if it passes only `comment coverage`, `citation audit result: pass`, or `officecli validate` while source-to-final review-artifact and citation-run diffs are missing.
- Citation-run diffs must preserve hyperlink/bookmark hosts only when the source marker had such a host. If the source citation marker was already a plain superscript marker without a hyperlink host, the diff must record source-existing-only hyperlink scope instead of falsely reporting a lost final hyperlink.
- A missing bookmark may be explicitly disposed only when the canonical review-artifact auditor receives a SHA-bound `empty-paragraph-bookmark-disposition` record proving the bookmark lived only on a deleted empty paragraph. The disposition must prove no visible text, image/table, section/page break, comment anchor, tracked change, field host, hyperlink, or citation marker was removed. Any other bookmark loss remains blocking.
- For `new-thesis-production` only, when an old-topic manuscript is used solely as a template/source carrier and the final manuscript intentionally replaces the subject, the canonical auditor may accept a SHA-bound `new-thesis-source-artifact-disposition` record. That record must bind source/final DOCX paths and hashes, confirm `selected_thesis_workflow=new-thesis-production`, list every allowed missing bookmark, field host, and hyperlink host, explicitly allow source-citation non-preservation, and require the final citation chain plus full-content reference entries to pass. This exception must not be used for `whole-thesis-revision`, `local-surface-repair`, comment-driven repair, or format-only work.

### EXEC-MAINT-060. Smoke Summaries Cannot Become Final Thesis Acceptance (Mandatory)

- A thesis acceptance record is final only when it is based on `assets/final-acceptance-template.md` and the exact record passes `scripts/validate_skill_gate.py --gate-record <that exact record>`.
- A file named `final-acceptance`, `final-verification`, `acceptance`, or `verification` is not final acceptance when it only checks `officecli validate`, PDF export, media counts, page-image existence, old-term counts, phrase presence, image-count deltas, or broad screenshot existence.
- Smoke-only records must be labelled as smoke or intermediate evidence, must carry `handoff blocked`, and must not set `handoff status: pass`, `known caveats: none`, `validation result: pass`, or equivalent pass-shaped fields.
- When user-reported or teacher-reported front-matter issues mention keywords or abstracts, final acceptance must include a user-reported issue ledger row and protected-surface evidence for `zh_keyword_line` and `en_keyword_line`. The evidence must prove label/content run separation, baseline comparison, and the exact final DOCX SHA256; visible paragraph text or paragraph-level style alone is not enough.
- When user-reported or teacher-reported issues mention figures, screenshots, generated images, sample images, placeholder images, or mismatched captions, final acceptance must include a whole-DOCX figure inventory, figure asset manifest, caption-to-asset mapping, and per-figure provenance review. Generated info cards, generated table images, schematic placeholders, or sample/mock images must not be accepted as runtime screenshots, code screenshots, data-result screenshots, or model-result screenshots.
- The project-local helper scanner must flag scripts that write pass-shaped `final-acceptance` or `final-verification` records from structural signals without invoking the canonical `validate_skill_gate.py --gate-record` path.
- The figure contract must treat Chinese captions and descriptions containing `截图`, `结果截图`, `样例图`, `示意图`, `模型训练结果`, `数据存储结果`, `数据清洗结果`, or equivalent wording as evidence-sensitive figure surfaces that need provenance, not as generic raster images.
- Validators or selftests must include a known-bad smoke acceptance fixture that says pass from structural signals only and prove that the canonical gate rejects it.

### EXEC-MAINT-061. Comment-Driven DOCX Revision Must Preserve Structure And Avoid Fixed Locators (Mandatory)

- A teacher/user comment-driven thesis revision is not a license to rebuild the whole manuscript, replay generic styles, or copy visible text into a new DOCX.
- If the source DOCX contains comments, comment anchors, tracked changes, bookmarks, fields, hyperlinks, or citation superscript runs, the mutation transaction must freeze and diff review artifacts before and after mutation.
- Comment-driven repairs must locate targets by comment id/anchor, caption text, heading scope, bookmark, field host, or another stable structural anchor. Fixed paragraph indexes, fixed Word COM paragraph numbers, `doc.paragraphs[n]`, and mutating `Paragraph.Next()` / `Paragraph.Previous()` chains are blocked.
- If the requested change cannot prove target anchors, source-to-final review-artifact preservation, and chapter-level format preservation, route the run to `audit-only` instead of mutating the DOCX.
- Validators and selftests must reject comment-driven transaction records that use fixed paragraph indexes or whole-document rewrite/rebuild wording without the `new-thesis-production` workflow.

### EXEC-MAINT-062. Figure Insertion Must Stay Out Of TOC And Front Matter (Mandatory)

- A thesis mutation may not insert or keep figure bodies, screenshots, drawings, figure captions, or generated image placeholders inside the cover, declaration, abstract, keyword, TOC, or other front-matter protected zones unless the official template explicitly contains that image surface and the transaction names it as an authorized protected-surface edit.
- Before any image insertion or replacement, the transaction must prove the target anchor is outside TOC/front matter with a `target_anchor_not_protected_surface_verdict` or equivalent pass verdict.
- The figure contract must scan the exact final DOCX and fail when a `w:drawing`, `w:object`, or legacy picture appears before the first real body chapter as an unapproved image surface.
- A caption count, media count, package validity check, PDF export, or generic rendered screenshot cannot override a protected-zone image insertion failure.

### EXEC-MAINT-063. Existing Figures May Not Be Redrawn Or Replaced Without Explicit Authorization (Mandatory)

- For an existing thesis image, the default action is preserve the original media relationship, target caption, and visible asset unless the user or teacher comment explicitly asks for replacement, redraw, or recapture.
- A replacement/redraw manifest row must record `mutation_intent`, explicit authorization source and scope, original/final relationship id, original/final media target, original/final media SHA256, target caption, target chapter or anchor, protected-surface location verdict, caption-to-asset binding, rendered evidence, and final DOCX relationship evidence.
- Structural replacement images must still obey the draw.io/SVG/raster figure contract. Mermaid, Pillow/PIL, hand-drawn PNG, or other quick raster drafts are not acceptable final sources for structural thesis figures.
- If any existing media relationship hash changes without the authorization and source-to-final evidence above, the figure lane is failed even when the new image looks plausible.
- The figure contract must scan all Word package image relationship parts under `word/**/_rels/*.rels`, including body, header, footer, comment, footnote, and endnote story parts. Scanning only `word/_rels/document.xml.rels` is not enough.
- Any final DOCX or manifest that indicates image insertion, replacement, redraw, recapture, or media diff requires top-level manifest bindings for `source_docx_path`, `source_docx_sha256`, `final_docx_path`, and `final_docx_sha256`; missing or mismatched bindings fail closed.
- Authorized image changes must bind both sides: original owner part/rid/target/SHA256 from the source DOCX and final owner part/rid/target/SHA256 from the exact final DOCX. Duplicate rid/target/SHA identities across story parts require an owner-part field.
- Source media removal and final-only media addition are rejected unless the manifest carries explicit authorization for that exact change. Only when both source and final DOCX have no media may media-replacement checks be skipped.
- The figure contract must also compare source/final DOCX drawing-object manifests. Changes to DrawingML size, VML/WPS `w:pict` shape size, inline/anchor mode, relationship id set, or image-to-caption adjacency are image mutations and require explicit drawing authorization plus original/final drawing SHA256 bindings, even when media hashes are unchanged.
- Available-text-width checks apply to inserted or changed drawings in every Word story part, not only body `word/document.xml`. Header, footer, comment, footnote, and endnote drawings must fail when their final display width exceeds the document text width unless the exact drawing is preserved from the source DOCX.
- Figure display-height checks are also mandatory. A drawing whose final extent occupies more than the safe page-height threshold must fail the figure contract even when its width is inside the text area, because it can push captions or explanations onto unstable pages.
- A helper that only changes visible picture size still performs a drawing mutation. Generic surface-format helpers may not accept `image_display_resize` or similar display-size plans unless the call is owned by a transaction/figure-manifest wrapper that binds source DOCX, final DOCX, original/final drawing hashes, and a pure-size comment ledger row.
- Helper scripts that replace pictures directly must require a figure manifest and source DOCX path, validate the manifest after mutation, and limit the target to the dedicated image paragraph immediately before the verified caption. Broad nearby-paragraph scans are banned.
- Validators and selftests must reject unauthorized existing-image replacement, structural Mermaid/Pillow final sources, and DOCX figure insertions that land before the first body chapter.

### EXEC-MAINT-064. Comment Done State Is Not Semantic Completion Evidence (Mandatory)

- Word/WPS comment `done` or `resolved` state is a UI marker only. It is not evidence that the teacher/user comment was implemented.
- A comment-driven thesis run must keep a comment-resolution ledger that binds the exact source DOCX, final DOCX, source SHA256, final SHA256, every comment id, final status, done-state authorization, evidence paths, canonical surface, subissue, detector id, detector report path, detector verdict, and blocker state.
- A fixed ledger row is invalid if it only names a prose evidence path. It must bind the teacher/user comment subissue to a detector report with a pass verdict for the exact final DOCX, or the row remains open/partial/blocked.
- Any `commentsExtended.xml` done-state change must be authorized by a fixed ledger row for the same comment id. Missing ledger, missing evidence, open/partial status, orphaned comments, or comments without matching extension-state evidence block all-comments-resolved claims.
- A final `all-comments-resolved` or `批注已改完` claim must also prove that the exact final DOCX has zero open comments in `word/commentsExtended.xml`. Fixed ledger rows alone are not enough when Word/WPS would still show the teacher comments as unresolved.
- A figure comment that mentions crop, source/provenance, model structure, redraw, content, explanation, readability, or reference support cannot be closed by size-only evidence. Display resizing may close only a pure size subissue.
- Final acceptance must fail when a handoff says comments are complete while the ledger has open, partial, blocked, unknown, missing, or orphan rows, or while the ledger/audit was generated before the last DOCX mutation.

### EXEC-MAINT-065. Explicit Skill Invocation Must Start A Fail-Closed Run Lock (Mandatory)

- When the current user explicitly invokes `graduation-project-builder`, the first execution artifact must be a skill-invocation lock created from `assets/skill-invocation-lock-template.md`.
- The lock must be created before any file mutation, final handoff, or project-local helper execution. Reading `SKILL.md` is not enough.
- The lock must record the invocation source, selected mode and subtask, project root, loaded entrypoint, routed references, active checklist path, agent manifest or sequential fallback evidence, helper preflight report, helper risk count and disposition, mutation transaction path when a DOCX can be touched, mutation allowed verdict, exact output path/SHA when known, and final gate record/command/verdict before completion.
- If the lock is missing, incomplete, stale, or blocked, the run is audit-only. It may diagnose, plan, or repair the canonical skill bundle, but it must not mutate a thesis DOCX or claim final acceptance.
- For thesis generation, thesis revision, and format repair, a risky project-local helper scan forces the lock verdict to `blocked` or `audit-only` unless a clean-source restart or canonical-helper replacement has been recorded in the lock and in the final acceptance record.
- Final acceptance must cite the lock path and lock verdict. Validators must reject pass-shaped final records when the lock path is blank, the lock is unreadable, the lock is blocked, or the lock contradicts project-local helper preflight evidence.

### EXEC-MAINT-066. DOCX Helpers Must Preserve OpenXML Compatibility (Mandatory)

- A thesis DOCX repair is not acceptable when the produced package is only tolerated by one editor but fails schema-sensitive OpenXML validation or cannot be loaded by the configured renderer.
- Canonical helpers that rewrite `word/document.xml`, `word/styles.xml`, headers, footers, comments, footnotes, endnotes, or other XML parts must preserve WordprocessingML property child order and namespace declarations required by `mc:Ignorable`.
- OpenXML compatibility recovery must include builtin styleId canonicalization when WPS or another editor renumbers built-in styles, for example converting `1` back to `Normal`, `6` back to `TOC1`, and `7` back to `TOC2` before validators compare body, heading, and TOC style ownership.
- Recovery passes may use `scripts/repair_docx_openxml_compat.py` only as a bounded package-preserving repair: selected XML parts, no visible text edits, no media changes, no relationship changes, no comment disposal, and no field/TOC rewrite.
- A final thesis handoff cannot treat PDF export failure as a renderer-only problem until the exact DOCX has been checked for OpenXML compatibility errors and any helper-induced property-order or missing-namespace drift is either repaired or recorded as a blocker.
- Validators and selftests must include a fixture where invalid `w:pPr` / `w:rPr` / `w:style` child order and missing `mc:Ignorable` namespace bindings are repaired without changing thesis content.

### EXEC-MAINT-067. Live TOC Requirement Must Not Be Downgraded By Acceptance Generation (Mandatory)

- Whole-thesis production, whole-thesis revision, and submission-style thesis repair records must preserve the live TOC requirement unless the user explicitly approves a static-TOC exception and the handoff is blocked or labelled as static-fallback.
- A final acceptance generator must not hard-code `live TOC required this round?: no` for whole-thesis outputs. It must inspect the exact final DOCX for a standard TOC field instruction and record the live TOC field count and verdict.
- Static TOC entries made of `HYPERLINK` / `PAGEREF` rows are not a standard live TOC field. They may be recorded as intermediate compatibility evidence, but they cannot satisfy a live TOC pass claim.
- The gate must reject pass-shaped whole-thesis records that mark live TOC as not required, or that claim live TOC pass while the exact final DOCX lacks a `TOC` field instruction.

### EXEC-MAINT-068. Sample Self-Check Detectors Must Bind The Intended Source And Surface (Mandatory)

- `scripts/sample_self_check.py` figure-contract checks must use the `source_docx_path` and `source_docx_sha256` bound in the figure manifest when those fields exist. They must not silently validate figure media, drawing size, or caption adjacency against `reference_docx` as a fallback source.
- Sample self-check caption detectors must distinguish formal captions from explanatory body prose. A body paragraph such as `图3-1展示了...` or `图 3-1 展示了...` is a narrative reference, not a caption, and must not create a missing-caption, orphan-drawing, or figure-follow-up failure by itself.
- Body-style audits and comment-repair helpers must use the same distinction: paragraphs such as `图3-1展示了...`, `图4-4中的...`, or `图5-3从...` remain body prose and must still receive body line spacing, first-line indent, style binding, and mixed-script font checks. Only formal caption forms such as `图3-1  系统流程图` or `图3-1：系统流程图` may be excluded from body-prose auditing.
- Comment-content repair helpers must fail closed when a plan targets media replacement, image insertion, or drawing display extents without `--source-docx`, `--transaction-record`, and `--figure-manifest`; when a body-format repair anchor is non-unique or otherwise produces a non-unique anchor match; when the target is a protected non-body surface such as abstract, keyword, caption, TOC, references, or acknowledgement; or when `all_ascii_runs` font repair lacks an explicit `surface_allowlist` or explicit global authorization.
- Cover identity value-line checks must bind only to real cover identity fields and must exclude bibliography legends, reference-format tables, and rows such as `文献类型` / `参考文献类型` from cover value-line candidates.
- Body image-size detectors must reject disabled or placeholder wording that claims image dimension checks are off, skipped, or not applicable when body images exist. A disabled-text pass is not equivalent to measured drawing extents.
- Final-DOCX figure extent audits must also reject undersized structural figures. Architecture diagrams, process chains, rule-logic diagrams, risk-mapping diagrams, evidence-chain diagrams, and similar structural figures must meet the minimum readable width and structural minimum readable height instead of passing only because they are not oversized.
- Abstract and keyword donor checks must reject donor pollution from TOC rows, visible template instruction notes, bibliography/reference legends, red annotations, or the target paragraph itself when a separate template/reference donor is required.
- Every new sample self-check detector for these surfaces must have a named `scripts/selftest_skill_flow.py` case and an owner-map entry before the detector can be treated as durable.

### EXEC-MAINT-069. Baseline, TOC, And Figure Family Detectors Must Be Gate-Required (Mandatory)

- `scripts/sample_self_check.py` must emit Detector Registry entries for `heading.baseline-contract`, `toc.visible-format-contract`, and `figure.family-style-contract` on every thesis self-check report, using explicit evidence objects even when a surface is not applicable.
- `scripts/validate_skill_gate_record_gate.py` must require those three detector ids through `REQUIRED_SAMPLE_SELF_CHECK_DETECTORS`, the same required-detector path that requires `chapter.format-preservation-contract`; a pass-shaped final record must fail when any required detector is missing or lacks an evidence object.
- These detector ids are part of the durable sample-self-check closure chain, not separate prose-only rules. Do not create duplicate owner fragments in thesis layout, TOC, figure, or final-QA files unless a more specific surface rule is being changed.
- A pass037 maintenance record must explain whether multi-agent review was authorized, whether the source was clean or contaminated by project-local helpers, and whether the repair stayed canonical-helper-only. This record is maintenance evidence only and must not mutate a thesis DOCX.

### EXEC-MAINT-070. Scoped Thesis Repair Cannot Become A Release Baseline Until Promotion Gates Close (Mandatory)

- A local-surface, specialty, content-only, or otherwise scoped thesis repair is a candidate artifact until the release/baseline promotion gate is explicitly closed. A scoped pass may only claim the scoped target surfaces; it must not become the next master manuscript, release baseline, or handoff baseline by implication.
- Final acceptance and format-task records must carry `baseline promotion status`, `baseline promotion evidence path`, `release blocker ledger path`, `unresolved release blocker count`, and `scoped artifact next-baseline verdict`. Missing fields, nonnumeric blocker counts, unresolved blockers, or candidate/audit-only next-baseline wording blocks pass-shaped handoff.
- Before any scoped artifact is promoted, the run must bind the target surface list and the sibling protected-surface list, rerun executable audits for sibling/cross-surface regression, whole-document pagination, high-risk surface matrix, table/figure/caption/citation surfaces when present, and all known user-reported issues.
- Risky project-local helper preflight keeps baseline promotion blocked unless the final record proves a completed clean-source restart or canonical-helper-only restart. A contaminated or helper-risky scoped candidate must be recorded as audit-only or candidate-only, not as a release baseline.
- Generators and validators must reject records that combine `handoff status: pass` with unresolved release blockers, non-pass baseline promotion status, missing release-blocker ledger evidence, or a scoped-artifact next-baseline verdict that says candidate-only, audit-only, blocked, unverified, stale, or not promoted.

### EXEC-MAINT-071. Explicit Skill Invocation Cannot Be Downgraded To Reference-Only Execution (Mandatory)

- An explicit `graduation-project-builder` invocation is a workflow takeover event. It is not satisfied by reading `SKILL.md`, quoting the skill, or using the skill as background advice while continuing through ad hoc local scripts.
- The only actions allowed after recognizing explicit invocation and before the skill-invocation lock exists are: read the routed skill entry/reference files, create the lock, create the active checklist, create the agent/audit record, and classify whether the run is mutation-allowed or audit-only.
- If a previous continuation turn lacks a current lock or the lock is stale, blocked, missing routed references, missing active checklist, missing audit record, or missing canonical gate binding, the controller must restart from the lock before any new mutation or final handoff.
- A self-authored narrow checker, smoke checker, screenshot matrix, `officecli validate`, PDF export, page count, phrase scan, field-leak scan, or local helper report cannot replace `assets/final-acceptance-template.md` plus `scripts/validate_skill_gate.py --gate-record <exact record>` for a substantial handoff.
- Any failed, blocked, stale, missing, or contradictory evidence generated under the invoked skill must dominate the handoff state. It may not be moved into `residual notes`, `known caveats`, or a prose explanation while the handoff remains pass-shaped.
- The skill-invocation lock and final acceptance record must carry anti-bypass fields for invocation source type, skill activation status, rule-engine takeover, prohibited bypass checks, canonical gate requirement, narrow/smoke substitute use, failed-evidence escalation, no mutation before lock, and final handoff allowance.
- Validators and selftests must reject pass-shaped records when these anti-bypass fields are missing, contradictory, say that a narrow/smoke substitute gate was used, or record blocked evidence as caveat-only.

### EXEC-MAINT-072. Explicit Invocation Bootstrap Must Precede Project Work (Mandatory)

- After recognizing an explicit `graduation-project-builder` invocation, the run must enter a bootstrap phase before any ordinary project work.
- Bootstrap actions are limited to routed skill reading, skill-invocation lock creation, active checklist creation, agent/audit record creation, mode classification, and mutation-allowance classification.
- Project inspection, repository-wide text search, DOCX/PDF inspection, browser checks, system startup, helper-script execution, smoke tests, generated evidence, and user-facing final summaries are non-control actions. They may begin only after the lock/checklist/audit bootstrap exists and names their scope.
- If a non-control action happened before the bootstrap artifacts existed, the current run is contaminated/reference-only drift. The controller must record that drift, stop substantive execution, and restart from a fresh lock before mutation or handoff.
- A contaminated explicit-invocation run may continue only as canonical skill maintenance or audit-only diagnosis until the fresh lock records the contamination disposition and the routed rules are active.
- Handoff text must not claim that the skill was truly invoked unless it can name the current lock path, active checklist path, agent/audit record path, and final gate binding created before non-control execution.
- The skill-invocation lock and final acceptance record must carry `no non-control action before lock?`; a pass-shaped handoff must fail when that field is missing or anything other than `yes`.

### EXEC-MAINT-073. Pagination And Reference False-Pass Evidence Must Fail Closed (Mandatory)

- Whole-document pagination evidence must not clear `raw_unexpected_near_empty_pages` merely because `--allow-content-growth` is active. Each raw unexpected near-empty page requires an explicit page allowlist entry plus page-class/root-cause evidence; otherwise it remains a blocker.
- Acceptance generators must not rewrite unexpected near-empty pages into `unexpected_near_empty_pages=[]` after content-growth reconciliation.
- Protected-surface evidence for `references_title` and `references_entries` must compare template and actual hard metrics for style, font, size, weight, alignment, first-line indent, hanging indent, left/right indent, and tab stops. A pass-shaped verdict cannot override numeric drift in these fields.
- Reference-entry font audits must not use an explicit half-point override when a template/reference DOCX already provides a bibliography donor size. Use the template-derived size, or use an explicit named-size policy with all-entry WPS/Word evidence when the current user or official school rule requires that named size.
- Reference-entry scans must stop at the first real terminal sibling block such as `附录`, `致谢`, `谢辞`, `Acknowledgements`, or `Appendix`. A numbered/list-form appendix or acknowledgement title must never be counted as a bibliography entry merely because it appears after `参考文献` and before another tail block.
- `sample_self_check.py` must treat Word automatic numbering (`w:numPr`) as a valid bibliography baseline source. A template that uses automatic bibliography numbering must not downgrade reference entry or tail-block baseline checks to `not-applicable`.
- Reference-section pagination is not proven by page count, blank-page absence, or a generic `tail block rendered map`. The `references` opener must have an explicit rendered physical page, `references_page_found=yes`, `references_fresh_page_verdict=pass`, `references previous content physical page=...`, `references_prior_block_separation_verdict=pass`, and `references_opener_owner_evidence` bound to `sample_self_check.py` detector `tail-block.pagination-contract`.
- `sample_self_check.py` must fail closed when the `references` opener has no single pagination owner (`w:pageBreakBefore` on the opener, or exactly one previous-paragraph page/section break), when the previous real content block cannot be mapped to a rendered physical page, or when the previous real content page is the same as or after the formal `references` opener page. Missing, duplicated, stale, or generic opener-owner / prior-block separation evidence is a blocking defect.
- The same `tail-block.pagination-contract` detector must cover acknowledgement and appendix openers when those blocks exist. Appendix must not be audited as a body-chapter opener, and a final report must not pass if the appendix opener lacks a single pagination owner, renders before or on the same page as acknowledgement, or is absent from detector evidence while appendix content exists.
- Acceptance generators and validators must reject whole-document pagination JSON or final records that hide lost reference pagination behind pass-shaped fields, including `references previous content physical page=missing`, `references physical page=missing`, `references_page_found=no`, `references_fresh_page_verdict=fail`, `references_prior_block_separation_verdict=fail`, missing `tail-block.pagination-contract`, or a `tail_block_opener_page_map` without the required reference-opener and prior-block separation tokens.
- Validators and selftests must reject: unapproved raw unexpected near-empty pages, reference title first-line drift, reference entry title-as-entry/target drift, normal-to-bold reference typography drift, explicit half-point size override against a template donor, automatic-numbered bibliography templates that become not-applicable, missing tail-block pagination detector evidence, reference-opener pagination loss, and pass-shaped reference maps that omit the previous-content page or prior-block separation verdict.

### EXEC-MAINT-074. Body Opener And Running Header Titles Must Be Cross-Checked (Mandatory)

- Any thesis repair that touches page flow, section boundaries, chapter-start owners, body headings, headers, footers, end-matter openers, abstract or TOC pagination, style-blast-radius surfaces, or any user-reported title/header/page drift must treat body opener title text and running header/footer title text as one protected cross-surface family.
- Final acceptance cannot pass when the rendered body opener names one chapter or section while the running header/footer on the same physical page names another chapter or section.
- The canonical failure example is a rendered body opener/title `绪论` with a running header/title `结论`; this is a hard failure even if the local surface that was edited appears fixed.
- Evidence must come from rendered physical pages, or from an equivalent sentinel-mapped physical page review that binds body opener text, running header/footer text, expected title, observed title, and final verdict to the exact final DOCX/PDF.
- Final acceptance must include `body opener/header title consistency evidence path` and `body opener/header title consistency verdict` whenever title, heading, header, footer, chapter opener, pagination, or cross-surface drift is reported or can be affected by the mutation.
- User-reported issue ledgers must include this surface whenever the user reports title, header, footer, chapter opener, section opener, page-flow, pagination, or repeated cross-surface drift.
- Validators and selftests must reject missing title/header consistency evidence, pass-shaped records whose title/header verdict is missing, blocked, stale, sampled-only, or not checked, and evidence that does not name both the body opener/title and the running header/footer title.

### EXEC-MAINT-075. Protected Visual Surface False-Passes Must Fail Closed (Mandatory)

- A thesis handoff cannot pass when TOC typography, abstract typography, header line/text, footer, page-number position, or any other protected visual surface is checked only by structural signals such as field presence, style-name presence, page count, PDF export success, page-image existence, or body-style audit success.
- `scripts/sample_self_check.py` must emit the blocking detector `header-footer.page-number-template-contract` for header/footer/page-number template parity and must keep TOC visible format, abstract template-style, heading baseline, body style, table/figure, tail-block pagination, and common pre-submission detectors gate-required.
- Every missing or failed required detector must be recorded as a sample self-check detector gate failure in generated acceptance records.
- `scripts/generate_thesis_acceptance_record.py` must use the same required detector list as `scripts/validate_skill_gate_record_gate.py`. If any required detector is missing/failed, if the font audit fails, or if the exact `validate_skill_gate.py --gate-record` run fails, the generated acceptance record and skill lock must switch all handoff fields to blocked/fail.
- Validators must reject `passed with limitations`, `core checks passed`, `structural pass only`, `font audit limitation accepted`, or similar caveat wording when protected visual surfaces are in scope.
- Selftests must include false-pass fixtures for missing header/footer/page-number detector evidence, abstract/header/footer visual complaints without rendered geometry binding, and limitation wording that tries to keep a pass-shaped handoff.

### EXEC-MAINT-076. Whole-Thesis DOCX Structural Format Gate Is Required For Release (Mandatory)

- A whole-thesis generation or whole-thesis revision cannot pass final acceptance from isolated checks for text length, reference count, media count, PDF export, font color, body style, live TOC presence, or page screenshots. It must also pass the exact-output whole-DOCX structural format gate.
- The structural gate must bind to the handed-off DOCX path and SHA256 and must check, at minimum, section topology, front-matter/body/end-matter order, live or template-authorized TOC implementation, footer PAGE fields, front-matter/body page-number chain, header/footer part binding, builder-owned style contamination, and excessive style-less body paragraphs.
- A school/template static running header, such as an institutional thesis header reused on cover, TOC, and body sections, is not automatically a section leak. The whole-format gate must distinguish safe static institutional header text from body/TOC heading leakage, visible PAGE fields on the cover section, unresolved placeholders, and chapter-title contamination.
- When the locked donor is a converted template with empty front-matter headers, the gate may record that an empty converted template may allow safe static institutional header text only when the final header is the approved institutional header and no body/TOC heading leak is present.
- Whole-format repair helpers must derive header text and header relationship topology from the current DOCX, locked template, or approved sample; in short, header repair derives from current DOCX or template. They must not hard-code a school name, `目 录`, or the first body chapter title as a global header repair policy.
- TOC/body section repair must recognize localized heading styles, localized Heading1 style names, and outline levels, including Chinese `一级标题` style names or template-local style ids, before deciding that the body start is missing.
- Whole-format reference-entry evidence must use formal bibliography-entry recognition, not `w:numPr` alone. List/numbering properties on `附录`, `致谢`, or other terminal titles are tail-title format defects, not bibliography-entry rows.
- If the user reports blue text, cover format errors, TOC errors, header/footer errors, page-number errors, global style drift, or mixed table/figure/body formatting, the run must treat the complaint as a whole-document structural risk until `scripts/audit_docx_whole_format_gate.py` passes on the exact final DOCX.
- A gate report from an earlier draft, a review copy that is not promoted to the final path, a manually edited JSON report, or a report whose SHA256 does not match the final DOCX cannot satisfy release acceptance.
- Final acceptance must include `final DOCX whole-format structural audit path` and `final DOCX whole-format structural audit verdict`; validators must reject missing, failed, stale, path-mismatched, or SHA-mismatched reports for whole-thesis workflows.

### EXEC-MAINT-077. Body Run Font-Slot Pollution Must Not Hide Behind Passing Style Binding (Mandatory)

- A body paragraph may still be visually polluted even when it is bound to the correct body style, because direct `w:rFonts` on visible runs can override the style family.
- Body-style, content-repair, and surface-format repair evidence must treat visible body run font slots as part of the protected body text surface whenever the user reports a paragraph that "looks different" from nearby正文.
- A repair that splits mixed Chinese/Latin runs must assign Chinese-visible segments to the approved East Asian body font slot and Latin/digit segments to the approved Western body font slot. It may not let Chinese text inherit `Times New Roman`, theme-only ownership, or a copied polluted run model.
- Citation runs, hyperlink wrappers, bookmarks, fields, drawings, comments, captions, references, TOC, abstracts, and tail-block titles remain protected. A body run font-slot repair must skip or preserve those objects unless the transaction explicitly owns them.
- Final handoff after such a repair must include citation audit evidence proving superscript/hyperlink preservation and body-style or font-slot evidence proving the polluted visible body runs were corrected on the exact final DOCX.

### EXEC-MAINT-078. Footer PAGE Field Repair Must Match The Whole-Format Gate Size Contract (Mandatory)

- A footer PAGE-field repair helper must not repair page-number fields to a different direct run size than the exact-output whole-format gate expects.
- The `footer-page-number-font-size` operation in `scripts/repair_thesis_frontmatter_toc_structure.py` must follow this rule whenever it is selected.
- The canonical thesis footer PAGE-field size contract is 21 half-points (五号) unless an active official template/profile or current user instruction supplies a stronger explicit value.
- Do not normalize footer PAGE fields to 18 half-points unless an active official template/profile explicitly requires that value; 18 half-points is treated as a mismatch under the default thesis footer contract.
- Any repair helper that changes PAGE-field typography must expose the selected half-point size in its report, bind that value to the exact final DOCX, and be covered by a selftest that fails if the helper reintroduces previous 18 or 24 half-point defaults.
- `scripts/audit_docx_whole_format_gate.py` remains the release authority for footer PAGE-field size acceptance; a helper report alone cannot clear footer/page-number handoff.

### EXEC-MAINT-079. Tail-Block Titles Must Be Unlisted Explicit Title Surfaces (Mandatory)

- Reference, appendix, and acknowledgement title repair may not directly inherit `w:numPr`, `w:ilvl`, `w:numId`, outline state, first-line body indent, or bibliography-entry formatting from a converted school template.
- Appendix figure/table captions such as `图A.1` and `表A.1` are formal caption surfaces, not ordinary appendix body paragraphs. Self-check and repair helpers must not force those captions into first-line-indented body style, and table repair helpers must include appendix table titles when enforcing three-line-table structure.
- `参考文献`, `附录`, `致谢`, `谢辞`, `References`, `Appendix`, and `Acknowledgements` are terminal title surfaces. They must be represented as explicit title paragraphs with title-level style binding or direct title metrics and explicit title-size evidence before whole-format acceptance.
- If a template donor exposes a numbered title style, the canonical repair helper must strip list/numbering state after cloning donor metrics and then apply the accepted title baseline for the final manuscript. Do not pass the raw template numbering state through to the final document.
- Self-check baseline comparison must follow the same rule: a numbered template donor such as `附录多级编号2` is evidence of old donor numbering state, not an instruction to reject final unlisted `Heading1` tail titles that carry explicit 30 half-point title-size evidence.
- A whole-format gate or evidence helper must treat a numbered/list tail title as a tail-title style defect. It must not reinterpret that title as a reference entry, body heading, or appendix body paragraph to make another detector pass.
- Selftests must cover both sides of the defect: the repair helper removes tail-title numbering while preserving explicit title size, sample self-check accepts final unlisted explicit-title surfaces while rejecting numbered/list final titles, and the whole-format gate stops reference-entry scanning at appendix/acknowledgement boundaries.

### EXEC-MAINT-080. Cover Identity Placeholder Lines Must Not Survive Cover Repair (Mandatory)

- Cover identity paragraphs such as `学生姓名：________________`, `学号：________________`, `班级：________________`, and `指导教师：________________` are unresolved template placeholders, not valid identity values.
- Cover repair must start from a real cover authority: the official school template cover, an accepted local sample cover, or a locked current-manuscript cover baseline that has already been recorded as the donor for this run. Record the cover provenance, donor path/SHA, or the locked baseline evidence before mutating cover identity fields, cover tables, cover spacing, or the cover section boundary.
- A cover repair may replay cover donor/baseline paragraph, run, table-cell, underline, spacing, and section metrics onto the cover surface. It must not paste body text, abstract paragraphs, TOC rows, reference entries, acknowledgement text, appendix rows, or other non-cover blocks into the cover and then style them to look like cover content.
- If no real template/sample donor or locked cover baseline is available, the cover lane is blocked for mutation. Do not invent a cover layout from generic body formatting, default Word styles, school-name guesses, or nearby front-matter paragraphs.
- If real identity values are unavailable, the canonical cover repair helper must remove the underline/square placeholder value and preserve the label-only line instead of inventing names, student numbers, class names, advisors, or `未填写` stand-ins.
- Cover placeholder cleanup must be bounded to the cover/front-matter identity zone before the abstract, must preserve paragraph/run formatting where possible, and must not rewrite body content, references, figures, tables, media, comments, relationships, or headers/footers.
- Cover placeholder cleanup must preserve the locked cover-only section behavior. It may not borrow a body/front-matter section break, header, footer, or page-number model merely because that makes pagination easier.
- Whole-thesis release remains blocked until `scripts/audit_docx_whole_format_gate.py` passes the cover placeholder contract on the exact final DOCX.
- Selftests must include a DOCX fixture where `scripts/repair_template_surface_baselines.py --surfaces cover` removes cover identity placeholder lines without creating invented identity values.

### EXEC-MAINT-081. Live TOC Fields Must Be Locked When Cache Rows Are Preserved (Mandatory)

- When a thesis final relies on reviewed TOC cache rows, especially front-matter rows such as `摘要` and `ABSTRACT`, the standard `TOC` field must exist and its begin field must carry `w:fldLock="true"` before handoff.
- A visible TOC, static dotted-leader rows, or `HYPERLINK` / `PAGEREF` rows cannot satisfy the live-TOC requirement when the current workflow is whole-thesis production or whole-thesis revision.
- `scripts/audit_docx_whole_format_gate.py --require-toc-field` must fail when the final DOCX has no standard `TOC` field or when any standard `TOC` field is not locked.
- `scripts/validate_skill_gate_record_gate.py` must independently inspect the exact final DOCX named by the acceptance record and reject pass-shaped records that rely on visible TOC wording while the live TOC is absent or unlocked.
- When the repair path uses `toc-field-lock`, `toc-frontmatter-cache-compact`, `frontmatter-title-outline-exclusion`, or `toc-frontmatter-cache-exclusion`, the final acceptance record must bind the repair report, live TOC field count, locked TOC field count, sample self-check, whole-format gate, and pagination evidence to the exact final DOCX SHA256.

### EXEC-MAINT-082. Machine Evidence JSON Must Survive Windows Default Readers (Mandatory)

- Thesis audit and repair helpers that write JSON evidence for final handoff must serialize non-ASCII paths, Chinese text, and template labels with JSON escapes (`ensure_ascii=True`) unless a stronger consumer contract explicitly requires raw UTF-8 text.
- UTF-8 without BOM remains the required file encoding for skill artifacts, but final machine evidence must also parse under Windows PowerShell `Get-Content | ConvertFrom-Json` default-reader behavior when possible.
- A pass-shaped JSON report that is valid only when the reader manually specifies UTF-8, or that fails strict JSON parsing after a normal Windows shell read because Chinese path text changed into invalid backslash escape sequences, cannot be the only final acceptance evidence.
- The canonical final-thesis evidence scripts for list pollution, whole-format gate, font color, and bibliography school requirements must emit ASCII-safe JSON report files and console summaries so repeated cover/abstract/TOC/reference-format repair runs do not leave evidence that looks passed but is not machine-recheckable.

### EXEC-MAINT-083. Pagination Or Page-Number Complaints Require Whole-Document Pagination Evidence (Mandatory)

- When the user reports that the thesis has no pagination, no page numbers, wrong page breaks, frequent pagination errors, tail-block pagination loss, or equivalent page-flow defects, the next repair cannot close from a PDF page count or a generic `officecli validate` pass.
- The format lane must inspect and record the DOCX section map, hard page-break and section-break owners, header/footer references, footer `PAGE` fields, front-matter/body page-number restart state, TOC/page sync, tail-block opener ownership, and rendered physical page map on the exact final DOCX/PDF.
- Pagination evidence JSON must be ASCII-safe and machine-parseable under default Windows readers. A pagination report that becomes invalid JSON after `Get-Content | ConvertFrom-Json` because paths or non-ASCII text were written without JSON escaping is stale evidence and cannot close a page-number complaint.
- If the school template expects visible page numbers, every page class where page numbers should appear must have a Word/WPS page-number field or template-approved equivalent; hand-typed page numbers and missing footer fields are hard failures.
- The footer PAGE field must also carry the template/profile font-size contract directly on the field runs; a rendered PDF page number that looks present is not enough if the DOCX field can refresh into a different size.
- The PAGE field direct run size is part of the evidence surface; footer repairs must audit the actual PAGE field direct run size rather than only surrounding paragraph or rendered PDF appearance.
- When the official school format file is a prose rule sheet rather than a finished thesis sample, the prose-specified section geometry overrides that rule sheet DOCX's incidental section margins. The override must be explicit in the pagination evidence, for example by naming the accepted margin profile, and it must not hide missing PAGE fields, wrong page-number restarts, missing rendered footer maps, or tail-block opener defects.
- If the school template expects body/tail blocks to start on fresh pages, the opener owner must be explicit and singular: one `w:pageBreakBefore`, one preceding page/section break, or another template-proven owner. Duplicate or missing opener ownership is a pagination failure even if the exported PDF happens to look acceptable once.
- Chapter-opener checks must recognize Chinese chapter headings such as `第1章 ...` as first-level body openers, while avoiding false positives from calculation/result sentences that merely start with large numbers such as `96 ...`. A body pagination repair that misses `第N章` headings or treats calculation enumerators as chapters is stale.
- Tail-block previous-content page mapping must be robust to rendered PDF line wrapping. The checker may use bounded fuzzy fragments from the exact preceding DOCX paragraph, but it must search only before the formal tail-block opener page; a hit elsewhere in the thesis cannot prove reference-section separation.
- Mechanical design final chapters may use title variants such as `结论与展望` or `结论与后续优化` when the chapter is an independent body chapter before references. The pre-submission detector should reject merged test/summary chapters, not force every discipline into one literal `总结与展望` title.
- Final acceptance for such a run must bind `inspect_docx_pagination_structure.py` evidence, full rendered page/footer evidence, and the whole-format gate to the same final DOCX/PDF SHA256 pair. A first/last-page sample, page count, or page-footer excerpt is not enough after a user says the whole document has no pagination.
- The rendered evidence must include a complete physical-page footer/page-number map, not only sample pages. If the rendered footer map is generated before a field refresh, PDF export, formula-number repair, bibliography repair, or section-break repair, it is stale and must be regenerated.

### EXEC-MAINT-084. Final Gate Records Must Use Recheckable UTF-8 Paths And Clean Role Aliases (Mandatory)

- Final acceptance records, skill-invocation locks, agent manifests, task cards, and evidence ledgers must store real Windows absolute paths, not mojibake strings copied from a lossy console, PowerShell default-reader output, or a previously corrupted record.
- Before a pass-shaped thesis handoff, the audit lane must re-open every path-bearing field that the validator will read, including final DOCX/PDF paths, template profile paths, whole-format/list-pollution/font-color reports, CAD package paths, CAD appendix binding reports, rendered page evidence, task-card paths, and the skill lock path.
- `active_template_profile_path` must point to an existing generated profile file; `none`, empty strings, stale template paths, or corrupted template filenames block handoff.
- Role aliases in records must be regenerated from the clean role map rather than copied from terminal output. The audit role must be recorded as `审核` in final gate records when validators require readable Chinese role names.
- A final record cannot carry old failure text such as `blocked`, `failed`, `unresolved`, or `rendered sample still has unresolved thesis-quality issues` in pass-critical verdict, status, summary, handoff, or blocker fields. If the issue has been fixed, regenerate the record from the current evidence; if it has not been fixed, keep the handoff blocked.
- When the user reports repeated reference-format, pagination, formula-number, cover, or CAD appendix failures, the next repair must explicitly audit the final record itself for path encoding, stale blocker text, and exact-output evidence binding before running `validate_skill_gate.py`.

### EXEC-MAINT-085. External Format Reports Are Gate Inputs, Not Reference Notes (Mandatory)

- When a user supplies a school, teacher, platform, or third-party thesis format-check report, the report must be parsed into a normalized issue ledger before DOCX mutation or handoff, and the run must complete report-equivalent final DOCX verification.
- The issue ledger must preserve both expanded paragraph-level rows and report/statistics-page structural rows. Structure order, TOC, header, footer, page-number, table, caption, reference-label, abstract, keyword, heading, and acknowledgement/title rows from the report become user-reported protected surfaces.
- The final DOCX must be checked against a report-equivalent verifier bound to the exact final DOCX path and SHA256. A whole-format gate, body-style audit, table audit, PDF export, or sampled screenshot cannot clear a report-driven repair if the external report ledger was never consumed.
- The parser must distinguish report overview error counts from expanded sub-issue rows so table-cell fan-out does not inflate or hide the official reported error total.
- Content reminders such as abstract word-count warnings must be recorded as manual/content decisions. A format repair helper must not silently rewrite abstract content merely to satisfy a length reminder.
- Final acceptance for a report-driven repair must bind the report zip/path, extracted report files, normalized issue ledger, report-equivalent audit path/verdict, final DOCX SHA256, and any deliberately out-of-scope content reminders.
- Validators or selftests must fail closed when a report-driven final record lacks a normalized issue ledger, lacks exact-output report-equivalent evidence, or treats stats-page structure/header/footer/page-number findings as advisory only.

### EXEC-MAINT-086. Project-Local Helper Scans Must Distinguish Active Runs From Historical Archives (Mandatory)

- When a thesis run has a locked current `.codex/graduation-project-builder/<run>` directory, the project-local helper preflight must record that active run directory and may treat sibling run directories under `.codex/graduation-project-builder` as historical archives for the current pass.
- Historical sibling archives may still be listed in diagnostic scans, but they must not block a clean canonical-source pass when the current run uses only canonical skill scripts plus adapter data.
- Generic archive packaging markers such as `zipfile` or `ZipFile` are not enough to classify a helper as DOCX/OOXML mutation. The scanner must require actual DOCX package surfaces such as `.docx`, `word/document.xml`, Word headers/footers, section/page XML, or equivalent OOXML target evidence before raising a DOCX rewrite risk.
- CAD, PDF, DXF, DWG, PNG, or delivery-ZIP helpers must still fail if they insert or rewrite thesis DOCX surfaces outside canonical helpers; the active-run filter does not exempt a current helper from real DOCX mutation risks.
- Final helper-preflight evidence for an active run must include the project root, active run directory when used, scanner command, risk count, and scanner exit status.

### EXEC-MAINT-087. Heading-Completion Passes Must Not Corrupt TOC Or Body Typography (Mandatory)

- A pass that adds missing second-level or third-level headings must be body-surface bounded. It must locate the real first body chapter from section/page evidence, live TOC field boundaries, and heading style ownership before inserting anything.
- Project-local helpers must not search the whole `word/document.xml` for heading anchors with unbounded substring matching. If an anchor also appears in the TOC, abstract, figure caption, table title, reference entry, acknowledgement, appendix, header, footer, or formula/table text, the helper must skip that anchor until a canonical body-only locator proves a unique body paragraph.
- A final proof for `uses三级标题` must count only body heading rows after the TOC/body boundary and must separately report any `Heading3`/body-heading style paragraph that appears inside the TOC block or live TOC field result. A global `Heading3` count is not acceptance evidence.
- TOC cache rows are not body headings. Every visible TOC row, including level-3 and level-4 rows, must have a TOC-owned style/metrics record, a right dotted tab leader, and a visible page number. Aggregate dotted-leader counts cannot clear individual missing-row failures.
- The TOC cache must be checked against the real body-heading map. Existing TOC rows with valid leaders/page numbers do not prove that every body level-1/2/3 heading, especially newly inserted `Heading3` rows, is represented in the TOC.
- Body typography pollution is a release blocker when direct run size, style size, paragraph metrics, or caption/table sibling formatting makes body prose visibly smaller, denser, or otherwise different from the locked template/body baseline. The body-style audit must fail both over-large and under-size direct typography drift.
- If a user reports TOC row defects, body font-size pollution, or abnormal blank/near-empty pages, the next handoff must bind exact-output whole-format, body-style, and whole-document pagination evidence to the final DOCX/PDF SHA256 pair. A formula audit, figure audit, heading count, `officecli validate`, PDF export success, or package hash cannot close those user-reported visual surfaces.
- When the exact final PDF is available, whole-format evidence must scan rendered PDF near-empty pages directly or bind a canonical whole-document pagination JSON that does so. A page that contains only a footer/page number or other tiny residue is not acceptable unless its physical page number is explicitly allowlisted with independent review evidence.

### EXEC-MAINT-088. Renderers Must Not Mutate The Visible Final DOCX (Mandatory)

- A thesis render, PDF export, field refresh, Word/WPS/LibreOffice/officecli conversion, or screenshot pass must never operate on the visible handoff DOCX path directly.
- Before any rendered review, copy the exact final DOCX to a disposable render-source path and render that copy. The transaction record must name the disposable render-source DOCX path and SHA256.
- The visible handoff DOCX SHA256 must be recorded immediately before and immediately after the render/export/refresh step. The two hashes must be identical and must match `final_docx_sha256`.
- If a renderer, field refresh, or export tool changes the visible handoff DOCX, earlier DOCX-bound evidence is stale. Do not "rebind" the mutated file as accepted; restart from the clean source or rerun the full transaction after explicitly classifying the renderer mutation.
- Final acceptance must include `visible_final_docx_renderer_mutation_verdict=pass`, `final_docx_hash_before_render`, `final_docx_hash_after_render`, `render_source_docx_path`, and `render_source_docx_disposable_copy_verdict`.
- `scripts/validate_thesis_mutation_transaction.py` must reject mutating thesis transactions whose rendered-review evidence lacks these fields, whose render-source path equals the final DOCX path, or whose before/after/final hashes differ.

### EXEC-MAINT-089. Body Visual Pollution Evidence Must Prove Effective Typography And TOC Preservation (Mandatory)

- When the user reports body text that "looks different", body font-size pollution, heading/body visual drift, or similar rendered-body defects, XML direct-format cleanup is not enough.
- Body repair evidence must prove the effective font chain, including paragraph style, style inheritance, direct run `w:rFonts`, East Asian and Western font slots, `w:sz`/`w:szCs`, bold/boldCs, paragraph spacing, indentation, and nearby normal-body donor comparison.
- The body gate must also include rendered comparison evidence for the exact affected page or paragraph family. A pass that only says direct 24 half-point size or direct bold was removed is a false pass when the rendered text still differs.
- `--strict-direct-visible-metrics` or equivalent strict visible metrics must be enabled for user-reported body typography repairs. A disabled strict-direct-visible-metrics status cannot close a user-reported body pollution complaint.
- Citation runs, bookmarks, hyperlinks, fields, drawings, TOC rows, captions, tables, references, headers, and footers remain protected while body typography is repaired.
- If TOC is a protected sibling during such a repair, the transaction must prove the TOC field/cache rows, hyperlinks/bookmarks, page-number column, and rendered TOC/body heading sync were unchanged or intentionally refreshed under TOC ownership.
- When the user reports TOC digits, cached rows, page-number columns, bookmarks, or hyperlink anchors were damaged, final acceptance must bind a TOC digits/cache preservation evidence record for the exact final DOCX SHA256 with passing digit-spacing, page-number-column, bookmark-anchor, hyperlink-anchor, and before/after-cache verdicts. A dummy evidence file or pass-shaped prose cannot close the complaint.
- `scripts/validate_thesis_mutation_transaction.py` must reject mutating body-format transactions whose `chapter_format_preservation_report` lacks passing `effective_font_slot_verdict`, `strict_direct_visible_metrics_verdict`, `rendered_body_typography_comparison_verdict`, and `neighbor_body_baseline_comparison_verdict`, or whose TOC-protected transaction lacks TOC field/cache preservation evidence.

### EXEC-MAINT-090. Protected-Surface Diff And Review/Citation Evidence Must Be Parsed (Mandatory)

- A thesis mutation transaction cannot pass from `post_mutation_surface_diff verdict: pass` or a path-only check. The transaction validator must parse canonical protected-surface JSON, verify the schema, source/final DOCX paths and SHA256 values, full canonical surface-id coverage, changed package parts, unauthorized non-target surface rows, style-bearing part drift, keyword run split verdict, review-artifact verdict, citation-run verdict, and stale-evidence verdict.
- The protected-surface diff record must name every canonical surface, including cover, declaration/title front matter, Chinese/English abstracts and keywords, TOC title/entries/leaders/page-number column, body headings/text, figure/table holders, body citation superscripts, review comments/change marks, references, acknowledgement, appendix, headers, footers, page numbers, and whole-document pagination.
- Review-artifact and citation-run transaction fields must be re-opened and recomputed against the exact transaction source and final DOCX files. Fake markdown containing only `result: pass`, stale report paths, or reports for a different final DOCX must block handoff.
- Citation superscript preservation must be source-relative and run-level. If a source body citation marker is superscript, the final occurrence cannot pass when it becomes plain text, loses its marker run, moves to an unrelated host paragraph, or keeps a pass-shaped stale citation report.
- Selftest fixtures for transaction positives must generate canonical protected-surface diff JSON plus real source-to-final review/citation reports, so the fixture cannot hide a validator downgrade back to path checks.

### EXEC-MAINT-091. Caption Baseline Repair And Audit Must Share One Contract (Mandatory)

- Formal figure-caption and table-caption spacing baselines are protected template surfaces. A repair helper, body-style audit, sample self-check, and whole-format gate must use the same accepted caption baseline for the current school/template profile.
- For the current IMUST mechanical-design baseline, `repair_template_surface_baselines.py` and `audit_docx_body_style.py` must agree on figure captions before `0` / after `240` / line `360`, and table captions before `240` / after `0` / line `360` with `keepNext`.
- Do not require direct zero-indent attributes on captions when the locked baseline leaves those attributes absent. Adding explicit zero-indent attributes is itself a style drift unless a stronger template donor proves they are required.
- A body-style audit that still expects an older generic caption pattern such as before `120` / after `120` / line `240` is stale evidence for this thesis profile and must be fixed before it can block or clear handoff.
- If a future school template has a different caption baseline, first record the template/profile evidence, then update both the repair helper and all caption audit/self-check validators together. Do not patch only the document or only one checker.

### EXEC-MAINT-092. Tail-Block Run Fonts Must Not Be Reported As Paragraph Drift (Mandatory)

- Tail-block body checks for references, acknowledgement, and appendix must keep paragraph metrics and run typography as separate comparison surfaces.
- If `sample_self_check.py` materializes first-run font slots into a paragraph signature for detection convenience, tail-block body paragraph baseline comparison must not report those run-derived slots as paragraph-level drift.
- Appendix body paragraphs may carry explicit Chinese run fonts such as `宋体` when the approved donor/run model allows that visible typography. The checker must compare those slots through the run-signature contract, not through paragraph metric keys such as `eastAsia`, `ascii`, `hAnsi`, `size`, or `bold`.
- A true tail-block failure is still blocking when paragraph properties drift, style binding drifts, or the run-signature contract itself drifts. This rule only prevents false failures caused by mixing run-derived font materialization into paragraph-baseline diffs.

### EXEC-MAINT-093. Section Page Setup Must Be Materialized Before Rendered Geometry Acceptance (Mandatory)

- Whole-thesis DOCX outputs must materialize page setup on every `w:sectPr`; empty section properties or sections missing `w:pgSz` / `w:pgMar` are hard failures even when Word or WPS appears to inherit a usable layout on screen.
- When the locked school template exposes `w:cols` and `w:docGrid`, final sections must also carry those nodes unless a template/profile decision explicitly records why a section is exempt.
- Reference-entry rendered geometry drift can be caused by document page-box fallback, not only by bibliography paragraph indentation. After a bibliography-label, reference-format, pagination, or section-boundary repair, compare section page setup against the active template before interpreting reference `x0` drift as a reference-paragraph issue.
- A final PDF that renders thesis pages as US Letter or any non-template page box fails even if DOCX paragraph metrics, font audits, and citation audits pass. The rendered PDF page box must match the template A4/profile page size within an explicit tolerance.
- Before final handoff after page-flow or reference-format complaints, run `scripts/audit_docx_section_page_setup.py` on the exact final DOCX, with the active template/reference DOCX and final PDF when available. The evidence must bind the final DOCX/PDF SHA256 values.
- Repair helpers that add section breaks or clone front/body/tail blocks must copy or materialize `pgSz`, `pgMar`, `cols`, and `docGrid` from the active template/profile instead of leaving blank `sectPr` nodes for the renderer to guess.

### EXEC-MAINT-094. Review-Artifact Detection Must Not Confuse Field Instructions With Tracked Changes (Mandatory)

- Transaction validators must distinguish real Word tracked-change elements from field-instruction elements. The string prefix `<w:ins` is not enough because `<w:instrText>` belongs to Word field instructions such as TOC or hyperlink fields, not to revision insertion markup.
- A DOCX must be classified as containing review artifacts only when it contains actual comment parts with comment records, comment anchors, or exact tracked-change tags such as `<w:ins>` / `<w:del>` with a tag boundary. Field instruction text, TOC field caches, PAGE fields, and hyperlink field instructions must not trigger comment-resolution ledger requirements by themselves.
- A transaction that protects the `review_comments_and_change_marks` surface but whose source and final DOCX contain no comments and no real tracked-change tags may rely on source/final review-artifact inventory and diff evidence; it must not be blocked solely because the protected-surface list contains the word `comments`.
- If a run explicitly claims teacher-comment repair, all-comments resolution, or a non-not-applicable comment scope, the ledger gate remains mandatory even when the DOCX currently has no visible open comments.

### EXEC-MAINT-095. Durable Rules Must Be Wired Into Workflow Evidence Before Handoff (Mandatory)

- A new or repaired durable rule cannot remain as passive guidance. The same turn that adds or changes the rule must connect it to the execution chain that will enforce it.
- Required chain for each durable rule or user-reported defect class:
  - router exposure from `SKILL.md` or a directly loaded parent reference
  - one canonical owner file and `references/rule-owner-map.json` row
  - active checklist or task-card row naming the rule or defect class
  - agent run manifest entry naming the lane that owns execution and the audit lane that owns rejection
  - evidence field or evidence file path bound to the exact final artifact
  - validator or selftest owner when the check is automatable
  - manual rendered-review record when the check is visual or non-automatable
  - final acceptance or handoff text that states pass, fail, or explicitly skipped with reason
- Do not mark a final handoff as passed while the active checklist still has unchecked required rows, the manifest still says `pending` for the same surface, or the final evidence was generated against an older candidate.
- For thesis/CAD work, process-chain evidence must include the final DOCX/PDF/DWG/package path and SHA256 binding whenever the rule concerns formatting, figures, CAD borders, CAD overlap, CAD appendix binding, references, formulas, pagination, or media relationships.
- If a validator is unavailable, the workflow must fail closed to `audit-only` or bind a manual review record that names the exact rendered pages, CAD sheets, screenshots, or exported PDFs reviewed. A generic statement such as `checked visually` is not enough.
- When the user reports that rules exist but the flow did not follow them, the next skill-maintenance pass must audit the process chain itself before touching the thesis or drawing output. The audit must check router exposure, checklist rows, manifest fields, evidence freshness, final hash binding, and stale pass/failure text.
- A stale checklist can invalidate an otherwise repaired artifact for skill-controlled handoff. Before final packaging, either all required rows are checked with evidence or the unchecked rows are reclassified with explicit `not-applicable`, `superseded`, or `blocked` reasons.

### EXEC-MAINT-096. Flowchart And Table Rules Must Close Through Manifest, Validator, And Acceptance (Mandatory)

- When the user reports that flowchart, diagram, table, draw.io, generated-image, or table-format rules were bypassed, the next skill-maintenance or thesis run must wire the complaint through `SKILL.md`, `references/thesis/figure-rules/workflow-and-checklists.md`, `references/thesis-table-style-memory.md`, `scripts/thesis_figure_contract.py`, agent task cards, the run manifest, selftests, and final acceptance fields before claiming repair.
- Any thesis figure whose caption, title, nearby prose, manifest fields, teacher comment, or user comment contains workflow, process, step, chain, sequence, or flowchart semantics must be treated as a `flowchart` structural figure unless a stronger approved sample explicitly proves a different family. It cannot be recorded as a loose raster image, generic chart, AI-generated PNG, Mermaid final export, Pillow image, or unclassified screenshot.
- A `.drawio` file that only contains an imported bitmap/SVG image, `shape=image`, `data:image`, image URL cells, or pasted raster artwork is not a valid draw.io structural source. Native mxGraph shape vertices and boundary-bound orthogonal connectors are mandatory; wrapping a generated PNG inside draw.io is a hard failure.
- User-provided material such as `素材.doc` is provenance or style/source evidence only unless the current user explicitly locks `CORE-FIGURE-019` material-only reuse. In the default structural-figure path it does not replace the required draw.io source, SVG export, raster fallback, geometry report, rendered-page evidence, and final DOCX relationship binding for flowcharts and other structural figures.
- Every default-path flowchart must have a draw.io source file, SVG export, raster fallback, source-scale geometry validation report, source-to-inserted geometry evidence, post-insertion rendered-page evidence, final DOCX relationship evidence, visible start/end terminators, orthogonal boundary-to-boundary connectors, pass collision verdict, pass insertion status, and exact final DOCX binding. Missing any required field fails closed.
- Every material-only flowchart or structural figure must instead bind the primary material source, any approved supplemental source, missing-primary-source reason for supplemental use, extracted image SHA256, final embedded media SHA256, generated-substitute rejection verdict, and paragraph-width placement evidence. In that mode, requiring draw.io/native reconstruction is the bypass.
- Every real body table present in scope, generated, repaired, preserved, or detected in the final DOCX must have a table manifest row. A row containing only a caption, title, table number, or row count is not evidence and must fail validation.
- Each table manifest row must record active table authority lock, authority source type/path or pass no-template authority verdict, manuscript-binding proof, title/caption mode, border-family verdict, header separator verdict, vertical separator verdict, body-row separator verdict, table-local structure verdict, rendered table evidence, pagination/continuation verdict, insertion status, rendered-page status, final DOCX table evidence, and exact-output SHA binding.
- If the final DOCX contains any `w:tbl` body table or table caption, the acceptance gate must treat the table lane as touched even when a format task card forgot to mark `table`. The table evidence fields, table manifest path, and per-table inventory are then mandatory.
- The owner map entry for this rule must name validator and selftest owners. A skill-maintenance pass that only edits prose rules without updating `references/rule-owner-map.json`, templates, and validator/selftest coverage is incomplete.
