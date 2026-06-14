# Thesis Table Style Memory

This file stores the active durable table-style authority for thesis work.

## Enforcement Status

- Every active table-style rule in this file is mandatory when table generation, redraw, border repair, caption repair, or table QA is in scope.
- If a stronger template-specific rule or an explicit current-user instruction exists, that higher-precedence source may override this file.
- Otherwise, this file is the single source of truth for thesis three-line tables under this skill.

## Active Table Authority

- The template-specific rendered sample or approved manuscript sample always overrides the generic table fallback in this file.
- When no stronger template-specific authority exists, the preserved generic thesis three-line-table fallback in this skill is the WPS-origin fallback family documented here.
- Do not keep, infer, or revive any older remembered three-line-table geometry from historical notes, screenshots, fallback samples, or prior project-specific rules.
- If a run cannot prove that it is following either the stronger template-specific authority or the locked fallback family in this file, the table lane is not ready.

## Authority Lock Rule

Before generating or repairing any thesis table, explicitly lock all of the following:

- exact manuscript path
- exact review-copy path
- active table authority
- authority source type
- authority source file path when a rendered sample is used
- manuscript-binding proof
- authority rationale

Do not let table repair proceed on memory alone without this lock.

## Editor-Side Authority

- When the current run uses a thesis three-line table and no stronger template-specific sample overrides it, the editor-side preferred authority is the WPS built-in `second three-line table` preset.
- Do not silently substitute:
  - a generic Word `TableGrid`
  - a generic officecli grid table
  - a manually remembered old three-line geometry
  - an older screenshot-derived table family
- If WPS automation can apply or structurally clone the preset safely, that path is preferred.
- If the live template/sample clearly shows a no-vertical-line three-line table, that visible template/sample geometry wins over any looser memory of the WPS preset.

## Rendered Acceptance Geometry

The accepted rendered border family for this WPS preset is:

- strong full-width top border
- strong separator line directly under the header row
- strong full-width bottom border
- no internal vertical separators between columns
- no body-row horizontal separators
- no outer left border
- no outer right border
- no visible per-cell or per-word underline on header text; header separation must come from the table/header border surface, not from run underline or paragraph-border residue

Treat this rendered geometry as the only accepted three-line-table family stored in this file.

## Title Mode And Caption Rule

- The active template, teacher-approved sample, or current user-identified table screenshot owns the table title mode.
- Valid title modes include an external standalone caption paragraph above the table and an in-table first merged title row.
- If the locked donor uses an in-table first merged title row, preserve that row inside the table, including its `gridSpan`, row/cell properties, paragraph properties, run properties, and the top-rule relationship around it.
- If the locked donor uses an external standalone caption paragraph, preserve that paragraph immediately above the table, centered, zero-indent, and keep-with-next.
- Do not convert between external-title and in-table-title modes unless the locked donor proves the current table is in the wrong mode.
- The generic fallback in this file may not override a current template, approved sample, current user screenshot, or explicit user correction about title mode.
- Do not keep editorial notes, parenthetical build notes, or standalone `数据来源：...` companion paragraphs bound to the table block unless a higher-precedence source explicitly requires them.
- If a table spans onto a later page and the active template/sample uses continuation titles such as `续表 3-2`, every continuation page must also start with a standalone continuation-title paragraph above the carried table fragment.
- Continuation-title paragraphs must:
  - stay outside the table grid
  - remain centered
  - keep first-line indent explicitly zero
  - keep-with-next into the continued table fragment
- Do not silently treat a repeated header row as a substitute for the required `续表` title when the approved sample expects that title family.

## Cell Text Rule

- Table-cell paragraph formatting is independent from nearby body paragraphs.
- Required default cell state for this active table family:
  - first-line indent explicitly zero
  - left and right indent explicitly zero
  - spacing before and after explicitly zero unless a stronger sample proves otherwise
  - no inherited thesis-body indent residue
  - header cells centered and visually stronger
  - short numeric/typed fields centered
  - longer descriptive text kept readable rather than stretched

If table text still looks visually offset after paragraph-indent repair, inspect and repair table/cell margins as well.

## WPS Compatibility Rule

- Distinguish editor-side preset authority from rendered acceptance evidence.
- Selected-table pale guides in WPS are not final printable borders.
- Final acceptance must be judged from:
  - an unselected page view
  - or an exported PDF
- Do not claim `the WPS built-in preset was applied` when the run actually used a rendered-geometry fallback only.

## Fallback Rule

- If direct WPS COM write attempts for table preset application hang or fail, the run may use a narrow rendered-geometry fallback only after recording that the WPS preset lane is blocked for the pass.
- That fallback is accepted only when the visible result matches the rendered acceptance geometry in this file exactly.
- The fallback must not introduce any other remembered three-line-table family.

## Full-Manuscript QA Rule

- Final table QA must enumerate every real body table on the exact review-copy path.
- Required QA sequence:
  - record total table count
  - inspect structural border family of every real body table
  - inspect header row run properties, paragraph borders, and cell borders so a per-word underline is not mistaken for the header separator rule
  - render every page that contains a table
  - verify the visible result on those rendered pages
- If a table spans multiple rendered pages and the active sample/template uses a continuation-title family, render and verify every continuation page separately.
- On each continuation page, verify:
  - the visible continuation title such as `续表 3-2`
  - the continuation title is outside the grid
  - the continuation title does not inherit body-text first-line indent
- For every cross-page table review, record a machine-readable continuation evidence row with:
  - table id/title
  - first rendered page and continuation rendered pages
  - rendered page-image or crop paths for each page
  - continuation-title text and whether it is outside the table grid
  - continuation-title alignment, first-line indent, and `keepNext` verdict
  - header repeat verdict (`tblHeader` or donor-backed equivalent)
  - row split verdict (`cantSplit` or donor-backed equivalent)
- If no real body table spans pages, final QA must still record `no cross-page body tables detected` with the rendered table page list. Silence on continuation evidence is not acceptable.
- Short tables must not be treated as cross-page tables only because their header or body tokens reappear on later PDF pages. Unless DOCX/rendered evidence proves that the actual table object crosses a page boundary, tables below the active long-table threshold must record `no cross-page body tables detected` instead of requiring a `续表` title.
- A run fails if even one table remains on a Word default full-grid family, keeps visible vertical separators, keeps body-row horizontal separators, or otherwise deviates from the locked three-line border family.

## Table Manifest And Validator Gate

- Every real body table that is generated, repaired, preserved, detected in the final DOCX, or otherwise in scope must have a `tables` entry in the figure asset manifest or the active thesis surface manifest.
- A table entry with only `caption`, `title`, `table number`, or `rows` is a false pass. It must fail until the row includes table authority, rendered evidence, and final DOCX binding.
- Required per-table manifest fields:
  - active table authority lock
  - authority source type and source path, or a pass no-template authority verdict
  - manuscript-binding proof between the authority and this table
  - title/caption mode
  - border-family verdict
  - header separator verdict
  - vertical separator verdict
  - body-row separator verdict
  - table-local structure verdict
  - rendered table evidence path
  - pagination or continuation verdict
  - insertion status
  - rendered-page status
  - final DOCX table evidence or final DOCX relationship/binding evidence
  - final DOCX SHA256 binding in the acceptance record
- If the final DOCX contains `w:tbl` or a rendered table caption, the table lane is considered touched even when the format-repair task card forgot to mark table work.
- Final acceptance must name the table surface inventory path, table manifest path, table manifest contract verdict, per-table evidence rows, exact-output DOCX binding evidence, and final DOCX SHA256 binding.

## Active Rendered Evidence

- `references/visual-style-samples/tables/table-style-sample-03-20260419-page.png`

Use this file only as rendered acceptance evidence for the WPS preset geometry. It is not a replacement for the WPS preset identity itself.
