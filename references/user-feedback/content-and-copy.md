# User Feedback Persistence: Content And Copy

Use this file for durable corrections about thesis wording, frontend copy, user-facing descriptions, and narrative perspective.

## Enforcement Status

- Every numbered rule in this file is mandatory when this file is loaded for the current subtask.
- Apply these rules together with `references/user-feedback-persistence.md`.

## Content And Copy Rules

### FB-COPY-001 (legacy 64). User-Facing Thesis And Frontend Copy Must Use User Perspective (Mandatory)

- When drafting thesis prose that describes pages, functions, or usage flow, prefer user-perspective wording such as what the user can view, select, query, compare, export, or obtain.
- Do not fill sections with product-introduction phrasing that only says "the system supports", "the system can", or "the system adopts" without explaining the user-facing action, purpose, and result.
- For frontend page descriptions, write from the interaction path first: what the user sees on entry, what inputs are available, what result is returned, and what value the result provides.
- For thesis chapters, architecture and implementation sections may still describe technical structure, but page descriptions, functional descriptions, and testing descriptions must not read like marketing copy or a generic system overview.
- If the user points out that current wording is too system-centric, treat that feedback as a durable correction and rewrite the affected sections before further formatting work.

### FB-COPY-002 (legacy 65). User-Facing Program Copy Must Start From Immediate User Intent (Mandatory)

- On business pages, login pages, homepage banners, card intros, and section helper text, write to the user's immediate goal first:
  - what the user wants to find
  - what the user wants to compare
  - what the user wants to manage
  - what the user wants to watch, decide, or act on now
- Do not use user-facing copy to introduce the system, the project, the graduation design, the school, or the implementation background unless the page itself is explicitly a documentation or about page.
- If the user explicitly says they want to see a specific thing, answer that thing directly in UI copy and conversational follow-up instead of drifting back into system introduction.
- If a hero, banner, lead paragraph, or helper text still reads like "the system is..." or "this project is..." on a business surface, treat the copy as unfinished and rewrite it.

### FB-COPY-003 (legacy 66). Business Images Must Be Semantically Relevant To The Actual Subject (Mandatory)

- If a page, banner, card, poster, screenshot placeholder, or visual slot uses an image, the image must match the actual business subject of that page or module.
- For graduation-project program pages, do not use generic office, classroom, meeting-room, laptop, abstract workspace, or unrelated stock photos when the page is about a specific business subject.
- Prefer in strict order:
  - system-generated domain image
  - searched image clearly matching the domain subject
  - conservative domain-relevant fallback
- Treat irrelevant placeholder imagery as a rule violation even if the layout otherwise looks complete.

### FB-COPY-004 (legacy 67). Every Thesis Figure And Table Must Have A 200-Character Explanation Below It (Mandatory)

- For every figure and every table in the thesis body, add a dedicated explanatory paragraph near the figure or table block.
- The explanatory paragraph should normally be no less than 200 Chinese characters.
- The paragraph must explain what the reader should look at, what the key data or structure means, and why the figure or table matters to the current section.
- Do not count the caption itself, a one-sentence note, or a generic transition sentence as satisfying this rule.
- Default placement rule:
  - for figures, keep the caption immediately below the image with no explanatory paragraph inserted between the image and caption
  - place the explanatory paragraph before the image by default, unless the current user explicitly asks for another layout
  - for tables, keep the caption attached to the table block and place the explanatory paragraph before the caption by default
- If the user reports large cross-page white space around a figure/table and explicitly asks to fill it with text, the content lane may place or split the explanatory body prose after the caption or before the next figure/table block when that is the only way to repair page occupancy. The caption must remain attached to the figure/table, and the added prose must stay in the approved body-text class rather than caption, heading, or placeholder formatting.
- Figure/table-adjacent fill text must be real thesis content tied to the current subsection, project implementation, data flow, or design decision. Do not use filler, meta-evaluation of the paper, delivery-process narration, empty paragraphs, or generic praise to occupy the page.
- If a figure or table is inserted without a sufficient explanatory paragraph, treat the thesis output as unfinished and return to content repair before format-only work.
- If the explanatory paragraph causes the caption to separate from the figure or table, the layout is wrong even if the explanation text itself is adequate.

### FB-COPY-005 (legacy 68). The Thesis Implementation Chapter Must Explain How Features Are Implemented, Not Read Like A Test Evaluation (Mandatory)

- When revising a thesis implementation chapter, organize the main prose around implementation steps such as route entry, validation logic, service processing, data persistence, page rendering, and output flow.
- Do not let the implementation chapter be dominated by evaluation-style lead-ins that belong more naturally in the testing chapter.
- Figure-adjacent paragraphs inside the implementation chapter should explain how the page or module is realized and which backend logic it corresponds to, rather than using the figure mainly to argue that the system has passed evaluation.
- When one implementation subsection discusses multiple pages or modules with multiple screenshots, build the subsection as repeated local units:
  - page or module explanation paragraph
  - the corresponding screenshot
  - the corresponding caption
- Do not write all page explanations first and then group all screenshots at the end of the subsection unless the user explicitly asks for a gallery-style block.
- If the user points out that screenshots have been piled at the bottom of a subsection, treat that as a durable execution rule and rebuild the subsection into alternating paragraph-figure blocks before further formatting work.
- If the user explicitly says the implementation chapter sounds like testing or assessment, treat that as a durable correction and rewrite the affected chapter before further formatting work.

### FB-COPY-006 (legacy 69). Business-Surface Copy Must Lead With The User's Immediate Task, Not With Platform Introduction (Mandatory)

- On login pages, dashboards, module homepages, shared shells, sidebars, top banners, hero areas, helper text, and empty states, start with what the current user wants to see or do now.
- Treat the user as the real operator of the product, such as an administrator, analyst, reviewer, dispatcher, or manager for that domain, not as a student defending the project or a presenter walking through slides.
- Prefer wording such as:
  - what the user can check first
  - what the user can compare
  - what the user can decide
  - what action the user can take next
- Do not spend the first visible copy block on describing the platform, architecture, technology stack, project background, or module inventory unless the page is explicitly an about, README, or documentation page.
- If a shared layout shell is used across many pages, its copy must still obey the same user-intent rule; do not hide a system-introduction paragraph inside the shared shell and then reuse it everywhere.
- If the user says page content should focus on `what I want to see` rather than `what the system does`, treat that as a durable correction and rewrite every affected business surface before handoff.
- If the user says the copy should read like a normal usable project rather than a graduation-defense script, remove defense-stage framing such as `答辩`, `演示者`, `上台讲解`, or similar presenter-first wording from business surfaces before handoff.

### FB-COPY-007 (legacy 70). User-Perspective Copy Must Be Reviewed As A Release Gate For Program UI Work (Mandatory)

- For any `program-only` or program-phase run that creates or revises frontend copy, do not stop after functional verification alone.
- Before handoff, review every visible business page surface at minimum for:
  - whether the first heading or lead text answers the current user's immediate goal
  - whether shared shell copy avoids generic platform introduction
  - whether helper text explains what the user can input, view, or obtain
  - whether empty states and section intros are action-oriented rather than system-oriented
- If any business page still reads mainly like `the system supports`, `the platform is`, `this project uses`, or other system-introduction phrasing, the UI copy review fails and the run must continue.
- Treat this review as mandatory even when the layout, routing, APIs, and tests already pass.

### FB-COPY-008. Thesis Body Humanizer Must Remove Meta-Evaluation Voice (Mandatory)

- When a thesis run writes or rewrites Chinese body prose, the content lane must apply the `humanizer-zh` route after factual consistency is checked.
- The humanizer pass must explicitly remove or rewrite wording that evaluates the thesis, the defense, the reviewer, the delivery process, Word assembly, formatting closure, or project management process unless that wording is truly part of a required reflection or acknowledgement section.
- Body chapters should describe project facts: requirements, design choices, data structures, implementation paths, interface behavior, test scope, observed results, limitations, and future technical work.
- Do not leave paragraphs that read as if they are praising, grading, defending, or externally assessing the paper itself. Phrases such as `论文质量`, `评审`, `答辩展示`, `交付完整`, `Word 草稿组装`, `格式收口`, and `证据化交付` are high-risk signals and must be rewritten into thesis-native technical narration or removed.
- Meta-writing phrases such as `从论文组织方法上看`, `写法的优点`, `写作逻辑`, `对于毕业设计论文而言`, `论文评价`, `加分价值`, `综合能力`, and `质量点` are also high-risk in body chapters. Rewrite them into concrete system facts, module responsibilities, data flow, test scope, or result boundaries instead of explaining why the paper is well organized or valuable.
- Thesis-body cleanup after review-feedback repair must also scan for process-pollution signals introduced by the repair itself. High-risk terms include `本地原型`, `可运行`, `可展示`, `交付`, `验收`, `完成度`, `答辩`, `评委`, `证明了系统`, `界面截图不能`, `论文在解释`, `当前阶段`, `后续继续完善论文和系统`, `不需要完全重写`, `不再把人工生成文本作为研究数据来源`, and `人工编造`. These terms block handoff when they appear in thesis body prose as defense, workload explanation, reviewer response, delivery-process narration, or acceptance evidence rather than as quoted source feedback or audit notes outside the manuscript.
- Mechanical-design theses must also reject delivery-process narration such as `先完成PDF+DWG工程图纸包`, `再按图纸参数反推计算书和说明书内容`, `交付PDF与DWG源文件`, `已作为独立A0图纸交付`, `源文件包含PDF、DWG和DXF格式`, or equivalent wording that describes how the artifact was generated rather than the machine design itself. Rewrite it into thesis-native statements about parameter coordination, structural design, calculation verification, manufacturing requirements, and drawing consistency.
- Rewrite contaminated body paragraphs into neutral academic statements about data source, method boundary, sample scope, model behavior, observed metrics, or future technical work. Do not explain why the project is acceptable for a graduation design, why a simpler model was chosen to reduce workload, why the paper is easier to inspect, or why screenshots prove the system works.
- After any thesis body rewrite, the content lane must produce a touched-scope pollution scan that lists the searched terms, paragraph identifiers, remaining counts, and disposition. A nonzero remaining count is acceptable only when the evidence names the exact quoted feedback/reference/audit context and proves it is outside the final thesis body prose; otherwise the paragraph must be rewritten before handoff.
- The acceptance record must include `assets/humanizer-evidence-template.md`-style evidence for every processed paragraph group: before text, after text, paragraph identifier, target language, skill name `humanizer-zh`, and a verdict that no meta-evaluation voice remains in the touched scope.
- A thesis text rewrite with only a general note such as `已润色` or `已 humanize` is incomplete.

### FB-COPY-009. Design And Implementation Chapter Responsibilities Must Stay Separated (Mandatory)

- When a thesis has separate design and implementation chapters, design material must live in the design chapter and implementation material must live in the implementation chapter.
- Architecture, function-module design, data model, database design, interface boundary, permissions model, configuration/logging design, workflow diagrams, ER diagrams, and other "why/what is designed" content belong to the design chapter.
- Page screenshots, route handling, validation logic, service processing, data persistence, component rendering, runtime interaction, and other "how it was implemented" content belong to the implementation chapter.
- If a run moves design content out of an implementation chapter, it must also update heading numbers, static or live TOC entries, figure/table captions, in-text cross-references, chapter-start pagination evidence, and any testing/conclusion paragraphs that mention the old chapter responsibilities.
- The implementation chapter must not be used as a mixed `设计与实现` bucket when the current user or active outline requires a dedicated `系统设计` chapter followed by a dedicated `系统实现` chapter.
- Handoff is blocked if the heading tree, TOC, captions, or cross-references still imply the old mixed chapter responsibility after the restructure.

### FB-COPY-010. Chapter-Length Expansion Must Be Evidence-Led, Not Filler (Mandatory)

- If the user or teacher asks a thesis chapter to reach a minimum page count, expand it with concrete project evidence, not broad background padding.
- Suitable expansion material includes real dataset source, fixed dataset filename, model family, training script, prediction script, routes, services, database tables, model metrics, screenshots, and implementation boundaries.
- When a model or metric section is revised, state which model produced the metrics, which code produced them, where the metric artifact is stored, and how the prediction path uses the trained artifact.
- If the revision mentions a code block or figure, add an explanation that ties the code/figure to the actual repository flow instead of describing it generically.
- After expansion, regenerate the final document and verify the requested chapter length in rendered PDF pages.

### FB-COPY-011. Implementation Chapter Sections Must Match Teacher-Requested Module Boundaries (Mandatory)

- If teacher feedback says a chapter section should represent one specific module, keep that section focused on that module instead of mixing adjacent functions into it.
- If the teacher asks to add a new section for a specific feature, create a separate heading for that feature and move the relevant explanation there, then renumber or regenerate downstream section numbers through the document builder.
- For system implementation chapters, avoid combining analysis dashboards, prediction functions, model management, and delivery scripts under one overloaded subsection when the feedback asks for separate sections.
- After section restructuring, regenerate the TOC and final DOCX/PDF, then verify both the TOC entry and the body heading for the newly inserted section.
