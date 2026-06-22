# graduation-project-builder--毕设论文一键化

这是从 Codex 同步到当前项目的 `graduation-project-builder` 完整副本，供当前 WorkBuddy 项目直接复用。

## 前置依赖和推荐配套 skills

使用本 skill 前，建议先确认当前 Codex 环境具备任务所需的前置能力。不是每个任务都需要全部前置，但涉及对应交付物时应提前安装或启用：

- 图表、流程图、ER 图、UML、BPMN、网络图等：默认需要 draw.io / diagrams.net 和 `drawio-diagrams-enhanced`；需要批量导出、CI 导出、SVG/PNG 导出时配套 `drawio-export-tools`。论文结构图不能默认用 Mermaid、Pillow、AI 生成 PNG 或贴图替代 draw.io 源文件，除非用户明确锁定材料复用或 draw.io 确实不适合。
- Word 论文、DOCX 模板、批注、目录、页眉页脚、引用和格式修复：建议配套 `Word / DOCX` 或 `officecli`，并准备可用于渲染检查的 Word、WPS 或 LibreOffice 环境。
- Excel、CSV、实验数据、统计表和明细表：建议配套 `Excel / XLSX`。
- Mermaid 架构图、流程图、时序图和 ERD：建议配套 `mermaid-diagram-specialist`。
- 论文文字降 AI 腔、中文润色和答辩稿自然化：中文优先配套 `humanizer-zh`，英文内容配套 `humanizer`。
- 计算机专业论文结构、摘要、相关工作、实验与结论写作：建议配套 `academic-writing-cs`。
- 需要生成或编辑位图素材、截图替代图、纹理或视觉资产时：建议配套 `imagegen`。
- 安装或更新 GitHub skill 时：建议配套 `skill-installer`。

如果缺少上述前置，Codex 应先说明缺口，再选择可执行的替代路径；涉及正式论文、结构图、DOCX 或模板验收时，不应静默跳过对应前置检查。

## 建议 Codex 使用

建议用户使用 Codex 进行任务处理，并在 Codex 中优先使用本 skill 处理毕业设计、毕业论文、课程设计、开题/中期/答辩材料，以及“系统 + 论文”一体化交付任务。

建议使用中转站降低成本，推荐中转站：https://api.apichatgpt.top

## Codex 个性化设置提示词

建议把以下内容添加到 Codex 的个性化设置内，作为使用本 skill 时的默认执行要求：

```text
所有会话默认授权多 agent，所有任务必须多 agent 并行执行，并且每次任务至少保留一个审核 agent，专门检查内容是否符合 skill 内部样式和限制要求，以及是否符合用户提供模板的样式和限制要求。
```

## 调用话术

首次完成毕设和论文时，按以下方式调用：

1. 先把干净的论文格式范文拖入聊天框。
2. 把任务书或者开题报告拖入聊天框。
3. 输入以下内容：

```text
$graduation-project-builder 多agent协同，根据格式模板和开题报告/任务书，先完成毕设，然后根据毕设1：1完成论文，注意需要有完整的论文画图以及系统截图，同时1：1和格式模板对齐格式。
```

如果是修改论文，请输入以下内容：

```text
$graduation-project-builder 多agent协同，根据格式模板和论文批注修改论文，注意不要破坏原有的论文格式。
```

建议配合目标模式使用，看情况可以同时叠加计划模式和目标模式。如果单用可能会导致修改出错。

## 一、建议优先阅读顺序

### 1. 核心入口

- `SKILL.md`
- `memory.md`

用途：

- `SKILL.md` 定义毕设主工作流、任务模式，以及程序与论文的默认执行顺序。
- `memory.md` 存放跨项目复用的稳定经验规则，适合用来避免重复踩坑。
- `agents/` 包含 agent orchestration roles，用于支持大任务拆分与独立审查。

### 2. 程序交付规则

- `references/program/workflow-standard.md`
- `references/program/verification-matrix.md`
- `references/program/stack-adaptation.md`
- `references/program/packaging-rules.md`
- `references/tooling-dependencies.md`
- `references/review-program-checklist.md`
- `references/review-delivery-bundle-checklist.md`
- `assets/program-gap-checklist.md`

用途：

- 判断项目是否仍停留在 demo 阶段
- 决定最小验证路径
- 决定交付包应包含什么
- 外化当前程序缺口清单
- 明确程序验证、打包与工具依赖
- 在交付前做显式 checklist 审查，减少漏项

### 3. 论文生产与格式修复

- `references/thesis/thesis-production-workflow.md`
- `references/thesis/thesis-format-rules.md`
- `references/thesis/thesis-format-sop.md`
- `references/thesis/thesis-format-class-review.md`
- `references/thesis/thesis-figure-generation-rules.md`
- `references/thesis/thesis-template-learning.md`
- `references/thesis/thesis-troubleshooting-log.md`
- `references/tooling-dependencies.md`
- `references/review-thesis-format-checklist.md`
- `references/review-thesis-content-consistency-checklist.md`
- `assets/thesis-blueprint-template.md`
- `assets/format-repair-task-template.md`
- `assets/final-acceptance-template.md`

用途：

- 真实项目到论文成稿的完整流程
- 格式修复 SOP
- 图表生成与验收规则
- 模板学习与故障回退经验
- 内容一致性与格式一致性的双重审查
- 在交付前做显式格式审查与最终验收汇总

### 4. 样式与视觉参考

- `references/thesis-layout-visual-memory.md`
- `references/thesis-figure-style-memory.md`
- `references/thesis-table-style-memory.md`
- `references/thesis-formula-style-memory.md`
- `references/visual-style-samples/STYLE-INDEX.md`
- `references/review-figure-style-checklist.md`
- `assets/figure-task-template.md`
- `assets/figure-plan-template.md`

用途：

- 作为论文图、表、公式、目录的默认视觉基线
- 当用户未提供更强样例时，优先参考这些文件
- 在绘图前锁定样式来源与任务边界
- 在插入前后做双重样式验收，减少样式漂移
- 先做单图 intake，再做多图计划，减少边画边猜

### 5. 用户反馈沉淀

- `references/user-feedback-persistence.md`

用途：

- 记录用户跨项目通用修正意见
- 防止有效反馈只停留在当前项目里

### 6. 当前采用的结构化模式

当前项目级 skill 已经引入 5 种结构化设计模式：

- Tool Wrapper：按需加载 references，不把所有规则堆进主 prompt
- Generator：通过 `assets/` 模板约束输出结构
- Reviewer：通过 `review-*.md` checklist 显式审查
- Inversion：高风险任务先锁定样式源、任务模式和边界
- Pipeline：程序、论文、图表任务按步骤执行，不能跳过 gate

## 二、当前项目建议保留的核心文件

### 必保留

- `SKILL.md`
- `memory.md`
- `references/`
- `scripts/`
- `agents/openai.yaml`
- `BASELINE-20260327.md`

### 可选保留

- `INSTALL.md`
- `MIRROR.md`
- `PUBLISHING.md`

这些文件不是核心执行规则，但保留有助于理解 skill 的镜像、安装与发布背景。

## 三、建议清理的冗余文件

以下文件主要是历史备份，不属于当前项目直接执行所需的核心资源：

- `SKILL.md.bak-20260323-221556`
- `SKILL.md.bak-clean-20260325-231050`
- `memory.md.bak-clean-20260325-231050`
- `memory.pre-clean-20260328.md`

如果你希望项目级 skill 更干净，可以删除这些备份文件。

## 四、当前项目使用建议

- 遇到“系统 + 论文”任务时，默认先看 `SKILL.md` 和 `memory.md`，再按需进入 `references/program/` 或 `references/thesis/`。
- 做论文格式修复时，不要只看主 `SKILL.md`，务必同时看 `thesis-format-rules.md`、`thesis-format-sop.md` 和 `thesis-format-class-review.md`。
- 做图表、结构图、ER 图、时序图时，优先看 `thesis-figure-generation-rules.md` 和 `STYLE-INDEX.md`。
- 接收到用户新的高价值修正意见后，应同步更新当前项目记忆或相关参考，而不是只留在一次对话里。
