# graduation-project-builder--毕设论文一键化

这是从 Codex 同步到当前项目的 `graduation-project-builder` 完整副本，供当前 WorkBuddy 项目直接复用。

## 建议 Codex 使用

建议在 Codex 中优先使用本 skill 处理毕业设计、毕业论文、课程设计、开题/中期/答辩材料，以及“系统 + 论文”一体化交付任务。所有会话默认授权多 agent，所有任务必须多 agent 并行执行，并且每次任务至少保留一个审核 agent，专门检查内容是否符合 skill 内部样式和限制要求，以及是否符合用户提供模板的样式和限制要求。

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
