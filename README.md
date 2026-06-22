# graduation-project-builder--毕设论文一键化

`graduation-project-builder` 是面向毕设、毕业论文、课程设计和配套软件项目的 Codex skill。它把程序交付、论文撰写、格式修复、图表处理、验收记录和最终打包放在同一套流程里管理，目标是让项目可以运行、论文可以解释、证据可以复查、交付物可以验收。

## 前置依赖和配套 skills

使用本 skill 前，建议先确认当前 Codex 环境具备任务所需的前置能力。不是每个任务都需要全部前置，但涉及对应交付物时应提前安装或启用：

- 图表、流程图、ER 图、UML、BPMN、网络图等：建议配套 `drawio-diagrams-enhanced`，并准备 draw.io / diagrams.net；需要批量导出或 CI 导出时配套 `drawio-export-tools`。
- Word 论文、DOCX 模板、批注、目录、页眉页脚、引用和格式修复：建议配套 `Word / DOCX` 或 `officecli`，并准备可用于渲染检查的 Word、WPS 或 LibreOffice 环境。
- Excel、CSV、实验数据、统计表和明细表：建议配套 `Excel / XLSX`。
- Mermaid 架构图、流程图、时序图和 ERD：建议配套 `mermaid-diagram-specialist`。
- 论文文字降 AI 腔、中文润色和答辩稿自然化：中文优先配套 `humanizer-zh`，英文内容配套 `humanizer`。
- 计算机专业论文结构、摘要、相关工作、实验与结论写作：建议配套 `academic-writing-cs`。
- 需要生成或编辑位图素材、截图替代图、纹理或视觉资产时：建议配套 `imagegen`。
- 安装或更新 GitHub skill 时：建议配套 `skill-installer`。

如果缺少上述前置，Codex 应先说明缺口，再选择可执行的替代路径；涉及正式论文、图表、DOCX 或模板验收时，不应跳过对应前置检查。

## 建议 Codex 使用

建议在 Codex 中优先使用本 skill 处理毕业设计、毕业论文、课程设计、开题/中期/答辩材料，以及“系统 + 论文”一体化交付任务。调用时可以直接写：

```text
$graduation-project-builder 帮我完成这个毕设项目
```

或安装后让 Codex 在相关任务中默认调用 `graduation-project-builder`。使用时请保持以下个性化约束：

- 所有会话默认授权多 agent。
- 所有任务必须多 agent 并行执行。
- 每次任务至少保留一个审核 agent，专门检查内容是否符合 skill 内部样式和限制要求，以及是否符合用户提供模板的样式和限制要求。

## 适用任务

- 补全、调试、启动和打包毕业设计程序。
- 生成或修订毕业论文、课程设计文档和答辩材料。
- 按学校模板修复 DOCX 格式、目录、页眉页脚、图表、引用和参考文献。
- 将程序截图、结构图、流程图、表格、公式和验收证据绑定到最终交付物。
- 对最终输出进行清单化验收，减少漏项和样式漂移。

## Codex 个性化入口

Codex 展示和默认提示词位于：

```text
agents/openai.yaml
```

技能执行入口保持为：

```text
SKILL.md
```

不要把 `SKILL.md` frontmatter 中的 `name: graduation-project-builder` 改成中文名，否则会影响 `$graduation-project-builder` 的调用兼容性。中文尾缀用于 GitHub 展示名、README 标题、仓库描述和 Codex 展示文案。

## 安装

发布到 GitHub 后，可通过 skill installer 安装：

```text
$skill-installer install-skill-from-github.py --repo <owner>/<repo> --path <path/to/graduation-project-builder>
```

也可以使用 GitHub tree URL：

```text
$skill-installer install https://github.com/<owner>/<repo>/tree/<ref>/<path/to/graduation-project-builder>
```

安装后重启 Codex，让平台刷新可用 skills。
