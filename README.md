# graduation-project-builder--毕设论文一键化

`graduation-project-builder` 是面向毕设、毕业论文、课程设计和配套软件项目的 Codex skill。它把程序交付、论文撰写、格式修复、图表处理、验收记录和最终打包放在同一套流程里管理，目标是让项目可以运行、论文可以解释、证据可以复查、交付物可以验收。

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
