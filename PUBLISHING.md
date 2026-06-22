# Publishing Notes

This folder is the GitHub-ready publishing copy of `graduation-project-builder`.

## 前置依赖和推荐配套 skills

发布说明开头必须提示使用者：本 skill 是总控交付 skill，正式运行前要按任务检查前置依赖和配套 skills。

- 结构图、流程图、ER 图、UML、BPMN、网络图等默认需要 draw.io / diagrams.net 和 `drawio-diagrams-enhanced`；批量导出、CI 导出、SVG/PNG 导出推荐配套 `drawio-export-tools`。
- 论文结构图不能默认用 Mermaid、Pillow、AI 生成 PNG 或贴图替代 draw.io 源文件，除非用户明确锁定材料复用或 draw.io 确实不适合。
- DOCX 论文、模板、目录、页眉页脚、批注、引用和格式修复推荐配套 `Word / DOCX` 或 `officecli`，并准备 Word、WPS 或 LibreOffice 渲染环境。
- Excel、CSV、实验数据和统计表推荐配套 `Excel / XLSX`。
- Mermaid 架构图和时序图推荐配套 `mermaid-diagram-specialist`。
- 中文论文润色、降 AI 腔、答辩稿自然化优先配套 `humanizer-zh`；英文内容配套 `humanizer`；计算机论文结构写作配套 `academic-writing-cs`。
- 位图素材生成或编辑配套 `imagegen`；从 GitHub 安装或更新本 skill 配套 `skill-installer`。

缺少前置时，Codex 应先说明缺口和替代路径；涉及正式论文、结构图、DOCX 或模板验收时，不应静默跳过检查。

## Recommended repo layout

You can publish it in either of these ways:

1. As a standalone repository:

```text
graduation-project-builder/
  SKILL.md
  agents/
  references/
  scripts/
```

2. As a subfolder inside a skill collection repository:

```text
skills/
  graduation-project-builder/
    SKILL.md
    agents/
    references/
    scripts/
```

## Installation after publishing

After uploading to GitHub, install it with the system skill installer using either:

```text
$skill-installer install-skill-from-github.py --repo <owner>/<repo> --path <path/to/graduation-project-builder>
```

or a GitHub tree URL:

```text
$skill-installer install https://github.com/<owner>/<repo>/tree/<ref>/<path/to/graduation-project-builder>
```

## Notes

- This publishing copy intentionally omits local-only mirror markers and cache folders.
- The bundle includes agent orchestration roles under `agents/`.
- The canonical maintained development copy remains local.
