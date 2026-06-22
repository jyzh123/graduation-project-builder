# Install Guide

After this skill is published to GitHub, install it inside Codex with one of these patterns:

## 安装前确认

本 skill 会调用多个前置能力和配套 skills。安装或启用前，请按任务类型确认：

- 结构图、流程图、ER 图、UML、BPMN、网络图等默认需要 draw.io / diagrams.net 和 `drawio-diagrams-enhanced`；批量导出、CI 导出、SVG/PNG 导出推荐 `drawio-export-tools`。
- 论文结构图不能默认用 Mermaid、Pillow、AI 生成 PNG 或贴图替代 draw.io 源文件，除非用户明确锁定材料复用或 draw.io 确实不适合。
- DOCX 论文、模板、目录、页眉页脚、批注、引用和格式修复推荐 `Word / DOCX` 或 `officecli`，并准备 Word、WPS 或 LibreOffice 渲染环境。
- Excel、CSV、实验数据和统计表推荐 `Excel / XLSX`。
- Mermaid 架构图和时序图推荐 `mermaid-diagram-specialist`。
- 中文论文润色、降 AI 腔、答辩稿自然化优先 `humanizer-zh`；英文内容用 `humanizer`；计算机论文结构写作用 `academic-writing-cs`。
- 位图素材生成或编辑用 `imagegen`；从 GitHub 安装或更新本 skill 用 `skill-installer`。

缺少前置时，Codex 应先说明缺口和替代路径；涉及正式论文、结构图、DOCX 或模板验收时，不应静默跳过检查。

## By repo and path

```text
$skill-installer install-skill-from-github.py --repo <owner>/<repo> --path <path/to/graduation-project-builder>
```

## By GitHub URL

```text
$skill-installer install https://github.com/<owner>/<repo>/tree/<ref>/<path/to/graduation-project-builder>
```

## After install

Restart Codex so the platform refreshes available skills.
