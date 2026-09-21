# 首咨选校方案

这是 YUSHI 用于生成“首咨选校方案”的可复用 Skill 源码包。

## 用途
输入学生背景 + 目标国家/地区 + 目标专业，输出固定格式的首咨选校规划 Word 文档（.docx）。

## 公开仓库原则
本仓库只保留匿名模板、规则、脚本和通用品牌素材，不上传真实学生 PDF 或学生个人数据。

## 目录
- `SKILL.md`：核心工作流
- `references/`：输入结构、核验规则、版式规范、QA 清单
- `assets/`：通用素材与模板（`template.docx` 为最终输出的版式基准；`template.pdf` 仅作为历史视觉参考）
- `scripts/`：确定性文档生成脚本
- `agents/openai.yaml`：Skill 显示信息

## 使用方式
- 在支持 Skills 的 ChatGPT/Codex 环境中安装或上传此 Skill 文件夹。
- 在 Codex 仓库中，也可放在 `.agents/skills/yushi-initial-consultation-school-selection/` 下。
- GitHub 主要用于版本管理与跨账号分发；具体安装方式取决于目标 ChatGPT 账号/工作区是否支持 Skills。
