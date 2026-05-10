Add README.md

# cui-tools —— 个人技能工具集

本仓库是我（cui）收集、编写和维护的 AI 助手技能集合。每个子文件夹都是一个独立的技能，可下载后直接使用。

## 包含技能

### 1. word-to-md-converter-portable

将 Word (.docx) 文件批量转换为 Markdown (.md)，支持递归子文件夹、保留目录结构，并自动查找 `markitdown.exe`。

- **特点**：自动适配不同电脑的安装路径，找不到时给出清晰提示
- **依赖**：需要先安装 `markitdown[all]`（见下文）
- **使用方法**：
  1. 将 `word-to-md-converter-portable` 文件夹复制到你的 AI 助手技能目录（例如：`C:\Users\你的用户名\.workbuddy\skills\`）
  2. 在 AI 助手中对话：`帮我把 D:\文档 里的 docx 全部转成 md，输出到 E:\输出`

## 依赖安装（必须）

本技能依赖微软开源工具 `markitdown`。在 **PowerShell** 中执行以下命令完成**完整安装**（支持 Word、PDF、PPT、图片、音频等）：

```powershell
pip install 'markitdown[all]' 
```
