---
name: word-to-md-converter-portable
description: 通用版，自动查找 markitdown.exe，可在任意电脑上使用。支持批量将 Word (.docx) 文件转换为 Markdown (.md)，递归子文件夹并保留目录结构。找不到时可选全盘搜索，默认不执行。
version: 1.0.4
agent_created: true
---

# 角色与目标
你是一位专业的文档处理助手，具备批量将 Word (.docx) 文件转换为 Markdown (.md) 格式的能力，并能够智能地处理文件夹结构。

## 核心原则
1. **严格遵循用户提供的源路径和目标路径**。
2. **自动创建目标文件夹**（如果不存在）。
3. **自动查找 markitdown.exe**，优先搜索常见位置，找不到时询问用户是否进行全盘搜索（默认否）。
4. **默认只转换当前文件夹（不递归）**。如果用户明确要求"包含子文件夹"、"递归"或"包括所有子目录"，则启用递归模式。
5. **当启用递归时，必须在输出目录下保留源目录的完整相对路径结构**，避免所有 .md 文件混在同一层级导致重名覆盖。
6. **提供清晰的反馈**：报告找到的文件总数、成功数量、失败文件名及原因，并列出失败的详细列表。

## 执行脚本
当用户提出转换需求时，请根据用户输入的路径生成以下 PowerShell 脚本并执行。
脚本中的 `$sourceDir`、`$outputDir`、`$recurse` 需根据用户指令替换为实际值。

```powershell
# ===== 用户输入参数（请替换为实际值）=====
$sourceDir = "这里填用户给的源文件夹路径"
$outputDir = "这里填用户给的目标文件夹路径"
$recurse = $false   # 如果用户要求包含子文件夹，改为 $true
# =========================================

# ----- 自动查找 markitdown.exe（常见位置）-----
function Find-MarkitdownExe {
    # 构建候选路径列表，兼容脚本运行和直接运行两种场景
    $candidates = @()

    # 1. $PSScriptRoot 有效时（作为 .ps1 脚本运行），检查常见虚拟环境目录
    if ($PSScriptRoot) {
        $candidates += (Join-Path $PSScriptRoot ".venv\Scripts\markitdown.exe")
        $candidates += (Join-Path $PSScriptRoot "venv\Scripts\markitdown.exe")
        $candidates += (Join-Path $PSScriptRoot "markitdown\Scripts\markitdown.exe")
    }

    # 2. 用户目录及常见安装位置
    $candidates += "$env:USERPROFILE\markitdown\Scripts\markitdown.exe"
    $candidates += "F:\markitdown\Scripts\markitdown.exe"
    $candidates += "C:\markitdown\Scripts\markitdown.exe"

    # 3. 系统 PATH 中的 markitdown 命令
    $cmd = Get-Command markitdown -ErrorAction SilentlyContinue
    if ($cmd) { $candidates += $cmd.Source }

    # 依次检查候选路径
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    # 最后尝试：markitdown 在 PATH 中但 Source 属性为空的情况
    if (Get-Command markitdown -ErrorAction SilentlyContinue) {
        return "markitdown"
    }

    return $null
}

# ----- 全盘搜索 markitdown.exe（可选，耗时较长）-----
function Search-AllDrivesForMarkitdown {
    Write-Host "⚠️ 全盘搜索正在运行，可能需要数分钟，请耐心等待..."
    $drives = Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root
    foreach ($drive in $drives) {
        Write-Host "  搜索驱动器 $drive ..."
        try {
            $found = Get-ChildItem -Path $drive -Filter "markitdown.exe" -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.FullName -match "Scripts" } |
                    Select-Object -First 1
            if ($found) {
                Write-Host "  ✅ 找到: $($found.FullName)"
                return $found.FullName
            }
        } catch { }
    }
    return $null
}

# ===== 主流程 =====
$markitdownExe = Find-MarkitdownExe

# 常见位置未找到，询问用户是否全盘搜索
if (-not $markitdownExe) {
    Write-Host "❌ 未在常见位置找到 markitdown.exe"
    Write-Host "   常见搜索位置：虚拟环境 .venv/venv、USERPROFILE\markitdown、PATH"
    $response = Read-Host "是否进行全盘搜索？（y/N，默认否，按 Enter 跳过）"
    if ($response -eq 'y' -or $response -eq 'Y') {
        $markitdownExe = Search-AllDrivesForMarkitdown
    }
}

if (-not $markitdownExe) {
    Write-Host "❌ 未找到 markitdown.exe，请先安装：pip install 'markitdown[all]'"
    exit 1
}
Write-Host "✅ 使用 markitdown: $markitdownExe"
# -------------------------------------------------

# 1. 确保输出目录存在
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 2. 获取所有 .docx 文件
if ($recurse) {
    $files = Get-ChildItem -Path $sourceDir -Filter *.docx -Recurse
} else {
    $files = Get-ChildItem -Path $sourceDir -Filter *.docx
}

$total = $files.Count
if ($total -eq 0) {
    Write-Host "在 $sourceDir 中没有找到任何 .docx 文件"
    exit
}

Write-Host "找到 $total 个 Word 文件，开始转换..."

$success = 0
$failures = @()

foreach ($file in $files) {
    # 计算输出子目录（保留相对路径）
    if ($recurse) {
        $relativePath = $file.DirectoryName.Substring($sourceDir.Length).TrimStart("\")
        $targetDir = if ([string]::IsNullOrEmpty($relativePath)) { $outputDir } else { Join-Path $outputDir $relativePath }
    } else {
        $targetDir = $outputDir
    }

    # 确保目标子目录存在
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    $outputName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".md"
    $outputPath = Join-Path $targetDir $outputName

    try {
        $convertOutput = & "$markitdownExe" "$($file.FullName)" -o "$outputPath" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "转换成功: $($file.FullName) -> $outputPath"
            $success++
        } else {
            throw ($convertOutput -join "`n")
        }
    }
    catch {
        Write-Host "转换失败: $($file.Name) - 错误: $_"
        $failures += $file.FullName
    }
}

# 3. 输出汇总报告
Write-Host "`n========== 转换完成 =========="
Write-Host "共找到: $total 个文件"
Write-Host "成功: $success 个"
if ($failures.Count -gt 0) {
    Write-Host "失败: $($failures.Count) 个"
    Write-Host "失败文件列表:"
    foreach ($f in $failures) { Write-Host "  - $f" }
} else {
    Write-Host "全部成功！"
}
```
