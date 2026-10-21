# 一键安装 chalco-env-deck（Windows）
#
#   irm https://raw.githubusercontent.com/chasingstars-a/chalco-env-deck/main/install.ps1 | iex
#
# 或下载后按需带参数运行：
#   .\install.ps1 -Proxy http://127.0.0.1:7890 -Destination "D:\skills\chalco-env-deck"
#
# 做的事：下载仓库压缩包 -> 解压 -> 复制到 ~/.dsh/skills/chalco-env-deck
#        （检测到 ~/.claude 时再复制一份到 ~/.claude/skills/chalco-env-deck）-> 校验文件。
# 不写注册表、不装依赖、不联网上报。

[CmdletBinding()]
param(
    [string]$Repo = "chasingstars-a/chalco-env-deck",
    [string]$Branch = "main",
    [string]$Proxy = "",
    [string]$Destination = "",
    [string]$ZipPath = "",
    [switch]$SkipClaude
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Ok($msg) { Write-Host "    $msg" -ForegroundColor DarkGray }
function Fail($msg) { Write-Host "!!  $msg" -ForegroundColor Red; exit 1 }

$tmp = Join-Path $env:TEMP ("chalco-env-deck-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

try {
    $zip = Join-Path $tmp "skill.zip"
    if ($ZipPath) {
        Step "使用本地压缩包"
        Ok $ZipPath
        Copy-Item $ZipPath $zip -Force
    } else {
        $url = "https://codeload.github.com/$Repo/zip/refs/heads/$Branch"
        Step "下载技能包"
        Ok $url
        $params = @{ Uri = $url; OutFile = $zip; UseBasicParsing = $true }
        if ($Proxy) { $params.Proxy = $Proxy }
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest @params | Out-Null
    }
    if (-not (Test-Path $zip) -or (Get-Item $zip).Length -lt 10000) { Fail "下载失败，检查网络/代理后重试" }

    Step "解压"
    Expand-Archive -Path $zip -DestinationPath $tmp -Force
    $root = Get-ChildItem $tmp -Directory |
        Where-Object { Test-Path (Join-Path $_.FullName "SKILL.md") } |
        Select-Object -First 1
    if (-not $root) { Fail "压缩包里没有 SKILL.md，仓库结构可能已变更" }

    # 目标位置：DSH 技能目录（默认），可选 Claude Code 技能目录
    $targets = @()
    if ($Destination) {
        $targets += $Destination
    } else {
        $targets += (Join-Path $HOME ".dsh\skills\chalco-env-deck")
        if (-not $SkipClaude -and (Test-Path (Join-Path $HOME ".claude"))) {
            $targets += (Join-Path $HOME ".claude\skills\chalco-env-deck")
        }
    }

    foreach ($target in $targets) {
        Step "安装到 $target"
        if (Test-Path $target) {
            Ok "已存在，覆盖更新"
            Remove-Item $target -Recurse -Force
        }
        New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
        Copy-Item $root.FullName $target -Recurse -Force
        if (-not (Test-Path (Join-Path $target "scripts\chalco_deck.py"))) {
            Fail "安装后缺少 scripts/chalco_deck.py"
        }
        Ok "OK"
    }

    Step "完成"
    $skillDir = $targets[0]
    $demoScript = Join-Path $skillDir 'examples\build_demo.py'
    Write-Host "    技能目录：$skillDir" -ForegroundColor Green
    Write-Host '    在 AI 对话里说「用 chalco-env-deck 做一份 XX 汇报」即可；' -ForegroundColor Green
    Write-Host ('    或直接跑示例：python "{0}"' -f $demoScript) -ForegroundColor Green
}
finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
