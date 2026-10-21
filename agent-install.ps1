# 面向 AI Agent 的零参数安装器（Windows）
#
#   irm https://raw.githubusercontent.com/chasingstars-a/chalco-env-deck/main/agent-install.ps1 | iex
#   # 或非交互执行（推荐给 Agent）：
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\agent-install.ps1 -WithDeps
#
# 行为：下载 -> 解压 -> 安装到已存在的 ~/.dsh 或 ~/.claude（都没有则建 ~/.dsh/skills）
#      -> 可选装依赖 -> import 自检 -> 输出一行 `RESULT {json}`
# 失败用 throw（`irm | iex` 下也不会关掉用户的窗口），成功正常结束（退出码 0）。

[CmdletBinding()]
param(
    [string]$Repo = "chasingstars-a/chalco-env-deck",
    [string]$Ref = "main",
    [string]$Destination = "",
    [string]$ZipPath = "",
    [string]$Proxy = "",
    [switch]$WithDeps,
    [switch]$NoDeps,
    [switch]$KeepExistingAssets,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$skill = "chalco-env-deck"
$version = "1.1.0"

function Write-Log([string]$msg) {
    if (-not $Quiet) { Write-Host "  $msg" -ForegroundColor DarkGray }
}

function Stop-Install([string]$msg, $extra) {
    $r = [ordered]@{ ok = $false; skill = $script:skill; version = $script:version; error = $msg }
    if ($extra) { foreach ($k in $extra.Keys) { $r[$k] = $extra[$k] } }
    Write-Output ("RESULT " + ($r | ConvertTo-Json -Compress))
    throw "chalco-env-deck 安装失败：$msg"
}

function Invoke-Py {
    param([string[]]$PyCmd, [string[]]$PyArgs)
    $exe = $PyCmd[0]
    $pre = @()
    if ($PyCmd.Length -gt 1) { $pre = $PyCmd[1..($PyCmd.Length - 1)] }
    & $exe @pre @PyArgs
}

# ---------- 1. 目标目录 ----------
$targets = @()
if ($Destination) {
    $targets = @($Destination)
} else {
    if (Test-Path (Join-Path $HOME ".dsh")) { $targets += (Join-Path $HOME ".dsh\skills\$skill") }
    if (Test-Path (Join-Path $HOME ".claude")) { $targets += (Join-Path $HOME ".claude\skills\$skill") }
    if ($targets.Count -eq 0) { $targets += (Join-Path $HOME ".dsh\skills\$skill") }
}
Write-Log ("目标：" + ($targets -join " ; "))

$tmp = Join-Path $env:TEMP ("chalco-ai-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$zip = Join-Path $tmp "skill.zip"

try {
    # ---------- 2. 取包 ----------
    if ($ZipPath) {
        Write-Log "使用本地 zip：$ZipPath"
        if (-not (Test-Path $ZipPath)) {
            Stop-Install "本地压缩包不存在：$ZipPath" @{ hint = "检查 -ZipPath 路径" }
        }
        try { Copy-Item $ZipPath $zip -Force }
        catch { Stop-Install "无法读取本地压缩包：$ZipPath" @{ hint = "检查文件权限" } }
    } else {
        $url = "https://codeload.github.com/$Repo/zip/refs/heads/$Ref"
        Write-Log "下载：$url"
        $p = @{ Uri = $url; OutFile = $zip; UseBasicParsing = $true }
        if ($Proxy) { $p.Proxy = $Proxy } elseif ($env:HTTPS_PROXY) { $p.Proxy = $env:HTTPS_PROXY }
        try { Invoke-WebRequest @p | Out-Null }
        catch { Stop-Install "下载失败（网络或代理）" @{ hint = "用 -ZipPath 指定本地压缩包" } }
    }
    if (-not (Test-Path $zip) -or (Get-Item $zip).Length -lt 10000) {
        Stop-Install "压缩包异常（过小或不存在）" @{ hint = "用 -ZipPath 指定本地压缩包" }
    }

    # ---------- 3. 解压 ----------
    Expand-Archive -Path $zip -DestinationPath $tmp -Force
    $root = Get-ChildItem $tmp -Directory |
        Where-Object { Test-Path (Join-Path $_.FullName "SKILL.md") } |
        Select-Object -First 1
    if (-not $root) { Stop-Install "压缩包里没有 SKILL.md" }

    # ---------- 4. 安装 ----------
    $installed = @()
    foreach ($t in $targets) {
        $kept = $null
        if ($KeepExistingAssets -and (Test-Path (Join-Path $t "assets"))) {
            $kept = Join-Path $tmp ("assets-keep-" + [guid]::NewGuid().ToString("N").Substring(0, 6))
            Copy-Item (Join-Path $t "assets") $kept -Recurse -Force
        }
        if (Test-Path $t) { Remove-Item $t -Recurse -Force }
        New-Item -ItemType Directory -Force -Path (Split-Path $t) | Out-Null
        Copy-Item $root.FullName $t -Recurse -Force
        if ($kept) {
            Remove-Item (Join-Path $t "assets") -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item $kept (Join-Path $t "assets") -Recurse -Force
            Write-Log "保留原有 assets/"
        }
        if (-not (Test-Path (Join-Path $t "scripts\chalco_deck.py"))) {
            Stop-Install "安装不完整：缺 scripts/chalco_deck.py"
        }
        $installed += $t
        Write-Log "已安装：$t"
    }
    $files = (Get-ChildItem $installed[0] -Recurse -File | Measure-Object).Count

    # ---------- 5. Python 与依赖 ----------
    # 候选顺序：CHALCO_PYTHON 环境变量 -> python -> py -3 -> python3，
    # 优先挑"已经装好 python-pptx + pillow"的解释器，全都没有时用第一个可用的并自动补装依赖
    $candidates = @()
    if ($env:CHALCO_PYTHON) { $candidates += , @($env:CHALCO_PYTHON) }
    $candidates += , @("python")
    $candidates += , @("py", "-3")
    $candidates += , @("python3")

    $pyCmd = $null
    $pyVer = ""
    $hasDeps = $false
    foreach ($cand in $candidates) {
        try {
            $v = Invoke-Py -PyCmd $cand -PyArgs @("--version") 2>&1
        } catch { continue }
        if ($LASTEXITCODE -ne 0) { continue }
        try {
            $probeDeps = Invoke-Py -PyCmd $cand -PyArgs @("-c", "import pptx, PIL") 2>&1
            $okDeps = ($LASTEXITCODE -eq 0)
        } catch { $okDeps = $false }
        if ($okDeps) {
            $pyCmd = $cand; $pyVer = ($v | Out-String).Trim(); $hasDeps = $true; break
        }
        if (-not $pyCmd) { $pyCmd = $cand; $pyVer = ($v | Out-String).Trim() }
    }

    $depsInstalled = $false
    $importOk = $false
    if ($pyCmd) {
        Write-Log "Python：$pyVer（依赖：$(if ($hasDeps) { '已就绪' } else { '缺失' })）"
        if (-not $hasDeps -and -not $NoDeps) {
            Write-Log "安装依赖 python-pptx / pillow"
            try {
                Invoke-Py -PyCmd $pyCmd -PyArgs @("-m", "pip", "install", "--quiet",
                    "--disable-pip-version-check", "python-pptx>=1.0", "pillow>=9.0") | Out-Null
                $depsInstalled = ($LASTEXITCODE -eq 0)
            } catch { $depsInstalled = $false }
        }
        $probe = "import sys;sys.path.insert(0,r'$($installed[0])\scripts');import chalco_deck;print(chalco_deck.SW,chalco_deck.SH)"
        try {
            $res = Invoke-Py -PyCmd $pyCmd -PyArgs @("-c", $probe) 2>&1
            $importOk = ("$res" -match "13\.333")
            if (-not $importOk) { Write-Log ("import 自检失败：" + (("$res" -split "`n")[0])) }
        } catch { $importOk = $false }
    } else {
        Write-Log "未找到 python，跳过依赖与自检"
    }

    # ---------- 6. 结果 ----------
    $ok = $importOk -or (-not $pyCmd)
    $result = [ordered]@{
        ok            = $ok
        skill         = $skill
        version       = $version
        ref           = $Ref
        targets       = $installed
        files         = $files
        python        = $pyVer
        deps_install  = $depsInstalled
        import_ok     = $importOk
        next          = "新开一个会话，技能列表里会出现 $skill"
    }
    if (-not $ok) {
        $result["hint"] = 'python -m pip install "python-pptx>=1.0" "pillow>=9.0"，然后重跑 -WithDeps'
    }
    Write-Output ("RESULT " + ($result | ConvertTo-Json -Compress))
    if (-not $ok) { throw "依赖未就绪：import chalco_deck 失败" }
}
finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
