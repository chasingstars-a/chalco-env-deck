# 给 AI Agent 的安装说明（chalco-env-deck）

你要把这个技能装到用户机器上，目标：`~/.dsh/skills/chalco-env-deck`（若存在 `~/.claude` 则再放一份）。
**不需要问用户任何问题**，按下面做，用最后一行 `RESULT {...}` 判断成功。

## 一键（推荐）

Windows：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -UseBasicParsing https://raw.githubusercontent.com/chasingstars-a/chalco-env-deck/main/agent-install.ps1 -OutFile $env:TEMP\chalco-ai.ps1; & $env:TEMP\chalco-ai.ps1"
```

macOS / Linux：

```bash
curl -fsSL https://raw.githubusercontent.com/chasingstars-a/chalco-env-deck/main/agent-install.sh | bash
```

脚本会：下载仓库 zip → 解压 → 复制到检测到的技能目录 → **发现缺 `python-pptx`/`pillow` 就自动 `pip install`**（可用 `-NoDeps` / `--no-deps` 关掉）→ import 自检 → 输出：

```
RESULT {"ok":true,"skill":"chalco-env-deck","version":"1.1.0","targets":["C:\\Users\\you\\.dsh\\skills\\chalco-env-deck"],"files":39,"python":"Python 3.12.14","deps_install":false,"import_ok":true,"next":"..."}
```

## 成功 / 失败判据

| 结果 | 含义 | 下一步 |
|---|---|---|
| 退出码 0，`ok:true`、`import_ok:true` | 装好了 | 告诉用户重开会话；可选跑一次 demo 自检 |
| 退出码 0，`ok:true`、`python:null` | 机器没有 Python，只装了文件 | 告知用户需装 Python 3.9+ 才能生成 PPT |
| 退出码 1，`ok:false` | 失败，看 `error` / `hint` | 按 hint 处理后重跑；网络问题改用离线 zip |

常见 `error`：`下载失败（网络或代理）`、`压缩包里没有 SKILL.md`、`安装不完整`、`依赖未就绪`。

## 下载回退

Windows 按 Invoke-WebRequest → python → curl.exe 依次尝试，POSIX 按 curl → wget → python3；公司网络里 schannel/代理常把第一条打掉，脚本会自动换路。全部失败时给 download_failed，改用 -ZipPath / --zip 离线包。

## 常用参数

| 需求 | PowerShell | POSIX |
|---|---|---|
| 指定安装目录 | `-Destination D:\skills\x` | `--dest ~/skills/x` |
| 离线/已有 zip | `-ZipPath .\repo.zip` | `--zip ./repo.zip` |
| 走代理 | `-Proxy http://127.0.0.1:7890` | `export HTTPS_PROXY=...` |
| 不动已有品牌素材 | `-KeepExistingAssets` | `--keep-assets` |
| 跳过 pip | `-NoDeps` | `--no-deps` |
| 指定解释器 | `$env:CHALCO_PYTHON="C:\py\python.exe"` | `CHALCO_PYTHON=/usr/bin/python3` |

## 手工兜底（脚本不可用时）

1. 下载 `https://codeload.github.com/chasingstars-a/chalco-env-deck/zip/refs/heads/main`
2. 解压，把里面的 `chalco-env-deck/` 整个目录放到 `~/.dsh/skills/chalco-env-deck`
3. 校验存在 `SKILL.md`、`scripts/chalco_deck.py`、`assets/logo_emblem.png`
4. `python -m pip install "python-pptx>=1.0" "pillow>=9.0"`
5. `cd <skill>/scripts && python -c "import chalco_deck;print(chalco_deck.SW,chalco_deck.SH)"` → 输出 `13.333 7.5`

## 装完自检（可选但推荐）

```bash
mkdir -p /tmp/chalco-check && cd /tmp/chalco-check
python <skill>/examples/build_demo.py
python <skill>/scripts/check_deck.py demo/*.pptx
```

期望最后一行：`结果：0 个错误，0 个警告`。

## 环境注意

- **Windows 执行策略**：只用 `-NoProfile -ExecutionPolicy Bypass -File` 单次调用，不要 `Set-ExecutionPolicy`。
- **`irm | iex` 与 `-File` 的区别**：前者失败时用 `throw` 报错，不会关掉用户窗口；后者退出码为 1。
- **Python 选择**：脚本优先选已装好依赖的解释器（`CHALCO_PYTHON` → `python` → `py -3` → `python3`）。
- **字体**：Windows 有微软雅黑；mac/Linux 没有会让版式偏移，必要时改 `scripts/chalco_deck.py` 顶部 `FONT`。
- **品牌素材**：公开仓库里是占位图，用 `-KeepExistingAssets` / `--keep-assets` 可保住用户已有的 `assets/`。

## 不要做

- 不要改 `SKILL.md` 的 `name: chalco-env-deck`（技能名必须一致）。
- 不要把用户的品牌素材提交回公开仓库。
- 不要用来源不明的图片替换 `assets/`。
