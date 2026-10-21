#!/usr/bin/env bash
# 一键安装 chalco-env-deck（macOS / Linux）
#
#   curl -fsSL https://raw.githubusercontent.com/chasingstars-a/chalco-env-deck/main/install.sh | bash
#
# 可带参数：
#   ./install.sh --proxy http://127.0.0.1:7890 --dest ~/skills/chalco-env-deck
#
# 做的事：下载仓库压缩包 -> 解压 -> 复制到 ~/.dsh/skills/chalco-env-deck
#        （检测到 ~/.claude 时再复制一份到 ~/.claude/skills/chalco-env-deck）-> 校验文件。

set -euo pipefail

REPO="${REPO:-chasingstars-a/chalco-env-deck}"
BRANCH="${BRANCH:-main}"
PROXY=""
DEST=""
ZIP_FILE=""
SKIP_CLAUDE=0

while [ $# -gt 0 ]; do
  case "$1" in
    --proxy) PROXY="$2"; shift 2 ;;
    --dest) DEST="$2"; shift 2 ;;
    --zip) ZIP_FILE="$2"; shift 2 ;;
    --skip-claude) SKIP_CLAUDE=1; shift ;;
    -h|--help) sed -n '2,12p' "$0"; exit 0 ;;
    *) echo "未知参数：$1"; exit 1 ;;
  esac
done

step() { printf '\033[36m==> %s\033[0m\n' "$1"; }
ok()   { printf '    %s\n' "$1"; }
fail() { printf '\033[31m!!  %s\033[0m\n' "$1"; exit 1; }

TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

if [ -n "$ZIP_FILE" ]; then
  step "使用本地压缩包"
  ok "$ZIP_FILE"
  cp "$ZIP_FILE" "$TMP/skill.zip"
else
  step "下载技能包"
  URL="https://codeload.github.com/${REPO}/zip/refs/heads/${BRANCH}"
  ok "$URL"
  CURL_ARGS=(-fsSL --retry 3 -o "$TMP/skill.zip")
  [ -n "$PROXY" ] && CURL_ARGS+=(--proxy "$PROXY")
  curl "${CURL_ARGS[@]}" "$URL" || fail "下载失败，检查网络/代理后重试"
fi

step "解压"
unzip -q "$TMP/skill.zip" -d "$TMP" || fail "解压失败（需要 unzip）"
ROOT="$(find "$TMP" -maxdepth 1 -type d -name '*-*' | head -n 1)"
[ -n "${ROOT:-}" ] && [ -f "$ROOT/SKILL.md" ] || fail "压缩包里没有 SKILL.md，仓库结构可能已变更"

TARGETS=()
if [ -n "$DEST" ]; then
  TARGETS+=("$DEST")
else
  TARGETS+=("$HOME/.dsh/skills/chalco-env-deck")
  if [ "$SKIP_CLAUDE" -eq 0 ] && [ -d "$HOME/.claude" ]; then
    TARGETS+=("$HOME/.claude/skills/chalco-env-deck")
  fi
fi

for target in "${TARGETS[@]}"; do
  step "安装到 $target"
  rm -rf "$target"
  mkdir -p "$(dirname "$target")"
  cp -R "$ROOT" "$target"
  [ -f "$target/scripts/chalco_deck.py" ] || fail "安装后缺少 scripts/chalco_deck.py"
  ok "OK"
done

step "完成"
printf '\033[32m    技能目录：%s\n' "${TARGETS[0]}"
printf '    在 AI 对话里说「用 chalco-env-deck 做一份 XX 汇报」即可；\n'
printf '    或直接跑示例：python "%s/examples/build_demo.py"\033[0m\n' "${TARGETS[0]}"
