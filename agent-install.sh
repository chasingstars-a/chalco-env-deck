#!/usr/bin/env bash
# 面向 AI Agent 的零参数安装器（macOS / Linux）
#
#   curl -fsSL https://raw.githubusercontent.com/chasingstars-a/chalco-env-deck/main/agent-install.sh | bash
#   curl -fsSL .../agent-install.sh | bash -s -- --with-deps
#
# 行为：下载 -> 解压 -> 安装到已存在的 ~/.dsh 或 ~/.claude（都没有则建 ~/.dsh/skills）
#      -> 可选装依赖 -> import 自检 -> stdout 最后一行输出 `RESULT {json}`
# 退出码：0 成功 / 1 失败。日志写 stderr。

set -euo pipefail

REPO="${REPO:-chasingstars-a/chalco-env-deck}"
REF="${REF:-main}"
SKILL="chalco-env-deck"
VERSION="1.1.0"
DEST=""
ZIP_FILE=""
WITH_DEPS=0
KEEP_ASSETS=0
NO_DEPS=0
QUIET=0

while [ $# -gt 0 ]; do
  case "$1" in
    --dest) DEST="$2"; shift 2 ;;
    --zip) ZIP_FILE="$2"; shift 2 ;;
    --with-deps) shift ;;
    --keep-assets) KEEP_ASSETS=1; shift ;;
    --no-deps) NO_DEPS=1; shift ;;
    --quiet) QUIET=1; shift ;;
    --help|-h) sed -n '2,12p' "$0"; exit 0 ;;
    *) echo "unknown arg: $1" >&2; exit 1 ;;
  esac
done

log() { [ "$QUIET" -eq 1 ] || echo "  $*" >&2; }
result() { printf 'RESULT %s\n' "$1"; }
fail() {
  result "{\"ok\":false,\"skill\":\"$SKILL\",\"version\":\"$VERSION\",\"error\":\"$1\"}"
  exit 1
}

# ---------- 1. 目标 ----------
targets=()
if [ -n "$DEST" ]; then
  targets=("$DEST")
else
  [ -d "$HOME/.dsh" ] && targets+=("$HOME/.dsh/skills/$SKILL")
  [ -d "$HOME/.claude" ] && targets+=("$HOME/.claude/skills/$SKILL")
  [ ${#targets[@]} -eq 0 ] && targets+=("$HOME/.dsh/skills/$SKILL")
fi
log "目标：${targets[*]}"

# ---------- 2. 取包 ----------
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
if [ -n "$ZIP_FILE" ]; then
  log "使用本地 zip：$ZIP_FILE"
  cp "$ZIP_FILE" "$TMP/skill.zip"
else
  URL="https://codeload.github.com/${REPO}/zip/refs/heads/${REF}"
  log "下载：$URL"
  curl -fsSL --retry 3 -o "$TMP/skill.zip" "$URL" || fail "download_failed"
fi
[ -s "$TMP/skill.zip" ] || fail "empty_zip"

# ---------- 3. 解压 ----------
(cd "$TMP" && unzip -q skill.zip) || fail "unzip_failed"
ROOT="$(find "$TMP" -maxdepth 2 -name SKILL.md -print -quit | xargs -r dirname)"
[ -f "$ROOT/SKILL.md" ] || fail "no_skill_md"

# ---------- 4. 安装 ----------
installed=()
for t in "${targets[@]}"; do
  kept=""
  if [ "$KEEP_ASSETS" -eq 1 ] && [ -d "$t/assets" ]; then
    kept="$TMP/kept-assets"
    cp -R "$t/assets" "$kept"
  fi
  rm -rf "$t"
  mkdir -p "$(dirname "$t")"
  cp -R "$ROOT" "$t"
  if [ -n "$kept" ]; then
    rm -rf "$t/assets"
    cp -R "$kept" "$t/assets"
    log "保留原有 assets/"
  fi
  [ -f "$t/scripts/chalco_deck.py" ] || fail "incomplete_install"
  installed+=("$t")
  log "已安装：$t"
done
FILES=$(find "${installed[0]}" -type f | wc -l | tr -d ' ')

# ---------- 5. Python 与依赖 ----------
# 优先挑已装好 python-pptx + pillow 的解释器；都没有时用第一个可用的并自动补装（--no-deps 可关闭）
PY=""
PYVER=""
HAS_DEPS=0
if [ -n "${CHALCO_PYTHON:-}" ]; then
  if command -v "$CHALCO_PYTHON" >/dev/null 2>&1 || [ -x "$CHALCO_PYTHON" ]; then
    PY="$CHALCO_PYTHON"; PYVER="$("$PY" --version 2>&1)"
    "$PY" -c 'import pptx, PIL' >/dev/null 2>&1 && HAS_DEPS=1
  fi
fi
if [ -z "$PY" ]; then
  for cand in python3 python; do
    if command -v "$cand" >/dev/null 2>&1; then
      if "$cand" -c 'import pptx, PIL' >/dev/null 2>&1; then
        PY="$cand"; PYVER="$("$cand" --version 2>&1)"; HAS_DEPS=1; break
      fi
      if [ -z "$PY" ]; then PY="$cand"; PYVER="$("$cand" --version 2>&1)"; fi
    fi
  done
fi

DEPS=0
IMPORT_OK=false
if [ -n "$PY" ]; then
  log "Python：$PYVER（依赖：$([ "$HAS_DEPS" -eq 1 ] && echo 已就绪 || echo 缺失)）"
  if [ "$HAS_DEPS" -eq 0 ] && [ "$NO_DEPS" -eq 0 ]; then
    log "安装依赖 python-pptx / pillow"
    if "$PY" -m pip install --quiet --disable-pip-version-check "python-pptx>=1.0" "pillow>=9.0" >/dev/null 2>&1; then
      DEPS=1
    fi
  fi
  if (cd "${installed[0]}/scripts" && "$PY" -c "import chalco_deck;print(chalco_deck.SW,chalco_deck.SH)" 2>/dev/null | grep -q '13.333'); then
    IMPORT_OK=true
  fi
else
  log "未找到 python3，跳过依赖与自检"
fi

# ---------- 6. 结果 ----------
if [ "$IMPORT_OK" = true ] || [ -z "$PY" ]; then OK=true; else OK=false; fi
TARGETS_JSON=$(printf '"%s",' "${installed[@]}" | sed 's/,$//')
HINT=""
[ "$OK" = false ] && HINT=',"hint":"python3 -m pip install \"python-pptx>=1.0\" \"pillow>=9.0\" 后重跑"'
result "{\"ok\":$OK,\"skill\":\"$SKILL\",\"version\":\"$VERSION\",\"ref\":\"$REF\",\"targets\":[$TARGETS_JSON],\"files\":$FILES,\"python\":\"$PYVER\",\"deps_install\":$DEPS,\"import_ok\":$IMPORT_OK,\"next\":\"新开一个会话，技能列表里会出现 $SKILL\"$HINT}"
[ "$OK" = true ] || exit 1
