#!/usr/bin/env bash
# ติดตั้ง skills ทั้งหมดจาก repo นี้ไปไว้ที่ ~/.claude/skills (และ commands ไปที่ ~/.claude/commands)
#
# ใช้งาน:
#   ./install.sh                     # รันจากโฟลเดอร์ repo ที่ clone มาแล้ว
#   ./install.sh --skip-in-progress  # ไม่ติดตั้ง skill กลุ่มทดลอง (in-progress)
#   t=$(mktemp -d) && gh repo clone jatupornrmutsb/claude-skills "$t" -- --depth 1 --quiet && "$t/install.sh"; rm -rf "$t"

set -euo pipefail

REPO_URL='https://github.com/jatupornrmutsb/claude-skills.git'
DEST="${HOME}/.claude/skills"
CMD_DEST="${HOME}/.claude/commands"
SKIP_IN_PROGRESS=0

for arg in "$@"; do
    case "$arg" in
        --skip-in-progress) SKIP_IN_PROGRESS=1 ;;
        *) echo "ไม่รู้จักตัวเลือก: $arg" >&2; exit 1 ;;
    esac
done

# ── หาโฟลเดอร์ต้นทาง: ใช้ repo ที่สคริปต์อยู่ ถ้าไม่มีให้ clone ชั่วคราว ─────
SRC=''
TMP=''
if [ -n "${BASH_SOURCE[0]:-}" ] && [ -f "${BASH_SOURCE[0]}" ]; then
    SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
    [ -d "$SCRIPT_DIR/skills" ] && SRC="$SCRIPT_DIR"
fi
if [ -z "$SRC" ]; then
    TMP="$(mktemp -d)"
    trap 'rm -rf "$TMP"' EXIT
    echo "กำลัง clone $REPO_URL ..."
    git clone --depth 1 --quiet "$REPO_URL" "$TMP/repo"
    SRC="$TMP/repo"
fi

mkdir -p "$DEST" "$CMD_DEST"

# ── คัดลอก skill ทีละตัว (skills/<หมวด>/<ชื่อ>/ และ skill ของเราเองที่ root) ──
count=0
install_skill() {
    local dir="${1%/}" name
    name="$(basename "$dir")"
    rm -rf "${DEST:?}/$name"
    cp -R "$dir" "$DEST/$name"
    echo "  ติดตั้ง $name"
    count=$((count + 1))
}

for dir in "$SRC"/skills/*/*/; do
    [ -f "$dir/SKILL.md" ] || continue
    if [ "$SKIP_IN_PROGRESS" = 1 ] && [[ "$dir" == */skills/in-progress/* ]]; then
        continue
    fi
    install_skill "$dir"
done

for dir in "$SRC"/*/; do
    [ -f "$dir/SKILL.md" ] && install_skill "$dir"
done

# ── ติดตั้งแพ็กเกจ Python ที่ skill ต้องใช้ (ถ้ามี requirements.txt) ──
PY="$(command -v python3 || command -v python || true)"
for req in "$DEST"/*/requirements.txt; do
    [ -f "$req" ] || continue
    name="$(basename "$(dirname "$req")")"
    if [ -n "$PY" ]; then
        echo "  ติดตั้งแพ็กเกจ Python ของ $name"
        "$PY" -m pip install --quiet --user -r "$req" || echo "คำเตือน: ติดตั้งแพ็กเกจของ $name ไม่สำเร็จ" >&2
    else
        echo "คำเตือน: ไม่พบ Python: skill $name ต้องใช้แพ็กเกจใน $req" >&2
    fi
done

# ── คัดลอก slash commands ─────────────────────────────────────────────────
if [ -d "$SRC/commands" ]; then
    cp "$SRC"/commands/*.md "$CMD_DEST/" 2>/dev/null || true
fi

echo "เสร็จแล้ว ติดตั้ง $count skill ไปที่ $DEST"
echo "เปิด Claude Code ใหม่ (หรือเริ่ม session ใหม่) เพื่อให้เห็น skill"
