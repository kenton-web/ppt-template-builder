#!/bin/bash
set -e

echo "📦 Installing pptx-deck-builder skill..."

# 1. Python dependencies
echo "→ Installing Python dependencies..."
pip install python-pptx pillow lxml -q

# 2. Skill files
SKILL_DIR="$HOME/.claude/skills/pptx-deck-builder"
mkdir -p "$SKILL_DIR"

BASE="https://raw.githubusercontent.com/kenton-web/ppt-template-builder/main/skills/pptx-deck-builder"

echo "→ Downloading skill files..."
curl -sL "$BASE/SKILL.md"      -o "$SKILL_DIR/SKILL.md"
curl -sL "$BASE/build_deck.py" -o "$SKILL_DIR/build_deck.py"

echo ""
echo "✅ Done! Restart Claude Code, then try:"
echo "   \"Use my template at ~/Downloads/MyTemplate.pptx to build a deck for [client]\""
