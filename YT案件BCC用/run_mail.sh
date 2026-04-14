#!/bin/zsh
# ====================================================
# 定例会メール自動作成 Automator/ランチャー用シェルスクリプト
# ====================================================

SCRIPT_DIR="/Users/staff7/Documents/Antigravity/YT案件BCC用"
PYTHON_SCRIPT="$SCRIPT_DIR/periodic_meeting_mail.py"
PYTHON_BIN="/usr/bin/python3"

# ターミナルで実行（ログが見えるようにするため）
cd "$SCRIPT_DIR"
$PYTHON_BIN "$PYTHON_SCRIPT"
