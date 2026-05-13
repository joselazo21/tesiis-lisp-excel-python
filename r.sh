#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

python3 render_ast_html.py

mkdir -p ast
cp -f AST_MANUAL.html "ast/AST_MANUAL.html"

cd ast
python3 -m http.server 8000 --bind 127.0.0.1
