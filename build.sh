#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTPUT_NAME="accessMenu.nvda-addon"

cd "$ROOT_DIR"

if ! command -v zip >/dev/null 2>&1; then
  echo "zip is required. Install it with: sudo apt-get install zip" >&2
  exit 1
fi

rm -f "$OUTPUT_NAME"
(
  cd "$ROOT_DIR/addon"
  zip -r "$ROOT_DIR/$OUTPUT_NAME" . >/dev/null
)

echo "Built $OUTPUT_NAME in $ROOT_DIR"
