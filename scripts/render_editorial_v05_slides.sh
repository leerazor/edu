#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
deck_path="$task_root/deliverables/editorial-v05/learning-fair-2026-editorial-v05.pptx"
preview_dir="$task_root/deliverables/editorial-v05/preview"
if [[ ! -f "$deck_path" ]]; then
  printf 'Build the editorial v05 PPTX before rendering: %s\n' "$deck_path" >&2
  exit 1
fi
mkdir -p "$preview_dir"
find "$preview_dir" -maxdepth 1 -type f -name 'slide-*.png' -delete
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass \
  -File "$(wslpath -w "$task_root/scripts/export_powerpoint.ps1")" \
  -SourcePath "$(wslpath -w "$deck_path")" \
  -PdfPath "$(wslpath -w "$task_root/deliverables/editorial-v05/learning-fair-2026-editorial-v05.pdf")" \
  -PreviewDir "$(wslpath -w "$preview_dir")"
