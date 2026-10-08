#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass \
  -File "$(wslpath -w "$task_root/scripts/export_powerpoint.ps1")" \
  -SourcePath "$(wslpath -w "$task_root/deliverables/learning-fair-2026.pptx")" \
  -PdfPath "$(wslpath -w "$task_root/deliverables/learning-fair-2026.pdf")" \
  -PreviewDir "$(wslpath -w "$task_root/deliverables/preview")"
