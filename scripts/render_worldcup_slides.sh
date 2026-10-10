#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass \
  -File "$(wslpath -w "$task_root/scripts/export_powerpoint.ps1")" \
  -SourcePath "$(wslpath -w "$task_root/deliverables/worldcup/learning-fair-2026-worldcup.pptx")" \
  -PdfPath "$(wslpath -w "$task_root/deliverables/worldcup/learning-fair-2026-worldcup.pdf")" \
  -PreviewDir "$(wslpath -w "$task_root/deliverables/worldcup/preview")"
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass \
  -File "$(wslpath -w "$task_root/scripts/export_powerpoint.ps1")" \
  -SourcePath "$(wslpath -w "$task_root/deliverables/worldcup/choice-states.pptx")" \
  -PdfPath "$(wslpath -w "$task_root/deliverables/worldcup/choice-states.pdf")" \
  -PreviewDir "$(wslpath -w "$task_root/deliverables/worldcup/choice-preview")"
