#!/usr/bin/env bash
# Render only the balance package; preserve the two older packages.
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
package_dir="$task_root/deliverables/balance-v03"

render_deck() {
  local deck_path="$1"
  local pdf_path="$2"
  local preview_dir="$3"
  if [[ ! -f "$deck_path" ]]; then
    printf 'Build the source PPTX before rendering: %s\n' "$deck_path" >&2
    exit 1
  fi
  mkdir -p "$preview_dir"
  find "$preview_dir" -maxdepth 1 -type f -name 'slide-*.png' -delete
  powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass \
    -File "$(wslpath -w "$task_root/scripts/export_powerpoint.ps1")" \
    -SourcePath "$(wslpath -w "$deck_path")" \
    -PdfPath "$(wslpath -w "$pdf_path")" \
    -PreviewDir "$(wslpath -w "$preview_dir")"
}

render_deck "$package_dir/learning-fair-2026-balance-v03.pptx" \
  "$package_dir/learning-fair-2026-balance-v03.pdf" "$package_dir/preview"
render_deck "$package_dir/choice-states.pptx" \
  "$package_dir/choice-states.pdf" "$package_dir/choice-preview"
