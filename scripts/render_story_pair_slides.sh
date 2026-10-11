#!/usr/bin/env bash
# Render the two new main decks and the separate tournament reveal states.
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

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

# Sequential COM calls keep the PowerPoint application from competing exports.
for package_name in story-v02 worldcup-v02; do
  package_dir="$task_root/deliverables/$package_name"
  render_deck "$package_dir/learning-fair-2026-$package_name.pptx" \
    "$package_dir/learning-fair-2026-$package_name.pdf" "$package_dir/preview"
done
choice_dir="$task_root/deliverables/worldcup-v02"
render_deck "$choice_dir/choice-states.pptx" "$choice_dir/choice-states.pdf" \
  "$choice_dir/choice-preview"
