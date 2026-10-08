# Repository Guidelines

## Project Structure & Module Organization

This repository holds the 2026 Learning Fair video production package for the Data Center team, not an application. `docs/` contains production guidance, concept, facts, sources, storyboard, script, and handoff. `content/film.json` is the source for scene timing, dialogue, presenters, and slide cues. `assets/references/` contains attributed external reference images. `deliverables/` contains the editable PPTX, actual PowerPoint PDF/PNG renders, teleprompter, rehearsal CSV, and validation reports. `scripts/` contains Python generation/validation and the WSL/Windows PowerPoint export helper. Every new Markdown document should explain what its reader must do.

## Build, Test, and Development Commands

Run from the repository root. Install: `uv venv .venv` then `uv pip install --python .venv/bin/python -r requirements.txt`. Build: `.venv/bin/python scripts/build_package.py`. Render: `bash scripts/render_slides.sh` (WSL plus installed Windows Microsoft PowerPoint). Validate after rendering: `.venv/bin/python scripts/validate_package.py`. For other systems, export the deck to the documented PDF and 1600×900 PNG paths using PowerPoint before validation. See README for the standard venv/pip alternative.

## Coding Style & Naming Conventions

No formatter or linter is configured. Use Python with four-space indentation and snake_case. Use UTF-8 Korean production documents, kebab-case filenames, and source IDs P01–P10 for main slides, A01–A03 for hidden appendices, S01–S10 for scenes, IMG01–IMG14 for internal assets. External image IDs use IMG-01 etc. Maintain slide/cue/speaker/timing agreement across generated outputs. Rebuild rather than hand-edit generated storyboard, script, PPTX, teleprompter, and timing reports.

## Testing Guidelines

No general test framework or coverage target is configured. The focused validation command checks 300-second timing continuity, pronunciation-based estimates, scene cues, internal asset references, PPTX integrity/layout/notes, hidden appendices, actual PDF/PNG renders, and reference asset hashes. Inspect the contact sheet after rendering. Treat reading time as an estimate until rehearsal measurements exist; never substitute participant counts for performance gains. Add separate tests only for material new behavior that this validation does not cover.

## Commit & Pull Request Guidelines

Use concise Conventional Commits such as `feat: add learning fair production package`. No PR template exists. The user authorized committing and pushing all production outputs to `https://github.com/leerazor/edu` on `main`. Preserve existing remote history; never force-push. A future PR should explain production changes and include relevant preview images.

## Security & Configuration

Keep credentials and local secrets out of version control. Provide placeholder-only example configuration when environment variables are introduced, and document required variables without exposing their values.

Actual internal photos/screens are not yet supplied. Never present generic external images as this team's real facilities or events. Attribute external images with their license and source. October OutSystems is pending as of 2026-10-09 and scheduled before filming; update only after confirmation. IAP's log-pattern detection and formatted messenger broadcast are confirmed examples; expanding to all team infrastructure is a goal. Do not invent time savings, satisfaction, testimonials, root-cause analysis, or automatic incident recovery.
