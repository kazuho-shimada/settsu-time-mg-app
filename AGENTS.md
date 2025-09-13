# language
日本語で簡潔かつ丁寧に回答してください

# Repository Guidelines
## Project Structure & Module Organization
- `src/` — Google Apps Script sources: `code.gs` (server), `index.html` and `admin.html` (templates), `appsscript.json` (manifest).
- `docs/` — operational notes and deployment tips (e.g., clasp commands).
- `config/` — reserved for environment/config files (no secrets committed).
- Keep GAS entry points (`doGet`, etc.) at top of `code.gs`, followed by feature blocks: initialization, utilities, cache, data access, UI helpers.

## Build, Deploy, and Local Development
- `clasp status` — show local vs. remote changes.
- `clasp push` — push local `src/` to Apps Script project.
- `clasp deploy --description "<note>"` — create a new deployment.
- `clasp open --webapp` — open current web app for manual testing.
- `clasp pull` — sync down remote changes if edits were made in the Apps Script editor.

## Coding Style & Naming Conventions
- JavaScript (V8). Indentation: 2 spaces; always use semicolons.
- Prefer `const`/`let` (no `var`). Functions in `camelCase`; constants in `UPPER_SNAKE_CASE` (e.g., `SPREADSHEET_ID`, `SHEETS`).
- Group related helpers; avoid large anonymous blocks. Keep functions short and single‑purpose.
- HTML templates: keep inline scripts minimal; prefer server utilities in `code.gs`.

## Testing Guidelines
- No automated tests yet. Perform manual checks via `clasp open --webapp`:
  - First‑load init, employee switching, check‑in/out flows, admin page navigation.
  - Cache behaviors (e.g., after updates). Use `console.log` for traces.
- When changing sheet schema, document headers and migration steps in the PR.

## Commit & Pull Request Guidelines
- Commits: concise, imperative mood; English or Japanese accepted. Optional prefixes: `Fix:`, `Refactor:`, `Feat:`.
- PRs must include: summary, affected screens (screenshots of `index`/`admin`), steps to verify, and any deployment or spreadsheet changes. Link related issues.

## Security & Configuration Tips
- Do not commit secrets. Prefer `PropertiesService` for IDs/keys over hard‑coding.
- Keep `appsscript.json` minimal; do not change `runtimeVersion` or `timeZone` without discussion.

## Agent‑Specific Notes
- Limit changes to `src/` unless otherwise requested. Avoid mass reformatting. Preserve existing public function names used by the web app.
