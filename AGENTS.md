# Repository Guidelines

## Project Structure

`qianchuan-material-collector.user.js` is the standalone Tampermonkey entry
point for Qianchuan material analysis and DouDian Compass self-video. Keep each
site integration in its own IIFE so hooks and state remain isolated. Tests
live in `tests/`. Root-level HAR captures and exports must stay out of Git.

The panels use pinned Vue 3, Element Plus, and SheetJS CDN resources declared
in the userscript metadata. Inject framework CSS only into each Shadow DOM;
never style the host commerce pages globally. Cross-page export snapshots use
Tampermonkey storage keys; store final rows only, never request credentials.

## Development and Testing

Run these commands from the repository root:

- `node --check qianchuan-material-collector.user.js` checks JavaScript syntax.
- `node --test tests/userscript.test.cjs` verifies replay parameters, response
  parsing, boundaries, dates, draggable UI geometry, and CSV/XLSX output.
- `git diff --check` detects whitespace errors before staging.
- `git status --short --branch` confirms the branch and intended changes.

For browser verification, install the script and reload each target page. On
Qianchuan, capture both promotion types and verify both filters. On Compass,
query Video Details once and confirm replay changes only dates, page number,
and descending `watch_cnt` sorting. A value at or below 20,000 ends a valid
descending scan; its CSV contains only `标题` and `观看次数`.
Check both `SUN` and `Dark` themes at desktop and narrow widths. Theme changes,
panel dragging, and date controls must not alter request-capture behavior.
Verify the combined Excel button stays disabled until both page snapshots use
the same dates, while each page's CSV remains independently downloadable.

## Coding Style and UI Copy

Use two-space indentation, single-quoted JavaScript strings, semicolons, and
descriptive camelCase names. Keep userscript metadata at the top. Preserve
captured request fields unless the feature explicitly requires changing them;
never log cookies, signatures, tokens, or full signed URLs.

Separate adjacent short UI labels, values, and statuses with one ordinary
space. When editing an existing page, replace middle dots (U+00B7) used as
separators with one ordinary space. Follow the project's existing conventions
for fonts, sizes, colors, buttons, and layout. Keep only copy and markers that
help users understand information, navigate, or take action.

## Commits and Pull Requests

Use a short imperative subject, such as `Add self-video collection`. Pull
requests should explain the behavior change, list checks run, link relevant
issues, and include screenshots for visible panel changes. Never commit HARs,
exports, account data, or credentials. Increase userscript `@version` for every
published behavior change.
