# Repository Guidelines

## Project Structure

`qianchuan-material-collector.user.js` is the standalone Tampermonkey entry
point. One userscript supports both the Qianchuan material analysis page and
the DouDian Compass self-video page. Keep each site integration in its own
IIFE so request hooks, state, and panels remain isolated. Automated checks
live in `tests/`. Root-level HAR captures and generated CSV files are local
fixtures and must stay out of Git.

## Development and Testing

Run these commands from the repository root:

- `node --check qianchuan-material-collector.user.js` checks JavaScript syntax.
- `node --test tests/userscript.test.cjs` verifies replay parameters, response
  parsing, boundaries, dates, and CSV output.
- `git diff --check` detects whitespace errors before staging.
- `git status --short --branch` confirms the branch and intended changes.

For browser verification, install the script in Tampermonkey and reload the
target page before querying the list. On Qianchuan, capture both promotion
types and verify the spend and impression filters. On Compass, query Video
Details once, then verify that replay keeps captured filters while changing
only dates, page number, and descending `watch_cnt` sorting. Confirm that the
first row at or below 20,000 stops a valid descending scan and that the CSV
contains only `标题` and `观看次数`.

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
exports, account data, or credentials.
