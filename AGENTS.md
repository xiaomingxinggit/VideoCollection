# Repository Guidelines

## Project Structure

`qianchuan-material-collector.user.js` is the standalone Tampermonkey entry
point. It hooks Fetch and XMLHttpRequest on the Qianchuan material analysis
page and collects matching video materials from the live and product tabs.
Keep local HAR captures and generated exports out of Git.

## Development and Testing

Run `git status --short --branch` from the repository root to check your
branch and changed files. Run `git diff --check` before staging to catch
whitespace errors in tracked edits and `git diff --cached --check` to check
the staged changes.

Run `node --check qianchuan-material-collector.user.js` to check syntax. For
browser verification, install it in Tampermonkey, reload the analysis page,
capture both promotion types, choose dates, and start collection. Dates
default to the month containing yesterday
in Asia/Shanghai. After a recent page query is captured, the month shortcut
uses that query's cutoff date when it falls within the previous seven days.
On the first day of a month, the fallback is the full previous month. Check
that both requests use the selected `StartTime` and `EndTime`. Changing dates
resets the previous result.
When the latest page query uses different dates, the panel must explain the
difference before comparing row positions. Changing a captured business
filter invalidates old results, including changes during the final request.
DevTools should show 100-row replay requests sorted by descending impressions.
The live tab uses
`live_show_count_for_roi2_v2`. The product tab uses `product_show_count_for_roi2`.
The panel reports the first row below 20,000 impressions and stops after that
page while sorting remains descending. The CSV keeps only rows with at least
20,000 impressions and ¥1,000 spend. Missing or out-of-order
impressions fall back to reading all pages for that tab.

## Coding Style and UI Copy

Use two-space indentation, single-quoted JavaScript strings, and descriptive
camelCase names for functions and variables. Keep the userscript metadata at
the top of the file.

Separate adjacent short UI labels, values, and statuses with one ordinary
space. When editing an existing page, replace middle dots (U+00B7) used as
separators with one ordinary space. Follow the project's existing conventions
for fonts, sizes, colors, buttons, and layout. Keep only copy and markers that
help users understand information, navigate, or take action.

## Commits and Pull Requests

Use a short imperative commit subject, such as `Add contributor guidelines`.
In each pull request, summarize what changed and why, list the checks you ran,
link relevant issues, and include screenshots for visible UI changes.
