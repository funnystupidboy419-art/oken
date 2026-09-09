# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**通信表・指導要録 所見作成支援アプリ** — A web application that automatically generates student remarks for school communications and guidance records (Japanese primary/secondary school context).

The app combines teacher input data (committee roles, behaviors, observations, etc.) with a templated text database to generate natural, appropriate remarks in both polite (敬体/comunicación) and plain (常体/guidance record) styles, with automatic character count management.

## Key Architecture

### Data Flow
1. **Input**: Student data (attendance number only; no personal data) entered via 8-tab interface or Excel paste
2. **Storage**: All data in browser localStorage (key: `shoken-app-v1`)
3. **Generation**: `Generator.build()` and related functions combine template text with student data
4. **Output**: Remarks in 6 varieties (2 styles × 3 domains) reviewed in tab 8, exported as Excel/CSV

### Critical Architectural Decisions

**Template System** (`js/templates.js`)
- Text is organized by category (committee, behavior, moral, etc.) with long/short variants and styling syntax
- Syntax: `{敬体|常体}` for style variants, `{{key}}` for data substitution
- All template text passes through `Util.render()` before display, which handles variable substitution and style conversion
- **Recent constraint**: Avoid abstract language ("多面的に考え", "心が育つ"). Prioritize concrete, observable student behaviors ("様々な立場から考え", "実感し").

**Text Assembly Algorithm** (`js/generate.js`)
- `Generator.build()` assembles remarks by:
  1. Collecting candidate sentences by category (committee, kakari/class role, events, behavior marks, observations)
  2. Ordering by priority (from `settings.order`)
  3. **Ending collision avoidance**: `renderPartVaried()` uses `endingKey()` (last ~9 chars) to skip templates that would repeat the previous sentence ending, ensuring varied language
  4. Character limit enforcement: first swap long→short, then delete low-priority items
- Similar logic in `Generator.buildMoral()` and `Generator.buildSougou()` for 道徳 and 総合
- **Critical**: Always maintain the ending-collision check when modifying templates or adding new sentences

**Excel Paste Parsing** (`js/util.js`, `app.js`)
- Tab-or-comma detection; quoted field handling for embedded newlines
- Behavior mark normalization (◎ ○ A B @ o 〇 → ◎ ○)
- Column header matching for behavior grids and multi-item tables

**Style Conversion** (`js/util.js`, `Util.toPlain()` / `Util.render()`)
- Audio-suffix handling: ～た → ～た (no change needed for most), but verbs need five-grade conjugation awareness
- "そのまま使う" observations: mechanical conversion for guidance records; **must be reviewed by teacher**

### File Roles

| File | Purpose |
|------|---------|
| `index.html` | Page structure; all JS and CSS bundled inline |
| `css/style.css` | Styling; `:root` CSS tokens for light/dark themes; `prefers-color-scheme` + `data-theme` attribute |
| `js/app.js` | Tab navigation, grid/table input, paste handling, state management, localStorage sync |
| `js/util.js` | Character counting (including full-width), text rendering, style conversion, file I/O, HTML escaping |
| `js/templates.js` | All remark text; organized by category and mark level (◎ ○) or free input (observations, moral, etc.) |
| `js/generate.js` | Remark assembly (`Generator.build*()` methods), ending-collision avoidance, character trimming logic |
| `js/xlsx.js` | Excel workbook generation (SpreadsheetML, no external libs); 6 sheets (敬体/常体 × 3 domains + metadata) |

## Common Development Tasks

### Running Locally
```bash
# Option 1: Direct (beware: ブラウザによっては出力ファイル名が download になることがある)
open index.html

# Option 2: Local server (recommended for file export testing)
python3 -m http.server 8080
# then visit http://localhost:8080
```

### Modifying Remark Templates
**All template text lives in `js/templates.js`.** Edit category objects (e.g., `T.committee.long`, `T.behavior['◎']`, `T.moral.material.noNote.short`):
- Use `{敬体|常体}` to vary by style
- Use `{{key}}` for variable substitution (keys defined in generate.js part functions)
- **Add variability**: multiple candidate strings per mark level reduce repetition when combined with ending-collision logic
- **Check impact**: If adding/removing templates, verify the ending-collision avoidance still works (especially for long texts)

### Testing Remarks Generation
1. Tab 1 (設定): Set character limits, item order, behavior order
2. Tab 2–7: Enter sample data or paste from Excel
3. Tab 8: Review generated remarks in both styles; check "文字数の都合で未使用" to verify trimming is correct
4. Use browser DevTools Console: `console.log(state)` to inspect in-memory state; check localStorage after save

### Exporting to Excel
- Tab 8 → "エクセルに出力" generates a SpreadsheetML `.xlsx` file with 6 sheets
- **No external libraries**: `xlsx.js` constructs the XML manually
- Sheets: 通信表所見(敬体), 指導要録所見(常体), チェック一覧, 入力データ, 道徳, 総合

## Important Constraints & Caveats

1. **No external libraries**: App must remain pure JS + HTML + CSS. This is a feature (no npm, no build step, no network calls).
2. **No personal data**: Strictly manage the no-name-input constraint. Tab 8 warnings catch "〜さん" patterns; respect this boundary.
3. **Browser storage only**: localStorage limit is ~5–10MB per domain; school data rarely exceeds this.
4. **Ending collision avoidance is critical**: When sentences end identically (e.g., "～ていました。"), the app reads as repetitive or copy-pasted. The `endingKey()` → `renderPartVaried()` logic prevents this. **Never disable or weaken it.**
5. **Style conversion is mechanical**: "そのまま使う" observations use `Util.toPlain()` to convert observations from polite to plain. This is imperfect (especially with rare verb forms). Teachers must review before using in guidance records.
6. **Character counting includes full-width**: `Util.countChars()` counts full-width chars as 2, half-width as 1, matching typical school measurement.
7. **Template language priority**: Avoid abstract concepts ("多面的", "心が育つ", "意欲的に"); describe concrete, observable behaviors instead.

## Data Structure (localStorage)

State object under key `shoken-app-v1`:
```javascript
{
  settings: { year, grade, className, term, count, order, enabled, behaviorOrder, behaviorMax, behaviorPick, limitReport, limitYouroku, moralLimitReport, moralLimitYouroku, sougouLimitReport, sougouLimitYouroku, useClosing, moralIntro, moralClosing, sougouClosing, warnShort, keepEdited },
  students: { "1": { no, committee, committeeRole, committeeWork, kakari, kakariWork, events[], behavior{}, observations[], moral[], sougou[], manual{}, bumps{} }, ... },
  moral: [ { studentNo, material, value, note }, ... ],
  sougou: [ { studentNo, unit, theme, activity, result }, ... ]
}
```

## Recent Changes & Priorities

**Latest work (as of Sep 2026)**: Improved template language to prioritize concrete, observable student behaviors over abstract descriptors.
- Replaced "多面的に考え" with "様々な立場や見方から考え"
- Replaced "心が育つ" with "実感し" / "積み重ね"
- Replaced "意欲的に取り組み" with specific actions ("進んで参加し、繰り返し練習して…")

These changes align with the app's goal: remarks grounded in what teachers actually see students do, not vague psychological language.

## Branch Strategy

Development branch: `claude/shinkoku-auto-remarks-app-sbxwvw`
- All changes merge to main via PR when ready
- Commits describe what changed and why (e.g., "Replace vague expressions with observable behaviors in templates")
