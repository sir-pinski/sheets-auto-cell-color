# Sheets Auto Cell Color

A Google Sheets Apps Script extension that colors cells based on their contents so you don't have to manually maintain conditional formatting rules.

## What this does

- Adds a **Cell Colors** menu in Google Sheets.
- Lets you **auto-color the active range** based on content type and value.
- Lets you **clear colors** from the active range.
- Lets you **set a custom palette** for positive, negative, zero, text, and blank cells.
- Includes an **onEdit trigger** that automatically colors edited cells.

## Color rules (default)

- **Number > 0**: light green (`#d9ead3`)
- **Number < 0**: light red (`#f4cccc`)
- **Number = 0**: light yellow (`#fff2cc`)
- **Text**: light blue (`#cfe2f3`)
- **Blank**: white (`#ffffff`)

## Setup

1. Open a Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Create files named `Code.gs` and `appsscript.json`.
4. Copy this repository's `Code.gs` and `appsscript.json` into the Apps Script project.
5. Save and reload the sheet.

## Usage

1. Select a range.
2. Click **Cell Colors → Color active range by content**.
3. Optionally configure colors via **Cell Colors → Set custom palette**.

For automatic updates when editing:

1. In Apps Script, click **Triggers**.
2. Add a trigger for `onEdit` (event source: spreadsheet, event type: on edit).

## Notes

- The script only changes background colors.
- It does not overwrite cell values or formulas.
- If you already use conditional formatting, this script can still run, but conditional formatting may visually override static background colors.
