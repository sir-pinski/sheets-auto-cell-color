/**
 * Google Sheets extension: auto-colors cells by content.
 */

const DEFAULT_PALETTE = {
  positive: '#d9ead3',
  negative: '#f4cccc',
  zero: '#fff2cc',
  text: '#cfe2f3',
  blank: '#ffffff',
};

/**
 * Adds custom menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cell Colors')
    .addItem('Color active range by content', 'colorActiveRangeByContent')
    .addItem('Clear colors in active range', 'clearActiveRangeColors')
    .addSeparator()
    .addItem('Set custom palette', 'promptForPalette')
    .addItem('Reset palette to defaults', 'resetPalette')
    .addToUi();
}

/**
 * Colors the currently active range using the configured palette.
 */
function colorActiveRangeByContent() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('No active range selected.');
    return;
  }

  colorRange_(range);
}

/**
 * Clears background colors in the active range.
 */
function clearActiveRangeColors() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('No active range selected.');
    return;
  }

  range.setBackground(null);
}

/**
 * Trigger-compatible edit handler.
 * Colors just the edited range for fast incremental updates.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  colorRange_(e.range);
}

/**
 * Prompts user to set custom color palette.
 */
function promptForPalette() {
  const ui = SpreadsheetApp.getUi();
  const current = getPalette_();

  const positive = askColor_(ui, 'Positive numbers (> 0)', current.positive);
  if (positive === null) return;

  const negative = askColor_(ui, 'Negative numbers (< 0)', current.negative);
  if (negative === null) return;

  const zero = askColor_(ui, 'Zero values (= 0)', current.zero);
  if (zero === null) return;

  const text = askColor_(ui, 'Text values', current.text);
  if (text === null) return;

  const blank = askColor_(ui, 'Blank cells', current.blank);
  if (blank === null) return;

  const palette = { positive, negative, zero, text, blank };
  savePalette_(palette);
  ui.alert('Palette saved.');
}

/**
 * Resets palette to defaults.
 */
function resetPalette() {
  savePalette_(DEFAULT_PALETTE);
  SpreadsheetApp.getUi().alert('Palette reset to defaults.');
}

/**
 * Colors every cell in a range according to content.
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function colorRange_(range) {
  const values = range.getValues();
  const palette = getPalette_();

  const backgrounds = values.map((row) =>
    row.map((value) => pickColorForValue_(value, palette))
  );

  range.setBackgrounds(backgrounds);
}

/**
 * Returns a color for an individual cell value.
 * @param {*} value
 * @param {{positive: string, negative: string, zero: string, text: string, blank: string}} palette
 * @returns {string}
 */
function pickColorForValue_(value, palette) {
  if (value === '' || value === null) {
    return palette.blank;
  }

  if (typeof value === 'number') {
    if (value > 0) return palette.positive;
    if (value < 0) return palette.negative;
    return palette.zero;
  }

  return palette.text;
}

/**
 * Gets persisted palette or default palette.
 * @returns {{positive: string, negative: string, zero: string, text: string, blank: string}}
 */
function getPalette_() {
  const raw = PropertiesService.getDocumentProperties().getProperty('CELL_COLOR_PALETTE');
  if (!raw) {
    return { ...DEFAULT_PALETTE };
  }

  try {
    const parsed = JSON.parse(raw);
    return {
      positive: normalizeColor_(parsed.positive, DEFAULT_PALETTE.positive),
      negative: normalizeColor_(parsed.negative, DEFAULT_PALETTE.negative),
      zero: normalizeColor_(parsed.zero, DEFAULT_PALETTE.zero),
      text: normalizeColor_(parsed.text, DEFAULT_PALETTE.text),
      blank: normalizeColor_(parsed.blank, DEFAULT_PALETTE.blank),
    };
  } catch (_error) {
    return { ...DEFAULT_PALETTE };
  }
}

/**
 * Saves palette.
 * @param {{positive: string, negative: string, zero: string, text: string, blank: string}} palette
 */
function savePalette_(palette) {
  PropertiesService.getDocumentProperties().setProperty(
    'CELL_COLOR_PALETTE',
    JSON.stringify(palette)
  );
}

/**
 * Asks user for a HEX color value.
 * @param {GoogleAppsScript.Base.Ui} ui
 * @param {string} label
 * @param {string} defaultColor
 * @returns {string|null}
 */
function askColor_(ui, label, defaultColor) {
  const response = ui.prompt(
    'Set color',
    `${label}\nEnter a HEX color like #aabbcc`,
    ui.ButtonSet.OK_CANCEL
  );

  const selectedButton = response.getSelectedButton();
  if (selectedButton !== ui.Button.OK) {
    return null;
  }

  const normalized = normalizeColor_(response.getResponseText(), defaultColor);
  if (!normalized) {
    ui.alert(`Invalid color for "${label}". Keeping ${defaultColor}.`);
    return defaultColor;
  }

  return normalized;
}

/**
 * Normalizes a color value to #rrggbb format.
 * @param {string} value
 * @param {string} fallback
 * @returns {string}
 */
function normalizeColor_(value, fallback) {
  if (!value || typeof value !== 'string') {
    return fallback;
  }

  const trimmed = value.trim();
  const shortHex = /^#([0-9a-fA-F]{3})$/;
  const longHex = /^#([0-9a-fA-F]{6})$/;

  if (longHex.test(trimmed)) {
    return trimmed.toLowerCase();
  }

  const shortMatch = trimmed.match(shortHex);
  if (shortMatch) {
    const [r, g, b] = shortMatch[1].split('');
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }

  return fallback;
}
