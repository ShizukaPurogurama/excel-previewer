export const PREVIEW_ROW_LIMIT = 24;
export const MAX_VISIBLE_COLUMNS = 10;
export const THEME_STORAGE_KEY = 'excel-viewer-theme';
export const COLUMN_STORAGE_KEY = 'excel-viewer-columns';
export const MONEY_MAPPING_STORAGE_KEY = 'excel-viewer-money-mappings';

export const CHANGELOG_ENTRIES = Object.freeze([
  Object.freeze({
    version: '1.3.0',
    dateLabel: 'March 17, 2026',
    stage: 'Current release',
    headline:
      "Cleaner money display, steadier column scrolling, and a built-in update page.",
    summary:
      'This release focuses on everyday workbook review. It makes payment amounts easier to read, keeps long column lists more stable while you browse them, and adds a plain-language place to understand what changed.',
    audience:
      'Operations, finance, QA, and support teammates who need quick spreadsheet review without technical release notes.',
    focus:
      'Clarity first: fewer surprises while scrolling, clearer amount fields, and easier communication for non-technical users.',
    highlights: [
      'Amount columns can automatically show $ for USD and ៛ for KHR when the sheet includes a matching currency column.',
      'The column picker keeps its place more reliably, reducing the jump-back effect in long lists.',
      "A new What's new page explains updates in simple language and shows the app version in a visible place.",
    ],
    impact: [
      'Payment and settlement files are easier to scan because the currency sign is visible at a glance.',
      'Long column lists feel more dependable during review and cleanup work.',
      'Anyone on the team can quickly see what improved without reading code or commit history.',
    ],
    notes: [
      'Your files still stay local in the browser unless you choose to reopen or reload them from your own device.',
      'The parsed table remains the main working view. The changelog page is read-only and there to keep updates easy to understand.',
    ],
  }),
]);

export const APP_VERSION = 'v' + CHANGELOG_ENTRIES[0].version;

export const COLUMN_PRESETS = {
  '*': [
    'TestID',
    'Service Name',
    'Sender Account',
    'Sender Amount',
    'Fee',
    'Receiver Account',
    'Receiver Amount',
    'TID',
    'Status',
  ],
};

export const SUPPORTED_CURRENCY_CODES = Object.freeze({
  USD: true,
  KHR: true,
});

export const FILE_PICKER_OPTIONS = {
  multiple: false,
  excludeAcceptAllOption: false,
  types: [
    {
      description: 'Excel and CSV files',
      accept: {
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
        'application/vnd.ms-excel': ['.xls', '.xlsb'],
        'application/vnd.oasis.opendocument.spreadsheet': ['.ods'],
        'text/csv': ['.csv'],
      },
    },
  ],
};

export const RELOAD_INTERVAL_OPTIONS = [
  { value: 1000, label: '1s' },
  { value: 3000, label: '3s' },
  { value: 5000, label: '5s' },
  { value: 10000, label: '10s' },
  { value: 30000, label: '30s' },
  { value: 60000, label: '1min' },
];

export const FILE_INPUT_ACCEPT = '.xlsx,.xls,.xlsb,.ods,.csv';
