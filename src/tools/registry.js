import { lazy } from 'react';

export const TOOLS = [
  {
    id: 'excel-previewer',
    name: 'Excel Previewer',
    description:
      'Upload and inspect Excel, CSV, and ODS workbooks entirely in the browser. No server, no upload.',
    icon: '📊',
    path: '/tools/excel-previewer',
    component: lazy(() => import('./excel-previewer/index.jsx')),
  },
  {
    id: 'hyperlink-generator',
    name: 'Hyperlink Generator',
    description:
      'Bulk-generate Excel =HYPERLINK() formulas from a column of values. Pick a source column, set a base path, and download the result.',
    icon: '🔗',
    path: '/tools/hyperlink-generator',
    component: lazy(() => import('./hyperlink-generator/index.jsx')),
  },
  {
    id: 'conditional-formula-filler',
    name: 'Conditional Formula Filler',
    description:
      'Fill cells based on multi-column IF/AND conditions without writing formulas. Define lookup values, set condition rules, pick target columns, and export a filled workbook.',
    icon: '🔀',
    path: '/tools/conditional-formula-filler',
    component: lazy(() => import('./conditional-formula-filler/index.jsx')),
  },
  {
    id: 'coming-soon',
    name: 'More tools coming',
    description: 'New browser-based tools will appear here as they are added.',
    icon: '🔧',
    path: '/tools/coming-soon',
    component: lazy(() => import('./coming-soon/index.jsx')),
    comingSoon: true,
  },
];
