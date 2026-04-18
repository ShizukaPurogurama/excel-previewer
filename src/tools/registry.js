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
    id: 'coming-soon',
    name: 'More tools coming',
    description: 'New browser-based tools will appear here as they are added.',
    icon: '🔧',
    path: '/tools/coming-soon',
    component: lazy(() => import('./coming-soon/index.jsx')),
    comingSoon: true,
  },
];
