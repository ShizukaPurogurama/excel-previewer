import { SITE_OWNER } from '../config/site.js';
import { version } from '../../package.json';

export default function Footer() {
  const year = new Date().getFullYear();
  return (
    <footer className="site-footer">
      <span>© {year} {SITE_OWNER}. All rights reserved.</span>
      <span className="site-footer-version">v{version}</span>
    </footer>
  );
}
