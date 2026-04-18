import { APP_VERSION } from '../../../constants.js';

export default function AppFooter() {
  return (
    <footer className="app-footer">
      <span>Excel Previewer</span>
      <span className="footer-separator" aria-hidden="true">
        •
      </span>
      <span>
        Version <span id="footer-version">{APP_VERSION}</span>
      </span>
      <span className="footer-separator" aria-hidden="true">
        •
      </span>
      <span>Copyright &copy; SEANG SENGLY</span>
    </footer>
  );
}
