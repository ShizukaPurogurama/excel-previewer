// Kept: clipboard write helpers (plain text + HTML) and HYPERLINK formula parser — no UI.
export function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function getClipboardCellText(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

export function getClipboardCellHtml(value) {
  const text = getClipboardCellText(value);
  const hyperMatch = text
    .trim()
    .match(/^=HYPERLINK\(\s*"([^"]+)"\s*(?:,\s*"([^"]+)")?\s*\)$/i);

  if (hyperMatch) {
    const url = escapeHtml(hyperMatch[1]);
    const label = escapeHtml(hyperMatch[2] || hyperMatch[1]);
    return '<a href="' + url + '">' + label + '</a>';
  }

  return escapeHtml(text);
}

export function buildClipboardTableHtml(rows) {
  return (
    '<table>' +
    rows
      .map(
        (row) =>
          '<tr>' +
          row
            .map(
              (value) =>
                '<td style="mso-number-format:\'\\@\'">' +
                getClipboardCellHtml(value) +
                '</td>'
            )
            .join('') +
          '</tr>'
      )
      .join('') +
    '</table>'
  );
}

export function copyHtmlToClipboard(plainText, htmlText) {
  if (navigator.clipboard && window.ClipboardItem) {
    const typeText = 'text/plain';
    const typeHtml = 'text/html';
    try {
      const item = new ClipboardItem({
        [typeText]: new Blob([plainText], { type: typeText }),
        [typeHtml]: new Blob([htmlText], { type: typeHtml }),
      });
      return navigator.clipboard.write([item]);
    } catch (err) {
      // fallthrough to execCommand path
    }
  }

  return new Promise((resolve, reject) => {
    try {
      const helper = document.createElement('textarea');
      helper.value = plainText;
      helper.setAttribute('readonly', '');
      helper.style.position = 'fixed';
      helper.style.opacity = '0';
      document.body.append(helper);
      helper.select();
      const success = document.execCommand('copy');
      helper.remove();

      if (!success) {
        reject(new Error('Copy command failed'));
        return;
      }

      resolve();
    } catch (error) {
      reject(error);
    }
  });
}

export function parseHyperlinkFormula(text) {
  if (typeof text !== 'string') return null;
  const match = text
    .trim()
    .match(/^=HYPERLINK\(\s*"([^"]+)"\s*(?:,\s*"([^"]+)")?\s*\)$/i);
  if (!match) return null;
  return { url: match[1], label: match[2] || match[1] };
}
