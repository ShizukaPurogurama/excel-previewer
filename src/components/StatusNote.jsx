export default function StatusNote({ message, tone }) {
  const classes = ['status-note'];
  if (tone === 'error') classes.push('is-error');
  if (tone === 'success') classes.push('is-success');

  return (
    <p id="status-note" className={classes.join(' ')}>
      {message || 'Upload a workbook to begin.'}
    </p>
  );
}
