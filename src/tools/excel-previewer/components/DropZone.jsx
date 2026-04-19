export default function DropZone({ isBusy, onPickFile }) {
  return (
    <div className="ep-dropzone" aria-label="Open a file to get started">
      <button
        type="button"
        className="ep-dropzone-pick"
        disabled={isBusy}
        onClick={onPickFile}
      >
        {isBusy ? 'Reading…' : 'Open a file'}
      </button>
      <p className="ep-dropzone-hint">
        or drag and drop anywhere&ensp;·&ensp;.xlsx&thinsp; .xls&thinsp; .xlsb&thinsp; .ods&thinsp; .csv
      </p>
    </div>
  );
}
