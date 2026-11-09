/** The start screen: a drop zone with "Create new spreadsheet" and "Upload file". */
import { useState } from 'react';
import { Icon } from './Icon';
import type { UIState } from '@sumline/core/views';

export function StartScreen({ onNew, onOpen, onFile, restore, onRestore, onDiscard }: {
  onNew: () => void;
  onOpen: () => void;
  onFile: (file: File) => void;
  restore: UIState['restore'];
  onRestore: () => void;
  onDiscard: () => void;
}) {
  const [dragging, setDragging] = useState(false);
  return (
    <main className="sl-start" data-testid="start-screen">
      <div className="sl-start-head">
        <h1>Open and edit Excel spreadsheets in your browser</h1>
        <p>Open an XLSX, edit it the way you would in Excel and save it back. Formulas, charts, conditional formatting, data validation and multiple worksheets, on desktop and on your phone. Your file stays on your device.</p>
      </div>
      <div
        className={`sl-dropzone${dragging ? ' dragging' : ''}`}
        data-testid="dropzone"
        onDragOver={(e) => {
          if ([...e.dataTransfer.types].includes('Files')) {
            e.preventDefault();
            setDragging(true);
          }
        }}
        onDragLeave={(e) => {
          if (!e.currentTarget.contains(e.relatedTarget as Node)) setDragging(false);
        }}
        onDrop={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setDragging(false);
          const file = e.dataTransfer.files[0];
          if (file) onFile(file);
        }}
      >
        <div>
          <Icon name="upload_file" />
          <p><strong>Drop your Excel file here</strong></p>
          <p className="caption">XLSX · XLS · XLSM · XLSB · CSV</p>
          <div className="sl-start-actions">
            <button type="button" className="sl-btn primary" data-testid="start-new" onClick={onNew}>Create new spreadsheet</button>
            <button type="button" className="sl-btn" data-testid="start-upload" onClick={onOpen}>Upload file</button>
          </div>
        </div>
      </div>
      {restore && (
        <div className="sl-start-recent" role="alert" data-testid="restore-banner">
          <div className="sl-validation-card">
            <div>
              <div><b>Unsaved changes found</b> in “{restore.fileName}”</div>
              <div className="meta">{new Date(restore.savedAt).toLocaleString()}</div>
            </div>
            <div style={{ display: 'flex', gap: 6 }}>
              <button type="button" className="sl-button" data-testid="restore-recover" onClick={onRestore}>Recover</button>
              <button type="button" className="sl-button light" data-testid="restore-discard" onClick={onDiscard}>Discard</button>
            </div>
          </div>
        </div>
      )}
      <div className="sl-start-notes">
        <div>
          <h3>1. Upload or start fresh</h3>
          <p>Drag an XLSX, XLS, XLSM, XLSB or CSV file here, or start with a blank workbook.</p>
        </div>
        <div>
          <h3>2. Edit and format</h3>
          <p>Type formulas, sort and filter, add charts, shapes, comments and data validation.</p>
        </div>
        <div>
          <h3>3. Download or print</h3>
          <p>Save as XLSX, CSV, TSV or PDF, or print. Nothing is uploaded anywhere.</p>
        </div>
      </div>
    </main>
  );
}
