/**
 * The footer row (View > Toggle footer row): a band under the grid aligned with the columns,
 * showing a label or the result of a formula per column (e.g. =SUM(B:B)).
 */
import { useEffect, useState } from 'react';
import { ArrayValue, CellError, formatValue, RefValue, type Scalar, type ViewTransform } from '@sumline/core/views';
import { useUI } from './context';

export function FooterRow() {
  const ui = useUI();
  const { doc, grid, state } = ui;
  const [vt, setVt] = useState<ViewTransform | null>(() => grid?.transform() ?? null);
  useEffect(() => {
    const off = ui.ui.onGridPaint((v) => setVt(v));
    grid?.requestRender();
    return () => {
      off();
    };
  }, [ui.ui, grid]);
  const footer = doc.sheet.footer;
  if (!footer || !vt) return null;
  void state.version;
  const value = (text: string, c: number): string => {
    if (!text.startsWith('=')) return text;
    try {
      let v = doc.engine.evaluateText(text.slice(1), doc.sheet, 1048575, c) as unknown;
      if (v instanceof RefValue) v = doc.engine.cellValue(v.single.sheet, v.single.range.r0, v.single.range.c0);
      if (v instanceof ArrayValue) v = v.get(0, 0);
      if (v instanceof CellError) return v.code;
      const numFmt = doc.styleAt(Math.max(0, (doc.sheet.usedRange()?.r1 ?? 0)), c).numFmt;
      return formatValue(v as Scalar, typeof v === 'number' ? numFmt : 'General', doc.wb.date1904).text;
    } catch {
      return '#VALUE!';
    }
  };
  const { cols } = vt.visibleMain();
  const frozen = Array.from({ length: vt.fc }, (_, i) => i);
  const columns = [...frozen, ...Array.from({ length: Math.max(0, cols[1] - cols[0] + 1) }, (_, i) => cols[0] + i)];
  return (
    <div className="sl-footer-row" data-testid="footer-row" aria-label="Footer row">
      <div className="sl-footer-cell label" style={{ left: 0, width: vt.v.headerW }}>Σ</div>
      {columns.map((c) => {
        const rect = vt.cellRect(0, c);
        if (rect.w <= 0 || c >= footer.length) return null;
        const text = footer[c] ?? '';
        if (!text) return null;
        const isLabel = !text.startsWith('=');
        return (
          <div key={c} className={`sl-footer-cell${isLabel ? ' label' : ''}`} style={{ left: rect.x, width: rect.w }} title={text}>
            {value(text, c)}
          </div>
        );
      })}
    </div>
  );
}
