/**
 * Print: a preview of the printed pages with the page options (scope, paper, orientation,
 * scaling, margins, page order, alignment, gridlines, notes, headings, header and footer,
 * repeated frozen rows and columns). The pages can be saved as a PDF or sent to the printer.
 */
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  buildPages, canvasMeasure, DEFAULT_PRINT_OPTIONS, loadPrintFonts, type Measure, type Page, pagesToPdf, pageSvg, PRINT_CHOICES, PRINT_SWITCHES,
  type PrintOptions, printPages, PT_TO_PX, rasterize, saveFile,
} from '@sumline/core/views';
import { Icon } from '../Icon';
import { Modal, Switch } from '../Modal';
import { usePersistent, useUI } from '../context';

function PagePreview({ page, index, width, root }: { page: Page; index: number; width: number; root: HTMLElement | null }) {
  const ref = useRef<HTMLDivElement>(null);
  const [visible, setVisible] = useState(index < 2);
  useEffect(() => {
    const el = ref.current;
    if (!el || visible || typeof IntersectionObserver === 'undefined') {
      if (!visible) setVisible(true);
      return;
    }
    const io = new IntersectionObserver((entries) => {
      if (entries.some((e) => e.isIntersecting)) {
        setVisible(true);
        io.disconnect();
      }
    }, { root, rootMargin: '600px 0px' });
    io.observe(el);
    return () => io.disconnect();
  }, [visible, root]);
  const height = (width * page.height) / page.width;
  const markup = useMemo(() => (visible ? pageSvg(page, { width: '100%', height: '100%' }) : ''), [visible, page]);
  return (
    <div
      ref={ref}
      className="sl-print-page"
      data-page={index + 1}
      data-testid="print-page"
      style={{ width, height }}
      // Generated from the document by the print layout (escaped text).
      dangerouslySetInnerHTML={{ __html: markup }}
    />
  );
}

export function PrintDialog({ onClose }: { onClose: () => void }) {
  const { doc, phone, notify } = useUI();
  const [stored, setStored] = usePersistent<Partial<PrintOptions>>('print-options', {});
  const opts: PrintOptions = useMemo(() => ({ ...DEFAULT_PRINT_OPTIONS, ...stored }), [stored]);
  const set = <K extends keyof PrintOptions>(key: K, value: PrintOptions[K]) => setStored({ ...stored, [key]: value });
  const [pages, setPages] = useState<Page[] | null>(null);
  const [busy, setBusy] = useState(false);
  const [zoom, setZoom] = useState<number | null>(null);
  const [current, setCurrent] = useState(1);
  const [paneWidth, setPaneWidth] = useState(600);
  const [showOptions, setShowOptions] = useState(!phone);
  const measure = useRef<Measure | null>(null);
  const pane = useRef<HTMLDivElement>(null);
  const title = doc.fileName.replace(/\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i, '');

  useEffect(() => {
    let alive = true;
    const t = setTimeout(async () => {
      await loadPrintFonts();
      if (!alive) return;
      measure.current ??= canvasMeasure();
      setPages(buildPages(doc, opts, measure.current, { title: doc.fileName }));
    }, 30);
    return () => {
      alive = false;
      clearTimeout(t);
    };
  }, [doc, opts]);

  useEffect(() => {
    const el = pane.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setPaneWidth(el.clientWidth));
    ro.observe(el);
    setPaneWidth(el.clientWidth);
    return () => ro.disconnect();
  }, []);

  const first = pages?.[0];
  const fit = first ? Math.max(0.1, (paneWidth - 36) / (first.width * PT_TO_PX)) : 1;
  const scale = zoom ?? fit;
  const pageWidth = (p: Page) => p.width * PT_TO_PX * scale;

  const onScroll = () => {
    const el = pane.current;
    if (!el || !pages) return;
    const mid = el.scrollTop + el.clientHeight / 3;
    let y = 18;
    for (let i = 0; i < pages.length; i++) {
      const h = (pageWidth(pages[i]) * pages[i].height) / pages[i].width + 16;
      if (mid < y + h) {
        setCurrent(i + 1);
        return;
      }
      y += h;
    }
    setCurrent(pages.length);
  };

  const download = async () => {
    if (!pages || busy) return;
    setBusy(true);
    try {
      const bytes = await pagesToPdf(pages, { rasterize, measure: measure.current ?? canvasMeasure() }, { title });
      saveFile(bytes, `${title || 'spreadsheet'}.pdf`, 'application/pdf');
    } catch (e) {
      notify(`The PDF could not be created: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  const print = () => {
    if (pages) printPages(pages);
  };

  const zoomBy = (f: number) => setZoom(Math.max(0.25, Math.min(4, Math.round(scale * f * 100) / 100)));

  return (
    <Modal title="Print" icon="print" onClose={onClose} width={Math.min(1300, window.innerWidth - 16)} height={phone ? '100%' : 'min(820px, calc(100vh - 40px))'} testId="print-dialog" bodyClassName="sl-print-body">
      <div className={`sl-print${showOptions ? '' : ' no-options'}`}>
        <div className="sl-print-viewer">
          <div className="sl-print-toolbar">
            <button type="button" className="sl-print-tool" aria-label={showOptions ? 'Hide options' : 'Show options'} onClick={() => setShowOptions(!showOptions)}>
              <Icon name="menu" />
            </button>
            <span className="sl-print-name">{title || 'spreadsheet'}.pdf</span>
            <span className="sl-print-count" data-testid="print-page-count">
              <span className="current">{pages ? current : 0}</span> / {pages?.length ?? 0}
            </span>
            <span className="sep" />
            <button type="button" className="sl-print-tool" aria-label="Zoom out" onClick={() => zoomBy(1 / 1.2)}>
              <Icon name="remove" />
            </button>
            <span className="sl-print-zoom">{Math.round(scale * 100)}%</span>
            <button type="button" className="sl-print-tool" aria-label="Zoom in" onClick={() => zoomBy(1.2)}>
              <Icon name="add" />
            </button>
            <button type="button" className="sl-print-tool" aria-label="Fit to width" onClick={() => setZoom(null)}>
              <Icon name="fit_screen" />
            </button>
            <span className="grow" />
            <button type="button" className="sl-print-tool" aria-label="Download PDF" title="Download PDF" data-testid="print-download" disabled={!pages || busy} onClick={download}>
              <Icon name={busy ? 'hourglass_empty' : 'download'} />
            </button>
            <button type="button" className="sl-print-tool" aria-label="Print" title="Print" data-testid="print-print" disabled={!pages} onClick={print}>
              <Icon name="print" />
            </button>
          </div>
          <div className="sl-print-preview" ref={pane} onScroll={onScroll}>
            {!pages && <div className="sl-print-wait">Preparing pages…</div>}
            {pages?.map((p, i) => <PagePreview key={`${i}-${pages.length}`} page={p} index={i} width={pageWidth(p)} root={pane.current} />)}
          </div>
        </div>
        {showOptions && (
          <div className="sl-print-options" data-testid="print-options">
            {PRINT_CHOICES.map((c) => (
              <label key={c.key} className="opt">
                <span>{c.label}:</span>
                <select className="sl-select" value={String(opts[c.key])} data-testid={`print-${c.key}`} onChange={(e) => set(c.key, e.target.value as never)}>
                  {c.options.map((o) => (
                    <option key={String(o.value)} value={String(o.value)}>{o.label}</option>
                  ))}
                </select>
              </label>
            ))}
            <div className="sl-print-switches">
              {PRINT_SWITCHES.map((s) => (
                <Switch key={s.key} label={s.label} checked={!!opts[s.key]} testId={`print-${s.key}`} onChange={(v) => set(s.key, v as never)} />
              ))}
            </div>
            <div className="sl-print-actions">
              <button type="button" className="sl-button" disabled={!pages || busy} onClick={download}>{busy ? 'Creating PDF…' : 'Download PDF'}</button>
              <button type="button" className="sl-button" disabled={!pages} onClick={print}>Print</button>
            </div>
          </div>
        )}
      </div>
    </Modal>
  );
}
