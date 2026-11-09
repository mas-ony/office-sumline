/** Search and replace: find text in this or all worksheets, step through matches, replace. */
import { useEffect, useMemo, useState } from 'react';
import { cellName, type SearchHit } from '@sumline/core/views';
import { useUI } from '../context';
import { Modal, Switch } from '../Modal';

export function SearchDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc, grid } = ui;
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [scope, setScope] = useState<'sheet' | 'all'>('sheet');
  const [matchCase, setMatchCase] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [inFormulas, setInFormulas] = useState(false);
  const [hiddenToo, setHiddenToo] = useState(false);
  const [hits, setHits] = useState<SearchHit[] | null>(null);
  const [index, setIndex] = useState(0);
  const [status, setStatus] = useState('');
  const opts = useMemo(() => ({ matchCase, wholeCell, inFormulas, allSheets: scope === 'all' }), [matchCase, wholeCell, inFormulas, scope]);

  // Options or the query changed: the result list is stale.
  useEffect(() => {
    setHits(null);
    setStatus('');
  }, [query, opts, hiddenToo]);

  useEffect(() => () => {
    if (grid) {
      grid.searchHit = undefined;
      grid.requestRender();
    }
  }, [grid]);

  const run = (): SearchHit[] => {
    let found = doc.find(query, opts);
    if (!hiddenToo) found = found.filter((h) => !h.sheet.isRowHidden(h.r) && !h.sheet.isColHidden(h.c));
    setHits(found);
    return found;
  };

  const show = (list: SearchHit[], i: number) => {
    if (!list.length) {
      setStatus('No matches');
      return;
    }
    const k = ((i % list.length) + list.length) % list.length;
    setIndex(k);
    const hit = list[k];
    const si = doc.wb.sheets.indexOf(hit.sheet);
    if (si !== doc.activeIndex) doc.setActiveSheet(si);
    doc.selectCell(hit.r, hit.c);
    if (grid) {
      grid.searchHit = { r: hit.r, c: hit.c };
      grid.requestRender();
    }
    setStatus(`${k + 1} of ${list.length}${scope === 'all' ? ` · ${hit.sheet.name}!${cellName(hit.r, hit.c)}` : ''}`);
  };

  const search = () => show(run(), 0);
  const step = (d: number) => {
    const list = hits ?? run();
    show(list, hits ? index + d : d > 0 ? 0 : list.length - 1);
  };
  const replaceOne = () => {
    const list = hits ?? run();
    if (!list.length) {
      setStatus('No matches');
      return;
    }
    const hit = list[index] ?? list[0];
    doc.replace(hit, query, replacement, opts);
    const next = run();
    show(next, Math.min(index, next.length - 1));
  };
  const replaceAll = () => {
    if (!query) return;
    let n = 0;
    doc.transact('Replace all', () => {
      for (const hit of run()) if (doc.replace(hit, query, replacement, opts)) n++;
    });
    setHits(null);
    setStatus(n ? `${n} cell${n === 1 ? '' : 's'} replaced` : 'No matches');
  };

  const disabled = !query;
  return (
    <Modal title="Search and replace" onClose={onClose} width={560} testId="search-dialog">
      <div className="sl-form" style={{ gridTemplateColumns: 'auto 1fr', alignItems: 'center', columnGap: 12 }}>
        <label htmlFor="sl-find">Find</label>
        <input id="sl-find" className="sl-input" value={query} data-testid="search-find" data-autofocus onChange={(e) => setQuery(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && (hits ? step(1) : search())} />
        <label htmlFor="sl-repl">Replace by</label>
        <input id="sl-repl" className="sl-input" value={replacement} data-testid="search-replace" onChange={(e) => setReplacement(e.target.value)} />
        <label htmlFor="sl-scope">Search</label>
        <select id="sl-scope" className="sl-select" value={scope} data-testid="search-scope" onChange={(e) => setScope(e.target.value as 'sheet' | 'all')}>
          <option value="sheet">This worksheet</option>
          <option value="all">All worksheets</option>
        </select>
        <span />
        <div style={{ display: 'grid', gap: 10 }}>
          <Switch label="Case sensitive" checked={matchCase} onChange={setMatchCase} testId="search-case" />
          <Switch label="Match all cell contents" checked={wholeCell} onChange={setWholeCell} testId="search-whole" />
          <Switch label="Search inside formulas" checked={inFormulas} onChange={setInFormulas} testId="search-formulas" />
          <Switch label="Replace cells that were hidden by a filter" checked={hiddenToo} onChange={setHiddenToo} testId="search-hidden" />
        </div>
      </div>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginTop: 22 }}>
        <button type="button" className="sl-button" disabled={disabled} data-testid="search-go" onClick={search}>Search</button>
        <button type="button" className="sl-button" disabled={disabled} data-testid="search-prev" onClick={() => step(-1)}>Prev</button>
        <button type="button" className="sl-button" disabled={disabled} data-testid="search-next" onClick={() => step(1)}>Next</button>
        <button type="button" className="sl-button" disabled={disabled || ui.state.readOnly} data-testid="search-replace-one" onClick={replaceOne}>Replace</button>
        <button type="button" className="sl-button" disabled={disabled || ui.state.readOnly} data-testid="search-replace-all" onClick={replaceAll}>Replace All</button>
      </div>
      <div className="sl-hint" style={{ minHeight: 20, marginTop: 10 }} role="status" data-testid="search-status">{status}</div>
    </Modal>
  );
}
