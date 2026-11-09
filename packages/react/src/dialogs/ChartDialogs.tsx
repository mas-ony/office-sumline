/**
 * Charts: the Insert chart dialog (data range and a gallery of chart types) and the chart settings
 * panel (Configuration, Style and Series tabs).
 */
import { useMemo, useState } from 'react';
import {
  CHART_SIDES, CHART_TYPE_INFO, type ChartConfig, chartLines, chartSettingsOf, chartSettingsPatch, type ChartStyle, type ChartType, insertChart,
  LINE_DASHES, type LineDash, localRef, numText, parseNum, rangeName, readChartData, relayoutChart, resolveRef, type SeriesStyle, type Side,
  type Trendline, TRENDLINES, trimRange, withSeriesStyle,
} from '@sumline/core/views';
import { ColorField } from '../ColorField';
import { useUI } from '../context';
import { Icon } from '../Icon';
import { Modal, Switch } from '../Modal';

export function ChartInsertDialog({ onClose }: { onClose: () => void }) {
  const { doc, grid, phone, selectObject } = useUI();
  const sheet = doc.sheet;
  const [range, setRange] = useState(() => rangeName(trimRange(sheet, doc.range)));
  const [type, setType] = useState<ChartType>('line');
  const [error, setError] = useState('');

  const insert = () => {
    const res = insertChart(doc, range, type, phone);
    if ('error' in res) {
      setError(res.error);
      return;
    }
    onClose();
    if (res.chart) {
      grid?.scrollIntoView(res.chart.from.r, res.chart.from.c);
      selectObject(res.chart.id);
    }
  };

  return (
    <Modal title="Insert chart" icon="addchart" onClose={onClose} width={430} testId="chart-insert-dialog">
      <div className="sl-form">
        <div className="sl-field">
          <label htmlFor="chart-range">Data to analyze</label>
          <div className="sl-range-input">
            <input id="chart-range" className="sl-input" value={range} spellCheck={false} data-testid="chart-range" onChange={(e) => { setRange(e.target.value); setError(''); }} />
            <button type="button" title="Use the selection" aria-label="Use the selection" onClick={() => setRange(rangeName(trimRange(sheet, doc.range)))}>
              <Icon name="highlight_alt" size={18} />
            </button>
          </div>
        </div>
      </div>
      <div className="sl-chart-types" role="listbox" aria-label="Chart type">
        {CHART_TYPE_INFO.map((t) => (
          <button
            key={t.type}
            type="button"
            role="option"
            aria-selected={type === t.type}
            className={`sl-chart-type${type === t.type ? ' selected' : ''}`}
            data-testid={`chart-type-${t.type}`}
            onClick={() => setType(t.type)}
            onDoubleClick={insert}
          >
            <span className="tile"><Icon name={t.icon} /></span>
            <span>{t.label}</span>
          </button>
        ))}
      </div>
      {error && <div className="sl-error" role="alert">{error}</div>}
      <div className="sl-stack-buttons">
        <button type="button" className="sl-button wide" data-testid="chart-insert" onClick={insert}>Insert</button>
        <button type="button" className="sl-button wide" onClick={onClose}>Cancel</button>
      </div>
    </Modal>
  );
}

// -------------------------------------------------------------------------------------------------
// Chart settings

export function ChartSettings({ id, onClose }: { id: string; onClose: () => void }) {
  const { doc, phone } = useUI();
  const sheet = doc.sheet;
  const chart = sheet.charts.find((c) => c.id === id);
  const [tab, setTab] = useState<'config' | 'style' | 'series'>('config');
  const [initial] = useState(() => chartSettingsOf(doc, chart, sheet));
  const [config, setConfig] = useState<ChartConfig>(initial.config);
  const [custom, setCustom] = useState(initial.custom);
  const [style, setStyle] = useState<ChartStyle>(initial.style);
  const [seriesStyles, setSeriesStyles] = useState<SeriesStyle[]>(initial.seriesStyles);

  const target = useMemo(() => resolveRef(doc, config.range, sheet), [doc, config.range, sheet]);
  const lines = useMemo(() => (target ? chartLines(doc, target.sheet, target.range, config.seriesIn, config.headers) : []), [doc, target, config.seriesIn, config.headers]);
  const data = useMemo(() => (chart ? readChartData(doc, chart) : null), [doc, chart]);

  if (!chart) return null;

  /** Re-derives labels and series when the range, orientation or header option changes. */
  const relayout = (patch: Partial<ChartConfig>) => {
    const next = relayoutChart(doc, sheet, config, patch);
    if (next.laidOut) setCustom(false);
    setConfig(next.config);
  };

  const seriesNames: string[] = custom ? (data?.series.map((s) => s.name) ?? []) : config.series.map((i) => lines[i]?.name ?? `Series ${i + 1}`);

  const save = () => {
    doc.updateChart(id, chartSettingsPatch(chart, { config, custom, style, seriesStyles }, target));
    onClose();
  };

  const setSeriesStyle = (i: number, p: Partial<SeriesStyle>) => setSeriesStyles((list) => withSeriesStyle(list, i, p));

  const footer = (
    <div className="sl-dialog-buttons">
      <button type="button" className="sl-button wide" data-testid="chart-save" onClick={save}>Save</button>
      <button type="button" className="sl-button wide" onClick={onClose}>Cancel</button>
    </div>
  );

  return (
    <Modal title="Chart settings" icon="bar_chart" onClose={onClose} panel width={phone ? window.innerWidth : 440} testId="chart-settings">
      <div className="sl-subtabs" role="tablist">
        {(['config', 'style', 'series'] as const).map((t) => (
          <button key={t} type="button" role="tab" aria-selected={tab === t} className={tab === t ? 'selected' : ''} data-testid={`chart-tab-${t}`} onClick={() => setTab(t)}>
            {t === 'config' ? 'Configuration' : t === 'style' ? 'Style' : 'Series'}
          </button>
        ))}
      </div>
      {tab === 'config' && (
        <div className="sl-form">
          <div className="sl-row">
            <div className="sl-field">
              <label htmlFor="chart-type">Chart type</label>
              <select id="chart-type" className="sl-select" value={config.type} data-testid="chart-type" onChange={(e) => setConfig({ ...config, type: e.target.value as ChartType })}>
                {CHART_TYPE_INFO.map((t) => <option key={t.type} value={t.type}>{t.name}</option>)}
              </select>
            </div>
            <div className="sl-field">
              <label htmlFor="chart-orientation">Data orientation</label>
              <select id="chart-orientation" className="sl-select" value={config.seriesIn} data-testid="chart-orientation" disabled={custom && !target} onChange={(e) => relayout({ seriesIn: e.target.value as 'rows' | 'columns' })}>
                <option value="rows">Horizontal</option>
                <option value="columns">Vertical</option>
              </select>
            </div>
          </div>
          <div className="sl-field">
            <label htmlFor="chart-data-range">Data range</label>
            <div className="sl-range-input">
              <input id="chart-data-range" className={`sl-input${config.range && !target ? ' invalid' : ''}`} value={config.range} spellCheck={false} data-testid="chart-data-range" onChange={(e) => setConfig({ ...config, range: e.target.value })} onBlur={() => target && relayout({})} />
              <button type="button" title="Use the selection" aria-label="Use the selection" onClick={() => relayout({ range: rangeName(trimRange(sheet, doc.range)) })}>
                <Icon name="highlight_alt" size={18} />
              </button>
            </div>
          </div>
          {!custom && target ? (
            <>
              <div className="sl-field">
                <label htmlFor="chart-labels">Labels</label>
                <select id="chart-labels" className="sl-select" value={config.labels ?? ''} data-testid="chart-labels" onChange={(e) => setConfig({ ...config, labels: e.target.value === '' ? null : Number(e.target.value) })}>
                  <option value="">None</option>
                  {lines.map((l) => <option key={l.index} value={l.index}>{l.name}</option>)}
                </select>
              </div>
              <div className="sl-field">
                <div className="label">Series</div>
                {config.series.map((s, k) => (
                  <div key={k} className="sl-series-row">
                    <select className="sl-select" value={s} aria-label={`Series ${k + 1}`} data-testid="chart-series" onChange={(e) => setConfig({ ...config, series: config.series.map((x, j) => (j === k ? Number(e.target.value) : x)) })}>
                      {lines.map((l) => <option key={l.index} value={l.index}>{l.name}</option>)}
                    </select>
                    <button type="button" className="sl-icon-btn" aria-label="Remove series" data-testid="chart-series-remove" onClick={() => {
                      setConfig({ ...config, series: config.series.filter((_, j) => j !== k) });
                      setSeriesStyles((list) => list.filter((_, j) => j !== k));
                    }}>
                      <Icon name="close" />
                    </button>
                  </div>
                ))}
                <button type="button" className="sl-link" data-testid="chart-series-add" onClick={() => {
                  const unused = lines.find((l) => l.index !== config.labels && !config.series.includes(l.index)) ?? lines[0];
                  if (unused) setConfig({ ...config, series: [...config.series, unused.index] });
                }}>Add new serie</button>
              </div>
              <Switch label={config.seriesIn === 'columns' ? 'Use the first row as headers' : 'Use the first column as headers'} checked={config.headers} testId="chart-headers" onChange={(v) => setConfig({ ...config, headers: v })} />
            </>
          ) : (
            <div className="sl-field">
              <div className="label">Series</div>
              {(data?.series ?? []).map((s) => <div key={s.index} className="sl-hint">{s.name}: {localRef(s.source.values, sheet)}</div>)}
            </div>
          )}
        </div>
      )}
      {tab === 'style' && (
        <div className="sl-form">
          <div className="sl-row">
            <div className="sl-field">
              <label htmlFor="chart-title">Chart title</label>
              <input id="chart-title" className="sl-input" value={style.title ?? ''} data-testid="chart-title" onChange={(e) => setStyle({ ...style, title: e.target.value })} />
            </div>
            <div className="sl-field">
              <label htmlFor="chart-title-pos">Title position</label>
              <select id="chart-title-pos" className="sl-select" value={style.titlePosition ?? ''} onChange={(e) => setStyle({ ...style, titlePosition: (e.target.value || undefined) as Side | undefined })}>
                <option value="" />
                {CHART_SIDES.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
              </select>
            </div>
          </div>
          <div className="sl-field">
            <label htmlFor="chart-subtitle">Subtitle</label>
            <input id="chart-subtitle" className="sl-input" value={style.subtitle ?? ''} onChange={(e) => setStyle({ ...style, subtitle: e.target.value })} />
          </div>
          <div className="sl-row">
            <div className="sl-field">
              <label htmlFor="chart-legend-title">Legend title</label>
              <input id="chart-legend-title" className="sl-input" value={style.legendTitle ?? ''} onChange={(e) => setStyle({ ...style, legendTitle: e.target.value })} />
            </div>
            <div className="sl-field">
              <label htmlFor="chart-legend-pos">Legend position</label>
              <select id="chart-legend-pos" className="sl-select" value={style.legendPosition ?? ''} data-testid="chart-legend-position" onChange={(e) => setStyle({ ...style, legendPosition: (e.target.value || undefined) as Side | undefined })}>
                <option value="" />
                {CHART_SIDES.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
              </select>
            </div>
          </div>
          <div className="sl-row">
            <Switch label="Show legend" checked={style.showLegend} testId="chart-show-legend" onChange={(v) => setStyle({ ...style, showLegend: v })} />
            <Switch label="Gridlines" checked={style.gridlines} testId="chart-gridlines" onChange={(v) => setStyle({ ...style, gridlines: v })} />
          </div>
          <Switch label="Value labels" checked={style.valueLabels} testId="chart-value-labels" onChange={(v) => setStyle({ ...style, valueLabels: v })} />
          <div className="sl-row">
            <div className="sl-field">
              <label htmlFor="chart-x-title">X axis title</label>
              <input id="chart-x-title" className="sl-input" value={style.xAxisTitle ?? ''} onChange={(e) => setStyle({ ...style, xAxisTitle: e.target.value })} />
            </div>
            <div className="sl-field">
              <label htmlFor="chart-y-title">Y axis title</label>
              <input id="chart-y-title" className="sl-input" value={style.yAxisTitle ?? ''} onChange={(e) => setStyle({ ...style, yAxisTitle: e.target.value })} />
            </div>
          </div>
          <div className="sl-field">
            <label htmlFor="chart-sec-title">Secondary axis title</label>
            <input id="chart-sec-title" className="sl-input" value={style.secondaryAxisTitle ?? ''} onChange={(e) => setStyle({ ...style, secondaryAxisTitle: e.target.value })} />
          </div>
          <div className="sl-row">
            <div className="sl-field">
              <label htmlFor="chart-ymin">Y min</label>
              <input id="chart-ymin" className="sl-input" inputMode="decimal" value={style.yMin} onChange={(e) => setStyle({ ...style, yMin: e.target.value })} />
            </div>
            <div className="sl-field">
              <label htmlFor="chart-ymax">Y max</label>
              <input id="chart-ymax" className="sl-input" inputMode="decimal" value={style.yMax} onChange={(e) => setStyle({ ...style, yMax: e.target.value })} />
            </div>
          </div>
        </div>
      )}
      {tab === 'series' && (
        <div className="sl-form">
          {seriesNames.map((name, i) => {
            const st = seriesStyles[i] ?? {};
            return (
              <div key={i} className="sl-series-style">
                <h4>{name}</h4>
                <div className="sl-row">
                  <div className="sl-field">
                    <label>Color</label>
                    <ColorField value={st.color ?? ''} testId={`chart-series-color-${i}`} onChange={(hex) => setSeriesStyle(i, { color: hex || undefined })} />
                  </div>
                  <div className="sl-field">
                    <label htmlFor={`sw-${i}`}>Line width</label>
                    <input id={`sw-${i}`} className="sl-input" inputMode="numeric" value={numText(st.lineWidth)} onChange={(e) => setSeriesStyle(i, { lineWidth: parseNum(e.target.value) })} />
                  </div>
                </div>
                <div className="sl-row">
                  <div className="sl-field">
                    <label htmlFor={`sd-${i}`}>Line style</label>
                    <select id={`sd-${i}`} className="sl-select" value={st.lineStyle ?? 'solid'} onChange={(e) => setSeriesStyle(i, { lineStyle: e.target.value as LineDash })}>
                      {LINE_DASHES.map((d) => <option key={d.key} value={d.key}>{d.label}</option>)}
                    </select>
                  </div>
                  <div className="sl-field">
                    <label htmlFor={`st-${i}`}>Trendline</label>
                    <select id={`st-${i}`} className="sl-select" value={st.trendline ?? 'none'} data-testid={`chart-series-trend-${i}`} onChange={(e) => setSeriesStyle(i, { trendline: e.target.value as Trendline })}>
                      {TRENDLINES.map((d) => <option key={d.key} value={d.key}>{d.label}</option>)}
                    </select>
                  </div>
                </div>
                <div className="sl-row">
                  <Switch label="Markers" checked={!!st.markers} onChange={(v) => setSeriesStyle(i, { markers: v })} />
                  <Switch label="Secondary axis" checked={!!st.secondaryAxis} onChange={(v) => setSeriesStyle(i, { secondaryAxis: v })} />
                </div>
              </div>
            );
          })}
        </div>
      )}
      {footer}
    </Modal>
  );
}
