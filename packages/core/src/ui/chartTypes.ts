/** Chart types as listed in the Insert chart gallery and the chart settings. */
import type { ChartType } from '../core/model/drawing';
import type { IconName } from './icons';

export interface ChartTypeInfo {
  type: ChartType;
  /** Gallery label. */
  label: string;
  /** Name in the chart settings type list. */
  name: string;
  icon: IconName;
}

export const CHART_TYPE_INFO: ChartTypeInfo[] = [
  { type: 'line', label: 'Line', name: 'Line Chart', icon: 'show_chart' },
  { type: 'area', label: 'Area', name: 'Area Chart', icon: 'area_chart' },
  { type: 'stacked-area', label: 'Stacked Area', name: 'Stacked Area Chart', icon: 'stacked_line_chart' },
  { type: 'percent-area', label: '100% Stacked Area', name: '100% Stacked Area Chart', icon: 'ssid_chart' },
  { type: 'bar', label: 'Bar', name: 'Bar Chart', icon: 'align_horizontal_left' },
  { type: 'stacked-bar', label: 'Stacked Bar', name: 'Stacked Bar Chart', icon: 'horizontal_split' },
  { type: 'percent-bar', label: '100% Stacked Bar', name: '100% Stacked Bar Chart', icon: 'view_headline' },
  { type: 'column', label: 'Column', name: 'Column Chart', icon: 'bar_chart' },
  { type: 'stacked-column', label: 'Stacked Column', name: 'Stacked Column Chart', icon: 'stacked_bar_chart' },
  { type: 'percent-column', label: '100% Stacked Column', name: '100% Stacked Column Chart', icon: 'leaderboard' },
  { type: 'pie', label: 'Pie', name: 'Pie Chart', icon: 'pie_chart' },
  { type: 'doughnut', label: 'Doughnut', name: 'Doughnut Chart', icon: 'donut_large' },
  { type: 'scatter', label: 'Scatter', name: 'Scatter Chart', icon: 'scatter_plot' },
  { type: 'bubble', label: 'Bubble', name: 'Bubble Chart', icon: 'bubble_chart' },
  { type: 'radar', label: 'Radar', name: 'Radar Chart', icon: 'track_changes' },
  { type: 'filled-radar', label: 'Filled radar', name: 'Filled Radar Chart', icon: 'radar' },
  { type: 'histogram', label: 'Histogram', name: 'Histogram', icon: 'equalizer' },
  { type: 'pareto', label: 'Pareto', name: 'Pareto Chart', icon: 'auto_graph' },
  { type: 'box-whisker', label: 'Box and whisker', name: 'Box and Whisker Chart', icon: 'vertical_align_center' },
  { type: 'waterfall', label: 'Waterfall', name: 'Waterfall Chart', icon: 'waterfall_chart' },
  { type: 'funnel', label: 'Funnel', name: 'Funnel Chart', icon: 'filter_alt' },
  { type: 'treemap', label: 'Treemap', name: 'Treemap Chart', icon: 'view_quilt' },
  { type: 'stock', label: 'Stock', name: 'Stock Chart', icon: 'candlestick_chart' },
];

export function chartTypeInfo(type: ChartType): ChartTypeInfo {
  return CHART_TYPE_INFO.find((t) => t.type === type) ?? CHART_TYPE_INFO[0];
}
