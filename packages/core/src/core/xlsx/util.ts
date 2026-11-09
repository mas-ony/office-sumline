import { serialFromYMD } from '../formula/dates';

/** Serial number of an ISO 8601 date-time ("2024-01-15T10:30:00Z"), as used by t="d" cells. */
export function serialFromDate(iso: string): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(iso.trim());
  if (!m) {
    const t = /^(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?/.exec(iso.trim());
    if (!t) return null;
    return (Number(t[1]) * 3600 + Number(t[2]) * 60 + Number(t[3] ?? 0)) / 86400;
  }
  const day = serialFromYMD(Number(m[1]), Number(m[2]), Number(m[3]));
  const secs = Number(m[4] ?? 0) * 3600 + Number(m[5] ?? 0) * 60 + Number(m[6] ?? 0);
  return day + secs / 86400;
}

/** EMU per pixel at 96 DPI. */
export const EMU_PER_PX = 9525;

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}
