/** A draggable modal window with a title bar (icon, title, close), or a panel on the right. */
import { useEffect, useLayoutEffect, useRef, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import type { IconName } from '@sumline/core/views';
import { Icon } from './Icon';
import { useUI } from './context';

export interface ModalProps {
  title: string;
  icon?: IconName;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
  width?: number;
  height?: number | string;
  /** Docked to the right edge, full height (chart settings). */
  panel?: boolean;
  /** Dims and blocks the page behind. */
  blocking?: boolean;
  testId?: string;
  bodyClassName?: string;
  bodyStyle?: React.CSSProperties;
}

export function Modal({ title, icon, onClose, children, footer, width = 430, height, panel, blocking, testId, bodyClassName, bodyStyle }: ModalProps) {
  const { layer } = useUI();
  const ref = useRef<HTMLDivElement>(null);
  const drag = useRef<{ dx: number; dy: number; id: number } | null>(null);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el || panel) return;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    el.style.left = `${Math.max(8, (window.innerWidth - w) / 2)}px`;
    el.style.top = `${Math.max(8, Math.min((window.innerHeight - h) / 2, 120))}px`;
  }, [panel]);

  useEffect(() => {
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !e.defaultPrevented) {
        // Escape inside an open drop-down of the dialog closes the drop-down only.
        if ((e.target as HTMLElement | null)?.closest?.('.sl-picker, .sl-color, .sl-popup')) return;
        e.preventDefault();
        onClose();
      }
    };
    window.addEventListener('keydown', key);
    const first = ref.current?.querySelector<HTMLElement>('[data-autofocus], input:not([type=hidden]):not([disabled]), select, textarea');
    if (first && window.matchMedia?.('(pointer: fine)').matches) first.focus({ preventScroll: true });
    return () => window.removeEventListener('keydown', key);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const startDrag = (e: React.PointerEvent) => {
    const el = ref.current;
    if (!el || panel || (e.target as HTMLElement).closest('button')) return;
    const rect = el.getBoundingClientRect();
    drag.current = { dx: e.clientX - rect.left, dy: e.clientY - rect.top, id: e.pointerId };
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
  };
  const moveDrag = (e: React.PointerEvent) => {
    const d = drag.current;
    const el = ref.current;
    if (!d || !el || d.id !== e.pointerId) return;
    const x = Math.max(-el.offsetWidth + 80, Math.min(window.innerWidth - 80, e.clientX - d.dx));
    const y = Math.max(0, Math.min(window.innerHeight - 40, e.clientY - d.dy));
    el.style.left = `${x}px`;
    el.style.top = `${y}px`;
  };

  return createPortal(
    <div className={`sl-modal-layer${blocking ? ' blocking' : ''}`} onPointerDown={(e) => blocking && e.target === e.currentTarget && onClose()}>
      <div
        ref={ref}
        className={`sl-modal${panel ? ' panel-right' : ''}`}
        role="dialog"
        aria-modal={blocking || undefined}
        aria-label={title}
        data-testid={testId}
        style={{ width: panel ? Math.min(width, window.innerWidth) : width, height }}
      >
        <div className="sl-modal-title" onPointerDown={startDrag} onPointerMove={moveDrag} onPointerUp={() => (drag.current = null)}>
          {icon && <Icon name={icon} />}
          <span className="text">{title}</span>
          <button type="button" className="sl-modal-close" aria-label="Close" data-testid="modal-close" onClick={onClose}>
            <Icon name="close" />
          </button>
        </div>
        <div className={`sl-modal-body${bodyClassName ? ` ${bodyClassName}` : ''}`} style={bodyStyle}>{children}</div>
        {footer && <div className="sl-modal-foot">{footer}</div>}
      </div>
    </div>,
    layer,
  );
}

/** Labeled on/off switch. */
export function Switch({ label, checked, onChange, testId }: { label: string; checked: boolean; onChange: (v: boolean) => void; testId?: string }) {
  return (
    <label className="sl-switch">
      <input type="checkbox" role="switch" checked={checked} onChange={(e) => onChange(e.target.checked)} data-testid={testId} />
      <span>{label}</span>
    </label>
  );
}
