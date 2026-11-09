/** An icon of the editor's set, as an inline SVG that takes the text colour. */
import { type IconName, iconSvg } from '@sumline/core/views';

export function Icon({ name, size, className }: { name: IconName; size?: number; className?: string }) {
  return (
    <span
      className={`mi${className ? ` ${className}` : ''}`}
      aria-hidden="true"
      style={size ? { width: size, height: size } : undefined}
      // The SVG strings are bundled package assets, not user content.
      dangerouslySetInnerHTML={{ __html: iconSvg(name) }}
    />
  );
}
