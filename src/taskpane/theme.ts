import { createDarkTheme, createLightTheme, type BrandVariants, type Theme } from '@fluentui/react-components';

/** Builds a 16-step Fluent brand ramp around Excel green (#107C41 at step 80). */
function excelRamp(): BrandVariants {
  const base = [0x10, 0x7c, 0x41];
  const mix = (target: number, t: number) => base.map((c) => Math.round(c + (target - c) * t));
  const hex = (rgb: number[]) => `#${rgb.map((c) => c.toString(16).padStart(2, '0')).join('')}`;
  const darker = [0.86, 0.76, 0.66, 0.56, 0.45, 0.33, 0.17];
  const lighter = [0.12, 0.24, 0.36, 0.48, 0.6, 0.72, 0.84, 0.94];
  const steps = [...darker.map((t) => hex(mix(0, t))), hex(base), ...lighter.map((t) => hex(mix(255, t)))];
  const keys = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160] as const;
  return Object.fromEntries(keys.map((k, i) => [k, steps[i]])) as unknown as BrandVariants;
}

const brand = excelRamp();

export const lightTheme: Theme = createLightTheme(brand);

export const darkTheme: Theme = {
  ...createDarkTheme(brand),
  colorBrandForeground1: brand[110],
  colorBrandForeground2: brand[120],
  colorBrandForegroundLink: brand[110],
  colorBrandForegroundLinkHover: brand[120],
};

function luminance(hex: string): number {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return 1;
  const n = parseInt(m[1], 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
}

/** Excel's own theme when Office reports it, otherwise the system appearance. */
export function prefersDark(): boolean {
  try {
    const bg = typeof Office !== 'undefined' ? Office.context?.officeTheme?.bodyBackgroundColor : undefined;
    if (bg) return luminance(bg) < 0.45;
  } catch {
    // officeTheme isn't available on every platform.
  }
  return typeof matchMedia !== 'undefined' && matchMedia('(prefers-color-scheme: dark)').matches;
}
