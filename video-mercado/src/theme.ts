import { Easing } from "remotion";

export const theme = {
  colors: {
    bg: "#070B16",
    panel: "rgba(255,255,255,0.05)",
    line: "rgba(255,255,255,0.12)",
    text: "#EEF2F8",
    muted: "#9AA6BA",
    dim: "#64708A",
    teal: "#1FE0C8",
    sky: "#60A5FA",
    pink: "#EC4899",
    violet: "#8B5CF6",
    ok: "#34D399",
    warn: "#FBBF24",
    red: "#FB7185",
  },
  grad: "linear-gradient(120deg, #1FE0C8, #38BDF8 50%, #EC4899)",
  fonts: { display: "Unbounded, sans-serif", body: "Manrope, sans-serif", mono: "'Courier New', monospace" },
  ease: {
    out: Easing.bezier(0.16, 1, 0.3, 1),
    inOut: Easing.bezier(0.83, 0, 0.17, 1),
    in: Easing.bezier(0.7, 0, 0.84, 0),
  },
  spring: {
    snappy: { damping: 14, stiffness: 160, mass: 0.6 },
    smooth: { damping: 20, stiffness: 90, mass: 1 },
    bouncy: { damping: 11, stiffness: 170, mass: 0.7 },
  },
} as const;

export const FPS = 30;
export const SCENES = [123, 263, 340, 385, 284, 169, 385, 246, 246] as const;
export const STARTS = SCENES.reduce<number[]>((a, d, i) => [...a, i === 0 ? 0 : a[i - 1] + SCENES[i - 1]], []);
export const TOTAL = SCENES.reduce((a, b) => a + b, 0);
