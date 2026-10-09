import { Easing } from "remotion";

export const theme = {
  colors: {
    bg: "#1A1617",
    panel: "rgba(255,255,255,0.05)",
    line: "rgba(255,255,255,0.12)",
    text: "#EEF2F8",
    muted: "#9AA6BA",
    dim: "#64708A",
    teal: "#EC6707",
    sky: "#F39200",
    pink: "#FF8A3D",
    violet: "#F39200",
    ok: "#34D399",
    warn: "#FBBF24",
    red: "#FB7185",
  },
  grad: "linear-gradient(120deg, #EC6707, #F39200)",
  fonts: { display: "Harabara, sans-serif", body: "'Source Sans Pro', sans-serif", mono: "'Courier New', monospace" },
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
export const SCENES = [284, 320, 412, 647, 467, 335, 315, 205, 270] as const;
export const STARTS = SCENES.reduce<number[]>((a, d, i) => [...a, i === 0 ? 0 : a[i - 1] + SCENES[i - 1]], []);
export const TOTAL = SCENES.reduce((a, b) => a + b, 0);
