import { Easing } from "remotion";

export const theme = {
  colors: {
    // Identidade do Future PDV (landing.css)
    bg: "#08070F",
    panel: "rgba(255,255,255,0.05)",
    line: "rgba(255,255,255,0.10)",
    text: "#EEF2F8",
    muted: "#9AA6BA",
    dim: "#64708A",
    teal: "#FF7A18",
    sky: "#22D3EE",
    pink: "#FF2E9A",
    violet: "#8B5CF6",
    ok: "#34D399",
    warn: "#FBBF24",
  },
  grad: "linear-gradient(120deg, #FF7A18, #FF2E9A 55%, #8B5CF6)",
  fonts: { display: "Unbounded, sans-serif", body: "Manrope, sans-serif" },
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
export const W = 1920;
export const H = 1080;
export const SCENES = [246, 263, 342, 250, 119] as const;
export const STARTS = SCENES.reduce<number[]>((a, d, i) => [...a, i === 0 ? 0 : a[i - 1] + SCENES[i - 1]], []);
export const TOTAL = SCENES.reduce((a, b) => a + b, 0);
