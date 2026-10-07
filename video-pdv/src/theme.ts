import { Easing } from "remotion";

export const theme = {
  colors: {
    // Identidade do Future PDV (landing.css)
    bg: "#070B16",
    panel: "rgba(255,255,255,0.05)",
    line: "rgba(255,255,255,0.10)",
    text: "#EEF2F8",
    muted: "#9AA6BA",
    dim: "#64708A",
    teal: "#00E0C6",
    sky: "#38BDF8",
    pink: "#FF2E9A",
    violet: "#7C3AED",
    ok: "#34D399",
    warn: "#FBBF24",
  },
  grad: "linear-gradient(120deg, #00E0C6, #38BDF8 45%, #FF2E9A)",
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
export const SCENES = [265, 327, 261, 258, 299, 258, 198, 402, 363, 272, 284, 392, 253, 253, 258] as const;
export const STARTS = SCENES.reduce<number[]>((a, d, i) => [...a, i === 0 ? 0 : a[i - 1] + SCENES[i - 1]], []);
export const TOTAL = SCENES.reduce((a, b) => a + b, 0);
