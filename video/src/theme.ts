import { Easing } from "remotion";

export const theme = {
  colors: {
    // Brandbook Mundial Logistics 2024
    bg: "#F7F7F7",
    bgAlt: "#E6E6E6",
    card: "#FFFFFF",
    line: "rgba(35,31,32,0.10)",
    primary: "#EC6707",
    primary2: "#F39200",
    accent: "#F39200",
    ok: "#2F9E6E",
    text: "#231F20",
    textDim: "#6B6665",
    glow: "rgba(236,103,7,0.22)",
  },
  fonts: { display: "Harabara, sans-serif", body: "'Source Sans Pro', sans-serif" },
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

// frames por cena (narração + respiro)
export const SCENES = [193, 304, 253, 256, 594, 229, 358, 311, 409, 462, 143] as const;
export const STARTS = SCENES.reduce<number[]>((a, d, i) => [...a, i === 0 ? 0 : a[i - 1] + SCENES[i - 1]], []);
export const TOTAL = SCENES.reduce((a, b) => a + b, 0);
