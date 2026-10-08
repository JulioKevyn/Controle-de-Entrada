import { Easing } from "remotion";

export const theme = {
  colors: {
    bg: "#12060A",
    panel: "rgba(255,255,255,0.05)",
    line: "rgba(217,178,111,0.35)",
    text: "#F7EFE6",
    muted: "#C9B9A8",
    gold: "#D9B26F",
    wine: "#7A1230",
    rose: "#E58AA0",
  },
  fonts: { display: "Playfair, serif", body: "Montserrat, sans-serif", script: "Cormorant, serif" },
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
export const SCENES = [186, 279, 210, 155] as const;
export const STARTS = SCENES.reduce<number[]>((a, d, i) => [...a, i === 0 ? 0 : a[i - 1] + SCENES[i - 1]], []);
export const TOTAL = SCENES.reduce((a, b) => a + b, 0);
