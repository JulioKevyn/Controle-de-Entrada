import { Easing } from "remotion";

export const theme = {
  colors: {
    bg: "#0B0A09",
    bgAlt: "#17130F",
    card: "#1C1713",
    line: "rgba(255,255,255,0.08)",
    primary: "#EE7500",
    accent: "#38BDF8",
    ok: "#34D399",
    text: "#F5F1EC",
    textDim: "#A8A29A",
    glow: "rgba(238,117,0,0.45)",
  },
  fonts: { display: "Sora, sans-serif", body: "Inter, sans-serif" },
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
export const SCENES = [212, 217, 204, 230, 156, 227, 135, 128] as const;
export const STARTS = SCENES.reduce<number[]>((a, d, i) => [...a, i === 0 ? 0 : a[i - 1] + SCENES[i - 1]], []);
export const TOTAL = SCENES.reduce((a, b) => a + b, 0);
