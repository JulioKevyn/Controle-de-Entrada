import React from "react";
import { AbsoluteFill, Img, interpolate, spring, staticFile, useCurrentFrame, useVideoConfig } from "remotion";
import { theme } from "./theme";


const FONTS: [string, number][] = [["Sora", 600], ["Sora", 700], ["Sora", 800], ["Inter", 400], ["Inter", 500], ["Inter", 600]];
if (typeof document !== "undefined") {
  FONTS.forEach(([fam, w]) => {
    const ff = new FontFace(fam, `url(${staticFile(`fonts/${fam}-${w}.ttf`)})`, { weight: String(w) });
    ff.load().then((f) => document.fonts.add(f));
  });
}

const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;

export const useSpring = (delay = 0, cfg: keyof typeof theme.spring = "smooth") => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  return spring({ frame: frame - delay, fps, config: theme.spring[cfg] });
};

export const Entrance: React.FC<{
  delay?: number;
  y?: number;
  x?: number;
  cfg?: keyof typeof theme.spring;
  style?: React.CSSProperties;
  children: React.ReactNode;
}> = ({ delay = 0, y = 40, x = 0, cfg = "smooth", style, children }) => {
  const p = useSpring(delay, cfg);
  return (
    <div
      style={{
        opacity: p,
        transform: `translate(${interpolate(p, [0, 1], [x, 0])}px, ${interpolate(p, [0, 1], [y, 0])}px) scale(${interpolate(p, [0, 1], [0.94, 1])})`,
        ...style,
      }}
    >
      {children}
    </div>
  );
};

export const WordReveal: React.FC<{
  text: string;
  delay?: number;
  per?: number;
  size?: number;
  highlight?: string[];
  align?: "flex-start" | "center";
  weight?: number;
}> = ({ text, delay = 0, per = 3, size = 96, highlight = [], align = "flex-start", weight = 700 }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  return (
    <div
      style={{
        display: "flex",
        flexWrap: "wrap",
        columnGap: size * 0.26,
        justifyContent: align,
        fontFamily: theme.fonts.display,
        fontWeight: weight,
        fontSize: size,
        lineHeight: 1.08,
        letterSpacing: "-0.03em",
        color: theme.colors.text,
      }}
    >
      {text.split(" ").map((w, i) => {
        const p = spring({ frame: frame - delay - i * per, fps, config: theme.spring.snappy });
        const hl = highlight.includes(w);
        return (
          <span
            key={i}
            style={{
              display: "inline-block",
              opacity: p,
              transform: `translateY(${interpolate(p, [0, 1], [34, 0])}px)`,
              color: hl ? theme.colors.primary : undefined,
              textShadow: hl ? `0 0 50px ${theme.colors.glow}` : undefined,
            }}
          >
            {w}
          </span>
        );
      })}
    </div>
  );
};

// Fundo vivo (camada 1)
export const BgMesh: React.FC = () => {
  const frame = useCurrentFrame();
  const d1 = Math.sin(frame / 55) * 60;
  const d2 = Math.cos(frame / 70) * 50;
  return (
    <AbsoluteFill style={{ background: theme.colors.bg }}>
      <div style={{ position: "absolute", width: 1300, height: 1300, borderRadius: "50%", top: -600, left: -350 + d1, filter: "blur(60px)", background: `radial-gradient(circle, ${theme.colors.primary}30, transparent 62%)` }} />
      <div style={{ position: "absolute", width: 1000, height: 1000, borderRadius: "50%", bottom: -500, right: -300 - d2, filter: "blur(80px)", background: `radial-gradient(circle, ${theme.colors.accent}18, transparent 65%)` }} />
      <AbsoluteFill style={{ backgroundImage: "linear-gradient(rgba(255,255,255,0.025) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.025) 1px, transparent 1px)", backgroundSize: "80px 80px", backgroundPosition: `${-frame * 0.4}px ${-frame * 0.2}px` }} />
    </AbsoluteFill>
  );
};

export const Grade: React.FC = () => (
  <AbsoluteFill style={{ pointerEvents: "none" }}>
    <AbsoluteFill style={{ backgroundColor: theme.colors.primary, mixBlendMode: "soft-light", opacity: 0.18 }} />
    <AbsoluteFill style={{ background: "linear-gradient(180deg, rgba(0,0,0,0.12), transparent 28%, transparent 72%, rgba(0,0,0,0.22))" }} />
  </AbsoluteFill>
);

export const Grain: React.FC = () => {
  const frame = useCurrentFrame();
  const noise = `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='220' height='220'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='2'/%3E%3C/filter%3E%3Crect width='220' height='220' filter='url(%23n)' opacity='0.5'/%3E%3C/svg%3E")`;
  return (
    <AbsoluteFill style={{ pointerEvents: "none", backgroundImage: noise, backgroundSize: "220px", backgroundPosition: `${(frame * 7) % 220}px ${(frame * 13) % 220}px`, opacity: 0.05, mixBlendMode: "overlay" }} />
  );
};

export const Vignette: React.FC = () => (
  <AbsoluteFill style={{ pointerEvents: "none", background: "radial-gradient(ellipse at center, transparent 56%, rgba(0,0,0,0.35) 100%)" }} />
);

// Saída rápida da cena
export const Scene: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();
  const t: [number, number] = [durationInFrames - 10, durationInFrames - 1];
  const o = interpolate(frame, t, [1, 0], clamp);
  const y = interpolate(frame, t, [0, -30], { ...clamp, easing: theme.ease.in });
  return <AbsoluteFill style={{ opacity: o, transform: `translateY(${y}px)` }}>{children}</AbsoluteFill>;
};

export const Counter: React.FC<{ to: number; delay?: number; style?: React.CSSProperties }> = ({ to, delay = 0, style }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const p = spring({ frame: frame - delay, fps, config: { damping: 30, stiffness: 60 } });
  return <span style={{ fontVariantNumeric: "tabular-nums", ...style }}>{Math.round(interpolate(p, [0, 1], [0, to]))}</span>;
};

export const breathe = (frame: number, amp = 0.015, speed = 22) => 1 + Math.sin(frame / speed) * amp;
export const float = (frame: number, amp = 3, speed = 30) => Math.sin(frame / speed) * amp;

// Logo da empresa (public/logo.png)
export const Mark: React.FC<{ size?: number; delay?: number }> = ({ size = 160, delay = 0 }) => {
  const frame = useCurrentFrame();
  const sc = useSpring(delay, "bouncy");
  return (
    <div style={{ opacity: sc, transform: `scale(${interpolate(sc, [0, 1], [0.7, 1]) * breathe(frame, 0.012)})`, background: "#FFFFFF", borderRadius: size * 0.16, padding: size * 0.14, boxShadow: `0 30px 80px -20px rgba(0,0,0,0.7), 0 0 ${size * 0.5}px ${theme.colors.glow}` }}>
      <Img src={staticFile("logo.png")} style={{ height: size, width: "auto", maxWidth: size * 3.2, objectFit: "contain", display: "block" }} />
    </div>
  );
};

export const Chip: React.FC<{ color: string; children: React.ReactNode }> = ({ color, children }) => (
  <span style={{ display: "inline-block", padding: "6px 16px", borderRadius: 999, background: `${color}22`, border: `1px solid ${color}66`, color, fontFamily: theme.fonts.body, fontWeight: 600, fontSize: 22 }}>{children}</span>
);

export const Label: React.FC<{ children: React.ReactNode; delay?: number }> = ({ children, delay = 0 }) => (
  <Entrance delay={delay} y={20}>
    <div style={{ fontFamily: theme.fonts.body, fontWeight: 600, fontSize: 26, letterSpacing: "0.22em", textTransform: "uppercase", color: theme.colors.primary }}>{children}</div>
  </Entrance>
);
