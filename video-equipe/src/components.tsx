import React from "react";
import { AbsoluteFill, interpolate, interpolateColors, spring, staticFile, useCurrentFrame, useVideoConfig } from "remotion";
import { theme } from "./theme";

const FONTS: [string, string, number][] = [
  ...[400, 600, 700, 800].map((w) => ["Harabara", "Harabara.ttf", w] as [string, string, number]),
  ...[400, 600, 700, 800].map((w) => ["Source Sans Pro", `SourceSans-${Math.min(w, 700)}.ttf`, w] as [string, string, number]),
];
if (typeof document !== "undefined") {
  FONTS.forEach(([fam, file, w]) => {
    const ff = new FontFace(fam, `url(${staticFile(`fonts/${file}`)})`, { weight: String(w) });
    ff.load().then((f) => document.fonts.add(f));
  });
}

export const useV = () => { const { width, height } = useVideoConfig(); return { V: height > width, w: width, h: height }; };
export const C = theme.colors;
export const F = theme.fonts;
export const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;
export const panel: React.CSSProperties = {
  background: C.panel, border: `1px solid ${C.line}`, borderRadius: 26,
  boxShadow: "0 30px 70px -24px rgba(0,0,0,0.7)", backdropFilter: "blur(12px)",
};

export const useSpring = (delay = 0, cfg: keyof typeof theme.spring = "smooth") => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  return spring({ frame: frame - delay, fps, config: theme.spring[cfg] });
};
export const breathe = (frame: number, amp = 0.015, speed = 22) => 1 + Math.sin(frame / speed) * amp;
export const float = (frame: number, amp = 3, speed = 30) => Math.sin(frame / speed) * amp;

export const Entrance: React.FC<{ delay?: number; y?: number; x?: number; cfg?: keyof typeof theme.spring; style?: React.CSSProperties; children: React.ReactNode }> = ({ delay = 0, y = 40, x = 0, cfg = "smooth", style, children }) => {
  const p = useSpring(delay, cfg);
  return (
    <div style={{ opacity: Math.min(1, p), transform: `translate(${interpolate(p, [0, 1], [x, 0])}px, ${interpolate(p, [0, 1], [y, 0])}px) scale(${interpolate(p, [0, 1], [0.94, 1])})`, ...style }}>{children}</div>
  );
};

export const GradText: React.FC<{ children: React.ReactNode; style?: React.CSSProperties }> = ({ children, style }) => (
  <span style={{ backgroundImage: theme.grad, WebkitBackgroundClip: "text", backgroundClip: "text", color: "transparent", ...style }}>{children}</span>
);

export const WordReveal: React.FC<{ text: string; delay?: number; per?: number; size?: number; highlight?: string[]; align?: "flex-start" | "center"; weight?: number }> = ({ text, delay = 0, per = 3, size = 80, highlight = [], align = "flex-start", weight = 700 }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  return (
    <div style={{ display: "flex", flexWrap: "wrap", columnGap: size * 0.28, justifyContent: align, fontFamily: F.display, fontWeight: weight, fontSize: size, lineHeight: 1.15, letterSpacing: "-0.02em", color: C.text }}>
      {text.split(" ").map((w, i) => {
        const p = spring({ frame: frame - delay - i * per, fps, config: theme.spring.snappy });
        const hl = highlight.includes(w);
        return (
          <span key={i} style={{ display: "inline-block", opacity: Math.min(1, p), transform: `translateY(${interpolate(p, [0, 1], [30, 0])}px)` }}>
            {hl ? <GradText>{w}</GradText> : w}
          </span>
        );
      })}
    </div>
  );
};

export const BgMesh: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <AbsoluteFill style={{ background: C.bg }}>
      <div style={{ position: "absolute", width: 1200, height: 1200, borderRadius: "50%", top: -560, left: -300 + Math.sin(frame / 60) * 60, filter: "blur(80px)", background: `radial-gradient(circle, ${C.teal}2E, transparent 62%)` }} />
      <div style={{ position: "absolute", width: 1100, height: 1100, borderRadius: "50%", bottom: -520, right: -260 + Math.cos(frame / 75) * 50, filter: "blur(90px)", background: `radial-gradient(circle, ${C.pink}26, transparent 62%)` }} />
      <div style={{ position: "absolute", width: 900, height: 900, borderRadius: "50%", top: 200, right: 300, filter: "blur(110px)", background: `radial-gradient(circle, ${C.violet}1F, transparent 65%)` }} />
      <AbsoluteFill style={{ backgroundImage: "linear-gradient(rgba(255,255,255,0.03) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.03) 1px, transparent 1px)", backgroundSize: "80px 80px", backgroundPosition: `${-frame * 0.35}px ${-frame * 0.18}px` }} />
    </AbsoluteFill>
  );
};
export const Grain: React.FC = () => {
  const frame = useCurrentFrame();
  const noise = `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='220' height='220'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='2'/%3E%3C/filter%3E%3Crect width='220' height='220' filter='url(%23n)' opacity='0.5'/%3E%3C/svg%3E")`;
  return <AbsoluteFill style={{ pointerEvents: "none", backgroundImage: noise, backgroundSize: "220px", backgroundPosition: `${(frame * 7) % 220}px ${(frame * 13) % 220}px`, opacity: 0.05, mixBlendMode: "overlay" }} />;
};
export const Vignette: React.FC = () => (
  <AbsoluteFill style={{ pointerEvents: "none", background: "radial-gradient(ellipse at center, transparent 55%, rgba(0,0,0,0.4) 100%)" }} />
);

export const Scene: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();
  const t: [number, number] = [durationInFrames - 10, durationInFrames - 1];
  const zoom = interpolate(frame, [0, 22], [1.07, 1], { ...clamp, easing: theme.ease.out });
  const sweep = interpolate(frame, [0, 16], [-40, 140], { ...clamp, easing: theme.ease.inOut });
  const sweepO = interpolate(frame, [0, 4, 16], [0, 1, 0], clamp);
  return (
    <AbsoluteFill style={{ opacity: interpolate(frame, t, [1, 0], clamp), transform: `translateY(${interpolate(frame, t, [0, -30], { ...clamp, easing: theme.ease.in })}px) scale(${zoom})` }}>
      {children}
      <AbsoluteFill style={{ pointerEvents: "none", overflow: "hidden", opacity: sweepO }}>
        <div style={{ position: "absolute", top: -200, bottom: -200, width: 420, left: `${sweep}%`, transform: "skewX(-18deg)", background: `linear-gradient(90deg, transparent, ${C.teal}55, ${C.pink}44, transparent)`, filter: "blur(18px)" }} />
      </AbsoluteFill>
    </AbsoluteFill>
  );
};

// Partículas / brilhos flutuando (determinístico)
export const Sparkles: React.FC = () => {
  const frame = useCurrentFrame();
  const { w, h } = useV();
  return (
    <AbsoluteFill style={{ pointerEvents: "none" }}>
      {Array.from({ length: 34 }).map((_, i) => {
        const rnd = (n: number) => { const v = Math.sin(n * 12.9898 + 78.233) * 43758.5453; return v - Math.floor(v); };
        const x = rnd(i + 1) * (w - 20);
        const speed = 0.35 + ((i * 37) % 10) / 14;
        const y = h + 80 - ((frame * speed * 1.6 + rnd(i + 50) * (h + 180)) % (h + 180));
        const tw = 0.35 + 0.65 * Math.abs(Math.sin(frame / (14 + (i % 7)) + i));
        const size = 3 + (i % 4) * 2;
        const col = [C.teal, C.sky, C.pink, C.violet][i % 4];
        return <div key={i} style={{ position: "absolute", left: x, top: y, width: size, height: size, borderRadius: "50%", background: col, opacity: tw * 0.55, boxShadow: `0 0 ${size * 4}px ${col}` }} />;
      })}
    </AbsoluteFill>
  );
};

// Legendas (timing aproximado da fala)
export const Captions: React.FC<{ chunks: { t: string; a: number; b: number }[] }> = ({ chunks }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const { V } = useV();
  const t = frame / fps;
  const idx = chunks.findIndex((c) => t >= c.a && t < c.b);
  if (idx < 0) return null;
  const c = chunks[idx];
  const p = spring({ frame: frame - Math.round(c.a * fps), fps, config: theme.spring.snappy });
  return (
    <div style={{ position: "absolute", left: 0, right: 0, bottom: V ? 120 : 34, display: "flex", justifyContent: "center", opacity: Math.min(1, p), transform: `translateY(${interpolate(p, [0, 1], [18, 0])}px) scale(${interpolate(p, [0, 1], [0.94, 1])})` }}>
      <div style={{ padding: "12px 34px", borderRadius: 22, background: "rgba(26,22,23,0.78)", border: `1px solid ${C.line}`, backdropFilter: "blur(10px)", fontFamily: F.display, fontWeight: 700, fontSize: V ? 40 : 40, color: C.text, letterSpacing: "-0.01em", textShadow: "0 2px 18px rgba(0,0,0,0.6)", whiteSpace: V ? "normal" : "nowrap", maxWidth: V ? 980 : undefined, textAlign: "center", lineHeight: 1.2 }}>{c.t}</div>
    </div>
  );
};

export const Counter: React.FC<{ to: number; delay?: number; style?: React.CSSProperties }> = ({ to, delay = 0, style }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const p = spring({ frame: frame - delay, fps, config: { damping: 30, stiffness: 60 } });
  return <span style={{ fontVariantNumeric: "tabular-nums", ...style }}>{Math.round(interpolate(p, [0, 1], [0, to])).toLocaleString("pt-BR")}</span>;
};

export const Chip: React.FC<{ color: string; children: React.ReactNode; size?: number }> = ({ color, children, size = 22 }) => (
  <span style={{ display: "inline-block", padding: "7px 18px", borderRadius: 999, background: `${color}1F`, border: `1px solid ${color}80`, color, fontFamily: F.body, fontWeight: 700, fontSize: size, whiteSpace: "nowrap" }}>{children}</span>
);
export const Label: React.FC<{ children: React.ReactNode; delay?: number; color?: string }> = ({ children, delay = 0, color = C.teal }) => (
  <Entrance delay={delay} y={20}>
    <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 24, letterSpacing: "0.24em", textTransform: "uppercase", color }}>{children}</div>
  </Entrance>
);

export const AiBadge: React.FC<{ delay?: number; big?: boolean }> = ({ delay = 0, big }) => {
  const p = useSpring(delay, "bouncy");
  return (
    <div style={{ opacity: Math.min(1, p), transform: `scale(${p})`, transformOrigin: "left center", display: "inline-flex", alignItems: "center", columnGap: 12, padding: big ? "12px 28px" : "8px 20px", borderRadius: 999, background: `linear-gradient(120deg, ${C.violet}, ${C.pink})`, color: "#fff", fontFamily: F.body, fontWeight: 800, fontSize: big ? 30 : 22, letterSpacing: "0.04em", boxShadow: `0 10px 40px -8px ${C.pink}99` }}>
      <span style={{ fontSize: big ? 34 : 24 }}>✦</span> Inteligência Artificial
    </div>
  );
};
