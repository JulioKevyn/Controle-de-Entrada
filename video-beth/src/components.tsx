import React from "react";
import { AbsoluteFill, Img, interpolate, spring, staticFile, useCurrentFrame, useVideoConfig } from "remotion";
import { theme } from "./theme";

if (typeof document !== "undefined") {
  [["Playfair", "Playfair.ttf", "400 900", "normal"], ["Montserrat", "Montserrat.ttf", "100 900", "normal"], ["Cormorant", "CormorantItalic.ttf", "300 700", "italic"]].forEach(([fam, file, w, st]) => {
    const ff = new FontFace(fam, `url(${staticFile(`fonts/${file}`)})`, { weight: w, style: st });
    ff.load().then((f) => document.fonts.add(f));
  });
}

export const C = theme.colors;
export const F = theme.fonts;
export const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;

export const useSpring = (delay = 0, cfg: keyof typeof theme.spring = "smooth") => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  return spring({ frame: frame - delay, fps, config: theme.spring[cfg] });
};

export const Entrance: React.FC<{ delay?: number; y?: number; x?: number; cfg?: keyof typeof theme.spring; style?: React.CSSProperties; children: React.ReactNode }> = ({ delay = 0, y = 40, x = 0, cfg = "smooth", style, children }) => {
  const p = useSpring(delay, cfg);
  return (
    <div style={{ opacity: Math.min(1, p), transform: `translate(${interpolate(p, [0, 1], [x, 0])}px, ${interpolate(p, [0, 1], [y, 0])}px) scale(${interpolate(p, [0, 1], [0.94, 1])})`, ...style }}>{children}</div>
  );
};

export const WordReveal: React.FC<{ text: string; delay?: number; per?: number; size?: number; highlight?: string[]; italic?: boolean; weight?: number }> = ({ text, delay = 0, per = 4, size = 90, highlight = [], italic, weight = 700 }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  return (
    <div style={{ display: "flex", flexWrap: "wrap", columnGap: size * 0.26, justifyContent: "center", textAlign: "center", fontFamily: italic ? F.script : F.display, fontStyle: italic ? "italic" : "normal", fontWeight: weight, fontSize: size, lineHeight: 1.12, color: C.text, textShadow: "0 4px 30px rgba(0,0,0,0.7)" }}>
      {text.split(" ").map((w, i) => {
        const p = spring({ frame: frame - delay - i * per, fps, config: theme.spring.snappy });
        const hl = highlight.includes(w);
        return (
          <span key={i} style={{ display: "inline-block", opacity: Math.min(1, p), transform: `translateY(${interpolate(p, [0, 1], [34, 0])}px)`, color: hl ? C.gold : undefined }}>{w}</span>
        );
      })}
    </div>
  );
};

// Foto com borda dourada e movimento lento (ken burns)
export const Photo: React.FC<{ src: string; w: number; h: number; rot?: number; delay?: number; x?: number; y?: number; drift?: number; style?: React.CSSProperties }> = ({ src, w, h, rot = 0, delay = 0, x = 60, y = 80, drift = 1, style }) => {
  const frame = useCurrentFrame();
  const p = useSpring(delay, "smooth");
  const z = 1.04 + (frame / 300) * 0.08 * drift;
  return (
    <div style={{ width: w, height: h, borderRadius: 14, overflow: "hidden", border: `3px solid ${C.gold}`, boxShadow: "0 40px 90px -20px rgba(0,0,0,0.85)", opacity: Math.min(1, p), transform: `translate(${interpolate(p, [0, 1], [x, 0])}px, ${interpolate(p, [0, 1], [y, 0])}px) rotate(${interpolate(p, [0, 1], [rot * 2.4, rot])}deg)`, ...style }}>
      <Img src={staticFile(`img/${src}.jpg`)} style={{ width: "100%", height: "100%", objectFit: "cover", transform: `scale(${z})` }} />
    </div>
  );
};

export const BgBlur: React.FC<{ src: string }> = ({ src }) => {
  const frame = useCurrentFrame();
  return (
    <AbsoluteFill style={{ background: C.bg, overflow: "hidden" }}>
      <Img src={staticFile(`img/${src}.jpg`)} style={{ position: "absolute", inset: -120, width: "calc(100% + 240px)", height: "calc(100% + 240px)", objectFit: "cover", filter: "blur(46px) saturate(1.2)", opacity: 0.55, transform: `scale(${1.1 + frame / 2400})` }} />
      <AbsoluteFill style={{ background: "linear-gradient(180deg, rgba(18,6,10,0.55), rgba(18,6,10,0.88))" }} />
    </AbsoluteFill>
  );
};

export const Grain: React.FC = () => {
  const frame = useCurrentFrame();
  const noise = `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='220' height='220'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='2'/%3E%3C/filter%3E%3Crect width='220' height='220' filter='url(%23n)' opacity='0.5'/%3E%3C/svg%3E")`;
  return <AbsoluteFill style={{ pointerEvents: "none", backgroundImage: noise, backgroundSize: "220px", backgroundPosition: `${(frame * 7) % 220}px ${(frame * 13) % 220}px`, opacity: 0.06, mixBlendMode: "overlay" }} />;
};
export const Vignette: React.FC = () => (
  <AbsoluteFill style={{ pointerEvents: "none", background: "radial-gradient(ellipse at center, transparent 50%, rgba(0,0,0,0.5) 100%)" }} />
);

export const Scene: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();
  const zoom = interpolate(frame, [0, 22], [1.06, 1], { ...clamp, easing: theme.ease.out });
  const t: [number, number] = [durationInFrames - 10, durationInFrames - 1];
  const flash = interpolate(frame, [0, 3, 14], [0, 0.5, 0], clamp);
  return (
    <AbsoluteFill style={{ opacity: interpolate(frame, t, [1, 0], clamp), transform: `scale(${zoom})` }}>
      {children}
      <AbsoluteFill style={{ pointerEvents: "none", background: C.gold, opacity: flash * 0.35 }} />
    </AbsoluteFill>
  );
};

// Brilhos dourados (bokeh) subindo
export const Sparkles: React.FC = () => {
  const frame = useCurrentFrame();
  const { width: w, height: h } = useVideoConfig();
  return (
    <AbsoluteFill style={{ pointerEvents: "none" }}>
      {Array.from({ length: 30 }).map((_, i) => {
        const rnd = (n: number) => { const v = Math.sin(n * 12.9898 + 78.233) * 43758.5453; return v - Math.floor(v); };
        const x = rnd(i + 1) * (w - 20);
        const speed = 0.3 + ((i * 37) % 10) / 16;
        const y = h + 80 - ((frame * speed * 1.5 + rnd(i + 50) * (h + 180)) % (h + 180));
        const tw = 0.3 + 0.7 * Math.abs(Math.sin(frame / (16 + (i % 7)) + i));
        const size = 5 + (i % 4) * 4;
        return <div key={i} style={{ position: "absolute", left: x, top: y, width: size, height: size, borderRadius: "50%", background: C.gold, opacity: tw * 0.5, boxShadow: `0 0 ${size * 4}px ${C.gold}`, filter: i % 3 === 0 ? "blur(2px)" : undefined }} />;
      })}
    </AbsoluteFill>
  );
};

export const Captions: React.FC<{ chunks: { t: string; a: number; b: number }[] }> = ({ chunks }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const t = frame / fps;
  const idx = chunks.findIndex((c) => t >= c.a && t < c.b);
  if (idx < 0) return null;
  const c = chunks[idx];
  const p = spring({ frame: frame - Math.round(c.a * fps), fps, config: theme.spring.snappy });
  return (
    <div style={{ position: "absolute", left: 0, right: 0, bottom: 150, display: "flex", justifyContent: "center", opacity: Math.min(1, p), transform: `translateY(${interpolate(p, [0, 1], [18, 0])}px) scale(${interpolate(p, [0, 1], [0.94, 1])})` }}>
      <div style={{ padding: "14px 36px", borderRadius: 22, background: "rgba(18,6,10,0.72)", border: `1px solid ${C.line}`, fontFamily: F.body, fontWeight: 600, fontSize: 42, color: C.text, maxWidth: 960, textAlign: "center", lineHeight: 1.25 }}>{c.t}</div>
    </div>
  );
};

export const Chip: React.FC<{ children: React.ReactNode; delay?: number }> = ({ children, delay = 0 }) => (
  <Entrance delay={delay} y={24}>
    <span style={{ display: "inline-block", padding: "12px 28px", borderRadius: 999, background: "rgba(217,178,111,0.12)", border: `1.5px solid ${C.gold}`, color: C.gold, fontFamily: F.body, fontWeight: 600, fontSize: 32, letterSpacing: "0.05em" }}>{children}</span>
  </Entrance>
);
