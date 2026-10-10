import React from "react";
import { AbsoluteFill, interpolate, spring, staticFile, useCurrentFrame, useVideoConfig } from "remotion";

if (typeof document !== "undefined") {
  [["Anton", "Anton.ttf", "400"], ["Montserrat", "Montserrat.ttf", "100 900"]].forEach(([fam, file, w]) => {
    const ff = new FontFace(fam, `url(${staticFile(`fonts/${file}`)})`, { weight: w });
    ff.load().then((f) => document.fonts.add(f));
  });
}

export const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;
export const rnd = (n: number) => { const v = Math.sin(n * 12.9898 + 78.233) * 43758.5453; return v - Math.floor(v); };
export const easeOut = (x: number) => 1 - Math.pow(1 - Math.min(1, Math.max(0, x)), 4);
export const easeInOut = (x: number) => { const t = Math.min(1, Math.max(0, x)); return t < 0.5 ? 8 * t * t * t * t : 1 - Math.pow(-2 * t + 2, 4) / 2; };

// Relógio em batidas. Troque o BPM para casar com outra música: todas as cenas seguem a grade.
export const useBeat = (bpm: number) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const b = (frame / fps) * (bpm / 60);
  const phase = b - Math.floor(b);
  return { frame, fps, b, phase, pulse: Math.exp(-phase * 6), f: (beats: number) => Math.round((beats * 60 * fps) / bpm) };
};

export const Pop: React.FC<{ at: number; b: number; children: React.ReactNode; style?: React.CSSProperties }> = ({ at, b, children, style }) => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const p = spring({ frame: frame - Math.round((at * 60 * fps) / b), fps, config: { damping: 11, stiffness: 190, mass: 0.6 } });
  return <div style={{ opacity: Math.min(1, p * 1.5), transform: `scale(${0.6 + 0.4 * p}) translateY(${(1 - p) * 40}px)`, ...style }}>{children}</div>;
};

export const Stars: React.FC<{ n?: number }> = ({ n = 90 }) => {
  const frame = useCurrentFrame();
  return (
    <AbsoluteFill>
      {Array.from({ length: n }).map((_, i) => (
        <div key={i} style={{ position: "absolute", left: rnd(i + 1) * 1080, top: ((rnd(i + 200) * 1920 + frame * (0.15 + rnd(i + 9) * 0.5)) % 1920), width: 2 + rnd(i + 50) * 3, height: 2 + rnd(i + 50) * 3, borderRadius: "50%", background: "#fff", opacity: 0.25 + 0.6 * Math.abs(Math.sin(frame / (20 + (i % 9)) + i)) }} />
      ))}
    </AbsoluteFill>
  );
};

export const Grain: React.FC = () => {
  const frame = useCurrentFrame();
  const noise = `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='220' height='220'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='2'/%3E%3C/filter%3E%3Crect width='220' height='220' filter='url(%23n)' opacity='0.5'/%3E%3C/svg%3E")`;
  return <AbsoluteFill style={{ pointerEvents: "none", backgroundImage: noise, backgroundSize: "220px", backgroundPosition: `${(frame * 7) % 220}px ${(frame * 13) % 220}px`, opacity: 0.06, mixBlendMode: "overlay" }} />;
};

export const Title: React.FC<{ children: React.ReactNode; size?: number; color?: string }> = ({ children, size = 120, color = "#fff" }) => (
  <div style={{ fontFamily: "Anton, Impact, sans-serif", fontSize: size, lineHeight: 1.02, textAlign: "center", color, textTransform: "uppercase", textShadow: "0 6px 40px rgba(0,0,0,0.8)" }}>{children}</div>
);
export const Small: React.FC<{ children: React.ReactNode; color?: string; size?: number }> = ({ children, color = "#c8d0e0", size = 38 }) => (
  <div style={{ fontFamily: "Montserrat, sans-serif", fontWeight: 700, fontSize: size, textAlign: "center", color, letterSpacing: "0.04em" }}>{children}</div>
);
