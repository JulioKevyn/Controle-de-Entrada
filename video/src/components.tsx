import React from "react";
import { AbsoluteFill, Img, interpolate, spring, staticFile, useCurrentFrame, useVideoConfig } from "remotion";
import { theme } from "./theme";


const FONTS: [string, string, number][] = [
  ...[400, 600, 700, 800].map((w) => ["Harabara", "Harabara.ttf", w] as [string, string, number]),
  ...[400, 600, 700].map((w) => ["Source Sans Pro", `SourceSans-${w}.ttf`, w] as [string, string, number]),
];
if (typeof document !== "undefined") {
  FONTS.forEach(([fam, file, w]) => {
    const ff = new FontFace(fam, `url(${staticFile(`fonts/${file}`)})`, { weight: String(w) });
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
  color?: string;
  hlColor?: string;
}> = ({ text, delay = 0, per = 3, size = 96, highlight = [], align = "flex-start", weight = 700, color = theme.colors.text, hlColor = theme.colors.primary }) => {
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
        letterSpacing: "0.015em",
        color,
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
              color: hl ? hlColor : undefined,
            }}
          >
            {w}
          </span>
        );
      })}
    </div>
  );
};

// Fundo vivo (camada 1): degradê claro da marca + grafismo de elipses
export const BgMesh: React.FC = () => {
  const frame = useCurrentFrame();
  const rot = frame * 0.12;
  return (
    <AbsoluteFill style={{ background: `linear-gradient(135deg, #E4E4E4 0%, #F4F4F4 45%, #FFFFFF 100%)` }}>
      <div style={{ position: "absolute", width: 1300, height: 1300, borderRadius: "50%", top: -650, left: -400 + Math.sin(frame / 55) * 50, filter: "blur(90px)", background: `radial-gradient(circle, ${theme.colors.primary}26, transparent 62%)` }} />
      <svg width={1100} height={1100} viewBox="-550 -550 1100 1100" style={{ position: "absolute", right: -380, bottom: -380, opacity: 0.16 }}>
        <g transform={`rotate(${rot})`} fill="none" stroke={theme.colors.primary} strokeWidth={2.2}>
          {Array.from({ length: 22 }).map((_, i) => (
            <ellipse key={i} cx={0} cy={0} rx={520 - i * 4} ry={520 - i * 21} transform={`rotate(${i * 2.2 - 25})`} />
          ))}
        </g>
      </svg>
    </AbsoluteFill>
  );
};

export const FooterBar: React.FC = () => (
  <div style={{ position: "absolute", left: 0, right: 0, bottom: 0, height: 30, background: `linear-gradient(90deg, ${theme.colors.primary}, ${theme.colors.primary2})` }} />
);

export const OrangeCover: React.FC = () => {
  const frame = useCurrentFrame();
  return <AbsoluteFill style={{ background: `linear-gradient(${100 + Math.sin(frame / 60) * 8}deg, ${theme.colors.primary} 0%, ${theme.colors.primary2} 100%)` }} />;
};

export const Grade: React.FC = () => (
  <AbsoluteFill style={{ pointerEvents: "none" }}>
    <AbsoluteFill style={{ background: "linear-gradient(180deg, rgba(35,31,32,0.04), transparent 28%, transparent 72%, rgba(35,31,32,0.06))" }} />
  </AbsoluteFill>
);

export const Grain: React.FC = () => {
  const frame = useCurrentFrame();
  const noise = `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='220' height='220'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='2'/%3E%3C/filter%3E%3Crect width='220' height='220' filter='url(%23n)' opacity='0.5'/%3E%3C/svg%3E")`;
  return (
    <AbsoluteFill style={{ pointerEvents: "none", backgroundImage: noise, backgroundSize: "220px", backgroundPosition: `${(frame * 7) % 220}px ${(frame * 13) % 220}px`, opacity: 0.05, mixBlendMode: "multiply" }} />
  );
};

export const Vignette: React.FC = () => (
  <AbsoluteFill style={{ pointerEvents: "none", background: "radial-gradient(ellipse at center, transparent 60%, rgba(35,31,32,0.10) 100%)" }} />
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
  return <span style={{ fontVariantNumeric: "tabular-nums", ...style }}>{Math.round(interpolate(p, [0, 1], [0, to])).toLocaleString("pt-BR")}</span>;
};

export const breathe = (frame: number, amp = 0.015, speed = 22) => 1 + Math.sin(frame / speed) * amp;
export const float = (frame: number, amp = 3, speed = 30) => Math.sin(frame / speed) * amp;

// Logo da empresa: "circle" (principal, colorido) ou "white" (secundário, sobre laranja)
export const Mark: React.FC<{ size?: number; delay?: number; variant?: "circle" | "white"; shadow?: boolean }> = ({ size = 160, delay = 0, variant = "circle", shadow = true }) => {
  const frame = useCurrentFrame();
  const sc = useSpring(delay, "bouncy");
  const src = variant === "circle" ? "logo.png" : "logo-white.png";
  return (
    <div style={{ opacity: Math.min(1, sc), transform: `scale(${interpolate(sc, [0, 1], [0.7, 1]) * breathe(frame, 0.012)})`, filter: shadow ? "drop-shadow(0 24px 40px rgba(35,31,32,0.25))" : undefined }}>
      <Img src={staticFile(src)} style={{ height: size, width: "auto", display: "block" }} />
    </div>
  );
};

export const Chip: React.FC<{ color: string; children: React.ReactNode }> = ({ color, children }) => (
  <span style={{ display: "inline-block", padding: "6px 16px", borderRadius: 999, background: `${color}1A`, border: `1px solid ${color}77`, color, fontFamily: theme.fonts.body, fontWeight: 600, fontSize: 22 }}>{children}</span>
);

export const Label: React.FC<{ children: React.ReactNode; delay?: number }> = ({ children, delay = 0 }) => (
  <Entrance delay={delay} y={20}>
    <div style={{ fontFamily: theme.fonts.body, fontWeight: 700, fontSize: 26, letterSpacing: "0.22em", textTransform: "uppercase", color: theme.colors.primary }}>{children}</div>
  </Entrance>
);
