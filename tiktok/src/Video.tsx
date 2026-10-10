import React from "react";
import { AbsoluteFill, Audio, interpolate, staticFile, useCurrentFrame, useVideoConfig } from "remotion";
import type { VideoConfig } from "./types";

if (typeof document !== "undefined") {
  const ff = new FontFace("Anton", `url(${staticFile("fonts/Anton.ttf")})`);
  ff.load().then((f) => document.fonts.add(f));
}

const rnd = (n: number) => { const v = Math.sin(n * 12.9898 + 78.233) * 43758.5453; return v - Math.floor(v); };
// tempo (s) desde o último evento da lista (ou Infinity)
const since = (list: number[], t: number) => {
  let last = -1;
  for (const x of list) { if (x <= t) last = x; else break; }
  return last < 0 ? Infinity : t - last;
};

export const Video: React.FC<{ cfg: VideoConfig }> = ({ cfg }) => {
  const frame = useCurrentFrame();
  const { fps, width, height } = useVideoConfig();
  const t = frame / fps;
  const beat = 60 / cfg.beats.bpm;
  const bt = t / beat; // tempo em batidas
  const kick = Math.exp(-since(cfg.beats.kick, t) * 9);
  const snare = Math.exp(-since(cfg.beats.snare, t) * 14);
  const dropped = bt >= cfg.dropBeat;
  const line = cfg.lines.find((l) => bt >= l.b && bt < l.b + l.d) ?? cfg.lines[cfg.lines.length - 1];
  const lf = (bt - line.b) * beat * fps; // frames desde o início da linha
  const words = line.t.split(" ");
  const longest = Math.max(...words.map((w) => w.length));
  const lines = line.t.length > 9 && words.length > 1 ? words : [line.t];
  const maxLen = Math.max(...lines.map((w) => w.length));
  const size = Math.min(330, 900 / (maxLen * 0.47), 330);
  const scale = 1 + 0.13 * kick + (lf < 3 ? 0.08 : 0);
  const shake = kick * 14;
  const sx = (rnd(frame) - 0.5) * shake, sy = (rnd(frame + 99) - 0.5) * shake;
  const glitch = lf < 4 ? (4 - lf) * 7 : 0;
  const glow = dropped ? 0.55 + 0.35 * kick : 0.18 + 0.2 * kick;
  const style: React.CSSProperties = { fontFamily: "Anton, Impact, sans-serif", fontSize: size, lineHeight: 1.02, textAlign: "center", color: "#fff", letterSpacing: "0.01em", textTransform: "uppercase", position: "absolute", left: 0, right: 0 };
  const inner = lines.map((w, i) => <div key={i}>{w}</div>);
  return (
    <AbsoluteFill style={{ background: "#070707", overflow: "hidden" }}>
      <AbsoluteFill style={{ background: `radial-gradient(ellipse at 50% 45%, ${cfg.accent}${Math.round(glow * 255).toString(16).padStart(2, "0")} 0%, transparent 62%)` }} />
      <AbsoluteFill style={{ backgroundImage: "repeating-linear-gradient(115deg, rgba(255,255,255,0.035) 0 2px, transparent 2px 38px)", backgroundPosition: `${frame * 3}px 0` }} />
      <AbsoluteFill style={{ display: "flex", alignItems: "center", justifyContent: "center", paddingBottom: 220, transform: `translate(${sx}px, ${sy}px) scale(${scale})` }}>
        <div style={{ position: "relative", width }}>
          {glitch > 0 && <div style={{ ...style, position: "absolute", top: 0, color: "#ff2a2a", transform: `translateX(${-glitch}px)`, mixBlendMode: "screen", opacity: 0.9 }}>{inner}</div>}
          {glitch > 0 && <div style={{ ...style, position: "absolute", top: 0, color: "#22d3ee", transform: `translateX(${glitch}px)`, mixBlendMode: "screen", opacity: 0.9 }}>{inner}</div>}
          <div style={{ ...style, position: "relative", textShadow: dropped ? `0 0 40px ${cfg.accent}` : "0 0 30px rgba(255,255,255,0.15)" }}>{inner}</div>
        </div>
      </AbsoluteFill>
      <AbsoluteFill style={{ background: "#fff", opacity: snare * 0.32, pointerEvents: "none" }} />
      <AbsoluteFill style={{ background: cfg.accent, opacity: interpolate(bt, [cfg.dropBeat, cfg.dropBeat + 0.5], [0.6, 0], { extrapolateLeft: "clamp", extrapolateRight: "clamp" }) * (dropped ? 1 : 0), pointerEvents: "none" }} />
      <AbsoluteFill style={{ background: "radial-gradient(ellipse at center, transparent 55%, rgba(0,0,0,0.7) 100%)", pointerEvents: "none" }} />
      <Audio src={staticFile(cfg.audio)} volume={1} />
    </AbsoluteFill>
  );
};
