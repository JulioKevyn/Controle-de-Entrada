import React from "react";
import { AbsoluteFill } from "remotion";
import { Grain, Vignette } from "./components";

const PAL = ["#ef4444", "#f59e0b", "#22c55e", "#3b82f6", "#a855f7", "#ec4899", "#14b8a6", "#f97316"];
const rnd = (n: number) => { const v = Math.sin(n * 12.9898 + 78.233) * 43758.5453; return v - Math.floor(v); };

const Product: React.FC<{ i: number; x: number }> = ({ i, x }) => {
  const t = Math.floor(rnd(i * 3 + 1) * 3), c = PAL[Math.floor(rnd(i + 7) * PAL.length)], h = 90 + rnd(i + 3) * 70, w = 64 + rnd(i + 5) * 40;
  if (t === 0) return (<div style={{ position: "absolute", left: x, bottom: 0, width: w, height: h, borderRadius: 8, background: c, boxShadow: "inset -10px 0 0 rgba(0,0,0,0.15)" }}><div style={{ margin: "26% 12%", height: "30%", background: "rgba(255,255,255,0.85)", borderRadius: 4 }} /></div>);
  if (t === 1) return (<div style={{ position: "absolute", left: x, bottom: 0, width: w * 0.8, height: h * 0.8, borderRadius: 10, background: "linear-gradient(90deg,#d1d5db,#f3f4f6 40%,#9ca3af)" }}><div style={{ margin: "30% 0", height: "36%", background: c }} /></div>);
  return (<div style={{ position: "absolute", left: x, bottom: 0, width: w * 0.7, height: h * 1.1 }}><div style={{ margin: "0 32%", height: "26%", background: c, borderRadius: "6px 6px 0 0" }} /><div style={{ height: "74%", borderRadius: 14, background: c, boxShadow: "inset -9px 0 0 rgba(0,0,0,0.18)" }}><div style={{ margin: "30% 14%", height: "28%", background: "rgba(255,255,255,0.8)", borderRadius: 4 }} /></div></div>);
};
const Shelf: React.FC<{ row: number; y: number }> = ({ row, y }) => {
  const items: React.ReactNode[] = []; let x = 10, i = row * 40;
  while (x < 1100) { const w = 70 + rnd(i) * 30; items.push(<Product key={i} i={i} x={x} />); x += w; i++; }
  return (
    <div style={{ position: "absolute", left: 0, right: 0, top: y, height: 190 }}>
      <div style={{ position: "absolute", left: 0, right: 0, bottom: 0, height: 190 }}>{items}</div>
      <div style={{ position: "absolute", left: 0, right: 0, bottom: -16, height: 22, background: "linear-gradient(180deg,#e5e7eb,#9ca3af)", boxShadow: "0 12px 30px rgba(0,0,0,0.5)" }} />
      <div style={{ position: "absolute", left: 120, bottom: -12, width: 54, height: 20, background: "#ffe600" }} />
      <div style={{ position: "absolute", left: 520, bottom: -12, width: 54, height: 20, background: "#ffe600" }} />
      <div style={{ position: "absolute", left: 900, bottom: -12, width: 54, height: 20, background: "#ffe600" }} />
    </div>
  );
};
const Store: React.FC<{ blur?: number; dim?: number }> = ({ blur = 0, dim = 0.35 }) => (
  <AbsoluteFill style={{ background: "linear-gradient(180deg,#0f172a 0%,#1e3a5f 38%,#0b3b3a 100%)" }}>
    <div style={{ position: "absolute", top: 60, left: 80, right: 80, height: 18, borderRadius: 9, background: "rgba(255,255,255,0.9)", boxShadow: "0 0 90px 30px rgba(180,220,255,0.35)" }} />
    <div style={{ position: "absolute", top: 220, left: 0, right: 0, height: 1000, filter: `blur(${blur}px)` }}>
      {[0, 1, 2, 3].map((r) => <Shelf key={r} row={r} y={r * 250} />)}
    </div>
    <AbsoluteFill style={{ background: `rgba(5,8,18,${dim})` }} />
  </AbsoluteFill>
);

const SKIN = "#d9a77c", SKIN2 = "#c18e63";
const Screen: React.FC<{ sc?: number }> = ({ sc = 1 }) => (
  <div style={{ position: "absolute", left: 0, top: 0, width: `${100 / sc}%`, height: `${100 / sc}%`, transform: `scale(${sc})`, transformOrigin: "top left", background: "#0a1020", fontFamily: "Manrope, sans-serif", color: "#fff", padding: "40px 22px 26px" }}>
    <div style={{ fontWeight: 800, fontSize: 22, color: "#9aa6ba", marginBottom: 8 }}>Hortifruti · por quilo</div>
    <div style={{ fontFamily: "Unbounded, sans-serif", fontWeight: 700, fontSize: 30 }}>Tomate italiano</div>
    <div style={{ fontFamily: "Unbounded, sans-serif", fontWeight: 800, fontSize: 92, lineHeight: 1.1, marginTop: 14 }}>0,640<span style={{ fontSize: 30, color: "#9aa6ba" }}> kg</span></div>
    <div style={{ fontWeight: 700, fontSize: 22, color: "#9aa6ba" }}>R$ 8,90 / kg</div>
    <div style={{ marginTop: 16, borderRadius: 16, padding: "16px 18px", background: "linear-gradient(120deg,#1FE0C8,#38BDF8)", color: "#06101C", display: "flex", justifyContent: "space-between", fontFamily: "Unbounded, sans-serif", fontWeight: 800, fontSize: 24 }}><span>Valor</span><span>R$ 5,70</span></div>
    {[["Pão francês · 0,350 kg", "R$ 5,57"], ["Arroz Tipo 1 5 kg", "R$ 28,90"]].map(([a, b]) => <div key={a} style={{ marginTop: 12, display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 22, color: "#cbd5e1", borderBottom: "1px solid rgba(255,255,255,0.1)", paddingBottom: 10 }}><span>{a}</span><span>{b}</span></div>)}
    <div style={{ marginTop: 22, borderRadius: 14, padding: "16px 0", textAlign: "center", background: "#10b981", fontWeight: 800, fontSize: 24 }}>Cobrar R$ 68,87</div>
  </div>
);
const Phone: React.FC<{ w: number; h: number; sc?: number; style?: React.CSSProperties; children?: React.ReactNode }> = ({ w, h, sc = 1, style, children }) => (
  <div style={{ position: "absolute", width: w, height: h, borderRadius: 56, background: "#05070d", border: "10px solid #0f1424", boxShadow: "0 50px 110px -20px rgba(0,0,0,0.9), 0 0 0 2px rgba(255,255,255,0.08)", overflow: "hidden", ...style }}>
    <div style={{ position: "absolute", inset: 0, borderRadius: 46, overflow: "hidden" }}><Screen sc={sc} />{children}</div>
    <div style={{ position: "absolute", top: 10, left: "50%", marginLeft: -48, width: 96, height: 24, borderRadius: 14, background: "#000" }} />
    <div style={{ position: "absolute", inset: 0, background: "linear-gradient(115deg,rgba(255,255,255,0.14) 0%,transparent 35%)", borderRadius: 46 }} />
  </div>
);
const Finger: React.FC<{ x: number; y: number; w?: number; r?: number }> = ({ x, y, w = 150, r = -8 }) => (
  <div style={{ position: "absolute", left: x, top: y, width: w, height: 64, borderRadius: 32, background: `linear-gradient(180deg,${SKIN},${SKIN2})`, transform: `rotate(${r}deg)`, boxShadow: "0 6px 14px rgba(0,0,0,0.35), inset 0 -6px 10px rgba(120,70,40,0.25)" }}><div style={{ position: "absolute", right: 14, top: 12, width: 28, height: 22, borderRadius: "50% 50% 40% 40%", background: "rgba(255,235,220,0.55)" }} /></div>
);

// A: mão segurando o celular no corredor (fundo desfocado)
export const IlusA: React.FC = () => (
  <AbsoluteFill>
    <Store blur={7} dim={0.3} />
    <div style={{ position: "absolute", left: 270, top: 270, width: 540, height: 1000, transform: "rotate(-5deg)" }}>
      <Phone w={540} h={1000} style={{ left: 0, top: 0 }} />
      {/* polegar (direita) e dedos (esquerda) */}
      <div style={{ position: "absolute", left: 500, top: 640, width: 150, height: 74, borderRadius: 37, background: `linear-gradient(180deg,${SKIN},${SKIN2})`, transform: "rotate(-20deg)", boxShadow: "0 6px 14px rgba(0,0,0,0.35)" }} />
      <Finger x={-70} y={700} w={150} r={6} /><Finger x={-78} y={770} w={160} r={4} /><Finger x={-70} y={840} w={150} r={2} />
      <div style={{ position: "absolute", left: 40, top: 880, width: 470, height: 360, borderRadius: "40px 40px 200px 200px", background: `linear-gradient(180deg,${SKIN},${SKIN2})`, boxShadow: "inset 0 -30px 50px rgba(120,70,40,0.35)" }} />
    </div>
    <Grain /><Vignette />
  </AbsoluteFill>
);

// B: balança do caixa com tomates e o celular
const Tomato: React.FC<{ x: number; y: number; s: number }> = ({ x, y, s }) => (
  <div style={{ position: "absolute", left: x, top: y, width: s, height: s * 0.9, borderRadius: "50% 50% 46% 46%", background: "radial-gradient(circle at 35% 30%,#ff7a6b,#dc2626 55%,#8f1414)", boxShadow: "0 10px 20px rgba(0,0,0,0.4)" }}><div style={{ position: "absolute", left: "34%", top: -s * 0.06, width: "32%", height: s * 0.16, background: "#16a34a", borderRadius: "50% 50% 10% 10%" }} /></div>
);
export const IlusB: React.FC = () => (
  <AbsoluteFill>
    <Store blur={10} dim={0.45} />
    <div style={{ position: "absolute", left: 0, right: 0, top: 880, bottom: 0, background: "linear-gradient(180deg,#a16a3c,#6b4423 45%,#3d2512)", boxShadow: "0 -30px 80px rgba(0,0,0,0.6)" }} />
    <div style={{ position: "absolute", left: 90, top: 700, width: 520, height: 260, borderRadius: 26, background: "linear-gradient(180deg,#e5e7eb,#9ca3af)", boxShadow: "0 30px 60px rgba(0,0,0,0.6)" }}>
      <div style={{ position: "absolute", left: 40, right: 40, top: -22, height: 46, borderRadius: 14, background: "linear-gradient(180deg,#f3f4f6,#cbd5e1)" }} />
      <div style={{ position: "absolute", left: 340, top: 80, width: 150, height: 90, borderRadius: 12, background: "#0b1224", color: "#1FE0C8", fontFamily: "Unbounded, sans-serif", fontWeight: 800, fontSize: 34, display: "flex", alignItems: "center", justifyContent: "center" }}>0,640</div>
    </div>
    <Tomato x={150} y={560} s={150} /><Tomato x={290} y={590} s={140} /><Tomato x={420} y={550} s={150} /><Tomato x={230} y={500} s={140} /><Tomato x={360} y={470} s={140} />
    <Phone w={400} h={780} sc={0.72} style={{ left: 640, top: 380, transform: "rotate(8deg)" }} />
    <div style={{ position: "absolute", left: 650, top: 740, width: 400, height: 8, borderRadius: 4, background: "#ff3b6b", boxShadow: "0 0 30px 8px rgba(255,59,107,0.6)", transform: "rotate(8deg)", opacity: 0.0 }} />
    <Grain /><Vignette />
  </AbsoluteFill>
);

// C: prateleira com alertas do sistema
const Notif: React.FC<{ x: number; y: number; w: number; c: string; icon: string; t: string; s: string; r?: number }> = ({ x, y, w, c, icon, t, s, r = 0 }) => (
  <div style={{ position: "absolute", left: x, top: y, width: w, padding: "26px 32px", borderRadius: 30, background: "rgba(10,16,32,0.86)", border: `2px solid ${c}`, boxShadow: `0 30px 80px -20px rgba(0,0,0,0.9), 0 0 60px ${c}44`, display: "flex", gap: 24, alignItems: "center", transform: `rotate(${r}deg)`, backdropFilter: "blur(12px)", fontFamily: "Manrope, sans-serif" }}>
    <div style={{ fontSize: 66 }}>{icon}</div>
    <div><div style={{ fontFamily: "Unbounded, sans-serif", fontWeight: 800, fontSize: 36, color: "#fff", lineHeight: 1.15 }}>{t}</div><div style={{ fontWeight: 700, fontSize: 28, color: c, marginTop: 6 }}>{s}</div></div>
  </div>
);
export const IlusC: React.FC = () => (
  <AbsoluteFill>
    <Store blur={2} dim={0.2} />
    <Notif x={90} y={330} w={800} c="#fbbf24" icon="⏳" t="12 itens vencendo" s="R$ 372 em risco" r={-2} />
    <Notif x={170} y={600} w={800} c="#1FE0C8" icon="🛒" t="Repor primeiro:" s="arroz e leite" r={1.5} />
    <Notif x={90} y={870} w={800} c="#a78bfa" icon="✦" t="Produtos parados" s="há 30 dias sem giro" r={-1} />
    <Grain /><Vignette />
  </AbsoluteFill>
);
