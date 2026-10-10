import React from "react";
import { AbsoluteFill, interpolate } from "remotion";
import { clamp, easeInOut, easeOut, Grain, Pop, Small, Stars, Title, useBeat } from "./kit";

export const BPM = 120;
export const BEATS = 40;
const E = 28; // raio da Terra em px (escala base)
const J = E * 10.97; // Júpiter: 10,97x o diâmetro da Terra
const S = E * 109.1; // Sol: 109,1x

export const Sun01: React.FC = () => {
  const { frame, b, pulse, f } = useBeat(BPM);
  // câmera: 1 -> afasta para mostrar o Sol inteiro
  const zoom1 = interpolate(b, [4, 5], [0, 1], clamp);
  const zoomOut = easeInOut((b - 10) / 4);
  const scale = 1 - zoomOut * 0.86;
  const sunIn = easeOut((b - 10.5) / 3);
  const jupX = interpolate(easeOut((b - 4) / 1.2), [0, 1], [900, 190]);
  const fill = easeOut((b - 18) / 8);
  const count = Math.round(1300000 * easeOut((b - 18) / 9));
  const dots = Math.round(fill * 360);
  return (
    <AbsoluteFill style={{ background: "radial-gradient(ellipse at 50% 40%, #0b1330 0%, #04060f 70%)" }}>
      <Stars />
      <AbsoluteFill style={{ transform: `scale(${scale}) translate(${zoomOut * 150}px, ${zoomOut * 130}px)`, transformOrigin: "50% 52%" }}>
        {/* Sol */}
        <div style={{ position: "absolute", left: 540 - S * sunIn - 0, top: 990 - S * sunIn, width: S * 2 * sunIn, height: S * 2 * sunIn, borderRadius: "50%", background: "radial-gradient(circle at 40% 38%, #fff6c8 0%, #ffd23c 22%, #ff9a1f 55%, #e2420f 100%)", boxShadow: `0 0 ${240 + pulse * 120}px ${80 + pulse * 40}px rgba(255,150,30,0.55)`, opacity: sunIn > 0 ? 1 : 0 }}>
          {b > 18 && (
            <div style={{ position: "absolute", inset: 0, borderRadius: "50%", overflow: "hidden" }}>
              {Array.from({ length: dots }).map((_, i) => {
                const a = i * 2.399963;
                const r = Math.sqrt((i + 0.5) / 360) * S * 0.92;
                return <div key={i} style={{ position: "absolute", left: S + Math.cos(a) * r - 20, top: S + Math.sin(a) * r - 20, width: 40, height: 40, borderRadius: "50%", background: "radial-gradient(circle at 35% 35%, #7fd1ff, #1b6bd6 60%, #0b2f73)", boxShadow: "0 0 10px rgba(0,0,0,0.5)" }} />;
              })}
            </div>
          )}
        </div>
        {/* Júpiter */}
        <div style={{ position: "absolute", left: 540 + jupX - J - 0, top: 990 - J, width: J * 2, height: J * 2, borderRadius: "50%", background: "linear-gradient(180deg, #e8c9a0 0%, #c78f5b 18%, #efd6b4 32%, #b9763f 46%, #e3b98a 60%, #a8683a 76%, #d9ac7c 100%)", boxShadow: "inset -60px -40px 120px rgba(0,0,0,0.55)", opacity: b > 4 ? 1 : 0, transform: `translateX(${b > 10 ? zoomOut * -260 : 0}px)` }} />
        {/* Terra */}
        <div style={{ position: "absolute", left: 540 - E - (b > 4 ? 380 * easeOut((b - 4) / 1.2) : 0), top: 990 - E, width: E * 2, height: E * 2, borderRadius: "50%", background: "radial-gradient(circle at 35% 35%, #8fe0ff, #1b6bd6 55%, #0b2f73)", boxShadow: `0 0 ${30 + pulse * 20}px rgba(80,170,255,0.8)`, transform: `scale(${(b < 4 ? 0.2 + easeOut(b / 1.5) * 0.8 : 1) * (1 + pulse * 0.06)})` }} />
      </AbsoluteFill>
      {/* textos */}
      <AbsoluteFill style={{ padding: "230px 60px 0" }}>
        {b < 4 && <Pop at={0} b={BPM}><Title size={130}>HOW BIG<br />IS THE SUN?</Title></Pop>}
        {b >= 4 && b < 10 && (<Pop at={4} b={BPM}><Title size={110}>JUPITER<br /><span style={{ color: "#ffb347" }}>11×</span> WIDER<br />THAN EARTH</Title></Pop>)}
        {b >= 10 && b < 18 && (<Pop at={14} b={BPM}><Title size={110}>THE SUN IS<br /><span style={{ color: "#ffd23c" }}>109×</span> WIDER</Title></Pop>)}
        {b >= 18 && b < 30 && (<div><Title size={170} color="#ffd23c">{count.toLocaleString("en-US")}</Title><div style={{ height: 14 }} /><Small size={46}>EARTHS FIT INSIDE THE SUN</Small></div>)}
        {b >= 30 && (<Pop at={30} b={BPM}><Title size={130}>1.3 MILLION<br />EARTHS<br /><span style={{ color: "#ffd23c" }}>IN ONE SUN</span></Title></Pop>)}
      </AbsoluteFill>
      <Grain />
    </AbsoluteFill>
  );
};
