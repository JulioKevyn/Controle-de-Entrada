import React from "react";
import { AbsoluteFill } from "remotion";
import { clamp, easeOut, Grain, Pop, Small, Title, useBeat } from "./kit";

export const BPM = 120;
export const BEATS = 40;
const R = 0.1 / 12; // 10% ao ano, capitalização mensal (ilustrativo)
const PMT = 300;
const fv = (years: number) => (years === 0 ? 0 : PMT * ((Math.pow(1 + R, years * 12) - 1) / R));
const MAX = fv(30);
const fmt = (n: number) => "$" + Math.round(n).toLocaleString("en-US");

export const Money01: React.FC = () => {
  const { b, pulse } = useBeat(BPM);
  const yearsF = Math.min(30, Math.max(0, ((b - 6) / 20) * 30)); // anos animados de b6 a b26
  const grow = easeOut((b - 6) / 20);
  const years = Math.min(30, 30 * (b < 6 ? 0 : (b - 6) / 20));
  const total = fv(years);
  const put = PMT * 12 * years;
  const H = 820;
  return (
    <AbsoluteFill style={{ background: "radial-gradient(ellipse at 50% 30%, #0b2a1c 0%, #040a07 72%)" }}>
      <AbsoluteFill style={{ padding: "220px 60px 0" }}>
        {b < 6 && (<Pop at={0} b={BPM}><Title size={150}>$300<br />A MONTH</Title><div style={{ height: 30 }} /><Pop at={3} b={BPM}><Small size={52}>FOR 30 YEARS. WHAT HAPPENS?</Small></Pop></Pop>)}
        {b >= 6 && b < 28 && (<div><Title size={170} color="#4ade80">{fmt(total)}</Title><div style={{ height: 8 }} /><Small size={42}>YEAR {Math.floor(years)} · YOU PUT IN {fmt(put)}</Small></div>)}
        {b >= 28 && b < 36 && (<Pop at={28} b={BPM}><Title size={120}><span style={{ color: "#9aa7b8" }}>YOU PUT IN</span><br />$108,000</Title><div style={{ height: 24 }} /><Pop at={32} b={BPM}><Title size={120}><span style={{ color: "#4ade80" }}>INTEREST MADE</span><br />{fmt(fv(30) - 108000)}</Title></Pop></Pop>)}
        {b >= 36 && (<Pop at={36} b={BPM}><Title size={150}>START<br /><span style={{ color: "#4ade80" }}>EARLY.</span></Title></Pop>)}
      </AbsoluteFill>
      {/* gráfico */}
      {b >= 6 && b < 36 && (
        <div style={{ position: "absolute", left: 40, right: 40, bottom: 330, height: H, display: "flex", alignItems: "flex-end", gap: 6 }}>
          {Array.from({ length: 30 }).map((_, i) => {
            const y = i + 1;
            const vis = Math.min(1, Math.max(0, years - i));
            const tot = fv(y);
            const contrib = PMT * 12 * y;
            const h = (tot / MAX) * H * vis;
            const hc = (contrib / MAX) * H * vis;
            return (
              <div key={i} style={{ flex: 1, height: h, position: "relative", opacity: vis > 0 ? 1 : 0 }}>
                <div style={{ position: "absolute", bottom: 0, left: 0, right: 0, height: hc, background: "linear-gradient(180deg,#8896a8,#5a6676)", borderRadius: "3px 3px 0 0" }} />
                <div style={{ position: "absolute", bottom: hc, left: 0, right: 0, height: Math.max(0, h - hc), background: "linear-gradient(180deg,#86efac,#16a34a)", borderRadius: "5px 5px 0 0", boxShadow: i === Math.floor(years) ? `0 0 ${20 + pulse * 30}px #4ade80` : undefined }} />
              </div>
            );
          })}
        </div>
      )}
      {b >= 6 && b < 36 && (
        <div style={{ position: "absolute", left: 60, bottom: 280, display: "flex", gap: 40 }}>
          <Small color="#8896a8" size={30}>■ YOU PUT IN</Small>
          <Small color="#4ade80" size={30}>■ INTEREST</Small>
        </div>
      )}
      <div style={{ position: "absolute", left: 0, right: 0, bottom: 200 }}><Small size={24} color="#7b8696">10%/yr return, illustrative only. Not financial advice.</Small></div>
      <Grain />
    </AbsoluteFill>
  );
};
