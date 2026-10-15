import React from "react";
import { AbsoluteFill } from "remotion";
import { BgMesh, C, F, GradText, Grain, Vignette } from "./components";

export const Capa: React.FC = () => (
  <AbsoluteFill>
    <BgMesh />
    <AbsoluteFill style={{ padding: "110px 80px", justifyContent: "space-between" }}>
      <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 30, letterSpacing: "0.26em", color: C.teal, textTransform: "uppercase" }}>Future PDV · mercados</div>
      <div>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 92, lineHeight: 1.08, color: C.text, letterSpacing: "-0.03em" }}>
          Trabalhei <GradText>30 dias</GradText> num supermercado
        </div>
        <div style={{ fontFamily: F.display, fontWeight: 600, fontSize: 56, lineHeight: 1.2, color: C.muted, marginTop: 40 }}>para montar um sistema.</div>
      </div>
      <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 36, color: C.muted }}>E aprendi o que nenhum requisito me contaria.</div>
    </AbsoluteFill>
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
