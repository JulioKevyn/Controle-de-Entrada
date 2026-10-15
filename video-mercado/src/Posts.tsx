import React from "react";
import { AbsoluteFill } from "remotion";
import { BgMesh, C, F, GradText, Grain, Vignette } from "./components";

const Frame: React.FC<{ tag: string; title: React.ReactNode; children: React.ReactNode }> = ({ tag, title, children }) => (
  <AbsoluteFill>
    <BgMesh />
    <AbsoluteFill style={{ padding: "90px 70px 70px", flexDirection: "column" }}>
      <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 30, letterSpacing: "0.26em", color: C.teal, textTransform: "uppercase" }}>{tag}</div>
      <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 68, lineHeight: 1.08, color: C.text, letterSpacing: "-0.03em", marginTop: 26 }}>{title}</div>
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>{children}</div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontFamily: F.body, fontWeight: 800, fontSize: 28, color: C.muted }}>
        <span>Future<GradText>PDV</GradText> · versão para mercados</span>
        <span style={{ fontSize: 22, fontWeight: 700, color: C.dim }}>dados de exemplo</span>
      </div>
    </AbsoluteFill>
    <Grain />
    <Vignette />
  </AbsoluteFill>
);

const card: React.CSSProperties = { background: "rgba(255,255,255,0.06)", border: "1.5px solid rgba(255,255,255,0.14)", borderRadius: 34, boxShadow: "0 40px 90px -30px rgba(0,0,0,0.85)", backdropFilter: "blur(10px)" };

export const Post1: React.FC = () => (
  <Frame tag="Venda por peso" title={<>Tomate não é<br />“1 unidade”.<br /><GradText>É peso.</GradText></>}>
    <div style={{ ...card, width: 900, padding: 44 }}>
      <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 30, color: C.muted }}>Hortifruti · por quilo</div>
      <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 52, color: C.text, marginTop: 8 }}>Tomate italiano</div>
      <div style={{ display: "flex", alignItems: "baseline", gap: 18, marginTop: 28 }}>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 150, lineHeight: 1, color: C.text }}>0,640</div>
        <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 60, color: C.muted }}>kg</div>
      </div>
      <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 34, color: C.muted, marginTop: 6 }}>R$ 8,90 / kg</div>
      <div style={{ marginTop: 34, padding: "26px 34px", borderRadius: 24, background: "linear-gradient(120deg,#1FE0C8,#38BDF8)", display: "flex", justifyContent: "space-between", alignItems: "center", fontFamily: F.display, fontWeight: 800, color: "#06101C" }}>
        <span style={{ fontSize: 38 }}>Valor do item</span><span style={{ fontSize: 62 }}>R$ 5,70</span>
      </div>
      <div style={{ display: "flex", gap: 16, marginTop: 30 }}>
        {[["PIX", C.ok], ["Dinheiro", C.warn], ["Cartão", C.sky]].map(([n, c]) => <span key={n} style={{ padding: "12px 30px", borderRadius: 999, border: `2px solid ${c}`, color: c, fontFamily: F.body, fontWeight: 800, fontSize: 30 }}>{n}</span>)}
      </div>
    </div>
  </Frame>
);

export const Post2: React.FC = () => {
  const rows = [["Iogurte natural 170 g", "venceu há 2 dias", "R$ 34,20", C.red], ["Queijo mussarela fatiado", "vence em 2 dias", "R$ 240,00", C.warn], ["Pão de forma", "vence em 6 dias", "R$ 98,40", C.ok]];
  return (
    <Frame tag="Validade por lote" title={<>Veja o que vence <GradText>antes de virar prejuízo</GradText></>}>
      <div style={{ width: 920, display: "flex", flexDirection: "column", gap: 18 }}>
        <div style={{ ...card, padding: "26px 44px", border: `2px solid ${C.warn}`, boxShadow: `0 0 80px ${C.warn}44` }}>
          <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 28, color: C.warn, letterSpacing: "0.1em" }}>VALOR EM RISCO</div>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 104, lineHeight: 1.05, color: C.text, whiteSpace: "nowrap" }}>R$ 372,60</div>
        </div>
        {rows.map(([n, s, v, c]) => (
          <div key={n} style={{ ...card, padding: "20px 36px", display: "flex", justifyContent: "space-between", alignItems: "center", borderLeft: `10px solid ${c}` }}>
            <div>
              <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 32, color: C.text }}>{n}</div>
              <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 26, color: c, marginTop: 4 }}>{s}</div>
            </div>
            <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 36, color: C.text, whiteSpace: "nowrap" }}>{v}</div>
          </div>
        ))}
      </div>
    </Frame>
  );
};

export const Post3: React.FC = () => (
  <Frame tag="Fiado" title={<>Chega de <GradText>caderneta perdida</GradText></>}>
    <div style={{ width: 900, display: "flex", flexDirection: "column", gap: 26 }}>
      <div style={{ display: "flex", gap: 24 }}>
        <div style={{ flex: 1.4, borderRadius: 34, padding: "36px 40px", background: "linear-gradient(135deg,#10b981,#0e9f8e)", boxShadow: "0 30px 80px -20px rgba(16,185,129,0.6)" }}>
          <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 28, color: "rgba(255,255,255,0.85)" }}>A RECEBER</div>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 66, color: "#fff", lineHeight: 1.1, whiteSpace: "nowrap" }}>R$ 725,50</div>
          <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 28, color: "rgba(255,255,255,0.85)", marginTop: 6 }}>2 clientes com saldo</div>
        </div>
        <div style={{ flex: 1, ...card, padding: "36px 36px", border: `2px solid ${C.red}` }}>
          <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 28, color: C.red }}>ATRASADO</div>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 40, color: C.text, marginTop: 12, whiteSpace: "nowrap" }}>R$ 240,00</div>
        </div>
      </div>
      {[["Dona Marlene", "R$ 640,00", "R$ 240,00 atrasado"], ["Seu Joaquim", "R$ 85,50", ""]].map(([n, v, a]) => (
        <div key={n} style={{ ...card, padding: "30px 40px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 36, color: C.text }}>{n}</div>
            {a && <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 26, color: C.red, marginTop: 4 }}>{a}</div>}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 26 }}>
            <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 36, color: C.text, whiteSpace: "nowrap" }}>{v}</div>
            <div style={{ padding: "14px 26px", borderRadius: 16, background: "#10b981", color: "#fff", fontFamily: F.body, fontWeight: 800, fontSize: 28 }}>Receber</div>
          </div>
        </div>
      ))}
    </div>
  </Frame>
);

const Label: React.FC<{ n: string; whole: string; cents: string; unit?: string; rot: number }> = ({ n, whole, cents, unit, rot }) => (
  <div style={{ width: 700, background: "#FFE600", borderRadius: 14, padding: "20px 32px 16px", transform: `rotate(${rot}deg)`, boxShadow: "0 30px 70px -20px rgba(0,0,0,0.8)", color: "#111", fontFamily: F.display }}>
    <div style={{ fontWeight: 800, fontSize: 34 }}>{n}</div>
    <div style={{ display: "flex", alignItems: "flex-start", gap: 12, lineHeight: 1 }}>
      <span style={{ fontWeight: 800, fontSize: 40, marginTop: 20 }}>R$</span>
      <span style={{ fontWeight: 800, fontSize: 120 }}>{whole}</span>
      <span style={{ fontWeight: 800, fontSize: 58, marginTop: 12 }}>,{cents}</span>
      {unit && <span style={{ fontWeight: 700, fontSize: 34, marginTop: 80 }}>{unit}</span>}
    </div>
    <div style={{ display: "flex", gap: 4, height: 40, marginTop: 8 }}>{Array.from({ length: 58 }).map((_, i) => <div key={i} style={{ width: 2 + ((i * 7) % 5), background: "#111", height: "100%", marginRight: 2 }} />)}</div>
  </div>
);
export const Post4: React.FC = () => (
  <Frame tag="Etiquetas de gôndola" title={<>Etiqueta de gôndola, <GradText>pronta para imprimir</GradText></>}>
    <div style={{ display: "flex", flexDirection: "column", gap: 20, alignItems: "center" }}>
      <Label n="ARROZ TIPO 1 5 KG" whole="28" cents="90" rot={-2} />
      <Label n="TOMATE ITALIANO" whole="8" cents="90" unit="/kg" rot={1.5} />
      <Label n="PÃO FRANCÊS" whole="15" cents="90" unit="/kg" rot={-1} />
    </div>
  </Frame>
);

export const Post5: React.FC = () => {
  const rows: [string, number][] = [["Arroz Tipo 1 5 kg", 89.6], ["Pão francês", 81.1], ["Sabão em pó 1 kg", 70], ["Refrigerante 2 L", 68.4], ["Tomate italiano", 66.3], ["Leite integral 1 L", 42.3]];
  return (
    <Frame tag="Lucro por produto" title={<>Quanto você lucra <GradText>em cada produto</GradText></>}>
      <div style={{ ...card, width: 920, padding: "44px 48px", display: "flex", flexDirection: "column", gap: 30 }}>
        <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 28, color: C.muted, letterSpacing: "0.1em" }}>MARGEM</div>
        {rows.map(([n, m], i) => (
          <div key={n}>
            <div style={{ display: "flex", justifyContent: "space-between", fontFamily: F.body, fontWeight: 800, fontSize: 32, color: C.text, marginBottom: 10 }}>
              <span>{n}</span><span style={{ color: m > 60 ? C.ok : C.warn }}>{m.toFixed(1).replace(".", ",")}%</span>
            </div>
            <div style={{ height: 30, borderRadius: 15, background: "rgba(255,255,255,0.08)" }}>
              <div style={{ width: `${m}%`, height: "100%", borderRadius: 15, background: m > 60 ? "linear-gradient(90deg,#1FE0C8,#10b981)" : "linear-gradient(90deg,#FBBF24,#f97316)", boxShadow: `0 0 24px ${m > 60 ? "#1FE0C888" : "#FBBF2488"}` }} />
            </div>
          </div>
        ))}
      </div>
    </Frame>
  );
};

export const Post6: React.FC = () => {
  const b = (me: boolean, t: string) => (
    <div style={{ display: "flex", justifyContent: me ? "flex-end" : "flex-start" }}>
      <div style={{ maxWidth: 640, padding: "28px 36px", borderRadius: 34, borderBottomRightRadius: me ? 8 : 34, borderBottomLeftRadius: me ? 34 : 8, background: me ? "linear-gradient(120deg,#0e9f8e,#1b6bd6)" : "rgba(255,255,255,0.1)", fontFamily: F.body, fontWeight: 700, fontSize: 40, color: "#fff", lineHeight: 1.25 }}>{t}</div>
    </div>
  );
  return (
    <Frame tag="Plano Master IA" title={<>Pergunte para <GradText>o seu mercado</GradText></>}>
      <div style={{ ...card, width: 900, padding: "40px 44px", display: "flex", flexDirection: "column", gap: 26 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 40, color: C.text }}>Assistente IA</div>
          <div style={{ padding: "10px 26px", borderRadius: 999, background: "linear-gradient(120deg,#8B5CF6,#EC4899)", color: "#fff", fontFamily: F.body, fontWeight: 800, fontSize: 26 }}>✦ Inteligência Artificial</div>
        </div>
        {b(true, "O que eu preciso repor primeiro?")}
        {b(false, "Repor primeiro: arroz e leite.")}
        {b(true, "Tem algo vencendo?")}
        {b(false, "12 itens vencendo, com R$ 372 em risco.")}
        <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 28, color: C.muted }}>no sistema ou no WhatsApp</div>
      </div>
    </Frame>
  );
};
