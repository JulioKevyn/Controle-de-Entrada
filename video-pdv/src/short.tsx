import React from "react";
import { AbsoluteFill, Audio, Sequence, interpolate, staticFile, useCurrentFrame } from "remotion";
import { AiBadge, BgMesh, C, Captions, Chip, Counter, Entrance, F, Grain, GradText, Mark, Scene, Sparkles, Vignette, WordReveal, breathe, clamp, float, panel, useSpring } from "./components";
import { theme } from "./theme";

// Cortes (em quadros) alinhados à fala; total = 24 s
const CUTS = [0, 93, 216, 321, 396, 495, 720];
export const SHORT_TOTAL = 720;
const CAPS = [
  { t: "Sua loja sabe quanto realmente lucra?", a: 0, b: 3.3 },
  { t: "Conheça o Future PDV,", a: 3.4, b: 5.2 },
  { t: "com inteligência artificial!", a: 5.2, b: 7.2 },
  { t: "Venda em segundos,", a: 7.3, b: 8.9 },
  { t: "veja o lucro de verdade,", a: 8.9, b: 10.7 },
  { t: "e pergunte pra IA", a: 10.8, b: 12.0 },
  { t: "o que está parado.", a: 12.0, b: 13.2 },
  { t: "Alertas no WhatsApp", a: 13.4, b: 14.9 },
  { t: "e a cara da sua marca.", a: 14.9, b: 16.5 },
  { t: "Teste grátis por sete dias!", a: 16.7, b: 20 },
];

const Center: React.FC<{ children: React.ReactNode; gap?: number; top?: number }> = ({ children, gap = 40, top }) => (
  <AbsoluteFill style={{ alignItems: "center", justifyContent: top ? "flex-start" : "center", flexDirection: "column", rowGap: gap, paddingTop: top, padding: "0 50px" }}>{children}</AbsoluteFill>
);

const A: React.FC = () => {
  const frame = useCurrentFrame();
  const q = useSpring(14, "bouncy");
  const flick = Math.floor(frame / 4) % 3;
  return (
    <Scene>
      <Center gap={50}>
        <WordReveal text="Sua loja sabe quanto realmente lucra?" delay={2} size={92} align="center" highlight={["lucra?"]} per={2} />
        <div style={{ opacity: Math.min(1, q), transform: `scale(${interpolate(q, [0, 1], [0.4, 1]) * breathe(frame, 0.04, 8)})`, fontFamily: F.display, fontWeight: 800, fontSize: 190, color: C.pink, textShadow: `0 0 90px ${C.pink}99` }}>R$ {["???", "?,?", "???"][flick]}</div>
      </Center>
    </Scene>
  );
};

const B: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <Scene>
      <Center gap={44}>
        <Mark size={260} />
        <Entrance delay={8} y={30} cfg="snappy">
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 126, letterSpacing: "-0.03em", lineHeight: 1, transform: `translateY(${float(frame, 3)}px)` }}>
            <span style={{ color: C.text }}>Future </span><GradText>PDV</GradText>
          </div>
        </Entrance>
        <AiBadge delay={24} big />
        <Entrance delay={40} y={20}><div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 36, color: C.muted, textAlign: "center" }}>Vendas, estoque e gestão da sua loja</div></Entrance>
      </Center>
    </Scene>
  );
};

const Cc: React.FC = () => {
  const frame = useCurrentFrame();
  const done = useSpring(46, "bouncy");
  const bars = [52, 60, 82, 66, 78, 72];
  return (
    <Scene>
      <Center gap={34}>
        <Entrance delay={0} y={60}>
          <div style={{ ...panel, width: 960, padding: "30px 38px" }}>
            <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 24, letterSpacing: "0.2em", color: C.teal, marginBottom: 14 }}>PDV</div>
            {[["1× Camiseta Preta G", "89,90"], ["1× Bermuda Azul 42", "149,90"], ["2× Boné", "79,80"]].map(([n, v]) => (
              <div key={n} style={{ display: "flex", justifyContent: "space-between", fontFamily: F.body, fontSize: 32, color: C.text, height: 62, alignItems: "center", borderBottom: `1px solid ${C.line}` }}><span>{n}</span><b>R$ {v}</b></div>
            ))}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 18 }}>
              <div style={{ fontFamily: F.display, fontSize: 44, color: C.text }}>Total <GradText>R$ 319,60</GradText></div>
              <div style={{ opacity: Math.min(1, done), transform: `scale(${done})` }}><Chip color={C.ok} size={26}>✓ Venda concluída</Chip></div>
            </div>
          </div>
        </Entrance>
        <Entrance delay={26} y={60}>
          <div style={{ ...panel, width: 960, padding: "30px 38px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 24, letterSpacing: "0.2em", color: C.teal }}>LUCRO DE VERDADE</div>
              <GradText style={{ fontFamily: F.display, fontWeight: 800, fontSize: 56 }}><Counter to={38} delay={40} />%</GradText>
            </div>
            <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", height: 260 }}>
              {bars.map((b, i) => {
                const g = useSpring(34 + i * 5, "smooth");
                return (
                  <div key={i} style={{ display: "flex", alignItems: "flex-end", columnGap: 6 }}>
                    {[[b, C.sky], [b * 0.68, C.pink], [b * 0.32, C.teal]].map(([v, c], k) => <div key={k} style={{ width: 34, height: (v as number) * 2.6 * g, borderRadius: "8px 8px 2px 2px", background: c as string }} />)}
                  </div>
                );
              })}
            </div>
          </div>
        </Entrance>
      </Center>
    </Scene>
  );
};

const typed = (text: string, frame: number, from: number, cps = 1.6) => text.slice(0, Math.max(0, Math.min(text.length, Math.floor((frame - from) * cps))));
const D: React.FC = () => {
  const frame = useCurrentFrame();
  const a = "O Boné está há mais de 30 dias parado, com 14 unidades. Que tal um combo com a camiseta?";
  return (
    <Scene>
      <Center gap={30}>
        <AiBadge delay={0} big />
        <WordReveal text="Pergunte pra sua loja." delay={4} size={80} align="center" highlight={["loja."]} per={2} />
        <div style={{ ...panel, width: 960, padding: "30px 36px", minHeight: 620 }}>
          <Entrance delay={10} y={30}>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 20 }}>
              <div style={{ padding: "20px 28px", borderRadius: "24px 24px 6px 24px", background: theme.grad, color: C.bg, fontFamily: F.body, fontWeight: 800, fontSize: 34 }}>Tem alguma coisa parada?</div>
            </div>
          </Entrance>
          {frame >= 28 && (
            <div style={{ maxWidth: 780, padding: "22px 30px", borderRadius: "24px 24px 24px 6px", background: "rgba(255,255,255,0.08)", border: `1px solid ${C.line}`, fontFamily: F.body, fontWeight: 600, fontSize: 34, lineHeight: 1.4, color: C.text }}>{typed(a, frame, 28, 1.7)}</div>
          )}
        </div>
      </Center>
    </Scene>
  );
};

const E: React.FC = () => {
  const frame = useCurrentFrame();
  const wa = useSpring(2, "bouncy");
  const sw = useSpring(34, "bouncy");
  const step = Math.min(3, Math.floor(Math.max(0, frame - 40) / 16));
  const cols = ["#00E0C6", "#F4A261", "#FF2E9A", "#7C3AED"];
  return (
    <Scene>
      <Center gap={44}>
        <div style={{ opacity: Math.min(1, wa), transform: `scale(${wa})`, width: 960 }}>
          <div style={{ ...panel, padding: "28px 34px" }}>
            <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 28, color: C.ok, marginBottom: 14 }}>● WhatsApp · Future PDV</div>
            <div style={{ padding: "20px 26px", borderRadius: 20, background: "rgba(255,255,255,0.08)", border: `1px solid ${C.line}`, fontFamily: F.body, fontWeight: 600, fontSize: 34, color: C.text, lineHeight: 1.35 }}>⚠ Estoque baixo: Camiseta Branca G (1 un.)</div>
          </div>
        </div>
        <div style={{ opacity: Math.min(1, sw), transform: `scale(${sw})`, width: 960 }}>
          <div style={{ ...panel, padding: "30px 34px", display: "flex", alignItems: "center", columnGap: 28, border: `2px solid ${cols[step]}`, boxShadow: `0 0 70px ${cols[step]}55` }}>
            <div style={{ width: 96, height: 96, borderRadius: 28, background: cols[step], color: "#0B1020", fontFamily: F.display, fontWeight: 800, fontSize: 28, display: "flex", alignItems: "center", justifyContent: "center" }}>SUA</div>
            <div><div style={{ fontFamily: F.display, fontSize: 40, color: C.text }}>A cara da sua marca</div><div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 28, color: C.muted, marginTop: 6 }}>Logo, cores e tema</div></div>
          </div>
        </div>
        <div style={{ display: "flex", columnGap: 22 }}>
          {cols.map((c, i) => <div key={c} style={{ width: 70, height: 70, borderRadius: 35, background: c, border: `4px solid ${i === step ? "#fff" : "transparent"}`, transform: `scale(${i === step ? 1.2 : 1})` }} />)}
        </div>
      </Center>
    </Scene>
  );
};

const F6: React.FC = () => {
  const frame = useCurrentFrame();
  const btn = useSpring(18, "bouncy");
  return (
    <Scene>
      <Center gap={44}>
        <Mark size={200} />
        <WordReveal text="Teste grátis por 7 dias" delay={6} size={92} align="center" highlight={["grátis"]} per={3} />
        <div style={{ opacity: Math.min(1, btn), transform: `scale(${btn * breathe(frame, 0.03, 10)})`, padding: "26px 64px", borderRadius: 999, background: theme.grad, color: C.bg, fontFamily: F.display, fontWeight: 800, fontSize: 48, boxShadow: `0 20px 70px -10px ${C.teal}99` }}>Criar minha loja</div>
        <Entrance delay={40} y={20}><Chip color={C.ok} size={30}>sem cartão · sem fidelidade</Chip></Entrance>
        <Entrance delay={56} y={20}><AiBadge big /></Entrance>
        <Entrance delay={70} y={20}><div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 28, color: C.muted, textAlign: "center" }}>Julio Clemente · (11) 96620-9914</div></Entrance>
      </Center>
    </Scene>
  );
};

const LIST = [A, B, Cc, D, E, F6];
export const ShortMain: React.FC = () => (
  <AbsoluteFill>
    <BgMesh />
    <Sparkles />
    {LIST.map((S, i) => (
      <Sequence key={i} from={CUTS[i]} durationInFrames={CUTS[i + 1] - CUTS[i]}>
        <S />
        <Audio src={staticFile("audio/whoosh.wav")} volume={0.4} />
      </Sequence>
    ))}
    <Captions chunks={CAPS} />
    <Audio src={staticFile("audio/sh.mp3")} volume={1} />
    <Audio src={staticFile("audio/music_short.wav")} volume={0.34} />
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
