import React from "react";
import { AbsoluteFill, Img, interpolate, staticFile, useCurrentFrame, useVideoConfig } from "remotion";
import { C, clamp, Counter, Entrance, F, GradText, Scene, useSpring, WordReveal } from "./components";
import { theme } from "./theme";

const useAt = () => { const { durationInFrames: d } = useVideoConfig(); return (f: number) => Math.round(d * f); };
const Logo: React.FC<{ h?: number }> = ({ h = 70 }) => <Img src={staticFile("logo-white.png")} style={{ height: h, objectFit: "contain" }} />;
const Label: React.FC<{ children: React.ReactNode; delay?: number }> = ({ children, delay = 0 }) => (
  <Entrance delay={delay} y={20}><div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 26, letterSpacing: "0.28em", color: C.teal, textTransform: "uppercase" }}>{children}</div></Entrance>
);
const Head: React.FC<{ k: string; t: string }> = ({ k, t }) => (
  <div style={{ position: "absolute", top: 70, left: 0, right: 0, textAlign: "center" }}>
    <Label>{k}</Label>
    <Entrance delay={6} y={24}><div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 64, color: C.text, marginTop: 12 }}>{t}</div></Entrance>
  </div>
);
const Card: React.FC<{ children: React.ReactNode; color?: string; w?: number; h?: number; delay?: number; style?: React.CSSProperties }> = ({ children, color = C.teal, w, h, delay = 0, style }) => {
  const p = useSpring(delay, "bouncy");
  return (
    <div style={{ width: w, height: h, opacity: Math.min(1, p), transform: `scale(${interpolate(p, [0, 1], [0.8, 1])})`, borderRadius: 28, background: C.panel, border: `2px solid ${color}`, boxShadow: `0 0 40px ${color}33, 0 30px 70px -24px rgba(0,0,0,0.8)`, boxSizing: "border-box", padding: "30px 34px", ...style }}>{children}</div>
  );
};

const PILLARS = ["Inovação", "Inteligência Artificial", "Automação"];

export const E1: React.FC = () => {
  const p = useSpring(2, "bouncy");
  return (
    <Scene>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", rowGap: 22 }}>
        <Entrance delay={0}><Logo h={86} /></Entrance>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 340, lineHeight: 1, letterSpacing: "0.05em", transform: `scale(${interpolate(p, [0, 1], [0.6, 1])})`, opacity: Math.min(1, p), filter: "drop-shadow(0 0 50px rgba(236,103,7,0.4))" }}><GradText>IIA</GradText></div>
        <Entrance delay={26} y={20}><div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 52, color: C.text, letterSpacing: "0.06em" }}>MUNDIAL LOGISTICS</div></Entrance>
        <div style={{ display: "flex", gap: 24, marginTop: 14 }}>
          {PILLARS.map((x, i) => (
            <Entrance key={x} delay={60 + i * 30} y={26}>
              <div style={{ padding: "14px 34px", borderRadius: 999, border: `2px solid ${C.teal}`, background: `${C.teal}1F`, fontFamily: F.body, fontWeight: 700, fontSize: 36, color: C.text }}>{x}</div>
            </Entrance>
          ))}
        </div>
        <Entrance delay={170} y={20}><div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 38, color: C.muted, marginTop: 10 }}>A equipe que automatiza processos</div></Entrance>
      </AbsoluteFill>
    </Scene>
  );
};

export const E2: React.FC = () => {
  const frame = useCurrentFrame();
  const at = useAt();
  return (
    <Scene>
      <AbsoluteFill style={{ flexDirection: "row", alignItems: "center", justifyContent: "center", gap: 120 }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 420, lineHeight: 1, filter: "drop-shadow(0 0 50px rgba(236,103,7,0.4))" }}><GradText><Counter to={48} delay={8} /></GradText></div>
          <Entrance delay={40} y={20}><div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 64, color: C.text, marginTop: 6 }}>automações</div></Entrance>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(8, 84px)", gap: 14 }}>
          {Array.from({ length: 48 }).map((_, i) => {
            const t = interpolate(frame, [10 + i * 2.6, 22 + i * 2.6], [0, 1], clamp);
            return <div key={i} style={{ width: 84, height: 84, borderRadius: 18, background: `rgba(236,103,7,${0.12 + 0.55 * t})`, border: `2px solid ${t > 0.5 ? C.teal : C.line}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 38, color: "#fff", transform: `scale(${0.8 + 0.2 * t})`, boxShadow: t > 0.5 ? `0 0 18px ${C.teal}66` : undefined }}>{t > 0.6 ? "⚙" : ""}</div>;
          })}
        </div>
      </AbsoluteFill>
      <Entrance delay={at(0.55)} y={20} style={{ position: "absolute", bottom: 135, left: 0, right: 0, textAlign: "center" }}>
        <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 40, color: C.muted }}>menos planilha manual, menos retrabalho</div>
      </Entrance>
    </Scene>
  );
};

const STEPS = [
  { t: "Necessidade", s: "da operação", i: "🎯", c: C.sky },
  { t: "Desenvolvimento", s: "branch + commits", i: "💻", c: C.teal },
  { t: "Homologação", s: "testes e aprovação", i: "🧪", c: C.warn },
  { t: "Servidor", s: "entrega segura", i: "🖥️", c: C.ok },
];

export const E3: React.FC = () => {
  const frame = useCurrentFrame();
  const at = useAt();
  return (
    <Scene>
      <Head k="Como entregamos" t="Qualidade e segurança" />
      <AbsoluteFill style={{ flexDirection: "row", alignItems: "center", justifyContent: "center", gap: 70, paddingTop: 40 }}>
        {STEPS.map((s, i) => {
          const d = at(0.12 + i * 0.2);
          const ln = interpolate(frame, [d + 20, d + 45], [0, 1], clamp);
          return (
            <React.Fragment key={s.t}>
              <Card color={s.c} w={340} h={300} delay={d} style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", rowGap: 10, textAlign: "center" }}>
                <div style={{ fontSize: 84 }}>{s.i}</div>
                <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 34, color: s.c }}>{s.t}</div>
                <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 26, color: C.muted }}>{s.s}</div>
              </Card>
              {i < 3 && <div style={{ position: "absolute", left: 0 }} />}
            </React.Fragment>
          );
        })}
      </AbsoluteFill>
    </Scene>
  );
};

const CHECKS = [
  ["CNPJ / CPF", "ativo na Receita"],
  ["CEP e endereço", "planilha × cadastro"],
  ["Saldo de estoque", "por produto e classe"],
  ["Vínculo", "destinatário × depositante"],
];

export const E4: React.FC = () => {
  const at = useAt();
  return (
    <Scene>
      <Head k="Exemplo real" t="Solicitação de Materiais" />
      <AbsoluteFill style={{ flexDirection: "row", alignItems: "center", justifyContent: "center", gap: 90, paddingTop: 90 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 28 }}>
          <Card w={420} h={230} delay={at(0.1)} color={C.sky} style={{ textAlign: "center" }}>
            <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 120, lineHeight: 1, color: C.sky }}>16</div>
            <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 32, color: C.text, marginTop: 8 }}>layouts de planilha</div>
          </Card>
          <Card w={420} h={230} delay={at(0.2)} color={C.teal} style={{ textAlign: "center" }}>
            <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 120, lineHeight: 1, color: C.teal }}>100+</div>
            <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 32, color: C.text, marginTop: 8 }}>depositantes</div>
          </Card>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
          {CHECKS.map(([t, s], i) => (
            <Card key={t} w={780} h={118} delay={at(0.4 + i * 0.12)} color={C.ok} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 36px" }}>
              <div>
                <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 38, color: C.text }}>{t}</div>
                <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 26, color: C.muted }}>{s}</div>
              </div>
              <div style={{ fontSize: 60, color: C.ok }}>✔</div>
            </Card>
          ))}
        </div>
      </AbsoluteFill>
    </Scene>
  );
};

export const E5: React.FC = () => {
  const frame = useCurrentFrame();
  const at = useAt();
  const bar1 = interpolate(frame, [at(0.12), at(0.3)], [0, 1], { ...clamp, easing: theme.ease.out });
  const bar2 = interpolate(frame, [at(0.3), at(0.42)], [0, 1], { ...clamp, easing: theme.ease.out });
  return (
    <Scene>
      <Head k="Resultado" t="Ganho real de tempo" />
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", paddingTop: 70, rowGap: 34 }}>
        <div style={{ width: 1400 }}>
          <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 34, color: C.muted, marginBottom: 10 }}>Antes · 4h30</div>
          <div style={{ height: 70, borderRadius: 16, background: `linear-gradient(90deg, ${C.red}, #FB923C)`, width: `${bar1 * 100}%` }} />
        </div>
        <div style={{ width: 1400 }}>
          <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 34, color: C.ok, marginBottom: 10 }}>Agora · 27 min</div>
          <div style={{ height: 70, borderRadius: 16, background: C.ok, width: `${bar2 * 100 * 0.1}%`, minWidth: bar2 > 0 ? 8 : 0 }} />
        </div>
        <div style={{ display: "flex", gap: 40, marginTop: 30 }}>
          <Card w={420} h={200} delay={at(0.45)} color={C.ok} style={{ textAlign: "center" }}>
            <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 96, color: C.ok, lineHeight: 1 }}>−90%</div>
            <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 30, color: C.text, marginTop: 8 }}>de tempo</div>
          </Card>
          <Card w={420} h={200} delay={at(0.6)} color={C.teal} style={{ textAlign: "center" }}>
            <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 96, color: C.teal, lineHeight: 1 }}><Counter to={1904} delay={at(0.6)} /></div>
            <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 30, color: C.text, marginTop: 8 }}>pacotes</div>
          </Card>
          <Card w={420} h={200} delay={at(0.72)} color={C.sky} style={{ textAlign: "center" }}>
            <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 96, color: C.sky, lineHeight: 1 }}><Counter to={27787} delay={at(0.72)} /></div>
            <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 30, color: C.text, marginTop: 8 }}>destinos</div>
          </Card>
        </div>
      </AbsoluteFill>
    </Scene>
  );
};

const Benefits: React.FC<{ k: string; t: string; color: string; items: [string, string, string][] }> = ({ k, t, color, items }) => {
  const at = useAt();
  return (
    <Scene>
      <Head k={k} t={t} />
      <AbsoluteFill style={{ flexDirection: "row", alignItems: "center", justifyContent: "center", gap: 50, paddingTop: 80 }}>
        {items.map(([ic, ti, su], i) => (
          <Card key={ti} w={500} h={400} delay={at(0.14 + i * 0.2)} color={color} style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", rowGap: 14, textAlign: "center" }}>
            <div style={{ fontSize: 110 }}>{ic}</div>
            <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 44, color }}>{ti}</div>
            <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 30, color: C.muted, lineHeight: 1.25 }}>{su}</div>
          </Card>
        ))}
      </AbsoluteFill>
    </Scene>
  );
};

export const E6: React.FC = () => (
  <Benefits k="Para o cliente" t="Mais valor a cada pedido" color={C.sky} items={[["⚡", "Rapidez", "pedido validado e orçado em minutos"], ["🎯", "Menos erros", "validação automática de cada linha"], ["👁️", "Transparência", "compara fretes e aprova no portal"]]} />
);
export const E7: React.FC = () => (
  <Benefits k="Para a empresa" t="Mais eficiência na operação" color={C.teal} items={[["⏱️", "Tempo liberado", "equipe focada no que importa"], ["♻️", "Menos retrabalho", "conferência feita pelo sistema"], ["🔍", "Rastreabilidade", "histórico de cada etapa"]]} />
);

export const E8: React.FC = () => {
  const at = useAt();
  const p = useSpring(at(0.5), "bouncy");
  return (
    <Scene>
      <AbsoluteFill style={{ flexDirection: "row", alignItems: "center", justifyContent: "center", gap: 60 }}>
        <Card w={520} h={300} delay={at(0.05)} color={C.sky} style={{ display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", rowGap: 10 }}>
          <div style={{ fontSize: 100, color: C.ok }}>✔</div>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 60, color: C.sky }}>Cliente</div>
        </Card>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 120, color: C.text, transform: `scale(${p})` }}>+</div>
        <Card w={520} h={300} delay={at(0.25)} color={C.teal} style={{ display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", rowGap: 10 }}>
          <div style={{ fontSize: 100, color: C.ok }}>✔</div>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 60, color: C.teal }}>Empresa</div>
        </Card>
      </AbsoluteFill>
      <div style={{ position: "absolute", bottom: 290, left: 0, right: 0, display: "flex", justifyContent: "center" }}>
        <WordReveal text="Ganho real para os dois lados" size={64} delay={at(0.5)} align="center" highlight={["real"]} />
      </div>
    </Scene>
  );
};

export const E9: React.FC = () => {
  const p = useSpring(2, "bouncy");
  return (
    <Scene>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", rowGap: 20 }}>
        <Entrance delay={0}><Logo h={96} /></Entrance>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 300, lineHeight: 1, letterSpacing: "0.05em", transform: `scale(${interpolate(p, [0, 1], [0.6, 1])})`, opacity: Math.min(1, p), filter: "drop-shadow(0 0 60px rgba(236,103,7,0.45))" }}><GradText>IIA</GradText></div>
        <WordReveal text="Inovação · Inteligência Artificial · Automação" size={50} delay={20} align="center" highlight={["Inovação", "Inteligência", "Artificial", "Automação"]} />
        <Entrance delay={60} y={20}><div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 44, color: C.muted, marginTop: 20 }}>Fazendo marcas venderem mais</div></Entrance>
      </AbsoluteFill>
    </Scene>
  );
};
