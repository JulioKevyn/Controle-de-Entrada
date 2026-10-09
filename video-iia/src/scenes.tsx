import React from "react";
import { AbsoluteFill, interpolate, useCurrentFrame, useVideoConfig } from "remotion";
import { AiBadge, C, clamp, Entrance, F, GradText, Scene, useSpring, useV, WordReveal } from "./components";
import { theme } from "./theme";

type N = { title: string; sub?: string; color: string; at: number; tag?: string; icon?: string };
type E = { label: string; at: number; dash?: boolean };

const NodeCard: React.FC<{ n: N; w: number; h: number; V: boolean }> = ({ n, w, h, V }) => {
  const p = useSpring(n.at, "bouncy");
  const frame = useCurrentFrame();
  const glow = 0.5 + 0.5 * Math.sin((frame - n.at) / 12);
  return (
    <div style={{ width: w, height: h, opacity: Math.min(1, p), transform: `scale(${interpolate(p, [0, 1], [0.7, 1])})`, borderRadius: 26, background: C.panel, border: `2px solid ${n.color}`, boxShadow: `0 0 ${24 + glow * 24}px ${n.color}55, 0 30px 70px -24px rgba(0,0,0,0.8)`, display: "flex", flexDirection: V ? "row" : "column", alignItems: "center", justifyContent: "center", columnGap: 26, rowGap: 8, padding: "0 18px", boxSizing: "border-box", backdropFilter: "blur(10px)" }}>
      {n.icon && <div style={{ fontSize: V ? 70 : 64, lineHeight: 1 }}>{n.icon}</div>}
      <div style={{ textAlign: V ? "left" : "center" }}>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: V ? 42 : 34, color: n.color, letterSpacing: "-0.01em" }}>{n.title}</div>
        {n.sub && <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: V ? 28 : 24, color: C.muted, marginTop: 6 }}>{n.sub}</div>}
        {n.tag && <div style={{ marginTop: 10, fontFamily: F.body, fontWeight: 800, fontSize: 20, color: n.color, letterSpacing: "0.1em" }}>{n.tag}</div>}
      </div>
    </div>
  );
};

export const Flow: React.FC<{ nodes: N[]; edges: E[]; x: number; y: number; w: number; h: number }> = ({ nodes, edges, x, y, w, h }) => {
  const { V } = useV();
  const frame = useCurrentFrame();
  const N_ = nodes.length;
  const nw = V ? Math.min(w, 800) : Math.min(N_ <= 3 ? 380 : 290, (w - 90 * (N_ - 1)) / N_);
  const nh = V ? Math.min(180, (h - 100 * (N_ - 1)) / N_) : 220;
  const gap = V ? (h - N_ * nh) / Math.max(1, N_ - 1) : (w - N_ * nw) / Math.max(1, N_ - 1);
  const pos = nodes.map((_, i) => V ? { x: x + (w - nw) / 2, y: y + i * (nh + gap) } : { x: x + i * (nw + gap), y: y + (h - nh) / 2 });
  return (
    <>
      {edges.map((e, i) => {
        const a = pos[i], b = pos[i + 1];
        if (!b) return null;
        const p = interpolate(frame, [e.at, e.at + 22], [0, 1], { ...clamp, easing: theme.ease.out });
        const x1 = V ? a.x + nw / 2 : a.x + nw, y1 = V ? a.y + nh : a.y + nh / 2;
        const x2 = V ? b.x + nw / 2 : b.x, y2 = V ? b.y : b.y + nh / 2;
        const t = ((frame - e.at - 22) % 40 + 40) % 40 / 40;
        const col = nodes[i + 1].color;
        return (
          <React.Fragment key={i}>
            <div style={{ position: "absolute", left: V ? x1 - 3 : x1, top: V ? y1 : y1 - 3, width: V ? 6 : (x2 - x1) * p, height: V ? (y2 - y1) * p : 6, borderRadius: 3, background: e.dash ? "transparent" : `linear-gradient(${V ? "180deg" : "90deg"}, ${nodes[i].color}, ${col})`, borderLeft: e.dash && V ? `4px dashed ${col}` : undefined, borderTop: e.dash && !V ? `4px dashed ${col}` : undefined, opacity: p }} />
            {p >= 1 && <div style={{ position: "absolute", left: (V ? x1 : x1 + (x2 - x1) * t) - 11, top: (V ? y1 + (y2 - y1) * t : y1) - 11, width: 22, height: 22, borderRadius: "50%", background: "#fff", boxShadow: `0 0 24px 6px ${col}` }} />}
            <div style={{ position: "absolute", left: V ? x1 + 26 : (x1 + x2) / 2 - 60, top: V ? (y1 + y2) / 2 - 18 : y1 - 52, width: V ? 360 : 120, textAlign: V ? "left" : "center", opacity: p, fontFamily: F.body, fontWeight: 800, fontSize: 26, color: col, letterSpacing: "0.08em" }}>{e.label}</div>
          </React.Fragment>
        );
      })}
      {nodes.map((n, i) => <div key={i} style={{ position: "absolute", left: pos[i].x, top: pos[i].y }}><NodeCard n={n} w={nw} h={nh} V={V} /></div>)}
    </>
  );
};

const Title: React.FC<{ k: string; t: string; delay?: number }> = ({ k, t, delay = 0 }) => {
  const { V } = useV();
  return (
    <Entrance delay={delay} y={24} style={{ position: "absolute", top: V ? 130 : 70, left: 0, right: 0, textAlign: "center" }}>
      <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 26, letterSpacing: "0.3em", color: C.teal, textTransform: "uppercase" }}>{k}</div>
      <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: V ? 64 : 58, color: C.text, marginTop: 12, lineHeight: 1.1 }}>{t}</div>
    </Entrance>
  );
};

const PILLARS = [
  { t: "Inovação", c: C.teal, i: "💡" },
  { t: "Inteligência Artificial", c: C.violet, i: "🤖" },
  { t: "Automação", c: C.ok, i: "⚙️" },
];

export const I1: React.FC = () => {
  const { V } = useV();
  const p = useSpring(2, "bouncy");
  return (
    <Scene>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", rowGap: V ? 46 : 28 }}>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: V ? 380 : 330, lineHeight: 1, letterSpacing: "0.02em", transform: `scale(${interpolate(p, [0, 1], [0.6, 1])})`, opacity: Math.min(1, p), filter: "drop-shadow(0 0 50px rgba(34,211,238,0.35))" }}><GradText>IIA</GradText></div>
        <div style={{ display: "flex", flexDirection: V ? "column" : "row", gap: 22, alignItems: "center" }}>
          {PILLARS.map((x, i) => (
            <Entrance key={x.t} delay={50 + i * 36} y={30}>
              <div style={{ display: "flex", alignItems: "center", columnGap: 16, padding: "16px 34px", borderRadius: 999, border: `2px solid ${x.c}`, background: `${x.c}1A`, fontFamily: F.display, fontWeight: 700, fontSize: V ? 46 : 36, color: x.c }}><span style={{ fontSize: V ? 52 : 40 }}>{x.i}</span>{x.t}</div>
            </Entrance>
          ))}
        </div>
        <Entrance delay={190} y={20}>
          <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: V ? 42 : 34, color: C.muted, marginTop: 14 }}>Nosso processo de desenvolvimento</div>
        </Entrance>
      </AbsoluteFill>
    </Scene>
  );
};

const Steps: React.FC<{ items: [string, number][]; x: number; y: number; w: number }> = ({ items, x, y, w }) => {
  const { V } = useV();
  return (
    <div style={{ position: "absolute", left: x, top: y, width: w, display: "flex", flexDirection: V ? "column" : "row", gap: 18, justifyContent: "center" }}>
      {items.map(([t, at], i) => (
        <Entrance key={i} delay={at} y={24} style={{ flex: V ? undefined : 1 }}>
          <div style={{ display: "flex", alignItems: "center", columnGap: 18, padding: "16px 24px", borderRadius: 20, background: C.panel, border: `1px solid ${C.line}`, fontFamily: F.body, fontWeight: 700, fontSize: V ? 34 : 28, color: C.text }}>
            <span style={{ width: 46, height: 46, borderRadius: "50%", background: theme.grad, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: F.display, fontWeight: 800, fontSize: 24, color: "#06101C", flexShrink: 0 }}>{i + 1}</span>{t}
          </div>
        </Entrance>
      ))}
    </div>
  );
};

export const I2: React.FC = () => {
  const { V, w, h } = useV();
  return (
    <Scene>
      <Title k="Passos 1 a 4" t="Da branch até a homologação" />
      <Flow
        x={V ? 140 : 360} y={V ? 400 : 250} w={V ? w - 280 : w - 720} h={V ? 700 : 300}
        nodes={[
          { title: "bugfix | feature", sub: "bugfix/<nome> · feature/<nome>", color: C.teal, at: 20, icon: "🌿" },
          { title: "homolog", sub: "testes · homologação", color: C.warn, at: 210, icon: "🧪" },
        ]}
        edges={[{ label: "PR", at: 170 }]}
      />
      <Steps x={V ? 100 : 160} y={V ? 1180 : 640} w={V ? w - 200 : w - 320} items={[["Criar a branch", 30], ["Desenvolver e commitar", 90], ["Abrir PR para homolog", 170], ["Homologar", 250]]} />
    </Scene>
  );
};

const Person: React.FC<{ name: string; color: string; delay: number; check?: boolean }> = ({ name, color, delay, check }) => {
  const { V } = useV();
  const p = useSpring(delay, "bouncy");
  const ck = useSpring(delay + 40, "bouncy");
  return (
    <div style={{ opacity: Math.min(1, p), transform: `scale(${interpolate(p, [0, 1], [0.7, 1])})`, display: "flex", alignItems: "center", columnGap: 18, padding: "18px 34px", borderRadius: 999, background: `${color}1F`, border: `2px solid ${color}`, fontFamily: F.display, fontWeight: 700, fontSize: V ? 44 : 36, color }}>
      <span style={{ width: V ? 64 : 54, height: V ? 64 : 54, borderRadius: "50%", background: color, color: "#06101C", display: "flex", alignItems: "center", justifyContent: "center" }}>{name[0]}</span>
      {name}
      {check && <span style={{ fontSize: V ? 54 : 44, color: C.ok, transform: `scale(${ck})`, opacity: Math.min(1, ck), marginLeft: 8 }}>✔</span>}
    </div>
  );
};

export const I3: React.FC = () => {
  const { V } = useV();
  const dir = V ? "column" : "row";
  const Col: React.FC<{ label: string; children: React.ReactNode; delay: number }> = ({ label, children, delay }) => (
    <Entrance delay={delay} y={20}>
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", rowGap: 16 }}>
        <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 24, letterSpacing: "0.22em", color: C.muted, textTransform: "uppercase" }}>{label}</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>{children}</div>
      </div>
    </Entrance>
  );
  return (
    <Scene>
      <Title k="Aprovação" t="PR para homolog" />
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: dir, gap: V ? 50 : 100, paddingTop: V ? 160 : 120 }}>
        <Col label="Autor" delay={10}>
          <Person name="Rafael" color={C.sky} delay={14} />
          <Person name="Gustavo" color={C.sky} delay={30} />
        </Col>
        <Entrance delay={50}><div style={{ fontSize: 90, color: C.teal, transform: V ? "rotate(90deg)" : undefined }}>➜</div></Entrance>
        <Col label="Aprovam" delay={60}>
          <Person name="Júlio" color={C.violet} delay={70} check />
          <Person name="Sidney" color={C.violet} delay={90} check />
        </Col>
      </AbsoluteFill>
    </Scene>
  );
};

const DIFF: [string, string][] = [
  ["  def solicitar(item):", C.dim],
  ["-   total = qtd * 1", C.red],
  ["+   total = qtd * fator", C.ok],
  ["+   registrar(item, total)", C.ok],
];

export const I4: React.FC = () => {
  const { V, w, h } = useV();
  const frame = useCurrentFrame();
  return (
    <Scene>
      <Title k="Passos 5 a 7" t="De homolog até o servidor" />
      <Flow
        x={V ? 140 : 60} y={V ? 360 : 230} w={V ? w - 280 : w - 120} h={V ? 1000 : 300}
        nodes={[
          { title: "branch", color: C.teal, at: 4, icon: "🌿" },
          { title: "homolog", color: C.warn, at: 20, icon: "🧪", sub: "homologado" },
          { title: "development", color: C.sky, at: 50, icon: "🛠️", sub: "após homologado" },
          { title: "main", color: C.violet, at: 160, icon: "🔒", sub: "após aprovação" },
          { title: "Servidor", color: C.ok, at: 265, icon: "🖥️", sub: "linha por linha" },
        ]}
        edges={[{ label: "PR", at: 12 }, { label: "PR", at: 30 }, { label: "PR", at: 140 }, { label: V ? "cópia manual" : "cópia", at: 245 }]}
      />
      <div style={{ position: "absolute", left: V ? 140 : 560, right: V ? 140 : 560, top: V ? 1420 : 620, opacity: interpolate(frame, [290, 305], [0, 1], clamp) }}>
        <div style={{ borderRadius: 18, background: "rgba(0,0,0,0.55)", border: `1px solid ${C.line}`, padding: "20px 28px", fontFamily: F.mono, fontSize: V ? 34 : 28 }}>
          {DIFF.map(([t, c], i) => {
            const a = 300 + i * 26;
            return <div key={i} style={{ color: c, whiteSpace: "pre", opacity: interpolate(frame, [a, a + 8], [0, 1], clamp), background: i === 1 || i === 2 || i === 3 ? `${c}14` : undefined, padding: "2px 6px" }}>{t}</div>;
          })}
        </div>
      </div>
    </Scene>
  );
};

export const I5: React.FC = () => {
  const { V, w } = useV();
  return (
    <Scene>
      <Title k="Fluxo atual" t="Solicitação de Materiais" />
      <Flow
        x={V ? 140 : 200} y={V ? 400 : 270} w={V ? w - 280 : w - 400} h={V ? 900 : 300}
        nodes={[
          { title: "bugfix | feature", color: C.teal, at: 10, icon: "🌿", sub: "branch" },
          { title: "homolog", color: C.warn, at: 60, icon: "🧪", sub: "testes · homologação" },
          { title: "Servidor", color: C.ok, at: 140, icon: "🖥️", sub: "linha por linha" },
        ]}
        edges={[{ label: "PR", at: 40 }, { label: "homologado", at: 120 }]}
      />
      <Entrance delay={170} y={20} style={{ position: "absolute", left: 0, right: 0, bottom: V ? 380 : 190, textAlign: "center" }}>
        <span style={{ display: "inline-block", padding: "14px 34px", borderRadius: 999, background: "rgba(251,191,36,0.14)", border: `2px solid ${C.warn}`, color: C.warn, fontFamily: F.body, fontWeight: 800, fontSize: V ? 36 : 30 }}>Direto: sem development e main</span>
      </Entrance>
    </Scene>
  );
};

export const I6: React.FC = () => {
  const { V } = useV();
  const p = useSpring(2, "bouncy");
  return (
    <Scene>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", rowGap: 30 }}>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: V ? 380 : 320, lineHeight: 1, transform: `scale(${interpolate(p, [0, 1], [0.6, 1])})`, opacity: Math.min(1, p), filter: "drop-shadow(0 0 60px rgba(139,92,246,0.45))" }}><GradText>IIA</GradText></div>
        <WordReveal text={V ? "Inovação, Inteligência Artificial e Automação" : "Inovação · Inteligência Artificial · Automação"} size={V ? 52 : 44} delay={20} align="center" highlight={["Inovação", "Inteligência", "Artificial", "Automação"]} />
        <div style={{ marginTop: 10 }}><AiBadge delay={60} big /></div>
      </AbsoluteFill>
    </Scene>
  );
};
