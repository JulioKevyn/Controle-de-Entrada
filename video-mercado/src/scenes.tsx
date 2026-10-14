import React from "react";
import { AbsoluteFill, Img, interpolate, staticFile, useCurrentFrame, useVideoConfig } from "remotion";
import { AiBadge, C, Chip, clamp, Entrance, F, GradText, panel, Scene, useSpring, useV, WordReveal } from "./components";
import { theme } from "./theme";

const useAt = () => { const { durationInFrames: d } = useVideoConfig(); return (f: number) => Math.round(d * f); };
const ease = (x: number) => 1 - Math.pow(1 - Math.min(1, Math.max(0, x)), 3);

const Logo: React.FC<{ size?: number }> = ({ size = 64 }) => (
  <div style={{ display: "inline-flex", alignItems: "center", columnGap: size * 0.28 }}>
    <div style={{ width: size, height: size, borderRadius: size * 0.28, background: theme.grad, display: "flex", alignItems: "center", justifyContent: "center", fontSize: size * 0.6, color: "#06101C" }}>⚡</div>
    <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: size * 0.85, color: C.text, letterSpacing: "-0.02em" }}>Future<GradText>PDV</GradText></div>
  </div>
);

// Moldura de celular; `s` reduz o conjunto (usado no formato horizontal)
const Phone: React.FC<{ children: React.ReactNode; w?: number; h?: number; delay?: number; s?: number }> = ({ children, w = 640, h = 1180, delay = 0, s = 1 }) => {
  const p = useSpring(delay, "smooth");
  return (
    <div style={{ width: w * s, height: h * s }}>
      <div style={{ width: w, height: h, transformOrigin: "top left", transform: `scale(${s})` }}>
        <div style={{ width: w, height: h, borderRadius: 70, background: "#0b1224", border: "10px solid #1b2440", boxShadow: "0 60px 120px -30px rgba(0,0,0,0.9), 0 0 80px rgba(31,224,200,0.15)", overflow: "hidden", opacity: Math.min(1, p), transform: `translateY(${(1 - p) * 120}px) scale(${0.92 + 0.08 * p})`, position: "relative", boxSizing: "border-box" }}>
          <div style={{ position: "absolute", top: 16, left: "50%", marginLeft: -60, width: 120, height: 30, borderRadius: 20, background: "#000", zIndex: 5 }} />
          {children}
        </div>
      </div>
    </div>
  );
};

const Tag: React.FC<{ delay: number; children: React.ReactNode; color?: string; style?: React.CSSProperties }> = ({ delay, children, color = C.teal, style }) => {
  const p = useSpring(delay, "bouncy");
  return <div style={{ opacity: Math.min(1, p), transform: `scale(${0.7 + 0.3 * p})`, padding: "14px 28px", borderRadius: 999, background: "rgba(7,11,22,0.9)", border: `2px solid ${color}`, color: C.text, fontFamily: F.body, fontWeight: 700, fontSize: 32, boxShadow: `0 10px 40px -8px ${color}88`, whiteSpace: "nowrap", ...style }}>{children}</div>;
};

type TagItem = { d: number; c?: string; t: string; v: React.CSSProperties };
// Layout padrão de cena: vertical = título no topo + visual + tags posicionadas; horizontal = texto à esquerda, visual à direita
const Layout: React.FC<{ k: string; t: string; tags?: TagItem[]; vTop: number; hGap?: number; children: React.ReactNode; leftW?: number }> = ({ k, t, tags = [], vTop, children, leftW = 700 }) => {
  const { V } = useV();
  if (V) {
    return (
      <>
        <div style={{ position: "absolute", top: 130, left: 60, right: 60, textAlign: "center" }}>
          <Entrance delay={0} y={20}><div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 28, letterSpacing: "0.26em", color: C.teal, textTransform: "uppercase" }}>{k}</div></Entrance>
          <Entrance delay={6} y={24}><div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 64, color: C.text, marginTop: 14, lineHeight: 1.12 }}>{t}</div></Entrance>
        </div>
        <AbsoluteFill style={{ alignItems: "center", paddingTop: vTop }}>{children}</AbsoluteFill>
        {tags.map((x, i) => <Tag key={i} delay={x.d} color={x.c} style={{ position: "absolute", ...x.v }}>{x.t}</Tag>)}
      </>
    );
  }
  return (
    <AbsoluteFill style={{ flexDirection: "row", alignItems: "center", justifyContent: "center", padding: "0 100px 60px", columnGap: 90 }}>
      <div style={{ width: leftW, display: "flex", flexDirection: "column", rowGap: 22 }}>
        <Entrance delay={0} y={20}><div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 26, letterSpacing: "0.26em", color: C.teal, textTransform: "uppercase" }}>{k}</div></Entrance>
        <Entrance delay={6} y={24}><div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 66, color: C.text, lineHeight: 1.12 }}>{t}</div></Entrance>
        <div style={{ display: "flex", flexDirection: "column", rowGap: 18, alignItems: "flex-start", marginTop: 14 }}>
          {tags.map((x, i) => <Tag key={i} delay={x.d} color={x.c} style={{ fontSize: 30 }}>{x.t}</Tag>)}
        </div>
      </div>
      <div style={{ flex: 1, display: "flex", justifyContent: "center", alignItems: "center" }}>{children}</div>
    </AbsoluteFill>
  );
};

export const M1: React.FC = () => {
  const frame = useCurrentFrame();
  const { V } = useV();
  const shelves = (
    <div style={{ marginTop: V ? 90 : 0, width: 860 }}>
      {[0, 1, 2].map((r) => (
        <Entrance key={r} delay={36 + r * 10} y={40}>
          <div style={{ position: "relative", height: 150, marginBottom: 46 }}>
            <div style={{ position: "absolute", bottom: 0, left: 0, right: 0, height: 16, borderRadius: 8, background: "linear-gradient(180deg,#3b4a6b,#1d2740)" }} />
            <div style={{ position: "absolute", bottom: 16, left: 20, display: "flex", gap: 18 }}>
              {Array.from({ length: 6 }).map((_, i) => (
                <div key={i} style={{ width: 100, height: 70 + ((i * 29 + r * 17) % 60), borderRadius: 12, background: [C.teal, C.sky, C.pink, C.warn, C.ok, C.violet][(i + r) % 6] + "cc", border: "2px solid rgba(255,255,255,0.18)" }} />
              ))}
            </div>
            <div style={{ position: "absolute", right: 10, top: -6, width: 74, height: 74, borderRadius: "50%", background: "#0b1224", border: `3px solid ${C.warn}`, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: F.display, fontWeight: 800, fontSize: 44, color: C.warn, transform: `scale(${1 + 0.08 * Math.sin((frame + r * 9) / 6)})` }}>?</div>
          </div>
        </Entrance>
      ))}
    </div>
  );
  return (
    <Scene>
      {V ? (
        <AbsoluteFill style={{ alignItems: "center", paddingTop: 300 }}>
          <WordReveal text="Quanto o seu mercado lucra" size={78} delay={4} align="center" highlight={["lucra"]} />
          <WordReveal text="em cada prateleira?" size={78} delay={24} align="center" highlight={["prateleira?"]} />
          {shelves}
        </AbsoluteFill>
      ) : (
        <AbsoluteFill style={{ flexDirection: "row", alignItems: "center", justifyContent: "center", padding: "0 100px 60px", columnGap: 90 }}>
          <div style={{ width: 760 }}>
            <WordReveal text="Quanto o seu mercado lucra em cada prateleira?" size={92} delay={4} highlight={["lucra", "prateleira?"]} />
          </div>
          {shelves}
        </AbsoluteFill>
      )}
    </Scene>
  );
};

export const M2: React.FC = () => {
  const at = useAt();
  const { V } = useV();
  const k = V ? 0.62 : 0.9;
  const devices = (
    <div style={{ display: "flex", gap: 26, alignItems: "flex-end", marginTop: V ? 70 : 0 }}>
      {[{ w: 240, h: 440, n: "Celular" }, { w: 420, h: 560, n: "Tablet" }, { w: 540, h: 360, n: "Computador" }].map((d, i) => (
        <Entrance key={d.n} delay={at(0.4) + i * 14} y={60}>
          <div style={{ width: d.w * k, height: d.h * k, borderRadius: 28, background: "#0b1224", border: "6px solid #232d4d", boxShadow: "0 30px 70px -20px rgba(0,0,0,0.9)", position: "relative", overflow: "hidden" }}>
            <div style={{ position: "absolute", inset: 14, borderRadius: 14, background: "linear-gradient(160deg, rgba(31,224,200,0.25), rgba(236,72,153,0.2))" }} />
            {[0, 1, 2, 3].map((r) => <div key={r} style={{ position: "absolute", left: 26, right: 26, top: 40 + r * 44, height: 22, borderRadius: 8, background: "rgba(255,255,255,0.15)" }} />)}
          </div>
          <div style={{ textAlign: "center", marginTop: 14, fontFamily: F.body, fontWeight: 800, fontSize: 28, color: C.muted }}>{d.n}</div>
        </Entrance>
      ))}
    </div>
  );
  return (
    <Scene>
      {V ? (
        <AbsoluteFill style={{ alignItems: "center", paddingTop: 280, rowGap: 36 }}>
          <Entrance delay={4}><Logo size={92} /></Entrance>
          <div style={{ marginTop: 20 }}><WordReveal text="O caixa do seu mercado," size={74} delay={20} align="center" /></div>
          <WordReveal text="rápido de verdade." size={74} delay={44} align="center" highlight={["rápido", "de", "verdade."]} />
          {devices}
        </AbsoluteFill>
      ) : (
        <AbsoluteFill style={{ flexDirection: "row", alignItems: "center", justifyContent: "center", padding: "0 100px 60px", columnGap: 80 }}>
          <div style={{ width: 800, display: "flex", flexDirection: "column", rowGap: 34 }}>
            <Entrance delay={4}><Logo size={96} /></Entrance>
            <WordReveal text="O caixa do seu mercado, rápido de verdade." size={72} delay={20} highlight={["rápido", "de", "verdade."]} />
          </div>
          {devices}
        </AbsoluteFill>
      )}
    </Scene>
  );
};

export const M3: React.FC = () => {
  const frame = useCurrentFrame();
  const at = useAt();
  const { V } = useV();
  const peso = ease((frame - at(0.22)) / 40) * 0.64;
  const scan = interpolate(frame % 60, [0, 30, 60], [0, 1, 0]);
  const pay = [{ n: "PIX", c: C.ok }, { n: "Dinheiro", c: C.warn }, { n: "Cartão", c: C.sky }];
  return (
    <Scene>
      <Layout k="Caixa" t="Venda por peso ou unidade" vTop={470}
        tags={[{ d: at(0.3), t: "Etiqueta de balança lida", v: { left: 40, top: 1100 } }, { d: at(0.5), c: C.sky, t: "Câmera do celular", v: { right: 40, top: 1280 } }]}>
        <Phone w={640} h={1060} delay={6} s={V ? 1 : 0.84}>
          <div style={{ padding: "70px 28px 0", fontFamily: F.body }}>
            <div style={{ ...panel, padding: 24, borderRadius: 22 }}>
              <div style={{ color: C.muted, fontWeight: 700, fontSize: 22 }}>Hortifruti · por quilo</div>
              <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 38, color: C.text, marginTop: 6 }}>Tomate italiano</div>
              <div style={{ marginTop: 18, padding: 18, borderRadius: 16, background: "rgba(255,255,255,0.06)" }}>
                <div style={{ color: C.muted, fontWeight: 700, fontSize: 20 }}>Peso</div>
                <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 86, color: C.text, lineHeight: 1.05 }}>{peso.toFixed(3).replace(".", ",")}<span style={{ fontSize: 34, color: C.muted }}> kg</span></div>
                <div style={{ color: C.muted, fontWeight: 700, fontSize: 24 }}>R$ 8,90 / kg</div>
              </div>
              <div style={{ marginTop: 14, padding: "14px 18px", borderRadius: 14, background: "rgba(31,224,200,0.14)", display: "flex", justifyContent: "space-between", fontWeight: 800, fontSize: 26 }}>
                <span style={{ color: C.teal }}>Valor do item</span><span style={{ color: C.teal }}>R$ {(peso * 8.9).toFixed(2).replace(".", ",")}</span>
              </div>
            </div>
            <div style={{ marginTop: 22, display: "flex", gap: 14, alignItems: "center" }}>
              <div style={{ flex: 1, height: 92, borderRadius: 16, background: "#fff", position: "relative", overflow: "hidden", display: "flex", gap: 5, padding: "14px 16px" }}>
                {Array.from({ length: 36 }).map((_, i) => <div key={i} style={{ width: 3 + (i % 4), background: "#111", height: "100%" }} />)}
                <div style={{ position: "absolute", left: 0, right: 0, top: `${scan * 90}%`, height: 5, background: "#ff3b6b", boxShadow: "0 0 16px #ff3b6b" }} />
              </div>
            </div>
            <div style={{ marginTop: 22, display: "flex", gap: 12 }}>
              {pay.map((p, i) => <Entrance key={p.n} delay={at(0.62) + i * 8} y={20}><Chip color={p.c} size={26}>{p.n}</Chip></Entrance>)}
            </div>
            <Entrance delay={at(0.75)} y={30}>
              <div style={{ marginTop: 26, padding: "26px 0", borderRadius: 20, background: theme.grad, textAlign: "center", fontFamily: F.display, fontWeight: 800, fontSize: 40, color: "#06101C" }}>Cobrar R$ 68,87</div>
            </Entrance>
          </div>
        </Phone>
      </Layout>
    </Scene>
  );
};

export const M4: React.FC = () => {
  const frame = useCurrentFrame();
  const at = useAt();
  const { V } = useV();
  const lots = [{ n: "Lote 1", s: "vence primeiro", c: C.warn }, { n: "Lote 2", s: "vence depois", c: C.sky }, { n: "Lote 3", s: "vence por último", c: C.ok }];
  const sold = ease((frame - at(0.62)) / 30);
  return (
    <Scene>
      <Layout k="Estoque" t="Validade por lote" vTop={500} tags={[{ d: at(0.8), c: C.warn, t: "12 itens vencendo em 7 dias", v: { left: 120, top: 1430 } }]}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", rowGap: 28 }}>
          <div style={{ display: "flex", gap: 20, alignItems: "flex-end" }}>
            {["Fardo", "Caixa", "Unidade"].map((u, i) => (<Entrance key={u} delay={10 + i * 10}><Chip color={C.teal} size={34}>{u}</Chip></Entrance>))}
          </div>
          <Entrance delay={at(0.14)}><div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 34, color: C.muted }}>baixa automática no estoque</div></Entrance>
          <div style={{ marginTop: V ? 40 : 14, width: 880, display: "flex", flexDirection: "column", gap: 22 }}>
            {lots.map((l, i) => (
              <Entrance key={l.n} delay={at(0.28) + i * 14}>
                <div style={{ ...panel, display: "flex", alignItems: "center", justifyContent: "space-between", padding: "28px 36px", borderRadius: 24, border: `2px solid ${i === 0 && sold > 0 ? l.c : C.line}`, boxShadow: i === 0 && sold > 0 ? `0 0 50px ${l.c}55` : undefined, transform: i === 0 ? `translateX(${sold * 40}px)` : undefined }}>
                  <div>
                    <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 40, color: C.text }}>{l.n}</div>
                    <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 28, color: l.c }}>{l.s}</div>
                  </div>
                  <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 30, color: i === 0 ? C.warn : C.muted }}>{i === 0 ? (sold > 0.5 ? "SAIU PRIMEIRO ✔" : "ALERTA ⚠") : ""}</div>
                </div>
              </Entrance>
            ))}
          </div>
        </div>
      </Layout>
    </Scene>
  );
};

export const M5: React.FC = () => {
  const at = useAt();
  const { V } = useV();
  const cards = [
    { i: "📒", t: "Fiado", s: "R$ 640 a receber", c: C.warn },
    { i: "🚚", t: "Entrada de mercadoria", s: "nota → estoque, custo, validade e margem", c: C.sky },
    { i: "🗑️", t: "Perdas e quebras", s: "vencidos, quebras, furtos e consumo", c: C.pink },
  ];
  return (
    <Scene>
      {V ? (
        <>
          <div style={{ position: "absolute", top: 130, left: 60, right: 60, textAlign: "center" }}>
            <Entrance delay={0} y={20}><div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 28, letterSpacing: "0.26em", color: C.teal, textTransform: "uppercase" }}>Tudo no mesmo lugar</div></Entrance>
            <Entrance delay={6} y={24}><div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 64, color: C.text, marginTop: 14 }}>Fiado, compras e perdas</div></Entrance>
          </div>
          <AbsoluteFill style={{ alignItems: "center", paddingTop: 520, rowGap: 34 }}>
            {cards.map((c, i) => (
              <Entrance key={c.t} delay={at(0.16) + i * at(0.2)} y={50}>
                <div style={{ ...panel, width: 900, padding: "36px 44px", borderRadius: 30, border: `2px solid ${c.c}`, display: "flex", alignItems: "center", columnGap: 34, boxShadow: `0 0 50px ${c.c}33` }}>
                  <div style={{ fontSize: 84 }}>{c.i}</div>
                  <div>
                    <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 46, color: c.c }}>{c.t}</div>
                    <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 32, color: C.muted, marginTop: 8, lineHeight: 1.2 }}>{c.s}</div>
                  </div>
                </div>
              </Entrance>
            ))}
          </AbsoluteFill>
        </>
      ) : (
        <AbsoluteFill style={{ alignItems: "center", paddingTop: 130 }}>
          <Entrance delay={0} y={20}><div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 26, letterSpacing: "0.26em", color: C.teal, textTransform: "uppercase" }}>Tudo no mesmo lugar</div></Entrance>
          <Entrance delay={6} y={24}><div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 70, color: C.text, marginTop: 14 }}>Fiado, compras e perdas</div></Entrance>
          <div style={{ display: "flex", gap: 40, marginTop: 70 }}>
            {cards.map((c, i) => (
              <Entrance key={c.t} delay={at(0.16) + i * at(0.2)} y={50}>
                <div style={{ ...panel, width: 520, height: 400, padding: "40px 40px", borderRadius: 30, border: `2px solid ${c.c}`, display: "flex", flexDirection: "column", rowGap: 18, boxShadow: `0 0 50px ${c.c}33` }}>
                  <div style={{ fontSize: 96 }}>{c.i}</div>
                  <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 42, color: c.c, lineHeight: 1.1 }}>{c.t}</div>
                  <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 30, color: C.muted, lineHeight: 1.25 }}>{c.s}</div>
                </div>
              </Entrance>
            ))}
          </div>
        </AbsoluteFill>
      )}
    </Scene>
  );
};

export const M6: React.FC = () => {
  const frame = useCurrentFrame();
  const at = useAt();
  const months = ["J", "F", "M", "A", "M", "J"];
  const fat = [0.55, 0.62, 0.6, 0.72, 0.78, 0.9];
  const cus = [0.35, 0.4, 0.38, 0.44, 0.47, 0.52];
  return (
    <Scene>
      <Layout k="Análise de lucro" t="Mês a mês" vTop={540} tags={[{ d: at(0.7), c: C.ok, t: "faturamento · custo · despesas", v: { left: 140, top: 1280 } }]}>
        <div style={{ ...panel, width: 900, height: 640, padding: "40px 44px", borderRadius: 34, position: "relative" }}>
          <div style={{ display: "flex", gap: 30, fontFamily: F.body, fontWeight: 800, fontSize: 26 }}>
            <span style={{ color: C.teal }}>■ Faturamento</span><span style={{ color: C.pink }}>■ Custo</span><span style={{ color: C.ok }}>■ Lucro</span>
          </div>
          <div style={{ position: "absolute", left: 44, right: 44, bottom: 60, height: 440, display: "flex", alignItems: "flex-end", justifyContent: "space-between" }}>
            {months.map((m, i) => {
              const p = ease((frame - 14 - i * 8) / 26);
              return (
                <div key={i} style={{ width: 120, display: "flex", flexDirection: "column", alignItems: "center", gap: 10 }}>
                  <div style={{ display: "flex", gap: 8, alignItems: "flex-end", height: 400 }}>
                    <div style={{ width: 34, height: fat[i] * 400 * p, borderRadius: "8px 8px 0 0", background: C.teal }} />
                    <div style={{ width: 34, height: cus[i] * 400 * p, borderRadius: "8px 8px 0 0", background: C.pink }} />
                    <div style={{ width: 34, height: (fat[i] - cus[i]) * 400 * p, borderRadius: "8px 8px 0 0", background: C.ok }} />
                  </div>
                  <div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 28, color: C.muted }}>{m}</div>
                </div>
              );
            })}
          </div>
        </div>
      </Layout>
    </Scene>
  );
};

const Bubble: React.FC<{ me?: boolean; delay: number; children: React.ReactNode }> = ({ me, delay, children }) => (
  <Entrance delay={delay} y={30}>
    <div style={{ display: "flex", justifyContent: me ? "flex-end" : "flex-start" }}>
      <div style={{ maxWidth: 520, padding: "22px 28px", borderRadius: 26, borderBottomRightRadius: me ? 6 : 26, borderBottomLeftRadius: me ? 26 : 6, background: me ? "linear-gradient(120deg,#0e9f8e,#1b6bd6)" : "rgba(255,255,255,0.09)", fontFamily: F.body, fontWeight: 700, fontSize: 32, color: "#fff", lineHeight: 1.25 }}>{children}</div>
    </div>
  </Entrance>
);

export const M7: React.FC = () => {
  const at = useAt();
  const { V } = useV();
  return (
    <Scene>
      <Layout k="Plano Master IA" t="Pergunte para o seu mercado" vTop={500}
        tags={[{ d: at(0.72), c: C.pink, t: "produtos parados há 30 dias", v: { left: 40, top: 1480 } }, { d: at(0.82), c: C.ok, t: "no sistema ou no WhatsApp", v: { right: 40, top: 1400 } }]}>
        <Phone w={640} h={1000} delay={4} s={V ? 1 : 0.86}>
          <div style={{ padding: "76px 26px 0", display: "flex", flexDirection: "column", gap: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 30, color: C.text }}>Assistente IA</div>
              <AiBadge delay={10} />
            </div>
            <Bubble me delay={at(0.2)}>O que eu preciso repor primeiro?</Bubble>
            <Bubble delay={at(0.34)}>Repor primeiro: arroz e leite.</Bubble>
            <Bubble me delay={at(0.5)}>Tem algo vencendo?</Bubble>
            <Bubble delay={at(0.62)}>12 itens vencendo, com R$ 372 em risco.</Bubble>
          </div>
        </Phone>
      </Layout>
    </Scene>
  );
};

export const M8: React.FC = () => {
  const frame = useCurrentFrame();
  const at = useAt();
  const { V } = useV();
  const pan = interpolate(frame, [0, at(1)], [0, -260], clamp);
  const k = V ? 1 : 1.12;
  return (
    <Scene>
      <Layout k="Controle" t="Mais controle, menos prejuízo" vTop={520} leftW={620}
        tags={[{ d: at(0.3), c: C.warn, t: "🔒 Preço mínimo com senha do gerente", v: { left: 40, top: 1330 } }, { d: at(0.55), c: C.sky, t: "Quem fez cada venda, registrado", v: { right: 40, top: 1430 } }]}>
        <Entrance delay={6} y={60}>
          <div style={{ width: 920 * k, height: 760 * (V ? 1 : 0.95), borderRadius: 30, overflow: "hidden", border: `3px solid ${C.line}`, boxShadow: "0 50px 100px -30px rgba(0,0,0,0.9)", position: "relative", transform: "perspective(1400px) rotateX(4deg)" }}>
            <Img src={staticFile("img/inicio.jpg")} style={{ position: "absolute", width: 1500 * k, left: (-230 + pan * 0.6) * k, top: (-30 + pan * 0.4) * k }} />
          </div>
        </Entrance>
      </Layout>
    </Scene>
  );
};

export const M9: React.FC = () => {
  const frame = useCurrentFrame();
  const { V } = useV();
  const pb = useSpring(60, "bouncy");
  return (
    <Scene>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", rowGap: V ? 36 : 28, paddingBottom: V ? 0 : 50 }}>
        <Entrance delay={0}><Logo size={V ? 104 : 96} /></Entrance>
        <WordReveal text="Teste 7 dias grátis" size={V ? 84 : 92} delay={14} align="center" highlight={["grátis"]} />
        <Entrance delay={34}><div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 40, color: C.muted }}>sem cartão · sem fidelidade</div></Entrance>
        <div style={{ transform: `scale(${pb * (1 + Math.sin(frame / 7) * 0.02)})`, opacity: Math.min(1, pb), marginTop: 20 }}>
          <div style={{ padding: "30px 70px", borderRadius: 999, background: theme.grad, fontFamily: F.display, fontWeight: 800, fontSize: 46, color: "#06101C", boxShadow: "0 20px 70px -10px rgba(31,224,200,0.55)" }}>Criar meu mercado grátis</div>
        </div>
        <Entrance delay={100}><div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 36, color: C.teal, marginTop: 10 }}>pdv-do-futuro.onrender.com/mercado</div></Entrance>
      </AbsoluteFill>
    </Scene>
  );
};
