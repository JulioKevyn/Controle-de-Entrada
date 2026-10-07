import React from "react";
import { AbsoluteFill, interpolate, useCurrentFrame } from "remotion";
import { AiBadge, C, Chip, Counter, Entrance, F, GradText, Label, Mark, Scene, WordReveal, breathe, clamp, float, panel, useSpring } from "./components";
import { theme } from "./theme";

const typed = (text: string, frame: number, from: number, cps = 1.6) =>
  text.slice(0, Math.max(0, Math.min(text.length, Math.floor((frame - from) * cps))));

const Left: React.FC<{ top?: number; width?: number; children: React.ReactNode }> = ({ top = 330, width = 700, children }) => (
  <div style={{ position: "absolute", left: 130, top, width }}>{children}</div>
);
const Right: React.FC<{ top?: number; left?: number; width?: number; children: React.ReactNode; delay?: number }> = ({ top = 170, left = 900, width = 900, children, delay = 4 }) => (
  <Entrance delay={delay} x={80} y={0} style={{ position: "absolute", left, top, width }}>{children}</Entrance>
);

// 1 ---------- Abertura ----------
export const S1: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <Scene>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", rowGap: 34 }}>
        <Mark size={170} />
        <Entrance delay={16} y={30}>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 150, letterSpacing: "-0.03em", lineHeight: 1, transform: `translateY(${float(frame, 3)}px)` }}>
            <span style={{ color: C.text }}>Future </span><GradText>PDV</GradText>
          </div>
        </Entrance>
        <Entrance delay={34} y={20}>
          <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 38, color: C.muted }}>Vendas, estoque e gestão da sua loja na nuvem</div>
        </Entrance>
        <div style={{ marginTop: 10 }}><AiBadge delay={56} big /></div>
      </AbsoluteFill>
    </Scene>
  );
};

// 2 ---------- Problema ----------
const CHAOS = [
  ["planilha_vendas_FINAL_v7.xlsx", -4, 0, 0],
  ["caderno do caixa", 4, 220, 130],
  ["estoque ≠ prateleira", -3, 0, 250],
  ["custo? despesa? lucro?", 4, 210, 380],
] as const;
export const S2: React.FC = () => {
  const frame = useCurrentFrame();
  const q = useSpring(150, "bouncy");
  const flick = Math.floor(frame / 5) % 3;
  return (
    <Scene>
      <Left top={300}>
        <Label color={C.pink}>O problema</Label>
        <div style={{ marginTop: 18 }}><WordReveal text="No fim do mês, quanto sobrou?" delay={6} size={72} highlight={["sobrou?"]} /></div>
      </Left>
      <div style={{ position: "absolute", left: 960, top: 190, width: 860, height: 700 }}>
        {CHAOS.map(([t, r, x, y], i) => {
          const p = useSpring(24 + i * 18, "bouncy");
          return (
            <div key={t} style={{ position: "absolute", left: x, top: y, ...panel, borderRadius: 18, padding: "22px 30px", fontFamily: F.body, fontWeight: 700, fontSize: 30, color: C.muted, opacity: Math.min(1, p), transform: `rotate(${r * p}deg) scale(${interpolate(p, [0, 1], [0.6, 1])}) translateY(${float(frame + i * 12, 6, 26)}px)`, whiteSpace: "nowrap" }}>{t}</div>
          );
        })}
        <div style={{ position: "absolute", left: 60, top: 530, opacity: Math.min(1, q), transform: `scale(${interpolate(q, [0, 1], [0.5, 1])})` }}>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 120, color: C.pink, textShadow: `0 0 60px ${C.pink}88` }}>R$ {["???", "?,?", "???"][flick]}</div>
        </div>
      </div>
    </Scene>
  );
};

// 3 ---------- PDV ----------
const ITEMS = [["Camiseta Preta G", "1×", "89,90"], ["Bermuda Azul 42", "1×", "149,90"], ["Boné", "2×", "79,80"]];
export const S3: React.FC = () => {
  const frame = useCurrentFrame();
  const code = "7891234567895";
  const press = interpolate(frame, [200, 204, 210], [1, 0.94, 1], clamp);
  const done = useSpring(212, "bouncy");
  const payHi = frame > 150;
  return (
    <Scene>
      <Left><Label>PDV</Label><div style={{ marginTop: 18 }}><WordReveal text="Venda em segundos." delay={6} size={80} highlight={["segundos."]} /></div></Left>
      <Right top={120}>
        <div style={{ ...panel, padding: "34px 40px" }}>
          <div style={{ display: "flex", alignItems: "center", columnGap: 16, height: 70, borderRadius: 16, background: "rgba(255,255,255,0.06)", border: `1px solid ${C.line}`, padding: "0 24px", fontFamily: F.body, fontSize: 30, color: C.text, fontVariantNumeric: "tabular-nums", marginBottom: 22 }}>
            <span style={{ color: C.teal }}>▮▯▮▮▯</span>{typed(code, frame, 70, 0.6)}<span style={{ opacity: Math.floor(frame / 8) % 2, color: C.teal }}>|</span>
          </div>
          {ITEMS.map(([n, q, v], i) => {
            const p = useSpring(96 + i * 18, "snappy");
            return (
              <div key={n} style={{ opacity: Math.min(1, p), transform: `translateX(${interpolate(p, [0, 1], [30, 0])}px)`, display: "flex", justifyContent: "space-between", alignItems: "center", height: 70, borderBottom: `1px solid ${C.line}`, fontFamily: F.body, fontSize: 30, color: C.text }}>
                <span><span style={{ color: C.muted, marginRight: 14 }}>{q}</span>{n}</span><span style={{ fontVariantNumeric: "tabular-nums", fontWeight: 700 }}>R$ {v}</span>
              </div>
            );
          })}
          <div style={{ display: "flex", columnGap: 12, margin: "24px 0" }}>
            {["PIX", "Crédito", "Débito", "Dinheiro"].map((m, i) => {
              const on = payHi && (i === 0 || i === 1);
              return <div key={m} style={{ flex: 1, textAlign: "center", padding: "14px 0", borderRadius: 14, fontFamily: F.body, fontWeight: 700, fontSize: 26, color: on ? C.bg : C.muted, background: on ? C.teal : "rgba(255,255,255,0.06)", border: `1px solid ${on ? C.teal : C.line}` }}>{m}</div>;
            })}
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ fontFamily: F.display, fontSize: 44, color: C.text }}>Total <GradText>R$ 319,60</GradText></div>
            {done < 0.05 ? (
              <div style={{ padding: "16px 34px", borderRadius: 16, background: theme.grad, color: C.bg, fontFamily: F.body, fontWeight: 800, fontSize: 28, transform: `scale(${press})` }}>Finalizar venda</div>
            ) : (
              <div style={{ opacity: done, transform: `scale(${done})` }}><Chip color={C.ok} size={26}>✓ Venda concluída · recibo</Chip></div>
            )}
          </div>
        </div>
      </Right>
    </Scene>
  );
};

// 4 ---------- Caixa ----------
const CASH = [["08:57", "Abertura do caixa", "Maria", C.teal], ["13:10", "Sangria · Banco", "Maria", C.warn], ["15:42", "Venda registrada", "PIN 04", C.sky], ["18:30", "Fechamento", "Maria", C.pink]] as const;
export const S4: React.FC = () => {
  const frame = useCurrentFrame();
  const pin = Math.floor(interpolate(frame, [150, 180], [0, 4], clamp));
  return (
    <Scene>
      <Left><Label>Caixa e equipe</Label><div style={{ marginTop: 18 }}><WordReveal text="O caixa fecha certo." delay={6} size={80} highlight={["certo."]} /></div></Left>
      <Right top={150}>
        <div style={{ ...panel, padding: "30px 40px" }}>
          {CASH.map(([h, t, w, c], i) => {
            const p = useSpring(24 + i * 38, "snappy");
            return (
              <div key={t} style={{ display: "flex", alignItems: "center", columnGap: 24, height: 104, borderBottom: i < 3 ? `1px solid ${C.line}` : undefined, opacity: Math.min(1, p), transform: `translateX(${interpolate(p, [0, 1], [30, 0])}px)` }}>
                <div style={{ width: 16, height: 16, borderRadius: 8, background: c, boxShadow: `0 0 20px ${c}` }} />
                <div style={{ fontFamily: F.display, fontSize: 30, color: C.muted, width: 130 }}>{h}</div>
                <div style={{ flex: 1, fontFamily: F.body, fontWeight: 700, fontSize: 32, color: C.text }}>{t}</div>
                <Chip color={c}>{w}</Chip>
              </div>
            );
          })}
          <div style={{ display: "flex", alignItems: "center", columnGap: 20, marginTop: 24 }}>
            <span style={{ fontFamily: F.body, fontWeight: 700, fontSize: 26, color: C.muted }}>PIN do vendedor</span>
            {[0, 1, 2, 3].map((k) => <div key={k} style={{ width: 22, height: 22, borderRadius: 11, background: k < pin ? C.teal : "rgba(255,255,255,0.15)", boxShadow: k < pin ? `0 0 16px ${C.teal}` : undefined }} />)}
          </div>
        </div>
      </Right>
    </Scene>
  );
};

// 5 ---------- Estoque ----------
const VARS = [["Preta · M", 5], ["Preta · G", 7], ["Branca · M", 4], ["Branca · G", 1]] as const;
export const S5: React.FC = () => {
  const frame = useCurrentFrame();
  const sold = frame > 100;
  const alert = useSpring(215, "bouncy");
  return (
    <Scene>
      <Left><Label>Estoque</Label><div style={{ marginTop: 18 }}><WordReveal text="Por cor e por tamanho." delay={6} size={76} highlight={["tamanho."]} /></div></Left>
      <Right top={170}>
        <div style={{ ...panel, padding: "34px 40px" }}>
          <div style={{ fontFamily: F.display, fontSize: 36, color: C.text, marginBottom: 6 }}>Camiseta</div>
          <div style={{ fontFamily: F.body, fontSize: 24, color: C.muted, marginBottom: 22 }}>4 variações</div>
          {VARS.map(([n, q], i) => {
            const p = useSpring(24 + i * 12, "snappy");
            const qty = n === "Preta · G" && sold ? 6 : q;
            const hit = n === "Preta · G" && frame > 100 && frame < 190;
            const low = n === "Branca · G";
            return (
              <div key={n} style={{ display: "flex", alignItems: "center", columnGap: 22, height: 84, borderRadius: 14, padding: "0 20px", marginBottom: 10, background: hit ? `${C.teal}1F` : "rgba(255,255,255,0.04)", border: `1px solid ${hit ? C.teal : C.line}`, opacity: Math.min(1, p), transform: `translateX(${interpolate(p, [0, 1], [30, 0])}px)` }}>
                <div style={{ width: 190, fontFamily: F.body, fontWeight: 700, fontSize: 30, color: C.text }}>{n}</div>
                <div style={{ flex: 1, height: 14, borderRadius: 7, background: "rgba(255,255,255,0.08)", overflow: "hidden" }}>
                  <div style={{ width: `${(qty / 8) * 100}%`, height: "100%", borderRadius: 7, background: low ? C.warn : theme.grad }} />
                </div>
                <div style={{ width: 70, textAlign: "right", fontFamily: F.display, fontSize: 36, color: low ? C.warn : C.text, fontVariantNumeric: "tabular-nums" }}>{qty}</div>
              </div>
            );
          })}
          <div style={{ marginTop: 14, height: 56, opacity: Math.min(1, alert), transform: `scale(${alert})`, transformOrigin: "left center" }}><Chip color={C.warn} size={26}>⚠ Branca · G está acabando</Chip></div>
        </div>
      </Right>
    </Scene>
  );
};

// 6 ---------- Lucro ----------
const MONTHS = [["Abr", 52, 38, 14], ["Mai", 60, 42, 18], ["Jun", 82, 55, 27], ["Jul", 66, 47, 19], ["Ago", 78, 52, 26], ["Set", 72, 49, 23]] as const;
export const S6: React.FC = () => (
  <Scene>
    <Left><Label>Lucro de verdade</Label><div style={{ marginTop: 18 }}><WordReveal text="Quanto sobrou, não só quanto vendeu." delay={6} size={60} highlight={["sobrou,"]} /></div></Left>
    <Right top={140}>
      <div style={{ ...panel, padding: "34px 40px" }}>
        <div style={{ display: "flex", columnGap: 26, marginBottom: 22 }}>
          {[["Faturamento", C.sky], ["Custos e despesas", C.pink], ["Lucro líquido", C.teal]].map(([l, c]) => (
            <div key={l} style={{ display: "flex", alignItems: "center", columnGap: 10, fontFamily: F.body, fontWeight: 700, fontSize: 22, color: C.muted }}><div style={{ width: 14, height: 14, borderRadius: 4, background: c }} />{l}</div>
          ))}
        </div>
        <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", height: 420 }}>
          {MONTHS.map(([m, a, b, c], i) => {
            const g = (d: number) => useSpring(20 + i * 8 + d, "smooth");
            const g1 = g(0), g2 = g(4), g3 = g(8);
            return (
              <div key={m} style={{ width: 120, textAlign: "center" }}>
                <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "center", columnGap: 6, height: 380 }}>
                  {[[a, C.sky, g1], [b, C.pink, g2], [c, C.teal, g3]].map(([v, col, gg], k) => (
                    <div key={k} style={{ width: 30, height: `${(v as number) * 4.2 * (gg as number)}px`, borderRadius: "8px 8px 2px 2px", background: col as string, boxShadow: k === 2 ? `0 0 24px ${C.teal}88` : undefined }} />
                  ))}
                </div>
                <div style={{ marginTop: 10, fontFamily: F.body, fontWeight: 700, fontSize: 24, color: C.muted }}>{m}</div>
              </div>
            );
          })}
        </div>
        <div style={{ marginTop: 18, display: "flex", alignItems: "baseline", columnGap: 16 }}>
          <span style={{ fontFamily: F.body, fontWeight: 700, fontSize: 26, color: C.muted }}>Margem</span>
          <GradText style={{ fontFamily: F.display, fontWeight: 800, fontSize: 64 }}><Counter to={38} delay={120} />%</GradText>
        </div>
      </div>
    </Right>
  </Scene>
);

// 7 ---------- Horários ----------
const HOURS = [20, 28, 35, 42, 38, 55, 72, 96, 100, 84, 48, 30];
export const S7: React.FC = () => {
  const pk = useSpring(100, "bouncy");
  return (
    <Scene>
      <Left><Label>Dados da sua loja</Label><div style={{ marginTop: 18 }}><WordReveal text="Quando você mais vende." delay={6} size={72} highlight={["vende."]} /></div></Left>
      <Right top={200}>
        <div style={{ ...panel, padding: "34px 40px" }}>
          <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", height: 340 }}>
            {HOURS.map((h, i) => {
              const g = useSpring(20 + i * 5, "smooth");
              const peak = i >= 7 && i <= 9;
              return (
                <div key={i} style={{ width: 52, textAlign: "center" }}>
                  <div style={{ height: `${h * 3 * g}px`, borderRadius: "10px 10px 3px 3px", background: peak ? theme.grad : "rgba(255,255,255,0.16)", boxShadow: peak ? `0 0 30px ${C.teal}66` : undefined }} />
                  <div style={{ marginTop: 10, fontFamily: F.body, fontWeight: 700, fontSize: 18, color: C.dim }}>{9 + i}h</div>
                </div>
              );
            })}
          </div>
          <div style={{ marginTop: 22, opacity: Math.min(1, pk), transform: `scale(${pk})`, transformOrigin: "left center" }}><Chip color={C.teal} size={26}>Sábado 15h–18h é o seu pico</Chip></div>
        </div>
      </Right>
    </Scene>
  );
};

// 8 ---------- IA: chat ----------
const Bubble: React.FC<{ me?: boolean; children: React.ReactNode; delay: number }> = ({ me, children, delay }) => {
  const p = useSpring(delay, "snappy");
  return (
    <div style={{ display: "flex", justifyContent: me ? "flex-end" : "flex-start", marginBottom: 18, opacity: Math.min(1, p), transform: `translateY(${interpolate(p, [0, 1], [24, 0])}px)` }}>
      <div style={{ maxWidth: 640, padding: "20px 28px", borderRadius: me ? "24px 24px 6px 24px" : "24px 24px 24px 6px", background: me ? theme.grad : "rgba(255,255,255,0.08)", color: me ? C.bg : C.text, border: me ? undefined : `1px solid ${C.line}`, fontFamily: F.body, fontWeight: me ? 800 : 600, fontSize: 29, lineHeight: 1.4 }}>{children}</div>
    </div>
  );
};
export const S8: React.FC = () => {
  const frame = useCurrentFrame();
  const a1 = "Foram 142 vendas. O mais vendido foi a Camiseta, principalmente no tamanho G, e ela já está com estoque baixo.";
  const a2 = "Sim: o Boné está há mais de 30 dias sem vender e tem 14 unidades. Que tal um combo com a camiseta, que é a que mais sai?";
  const dots = (f0: number, f1: number) => frame >= f0 && frame < f1;
  return (
    <Scene>
      <Left top={250}>
        <AiBadge delay={4} big />
        <div style={{ marginTop: 26 }}><WordReveal text="Pergunte pra sua loja." delay={14} size={68} highlight={["loja."]} /></div>
        <Entrance delay={40} y={20}><div style={{ marginTop: 22, fontFamily: F.body, fontWeight: 600, fontSize: 30, color: C.muted }}>Em português, com os números reais do seu negócio.</div></Entrance>
      </Left>
      <Right top={90} width={940} left={880}>
        <div style={{ ...panel, padding: "30px 36px", minHeight: 860 }}>
          <div style={{ display: "flex", alignItems: "center", columnGap: 14, paddingBottom: 18, marginBottom: 22, borderBottom: `1px solid ${C.line}`, fontFamily: F.display, fontSize: 26, color: C.text }}>
            <span style={{ color: C.pink, fontSize: 30 }}>✦</span> Assistente da loja
          </div>
          {frame >= 110 && <Bubble me delay={110}>Quanto vendi essa semana?</Bubble>}
          {dots(150, 180) && <Bubble delay={150}><span style={{ letterSpacing: 6, color: C.teal }}>{"•".repeat(1 + (Math.floor(frame / 6) % 3))}</span></Bubble>}
          {frame >= 180 && <Bubble delay={180}>{typed(a1, frame, 180, 1.4)}</Bubble>}
          {frame >= 232 && <Bubble me delay={232}>Tem alguma coisa parada?</Bubble>}
          {dots(262, 285) && <Bubble delay={262}><span style={{ letterSpacing: 6, color: C.teal }}>{"•".repeat(1 + (Math.floor(frame / 6) % 3))}</span></Bubble>}
          {frame >= 285 && <Bubble delay={285}>{typed(a2, frame, 285, 1.3)}</Bubble>}
        </div>
      </Right>
    </Scene>
  );
};

// 9 ---------- IA: recursos + WhatsApp ----------
const AIF = [
  ["Insights automáticos", "Resumo do período com pontos de atenção, sem você perguntar.", 10],
  ["O que repor primeiro", "Cruza o que mais vende com o que está acabando.", 82],
  ["Ranking preditivo", "Os produtos com mais chance de vender.", 133],
] as const;
export const S9: React.FC = () => {
  const frame = useCurrentFrame();
  const ph = useSpring(214, "smooth");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 120 }}>
        <AiBadge delay={2} />
        <div style={{ marginTop: 20 }}><WordReveal text="A IA trabalha por você." delay={8} size={64} highlight={["IA"]} /></div>
      </div>
      <div style={{ position: "absolute", left: 130, top: 330, width: 1000 }}>
        {AIF.map(([t, d, at], i) => (
          <Entrance key={t} delay={at} y={30} style={{ marginBottom: 22 }}>
            <div style={{ ...panel, padding: "26px 34px", display: "flex", alignItems: "center", columnGap: 26, transform: `translateY(${float(frame + i * 14, 4, 28)}px)` }}>
              <div style={{ width: 70, height: 70, borderRadius: 20, background: "linear-gradient(120deg, #7C3AED, #FF2E9A)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 36, color: "#fff" }}>✦</div>
              <div><div style={{ fontFamily: F.display, fontSize: 32, color: C.text }}>{t}</div><div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 25, color: C.muted, marginTop: 6 }}>{d}</div></div>
            </div>
          </Entrance>
        ))}
      </div>
      <div style={{ position: "absolute", left: 1240, top: 190, width: 540, opacity: Math.min(1, ph), transform: `translateX(${interpolate(ph, [0, 1], [70, 0])}px) scale(${interpolate(ph, [0, 1], [0.92, 1])})` }}>
        <div style={{ ...panel, borderRadius: 44, padding: "28px 26px", minHeight: 720, border: `2px solid ${C.line}` }}>
          <div style={{ display: "flex", alignItems: "center", columnGap: 12, paddingBottom: 16, marginBottom: 18, borderBottom: `1px solid ${C.line}`, fontFamily: F.body, fontWeight: 800, fontSize: 26, color: C.ok }}>● WhatsApp · Future PDV</div>
          {[["Resumo do dia: 38 vendas, ticket médio em alta. ✦", 236], ["⚠ Estoque baixo: Camiseta Branca G (1 un.)", 270], ["Qual produto está parado?", 304, true]].map(([t, at, me]) => (
            <div key={t as string} style={{ display: "flex", justifyContent: me ? "flex-end" : "flex-start", marginBottom: 16, opacity: Math.min(1, Math.max(0, (frame - (at as number)) / 8)), transform: `translateY(${interpolate(frame, [(at as number), (at as number) + 10], [16, 0], clamp)}px)` }}>
              <div style={{ maxWidth: 400, padding: "16px 22px", borderRadius: 20, background: me ? `${C.ok}33` : "rgba(255,255,255,0.08)", border: `1px solid ${me ? C.ok + "66" : C.line}`, fontFamily: F.body, fontWeight: 600, fontSize: 24, color: C.text, lineHeight: 1.35 }}>{t}</div>
            </div>
          ))}
        </div>
      </div>
    </Scene>
  );
};

// 10 ---------- Filiais ----------
const BRANCH = ["Matriz", "Shopping", "Praia"];
export const S10: React.FC = () => {
  const frame = useCurrentFrame();
  const sel = frame < 130 ? 0 : frame < 200 ? 1 : 2;
  return (
    <Scene>
      <Left top={300}><Label>Multi-filial</Label><div style={{ marginTop: 18 }}><WordReveal text="Todas as lojas, uma conta." delay={6} size={68} highlight={["uma", "conta."]} /></div></Left>
      <Right top={130}>
        <div style={{ display: "flex", columnGap: 18, marginBottom: 22 }}>
          {BRANCH.map((b, i) => {
            const on = i === sel;
            const p = useSpring(24 + i * 10, "bouncy");
            return (
              <div key={b} style={{ flex: 1, ...panel, padding: "26px 24px", opacity: Math.min(1, p), transform: `scale(${interpolate(p, [0, 1], [0.85, 1])})`, border: `2px solid ${on ? C.teal : C.line}`, boxShadow: on ? `0 0 40px ${C.teal}44` : undefined }}>
                <div style={{ fontFamily: F.display, fontSize: 30, color: on ? C.teal : C.text }}>{b}</div>
                <div style={{ marginTop: 14, fontFamily: F.body, fontWeight: 600, fontSize: 22, color: C.muted }}>Caixa · Estoque · Equipe</div>
                <div style={{ marginTop: 16, height: 10, borderRadius: 5, background: "rgba(255,255,255,0.1)" }}><div style={{ width: `${[78, 54, 66][i]}%`, height: "100%", borderRadius: 5, background: theme.grad }} /></div>
              </div>
            );
          })}
        </div>
        <Entrance delay={250} y={30}><div style={{ ...panel, padding: "22px 30px", marginBottom: 16, display: "flex", alignItems: "center", columnGap: 18, fontFamily: F.body, fontWeight: 700, fontSize: 28, color: C.text }}><Chip color={C.ok}>✓</Chip> Backup automático dos dados</div></Entrance>
        <Entrance delay={320} y={30}><div style={{ ...panel, padding: "22px 30px", fontFamily: F.body, fontWeight: 700, fontSize: 26, color: C.text }}>
          <div style={{ color: C.muted, marginBottom: 10, fontSize: 22 }}>REGISTRO DE AUDITORIA</div>
          <div>Preço alterado · <span style={{ color: C.teal }}>por Ana</span></div>
          <div style={{ marginTop: 6 }}>Estoque ajustado · <span style={{ color: C.warn }}>com autorização do dono</span></div>
        </div></Entrance>
      </Right>
    </Scene>
  );
};

// 11 ---------- Clientes ----------
const QUOTES = [
  ["Boutique Malu", "O que mais ajudou foi conseguir ver o lucro de cada produto, porque antes eu só olhava o valor das vendas."],
  ["Bella Store", "A parte de lucro e estoque ajuda muito nas compras, principalmente pra não ficar colocando dinheiro em produto que não gira."],
  ["Empório da Praça", "Consigo acompanhar as vendas e entender melhor onde está meu lucro sem precisar ficar fazendo conta toda hora."],
];
export const S11: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 110 }}><Label>Quem usa</Label><div style={{ marginTop: 14 }}><WordReveal text="Aprovado por lojistas." delay={6} size={68} highlight={["lojistas."]} /></div></div>
      {QUOTES.map(([n, q], i) => (
        <Entrance key={n} delay={30 + i * 22} y={50} style={{ position: "absolute", left: 130 + i * 596, top: 340, width: 570 }}>
          <div style={{ ...panel, padding: "34px 34px", minHeight: 470, transform: `translateY(${float(frame + i * 15, 5, 28)}px)` }}>
            <div style={{ color: C.warn, fontSize: 32, letterSpacing: 4 }}>★★★★★</div>
            <div style={{ marginTop: 18, fontFamily: F.body, fontWeight: 600, fontSize: 30, lineHeight: 1.45, color: C.text }}>“{q}”</div>
            <div style={{ marginTop: 26, fontFamily: F.display, fontSize: 26, color: C.teal }}>{n}</div>
            <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 22, color: C.muted, marginTop: 4 }}>Cliente do Future PDV</div>
          </div>
        </Entrance>
      ))}
    </Scene>
  );
};

// 12 ---------- Planos ----------
const PLANS = [
  { n: "Básico", p: 100, f: ["PDV e controle de caixa", "Estoque por cor e tamanho", "Dashboard de BI", "Relatórios e auditoria"], hot: false },
  { n: "Pro", p: 150, f: ["Tudo do Básico", "Importar e exportar CSV", "Histórico de estoque", "Comissões por vendedor"], hot: false },
  { n: "Master IA", p: 350, f: ["Tudo do Pro", "Chat com a IA da loja", "Insights e ranking preditivo", "Alertas no WhatsApp"], hot: true },
];
export const S12: React.FC = () => {
  const frame = useCurrentFrame();
  const free = useSpring(150, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 0, right: 0, top: 90, display: "flex", justifyContent: "center" }}><WordReveal text="Comece grátis. Cresça quando quiser." delay={4} size={56} align="center" highlight={["grátis."]} /></div>
      {PLANS.map((pl, i) => (
        <Entrance key={pl.n} delay={20 + i * 14} y={60} style={{ position: "absolute", left: 130 + i * 600, top: pl.hot ? 250 : 290, width: 540 }}>
          <div style={{ ...panel, padding: "36px 36px", minHeight: pl.hot ? 600 : 540, border: `2px solid ${pl.hot ? C.pink : C.line}`, boxShadow: pl.hot ? `0 0 70px ${C.pink}55` : undefined, transform: `scale(${pl.hot ? breathe(frame, 0.008) : 1})` }}>
            {pl.hot ? <AiBadge delay={50} /> : <div style={{ height: 40 }} />}
            <div style={{ fontFamily: F.display, fontSize: 40, color: C.text, marginTop: 14 }}>{pl.n}</div>
            <div style={{ display: "flex", alignItems: "baseline", columnGap: 8, margin: "10px 0 22px" }}>
              <span style={{ fontFamily: F.body, fontWeight: 700, fontSize: 28, color: C.muted }}>R$</span>
              <GradText style={{ fontFamily: F.display, fontWeight: 800, fontSize: 84 }}>{pl.p}</GradText>
              <span style={{ fontFamily: F.body, fontWeight: 700, fontSize: 26, color: C.muted }}>/mês</span>
            </div>
            {pl.f.map((x) => <div key={x} style={{ fontFamily: F.body, fontWeight: 600, fontSize: 26, color: C.text, marginBottom: 12 }}><span style={{ color: C.teal, marginRight: 12 }}>✓</span>{x}</div>)}
          </div>
        </Entrance>
      ))}
      <div style={{ position: "absolute", left: 0, right: 0, bottom: 90, display: "flex", justifyContent: "center", opacity: Math.min(1, free), transform: `scale(${free})` }}><Chip color={C.ok} size={30}>7 dias grátis · sem cartão · sem fidelidade</Chip></div>
    </Scene>
  );
};

// 13 ---------- CTA ----------
export const S13: React.FC = () => {
  const frame = useCurrentFrame();
  const btn = useSpring(110, "bouncy");
  return (
    <Scene>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", rowGap: 34 }}>
        <Mark size={130} />
        <WordReveal text="Sua loja merece saber quanto realmente lucra." delay={14} size={70} align="center" highlight={["lucra."]} />
        <div style={{ opacity: Math.min(1, btn), transform: `scale(${btn * breathe(frame, 0.012)})`, padding: "22px 60px", borderRadius: 999, background: theme.grad, color: C.bg, fontFamily: F.display, fontWeight: 800, fontSize: 40, boxShadow: `0 20px 60px -10px ${C.teal}88` }}>Testar 7 dias grátis</div>
        <Entrance delay={150} y={20}><AiBadge big /></Entrance>
        <Entrance delay={170} y={20}><div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 28, color: C.muted }}>Future PDV · desenvolvido por Julio Clemente · (11) 96620-9914</div></Entrance>
      </AbsoluteFill>
    </Scene>
  );
};
