import React from "react";
import { AbsoluteFill, interpolate, useCurrentFrame, useVideoConfig } from "remotion";
import { Chip, Counter, OrangeCover, Entrance, Label, Mark, Scene, WordReveal, breathe, float, useSpring } from "./components";
import { theme } from "./theme";

const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;
const F = theme.fonts;
const C = theme.colors;

const panel: React.CSSProperties = {
  background: C.card,
  border: `1px solid ${C.line}`,
  borderRadius: 28,
  boxShadow: "0 30px 60px -20px rgba(35,31,32,0.28)",
};


// ---------- 1. Abertura ----------
const Pill: React.FC<{ delay: number; children: React.ReactNode }> = ({ delay, children }) => (
  <Entrance delay={delay} y={20}>
    <div style={{ background: "#fff", color: C.primary, fontFamily: F.display, fontSize: 40, padding: "18px 56px", borderRadius: 999, boxShadow: "0 20px 40px -16px rgba(35,31,32,0.35)" }}>{children}</div>
  </Entrance>
);

export const S1: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <Scene>
      <OrangeCover />
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 40 }}>
        <Mark size={150} variant="white" shadow={false} />
        <div style={{ transform: `translateY(${float(frame, 3)}px)`, marginTop: 10 }}>
          <WordReveal text="Solicitação de Materiais" delay={30} size={132} align="center" highlight={["Materiais"]} color="#fff" hlColor={C.text} />
        </div>
        <Pill delay={74}>Sistema de controle de pedidos e envios</Pill>
      </AbsoluteFill>
    </Scene>
  );
};

// ---------- 2. Problema ----------
const SHEETS = ["NIVEA", "DANONE", "BAYER", "MONDELEZ", "PERNOD RICARD", "OUTROS"];
const SheetCard: React.FC<{ i: number; name: string }> = ({ i, name }) => {
  const frame = useCurrentFrame();
  const p = useSpring(18 + i * 6, "bouncy");
  const bad = useSpring(233 + i * 8, "bouncy");
  const rot = ((i * 37) % 13) - 6;
  const pos = [[0, 0], [300, 30], [600, 0], [60, 230], [360, 250], [660, 240]][i];
  const cols = 3 + (i % 3);
  return (
    <div style={{ position: "absolute", left: pos[0], top: pos[1], width: 250, height: 170, ...panel, borderRadius: 16, padding: 14, opacity: p, transform: `translateY(${float(frame + i * 9, 6, 26)}px) rotate(${rot * p}deg) scale(${interpolate(p, [0, 1], [0.6, 1])})` }}>
      <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 16, color: C.textDim, letterSpacing: "0.1em", marginBottom: 10 }}>{name}.xlsx</div>
      <div style={{ display: "grid", gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: 5 }}>
        {Array.from({ length: cols * 5 }).map((_, k) => (
          <div key={k} style={{ height: 12, borderRadius: 3, background: k < cols ? `${C.accent}88` : "rgba(35,31,32,0.09)" }} />
        ))}
      </div>
      <div style={{ position: "absolute", right: -14, top: -14, width: 40, height: 40, borderRadius: 20, background: C.primary, color: "#fff", fontFamily: F.display, fontSize: 24, display: "flex", alignItems: "center", justifyContent: "center", transform: `scale(${bad})`, boxShadow: "0 8px 20px -6px rgba(236,103,7,0.6)" }}>!</div>
    </div>
  );
};

const BigStat: React.FC<{ to: number; plus?: boolean; label: string; delay: number }> = ({ to, plus, label, delay }) => (
  <Entrance delay={delay} y={40} cfg="snappy">
    <div style={{ display: "flex", alignItems: "baseline", columnGap: 22 }}>
      <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 190, lineHeight: 1, color: C.primary, textShadow: `0 8px 40px ${C.glow}` }}>
        {plus ? "+" : ""}<Counter to={to} delay={delay} />
      </div>
      <div style={{ fontFamily: F.display, fontSize: 56, color: C.text }}>{label}</div>
    </div>
  </Entrance>
);

export const S2: React.FC = () => (
  <Scene>
    <div style={{ position: "absolute", left: 130, top: 190, width: 800 }}>
      <Label>O desafio</Label>
      <div style={{ marginTop: 20 }}>
        <WordReveal text="Cada cliente, um jeito de enviar." delay={6} size={72} highlight={["jeito"]} />
      </div>
      <div style={{ marginTop: 50, display: "flex", flexDirection: "column", rowGap: 14 }}>
        <BigStat to={16} label="layouts" delay={100} />
        <BigStat to={100} plus label="depositantes" delay={170} />
      </div>
    </div>
    <div style={{ position: "absolute", left: 960, top: 270, width: 960, height: 520 }}>
      {SHEETS.map((n, i) => <SheetCard key={n} i={i} name={n} />)}
    </div>
  </Scene>
);

// ---------- 3. Envio da planilha ----------
const FINALIDADES = ["Envio", "Descarte", "Coleta", "Transferência", "Retirada"];
const Field: React.FC<{ label: string; children: React.ReactNode; delay: number }> = ({ label, children, delay }) => (
  <Entrance delay={delay} y={24}>
    <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 20, letterSpacing: "0.12em", color: C.textDim, textTransform: "uppercase", marginBottom: 10 }}>{label}</div>
    {children}
  </Entrance>
);
const Input: React.FC<{ children?: React.ReactNode }> = ({ children }) => (
  <div style={{ height: 66, borderRadius: 14, background: "#F0EEEC", border: `1px solid ${C.line}`, display: "flex", alignItems: "center", padding: "0 22px", fontFamily: F.body, fontSize: 28, color: C.text, justifyContent: "space-between" }}>{children}</div>
);

export const S3: React.FC = () => {
  const frame = useCurrentFrame();
  const sel = frame > 70 ? 0 : -1;
  const client = "Nivea";
  const typed = client.slice(0, Math.floor(interpolate(frame, [95, 120], [0, client.length + 0.99], clamp)));
  const prog = interpolate(frame, [130, 168], [0, 1], { ...clamp, easing: theme.ease.inOut });
  const det = useSpring(178, "bouncy");
  const sso = useSpring(14, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 330, width: 640 }}>
        <Label>Envio de planilha</Label>
        <div style={{ marginTop: 18 }}>
          <WordReveal text="Tudo começa na planilha." delay={6} size={84} highlight={["planilha."]} />
        </div>
      </div>
      <Entrance delay={4} x={80} y={0} style={{ position: "absolute", left: 860, top: 170, width: 940 }}>
        <div style={{ ...panel, padding: "36px 44px", display: "flex", flexDirection: "column", rowGap: 26 }}>
          <div style={{ display: "flex", alignItems: "center", columnGap: 16, opacity: sso, transform: `scale(${interpolate(sso, [0, 1], [0.9, 1])})`, transformOrigin: "left center" }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 3, width: 34, height: 34 }}>
              {["#F25022", "#7FBA00", "#00A4EF", "#FFB900"].map((c) => <div key={c} style={{ background: c }} />)}
            </div>
            <span style={{ fontFamily: F.body, fontWeight: 600, fontSize: 28, color: C.text }}>Entrou com a conta Microsoft</span>
            <Chip color={C.ok}>✓ Acesso por perfil</Chip>
          </div>
          <Field label="Finalidade da solicitação" delay={34}>
            <div style={{ display: "flex", columnGap: 12 }}>
              {FINALIDADES.map((f, i) => (
                <div key={f} style={{ flex: 1, height: 70, borderRadius: 14, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: F.body, fontWeight: 700, fontSize: 23, color: i === sel ? "#fff" : C.text, background: i === sel ? C.primary : "#F0EEEC", border: `1px solid ${i === sel ? C.primary : C.line}`, boxShadow: i === sel ? `0 12px 24px -10px ${C.primary}` : undefined }}>{f}</div>
              ))}
            </div>
          </Field>
          <Field label="Cliente / Depositante" delay={60}>
            <Input>
              <span>{typed}<span style={{ display: "inline-block", width: 2, height: 30, background: C.text, marginLeft: 3, verticalAlign: "middle", opacity: frame < 123 ? Math.floor(frame / 8) % 2 : 0 }} /></span>
              <span style={{ color: C.textDim }}>⌄</span>
            </Input>
          </Field>
          <Field label="Planilha do pedido" delay={90}>
            <Input>
              <span>pedido_nivea.xlsx</span>
              <div style={{ width: 220, height: 10, borderRadius: 5, background: "rgba(35,31,32,0.1)", overflow: "hidden" }}>
                <div style={{ width: `${prog * 100}%`, height: "100%", background: `linear-gradient(90deg, ${C.primary}, ${C.primary2})` }} />
              </div>
            </Input>
          </Field>
          <div style={{ height: 56, transform: `scale(${det})`, opacity: det, transformOrigin: "left center" }}>
            <Chip color={C.ok}>✓ Template detectado: layout Nivea</Chip>
          </div>
        </div>
      </Entrance>
    </Scene>
  );
};

// ---------- 4. Validação antecipada ----------
const STEPS = [
  { t: "Identificar cliente / template", at: 30 },
  { t: "Extrair dados da planilha", at: 65 },
  { t: "Validar CNPJ/CPF na Receita Federal", at: 100 },
  { t: "Validar cadastro, endereço e estoque", at: 140 },
  { t: "Gerar orçamento / grade", at: 185 },
];
export const S4: React.FC = () => {
  const frame = useCurrentFrame();
  const done = STEPS.filter((s) => frame >= s.at).length;
  const pct = interpolate(frame, [20, 190], [0, 1], clamp);
  const note = useSpring(150, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 300, width: 700 }}>
        <Label>Validação antecipada</Label>
        <div style={{ marginTop: 18 }}>
          <WordReveal text="Descubra o erro antes de enviar." delay={6} size={80} highlight={["antes"]} />
        </div>
      </div>
      <Entrance delay={4} x={80} y={0} style={{ position: "absolute", left: 900, top: 190, width: 900 }}>
        <div style={{ ...panel, padding: "36px 44px" }}>
          <div style={{ fontFamily: F.display, fontSize: 38, color: C.text, marginBottom: 26 }}>Validando sua solicitação</div>
          <div style={{ height: 10, borderRadius: 5, background: "rgba(35,31,32,0.1)", overflow: "hidden", marginBottom: 28 }}>
            <div style={{ width: `${pct * 100}%`, height: "100%", background: `linear-gradient(90deg, ${C.primary}, ${C.primary2})` }} />
          </div>
          {STEPS.map((s, i) => {
            const ok = frame >= s.at;
            const active = i === done;
            const p = useSpring(14 + i * 5, "snappy");
            return (
              <div key={s.t} style={{ display: "flex", alignItems: "center", columnGap: 20, height: 72, opacity: p, borderBottom: i < STEPS.length - 1 ? `1px solid ${C.line}` : undefined }}>
                <div style={{ width: 38, height: 38, borderRadius: 19, border: `3px solid ${ok ? C.ok : active ? C.primary : C.line}`, background: ok ? C.ok : "transparent", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, fontWeight: 700, position: "relative" }}>
                  {ok ? "✓" : active ? <div style={{ position: "absolute", inset: -3, borderRadius: 19, border: `3px solid transparent`, borderTopColor: C.primary, transform: `rotate(${frame * 14}deg)` }} /> : null}
                </div>
                <span style={{ fontFamily: F.body, fontSize: 30, fontWeight: ok ? 600 : 400, color: ok ? C.text : C.textDim }}>{s.t}</span>
              </div>
            );
          })}
          <div style={{ marginTop: 26, height: 56, opacity: note, transform: `scale(${note})`, transformOrigin: "left center" }}>
            <Chip color={C.primary}>Nada é gravado até você enviar</Chip>
          </div>
        </div>
      </Entrance>
    </Scene>
  );
};

// ---------- 5. Quatro conferências ----------
const CHECKS = [
  { k: "CNPJ / CPF", d: "Situação cadastral na Receita Federal", ok: "Ativo", at: 98 },
  { k: "CEP e endereço", d: "Planilha × cadastro × Sintegra", ok: "Endereço confere", at: 222 },
  { k: "Saldo de estoque", d: "Por SKU, classe e depositante", ok: "Saldo disponível", at: 380 },
  { k: "Vínculo", d: "Destinatário × depositante", ok: "Vínculo cadastrado", at: 486 },
];
export const S5V: React.FC = () => {
  const frame = useCurrentFrame();
  const alt = useSpring(440, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 330, width: 640 }}>
        <Label>Validações</Label>
        <div style={{ marginTop: 18 }}>
          <WordReveal text="Quatro conferências por linha." delay={6} size={80} highlight={["Quatro"]} />
        </div>
      </div>
      <Entrance delay={4} x={80} y={0} style={{ position: "absolute", left: 840, top: 150, width: 960 }}>
        <div style={{ ...panel, padding: "18px 36px" }}>
          {CHECKS.map((c, i) => {
            const on = useSpring(c.at, "snappy");
            const ok = useSpring(c.at + 40, "bouncy");
            const active = frame >= c.at && frame < c.at + 120;
            return (
              <div key={c.k} style={{ display: "flex", alignItems: "center", columnGap: 24, height: 190, borderBottom: i < 3 ? `1px solid ${C.line}` : undefined, opacity: 0.35 + 0.65 * on }}>
                <div style={{ width: 84, height: 84, borderRadius: 42, background: active ? C.primary : `${C.primary}1F`, color: active ? "#fff" : C.primary, fontFamily: F.display, fontSize: 40, display: "flex", alignItems: "center", justifyContent: "center", transform: `scale(${interpolate(on, [0, 1], [0.8, 1])})` }}>{i + 1}</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontFamily: F.display, fontSize: 40, color: C.text }}>{c.k}</div>
                  <div style={{ fontFamily: F.body, fontSize: 26, color: C.textDim, marginTop: 6 }}>{c.d}</div>
                  {i === 2 && (
                    <div style={{ marginTop: 8, opacity: alt, transform: `translateY(${interpolate(alt, [0, 1], [8, 0])}px)`, fontFamily: F.body, fontWeight: 600, fontSize: 22, color: C.primary }}>Sem saldo? Sugere outra classe com estoque</div>
                  )}
                </div>
                <div style={{ opacity: ok, transform: `scale(${ok})`, transformOrigin: "right center" }}><Chip color={C.ok}>✓ {c.ok}</Chip></div>
              </div>
            );
          })}
        </div>
      </Entrance>
    </Scene>
  );
};

// ---------- 6. Correção guiada ----------
export const S5: React.FC = () => {
  const frame = useCurrentFrame();
  const OPTS = ["Manter o endereço da planilha", "Ajustar conforme o Sintegra", "Ajustar conforme o WebClient"];
  const sel = frame > 70 ? 1 : -1;
  const press = interpolate(frame, [112, 116, 122], [1, 0.93, 1], clamp);
  const okIn = useSpring(138, "bouncy");
  const card = interpolate(frame, [134, 146], [1, 0], { ...clamp, easing: theme.ease.in });
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 330, width: 700 }}>
        <Label>Correção guiada</Label>
        <div style={{ marginTop: 18 }}>
          <WordReveal text="O sistema pergunta. Você decide." delay={6} size={84} highlight={["decide."]} />
        </div>
      </div>
      <Entrance delay={4} x={80} y={0} style={{ position: "absolute", left: 900, top: 200, width: 900, opacity: undefined }}>
        <div style={{ ...panel, padding: "38px 44px", opacity: card, transform: `scale(${interpolate(card, [0, 1], [0.92, 1])})` }}>
          <div style={{ display: "flex", alignItems: "center", columnGap: 16, marginBottom: 8 }}>
            <div style={{ width: 44, height: 44, borderRadius: 22, background: C.primary, color: "#fff", fontFamily: F.display, fontSize: 28, display: "flex", alignItems: "center", justifyContent: "center" }}>!</div>
            <div style={{ fontFamily: F.display, fontSize: 34, color: C.text }}>Endereço divergente</div>
          </div>
          <div style={{ fontFamily: F.body, fontSize: 26, color: C.textDim, marginBottom: 26 }}>Planilha, WebClient e Sintegra trazem endereços diferentes. Qual usar?</div>
          {OPTS.map((o, i) => {
            const on = i === sel;
            const p = useSpring(24 + i * 8, "snappy");
            return (
              <div key={o} style={{ opacity: p, transform: `translateX(${interpolate(p, [0, 1], [24, 0])}px)`, display: "flex", alignItems: "center", columnGap: 18, height: 74, borderRadius: 14, padding: "0 24px", marginBottom: 12, background: on ? `${C.primary}12` : "#F0EEEC", border: `2px solid ${on ? C.primary : "transparent"}`, fontFamily: F.body, fontSize: 30, fontWeight: on ? 700 : 400, color: C.text }}>
                <div style={{ width: 28, height: 28, borderRadius: 14, border: `3px solid ${on ? C.primary : C.textDim}`, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  {on && <div style={{ width: 14, height: 14, borderRadius: 7, background: C.primary }} />}
                </div>
                {o}
              </div>
            );
          })}
          <div style={{ marginTop: 22, display: "flex", justifyContent: "flex-end" }}>
            <div style={{ padding: "16px 38px", borderRadius: 14, background: C.primary, color: "#fff", fontFamily: F.body, fontWeight: 700, fontSize: 28, transform: `scale(${press})`, boxShadow: `0 12px 30px ${C.glow}` }}>✓ Confirmar e revalidar</div>
          </div>
        </div>
      </Entrance>
      <AbsoluteFill style={{ left: 900, width: 900, top: 200, height: 640, alignItems: "center", justifyContent: "center", position: "absolute" }}>
        <div style={{ opacity: okIn, transform: `scale(${interpolate(okIn, [0, 1], [0.5, 1])})`, textAlign: "center" }}>
          <div style={{ width: 150, height: 150, borderRadius: 75, margin: "0 auto 26px", background: `${C.ok}1F`, border: `3px solid ${C.ok}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 84, color: C.ok }}>✓</div>
          <div style={{ fontFamily: F.display, fontSize: 52, color: C.text }}>Revalidado, sem pendências</div>
        </div>
      </AbsoluteFill>
    </Scene>
  );
};

// ---------- 6. Orçamento (Kanban) ----------
const COLS = ["Aguardando Orçamento", "Orçamento em Elaboração", "Aguardando Aprovação do Cliente", "Aprovado"];
export const S6: React.FC = () => {
  const frame = useCurrentFrame();
  const colW = 400, gap = 30, x0 = 115;
  const stops = [0, 0, 1, 1, 2, 2, 3];
  const t = [0, 140, 160, 245, 262, 300, 316];
  const pos = interpolate(frame, t, stops, { ...clamp, easing: theme.ease.inOut });
  const colNow = Math.round(pos);
  const hx = x0 + pos * (colW + gap) + 16;
  const mail = useSpring(268, "bouncy");
  const appr = useSpring(325, "bouncy");
  const pdf = useSpring(218, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 120 }}>
        <Label>Orçamento</Label>
        <div style={{ marginTop: 14 }}>
          <WordReveal text="Do pacote ao orçamento." delay={6} size={72} highlight={["orçamento."]} />
        </div>
      </div>
      {COLS.map((c, i) => (
        <Entrance key={c} delay={10 + i * 6} y={40} style={{ position: "absolute", left: x0 + i * (colW + gap), top: 310, width: colW }}>
          <div style={{ ...panel, height: 560, padding: 18, borderRadius: 22, background: i === 3 ? `${C.ok}0F` : "#FFFFFFB3" }}>
            <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 25, color: i === 3 ? C.ok : C.text, height: 64, lineHeight: 1.2 }}>{c}</div>
            {i < 3 && [0, 1].map((k) => (
              <div key={k} style={{ height: 100, borderRadius: 14, background: "#F0EEEC", marginBottom: 14, padding: 16, opacity: 0.8 }}>
                <div style={{ height: 14, width: "60%", borderRadius: 7, background: "rgba(35,31,32,0.14)", marginBottom: 12 }} />
                <div style={{ height: 12, width: "85%", borderRadius: 6, background: "rgba(35,31,32,0.08)" }} />
              </div>
            ))}
          </div>
        </Entrance>
      ))}
      <Entrance delay={26} y={30} style={{ position: "absolute", left: hx, top: 400, width: colW - 32 }}>
        <div style={{ ...panel, borderRadius: 16, padding: "18px 22px", border: `2px solid ${C.primary}`, boxShadow: `0 24px 50px -16px ${C.glow}, 0 30px 60px -20px rgba(35,31,32,0.3)`, transform: `translateY(${float(frame, 3, 24)}px) rotate(${(frame > 140 && frame < 160) || (frame > 245 && frame < 262) || (frame > 300 && frame < 316) ? -2 : 0}deg)` }}>
          <div style={{ fontFamily: F.display, fontSize: 28, color: C.text }}>Pacote · Job Nivea</div>
          <div style={{ fontFamily: F.body, fontSize: 21, color: C.textDim, marginTop: 6 }}>{colNow === 0 ? "Pacote criado após a validação" : colNow === 1 ? "Modalidades de frete preenchidas" : colNow === 2 ? "Cliente notificado por e-mail" : "Aprovado pelo cliente"}</div>
        </div>
      </Entrance>
      <div style={{ position: "absolute", left: x0 + 1 * (colW + gap) + 16, top: 640, opacity: pdf, transform: `scale(${pdf})`, transformOrigin: "left center" }}>
        <Chip color={C.primary}>PDF + planilha de fretes</Chip>
      </div>
      <div style={{ position: "absolute", left: x0 + 2 * (colW + gap) + 16, top: 640, opacity: mail, transform: `scale(${mail})`, transformOrigin: "left center" }}>
        <Chip color={C.primary}>✉ E-mail enviado ao cliente</Chip>
      </div>
      <div style={{ position: "absolute", left: x0 + 3 * (colW + gap) + 16, top: 520, opacity: appr, transform: `scale(${appr})`, transformOrigin: "left center" }}>
        <Chip color={C.ok}>✓ Segue para a operação</Chip>
      </div>
    </Scene>
  );
};

// ---------- 8. Aprovação do cliente ----------
const MODAL = [["Rodoviário Convencional", "R$ 1.280,00"], ["Aéreo Expresso", "R$ 2.940,00"], ["Carro Exclusivo", "R$ 3.610,00"]];
export const S8A: React.FC = () => {
  const frame = useCurrentFrame();
  const sel = frame > 70 ? 0 : -1;
  const press = interpolate(frame, [150, 154, 160], [1, 0.92, 1], clamp);
  const done = useSpring(166, "bouncy");
  const p1 = useSpring(186, "bouncy");
  const p2 = useSpring(206, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 330, width: 640 }}>
        <Label>Portal do cliente</Label>
        <div style={{ marginTop: 18 }}>
          <WordReveal text="Aprova ou recusa. Com motivo." delay={6} size={80} highlight={["Com", "motivo."]} />
        </div>
        <div style={{ marginTop: 40, display: "flex", columnGap: 16 }}>
          <div style={{ opacity: p1, transform: `scale(${p1})`, transformOrigin: "left center" }}><Chip color={C.primary}>Custos extras</Chip></div>
          <div style={{ opacity: p2, transform: `scale(${p2})`, transformOrigin: "left center" }}><Chip color={C.primary}>Manuseio tabelado</Chip></div>
        </div>
      </div>
      <Entrance delay={4} x={80} y={0} style={{ position: "absolute", left: 860, top: 170, width: 940 }}>
        <div style={{ ...panel, padding: "34px 44px" }}>
          <div style={{ fontFamily: F.display, fontSize: 36, color: C.text, marginBottom: 20 }}>Orçamento · Job Nivea</div>
          {MODAL.map(([m, v], i) => {
            const on = i === sel;
            const e = useSpring(26 + i * 10, "snappy");
            return (
              <div key={m} style={{ opacity: e, transform: `translateX(${interpolate(e, [0, 1], [24, 0])}px)`, display: "flex", alignItems: "center", justifyContent: "space-between", height: 78, borderRadius: 14, padding: "0 24px", marginBottom: 12, background: on ? `${C.primary}12` : "#F0EEEC", border: `2px solid ${on ? C.primary : "transparent"}`, fontFamily: F.body, fontSize: 30, color: C.text, fontWeight: on ? 700 : 400 }}>
                <span>{m}</span><span style={{ fontVariantNumeric: "tabular-nums" }}>{v}</span>
              </div>
            );
          })}
          <div style={{ display: "flex", justifyContent: "flex-end", columnGap: 16, marginTop: 22, height: 70, alignItems: "center" }}>
            {done < 0.05 ? (
              <>
                <div style={{ padding: "14px 34px", borderRadius: 14, background: "#F0EEEC", border: `1px solid ${C.line}`, color: C.text, fontFamily: F.body, fontWeight: 700, fontSize: 26 }}>Recusar</div>
                <div style={{ padding: "14px 38px", borderRadius: 14, background: C.primary, color: "#fff", fontFamily: F.body, fontWeight: 700, fontSize: 26, transform: `scale(${press})`, boxShadow: `0 12px 30px ${C.glow}` }}>Aprovar</div>
              </>
            ) : (
              <div style={{ transform: `scale(${done})`, opacity: done }}><Chip color={C.ok}>✓ Orçamento aprovado</Chip></div>
            )}
          </div>
        </div>
      </Entrance>
    </Scene>
  );
};

// ---------- 9. Operação ----------
const MODS = [
  ["Distribuição de Materiais", "Materiais por classe"],
  ["Gestão de Romaneios", "Pedidos por romaneio"],
  ["Operação de Coletas", "Da ordem ao recebimento"],
  ["Agenda de Retiradas", "Agendamento por WhatsApp"],
  ["Operação de Descarte", "Laudo e acompanhamento"],
  ["Transferência entre Bases", "Entre filiais"],
  ["Alertas de Vencimento", "Antes de vencer"],
  ["Roteirização", "Rotas e veículos"],
];
export const S7: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <Scene>
      <div style={{ position: "absolute", left: 0, right: 0, top: 130, display: "flex", justifyContent: "center" }}>
        <WordReveal text="Aprovado, a operação assume." delay={20} size={88} align="center" highlight={["operação"]} />
      </div>
      {MODS.map(([t, d], i) => {
        const col = i % 4, row = Math.floor(i / 4);
        return (
          <Entrance key={t} delay={[70, 112, 150, 186, 250, 285, 320, 345][i]} y={40} style={{ position: "absolute", left: 130 + col * 420, top: 380 + row * 220, width: 390 }}>
            <div style={{ ...panel, height: 190, borderRadius: 22, padding: "26px 28px", position: "relative", overflow: "hidden", transform: `translateY(${float(frame + i * 13, 4, 28)}px)` }}>
              <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: 8, background: `linear-gradient(180deg, ${C.primary}, ${C.primary2})` }} />
              <div style={{ fontFamily: F.display, fontSize: 31, lineHeight: 1.15, color: C.text, marginLeft: 8 }}>{t}</div>
              <div style={{ fontFamily: F.body, fontSize: 24, color: C.textDim, marginTop: 12, marginLeft: 8 }}>{d}</div>
            </div>
          </Entrance>
        );
      })}
    </Scene>
  );
};

// ---------- 10. Resultados ----------
const BarRow: React.FC<{ label: string; from: number; w: number; color: string; text: string; sub: string }> = ({ label, from, w, color, text, sub }) => {
  const frame = useCurrentFrame();
  const grow = interpolate(frame, [from, from + 36], [0, 1], { ...clamp, easing: theme.ease.inOut });
  const t = useSpring(from + 30, "bouncy");
  return (
    <div style={{ marginBottom: 34 }}>
      <div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 24, letterSpacing: "0.14em", textTransform: "uppercase", color: C.textDim, marginBottom: 10, opacity: Math.min(1, grow * 4) }}>{label}</div>
      <div style={{ display: "flex", alignItems: "center", columnGap: 26 }}>
        <div style={{ height: 78, width: Math.max(0.001, w * grow), borderRadius: 16, background: color, boxShadow: color === C.primary ? `0 14px 34px -10px ${C.primary}` : undefined }} />
        <div style={{ opacity: t, transform: `scale(${interpolate(t, [0, 1], [0.7, 1])})`, transformOrigin: "left center", whiteSpace: "nowrap" }}>
          <div style={{ fontFamily: F.display, fontSize: 60, lineHeight: 1, color: C.text }}>{text}</div>
          <div style={{ fontFamily: F.body, fontSize: 22, color: C.textDim, marginTop: 4 }}>{sub}</div>
        </div>
      </div>
    </div>
  );
};

export const S10R: React.FC = () => {
  const frame = useCurrentFrame();
  const pct = useSpring(214, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 120 }}>
        <Label>Resultados</Label>
        <div style={{ marginTop: 14 }}>
          <WordReveal text="Menos tempo. Mais pedidos." delay={6} size={72} highlight={["Mais", "pedidos."]} />
        </div>
      </div>
      <div style={{ position: "absolute", left: 130, top: 340, width: 1130 }}>
        <BarRow label="Antes" from={58} w={640} color="#B9B3AE" text="4h 30min" sub="processo manual" />
        <BarRow label="Agora" from={142} w={64} color={C.primary} text="27 min" sub="com a ferramenta" />
      </div>
      <div style={{ position: "absolute", left: 1330, top: 330, width: 470, textAlign: "center", opacity: Math.min(1, pct), transform: `scale(${interpolate(pct, [0, 1], [0.6, 1])})` }}>
        <div style={{ fontFamily: F.display, fontSize: 210, lineHeight: 1, color: C.primary, textShadow: `0 10px 50px ${C.glow}` }}>−90%</div>
        <div style={{ fontFamily: F.display, fontSize: 40, color: C.text, marginTop: 10 }}>no tempo do processo</div>
      </div>
      {[["pacotes", 1904, 291], ["destinos", 27787, 322]].map(([l, n, d], i) => (
        <Entrance key={l as string} delay={d as number} y={40} style={{ position: "absolute", left: 130 + i * 640, top: 760 }}>
          <div style={{ ...panel, padding: "26px 36px", display: "flex", alignItems: "baseline", columnGap: 20, whiteSpace: "nowrap", transform: `translateY(${float(frame + i * 14, 4, 28)}px)` }}>
            <div style={{ fontFamily: F.display, fontSize: 84, lineHeight: 1, color: C.text, fontVariantNumeric: "tabular-nums" }}>
              <Counter to={n as number} delay={d as number} />
            </div>
            <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 32, color: C.textDim, whiteSpace: "nowrap" }}>{l} processados</div>
          </div>
        </Entrance>
      ))}
    </Scene>
  );
};

// ---------- 11. Encerramento ----------
export const S8: React.FC = () => {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();
  const out = interpolate(frame, [durationInFrames - 14, durationInFrames - 1], [1, 0], { ...clamp, easing: theme.ease.in });
  const o = useSpring(0, "smooth");
  return (
    <AbsoluteFill style={{ opacity: out }}>
      <AbsoluteFill style={{ opacity: o }}><OrangeCover /></AbsoluteFill>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 44 }}>
        <Mark size={170} variant="white" shadow={false} delay={4} />
        <div style={{ transform: `scale(${breathe(frame, 0.008)})` }}>
          <WordReveal text="Fazendo marcas venderem mais." delay={22} size={92} align="center" highlight={["venderem", "mais."]} color="#fff" hlColor={C.text} />
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
};
