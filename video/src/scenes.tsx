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
  const bad = useSpring(150 + i * 12, "bouncy");
  const rot = ((i * 37) % 13) - 6;
  const pos = [[0, 0], [310, 30], [620, 0], [100, 230], [410, 250], [720, 240]][i];
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

export const S2: React.FC = () => (
  <Scene>
    <div style={{ position: "absolute", left: 130, top: 250, width: 760 }}>
      <Label>O desafio</Label>
      <div style={{ marginTop: 20 }}>
        <WordReveal text="Um layout para cada cliente." delay={6} size={84} highlight={["cada", "cliente."]} />
      </div>
      <div style={{ marginTop: 44, display: "flex", flexDirection: "column", rowGap: 18 }}>
        {["CNPJ está ativo?", "O endereço confere?", "Tem estoque?"].map((t, i) => (
          <Entrance key={t} delay={120 + i * 40} x={-30} y={0} style={{ alignSelf: "flex-start" }}>
            <div style={{ ...panel, borderRadius: 18, padding: "16px 32px", fontFamily: F.body, fontWeight: 600, fontSize: 36, color: C.text, display: "flex", alignItems: "center", columnGap: 18 }}>
              <span style={{ color: C.primary, fontFamily: F.display, fontSize: 40 }}>?</span>{t}
            </div>
          </Entrance>
        ))}
      </div>
    </div>
    <div style={{ position: "absolute", left: 900, top: 270, width: 960, height: 520 }}>
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
  const sel = frame > 50 ? 0 : -1;
  const client = "Nivea";
  const typed = client.slice(0, Math.floor(interpolate(frame, [80, 105], [0, client.length + 0.99], clamp)));
  const prog = interpolate(frame, [130, 175], [0, 1], { ...clamp, easing: theme.ease.inOut });
  const det = useSpring(185, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 330, width: 640 }}>
        <Label>Envio de planilha</Label>
        <div style={{ marginTop: 18 }}>
          <WordReveal text="Tudo começa na planilha." delay={6} size={84} highlight={["planilha."]} />
        </div>
      </div>
      <Entrance delay={4} x={80} y={0} style={{ position: "absolute", left: 860, top: 170, width: 940 }}>
        <div style={{ ...panel, padding: "40px 44px", display: "flex", flexDirection: "column", rowGap: 30 }}>
          <Field label="Finalidade da solicitação" delay={14}>
            <div style={{ display: "flex", columnGap: 12 }}>
              {FINALIDADES.map((f, i) => (
                <div key={f} style={{ flex: 1, height: 70, borderRadius: 14, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: F.body, fontWeight: 700, fontSize: 23, color: i === sel ? "#fff" : C.text, background: i === sel ? C.primary : "#F0EEEC", border: `1px solid ${i === sel ? C.primary : C.line}`, boxShadow: i === sel ? `0 12px 24px -10px ${C.primary}` : undefined }}>{f}</div>
              ))}
            </div>
          </Field>
          <Field label="Cliente / Depositante" delay={34}>
            <Input>
              <span>{typed}<span style={{ display: "inline-block", width: 2, height: 30, background: C.text, marginLeft: 3, verticalAlign: "middle", opacity: frame < 108 ? Math.floor(frame / 8) % 2 : 0 }} /></span>
              <span style={{ color: C.textDim }}>⌄</span>
            </Input>
          </Field>
          <Field label="Planilha do pedido" delay={54}>
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
  { t: "Identificar cliente / template", at: 40 },
  { t: "Extrair dados da planilha", at: 90 },
  { t: "Validar CNPJ/CPF na Receita Federal", at: 140 },
  { t: "Validar cadastro, endereço e estoque", at: 200 },
  { t: "Gerar orçamento / grade", at: 255 },
];
export const S4: React.FC = () => {
  const frame = useCurrentFrame();
  const done = STEPS.filter((s) => frame >= s.at).length;
  const pct = interpolate(frame, [20, 260], [0, 1], clamp);
  const note = useSpring(270, "bouncy");
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

// ---------- 5. Correção guiada ----------
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
  const t = [0, 50, 70, 120, 140, 180, 200];
  const pos = interpolate(frame, t, stops, { ...clamp, easing: theme.ease.inOut });
  const colNow = Math.round(pos);
  const hx = x0 + pos * (colW + gap) + 16;
  const mail = useSpring(140, "bouncy");
  const appr = useSpring(205, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 120 }}>
        <Label>Orçamento</Label>
        <div style={{ marginTop: 14 }}>
          <WordReveal text="Do pedido à aprovação." delay={6} size={72} highlight={["aprovação."]} />
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
        <div style={{ ...panel, borderRadius: 16, padding: "18px 22px", border: `2px solid ${C.primary}`, boxShadow: `0 24px 50px -16px ${C.glow}, 0 30px 60px -20px rgba(35,31,32,0.3)`, transform: `translateY(${float(frame, 3, 24)}px) rotate(${(frame > 50 && frame < 70) || (frame > 140 && frame < 180) ? -2 : 0}deg)` }}>
          <div style={{ fontFamily: F.display, fontSize: 28, color: C.text }}>Pacote · Job Nivea</div>
          <div style={{ fontFamily: F.body, fontSize: 21, color: C.textDim, marginTop: 6 }}>{colNow === 0 ? "Recebido, aguardando análise" : colNow === 1 ? "Planilha .xlsm e PDF anexados" : colNow === 2 ? "Cliente notificado por e-mail" : "Aprovado pelo cliente"}</div>
        </div>
      </Entrance>
      <div style={{ position: "absolute", left: x0 + 2 * (colW + gap) + 16, top: 640, opacity: mail, transform: `scale(${mail})`, transformOrigin: "left center" }}>
        <Chip color={C.primary}>✉ E-mail enviado ao cliente</Chip>
      </div>
      <div style={{ position: "absolute", left: x0 + 3 * (colW + gap) + 16, top: 520, opacity: appr, transform: `scale(${appr})`, transformOrigin: "left center" }}>
        <Chip color={C.ok}>✓ Segue para a operação</Chip>
      </div>
    </Scene>
  );
};

// ---------- 7. Operação ----------
const MODS = [
  ["Distribuição de Materiais", "Materiais por classe"],
  ["Gestão de Romaneios", "Pedidos por romaneio"],
  ["Operação de Coletas", "Da ordem ao recebimento"],
  ["Agenda de Retiradas", "Baixa e assinatura"],
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
        <WordReveal text="Aprovado, a operação assume." delay={4} size={88} align="center" highlight={["operação"]} />
      </div>
      {MODS.map(([t, d], i) => {
        const col = i % 4, row = Math.floor(i / 4);
        return (
          <Entrance key={t} delay={34 + i * 16} y={40} style={{ position: "absolute", left: 130 + col * 420, top: 380 + row * 220, width: 390 }}>
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

// ---------- 8. Encerramento ----------
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
