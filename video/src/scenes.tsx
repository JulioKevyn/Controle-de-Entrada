import React from "react";
import { AbsoluteFill, Img, interpolate, staticFile, useCurrentFrame, useVideoConfig } from "remotion";
import { Chip, Counter, Entrance, Label, Mark, Scene, WordReveal, breathe, float, useSpring } from "./components";
import { theme } from "./theme";

const clamp = { extrapolateLeft: "clamp", extrapolateRight: "clamp" } as const;
const F = theme.fonts;
const C = theme.colors;

const panel: React.CSSProperties = {
  background: C.card,
  border: `1px solid ${C.line}`,
  borderRadius: 28,
  boxShadow: "0 40px 80px -20px rgba(0,0,0,0.6)",
};

const CLIENTS = ["lg", "bayer", "heineken", "mondelez", "whirlpool", "danone", "diageo", "nivea", "opella", "unilever"];
const NAMES = ["LG", "BAYER", "HEINEKEN", "MONDELEZ", "WHIRLPOOL", "DANONE", "DIAGEO", "NIVEA", "OPELLA", "UNILEVER"];

// ---------- 1. Abertura ----------
export const S1: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <Scene>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 34 }}>
        <Mark size={120} />
        <div style={{ transform: `translateY(${float(frame, 3)}px)` }}>
          <WordReveal text="Solicitação de Materiais" delay={36} size={140} align="center" highlight={["Materiais"]} />
        </div>
        <Entrance delay={78} y={20}>
          <div style={{ fontFamily: F.body, fontSize: 36, color: C.textDim }}>Sistema de controle de pedidos e envios</div>
        </Entrance>
      </AbsoluteFill>
    </Scene>
  );
};

// ---------- 2. Problema ----------
const SheetCard: React.FC<{ i: number; name: string }> = ({ i, name }) => {
  const frame = useCurrentFrame();
  const p = useSpring(18 + i * 5, "bouncy");
  const rot = ((i * 37) % 17) - 8;
  const pos = [
    [0, 0], [310, 20], [620, 0], [150, 160], [460, 170],
    [0, 330], [310, 340], [620, 330], [150, 500], [460, 510],
  ][i];
  const cols = 3 + (i % 3);
  return (
    <div
      style={{
        position: "absolute", left: pos[0], top: pos[1], width: 240, height: 160, ...panel, borderRadius: 16, padding: 14,
        opacity: p,
        transform: `translateY(${float(frame + i * 9, 6, 26)}px) rotate(${rot * p}deg) scale(${interpolate(p, [0, 1], [0.6, 1])})`,
      }}
    >
      <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 16, color: C.textDim, letterSpacing: "0.1em", marginBottom: 10 }}>{name}.xlsx</div>
      <div style={{ display: "grid", gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: 5 }}>
        {Array.from({ length: cols * 5 }).map((_, k) => (
          <div key={k} style={{ height: 12, borderRadius: 3, background: k < cols ? `${C.accent}55` : "rgba(255,255,255,0.09)" }} />
        ))}
      </div>
    </div>
  );
};

export const S2: React.FC = () => (
  <Scene>
    <div style={{ position: "absolute", left: 130, top: 230, width: 760 }}>
      <Label>O desafio</Label>
      <div style={{ display: "flex", alignItems: "baseline", columnGap: 28, marginTop: 20 }}>
        <Entrance delay={8} cfg="snappy">
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 260, lineHeight: 1, letterSpacing: "-0.05em", color: C.primary, textShadow: `0 0 70px ${C.glow}` }}>
            <Counter to={10} delay={8} />
          </div>
        </Entrance>
        <WordReveal text="clientes" delay={14} size={90} />
      </div>
      <div style={{ marginTop: 30 }}>
        <WordReveal text="Dezenas de planilhas, cada uma com seu formato." delay={60} size={54} weight={600} />
      </div>
    </div>
    <div style={{ position: "absolute", left: 940, top: 250, width: 900, height: 600 }}>
      {NAMES.map((n, i) => <SheetCard key={n} i={i} name={n} />)}
    </div>
  </Scene>
);

// ---------- 3. Acesso por cliente ----------
export const S3: React.FC = () => {
  const frame = useCurrentFrame();
  const modal = useSpring(70, "bouncy");
  const modalOut = interpolate(frame, [168, 182], [1, 0], { ...clamp, easing: theme.ease.in });
  const dots = Math.floor(interpolate(frame, [92, 128], [0, 6], clamp));
  const ok = useSpring(138, "bouncy");
  const dim = interpolate(frame, [66, 84], [0, 0.65], clamp) * modalOut;
  return (
    <Scene>
      <div style={{ position: "absolute", left: 150, top: 110 }}>
        <Label>Acesso</Label>
        <div style={{ marginTop: 14 }}>
          <WordReveal text="Um módulo por cliente." delay={6} size={84} highlight={["cliente."]} />
        </div>
      </div>
      {CLIENTS.map((c, i) => {
        const col = i % 5, row = Math.floor(i / 5);
        return (
          <Entrance key={c} delay={24 + i * 4} y={50} style={{ position: "absolute", left: 150 + col * 330, top: 400 + row * 200, width: 300, height: 170 }}>
            <div style={{ width: 300, height: 170, borderRadius: 24, background: "#FFFFFF", display: "flex", alignItems: "center", justifyContent: "center", padding: 28, boxShadow: "0 24px 50px -16px rgba(0,0,0,0.6)", transform: `translateY(${float(frame + i * 11, 4, 28)}px)` }}>
              <Img src={staticFile(`logos/${c}.png`)} style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain" }} />
            </div>
          </Entrance>
        );
      })}
      <AbsoluteFill style={{ background: `rgba(0,0,0,${dim})` }} />
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", opacity: modal * modalOut, transform: `scale(${interpolate(modal, [0, 1], [0.85, 1])})` }}>
        <div style={{ ...panel, width: 620, padding: 48, textAlign: "center", boxShadow: `0 40px 100px rgba(0,0,0,0.7), 0 0 90px ${C.glow}` }}>
          <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 38, color: C.text }}>Senha para LG</div>
          <div style={{ margin: "32px 0", height: 74, borderRadius: 14, background: "#2a2a2a", border: `1px solid ${C.line}`, display: "flex", alignItems: "center", justifyContent: "center", columnGap: 16 }}>
            {Array.from({ length: 6 }).map((_, k) => (
              <div key={k} style={{ width: 18, height: 18, borderRadius: 9, background: k < dots ? C.text : "rgba(255,255,255,0.12)" }} />
            ))}
          </div>
          <div style={{ height: 56, display: "flex", justifyContent: "center", alignItems: "center", transform: `scale(${ok})`, opacity: ok }}>
            <Chip color={C.ok}>✓ Acesso liberado</Chip>
          </div>
        </div>
      </AbsoluteFill>
    </Scene>
  );
};

// ---------- tabela compartilhada ----------
const ROWS = [
  { nf: "48213", base: "São Paulo", d: "12/09", pend: true },
  { nf: "48217", base: "Campinas", d: "12/09", pend: true },
  { nf: "48220", base: "Curitiba", d: "13/09", pend: false },
  { nf: "48231", base: "Campinas", d: "14/09", pend: true },
  { nf: "48236", base: "Recife", d: "14/09", pend: true },
  { nf: "48244", base: "Campinas", d: "15/09", pend: true },
  { nf: "48250", base: "Salvador", d: "15/09", pend: false },
];
const cols = "150px 1fr 140px 250px";

const HeadRow: React.FC<{ c?: string }> = ({ c = cols }) => (
  <div style={{ display: "grid", gridTemplateColumns: c, padding: "0 32px", height: 64, alignItems: "center", fontFamily: F.body, fontWeight: 600, fontSize: 20, letterSpacing: "0.12em", color: C.textDim, borderBottom: `1px solid ${C.line}` }}>
    <span>NF</span><span>BASE / CIDADE</span><span>DATA</span><span>STATUS</span>
  </div>
);

export const S4: React.FC = () => {
  const frame = useCurrentFrame();
  const scan = interpolate(frame, [60, 130], [0, 1], { ...clamp, easing: theme.ease.inOut });
  const scanOn = frame >= 60 && frame <= 134;
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 300, width: 640 }}>
        <Label>Leitura automática</Label>
        <div style={{ marginTop: 18 }}>
          <WordReveal text="Colunas e datas tratadas sozinhas." delay={6} size={70} highlight={["sozinhas."]} />
        </div>
        <Entrance delay={50} y={20}>
          <div style={{ marginTop: 28, fontFamily: F.body, fontSize: 30, color: C.textDim }}>Planilhas .xlsx → pendências prontas.</div>
        </Entrance>
      </div>
      <Entrance delay={4} x={80} y={0} style={{ position: "absolute", left: 840, top: 190, width: 960 }}>
        <div style={{ ...panel, overflow: "hidden", position: "relative", transform: `translateY(${float(frame, 4)}px)` }}>
          <HeadRow />
          {ROWS.map((r, i) => {
            const rowY = (i + 0.5) / ROWS.length;
            const reveal = scan > rowY;
            const e = useSpring(16 + i * 5, "snappy");
            return (
              <div key={r.nf} style={{ display: "grid", gridTemplateColumns: cols, padding: "0 32px", height: 78, alignItems: "center", fontFamily: F.body, fontSize: 28, color: C.text, opacity: e, transform: `translateX(${interpolate(e, [0, 1], [30, 0])}px)`, borderBottom: `1px solid ${C.line}`, background: reveal && r.pend ? `${C.primary}14` : "transparent" }}>
                <span style={{ fontWeight: 600, fontVariantNumeric: "tabular-nums" }}>{r.nf}</span>
                <span>{r.base}</span>
                <span style={{ color: C.textDim }}>{r.d}</span>
                <span style={{ opacity: reveal ? 1 : 0 }}>{r.pend ? <Chip color={C.primary}>Sem baixa</Chip> : <Chip color={C.ok}>Baixada</Chip>}</span>
              </div>
            );
          })}
          {scanOn && (
            <div style={{ position: "absolute", left: 0, right: 0, top: 64 + scan * 78 * ROWS.length, height: 3, background: C.accent, boxShadow: `0 0 30px 6px ${C.accent}88` }} />
          )}
        </div>
      </Entrance>
    </Scene>
  );
};

// ---------- 5. Filtro ----------
export const S5: React.FC = () => {
  const frame = useCurrentFrame();
  const word = "Campinas";
  const typed = word.slice(0, Math.floor(interpolate(frame, [26, 56], [0, word.length + 0.99], clamp)));
  const press = interpolate(frame, [62, 66, 72], [1, 0.93, 1], clamp);
  const f = interpolate(frame, [70, 92], [0, 1], { ...clamp, easing: theme.ease.inOut });
  const pend = ROWS.filter((r) => r.pend);
  return (
    <Scene>
      <div style={{ position: "absolute", left: 130, top: 270, width: 700 }}>
        <Label>Filtros dinâmicos</Label>
        <div style={{ marginTop: 18 }}>
          <WordReveal text="Só o que exige ação." delay={6} size={80} highlight={["ação."]} />
        </div>
        <Entrance delay={20} y={24} style={{ marginTop: 44 }}>
          <div style={{ ...panel, padding: 24, display: "flex", alignItems: "center", columnGap: 18, borderRadius: 20 }}>
            <span style={{ fontFamily: F.body, fontSize: 26, color: C.textDim }}>Filtrar base:</span>
            <div style={{ flex: 1, height: 58, borderRadius: 12, background: "#2a2a2a", display: "flex", alignItems: "center", padding: "0 18px", fontFamily: F.body, fontSize: 28, color: C.text }}>
              {typed}
              <span style={{ width: 2, height: 30, background: C.text, marginLeft: 3, opacity: Math.floor(frame / 8) % 2 }} />
            </div>
            <div style={{ padding: "14px 28px", borderRadius: 12, background: C.primary, color: "#fff", fontFamily: F.body, fontWeight: 600, fontSize: 26, transform: `scale(${press})` }}>Filtrar</div>
          </div>
        </Entrance>
      </div>
      <Entrance delay={6} x={80} y={0} style={{ position: "absolute", left: 900, top: 230, width: 900 }}>
        <div style={{ ...panel, overflow: "hidden" }}>
          <HeadRow c="130px 1fr 120px 230px" />
          {ROWS.map((r, i) => {
            const keep = r.base === word;
            const k = keep ? 1 : 1 - f;
            return (
              <div key={r.nf} style={{ display: "grid", gridTemplateColumns: "130px 1fr 120px 230px", padding: "0 32px", height: 78 * k, overflow: "hidden", alignItems: "center", fontFamily: F.body, fontSize: 28, color: C.text, opacity: k, borderBottom: `1px solid ${C.line}`, background: keep && f > 0.9 ? `${C.primary}14` : "transparent" }}>
                <span style={{ fontWeight: 600 }}>{r.nf}</span><span>{r.base}</span><span style={{ color: C.textDim }}>{r.d}</span>
                <span>{r.pend ? <Chip color={C.primary}>Sem baixa</Chip> : <Chip color={C.ok}>Baixada</Chip>}</span>
              </div>
            );
          })}
        </div>
        <Entrance delay={98} y={20} style={{ marginTop: 30 }}>
          <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 44, color: C.text }}>
            <Counter to={pend.filter((r) => r.base === word).length} delay={98} style={{ color: C.primary }} /> notas pendentes em Campinas
          </div>
        </Entrance>
      </Entrance>
    </Scene>
  );
};

// ---------- 6. Envio ----------
export const S6: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const lines = [
    "Prezados, segue a relação de notas entregues sem baixa:",
    "NF 48217 · 12/09",
    "NF 48231 · 14/09",
    "NF 48244 · 15/09",
  ];
  const press = interpolate(frame, [128, 132, 138], [1, 0.92, 1], clamp);
  const fly = interpolate(frame, [142, 168], [0, 1], { ...clamp, easing: theme.ease.in });
  const sent = useSpring(166, "bouncy");
  return (
    <Scene>
      <div style={{ position: "absolute", left: 0, right: 0, top: 80, display: "flex", justifyContent: "center" }}>
        <Label>Envio automatizado</Label>
      </div>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
        <Entrance delay={6} y={60} style={{ width: 1160, transform: undefined }}>
          <div style={{ ...panel, width: 1160, overflow: "hidden", transform: `translate(${fly * 900}px, ${-fly * 500}px) scale(${1 - fly * 0.5}) rotate(${fly * 8}deg)`, opacity: 1 - fly, filter: `blur(${fly * 6}px)` }}>
            <div style={{ height: 56, background: "#241d16", display: "flex", alignItems: "center", padding: "0 24px", columnGap: 10 }}>
              {["#ff5f57", "#febc2e", "#28c840"].map((c) => <div key={c} style={{ width: 14, height: 14, borderRadius: 7, background: c }} />)}
              <span style={{ marginLeft: 20, fontFamily: F.body, fontSize: 22, color: C.textDim }}>Nova mensagem — Outlook</span>
            </div>
            <div style={{ padding: "30px 44px", fontFamily: F.body, fontSize: 28, color: C.text }}>
              <div style={{ color: C.textDim, marginBottom: 12 }}>Para: <span style={{ color: C.text }}>responsável da base Campinas</span></div>
              <div style={{ color: C.textDim, paddingBottom: 18, borderBottom: `1px solid ${C.line}` }}>Assunto: <span style={{ color: C.text, fontWeight: 600 }}>[NF] Notas entregues sem baixa – Campinas</span></div>
              <div style={{ marginTop: 24, minHeight: 190 }}>
                {lines.map((l, i) => {
                  const p = useSpring(26 + i * 14, "smooth");
                  return <div key={i} style={{ opacity: p, transform: `translateY(${interpolate(p, [0, 1], [14, 0])}px)`, marginBottom: 12, color: i === 0 ? C.text : C.textDim, fontVariantNumeric: "tabular-nums" }}>{l}</div>;
                })}
              </div>
              <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
                <div style={{ padding: "16px 40px", borderRadius: 14, background: C.primary, color: "#fff", fontWeight: 600, fontSize: 28, transform: `scale(${press * breathe(frame, 0.012)})`, boxShadow: `0 0 50px ${C.glow}` }}>Confirmar Envio</div>
              </div>
            </div>
          </div>
        </Entrance>
      </AbsoluteFill>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center" }}>
        <div style={{ opacity: sent, transform: `scale(${interpolate(sent, [0, 1], [0.5, 1])})`, textAlign: "center" }}>
          <div style={{ width: 150, height: 150, borderRadius: 75, margin: "0 auto 28px", background: `${C.ok}22`, border: `3px solid ${C.ok}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 84, color: C.ok }}>✓</div>
          <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 64, color: C.text }}>Solicitação enviada</div>
        </div>
      </AbsoluteFill>
    </Scene>
  );
};

// ---------- 7. Benefícios ----------
export const S7: React.FC = () => (
  <Scene>
    <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", rowGap: 22 }}>
      <WordReveal text="Menos retrabalho." delay={4} size={130} align="center" />
      <WordReveal text="Mais controle." delay={30} size={130} align="center" />
      <WordReveal text="Mais prazo cumprido." delay={56} size={130} align="center" highlight={["prazo", "cumprido."]} />
    </AbsoluteFill>
  </Scene>
);

// ---------- 8. Encerramento ----------
export const S8: React.FC = () => {
  const frame = useCurrentFrame();
  const { durationInFrames } = useVideoConfig();
  const out = interpolate(frame, [durationInFrames - 14, durationInFrames - 1], [1, 0], { ...clamp, easing: theme.ease.in });
  return (
    <AbsoluteFill style={{ opacity: out }}>
      <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 34 }}>
        <Mark size={170} />
        <Entrance delay={24} y={20}>
          <div style={{ fontFamily: F.display, fontWeight: 700, fontSize: 80, letterSpacing: "-0.02em", color: C.text, transform: `scale(${breathe(frame, 0.01)})` }}>Fazendo marcas <span style={{ color: C.primary, textShadow: `0 0 50px ${C.glow}` }}>venderem mais.</span></div>
        </Entrance>
      </AbsoluteFill>
    </AbsoluteFill>
  );
};
