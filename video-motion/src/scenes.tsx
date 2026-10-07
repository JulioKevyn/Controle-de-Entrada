import React from "react";
import { AbsoluteFill, Img, interpolate, staticFile, useCurrentFrame } from "remotion";
import { C, Chip, Entrance, F, GradText, Label, Scene, WordReveal, breathe, clamp, float, panel, useSpring, useV } from "./components";
import { theme } from "./theme";

const Center: React.FC<{ children: React.ReactNode; gap?: number }> = ({ children, gap = 36 }) => (
  <AbsoluteFill style={{ alignItems: "center", justifyContent: "center", flexDirection: "column", rowGap: gap, padding: "0 60px" }}>{children}</AbsoluteFill>
);

// Formas animadas de fundo (anéis girando)
const Rings: React.FC<{ strong?: boolean }> = ({ strong }) => {
  const frame = useCurrentFrame();
  const { w, h } = useV();
  return (
    <svg width={w} height={h} style={{ position: "absolute", inset: 0, opacity: strong ? 0.5 : 0.22 }}>
      <defs>
        <linearGradient id="rg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stopColor="#FF7A18" /><stop offset="0.5" stopColor="#FF2E9A" /><stop offset="1" stopColor="#8B5CF6" /></linearGradient>
      </defs>
      {[0, 1, 2, 3].map((i) => (
        <circle key={i} cx={w / 2} cy={h / 2} r={220 + i * 130 + Math.sin(frame / 20 + i) * 14} fill="none" stroke="url(#rg)" strokeWidth={3} strokeDasharray={`${40 + i * 30} ${70 + i * 20}`} transform={`rotate(${frame * (i % 2 ? -0.8 : 0.9) + i * 40} ${w / 2} ${h / 2})`} />
      ))}
    </svg>
  );
};

// 1 ---------- Gancho ----------
export const M1: React.FC = () => {
  const frame = useCurrentFrame();
  const { V } = useV();
  const a = useSpring(79, "bouncy");
  const b = useSpring(144, "bouncy");
  return (
    <Scene>
      <Rings strong />
      <Center gap={30}>
        <Entrance delay={0} y={20}><div style={{ fontFamily: F.body, fontWeight: 800, fontSize: 30, letterSpacing: "0.34em", color: C.teal }}>VÍDEOS MOTION</div></Entrance>
        <WordReveal text="Seu negócio em movimento." delay={4} size={V ? 112 : 132} align="center" highlight={["movimento."]} per={4} />
        <div style={{ opacity: Math.min(1, a), transform: `translateY(${interpolate(a, [0, 1], [30, 0])}px)`, fontFamily: F.display, fontWeight: 700, fontSize: V ? 48 : 54, color: C.text, textAlign: "center" }}>que chamam atenção</div>
        <div style={{ opacity: Math.min(1, b), transform: `translateY(${interpolate(b, [0, 1], [30, 0])}px) scale(${breathe(frame, 0.02)})`, fontFamily: F.display, fontWeight: 700, fontSize: V ? 48 : 54, textAlign: "center" }}><GradText>e fazem o cliente parar.</GradText></div>
      </Center>
    </Scene>
  );
};

// 2 ---------- Timeline ----------
const LAYERS = [["Texto animado", 6, C.teal], ["Cores da marca", 38, C.pink], ["Trilha sonora", 80, C.violet], ["Legendas", 108, C.sky]] as const;
export const M2: React.FC = () => {
  const frame = useCurrentFrame();
  const { V } = useV();
  const head = interpolate(frame, [0, 250], [0, 1], clamp);
  const end = useSpring(142, "bouncy");
  const W = V ? 960 : 1380;
  const pv = useSpring(2, "smooth");
  return (
    <Scene>
      <Center gap={32}>
        <div style={{ ...panel, width: W, padding: "26px 34px", opacity: Math.min(1, pv), transform: `scale(${interpolate(pv, [0, 1], [0.92, 1])})` }}>
          <div style={{ display: "flex", alignItems: "center", columnGap: 12, marginBottom: 18 }}>
            {["#FF5F57", "#FEBC2E", "#28C840"].map((c) => <div key={c} style={{ width: 14, height: 14, borderRadius: 7, background: c }} />)}
            <span style={{ fontFamily: F.body, fontWeight: 700, fontSize: 22, color: C.muted, marginLeft: 12 }}>Timeline</span>
          </div>
          <div style={{ position: "relative" }}>
            {LAYERS.map(([n, at, c], i) => {
              const grow = interpolate(frame, [at, at + 28], [0, 1], { ...clamp, easing: theme.ease.out });
              const lw = V ? 220 : 280;
              return (
                <div key={n} style={{ display: "flex", alignItems: "center", height: V ? 92 : 100, borderBottom: i < 3 ? `1px solid ${C.line}` : undefined }}>
                  <div style={{ width: lw, fontFamily: F.body, fontWeight: 700, fontSize: V ? 26 : 30, color: C.text }}>{n}</div>
                  <div style={{ flex: 1, position: "relative", height: 48 }}>
                    <div style={{ position: "absolute", left: `${(i * 9) % 30}%`, width: `${(45 + ((i * 13) % 25)) * grow}%`, height: 48, borderRadius: 12, background: `linear-gradient(90deg, ${c}, ${c}AA)`, boxShadow: `0 0 28px ${c}66`, display: "flex", alignItems: "center", paddingLeft: 14 }}>
                      {[0, 1, 2].map((k) => <div key={k} style={{ width: 12, height: 12, background: "#fff", transform: "rotate(45deg)", marginRight: 26, opacity: grow > 0.9 ? 1 : 0 }} />)}
                    </div>
                  </div>
                </div>
              );
            })}
            <div style={{ position: "absolute", top: -6, bottom: -6, left: `calc(${V ? 220 : 280}px + ${head} * (100% - ${V ? 220 : 280}px))`, width: 3, background: "#fff", boxShadow: "0 0 18px #fff" }} />
          </div>
        </div>
        <WordReveal text="Tudo pensado para mostrar o que você faz em segundos." delay={142} size={V ? 52 : 54} align="center" highlight={["segundos."]} per={3} />
        <div style={{ opacity: Math.min(1, end), transform: `scale(${end})` }}><Chip color={C.teal} size={28}>animação · cor · som · legenda</Chip></div>
      </Center>
    </Scene>
  );
};

// 3 ---------- Tipos de vídeo ----------
const Tile: React.FC<{ i: number; at: number; t: string; icon: React.ReactNode; col: string }> = ({ i, at, t, icon, col }) => {
  const frame = useCurrentFrame();
  const p = useSpring(at, "bouncy");
  return (
    <div style={{ ...panel, width: "100%", height: "100%", border: `2px solid ${col}88`, boxShadow: `0 0 50px ${col}33`, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", rowGap: 18, opacity: Math.min(1, p), transform: `scale(${interpolate(p, [0, 1], [0.6, 1])}) translateY(${float(frame + i * 9, 6, 24)}px)` }}>
      <div style={{ width: 110, height: 110, position: "relative" }}>{icon}</div>
      <div style={{ fontFamily: F.display, fontSize: 28, color: C.text, textAlign: "center", padding: "0 14px", lineHeight: 1.2 }}>{t}</div>
    </div>
  );
};
const Ic = {
  company: (f: number, c: string) => <>{[0, 1, 2].map((k) => <div key={k} style={{ position: "absolute", bottom: 0, left: 10 + k * 34, width: 28, height: (40 + k * 24) * (0.7 + 0.3 * Math.sin(f / 10 + k)), background: c, borderRadius: 5 }} />)}</>,
  ads: (f: number, c: string) => <div style={{ position: "absolute", left: 28, top: 0, width: 54, height: 110, border: `5px solid ${c}`, borderRadius: 14 }}><div style={{ position: "absolute", left: 14, top: 38, width: 0, height: 0, borderLeft: `20px solid ${c}`, borderTop: "13px solid transparent", borderBottom: "13px solid transparent", transform: `scale(${1 + 0.15 * Math.sin(f / 6)})` }} /></div>,
  launch: (f: number, c: string) => <div style={{ position: "absolute", left: 22, top: 22, width: 66, height: 66, background: c, borderRadius: 12, transform: `rotate(${f * 2}deg)` }} />,
  brand: (f: number, c: string) => <><div style={{ position: "absolute", inset: 8, border: `6px solid ${c}`, borderRadius: "50%", transform: `rotate(${f * 3}deg)`, borderTopColor: "transparent" }} /><div style={{ position: "absolute", left: 38, top: 38, width: 34, height: 34, borderRadius: "50%", background: c }} /></>,
  explain: (f: number, c: string) => <>{[0, 1, 2, 3].map((k) => <div key={k} style={{ position: "absolute", bottom: 0, left: 4 + k * 27, width: 20, height: 30 + ((k * 23 + f * 1.4) % 70), background: c, borderRadius: 4 }} />)}</>,
  more: (f: number, c: string) => <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: F.display, fontWeight: 800, fontSize: 96, color: c, transform: `scale(${1 + 0.1 * Math.sin(f / 7)})` }}>+</div>,
};
export const M3: React.FC = () => {
  const frame = useCurrentFrame();
  const { V } = useV();
  const T: [string, number, keyof typeof Ic, string][] = [
    ["Apresentação de empresa", 6, "company", C.teal],
    ["Anúncios para redes sociais", 48, "ads", C.pink],
    ["Lançamento de produto", 101, "launch", C.violet],
    ["Abertura de marca", 140, "brand", C.sky],
    ["Vídeo explicativo", 190, "explain", C.teal],
    ["E muito mais", 246 - 22, "more", C.pink],
  ];
  const cols = V ? 2 : 3;
  const tw = V ? 460 : 520, th = V ? 330 : 330, gx = 28;
  const ban = useSpring(228, "bouncy");
  return (
    <Scene>
      <Rings />
      <Center gap={40}>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: V ? 64 : 64, color: C.text, textAlign: "center" }}>Qualquer vídeo <GradText>motion</GradText></div>
        <div style={{ display: "grid", gridTemplateColumns: `repeat(${cols}, ${tw}px)`, gap: gx }}>
          {T.map(([t, at, ic, col], i) => (
            <div key={t} style={{ width: tw, height: th }}><Tile i={i} at={at} t={t} icon={Ic[ic](frame, col)} col={col} /></div>
          ))}
        </div>
        <div style={{ opacity: Math.min(1, ban), transform: `scale(${ban})` }}><Chip color={C.teal} size={32}>que você precisar</Chip></div>
      </Center>
    </Scene>
  );
};

// 4 ---------- Formatos ----------
const Slides: React.FC<{ srcs: string[]; w: number; h: number; every?: number }> = ({ srcs, w, h, every = 60 }) => {
  const frame = useCurrentFrame();
  const k = Math.floor(frame / every) % srcs.length;
  const local = (frame % every) / every;
  return (
    <div style={{ width: w, height: h, position: "relative", overflow: "hidden", background: "#000" }}>
      {srcs.map((s, i) => (
        <Img key={s} src={staticFile(s)} style={{ position: "absolute", inset: 0, width: "100%", height: "100%", objectFit: "cover", opacity: i === k ? 1 : 0, transform: `scale(${1 + 0.06 * (i === k ? local : 0)})` }} />
      ))}
    </div>
  );
};
export const M4: React.FC = () => {
  const frame = useCurrentFrame();
  const { V } = useV();
  const lap = useSpring(8, "bouncy");
  const ph = useSpring(75, "bouncy");
  const plats = [["Instagram", 131], ["Reels", 160], ["Stories", 183], ["YouTube", 210], ["Seu site", 227]] as const;
  const LW = V ? 880 : 760, LH = LW * 0.5625;
  const PW = V ? 330 : 320, PH = PW * 1.78;
  return (
    <Scene>
      <Center gap={V ? 36 : 40}>
        <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: V ? 60 : 64, color: C.text, textAlign: "center" }}>No <GradText>computador</GradText> e no <GradText>celular</GradText></div>
        <div style={{ display: "flex", alignItems: "flex-end", columnGap: 40, flexDirection: V ? "column" : "row", rowGap: 30, alignItemsCenter: undefined } as React.CSSProperties}>
          <div style={{ opacity: Math.min(1, lap), transform: `translateY(${interpolate(lap, [0, 1], [60, 0])}px)`, filter: `drop-shadow(0 30px 60px ${C.pink}55)` }}>
            <div style={{ padding: 12, borderRadius: 18, background: "#1b1b24", border: `2px solid ${C.line}` }}><Slides srcs={["img/land_22.jpg", "img/land_66.jpg", "img/land_86.jpg", "img/land_112.jpg"]} w={LW} h={LH} /></div>
            <div style={{ height: 14, background: "#2a2a36", borderRadius: "0 0 14px 14px", margin: "0 -30px" }} />
          </div>
          <div style={{ opacity: Math.min(1, ph), transform: `translateY(${interpolate(ph, [0, 1], [80, 0])}px) scale(${breathe(frame, 0.01)})`, filter: `drop-shadow(0 30px 60px ${C.teal}55)` }}>
            <div style={{ padding: 10, borderRadius: 40, background: "#1b1b24", border: `2px solid ${C.line}` }}><div style={{ borderRadius: 30, overflow: "hidden" }}><Slides srcs={["img/vert_5.jpg", "img/vert_9.jpg", "img/vert_12.jpg", "img/vert_15.jpg"]} w={PW} h={PH} every={50} /></div></div>
          </div>
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "center", gap: 14, maxWidth: 900 }}>
          {plats.map(([n, at]) => <PlatChip key={n} n={n} at={at} />)}
        </div>
        <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 24, color: C.dim }}>Exemplo: vídeo do Future PDV</div>
      </Center>
    </Scene>
  );
};
const PlatChip: React.FC<{ n: string; at: number }> = ({ n, at }) => {
  const p = useSpring(at - 0, "bouncy");
  return <div style={{ opacity: Math.min(1, p), transform: `scale(${p})` }}><Chip color={C.teal} size={30}>{n}</Chip></div>;
};

// 5 ---------- CTA ----------
export const M5: React.FC = () => {
  const frame = useCurrentFrame();
  const { V } = useV();
  const btn = useSpring(14, "bouncy");
  return (
    <Scene>
      <Rings strong />
      <Center gap={44}>
        <WordReveal text="Seu negócio em movimento." delay={0} size={V ? 100 : 108} align="center" highlight={["movimento."]} per={3} />
        <div style={{ opacity: Math.min(1, btn), transform: `scale(${btn * breathe(frame, 0.035, 9)})`, display: "flex", alignItems: "center", columnGap: 22, padding: "26px 60px", borderRadius: 999, background: "linear-gradient(120deg, #25D366, #128C7E)", color: "#fff", fontFamily: F.display, fontWeight: 800, fontSize: V ? 46 : 52, boxShadow: "0 20px 70px -10px #25D36699" }}>
          <span style={{ fontSize: 56 }}>●</span> Chama no WhatsApp
        </div>
        <Entrance delay={40} y={20}><div style={{ fontFamily: F.display, fontSize: V ? 44 : 48, color: C.text, textAlign: "center" }}>Julio Clemente</div></Entrance>
        <Entrance delay={52} y={20}><div style={{ fontFamily: F.body, fontWeight: 700, fontSize: 40, color: C.teal }}>(11) 96620-9914</div></Entrance>
      </Center>
    </Scene>
  );
};
