import React from "react";
import { AbsoluteFill, Img, interpolate, staticFile, useCurrentFrame } from "remotion";
import { BgBlur, Chip, clamp, C, Entrance, F, Photo, Scene, useSpring, WordReveal } from "./components";
import { theme } from "./theme";

const Gold: React.FC<{ w?: number }> = ({ w = 220 }) => {
  const p = useSpring(10);
  return <div style={{ width: w * p, height: 3, background: C.gold, borderRadius: 2 }} />;
};

export const B1: React.FC = () => {
  const frame = useCurrentFrame();
  return (
    <Scene>
      <BgBlur src="natal" />
      <AbsoluteFill style={{ alignItems: "center", paddingTop: 170 }}>
        <Photo src="natal" w={820} h={1090} rot={-2} drift={1.4} y={120} x={0} />
      </AbsoluteFill>
      <AbsoluteFill style={{ justifyContent: "flex-end", alignItems: "center", paddingBottom: 330, background: "linear-gradient(180deg, transparent 45%, rgba(18,6,10,0.92) 82%)" }}>
        <div style={{ opacity: interpolate(frame, [0, 20], [0, 1], clamp) }}>
          <WordReveal text="Momentos passam rápido." size={72} delay={6} />
        </div>
        <div style={{ height: 26 }} />
        <WordReveal text="Nós transformamos em memórias" size={64} italic highlight={["memórias"]} delay={64} weight={600} />
      </AbsoluteFill>
    </Scene>
  );
};

const SERVICES = ["Retratos", "Ensaios", "15 Anos", "Eventos", "Corporativo"];

export const B2: React.FC = () => {
  const frame = useCurrentFrame();
  const pr = useSpring(4, "bouncy");
  return (
    <Scene>
      <BgBlur src="escada" />
      <AbsoluteFill style={{ alignItems: "center", paddingTop: 150 }}>
        <div style={{ position: "relative", width: 320, height: 320, transform: `scale(${pr})` }}>
          <div style={{ position: "absolute", inset: -14, borderRadius: "50%", border: `3px solid ${C.gold}`, boxShadow: `0 0 60px ${C.gold}66` }} />
          <Img src={staticFile("img/avatar.jpg")} style={{ width: "100%", height: "100%", borderRadius: "50%", objectFit: "cover" }} />
        </div>
        <Entrance delay={14} y={30} style={{ marginTop: 44, textAlign: "center" }}>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 118, color: C.text, lineHeight: 1 }}>Beth <span style={{ color: C.gold, fontFamily: F.script, fontStyle: "italic", fontWeight: 500 }}>&</span> Isa</div>
        </Entrance>
        <Entrance delay={24} y={20} style={{ marginTop: 18 }}>
          <div style={{ fontFamily: F.script, fontStyle: "italic", fontSize: 58, color: C.muted }}>mãe e filha, fotógrafas</div>
        </Entrance>
        <div style={{ marginTop: 26 }}><Gold w={260} /></div>
        <Entrance delay={40} y={20} style={{ marginTop: 26 }}>
          <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 36, letterSpacing: "0.22em", color: C.text, textTransform: "uppercase" }}>Guarulhos · SP</div>
        </Entrance>
        <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "center", gap: 18, marginTop: 56, maxWidth: 900 }}>
          {SERVICES.map((s, i) => <Chip key={s} delay={110 + i * 22}>{s}</Chip>)}
        </div>
        <div style={{ position: "absolute", bottom: 260, display: "flex", gap: 26 }}>
          <Photo src="dupla" w={300} h={400} rot={-5} delay={190} drift={0.6} />
          <Photo src="alice" w={300} h={400} rot={3} delay={200} drift={0.6} y={100} />
          <Photo src="quinze" w={300} h={400} rot={-3} delay={210} drift={0.6} />
        </div>
      </AbsoluteFill>
    </Scene>
  );
};

const COL1 = ["escada", "salao", "milho", "concerto"];
const COL2 = ["violino", "cavalo", "alice", "quinze"];
const COL3 = ["natal", "cavalobw", "dupla", "escada"];

const Col: React.FC<{ imgs: string[]; dir: 1 | -1; speed: number; off: number }> = ({ imgs, dir, speed, off }) => {
  const frame = useCurrentFrame();
  const h = 560;
  const total = imgs.length * (h + 24);
  const y = ((frame * speed * dir + off) % total + total) % total;
  const list = [...imgs, ...imgs, ...imgs];
  return (
    <div style={{ width: 330, overflow: "visible", transform: `translateY(${-y - total + 0}px)` }}>
      {list.map((s, i) => (
        <div key={i} style={{ width: 330, height: h, marginBottom: 24, borderRadius: 14, overflow: "hidden", border: `2px solid ${C.gold}88` }}>
          <Img src={staticFile(`img/${s}.jpg`)} style={{ width: "100%", height: "100%", objectFit: "cover" }} />
        </div>
      ))}
    </div>
  );
};

export const B3: React.FC = () => {
  const p = useSpring(2);
  return (
    <Scene>
      <BgBlur src="concerto" />
      <AbsoluteFill style={{ flexDirection: "row", justifyContent: "center", gap: 24, overflow: "hidden", opacity: p, transform: "rotate(-6deg) scale(1.25)" }}>
        <Col imgs={COL1} dir={1} speed={3.2} off={0} />
        <Col imgs={COL2} dir={-1} speed={3.8} off={300} />
        <Col imgs={COL3} dir={1} speed={2.8} off={600} />
      </AbsoluteFill>
      <AbsoluteFill style={{ background: "radial-gradient(ellipse at center, rgba(18,6,10,0.78) 20%, rgba(18,6,10,0.25) 80%)" }} />
      <AbsoluteFill style={{ justifyContent: "center", alignItems: "center", padding: "0 70px" }}>
        <WordReveal text="Cada clique" size={118} delay={14} />
        <WordReveal text="com carinho," size={100} italic delay={28} weight={600} highlight={["carinho,"]} />
        <div style={{ height: 30 }} />
        <WordReveal text="luz e olhar" size={86} delay={92} />
        <WordReveal text="de quem ama o que faz" size={62} delay={110} italic weight={500} />
      </AbsoluteFill>
    </Scene>
  );
};

export const B4: React.FC = () => {
  const frame = useCurrentFrame();
  const pulse = 1 + Math.sin(frame / 7) * 0.025;
  const pb = useSpring(70, "bouncy");
  return (
    <Scene>
      <BgBlur src="dupla" />
      <AbsoluteFill style={{ alignItems: "center", paddingTop: 140 }}>
        <Photo src="dupla" w={640} h={850} rot={-2} delay={0} x={0} y={90} drift={1.2} />
        <Entrance delay={26} y={30} style={{ marginTop: 56, textAlign: "center" }}>
          <div style={{ fontFamily: F.display, fontWeight: 800, fontSize: 84, color: C.text, lineHeight: 1.08 }}>Eternize o seu <span style={{ color: C.gold }}>momento</span></div>
        </Entrance>
        <div style={{ marginTop: 44, transform: `scale(${pb * pulse})`, opacity: Math.min(1, pb) }}>
          <div style={{ padding: "26px 64px", borderRadius: 999, background: "linear-gradient(120deg, #1FAF5A, #25D366)", fontFamily: F.body, fontWeight: 700, fontSize: 46, color: "#fff", boxShadow: "0 20px 60px -10px rgba(37,211,102,0.55)" }}>Chama no WhatsApp</div>
        </div>
        <Entrance delay={100} y={20} style={{ marginTop: 40, textAlign: "center" }}>
          <div style={{ fontFamily: F.body, fontWeight: 600, fontSize: 36, color: C.gold, letterSpacing: "0.04em" }}>@bethmattosretratoseeventos</div>
          <div style={{ fontFamily: F.script, fontStyle: "italic", fontSize: 44, color: C.muted, marginTop: 8 }}>Guarulhos · SP</div>
        </Entrance>
      </AbsoluteFill>
    </Scene>
  );
};
