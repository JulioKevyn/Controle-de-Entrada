import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { BgMesh, FooterBar, Grade, Grain, Mark, Vignette } from "./components";
import { S1, S2, S3, S4, S5V, S5, S6, S8A, S7, S8 } from "./scenes";
import { SCENES, STARTS, TOTAL } from "./theme";

const list = [S1, S2, S3, S4, S5V, S5, S6, S8A, S7, S8];

export const Main: React.FC = () => (
  <AbsoluteFill>
    <BgMesh />
    <FooterBar />
    <Sequence from={STARTS[1]} durationInFrames={TOTAL - STARTS[1] - SCENES[SCENES.length - 1]}>
      <div style={{ position: "absolute", right: 64, top: 44 }}><Mark size={96} shadow={false} /></div>
    </Sequence>
    {list.map((S, i) => (
      <Sequence key={i} from={STARTS[i]} durationInFrames={SCENES[i]}>
        <S />
        <Audio src={staticFile(`audio/s${i + 1}.mp3`)} volume={1} />
        <Sequence from={0}>
          <Audio src={staticFile("audio/whoosh.wav")} volume={0.35} />
        </Sequence>
      </Sequence>
    ))}
    <Audio src={staticFile("audio/music.wav")} volume={0.3} />
    <Grade />
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
