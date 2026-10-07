import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { BgMesh, Grade, Grain, Vignette } from "./components";
import { S1, S2, S3, S4, S5, S6, S7, S8 } from "./scenes";
import { SCENES, STARTS, TOTAL } from "./theme";

const list = [S1, S2, S3, S4, S5, S6, S7, S8];

export const Main: React.FC = () => (
  <AbsoluteFill>
    <BgMesh />
    {list.map((S, i) => (
      <Sequence key={i} from={STARTS[i]} durationInFrames={SCENES[i]}>
        <S />
        <Audio src={staticFile(`audio/s${i + 1}.mp3`)} volume={1} />
        <Sequence from={0}>
          <Audio src={staticFile("audio/whoosh.wav")} volume={0.35} />
        </Sequence>
      </Sequence>
    ))}
    <Audio src={staticFile("audio/music.wav")} volume={0.22} />
    <Grade />
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
