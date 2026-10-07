import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { BgMesh, Grain, Vignette } from "./components";
import { S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13 } from "./scenes";
import { SCENES, STARTS } from "./theme";

const list = [S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13];

export const Main: React.FC = () => (
  <AbsoluteFill>
    <BgMesh />
    {list.map((S, i) => (
      <Sequence key={i} from={STARTS[i]} durationInFrames={SCENES[i]}>
        <S />
        <Audio src={staticFile(`audio/s${i + 1}.mp3`)} volume={1} />
        <Audio src={staticFile("audio/whoosh.wav")} volume={0.3} />
      </Sequence>
    ))}
    <Audio src={staticFile("audio/music.wav")} volume={0.3} />
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
