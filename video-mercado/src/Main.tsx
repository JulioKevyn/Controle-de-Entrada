import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { BgMesh, Captions, Grain, Sparkles, Vignette } from "./components";
import { CAPTIONS } from "./captionsData";
import { M1, M2, M3, M4, M5, M6, M7, M8, M9 } from "./scenes";
import { SCENES, STARTS } from "./theme";

const list = [M1, M2, M3, M4, M5, M6, M7, M8, M9];

export const Main: React.FC = () => (
  <AbsoluteFill>
    <BgMesh />
    <Sparkles />
    {list.map((S, i) => (
      <Sequence key={i} from={STARTS[i]} durationInFrames={SCENES[i]}>
        <S />
        <Captions chunks={CAPTIONS[i]} />
        <Audio src={staticFile(`audio/m${i + 1}.mp3`)} volume={1} />
        <Audio src={staticFile("audio/whoosh.wav")} volume={0.3} />
      </Sequence>
    ))}
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
