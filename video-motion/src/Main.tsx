import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { BgMesh, Captions, Grain, Sparkles, Vignette } from "./components";
import { CAPTIONS } from "./captionsData";
import { M1, M2, M3, M4, M5 } from "./scenes";
import { SCENES, STARTS, TOTAL } from "./theme";

const list = [M1, M2, M3, M4, M5];
const AUDIO = ["m1", "m2", "m3", "m4", "m5"];

export const Main: React.FC = () => (
  <AbsoluteFill>
    <BgMesh />
    <Sparkles />
    {list.map((S, i) => (
      <Sequence key={i} from={STARTS[i]} durationInFrames={SCENES[i]}>
        <S />
        <Captions chunks={CAPTIONS[i]} />
        <Audio src={staticFile(`audio/${AUDIO[i]}.mp3`)} volume={1} />
        <Audio src={staticFile("audio/whoosh.wav")} volume={0.4} />
      </Sequence>
    ))}
    <Audio src={staticFile("audio/music.wav")} volume={0.34} />
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
