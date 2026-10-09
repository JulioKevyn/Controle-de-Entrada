import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { BgMesh, Captions, Grain, Sparkles, Vignette } from "./components";
import { CAPTIONS } from "./captionsData";
import { I1, I2, I3, I4, I5, I6 } from "./scenes";
import { SCENES, STARTS } from "./theme";

const list = [I1, I2, I3, I4, I5, I6];
const AUDIO = ["i1", "i2", "i3", "i4", "i5", "i6"];

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
    <Audio src={staticFile("audio/music.wav")} volume={0.3} />
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
