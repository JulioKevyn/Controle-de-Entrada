import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { BgMesh, Captions, Grain, Sparkles, Vignette } from "./components";
import { CAPTIONS } from "./captionsData";
import { S1, S2, S3, S4, S5, S6, S7, S8, S9, SCust, S10, S11, S12, S13 } from "./scenes";
import { SCENES, STARTS } from "./theme";

const list = [S1, S2, S3, S4, S5, S6, S7, S8, S9, SCust, S10, S11, S12, S13];
const AUDIO = ["s1", "s2", "s3", "s4", "s5", "s6", "s7", "s8", "s9", "sp", "s10", "s11", "s12", "s13"];

export const Main: React.FC = () => (
  <AbsoluteFill>
    <BgMesh />
    <Sparkles />
    {list.map((S, i) => (
      <Sequence key={i} from={STARTS[i]} durationInFrames={SCENES[i]}>
        <S />
        <Captions chunks={CAPTIONS[i]} />
        <Audio src={staticFile(`audio/${AUDIO[i]}.mp3`)} volume={1} />
        <Audio src={staticFile("audio/whoosh.wav")} volume={0.35} />
      </Sequence>
    ))}
    <Audio src={staticFile("audio/music.wav")} volume={0.3} />
    <Grain />
    <Vignette />
  </AbsoluteFill>
);
