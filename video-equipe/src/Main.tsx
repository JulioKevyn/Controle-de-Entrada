import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { BgMesh, Captions, Grain, Sparkles, Vignette } from "./components";
import { CAPTIONS } from "./captionsData";
import { E1, E2, E3, E4, E5, E6, E7, E8, E9 } from "./scenes";
import { SCENES, STARTS } from "./theme";

const list = [E1, E2, E3, E4, E5, E6, E7, E8, E9];
const AUDIO = ["e1","e2","e3","e4","e5","e6","e7","e8","e9"];

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
