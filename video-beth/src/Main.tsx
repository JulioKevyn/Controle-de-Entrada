import React from "react";
import { AbsoluteFill, Audio, Sequence, staticFile } from "remotion";
import { Captions, Grain, Sparkles, Vignette } from "./components";
import { CAPTIONS } from "./captionsData";
import { B1, B2, B3, B4 } from "./scenes";
import { SCENES, STARTS } from "./theme";

const list = [B1, B2, B3, B4];
const AUDIO = ["b1", "b2", "b3", "b4"];

export const Main: React.FC = () => (
  <AbsoluteFill style={{ background: "#12060A" }}>
    {list.map((S, i) => (
      <Sequence key={i} from={STARTS[i]} durationInFrames={SCENES[i]}>
        <S />
        <Sparkles />
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
