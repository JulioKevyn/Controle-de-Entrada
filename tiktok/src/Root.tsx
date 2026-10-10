import React from "react";
import { Composition } from "remotion";
import { Video } from "./Video";
import motiv01 from "./videos/motiv01";
import { Sun01, BEATS as SB, BPM as SBPM } from "./Sun01";
import { Money01, BEATS as MB, BPM as MBPM } from "./Money01";

export const Root: React.FC = () => (
  <>
    <Composition id={motiv01.id} component={Video} defaultProps={{ cfg: motiv01 }} durationInFrames={Math.round(motiv01.beats.duration * 30)} fps={30} width={1080} height={1920} />
    <Composition id="Sun01" component={Sun01} durationInFrames={Math.round((SB * 60 * 30) / SBPM)} fps={30} width={1080} height={1920} />
    <Composition id="Money01" component={Money01} durationInFrames={Math.round((MB * 60 * 30) / MBPM)} fps={30} width={1080} height={1920} />
  </>
);
