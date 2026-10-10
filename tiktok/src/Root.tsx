import React from "react";
import { Composition } from "remotion";
import { Video } from "./Video";
import motiv01 from "./videos/motiv01";

const ALL = [motiv01];

export const Root: React.FC = () => (
  <>
    {ALL.map((c) => (
      <Composition key={c.id} id={c.id} component={Video} defaultProps={{ cfg: c }} durationInFrames={Math.round(c.beats.duration * 30)} fps={30} width={1080} height={1920} />
    ))}
  </>
);
