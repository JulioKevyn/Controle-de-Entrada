import React from "react";
import { Composition } from "remotion";
import { Main } from "./Main";
import { ShortMain, SHORT_TOTAL } from "./short";
import { FPS, H, TOTAL, W } from "./theme";

export const Root: React.FC = () => (
  <>
    <Composition id="Main" component={Main} durationInFrames={TOTAL} fps={FPS} width={W} height={H} />
    <Composition id="Vertical" component={Main} durationInFrames={TOTAL} fps={FPS} width={1080} height={1920} />
    <Composition id="Short" component={ShortMain} durationInFrames={SHORT_TOTAL} fps={FPS} width={1080} height={1920} />
  </>
);
