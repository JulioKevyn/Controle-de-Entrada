import React from "react";
import { Composition } from "remotion";
import { Main } from "./Main";
import { Capa } from "./Capa";
import { Post1, Post2, Post3, Post4, Post5, Post6 } from "./Posts";
import { FPS, TOTAL } from "./theme";

export const Root: React.FC = () => (
  <>
    <Composition id="Vertical" component={Main} durationInFrames={TOTAL} fps={FPS} width={1080} height={1920} />
    <Composition id="Horizontal" component={Main} durationInFrames={TOTAL} fps={FPS} width={1920} height={1080} />
    <Composition id="Capa" component={Capa} durationInFrames={1} fps={30} width={1080} height={1350} />
    <Composition id="Post1" component={Post1} durationInFrames={1} fps={30} width={1080} height={1350} />
    <Composition id="Post2" component={Post2} durationInFrames={1} fps={30} width={1080} height={1350} />
    <Composition id="Post3" component={Post3} durationInFrames={1} fps={30} width={1080} height={1350} />
    <Composition id="Post4" component={Post4} durationInFrames={1} fps={30} width={1080} height={1350} />
    <Composition id="Post5" component={Post5} durationInFrames={1} fps={30} width={1080} height={1350} />
    <Composition id="Post6" component={Post6} durationInFrames={1} fps={30} width={1080} height={1350} />
  </>
);
