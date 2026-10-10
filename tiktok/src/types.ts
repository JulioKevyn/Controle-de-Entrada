export type Line = { t: string; b: number; d: number };
export type VideoConfig = {
  id: string;
  audio: string;
  beats: { bpm: number; bars: number; duration: number; kick: number[]; snare: number[]; bar: number[] };
  accent: string;
  dropBeat: number;
  lines: Line[];
};
