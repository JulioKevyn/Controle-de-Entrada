import beats from "../beats_motiv01.json";
import type { VideoConfig } from "../types";

// Cada linha entra EXATAMENTE em uma batida (b = batida inicial, d = duração em batidas).
const config: VideoConfig = {
  id: "Motiv01",
  audio: "audio/motiv01.wav",
  beats,
  accent: "#FF2A2A",
  dropBeat: 16,
  lines: [
    { t: "THEY SLEPT.", b: 0, d: 4 },
    { t: "YOU GRINDED.", b: 4, d: 4 },
    { t: "NOBODY CLAPPED.", b: 8, d: 4 },
    { t: "NOBODY CARED.", b: 12, d: 4 },
    { t: "STILL", b: 16, d: 1 },
    { t: "YOU", b: 17, d: 1 },
    { t: "KEPT", b: 18, d: 1 },
    { t: "GOING.", b: 19, d: 1 },
    { t: "NOW THEY", b: 20, d: 2 },
    { t: "ASK", b: 22, d: 2 },
    { t: "HOW", b: 24, d: 1 },
    { t: "DID", b: 25, d: 1 },
    { t: "YOU", b: 26, d: 1 },
    { t: "DO", b: 27, d: 1 },
    { t: "IT?", b: 28, d: 2 },
    { t: "SIMPLE.", b: 30, d: 2 },
    { t: "YOU NEVER", b: 32, d: 4 },
    { t: "STOPPED.", b: 36, d: 4 },
    { t: "KEEP GOING.", b: 40, d: 4 },
    { t: "THEY SLEPT.", b: 44, d: 4 },
  ],
};
export default config;
