# Gera frames por cena e legendas (timing aproximado, proporcional à fala) a partir dos áudios.
import json, math, re, subprocess, sys
sys.path.insert(0, "scripts")
ORDER = ["b1","b2","b3","b4"]
src = open("scripts/vo.py", encoding="utf-8").read()
LINES = {}
for m in re.finditer(r'"(\w+)":\s*"((?:[^"\\]|\\.)*)"', src):
    LINES[m.group(1)] = m.group(2)

def dur(k):
    return float(subprocess.check_output(["ffprobe","-v","error","-show_entries","format=duration","-of","csv=p=0",f"public/audio/{k}.mp3"]))

def speech_span(k, d):
    out = subprocess.run(["ffmpeg","-v","info","-i",f"public/audio/{k}.mp3","-af","silencedetect=noise=-35dB:d=0.12","-f","null","-"],capture_output=True,text=True).stderr
    s0, s1 = 0.0, d
    starts = [float(x) for x in re.findall(r"silence_start: ([0-9.]+)", out)]
    ends = [float(x) for x in re.findall(r"silence_end: ([0-9.]+)", out)]
    if starts and starts[0] < 0.05 and ends: s0 = ends[0]
    if starts and starts[-1] > d - 1.0 and (len(ends) < len(starts) or ends[-1] > d - 0.05): s1 = starts[-1]
    return s0, s1

frames, caps = [], []
for k in ORDER:
    d = dur(k)
    frames.append(math.ceil((d + 0.7) * 30))
    text = re.sub(r"\[[^\]]*\]\s*", "", LINES[k])
    words = text.split()
    wts = []
    for w in words:
        wt = len(re.sub(r"\W", "", w)) + 1.0
        if w.endswith((".", "!", "?")): wt += 5
        elif w.endswith((",", ":", ";")): wt += 2.5
        wts.append(wt)
    s0, s1 = speech_span(k, d)
    tot = sum(wts); acc = 0; times = []
    for w, wt in zip(words, wts):
        times.append((s0 + (s1 - s0) * acc / tot, s0 + (s1 - s0) * (acc + wt) / tot)); acc += wt
    chunks, cur = [], []
    for i, w in enumerate(words):
        cur.append(i)
        if len(cur) >= 6 or w.endswith((".", "!", "?", ",", ":")) and len(cur) >= 3 or i == len(words) - 1:
            chunks.append(cur); cur = []
    if len(chunks) > 1 and len(chunks[-1]) < 3:
        chunks[-2] += chunks.pop()
    out = []
    for ci, ch in enumerate(chunks):
        a = times[ch[0]][0]
        b = times[chunks[ci + 1][0]][0] if ci + 1 < len(chunks) else min(d, times[ch[-1]][1] + 0.35)
        out.append({"t": " ".join(words[i] for i in ch), "a": round(a, 2), "b": round(b, 2)})
    caps.append(out)
open("src/captionsData.ts", "w", encoding="utf-8").write("export const CAPTIONS: { t: string; a: number; b: number }[][] = " + json.dumps(caps, ensure_ascii=False) + ";\n")
th = open("src/theme.ts", encoding="utf-8").read()
th = re.sub(r"export const SCENES = \[.*?\] as const;", "export const SCENES = [" + ", ".join(map(str, frames)) + "] as const;", th)
open("src/theme.ts", "w", encoding="utf-8").write(th)
print(frames, sum(frames) / 30)
