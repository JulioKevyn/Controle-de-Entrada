# Gera uma faixa estilo phonk (original, sem direitos autorais) + grade de batidas para sincronizar o motion.
# uso: python3 scripts/phonk.py <bpm> <compassos> <semente> <saida_wav> <saida_json>
import sys, json, wave
import numpy as np

SR = 44100
BPM = float(sys.argv[1]); BARS = int(sys.argv[2]); SEED = int(sys.argv[3])
OUT, OUTJ = sys.argv[4], sys.argv[5]
rng = np.random.default_rng(SEED)
beat = 60 / BPM; step = beat / 4
n = int(SR * beat * 4 * BARS)
mix = np.zeros(n); drums = np.zeros(n)

def add(buf, sig, t, g=1.0):
    i = int(t * SR); j = min(n, i + len(sig))
    if i < n: buf[i:j] += sig[: j - i] * g

def fft_filter(x, lo=None, hi=None):
    X = np.fft.rfft(x); f = np.fft.rfftfreq(len(x), 1 / SR)
    if lo: X[f < lo] = 0
    if hi: X[f > hi] = 0
    return np.fft.irfft(X, len(x))

def kick():
    t = np.arange(int(SR * 0.45)) / SR
    f = 45 + 110 * np.exp(-t * 28)
    s = np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t * 7)
    s += 0.5 * np.exp(-t * 90) * rng.standard_normal(len(t))
    return np.tanh(s * 2.2)

def clap():
    t = np.arange(int(SR * 0.28)) / SR
    x = fft_filter(rng.standard_normal(len(t)), lo=900, hi=7000)
    e = np.exp(-t * 14) * (1 + 0.8 * (np.mod(t, 0.012) < 0.004))
    return x * e * 0.6

def hat(open_=False):
    L = 0.18 if open_ else 0.05
    t = np.arange(int(SR * L)) / SR
    return fft_filter(rng.standard_normal(len(t)), lo=6500) * np.exp(-t * (18 if open_ else 70)) * 0.25

def cowbell(freq):
    t = np.arange(int(SR * 0.35)) / SR
    s = np.sign(np.sin(2 * np.pi * freq * t)) + np.sign(np.sin(2 * np.pi * freq * 1.505 * t))
    s = fft_filter(s, lo=500, hi=4200)
    return s * np.exp(-t * 11) * 0.28

def b808(freq, L):
    t = np.arange(int(SR * L)) / SR
    f = freq * (1 + 0.6 * np.exp(-t * 18))
    s = np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t * 2.4)
    return np.tanh(s * 3.0) * 0.8

def pad(notes, L):
    t = np.arange(int(SR * L)) / SR
    s = sum(np.sin(2 * np.pi * 440 * 2 ** ((m - 69) / 12) * t * (1 + d)) for m in notes for d in (-0.004, 0.004))
    s = fft_filter(s, hi=1400) / (len(notes) * 2)
    return s * np.minimum(1, t / 0.4) * np.minimum(1, (L - t) / 0.5)

midi = lambda m: 440 * 2 ** ((m - 69) / 12)
# Fá menor: Fm - Db - Ab - Eb (2 compassos cada acorde... 1 compasso aqui)
prog = [(41, [53, 56, 60]), (37, [49, 53, 56]), (44, [56, 60, 63]), (39, [51, 55, 58])]
mel = [(0, 77), (3, 80), (6, 77), (8, 75), (11, 72), (14, 75)]  # (passo, nota)
K, C, H, HO = kick(), clap(), hat(), hat(True)
events = {"kick": [], "snare": [], "bar": []}
for bar in range(BARS):
    t0 = bar * 4 * beat
    drop = bar >= 4
    events["bar"].append(round(t0, 4))
    root, chord = prog[bar % 4]
    add(mix, pad(chord, 4 * beat), t0, 0.5)
    for s_, nt in mel:
        if bar < 2 and s_ not in (0, 8): continue
        add(mix, cowbell(midi(nt - 12 + (0 if bar % 4 < 2 else -2)) * 1.0), t0 + s_ * step, 0.9)
    kicks = [0, 10] if bar < 4 else [0, 7, 10]
    for s_ in kicks:
        add(drums, K, t0 + s_ * step, 1.0); events["kick"].append(round(t0 + s_ * step, 4))
        add(mix, b808(midi(root - 12 if root > 40 else root), 1.4 * beat), t0 + s_ * step, 0.9)
    snares = [4, 12] if drop else ([12] if bar == 3 else [])
    for s_ in snares:
        add(drums, C, t0 + s_ * step, 0.9); events["snare"].append(round(t0 + s_ * step, 4))
    if drop:
        for s_ in range(0, 16, 2): add(drums, H, t0 + s_ * step, 0.8)
        if bar % 2 == 1:
            for s_ in (13, 14, 15): add(drums, H, t0 + s_ * step, 1.0)
        add(drums, HO, t0 + 6 * step, 0.7)
mix = mix + drums * 1.0
mix = np.tanh(mix * 1.1)
# fade curto no fim para loop sem estalo
f = int(0.04 * SR); mix[-f:] *= np.linspace(1, 0, f)
mix = mix / np.max(np.abs(mix)) * 0.89
with wave.open(OUT, "wb") as w:
    w.setnchannels(1); w.setsampwidth(2); w.setframerate(SR)
    w.writeframes((mix * 32767).astype(np.int16).tobytes())
json.dump({"bpm": BPM, "bars": BARS, "duration": n / SR, **events}, open(OUTJ, "w"))
print(f"{n / SR:.2f}s, {len(events['kick'])} kicks")
