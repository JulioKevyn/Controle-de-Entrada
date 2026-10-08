# Trilha de fundo procedural (corporativa, ~100 BPM), sem direitos autorais.
import sys, wave
import numpy as np

SR = 44100
BPM = 84
DUR = float(sys.argv[1]) if len(sys.argv) > 1 else 110.0
OUT = sys.argv[2] if len(sys.argv) > 2 else "public/audio/music.wav"
beat = 60 / BPM
n = int(SR * DUR)
t = np.arange(n) / SR
mix = np.zeros(n)

def hz(m): return 440 * 2 ** ((m - 69) / 12)
def add(sig, start):
    i = int(start * SR)
    j = min(n, i + len(sig))
    if i < n: mix[i:j] += sig[: j - i]

def env(L, a, r):
    e = np.ones(L); A = int(a * SR); R = int(r * SR)
    e[:A] = np.linspace(0, 1, A); e[-R:] = np.linspace(1, 0, R)
    return e

# Am - F - C - G (4 compassos cada acorde = 2 compassos)
chords = [(48, [55, 60, 64]), (55, [55, 59, 62]), (57, [57, 60, 64]), (53, [53, 57, 60])]
bar = 4 * beat
rng = np.random.default_rng(7)

def pad(notes, L):
    tt = np.arange(L) / SR
    s = np.zeros(L)
    for m in notes:
        for d in (-0.15, 0.15):
            f = hz(m) * 2 ** (d / 12)
            s += np.sin(2 * np.pi * f * tt) + 0.3 * np.sin(4 * np.pi * f * tt)
    return s / len(notes) / 2 * env(L, 0.6, 0.8)

def pluck(m, L):
    tt = np.arange(L) / SR
    f = hz(m)
    s = np.sin(2 * np.pi * f * tt) + 0.4 * np.sin(4 * np.pi * f * tt) + 0.15 * np.sin(6 * np.pi * f * tt)
    return s * np.exp(-tt * 7)

def kick(L=int(0.25 * SR)):
    tt = np.arange(L) / SR
    f = 110 * np.exp(-tt * 18) + 45
    return np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-tt * 14)

def hat(L=int(0.05 * SR)):
    return rng.standard_normal(L) * np.exp(-np.arange(L) / SR * 90)

total_bars = int(DUR / bar) + 1
for b in range(total_bars):
    root, notes = chords[(b // 2) % 4]
    st = b * bar
    if b % 2 == 0:
        add(pad(notes + [notes[0] + 12], int(2 * bar * SR)) * 0.55, st)
    # baixo
    for k in (0, 2):
        add(np.sin(2 * np.pi * hz(root - 12) * np.arange(int(beat * 1.8 * SR)) / SR) * env(int(beat * 1.8 * SR), 0.01, 0.2) * 0.55, st + k * beat)
    # arpejo em colcheias (entra depois do 1o compasso)
    if b >= 2:
        pat = [0, 1, 2, 1, 2, 1, 2, 1]
        for k, p in enumerate(pat):
            add(pluck(notes[p % 3] + 12, int(0.5 * SR)) * 0.22, st + k * beat / 2)
    # percussão entra a partir do compasso 4
    if b >= 4:
        for k in range(4):
            add(kick() * 0.5, st + k * beat)
            add(hat() * 0.12, st + k * beat + beat / 2)

mix = np.tanh(mix * 0.9)
mix *= env(n, 2.0, 4.0)
mix = mix / np.max(np.abs(mix)) * 0.9
pcm = (mix * 32767).astype(np.int16)
with wave.open(OUT, "wb") as w:
    w.setnchannels(1); w.setsampwidth(2); w.setframerate(SR); w.writeframes(pcm.tobytes())
print("ok", OUT, round(DUR, 1))
