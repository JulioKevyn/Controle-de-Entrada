import os, json, sys, urllib.request
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "cgSgspJ2msm6clMCkdW9"  # Jessica (modelo v3)
LINES = {
 "b1": "[warm] Alguns momentos passam rápido demais. A gente transforma eles em memórias.",
 "b2": "[confident] Beth e Isa, mãe e filha, fotógrafas em Guarulhos. Retratos, ensaios, festas de quinze anos, eventos e fotos corporativas.",
 "b3": "[warm] Cada clique feito com carinho, luz e um olhar de quem ama o que faz.",
 "b4": "[excited] Quer eternizar o seu momento? Chama no WhatsApp e garanta a sua data!",
}
for k, t in LINES.items():
    if len(sys.argv) > 1 and k not in sys.argv[1:]: continue
    body = json.dumps({"text": t, "model_id": "eleven_v3", "voice_settings": {"stability": 0.4, "similarity_boost": 0.75, "style": 0.5}}).encode()
    req = urllib.request.Request(f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE}?output_format=mp3_44100_128", data=body, headers={"xi-api-key": KEY, "Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r, open(f"public/audio/{k}.mp3", "wb") as f: f.write(r.read())
    print(k, "ok")
