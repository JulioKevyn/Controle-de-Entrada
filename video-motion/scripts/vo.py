import os, json, sys, urllib.request
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "cgSgspJ2msm6clMCkdW9"  # Jessica (modelo v3)
LINES = {
 "m1": "[excited] Seu negócio em movimento! Vídeos motion que chamam atenção, e fazem o seu cliente parar para assistir.",
 "m2": "[confident] Texto animado, as cores da sua marca, trilha sonora e legendas. Tudo pensado para mostrar o que você faz em poucos segundos.",
 "m3": "[confident] Apresentação de empresa, anúncios para redes sociais, lançamento de produto, abertura de marca, vídeo explicativo. Qualquer vídeo motion que você precisar.",
 "m4": "[confident] Em formato para o computador e para o celular, prontos para Instagram, Reels, Stories, YouTube e para o seu site.",
 "m5": "[excited] Quer ver o seu negócio em movimento? Chama no WhatsApp!",
}
for k, t in LINES.items():
    if len(sys.argv) > 1 and k not in sys.argv[1:]: continue
    body = json.dumps({"text": t, "model_id": "eleven_v3", "voice_settings": {"stability": 0.4, "similarity_boost": 0.75, "style": 0.5}}).encode()
    req = urllib.request.Request(f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE}?output_format=mp3_44100_128", data=body, headers={"xi-api-key": KEY, "Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r, open(f"public/audio/{k}.mp3", "wb") as f: f.write(r.read())
    print(k, "ok")
