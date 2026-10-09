import os, json, sys, urllib.request
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "cgSgspJ2msm6clMCkdW9"  # Jessica (modelo v3)
LINES = {
 "i1": "[confident] I, I, A. Inovação, inteligência artificial e automação. Veja como cada alteração chega com segurança até o servidor.",
 "i2": "[confident] Tudo começa em uma branch: bugfix para corrigir, feature para criar. Depois de desenvolver, abrimos um PR para a homolog, onde a alteração é testada e homologada.",
 "i3": "[confident] Quando o autor é o Rafael ou o Gustavo, o PR exige a aprovação do Júlio e do Sidney.",
 "i4": "[confident] Homologado, abrimos um PR da homolog para a development. Após aprovação, outro PR para a main. Com a alteração na main, aplicamos no servidor, copiando, linha por linha, o que foi alterado em cada arquivo.",
 "i5": "[confident] Na Solicitação de Materiais, o fluxo atual é mais direto: da branch para a homolog, e depois de homologado, direto para o servidor.",
 "i6": "[excited] Inovação, inteligência artificial e automação, em cada entrega!",
}
for k, t in LINES.items():
    if len(sys.argv) > 1 and k not in sys.argv[1:]: continue
    body = json.dumps({"text": t, "model_id": "eleven_v3", "voice_settings": {"stability": 0.4, "similarity_boost": 0.75, "style": 0.5}}).encode()
    req = urllib.request.Request(f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE}?output_format=mp3_44100_128", data=body, headers={"xi-api-key": KEY, "Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r, open(f"public/audio/{k}.mp3", "wb") as f: f.write(r.read())
    print(k, "ok")
