import os, json, urllib.request, sys
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "EXAVITQu4vr4xnSDxMaL"  # Sarah
LINES = {
 "s1": "Mundial Logistics apresenta: o sistema de solicitação de materiais e controle de pedidos.",
 "s2": "Dez clientes. Dezenas de planilhas. Cada uma com seu formato, suas colunas e suas datas.",
 "s3": "Agora, cada cliente tem seu próprio módulo, protegido por senha. Escolha a base e entre.",
 "s4": "O sistema lê as planilhas automaticamente, trata colunas e datas, e separa o que está pendente.",
 "s5": "Filtre por base ou cidade, e veja apenas o que precisa de ação.",
 "s6": "Com um clique, o Outlook dispara a solicitação para cada responsável, já com a lista de notas.",
 "s7": "Menos retrabalho. Mais controle. Mais prazo cumprido.",
 "s8": "Mundial Logistics. Fazendo marcas venderem mais.",
}
os.makedirs("public/audio", exist_ok=True)
for k, t in LINES.items():
    if len(sys.argv) > 1 and k not in sys.argv[1:]: continue
    body = json.dumps({"text": t, "model_id": "eleven_multilingual_v2",
        "voice_settings": {"stability": 0.5, "similarity_boost": 0.75, "style": 0.2}}).encode()
    req = urllib.request.Request(f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE}?output_format=mp3_44100_128",
        data=body, headers={"xi-api-key": KEY, "Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r, open(f"public/audio/{k}.mp3", "wb") as f:
        f.write(r.read())
    print(k, "ok")
