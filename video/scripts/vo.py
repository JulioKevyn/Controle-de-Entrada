import os, json, urllib.request, sys
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "EXAVITQu4vr4xnSDxMaL"  # Sarah
LINES = {
 "s1": "Mundial Logistics apresenta: o Sistema de Solicitação de Materiais.",
 "s2": "Cada cliente envia sua planilha em um layout diferente. Conferir CNPJ, endereço e estoque manualmente toma tempo e abre espaço para erro.",
 "s3": "Agora tudo começa no envio da planilha. Você escolhe a finalidade, o cliente, e o sistema reconhece o modelo sozinho.",
 "s4": "Antes mesmo de enviar, a validação roda em segundo plano: Receita Federal, endereço, estoque e cadastro, tudo em tempo real.",
 "s5": "Encontrou uma pendência? O sistema pergunta, você escolhe a correção, e ele revalida na hora.",
 "s6": "Depois, o orçamento segue no quadro de acompanhamento, da elaboração até a aprovação do cliente.",
 "s7": "Aprovado, a operação assume: distribuição, romaneios, coletas, retiradas e descarte, tudo no mesmo portal.",
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
