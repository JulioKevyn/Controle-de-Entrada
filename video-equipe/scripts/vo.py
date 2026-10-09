import os, json, sys, urllib.request
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "cgSgspJ2msm6clMCkdW9"  # Jessica (modelo v3)
LINES = {
 "e1": "[confident] I, I, A, Mundial Logistics. Inovação, inteligência artificial e automação. Esta é a nossa equipe.",
 "e2": "[excited] Quarenta e oito automações criadas pela nossa equipe. Processos que dependiam de planilhas, conferências manuais e retrabalho, agora rodam com automação.",
 "e3": "[confident] Toda automação nasce de uma necessidade real da operação. É desenvolvida, passa por homologação e aprovação, e só então chega ao servidor. Qualidade e segurança em cada entrega.",
 "e4": "[confident] Um exemplo real: a Solicitação de Materiais. Cada cliente envia a planilha de um jeito. São dezesseis layouts diferentes e mais de cem depositantes. O sistema reconhece o layout e valida cada linha automaticamente: o documento, o endereço, o estoque e o vínculo do destinatário.",
 "e5": "[excited] E o resultado? O que levava quatro horas e trinta minutos, agora leva vinte e sete. Noventa por cento menos tempo. Mais de mil e novecentos pacotes, e mais de vinte e sete mil destinos.",
 "e6": "[warm] Para o cliente: mais rapidez, menos erros e total transparência. Ele compara as modalidades de frete no portal, e aprova ou recusa em poucos cliques.",
 "e7": "[warm] Para a empresa: o tempo da equipe liberado para o que realmente importa, menos retrabalho, conferência automática e rastreabilidade em cada etapa.",
 "e8": "[confident] Cliente e empresa ganhando juntos. É assim que medimos o sucesso de cada automação.",
 "e9": "[excited] I, I, A, Mundial Logistics. Inovação, inteligência artificial e automação. Fazendo marcas venderem mais!",
}
for k, t in LINES.items():
    if len(sys.argv) > 1 and k not in sys.argv[1:]: continue
    body = json.dumps({"text": t, "model_id": "eleven_v3", "voice_settings": {"stability": 0.4, "similarity_boost": 0.75, "style": 0.5}}).encode()
    req = urllib.request.Request(f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE}?output_format=mp3_44100_128", data=body, headers={"xi-api-key": KEY, "Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r, open(f"public/audio/{k}.mp3", "wb") as f: f.write(r.read())
    print(k, "ok")
