import os, json, sys, urllib.request
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "RGymW84CSmfVugnA5tvA"  # Roberta (pt-BR)
LINES = {
 "m1": "[confident] Você sabe quanto o seu mercado lucra em cada prateleira?",
 "m2": "[confident] Conheça o Future PDV, o caixa do seu mercado, rápido de verdade. Funciona no celular, no tablet e no computador.",
 "m3": "[confident] Venda por peso ou por unidade. Leia o código de barras pela câmera do celular, ou pela etiqueta da balança, e receba em PIX, dinheiro e cartão, inclusive dividido.",
 "m4": "[confident] O estoque dá baixa sozinho, por fardo, caixa ou unidade. A validade é controlada por lote, o sistema avisa o que está vencendo, e vende primeiro o lote que vence antes.",
 "m5": "[confident] Fiado na caderneta digital, fornecedores e notas de entrada, perdas, quebras e vencidos. Tudo registrado, no mesmo lugar.",
 "m6": "[excited] E a análise de lucro mostra faturamento, custo e despesas, mês a mês.",
 "m7": "[excited] No plano Master IA, um assistente responde perguntas sobre as suas vendas, em português, no sistema ou no WhatsApp. Mostra o que está parado, o que repor primeiro, e avisa quando o estoque está baixo.",
 "m8": "[confident] Preço mínimo protegido por senha do gerente, e o registro de quem fez cada venda. Mais controle, menos prejuízo.",
 "m9": "[excited] Teste sete dias grátis, sem cartão. Seu mercado merece saber quanto lucra em cada prateleira. Future PDV!",
}
for k, t in LINES.items():
    if len(sys.argv) > 1 and k not in sys.argv[1:]: continue
    body = json.dumps({"text": t, "model_id": "eleven_v3", "voice_settings": {"stability": 0.4, "similarity_boost": 0.75, "style": 0.5}}).encode()
    req = urllib.request.Request(f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE}?output_format=mp3_44100_128", data=body, headers={"xi-api-key": KEY, "Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r, open(f"public/audio/{k}.mp3", "wb") as f: f.write(r.read())
    print(k, "ok")
