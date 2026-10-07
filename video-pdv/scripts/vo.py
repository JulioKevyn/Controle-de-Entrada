import os, json, sys, urllib.request
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "cgSgspJ2msm6clMCkdW9"  # Jessica (modelo v3)
LINES = {
 "s1": "[excited] Apresentamos o Future PDV: o sistema de vendas, estoque e gestão da sua loja, agora com inteligência artificial.",
 "s2": "[confident] Você vende o dia inteiro, mas no fim do mês não sabe quanto realmente sobrou. Planilha, caderno, estoque que nunca bate.",
 "s3": "[confident] No balcão, a venda sai em segundos. Bipe o código de barras, misture formas de pagamento, e o recibo sai na hora.",
 "s4": "[confident] O caixa fecha certo. Abertura, sangria e fechamento registrados, e cada vendedor entra com o próprio PIN.",
 "s5": "[confident] O estoque é por cor e tamanho. Vendeu a preta G, baixa só a preta G. E o sistema avisa antes de acabar.",
 "s6": "[confident] E você vê o lucro de verdade: faturamento, custo da mercadoria e despesas separados, mês a mês.",
 "s7": "[confident] Descubra quando e por que você vende, com os horários mais fortes e o ranking dos produtos.",
 "s8": "[excited] E aqui entra a inteligência artificial! Pergunte em português para a sua loja: quanto eu vendi, o que está parado. Ela responde com os números reais, e ainda sugere o que fazer.",
 "s9": "[excited] A IA cria insights sozinha, mostra o que repor primeiro, indica os produtos que mais vão vender, e manda o resumo do dia e os alertas de estoque direto no seu WhatsApp.",
 "s10": "[confident] Tem mais de uma loja? Cada filial com caixa, estoque e equipe próprios, tudo numa conta só, com backup automático e registro de quem fez o quê.",
 "s11": "[warm] Quem usa aprova. Ver o lucro de cada produto, e não só o faturamento, muda a forma de administrar a loja.",
 "s12": "[confident] Comece com sete dias grátis, sem cartão. Escolha o Básico, o Pro, ou o Master I A, com toda a inteligência artificial.",
 "sp": "[excited] E ainda fica com a cara da sua marca! Escolha o logo, as cores e o tema, e o sistema inteiro muda para o estilo da sua loja.",
 "si": "[confident] Vende online também? No Master I A, o Future PDV se integra com TikTok Shop, Nuvemshop, Shopify e outros.",
 "s13": "[excited] Future PDV. Sua loja merece saber quanto realmente lucra. Teste grátis por sete dias!",
}
os.makedirs("public/audio", exist_ok=True)
for k, t in LINES.items():
    if len(sys.argv) > 1 and k not in sys.argv[1:]: continue
    body = json.dumps({"text": t, "model_id": "eleven_v3",
        "voice_settings": {"stability": 0.4, "similarity_boost": 0.75, "style": 0.5}}).encode()
    req = urllib.request.Request(f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE}?output_format=mp3_44100_128",
        data=body, headers={"xi-api-key": KEY, "Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r, open(f"public/audio/{k}.mp3", "wb") as f:
        f.write(r.read())
    print(k, "ok")
