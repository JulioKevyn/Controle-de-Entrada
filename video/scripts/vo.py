import os, json, urllib.request, sys
KEY = os.environ["ELEVEN_API_KEY"]
VOICE = "cgSgspJ2msm6clMCkdW9"  # Jessica (modelo v3)
LINES = {
 "s1": "[confident] Mundial Logística apresenta: o Sistema de Solicitação de Materiais.",
 "s2": "[confident] Cada cliente envia a planilha do seu jeito. São dezesseis layouts diferentes, e mais de cem depositantes, todos precisando de conferência.",
 "s3": "[confident] Você entra com a sua conta Microsoft, escolhe a finalidade, o cliente, e anexa a planilha. O sistema reconhece o layout sozinho.",
 "s4": "[confident] E antes mesmo de enviar, a validação roda em segundo plano, em tempo real. Nada é gravado até você confirmar.",
 "s5": "[confident] Cada linha passa por quatro conferências. O CNPJ ou CPF, ativo na Receita Federal. O CEP e o endereço, comparando planilha, cadastro e Sintegra. O saldo de estoque, por produto e por classe. E o vínculo do destinatário com o depositante.",
 "s6": "[warm] Encontrou uma pendência? O sistema pergunta, você escolhe a correção, e ele revalida na hora.",
 "s7": "[confident] Validado o pedido, nasce o pacote e a planilha de orçamento. O time preenche as modalidades de frete, anexa o PDF, e o cliente recebe um e-mail com o valor.",
 "s8": "[confident] O cliente compara as modalidades no portal, e aprova ou recusa, informando o motivo. Custos extras e manuseio tabelado entram no mesmo fluxo.",
 "s9": "[excited] Aprovado, a operação assume: distribuição de materiais, romaneios, coletas, retiradas agendadas pelo WhatsApp, descarte e alertas de vencimento. Tudo no mesmo portal.",
 "s10": "[excited] Mundial Logística. Fazendo marcas venderem mais!",
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
