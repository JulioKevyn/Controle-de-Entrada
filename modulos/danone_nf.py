import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, scrolledtext
import win32com.client as win32
import os

from .utils import carregar_tema, atualizar_estilo_text_area, buscar_emails_generico
from .config import CONFIG

def executar_interface_nf():
    file_path = CONFIG["danone_nf"]["file_path"]
    sheet = CONFIG["danone_nf"]["sheet"]
    header = CONFIG["danone_nf"].get("header", 1)
    emails_df_path = CONFIG["danone_nf"]["emails_path"]
    log_path = CONFIG["log_path"]
    historico_path = CONFIG["historico_path"]
    config_path = CONFIG["config_tema_path"]

    try:
        df = pd.read_excel(file_path, sheet_name=sheet, header=header)

        df.columns = df.columns.str.strip().str.lower().str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

        df = df[df["data da entrega"].notna() & df["data da entrada"].isna()]
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir a planilha:\n{e}")
        return

    try:
        emails_df = pd.read_excel(emails_df_path)
        emails_df.columns = emails_df.columns.str.strip().str.lower()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir planilha de e-mails:\n\n{str(e)}")
        return

    df.rename(columns={
        "Data da Entrada": "data da entrada",
        "data classificacao": "data classificacao",
        "Data da entrega": "data da entrega"
    }, inplace=True)

    if df.empty:
        messagebox.showinfo("Aviso", "Nenhuma NF entregue sem entrada encontrada.")
        return

    df["base/cidade"] = df["base"].astype(str) + " - " + df["dep"].astype(str)
    agrupado = df.groupby(df["base/cidade"].str.lower().str.strip())

    def enviar_emails():
        if not messagebox.askyesno("Confirmação", "Deseja enviar os e-mails agora?"):
            return

        try:
            outlook = win32.Dispatch("Outlook.Application")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o Outlook:\n\n{e}")
            return

        log_linhas = []

        for base, group in agrupado:
            emails = buscar_emails_generico(emails_df, base)
            if not emails:
                continue

            corpo_lista = ""
            for _, row in group.iterrows():
                corpo_lista += f"- Pedido {row['pedido']} – NF {row['nf']} – Entregue em {row['data da entrega'].strftime('%d/%m/%Y')}\n"

            corpo_email = f"""
Olá,

Conforme nossa análise, identificamos que as seguintes notas fiscais foram entregues, porém ainda não constam como baixadas no sistema:

{corpo_lista}
Por gentileza, verificar a baixa dessas notas fiscais e nos retornar com um posicionamento.

Atenciosamente,
Equipe Mundial Logistics
"""

            enviados_para = []
            for email in emails.split(";"):
                email = email.strip()
                if not email or email in enviados_para:
                    continue
                enviados_para.append(email)
                try:
                    mail = outlook.CreateItem(0)
                    mail.To = email
                    mail.CC = "julio.clemente@mundiallogistics.com.br"
                    mail.Subject = f"[NF] Notas entregues sem baixa – {group.iloc[0]['base/cidade']}"
                    mail.Body = corpo_email
                    mail.Send()
                    log_linhas.append(f"{datetime.now():%d/%m/%Y %H:%M:%S} | {group.iloc[0]['base/cidade']} -> {email}")
                except Exception as ex:
                    log_linhas.append(f"{datetime.now():%d/%m/%Y %H:%M:%S} | ERRO {email}: {ex}")

        if log_linhas:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write("\n".join(log_linhas) + "\n")

            historico_data = [{
                "Data envio": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "Base": row["base/cidade"],
                "Pedido": row["pedido"],
                "NF": row["nf"],
                "Data da entrega": row["data da entrega"],
                "E-mails": buscar_emails_generico(emails_df, row["base/cidade"])
            } for _, row in df.iterrows()]
            historico_df = pd.DataFrame(historico_data)

            if os.path.exists(historico_path):
                try:
                    antigo = pd.read_excel(historico_path)
                    historico_df = pd.concat([antigo, historico_df], ignore_index=True)
                except:
                    pass
            historico_df.to_excel(historico_path, index=False)

        messagebox.showinfo("Sucesso", "E-mails enviados com sucesso!")
        root.destroy()

    # Interface visual
    tema_claro = {"bg": "#ffffff", "fg": "#000000", "area_texto": "#ffffff", "texto_fg": "#000000", "negrito": "#ee7500"}
    tema_escuro = {"bg": "#1e1e1e", "fg": "#ffffff", "area_texto": "#2e2e2e", "texto_fg": "#ffffff", "negrito": "#ee7500"}
    tema = carregar_tema(config_path, tema_escuro, tema_claro)

    root = tk.Tk()
    root.title("Cobrança de NF – Danone")
    root.configure(bg=tema["bg"])

    tk.Label(root, text="Notas entregues sem baixa:", font=("Arial", 12, "bold"), bg=tema["bg"], fg=tema["negrito"]).pack(pady=(10,0))

    filtro_frame = tk.Frame(root, bg=tema["bg"])
    filtro_frame.pack(pady=5)
    tk.Label(filtro_frame, text="Filtrar base:", bg=tema["bg"], fg=tema["fg"]).pack(side=tk.LEFT)
    filtro_entry = tk.Entry(filtro_frame, width=30)
    filtro_entry.pack(side=tk.LEFT, padx=5)

    def preencher_text_area(filtro=""):
        text_area.delete("1.0", tk.END)
        for base, group in agrupado:
            if filtro and filtro.lower() not in base:
                continue
            for _, row in group.iterrows():
                emails = buscar_emails_generico(emails_df, row["base/cidade"])
                text_area.insert(tk.END, "BASE: ", "negrito")
                text_area.insert(tk.END, row["base/cidade"].upper(), "negrito")
                text_area.insert(tk.END, f" | PEDIDO: {row['pedido']}")
                text_area.insert(tk.END, f" | NF: {row['nf']}")
                text_area.insert(tk.END, f" | ENTREGA: {row['data da entrega'].strftime('%d/%m/%Y')} | EMAILS: {emails if emails else 'Não encontrado'}\n")

    tk.Button(filtro_frame, text="Filtrar", bg=tema["negrito"], fg="white", command=lambda: preencher_text_area(filtro_entry.get())).pack(side=tk.LEFT, padx=5)
    tk.Button(filtro_frame, text="Limpar Filtro", bg="#888888", fg="white", command=lambda: [filtro_entry.delete(0, tk.END), preencher_text_area()]).pack(side=tk.LEFT, padx=5)

    text_area = scrolledtext.ScrolledText(root, width=90, height=20, bg=tema["area_texto"], fg=tema["texto_fg"], font=("Arial", 10))
    text_area.pack(padx=10, pady=10)
    atualizar_estilo_text_area(text_area, tema["negrito"])

    preencher_text_area()

    button_frame = tk.Frame(root, bg=tema["bg"])
    button_frame.pack(pady=(0,10))
    tk.Button(button_frame, text="Confirmar Envio", command=enviar_emails, font=("Arial", 10, "bold"), bg=tema["negrito"], fg="white").pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="Cancelar", command=root.destroy, font=("Arial", 10), bg="#cccccc", fg="black").pack(side=tk.RIGHT, padx=10)

    root.mainloop()
