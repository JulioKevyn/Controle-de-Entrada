import sys
import os
import json
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
import win32com.client as win32     


sys.path.append(os.path.dirname(__file__))

from modulos import lg_nf, bayer_nf, heineken_nf, mondelez_nf, whirlpool_nf, danone_nf, diageo_nf, nivea_nf, opella_nf, unilever_nf

clientes = {
    "LG – NF": lg_nf,
    "BAYER – NF": bayer_nf,
    "HEINEKEN – NF": heineken_nf,
    "MONDELEZ – NF": mondelez_nf,
    "WHIRLPOOL – NF": whirlpool_nf,
    "DANONE – NF": danone_nf,
    "DIAGEO – NF": diageo_nf,
    "OPELLA – NF": opella_nf,
    "NIVEA – NF": nivea_nf,
    "UNILEVER – NF": unilever_nf
}

def carregar_senhas():
    caminho = r"W:\CONTROLE DE BASES\001 JULIO\AUTOMAÇÃO\ANOTAÇÕES\ANOTAÇÃO.json"
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

senhas = carregar_senhas()

def abrir_sistema(cliente):
    def verificar_senha(event=None):
        senha = entry_senha.get()
        if senha == senhas.get(cliente, ""):
            janela_senha.destroy()
            root.destroy()
            clientes[cliente].executar_interface_nf()
        else:
            messagebox.showerror("Erro", "Senha incorreta!")
    janela_senha = tk.Toplevel(root)
    janela_senha.title(f"Senha – {cliente}")
    janela_senha.geometry("320x200")
    janela_senha.configure(bg="#1e1e1e")
    tk.Label(
        janela_senha,
        text=f"Senha para {cliente}",
        bg="#1e1e1e", fg="#ee7500",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(20,10))
    entry_senha = tk.Entry(
        janela_senha, show="*", font=("Segoe UI", 12),
        bg="#2a2a2a", fg="white", width=20
    )
    entry_senha.pack(pady=(0,10))
    entry_senha.focus()
    entry_senha.bind("<Return>", verificar_senha)
    tk.Button(
        janela_senha, text="Confirmar",
        bg="#ee7500", fg="white",
        font=("Segoe UI", 12, "bold"),
        command=verificar_senha
    ).pack()

def carregar_imagem(path, size=(32,32)):
    img = Image.open(path).resize(size, Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(img)


def abrir_ajuda():
    help_win = tk.Toplevel(root)
    help_win.title("Suporte")
    help_win.geometry("400x450")
    help_win.configure(bg="#1e1e1e")

    tk.Label(help_win, text="Nome:", bg="#1e1e1e", fg="white").pack(anchor="w", padx=10, pady=(10,0))
    nome_var = tk.StringVar()
    tk.Entry(help_win, textvariable=nome_var, width=40).pack(padx=10, pady=5)

    tk.Label(help_win, text="Print do problema:", bg="#1e1e1e", fg="white").pack(anchor="w", padx=10, pady=(10,0))
    img_path_var = tk.StringVar()
    lbl_img = tk.Label(help_win, text="", bg="#1e1e1e", fg="white")
    lbl_img.pack(padx=10, pady=(0,5))

    def escolher_arquivo():
        p = filedialog.askopenfilename(
            filetypes=[("Imagens","*.png;*.jpg;*.jpeg;*.bmp")]
        )
        if p:
            img_path_var.set(p)
            lbl_img.config(text=os.path.basename(p))

            help_win.lift()
            help_win.focus_force()

    tk.Button(help_win, text="Selecionar imagem", command=escolher_arquivo).pack(padx=10, pady=5)

    tk.Label(help_win, text="Descrição:", bg="#1e1e1e", fg="white").pack(anchor="w", padx=10, pady=(10,0))
    txt_desc = tk.Text(help_win, height=8, width=45, bg="#2a2a2a", fg="white")
    txt_desc.pack(padx=10, pady=5)

    def enviar_suporte():
        nome = nome_var.get().strip()
        desc = txt_desc.get("1.0", "end").strip()
        img_path = img_path_var.get()
        if not nome or not desc or not img_path:
            messagebox.showwarning("Atenção", "Preencha todos os campos e selecione uma imagem.")
            return
        try:
            mail = win32.Dispatch("Outlook.Application").CreateItem(0)
            mail.To = "julio.clemente@mundiallogistics.com.br"
            mail.Subject = f"Suporte: {nome}"
            mail.Body = f"Nome: {nome}\n\nDescrição:\n{desc}"
            mail.Attachments.Add(img_path)
            mail.Send()
            messagebox.showinfo("Sucesso", "Mensagem de suporte enviada!")
            help_win.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao enviar e-mail:\n{e}")

    tk.Button(help_win, text="Enviar", bg="#ee7500", fg="white",
              command=enviar_suporte).pack(pady=10)


root = tk.Tk()
root.title("Sistema Unificado de Cobrança")
root.minsize(400,450)
root.configure(bg="#1e1e1e")

largura = 500
altura = 500
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()
x = (largura_tela // 2) - (largura // 2)
y = (altura_tela // 2) - (altura // 2)
root.geometry(f"{largura}x{altura}+{x}+{y}")


menubar = tk.Menu(root)
helpmenu = tk.Menu(menubar, tearoff=0)
helpmenu.add_command(label="Informações", command=lambda: messagebox.showinfo(
    "Informações",
    "Para o sistema funcionar corretamente:\n\n"
    "- Atualize a planilha de pedidos mensalmente com os dados mais recentes;\n"
    "- Mantenha o nome da planilha como 'ATUAL' (ou conforme configuração do cliente);\n"
    "- Certifique-se de que a planilha de e-mails contenha os e-mails corretos do respectivo cliente.\n\n"
    "Qualquer dúvida, acesse o menu Ajuda / Suporte."
))
helpmenu.add_command(label="Ajuda / Suporte", command=abrir_ajuda)
menubar.add_cascade(label="Ajuda", menu=helpmenu)
root.config(menu=menubar)


tk.Label(root, text="Selecione o cliente",
         bg="#1e1e1e", fg="#ee7500",
         font=("Segoe UI", 18, "bold")).pack(pady=(20,5))


search_var = tk.StringVar()
def filtrar(*_):
    termo = search_var.get().lower()
    for nome, btn in botoes.items():
        btn.pack_forget()
        if termo in nome.lower():
            btn.pack(pady=8, padx=50, fill="x", expand=True, ipady=15)

search_entry = tk.Entry(root, textvariable=search_var, font=("Segoe UI",12),
                        bg="white", fg="black", insertbackground="black")
search_entry.pack(fill="x", padx=50, pady=(0,10))
search_entry.insert(0, "Pesquisar...")
def on_focus_in(e):
    if search_entry.get() == "Pesquisar...":
        search_entry.delete(0, "end")
search_entry.bind("<FocusIn>", on_focus_in)
search_var.trace_add("write", filtrar)


container = tk.Frame(root, bg="#1e1e1e")
container.pack(fill="both", expand=True, padx=10, pady=10)
canvas = tk.Canvas(container, bg="#1e1e1e", highlightthickness=0)
vsb = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=vsb.set)
vsb.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)
buttons_frame = tk.Frame(canvas, bg="#1e1e1e")
win_id = canvas.create_window((0,0), window=buttons_frame, anchor="nw")
def on_canvas_resize(e):
    canvas.itemconfig(win_id, width=e.width)
canvas.bind("<Configure>", on_canvas_resize)
buttons_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))


icones = {nome: carregar_imagem(f"modulos/img/{nome.lower()}.png") for nome in clientes}


botoes = {}
for nome in clientes:
    btn = tk.Button(
        buttons_frame,
        text="  "+nome,
        image=icones[nome],
        compound="left",
        anchor="w",
        padx=10,
        bg="#ee7500", fg="white",
        font=("Segoe UI",12,"bold"),
        activebackground="#ff8c1a",
        relief="flat", bd=0,
        command=lambda n=nome: abrir_sistema(n)
    )
    btn.bind("<Enter>", lambda e, b=btn: b.configure(bg="#ff8c1a"))
    btn.bind("<Leave>", lambda e, b=btn: b.configure(bg="#ee7500"))
    botoes[nome] = btn
    btn.pack(pady=8, padx=50, fill="x", expand=True, ipady=15)

root.mainloop()
