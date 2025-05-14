import os
import pandas as pd
import unicodedata
from tkinter import messagebox

def remover_acentos(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize('NFKD', texto)
    return "".join(c for c in texto if not unicodedata.combining(c)).lower().strip()

def normalizar_base(nome):
    if pd.isna(nome):
        return ""
    nome = str(nome).strip()
    nome = " ".join(nome.split())
    partes = nome.split("-")[0]
    return remover_acentos(partes.strip())

def buscar_emails_generico(emails_df, base_cidade):
    base_cidade = base_cidade.strip().lower()


    def normalizar(texto):
        return (
            str(texto).strip().lower()
            .encode('ascii', errors='ignore')
            .decode('utf-8')
        )

    base_normalizada = normalizar(base_cidade)
    emails_df["base_normalizada"] = emails_df["base"].apply(normalizar)


    emails_exatos = emails_df[emails_df["base_normalizada"] == base_normalizada]
    if not emails_exatos.empty:
        return emails_exatos.iloc[0]["email"]


    primeira_palavra = base_normalizada.split()[0]
    emails_parciais = emails_df[emails_df["base_normalizada"].str.startswith(primeira_palavra)]
    if not emails_parciais.empty:
        return emails_parciais.iloc[0]["email"]


    emails_contem = emails_df[emails_df["base_normalizada"].str.contains(primeira_palavra)]
    if not emails_contem.empty:
        return emails_contem.iloc[0]["email"]

    return None


def carregar_tema(caminho_config, tema_escuro, tema_claro):
    if os.path.exists(caminho_config):
        with open(caminho_config, "r") as f:
            tema = f.read().strip()
            return tema_escuro if tema == "escuro" else tema_claro
    return tema_claro

def salvar_tema(caminho_config, tema, tema_escuro):
    with open(caminho_config, "w") as f:
        f.write("escuro" if tema == tema_escuro else "claro")

def atualizar_estilo_text_area(text_area, cor_negrito):
    text_area.tag_configure("negrito", font=("Arial", 10, "bold"), foreground=cor_negrito)

def aplicar_tema(root, widgets, tema):
    for widget in widgets:
        if "widget" in widget:
            widget["widget"].configure(**widget["config"][tema])
