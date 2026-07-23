# -*- coding: utf-8 -*-
"""
UFISCAL - Análise de Extratos de Cartões de Colaboradores
Consolida múltiplos arquivos XLSX de extrato e gera análises:
  1) Consolidado geral
  2) Ranking de Transferência de Saldo a Crédito por colaborador
  3) Ranking de gastos (Transação Compra) por colaborador
  4) Ranking de gastos por Ramo, individualizado por colaborador
Executar: streamlit run analise_cartoes.py
"""

import io
import re
import unicodedata

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="UFISCAL | Análise de Cartões",
    page_icon="💳",
    layout="wide",
)

# ---------------------------------------------------------------- utilidades

COLUNAS_ESPERADAS = [
    "Data", "Nome", "Número de série", "Tipo de transação",
    "Estabelecimento", "Ramo", "Valor",
]


def parse_valor(v):
    """Converte 'R$ 1.000,00', '-R$ 45,90' e até formatos inconsistentes
    como '-R$ 3,253,60' em float. Regra: últimos 2 dígitos = centavos."""
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v)
    negativo = "-" in s
    digitos = re.sub(r"\D", "", s)
    if not digitos:
        return 0.0
    valor = int(digitos) / 100.0
    return -valor if negativo else valor


def fmt_brl(v):
    """Formata float como moeda brasileira."""
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    sinal = "-" if v < 0 else ""
    return f"{sinal}R$ {s}"


def slug_sheet(nome, usados):
    """Nome de aba válido no Excel (<=31 chars, sem caracteres proibidos)."""
    s = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode()
    s = re.sub(r"[\\/*?:\[\]]", "", s).strip()[:31] or "Colaborador"
    base, i = s, 1
    while s in usados:
        s = f"{base[:28]}_{i}"
        i += 1
    usados.add(s)
    return s


@st.cache_data(show_spinner=False)
def consolidar(arquivos_bytes):
    frames = []
    for nome_arq, conteudo in arquivos_bytes:
        df = pd.read_excel(io.BytesIO(conteudo))
        df.columns = [str(c).strip() for c in df.columns]
        df["Arquivo de origem"] = nome_arq
        frames.append(df)
    df = pd.concat(frames, ignore_index=True)

    faltantes = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            f"Colunas ausentes nos arquivos: {', '.join(faltantes)}"
        )

    df["Valor (R$)"] = df["Valor"].apply(parse_valor)
    df["Data/Hora"] = pd.to_datetime(
        df["Data"], format="%d/%m/%y %H:%M:%S", errors="coerce"
    )
    df = df.sort_values("Data/Hora", ascending=False).reset_index(drop=True)
    df["Nome"] = df["Nome"].astype(str).str.strip()
    df["Ramo"] = df["Ramo"].fillna("(sem ramo)").replace("", "(sem ramo)")
    return df


def gerar_excel(df, transf, compras, por_ramo):
    """Gera o XLSX consolidado com todas as análises em abas."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export = df.drop(columns=["Data/Hora"])
        df_export.to_excel(writer, sheet_name="Consolidado", index=False)
        transf.to_excel(writer, sheet_name="Transferencias Credito", index=False)
        compras.to_excel(writer, sheet_name="Ranking Gastos", index=False)
        usados = {"Consolidado", "Transferencias Credito", "Ranking Gastos"}
        for nome, tabela in por_ramo.items():
            aba = slug_sheet(nome.title(), usados)
            tabela.to_excel(writer, sheet_name=aba, index=False)
        # largura básica de colunas
        for ws in writer.book.worksheets:
            for col in ws.columns:
                largura = max((len(str(c.value)) for c in col if c.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(largura + 2, 40)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------- interface

st.title("💳 Análise de Extratos de Cartões — Colaboradores")
st.caption("Envie os arquivos XLSX de extrato (um por colaborador ou vários). "
           "O sistema consolida tudo e gera as análises abaixo.")

uploads = st.file_uploader(
    "Arquivos de extrato (.xlsx)", type=["xlsx"], accept_multiple_files=True
)

if not uploads:
    st.info("⬆️ Envie ao menos um arquivo XLSX para iniciar.")
    st.stop()

try:
    df = consolidar(tuple((u.name, u.getvalue()) for u in uploads))
except Exception as e:
    st.error(f"Erro ao ler os arquivos: {e}")
    st.stop()

# --------- bases de análise
mask_transf = df["Tipo de transação"].str.contains(
    "Transferência de Saldo a Crédito", case=False, na=False
)
mask_compra = df["Tipo de transação"].str.contains(
    "Transação Compra", case=False, na=False
)

transf = (
    df[mask_transf]
    .groupby("Nome", as_index=False)
    .agg(**{
        "Qtde de Transferências": ("Valor (R$)", "count"),
        "Total Recebido (R$)": ("Valor (R$)", "sum"),
    })
    .sort_values("Total Recebido (R$)", ascending=False)
    .reset_index(drop=True)
)
transf.insert(0, "Ranking", range(1, len(transf) + 1))

compras_df = df[mask_compra].copy()
compras_df["Gasto (R$)"] = compras_df["Valor (R$)"].abs()

compras = (
    compras_df
    .groupby("Nome", as_index=False)
    .agg(**{
        "Qtde de Compras": ("Gasto (R$)", "count"),
        "Total Gasto (R$)": ("Gasto (R$)", "sum"),
        "Ticket Médio (R$)": ("Gasto (R$)", "mean"),
        "Maior Compra (R$)": ("Gasto (R$)", "max"),
    })
    .sort_values("Total Gasto (R$)", ascending=False)
    .reset_index(drop=True)
)
compras.insert(0, "Ranking", range(1, len(compras) + 1))

colaboradores = sorted(df["Nome"].unique())
por_ramo = {}
for nome in colaboradores:
    base = compras_df[compras_df["Nome"] == nome]
    tab = (
        base.groupby("Ramo", as_index=False)
        .agg(**{
            "Qtde de Compras": ("Gasto (R$)", "count"),
            "Total Gasto (R$)": ("Gasto (R$)", "sum"),
        })
        .sort_values("Total Gasto (R$)", ascending=False)
        .reset_index(drop=True)
    )
    total = tab["Total Gasto (R$)"].sum()
    tab["% do Total"] = (tab["Total Gasto (R$)"] / total * 100).round(2) if total else 0
    tab.insert(0, "Ranking", range(1, len(tab) + 1))
    por_ramo[nome] = tab

# --------- download do consolidado
excel_buffer = gerar_excel(df, transf, compras, por_ramo)
st.download_button(
    "⬇️ Baixar XLSX consolidado com todas as análises",
    data=excel_buffer,
    file_name="extratos_consolidados_analises.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# --------- abas
abas = st.tabs([
    "📋 Consolidado",
    "💰 Transferências a Crédito",
    "🛒 Ranking de Gastos",
    "🏷️ Gastos por Ramo (por colaborador)",
])

# ---- Aba 1: Consolidado
with abas[0]:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Transações", len(df))
    c2.metric("Colaboradores", df["Nome"].nunique())
    c3.metric("Total Recebido (Crédito)", fmt_brl(df.loc[mask_transf, "Valor (R$)"].sum()))
    c4.metric("Total Gasto (Compras)", fmt_brl(compras_df["Gasto (R$)"].sum()))

    col_f1, col_f2 = st.columns(2)
    filtro_nome = col_f1.multiselect("Filtrar colaborador", colaboradores)
    tipos = sorted(df["Tipo de transação"].dropna().unique())
    filtro_tipo = col_f2.multiselect("Filtrar tipo de transação", tipos)

    visao = df.copy()
    if filtro_nome:
        visao = visao[visao["Nome"].isin(filtro_nome)]
    if filtro_tipo:
        visao = visao[visao["Tipo de transação"].isin(filtro_tipo)]

    visao_show = visao.drop(columns=["Data/Hora"]).copy()
    visao_show["Valor (R$)"] = visao_show["Valor (R$)"].map(fmt_brl)
    st.dataframe(visao_show, use_container_width=True, hide_index=True)
    st.caption(f"{len(visao)} transações exibidas | "
               f"Saldo líquido do filtro: {fmt_brl(visao['Valor (R$)'].sum())}")

# ---- Aba 2: Transferências
with abas[1]:
    st.subheader("Ranking — Transferência de Saldo a Crédito recebida")
    t_show = transf.copy()
    t_show["Total Recebido (R$)"] = t_show["Total Recebido (R$)"].map(fmt_brl)
    st.dataframe(t_show, use_container_width=True, hide_index=True)
    st.bar_chart(transf.set_index("Nome")["Total Recebido (R$)"])
    st.metric("Total geral transferido", fmt_brl(df.loc[mask_transf, "Valor (R$)"].sum()))

# ---- Aba 3: Ranking de gastos
with abas[2]:
    st.subheader("Ranking — Total gasto em Transação Compra")
    g_show = compras.copy()
    for c in ["Total Gasto (R$)", "Ticket Médio (R$)", "Maior Compra (R$)"]:
        g_show[c] = g_show[c].map(fmt_brl)
    st.dataframe(g_show, use_container_width=True, hide_index=True)
    st.bar_chart(compras.set_index("Nome")["Total Gasto (R$)"])

# ---- Aba 4: por colaborador / ramo
with abas[3]:
    sub = st.tabs([n.title() for n in colaboradores])
    for aba_colab, nome in zip(sub, colaboradores):
        with aba_colab:
            tab = por_ramo[nome]
            c1, c2, c3 = st.columns(3)
            c1.metric("Total gasto", fmt_brl(tab["Total Gasto (R$)"].sum()))
            c2.metric("Compras", int(tab["Qtde de Compras"].sum()))
            recebido = df.loc[mask_transf & (df["Nome"] == nome), "Valor (R$)"].sum()
            c3.metric("Recebido a crédito", fmt_brl(recebido))

            t_show = tab.copy()
            t_show["Total Gasto (R$)"] = t_show["Total Gasto (R$)"].map(fmt_brl)
            t_show["% do Total"] = t_show["% do Total"].map(lambda x: f"{x:.2f}%")
            st.dataframe(t_show, use_container_width=True, hide_index=True)
            st.bar_chart(tab.set_index("Ramo")["Total Gasto (R$)"])
