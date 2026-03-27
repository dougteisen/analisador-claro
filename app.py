import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

st.set_page_config(layout="wide")

# ===== CSS SaaS PROFISSIONAL =====
st.markdown("""
<style>

.main {
    background: linear-gradient(180deg, #0f172a 0%, #020617 100%);
}

h1, h2, h3 {
    color: white;
}

.block-container {
    padding-top: 1.5rem;
}

.upload-box {
    border: 2px dashed #334155;
    border-radius: 16px;
    padding: 40px;
    text-align: center;
    background: #020617;
    transition: 0.3s;
}

.upload-box:hover {
    border-color: #22c55e;
    transform: scale(1.01);
}

.upload-icon {
    font-size: 50px;
    animation: pulse 2s infinite;
}

@keyframes pulse {
    0% {opacity: 0.5;}
    50% {opacity: 1;}
    100% {opacity: 0.5;}
}

.stMetric {
    background: linear-gradient(145deg, #111827, #1f2937);
    padding: 18px;
    border-radius: 14px;
}

.stDownloadButton>button {
    background: linear-gradient(90deg, #16a34a, #22c55e);
    color: white;
    border-radius: 12px;
    height: 55px;
    font-weight: bold;
}

</style>
""", unsafe_allow_html=True)

# ===== HEADER =====
col1, col2 = st.columns([2, 4])

with col1:
    st.image("logo.png", width=240)

with col2:
    st.markdown("# TARGET TELECOM")
    st.markdown("### Inteligência em Faturas Corporativas")

st.markdown("---")

# ===== UPLOAD =====
st.markdown("""
<div class="upload-box">
    <div class="upload-icon">📎</div>
    <h3>Arraste sua fatura ou clique para enviar</h3>
    <p>PDF • Seguro • Processamento automático</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("", type="pdf", accept_multiple_files=True)

# ===== FUNÇÕES =====

def extrair_cliente(texto):
    linhas = texto.split("\n")
    for i, linha in enumerate(linhas):
        if "nº do cliente" in linha.lower():
            if i >= 1:
                nome = linhas[i - 1].strip().upper()
                nome = re.sub(r'[\\/:*?"<>|]', "", nome)
                nome = re.sub(r"\s+", " ", nome)
                return nome
    return "CLIENTE"

def extrair_mensalidades(texto):
    blocos = re.split(r"DETALHAMENTO DE LIGAÇÕES E SERVIÇOS DO CELULAR", texto)
    mapa = {}
    for bloco in blocos:
        num = re.search(r"\(\d{2}\)\s\d{5}\s\d{4}", bloco)
        if not num:
            continue
        linha = num.group(0).replace("(", "").replace(")", "").replace(" ", "")
        total = re.search(r"TOTAL\s*R\$\s*([\d\.,]+)", bloco)
        if total:
            mapa[linha] = total.group(1)
    return mapa

def extrair_pacote_e_passaporte(texto):
    blocos = re.split(r"DETALHAMENTO DE LIGAÇÕES E SERVIÇOS DO CELULAR", texto)
    resultado = {}

    for bloco in blocos:
        num = re.search(r"\(\d{2}\)\s\d{5}\s\d{4}", bloco)
        if not num:
            continue

        linha = num.group(0).replace("(", "").replace(")", "").replace(" ", "")

        pacote = "-"
        passaporte = "-"
        valor_passaporte = "0"

        # 🔥 CORREÇÃO AQUI (SUPORTE LIFE)
        m = re.search(r"(Claro (Pós|Life Ilimitado)\s*\d+GB)", bloco)
        if m:
            pacote = m.group(1)

        for linha_bloco in bloco.split("\n"):
            linha_limpa = linha_bloco.strip()

            if not linha_limpa.startswith("Claro Passaporte"):
                continue
            if "-" in linha_limpa:
                continue

            m = re.search(r"Claro (Passaporte .*?GB)\s+([\d]+,\d{2})$", linha_limpa)
            if m:
                passaporte = m.group(1)
                valor_passaporte = m.group(2)
                break

        resultado[linha] = {
            "Pacote": pacote,
            "Passaporte": passaporte,
            "Valor Passaporte": valor_passaporte
        }

    return resultado

def extrair_detalhamento(texto):
    blocos = re.split(r"DETALHAMENTO DE LIGAÇÕES E SERVIÇOS DO CELULAR", texto)
    mapa = {}

    for bloco in blocos:
        num = re.search(r"\(\d{2}\)\s\d{5}\s\d{4}", bloco)
        if not num:
            continue

        linha = num.group(0).replace("(", "").replace(")", "").replace(" ", "")

        internet = "0"

        m = re.search(r"Internet\s+([\d\.,]+)", bloco, re.IGNORECASE)
        if not m:
            m = re.search(r"Internet.*?([\d\.,]+)", bloco, re.IGNORECASE)

        if m:
            internet = m.group(1)
        else:
            m = re.search(r"Subtotal\s([\d\.,]+)", bloco)
            if m:
                internet = m.group(1)

        minutos = "0"
        m = re.search(r"TOTAL\s([\dminseg:s]+)", bloco)
        if m:
            minutos = m.group(1)

        mapa[linha] = {
            "Internet (MB)": internet,
            "Minutos": minutos
        }

    return mapa

def extrair_linhas(texto):
    linhas = re.findall(r"\(\d{2}\)\s\d{5}\s\d{4}", texto)
    lista = []
    for l in linhas:
        num = l.replace("(", "").replace(")", "").replace(" ", "")
        if num not in lista:
            lista.append(num)
    return lista

def to_float(valor):
    valor = str(valor).replace(".", "").replace(",", ".")
    try:
        return float(valor)
    except:
        return 0

def extrair_gb_pacote(pacote):
    m = re.search(r"(\d+)\s*GB", str(pacote))
    if m:
        return int(m.group(1))
    return 0

def processar_pdf(file):
    texto = ""

    with pdfplumber.open(file) as pdf:

        total_paginas = len(pdf.pages)
        progresso = st.progress(0)
        status = st.empty()

        for i, page in enumerate(pdf.pages):
            status.text(f"📄 Processando página {i+1} de {total_paginas}")

            t = page.extract_text()
            if t:
                texto += t + "\n"

            progresso.progress((i + 1) / total_paginas)

        status.text("🔍 Extraindo dados...")

    cliente = extrair_cliente(texto)
    linhas = extrair_linhas(texto)
    mensalidades = extrair_mensalidades(texto)
    detalhamento = extrair_detalhamento(texto)
    pacotes = extrair_pacote_e_passaporte(texto)

    dados = []

    for linha in linhas:

        total = to_float(mensalidades.get(linha, "0"))
        valor_passaporte = to_float(pacotes.get(linha, {}).get("Valor Passaporte", "0"))

        valor_plano = total - valor_passaporte

        dados.append({
            "Linha": linha,
            "Internet (MB)": detalhamento.get(linha, {}).get("Internet (MB)", "0"),
            "Pacote de dados": pacotes.get(linha, {}).get("Pacote", "-"),
            "Mensalidade": f"R$ {valor_plano:.2f}".replace(".", ","),
            "Passaporte": pacotes.get(linha, {}).get("Passaporte", "-"),
            "Mensalidade Passaporte": f"R$ {valor_passaporte:.2f}".replace(".", ",") if valor_passaporte else "-",
            "Total por linha": f"R$ {total:.2f}".replace(".", ","),
            "Minutos": detalhamento.get(linha, {}).get("Minutos", "0")
        })

    df = pd.DataFrame(dados)

    df["Internet (MB)"] = (
        df["Internet (MB)"]
        .astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    df["Internet (MB)"] = pd.to_numeric(df["Internet (MB)"], errors="coerce").fillna(0)

    def classificar(x):
        if x > 10000:
            return "🔴 Alto"
        elif x > 3000:
            return "🟡 Médio"
        else:
            return "⚪ Baixo"

    df["Perfil"] = df["Internet (MB)"].apply(classificar)

    def em_uso(row):
        minutos = str(row["Minutos"]).lower()
        if row["Internet (MB)"] == 0 and minutos in ["0", "", "0min", "0:00", "0seg"]:
            return "Não"
        return "Sim"

    df["Em Uso"] = df.apply(em_uso, axis=1)

    def estrategia(row):
        if row["Em Uso"] == "Não":
            return "⚪ Manter"

        pacote_gb = extrair_gb_pacote(row["Pacote de dados"])
        uso_gb = row["Internet (MB)"] / 1024 if row["Internet (MB)"] else 0

        if pacote_gb > 0 and uso_gb >= pacote_gb * 0.9:
            return "🔵 Upsell → Aumento recomendado"

        if "Baixo" in row["Perfil"]:
            return "🟡 Sustentar plano"
        if "Médio" in row["Perfil"]:
            return "🟢 Bem dimensionado"
        if "Alto" in row["Perfil"]:
            return "🟢 Bem dimensionado"

        return ""

    df["Estratégia Comercial"] = df.apply(estrategia, axis=1)

    return df, cliente

# ===== EXECUÇÃO =====

if uploaded_files:

    df_total = pd.DataFrame()
    cliente_nome = "CLIENTE"

    progress = st.progress(0)
    total_files = len(uploaded_files)

    for i, file in enumerate(uploaded_files):
        with st.spinner(f"Processando {file.name}..."):
            df, cliente = processar_pdf(file)
            df_total = pd.concat([df_total, df])
            cliente_nome = cliente

        progress.progress((i + 1) / total_files)

    if not df_total.empty:

        col1, col2, col3, col4 = st.columns(4)

        total_linhas = len(df_total)
        em_uso = (df_total["Em Uso"] == "Sim").sum()
        total_gb = df_total["Internet (MB)"].sum() / 1024
        media_gb = total_gb / total_linhas if total_linhas else 0

        col1.metric("Linhas", total_linhas)
        col2.metric("Em uso", em_uso)
        col3.metric("Total GB", round(total_gb, 1))
        col4.metric("Média GB", round(media_gb, 1))

        st.markdown("---")
        st.dataframe(df_total)

        excel = gerar_excel(df_total)

        st.download_button(
            "📥 Baixar Relatório Excel",
            data=excel,
            file_name=f"Analise_Target_{cliente_nome}.xlsx"
        )
