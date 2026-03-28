import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

st.set_page_config(
    layout="wide",
    page_title="Target Telecom · Análise de Faturas",
    page_icon="📡"
)

# ===== CSS REDESIGN =====
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background: #060d1a !important;
}
.main {
    background: #060d1a !important;
}

.block-container {
    padding-top: 2rem !important;
    padding-bottom: 3rem !important;
    max-width: 1400px !important;
}

*, h1, h2, h3, h4, p, span, div, label {
    font-family: 'DM Sans', sans-serif !important;
    color: #e2e8f0;
}

.tt-section-title {
    font-family: 'Syne', sans-serif !important;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #475569 !important;
    margin: 8px 0 16px 2px;
}

.tt-divider {
    height: 1px;
    background: linear-gradient(90deg, rgba(16,185,129,0.3), rgba(59,130,246,0.15), transparent);
    border: none;
    margin: 24px 0;
}
</style>
""", unsafe_allow_html=True)

# ===== HEADER =====
st.markdown("""
<div class="tt-header">
    <div class="tt-logo-placeholder">📡</div>
    <div class="tt-brand">
        <h1>TARGET TELECOM</h1>
        <p>Inteligência em Faturas Corporativas</p>
    </div>
    <div class="tt-badge">✦ Análise Automática</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="tt-divider"></div>', unsafe_allow_html=True)

# ===== UPLOAD =====
st.markdown('<p class="tt-section-title">📂 Carregar Fatura</p>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Arraste o PDF da fatura ou clique para selecionar — múltiplos arquivos suportados",
    type="pdf",
    accept_multiple_files=True,
)

# ===== FUNÇÕES (BASE INTACTA) =====

def normalizar_numero(num_str: str) -> str:
    return num_str.replace("(", "").replace(")", "").replace(" ", "")

def extrair_blocos_por_linha(texto: str) -> dict:
    blocos = re.split(r"DETALHAMENTO DE LIGAÇÕES E SERVIÇOS DO CELULAR", texto)
    resultado = {}
    for bloco in blocos:
        num = re.search(r"\(\d{2}\)\s\d{5}\s\d{4}", bloco)
        if num:
            chave = normalizar_numero(num.group(0))
            resultado[chave] = bloco
    return resultado

def extrair_cliente(texto: str) -> str:
    linhas = texto.split("\n")
    for i, linha in enumerate(linhas):
        if "nº do cliente" in linha.lower():
            if i >= 1:
                nome = linhas[i - 1].strip().upper()
                nome = re.sub(r'[\\/:*?"<>|]', "", nome)
                nome = re.sub(r"\s+", " ", nome)
                return nome
    return "CLIENTE"

def extrair_vencimento(texto: str) -> str:
    match = re.search(r"Nº da conta:.*?(\d{2}/\d{2}/\d{4})", texto, re.DOTALL)
    if match:
        return match.group(1)
    match = re.search(r"Vencimento\s*\n\s*(\d{2}/\d{2}/\d{4})", texto)
    if match:
        return match.group(1)
    return ""

def extrair_linhas(texto: str) -> list:
    linhas = re.findall(r"\(\d{2}\)\s\d{5}\s\d{4}", texto)
    lista = []
    for l in linhas:
        num = normalizar_numero(l)
        if num not in lista:
            lista.append(num)
    return lista

def extrair_mensalidades(blocos: dict) -> dict:
    mapa = {}
    for linha, bloco in blocos.items():
        total_m = re.search(r"TOTAL\s*R\$\s*([\d\.,]+)", bloco)
        total_pdf = float(total_m.group(1).replace(".", "").replace(",", ".")) if total_m else 0.0

        secao = re.search(r"Mensalidades e Pacotes Promocionais(.*?)TOTAL\s*R\$", bloco, re.DOTALL)
        soma_positivos = 0.0
        if secao:
            for lb in secao.group(1).split("\n"):
                m = re.search(r"(-?[\d]+,\d{2})$", lb.strip())
                if m:
                    try:
                        val = float(m.group(1).replace(".", "").replace(",", "."))
                        if val > 0:
                            soma_positivos += val
                    except:
                        pass

        if soma_positivos > total_pdf + 0.01:
            mapa[linha] = f"{soma_positivos:.2f}".replace(".", ",")
        elif total_m:
            mapa[linha] = total_m.group(1)

    return mapa

def extrair_pacote_e_passaporte(blocos: dict) -> dict:
    resultado = {}
    for linha, bloco in blocos.items():
        pacote = "-"
        passaporte = "-"
        valor_passaporte = "0"

        m = re.search(
            r"(Claro\s+(?:Pós|Life Ilimitado|Controle)\s+\d+\s*GB)",
            bloco
        )
        if m:
            pacote = m.group(1).strip()

        secao = re.search(
            r"Mensalidades e Pacotes Promocionais(.*?)TOTAL\s+R\$",
            bloco,
            re.DOTALL
        )
        if secao:
            for linha_bloco in secao.group(1).split("\n"):
                if "Claro Passaporte" not in linha_bloco:
                    continue
                m = re.search(r"(Claro Passaporte.*?GB).*?([\d]+,\d{2})$", linha_bloco.strip())
                if m:
                    passaporte = m.group(1).strip()
                    valor_passaporte = m.group(2).strip()
                    break

        resultado[linha] = {
            "Pacote": pacote,
            "Passaporte": passaporte,
            "Valor Passaporte": valor_passaporte
        }
    return resultado

def extrair_detalhamento(blocos: dict) -> dict:
    mapa = {}
    for linha, bloco in blocos.items():
        internet = "0"
        m = re.search(
            r"Serviços \(Torpedos.*?^Internet\s+([\d\.,]+)\s+0,00",
            bloco, re.DOTALL | re.MULTILINE
        )
        if m:
            internet = m.group(1)

        minutos = "0"
        m = re.search(r"TOTAL\s+([\d]+min[\d]*s?|[\d]+s)", bloco)
        if m:
            minutos = m.group(1)

        mapa[linha] = {
            "Internet (MB)": internet,
            "Minutos": minutos
        }
    return mapa

def to_float(valor) -> float:
    try:
        return float(str(valor).replace(".", "").replace(",", "."))
    except:
        return 0.0

def extrair_gb_pacote(pacote: str) -> int:
    m = re.search(r"(\d+)\s*GB", str(pacote))
    return int(m.group(1)) if m else 0

def processar_pdf(file):
    texto = ""
    placeholder = st.empty()
    with pdfplumber.open(file) as pdf:
        total_paginas = len(pdf.pages)
        progresso = st.progress(0)
        for i, page in enumerate(pdf.pages):
            placeholder.markdown(f"### 📄 Processando página {i+1} de {total_paginas}")
            t = page.extract_text()
            if t:
                texto += t + "\n"
            progresso.progress((i + 1) / total_paginas)
        placeholder.text("🔍 Extraindo dados...")

    cliente = extrair_cliente(texto)
    vencimento = extrair_vencimento(texto)
    linhas = extrair_linhas(texto)

    blocos = extrair_blocos_por_linha(texto)

    mensalidades = extrair_mensalidades(blocos)
    detalhamento = extrair_detalhamento(blocos)
    pacotes = extrair_pacote_e_passaporte(blocos)

    progresso.empty()
    placeholder.empty()

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
        minutos_str = str(row["Minutos"]).strip().lower()
        sem_minutos = re.fullmatch(r"0[^\d]*|", minutos_str) is not None
        if row["Internet (MB)"] == 0 and sem_minutos:
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

    return df, cliente, vencimento

def gerar_excel(df: pd.DataFrame) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Detalhamento"

    df_reset = df.reset_index(drop=True)

    for r in dataframe_to_rows(df_reset, index=False, header=True):
        ws.append(r)

    borda = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="333333", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        cell.border = borda

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ===== EXECUÇÃO =====
if uploaded_files:
    df_total = pd.DataFrame()
    cliente_nome = "CLIENTE"
    vencimento_fatura = ""

    for file in uploaded_files:
        df, cliente, vencimento = processar_pdf(file)
        df_total = pd.concat([df_total, df], ignore_index=True)
        cliente_nome = cliente
        vencimento_fatura = vencimento

    if not df_total.empty:

        st.markdown('<p class="tt-section-title">📊 Resumo da Fatura</p>', unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4)

        total_linhas = len(df_total)
        em_uso = (df_total["Em Uso"] == "Sim").sum()
        total_gb = df_total["Internet (MB)"].sum() / 1024
        media_gb = total_gb / total_linhas if total_linhas else 0

        col1.metric("Total de Linhas", total_linhas)
        col2.metric("Linhas em Uso", em_uso)
        col3.metric("Consumo Total de Dados", f"{round(total_gb,1)} GB")
        col4.metric("Consumo Médio por Linha", f"{round(media_gb,1)} GB")

        # ===== INSIGHTS =====
        st.markdown('<p class="tt-section-title">🧠 Insights Automáticos</p>', unsafe_allow_html=True)

        upsell = (df_total["Estratégia Comercial"].str.contains("Upsell")).sum()
        ociosidade = (df_total["Em Uso"] == "Não").sum()

        col_i1, col_i2 = st.columns(2)

        col_i1.metric("Upsell", upsell)
        col_i2.metric("Linhas sem uso", ociosidade)

        st.markdown(f"""
        <div style="padding:14px;border-radius:12px;
                    background:rgba(16,185,129,0.08);
                    border:1px solid rgba(16,185,129,0.25);">
            <strong style="color:#10b981;">Resumo executivo:</strong>
            {upsell} oportunidades de upgrade e {ociosidade} linhas com possível economia.
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<p class="tt-section-title">📋 Detalhamento</p>', unsafe_allow_html=True)
        st.dataframe(df_total, use_container_width=True)

        excel = gerar_excel(df_total)

        st.download_button(
            "📥 Baixar Excel",
            data=excel,
            file_name="analise.xlsx"
        )
