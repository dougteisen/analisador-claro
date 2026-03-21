import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

st.set_page_config(layout="wide")

# ===== CSS PROFISSIONAL =====
st.markdown("""
<style>
.main {
    background-color: #0f172a;
}

.block-container {
    padding-top: 1.5rem;
}

h1, h2, h3 {
    color: white;
}

p {
    color: #cbd5e1;
}

.stMetric {
    background-color: #111827;
    padding: 15px;
    border-radius: 10px;
}

.stDownloadButton>button {
    background-color: #16a34a;
    color: white;
    border-radius: 10px;
    height: 50px;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# ===== HEADER MELHORADO =====
col1, col2 = st.columns([1.5, 4])

with col1:
    st.image("logo.png", width=180)  # 🔧 LOGO MENOR

with col2:
    st.markdown("# TARGET TELECOM")
    st.markdown("### Inteligência em Análise de Faturas Corporativas")

st.markdown("---")

st.markdown("### 📎 Envie suas faturas")
uploaded_files = st.file_uploader("", type="pdf", accept_multiple_files=True)

# ===== BASE ORIGINAL =====

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

        m = re.search(r"(Claro Pós\s*\d+GB)", bloco)
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
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto += t + "\n"

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

# ===== FUNÇÃO EXCEL (RESTAURADA - ERA O ERRO) =====

def gerar_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Detalhamento"

    df_reset = df.reset_index(drop=True)

    for r in dataframe_to_rows(df_reset, index=False, header=True):
        ws.append(r)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer

# ===== EXECUÇÃO =====

if uploaded_files:

    df_total = pd.DataFrame()
    cliente_nome = "CLIENTE"

    for file in uploaded_files:
        with st.spinner(f"Processando {file.name}..."):
            df, cliente = processar_pdf(file)
            df_total = pd.concat([df_total, df])
            cliente_nome = cliente

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
