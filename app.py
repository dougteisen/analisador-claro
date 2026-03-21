import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

st.set_page_config(layout="wide")

# ===== LOGO =====
st.image("logo.png", width=180)

# ===== TÍTULO =====
st.title("📊 Target | Análise de Faturas Corporativas")

# ===== TEXTO COMERCIAL =====
st.markdown("""
### 📈 Otimize seus custos e identifique oportunidades

Envie suas faturas da operadora e receba uma análise completa com:

- 📊 Consumo por linha
- 💰 Visão financeira detalhada
- 🚀 Recomendações comerciais inteligentes
""")

# ===== UPLOAD MELHORADO =====
st.markdown("### 📎 Envie suas faturas (PDF)")

uploaded_files = st.file_uploader(
    "Arraste ou selecione os arquivos",
    type="pdf",
    accept_multiple_files=True
)

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
        if "Baixo" in row["Perfil"]:
            return "🟡 Sustentar plano"
        if "Médio" in row["Perfil"]:
            return "🟢 Bem dimensionado"
        if "Alto" in row["Perfil"]:
            return "🔵 Upsell → Aumento recomendado"
        return ""

    df["Estratégia Comercial"] = df.apply(estrategia, axis=1)

    return df, cliente

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

if uploaded_files:

    st.write("🚀 Iniciando processamento...")

    df_total = pd.DataFrame()
    cliente_nome = "CLIENTE"

    for file in uploaded_files:
        try:
            df, cliente = processar_pdf(file)
            df_total = pd.concat([df_total, df])
            cliente_nome = cliente
        except Exception as e:
            st.error(f"Erro ao processar arquivo: {e}")

    st.write("✅ Processamento finalizado")

    if not df_total.empty:
        st.success(f"{len(df_total)} linhas processadas")
        st.dataframe(df_total)

        excel = gerar_excel(df_total)

        nome_arquivo = f"Analise_Target_{cliente_nome}.xlsx"

        st.download_button(
            "📥 Baixar Excel",
            data=excel,
            file_name=nome_arquivo
        )
    else:
        st.warning("⚠️ Nenhum dado foi extraído")
