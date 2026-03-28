import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

st.set_page_config(layout="wide")

# ===== CSS =====
st.markdown("""
<style>
.main {
    background: linear-gradient(180deg, #020617 0%, #020617 60%, #0f172a 100%);
}

/* Títulos */
h1 { font-size: 2.2rem; font-weight: 700; }
h2 { font-size: 1.6rem; }
h3 { font-size: 1.2rem; }

/* Container geral */
.block-container {
    padding-top: 1.2rem;
    max-width: 1200px;
}

/* Cards de métricas */
.stMetric {
    background: linear-gradient(145deg, #0f172a, #111827);
    padding: 18px;
    border-radius: 16px;
    border: 1px solid #1f2937;
    box-shadow: 0px 6px 25px rgba(0,0,0,0.5);
}

/* Upload */
[data-testid="stFileUploader"] {
    border: 2px dashed #334155;
    border-radius: 16px;
    background: #020617;
    padding: 12px;
    transition: 0.3s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #22c55e;
    box-shadow: 0 0 20px rgba(34,197,94,0.2);
}

/* Botão download */
.stDownloadButton>button {
    background: linear-gradient(90deg, #16a34a, #22c55e);
    color: white;
    border-radius: 12px;
    height: 55px;
    font-weight: bold;
    transition: 0.3s;
    border: none;
}
.stDownloadButton>button:hover {
    transform: scale(1.04);
    box-shadow: 0 0 15px rgba(34,197,94,0.4);
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 12px;
    border: 1px solid #1f2937;
}
</style>
""", unsafe_allow_html=True)

# ===== HEADER =====
col1, col2 = st.columns([1, 5])

with col1:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=120)

with col2:
    st.markdown("""
    <div style="padding-top:10px">
        <h1 style="margin-bottom:0;">TARGET TELECOM</h1>
        <p style="color:#94a3b8; margin-top:4px;">
            Inteligência em Faturas Corporativas • Análise Automatizada
        </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

uploaded_files = st.file_uploader(
    "📎 Envie suas faturas em PDF (uma ou várias)",
    type="pdf",
    accept_multiple_files=True
)

st.caption("💡 O sistema identifica automaticamente linhas, consumo, planos e oportunidades comerciais.")

# ===== UTILITÁRIOS =====

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

    borda = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    header_fill = PatternFill(start_color="333333", fill_type="solid")
    zebra = PatternFill(start_color="F2F2F2", fill_type="solid")
    vermelho = PatternFill(start_color="FF4C4C", fill_type="solid")
    verde = PatternFill(start_color="C6EFCE", fill_type="solid")
    amarelo = PatternFill(start_color="FFF3B0", fill_type="solid")
    azul = PatternFill(start_color="BDD7EE", fill_type="solid")
    cinza = PatternFill(start_color="D9D9D9", fill_type="solid")

    headers = [cell.value for cell in ws[1]]
    col_idx = {v: i for i, v in enumerate(headers)}

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda

    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for j, cell in enumerate(row):
            coluna = headers[j]
            if coluna in ("Perfil", "Estratégia Comercial"):
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif coluna == "Minutos":
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = borda

        if i % 2 == 0:
            for cell in row:
                cell.fill = zebra

        perfil = str(row[col_idx["Perfil"]].value)
        uso = str(row[col_idx["Em Uso"]].value)
        estrategia = str(row[col_idx["Estratégia Comercial"]].value)

        if "Alto" in perfil:
            row[col_idx["Perfil"]].fill = vermelho
        elif "Médio" in perfil:
            row[col_idx["Perfil"]].fill = amarelo

        if uso == "Não":
            row[col_idx["Em Uso"]].fill = vermelho
        else:
            row[col_idx["Em Uso"]].fill = verde

        if "Manter" in estrategia:
            row[col_idx["Estratégia Comercial"]].fill = cinza
        elif "Sustentar" in estrategia:
            row[col_idx["Estratégia Comercial"]].fill = amarelo
        elif "Bem dimensionado" in estrategia:
            row[col_idx["Estratégia Comercial"]].fill = verde
        elif "Upsell" in estrategia:
            row[col_idx["Estratégia Comercial"]].fill = azul

    for col in ws.columns:
        max_length = max((len(str(cell.value)) for cell in col if cell.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ===== EXECUÇÃO =====
if uploaded_files:
    df_total = pd.DataFrame()
    cliente_nome = "CLIENTE"
    vencimento_fatura = ""

    progress = st.progress(0)
    total_files = len(uploaded_files)

    for i, file in enumerate(uploaded_files):
        try:
            with st.spinner(f"Processando {file.name}..."):
                df, cliente, vencimento = processar_pdf(file)
                df_total = pd.concat([df_total, df], ignore_index=True)
                cliente_nome = cliente
                vencimento_fatura = vencimento
        except Exception as e:
            st.error(f"❌ Erro ao processar **{file.name}**: {e}")
            continue

        progress.progress((i + 1) / total_files)

    if not df_total.empty:
        st.markdown("## 📊 Resumo da Fatura")

        col1, col2, col3, col4 = st.columns(4)
        total_linhas = len(df_total)
        em_uso = (df_total["Em Uso"] == "Sim").sum()
        total_gb = df_total["Internet (MB)"].sum() / 1024
        media_gb = total_gb / total_linhas if total_linhas else 0

        col1.metric("📱 Linhas", total_linhas)
        col2.metric("📡 Em uso", em_uso)
        col3.metric("🌐 Total GB", round(total_gb, 1))
        col4.metric("📊 Média GB", round(media_gb, 1))

        st.markdown("## 📋 Detalhamento das Linhas")
        st.dataframe(df_total)

        st.markdown("## 📥 Exportação")

        excel = gerar_excel(df_total)

        nome_arquivo = (
            f"Analise_Target_{cliente_nome}_{vencimento_fatura}.xlsx"
            if vencimento_fatura
            else f"Analise_Target_{cliente_nome}.xlsx"
        )

        st.download_button(
            "📥 Baixar Relatório Excel",
            data=excel,
            file_name=nome_arquivo
        )
