import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# ===== CONFIG =====
st.set_page_config(layout="wide")

# ===== HEADER PROFISSIONAL =====
col1, col2 = st.columns([1, 5])

with col1:
    st.image("logo.png", width=120)

with col2:
    st.markdown("## TARGET TELECOM")
    st.markdown("### Análise Inteligente de Faturas Corporativas")

st.markdown("---")

st.markdown("""
📊 **Transforme faturas em estratégia comercial**

Envie suas faturas e identifique oportunidades de:
- Otimização de consumo
- Ajustes comerciais
- Expansão de planos (Upsell inteligente)
""")

st.markdown("### 📎 Envie os PDFs")
uploaded_files = st.file_uploader("", type="pdf", accept_multiple_files=True)

# ===== BASE ORIGINAL (INALTERADA) =====

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

# ===== RESTANTE 100% BASE ORIGINAL =====

def gerar_excel(df):

    wb = Workbook()
    ws = wb.active
    ws.title = "Detalhamento"

    df_reset = df.reset_index(drop=True)

    for r in dataframe_to_rows(df_reset, index=False, header=True):
        ws.append(r)

    borda = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    header_fill = PatternFill(start_color="333333", fill_type="solid")
    zebra = PatternFill(start_color="F2F2F2", fill_type="solid")

    vermelho = PatternFill(start_color="FF4C4C", fill_type="solid")
    verde = PatternFill(start_color="C6EFCE", fill_type="solid")
    amarelo = PatternFill(start_color="FFF3B0", fill_type="solid")
    azul = PatternFill(start_color="BDD7EE", fill_type="solid")
    cinza = PatternFill(start_color="D9D9D9", fill_type="solid")

    headers = [cell.value for cell in ws[1]]

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda

    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):

        for j, cell in enumerate(row):
            coluna = headers[j]

            if coluna == "Perfil":
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif coluna == "Minutos":
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif coluna == "Estratégia Comercial":
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            cell.border = borda

        if i % 2 == 0:
            for cell in row:
                cell.fill = zebra

        perfil = str(row[8].value)
        uso = str(row[9].value)
        estrategia = str(row[10].value)

        if "Alto" in perfil:
            row[8].fill = vermelho
        elif "Médio" in perfil:
            row[8].fill = amarelo

        if uso == "Não":
            row[9].fill = vermelho
        else:
            row[9].fill = verde

        if "Manter" in estrategia:
            row[10].fill = cinza
        elif "Sustentar" in estrategia:
            row[10].fill = amarelo
        elif "Bem dimensionado" in estrategia:
            row[10].fill = verde
        elif "Upsell" in estrategia:
            row[10].fill = azul

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_length + 3

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    resumo = wb.create_sheet(title="Resumo Executivo")

    total_linhas = len(df)
    em_uso = (df["Em Uso"] == "Sim").sum()
    sem_uso = (df["Em Uso"] == "Não").sum()

    total_mb = df["Internet (MB)"].sum()
    total_gb = total_mb / 1024
    media_gb = total_gb / total_linhas if total_linhas > 0 else 0

    def limpar_valor(v):
        v = str(v).replace("R$", "").replace(".", "").replace(",", ".").strip()
        try:
            return float(v)
        except:
            return 0

    total_valor = (
        df["Mensalidade"].apply(limpar_valor).sum() +
        df["Mensalidade Passaporte"].apply(limpar_valor).sum()
    )

    dados_resumo = [
        ["Total de Linhas", total_linhas],
        ["Linhas em Uso", em_uso],
        ["Linhas sem Uso", sem_uso],
        ["Consumo Médio (GB)", round(media_gb, 2)],
        ["Total de GB", round(total_gb, 2)],
        ["Valor Total do Plano (R$)", round(total_valor, 2)],
    ]

    for linha in dados_resumo:
        resumo.append(linha)

    for row in resumo.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="left")

    resumo.column_dimensions["A"].width = 35
    resumo.column_dimensions["B"].width = 20

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer

# ===== EXECUÇÃO =====

if uploaded_files:

    st.write("🚀 Iniciando processamento...")

    df_total = pd.DataFrame()
    cliente_nome = "CLIENTE"

    for file in uploaded_files:
        with st.spinner(f"Processando {file.name}..."):
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
