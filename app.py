import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# ── Google Vision (OCR para PDFs de imagem) ──────────────────────────────────
try:
    from google.cloud import vision
    from google.oauth2 import service_account
    _VISION_DISPONIVEL = True
except ImportError:
    _VISION_DISPONIVEL = False

# FIX 1: st.set_page_config() DEVE ser a primeira chamada Streamlit — sem nada antes
st.set_page_config(
    layout="wide",
    page_title="Target Telecom · Análise de Faturas",
    page_icon="📡"
)

# ===== CSS REDESIGN =====
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

/* ── Base & Background ── */
html, body, [data-testid="stAppViewContainer"] {
    background: #060d1a !important;
}
.main {
    background: #060d1a !important;
}
[data-testid="stAppViewContainer"]::before {
    content: '';
    position: fixed;
    top: -20%;
    left: -10%;
    width: 55%;
    height: 55%;
    background: radial-gradient(ellipse, rgba(16,185,129,0.07) 0%, transparent 70%);
    pointer-events: none;
    z-index: 0;
}
[data-testid="stAppViewContainer"]::after {
    content: '';
    position: fixed;
    bottom: -10%;
    right: -5%;
    width: 40%;
    height: 40%;
    background: radial-gradient(ellipse, rgba(59,130,246,0.06) 0%, transparent 70%);
    pointer-events: none;
    z-index: 0;
}
.block-container {
    padding-top: 2rem !important;
    padding-bottom: 3rem !important;
    max-width: 1400px !important;
}

/* ── Tipografia global ── */
*, h1, h2, h3, h4, p, span, div, label {
    font-family: 'DM Sans', sans-serif !important;
    color: #e2e8f0;
}

/* ── Header personalizado ── */
.tt-header {
    display: flex;
    align-items: center;
    gap: 20px;
    padding: 28px 36px;
    background: linear-gradient(135deg, #0d1f2d 0%, #0a1628 100%);
    border: 1px solid rgba(16,185,129,0.18);
    border-radius: 20px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.tt-header::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, transparent, #10b981, #3b82f6, transparent);
}
.tt-logo-placeholder {
    width: 56px; height: 56px;
    background: linear-gradient(135deg, #10b981, #059669);
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-size: 26px;
    flex-shrink: 0;
    box-shadow: 0 0 24px rgba(16,185,129,0.35);
}
.tt-brand h1 {
    font-family: 'Syne', sans-serif !important;
    font-size: 1.75rem !important;
    font-weight: 800 !important;
    letter-spacing: -0.02em;
    margin: 0 !important; padding: 0 !important;
    background: linear-gradient(90deg, #ffffff, #a7f3d0);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    line-height: 1.1 !important;
}
.tt-brand p {
    font-size: 0.82rem;
    color: #64748b !important;
    margin: 4px 0 0 0;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    font-weight: 500;
}
.tt-badge {
    margin-left: auto;
    padding: 6px 14px;
    background: rgba(16,185,129,0.1);
    border: 1px solid rgba(16,185,129,0.25);
    border-radius: 20px;
    font-size: 0.72rem;
    color: #10b981 !important;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    font-weight: 600;
}

/* ── Divider ── */
hr {
    border: none !important;
    border-top: 1px solid rgba(255,255,255,0.06) !important;
    margin: 1.5rem 0 !important;
}

/* ── Upload Area ── */
[data-testid="stFileUploader"] {
    background: linear-gradient(135deg, #0d1f2d 0%, #091523 100%) !important;
    border: 1.5px dashed rgba(16,185,129,0.35) !important;
    border-radius: 18px !important;
    padding: 8px 16px !important;
    transition: all 0.3s ease !important;
    position: relative;
}
[data-testid="stFileUploader"]:hover {
    border-color: rgba(16,185,129,0.75) !important;
    background: linear-gradient(135deg, #0f2535 0%, #0c1d2e 100%) !important;
    box-shadow: 0 0 30px rgba(16,185,129,0.08) !important;
}
[data-testid="stFileUploaderDropzone"] {
    background: transparent !important;
    border: none !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] p,
[data-testid="stFileUploaderDropzoneInstructions"] span {
    color: #94a3b8 !important;
    font-size: 0.9rem !important;
    font-family: 'DM Sans', sans-serif !important;
}
[data-testid="stFileUploaderDropzone"] svg {
    fill: #10b981 !important;
    opacity: 0.8;
}
[data-testid="stFileUploader"] label {
    font-size: 0.95rem !important;
    color: #cbd5e1 !important;
    font-weight: 500 !important;
}

/* ── Métricas ── */
[data-testid="metric-container"] {
    background: linear-gradient(145deg, #0d1f2d, #0a1929) !important;
    border: 1px solid rgba(255,255,255,0.07) !important;
    border-radius: 16px !important;
    padding: 20px 24px !important;
    position: relative;
    overflow: hidden;
    transition: transform 0.2s ease, border-color 0.2s ease;
}
[data-testid="metric-container"]:hover {
    transform: translateY(-2px);
    border-color: rgba(16,185,129,0.25) !important;
}
[data-testid="metric-container"]::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, #10b981, #3b82f6);
    opacity: 0.6;
}
[data-testid="stMetricLabel"] {
    font-size: 0.72rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.1em !important;
    color: #64748b !important;
    font-weight: 600 !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Syne', sans-serif !important;
    font-size: 2rem !important;
    font-weight: 700 !important;
    color: #f1f5f9 !important;
    line-height: 1.2 !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(255,255,255,0.07) !important;
    border-radius: 16px !important;
    overflow: hidden !important;
    background: #0a1628 !important;
}
[data-testid="stDataFrame"] table {
    background: transparent !important;
}
[data-testid="stDataFrame"] thead th {
    background: #0d1f2d !important;
    color: #94a3b8 !important;
    font-size: 0.72rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.08em !important;
    border-bottom: 1px solid rgba(255,255,255,0.08) !important;
    padding: 12px 16px !important;
}
[data-testid="stDataFrame"] tbody tr:hover td {
    background: rgba(16,185,129,0.05) !important;
}

/* ── Botão de Download ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 12px !important;
    height: 52px !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.03em !important;
    width: 100% !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 4px 20px rgba(16,185,129,0.25) !important;
}
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 28px rgba(16,185,129,0.4) !important;
    background: linear-gradient(135deg, #047857 0%, #059669 100%) !important;
}

/* ── Spinner e Progress ── */
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #10b981, #3b82f6) !important;
    border-radius: 4px !important;
}
[data-testid="stProgress"] {
    background: rgba(255,255,255,0.06) !important;
    border-radius: 4px !important;
}

/* ── Alerts / Errors ── */
[data-testid="stAlert"] {
    border-radius: 12px !important;
    border-left: 3px solid #ef4444 !important;
    background: rgba(239,68,68,0.08) !important;
}

/* ── Sidebar (se abrir) ── */
[data-testid="stSidebar"] {
    background: #060d1a !important;
    border-right: 1px solid rgba(255,255,255,0.06) !important;
}

/* ── Seção de resultados ── */
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
if os.path.exists("logo.png"):
    col1, col2 = st.columns([1, 5])
    with col1:
        st.image("logo.png", width=180)
    with col2:
        st.markdown("""
        <div style="padding: 12px 0;">
            <div style="font-family:'Syne',sans-serif;font-size:1.6rem;font-weight:800;
                        background:linear-gradient(90deg,#fff,#a7f3d0);
                        -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                        background-clip:text;letter-spacing:-0.02em;">TARGET TELECOM</div>
            <div style="font-size:0.78rem;color:#64748b;text-transform:uppercase;
                        letter-spacing:0.1em;font-weight:500;margin-top:4px;">
                Inteligência em Faturas Corporativas
            </div>
        </div>
        """, unsafe_allow_html=True)
else:
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
    label_visibility="visible"
)

# ===== GOOGLE VISION — cliente único (singleton) =====
# FIX 2: instanciado UMA vez aqui, não repetido dentro de cada função
_vision_client = None

def _get_vision_client():
    """Retorna o cliente Vision, criando-o na primeira chamada."""
    global _vision_client
    if _vision_client is not None:
        return _vision_client
    if not _VISION_DISPONIVEL:
        return None
    try:
        creds_dict = st.secrets["GOOGLE_CREDENTIALS"]
        credentials = service_account.Credentials.from_service_account_info(creds_dict)
        _vision_client = vision.ImageAnnotatorClient(credentials=credentials)
        return _vision_client
    except Exception as e:
        st.warning(f"⚠️ Google Vision não configurado: {e}. PDFs de imagem não serão suportados.")
        return None

def extrair_texto_com_ocr(img_bytes: bytes) -> str:
    """
    Envia imagem PNG (bytes) para Google Vision e retorna o texto extraído.
    FIX 6: recebe bytes puros (não BytesIO) — chamador deve usar .getvalue()
    FIX 2: usa cliente singleton, não recria a cada chamada
    """
    client = _get_vision_client()
    if client is None:
        return ""
    try:
        image = vision.Image(content=img_bytes)
        response = client.document_text_detection(image=image)
        if response.error.message:
            return ""
        return response.full_text_annotation.text or ""
    except Exception:
        return ""

# ===== UTILITÁRIOS =====

def normalizar_numero(num_str: str) -> str:
    """Remove parênteses e espaços do número de telefone."""
    return num_str.replace("(", "").replace(")", "").replace(" ", "")

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
    # FIX #5: usa re.DOTALL para capturar mesmo com quebra de linha
    match = re.search(r"Nº da conta:.*?(\d{2}/\d{2}/\d{4})", texto, re.DOTALL)
    if match:
        return match.group(1)
    # fallback: busca a data de vencimento diretamente
    match = re.search(r"Vencimento\s*\n\s*(\d{2}/\d{2}/\d{4})", texto)
    if match:
        return match.group(1)
    return ""

def extrair_linhas(texto: str) -> list:
    # Padrão principal: (11) 98936 0484
    linhas = re.findall(r"\(\d{2}\)\s\d{5}\s\d{4}", texto)
    # Fallback OCR: número pode vir sem parênteses ou com espaçamento diferente
    if not linhas:
        linhas = re.findall(r"\d{2}\s\d{5}\s\d{4}", texto)
    lista = []
    for l in linhas:
        num = normalizar_numero(l)
        if num not in lista:
            lista.append(num)
    return lista

def extrair_blocos_por_linha(texto: str) -> dict:
    """
    Divide o texto em blocos por linha telefônica.
    Suporta PDFs digitais e texto OCR (com variações de espaçamento).
    """
    blocos = re.split(r"DETALHAMENTO DE LIGAÇÕES E SERVIÇOS DO CELULAR", texto)
    resultado = {}
    for bloco in blocos:
        # Padrão principal com parênteses
        num = re.search(r"\(\d{2}\)\s\d{5}\s\d{4}", bloco)
        if num:
            chave = normalizar_numero(num.group(0))
            resultado[chave] = bloco
        else:
            # Fallback: número sem parênteses (variação OCR)
            num = re.search(r"^\s*\d{2}\s\d{5}\s\d{4}", bloco, re.MULTILINE)
            if num:
                chave = normalizar_numero(num.group(0))
                resultado[chave] = bloco
    return resultado

def extrair_mensalidades(blocos: dict) -> dict:
    """
    Captura o TOTAL de cada linha com múltiplos fallbacks para suportar
    variações de formatação em PDFs digitais e PDFs de imagem (OCR).
    """
    mapa = {}
    for linha, bloco in blocos.items():
        total_pdf = 0.0
        total_str = None

        # Estratégia 1: "TOTAL R$59,15" ou "TOTAL R$ 59,15" (com ou sem espaço)
        m = re.search(r"TOTAL\s*R\$\s*([\d\.,]+)", bloco)
        if m:
            total_str = m.group(1)
            try:
                total_pdf = float(total_str.replace(".", "").replace(",", "."))
            except (ValueError, TypeError):
                pass

        # Estratégia 2: "TOTAL RS59,15" — OCR confunde $ com S
        if total_pdf == 0.0:
            m = re.search(r"TOTAL\s+R[S$]\s*([\d\.,]+)", bloco, re.IGNORECASE)
            if m:
                total_str = m.group(1)
                try:
                    total_pdf = float(total_str.replace(".", "").replace(",", "."))
                except (ValueError, TypeError):
                    pass

        # Estratégia 3: somar valores positivos da seção de mensalidades
        # — fallback robusto quando TOTAL não é detectado e também
        #   serve para ignorar descontos negativos (ex: Desconto Dados GPRS)
        secao = re.search(r"Mensalidades e Pacotes Promocionais(.*?)TOTAL", bloco, re.DOTALL)
        soma_positivos = 0.0
        if secao:
            for lb in secao.group(1).split("\n"):
                mv = re.search(r"(-?[\d]+,\d{2})$", lb.strip())
                if mv:
                    try:
                        val = float(mv.group(1).replace(".", "").replace(",", "."))
                        if val > 0:
                            soma_positivos += val
                    except (ValueError, TypeError):
                        pass

        # Se descontos reduziram o total, ou se total não foi detectado, usa soma
        if soma_positivos > total_pdf + 0.01:
            mapa[linha] = f"{soma_positivos:.2f}".replace(".", ",")
        elif total_str:
            mapa[linha] = total_str

    return mapa

def extrair_pacote_e_passaporte(blocos: dict) -> dict:
    resultado = {}
    for linha, bloco in blocos.items():
        pacote = "-"
        passaporte = "-"
        valor_passaporte = "0"

        # Captura todos os planos possíveis: Claro Pós/Life Ilimitado/Controle e Plano Wi-Fi
        m = re.search(
            r"((?:Claro\s+(?:Pós|Life Ilimitado|Controle)|Plano de Internet Wi-Fi)\s+\d+\s*GB)",
            bloco
        )
        if m:
            pacote = m.group(1).strip()

        # Busca seção de mensalidades e pacotes
        secao = re.search(
            r"Mensalidades e Pacotes Promocionais(.*?)TOTAL\s+R\$",
            bloco,
            re.DOTALL
        )
        if secao:
            trecho = secao.group(1)
            for linha_bloco in trecho.split("\n"):
                linha_limpa = linha_bloco.strip()
                if "Claro Passaporte" not in linha_limpa:
                    continue
                # Captura nome do passaporte e valor no final da linha
                m = re.search(r"(Claro Passaporte.*?GB).*?([\d]+,\d{2})$", linha_limpa)
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

def extrair_detalhamento(blocos: dict, texto: str) -> dict:
    """
    Nova versão robusta para PDFs imagem da Claro
    """

    resultado = {}

    # 🔹 1. INTERNET GLOBAL (pega todos os subtotais)
    internet_vals = re.findall(r"Subtotal\s+([\d\.]+,\d+)", texto)

    # 🔹 2. MINUTOS GLOBAL
    minutos_vals = re.findall(r"TOTAL\s+(\d+min\d+s)", texto)

    # 🔹 3. fallback minutos simples
    if not minutos_vals:
        minutos_vals = re.findall(r"(\d+min\d+s)", texto)

    linhas = list(blocos.keys())

    for i, linha in enumerate(linhas):

        internet = "0"
        minutos = "0"

        # INTERNET
        if i < len(internet_vals):
            internet = internet_vals[i]

        # MINUTOS
        if i < len(minutos_vals):
            minutos = minutos_vals[i]

        resultado[linha] = {
            "Internet (MB)": internet,
            "Minutos": minutos
        }

    return resultado

def to_float(valor) -> float:
    # FIX #2: except específico
    try:
        return float(str(valor).replace(".", "").replace(",", "."))
    except (ValueError, TypeError):
        return 0.0

def extrair_gb_pacote(pacote: str) -> int:
    m = re.search(r"(\d+)\s*GB", str(pacote))
    return int(m.group(1)) if m else 0

def processar_pdf(file):
    """
    Processa um PDF da Claro extraindo texto por duas estratégias:
    1. pdfplumber (PDFs com texto digital) — rápido e preciso
    2. Google Vision OCR (PDFs de imagem escaneados) — fallback por página
    FIX 4/8: barra de progresso avança em TODAS as páginas, não só nas de OCR
    FIX 7: texto concatenado por página mantém contexto correto
    """
    texto = ""
    placeholder = st.empty()
    usou_ocr = False

    with pdfplumber.open(file) as pdf:
        total_paginas = len(pdf.pages)
        progresso = st.progress(0)

        for i, page in enumerate(pdf.pages):
            placeholder.text(f"📄 Processando página {i+1} de {total_paginas}...")

            t = page.extract_text()

            if t and t.strip():
                # PDF com texto digital — extração direta
                texto += t + "\n"
            else:
                # PDF de imagem — aplica OCR via Google Vision
                if not usou_ocr:
                    usou_ocr = True
                    placeholder.text(f"🔍 PDF de imagem detectado — aplicando OCR na página {i+1}...")

                img_buf = io.BytesIO()
                page.to_image(resolution=300).original.save(img_buf, format="PNG")
                texto_ocr = extrair_texto_com_ocr(img_buf.getvalue())
                texto += texto_ocr + "\n"

            # FIX 4: progresso avança para TODA página, não só as de OCR
            progresso.progress((i + 1) / total_paginas)

        placeholder.text("🔎 Extraindo dados...")
        st.text_area("DEBUG OCR TEXTO", texto[:5000])

    if usou_ocr and not texto.strip():
        progresso.empty()
        placeholder.empty()
        raise ValueError(
            "Não foi possível extrair texto do PDF. "
            "Verifique se o Google Vision está configurado nos secrets."
        )

    cliente = extrair_cliente(texto)
    vencimento = extrair_vencimento(texto)
    linhas = extrair_linhas(texto)
    blocos = extrair_blocos_por_linha(texto)

    mensalidades = extrair_mensalidades(blocos)
    detalhamento = extrair_detalhamento(blocos, texto)
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

    # FIX #12: reset_index com ignore_index
    df_reset = df.reset_index(drop=True)

    for r in dataframe_to_rows(df_reset, index=False, header=True):
        ws.append(r)

    borda = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    header_fill = PatternFill(start_color="333333", fill_type="solid")
    zebra       = PatternFill(start_color="F2F2F2", fill_type="solid")
    vermelho    = PatternFill(start_color="FF4C4C", fill_type="solid")
    verde       = PatternFill(start_color="C6EFCE", fill_type="solid")
    amarelo     = PatternFill(start_color="FFF3B0", fill_type="solid")
    azul        = PatternFill(start_color="BDD7EE", fill_type="solid")
    cinza       = PatternFill(start_color="D9D9D9", fill_type="solid")

    # FIX #1: índices por nome de coluna, não posição hardcoded
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

        # FIX #1: usa índice por nome
        perfil     = str(row[col_idx["Perfil"]].value) if "Perfil" in col_idx else ""
        uso        = str(row[col_idx["Em Uso"]].value) if "Em Uso" in col_idx else ""
        estrategia = str(row[col_idx["Estratégia Comercial"]].value) if "Estratégia Comercial" in col_idx else ""

        if "Alto" in perfil:
            row[col_idx["Perfil"]].fill = vermelho
        elif "Médio" in perfil:
            row[col_idx["Perfil"]].fill = amarelo

        if "Em Uso" in col_idx:
            if uso == "Não":
                row[col_idx["Em Uso"]].fill = vermelho
            else:
                row[col_idx["Em Uso"]].fill = verde

        if "Estratégia Comercial" in col_idx:
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
        # FIX #11: tratamento de erro no processamento
        try:
            with st.spinner(f"Processando {file.name}..."):
                df, cliente, vencimento = processar_pdf(file)
                # FIX #12: ignore_index=True no concat
                df_total = pd.concat([df_total, df], ignore_index=True)
                cliente_nome = cliente
                vencimento_fatura = vencimento
        except Exception as e:
            st.error(f"❌ Erro ao processar **{file.name}**: {e}")
            continue

        progress.progress((i + 1) / total_files)

    if not df_total.empty:
        st.markdown('<div class="tt-divider"></div>', unsafe_allow_html=True)
        st.markdown('<p class="tt-section-title">📊 Resumo da Fatura</p>', unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4)
        total_linhas = len(df_total)
        em_uso       = (df_total["Em Uso"] == "Sim").sum()
        inativas     = total_linhas - em_uso
        total_gb     = df_total["Internet (MB)"].sum() / 1024
        media_gb     = total_gb / total_linhas if total_linhas else 0

        col1.metric("Total de Linhas", total_linhas)
        col2.metric("Linhas Ativas", em_uso, delta=f"{inativas} inativas" if inativas else None,
                    delta_color="inverse")
        col3.metric("Consumo Total", f"{round(total_gb, 1)} GB")
        col4.metric("Média por Linha", f"{round(media_gb, 1)} GB")

        st.markdown('<div class="tt-divider"></div>', unsafe_allow_html=True)
        st.markdown('<p class="tt-section-title">📋 Detalhamento por Linha</p>', unsafe_allow_html=True)

        st.dataframe(
            df_total,
            use_container_width=True,
            height=min(600, 60 + len(df_total) * 38),
            hide_index=True,
        )

        st.markdown('<div class="tt-divider"></div>', unsafe_allow_html=True)

        col_info, col_btn = st.columns([3, 1])
        with col_info:
            st.markdown(f"""
            <div style="padding:14px 18px;background:rgba(16,185,129,0.06);
                        border:1px solid rgba(16,185,129,0.15);border-radius:12px;">
                <div style="font-size:0.72rem;color:#64748b;text-transform:uppercase;
                            letter-spacing:0.08em;font-weight:600;margin-bottom:4px;">
                    Relatório pronto
                </div>
                <div style="font-size:0.9rem;color:#cbd5e1;">
                    <strong style="color:#10b981;">{total_linhas}</strong> linhas ·
                    cliente <strong style="color:#e2e8f0;">{cliente_nome}</strong> ·
                    vencimento <strong style="color:#e2e8f0;">{vencimento_fatura or '—'}</strong>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_btn:
            excel = gerar_excel(df_total)
            nome_arquivo = (
                f"Analise_Target_{cliente_nome}_{vencimento_fatura}.xlsx"
                if vencimento_fatura
                else f"Analise_Target_{cliente_nome}.xlsx"
            )
            st.download_button(
                "📥  Baixar Relatório Excel",
                data=excel,
                file_name=nome_arquivo,
                use_container_width=True,
            )
