import streamlit as st
from datetime import date
from io import BytesIO
from docx import Document
import sqlite3
import json
import os
import requests

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Laudo de Morte Violenta | ICRIM/NPT",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

AUTORIDADES_DEFAULT = ["Delegado(a) Plantonista", "Policia Civil"]
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
EXTENSO = {
    1: "um", 2: "dois", 3: "três", 4: "quatro", 5: "cinco", 6: "seis",
    7: "sete", 8: "oito", 9: "nove", 10: "dez", 11: "onze", 12: "doze",
    13: "treze", 14: "catorze", 15: "quinze", 16: "dezesseis", 17: "dezessete",
    18: "dezoito", 19: "dezenove", 20: "vinte", 21: "vinte e um", 22: "vinte e dois",
    23: "vinte e três", 24: "vinte e quatro", 25: "vinte e cinco", 26: "vinte e seis",
    27: "vinte e sete", 28: "vinte e oito", 29: "vinte e nove", 30: "trinta", 31: "trinta e um",
    2025: "dois mil e vinte e cinco", 2026: "dois mil e vinte e seis",
    2027: "dois mil e vinte e sete", 2028: "dois mil e vinte e oito",
    2029: "dois mil e vinte e nove", 2030: "dois mil e trinta",
}


def num_extenso(n): return EXTENSO.get(n, str(n))


def data_extenso(d):
    return (f"Aos {d.day:02d} ({num_extenso(d.day)}) dia(s) do mês de "
            f"{MESES[d.month-1]} do ano {d.year} ({num_extenso(d.year)})")


def data_simples(d): return d.strftime("%d/%m/%Y")
def v(val): return str(val).strip() if str(val).strip() else "________"


# ═══════════════════════════════════════════════════════════
# MÓDULO DE INTEGRAÇÃO GOOGLE GEMINI AI
# ═══════════════════════════════════════════════════════════
# GEMINI VISION & IA HELPERS
# ═══════════════════════════════════════════════════════════
def get_gemini_api_key():
    """Recupera a chave de API do Gemini de CHAVE.txt, secrets, ambiente ou session_state."""
    for fname in ["CHAVE.txt", "chave.txt", "CHAVE.TXT"]:
        fpath = os.path.join(os.path.dirname(__file__), fname)
        if os.path.exists(fpath):
            try:
                with open(fpath, "r", encoding="utf-8", errors="ignore") as fp:
                    k = fp.read().strip()
                    if k: return k
            except Exception: pass
    try:
        if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
            return st.secrets["GEMINI_API_KEY"]
    except Exception: pass
    env_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if env_key: return env_key
    if st.session_state.get("user_gemini_key"):
        return st.session_state.get("user_gemini_key")
    if st.session_state.get("gemini_api_key_input"):
        return st.session_state.get("gemini_api_key_input")
    return ""


def convert_bytes_to_pil_images(file_bytes, file_name, file_type=""):
    """
    Converte bytes de arquivo (PDF ou Imagem) em lista de objetos PIL Image.
    """
    from io import BytesIO
    from PIL import Image

    file_name_lower = file_name.lower()
    is_pdf = file_type == "application/pdf" or file_name_lower.endswith(".pdf")
    images = []

    if is_pdf:
        # 1. PyMuPDF (fitz) - alta fidelidade e velocidade
        try:
            import fitz
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page in doc:
                pix = page.get_pixmap(dpi=150)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                images.append(img)
        except Exception:
            pass

        # 2. Fallback pdfplumber
        if not images:
            try:
                import pdfplumber
                with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                    for page in pdf.pages:
                        img = page.to_image(resolution=150).original
                        images.append(img.convert("RGB"))
            except Exception:
                pass
    else:
        try:
            img = Image.open(BytesIO(file_bytes))
            images.append(img.convert("RGB"))
        except Exception:
            pass

    return images




def carregar_dados_teste_exemplo():
    """Carrega dados de teste completos e 3 fotografias demonstrativas no session_state."""
    import datetime
    st.session_state["num_laudo"] = "2026-ICRIM-0987"
    st.session_state["ocorrencia"] = "BO-12345/2026"
    st.session_state["requisicao"] = "REQ-5544/2026"
    st.session_state["perito"] = "Dr. Carlos Eduardo Silva - Perito Criminal"
    st.session_state["destino"] = "1ª Delegacia de Polícia Civil de Imperatriz"
    st.session_state["autoridade_sel"] = "Dr. João Mendes - Delegado de Polícia"
    st.session_state["data_pericia_input"] = datetime.date.today()
    st.session_state["data_atendimento_input"] = datetime.date.today()
    st.session_state["horario"] = datetime.time(14, 30)
    st.session_state["horario_atendimento"] = datetime.time(14, 45)
    
    st.session_state["endereco"] = "Rua das Palmeiras, nº 450, Bairro Maranhão Novo"
    st.session_state["municipio"] = "Imperatriz - MA"
    st.session_state["latitude"] = "-5.5264"
    st.session_state["longitude"] = "-47.4721"
    st.session_state["ponto_referencia"] = "Próximo à Praça da Cultura"
    st.session_state["area"] = "Via Pública Comercial / Asfalto"
    st.session_state["pavimento"] = "Asfalto"
    st.session_state["delimitacoes"] = "Local aberto, delimitado ao norte pela via pública e ao sul por imóvel comercial."
    st.session_state["iso_estado"] = "1. Local Preservado e Isolado"
    st.session_state["iso_meio"] = "fita zebrada e cordão de isolamento da Polícia Militar"
    st.session_state["clima"] = "Seco / Ensolarado"
    st.session_state["visibilidade"] = "Boa (Luz Natural Solar)"
    st.session_state["iluminacao"] = "Natural"
    st.session_state["equipe_pm"] = "VTR 14-020 (Sgt. Oliveira e Sd. Santos)"
    st.session_state["equipe_pc"] = "Equipe de Homicídios (Inv. Lima)"
    st.session_state["autoridade_local"] = "Dr. João Mendes - Delegado de Polícia"

    st.session_state["vitimas"] = [{
        "nome": "João da Silva Sauro",
        "sexo": "Masculino",
        "idade": "35 anos",
        "naturalidade": "Imperatriz - MA",
        "rg": "1234567 SSP/MA",
        "cpf": "000.111.222-33",
        "filiacao": "Maria da Silva Sauro",
        "posicao": "Decúbito dorsal",
        "cabeca": "Voltada para a direita",
        "membros": "Superiores estendidos ao longo do tronco; inferiores paralelos",
        "vestes": "Camisa polo branca e calça jeans azul",
        "fenomenos": "Rigidez cadavérica em início de fixação; livores de hipóstase dorsais",
        "lesoes": ["Ferimento perfurocontundente com orifício de entrada em região parietal direita.",
                   "Orifício de saída em região temporal esquerda associado a fratura de calota craniana."]
    }]

    st.session_state["vestigios"] = [
        {
            "tipo": "Vestígio Balístico",
            "localizacao": "Solo asfáltico, a 1,20m do cadáver",
            "orientacao": "Norte",
            "coordenadas": "X: 2.5m, Y: 1.8m",
            "detalhes": "Estojo percutido e deflagrado de munição calibre 9mm, marca CBC."
        },
        {
            "tipo": "Mancha de Sangue",
            "localizacao": "Solo abaixo da região cefálica",
            "orientacao": "Centro",
            "coordenadas": "X: 0m, Y: 0m",
            "detalhes": "Poça de sangue com padrão de poça e escorrimento por gravidade."
        }
    ]

    st.session_state["inst_acao"] = "1. Perfurocontundente (Ex: Projétil de Arma de Fogo - PAF)"
    st.session_state["inst_agente"] = "projéteis de arma de fogo (PAF)"
    st.session_state["inst_extra"] = "com a recuperação de um projétil de arma de fogo durante o exame perinecroscópico e confirmação de disparo a curta distância"

    st.session_state["consideracoes_extras"] = [
        {
            "titulo": "Do Exame das Vestes da Vítima",
            "texto": "As vestes da vítima apresentavam perfurações com orlas de enxugo e esfumaçamento compatíveis com a entrada de projétil disparado por arma de fogo."
        },
        {
            "titulo": "Das Marcas de Frenagem Pneumática",
            "texto": "Foi constatada marca de frenagem pneumática impressa na pista de rolamento com extensão de 4,20 metros, sugerindo tentativa de desaceleração veicular no momento dos fatos."
        }
    ]

    st.session_state["quesitos_list"] = [
        {"pergunta": "Qual a causa da morte?", "resposta": "Traumatismo cranioencefálico decorrente de ferimento por projétil de arma de fogo."},
        {"pergunta": "Qual o instrumento ou meio utilizado?", "resposta": "Instrumento em ação perfurocontundente (projétil de arma de fogo)."}
    ]

    st.session_state["dinamica_fatos"] = (
        "Diante dos elementos materiais coligidos no local dos fatos e exames perinecroscópicos realizados, "
        "infere-se que a vítima transitava pela via pública quando foi surpreendida pelo agressor. "
        "Foram efetuados disparos de arma de fogo de curto alcance, vindo a atingir a região cefálica da vítima, "
        "ocasionando sua queda imediata em decúbito dorsal e o consequente óbito no local por traumatismo cranioencefálico."
    )

    st.session_state["fotos"] = [
        {
            "b64": "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wIf/AOS6P+Eo/Zg/6Jz4o/8AAh//AJLo9s/5X/XzD2S/mR4PRX0tpWhfA3xt8PfG+q+EfBeqWF5oOky3CyXt1KMSGKVo2UCdwcGIkhhjp15r5pqoVFO+lrEzhy213CvePGv/ACZZ4H/7Dsv/AKHe14PXvHjX/kyzwP8A9h2X/wBDvamtvH1/zKpbS9DweiiitjIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAPeP2af8Akk/xi/7AQ/8ARF3Xg9e8fs0/8kn+MX/YCH/oi7rwesafxz+X5Gs/gj8/zCvePGv/ACZZ4H/7Dsv/AKHe14PXvHjX/kyzwP8A9h2X/wBDvaK28fX/ADCltL0PB6KKK2MgooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA94/Zp/5JP8AGL/sBD/0Rd14PXvH7NP/ACSf4xf9gIf+iLuvB6xp/HP5fkaz+CPz/MK948a/8mWeB/8AsOy/+h3teD17x41/5Ms8D/8AYdl/9DvaK28fX/MKW0vQ8HooorYyCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD3j9mn/AJJP8Yv+wEP/AERd14PXvH7NP/JJ/jF/2Ah/6Iu68HrGn8c/l+RrP4I/P8wr3jxr/wAmWeB/+w7L/wCh3teD17x41/5Ms8D/APYdl/8AQ72itvH1/wAwpbS9DweiiitjIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAPeP2af+ST/ABi/7AQ/9EXdeD17x+zT/wAkn+MX/YCH/oi7rwesafxz+X5Gs/gj8/zCvePGv/Jlngf/ALDsv/od7Xg9e8eNf+TLPA//AGHZf/Q72itvH1/zCltL0PB6KKK2MgooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA94/Zp/wCST/GL/sBD/wBEXdeD17x+zT/ySf4xf9gIf+iLuvB6xp/HP5fkaz+CPz/MK948a/8AJlngf/sOy/8Aod7Xg9e8eNf+TLPA/wD2HZf/AEO9orbx9f8AMKW0vQ8HooorYyCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD3j9mn/kk/wAYv+wEP/RF3Xg9e8fs0/8AJJ/jF/2Ah/6Iu68HrGn8c/l+RrP4I/P8wr3jxr/yZZ4H/wCw7L/6He14PXvHjX/kyzwP/wBh2X/0O9orbx9f8wpbS9DweiiitjIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAPeP2af8Akk/xi/7AQ/8ARF3Xg9e8fs0/8kn+MX/YCH/oi7rwesafxz+X5Gs/gj8/zCvePGv/ACZZ4H/7Dsv/AKHe14PXvHjX/kyzwP8A9h2X/wBDvaK28fX/ADCltL0PB6KKs6TYXeq6raaXYRedeXk6W8Ee4LvkdgqjJIAySOSQK2MitRW7H4T1mS4aGL+zZdkXmySx6pbNFEuQuXkEmxMkgAMQSTxmpbTwV4hub+ewjgsVuIIjMyS6lbR7ohGZDIhaQB02AtuXIwM5qeZdx8r7HO0V0DeDPEK25nFtauPLeVEjv4HkljTO540DlpEG1vmUEfKeeDVKy0DV7zQr7Xbeyd9OsNv2icsqquWRMDJyxzImQucbgTgGnzLuHKzMoqzbWF3c2V3ewxb4LQIZ2DDKBm2qcZyRnAyOmRnqK028JeIF+w7tPwL/AHfZyZo8ELDHOxb5vkAimjclsAA89DgugszDoro7bwR4hubt7WKLT/NWA3A36paorxKjO0iMZAHUKjEspIG055BqjceHtWg0q51VreJ7G2uRayzxXEcieYRnAKsdwxj5lyBkAnJGTmXcOVmV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",
            "legenda": "Fotografia 01 — Visão geral do local de crime evidenciando o cadáver no solo asfáltico.",
            "incluir": True
        },
        {
            "b64": "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",
            "legenda": "Fotografia 02 — Posicionamento inicial do cadáver em decúbito dorsal no momento do exame pericial.",
            "incluir": True
        },
        {
            "b64": "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EKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAPb/wBnX/kl3xb/AOwKP/RF1XiFe3/s6/8AJLvi3/2BR/6Iuq8Qq5fDE2qfw4fP8wr2/wAYf8meeDP+w1J/6HeV4hXt/jD/AJM88Gf9hqT/ANDvKIdfQKO0vT/I8QoooqDEKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAPb/ANnX/kl3xb/7Ao/9EXVeIV7f+zr/AMku+Lf/AGBR/wCiLqvEKuXwxNqn8OHz/MK9v8Yf8meeDP8AsNSf+h3leIV7f4w/5M88Gf8AYak/9DvKIdfQKO0vT/I8QoooqDEKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAPb/2df8Akl3xb/7Ao/8ARF1XiFe3/s6/8ku+Lf8A2BR/6Iuq8Qq5fDE2qfw4fP8AMK9v8Yf8meeDP+w1J/6HeV4hXt/jD/kzzwZ/2GpP/Q7yiHX0CjtL0/yPEKKKKgxCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD2/9nX/kl3xb/wCwKP8A0RdV4hXt/wCzr/yS74t/9gUf+iLqvEKuXwxNqn8OHz/MK9v8Yf8AJnngz/sNSf8Aod5XiFe3+MP+TPPBn/Yak/8AQ7yiHX0CjtL0/wAjxCiiioMQooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA9v/AGdf+SXfFv8A7Ao/9EXVeIV7f+zr/wAku+Lf/YFH/oi6rxCrl8MTap/Dh8/zCvb/ABh/yZ54M/7DUn/od5XiFe3+MP8AkzzwZ/2GpP8A0O8oh19Ao7S9P8jxCiiioMQooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA9v8A2df+SXfFv/sCj/0RdV4hXt/7Ov8AyS74t/8AYFH/AKIuq8Qq5fDE2qfw4fP8wr1PwJ8dPFvg3wpZ+G9M07Q5rSz3+W9xDK0h3yM5yVkA6se3TFeWUVMZOOxnCcoO8We3/wDDTfjz/oEeGv8AwGn/APj1H/DTfjz/AKBHhr/wGn/+PV4hRVe1n3NPrNX+Y9v/AOGm/Hn/AECPDX/gNP8A/HqP+Gm/Hn/QI8Nf+A0//wAerxCij2s+4fWav8x7f/w0348/6BHhr/wGn/8Aj1H/AA0348/6BHhr/wABp/8A49XiFFHtZ9w+s1f5j2//AIab8ef9Ajw1/wCA0/8A8eo/4ab8ef8AQI8Nf+A0/wD8erxCij2s+4fWav8AMe3/APDTfjz/AKBHhr/wGn/+PUf8NN+PP+gR4a/8Bp//AI9XiFFHtZ9w+s1f5j2//hpvx5/0CPDX/gNP/wDHqP8Ahpvx5/0CPDX/AIDT/wDx6vEKKPaz7h9Zq/zHt/8Aw0348/6BHhr/AMBp/wD49R/w0348/wCgR4a/8Bp//j1eIUUe1n3D6zV/mPb/APhpvx5/0CPDX/gNP/8AHqP+Gm/Hn/QI8Nf+A0//AMerxCij2s+4fWav8x7f/wANN+PP+gR4a/8AAaf/AOPUf8NN+PP+gR4a/wDAaf8A+PV4hRR7WfcPrNX+Y9v/AOGm/Hn/AECPDX/gNP8A/HqP+Gm/Hn/QI8Nf+A0//wAerxCij2s+4fWav8x7f/w0348/6BHhr/wGn/8Aj1H/AA0348/6BHhr/wABp/8A49XiFFHtZ9w+s1f5j2//AIab8ef9Ajw1/wCA0/8A8eo/4ab8ef8AQI8Nf+A0/wD8erxCij2s+4fWav8AMe3/APDTfjz/AKBHhr/wGn/+PUf8NN+PP+gR4a/8Bp//AI9XiFFHtZ9w+s1f5j2//hpvx5/0CPDX/gNP/wDHqP8Ahpvx5/0CPDX/AIDT/wDx6vEKKPaz7h9Zq/zHt/8Aw0348/6BHhr/AMBp/wD49R/w0348/wCgR4a/8Bp//j1eIUUe1n3D6zV/mPb/APhpvx5/0CPDX/gNP/8AHqP+Gm/Hn/QI8Nf+A0//AMerxCij2s+4fWav8x7f/wANN+PP+gR4a/8AAaf/AOPUf8NN+PP+gR4a/wDAaf8A+PV4hRR7WfcPrNX+Y9v/AOGm/Hn/AECPDX/gNP8A/HqP+Gm/Hn/QI8Nf+A0//wAerxCij2s+4fWav8x7f/w0348/6BHhr/wGn/8Aj1H/AA0348/6BHhr/wABp/8A49XiFFHtZ9w+s1f5j2//AIab8ef9Ajw1/wCA0/8A8eo/4ab8ef8AQI8Nf+A0/wD8erxCij2s+4fWav8AMe3/APDTfjz/AKBHhr/wGn/+PUf8NN+PP+gR4a/8Bp//AI9XiFFHtZ9w+s1f5j2//hpvx5/0CPDX/gNP/wDHqP8Ahpvx5/0CPDX/AIDT/wDx6vEKKPaz7h9Zq/zHt/8Aw0348/6BHhr/AMBp/wD49R/w0348/wCgR4a/8Bp//j1eIUUe1n3D6zV/mPYfEH7RHjXW9B1DRrrS/D6W9/ay2srRW8wcJIhUlSZSM4PGQa8eooqZSctyJ1JT+J3P/9k=",
            "legenda": "Fotografia 03 — Detalhe aproximado do vestígio balístico (estojo cal. 9mm) arrecadado no solo.",
            "incluir": True
        }
    ]



def render_action_buttons(prefix):
        st.markdown('<br>', unsafe_allow_html=True)
        btn1, btn2, btn3, btn4, btn5, btn6, btn7 = st.columns([1, 1.2, 0.9, 1.1, 1.1, 1.1, 1.2])

        # 1. Salvar no Sistema
        if btn1.button("💾 Salvar", use_container_width=True, key=f"btn_salvar_{prefix}"):
            import sqlite3, json, os, datetime
            DB_PATH = os.path.join(os.path.dirname(__file__), "banco_laudos.sqlite")
            dados = {}
            for k, v in st.session_state.items():
                if k not in ['autenticado', 'preview_open']:
                    try:
                        if isinstance(v, (datetime.date, datetime.time)):
                            dados[k] = v.isoformat()
                        else:
                            dados[k] = v
                    except: pass
            if 'vitimas' in st.session_state: dados['vitimas'] = st.session_state['vitimas']
            if 'vestigios' in st.session_state: dados['vestigios'] = st.session_state['vestigios']
            if 'fotos' in st.session_state: dados['fotos'] = st.session_state['fotos']
            if 'quesitos_list' in st.session_state: dados['quesitos_list'] = st.session_state['quesitos_list']

            try:
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                mob_id = dados.get('ocorrencia', f"laudo_{len(dados)}")
                now_str = datetime.datetime.now().isoformat()
                c.execute('INSERT OR REPLACE INTO laudos (mobile_id, data_sincronizacao, dados_json) VALUES (?, ?, ?)',
                          (mob_id, now_str, json.dumps(dados)))
                conn.commit()
                conn.close()
                st.success("✅ Dados salvos com sucesso no sistema!")
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")

        # 2. Gerar Laudo
        gerar_clicked = btn2.button("🚀 Gerar Laudo (.docx)", type="primary", use_container_width=True, key=f"btn_gerar_{prefix}")

        # 3. Limpar Formulário
        if btn3.button("🗑️ Limpar", use_container_width=True, key=f"btn_limpar_{prefix}"):
            for k in list(st.session_state.keys()):
                if k not in ['autenticado']:
                    del st.session_state[k]
            st.rerun()

        # 4. Buscar Ocorrências (Modal)
        if btn4.button("🔍 Buscar", use_container_width=True, key=f"btn_sync_{prefix}"):
            modal_ocorrencias()

        # 5. Auditoria de Inconsistências (Checkup do Laudo)
        if btn5.button("🔍 Auditoria", use_container_width=True, key=f"btn_audit_{prefix}"):
            modal_auditoria_inconsistencias()

        # 6. Sugestões & Melhorias
        if btn6.button("💡 Sugestões", use_container_width=True, key=f"btn_sugestoes_{prefix}"):
            modal_sugestoes()

        return gerar_clicked


SUGESTOES_FILE = os.path.join(os.path.dirname(__file__), "sugestoes.json")

def carregar_sugestoes():
    if os.path.exists(SUGESTOES_FILE):
        try:
            with open(SUGESTOES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def salvar_sugestoes(lista):
    try:
        with open(SUGESTOES_FILE, "w", encoding="utf-8") as f:
            json.dump(lista, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar sugestões: {e}")
        return False

@st.dialog("💬 Central de Sugestões, Erros e Melhorias")
def modal_sugestoes():
    sugestoes = carregar_sugestoes()
    tab1, tab2 = st.tabs(["➕ Cadastrar Sugestão / Erro", f"📋 Acompanhar Sugestões ({len(sugestoes)})"])

    with tab1:
        st.markdown("Use este formulário para reportar bugs, divergências ou sugerir novas melhorias para o sistema:")
        tipo = st.selectbox("Tipo de Comunicação", ["💡 Sugestão de Melhoria", "🐛 Reportar Erro / Problema", "✨ Nova Funcionalidade", "Outro"], key="sug_tipo")
        titulo = st.text_input("Título / Resumo", placeholder="Ex: Ajustar layout dos cartões de vestígios", key="sug_titulo")
        autor = st.text_input("Seu Nome / Perito", value=st.session_state.get("perito", ""), placeholder="Ex: Dr. Carlos", key="sug_autor")
        descricao = st.text_area("Descrição detalhada", placeholder="Descreva com detalhes a sua sugestão ou o erro observado...", height=100, key="sug_desc")
        uploaded_imgs = st.file_uploader("📸 Anexar Imagens / Prints (Opcional)", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key="sug_imgs_input")

        if st.button("🚀 Cadastrar Sugestão", type="primary", use_container_width=True, key="btn_cadastrar_sugestao"):
            if not titulo.strip() or not descricao.strip():
                st.warning("⚠️ Por favor, preencha o título e a descrição da sugestão.")
            else:
                import datetime, base64
                imgs_b64 = []
                if uploaded_imgs:
                    for f_img in uploaded_imgs:
                        try:
                            imgs_b64.append(base64.b64encode(f_img.read()).decode("utf-8"))
                        except Exception: pass

                novo_item = {
                    "id": datetime.datetime.now().strftime("%Y%m%d%H%M%S"),
                    "data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "tipo": tipo,
                    "titulo": titulo.strip(),
                    "autor": autor.strip() if autor.strip() else "Perito Anônimo",
                    "descricao": descricao.strip(),
                    "imagens": imgs_b64,
                    "status": "🟡 Pendente"
                }
                sugestoes.insert(0, novo_item)
                salvar_sugestoes(sugestoes)
                st.success("✅ Sua sugestão foi cadastrada com sucesso! Você pode acompanhar o status na aba 'Acompanhar Sugestões'.")
                st.rerun()

    with tab2:
        if not sugestoes:
            st.info("Nenhuma sugestão cadastrada até o momento.")
        else:
            st.markdown("##### 📋 Sugestões Cadastradas e Acompanhamento de Status")
            opcoes_status = ["🟡 Pendente", "🔵 Em Análise", "🟠 Em Desenvolvimento", "🟢 Concluído / Executado", "🔴 Não Aplicável"]

            for idx, item in enumerate(sugestoes):
                with st.expander(f"{item.get('status', '🟡 Pendente')} | {item.get('titulo')} ({item.get('data')})"):
                    st.markdown(f"**Tipo:** {item.get('tipo')}")
                    st.markdown(f"**Autor:** {item.get('autor')}")
                    st.markdown(f"**Descrição:**\n{item.get('descricao')}")

                    if item.get("imagens"):
                        st.markdown("**📸 Imagens / Prints Anexados:**")
                        import base64
                        cols_img = st.columns(min(len(item["imagens"]), 3))
                        for i_idx, b64_str in enumerate(item["imagens"]):
                            with cols_img[i_idx % 3]:
                                try:
                                    img_data = base64.b64decode(b64_str)
                                    st.image(img_data, use_container_width=True, caption=f"Print #{i_idx+1}")
                                except Exception: pass

                    st.markdown("---")
                    st.markdown("**⚙️ Alterar Status (Administrador):**")
                    col_st1, col_st2 = st.columns([3, 1])
                    with col_st1:
                        curr_st = item.get("status", "🟡 Pendente")
                        curr_idx = opcoes_status.index(curr_st) if curr_st in opcoes_status else 0
                        novo_st = st.selectbox("Status", opcoes_status, index=curr_idx, key=f"sel_st_{item['id']}_{idx}")
                    with col_st2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("Salvar", key=f"btn_save_st_{item['id']}_{idx}", use_container_width=True):
                            sugestoes[idx]["status"] = novo_st
                            salvar_sugestoes(sugestoes)
                            st.success("Status atualizado!")
                            st.rerun()


def call_gemini_text(prompt, system_instruction=""):
    key = get_gemini_api_key()
    if not key:
        return None, "Chave GEMINI_API_KEY não configurada. Verifique se o arquivo CHAVE.txt está presente."
    last_err = None
    try:
        import google.generativeai as genai
        genai.configure(api_key=key)

        models = ["gemini-2.5-flash", "gemini-flash-latest", "gemini-3.6-flash"]
        for m in models:
            try:
                model = genai.GenerativeModel(
                    m,
                    system_instruction=system_instruction if system_instruction else None
                )
                resp = model.generate_content(prompt)
                if resp and resp.text:
                    return resp.text, None
            except Exception as e:
                last_err = e
                continue
    except Exception as e:
        last_err = e

    # Fallback via REST API
    try:
        import requests
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={key}"
        full_text = (system_instruction + "\n\n" + prompt) if system_instruction else prompt
        payload = {"contents": [{"parts": [{"text": full_text}]}]}
        r = requests.post(url, json=payload, timeout=35)
        if r.status_code == 200:
            res_j = r.json()
            txt = res_j['candidates'][0]['content']['parts'][0]['text']
            return txt, None
        else:
            return None, f"Erro na API do Gemini (HTTP {r.status_code}): {r.text}"
    except Exception as e_rest:
        return None, f"Erro ao conectar ao Gemini: {last_err or e_rest}"


def call_gemini_vision(prompt, image_input, mime_type="image/jpeg", json_mode=False):
    """
    Executa chamada multimodal ao Google Gemini API com safety_settings desativados (BLOCK_NONE) para uso médico-legal/forense.
    """
    key = get_gemini_api_key()
    if not key:
        return None, "Chave GEMINI_API_KEY não configurada. Verifique se o arquivo CHAVE.txt está presente."

    try:
        import google.generativeai as genai
        from google.generativeai.types import HarmCategory, HarmBlockThreshold

        genai.configure(api_key=key)

        generation_config = {}
        if json_mode:
            generation_config["response_mime_type"] = "application/json"

        safety_settings = {
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
        }

        models_to_try = ["gemini-2.5-flash", "gemini-flash-latest", "gemini-3.6-flash"]

        contents = [prompt]
        if isinstance(image_input, list):
            contents.extend(image_input)
        elif hasattr(image_input, "save") or str(type(image_input)).find("Image") != -1:
            contents.append(image_input)
        elif isinstance(image_input, bytes):
            contents.append({"mime_type": mime_type, "data": image_input})
        else:
            contents.append(image_input)

        last_err = None
        for m_name in models_to_try:
            try:
                model = genai.GenerativeModel(m_name, generation_config=generation_config if json_mode else None)
                resp = model.generate_content(contents, safety_settings=safety_settings)
                if resp and resp.text:
                    return resp.text, None
            except Exception as e:
                last_err = e
                continue

        return None, f"Erro ao executar Gemini Vision: {last_err}"
    except Exception as err:
        return None, str(err)



def ler_anotacoes_livres_gemini(texto_bruto):
    """
    Analisa anotações livres/brutas de campo usando Gemini IA e extrai dicionário estruturado.
    """
    prompt = f"""
Você é um perito criminalístico especialista em estruturação de dados de local de crime.
Examine as anotações livres/brutas fornecidas pelo perito e extraia os campos no formato JSON estrito:

{{
  "num_laudo": "Número do laudo se mencionado",
  "ocorrencia": "Número da ocorrência/BO",
  "requisicao": "Número da requisição",
  "perito": "Nome do perito relator",
  "autoridade_local": "Nome do delegado ou autoridade presente",
  "endereco": "Endereço completo do local",
  "ponto_referencia": "Ponto de referência",
  "area": "Descrição do tipo de área/pavimento",
  "delimitacoes": "Delimitações/limites físicos do local",
  "equipe_pm": "Equipe da Polícia Militar presente",
  "equipe_pc": "Equipe da Polícia Civil presente",
  "vitima_nome": "Nome da vítima se constar",
  "vitima_vestes": "Descrição das vestes da vítima",
  "vitima_posicao": "Posição/decúbito do cadáver",
  "vitima_lesoes": "Descrição das lesões na vítima",
  "vestigios_resumo": "Resumo dos vestígios encontrados",
  "dinamica_fatos": "Esboço ou rascunho da dinâmica dos fatos"
}}

Anotações do perito:
{texto_bruto}

Retorne APENAS o JSON estrito. Se algum dado não estiver mencionado, retorne string vazia.
"""

    res, err = call_gemini_text(prompt, system_instruction="Você é um assistente pericial de estruturação de dados. Responda apenas com JSON estrito.")
    if err or not res:
        return {}, err or "Sem resposta da IA"

    import json, re
    cleaned = res.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        return json.loads(cleaned), None
    except Exception as e:
        match = re.search(r"\{[\s\S]*\}", res)
        if match:
            try: return json.loads(match.group(0)), None
            except: pass
        return {}, f"Erro ao decodificar dados da IA: {e}"


def ler_requisicao_gemini_vision(file_bytes, file_name, file_type=""):
    """
    Lê uma Requisição Pericial (PDF ou Imagem) com Gemini Vision e retorna (extracted_text, dados_dict, error_msg).
    Prompt estruturado para extrair JSON com Delegacia, Delegado, Ocorrência, Requisição e Quesitos.
    """
    images = convert_bytes_to_pil_images(file_bytes, file_name, file_type)
    if not images:
        return "", {}, "Não foi possível converter o documento em imagem para análise visual pelo Gemini Vision."

    prompt = """
Você é um perito criminalístico especialista em análise documental pericial brasileira.
Examine a(s) imagem(ns) fornecida(s) da Requisição Pericial e extraia os seguintes dados no formato JSON estrito:

{
  "delegacia": "Nome da Delegacia de Polícia ou Órgão Solicitante",
  "delegado": "Nome da Autoridade Solicitante / Delegado(a) de Polícia (incluindo cargo ou 'Dr.(a)' se houver)",
  "ocorrencia": "Número do Boletim de Ocorrência / Ocorrência Policial",
  "requisicao": "Número da Requisição Pericial",
  "quesitos": "Texto integral de todos os quesitos / perguntas formuladas pela autoridade solicitante, numerados linha a linha"
}

Regras de extração:
1. Transcreva com fidelidade absoluta os números de ocorrência, requisição e quesitos.
2. Se algum campo não estiver visível na imagem, coloque string vazia "".
3. Retorne APENAS o JSON válido.
"""

    resp_text, err = call_gemini_vision(prompt, images, json_mode=True)
    if err or not resp_text:
        return "", {}, err or "Sem resposta do Gemini"

    import json
    import re
    cleaned = resp_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    try:
        dados = json.loads(cleaned)
        dados["origem"] = "Gemini Vision IA"
        return resp_text, dados, None
    except Exception as e:
        match = re.search(r"\{[\s\S]*\}", resp_text)
        if match:
            try:
                dados = json.loads(match.group(0))
                dados["origem"] = "Gemini Vision IA"
                return resp_text, dados, None
            except Exception:
                pass
        return resp_text, {}, f"Erro ao interpretar JSON do Gemini: {e}"


def ler_necropsia_gemini_vision(file_bytes, file_name, file_type=""):
    """
    Lê um Laudo de Necropsia / Exame Cadavérico (PDF ou Imagem) usando Gemini Vision.
    Extrai em JSON: causa_mortis, instrumento_lesivo, agente_instrumento, lesoes (lista), nome_vitima, documento_vitima, numero_laudo_iml.
    """
    images = convert_bytes_to_pil_images(file_bytes, file_name, file_type)
    if not images:
        return {}, "Não foi possível converter o arquivo em imagem para o Gemini Vision."

    prompt = """
Você é um médico legista e perito criminal especialista em tanatologia pericial e necropsia.
Examine a(s) imagem(ns) do Laudo de Necropsia / Auto de Exame Cadavérico e extraia as informações em formato JSON estrito:

{
  "causa_mortis": "Descrição exata e detalhada da causa mortis (ex: Traumatismo cranioencefálico decorrente de PAF, Choque hipovolêmico por ferimento perfurocortante, etc.)",
  "instrumento_lesivo": "Tipo de ação do instrumento lesivo (escolha preferencialmente uma das opções: Perfurocontundente, Cortante, Perfurante, Perfurocortante, Contundente, Cortocontundente, Ação Térmica)",
  "agente_instrumento": "Agente/instrumento específico se mencionado (ex: projétil de arma de fogo, lâmina de faca, instrumento contundente maciço)",
  "lesoes": [
    "Descrição detalhada e técnica da lesão 1 (localização anatômica, tipo, formato, dimensões)",
    "Descrição detalhada e técnica da lesão 2",
    "..."
  ],
  "nome_vitima": "Nome da vítima / examinando se constar",
  "documento_vitima": "RG/CPF ou número de identificação da vítima se constar",
  "numero_laudo_iml": "Número do Laudo do IML / Necropsia"
}

Regras:
1. Analise o exame necroscópico minuciosamente.
2. Extraia cada lesão identificada como um item detalhado na lista "lesoes".
3. Retorne apenas o JSON estrito.
"""

    resp_text, err = call_gemini_vision(prompt, images, json_mode=True)
    if err or not resp_text:
        return {}, err or "Sem resposta do Gemini"

    import json
    import re
    cleaned = resp_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    try:
        dados = json.loads(cleaned)
        return dados, None
    except Exception as e:
        match = re.search(r"\{[\s\S]*\}", resp_text)
        if match:
            try:
                dados = json.loads(match.group(0))
                return dados, None
            except Exception:
                pass
        return {}, f"Erro ao decodificar JSON de Necropsia: {e}"


def gerar_legenda_foto_gemini(b64_img, prompt_custom=None):
    """
    Gera legenda descritiva pericial técnica para uma fotografia de local/vestígio usando Gemini Vision.
    Possui fallback automático para laudos forenses em caso de restrição.
    """
    import base64
    from io import BytesIO
    from PIL import Image

    try:
        img_bytes = base64.b64decode(b64_img)
        pil_img = Image.open(BytesIO(img_bytes)).convert("RGB")
    except Exception as e:
        return "Visualização do registro fotográfico pericial no local dos fatos.", f"Erro ao processar imagem: {e}"

    prompt = prompt_custom or """
Você é um perito criminal relator da Polícia Científica e Medicina Legal.
Esta fotografia pertence a um procedimento oficial de investigação criminal e exame de local de crime / necropsia.
Descreva a imagem em uma frase técnica, objetiva, imparcial e formal para constar no Apêndice Fotográfico do Laudo Pericial.

Exemplos de formato pericial:
- "Visualização do corpo da vítima em decúbito dorsal na posição inicial do exame pericial."
- "Visão geral do sítio pericial, evidenciando o cadáver no solo."
- "Detalhe aproximado do vestígio/elemento registrado na cena dos fatos."

Regras:
1. Forneça APENAS o texto da legenda pericial (1 frase).
2. Não use aspas ou prefixos.
"""

    resp_text, err = call_gemini_vision(prompt, pil_img)
    if resp_text and resp_text.strip():
        cleaned = resp_text.strip().strip('"').strip("'")
        return cleaned, None

    # Fallback padronizado pericial caso o filtro de visão omita o retorno
    fallback_caption = "Visualização do corpo da vítima em decúbito dorsal na posição inicial do exame pericial."
    return fallback_caption, None


@st.dialog("✨ Polir Redação com IA (Tom Formal-Jurídico)")
def modal_polir_redacao(field_key, field_label):
    st.markdown("### ✨ Polimento de Redação Forense com IA")
    st.markdown(f"**Campo Selecionado:** `{field_label}`")

    texto_atual = st.session_state.get(field_key, "")
    if isinstance(texto_atual, list):
        texto_atual = "\n".join([str(x) for x in texto_atual if str(x).strip()])

    if not str(texto_atual).strip():
        st.warning("⚠️ O campo selecionado está vazio. Digite um texto antes de solicitar o polimento.")
        return

    st.caption("Texto Original Atual:")
    st.text_area("Texto Original", value=str(texto_atual), height=110, disabled=True, key=f"polir_orig_{field_key}")

    key_temp = f"polished_temp_{field_key}"
    if key_temp not in st.session_state:
        st.session_state[key_temp] = ""

    col1, col2 = st.columns(2)
    with col1:
        if st.button("✨ Executar Polimento com IA", type="primary", use_container_width=True, key=f"btn_exec_polir_{field_key}"):
            with st.spinner("Refinando redação e tom formal-jurídico com a IA Gemini..."):
                sys_instruction = (
                    "Você é um perito criminal especialista e revisor de laudos periciais oficiais da Polícia Civil. "
                    "Reescreva o texto a seguir elevando o padrão para a norma culta formal-jurídica, "
                    "utilizando terminologia pericial precisa, tom impessoal e clareza técnico-científica. "
                    "NUNCA invente fatos, dados, objetos ou lesões adicionais. Mantenha 100% dos fatos originais."
                )
                prompt = f"Reescreva e polia o texto pericial abaixo mantendo integralmente todos os fatos e dados técnicos:\n\n{texto_atual}"
                res, err = call_gemini_text(prompt, system_instruction=sys_instruction)
                if err:
                    st.error(f"Erro ao polir texto: {err}")
                else:
                    st.session_state[key_temp] = res
                    st.success("✅ Texto polido gerado com sucesso!")

    with col2:
        if st.session_state.get(key_temp):
            if st.button("✅ Aplicar Texto Polido no Laudo", type="secondary", use_container_width=True, key=f"btn_apply_polir_{field_key}"):
                st.session_state[field_key] = st.session_state[key_temp]
                st.session_state[key_temp] = ""
                st.toast("✅ Texto polido aplicado ao laudo!")
                st.rerun()

    if st.session_state.get(key_temp):
        st.markdown("**Resultado da Redação Polida (Sugestão IA):**")
        st.text_area("Texto Polido", value=st.session_state[key_temp], height=150, key=f"polir_res_{field_key}")


@st.dialog("🔍 Auditoria de Inconsistências (Checkup do Laudo)")
def modal_auditoria_inconsistencias():
    st.markdown("### 🛡️ Checkup de Qualidade e Coerência Forense")
    st.markdown("O sistema executa verificações determinísticas de formulário e uma análise lógica profunda via Gemini IA:")

    num_laudo = st.session_state.get("num_laudo", "")
    ocorrencia = st.session_state.get("ocorrencia", "")
    perito = st.session_state.get("perito", "")
    autoridade = st.session_state.get("autoridade_sel", "")
    requisicao = st.session_state.get("requisicao", "")
    endereco = st.session_state.get("endereco", "")
    delimitacoes = st.session_state.get("delimitacoes", "")
    iso_estado = st.session_state.get("iso_estado", "")
    inst_acao = st.session_state.get("inst_acao", "")
    vitimas = st.session_state.get("vitimas", [])
    vestigios = st.session_state.get("vestigios", [])
    quesitos = st.session_state.get("quesitos_list", [])
    dinamica = st.session_state.get("dinamica_fatos", "")

    # Rule-based validation engine
    erros = []
    alertas = []
    ok_items = []

    # 1. Header & Essenciais
    if not str(num_laudo).strip(): erros.append("Nº do Laudo não preenchido.")
    else: ok_items.append("Nº do Laudo preenchido.")
    if not str(ocorrencia).strip(): erros.append("Nº da Ocorrência não preenchido.")
    else: ok_items.append("Nº da Ocorrência preenchido.")
    if not str(perito).strip(): erros.append("Perito(a) Relator(a) não informado.")
    else: ok_items.append("Perito Relator informado.")
    if not str(autoridade).strip(): erros.append("Autoridade Solicitante não selecionada.")
    else: ok_items.append("Autoridade Solicitante informada.")
    if not str(requisicao).strip(): erros.append("Número da Requisição não preenchido.")
    else: ok_items.append("Número de Requisição informado.")
    if not str(endereco).strip(): erros.append("Endereço do local não informado.")
    else: ok_items.append("Endereço informado.")

    # 2. Vítimas & Lesões vs Instrumento
    todas_lesoes = []
    for idx_v, vt in enumerate(vitimas, 1):
        if not vt.get("nome", "").strip():
            alertas.append(f"Vítima #{idx_v} está sem nome/identificação definida.")
        lesoes = [l.strip() for l in vt.get("lesoes", []) if l.strip()]
        if not lesoes:
            alertas.append(f"Vítima #{idx_v} ({vt.get('nome') or 'N/I'}) não possui lesões descritas.")
        todas_lesoes.extend(lesoes)

    lesoes_concat = " ".join(todas_lesoes).lower()
    inst_str = str(inst_acao).lower()
    if "perfurocontundente" in inst_str or inst_str.startswith("1"):
        if "paf" not in lesoes_concat and "perfurocontundente" not in lesoes_concat and "projétil" not in lesoes_concat and "entrada" not in lesoes_concat and "disparo" not in lesoes_concat:
            alertas.append("Instrumento selecionado é Perfurocontundente (PAF), mas a descrição de lesões da(s) vítima(s) não menciona PAF, disparos ou orifícios de entrada/saída.")
    elif "cortante" in inst_str or inst_str.startswith("2"):
        if "corte" not in lesoes_concat and "incisa" not in lesoes_concat and "gume" not in lesoes_concat and "faca" not in lesoes_concat:
            alertas.append("Instrumento selecionado é Cortante, mas lesões não descrevem feridas incisas ou lâmina/gume.")

    # 3. Vestígios
    if not vestigios:
        alertas.append("Nenhum vestígio registrado no laudo.")
    else:
        ok_items.append(f"{len(vestigios)} vestígio(s) registrado(s).")
        for idx_vest, vest in enumerate(vestigios, 1):
            if not vest.get("tipo"):
                erros.append(f"Vestígio #{idx_vest} não possui Tipo selecionado.")
            elif not vest.get("descricao") and not vest.get("localizacao") and not vest.get("subtipo"):
                alertas.append(f"Vestígio #{idx_vest} ({vest.get('tipo')}) possui descrição/localização incompleta.")

    # 4. Quesitos
    quesitos_sem_resposta = [i+1 for i, q in enumerate(quesitos) if q.get("pergunta", "").strip() and not q.get("resposta", "").strip()]
    if quesitos_sem_resposta:
        alertas.append(f"Existem quesitos sem resposta formulada: Quesito(s) nº {quesitos_sem_resposta}.")

    # Render Rule-Based Metrics
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("🔴 Erros/Incompletos", len(erros))
    c_m2.metric("🟡 Alertas & Atenção", len(alertas))
    c_m3.metric("🟢 Verificados OK", len(ok_items))

    if erros:
        st.error("**Inconsistências Críticas:**\n" + "\n".join([f"• {e}" for e in erros]))
    if alertas:
        st.warning("**Alertas e Recomendações:**\n" + "\n".join([f"• {a}" for a in alertas]))
    if ok_items and not erros:
        st.success("**Itens de Formulário Checados:**\n" + "\n".join([f"✓ {o}" for o in ok_items]))

    st.divider()

    st.markdown("### 🤖 Auditoria Lógica Profunda com Gemini IA")
    if st.button("🚀 Executar Checkup Forense com IA", type="primary", use_container_width=True, key="btn_run_ai_audit"):
        with st.spinner("Analisando dados com a IA Gemini em busca de incoerências e omissões..."):
            dados_completos = {
                "num_laudo": num_laudo, "ocorrencia": ocorrencia, "perito": perito, "autoridade": autoridade,
                "requisicao": requisicao, "endereco": endereco, "delimitacoes": delimitacoes,
                "isolamento": iso_estado, "instrumento": inst_acao, "vitimas": vitimas,
                "vestigios": vestigios, "quesitos": quesitos, "dinamica_fatos": dinamica
            }
            prompt_audit = (
                "Você é um perito auditor sênior em criminalística e medicina legal. Analise o conjunto de dados do laudo pericial a seguir "
                "e verifique rigorosamente se existem:\n"
                "1. Contradições lógicas entre a posição da vítima, lesões e vestígios encontrados;\n"
                "2. Incompatibilidades entre a ação do instrumento e os exames perinicroscópicos;\n"
                "3. Incoerências na preservação/isolamento e na cadeia de custódia;\n"
                "4. Falhas ou omissões importantes que possam gerar nulidades no processo penal.\n\n"
                f"DADOS DO LAUDO:\n{json.dumps(dados_completos, ensure_ascii=False, indent=2, default=str)}\n\n"
                "Retorne o parecer técnico formatado com seções claras: 🚨 CONTRADIÇÕES LÓGICAS, ⚠️ RECOMENDAÇÕES DE AJUSTE, e 💡 SUGESTÕES COMPLEMENTARES."
            )
            sys_audit = "Você é um auditor pericial rigoroso. Seja imparcial, objetivo, técnico e ajude o perito a blindar o laudo contra erros ou contestação jurídica."
            res, err = call_gemini_text(prompt_audit, system_instruction=sys_audit)
            if err:
                st.error(f"Erro na auditoria IA: {err}")
            else:
                st.session_state["resultado_auditoria_ia"] = res
                st.success("✅ Auditoria IA concluída com sucesso!")

    if st.session_state.get("resultado_auditoria_ia"):
        st.markdown(st.session_state["resultado_auditoria_ia"])


def modal_auditoria_ia():
    modal_auditoria_inconsistencias()


def get_logo_base64():
    import base64
    logo_path = os.path.join(os.path.dirname(__file__), "logo_pericia.png")
    if os.path.exists(logo_path):
        try:
            with open(logo_path, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            return ""
    return ""


def ler_requisicao_pericial(file_bytes, file_name, file_type=""):
    """
    Extrai texto e campos (Delegacia, Delegado, Ocorrência, Requisição, Quesitos)
    de arquivos de Requisição Pericial em formato PDF ou Imagem.
    Prioriza o Gemini Vision IA quando a chave está configurada e faz fallback para OCR/Regex.
    """
    # 1. Tenta extração via Gemini Vision IA se a API Key estiver configurada
    if get_gemini_api_key():
        resp_text, dados_g, err_g = ler_requisicao_gemini_vision(file_bytes, file_name, file_type)
        if dados_g and any(dados_g.values()):
            return resp_text, dados_g

    # 2. Fallback para métodos locais (pdfplumber, pypdf, pytesseract OCR + Regex)
    from io import BytesIO
    import re

    extracted_text = ""
    file_name_lower = file_name.lower()
    is_pdf = file_type == "application/pdf" or file_name_lower.endswith(".pdf")
    is_img = file_type.startswith("image/") or file_name_lower.endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff"))

    if is_pdf:
        try:
            import pdfplumber
            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        extracted_text += t + "\n"
        except Exception:
            pass

        if not extracted_text.strip():
            try:
                import pypdf
                reader = pypdf.PdfReader(BytesIO(file_bytes))
                for page in reader.pages:
                    t = page.extract_text()
                    if t:
                        extracted_text += t + "\n"
            except Exception:
                pass

        if not extracted_text.strip():
            try:
                import pdfplumber
                import pytesseract
                with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                    for page in pdf.pages:
                        pil_img = page.to_image(resolution=150).original
                        try:
                            t = pytesseract.image_to_string(pil_img, lang="por")
                        except Exception:
                            t = pytesseract.image_to_string(pil_img)
                        if t:
                            extracted_text += t + "\n"
            except Exception:
                pass

    elif is_img:
        try:
            import pytesseract
            from PIL import Image
            img = Image.open(BytesIO(file_bytes))
            try:
                extracted_text = pytesseract.image_to_string(img, lang="por")
            except Exception:
                extracted_text = pytesseract.image_to_string(img)
        except Exception:
            pass

    dados = {}
    if extracted_text.strip():
        m_del = re.search(r'(?:DELEGACIA(?: DE POLÍCIA)?|UNIDADE SOLICITANTE|ÓRGÃO SOLICITANTE|ORIGEM|DESTINO)[\s:]+([^\n\r]+)', extracted_text, re.IGNORECASE)
        if not m_del:
            m_del = re.search(r'(DELEGACIA DE POLÍCIA [^\n\r]+)', extracted_text, re.IGNORECASE)
        if not m_del:
            m_del = re.search(r'(DELEGACIA [^\n\r]+)', extracted_text, re.IGNORECASE)
        if m_del:
            dados["delegacia"] = m_del.group(1).strip()

        # Extrai Delegado / Autoridade Solicitante
        m_aut = re.search(r'(?:DELEGAD[OA](?: DE POLÍCIA)?|AUTORIDADE SOLICITANTE|AUTORIDADE POLICIAL|SOLICITANTE)[\s:]+([^\n\r]+)', extracted_text, re.IGNORECASE)
        if not m_aut:
            m_aut = re.search(r'(?:Dr[a]?\.\s+)([^\n\r]+)', extracted_text, re.IGNORECASE)
        if m_aut:
            dados["delegado"] = m_aut.group(1).strip()

        # Extrai Ocorrência / BO
        m_oco = re.search(r'(?:OCORRÊNCIA|OCORRENCIA|BOLETIM DE OCORRÊNCIA|BO)[\s\:\.\º\°]*(?:N[º°\.\:]*)?[\s\:\.\º\°]*([0-9A-Za-z\/\-\.]{3,})', extracted_text, re.IGNORECASE)
        if m_oco:
            dados["ocorrencia"] = m_oco.group(1).strip()

        # Extrai Requisição
        m_req = re.search(r'(?:REQUISIÇÃO|REQUISICAO|REQ\.?)[\s\:\.\º\°]*(?:N[º°\.\:]*)?[\s\:\.\º\°]*([0-9A-Za-z\/\-\.]{3,})', extracted_text, re.IGNORECASE)
        if m_req:
            dados["requisicao"] = m_req.group(1).strip()

        # Extrai Quesitos
        m_que = re.search(r'(?:QUESITOS(?: DA AUTORIDADE| FORMULADOS)?|PERGUNTAS)[\s\:\-\n\r]+([\s\S]+?)(?:\n\s*\n[A-Z0-9\s]{4,}:|\Z)', extracted_text, re.IGNORECASE)
        if m_que:
            dados["quesitos"] = m_que.group(1).strip()
        else:
            num_q = re.findall(r'(\d+[\.\)-]\s*[^\n]+)', extracted_text)
            if num_q:
                dados["quesitos"] = "\n".join(num_q)

    return extracted_text, dados


# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
for key, default in {
    "autoridades":   AUTORIDADES_DEFAULT.copy(),
    "preview_open":  False,
    "show_nova_aut": False,
    "docx_bytes":    None,
    "docx_filename": "",
    "quesitos":      "",
    "quesitos_list": [{"pergunta": "", "resposta": ""}],
    "vitimas":       [{"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "", "naturalidade": "",
                       "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "",
                       "fenomenos": "", "lesoes": [""]}],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ── Reset e Base ── */
html, body, [class*="css"], .stApp {
    font-family: 'Inter', sans-serif !important;
    background-color: #e2e8f0 !important;
    color: #0f172a !important;
}


/* ── App Header (Título) ── */
.app-header {
    background-color: #ffffff;
    padding: 24px;
    border-radius: 16px;
    border: 1px solid #e2e8f0;
    margin-bottom: 24px;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
}
.app-header-brand {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 16px;
}
.app-header-badge {
    background: #eff6ff;
    color: #2563eb;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 700;
    width: fit-content;
}
.app-header-title {
    font-size: 24px;
    font-weight: 800;
    color: #0f172a;
    line-height: 1.2;
}
.app-header-sub {
    font-size: 14px;
    color: #64748b;
}


/* ── Forçar Colunas em Grid no Mobile ── */
@media (max-width: 768px) {
    div[data-testid="column"] {
        width: 50% !important;
        flex: 1 1 calc(50% - 1rem) !important;
        min-width: 40% !important;
    }
    div[data-testid="stHorizontalBlock"] {
        flex-wrap: wrap !important;
        flex-direction: row !important;
    }
}

/* Esconder o Header Padrão do Streamlit */
header[data-testid="stHeader"] { display: none !important; }

/* ── Containers com borda (st.container border=True) ── */
div[data-testid="stVerticalBlock"]:has(.custom-border-marker) {
    background-color: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 16px !important;
    box-shadow: 0 10px 15px -3px rgba(0,0,0,0.05), 0 4px 6px -4px rgba(0,0,0,0.05) !important;
    padding: 24px !important;
    margin-bottom: 24px !important;
}


/* ── Título das seções ── */
.section-title {
    font-size: 22px !important;
    font-weight: 800 !important;
    color: #0f172a !important;
    margin-top: 12px !important;
    margin-bottom: 20px !important;
    padding-bottom: 10px !important;
    border-bottom: 2px solid #cbd5e1 !important;
    display: flex !important;
    align-items: center !important;
    gap: 10px !important;
    line-height: 1.3 !important;
}


/* ── Botão primário ── */
.stButton > button[kind="primary"] {
    background: #2563eb !important;
    border: none !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    color: #fff !important;
    padding: 12px 24px !important;
    transition: all 0.2s !important;
}
.stButton > button[kind="primary"]:hover {
    background: #1d4ed8 !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 10px 15px -3px rgba(37,99,235,0.3) !important;
}

/* ── Botão secundário ── */
.stButton > button[kind="secondary"] {
    background: #f8fafc !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    color: #334155 !important;
    transition: all 0.2s !important;
}
.stButton > button[kind="secondary"]:hover {
    background: #e2e8f0 !important;
}

/* ── Inputs e Selects ── */
div[data-testid="stTextInput"] input,
div[data-testid="stTextArea"] textarea,
div[data-testid="stDateInput"] input,
div[data-testid="stTimeInput"] input,
div[data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 10px !important;
    box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05) !important;
    color: #0f172a !important;
}

/* Labels */
div[data-testid="stTextInput"] label,
div[data-testid="stTextArea"] label,
div[data-testid="stSelectbox"] label,
div[data-testid="stDateInput"] label,
div[data-testid="stTimeInput"] label {
    font-size: 13px !important;
    font-weight: 600 !important;
    color: #475569 !important;
    margin-bottom: 4px !important;
}

/* Expanders */
.streamlit-expanderHeader {
    background-color: #ffffff !important;
    border-radius: 12px !important;
    border: 1px solid #e2e8f0 !important;
    color: #1e293b !important;
    font-weight: 600 !important;
}
div[data-testid="stExpander"] {
    border: none !important;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05) !important;
    border-radius: 12px !important;
    overflow: hidden;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background-color: #ffffff !important;
    border-right: 1px solid #e2e8f0 !important;
}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# AUTENTICAÇÃO DO SISTEMA
# ═══════════════════════════════════════════════════════════
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = True

# Força bypass:
st.session_state["autenticado"] = True

if not st.session_state["autenticado"]:
    st.markdown("<h2 style='text-align: center; color: #1a3a6e; margin-top: 50px;'>🔒 Acesso Restrito ao Sistema</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b; margin-bottom: 30px;'>Por favor, insira as credenciais de acesso para prosseguir.</p>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.container(border=True):
            with st.form("login_form"):
                user = st.text_input("Usuário")
                pwd = st.text_input("Senha", type="password")
                submitted = st.form_submit_button(
                    "Entrar", type="primary", use_container_width=True)

                if submitted:
                    if (user.lower() in ["alvaro", "perito"]) and pwd in ["123", "icrim123"]:
                        st.session_state["autenticado"] = True
                        st.rerun()
                    else:
                        st.error("❌ Usuário ou senha incorretos.")

    st.stop()  # Bloqueia a execução do restante do código


# ═══════════════════════════════════════════════════════════
# MODO PREVIEW
# ═══════════════════════════════════════════════════════════
if st.session_state.preview_open:

    def ss(key, default=""):
        return st.session_state.get(key, default) or default

    num_laudo = ss("num_laudo")
    ocorrencia = ss("ocorrencia")
    perito = ss("perito")
    autoridade = ss("autoridade_sel")
    requisicao = ss("requisicao")
    referencia = ss("referencia") or "(sem referência)"
    destino = ss("destino")

    vitimas_list = st.session_state.get("vitimas", [{"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "",
                                        "naturalidade": "", "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "", "fenomenos": "", "lesoes": [""]}])
    nomes = [vt["nome"] for vt in vitimas_list if vt["nome"].strip()]
    vitima = ", ".join(nomes) if nomes else "________"

    vitimas_detalhes_html = ""
    for idx, vt in enumerate(vitimas_list, 1):
        nome_v = vt.get("nome", "").strip() or "Não identificado"
        doc_v = vt.get("documento", "").strip() or "Não informado"
        sexo_v = vt.get("sexo", "").strip() or "Não informado"
        dob_v = vt.get("data_nascimento", None)
        dob_str = data_simples(dob_v) if dob_v else "Não informado"
        fili_v = vt.get("filicao", "").strip() or "Não informada"
        nat_v = vt.get("naturalidade", "").strip() or "Não informada"
        vest_v = vt.get("vestes", "").strip() or "Não descritas"
        pert_v = vt.get("pertences", "").strip() or "Não descritos"
        loc_v = vt.get("localizacao", "").strip() or "Não descrita"
        pos_v = vt.get("posicao", "").strip() or "Não descrita"
        cab_v = vt.get("cabeca", "").strip() or "Não descrita"
        memb_v = vt.get("membros", "").strip() or "Não descritos"
        fen_v = vt.get("fenomenos", "").strip() or "Não descritos"

        lesoes_html = ""
        for les in vt.get("lesoes", [""]):
            if les.strip():
                lesoes_html += f"• {les.strip()}<br>"
        if not lesoes_html:
            lesoes_html = "• Não descritas<br>"

        vitimas_detalhes_html += f"""
                <div style="font-weight:bold; margin-top:0.6cm; font-size:11.5pt; color:#0f2044; text-indent:0px;">
                3.3.{idx} – Descrição da Vítima {idx}
                </div>

                <div style="font-size:10.5pt; line-height:1.6; padding-left:1.5cm; text-indent:0px;">
                <div style="font-weight:bold; margin-top:0.3cm; color:#0f2044; border-bottom:1px solid #e2e8f0; width:100%;">Identificação</div>
                a) Nome: {nome_v}<br>
                b) Documento: {doc_v}<br>
                c) Sexo: {sexo_v}<br>
                d) Data de Nascimento: {dob_str}<br>
                e) Filiação: {fili_v}<br>
                f) Naturalidade: {nat_v}<br>

                <div style="font-weight:bold; margin-top:0.3cm; color:#0f2044; border-bottom:1px solid #e2e8f0; width:100%;">Vestes e Pertences</div>
                a) Vestes: {vest_v}<br>
                b) Pertences: {pert_v}<br>

                <div style="font-weight:bold; margin-top:0.3cm; color:#0f2044; border-bottom:1px solid #e2e8f0; width:100%;">Posição e Localização (In Situ)</div>
                a) Localização: {loc_v}<br>
                b) Posição: {pos_v}<br>
                c) Cabeça: {cab_v}<br>
                d) Membros: {memb_v}<br>

                <div style="font-weight:bold; margin-top:0.3cm; color:#0f2044; border-bottom:1px solid #e2e8f0; width:100%;">Exame Perinecroscópico e Estado de Conservação</div>
                a) Fenômenos Transformativos: {fen_v}<br>
                b) Lesões:<br>
                {lesoes_html}
                </div>
                """
        dp_val = st.session_state.get("data_pericia_input",     date.today())
        da_val = st.session_state.get("data_atendimento_input", date.today())
        horario_val = st.session_state.get("horario",               None)
        horario_a_val = st.session_state.get("horario_atendimento",   None)
        endereco = ss("endereco")
        latitude = ss("latitude")
        longitude = ss("longitude")
        ponto = ss("ponto_referencia")
        area = ss("area")
        pavimento = ss("pavimento")
        delimitacoes = ss("delimitacoes")
        iso_s = ss("isolamento")
        isolamento = ss("isolamento_outros") if iso_s == "Outros" else iso_s
        clim_s = ss("clima")
        clima = ss("clima_outros") if clim_s == "Outros" else clim_s
        visibilidade = ss("visibilidade")
        iluminacao = ss("iluminacao")
        equipe_pm = ss("equipe_pm")
        equipe_pc = ss("equipe_pc")
        aut_local = ss("autoridade_local")

        dp_ext = data_extenso(dp_val)
        da_simp = data_simples(da_val)
        h_str = horario_val.strftime("%H:%M") if horario_val else "____"
        ha_str = horario_a_val.strftime("%H:%M") if horario_a_val else "____"

        tipo_local_val = st.session_state.get("tipo_local", "")
        ano = date.today().year

        # Barra superior
        st.markdown(f"""
                <div style="background:linear-gradient(135deg,#0f2044,#1a3a6e); padding:0 52px;
                height:68px; display:flex; align-items:center; justify-content:space-between;
                box-shadow:0 2px 16px rgba(0,0,0,0.25);">
                <div style="display:flex; align-items:center; gap:18px;">
                <span style="background:rgba(255,255,255,0.12); border:1px solid rgba(255,255,255,0.25);
                border-radius:6px; padding:5px 13px; font-size:10px; font-weight:700;
                color:#a8c4f0; letter-spacing:1.8px;">ICRIM / NPT</span>
                <div>
                <div style="font-size:17px; font-weight:700; color:#fff;">Pré-visualização do Laudo</div>
                <div style="font-size:12px; color:#7fa8d4;">Modo leitura — formulário desativado</div>
                </div>
                </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

        _, btn_col, _ = st.columns([3, 2, 3])
        with btn_col:
            if st.button("✕  Fechar e Voltar ao Formulário", type="primary", use_container_width=True):
                st.session_state.preview_open = False
                st.rerun()

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

        st.markdown(f"""
                <div class="a4-wrapper">
                <div class="a4-sheet">
                <div style="text-align:center; font-weight:bold; font-size:10.5pt;
                margin-bottom:1.2cm; line-height:1.6; color:#0f2044;">
                GOVERNO DO ESTADO DO MARANHÃO<br>
                SECRETARIA DE ESTADO DA SEGURANÇA PÚBLICA<br>
                POLÍCIA CIVIL<br>
                SUPERINTENDÊNCIA DE POLÍCIA TÉCNICO-CIENTÍFICA<br>
                INSTITUTO DE CRIMINALÍSTICA DE IMPERATRIZ
                </div>
                <div style="text-align:center; font-weight:bold; font-size:12pt;
                margin-bottom:1cm; text-decoration:underline; color:#0f2044;">
                LAUDO DE EXAME PERICIAL EM LOCAL DE MORTE VIOLENTA<br>
                Nº {v(num_laudo)}.{ano}/PO
                </div>
                <table style="width:100%; font-size:10.5pt; border-collapse:collapse; margin-bottom:1cm;">
                {''.join(f'<tr><td style="font-weight:bold;width:195px;padding:5px 0;vertical-align:top;color:#0f2044;">{lbl}:</td><td style="padding:5px 0;">{val}</td></tr>'
                for lbl, val in [
                ("NATUREZA DA PERÍCIA","Local de Morte Violenta"),
                ("OCORRÊNCIA Nº", v(ocorrencia)),
                ("PERITO(A) RELATOR(A)", v(perito)),
                ("AUTORIDADE SOLICITANTE", v(autoridade)),
                ("REQUISIÇÃO", v(requisicao)),
                ("VÍTIMA(S)", v(vitima)),
                ("INVESTIGADO(S)", investigado),
                ("IP/PM/TCO/BO/PROC", referencia),
                ("UNIDADE DESTINO", v(destino)),
                ])}
                </table>
                <div style="font-weight:bold;margin-top:1cm;margin-bottom:0.4cm;font-size:12pt;
                color:#0f2044;border-bottom:1px solid #cbd5e1;padding-bottom:4px;">
                1 — HISTÓRICO
                </div>
                <p style="text-align:justify;text-indent:1.5cm;margin-bottom:0.5cm;font-size:10.5pt;">
                {dp_ext}, neste Estado do Maranhão, o Diretor do Instituto de Criminalística de
                Imperatriz (ICRIM) designou o(a) Perito(a) Criminal <b>{v(perito)}</b>, para realização
                de exame pericial em Local de Morte Violenta, referente à Ocorrência nº
                <b>{v(ocorrencia)}</b>, requisitada por <b>{v(autoridade)}</b>, sob Requisição nº
                <b>{v(requisicao)}</b>.
                </p>
                <p style="text-align:justify;text-indent:1.5cm;margin-bottom:0.5cm;font-size:10.5pt;">
                Atendendo à solicitação, os peritos signatários compareceram ao local
                no dia <b>{da_simp}</b> às <b>{ha_str}</b>, encontrando as condições descritas a seguir.
                </p>
                <div style="font-weight:bold;margin-top:1cm;margin-bottom:0.4cm;font-size:12pt;
                color:#0f2044;border-bottom:1px solid #cbd5e1;padding-bottom:4px;">
                2 — DO LOCAL
                </div>
                <p style="text-align:justify;text-indent:1.5cm;font-size:10.5pt;line-height:2;">
                <b>ENDEREÇO:</b> {v(endereco)}<br>
                <b>COORDENADAS:</b> Lat: {v(latitude)} / Lon: {v(longitude)}<br>
                <b>PONTO DE REFERÊNCIA:</b> {v(ponto)}<br>
                <b>ÁREA:</b> {v(area)} &nbsp; <b>PAVIMENTO:</b> {v(pavimento)}<br>
                <b>DELIMITAÇÕES:</b> {v(delimitacoes)}<br>
                <b>ISOLAMENTO:</b> {v(isolamento)}<br>
                <b>CLIMA:</b> {v(clima)} &nbsp; <b>VISIBILIDADE:</b> {v(visibilidade)} &nbsp; <b>ILUMINAÇÃO:</b> {v(iluminacao)}<br>
                <b>EQUIPE PM:</b> {v(equipe_pm)} &nbsp; <b>EQUIPE PC:</b> {v(equipe_pc)}<br>
                <b>AUTORIDADE NO LOCAL:</b> {v(aut_local)}
                </p>

                <div style="font-weight:bold;margin-top:1.2cm;margin-bottom:0.4cm;font-size:12pt;
                color:#0f2044;border-bottom:1px solid #cbd5e1;padding-bottom:4px;">
                3 — EXAMES PERICIAIS
                </div>
                <div style="font-weight:bold;margin-top:0.5cm;margin-bottom:0.3cm;font-size:11.5pt;
                color:#0f2044;">
                3.3 — Das Vítimas
                </div>
                {vitimas_detalhes_html}

                <div style="text-align:center;font-size:9pt;margin-top:3cm;
                border-top:1px solid #94a3b8;padding-top:12px;color:#64748b;">
                Instituto de Criminalística de Imperatriz — ICRIM/NPT &nbsp;|&nbsp; Polícia Civil do Estado do Maranhão
                </div>
                </div>
                </div>
                """, unsafe_allow_html=True)


else:
    # ═══════════════════════════════════════════════════════════
    # MODO FORMULÁRIO (LOGIN DESABILITADO TEMPORARIAMENTE)
    # ═══════════════════════════════════════════════════════════
    st.session_state['logged_in'] = True

    _, main, _ = st.columns([0.04, 0.92, 0.04])


GAS_URL = "https://script.google.com/macros/s/AKfycbx9N4hidxHbfAUUpHVDAOadZBF5SRB9x9UzC0nt3k-wW0pgLThCHwBQsaUMJjtcTL1QJw/exec"

with main:
    def sanitize_datetime_state():
        from datetime import datetime, date, time
        for k in ["data_pericia_input", "data_atendimento_input"]:
            if k in st.session_state and isinstance(st.session_state[k], str):
                try:
                    st.session_state[k] = datetime.strptime(
                        st.session_state[k], "%Y-%m-%d").date()
                except:
                    st.session_state[k] = date.today()
        for k in ["horario", "horario_atendimento"]:
            if k in st.session_state and isinstance(st.session_state[k], str):
                try:
                    s = st.session_state[k]
                    if len(s.split(':')) == 2:
                        s += ":00"
                    st.session_state[k] = datetime.strptime(
                        s, "%H:%M:%S").time()
                except:
                    st.session_state[k] = None

    sanitize_datetime_state()


    logo_b64 = get_logo_base64()
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:65px; width:auto; object-fit:contain;" />' if logo_b64 else ''

    st.markdown(f'''
    <div class="app-header">
      <div class="app-header-brand">
        {logo_html}
        <div>
          <div class="app-header-title">Laudo de Local de Morte Violenta</div>
          <div class="app-header-sub">Instituto de Criminalística de Imperatriz &nbsp;|&nbsp; Perícia Oficial do Maranhão</div>
        </div>
      </div>
    </div>
    <div style="height:16px"></div>
    ''', unsafe_allow_html=True)


    @st.dialog("☁️ Carregar Ocorrência (Banco Local / Nuvem)")
    def modal_ocorrencias():
        st.markdown(
            "Busca ocorrências no Google Drive (App Mobile) ou banco local.")
        import sqlite3
        import json
        import os
        import requests
        DB_PATH = os.path.join(os.path.dirname(
            __file__), "banco_laudos.sqlite")

        # 1. Fetch from Cloud
        ocorrencias_cloud = []
        try:
            resp = requests.get(
                GAS_URL, params={"key": "perito:icrim123"}, timeout=5)
            if resp.status_code == 200:
                ocorrencias_cloud = resp.json()
        except:
            pass

        if ocorrencias_cloud and isinstance(ocorrencias_cloud, list):
            for item in ocorrencias_cloud:
                try:
                    conn = sqlite3.connect(DB_PATH)
                    c = conn.cursor()
                    mob_id = item.get("mobile_id", "")
                    dados_j = item.get("dados_json", "{}")
                    dt = item.get("data_sincronizacao", "")
                    if isinstance(dados_j, dict):
                        dados_j = json.dumps(dados_j)
                    c.execute('INSERT OR REPLACE INTO laudos (mobile_id, data_sincronizacao, dados_json) VALUES (?, ?, ?)',
                              (mob_id, dt, dados_j))
                    conn.commit()
                    conn.close()
                except:
                    pass

        # 2. Show from SQLite
        db_laudos = []
        if os.path.exists(DB_PATH):
            try:
                conn = sqlite3.connect(DB_PATH)
                conn.row_factory = sqlite3.Row
                c = conn.cursor()
                c.execute(
                    'SELECT id, mobile_id, data_sincronizacao, dados_json FROM laudos ORDER BY id DESC')
                rows = c.fetchall()
                conn.close()
                for r in rows:
                    db_laudos.append({"id": r["id"], "mobile_id": r["mobile_id"],
                                     "data": r["data_sincronizacao"], "dados": json.loads(r["dados_json"])})
            except:
                pass

        if db_laudos:
            filtro = st.text_input(
                "🔍 Buscar por perito, ocorrência ou data:",
                placeholder="Digite para filtrar por perito, ocorrência ou data...",
                key="filtro_ocorrencias_modal"
            ).strip().lower()

            db_laudos_filtrados = []
            for l in db_laudos:
                oc_val = str(l['dados'].get('ocorrencia', '')).lower()
                per_val = str(l['dados'].get('perito', '')).lower()
                dt_val = str(l.get('data', '')).lower()
                if not filtro or (filtro in oc_val or filtro in per_val or filtro in dt_val):
                    db_laudos_filtrados.append(l)

            if db_laudos_filtrados:
                opcoes = {
                    f"Ocorrência {l['dados'].get('ocorrencia', 'S/N')} — {l['dados'].get('perito', 'N/I')} ({l['data'][:16]})": l
                    for l in db_laudos_filtrados
                }
                escolhida = st.selectbox(
                    f"Resultados ({len(db_laudos_filtrados)} encontrada(s)):", list(opcoes.keys()))
                if st.button("Carregar Dados", use_container_width=True, type="primary"):
                    item = opcoes[escolhida]
                    data_obj = item["dados"]
                    for k, v in data_obj.items():
                        if k not in ["vitimas", "vestigios", "fotos", "quesitos_list"]:
                            st.session_state[k] = v
                    if "vitimas" in data_obj:
                        st.session_state["vitimas"] = data_obj["vitimas"]
                    if "vestigios" in data_obj:
                        st.session_state["vestigios"] = data_obj["vestigios"]
                    if "fotos" in data_obj:
                        st.session_state["fotos"] = data_obj["fotos"]
                    if "quesitos_list" in data_obj:
                        st.session_state["quesitos_list"] = data_obj["quesitos_list"]
                    st.rerun()
            else:
                st.warning("⚠️ Nenhuma ocorrência encontrada para o filtro informado.")
        else:
            st.info("Nenhuma ocorrência encontrada no banco.")





    @st.dialog("🩸 Guia Visual de Manchas de Sangue")
    def modal_guia_sangue():
        st.markdown("Utilize as imagens de referência abaixo para guiar a identificação do tipo de mancha. Passe o mouse sobre a imagem e use o ícone ⛶ (tela cheia) ou selecione um padrão abaixo para ampliar e baixar:")
        img_dir = os.path.join(os.path.dirname(__file__), "vestigio sangue")
        if os.path.exists(img_dir):
            files = sorted([os.path.join(img_dir, f) for f in os.listdir(img_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg'))])
            if files:
                cols = st.columns(2)
                for idx, img_p in enumerate(files):
                    with cols[idx % 2]:
                        st.image(img_p, use_container_width=True, caption=f"Padrão Referência #{idx+1}")

                st.markdown("<hr style='margin:16px 0; border-color:#e2e8f0;'>", unsafe_allow_html=True)
                st.markdown("##### 🔍 Ampliar e Baixar Padrão Específico")
                opcoes_img = {f"Padrão Referência #{i+1}": img_p for i, img_p in enumerate(files)}
                sel = st.selectbox("Escolha o padrão para ampliar:", list(opcoes_img.keys()), key="sel_zoom_sangue")
                if sel:
                    st.image(opcoes_img[sel], use_container_width=True, caption=f"Visualização Detalhada Ampliada — {sel}")
                    try:
                        with open(opcoes_img[sel], "rb") as file_img:
                            st.download_button(
                                label=f"📥 Baixar Imagem do {sel}",
                                data=file_img.read(),
                                file_name=os.path.basename(opcoes_img[sel]),
                                mime="image/jpeg",
                                use_container_width=True,
                                key=f"btn_dl_img_{sel}"
                            )
                    except Exception as e_dl:
                        st.error(f"Erro ao disponibilizar download: {e_dl}")
        else:
            st.info("Pasta 'vestigio sangue' não localizada.")


    gerar_top = render_action_buttons("top")
    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)

    with st.container():
        # ══ SEÇÃO 1: DA OCORRÊNCIA ════════════════════════════
        st.markdown(
            '<div class="section-title">📋 &nbsp; Da Ocorrência</div>', unsafe_allow_html=True)
        with st.container():

            col_tools1, col_tools2 = st.columns(2)
            with col_tools1:
                with st.expander("📝 Bloco de Notas / Rascunho Livre (IA)", expanded=False):
                    st.markdown("<p style='font-size:13px; color:#475569;'>Cole ou digite suas anotações brutas de campo (anotações do celular, rascunhos, transcrições) e a IA Gemini preencherá o laudo automaticamente.</p>", unsafe_allow_html=True)
                    raw_notes = st.text_area("Anotações de Campo do Perito", placeholder="Ex: ocorrencia 1234, perito carlos, local rua das flores 50, corpo masculino de brucos...", height=120, key="txt_raw_field_notes")
                    if st.button("🤖 Processar Rascunho com IA", type="primary", key="btn_parse_raw_notes", use_container_width=True):
                        if not raw_notes.strip():
                            st.warning("⚠️ Insira o texto das anotações antes de processar.")
                        else:
                            with st.spinner("Analisando rascunho com IA Gemini..."):
                                dados_parsed, err_p = ler_anotacoes_livres_gemini(raw_notes)
                                if err_p:
                                    st.error(f"Erro ao processar anotações: {err_p}")
                                elif dados_parsed:
                                    count_fields = 0
                                    mapping = {
                                        "num_laudo": "num_laudo", "ocorrencia": "ocorrencia", "requisicao": "requisicao",
                                        "perito": "perito", "autoridade_local": "autoridade_local", "endereco": "endereco",
                                        "ponto_referencia": "ponto_referencia", "area": "area", "delimitacoes": "delimitacoes",
                                        "equipe_pm": "equipe_pm", "equipe_pc": "equipe_pc", "dinamica_fatos": "dinamica_fatos"
                                    }
                                    for k_json, k_state in mapping.items():
                                        if dados_parsed.get(k_json):
                                            st.session_state[k_state] = dados_parsed[k_json]
                                            count_fields += 1

                                    if dados_parsed.get("vitima_nome") or dados_parsed.get("vitima_vestes") or dados_parsed.get("vitima_lesoes"):
                                        if "vitimas" in st.session_state and st.session_state["vitimas"]:
                                            v0 = st.session_state["vitimas"][0]
                                            if dados_parsed.get("vitima_nome"): v0["nome"] = dados_parsed["vitima_nome"]
                                            if dados_parsed.get("vitima_vestes"): v0["vestes"] = dados_parsed["vitima_vestes"]
                                            if dados_parsed.get("vitima_posicao"): v0["posicao"] = dados_parsed["vitima_posicao"]
                                            if dados_parsed.get("vitima_lesoes"): v0["lesoes"] = [dados_parsed["vitima_lesoes"]]
                                            count_fields += 1

                                    st.success(f"✅ Sucesso! {count_fields} campo(s) preenchido(s):")
                                    st.json(dados_parsed)
                                    st.rerun()

            with col_tools2:
                with st.expander("📄 Leitor de Requisição (PDF / Imagem)", expanded=False):
                    if get_gemini_api_key():
                        st.info("✨ Gemini Vision: Extração inteligente de Delegacia, Delegado, Ocorrência, Requisição e Quesitos.")
                    else:
                        st.caption("💡 Para extração com IA Gemini Vision, verifique o arquivo CHAVE.txt.")

                    req_file = st.file_uploader("Carregar Requisição (PDF / Imagem)", type=["pdf", "png", "jpg", "jpeg", "bmp", "tiff"], key="req_file_input")
                    if req_file is not None:
                        file_bytes = req_file.read()
                        extracted_text, dados = ler_requisicao_pericial(file_bytes, req_file.name, req_file.type)
                        if dados:
                            if dados.get("delegacia"):
                                st.session_state["destino"] = dados["delegacia"]
                            if dados.get("delegado"):
                                del_nome = dados["delegado"]
                                if del_nome not in st.session_state.autoridades:
                                    st.session_state.autoridades.append(del_nome)
                                st.session_state["autoridade_sel"] = del_nome
                            if dados.get("ocorrencia"):
                                st.session_state["ocorrencia"] = dados["ocorrencia"]
                            if dados.get("requisicao"):
                                st.session_state["requisicao"] = dados["requisicao"]
                            if dados.get("quesitos"):
                                st.session_state["quesitos"] = dados["quesitos"]
                                raw_lines = [q.strip() for q in dados["quesitos"].split("\n") if q.strip()]
                                if raw_lines:
                                    st.session_state["quesitos_list"] = [{"pergunta": q_line, "resposta": ""} for q_line in raw_lines]

                            orig = dados.get("origem", "Automático")
                            st.success(f"✅ Requisição lida ({orig}):")
                            st.json(dados)
                        elif extracted_text:
                            st.warning("⚠️ Texto extraído da requisição, mas nenhum campo de formulário reconhecido por regex.")
                            with st.expander("Ver Texto Extraído"):
                                st.text(extracted_text)
                        else:
                            st.error("❌ Não foi possível extrair texto do arquivo fornecido.")

            c1, c2, c3, c4 = st.columns(4)
            num_laudo = c1.text_input(
                "Nº do Laudo",         placeholder="0092258",    key="num_laudo")
            ocorrencia = c2.text_input(
                "Nº da Ocorrência",    placeholder="12345/2026", key="ocorrencia")
            requisicao = c3.text_input(
                "Requisição",           key="requisicao")
            referencia = c4.text_input(
                "IP / PM / BO / Proc.", key="referencia")

            c5, c6, c7, c8 = st.columns(4)
            data_pericia_val = c5.date_input(
                "Data da Perícia",        value=date.today(), format="DD/MM/YYYY", key="data_pericia_input")
            horario_val = c6.time_input(
                "Horário da Perícia",     key="horario",      step=300)
            data_atendimento_val = c7.date_input(
                "Data de Atendimento",    value=date.today(), format="DD/MM/YYYY", key="data_atendimento_input")
            horario_atend_val = c8.time_input(
                "Horário de Atendimento", key="horario_atendimento", step=300)

            cp, ca, cb = st.columns([2, 2, 0.7])
            perito = cp.text_input("Perito(a) Relator(a)", key="perito")
            with ca:
                autoridade = st.selectbox("Autoridade Solicitante",
                                          options=[""] + st.session_state.autoridades, key="autoridade_sel")
            with cb:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Nova", key="btn_nova_aut", use_container_width=True):
                    st.session_state.show_nova_aut = True

        if st.session_state.show_nova_aut:
            na1, na2, na3 = st.columns([3, 1, 1])
            nova_aut = na1.text_input("Nome da autoridade:", key="nova_aut_input",
                                      label_visibility="collapsed",
                                      placeholder="Ex: Delegado(a) Titular Silva")
            if na2.button("✓ Salvar", type="primary", key="btn_ok_aut", use_container_width=True):
                if nova_aut.strip() and nova_aut.strip() not in st.session_state.autoridades:
                    st.session_state.autoridades.append(nova_aut.strip())
                st.session_state.show_nova_aut = False
                st.rerun()
            if na3.button("Cancelar", key="btn_cancel_aut", use_container_width=True):
                st.session_state.show_nova_aut = False
                st.rerun()

        cv1, cv2 = st.columns(2)
        investigado = cv1.text_area(
            "Investigado(s)",  placeholder="Nomes dos investigados...", key="investigado", height=80)
        destino = cv2.text_input(
            "Unidade Destino", value="Delegacia de Homicídios de Imperatriz", key="destino")

    with st.container():
        # ══ SEÇÃO 2: DO LOCAL ════════════════════════════════
        st.markdown(
            '<div class="section-title">📍 &nbsp; Do Local</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>',
                        unsafe_allow_html=True)

            ce1, ce2, ce3, ce4 = st.columns([3, 1.2, 1.2, 0.6])
            endereco = ce1.text_input(
                "Endereço completo", placeholder="Rua, Nº, Bairro, Município", key="endereco")
            latitude = ce2.text_input(
                "Latitude",  placeholder="-5.518600",  key="latitude")
            longitude = ce3.text_input(
                "Longitude", placeholder="-47.477600", key="longitude")
            with ce4:
                st.markdown("<br>", unsafe_allow_html=True)
                st.button("📍 GPS", key="btn_gps", use_container_width=True,
                          help="GPS disponível via HTTPS")

            cr1, cr2, cr3 = st.columns([2, 1, 1])
            ponto_referencia = cr1.text_input(
                "Ponto de Referência", placeholder="Ex: Próximo ao posto de saúde", key="ponto_referencia")
            area = cr2.text_input(
                "Área",      placeholder="Ex: Externa", key="area")
            pavimento = cr3.text_input(
                "Pavimento", placeholder="Ex: Asfalto", key="pavimento")

            cdel1, cdel2 = st.columns([4, 1])
            with cdel1:
                delimitacoes = st.text_area("Delimitações do local",
                                            placeholder="Descreva os limites físicos do local periciado...",
                                            key="delimitacoes", height=70)
            with cdel2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("✨ Polir IA", key="btn_polir_delimitacoes", use_container_width=True, help="Polir redação das delimitações com IA"):
                    modal_polir_redacao("delimitacoes", "Delimitações do Local")

            cl1, cl2, cl3, cl4 = st.columns(4)
            iso_opts = ["", "Preservado e Isolado", "Parcialmente Preservado e Isolado",
                        "Preservado e Parcialmente Isolado", "Não Preservado e Não Isolado", "Outros"]
            isolamento_sel = cl1.selectbox(
                "Isolamento", iso_opts, key="isolamento")

            clima_opts = ["", "Aberto diurno", "Aberto noturno", "Nublado diurno",
                          "Nublado noturno", "Chuvoso diurno", "Chuvoso noturno", "Outros"]
            clima_sel = cl2.selectbox("Clima", clima_opts, key="clima")

            visibilidade = cl3.selectbox(
                "Visibilidade", ["", "Ampla", "Reduzida"], key="visibilidade")
            iluminacao = cl4.selectbox("Iluminação",
                                       ["", "Natural Satisfatória", "Natural Insatisfatória",
                                        "Artificial Satisfatória", "Artificial Insatisfatória"], key="iluminacao")

            if isolamento_sel == "Outros" or clima_sel == "Outros":
                ot1, ot2 = st.columns(2)
                if isolamento_sel == "Outros":
                    isolamento_txt = ot1.text_input(
                        "Especifique o Isolamento", key="isolamento_outros")
                    isolamento = isolamento_txt or "Outros"
                else:
                    isolamento = isolamento_sel
                if clima_sel == "Outros":
                    clima_txt = ot2.text_input(
                        "Especifique o Clima", key="clima_outros")
                    clima = clima_txt or "Outros"
                else:
                    clima = clima_sel
            else:
                isolamento = isolamento_sel
                clima = clima_sel

            cq1, cq2, cq3 = st.columns(3)
            equipe_pm = cq1.text_input(
                "Equipe PM", placeholder="VTR, Comandante...", key="equipe_pm")
            equipe_pc = cq2.text_input(
                "Equipe PC", placeholder="Investigador, Delegado...", key="equipe_pc")
            autoridade_local = cq3.text_input(
                "Autoridade no Local", placeholder="Nome da autoridade...", key="autoridade_local")

    with st.container():
        # ══ SEÇÃO: DAS VÍTIMAS (Item 3.3) ═════════════════════
        st.markdown(
            '<div class="section-title">👥 &nbsp; Das Vítimas (Item 3.3)</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>',
                        unsafe_allow_html=True)

            with st.expander("📄 Leitor Gemini Vision - Laudo de Necropsia / Exame Cadavérico (PDF / Imagem)", expanded=False):
                st.markdown("<p style='font-size:13px; color:#475569;'>Faça o upload do Laudo de Necropsia ou Auto de Exame Cadavérico (IML) em PDF ou imagem para extração automática da Causa Mortis, Instrumento Lesivo e Lesões Detalhadas com Gemini Vision IA.</p>", unsafe_allow_html=True)
                nec_file = st.file_uploader("Carregar Laudo de Necropsia / Exame Cadavérico", type=["pdf", "png", "jpg", "jpeg", "bmp", "tiff"], key="necropsia_file_input")
                if nec_file is not None:
                    if st.button("✨ Processar Necropsia com Gemini Vision", type="primary", key="btn_process_necropsia", use_container_width=True):
                        if not get_gemini_api_key():
                            st.error("⚠️ Chave GEMINI_API_KEY não configurada. Verifique se o arquivo CHAVE.txt está presente.")
                        else:
                            with st.spinner("Analisando Laudo de Necropsia com Gemini Vision..."):
                                nec_bytes = nec_file.read()
                                dados_nec, err_nec = ler_necropsia_gemini_vision(nec_bytes, nec_file.name, nec_file.type)
                                if err_nec:
                                    st.error(f"❌ Erro ao analisar laudo de necropsia: {err_nec}")
                                elif dados_nec:
                                    if dados_nec.get("causa_mortis"):
                                        st.session_state["resultado_laudo_IML"] = dados_nec["causa_mortis"]
                                    if dados_nec.get("numero_laudo_iml"):
                                        st.session_state["numero_laudo_necropsia"] = dados_nec["numero_laudo_iml"]

                                    inst_extraido = dados_nec.get("instrumento_lesivo", "")
                                    if inst_extraido:
                                        map_inst = {
                                            "perfurocontundente": "1. Perfurocontundente (Arma de Fogo)",
                                            "cortante": "2. Cortante (Feridas Incisas - Faca, Navalha, Estilete, Vidro)",
                                            "perfurante": "3. Perfurante (Feridas Punctórias - Espeto, Estilete, Chave de Fenda)",
                                            "perfurocortante": "4. Perfurocortante (Feridas Perfuroincisas - Faca, Punhal, Canivete)",
                                            "contundente": "5. Contundente (Feridas Contusas/Fraturas - Madeira, Pedra, Veículo, Piso)",
                                            "cortocontundente": "6. Cortocontundente (Feridas Contuso-Incisas - Machado, Facão, Foice)",
                                            "térmica": "7. Ação Térmica (Queimaduras - Chama Direta, Líquido Fervente, Superfície Aquecida)",
                                            "termica": "7. Ação Térmica (Queimaduras - Chama Direta, Líquido Fervente, Superfície Aquecida)"
                                        }
                                        for k_m, v_m in map_inst.items():
                                            if k_m in inst_extraido.lower():
                                                st.session_state["inst_acao"] = v_m
                                                break
                                    if dados_nec.get("agente_instrumento"):
                                        st.session_state["inst_agente"] = dados_nec["agente_instrumento"]

                                    if "vitimas" in st.session_state and st.session_state["vitimas"]:
                                        vit0 = st.session_state["vitimas"][0]
                                        if dados_nec.get("nome_vitima") and not vit0.get("nome"):
                                            vit0["nome"] = dados_nec["nome_vitima"]
                                        if dados_nec.get("documento_vitima") and not vit0.get("documento"):
                                            vit0["documento"] = dados_nec["documento_vitima"]
                                        if dados_nec.get("lesoes") and isinstance(dados_nec["lesoes"], list) and len(dados_nec["lesoes"]) > 0:
                                            vit0["lesoes"] = [str(l).strip() for l in dados_nec["lesoes"] if str(l).strip()]

                                    st.success("✅ Laudo de Necropsia analisado com sucesso pelo Gemini Vision! Dados preenchidos no formulário:")
                                    st.json(dados_nec)
                                    st.rerun()

            vitimas_list = st.session_state.get("vitimas", [{"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "",
                                                "naturalidade": "", "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "", "fenomenos": "", "lesoes": [""]}])

            for i, vit in enumerate(vitimas_list):
                st.markdown(f"##### 👤 Vítima #{i+1}")

                st.markdown(
                    "<p style='font-size:12px; font-weight:700; color:#1a3a6e; margin-bottom:8px;'>IDENTIFICAÇÃO</p>", unsafe_allow_html=True)
                c_nome, c_doc, c_sexo = st.columns([2, 1, 1])
                vit["nome"] = c_nome.text_input(
                    "Nome completo", value=vit.get("nome", ""), key=f"vit_nome_{i}")
                vit["documento"] = c_doc.text_input("Nº Documento (RG/CPF)", value=vit.get(
                    "documento", ""), key=f"vit_doc_{i}", placeholder="RG ou CPF...")
                vit["sexo"] = c_sexo.selectbox("Sexo", ["", "Masculino", "Feminino", "Outro"], index=[
                                               "", "Masculino", "Feminino", "Outro"].index(vit.get("sexo", "")), key=f"vit_sexo_{i}")

                c_dob, c_fili, c_nat = st.columns([1, 2, 1])
                dob_val = vit.get("data_nascimento", None)
                if not isinstance(dob_val, date) and dob_val is not None:
                    dob_val = date.today()
                vit["data_nascimento"] = c_dob.date_input(
                    "Data de Nascimento", value=dob_val, key=f"vit_dob_{i}", format="DD/MM/YYYY")
                vit["filicao"] = c_fili.text_input("Filiação", value=vit.get(
                    "filicao", ""), key=f"vit_fili_{i}", placeholder="Nome da mãe e/ou pai...")
                vit["naturalidade"] = c_nat.text_input("Naturalidade", value=vit.get(
                    "naturalidade", ""), key=f"vit_nat_{i}", placeholder="Cidade-UF...")

                st.markdown(
                    "<p style='font-size:12px; font-weight:700; color:#1a3a6e; margin:12px 0 8px 0;'>VESTES E PERTENCES</p>", unsafe_allow_html=True)
                c_vest, c_pert = st.columns(2)
                vit["vestes"] = c_vest.text_area("Vestes", value=vit.get(
                    "vestes", ""), key=f"vit_vest_{i}", height=70, placeholder="Descrição das roupas que usava...")
                vit["pertences"] = c_pert.text_area("Pertences", value=vit.get(
                    "pertences", ""), key=f"vit_pert_{i}", height=70, placeholder="Celular, chaves, carteira, joias, etc...")

                st.markdown(
                    "<p style='font-size:12px; font-weight:700; color:#1a3a6e; margin:12px 0 8px 0;'>POSIÇÃO E LOCALIZAÇÃO (IN SITU)</p>", unsafe_allow_html=True)
                c_loc, c_pos = st.columns(2)
                vit["localizacao"] = c_loc.text_area("Localização", value=vit.get(
                    "localizacao", ""), key=f"vit_loc_{i}", height=70, placeholder="Ex: No compartimento traseiro do carro...")
                vit["posicao"] = c_pos.text_area("Posição", value=vit.get(
                    "posicao", ""), key=f"vit_pos_{i}", height=70, placeholder="Ex: Decúbito lateral direito...")

                c_cab, c_memb = st.columns(2)
                vit["cabeca"] = c_cab.text_input("Cabeça", value=vit.get(
                    "cabeca", ""), key=f"vit_cab_{i}", placeholder="Ex: Voltada a leste...")
                vit["membros"] = c_memb.text_input("Membros", value=vit.get(
                    "membros", ""), key=f"vit_memb_{i}", placeholder="Ex: Membros superiores fletidos...")

                st.markdown(
                    "<p style='font-size:12px; font-weight:700; color:#1a3a6e; margin:12px 0 8px 0;'>EXAME PERINECROSCÓPICO E ESTADO DE CONSERVAÇÃO</p>", unsafe_allow_html=True)
                vit["fenomenos"] = st.text_input("Fenômenos Transformativos", value=vit.get(
                    "fenomenos", ""), key=f"vit_fen_{i}", placeholder="Ex: Mancha verde abdominal, rigidez muscular...")

                st.markdown(
                    "<p style='font-size:11px; font-weight:600; color:#64748b; margin-top:8px;'>LESÕES CONSTATADAS</p>", unsafe_allow_html=True)
                les_list = vit.get("lesoes", [""])
                for j, lesao in enumerate(les_list):
                    c_les, c_rem_l = st.columns([12, 1])
                    les_list[j] = c_les.text_input(
                        f"Lesão #{j+1}", value=lesao, key=f"vit_{i}_les_{j}", label_visibility="collapsed", placeholder=f"Descrição da lesão #{j+1}...")
                    if len(les_list) > 1:
                        if c_rem_l.button("❌", key=f"vit_{i}_rem_les_{j}", help="Remover esta lesão"):
                            les_list.pop(j)
                            st.rerun()
                vit["lesoes"] = les_list

                if st.button("➕ Adicionar Lesão", key=f"vit_{i}_add_les"):
                    vit.setdefault("lesoes", [""]).append("")
                    st.rerun()

                if len(vitimas_list) > 1:
                    st.markdown("<div style='height:8px'></div>",
                                unsafe_allow_html=True)
                    if st.button(f"🗑️ Remover Vítima #{i+1}", key=f"remove_vit_{i}", type="secondary"):
                        st.session_state.vitimas.pop(i)
                        st.rerun()
                st.markdown(
                    "<hr style='margin:20px 0; border-color:#2563eb;'>", unsafe_allow_html=True)

            if st.button("➕ Adicionar Outra Vítima", key="btn_add_vitima"):
                st.session_state.vitimas.append({"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "", "naturalidade": "",
                                                "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "", "fenomenos": "", "lesoes": [""]})
                st.rerun()

    with st.container():
        st.markdown(
            '<div class="section-title">🔍 &nbsp; Dos Vestígios</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>',
                        unsafe_allow_html=True)

            manchas_dict = {
                "1. FORMAÇÃO PASSIVA": ["Gotejamento Isolado", "Gotejamento Sucessivo", "Gotejamento Estático", "Gotejamento Dinâmico", "Escorrimento", "Saturação", "Empoçamento"],
                "2. FORMAÇÃO ATIVA": ["Espargimento Por Impacto", "Espargimento Por Expiração", "Espargimento Por Êmese", "Espargimento De Retorno", "Espargimento De Saída", "Transferência Por Contato", "Transferência Por Arrastamento", "Inercial Por Desprendimento Centrífugo", "Inercial Por Parada", "Projeção por Diferença de Pressão", "Espalhamento de Volume em Queda Livre"],
                "3. ALTERADAS": ["Por Arrastamento", "Silhueta de Mancha", "Por Limpeza", "Por Inseto"],
                "4. AUSÊNCIA": ["Área de Sombra", "Sangue Latente"]
            }

            if "vestigios" not in st.session_state or not st.session_state.vestigios:
                st.session_state["vestigios"] = [{"tipo": "", "categoria": "", "subtipo": "", "localizacao": "", "descricao": "", "elemento_tipo": "",
                                                  "quantidade": "", "envelope": "", "arma_tipo": "", "municiada": "", "marca": "", "cor": "", "funcionamento": "", "instrumento_tipo": ""}]
            vestigios_list = st.session_state["vestigios"]

            for i, vest in enumerate(vestigios_list):
                st.markdown(f"##### 🔎 Vestígio #{i+1}")
                tipo_opts = ["", "Manchas de Sangue", "Elemento Balístico", "Arma de Fogo",
                             "Aparelho Telefônico", "Instrumento Lesivo", "Impressão Datiloscópica", "Outros"]
                tipo_idx = tipo_opts.index(vest.get("tipo", "")) if vest.get(
                    "tipo", "") in tipo_opts else 0
                vest["tipo"] = st.selectbox(
                    "Tipo de Vestígio", tipo_opts, index=tipo_idx, key=f"vest_tipo_{i}")

                if vest["tipo"] in ["Manchas de Sangue", "Mancha de Sangue"]:
                    if st.button("🩸 Abrir Guia de Manchas de Sangue", key=f"btn_guia_sangue_{i}"):
                        modal_guia_sangue()
                    c_cat, c_sub = st.columns(2)
                    cat_opts = [""] + list(manchas_dict.keys())
                    cat_idx = cat_opts.index(vest.get("categoria", "")) if vest.get(
                        "categoria", "") in cat_opts else 0
                    vest["categoria"] = c_cat.selectbox(
                        "Categoria", cat_opts, index=cat_idx, key=f"vest_cat_{i}")
                    sub_opts = [""] + \
                        manchas_dict.get(vest.get("categoria", ""), [])
                    sub_idx = sub_opts.index(vest.get("subtipo", "")) if vest.get(
                        "subtipo", "") in sub_opts else 0
                    vest["subtipo"] = c_sub.selectbox(
                        "Subtipo", sub_opts, index=sub_idx, key=f"vest_sub_{i}")
                    if vest["subtipo"]:
                        image_path = os.path.join(os.path.dirname(
                            __file__), "imagens_manchas", f"{vest['subtipo']}.jpg")
                        if os.path.exists(image_path):
                            st.image(image_path, width=300,
                                     caption=f"Guia Visual: {vest['subtipo']}")
                        else:
                            st.caption(
                                f"(Dica: Salve o recorte como 'imagens_manchas\\{vest['subtipo']}.jpg' para ver o guia visual aqui)")

                    vest["localizacao"] = st.text_input("Localização", value=vest.get(
                        "localizacao", ""), key=f"vest_loc_{i}", placeholder="Ex: Parede leste...")
                    vest["descricao"] = st.text_area("Descrição Livre", value=vest.get(
                        "descricao", ""), key=f"vest_desc_{i}", height=70)

                elif vest["tipo"] == "Elemento Balístico":
                    c_tipo, c_qtd = st.columns([3, 1])
                    elem_opts = ["", "Estojo",
                                 "Munição", "Jaqueta", "Projétil"]
                    elem_idx = elem_opts.index(vest.get("elemento_tipo", "")) if vest.get(
                        "elemento_tipo", "") in elem_opts else 0
                    vest["elemento_tipo"] = c_tipo.selectbox(
                        "Tipo de Elemento", elem_opts, index=elem_idx, key=f"vest_elem_{i}")
                    vest["quantidade"] = c_qtd.text_input(
                        "Quantidade", value=vest.get("quantidade", ""), key=f"vest_qtd_{i}")
                    vest["envelope"] = st.text_input("Número Envelope de Segurança", value=vest.get(
                        "envelope", ""), key=f"vest_env_{i}")
                    vest["descricao"] = st.text_area("Descrição (geral)", value=vest.get(
                        "descricao", ""), key=f"vest_desc_{i}", height=70)

                elif vest["tipo"] == "Arma de Fogo":
                    c_arma, c_mun = st.columns([3, 1])
                    arma_opts = ["", "Revólver", "Espingarda",
                                 "Rifle", "Pistola", "Caseira"]
                    arma_idx = arma_opts.index(vest.get("arma_tipo", "")) if vest.get(
                        "arma_tipo", "") in arma_opts else 0
                    vest["arma_tipo"] = c_arma.selectbox(
                        "Tipo de Arma", arma_opts, index=arma_idx, key=f"vest_arma_{i}")
                    mun_opts = ["", "Sim", "Não"]
                    mun_idx = mun_opts.index(vest.get("municiada", "")) if vest.get(
                        "municiada", "") in mun_opts else 0
                    vest["municiada"] = c_mun.selectbox(
                        "Municiada", mun_opts, index=mun_idx, key=f"vest_mun_{i}")
                    vest["envelope"] = st.text_input("Número Envelope de Segurança", value=vest.get(
                        "envelope", ""), key=f"vest_env_{i}")
                    vest["descricao"] = st.text_area("Descrição", value=vest.get(
                        "descricao", ""), key=f"vest_desc_{i}", height=70)

                elif vest["tipo"] == "Outros":
                    vest["descricao"] = st.text_area("Descrição", value=vest.get(
                        "descricao", ""), key=f"vest_desc_{i}", height=70, placeholder="Descreva o vestígio...")

                elif vest["tipo"] == "Impressão Datiloscópica":
                    vest["descricao"] = st.text_area("Descrição", value=vest.get(
                        "descricao", ""), key=f"vest_desc_{i}", height=70, placeholder="Descreva os detalhes da impressão datiloscópica, superfície, revelação, etc.")

                elif vest["tipo"] == "Instrumento Lesivo":
                    vest["instrumento_tipo"] = st.text_input("Tipo (ex: Faca, pedaço de madeira, corda)", value=vest.get(
                        "instrumento_tipo", ""), key=f"vest_inst_{i}")
                    vest["envelope"] = st.text_input("Número Envelope de Segurança", value=vest.get(
                        "envelope", ""), key=f"vest_env_{i}")
                    vest["descricao"] = st.text_area("Descrição", value=vest.get(
                        "descricao", ""), key=f"vest_desc_{i}", height=70)

                elif vest["tipo"] == "Aparelho Telefônico":
                    c_marca, c_cor = st.columns(2)
                    vest["marca"] = c_marca.text_input(
                        "Marca ou Modelo", value=vest.get("marca", ""), key=f"vest_marca_{i}")
                    vest["cor"] = c_cor.text_input(
                        "Cor", value=vest.get("cor", ""), key=f"vest_cor_{i}")
                    c_func, c_env = st.columns(2)
                    func_opts = ["", "Ligado", "Desligado"]
                    func_idx = func_opts.index(vest.get("funcionamento", "")) if vest.get(
                        "funcionamento", "") in func_opts else 0
                    vest["funcionamento"] = c_func.selectbox(
                        "Funcionamento", func_opts, index=func_idx, key=f"vest_func_{i}")
                    vest["envelope"] = c_env.text_input(
                        "Número Envelope de Segurança", value=vest.get("envelope", ""), key=f"vest_env_{i}")
                    vest["descricao"] = st.text_area("Descrição", value=vest.get(
                        "descricao", ""), key=f"vest_desc_{i}", height=70)

                if st.button(f"🗑️ Remover Vestígio #{i+1}", key=f"remove_vest_{i}", type="secondary"):
                    st.session_state.vestigios.pop(i)
                    st.rerun()
                st.markdown(
                    "<hr style='margin:20px 0; border-color:#2563eb;'>", unsafe_allow_html=True)

            if st.button("➕ Adicionar Vestígio", key="btn_add_vest"):
                if "vestigios" not in st.session_state:
                    st.session_state.vestigios = []
                st.session_state.vestigios.append({"tipo": "", "categoria": "", "subtipo": "", "localizacao": "", "descricao": "", "elemento_tipo": "",
                                                  "quantidade": "", "envelope": "", "arma_tipo": "", "municiada": "", "marca": "", "cor": "", "funcionamento": "", "instrumento_tipo": ""})
                st.rerun()

    # ══ BARRA DE AÇÕES (RODAPÉ) ══════════════════════════

    with st.container():
        st.markdown(
            '<div class="section-title">📝 &nbsp; 4. Considerações Técnicas</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>',
                        unsafe_allow_html=True)

            st.subheader("a) Isolamento e Preservação")

            opcoes_isolamento = [
                "1. Local Preservado e Isolado",
                "2. Local Parcialmente Preservado e Isolado",
                "3. Local Preservado e Parcialmente Isolado",
                "4. Local Parcialmente Preservado e Parcialmente Isolado",
                "5. Local Não Preservado e Não Isolado (Devassado)"
            ]

            iso_idx = 0
            curr_iso = st.session_state.get("iso_estado", "")
            if curr_iso in opcoes_isolamento:
                iso_idx = opcoes_isolamento.index(curr_iso)

            iso_estado = st.selectbox(
                "Estado do Isolamento", opcoes_isolamento, index=iso_idx, key="iso_estado")

            # Condicionais
            st.session_state["iso_meio"] = st.session_state.get("iso_meio", "")
            st.session_state["iso_alteracao"] = st.session_state.get(
                "iso_alteracao", "")
            st.session_state["iso_falha"] = st.session_state.get(
                "iso_falha", "")

            if iso_estado.startswith("1"):
                st.text_input("Meio de Isolamento", value=st.session_state["iso_meio"], key="iso_meio_ui",
                              help="Ex: fita zebrada, cordão de isolamento, guarnição física da PM")
                st.session_state["iso_meio"] = st.session_state.get(
                    "iso_meio_ui", "")
            elif iso_estado.startswith("2"):
                st.text_input(
                    "Meio de Isolamento", value=st.session_state["iso_meio"], key="iso_meio_ui")
                st.text_area("Descrever Alteração", value=st.session_state["iso_alteracao"],
                             key="iso_alteracao_ui", help="Ex: movimentação da vítima por equipes de socorro")
                st.session_state["iso_meio"] = st.session_state.get(
                    "iso_meio_ui", "")
                st.session_state["iso_alteracao"] = st.session_state.get(
                    "iso_alteracao_ui", "")
            elif iso_estado.startswith("3"):
                st.text_input("Descrever Falha no Isolamento",
                              value=st.session_state["iso_falha"], key="iso_falha_ui", help="Ex: ausência de barreiras em rotas de fuga")
                st.session_state["iso_falha"] = st.session_state.get(
                    "iso_falha_ui", "")
            elif iso_estado.startswith("4"):
                st.text_input("Descrever Falha no Isolamento",
                              value=st.session_state["iso_falha"], key="iso_falha_ui", help="Ex: coibir o fluxo de pessoas não autorizadas")
                st.text_area("Descrever Alteração", value=st.session_state["iso_alteracao"],
                             key="iso_alteracao_ui", help="Ex: marcas de pneus sobrepostas a manchas de sangue")
                st.session_state["iso_falha"] = st.session_state.get(
                    "iso_falha_ui", "")
                st.session_state["iso_alteracao"] = st.session_state.get(
                    "iso_alteracao_ui", "")
            elif iso_estado.startswith("5"):
                st.text_area("Descrever Alteração (Devassado)",
                             value=st.session_state["iso_alteracao"], key="iso_alteracao_ui")
                st.session_state["iso_alteracao"] = st.session_state.get(
                    "iso_alteracao_ui", "")

            st.divider()
            st.subheader("c) Laudo de Necropsia (IML)")
            c_iml1, c_iml2 = st.columns(2)
            with c_iml1:
                st.session_state["numero_laudo_necropsia"] = st.text_input("Número Laudo IML", value=st.session_state.get(
                    "numero_laudo_necropsia", ""), placeholder="Ex: 123/2026 - IML")
            with c_iml2:
                st.session_state["resultado_laudo_IML"] = st.text_input("Resultado do Laudo IML", value=st.session_state.get(
                    "resultado_laudo_IML", ""), placeholder="Ex: traumatismo cranioencefálico por PAF")

            st.divider()
            st.subheader("d) Instrumento Utilizado")

            opcoes_inst = [
                "1. Perfurocontundente (Arma de Fogo)",
                "2. Cortante (Feridas Incisas - Faca, Navalha, Estilete, Vidro)",
                "3. Perfurante (Feridas Punctórias - Espeto, Estilete, Chave de Fenda)",
                "4. Perfurocortante (Feridas Perfuroincisas - Faca, Punhal, Canivete)",
                "5. Contundente (Feridas Contusas/Fraturas - Madeira, Pedra, Veículo, Piso)",
                "6. Cortocontundente (Feridas Contuso-Incisas - Machado, Facão, Foice)",
                "7. Ação Térmica (Queimaduras - Chama Direta, Líquido Fervente, Superfície Aquecida)"
            ]

            inst_idx = 0
            curr_inst = st.session_state.get("inst_acao", "")
            if curr_inst in opcoes_inst:
                inst_idx = opcoes_inst.index(curr_inst)

            inst_acao = st.selectbox(
                "Tipo de Ação do Instrumento", opcoes_inst, index=inst_idx, key="inst_acao")

            sugestoes_agentes = {
                "1": ["projéteis de arma de fogo (PAF)"],
                "2": ["faca", "navalha", "fragmento de vidro", "lâmina de estilete"],
                "3": ["espeto", "estilete", "chave de fenda", "sovela"],
                "4": ["faca", "punhal", "canivete"],
                "5": ["barra de ferro", "pedaço de madeira", "pedra", "veículo automotor em movimento", "piso"],
                "6": ["machado", "facão", "foice", "enxada"],
                "7": ["chama direta", "líquido fervente", "superfície superaquecida"]
            }

            inst_key = inst_acao[0]
            if inst_key != "1":
                sug = sugestoes_agentes.get(inst_key, [])
                st.session_state["inst_agente"] = st.text_input("Agente Compatível", value=st.session_state.get(
                    "inst_agente", ""), placeholder=f"Ex: {', '.join(sug[:3])}")
            else:
                st.session_state["inst_agente"] = "projéteis de arma de fogo (PAF)"

            placeholders_extras = {
                "1": "Ex: com a recuperação de um projétil em cada cadáver durante a necrópsia / com orifícios de entrada compatíveis com disparos a curta distância",
                "2": "Ex: evidenciando cauda de escoriação que indica a direção do corte / apresentando características típicas de lesão de defesa",
                "3": "Ex: com transfixação de estruturas vitais profundas / sem a recuperação ou retenção de fragmentos do instrumento no interior da cavidade",
                "4": "Ex: contabilizando-se 3 golpes desferidos / com a recuperação de um fragmento metálico da lâmina retido na estrutura óssea",
                "5": "Ex: resultando em traumatismo cranioencefálico severo / não havendo elementos materiais desprendidos do instrumento no cadáver",
                "6": "Ex: provocando lacerações profundas associadas a fraturas ósseas",
                "7": "Ex: configurando queimaduras de 3º grau, com presença de fuligem em vias aéreas"
            }

            st.session_state["inst_extra"] = st.text_area(
                "Achados Extras / Observações do Instrumento", value=st.session_state.get("inst_extra", ""), help=placeholders_extras.get(inst_key, ""))

            st.divider()
            
            st.markdown("<hr style='margin:20px 0; border-color:#cbd5e1;'>", unsafe_allow_html=True)
            st.subheader("Considerações Técnicas Adicionais")
            st.markdown("<p style='font-size:13px; color:#475569; margin-bottom:12px;'>Adicione quantas considerações personalizadas desejar. A numeração por letras (e, f, g...) será gerada automaticamente em ordem crescente.</p>", unsafe_allow_html=True)

            if "consideracoes_extras" not in st.session_state:
                st.session_state["consideracoes_extras"] = []

            cons_extras = st.session_state["consideracoes_extras"]

            for idx_extra, item_extra in enumerate(cons_extras):
                letra = chr(ord('e') + idx_extra)
                st.markdown(f"##### {letra}) Consideração Adicional #{idx_extra + 1}")
                
                item_extra["titulo"] = st.text_input(
                    "Título / Assunto (Opcional)",
                    value=item_extra.get("titulo", ""),
                    key=f"extra_cons_tit_{idx_extra}",
                    placeholder="Ex: Do Exame das Vestes / Das Marcas de Pneu"
                )
                
                item_extra["texto"] = st.text_area(
                    "Texto da Consideração",
                    value=item_extra.get("texto", ""),
                    key=f"extra_cons_txt_{idx_extra}",
                    height=100,
                    placeholder="Digite o texto da consideração técnica..."
                )
                
                if st.button(f"🗑️ Remover Consideração ({letra})", key=f"btn_rem_extra_{idx_extra}", type="secondary"):
                    st.session_state["consideracoes_extras"].pop(idx_extra)
                    st.rerun()
                st.markdown("<hr style='margin:16px 0; border-color:#e2e8f0;'>", unsafe_allow_html=True)

            if st.button("➕ Adicionar Consideração", key="btn_add_consideracao_extra", type="primary"):
                st.session_state["consideracoes_extras"].append({"titulo": "", "texto": ""})
                st.rerun()

            st.divider()
            st.subheader("e) Quesitos e Respostas")
            st.markdown("<p style='font-size:13px; color:#475569; margin-bottom:12px;'>Gerenciamento de quesitos formulados pela Autoridade Solicitante e respostas da perícia.</p>", unsafe_allow_html=True)

            if "quesitos_list" not in st.session_state or not st.session_state["quesitos_list"]:
                st.session_state["quesitos_list"] = [{"pergunta": "", "resposta": ""}]

            quesitos_list = st.session_state["quesitos_list"]

            for i, q_item in enumerate(quesitos_list):
                st.markdown(f"##### ❓ Quesito #{i+1}")
                q_item["pergunta"] = st.text_input(
                    "Pergunta",
                    value=q_item.get("pergunta", ""),
                    key=f"quesito_perg_{i}",
                    placeholder="Ex: Qual a causa da morte?"
                )
                q_item["resposta"] = st.text_area(
                    "Resposta",
                    value=q_item.get("resposta", ""),
                    key=f"quesito_resp_{i}",
                    height=75,
                    placeholder="Ex: Traumatismo cranioencefálico decorrente de PAF."
                )

                if len(quesitos_list) > 1:
                    if st.button(f"🗑️ Remover Quesito #{i+1}", key=f"remove_quesito_{i}", type="secondary"):
                        st.session_state.quesitos_list.pop(i)
                        st.rerun()
                st.markdown("<hr style='margin:16px 0; border-color:#cbd5e1;'>", unsafe_allow_html=True)

            if st.button("➕ Adicionar Quesito", key="btn_add_quesito"):
                st.session_state.quesitos_list.append({"pergunta": "", "resposta": ""})
                st.rerun()

            st.divider()
            st.subheader("f) Dinâmica dos Fatos / Conclusão Pericial")
            st.markdown("<p style='font-size:13px; color:#475569; margin-bottom:12px;'>Elabore a narrativa técnico-científica dos acontecimentos ou solicite à IA Gemini uma sugestão automatizada baseada nos dados do laudo.</p>", unsafe_allow_html=True)

            col_d1, col_d2 = st.columns([2.2, 1])
            with col_d1:
                if st.button("🤖 Gerar Sugestão de Dinâmica dos Fatos com IA", type="primary", use_container_width=True, key="btn_gerar_dinamica_ia"):
                    with st.spinner("Sintetizando vestígios, lesões e local para sugerir a Dinâmica dos Fatos..."):
                        dados_dinamica = {
                            "ocorrencia": st.session_state.get("ocorrencia"),
                            "endereco": st.session_state.get("endereco"),
                            "delimitacoes": st.session_state.get("delimitacoes"),
                            "isolamento": st.session_state.get("iso_estado"),
                            "instrumento": st.session_state.get("inst_acao"),
                            "agente_compativel": st.session_state.get("inst_agente"),
                            "achados_extras": st.session_state.get("inst_extra"),
                            "vitimas": st.session_state.get("vitimas"),
                            "vestigios": st.session_state.get("vestigios"),
                            "quesitos": st.session_state.get("quesitos_list")
                        }
                        prompt_dinamica = (
                            "Com base exclusivamente nos dados técnicos do laudo pericial fornecidos abaixo, elabore o texto formal "
                            "da 'Dinâmica dos Fatos / Conclusão Pericial'. A narrativa deve correlacionar a posição e lesões da(s) vítima(s), "
                            "os vestígios materiais coletados, o instrumento utilizado e o cenário do local, apresentando em sequência cronológica e lógica "
                            "provável como os fatos ocorreram, mantendo tom imparcial, técnico-científico e jurídico de laudo oficial da Polícia Civil.\n\n"
                            f"DADOS DO LAUDO:\n{json.dumps(dados_dinamica, ensure_ascii=False, indent=2, default=str)}"
                        )
                        sys_dinamica = (
                            "Você é um perito criminal relator especialista em homicídios e locais de morte violenta. "
                            "Redija uma Dinâmica dos Fatos formal, coesa e com vocabulário técnico rigoroso para constar no laudo oficial."
                        )
                        res, err = call_gemini_text(prompt_dinamica, system_instruction=sys_dinamica)
                        if err:
                            st.error(f"Erro ao gerar dinâmica com IA: {err}")
                        else:
                            st.session_state["dinamica_fatos"] = res
                            st.success("✅ Sugestão de Dinâmica dos Fatos gerada com sucesso!")

            with col_d2:
                if st.button("✨ Polir Redação da Dinâmica", use_container_width=True, key="btn_polir_dinamica"):
                    modal_polir_redacao("dinamica_fatos", "Dinâmica dos Fatos")

            dinamica_text = st.text_area(
                "Descrição da Dinâmica dos Fatos",
                value=st.session_state.get("dinamica_fatos", ""),
                height=180,
                key="dinamica_fatos_ui",
                placeholder="Clique no botão acima para gerar a narrativa por IA ou digite a síntese da dinâmica..."
            )
            st.session_state["dinamica_fatos"] = dinamica_text
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    with st.container():
        st.markdown(
            '<div class="section-title">📷 &nbsp; 5. Fotografias (Apêndice A)</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>',
                        unsafe_allow_html=True)

            if "fotos" not in st.session_state:
                st.session_state["fotos"] = []

            st.markdown("<p style='font-size:13px; color:#475569; margin-bottom:12px;'>Faça o upload de fotos do local para compor o Apêndice A (Tomadas Fotográficas) do laudo.</p>", unsafe_allow_html=True)

            uploaded_files = st.file_uploader(
                "Selecione uma ou mais imagens (PNG, JPG, JPEG)",
                type=["png", "jpg", "jpeg"],
                accept_multiple_files=True,
                key="foto_uploader"
            )

            if uploaded_files:
                import base64
                novas_adicionadas = False
                for up_file in uploaded_files:
                    filename = up_file.name
                    if not any(f.get("nome") == filename for f in st.session_state["fotos"]):
                        bytes_data = up_file.read()
                        b64_str = base64.b64encode(bytes_data).decode("utf-8")
                        st.session_state["fotos"].append({
                            "nome": filename,
                            "b64": b64_str,
                            "descricao": f"Visão geral do local ({filename})",
                            "incluir": True
                        })
                        novas_adicionadas = True
                if novas_adicionadas:
                    st.toast("📷 Fotografias adicionadas com sucesso!")
                    st.rerun()

            if st.session_state["fotos"]:
                st.markdown("<hr style='margin:16px 0; border-color:#e2e8f0;'>", unsafe_allow_html=True)

                c_gal_hdr, c_gal_btn = st.columns([2, 1])
                with c_gal_hdr:
                    st.markdown(f"##### 🖼️ Galeria de Fotografias ({len(st.session_state['fotos'])} item(ns))")
                with c_gal_btn:
                    if st.button("✨ Gerar Legendas com IA Gemini", type="primary", key="btn_gemini_legendas_todas", use_container_width=True):
                        if not get_gemini_api_key():
                            st.error("⚠️ Chave GEMINI_API_KEY não configurada. Verifique se o arquivo CHAVE.txt está presente.")
                        else:
                            with st.spinner("✨ Analisando fotografias e gerando legendas periciais formais com Gemini Vision..."):
                                count_sucesso = 0
                                for idx_f, foto in enumerate(st.session_state["fotos"]):
                                    legenda, err_g = gerar_legenda_foto_gemini(foto["b64"])
                                    if legenda:
                                        foto["descricao"] = legenda
                                        count_sucesso += 1
                                if count_sucesso > 0:
                                    st.toast(f"✨ Legendas geradas com sucesso para {count_sucesso} foto(s)!")
                                    st.rerun()
                                else:
                                    st.error("Não foi possível gerar legendas para as fotos.")

                indices_para_remover = []
                for idx_f, foto in enumerate(st.session_state["fotos"]):
                    col_img, col_info = st.columns([1, 2])
                    with col_img:
                        try:
                            img_bytes = base64.b64decode(foto["b64"])
                            st.image(img_bytes, use_container_width=True)
                        except Exception:
                            st.error("Erro ao carregar imagem")
                    with col_info:
                        foto["incluir"] = st.checkbox(
                            "Incluir no Laudo (.docx)",
                            value=foto.get("incluir", True),
                            key=f"foto_inc_{idx_f}"
                        )
                        c_lbl, c_btn_sing = st.columns([2, 1])
                        with c_lbl:
                            st.markdown("<label style='font-size:13px; font-weight:600; color:#334155;'>Legenda / Descrição da Fotografia</label>", unsafe_allow_html=True)
                        with c_btn_sing:
                            if st.button("✨ Legenda IA", key=f"btn_single_gemini_{idx_f}", use_container_width=True, help="Gerar legenda pericial automática com Gemini Vision"):
                                if not get_gemini_api_key():
                                    st.error("⚠️ Chave Gemini API não configurada.")
                                else:
                                    with st.spinner("Gerando legenda..."):
                                        leg_s, err_s = gerar_legenda_foto_gemini(foto["b64"])
                                        if leg_s:
                                            foto["descricao"] = leg_s
                                            st.toast(f"✨ Legenda gerada para Foto #{idx_f+1}!")
                                            st.rerun()
                                        else:
                                            st.error(f"Erro: {err_s}")

                        foto["descricao"] = st.text_area(
                            "Legenda / Descrição da Fotografia",
                            value=foto.get("descricao", ""),
                            key=f"foto_desc_{idx_f}",
                            height=80,
                            label_visibility="collapsed",
                            placeholder="Descreva o elemento ou visão registrada nesta foto..."
                        )
                        if st.button(f"🗑️ Remover Foto #{idx_f+1}", key=f"btn_del_foto_{idx_f}", type="secondary"):
                            indices_para_remover.append(idx_f)
                    st.markdown("<hr style='margin:12px 0; border-color:#cbd5e1;'>", unsafe_allow_html=True)

                if indices_para_remover:
                    for i in reversed(indices_para_remover):
                        st.session_state["fotos"].pop(i)
                    st.rerun()

                if st.button("🗑️ Limpar Todas as Fotografias", type="secondary", key="btn_clear_all_fotos"):
                    st.session_state["fotos"] = []
                    st.rerun()
            else:
                st.info("Nenhuma fotografia adicionada ainda. Utilize o campo acima para enviar as fotos do laudo.")

    # ── BOTÕES INFERIORES ──
    gerar_bottom = render_action_buttons("bottom")

    gerar_clicked = gerar_top or gerar_bottom

    if gerar_clicked:
        try:
            import os
            template_path = os.path.join(
                os.path.dirname(__file__), "Modelo Laudo.docx")
            doc = Document(template_path)

            # Fetch dynamically selected data
            dp_val = st.session_state.get("data_pericia_input", date.today())
            da_val = st.session_state.get(
                "data_atendimento_input", date.today())
            horario_val = st.session_state.get("horario", None)
            horario_a_val = st.session_state.get("horario_atendimento", None)

            from datetime import datetime, date, time as dt_time

            def safe_date_obj(d):
                if isinstance(d, str):
                    try:
                        return datetime.strptime(d, "%Y-%m-%d").date()
                    except:
                        pass
                return d

            def safe_time_str(t):
                if not t:
                    return "____"
                if isinstance(t, str):
                    return t[:5]
                try:
                    return t.strftime("%H:%M")
                except:
                    return str(t)

            dp_val = safe_date_obj(dp_val)
            da_val = safe_date_obj(da_val)

            dp_extenso = data_extenso(dp_val)
            dp_simples = data_simples(dp_val)
            da_simples = data_simples(da_val)
            h_str = safe_time_str(horario_val)
            ha_str = safe_time_str(horario_a_val)

            tipo_local_val = st.session_state.get("tipo_local", "")

            vitimas_list = st.session_state.get("vitimas", [{"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "",
                                                "naturalidade": "", "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "", "fenomenos": "", "lesoes": [""]}])
            nomes_v = [vt["nome"] for vt in vitimas_list if vt["nome"].strip()]
            vitima_header = ", ".join(nomes_v) if nomes_v else "________"

            vitimas_detalhes_list = []
            individual_vars = {}
            for idx, vt in enumerate(vitimas_list, 1):
                suffix = f"_vitima{idx}"
                membros_key = "Membros_vitima!" if idx == 1 else f"Membros_vitima{idx}"

                dob_v = vt.get("data_nascimento", None)
                dob_str = data_simples(dob_v) if dob_v else "________"

                lesoes_list = vt.get("lesoes", [""])
                lesao1_val = lesoes_list[0] if len(lesoes_list) > 0 else ""
                lesao2_val = lesoes_list[1] if len(lesoes_list) > 1 else ""

                individual_vars[f"nome{suffix}"] = vt.get("nome", "")
                individual_vars[f"numero_documento{suffix}"] = vt.get(
                    "documento", "")
                individual_vars[f"sexo{suffix}"] = vt.get("sexo", "")
                individual_vars[f"data_nascimento{suffix}"] = dob_str
                individual_vars[f"filicao{suffix}"] = vt.get("filicao", "")
                individual_vars[f"naturalidade{suffix}"] = vt.get(
                    "naturalidade", "")
                individual_vars[f"veste{suffix}"] = vt.get("vestes", "")
                individual_vars[f"pertence{suffix}"] = vt.get("pertences", "")
                individual_vars[f"localização{suffix}"] = vt.get(
                    "localizacao", "")
                individual_vars[f"posição{suffix}"] = vt.get("posicao", "")
                individual_vars[f"cabeça{suffix}"] = vt.get("cabeca", "")
                individual_vars[membros_key] = vt.get("membros", "")
                individual_vars[f"fenômenos_transformativo{suffix}"] = vt.get(
                    "fenomenos", "")
                individual_vars[f"Lesao1{suffix}"] = lesao1_val
                individual_vars[f"Lesao2{suffix}"] = lesao2_val

                lesoes_txt_list = []
                for les in lesoes_list:
                    if les.strip():
                        lesoes_txt_list.append(f"• {les.strip()}")
                lesoes_text = "\n".join(
                    lesoes_txt_list) if lesoes_txt_list else "• Não descritas"

                text = (
                    f"3.3.{idx} – Descrição da Vítima {idx}\n"
                    f"Identificação\n"
                    f"a) Nome: {vt.get('nome') or '________'}\n"
                    f"b) Documento: {vt.get('documento') or '________'}\n"
                    f"c) Sexo: {vt.get('sexo') or '________'}\n"
                    f"d) Data de Nascimento: {dob_str}\n"
                    f"e) Filiação: {vt.get('filicao') or '________'}\n"
                    f"f) Naturalidade: {vt.get('naturalidade') or '________'}\n\n"
                    f"Vestes e Pertences\n"
                    f"a) Vestes: {vt.get('vestes') or '________'}\n"
                    f"b) Pertences: {vt.get('pertences') or '________'}\n\n"
                    f"Posição e Localização (In Situ)\n"
                    f"a) Localização: {vt.get('localizacao') or '________'}\n"
                    f"b) Posição: {vt.get('posicao') or '________'}\n"
                    f"c) Cabeça: {vt.get('cabeca') or '________'}\n"
                    f"d) Membros: {vt.get('membros') or '________'}\n\n"
                    f"Exame Perinecroscópico e Estado de Conservação\n"
                    f"a) Fenômenos Transformativos: {vt.get('fenomenos') or '________'}\n"
                    f"b) Lesões:\n"
                    f"{lesoes_text}"
                )
                vitimas_detalhes_list.append(text)

            vitimas_detalhes = "\n\n".join(vitimas_detalhes_list)

            vestigios_list = st.session_state.get("vestigios", [])
            vestigios_detalhes_list = []
            for idx_v, vest in enumerate(vestigios_list, 1):
                if vest.get("tipo") == "Manchas de Sangue":
                    txt = (f"Vestígio {idx_v} (Mancha de Sangue):\n"
                           f"Categoria: {vest.get('categoria') or '________'} - Subtipo: {vest.get('subtipo') or '________'}\n"
                           f"Localização: {vest.get('localizacao') or '________'}\n"
                           f"Descrição: {vest.get('descricao') or '________'}")
                    vestigios_detalhes_list.append(txt)
                elif vest.get("tipo") == "Elemento Balístico":
                    txt = (f"Vestígio {idx_v} (Elemento Balístico):\n"
                           f"Tipo: {vest.get('elemento_tipo') or '________'} - Quantidade: {vest.get('quantidade') or '________'}\n"
                           f"Envelope de Segurança: {vest.get('envelope') or '________'}\n"
                           f"Descrição: {vest.get('descricao') or '________'}")
                    vestigios_detalhes_list.append(txt)
                elif vest.get("tipo") == "Arma de Fogo":
                    txt = (f"Vestígio {idx_v} (Arma de Fogo):\n"
                           f"Tipo: {vest.get('arma_tipo') or '________'} - Municiada: {vest.get('municiada') or '________'}\n"
                           f"Envelope de Segurança: {vest.get('envelope') or '________'}\n"
                           f"Descrição: {vest.get('descricao') or '________'}")
                    vestigios_detalhes_list.append(txt)
                elif vest.get("tipo") == "Outros":
                    txt = (f"Vestígio {idx_v} (Outros):\n"
                           f"Descrição: {vest.get('descricao') or '________'}")
                    vestigios_detalhes_list.append(txt)
                elif vest.get("tipo") == "Impressão Datiloscópica":
                    txt = (f"Vestígio {idx_v} (Impressão Datiloscópica):\n"
                           f"Descrição: {vest.get('descricao') or '________'}")
                    vestigios_detalhes_list.append(txt)
                elif vest.get("tipo") == "Instrumento Lesivo":
                    txt = (f"Vestígio {idx_v} (Instrumento Lesivo):\n"
                           f"Tipo: {vest.get('instrumento_tipo') or '________'}\n"
                           f"Envelope de Segurança: {vest.get('envelope') or '________'}\n"
                           f"Descrição: {vest.get('descricao') or '________'}")
                    vestigios_detalhes_list.append(txt)
                elif vest.get("tipo") == "Aparelho Telefônico":
                    txt = (f"Vestígio {idx_v} (Aparelho Telefônico):\n"
                           f"Marca/Modelo: {vest.get('marca') or '________'} - Cor: {vest.get('cor') or '________'}\n"
                           f"Funcionamento: {vest.get('funcionamento') or '________'}\n"
                           f"Envelope de Segurança: {vest.get('envelope') or '________'}\n"
                           f"Descrição: {vest.get('descricao') or '________'}")
                    vestigios_detalhes_list.append(txt)
            vestigios_detalhes = "\n\n".join(
                vestigios_detalhes_list) if vestigios_detalhes_list else "Não foram descritos vestígios."

            quesitos_formatted_list = []
            for idx_q, q_item in enumerate(st.session_state.get("quesitos_list", []), 1):
                p_val = q_item.get("pergunta", "").strip()
                r_val = q_item.get("resposta", "").strip()
                if p_val or r_val:
                    p_text = p_val if p_val else "________"
                    r_text = r_val if r_val else "________"
                    quesitos_formatted_list.append(f"Quesito {idx_q}: {p_text}\nResposta: {r_text}")

            quesitos_final_str = "\n\n".join(quesitos_formatted_list) if quesitos_formatted_list else "Não foram formulados quesitos específicos."
            st.session_state["quesitos"] = quesitos_final_str

            VARS = {
                "num_laudo": st.session_state.get("num_laudo", ""),
                "ocorrencia": st.session_state.get("ocorrencia", ""),
                "ocorrência": st.session_state.get("ocorrencia", ""),
                "perito": st.session_state.get("perito", ""),
                "autoridade": st.session_state.get("autoridade_sel", ""),
                "requisicao": st.session_state.get("requisicao", ""),
                "vitima": vitima_header,
                "investigado": st.session_state.get("investigado", ""),
                "referencia": st.session_state.get("referencia", ""),
                "destino": st.session_state.get("destino", ""),
                "data_extenso": dp_extenso,
                "data_simples": da_simples,
                "horario": h_str,
                "horario_atend": ha_str,
                "endereco": st.session_state.get("endereco", ""),
                "latitude": st.session_state.get("latitude", ""),
                "longitude": st.session_state.get("longitude", ""),
                "ponto": st.session_state.get("ponto_referencia", ""),
                "area": st.session_state.get("area", ""),
                "pavimento": st.session_state.get("pavimento", ""),
                "delimitacoes": st.session_state.get("delimitacoes", ""),
                "isolamento": st.session_state.get("isolamento_outros", "") if st.session_state.get("isolamento", "") == "Outros" else st.session_state.get("isolamento", ""),
                "clima": st.session_state.get("clima_outros", "") if st.session_state.get("clima", "") == "Outros" else st.session_state.get("clima", ""),
                "visibilidade": st.session_state.get("visibilidade", ""),
                "iluminacao": st.session_state.get("iluminacao", ""),
                "equipe_pm": st.session_state.get("equipe_pm", ""),
                "equipe_pc": st.session_state.get("equipe_pc", ""),
                "aut_local": st.session_state.get("autoridade_local", ""),
                "vitimas_detalhes": vitimas_detalhes,
                "vestigios_detalhes": vestigios_detalhes,
                "quesitos": quesitos_final_str,
                "quesitos_respostas": quesitos_final_str,
                "dinamica_fatos": st.session_state.get("dinamica_fatos", "") or "Dinâmica não descrita.",
                "dinâmica_dos_fatos": st.session_state.get("dinamica_fatos", "") or "Dinâmica não descrita.",
                "dinamica": st.session_state.get("dinamica_fatos", "") or "Dinâmica não descrita.",
            }
            VARS.update(individual_vars)
            VARS["tipo_local"] = tipo_local_val

            # --- Lógica de Isolamento ---
            tpl_isolamento = {
                "1": "No momento da chegada da equipe pericial, o local encontrava-se devidamente isolado por meio de {meio_de_isolamento}, impedindo o acesso de pessoas não autorizadas ao perímetro de interesse. Constatou-se a integral preservação do estado das coisas, não havendo quaisquer indícios de alteração, supressão, contaminação ou acréscimo de vestígios. Tais condições atestam o fiel cumprimento ao Art. 6º, inciso I, e Art. 169 do Código de Processo Penal, garantindo a idoneidade da etapa de isolamento da Cadeia de Custódia (Art. 158-A, § 2º) e conferindo total confiabilidade ao levantamento pericial e à dinâmica interpretada.",
                "2": "O sítio pericial apresentava-se isolado por {meio_de_isolamento}, restringindo o trânsito de pessoas estranhas aos exames naquele momento. Contudo, constatou-se que o ambiente encontrava-se apenas parcialmente preservado. Notabilizou-se a alteração do estado original das coisas, consubstanciada por {descrever_alteracao}, ocorrida em momento anterior ao isolamento efetivo. Embora a restrição de acesso atual seja adequada, a descaracterização pretérita impõe ressalvas na interpretação absoluta da dinâmica do evento, sem, contudo, inviabilizar a análise dos vestígios remanescentes.",
                "3": "O perímetro de interesse criminalístico encontrava-se parcialmente isolado, apresentando vulnerabilidades na delimitação física, tais como {descrever_falha_isolamento}, permitindo potencial acesso de curiosos. Não obstante a fragilidade da contenção perimetral, o núcleo do evento encontrava-se preservado. As evidências materiais e a disposição dos elementos encontravam-se em seu estado aparente original, sem sinais de manipulação, contaminação ou descaracterização, permitindo a devida fixação, coleta e garantia da cadeia de custódia dos vestígios encontrados.",
                "4": "O local dos fatos encontrava-se parcialmente isolado, com delimitação perimetral incipiente e insuficiente para {descrever_falha_isolamento}. Concomitantemente, verificou-se a preservação apenas parcial do ambiente, evidenciada por {descrever_alteracao}. Esta conjugação de ineficiência no resguardo do perímetro e a consequente alteração do estado original mitigam a robustez da análise da dinâmica delitiva, configurando inobservância parcial aos ritos de inalterabilidade exigidos pela legislação processual penal vigente.",
                "5": "No momento do acionamento e chegada desta equipe, o sítio pericial encontrava-se desprovido de qualquer isolamento e totalmente devassado. Constatou-se a ausência absoluta de preservação do estado de coisas, com evidências de intensa e irreversível modificação do cenário original, notadamente por {descrever_alteracao}. Tal inobservância à garantia da inalterabilidade do sítio (Art. 6º, I e Art. 169 do CPP) compromete severamente as etapas iniciais da Cadeia de Custódia, prejudicando o estabelecimento inconteste do nexo de causalidade e limitando os achados periciais apenas aos vestígios intrínsecos."
            }
            iso_val_str = str(st.session_state.get("iso_estado", "1")).strip()
            iso_key = iso_val_str[0] if iso_val_str else "1"
            iso_text = tpl_isolamento.get(iso_key, "")
            iso_text = iso_text.replace(
                "{meio_de_isolamento}", st.session_state.get("iso_meio") or "________")
            iso_text = iso_text.replace(
                "{descrever_alteracao}", st.session_state.get("iso_alteracao") or "________")
            iso_text = iso_text.replace(
                "{descrever_falha_isolamento}", st.session_state.get("iso_falha") or "________")

            VARS["isolamento_detalhes"] = iso_text

            # Map IML variables
            VARS["número_laudo_necropsia"] = st.session_state.get(
                "numero_laudo_necropsia") or "________"
            VARS["numero_laudo_necropsia"] = st.session_state.get(
                "numero_laudo_necropsia") or "________"
            VARS["resultado_laudo_IML"] = st.session_state.get(
                "resultado_laudo_IML") or "________"

            # --- Lógica de Instrumento Utilizado ---
            tpl_instrumento = {
                "1": "As lesões constatadas na(s) vítima(s) são características daquelas produzidas por instrumento em ação perfurocontundente, comprovadamente produzidas por projéteis de arma de fogo (PAF), {achado_extra}.",
                "2": "As lesões constatadas na(s) vítima(s) são características daquelas produzidas por instrumento em ação cortante (feridas incisas), comprovadamente produzidas por deslizamento de gume afiado, compatível com {agente_compativel}, {achado_extra}.",
                "3": "As lesões constatadas na(s) vítima(s) são características daquelas produzidas por instrumento em ação perfurante (feridas punctórias), comprovadamente produzidas por agente de ponta fina, compatível com {agente_compativel}, {achado_extra}.",
                "4": "As lesões constatadas na(s) vítima(s) são características daquelas produzidas por instrumento em ação perfurocortante (feridas perfuroincisas), comprovadamente produzidas por arma branca dotada de ponta e gume(s), compatível com {agente_compativel}, {achado_extra}.",
                "5": "As lesões constatadas na(s) vítima(s) são características daquelas produzidas por instrumento em ação contundente (feridas contusas / equimoses / fraturas), comprovadamente produzidas por choque ou impacto contra superfície rígida, compatível com {agente_compativel}, {achado_extra}.",
                "6": "As lesões constatadas na(s) vítima(s) são características daquelas produzidas por instrumento em ação cortocontundente (feridas contuso-incisas), comprovadamente produzidas por agente dotado de massa expressiva e gume, compatível com {agente_compativel}, {achado_extra}.",
                "7": "As lesões constatadas na(s) vítima(s) são características daquelas produzidas por energia de ordem física (ação térmica), comprovadamente decorrentes de exposição a {agente_compativel}, {achado_extra}."
            }

            inst_val_str = str(st.session_state.get("inst_acao", "1")).strip()
            inst_k = inst_val_str[0] if inst_val_str else "1"
            inst_text = tpl_instrumento.get(inst_k, "")
            agente_val = st.session_state.get("inst_agente") or "________"
            extra_val = st.session_state.get("inst_extra") or "________"

            inst_text = inst_text.replace("{agente_compativel}", agente_val)
            inst_text = inst_text.replace("{achado_extra}", extra_val)

            # Cleanup if extra ends with double period or unformatted spaces
            inst_text = inst_text.replace("..", ".")

            VARS["instrumento_detalhes"] = inst_text
            # Map consideracoes_extras in VARS for DOCX
            formatted_extras = []
            for idx_extra, item_extra in enumerate(st.session_state.get("consideracoes_extras", [])):
                letra = chr(ord('e') + idx_extra)
                tit = item_extra.get("titulo", "").strip()
                txt = item_extra.get("texto", "").strip()
                if tit and txt:
                    formatted_extras.append(f"{letra}) {tit}\n{txt}")
                elif txt:
                    formatted_extras.append(f"{letra}) {txt}")

            VARS["consideracoes_extras"] = "\n\n".join(formatted_extras) if formatted_extras else ""

            # Also map to target paragraph if tag is missing in docx
            VARS[
                "d. As lesões constatadas nas vítimas são características daquelas produzidas por instrumento em ação perfurocontundente, inequivocamente produzidas por projéteis de arma de fogo (PAF), com a recuperação de um projétil em cada cadáver durante a necrópsia."] = f"d. {inst_text}"

            def replace_para(para, var_dict):
                for k, v in var_dict.items():
                    target = "{" + k + "}"
                    if target in para.text:
                        val_str = str(v) if v else "________"
                        if "\\n" in val_str or "\n" in val_str:
                            val_str = val_str.replace("\\n", "\n")
                            lines = val_str.split("\n")
                            parts = para.text.split(target, 1)
                            prefix = parts[0]
                            suffix = parts[1] if len(parts) > 1 else ""

                            # Original para will become the LAST line.
                            # So we insert all lines except the last one BEFORE para.
                            for i, line in enumerate(lines[:-1]):
                                p_new = para.insert_paragraph_before(
                                    style=para.style)
                                if i == 0:
                                    p_new.add_run(prefix + line)
                                else:
                                    p_new.add_run(line)

                            # Now modify the original para to hold the last line
                            para.text = ""
                            if len(lines) == 1:
                                para.add_run(prefix + lines[0] + suffix)
                            else:
                                para.add_run(lines[-1] + suffix)

                            # Recursively call in case there are other tags in the same paragraph
                            replace_para(para, var_dict)
                            return
                        else:
                            new = para.text.replace(target, val_str)
                            if len(para.runs) > 0:
                                para.runs[0].text = new
                                for r in para.runs[1:]:
                                    r.text = ""
                            else:
                                para.add_run(new)

            for para in doc.paragraphs:
                replace_para(para, VARS)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            replace_para(para, VARS)
            for section in doc.sections:
                if section.header:
                    for para in section.header.paragraphs:
                        replace_para(para, VARS)
                if section.footer:
                    for para in section.footer.paragraphs:
                        replace_para(para, VARS)

            # --- Inserção de Fotografias no Apêndice A ---
            fotos_para_laudo = [f for f in st.session_state.get(
                "fotos", []) if f.get("incluir", True)]
            if fotos_para_laudo:
                import base64
                from io import BytesIO
                from docx.shared import Inches, Pt
                from docx.enum.text import WD_ALIGN_PARAGRAPH

                # Locate Apêndice A paragraph or append to document
                apendice_p = None
                for p in doc.paragraphs:
                    if "APÊNDICE A" in p.text.upper() or "TOMADAS FOTOGRÁFICAS" in p.text.upper():
                        apendice_p = p

                # If not found, add header
                if not apendice_p:
                    apendice_p = doc.add_paragraph(
                        "APÊNDICE A – TOMADAS FOTOGRÁFICAS")
                    apendice_p.runs[0].bold = True

                # Clear placeholders after apendice_p if any
                # Append each photo with description ABOVE and picture BELOW
                for f_idx, foto in enumerate(fotos_para_laudo, 1):
                    # Add Description Paragraph (ABOVE)
                    p_desc = doc.add_paragraph()
                    p_desc.paragraph_format.space_before = Pt(12)
                    p_desc.paragraph_format.space_after = Pt(4)

                    r_num = p_desc.add_run(f"Fotografia {f_idx} – ")
                    r_num.bold = True
                    r_desc = p_desc.add_run(foto.get("descricao") or "")

                    # Add Image Paragraph (BELOW)
                    try:
                        img_data = base64.b64decode(foto["b64"])
                        img_stream = BytesIO(img_data)
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_after = Pt(18)
                        r_img = p_img.add_run()
                        r_img.add_picture(img_stream, width=Inches(5.5))
                    except Exception as img_err:
                        st.warning(
                            f"Não foi possível inserir a fotografia {f_idx}: {str(img_err)}")

            buf = BytesIO()
            doc.save(buf)
            buf.seek(0)

            ocr = st.session_state.get("ocorrencia", "")
            nome = f"Laudo_{ocr.replace('/','_') if ocr else 'sem_numero'}.docx"
            st.session_state.docx_bytes = buf.getvalue()
            st.session_state.docx_filename = nome
            st.success("✅ Laudo gerado com sucesso!")
        except Exception as e:
            st.error(f"❌ Erro ao gerar o laudo: {str(e)}")

        if st.session_state.get("docx_bytes"):
            st.download_button(
                label="⬇️  Baixar Laudo Gerado",
                data=st.session_state.docx_bytes,
                file_name=st.session_state.docx_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,

            )
