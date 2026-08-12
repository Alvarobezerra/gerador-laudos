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
MESES = ["janeiro","fevereiro","março","abril","maio","junho",
         "julho","agosto","setembro","outubro","novembro","dezembro"]
EXTENSO = {
    1:"um",2:"dois",3:"três",4:"quatro",5:"cinco",6:"seis",
    7:"sete",8:"oito",9:"nove",10:"dez",11:"onze",12:"doze",
    13:"treze",14:"catorze",15:"quinze",16:"dezesseis",17:"dezessete",
    18:"dezoito",19:"dezenove",20:"vinte",21:"vinte e um",22:"vinte e dois",
    23:"vinte e três",24:"vinte e quatro",25:"vinte e cinco",26:"vinte e seis",
    27:"vinte e sete",28:"vinte e oito",29:"vinte e nove",30:"trinta",31:"trinta e um",
    2025:"dois mil e vinte e cinco",2026:"dois mil e vinte e seis",
    2027:"dois mil e vinte e sete",2028:"dois mil e vinte e oito",
    2029:"dois mil e vinte e nove",2030:"dois mil e trinta",
}

def num_extenso(n):    return EXTENSO.get(n, str(n))
def data_extenso(d):
    return (f"Aos {d.day:02d} ({num_extenso(d.day)}) dia(s) do mês de "
            f"{MESES[d.month-1]} do ano {d.year} ({num_extenso(d.year)})")
def data_simples(d):   return d.strftime("%d/%m/%Y")
def v(val):            return str(val).strip() if str(val).strip() else "________"

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
for key, default in {
    "autoridades":   AUTORIDADES_DEFAULT.copy(),
    "preview_open":  False,
    "show_nova_aut": False,
    "docx_bytes":    None,
    "docx_filename": "",
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
    background-color: #f8fafc !important;
    color: #0f172a !important;
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
    font-size: 16px !important;
    font-weight: 700 !important;
    color: #1e293b !important;
    margin-top: 15px !important;
    margin-bottom: -10px !important;
    display: flex !important;
    align-items: center !important;
    gap: 8px !important;
    padding-left: 8px;
    border-left: 4px solid #2563eb;
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
                submitted = st.form_submit_button("Entrar", type="primary", use_container_width=True)
        
                if submitted:
                    if (user.lower() in ["alvaro", "perito"]) and pwd in ["123", "icrim123"]:
                        st.session_state["autenticado"] = True
                        st.rerun()
                    else:
                        st.error("❌ Usuário ou senha incorretos.")
    
    st.stop() # Bloqueia a execução do restante do código


# ═══════════════════════════════════════════════════════════
# MODO PREVIEW
# ═══════════════════════════════════════════════════════════
if st.session_state.preview_open:

    def ss(key, default=""):
        return st.session_state.get(key, default) or default

    num_laudo      = ss("num_laudo")
    ocorrencia     = ss("ocorrencia")
    perito         = ss("perito")
    autoridade     = ss("autoridade_sel")
    requisicao     = ss("requisicao")
    referencia     = ss("referencia")  or "(sem referência)"
    destino        = ss("destino")
    
    vitimas_list   = st.session_state.get("vitimas", [{"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "", "naturalidade": "", "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "", "fenomenos": "", "lesoes": [""]}])
    nomes          = [vt["nome"] for vt in vitimas_list if vt["nome"].strip()]
    vitima         = ", ".join(nomes) if nomes else "________"

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
                dp_val         = st.session_state.get("data_pericia_input",     date.today())
                da_val         = st.session_state.get("data_atendimento_input", date.today())
                horario_val    = st.session_state.get("horario",               None)
                horario_a_val  = st.session_state.get("horario_atendimento",   None)
                endereco       = ss("endereco")
                latitude       = ss("latitude")
                longitude      = ss("longitude")
                ponto          = ss("ponto_referencia")
                area           = ss("area")
                pavimento      = ss("pavimento")
                delimitacoes   = ss("delimitacoes")
                iso_s          = ss("isolamento")
                isolamento     = ss("isolamento_outros") if iso_s == "Outros" else iso_s
                clim_s         = ss("clima")
                clima          = ss("clima_outros") if clim_s == "Outros" else clim_s
                visibilidade   = ss("visibilidade")
                iluminacao     = ss("iluminacao")
                equipe_pm      = ss("equipe_pm")
                equipe_pc      = ss("equipe_pc")
                aut_local      = ss("autoridade_local")

                dp_ext  = data_extenso(dp_val)
                da_simp = data_simples(da_val)
                h_str   = horario_val.strftime("%H:%M")   if horario_val  else "____"
                ha_str  = horario_a_val.strftime("%H:%M") if horario_a_val else "____"

                tipo_local_val = st.session_state.get("tipo_local", "")
                ano     = date.today().year

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


    # ═══════════════════════════════════════════════════════════
    # MODO FORMULÁRIO (LOGIN DESABILITADO TEMPORARIAMENTE)
    # ═══════════════════════════════════════════════════════════
    st.session_state['logged_in'] = True

    # ── Cabeçalho ──────────────────────────────────────────────
    st.markdown("""
    <div class="app-header">
      <div class="app-header-brand">
        <div class="app-header-badge">ICRIM / NPT</div>
        <div>
          <div class="app-header-title">Laudo de Local de Morte Violenta</div>
          <div class="app-header-sub">Instituto de Criminalística de Imperatriz &nbsp;|&nbsp; Polícia Civil — Maranhão</div>
    </div>
  </div>
</div>
<div style="height:24px"></div>
""", unsafe_allow_html=True)

# ── Margem lateral para conteúdo ──────────────────────────
_, main, _ = st.columns([0.04, 0.92, 0.04])

# ── Sidebar Sincronização Mobile ──────────────────────────
with st.sidebar:
    st.markdown("## AÇÕES DA OCORRÊNCIA")
    st.divider()
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    if st.button("Gerar Prévia do Laudo", use_container_width=True, type="primary"):
        st.session_state.preview_open = True
        st.rerun()

    st.divider()
    if st.button("💾 Salvar Rascunho no Banco", use_container_width=True):
        DB_PATH = os.path.join(os.path.dirname(__file__), "banco_laudos.sqlite")
        try:
            import time, sqlite3, json
            from datetime import datetime, date, time as dt_time
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute('CREATE TABLE IF NOT EXISTS laudos (id INTEGER PRIMARY KEY AUTOINCREMENT, mobile_id TEXT UNIQUE, data_sincronizacao TEXT, dados_json TEXT)')
            mob_id = st.session_state.get('ocorrencia')
            if not mob_id:
                mob_id = f"manual_{int(time.time())}"
            now_str = datetime.now().isoformat()
            data_to_save = {}
            for k, v in st.session_state.items():
                if isinstance(v, (str, int, float, bool, list, dict)):
                    data_to_save[k] = v
                elif isinstance(v, (date, datetime, dt_time)):
                    data_to_save[k] = str(v)
            c.execute('INSERT OR REPLACE INTO laudos (mobile_id, data_sincronizacao, dados_json) VALUES (?, ?, ?)',
                      (mob_id, now_str, json.dumps(data_to_save)))
            conn.commit()
            conn.close()
            st.success("✅ Formulário salvo no banco local!")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

GAS_URL = "https://script.google.com/macros/s/AKfycbx9N4hidxHbfAUUpHVDAOadZBF5SRB9x9UzC0nt3k-wW0pgLThCHwBQsaUMJjtcTL1QJw/exec"

with main:
    with st.expander("📂 Ocorrências Salvas Localmente", expanded=False):
        st.markdown('<div class="section-title">📂 &nbsp; Minhas Ocorrências</div>', unsafe_allow_html=True)
        st.caption("Ocorrências sincronizadas do celular via Google Sheets e salvas no banco local.")

        # Fetch from Google Sheets
        ocorrencias_cloud = []
        try:
            resp = requests.get(GAS_URL, params={"key": "perito:icrim123"}, timeout=10)
            if resp.status_code == 200:
                ocorrencias_cloud = resp.json()
        except Exception:
            pass

        if ocorrencias_cloud and isinstance(ocorrencias_cloud, list):
            st.success(f"☁️ {len(ocorrencias_cloud)} ocorrência(s) encontrada(s) na planilha do Google Drive.")
            for idx, item in enumerate(ocorrencias_cloud):
                try:
                    dados = json.loads(item.get("dados_json", "{}")) if isinstance(item.get("dados_json"), str) else item.get("dados_json", {})
                except Exception:
                    dados = {}
                oc_num = dados.get("ocorrencia", item.get("mobile_id", "S/N"))
                perito_nome = dados.get("perito", "N/I")
                tipo = dados.get("tipo_local", "N/I")
                data_sync = item.get("data_sincronizacao", "")[:16]

                with st.container(border=True):
                    c1, c2, c3 = st.columns([3, 1.5, 1])
                    c1.markdown(f"**Ocorrência: {oc_num}**")
                    c1.caption(f"Perito: {perito_nome} | Tipo: {tipo}")
                    c2.caption(f"Sincronizado: {data_sync}")
                    if c3.button("📥 Carregar", key=f"btn_cloud_{idx}", use_container_width=True):
                        for k, val in dados.items():
                            if k not in ["vitimas", "vestigios", "fotos"]:
                                st.session_state[k] = val
                        if "vitimas" in dados: st.session_state["vitimas"] = dados["vitimas"]
                        if "vestigios" in dados: st.session_state["vestigios"] = dados["vestigios"]
                        if "fotos" in dados: st.session_state["fotos"] = dados["fotos"]
                        st.rerun()
        else:
            st.info("Nenhuma ocorrência sincronizada encontrada na planilha do Google Drive.")

    with st.expander("☁️ Sincronização com o App (Nuvem)", expanded=False):
        st.markdown('<div class="section-title">☁️ &nbsp; Sincronização e Rascunhos de Ocorrências</div>', unsafe_allow_html=True)
        
        import sqlite3, json, os
        DB_PATH = os.path.join(os.path.dirname(__file__), "banco_laudos.sqlite")
        
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            with st.container(border=True):
                st.subheader("🗄️ Ocorrências no Banco de Dados Local")
                st.caption("Ocorrências salvas localmente pelo app web ou API.")
                
                db_laudos = []
                if os.path.exists(DB_PATH):
                    try:
                        conn = sqlite3.connect(DB_PATH)
                        conn.row_factory = sqlite3.Row
                        c = conn.cursor()
                        c.execute('SELECT id, mobile_id, data_sincronizacao, dados_json FROM laudos ORDER BY id DESC')
                        rows = c.fetchall()
                        conn.close()
                        for r in rows:
                            db_laudos.append({
                                "id": r["id"],
                                "mobile_id": r["mobile_id"],
                                "data": r["data_sincronizacao"],
                                "dados": json.loads(r["dados_json"])
                            })
                    except Exception as e:
                        st.error(f"Erro ao acessar banco sqlite: {e}")
                
                if db_laudos:
                    opcoes = {f"Ocorrência {l['dados'].get('ocorrencia', 'S/N')} — Perito: {l['dados'].get('perito', 'N/I')} ({l['data'][:16]})": l for l in db_laudos}
                    escolhida_key = st.selectbox("Selecione a ocorrência salva:", list(opcoes.keys()))
                    
                    if escolhida_key:
                        item = opcoes[escolhida_key]
                        st.info(f"📍 Tipo: {item['dados'].get('tipo_local', 'N/I')} | Vítimas: {len(item['dados'].get('vitimas', []))} | Fotos: {len(item['dados'].get('fotos', []))}")
                        
                        if st.button("📥 CARREGAR ESTA OCORRÊNCIA NO LAUDO", type="primary", key="btn_load_sqlite", use_container_width=True):
                            data_obj = item["dados"]
                            for k, v in data_obj.items():
                                if k not in ["vitimas", "vestigios", "fotos"]:
                                    st.session_state[k] = v
                            if "vitimas" in data_obj: st.session_state["vitimas"] = data_obj["vitimas"]
                            if "vestigios" in data_obj: st.session_state["vestigios"] = data_obj["vestigios"]
                            if "fotos" in data_obj: st.session_state["fotos"] = data_obj["fotos"]
                            
                            st.success("✅ Ocorrência carregada com sucesso no formulário!")
                            st.rerun()
                else:
                    st.warning("Nenhuma ocorrência encontrada no banco de dados local ainda.")

        with col_right:
            with st.container(border=True):
                st.subheader("📥 Importar Arquivo .JSON do Celular")
                st.caption("Se você baixou o arquivo .json do celular clicando em 'Exportar Caso (.json)', selecione-o aqui.")
                
                json_file = st.file_uploader("Selecione o arquivo de ocorrência (.json)", type=["json"], key="uploader_json_sync_2")
                if json_file:
                    try:
                        imported_json_data = json.load(json_file)
                        st.success(f"Arquivo '{imported_json_data.get('ocorrencia', '')}' carregado!")
                        
                        if st.button("🚀 IMPORTAR DADOS DO ARQUIVO JSON", type="primary", key="btn_load_json_file", use_container_width=True):
                            for k, v in imported_json_data.items():
                                if k not in ["vitimas", "vestigios", "fotos"]:
                                    st.session_state[k] = v
                            if "vitimas" in imported_json_data: st.session_state["vitimas"] = imported_json_data["vitimas"]
                            if "vestigios" in imported_json_data: st.session_state["vestigios"] = imported_json_data["vestigios"]
                            if "fotos" in imported_json_data: st.session_state["fotos"] = imported_json_data["fotos"]
                            
                            # Also save to SQLite so it stays in the database
                            try:
                                conn = sqlite3.connect(DB_PATH)
                                c = conn.cursor()
                                mob_id = imported_json_data.get('ocorrencia', f"mob_{len(db_laudos)+1}")
                                now_str = datetime.now().isoformat()
                                c.execute('INSERT OR REPLACE INTO laudos (mobile_id, data_sincronizacao, dados_json) VALUES (?, ?, ?)',
                                          (mob_id, now_str, json.dumps(imported_json_data)))
                                conn.commit()
                                conn.close()
                            except Exception: pass
                            
                            st.success("✅ Ocorrência importada e salva no banco de dados!")
                            st.rerun()
                    except Exception as err:
                        st.error(f"Erro ao ler arquivo: {err}")


    def sanitize_datetime_state():
        from datetime import datetime, date, time
        for k in ["data_pericia_input", "data_atendimento_input"]:
            if k in st.session_state and isinstance(st.session_state[k], str):
                try:
                    st.session_state[k] = datetime.strptime(st.session_state[k], "%Y-%m-%d").date()
                except:
                    st.session_state[k] = date.today()
        for k in ["horario", "horario_atendimento"]:
            if k in st.session_state and isinstance(st.session_state[k], str):
                try:
                    s = st.session_state[k]
                    if len(s.split(':')) == 2: s += ":00"
                    st.session_state[k] = datetime.strptime(s, "%H:%M:%S").time()
                except:
                    st.session_state[k] = None

    sanitize_datetime_state()

    with st.container():
        # ══ SEÇÃO 1: DA OCORRÊNCIA ════════════════════════════
        st.markdown('<div class="section-title">📋 &nbsp; Da Ocorrência</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>', unsafe_allow_html=True)

            c1, c2, c3, c4 = st.columns(4)
            num_laudo  = c1.text_input("Nº do Laudo",         placeholder="0092258",    key="num_laudo")
            ocorrencia = c2.text_input("Nº da Ocorrência",    placeholder="12345/2026", key="ocorrencia")
            requisicao = c3.text_input("Requisição",           key="requisicao")
            referencia = c4.text_input("IP / PM / BO / Proc.", key="referencia")

            c5, c6, c7, c8 = st.columns(4)
            data_pericia_val     = c5.date_input("Data da Perícia",        value=date.today(), key="data_pericia_input")
            horario_val          = c6.time_input("Horário da Perícia",     key="horario",      step=300)
            data_atendimento_val = c7.date_input("Data de Atendimento",    value=date.today(), key="data_atendimento_input")
            horario_atend_val    = c8.time_input("Horário de Atendimento", key="horario_atendimento", step=300)

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
        investigado = cv1.text_area("Investigado(s)",  placeholder="Nomes dos investigados...", key="investigado", height=80)
        destino     = cv2.text_input("Unidade Destino", value="Delegacia de Homicídios de Imperatriz", key="destino")

    with st.container():
        # ══ SEÇÃO 2: DO LOCAL ════════════════════════════════
        st.markdown('<div class="section-title">📍 &nbsp; Do Local</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>', unsafe_allow_html=True)

            ce1, ce2, ce3, ce4 = st.columns([3, 1.2, 1.2, 0.6])
            endereco  = ce1.text_input("Endereço completo", placeholder="Rua, Nº, Bairro, Município", key="endereco")
            latitude  = ce2.text_input("Latitude",  placeholder="-5.518600",  key="latitude")
            longitude = ce3.text_input("Longitude", placeholder="-47.477600", key="longitude")
            with ce4:
                st.markdown("<br>", unsafe_allow_html=True)
                st.button("📍 GPS", key="btn_gps", use_container_width=True,
                          help="GPS disponível via HTTPS")

            cr1, cr2, cr3 = st.columns([2, 1, 1])
            ponto_referencia = cr1.text_input("Ponto de Referência", placeholder="Ex: Próximo ao posto de saúde", key="ponto_referencia")
            area             = cr2.text_input("Área",      placeholder="Ex: Externa", key="area")
            pavimento        = cr3.text_input("Pavimento", placeholder="Ex: Asfalto", key="pavimento")

            delimitacoes = st.text_area("Delimitações do local",
                placeholder="Descreva os limites físicos do local periciado...",
                key="delimitacoes", height=70)

            cl1, cl2, cl3, cl4 = st.columns(4)
            iso_opts = ["", "Preservado e Isolado", "Parcialmente Preservado e Isolado",
                        "Preservado e Parcialmente Isolado", "Não Preservado e Não Isolado", "Outros"]
            isolamento_sel = cl1.selectbox("Isolamento", iso_opts, key="isolamento")

            clima_opts = ["", "Aberto diurno", "Aberto noturno", "Nublado diurno",
                          "Nublado noturno", "Chuvoso diurno", "Chuvoso noturno", "Outros"]
            clima_sel = cl2.selectbox("Clima", clima_opts, key="clima")

            visibilidade = cl3.selectbox("Visibilidade", ["", "Ampla", "Reduzida"], key="visibilidade")
            iluminacao   = cl4.selectbox("Iluminação",
                ["", "Natural Satisfatória", "Natural Insatisfatória",
                 "Artificial Satisfatória", "Artificial Insatisfatória"], key="iluminacao")

            if isolamento_sel == "Outros" or clima_sel == "Outros":
                ot1, ot2 = st.columns(2)
                if isolamento_sel == "Outros":
                    isolamento_txt = ot1.text_input("Especifique o Isolamento", key="isolamento_outros")
                    isolamento = isolamento_txt or "Outros"
                else:
                    isolamento = isolamento_sel
                if clima_sel == "Outros":
                    clima_txt = ot2.text_input("Especifique o Clima", key="clima_outros")
                    clima = clima_txt or "Outros"
                else:
                    clima = clima_sel
            else:
                isolamento = isolamento_sel
                clima      = clima_sel

            cq1, cq2, cq3 = st.columns(3)
            equipe_pm        = cq1.text_input("Equipe PM", placeholder="VTR, Comandante...", key="equipe_pm")
            equipe_pc        = cq2.text_input("Equipe PC", placeholder="Investigador, Delegado...", key="equipe_pc")
            autoridade_local = cq3.text_input("Autoridade no Local", placeholder="Nome da autoridade...", key="autoridade_local")


    with st.container():
# ══ SEÇÃO: DAS VÍTIMAS (Item 3.3) ═════════════════════
        st.markdown('<div class="section-title">👥 &nbsp; Das Vítimas (Item 3.3)</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>', unsafe_allow_html=True)
        
            vitimas_list = st.session_state.get("vitimas", [{"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "", "naturalidade": "", "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "", "fenomenos": "", "lesoes": [""]}])
        
            for i, vit in enumerate(vitimas_list):
                st.markdown(f"##### 👤 Vítima #{i+1}")
        
                st.markdown("<p style='font-size:12px; font-weight:700; color:#1a3a6e; margin-bottom:8px;'>IDENTIFICAÇÃO</p>", unsafe_allow_html=True)
                c_nome, c_doc, c_sexo = st.columns([2, 1, 1])
                vit["nome"]  = c_nome.text_input("Nome completo", value=vit.get("nome", ""), key=f"vit_nome_{i}")
                vit["documento"] = c_doc.text_input("Nº Documento (RG/CPF)", value=vit.get("documento", ""), key=f"vit_doc_{i}", placeholder="RG ou CPF...")
                vit["sexo"]  = c_sexo.selectbox("Sexo", ["", "Masculino", "Feminino", "Outro"], index=["", "Masculino", "Feminino", "Outro"].index(vit.get("sexo", "")), key=f"vit_sexo_{i}")
        
                c_dob, c_fili, c_nat = st.columns([1, 2, 1])
                dob_val = vit.get("data_nascimento", None)
                if not isinstance(dob_val, date) and dob_val is not None:
                    dob_val = date.today()
                vit["data_nascimento"] = c_dob.date_input("Data de Nascimento", value=dob_val, key=f"vit_dob_{i}")
                vit["filicao"] = c_fili.text_input("Filiação", value=vit.get("filicao", ""), key=f"vit_fili_{i}", placeholder="Nome da mãe e/ou pai...")
                vit["naturalidade"] = c_nat.text_input("Naturalidade", value=vit.get("naturalidade", ""), key=f"vit_nat_{i}", placeholder="Cidade-UF...")
        
                st.markdown("<p style='font-size:12px; font-weight:700; color:#1a3a6e; margin:12px 0 8px 0;'>VESTES E PERTENCES</p>", unsafe_allow_html=True)
                c_vest, c_pert = st.columns(2)
                vit["vestes"] = c_vest.text_area("Vestes", value=vit.get("vestes", ""), key=f"vit_vest_{i}", height=70, placeholder="Descrição das roupas que usava...")
                vit["pertences"] = c_pert.text_area("Pertences", value=vit.get("pertences", ""), key=f"vit_pert_{i}", height=70, placeholder="Celular, chaves, carteira, joias, etc...")
        
                st.markdown("<p style='font-size:12px; font-weight:700; color:#1a3a6e; margin:12px 0 8px 0;'>POSIÇÃO E LOCALIZAÇÃO (IN SITU)</p>", unsafe_allow_html=True)
                c_loc, c_pos = st.columns(2)
                vit["localizacao"] = c_loc.text_area("Localização", value=vit.get("localizacao", ""), key=f"vit_loc_{i}", height=70, placeholder="Ex: No compartimento traseiro do carro...")
                vit["posicao"] = c_pos.text_area("Posição", value=vit.get("posicao", ""), key=f"vit_pos_{i}", height=70, placeholder="Ex: Decúbito lateral direito...")
        
                c_cab, c_memb = st.columns(2)
                vit["cabeca"] = c_cab.text_input("Cabeça", value=vit.get("cabeca", ""), key=f"vit_cab_{i}", placeholder="Ex: Voltada a leste...")
                vit["membros"] = c_memb.text_input("Membros", value=vit.get("membros", ""), key=f"vit_memb_{i}", placeholder="Ex: Membros superiores fletidos...")
        
                st.markdown("<p style='font-size:12px; font-weight:700; color:#1a3a6e; margin:12px 0 8px 0;'>EXAME PERINECROSCÓPICO E ESTADO DE CONSERVAÇÃO</p>", unsafe_allow_html=True)
                vit["fenomenos"] = st.text_input("Fenômenos Transformativos", value=vit.get("fenomenos", ""), key=f"vit_fen_{i}", placeholder="Ex: Mancha verde abdominal, rigidez muscular...")
        
                st.markdown("<p style='font-size:11px; font-weight:600; color:#64748b; margin-top:8px;'>LESÕES CONSTATADAS</p>", unsafe_allow_html=True)
                les_list = vit.get("lesoes", [""])
                for j, lesao in enumerate(les_list):
                    c_les, c_rem_l = st.columns([12, 1])
                    les_list[j] = c_les.text_input(f"Lesão #{j+1}", value=lesao, key=f"vit_{i}_les_{j}", label_visibility="collapsed", placeholder=f"Descrição da lesão #{j+1}...")
                    if len(les_list) > 1:
                        if c_rem_l.button("❌", key=f"vit_{i}_rem_les_{j}", help="Remover esta lesão"):
                            les_list.pop(j)
                            st.rerun()
                vit["lesoes"] = les_list
        
                if st.button("➕ Adicionar Lesão", key=f"vit_{i}_add_les"):
                    vit.setdefault("lesoes", [""]).append("")
                    st.rerun()
        
                if len(vitimas_list) > 1:
                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                    if st.button(f"🗑️ Remover Vítima #{i+1}", key=f"remove_vit_{i}", type="secondary"):
                        st.session_state.vitimas.pop(i)
                        st.rerun()
                st.markdown("<hr style='margin:20px 0; border-color:#2563eb;'>", unsafe_allow_html=True)
        
            if st.button("➕ Adicionar Outra Vítima", key="btn_add_vitima"):
                st.session_state.vitimas.append({"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "", "naturalidade": "", "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "", "fenomenos": "", "lesoes": [""]})
                st.rerun()



    with st.container():
        st.markdown('<div class="section-title">🔍 &nbsp; Dos Vestígios</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>', unsafe_allow_html=True)
            
            manchas_dict = {
                "1. FORMAÇÃO PASSIVA": ["Gotejamento Isolado", "Gotejamento Sucessivo", "Gotejamento Estático", "Gotejamento Dinâmico", "Escorrimento", "Saturação", "Empoçamento"],
                "2. FORMAÇÃO ATIVA": ["Espargimento Por Impacto", "Espargimento Por Expiração", "Espargimento Por Êmese", "Espargimento De Retorno", "Espargimento De Saída", "Transferência Por Contato", "Transferência Por Arrastamento", "Inercial Por Desprendimento Centrífugo", "Inercial Por Parada", "Projeção por Diferença de Pressão", "Espalhamento de Volume em Queda Livre"],
                "3. ALTERADAS": ["Por Arrastamento", "Silhueta de Mancha", "Por Limpeza", "Por Inseto"],
                "4. AUSÊNCIA": ["Área de Sombra", "Sangue Latente"]
            }
            
            if "vestigios" not in st.session_state or not st.session_state.vestigios:
                st.session_state["vestigios"] = [{"tipo": "", "categoria": "", "subtipo": "", "localizacao": "", "descricao": "", "elemento_tipo": "", "quantidade": "", "envelope": "", "arma_tipo": "", "municiada": "", "marca": "", "cor": "", "funcionamento": "", "instrumento_tipo": ""}]
            vestigios_list = st.session_state["vestigios"]
            
            for i, vest in enumerate(vestigios_list):
                st.markdown(f"##### 🔎 Vestígio #{i+1}")
                tipo_opts = ["", "Manchas de Sangue", "Elemento Balístico", "Arma de Fogo", "Aparelho Telefônico", "Instrumento Lesivo", "Impressão Datiloscópica", "Outros"]
                tipo_idx = tipo_opts.index(vest.get("tipo", "")) if vest.get("tipo", "") in tipo_opts else 0
                vest["tipo"] = st.selectbox("Tipo de Vestígio", tipo_opts, index=tipo_idx, key=f"vest_tipo_{i}")
                
                if vest["tipo"] == "Manchas de Sangue":
                    c_cat, c_sub = st.columns(2)
                    cat_opts = [""] + list(manchas_dict.keys())
                    cat_idx = cat_opts.index(vest.get("categoria", "")) if vest.get("categoria", "") in cat_opts else 0
                    vest["categoria"] = c_cat.selectbox("Categoria", cat_opts, index=cat_idx, key=f"vest_cat_{i}")
                    sub_opts = [""] + manchas_dict.get(vest.get("categoria", ""), [])
                    sub_idx = sub_opts.index(vest.get("subtipo", "")) if vest.get("subtipo", "") in sub_opts else 0
                    vest["subtipo"] = c_sub.selectbox("Subtipo", sub_opts, index=sub_idx, key=f"vest_sub_{i}")
                    if vest["subtipo"]:
                        image_path = os.path.join(os.path.dirname(__file__), "imagens_manchas", f"{vest['subtipo']}.jpg")
                        if os.path.exists(image_path):
                            st.image(image_path, width=300, caption=f"Guia Visual: {vest['subtipo']}")
                        else:
                            st.caption(f"(Dica: Salve o recorte como 'imagens_manchas\\{vest['subtipo']}.jpg' para ver o guia visual aqui)")

                    vest["localizacao"] = st.text_input("Localização", value=vest.get("localizacao", ""), key=f"vest_loc_{i}", placeholder="Ex: Parede leste...")
                    vest["descricao"] = st.text_area("Descrição Livre", value=vest.get("descricao", ""), key=f"vest_desc_{i}", height=70)
                
                elif vest["tipo"] == "Elemento Balístico":
                    c_tipo, c_qtd = st.columns([3,1])
                    elem_opts = ["", "Estojo", "Munição", "Jaqueta", "Projétil"]
                    elem_idx = elem_opts.index(vest.get("elemento_tipo", "")) if vest.get("elemento_tipo", "") in elem_opts else 0
                    vest["elemento_tipo"] = c_tipo.selectbox("Tipo de Elemento", elem_opts, index=elem_idx, key=f"vest_elem_{i}")
                    vest["quantidade"] = c_qtd.text_input("Quantidade", value=vest.get("quantidade", ""), key=f"vest_qtd_{i}")
                    vest["envelope"] = st.text_input("Número Envelope de Segurança", value=vest.get("envelope", ""), key=f"vest_env_{i}")
                    vest["descricao"] = st.text_area("Descrição (geral)", value=vest.get("descricao", ""), key=f"vest_desc_{i}", height=70)
                    
                elif vest["tipo"] == "Arma de Fogo":
                    c_arma, c_mun = st.columns([3,1])
                    arma_opts = ["", "Revólver", "Espingarda", "Rifle", "Pistola", "Caseira"]
                    arma_idx = arma_opts.index(vest.get("arma_tipo", "")) if vest.get("arma_tipo", "") in arma_opts else 0
                    vest["arma_tipo"] = c_arma.selectbox("Tipo de Arma", arma_opts, index=arma_idx, key=f"vest_arma_{i}")
                    mun_opts = ["", "Sim", "Não"]
                    mun_idx = mun_opts.index(vest.get("municiada", "")) if vest.get("municiada", "") in mun_opts else 0
                    vest["municiada"] = c_mun.selectbox("Municiada", mun_opts, index=mun_idx, key=f"vest_mun_{i}")
                    vest["envelope"] = st.text_input("Número Envelope de Segurança", value=vest.get("envelope", ""), key=f"vest_env_{i}")
                    vest["descricao"] = st.text_area("Descrição", value=vest.get("descricao", ""), key=f"vest_desc_{i}", height=70)
                
                elif vest["tipo"] == "Outros":
                    vest["descricao"] = st.text_area("Descrição", value=vest.get("descricao", ""), key=f"vest_desc_{i}", height=70, placeholder="Descreva o vestígio...")
                
                elif vest["tipo"] == "Impressão Datiloscópica":
                    vest["descricao"] = st.text_area("Descrição", value=vest.get("descricao", ""), key=f"vest_desc_{i}", height=70, placeholder="Descreva os detalhes da impressão datiloscópica, superfície, revelação, etc.")
                
                elif vest["tipo"] == "Instrumento Lesivo":
                    vest["instrumento_tipo"] = st.text_input("Tipo (ex: Faca, pedaço de madeira, corda)", value=vest.get("instrumento_tipo", ""), key=f"vest_inst_{i}")
                    vest["envelope"] = st.text_input("Número Envelope de Segurança", value=vest.get("envelope", ""), key=f"vest_env_{i}")
                    vest["descricao"] = st.text_area("Descrição", value=vest.get("descricao", ""), key=f"vest_desc_{i}", height=70)
                
                elif vest["tipo"] == "Aparelho Telefônico":
                    c_marca, c_cor = st.columns(2)
                    vest["marca"] = c_marca.text_input("Marca ou Modelo", value=vest.get("marca", ""), key=f"vest_marca_{i}")
                    vest["cor"] = c_cor.text_input("Cor", value=vest.get("cor", ""), key=f"vest_cor_{i}")
                    c_func, c_env = st.columns(2)
                    func_opts = ["", "Ligado", "Desligado"]
                    func_idx = func_opts.index(vest.get("funcionamento", "")) if vest.get("funcionamento", "") in func_opts else 0
                    vest["funcionamento"] = c_func.selectbox("Funcionamento", func_opts, index=func_idx, key=f"vest_func_{i}")
                    vest["envelope"] = c_env.text_input("Número Envelope de Segurança", value=vest.get("envelope", ""), key=f"vest_env_{i}")
                    vest["descricao"] = st.text_area("Descrição", value=vest.get("descricao", ""), key=f"vest_desc_{i}", height=70)
                
                if st.button(f"🗑️ Remover Vestígio #{i+1}", key=f"remove_vest_{i}", type="secondary"):
                    st.session_state.vestigios.pop(i)
                    st.rerun()
                st.markdown("<hr style='margin:20px 0; border-color:#2563eb;'>", unsafe_allow_html=True)
                
            if st.button("➕ Adicionar Vestígio", key="btn_add_vest"):
                if "vestigios" not in st.session_state:
                    st.session_state.vestigios = []
                st.session_state.vestigios.append({"tipo": "", "categoria": "", "subtipo": "", "localizacao": "", "descricao": "", "elemento_tipo": "", "quantidade": "", "envelope": "", "arma_tipo": "", "municiada": "", "marca": "", "cor": "", "funcionamento": "", "instrumento_tipo": ""})
                st.rerun()



    # ══ BARRA DE AÇÕES (RODAPÉ) ══════════════════════════
    
    with st.container():
        st.markdown('<div class="section-title">📝 &nbsp; 4. Considerações Técnicas</div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<span class="custom-border-marker"></span>', unsafe_allow_html=True)
        
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
            
            iso_estado = st.selectbox("Estado do Isolamento", opcoes_isolamento, index=iso_idx, key="iso_estado")
        
            # Condicionais
            st.session_state["iso_meio"] = st.session_state.get("iso_meio", "")
            st.session_state["iso_alteracao"] = st.session_state.get("iso_alteracao", "")
            st.session_state["iso_falha"] = st.session_state.get("iso_falha", "")
        
            if iso_estado.startswith("1"):
                st.text_input("Meio de Isolamento", value=st.session_state["iso_meio"], key="iso_meio_ui", help="Ex: fita zebrada, cordão de isolamento, guarnição física da PM")
                st.session_state["iso_meio"] = st.session_state.get("iso_meio_ui", "")
            elif iso_estado.startswith("2"):
                st.text_input("Meio de Isolamento", value=st.session_state["iso_meio"], key="iso_meio_ui")
                st.text_area("Descrever Alteração", value=st.session_state["iso_alteracao"], key="iso_alteracao_ui", help="Ex: movimentação da vítima por equipes de socorro")
                st.session_state["iso_meio"] = st.session_state.get("iso_meio_ui", "")
                st.session_state["iso_alteracao"] = st.session_state.get("iso_alteracao_ui", "")
            elif iso_estado.startswith("3"):
                st.text_input("Descrever Falha no Isolamento", value=st.session_state["iso_falha"], key="iso_falha_ui", help="Ex: ausência de barreiras em rotas de fuga")
                st.session_state["iso_falha"] = st.session_state.get("iso_falha_ui", "")
            elif iso_estado.startswith("4"):
                st.text_input("Descrever Falha no Isolamento", value=st.session_state["iso_falha"], key="iso_falha_ui", help="Ex: coibir o fluxo de pessoas não autorizadas")
                st.text_area("Descrever Alteração", value=st.session_state["iso_alteracao"], key="iso_alteracao_ui", help="Ex: marcas de pneus sobrepostas a manchas de sangue")
                st.session_state["iso_falha"] = st.session_state.get("iso_falha_ui", "")
                st.session_state["iso_alteracao"] = st.session_state.get("iso_alteracao_ui", "")
            elif iso_estado.startswith("5"):
                st.text_area("Descrever Alteração (Devassado)", value=st.session_state["iso_alteracao"], key="iso_alteracao_ui")
                st.session_state["iso_alteracao"] = st.session_state.get("iso_alteracao_ui", "")

            st.divider()
            st.subheader("c) Laudo de Necropsia (IML)")
            c_iml1, c_iml2 = st.columns(2)
            with c_iml1:
                st.session_state["numero_laudo_necropsia"] = st.text_input("Número Laudo IML", value=st.session_state.get("numero_laudo_necropsia", ""), placeholder="Ex: 123/2026 - IML")
            with c_iml2:
                st.session_state["resultado_laudo_IML"] = st.text_input("Resultado do Laudo IML", value=st.session_state.get("resultado_laudo_IML", ""), placeholder="Ex: traumatismo cranioencefálico por PAF")

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
            
            inst_acao = st.selectbox("Tipo de Ação do Instrumento", opcoes_inst, index=inst_idx, key="inst_acao")
        
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
                st.session_state["inst_agente"] = st.text_input("Agente Compatível", value=st.session_state.get("inst_agente", ""), placeholder=f"Ex: {', '.join(sug[:3])}")
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
        
            st.session_state["inst_extra"] = st.text_area("Achados Extras / Observações do Instrumento", value=st.session_state.get("inst_extra", ""), help=placeholders_extras.get(inst_key, ""))
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    with st.container(border=True):
        ba1, ba2 = st.columns(2)

        with ba1:
            if st.button("👁️  Visualizar Laudo", use_container_width=True):
                st.session_state.preview_open = True
                st.rerun()

        with ba2:
            gerar_clicked = st.button("🚀  Gerar Laudo (.docx)", type="primary", use_container_width=True)

        if gerar_clicked:
            try:
                import os
                template_path = os.path.join(os.path.dirname(__file__), "Modelo Laudo.docx")
                doc = Document(template_path)
                
                # Fetch dynamically selected data
                dp_val = st.session_state.get("data_pericia_input", date.today())
                da_val = st.session_state.get("data_atendimento_input", date.today())
                horario_val = st.session_state.get("horario", None)
                horario_a_val = st.session_state.get("horario_atendimento", None)

                from datetime import datetime, date, time as dt_time
                def safe_date_obj(d):
                    if isinstance(d, str):
                        try: return datetime.strptime(d, "%Y-%m-%d").date()
                        except: pass
                    return d

                def safe_time_str(t):
                    if not t: return "____"
                    if isinstance(t, str): return t[:5]
                    try: return t.strftime("%H:%M")
                    except: return str(t)

                dp_val = safe_date_obj(dp_val)
                da_val = safe_date_obj(da_val)

                dp_extenso = data_extenso(dp_val)
                dp_simples  = data_simples(dp_val)
                da_simples  = data_simples(da_val)
                h_str       = safe_time_str(horario_val)
                ha_str      = safe_time_str(horario_a_val)

                tipo_local_val = st.session_state.get("tipo_local", "")

                vitimas_list = st.session_state.get("vitimas", [{"nome": "", "cad": "", "documento": "", "sexo": "", "data_nascimento": None, "filicao": "", "naturalidade": "", "vestes": "", "pertences": "", "localizacao": "", "posicao": "", "cabeca": "", "membros": "", "fenomenos": "", "lesoes": [""]}])
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
                    individual_vars[f"numero_documento{suffix}"] = vt.get("documento", "")
                    individual_vars[f"sexo{suffix}"] = vt.get("sexo", "")
                    individual_vars[f"data_nascimento{suffix}"] = dob_str
                    individual_vars[f"filicao{suffix}"] = vt.get("filicao", "")
                    individual_vars[f"naturalidade{suffix}"] = vt.get("naturalidade", "")
                    individual_vars[f"veste{suffix}"] = vt.get("vestes", "")
                    individual_vars[f"pertence{suffix}"] = vt.get("pertences", "")
                    individual_vars[f"localização{suffix}"] = vt.get("localizacao", "")
                    individual_vars[f"posição{suffix}"] = vt.get("posicao", "")
                    individual_vars[f"cabeça{suffix}"] = vt.get("cabeca", "")
                    individual_vars[membros_key] = vt.get("membros", "")
                    individual_vars[f"fenômenos_transformativo{suffix}"] = vt.get("fenomenos", "")
                    individual_vars[f"Lesao1{suffix}"] = lesao1_val
                    individual_vars[f"Lesao2{suffix}"] = lesao2_val
                
                    lesoes_txt_list = []
                    for les in lesoes_list:
                        if les.strip():
                            lesoes_txt_list.append(f"• {les.strip()}")
                    lesoes_text = "\n".join(lesoes_txt_list) if lesoes_txt_list else "• Não descritas"
                
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
                vestigios_detalhes = "\n\n".join(vestigios_detalhes_list) if vestigios_detalhes_list else "Não foram descritos vestígios."


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
                iso_text = iso_text.replace("{meio_de_isolamento}", st.session_state.get("iso_meio") or "________")
                iso_text = iso_text.replace("{descrever_alteracao}", st.session_state.get("iso_alteracao") or "________")
                iso_text = iso_text.replace("{descrever_falha_isolamento}", st.session_state.get("iso_falha") or "________")
                
                VARS["isolamento_detalhes"] = iso_text

                # Map IML variables
                VARS["número_laudo_necropsia"] = st.session_state.get("numero_laudo_necropsia") or "________"
                VARS["numero_laudo_necropsia"] = st.session_state.get("numero_laudo_necropsia") or "________"
                VARS["resultado_laudo_IML"] = st.session_state.get("resultado_laudo_IML") or "________"
                
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
                # Also map to target paragraph if tag is missing in docx
                VARS["d. As lesões constatadas nas vítimas são características daquelas produzidas por instrumento em ação perfurocontundente, inequivocamente produzidas por projéteis de arma de fogo (PAF), com a recuperação de um projétil em cada cadáver durante a necrópsia."] = f"d. {inst_text}"


                
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
                                    p_new = para.insert_paragraph_before(style=para.style)
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
                                    for r in para.runs[1:]: r.text = ""
                                else:
                                    para.add_run(new)

                for para in doc.paragraphs: replace_para(para, VARS)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs: replace_para(para, VARS)
                for section in doc.sections:
                    if section.header:
                        for para in section.header.paragraphs: replace_para(para, VARS)
                    if section.footer:
                        for para in section.footer.paragraphs: replace_para(para, VARS)

                # --- Inserção de Fotografias no Apêndice A ---
                fotos_para_laudo = [f for f in st.session_state.get("fotos", []) if f.get("incluir", True)]
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
                        apendice_p = doc.add_paragraph("APÊNDICE A – TOMADAS FOTOGRÁFICAS")
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
                            st.warning(f"Não foi possível inserir a fotografia {f_idx}: {str(img_err)}")

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
