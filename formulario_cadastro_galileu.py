# -*- coding: utf-8 -*-
"""
Formulário de Cadastro Galileu - Sistema LAF
Formulário standalone para preenchimento dos dados que o robô utiliza.
- Part 1: Dados gerais da solicitação (etapas 4, 5 e 7)
- Part 2: Evidências (escolha tipo + campos por tipo, similar ao Gerar Laudo)

Executar: streamlit run formulario_cadastro_galileu.py
"""
import os
import json
from datetime import datetime
import streamlit as st

# =============================================================================
# CONSTANTES (sincronizadas com robo_galileu_cadastro_teste.py)
# =============================================================================

TIPOS_DOCUMENTO = (
    "GUIA - Guia", "GCOMP - Guia Complementar", "OCO - Ocorrência",
    "BO - Boletim de Ocorrência", "OF - Ofício", "MDJ - Mandado Judicial",
    "PROC - Processo", "RDIO - Rádio", "CELP - Celular Plantão",
    "IP - Inquérito Policial", "INT - Requisição Interna",
    "REQ - Requisição de Exame Pericial", "NPROC - Número do Procedimento",
    "CI - Comunicação Interna", "SEI - Processo SEI",
)

CARGOS_AUTORIDADE = (
    "Delegado(a) de Polícia Civil", "Delegado (a) de Polícia Federal",
    "Promotor de Justiça", "Juiz", "Perito(a) Oficial",
    "Perito(a) Oficial - Chefe de Serviço", "Perito(a) Oficial - Diretor",
    "Perito(a) Oficial - Perito Geral", "Comandante", "Coordenador",
    "Procurador de Justiça", "Secretário de Estado",
)

INVOLUCROS = (
    "Ampola", "Balde", "Blister", "Caixa de Papelão", "Coletor à vácuo",
    "Coletor universal", "Embalagem para Suabe", "Embalagem Plástica",
    "Envelope de papel", "Eppendorf", "Fita de empacotamento",
    "Fita ou película adesiva", "Frasco de Penicilina", "Frasco de Vidro",
    "Frasco Plástico", "Galão", "Garrafa de vidro", "Garrafão",
    "Garrafa pet", "Garrafa Plástica", "Invólucro de Papel", "Jarro",
    "Lacre de Evidência", "Mala", "Material avulso", "Mochila",
    "Não Embalada", "Outros", "Papel Alumínio", "Plástico Filme",
    "Pote Plástico", "PVC", "Sacola Plástica", "Saco Plástico",
    "Seringa", "Tubo de vidro", "Zip lock",
)

SETORES_DESTINO = (
    "Seção de Química Forense",
    "QUI ICRIM IMP - Seção de Química Forense",
)

TIPOS_EXAME = (
    "Químico para Identificação de THC - Definitivo",
    "Químico para Identificação de THC - Preliminar",
)

TIPOS_MATERIAL = ("Vegetal", "Líquido", "Sólido", "Pó", "Outros")
SUBSTANCIAS = ("Maconha", "Cocaína", "Crack", "THC", "Outros")
UNIDADES_MEDIDA = ("mg", "g", "ml", "L")


def _dir_script():
    return os.path.dirname(os.path.abspath(__file__))


def _arquivo_cadastros():
    return os.path.join(_dir_script(), "cadastros_galileu.json")


def _carregar_cadastros():
    path = _arquivo_cadastros()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _salvar_cadastros(lista):
    path = _arquivo_cadastros()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(lista, f, ensure_ascii=False, indent=2)


def _chave_item(dados_gerais):
    # Chave estável para update: número + ano + tipo_documento.
    return "|".join(
        [
            str(dados_gerais.get("numero", "")).strip(),
            str(dados_gerais.get("ano", "")).strip(),
            str(dados_gerais.get("tipo_documento", "")).strip(),
        ]
    )


def carregar_opcoes(arquivo, excluir=("[Filtrar por ]", "[Selecione", "Selecione ")):
    """Carrega opções de um arquivo (uma por linha)."""
    caminho = os.path.join(_dir_script(), arquivo)
    if not os.path.exists(caminho):
        return []
    with open(caminho, "r", encoding="utf-8") as f:
        linhas = [ln.strip() for ln in f.readlines() if ln.strip()]
    return [x for x in linhas if not any(x.startswith(p) for p in excluir)]


def carregar_unidades():
    """Carrega unidades policial e destino dos arquivos."""
    raw_pol = carregar_opcoes("unidade_policial_galileu.txt")
    raw_dest = carregar_opcoes("unidades_destino_galileu.txt")
    # Filtra apenas linhas no formato "COD - Descrição" (unidades reais)
    def filtrar_unidades(lst):
        return [x for x in lst if " - " in x and len(x) > 10 and not x.startswith("[")]
    unidade_pol = filtrar_unidades(raw_pol) if raw_pol else []
    unidade_dest = filtrar_unidades(raw_dest) if raw_dest else []
    if not unidade_pol:
        unidade_pol = ["10DP SLZ - 10º Distrito de Polícia da Capital Coroadinho"]
    if not unidade_dest:
        unidade_dest = ["QUI ICRIM IMP - Seção de Química Forense"]
    return unidade_pol, unidade_dest


# =============================================================================
# DADOS DE TESTE (preenchimento geral)
# =============================================================================

DADOS_TESTE = {
    "setor_destino": "Seção de Química Forense",
    "tipo_exame": "Químico para Identificação de THC - Definitivo",
    "tipo_documento": "GUIA - Guia",
    "numero": "88888",
    "s_n": False,
    "ano": "2026",
    "referencia": "Ref. teste formulário LAF - Verificação de inserção",
    "unidade_requisitante": "10DP SLZ - 10º Distrito de Polícia da Capital Coroadinho",
    "unidade_destino": "QUI ICRIM IMP - Seção de Química Forense",
    "cargo": "Delegado(a) de Polícia Civil",
    "autoridade": "Dr. João Silva Santos",
    "data_documento": "05/03/2026",
    "data_hora_recebimento": "05/03/2026 14:30",
}


def preencher_dados_teste():
    """Preenche todos os campos com dados de teste e adiciona evidências exemplo."""
    for k, v in DADOS_TESTE.items():
        st.session_state[k] = v
    st.session_state["evidencias_cadastro"] = [
        {
            "tipo": "Drogas de Abuso",
            "dados": {
                "tipo_material": "Vegetal",
                "quantidade_embalagens": "5",
                "involucro": "Sacola Plástica",
                "substancia": "Maconha",
                "qtd_apresentada_material": "200",
                "unidade_apresentada": "mg",
                "qtd_massa_liquida": "300",
                "unidade_massa_liquida": "mg",
                "qtd_utilizada": "150",
                "unidade_utilizada": "mg",
                "servico": "Química",
                "exame": "THC",
            },
        },
        {
            "tipo": "Pessoa Envolvida",
            "dados": {
                "tipo_envolvimento": "Implicado",
                "tipo_implicado": "Investigado",
                "nome": "Maria Santos Oliveira",
            },
        },
    ]


# =============================================================================
# INICIALIZAÇÃO
# =============================================================================

st.set_page_config(page_title="Cadastro Galileu", page_icon="📋", layout="wide")

if "evidencias_cadastro" not in st.session_state:
    st.session_state["evidencias_cadastro"] = []  # lista de dicts: {tipo, dados}

unidade_pol, unidade_dest = carregar_unidades()

# Garantir que opções de teste existam nas listas
if unidade_pol and DADOS_TESTE["unidade_requisitante"] not in unidade_pol:
    unidade_pol = [DADOS_TESTE["unidade_requisitante"]] + list(unidade_pol)
if unidade_dest and DADOS_TESTE["unidade_destino"] not in unidade_dest:
    unidade_dest = [DADOS_TESTE["unidade_destino"]] + list(unidade_dest)

# =============================================================================
# PART 1: DADOS GERAIS DA SOLICITAÇÃO
# =============================================================================

c_tit, c_btn_teste = st.columns([4, 1])
with c_tit:
    st.markdown("## 📋 Formulário de Cadastro Galileu")
    st.caption("Campos que o robô preenche no Galileu. Distinção: **seleção** vs **digitação**.")
with c_btn_teste:
    st.write("")
    if st.button("🧪 Preencher com dados de teste", type="secondary", use_container_width=True, key="btn_teste"):
        preencher_dados_teste()
        st.toast("Formulário preenchido com dados de teste! Verifique a inserção.", icon="✅")
        st.rerun()

st.markdown("---")
st.markdown("### 1. Dados Gerais da Solicitação")

st.markdown("**Etapas 4 e 5 – Seleções iniciais**")
c_s1, c_s2 = st.columns(2)
with c_s1:
    setor_destino = st.selectbox(
        "Setor de Destino *(seleção)*",
        options=SETORES_DESTINO,
        key="setor_destino",
        help="Dropdown no Galileu",
    )
with c_s2:
    tipo_exame = st.selectbox(
        "Tipo de Exame *(seleção)*",
        options=TIPOS_EXAME,
        key="tipo_exame",
        help="Dropdown no Galileu",
    )

st.markdown("**Etapa 7 – Formulário da ocorrência**")
st.markdown("*Campos de seleção (dropdown)*")
c1, c2, c3 = st.columns(3)
with c1:
    tipo_documento = st.selectbox("Tipo de Documento", options=TIPOS_DOCUMENTO, key="tipo_doc")
with c2:
    unidade_requisitante = st.selectbox(
        "Unidade Requisitante",
        options=unidade_pol,
        key="unidade_req",
    )
with c3:
    unidade_destino = st.selectbox(
        "Unidade Destino",
        options=unidade_dest,
        key="unidade_dest",
    )

c4, c5, c6 = st.columns(3)
with c4:
    cargo_autoridade = st.selectbox("Cargo da Autoridade Requisitante", options=CARGOS_AUTORIDADE, key="cargo")
with c5:
    st.write("")  # alinhamento
with c6:
    st.write("")

st.markdown("*Campos de digitação (texto/número)*")
c7, c8, c9, c10 = st.columns(4)
with c7:
    numero = st.text_input("Número *(digitar)*", placeholder="Ex: 12345", key="numero")
with c8:
    s_n_desconhecido = st.checkbox("S/N? (desconhecido)", key="s_n")
with c9:
    ano = st.text_input("Ano *(digitar)*", placeholder="Ex: 2026", key="ano")
with c10:
    referencia = st.text_input("Referência *(digitar)*", placeholder="Ex: Ref. processo", key="ref")

c11, c12 = st.columns(2)
with c11:
    autoridade = st.text_input("Autoridade Requisitante *(digitar)*", placeholder="Nome da autoridade", key="autoridade")
with c12:
    data_documento = st.text_input("Data do Documento *(digitar)*", placeholder="dd/mm/aaaa", key="data_doc")

data_hora_recebimento = st.text_input(
    "Data e Hora de Recebimento *(digitar)*",
    placeholder="dd/mm/aaaa HH:mm",
    key="data_hora_rec",
)

# Resumo Part 1
dados_gerais = {
    "setor_destino": setor_destino,
    "tipo_exame": tipo_exame,
    "tipo_documento": tipo_documento,
    "numero": numero,
    "s_n_desconhecido": s_n_desconhecido,
    "ano": ano,
    "referencia": referencia,
    "unidade_requisitante": unidade_requisitante,
    "unidade_destino": unidade_destino,
    "cargo_autoridade": cargo_autoridade,
    "autoridade": autoridade,
    "data_documento": data_documento,
    "data_hora_recebimento": data_hora_recebimento,
}

# =============================================================================
# PART 2: EVIDÊNCIAS (similar ao Gerar Laudo)
# =============================================================================

st.markdown("---")
st.markdown("### 2. Evidências / Vestígios")
st.caption("Escolha o tipo de evidência a adicionar e preencha os campos. Similar ao fluxo Gerar Laudo.")

c_tipo, c_btn = st.columns([3, 1])
with c_tipo:
    tipo_evidencia_sel = st.selectbox(
        "Tipo de evidência a adicionar",
        options=["Drogas de Abuso", "Pessoa Envolvida"],
        key="tipo_evidencia_novo",
    )
with c_btn:
    st.write("")
    st.write("")
    if st.button("➕ Adicionar Evidência", type="primary", use_container_width=True, key="btn_add_evid"):
        novo = {"tipo": tipo_evidencia_sel, "dados": {}}
        st.session_state["evidencias_cadastro"].append(novo)
        st.rerun()

# Lista de evidências já adicionadas
if st.session_state["evidencias_cadastro"]:
    st.markdown("#### Evidências adicionadas")
    indices_remover = []
    for idx, ev in enumerate(st.session_state["evidencias_cadastro"]):
        with st.expander(f"**{ev['tipo']}** #{idx + 1}", expanded=True):
            col_del, _ = st.columns([1, 5])
            with col_del:
                if st.button("🗑️ Remover", key=f"del_ev_{idx}"):
                    indices_remover.append(idx)

            if ev["tipo"] == "Drogas de Abuso":
                d = ev["dados"]
                # Campos do vestígio Drogas de Abuso
                c1, c2 = st.columns(2)
                with c1:
                    ev["dados"]["tipo_material"] = st.selectbox(
                        "Tipo de Material", TIPOS_MATERIAL,
                        index=TIPOS_MATERIAL.index(d.get("tipo_material", "Vegetal")) if d.get("tipo_material") in TIPOS_MATERIAL else 0,
                        key=f"ev_{idx}_tm",
                    )
                    ev["dados"]["quantidade_embalagens"] = st.text_input("Quantidade de Embalagens", value=d.get("quantidade_embalagens", "5"), key=f"ev_{idx}_qe")
                    ev["dados"]["involucro"] = st.selectbox("Invólucro", INVOLUCROS, index=INVOLUCROS.index(d["involucro"]) if d.get("involucro") in INVOLUCROS else 0, key=f"ev_{idx}_inv")
                    ev["dados"]["substancia"] = st.selectbox("Substância", SUBSTANCIAS, index=SUBSTANCIAS.index(d["substancia"]) if d.get("substancia") in SUBSTANCIAS else 0, key=f"ev_{idx}_sub")
                    ev["dados"]["qtd_apresentada_material"] = st.text_input("Qtd. Apresentada Material", value=d.get("qtd_apresentada_material", "200"), key=f"ev_{idx}_qam")
                with c2:
                    ev["dados"]["unidade_apresentada"] = st.selectbox("Unidade (apresentada)", UNIDADES_MEDIDA, index=UNIDADES_MEDIDA.index(d["unidade_apresentada"]) if d.get("unidade_apresentada") in UNIDADES_MEDIDA else 0, key=f"ev_{idx}_ua")
                    ev["dados"]["qtd_massa_liquida"] = st.text_input("Qtd. Massa Líquida", value=d.get("qtd_massa_liquida", "300"), key=f"ev_{idx}_qml")
                    ev["dados"]["unidade_massa_liquida"] = st.selectbox("Unidade Massa Líquida", UNIDADES_MEDIDA, index=UNIDADES_MEDIDA.index(d["unidade_massa_liquida"]) if d.get("unidade_massa_liquida") in UNIDADES_MEDIDA else 0, key=f"ev_{idx}_uml")
                    ev["dados"]["qtd_utilizada"] = st.text_input("Quantidade Utilizada", value=d.get("qtd_utilizada", "150"), key=f"ev_{idx}_qu")
                    ev["dados"]["unidade_utilizada"] = st.selectbox("Unidade Utilizada", UNIDADES_MEDIDA, index=UNIDADES_MEDIDA.index(d["unidade_utilizada"]) if d.get("unidade_utilizada") in UNIDADES_MEDIDA else 0, key=f"ev_{idx}_uu")
                ev["dados"]["servico"] = st.text_input("Serviço", value=d.get("servico", "Química"), key=f"ev_{idx}_serv")
                ev["dados"]["exame"] = st.text_input("Exame", value=d.get("exame", "THC"), key=f"ev_{idx}_exam")

            elif ev["tipo"] == "Pessoa Envolvida":
                d = ev["dados"]
                te_opts = ["Implicado", "Vítima", "Testemunha", "Outros"]
                ti_opts = ["Investigado", "Indiciado", "Outros"]
                ev["dados"]["tipo_envolvimento"] = st.selectbox(
                    "Tipo de Envolvimento", te_opts,
                    index=te_opts.index(d["tipo_envolvimento"]) if d.get("tipo_envolvimento") in te_opts else 0,
                    key=f"ev_{idx}_te",
                )
                ev["dados"]["tipo_implicado"] = st.selectbox(
                    "Tipo Implicado (se Implicado)", ti_opts,
                    index=ti_opts.index(d["tipo_implicado"]) if d.get("tipo_implicado") in ti_opts else 0,
                    key=f"ev_{idx}_ti",
                )
                ev["dados"]["nome"] = st.text_input("Nome", value=d.get("nome", ""), placeholder="Nome completo", key=f"ev_{idx}_nome")

    for i in sorted(indices_remover, reverse=True):
        st.session_state["evidencias_cadastro"].pop(i)
        st.rerun()
else:
    st.info("Nenhuma evidência adicionada. Selecione o tipo e clique em **Adicionar Evidência**.")

# =============================================================================
# EXPOR DADOS (para integração futura)
# =============================================================================

st.markdown("---")
st.markdown("### 3. Resumo / Exportação")
with st.expander("📄 Visualizar dados para o robô"):
    saida = {
        "dados_gerais": dados_gerais,
        "evidencias": st.session_state["evidencias_cadastro"],
    }
    st.json(saida)
    st.caption("Este JSON pode ser usado para alimentar o robô quando a integração for feita.")

# Fluxo separado: salvar/atualizar sempre, e gerar laudo opcional.
gerar_laudo = st.checkbox(
    "Gerar laudo após salvar/atualizar",
    value=False,
    key="ck_gerar_laudo",
    help="Desmarcado = apenas salva/atualiza o item.",
)

c_acoes_1, c_acoes_2 = st.columns(2)
with c_acoes_1:
    if st.button("💾 Salvar/Atualizar item", key="btn_salvar_atualizar", use_container_width=True):
        item = {
            "chave_item": _chave_item(dados_gerais),
            "atualizado_em": datetime.now().isoformat(timespec="seconds"),
            "dados_gerais": dados_gerais,
            "evidencias": st.session_state["evidencias_cadastro"],
            "gerar_laudo": bool(gerar_laudo),
        }
        cadastros = _carregar_cadastros()
        idx_existente = next(
            (i for i, x in enumerate(cadastros) if str(x.get("chave_item", "")) == item["chave_item"]),
            None,
        )
        if idx_existente is None:
            cadastros.append(item)
            acao = "salvo"
        else:
            cadastros[idx_existente] = item
            acao = "atualizado"
        _salvar_cadastros(cadastros)

        if gerar_laudo:
            st.success(f"Item {acao} com sucesso. Fluxo marcado para gerar laudo.")
        else:
            st.success(f"Item {acao} com sucesso. Apenas atualização (sem gerar laudo).")

with c_acoes_2:
    if st.button("🧪 Gerar laudo agora", key="btn_gerar_laudo", use_container_width=True):
        st.info("Geração de laudo acionada. Enquanto a integração não está ligada, o item também fica salvo/atualizado no cadastro local.")
