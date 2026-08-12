# -*- coding: utf-8 -*-
"""
Robô Galileu - Cadastro (teste)
Inverso da extração: cadastrar dados no Galileu a partir de valores informados.
- URL: galileu-treinamento (ambiente de teste)
- Login/senha: mesmos do robo_galileu
- headless=False para acompanhar no navegador
- Simulação com dados aleatórios (rodar pelo terminal)
"""
import os
import time
import sys
import random
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# =============================================================================
# CONFIGURAÇÃO
# =============================================================================
BASE_URL = "https://galileu.ssp.ma.gov.br"
URL_TREINAMENTO = "https://galileu.ssp.ma.gov.br/galileu-treinamento/"
URL_MODULO_CARTORIO = "https://galileu.ssp.ma.gov.br/galileu-treinamento/pages/cruds/pericia/moduloEntradaCartorio.faces#no-back-button"
USUARIO = "04360213310"
SENHA = "1810Pl@."

ARQUIVO_UNIDADES_DESTINO = "unidades_destino_galileu.txt"
ARQUIVO_UNIDADE_POLICIAL = "unidade_policial_galileu.txt"
ARQUIVO_TIPO_DOCUMENTO = "tipo_documento_galileu.txt"
ARQUIVO_CARGO_AUTORIDADE = "cargo_autoridade_galileu.txt"
ARQUIVO_INVOLUCRO = "involucro_galileu.txt"

# Valores válidos para Tipo de Documento (Galileu)
TIPOS_DOCUMENTO_VALIDOS = (
    "GUIA - Guia",
    "GCOMP - Guia Complementar",
    "OCO - Ocorrência",
    "BO - Boletim de Ocorrência",
    "OF - Ofício",
    "MDJ - Mandado Judicial",
    "PROC - Processo",
    "RDIO - Rádio",
    "CELP - Celular Plantão",
    "IP - Inquérito Policial",
    "INT - Requisição Interna",
    "REQ - Requisição de Exame Pericial",
    "NPROC - Número do Procedimento",
    "CI - Comunicação Interna",
    "SEI - Processo SEI",
)

# Valores válidos para Cargo da Autoridade Requisitante (obrigatório)
CARGOS_AUTORIDADE_VALIDOS = (
    "Delegado(a) de Polícia Civil",
    "Delegado (a) de Polícia Federal",
    "Promotor de Justiça",
    "Juiz",
    "Perito(a) Oficial",
    "Perito(a) Oficial - Chefe de Serviço",
    "Perito(a) Oficial - Diretor",
    "Perito(a) Oficial - Perito Geral",
    "Comandante",
    "Coordenador",
    "Procurador de Justiça",
    "Secretário de Estado",
)

# Valores válidos para Invólucro (apresentacaoMaterialVestigio)
INVOLUCROS_VALIDOS = (
    "Ampola", "Balde", "Blister", "Caixa de Papelão", "Coletor à vácuo", "Coletor universal",
    "Embalagem para Suabe", "Embalagem Plástica", "Envelope de papel", "Eppendorf",
    "Fita de empacotamento", "Fita ou película adesiva", "Frasco de Penicilina", "Frasco de Vidro",
    "Frasco Plástico", "Galão", "Garrafa de vidro", "Garrafão", "Garrafa pet", "Garrafa Plástica",
    "Invólucro de Papel", "Jarro", "Lacre de Evidência", "Mala", "Material avulso", "Mochila",
    "Não Embalada", "Outros", "Papel Alumínio", "Plástico Filme", "Pote Plástico", "PVC",
    "Sacola Plástica", "Saco Plástico", "Seringa", "Tubo de vidro", "Zip lock",
)

# Valores fixos para o teste de cadastro (ajustar conforme o formulário real)
DADOS_TESTE = {
    "ocorrencia": "999999",
    "setor_destino": "Seção de Química Forense",
    "tipo_exame": "Químico para Identificação de THC - Definitivo",
    "origem": "",
    "destino": "",
    "autoridade": "",
    "envolvidos": "",
    "evidencias": "",
}


def log(msg):
    print(msg, flush=True)


def _dir_script():
    return os.path.dirname(os.path.abspath(__file__))


def carregar_opcoes(arquivo):
    """Carrega opções de um arquivo extraído (uma por linha). Filtra placeholders."""
    caminho = os.path.join(_dir_script(), arquivo)
    if not os.path.exists(caminho):
        return []
    with open(caminho, "r", encoding="utf-8") as f:
        linhas = [ln.strip() for ln in f.readlines() if ln.strip()]
    excluir = ("[Filtrar por ]", "[Selecione", "Selecione ")
    return [x for x in linhas if not any(x.startswith(p) for p in excluir)]


# -----------------------------------------------------------------------------
# Geração de dados aleatórios para simulação
# -----------------------------------------------------------------------------
NOMES = ("João Silva", "Maria Santos", "Carlos Oliveira", "Ana Costa", "Pedro Ferreira",
         "Lucia Pereira", "Roberto Alves", "Fernanda Lima", "Antonio Souza", "Juliana Rocha")


def gerar_numero_aleatorio(min_val=1000, max_val=99999):
    return str(random.randint(min_val, max_val))


def gerar_ano_aleatorio(min_ano=2020, max_ano=None):
    max_ano = max_ano or datetime.now().year
    return str(random.randint(min_ano, max_ano))


def gerar_texto_aleatorio(prefixo="Ref", tamanho=8):
    chars = "abcdefghijklmnopqrstuvwxyz0123456789"
    return prefixo + "-" + "".join(random.choices(chars, k=tamanho))


def gerar_nome_aleatorio():
    return random.choice(NOMES)


def gerar_data_aleatoria():
    """Retorna data no formato DD/MM/YYYY."""
    dias_atras = random.randint(0, 365)
    d = datetime.now() - timedelta(days=dias_atras)
    return d.strftime("%d/%m/%Y")


def gerar_data_hora_aleatoria():
    """Retorna data e hora no formato DD/MM/YYYY HH:MM."""
    dias_atras = random.randint(0, 30)
    d = datetime.now() - timedelta(days=dias_atras)
    d = d.replace(hour=random.randint(8, 18), minute=random.choice([0, 15, 30, 45]))
    return d.strftime("%d/%m/%Y %H:%M")


def selecionar_dropdown_por_valor(page, seletor_label, valor):
    """Seleciona opção em dropdown PrimeFaces digitando o valor e Enter."""
    page.wait_for_selector(seletor_label, state="visible", timeout=5000)
    page.click(seletor_label)
    time.sleep(0.5)
    page.keyboard.type(valor, delay=50)
    time.sleep(0.3)
    page.keyboard.press("Enter")
    limpar_bloqueios(page)
    time.sleep(0.3)


def preencher_sem_scroll(page, seletor, valor, evita_focus=False):
    """
    Preenche input via JavaScript para evitar scroll automático.
    evita_focus=True: não dá focus (evita abrir date picker).
    """
    page.evaluate("""
        ([sel, val, semFocus]) => {
            const el = document.querySelector(sel);
            if (el) {
                if (!semFocus) el.focus();
                el.value = val;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('blur', { bubbles: true }));
            }
        }
    """, [seletor, str(valor), evita_focus])


# def extrair_dropdown(page, seletor_label, arquivo_saida, salvar_arquivo=True):
#     """Abre um dropdown e extrai todas as opções para padronização."""
#     page.wait_for_selector(seletor_label, state="visible", timeout=10000)
#     page.click(seletor_label)
#     time.sleep(0.8)
#     itens = page.locator(".ui-selectonemenu-panel li.ui-selectonemenu-item").all_inner_texts()
#     if salvar_arquivo and itens:
#         caminho = os.path.join(_dir_script(), arquivo_saida)
#         with open(caminho, "w", encoding="utf-8") as f:
#             for txt in itens:
#                 f.write(txt.strip() + "\n")
#         log(f"   {len(itens)} opcoes salvas em {arquivo_saida}")
#     page.keyboard.press("Escape")
#     time.sleep(0.3)
#     return itens
#
#
# def extrair_unidades_destino(page, salvar_arquivo=True):
#     """Extrai opções do dropdown Unidade Destino."""
#     return extrair_dropdown(page, "#unidadeDestino_label", ARQUIVO_UNIDADES_DESTINO, salvar_arquivo)


def limpar_bloqueios(page):
    """Remove overlays e fecha date pickers abertos."""
    try:
        page.wait_for_selector(".ui-widget-overlay", state="hidden", timeout=1000)
    except Exception:
        pass
    page.keyboard.press("Escape")
    time.sleep(0.2)
    page.keyboard.press("Escape")
    time.sleep(0.2)
    page.evaluate("""
        () => {
            document.querySelectorAll('.ui-widget-overlay').forEach(e => e.remove());
            document.querySelectorAll('.ui-dialog-mask').forEach(e => e.remove());
            document.querySelectorAll('.ui-datepicker').forEach(e => e.style.display = 'none');
            document.querySelectorAll('[id*="datepicker"]').forEach(e => { e.style.visibility = 'hidden'; });
        }
    """)
    time.sleep(0.15)


def preencher_formulario_cadastro(page):
    """Preenche o formulário 'Confirme os dados do documento' com dados aleatórios."""
    log("   [PREENCHENDO] Formulario com dados aleatorios...")
    dados = {}

    # Carregar opções (Tipo e Cargo usam listas fixas válidas)
    opcoes_tipo_doc = list(TIPOS_DOCUMENTO_VALIDOS)
    opcoes_unidade_pol = carregar_opcoes(ARQUIVO_UNIDADE_POLICIAL)
    opcoes_unidade_dest = carregar_opcoes(ARQUIVO_UNIDADES_DESTINO)
    opcoes_cargo = list(CARGOS_AUTORIDADE_VALIDOS)

    # Gerar dados aleatórios
    dados["tipo_documento"] = random.choice(opcoes_tipo_doc) if opcoes_tipo_doc else "GUIA - Guia"
    dados["numero"] = gerar_numero_aleatorio()
    dados["s_n"] = random.choice([True, False])
    dados["ano"] = gerar_ano_aleatorio()
    dados["referencia"] = gerar_texto_aleatorio()
    dados["unidade_requisitante"] = random.choice(opcoes_unidade_pol) if opcoes_unidade_pol else "10DP SLZ - 10º Distrito de Polícia da Capital Coroadinho"
    dados["unidade_destino"] = random.choice(opcoes_unidade_dest) if opcoes_unidade_dest else ""
    dados["cargo_autoridade"] = random.choice(opcoes_cargo) if opcoes_cargo else ""
    dados["autoridade"] = gerar_nome_aleatorio()
    dados["data_documento"] = gerar_data_aleatoria()
    dados["data_hora_recebimento"] = gerar_data_hora_aleatoria()

    for k, v in dados.items():
        log(f"      {k}: {v}")

    # 1. Tipo de Documento
    selecionar_dropdown_por_valor(page, "#tipoDocumentoPolicial_label", dados["tipo_documento"])
    limpar_bloqueios(page)
    time.sleep(1.5)  # aguarda possível atualização AJAX do PrimeFaces após seleção

    # 2. Número (via JS - evita scroll)
    page.wait_for_selector("#numeroGuia124_input", state="visible", timeout=5000)
    preencher_sem_scroll(page, "#numeroGuia124_input", dados["numero"])
    time.sleep(0.4)

    # 3. S/N? (checkbox)
    checkbox = page.locator("#desconhecidoCadastro1_input").first
    if dados["s_n"]:
        if not checkbox.is_checked():
            checkbox.click(force=True)

    # 4. Ano
    preencher_sem_scroll(page, "input[id*='anoGuia']", dados["ano"])
    time.sleep(0.2)

    # 5. Referência
    preencher_sem_scroll(page, "#referencia", dados["referencia"])
    time.sleep(0.2)

    # 6. Unidade Requisitante
    page.wait_for_selector("#unidadePolicial_label", state="visible", timeout=5000)
    selecionar_dropdown_por_valor(page, "#unidadePolicial_label", dados["unidade_requisitante"])
    time.sleep(0.4)

    # 7. Unidade Destino
    if dados["unidade_destino"]:
        selecionar_dropdown_por_valor(page, "#unidadeDestino_label", dados["unidade_destino"])
        time.sleep(0.4)

    # 8. Cargo da Autoridade Requisitante (obrigatório)
    selecionar_dropdown_por_valor(page, "#cargoAutoridadeRequisitante_label", dados["cargo_autoridade"])
    time.sleep(0.4)

    # 9. Autoridade Requisitante (texto)
    preencher_sem_scroll(page, "#autoridaderequisitante", dados["autoridade"])
    time.sleep(0.2)

    # 10. Data do Documento (sem focus para não abrir o calendário)
    preencher_sem_scroll(page, "input[id*='dataDocumento']", dados["data_documento"], evita_focus=True)
    time.sleep(0.2)

    # 11. Data e Hora de Recebimento (sem focus para não abrir o calendário)
    preencher_sem_scroll(page, "input[id*='dataHoraRecebimento']", dados["data_hora_recebimento"], evita_focus=True)

    # Fecha qualquer date picker que tenha aberto e remove overlays
    limpar_bloqueios(page)
    time.sleep(0.3)

    log("   [OK] Formulario preenchido.")
    return dados


def buscar_ocorrencia_cadastrada(page, numero_req, ano):
    """
    Após retorno à página moduloEntradaCartorio, busca a linha com o número da requisição
    cadastrada e retorna o número da ocorrência gerada (coluna j_idt418, primeiro valor).
    """
    try:
        page.wait_for_selector("#tabelaSolicitacoes_data", state="visible", timeout=15000)
        time.sleep(2)  # aguarda dados da tabela
        limpar_bloqueios(page)

        rows = page.locator("#tabelaSolicitacoes_data tr[role='row']")
        n = rows.count()
        for i in range(n):
            row = rows.nth(i)
            texto = row.inner_text()
            # Busca pela linha que contém o número da requisição e ano
            if str(numero_req) in texto and str(ano) in texto:
                try:
                    # Ocorrência está no label j_idt418 - formato "60051/2026 57469/2026"
                    label_oc = row.locator("label[id*='j_idt418']").first
                    valor = label_oc.inner_text(timeout=3000).strip()
                    # Primeiro número é a ocorrência
                    partes = valor.split()
                    if partes:
                        return partes[0].strip()
                except Exception:
                    pass
        return None
    except Exception as e:
        log(f"   [AVISO] Erro ao buscar ocorrência: {e}")
        return None


def abrir_modal_evidencia_e_verificar(page, ocorrencia):
    """
    Clica em 'Adicionar evidencia' na primeira linha (ocorrência recém-criada),
    abre o modal e verifica se o número da ocorrência no modal corresponde.
    Retorna True se OK.
    """
    try:
        # Número da ocorrência (ex: "60051/2026" -> "60051")
        num_oc = ocorrencia.split("/")[0] if "/" in ocorrencia else ocorrencia

        # Clicar Adicionar evidencia na primeira linha
        link_ev = page.locator("#tabelaSolicitacoes_data tr[role='row']").first.locator('a[aria-label="Adicionar evidencia"]')
        link_ev.wait_for(state="visible", timeout=5000)
        link_ev.click()
        time.sleep(2)
        limpar_bloqueios(page)

        # Aguardar modal PESSOAS / VESTÍGIOS / EVIDÊNCIAS
        page.wait_for_selector("form#formVestigio", state="visible", timeout=10000)
        page.wait_for_selector(".ui-dialog-title:has-text('PESSOAS')", state="visible", timeout=5000)

        # Verificar ocorrência no modal (primeiro fieldset = Dados da ocorrência)
        oc_modal = page.locator("form#formVestigio fieldset .ui-fieldset-content .ui-g-6").first.locator("label").nth(1)
        oc_modal_txt = oc_modal.inner_text(timeout=3000).strip()

        if num_oc in oc_modal_txt or oc_modal_txt == num_oc:
            log(f"   [OK] Modal de evidencia aberto - Ocorrencia {oc_modal_txt} confere.")
            return True
        log(f"   [AVISO] Modal aberto mas ocorrencia no modal ({oc_modal_txt}) difere da esperada ({num_oc}).")
        return True  # Modal abriu, continuamos
    except Exception as e:
        log(f"   [AVISO] Erro ao abrir modal evidencia: {e}")
        return False


def extrair_dropdown_opcoes(page, seletor_label, arquivo_saida):
    """Extrai opções de um dropdown PrimeFaces e salva em arquivo."""
    try:
        page.wait_for_selector(seletor_label, state="visible", timeout=5000)
        page.click(seletor_label)
        time.sleep(0.8)
        itens = page.locator(".ui-selectonemenu-panel li.ui-selectonemenu-item").all_inner_texts()
        if itens:
            caminho = os.path.join(_dir_script(), arquivo_saida)
            with open(caminho, "w", encoding="utf-8") as f:
                for txt in itens:
                    f.write(txt.strip() + "\n")
            log(f"   {len(itens)} opcoes extraidas e salvas em {arquivo_saida}")
        page.keyboard.press("Escape")
        time.sleep(0.3)
        return itens if itens else []
    except Exception as e:
        log(f"   [AVISO] Erro ao extrair dropdown: {e}")
        return []


def _preencher_vestigio_droga_abuso(page):
    """Preenche formulário de vestígio tipo Drogas de Abuso (Tipo dropdown já visível)."""
    label_el = page.locator("#tipoVestigio_label")
    if label_el.count() > 0:
        try:
            selecionar_dropdown_por_valor(page, "#tipoVestigio_label", "Drogas de Abuso")
        except Exception:
            page.locator("#tipoVestigio_input").select_option(value="MATERIAL_QUIMICO_BIOLOGICO")
    else:
        page.locator("#tipoVestigio_input").select_option(value="MATERIAL_QUIMICO_BIOLOGICO")
    time.sleep(1)
    limpar_bloqueios(page)
    log("   [OK] Tipo 'Drogas de Abuso' selecionado.")

    # Aguardar formulário de vestígio (Tipo material, Involucro, etc.)
    page.wait_for_selector("#apresentacaoMaterialVestigio_label", state="visible", timeout=8000)
    time.sleep(0.5)

    # Valores do cadastro (para verificação)
    involucro = random.choice(INVOLUCROS_VALIDOS)
    valores_evidencia = {
        "Tipo material": "Vegetal",
        "Quantidade": "5",
        "Involucro": involucro,
        "Substancia": "Maconha",
        "Qtd Apresentada Material": "200",
        "Unidade": "mg",
        "Qtd Apresentada Massa liquida": "300",
        "Unidade Massa liquida": "mg",
        "Quantidade Utilizada": "150",
        "Unidade Utilizado": "mg",
        "Servico": "Quimica",
        "Exame": "THC",
    }
    log("   Valores para cadastro da evidencia (verificar na tela):")
    for k, v in valores_evidencia.items():
        log(f"      {k}: {v}")

    # Tipo de material: Vegetal
    for seletor in ["#tipoMaterialVestigio_label", "#tipoMaterial_label", "label[id*='tipoMaterial']"]:
        if page.locator(seletor).count() > 0:
            selecionar_dropdown_por_valor(page, seletor, "Vegetal")
            break
    time.sleep(0.3)

    # Quantidade: 5 (quantidadeDeEmbalagens)
    preencher_sem_scroll(page, "#quantidadeDeEmbalagens_input", "5")
    time.sleep(0.2)

    # Invólucro: valor aleatório
    if involucro:
        selecionar_dropdown_por_valor(page, "#apresentacaoMaterialVestigio_label", involucro)
    time.sleep(0.3)

    # Substância: Maconha
    for seletor in ["#substanciaVestigio_label", "#substancia_label", "label[id*='substancia']"]:
        if page.locator(seletor).count() > 0:
            selecionar_dropdown_por_valor(page, seletor, "Maconha")
            break
    time.sleep(0.3)

    # Quantidade Apresentada de Material: 200
    preencher_sem_scroll(page, "#quantidadeApresentadaMaterial_input", "200")
    time.sleep(0.2)

    # Unidade: mg
    selecionar_dropdown_por_valor(page, "#unidadeMedidaApresentada_label", "mg")
    time.sleep(0.3)

    # Quantidade Apresentada Material Massa liquida: 300
    preencher_sem_scroll(page, "#quantidadeApresentadaMaterialMassaliquida_input", "300")
    time.sleep(0.2)

    # Unidade Medida Massa liquida: mg
    selecionar_dropdown_por_valor(page, "#unidadeMedidaApresentadaMassaliquida_label", "mg")
    time.sleep(0.3)

    # Quantidade Utilizada: 150
    preencher_sem_scroll(page, "#quantidadeUtilizado_input", "150")
    time.sleep(0.2)

    # Unidade de medida Utilizado: mg
    selecionar_dropdown_por_valor(page, "#unidadeMedidaUtilizado_label", "mg")
    time.sleep(0.3)

    # Serviço: Química
    try:
        sel_serv = page.locator("label.ui-selectonemenu-label:has-text('Selecione o serviço')").first
        sel_serv.click()
        time.sleep(0.5)
        page.keyboard.type("Química", delay=50)
        time.sleep(0.3)
        page.keyboard.press("Enter")
        limpar_bloqueios(page)
    except Exception:
        pass
    time.sleep(0.5)

    # THC (tipo de exame - após Serviço)
    try:
        thc_sel = page.locator("label[id*='2927']").first
        if thc_sel.count() > 0:
            selecionar_dropdown_por_valor(page, "label[id*='2927']", "THC")
        else:
            exam_labels = page.locator("label.ui-selectonemenu-label")
            for i in range(exam_labels.count()):
                lbl = exam_labels.nth(i)
                txt = lbl.inner_text()
                if "Selecione" in txt and "serviço" not in txt.lower():
                    lbl.click()
                    time.sleep(0.5)
                    page.keyboard.type("THC", delay=50)
                    time.sleep(0.3)
                    page.keyboard.press("Enter")
                    limpar_bloqueios(page)
                    break
    except Exception:
        pass
    time.sleep(0.3)

    # Clicar Adicionar
    page.click('button:has-text("Adicionar")')
    limpar_bloqueios(page)
    time.sleep(2)
    log("   [OK] Dados do vestigio preenchidos e Adicionar acionado.")

    # Clicar Salvar
    page.click("#btnSalvarEvidenciaVestigio")
    limpar_bloqueios(page)
    time.sleep(2)
    log("   [OK] Evidencia salva.")


def incluir_vestigio_tipo_droga(page):
    """Clica em Incluir, seleciona Drogas de Abuso e preenche o formulário."""
    try:
        page.click("#btnIncluirEnvolvidoVestigioSolicitacao")
        limpar_bloqueios(page)
        time.sleep(2)
        page.wait_for_selector("#tipoVestigio_input", state="attached", timeout=8000)
        time.sleep(0.5)
        _preencher_vestigio_droga_abuso(page)
    except Exception as e:
        log(f"   [AVISO] Erro ao incluir vestigio: {e}")


def extrair_vestigio_tipo_droga(page):
    """Clica em Extrair vestígio (primeira linha da tabela), seleciona Drogas de Abuso e preenche."""
    try:
        page.click("a[id*='cmdExtrairEvidenciaVestigio']")
        limpar_bloqueios(page)
        time.sleep(2)
        page.wait_for_selector("#tipoVestigio_input", state="attached", timeout=8000)
        time.sleep(0.5)
        _preencher_vestigio_droga_abuso(page)
    except Exception as e:
        log(f"   [AVISO] Erro ao extrair vestigio: {e}")


def incluir_vestigio_tipo_pessoa_envolvida(page):
    """
    Clica em Incluir, seleciona 'Pessoa Envolvida' e preenche os campos do envolvido.
    Modal de evidência deve permanecer aberto após a evidência anterior.
    """
    try:
        # Clicar Incluir (no mesmo modal)
        page.click("#btnIncluirEnvolvidoVestigioSolicitacao")
        limpar_bloqueios(page)
        time.sleep(2)

        page.wait_for_selector("#tipoVestigio_input", state="attached", timeout=8000)
        time.sleep(0.5)

        # Selecionar Pessoa Envolvida
        label_el = page.locator("#tipoVestigio_label")
        if label_el.count() > 0:
            try:
                selecionar_dropdown_por_valor(page, "#tipoVestigio_label", "Pessoa Envolvida")
            except Exception:
                page.locator("#tipoVestigio_input").select_option(value="PESSOA_ENVOLVIDA")
        else:
            page.locator("#tipoVestigio_input").select_option(value="PESSOA_ENVOLVIDA")
        time.sleep(1)
        limpar_bloqueios(page)
        log("   [OK] Tipo 'Pessoa Envolvida' selecionado.")

        # Valores para o envolvido (verificação)
        nome_env = gerar_nome_aleatorio()
        valores_envolvido = {
            "Tipo": "Pessoa Envolvida",
            "Tipo envolvimento": "Implicado",
            "Tipo Implicado": "Investigado",
            "Nome": nome_env,
        }
        log("   Valores para cadastro do envolvido (verificar na tela):")
        for k, v in valores_envolvido.items():
            log(f"      {k}: {v}")

        # Clicar Dados Gerais (accordion)
        page.click('div.ui-accordion-header:has-text("Dados Gerais")')
        time.sleep(0.8)

        # Tipo de envolvimento: Implicado (id com : precisa escape em CSS)
        try:
            selecionar_dropdown_por_valor(page, "[id='envolvidoAccordion:tipoEnvolvidoCadastro_label']", "Implicado")
        except Exception:
            page.locator("[id='envolvidoAccordion:tipoEnvolvidoCadastro_label']").click()
            time.sleep(0.5)
            page.keyboard.type("Implicado", delay=50)
            page.keyboard.press("Enter")
        time.sleep(0.5)

        # Tipo Implicado: Investigado
        try:
            selecionar_dropdown_por_valor(page, "[id='envolvidoAccordion:tipoImplicado_label']", "Investigado")
        except Exception:
            page.locator("[id='envolvidoAccordion:tipoImplicado_label']").click()
            time.sleep(0.5)
            page.keyboard.type("Investigado", delay=50)
            page.keyboard.press("Enter")
        time.sleep(0.3)

        # Nome (selectors comuns para formulário de envolvido)
        for seletor in ["input[id*='nome']", "input[id*='Nome']", "input[data-p-label*='Nome']"]:
            if page.locator(seletor).count() > 0:
                preencher_sem_scroll(page, seletor, nome_env)
                break
        time.sleep(0.3)

        # Clicar Salvar
        page.click("#btnSalvarEvidenciaVestigio")
        limpar_bloqueios(page)
        time.sleep(2)
        log("   [OK] Envolvido salvo.")
    except Exception as e:
        log(f"   [AVISO] Erro ao incluir envolvido: {e}")


def fechar_modal_vestigio(page):
    """Fecha o modal de vestígios/evidências clicando no botão X (btnFecharModalVestigio)."""
    try:
        page.click("#btnFecharModalVestigio")
        limpar_bloqueios(page)
        time.sleep(2)
        log("   [OK] Modal de vestigio fechado.")
    except Exception as e:
        log(f"   [AVISO] Erro ao fechar modal: {e}")


def solicitar_exames_ocorrencia(page, ocorrencia):
    """
    Busca a ocorrência pelo número, garante que só ela aparece na tela, e clica em
    'Solicitar exames' na primeira linha (tabelaSolicitacoes:0:cmdSolicitarExamesOcorrencia).
    """
    try:
        num_oc = ocorrencia.split("/")[0] if "/" in ocorrencia else ocorrencia

        # Buscar ocorrência para garantir que tabelaSolicitacoes:0 seja a nossa
        campo = page.locator("input[id$='campoNumeroRapido_input']")
        campo.wait_for(state="visible", timeout=8000)
        campo.fill("")
        campo.fill(num_oc)
        time.sleep(0.5)
        page.keyboard.press("Enter")
        log(f"   [OK] Busca por ocorrencia {num_oc} acionada.")

        page.wait_for_selector("#tabelaSolicitacoes_data", state="visible", timeout=10000)
        time.sleep(2)

        # Clicar em Solicitar exames (primeira linha = nossa ocorrência)
        link = page.locator("a[id*='cmdSolicitarExamesOcorrencia']").first
        link.wait_for(state="visible", timeout=5000)
        link.click()
        limpar_bloqueios(page)
        time.sleep(2)
        log("   [OK] Modal/tela de exames suplementares aberto.")
    except Exception as e:
        log(f"   [AVISO] Erro ao solicitar exames: {e}")


def incluir_solicitacao_exame_suplementar(page):
    """
    Na tela de exames suplementares: clica Solicitar Exame, seleciona Unidade e Exame,
    marca o vestígio da droga extraída (Vegetal - Id: ...) e clica Incluir.
    """
    try:
        # 1. Clicar em Solicitar Exame
        page.wait_for_selector("#btnIncluirSolicitacao", state="visible", timeout=8000)
        page.click("#btnIncluirSolicitacao")
        limpar_bloqueios(page)
        time.sleep(2)
        log("   [OK] Botao Solicitar Exame acionado.")

        # 2. Unidade: QUI ICRIM IMP - Seção de Química Forense
        page.wait_for_selector("#tipoExameSetorSolicitacao_label", state="visible", timeout=8000)
        selecionar_dropdown_por_valor(page, "#tipoExameSetorSolicitacao_label", "QUI ICRIM IMP - Seção de Química Forense")
        time.sleep(0.5)
        log("   [OK] Unidade selecionada.")

        # 3. Exame: Químico para identificação de THC - Definitivo
        page.wait_for_selector("#tipoExameSolicitacao_label", state="visible", timeout=8000)
        selecionar_dropdown_por_valor(page, "#tipoExameSolicitacao_label", "Químico para Identificação de THC - Definitivo")
        time.sleep(1)  # Aguarda árvore de vestígios carregar
        log("   [OK] Exame selecionado.")

        # 4. Marcar o vestígio da droga extraída (Vegetal - Id: ...)
        # Busca nó da árvore que contenha "Vegetal - Id:" (confere que é a droga de abuso extraída)
        no_vestigio = page.locator("li.ui-treenode-leaf:has(.ui-treenode-label:has-text('Vegetal - Id:'))").first
        no_vestigio.wait_for(state="visible", timeout=8000)
        # Clicar no checkbox para marcar
        no_vestigio.locator(".ui-chkbox-box").first.click()
        limpar_bloqueios(page)
        time.sleep(0.5)
        log("   [OK] Vestigio droga de abuso marcado (Vegetal - Id:).")

        # 5. Clicar em Incluir
        page.click("#btnSalvarSolicitacao")
        limpar_bloqueios(page)
        time.sleep(2)
        log("   [OK] Solicitacao de exame incluida.")
    except Exception as e:
        log(f"   [AVISO] Erro ao incluir solicitacao exame: {e}")


def executar_cadastro_teste():
    """Fluxo de teste: login + navegação + etapas de cadastro com valores fixos."""
    inicio = time.time()
    log("--- [INICIO] CADASTRO GALILEU (TESTE) ---")
    log(f"URL: {URL_TREINAMENTO}")
    log("headless=False - navegador visível")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors"]
        )
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()
        page.set_viewport_size({"width": 1920, "height": 1080})

        try:
            # 1. Login (página inicial = ambiente de treinamento)
            log("1. [LOGIN] Acessando pagina de login (treinamento)...")
            page.goto(URL_TREINAMENTO)
            time.sleep(1)

            # Tenta o seletor padrão do Galileu
            page.fill("input[id='nomeUsuario']", USUARIO)
            page.fill("input[id='senhaUsuario']", SENHA)
            with page.expect_navigation(timeout=60000):
                page.click("[id='BtEntrar']")
            log("   Login realizado.")

            time.sleep(2)

            # 2. Navegação para o módulo de entrada cartório
            log("2. [NAV] Acessando modulo de entrada cartorio...")
            page.goto(URL_MODULO_CARTORIO)
            page.wait_for_selector("#tabelaSolicitacoes_data, #novaSolicitacao", state="visible", timeout=60000)
            log("   Modulo carregado.")

            # 3. Clicar em Nova Ocorrência
            log("3. [NOVA OCORRENCIA] Clicando no botao Nova Ocorrencia...")
            page.click("#novaSolicitacao")
            limpar_bloqueios(page)
            time.sleep(2)

            # 4. Setor de destino: digitar + Enter
            log("4. [SETOR DESTINO] Digitando Seção de Química Forense...")
            page.wait_for_selector("#setorPericia_label", state="visible", timeout=10000)
            page.click("#setorPericia_label")
            time.sleep(0.5)
            page.keyboard.type(DADOS_TESTE["setor_destino"], delay=50)
            time.sleep(0.3)
            page.keyboard.press("Enter")
            limpar_bloqueios(page)
            log("   Setor selecionado.")

            # 5. Tipo de exame: digitar + Enter
            log("5. [TIPO EXAME] Digitando Químico para Identificação de THC - Definitivo...")
            page.wait_for_selector("#tipoExameSetor_label", state="visible", timeout=10000)
            page.click("#tipoExameSetor_label")
            time.sleep(0.5)
            page.keyboard.type(DADOS_TESTE["tipo_exame"], delay=50)
            time.sleep(0.3)
            page.keyboard.press("Enter")
            limpar_bloqueios(page)
            log("   Tipo de exame selecionado.")

            # 5b. Clicar em Confirmar
            log("5b. [CONFIRMAR] Clicando no botao Confirmar...")
            page.wait_for_selector("#linkConfirmar", state="visible", timeout=10000)
            page.click("#linkConfirmar")
            limpar_bloqueios(page)
            time.sleep(2)
            log("   Confirmar acionado.")

            # 6. Extração (inativo - base já extraída)
            # log("6. [EXTRAIR] Unidade Destino...")
            # extrair_unidades_destino(page)
            # extrair_dropdown(page, "#tipoDocumentoPolicial_label", ARQUIVO_TIPO_DOCUMENTO)
            # extrair_dropdown(page, "#cargoAutoridadeRequisitante_label", ARQUIVO_CARGO_AUTORIDADE)
            # extrair_dropdown(page, "#unidadePolicial_label", ARQUIVO_UNIDADE_POLICIAL)
            time.sleep(1)

            # 7. Preencher formulário com dados aleatórios
            log("7. [CADASTRO] Preenchendo formulario com dados aleatorios...")
            dados = preencher_formulario_cadastro(page)

            # 8. Clicar em Validar (garante que date picker e overlays estejam fechados)
            log("8. [VALIDAR] Clicando em Validar...")
            limpar_bloqueios(page)
            time.sleep(0.5)
            page.click('button:has-text("Validar")', force=True)
            limpar_bloqueios(page)
            time.sleep(2)
            log("   Validar acionado.")

            # 9. Clicar em Confirmar (surgido após validação)
            log("9. [CONFIRMAR] Clicando em Confirmar...")
            page.wait_for_selector('button[title="confirmar1"]', state="visible", timeout=10000)
            page.click('button[title="confirmar1"]')
            limpar_bloqueios(page)
            time.sleep(3)
            log("   Confirmar acionado.")

            # 10. Aguardar retorno à página (modal fecha, lista atualiza) e buscar ocorrência gerada
            log("10. [VERIFICAR] Aguardando retorno à página e buscando ocorrência...")
            limpar_bloqueios(page)
            ocorrencia = buscar_ocorrencia_cadastrada(page, dados["numero"], dados["ano"])
            if ocorrencia:
                log("")
                log("   ========================================")
                log(f"   OCORRENCIA GERADA: {ocorrencia}")
                log("   ========================================")
                log("")

                # 11. Abrir modal de evidência (primeira linha = ocorrência criada) e verificar
                log("11. [EVIDENCIA] Abrindo modal Adicionar evidencia...")
                abrir_modal_evidencia_e_verificar(page, ocorrencia)

                # 12. Evidência: Drogas de Abuso
                log("12. [EVIDENCIA] Incluindo evidencia tipo Drogas de Abuso...")
                incluir_vestigio_tipo_droga(page)

                # 13. Extrair vestígio (da evidência na tabela)
                log("13. [EXTRAIR VESTIGIO] Clicando Extrair vestigio e preenchendo Drogas de Abuso...")
                extrair_vestigio_tipo_droga(page)

                # 14. Envolvido: Pessoa Envolvida
                log("14. [ENVOLVIDO] Incluindo envolvido tipo Pessoa Envolvida...")
                incluir_vestigio_tipo_pessoa_envolvida(page)

                # 15. Fechar modal de vestígios
                log("15. [FECHAR] Fechando modal de vestigios...")
                fechar_modal_vestigio(page)

                # 16. Exames suplementares (buscar ocorrência e clicar Solicitar exames)
                log("16. [EXAMES SUPLEMENTARES] Buscando ocorrencia e abrindo solicitar exames...")
                solicitar_exames_ocorrencia(page, ocorrencia)

                # 17. Incluir solicitação de exame (Unidade, Exame, vestígio, Incluir)
                log("17. [SOLICITAR EXAME] Clicando Solicitar Exame e preenchendo...")
                incluir_solicitacao_exame_suplementar(page)
            else:
                log("   [AVISO] Registro não encontrado na tabela (verifique numero/ano).")

            log("   Pausa para inspeção - feche o navegador ou aguarde 30s.")
            time.sleep(30)

        except Exception as e:
            log(f"[ERRO] {e}")
            import traceback
            traceback.print_exc()
            time.sleep(15)
        finally:
            try:
                browser.close()
            except Exception:
                pass
        duracao = time.time() - inicio
        mins = int(duracao // 60)
        segs = int(duracao % 60)
        log(f"--- [FIM] Tempo total: {mins}min {segs}s ---")


if __name__ == "__main__":
    executar_cadastro_teste()
