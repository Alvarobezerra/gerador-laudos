# Manual de Implantação - Automação do Galileu (Cadastro de Ocorrências)

Este manual descreve o procedimento para implantar e executar a automação de cadastro do sistema Galileu em outro computador.

---

## 1. Arquivos Necessários

Para que a automação funcione no novo computador, copie os seguintes arquivos (que já estão empacotados no arquivo `automacao_galileu_cadastro.zip`):

* **Formulário Streamlit (Interface):**
  * `formulario_cadastro_galileu.py` (formulário para digitação e salvamento dos dados)
* **Robô de Automação (Playwright):**
  * `robo_galileu_cadastro_teste.py` (script que faz o login, preenche os dados e cadastra)
* **Arquivos de Configuração e Opções:**
  * `opcoes_dropdown_mapeados.json` (mapeamento de campos extraídos da página de registro)
  * `cargo_autoridade_galileu.txt` (opções válidas de cargo de autoridade no Galileu)
  * `involucro_galileu.txt` (opções de invólucros)
  * `tipo_documento_galileu.txt` (opções de tipo de documento)
  * `unidade_policial_galileu.txt` (lista de unidades policiais requisitantes)
  * `unidades_destino_galileu.txt` (lista de seções de destino/perícia)
* **Instalação:**
  * `requirements_cadastro.txt` (bibliotecas necessárias)

---

## 2. Pré-requisitos de Instalação

No novo computador, execute os seguintes passos para preparar o ambiente:

### Passo 1: Instalar o Python
* Baixe e instale o **Python 3.10 ou superior** (marcando a opção *"Add Python to PATH"* durante a instalação).

### Passo 2: Criar e ativar o Ambiente Virtual (.venv)
Abra o Prompt de Comando (CMD) ou PowerShell na pasta do projeto e execute:
```bash
# Criar o ambiente virtual
python -m venv .venv

# Ativar o ambiente virtual
# No Prompt de Comando (CMD):
.venv\Scripts\activate

# No PowerShell:
.venv\Scripts\Activate.ps1
```

### Passo 3: Instalar as Dependências
Com o ambiente virtual ativado, instale as bibliotecas necessárias:
```bash
pip install -r requirements_cadastro.txt
```

### Passo 4: Instalar e configurar os navegadores do Playwright
O Playwright precisa baixar seus próprios binários do navegador Chromium para poder rodar. Execute o comando:
```bash
playwright install chromium
```

---

## 3. Como Executar os Módulos

### 1. Executando o Formulário (Streamlit)
O formulário serve para que o perito/usuário digite os dados da ocorrência e salve os dados localmente.
```bash
streamlit run formulario_cadastro_galileu.py
```
Isso abrirá uma página web local no seu navegador padrão (`http://localhost:8501`) contendo a interface de preenchimento.

### 2. Executando o Robô de Cadastro (Playwright)
Com o robô configurado com os dados reais ou em ambiente de testes:
```bash
python robo_galileu_cadastro_teste.py
```
*(Nota: Certifique-se de ajustar as variáveis `USUARIO`, `SENHA` e a `URL` dentro de `robo_galileu_cadastro_teste.py` antes de rodar).*

---

## 4. Estrutura do Mapeamento de Ocorrência
O arquivo `opcoes_dropdown_mapeados.json` contém as opções válidas que o robô utiliza para preencher os seletores da página `moduloRegistroOcorrencia.faces`.

Desenvolvido para apoio às equipes do Instituto de Criminalística (ICRIM).
