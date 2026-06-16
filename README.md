# 🎓 Automação de Rubricas — Google Classroom → Excel

Ferramenta de automação que **extrai rubricas do Google Classroom** e **preenche automaticamente planilhas Excel**, eliminando o trabalho manual de copiar avaliações critério a critério.

---

## 📋 Visão Geral

O fluxo de trabalho é dividido em duas etapas:

1. **Extração** — O sistema abre o Chrome com um perfil de usuário já logado no Google Classroom, navega pela tabela de alunos e captura os critérios e graus avaliados (G1–G5) de cada rubrica, salvando tudo em `rubricas.json`.

2. **Preenchimento** — O professor informa a **autonomia** (Autônomo / Parcialmente autônomo / Apoiado / Não satisfatório) de cada aluno. O sistema então combina a **autonomia** (faixa de pontuação) com o **grau** (G1–G5, subnível) para calcular a pontuação interna (1–20) e marca a célula correspondente na planilha Excel aberta.

### Mapeamento de pontuação

| Autonomia | Faixa |
|---|---|
| Não satisfatório | 1 – 5 |
| Apoiado | 6 – 10 |
| Parcialmente autônomo | 11 – 15 |
| Autônomo | 16 – 20 |

O grau (G1–G5) define o subnível dentro da faixa: G1 → posição 1, G2 → +1, …, G5 → +4.

---

## 🏗️ Arquitetura

```
py copia rubricas/
├── app.py                  # Servidor Flask (API REST)
├── extrair_rubrica.py      # Selenium: extrai rubricas do Classroom
├── preencher_planilha.py   # Lógica de pontuação + automação Excel (win32com)
├── update_planilha.py      # Orquestrador do preenchimento com autonomia do professor
├── requirements.txt        # Dependências Python
├── rubricas.json           # ⚠️ Gerado em tempo de execução (não versionado)
└── frontend/               # Interface React + Vite + TailwindCSS
    ├── src/
    │   ├── App.tsx          # Componente principal (UI completa)
    │   └── main.tsx
    ├── package.json
    └── vite.config.ts
```

---

## ⚙️ Pré-requisitos

| Requisito | Versão mínima | Observação |
|---|---|---|
| **Windows** | 10 / 11 | Obrigatório — usa `win32com` para Excel |
| **Python** | 3.10+ | |
| **Node.js** | 18+ | Para o frontend |
| **Google Chrome** | Qualquer versão recente | Deve estar instalado |
| **Microsoft Excel** | Qualquer versão Desktop | Deve estar aberto durante o preenchimento |

---

## 🚀 Instalação e Configuração

### 1. Clone o repositório

```bash
git clone <url-do-repositorio>
cd "py copia rubricas"
```

### 2. Backend Python

```bash
# Crie e ative um ambiente virtual (recomendado)
python -m venv .venv
.venv\Scripts\activate

# Instale as dependências
pip install -r requirements.txt
```

### 3. Configurar o perfil do Chrome para o Selenium

O Selenium reutiliza um perfil Chrome já logado no Google para evitar login manual a cada execução.

```bash
# Crie a pasta do perfil (se não existir)
mkdir C:\selenium\chrome-profile
```

Na primeira execução, abra o Chrome **manualmente** com este perfil e faça login na conta Google que tem acesso ao Classroom:

```bash
chrome.exe --user-data-dir="C:\selenium\chrome-profile"
```

> **Importante:** Feche o Chrome completamente antes de rodar a automação, pois o Selenium precisa controlar o processo.

### 4. Frontend

```bash
cd frontend
npm install
```

---

## ▶️ Executando o projeto

### Terminal 1 — Backend

```bash
# Na raiz do projeto
python app.py
```

O servidor Flask inicia em `http://localhost:5000`.

### Terminal 2 — Frontend

```bash
cd frontend
npm run dev
```

A interface abre em `http://localhost:5173`.

---

## 🖥️ Como usar

### Passo 1 — Extrair Rubricas

1. Abra a interface em `http://localhost:5173`
2. Clique em **"1. Extrair Rubricas"** — o Chrome abre automaticamente no Classroom
3. Navegue até a atividade com a tabela de alunos
4. Clique em **"Confirmar Página do Classroom"** na interface
5. O sistema percorre cada aluno automaticamente e salva `rubricas.json`

### Passo 2 — Preencher Planilha

1. Abra a planilha Excel e deixe-a **aberta e visível**
2. Clique em **"2. Preencher Planilha"**
3. Um modal exibe os alunos com graus avaliados — selecione a **autonomia** de cada um
4. Clique em **"Confirmar e Preencher"**
5. O sistema preenche a planilha e exibe o log no console

---

## 📁 Estrutura da Planilha Excel Esperada

A planilha deve seguir este layout:

| Configuração | Valor |
|---|---|
| Coluna dos critérios/capacidades | **G (coluna 7)** |
| Início da faixa de avaliação | **H (coluna 8)** — representa pontuação 1 |
| Fim da faixa de avaliação | **AB (coluna 27)** — representa pontuação 20 |
| Linha inicial de busca | **21** |
| Nome do aluno | **M10** (linha 10, coluna 13) |

Cada critério recebe um `✓` na coluna correspondente à pontuação. Cada capacidade recebe um `X` com base na mediana dos critérios do seu bloco.

---

## 📡 API REST (Flask)

| Método | Rota | Descrição |
|---|---|---|
| `GET` | `/api/status` | Verifica se há processo em execução |
| `GET` | `/api/logs` | Retorna logs acumulados desde a última leitura |
| `POST` | `/api/extrair` | Inicia a extração do Classroom |
| `POST` | `/api/extrair/confirmar` | Confirma que a página do Classroom está pronta |
| `GET` | `/api/preencher/check` | Lista alunos com graus válidos (precisam de autonomia) |
| `POST` | `/api/preencher/execute` | Executa o preenchimento com as autonomias informadas |
| `POST` | `/api/stop` | Solicita parada do processo em curso |

### Exemplo — `/api/preencher/execute`

```json
POST /api/preencher/execute
{
  "autonomias": {
    "João da Silva": "Autônomo",
    "Maria Souza": "Apoiado"
  }
}
```

---

## 🔧 Variáveis de configuração

Todas as constantes de layout da planilha estão centralizadas no topo de [`preencher_planilha.py`](preencher_planilha.py):

```python
COLUNA_CRITERIO      = 7   # Coluna G
COL_INICIO_AVALIACAO = 8   # Coluna H (pontuação 1)
COL_FIM_AVALIACAO    = 27  # Coluna AB (pontuação 20)
LINHA_INICIAL_BUSCA  = 21  # Primeira linha pesquisável
LINHA_NOME_ALUNO     = 10  # Linha do nome na aba
COLUNA_NOME_ALUNO    = 13  # Coluna M
```

Se o layout da planilha mudar, ajuste apenas essas constantes.

---

## ⚠️ Arquivos sensíveis / não versionados

| Arquivo | Motivo |
|---|---|
| `rubricas.json` | Contém dados pessoais dos alunos — gerado em tempo de execução |
| `*.json` (raiz) | Todos os JSONs da raiz são dados de execução |
| `frontend/node_modules/` | Dependências — reinstale com `npm install` |
| `.venv/` | Ambiente virtual Python |
| `C:\selenium\chrome-profile` | Perfil Chrome com sessão logada — configurar localmente |

---

## 🐛 Problemas comuns

**`win32com` não encontrado**
```bash
pip install pywin32
python -m pywin32_postinstall -install
```

**Chrome já aberto ao iniciar extração**
> Feche todas as janelas do Chrome antes de rodar. O Selenium precisa controlar o processo do zero.

**Aluno não encontrado na planilha**
> Verifique se o nome na célula M10 da aba do aluno corresponde exatamente (ou parcialmente) ao nome retornado pelo Classroom.

**Planilha não encontrada**
> O Excel deve estar aberto com a planilha correta antes de clicar em "Preencher Planilha".
