# 🎓 Automação de Rubricas — Google Classroom → Excel (Monorepo)

Ferramenta de automação que **extrai rubricas do Google Classroom** (inclusive em segundo plano) e **preenche automaticamente planilhas Excel**, eliminando o trabalho manual de copiar avaliações critério a critério.

---

## 📋 Visão Geral

O fluxo de trabalho é dividido em duas etapas:

1. **Extração** — O sistema abre o Chrome de forma invisível (modo oculto / headless) ou visível. Ele navega pela tabela de alunos e captura os critérios e graus avaliados (G1–G5) de cada rubrica, salvando tudo em `rubricas.json` na raiz do projeto.
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

## 🏗️ Arquitetura do Monorepo

O projeto está organizado em uma arquitetura monorepo limpa:

```
Extrair-Rubricas-Preencher-Excel/ (raiz)
├── package.json            # Scripts de automação unificados do monorepo
├── rubricas.json           # ⚠️ Gerado em tempo de execução (dados de rubricas extraídos)
├── backend/                # Servidor Backend Python (Flask + Selenium + win32com)
│   ├── app.py              # Ponto de entrada da API REST
│   ├── extrair_rubrica.py  # Selenium: extrai rubricas do Classroom (headless/visível)
│   ├── preencher_planilha.py # Lógica de pontuação + automação Excel (win32com)
│   ├── update_planilha.py  # Orquestrador do preenchimento com autonomia do professor
│   └── requirements.txt    # Dependências Python do backend
└── frontend/               # Interface de usuário (Vite + React + TypeScript)
    ├── src/
    │   ├── App.tsx         # Componente principal (UI unificada)
    │   └── main.tsx
    ├── package.json        # Dependências e scripts do frontend
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

Graças aos scripts unificados do Monorepo, você pode configurar tudo com poucos comandos a partir da raiz:

### 1. Instalar dependências (Monorepo Setup)
Execute na raiz do projeto:
```bash
# Instala as dependências do React (frontend) e do Python (backend) automaticamente
npm run setup
```

### 2. Configurar o perfil do Chrome para o Selenium
O Selenium reutiliza um perfil do Chrome para evitar login manual a cada execução.
```bash
# Crie a pasta do perfil no seu sistema (se não existir)
mkdir C:\selenium\chrome-profile
```
Na primeira execução, abra o Chrome **manualmente** uma vez com este perfil e faça login na conta Google que tem acesso ao Classroom:
```bash
chrome.exe --user-data-dir="C:\selenium\chrome-profile"
```
*(Certifique-se de fechar este navegador antes de iniciar as automações).*

---

## ▶️ Executando o projeto

Inicie ambos os servidores (Frontend e Backend) com um único comando na raiz do projeto:

```bash
npm run dev
```

* O **Frontend** iniciará em: `http://localhost:5173`
* O **Backend** Flask iniciará em: `http://localhost:5000`

---

## 🖥️ Como usar

### Passo 1 — Extrair Rubricas (Classroom)
Você pode escolher entre duas formas de extração na interface:

* **Modo Oculto (Recomendado):** 
  1. Copie a URL da atividade do Classroom direto da barra de endereço do seu navegador.
  2. Cole-a no campo **"URL da Atividade (Classroom)"** na interface da ferramenta.
  3. Clique em **"1. Extrair Rubricas"**. O processo ocorrerá inteiramente em segundo plano sem abrir nenhuma janela na sua tela.
  *(Caso você não esteja autenticado, o robô abrirá o navegador uma única vez para você logar e continuará o processo sozinho).*
* **Modo Visível (Manual):**
  1. Deixe o campo de URL em branco.
  2. Clique em **"1. Extrair Rubricas"**. O navegador se abrirá fisicamente.
  3. Navegue até a atividade desejada no Classroom e clique em **"Confirmar Página do Classroom"** no painel da ferramenta.

O sistema processará os alunos que entregaram e salvará as informações em `rubricas.json` na raiz do projeto.

### Passo 2 — Preencher Planilha (Excel)
1. Abra a planilha Excel no seu computador.
2. Na ferramenta, clique em **"2. Preencher Planilha"**.
3. O sistema detectará os alunos avaliados e solicitará que você defina a **autonomia** de cada um.
4. Clique em **"Confirmar e Preencher"**. O robô fará uma varredura nas linhas, apagará marcações anteriores de notas/capacidades naquela linha para evitar duplicidade, e preencherá as novas notas.

---

## 🔧 Variáveis de configuração da Planilha

As configurações de layout da planilha estão centralizadas no topo de `backend/preencher_planilha.py`:

```python
COLUNA_CRITERIO      = 7   # Coluna G (Onde pesquisa o título dos critérios)
COL_INICIO_AVALIACAO = 8   # Coluna H (Posição de nota 1)
COL_FIM_AVALIACAO    = 27  # Coluna AA (Posição de nota 20)
LINHA_INICIAL_BUSCA  = 21  # Primeira linha de busca na aba do aluno
LINHA_NOME_ALUNO     = 10  # Linha onde fica o nome do aluno
COLUNA_NOME_ALUNO    = 13  # Coluna M (Onde fica o nome do aluno)
```

Caso o design ou formato da planilha mude, basta atualizar estes valores.
