# Importa o módulo principal do Selenium responsável por controlar o navegador
from selenium import webdriver

# Importa a classe de opções do Chrome, usada para configurar o navegador antes de iniciá-lo
from selenium.webdriver.chrome.options import Options

# Importa a classe By, utilizada para localizar elementos na página
# por CSS_SELECTOR, XPATH, ID, NAME etc.
from selenium.webdriver.common.by import By

# Importa WebDriverWait para esperas explícitas
# (neste código ainda não está sendo usado, mas pode ser útil futuramente)
from selenium.webdriver.support.ui import WebDriverWait

# Importa condições esperadas para usar junto com WebDriverWait
# (também não está sendo usado no fluxo atual)
from selenium.webdriver.support import expected_conditions as EC

# Importa o BeautifulSoup para fazer o parsing do HTML e facilitar a extração de dados
from bs4 import BeautifulSoup

# Importa o módulo time para fazer pausas durante a automação
import time

# Importa o módulo json para salvar os dados extraídos em arquivo .json
import json

def iniciar_driver():
    """
    Inicia o navegador Chrome com configurações personalizadas.

    O objetivo principal aqui é:
    - reutilizar um perfil já existente do Chrome
    - abrir o navegador maximizado

    Retorno:
        webdriver.Chrome: instância do navegador pronta para uso.
    """

    # Cria um objeto de opções para configurar o Chrome antes de iniciar
    options = Options()

    # Define o diretório do perfil de usuário do Chrome que será reutilizado.
    # Isso é útil para manter sessão logada, cookies e configurações já existentes.
    # Exemplo: evita precisar fazer login no Google Classroom toda vez.
    options.add_argument(r"--user-data-dir=C:\selenium\chrome-profile")

    # Faz o navegador abrir maximizado para facilitar a visualização
    # e evitar problemas com elementos escondidos por layout responsivo.
    options.add_argument("--start-maximized")

    # Retorna a instância do Chrome já configurada
    return webdriver.Chrome(options=options)

def coletar_alunos_da_tabela(driver):
    """
    Lê a tabela de estudantes e retorna uma lista com nome e link
    de cada aluno.
    """
    alunos = []

    # pega apenas linhas que realmente representam aluno
    linhas = driver.find_elements(By.CSS_SELECTOR, 'tr[data-student-id]')

    print(f"Alunos encontrados na tabela: {len(linhas)}")

    for linha in linhas:
        try:
            link = linha.find_element(By.CSS_SELECTOR, 'td.TAjiIf a[href*="/student/"]')
            nome = linha.find_element(By.CSS_SELECTOR, 'span.YVvGBb').text.strip()
            href = link.get_attribute("href")

            if nome and href:
                alunos.append({
                    "nome": nome,
                    "href": href
                })

        except Exception as e:
            print(f"Não consegui ler uma linha da tabela: {e}")

    return alunos

def esperar_rubrica_carregar(driver, timeout=20):
    """
    Aguarda a página do aluno carregar e diagnostica em qual etapa a rubrica falha.
    """
    wait = WebDriverWait(driver, timeout)

    try:
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        print("Etapa 1 OK: document.readyState = complete")
    except Exception as e:
        print(f"Falha na etapa 1 (document.readyState): {e}")
        return False

    try:
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, '[data-criterion-id]')) > 0)
        qtd_criterios = len(driver.find_elements(By.CSS_SELECTOR, '[data-criterion-id]'))
        print(f"Etapa 2 OK: critérios encontrados = {qtd_criterios}")
    except Exception as e:
        print(f"Falha na etapa 2 (data-criterion-id): {e}")
        return False

    try:
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, '[role="menuitemradio"]')) > 0)
        qtd_opcoes = len(driver.find_elements(By.CSS_SELECTOR, '[role="menuitemradio"]'))
        print(f"Etapa 3 OK: opções de nível encontradas = {qtd_opcoes}")
    except Exception as e:
        print(f"Falha na etapa 3 (menuitemradio): {e}")
        return False

    print("Rubrica carregada com sucesso.")
    return True

def extrair_rubrica_do_aluno(driver, nome_aluno, href):
    """
    Abre a página do aluno, extrai a rubrica
    e devolve um dicionário com os dados.
    """
    print(f"\nAbrindo aluno: {nome_aluno}")
    driver.get(href)

    wait = WebDriverWait(driver, 20)

    try:
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, '[data-criterion-id]')) > 0)
    except Exception as e:
        print(f"Página do aluno carregou, mas os critérios não ficaram disponíveis: {e}")
        return {
            "aluno": nome_aluno,
            "link": href,
            "rubrica": [],
            "observacao": "Critérios da rubrica não encontrados."
        }

    time.sleep(1)

    # scroll leve para incentivar renderização
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.4);")
    time.sleep(0.5)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.5)

    html = driver.page_source

    print(f"HTML capturado com {len(html)} caracteres")
    print(f"Quantidade de 'data-criterion-id' no HTML: {html.count('data-criterion-id')}")
    print(f"Quantidade de 'menuitemradio' no HTML: {html.count('menuitemradio')}")
    print(f"Quantidade de 'aria-checked=\"true\"' no HTML: {html.count('aria-checked=\"true\"')}")

    rubrica = copiar_rubrica(driver)

    return {
        "aluno": nome_aluno,
        "link": href,
        "rubrica": rubrica
    }

def copiar_rubrica(driver):
    """
    Extrai critérios e níveis diretamente do DOM usando a ligação
    entre o botão do critério (aria-controls) e o painel de opções (id).
    """
    dados = driver.execute_script("""
        const resultado = [];
        const vistos = new Set();

        const blocos = Array.from(document.querySelectorAll('.stS1kf.Lzvjbf[data-criterion-id], [data-criterion-id]'));

        for (const bloco of blocos) {
            const criterionId = (bloco.getAttribute('data-criterion-id') || '').trim();
            if (!criterionId || vistos.has(criterionId)) continue;

            let criterio = 'Critério não encontrado';

            // tenta achar o texto do critério dentro do bloco
            const candidatosTitulo = Array.from(bloco.querySelectorAll('.K0lUWd'));
            const tituloValido = candidatosTitulo
                .map(el => (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim())
                .find(t => t.length > 0);

            if (tituloValido) {
                criterio = tituloValido;
            }

            // acha o botão do critério que aponta para o painel correto
            let botaoControlador = Array.from(bloco.querySelectorAll('[aria-controls]'))
                .find(el => (el.getAttribute('aria-controls') || '').trim().length > 0);

            // fallback: procurar próximo do bloco
            if (!botaoControlador && bloco.parentElement) {
                botaoControlador = Array.from(bloco.parentElement.querySelectorAll('[aria-controls]'))
                    .find(el => (el.getAttribute('aria-controls') || '').trim().length > 0);
            }

            let nivel = 'Nenhum nível marcado';
            let opcoes = [];

            if (botaoControlador) {
                const panelId = (botaoControlador.getAttribute('aria-controls') || '').trim();
                if (panelId) {
                    const painel = document.getElementById(panelId);
                    if (painel) {
                        opcoes = Array.from(painel.querySelectorAll('[role="menuitemradio"]'));
                    }
                }
            }

            // fallback final: procura no contêiner visual mais próximo
            if (opcoes.length === 0) {
                const container = bloco.closest('.NBQ1Tb') || bloco.parentElement;
                if (container) {
                    opcoes = Array.from(container.querySelectorAll('[role="menuitemradio"]'));
                }
            }

            if (opcoes.length === 0) {
                continue;
            }

            let marcada = opcoes.find(el => (el.getAttribute('aria-checked') || '').toLowerCase() === 'true');

            if (!marcada) {
                marcada = opcoes.find(el => el.classList.contains('KKjvXb'));
            }

            if (marcada) {
                nivel =
                    (marcada.getAttribute('aria-label') || '').trim() ||
                    (marcada.innerText || marcada.textContent || '').replace(/\\s+/g, ' ').trim() ||
                    'Nível sem rótulo';
            }

            vistos.add(criterionId);

            resultado.push({
                criterion_id: criterionId,
                criterio: criterio,
                nivel: nivel
            });
        }

        return resultado;
    """)

    print(f"Critérios válidos extraídos: {len(dados)}")
    for i, item in enumerate(dados, start=1):
        print(f"Critério {i}: {item['criterio']} => {item['nivel']}")

    return dados


def salvar_resultados(dados, nome_arquivo="rubrica.json"):
    """
    Salva os dados extraídos em um arquivo JSON.

    Parâmetros:
        dados: estrutura de dados que será salva no arquivo.
        nome_arquivo (str): nome do arquivo de saída.
    """

    # Abre o arquivo em modo de escrita com codificação UTF-8
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        # Salva os dados em formato JSON com:
        # - ensure_ascii=False: preserva acentos
        # - indent=2: deixa o JSON formatado e legível
        json.dump(dados, f, ensure_ascii=False, indent=2)


def extrair_rubricas_todos_alunos(confirmar_inicio_callback=None, should_stop_callback=None):
    driver = iniciar_driver()

    try:
        driver.get("https://classroom.google.com/")
        if confirmar_inicio_callback is None:
            input("Abra a página da atividade com a tabela de estudantes e pressione ENTER... ")
        else:
            # Quando usado por interface gráfica, a confirmação vem por callback
            # (ex.: botão/caixa de diálogo), evitando travar em input().
            deve_continuar = confirmar_inicio_callback()
            if deve_continuar is False:
                print("Processo interrompido pelo usuário antes da extração.")
                return

        time.sleep(2)

        # coleta todos os alunos da tabela
        alunos = coletar_alunos_da_tabela(driver)

        if not alunos:
            print("Nenhum aluno encontrado na tabela.")
            return

        resultados = []

        for i, aluno in enumerate(alunos, start=1):
            # Permite cancelamento cooperativo quando a função é chamada pela interface.
            # Se o usuário clicar em "Parar", interrompemos antes de abrir o próximo aluno.
            if should_stop_callback is not None and should_stop_callback():
                print("Processo interrompido pelo usuário durante a extração.")
                break

            print(f"\n===== Aluno {i}/{len(alunos)} =====")

            try:
                resultado = extrair_rubrica_do_aluno(
                    driver,
                    aluno["nome"],
                    aluno["href"]
                )
                resultados.append(resultado)

                print(f"Rubrica extraída de: {aluno['nome']}")

            except Exception as e:
                print(f"Erro ao extrair rubrica de {aluno['nome']}: {e}")
                resultados.append({
                    "aluno": aluno["nome"],
                    "link": aluno["href"],
                    "erro": str(e)
                })

        # mostra no terminal
        print("\nRESULTADO FINAL:\n")
        for aluno in resultados:
            print(f"\nAluno: {aluno.get('aluno')}")
            if "rubrica" in aluno:
                for item in aluno["rubrica"]:
                    print(f"  - {item['criterio']} => {item['nivel']}")
            else:
                print(f"  Erro: {aluno.get('erro')}")

        salvar_resultados(resultados, "rubricas.json")

    finally:
        driver.quit()


# Este bloco garante que a função extrair_rubricas_todos_alunos() só será executada
# automaticamente quando este arquivo for rodado diretamente.
# Se o arquivo for importado em outro módulo (como main.py), esse trecho não roda.
if __name__ == "__main__":
    extrair_rubricas_todos_alunos()