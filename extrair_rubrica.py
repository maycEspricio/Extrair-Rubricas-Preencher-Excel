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
    - ocultar a flag de automação para evitar bloqueios do Classroom

    Retorno:
        webdriver.Chrome: instância do navegador pronta para uso.
    """

    # Cria um objeto de opções para configurar o Chrome antes de iniciar
    options = Options()

    # Define o diretório do perfil de usuário do Chrome que será reutilizado.
    options.add_argument(r"--user-data-dir=C:\selenium\chrome-profile")

    # Faz o navegador abrir maximizado para facilitar a visualização
    options.add_argument("--start-maximized")

    # Oculta a automação do Selenium para evitar bloqueio do Classroom
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=options)

    # Executa script CDP para ocultar o navigator.webdriver
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })

    return driver

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

def extrair_rubrica_do_aluno(driver, nome_aluno, href, force_navigation=False):
    """
    Seleciona o aluno (por clique físico ou URL), expande a rubrica se necessário,
    espera o carregamento das notas e extrai a rubrica.
    """
    print(f"\nSelecionando aluno: {nome_aluno} (Navegação forçada: {force_navigation})")
    
    # Extrai o ID do estudante da URL para sincronização de SPA
    import re
    match_id = re.search(r'/student/([A-Za-z0-9]+)', href)
    student_id = match_id.group(1) if match_id else None

    clicado = False
    if not force_navigation:
        try:
            # Busca elementos de texto que possam conter o nome do aluno na tela
            normalized_name = nome_aluno.strip().lower()
            xpath_expr = f"//*[(self::span or self::div or self::a) and contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÀÈÌÒÙÁÉÍÓÚÂÊÎÔÛÃÕÇ', 'abcdefghijklmnopqrstuvwxyzàèìòùáéíóúâêîôûãõç'), '{normalized_name}')]"
            candidatos = driver.find_elements(By.XPATH, xpath_expr)
            
            if not candidatos:
                # Fallback secundário buscando span com classe padrão
                candidatos = [el for el in driver.find_elements(By.CSS_SELECTOR, "span.YVvGBb") if normalized_name in el.text.lower()]
                
            for c in candidatos:
                if c.is_displayed():
                    # Centraliza o elemento na tela
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", c)
                    time.sleep(0.3)
                    
                    # Encontra o ancestral clicável correto via JS (A, BUTTON, role=button, etc)
                    clickable_element = driver.execute_script("""
                        let start = arguments[0];
                        let p = start;
                        for (let i = 0; i < 6; i++) {
                            if (!p) break;
                            if (p.tagName === 'A' || 
                                p.tagName === 'BUTTON' || 
                                p.getAttribute('role') === 'button' || 
                                p.getAttribute('role') === 'option' || 
                                p.getAttribute('role') === 'listitem' ||
                                p.getAttribute('tabindex') === '0' ||
                                p.tagName === 'TR') {
                                return p;
                            }
                            p = p.parentElement;
                        }
                        return start;
                    """, c)
                    
                    if clickable_element:
                        # Executa clique físico/real via ActionChains para garantir que o Classroom processe o clique
                        from selenium.webdriver.common.action_chains import ActionChains
                        actions = ActionChains(driver)
                        actions.move_to_element(clickable_element).click().perform()
                        clicado = True
                        print(f"Clique físico executado com sucesso no aluno {nome_aluno}.")
                        break
        except Exception as e:
            print(f"Erro ao tentar clicar na barra lateral para o aluno {nome_aluno}: {e}")

    if not clicado:
        # Navega diretamente pelo link (Primeiro load ou caso clique falhe)
        print(f"Navegando via link: {href}")
        driver.get(href)
        
    # Sincronização de SPA: Aguarda a URL mudar para a do aluno selecionado
    if student_id:
        try:
            WebDriverWait(driver, 10).until(EC.url_contains(student_id))
            print(f"URL confirmada para o aluno ID: {student_id}")
        except Exception:
            print(f"Aviso: A URL não atualizou para o ID {student_id} a tempo.")

    # Tempo de segurança para o DOM do aluno anterior descarregar e o novo renderizar
    time.sleep(1.5)

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

    # Clica em todos os botões/cabeçalhos de critério que estão colapsados (aria-expanded="false")
    try:
        driver.execute_script("""
            const botoes = Array.from(document.querySelectorAll('[data-criterion-id] [aria-expanded="false"], [data-criterion-id][aria-expanded="false"], [aria-controls][aria-expanded="false"]'));
            botoes.forEach(b => {
                try {
                    b.click();
                } catch(e) {
                    console.error("Erro ao expandir critério:", e);
                }
            });
        """)
        print("Critérios colapsados expandidos programaticamente.")
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao tentar expandir os critérios: {e}")

    # Agora espera que as opções de nível sejam renderizadas após a expansão
    try:
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, '[role="menuitemradio"]')) > 0)
    except Exception as e:
        print(f"Opções de nível não carregaram após expansão: {e}")

    # scroll leve para incentivar renderização de todos os elementos e carregamento de estados
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.4);")
    time.sleep(0.5)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.5)

    # Espera inteligente para ver se alguma opção está marcada (comportamento assíncrono do Classroom ao carregar nota)
    try:
        wait_marcado = WebDriverWait(driver, 6)
        wait_marcado.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, '[role="menuitemradio"][aria-checked="true"], [role="menuitemradio"].KKjvXb')) > 0)
        print("Opção marcada encontrada/carregada com sucesso!")
    except Exception:
        print("Nenhuma opção marcada carregada após 6 segundos (aluno pode não ter sido avaliado ainda).")

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
    Extrai critérios e níveis diretamente do DOM de forma resiliente a mudanças de classes do Classroom.
    """
    dados = driver.execute_script("""
        const resultado = [];
        const vistos = new Set();

        const blocos = Array.from(document.querySelectorAll('[data-criterion-id]')).filter(el => el.offsetHeight > 0);

        for (const bloco of blocos) {
            const criterionId = (bloco.getAttribute('data-criterion-id') || '').trim();
            if (!criterionId || vistos.has(criterionId)) continue;

            // Busca o título do critério de forma resiliente
            let criterio = 'Critério não encontrado';
            const cabecalho = bloco.querySelector('h2, h3, h4, [role="heading"], .K0lUWd');
            if (cabecalho) {
                criterio = (cabecalho.innerText || cabecalho.textContent || '').replace(/\\s+/g, ' ').trim();
            } else {
                // Fallback: pega o primeiro texto curto no topo do bloco
                const primeiroTexto = Array.from(bloco.querySelectorAll('div, span'))
                    .map(el => (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim())
                    .find(t => t.length > 0 && t.length < 100);
                if (primeiroTexto) {
                    criterio = primeiroTexto;
                }
            }

            let nivel = 'Nenhum nível marcado';
            let opcoes = [];
            
            // 1. Tenta achar via aria-controls (apenas visíveis)
            const botaoControlador = bloco.querySelector('[aria-controls]') || 
                                     (bloco.getAttribute('aria-controls') ? bloco : null);
            if (botaoControlador) {
                const panelId = botaoControlador.getAttribute('aria-controls');
                const painel = document.getElementById(panelId);
                if (painel) {
                    opcoes = Array.from(painel.querySelectorAll('[role="menuitemradio"], [role="radio"], [role="button"], [aria-checked]'))
                                  .filter(el => el.offsetHeight > 0);
                }
            }
            
            // 2. Se não achar, procura subindo a árvore DOM até encontrar um container com opções (apenas visíveis)
            if (opcoes.length === 0) {
                let container = bloco.parentElement;
                while (container && container.tagName !== 'BODY') {
                    const ops = Array.from(container.querySelectorAll('[role="menuitemradio"], [role="radio"], [aria-checked]'))
                                     .filter(el => el.offsetHeight > 0);
                    if (ops.length > 0) {
                        opcoes = ops;
                        break;
                    }
                    container = container.parentElement;
                }
            }

            // Procura diretamente a opção marcada entre as encontradas
            let marcada = opcoes.find(el => {
                // Verifica se o próprio elemento está marcado
                const selfChecked = (el.getAttribute('aria-checked') || '').toLowerCase() === 'true' || 
                                     (el.getAttribute('aria-selected') || '').toLowerCase() === 'true' ||
                                     el.classList.contains('KKjvXb');
                if (selfChecked) return true;

                // Verifica se algum filho/descendente está marcado (comum no Classroom estruturado)
                const descChecked = el.querySelector('[aria-checked="true"], [aria-selected="true"], .KKjvXb');
                if (descChecked) return true;

                return false;
            });

            if (marcada) {
                // Extrai o texto do aria-label ou data-tooltip (já que o innerText costuma vir vazio no DOM)
                nivel =
                    (marcada.getAttribute('aria-label') || '').trim() ||
                    (marcada.getAttribute('data-tooltip') || '').trim() ||
                    (marcada.innerText || marcada.textContent || '').replace(/\\s+/g, ' ').trim() ||
                    'Nível sem rótulo';
            }

            // Coleta diagnóstico das opções para depuração no console do Python
            const opcoesDebug = opcoes.map(el => {
                const attrs = {};
                for (let i = 0; i < el.attributes.length; i++) {
                    attrs[el.attributes[i].name] = el.attributes[i].value;
                }
                return {
                    texto: (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim(),
                    classes: el.className || '',
                    atributos: attrs,
                    html: el.outerHTML
                };
            });

            vistos.add(criterionId);

            resultado.push({
                criterion_id: criterionId,
                criterio: criterio,
                nivel: nivel,
                opcoes_debug: opcoesDebug
            });
        }

        return resultado;
    """)

    print(f"Critérios válidos extraídos: {len(dados)}")
    for i, item in enumerate(dados, start=1):
        print(f"Critério {i}: {item['criterio']} => {item['nivel']}")
        print(f"  Diagnóstico de opções do Critério {i}:")
        for op in item.get('opcoes_debug', []):
            is_marcada = op['atributos'].get('aria-checked') == 'true' or op['atributos'].get('aria-selected') == 'true' or 'KKjvXb' in op['classes']
            marcador = " [MARCADA]" if is_marcada else ""
            print(f"    - Nível: '{op['texto']}' | Classes: '{op['classes']}' | Atributos: {op['atributos']}{marcador}")
            print(f"      HTML Completo da Opção: {op['html']}")

    # Remove o campo debug antes de retornar para não poluir o JSON final
    for item in dados:
        if 'opcoes_debug' in item:
            del item['opcoes_debug']

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
                # O primeiro aluno sempre força a navegação via URL para abrir a tela de correção.
                # Os próximos navegam via cliques na barra lateral (mais rápido e evita bugs).
                resultado = extrair_rubrica_do_aluno(
                    driver,
                    aluno["nome"],
                    aluno["href"],
                    force_navigation=(i == 1)
                )

                # Workaround para o bug do Classroom: se for o primeiro aluno e a rubrica não carregou (veio vazia),
                # nós clicamos no segundo aluno para acordar a interface e depois voltamos para o primeiro.
                if i == 1 and (not resultado.get("rubrica") or "não encontrados" in resultado.get("observacao", "")) and len(alunos) > 1:
                    print("\n[WORKAROUND] Rubrica do primeiro aluno não carregou no load inicial.")
                    print("[WORKAROUND] Selecionando o segundo aluno temporariamente para ativar o painel...")
                    # Seleciona o segundo aluno via clique
                    extrair_rubrica_do_aluno(driver, alunos[1]["nome"], alunos[1]["href"], force_navigation=False)
                    time.sleep(1.5)
                    
                    print("[WORKAROUND] Retornando ao primeiro aluno para nova tentativa...")
                    # Tenta extrair o primeiro aluno novamente via clique
                    resultado = extrair_rubrica_do_aluno(
                        driver,
                        aluno["nome"],
                        aluno["href"],
                        force_navigation=False
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