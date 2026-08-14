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

def iniciar_driver(headless=True):
    """
    Inicia o navegador Chrome com configurações personalizadas.

    O objetivo principal aqui é:
    - reutilizar um perfil já existente do Chrome
    - abrir o navegador oculto ou maximizado conforme parâmetro
    - ocultar a flag de automação para evitar bloqueios do Classroom

    Retorno:
        webdriver.Chrome: instância do navegador pronta para uso.
    """

    # Cria um objeto de opções para configurar o Chrome antes de iniciar
    options = Options()

    # Define o diretório do perfil de usuário do Chrome que será reutilizado.
    options.add_argument(r"--user-data-dir=C:\selenium\chrome-profile")

    if headless:
        # Configura o modo headless (oculto) para rodar em segundo plano sem abrir janela física
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
    else:
        # Abre em modo visível e maximizado para permitir interação/login manual
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
    de cada aluno, classificando-os por status (bucket) e priorizando os devolvidos.
    """
    alunos = []

    # Encontra todas as linhas de cabeçalho e de alunos na tabela ordenadamente
    linhas = driver.find_elements(By.CSS_SELECTOR, 'tr.tYQn5c, tr.DC55n, tr[data-bucket], tr[data-student-id]')
    
    print(f"Linhas totais na tabela de alunos: {len(linhas)}")
    current_bucket = "unknown"

    for linha in linhas:
        try:
            classes = linha.get_attribute("class") or ""
            bucket_attr = linha.get_attribute("data-bucket")
            
            # Se for uma linha de cabeçalho (atribuído, entregue, devolvido, etc.)
            if "tYQn5c" in classes or bucket_attr:
                current_bucket = bucket_attr or current_bucket
                continue
            
            # Se for uma linha de aluno
            if "DC55n" in classes or linha.get_attribute("data-student-id"):
                link = linha.find_element(By.CSS_SELECTOR, 'td.TAjiIf a[href*="/student/"]')
                nome = linha.find_element(By.CSS_SELECTOR, 'span.YVvGBb').text.strip()
                href = link.get_attribute("href")

                if nome and href:
                    alunos.append({
                        "nome": nome,
                        "href": href,
                        "bucket": current_bucket
                    })
        except Exception as e:
            # Silenciosamente tenta a próxima linha
            pass

    print(f"Alunos mapeados no total: {len(alunos)}")
    
    # Define a ordenação: "returned" (devolvido) primeiro, depois outros, "not-done" (atribuído/não feito) por último
    def obter_prioridade(aluno):
        b = aluno.get("bucket", "unknown")
        if b == "returned":
            return 0
        elif b == "graded":
            return 1
        elif b == "not-done":
            return 3
        return 2

    alunos.sort(key=obter_prioridade)
    # Filtra para manter apenas os alunos dos blocos "Atividades devolvidas" (returned) e "Com nota" (graded)
    alunos_filtrados = [a for a in alunos if a.get("bucket") in ("returned", "graded")]
    print(f"Alunos filtrados (apenas devolvidos/com nota): {len(alunos_filtrados)} de {len(alunos)} no total.")
    return alunos_filtrados

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

    # Rolagem incremental do painel de notas (barra lateral direita que contém a rubrica)
    try:
        # Encontra o container interno com overflow do Classroom que realmente rola na direita
        container_found = driver.execute_script("""
            const crit = document.querySelector('[data-criterion-id]');
            if (!crit) return null;
            let parent = crit.parentElement;
            while (parent && parent !== document.body) {
                const style = window.getComputedStyle(parent);
                const overflow = style.overflow + style.overflowY;
                if (parent.scrollHeight > parent.clientHeight && (overflow.includes('auto') || overflow.includes('scroll'))) {
                    parent.classList.add('selenium-scroll-container');
                    return true;
                }
                parent = parent.parentElement;
            }
            return false;
        """)

        if container_found:
            altura_total = driver.execute_script("return document.querySelector('.selenium-scroll-container').scrollHeight")
            posicao_atual = 0
            passo = 250
            while posicao_atual < altura_total:
                posicao_atual += passo
                driver.execute_script(f"document.querySelector('.selenium-scroll-container').scrollTop = {posicao_atual};")
                time.sleep(0.15)
                altura_total = driver.execute_script("return document.querySelector('.selenium-scroll-container').scrollHeight")
            
            driver.execute_script("document.querySelector('.selenium-scroll-container').scrollTop = 0;")
            time.sleep(0.3)
            print("Rolagem incremental da barra lateral de notas concluída.")
        else:
            # Fallback scroll básico se não achar o container específico
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.3)
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.3)
    except Exception as e:
        print(f"Aviso: Falha na rolagem do painel de critérios: {e}")

    # Clica em todos os botões/cabeçalhos de critério que estão colapsados (aria-expanded="false")
    # Nota: Usamos apenas seletores explícitos com aria-expanded="false" para garantir que NUNCA 
    # cliquemos nas opções de nota/níveis (que são botões comuns ou radio buttons) e evitamos alterar avaliações.
    try:
        driver.execute_script("""
            // Filtra para pegar apenas os blocos que estão visíveis na tela (do aluno ativo)
            const blocos = Array.from(document.querySelectorAll('[data-criterion-id]'))
                                .filter(el => el.offsetHeight > 0);
            blocos.forEach(bloco => {
                // Para segurança absoluta, selecionamos apenas elementos 'button' que controlam painéis (aria-controls)
                // e estão fechados (aria-expanded="false"). Botões de atribuição de nota não possuem essas duas propriedades juntas.
                const btn = bloco.querySelector('button[aria-expanded="false"][aria-controls]');
                if (btn) {
                    try {
                        btn.click();
                    } catch(e) {
                        console.error("Erro ao expandir critério:", e);
                    }
                }
            });
        """)
        print("Critérios colapsados expandidos programaticamente de forma segura (um clique por critério).")
        time.sleep(0.8)
    except Exception as e:
        print(f"Erro ao tentar expandir os critérios: {e}")

    # Agora espera que as opções de nível sejam renderizadas após a expansão
    try:
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, '[role="menuitemradio"]')) > 0)
    except Exception as e:
        print(f"Opções de nível não carregaram após expansão: {e}")

    # Outra rolagem rápida no container de rolagem após expandir para carregar dados virtuais
    try:
        driver.execute_script("""
            const sc = document.querySelector('.selenium-scroll-container');
            if (sc) {
                sc.scrollTop = sc.scrollHeight * 0.5;
                setTimeout(() => { sc.scrollTop = 0; }, 150);
            }
        """)
        time.sleep(0.3)
    except Exception as e:
        pass

    # Espera inteligente para ver se alguma opção está marcada (comportamento assíncrono do Classroom ao carregar nota)
    try:
        wait_marcado = WebDriverWait(driver, 2)
        wait_marcado.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, '[role="menuitemradio"][aria-checked="true"], [role="menuitemradio"].KKjvXb')) > 0)
        print("Opção marcada encontrada/carregada com sucesso!")
    except Exception:
        print("Nenhuma opção marcada carregada após 2 segundos (aluno pode não ter sido avaliado ainda).")

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

        // Filtramos para pegar apenas blocos visíveis na tela (do aluno ativo)
        // Evita ler elementos de alunos anteriores que continuam ocultos no DOM do Classroom
        const blocos = Array.from(document.querySelectorAll('[data-criterion-id]'))
                            .filter(el => el.offsetHeight > 0);

        for (const bloco of blocos) {
            const criterionId = (bloco.getAttribute('data-criterion-id') || '').trim();
            if (!criterionId || vistos.has(criterionId)) continue;

            // Busca o título do critério de forma resiliente
            let criterio = '';
            
            // Prioridade 1: Tenta obter o span específico que contém o texto principal no Classroom (jsname="V67aGc")
            const tituloEl = bloco.querySelector('span[jsname="V67aGc"], .mUIrbf-vQzf8d');
            if (tituloEl) {
                criterio = (tituloEl.innerText || tituloEl.textContent || '').replace(/\\s+/g, ' ').trim();
            }

            // Prioridade 2: Tenta cabeçalho h2, h3, h4 ou similar
            if (!criterio) {
                const cabecalho = bloco.querySelector('h2, h3, h4, [role="heading"], .K0lUWd');
                if (cabecalho && !cabecalho.innerText.toLowerCase().includes('rubrica')) {
                    criterio = (cabecalho.innerText || cabecalho.textContent || '').replace(/\\s+/g, ' ').trim();
                }
            }

            // Fallback: pega o primeiro texto válido, ignorando tooltips e botões com a frase "Expandir critério"
            if (!criterio) {
                const primeiroTexto = Array.from(bloco.querySelectorAll('div, span, button'))
                    .filter(el => !el.matches('[role="tooltip"], .ne2Ple-oshW8e-V67aGc, [aria-label*="Expandir"]'))
                    .map(el => (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim())
                    .find(t => t.length > 0 && t.toLowerCase() !== 'expandir critério' && t.toLowerCase() !== 'expandir criterio');
                if (primeiroTexto) {
                    criterio = primeiroTexto;
                } else {
                    criterio = 'Critério não encontrado';
                }
            }

            let nivel = 'Nenhum nível marcado';
            let opcoes = [];
            
            // 1. Busca opções diretamente dentro do próprio bloco do critério (escopo restrito ao critério)
            opcoes = Array.from(bloco.querySelectorAll('[role="menuitemradio"], [role="radio"], [aria-checked]'))
                          .filter(el => el.offsetHeight > 0);
            
            // 2. Fallback: Tenta achar via aria-controls se não encontrou dentro do bloco
            if (opcoes.length === 0) {
                const botaoControlador = bloco.querySelector('[aria-controls]') || 
                                         (bloco.getAttribute('aria-controls') ? bloco : null);
                if (botaoControlador) {
                    const panelId = botaoControlador.getAttribute('aria-controls');
                    const painel = document.getElementById(panelId);
                    if (painel) {
                        opcoes = Array.from(painel.querySelectorAll('[role="menuitemradio"], [role="radio"], [aria-checked]'))
                                      .filter(el => el.offsetHeight > 0);
                    }
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
        opcoes_formatadas = []
        for op in item.get('opcoes_debug', []):
            is_marcada = op['atributos'].get('aria-checked') == 'true' or op['atributos'].get('aria-selected') == 'true' or 'KKjvXb' in op['classes']
            marcador = " (MARCADA)" if is_marcada else ""
            opcoes_formatadas.append(f"'{op['texto']}'{marcador}")
        
        print(f"Critério {i}: {item['criterio']} => {item['nivel']}")
        print(f"  Níveis lidos: [{', '.join(opcoes_formatadas)}]")

    # Remove o campo debug antes de retornar para não poluir o JSON final, limpa o texto do critério e filtra vazios
    import re
    dados_limpos = []
    for item in dados:
        if 'opcoes_debug' in item:
            del item['opcoes_debug']
        
        criterio_original = item.get("criterio", "")
        # Remove "Expandir critério" / "Expandir criterio" case-insensitively
        criterio_limpo = re.sub(r'(?i)expandir\s+crit[eé]rio', '', criterio_original)
        criterio_limpo = " ".join(criterio_limpo.split()).strip()
        
        if criterio_limpo and criterio_limpo.lower() != "critério não encontrado":
            item["criterio"] = criterio_limpo
            dados_limpos.append(item)
        else:
            print(f"[INFO] Removido item de controle ou vazio após limpeza: '{criterio_original}'")

    return dados_limpos


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


def extrair_rubricas_todos_alunos(url=None, confirmar_inicio_callback=None, should_stop_callback=None):
    # Se uma URL foi fornecida, tentamos rodar oculto (headless=True), caso contrário visível
    headless_mode = True if url else False
    driver = iniciar_driver(headless=headless_mode)

    try:
        if url:
            print("Verificando status de autenticação no Google Classroom...")
            driver.get("https://classroom.google.com/")
            time.sleep(2.5)

            # Verifica se foi redirecionado para a tela de login do Google Accounts
            if "accounts.google.com" in driver.current_url or "signin" in driver.current_url:
                print("\n[AVISO] Você não está logado! Abrindo janela do navegador para você realizar o login...")
                driver.quit()
                
                # Abre o navegador visível para o usuário interagir
                driver = iniciar_driver(headless=False)
                driver.get("https://classroom.google.com/")
                
                # Fica em loop monitorando a URL até que o usuário saia das páginas de login
                while "accounts.google.com" in driver.current_url or "signin" in driver.current_url:
                    time.sleep(1)
                    if should_stop_callback is not None and should_stop_callback():
                        print("Processo abortado durante a tela de login.")
                        return
                
                print("[SISTEMA] Login detectado com sucesso! Prosseguindo para a extração...")
                time.sleep(1.5)

            print(f"Navegando diretamente para a URL da atividade: {url}")
            driver.get(url)
            print("Aguardando carregamento da tabela de estudantes (5 segundos)...")
            time.sleep(5)
        else:
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

                # Validação crítica das rubricas e níveis
                rubrica_extraida = resultado.get("rubrica", [])
                
                # Se não há rubrica ou se houver critérios com nível inválido/vazio
                niveis_invalidos = [
                    item for item in rubrica_extraida
                    if item.get("nivel") == "Nenhum nível marcado" or not item.get("nivel")
                ]
                
                if not rubrica_extraida:
                    raise ValueError("Nenhum critério foi extraído da página.")
                
                if niveis_invalidos:
                    raise ValueError("Não foi encontrado nível (nota) preenchido para um ou mais critérios.")

                resultados.append(resultado)
                print(f"Rubrica extraída com sucesso de: {aluno['nome']}")

            except Exception as e:
                # Evidencia o erro e interrompe a automação na hora
                print(f"\n[ERRO CRÍTICO] Falha na extração do aluno {aluno['nome']}: {e}")
                raise e

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