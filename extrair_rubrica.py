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


def expandir_criterios(driver):
    """
    Procura todos os botões de expandir critério na página e tenta clicar em cada um.

    Isso é necessário porque alguns dados da rubrica podem estar ocultos
    até que o critério seja expandido visualmente.

    Parâmetros:
        driver: instância ativa do Selenium WebDriver.
    """

    # Busca todos os elementos que possuam aria-label="Expandir critério"
    # usando seletor CSS.
    botoes = driver.find_elements(By.CSS_SELECTOR, '[aria-label="Expandir critério"]')

    # Exibe quantos botões foram encontrados para depuração
    print(f"Botões de expandir encontrados: {len(botoes)}")

    # Percorre todos os botões encontrados
    for i, botao in enumerate(botoes, start=1):
        try:
            # Usa JavaScript para clicar no botão.
            # Isso costuma funcionar melhor do que botao.click()
            # em elementos que estão sobrepostos ou com problemas de interação.
            driver.execute_script("arguments[0].click();", botao)

            # Mensagem de sucesso para acompanhamento da execução
            print(f"Critério {i} expandido")

        except Exception as e:
            # Caso ocorra erro ao clicar, o sistema não interrompe o processo inteiro.
            # Apenas registra o problema e continua com os demais botões.
            print(f"Não consegui expandir o critério {i}: {e}")


def copiar_rubrica(html: str):
    """
    Recebe o HTML completo da página e extrai os critérios e níveis selecionados da rubrica.

    A função faz o parsing do HTML com BeautifulSoup e tenta identificar:
    - o texto do critério
    - o nível marcado para cada critério

    Parâmetros:
        html (str): código-fonte HTML da página.

    Retorno:
        list[dict]: lista de dicionários com:
            - criterion_id
            - criterio
            - nivel
    """

    # Cria o objeto BeautifulSoup para navegar e consultar o HTML
    soup = BeautifulSoup(html, "html.parser")

    # Conjunto com os nomes dos níveis conhecidos da rubrica.
    # Isso é usado para evitar confundir o texto de um nível com o texto do critério.
    NIVEIS_CONHECIDOS = {
        "Não satisfatório",
        "Apoiado",
        "Parcialmente Autônomo",
        "Autônomo",
    }

    # Procura todos os blocos que possuem o atributo data-criterion-id.
    # Cada um desses blocos representa um critério da rubrica.
    blocos_criterio = soup.find_all(attrs={"data-criterion-id": True})

    # Dicionário usado para armazenar os critérios sem duplicação.
    # A chave será o criterion_id.
    criterios = {}

    # Percorre todos os blocos identificados como critério
    for bloco in blocos_criterio:

        # Obtém o ID do critério e remove espaços extras
        criterio_id = (bloco.get("data-criterion-id") or "").strip()

        # Se não houver ID, ignora o bloco, pois ele não pode ser identificado corretamente
        if not criterio_id:
            continue

        # Cria a estrutura inicial do critério apenas se ele ainda não existir no dicionário.
        # Isso evita duplicidade quando o mesmo criterion_id aparece mais de uma vez no HTML.
        if criterio_id not in criterios:
            criterios[criterio_id] = {
                "criterion_id": criterio_id,
                "criterio": "Critério não encontrado",
                "nivel": "Nenhum nível marcado"
            }

        # ==========================================================
        # 1) TENTATIVA DE IDENTIFICAR O TEXTO DO CRITÉRIO
        # ==========================================================
        #
        # Aqui buscamos possíveis elementos que contenham o texto da descrição/pergunta
        # do critério. Os seletores foram definidos com base na estrutura observada
        # no HTML do Google Classroom.
        #
        # "span.NPEfkd div.K0lUWd" e "div.K0lUWd" são candidatos a conter o texto desejado.
        candidatos = bloco.select("span.NPEfkd div.K0lUWd, div.K0lUWd")

        # Percorre os candidatos até encontrar um texto que pareça ser o critério
        for tag in candidatos:

            # Extrai o texto do elemento, unindo possíveis quebras com espaço
            texto = tag.get_text(" ", strip=True)

            # Normaliza múltiplos espaços em branco para um único espaço
            texto = " ".join(texto.split()).strip()

            # Regras para considerar esse texto como possível critério:
            # - precisa existir
            # - precisa ter mais de 15 caracteres (heurística para evitar textos curtos irrelevantes)
            # - não pode ser igual a um dos níveis conhecidos
            if (
                texto
                and len(texto) > 15
                and texto not in NIVEIS_CONHECIDOS
            ):
                # Só substitui o valor padrão se ainda não encontrou
                # um critério considerado válido antes
                if criterios[criterio_id]["criterio"] == "Critério não encontrado":
                    criterios[criterio_id]["criterio"] = texto

                # Encerra o loop ao encontrar o primeiro candidato aceitável
                break

        # ==========================================================
        # 2) TENTATIVA DE IDENTIFICAR O NÍVEL MARCADO
        # ==========================================================
        #
        # Cada opção de nível da rubrica costuma aparecer como um elemento
        # com role="menuitemradio". Entre essas opções, procuramos a que está marcada.
        opcoes = bloco.find_all(attrs={"role": "menuitemradio"})

        # Se encontrou opções, mostra informações para depuração
        if opcoes:
            print(f"\n[DEBUG] criterion_id={criterio_id}")
            print(f"[DEBUG] opções encontradas: {len(opcoes)}")

        # Percorre cada opção disponível dentro do critério
        for opcao in opcoes:

            # Verifica se a opção está marcada
            aria_checked = opcao.get("aria-checked")

            # Obtém o rótulo do nível, se existir
            aria_label = (opcao.get("aria-label") or "").strip()

            # Extrai também o texto visível da opção, caso o aria-label não exista
            texto_opcao = opcao.get_text(" ", strip=True)

            # Normaliza os espaços no texto da opção
            texto_opcao = " ".join(texto_opcao.split()).strip()

            # Exibe detalhes de cada opção para facilitar depuração
            print(
                f"aria-checked={aria_checked!r} | "
                f"aria-label={aria_label!r} | "
                f"texto={texto_opcao!r}"
            )

            # Se a opção estiver marcada, salva o nível correspondente
            if aria_checked == "true":
                criterios[criterio_id]["nivel"] = aria_label or texto_opcao or "Nível sem rótulo"
                break

    # Retorna os critérios como lista de dicionários
    return list(criterios.values())


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


def extrair_rubrica():
    """
    Função principal responsável por:
    1. iniciar o navegador
    2. abrir o Google Classroom
    3. aguardar o usuário abrir a rubrica
    4. expandir os critérios
    5. extrair os dados da página
    6. exibir o resultado no terminal
    7. salvar os dados em JSON

    Observação:
        O fechamento do navegador é controlado pelo usuário ao final.
    """

    # Inicia o navegador com o perfil configurado
    driver = iniciar_driver()

    try:
        # Abre a página principal do Google Classroom
        driver.get("https://classroom.google.com/")

        # Pausa para o usuário navegar manualmente até a rubrica desejada
        input("Abra a rubrica e pressione ENTER... ")

        # Aguarda um pequeno tempo extra para a página estabilizar
        time.sleep(2)

        # Expande todos os critérios encontrados
        expandir_criterios(driver)

        # Aguarda a interface atualizar após a expansão
        time.sleep(2)

        # Captura o HTML atual da página
        html = driver.page_source

        # Extrai os dados da rubrica a partir do HTML
        dados = copiar_rubrica(html)

        # Exibe no terminal os critérios e seus respectivos níveis
        print("\nRESULTADO:\n")
        for item in dados:
            print(f"{item['criterio']} => {item['nivel']}")

        # Salva os dados extraídos em arquivo JSON
        salvar_resultados(dados)

    finally:
            driver.quit()


# Este bloco garante que a função extrair_rubrica() só será executada
# automaticamente quando este arquivo for rodado diretamente.
# Se o arquivo for importado em outro módulo (como main.py), esse trecho não roda.
if __name__ == "__main__":
    extrair_rubrica()