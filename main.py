from extrair_rubrica import extrair_rubrica
from preencher_planilha import preencher_planilha_excel_aberta


def main():
    print("=== ETAPA 1: EXTRAÇÃO DA RUBRICA ===")
    extrair_rubrica()

    print("\n=== ETAPA 2: PREENCHIMENTO DA PLANILHA ===")
    print("Agora abra a planilha no Excel Desktop e deixe a aba correta ativa.")
    input("Quando estiver tudo pronto, pressione ENTER para continuar... ")

    preencher_planilha_excel_aberta()

    print("\nProcesso concluído com sucesso.")


if __name__ == "__main__":
    main()