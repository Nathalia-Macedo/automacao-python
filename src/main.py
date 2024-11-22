from excel_reader import ler_dados_excel, imprimir_dados

def main():
    print("Lendo dados do formulário de origem...")
    dados = ler_dados_excel()
    print("\nDados lidos:")
    imprimir_dados(dados)

if __name__ == "__main__":
    main()