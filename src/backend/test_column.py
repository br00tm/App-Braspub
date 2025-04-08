import pandas as pd
import os
import sys
from organizador import processar_planilha

def test_planilha(caminho_arquivo):
    """
    Testa o processamento da planilha e imprime informações úteis para debug.
    """
    print(f"Testando planilha: {caminho_arquivo}")
    
    try:
        # Carregar a planilha diretamente para verificar colunas
        df = pd.read_excel(caminho_arquivo)
        print("\nColunas disponíveis na planilha:")
        for idx, col in enumerate(df.columns):
            print(f"{idx+1}. '{col}'")
        
        # Verificar se as colunas específicas existem
        tipo_midia_coluna = 'Tipo da mídia'
        tipo_midia_maiusculo = 'TIPO DE MÍDIA'
        
        print(f"\nColuna '{tipo_midia_coluna}' existe: {tipo_midia_coluna in df.columns}")
        print(f"Coluna '{tipo_midia_maiusculo}' existe: {tipo_midia_maiusculo in df.columns}")
        
        # Executar a função de processamento
        print("\nExecutando função processar_planilha...")
        resultado = processar_planilha(caminho_arquivo)
        
        print(f"Status: {resultado['status']}")
        if resultado['status'] == 'sucesso':
            tipos = list(resultado['dados'].keys())
            total_itens = sum(len(items) for items in resultado['dados'].values())
            print(f"Tipos encontrados: {len(tipos)}")
            print(f"Total de itens: {total_itens}")
            print(f"Tipos: {', '.join(tipos)}")
        else:
            print(f"Erro: {resultado['mensagem']}")
            
    except Exception as e:
        print(f"Erro ao processar a planilha: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python test_column.py caminho_para_planilha.xlsx")
        sys.exit(1)
        
    arquivo = sys.argv[1]
    if not os.path.exists(arquivo):
        print(f"Arquivo não encontrado: {arquivo}")
        sys.exit(1)
        
    test_planilha(arquivo) 