import pandas as pd
import os
import json
import sys
from datetime import datetime, date, time

def json_serial(obj):
    """
    Função para serializar objetos que não são nativamente serializáveis pelo JSON
    como datetime, date e time.
    """
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    elif isinstance(obj, time):
        return obj.isoformat()
    raise TypeError(f"Tipo não serializável: {type(obj)}")

def processar_planilha(caminho_arquivo):
    """
    Processa a planilha Excel e organiza os dados por tipo de mídia.
    Mantém apenas as colunas:
    - Nome do Cliente (Coluna Assunto - C)
    - Data de Inclusão (Coluna E)
    - Título da Matéria (Coluna L)
    - Link da Matéria Cadastrada (Coluna P)
    - Veículo (Coluna H)
    - Tipo de Mídia (Coluna J)
    
    Args:
        caminho_arquivo: Caminho para o arquivo Excel a ser processado
        
    Returns:
        Um dicionário com os dados organizados por tipo de mídia
    """
    try:
        # Carregar a planilha Excel
        df = pd.read_excel(caminho_arquivo)
        
        # Exibir colunas disponíveis para debug (será removido no código final)
        print("Colunas disponíveis:", df.columns.tolist())
        
        # Mapear colunas originais para novos nomes simplificados
        # Incluindo possíveis variações de nomes de colunas
        mapeamento_colunas = {
            'Assunto': 'Nome do Cliente',
            'Data de inclusão': 'Data de Inclusão',
            'Título': 'Título da Matéria',
            'Link compartilhável': 'Link da Matéria',
            'Link': 'Link da Matéria',
            'Link da Matéria': 'Link da Matéria',
            'Veículo': 'Veículo',
            'Tipo da mídia': 'Tipo de Mídia'
        }
        
        # Se estamos lidando com índices de colunas em vez de nomes
        # Usar iloc para acessar as colunas por índice se necessário
        colunas_por_indice = {}
        try:
            # Coluna C (2) - Nome do Cliente (Assunto)
            colunas_por_indice[df.columns[2]] = 'Nome do Cliente'
            # Coluna E (4) - Data de Inclusão
            colunas_por_indice[df.columns[4]] = 'Data de Inclusão'
            # Coluna L (11) - Título da Matéria
            colunas_por_indice[df.columns[11]] = 'Título da Matéria'
            # Coluna P (15) - Link da Matéria
            colunas_por_indice[df.columns[15]] = 'Link da Matéria'
            # Coluna H (7) - Veículo
            colunas_por_indice[df.columns[7]] = 'Veículo'
            # Coluna J (9) - Tipo de Mídia
            colunas_por_indice[df.columns[9]] = 'Tipo de Mídia'
        except IndexError:
            # Algumas colunas podem não existir, o que é aceitável
            pass
        
        # Adicionar as colunas por índice ao mapeamento
        mapeamento_colunas.update(colunas_por_indice)
        
        # Verificar se a coluna 'Tipo da mídia' existe, por nome ou índice
        tipo_midia_encontrado = False
        if 'Tipo da mídia' in df.columns:
            tipo_midia_encontrado = True
        elif len(df.columns) > 9:  # Verificar coluna J (índice 9)
            tipo_midia_encontrado = True
            df['Tipo da mídia'] = df.iloc[:, 9]
            
        if not tipo_midia_encontrado:
            return {'status': 'erro', 'mensagem': 'Formato de planilha inválido. Coluna de Tipo da mídia não encontrada.'}
        
        # Limpar dados
        df = df.fillna('')
        
        # Converter campos datetime, date e time para string
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%Y-%m-%d')
            elif df[col].dtype == 'object':
                # Tentativa de identificar e converter colunas com objetos time
                sample_vals = df[col].dropna().head()
                if len(sample_vals) > 0 and any(isinstance(val, (datetime, date, time)) for val in sample_vals):
                    df[col] = df[col].apply(lambda x: x.isoformat() if isinstance(x, (datetime, date, time)) else x)
        
        # Criar um novo DataFrame com as colunas mapeadas
        novo_df = pd.DataFrame()
        
        # Mapear as colunas por nome
        for col_original, col_nova in mapeamento_colunas.items():
            if col_original in df.columns:
                novo_df[col_nova] = df[col_original]
        
        # Adicionar colunas por índice se necessário
        # Coluna C (2) - Nome do Cliente
        if 'Nome do Cliente' not in novo_df.columns and len(df.columns) > 2:
            novo_df['Nome do Cliente'] = df.iloc[:, 2]
            
        # Coluna E (4) - Data de Inclusão
        if 'Data de Inclusão' not in novo_df.columns and len(df.columns) > 4:
            novo_df['Data de Inclusão'] = df.iloc[:, 4]
            
        # Coluna L (11) - Título da Matéria
        if 'Título da Matéria' not in novo_df.columns and len(df.columns) > 11:
            novo_df['Título da Matéria'] = df.iloc[:, 11]
            
        # Coluna P (15) - Link da Matéria
        if 'Link da Matéria' not in novo_df.columns and len(df.columns) > 15:
            novo_df['Link da Matéria'] = df.iloc[:, 15]
            
        # Coluna H (7) - Veículo
        if 'Veículo' not in novo_df.columns and len(df.columns) > 7:
            novo_df['Veículo'] = df.iloc[:, 7]
            
        # Coluna J (9) - Tipo de Mídia (se não foi adicionada anteriormente)
        if 'Tipo de Mídia' not in novo_df.columns and len(df.columns) > 9:
            novo_df['Tipo de Mídia'] = df.iloc[:, 9]
        elif 'Tipo de Mídia' not in novo_df.columns and 'Tipo da mídia' in df.columns:
            novo_df['Tipo de Mídia'] = df['Tipo da mídia']
        
        # Organizar os dados por tipo de mídia
        resultado = {}
        
        # Agrupar por tipo de mídia
        if 'Tipo de Mídia' in novo_df.columns:
            tipos_midia = novo_df['Tipo de Mídia'].unique()
        else:
            # Como não temos a coluna "Tipo de Mídia", usamos a original
            tipos_midia = df['Tipo da mídia'].unique() if 'Tipo da mídia' in df.columns else ['Sem classificação']
            
        for tipo in tipos_midia:
            if not tipo or str(tipo).strip() == '':
                continue
                
            # Filtrar o dataframe pelo tipo de mídia
            if 'Tipo de Mídia' in novo_df.columns:
                df_tipo = novo_df[novo_df['Tipo de Mídia'] == tipo].copy()
            elif 'Tipo da mídia' in df.columns:
                # Filtramos no DataFrame original
                indices = df[df['Tipo da mídia'] == tipo].index
                df_tipo = novo_df.loc[indices].copy()
            else:
                # Se não temos como filtrar, usamos todo o DataFrame
                df_tipo = novo_df.copy()
            
            # Converter para lista de dicionários
            registros = df_tipo.to_dict(orient='records')
            
            # Adicionar ao resultado
            resultado[str(tipo)] = registros
        
        return {'status': 'sucesso', 'dados': resultado}
        
    except Exception as e:
        import traceback
        traceback_str = traceback.format_exc()
        return {'status': 'erro', 'mensagem': str(e), 'traceback': traceback_str}

def exportar_planilha(dados, caminho_saida):
    """
    Exporta os dados processados para uma planilha Excel com abas separadas por tipo de mídia.
    Mantém apenas as colunas simplificadas e as ordena:
    - Nome do Cliente
    - Data de Inclusão
    - Título da Matéria
    - Link da Matéria
    - Veículo
    - Tipo de Mídia
    
    Args:
        dados: Dicionário com dados organizados por tipo de mídia
        caminho_saida: Caminho onde a planilha será salva
        
    Returns:
        Status da operação
    """
    try:
        # Definir a ordem desejada das colunas
        ordem_colunas = [
            'Nome do Cliente',
            'Data de Inclusão',
            'Título da Matéria',
            'Link da Matéria',
            'Veículo',
            'Tipo de Mídia'
        ]
        
        # Criar um objeto ExcelWriter
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            # Para cada tipo de mídia, criar uma aba
            for tipo, registros in dados.items():
                # Converter a lista de dicionários de volta para DataFrame
                df = pd.DataFrame(registros)
                
                # Organizar as colunas na ordem desejada (apenas as que existem)
                colunas_finais = [col for col in ordem_colunas if col in df.columns]
                if colunas_finais:
                    df = df[colunas_finais]
                
                # Substituir o nome da aba por algo válido
                nome_aba = str(tipo)[:31]  # Excel limita o nome da aba a 31 caracteres
                nome_aba = nome_aba.replace('/', '_').replace('\\', '_').replace('?', '').replace('*', '')
                nome_aba = nome_aba.replace('[', '').replace(']', '').replace(':', '')
                
                # Escrever na aba
                df.to_excel(writer, sheet_name=nome_aba, index=False)
        
        return {'status': 'sucesso', 'mensagem': f'Planilha salva com sucesso em {caminho_saida}'}
        
    except Exception as e:
        return {'status': 'erro', 'mensagem': str(e)}

def main():
    # Verificar argumentos
    if len(sys.argv) < 2:
        print(json.dumps({'status': 'erro', 'mensagem': 'Argumentos insuficientes'}))
        return
        
    # Primeiro argumento: caminho do arquivo a processar
    caminho_arquivo = sys.argv[1]
    
    # Se o arquivo de entrada for um JSON, é para exportação
    if caminho_arquivo.endswith('.json'):
        # Verificar se o segundo argumento (caminho de saída) foi fornecido
        if len(sys.argv) < 3:
            print(json.dumps({'status': 'erro', 'mensagem': 'Caminho de saída não fornecido'}))
            return
            
        # Caminho de saída
        caminho_saida = sys.argv[2]
        
        # Ler os dados do JSON
        try:
            with open(caminho_arquivo, 'r') as f:
                dados = json.load(f)
            
            # Exportar para Excel
            resultado = exportar_planilha(dados, caminho_saida)
            print(json.dumps(resultado))
            
        except Exception as e:
            print(json.dumps({'status': 'erro', 'mensagem': str(e)}))
            
    else:
        # Processar a planilha
        resultado = processar_planilha(caminho_arquivo)
        print(json.dumps(resultado, default=json_serial))

if __name__ == "__main__":
    main() 