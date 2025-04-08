from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import json
import uuid
from datetime import datetime, date, time
import pandas as pd
import random
import logging
from urllib.parse import urlparse, urlunparse
import requests
import shutil
from pathlib import Path
import re
import time
from bs4 import BeautifulSoup
from urllib.parse import unquote

app = Flask(__name__)
CORS(app)  # Habilitar CORS para todas as rotas

# Configuração de logging
LOG_DIR = 'logs'
os.makedirs(LOG_DIR, exist_ok=True)
log_file = os.path.join(LOG_DIR, f'braspub_api_{datetime.now().strftime("%Y%m%d")}.log')

# Configurar o logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()  # Para também exibir logs no console
    ]
)
logger = logging.getLogger('braspub_api')
logger.info(f"Iniciando API - Log configurado em {log_file}")

# Diretório para arquivos temporários
TEMP_DIR = os.path.join(tempfile.gettempdir(), 'organizador_planilhas')
os.makedirs(TEMP_DIR, exist_ok=True)
logger.info(f"Diretório temporário configurado: {TEMP_DIR}")

# Função auxiliar para serializar objetos complexos para JSON
def serializar_para_json(dados):
    def conversor_personalizado(obj):
        if isinstance(obj, (datetime, date, time)):
            return obj.isoformat()
        elif hasattr(obj, 'to_dict'):
            return obj.to_dict()
        elif isinstance(obj, set):
            return list(obj)
        elif not isinstance(obj, (str, int, float, bool, list, dict, type(None))):
            try:
                return dict(obj)
            except (TypeError, ValueError):
                return str(obj)
        return f"[Objeto não serializável: {type(obj).__name__}]"
    
    try:
        return json.dumps(dados, default=conversor_personalizado, ensure_ascii=False)
    except Exception as e:
        print(f"Erro na serialização JSON: {str(e)}")
        return json.dumps({'status': 'erro', 'mensagem': f'Erro na serialização: {str(e)}'})

@app.route('/api/processar', methods=['POST'])
def api_processar():
    """Recebe um arquivo Excel, processa e retorna os dados organizados."""
    try:
        # Verificar se existe um arquivo válido na requisição
        if 'arquivo' not in request.files:
            return jsonify({'status': 'erro', 'mensagem': 'Nenhum arquivo enviado'}), 400
            
        arquivo = request.files['arquivo']
        if arquivo.filename == '':
            return jsonify({'status': 'erro', 'mensagem': 'Nome de arquivo vazio'}), 400
        if not arquivo.filename.endswith(('.xls', '.xlsx')):
            return jsonify({'status': 'erro', 'mensagem': 'Formato de arquivo inválido. Use .xls ou .xlsx'}), 400
        
        # Salvar o arquivo temporariamente
        temp_path = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}_{arquivo.filename}")
        arquivo.save(temp_path)
        print(f"Arquivo salvo em: {temp_path}")
        
        try:
            # Verificar se o arquivo existe e pode ser lido
            if not os.path.exists(temp_path):
                return jsonify({'status': 'erro', 'mensagem': 'Arquivo não encontrado após upload'}), 400
            
            # Processar o arquivo
            df = pd.read_excel(temp_path)
            print(f"Arquivo lido: {df.shape[0]} linhas, {df.shape[1]} colunas")
            
            # Identificar colunas importantes
            colunas = {
                'url': next((col for col in df.columns if 'url' in str(col).lower() or ('link' in str(col).lower() and 'web' not in str(col).lower())), df.columns[0] if len(df.columns) > 0 else None),
                'link_web_texto': next((col for col in df.columns if 'link web' in str(col).lower() and 'texto' in str(col).lower()), None),
                'link_web_imagem': next((col for col in df.columns if ('link web' in str(col).lower() and 'imagem' in str(col).lower()) or 'link_web_imagem' in str(col).lower()), None),
                'tipo_midia': next((col for col in df.columns if 'tipo' in str(col).lower() and 'midia' in str(col).lower()), None)
            }
            
            # Processar linhas e organizar por tipo de mídia
            tipos_midia = ['Portal', 'Impresso', 'TV', 'Rádio']
            resultado = {tipo: [] for tipo in tipos_midia}
            
            for idx, row in df.iterrows():
                if idx >= 50:  # Limitar a 50 linhas
                    break
                
                # Obter valores da linha
                url_base = str(row[colunas['url']]).strip() if colunas['url'] and pd.notna(row[colunas['url']]) else ""
                link_web_texto = str(row[colunas['link_web_texto']]).strip() if colunas['link_web_texto'] and pd.notna(row[colunas['link_web_texto']]) else ""
                link_web_imagem = str(row[colunas['link_web_imagem']]).strip() if colunas['link_web_imagem'] and pd.notna(row[colunas['link_web_imagem']]) else ""
                tipo_midia = str(row[colunas['tipo_midia']]).strip() if colunas['tipo_midia'] and pd.notna(row[colunas['tipo_midia']]) else "Portal"
                
                # Determinar Link da Matéria baseado na prioridade correta
                link_materia = "Materia Não Cadastrada"  # Valor padrão quando não há links
                
                # Para qualquer tipo, priorizar link_web_texto
                if link_web_texto and link_web_texto.startswith(('http://', 'https://')):
                    link_materia = link_web_texto
                # Para Impresso, se não há link_web_texto, usar link_web_imagem
                elif tipo_midia == 'Impresso' and link_web_imagem and link_web_imagem.startswith(('http://', 'https://')):
                    link_materia = link_web_imagem
                # Caso contrário, mantém "Materia Não Cadastrada"
                
                # Criar registro base
                registro_base = {
                    'Nome do Cliente': 'Cliente',
                    'Data de Inclusão': datetime.now().strftime('%Y-%m-%d'),
                    'Título da Matéria': f"Matéria {idx+1}",
                    'Link da Matéria': link_materia,
                    'Veículo': 'Veículo padrão',
                    'Tipo de Mídia': tipo_midia
                }
                
                # Adicionar ao resultado
                if tipo_midia in tipos_midia:
                    resultado[tipo_midia].append(registro_base)
                else:
                    resultado['Portal'].append(registro_base)
            
            # Salvar resultado em planilha processada
            output_path = f"{os.path.splitext(temp_path)[0]}_processado.xlsx"
            with pd.ExcelWriter(output_path) as writer:
                for tipo, registros in resultado.items():
                    if registros:
                        df_tipo = pd.DataFrame(registros)
                        df_tipo.to_excel(writer, sheet_name=tipo, index=False)
            
            return jsonify({
                'status': 'sucesso',
                'mensagem': 'Arquivo processado com sucesso',
                'dados': resultado
            })
            
        finally:
            # Remover arquivo temporário
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
    except Exception as e:
        print(f"Erro ao processar: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

@app.route('/api/processar_keywords', methods=['POST'])
def api_processar_keywords():
    """Recebe um arquivo Excel de palavras-chave e organiza por palavras-chave e tipo de mídia."""
    try:
        # Verificar se existe um arquivo válido na requisição
        if 'arquivo' not in request.files:
            return jsonify({'status': 'erro', 'mensagem': 'Nenhum arquivo enviado'}), 400
            
        arquivo = request.files['arquivo']
        if arquivo.filename == '':
            return jsonify({'status': 'erro', 'mensagem': 'Nome de arquivo vazio'}), 400
        if not arquivo.filename.endswith(('.xls', '.xlsx')):
            return jsonify({'status': 'erro', 'mensagem': 'Formato de arquivo inválido. Use .xls ou .xlsx'}), 400
        
        # Salvar o arquivo temporariamente
        temp_path = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}_{arquivo.filename}")
        arquivo.save(temp_path)
        print(f"Arquivo salvo em: {temp_path}")
        
        try:
            # Verificar arquivo
            if not os.path.exists(temp_path):
                return jsonify({'status': 'erro', 'mensagem': 'Arquivo não encontrado após upload'}), 400
            
            # Processar o arquivo
            df = pd.read_excel(temp_path)
            print(f"Arquivo lido: {df.shape[0]} linhas, {df.shape[1]} colunas")
            
            # Identificar coluna de palavras-chave
            palavras_chave_col = None
            for col in df.columns:
                col_lower = str(col).lower()
                if any(termo in col_lower for termo in ['palavra', 'chave', 'keyword']):
                    palavras_chave_col = col
                    logger.info(f"Coluna de palavras-chave encontrada: {col}")
                    break
            
            # Se não encontrou, usar a primeira coluna
            if not palavras_chave_col and len(df.columns) > 0:
                palavras_chave_col = df.columns[0]
                print(f"Usando primeira coluna como palavras-chave: {palavras_chave_col}")
            
            # Identificar outras colunas importantes com verificação mais precisa
            colunas = {
                'link': None,
                'link_web_texto': None,
                'link_web_imagem': None,
                'tipo_midia': None
            }
            
            # Buscar todas as colunas por nome mais específico
            for col in df.columns:
                col_lower = str(col).lower()
                
                # Para link_web_texto
                if ('link web' in col_lower and 'texto' in col_lower) or 'link_web_texto' in col_lower:
                    colunas['link_web_texto'] = col
                    print(f"Coluna Link web - Texto encontrada: {col}")
                
                # Para link_web_imagem
                elif ('link web' in col_lower and 'imagem' in col_lower) or 'link_web_imagem' in col_lower:
                    colunas['link_web_imagem'] = col
                    print(f"Coluna Link web - Imagem encontrada: {col}")
                
                # Para tipo_midia
                elif 'tipo' in col_lower and 'midia' in col_lower:
                    colunas['tipo_midia'] = col
                    print(f"Coluna Tipo de Mídia encontrada: {col}")
                
                # Para link genérico (apenas se não conflitar com as colunas específicas)
                elif 'link' in col_lower and 'web' not in col_lower and 'texto' not in col_lower and 'imagem' not in col_lower:
                    colunas['link'] = col
                    print(f"Coluna Link genérico encontrada: {col}")
            
            print(f"Colunas identificadas: {colunas}")
            
            # Extrair palavras-chave únicas
            palavras_unicas = set()
            for idx, row in df.iterrows():
                if pd.notna(row[palavras_chave_col]):
                    palavras_cell = str(row[palavras_chave_col]).strip()
                    # Dividir a célula por vírgulas e adicionar cada parte como uma palavra-chave
                    for palavra in palavras_cell.split(','):
                        palavra_limpa = palavra.strip()
                        if palavra_limpa:  # Verificar se não está vazia após a limpeza
                            palavras_unicas.add(palavra_limpa)
            
            # Converter para lista sem limitar a quantidade
            palavras_unicas = list(palavras_unicas)
            print(f"Palavras-chave encontradas: {len(palavras_unicas)}")
            print(f"Palavras: {', '.join(palavras_unicas[:5] if len(palavras_unicas) >= 5 else palavras_unicas)}")
            
            # Extrair links e tipos de mídia para cada palavra
            info_palavras = {}
            for palavra in palavras_unicas:
                # Buscar linhas que contenham a palavra-chave (não apenas iguais)
                linhas_palavra = df[df[palavras_chave_col].astype(str).str.contains(palavra, case=False, regex=False)]
                print(f"Palavra '{palavra}': {len(linhas_palavra)} linhas encontradas")
                
                info_palavras[palavra] = {
                    'links_por_linha': [],  # Armazenar pares (link_texto, link_imagem) por linha
                    'tipos_midia': []
                }
                
                for idx, row in linhas_palavra.iterrows():
                    # Extrair links da linha atual
                    link_texto = ""
                    link_imagem = ""
                    
                    # Obter link_web_texto
                    if colunas['link_web_texto'] and pd.notna(row[colunas['link_web_texto']]):
                        link_texto = str(row[colunas['link_web_texto']]).strip()
                        if not link_texto.startswith(('http://', 'https://')):
                            link_texto = ""
                    
                    # Obter link_web_imagem
                    if colunas['link_web_imagem'] and pd.notna(row[colunas['link_web_imagem']]):
                        link_imagem = str(row[colunas['link_web_imagem']]).strip()
                        if not link_imagem.startswith(('http://', 'https://')):
                            link_imagem = ""
                    
                    # Extrair tipo_midia se disponível
                    if colunas['tipo_midia'] and pd.notna(row[colunas['tipo_midia']]):
                        tipo = str(row[colunas['tipo_midia']]).strip()
                        if tipo:
                            info_palavras[palavra]['tipos_midia'].append(tipo)
                    
                    # Armazenar o par de links desta linha
                    if link_texto or link_imagem:
                        info_palavras[palavra]['links_por_linha'].append((link_texto, link_imagem))
                
                print(f"Informações para '{palavra}':")
                print(f"  - Pares de links: {len(info_palavras[palavra]['links_por_linha'])}")
                print(f"  - Tipos de Mídia: {len(info_palavras[palavra]['tipos_midia'])}")
            
            # Preparar dados para Excel
            # Tipos de mídia padrão (alterando de 'Rádio' para 'Online')
            tipos_midia = ['Portal', 'Impresso', 'TV', 'Online']
            registros = []
            
            # Para cada palavra-chave, criar exatamente 4 registros (um para cada tipo de mídia)
            for palavra in palavras_unicas:
                # Selecionar o melhor par de links para esta palavra-chave
                melhor_link_texto = ""
                melhor_link_imagem = ""
                
                # Se há algum par de links disponível, use o primeiro
                if info_palavras[palavra]['links_por_linha']:
                    melhor_link_texto, melhor_link_imagem = info_palavras[palavra]['links_por_linha'][0]
                
                # Determinar compatibilidade de tipos de mídia com base nas extensões
                tipos_compatíveis = {}
                
                if melhor_link_imagem:
                    link_lower = melhor_link_imagem.lower()
                    # Verificar extensões para determinar compatibilidade
                    if link_lower.endswith('.mp4'):
                        tipos_compatíveis['Online'] = True
                    elif link_lower.endswith('.mp3'):
                        tipos_compatíveis['TV'] = True
                    elif any(link_lower.endswith(ext) for ext in ['.jpg', '.jpeg', '.png']):
                        tipos_compatíveis['Portal'] = True
                        tipos_compatíveis['Impresso'] = True
                
                # Criar exatamente um registro para cada tipo de mídia
                for tipo in tipos_midia:
                    # Determinar o LINK DA MATÉRIA CADASTRADA
                    if tipo in tipos_compatíveis and melhor_link_texto and melhor_link_texto.lower().startswith(('http://', 'https://')):
                        link_materia = melhor_link_texto
                        logger.info(f"[Registro] '{palavra}' tipo '{tipo}': Usando LINK WEB - TEXTO (compatível)")
                    else:
                        link_materia = "Materia Não Cadastrada"
                        if tipo not in tipos_compatíveis:
                            logger.info(f"[Registro] '{palavra}' tipo '{tipo}': Tipo não compatível com os links disponíveis")
                        else:
                            logger.info(f"[Registro] '{palavra}' tipo '{tipo}': Link WEB - TEXTO inválido ou ausente")
                    
                    # Criar registro
                    registro = {
                        'PALAVRAS-CHAVE': palavra,
                        'DATA DE INCLUSÃO': datetime.now().strftime('%Y-%m-%d'),
                        'TÍTULO DA MATÉRIA': f"Matéria sobre {palavra}",
                        'TIPO DE MÍDIA': tipo,
                        'LINK DA MATÉRIA CADASTRADA': link_materia
                    }
                    registros.append(registro)
            
            # Converter para DataFrame
            df_resultado = pd.DataFrame(registros)
            logger.info(f"Registros gerados: {len(registros)}")
            
            # Salvar em Excel
            output_path = f"{os.path.splitext(temp_path)[0]}_keywords.xlsx"
            df_resultado.to_excel(output_path, index=False)
            logger.info(f"Arquivo Excel salvo em: {output_path}")
            
            # Organizar registros por palavra-chave para exportação posterior
            dados_por_palavra = {}
            for registro in registros:
                palavra = registro['PALAVRAS-CHAVE']
                if palavra not in dados_por_palavra:
                    dados_por_palavra[palavra] = []
                dados_por_palavra[palavra].append(registro)
            
            return jsonify({
                'status': 'sucesso',
                'mensagem': 'Arquivo de palavras-chave processado com sucesso',
                'total_registros': len(registros),
                'palavras_processadas': len(palavras_unicas),
                'output_path': output_path,
                'dados': dados_por_palavra  # Adiciona dados organizados por palavra-chave
            })
            
        finally:
            # Remover arquivo temporário
            if os.path.exists(temp_path):
                os.remove(temp_path)
        
    except Exception as e:
        logger.error(f"Erro ao processar keywords: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

@app.route('/api/exportar', methods=['POST'])
def api_exportar():
    """Recebe dados processados e retorna um arquivo Excel."""
    if not request.json or 'dados' not in request.json:
        return jsonify({'status': 'erro', 'mensagem': 'Dados não fornecidos'}), 400
    
    try:
        dados = request.json['dados']
        temp_xlsx = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}.xlsx")
        
        # Verificando e registrando o formato dos dados para debug
        logger.info(f"Estrutura dos dados: {type(dados)}")
        for tipo, registros in dados.items():
            logger.info(f"Tipo de mídia: {tipo}, número de registros: {len(registros) if isinstance(registros, list) else 'não é lista'}")
            
            # Converter para lista se não for
            if not isinstance(registros, list):
                logger.warning(f"Convertendo registros para {tipo} que não é lista...")
                # Se for dicionário, tentar transformar em lista
                if isinstance(registros, dict):
                    dados[tipo] = [registros]
                else:
                    # Se não for dict nem list, criar um registro vazio
                    logger.error(f"Registros para {tipo} não é um formato válido")
                    dados[tipo] = []
        
        # Criar a planilha Excel
        with pd.ExcelWriter(temp_xlsx) as writer:
            abas_escritas = False
            
            for tipo, registros in dados.items():
                if registros and isinstance(registros, list):
                    # Convertendo para DataFrame
                    df_tipo = pd.DataFrame(registros)
                    
                    # Verificando e registrando as colunas para debug
                    logger.info(f"Colunas no DataFrame para {tipo}: {df_tipo.columns.tolist()}")
                    
                    # Limitar o tamanho do nome da aba para evitar erros do Excel
                    nome_aba = str(tipo)[:31]  # Excel limita o nome da aba a 31 caracteres
                    nome_aba = nome_aba.replace('/', '_').replace('\\', '_').replace('?', '').replace('*', '')
                    nome_aba = nome_aba.replace('[', '').replace(']', '').replace(':', '')
                    
                    # Salvar os dados na aba
                    df_tipo.to_excel(writer, sheet_name=nome_aba, index=False)
                    abas_escritas = True
                    
                    # Ajustar largura das colunas
                    worksheet = writer.sheets[nome_aba]
                    for idx, col in enumerate(df_tipo.columns):
                        max_length = max(
                            df_tipo[col].astype(str).map(len).max(),
                            len(str(col))
                        ) + 2
                        col_letter = chr(65 + idx) if idx < 26 else chr(64 + idx // 26) + chr(65 + idx % 26)
                        worksheet.column_dimensions[col_letter].width = min(max_length, 100)
            
            # Se nenhuma aba foi escrita, criar uma aba vazia para evitar erro
            if not abas_escritas:
                logger.warning("Nenhuma aba foi escrita. Criando aba vazia.")
                pd.DataFrame(columns=['Aviso']).to_excel(writer, sheet_name='Sem Dados', index=False)
        
        # Verificar se o arquivo foi criado
        if not os.path.exists(temp_xlsx):
            logger.error("Arquivo Excel não foi criado")
            return jsonify({'status': 'erro', 'mensagem': 'Erro ao criar arquivo Excel'}), 500
        
        logger.info(f"Arquivo Excel criado com sucesso: {temp_xlsx}")
        
        # Enviar o arquivo
        try:
            response = send_file(
                temp_xlsx,
                as_attachment=True,
                download_name="Planilha_Organizada.xlsx",
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
            # Remover o arquivo após envio
            @response.call_on_close
            def remove_temp_file():
                if os.path.exists(temp_xlsx):
                    os.remove(temp_xlsx)
                    logger.info(f"Arquivo temporário removido: {temp_xlsx}")
            
            return response
            
        except Exception as e:
            logger.error(f"Erro ao enviar arquivo: {str(e)}")
            # Se falhar ao enviar o arquivo, tente retornar erro em JSON
            return jsonify({'status': 'erro', 'mensagem': f'Erro ao enviar arquivo: {str(e)}'}), 500
        
    except Exception as e:
        logger.error(f"Erro ao exportar: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

@app.route('/api/exportar_keywords', methods=['POST'])
def api_exportar_keywords():
    """Recebe dados de palavras-chave processados e retorna um arquivo Excel."""
    if not request.json or 'dados' not in request.json:
        return jsonify({'status': 'erro', 'mensagem': 'Dados não fornecidos'}), 400
    
    try:
        dados = request.json['dados']
        temp_xlsx = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}.xlsx")
        
        # Criar DataFrame com todas as palavras-chave
        colunas_necessarias = [
            'PALAVRAS-CHAVE', 'DATA DE CADASTRO', 'TÍTULO DA MATÉRIA',
            'TIPO DE MÍDIA', 'LINK DA MATÉRIA CADASTRADA'
        ]
        df_final = pd.DataFrame(columns=colunas_necessarias)
        
        # Verificar e logar os dados recebidos
        print(f"Dados recebidos para exportação: {len(dados)} palavras-chave")
        
        # Adicionar todas as palavras-chave ao DataFrame
        registros_totais = []
        for palavra, registros in dados.items():
            print(f"Palavra '{palavra}': {len(registros)} registros")
            # Log das colunas no primeiro registro
            if registros and len(registros) > 0:
                print(f"Colunas disponíveis: {list(registros[0].keys())}")
            
            for reg in registros:
                # Garantir que todos os campos necessários existam
                registro_completo = {col: "" for col in colunas_necessarias}
                
                # Converter DATA DE INCLUSÃO para DATA DE CADASTRO se necessário
                if 'DATA DE INCLUSÃO' in reg and reg['DATA DE INCLUSÃO'] and not reg['DATA DE CADASTRO']:
                    reg['DATA DE CADASTRO'] = reg['DATA DE INCLUSÃO']
                
                registro_completo.update(reg)
                registros_totais.append(registro_completo)
        
        print(f"Total de registros a exportar: {len(registros_totais)}")
        
        # Criar DataFrame apenas se houver registros
        if registros_totais:
            df_final = pd.DataFrame(registros_totais)
            print(f"DataFrame criado com {len(df_final)} linhas e {len(df_final.columns)} colunas")
            print(f"Colunas no DataFrame: {df_final.columns.tolist()}")
        else:
            print("AVISO: Nenhum registro encontrado para exportar!")
        
        # Salvar para Excel
        with pd.ExcelWriter(temp_xlsx) as writer:
            df_final.to_excel(writer, sheet_name='Palavras-Chave', index=False)
            
            # Ajustar largura das colunas
            worksheet = writer.sheets['Palavras-Chave']
            for idx, col in enumerate(df_final.columns):
                max_length = max(
                    df_final[col].astype(str).map(len).max(),
                    len(str(col))
                ) + 2
                col_letter = chr(65 + idx) if idx < 26 else chr(64 + idx // 26) + chr(65 + idx % 26)
                worksheet.column_dimensions[col_letter].width = min(max_length, 100)
        
        # Verificar se o arquivo foi criado
        if not os.path.exists(temp_xlsx):
            return jsonify({'status': 'erro', 'mensagem': 'Erro ao criar arquivo Excel'}), 500
        
        # Enviar o arquivo
        response = send_file(
            temp_xlsx,
            as_attachment=True,
            download_name="Palavras_Chave_Organizadas.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        # Remover o arquivo após envio
        @response.call_on_close
        def remove_temp_file():
            if os.path.exists(temp_xlsx):
                os.remove(temp_xlsx)
                
        return response
        
    except Exception as e:
        print(f"Erro ao exportar: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

@app.route('/api/status', methods=['GET'])
def api_status():
    """Verifica se a API está funcionando."""
    return jsonify({'status': 'online', 'mensagem': 'API do Organizador de Planilhas está funcionando'})

@app.route('/api/processar-planilha-download', methods=['POST'])
def api_processar_planilha_download():
    """Recebe um arquivo Excel, extrai links para download e retorna os dados organizados."""
    try:
        # Verificar se existe um arquivo válido na requisição
        if 'arquivo' not in request.files:
            return jsonify({'status': 'erro', 'mensagem': 'Nenhum arquivo enviado'}), 400
            
        arquivo = request.files['arquivo']
        if arquivo.filename == '':
            return jsonify({'status': 'erro', 'mensagem': 'Nome de arquivo vazio'}), 400
        if not arquivo.filename.endswith(('.xls', '.xlsx')):
            return jsonify({'status': 'erro', 'mensagem': 'Formato de arquivo inválido. Use .xls ou .xlsx'}), 400
        
        # Salvar o arquivo temporariamente
        temp_path = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}_{arquivo.filename}")
        arquivo.save(temp_path)
        logger.info(f"Arquivo salvo em: {temp_path}")
        
        try:
            # Verificar arquivo
            if not os.path.exists(temp_path):
                return jsonify({'status': 'erro', 'mensagem': 'Arquivo não encontrado após upload'}), 400
            
            # Processar o arquivo para extrair links
            resultado = processar_planilha_download(temp_path)
            
            return jsonify({
                'status': 'sucesso',
                'mensagem': 'Links extraídos com sucesso',
                'dados': resultado
            })
            
        finally:
            # Remover arquivo temporário
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
    except Exception as e:
        logger.error(f"Erro ao processar planilha para download: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

@app.route('/api/processar_planilha_download', methods=['POST'])
def api_processar_planilha_download_compat():
    """Versão compatível da rota para o frontend."""
    return api_processar_planilha_download()

@app.route('/api/baixar-arquivos', methods=['POST'])
def api_baixar_arquivos():
    """Recebe dados com links, baixa os arquivos e organiza em pastas."""
    try:
        # Verificar se há dados
        if not request.json:
            return jsonify({'status': 'erro', 'mensagem': 'Nenhum dado fornecido'}), 400
        
        dados = request.json
        
        # Baixar arquivos
        resultado = baixar_arquivos(dados)
        
        return jsonify({
            'status': 'sucesso',
            'mensagem': 'Arquivos baixados com sucesso',
            'detalhes': resultado
        })
            
    except Exception as e:
        logger.error(f"Erro ao baixar arquivos: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

@app.route('/api/baixar_arquivos', methods=['POST'])
def api_baixar_arquivos_compat():
    """Versão compatível da rota para o frontend."""
    return api_baixar_arquivos()

# Função para processar a planilha e extrair links para download
def processar_planilha_download(arquivo_path):
    """
    Processa a planilha para extrair links para download.
    
    Args:
        arquivo_path: Caminho para o arquivo Excel
        
    Returns:
        Um dicionário com os links organizados por data e tipo de mídia
    """
    logger.info(f"Processando planilha para download: {arquivo_path}")
    
    # Ler arquivo Excel
    df = pd.read_excel(arquivo_path)
    logger.info(f"Planilha lida: {df.shape[0]} linhas, {df.shape[1]} colunas")
    
    # Identificar colunas importantes
    colunas = {}
    for col in df.columns:
        col_lower = str(col).lower()
        # Coluna de título
        if any(termo in col_lower for termo in ['título', 'titulo', 'title']):
            colunas['titulo'] = col
            logger.info(f"Coluna de título encontrada: {col}")
        
        # Coluna de data
        elif any(termo in col_lower for termo in ['data', 'date', 'inclusão', 'inclusao']):
            colunas['data'] = col
            logger.info(f"Coluna de data encontrada: {col}")
        
        # Coluna de tipo de mídia
        elif 'tipo' in col_lower and 'mídia' in col_lower:
            colunas['tipo_midia'] = col
            logger.info(f"Coluna de tipo de mídia encontrada: {col}")
        
        # Coluna de link web imagem
        elif ('link web' in col_lower and 'imagem' in col_lower) or 'link_web_imagem' in col_lower:
            colunas['link_web_imagem'] = col
            logger.info(f"Coluna de link web imagem encontrada: {col}")
    
    # Verificar se encontrou as colunas necessárias
    colunas_obrigatorias = ['titulo', 'data', 'tipo_midia', 'link_web_imagem']
    colunas_faltando = [col for col in colunas_obrigatorias if col not in colunas]
    
    if colunas_faltando:
        mensagem = f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}"
        logger.error(mensagem)
        raise ValueError(mensagem)
    
    # Organizar links por data e tipo de mídia
    resultado = {}
    
    for idx, row in df.iterrows():
        # Obter valores da linha
        try:
            titulo = str(row[colunas['titulo']]).strip() if pd.notna(row[colunas['titulo']]) else f"Matéria {idx+1}"
            data_str = str(row[colunas['data']]).strip() if pd.notna(row[colunas['data']]) else datetime.now().strftime('%Y-%m-%d')
            tipo_midia = str(row[colunas['tipo_midia']]).strip() if pd.notna(row[colunas['tipo_midia']]) else "Outros"
            link_web_imagem = str(row[colunas['link_web_imagem']]).strip() if pd.notna(row[colunas['link_web_imagem']]) else ""
            
            # Processar a data para formato yyyy-mm-dd
            try:
                if isinstance(row[colunas['data']], datetime):
                    data_formatada = row[colunas['data']].strftime('%Y-%m-%d')
                else:
                    # Tentar converter várias formatações de data
                    data_obj = pd.to_datetime(data_str, errors='coerce')
                    if pd.isna(data_obj):
                        data_formatada = datetime.now().strftime('%Y-%m-%d')
                    else:
                        data_formatada = data_obj.strftime('%Y-%m-%d')
            except:
                data_formatada = datetime.now().strftime('%Y-%m-%d')
            
            # Verificar se o link é válido
            if link_web_imagem and link_web_imagem.startswith(('http://', 'https://')):
                # Criar estrutura para a data se não existir
                if data_formatada not in resultado:
                    resultado[data_formatada] = {}
                
                # Criar estrutura para o tipo de mídia se não existir
                if tipo_midia not in resultado[data_formatada]:
                    resultado[data_formatada][tipo_midia] = []
                
                # Adicionar link ao resultado
                arquivo_info = {
                    'titulo': titulo,
                    'link': link_web_imagem
                }
                resultado[data_formatada][tipo_midia].append(arquivo_info)
                logger.info(f"Link encontrado para download: {tipo_midia} / {titulo}")
            
        except Exception as e:
            logger.error(f"Erro ao processar linha {idx}: {str(e)}")
    
    # Calcular estatísticas
    total_links = 0
    total_datas = len(resultado)
    total_tipos = 0
    
    for data in resultado:
        for tipo in resultado[data]:
            total_tipos += 1
            total_links += len(resultado[data][tipo])
    
    logger.info(f"Processamento concluído. Encontrados {total_links} links em {total_datas} datas e {total_tipos} tipos de mídia.")
    
    return resultado

# Função para baixar arquivos e organizá-los em pastas
def baixar_arquivos(dados):
    """
    Baixa arquivos a partir dos links fornecidos e organiza em pastas.
    
    Args:
        dados: Dicionário com links organizados por data e tipo de mídia
        
    Returns:
        Um dicionário com estatísticas de download
    """
    logger.info("Iniciando processo de download de arquivos")
    
    # Criar diretório para downloads
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads", "BrasPub_Downloads")
    os.makedirs(download_dir, exist_ok=True)
    logger.info(f"Diretório para downloads: {download_dir}")
    
    # Estatísticas
    status = {}
    total_arquivos = 0
    total_baixados = 0
    total_erros = 0
    
    # Processar cada data
    for data in dados:
        data_dir = os.path.join(download_dir, data)
        os.makedirs(data_dir, exist_ok=True)
        
        # Processar cada tipo de mídia
        for tipo_midia in dados[data]:
            tipo_dir = os.path.join(data_dir, tipo_midia)
            os.makedirs(tipo_dir, exist_ok=True)
            
            # Inicializar estatísticas para este tipo
            if tipo_midia not in status:
                status[tipo_midia] = {'total': 0, 'baixados': 0, 'erros': 0}
            
            # Processar cada arquivo
            for arquivo_info in dados[data][tipo_midia]:
                total_arquivos += 1
                status[tipo_midia]['total'] += 1
                
                try:
                    # Obter informações do arquivo
                    titulo = arquivo_info['titulo']
                    link = arquivo_info['link']
                    
                    # Criar nome de arquivo seguro
                    nome_arquivo = limpar_nome_arquivo(titulo)
                    
                    # Determinar extensão de arquivo baseada no URL
                    extensao = determinar_extensao(link)
                    
                    # Caminho completo do arquivo
                    caminho_arquivo = os.path.join(tipo_dir, f"{nome_arquivo}{extensao}")
                    
                    # Baixar o arquivo
                    logger.info(f"Baixando arquivo de {link} para {caminho_arquivo}")
                    
                    # Adicionar atraso entre downloads para evitar bloqueios
                    time.sleep(0.5)
                    
                    # Tentar baixar o arquivo
                    baixar_arquivo(link, caminho_arquivo)
                    
                    # Atualizar estatísticas
                    total_baixados += 1
                    status[tipo_midia]['baixados'] += 1
                    logger.info(f"Arquivo baixado com sucesso: {caminho_arquivo}")
                    
                except Exception as e:
                    total_erros += 1
                    status[tipo_midia]['erros'] += 1
                    logger.error(f"Erro ao baixar arquivo: {str(e)}")
    
    # Resumo
    logger.info(f"Download concluído. Total: {total_arquivos}, Baixados: {total_baixados}, Erros: {total_erros}")
    
    return status

# Função para limpar nome de arquivo
def limpar_nome_arquivo(nome):
    """Remove caracteres inválidos do nome do arquivo."""
    # Substituir caracteres não permitidos
    nome_limpo = re.sub(r'[\\/*?:"<>|]', "", nome)
    # Substituir espaços por underscores
    nome_limpo = re.sub(r'\s+', '_', nome_limpo)
    # Limitar tamanho
    if len(nome_limpo) > 100:
        nome_limpo = nome_limpo[:100]
    return nome_limpo

# Função para determinar extensão de arquivo com base no URL
def determinar_extensao(url):
    """Determina a extensão do arquivo com base no URL."""
    # Extrair nome do arquivo da URL
    parsed_url = urlparse(url)
    path = unquote(parsed_url.path)
    
    # Verificar extensão no caminho
    nome_arquivo = os.path.basename(path)
    _, extensao = os.path.splitext(nome_arquivo)
    
    # Se encontrou extensão, retornar
    if extensao and extensao.startswith('.'):
        return extensao
    
    # Verificar tipo de mídia com base na URL
    if any(ext in url.lower() for ext in ['.jpg', '.jpeg', '.png', '.gif']):
        return '.jpg'
    elif '.mp3' in url.lower():
        return '.mp3'
    elif '.mp4' in url.lower():
        return '.mp4'
    elif '.pdf' in url.lower():
        return '.pdf'
    
    # Padrão para links sem extensão definida
    return '.jpg'

# Função para baixar arquivo
def baixar_arquivo(url, caminho_destino):
    """Baixa arquivo da URL para o destino especificado."""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        # Fazer requisição
        response = requests.get(url, headers=headers, stream=True, timeout=30)
        response.raise_for_status()
        
        # Salvar arquivo
        with open(caminho_destino, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        return True
    except Exception as e:
        logger.error(f"Erro ao baixar {url}: {str(e)}")
        raise

if __name__ == "__main__":
    logger.info("Iniciando servidor Flask para API de processamento...")
    app.run(debug=True, port=5000) 