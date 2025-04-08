from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import json
import uuid
from datetime import datetime, date, time
from organizador import processar_planilha, exportar_planilha, json_serial
from organizador_keywords import processar_planilha_keywords, exportar_planilha_keywords

app = Flask(__name__)
CORS(app)  # Habilitar CORS para todas as rotas

# Diretório para arquivos temporários
TEMP_DIR = os.path.join(tempfile.gettempdir(), 'organizador_planilhas')
os.makedirs(TEMP_DIR, exist_ok=True)

# Função auxiliar para serializar objetos datetime e time
def serializar_para_json(dados):
    return json.dumps(dados, default=lambda obj: obj.isoformat() if isinstance(obj, (datetime, date, time)) else str(obj) if not isinstance(obj, (str, int, float, bool, list, dict, type(None))) else None)

@app.route('/api/processar', methods=['POST'])
def api_processar():
    """
    Recebe um arquivo Excel, processa e retorna os dados organizados.
    """
    # Verificar se existe um arquivo na requisição
    if 'arquivo' not in request.files:
        return jsonify({'status': 'erro', 'mensagem': 'Nenhum arquivo enviado'}), 400
        
    arquivo = request.files['arquivo']
    
    # Verificar se o nome do arquivo não está vazio
    if arquivo.filename == '':
        return jsonify({'status': 'erro', 'mensagem': 'Nome de arquivo vazio'}), 400
        
    # Verificar a extensão do arquivo
    if not arquivo.filename.endswith(('.xls', '.xlsx')):
        return jsonify({'status': 'erro', 'mensagem': 'Formato de arquivo inválido. Use .xls ou .xlsx'}), 400
    
    try:
        # Salvar o arquivo temporariamente
        temp_path = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}_{arquivo.filename}")
        arquivo.save(temp_path)
        
        # Processar a planilha
        resultado = processar_planilha(temp_path)
        
        # Remover o arquivo temporário
        os.remove(temp_path)
        
        # Usar a função de serialização manual
        return app.response_class(
            response=serializar_para_json(resultado),
            status=200,
            mimetype='application/json'
        )
        
    except Exception as e:
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

@app.route('/api/exportar', methods=['POST'])
def api_exportar():
    """
    Recebe dados processados e retorna um arquivo Excel.
    """
    # Verificar se os dados foram enviados
    if not request.json or 'dados' not in request.json:
        return app.response_class(
            response=serializar_para_json({'status': 'erro', 'mensagem': 'Dados não fornecidos'}),
            status=400,
            mimetype='application/json'
        )
    
    try:
        # Obter os dados
        dados = request.json['dados']
        
        # Gerar nomes de arquivos temporários
        temp_xlsx = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}.xlsx")
        
        # Exportar para Excel
        resultado = exportar_planilha(dados, temp_xlsx)
        
        if resultado['status'] == 'sucesso':
            # Enviar o arquivo
            response = send_file(
                temp_xlsx,
                as_attachment=True,
                download_name="Planilha_Organizada.xlsx",
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
            # Configurar callback para remover o arquivo após o envio
            @response.call_on_close
            def remove_temp_file():
                if os.path.exists(temp_xlsx):
                    os.remove(temp_xlsx)
                    
            return response
        else:
            return app.response_class(
                response=serializar_para_json(resultado),
                status=500,
                mimetype='application/json'
            )
        
    except Exception as e:
        return app.response_class(
            response=serializar_para_json({'status': 'erro', 'mensagem': str(e)}),
            status=500,
            mimetype='application/json'
        )

@app.route('/api/processar_keywords', methods=['POST'])
def api_processar_keywords():
    """
    Recebe um arquivo Excel de palavras-chave, processa e retorna os dados organizados.
    Para cada palavra-chave, são gerados quatro tipos de mídia (Portal, Impresso, TV, Rádio),
    cada um com seu formato específico de arquivo, extraído diretamente da página da matéria:
    - Portal: Arquivo PDF ou link da página web
    - Impresso: Imagem JPEG/JPG da matéria impressa
    - TV: Vídeo MP4 ou link para vídeo
    - Rádio: Áudio MP3 ou link para áudio
    """
    # Verificar se existe um arquivo na requisição
    if 'arquivo' not in request.files:
        return jsonify({'status': 'erro', 'mensagem': 'Nenhum arquivo enviado'}), 400
        
    arquivo = request.files['arquivo']
    
    # Verificar se o nome do arquivo não está vazio
    if arquivo.filename == '':
        return jsonify({'status': 'erro', 'mensagem': 'Nome de arquivo vazio'}), 400
        
    # Verificar a extensão do arquivo
    if not arquivo.filename.endswith(('.xls', '.xlsx')):
        return jsonify({'status': 'erro', 'mensagem': 'Formato de arquivo inválido. Use .xls ou .xlsx'}), 400
    
    try:
        # Salvar o arquivo temporariamente
        temp_path = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}_{arquivo.filename}")
        arquivo.save(temp_path)
        
        # Processar a planilha de palavras-chave
        resultado = processar_planilha_keywords(temp_path)
        
        # Remover o arquivo temporário
        os.remove(temp_path)
        
        # Usar a função de serialização manual
        return app.response_class(
            response=serializar_para_json(resultado),
            status=200,
            mimetype='application/json'
        )
        
    except Exception as e:
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

@app.route('/api/exportar_keywords', methods=['POST'])
def api_exportar_keywords():
    """
    Recebe dados de palavras-chave processados e retorna um arquivo Excel organizado por palavras-chave.
    Cada palavra-chave aparecerá somente uma vez. Para cada palavra-chave, serão gerados
    quatro tipos de mídia, cada um com seu link específico extraído da página da matéria:
    - Portal: Arquivo PDF ou link da página web
    - Impresso: Imagem JPEG/JPG da matéria impressa
    - TV: Vídeo MP4 ou link para vídeo
    - Rádio: Áudio MP3 ou link para áudio
    """
    # Verificar se os dados foram enviados
    if not request.json or 'dados' not in request.json:
        return app.response_class(
            response=serializar_para_json({'status': 'erro', 'mensagem': 'Dados não fornecidos'}),
            status=400,
            mimetype='application/json'
        )
    
    try:
        # Obter os dados
        dados = request.json['dados']
        
        # Gerar nomes de arquivos temporários
        temp_xlsx = os.path.join(TEMP_DIR, f"{uuid.uuid4().hex}.xlsx")
        
        # Exportar para Excel usando a função específica para palavras-chave
        resultado = exportar_planilha_keywords(dados, temp_xlsx)
        
        if resultado['status'] == 'sucesso':
            # Enviar o arquivo
            response = send_file(
                temp_xlsx,
                as_attachment=True,
                download_name="Palavras_Chave_Organizadas.xlsx",
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
            # Configurar callback para remover o arquivo após o envio
            @response.call_on_close
            def remove_temp_file():
                if os.path.exists(temp_xlsx):
                    os.remove(temp_xlsx)
                    
            return response
        else:
            return app.response_class(
                response=serializar_para_json(resultado),
                status=500,
                mimetype='application/json'
            )
        
    except Exception as e:
        return app.response_class(
            response=serializar_para_json({'status': 'erro', 'mensagem': str(e)}),
            status=500,
            mimetype='application/json'
        )

@app.route('/api/status', methods=['GET'])
def api_status():
    """
    Verifica se a API está funcionando.
    """
    return app.response_class(
        response=serializar_para_json({'status': 'online', 'mensagem': 'API do Organizador de Planilhas está funcionando'}),
        status=200,
        mimetype='application/json'
    )

if __name__ == '__main__':
    # Iniciar o servidor na porta 5000
    app.run(host='0.0.0.0', port=5000) 