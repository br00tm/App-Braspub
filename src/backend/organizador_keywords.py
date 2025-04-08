import pandas as pd
import os
import json
import sys
import re
import logging
from datetime import datetime, date, time
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("extrator_midia.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("ExtractorMidia")

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

def converter_para_url_absoluta(url, base_url):
    """
    Converte uma URL relativa para absoluta usando a URL base.
    
    Args:
        url: URL que pode ser relativa
        base_url: URL base para converter URLs relativas
        
    Returns:
        URL absoluta
    """
    # Verificar se a URL já é absoluta
    if bool(urlparse(url).netloc):
        return url
    
    # Converter URL relativa para absoluta
    return urljoin(base_url, url)

def obter_link_por_tipo_midia(url_base, tipo_midia):
    """
    Obtém o link correto para o tipo de mídia a partir da página.
    Busca especificamente elementos dentro das divs na estrutura DOM da página.
    
    Args:
        url_base: URL base da matéria
        tipo_midia: Tipo de mídia (Portal, Impresso, TV, Rádio)
        
    Returns:
        URL para o arquivo de mídia adequado ou URL base se não conseguir extrair
    """
    try:
        # Verificar se a URL é válida
        if not url_base or not url_base.startswith(('http://', 'https://')):
            # Retornar URL padrão com extensão adequada
            extensoes = {
                'Portal': '.pdf',
                'Impresso': '.jpg',
                'TV': '.mp4',
                'Rádio': '.mp3'
            }
            logger.warning(f"URL inválida para {tipo_midia}: {url_base}. Usando URL padrão.")
            return url_base + extensoes.get(tipo_midia, '')
        
        logger.info(f"Buscando mídia para: {tipo_midia} na URL: {url_base}")
        
        # Fazer requisição para a página
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        logger.info(f"Fazendo requisição para: {url_base}")
        try:
            response = requests.get(url_base, headers=headers, timeout=15)
            if response.status_code != 200:
                logger.warning(f"Falha ao acessar URL: {url_base}, status: {response.status_code}")
                return url_base
            
            logger.info(f"Página acessada com sucesso. Analisando HTML para o tipo: {tipo_midia}")
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Dependendo do tipo de mídia, buscar elementos diferentes
            if tipo_midia == 'Portal':
                # Buscar elementos <a> com href que contenha "getPDF"
                pdf_links = []
                for link in soup.find_all('a', href=True):
                    href = link['href']
                    if 'getpdf' in href.lower():
                        pdf_links.append(href)
                        logger.info(f"Encontrado link PDF com getPDF: {href}")
                
                # Se encontrou algum link com getPDF, retornar o primeiro
                if pdf_links:
                    # Converter URL relativa para absoluta se necessário
                    result = converter_para_url_absoluta(pdf_links[0], url_base)
                    logger.info(f"Retornando link PDF direto: {result}")
                    return result
                
                # Se não encontrou link com getPDF, buscar outros links de PDF
                for link in soup.find_all('a', href=True):
                    href = link['href'].lower()
                    if '.pdf' in href:
                        pdf_links.append(link['href'])
                        logger.info(f"Encontrado link PDF: {link['href']}")
                
                # Buscar também em iframes ou objetos embed que possam conter PDFs
                for embed in soup.find_all(['embed', 'iframe', 'object'], src=True):
                    src = embed['src'].lower()
                    if '.pdf' in src:
                        pdf_links.append(embed['src'])
                        logger.info(f"Encontrado embed PDF: {embed['src']}")
                
                # Se encontrou algum PDF, retornar o primeiro
                if pdf_links:
                    # Converter URL relativa para absoluta
                    result = converter_para_url_absoluta(pdf_links[0], url_base)
                    logger.info(f"Retornando link PDF: {result}")
                    return result
                
                # Se não encontrou nenhum PDF, retornar a URL original
                logger.info(f"Nenhum PDF encontrado. Retornando URL original: {url_base}")
                return url_base
                
            elif tipo_midia == 'Impresso':
                # Buscar imagens em divs específicas
                img_links = []
                
                # Buscar imagens na div da classe imagem-container e outras classes relevantes
                for img_container in soup.select('div.imagem-container, div.image-container, div[data-v-6c6e7f38] div.imagem-container, div.figura'):
                    # Buscar imagens dentro da div
                    for img in img_container.find_all('img', src=True):
                        img_links.append(img['src'])
                        logger.info(f"Encontrada imagem em container específico: {img['src']}")
                    
                    # Buscar também elementos com data-src
                    for img in img_container.find_all(attrs={'data-src': True}):
                        img_links.append(img['data-src'])
                        logger.info(f"Encontrada imagem com data-src em container: {img['data-src']}")
                
                # Buscar imagens com classe específica
                if not img_links:
                    for img in soup.select('img.imagem-full, img.materia-imagem, img.full-image'):
                        if img.get('src'):
                            img_links.append(img['src'])
                            logger.info(f"Encontrada imagem com classe específica: {img['src']}")
                
                # Buscar imagens com nomes específicos
                if not img_links:
                    for img in soup.find_all('img', src=True):
                        src = img['src'].lower()
                        if any(pattern in src for pattern in ['site.jpg', 'impresso.jpg', 'noticia.jpg', 'materia']):
                            img_links.append(img['src'])
                            logger.info(f"Encontrada imagem com nome específico: {img['src']}")
                
                # Buscar qualquer imagem grande que possa ser a principal
                if not img_links:
                    for img in soup.find_all('img', src=True):
                        src = img['src'].lower()
                        if '.jpg' in src or '.jpeg' in src or '.png' in src:
                            # Evitar ícones pequenos
                            if 'icon' not in src and 'logo' not in src:
                                img_links.append(img['src'])
                                logger.info(f"Encontrada imagem JPG/JPEG/PNG: {img['src']}")
                
                # Se encontrou alguma imagem, retornar a primeira
                if img_links:
                    result = converter_para_url_absoluta(img_links[0], url_base)
                    logger.info(f"Retornando link de imagem: {result}")
                    return result
                
                # Se não encontrou nenhuma imagem, retornar URL com extensão .jpg
                result = construir_url_padrao(url_base, tipo_midia)
                logger.warning(f"Nenhuma imagem encontrada. Usando URL padrão: {result}")
                return result
                
            elif tipo_midia == 'TV':
                # Buscar vídeos em divs específicas
                video_links = []
                
                # Buscar em divs específicas que possam conter vídeos
                divs_video = soup.select('div.video-container, div.player, div.materia-video, div[data-v-6c6e7f38]')
                
                for div in divs_video:
                    # Buscar elementos de vídeo dentro dessas divs
                    for video in div.find_all('video'):
                        # Buscar source mp4
                        for source in video.find_all('source'):
                            if source.get('src') and '.mp4' in source['src'].lower():
                                video_links.append(source['src'])
                                logger.info(f"Encontrado source MP4 em div específica: {source['src']}")
                        
                        # Verificar se o próprio vídeo tem src
                        if video.get('src') and '.mp4' in video['src'].lower():
                            video_links.append(video['src'])
                            logger.info(f"Encontrado vídeo com src em div específica: {video['src']}")
                    
                    # Buscar links mp4
                    for link in div.find_all('a', href=True):
                        if '.mp4' in link['href'].lower():
                            video_links.append(link['href'])
                            logger.info(f"Encontrado link MP4 em div específica: {link['href']}")
                    
                    # Buscar iframes (YouTube, etc.)
                    for iframe in div.find_all('iframe', src=True):
                        video_links.append(iframe['src'])
                        logger.info(f"Encontrado iframe em div específica: {iframe['src']}")
                
                # Se não encontrou em divs específicas, buscar em toda a página
                if not video_links:
                    # Buscar sources de vídeo MP4
                    for source in soup.select('video source[src]'):
                        src = source['src'].lower()
                        if '.mp4' in src:
                            video_links.append(source['src'])
                            logger.info(f"Encontrado source de vídeo: {source['src']}")
                    
                    # Buscar links para arquivos MP4
                    for link in soup.find_all('a', href=True):
                        href = link['href'].lower()
                        if '.mp4' in href:
                            video_links.append(link['href'])
                            logger.info(f"Encontrado link para vídeo: {link['href']}")
                    
                    # Buscar player de vídeo
                    for video in soup.find_all('video', src=True):
                        src = video['src'].lower()
                        if '.mp4' in src:
                            video_links.append(video['src'])
                            logger.info(f"Encontrado elemento de vídeo: {video['src']}")
                    
                    # Buscar iframes para serviços de vídeo (YouTube, Vimeo)
                    for iframe in soup.find_all('iframe', src=True):
                        src = iframe['src'].lower()
                        if 'youtube' in src or 'vimeo' in src or 'video' in src:
                            video_links.append(iframe['src'])
                            logger.info(f"Encontrado iframe de vídeo: {iframe['src']}")
                
                # Se encontrou algum vídeo, retornar o primeiro
                if video_links:
                    result = converter_para_url_absoluta(video_links[0], url_base)
                    logger.info(f"Retornando link de vídeo: {result}")
                    return result
                
                # Se não encontrou nenhum vídeo, retornar URL com extensão .mp4
                result = construir_url_padrao(url_base, tipo_midia)
                logger.warning(f"Nenhum vídeo encontrado. Usando URL padrão: {result}")
                return result
                
            elif tipo_midia == 'Rádio':
                # Buscar áudios em divs específicas
                audio_links = []
                
                # Buscar em divs específicas que possam conter áudios
                divs_audio = soup.select('div.audio-container, div.player, div.materia-audio, div[data-v-6c6e7f38]')
                
                for div in divs_audio:
                    # Buscar elementos de áudio dentro dessas divs
                    for audio in div.find_all('audio'):
                        # Buscar source mp3
                        for source in audio.find_all('source'):
                            if source.get('src') and '.mp3' in source['src'].lower():
                                audio_links.append(source['src'])
                                logger.info(f"Encontrado source MP3 em div específica: {source['src']}")
                        
                        # Verificar se o próprio áudio tem src
                        if audio.get('src') and '.mp3' in audio['src'].lower():
                            audio_links.append(audio['src'])
                            logger.info(f"Encontrado áudio com src em div específica: {audio['src']}")
                    
                    # Buscar links mp3
                    for link in div.find_all('a', href=True):
                        if '.mp3' in link['href'].lower():
                            audio_links.append(link['href'])
                            logger.info(f"Encontrado link MP3 em div específica: {link['href']}")
                
                # Se não encontrou em divs específicas, buscar em toda a página
                if not audio_links:
                    # Buscar sources de áudio MP3
                    for source in soup.select('audio source[src]'):
                        src = source['src'].lower()
                        if '.mp3' in src:
                            audio_links.append(source['src'])
                            logger.info(f"Encontrado source de áudio: {source['src']}")
                    
                    # Buscar links para arquivos MP3
                    for link in soup.find_all('a', href=True):
                        href = link['href'].lower()
                        if '.mp3' in href:
                            audio_links.append(link['href'])
                            logger.info(f"Encontrado link para áudio: {link['href']}")
                    
                    # Buscar player de áudio
                    for audio in soup.find_all('audio', src=True):
                        src = audio['src'].lower()
                        if '.mp3' in src:
                            audio_links.append(audio['src'])
                            logger.info(f"Encontrado elemento de áudio: {audio['src']}")
                
                # Se encontrou algum áudio, retornar o primeiro
                if audio_links:
                    result = converter_para_url_absoluta(audio_links[0], url_base)
                    logger.info(f"Retornando link de áudio: {result}")
                    return result
                
                # Se não encontrou nenhum áudio, retornar URL com extensão .mp3
                result = construir_url_padrao(url_base, tipo_midia)
                logger.warning(f"Nenhum áudio encontrado. Usando URL padrão: {result}")
                return result
            
            # Tipo de mídia não reconhecido ou não encontrou nada específico
            logger.warning(f"Tipo de mídia não reconhecido: {tipo_midia}")
            return url_base
            
        except Exception as e:
            logger.error(f"Erro ao acessar URL: {str(e)}")
            return url_base
            
    except Exception as e:
        logger.error(f"Erro ao extrair link para {tipo_midia}: {str(e)}", exc_info=True)
        # Em caso de erro, retornar URL base
        logger.info(f"Retornando URL base devido a erro: {url_base}")
        return url_base

def construir_url_padrao(url_base, tipo_midia):
    """
    Constrói uma URL padrão com a extensão adequada para o tipo de mídia.
    
    Args:
        url_base: URL base
        tipo_midia: Tipo de mídia
        
    Returns:
        URL com extensão adequada
    """
    extensoes = {
        'Portal': '.pdf',
        'Impresso': '.jpg',
        'TV': '.mp4',
        'Rádio': '.mp3'
    }
    
    # Remover extensão existente se houver
    for ext in ['.pdf', '.jpg', '.jpeg', '.mp4', '.mp3', '.html', '.htm']:
        if url_base.lower().endswith(ext):
            url_base = url_base[:-len(ext)]
    
    # Adicionar a extensão correta
    return url_base + extensoes.get(tipo_midia, '')

def processar_planilha_keywords(caminho_arquivo):
    """
    Processa a planilha Excel com palavras-chave e organiza os dados.
    Cada palavra-chave terá exatamente um registro para cada tipo de mídia: Portal, Impresso, TV e Rádio.
    O link para cada tipo de mídia será extraído da página correspondente.
    
    Args:
        caminho_arquivo: Caminho para o arquivo Excel a ser processado
        
    Returns:
        Um dicionário com os dados organizados por palavra-chave
    """
    try:
        # Carregar a planilha Excel
        df = pd.read_excel(caminho_arquivo)
        
        # Exibir colunas disponíveis para debug
        logger.info(f"Colunas disponíveis: {df.columns.tolist()}")
        print("Colunas disponíveis:", df.columns.tolist())
        
        # Debug: Mostrar as primeiras linhas da planilha
        logger.info("Primeiras linhas da planilha:")
        for idx, row in df.head().iterrows():
            logger.info(f"Linha {idx}: {dict(row)}")
        
        # Mapear colunas originais para novos nomes simplificados conforme imagem
        # Incluindo possíveis variações de nomes de colunas
        mapeamento_colunas = {
            'Palavra-chave': 'PALAVRAS-CHAVE',
            'Palavra chave': 'PALAVRAS-CHAVE',
            'Keyword': 'PALAVRAS-CHAVE',
            'PALAVRAS-CHAVE': 'PALAVRAS-CHAVE',
            'Assunto': 'PALAVRAS-CHAVE',
            'Cliente': 'PALAVRAS-CHAVE',
            'Data de inclusão': 'DATA DE CADASTRO',
            'Data': 'DATA DE CADASTRO',
            'DATA DE CADASTRO': 'DATA DE CADASTRO',
            'Título': 'TÍTULO DA MATÉRIA',
            'TÍTULO DA MATÉRIA': 'TÍTULO DA MATÉRIA',
            'Veículo': 'VEÍCULO',
            'Tipo da mídia': 'TIPO DE MÍDIA',
            'TIPO DE MÍDIA': 'TIPO DE MÍDIA',
            'Portal': 'TIPO DE MÍDIA',
            'Online': 'TIPO DE MÍDIA',
            'Link da matéria cadastrada': 'LINK DA MATÉRIA CADASTRADA',
            'LINK DA MATÉRIA CADASTRADA': 'LINK DA MATÉRIA CADASTRADA',
            'Link original': 'LINK ORIGINAL',
            'LINK ORIGINAL': 'LINK ORIGINAL',
            'Link': 'LINK DA MATÉRIA CADASTRADA',
            'URL': 'LINK DA MATÉRIA CADASTRADA',
            'Endereço': 'LINK DA MATÉRIA CADASTRADA',
            'Link web - Imagem': 'LINK_WEB_IMAGEM',
            'Link web - Texto': 'LINK_WEB_TEXTO',
            'Link Materia': 'LINK_WEB_TEXTO'
        }
        
        # Se estamos lidando com índices de colunas em vez de nomes
        # Tentar identificar colunas por padrões nas planilhas de exemplo
        colunas_por_indice = {}
        try:
            for i, col in enumerate(df.columns):
                col_lower = str(col).lower()
                if 'palavra' in col_lower or 'chave' in col_lower or 'assunto' in col_lower:
                    colunas_por_indice[col] = 'PALAVRAS-CHAVE'
                elif 'data' in col_lower or 'cadastro' in col_lower or 'inclusão' in col_lower:
                    colunas_por_indice[col] = 'DATA DE CADASTRO'
                elif 'título' in col_lower or 'matéria' in col_lower or 'assunto' in col_lower:
                    colunas_por_indice[col] = 'TÍTULO DA MATÉRIA'
                elif 'tipo' in col_lower or 'mídia' in col_lower or 'portal' in col_lower:
                    colunas_por_indice[col] = 'TIPO DE MÍDIA'
                elif 'link web' in col_lower and 'imagem' in col_lower:
                    colunas_por_indice[col] = 'LINK_WEB_IMAGEM'
                elif 'link web' in col_lower and 'texto' in col_lower:
                    colunas_por_indice[col] = 'LINK_WEB_TEXTO'
                elif 'link materia' in col_lower:
                    colunas_por_indice[col] = 'LINK_WEB_TEXTO'
                elif 'link da matéria cadastrada' in col_lower:
                    colunas_por_indice[col] = 'LINK DA MATÉRIA CADASTRADA'
                elif 'link original' in col_lower:
                    colunas_por_indice[col] = 'LINK ORIGINAL'
                elif 'link' in col_lower or 'url' in col_lower or 'endereço' in col_lower:
                    # Se já encontramos links mais específicos, não sobrescrever
                    if not any(nome in colunas_por_indice.values() for nome in ['LINK DA MATÉRIA CADASTRADA', 'LINK ORIGINAL', 'LINK_WEB_IMAGEM', 'LINK_WEB_TEXTO']):
                        colunas_por_indice[col] = 'LINK DA MATÉRIA CADASTRADA'
        except Exception as e:
            # Algumas colunas podem causar erro, o que é aceitável
            logger.error(f"Erro ao processar colunas: {str(e)}")
            pass
        
        # Adicionar as colunas por índice ao mapeamento
        mapeamento_colunas.update(colunas_por_indice)
        
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
                logger.info(f"Mapeada coluna: {col_original} -> {col_nova}")
        
        # Debug: Mostrando colunas no novo DataFrame
        logger.info(f"Colunas no novo DataFrame: {novo_df.columns.tolist()}")
        
        # Verificar se temos a coluna de palavras-chave
        if 'PALAVRAS-CHAVE' not in novo_df.columns:
            # Se não temos a coluna de palavras-chave, procurar por ela de outra forma
            for col in df.columns:
                col_lower = str(col).lower()
                if any(termo in col_lower for termo in ['palavra', 'chave', 'keyword', 'assunto']):
                    novo_df['PALAVRAS-CHAVE'] = df[col]
                    break
            
            # Se ainda não encontramos, tentar usar a primeira coluna
            if 'PALAVRAS-CHAVE' not in novo_df.columns and len(df.columns) > 0:
                novo_df['PALAVRAS-CHAVE'] = df.iloc[:, 0]
        
        # Se ainda não temos a coluna PALAVRAS-CHAVE, não podemos continuar
        if 'PALAVRAS-CHAVE' not in novo_df.columns:
            return {'status': 'erro', 'mensagem': 'Não foi possível identificar a coluna de PALAVRAS-CHAVE.'}
        
        # Verificar se temos a coluna Link web - Imagem
        if 'LINK_WEB_IMAGEM' not in novo_df.columns:
            # Tentar encontrar a coluna de links web imagem de outra forma
            for col in df.columns:
                col_lower = str(col).lower()
                if 'link web' in col_lower and 'imagem' in col_lower:
                    novo_df['LINK_WEB_IMAGEM'] = df[col]
                    logger.info(f"Coluna 'Link web - Imagem' encontrada como: {col}")
                    # Debug: Mostrar valores da coluna encontrada
                    logger.info(f"Valores da coluna '{col}': {df[col].tolist()}")
                    break
            
            # Se ainda não encontramos, procurar qualquer coluna que contenha "imagem"
            if 'LINK_WEB_IMAGEM' not in novo_df.columns:
                for col in df.columns:
                    col_lower = str(col).lower()
                    if 'imagem' in col_lower and ('link' in col_lower or 'url' in col_lower):
                        novo_df['LINK_WEB_IMAGEM'] = df[col]
                        logger.info(f"Coluna similar a 'Link web - Imagem' encontrada como: {col}")
                        # Debug: Mostrar valores da coluna encontrada
                        logger.info(f"Valores da coluna '{col}': {df[col].tolist()}")
                        break
        else:
            # Se já temos a coluna, mostrar os valores
            logger.info(f"Valores da coluna 'LINK_WEB_IMAGEM': {novo_df['LINK_WEB_IMAGEM'].tolist()}")
        
        # Verificar se temos a coluna Link web - Texto
        if 'LINK_WEB_TEXTO' not in novo_df.columns:
            # Tentar encontrar a coluna de links web texto de outra forma
            for col in df.columns:
                col_lower = str(col).lower()
                if 'link web' in col_lower and 'texto' in col_lower:
                    novo_df['LINK_WEB_TEXTO'] = df[col]
                    logger.info(f"Coluna 'Link web - Texto' encontrada como: {col}")
                    # Debug: Mostrar valores da coluna encontrada
                    logger.info(f"Valores da coluna '{col}': {df[col].tolist()}")
                    break
                elif 'link materia' in col_lower:
                    novo_df['LINK_WEB_TEXTO'] = df[col]
                    logger.info(f"Coluna 'Link Materia' encontrada como: {col}")
                    # Debug: Mostrar valores da coluna encontrada
                    logger.info(f"Valores da coluna '{col}': {df[col].tolist()}")
                    break
            
            # Se ainda não encontramos, procurar qualquer coluna que contenha "texto" ou "materia"
            if 'LINK_WEB_TEXTO' not in novo_df.columns:
                for col in df.columns:
                    col_lower = str(col).lower()
                    if ('texto' in col_lower or 'materia' in col_lower) and ('link' in col_lower or 'url' in col_lower):
                        novo_df['LINK_WEB_TEXTO'] = df[col]
                        logger.info(f"Coluna similar a 'Link web - Texto' encontrada como: {col}")
                        # Debug: Mostrar valores da coluna encontrada
                        logger.info(f"Valores da coluna '{col}': {df[col].tolist()}")
                        break
        else:
            # Se já temos a coluna, mostrar os valores
            logger.info(f"Valores da coluna 'LINK_WEB_TEXTO': {novo_df['LINK_WEB_TEXTO'].tolist()}")
        
        # Verificar se temos a coluna de links cadastrados
        if 'LINK DA MATÉRIA CADASTRADA' not in novo_df.columns:
            # Tentar encontrar a coluna de links de outra forma
            for col in df.columns:
                col_lower = str(col).lower()
                if 'link' in col_lower and ('matéria' in col_lower or 'cadastrada' in col_lower):
                    novo_df['LINK DA MATÉRIA CADASTRADA'] = df[col]
                    break
                elif 'link original' in col_lower:
                    novo_df['LINK ORIGINAL'] = df[col]
                elif any(termo in col_lower for termo in ['link', 'url', 'endereço', 'http']):
                    novo_df['LINK DA MATÉRIA CADASTRADA'] = df[col]
                    break
            
            # Se ainda não encontramos, criar coluna vazia
            if 'LINK DA MATÉRIA CADASTRADA' not in novo_df.columns:
                novo_df['LINK DA MATÉRIA CADASTRADA'] = ''
        
        # Limpar e normalizar as palavras-chave
        novo_df['PALAVRAS-CHAVE'] = novo_df['PALAVRAS-CHAVE'].astype(str).apply(lambda x: x.strip())
        
        # Garantir que temos a coluna TIPO DE MÍDIA
        if 'TIPO DE MÍDIA' not in novo_df.columns:
            # Procurar por colunas relacionadas a tipo de mídia
            for col in df.columns:
                col_lower = str(col).lower()
                if any(termo in col_lower for termo in ['tipo', 'mídia', 'portal', 'impresso', 'tv', 'rádio']):
                    novo_df['TIPO DE MÍDIA'] = df[col]
                    break
            
            # Se ainda não encontramos, criar com valor padrão
            if 'TIPO DE MÍDIA' not in novo_df.columns:
                novo_df['TIPO DE MÍDIA'] = 'Portal'  # Valor padrão
        
        # Converter 'Online' para 'Portal' na coluna TIPO DE MÍDIA
        if 'TIPO DE MÍDIA' in novo_df.columns:
            novo_df['TIPO DE MÍDIA'] = novo_df['TIPO DE MÍDIA'].replace('Online', 'Portal')
        
        # Adicionar DATA DE CADASTRO se não existir
        if 'DATA DE CADASTRO' not in novo_df.columns:
            # Procurar por colunas relacionadas a datas
            for col in df.columns:
                col_lower = str(col).lower()
                if any(termo in col_lower for termo in ['data', 'cadastro', 'inclusão', 'publicação']):
                    novo_df['DATA DE CADASTRO'] = df[col]
                    break
            
            # Se ainda não encontramos, criar coluna vazia
            if 'DATA DE CADASTRO' not in novo_df.columns:
                novo_df['DATA DE CADASTRO'] = ''
        
        # Adicionar TÍTULO DA MATÉRIA se não existir
        if 'TÍTULO DA MATÉRIA' not in novo_df.columns:
            # Procurar por colunas relacionadas a títulos
            for col in df.columns:
                col_lower = str(col).lower()
                if any(termo in col_lower for termo in ['título', 'matéria', 'assunto']):
                    novo_df['TÍTULO DA MATÉRIA'] = df[col]
                    break
            
            # Se ainda não encontramos, criar coluna vazia
            if 'TÍTULO DA MATÉRIA' not in novo_df.columns:
                novo_df['TÍTULO DA MATÉRIA'] = 'Matéria Não Cadastrada'
        
        # Definir os tipos de mídia padrão como na imagem
        tipos_midia_padrao = ['Portal', 'Impresso', 'TV', 'Rádio']
        
        # Extrair palavras-chave únicas, removendo duplicatas e espaços em branco
        palavras_chave = []
        for palavra in novo_df['PALAVRAS-CHAVE'].unique():
            if palavra and str(palavra).strip() != '':
                palavras_chave.append(str(palavra).strip())
        
        # Organizar os dados por palavra-chave
        resultado = {}
        
        # Para cada palavra-chave, criar registros para cada tipo de mídia
        for palavra in palavras_chave:
            # Filtrar os dados pela palavra-chave
            df_palavra = novo_df[novo_df['PALAVRAS-CHAVE'] == palavra].copy()
            logger.info(f"Processando palavra-chave: {palavra}")
            logger.info(f"Colunas para esta palavra: {df_palavra.columns.tolist()}")
            
            # Primeiro verificar se temos o "Link web - Imagem"
            link_web_imagem = None
            if 'LINK_WEB_IMAGEM' in df_palavra.columns:
                for idx, row in df_palavra.iterrows():
                    if row['LINK_WEB_IMAGEM'] and str(row['LINK_WEB_IMAGEM']).strip():
                        link_web_imagem = str(row['LINK_WEB_IMAGEM']).strip()
                        logger.info(f"Link web imagem encontrado para {palavra}: {link_web_imagem}")
                        break
                
                # Se não encontrou link web imagem nos registros, mostrar todos os valores
                if not link_web_imagem:
                    logger.warning(f"Link web imagem não encontrado para {palavra} nos registros")
                    logger.warning(f"Valores disponíveis: {df_palavra['LINK_WEB_IMAGEM'].tolist()}")
            else:
                logger.warning(f"Coluna LINK_WEB_IMAGEM não encontrada para {palavra}")
            
            # Verificar se temos o "Link web - Texto"
            link_web_texto = None
            if 'LINK_WEB_TEXTO' in df_palavra.columns:
                for idx, row in df_palavra.iterrows():
                    if row['LINK_WEB_TEXTO'] and str(row['LINK_WEB_TEXTO']).strip():
                        link_web_texto = str(row['LINK_WEB_TEXTO']).strip()
                        logger.info(f"Link web texto encontrado para {palavra}: {link_web_texto}")
                        break
                
                # Se não encontrou link web texto nos registros, mostrar todos os valores
                if not link_web_texto:
                    logger.warning(f"Link web texto não encontrado para {palavra} nos registros")
                    logger.warning(f"Valores disponíveis: {df_palavra['LINK_WEB_TEXTO'].tolist()}")
            else:
                logger.warning(f"Coluna LINK_WEB_TEXTO não encontrada para {palavra}")
            
            # Detectar o tipo de mídia a partir do link web imagem se disponível
            tipo_midia_detectado = None
            if link_web_imagem and link_web_imagem.startswith(('http://', 'https://')):
                tipo_midia_detectado = detectar_tipo_midia(link_web_imagem)
                logger.info(f"Tipo de mídia detectado para {palavra}: {tipo_midia_detectado}")
            
            # Encontrar o link base para esta palavra-chave
            link_base = ''
            
            # Priorizar o link web texto se disponível
            if link_web_texto:
                link_base = link_web_texto
                logger.info(f"Usando link web texto como base: {link_base}")
            # Se não tem link web texto, priorizar o link web imagem
            elif link_web_imagem:
                link_base = link_web_imagem
                logger.info(f"Usando link web imagem como base: {link_base}")
            else:
                # Tentar obter o link da matéria cadastrada
                for idx, row in df_palavra.iterrows():
                    if 'LINK DA MATÉRIA CADASTRADA' in row and row['LINK DA MATÉRIA CADASTRADA'] and str(row['LINK DA MATÉRIA CADASTRADA']).strip():
                        link_base = str(row['LINK DA MATÉRIA CADASTRADA']).strip()
                        break
                
                # Se não encontrou, tentar o link original
                if not link_base and 'LINK ORIGINAL' in df_palavra.columns:
                    for idx, row in df_palavra.iterrows():
                        if row['LINK ORIGINAL'] and str(row['LINK ORIGINAL']).strip():
                            link_base = str(row['LINK ORIGINAL']).strip()
                            break
            
            # Se ainda não encontrou, usar um link padrão baseado na palavra-chave
            if not link_base:
                link_base = f"https://braspub.com.br/materias/{palavra.replace(' ', '_').lower()}"
            
            # Criar uma lista para armazenar os registros desta palavra-chave
            registros_palavra = []
            
            # Extrair data de cadastro e título da matéria
            data_cadastro = df_palavra['DATA DE CADASTRO'].iloc[0] if len(df_palavra) > 0 and 'DATA DE CADASTRO' in df_palavra.columns else ''
            titulo_materia = df_palavra['TÍTULO DA MATÉRIA'].iloc[0] if len(df_palavra) > 0 and 'TÍTULO DA MATÉRIA' in df_palavra.columns and str(df_palavra['TÍTULO DA MATÉRIA'].iloc[0]).strip() != '' else 'Matéria Não Cadastrada'
            
            # Para cada tipo de mídia, verificar se existe registro ou criar um vazio
            for tipo_midia in tipos_midia_padrao:
                # Se o tipo de mídia atual é o mesmo que detectamos no link web imagem
                # usar o link web imagem diretamente para Impresso, TV e Rádio
                if tipo_midia_detectado and tipo_midia == tipo_midia_detectado and link_web_imagem and tipo_midia in ['Impresso', 'TV', 'Rádio']:
                    link_especifico = link_web_imagem
                    logger.info(f"Usando link web imagem diretamente para {tipo_midia}: {link_especifico}")
                # Para Portal, priorizar usar o Link web - Texto
                elif tipo_midia == 'Portal' and link_web_texto:
                    link_especifico = link_web_texto
                    logger.info(f"Usando link web texto diretamente para Portal: {link_especifico}")
                else:
                    # Caso contrário, obter o link específico para este tipo de mídia
                    link_especifico = obter_link_por_tipo_midia(link_base, tipo_midia)
                    logger.info(f"Usando link processado para {tipo_midia}: {link_especifico}")
                
                registros_tipo = df_palavra[df_palavra['TIPO DE MÍDIA'] == tipo_midia]
                
                if len(registros_tipo) > 0:
                    # Usar o primeiro registro encontrado, mas com o link específico para este tipo de mídia
                    registro = registros_tipo.iloc[0].to_dict()
                    registro['LINK DA MATÉRIA CADASTRADA'] = link_especifico
                else:
                    # Criar um registro para este tipo de mídia
                    registro = {
                        'PALAVRAS-CHAVE': palavra,
                        'DATA DE CADASTRO': data_cadastro,
                        'TÍTULO DA MATÉRIA': titulo_materia,
                        'TIPO DE MÍDIA': tipo_midia,
                        'LINK DA MATÉRIA CADASTRADA': link_especifico
                    }
                
                # Adicionar o registro à lista
                registros_palavra.append(registro)
            
            # Adicionar os registros ao resultado
            resultado[palavra] = registros_palavra
        
        return {'status': 'sucesso', 'dados': resultado}
        
    except Exception as e:
        import traceback
        traceback_str = traceback.format_exc()
        logger.error(f"Erro ao processar planilha: {str(e)}\n{traceback_str}")
        return {'status': 'erro', 'mensagem': str(e), 'traceback': traceback_str}

def exportar_planilha_keywords(dados, caminho_saida):
    """
    Exporta os dados de palavras-chave processados para uma planilha Excel.
    Cada palavra-chave aparece apenas uma vez. Para cada palavra-chave, são gerados 
    quatro tipos de mídia (Portal, Impresso, TV, Rádio), cada um com seu formato específico de arquivo
    extraído da página da matéria.
    
    Args:
        dados: Dicionário com dados organizados por palavra-chave
        caminho_saida: Caminho onde a planilha será salva
        
    Returns:
        Status da operação
    """
    try:
        # Definir a ordem desejada das colunas conforme a imagem da planilha
        ordem_colunas = [
            'PALAVRAS-CHAVE',
            'DATA DE CADASTRO',
            'TÍTULO DA MATÉRIA',
            'TIPO DE MÍDIA',
            'LINK DA MATÉRIA CADASTRADA'
        ]
        
        # Definir a ordem dos tipos de mídia
        ordem_midia = ['Portal', 'Impresso', 'TV', 'Rádio']
        
        # Criar um DataFrame final para armazenar todos os dados
        final_df = pd.DataFrame(columns=ordem_colunas)
        
        # Ordenar as palavras-chave para melhor organização
        palavras_chave_ordenadas = sorted(dados.keys())
        
        # Processar cada palavra-chave
        for palavra in palavras_chave_ordenadas:
            registros = dados[palavra]
            df_palavra = pd.DataFrame(registros)
            
            # Encontrar o link base para esta palavra-chave
            link_base = ''
            for _, row in df_palavra.iterrows():
                if 'LINK DA MATÉRIA CADASTRADA' in row and row['LINK DA MATÉRIA CADASTRADA'] and str(row['LINK DA MATÉRIA CADASTRADA']).strip():
                    link_base = str(row['LINK DA MATÉRIA CADASTRADA']).strip()
                    break
            
            # Se não encontrou um link, usar um padrão
            if not link_base:
                link_base = f"https://braspub.com.br/materias/{palavra.replace(' ', '_').lower()}"
            
            # Criar registros ordenados para cada tipo de mídia
            registros_ordenados = []
            for tipo in ordem_midia:
                # Obter o link específico para este tipo de mídia
                link_especifico = obter_link_por_tipo_midia(link_base, tipo)
                
                # Filtrar por tipo de mídia
                filtro = df_palavra['TIPO DE MÍDIA'] == tipo
                if filtro.any():
                    # Obter o registro deste tipo de mídia
                    registro = df_palavra[filtro].iloc[0].to_dict()
                    # Definir o link específico para este tipo de mídia
                    registro['LINK DA MATÉRIA CADASTRADA'] = link_especifico
                else:
                    # Obter valores padrão da primeira linha
                    data_cadastro = df_palavra['DATA DE CADASTRO'].iloc[0] if 'DATA DE CADASTRO' in df_palavra.columns and len(df_palavra) > 0 else ''
                    titulo = df_palavra['TÍTULO DA MATÉRIA'].iloc[0] if 'TÍTULO DA MATÉRIA' in df_palavra.columns and len(df_palavra) > 0 and str(df_palavra['TÍTULO DA MATÉRIA'].iloc[0]).strip() != '' else 'Matéria Não Cadastrada'
                    
                    # Criar registro para este tipo de mídia
                    registro = {
                        'PALAVRAS-CHAVE': palavra,
                        'DATA DE CADASTRO': data_cadastro,
                        'TÍTULO DA MATÉRIA': titulo,
                        'TIPO DE MÍDIA': tipo,
                        'LINK DA MATÉRIA CADASTRADA': link_especifico
                    }
                
                # Adicionar à lista de registros ordenados
                registros_ordenados.append(registro)
            
            # Adicionar ao DataFrame final
            df_linhas = pd.DataFrame(registros_ordenados)
            final_df = pd.concat([final_df, df_linhas], ignore_index=True)
        
        # Criar um objeto ExcelWriter
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            # Escrever o DataFrame na planilha
            final_df.to_excel(writer, sheet_name='Palavras-Chave', index=False)
            
            # Ajustar a largura das colunas automaticamente
            worksheet = writer.sheets['Palavras-Chave']
            
            # Dicionário para armazenar a largura máxima de cada coluna
            max_width = {}
            
            # Inicializar o dicionário com os tamanhos dos cabeçalhos
            for idx, col in enumerate(final_df.columns):
                max_width[idx] = len(str(col)) + 2  # +2 para dar um pouco de espaço extra
            
            # Calcular a largura máxima para cada coluna baseada nos dados
            for idx, col in enumerate(final_df.columns):
                # Converter todos os valores para string e obter o comprimento
                column_width = max(
                    final_df[col].astype(str).map(len).max(),  # Maior valor dos dados
                    max_width[idx]  # Largura atual (cabeçalho)
                )
                
                # Limitar a uma largura máxima razoável (opcional)
                max_width[idx] = min(column_width + 2, 100)  # +2 para espaço e limite de 100
            
            # Aplicar as larguras às colunas
            for idx, width in max_width.items():
                col_letter = chr(65 + idx) if idx < 26 else chr(64 + idx // 26) + chr(65 + idx % 26)
                worksheet.column_dimensions[col_letter].width = width
    
        return {'status': 'sucesso', 'mensagem': f'Planilha de palavras-chave salva com sucesso em {caminho_saida}'}
        
    except Exception as e:
        return {'status': 'erro', 'mensagem': str(e)}

def extrair_keywords_da_pagina(url_base):
    """
    Extrai as palavras-chave de uma página HTML, buscando dentro de
    divs com a classe 'q-chip__content'.
    
    Args:
        url_base: URL da página
        
    Returns:
        Lista de palavras-chave encontradas
    """
    try:
        # Verificar se a URL é válida
        if not url_base or not url_base.startswith(('http://', 'https://')):
            logger.warning(f"URL inválida para extração de keywords: {url_base}")
            return []
        
        logger.info(f"Buscando keywords na URL: {url_base}")
        
        # Fazer requisição para a página
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url_base, headers=headers, timeout=15)
        if response.status_code != 200:
            logger.warning(f"Falha ao acessar URL: {url_base}, status: {response.status_code}")
            return []
        
        logger.info(f"Página acessada com sucesso. Extraindo keywords.")
        
        # Analisar o HTML da página
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Buscar as divs com a classe 'q-chip__content'
        keywords = []
        for div in soup.select('div.q-chip__content'):
            keyword = div.get_text().strip()
            if keyword:
                keywords.append(keyword)
                logger.info(f"Encontrada keyword: {keyword}")
        
        # Se não encontrou nas divs específicas, tentar outras abordagens
        if not keywords:
            # Tentar buscar em elementos com classes que possam conter palavras-chave
            for elem in soup.select('.tag, .keyword, .palavra-chave, .assunto'):
                keyword = elem.get_text().strip()
                if keyword:
                    keywords.append(keyword)
                    logger.info(f"Encontrada keyword em outro elemento: {keyword}")
            
            # Tentar buscar em meta tags
            meta_keywords = soup.find('meta', {'name': 'keywords'})
            if meta_keywords and meta_keywords.get('content'):
                content = meta_keywords['content']
                for keyword in content.split(','):
                    keyword = keyword.strip()
                    if keyword:
                        keywords.append(keyword)
                        logger.info(f"Encontrada keyword em meta tag: {keyword}")
        
        return keywords
        
    except Exception as e:
        logger.error(f"Erro ao extrair keywords: {str(e)}", exc_info=True)
        return []

# Modificar a função de detecção de mídia para considerar elementos específicos
def detectar_tipo_midia(url_base):
    """
    Detecta o tipo de mídia predominante na página.
    
    Args:
        url_base: URL da página
        
    Returns:
        String com o tipo de mídia ('Portal', 'Impresso', 'TV', 'Rádio')
    """
    try:
        # Verificar se a URL é válida
        if not url_base or not url_base.startswith(('http://', 'https://')):
            logger.warning(f"URL inválida para detecção de mídia: {url_base}")
            return 'Portal'  # Valor padrão
        
        logger.info(f"Detectando tipo de mídia na URL: {url_base}")
        
        # Fazer requisição para a página
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url_base, headers=headers, timeout=15)
        if response.status_code != 200:
            logger.warning(f"Falha ao acessar URL: {url_base}, status: {response.status_code}")
            return 'Portal'
        
        # Analisar o HTML da página
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Verificar elementos específicos para determinar o tipo de mídia
        
        # Verificar se há elementos de vídeo
        videos = soup.find_all(['video', 'iframe'])
        video_links = soup.find_all('a', href=lambda href: href and any(ext in href.lower() for ext in ['.mp4', 'youtube', 'vimeo']))
        if videos or video_links:
            logger.info(f"Detectado tipo de mídia: TV")
            return 'TV'
        
        # Verificar se há elementos de áudio
        audios = soup.find_all('audio')
        audio_links = soup.find_all('a', href=lambda href: href and '.mp3' in href.lower())
        if audios or audio_links:
            logger.info(f"Detectado tipo de mídia: Rádio")
            return 'Rádio'
        
        # Verificar se é uma página de conteúdo impresso
        # Verificar padrões típicos de conteúdo impresso
        if any(termo in response.text.lower() for termo in ['jornal impresso', 'versão impressa', 'edição impressa']):
            logger.info(f"Detectado tipo de mídia: Impresso")
            return 'Impresso'
        
        # Verificar a presença de muitas imagens (típico de conteúdo impresso)
        images = soup.find_all('img', src=lambda src: src and any(ext in src.lower() for ext in ['.jpg', '.jpeg', '.png']))
        if len(images) > 5:  # Se tiver muitas imagens, provavelmente é impresso
            logger.info(f"Detectado tipo de mídia: Impresso (muitas imagens)")
            return 'Impresso'
        
        # Se não detectou nenhum dos anteriores, considerar como Portal (padrão)
        logger.info(f"Tipo de mídia padrão: Portal")
        return 'Portal'
        
    except Exception as e:
        logger.error(f"Erro ao detectar tipo de mídia: {str(e)}", exc_info=True)
        return 'Portal'  # Valor padrão em caso de erro

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
            resultado = exportar_planilha_keywords(dados, caminho_saida)
            print(json.dumps(resultado))
            
        except Exception as e:
            print(json.dumps({'status': 'erro', 'mensagem': str(e)}))
            
    else:
        # Processar a planilha
        resultado = processar_planilha_keywords(caminho_arquivo)
        print(json.dumps(resultado, default=json_serial))

if __name__ == "__main__":
    main() 