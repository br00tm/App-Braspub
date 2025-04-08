# Organizador de Planilhas BrasPub

Um aplicativo desktop moderno para processar e organizar planilhas Excel por tipo de mídia. Desenvolvido com React, Electron e Python.

## Características

- Interface moderna e intuitiva
- Processamento rápido de planilhas Excel
- Organização automática por tipo de mídia
- Exportação para Excel com abas separadas
- Estatísticas detalhadas do processamento
- Compatível com Windows

## Requisitos

- Python 3.7 ou superior
- Node.js 14 ou superior
- Windows 10 ou superior

## Estrutura do Projeto

```
├── compilar.bat           # Script para compilar o aplicativo
├── executar_dev.bat       # Script para executar em modo de desenvolvimento
├── venv/                  # Ambiente virtual Python (criado automaticamente)
├── src/
│   ├── backend/           # Backend Python
│   │   ├── api.py         # API Flask 
│   │   ├── organizador.py # Processador de planilhas
│   │   └── requirements.txt
│   └── frontend/          # Frontend React + Electron
│       ├── public/        # Recursos públicos
│       │   ├── electron.js # Script principal do Electron
│       │   └── index.html
│       ├── src/           # Código fonte React
│       │   ├── App.js     # Componente principal
│       │   ├── App.css    # Estilos CSS
│       │   ├── index.js   # Ponto de entrada
│       │   └── index.css  # Estilos globais
│       └── package.json   # Configuração npm
```

## Instruções de Uso

### Executar em Modo de Desenvolvimento

1. Clone o repositório
2. Execute o script `executar_dev.bat` para iniciar a aplicação em modo de desenvolvimento
   - O script criará automaticamente um ambiente virtual Python (venv) se não existir
   - Instalará as dependências Python no ambiente virtual
   - Instalará as dependências do Node.js
   - Iniciará o backend e o frontend
3. A aplicação será iniciada com o backend Python e o frontend React+Electron

### Compilar o Aplicativo

1. Execute o script `compilar.bat` para compilar o aplicativo
   - O script criará automaticamente um ambiente virtual Python (venv) se não existir
   - Instalará as dependências Python e PyInstaller no ambiente virtual
   - Compilará o backend em um executável
   - Compilará o frontend com Electron
2. O instalador será gerado na pasta `dist/`
3. Execute o instalador para instalar o aplicativo no seu sistema

## Como Funciona

1. **Selecione uma planilha Excel**: A planilha deve conter uma coluna chamada "TIPO DE MÍDIA"
2. **Clique em Processar**: A aplicação processará a planilha, identificando os diferentes tipos de mídia
3. **Escolha onde salvar**: Selecione o local para salvar a planilha processada
4. **Pronto!**: A planilha será exportada com uma aba para cada tipo de mídia encontrado

## Detalhes Técnicos

### Backend (Python)

- **Flask**: API RESTful para comunicação com o frontend
- **Pandas**: Processamento de dados tabulares
- **OpenPyXL**: Manipulação de arquivos Excel
- **PyInstaller**: Compilação para executável
- **Ambiente Virtual**: Isolamento de dependências Python

### Frontend (React + Electron)

- **React**: Interface de usuário reativa
- **Electron**: Framework para aplicações desktop
- **Bootstrap**: Componentes de UI
- **Electron Builder**: Empacotamento e distribuição

## Solução de Problemas

### O aplicativo não inicia

- Verifique se o Python 3.7+ está instalado e configurado no PATH
- Verifique se o Node.js 14+ está instalado e configurado no PATH
- Verifique se você tem permissões para criar um ambiente virtual
- Reinstale o aplicativo se necessário

### Erro ao processar planilha

- Verifique se a planilha contém uma coluna chamada "TIPO DE MÍDIA"
- Verifique se o formato da planilha é .xlsx ou .xls
- Tente salvar a planilha em um formato mais recente se estiver usando um formato antigo

## Licença

Este software é propriedade da BrasPub. Todos os direitos reservados.
