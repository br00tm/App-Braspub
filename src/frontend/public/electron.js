const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const isDev = require('electron-is-dev');
const { spawn, exec } = require('child_process');
const fs = require('fs');
const os = require('os');
const axios = require('axios');
const FormData = require('form-data');
const url = require('url');

// Configuração da janela principal
let mainWindow;
let backendProcess;
const API_URL = isDev ? 'http://127.0.0.1:5000' : 'http://localhost:5000'; // Mudar para a URL da sua VPS em produção

// Configurar o servidor backend
function setupBackend() {
  if (isDev) {
    console.log('Executando em modo de desenvolvimento, não é necessário iniciar o backend.');
    return null;
  }

  try {
    const backendPath = path.join(process.resourcesPath, 'OrganizadorPlanilhas.exe');
    
    if (fs.existsSync(backendPath)) {
      console.log(`Iniciando backend: ${backendPath}`);
      return spawn(backendPath, [], {
        detached: false,
        stdio: 'ignore'
      });
    } else {
      console.error(`Backend não encontrado: ${backendPath}`);
      dialog.showErrorBox(
        'Erro ao iniciar aplicação',
        `O arquivo do backend não foi encontrado: ${backendPath}`
      );
      return null;
    }
  } catch (error) {
    console.error('Erro ao iniciar o backend:', error);
    return null;
  }
}

// Função para verificar se o backend está online
async function checkBackendStatus() {
  try {
    const response = await axios.get(`${API_URL}/api/status`);
    return response.data.status === 'online';
  } catch (error) {
    console.error('Backend não está respondendo:', error.message);
    return false;
  }
}

// Função para criar a janela principal
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    show: false, // Importante: não mostrar até estar pronto
    backgroundColor: '#f5f5f5', // Fundo para evitar flash branco
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
      preload: path.join(__dirname, 'preload.js')
    },
    icon: path.join(__dirname, isDev ? 'icon.ico' : '../build/icon.ico')
  });

  // Carregar URL
  const startUrl = isDev
    ? 'http://localhost:3000'
    : `file://${path.join(__dirname, '../build/index.html')}`;

  mainWindow.loadURL(startUrl);

  // Mostrar a janela quando estiver pronta para evitar flash branco
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    mainWindow.focus();
  });

  // Abrir DevTools em desenvolvimento
  if (isDev) {
    mainWindow.webContents.openDevTools();
  } else {
    // Remover menu em produção
    mainWindow.setMenu(null);
  }

  // Evento quando a janela é fechada
  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// Manipular eventos de ciclo de vida do app
app.on('ready', () => {
  // Iniciar o backend em produção
  if (!isDev) {
    backendProcess = setupBackend();
  }
  
  // Verificar status do backend e criar a janela principal
  const checkAndCreateWindow = async () => {
    try {
      // Em desenvolvimento o backend já deve estar rodando
      if (isDev) {
        createWindow();
        return;
      }
      
      // Em produção, verificar se o backend está online
      let attempts = 0;
      const maxAttempts = 10;
      
      const checkStatus = async () => {
        try {
          const isOnline = await checkBackendStatus();
          if (isOnline) {
            createWindow();
          } else {
            attempts++;
            if (attempts < maxAttempts) {
              console.log(`Backend ainda não está online. Tentativa ${attempts}/${maxAttempts}`);
              setTimeout(checkStatus, 1000);
            } else {
              console.error('Não foi possível conectar ao backend após várias tentativas');
              dialog.showErrorBox(
                'Erro de Conexão',
                'Não foi possível conectar ao servidor backend. A aplicação será encerrada.'
              );
              app.quit();
            }
          }
        } catch (error) {
          console.error('Erro ao verificar status do backend:', error);
          setTimeout(checkStatus, 1000);
        }
      };
      
      checkStatus();
    } catch (error) {
      console.error('Erro ao iniciar aplicação:', error);
      dialog.showErrorBox('Erro ao Iniciar', `Ocorreu um erro ao iniciar a aplicação: ${error.message}`);
    }
  };
  
  checkAndCreateWindow();
});

// Processar planilha Excel
ipcMain.handle('processar-planilha', async (event, filePath) => {
  try {
    console.log(`Iniciando processamento de planilha: ${filePath}`);
    
    // Se não recebemos um caminho válido (pode ser chamada de verificação)
    if (!filePath) {
      console.log('Verificação de conectividade com o backend');
      // Verificar se o backend está online
      const isOnline = await checkBackendStatus();
      if (!isOnline) {
        console.log('Backend não está respondendo na verificação de conexão');
        return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
      }
      console.log('Backend está online');
      return { status: 'sucesso', mensagem: 'Backend está online' };
    }
    
    // Verificar se o arquivo existe
    if (!fs.existsSync(filePath)) {
      console.error(`Arquivo não encontrado: ${filePath}`);
      return { status: 'erro', mensagem: `Arquivo não encontrado: ${filePath}` };
    }
    
    // Verificar se o backend está online
    const isOnline = await checkBackendStatus();
    if (!isOnline) {
      console.error('Backend não está respondendo');
      return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
    }
    
    console.log('Criando objeto FormData');
    // Criar um objeto FormData
    const formData = new FormData();
    formData.append('arquivo', fs.createReadStream(filePath));
    
    console.log('Enviando requisição para o backend');
    // Enviar para a API com timeout e retry
    try {
      const response = await axios.post(`${API_URL}/api/processar`, formData, {
        headers: {
          ...formData.getHeaders()
        },
        maxContentLength: Infinity,
        maxBodyLength: Infinity,
        timeout: 60000, // 60 segundos de timeout
        validateStatus: (status) => status < 500 // Aceitar status codes menores que 500
      });
      
      console.log(`Resposta recebida do backend: status ${response.status}`);
      
      // Verificar status da resposta
      if (response.status !== 200) {
        console.error(`Erro na resposta: ${response.status}`, response.data);
        return { 
          status: 'erro', 
          mensagem: response.data?.mensagem || `Erro no servidor: ${response.status}` 
        };
      }
      
      // Verificar se os dados da resposta são válidos
      if (!response.data || typeof response.data !== 'object') {
        console.error('Resposta do servidor inválida', response.data);
        return { 
          status: 'erro', 
          mensagem: 'Resposta do servidor inválida'
        };
      }
      
      console.log('Processamento concluído com sucesso');
      return response.data;
      
    } catch (requestError) {
      console.error('Erro na requisição:', requestError);
      
      // Tratar diferentes tipos de erros de requisição
      if (requestError.code === 'ECONNREFUSED') {
        return { status: 'erro', mensagem: 'Não foi possível conectar ao servidor backend. Verifique se o servidor está em execução.' };
      } 
      else if (requestError.code === 'ETIMEDOUT') {
        return { status: 'erro', mensagem: 'A conexão com o servidor expirou. O processamento pode ser muito longo ou o servidor está sobrecarregado.' };
      }
      else if (requestError.response) {
        // O servidor respondeu com um código de status fora do intervalo 2xx
        console.error('Erro na resposta do servidor:', {
          status: requestError.response.status,
          data: requestError.response.data
        });
        return { 
          status: 'erro', 
          mensagem: requestError.response.data?.mensagem || `Erro no servidor: ${requestError.response.status}`
        };
      }
      
      return { 
        status: 'erro', 
        mensagem: `Erro na comunicação com o servidor: ${requestError.message}` 
      };
    }
  } catch (error) {
    console.error('Erro geral ao processar planilha:', error);
    return { 
      status: 'erro', 
      mensagem: `Erro ao processar planilha: ${error.message}` 
    };
  }
});

// Processar planilha de palavras-chave
ipcMain.handle('processar-planilha-keywords', async (event, filePath) => {
  try {
    // Verificar se o backend está online
    const isOnline = await checkBackendStatus();
    if (!isOnline) {
      return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
    }

    // Criar um objeto FormData
    const formData = new FormData();
    formData.append('arquivo', fs.createReadStream(filePath));

    // Enviar para a API específica de palavras-chave
    const response = await axios.post(`${API_URL}/api/processar_keywords`, formData, {
      headers: {
        ...formData.getHeaders()
      },
      maxContentLength: Infinity,
      maxBodyLength: Infinity
    });

    return response.data;
  } catch (error) {
    console.error('Erro ao processar planilha de palavras-chave:', error);
    return { 
      status: 'erro', 
      mensagem: error.response ? error.response.data.mensagem : error.message 
    };
  }
});

// Exportar planilha organizada
ipcMain.handle('exportar-planilha', async (event, dados) => {
  try {
    console.log('Iniciando exportação de planilha');
    
    // Verificar se os dados foram fornecidos
    if (!dados || typeof dados !== 'object') {
      console.error('Dados inválidos para exportação');
      return { status: 'erro', mensagem: 'Dados inválidos para exportação' };
    }
    
    // Verificar estrutura dos dados para debug
    console.log(`Estrutura dos dados: ${typeof dados}`);
    console.log(`Chaves: ${Object.keys(dados).join(', ')}`);
    
    // Log detalhado dos dados para diagnóstico
    console.log('Detalhes dos dados recebidos:');
    for (const tipo in dados) {
      console.log(`- Tipo: ${tipo}`);
      if (Array.isArray(dados[tipo])) {
        console.log(`  - Quantidade: ${dados[tipo].length} registros`);
        if (dados[tipo].length > 0) {
          console.log(`  - Exemplo do primeiro registro:`);
          console.log(JSON.stringify(dados[tipo][0], null, 2).substring(0, 200) + '...');
        }
      } else {
        console.log(`  - PROBLEMA: Não é um array, é ${typeof dados[tipo]}`);
      }
    }
    
    for (const tipo in dados) {
      console.log(`Tipo: ${tipo}, registros: ${Array.isArray(dados[tipo]) ? dados[tipo].length : 'não é array'}`);
      
      // Verificar se cada tipo tem uma lista de registros
      if (!Array.isArray(dados[tipo])) {
        console.error(`Erro: dados para o tipo "${tipo}" não é um array`);
        console.log(`Convertendo dados para o tipo "${tipo}" em um formato válido...`);
        dados[tipo] = [];
      }
    }
    
    console.log('Solicitando local para salvar o arquivo');
    // Solicitar ao usuário onde salvar o arquivo
    const resultado = await dialog.showSaveDialog(mainWindow, {
      title: 'Salvar planilha organizada',
      defaultPath: path.join(os.homedir(), 'Desktop', 'Planilha_Organizada.xlsx'),
      filters: [
        { name: 'Arquivos Excel', extensions: ['xlsx'] }
      ]
    });
    
    if (resultado.canceled) {
      console.log('Exportação cancelada pelo usuário');
      return { status: 'cancelado' };
    }
    
    const caminho_saida = resultado.filePath;
    console.log(`Local de salvamento selecionado: ${caminho_saida}`);
    
    // Criar Excel diretamente no frontend sem usar o backend
    try {
      // Verificar se o módulo xlsx está disponível
      let XLSX;
      try {
        XLSX = require('xlsx');
        console.log('Módulo XLSX carregado com sucesso');
      } catch (moduleError) {
        console.error('Erro ao carregar módulo XLSX:', moduleError);
        
        // Tentar carregar o módulo de outros locais
        try {
          const xlsxPath = path.join(app.getAppPath(), 'node_modules', 'xlsx');
          console.log(`Tentando carregar xlsx de: ${xlsxPath}`);
          XLSX = require(xlsxPath);
          console.log('Módulo XLSX carregado com sucesso de caminho alternativo');
        } catch (secError) {
          console.error('Erro ao carregar módulo XLSX de caminho alternativo:', secError);
          return { 
            status: 'erro', 
            mensagem: `Não foi possível carregar o módulo XLSX. Erro: ${moduleError.message}` 
          };
        }
      }
      
      // Verificar se temos dados válidos
      let temDados = false;
      let tipoColunas = {};
      
      // Garantir que as colunas obrigatórias existam em todos os registros
      const colunasObrigatorias = [
        'Nome do Cliente',
        'Data de Inclusão',
        'Título da Matéria',
        'Link da Matéria',
        'Veículo',
        'Tipo de Mídia'
      ];
      
      for (const tipo in dados) {
        if (Array.isArray(dados[tipo]) && dados[tipo].length > 0) {
          temDados = true;
          
          // Verificar e corrigir estrutura dos itens
          const dadosCorrigidos = [];
          
          for (let i = 0; i < dados[tipo].length; i++) {
            const item = dados[tipo][i];
            // Verificar se o item é um objeto válido
            if (item && typeof item === 'object') {
              // Garantir que todos os campos obrigatórios existam
              const itemCompleto = { ...item };
              
              for (const coluna of colunasObrigatorias) {
                if (!itemCompleto[coluna]) {
                  itemCompleto[coluna] = coluna === 'Data de Inclusão' 
                    ? new Date().toISOString().split('T')[0]
                    : '';
                }
              }
              
              dadosCorrigidos.push(itemCompleto);
              
              // Coletar todas as colunas para cada tipo
              Object.keys(itemCompleto).forEach(coluna => {
                if (!tipoColunas[tipo]) tipoColunas[tipo] = new Set();
                tipoColunas[tipo].add(coluna);
              });
            } else {
              console.error(`Item inválido encontrado: ${JSON.stringify(item)}`);
              // Adicionar item vazio mas completo para manter a contagem
              const itemVazio = {};
              colunasObrigatorias.forEach(col => {
                itemVazio[col] = col === 'Data de Inclusão'
                  ? new Date().toISOString().split('T')[0]
                  : '';
              });
              dadosCorrigidos.push(itemVazio);
            }
          }
          
          // Substituir os dados originais pelos corrigidos
          dados[tipo] = dadosCorrigidos;
        }
      }
      
      if (!temDados) {
        console.error('Nenhum dado válido para exportar');
        return { status: 'erro', mensagem: 'Nenhum dado válido para exportar' };
      }
      
      // Ordenar colunas para cada tipo
      for (const tipo in tipoColunas) {
        console.log(`Colunas para ${tipo}: ${[...tipoColunas[tipo]].join(', ')}`);
      }
      
      console.log('Criando workbook Excel');
      // Criar um novo workbook
      const wb = XLSX.utils.book_new();
      
      // Garantir que todos os tipos de mídia padrão existam
      const tiposMidiaPadrao = ['Portal', 'Impresso', 'TV', 'Rádio'];
      for (const tipoMidia of tiposMidiaPadrao) {
        // Verificar se o tipo existe nos dados
        if (!dados[tipoMidia] || !Array.isArray(dados[tipoMidia])) {
          console.log(`Tipo ${tipoMidia} não existe ou não é array. Criando array vazio.`);
          dados[tipoMidia] = [];
        }
        
        // Se o array estiver vazio, adicionar um registro vazio para garantir que a aba seja criada
        if (dados[tipoMidia].length === 0) {
          console.log(`Adicionando registro vazio para o tipo: ${tipoMidia}`);
          const registroVazio = {};
          colunasObrigatorias.forEach(col => {
            if (col === 'Nome do Cliente') registroVazio[col] = 'Cliente';
            else if (col === 'Data de Inclusão') registroVazio[col] = new Date().toISOString().split('T')[0];
            else if (col === 'Título da Matéria') registroVazio[col] = `Nenhuma matéria de ${tipoMidia} encontrada`;
            else if (col === 'Link da Matéria') registroVazio[col] = 'N/A';
            else if (col === 'Veículo') registroVazio[col] = 'N/A';
            else if (col === 'Tipo de Mídia') registroVazio[col] = tipoMidia;
            else registroVazio[col] = '';
          });
          dados[tipoMidia].push(registroVazio);
        }
      }
      
      // Para cada tipo de mídia, adicionar uma worksheet
      for (const tipo in dados) {
        if (Array.isArray(dados[tipo]) && dados[tipo].length > 0) {
          console.log(`Processando tipo: ${tipo} com ${dados[tipo].length} registros`);
          
          try {
            // Ordenar as colunas para que as obrigatórias apareçam primeiro e na ordem correta
            const todascolunas = [...tipoColunas[tipo] || []];
            const colunasordenadas = [
              ...colunasObrigatorias.filter(col => todascolunas.includes(col)),
              ...todascolunas.filter(col => !colunasObrigatorias.includes(col))
            ];
            
            // Reorganizar os dados para seguir a ordem das colunas
            const dadosOrdenados = dados[tipo].map(item => {
              const novoItem = {};
              for (const col of colunasordenadas) {
                novoItem[col] = item[col] || '';
              }
              return novoItem;
            });
            
            // Criar uma worksheet a partir dos dados ordenados
            const ws = XLSX.utils.json_to_sheet(dadosOrdenados);
            
            // Limitar nome da aba a 31 caracteres e remover caracteres inválidos
            let nome_aba = String(tipo).slice(0, 31);
            nome_aba = nome_aba.replace(/[\\/*?:[\]]/g, '_');
            
            console.log(`Adicionando aba: ${nome_aba}`);
            // Adicionar a worksheet
            XLSX.utils.book_append_sheet(wb, ws, nome_aba);
          } catch (sheetError) {
            console.error(`Erro ao criar aba para ${tipo}:`, sheetError);
            
            // Tentar uma abordagem alternativa
            try {
              console.log('Tentando abordagem alternativa para criar planilha');
              // Criar planilha vazia e preencher manualmente
              const ws = XLSX.utils.aoa_to_sheet([
                // Cabeçalho com nomes das colunas
                colunasObrigatorias
              ]);
              
              // Limitar nome da aba a 31 caracteres e remover caracteres inválidos
              let nome_aba = String(tipo).slice(0, 31);
              nome_aba = nome_aba.replace(/[\\/*?:[\]]/g, '_');
              
              console.log(`Adicionando aba alternativa: ${nome_aba}`);
              // Adicionar a worksheet
              XLSX.utils.book_append_sheet(wb, ws, nome_aba);
            } catch (altSheetError) {
              console.error('Erro também na abordagem alternativa:', altSheetError);
            }
          }
        }
      }
      
      console.log(`Salvando workbook em: ${caminho_saida}`);
      
      // Verificar se o workbook tem pelo menos uma planilha
      if (wb.SheetNames.length === 0) {
        const ws = XLSX.utils.aoa_to_sheet([colunasObrigatorias, ['Não foi possível processar os dados', '', '', '', '', '']]);
        XLSX.utils.book_append_sheet(wb, ws, 'Erro');
      }
      
      // Escrever o arquivo
      XLSX.writeFile(wb, caminho_saida);
      console.log('Arquivo Excel salvo com sucesso');
      
      // Abrir a pasta contendo o arquivo
      shell.showItemInFolder(caminho_saida);
      
      return { 
        status: 'sucesso', 
        mensagem: `Planilha exportada com sucesso para: ${caminho_saida}` 
      };
    } catch (error) {
      console.error('Erro ao criar arquivo Excel:', error);
      return { 
        status: 'erro', 
        mensagem: `Erro ao criar arquivo Excel: ${error.message}` 
      };
    }
  } catch (error) {
    console.error('Erro geral na exportação:', error);
    return { 
      status: 'erro', 
      mensagem: `Erro ao exportar: ${error.message}` 
    };
  }
});

// Exportar planilha de palavras-chave organizada
ipcMain.handle('exportar-planilha-keywords', async (event, dados) => {
  try {
    console.log('Iniciando exportação de planilha de palavras-chave');
    
    // Verificar se os dados foram fornecidos
    if (!dados || typeof dados !== 'object') {
      console.error('Dados inválidos para exportação de palavras-chave');
      return { status: 'erro', mensagem: 'Dados não fornecidos ou inválidos' };
    }
    
    // Verificar estrutura dos dados para debug
    console.log(`Estrutura dos dados de palavras-chave: ${typeof dados}`);
    console.log(`Chaves: ${Object.keys(dados).join(', ')}`);
    
    // Solicitar ao usuário onde salvar o arquivo
    const resultado = await dialog.showSaveDialog(mainWindow, {
      title: 'Salvar planilha de palavras-chave organizada',
      defaultPath: path.join(os.homedir(), 'Desktop', 'Palavras_Chave_Organizadas.xlsx'),
      filters: [
        { name: 'Arquivos Excel', extensions: ['xlsx'] }
      ]
    });
    
    if (resultado.canceled) {
      console.log('Exportação cancelada pelo usuário');
      return { status: 'cancelado' };
    }
    
    const caminho_saida = resultado.filePath;
    console.log(`Local de salvamento selecionado: ${caminho_saida}`);
    
    // Usar o método alternativo com XLSX que funciona para o tipo mídia
    try {
      // Verificar se o módulo xlsx está disponível
      let XLSX;
      try {
        XLSX = require('xlsx');
        console.log('Módulo XLSX carregado com sucesso para palavras-chave');
      } catch (moduleError) {
        console.error('Erro ao carregar módulo XLSX para palavras-chave:', moduleError);
        
        // Tentar carregar o módulo de outros locais
        try {
          const xlsxPath = path.join(app.getAppPath(), 'node_modules', 'xlsx');
          console.log(`Tentando carregar xlsx de: ${xlsxPath}`);
          XLSX = require(xlsxPath);
          console.log('Módulo XLSX carregado com sucesso de caminho alternativo');
        } catch (secError) {
          console.error('Erro ao carregar módulo XLSX de caminho alternativo:', secError);
          return { 
            status: 'erro', 
            mensagem: `Não foi possível carregar o módulo XLSX. Erro: ${moduleError.message}` 
          };
        }
      }
      
      // Criar planilha com todas as palavras-chave
      const colunas_necessarias = [
        'PALAVRAS-CHAVE', 'DATA DE CADASTRO', 'TÍTULO DA MATÉRIA',
        'TIPO DE MÍDIA', 'LINK DA MATÉRIA CADASTRADA'
      ];
      
      console.log('Extraindo registros de todas as palavras-chave');
      
      // Extrair todos os registros
      let registros_totais = [];
      let temDados = false;
      
      // Garantir que cada palavra-chave tenha todos os tipos de mídia
      const tiposMidiaPadrao = ['Portal', 'Impresso', 'TV', 'Rádio'];
      
      for (const palavra in dados) {
        console.log(`Processando palavra: ${palavra}`);
        if (Array.isArray(dados[palavra])) {
          temDados = true;
          
          // Verificar quais tipos de mídia existem para essa palavra
          const tiposExistentes = new Set(dados[palavra].map(reg => reg['TIPO DE MÍDIA']));
          
          // Adicionar os registros desta palavra-chave
          for (const reg of dados[palavra]) {
            // Garantir que todos os campos necessários existam
            const registro_completo = {...reg};
            colunas_necessarias.forEach(col => {
              if (!registro_completo[col]) registro_completo[col] = "";
            });
            
            registros_totais.push(registro_completo);
          }
          
          // Adicionar registros para tipos de mídia que faltam
          for (const tipo of tiposMidiaPadrao) {
            if (!tiposExistentes.has(tipo)) {
              console.log(`Adicionando registro vazio para palavra "${palavra}" com tipo "${tipo}"`);
              const registro_vazio = {
                'PALAVRAS-CHAVE': palavra,
                'DATA DE CADASTRO': new Date().toISOString().split('T')[0],
                'TÍTULO DA MATÉRIA': `Sem matéria para ${palavra}`,
                'TIPO DE MÍDIA': tipo,
                'LINK DA MATÉRIA CADASTRADA': 'N/A'
              };
              registros_totais.push(registro_vazio);
            }
          }
        } else {
          console.warn(`Palavra "${palavra}" não tem registros válidos`);
        }
      }
      
      console.log(`Total de registros extraídos: ${registros_totais.length}`);
      
      if (!temDados || registros_totais.length === 0) {
        console.error('Nenhum dado válido para exportar');
        return { status: 'erro', mensagem: 'Nenhum dado válido encontrado nas palavras-chave' };
      }
      
      // Criar workbook e adicionar planilha
      console.log('Criando workbook Excel para palavras-chave');
      const wb = XLSX.utils.book_new();
      
      // Adicionar todos os registros em uma única planilha
      try {
        // Garantir que as colunas apareçam na ordem correta
        const dadosOrdenados = registros_totais.map(registro => {
          // Criar um novo objeto com as colunas na ordem desejada
          const registroOrdenado = {};
          
          colunas_necessarias.forEach(coluna => {
            registroOrdenado[coluna] = registro[coluna] || '';
          });
          return registroOrdenado;
        });
        
        // Criar planilha com dados ordenados
        const ws = XLSX.utils.json_to_sheet(dadosOrdenados);
        XLSX.utils.book_append_sheet(wb, ws, 'Palavras-Chave');
        
        console.log(`Salvando workbook em: ${caminho_saida}`);
        XLSX.writeFile(wb, caminho_saida);
        console.log('Arquivo Excel de palavras-chave salvo com sucesso');
        
        // Abrir a pasta contendo o arquivo
        shell.showItemInFolder(caminho_saida);
        
        return { 
          status: 'sucesso', 
          mensagem: `Planilha de palavras-chave exportada com sucesso para: ${caminho_saida}` 
        };
      } catch (error) {
        console.error('Erro ao criar planilha de palavras-chave:', error);
        return { 
          status: 'erro', 
          mensagem: `Erro ao criar planilha de palavras-chave: ${error.message}` 
        };
      }
    } catch (error) {
      console.error('Erro ao criar arquivo Excel de palavras-chave:', error);
      return { 
        status: 'erro', 
        mensagem: `Erro ao criar arquivo Excel de palavras-chave: ${error.message}` 
      };
    }
  } catch (error) {
    console.error('Erro geral na exportação de palavras-chave:', error);
    return { 
      status: 'erro', 
      mensagem: `Erro ao exportar palavras-chave: ${error.message}` 
    };
  }
});

// Processar planilha para download
ipcMain.handle('processar-planilha-download', async (event, filePath) => {
  try {
    console.log(`Processando planilha para download: ${filePath}`);
    
    // Verificar se o backend está online
    const isOnline = await checkBackendStatus();
    if (!isOnline) {
      console.error('Backend não está respondendo');
      return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
    }
    
    console.log('Criando objeto FormData');
    // Criar um objeto FormData
    const formData = new FormData();
    formData.append('arquivo', fs.createReadStream(filePath));
    
    console.log('Enviando requisição para o backend');
    // Enviar para a API com timeout e retry
    try {
      const response = await axios.post(`${API_URL}/api/processar_planilha_download`, formData, {
        headers: {
          ...formData.getHeaders()
        },
        maxContentLength: Infinity,
        maxBodyLength: Infinity,
        timeout: 60000 // 60 segundos de timeout
      });
      
      console.log(`Resposta recebida do backend: status ${response.status}`);
      return response.data;
      
    } catch (requestError) {
      console.error('Erro na requisição:', requestError);
      
      if (requestError.response) {
        return { 
          status: 'erro', 
          mensagem: requestError.response.data?.mensagem || `Erro no servidor: ${requestError.response.status}`
        };
      }
      
      return { 
        status: 'erro', 
        mensagem: `Erro na comunicação com o servidor: ${requestError.message}` 
      };
    }
  } catch (error) {
    console.error('Erro geral ao processar planilha para download:', error);
    return { 
      status: 'erro', 
      mensagem: `Erro ao processar planilha para download: ${error.message}` 
    };
  }
});

// Processar planilha para download com nome diferente (mantendo compatibilidade)
ipcMain.handle('processar_planilha_download', async (event, filePath) => {
  try {
    console.log(`Processando planilha para download (via underscore): ${filePath}`);
    
    // Verificar se o backend está online
    const isOnline = await checkBackendStatus();
    if (!isOnline) {
      console.error('Backend não está respondendo');
      return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
    }
    
    console.log('Criando objeto FormData');
    // Criar um objeto FormData
    const formData = new FormData();
    formData.append('arquivo', fs.createReadStream(filePath));
    
    console.log('Enviando requisição para o backend');
    // Enviar para a API com timeout e retry
    try {
      const response = await axios.post(`${API_URL}/api/processar_planilha_download`, formData, {
        headers: {
          ...formData.getHeaders()
        },
        maxContentLength: Infinity,
        maxBodyLength: Infinity,
        timeout: 60000 // 60 segundos de timeout
      });
      
      console.log(`Resposta recebida do backend: status ${response.status}`);
      return response.data;
      
    } catch (requestError) {
      console.error('Erro na requisição:', requestError);
      
      if (requestError.response) {
        return { 
          status: 'erro', 
          mensagem: requestError.response.data?.mensagem || `Erro no servidor: ${requestError.response.status}`
        };
      }
      
      return { 
        status: 'erro', 
        mensagem: `Erro na comunicação com o servidor: ${requestError.message}` 
      };
    }
  } catch (error) {
    console.error('Erro geral ao processar planilha para download:', error);
    return { 
      status: 'erro', 
      mensagem: `Erro ao processar planilha para download: ${error.message}` 
    };
  }
});

// Baixar arquivos
ipcMain.handle('baixar-arquivos', async (event, urls) => {
  try {
    console.log(`Iniciando download de ${urls.length} arquivos`);
    
    // Verificar se o backend está online
    const isOnline = await checkBackendStatus();
    if (!isOnline) {
      console.error('Backend não está respondendo');
      return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
    }
    
    // Solicitar ao usuário onde salvar os arquivos
    const resultado = await dialog.showOpenDialog(mainWindow, {
      title: 'Selecionar pasta para salvar os arquivos',
      properties: ['openDirectory']
    });
    
    if (resultado.canceled) {
      console.log('Download cancelado pelo usuário');
      return { status: 'cancelado' };
    }
    
    const pastaDestino = resultado.filePaths[0];
    console.log(`Pasta de destino selecionada: ${pastaDestino}`);
    
    // Enviar requisição para API de download
    const response = await axios.post(`${API_URL}/api/baixar_arquivos`, {
      urls: urls,
      pasta_destino: pastaDestino
    });
    
    console.log('Resposta do download recebida');
    
    // Abrir a pasta de destino
    shell.openPath(pastaDestino);
    
    return response.data;
    
  } catch (error) {
    console.error('Erro ao baixar arquivos:', error);
    return { 
      status: 'erro', 
      mensagem: `Erro ao baixar arquivos: ${error.message}` 
    };
  }
});