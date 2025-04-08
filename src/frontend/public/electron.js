const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const isDev = require('electron-is-dev');
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const axios = require('axios');
const FormData = require('form-data');

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
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, isDev ? 'icon.ico' : '../build/icon.ico')
  });

  // Carregar URL
  const startUrl = isDev
    ? 'http://localhost:3000'
    : `file://${path.join(__dirname, '../build/index.html')}`;

  mainWindow.loadURL(startUrl);

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

// Processar planilha Excel
ipcMain.handle('processar-planilha', async (event, filePath) => {
  try {
    // Verificar se o backend está online
    const isOnline = await checkBackendStatus();
    if (!isOnline) {
      return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
    }

    // Criar um objeto FormData
    const formData = new FormData();
    formData.append('arquivo', fs.createReadStream(filePath));

    // Enviar para a API
    const response = await axios.post(`${API_URL}/api/processar`, formData, {
      headers: {
        ...formData.getHeaders()
      },
      maxContentLength: Infinity,
      maxBodyLength: Infinity
    });

    return response.data;
  } catch (error) {
    console.error('Erro ao processar planilha:', error);
    return { 
      status: 'erro', 
      mensagem: error.response ? error.response.data.mensagem : error.message 
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
    // Verificar se o backend está online
    const isOnline = await checkBackendStatus();
    if (!isOnline) {
      return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
    }

    // Solicitar ao usuário onde salvar o arquivo
    const resultado = await dialog.showSaveDialog(mainWindow, {
      title: 'Salvar planilha organizada',
      defaultPath: path.join(os.homedir(), 'Desktop', 'Planilha_Organizada.xlsx'),
      filters: [
        { name: 'Arquivos Excel', extensions: ['xlsx'] }
      ]
    });
    
    if (resultado.canceled) {
      return { status: 'cancelado' };
    }
    
    const caminho_saida = resultado.filePath;
    
    // Solicitar a exportação
    const response = await axios.post(
      `${API_URL}/api/exportar`, 
      { dados },
      { responseType: 'arraybuffer' }
    );

    // Salvar o arquivo recebido
    fs.writeFileSync(caminho_saida, Buffer.from(response.data));
    
    // Abrir o diretório do arquivo
    shell.showItemInFolder(caminho_saida);
    
    return { 
      status: 'sucesso', 
      mensagem: `Planilha exportada com sucesso para: ${caminho_saida}` 
    };
  } catch (error) {
    console.error('Erro ao exportar planilha:', error);
    return { 
      status: 'erro', 
      mensagem: error.response ? error.response.data.mensagem : error.message 
    };
  }
});

// Exportar planilha de palavras-chave organizada
ipcMain.handle('exportar-planilha-keywords', async (event, dados) => {
  try {
    // Verificar se o backend está online
    const isOnline = await checkBackendStatus();
    if (!isOnline) {
      return { status: 'erro', mensagem: 'Backend não está respondendo. Verifique a conexão.' };
    }

    // Solicitar ao usuário onde salvar o arquivo
    const resultado = await dialog.showSaveDialog(mainWindow, {
      title: 'Salvar planilha de palavras-chave organizada',
      defaultPath: path.join(os.homedir(), 'Desktop', 'Palavras_Chave_Organizadas.xlsx'),
      filters: [
        { name: 'Arquivos Excel', extensions: ['xlsx'] }
      ]
    });
    
    if (resultado.canceled) {
      return { status: 'cancelado' };
    }
    
    const caminho_saida = resultado.filePath;
    
    // Solicitar a exportação usando o endpoint específico para palavras-chave
    const response = await axios.post(
      `${API_URL}/api/exportar_keywords`, 
      { dados },
      { responseType: 'arraybuffer' }
    );

    // Salvar o arquivo recebido
    fs.writeFileSync(caminho_saida, Buffer.from(response.data));
    
    // Abrir o diretório do arquivo
    shell.showItemInFolder(caminho_saida);
    
    return { 
      status: 'sucesso', 
      mensagem: `Planilha de palavras-chave exportada com sucesso para: ${caminho_saida}` 
    };
  } catch (error) {
    console.error('Erro ao exportar planilha de palavras-chave:', error);
    return { 
      status: 'erro', 
      mensagem: error.response ? error.response.data.mensagem : error.message 
    };
  }
});

// Iniciar o app quando o Electron estiver pronto
app.on('ready', () => {
  // Iniciar o backend em produção
  if (!isDev) {
    backendProcess = setupBackend();
  }
  
  // Criar a janela principal após um breve intervalo para dar tempo ao backend iniciar
  setTimeout(createWindow, 1000);
});

// Sair quando todas as janelas estiverem fechadas
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    // Finalizar o processo do backend se ele existir
    if (backendProcess) {
      try {
        process.kill(backendProcess.pid);
      } catch (error) {
        console.error('Erro ao finalizar o backend:', error);
      }
    }
    
    app.quit();
  }
});

// Recriar a janela no macOS
app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
  }
}); 