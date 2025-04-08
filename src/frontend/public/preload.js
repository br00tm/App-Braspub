const { contextBridge, ipcRenderer } = require('electron');

// Expor API segura para o navegador
contextBridge.exposeInMainWorld(
  'electron',
  {
    ipcRenderer: {
      // Apenas as funções que precisamos expor para nosso aplicativo
      invoke: (channel, ...args) => {
        // Lista de canais permitidos que podemos invocar
        const validChannels = [
          'processar-planilha',
          'processar-planilha-keywords',
          'exportar-planilha',
          'exportar-planilha-keywords',
          'processar-planilha-download',
          'baixar-arquivos'
        ];
        
        if (validChannels.includes(channel)) {
          return ipcRenderer.invoke(channel, ...args);
        }
        
        return Promise.reject(new Error(`Canal ipcRenderer não permitido: ${channel}`));
      }
    },
    // Adicione qualquer outra funcionalidade necessária do Electron aqui
    isElectron: true
  }
);

// Adicionar variáveis úteis
window.isElectron = true;

// Injetar informações de funcionamento
console.log('Preload script loaded successfully'); 