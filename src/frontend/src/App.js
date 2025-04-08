import React, { useState, useEffect } from 'react';
import 'bootstrap/dist/css/bootstrap.min.css';
import {
  Container,
  Row,
  Col,
  Form,
  Button,
  Alert,
  Spinner,
  Card,
  ProgressBar,
  Navbar,
  Nav,
  ButtonGroup,
  ToggleButton
} from 'react-bootstrap';
import './App.css';

// Acessar API do Electron
const ipcRenderer = (() => {
  // Verifica se estamos em um ambiente Electron
  if (window.electron) {
    return window.electron.ipcRenderer;
  } else if (typeof window !== 'undefined' && window.require) {
    return window.require('electron').ipcRenderer;
  } else {
    // Fallback para ambiente de desenvolvimento web
    console.warn('Executando fora do ambiente Electron - algumas funcionalidades não estarão disponíveis');
    // Mock de ipcRenderer para desenvolvimento
    return {
      invoke: async (channel, ...args) => {
        console.log(`Mock ipcRenderer.invoke chamado: ${channel}`, args);
        // Retornar valores simulados para testes
        if (channel === 'processar-planilha' || channel === 'processar-planilha-keywords') {
          return { status: 'simulado', mensagem: 'Executando em ambiente web - simulação' };
        }
        return null;
      }
    };
  }
})();

function App() {
  // Estados gerais
  const [modoAtual, setModoAtual] = useState('tipoMidia'); // 'tipoMidia', 'palavrasChave', 'baixarArquivos'
  const [loading, setLoading] = useState(false);
  const [processStep, setProcessStep] = useState('');
  const [processingProgress, setProcessingProgress] = useState(0);
  
  // Estados do modo Tipo de Mídia
  const [fileTipoMidia, setFileTipoMidia] = useState(null);
  const [errorTipoMidia, setErrorTipoMidia] = useState(null);
  const [successTipoMidia, setSuccessTipoMidia] = useState(null);
  const [processedDataTipoMidia, setProcessedDataTipoMidia] = useState(null);
  
  // Estados do modo Palavras-chave
  const [filePalavrasChave, setFilePalavrasChave] = useState(null);
  const [errorPalavrasChave, setErrorPalavrasChave] = useState(null);
  const [successPalavrasChave, setSuccessPalavrasChave] = useState(null);
  const [processedDataPalavrasChave, setProcessedDataPalavrasChave] = useState(null);
  
  // Estados do modo Baixar Arquivos
  const [fileBaixarArquivos, setFileBaixarArquivos] = useState(null);
  const [errorBaixarArquivos, setErrorBaixarArquivos] = useState(null);
  const [successBaixarArquivos, setSuccessBaixarArquivos] = useState(null);
  const [processedDataBaixarArquivos, setProcessedDataBaixarArquivos] = useState(null);
  const [downloadStatus, setDownloadStatus] = useState({}); // Para rastrear status de downloads
  
  // Verificar se estamos no modo de palavras-chave
  const modoKeywords = modoAtual === 'palavrasChave';
  // Verificar se estamos no modo de baixar arquivos
  const modoBaixarArquivos = modoAtual === 'baixarArquivos';
  
  // Estados atuais baseados no modo selecionado
  const file = modoBaixarArquivos 
    ? fileBaixarArquivos 
    : (modoKeywords ? filePalavrasChave : fileTipoMidia);
    
  const setFile = modoBaixarArquivos 
    ? setFileBaixarArquivos 
    : (modoKeywords ? setFilePalavrasChave : setFileTipoMidia);
    
  const error = modoBaixarArquivos 
    ? errorBaixarArquivos 
    : (modoKeywords ? errorPalavrasChave : errorTipoMidia);
    
  const setError = modoBaixarArquivos 
    ? setErrorBaixarArquivos 
    : (modoKeywords ? setErrorPalavrasChave : setErrorTipoMidia);
    
  const success = modoBaixarArquivos 
    ? successBaixarArquivos 
    : (modoKeywords ? successPalavrasChave : successTipoMidia);
    
  const setSuccess = modoBaixarArquivos 
    ? setSuccessBaixarArquivos 
    : (modoKeywords ? setSuccessPalavrasChave : setSuccessTipoMidia);
    
  const processedData = modoBaixarArquivos 
    ? processedDataBaixarArquivos 
    : (modoKeywords ? processedDataPalavrasChave : processedDataTipoMidia);
    
  const setProcessedData = modoBaixarArquivos 
    ? setProcessedDataBaixarArquivos 
    : (modoKeywords ? setProcessedDataPalavrasChave : setProcessedDataTipoMidia);

  // Verificar conexão com o backend a cada 10 segundos
  useEffect(() => {
    const checkConnection = async () => {
      try {
        await ipcRenderer.invoke('processar-planilha', null);
        // Se não houver erro, estamos conectados
        setError(null);
      } catch (err) {
        if (err.message && err.message.includes('Backend não está respondendo')) {
          setError('Sem conexão com o servidor. Verifique se o servidor está online.');
        }
      }
    };

    // Verificar inicialmente
    checkConnection();
    
    // Configurar verificação periódica
    const interval = setInterval(checkConnection, 10000);
    
    // Limpar intervalo ao desmontar
    return () => clearInterval(interval);
  }, []);

  // Adicionar um useEffect para garantir reset de loading ao mudar de modo
  useEffect(() => {
    // Reset completo do estado quando mudar de modo
    setLoading(false);
    setProcessingProgress(0);
    setProcessStep('');
  }, [modoAtual]);

  // Manipular seleção de arquivo
  const handleFileSelect = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setError(null);
      setSuccess(null);
      setProcessedData(null);
    }
  };

  // Adicionar função para remover o arquivo selecionado
  const handleRemoveFile = () => {
    setFile(null);
    // Resetar o input file
    const fileInput = document.getElementById('formFile');
    if (fileInput) fileInput.value = '';
  };

  // Processar planilha e exportar diretamente
  const processAndExportExcel = async () => {
    if (!file) {
      setError('Por favor, selecione um arquivo Excel para processar.');
      return;
    }

    setLoading(true);
    setError(null);
    setSuccess(null);
    setProcessStep('processando');
    setProcessingProgress(10);

    try {
      // Primeiro, processar a planilha para obter os dados organizados
      setProcessStep(modoKeywords ? 'Processando planilha de palavras-chave...' : 'Processando planilha...');
      setProcessingProgress(30);
      
      // Escolher o método de processamento baseado no modo
      const processMethod = modoKeywords ? 'processar-planilha-keywords' : 'processar-planilha';
      console.log(`Chamando método: ${processMethod} para arquivo: ${file.path}`);
      const result = await ipcRenderer.invoke(processMethod, file.path);
      
      if (result && result.status === 'sucesso') {
        console.log('Processamento bem-sucedido, dados:', result.dados);
        setProcessedData(result.dados);
        setProcessingProgress(60);
        
        // Verificar estrutura dos dados para diagnóstico
        console.log('Verificando estrutura dos dados processados:');
        if (result.dados && typeof result.dados === 'object') {
          for (const key in result.dados) {
            console.log(`- ${key}: ${Array.isArray(result.dados[key]) ? result.dados[key].length + ' itens' : typeof result.dados[key]}`);
            if (Array.isArray(result.dados[key]) && result.dados[key].length > 0) {
              console.log(`  Exemplo de item: ${JSON.stringify(result.dados[key][0]).substring(0, 100)}...`);
            }
          }
        }
        
        // Imediatamente iniciar o processo de exportação
        setProcessStep(modoKeywords ? 'Exportando planilha de palavras-chave...' : 'Exportando planilha...');
        setProcessingProgress(80);
        
        // Escolher o método de exportação baseado no modo
        const exportMethod = modoKeywords ? 'exportar-planilha-keywords' : 'exportar-planilha';
        console.log(`Chamando método de exportação: ${exportMethod}`);
        
        // Adicionar timeout mais longo para exportação
        const exportPromise = ipcRenderer.invoke(exportMethod, result.dados);
        
        // Timeout para monitorar se a exportação está demorando muito
        const timeoutPromise = new Promise((_, reject) => {
          setTimeout(() => reject(new Error('Exportação demorou muito tempo')), 60000);
        });
        
        // Aguardar exportação ou timeout
        try {
          const exportResult = await Promise.race([exportPromise, timeoutPromise]);
          setProcessingProgress(100);
          
          console.log('Resultado da exportação:', exportResult);
          
          if (exportResult && exportResult.status === 'sucesso') {
            setSuccess(`Planilha ${modoKeywords ? 'de palavras-chave ' : ''}processada e exportada com sucesso! ${exportResult.mensagem || ''}`);
          } else if (exportResult && exportResult.status === 'cancelado') {
            setError('Exportação cancelada pelo usuário.');
          } else {
            setError(`Erro ao exportar a planilha: ${exportResult?.mensagem || 'Resposta inválida do servidor'}`);
          }
        } catch (timeoutErr) {
          console.error('Timeout na exportação:', timeoutErr);
          setError('A exportação demorou muito tempo. Tente novamente ou use um arquivo menor.');
        }
        
      } else {
        setError(`Erro ao processar a planilha: ${result?.mensagem || 'Resposta inválida do servidor'}`);
      }
    } catch (err) {
      console.error('Erro ao processar/exportar:', err);
      setError(`Erro: ${err.message || 'Ocorreu um erro inesperado'}`);
    } finally {
      setLoading(false);
      setProcessStep('');
    }
  };

  // Adicionar função para baixar arquivos
  const processAndDownloadFiles = async () => {
    if (!file) {
      setError('Por favor, selecione um arquivo Excel para processar.');
      return;
    }

    setLoading(true);
    setError(null);
    setSuccess(null);
    setProcessStep('processando');
    setProcessingProgress(10);

    try {
      // Primeiro, processar a planilha para obter os dados organizados
      setProcessStep('Analisando planilha para download de arquivos...');
      setProcessingProgress(20);
      
      // Chamar o método de processamento para download
      console.log(`Preparando para download de arquivos da planilha: ${file.path}`);
      const result = await ipcRenderer.invoke('processar-planilha-download', file.path);
      
      if (result && result.status === 'sucesso') {
        console.log('Análise bem-sucedida, links encontrados:', result.dados);
        setProcessedData(result.dados);
        setProcessingProgress(40);
        
        // Iniciar processo de download
        setProcessStep('Baixando arquivos...');
        
        // Chamar a função de download
        const downloadResult = await ipcRenderer.invoke('baixar-arquivos', result.dados);
        
        setProcessingProgress(100);
        
        if (downloadResult && downloadResult.status === 'sucesso') {
          setSuccess(`Arquivos baixados com sucesso! ${downloadResult.mensagem || ''}`);
          
          // Atualizar status de download
          if (downloadResult.detalhes) {
            setDownloadStatus(downloadResult.detalhes);
          }
        } else {
          setError(`Erro ao baixar arquivos: ${downloadResult?.mensagem || 'Resposta inválida do servidor'}`);
        }
      } else {
        setError(`Erro ao processar a planilha: ${result?.mensagem || 'Resposta inválida do servidor'}`);
      }
    } catch (err) {
      console.error('Erro ao processar/baixar:', err);
      setError(`Erro: ${err.message || 'Ocorreu um erro inesperado'}`);
    } finally {
      setLoading(false);
      setProcessStep('');
    }
  };

  // Estatísticas da planilha processada
  const renderProcessedStats = () => {
    if (!processedData) return null;
    
    const totalItems = Object.keys(processedData).length;
    const totalRegistros = Object.values(processedData).reduce(
      (sum, items) => sum + items.length, 0
    );
    
    return (
      <Card className="stat-card mb-4">
        <Card.Body>
          <Card.Title>Estatísticas do Processamento</Card.Title>
          <div className="stats-grid">
            <div className="stat-item">
              <div className="stat-value">{totalItems}</div>
              <div className="stat-label">{modoKeywords ? 'Palavras-chave' : 'Tipos de mídia'}</div>
            </div>
            <div className="stat-item">
              <div className="stat-value">{totalRegistros}</div>
              <div className="stat-label">Itens processados</div>
            </div>
          </div>
          <div className="mt-3">
            <h6>{modoKeywords ? 'Palavras-chave encontradas:' : 'Tipos encontrados:'}</h6>
            <div className="tipos-container">
              {Object.keys(processedData).map(item => (
                <div key={item} className="tipo-badge">
                  {item} <span className="badge">{processedData[item].length}</span>
                </div>
              ))}
            </div>
          </div>
        </Card.Body>
      </Card>
    );
  };
  
  // Estatísticas dos downloads
  const renderDownloadStats = () => {
    if (!processedData) return null;
    
    // Calcular estatísticas de download
    const totalTipos = Object.keys(downloadStatus).length;
    const totalArquivos = Object.values(downloadStatus).reduce(
      (sum, tipo) => sum + (tipo.baixados || 0), 0
    );
    const totalErros = Object.values(downloadStatus).reduce(
      (sum, tipo) => sum + (tipo.erros || 0), 0
    );
    
    return (
      <Card className="stat-card mb-4">
        <Card.Body>
          <Card.Title>Estatísticas de Download</Card.Title>
          <div className="stats-grid">
            <div className="stat-item">
              <div className="stat-value">{totalTipos}</div>
              <div className="stat-label">Tipos de mídia</div>
            </div>
            <div className="stat-item">
              <div className="stat-value">{totalArquivos}</div>
              <div className="stat-label">Arquivos baixados</div>
            </div>
            <div className="stat-item">
              <div className="stat-value">{totalErros}</div>
              <div className="stat-label">Erros</div>
            </div>
          </div>
          
          {Object.keys(downloadStatus).length > 0 && (
            <div className="mt-3">
              <h6>Status por tipo:</h6>
              <div className="tipos-container">
                {Object.entries(downloadStatus).map(([tipo, status]) => (
                  <div key={tipo} className="tipo-badge">
                    {tipo} <span className="badge">{status.baixados}/{status.total}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </Card.Body>
      </Card>
    );
  };
  
  return (
    <div className="app-container">
      {/* Barra de navegação */}
      <Navbar bg="primary" variant="dark" expand="lg" className="mb-4">
        <Container>
          <Navbar.Brand href="#home">
            <img
              src={process.env.PUBLIC_URL + '/logo192.png'}
              width="30"
              height="30"
              className="d-inline-block align-top me-2"
              alt="BrasPub logo"
            />
            Organizador de Planilhas BrasPub
          </Navbar.Brand>
          <Navbar.Toggle aria-controls="basic-navbar-nav" />
          <Navbar.Collapse id="basic-navbar-nav">
            <Nav className="ms-auto">
            </Nav>
          </Navbar.Collapse>
        </Container>
      </Navbar>
      
      <Container>
        {/* Cabeçalho */}
        <Row className="mb-4 text-center">
          <Col>
            <h1 className="display-5">Organizador de Planilhas</h1>
            <p className="lead">Processe e exporte planilhas organizadas</p>
          </Col>
        </Row>
        
        {/* Seletor de modo */}
        <Row className="mb-4 justify-content-center">
          <Col md={8} className="text-center">
            <ButtonGroup className="mode-selector">
              <Button
                variant={modoAtual === 'tipoMidia' ? "primary" : "outline-primary"}
                onClick={() => {
                  setModoAtual('tipoMidia');
                  setLoading(false);
                  setProcessingProgress(0);
                  setProcessStep('');
                }}
                className="mode-btn"
              >
                Tipo de Mídia
              </Button>
              <Button
                variant={modoAtual === 'palavrasChave' ? "primary" : "outline-primary"}
                onClick={() => {
                  setModoAtual('palavrasChave');
                  setLoading(false);
                  setProcessingProgress(0);
                  setProcessStep('');
                }}
                className="mode-btn"
              >
                Palavras-chave
              </Button>
              <Button
                variant={modoAtual === 'baixarArquivos' ? "primary" : "outline-primary"}
                onClick={() => {
                  setModoAtual('baixarArquivos');
                  setLoading(false);
                  setProcessingProgress(0);
                  setProcessStep('');
                }}
                className="mode-btn"
              >
                Baixar Arquivos
              </Button>
            </ButtonGroup>
          </Col>
        </Row>
        
        {/* Conteúdo principal */}
        <Row>
          {/* Coluna da esquerda - Upload e processamento */}
          <Col lg={7} className="mb-4">
            <Card className="main-card">
              <Card.Body>
                <Form.Group controlId="formFile" className="mb-4">
                  <Form.Label className="upload-label">
                    <span className="upload-icon">📄</span>
                    <strong>Selecione sua planilha Excel</strong>
                  </Form.Label>
                  <Form.Control 
                    type="file" 
                    accept=".xlsx,.xls" 
                    onChange={handleFileSelect} 
                    disabled={loading}
                    className="file-upload"
                  />
                  <Form.Text className="text-muted">
                    {modoKeywords 
                      ? "Formato esperado: Planilha com coluna de Palavras-chave" 
                      : "Formato esperado: Planilha com coluna 'Tipo da mídia'"}
                  </Form.Text>
                </Form.Group>

                {file && (
                  <div className="selected-file mb-3">
                    <div className="file-icon">📑</div>
                    <div className="file-info">
                      <div className="file-name">{file.name}</div>
                      <div className="file-size">{(file.size / 1024 / 1024).toFixed(2)} MB</div>
                    </div>
                    <Button 
                      variant="danger" 
                      size="sm" 
                      className="remove-file-btn"
                      onClick={handleRemoveFile}
                      title="Remover arquivo"
                    >
                      <span aria-hidden="true">−</span>
                    </Button>
                  </div>
                )}

                {loading && (
                  <div className="mb-4">
                    <div className="d-flex justify-content-between align-items-center mb-2">
                      <strong>{processStep}</strong>
                      <span>{processingProgress}%</span>
                    </div>
                    <ProgressBar 
                      animated 
                      now={processingProgress} 
                      variant="primary" 
                      className="custom-progress" 
                    />
                  </div>
                )}

                <Button 
                  variant="primary" 
                  onClick={modoBaixarArquivos ? processAndDownloadFiles : processAndExportExcel} 
                  disabled={!file || loading}
                  className="w-100 action-button"
                >
                  {loading ? (
                    <>
                      <Spinner
                        as="span"
                        animation="border"
                        size="sm"
                        role="status"
                        aria-hidden="true"
                        className="me-2"
                      />
                      Processando...
                    </>
                  ) : modoBaixarArquivos 
                      ? "Baixar Arquivos dos Links" 
                      : `Processar e Exportar Planilha ${modoKeywords ? 'de Palavras-chave' : ''}`}
                </Button>
              </Card.Body>
            </Card>

            {/* Mensagem de erro */}
            {error && (
              <Alert variant="danger" className="mt-3 custom-alert">
                <div className="alert-icon">⚠️</div>
                <div className="alert-content">{error}</div>
              </Alert>
            )}
            
            {/* Mensagem de sucesso */}
            {success && (
              <Alert variant="success" className="mt-3 custom-alert">
                <div className="alert-icon">✅</div>
                <div className="alert-content">{success}</div>
              </Alert>
            )}
          </Col>

          {/* Coluna da direita - Estatísticas e instruções */}
          <Col lg={5}>
            {processedData && (modoBaixarArquivos ? renderDownloadStats() : renderProcessedStats())}
            
            {/* Instruções */}
            <Card className="instruction-card">
              <Card.Body>
                <Card.Title>Como funciona</Card.Title>
                <ol className="instruction-list">
                  <li>
                    <div className="step-icon">1</div>
                    <div className="step-content">
                      <strong>Selecione o modo</strong>
                      <p>Escolha entre organizar por Tipo de Mídia ou Palavras-chave</p>
                    </div>
                  </li>
                  <li>
                    <div className="step-icon">2</div>
                    <div className="step-content">
                      <strong>Selecione o arquivo Excel</strong>
                      <p>{modoBaixarArquivos 
                        ? "Escolha uma planilha com links para download" 
                        : (modoKeywords 
                          ? "Escolha uma planilha com coluna de Palavras-chave" 
                          : "Escolha uma planilha com coluna 'Tipo da mídia'")}
                      </p>
                    </div>
                  </li>
                  <li>
                    <div className="step-icon">3</div>
                    <div className="step-content">
                      <strong>Clique em Processar</strong>
                      <p>A aplicação processará os dados automaticamente</p>
                    </div>
                  </li>
                  <li>
                    <div className="step-icon">4</div>
                    <div className="step-content">
                      <strong>Escolha onde salvar</strong>
                      <p>Selecione a localização para o arquivo processado</p>
                    </div>
                  </li>
                  <li>
                    <div className="step-icon">5</div>
                    <div className="step-content">
                      <strong>Pronto!</strong>
                      <p>{modoBaixarArquivos 
                        ? "Os arquivos serão baixados e organizados por data e tipo de mídia" 
                        : `A planilha será organizada com uma aba para cada ${modoKeywords ? 'palavra-chave' : 'tipo de mídia'}`}
                      </p>
                    </div>
                  </li>
                </ol>
              </Card.Body>
            </Card>
          </Col>
        </Row>
        
        {/* Rodapé */}
        <Row className="mt-4 mb-4">
          <Col className="text-center">
            <p className="text-muted">© {new Date().getFullYear()} BrasPub - Todos os direitos reservados</p>
          </Col>
        </Row>
      </Container>
    </div>
  );
}

export default App; 