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
const { ipcRenderer } = window.require('electron');

function App() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [processedData, setProcessedData] = useState(null);
  const [processStep, setProcessStep] = useState('');
  const [processingProgress, setProcessingProgress] = useState(0);
  const [modoKeywords, setModoKeywords] = useState(false);

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
        
        // Imediatamente iniciar o processo de exportação
        setProcessStep(modoKeywords ? 'Exportando planilha de palavras-chave...' : 'Exportando planilha...');
        setProcessingProgress(80);
        
        // Escolher o método de exportação baseado no modo
        const exportMethod = modoKeywords ? 'exportar-planilha-keywords' : 'exportar-planilha';
        console.log(`Chamando método de exportação: ${exportMethod}`);
        const exportResult = await ipcRenderer.invoke(exportMethod, result.dados);
        
        setProcessingProgress(100);
        
        if (exportResult && exportResult.status === 'sucesso') {
          setSuccess(`Planilha ${modoKeywords ? 'de palavras-chave ' : ''}processada e exportada com sucesso! ${exportResult.mensagem || ''}`);
        } else if (exportResult && exportResult.status === 'cancelado') {
          setError('Exportação cancelada pelo usuário.');
        } else {
          setError(`Erro ao exportar a planilha: ${exportResult?.mensagem || 'Resposta inválida do servidor'}`);
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
              <Nav.Link href="#home" active>Início</Nav.Link>
              <Nav.Link href="#about">Sobre</Nav.Link>
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
          <Col md={6} className="text-center">
            <ButtonGroup className="mode-selector">
              <Button
                variant={modoKeywords ? "outline-primary" : "primary"}
                onClick={() => {
                  setModoKeywords(false);
                  setFile(null);
                  setProcessedData(null);
                  setError(null);
                  setSuccess(null);
                }}
                className="mode-btn"
              >
                Tipo de Mídia
              </Button>
              <Button
                variant={modoKeywords ? "primary" : "outline-primary"}
                onClick={() => {
                  setModoKeywords(true);
                  setFile(null);
                  setProcessedData(null);
                  setError(null);
                  setSuccess(null);
                }}
                className="mode-btn"
              >
                Palavras-chave
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
                  onClick={processAndExportExcel} 
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
                  ) : `Processar e Exportar Planilha ${modoKeywords ? 'de Palavras-chave' : ''}`}
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
            {processedData && renderProcessedStats()}
            
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
                      <p>{modoKeywords 
                        ? "Escolha uma planilha com coluna de Palavras-chave" 
                        : "Escolha uma planilha com coluna 'Tipo da mídia'"}
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
                      <p>A planilha será organizada com uma aba para cada {modoKeywords ? 'palavra-chave' : 'tipo de mídia'}</p>
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