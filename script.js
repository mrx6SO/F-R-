// Todas as referências para os elementos HTML
const fileInput = document.getElementById('fileInput');
const viewButton = document.getElementById('viewButton');
const reloadButton = document.getElementById('reloadButton');
const speakButton = document.getElementById('speakButton');
const volumeRange = document.getElementById('volumeRange');
const rateRange = document.getElementById('rateRange');
const pauseButton = document.getElementById('pauseButton');
const continueButton = document.getElementById('continueButton');
const fileContent = document.getElementById('fileContent');
const pdfViewerContainer = document.getElementById('pdfViewerContainer');
const controlButtons = document.getElementById('controlButtons');
const voiceList = document.getElementById('voiceList');
const downloadButton = document.getElementById('downloadButton');
const pdfDownloadButton = document.getElementById('pdfDownloadButton');
const wordDownloadButton = document.getElementById('wordDownloadButton');
const txtDownloadButton = document.getElementById('txtDownloadButton');
const volumeButton = document.getElementById('volumeButtons')

// Adicionar eventos para os botões de download
downloadButton.addEventListener('click', () => {
  // O formato de download selecionado
  const selectedFormat = document.getElementById('downloadFormat').value;
  // Verificar se há um arquivo selecionado antes de iniciar o download
  if (selectedFile) {
    downloadFile(selectedFile, selectedFormat);
  }
});

// Adicionar eventos para os botões de download específicos
pdfDownloadButton.addEventListener('click', () => {
  if (selectedFile) {
    // Inicia o download do arquivo PDF
    downloadFile(selectedFile, 'pdf');
  }
});

wordDownloadButton.addEventListener('click', () => {
  if (selectedFile) {
    // Inicia o download do arquivo Word (docx)
    downloadFile(selectedFile, 'docx');
  }
});

txtDownloadButton.addEventListener('click', () => {
  if (selectedFile) {
    // Inicia o download do arquivo de texto (txt)
    downloadFile(selectedFile, 'txt');
  }
});

// Inicializa variáveis para o arquivo selecionado e para a síntese de fala
let selectedFile = null;
let speechUtterance = new SpeechSynthesisUtterance();

// Obtém referência ao botão de upload
const uploadButton = document.getElementById('uploadButton');

// Adiciona um ouvinte de eventos para o botão de upload
uploadButton.addEventListener('click', () => {
  // Aciona o clique no input de arquivo para permitir a seleção de um arquivo
  fileInput.click();
});

// Adiciona um ouvinte de eventos para o input de arquivo
fileInput.addEventListener('change', (event) => {
  // Obtém o arquivo selecionado a partir do evento
  const file = event.target.files[0];
  // Verifica se um arquivo foi selecionado
  if (file) {
    // Exibe informações sobre o arquivo selecionado
    document.getElementById('fileInfo').style.display = 'block';
    document.getElementById('fileName').textContent = file.name;
    // Armazena o arquivo selecionado na variável
    selectedFile = file;
  } else {
    // Oculta as informações do arquivo se nenhum arquivo for selecionado
    document.getElementById('fileInfo').style.display = 'none';
    selectedFile = null;
  }
});

// Função para obter o formato do arquivo com base no tipo MIME
function getFileFormat(mimeType) {
  if (mimeType === 'application/pdf') {
    return 'PDF';
  } else if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
    return 'Word (docx)';
  } else if (mimeType === 'text/plain') {
    return 'TXT';
  }
  return 'Desconhecido';
}

// Adiciona um ouvinte de eventos para o botão de visualização
viewButton.addEventListener('click', () => {
  // Verifica se há um arquivo selecionado antes de iniciar a visualização
  if (selectedFile) {
    // Oculta as informações do arquivo
    document.getElementById('fileInfo').style.display = 'none';
    // Chama a função para ler e exibir o conteúdo do arquivo
    readFileContent(selectedFile);
    // Exibe os botões de controle
    controlButtons.style.display = 'block';
  }
});

// Adiciona um ouvinte de eventos para o botão de recarregar
reloadButton.addEventListener('click', () => {
  // Oculta os botões de controle
  controlButtons.style.display = 'none';
  // Limpa o conteúdo da div do arquivo e do visualizador PDF
  fileContent.innerHTML = ''; 
  pdfViewerContainer.innerHTML = ''; 
  // Recarrega a página
  location.reload();
});

// Adiciona um ouvinte de eventos para o botão de fala
speakButton.addEventListener('click', () => {
  // Chama a função para ler o texto com voz
  readTextWithVoice();
});

// Adiciona um ouvinte de eventos para o controle de volume
volumeRange.addEventListener('input', () => {
  // Ajusta o volume da síntese de fala com base no valor do controle deslizante
  speechUtterance.volume = parseFloat(volumeRange.value);
});

// Adiciona um ouvinte de eventos para o controle de taxa
rateRange.addEventListener('input', () => {
  // Ajusta a taxa de fala da síntese de fala com base no valor do controle deslizante
  speechUtterance.rate = parseFloat(rateRange.value);
});

// Adiciona um ouvinte de eventos para o botão de pausa
pauseButton.addEventListener('click', () => {
  // Pausa a síntese de fala
  speechSynthesis.pause();
});

// Adiciona um ouvinte de eventos para o botão de continuar
continueButton.addEventListener('click', () => {
  // Continua a síntese de fala
  speechSynthesis.resume();
});

// Função assíncrona para ler e exibir o conteúdo do arquivo
async function readFileContent(file) {
  // Verifica o tipo de arquivo e executa a ação correspondente
  if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
    // Converte o arquivo Word para HTML e exibe o conteúdo
    const htmlContent = await convertWordToHTML(file);
    document.getElementById('textContent').innerHTML = htmlContent;
  } else if (file.type === 'text/plain') {
    // Converte o arquivo de texto para texto simples e exibe o conteúdo
    const textContent = await convertTxtToText(file);
    document.getElementById('textContent').textContent = textContent;
  } else if (file.type === 'application/pdf') {
    // Exibe o arquivo PDF em um visualizador embutido
    pdfViewerContainer.innerHTML = '';
    const pdfViewer = document.createElement('iframe');
    pdfViewer.src = URL.createObjectURL(file);
    pdfViewer.style.width = '600px';
    pdfViewer.style.height = '400px';
    pdfViewerContainer.appendChild(pdfViewer);
  }
  // Mostra os botões inferiores quando um arquivo PDF é visualizado
  document.getElementById('bottomButtons').style.display = 'block';
  document.getElementById('volumeButtons').style.display = 'block';
}

// Função assíncrona para converter um arquivo Word para HTML
async function convertWordToHTML(file) {
  const arrayBuffer = await file.arrayBuffer();
  const wordArray = new Uint8Array(arrayBuffer);
  try {
    // Utiliza a biblioteca mammoth para converter o Word para HTML
    const converted = await mammoth.convertToHtml({ arrayBuffer: wordArray });
    return converted.value;
  } catch (error) {
    console.error('Erro na conversão do Word:', error);
    return null;
  }
}

// Função assíncrona para converter um arquivo de texto para texto simples
async function convertTxtToText(file) {
  const text = await file.text();
  return text;
}

// Função assíncrona para converter um arquivo de texto para PDF
async function convertTxtToPdf(file) {
  const textContent = await convertTxtToText(file);
  const pdfBlob = await createPdfFromText(textContent);
  return pdfBlob;
}

// Função para criar um Blob de PDF a partir do texto
function createPdfFromText(textContent) {
  return new Promise(resolve => {
    const pdfDoc = new jsPDF();
    pdfDoc.text(textContent, 10, 10);
    resolve(pdfDoc.output('blob'));
  });
}

// Função para ler o texto com voz
function readTextWithVoice() {
  // Verifica se há um arquivo selecionado e se há texto na div do conteúdo
  if (selectedFile) {
    const text = document.getElementById('textContent').textContent;
    // Verifica se há texto a ser lido
    if (text) {
      // Configura o texto no objeto de síntese de fala e inicia a leitura
      speechUtterance.text = text;
      speechSynthesis.cancel();
      speechSynthesis.speak(speechUtterance);
    }
  }
}

// Função assíncrona para converter HTML para DOCX
async function convertHtmlToDocx(htmlContent) {
  return new Promise((resolve, reject) => {
    // Implemente a lógica para converter HTML para DOCX aqui
    // Exemplo simples: criando um Blob com conteúdo HTML
    const blob = new Blob([htmlContent], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    resolve(blob);
    // Em uma implementação real, você precisará usar uma biblioteca ou ferramenta adequada
    // para realizar a conversão adequada de HTML para DOCX
  });
}

// Função assíncrona para converter um arquivo Word para um objeto Blob
async function convertWordToBlob(file) {
  const arrayBuffer = await file.arrayBuffer();
  const wordArray = new Uint8Array(arrayBuffer);
  try {
    // Utiliza a biblioteca mammoth para converter o Word para HTML
    const converted = await mammoth.convertToHtml({ arrayBuffer: wordArray });
    const htmlContent = converted.value;
    // Converte o conteúdo HTML para DOCX
    const docxBlob = await convertHtmlToDocx(htmlContent);
    return docxBlob;
  } catch (error) {
    console.error('Erro na conversão do Word para Blob:', error);
    return null;
  }
}

// Função assíncrona para fazer o download do arquivo
async function downloadFile(file, format) {
  const blob = await convertToBlob(file, format);

  // Verifica se o Blob foi criado com sucesso
  if (blob) {
    // Cria uma URL para o Blob e simula um clique no link de download
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = getDownloadFileName(file.name, format);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    // Revoga a URL do objeto
    URL.revokeObjectURL(url);
  }
}


// Função para obter o nome de download com base no formato
function getDownloadFileName(originalFileName, format) {
  const baseName = originalFileName.replace(/\.[^/.]+$/, '');
  return format === 'docx' ? baseName + '.docx' :
         format === 'txt' ? baseName + '.txt' :
         format === 'pdf' ? baseName + '.pdf' :
         baseName + '.html';
}

// Função assíncrona para converter o arquivo para Blob com base no formato
async function convertToBlob(file, format) {
  if (format === 'docx') {
    return await convertWordToBlob(file);
  } else if (format === 'txt') {
    return new Blob([await convertTxtToText(file)], { type: 'text/plain' });
  } else if (format === 'pdf') {
    return new Blob([file], { type: 'application/pdf' });
  } else {
    return new Blob([await convertToHTML(file)], { type: 'text/html' });
  }
}




// Função para configurar as vozes disponíveis na lista suspensa
function voices() {
  for (let voice of speechSynthesis.getVoices()) {
    let selected = voice.name === "Google US English" ? "selected" : "";
    let option = `<option value="${voice.name}" ${selected}>${voice.name} (${voice.lang})</option>`;
    voiceList.insertAdjacentHTML("beforeend", option);
  }
}

// Função para atualizar a exibição dos botões de download com base no tipo de arquivo selecionado
function updateDownloadButtons() {
  if (selectedFile) {
    pdfDownloadButton.style.display = selectedFile.type === 'application/pdf' ? 'block' : 'none';
    wordDownloadButton.style.display = selectedFile.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ? 'block' : 'none';
    txtDownloadButton.style.display = selectedFile.type === 'text/plain' ? 'block' : 'none';
  } else {
    pdfDownloadButton.style.display = 'none';
    wordDownloadButton.style.display = 'none';
    txtDownloadButton.style.display = 'none';
  }
}

// Chama a função de atualização dos botões de download
updateDownloadButtons();
