import { lerDocumento } from './modules/reader.js';
import { gerarDocxPadronizado } from './modules/generator.js';

const zonaEnvio = document.getElementById('zona-envio');
const entradaArquivo = document.getElementById('entrada-arquivo');
const painelStatus = document.getElementById('painel-status');
const painelResultado = document.getElementById('painel-resultado');
const painelErro = document.getElementById('painel-erro');
const textoErro = document.getElementById('texto-erro');
const informacaoArquivo = document.getElementById('informacao-arquivo');
const botaoBaixar = document.getElementById('botao-baixar');

let arquivoGerado = null;
let nomeArquivoOriginal = '';

// Eventos de Arrastar e Soltar (Drag & Drop)
['dragover', 'dragenter'].forEach((nomeEvento) => {
  zonaEnvio.addEventListener(nomeEvento, (e) => {
    e.preventDefault();
    zonaEnvio.classList.add('drop-zone--active');
  });
});

['dragleave', 'dragend'].forEach((nomeEvento) => {
  zonaEnvio.addEventListener(nomeEvento, () => {
    zonaEnvio.classList.remove('drop-zone--active');
  });
});

zonaEnvio.addEventListener('drop', (e) => {
  e.preventDefault();
  zonaEnvio.classList.remove('drop-zone--active');

  if (e.dataTransfer.files.length) {
    entradaArquivo.files = e.dataTransfer.files;
    processarArquivo(e.dataTransfer.files[0]);
  }
});

// Evento disparado assim que o usuário escolhe um arquivo
entradaArquivo.addEventListener('change', () => {
  if (entradaArquivo.files.length) {
    processarArquivo(entradaArquivo.files[0]);
  }
});

async function processarArquivo(arquivo) {
  reiniciarInterface();
  exibirStatus('Extraindo texto e aplicando modelo padronizado...');

  try {
    nomeArquivoOriginal = arquivo.name.substring(0, arquivo.name.lastIndexOf('.'));
    
    // Extração do conteúdo
    const dadosExtraidos = await lerDocumento(arquivo);

    // Geração do novo arquivo formatado
    arquivoGerado = await gerarDocxPadronizado(dadosExtraidos);

    // Sucesso
    ocultarStatus();
    informacaoArquivo.textContent = `Arquivo processado: "${arquivo.name}"`;
    painelResultado.classList.remove('hidden');
  } catch (erro) {
    ocultarStatus();
    exibirErro(erro.message || 'Ocorreu um erro ao processar o arquivo.');
  }
}

botaoBaixar.addEventListener('click', () => {
  if (!arquivoGerado) return;

  const link = URL.createObjectURL(arquivoGerado);
  const elementoLink = document.createElement('a');
  elementoLink.href = link;
  elementoLink.download = `${nomeArquivoOriginal}_Padronizado.docx`;
  document.body.appendChild(elementoLink);
  elementoLink.click();
  document.body.removeChild(elementoLink);
  URL.revokeObjectURL(link);
});

function reiniciarInterface() {
  painelStatus.classList.add('hidden');
  painelResultado.classList.add('hidden');
  painelErro.classList.add('hidden');
  arquivoGerado = null;
}

function exibirStatus(mensagem) {
  document.getElementById('texto-status').textContent = mensagem;
  painelStatus.classList.remove('hidden');
}

function ocultarStatus() {
  painelStatus.classList.add('hidden');
}

function exibirErro(mensagem) {
  textoErro.textContent = mensagem;
  painelErro.classList.remove('hidden');
}