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

const opcaoRepublicado = document.getElementById('opcao-republicado');
const camposRepublicado = document.getElementById('campos-republicado');
const campoRevista = document.getElementById('campo-revista');
const campoLicenca = document.getElementById('campo-licenca');
const entradaGraficos = document.getElementById('entrada-graficos');
const listaGraficos = document.getElementById('lista-graficos');
const listaAvisos = document.getElementById('lista-avisos');

let dadosExtraidos = null;
let nomeArquivoOriginal = '';
let graficosAnexados = [];

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

// Opção "artigo republicado": mostra/esconde os campos de revista e licença
opcaoRepublicado.addEventListener('change', () => {
  camposRepublicado.classList.toggle('hidden', !opcaoRepublicado.checked);
});

// Gráficos anexados pelo usuário (imagens)
entradaGraficos.addEventListener('change', async () => {
  graficosAnexados = [];
  for (const arquivo of Array.from(entradaGraficos.files)) {
    try {
      graficosAnexados.push(await lerImagem(arquivo));
    } catch (erro) {
      exibirErro(`Não foi possível ler a imagem "${arquivo.name}". Use PNG, JPG, GIF ou BMP.`);
    }
  }
  listaGraficos.textContent = graficosAnexados.length
    ? `${graficosAnexados.length} imagem(ns) anexada(s): ${graficosAnexados.map((g) => g.nome).join(', ')}`
    : '';
});

async function lerImagem(arquivo) {
  const nome = arquivo.name;
  const e = nome.split('.').pop().toLowerCase();
  const extensao = e === 'jpeg' ? 'jpg' : e;
  if (!['png', 'jpg', 'gif', 'bmp'].includes(extensao)) {
    throw new Error('Formato de imagem não suportado.');
  }
  const dados = new Uint8Array(await arquivo.arrayBuffer());
  const bitmap = await createImageBitmap(arquivo);
  const imagem = { nome, dados, extensao, largura: bitmap.width, altura: bitmap.height };
  bitmap.close();
  return imagem;
}

async function processarArquivo(arquivo) {
  reiniciarInterface();
  exibirStatus('Extraindo texto e aplicando modelo padronizado...');

  try {
    nomeArquivoOriginal = arquivo.name.substring(0, arquivo.name.lastIndexOf('.'));

    // Extração do conteúdo
    dadosExtraidos = await lerDocumento(arquivo);

    // Sugere a revista de origem (detectada na citação "Como citar este artigo")
    campoRevista.value = dadosExtraidos.revistaDetectada || '';

    // Sucesso
    ocultarStatus();
    informacaoArquivo.textContent = `Arquivo processado: "${arquivo.name}"`;
    mostrarAvisos(dadosExtraidos.avisos || []);
    painelResultado.classList.remove('hidden');
  } catch (erro) {
    ocultarStatus();
    exibirErro(erro.message || 'Ocorreu um erro ao processar o arquivo.');
  }
}

botaoBaixar.addEventListener('click', async () => {
  if (!dadosExtraidos) return;

  try {
    // O .docx é gerado na hora do clique, para respeitar as opções marcadas na tela
    const arquivoGerado = await gerarDocxPadronizado(dadosExtraidos, {
      republicado: opcaoRepublicado.checked,
      revista: campoRevista.value,
      licenca: campoLicenca.value,
      graficos: graficosAnexados
    });

    const link = URL.createObjectURL(arquivoGerado);
    const elementoLink = document.createElement('a');
    elementoLink.href = link;
    elementoLink.download = `${nomeArquivoOriginal}_Padronizado.docx`;
    document.body.appendChild(elementoLink);
    elementoLink.click();
    document.body.removeChild(elementoLink);
    URL.revokeObjectURL(link);
  } catch (erro) {
    exibirErro(erro.message || 'Ocorreu um erro ao gerar o arquivo.');
  }
});

function mostrarAvisos(avisos) {
  listaAvisos.innerHTML = '';
  avisos.forEach((texto) => {
    const item = document.createElement('li');
    item.textContent = texto;
    listaAvisos.appendChild(item);
  });
  listaAvisos.classList.toggle('hidden', avisos.length === 0);
}

function reiniciarInterface() {
  painelStatus.classList.add('hidden');
  painelResultado.classList.add('hidden');
  painelErro.classList.add('hidden');
  dadosExtraidos = null;
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