import JSZip from 'jszip';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_PARSER = () => new DOMParser();

export async function lerDocumento(arquivo) {
  const extensao = arquivo.name.split('.').pop().toLowerCase();
  if (extensao === 'docx') return await lerDocx(arquivo);
  if (extensao === 'pdf') return await lerPdf(arquivo);
  throw new Error('Formato não suportado. Envie apenas arquivos .docx ou .pdf.');
}

async function lerDocx(arquivo) {
  const vetorBuffer = await arquivo.arrayBuffer();
  const zip = await JSZip.loadAsync(vetorBuffer);
  const entradaDocumento = zip.file('word/document.xml');
  if (!entradaDocumento) throw new Error('O DOCX não possui word/document.xml.');

  const xml = await entradaDocumento.async('text');
  const doc = XML_PARSER().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Não foi possível ler a estrutura XML do DOCX.');

  const paragrafos = Array.from(doc.getElementsByTagNameNS(NS_W, 'p'));
  const blocos = [];
  const contadoresNumeracao = new Map();
  let indice = 0;

  for (const paragrafo of paragrafos) {
    const texto = extrairTextoParagrafo(paragrafo, { ignorarReferenciaDeNota: true }).trim();
    if (!texto) {
      indice++;
      continue;
    }

    const emTabela = estaDentroDeTabela(paragrafo);
    const recuo = obterRecuoEsquerdo(paragrafo);
    const tamanhoFonte = obterTamanhoFonteParagrafo(paragrafo);
    const numeracao = obterMarcadorNumeracao(paragrafo, doc, contadoresNumeracao);
    const estilo = obterEstiloParagrafo(paragrafo);
    const eTitulo = identificarTituloDocx(paragrafo, texto, estilo);
    const eCitacaoLonga = recuo >= 1500 && tamanhoFonte <= 20;
    const eNotaSolta = identificarParagrafoDeNotaSolta(paragrafo, texto, tamanhoFonte);

    blocos.push({
      tipo: eTitulo ? 'topico' : 'paragrafo',
      texto: numeracao ? `${numeracao} ${texto}` : texto,
      textoOriginal: texto,
      citacaoLonga: eCitacaoLonga,
      indice,
      emTabela,
      recuo,
      tamanhoFonte,
      estilo,
      eNotaSolta,
    });
    indice++;
  }

  if (!blocos.length) throw new Error('O arquivo DOCX está vazio ou não possui texto extraível.');

  const estrutura = montarEstrutura(blocos);
  estrutura.tipoEntrada = 'docx';
  estrutura.pacoteDocx = zip;
  estrutura.xmlDocumentoOriginal = xml;
  estrutura.indices = {
    citacaoAbnt: localizarIndiceCitacao(blocos),
    biografias: localizarIndicesBiografias(blocos, estrutura.autores),
    notasSoltas: blocos.filter((b) => b.eNotaSolta).map((b) => b.indice),
    paragrafos: blocos,
  };
  return estrutura;
}

async function lerPdf(arquivo) {
  const vetorBuffer = await arquivo.arrayBuffer();
  const pdfjsLib = window['pdfjs-dist/build/pdf'];
  if (!pdfjsLib) throw new Error('A biblioteca de leitura de PDF não foi carregada corretamente.');

  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const pdf = await pdfjsLib.getDocument({ data: vetorBuffer }).promise;
  const todasAsLinhas = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    const pagina = await pdf.getPage(i);
    const viewport = pagina.getViewport({ scale: 1 });
    const conteudo = await pagina.getTextContent();
    todasAsLinhas.push(...agruparEmLinhas(conteudo.items, viewport.height));
  }

  if (!todasAsLinhas.length) throw new Error('O PDF enviado não possui texto selecionável.');
  const perfil = calcularPerfilTipografico(todasAsLinhas);
  const linhasClassificadas = todasAsLinhas
    .map((linha) => classificarLinha(linha, perfil))
    .filter((linha) => linha.categoria !== 'rodape');
  const blocos = reconstruirParagrafos(linhasClassificadas, perfil);
  if (!blocos.length) throw new Error('Não foi possível identificar conteúdo no arquivo.');

  const dados = montarEstrutura(blocos);
  dados.tipoEntrada = 'pdf';
  dados.avisoPdf = 'PDF: tabelas e gráficos precisam ser conferidos/refeitos manualmente.';
  return dados;
}

function agruparEmLinhas(items, alturaPagina) {
  const linhas = [];
  let linhaAtual = null;
  for (const item of items) {
    const texto = item.str;
    const x = item.transform[4];
    const y = item.transform[5];
    const escala = Math.abs(item.transform[3]);
    if (!linhaAtual || Math.abs(y - linhaAtual.y) > 2) {
      linhaAtual = { y, alturaPagina, itens: [] };
      linhas.push(linhaAtual);
    }
    if (texto.trim()) linhaAtual.itens.push({ texto, x, escala, fonte: item.fontName });
  }
  return linhas.filter((l) => l.itens.length).map((l) => ({
    y: l.y, alturaPagina: l.alturaPagina, x: l.itens[0].x,
    texto: l.itens.map((i) => i.texto).join(' ').replace(/\s+/g, ' ').trim(),
    escala: media(l.itens.map((i) => i.escala)), fonte: modaTexto(l.itens.map((i) => i.fonte)),
  })).filter((l) => l.texto);
}

function calcularPerfilTipografico(linhas) {
  const escalaCorpo = moda(linhas.map((l) => Math.round(l.escala)));
  const linhasDoCorpo = linhas.filter((l) => Math.round(l.escala) === escalaCorpo);
  return { escalaCorpo, fonteCorpo: modaTexto(linhasDoCorpo.map((l) => l.fonte)), margemEsquerda: moda(linhasDoCorpo.map((l) => Math.round(l.x))) };
}

function classificarLinha(linha, perfil) {
  const ehMarcadorDeNota = linha.escala < perfil.escalaCorpo * 0.7;
  if (ehMarcadorDeNota) return { ...linha, categoria: 'rodape' };
  const fonteMenorQueCorpo = linha.escala < perfil.escalaCorpo - 1.5;
  if (fonteMenorQueCorpo) {
    if (linha.y < linha.alturaPagina * 0.25) return { ...linha, categoria: 'rodape' };
    const deslocamento = linha.x - perfil.margemEsquerda;
    if (deslocamento > 50 && deslocamento < 250) return { ...linha, categoria: 'citacao' };
  }
  const fonteDiferente = linha.fonte !== perfil.fonteCorpo && linha.escala >= perfil.escalaCorpo - 1;
  const pareceCaixaAlta = linha.texto === linha.texto.toUpperCase();
  if (fonteDiferente && pareceCaixaAlta && linha.texto.length < 120) return { ...linha, categoria: 'topico' };
  return { ...linha, categoria: 'corpo' };
}

function reconstruirParagrafos(linhasClassificadas, perfil) {
  const blocos = [];
  let atual = '';
  let citacao = false;
  const fechar = () => {
    if (atual.trim()) blocos.push({ tipo: 'paragrafo', texto: atual.trim(), citacaoLonga: citacao });
    atual = ''; citacao = false;
  };
  for (const linha of linhasClassificadas) {
    if (linha.categoria === 'topico') { fechar(); blocos.push({ tipo: 'topico', texto: linha.texto }); continue; }
    if (linha.categoria === 'citacao') {
      if (!citacao) fechar();
      citacao = true; atual += (atual ? ' ' : '') + linha.texto; continue;
    }
    if (citacao) fechar();
    const deslocamentoGrande = linha.x > perfil.margemEsquerda + 100;
    const temRecuo = linha.x > perfil.margemEsquerda + 15;
    if (deslocamentoGrande || temRecuo) fechar();
    else if (atual) atual += ' ';
    atual += linha.texto;
  }
  fechar();
  return blocos;
}

function montarEstrutura(blocos) {
  const autores = [];
  let tituloArtigo = '';
  let resumoTexto = '';
  let palavrasChave = '';
  const topicos = [];
  const referenciasBrutas = [];
  let topicoAtual = null;
  let fase = 'cabecalho';
  let aguardandoPalavrasChave = false;

  for (const bloco of blocos) {
    const texto = bloco.texto.trim();
    if (!texto || bloco.eNotaSolta) continue;

    if (fase === 'cabecalho') {
      if (/^resumo\b/i.test(texto)) { fase = 'resumo'; continue; }
      if (bloco.tipo === 'topico') { if (!tituloArtigo) tituloArtigo = texto; continue; }
      if (ehNaoAutor(texto)) continue;
      const candidatos = texto.split(/;+/).map((parte) => parte.trim()).filter(Boolean);
      for (const candidato of candidatos) {
        const autorLimpo = limparAutor(candidato);
        if (autorLimpo) autores.push(autorLimpo);
      }
      continue;
    }

    if (fase === 'resumo') {
      if (/^palavras[-\s]?chave\b/i.test(texto)) {
        const resto = texto.replace(/^palavras[-\s]?chave\s*:?\s*/i, '').trim();
        if (resto) palavrasChave = resto;
        else aguardandoPalavrasChave = true;
        continue;
      }
      if (aguardandoPalavrasChave) { palavrasChave = texto; aguardandoPalavrasChave = false; continue; }
      if (/^abstract\b/i.test(texto)) { fase = 'abstract'; continue; }
      if (bloco.tipo === 'topico') { fase = 'corpo'; }
      else { resumoTexto += (resumoTexto ? ' ' : '') + texto; continue; }
    }

    if (fase === 'abstract') {
      if (/^keywords?\b/i.test(texto)) continue;
      if (bloco.tipo === 'topico') fase = 'corpo'; else continue;
    }

    if (fase === 'referencias') {
      referenciasBrutas.push(bloco);
      continue;
    }

    if (bloco.tipo === 'topico') {
      if (/^\d+\.?\s*referências?$/i.test(texto) || /^referências?$/i.test(texto)) { fase = 'referencias'; continue; }
      if (fase === 'corpo' && topicoAtual && topicoAtual.paragrafos.length === 0) { topicoAtual.titulo += ' ' + texto; continue; }
      fase = 'corpo';
      topicoAtual = { titulo: texto, paragrafos: [] };
      topicos.push(topicoAtual);
      continue;
    }

    if (!topicoAtual) { topicoAtual = { titulo: 'Introdução', paragrafos: [] }; topicos.push(topicoAtual); }
    topicoAtual.paragrafos.push({ texto, eCitacaoLonga: bloco.citacaoLonga ?? pareceCitacaoLonga(texto), indice: bloco.indice });
  }

  const referenciasSeparadas = separarReferencias(referenciasBrutas, autores);
  topicos.forEach((topico, indice) => { topico.titulo = `${indice + 1}. ${limparNumeracaoExistente(topico.titulo)}`; });
  const numeroUltimoTopico = topicos.length + 1;
  const sumarioLista = topicos.map((t) => t.titulo);
  if (referenciasSeparadas.length) sumarioLista.push(`${numeroUltimoTopico}. REFERÊNCIAS`);

  return {
    autores,
    autor: autores[0] || '',
    tituloArtigo,
    dataAtual: '',
    creditoPublicacao: '',
    citacaoAbnt: montarRascunhoCitacaoAbnt(autores, tituloArtigo),
    citacaoAbntOriginal: '',
    sumarioTexto: sumarioLista.join(' • '),
    resumoTexto,
    palavrasChave,
    topicos,
    referencias: referenciasSeparadas.map((r) => r.texto),
    numeroUltimoTopico,
    notaDeRodape: true,
  };
}

function ehNaoAutor(texto) {
  return /universidade|instituto|faculdade|\(.*(?:cidade|pr|sp|rj|brasil|uf).*\)|\bbrasil\b|e-?mail|recebimento|aceitação|como citar|how to cite|\bdoi\b|\bissn\b|\b\d{1,2}\s+[a-zç]{3,9}\.\s+\d{4}\b/i.test(texto) || /^abstract\b/i.test(texto);
}

function limparAutor(texto) {
  let limpo = texto.replace(/\s+/g, ' ').trim();
  limpo = limpo.replace(/(?:Universidade|Instituto|Faculdade)\b.*$/i, '').trim();
  limpo = limpo.replace(/[;,:]+$/, '').trim();
  if (!limpo || ehNaoAutor(limpo)) return '';
  if (limpo.length > 120 || /\d/.test(limpo)) return '';
  return limpo;
}

function montarRascunhoCitacaoAbnt(autores, tituloArtigo) {
  const nomes = autores.map(formatarAutorAbnt).join('; ');
  return nomes ? `${nomes}. ${limparNumeracaoExistente(tituloArtigo)}. Doutrin, [DATA DE PUBLICAÇÃO]. Disponível em: [LINK].` : '';
}

function formatarAutorAbnt(nome) {
  const partes = nome.trim().split(/\s+/);
  if (partes.length < 2) return nome;
  const sobrenomesCompostos = new Set(['FILHO', 'JÚNIOR', 'JUNIOR', 'NETO', 'SOBRINHO']);
  const ultimo = partes[partes.length - 1];
  const penultimo = partes[partes.length - 2];
  const indice = sobrenomesCompostos.has(ultimo.toUpperCase()) && partes.length > 2 ? partes.length - 2 : partes.length - 1;
  const sobrenome = partes.slice(indice).join(' ');
  const resto = partes.slice(0, indice).join(' ');
  return `${sobrenome.toUpperCase()}, ${resto}`;
}

function separarReferencias(referencias, autores) {
  const resultado = [];
  let atual = null;
  for (const bloco of referencias) {
    const texto = bloco.texto.trim();
    if (ehBiografia(texto, autores)) break;
    if (!atual) { atual = { texto, indices: [bloco.indice] }; continue; }
    const continua = /(?:[:.]\s*$|\bIn:\.?\s*$|\bIn:\s*$)/i.test(atual.texto) || !/[.!?]$/.test(atual.texto);
    const novoAutor = /^[A-ZÀ-Ú][A-ZÀ-Ú' -]{1,},\s+[A-ZÀ-Úa-zà-ú]/.test(texto);
    if (continua && !novoAutor) { atual.texto += ` ${texto}`; atual.indices.push(bloco.indice); }
    else { resultado.push(finalizarReferencia(atual)); atual = { texto, indices: [bloco.indice] }; }
  }
  if (atual) resultado.push(finalizarReferencia(atual));
  return resultado.filter((r) => r.texto);
}

function finalizarReferencia(r) { return { ...r, texto: r.texto.replace(/\s+/g, ' ').replace(/\.+$/, '') + '.' }; }
function ehBiografia(texto, autores) { return autores.some((a) => texto.startsWith(a) && /\b(?:Advogado|Doutor|Professor|Professora|Mestre)\b/i.test(texto)); }
function localizarIndiceCitacao(blocos) { const i = blocos.findIndex((b) => /como citar este artigo|how to cite this article/i.test(b.texto)); return i >= 0 ? i + 1 : -1; }
function localizarIndicesBiografias(blocos, autores) { return blocos.filter((b) => ehBiografia(b.texto, autores)).map((b) => b.indice); }

function identificarTituloDocx(paragrafo, texto, estilo) {
  if (/heading|title|subtitle/i.test(estilo)) return true;
  return /^\d+\.?\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ0-9][A-ZÁÀÂÃÉÊÍÓÔÕÚÇ0-9 .–—-]{2,}$/.test(texto) && texto.length < 150;
}
function obterEstiloParagrafo(p) { const pPr = filho(p, 'pPr'); const pStyle = pPr && filho(pPr, 'pStyle'); return pStyle?.getAttributeNS(NS_W, 'val') || ''; }
function obterRecuoEsquerdo(p) { const ind = filho(filho(p, 'pPr'), 'ind'); return Number(ind?.getAttributeNS(NS_W, 'left') || 0); }
function obterTamanhoFonteParagrafo(p) {
  const runs = Array.from(p.getElementsByTagNameNS(NS_W, 'r')); const tamanhos = runs.map((r) => Number(filho(filho(r, 'rPr'), 'sz')?.getAttributeNS(NS_W, 'val') || 24)).filter(Boolean);
  return tamanhos.length ? Math.round(tamanhos.reduce((a, b) => a + b, 0) / tamanhos.length) : 24;
}
function filho(el, nome) { return el ? Array.from(el.childNodes).find((n) => n.nodeType === 1 && n.localName === nome) : null; }
function extrairTextoParagrafo(p, opcoes = {}) {
  return Array.from(p.getElementsByTagNameNS(NS_W, 't')).map((t) => t.textContent).join('')
    .replace(/\s+/g, ' ').trim();
}
function obterMarcadorNumeracao(p, doc, contadores) {
  const numPr = filho(filho(p, 'pPr'), 'numPr');
  if (!numPr) return '';
  const numId = Number(filho(numPr, 'numId')?.getAttributeNS(NS_W, 'val'));
  const ilvl = Number(filho(numPr, 'ilvl')?.getAttributeNS(NS_W, 'val') || 0);
  if (!numId) return '';
  const nums = Array.from(doc.getElementsByTagNameNS(NS_W, 'num'));
  const num = nums.find((n) => Number(n.getAttributeNS(NS_W, 'numId')) === numId);
  const abstractId = Number(filho(num, 'abstractNumId')?.getAttributeNS(NS_W, 'val'));
  const abstracts = Array.from(doc.getElementsByTagNameNS(NS_W, 'abstractNum'));
  const abstractNum = abstracts.find((a) => Number(a.getAttributeNS(NS_W, 'abstractNumId')) === abstractId);
  const lvls = abstractNum ? Array.from(abstractNum.getElementsByTagNameNS(NS_W, 'lvl')) : [];
  const lvl = lvls.find((l) => Number(l.getAttributeNS(NS_W, 'ilvl')) === ilvl) || lvls[0];
  if (!lvl) return '';
  const chave = `${numId}:${ilvl}`;
  const atual = (contadores.get(chave) || 0) + 1;
  contadores.set(chave, atual);
  for (let nivel = ilvl + 1; nivel < 10; nivel++) contadores.delete(`${numId}:${nivel}`);
  const formato = filho(lvl, 'lvlText')?.getAttributeNS(NS_W, 'val') || '%1.';
  return formato.replace(/%([1-9])/g, (m, n) => String(contadores.get(`${numId}:${Number(n)-1}`) || atual));
}

function estaDentroDeTabela(p) {
  let atual = p.parentNode;
  while (atual) {
    if (atual.nodeType === 1 && atual.localName === 'tbl') return true;
    atual = atual.parentNode;
  }
  return false;
}

function identificarParagrafoDeNotaSolta(p, texto, tamanhoFonte) {
  if (tamanhoFonte > 20) return false;
  return /^(?:10|[1-9])\s+[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ“"']/.test(texto) && texto.length > 70;
}
function limparNumeracaoExistente(titulo) { return titulo.replace(/^\d+[.)]?\s*/, '').trim(); }
function pareceCitacaoLonga(texto) { const t = texto.trim(); return /^(?:["“'‘]|\[\.\.\.\])/.test(t) && t.length > 150; }
function media(numeros) { return numeros.reduce((a, b) => a + b, 0) / numeros.length; }
function moda(numeros) { const c = new Map(); for (const n of numeros) c.set(n, (c.get(n) || 0) + 1); return [...c.entries()].sort((a,b) => b[1]-a[1])[0]?.[0] ?? 0; }
function modaTexto(textos) { const c = new Map(); for (const t of textos) c.set(t, (c.get(t) || 0) + 1); return [...c.entries()].sort((a,b) => b[1]-a[1])[0]?.[0] ?? ''; }
