import JSZip from 'jszip';

/* ============================================================================
 * LEITOR DE DOCUMENTOS
 *
 * Entrada : arquivo .docx ou .pdf
 * Saída   : objeto "dados" com a estrutura do artigo (autores, citação ABNT,
 *           resumo, tópicos, referências, notas de rodapé...), que o
 *           generator.js transforma no .docx padronizado.
 *
 * Modelo de "run" (trecho de texto com formatação):
 *   { texto, negrito, italico, sobrescrito }   -> texto comum
 *   { nota: '3' }                              -> chamada de nota de rodapé nº 3
 *
 * Itens dentro de um tópico:
 *   { tipo: 'paragrafo', runs, lista }   texto corrido (lista = item numerado)
 *   { tipo: 'citacao',   runs }          citação longa (recuo de 4 cm, corpo 10)
 *   { tipo: 'legenda',   runs }          "Tabela 1 – ...", "Fonte: ..."
 *   { tipo: 'tabela',    linhas, ... }   tabela copiada do original
 *   { tipo: 'imagem',    dados, ... }    imagem/gráfico embutido no original
 *   { tipo: 'grafico-pendente', ... }    gráfico vetorial que não pode ser copiado
 * ========================================================================== */

export async function lerDocumento(arquivo) {
  const extensao = arquivo.name.split('.').pop().toLowerCase();

  if (extensao === 'docx') {
    return await lerDocx(arquivo);
  } else if (extensao === 'pdf') {
    return await lerPdf(arquivo);
  } else {
    throw new Error('Formato não suportado. Envie apenas arquivos .docx ou .pdf.');
  }
}

/* ============================================================================
 * 1. UTILITÁRIOS DE XML (OOXML)
 * ========================================================================== */

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function lerXml(texto) {
  const doc = new DOMParser().parseFromString(texto, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error('Não foi possível interpretar o conteúdo interno do arquivo .docx.');
  }
  return doc;
}

const filhos = (el, nome) =>
  el ? Array.from(el.children).filter((n) => n.localName === nome) : [];
const filho = (el, nome) => filhos(el, nome)[0] || null;
const attr = (el, nome) => (el ? el.getAttributeNS(NS_W, nome) || null : null);
const ligado = (el) => !!el && !['0', 'false', 'off', 'none'].includes(attr(el, 'val'));

async function lerArquivoZip(zip, caminho) {
  const entrada = zip.file(caminho);
  return entrada ? await entrada.async('string') : null;
}

/* ============================================================================
 * 2. ESTILOS E NUMERAÇÃO DO DOCX
 * ========================================================================== */

function carregarEstilos(doc) {
  const estilos = new Map();
  let padrao = null;
  let tamanhoPadrao = 10;
  if (!doc) return { estilos, padrao, tamanhoPadrao };

  const sz = doc.getElementsByTagNameNS(NS_W, 'docDefaults')[0];
  const szPadrao = sz && sz.getElementsByTagNameNS(NS_W, 'sz')[0];
  if (szPadrao) tamanhoPadrao = Number(attr(szPadrao, 'val')) / 2;

  for (const s of Array.from(doc.getElementsByTagNameNS(NS_W, 'style'))) {
    const id = attr(s, 'styleId');
    if (!id) continue;
    if (attr(s, 'type') === 'paragraph' && s.getAttributeNS(NS_W, 'default') === '1') padrao = id;
    estilos.set(id, {
      nome: attr(filho(s, 'name'), 'val') || '',
      base: attr(filho(s, 'basedOn'), 'val'),
      pPr: filho(s, 'pPr'),
      rPr: filho(s, 'rPr'),
    });
  }
  return { estilos, padrao, tamanhoPadrao };
}

function lerPPrBasico(pPr) {
  const r = {};
  if (!pPr) return r;

  const ind = filho(pPr, 'ind');
  if (ind) {
    const esq = attr(ind, 'left') ?? attr(ind, 'start');
    if (esq !== null) r.esq = Number(esq);
    const primeira = attr(ind, 'firstLine');
    const pendurado = attr(ind, 'hanging');
    if (primeira !== null) r.primeira = Number(primeira);
    else if (pendurado !== null) r.primeira = -Number(pendurado);
  }

  const jc = filho(pPr, 'jc');
  if (jc) r.jc = attr(jc, 'val');

  const numPr = filho(pPr, 'numPr');
  if (numPr) {
    const numId = attr(filho(numPr, 'numId'), 'val');
    const ilvl = attr(filho(numPr, 'ilvl'), 'val');
    if (numId && numId !== '0') r.numPr = { numId, ilvl: Number(ilvl || 0) };
  }
  return r;
}

function lerRPrBasico(rPr) {
  const r = {};
  if (!rPr) return r;
  const sz = filho(rPr, 'sz');
  if (sz) r.tamanho = Number(attr(sz, 'val')) / 2;
  if (filho(rPr, 'b')) r.negrito = ligado(filho(rPr, 'b'));
  if (filho(rPr, 'i')) r.italico = ligado(filho(rPr, 'i'));
  return r;
}

function propsDoEstilo(id, ctx, profundidade = 0) {
  if (!id || profundidade > 8 || !ctx.estilos.has(id)) return {};
  const s = ctx.estilos.get(id);
  return {
    ...propsDoEstilo(s.base, ctx, profundidade + 1),
    ...lerPPrBasico(s.pPr),
    ...lerRPrBasico(s.rPr),
    nomeEstilo: s.nome,
  };
}

function carregarNumeracao(doc) {
  const numeracao = { nums: new Map(), abstratos: new Map(), contadores: new Map() };
  if (!doc) return numeracao;

  for (const a of Array.from(doc.getElementsByTagNameNS(NS_W, 'abstractNum'))) {
    const niveis = new Map();
    for (const l of filhos(a, 'lvl')) {
      niveis.set(Number(attr(l, 'ilvl')), {
        fmt: attr(filho(l, 'numFmt'), 'val') || 'decimal',
        texto: attr(filho(l, 'lvlText'), 'val') || '%1.',
        inicio: Number(attr(filho(l, 'start'), 'val') || 1),
      });
    }
    numeracao.abstratos.set(attr(a, 'abstractNumId'), niveis);
  }
  for (const n of Array.from(doc.getElementsByTagNameNS(NS_W, 'num'))) {
    numeracao.nums.set(attr(n, 'numId'), attr(filho(n, 'abstractNumId'), 'val'));
  }
  return numeracao;
}

function formatarNumero(valor, fmt) {
  const letra = (n, base) => {
    let s = '';
    let x = n;
    while (x > 0) {
      x -= 1;
      s = String.fromCharCode(base + (x % 26)) + s;
      x = Math.floor(x / 26);
    }
    return s;
  };
  const romano = (n) => {
    const tabela = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
      [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let s = '';
    let x = n;
    for (const [v, r] of tabela) {
      while (x >= v) { s += r; x -= v; }
    }
    return s;
  };
  switch (fmt) {
    case 'lowerLetter': return letra(valor, 97);
    case 'upperLetter': return letra(valor, 65);
    case 'lowerRoman': return romano(valor);
    case 'upperRoman': return romano(valor).toUpperCase();
    default: return String(valor);
  }
}

/** Calcula o rótulo ("(a)", "1)", "•"...) do próximo item de uma lista numerada. */
function proximoRotulo(numPr, ctx) {
  const abstratoId = ctx.numeracao.nums.get(numPr.numId);
  const niveis = ctx.numeracao.abstratos.get(abstratoId);
  if (!niveis || !niveis.has(numPr.ilvl)) return '';

  const nivel = niveis.get(numPr.ilvl);
  if (nivel.fmt === 'bullet') return '•';

  const chave = numPr.numId;
  const contadores = ctx.numeracao.contadores.get(chave) || [];
  contadores[numPr.ilvl] = (contadores[numPr.ilvl] ?? nivel.inicio - 1) + 1;
  contadores.length = numPr.ilvl + 1; // zera os níveis mais profundos
  ctx.numeracao.contadores.set(chave, contadores);

  return nivel.texto.replace(/%(\d)/g, (_, n) => {
    const idx = Number(n) - 1;
    const nivelRef = niveis.get(idx) || nivel;
    return formatarNumero(contadores[idx] ?? nivelRef.inicio, nivelRef.fmt);
  });
}

/* ============================================================================
 * 3. LEITURA DE PARÁGRAFOS, TABELAS E IMAGENS
 * ========================================================================== */

const EMU_POR_PIXEL = 9525;

function lerParagrafo(p, ctx) {
  const pPr = filho(p, 'pPr');
  const idEstilo = attr(filho(pPr, 'pStyle'), 'val') || ctx.padrao;
  const props = { ...propsDoEstilo(idEstilo, ctx), ...lerPPrBasico(pPr) };

  const runs = [];
  const extras = [];

  const lerRun = (r) => {
    const rPr = filho(r, 'rPr');
    const fmt = {
      negrito: filho(rPr, 'b') ? ligado(filho(rPr, 'b')) : !!props.negrito,
      italico: filho(rPr, 'i') ? ligado(filho(rPr, 'i')) : !!props.italico,
      // "elevado" pode vir como sobrescrito de verdade ou como texto deslocado para cima (w:position > 0),
      // que é como os arquivos convertidos de PDF marcam o número da nota.
      sobrescrito:
        attr(filho(rPr, 'vertAlign'), 'val') === 'superscript' || Number(attr(filho(rPr, 'position'), 'val')) > 0,
      tamanho: filho(rPr, 'sz') ? Number(attr(filho(rPr, 'sz'), 'val')) / 2 : props.tamanho,
    };
    let texto = '';
    const fechar = () => {
      if (texto) runs.push({ texto, ...fmt });
      texto = '';
    };

    const tratarFilho = (c) => {
      switch (c.localName) {
        case 't':
          texto += c.textContent;
          break;
        case 'tab':
        case 'br':
        case 'cr':
          texto += ' ';
          break;
        case 'noBreakHyphen':
          texto += '-';
          break;
        case 'footnoteReference':
          fechar();
          runs.push({ notaReal: attr(c, 'id') });
          break;
        case 'drawing': {
          const extra = interpretarDesenho(c, ctx);
          if (extra) extras.push(extra);
          break;
        }
        case 'pict': {
          const extra = interpretarPict(c, ctx);
          if (extra) extras.push(extra);
          break;
        }
        case 'AlternateContent': {
          const escolha = filho(c, 'Choice') || filho(c, 'Fallback');
          if (escolha) Array.from(escolha.children).forEach(tratarFilho);
          break;
        }
        default:
          break;
      }
    };
    Array.from(r.children).forEach(tratarFilho);
    fechar();
  };

  const visitar = (no) => {
    for (const n of Array.from(no.children)) {
      switch (n.localName) {
        case 'r':
          lerRun(n);
          break;
        case 'hyperlink':
        case 'ins':
        case 'moveTo':
        case 'smartTag':
        case 'fldSimple':
        case 'customXml':
          visitar(n);
          break;
        case 'sdt': {
          const conteudo = filho(n, 'sdtContent');
          if (conteudo) visitar(conteudo);
          break;
        }
        default: // w:del, w:moveFrom, w:pPr, w:proofErr... são ignorados
          break;
      }
    }
  };
  visitar(p);

  const runsLimpos = normalizarRuns(runs);
  const blocos = [];

  if (runsLimpos.length > 0) {
    const textoTotal = runsLimpos.map((r) => r.texto || '').join('');
    const comTexto = runsLimpos.filter((r) => r.texto && r.texto.trim());
    const maior = comTexto.sort((a, b) => b.texto.length - a.texto.length)[0];

    let rotulo = '';
    if (props.numPr) rotulo = proximoRotulo(props.numPr, ctx);

    blocos.push({
      origem: 'docx',
      tipo: 'p',
      runs: runsLimpos,
      esq: props.esq ?? 0,
      primeira: props.primeira ?? 0,
      tamanho: (maior && maior.tamanho) || props.tamanho || ctx.tamanhoPadrao,
      nomeEstilo: props.nomeEstilo || '',
      rotulo,
      todoNegrito: comTexto.length > 0 && comTexto.every((r) => r.negrito),
      todoItalico: comTexto.length > 0 && comTexto.every((r) => r.italico),
      texto: textoTotal,
    });
  }
  blocos.push(...extras);
  return blocos;
}

/** Junta runs vizinhos com a mesma formatação e remove espaços duplicados. */
function normalizarRuns(runs) {
  const saida = [];
  for (const r of runs) {
    if (r.texto === undefined) {
      saida.push(r); // chamadas de nota ({ nota } ou { notaReal }) passam direto
      continue;
    }
    const texto = r.texto.replace(/[\u00A0\u2007\u202F]/g, ' ').replace(/\u00AD/g, '');
    const anterior = saida[saida.length - 1];
    if (
      anterior &&
      anterior.texto !== undefined &&
      anterior.negrito === r.negrito &&
      anterior.italico === r.italico &&
      anterior.sobrescrito === r.sobrescrito
    ) {
      anterior.texto += texto;
    } else {
      saida.push({ ...r, texto });
    }
  }
  // espaços duplicados e espaços nas pontas do parágrafo
  saida.forEach((r) => {
    if (r.texto !== undefined) r.texto = r.texto.replace(/ {2,}/g, ' ');
  });
  while (saida.length && saida[0].texto !== undefined && !saida[0].texto.trim()) saida.shift();
  if (saida.length && saida[0].texto !== undefined) saida[0].texto = saida[0].texto.replace(/^\s+/, '');
  while (saida.length && saida[saida.length - 1].texto !== undefined && !saida[saida.length - 1].texto.trim()) saida.pop();
  const ultimo = saida[saida.length - 1];
  if (ultimo && ultimo.texto !== undefined) ultimo.texto = ultimo.texto.replace(/\s+$/, '');
  return saida;
}

/** Desenho moderno (DrawingML): imagem simples vira "imagem"; grupo de formas vira "gráfico pendente". */
function interpretarDesenho(desenho, ctx) {
  const contar = (nome) => desenho.getElementsByTagNameNS('*', nome).length;
  const blips = Array.from(desenho.getElementsByTagNameNS('*', 'blip'));
  const complexo = contar('wgp') > 0 || contar('txbxContent') > 0 || contar('chart') > 0 || blips.length > 1;

  if (complexo) return { origem: 'docx', tipo: 'grafico-pendente' };
  if (blips.length === 0) return null;

  const idRel = blips[0].getAttributeNS(NS_R, 'embed');
  const alvo = ctx.rels.get(idRel);
  if (!alvo) return null;

  const extensao = extensaoImagem(alvo);
  if (!extensao) return { origem: 'docx', tipo: 'grafico-pendente' };

  const medida = desenho.getElementsByTagNameNS('*', 'extent')[0];
  const largura = medida ? Number(medida.getAttribute('cx')) / EMU_POR_PIXEL : 400;
  const altura = medida ? Number(medida.getAttribute('cy')) / EMU_POR_PIXEL : 300;
  return { origem: 'docx', tipo: 'imagem', alvo, extensao, largura, altura };
}

/** Desenho antigo (VML). */
function interpretarPict(pict, ctx) {
  const contar = (nome) => pict.getElementsByTagNameNS('*', nome).length;
  const imagens = Array.from(pict.getElementsByTagNameNS('*', 'imagedata'));
  const complexo = contar('group') > 0 || contar('txbxContent') > 0 || imagens.length > 1;

  if (complexo) return { origem: 'docx', tipo: 'grafico-pendente' };
  if (imagens.length === 0) return null;

  const idRel = imagens[0].getAttributeNS(NS_R, 'id');
  const alvo = ctx.rels.get(idRel);
  const extensao = alvo ? extensaoImagem(alvo) : null;
  if (!alvo || !extensao) return { origem: 'docx', tipo: 'grafico-pendente' };

  const estilo = (pict.getElementsByTagNameNS('*', 'shape')[0] || { getAttribute: () => '' }).getAttribute('style') || '';
  const larguraPt = /width:\s*([\d.]+)pt/.exec(estilo);
  const alturaPt = /height:\s*([\d.]+)pt/.exec(estilo);
  return {
    origem: 'docx',
    tipo: 'imagem',
    alvo,
    extensao,
    largura: larguraPt ? Number(larguraPt[1]) * (96 / 72) : 400,
    altura: alturaPt ? Number(alturaPt[1]) * (96 / 72) : 300,
  };
}

function extensaoImagem(caminho) {
  const e = caminho.split('.').pop().toLowerCase();
  if (e === 'jpeg' || e === 'jpg') return 'jpg';
  if (['png', 'gif', 'bmp'].includes(e)) return e;
  return null; // emf, wmf, svg etc. não são suportados pela biblioteca de geração
}

function lerTabela(tbl, ctx) {
  const bordasTabela = filho(filho(tbl, 'tblPr'), 'tblBorders');
  const semBorda = (el) => !el || ['nil', 'none'].includes(attr(el, 'val'));
  const bordasPadrao = {
    top: bordasTabela ? !semBorda(filho(bordasTabela, 'top')) : true,
    bottom: bordasTabela ? !semBorda(filho(bordasTabela, 'bottom')) : true,
    left: bordasTabela ? !semBorda(filho(bordasTabela, 'left')) : true,
    right: bordasTabela ? !semBorda(filho(bordasTabela, 'right')) : true,
    insideH: bordasTabela ? !semBorda(filho(bordasTabela, 'insideH')) : true,
    insideV: bordasTabela ? !semBorda(filho(bordasTabela, 'insideV')) : true,
  };

  const grade = filhos(filho(tbl, 'tblGrid'), 'gridCol').map((g) => Number(attr(g, 'w')) || 0);
  const linhas = [];

  for (const tr of filhos(tbl, 'tr')) {
    const celulas = [];
    for (const tc of filhos(tr, 'tc')) {
      const tcPr = filho(tc, 'tcPr');
      const bordasCelula = filho(tcPr, 'tcBorders');
      const bordas = {};
      if (bordasCelula) {
        for (const lado of ['top', 'bottom', 'left', 'right']) {
          const b = filho(bordasCelula, lado);
          if (b) bordas[lado] = !semBorda(b);
        }
      }
      const vm = filho(tcPr, 'vMerge');
      const paragrafos = [];
      for (const p of filhos(tc, 'p')) {
        const bloco = lerParagrafo(p, ctx).find((b) => b.tipo === 'p');
        if (bloco) {
          const jc = { ...propsDoEstilo(attr(filho(filho(p, 'pPr'), 'pStyle'), 'val'), ctx), ...lerPPrBasico(filho(p, 'pPr')) }.jc;
          paragrafos.push({ runs: bloco.runs, jc: jc || 'left' });
        }
      }
      celulas.push({
        paragrafos,
        span: Number(attr(filho(tcPr, 'gridSpan'), 'val') || 1),
        vmerge: vm ? (attr(vm, 'val') === 'restart' ? 'restart' : 'continue') : null,
        largura: Number(attr(filho(tcPr, 'tcW'), 'w') || 0),
        bordas,
      });
    }
    linhas.push(celulas);
  }

  return { origem: 'docx', tipo: 'tabela', linhas, grade, bordas: bordasPadrao };
}

function lerCorpo(corpo, ctx) {
  const blocos = [];
  const visitar = (no) => {
    for (const n of Array.from(no.children)) {
      if (n.localName === 'p') blocos.push(...lerParagrafo(n, ctx));
      else if (n.localName === 'tbl') blocos.push(lerTabela(n, ctx));
      else if (n.localName === 'sdt') {
        const conteudo = filho(n, 'sdtContent');
        if (conteudo) visitar(conteudo);
      }
    }
  };
  visitar(corpo);
  return blocos;
}

function lerNotasReais(doc, ctx) {
  const notas = new Map();
  if (!doc) return notas;
  for (const n of filhos(doc.documentElement, 'footnote')) {
    const tipo = attr(n, 'type');
    const id = attr(n, 'id');
    if (tipo === 'separator' || tipo === 'continuationSeparator' || Number(id) <= 0) continue;

    const runs = [];
    for (const p of filhos(n, 'p')) {
      const bloco = lerParagrafo(p, ctx).find((b) => b.tipo === 'p');
      if (!bloco) continue;
      if (runs.length) runs.push({ texto: ' ' });
      runs.push(...bloco.runs.filter((r) => r.texto !== undefined));
    }
    const limpos = normalizarRuns(runs);
    if (limpos.length) notas.set(id, limpos);
  }
  return notas;
}

/* ============================================================================
 * 4. LEITURA DO DOCX (visão geral)
 * ========================================================================== */

async function lerDocx(arquivo) {
  const buffer = await arquivo.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);

  const xmlDocumento = await lerArquivoZip(zip, 'word/document.xml');
  if (!xmlDocumento) throw new Error('O arquivo DOCX está corrompido ou não possui conteúdo.');

  const docDocumento = lerXml(xmlDocumento);
  const xmlEstilos = await lerArquivoZip(zip, 'word/styles.xml');
  const xmlNumeracao = await lerArquivoZip(zip, 'word/numbering.xml');
  const xmlRels = await lerArquivoZip(zip, 'word/_rels/document.xml.rels');
  const xmlNotas = await lerArquivoZip(zip, 'word/footnotes.xml');

  const rels = new Map();
  if (xmlRels) {
    for (const r of Array.from(lerXml(xmlRels).documentElement.children)) {
      const alvo = r.getAttribute('Target') || '';
      rels.set(r.getAttribute('Id'), alvo.startsWith('/') ? alvo.slice(1) : `word/${alvo}`);
    }
  }

  const { estilos, padrao, tamanhoPadrao } = carregarEstilos(xmlEstilos ? lerXml(xmlEstilos) : null);
  const ctx = {
    estilos,
    padrao,
    tamanhoPadrao,
    rels,
    numeracao: carregarNumeracao(xmlNumeracao ? lerXml(xmlNumeracao) : null),
  };

  const corpo = docDocumento.getElementsByTagNameNS(NS_W, 'body')[0];
  let blocos = lerCorpo(corpo, ctx);
  const notasReais = lerNotasReais(xmlNotas ? lerXml(xmlNotas) : null, ctx);

  if (blocos.filter((b) => b.tipo === 'p').length === 0) {
    throw new Error('O arquivo DOCX está vazio ou não possui texto extraível.');
  }

  const avisos = [];
  blocos = classificarBlocosDocx(blocos);
  const notas = processarNotas(blocos, notasReais);
  blocos = unirContinuacoes(blocos);

  // Carrega os bytes das imagens embutidas (leitura assíncrona do zip).
  let imagensIgnoradas = 0;
  for (const b of blocos) {
    if (b.tipo !== 'imagem') continue;
    const entrada = zip.file(b.alvo);
    if (!entrada) {
      b.tipo = 'grafico-pendente';
      imagensIgnoradas += 1;
      continue;
    }
    b.dados = await entrada.async('uint8array');
  }

  const pendentes = blocos.filter((b) => b.tipo === 'grafico-pendente').length;
  if (pendentes > 0) {
    avisos.push(
      `${pendentes} gráfico(s) do original são desenhos vetoriais e não puderam ser copiados. ` +
        'Foi deixada uma marcação amarela no local — anexe a imagem do gráfico abaixo ou insira-o manualmente.'
    );
  }
  if (imagensIgnoradas > 0) avisos.push(`${imagensIgnoradas} imagem(ns) não puderam ser lidas.`);

  return montarEstrutura(blocos, { notas, avisos });
}

/* ============================================================================
 * 5. CLASSIFICAÇÃO DOS PARÁGRAFOS DO DOCX
 * ========================================================================== */

const REGEX_LEGENDA = /^(tabela|quadro|gráfico|grafico|figura|ilustração|ilustracao|imagem)\s+\d+\s*[–—\-:.]/i;
const REGEX_FONTE = /^(fontes?|nota|notas|elaboração|elaboracao)\s*:/i;

function ehTituloDocx(b) {
  if (/^(heading\s*\d|title|subtitle)/i.test(b.nomeEstilo)) return true;
  const t = b.texto.trim();
  if (b.rotulo) return false;
  return (
    b.todoNegrito &&
    !b.todoItalico &&
    t.length > 0 &&
    t.length <= 150 &&
    !/[.;,]$/.test(t) &&
    (t === t.toUpperCase() || b.tamanho >= 13)
  );
}

function classificarBlocosDocx(blocos) {
  // Margem esquerda "normal" do corpo = recuo mais comum entre parágrafos longos.
  const recuos = blocos.filter((b) => b.tipo === 'p' && b.texto.length > 100).map((b) => b.esq);
  const margemCorpo = recuos.length ? moda(recuos) : 0;

  const saida = [];
  for (const b of blocos) {
    if (b.tipo !== 'p') {
      saida.push(b);
      continue;
    }
    const anterior = saida[saida.length - 1];
    const anteriorEhLegenda = anterior && (anterior.sub === 'legenda' || ['tabela', 'imagem', 'grafico-pendente'].includes(anterior.tipo));
    const t = b.texto.trim();

    const novo = { ...b };
    if (ehTituloDocx(b)) {
      novo.tipo = 'topico';
    } else if (b.rotulo) {
      novo.tipo = 'paragrafo';
      novo.lista = true;
      novo.runs = [{ texto: `${b.rotulo} ` }, ...b.runs];
    } else if (b.esq - margemCorpo >= 1000) {
      novo.tipo = 'paragrafo';
      novo.sub = 'citacao';
    } else if (
      REGEX_LEGENDA.test(t) ||
      (REGEX_FONTE.test(t) && (b.tamanho <= 10.5 || anteriorEhLegenda)) ||
      (/^\*/.test(t) && anteriorEhLegenda)
    ) {
      novo.tipo = 'paragrafo';
      novo.sub = 'legenda';
    } else {
      novo.tipo = 'paragrafo';
    }
    saida.push(novo);
  }
  return saida;
}

/* ============================================================================
 * 6. NOTAS DE RODAPÉ
 *
 * Duas situações:
 *  a) notas reais do Word (w:footnoteReference)  -> viram notas de rodapé
 *  b) "notas falsas": o arquivo veio de um PDF e a nota ficou como um parágrafo
 *     solto no meio do texto ("1 Texto da nota...", fonte 10), com o número em
 *     sobrescrito no corpo. Essas também viram notas de rodapé de verdade e o
 *     parágrafo solto é removido do corpo.
 * ========================================================================== */

function ehDigitos(texto) {
  return /^\d{1,3}$/.test(texto.trim());
}

function numeroInicialDaNota(b) {
  const primeiro = b.runs.find((r) => r.texto !== undefined && r.texto.trim());
  if (!primeiro) return null;
  if (primeiro.sobrescrito && ehDigitos(primeiro.texto)) return primeiro.texto.trim();
  const m = /^(\d{1,3})\s+\S/.exec(b.texto.trim());
  return m ? m[1] : null;
}

function processarNotas(blocos, notasReais) {
  const notas = {}; // id sequencial -> runs
  let contador = 0;

  // (b) 1º passo: candidatos a "nota falsa"
  const candidatos = new Map(); // bloco -> numero
  blocos.forEach((b) => {
    if (b.tipo !== 'paragrafo' || b.sub || b.lista) return;
    if (b.tamanho > 10.5 && !(b.runs[0] && b.runs[0].sobrescrito)) return;
    const n = numeroInicialDaNota(b);
    if (n !== null && b.esq < 1000) candidatos.set(b, n);
  });

  // marcadores = números sobrescritos nos parágrafos que NÃO são candidatos
  const marcadores = new Set();
  blocos.forEach((b) => {
    if (b.tipo !== 'paragrafo' || candidatos.has(b)) return;
    b.runs.forEach((r) => {
      if (r.sobrescrito && r.texto !== undefined && ehDigitos(r.texto)) marcadores.add(r.texto.trim());
    });
  });

  // confirma: só vira nota se existir marcador correspondente no texto
  const textosDasNotas = new Map(); // numero -> runs
  const paraRemover = new Set();
  candidatos.forEach((n, b) => {
    if (!marcadores.has(n) || textosDasNotas.has(n)) return;
    const runs = b.runs.map((r) => ({ ...r }));
    const primeiro = runs.find((r) => r.texto !== undefined && r.texto.trim());
    primeiro.texto = primeiro.texto.replace(/^\s*\d{1,3}\s*/, '');
    textosDasNotas.set(n, normalizarRuns(runs));
    paraRemover.add(b);
  });

  // 2º passo: substitui os marcadores pelas chamadas de nota (na ordem do documento)
  const usadas = new Set();
  blocos.forEach((b) => {
    if (b.tipo !== 'paragrafo' || paraRemover.has(b)) return;
    const novosRuns = [];
    for (const r of b.runs) {
      if (r.notaReal !== undefined) {
        const texto = notasReais.get(r.notaReal);
        if (texto) {
          contador += 1;
          notas[contador] = texto;
          novosRuns.push({ nota: String(contador) });
        }
      } else if (r.sobrescrito && r.texto !== undefined && ehDigitos(r.texto) && textosDasNotas.has(r.texto.trim())) {
        const n = r.texto.trim();
        contador += 1;
        notas[contador] = textosDasNotas.get(n);
        usadas.add(n);
        novosRuns.push({ nota: String(contador) });
        if (/\s$/.test(r.texto)) novosRuns.push({ texto: ' ' }); // não engole o espaço depois do número
      } else {
        novosRuns.push(r);
      }
    }
    b.runs = ajustarPontuacaoDasNotas(novosRuns);
    b.texto = b.runs.map((r) => r.texto || '').join('');
  });

  // notas falsas sem marcador no corpo seriam perdidas: são anexadas ao parágrafo anterior
  textosDasNotas.forEach((runs, n) => {
    if (usadas.has(n)) return;
    const destino = [...blocos].reverse().find((b) => b.tipo === 'paragrafo' && !paraRemover.has(b) && !b.sub);
    if (!destino) return;
    contador += 1;
    notas[contador] = runs;
    destino.runs.push({ nota: String(contador) });
  });

  // remove os parágrafos de nota do fluxo principal
  for (let i = blocos.length - 1; i >= 0; i--) {
    if (paraRemover.has(blocos[i])) blocos.splice(i, 1);
  }
  return notas;
}

/** Em português a chamada de nota vai depois da pontuação: ")1." -> ")." + nota. */
function ajustarPontuacaoDasNotas(runs) {
  const saida = runs.map((r) => ({ ...r }));
  for (let i = 0; i < saida.length - 1; i++) {
    if (saida[i].nota === undefined) continue;
    const prox = saida[i + 1];
    if (prox.texto === undefined) continue;
    const m = /^[.,;:!?]+/.exec(prox.texto);
    if (!m) continue;
    prox.texto = prox.texto.slice(m[0].length);
    const anterior = saida[i - 1];
    if (anterior && anterior.texto !== undefined) {
      anterior.texto += m[0];
    } else {
      saida.splice(i, 0, { texto: m[0] });
      i += 1;
    }
  }
  return saida.filter((r) => r.nota !== undefined || r.texto !== '');
}

/* ============================================================================
 * 7. UNIÃO DE PARÁGRAFOS QUEBRADOS
 *
 * Em arquivos convertidos de PDF, um mesmo parágrafo costuma ser cortado na
 * quebra de página. Aqui eles são costurados de volta.
 * ========================================================================== */

const REGEX_FIM_DE_FRASE = /[.!?…:]["”'’)\]»]*\s*$/;

function textoDoBloco(b) {
  return (b.runs || []).map((r) => r.texto || '').join('');
}

function unirContinuacoes(blocos) {
  // O cabeçalho (autores, "Recebimento:", "Como citar...") nunca é costurado:
  // a união só começa depois do título "RESUMO" (ou desde o início, se ele não existir).
  const temResumo = blocos.some((b) => b.tipo === 'topico' && /^resumo\b/i.test(textoDoBloco(b)));
  let corpoIniciado = !temResumo;

  const saida = [];
  for (const b of blocos) {
    if (b.tipo === 'topico' && /^resumo\b/i.test(textoDoBloco(b))) corpoIniciado = true;
    const anterior = saida[saida.length - 1];
    const mesmoTipo =
      corpoIniciado &&
      anterior &&
      anterior.tipo === 'paragrafo' &&
      b.tipo === 'paragrafo' &&
      !b.lista &&
      anterior.sub === b.sub &&
      (b.sub === undefined || b.sub === 'citacao');

    // parágrafo que começa em minúscula (e não é "a)", "b)"...) é continuação do anterior
    const comecaEmMinuscula = /^[a-zà-ú]/.test(textoDoBloco(b)) && !/^[a-z]\)/.test(textoDoBloco(b));
    // item de lista só aceita continuação "em minúscula" (ex.: "(a) ... [quebra de página] a despesa total)")
    const permitido = mesmoTipo && (!anterior.lista || comecaEmMinuscula);
    if (permitido && (comecaEmMinuscula || !REGEX_FIM_DE_FRASE.test(textoDoBloco(anterior)))) {
      const semRecuoDeParagrafo = comecaEmMinuscula || b.sub === 'citacao' || (b.primeira || 0) < 400;
      if (semRecuoDeParagrafo) {
        anterior.runs = normalizarRuns([...anterior.runs, { texto: ' ' }, ...b.runs]);
        anterior.texto = textoDoBloco(anterior);
        continue;
      }
    }
    saida.push(b);
  }
  return saida;
}

/* ============================================================================
 * 8. LEITURA DE PDF (inalterada: heurísticas de geometria via pdf.js)
 * ========================================================================== */

async function lerPdf(arquivo) {
  const vetorBuffer = await arquivo.arrayBuffer();

  const pdfjsLib = window['pdfjs-dist/build/pdf'];
  if (!pdfjsLib) {
    throw new Error('A biblioteca de leitura de PDF não foi carregada corretamente.');
  }

  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  const pdf = await pdfjsLib.getDocument({ data: vetorBuffer }).promise;

  const todasAsLinhas = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const pagina = await pdf.getPage(i);
    const viewport = pagina.getViewport({ scale: 1 });
    const conteudo = await pagina.getTextContent();
    const linhasDaPagina = agruparEmLinhas(conteudo.items, viewport.height);
    todasAsLinhas.push(...linhasDaPagina);
  }

  if (todasAsLinhas.length === 0) {
    throw new Error('O PDF enviado não possui texto selecionável.');
  }

  const perfil = calcularPerfilTipografico(todasAsLinhas);
  const linhasClassificadas = todasAsLinhas
    .map((linha) => classificarLinha(linha, perfil))
    .filter((linha) => linha.categoria !== 'rodape');

  const blocosBrutos = reconstruirParagrafos(linhasClassificadas, perfil);

  if (blocosBrutos.length === 0) {
    throw new Error('Não foi possível identificar conteúdo no arquivo.');
  }

  const blocos = blocosBrutos.map((b) => ({
    origem: 'pdf',
    tipo: b.tipo,
    runs: [{ texto: b.texto }],
    sub: b.citacaoLonga ? 'citacao' : undefined,
    inferirCitacao: true,
  }));

  return montarEstrutura(blocos, {
    notas: {},
    avisos: [
      'Arquivos PDF não carregam notas de rodapé, tabelas nem gráficos de forma confiável: ' +
        'confira o resultado e insira esses elementos manualmente, se houver.',
    ],
  });
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

    if (texto.trim().length > 0) {
      linhaAtual.itens.push({ texto, x, escala, fonte: item.fontName });
    }
  }

  return linhas
    .filter((l) => l.itens.length > 0)
    .map((l) => ({
      y: l.y,
      alturaPagina: l.alturaPagina,
      x: l.itens[0].x,
      texto: l.itens.map((i) => i.texto).join(' ').replace(/\s+/g, ' ').trim(),
      escala: media(l.itens.map((i) => i.escala)),
      fonte: modaTexto(l.itens.map((i) => i.fonte)),
    }))
    .filter((l) => l.texto.length > 0);
}

function calcularPerfilTipografico(linhas) {
  const escalaCorpo = moda(linhas.map((l) => Math.round(l.escala)));
  const linhasDoCorpo = linhas.filter((l) => Math.round(l.escala) === escalaCorpo);
  const fonteCorpo = modaTexto(linhasDoCorpo.map((l) => l.fonte));
  const margemEsquerda = moda(linhasDoCorpo.map((l) => Math.round(l.x)));

  return { escalaCorpo, fonteCorpo, margemEsquerda };
}

function classificarLinha(linha, perfil) {
  const ehMarcadorDeNota = linha.escala < perfil.escalaCorpo * 0.7;
  if (ehMarcadorDeNota) {
    return { ...linha, categoria: 'rodape' };
  }

  const fonteMenorQueCorpo = linha.escala < perfil.escalaCorpo - 1.5;
  if (fonteMenorQueCorpo) {
    const naFaixaInferiorDaPagina = linha.y < linha.alturaPagina * 0.25;
    if (naFaixaInferiorDaPagina) return { ...linha, categoria: 'rodape' };
    const deslocamento = linha.x - perfil.margemEsquerda;
    const pareceRecuoDeCitacao = deslocamento > 50 && deslocamento < 250;
    if (pareceRecuoDeCitacao) return { ...linha, categoria: 'citacao' };
  }

  const fonteDiferente = linha.fonte !== perfil.fonteCorpo && linha.escala >= perfil.escalaCorpo - 1;
  const pareceCaixaAlta = linha.texto === linha.texto.toUpperCase();
  if (fonteDiferente && pareceCaixaAlta && linha.texto.length < 120) {
    return { ...linha, categoria: 'topico' };
  }

  return { ...linha, categoria: 'corpo' };
}

function reconstruirParagrafos(linhasClassificadas, perfil) {
  const blocos = [];
  let paragrafoAtual = '';
  let dentroDeCitacao = false;

  const fecharParagrafo = () => {
    if (paragrafoAtual.trim()) {
      blocos.push({ tipo: 'paragrafo', texto: paragrafoAtual.trim(), citacaoLonga: dentroDeCitacao });
    }
    paragrafoAtual = '';
    dentroDeCitacao = false;
  };

  for (const linha of linhasClassificadas) {
    if (linha.categoria === 'topico') {
      fecharParagrafo();
      blocos.push({ tipo: 'topico', texto: linha.texto });
      continue;
    }

    if (linha.categoria === 'citacao') {
      if (!dentroDeCitacao) fecharParagrafo();
      dentroDeCitacao = true;
      if (paragrafoAtual) paragrafoAtual += ' ';
      paragrafoAtual += linha.texto;
      continue;
    }

    if (dentroDeCitacao) fecharParagrafo();

    const deslocamentoGrande = linha.x > perfil.margemEsquerda + 100;
    if (deslocamentoGrande) {
      fecharParagrafo();
      blocos.push({ tipo: 'paragrafo', texto: linha.texto });
      continue;
    }

    const temRecuoDeParagrafo = linha.x > perfil.margemEsquerda + 15;
    if (temRecuoDeParagrafo) {
      fecharParagrafo();
    } else if (paragrafoAtual) {
      paragrafoAtual += ' ';
    }
    paragrafoAtual += linha.texto;
  }
  fecharParagrafo();

  return blocos;
}

function media(numeros) {
  return numeros.reduce((a, b) => a + b, 0) / numeros.length;
}

function moda(numeros) {
  const contagem = new Map();
  for (const n of numeros) contagem.set(n, (contagem.get(n) || 0) + 1);
  return [...contagem.entries()].sort((a, b) => b[1] - a[1])[0][0];
}

function modaTexto(textos) {
  const contagem = new Map();
  for (const t of textos) contagem.set(t, (contagem.get(t) || 0) + 1);
  return [...contagem.entries()].sort((a, b) => b[1] - a[1])[0][0];
}

/* ============================================================================
 * 9. MONTAGEM DA ESTRUTURA DO ARTIGO (comum a DOCX e PDF)
 *
 * Fases: cabecalho -> resumo -> abstract -> corpo -> referencias -> fim
 * ========================================================================== */

const REGEX_AFILIACAO =
  /universidade|universidad|university|instituto|institute|faculdade|faculty|departamento|programa de|centro universit|escola|\bUF[A-Z]{1,3}\b|\bPUC\b|\bUSP\b|\bIDP\b|@|\((?:[^)]*\b(?:brasil|brazil)\b[^)]*)\)|advogad|professor|doutor|mestre|pesquisador/i;

const REGEX_NAO_E_AUTOR = /revista|journal|artigo|article|vol\.|ISSN|DOI|creative commons|copyright/i;

const CONECTIVOS_DE_NOME = new Set(['de', 'da', 'do', 'dos', 'das', 'e', 'di', 'del', 'van', 'von', 'la', 'le']);

function textoDe(bloco) {
  return (bloco.runs || []).map((r) => r.texto || '').join('').replace(/\s+/g, ' ').trim();
}

function pareceNomeDePessoa(texto) {
  if (!texto || texto.length > 60 || /[0-9:]/.test(texto)) return false;
  if (REGEX_AFILIACAO.test(texto) || REGEX_NAO_E_AUTOR.test(texto)) return false;
  const palavras = texto.split(/\s+/);
  if (palavras.length < 2 || palavras.length > 7) return false;
  return palavras.every((p) => /^[A-ZÀ-Ú]/.test(p) || CONECTIVOS_DE_NOME.has(p.toLowerCase()));
}

/**
 * Nome do autor. Nos arquivos das revistas o nome vem em negrito + itálico, e a
 * universidade vem em texto comum (às vezes na MESMA linha). Só a parte em
 * negrito+itálico é considerada nome.
 */
function extrairNomeAutor(bloco, texto) {
  let nome = '';
  for (const r of bloco.runs || []) {
    if (r.texto === undefined) continue;
    if (r.negrito && r.italico) nome += r.texto;
    else if (!r.texto.trim()) nome += r.texto;
    else if (nome.trim()) break;
  }
  nome = nome.replace(/\s+/g, ' ').trim();

  if (nome.length > 2 && !REGEX_AFILIACAO.test(nome)) return nome;
  return pareceNomeDePessoa(texto) ? texto : null;
}

function limparCitacaoFonte(texto) {
  return texto
    .replace(/\s*ISSN\s*[\d-]+[Xx]?\.?/gi, '')
    .replace(/\s*e-?ISSN\s*[\d-]+[Xx]?\.?/gi, '')
    .replace(/\s*DOI:?\s*\S+/gi, '')
    .replace(/[<>]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/\s+\./g, '.')
    .trim();
}

function detectarRevista(citacao) {
  const m = /\b((?:Revista|Cadernos|Anais|Boletim|Journal)\s[^,.]+)/i.exec(citacao);
  return m ? m[1].trim() : '';
}

function paraCaixaDeSentenca(texto) {
  if (texto !== texto.toUpperCase()) return texto;
  const minusculo = texto.toLowerCase();
  return minusculo.charAt(0).toUpperCase() + minusculo.slice(1);
}

function formatarAutorAbnt(nomeCompleto) {
  const partes = nomeCompleto.trim().split(/\s+/);
  if (partes.length < 2) return nomeCompleto;
  const sobrenome = partes[partes.length - 1];
  const resto = partes.slice(0, -1).join(' ');
  return `${sobrenome.toUpperCase()}, ${resto}`;
}

function montarRascunhoCitacaoAbnt(autores, tituloArtigo) {
  const nomes = autores.map(formatarAutorAbnt).join('; ') || 'SOBRENOME, Nome';
  const titulo = paraCaixaDeSentenca(tituloArtigo).replace(/\.+$/, '');
  return `${nomes}. ${titulo}${/[?!]$/.test(titulo) ? '' : '.'} Doutrin, [DATA DE PUBLICAÇÃO]. Disponível em: [LINK].`;
}

function ehTituloReferencias(texto) {
  return /^(\d+[.)]?\s*)?refer[eê]ncias?(\s+bibliogr[aá]ficas?)?\.?$/i.test(texto.trim());
}

function limparNumeracaoExistente(titulo) {
  return titulo.replace(/^\d+(\.\d+)*[.)]?\s*/, '').trim();
}

function pareceCitacaoLonga(texto) {
  const t = texto.trim();
  const comecaComAspasOuReticencias = /^(["“'‘]|\[\.\.\.\]|\[…\])/.test(t);
  return comecaComAspasOuReticencias && t.length > 150;
}

function ehBiografia(bloco, texto) {
  const primeiro = (bloco.runs || []).find((r) => r.texto !== undefined && r.texto.trim());
  return !!primeiro && primeiro.negrito && primeiro.italico && texto.length > 40;
}

function montarEstrutura(blocos, contexto = {}) {
  const autores = [];
  let tituloArtigo = '';
  let citacaoFonte = '';
  let aguardandoCitacao = false;
  let resumoTexto = '';
  let palavrasChave = '';
  let aguardandoPalavrasChave = false;
  const topicos = [];
  const fragmentosReferencia = [];
  let topicoAtual = null;
  let fase = 'cabecalho'; // cabecalho -> resumo -> abstract -> corpo -> referencias -> fim

  const novoTopico = (titulo) => {
    topicoAtual = { titulo, itens: [] };
    topicos.push(topicoAtual);
  };

  const adicionarAoCorpo = (bloco, texto) => {
    if (bloco.tipo === 'topico') {
      if (ehTituloReferencias(texto)) {
        fase = 'referencias';
        return;
      }
      if (topicoAtual && topicoAtual.itens.length === 0) {
        topicoAtual.titulo += ' ' + texto; // título de tópico quebrado em duas linhas
        return;
      }
      novoTopico(texto);
      return;
    }

    if (!topicoAtual) novoTopico('Introdução');

    if (['tabela', 'imagem', 'grafico-pendente'].includes(bloco.tipo)) {
      const ultimo = topicoAtual.itens[topicoAtual.itens.length - 1];
      // vários pedaços de um mesmo desenho viram um só marcador
      if (bloco.tipo === 'grafico-pendente' && ultimo && ultimo.tipo === 'grafico-pendente') return;
      topicoAtual.itens.push(bloco);
      return;
    }

    const ehCitacao =
      bloco.sub === 'citacao' || (bloco.inferirCitacao && bloco.sub === undefined && pareceCitacaoLonga(texto));
    if (ehCitacao) topicoAtual.itens.push({ tipo: 'citacao', runs: bloco.runs });
    else if (bloco.sub === 'legenda') topicoAtual.itens.push({ tipo: 'legenda', runs: bloco.runs });
    else topicoAtual.itens.push({ tipo: 'paragrafo', runs: bloco.runs, lista: !!bloco.lista });
  };

  for (const bloco of blocos) {
    const ehObjeto = ['tabela', 'imagem', 'grafico-pendente'].includes(bloco.tipo);
    const texto = ehObjeto ? '' : textoDe(bloco);
    if (!ehObjeto && !texto) continue;

    // ---------- CABEÇALHO ----------
    if (fase === 'cabecalho') {
      if (ehObjeto) continue;

      if (/^resumo\b/i.test(texto) && (bloco.tipo === 'topico' || texto.length < 25)) {
        fase = 'resumo';
        continue;
      }

      const iniciaCorpoSemResumo =
        bloco.tipo === 'topico' && tituloArtigo && /^(\d+[.)]?\s+|introdu[cç][aã]o)/i.test(texto);

      if (!iniciaCorpoSemResumo) {
        if (aguardandoCitacao && texto.length > 30) {
          citacaoFonte = texto; // parágrafo logo após "Como citar este artigo"
          aguardandoCitacao = false;
          continue;
        }
        if (/^como citar|^how to cite/i.test(texto)) {
          aguardandoCitacao = true;
          continue;
        }
        if (bloco.tipo === 'topico') {
          if (!tituloArtigo) tituloArtigo = texto;
          continue;
        }
        if (/^(recebimento|aceita[cç][aã]o|aprovado|received|accepted|submitted|publicado)\b/i.test(texto)) continue;

        const nome = extrairNomeAutor(bloco, texto);
        if (nome) autores.push(nome);
        continue;
      }
      fase = 'corpo';
    }

    // ---------- RESUMO ----------
    if (fase === 'resumo') {
      if (ehObjeto) continue;

      if (/^palavras[-\s]?chave/i.test(texto)) {
        const resto = texto.replace(/^palavras[-\s]?chave\s*:?\s*/i, '').trim();
        if (resto) palavrasChave = resto;
        else aguardandoPalavrasChave = true; // "PALAVRAS-CHAVE" é um título; o conteúdo vem no parágrafo seguinte
        continue;
      }
      if (aguardandoPalavrasChave && bloco.tipo !== 'topico') {
        palavrasChave = texto;
        aguardandoPalavrasChave = false;
        continue;
      }
      if (/^(abstract|resumen|r[eé]sum[eé])\b/i.test(texto) && (bloco.tipo === 'topico' || texto.length < 20)) {
        fase = 'abstract';
        continue;
      }
      if (bloco.tipo === 'topico') {
        fase = 'corpo';
      } else {
        resumoTexto += (resumoTexto ? ' ' : '') + texto;
        continue;
      }
    }

    // ---------- ABSTRACT (descartado) ----------
    if (fase === 'abstract') {
      if (ehObjeto) continue;
      if (/^(keywords?|palabras[-\s]?clave|mots[-\s]cl)/i.test(texto)) continue;
      if (bloco.tipo === 'topico' || /^(\d+[.)]?\s*)?introdu[cç][aã]o/i.test(texto)) {
        fase = 'corpo';
      } else {
        continue;
      }
    }

    // ---------- REFERÊNCIAS ----------
    if (fase === 'referencias') {
      if (ehObjeto) continue;
      if (bloco.tipo === 'topico' || ehBiografia(bloco, texto)) {
        fase = 'fim'; // minibiografias dos autores etc. ficam de fora
        continue;
      }
      fragmentosReferencia.push({ texto, deParagrafo: bloco.origem === 'docx' });
      continue;
    }

    if (fase === 'fim') continue;

    // ---------- CORPO ----------
    fase = 'corpo';
    adicionarAoCorpo(bloco, texto);
  }

  topicos.forEach((topico, indice) => {
    topico.titulo = `${indice + 1}. ${limparNumeracaoExistente(topico.titulo)}`;
  });

  const numeroUltimoTopico = topicos.length + 1;
  const sumarioLista = topicos.map((t) => t.titulo);
  const referencias = montarReferencias(fragmentosReferencia);
  if (referencias.length > 0) sumarioLista.push(`${numeroUltimoTopico}. REFERÊNCIAS`);

  const citacaoAbnt = citacaoFonte
    ? limparCitacaoFonte(citacaoFonte)
    : tituloArtigo
      ? montarRascunhoCitacaoAbnt(autores, tituloArtigo)
      : '';

  return {
    autores: autores.length ? autores : ['Nome Sobrenome'],
    tituloArtigo,
    citacaoAbnt,
    revistaDetectada: citacaoFonte ? detectarRevista(citacaoAbnt) : '',
    sumarioTexto: sumarioLista.join(' • '),
    resumoTexto,
    palavrasChave,
    topicos,
    referencias,
    numeroUltimoTopico,
    notas: contexto.notas || {},
    avisos: contexto.avisos || [],
  };
}

/* ============================================================================
 * 10. REFERÊNCIAS
 *
 * Uma nova referência começa quando o parágrafo abre com SOBRENOME EM CAIXA
 * ALTA seguido de vírgula/ponto ("ALSTON, Lee", "BRASIL. Câmara...").
 * O que não começa assim é continuação da anterior (ex.: "FGV Editora, 2008."
 * ou "<https://...>. Acesso em: ...").
 * ========================================================================== */

const REGEX_INICIO_REFERENCIA = /^[A-ZÀ-Ú]{3,}[A-ZÀ-Ú'’\- ]*[,.]/;

const TRECHOS_COLADOS = [
  [/Disponívelem/g, 'Disponível em'],
  [/Acessoem/g, 'Acesso em'],
  [/\b(Mestrado|Doutorado)em(?=[A-ZÀ-Ú])/g, '$1 em '],
  [/Políticaorçamentáriano/g, 'Política orçamentária no'],
];

/** Conserta textos em que o PDF/Word "engoliu" os espaços ("FIGUEIREDO,ArgelinaCheibub;LIMONGI,Fernando"). */
function repararEspacos(texto) {
  const espacos = (texto.match(/\s/g) || []).length;
  if (texto.length < 25 || espacos / texto.length >= 0.1) return texto;

  let t = texto;
  TRECHOS_COLADOS.forEach(([regex, troca]) => {
    t = t.replace(regex, troca);
  });
  return t
    .split(' ')
    .map((token) => {
      if (/https?:|www\./i.test(token)) return token;
      return token
        .replace(/([,;])(?=[A-Za-zÀ-ÿ])/g, '$1 ')
        .replace(/([a-zà-ú)\d])\.(?=[A-ZÀ-Ú][a-zà-ú])/g, '$1. ')
        .replace(/([a-zà-ú)]),(?=\d{4}\b)/g, '$1, ')
        .replace(/(\d)f\.(?=\s|[A-ZÀ-Ú])/g, '$1 f.')
        .replace(/([a-zà-ú])(?=[A-ZÀ-Ú][a-zà-ú])/g, '$1 ')
        .replace(/\)–(?=\S)/g, ') – ')
        .replace(/(\S)–(?=[A-ZÀ-Ú])/g, '$1 – ')
        .replace(/\)(?=[A-ZÀ-Ú])/g, ') ');
    })
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function montarReferencias(fragmentos) {
  const referencias = [];
  for (const f of fragmentos) {
    const texto = repararEspacos(f.texto);
    const anterior = referencias[referencias.length - 1];
    const abreNova = REGEX_INICIO_REFERENCIA.test(texto);
    // em PDF as linhas são cortadas à força: só abre referência nova se a anterior terminou
    const anteriorTerminou = !anterior || /[.>)\d]$/.test(anterior);

    if (!anterior || (abreNova && (f.deParagrafo || anteriorTerminou))) referencias.push(texto);
    else referencias[referencias.length - 1] = `${anterior} ${texto}`;
  }
  return referencias.map((r) => r.replace(/\s+/g, ' ').trim()).filter((r) => r.length > 0);
}