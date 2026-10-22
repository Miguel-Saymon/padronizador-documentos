import mammoth from 'mammoth';

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

const MAPA_ESTILOS_DOCX = [
  "p[style-name='Title'] => h1.topico:fresh",
  "p[style-name='Subtitle'] => h1.topico:fresh",
  "p[style-name='Heading 1'] => h1.topico:fresh",
  "p[style-name='Heading 2'] => h1.topico:fresh",
  "p[style-name='Heading 3'] => h1.topico:fresh",
];

async function lerDocx(arquivo) {
  const vetorBuffer = await arquivo.arrayBuffer();
  const resultado = await mammoth.convertToHtml(
    { arrayBuffer: vetorBuffer },
    { styleMap: MAPA_ESTILOS_DOCX }
  );
  const html = resultado.value;

  if (!html || html.trim().length === 0) {
    throw new Error('O arquivo DOCX está vazio ou não possui texto extraível.');
  }

  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('ol').forEach((ol) => {
    if (ol.querySelector('li[id^="footnote-"]')) ol.remove();
  });

  const blocos = [];
  doc.body.querySelectorAll(':scope > *').forEach((el) => {
    const ehTopico = el.tagName === 'H1';
    const clone = el.cloneNode(true);
    clone.querySelectorAll('sup').forEach((sup) => sup.remove());
    const texto = clone.textContent.replace(/\s+/g, ' ').trim();

    if (!texto) return;
    blocos.push({ tipo: ehTopico ? 'topico' : 'paragrafo', texto });
  });

  if (blocos.length === 0) {
    throw new Error('Não foi possível identificar conteúdo no arquivo.');
  }

  return montarEstrutura(blocos);
}

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

  const blocos = reconstruirParagrafos(linhasClassificadas, perfil);

  if (blocos.length === 0) {
    throw new Error('Não foi possível identificar conteúdo no arquivo.');
  }

  return montarEstrutura(blocos);
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

function montarEstrutura(blocos) {
  const autores = [];
  let tituloArtigo = '';
  let resumoTexto = '';
  let palavrasChave = '';
  const topicos = [];
  const referencias = [];
  let topicoAtual = null;
  let fase = 'cabecalho'; // cabecalho -> resumo -> abstract -> corpo -> referencias

  for (const bloco of blocos) {
    const texto = bloco.texto.trim();
    if (!texto) continue;

    if (fase === 'cabecalho') {
      if (/^resumo\b/i.test(texto)) {
        fase = 'resumo';
        continue;
      }
      if (bloco.tipo === 'topico') {

        if (!tituloArtigo) tituloArtigo = texto;
        continue;
      }

      if (texto.includes(':')) {
        continue;
      }
      autores.push(texto);
      continue;
    }

    if (fase === 'resumo') {
      if (/^palavras[-\s]?chave/i.test(texto)) {
        palavrasChave = texto.replace(/^palavras[-\s]?chave:?\s*/i, '').trim();
        continue;
      }
      if (/^abstract\b/i.test(texto)) {
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

    if (fase === 'abstract') {
      if (/^keywords?\b/i.test(texto)) {
        continue;
      }
      if (bloco.tipo === 'topico') {
        fase = 'corpo'; 
      } else {
        continue; 
      }
    }

    if (fase === 'referencias') {
      referencias.push(texto);
      continue;
    }

    if (bloco.tipo === 'topico') {
      if (/^referências?$/i.test(texto)) {
        fase = 'referencias';
        continue;
      }

      if (fase === 'corpo' && topicoAtual && topicoAtual.paragrafos.length === 0) {
        topicoAtual.titulo += ' ' + texto;
        continue;
      }
      fase = 'corpo';
      topicoAtual = { titulo: texto, paragrafos: [] };
      topicos.push(topicoAtual);
      continue;
    }

    if (!topicoAtual) {
      topicoAtual = { titulo: 'Introdução', paragrafos: [] };
      topicos.push(topicoAtual);
    }
    const eCitacaoLonga = bloco.citacaoLonga ?? pareceCitacaoLonga(texto);
    topicoAtual.paragrafos.push({ texto, eCitacaoLonga });
  }

  topicos.forEach((topico, indice) => {
    topico.titulo = `${indice + 1}. ${limparNumeracaoExistente(topico.titulo)}`;
  });

  const numeroUltimoTopico = topicos.length + 1;
  const sumarioLista = topicos.map((t) => t.titulo);
  const referenciasSeparadas = separarReferencias(referencias);
  if (referenciasSeparadas.length > 0) {
    sumarioLista.push(`${numeroUltimoTopico}. REFERÊNCIAS`);
  }

  return {
    autor: autores.join('; ') || 'Nome Sobrenome',
    tituloArtigo,
    dataAtual: '',
    creditoPublicacao: '',
    citacaoAbnt: tituloArtigo ? montarRascunhoCitacaoAbnt(autores, tituloArtigo) : '',
    sumarioTexto: sumarioLista.join(' • '),
    resumoTexto,
    palavrasChave,
    topicos,
    referencias: referenciasSeparadas,
    numeroUltimoTopico,
  };
}

function montarRascunhoCitacaoAbnt(autores, tituloArtigo) {
  const nomesFormatados = autores
    .filter((a) => a && !/^(group|abstract|resumo)/i.test(a))
    .map(formatarAutorAbnt)
    .join('; ');
  const autorParte = nomesFormatados || 'SOBRENOME, Nome';
  return `${autorParte}. ${tituloArtigo}. Doutrin, [DATA DE PUBLICAÇÃO]. Disponível em: [LINK].`;
}

function formatarAutorAbnt(nomeCompleto) {
  const partes = nomeCompleto.trim().split(/\s+/);
  if (partes.length < 2) return nomeCompleto;

  const sobrenome = partes[partes.length - 1];
  const resto = partes.slice(0, -1).join(' ');
  return `${sobrenome.toUpperCase()}, ${resto}`;
}

function separarReferencias(referencias) {
  const textoUnico = referencias.join(' ');
  return textoUnico
    .split(/(?<=[.:])\s+(?=[A-ZÀ-Ú]{2,}(?:[A-ZÀ-Úa-zà-ú\s]*)?,\s)/)
    .map((r) => r.trim())
    .filter((r) => r.length > 0);
}

function limparNumeracaoExistente(titulo) {
  return titulo.replace(/^\d+[.)]?\s*/, '').trim();
}

function pareceCitacaoLonga(texto) {
  const t = texto.trim();
  const comecaComAspasOuReticencias = /^(["“'‘]|\[\.\.\.\])/.test(t);
  return comecaComAspasOuReticencias && t.length > 150;
}