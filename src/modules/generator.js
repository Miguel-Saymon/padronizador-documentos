import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  LineRuleType,
  FootnoteReferenceRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  VerticalAlign,
  VerticalMergeType,
  convertInchesToTwip
} from 'docx';

/* ============================================================================
 * PADRÃO DE FORMATAÇÃO DO DOUTRIN
 * Unidades: fonte em meios-pontos (24 = 12pt); recuos em twips (567 = 1 cm).
 * ========================================================================== */

const DOUTRIN_CONFIG = {
  font: 'Helvetica Neue',

  // No Word: 240 = Simples / 360 = 1,5 Linhas
  lineSpacingSimple: 240,
  lineSpacingOneAndHalf: 360,

  size8pt: 16,
  size9pt: 18,
  size10pt: 20,
  size12pt: 24,
  size14pt: 28,

  firstLineIndent: 567,    // 1 cm — SOMENTE na primeira linha do parágrafo
  citationIndent: 2268,    // 4 cm de recuo à esquerda nas citações longas

  pageMargin: convertInchesToTwip(2.5 / 2.54),
  textWidth: 9071          // largura útil da página A4 com margens de 2,5 cm (16 cm)
};

const REGEX_NOVA_LEGENDA = /^(tabela|quadro|gráfico|grafico|figura|ilustração|ilustracao)\s+\d+/i;
const textoDoItem = (item) => (item.runs || []).map((r) => r.texto || '').join('').trim();

const LARGURA_MAXIMA_IMAGEM_PX = 600;
const ALTURA_MAXIMA_IMAGEM_PX = 700;

/* ============================================================================
 * FUNÇÃO PRINCIPAL
 *
 * dados  : objeto produzido pelo reader.js
 * opcoes : {
 *   republicado : boolean  -> inclui o crédito "Texto originalmente publicado..."
 *   revista     : string   -> nome da revista de origem
 *   licenca     : string   -> ex.: "Creative Commons Attribution"
 *   graficos    : [{ nome, dados, extensao, largura, altura }]  -> imagens anexadas pelo usuário
 * }
 * ========================================================================== */

export async function gerarDocxPadronizado(dados, opcoes = {}) {
  const elementos = [];
  const notasDeRodape = {};

  // Notas de rodapé reais (o texto delas NÃO vai para o corpo do documento)
  Object.entries(dados.notas || {}).forEach(([id, runs]) => {
    notasDeRodape[id] = {
      children: [
        new Paragraph({
          style: 'NotaDoutrin',
          alignment: AlignmentType.JUSTIFIED,
          children: [
            new TextRun({ text: ' ', font: DOUTRIN_CONFIG.font, size: DOUTRIN_CONFIG.size10pt }),
            ...criarRuns(runs, { size: DOUTRIN_CONFIG.size10pt })
          ]
        })
      ]
    };
  });

  const topicos = injetarGraficos(dados.topicos || [], opcoes.graficos || []);

  // PARTE 1: INICIAL — espaçamento simples, 8pt
  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size8pt, DOUTRIN_CONFIG.lineSpacingSimple));

  // Crédito de republicação (só quando o usuário marcar a opção)
  if (opcoes.republicado) {
    const revista = (opcoes.revista || '').trim() || '[NOME DA REVISTA]';
    const licenca = (opcoes.licenca || '').trim() || 'Creative Commons Attribution';
    elementos.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({
            text: `Texto originalmente publicado na ${revista}, sob a licença ${licenca}.`,
            font: DOUTRIN_CONFIG.font,
            size: DOUTRIN_CONFIG.size8pt,
            bold: true
          })
        ]
      })
    );
  }

  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size8pt, DOUTRIN_CONFIG.lineSpacingSimple));

  // Citação ABNT
  elementos.push(
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
      children: [
        new TextRun({
          text: dados.citacaoAbnt || 'SOBRENOME, Nome. Título do artigo. Doutrin, [DATA DE PUBLICAÇÃO]. Disponível em: [LINK].',
          font: DOUTRIN_CONFIG.font,
          size: DOUTRIN_CONFIG.size8pt
        })
      ]
    })
  );

  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size8pt, DOUTRIN_CONFIG.lineSpacingSimple));

  // PARTE 2: AUTORES — um nome por linha, 14pt negrito-itálico, alinhado à direita
  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

  (dados.autores && dados.autores.length ? dados.autores : ['Nome Sobrenome']).forEach((autor) => {
    elementos.push(
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({
            text: autor,
            font: DOUTRIN_CONFIG.font,
            size: DOUTRIN_CONFIG.size14pt,
            bold: true,
            italics: true
          })
        ]
      })
    );
  });

  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));
  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

  // PARTE 3: SUMÁRIO — espaçamento simples, 12pt
  if (dados.sumarioTexto) {
    elementos.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({
            text: dados.sumarioTexto.toUpperCase(),
            font: DOUTRIN_CONFIG.font,
            size: DOUTRIN_CONFIG.size12pt
          })
        ]
      })
    );
  }

  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingSimple));

  // PARTE 4: RESUMO E PALAVRAS-CHAVE
  if (dados.resumoTexto) {
    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

    elementos.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({ text: 'RESUMO', font: DOUTRIN_CONFIG.font, size: DOUTRIN_CONFIG.size14pt, bold: true })
        ]
      })
    );

    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

    elementos.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({ text: dados.resumoTexto, font: DOUTRIN_CONFIG.font, size: DOUTRIN_CONFIG.size12pt })
        ]
      })
    );

    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

    if (dados.palavrasChave) {
      elementos.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
          children: [
            new TextRun({
              text: `Palavras-chave: ${dados.palavrasChave}`,
              font: DOUTRIN_CONFIG.font,
              size: DOUTRIN_CONFIG.size12pt
            })
          ]
        })
      );
    }

    for (let i = 0; i < 3; i++) {
      elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));
    }
  }

  // PARTE 5: CORPO DO TEXTO E TÓPICOS
  topicos.forEach((topico) => {
    // Título do tópico: SEM recuo
    elementos.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({
            text: topico.titulo.toUpperCase(),
            font: DOUTRIN_CONFIG.font,
            size: DOUTRIN_CONFIG.size14pt,
            bold: true
          })
        ]
      })
    );

    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

    let tipoAnterior = null; // para decidir as linhas em branco entre blocos diferentes

    (topico.itens || []).forEach((item) => {
      const categoria = categoriaDoItem(item);

      // Linha em branco quando muda o tipo de bloco (texto <-> citação <-> tabela/gráfico) e entre citações seguidas
      const ehNovaLegenda = item.tipo === 'legenda' && REGEX_NOVA_LEGENDA.test(textoDoItem(item));
      if (tipoAnterior && (tipoAnterior !== categoria || categoria === 'citacao' || ehNovaLegenda)) {
        const espacoCitacao = tipoAnterior === 'citacao' || categoria === 'citacao';
        elementos.push(
          espacoCitacao
            ? criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingSimple)
            : criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf)
        );
      }

      elementos.push(...criarElementosDoItem(item));
      tipoAnterior = categoria;
    });

    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));
    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));
  });

  // PARTE 6: REFERÊNCIAS
  if (dados.referencias && dados.referencias.length > 0) {
    elementos.push(
      new Paragraph({
        spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({
            text: `${dados.numeroUltimoTopico}. REFERÊNCIAS`,
            font: DOUTRIN_CONFIG.font,
            size: DOUTRIN_CONFIG.size14pt,
            bold: true
          })
        ]
      })
    );

    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingSimple));
    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingSimple));

    dados.referencias.forEach((ref) => {
      elementos.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
          children: [
            new TextRun({
              text: ref.endsWith('.') ? ref : `${ref}.`,
              font: DOUTRIN_CONFIG.font,
              size: DOUTRIN_CONFIG.size10pt
            })
          ]
        })
      );
      elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size10pt, DOUTRIN_CONFIG.lineSpacingSimple));
    });
  }

  const documento = new Document({
    styles: {
      default: {
        document: { run: { font: DOUTRIN_CONFIG.font, size: DOUTRIN_CONFIG.size12pt } }
      },
      paragraphStyles: [
        {
          id: 'NotaDoutrin',
          name: 'Nota Doutrin',
          basedOn: 'Normal',
          run: { font: DOUTRIN_CONFIG.font, size: DOUTRIN_CONFIG.size10pt },
          paragraph: { spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, before: 0, after: 0 } }
        }
      ]
    },
    footnotes: notasDeRodape,
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: DOUTRIN_CONFIG.pageMargin,
              right: DOUTRIN_CONFIG.pageMargin,
              bottom: DOUTRIN_CONFIG.pageMargin,
              left: DOUTRIN_CONFIG.pageMargin
            }
          }
        },
        children: elementos
      }
    ]
  });

  return await Packer.toBlob(documento);
}

/* ============================================================================
 * BLOCOS DO CORPO
 * ========================================================================== */

/** Agrupa os itens em "famílias" para decidir onde entram as linhas em branco. */
function categoriaDoItem(item) {
  if (item.tipo === 'citacao') return 'citacao';
  if (['legenda', 'tabela', 'imagem', 'grafico-pendente'].includes(item.tipo)) return 'figura';
  return 'texto';
}

function criarElementosDoItem(item) {
  switch (item.tipo) {
    case 'citacao':
      // Citação longa: recuo de 4 cm, espaçamento simples, fonte 10
      return [
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: DOUTRIN_CONFIG.citationIndent },
          spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
          children: criarRuns(item.runs, { size: DOUTRIN_CONFIG.size10pt })
        })
      ];

    case 'legenda':
      return [
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
          children: criarRuns(item.runs, { size: DOUTRIN_CONFIG.size9pt })
        })
      ];

    case 'tabela':
      return [criarTabela(item)];

    case 'imagem':
      return [criarParagrafoImagem(item)];

    case 'grafico-pendente':
      return [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 120, after: 120 },
          children: [
            new TextRun({
              text: '[INSERIR GRÁFICO AQUI]',
              font: DOUTRIN_CONFIG.font,
              size: DOUTRIN_CONFIG.size10pt,
              bold: true,
              highlight: 'yellow'
            })
          ]
        })
      ];

    default:
      // Parágrafo comum: recuo de 1 cm SÓ na primeira linha; itens de lista ficam sem recuo
      return [
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          indent: item.lista ? undefined : { firstLine: DOUTRIN_CONFIG.firstLineIndent },
          spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
          children: criarRuns(item.runs, { size: DOUTRIN_CONFIG.size12pt })
        })
      ];
  }
}

/**
 * Converte os "runs" do leitor em TextRuns. Preserva negrito/itálico do original
 * e transforma { nota } em chamada de nota de rodapé de verdade.
 */
function criarRuns(runs, base) {
  return (runs || []).map((r) => {
    if (r.nota !== undefined) return new FootnoteReferenceRun(Number(r.nota));
    return new TextRun({
      text: r.texto,
      font: DOUTRIN_CONFIG.font,
      size: base.size,
      bold: base.bold || r.negrito || undefined,
      italics: base.italics || r.italico || undefined,
      superScript: r.sobrescrito || undefined
    });
  });
}

/* ============================================================================
 * TABELAS
 * ========================================================================== */

const ALINHAMENTO_CELULA = {
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  both: AlignmentType.JUSTIFIED,
  left: AlignmentType.LEFT
};

function criarTabela(item) {
  const linhas = item.linhas || [];
  const larguraDaGrade = (item.grade || []).reduce((a, b) => a + b, 0);
  const fator = larguraDaGrade > 0 ? Math.min(1, DOUTRIN_CONFIG.textWidth / larguraDaGrade) : 1;

  const linha = (estilo) => ({ style: estilo ? BorderStyle.SINGLE : BorderStyle.NONE, size: estilo ? 4 : 0, color: estilo ? '000000' : 'auto' });
  const bordasTabela = item.bordas || {};
  const colunasMaximas = Math.max(1, ...linhas.map((l) => l.reduce((soma, c) => soma + c.span, 0)));
  const grade =
    item.grade && item.grade.length === colunasMaximas
      ? item.grade.map((w) => Math.max(200, Math.round(w * fator)))
      : Array.from({ length: colunasMaximas }, () => Math.floor(DOUTRIN_CONFIG.textWidth / colunasMaximas));

  const rows = linhas.map((celulas) => {
    let coluna = 0;
    const children = celulas.map((celula) => {
      const largura = grade.slice(coluna, coluna + celula.span).reduce((a, b) => a + b, 0);
      coluna += celula.span;

      const bordasCelula = {};
      const nomes = { top: 'top', bottom: 'bottom', left: 'left', right: 'right' };
      Object.keys(nomes).forEach((lado) => {
        if (celula.bordas && celula.bordas[lado] !== undefined) bordasCelula[lado] = linha(celula.bordas[lado]);
      });

      const paragrafos = (celula.paragrafos.length ? celula.paragrafos : [{ runs: [], jc: 'left' }]).map(
        (p) =>
          new Paragraph({
            alignment: ALINHAMENTO_CELULA[p.jc] || AlignmentType.LEFT,
            spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
            children: criarRuns(p.runs, { size: DOUTRIN_CONFIG.size9pt })
          })
      );

      return new TableCell({
        children: paragrafos,
        width: { size: largura, type: WidthType.DXA },
        columnSpan: celula.span > 1 ? celula.span : undefined,
        verticalMerge:
          celula.vmerge === 'restart'
            ? VerticalMergeType.RESTART
            : celula.vmerge === 'continue'
              ? VerticalMergeType.CONTINUE
              : undefined,
        verticalAlign: VerticalAlign.CENTER,
        borders: Object.keys(bordasCelula).length ? bordasCelula : undefined
      });
    });
    return new TableRow({ children, cantSplit: true });
  });

  return new Table({
    rows,
    width: { size: grade.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: grade,
    alignment: AlignmentType.CENTER,
    borders: {
      top: linha(bordasTabela.top !== false),
      bottom: linha(bordasTabela.bottom !== false),
      left: linha(bordasTabela.left !== false),
      right: linha(bordasTabela.right !== false),
      insideHorizontal: linha(bordasTabela.insideH !== false),
      insideVertical: linha(bordasTabela.insideV !== false)
    }
  });
}

/* ============================================================================
 * IMAGENS E GRÁFICOS
 * ========================================================================== */

function criarParagrafoImagem(item) {
  let largura = item.largura || 400;
  let altura = item.altura || 300;
  const escala = Math.min(1, LARGURA_MAXIMA_IMAGEM_PX / largura, ALTURA_MAXIMA_IMAGEM_PX / altura);
  largura = Math.max(20, Math.round(largura * escala));
  altura = Math.max(20, Math.round(altura * escala));

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
    children: [
      new ImageRun({
        data: item.dados,
        type: item.extensao || 'png',
        transformation: { width: largura, height: altura }
      })
    ]
  });
}

/**
 * Encaixa nos tópicos as imagens de gráficos que o usuário anexou na tela.
 * O encaixe é feito pelo NÚMERO: "grafico3.png" -> legenda "Gráfico 3 – ...".
 *  - se existe um marcador "[INSERIR GRÁFICO AQUI]" logo após a legenda, ele é substituído;
 *  - se existe só a legenda, a imagem é colocada logo abaixo dela;
 *  - imagens sem legenda correspondente vão para o fim do último tópico.
 */
function injetarGraficos(topicos, graficos) {
  if (!graficos.length) return topicos;

  const pendentes = graficos.map((g) => ({
    ...g,
    numero: (/(\d+)/.exec(g.nome || '') || [])[1] || null,
    usado: false
  }));

  const REGEX_LEGENDA_GRAFICO = /^(gráfico|grafico|figura|ilustração|ilustracao)\s+(\d+)/i;

  const novosTopicos = topicos.map((topico) => {
    const itens = [];
    let numeroDaLegenda = null;

    (topico.itens || []).forEach((item, indice, todos) => {
      if (item.tipo === 'legenda') {
        const texto = (item.runs || []).map((r) => r.texto || '').join('').trim();
        const m = REGEX_LEGENDA_GRAFICO.exec(texto);
        numeroDaLegenda = m ? m[2] : null;
        itens.push(item);

        const proximo = todos[indice + 1];
        const seguidaDeFigura = proximo && (proximo.tipo === 'grafico-pendente' || proximo.tipo === 'imagem');
        if (numeroDaLegenda && !seguidaDeFigura) {
          const g = pendentes.find((x) => !x.usado && x.numero === numeroDaLegenda);
          if (g) {
            g.usado = true;
            itens.push({ tipo: 'imagem', ...g });
          }
        }
        return;
      }

      if (item.tipo === 'grafico-pendente' && numeroDaLegenda) {
        const g = pendentes.find((x) => !x.usado && x.numero === numeroDaLegenda);
        if (g) {
          g.usado = true;
          itens.push({ tipo: 'imagem', ...g });
          return;
        }
      }

      itens.push(item);
    });

    return { ...topico, itens };
  });

  // Sobras: anexadas ao final do último tópico, com o nome do arquivo como legenda
  const sobras = pendentes.filter((g) => !g.usado);
  if (sobras.length && novosTopicos.length) {
    const ultimo = novosTopicos[novosTopicos.length - 1];
    sobras.forEach((g) => {
      ultimo.itens.push({ tipo: 'legenda', runs: [{ texto: (g.nome || 'Gráfico').replace(/\.[^.]+$/, '') }] });
      ultimo.itens.push({ tipo: 'imagem', ...g });
    });
  }
  return novosTopicos;
}

/* ============================================================================
 * UTILITÁRIOS
 * ========================================================================== */

function criarParagrafoVazio(tamanhoFonte, espacamentoLinha, recuoEsquerda = 0) {
  return new Paragraph({
    indent: recuoEsquerda > 0 ? { left: recuoEsquerda } : undefined,
    spacing: { line: espacamentoLinha, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
    children: [
      new TextRun({
        text: '',
        font: DOUTRIN_CONFIG.font,
        size: tamanhoFonte
      })
    ]
  });
}