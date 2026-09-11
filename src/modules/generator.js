import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  LineRuleType,
  convertInchesToTwip
} from 'docx';

const DOUTRIN_CONFIG = {
  font: 'Helvetica Neue',
  
  // No Word: 240 = Simples / 360 = 1,5 Linhas
  lineSpacingSimple: 240, 
  lineSpacingOneAndHalf: 360, 
  
  size8pt: 16,
  size10pt: 20,
  size12pt: 24,
  size14pt: 28,

  marginBodyLeft: 567, // 1 cm
  citationIndentTotal: 3969, // 7 cm (4 cm recuo + 3 cm margem)
  
  pageMargin: convertInchesToTwip(2.5 / 2.54)
};

export async function gerarDocxPadronizado(dados) {
  const elementos = [];

  // PARTE 1: INICIAL (L1 a L5) - Espaçamento Simples, 8pt
  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size8pt, DOUTRIN_CONFIG.lineSpacingSimple));

  if (dados.creditoPublicacao) {
    elementos.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({
            text: dados.creditoPublicacao,
            font: DOUTRIN_CONFIG.font,
            size: DOUTRIN_CONFIG.size8pt,
            bold: true
          })
        ]
      })
    );
  }

  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size8pt, DOUTRIN_CONFIG.lineSpacingSimple));

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

  // PARTE 2: AUTOR (L6 a L9) - Espaçamento 1,5 linha, 14pt Bold Italic
  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

  elementos.push(
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
      children: [
        new TextRun({
          text: dados.autor || "Nome Sobrenome",
          font: DOUTRIN_CONFIG.font,
          size: DOUTRIN_CONFIG.size14pt,
          bold: true,
          italics: true
        })
      ]
    })
  );

  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));
  elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

  // PARTE 3: SUMÁRIO (L10) - Espaçamento Simples, 12pt
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

  // PARTE 4: RESUMO
  if (dados.resumoTexto) {
    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

    elementos.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({
            text: "RESUMO",
            font: DOUTRIN_CONFIG.font,
            size: DOUTRIN_CONFIG.size14pt,
            bold: true
          })
        ]
      })
    );

    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));

    elementos.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        children: [
          new TextRun({
            text: dados.resumoTexto,
            font: DOUTRIN_CONFIG.font,
            size: DOUTRIN_CONFIG.size12pt
          })
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

    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));
    elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf));
  }

  // PARTE 5: CORPO DO TEXTO & TÓPICOS
  if (dados.topicos && dados.topicos.length > 0) {
    dados.topicos.forEach((topico) => {
      elementos.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          indent: { left: DOUTRIN_CONFIG.marginBodyLeft },
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

      elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf, DOUTRIN_CONFIG.marginBodyLeft));

      topico.paragrafos.forEach((p) => {
        if (p.eCitacaoLonga) {
          elementos.push(
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              indent: { left: DOUTRIN_CONFIG.citationIndentTotal },
              spacing: { line: DOUTRIN_CONFIG.lineSpacingSimple, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
              children: [
                new TextRun({
                  text: p.texto,
                  font: DOUTRIN_CONFIG.font,
                  size: DOUTRIN_CONFIG.size10pt
                })
              ]
            })
          );
        } else {
          elementos.push(
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              indent: { left: DOUTRIN_CONFIG.marginBodyLeft },
              spacing: { line: DOUTRIN_CONFIG.lineSpacingOneAndHalf, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
              children: [
                new TextRun({
                  text: p.texto,
                  font: DOUTRIN_CONFIG.font,
                  size: DOUTRIN_CONFIG.size12pt
                })
              ]
            })
          );
        }
      });

      elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf, DOUTRIN_CONFIG.marginBodyLeft));
      elementos.push(criarParagrafoVazio(DOUTRIN_CONFIG.size12pt, DOUTRIN_CONFIG.lineSpacingOneAndHalf, DOUTRIN_CONFIG.marginBodyLeft));
    });
  }

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

function criarParagrafoVazio(tamanhoFonte, espacamentoLinha, recuoEsquerda = 0) {
  return new Paragraph({
    indent: recuoEsquerda > 0 ? { left: recuoEsquerda } : undefined,
    spacing: { line: espacamentoLinha, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
    children: [
      new TextRun({
        text: "",
        font: DOUTRIN_CONFIG.font,
        size: tamanhoFonte
      })
    ]
  });
}