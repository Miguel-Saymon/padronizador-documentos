# Padronizador de Documentos

Aplicação web estática para conversão e padronização automática de documentos (`.docx` e `.pdf`) para o modelo de formatação do Doutrin.

Todo o processamento acontece **100% no navegador do usuário** — nenhum arquivo é enviado a um servidor. Isso garante privacidade (o documento nunca sai do computador de quem usa) e custo zero de infraestrutura (o site inteiro é estático, hospedado no GitHub Pages).

**[🔗 Usar agora](https://Miguel-Saymon.github.io/padronizador-documentos/)**

## Como funciona

1. Você arrasta (ou seleciona) um arquivo `.docx` ou `.pdf`.
2. A ferramenta lê o conteúdo e identifica automaticamente: título, autor(es), resumo, palavras-chave, tópicos numerados, citações longas, notas de rodapé, tabelas, imagens e referências bibliográficas.
3. Se quiser, você marca **"Artigo republicado"** (para incluir o crédito da publicação original) e/ou **anexa imagens de gráficos**.
4. Um novo arquivo `.docx` é gerado já formatado no padrão do Doutrin.
5. Você baixa o resultado.

## ✨ O que a ferramenta aplica

- Fonte Helvetica Neue, página A4 com margens de 2,5 cm.
- Autores um por linha (somente o nome, sem a universidade), em 14 pt negrito-itálico, alinhados à direita.
- Corpo do texto justificado, espaçamento 1,5 e **recuo de 1 cm apenas na primeira linha** de cada parágrafo.
- Títulos de tópicos sem recuo, em 14 pt negrito.
- **Citações longas** com recuo de 4 cm, espaçamento simples, fonte 10 e uma linha em branco antes e depois.
- **Notas de rodapé reais** do Word (o texto da nota não vai para o corpo). Notas "soltas" no meio do texto, comuns em arquivos convertidos de PDF, também viram notas de rodapé.
- **Referências** separadas corretamente, com reparo de espaços colados.
- **Tabelas e imagens** do original são copiadas para o documento final.
- **Crédito de republicação** ("Texto originalmente publicado na ..., sob a licença ...") apenas quando a opção é marcada; a revista é sugerida a partir de "Como citar este artigo" e a licença é editável.
- **Gráficos anexados**: imagens PNG, JPG, GIF ou BMP são encaixadas pelo número da legenda. Um arquivo chamado `grafico3.png` vai para a legenda "Gráfico 3 – ...". Se não houver legenda correspondente, a imagem vai para o fim do último tópico.

## 🚀 Como executar localmente

1. **Instale as dependências:**
   ```bash
   npm install
   ```

2. **Rode o servidor de desenvolvimento:**
   ```bash
   npm run dev
   ```
   O terminal vai mostrar um endereço local (algo como `http://localhost:5173`) — abra no navegador.

3. **Gere a versão de produção** (opcional, usado pelo deploy automático):
   ```bash
   npm run build
   ```
   Os arquivos finais ficam em `dist/`.

## 🧱 Estrutura do projeto

```
src/
  index.html          # interface (arrastar-e-soltar, opções de republicação e gráficos)
  main.js             # liga a interface aos módulos de leitura/geração
  style.css
  modules/
    reader.js          # lê .docx (JSZip + XML do Word) e .pdf (pdf.js) e extrai a estrutura do texto
    generator.js       # monta o .docx final no padrão Doutrin (usa a lib docx)
```

- **Leitura de `.docx`**: abre o arquivo com o [JSZip](https://stuk.github.io/jszip/) e interpreta diretamente o XML do Word (estilos, recuos, numeração, runs em negrito/itálico/sobrescrito, tabelas, imagens e notas de rodapé). Isso permite distinguir citações (pelo recuo), títulos, legendas e notas, o que não era possível ao converter para HTML.
- **Leitura de `.pdf`**: usa o [pdf.js](https://mozilla.github.io/pdf.js/) (carregado via CDN) e reconstrói a estrutura do texto a partir de pistas geométricas (posição, tamanho de fonte) — já que um PDF não guarda informação semântica como um `.docx`.
- **Geração do `.docx` final**: usa a lib [docx](https://github.com/dolanmiu/docx) para montar o documento programaticamente, já com a formatação do padrão Doutrin aplicada.

## ⚠️ Limitações conhecidas

- **PDF é sempre mais frágil que DOCX** como entrada: a extração depende de heurísticas de geometria, então PDFs com layout muito diferente do testado podem não ser lidos perfeitamente. Além disso, **PDF não traz notas de rodapé, tabelas nem gráficos de forma confiável** — a tela exibe um aviso e esses elementos devem ser inseridos manualmente.
- **Gráficos vetoriais do Word** (desenhados como grupos de formas) não podem ser copiados. No lugar sai um marcador amarelo `[INSERIR GRÁFICO AQUI]`; anexe a imagem do gráfico na tela (nomeada com o número da legenda) ou insira-o manualmente.
- Imagens em formatos como EMF/WMF/SVG não são suportadas e também viram marcador.
- As **minibiografias dos autores** depois das referências são descartadas.
- A **data de publicação e o link (URL)** da citação ABNT no topo do documento não podem ser adivinhados quando o original não traz "Como citar este artigo" — eles saem como `[DATA DE PUBLICAÇÃO]` e `[LINK]`, para completar manualmente antes da publicação.
- O estilo dos marcadores de lista (ex.: `(a)` ou `a)`) segue o que está definido no `numbering.xml` do arquivo original.
- Sempre revise o resultado antes de publicar.

## 🛠️ Tecnologias

- [Vite](https://vitejs.dev/) — build e servidor de desenvolvimento
- [JSZip](https://stuk.github.io/jszip/) — leitura do conteúdo interno do `.docx`
- [pdf.js](https://mozilla.github.io/pdf.js/) — leitura de `.pdf`
- [docx](https://github.com/dolanmiu/docx) — geração do `.docx` final
- JavaScript puro (sem framework de UI) + HTML/CSS

## 🚢 Deploy

O deploy é automático: qualquer push na branch `main` dispara o workflow em `.github/workflows/deploy.yml`, que builda o projeto e publica o conteúdo de `dist/` no GitHub Pages.

## 🤝 Contribuindo

Contribuições são bem-vindas! Para propor uma mudança:

1. Faça um fork do repositório
2. Crie uma branch (`git checkout -b minha-melhoria`)
3. Commit suas mudanças (`git commit -m 'Descrição da mudança'`)
4. Push para a sua branch (`git push origin minha-melhoria`)
5. Abra um Pull Request

Para bugs ou sugestões, abra uma [issue](../../issues).

## 📄 Licença

Este projeto está licenciado sob a licença MIT — veja o arquivo [LICENSE](LICENSE) para mais detalhes.