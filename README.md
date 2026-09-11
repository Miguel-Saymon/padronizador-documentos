# Padronizador de Documentos

Aplicação web estática para conversão e padronização automática de documentos (`.docx` e `.pdf`) para o modelo de formatação do Doutrin.

Todo o processamento acontece **100% no navegador do usuário** — nenhum arquivo é enviado a um servidor. Isso garante privacidade (o documento nunca sai do computador de quem usa) e custo zero de infraestrutura (o site inteiro é estático, hospedado no GitHub Pages).

**[🔗 Usar agora](https://SEU_USUARIO.github.io/padronizador-documentos/)**

## Como funciona

1. Você arrasta (ou seleciona) um arquivo `.docx` ou `.pdf`.
2. A ferramenta lê o conteúdo e identifica automaticamente: título, autor(es), resumo, palavras-chave, tópicos numerados, citações longas e referências bibliográficas.
3. Um novo arquivo `.docx` é gerado já formatado no padrão do Doutrin (fonte, tamanho, espaçamento, recuos e alinhamento corretos).
4. Você baixa o resultado.

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
  index.html          # interface (área de arrastar-e-soltar)
  main.js             # liga a interface aos módulos de leitura/geração
  style.css
  modules/
    reader.js          # lê .docx (mammoth) e .pdf (pdf.js) e extrai a estrutura do texto
    generator.js       # monta o .docx final no padrão Doutrin (usa a lib docx)
```

- **Leitura de `.docx`**: usa o [mammoth](https://github.com/mwilliamson/mammoth.js) para converter o arquivo em HTML, preservando parágrafos, títulos (`Heading 1/2/3`) e notas de rodapé.
- **Leitura de `.pdf`**: usa o [pdf.js](https://mozilla.github.io/pdf.js/) e reconstrói a estrutura do texto a partir de pistas geométricas (posição, tamanho de fonte, negrito) — já que um PDF não guarda informação semântica como um `.docx`.
- **Geração do `.docx` final**: usa a lib [docx](https://github.com/dolanmiu/docx) para montar o documento programaticamente, já com a formatação do padrão Doutrin aplicada.

## ⚠️ Limitações conhecidas

- **PDF é sempre mais frágil que DOCX** como entrada: a extração depende de heurísticas de geometria (posição do texto, tamanho de fonte, negrito), então PDFs com layout muito diferente do testado podem não ser lidos perfeitamente. Sempre revise o resultado antes de publicar.
- Em casos raros, uma nota de rodapé muito longa pode ocupar uma faixa vertical da página que se sobrepõe à de uma citação longa, causando confusão entre as duas categorias. Vale conferir manualmente notas de rodapé extensas.
- A **data de publicação e o link (URL)** da citação ABNT no topo do documento não podem ser adivinhados pela ferramenta — eles saem preenchidos como `[DATA DE PUBLICAÇÃO]` e `[LINK]`, para completar manualmente antes da publicação.

## 🛠️ Tecnologias

- [Vite](https://vitejs.dev/) — build e servidor de desenvolvimento
- [mammoth](https://github.com/mwilliamson/mammoth.js) — leitura de `.docx`
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