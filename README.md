# Ajustes MELI

Ferramenta de página única que lê a planilha **Labor enviado ao MELI**, detecta as
movimentações ocorridas depois do corte do dia 15 e gera os ajustes que precisam
entrar na próxima fatura.

Todo o processamento acontece no navegador. Nenhum dado sai da máquina: não há
servidor, banco nem envio de arquivos.

```
Importar planilha  →  competência detectada  →  lista de ajustes  →  Exportar Excel
```

## A regra em uma frase

O faturamento congela um retrato do quadro no dia 15. Quem está ativo nesse dia é
cobrado pela competência inteira. O que acontece **depois** do congelamento —
uma saída no dia 20, uma entrada no dia 22 — não cabe mais naquela fatura e vira
ajuste na seguinte.

| Situação | Ajuste |
|---|---|
| Ativo em 15/08, desligado em 20/08 | **Descontar** 21/08 → 31/08 (11 dias) |
| Admitido em 20/08 | **Acrescentar** 20/08 → 31/08 (12 dias) |
| Admitido 20/08 e desligado 25/08 | **Acrescentar** 20/08 → 25/08 (6 dias) |
| Desligado em 15/08 (data fim é inclusiva) | **Descontar** 16/08 → 31/08 |
| Desligado em 14/08 | Nenhum — já não estava no retrato |
| Admissão antiga, sem saída | Nenhum — cobrança integral correta |

`FTE do ajuste = dias do ajuste ÷ dias corridos da competência × % rateio`,
negativo para desconto e positivo para acréscimo.

As regras completas, com os casos de borda, estão em [`docs/REGRAS.md`](docs/REGRAS.md).

## Como usar

Abra `index.html` no navegador — clicando duas vezes no arquivo já funciona.
Depois clique em **Importar planilha** e, quando a lista aparecer, em
**Exportar Excel**. Nada mais precisa ser preenchido.

Para servir localmente (recomendado, evita restrições de `file://` em alguns navegadores):

```bash
python3 -m http.server 8080
# abra http://localhost:8080
```

### Publicando no GitHub Pages

Em **Settings → Pages**, escolha `Deploy from a branch`, branch `main`, pasta `/ (root)`.
O arquivo `.nojekyll` já está no repositório para o Pages servir tudo sem processar.

## Estrutura

```
index.html            marcação da página e ordem de carga dos scripts
css/styles.css        estilos
js/config.js          nomes de abas, colunas aceitas, dia de corte
js/dates.js           datas como inteiro AAAAMMDD, imune a fuso horário
js/engine.js          motor de regras: valida, classifica, dedupe
js/competence.js      detecção automática da competência
js/import.js          leitura da planilha (SheetJS)
js/ui.js              renderização da tabela e do detalhe
js/export.js          geração do Excel herdando o estilo do arquivo original
js/app.js             inicialização e eventos
tests/                testes do motor, sem dependências
docs/REGRAS.md        regras de negócio detalhadas
```

Os arquivos são carregados como scripts clássicos, na ordem declarada no
`index.html`. Não há build, bundler nem instalação: é o código que roda.

### Bibliotecas

Carregadas por CDN, sem instalação:

- **SheetJS (xlsx)** — leitura da planilha importada.
- **ExcelJS** — escrita do arquivo exportado, porque a versão comunitária do
  SheetJS não grava formatação, e a exportação precisa preservar o visual MELI.

## Testes

```bash
npm test          # ou: node tests/engine.test.js
```

Não há dependências para instalar. Os testes carregam os mesmos arquivos de
`js/` que o navegador usa, num contexto compartilhado, e cobrem os casos
obrigatórios de regra, os limites da competência, a proteção contra ajuste
duplicado e a detecção automática da competência.

## Detecção da competência

Em ordem de força da evidência:

1. Se a aba `Retorno MELI` tiver `Data Trab.` com um mês concentrando **≥70%**
   das datas válidas, esse mês é a competência.
2. Caso contrário, o app localiza o mês-âncora — o mais recente com volume real
   de movimentações — e avalia apenas ele e os dois anteriores. Desligamentos
   antigos ficam fora da conta, e uma data solta (erro de digitação, data
   futura) não forma um mês significativo.
3. Se dois meses ficarem próximos, aparece **Alterar** ao lado da competência
   para escolher manualmente. A linha da competência é sempre clicável.

## Exportação

O arquivo gerado (`Ajustes MELI - Agosto 2026.xlsx`, aba `Ajustes MELI`) mantém
as dez colunas originais da planilha, copiadas da linha importada, e acrescenta
à direita apenas o bloco do ajuste: tipo, início, fim, dias, FTE, competência de
origem, competência de aplicação e motivo.

A formatação não é imitada: o exportador lê o estilo real do arquivo importado
(fonte, preenchimento, bordas, alinhamento, formatos de data e percentual) e o
reaplica célula a célula. Como o estilo vem da própria linha de origem, cores
associadas a categorias — regime de contrato, cargo, escala — são preservadas
automaticamente. Só entram no Excel os ajustes que permanecem marcados.

## Planilha de entrada

O app procura a aba `Labor enviado ao MELI` e localiza as colunas **pelo nome do
cabeçalho**, não pela posição:

`GROOT ID` · `NOME` · `MATRICULA` · `REGIME DE CONTRATO` · `CARGO` ·
`DATA DE INÍCIO` · `DATA FIM` · `% RATEIO` · `DIAS TRABALHADOS X FOLGA` · `ESCALA`

Datas são aceitas como data do Excel, número de série, `DD/MM/AAAA` ou
`AAAA-MM-DD`, e normalizadas internamente para inteiro `AAAAMMDD` — comparação
de datas nunca é feita como texto nem sofre com fuso horário.

Linhas com dados inconsistentes (data fim anterior à de início, GROOT vazio,
rateio inválido) não interrompem a análise: aparecem agrupadas abaixo da tabela
em "registros não puderam ser analisados".

## Licença

MIT — veja [LICENSE](LICENSE).
