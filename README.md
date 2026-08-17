# Ajustes MELI

Ferramenta de página única que lê a planilha **Labor enviado ao MELI** e responde
perguntas diferentes sobre ela, uma em cada aba.

Todo o processamento acontece no navegador. Nenhum dado sai da máquina: não há
servidor, banco nem envio de arquivos.

## As quatro abas

| Aba | Pergunta | Entrega |
|---|---|---|
| **Conciliação Faturas** | O que já foi cobrado está certo? | Ajustes projetados, ou a conferência de duas competências |
| **Fusão de Linhas** | O arquivo bate com o alvo do MELI? | Labor equalizado dia a dia contra o retorno |
| **Extração · Diarista** | Quem foram os diaristas do período? | Uma planilha por operação, no layout de origem |
| **Guia** | — | Resumo conceitual do que cada aba faz |

A aba é escolhida pelo topo da página e também pela URL: `index.html#fusao` abre
direto na segunda.

### Conciliação Faturas

Tem dois modos, escolhidos no topo da aba.

**Projetar ajustes** — uma fatura. Descobre quais ajustes deverão aparecer na
competência seguinte.

```
Importar planilha  →  competência detectada  →  lista de ajustes  →  Exportar Excel
```

**Conciliar duas faturas** — duas competências consecutivas. Para cada pessoa
reconstrói a história de faturamento — o que foi cobrado, o que o corte
congelou, o que se infere ter acontecido depois, que ajuste era devido e o que
de fato apareceu — e classifica o encontro em conciliado, ajuste ausente,
período divergente, FTE divergente, sinal incorreto, duplicado, retroativo sem
origem, identidade ambígua, indeterminado ou revisão manual. Um mesmo registro
pode acumular alertas. Cada apontamento mostra quais das cinco dimensões
comparadas fecharam (identidade, competência, período, sinal e FTE), o nível de
confiança e a razão dele, uma linha do tempo da competência de origem e um
diagnóstico em texto.

```
Fatura N + Fatura N+1  →  apontamentos  →  você decide  →  prévia  →  Fatura Conciliada
```

O app **não altera nada por conta própria**. Toda linha nasce em "manter como
está"; só o que for explicitamente aceito entra na prévia, e a exportação parte
da fatura N+1 para criar um arquivo novo — o original nunca é sobrescrito. O
arquivo gerado leva uma aba `CONCILIAÇÃO` documentando o que foi analisado,
o que você decidiu e o que de fato mudou, inclusive o que optou por não tratar.

### Fusão de Linhas

Compara, dia a dia, o headcount ativo no Labor com a quantidade que o MELI aponta
na aba `Retorno MELI` (`Qtd. PREF` e `Q Pós Comp.`). Onde sobra gente monta um
plano e o aplica — retirar linha, adiar início, encurtar fim, ou pausar e retomar
o contrato preservando GROOT e matrícula — e devolve o Labor corrigido no mesmo
layout do original, com abas `A_INCLUIR` e `REVISAR` para o que depende de
decisão humana. Aceita ainda dois arquivos opcionais: a base de diaristas, para
preencher os dias em falta, e o HCM Report, para achar quem está na base do MELI
e não tem cobertura nenhuma no Labor.

```
Soltar planilha  →  conciliação dia a dia  →  plano aplicado  →  Baixar Labor ajustado
```

### Extração · Diarista

Assistente de três passos sobre o controle de diaristas — o arquivo único com uma
aba por operação (Pouso Alegre SVC e XD, Poços de Caldas, Varginha, Divinópolis,
Patos de Minas). Filtra pela **data de solicitação**, opcionalmente por
**solicitante** (ID Logistics ou MELI), e resolve os **Groot IDs repetidos** em um
de três modos: uma linha por pessoa no período, uma por pessoa por dia, ou nenhuma
remoção. Exporta uma planilha por operação — ou as seis num arquivo só — no layout
de origem, nomeadas como `Varginha - Diaristas - Agosto.26.xlsx`.

```
Soltar planilha  →  período + filtros  →  placar por operação  →  Baixar .xlsx
```

## A regra em uma frase

O faturamento congela um retrato do quadro no dia 15 e **projeta** a cobrança de
cada pessoa, da sua data de início até o fim do mês — quem entrou em 07/05 e
estava ativo em 15/05 foi cobrado de 07/05 a 31/05, e não o mês cheio. O que
acontece **depois** do congelamento — uma saída no dia 20, uma entrada no dia
22 — não cabe mais naquela fatura e vira ajuste na seguinte.

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
index.html            as quatro abas e a ordem de carga dos scripts
css/styles.css        estilos das quatro abas
js/config.js          nomes de abas, colunas aceitas, dia de corte
js/dates.js           datas como inteiro AAAAMMDD, imune a fuso horário
js/engine.js          motor de regras: valida, classifica, dedupe
js/competence.js      detecção automática da competência
js/import.js          leitura da planilha (SheetJS)
js/ui.js              renderização da tabela e do detalhe
js/export.js          geração do Excel herdando o estilo do arquivo original
js/app.js             inicialização e eventos do modo "projetar ajustes"
js/billing.js               reconstrói a cobrança e classifica cada linha
js/competence-source.js     competência por ordem de evidência, com fonte visível
js/reconciliation.js        motor de conciliação entre duas faturas
js/reconciliation-ui.js     uploads duplos, apontamentos, decisões e prévia (IIFE)
js/reconciliation-export.js clona a fatura N+1 e aplica só o que foi aceito
js/fusao.js           aba Fusão de Linhas, inteira (IIFE)
js/extracao.js        aba Extração · Diarista, inteira (IIFE)
js/tabs.js            navegação entre as abas principais
tests/                testes do motor e da conciliação, sem dependências
docs/REGRAS.md        regras de negócio detalhadas
```

Os arquivos são carregados como scripts clássicos, na ordem declarada no
`index.html`. Não há build, bundler nem instalação: é o código que roda.

### Por que boa parte dos módulos é IIFE

As ferramentas nasceram como páginas independentes e colidiam ao dividir o mesmo
documento. O modo "projetar ajustes" está espalhado por sete arquivos que
compartilham o escopo global e continua assim; os demais foram fechados em IIFEs.

`js/fusao.js` precisava disso porque também define uma função `render`; publica em
`window.Fusao` apenas o que os handlers inline do HTML chamam. `js/extracao.js` já
nascera isolado. `js/reconciliation-ui.js` segue a mesma regra e publica
`window.Recon` — já `js/reconciliation.js` fica no escopo global de propósito,
porque é motor puro e os testes o carregam junto de `engine.js`.

Os `id` dos elementos levam prefixo — `fz-` na fusão, `ex-` na extração — porque
as abas coexistem no mesmo documento: fusão e conciliação tinham as duas um
`#result` e um `#btnExport`, e extração e conciliação tinham as duas um
`#fileInput`.

### Bibliotecas

Carregadas por CDN, sem instalação:

- **SheetJS (xlsx)** — leitura da planilha importada.
- **ExcelJS** — escrita do arquivo exportado, porque a versão comunitária do
  SheetJS não grava formatação, e a exportação precisa preservar o visual MELI.

## Testes

```bash
npm test          # roda os dois arquivos de teste
```

Não há dependências para instalar. Os testes carregam os mesmos arquivos de
`js/` que o navegador usa, num contexto compartilhado.

`tests/engine.test.js` cobre os casos obrigatórios de regra, os limites da
competência, a proteção contra ajuste duplicado e a detecção automática da
competência.

`tests/reconciliation.test.js` cobre a conciliação entre duas faturas: as
classificações de status, a reconstrução da cobrança original (o corte projeta
da admissão, não cria mês cheio), a classificação de linha normal × retroativa
nos dois sentidos de sinal, a competência por ordem de evidência, a
normalização de identificadores, a checagem de sequência, os alertas múltiplos,
a diferença exata em dias, a identidade da linha criada e a garantia de que
nenhum apontamento nasce aceito.

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
