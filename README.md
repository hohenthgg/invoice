# Ajustes MELI

Ferramenta de página única que lê a planilha **Labor enviado ao MELI** e responde
perguntas diferentes sobre ela, uma em cada aba.

Todo o processamento acontece no navegador. Nenhum dado sai da máquina: não há
servidor, banco nem envio de arquivos.

## As cinco abas

| Aba | Pergunta | Entrega |
|---|---|---|
| **Conciliação Faturas** | O que já foi cobrado está certo? | Ajustes projetados, ou a conferência de duas competências |
| **Fusão de Linhas** | O arquivo bate com o alvo do MELI? | Labor equalizado dia a dia contra o retorno |
| **Extração · Diarista** | Quem foram os diaristas do período? | Uma planilha por operação, no layout de origem |
| **Calcular ABS** | O absenteísmo ficou dentro do range? | Absenteísmo antes e pós diaristas, com o Excel gerencial |
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
remoção. A deduplicação é **global** — a mesma pessoa-dia em duas operações é uma
só, e o que foi descartado aparece num painel com GROOT, data e de onde veio. Exporta uma planilha por operação — ou as seis num arquivo só — no layout
de origem, nomeadas como `Varginha - Diaristas - Agosto.26.xlsx`.

As colunas são localizadas **pelo próprio cabeçalho**, não por posição a partir do
GROOT. A leitura antiga travava na primeira linha que contivesse "groot" e contava
offsets dali: em Varginha, uma legenda com `GROOT ID` numa coluna vazia acima do
cabeçalho real fez 119 linhas saírem com GROOT, NOME, CARGO e ESCALA **todos**
vazios. Agora o cabeçalho é a linha que reconhece mais colunas, e entre dois
candidatos para a mesma coluna ganha o que tem dado embaixo.

A saída sai **sem coluna vazia e sem célula obviamente preenchível**: coluna sem um
único valor na filial não é escrita; célula em branco numa coluna que só tem um
valor no arquivo é preenchida com ele — havendo dois valores distintos, fica em
branco, porque escolher seria inventar. Identificadores saem limpos de chaves,
colchetes e barras (`{2499441}` → `2499441`), o que **também corrige a
deduplicação**: o embrulhado e o limpo passam a ser a mesma pessoa. Todos esses
ajustes viram tópico no painel de revisão — nada é mexido em silêncio.

O formato de saída é o de `Modelo diaristas.xlsx` — oito colunas, copiadas dele
célula a célula (larguras, os dois tons de cinza do cabeçalho, fonte Aptos Narrow,
alinhamentos, sem congelar linha e sem autofiltro, porque o modelo não tem):

`MÊS SOLICITAÇÃO` · `DATA SOLICITAÇÃO` · `SOLICITANTE` · `EMPRESA DIARISTA` ·
`GROOT ID` · `NOME` · `CARGO` · `ESCALA`

**No modelo, `ESCALA` é o HORÁRIO** (`03:00 07:00 08:00 12:48` — entrada, início e
fim do intervalo, saída), não o `6x1`/`XD` do SIGO. E a coluna ESCALA do SIGO fala
uma língua por filial, então a resolução é **por linha** (`resolverEscala`,
`js/config.js`), nesta prioridade:

1. coluna explícita de horário ("Escala Horário" / "Escala Natural"), se preenchida;
2. horário escrito na própria ESCALA passa **verbatim** — Divinópolis
   (`01:00 04:00 05:00 09:20`) e Patos de Minas (`00:30 as 09:18`);
3. turno vira horário pela tabela da operação (`ESCALA_TOKEN_HORARIO`), levantada
   nas faturas 3PL: Varginha `AM` → `03:00 07:00 08:00 11:20` e `PM` →
   `10:00 15:00 16:00 19:48`; Pouso `svc` → horário SMG3 e `xd` → BRXMG3; Poços `AM`;
4. célula vazia recebe o padrão da operação (`ESCALA_HORARIO_PADRAO`);
5. turno **sem horário levantado** (`SD`, `FULL` no SVC, `PM` em Poços) sai como
   está e vira tópico de revisão — escrever o horário de outro turno no lugar
   seria pôr um dado errado com cara de certo.

Antes do passo 3 há uma regra de **pertencimento**, que é de negócio e não de
horário (`ESCALA_TURNOS_DA_OPERACAO`): o **SVC mistura-se** com `SD`, `FULL` e
`XD`; o **XD é apenas XD**. Um `SD` ou `FULL` na aba XD não é falta de horário —
é registro na aba errada. Sai como está e vira tópico **grave**, com uma linha por
ocorrência (nome, GROOT e data), porque cada diária precisa ser conferida
individualmente. Operação sem regra declarada não policia turno nenhum.

```
Soltar planilha  →  período + filtros  →  placar por operação  →  Baixar .xlsx
```

### Calcular ABS

Compara, dia a dia, o **Quadro S&OP (sem over)** com quem esteve **presente** no
período da fatura, e compensa o déficit de cada dia com os **diaristas
efetivamente solicitados naquele dia** — contados uma vez só, e apenas quando têm
Groot ID. O abate é limitado ao próprio déficit: um dia nunca fica positivo por
sobra de diarista.

Devolve o absenteísmo **antes** e **pós compensação** contra o **range contratual
de 2,5%**, a lista dos dias críticos e um Excel com a aba `Diaristas`, um
`Unificado` por operação e um `Resumo` gerencial com fórmulas vivas — o
`COUNTIFS` dos diaristas aponta para a própria aba `Diaristas`, então a planilha
recalcula sozinha se alguém editar. Além do unificado há **um download por
filial** — o arquivo que vai para cada gerente sem levar junto o resto da rede.

A linha **Contratado** do resumo vem da linha `Contratado` do bloco de headcount
da própria planilha mensal — é dado de origem, não algo derivável da grade:
contar marcas não-DF dava 179 num dia em que a origem diz 137. Só no modelo
antigo, sem o bloco, ela continua derivada — e rotulada como derivada.

Dias sem S&OP, ou sem lançamento na base de absenteísmo, ficam de fora do
percentual em vez de entrar como zero.

```
3 bases  →  período  →  filiais e operações  →  antes × pós  →  Baixar .xlsx
```

Precisa de **três bases**: duas planilhas mensais de absenteísmo (abas por
operação, ex. `PAXD Jul`) e a base SIGO de diaristas.

O **S&OP diário vem da própria aba da operação**, na linha `Esperado` do bloco de
headcount que precede a grade de pessoas:

```
linha  1  S&OP — HEADCOUNT DIÁRIO · Pouso Alegre · Julho/2026
linha  2  Métrica  …  1 (Qua) | 2 (Qui) | …
linha  3  Esperado …      88  |     88  | …   ← o S&OP do dia
linha 10  COLABORADORES — PRESENÇAS E FALTAS
linha 11  Mat. | Groot ID | Nome | … | 1 Qua | 2 Qui | …
```

Nada disso é posição fixa: `localizarBlocos` procura a linha `Mat./Nome` e, se
existir, a linha `Esperado` com o cabeçalho de dias logo acima — por isso o
**modelo antigo continua valendo**, com a grade começando na linha 1 e o S&OP
vindo do arquivo de headcount, que virou **opcional**. Cada aba é uma operação,
então vira uma seção com o nome do turno (`SVC`, `XD`, `SD`, `FULL`) e uma filial
sem turno vira a seção `TOTAL`. Quando as duas fontes trazem o mesmo dia, o
arquivo avulso prevalece — carregá-lo é um ato explícito — e o seletor diz de onde
o número veio.

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
index.html            as cinco abas e a ordem de carga dos scripts
css/styles.css        estilos das cinco abas
js/identity.js        normalização de GROOT e a chave pessoa-dia, para o app inteiro
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
js/extraction-dedup.js  deduplicação pessoa-dia, global e testável
js/abs.js             aba Calcular ABS, inteira (IIFE)
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

`js/fusao.js` e `js/abs.js` precisavam disso porque também definem uma função
`render`; a fusão publica em `window.Fusao` o que os handlers inline chamam, e a
de ABS não publica nada, porque não tem handler inline. `js/extracao.js` já
nascera isolado. `js/reconciliation-ui.js` segue a mesma regra e publica
`window.Recon` — já `js/reconciliation.js` fica no escopo global de propósito,
porque é motor puro e os testes o carregam junto de `engine.js`.

Os `id` dos elementos levam prefixo — `fz-` na fusão, `ex-` na extração — porque
as abas coexistem no mesmo documento: fusão e conciliação tinham as duas um
`#result` e um `#btnExport`, e extração e conciliação tinham as duas um
`#fileInput`. A aba de ABS usa `abs-` pela mesma razão, e suas classes genéricas
(`.card`, `.top`, `.drop`, `.row`, `.note`) ficam sob `#panel-abs` porque existem
em outras abas com intenção diferente.

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

`tests/export.test.js` cobre a geração da Fatura Conciliada, com um stub de
worksheet que reproduz o deslocamento de índices do `spliceRows`. Prova que uma
linha nova nunca herda identidade de outra pessoa, mesmo depois de várias
remoções acima dela.

`tests/extraction.test.js` cobre a deduplicação pessoa-dia da Extração ·
Diarista: mesmo GROOT normalizado e mesma data são uma única pessoa-dia,
independentemente de operação, aba ou arquivo de origem.

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
