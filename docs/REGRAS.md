# Regras de negócio

Este documento descreve o que o motor faz e por quê. Ele é a referência para
mudanças em `js/engine.js` — se uma alteração de código contradisser algo aqui,
uma das duas coisas está errada.

## O conceito

A apuração trabalha com três datas diferentes, e confundi-las é a origem de
quase todo erro de conferência:

- **Período de folha:** dia 16 do mês anterior até o dia 15 do mês corrente.
- **Competência:** dia 01 até o último dia do mês.
- **Data de corte (snapshot):** dia 15.

O quadro existente no dia 15 é congelado e serve de base para a competência
inteira. A pergunta que o motor responde **não** é "quantos dias essa pessoa
trabalhou no mês", e sim:

> O que foi congelado no retrato do dia 15, e qual fato posterior ao
> congelamento precisa ser compensado na apuração seguinte?

## Ativo no snapshot

```
DATA DE INÍCIO ≤ dia 15
E (DATA FIM vazia OU DATA FIM ≥ dia 15)
```

`DATA FIM` é **inclusiva**: é o último dia em que a pessoa permanece ativa. Quem
tem data fim 20/08 esteve ativo em 20/08, e a ausência começa em 21/08.

### O corte congela uma projeção, não cria mês cheio

Estando ativa no dia 15, a pessoa é apurada da sua **data de início** até o fim
da competência — não do dia 01. O snapshot congela o que se sabia naquele dia e
**projeta** a cobrança até o fim do mês; ele não transforma todo mundo em
31/31.

| Situação em maio | Cobrança reconstruída |
|---|---|
| Início 01/05, ativa em 15/05 | 01/05 → 31/05 (31 dias) |
| Início 07/05, ativa em 15/05 | **07/05 → 31/05 (25 dias)**, não 31/31 |
| Início 07/05, saída 12/05 já conhecida no corte | 07/05 → 12/05 (6 dias) |
| Início 01/05, desligamento 20/05 reconhecido depois do corte | 01/05 → 31/05 — e desconto de 21/05 a 31/05 na competência seguinte |
| Início 16/05 | nada — não entrou na fotografia do corte |

Não se calcula pró-rata de saída posterior ao corte dentro da própria
competência: isso é justamente o que vira ajuste na seguinte.

## Os dois ajustes

### Desconto — desligamento pós-corte

A pessoa estava no retrato e entrou integralmente na competência, mas saiu antes
do fim do mês. Os dias posteriores à saída foram computados a mais.

```
início do ajuste = DATA FIM + 1 dia
fim do ajuste    = último dia da competência
```

### Acréscimo — admissão pós-corte

A pessoa não estava no retrato, logo não foi computada, mas trabalhou parte da
competência.

```
início do ajuste = DATA DE INÍCIO
fim do ajuste    = DATA FIM, ou o último dia da competência se ainda ativa
```

O caso em que alguém **entra e sai entre dois cortes** é a razão pela qual o
motor nunca compara apenas dois retratos: essa pessoa não aparece no snapshot
anterior nem no seguinte, e mesmo assim houve uso de mão de obra. A detecção é
sempre feita linha a linha, sobre `DATA DE INÍCIO` e `DATA FIM`.

## Cálculo

```
FTE do ajuste = dias do ajuste ÷ dias corridos da competência × % rateio
```

Negativo para desconto, positivo para acréscimo. Agosto tem 31 dias, então 11
dias de desconto com rateio integral resultam em −0,3548; com rateio 50%,
−0,1774.

O período do ajuste nunca extrapola a competência: nem antes do dia 01, nem
depois do último dia do mês. Ajustes de zero dia ou menos são descartados.

## Casos que não geram ajuste

- Admissão anterior ao dia 15 e sem desligamento: apuração integral correta.
- Desligamento antes do dia 15: a pessoa não entrou no retrato, então não há o
  que compensar nesta competência.
- Desligamento no último dia do mês ou depois: nada foi computado a mais.

Nenhum desses aparece na tela. A interface mostra apenas ação necessária.

## Proteções contra ajuste duplicado

- Duas linhas que produziriam o mesmo ajuste (mesma pessoa, mesmo tipo, mesmo
  período) geram **um** ajuste, com aviso.
- Duas linhas da mesma pessoa com períodos **sobrepostos** do mesmo tipo não são
  resolvidas automaticamente: vão para a lista de registros não analisados, para
  decisão humana.
- Períodos distintos da mesma pessoa são mantidos, com aviso de código de
  colaborador ou matrícula repetidos.

## Erros de dados

Não se misturam com ajustes. São listados à parte e não entram na exportação:

- código do colaborador ou matrícula vazios
- `DATA DE INÍCIO` vazia ou inválida
- `DATA FIM` inválida
- `DATA FIM` anterior à `DATA DE INÍCIO`
- rateio inválido ou fora do intervalo 0–100%

Uma linha problemática nunca interrompe a análise das demais.

## Conciliação entre duas competências

Tudo acima responde a uma pergunta de projeção: dada uma competência, que ajuste
ela obriga na seguinte. Com duas competências consecutivas em mãos, a pergunta
muda: **o ajuste devido realmente apareceu?**

O ajuste esperado é exatamente o que o motor de projeção calcula sobre a
competência N. O trabalho novo é achar, na competência N+1, o lançamento que o
corresponde.

### Classificação de cada linha

Toda linha recebe uma classificação explícita, com o motivo em texto:

| Classificação | Significado |
|---|---|
| `CURRENT_COMPETENCE` | Competência corrente |
| `RETROACTIVE_ADD` | Acréscimo retroativo |
| `RETROACTIVE_DISCOUNT` | Desconto retroativo |
| `UNDETERMINED` | Indeterminado — revisão manual |

### Linha de competência × linha retroativa

Uma linha de ajuste retroativo pode legitimamente ter `% RATEIO` **negativo**.
Isso vale só para ela: numa linha de competência, rateio fora de 0–100% continua
sendo erro de dados.

O que separa as duas **não é o sinal**, é o período. Uma linha cujo período
termina antes do primeiro dia da competência do arquivo não pode pertencer àquela
competência — só pode ser retroativa. O sinal entra depois, só para separar
acréscimo de desconto.

Isso importa nos dois sentidos. Um retroativo pode ser **positivo**: uma linha de
28/05 a 31/05 com rateio +1 dentro da fatura de junho é acréscimo retroativo de
maio, não competência corrente. E um rateio negativo cujo período cai dentro da
própria competência do arquivo não é classificado às pressas — fica
**indeterminado** e vai para revisão.

Isso importa porque uma linha perfeitamente normal costuma carregar a `DATA DE
INÍCIO` real da pessoa, que pode ser de meses atrás. Quem foi admitido em 20/05 e
segue ativo aparece na fatura de junho com início 20/05 e **`DATA FIM` vazia** —
é linha de competência, não retroativo. O retroativo dessa mesma admissão é outra
linha, fechada: 20/05 a 31/05.

### De onde vem a competência de cada fatura

Detectar competência por cluster de datas erra justamente na fatura N+1, que
carrega os retroativos do mês anterior. A competência passa a ser buscada em
ordem de força da evidência, e a fonte usada aparece na tela junto do nível de
confiança:

1. **Campo explícito da planilha** (aba `Resumo`, cabeçalho, rótulo
   "Competência") — confiança alta;
2. **Nome do arquivo** (`FATURA_..._Junho_2026`) — confiança alta;
3. **Sequência com a outra fatura** já carregada — confiança média;
4. **Heurística de datas**, calculada apenas sobre as linhas de competência
   corrente — confiança média;
5. **Confirmação manual** — sempre disponível.

A competência detectada nunca é escondida: fonte e confiança ficam à vista, e o
seletor de mês e ano está sempre habilitado.

### Identificação da pessoa

`GROOT ID` é a chave; matrícula é o segundo recurso e só serve de ponte quando
aponta para um único GROOT. **Nome nunca é chave.** Identificadores são
normalizados antes da comparação, porque o Excel devolve o mesmo número ora como
texto, ora como número: `123456`, `123456.0` e `"123456"` são a mesma pessoa.

Uma matrícula ligada a dois GROOTs distintos é conflito: os registros **não são
unidos**, e o apontamento cai em revisão manual.

### Classificações

Para dar um caso por conciliado não basta achar uma linha parecida: são
comparadas cinco dimensões — **identidade, competência de origem, período, sinal
e FTE** — e a tela mostra quais bateram.

| Situação | Status |
|---|---|
| As cinco dimensões conferem | **Conciliado** |
| Nada correspondente na competência seguinte | **Ajuste ausente** |
| Períodos se sobrepõem mas não coincidem | **Período divergente** |
| Período certo, rateio diferente | **FTE divergente** |
| Período certo, sentido oposto | **Sinal incorreto** |
| O mesmo período lançado mais de uma vez | **Possível duplicidade** |
| Retroativo na N+1 sem fato na N que o explique | **Retroativo sem origem** |
| Matrícula ligada a mais de um GROOT | **Identidade ambígua** |
| Linha que não pôde ser classificada | **Indeterminado** |
| Mais de uma leitura possível | **Revisão manual** |

Um mesmo registro pode acumular alertas — período divergente **e** sinal
incorreto, por exemplo. FTE só vira alerta próprio quando o **rateio** difere:
um período menor já produz FTE menor por aritmética, e apontar isso à parte
seria ruído.

Cada apontamento carrega um nível de confiança — **alta**, **média** ou
**revisão necessária** — acompanhado da razão que o justifica.

### Identificador serve para achar, não para calcular

Quem define o ajuste devido são as **datas**. `GROOT ID` e matrícula servem para
**achar a pessoa na fatura seguinte**. Confundir as duas coisas custava caro: o
motor de projeção exige os dois identificadores e, sem eles, devolve erro —
descartando um ajuste que era perfeitamente calculável.

Na conciliação isso passa a ser tratado por partes:

| Linha na fatura de origem | O que acontece |
|---|---|
| Tem GROOT **ou** matrícula, falta o outro | Rastreável: ajuste calculado e conciliado normalmente |
| Não tem nenhum dos dois, mas tem ajuste devido | Status **Sem identificador**: o ajuste é mostrado, e o texto diz que falta a chave de busca — não o dado do cálculo |
| Não tem nenhum dos dois e não tem ajuste devido | Nada é dito — não há o que conferir |
| Datas ilegíveis | **Revisão manual**: sem data não há como saber se havia ajuste |

Nome nunca preenche essa lacuna: homônimo existe, e unir registros por nome
produziria conciliação falsa.

### O que não vira apontamento

Fatura real tem rodapé, subtotal e linha de formatação. Uma linha **sem nome, sem
identificador e sem período** não é um colaborador: não há o que conciliar nem o
que revisar. Essas linhas são descartadas da análise e apenas **contadas**, para
que o total continue batendo com o arquivo.

Pela mesma razão, uma linha da competência seguinte sem período válido não é
classificada como "indeterminada" — ela simplesmente não pode ser um retroativo,
e apontá-la numa fatura de centenas de linhas seria ruído. `Indeterminado` fica
reservado para a linha que **tem** período e cujos sinais se contradizem. E um
erro de preenchimento em linha anônima não vira `Revisão manual`: sem saber de
quem é a linha, não há o que procurar na fatura seguinte.

### Diferença exata

Quando o período não fecha, o app não diz apenas "divergente": aponta o trecho.
Esperado 21/05 a 31/05 e encontrado 22/05 a 31/05 vira *"não conciliado: 21/05"*,
e as ações oferecidas são substituir a linha, complementar **somente 21/05**,
manter o original ou revisar — nenhuma marcada de antemão.

### O limite da evidência

Duas faturas, sozinhas, nem sempre provam qual era a situação operacional real.
Por isso o app não escreve "a fatura está errada". Ele escreve o que de fato
sabe: *ajuste esperado pela regra do snapshot não localizado*, *retroativo
encontrado sem origem identificável nas duas faturas*, *informação insuficiente —
requer validação da movimentação*.

Pela mesma razão, uma data de desligamento deduzida de um retroativo é sempre
apresentada como inferência, com a base à vista: *"último dia faturável inferido:
17/05 — base da inferência: desconto encontrado de 18/05 a 31/05"*. O app não
escreve "esta pessoa foi desligada em 17/05".

Uma terceira fonte (base de movimentações de RH) pode, no futuro, corroborar um
retroativo sem origem. O motor já aceita essa fonte como parâmetro opcional; ela
não é necessária para o resto funcionar.

### Impacto financeiro

Quando a planilha traz uma tarifa inequívoca, o app converte FTE em valor e
mostra original, esperado, encontrado e diferença. Quando não traz, diz
exatamente isso — *"impacto financeiro não calculado: tarifa insuficiente ou
ambígua"* — e nunca inventa um número.

### Quem decide

O app detecta, calcula, compara, explica e sugere. **Ele não decide que uma
fatura deve ser alterada.** Todo apontamento nasce em "manter como está"; a
correção só é calculada quando o usuário pede, e ainda assim pode ser editada
antes de valer. A exportação parte da competência N+1, gera um arquivo novo e
preserva o restante do documento — abas, fórmulas, estilos e formatos. O que
mudou, e o que o usuário decidiu não tratar, fica registrado numa aba
`CONCILIAÇÃO` dentro da cópia, com competências, classificação da linha, status,
alertas, confiança e sua base, cobrança original, ajuste esperado e encontrado,
decisão do usuário, período e FTE finais, alteração aplicada, impacto financeiro
e observação.

### Pessoa-dia: a unidade de identidade do aplicativo

`GROOT ID` representa **uma pessoa**. Logo:

```
pessoa-dia = GROOT normalizado + data normalizada
```

Se a mesma combinação aparece em duas operações, duas abas ou dois arquivos,
não são duas pessoas nem duas diárias — é a mesma pessoa-dia repetida. Contá-la
duas vezes produz dupla cobrança do diarista, dupla contabilização e compensação
indevida de ABS.

A normalização vive num lugar só, `js/identity.js`, e é usada por todas as abas.
Duas implementações da mesma regra é como se cria a situação em que a
Conciliação acha que são a mesma pessoa e a Extração acha que não.

| Entrada | Normalizado |
|---|---|
| `123456`, `"123456"`, `123456.0`, `" 123456 "` | `123456` |
| `ABC123` | `ABC123` — alfanumérico não é convertido |
| `00123456` | `00123456` — zeros à esquerda preservados |

Zeros à esquerda **não** são removidos: `00123456` pode ser um identificador
diferente de `123456`, e não há como saber pelo dado. Unir duas pessoas por
engano é pior do que deixar de unir duas grafias da mesma.

Registro **sem GROOT nunca é deduplicado**. Tratar todos os vazios como a mesma
chave apagaria pessoas diferentes de uma vez — o oposto do objetivo. Eles são
preservados e marcados como *GROOT ausente — revisar*.

Na deduplicação, fica sempre a **primeira ocorrência encontrada**, e o que foi
descartado aparece num painel com GROOT, data, operação mantida e operação
descartada. Nada sai em silêncio.

### Estrutura e identidade de uma linha criada

Quando um ajuste ausente vira linha nova, as duas coisas vêm de lugares
diferentes, e confundi-las já produziu linha com o nome de outra pessoa:

- a **estrutura** — posição das colunas, estilos, formatos, fórmulas — vem
  sempre da fatura N+1, que é o documento sendo gerado;
- a **identidade** — GROOT, matrícula, nome, cargo, regime, escala — vem do
  registro conciliado da pessoa, que quando ela não existe na N+1 só existe na
  fatura N.

O template **nunca** fornece valor de célula: dele se herda estilo, borda,
formato e altura. Todos os valores vêm da identidade lógica da pessoa.

Isso importa por causa de um detalhe do Excel: `spliceRows` **desloca os índices
das linhas seguintes**. Um índice guardado antes de uma remoção passa a apontar
para outra pessoa — e uma linha nova de João sairia com o GROOT, o nome e a
matrícula de Maria, erro nominal de faturamento, silencioso.

A defesa é dupla, de propósito:

1. **Fase 0** — antes de qualquer operação estrutural, cada inclusão captura um
   snapshot imutável do template (só aparência) e a identidade lógica da pessoa.
   Depois disso nenhum índice antigo é consultado.
2. **Identidade sempre reescrita** — ainda que o snapshot viesse da linha errada,
   o que se herda é aparência.

Antes de dar a linha por boa, os campos nominais gravados são conferidos contra
os esperados, com o GROOT comparado normalizado. Divergiu, a inclusão é
**abortada e registrada** — nunca gravada em silêncio.

## Datas

Toda data é normalizada para um inteiro `AAAAMMDD` e comparada numericamente.
Isso elimina de uma vez comparação de texto, ambiguidade entre `DD/MM` e `MM/DD`
e o erro clássico de fuso horário que transforma 15/08 em 14/08.

Entradas aceitas: objeto `Date` do leitor de planilha, número de série do Excel,
`DD/MM/AAAA`, `AAAA-MM-DD`.
