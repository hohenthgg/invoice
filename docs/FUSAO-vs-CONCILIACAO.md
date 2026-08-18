# Fusão de Linhas × Conciliação Faturas

As duas abas olham para o **mesmo Labor** e mesmo assim respondem perguntas
diferentes. Confundi-las é o erro mais caro do app: uma trata do **quadro de
pessoas** que o MELI vai aceitar, a outra trata do **dinheiro** que já foi (ou
não foi) cobrado. Este documento explica a Fusão de Linhas por dentro e coloca
as duas lado a lado.

---

## Em uma frase cada

| Aba | Pergunta que responde | Eixo |
|---|---|---|
| **Fusão de Linhas** | Meu Labor bate com o quadro que o MELI está usando? E, se não bate, como eu conserto o arquivo? | **Headcount por dia** |
| **Conciliação Faturas** | O ajuste retroativo que eu deveria ter recebido apareceu mesmo na fatura seguinte? | **Dinheiro por competência** |

A Fusão é **pré-envio** — arruma o arquivo antes de ele virar cobrança.
A Conciliação é **pós-fatura** — audita o que já foi cobrado.

---

## Fusão de Linhas, por dentro

### 1. O que ela lê

Um único `.xlsx` que precisa conter duas coisas, em abas quaisquer (o app acha
pelo cabeçalho, não pela posição):

- **Labor** — cabeçalho com `GROOT ID` e `DATA DE INÍCIO`. Também usa `NOME`,
  `MATRICULA`, `CARGO`, `DATA FIM`, `% RATEIO`, `REGIME`.
- **Retorno MELI** — cabeçalho com `Data Trab.` e `Qtd. PREF`. Também usa
  `Q Pós Comp.`, `Desvio`, `Employee type`, `Ocorrência`.

Linhas de absenteísmo (`GROOT` = `ABS`, ou nome contendo "absenteísmo") são
descartadas: não são pessoas do quadro.

### 2. A contagem diária

Para cada dia do Retorno, a aba conta quantas pessoas do seu Labor estavam
ativas — `início ≤ dia ≤ fim` (fim vazio = em aberto), somando o `% RATEIO`,
não a linha:

- só os **cargos que o MELI considera no PREF** (`CARGOS_MELI`, editável no
  topo de `js/fusao.js`);
- **full-time e part-time contam em trilhos separados** — cada dia do Retorno
  vem marcado por tipo e é comparado apenas com o seu.

O **alvo do MELI** é `Q Pós Comp.` quando a coluna existe; senão, `Qtd. PREF`.

```
Dif. = (headcount do seu Labor no dia)  −  (alvo do MELI no dia)
```

### 3. De quem é a culpa

Cada dia recebe um dos três diagnósticos:

| Resultado | Status | Leitura |
|---|---|---|
| `Dif. ≠ 0` | **divergência sua** | linha faltando, `DATA FIM` cortada cedo demais, ou gente a mais no arquivo |
| `Dif. = 0` e `Desvio > 3` | **lado MELI** | o Scheduling/Rostering não reconheceu a escala do dia — cobrança é com o BC deles |
| `Dif. = 0` e `Desvio < 0` | **lado MELI** | operou mais gente do que o declarado — pedir nova versão do PREF |
| resto | **OK** | contagem e desvio sob controle |

Essa separação é o ponto da aba: **nem toda diferença é sua**. O app diz quando
o arquivo está certo e o erro é do outro lado, para você não "consertar" um
Labor que já estava correto.

### 4. O plano de equalização

Onde **sobra** gente, o app monta um plano em **quatro fases, nesta ordem**,
sempre escolhendo primeiro quem entrou mais recentemente. A regra que atravessa
todas: **nenhuma fase pode empurrar um dia para o negativo** — cortar excesso de
segunda não pode criar falta na terça.

| Fase | Ação | Quando se aplica |
|---|---|---|
| 1 | **Retirar** | o período ativo inteiro da pessoa cabe dentro do excesso — a linha sai do arquivo |
| 2 | **Adiar início** | o excesso está no começo do contrato — a `DATA DE INÍCIO` anda para frente *(desativável na tela)* |
| 3 | **Pausar / Retomar** | o alvo **cai e volta a subir** (um "vale") — em vez de demitir, fecha o contrato no início do vale e reabre no fim: mesma pessoa, mesmo GROOT, duas linhas *(exige a data "pausar a partir de")* |
| 4 | **Encurtar fim** | sobra residual no fim do contrato, sem retomada — a `DATA FIM` é antecipada, e **só** para quem já terminaria ali de qualquer forma |

Dias com **alvo 0** (escala não publicada, Go-Live) ficam de fora de tudo: não
geram corte nem inclusão — zero não é demanda, é ausência de informação.

Onde **falta** gente, não há o que cortar: o dia vai para a lista **A incluir**,
com a quantidade por dia. Aí você sobe a **base de diaristas** e o app monta um
`Labor_Diaristas_….xlsx` já preenchido, sem repetir o mesmo GROOT no mesmo dia —
e com uma aba `DIAS SEM COBERTURA` quando a base não dá conta.

O que sobra sem solução limpa cai em **Revisar** (vale isolado que só zeraria
partindo um contrato em dois — o app não faz isso sozinho).

### 5. As verificações do arquivo

Independentes da conciliação, seis checagens que pegam erro de preenchimento:

- **GROOT duplicado** — mesma pessoa com períodos sobrepostos (conta duas vezes no dia);
- **GROOT zerado/faltante** — risco de deduplicação no sistema do MELI;
- **Datas invertidas** — `DATA FIM` antes do início;
- **Buracos de 1 dia** — ativo na véspera e no dia seguinte, sem linha no meio;
- **Cortes de período** — `DATA FIM` antes do fim do período, sem linha sucessora;
- **Desvio divergente** — o desvio informado pelo MELI ≠ o recalculado (versões diferentes do PREF).

### 6. O que sai

`Labor_ajustado_AAAAMMDD.xlsx`, preservando o cabeçalho original:

- aba **Labor enviado ao MELI** — as linhas já com o plano aplicado (pausas viram
  dois segmentos da mesma pessoa);
- aba **A_INCLUIR** — os dias e quantidades que faltam;
- aba **REVISAR** — os vales sem solução automática.

---

## Conciliação Faturas, por dentro (o contraste)

A aba 1 nunca conta pessoa por dia. Ela trabalha sobre uma regra de
faturamento: **o snapshot do dia 15 congela a cobrança da competência**.

Ela tem dois modos:

**Projetar ajustes** (uma fatura) — pergunta o que *deveria* aparecer no mês
seguinte. Para cada pessoa:

- **ativa no dia 15** e desligada antes do fim do mês → **DESCONTAR** os dias
  entre o desligamento e o fim da competência;
- **admitida depois do dia 15** → **ACRESCENTAR** o período trabalhado, que
  nenhum snapshot pegou;
- ativa no dia 15 e até o fim do mês → nada a ajustar.

O resultado é medido em **FTE** (`dias ÷ dias da competência × rateio`), não em
cabeças. Junto vêm as travas contra cobrança dupla: dedupe exato, detecção de
períodos sobrepostos da mesma pessoa e separação de linhas com cadastro inválido.

**Conciliar duas faturas** (N e N+1) — pergunta se o ajuste devido *apareceu*.
Reconstrói, pessoa a pessoa, o que foi efetivamente faturado em cada
competência, separa o que é da competência corrente do que é retroativo, e
classifica: `Conciliado`, `Ajuste ausente`, `Período divergente`, `FTE
divergente`, `Sinal incorreto`, `Possível duplicidade`, `Retroativo sem
origem`, `Identidade ambígua`… A identidade é por **GROOT ID**, com matrícula
como ponte segura; **nome nunca une registros**.

---

## Lado a lado

| | **Fusão de Linhas** | **Conciliação Faturas** |
|---|---|---|
| **Entrada** | 1 arquivo: Labor + Retorno MELI | 1 fatura (projetar) ou 2 faturas consecutivas (conciliar) |
| **Unidade** | pessoas por dia (com rateio) | FTE por competência |
| **Referência externa** | `Qtd. PREF` / `Q Pós Comp.` do MELI | a regra do snapshot do dia 15 |
| **Granularidade** | dia | competência (mês) |
| **Pergunta** | o quadro bate? | o dinheiro bate? |
| **Momento** | antes de enviar o Labor | depois de a fatura chegar |
| **Altera arquivo?** | **sim** — gera o Labor ajustado pronto para importar | **não** — aponta, explica e sugere; a decisão é sua |
| **Culpa compartilhada** | separa "erro seu" de "erro do MELI" (desvio) | separa "ajuste ausente" de "retroativo sem origem" |
| **Saída** | `Labor_ajustado_….xlsx` (+ `A_INCLUIR`, `REVISAR`) e `Labor_Diaristas_….xlsx` | relatório de conciliação por pessoa, exportável |

### A diferença que mais importa

A **Fusão escreve**. Ela devolve um arquivo modificado — linhas removidas,
datas mexidas, contratos pausados — porque o objetivo é entregar ao MELI um
Labor que o sistema deles aceite sem desvio.

A **Conciliação não escreve nada**. Ela é uma auditoria: mostra a divergência,
diz por quê, e para por aí. Mudar uma fatura é decisão humana, e o app foi feito
para não tomá-la.

### Como elas se encadeiam

```
Fusão de Linhas  →  Labor ajustado  →  vira fatura no MELI  →  Conciliação Faturas
   (o quadro)                                                      (o dinheiro)
```

Um Labor que passou pela Fusão chega à Conciliação com menos ruído: sem GROOT
duplicado, sem `DATA FIM` cortada por engano, sem buraco de um dia. Quase todo
apontamento de "ajuste ausente" que sobra na aba 1 é ajuste de verdade, não
sujeira de preenchimento.

---

## Onde isso mora no código

| Assunto | Arquivo |
|---|---|
| Fusão inteira (leitura, conciliação, plano, exportações, HCM) | `js/fusao.js` |
| Regra do dia 15 / detecção de ajustes | `js/engine.js` |
| Reconstrução do que foi faturado | `js/billing.js` |
| Confronto entre duas faturas | `js/reconciliation.js` |
| Detecção da competência | `js/competence.js` |
| Identidade (GROOT / matrícula) — uma só no app inteiro | `js/identity.js` |
