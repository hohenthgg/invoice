# Regras de negócio

Este documento descreve o que o motor faz e por quê. Ele é a referência para
mudanças em `js/engine.js` — se uma alteração de código contradisser algo aqui,
uma das duas coisas está errada.

## O conceito

O faturamento MELI trabalha com três datas diferentes, e confundi-las é a
origem de quase todo erro de conferência:

- **Período de folha:** dia 16 do mês anterior até o dia 15 do mês corrente.
- **Competência de faturamento:** dia 01 até o último dia do mês.
- **Data de corte (snapshot):** dia 15.

O quadro existente no dia 15 é congelado e serve de base para a cobrança da
competência inteira. A pergunta que o motor responde **não** é "quantos dias
essa pessoa trabalhou no mês", e sim:

> O que foi congelado no retrato do dia 15, e qual fato posterior ao
> congelamento precisa ser compensado na próxima fatura?

## Ativo no snapshot

```
DATA DE INÍCIO ≤ dia 15
E (DATA FIM vazia OU DATA FIM ≥ dia 15)
```

`DATA FIM` é **inclusiva**: é o último dia em que a pessoa permanece ativa. Quem
tem data fim 20/08 esteve ativo em 20/08, e a ausência começa em 21/08.

Estando ativa no dia 15, a pessoa é cobrada por 100% da competência × % rateio.
Não se calcula pró-rata de admissão anterior ao corte — o snapshot é uma regra
comercial, não um cálculo de folha.

## Os dois ajustes

### Desconto — desligamento pós-corte

A pessoa estava no retrato e foi cobrada integralmente, mas saiu antes do fim do
mês. Os dias posteriores à saída foram cobrados indevidamente.

```
início do ajuste = DATA FIM + 1 dia
fim do ajuste    = último dia da competência
```

### Acréscimo — admissão pós-corte

A pessoa não estava no retrato, logo não foi cobrada, mas trabalhou parte da
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

- Admissão anterior ao dia 15 e sem desligamento: cobrança integral correta.
- Desligamento antes do dia 15: a pessoa não entrou no retrato, então não há o
  que compensar nesta competência.
- Desligamento no último dia do mês ou depois: nada foi cobrado a mais.

Nenhum desses aparece na tela. A interface mostra apenas ação necessária.

## Proteções contra cobrança dupla

- Duas linhas que produziriam o mesmo ajuste (mesma pessoa, mesmo tipo, mesmo
  período) geram **um** ajuste, com aviso.
- Duas linhas da mesma pessoa com períodos **sobrepostos** do mesmo tipo não são
  resolvidas automaticamente: vão para a lista de registros não analisados, para
  decisão humana.
- Períodos distintos da mesma pessoa são mantidos, com aviso de GROOT ou
  matrícula repetidos.

## Erros de dados

Não se misturam com ajustes. São listados à parte e não entram na exportação:

- GROOT ID ou matrícula vazios
- `DATA DE INÍCIO` vazia ou inválida
- `DATA FIM` inválida
- `DATA FIM` anterior à `DATA DE INÍCIO`
- rateio inválido ou fora do intervalo 0–100%

Uma linha problemática nunca interrompe a análise das demais.

## Datas

Toda data é normalizada para um inteiro `AAAAMMDD` e comparada numericamente.
Isso elimina de uma vez comparação de texto, ambiguidade entre `DD/MM` e `MM/DD`
e o erro clássico de fuso horário que transforma 15/08 em 14/08.

Entradas aceitas: objeto `Date` do SheetJS, número de série do Excel,
`DD/MM/AAAA`, `AAAA-MM-DD`.
