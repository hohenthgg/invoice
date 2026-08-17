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

Estando ativa no dia 15, a pessoa entra por 100% da competência × % rateio. Não
se calcula pró-rata de admissão anterior ao corte — o snapshot é uma regra de
apuração, não um cálculo de folha.

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

### Linha de competência × linha retroativa

Uma linha de ajuste retroativo pode legitimamente ter `% RATEIO` **negativo**.
Isso vale só para ela: numa linha de competência, rateio fora de 0–100% continua
sendo erro de dados.

O que separa as duas **não é o sinal**, é o período. Uma linha cujo período
termina antes do primeiro dia da competência do arquivo não pode pertencer àquela
competência — só pode ser retroativa. O sinal entra como evidência adicional,
nunca sozinho.

Isso importa porque uma linha perfeitamente normal costuma carregar a `DATA DE
INÍCIO` real da pessoa, que pode ser de meses atrás. Quem foi admitido em 20/05 e
segue ativo aparece na fatura de junho com início 20/05 e **`DATA FIM` vazia** —
é linha de competência, não retroativo. O retroativo dessa mesma admissão é outra
linha, fechada: 20/05 a 31/05.

### Identificação da pessoa

`GROOT ID` é a chave; matrícula é o segundo recurso e só serve de ponte quando
aponta para um único GROOT. **Nome nunca é chave.** Identificadores são
normalizados antes da comparação, porque o Excel devolve o mesmo número ora como
texto, ora como número: `123456`, `123456.0` e `"123456"` são a mesma pessoa.

Uma matrícula ligada a dois GROOTs distintos é conflito: os registros **não são
unidos**, e o apontamento cai em revisão manual.

### Classificações

| Situação | Status |
|---|---|
| Mesmo período, mesmo sentido, mesmo FTE | **Conciliado** |
| Nada correspondente na competência seguinte | **Ajuste ausente** |
| Períodos se sobrepõem mas não coincidem | **Ajuste parcial** |
| Período certo, quantidade diferente | **Rateio/FTE divergente** |
| Período certo, sentido oposto | **Sinal incorreto** |
| O mesmo período lançado mais de uma vez | **Possível duplicidade** |
| Retroativo na N+1 sem fato na N que o explique | **Retroativo sem origem** |
| Identificação ambígua ou mais de uma leitura possível | **Revisão manual** |

Cada apontamento carrega ainda um nível de confiança: **alta** (mesma pessoa,
mesmo período, mesmo FTE), **média** (correspondência parcial) e **revisão
necessária** (identificação ambígua ou vários candidatos).

### O limite da evidência

Duas faturas, sozinhas, nem sempre provam qual era a situação operacional real.
Por isso o app não escreve "a fatura está errada". Ele escreve o que de fato
sabe: *ajuste esperado pela regra do snapshot não localizado*, *retroativo
encontrado sem origem identificável nas duas faturas*, *informação insuficiente —
requer validação da movimentação*.

Uma terceira fonte (base de movimentações de RH) pode, no futuro, corroborar um
retroativo sem origem. O motor já aceita essa fonte como parâmetro opcional; ela
não é necessária para o resto funcionar.

### Quem decide

O app detecta, calcula, compara, explica e sugere. **Ele não decide que uma
fatura deve ser alterada.** Todo apontamento nasce em "manter como está"; a
correção só é calculada quando o usuário pede, e ainda assim pode ser editada
antes de valer. A exportação parte da competência N+1, gera um arquivo novo e
preserva o restante do documento — abas, fórmulas, estilos e formatos. O que
mudou, e o que o usuário decidiu não tratar, fica registrado numa aba
`CONCILIAÇÃO` dentro da cópia.

## Datas

Toda data é normalizada para um inteiro `AAAAMMDD` e comparada numericamente.
Isso elimina de uma vez comparação de texto, ambiguidade entre `DD/MM` e `MM/DD`
e o erro clássico de fuso horário que transforma 15/08 em 14/08.

Entradas aceitas: objeto `Date` do leitor de planilha, número de série do Excel,
`DD/MM/AAAA`, `AAAA-MM-DD`.
