# Ajustes de Fatura

Ferramenta de página única que lê a planilha de **Labor enviado** e responde
perguntas diferentes sobre ela, uma em cada aba.

Todo o processamento acontece no navegador. Nenhum dado sai da máquina: não há
servidor, banco nem envio de arquivos.

## As seis abas

| Aba | Pergunta | Entrega |
|---|---|---|
| **Conciliação Faturas** | O que já foi cobrado está certo? | Ajustes projetados, ou a conferência de duas competências |
| **Fusão de Linhas** | O arquivo bate com o alvo do cliente? | Labor equalizado dia a dia contra o retorno |
| **Extração · Diarista** | Quem foram os diaristas do período? | Uma planilha por operação, no layout de origem |
| **Calcular ABS** | O absenteísmo ficou dentro do range? | Absenteísmo antes e pós diaristas, com o Excel gerencial |
| **Validação Template** | A fatura tem inconsistência? E quanto o cliente deve reconhecer? | Apontamentos explicados; e o retorno previsto dia a dia |
| **Guia** | — | Resumo conceitual do que cada aba faz |

A aba é escolhida pelo topo da página e também pela URL: `index.html#fusao` abre
direto na segunda.

**O app não escreve valor monetário em lugar nenhum** — nem na tela, nem em arquivo
exportado. Tudo é medido em pessoas, dias, FTE e headcount. Onde a conclusão depende
de um lançamento, o que entra na conta é o **sinal** dele: positivo é cobrança,
negativo é estorno, e é essa diferença que separa dupla cobrança de ajuste
retroativo. A grandeza não muda nenhuma conclusão do app, então não é escrita.

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

Compara, dia a dia, o headcount ativo no Labor com a quantidade que o cliente
aponta na aba de Retorno (`Qtd. PREF` e `Q Pós Comp.`). Onde sobra gente monta um
plano e o aplica — retirar linha, adiar início, encurtar fim, ou pausar e retomar
o contrato preservando GROOT e matrícula — e devolve o Labor corrigido no mesmo
layout do original, com abas `A_INCLUIR` e `REVISAR` para o que depende de
decisão humana. Aceita ainda dois arquivos opcionais: a base de diaristas, para
preencher os dias em falta, e o HCM Report, para achar quem está na base do
cliente e não tem cobertura nenhuma no Labor.

```
Soltar planilha  →  conciliação dia a dia  →  plano aplicado  →  Baixar Labor ajustado
```

### Extração · Diarista

Assistente de três passos sobre o controle de diaristas — o arquivo único com uma
aba por operação (Pouso Alegre SVC e XD, Poços de Caldas, Varginha, Divinópolis,
Patos de Minas). Filtra pela **data de solicitação**, opcionalmente por
**solicitante** (interno ou cliente), e resolve os **Groot IDs repetidos** em um
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
3. turno vira horário pelo turno declarado da operação (`ESCALA_TURNO_OPERACAO`),
   levantado nas faturas 3PL. Cada filial tem uma **manhã** e uma **tarde**; o que muda
   de uma para outra é o nome que a coluna ESCALA dá a cada turno:

   | operação | manhã | tarde |
   |---|---|---|
   | Varginha | `AM` → `03:00 07:00 08:00 11:20` | `PM` → `11:00 14:00 15:00 20:00` |
   | Poços de Caldas | `AM` → `03:00 07:00 08:00 12:48` | `PM`, `SD` → *não levantada* |
   | Pouso Alegre SVC | `SVC` → `03:00 07:00 08:00 12:48` | `XD` → `13:00 17:00 18:00 22:45` |
   | Pouso Alegre XD | — | `XD` → `13:00 17:00 18:00 22:45` |

   O turno é declarado primeiro e os apelidos depois, para a equivalência ("em Poços,
   `SD` é a tarde"; "em Pouso, `XD` é a tarde") ficar escrita num lugar só em vez de
   virar dois horários iguais em duas chaves;
4. célula vazia recebe o padrão da operação (`ESCALA_HORARIO_PADRAO`);
5. turno **sem horário levantado** (`SD` e `FULL` no SVC, e a **tarde de Poços**) sai
   como está e vira tópico de revisão — escrever o horário de outro turno no lugar
   seria pôr um dado errado com cara de certo.

O `PM` de Varginha já esteve errado: valia `10:00 15:00 16:00 19:48`, tirado de 30 linhas
da aba DIARISTAS de julho, quando o turno da tarde da unidade é `11:00 14:00 15:00 20:00`
— as 28 pessoas do LABOR de agosto em "1º e 2º Turno". `PM` é 1.618 dos 3.683 registros de
Varginha no SIGO, 44% da filial, saindo com o horário de outro turno.

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
daquele dia** — contados uma vez só, e apenas quando têm Groot ID. Ao clicar em
**Validar**, um diálogo pergunta quais diaristas podem compensar: **somente os
internos** (o padrão — diarista pedido pelo cliente é custo do cliente),
**somente os do cliente**, ou **ambos**. Em **ambos o interno tem prioridade**:
o do cliente só entra no dia depois de esgotados todos os internos, e apenas no
que sobrou do déficit — o total abatido é o mesmo, muda a atribuição, e o
resultado diz quanto veio de cada fonte.
A agência não importa; quem não tem solicitante fica sempre fora. O abate é limitado ao próprio déficit: um dia
nunca fica positivo por sobra de diarista.

A planilha de absenteísmo é aceita nos dois formatos: **um arquivo por período**,
com abas nomeadas por intervalo (`PASVC 16.07-15.08`) cobrindo dois meses, ou os
**dois arquivos mensais** antigos (`PASVC Jul` + `PASVC Ago`), um em cada campo.
No passo 1 as bases **obrigatórias** (absenteísmo e SIGO) vêm destacadas em
âmbar, com a razão de serem obrigatórias; as opcionais ficam apagadas.

O pool é ainda filtrado pela **ESCALA da operação**, não pela aba do SIGO: as abas
se misturam — a aba `Pouso Alegre XD` carrega linhas `SD` e `FULL`, e o cartão
Same Day nem tinha aba própria (caía na do Service). Onde a filial marca a
operação na ESCALA, cada cartão leva só a sua; onde marca outra coisa (Varginha e
Poços usam `AM`/`PM`, Divinópolis usa horário), o cartão leva tudo **menos** o que
pertence às outras operações da filial. Filial que não distingue a operação na
ESCALA recebe um aviso no cartão: os diaristas são os mesmos dos irmãos e podem
estar contados em dois.

A tela de resultado por operação foi **ocultada**: depois de validar sobra só a
confirmação e os avisos que merecem atenção antes de abrir o arquivo — o resultado
completo, por dia e por operação, está na planilha. Ela traz o absenteísmo
**antes** e **pós compensação** contra o **range contratual de 2,5%**, a lista dos
dias críticos, no formato do modelo de
referência: uma aba `Diaristas <operação>` no layout do SIGO, um `Unificado` por
operação e um `Resumo Gerencial` com fórmulas vivas — o `COUNTIFS` aponta para a
aba de diaristas da própria operação e filtra a ESCALA, então a planilha
recalcula sozinha se alguém editar.

As datas gravadas no Excel são **sempre meia-noite UTC** (`dataExcel`). Meia-noite
local vira `03:00Z` em UTC-3, ou seja o serial `46223,125` em vez de `46223` — e o
`COUNTIFS` do resumo, que compara com `DATE(2026,7,20)` (serial inteiro), não casa
com nada: **todo o abate recalculava para zero ao abrir a planilha**, embora os
valores em cache estivessem certos. Um ambiente em UTC não reproduz o defeito, por
isso `tests/abs-datas.test.js` força o fuso de São Paulo. Além do unificado há **um download por
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

### Validação Template

Dois modos, escolhidos no topo da aba: **auditar a fatura** e **simular o retorno**.

#### Auditar inconsistências

Recebe uma fatura no template padrão e a audita **como um auditor faria**: cruza GROOT ID, nome,
cargo, vínculo, datas e o sinal dos lançamentos entre as abas `LABOR` e `DIARISTAS` e separa o que
é erro do que é movimentação legítima de contrato ou ajuste retroativo. `RESUMO` e `HORA EXTRA` entram como apoio.

O ponto da aba não são as regras isoladas — é a **combinação de evidências**. Duas linhas com o
mesmo GROOT ID têm exatamente os mesmos campos preenchidos e podem ser coisas opostas:

```
GROOT 2110045  MARIANA COSTA DO AMARAL SANTOS  Temporário  30/03→10/08
GROOT 2110045  MARIANA COSTA DO AMARAL             Efetivo     11/08→aberto   → efetivação

GROOT 2110046  BRUNO TEIXEIRA LOPES   Efetivo     Líder      23/03→aberto
GROOT 2110046  SERGIO ALMEIDA CUNHA       Temporário  Auxiliar   03/06→04/08   → duas pessoas
```

No primeiro caso um nome é o outro com um sobrenome a mais, os períodos se **encostam** sem
sobrepor e o vínculo evolui de Temporário para Efetivo: o app classifica como **Informativo** e
sugere consolidar numa linha só, apagando a DATA FIM. No segundo os nomes não têm relação, os
cargos diferem e os períodos **se sobrepõem** com as duas linhas positivas: **Crítico**.

O sinal do lançamento muda a conclusão. A mesma sobreposição LABOR × DIARISTA é dupla cobrança quando o
LABOR é positivo e cai para **Revisar** quando é negativo — negativo costuma ser estorno,
desligamento retroativo ou a própria devolução do fixo para pagar os dias como diária. O app não
decide que um lançamento negativo está errado.

O que a aba procura: GROOT compartilhado por pessoas diferentes · transição Temporário → Efetivo ·
variação cadastral de nome · períodos sobrepostos · LABOR × DIARISTA · GROOT ausente (com destaque
para auxiliar e operador, os cargos conciliados pessoa a pessoa) · GROOT fora do padrão
**aprendido da própria planilha** · homônimos · diária duplicada no mesmo dia · `VALOR FINAL` ≠
quantidade × unitário · datas invertidas · data fora do período · campos obrigatórios vazios ·
grafia inconsistente do regime · tarifa destoante da mediana do mesmo tipo · linhas idênticas ·
hora extra de quem não está no LABOR.

Os apontamentos são reunidos em **blocos por categoria** — mesmo GROOT com pessoas diferentes, com
nomes parcialmente coincidentes, com a mesma pessoa, sem GROOT, homônimos, cobrança em duplicidade,
identificador fora do padrão, datas, coerência entre colunas e qualidade de cadastro. Cada bloco traz
o total e a quebra por severidade no cabeçalho; clicar abre os casos um a um. Uma fatura que geraria
trinta cartões em coluna cabe em sete blocos fechados.

Cada achado sai com quatro coisas: a **severidade** (Crítico, Revisar, Cadastro, Informativo), o
**raciocínio em texto** dizendo quais evidências levaram àquela conclusão, uma **sugestão de
correção** e os **registros envolvidos lado a lado**. As opções de decisão (manter, consolidar,
excluir, marcar como ajuste, ignorar) ficam registradas e saem no relatório.

**Nenhuma linha da fatura é alterada.** A aba aponta, explica e sugere; a correção é feita por
você na planilha.

```
Soltar a fatura  →  apontamentos por severidade  →  decidir caso a caso  →  Baixar relatório (.xlsx)
```

#### Simular o retorno

Antecipa o confronto que a **Fusão de Linhas** só consegue fazer depois que o retorno oficial
chega. Reconstrói o **PREF** a partir da aba `LABOR` da fatura, soma o **S&OP diário** de todas as
operações da planilha operacional e prevê o que o cliente tende a reconhecer.

O PREF é reconstruído com a **mesma regra da Fusão de Linhas** — cargo na lista, dia dentro de
`[início, fim]` com fim vazio valendo até o fim do período, e **`% RATEIO` maior que zero**.

O rateio ≤ 0 é **estorno**: devolução do que já foi cobrado, não gente no quadro do dia. O confronto
com o S&OP e com o QF pergunta **quantas pessoas o dia tem**, e para essa pergunta o estorno não é
relevante — a linha continua no arquivo exportado, apenas fora da conta. A Fusão de Linhas aplica a
mesma regra no `elig` do seu Labor, porque as duas reconstroem o mesmo número a partir do mesmo
arquivo: divergir aqui daria dois PREFs para a mesma fatura. Quantas linhas ficaram de fora aparece
como aviso na tela — ficar fora da conta não é ficar escondido.

O S&OP tem **duas fontes**, escolhidas na tela: a **planilha operacional**, dia a dia somando as
operações, ou um **valor fixo do mês** digitado à mão, igual em todos os dias — nesse caso a
planilha não é necessária e o período sai da competência da própria fatura.

A regra da previsão é **assimétrica de propósito**:

```
Q Pós previsto = MIN(PREF, S&OP)
```

O cliente pode **cortar** o que foi apresentado acima do planejamento, mas **não paga** o que nem
sequer foi enviado. PREF 130 contra S&OP 138 não vira 138 — vira 130, e os 8 de folga são um
alerta de subfaturamento para conferir o Labor, não receita a receber. Prever `Q Pós = S&OP` nos
dois sentidos seria otimista no lado errado, e é o erro que o módulo existe para não cometer.

Duas coisas que a reconstrução do PREF **não** faz:

- **Não soma o S&OP por posição de coluna.** Cada coluna de cada aba operacional é resolvida para
  uma data completa antes de qualquer soma, e a soma é por data. Uma aba com uma coluna a mais no
  começo somaria o dia 16 de uma com o dia 17 da outra, em silêncio. Se as datas das abas não
  coincidirem, o app **recusa** em vez de somar.
- **Não chuta cargo para dentro.** Entram só os de `CARGOS_PREF` (`js/config.js`) — auxiliar e
  operador. Liderança e indiretos saem pelo cargo, sem alarme; cargo desconhecido sai **com**
  alarme, nomeado, porque um PREF subestimado em silêncio é pior que um alerta.

Célula de S&OP vazia é **ausência, não zero**: o dia sai como `REVISÃO NECESSÁRIA`, sem número
previsto. Zero silencioso inventaria um risco de corte que não existe.

O que vai ao confronto é o **quadro do dia**: as linhas do `LABOR` mais as **diárias já lançadas** na
aba `DIARISTAS` da fatura. Contar só o fixo aqui, enquanto a equalização conta os dois, dava duas
respostas para o mesmo dia na mesma tela — a fatura já equalizada voltava com 170 contra 182 e
"possível subfaturamento de −12", quando 170 fixos + 12 diárias fecham exatamente em 182. A célula
mostra o total com a composição embaixo, e o diagnóstico repete a conta.

Cada dia sai com S&OP por operação, S&OP total, quadro, Q Pós previsto, gap, correção prevista,
status e diagnóstico em texto. O painel separa **provável redução**, **alinhado** e **possível
subfaturamento**, do maior desvio para o menor.

#### Sugestões de equalização — o motor da Fusão, dentro da Validação

Quando a fonte é o **valor fixo** — o QF que o cliente reconhece — a Validação para de só apontar o
excesso e passa a resolvê-lo: quais linhas do `LABOR` **retirar**, de quem **adiar o início**, de
quem **antecipar o fim**, quem **pausar e retomar**. Cada sugestão abre com o raciocínio, o impacto
em HC e dias, e a vigência atual da linha.

Isto não é uma segunda heurística parecida com a da Fusão de Linhas: é **a mesma função**. A
inteligência saiu de dentro de `js/fusao.js` para `js/equalizacao.js`, e as duas abas chamam a
mesma:

```
                    eqEqualizar()
                         │
              ┌──────────┴──────────┐
        Fusão de Linhas       Validação Template
     (alvo = retorno oficial,  (alvo = QF do cliente,
      dia a dia, por tipo)      constante no período)
```

A única diferença é o vocabulário de entrada — a Fusão fala `Date` e retorno dia a dia; a Validação
fala `AAAAMMDD` e um número só. O motor não conhece data nenhuma: trabalha num eixo de inteiros onde
`d+1` é o dia seguinte, e as duas conversões moram no próprio `equalizacao.js`. Mesmo Labor e mesmo
alvo dão o mesmo plano, e o teste confere isso ação por ação.

O que o motor faz — e o que ele se recusa a fazer — está descrito em `js/equalizacao.js`. O resumo é
que ele olha o **período inteiro**, nunca o dia isolado: só retira quem cabe no excesso em *todos* os
dias em que está ativo, só antecipa o fim de quem já terminaria dentro do trecho em sobra, e trata
queda-e-retomada de demanda como **pausa**, não como desligamento. Excesso sem solução vai para
**Revisar**; dia abaixo do QF vai para **Falta**, porque equalizar não inventa pessoa.

Na Validação o plano é **sugestão**: nada é aplicado à fatura. A matemática fecha a curva; quem sabe
se a movimentação aconteceu de verdade é a operação — e o aviso está na tela.

#### A fatura equalizada, pronta

O botão no fim da seção devolve **a fatura inteira**: as mesmas abas que entraram, com estilos e
fórmulas, e o `LABOR` e o `DIARISTAS` reescritos com o resultado do plano. Exportar duas abas soltas
obrigava a colar de volta à mão, e colar à mão é onde a correção volta a virar erro. As quatro abas
de documentação — `EQUALIZACAO`, `INCLUSOES`, `REVISAR`, `METADADOS` — entram no fim.

A reescrita é **no lugar**, célula a célula: `spliceRows` do ExcelJS não remove nada numa aba cujas
linhas carregam fórmula — sai em silêncio, sem efeito — e o arquivo saía com o Labor **duplicado**.
Escrever por cima ainda tem a vantagem de cada célula guardar o próprio formato. Data vira **serial**,
nunca objeto `Date`, porque o ExcelJS grava `Date` pelo fuso local. O que muda além das duas abas são
só **resultados de fórmula em cache**, que o Excel recalcula ao abrir.

**Quinta ação: retirar diária acima do QF.** As quatro fases do motor mexem no quadro fixo. Quando
elas terminam e ainda sobra excesso, esse excesso não é fixo: é diária lançada por cima de um quadro
já no teto, e cortar mais gente fixa criaria falta em outro dia. Então sai a diária excedente — nunca
mais que o excesso do dia, nunca num dia sem excesso. A ordem é a do abate "ambos" **lida ao
contrário**: lá se gasta primeiro o diarista interno e o do cliente por último; aqui, para devolver
excesso, sai primeiro o interno e o do cliente fica por último, pelo mesmo motivo dos dois lados.

O quadro do dia é **o fixo do `LABOR` mais as diárias que a fatura já lança** na aba `DIARISTAS`, e é
esse total que vai ao confronto com o QF. A diária entra no motor como pessoa *imutável*: conta na
curva — vaga ocupada é vaga ocupada — mas nenhuma das quatro fases pode escolhê-la, porque quem
equaliza mexe no quadro fixo, não em diária que já aconteceu.

Ignorar essa aba foi um defeito real, e caro: num dia em que a fatura já pagava 15 diárias o app pedia
mais 23, e **29 pessoa-dia apareciam nas duas listas** — a mesma pessoa, no mesmo dia, cobrada duas
vezes. A exportação também substituía as diárias existentes em vez de somar a elas.

E o lado da falta **já vem preenchido**, com a base SIGO carregada. Nos dias abaixo do QF o app escolhe
diaristas **solicitados naquele dia e sem cobrança naquele dia** — nem no `LABOR`, nem como diária já
lançada nesta fatura —, **primeiro os internos até acabarem** e só então os do cliente — a mesma prioridade do abate "ambos" da aba Calcular ABS, pelo mesmo motivo:
gastar diarista do cliente com interno sobrando escolhe a fonte errada, e o total não denuncia. Nunca
entra mais gente do que a falta do dia.

**Ninguém entra num dia em que não foi solicitado.** Cada pessoa-dia escolhido é reconferido contra
um índice montado do zero a partir das linhas cruas do SIGO — não contra o mapa que produziu a
escolha, porque conferir com a mesma fonte não confere nada — antes de virar linha do arquivo. O que
não passasse seria descartado e contado, e o total de conferidos e recusados sai no `METADADOS`. A
garantia já valia por construção; o que ela não tinha era quem a cobrasse, e garantia por construção
é a que um refactor apaga sem ninguém ver — bastaria alguém emendar o intervalo entre dois dias
soltos da mesma pessoa para inventar diária que a base não tem.

Essa gente sai na aba **`Diaristas`**, uma linha por pessoa-dia, cargo `Diarista` e quantidade 1 —
como a fatura lança diária. **Não** no `Labor`: quem cobre a falta é diarista, não quadro fixo, e
lançá-lo como fixo seria cobrar outra coisa. A consequência é visível e está dita no `METADADOS`: o
`Labor` exportado fica **abaixo do QF** nos dias cobertos por diária, e é assim mesmo.

A `ESCALA` do diarista é a da **operação da unidade** (`ESCALA_HORARIO_PADRAO`); unidade que casa com
mais de uma operação — Pouso Alegre tem SVC e XD, com horários diferentes — sai em branco, porque
escolher um turno seria chutar qual. Matrícula, regime, dias trabalhados e turno não são deduzíveis do
SIGO e não são escritos. Falta que não houver diarista livre para cobrir sai declarada em `REVISAR`,
com o tamanho: um número que some seria pior que um número feio.

A aba `Diaristas` do arquivo traz **as diárias que a fatura já tinha, como estão, mais as novas** —
é a fatura equalizada inteira, não só o que o app acrescentou.

Junto vão, sempre, as abas que explicam o que houve: `EQUALIZACAO` (cada ação com o motivo, o impacto
e as datas de antes e depois), `INCLUSOES`, `REVISAR` e `METADADOS`. Em `REVISAR`, o excesso que cabe
dentro das diárias do dia sai com a explicação certa — o quadro fixo já está no teto e o que passa é
diária já lançada — em vez de mandar o usuário procurar um contrato para partir.

Com a base SIGO carregada, cada sugestão mostra ainda se a pessoa **também aparece como diarista**
justamente nos dias que o plano tira do fixo. É a explicação operacional da correção — transição de
fixo para diária — e é onde uma dupla cobrança apareceria. A base de diaristas **não muda a
matemática**; só acrescenta contexto.

#### Diaristas disponíveis no dia (opcional)

Soltando também a base SIGO de diaristas, aparece ao lado do gap a coluna **Diaristas disp.**:
quantos havia naquele dia para cobrir a falta, quebrados entre **internos** e **do cliente**. A aba
do SIGO é escolhida pela **unidade da própria fatura**, lida do `RESUMO` — não pelo nome do arquivo.

Disponível quer dizer **ainda não cobrado**. Cada Groot ID solicitado é cruzado com o `LABOR` da
mesma fatura **naquele dia**, e quem já está lá sai da conta — a coluna mostra `N já no LABOR`, para
o desconto ficar visível em vez de implícito. Vale a mesma regra do rateio: o estorno ao lado do
fixo **não** devolve a pessoa ao mercado — se há lançamento positivo ativo no dia, ela segue cobrada,
e contá-la como diarista disponível seria contá-la duas vezes. A
mesma pessoa pedida pelos dois lados no mesmo dia é **uma pessoa, contada como interna** — a mesma
prioridade do abate "ambos" da aba Calcular ABS.

A cobertura possível de cada dia é `MIN(disponíveis, falta)`, então um dia nunca fica positivo por
sobra de diarista, e o diagnóstico do dia passa a dizer quantos havia e quanto dariam para cobrir.
É **leitura, não abate**: nada entra no PREF nem no Q Pós previsto.

A exportação gera quatro abas: `COMPARATIVO` (leitura humana), `RETORNO SIMULADO` (com os mesmos
cabeçalhos que a Fusão de Linhas procura, para servir de retorno de teste lá), `METADADOS` (o
aviso de que o arquivo é simulado e as fórmulas usadas) e `FORA DO PREF` (toda linha excluída e
por quê).

**É previsão, nunca retorno confirmado** — o rótulo aparece na tela e no arquivo.

```
Fatura + planilha S&OP  →  PREF reconstruído × S&OP  →  desvios  →  Exportar Retorno Simulado
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
index.html            as seis abas e a ordem de carga dos scripts
css/styles.css        estilos das seis abas
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
js/validacao.js       motor de auditoria da fatura — puro, sem DOM
js/validacao-ui.js    leitura do .xlsx e tela da Validação Template (IIFE)
js/simulacao.js       previsão do retorno a partir de PREF × S&OP — puro, sem DOM
js/equalizacao.js     motor de equalização — compartilhado pela Fusão e pela Validação
js/simulacao-ui.js    leitura dos dois arquivos e tela do Retorno Simulado (IIFE)
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
  SheetJS não grava formatação, e a exportação precisa preservar o visual do modelo.

## Testes

```bash
npm test          # roda todos os arquivos de teste
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

`tests/abs-prioridade.test.js` cobre a prioridade do solicitante interno no
abate "ambos" da aba Calcular ABS: o do cliente nunca entra com diarista interno
sobrando no dia, o interno é sempre consumido até o teto, e o total abatido
continua sendo
`min(pool inteiro, déficit)`. O defeito que ele pega é invisível no total — só
aparece na repartição entre as duas fontes.

`tests/validacao.test.js` cobre o julgamento da Validação Template — a parte que decide se duas
linhas do mesmo GROOT ID são uma efetivação ou duas pessoas. Reproduz casos de uma fatura real e
prova os dois sentidos do erro: que uma efetivação normal não vira Crítico, e que pessoas
diferentes com períodos sobrepostos e ambas as linhas positivas viram. Cobre também a queda de
severidade quando há estorno em jogo, o aprendizado do padrão de GROOT a partir da própria
planilha, e que uma fatura limpa não gera achado nenhum.

`tests/simulacao.test.js` cobre a previsão do retorno. Prova a assimetria (`MIN(PREF, S&OP)`) nos
dois sentidos com os números reais de Varginha — 16/07 = 120 + 16, 19/07 = 80 + 12, 10/08 = 120 +
17 — e prova que retroativo negativo não reduz o headcount do dia, que liderança e indiretos não
entram sozinhos, que cargo desconhecido vira aviso em vez de palpite, e que célula de S&OP vazia
vira revisão em vez de zero. Um bloco garante a **paridade com a Fusão de Linhas** — o PREF de cada
dia é confrontado com a conta da Fusão, dia a dia, e o teste morde: o dia do estorno muda de valor
se o rateio ≤ 0 voltar a entrar com o sinal. Outro cobre os **diaristas disponíveis**: quem já está
cobrado no LABOR do dia não conta, o estorno ao lado do fixo não libera ninguém, o mesmo Groot pedido
pelos dois lados é uma pessoa contada como interna, e sem a planilha o resultado não inventa
disponibilidade.

`tests/equalizacao.test.js` cobre o motor de equalização, que até então rodava só no navegador e
nunca teve teste. Metade dos casos são as **recusas** — não retira quem cobre um dia sem excesso, não
antecipa o fim de quem trabalha depois do trecho, não pausa quem não volta, não trata alvo 0 como
demanda zero — porque é delas que sai um plano que fecha a curva sem furá-la para baixo. A outra
metade é a **paridade**: o mesmo Labor montado nos dois vocabulários, os dois caminhos completos
rodados de ponta a ponta, e plano idêntico exigido ação por ação, com um caso de controle provando
que alvo diferente dá plano diferente — senão "iguais" poderia ser só duas listas vazias.

`tests/reconciliation.test.js` cobre a conciliação entre duas faturas: as
classificações de status, a reconstrução da cobrança original (o corte projeta
da admissão, não cria mês cheio), a classificação de linha normal × retroativa
nos dois sentidos de sinal, a competência por ordem de evidência, a
normalização de identificadores, a checagem de sequência, os alertas múltiplos,
a diferença exata em dias, a identidade da linha criada e a garantia de que
nenhum apontamento nasce aceito.

## Detecção da competência

Em ordem de força da evidência:

1. Se a aba de Retorno tiver `Data Trab.` com um mês concentrando **≥70%**
   das datas válidas, esse mês é a competência.
2. Caso contrário, o app localiza o mês-âncora — o mais recente com volume real
   de movimentações — e avalia apenas ele e os dois anteriores. Desligamentos
   antigos ficam fora da conta, e uma data solta (erro de digitação, data
   futura) não forma um mês significativo.
3. Se dois meses ficarem próximos, aparece **Alterar** ao lado da competência
   para escolher manualmente. A linha da competência é sempre clicável.

## Exportação

O arquivo gerado — nomeado pela competência, com uma única aba de ajustes — mantém
as dez colunas originais da planilha, copiadas da linha importada, e acrescenta
à direita apenas o bloco do ajuste: tipo, início, fim, dias, FTE, competência de
origem, competência de aplicação e motivo.

A formatação não é imitada: o exportador lê o estilo real do arquivo importado
(fonte, preenchimento, bordas, alinhamento, formatos de data e percentual) e o
reaplica célula a célula. Como o estilo vem da própria linha de origem, cores
associadas a categorias — regime de contrato, cargo, escala — são preservadas
automaticamente. Só entram no Excel os ajustes que permanecem marcados.

## Planilha de entrada

O app procura a aba de Labor do arquivo e localiza as colunas **pelo nome do
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
