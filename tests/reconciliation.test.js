/* Testes do motor de conciliação entre duas faturas subsequentes.
   Sem dependências externas: basta `node tests/reconciliation.test.js`. */
"use strict";
const { load } = require("./load");
const ctx = load(["config.js", "dates.js", "engine.js", "competence.js", "billing.js",
                  "competence-source.js", "reconciliation.js"],
                 ["RECON_STATUS", "RECON_CONF", "RECON_META", "LINE_CLASS", "COMP_SOURCE"]);
const { ymd, fmtShort, fmtYmd, buildCompetence, reconcile, sugerirCorrecao,
        RECON_STATUS, RECON_CONF, LINE_CLASS, COMP_SOURCE, normId, splitInvoiceLines,
        checkSequence, classifyLine, originalBilling, parseCompetenceText, fromFileName } = ctx;

let pass = 0, fail = 0;
function check(ok, label, extra) {
  if (ok) { pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}

const MAIO  = buildCompetence(2026, 5);   // 31 dias, corte 15/05
const JUNHO = buildCompetence(2026, 6);   // 30 dias, corte 15/06

/** linha de Labor genérica */
function linha(o) {
  return { srcRow: o.srcRow || 2, groot: o.groot || "1001", nome: o.nome || "Pessoa Teste",
           matricula: o.matricula || "m1", inicio: o.inicio, fim: o.fim === undefined ? null : o.fim,
           rateio: o.rateio === undefined ? 1 : o.rateio, raw: {} };
}
/** roda a conciliação e devolve o item da pessoa indicada */
function conciliar(linhasMaio, linhasJunho) {
  return reconcile({ origem: { employees: linhasMaio, comp: MAIO },
                     seguinte: { employees: linhasJunho, comp: JUNHO } });
}
function itemDe(res, nome) { return res.items.find(i => i.nome === nome); }
function resumo(it) {
  if (!it) return "(sem item)";
  const e = it.esperado ? `${it.esperado.kind} ${fmtShort(it.esperado.start)}→${fmtShort(it.esperado.end)}` : "—";
  const a = it.achado ? `${it.achado.kind} ${fmtShort(it.achado.start)}→${fmtShort(it.achado.end)}` : "—";
  return `${it.status} | esp ${e} | ach ${a}`;
}

/* ================================================================ */
console.log("\nNormalização de identificadores");
check(normId("123456.0") === normId(123456), "123456.0 e 123456 são o mesmo GROOT",
      normId("123456.0") + " vs " + normId(123456));
check(normId(" 77201 ") === "77201", "espaços em volta são ignorados");
check(normId("") === "" && normId(null) === "", "vazio e nulo não viram identificador");

/* ================================================================ */
console.log("\nLinha normal × linha retroativa");
{
  // linha normal de junho de quem foi admitido em maio: DATA FIM vazia
  const normal = linha({ inicio: ymd(2026,5,20), fim: null });
  // linha retroativa de maio dentro da fatura de junho
  const retro  = linha({ inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1 });
  const s = splitInvoiceLines([normal, retro], JUNHO);
  check(s.normais.length === 1 && s.normais[0] === normal,
        "admissão de maio com DATA FIM vazia é linha normal de junho");
  check(s.retroativos.length === 1 && s.retroativos[0] === retro,
        "período fechado dentro de maio é retroativo");
  // rateio negativo não pode cair em 'erro de dados'
  const res = conciliar([], [retro]);
  const it = res.items[0];
  check(it && it.status === RECON_STATUS.SEM_ORIGEM,
        "rateio negativo não é tratado como erro de dados", resumo(it));
}

/* ================================================================ */
console.log("\nSequência de competências");
check(checkSequence(MAIO, JUNHO).ok, "Maio → Junho é sequência válida");
check(checkSequence(JUNHO, MAIO).kind === "invertida", "Junho → Maio acusa inversão");
check(checkSequence(MAIO, buildCompetence(2026,8)).kind === "distante",
      "Maio → Agosto acusa competências não subsequentes");
check(!checkSequence(MAIO, buildCompetence(2026,5)).ok, "mesma competência duas vezes não é válida");

/* ================================================================ */
console.log("\nCasos de conciliação");

// Caso 1 — desconto correto
{
  const maio  = [linha({ nome: "Caso1", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Caso1", inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Caso1");
  check(it && it.status === RECON_STATUS.CONCILIADO && it.confianca === RECON_CONF.ALTA,
        "1 · ativo em 15/05, saiu 20/05, junho traz −21→31/05 → CONCILIADO", resumo(it));
}

// Caso 2 — desconto ausente
{
  const maio  = [linha({ nome: "Caso2", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const it = itemDe(conciliar(maio, []), "Caso2");
  check(it && it.status === RECON_STATUS.AUSENTE && it.esperado.kind === "DESCONTAR"
        && it.esperado.days === 11,
        "2 · desconto esperado e nada em junho → AJUSTE AUSENTE (11 dias)", resumo(it));
}

// Caso 3 — desconto parcial
{
  const maio  = [linha({ nome: "Caso3", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Caso3", inicio: ymd(2026,5,22), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Caso3");
  check(it && it.status === RECON_STATUS.PERIODO && it.diffDias === -1,
        "3 · esperado 21→31, encontrado 22→31 → PERÍODO DIVERGENTE, diferença de 1 dia",
        resumo(it) + " diff=" + (it && it.diffDias));
}

// Caso 4 — acréscimo correto
{
  const maio  = [linha({ nome: "Caso4", inicio: ymd(2026,5,28), fim: null })];
  const junho = [linha({ nome: "Caso4", inicio: ymd(2026,5,28), fim: ymd(2026,5,31), rateio: 1 })];
  const it = itemDe(conciliar(maio, junho), "Caso4");
  check(it && it.status === RECON_STATUS.CONCILIADO && it.esperado.kind === "ACRESCENTAR",
        "4 · admissão 28/05, junho traz +28→31/05 → CONCILIADO", resumo(it));
}

// Caso 5 — entrada e saída entre cortes
{
  const maio  = [linha({ nome: "Caso5", inicio: ymd(2026,5,16), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Caso5", inicio: ymd(2026,5,16), fim: ymd(2026,5,20), rateio: 1 })];
  const it = itemDe(conciliar(maio, junho), "Caso5");
  check(it && it.status === RECON_STATUS.CONCILIADO && it.esperado.days === 5,
        "5 · entrou 16/05 e saiu 20/05, junho traz +16→20/05 → CONCILIADO (5 dias)", resumo(it));
}

// Caso 6 — ajuste duplicado
{
  const maio  = [linha({ nome: "Caso6", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Caso6", srcRow: 2, inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1 }),
                 linha({ nome: "Caso6", srcRow: 3, inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Caso6");
  check(it && it.status === RECON_STATUS.DUPLICADO && it.achados.length === 2,
        "6 · duas linhas iguais em junho → DUPLICADO (nunca removido sozinho)", resumo(it));
}

// Caso 7 — sinal errado
{
  const maio  = [linha({ nome: "Caso7", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Caso7", inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: 1 })];
  const it = itemDe(conciliar(maio, junho), "Caso7");
  check(it && it.status === RECON_STATUS.SINAL,
        "7 · esperado desconto, encontrado acréscimo → SINAL INCORRETO", resumo(it));
}

// Caso 8 — retroativo sem origem
{
  const maio  = [linha({ nome: "Caso8", inicio: ymd(2025,1,10), fim: null })];   // sem movimentação
  const junho = [linha({ nome: "Caso8", inicio: ymd(2026,5,22), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Caso8");
  check(it && it.status === RECON_STATUS.SEM_ORIGEM && it.confianca === RECON_CONF.REVISAO,
        "8 · retroativo em junho sem fato em maio → SEM ORIGEM / revisão", resumo(it));
  check(it && /sem origem identificada/i.test(it.diagnostico)
        && /Requer validação/i.test(it.diagnostico)
        && !/\berro\b/i.test(it.diagnostico) && !/est[áa] errad/i.test(it.diagnostico),
        "8 · o texto pede validação e não chama o lançamento de erro",
        it && it.diagnostico.slice(0, 90));
}

// Caso 9 — rateio divergente
{
  const maio  = [linha({ nome: "Caso9", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Caso9", inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -0.5 })];
  const it = itemDe(conciliar(maio, junho), "Caso9");
  const esperadoFte = -11/31, achadoFte = -11/31*0.5;
  check(it && it.status === RECON_STATUS.FTE
        && Math.abs(it.esperado.fte - esperadoFte) < 1e-9
        && Math.abs(it.achado.fte - achadoFte) < 1e-9,
        "9 · mesmo período com rateio 50% → FTE DIVERGENTE", resumo(it));
}

/* ================================================================ */
console.log("\nDecisão do usuário");
{
  const maio  = [linha({ nome: "Caso10", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const res = conciliar(maio, []);
  const it = itemDe(res, "Caso10");
  // Caso 10 — nada vem pré-aceito
  check(it && it.decisao === "MANTER" && it.sugestao === null,
        "10 · apontamento nasce em 'manter como está', sem sugestão aplicada");
  check(res.items.every(i => i.decisao === "MANTER"),
        "10 · nenhum item é aceito automaticamente");

  // Caso 11 — ao aceitar, a sugestão é calculada mas continua sendo só proposta
  const sug = sugerirCorrecao(it, null);
  check(sug && sug.acao === "INCLUIR" && sug.start === ymd(2026,5,21)
        && sug.end === ymd(2026,5,31) && Math.abs(sug.fte + 11/31) < 1e-9,
        "11 · aceitar sugestão propõe INCLUIR −21→31/05 com FTE −0,3548",
        sug && JSON.stringify({a:sug.acao, d:sug.days, f:sug.fte}));
  check(it.decisao === "MANTER",
        "11 · calcular a sugestão não altera a decisão por conta própria");
}

/* ================================================================ */
console.log("\nSugestões por tipo de divergência");
{
  const maio  = [linha({ nome: "Sub", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Sub", srcRow: 7, inicio: ymd(2026,5,22), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Sub");
  const sub = sugerirCorrecao(it, "SUBSTITUIR");
  check(sub && sub.acao === "SUBSTITUIR" && sub.alvoRow === 7 && sub.days === 11,
        "período divergente · substituir aponta a linha existente e o período completo");
  const comp = sugerirCorrecao(it, "COMPLEMENTAR");
  check(comp && comp.acao === "INCLUIR" && comp.start === ymd(2026,5,21)
        && comp.end === ymd(2026,5,21) && comp.days === 1,
        "período divergente · complementar cobre só o dia que faltou",
        comp && `${fmtShort(comp.start)}→${fmtShort(comp.end)} ${comp.days}d`);
}
{
  const maio  = [linha({ nome: "Dup", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Dup", srcRow: 4, inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1 }),
                 linha({ nome: "Dup", srcRow: 9, inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Dup");
  const sug = sugerirCorrecao(it, null);
  check(sug && sug.acao === "REMOVER" && sug.manterRow === 4
        && sug.alvoRows.length === 1 && sug.alvoRows[0] === 9,
        "duplicado · sugere remover a repetição e manter a primeira linha");
}

/* ================================================================ */
console.log("\nLinha-modelo para inclusões (identidade da linha criada)");
{
  // pessoa SEM nenhuma linha na fatura seguinte: não pode haver modelo, senão a
  // linha incluída herdaria o cadastro de outra pessoa do arquivo de destino
  const maio  = [linha({ nome: "Somente-Maio", groot: "8001", matricula: "n1",
                         inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Outra Pessoa", groot: "9009", matricula: "n9", srcRow: 2,
                         inicio: ymd(2026,6,1), fim: null })];
  const it = itemDe(conciliar(maio, junho), "Somente-Maio");
  check(it && it.status === RECON_STATUS.AUSENTE && it.modeloRow === null,
        "pessoa ausente da fatura seguinte não recebe linha-modelo",
        it && "modeloRow=" + it.modeloRow);
}
{
  // pessoa COM linha normal na fatura seguinte: o modelo é a linha dela
  const maio  = [linha({ nome: "Com-Linha", groot: "8002", matricula: "n2",
                         inicio: ymd(2026,5,28), fim: null })];
  const junho = [linha({ nome: "Ruido", groot: "9009", matricula: "n9", srcRow: 2,
                         inicio: ymd(2026,6,1), fim: null }),
                 linha({ nome: "Com-Linha", groot: "8002", matricula: "n2", srcRow: 5,
                         inicio: ymd(2026,5,28), fim: null })];
  const it = itemDe(conciliar(maio, junho), "Com-Linha");
  check(it && it.modeloRow === 5,
        "pessoa presente na fatura seguinte aponta a própria linha como modelo",
        it && "modeloRow=" + it.modeloRow);
}

/* ================================================================ */
console.log("\nIdentificação ambígua");
{
  // mesma matrícula apontando para dois GROOTs distintos
  const maio = [linha({ nome: "Amb", groot: "5001", matricula: "dup", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Outro", groot: "5002", matricula: "dup", inicio: ymd(2026,6,1), fim: null })];
  const it = itemDe(conciliar(maio, junho), "Amb");
  check(it && it.confianca === RECON_CONF.REVISAO,
        "matrícula ligada a dois GROOTs → confiança rebaixada para revisão", resumo(it));
}

/* ================================================================ */
console.log("\nContexto e resumo");
{
  const maio  = [linha({ nome: "C1", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) }),
                 linha({ nome: "C2", groot: "2002", matricula: "m2", inicio: ymd(2026,5,28), fim: null })];
  const junho = [linha({ nome: "C1", inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1 })];
  const res = conciliar(maio, junho);
  check(res.summary.total === 2 && res.summary.conciliados === 1 && res.summary.pendencias === 1,
        "resumo conta conciliados e pendências", JSON.stringify(res.summary));
  check(res.summary.acrescimosAusentes === 1, "acréscimo ausente é contado à parte");
  check(res.contexto.sequencia.ok && res.contexto.retroativosNaSeguinte === 1,
        "contexto traz a sequência e o total de retroativos lidos");
  check(res.items[0].status !== RECON_STATUS.CONCILIADO,
        "pendências aparecem antes dos conciliados");
}

/* ================================================================
   RECONSTRUÇÃO DA COBRANÇA ORIGINAL
   O snapshot congela a projeção conhecida; ele NÃO cobra mês cheio
   de quem entrou no meio do mês.
   ================================================================ */
console.log("\nCobrança original — o corte não vira mês cheio");
function cobranca(o){ return originalBilling(linha(o), MAIO); }
{
  // Caso A — admissão 01/05, ativa no corte
  const a = cobranca({ inicio: ymd(2026,5,1), fim: null });
  check(a.cobrado && a.start === ymd(2026,5,1) && a.end === ymd(2026,5,31) && a.days === 31,
        "A · admissão 01/05 e ativa em 15/05 → cobrado 01/05 a 31/05 (31 dias)",
        `${fmtYmd(a.start)}→${fmtYmd(a.end)} ${a.days}d`);

  // Caso B — admissão 07/05, ativa no corte: 25 dias, NÃO 31
  const b = cobranca({ inicio: ymd(2026,5,7), fim: null });
  check(b.cobrado && b.start === ymd(2026,5,7) && b.end === ymd(2026,5,31) && b.days === 25,
        "B · admissão 07/05 e ativa em 15/05 → cobrado 07/05 a 31/05 (25 dias), não 31/31",
        `${fmtYmd(b.start)}→${fmtYmd(b.end)} ${b.days}d`);
  check(Math.abs(b.fte - 25/31) < 1e-9, "B · FTE da cobrança original é 25/31",
        b.fte.toFixed(4));

  // Caso C — saída conhecida antes do corte
  const c = cobranca({ inicio: ymd(2026,5,7), fim: ymd(2026,5,12) });
  check(c.cobrado && c.start === ymd(2026,5,7) && c.end === ymd(2026,5,12) && c.days === 6,
        "C · admissão 07/05 com saída 12/05 conhecida no corte → cobrado 07/05 a 12/05",
        `${fmtYmd(c.start)}→${fmtYmd(c.end)} ${c.days}d`);

  // Caso D — desligamento reconhecido depois do corte: cobrança segue integral
  const d = cobranca({ inicio: ymd(2026,5,1), fim: ymd(2026,5,20) });
  check(d.cobrado && d.end === ymd(2026,5,31) && d.days === 31,
        "D · desligamento 20/05 reconhecido após o corte → maio ainda cobrou até 31/05",
        `${fmtYmd(d.start)}→${fmtYmd(d.end)} ${d.days}d`);

  // Caso E — admissão pós-corte não entra no snapshot
  const e = cobranca({ inicio: ymd(2026,5,16), fim: null });
  check(!e.cobrado && e.days === 0,
        "E · admissão 16/05 não entra na fotografia do corte → nada cobrado em maio",
        e.base);

  // Caso F — entrada e saída entre cortes
  const f = cobranca({ inicio: ymd(2026,5,16), fim: ymd(2026,5,20) });
  check(!f.cobrado && f.days === 0,
        "F · entrada 16/05 e saída 20/05 não aparecem em nenhum snapshot", f.base);
}
{
  // O caso Michael, ponta a ponta
  const maio  = [linha({ nome: "Michael", inicio: ymd(2026,5,7), fim: ymd(2026,5,17) })];
  const junho = [linha({ nome: "Michael", inicio: ymd(2026,5,18), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Michael");
  check(it && it.cobrancaOriginal.days === 25 && it.esperado.days === 14
        && it.status === RECON_STATUS.CONCILIADO,
        "Michael · cobrado 25 dias em maio, desconto de 14 dias em junho → CONCILIADO",
        it && `orig ${it.cobrancaOriginal.days}d, esperado ${it.esperado.days}d, ${it.status}`);
  check(it && /17\/05/.test(it.movimentacaoInferida.texto)
        && /inferido/i.test(it.movimentacaoInferida.texto),
        "Michael · último dia faturável é apresentado como INFERIDO, não como fato",
        it && it.movimentacaoInferida.texto);
}

/* ================================================================
   CLASSIFICAÇÃO DE LINHA
   ================================================================ */
console.log("\nClassificação de linha na fatura de junho");
function classe(o){ return classifyLine(linha(o), JUNHO); }
{
  const a = classe({ inicio: ymd(2026,5,28), fim: ymd(2026,5,31), rateio: 1 });
  check(a.classe === LINE_CLASS.RETRO_ADD,
        "retroativo POSITIVO de maio é acréscimo retroativo, não competência corrente", a.classe);
  check(/pertence a Maio\/2026/.test(a.motivo), "…e explica por quê", a.motivo.slice(0,80));

  const b = classe({ inicio: ymd(2026,5,18), fim: ymd(2026,5,31), rateio: -1 });
  check(b.classe === LINE_CLASS.RETRO_DISC, "retroativo negativo é desconto retroativo", b.classe);

  const c = classe({ inicio: ymd(2026,5,7), fim: null, rateio: 1 });
  check(c.classe === LINE_CLASS.CURRENT,
        "linha com admissão em maio e DATA FIM vazia é competência corrente de junho", c.classe);
  check(/DATA FIM está vazia/.test(c.motivo), "…e explica que a admissão anterior não a torna retroativa");

  const d = classe({ inicio: ymd(2026,6,1), fim: ymd(2026,6,30), rateio: 1 });
  check(d.classe === LINE_CLASS.CURRENT, "período dentro de junho é competência corrente", d.classe);

  const e = classe({ inicio: ymd(2026,6,5), fim: ymd(2026,6,20), rateio: -1 });
  check(e.classe === LINE_CLASS.UNDETERMINED,
        "negativo com período dentro da própria competência fica INDETERMINADO", e.classe);
  check(/[Rr]evisão manual/.test(e.motivo), "…e pede revisão manual");
}
{
  // retroativos não podem distorcer a competência detectada pelas linhas correntes
  const junho = [
    linha({ nome:"corrente1", groot:"1", matricula:"a", inicio: ymd(2026,6,1), fim: null }),
    linha({ nome:"corrente2", groot:"2", matricula:"b", inicio: ymd(2026,6,3), fim: null }),
    linha({ nome:"corrente3", groot:"3", matricula:"c", inicio: ymd(2026,6,8), fim: null }),
    linha({ nome:"retroPos", groot:"4", matricula:"d", inicio: ymd(2026,5,28), fim: ymd(2026,5,31), rateio: 1 }),
    linha({ nome:"retroNeg", groot:"5", matricula:"e", inicio: ymd(2026,5,18), fim: ymd(2026,5,31), rateio: -1 })
  ];
  const s = splitInvoiceLines(junho, JUNHO);
  check(s.normais.length === 3 && s.retroativos.length === 2,
        "retroativos positivo e negativo saem das linhas de competência corrente",
        `normais=${s.normais.length} retro=${s.retroativos.length}`);
}

/* ================================================================
   COMPETÊNCIA — ORIGEM DA EVIDÊNCIA
   ================================================================ */
console.log("\nCompetência: prioridade da evidência");
check(parseCompetenceText("Junho/2026").m === 6, "texto 'Junho/2026' é lido");
check(parseCompetenceText("FATURA_CONCILIADA_SMG3_Junho_2026").m === 6, "nome de arquivo com mês por extenso");
check(parseCompetenceText("2026-06").m === 6, "formato AAAA-MM");
check(parseCompetenceText("06/2026").m === 6, "formato MM/AAAA");
check(parseCompetenceText("Junho") === null, "só o nome do mês, sem ano, não basta");
{
  const f = fromFileName("FATURA SMG3 Junho 2026.xlsx");
  check(f && f.comp.y === 2026 && f.comp.m === 6, "competência extraída do nome do arquivo");
  check(COMP_SOURCE.ARQUIVO.confianca === "ALTA" && COMP_SOURCE.HEURISTICA.confianca === "MEDIA",
        "nome do arquivo tem confiança maior que heurística de datas");
}

/* ================================================================
   ALERTAS MÚLTIPLOS E DIFERENÇA EXATA
   ================================================================ */
console.log("\nAlertas múltiplos e diferença exata");
{
  // período divergente E FTE divergente ao mesmo tempo
  const maio  = [linha({ nome: "Multi", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Multi", inicio: ymd(2026,5,22), fim: ymd(2026,5,31), rateio: -0.5 })];
  const it = itemDe(conciliar(maio, junho), "Multi");
  check(it && it.alerts.includes(RECON_STATUS.PERIODO) && it.alerts.includes(RECON_STATUS.FTE),
        "um mesmo registro pode acumular PERÍODO DIVERGENTE e FTE DIVERGENTE",
        it && it.alerts.join(" + "));
}
{
  // período menor com o MESMO rateio não deve gerar alerta de FTE: a diferença
  // de FTE é consequência aritmética do período, não um segundo problema
  const maio  = [linha({ nome: "SoPeriodo", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "SoPeriodo", inicio: ymd(2026,5,22), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "SoPeriodo");
  check(it && it.alerts.includes(RECON_STATUS.PERIODO) && !it.alerts.includes(RECON_STATUS.FTE),
        "período divergente com rateio igual NÃO acumula alerta de FTE",
        it && it.alerts.join(" + "));
}
{
  const maio  = [linha({ nome: "Falta1", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Falta1", inicio: ymd(2026,5,22), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Falta1");
  check(it && it.faltantes.length === 1 && it.faltantes[0].start === ymd(2026,5,21)
        && it.faltantes[0].days === 1,
        "o trecho não conciliado é apontado com data exata (21/05)",
        it && it.faltantes.map(f=>fmtYmd(f.start)+"→"+fmtYmd(f.end)).join(", "));
  check(it && /21\/05/.test(it.diagnostico), "…e aparece no diagnóstico em texto");
}
{
  const maio  = [linha({ nome: "Dims", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Dims", inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1 })];
  const it = itemDe(conciliar(maio, junho), "Dims");
  check(it && it.dimensoes.identidade && it.dimensoes.competencia && it.dimensoes.periodo
        && it.dimensoes.sinal && it.dimensoes.fte,
        "CONCILIADO informa quais dimensões bateram", it && JSON.stringify(it.dimensoes));
  check(it && it.confiancaMotivo.length > 0, "…e de onde veio a confiança", it && it.confiancaMotivo);
}

/* ================================================================
   IDENTIDADE DA LINHA CRIADA
   ================================================================ */
console.log("\nIdentidade da linha criada (estrutura N+1, identidade N)");
{
  const joao = linha({ nome: "JOÃO SILVA", groot: "123", matricula: "j1",
                       inicio: ymd(2026,5,1), fim: ymd(2026,5,20) });
  joao.raw = { groot: 123, nome: "JOÃO SILVA", matricula: "j1", cargo: "Auxiliar", regime: "CLT", escala: "5x2" };
  const outro = linha({ nome: "OUTRA PESSOA", groot: "999", matricula: "z9", srcRow: 2,
                        inicio: ymd(2026,6,1), fim: null });
  const it = itemDe(conciliar([joao], [outro]), "JOÃO SILVA");
  check(it && it.status === RECON_STATUS.AUSENTE, "João precisa de retroativo em junho");
  check(it && it.modeloRow === null,
        "sem linha de João em junho, não há template de estrutura da própria pessoa",
        it && "modeloRow=" + it.modeloRow);
  check(it && it.identidadeCampos && it.identidadeCampos.groot === "123"
        && it.identidadeCampos.nome === "JOÃO SILVA",
        "a identidade da linha nova vem do registro de João na fatura N",
        it && JSON.stringify(it.identidadeCampos));
  check(it && it.identidadeRaw && it.identidadeRaw.cargo === "Auxiliar",
        "…inclusive cargo, regime e escala do cadastro correto");
  check(it && it.groot !== "999" && it.nome !== "OUTRA PESSOA",
        "nenhuma identidade de outro colaborador é reutilizada");
}
{
  const p = linha({ nome: "ComLinha", groot: "555", matricula: "p5",
                    inicio: ymd(2026,5,28), fim: null });
  const junho = [linha({ nome: "Ruido", groot: "999", matricula: "z9", srcRow: 2,
                         inicio: ymd(2026,6,1), fim: null }),
                 linha({ nome: "ComLinha", groot: "555", matricula: "p5", srcRow: 7,
                         inicio: ymd(2026,5,28), fim: null })];
  const it = itemDe(conciliar([p], junho), "ComLinha");
  check(it && it.modeloRow === 7,
        "existindo linha da pessoa na N+1, ela é o template de estrutura",
        it && "modeloRow=" + it.modeloRow);
}

/* ================================================================
   IMPACTO FINANCEIRO
   ================================================================ */
console.log("\nImpacto financeiro");
{
  const maio  = [linha({ nome: "SemTarifa", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const it = itemDe(conciliar(maio, []), "SemTarifa");
  check(it && it.impacto.calculado === false && /não calculado/i.test(it.impacto.motivo),
        "sem tarifa confiável, o impacto não é calculado nem inventado",
        it && it.impacto.motivo.slice(0,60));
}

/* ================================================================
   RUÍDO DE PLANILHA
   Fatura real tem rodapé, subtotal e linhas de formatação. Elas não
   podem virar apontamento — nem "indeterminado", nem "revisão".
   ================================================================ */
console.log("\nLinhas que não são colaborador");
{
  const sobra = { srcRow: 500, groot: null, nome: "", matricula: null,
                  inicio: null, fim: null, rateio: null, raw: {} };
  const sobra2 = { srcRow: 501, groot: "", nome: "   ", matricula: "",
                   inicio: NaN, fim: NaN, rateio: null, raw: {} };
  const real  = linha({ nome: "Gente", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) });

  const s = splitInvoiceLines([sobra, sobra2, real], JUNHO);
  check(s.ignoradas.length === 2 && s.indeterminados.length === 0,
        "rodapé sem nome, sem identificador e sem datas é ignorado, não indeterminado",
        `ignoradas=${s.ignoradas.length} indeterminados=${s.indeterminados.length}`);

  const res = conciliar([real, sobra, sobra2], [sobra, sobra2]);
  check(!res.items.some(i => i.nome === "(sem nome)"),
        "nenhum apontamento '(sem nome)' chega à tela",
        res.items.map(i => i.nome).join(", "));
  check(res.items.every(i => i.status !== RECON_STATUS.INDETERMINADO),
        "sobra de planilha não gera INDETERMINADO");
  check(res.contexto.ignoradas === 4,
        "as linhas ignoradas são contadas e informadas", "ignoradas=" + res.contexto.ignoradas);
}
{
  // uma linha COM período ambíguo continua sendo apontada — esse sinal é real
  const ambigua = linha({ nome: "Ambigua", inicio: ymd(2026,6,5), fim: ymd(2026,6,20), rateio: -1 });
  const res = conciliar([], [ambigua]);
  const it = itemDe(res, "Ambigua");
  check(it && it.status === RECON_STATUS.INDETERMINADO,
        "linha com período contraditório continua sendo apontada como indeterminada",
        it && it.status);
  check(it && /aparece aqui porque/.test(it.diagnostico) && /não entra em nenhuma comparação/.test(it.diagnostico),
        "…e o texto explica por que ela está ali e o que isso implica",
        it && it.diagnostico.slice(0, 100));
}
{
  // erro de dados em linha anônima não é acionável; em linha identificada, é
  const anonima = { srcRow: 9, groot: null, nome: "", matricula: null,
                    inicio: ymd(2026,5,1), fim: ymd(2026,5,20), rateio: 1, raw: {} };
  const res = conciliar([anonima], []);
  check(!res.items.some(i => i.status === RECON_STATUS.REVISAO),
        "erro de dados em linha sem nome não vira apontamento de revisão",
        res.items.map(i => i.status).join(", "));

  const comNome = { srcRow: 9, groot: null, nome: "Fulano Sem Groot", matricula: null,
                    inicio: ymd(2026,5,1), fim: ymd(2026,5,20), rateio: 1, raw: {} };
  const res2 = conciliar([comNome], []);
  const it = itemDe(res2, "Fulano Sem Groot");
  check(it && it.status === RECON_STATUS.SEM_ID,
        "pessoa identificável só pelo nome vira 'Sem identificador', com o ajuste calculado",
        it && it.status);
  check(it && /aparece aqui porque/.test(it.diagnostico) && it.esperado,
        "…com explicação do porquê e o ajuste devido à vista",
        it && it.diagnostico.slice(0, 100));
}

/* ================================================================
   IDENTIFICADOR FALTANDO ≠ AJUSTE INDETERMINÁVEL
   Quem define o ajuste são as datas. O identificador serve para achar
   a pessoa na fatura seguinte.
   ================================================================ */
console.log("\nLinha sem identificador");
{
  // nome + datas, sem GROOT e sem matrícula: o ajuste É calculável
  const p = { srcRow: 120, groot: null, nome: "ALDERINO DE JESUS", matricula: null,
              inicio: ymd(2026,5,1), fim: ymd(2026,5,20), rateio: 1, raw: {} };
  const it = itemDe(conciliar([p], []), "ALDERINO DE JESUS");
  check(it && it.status === RECON_STATUS.SEM_ID,
        "sem GROOT nem matrícula → status próprio 'Sem identificador', não 'Revisão manual'",
        it && it.status);
  check(it && it.esperado && it.esperado.kind === "DESCONTAR" && it.esperado.days === 11
        && Math.abs(it.esperado.fte + 11/31) < 1e-9,
        "o ajuste devido é calculado e mostrado (11 dias, FTE −0,3548), não escondido",
        it && it.esperado && `${it.esperado.kind} ${it.esperado.days}d ${it.esperado.fte.toFixed(4)}`);
  check(it && /não é usado como chave/.test(it.diagnostico)
        && /procurá-lo/.test(it.diagnostico),
        "o texto explica que falta a CHAVE de busca, não o dado do cálculo",
        it && it.diagnostico.slice(0, 110));
  check(it && it.cobrancaOriginal && it.cobrancaOriginal.days === 31,
        "a cobrança original também é reconstruída para essa pessoa");
}
{
  // sem identificador E sem ajuste devido → nada a dizer
  const p = { srcRow: 121, groot: null, nome: "Sem Movimento", matricula: null,
              inicio: ymd(2024,1,1), fim: null, rateio: 1, raw: {} };
  const res = conciliar([p], []);
  check(res.items.length === 0,
        "sem identificador e sem ajuste devido não gera apontamento nenhum",
        res.items.map(i => i.status + ":" + i.nome).join(", "));
}
{
  // GROOT presente, matrícula vazia: é rastreável, entra no confronto normal
  const p = { srcRow: 122, groot: "4242", nome: "Tem Groot", matricula: null,
              inicio: ymd(2026,5,1), fim: ymd(2026,5,20), rateio: 1, raw: {} };
  const junho = [{ srcRow: 5, groot: "4242", nome: "Tem Groot", matricula: null,
                   inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -1, raw: {} }];
  const it = itemDe(conciliar([p], junho), "Tem Groot");
  check(it && it.status === RECON_STATUS.CONCILIADO,
        "GROOT presente e matrícula vazia continua rastreável e concilia normalmente",
        it && it.status);
}
{
  // erro que REALMENTE impede o cálculo continua em revisão, com texto correto
  const p = { srcRow: 123, groot: "5555", nome: "Data Ruim", matricula: "m5",
              inicio: null, fim: ymd(2026,5,20), rateio: 1, raw: {} };
  const it = itemDe(conciliar([p], []), "Data Ruim");
  check(it && it.status === RECON_STATUS.REVISAO,
        "DATA DE INÍCIO vazia continua sendo revisão manual", it && it.status);
  check(it && /datas/.test(it.diagnostico) && !/identificador|GROOT|matrícula/.test(it.diagnostico),
        "…e o texto atribui a causa às datas, nunca ao identificador",
        it && it.diagnostico.slice(0, 110));
}

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
