/* Testes do motor de conciliação entre duas faturas subsequentes.
   Sem dependências externas: basta `node tests/reconciliation.test.js`. */
"use strict";
const { load } = require("./load");
const ctx = load(["config.js", "dates.js", "engine.js", "competence.js", "reconciliation.js"],
                 ["RECON_STATUS", "RECON_CONF", "RECON_META"]);
const { ymd, fmtShort, buildCompetence, reconcile, sugerirCorrecao,
        RECON_STATUS, RECON_CONF, normId, splitInvoiceLines, checkSequence } = ctx;

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
  check(it && it.status === RECON_STATUS.PARCIAL && it.diffDias === -1,
        "3 · esperado 21→31, encontrado 22→31 → AJUSTE PARCIAL, diferença de 1 dia",
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
  check(it && /Informação insuficiente/.test(it.diagnostico),
        "8 · o texto não afirma que a fatura está errada");
}

// Caso 9 — rateio divergente
{
  const maio  = [linha({ nome: "Caso9", inicio: ymd(2026,5,1), fim: ymd(2026,5,20) })];
  const junho = [linha({ nome: "Caso9", inicio: ymd(2026,5,21), fim: ymd(2026,5,31), rateio: -0.5 })];
  const it = itemDe(conciliar(maio, junho), "Caso9");
  const esperadoFte = -11/31, achadoFte = -11/31*0.5;
  check(it && it.status === RECON_STATUS.RATEIO
        && Math.abs(it.esperado.fte - esperadoFte) < 1e-9
        && Math.abs(it.achado.fte - achadoFte) < 1e-9,
        "9 · mesmo período com rateio 50% → RATEIO/FTE DIVERGENTE", resumo(it));
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
        "parcial · substituir aponta a linha existente e o período completo");
  const comp = sugerirCorrecao(it, "COMPLEMENTAR");
  check(comp && comp.acao === "INCLUIR" && comp.start === ymd(2026,5,21)
        && comp.end === ymd(2026,5,21) && comp.days === 1,
        "parcial · complementar cobre só o dia que faltou",
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

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
