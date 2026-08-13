/* Testes do motor de regras e da detecção de competência.
   Sem dependências externas: basta `node tests/engine.test.js`. */
"use strict";
const { load } = require("./load");
const ctx = load(["config.js", "dates.js", "engine.js", "competence.js"]);
const { ymd, fmtShort, buildCompetence, detectAdjustment, analyze, detectCompetence } = ctx;

let pass = 0, fail = 0;
function check(ok, label, extra) {
  if (ok) { pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}

const AGO = buildCompetence(2026, 8);   // 31 dias, corte 15/08

function pessoa(inicio, fim, rateio, groot, matricula) {
  return { srcRow: 2, groot: groot || "1", nome: "Teste", matricula: matricula || "m1",
           inicio, fim, rateio: rateio === undefined ? 1 : rateio, raw: {} };
}
function caso(label, emp, tipoEsperado, diasEsperados, fteEsperado) {
  const r = detectAdjustment(emp, AGO);
  const tipo = r === null ? "NENHUM" : r.kind;
  const dias = (r && r.days) || 0;
  const fte  = (r && r.fte) || 0;
  const ok = tipo === tipoEsperado && dias === diasEsperados
          && (fteEsperado === undefined || Math.abs(fte - fteEsperado) < 1e-9);
  check(ok, label, `${tipo} ${dias}d ${r && r.start ? fmtShort(r.start) + "→" + fmtShort(r.end) : ""}`);
}

console.log("\nRegras de ajuste (competência Agosto/2026, corte 15/08)");
caso("A · entrou 01/08, saiu 20/08 → descontar 21–31/08 (11 dias)",
     pessoa(ymd(2026,8,1), ymd(2026,8,20)), "DESCONTAR", 11, -11/31);
caso("B · entrou 20/08, sem saída → acrescentar 20–31/08 (12 dias)",
     pessoa(ymd(2026,8,20), null), "ACRESCENTAR", 12, 12/31);
caso("C · entrou 20/08, saiu 25/08 → acrescentar 20–25/08 (6 dias)",
     pessoa(ymd(2026,8,20), ymd(2026,8,25)), "ACRESCENTAR", 6, 6/31);
caso("D · admissão antiga, sem saída → sem ajuste",
     pessoa(ymd(2025,3,10), null), "NENHUM", 0);
caso("E · saiu 14/08 (antes do corte) → sem ajuste",
     pessoa(ymd(2026,2,1), ymd(2026,8,14)), "NENHUM", 0);
caso("F · saiu 15/08 (DATA FIM inclusiva) → descontar 16–31/08 (16 dias)",
     pessoa(ymd(2026,1,5), ymd(2026,8,15)), "DESCONTAR", 16, -16/31);
caso("H · rateio 50%, saiu 20/08 → FTE −(11/31 × 0,50)",
     pessoa(ymd(2026,8,1), ymd(2026,8,20), 0.5), "DESCONTAR", 11, -(11/31*0.5));

const g = detectAdjustment(pessoa(ymd(2026,8,28), ymd(2026,8,25)), AGO);
check(g.kind === "ERRO", "G · início 28/08 depois do fim 25/08 → erro de dados", g.errors && g.errors.join("; "));

console.log("\nLimites da competência");
caso("saiu no último dia do mês → sem ajuste",
     pessoa(ymd(2026,1,1), ymd(2026,8,31)), "NENHUM", 0);
caso("saída no mês seguinte não gera desconto nesta competência",
     pessoa(ymd(2026,1,1), ymd(2026,9,5)), "NENHUM", 0);
caso("entrou 16/08 (primeiro dia pós-corte) → 16 dias",
     pessoa(ymd(2026,8,16), null), "ACRESCENTAR", 16, 16/31);
caso("entrou e saiu em 31/08 → 1 dia",
     pessoa(ymd(2026,8,31), ymd(2026,8,31)), "ACRESCENTAR", 1, 1/31);

console.log("\nDuplicidade, ordenação e erros");
{
  const base = [
    pessoa(ymd(2026,8,1),  ymd(2026,8,20), 1, "A", "10"),   // desconto 21–31
    pessoa(ymd(2026,8,1),  ymd(2026,8,20), 1, "A", "10"),   // linha idêntica → não duplica
    pessoa(ymd(2026,8,20), null,           1, "B", "11"),   // acréscimo 20–31
    pessoa(ymd(2026,8,18), ymd(2026,8,25), 1, "B", "11"),   // período sobreposto → revisão
    pessoa(ymd(2026,8,22), ymd(2026,8,23), 1, "C", "12"),   // acréscimo 22–23
    pessoa(ymd(2026,8,1),  ymd(2026,8,18), 1, "D", "13"),   // desconto 19–31
    { srcRow: 9, groot: "", nome: "Sem dados", matricula: "", inicio: null, fim: null, rateio: 1, raw: {} }
  ];
  const { adjustments, errors } = analyze(base, AGO);
  check(adjustments.length === 4, "linha idêntica não gera ajuste duplicado", adjustments.length + " ajustes");
  check(errors.length === 2, "sobreposição e linha inválida vão para erros", errors.length + " erros");
  check(adjustments[0].kind === "DESCONTAR" && adjustments[3].kind === "ACRESCENTAR",
        "descontos aparecem antes dos acréscimos");
  check(adjustments[0].movDate <= adjustments[1].movDate,
        "dentro do grupo, ordenado pela data da movimentação");
}

console.log("\nDetecção automática da competência");
{
  const antigo = pessoa(ymd(2020,1,1), ymd(2020,1,20));
  let d = detectCompetence([antigo], { top: { y: 2026, m: 8 }, total: 100, topShare: 0.85 });
  check(d.comp.label === "Agosto/2026" && d.via === "retorno",
        "maioria ≥70% em Data Trab. decide sozinha", d.comp.label + " via " + d.via);

  const recentes = [pessoa(ymd(2026,8,1), ymd(2026,8,20), 1, "a", "1"),
                    pessoa(ymd(2026,8,2), ymd(2026,8,22), 1, "b", "2"),
                    pessoa(ymd(2026,8,3), ymd(2026,8,25), 1, "c", "3")];
  d = detectCompetence(recentes, { top: { y: 2026, m: 7 }, total: 30, topShare: 0.5 });
  check(d.via === "cluster", "Data Trab. dividida meio a meio cai para o cluster recente", d.comp.label + " via " + d.via);

  const historico = [];
  for (let i = 0; i < 80; i++) historico.push(pessoa(ymd(2025,1,1), ymd(2025,2,17 + (i % 10)), 1, "H" + i, "h" + i));
  for (let i = 0; i < 6;  i++) historico.push(pessoa(ymd(2026,8,1), ymd(2026,8,18 + (i % 8)), 1, "N" + i, "n" + i));
  d = detectCompetence(historico, null);
  check(d.comp.label === "Agosto/2026",
        "80 desligamentos antigos não elegem uma competência passada", d.comp.label);

  const comDataFutura = historico.slice(80).concat([pessoa(ymd(2026,8,1), ymd(2030,5,20), 1, "X", "x")]);
  d = detectCompetence(comDataFutura, null);
  check(d.comp.label === "Agosto/2026", "uma data futura isolada não sequestra a detecção", d.comp.label);
}

console.log(`\n${pass} passaram, ${fail} falharam\n`);
process.exit(fail ? 1 : 0);
