/* Testes da deduplicação da Extração · Diarista.

   O que estes testes provam:

       MESMO GROOT NORMALIZADO + MESMA DATA = UMA ÚNICA PESSOA-DIA

   independentemente de operação, aba ou arquivo de origem. Um `seen` por
   operação faria `SVC|123456|10/08` e `XD|123456|10/08` sobreviverem os dois,
   produzindo dupla contagem de diarista e compensação indevida de ABS. */
"use strict";
const { load } = require("./load");
const ctx = load(["identity.js", "extraction-dedup.js"], ["DEDUP_MODOS"]);
const { normalizeGroot, hasGroot, normalizeDateKey, personDayKey,
        deduplicarPessoaDia, DEDUP_MODOS } = ctx;

let pass = 0, fail = 0;
function check(ok, label, extra) {
  if (ok) { pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}

/** registro de diarista, como a leitura do SIGO produz */
const reg = (id, date, nome) => ({ id, date, nome: nome || "Diarista", solic: "id",
                                   empresa: "Empresa X", cargo: "Auxiliar", escala: "6x1",
                                   solicitante: "ID Logistics" });
/** roda a dedup sobre operações nomeadas */
function dedup(porOp, modo) {
  return deduplicarPessoaDia(Object.entries(porOp).map(([op, rows]) => ({ op, rows })),
                             modo || DEDUP_MODOS.DIA);
}
const totalMantido = r => Object.values(r.porOperacao).reduce((a, o) => a + o.rows.length, 0);

/* ================================================================ */
console.log("\nNormalização de GROOT");
check(normalizeGroot(123456) === "123456", "número vira texto canônico");
check(normalizeGroot("123456.0") === "123456", "float do Excel perde o .0");
check(normalizeGroot("123456.00") === "123456", "…e o .00 também");
check(normalizeGroot(" 123456 ") === "123456", "espaços em volta somem");
check(normalizeGroot("123456") === normalizeGroot(123456)
      && normalizeGroot(123456) === normalizeGroot("123456.0")
      && normalizeGroot("123456.0") === normalizeGroot(" 123456 "),
      "todas as grafias equivalentes colapsam no mesmo identificador");
check(normalizeGroot("ABC123") === "ABC123", "identificador alfanumérico não é convertido");
check(normalizeGroot("00123456") === "00123456",
      "zeros à esquerda são preservados — podem ser parte do identificador",
      normalizeGroot("00123456"));
check(normalizeGroot("00123456") !== normalizeGroot("123456"),
      "…e por isso não são unidos silenciosamente a outro id");
check(normalizeGroot("") === "" && normalizeGroot(null) === "" && normalizeGroot(undefined) === "",
      "vazio, nulo e indefinido não identificam ninguém");
check(!hasGroot("0") && !hasGroot("") && hasGroot("123456"),
      "'0' e vazio não contam como GROOT presente");

console.log("\nNormalização de data");
check(normalizeDateKey("10/08/2026") === "2026-08-10", "DD/MM/AAAA");
check(normalizeDateKey("2026-08-10") === "2026-08-10", "AAAA-MM-DD");
check(normalizeDateKey(new Date(Date.UTC(2026, 7, 10))) === "2026-08-10", "objeto Date (UTC)");
check(normalizeDateKey(46244) === "2026-08-10", "serial do Excel", normalizeDateKey(46244));
check(normalizeDateKey(20260810) === "2026-08-10", "inteiro AAAAMMDD interno");
check(personDayKey(123456, "10/08/2026") === "123456|2026-08-10",
      "a chave pessoa-dia é GROOT|AAAA-MM-DD", personDayKey(123456, "10/08/2026"));
check(personDayKey("", "10/08/2026") === null, "sem GROOT não existe chave pessoa-dia");

/* ================================================================ */
console.log("\nDeduplicação pessoa-dia");

// Teste 1 — mesmo GROOT, mesma data, mesma operação
{
  const r = dedup({ SVC: [reg("123456", "2026-08-10"), reg("123456", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "1 · mesmo GROOT, mesma data, mesma operação → 1 registro",
        `mantidos=${totalMantido(r)} dups=${r.resumo.duplicados}`);
}

// Teste 2 — mesmo GROOT, mesma data, OPERAÇÕES DIFERENTES  (o defeito corrigido)
{
  const r = dedup({ SVC: [reg("123456", "2026-08-10")],
                    XD:  [reg("123456", "2026-08-10")],
                    SD:  [reg("123456", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 2,
        "2 · mesmo GROOT e data em SVC, XD e SD → 1 registro válido, 2 duplicados",
        `mantidos=${totalMantido(r)} dups=${r.resumo.duplicados}`);
  check(r.porOperacao.SVC.rows.length === 1 && r.porOperacao.XD.rows.length === 0
        && r.porOperacao.SD.rows.length === 0,
        "…e quem fica é a PRIMEIRA ocorrência encontrada");
  check(r.descartados.length === 2 && r.descartados[0].mantidaEm === "SVC"
        && r.descartados[0].descartadaDe === "XD",
        "…com registro de onde veio a descartada e onde a mantida ficou",
        JSON.stringify(r.descartados[0]));
}

// Teste 3 — mesmo GROOT, datas diferentes
{
  const r = dedup({ SVC: [reg("123456", "2026-08-10"), reg("123456", "2026-08-11")] });
  check(totalMantido(r) === 2 && r.resumo.duplicados === 0,
        "3 · mesmo GROOT em datas diferentes → 2 registros (duas diárias)");
}

// Teste 4 — GROOTs diferentes, mesma data
{
  const r = dedup({ SVC: [reg("123456", "2026-08-10"), reg("654321", "2026-08-10")] });
  check(totalMantido(r) === 2 && r.resumo.duplicados === 0,
        "4 · GROOTs diferentes na mesma data → 2 registros");
}

// Teste 5 — 123456 e 123456.0
{
  const r = dedup({ SVC: [reg(123456, "2026-08-10")], XD: [reg("123456.0", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "5 · 123456 e 123456.0 na mesma data → duplicado");
}

// Teste 6 — " 123456 " e 123456
{
  const r = dedup({ SVC: [reg(" 123456 ", "2026-08-10")], XD: [reg(123456, "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "6 · ' 123456 ' e 123456 na mesma data → duplicado");
}

// Teste 7 — mesma pessoa-dia vinda de arquivos diferentes
{
  // dois arquivos carregados na mesma extração chegam como operações distintas
  const r = dedup({ "Arquivo A · SVC": [reg("123456", "2026-08-10")],
                    "Arquivo B · SVC": [reg("123456", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "7 · mesma pessoa-dia em arquivos diferentes → 1 único registro");
}

// Teste 8 — GROOT vazio nunca é deduplicado
{
  const r = dedup({ SVC: [reg("", "2026-08-10", "Fulano"), reg("", "2026-08-10", "Beltrano"),
                          reg(null, "2026-08-10", "Sicrano")] });
  check(totalMantido(r) === 3 && r.resumo.duplicados === 0,
        "8 · três registros sem GROOT na mesma data → nenhum é removido",
        `mantidos=${totalMantido(r)}`);
  check(r.resumo.semGroot === 3 && r.semGroot.length === 3,
        "…são contados e marcados para revisão", `semGroot=${r.resumo.semGroot}`);

  /* Uma contagem não é revisável. Quem revisa precisa achar a pessoa na
     origem para preencher o GROOT, e para isso precisa do nome, da filial e
     da data — não de "119 registros". */
  check(r.semGroot.every(s => s.op === "SVC" && s.date === "2026-08-10"),
        "…com a filial e a data de cada um", JSON.stringify(r.semGroot[0]));
  check(r.semGroot.map(s => s.nome).join("|") === "Fulano|Beltrano|Sicrano",
        "…e o nome de cada pessoa, na ordem do arquivo",
        r.semGroot.map(s => s.nome).join("|"));
  check(r.semGroot.every(s => s.reg && s.reg.empresa === "Empresa X"
                              && s.reg.cargo === "Auxiliar" && s.reg.escala === "6x1"
                              && s.reg.solicitante === "ID Logistics"),
        "…e o registro inteiro, para exibir empresa, cargo, escala e solicitante",
        JSON.stringify(r.semGroot[0] && r.semGroot[0].reg));
  check(r.semGroot[0].reg === r.porOperacao.SVC.rows[0],
        "…apontando para o MESMO registro que ficou na saída, não para uma cópia");
}
{
  // filiais diferentes: o painel precisa dizer onde o problema está concentrado
  const r = dedup({ SVC: [reg("", "2026-08-10", "Fulano")],
                    XD:  [reg("", "2026-08-10", "Beltrano"), reg(null, "2026-08-11", "Sicrano")] });
  const porOp = r.semGroot.reduce((m, s) => (m[s.op] = (m[s.op] || 0) + 1, m), {});
  check(porOp.SVC === 1 && porOp.XD === 2,
        "…e dá para agrupar por filial a partir da lista", JSON.stringify(porOp));
}

/* ================================================================ */
console.log("\nModos e transparência");
{
  const entrada = { SVC: [reg("1", "2026-08-10"), reg("1", "2026-08-11")],
                    XD:  [reg("1", "2026-08-10")] };
  const porDia = dedup(entrada, DEDUP_MODOS.DIA);
  check(totalMantido(porDia) === 2,
        "modo 'mesma data': 10/08 fica uma vez, 11/08 é outra diária", totalMantido(porDia));

  const porPeriodo = dedup(entrada, DEDUP_MODOS.PERIODO);
  check(totalMantido(porPeriodo) === 1,
        "modo 'período': uma linha por pessoa no período inteiro, entre operações",
        totalMantido(porPeriodo));

  const semDedup = dedup(entrada, DEDUP_MODOS.NAO);
  check(totalMantido(semDedup) === 3 && semDedup.resumo.duplicados === 0,
        "modo 'não remover': tudo permanece como está no arquivo");
}
{
  const r = dedup({ SVC: [reg("1", "2026-08-10"), reg("1", "2026-08-10")],
                    XD:  [reg("", "2026-08-10")] });
  check(r.resumo.encontrados === 3 && r.resumo.unicos === 2
        && r.resumo.duplicados === 1 && r.resumo.semGroot === 1,
        "o resumo bate: encontrados = únicos + duplicados",
        JSON.stringify(r.resumo));
  check(r.resumo.global === true, "o resumo declara que a deduplicação é global");
}
{
  // data ilegível não é motivo para descartar ninguém
  const r = dedup({ SVC: [reg("123456", "data ruim"), reg("123456", "data ruim")] });
  check(totalMantido(r) === 2 && r.resumo.duplicados === 0,
        "sem data legível não há pessoa-dia — os registros são preservados");
}

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
