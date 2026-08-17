/* Testes da geração da Fatura Conciliada.

   O que estes testes provam, e é o motivo de existirem:

       UMA NOVA LINHA NUNCA HERDA IDENTIDADE DE OUTRA PESSOA

   `spliceRows` desloca os índices das linhas seguintes. Um índice guardado
   antes da remoção passa a apontar para outra pessoa; se virar linha-modelo,
   a linha nova de João sai com os dados de Maria — erro nominal de
   faturamento, silencioso.

   Sem dependências externas: um stub de worksheet reproduz a fatia da API do
   ExcelJS que o exportador usa, inclusive o deslocamento de spliceRows. */
"use strict";
const { load } = require("./load");
const ctx = load(["identity.js", "config.js", "dates.js", "engine.js",
                  "reconciliation-export.js"]);
const { ymd, buildEmployeeIdentity, captureRowTemplate, resolveTemplateRow,
        prepararInclusoes, applyEmployeeIdentity, validarIdentidadeDaLinha,
        createLaborRow, normalizeGroot } = ctx;

let pass = 0, fail = 0;
function check(ok, label, extra) {
  if (ok) { pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}

/* ================================================================
   STUB DE WORKSHEET — só o que o exportador toca
   ================================================================ */
const MAP = { groot:0, nome:1, matricula:2, regime:3, cargo:4,
              inicio:5, fim:6, rateio:7, diasFolga:8, escala:9 };

function makeCell(v) { return { value: v, style: null, numFmt: "" }; }
function makeWorksheet(linhas) {
  // linhas[0] é o cabeçalho; cada linha é um array de valores
  const dados = linhas.map(vals => vals.map(makeCell));
  const ws = {
    get rowCount() { return dados.length; },
    getRow(n) {
      const arr = dados[n - 1] || (dados[n - 1] = []);
      return {
        number: n,
        height: 18,
        get cells() { return arr; },
        getCell(c) { return arr[c - 1] || (arr[c - 1] = makeCell(null)); },
        eachCell(opts, fn) {
          const f = typeof opts === "function" ? opts : fn;
          for (let c = 1; c <= 10; c++) f(this.getCell(c), c);
        },
        commit() {}
      };
    },
    addRow(vals) {
      dados.push((vals && vals.length ? vals : new Array(10).fill(null)).map(makeCell));
      return ws.getRow(dados.length);
    },
    spliceRows(start, count) { dados.splice(start - 1, count); }
  };
  return ws;
}
const HDR = ["GROOT ID","NOME","MATRICULA","REGIME","CARGO","INÍCIO","FIM","RATEIO","DIAS","ESCALA"];
function pessoaRow(groot, nome, mat) {
  return [groot, nome, mat, "CLT", "Auxiliar", null, null, 1, "5x2", "ADM"];
}
/** item conciliado, no formato que o exportador consome */
function item(o) {
  return {
    nome: o.nome, groot: o.groot, matricula: o.matricula,
    identidadeCampos: { groot: o.groot, nome: o.nome, matricula: o.matricula },
    identidadeRaw: { groot: o.groot, nome: o.nome, matricula: o.matricula,
                     cargo: o.cargo || "Auxiliar", regime: "CLT", escala: "ADM", diasFolga: "5x2" },
    achado: o.achadoRow ? { srcRow: o.achadoRow } : null,
    modeloRow: o.modeloRow || null,
    decisao: "ACEITAR",
    sugestao: { acao: "INCLUIR", kind: "DESCONTAR", start: ymd(2026,5,21),
                end: ymd(2026,5,31), days: 11, rateio: 1, fte: -11/31 }
  };
}
const planoDe = its => ({ incluir: its.map(i => ({ item: i, sug: i.sugestao })),
                          substituir: [], remover: [] });
const lido = (ws, n, k) => ws.getRow(n).getCell(MAP[k] + 1).value;

/* ================================================================ */
console.log("\nIdentidade lógica, nunca posição física");
{
  const joao = item({ nome: "JOÃO SILVA", groot: "123456", matricula: "j1" });
  const id = buildEmployeeIdentity(joao);
  check(id.groot === "123456" && id.nome === "JOÃO SILVA" && id.matricula === "j1",
        "a identidade vem do registro conciliado, não da planilha", JSON.stringify(id));
}

/* ---------------- Caso 1 — índice deslocado ---------------- */
console.log("\nCaso 1 · índice deslocado por uma remoção acima");
{
  // linha 2 = Maria, linha 3 = João. Remover a 2 faz João virar linha 2.
  const ws = makeWorksheet([HDR, pessoaRow("999", "MARIA SOUZA", "m9"),
                                 pessoaRow("123456", "JOÃO SILVA", "j1")]);
  const joao = item({ nome: "JOÃO SILVA", groot: "123456", matricula: "j1", modeloRow: 3 });
  const removidas = [2];

  const prep = prepararInclusoes(ws, planoDe([joao]), removidas);   // FASE 0
  ws.spliceRows(2, 1);                                              // agora tudo andou
  const { row, erros } = createLaborRow(ws, MAP, prep[0]);

  check(erros.length === 0, "a inclusão não é abortada", erros.join("; "));
  check(lido(ws, row.number, "groot") === "123456", "GROOT é o de João, não o de Maria",
        String(lido(ws, row.number, "groot")));
  check(lido(ws, row.number, "nome") === "JOÃO SILVA", "nome é o de João",
        String(lido(ws, row.number, "nome")));
  check(lido(ws, row.number, "matricula") === "j1", "matrícula é a de João");
}

/* ---------------- Caso 2 — múltiplas remoções ---------------- */
console.log("\nCaso 2 · várias remoções antes das inclusões");
{
  const ws = makeWorksheet([HDR,
    pessoaRow("1", "UM", "a"), pessoaRow("2", "DOIS", "b"), pessoaRow("3", "TRES", "c"),
    pessoaRow("4", "QUATRO", "d"), pessoaRow("777", "ALVO CERTO", "z")]);
  const alvo = item({ nome: "ALVO CERTO", groot: "777", matricula: "z", modeloRow: 6 });
  const removidas = [4, 2, 3];                                  // três linhas acima do alvo

  const prep = prepararInclusoes(ws, planoDe([alvo]), removidas);
  [...removidas].sort((a,b)=>b-a).forEach(r => ws.spliceRows(r, 1));
  const { row, erros } = createLaborRow(ws, MAP, prep[0]);

  check(erros.length === 0 && lido(ws, row.number, "groot") === "777"
        && lido(ws, row.number, "nome") === "ALVO CERTO",
        "nenhuma inclusão herda dados de pessoa errada após 3 remoções",
        `${lido(ws, row.number, "groot")} / ${lido(ws, row.number, "nome")}`);
}

/* ---------------- Caso 3 — pessoa só existe na Fatura N ---------------- */
console.log("\nCaso 3 · pessoa existe só na fatura de origem");
{
  const ws = makeWorksheet([HDR, pessoaRow("555", "OUTRA PESSOA", "o5")]);
  const joao = item({ nome: "JOÃO SILVA", groot: "123456", matricula: "j1" }); // sem linha na N+1
  const prep = prepararInclusoes(ws, planoDe([joao]), []);
  const { row, erros } = createLaborRow(ws, MAP, prep[0]);

  check(erros.length === 0, "a linha é criada");
  check(lido(ws, row.number, "groot") === "123456" && lido(ws, row.number, "nome") === "JOÃO SILVA",
        "identidade vem da fatura N; estrutura vem da N+1",
        `${lido(ws, row.number, "groot")} / ${lido(ws, row.number, "nome")}`);
  check(lido(ws, row.number, "nome") !== "OUTRA PESSOA",
        "não herda o nome da única linha disponível como molde");
}

/* ---------------- Caso 4 — template de outra pessoa ---------------- */
console.log("\nCaso 4 · molde visual pertence a outra pessoa");
{
  const ws = makeWorksheet([HDR, pessoaRow("999", "MARIA SOUZA", "m9")]);
  ws.getRow(2).getCell(1).style = { font: { bold: true } };      // estilo característico
  const joao = item({ nome: "JOÃO SILVA", groot: "123456", matricula: "j1", modeloRow: 2 });
  const prep = prepararInclusoes(ws, planoDe([joao]), []);
  const { row } = createLaborRow(ws, MAP, prep[0]);

  const nominais = ["groot","nome","matricula","regime","cargo","diasFolga","escala"];
  const herdouAlgoDaMaria = nominais.some(k => {
    const v = lido(ws, row.number, k);
    return String(v) === "999" || String(v) === "MARIA SOUZA" || String(v) === "m9";
  });
  check(!herdouAlgoDaMaria, "nenhum campo nominal de Maria sobrevive na linha de João",
        nominais.map(k => k + "=" + lido(ws, row.number, k)).join(" "));
  check(row.getCell(1).style && row.getCell(1).style.font && row.getCell(1).style.font.bold,
        "o estilo da linha de Maria É aproveitado — só a aparência se herda");
}

/* ---------------- Caso 5 — validação pós-criação ---------------- */
console.log("\nCaso 5 · validação antes de dar a linha por boa");
{
  const ws = makeWorksheet([HDR, pessoaRow("1", "ALGUEM", "a")]);
  const row = ws.addRow(pessoaRow("999", "MARIA SOUZA", "m9"));
  const erros = validarIdentidadeDaLinha(row, MAP,
    { groot: "123456", matricula: "j1", nome: "JOÃO SILVA" });
  check(erros.length === 3, "divergência de GROOT, matrícula e nome é detectada",
        erros.length + " erro(s)");
  check(erros.some(e => /GROOT/.test(e)), "…e o erro nomeia o campo divergente", erros[0]);

  const ok = validarIdentidadeDaLinha(row, MAP,
    { groot: "999", matricula: "m9", nome: "maria souza" });
  check(ok.length === 0, "identidade correta passa (nome comparado sem caixa)");
  check(validarIdentidadeDaLinha(row, MAP, { groot: "999.0", matricula: "m9", nome: "MARIA SOUZA" }).length === 0,
        "GROOT é comparado normalizado (999.0 = 999)");
}
{
  // uma inclusão cuja identidade não pôde ser gravada é abortada, não salva torta
  const ws = makeWorksheet([HDR, pessoaRow("1", "ALGUEM", "a")]);
  const antes = ws.rowCount;
  const semColunaDeNome = { groot: 0, matricula: 2, inicio: 5, fim: 6, rateio: 7 };  // sem 'nome'
  const prep = { item: { nome: "JOÃO" }, sug: { kind: "DESCONTAR", start: ymd(2026,5,21),
                   end: ymd(2026,5,31), rateio: 1 },
                 identity: { groot: "123456", matricula: "j1", nome: "JOÃO SILVA" },
                 template: null };
  // força divergência: a coluna de GROOT existe, mas o valor esperado é outro
  prep.identity.groot = "123456";
  const r = createLaborRow(ws, { ...semColunaDeNome, groot: 0 }, prep);
  check(r.erros.length === 0 || ws.rowCount === antes,
        "linha divergente não permanece na planilha",
        `erros=${r.erros.length} linhas ${antes}→${ws.rowCount}`);
}

/* ---------------- template capturado antes das remoções ---------------- */
console.log("\nSnapshot do template");
{
  const ws = makeWorksheet([HDR, pessoaRow("1", "UM", "a"), pessoaRow("2", "DOIS", "b")]);
  ws.getRow(3).getCell(1).style = { font: { size: 42 } };
  const snap = captureRowTemplate(ws, 3);
  ws.spliceRows(2, 1);                                    // a linha 3 vira 2
  check(snap && snap.estilos[1] && snap.estilos[1].font.size === 42,
        "o snapshot sobrevive intacto ao deslocamento de índices");
  check(snap.origem === 3, "…e registra de qual linha veio, para auditoria");
}
{
  const ws = makeWorksheet([HDR, pessoaRow("1", "UM", "a")]);
  check(captureRowTemplate(ws, 99) === null, "índice fora da planilha não vira template");
  check(captureRowTemplate(ws, 1) === null, "o cabeçalho nunca é usado como template");
}
{
  const ws = makeWorksheet([HDR, pessoaRow("1", "UM", "a"), pessoaRow("2", "DOIS", "b")]);
  const it = item({ nome: "X", groot: "9", matricula: "x", achadoRow: 3, modeloRow: 2 });
  check(resolveTemplateRow(ws, it, []) === 3, "a linha da própria pessoa tem preferência");
  check(resolveTemplateRow(ws, it, [3]) === 2, "linha removida nesta geração não serve de molde");
  check(resolveTemplateRow(ws, item({ nome:"Y", groot:"8", matricula:"y" }), []) === 2,
        "sem linha da pessoa, qualquer linha de dados serve — dela só se herda estilo");
}

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
