/* Prioridade da ID Logistics no abate "Ambos" — aba Calcular ABS.
   ================================================================

   O que este teste protege:

       COM "AMBOS", O MELI SÓ ENTRA DEPOIS DE ESGOTADA A ID

   A escolha "Ambos" no diálogo de compensação junta dois pools de
   diaristas com custos diferentes: o pedido pela ID Logistics e o
   pedido pelo MELI. O total abatido num dia é o mesmo nos dois casos
   — nunca passa do déficit — mas a ATRIBUIÇÃO importa: diarista
   pedido pelo MELI é custo do MELI, e usá-lo antes de gastar a ID
   transfere para a conta errada um abate que a ID já cobria.

   Como o defeito é invisível no total (145 continua 145), só um teste
   sobre a repartição pega uma regressão aqui.

   Duas frentes:
     1. propriedades da repartição, sobre a mesma aritmética do módulo;
     2. um guarda no fonte de js/abs.js — a regra vive inline dentro de
        `processar`, então a única forma de garantir que ninguém a
        reescreveu de volta para "soma os dois pools e reparte" é ler
        as linhas que a implementam.
   ================================================================ */
"use strict";
const fs = require("fs");
const path = require("path");

let pass = 0, fail = 0;
function check(ok, label, extra) {
  if (ok) { pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}

/* A repartição do abate de um dia, exatamente como js/abs.js a calcula. */
function repartir(dispId, dispMeli, falta) {
  const usadosId = Math.min(dispId, falta);
  const usadosMeli = Math.min(dispMeli, falta - usadosId);
  return { usadosId, usadosMeli, usados: usadosId + usadosMeli };
}

console.log("\nRepartição do abate diário com \"Ambos\"");

check(repartir(5, 5, 3).usadosMeli === 0,
      "sobrando ID, o MELI não é tocado",
      JSON.stringify(repartir(5, 5, 3)));
check(repartir(2, 9, 6).usadosId === 2 && repartir(2, 9, 6).usadosMeli === 4,
      "a ID vai até o fim e o MELI completa só o que faltou",
      JSON.stringify(repartir(2, 9, 6)));
check(repartir(0, 4, 3).usadosId === 0 && repartir(0, 4, 3).usadosMeli === 3,
      "sem ID no dia, o MELI cobre sozinho",
      JSON.stringify(repartir(0, 4, 3)));
check(repartir(3, 4, 0).usados === 0,
      "dia sem déficit não consome diarista de fonte nenhuma",
      JSON.stringify(repartir(3, 4, 0)));
check(repartir(1, 1, 9).usados === 2,
      "pool menor que o déficit: abate tudo o que há, e para aí",
      JSON.stringify(repartir(1, 1, 9)));

console.log("\nPropriedades — valem para qualquer combinação");
let pTotal = true, pFallback = true, pIdMax = true, pNaoPiora = true;
for (let dispId = 0; dispId <= 8; dispId++)
  for (let dispMeli = 0; dispMeli <= 8; dispMeli++)
    for (let falta = 0; falta <= 12; falta++) {
      const r = repartir(dispId, dispMeli, falta);
      /* o total é o mesmo de um pool único: escolher "Ambos" não abate mais
         nem menos, só muda de quem veio */
      if (r.usados !== Math.min(dispId + dispMeli, falta)) pTotal = false;
      /* MELI só aparece com a ID esgotada */
      if (r.usadosMeli > 0 && r.usadosId !== dispId) pFallback = false;
      /* a ID é usada até o teto: ou acabou o pool, ou acabou o déficit */
      if (r.usadosId !== Math.min(dispId, falta)) pIdMax = false;
      /* somar MELI ao pool nunca reduz o que a ID abate — o abate atribuído
         à ID é o mesmo que ela teria sozinha, em "Somente ID Logistics" */
      if (r.usadosId !== repartir(dispId, 0, falta).usadosId) pNaoPiora = false;
    }
check(pTotal, "o total abatido é sempre min(pool inteiro, déficit)");
check(pFallback, "o MELI nunca entra com diarista da ID sobrando no dia");
check(pIdMax, "a ID é sempre consumida até o limite do pool ou do déficit");
check(pNaoPiora, "o abate atribuído à ID é o mesmo de \"Somente ID Logistics\"");

console.log("\nA regra está no fonte de js/abs.js");
const src = fs.readFileSync(path.join(__dirname, "..", "js", "abs.js"), "utf8");
const semEspaco = s => s.replace(/\s+/g, "");

check(semEspaco(src).includes(semEspaco("const usadosId=Math.min(dispId, falta)")),
      "o abate da ID é calculado primeiro, contra o déficit cheio");
check(semEspaco(src).includes(semEspaco("const usadosMeli=Math.min(dispMeli, falta-usadosId)")),
      "o do MELI só recebe o que sobrou do déficit");
/* No pool, o mesmo GROOT pode aparecer nos dois solicitantes no mesmo dia
   (pedido pela ID e pelo MELI). Contá-lo dos dois lados inflaria `disp` e
   ainda daria ao MELI uma pessoa que a ID já traz: a linha do MELI precisa
   olhar o conjunto da ID antes de aceitar alguém. */
check(/meliSet\.add/.test(src) && /!idSet\.has\(r\.id\)\s*&&\s*!meliSet\.has\(r\.id\)/.test(src),
      "quem já está no pool da ID no dia não é recontado no do MELI");

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
