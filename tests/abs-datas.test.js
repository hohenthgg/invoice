/* Datas escritas no Excel pela aba Calcular ABS.
   ================================================================

   O que este teste protege:

       O SERIAL DA DATA PRECISA SER INTEIRO

   O resumo compensa o absenteísmo com uma fórmula viva:

       COUNTIFS('Diaristas SVC'!$B:$B, DATE(2026,7,20), …)

   `DATE(2026,7,20)` é o serial 46223, inteiro. Se a célula da coluna B
   guardar 46223,125 — meia-noite local gravada como 03:00Z em UTC-3 — a
   comparação não casa com nada e TODO o abate recalcula para zero assim
   que alguém abre a planilha. Os valores em cache continuam certos, o que
   torna o defeito invisível até o Excel recalcular.

   Um ambiente em UTC não reproduz nada disso: lá meia-noite local já é
   meia-noite UTC. Por isso o teste força o fuso de São Paulo.
   ================================================================ */
"use strict";
process.env.TZ = "America/Sao_Paulo";

let pass = 0, fail = 0;
function check(ok, label, extra) {
  if (ok) { pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}

/* a mesma conversão usada em js/abs.js, na fronteira com o Excel */
const dataExcel = d => new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
/* como o ExcelJS transforma um Date em serial: instante UTC desde 1899-12-30 */
const serial = d => (d.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
/* o que a fórmula DATE(ano,mes,dia) do Excel vale */
const DATE = (y, m, d) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;

console.log("\nSerial de data no Excel (fuso America/Sao_Paulo)");
check(new Date(2026, 6, 20).getTimezoneOffset() === 180,
      "o fuso do teste é mesmo UTC-3 — sem isso o defeito não aparece",
      "offset=" + new Date(2026, 6, 20).getTimezoneOffset());

const local = new Date(2026, 6, 20);            // como as datas vivem no módulo
check(serial(local) === 46223.125,
      "meia-noite local vira 46223,125 — a fração de 3h que quebrava o COUNTIFS",
      String(serial(local)));
check(serial(dataExcel(local)) === 46223,
      "dataExcel() devolve o serial inteiro", String(serial(dataExcel(local))));
check(serial(dataExcel(local)) === DATE(2026, 7, 20),
      "…e ele casa com DATE(2026,7,20) da fórmula do resumo",
      serial(dataExcel(local)) + " vs " + DATE(2026, 7, 20));
check(serial(local) !== DATE(2026, 7, 20),
      "…enquanto a data local NÃO casa — era isso que zerava todo o abate");

console.log("\nO dia continua o mesmo depois da conversão");
for (const [y, m, d] of [[2026,7,20], [2026,1,1], [2026,12,31], [2026,10,18], [2026,2,21]]) {
  const orig = new Date(y, m - 1, d), conv = dataExcel(orig);
  check(conv.getUTCFullYear() === y && conv.getUTCMonth() === m - 1 && conv.getUTCDate() === d,
        `${String(d).padStart(2,"0")}/${String(m).padStart(2,"0")}/${y} sai como o mesmo dia`,
        conv.toISOString());
}
/* 18/10/2026 e 21/02/2026 estão perto das viradas de horário de verão que o
   Brasil já usou: se o horário de verão voltar, o offset muda no meio do
   período e uma conversão ingênua deslocaria o dia. */
check([[2026,10,18],[2026,2,21]].every(([y,m,d])=>{
        const c = dataExcel(new Date(y, m - 1, d));
        return c.getUTCDate() === d && Number.isInteger(serial(c));
      }),
      "inclusive nas datas de virada de horário de verão, com serial inteiro");

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
