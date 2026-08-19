/* Retorno Simulado — js/simulacao.js
   ================================================================

   O que estes testes protegem:

       A ASSIMETRIA DA PREVISÃO

           Q PÓS PREVISTO = MIN(PREF, S&OP)

   O cliente pode CORTAR o que foi apresentado acima do planejamento,
   mas não paga o que nem sequer foi enviado. A tentação de escrever
   `Q PÓS = S&OP` é grande — a fórmula fica mais simples e "prevê" mais
   receita — e está errada num dos dois sentidos: com PREF 130 contra
   S&OP 138 ela prometeria 138, uma cobrança que ninguém enviou.

   E o segundo, que já esteve errado aqui:

       O PREF É QUANTIDADE FATURADA, NÃO RETRATO DO TURNO

   Uma linha de rateio -1 referente a 27/07→31/07 REDUZ o que a fatura
   apresenta naqueles dias. Tratá-la como "ajuste financeiro que não é
   headcount" e tirá-la da conta inflava o PREF — na fatura real eram
   30 estornos, até 30 HC a mais num único dia de julho.

   A regra é a mesma da Fusão de Linhas, e tem de ser: as duas
   reconstroem o mesmo número a partir do mesmo Labor. Divergir aqui
   daria dois PREFs para a mesma fatura.
   ================================================================ */
"use strict";
const { load } = require("./load.js");

const ctx = load(["dates.js","config.js","simulacao.js"],
                 ["SIM_STATUS","SIM_AVISO_METADADOS","CARGOS_PREF"]);
const { simularRetorno, simClassificarLinhas, simLinhasRetorno,
        buildCompetence, ymd, SIM_STATUS } = ctx;

let pass = 0, fail = 0;
function check(ok, label, extra){
  if(ok){ pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}

const COMP = buildCompetence(2026,8);
const PER  = { ini: ymd(2026,7,16), fim: ymd(2026,8,15) };

/* Uma pessoa do quadro, ativa o período todo, salvo indicação. */
const P = o => Object.assign({ linha:0, groot:"7000001", nome:"PESSOA",
  cargo:"Auxiliar de Apoio LOG I", conta:"LABOR DIRETO", regime:"Efetivo",
  inicio: ymd(2026,1,1), fim:null, rateio:1 }, o);

/* Blocos de S&OP com um valor constante, ou um mapa dia→valor. */
function bloco(rotulo, valorOuMapa){
  const dias = {};
  for(let d = PER.ini; d <= PER.fim; d = ctx.addDays(d,1)){
    dias[d] = typeof valorOuMapa === "number" ? valorOuMapa
            : (Object.prototype.hasOwnProperty.call(valorOuMapa,d) ? valorOuMapa[d] : null);
  }
  return { rotulo, dias };
}
const rodar = (labor, blocos) =>
  simularRetorno({ labor, blocos, periodo:PER, comp:COMP });
const noDia = (sim, d) => sim.dias.find(x => x.data === d);

/* ================================================================
   1–3. A soma do S&OP é por DATA, e cobre todas as operações
   ================================================================ */
console.log("\nS&OP total = soma dos blocos da própria data");

/* Os três dias que o usuário pediu para conferir, com os valores reais
   da planilha operacional de Varginha. */
const svc = {}, sd = {};
svc[ymd(2026,7,16)] = 120; sd[ymd(2026,7,16)] = 16;
svc[ymd(2026,7,19)] = 80;  sd[ymd(2026,7,19)] = 12;
svc[ymd(2026,8,10)] = 120; sd[ymd(2026,8,10)] = 17;
const tresDias = rodar([P({ rateio:1 })], [bloco("SVC",svc), bloco("SD",sd)]);

check(noDia(tresDias, ymd(2026,7,16)).qCliente === 136,
      "16/07: SVC 120 + SD 16 = 136", String(noDia(tresDias, ymd(2026,7,16)).qCliente));
check(noDia(tresDias, ymd(2026,7,19)).qCliente === 92,
      "19/07: SVC 80 + SD 12 = 92", String(noDia(tresDias, ymd(2026,7,19)).qCliente));
check(noDia(tresDias, ymd(2026,8,10)).qCliente === 137,
      "10/08: SVC 120 + SD 17 = 137", String(noDia(tresDias, ymd(2026,8,10)).qCliente));

/* Só o SVC não é o S&OP de Varginha: a filial tem SVC + SD. */
const soSvc = rodar([P()], [bloco("SVC",120)]);
const comSd = rodar([P()], [bloco("SVC",120), bloco("SD",16)]);
check(noDia(soSvc, PER.ini).qCliente === 120 && noDia(comSd, PER.ini).qCliente === 136,
      "esquecer o bloco SD muda o S&OP total — os dois entram na soma");

/* Um bloco sem valor para o dia não vira zero: vira revisão. Zero
   silencioso seria um S&OP menor e um risco de corte inventado. */
const furo = {};
furo[ymd(2026,7,20)] = null;
const comFuro = rodar([P()], [bloco("SVC",120), bloco("SD",furo)]);
check(noDia(comFuro, ymd(2026,7,20)).status === SIM_STATUS.REVISAO,
      "dia sem valor num dos blocos sai como REVISÃO, não como zero");
check(noDia(comFuro, ymd(2026,7,20)).qCliente === null,
      "…e sem número previsto: não se fabrica valor");

/* ================================================================
   4–5. A previsão é assimétrica
   ================================================================ */
console.log("\nQ Pós previsto = MIN(PREF, S&OP)");

/* PREF 140 contra S&OP 136: o cliente pode cortar 4. */
const acima = rodar(Array.from({length:140},(_,i) => P({ groot:"A"+i })),
                    [bloco("SVC",120), bloco("SD",16)]);
const dAcima = noDia(acima, PER.ini);
check(dAcima.pref === 140 && dAcima.qCliente === 136,
      "PREF 140 contra S&OP 136", dAcima.pref+" / "+dAcima.qCliente);
check(dAcima.qPos === 136, "…Q Pós previsto é 136", String(dAcima.qPos));
check(dAcima.correcao === -4, "…correção prevista é -4", String(dAcima.correcao));
check(dAcima.gap === 4, "…e o gap é +4");
check(dAcima.status === SIM_STATUS.REDUCAO, "…classificado como possível correção");

/* PREF 130 contra S&OP 136: o cliente NÃO paga o que não foi enviado. */
const abaixo = rodar(Array.from({length:130},(_,i) => P({ groot:"B"+i })),
                     [bloco("SVC",120), bloco("SD",16)]);
const dAbaixo = noDia(abaixo, PER.ini);
check(dAbaixo.qPos === 130,
      "PREF 130 contra S&OP 136: Q Pós previsto é 130, NÃO 136 — o erro que este módulo existe para não cometer",
      String(dAbaixo.qPos));
check(dAbaixo.correcao === 0, "…correção automática prevista é 0", String(dAbaixo.correcao));
check(dAbaixo.gap === -6, "…e o gap acusa 6 de subfaturamento", String(dAbaixo.gap));
check(dAbaixo.status === SIM_STATUS.SUB, "…classificado como possível subfaturamento");
check(/não presumir que o cliente aumentará/i.test(dAbaixo.diagnostico),
      "…e o diagnóstico avisa para não presumir aumento automático");

const igual = rodar(Array.from({length:136},(_,i) => P({ groot:"C"+i })),
                    [bloco("SVC",120), bloco("SD",16)]);
const dIgual = noDia(igual, PER.ini);
check(dIgual.qPos === 136 && dIgual.correcao === 0 && dIgual.status === SIM_STATUS.ALINHADO,
      "PREF igual ao S&OP: alinhado, sem correção prevista");

/* O total em risco é a soma só do que está ACIMA — o que está abaixo
   não compensa, porque não vira receita. */
check(acima.totais.hcEmRisco === 4*31 && acima.totais.hcAbaixo === 0,
      "HC-dia em risco soma só os excessos, dia a dia",
      acima.totais.hcEmRisco+" / "+acima.totais.hcAbaixo);
check(abaixo.totais.hcEmRisco === 0 && abaixo.totais.hcAbaixo === 6*31,
      "…e o subfaturamento é contado separado, sem abater o risco");

/* ================================================================
   6. O estorno entra no PREF, subtraindo
   ================================================================ */
console.log("\nO PREF é o líquido: o rateio entra com o sinal que tem");

const comRetro = rodar([
  P({ groot:"D1" }),
  P({ groot:"D2" }),
  /* estorno de 27/07 a 31/07: reduz o que a fatura apresenta nesses dias */
  P({ groot:"D3", rateio:-1, inicio: ymd(2026,7,27), fim: ymd(2026,7,31) })
], [bloco("SVC",2)]);
check(noDia(comRetro, ymd(2026,7,29)).pref === 1,
      "no dia coberto pelo estorno o PREF cai de 2 para 1 — é o que a fatura cobra",
      String(noDia(comRetro, ymd(2026,7,29)).pref));
check(noDia(comRetro, ymd(2026,7,26)).pref === 2,
      "fora do período do estorno o PREF fica intacto",
      String(noDia(comRetro, ymd(2026,7,26)).pref));
check(noDia(comRetro, ymd(2026,7,29)).bruto === 2,
      "…e o BRUTO do mesmo dia continua 2: no turno havia duas pessoas",
      String(noDia(comRetro, ymd(2026,7,29)).bruto));
check(comRetro.avisos.some(a => a.tipo === "estorno"),
      "a presença de estornos vira aviso, para a queda no PREF não surpreender");
check(comRetro.totais.estornos === 1,
      "…e o total de linhas de estorno é informado", String(comRetro.totais.estornos));

/* Este é o defeito concreto que a mudança conserta: com o estorno fora
   da conta, o dia 29/07 sairia com 2 e o gap contra o S&OP seria 0 —
   um dia "alinhado" que na verdade está 1 abaixo. */
check(noDia(comRetro, ymd(2026,7,29)).gap === -1,
      "o gap do dia usa o líquido, não o bruto",
      String(noDia(comRetro, ymd(2026,7,29)).gap));

/* ================================================================
   7. Liderança e indiretos não entram sozinhos
   ================================================================ */
console.log("\nO que entra no PREF");

const comIndireto = rodar([
  P({ groot:"E1" }),
  P({ groot:"E2", conta:"LABOR LIDERANÇA E INDIRETOS", cargo:"Supervisor de Operação" }),
  P({ groot:"E3", conta:"LABOR LIDERANÇA E INDIRETOS", cargo:"Líder de Operação" })
], [bloco("SVC",1)]);
check(noDia(comIndireto, PER.ini).pref === 1,
      "liderança e indiretos ficam fora — o cargo deles não está na lista do PREF",
      String(noDia(comIndireto, PER.ini).pref));

const comAbs = rodar([
  P({ groot:"F1" }),
  P({ groot:"ABS", nome:" ABSENTEISMO", cargo:" ABSENTEISMO", regime:"FIXO/CLT" })
], [bloco("SVC",1)]);
check(noDia(comAbs, PER.ini).pref === 1,
      "a linha artificial de ABS não vira pessoa");

/* Cargo desconhecido não é chutado para dentro nem esquecido: fica de
   fora E vira aviso nominal. */
const cargoNovo = rodar([
  P({ groot:"G1" }),
  P({ groot:"G2", cargo:"Conferente de Expedição" })
], [bloco("SVC",2)]);
check(noDia(cargoNovo, PER.ini).pref === 1,
      "cargo fora da lista não é somado ao PREF — não se inventa elegibilidade");
const avisoCargo = cargoNovo.avisos.find(a => a.tipo === "cargo");
check(!!avisoCargo && /Conferente de Expedição/.test(avisoCargo.texto),
      "…e o aviso diz QUAL cargo ficou de fora, para o usuário decidir");
check(!!avisoCargo && /subestimado/.test(avisoCargo.texto),
      "…avisando que o PREF pode estar subestimado por causa disso");

/* Os dois cargos operacionais conhecidos entram. */
const doisCargos = rodar([
  P({ groot:"H1", cargo:"Auxiliar de Apoio LOG I" }),
  P({ groot:"H2", cargo:"Operador Transpaleteira" })
], [bloco("SVC",2)]);
check(noDia(doisCargos, PER.ini).pref === 2,
      "auxiliar e operador transpaleteira contam no PREF");

/* ================================================================
   O PREF por dia
   ================================================================ */
console.log("\nReconstrução do PREF, dia a dia");

const janela = rodar([
  P({ groot:"I1", inicio: ymd(2026,7,20), fim: ymd(2026,7,25) })
], [bloco("SVC",1)]);
check(noDia(janela, ymd(2026,7,19)).pref === 0
   && noDia(janela, ymd(2026,7,20)).pref === 1
   && noDia(janela, ymd(2026,7,25)).pref === 1
   && noDia(janela, ymd(2026,7,26)).pref === 0,
      "o intervalo é inclusivo nas duas pontas");

const aberto = rodar([P({ groot:"J1", inicio: ymd(2026,7,20), fim:null })], [bloco("SVC",1)]);
check(noDia(aberto, PER.fim).pref === 1,
      "sem DATA FIM, a pessoa conta até o fim do período");

const meio = rodar([P({ groot:"K1", rateio:0.5 }), P({ groot:"K2", rateio:0.5 })],
                   [bloco("SVC",1)]);
check(noDia(meio, PER.ini).pref === 1,
      "dois rateios de 0,5 somam 1 HC — o rateio entra na conta");

/* ================================================================
   S&OP DE VALOR FIXO

   Além do S&OP diário da planilha, a tela aceita um valor fixo do mês —
   o número que o contrato fecha — aplicado igual em todos os dias. Para
   o motor isso é só um bloco único e constante: a assimetria e o resto
   das regras não mudam, e é isso que se prova aqui.
   ================================================================ */
console.log("\nS&OP de valor fixo");

const fixo = rodar([
  ...Array.from({length:164},(_,i) => P({ groot:"X"+i }))
], [bloco("S&OP fixo", 182)]);
const dFixo = noDia(fixo, PER.ini);
check(dFixo.qCliente === 182, "o valor fixo vale para o dia", String(dFixo.qCliente));
check(fixo.dias.every(x => x.qCliente === 182), "…e para todos os 31 dias do período");
check(dFixo.pref === 164 && dFixo.qPos === 164,
      "PREF 164 contra fixo 182: o previsto é o PREF, não o S&OP",
      dFixo.pref+" / "+dFixo.qPos);
check(dFixo.gap === -18 && dFixo.correcao === 0,
      "…gap de 18 abaixo, e nenhuma correção presumida", dFixo.gap+" / "+dFixo.correcao);
check(dFixo.status === SIM_STATUS.SUB, "…classificado como possível subfaturamento");

/* Acima do fixo, a correção volta a aparecer — a assimetria é a mesma. */
const fixoAcima = rodar(Array.from({length:200},(_,i) => P({ groot:"Y"+i })),
                        [bloco("S&OP fixo", 182)]);
const dFA = noDia(fixoAcima, PER.ini);
check(dFA.qPos === 182 && dFA.correcao === -18,
      "PREF 200 contra fixo 182: previsto 182 e correção de -18",
      dFA.qPos+" / "+dFA.correcao);

/* ================================================================
   PARIDADE COM A FUSÃO DE LINHAS

   As duas abas reconstroem o MESMO número a partir do mesmo Labor: a
   Fusão para comparar com o retorno oficial, a simulação para prever
   esse retorno. Divergirem daria dois PREFs para a mesma fatura, e foi
   exatamente o que aconteceu enquanto este módulo excluía os estornos.

   A expressão abaixo é a de js/fusao.js, copiada:

       labor.reduce((s,p) => s + ((p.elig && p.pt===pt)
         && p.ini<=d && (p.fim==null || p.fim>=d) ? p.rateio : 0), 0)
   ================================================================ */
console.log("\nParidade com a Fusão de Linhas");

const elegivel = l => ctx.CARGOS_PREF.includes(
  String(l.cargo).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim());
const contaFusao = (labor, d) => labor.reduce((acc,p) =>
  acc + ((elegivel(p) && p.inicio <= d && (p.fim == null || p.fim >= d)) ? p.rateio : 0), 0);

/* Um Labor com os formatos que a fatura real tem: aberto, fechado,
   estorno, meio rateio, cargo de liderança e cargo desconhecido. */
const laborReal = [
  P({ groot:"F1", inicio: ymd(2026,3,23), fim:null }),
  P({ groot:"F2", inicio: ymd(2026,6,3),  fim: ymd(2026,8,4) }),
  P({ groot:"F3", inicio: ymd(2026,7,20), fim: ymd(2026,7,25) }),
  P({ groot:"F4", inicio: ymd(2026,7,27), fim: ymd(2026,7,31), rateio:-1 }),
  P({ groot:"F5", inicio: ymd(2026,8,1),  fim:null, rateio:0.5 }),
  P({ groot:"F6", inicio: ymd(2026,1,1),  fim:null, cargo:"Operador Transpaleteira" }),
  P({ groot:"F7", inicio: ymd(2026,1,1),  fim:null, cargo:"Supervisor de Operação",
      conta:"LABOR LIDERANÇA E INDIRETOS" }),
  P({ groot:"F8", inicio: ymd(2026,1,1),  fim:null, cargo:"Conferente de Expedição" })
];
const paridade = rodar(laborReal, [bloco("SVC",100)]);
const divergencias = [];
for(const dia of paridade.dias){
  const esperado = Math.round(contaFusao(laborReal, dia.data)*1e6)/1e6;
  if(dia.pref !== esperado) divergencias.push(ctx.fmtYmd(dia.data)+": "+dia.pref+" != "+esperado);
}
check(divergencias.length === 0,
      "o PREF de cada dia bate com a conta da Fusão de Linhas, dia a dia",
      divergencias.slice(0,3).join(" · "));

/* A prova de que o teste não é vazio: se os estornos voltassem a ser
   excluídos, os dias cobertos por eles divergiriam. */
const semEstorno = laborReal.filter(l => l.rateio > 0);
check(contaFusao(laborReal, ymd(2026,7,29)) !== contaFusao(semEstorno, ymd(2026,7,29)),
      "…e o dia do estorno realmente muda de valor entre as duas contas — o teste morde");

/* ================================================================
   A saída
   ================================================================ */
console.log("\nO arquivo simulado");

const retorno = simLinhasRetorno(acima);
check(retorno.length === 31, "uma linha por dia do período", String(retorno.length));
check(retorno.every(l => l.tipo === "FULL_TIME"), "Employee_Type é FULL_TIME neste MVP");
check(retorno[0].pref === 140 && retorno[0].qPos === 136 && retorno[0].desvio === 4,
      "Qtd. PREF, Q Pós e Desvio saem coerentes com o comparativo",
      JSON.stringify(retorno[0]));
check(retorno.every(l => /SIMULADA|REVIS/.test(l.ocorrencia)),
      "toda ocorrência se identifica como previsão — nunca como retorno confirmado");
check(/SIMULADO/.test(ctx.SIM_AVISO_METADADOS) && /Não corresponde ao retorno oficial/.test(ctx.SIM_AVISO_METADADOS),
      "o aviso de metadados diz que o arquivo é simulado e não é o oficial");

const comRevisao = simLinhasRetorno(comFuro);
const linhaRev = comRevisao.find(l => /REVIS/.test(l.ocorrencia));
check(!!linhaRev && linhaRev.qPos === null && linhaRev.desvio === null,
      "o dia em revisão sai sem Q Pós e sem desvio, em vez de sair com número inventado");

/* ================================================================
   Nada é alterado
   ================================================================ */
console.log("\nO simulador não altera nada");

const original = [P({ groot:"L1", rateio:1, inicio: ymd(2026,7,20), fim: ymd(2026,7,25) })];
const copia = JSON.parse(JSON.stringify(original));
rodar(original, [bloco("SVC",5)]);
check(JSON.stringify(original) === JSON.stringify(copia),
      "as linhas do Labor saem da simulação exatamente como entraram");

const semBloco = simularRetorno({ labor:[P()], blocos:[], periodo:PER, comp:COMP });
check(!!semBloco.erro, "sem nenhum bloco de S&OP o módulo recusa em vez de prever no vazio");

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
