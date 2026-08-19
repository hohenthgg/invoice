/* Retorno Simulado — o que o cliente provavelmente reconhecerá
   ================================================================

   Módulo PURO, sem DOM. Depende de js/dates.js, js/config.js e
   js/billing.js (classifyLine), carregados antes no index.html.

   O QUE ESTE MÓDULO RESPONDE

   Três números por dia, e a distância entre eles:

     PREF ENVIADO   o headcount que o Labor apresenta ao cliente
     Q CLIENTE      o que o S&OP do dia suporta (SVC + SD + …)
     Q PÓS PREVISTO o que provavelmente será reconhecido

   A regra da previsão é deliberadamente ASSIMÉTRICA:

       Q PÓS PREVISTO = MIN(PREF, Q CLIENTE)

   O cliente pode CORTAR o que foi apresentado acima do planejamento,
   mas não vai PAGAR o que nem sequer foi enviado. Enviar 130 contra um
   S&OP de 138 não vira 138 — vira 130, e os 8 de folga são um alerta
   de subfaturamento para conferir o Labor, não uma receita a receber.
   Prever `Q PÓS = Q CLIENTE` nos dois sentidos seria otimista no lado
   errado, e é o erro que este módulo existe para não cometer.

   ISTO É UMA PREVISÃO

   Não há retorno oficial nenhum aqui. Tudo o que sai daqui é rotulado
   como simulado, e o dia que não puder ser reconstruído com segurança
   sai como REVISÃO NECESSÁRIA em vez de sair com um número inventado.

   O QUE ENTRA NO PREF

   Só headcount operacional corrente:
     · conta LABOR DIRETO — liderança e indiretos não entram;
     · cargo na lista de CARGOS_PREF (js/config.js);
     · linha classificada como competência corrente por classifyLine.

   O último item é o que impede o erro mais caro do confronto: uma
   linha retroativa de rateio -1 referente a 20/07→31/07 é ajuste
   FINANCEIRO. Somá-la ao headcount diria que havia uma pessoa a menos
   trabalhando naqueles dias, o que é falso — ninguém desaparece do
   turno porque a fatura passada cobrou a mais.
   ================================================================ */
"use strict";

const SIM_STATUS = {
  REDUCAO:  "reducao",          // PREF acima do S&OP: risco de corte
  ALINHADO: "alinhado",
  SUB:      "subfaturamento",   // PREF abaixo do S&OP
  REVISAO:  "revisao"           // não dá para reconstruir com segurança
};
const SIM_STATUS_LABEL = {
  reducao:"Possível correção", alinhado:"Alinhado",
  subfaturamento:"Possível subfaturamento", revisao:"Revisão necessária"
};

/* Motivos de uma linha do Labor ficar fora do PREF. `avisa` marca os
   que o usuário precisa ver: sair por ser liderança é esperado, sair
   por cargo desconhecido não é. */
const SIM_FORA = {
  INDIRETO:    { chave:"indireto",    avisa:false, texto:"conta de liderança/indiretos — não é HC operacional do PREF" },
  ABS:         { chave:"abs",         avisa:false, texto:"linha de absenteísmo do template, não é pessoa" },
  RETRO:       { chave:"retro",       avisa:true,  texto:"lançamento retroativo — é ajuste financeiro, não HC do período" },
  INDEFINIDA:  { chave:"indefinida",  avisa:true,  texto:"não foi possível classificar a linha com segurança" },
  CARGO:       { chave:"cargo",       avisa:true,  texto:"cargo requer validação — não está na lista de cargos do PREF" },
  SEM_INICIO:  { chave:"sem_inicio",  avisa:true,  texto:"sem DATA DE INÍCIO legível" }
};

const simNorm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"")
  .toLowerCase().replace(/\s+/g," ").trim();
const simEhIndireto = conta => /lideranca|indireto/.test(simNorm(conta));
const simEhAbs = l => simNorm(l.groot) === "abs" || /absenteismo/.test(simNorm(l.nome));

/* ================================================================
   CLASSIFICAÇÃO DAS LINHAS DO LABOR
   ================================================================ */
/* A janela usada para separar corrente de retroativo é o PERÍODO
   FATURADO (16/07→15/08), não o mês-calendário da competência. Com o
   mês-calendário, toda linha que termina em julho cairia como
   retroativa — e julho é a primeira metade do próprio período cobrado.
   classifyLine só lê `first`, `last` e `label` do objeto, então basta
   entregar a janela certa. */
function simJanela(periodo, comp){
  return { y: comp ? comp.y : ymdParts(periodo.fim).y,
           m: comp ? comp.m : ymdParts(periodo.fim).m,
           first: periodo.ini, last: periodo.fim,
           label: fmtYmd(periodo.ini)+" a "+fmtYmd(periodo.fim) };
}

function simClassificarLinhas(labor, janela){
  const dentro = [], fora = [];
  for(const l of labor){
    const marca = motivo => fora.push({ linha:l, motivo });

    if(simEhAbs(l)){ marca(SIM_FORA.ABS); continue; }
    if(simEhIndireto(l.conta)){ marca(SIM_FORA.INDIRETO); continue; }
    if(!isValidYmd(l.inicio)){ marca(SIM_FORA.SEM_INICIO); continue; }

    /* Rateio negativo nunca é headcount, qualquer que seja o período.
       Vem antes de classifyLine de propósito: para o dinheiro, um
       negativo dentro da própria janela é ambíguo e a Conciliação o
       manda para revisão manual — mas para o HEADCOUNT não há
       ambiguidade nenhuma. Ninguém sai do turno porque a fatura
       passada cobrou a mais. */
    if(typeof l.rateio === "number" && l.rateio < 0){
      marca(SIM_FORA.RETRO); continue;
    }

    /* Daqui para baixo, a classificação é a MESMA de js/billing.js — a
       que a Conciliação usa para separar competência corrente de
       retroativo. Reimplementar aqui daria duas respostas para a mesma
       pergunta. */
    const cls = classifyLine(
      { inicio:l.inicio, fim:isValidYmd(l.fim)?l.fim:null, rateio:l.rateio }, janela);

    if(cls.classe === LINE_CLASS.RETRO_ADD || cls.classe === LINE_CLASS.RETRO_DISC){
      marca({ ...SIM_FORA.RETRO, detalhe:cls.motivo }); continue;
    }
    if(cls.classe === LINE_CLASS.UNDETERMINED){
      marca({ ...SIM_FORA.INDEFINIDA, detalhe:cls.motivo }); continue;
    }
    if(!CARGOS_PREF.includes(simNorm(l.cargo))){
      marca(SIM_FORA.CARGO); continue;
    }
    dentro.push(l);
  }
  return { dentro, fora };
}

/* ================================================================
   O CONFRONTO, DIA A DIA
   ================================================================ */
function simDiagnostico(dia){
  const n = v => (v === null || v === undefined || !isFinite(Number(v)))
    ? "—" : Number(Number(v).toFixed(2)).toLocaleString("pt-BR");

  /* O texto do dia em revisão não cita valores: é justamente por não
     ter valor confiável que ele está em revisão. */
  if(dia.status === SIM_STATUS.REVISAO){
    return "Não foi possível reconstruir este dia com segurança: "+dia.revisao.join(" ")
      + " Nenhum valor foi previsto para não inventar número.";
  }
  const fontes = dia.blocos.map(b => b.rotulo+" "+n(b.valor)).join(" + ");
  if(dia.status === SIM_STATUS.REDUCAO){
    return "Foram enviados "+n(dia.gap)+" HC acima do S&OP total deste dia ("+fontes
      + " = "+n(dia.qCliente)+"). Caso o cliente limite o reconhecimento ao planejamento, "
      + "há risco de redução de "+n(dia.gap)+" HC.";
  }
  if(dia.status === SIM_STATUS.SUB){
    return "O Labor enviado está "+n(-dia.gap)+" HC abaixo do S&OP disponível ("+fontes
      + " = "+n(dia.qCliente)+"). Não presumir que o cliente aumentará automaticamente a "
      + "cobrança: verifique se existem pessoas faturáveis faltando no Labor.";
  }
  return "PREF e S&OP total batem em "+n(dia.qCliente)+" ("+fontes+"). Sem correção prevista.";
}

function simularRetorno(dados){
  const labor = dados.labor || [], blocos = dados.blocos || [];
  const ini = dados.periodo.ini, fim = dados.periodo.fim;
  const comp = dados.comp;
  const avisos = [];

  if(!blocos.length) return { erro:"Nenhum bloco de S&OP foi identificado na planilha operacional." };

  const janela = simJanela({ ini, fim }, comp);
  const { dentro, fora } = simClassificarLinhas(labor, janela);

  /* Cargos fora da lista viram aviso nominal: o usuário precisa saber
     QUAL cargo ficou de fora para decidir se ele deveria entrar. */
  const cargosDuvidosos = [...new Set(fora.filter(f => f.motivo.chave === "cargo")
    .map(f => String(f.linha.cargo || "(cargo em branco)").trim()))];
  if(cargosDuvidosos.length){
    avisos.push({ tipo:"cargo", texto:
      "Cargo requer validação: "+cargosDuvidosos.join(", ")+". "
      + "Estas linhas estão em LABOR DIRETO mas o cargo não consta na lista de cargos do PREF, "
      + "então NÃO foram somadas — o PREF pode estar subestimado nos dias em que elas estão ativas. "
      + "Confirme se o cargo entra no PREF antes de usar a previsão." });
  }
  const nRetro = fora.filter(f => f.motivo.chave === "retro").length;
  if(nRetro) avisos.push({ tipo:"retro", texto:
    nRetro+" linha(s) retroativa(s) foram excluídas do PREF. Retroativo é ajuste financeiro de "
    + "outra competência: somá-lo ao headcount diria que havia gente a mais ou a menos "
    + "trabalhando nestes dias, o que não é verdade." });
  const nIndef = fora.filter(f => f.motivo.chave === "indefinida").length;
  if(nIndef) avisos.push({ tipo:"indefinida", texto:
    nIndef+" linha(s) não puderam ser classificadas entre competência corrente e retroativo. "
    + "Ficaram fora do PREF e os dias em que estão ativas foram marcados para revisão." });

  /* Dias em que há linha indefinida ativa: o PREF daquele dia não é
     confiável, então ele sai como REVISÃO em vez de sair com número. */
  const diasDuvidosos = new Set();
  for(const f of fora){
    if(f.motivo.chave !== "indefinida") continue;
    const l = f.linha;
    if(!isValidYmd(l.inicio)) continue;
    for(let d = Math.max(l.inicio, ini); d <= fim; d = addDays(d,1)){
      if(isValidYmd(l.fim) && d > l.fim) break;
      diasDuvidosos.add(d);
    }
  }

  const dias = [];
  let totalPref = 0, totalCliente = 0, totalPos = 0, hcAcima = 0, hcAbaixo = 0;
  const contagem = { reducao:0, alinhado:0, subfaturamento:0, revisao:0 };

  for(let d = ini; d <= fim; d = addDays(d,1)){
    /* Célula vazia do S&OP é AUSÊNCIA, não zero. `Number(null)` vale 0,
       e deixar isso passar transformaria um dia sem planejamento num
       S&OP zerado — o PREF inteiro daquele dia viraria "risco de corte"
       inventado. Por isso o valor só existe se for número de verdade. */
    const porBloco = blocos.map(b => {
      const bruto = Object.prototype.hasOwnProperty.call(b.dias, d) ? b.dias[d] : null;
      const n = Number(bruto);
      return { rotulo:b.rotulo,
               valor: (bruto === null || bruto === undefined || bruto === "" || !isFinite(n)) ? null : n };
    });
    const faltando = porBloco.filter(b => b.valor === null || !isFinite(b.valor));

    let pref = 0;
    for(const l of dentro){
      if(l.inicio > d) continue;
      if(isValidYmd(l.fim) && l.fim < d) continue;
      pref += (typeof l.rateio === "number" && isFinite(l.rateio)) ? l.rateio : 1;
    }
    pref = Math.round(pref*1e6)/1e6;

    const revisao = [];
    if(faltando.length) revisao.push("o S&OP de "+faltando.map(b=>b.rotulo).join(" e ")+" não tem valor para este dia.");
    if(diasDuvidosos.has(d)) revisao.push("há linha do Labor que não pôde ser classificada entre corrente e retroativo.");

    const dia = { data:d, blocos:porBloco, revisao, pref };
    if(revisao.length){
      dia.qCliente = null; dia.qPos = null; dia.gap = null; dia.correcao = null;
      dia.status = SIM_STATUS.REVISAO;
    } else {
      const qCliente = porBloco.reduce((s,b) => s + b.valor, 0);
      const qPos = Math.min(pref, qCliente);
      dia.qCliente = qCliente;
      dia.qPos = qPos;
      dia.gap = Math.round((pref - qCliente)*1e6)/1e6;
      dia.correcao = Math.round((qPos - pref)*1e6)/1e6;
      dia.status = dia.gap > 0 ? SIM_STATUS.REDUCAO
                 : dia.gap < 0 ? SIM_STATUS.SUB
                 : SIM_STATUS.ALINHADO;
      totalPref += pref; totalCliente += qCliente; totalPos += qPos;
      if(dia.gap > 0) hcAcima += dia.gap; else hcAbaixo += -dia.gap;
    }
    dia.diagnostico = simDiagnostico(dia);
    contagem[dia.status]++;
    dias.push(dia);
  }

  return {
    periodo:{ ini, fim }, comp, blocos:blocos.map(b => b.rotulo), dias, avisos,
    linhas:{ dentro:dentro.length, fora },
    totais:{
      pref: Math.round(totalPref*100)/100,
      cliente: Math.round(totalCliente*100)/100,
      pos: Math.round(totalPos*100)/100,
      /* O número que resume o risco: HC-dia que o cliente pode cortar. */
      hcEmRisco: Math.round(hcAcima*100)/100,
      hcAbaixo: Math.round(hcAbaixo*100)/100,
      dias: dias.length, ...contagem
    }
  };
}

/* ================================================================
   O ARQUIVO SIMULADO

   A aba de retorno é escrita com os MESMOS cabeçalhos que a Fusão de
   Linhas procura, para o arquivo poder ser usado lá como um retorno de
   teste. A Fusão localiza a aba pelo cabeçalho, não pelo nome, então o
   nome pode dizer com todas as letras que o arquivo é simulado.
   ================================================================ */
const SIM_AVISO_METADADOS =
  "ARQUIVO SIMULADO. Q Pós foi estimado por MIN(PREF, soma do S&OP dos blocos). "
  + "Não corresponde ao retorno oficial do cliente.";

function simLinhasRetorno(sim){
  const linhas = [];
  for(const dia of sim.dias){
    if(dia.status === SIM_STATUS.REVISAO){
      linhas.push({ data:dia.data, tipo:"FULL_TIME", pref:dia.pref, qPos:null, desvio:null,
        ocorrencia:"REVISÃO NECESSÁRIA — "+dia.diagnostico });
      continue;
    }
    linhas.push({ data:dia.data, tipo:"FULL_TIME", pref:dia.pref, qPos:dia.qPos,
      desvio: Math.round((dia.pref - dia.qPos)*1e6)/1e6,
      ocorrencia:"PREVISÃO SIMULADA — "+dia.diagnostico });
  }
  return linhas;
}

if(typeof module !== "undefined" && module.exports){
  module.exports = { simularRetorno, simClassificarLinhas, simLinhasRetorno,
                     SIM_STATUS, SIM_STATUS_LABEL, SIM_AVISO_METADADOS };
}
