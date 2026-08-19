/* Retorno Simulado — o que o cliente provavelmente reconhecerá
   ================================================================

   Módulo PURO, sem DOM. Depende de js/dates.js e js/config.js,
   carregados antes no index.html.

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

   A regra é a MESMA da Fusão de Linhas (js/fusao.js), e tem de ser:
   as duas reconstroem o mesmo número a partir do mesmo Labor.

     · cargo na lista de CARGOS_PREF (js/config.js) — auxiliar e
       operador; liderança e indiretos ficam fora por não estarem lá;
     · DATA DE INÍCIO legível, e o dia dentro de [início, fim], com fim
       vazio valendo até o fim do período;
     · o % RATEIO entra COM O SINAL QUE TEM.

   O sinal é onde este módulo já esteve errado. O PREF é uma quantidade
   FATURADA — o que a fatura apresenta, líquido dos estornos — e não um
   retrato do turno. Excluir os rateios negativos por serem "ajuste
   financeiro" inflava o PREF: nesta fatura eram 30 linhas de estorno,
   que somavam até 30 HC a mais num único dia de julho.

   As duas leituras continuam disponíveis: `pref` é o líquido, que vai
   ao confronto com o S&OP, e `bruto` conta só os positivos — quanta
   gente estava no turno. Só o primeiro é o PREF.
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
   que o usuário precisa ver: sair por cargo fora da lista é esperado
   para liderança, mas não para um cargo novo que ninguém declarou. */
const SIM_FORA = {
  ABS:         { chave:"abs",       texto:"linha de absenteísmo do template, não é pessoa" },
  /* Liderança e indiretos também saem por cargo, mas sair é o esperado
     deles — o aviso é para o cargo que ninguém declarou. A conta serve
     só para separar os dois casos, nunca para excluir: quem exclui é a
     lista de cargos, como na Fusão de Linhas. */
  INDIRETO:    { chave:"indireto",  texto:"cargo de liderança/indiretos — não entra no PREF" },
  CARGO:       { chave:"cargo",     texto:"cargo fora da lista do PREF" },
  SEM_INICIO:  { chave:"sem_inicio",texto:"sem DATA DE INÍCIO legível" }
};

const simNorm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"")
  .toLowerCase().replace(/\s+/g," ").trim();
const simEhIndireto = conta => /lideranca|indireto/.test(simNorm(conta));
const simEhAbs = l => simNorm(l.groot) === "abs" || /absenteismo/.test(simNorm(l.nome));

/* ================================================================
   QUEM ENTRA NO PREF

   Exatamente a regra da Fusão de Linhas (js/fusao.js): cargo na lista
   do PREF, DATA DE INÍCIO legível, fora a linha de ABS — e o rateio
   entra COM O SINAL QUE TEM.

   O sinal é o ponto, e é onde esta função já esteve errada. O PREF é
   uma quantidade FATURADA, não um retrato do turno: é o que a fatura
   apresenta ao cliente, líquido dos estornos. Uma linha de rateio -1
   referente a 27/07→31/07 reduz o que está sendo cobrado naqueles
   dias, e ignorá-la fazia o PREF sair alto — 30 linhas de estorno
   inflavam julho inteiro, até 30 HC num único dia.

   O headcount operacional bruto (só os positivos) continua sendo
   calculado e sai como `prefBruto`, porque as duas leituras são úteis:
   uma responde "quanto estou cobrando", a outra "quanta gente estava
   no turno". Só a primeira é o PREF.
   ================================================================ */
function simClassificarLinhas(labor){
  const dentro = [], fora = [];
  for(const l of labor){
    const marca = motivo => fora.push({ linha:l, motivo });
    if(simEhAbs(l)){ marca(SIM_FORA.ABS); continue; }
    if(!isValidYmd(l.inicio)){ marca(SIM_FORA.SEM_INICIO); continue; }
    if(!CARGOS_PREF.includes(simNorm(l.cargo))){
      marca(simEhIndireto(l.conta) ? SIM_FORA.INDIRETO : SIM_FORA.CARGO); continue;
    }
    dentro.push(l);
  }
  return { dentro, fora };
}

/* Rateio ausente ou ilegível vale 1, como na Fusão de Linhas: a linha
   existe e cobra uma pessoa. */
const simRateio = l =>
  (typeof l.rateio === "number" && isFinite(l.rateio)) ? l.rateio : 1;
const simAtivo = (l,d) => l.inicio <= d && (!isValidYmd(l.fim) || l.fim >= d);

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

  const { dentro, fora } = simClassificarLinhas(labor);

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
  const nNeg = dentro.filter(l => simRateio(l) < 0).length;
  if(nNeg) avisos.push({ tipo:"estorno", texto:
    nNeg+" linha(s) do PREF têm rateio negativo e entram SUBTRAINDO, como na Fusão de Linhas. "
    + "São estornos e ajustes retroativos, e o PREF é a quantidade faturada — o que a fatura "
    + "apresenta, líquido deles. O headcount bruto do turno, só com os lançamentos positivos, "
    + "aparece separado em cada dia." });
  const nSemInicio = fora.filter(f => f.motivo.chave === "sem_inicio").length;
  if(nSemInicio) avisos.push({ tipo:"indefinida", texto:
    nSemInicio+" linha(s) ficaram fora do PREF por não ter DATA DE INÍCIO legível. Sem a data "
    + "não há como saber em que dias elas estariam ativas." });

  const dias = [];
  let totalPref = 0, totalBruto = 0, totalCliente = 0, totalPos = 0, hcAcima = 0, hcAbaixo = 0;
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

    /* PREF é o líquido — a mesma soma da Fusão de Linhas, com o sinal
       do rateio. `bruto` conta só os positivos: é quanta gente estava
       no turno, e serve para explicar a diferença, nunca para comparar
       com o S&OP. */
    let pref = 0, bruto = 0;
    for(const l of dentro){
      if(!simAtivo(l,d)) continue;
      const r = simRateio(l);
      pref += r;
      if(r > 0) bruto += r;
    }
    pref = Math.round(pref*1e6)/1e6;
    bruto = Math.round(bruto*1e6)/1e6;

    const revisao = [];
    if(faltando.length) revisao.push("o S&OP de "+faltando.map(b=>b.rotulo).join(" e ")+" não tem valor para este dia.");

    const dia = { data:d, blocos:porBloco, revisao, pref, bruto };
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
      totalPref += pref; totalBruto += bruto; totalCliente += qCliente; totalPos += qPos;
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
      bruto: Math.round(totalBruto*100)/100,
      estornos: nNeg,
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
