/* Ajustes MELI — reconstrução da cobrança e classificação das linhas
   ================================================================

   Módulo puro, sem DOM. Responde três perguntas que o motor de ajustes
   não respondia, e que são o coração da auditoria entre competências:

     1. o que foi EFETIVAMENTE faturado numa competência?
     2. cada linha de uma fatura é da competência corrente ou é um
        retroativo de outro mês?
     3. por quê? — cada resposta vem com o motivo em texto.

   O ponto conceitual que este arquivo formaliza: o snapshot do dia 15
   NÃO transforma todo mundo em mês cheio. Ele congela o que se sabia
   naquele dia e PROJETA a cobrança até o fim da competência. Quem foi
   admitido em 07/05 e estava ativo em 15/05 foi cobrado de 07/05 a
   31/05 — 25 dias, não 31.
   ================================================================ */
"use strict";

/* Classificação de uma linha dentro da fatura em que ela aparece. */
const LINE_CLASS = {
  CURRENT: "CURRENT_COMPETENCE",
  RETRO_ADD: "RETROACTIVE_ADD",
  RETRO_DISC: "RETROACTIVE_DISCOUNT",
  UNDETERMINED: "UNDETERMINED"
};
const LINE_CLASS_LABEL = {
  CURRENT_COMPETENCE: "Competência corrente",
  RETROACTIVE_ADD: "Acréscimo retroativo",
  RETROACTIVE_DISCOUNT: "Desconto retroativo",
  UNDETERMINED: "Indeterminado"
};

/** Uma linha que não identifica ninguém e não traz período nenhum não é um
 *  colaborador: é sobra de planilha — rodapé, subtotal, linha de formatação.
 *  Fatura real tem várias. Elas não devem virar apontamento de espécie alguma,
 *  porque não há o que conciliar nem o que revisar. */
function semSubstancia(e){
  const temNome=String(e.nome??"").trim()!=="";
  const temGroot=String(e.groot??"").trim()!==""&&String(e.groot).trim()!=="0";
  const temMat=String(e.matricula??"").trim()!=="";
  const temData=isValidYmd(e.inicio)||isValidYmd(e.fim);
  return !temNome && !temGroot && !temMat && !temData;
}

/** Competência (ano/mês) a que um período pertence, quando ele cabe todo
 *  dentro de um único mês. Devolve null se o período cruza meses. */
function competenceOfPeriod(start, end){
  if(!isValidYmd(start)||!isValidYmd(end)) return null;
  const a=ymdParts(start), b=ymdParts(end);
  if(a.y!==b.y||a.m!==b.m) return null;
  return {y:a.y, m:a.m};
}

/* ================================================================
   CLASSIFICAÇÃO DE LINHA

   O discriminador principal é a relação TEMPORAL entre o período da
   linha e a competência do arquivo — nunca o sinal isolado. O sinal
   entra depois, para separar acréscimo de desconto, e a ausência de
   DATA FIM é o que denuncia uma linha corrente que apenas carrega a
   data real de admissão de um mês anterior.
   ================================================================ */
function classifyLine(e, comp){
  const temInicio=isValidYmd(e.inicio), temFim=isValidYmd(e.fim);
  const rateio=(typeof e.rateio==="number"&&isFinite(e.rateio))?e.rateio:null;
  const neg=rateio!==null&&rateio<0;
  const fmtC=c=>c?MONTHS[c.m-1]+"/"+c.y:"—";

  if(!temInicio){
    return {classe:LINE_CLASS.UNDETERMINED, compOrigem:null,
      motivo:"A linha não tem DATA DE INÍCIO legível. Não há informação suficiente "
        +"para distinguir competência corrente de retroativo."};
  }

  // --- 1) período fechado inteiramente antes da competência do arquivo ---
  if(temFim && e.fim < comp.first){
    const co=competenceOfPeriod(e.inicio,e.fim)
      ||{y:ymdParts(e.fim).y, m:ymdParts(e.fim).m};
    return {
      classe: neg?LINE_CLASS.RETRO_DISC:LINE_CLASS.RETRO_ADD,
      compOrigem:co,
      motivo:"O período da linha ("+fmtYmd(e.inicio)+" a "+fmtYmd(e.fim)+") termina antes do "
        +"primeiro dia de "+comp.label+", portanto pertence a "+fmtC(co)+". "
        +(neg?"O rateio é negativo, o que caracteriza desconto retroativo."
             :"A quantidade é positiva, o que caracteriza acréscimo retroativo.")
    };
  }

  // --- 2) rateio negativo: crédito/estorno, nunca linha de competência ---
  if(neg){
    if(temFim && e.fim>=comp.first && e.fim<=comp.last){
      // negativo com período dentro da própria competência: não dá para afirmar
      return {classe:LINE_CLASS.UNDETERMINED, compOrigem:null,
        motivo:"O rateio é negativo, mas o período está dentro da própria competência "
          +comp.label+". Não há informação suficiente para dizer se é estorno da competência "
          +"corrente ou retroativo lançado com datas da competência errada. Revisão manual necessária."};
    }
    const co=temFim?{y:ymdParts(e.fim).y,m:ymdParts(e.fim).m}:null;
    return {classe:LINE_CLASS.RETRO_DISC, compOrigem:co,
      motivo:"O rateio é negativo: a linha representa um estorno, e não uma cobrança "
        +"da competência "+comp.label+"."};
  }

  // --- 3) sem DATA FIM: linha aberta na competência corrente ---
  if(!temFim){
    return {classe:LINE_CLASS.CURRENT, compOrigem:{y:comp.y,m:comp.m},
      motivo:"A DATA FIM está vazia: a linha segue aberta em "+comp.label
        +". A DATA DE INÍCIO ("+fmtYmd(e.inicio)+") é a admissão real da pessoa e não torna "
        +"a linha retroativa."};
  }

  // --- 4) período que alcança ou ultrapassa a competência do arquivo ---
  if(e.fim>=comp.first){
    return {classe:LINE_CLASS.CURRENT, compOrigem:{y:comp.y,m:comp.m},
      motivo:"O período da linha alcança "+comp.label+" (termina em "+fmtYmd(e.fim)
        +"), portanto é cobrança da competência corrente."};
  }

  return {classe:LINE_CLASS.UNDETERMINED, compOrigem:null,
    motivo:"A combinação de datas e quantidade não permite classificar a linha com segurança."};
}

const isRetroClass=c=>c===LINE_CLASS.RETRO_ADD||c===LINE_CLASS.RETRO_DISC;

/* ================================================================
   RECONSTRUÇÃO DA COBRANÇA ORIGINAL

   O que a competência N efetivamente cobrou desta pessoa, segundo o
   que se sabia no dia 15:

   - ativa no corte  → projeta da admissão (ou do dia 01) até o fim do
                       mês. É aqui que mora a diferença entre "estava no
                       snapshot" e "foi cobrada o mês inteiro";
   - saída já conhecida antes do corte → cobra só até a saída;
   - admitida depois do corte → não entrou no snapshot, não foi cobrada.
   ================================================================ */
function originalBilling(e, comp){
  if(!isValidYmd(e.inicio)) return null;
  if(e.inicio>comp.last){
    return {start:null,end:null,days:0,fte:0,cobrado:false,
      base:"Admissão posterior à competência — nada foi cobrado em "+comp.label+"."};
  }
  const rateio=(typeof e.rateio==="number"&&isFinite(e.rateio)&&e.rateio>0)?e.rateio:1;

  if(isActiveOnSnapshot(e,comp.snapshot)){
    const start=Math.max(e.inicio,comp.first);
    const end=comp.last;                       // projeção do corte até o fim do mês
    const days=diffDaysInclusive(start,end);
    return {start,end,days,cobrado:true,rateio,
      fte:(days/comp.days)*rateio,
      base:"Ativa no snapshot de "+fmtShort(comp.snapshot)+": a cobrança foi projetada de "
        +fmtShort(start)+" até o fim da competência."};
  }

  // não estava ativa no corte
  if(isValidYmd(e.fim) && e.fim<comp.snapshot){
    if(e.fim<comp.first){
      return {start:null,end:null,days:0,fte:0,cobrado:false,
        base:"Saída anterior a "+comp.label+" — nada foi cobrado nesta competência."};
    }
    const start=Math.max(e.inicio,comp.first), end=e.fim;
    const days=diffDaysInclusive(start,end);
    return {start,end,days,cobrado:true,rateio,
      fte:(days/comp.days)*rateio,
      base:"Saída em "+fmtShort(e.fim)+" já era conhecida no snapshot de "
        +fmtShort(comp.snapshot)+": a cobrança foi só até esse dia."};
  }

  return {start:null,end:null,days:0,fte:0,cobrado:false,
    base:"Admissão em "+fmtShort(e.inicio)+", posterior ao snapshot de "
      +fmtShort(comp.snapshot)+": a pessoa não entrou na fotografia do corte e não foi "
      +"cobrada em "+comp.label+"."};
}

/** Movimentação inferida a partir de um retroativo encontrado.
 *  Texto deliberadamente no condicional: o retroativo é evidência do
 *  fato, não o fato — a data real vem do RH, não da fatura. */
function inferMovement(achado, compOrigem){
  if(!achado) return null;
  if(achado.kind==="DESCONTAR"){
    const ultimo=addDays(achado.start,-1);
    return {texto:"Último dia faturável inferido: "+fmtYmd(ultimo),
      base:"desconto encontrado de "+fmtShort(achado.start)+" a "+fmtShort(achado.end)
        +" na fatura seguinte."};
  }
  return {texto:"Início de cobrança inferido: "+fmtYmd(achado.start),
    base:"acréscimo encontrado de "+fmtShort(achado.start)+" a "+fmtShort(achado.end)
      +" na fatura seguinte."};
}

if(typeof module!=="undefined"&&module.exports){
  module.exports={LINE_CLASS, LINE_CLASS_LABEL, classifyLine, isRetroClass,
    originalBilling, inferMovement, competenceOfPeriod, semSubstancia};
}
