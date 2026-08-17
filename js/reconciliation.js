/* Ajustes MELI — motor de conciliação entre duas faturas subsequentes
   ================================================================

   O modo "Projetar ajustes" (js/engine.js) responde: dada UMA fatura,
   quais ajustes deveriam aparecer na competência seguinte? Este módulo
   responde a pergunta seguinte, com DUAS faturas na mão:

     o ajuste que deveria ter surgido na fatura N+1 realmente surgiu?

   O confronto não é "linha N versus linha N+1". Para cada pessoa o
   motor reconstrói uma pequena história de faturamento — o que foi
   cobrado, o que o corte congelou, o que se infere ter acontecido
   depois, que ajuste era devido e o que de fato apareceu — e só então
   classifica. A aritmética de datas vem de dates.js, o ajuste esperado
   de engine.js e a reconstrução da cobrança de billing.js: aqui não há
   nenhuma segunda implementação da regra do dia 15.

   O motor não decide nada. Detecta, compara, classifica e explica; a
   decisão de alterar uma fatura pertence ao usuário.
   ================================================================ */
"use strict";

const RECON_STATUS = {
  CONCILIADO:  "CONCILIADO",
  AUSENTE:     "AUSENTE",
  PERIODO:     "PERIODO_DIVERGENTE",
  FTE:         "FTE_DIVERGENTE",
  SINAL:       "SINAL_INCORRETO",
  DUPLICADO:   "DUPLICADO",
  SEM_ORIGEM:  "SEM_ORIGEM",
  AMBIGUA:     "IDENTIDADE_AMBIGUA",
  SEM_ID:      "SEM_IDENTIFICADOR",
  INDETERMINADO:"INDETERMINADO",
  REVISAO:     "REVISAO"
};
const RECON_CONF = { ALTA:"ALTA", MEDIA:"MEDIA", REVISAO:"REVISAO" };

/* Vermelho só para ausência ou conflito relevante; âmbar para o que é
   parcial; azul/cinza para o que é informação e pede validação. */
const RECON_META = {
  CONCILIADO:        {label:"Conciliado",              tone:"ok",   icon:"✓"},
  AUSENTE:           {label:"Ajuste ausente",          tone:"bad",  icon:"✖"},
  PERIODO_DIVERGENTE:{label:"Período divergente",      tone:"warn", icon:"⚠"},
  FTE_DIVERGENTE:    {label:"FTE divergente",          tone:"warn", icon:"⚠"},
  SINAL_INCORRETO:   {label:"Sinal incorreto",         tone:"bad",  icon:"✖"},
  DUPLICADO:         {label:"Possível duplicidade",    tone:"bad",  icon:"✖"},
  SEM_ORIGEM:        {label:"Retroativo sem origem",   tone:"info", icon:"?"},
  IDENTIDADE_AMBIGUA:{label:"Identidade ambígua",      tone:"info", icon:"?"},
  SEM_IDENTIFICADOR: {label:"Sem identificador",        tone:"warn", icon:"⚠"},
  INDETERMINADO:     {label:"Indeterminado",           tone:"info", icon:"?"},
  REVISAO:           {label:"Revisão manual",          tone:"info", icon:"?"}
};

const FTE_TOL = 1e-6;

/* Erros de validação que dizem respeito só ao cadastro, não ao cálculo. */
const ERROS_DE_IDENTIFICADOR = new Set(["GROOT ID vazio","matrícula vazia"]);
const PLACEHOLDER_ID = "§sem-id§";

/* ================================================================
   1. IDENTIDADE DAS PESSOAS

   Ordem: GROOT ID, depois matrícula. Nome NUNCA une registros — só
   entra como reforço visual. O Excel devolve o mesmo identificador ora
   como texto, ora como número: "123456", 123456 e 123456.0 são o mesmo.
   ================================================================ */
/* A regra de identidade é uma só no aplicativo inteiro — js/identity.js.
   Duas implementações da mesma normalização é como se cria a situação em que
   a Conciliação acha que são a mesma pessoa e a Extração acha que não. */
function normId(v){ return normalizeGroot(v); }
function hasId(v){ return hasGroot(v); }

function buildPersonIndex(lists){
  const midToGids=new Map(), gidToMids=new Map();
  lists.forEach(list=>list.forEach(e=>{
    const g=hasId(e.groot)?normId(e.groot):"", m=hasId(e.matricula)?normId(e.matricula):"";
    if(g&&m){
      if(!midToGids.has(m)) midToGids.set(m,new Set());
      midToGids.get(m).add(g);
      if(!gidToMids.has(g)) gidToMids.set(g,new Set());
      gidToMids.get(g).add(m);
    }
  }));
  const ambiguous=new Set();
  midToGids.forEach((gids,m)=>{ if(gids.size>1) ambiguous.add("M:"+m); });

  return {
    keyOf(e){
      if(hasId(e.groot)) return "G:"+normId(e.groot);
      if(hasId(e.matricula)){
        const m=normId(e.matricula), gids=midToGids.get(m);
        if(gids && gids.size===1) return "G:"+[...gids][0];   // ponte segura
        return "M:"+m;
      }
      return "";
    },
    fonteOf(e){
      if(hasId(e.groot)) return "GROOT";
      if(hasId(e.matricula)) return "MATRICULA";
      return "NENHUMA";
    },
    isAmbiguous(e){
      if(!hasId(e.groot) && hasId(e.matricula)) return ambiguous.has("M:"+normId(e.matricula));
      if(hasId(e.groot) && hasId(e.matricula)){
        const gids=midToGids.get(normId(e.matricula));
        return !!(gids && gids.size>1);
      }
      return false;
    },
    hasSplitMatricula(e){
      if(!hasId(e.groot)) return false;
      const mids=gidToMids.get(normId(e.groot));
      return !!(mids && mids.size>1);
    }
  };
}

/* ================================================================
   2. SEPARAÇÃO DAS LINHAS  (delega a classificação a billing.js)
   ================================================================ */
function isRetroLine(e, comp){ return isRetroClass(classifyLine(e,comp).classe); }

function splitInvoiceLines(employees, comp){
  const normais=[], retroativos=[], indeterminados=[], ignoradas=[];
  employees.forEach(e=>{
    /* Sobra de planilha não entra em nenhum bucket de análise: não é pessoa,
       não tem período, não há o que conciliar. Só é contada, para que o total
       de linhas lidas continue batendo com o arquivo. */
    if(semSubstancia(e)){ ignoradas.push(e); return; }
    const c=classifyLine(e,comp);
    e.lineClassification=c.classe;          // campo interno, usado pela interface
    e.lineClassReason=c.motivo;
    e.lineCompOrigem=c.compOrigem;
    if(isRetroClass(c.classe)) retroativos.push(e);
    else if(c.classe===LINE_CLASS.UNDETERMINED) indeterminados.push(e);
    else normais.push(e);
  });
  return {normais, retroativos, indeterminados, ignoradas};
}

/** Retroativo já normalizado para comparação com o esperado. */
function describeRetro(e, compAlvo){
  if(!isValidYmd(e.inicio)||!isValidYmd(e.fim)) return null;
  const start=e.inicio, end=e.fim;
  if(end<start) return null;
  const days=diffDaysInclusive(start,end);
  const rateio=(typeof e.rateio==="number"&&isFinite(e.rateio))?e.rateio:1;
  const kind=e.lineClassification===LINE_CLASS.RETRO_DISC?"DESCONTAR"
            :e.lineClassification===LINE_CLASS.RETRO_ADD?"ACRESCENTAR"
            :(rateio<0?"DESCONTAR":"ACRESCENTAR");
  return {emp:e, kind, start, end, days,
    rateio:Math.abs(rateio),
    fte:(days/compAlvo.days)*(kind==="DESCONTAR"?-Math.abs(rateio):Math.abs(rateio)),
    classe:e.lineClassification, classeMotivo:e.lineClassReason,
    dentroDaCompetencia: start>=compAlvo.first && end<=compAlvo.last};
}

/* ================================================================
   3. COMPETÊNCIAS SUBSEQUENTES
   ================================================================ */
function monthIndex(comp){ return comp.y*12+comp.m; }
function competencesAreSequential(compA, compB){
  return monthIndex(compB)-monthIndex(compA)===1;
}
function checkSequence(compA, compB){
  if(!compA||!compB) return {ok:false, kind:"incompleto", msg:"Carregue as duas faturas."};
  const d=monthIndex(compB)-monthIndex(compA);
  if(d===1) return {ok:true, kind:"valida", msg:"Sequência válida"};
  if(d===-1) return {ok:false, kind:"invertida",
    msg:"As faturas parecem estar invertidas — "+compB.label+" vem antes de "+compA.label+"."};
  if(d===0) return {ok:false, kind:"igual",
    msg:"As duas faturas são da mesma competência ("+compA.label+")."};
  return {ok:false, kind:"distante",
    msg:"Competências não subsequentes — "+compA.label+" → "+compB.label+" ("+Math.abs(d)+" meses de distância)."};
}

/* ================================================================
   4. DIFERENÇA ENTRE PERÍODOS

   Devolve os trechos que existem no esperado e não no encontrado (e
   vice-versa), para a interface dizer exatamente qual dia não fechou.
   ================================================================ */
function faltantes(exp, ach){
  const out=[];
  if(!ach) { if(exp) out.push({start:exp.start,end:exp.end,days:exp.days}); return out; }
  if(ach.start>exp.start) out.push({start:exp.start, end:addDays(ach.start,-1),
    days:diffDaysInclusive(exp.start,addDays(ach.start,-1))});
  if(ach.end<exp.end) out.push({start:addDays(ach.end,1), end:exp.end,
    days:diffDaysInclusive(addDays(ach.end,1),exp.end)});
  return out;
}
function excedentes(exp, ach){
  const out=[];
  if(!ach||!exp) return out;
  if(ach.start<exp.start) out.push({start:ach.start, end:addDays(exp.start,-1),
    days:diffDaysInclusive(ach.start,addDays(exp.start,-1))});
  if(ach.end>exp.end) out.push({start:addDays(exp.end,1), end:ach.end,
    days:diffDaysInclusive(addDays(exp.end,1),ach.end)});
  return out;
}

/* ================================================================
   5. MOTOR
   ================================================================ */
function reconcile(input){
  const origem=input.origem, seguinte=input.seguinte;
  const compN=origem.comp, compNext=seguinte.comp;
  const movimentacoes=input.movimentacoes||null;

  const idx=buildPersonIndex([origem.employees, seguinte.employees]);

  const splitN=splitInvoiceLines(origem.employees, compN);

  /* O motor de projeção exige GROOT E matrícula e, sem os dois, devolve ERRO —
     descartando o ajuste. Mas identificador não é o que DEFINE o ajuste: quem
     define são as datas. O identificador serve para achar a pessoa na fatura
     seguinte. Então quando o único problema da linha é identificador faltando,
     o ajuste é calculado de todo modo; o que muda é a capacidade de conferir. */
  const paraAnalisar=[], soFaltaId=[];
  splitN.normais.forEach(e=>{
    const errs=validateEmployee(e);
    if(errs.length && errs.every(x=>ERROS_DE_IDENTIFICADOR.has(x))) soFaltaId.push(e);
    else paraAnalisar.push(e);
  });

  const proj=analyze(paraAnalisar, compN);          // ajuste esperado = motor de projeção
  const esperados=proj.adjustments.slice();

  /* Recalcula os que só tropeçaram no identificador. Quem TEM ao menos um
     identificador é rastreável e entra no confronto normal; quem não tem nenhum
     vira apontamento próprio, porque não há como procurá-lo na fatura seguinte —
     nome nunca é chave. Se não houver ajuste devido, nada é dito. */
  const semIdComAjuste=[];
  soFaltaId.forEach(e=>{
    const suprido={...e, groot:e.groot||PLACEHOLDER_ID, matricula:e.matricula||PLACEHOLDER_ID};
    const r=detectAdjustment(suprido, compN);
    if(!r || r.kind==="ERRO") return;               // nada devido → nada a apontar
    r.emp=e;                                        // identidade real, não a suprida
    r.faltaIdentificador=true;
    if(hasId(e.groot)||hasId(e.matricula)) esperados.push(r);
    else semIdComAjuste.push(r);
  });

  const splitNext=splitInvoiceLines(seguinte.employees, compNext);
  const encontrados=[];
  splitNext.retroativos.forEach(e=>{
    const d=describeRetro(e, compN);
    if(!d) return;
    if(d.end<compN.first || d.start>compN.last) return;   // retroativo de outro mês
    d.consumido=false;
    encontrados.push(d);
  });

  const porPessoa=new Map();
  encontrados.forEach(d=>{
    const k=idx.keyOf(d.emp);
    if(!porPessoa.has(k)) porPessoa.set(k,[]);
    porPessoa.get(k).push(d);
  });

  /* Linha da MESMA pessoa dentro da fatura N+1: serve de TEMPLATE de
     estrutura para uma eventual inclusão. A identidade nunca vem daqui
     quando a pessoa não existe na N+1 — vem do registro dela na N. */
  const modeloPorPessoa=new Map();
  splitNext.normais.forEach(e=>{ const k=idx.keyOf(e);
    if(k&&!modeloPorPessoa.has(k)) modeloPorPessoa.set(k,e.srcRow); });
  splitNext.retroativos.forEach(e=>{ const k=idx.keyOf(e);
    if(k&&!modeloPorPessoa.has(k)) modeloPorPessoa.set(k,e.srcRow); });

  const linhasNextPorPessoa=new Map();
  seguinte.employees.forEach(e=>{ const k=idx.keyOf(e);
    if(!linhasNextPorPessoa.has(k)) linhasNextPorPessoa.set(k,[]);
    linhasNextPorPessoa.get(k).push(e); });

  const items=[];
  const overlap=(a,b)=>a.start<=b.end && b.start<=a.end;

  // --- 5.1 para cada ajuste devido, procurar o aplicado ---
  esperados.forEach(exp=>{
    const key=idx.keyOf(exp.emp);
    const cands=(porPessoa.get(key)||[]).filter(c=>!c.consumido);
    const ambiguo=idx.isAmbiguous(exp.emp)||!key;

    const exatos=cands.filter(c=>c.start===exp.start&&c.end===exp.end);
    const sobrepostos=cands.filter(c=>overlap(c,exp));

    let status, alerts=[], conf, confMotivo, achado=null, achados=[], diagnostico;
    let diffDias=0, diffFte=0;

    const dim={identidade:true, competencia:true, periodo:null, sinal:null, fte:null};

    if(exatos.length>1){
      status=RECON_STATUS.DUPLICADO; alerts=[status];
      conf=RECON_CONF.ALTA; confMotivo="mesma pessoa e período idêntico em mais de uma linha";
      achados=exatos; achado=exatos[0];
      exatos.forEach(c=>c.consumido=true);
      dim.periodo=true; dim.sinal=achado.kind===exp.kind; dim.fte=Math.abs(achado.fte-exp.fte)<=FTE_TOL;
      diagnostico="A fatura seguinte traz "+exatos.length+" lançamentos para o mesmo período ("
        +fmtShort(exp.start)+" a "+fmtShort(exp.end)+"). Possível "
        +(exp.kind==="DESCONTAR"?"desconto":"acréscimo")+" em duplicidade — confirme qual linha deve permanecer.";
    } else if(exatos.length===1){
      achado=exatos[0]; achado.consumido=true; achados=[achado];
      diffFte=achado.fte-exp.fte;
      dim.periodo=true;
      dim.sinal=achado.kind===exp.kind;
      dim.fte=Math.abs(diffFte)<=FTE_TOL;
      if(!dim.sinal) alerts.push(RECON_STATUS.SINAL);
      if(!dim.fte&&dim.sinal) alerts.push(RECON_STATUS.FTE);
      if(!alerts.length){
        status=RECON_STATUS.CONCILIADO; conf=RECON_CONF.ALTA;
        confMotivo="mesmo "+(idx.fonteOf(exp.emp)==="GROOT"?"GROOT":"identificador")
          +", competência de origem clara, período fechado e FTE idêntico";
        diagnostico="Linha compatível encontrada na fatura seguinte: identidade, competência de origem, "
          +"período, sinal e FTE conferem.";
      } else {
        status=alerts[0]; conf=RECON_CONF.ALTA;
        confMotivo="identidade e período fecham; a divergência está "+(dim.sinal?"na quantidade":"no sentido do lançamento");
        diagnostico=!dim.sinal
          ?"O período confere, mas o lançamento está no sentido oposto: esperado "
            +(exp.kind==="DESCONTAR"?"desconto":"acréscimo")+" e encontrado "
            +(achado.kind==="DESCONTAR"?"desconto":"acréscimo")+"."
          :"Período correto, quantidade diferente: esperado FTE "+fmtFte(exp.fte)
            +" (rateio "+fmtPct(exp.rateio)+") e encontrado "+fmtFte(achado.fte)
            +" (rateio "+fmtPct(achado.rateio)+").";
      }
    } else if(sobrepostos.length>1){
      status=RECON_STATUS.REVISAO; alerts=[status];
      conf=RECON_CONF.REVISAO; confMotivo="mais de uma linha candidata sobrepondo o período esperado";
      achados=sobrepostos; achado=sobrepostos[0];
      sobrepostos.forEach(c=>c.consumido=true);
      dim.periodo=false;
      diagnostico="Há mais de um lançamento retroativo sobrepondo o período esperado. "
        +"A correspondência comporta mais de uma leitura — requer conferência manual.";
    } else if(sobrepostos.length===1){
      achado=sobrepostos[0]; achado.consumido=true; achados=[achado];
      diffDias=achado.days-exp.days;
      diffFte=achado.fte-exp.fte;
      dim.periodo=false;
      dim.sinal=achado.kind===exp.kind;
      dim.fte=Math.abs(diffFte)<=FTE_TOL;
      if(!dim.sinal){ alerts.push(RECON_STATUS.SINAL); }
      alerts.push(RECON_STATUS.PERIODO);
      /* FTE só vira alerta próprio quando o RATEIO difere. Um período menor já
         produz FTE menor por aritmética — apontar isso à parte seria ruído. */
      if(dim.sinal && Math.abs(achado.rateio-exp.rateio)>FTE_TOL) alerts.push(RECON_STATUS.FTE);
      status=dim.sinal?RECON_STATUS.PERIODO:RECON_STATUS.SINAL;
      conf=RECON_CONF.MEDIA;
      confMotivo="identidade segura, período apenas parcialmente compatível";
      const falta=faltantes(exp,achado), sobra=excedentes(exp,achado);
      diagnostico="A fatura seguinte contém o "+(exp.kind==="DESCONTAR"?"desconto":"acréscimo")
        +", mas o período não coincide: esperado "+fmtShort(exp.start)+" a "+fmtShort(exp.end)
        +" ("+exp.days+" dia"+(exp.days===1?"":"s")+") e encontrado "+fmtShort(achado.start)+" a "
        +fmtShort(achado.end)+" ("+achado.days+" dia"+(achado.days===1?"":"s")+")."
        +(falta.length?" Não conciliado: "+falta.map(f=>descreverTrecho(f)).join(" e ")+".":"")
        +(sobra.length?" Excedente: "+sobra.map(f=>descreverTrecho(f)).join(" e ")+".":"");
    } else {
      status=RECON_STATUS.AUSENTE; alerts=[status];
      conf=ambiguo?RECON_CONF.REVISAO:RECON_CONF.ALTA;
      confMotivo=ambiguo?"identificação ambígua entre as faturas"
        :"identidade segura e nenhuma linha candidata na fatura seguinte";
      dim.periodo=false; dim.sinal=false; dim.fte=false;
      diffDias=-exp.days; diffFte=-exp.fte;
      diagnostico="Ajuste esperado pela regra do corte não localizado na fatura "+compNext.label+". "
        +"Nenhuma linha retroativa desta pessoa referente a "+compN.label+" foi encontrada.";
    }

    if(ambiguo){
      dim.identidade=false;
      conf=RECON_CONF.REVISAO;
      confMotivo="identificação ambígua: a matrícula aparece ligada a mais de um GROOT";
      if(!alerts.includes(RECON_STATUS.AMBIGUA)) alerts.push(RECON_STATUS.AMBIGUA);
      diagnostico+=" A identificação desta pessoa é ambígua entre as duas faturas — confirme GROOT e matrícula antes de agir.";
    } else if(idx.hasSplitMatricula(exp.emp) && conf===RECON_CONF.ALTA){
      conf=RECON_CONF.MEDIA;
      confMotivo="GROOT aparece com mais de uma matrícula (possível troca de contrato)";
    }

    items.push(makeItem({status, alerts, conf, confMotivo, esperado:exp, achado, achados,
      diagnostico, diffDias, diffFte, compN, compNext, emp:exp.emp, dim, idx,
      cobranca:originalBilling(exp.emp, compN),
      faltantes:faltantes(exp,achado), excedentes:excedentes(exp,achado),
      modeloRow:modeloPorPessoa.get(key)||null,
      identidadeOrigem:exp.emp,
      linhasNext:linhasNextPorPessoa.get(key)||[]}));
  });

  // --- 5.2 retroativos da N+1 que nenhum fato da N explica ---
  encontrados.filter(c=>!c.consumido).forEach(c=>{
    const key=idx.keyOf(c.emp);
    const corrobora=movimentacoes?buscarMovimentacao(movimentacoes,c,idx):null;
    const naOrigem=origem.employees.find(e=>idx.keyOf(e)===key)||null;
    items.push(makeItem({
      status:RECON_STATUS.SEM_ORIGEM, alerts:[RECON_STATUS.SEM_ORIGEM],
      conf:corrobora?RECON_CONF.MEDIA:RECON_CONF.REVISAO,
      confMotivo:corrobora?"movimentação externa compatível localizada"
        :"nenhuma evidência nas duas faturas justifica o lançamento",
      esperado:null, achado:c, achados:[c],
      diagnostico:"Retroativo sem origem identificada. A fatura "+compNext.label+" traz "
        +(c.kind==="DESCONTAR"?"um desconto":"um acréscimo")+" de "+fmtShort(c.start)+" a "
        +fmtShort(c.end)+" referente a "+compN.label+", mas não foi possível justificar este "
        +"lançamento apenas com as duas faturas carregadas. "
        +(corrobora?"Uma movimentação externa compatível foi localizada."
          :"Pode existir movimentação externa não presente nos arquivos. Requer validação."),
      diffDias:0, diffFte:0, compN, compNext, emp:c.emp, idx,
      dim:{identidade:!idx.isAmbiguous(c.emp), competencia:true, periodo:null, sinal:null, fte:null},
      cobranca:naOrigem?originalBilling(naOrigem, compN):null,
      identidadeOrigem:naOrigem||c.emp,
      modeloRow:modeloPorPessoa.get(key)||null,
      linhasNext:linhasNextPorPessoa.get(key)||[]}));
  });

  /* --- 5.3 linhas da N+1 que não deu para classificar ---
     Só vale apontar quando a linha TEM um período: aí ela poderia ser um
     retroativo e a dúvida é real. Sem período válido ela não pode ser
     retroativo de nada — não é "indeterminada", é inaplicável, e vira ruído
     numa fatura de centenas de linhas. */
  const semPeriodoNaSeguinte=[];
  splitNext.indeterminados.forEach(e=>{
    if(!isValidYmd(e.inicio)&&!isValidYmd(e.fim)){ semPeriodoNaSeguinte.push(e); return; }
    items.push(makeItem({
      status:RECON_STATUS.INDETERMINADO, alerts:[RECON_STATUS.INDETERMINADO],
      conf:RECON_CONF.REVISAO,
      confMotivo:"a linha tem período, mas a leitura dele comporta mais de uma interpretação",
      esperado:null, achado:null, achados:[],
      diagnostico:"Esta linha da fatura "+compNext.label+" aparece aqui porque tem um período "
        +"que poderia ser um retroativo de "+compN.label+", mas os sinais se contradizem. "
        +(e.lineClassReason||"")
        +" Enquanto isso não for esclarecido, ela não entra em nenhuma comparação: nem conta "
        +"como ajuste encontrado, nem como ajuste ausente.",
      diffDias:0, diffFte:0, compN, compNext, emp:e, idx,
      dim:{identidade:!idx.isAmbiguous(e), competencia:false, periodo:null, sinal:null, fte:null},
      cobranca:null, identidadeOrigem:e, modeloRow:e.srcRow, linhasNext:[e]}));
  });

  /* --- 5.3b pessoas sem nenhum identificador, mas com ajuste devido ---
     O ajuste é mostrado: ele é real e calculável. O que não é possível é
     conferi-lo automaticamente na fatura seguinte. */
  semIdComAjuste.forEach(r=>{
    items.push(makeItem({
      status:RECON_STATUS.SEM_ID, alerts:[RECON_STATUS.SEM_ID],
      conf:RECON_CONF.REVISAO,
      confMotivo:"a linha não tem GROOT nem matrícula — não há chave para procurar na fatura seguinte",
      esperado:r, achado:null, achados:[],
      diagnostico:"Esta pessoa aparece aqui porque a linha dela em "+compN.label+" não tem GROOT ID "
        +"nem matrícula. As datas são suficientes para calcular o ajuste devido, e ele está aí ao "
        +"lado — o que não dá para fazer é procurá-lo em "+compNext.label+": sem identificador a "
        +"única chave restante seria o nome, e nome não é usado como chave porque homônimo existe. "
        +"Confira este ajuste na mão, ou complete o cadastro na origem e reconcilie de novo.",
      diffDias:0, diffFte:0, compN, compNext, emp:r.emp, idx,
      dim:{identidade:false, competencia:true, periodo:null, sinal:null, fte:null},
      cobranca:originalBilling(r.emp, compN), identidadeOrigem:r.emp,
      modeloRow:null, linhasNext:[]}));
  });

  /* --- 5.3c pessoas da fatura N cujas datas não foram lidas ---
     Ficavam de fora sem aviso: classifyLine as marcava indeterminadas e só as
     linhas normais chegavam ao motor de projeção. Sem data legível não há como
     dizer se havia ajuste devido, e é justamente isso que precisa ser dito. */
  splitN.indeterminados.forEach(e=>{
    items.push(makeItem({
      status:RECON_STATUS.REVISAO, alerts:[RECON_STATUS.REVISAO],
      conf:RECON_CONF.REVISAO, confMotivo:"as datas da linha de origem não foram lidas",
      esperado:null, achado:null, achados:[],
      diagnostico:"Esta pessoa aparece aqui porque a linha dela em "+compN.label+" não teve as "
        +"datas lidas. "+(e.lineClassReason||"")+" São as datas que determinam o ajuste devido, "
        +"então sem elas não é possível dizer se algo deveria ter surgido em "+compNext.label+". "
        +"Corrija a linha na origem e reconcilie de novo.",
      diffDias:0, diffFte:0, compN, compNext, emp:e, idx,
      dim:{identidade:!idx.isAmbiguous(e), competencia:true, periodo:null, sinal:null, fte:null},
      cobranca:null, identidadeOrigem:e, modeloRow:null, linhasNext:[]}));
  });

  /* --- 5.4 linhas da fatura N que o motor não conseguiu ler ---
     Aparecem só quando dá para dizer DE QUEM é a linha. Um erro de dados numa
     linha anônima não é acionável: não há quem procurar na fatura seguinte. */
  proj.errors.filter(err=>String(err.nome||"").trim()&&err.nome!=="(sem nome)").forEach(err=>{
    items.push(makeItem({
      status:RECON_STATUS.REVISAO, alerts:[RECON_STATUS.REVISAO],
      conf:RECON_CONF.REVISAO, confMotivo:"a fatura de origem não traz o dado necessário",
      esperado:null, achado:null, achados:[],
      diagnostico:"Esta pessoa aparece aqui porque a linha dela na fatura "+compN.label
        +" tem um problema nas datas ou no rateio — "+err.reason+" — e são justamente esses "
        +"campos que determinam o ajuste devido. Sem eles não é possível dizer se algo deveria "
        +"ter surgido em "+compNext.label+", então o app prefere avisar a chutar. Corrija a "
        +"linha na origem e reconcilie de novo.",
      diffDias:0, diffFte:0, compN, compNext, idx,
      dim:{identidade:false, competencia:true, periodo:null, sinal:null, fte:null},
      cobranca:null, identidadeOrigem:null, modeloRow:null, linhasNext:[],
      emp:{nome:err.nome, groot:null, matricula:null}}));
  });

  ordenar(items);
  items.forEach((it,i)=>{ it.id=i; });
  return {items, summary:summarize(items), contexto:{compN, compNext,
    linhasOrigem:origem.employees.length, linhasSeguinte:seguinte.employees.length,
    correntesNaSeguinte:splitNext.normais.length,
    retroativosNaSeguinte:encontrados.length,
    indeterminadosNaSeguinte:splitNext.indeterminados.length,
    ignoradas:splitN.ignoradas.length+splitNext.ignoradas.length+semPeriodoNaSeguinte.length,
    esperados:esperados.length,
    sequencia:checkSequence(compN,compNext)}};
}

function descreverTrecho(t){
  return t.days===1?fmtShort(t.start):fmtShort(t.start)+" a "+fmtShort(t.end)
    +" ("+t.days+" dias)";
}

function buscarMovimentacao(movs, achado, idx){
  if(!Array.isArray(movs)) return null;
  const k=idx.keyOf(achado.emp);
  return movs.find(m=>idx.keyOf(m)===k &&
    ((isValidYmd(m.fim)&&Math.abs(diffDaysInclusive(m.fim,achado.start))<=2) ||
     (isValidYmd(m.inicio)&&m.inicio===achado.start))) || null;
}

/** A história de faturamento da pessoa, que a interface narra de ponta a ponta. */
function makeItem(o){
  const e=o.emp||{};
  const ident=o.identidadeOrigem||null;
  return {
    status:o.status, alerts:o.alerts&&o.alerts.length?o.alerts:[o.status],
    confianca:o.conf, confiancaMotivo:o.confMotivo||"",
    nome:e.nome||"(sem nome)", groot:e.groot||null, matricula:e.matricula||null,
    identidade:{fonte:o.idx?o.idx.fonteOf(e):"NENHUMA",
      ambigua:o.idx?o.idx.isAmbiguous(e):false},
    lineClassification:e.lineClassification||null,
    lineClassReason:e.lineClassReason||null,
    // história de faturamento
    cobrancaOriginal:o.cobranca||null,
    movimentacaoInferida:inferMovement(o.achado||o.esperado, o.compN),
    esperado:o.esperado?{kind:o.esperado.kind, start:o.esperado.start, end:o.esperado.end,
      days:o.esperado.days, fte:o.esperado.fte, rateio:o.esperado.rateio,
      motivo:o.esperado.motivo, sentence:o.esperado.sentence, mov:o.esperado.mov,
      emp:o.esperado.emp}:null,
    achado:o.achado?{kind:o.achado.kind, start:o.achado.start, end:o.achado.end,
      days:o.achado.days, fte:o.achado.fte, rateio:o.achado.rateio,
      classe:o.achado.classe, classeMotivo:o.achado.classeMotivo,
      srcRow:o.achado.emp?o.achado.emp.srcRow:null, emp:o.achado.emp}:null,
    achados:(o.achados||[]).map(a=>({start:a.start, end:a.end, days:a.days, fte:a.fte,
      kind:a.kind, srcRow:a.emp?a.emp.srcRow:null})),
    dimensoes:o.dim||{identidade:null,competencia:null,periodo:null,sinal:null,fte:null},
    faltantes:o.faltantes||[], excedentes:o.excedentes||[],
    diffDias:o.diffDias||0, diffFte:o.diffFte||0,
    diagnostico:o.diagnostico,
    compOrigem:o.compN.label, compAplicacao:o.compNext.label,
    compDias:o.compN.days, compFirst:o.compN.first, compLast:o.compN.last,
    compSnapshot:o.compN.snapshot,
    modeloRow:o.modeloRow||null,          // TEMPLATE de estrutura, dentro da N+1
    identidadeRaw:ident&&ident.raw?ident.raw:null,   // IDENTIDADE, preferencialmente da N
    identidadeCampos:ident?{groot:ident.groot, nome:ident.nome, matricula:ident.matricula}:null,
    impacto:impactoFinanceiro(o),
    // decisão do usuário — nasce sempre pendente, nunca em "aceitar"
    decisao:"MANTER", sugestao:null, observacao:""
  };
}

/* Impacto financeiro só quando a tarifa é inequívoca. Sem ela, o app diz
   que não calculou — nunca inventa valor. */
function impactoFinanceiro(o){
  const raw=o.identidadeOrigem&&o.identidadeOrigem.raw;
  const tarifa=raw&&typeof raw.valor==="number"&&isFinite(raw.valor)&&raw.valor>0?raw.valor:null;
  if(!tarifa){
    return {calculado:false,
      motivo:"Impacto financeiro não calculado: a planilha não traz uma tarifa que permita "
        +"converter FTE em valor com segurança."};
  }
  const orig=o.cobranca&&o.cobranca.cobrado?o.cobranca.fte*tarifa:null;
  const esp=o.esperado?o.esperado.fte*tarifa:null;
  const ach=o.achado?o.achado.fte*tarifa:null;
  return {calculado:true, tarifa, original:orig, esperado:esp, encontrado:ach,
    diferenca:(esp!==null&&ach!==null)?ach-esp:null};
}

const ORDEM_STATUS=[RECON_STATUS.AUSENTE, RECON_STATUS.SINAL, RECON_STATUS.DUPLICADO,
  RECON_STATUS.PERIODO, RECON_STATUS.FTE, RECON_STATUS.SEM_ORIGEM,
  RECON_STATUS.AMBIGUA, RECON_STATUS.SEM_ID, RECON_STATUS.INDETERMINADO, RECON_STATUS.REVISAO,
  RECON_STATUS.CONCILIADO];
function ordenar(items){
  items.sort((a,b)=>{
    const d=ORDEM_STATUS.indexOf(a.status)-ORDEM_STATUS.indexOf(b.status);
    if(d!==0) return d;
    return String(a.nome).localeCompare(String(b.nome),"pt-BR");
  });
}

function summarize(items){
  const s={total:items.length, conciliados:0, pendencias:0,
    descontosAusentes:0, acrescimosAusentes:0, divergencias:0,
    duplicados:0, semOrigem:0, revisao:0, indeterminados:0, semIdentificador:0};
  items.forEach(it=>{
    if(it.status===RECON_STATUS.CONCILIADO){ s.conciliados++; return; }
    s.pendencias++;
    if(it.status===RECON_STATUS.AUSENTE){
      if(it.esperado&&it.esperado.kind==="DESCONTAR") s.descontosAusentes++;
      else s.acrescimosAusentes++;
    }
    if(it.alerts.includes(RECON_STATUS.PERIODO)||it.alerts.includes(RECON_STATUS.FTE)||
       it.alerts.includes(RECON_STATUS.SINAL)) s.divergencias++;
    if(it.status===RECON_STATUS.DUPLICADO) s.duplicados++;
    if(it.status===RECON_STATUS.SEM_ORIGEM) s.semOrigem++;
    if(it.status===RECON_STATUS.INDETERMINADO) s.indeterminados++;
    if(it.status===RECON_STATUS.SEM_ID) s.semIdentificador++;
    if(it.status===RECON_STATUS.REVISAO||it.status===RECON_STATUS.AMBIGUA) s.revisao++;
  });
  return s;
}

/* ================================================================
   6. SUGESTÃO DE CORREÇÃO

   Calculada só quando o usuário escolhe "aceitar", e ainda assim é
   proposta: pode ser editada antes de entrar na prévia.
   ================================================================ */
function sugerirCorrecao(item, modo){
  const exp=item.esperado, ach=item.achado;
  if(item.status===RECON_STATUS.AUSENTE && exp){
    return {acao:"INCLUIR", kind:exp.kind, start:exp.start, end:exp.end,
      rateio:exp.rateio, days:exp.days, fte:exp.fte,
      motivo:exp.motivo||"Ajuste retroativo não localizado na fatura seguinte."};
  }
  if(item.alerts.includes(RECON_STATUS.PERIODO) && exp && ach){
    if(modo==="COMPLEMENTAR"){
      const trecho=(item.faltantes&&item.faltantes[0])||null;
      if(!trecho) return null;
      return {acao:"INCLUIR", kind:exp.kind, start:trecho.start, end:trecho.end,
        rateio:exp.rateio, ...medir(trecho.start, trecho.end, exp.rateio, exp.kind, item.compDias),
        motivo:"Complemento do trecho não contemplado no lançamento existente."};
    }
    return {acao:"SUBSTITUIR", alvoRow:ach.srcRow, kind:exp.kind, start:exp.start, end:exp.end,
      rateio:exp.rateio, days:exp.days, fte:exp.fte,
      motivo:"Período do lançamento existente ajustado para o esperado pela regra do corte."};
  }
  if((item.status===RECON_STATUS.FTE||item.status===RECON_STATUS.SINAL) && exp && ach){
    return {acao:"SUBSTITUIR", alvoRow:ach.srcRow, kind:exp.kind, start:exp.start, end:exp.end,
      rateio:exp.rateio, days:exp.days, fte:exp.fte,
      motivo:item.status===RECON_STATUS.SINAL
        ?"Sentido do lançamento corrigido para o esperado."
        :"Rateio/FTE ajustado para o esperado pela regra do corte."};
  }
  if(item.status===RECON_STATUS.DUPLICADO && item.achados.length>1){
    return {acao:"REMOVER", alvoRows:item.achados.slice(1).map(a=>a.srcRow),
      manterRow:item.achados[0].srcRow,
      motivo:"Lançamento repetido para o mesmo período."};
  }
  return null;
}

function medir(start, end, rateio, kind, compDias){
  const days=diffDaysInclusive(start,end);
  const sinal=kind==="DESCONTAR"?-1:1;
  return {days, fte:sinal*(days/compDias)*rateio};
}

function fmtFte(v){ return (v>0?"+":"")+v.toFixed(4).replace(".",","); }
function fmtPct(v){ return (v*100).toLocaleString("pt-BR",{maximumFractionDigits:1})+"%"; }

if(typeof module!=="undefined"&&module.exports){
  module.exports={RECON_STATUS, RECON_CONF, RECON_META, normId, buildPersonIndex,
    isRetroLine, splitInvoiceLines, describeRetro, checkSequence,
    competencesAreSequential, reconcile, sugerirCorrecao, faltantes, excedentes};
}
