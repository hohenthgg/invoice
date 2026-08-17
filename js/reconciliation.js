/* Ajustes MELI — motor de conciliação entre duas faturas subsequentes
   ================================================================

   O modo "Projetar ajustes" (js/engine.js) responde: dada UMA fatura,
   quais ajustes deveriam aparecer na competência seguinte? Este módulo
   responde a pergunta seguinte, com DUAS faturas na mão:

     o ajuste que deveria ter surgido na fatura N+1 realmente surgiu?

   O motor não decide nada. Ele detecta, compara, classifica e explica —
   a decisão de alterar uma fatura pertence sempre ao usuário, e vive na
   camada de interface (js/reconciliation-ui.js).

   Toda a aritmética de datas e o cálculo do ajuste esperado são
   reaproveitados de dates.js e engine.js: aqui não há uma segunda
   implementação da regra do dia 15.
   ================================================================ */
"use strict";

const RECON_STATUS = {
  CONCILIADO: "CONCILIADO",
  AUSENTE:    "AUSENTE",
  PARCIAL:    "PARCIAL",
  RATEIO:     "RATEIO_DIVERGENTE",
  SINAL:      "SINAL_INCORRETO",
  DUPLICADO:  "DUPLICADO",
  SEM_ORIGEM: "SEM_ORIGEM",
  REVISAO:    "REVISAO"
};
const RECON_CONF = { ALTA:"ALTA", MEDIA:"MEDIA", REVISAO:"REVISAO" };

/* Rótulos e semântica de cor. Vermelho só para ausência ou conflito
   relevante; âmbar para o que é parcial ou precisa de olho humano. */
const RECON_META = {
  CONCILIADO: {label:"Conciliado",            tone:"ok",   icon:"✓"},
  AUSENTE:    {label:"Ajuste ausente",        tone:"bad",  icon:"✖"},
  PARCIAL:    {label:"Ajuste parcial",        tone:"warn", icon:"⚠"},
  RATEIO_DIVERGENTE:{label:"Rateio/FTE divergente", tone:"warn", icon:"⚠"},
  SINAL_INCORRETO:{label:"Sinal incorreto",   tone:"bad",  icon:"✖"},
  DUPLICADO:  {label:"Possível duplicidade",  tone:"bad",  icon:"✖"},
  SEM_ORIGEM: {label:"Retroativo sem origem", tone:"info", icon:"?"},
  REVISAO:    {label:"Revisão manual",        tone:"info", icon:"?"}
};

const FTE_TOL = 1e-6;

/* ================================================================
   1. IDENTIDADE DAS PESSOAS

   O Excel devolve o mesmo identificador ora como texto, ora como
   número: "123456", 123456 e 123456.0 são a mesma pessoa. Já nomes
   NUNCA são chave — homônimos existem e grafia varia entre faturas.
   ================================================================ */
function normId(v){
  if(v===null||v===undefined) return "";
  const s=String(v).trim();
  if(!s) return "";
  if(/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(s)){
    const n=Number(s);
    // 123456.0 → "123456"; preserva não-inteiros como estão
    if(isFinite(n)) return Number.isInteger(n) ? String(n) : String(n);
  }
  return s.toUpperCase();
}
function hasId(v){ const s=normId(v); return s!=="" && s!=="0"; }

/** Índice de pessoas das duas faturas.
 *  GROOT é a chave preferida; matrícula só entra como ponte quando
 *  aponta para um único GROOT. Matrícula que aponta para dois GROOTs
 *  diferentes é conflito: as linhas ficam marcadas como ambíguas e
 *  NUNCA são unidas automaticamente. */
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
    /** chave canônica da pessoa, ou "" quando a linha não tem identificador */
    keyOf(e){
      if(hasId(e.groot)) return "G:"+normId(e.groot);
      if(hasId(e.matricula)){
        const m=normId(e.matricula), gids=midToGids.get(m);
        if(gids && gids.size===1) return "G:"+[...gids][0];   // ponte segura
        return "M:"+m;
      }
      return "";
    },
    /** a identificação desta linha comporta mais de uma leitura? */
    isAmbiguous(e){
      if(!hasId(e.groot) && hasId(e.matricula)) return ambiguous.has("M:"+normId(e.matricula));
      if(hasId(e.groot) && hasId(e.matricula)){
        const gids=midToGids.get(normId(e.matricula));
        return !!(gids && gids.size>1);
      }
      return false;
    },
    /** GROOT que aparece com duas matrículas distintas: sinal de troca de contrato */
    hasSplitMatricula(e){
      if(!hasId(e.groot)) return false;
      const mids=gidToMids.get(normId(e.groot));
      return !!(mids && mids.size>1);
    }
  };
}

/* ================================================================
   2. LINHA NORMAL × LINHA RETROATIVA

   Esta separação é o coração do modo. Uma linha de ajuste retroativo
   pode legitimamente ter % RATEIO negativo — o motor de projeção trata
   rateio negativo como erro de dados, e está certo para uma linha de
   competência, mas errado para um retroativo.

   O discriminador NÃO é o sinal. É o período: uma linha cujo período
   TERMINA antes do primeiro dia da competência do arquivo não pode
   pertencer àquela competência — só pode ser retroativo. Isso importa
   porque uma linha perfeitamente normal de junho costuma carregar a
   DATA DE INÍCIO real da pessoa (a admissão em maio); o que a denuncia
   como normal é a DATA FIM vazia ou dentro de junho.

   O sinal do rateio entra como evidência adicional, nunca sozinho.
   ================================================================ */
function isRetroLine(e, comp){
  const temFim=isValidYmd(e.fim);
  // período fechado antes do início da competência do arquivo
  if(temFim && e.fim < comp.first) return true;
  // crédito/estorno: rateio negativo não pertence a uma linha de competência
  if(typeof e.rateio==="number" && e.rateio < 0) return true;
  return false;
}

/** Separa as linhas de uma fatura entre as da própria competência e as retroativas. */
function splitInvoiceLines(employees, comp){
  const normais=[], retroativos=[];
  employees.forEach(e=>{ (isRetroLine(e,comp)?retroativos:normais).push(e); });
  return {normais, retroativos};
}

/** Descreve uma linha retroativa encontrada, já normalizada para comparação.
 *  `compAlvo` é a competência a que o período se refere (a fatura N). */
function describeRetro(e, compAlvo){
  if(!isValidYmd(e.inicio)||!isValidYmd(e.fim)) return null;
  const start=e.inicio, end=e.fim;
  if(end<start) return null;
  const days=diffDaysInclusive(start,end);
  const rateio=(typeof e.rateio==="number"&&isFinite(e.rateio))?e.rateio:1;
  // O sinal do rateio diz o sentido do lançamento; o período diz a que mês pertence.
  const kind=rateio<0?"DESCONTAR":"ACRESCENTAR";
  return {emp:e, kind, start, end, days,
    rateio:Math.abs(rateio),
    fte:(days/compAlvo.days)*rateio,
    dentroDaCompetencia: start>=compAlvo.first && end<=compAlvo.last};
}

/* ================================================================
   3. COMPETÊNCIAS SUBSEQUENTES
   ================================================================ */
function monthIndex(comp){ return comp.y*12+comp.m; }
function competencesAreSequential(compA, compB){
  return monthIndex(compB)-monthIndex(compA)===1;
}
/** Diagnóstico da dupla de competências, para a interface decidir o que dizer. */
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
   4. MOTOR

   reconcile({origem, seguinte, movimentacoes?})

     origem    = {employees, comp}   fatura N
     seguinte  = {employees, comp}   fatura N+1
     movimentacoes (opcional, ainda não obrigatório) = fonte externa de
       RH/HCM que, quando existir, corrobora retroativos sem origem.

   Devolve {items, summary, contexto}. Cada item é um apontamento com
   esperado, encontrado, diferença, confiança e uma explicação em texto.
   Nenhum item vem com decisão tomada: `decisao` nasce "MANTER".
   ================================================================ */
function reconcile(input){
  const origem=input.origem, seguinte=input.seguinte;
  const compN=origem.comp, compNext=seguinte.comp;
  const movimentacoes=input.movimentacoes||null;

  const idx=buildPersonIndex([origem.employees, seguinte.employees]);

  // --- 4.1 O que a fatura N obriga a ajustar na seguinte ---
  const splitN=splitInvoiceLines(origem.employees, compN);
  const proj=analyze(splitN.normais, compN);       // reaproveita o motor de projeção
  const esperados=proj.adjustments;

  // --- 4.2 O que a fatura N+1 traz de retroativo referente a N ---
  const splitNext=splitInvoiceLines(seguinte.employees, compNext);
  const encontrados=[];
  splitNext.retroativos.forEach(e=>{
    const d=describeRetro(e, compN);
    if(!d) return;
    // só interessam os retroativos que apontam para a competência N
    if(d.end<compN.first || d.start>compN.last) return;
    d.consumido=false;
    encontrados.push(d);
  });

  const porPessoa=new Map();
  encontrados.forEach(d=>{
    const k=idx.keyOf(d.emp);
    if(!porPessoa.has(k)) porPessoa.set(k,[]);
    porPessoa.get(k).push(d);
  });

  /* Linha da MESMA pessoa dentro da fatura N+1, para uma eventual inclusão
     herdar estilo e campos cadastrais. Índice de linha só faz sentido dentro
     do próprio arquivo: usar a linha da fatura N aqui produziria uma linha com
     os dados de outra pessoa. Linha normal tem preferência sobre retroativa. */
  const modeloPorPessoa=new Map();
  splitNext.normais.forEach(e=>{ const k=idx.keyOf(e);
    if(k&&!modeloPorPessoa.has(k)) modeloPorPessoa.set(k,e.srcRow); });
  splitNext.retroativos.forEach(e=>{ const k=idx.keyOf(e);
    if(k&&!modeloPorPessoa.has(k)) modeloPorPessoa.set(k,e.srcRow); });

  const items=[];
  const overlap=(a,b)=>a.start<=b.end && b.start<=a.end;

  // --- 4.3 Confronto: para cada ajuste esperado, procurar o aplicado ---
  esperados.forEach(exp=>{
    const key=idx.keyOf(exp.emp);
    const cands=(porPessoa.get(key)||[]).filter(c=>!c.consumido);
    const ambiguo=idx.isAmbiguous(exp.emp)||!key;

    const exatos=cands.filter(c=>c.start===exp.start&&c.end===exp.end);
    const sobrepostos=cands.filter(c=>overlap(c,exp));

    let status, conf, achado=null, achados=[], diagnostico, diffDias=0, diffFte=0;

    if(exatos.length>1){
      status=RECON_STATUS.DUPLICADO; conf=RECON_CONF.ALTA;
      achados=exatos; achado=exatos[0];
      exatos.forEach(c=>c.consumido=true);
      diagnostico="A fatura seguinte traz "+exatos.length+" lançamentos para o mesmo período ("
        +fmtShort(exp.start)+" a "+fmtShort(exp.end)+"). Possível "
        +(exp.kind==="DESCONTAR"?"desconto":"acréscimo")+" em duplicidade — confirme qual linha deve permanecer.";
    } else if(exatos.length===1){
      achado=exatos[0]; achado.consumido=true; achados=[achado];
      diffFte=achado.fte-exp.fte;
      if(achado.kind!==exp.kind){
        status=RECON_STATUS.SINAL; conf=RECON_CONF.ALTA;
        diagnostico="O período confere, mas o lançamento está no sentido oposto: esperado "
          +(exp.kind==="DESCONTAR"?"desconto":"acréscimo")+" e encontrado "
          +(achado.kind==="DESCONTAR"?"desconto":"acréscimo")+".";
      } else if(Math.abs(diffFte)>FTE_TOL){
        status=RECON_STATUS.RATEIO; conf=RECON_CONF.ALTA;
        diagnostico="Período correto, quantidade diferente: esperado FTE "+fmtFte(exp.fte)
          +" (rateio "+fmtPct(exp.rateio)+") e encontrado "+fmtFte(achado.fte)
          +" (rateio "+fmtPct(achado.rateio)+").";
      } else {
        status=RECON_STATUS.CONCILIADO; conf=RECON_CONF.ALTA;
        diagnostico="O ajuste esperado foi localizado na fatura seguinte, com o mesmo período e o mesmo FTE.";
      }
    } else if(sobrepostos.length>1){
      status=RECON_STATUS.REVISAO; conf=RECON_CONF.REVISAO;
      achados=sobrepostos; achado=sobrepostos[0];
      sobrepostos.forEach(c=>c.consumido=true);
      diagnostico="Há mais de um lançamento retroativo sobrepondo o período esperado. "
        +"A correspondência comporta mais de uma leitura — requer conferência manual.";
    } else if(sobrepostos.length===1){
      achado=sobrepostos[0]; achado.consumido=true; achados=[achado];
      diffDias=achado.days-exp.days;
      diffFte=achado.fte-exp.fte;
      if(achado.kind!==exp.kind){
        status=RECON_STATUS.SINAL; conf=RECON_CONF.ALTA;
        diagnostico="O lançamento encontrado cobre parte do período esperado, mas no sentido oposto.";
      } else {
        status=RECON_STATUS.PARCIAL; conf=RECON_CONF.MEDIA;
        diagnostico="A fatura seguinte contém o "+(exp.kind==="DESCONTAR"?"desconto":"acréscimo")
          +", mas o período não coincide: esperado "+fmtShort(exp.start)+" a "+fmtShort(exp.end)
          +" ("+exp.days+" dia"+(exp.days===1?"":"s")+") e encontrado "+fmtShort(achado.start)+" a "
          +fmtShort(achado.end)+" ("+achado.days+" dia"+(achado.days===1?"":"s")+"). Diferença de "
          +Math.abs(diffDias)+" dia"+(Math.abs(diffDias)===1?"":"s")+".";
      }
    } else {
      status=RECON_STATUS.AUSENTE;
      conf=ambiguo?RECON_CONF.REVISAO:RECON_CONF.ALTA;
      diffDias=-exp.days; diffFte=-exp.fte;
      diagnostico="Ajuste esperado pela regra do snapshot não localizado na fatura "+compNext.label+". "
        +"Nenhuma linha retroativa desta pessoa referente a "+compN.label+" foi encontrada.";
    }

    if(ambiguo){
      conf=RECON_CONF.REVISAO;
      diagnostico+=" A identificação desta pessoa é ambígua entre as duas faturas — confirme GROOT e matrícula antes de agir.";
    } else if(idx.hasSplitMatricula(exp.emp) && conf===RECON_CONF.ALTA){
      conf=RECON_CONF.MEDIA;
    }

    items.push(makeItem({status, conf, esperado:exp, achado, achados, diagnostico,
      diffDias, diffFte, compN, compNext, emp:exp.emp,
      modeloRow:modeloPorPessoa.get(key)||null}));
  });

  // --- 4.4 Retroativos da N+1 que nenhum fato da N explica ---
  encontrados.filter(c=>!c.consumido).forEach(c=>{
    const corrobora=movimentacoes?buscarMovimentacao(movimentacoes,c,idx):null;
    items.push(makeItem({
      status:RECON_STATUS.SEM_ORIGEM,
      conf:corrobora?RECON_CONF.MEDIA:RECON_CONF.REVISAO,
      esperado:null, achado:c, achados:[c],
      diagnostico:"Retroativo encontrado sem origem identificável nas duas faturas: a fatura "
        +compNext.label+" traz "+(c.kind==="DESCONTAR"?"um desconto":"um acréscimo")+" de "
        +fmtShort(c.start)+" a "+fmtShort(c.end)+" referente a "+compN.label
        +", mas nada na fatura de origem explica esse lançamento. "
        +(corrobora?"Uma movimentação externa compatível foi localizada.":
          "Informação insuficiente — requer validação da movimentação."),
      diffDias:0, diffFte:0, compN, compNext, emp:c.emp,
      modeloRow:modeloPorPessoa.get(idx.keyOf(c.emp))||null}));
  });

  // --- 4.5 Linhas da fatura N que o motor não conseguiu ler ---
  proj.errors.forEach(err=>{
    items.push(makeItem({
      status:RECON_STATUS.REVISAO, conf:RECON_CONF.REVISAO,
      esperado:null, achado:null, achados:[],
      diagnostico:"Informação insuficiente na fatura de origem: "+err.reason
        +". Sem isso não é possível afirmar qual ajuste seria devido.",
      diffDias:0, diffFte:0, compN, compNext,
      emp:{nome:err.nome, groot:null, matricula:null}}));
  });

  ordenar(items);
  items.forEach((it,i)=>{ it.id=i; });
  return {items, summary:summarize(items), contexto:{compN, compNext,
    linhasOrigem:origem.employees.length, linhasSeguinte:seguinte.employees.length,
    retroativosNaSeguinte:encontrados.length, esperados:esperados.length,
    sequencia:checkSequence(compN,compNext)}};
}

/* Gancho para a terceira fonte (RH/HCM), ainda opcional: quando uma base
   de movimentações for fornecida, um retroativo sem origem nas faturas
   pode ser corroborado por ela em vez de cair direto em revisão. */
function buscarMovimentacao(movs, achado, idx){
  if(!Array.isArray(movs)) return null;
  const k=idx.keyOf(achado.emp);
  return movs.find(m=>idx.keyOf(m)===k &&
    ((isValidYmd(m.fim)&&Math.abs(diffDaysInclusive(m.fim,achado.start))<=2) ||
     (isValidYmd(m.inicio)&&m.inicio===achado.start))) || null;
}

function makeItem(o){
  const e=o.emp||{};
  return {
    status:o.status, confianca:o.conf,
    nome:e.nome||"(sem nome)", groot:e.groot||null, matricula:e.matricula||null,
    esperado:o.esperado?{kind:o.esperado.kind, start:o.esperado.start, end:o.esperado.end,
      days:o.esperado.days, fte:o.esperado.fte, rateio:o.esperado.rateio,
      motivo:o.esperado.motivo, sentence:o.esperado.sentence, mov:o.esperado.mov,
      emp:o.esperado.emp}:null,
    achado:o.achado?{kind:o.achado.kind, start:o.achado.start, end:o.achado.end,
      days:o.achado.days, fte:o.achado.fte, rateio:o.achado.rateio,
      srcRow:o.achado.emp?o.achado.emp.srcRow:null, emp:o.achado.emp}:null,
    achados:(o.achados||[]).map(a=>({start:a.start, end:a.end, days:a.days, fte:a.fte,
      kind:a.kind, srcRow:a.emp?a.emp.srcRow:null})),
    diffDias:o.diffDias||0, diffFte:o.diffFte||0,
    diagnostico:o.diagnostico,
    compOrigem:o.compN.label, compAplicacao:o.compNext.label,
    compDias:o.compN.days,          // divisor do FTE, para recalcular sugestões
    modeloRow:o.modeloRow||null,    // linha desta pessoa DENTRO da fatura N+1
    compFirst:o.compN.first, compLast:o.compN.last,
    // decisão do usuário — nasce sempre pendente, nunca em "aceitar"
    decisao:"MANTER", sugestao:null, observacao:""
  };
}

/* Pendências primeiro; dentro do grupo, por nome. */
const ORDEM_STATUS=[RECON_STATUS.AUSENTE, RECON_STATUS.SINAL, RECON_STATUS.DUPLICADO,
  RECON_STATUS.PARCIAL, RECON_STATUS.RATEIO, RECON_STATUS.SEM_ORIGEM,
  RECON_STATUS.REVISAO, RECON_STATUS.CONCILIADO];
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
    duplicados:0, semOrigem:0, revisao:0};
  items.forEach(it=>{
    if(it.status===RECON_STATUS.CONCILIADO){ s.conciliados++; return; }
    s.pendencias++;
    if(it.status===RECON_STATUS.AUSENTE){
      if(it.esperado&&it.esperado.kind==="DESCONTAR") s.descontosAusentes++;
      else s.acrescimosAusentes++;
    }
    if(it.status===RECON_STATUS.PARCIAL||it.status===RECON_STATUS.RATEIO||
       it.status===RECON_STATUS.SINAL) s.divergencias++;
    if(it.status===RECON_STATUS.DUPLICADO) s.duplicados++;
    if(it.status===RECON_STATUS.SEM_ORIGEM) s.semOrigem++;
    if(it.status===RECON_STATUS.REVISAO) s.revisao++;
  });
  return s;
}

/* ================================================================
   5. SUGESTÃO DE CORREÇÃO

   Gerada sob demanda, quando o usuário escolhe "Aceitar sugestão".
   Nunca aplicada sozinha: descreve o que seria feito e permite edição
   antes da prévia. Para o parcial, a decisão entre substituir a linha e
   complementar a diferença é do usuário — as duas são oferecidas.
   ================================================================ */
function sugerirCorrecao(item, modo){
  const exp=item.esperado, ach=item.achado;
  if(item.status===RECON_STATUS.AUSENTE && exp){
    return {acao:"INCLUIR", kind:exp.kind, start:exp.start, end:exp.end,
      rateio:exp.rateio, days:exp.days, fte:exp.fte,
      motivo:exp.motivo||"Ajuste retroativo não localizado na fatura seguinte."};
  }
  if(item.status===RECON_STATUS.PARCIAL && exp && ach){
    if(modo==="COMPLEMENTAR"){
      // complementa só o trecho que faltou, preservando a linha existente
      const faltaInicio = ach.start>exp.start ? {start:exp.start, end:addDays(ach.start,-1)} : null;
      const faltaFim    = ach.end<exp.end     ? {start:addDays(ach.end,1), end:exp.end}     : null;
      const trecho=faltaInicio||faltaFim;
      if(!trecho) return null;
      return {acao:"INCLUIR", kind:exp.kind, start:trecho.start, end:trecho.end,
        rateio:exp.rateio, ...medir(trecho.start, trecho.end, exp.rateio, exp.kind, item.compDias),
        motivo:"Complemento do trecho não contemplado no lançamento existente."};
    }
    return {acao:"SUBSTITUIR", alvoRow:ach.srcRow, kind:exp.kind, start:exp.start, end:exp.end,
      rateio:exp.rateio, days:exp.days, fte:exp.fte,
      motivo:"Período do lançamento existente ajustado para o esperado pela regra do snapshot."};
  }
  if((item.status===RECON_STATUS.RATEIO||item.status===RECON_STATUS.SINAL) && exp && ach){
    return {acao:"SUBSTITUIR", alvoRow:ach.srcRow, kind:exp.kind, start:exp.start, end:exp.end,
      rateio:exp.rateio, days:exp.days, fte:exp.fte,
      motivo:item.status===RECON_STATUS.SINAL
        ?"Sentido do lançamento corrigido para o esperado."
        :"Rateio/FTE ajustado para o esperado pela regra do snapshot."};
  }
  if(item.status===RECON_STATUS.DUPLICADO && item.achados.length>1){
    // oferece remover as repetições, mantendo a primeira — o usuário escolhe
    return {acao:"REMOVER", alvoRows:item.achados.slice(1).map(a=>a.srcRow),
      manterRow:item.achados[0].srcRow,
      motivo:"Lançamento repetido para o mesmo período."};
  }
  return null;
}
/** Dias e FTE de um período — a mesma conta do motor de projeção,
 *  usada quando o usuário edita a sugestão à mão. */
function medir(start, end, rateio, kind, compDias){
  const days=diffDaysInclusive(start,end);
  const sinal=kind==="DESCONTAR"?-1:1;
  return {days, fte:sinal*(days/compDias)*rateio};
}

function fmtFte(v){ return (v>0?"+":"")+v.toFixed(4).replace(".",","); }
function fmtPct(v){ return (v*100).toLocaleString("pt-BR",{maximumFractionDigits:1})+"%"; }

/* Exportado para os testes (Node) sem atrapalhar o navegador. */
if(typeof module!=="undefined"&&module.exports){
  module.exports={RECON_STATUS, RECON_CONF, normId, buildPersonIndex,
    isRetroLine, splitInvoiceLines, describeRetro, checkSequence,
    competencesAreSequential, reconcile, sugerirCorrecao};
}
