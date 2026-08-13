/* Ajustes MELI — detecção automática da competência */
"use strict";

/* ================================================================
   DETECÇÃO AUTOMÁTICA DA COMPETÊNCIA

   Ordem das evidências:
   1) Retorno MELI / "Data Trab.": se um mês concentra >= 70% das datas
      válidas, ele É a competência (evidência temporal mais forte).
   2) Caso contrário, analisar apenas o CLUSTER RECENTE de movimentações
      (DATA FIM e admissões), ignorando o histórico antigo — assim dezenas
      de desligamentos de meses passados não elegem uma competência velha.
   3) Persistindo ambiguidade, marcar como não confiante → "Alterar".
   ================================================================ */
const RETORNO_MAJORITY = 0.70;   // maioria clara em Data Trab.
const CLUSTER_MONTHS   = 2;      // meses anteriores ao mês-âncora que entram no cluster
const MIN_CLUSTER_ABS  = 3;      // movimentações mínimas para um mês ser "real"
const MIN_CLUSTER_REL  = 0.05;   // ...ou 5% do mês de pico

/** Mês mais recente com volume real de movimentações.
 *  Uma ou duas datas soltas (erro de digitação, data futura) não formam
 *  um mês significativo e por isso não deslocam a âncora. */
function anchorMonth(employees){
  const cnt=new Map();
  const bump=v=>{ if(isValidYmd(v)){ const p=ymdParts(v); const k=p.y*12+p.m; cnt.set(k,(cnt.get(k)||0)+1); } };
  employees.forEach(e=>{ bump(e.fim); bump(e.inicio); });
  if(!cnt.size) return null;
  const peak=Math.max(...cnt.values());
  const min=Math.max(MIN_CLUSTER_ABS, MIN_CLUSTER_REL*peak);
  const keys=[...cnt.entries()].filter(([,n])=>n>=min).map(([k])=>k).sort((a,b)=>a-b);
  const k=keys.length?keys[keys.length-1]:[...cnt.keys()].sort((a,b)=>a-b).pop();
  return {y:Math.floor((k-1)/12), m:((k-1)%12)+1};
}

function detectCompetence(employees, retorno){
  // ---- 1) Retorno MELI com maioria clara decide sozinho ----
  if(retorno && retorno.total>0 && retorno.topShare>=RETORNO_MAJORITY){
    return {comp:buildCompetence(retorno.top.y, retorno.top.m), confident:true, via:"retorno"};
  }

  // ---- 2) Cluster recente de movimentações ----
  const anchor=anchorMonth(employees);
  if(!anchor){
    const now=new Date();
    return {comp:buildCompetence(now.getFullYear(), now.getMonth()+1), confident:false, via:"vazio"};
  }
  const aComp=buildCompetence(anchor.y,anchor.m);
  const startIdx=(anchor.y*12+anchor.m)-CLUSTER_MONTHS;
  const sy=Math.floor((startIdx-1)/12), sm=((startIdx-1)%12)+1;
  const from=ymd(sy,sm,1);        // início do cluster
  const to=aComp.last;            // fim do mês-âncora

  // candidatos: apenas meses dentro do cluster recente
  const cand=new Map();
  const addCand=v=>{ if(isValidYmd(v)&&v>=from&&v<=to){ const p=ymdParts(v); cand.set(p.y+"-"+p.m,{y:p.y,m:p.m}); } };
  employees.forEach(e=>{ addCand(e.fim); addCand(e.inicio); });
  if(retorno && retorno.top){
    const rk=retorno.top.y*12+retorno.top.m;
    if(rk>=startIdx && rk<=anchor.y*12+anchor.m) cand.set(retorno.top.y+"-"+retorno.top.m, retorno.top);
  }
  if(!cand.size) cand.set(anchor.y+"-"+anchor.m, anchor);

  const scored=[...cand.values()].map(c=>{
    const comp=buildCompetence(c.y,c.m);
    let fimsIn=0, inisIn=0, adj=0;
    employees.forEach(e=>{
      // só contam movimentações DENTRO do cluster recente
      if(isValidYmd(e.fim) && e.fim>=comp.first && e.fim<=comp.last && e.fim>=from && e.fim<=to) fimsIn++;
      if(isValidYmd(e.inicio) && e.inicio>=comp.first && e.inicio<=comp.last && e.inicio>=from && e.inicio<=to) inisIn++;
      const r=detectAdjustment(e,comp);
      if(r && r.kind!=="ERRO") adj++;
    });
    let score = 2.0*fimsIn + 0.5*inisIn + 1.0*adj;
    // Retorno MELI inconclusivo entra apenas como leve confirmação
    if(retorno && retorno.top && retorno.top.y===c.y && retorno.top.m===c.m) score *= 1.10;
    return {comp, score, fimsIn, adj};
  }).sort((a,b)=>b.score-a.score);

  const best=scored[0], second=scored[1];
  const confident = best.score>0 && (!second || second.score===0 || best.score>=1.35*second.score);
  return {comp:best.comp, confident, via:"cluster"};
}
