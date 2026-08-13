/* Ajustes MELI — motor de regras de ajuste */
"use strict";

/* ================================================================
   MOTOR DE REGRAS
   O snapshot do dia 15 congela a cobrança. O motor compara o que
   foi congelado com as movimentações posteriores ao congelamento.
   ================================================================ */
function validateEmployee(e){
  const errs=[];
  if(!e.groot) errs.push("GROOT ID vazio");
  if(!e.matricula) errs.push("matrícula vazia");
  if(e.inicio===null) errs.push("DATA DE INÍCIO vazia");
  if(Number.isNaN(e.inicio)) errs.push("DATA DE INÍCIO inválida");
  if(Number.isNaN(e.fim)) errs.push("DATA FIM inválida");
  if(isValidYmd(e.inicio)&&isValidYmd(e.fim)&&e.fim<e.inicio) errs.push("DATA FIM anterior à DATA DE INÍCIO");
  if(Number.isNaN(e.rateio)) errs.push("rateio inválido");
  if(typeof e.rateio==="number"&&(e.rateio<0||e.rateio>1)) errs.push("rateio fora de 0–100%");
  return errs;
}
function isActiveOnSnapshot(e,snapshot){
  // DATA FIM é inclusiva: último dia ativo
  return e.inicio<=snapshot && (e.fim===null||e.fim>=snapshot);
}

/** Retorna ajuste, null (sem ajuste) ou {kind:"ERRO"}. */
function detectAdjustment(e,comp){
  const errs=validateEmployee(e);
  if(errs.length) return {emp:e, kind:"ERRO", errors:errs};

  const warnings=[];
  let rateio=e.rateio;
  if(rateio===null){ rateio=1; warnings.push("Rateio vazio — assumido 100%"); }
  else if(rateio!==1) warnings.push("Rateio "+(rateio*100).toLocaleString("pt-BR",{maximumFractionDigits:1})+"%");

  if(isActiveOnSnapshot(e,comp.snapshot)){
    // Congelado: competência integral. Fato posterior: desligamento antes do fim do mês.
    if(e.fim!==null && e.fim<comp.last){
      const start=addDays(e.fim,1);                       // ausência começa no dia seguinte
      const end=comp.last;                                // nunca extrapola a competência
      const days=diffDaysInclusive(start,end);
      if(days<=0) return null;                            // guarda lógica
      return {emp:e, kind:"DESCONTAR", rateio, warnings,
        movDate:e.fim, mov:"Saiu "+fmtShort(e.fim),
        sentence:"Cobrado no snapshot de "+fmtShort(comp.snapshot)+"; desligamento em "+fmtShort(e.fim)
          +" → descontar "+fmtShort(start)+" a "+fmtShort(end)+".",
        motivo:"Desligamento pós-corte — descontar período posterior ao desligamento.",
        start, end, days, fte:-(days/comp.days)*rateio};
    }
    return null; // ativo no corte e até o fim do mês: cobrança integral correta
  }

  // Não congelado. Fato relevante: trabalho na janela pós-corte da própria competência.
  if(e.inicio>comp.snapshot && e.inicio<=comp.last){
    const start=Math.max(e.inicio,comp.first);            // limites da competência
    const end=(e.fim!==null && e.fim<comp.last)?e.fim:comp.last;
    const days=diffDaysInclusive(start,end);
    if(days<=0) return null;
    const inOut=(e.fim!==null && e.fim<comp.last);
    return {emp:e, kind:"ACRESCENTAR", rateio, warnings,
      movDate:e.inicio,
      mov: inOut?"Entrou "+fmtShort(e.inicio)+" • Saiu "+fmtShort(e.fim):"Entrou "+fmtShort(e.inicio),
      sentence:"Não existia no snapshot de "+fmtShort(comp.snapshot)+"; "
        +(inOut?"entrou "+fmtShort(e.inicio)+" e saiu "+fmtShort(e.fim):"admitido em "+fmtShort(e.inicio))
        +" → acrescentar "+fmtShort(start)+" a "+fmtShort(end)+".",
      motivo: inOut?"Admissão e desligamento entre cortes — período não contemplado em nenhum snapshot."
                   :"Admissão pós-corte — acrescentar período não contemplado no snapshot.",
      start, end, days, fte:(days/comp.days)*rateio};
  }
  return null; // desligado antes do corte, admissão futura etc.: sem ajuste nesta competência
}

/** Roda o motor na base inteira: dedupe, duplicidades e separação de erros. */
function analyze(employees,comp){
  const adjustments=[], errors=[];
  const seenAdj=new Map();      // groot|matricula|kind|start|end → dedupe exato
  const perPerson=new Map();    // pessoa → ajustes já gerados (sobreposição)
  const grootCount={}, matCount={};
  employees.forEach(e=>{
    if(e.groot) grootCount[e.groot]=(grootCount[e.groot]||0)+1;
    if(e.matricula) matCount[e.matricula]=(matCount[e.matricula]||0)+1;
  });

  employees.forEach(e=>{
    const r=detectAdjustment(e,comp);
    if(r===null) return;
    if(r.kind==="ERRO"){ errors.push({nome:e.nome||"(sem nome)", reason:r.errors.join("; ")}); return; }

    const person=(e.groot||"")+"§"+(e.matricula||"");
    const key=person+"§"+r.kind+"§"+r.start+"§"+r.end;
    if(seenAdj.has(key)){
      seenAdj.get(key).warnings.push("Linha duplicada no arquivo — ajuste gerado uma única vez");
      return; // segurança contra cobrança dupla
    }
    // sobreposição de períodos do mesmo tipo para a mesma pessoa → risco de duplicidade
    const prev=(perPerson.get(person)||[]).find(p=>p.kind===r.kind && r.start<=p.end && r.end>=p.start);
    if(prev){
      errors.push({nome:e.nome||"(sem nome)",
        reason:(e.groot?"GROOT "+e.groot:"matrícula "+e.matricula)+" duplicado com risco de ajuste duplicado (períodos sobrepostos)"});
      prev.warnings.push("Havia outra linha desta pessoa com período sobreposto — enviada para revisão");
      return;
    }
    if(e.groot && grootCount[e.groot]>1) r.warnings.push("GROOT ID aparece em "+grootCount[e.groot]+" linhas do arquivo");
    else if(e.matricula && matCount[e.matricula]>1) r.warnings.push("Matrícula aparece em "+matCount[e.matricula]+" linhas do arquivo");

    seenAdj.set(key,r);
    perPerson.set(person,(perPerson.get(person)||[]).concat([r]));
    adjustments.push(r);
  });

  // DESCONTAR primeiro, depois ACRESCENTAR; dentro do grupo, por data da movimentação
  adjustments.sort((a,b)=>{
    if(a.kind!==b.kind) return a.kind==="DESCONTAR"?-1:1;
    return a.movDate-b.movDate;
  });
  adjustments.forEach((r,i)=>{ r.id=i; r.include=true; });
  return {adjustments, errors};
}
