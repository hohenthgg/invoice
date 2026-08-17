/* Ajustes MELI — de onde vem a competência de uma fatura
   ================================================================

   A heurística por cluster de datas erra justamente onde mais dói: a
   fatura N+1 carrega os retroativos do mês N, todos datados do mês
   anterior, e a detecção escorrega para trás. Ignorar rateio negativo
   ajuda, mas não basta — um retroativo de ACRÉSCIMO é positivo.

   Por isso a competência passa a ser buscada numa ordem de evidência,
   da mais forte para a mais fraca, e a fonte usada é sempre mostrada
   ao usuário junto do nível de confiança. Incerteza não se esconde.
   ================================================================ */
"use strict";

const COMP_SOURCE = {
  PLANILHA:   {id:"PLANILHA",   label:"Campo da planilha",     confianca:"ALTA"},
  ARQUIVO:    {id:"ARQUIVO",    label:"Nome do arquivo",       confianca:"ALTA"},
  SEQUENCIA:  {id:"SEQUENCIA",  label:"Sequência com a outra fatura", confianca:"MEDIA"},
  HEURISTICA: {id:"HEURISTICA", label:"Heurística de datas",   confianca:"MEDIA"},
  MANUAL:     {id:"MANUAL",     label:"Confirmação manual",    confianca:"ALTA"},
  VAZIO:      {id:"VAZIO",      label:"Sem evidência",         confianca:"REVISAO"}
};

const MES_RE = "(janeiro|fevereiro|mar[cç]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)";
function mesPorNome(txt){
  const t=String(txt).normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase();
  const nomes=MONTHS.map(m=>m.normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase());
  const i=nomes.findIndex(n=>t.includes(n));
  return i>=0?i+1:0;
}

/** "Junho 2026", "Junho/2026", "Junho_2026", "06/2026", "2026-06" → {y,m} */
function parseCompetenceText(txt){
  if(txt===null||txt===undefined) return null;
  const s=String(txt).trim();
  if(!s) return null;
  const norm=s.normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase();

  let m=norm.match(new RegExp(MES_RE.normalize("NFD").replace(/[̀-ͯ]/g,"")+"[^0-9]{0,4}(\\d{4})"));
  if(m){ const mes=mesPorNome(m[1]); const ano=+m[2];
    if(mes&&ano>=1990&&ano<=2100) return {y:ano,m:mes}; }

  m=norm.match(/(\d{4})[-_/.](\d{1,2})(?!\d)/);          // 2026-06
  if(m){ const ano=+m[1], mes=+m[2];
    if(mes>=1&&mes<=12&&ano>=1990&&ano<=2100) return {y:ano,m:mes}; }

  m=norm.match(/(?:^|[^\d])(\d{1,2})[-_/.](\d{4})(?!\d)/); // 06/2026
  if(m){ const mes=+m[1], ano=+m[2];
    if(mes>=1&&mes<=12&&ano>=1990&&ano<=2100) return {y:ano,m:mes}; }

  // só o nome do mês, sem ano — insuficiente sozinho
  return null;
}

/* --------- PRIORIDADE 1: campo explícito na planilha --------- */
const ROTULOS_COMP=["competencia","competência","mes de referencia","mês de referência",
  "mes referencia","referencia","periodo de faturamento","período de faturamento","fatura"];
function fromWorkbook(wb){
  if(!wb||!wb.SheetNames) return null;
  const norm=s=>String(s??"").normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase().trim();
  for(const nome of wb.SheetNames){
    let rows;
    try{ rows=XLSX.utils.sheet_to_json(wb.Sheets[nome],{header:1,raw:false,defval:null}); }
    catch(err){ continue; }
    // varre as primeiras linhas em busca de um rótulo seguido do valor
    for(let i=0;i<Math.min(rows.length,12);i++){
      const row=rows[i]||[];
      for(let j=0;j<row.length;j++){
        const cel=norm(row[j]);
        if(!cel) continue;
        if(ROTULOS_COMP.some(r=>cel===norm(r)||cel.startsWith(norm(r)))){
          // valor na célula seguinte, ou na de baixo, ou no resto da própria célula
          const cands=[row[j+1], (rows[i+1]||[])[j], row[j]];
          for(const c of cands){
            const p=parseCompetenceText(c);
            if(p) return {comp:p, detalhe:'aba "'+nome+'"'};
          }
        }
      }
    }
    // título solto do tipo "Fatura — Junho/2026" nas duas primeiras linhas
    for(let i=0;i<Math.min(rows.length,2);i++){
      for(const c of (rows[i]||[])){
        const p=parseCompetenceText(c);
        if(p && /fatura|competenc|referenc/.test(norm(c))) return {comp:p, detalhe:'aba "'+nome+'"'};
      }
    }
  }
  return null;
}

/* --------- PRIORIDADE 2: nome do arquivo --------- */
function fromFileName(fileName){
  const p=parseCompetenceText(String(fileName||"").replace(/\.[^.]+$/,""));
  return p?{comp:p, detalhe:fileName}:null;
}

/* --------- PRIORIDADE 4: heurística sobre as linhas correntes --------- */
function fromCurrentLines(employees, wbHint){
  // Sem competência ainda conhecida, usa-se um palpite inicial só para separar
  // linhas correntes de retroativas e então redetectar sobre as correntes.
  const bruto=detectCompetence(employees, wbHint);
  const correntes=employees.filter(e=>{
    const c=classifyLine(e, bruto.comp);
    return !isRetroClass(c.classe);
  });
  const base=correntes.length>=Math.max(3,employees.length*0.2)?correntes:employees;
  const det=detectCompetence(base, wbHint);
  return {comp:{y:det.comp.y,m:det.comp.m}, confident:det.confident};
}

/** Resolve a competência de uma fatura na ordem de força da evidência.
 *  `par` (opcional) = {comp, papel:"origem"|"seguinte"} da outra fatura já
 *  carregada, usado como PRIORIDADE 3. */
function resolveCompetence(opts){
  const {wb, employees, fileName, retornoHint, par}=opts;

  const naPlanilha=fromWorkbook(wb);
  if(naPlanilha){
    return {comp:buildCompetence(naPlanilha.comp.y,naPlanilha.comp.m),
      fonte:COMP_SOURCE.PLANILHA, detalhe:naPlanilha.detalhe,
      confianca:COMP_SOURCE.PLANILHA.confianca};
  }

  const noNome=fromFileName(fileName);
  if(noNome){
    return {comp:buildCompetence(noNome.comp.y,noNome.comp.m),
      fonte:COMP_SOURCE.ARQUIVO, detalhe:noNome.detalhe,
      confianca:COMP_SOURCE.ARQUIVO.confianca};
  }

  const heur=employees&&employees.length?fromCurrentLines(employees,retornoHint):null;

  if(par&&par.comp){
    const esperado = par.papel==="origem"
      ? proximoMes(par.comp) : mesAnterior(par.comp);
    // a sequência decide quando a heurística não é confiante ou discorda
    if(!heur||!heur.confident||heur.comp.y!==esperado.y||heur.comp.m!==esperado.m){
      return {comp:buildCompetence(esperado.y,esperado.m),
        fonte:COMP_SOURCE.SEQUENCIA,
        detalhe:"assumida a partir de "+par.comp.label,
        confianca:COMP_SOURCE.SEQUENCIA.confianca};
    }
  }

  if(heur){
    return {comp:buildCompetence(heur.comp.y,heur.comp.m),
      fonte:COMP_SOURCE.HEURISTICA,
      detalhe:"cluster de movimentações das linhas de competência corrente",
      confianca:heur.confident?COMP_SOURCE.HEURISTICA.confianca:"REVISAO"};
  }

  const hoje=new Date();
  return {comp:buildCompetence(hoje.getUTCFullYear(),hoje.getUTCMonth()+1),
    fonte:COMP_SOURCE.VAZIO, detalhe:"nenhuma evidência no arquivo",
    confianca:"REVISAO"};
}

function proximoMes(c){ return c.m===12?{y:c.y+1,m:1}:{y:c.y,m:c.m+1}; }
function mesAnterior(c){ return c.m===1?{y:c.y-1,m:12}:{y:c.y,m:c.m-1}; }

if(typeof module!=="undefined"&&module.exports){
  module.exports={COMP_SOURCE, parseCompetenceText, fromFileName, resolveCompetence,
    proximoMes, mesAnterior};
}
