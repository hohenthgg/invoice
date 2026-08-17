/* Ajustes MELI — leitura da planilha */
"use strict";

/* ================================================================
   IMPORTAÇÃO
   ================================================================ */
const state={ employees:[], adjustments:[], errors:[], comp:null, confident:true, fileName:null };

function normHeader(h){ return String(h||"").trim().toUpperCase().replace(/\s+/g," "); }
function mapColumns(headerRow){
  const map={}, normed=headerRow.map(normHeader);
  for(const k in COLS){ map[k]=-1;
    for(const a of COLS[k]){ const i=normed.indexOf(normHeader(a)); if(i>=0){map[k]=i;break;} } }
  return map;
}
/** Localiza a aba de Labor e devolve as linhas cruas com o mapa de colunas.
 *  Função pura: não toca em estado global nem na tela — os dois modos da aba
 *  (projetar e conciliar) leem a planilha exatamente pelo mesmo caminho.
 *  Devolve {erro} quando a planilha não serve. */
function readLaborSheet(wb){
  let sheet=wb.Sheets[SHEET_NAME];
  let usedName=SHEET_NAME;
  if(!sheet){
    for(const n of wb.SheetNames){
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true});
      if(rows.length){ const m=mapColumns(rows[0]||[]); if(m.groot>=0&&m.inicio>=0){ sheet=wb.Sheets[n]; usedName=n; break; } }
    }
  }
  if(!sheet) return {erro:'Aba "'+SHEET_NAME+'" não encontrada.'};
  const rows=XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:null});
  if(!rows.length) return {erro:"Aba vazia."};
  const header=rows[0].map(h=>h===null?"":String(h).trim());
  const map=mapColumns(header);
  const missing=["groot","nome","inicio"].filter(k=>map[k]<0);
  if(missing.length) return {erro:"Cabeçalhos não encontrados: "+missing.join(", ")};
  return {sheet, sheetName:usedName, header, map, rows};
}

/** Converte as linhas cruas em colaboradores normalizados. */
function parseEmployees(rows,map){
  const employees=[];
  for(let i=1;i<rows.length;i++){
    const row=rows[i];
    if(!row||row.every(c=>c===null||c==="")) continue;
    const g=k=>map[k]>=0?row[map[k]]:null;
    employees.push({
      srcRow:i+1,                       // linha 1-based na planilha original (para herdar estilo)
      groot:g("groot")===null?null:String(g("groot")).trim(),
      nome:g("nome")===null?"":String(g("nome")).trim(),
      matricula:g("matricula")===null?null:String(g("matricula")).trim(),
      inicio:parseExcelDate(g("inicio")),
      fim:parseExcelDate(g("fim")),
      rateio:parseRateio(g("rateio")),
      // valores ORIGINAIS preservados para a exportação (§41)
      raw:{ groot:g("groot"), nome:g("nome"), matricula:g("matricula"), regime:g("regime"),
            cargo:g("cargo"), rateio:g("rateio"), diasFolga:g("diasFolga"), escala:g("escala") }
    });
  }
  return employees;
}

/** Lê a distribuição de "Data Trab." da aba Retorno MELI, evidência principal
 *  para a detecção automática da competência. Devolve null quando não há aba. */
function readRetornoHint(wb){
  const ret=wb.Sheets[RETORNO_SHEET];
  if(!ret) return null;
  const rrows=XLSX.utils.sheet_to_json(ret,{header:1,raw:true,defval:null});
  let hIdx=-1,cIdx=-1;
  for(let i=0;i<Math.min(rrows.length,10);i++){
    const j=(rrows[i]||[]).findIndex(c=>normHeader(c)==="DATA TRAB.");
    if(j>=0){ hIdx=i; cIdx=j; break; }
  }
  if(hIdx<0) return null;
  const freq={}; let total=0;
  for(let i=hIdx+1;i<rrows.length;i++){
    const d=parseExcelDate((rrows[i]||[])[cIdx]);
    if(isValidYmd(d)){ const p=ymdParts(d); const k=p.y+"-"+p.m; freq[k]=(freq[k]||0)+1; total++; }
  }
  const top=Object.entries(freq).sort((a,b)=>b[1]-a[1])[0];
  if(!top||!total) return null;
  const [y,m]=top[0].split("-").map(Number);
  return { top:{y,m}, total, topShare:top[1]/total };
}

function importWorkbook(data,fileName){
  const wb=XLSX.read(data,{type:"array",cellDates:true});
  const read=readLaborSheet(wb);
  if(read.erro){ alert(read.erro); return false; }
  const {map, rows}=read;
  const usedName=read.sheetName;

  const employees=parseEmployees(rows,map);
  state.employees=employees;
  state.fileName=fileName;
  state.srcMap=map;                     // chave → índice 0-based da coluna original
  state.srcSheetName=usedName;
  state.srcBuffer=data.slice().buffer;   // cópia exata dos bytes originais, só para ler estilos

  const det=detectCompetence(employees,readRetornoHint(wb));
  state.comp=det.comp;
  state.confident=det.confident;
  document.getElementById("fileName").textContent=fileName+" · "+employees.length+" linhas";
  return true;
}
