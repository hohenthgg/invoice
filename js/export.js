/* Ajustes MELI — exportação no padrão MELI */
"use strict";

/* ================================================================
   EXPORTAÇÃO — herda o estilo da planilha importada
   Em vez de imitar cores, o exportador lê a formatação real do
   arquivo de origem (fonte, preenchimento, bordas, alinhamento,
   formatos de data/percentual) e a reaplica célula a célula. Como o
   estilo vem da própria linha original, cores associadas a categorias
   (REGIME, CARGO, ESCALA...) são preservadas automaticamente.
   O bloco do motor à direita reaproveita o mesmo estilo, apenas com
   preenchimento de cabeçalho levemente diferenciado.
   ================================================================ */
const ADJ_HEADERS=["TIPO DE AJUSTE","INÍCIO DO AJUSTE","FIM DO AJUSTE","DIAS DE AJUSTE","FTE AJUSTE","COMPETÊNCIA ORIGEM","COMPETÊNCIA DE APLICAÇÃO","MOTIVO"];
const THIN={style:"thin",color:{argb:"FF000000"}};
const BORDER={top:THIN,bottom:THIN,left:THIN,right:THIN};
/* usados apenas se a planilha de origem não trouxer estilo algum */
const FALLBACK_HEADER={font:{name:"Calibri",size:9,bold:true,color:{argb:"FF000000"}},
  fill:{type:"pattern",pattern:"solid",fgColor:{argb:"FFFFFF00"}},
  border:BORDER, alignment:{horizontal:"center",vertical:"middle"}};
const FALLBACK_BODY={font:{name:"Calibri",size:9}, border:BORDER,
  alignment:{horizontal:"center",vertical:"middle"}};

const clone=o=>o?JSON.parse(JSON.stringify(o)):undefined;
const isDateFmt=f=>typeof f==="string" && /[dmy]/i.test(f) && !/^(general)$/i.test(f) && /d/i.test(f) && /y/i.test(f);

/** Lê a planilha de origem com ExcelJS só para extrair formatação. */
async function loadSourceStyles(){
  if(!state.srcBuffer) return null;
  try{
    const wb=new ExcelJS.Workbook();
    await wb.xlsx.load(state.srcBuffer);
    const ws=wb.getWorksheet(state.srcSheetName)||wb.worksheets[0];
    return ws||null;
  }catch(err){ return null; }
}

async function exportAdjustments(){
  const c=state.comp;
  const rows=state.adjustments.filter(r=>r.include);
  if(!rows.length) return;

  const src=await loadSourceStyles();
  const map=state.srcMap||{};
  const KEYS=["groot","nome","matricula","regime","cargo","inicio","fim","rateio","diasFolga","escala"];
  const srcCol=k=>(map[k]!==undefined&&map[k]>=0)?map[k]+1:null;   // 1-based no arquivo original

  // estilo do cabeçalho e do corpo, por coluna, colhidos da origem
  const headerStyle={}, bodyStyle={};
  KEYS.forEach(k=>{
    const col=srcCol(k);
    if(src&&col){
      const h=src.getRow(1).getCell(col);
      headerStyle[k]=clone(h.style)||clone(FALLBACK_HEADER);
      if(!headerStyle[k].border) headerStyle[k].border=clone(BORDER);
    } else headerStyle[k]=clone(FALLBACK_HEADER);
  });

  const wb=new ExcelJS.Workbook();
  const ws=wb.addWorksheet("Ajustes MELI",{views:[{state:"frozen",ySplit:1}]});

  // ---------- cabeçalho ----------
  const headers=MELI_HEADERS.concat(ADJ_HEADERS);
  const hr=ws.addRow(headers);
  // estilo de referência do bloco MELI (o mais típico entre os cabeçalhos)
  const alignCount={};
  KEYS.forEach(k=>{ const h=(headerStyle[k]&&headerStyle[k].alignment&&headerStyle[k].alignment.horizontal)||"—";
    alignCount[h]=(alignCount[h]||0)+1; });
  const modalAlign=Object.entries(alignCount).sort((a,b)=>b[1]-a[1])[0][0];
  const refKey=KEYS.find(k=>((headerStyle[k]&&headerStyle[k].alignment&&headerStyle[k].alignment.horizontal)||"—")===modalAlign)||KEYS[0];
  const refHeader=clone(headerStyle[refKey])||clone(FALLBACK_HEADER);
  hr.eachCell({includeEmpty:true},(cell,col)=>{
    if(col<=MELI_HEADERS.length){
      cell.style=clone(headerStyle[KEYS[col-1]])||clone(FALLBACK_HEADER);
    } else {
      const st=clone(refHeader);                     // mesma fonte/borda/alinhamento do MELI
      st.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF2F2F2"}}; // continuação discreta
      st.numFmt=undefined;
      cell.style=st;
    }
  });

  // ---------- linhas ----------
  rows.forEach(r=>{
    const e=r.emp, raw=e.raw;
    const values=[
      raw.groot??"", raw.nome??"", raw.matricula??"", raw.regime??"", raw.cargo??"",
      e.inicio!==null&&isValidYmd(e.inicio)?ymdToExcelDate(e.inicio):"",
      e.fim!==null&&isValidYmd(e.fim)?ymdToExcelDate(e.fim):"",
      raw.rateio??"", raw.diasFolga??"", raw.escala??"",
      r.kind, ymdToExcelDate(r.start), ymdToExcelDate(r.end), r.days,
      +r.fte.toFixed(6), c.label, c.nextLabel, r.motivo
    ];
    const row=ws.addRow(values);

    // 1) campos MELI: estilo copiado da MESMA célula da linha original,
    //    preservando cores de categoria (regime, cargo, escala...)
    let lastMeliStyle=clone(FALLBACK_BODY);
    KEYS.forEach((k,i)=>{
      const col=srcCol(k);
      let st=null;
      if(src&&col&&e.srcRow) st=clone(src.getRow(e.srcRow).getCell(col).style);
      if(!st||!Object.keys(st).length) st=clone(bodyStyle[k])||clone(FALLBACK_BODY);
      if(!st.font) st.font=clone(FALLBACK_BODY.font);
      if(!st.border) st.border=clone(BORDER);
      if(!st.alignment) st.alignment=clone(FALLBACK_BODY.alignment);
      row.getCell(i+1).style=st;
      lastMeliStyle=st;
    });

    // datas: mantém o formato do arquivo de origem; só define um se não houver
    [6,7].forEach(col=>{
      const cell=row.getCell(col);
      if(!isDateFmt(cell.numFmt)) cell.numFmt="dd/mm/yyyy";
    });
    // GROOT/matrícula: só intervém se o número for grande a ponto de virar notação científica
    [1,3].forEach(col=>{
      const cell=row.getCell(col);
      if(typeof cell.value==="number" && Math.abs(cell.value)>=1e10 && !/0/.test(cell.numFmt||"")) cell.numFmt="0";
    });
    // rateio: preserva o formato percentual original; senão aplica um coerente
    const rt=row.getCell(8);
    if(typeof rt.value==="number" && !/%/.test(rt.numFmt||"")){
      rt.numFmt=(Math.abs(rt.value*100-Math.round(rt.value*100))<1e-9)?"0%":"0.00%";
    }

    // 2) bloco do motor: mesma linguagem visual, sem preenchimento de categoria
    const dateFmt = isDateFmt(row.getCell(7).numFmt) ? row.getCell(7).numFmt : "dd/mm/yyyy";
    for(let col=MELI_HEADERS.length+1; col<=headers.length; col++){
      const st=clone(lastMeliStyle);
      delete st.fill;                    // corpo branco (§34)
      st.numFmt=undefined;
      if(st.font) st.font.bold=false;
      row.getCell(col).style=st;
    }
    row.getCell(12).numFmt=dateFmt;
    row.getCell(13).numFmt=dateFmt;
    row.getCell(14).numFmt="0";
    row.getCell(15).numFmt="0.0000";
    row.getCell(18).alignment={...(row.getCell(18).alignment||{}),horizontal:"left",wrapText:false};
    // destaque discreto somente na célula do tipo
    const tipo=row.getCell(11);
    tipo.font={...(tipo.font||{}),bold:true,
      color:{argb:r.kind==="DESCONTAR"?"FFB3372F":"FF1A7F4B"}};
  });

  // ---------- larguras: herdadas da origem quando existirem ----------
  KEYS.forEach((k,i)=>{
    const col=srcCol(k);
    const w=(src&&col&&src.getColumn(col).width)||[12,42,13,21,19,14,14,10,22,20][i];
    ws.getColumn(i+1).width=w;
  });
  [15,13,13,8,10,16,18,55].forEach((w,i)=>ws.getColumn(MELI_HEADERS.length+1+i).width=w);
  ws.autoFilter={from:{row:1,column:1},to:{row:ws.rowCount,column:headers.length}};

  const buf=await wb.xlsx.writeBuffer();
  const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="Ajustes MELI - "+c.fileLabel+".xlsx";
  a.click();
  URL.revokeObjectURL(a.href);
}
