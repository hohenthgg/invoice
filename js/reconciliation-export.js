/* Ajustes MELI — geração da Fatura Conciliada
   ================================================================

   Parte SEMPRE da fatura N+1 e SEMPRE gera um arquivo novo: o original
   não é tocado, nem é reconstruído do zero. O workbook é carregado
   inteiro pelo ExcelJS e devolvido inteiro — todas as abas, fórmulas,
   estilos, cores, larguras, merges e formatos —, e só as células
   estritamente necessárias da aba de Labor mudam.

   Só chega aqui o que o usuário aceitou explicitamente. Um item em
   "manter como está", "ignorar" ou "revisar" não produz alteração
   nenhuma; no máximo vira uma linha na aba de auditoria.
   ================================================================ */
"use strict";

/** Plano de alterações derivado das decisões — a mesma estrutura alimenta
 *  a prévia na tela e a escrita do arquivo, para que não haja como a
 *  prévia mostrar uma coisa e o Excel sair outra. */
function planoDeAlteracoes(items){
  const incluir=[], substituir=[], remover=[];
  items.forEach(it=>{
    if(it.decisao!=="ACEITAR"||!it.sugestao) return;
    const s=it.sugestao;
    if(s.acao==="INCLUIR")      incluir.push({item:it, sug:s});
    else if(s.acao==="SUBSTITUIR") substituir.push({item:it, sug:s});
    else if(s.acao==="REMOVER")    remover.push({item:it, sug:s});
  });
  const linhasRemovidas=remover.reduce((n,r)=>n+(r.sug.alvoRows||[]).length,0);
  return {incluir, substituir, remover, linhasRemovidas,
    total:incluir.length+substituir.length+linhasRemovidas};
}

/** Resumo textual da prévia: quantas linhas entram, mudam, saem e ficam. */
function resumoPrevia(plano, totalLinhas){
  const mexidas=plano.substituir.length+plano.linhasRemovidas;
  return {
    incluidas:plano.incluir.length,
    alteradas:plano.substituir.length,
    removidas:plano.linhasRemovidas,
    intactas:Math.max(0,totalLinhas-mexidas)
  };
}

const CONCIL_SHEET="CONCILIAÇÃO";
const CONCIL_HEADERS=["GROOT ID","NOME","STATUS","ESPERADO","ENCONTRADO",
  "DECISÃO DO USUÁRIO","ALTERAÇÃO APLICADA","MOTIVO"];

function descreverAjuste(a){
  if(!a) return "—";
  const sinal=a.kind==="DESCONTAR"?"−":"+";
  return sinal+fmtShort(a.start)+" a "+fmtShort(a.end)+" · "+a.days
    +" dia"+(a.days===1?"":"s")+" · FTE "+a.fte.toFixed(4).replace(".",",");
}
const DECISAO_LABEL={MANTER:"Manter como está", ACEITAR:"Aceitar sugestão",
  IGNORAR:"Ignorar apontamento", REVISAR:"Revisar manualmente"};

/** Texto do que efetivamente mudou no arquivo, por item. */
function descreverAlteracao(it){
  if(it.decisao!=="ACEITAR"||!it.sugestao) return "Nenhuma";
  const s=it.sugestao;
  if(s.acao==="INCLUIR")    return "Linha incluída: "+descreverAjuste(s);
  if(s.acao==="SUBSTITUIR") return "Linha "+s.alvoRow+" alterada para "+descreverAjuste(s);
  if(s.acao==="REMOVER")    return "Linha(s) removida(s): "+(s.alvoRows||[]).join(", ")
    +" (mantida a linha "+s.manterRow+")";
  return "Nenhuma";
}

/* ================================================================
   ESCRITA

   Ordem obrigatória das operações na aba de Labor:
     1) substituições  — alteram células, não deslocam nada;
     2) remoções       — de baixo para cima, senão os índices andam;
     3) inclusões      — sempre no fim, herdando o estilo de uma linha
                         existente da própria pessoa quando houver.
   ================================================================ */
async function gerarFaturaConciliada(ctx){
  const {buffer, sheetName, map, items, comp, fileNameBase}=ctx;
  const plano=planoDeAlteracoes(items);

  const wb=new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);                       // preserva o workbook inteiro
  const ws=wb.getWorksheet(sheetName)||wb.worksheets[0];
  if(!ws) throw new Error("Aba de Labor não encontrada no arquivo da fatura seguinte.");

  const col=k=>(map[k]!==undefined&&map[k]>=0)?map[k]+1:null;   // 1-based
  const cIni=col("inicio"), cFim=col("fim"), cRat=col("rateio");

  // ---------- 1) substituições ----------
  plano.substituir.forEach(({sug})=>{
    const row=ws.getRow(sug.alvoRow);
    if(cIni) escreverData(row.getCell(cIni), sug.start);
    if(cFim) escreverData(row.getCell(cFim), sug.end);
    if(cRat) escreverRateio(row.getCell(cRat), sug.kind==="DESCONTAR"?-sug.rateio:sug.rateio);
  });

  // ---------- 2) remoções, de baixo para cima ----------
  const paraRemover=[...new Set(plano.remover.flatMap(r=>r.sug.alvoRows||[]))]
    .sort((a,b)=>b-a);
  paraRemover.forEach(r=>ws.spliceRows(r,1));

  // ---------- 3) inclusões no fim ----------
  plano.incluir.forEach(({item,sug})=>{
    const modeloRow=acharModelo(ws, item, paraRemover);
    const novo=ws.addRow([]);
    if(modeloRow){
      // mesma pessoa nesta planilha: herda estilo E cadastro
      modeloRow.eachCell({includeEmpty:true},(cell,c)=>{
        const alvo=novo.getCell(c);
        alvo.style=JSON.parse(JSON.stringify(cell.style||{}));
        if(c!==cIni&&c!==cFim&&c!==cRat) alvo.value=cell.value;
      });
      novo.height=modeloRow.height;
    } else {
      // pessoa ausente desta fatura: só a aparência é herdada
      const visual=acharModeloVisual(ws, paraRemover);
      if(visual){
        visual.eachCell({includeEmpty:true},(cell,c)=>{
          novo.getCell(c).style=JSON.parse(JSON.stringify(cell.style||{}));
        });
        novo.height=visual.height;
      }
      preencherIdentificacao(novo, map, item);
    }
    if(cIni) escreverData(novo.getCell(cIni), sug.start);
    if(cFim) escreverData(novo.getCell(cFim), sug.end);
    if(cRat) escreverRateio(novo.getCell(cRat), sug.kind==="DESCONTAR"?-sug.rateio:sug.rateio);
    novo.commit&&novo.commit();
  });

  // ---------- aba de auditoria ----------
  montarAbaConciliacao(wb, items);

  const buf=await wb.xlsx.writeBuffer();
  const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="FATURA_CONCILIADA_"+(fileNameBase||"SMG3")+"_"+comp.fileLabel.replace(/ /g,"_")+".xlsx";
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();},2000);
  return {plano, arquivo:a.download};
}

/** Linha da MESMA pessoa dentro desta planilha, para a nova herdar estilo e
 *  campos cadastrais. Só valem índices da própria fatura N+1: a linha de
 *  origem vem do outro arquivo e apontaria para outra pessoa.
 *  Linhas removidas nesta mesma geração não servem de modelo. */
function acharModelo(ws, item, removidas){
  const candidatos=[item.achado&&item.achado.srcRow, item.modeloRow];
  for(const r of candidatos){
    if(r && !removidas.includes(r) && r>=2 && r<=ws.rowCount) return ws.getRow(r);
  }
  return null;
}
/** Sem linha da pessoa no arquivo de destino, herda só a APARÊNCIA de uma
 *  linha qualquer — os dados vêm do cadastro da fatura de origem. */
function acharModeloVisual(ws, removidas){
  for(let r=2;r<=ws.rowCount;r++) if(!removidas.includes(r)) return ws.getRow(r);
  return null;
}

function preencherIdentificacao(row, map, item){
  const set=(k,v)=>{ const c=map[k]; if(c!==undefined&&c>=0&&v!==null&&v!==undefined) row.getCell(c+1).value=v; };
  const emp=(item.esperado&&item.esperado.emp)||(item.achado&&item.achado.emp)||{};
  set("groot", emp.raw&&emp.raw.groot!==undefined?emp.raw.groot:item.groot);
  set("nome", item.nome);
  set("matricula", emp.raw&&emp.raw.matricula!==undefined?emp.raw.matricula:item.matricula);
  if(emp.raw){
    set("regime", emp.raw.regime); set("cargo", emp.raw.cargo);
    set("diasFolga", emp.raw.diasFolga); set("escala", emp.raw.escala);
  }
}

function escreverData(cell, v){
  cell.value=ymdToExcelDate(v);
  if(!/[dmy]/i.test(String(cell.numFmt||""))) cell.numFmt="dd/mm/yyyy";
}
function escreverRateio(cell, v){
  cell.value=v;
  if(!/%/.test(String(cell.numFmt||""))){
    cell.numFmt=(Math.abs(v*100-Math.round(v*100))<1e-9)?"0%":"0.00%";
  }
}

/** Aba de rastreabilidade: documenta TUDO que foi analisado, inclusive o
 *  que o usuário decidiu não tratar. Não participa de nenhum cálculo. */
function montarAbaConciliacao(wb, items){
  const antiga=wb.getWorksheet(CONCIL_SHEET);
  if(antiga) wb.removeWorksheet(antiga.id);
  const ws=wb.addWorksheet(CONCIL_SHEET,{views:[{state:"frozen",ySplit:1}]});
  ws.columns=[{width:14},{width:34},{width:24},{width:34},{width:34},{width:20},{width:40},{width:52}];

  const hr=ws.addRow(CONCIL_HEADERS);
  hr.height=24;
  hr.eachCell(cell=>{
    cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF141414"}};
    cell.font={bold:true,color:{argb:"FFFFFFFF"},size:10,name:"Calibri"};
    cell.alignment={horizontal:"center",vertical:"middle",wrapText:true};
  });

  items.forEach(it=>{
    const meta=(typeof RECON_META!=="undefined"&&RECON_META[it.status])||{label:it.status};
    const row=ws.addRow([
      it.groot||"", it.nome, meta.label,
      descreverAjuste(it.esperado), descreverAjuste(it.achado),
      DECISAO_LABEL[it.decisao]||it.decisao,
      descreverAlteracao(it),
      it.observacao?it.diagnostico+" | Observação: "+it.observacao:it.diagnostico
    ]);
    row.eachCell({includeEmpty:true},cell=>{
      cell.font={size:10,name:"Calibri"};
      cell.alignment={vertical:"top",wrapText:true};
    });
  });
  ws.autoFilter={from:{row:1,column:1},to:{row:1,column:CONCIL_HEADERS.length}};
  return ws;
}

if(typeof module!=="undefined"&&module.exports){
  module.exports={planoDeAlteracoes, resumoPrevia, descreverAjuste, descreverAlteracao};
}
