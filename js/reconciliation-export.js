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
const CONCIL_HEADERS=["GROOT ID","NOME","COMPETÊNCIA ORIGEM","COMPETÊNCIA DESTINO",
  "CLASSIFICAÇÃO DA LINHA","STATUS","ALERTAS","CONFIANÇA","BASE DA CONFIANÇA",
  "COBRANÇA ORIGINAL","AJUSTE ESPERADO","AJUSTE ENCONTRADO","DECISÃO DO USUÁRIO",
  "PERÍODO FINAL","FTE FINAL","ALTERAÇÃO APLICADA","IMPACTO FINANCEIRO",
  "OBSERVAÇÃO DO USUÁRIO","MOTIVO"];

function descreverAjuste(a){
  if(!a) return "—";
  const sinal=a.kind==="DESCONTAR"?"−":"+";
  return sinal+fmtShort(a.start)+" a "+fmtShort(a.end)+" · "+a.days
    +" dia"+(a.days===1?"":"s")+" · FTE "+a.fte.toFixed(4).replace(".",",");
}
function descreverCobranca(c){
  if(!c) return "—";
  if(!c.cobrado) return "não cobrada — "+c.base;
  return fmtShort(c.start)+" a "+fmtShort(c.end)+" · "+c.days+" dia"+(c.days===1?"":"s")
    +" · FTE "+c.fte.toFixed(4).replace(".",",");
}
function descreverImpacto(i){
  if(!i) return "—";
  if(!i.calculado) return i.motivo;
  const br=v=>v===null||v===undefined?"—":v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
  return "original "+br(i.original)+" · esperado "+br(i.esperado)
    +" · encontrado "+br(i.encontrado)+" · diferença "+br(i.diferenca);
}
/* Período e FTE que o usuário de fato escolheu, quando aceitou a sugestão. */
function periodoFinal(it){
  if(it.decisao!=="ACEITAR"||!it.sugestao) return "—";
  const s=it.sugestao;
  if(s.acao==="REMOVER") return "linha "+s.manterRow+" mantida";
  return fmtShort(s.start)+" a "+fmtShort(s.end)+" · "+s.days+" dia"+(s.days===1?"":"s");
}
function fteFinal(it){
  if(it.decisao!=="ACEITAR"||!it.sugestao) return "—";
  const s=it.sugestao;
  return typeof s.fte==="number"?s.fte.toFixed(4).replace(".",","):"—";
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
        // só a APARÊNCIA é herdada — nenhum valor de outra pessoa acompanha
        visual.eachCell({includeEmpty:true},(cell,c)=>{
          novo.getCell(c).style=JSON.parse(JSON.stringify(cell.style||{}));
        });
        novo.height=visual.height;
      }
      COLS_IDENTIDADE.forEach(k=>{ const c=map[k]; if(c!==undefined&&c>=0) novo.getCell(c+1).value=null; });
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

/** Identidade do colaborador na linha nova.
 *  A ESTRUTURA (colunas, estilo, formatos) vem sempre da fatura N+1; a
 *  IDENTIDADE vem do registro conciliado da pessoa, que quando ela não existe
 *  na N+1 só existe na fatura N. Nunca se reaproveita a identidade de outra
 *  pessoa da N+1 só porque a linha dela serviu de molde visual. */
function preencherIdentificacao(row, map, item){
  const set=(k,v)=>{ const c=map[k]; if(c!==undefined&&c>=0&&v!==null&&v!==undefined&&v!=="") row.getCell(c+1).value=v; };
  const raw=item.identidadeRaw||{};
  const campos=item.identidadeCampos||{};
  set("groot", raw.groot!==undefined&&raw.groot!==null?raw.groot:(campos.groot||item.groot));
  set("nome", raw.nome!==undefined&&raw.nome!==null?raw.nome:(campos.nome||item.nome));
  set("matricula", raw.matricula!==undefined&&raw.matricula!==null?raw.matricula:(campos.matricula||item.matricula));
  set("regime", raw.regime); set("cargo", raw.cargo);
  set("diasFolga", raw.diasFolga); set("escala", raw.escala);
}

/** Colunas que carregam identidade — precisam ser reescritas quando o molde
 *  visual veio de outra pessoa. */
const COLS_IDENTIDADE=["groot","nome","matricula","regime","cargo","diasFolga","escala"];

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
  ws.columns=[{width:14},{width:32},{width:16},{width:16},{width:22},{width:22},{width:28},
    {width:18},{width:44},{width:34},{width:34},{width:34},{width:20},{width:26},{width:12},
    {width:40},{width:44},{width:34},{width:60}];

  const hr=ws.addRow(CONCIL_HEADERS);
  hr.height=24;
  hr.eachCell(cell=>{
    cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF141414"}};
    cell.font={bold:true,color:{argb:"FFFFFFFF"},size:10,name:"Calibri"};
    cell.alignment={horizontal:"center",vertical:"middle",wrapText:true};
  });

  const CONF_LABEL={ALTA:"Alta",MEDIA:"Média",REVISAO:"Revisão necessária"};
  items.forEach(it=>{
    const meta=(typeof RECON_META!=="undefined"&&RECON_META[it.status])||{label:it.status};
    const alertas=(it.alerts||[]).map(a=>
      ((typeof RECON_META!=="undefined"&&RECON_META[a])||{label:a}).label).join(" + ");
    const classe=(typeof LINE_CLASS_LABEL!=="undefined"&&LINE_CLASS_LABEL[it.lineClassification])
      ||(it.achado&&typeof LINE_CLASS_LABEL!=="undefined"&&LINE_CLASS_LABEL[it.achado.classe])||"—";
    const row=ws.addRow([
      it.groot||"", it.nome, it.compOrigem, it.compAplicacao,
      classe, meta.label, alertas, CONF_LABEL[it.confianca]||it.confianca, it.confiancaMotivo||"",
      descreverCobranca(it.cobrancaOriginal),
      descreverAjuste(it.esperado), descreverAjuste(it.achado),
      DECISAO_LABEL[it.decisao]||it.decisao,
      periodoFinal(it), fteFinal(it),
      descreverAlteracao(it),
      descreverImpacto(it.impacto),
      it.observacao||"",
      it.diagnostico
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
