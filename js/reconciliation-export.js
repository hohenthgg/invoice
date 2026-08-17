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
   ESCRITA — identidade lógica, nunca posição física

   O perigo aqui é sutil: `spliceRows` desloca os índices das linhas
   seguintes. Um índice guardado antes da remoção passa a apontar para
   OUTRA pessoa. Se esse índice for usado como linha-modelo, a linha
   nova de João sai com o GROOT, o nome e a matrícula de Maria — erro
   nominal de faturamento, silencioso.

   A defesa é dupla, e a segunda torna a primeira redundante de
   propósito:

   1. FASE 0, antes de qualquer operação estrutural, cada inclusão
      captura um snapshot imutável do template (só estilos e altura) e a
      identidade lógica da pessoa. Depois disso nenhum índice antigo é
      consultado.

   2. O template NUNCA fornece valor de célula. Toda a identidade da
      linha nova vem do registro conciliado da pessoa. Ainda que o
      snapshot viesse da linha errada, o que se herda é aparência.

   E, antes de dar a linha por boa, os campos nominais escritos são
   conferidos contra os esperados. Divergiu, a inclusão é abortada e
   registrada — nunca gravada em silêncio.
   ================================================================ */

/** Identidade lógica do colaborador, obtida do registro conciliado —
 *  GROOT normalizado, matrícula como segundo recurso. Nunca da posição. */
function buildEmployeeIdentity(item){
  const raw=item.identidadeRaw||{};
  const campos=item.identidadeCampos||{};
  const pick=(a,b)=>(a!==undefined&&a!==null&&a!=="")?a:b;
  return {
    groot:      pick(raw.groot, pick(campos.groot, item.groot)),
    matricula:  pick(raw.matricula, pick(campos.matricula, item.matricula)),
    nome:       pick(raw.nome, pick(campos.nome, item.nome)),
    cargo:      raw.cargo, regime: raw.regime,
    diasFolga:  raw.diasFolga, escala: raw.escala
  };
}

/** Snapshot imutável de uma linha: só o que é aparência. Capturado ANTES
 *  de qualquer remoção, para que o deslocamento de índices não o afete. */
function captureRowTemplate(ws, rowIndex){
  if(!rowIndex||rowIndex<2||rowIndex>ws.rowCount) return null;
  const row=ws.getRow(rowIndex);
  const estilos=[];
  row.eachCell({includeEmpty:true},(cell,c)=>{
    estilos[c]=cell.style?JSON.parse(JSON.stringify(cell.style)):null;
  });
  return {estilos, height:row.height, origem:rowIndex};
}

/** Linha desta planilha que serve de molde visual. Preferência para a
 *  linha da própria pessoa; qualquer linha de dados serve de reserva,
 *  porque dela só se aproveita o estilo. */
function resolveTemplateRow(ws, item, removidas){
  const candidatos=[item.achado&&item.achado.srcRow, item.modeloRow];
  for(const r of candidatos){
    if(r&&!removidas.includes(r)&&r>=2&&r<=ws.rowCount) return r;
  }
  for(let r=2;r<=ws.rowCount;r++) if(!removidas.includes(r)) return r;
  return null;
}

/** FASE 0 — resolve tudo que depende de índice enquanto os índices ainda
 *  valem, e devolve um plano imutável. */
function prepararInclusoes(ws, plano, removidas){
  return plano.incluir.map(({item,sug})=>({
    item, sug,
    identity:buildEmployeeIdentity(item),
    template:captureRowTemplate(ws, resolveTemplateRow(ws, item, removidas))
  }));
}

function applyEmployeeIdentity(row, map, identity){
  const set=(k,v)=>{ const c=map[k];
    if(c!==undefined&&c>=0&&v!==undefined&&v!==null&&v!=="") row.getCell(c+1).value=v; };
  set("groot", identity.groot);
  set("nome", identity.nome);
  set("matricula", identity.matricula);
  set("regime", identity.regime);
  set("cargo", identity.cargo);
  set("diasFolga", identity.diasFolga);
  set("escala", identity.escala);
}

function applyAdjustmentFields(row, map, sug){
  const col=k=>(map[k]!==undefined&&map[k]>=0)?map[k]+1:null;
  const cIni=col("inicio"), cFim=col("fim"), cRat=col("rateio");
  if(cIni) escreverData(row.getCell(cIni), sug.start);
  if(cFim) escreverData(row.getCell(cFim), sug.end);
  if(cRat) escreverRateio(row.getCell(cRat), sug.kind==="DESCONTAR"?-sug.rateio:sug.rateio);
}

/** Confere o que foi realmente escrito contra o que era esperado.
 *  Devolve a lista de divergências — vazia quando a linha está correta. */
function validarIdentidadeDaLinha(row, map, identity){
  const erros=[];
  const lido=k=>{ const c=map[k]; return (c!==undefined&&c>=0)?row.getCell(c+1).value:undefined; };
  const igualId=(a,b)=>normalizeGroot(a)===normalizeGroot(b);
  const igualTexto=(a,b)=>String(a??"").trim().toUpperCase()===String(b??"").trim().toUpperCase();

  if(identity.groot!==undefined&&identity.groot!==null&&identity.groot!==""
     &&!igualId(lido("groot"), identity.groot))
    erros.push("GROOT gravado ("+lido("groot")+") diferente do esperado ("+identity.groot+")");
  if(identity.matricula!==undefined&&identity.matricula!==null&&identity.matricula!==""
     &&!igualId(lido("matricula"), identity.matricula))
    erros.push("matrícula gravada ("+lido("matricula")+") diferente da esperada ("+identity.matricula+")");
  if(identity.nome&&!igualTexto(lido("nome"), identity.nome))
    erros.push("nome gravado ("+lido("nome")+") diferente do esperado ("+identity.nome+")");
  return erros;
}

/** Cria a linha conciliada: estrutura do template, identidade da pessoa.
 *  Devolve {row, erros}. Com erros, a linha é removida e nada é gravado. */
function createLaborRow(ws, map, prep){
  const row=ws.addRow([]);
  if(prep.template){
    prep.template.estilos.forEach((st,c)=>{ if(st) row.getCell(c).style=JSON.parse(JSON.stringify(st)); });
    if(prep.template.height) row.height=prep.template.height;
  }
  applyEmployeeIdentity(row, map, prep.identity);
  applyAdjustmentFields(row, map, prep.sug);
  const erros=validarIdentidadeDaLinha(row, map, prep.identity);
  if(erros.length){ ws.spliceRows(row.number,1); return {row:null, erros}; }
  row.commit&&row.commit();
  return {row, erros:[]};
}

async function gerarFaturaConciliada(ctx){
  const {buffer, sheetName, map, items, comp, fileNameBase}=ctx;
  const plano=planoDeAlteracoes(items);

  const wb=new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);                       // preserva o workbook inteiro
  const ws=wb.getWorksheet(sheetName)||wb.worksheets[0];
  if(!ws) throw new Error("Aba de Labor não encontrada no arquivo da fatura seguinte.");

  const paraRemover=[...new Set(plano.remover.flatMap(r=>r.sug.alvoRows||[]))]
    .sort((a,b)=>b-a);

  // ---------- FASE 0: snapshots, enquanto os índices ainda valem ----------
  const inclusoes=prepararInclusoes(ws, plano, paraRemover);

  // ---------- 1) substituições (alteram células, não deslocam nada) ----------
  plano.substituir.forEach(({sug})=>{
    applyAdjustmentFields(ws.getRow(sug.alvoRow), map, sug);
  });

  // ---------- 2) remoções, de baixo para cima ----------
  paraRemover.forEach(r=>ws.spliceRows(r,1));

  // ---------- 3) inclusões, a partir dos snapshots ----------
  const falhas=[];
  inclusoes.forEach(prep=>{
    const {erros}=createLaborRow(ws, map, prep);
    if(erros.length){
      falhas.push({nome:prep.item.nome, erros});
      prep.item.exportacaoAbortada=erros.join("; ");
    }
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
  return {plano, arquivo:a.download, falhas};
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
  module.exports={planoDeAlteracoes, resumoPrevia, descreverAjuste, descreverAlteracao,
    buildEmployeeIdentity, captureRowTemplate, resolveTemplateRow, prepararInclusoes,
    applyEmployeeIdentity, applyAdjustmentFields, validarIdentidadeDaLinha, createLaborRow};
}
