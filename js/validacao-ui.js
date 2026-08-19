/* Validação Template — leitura do .xlsx e tela da aba
   ================================================================

   Este arquivo faz três coisas e nenhuma delas é raciocinar sobre a
   fatura: lê o arquivo, chama js/validacao.js e desenha o resultado.
   O raciocínio mora lá, sem DOM, para poder ser testado.

   A leitura localiza abas e colunas PELO NOME, nunca pela posição: o
   template evolui de versão (este veio como 6.1) e uma coluna nova no
   meio não pode deslocar tudo em silêncio.
   ================================================================ */
"use strict";
(function(){

const $ = id => document.getElementById(id);
const norm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"")
  .toUpperCase().replace(/\s+/g," ").trim();

/* Estado da aba. `marcas` guarda a decisão humana por achado. */
const V = { auditoria:null, filtro:"todos", marcas:{}, arquivo:"" };

/* ================================================================
   LEITURA DO ARQUIVO
   ================================================================ */
/* Uma célula de data pode vir como Date, serial numérico ou texto. */
function lerData(v){
  if(v === null || v === undefined || v === "") return null;
  if(v instanceof Date) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
  if(typeof v === "number"){
    const d = XLSX.SSF.parse_date_code(v);
    return d ? new Date(Date.UTC(d.y, d.m-1, d.d)) : null;
  }
  const m = String(v).match(/(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{1,4})/);
  if(!m) return null;
  const a = +m[1], b = +m[2], c = +m[3];
  if(a > 1000) return new Date(Date.UTC(a, b-1, c));                 // aaaa-mm-dd
  return new Date(Date.UTC(c < 100 ? c+2000 : c, b-1, a));           // dd/mm/aaaa
}
const num = v => { const n = Number(v); return isFinite(n) ? n : null; };

/* Acha a aba pelo nome, tolerando acento e caixa. */
function acharAba(wb, ...alvos){
  for(const alvo of alvos){
    const a = norm(alvo);
    const nome = wb.SheetNames.find(n => norm(n) === a) ||
                 wb.SheetNames.find(n => norm(n).includes(a));
    if(nome) return nome;
  }
  return null;
}
/* Acha a linha de cabeçalho: a primeira das 10 iniciais que reconhece
   todas as chaves pedidas. */
function acharCabecalho(rows, chaves){
  for(let i=0;i<Math.min(rows.length,10);i++){
    const cels = (rows[i]||[]).map(norm);
    if(chaves.every(k => cels.some(c => c.includes(norm(k))))) return { idx:i, cels };
  }
  return null;
}
const col = (cels,...alts) => {
  for(const a of alts){ const j = cels.findIndex(c => c === norm(a)); if(j >= 0) return j; }
  for(const a of alts){ const j = cels.findIndex(c => c.includes(norm(a))); if(j >= 0) return j; }
  return -1;
};
const cel = (row,i) => i >= 0 ? row[i] : "";

function lerLabor(wb){
  const nome = acharAba(wb,"LABOR");
  if(!nome) return { erro:'Não encontrei a aba "LABOR".' };
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[nome],{header:1,defval:""});
  const h = acharCabecalho(rows,["GROOT ID","NOME","DATA DE INICIO"]);
  if(!h) return { erro:'A aba "'+nome+'" não tem o cabeçalho esperado (GROOT ID, NOME, DATA DE INÍCIO).' };
  const c = h.cels;
  const iG=col(c,"GROOT ID"), iN=col(c,"NOME"), iM=col(c,"MATRICULA"), iR=col(c,"REGIME DE CONTRATO"),
        iC=col(c,"CARGO"), iI=col(c,"DATA DE INICIO"), iF=col(c,"DATA FIM"),
        iQ=col(c,"QUANTIDADE"), iV=col(c,"VALOR FINAL"), iU=col(c,"VALOR UNITARIO");
  const out = [];
  for(let i=h.idx+1;i<rows.length;i++){
    const r = rows[i]; if(!r) continue;
    const nomeP = String(cel(r,iN) ?? "").trim();
    const groot = String(cel(r,iG) ?? "").trim();
    const valor = num(cel(r,iV));
    /* Linha do template ainda em branco não é achado nenhum. */
    if(!nomeP && !groot && !valor) continue;
    out.push({ aba:"LABOR", linha:i+1, groot, nome:nomeP, matricula:String(cel(r,iM) ?? "").trim(),
      regime:String(cel(r,iR) ?? "").trim(), cargo:String(cel(r,iC) ?? "").trim(),
      ini:lerData(cel(r,iI)), fim:lerData(cel(r,iF)),
      qtd:num(cel(r,iQ)), unit:num(cel(r,iU)), valor });
  }
  return { aba:nome, linhas:out };
}

function lerDiaristas(wb){
  const nome = acharAba(wb,"DIARISTAS","DIARISTA");
  if(!nome) return { erro:'Não encontrei a aba "DIARISTAS".' };
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[nome],{header:1,defval:""});
  const h = acharCabecalho(rows,["GROOT ID","NOME","DATA"]);
  if(!h) return { erro:'A aba "'+nome+'" não tem o cabeçalho esperado (GROOT ID, NOME, DATA).' };
  const c = h.cels;
  const iG=col(c,"GROOT ID"), iN=col(c,"NOME"), iC=col(c,"CARGO"), iE=col(c,"ESCALA"),
        iD=col(c,"DATA"), iQ=col(c,"QUANTIDADE"), iU=col(c,"VALOR UNITARIO"), iV=col(c,"VALOR FINAL");
  const out = [];
  for(let i=h.idx+1;i<rows.length;i++){
    const r = rows[i]; if(!r) continue;
    const nomeP = String(cel(r,iN) ?? "").trim();
    const groot = String(cel(r,iG) ?? "").trim();
    const data = lerData(cel(r,iD));
    const valor = num(cel(r,iV));
    if(!nomeP && !groot && !data && !valor) continue;
    out.push({ aba:"DIARISTAS", linha:i+1, groot, nome:nomeP,
      cargo:String(cel(r,iC) ?? "").trim(), escala:String(cel(r,iE) ?? "").trim(),
      data, qtd:num(cel(r,iQ)), unit:num(cel(r,iU)), valor });
  }
  return { aba:nome, linhas:out };
}

function lerHoraExtra(wb){
  const nome = acharAba(wb,"HORA EXTRA","HORAS EXTRAS");
  if(!nome) return { linhas:[] };
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[nome],{header:1,defval:""});
  const h = acharCabecalho(rows,["GROOT ID","NOME"]);
  if(!h) return { linhas:[] };
  const c = h.cels;
  const iG=col(c,"GROOT ID"), iN=col(c,"NOME"), iC=col(c,"CARGO"),
        iQ=col(c,"QUANTIDADE"), iU=col(c,"VALOR UNITARIO"), iV=col(c,"VALOR FINAL");
  const out = [];
  for(let i=h.idx+1;i<rows.length;i++){
    const r = rows[i]; if(!r) continue;
    const nomeP = String(cel(r,iN) ?? "").trim();
    if(!nomeP && !String(cel(r,iG) ?? "").trim()) continue;
    const q = num(cel(r,iQ)), u = num(cel(r,iU));
    out.push({ aba:"HORA EXTRA", linha:i+1, groot:String(cel(r,iG) ?? "").trim(), nome:nomeP,
      cargo:String(cel(r,iC) ?? "").trim(), qtd:q, unit:u,
      valor: num(cel(r,iV)) ?? (q !== null && u !== null ? q*u : null) });
  }
  return { aba:nome, linhas:out };
}

/* O período faturado. A competência vem do RESUMO (ou do campo PERIODO
   das próprias abas) e o template fatura o ciclo 16 do mês anterior →
   15 do mês da competência. Quando nada disso aparece, cai para o
   intervalo observado nos dados — assim a regra de "data fora do
   período" nunca dispara por falta de informação. */
function lerPeriodo(wb, labor, diaristas){
  let comp = null;
  const resumo = acharAba(wb,"RESUMO");
  if(resumo){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[resumo],{header:1,defval:""});
    for(const r of rows.slice(0,20)){
      for(const v of r){
        const s = String(v ?? "").trim();
        if(/^\d{6}$/.test(s)){
          const y = +s.slice(0,4), m = +s.slice(4);
          if(y > 2000 && y < 2100 && m >= 1 && m <= 12){ comp = { y, m }; break; }
        }
      }
      if(comp) break;
    }
  }
  if(comp){
    const ini = new Date(Date.UTC(comp.m === 1 ? comp.y-1 : comp.y, comp.m === 1 ? 11 : comp.m-2, 16));
    const fim = new Date(Date.UTC(comp.y, comp.m-1, 15));
    return { ini, fim, comp, origem:"competência "+String(comp.m).padStart(2,"0")+"/"+comp.y };
  }
  const datas = [...diaristas.map(d=>d.data), ...labor.map(l=>l.ini), ...labor.map(l=>l.fim)]
    .filter(d => d instanceof Date).sort((a,b)=>a-b);
  if(!datas.length) return { ini:null, fim:null, origem:"não identificado" };
  return { ini:datas[0], fim:datas[datas.length-1], origem:"intervalo observado nos dados" };
}

/* ================================================================
   FLUXO
   ================================================================ */
function analisar(file){
  const err = $("vt-err");
  err.classList.add("hidden");
  $("vt-fname").textContent = file.name;
  const rd = new FileReader();
  rd.onload = e => {
    try{
      const wb = XLSX.read(e.target.result,{type:"array",cellDates:true});
      const L = lerLabor(wb), D = lerDiaristas(wb);
      if(L.erro) return falhar(L.erro);
      if(D.erro) return falhar(D.erro);
      const HE = lerHoraExtra(wb);
      const periodo = lerPeriodo(wb, L.linhas, D.linhas);
      V.arquivo = file.name;
      V.marcas = {};
      V.filtro = "todos";
      V.auditoria = auditarFatura({ labor:L.linhas, diaristas:D.linhas, horaExtra:HE.linhas, periodo });
      V.auditoria.periodo.origem = periodo.origem;
      render();
    }catch(ex){ falhar("Não consegui ler o arquivo: "+ex.message); }
  };
  rd.readAsArrayBuffer(file);
}
function falhar(msg){
  const err = $("vt-err");
  err.textContent = msg;
  err.classList.remove("hidden");
  $("vt-result").classList.add("hidden");
}

/* ================================================================
   TELA
   ================================================================ */
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const fmtD = d => d instanceof Date
  ? String(d.getUTCDate()).padStart(2,"0")+"/"+String(d.getUTCMonth()+1).padStart(2,"0")+"/"+d.getUTCFullYear()
  : "—";
const fmtM = v => (v === null || v === undefined || !isFinite(Number(v)))
  ? "—" : Number(v).toLocaleString("pt-BR",{style:"currency",currency:"BRL"});

const SEV_LABEL = { critico:"Crítico", revisar:"Revisar", cadastro:"Cadastro", info:"Informativo" };

function render(){
  const a = V.auditoria;
  if(!a) return;
  $("vt-drop").classList.add("compacto");
  $("vt-result").classList.remove("hidden");

  const p = a.periodo;
  $("vt-cabecalho").innerHTML =
    '<div class="vt-ident">'
    + '<div><span>Arquivo</span><b>'+esc(V.arquivo)+'</b></div>'
    + '<div><span>Período faturado</span><b>'+fmtD(p.ini)+" a "+fmtD(p.fim)+'</b>'
    + '<i>'+esc(p.origem || "")+'</i></div>'
    + '<div><span>LABOR</span><b>'+a.totais.labor+' linhas</b><i>'+fmtM(a.totais.valorLabor)+'</i></div>'
    + '<div><span>DIARISTAS</span><b>'+a.totais.diaristas+' linhas</b><i>'+fmtM(a.totais.valorDiaristas)+'</i></div>'
    + '</div>';

  const total = a.achados.length;
  const cards = [
    ["todos","Todos",total,"vt-c-todos"],
    ["critico","Crítico",a.resumo.critico,"vt-c-critico"],
    ["revisar","Revisar",a.resumo.revisar,"vt-c-revisar"],
    ["cadastro","Cadastro",a.resumo.cadastro,"vt-c-cadastro"],
    ["info","Informativo",a.resumo.info,"vt-c-info"]
  ];
  $("vt-cards").innerHTML = cards.map(([k,l,n,cls]) =>
    '<button class="vt-card '+cls+(V.filtro===k?" ativo":"")+'" data-filtro="'+k+'">'
    + '<div class="v">'+n+'</div><div class="l">'+l+'</div></button>').join("");
  [...$("vt-cards").querySelectorAll(".vt-card")].forEach(b =>
    b.onclick = () => { V.filtro = b.dataset.filtro; render(); });

  if(!total){
    $("vt-lista").innerHTML =
      '<div class="vt-vazio"><b>Nenhuma inconsistência encontrada</b>'
      + '<p>As abas LABOR e DIARISTAS passaram por todas as verificações sem apontamento.</p></div>';
    return;
  }

  const decididos = a.achados.filter(x => V.marcas[x.id]).length;
  $("vt-progresso").textContent = total + " apontamento(s) · " + decididos
    + " com decisão registrada · " + (total-decididos) + " pendente(s)";

  const lista = a.achados.filter(x => V.filtro === "todos" || x.severidade === V.filtro);
  $("vt-lista").innerHTML = lista.map(cardAchado).join("")
    || '<div class="vt-vazio"><b>Nada nesta categoria</b></div>';

  [...$("vt-lista").querySelectorAll(".vt-achado header")].forEach(h =>
    h.onclick = () => h.parentElement.classList.toggle("aberto"));
  [...$("vt-lista").querySelectorAll(".vt-acao")].forEach(b =>
    b.onclick = ev => {
      ev.stopPropagation();
      const id = b.dataset.id, marca = b.dataset.marca;
      V.marcas[id] = V.marcas[id] === marca ? null : marca;
      render();
    });
}

function cardAchado(x){
  const marca = V.marcas[x.id] || null;
  const meta = [];
  if(x.groot) meta.push(["GROOT", x.groot]);
  if(x.cargo) meta.push(["Cargo", x.cargo]);
  if(x.aba) meta.push(["Aba", x.aba]);
  if(x.datas) meta.push(["Datas", x.datas]);
  if(x.valor !== null && x.valor !== undefined) meta.push(["Valor", fmtM(x.valor)]);

  /* Toda opção é uma decisão registrável — inclusive "Manter lançamento",
     que também é uma escolha. Clicar de novo na mesma volta a Pendente. */
  const acoes = (x.opcoes || []).map(o =>
    '<button class="vt-acao'+(marca===o?" ativa":"")+'" data-id="'+esc(x.id)
      + '" data-marca="'+esc(o)+'">'+esc(o)+'</button>').join("");

  return '<article class="vt-achado sev-'+x.severidade+(marca?" marcado":"")+'">'
    + '<header>'
    +   '<span class="vt-chip sev-'+x.severidade+'">'+SEV_LABEL[x.severidade]+'</span>'
    +   '<div class="vt-tit"><b>'+esc(x.titulo)+'</b>'
    +     '<small>'+esc(x.regra)+(x.nome?" · "+esc(x.nome):"")+'</small></div>'
    +   (marca ? '<span class="vt-marca">'+esc(marca)+'</span>' : "")
    +   '<span class="vt-seta">▾</span>'
    + '</header>'
    + '<div class="vt-corpo">'
    +   '<div class="vt-meta">'+meta.map(([k,v]) =>
          '<div><span>'+k+'</span><b>'+esc(v)+'</b></div>').join("")+'</div>'
    +   '<div class="vt-razao"><span>Por que foi sinalizado</span><p>'+esc(x.raciocinio)+'</p></div>'
    +   (x.sugestao ? '<div class="vt-sug"><span>Sugestão</span><p>'+esc(x.sugestao)+'</p></div>' : "")
    +   (x.registros && x.registros.length ? tabelaRegistros(x.registros) : "")
    +   '<div class="vt-acoes">'+acoes+'</div>'
    + '</div></article>';
}

function tabelaRegistros(regs){
  const linhas = regs.map(r =>
    '<tr><td>'+esc(r.aba)+'<small> L'+r.linha+'</small></td>'
    + '<td>'+esc(r.groot || "—")+'</td>'
    + '<td>'+esc(r.nome || "—")+'</td>'
    + '<td>'+esc(r.cargo || "—")+(r.regime ? '<small> · '+esc(r.regime)+'</small>' : "")+'</td>'
    + '<td>'+(r.data instanceof Date ? fmtD(r.data) : fmtD(r.ini)+" → "+fmtD(r.fim))+'</td>'
    + '<td class="'+(Number(r.valor) < 0 ? "neg" : "")+'">'+fmtM(r.valor)+'</td></tr>').join("");
  return '<div class="vt-regs"><table>'
    + '<thead><tr><th>Origem</th><th>GROOT</th><th>Nome</th><th>Cargo</th><th>Período</th><th>Valor</th></tr></thead>'
    + '<tbody>'+linhas+'</tbody></table></div>';
}

/* ================================================================
   RELATÓRIO
   ================================================================ */
function exportar(){
  const a = V.auditoria;
  if(!a) return;
  const aoa = [["Severidade","Regra","Título","GROOT","Nome","Cargo","Aba","Datas","Valor",
                "Por que foi sinalizado","Sugestão","Decisão"]];
  for(const x of a.achados){
    aoa.push([ SEV_LABEL[x.severidade], x.regra, x.titulo, x.groot, x.nome, x.cargo, x.aba, x.datas,
      x.valor === null ? "" : Number(x.valor), x.raciocinio, x.sugestao,
      V.marcas[x.id] || "Pendente" ]);
  }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{wch:12},{wch:26},{wch:40},{wch:11},{wch:34},{wch:24},{wch:16},{wch:30},
                 {wch:14},{wch:120},{wch:80},{wch:12}];
  XLSX.utils.book_append_sheet(wb, ws, "Achados");

  const resumo = [["Arquivo", V.arquivo],
    ["Período", fmtD(a.periodo.ini)+" a "+fmtD(a.periodo.fim)],
    ["Origem do período", a.periodo.origem || ""],
    ["Linhas LABOR", a.totais.labor], ["Valor LABOR", a.totais.valorLabor],
    ["Linhas DIARISTAS", a.totais.diaristas], ["Valor DIARISTAS", a.totais.valorDiaristas],
    [], ["Crítico", a.resumo.critico], ["Revisar", a.resumo.revisar],
    ["Cadastro", a.resumo.cadastro], ["Informativo", a.resumo.info]];
  const wsR = XLSX.utils.aoa_to_sheet(resumo);
  wsR["!cols"] = [{wch:22},{wch:44}];
  XLSX.utils.book_append_sheet(wb, wsR, "Resumo");

  /* Âncora explícita em vez de XLSX.writeFile, e nome SEM ACENTO: num
     link para blob: URL, um acento no atributo `download` faz o Chromium
     descartar o nome inteiro e salvar como "download", em silêncio. */
  const base = String(V.arquivo).replace(/\.xlsx?$/i,"");
  const nome = ("Validação - "+base+".xlsx")
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[\\/:*?"<>|]/g,"-");
  const buf = XLSX.write(wb,{ bookType:"xlsx", type:"array" });
  const url = URL.createObjectURL(new Blob([buf],
    { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const link = document.createElement("a");
  link.href = url; link.download = nome;
  link.click();
  URL.revokeObjectURL(url);
}

/* ================================================================
   LIGAÇÕES
   ================================================================ */
const drop = $("vt-drop"), input = $("vt-file");
if(drop && input){
  drop.onclick = () => input.click();
  ["dragover","dragenter"].forEach(e => drop.addEventListener(e, ev => {
    ev.preventDefault(); drop.classList.add("over"); }));
  ["dragleave","drop"].forEach(e => drop.addEventListener(e, ev => {
    ev.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", ev => { if(ev.dataTransfer.files[0]) analisar(ev.dataTransfer.files[0]); });
  input.onchange = () => { if(input.files[0]) analisar(input.files[0]); };
}
const btn = $("vt-btnExport");
if(btn) btn.onclick = exportar;

})();
