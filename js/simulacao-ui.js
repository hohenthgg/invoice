/* Retorno Simulado — leitura dos dois arquivos e tela
   ================================================================

   Lê a aba LABOR da fatura e os blocos de S&OP da planilha
   operacional, chama js/simulacao.js e desenha. O raciocínio mora lá.

   A parte delicada da leitura é o S&OP: cada aba operacional traz uma
   linha "Esperado" com uma coluna por dia, e os dias vêm como número
   solto ("16 (Qui)"). Somar SVC com SD PELA POSIÇÃO da coluna é o erro
   fácil e silencioso — basta uma aba ter uma coluna a mais no começo
   para somar o dia 16 de uma com o dia 17 da outra. Por isso cada
   coluna é resolvida para uma DATA COMPLETA antes de qualquer soma, e
   a soma é feita por data.
   ================================================================ */
"use strict";
(function(){

const $ = id => document.getElementById(id);
const norm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"")
  .toUpperCase().replace(/\s+/g," ").trim();

const S = { fatura:null, sop:null, sim:null, nomeF:"", nomeS:"" };

/* ================================================================
   MODOS DA ABA
   ================================================================ */
const modos = $("vt-modos");
if(modos){
  [...modos.querySelectorAll(".vt-modo")].forEach(b => b.onclick = () => {
    [...modos.querySelectorAll(".vt-modo")].forEach(x => x.classList.toggle("ativo", x === b));
    document.querySelectorAll("#panel-validacao .vt-painel").forEach(p =>
      p.classList.toggle("ativo", p.id === "vt-painel-" + b.dataset.modo));
  });
}

/* ================================================================
   LEITURA — FATURA
   ================================================================ */
function acharAba(wb, ...alvos){
  for(const alvo of alvos){
    const a = norm(alvo);
    const n = wb.SheetNames.find(x => norm(x) === a) || wb.SheetNames.find(x => norm(x).includes(a));
    if(n) return n;
  }
  return null;
}
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

function lerFatura(wb){
  const nome = acharAba(wb,"LABOR");
  if(!nome) return { erro:'Não encontrei a aba "LABOR" na fatura.' };
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[nome],{header:1,defval:""});
  const h = acharCabecalho(rows,["GROOT ID","CARGO","DATA DE INICIO"]);
  if(!h) return { erro:'A aba "'+nome+'" não tem o cabeçalho esperado (GROOT ID, CARGO, DATA DE INÍCIO).' };
  const c = h.cels;
  const iG=col(c,"GROOT ID"), iN=col(c,"NOME"), iC=col(c,"CARGO"), iCt=col(c,"DESCRICAO CONTA"),
        iR=col(c,"REGIME DE CONTRATO"), iI=col(c,"DATA DE INICIO"), iF=col(c,"DATA FIM"),
        iRa=col(c,"% RATEIO","RATEIO"), iP=col(c,"PERIODO");
  const faltando = [];
  if(iI < 0) faltando.push("DATA DE INÍCIO");
  if(iF < 0) faltando.push("DATA FIM");
  if(iG < 0) faltando.push("GROOT ID");
  if(iC < 0) faltando.push("CARGO");
  if(iRa < 0) faltando.push("% RATEIO");
  if(faltando.length) return { erro:'A aba "'+nome+'" não tem a(s) coluna(s): '+faltando.join(", ")+"." };

  const linhas = [];
  let comp = null;
  for(let i=h.idx+1;i<rows.length;i++){
    const r = rows[i]; if(!r) continue;
    const nomeP = String(r[iN] ?? "").trim(), groot = String(r[iG] ?? "").trim();
    if(!nomeP && !groot) continue;
    if(!comp && iP >= 0){
      const s = String(r[iP] ?? "").trim();
      if(/^\d{6}$/.test(s)) comp = { y:+s.slice(0,4), m:+s.slice(4) };
    }
    linhas.push({ linha:i+1, groot, nome:nomeP, cargo:String(r[iC] ?? "").trim(),
      conta: iCt >= 0 ? String(r[iCt] ?? "").trim() : "",
      regime: iR >= 0 ? String(r[iR] ?? "").trim() : "",
      inicio: parseExcelDate(r[iI]), fim: parseExcelDate(r[iF]),
      rateio: parseRateio(r[iRa]) });
  }
  if(!linhas.length) return { erro:'A aba "'+nome+'" não tem linhas com dado.' };
  return { aba:nome, linhas, comp };
}

/* ================================================================
   LEITURA — S&OP

   Um bloco por aba operacional. A linha de cabeçalho é a que traz os
   dias; a linha "Esperado" logo abaixo traz o S&OP de cada um.
   ================================================================ */
const SUFIXOS_OP = ["SVC","FULL","XD","SD"];

/* Rótulo curto da aba: "VGSVC 16.07-15.08" → "SVC". Sem sufixo
   conhecido, fica o prefixo alfabético inteiro. */
function rotuloDaAba(nomeAba){
  const m = /^([A-Za-z]+)/.exec(nomeAba.trim());
  if(!m) return nomeAba.trim();
  const cod = m[1].toUpperCase();
  for(const suf of SUFIXOS_OP){
    if(cod.endsWith(suf) && cod.length > suf.length) return suf;
  }
  return cod;
}

/* "· 16/07 a 15/08/2026" no título, ou o intervalo no nome da aba. */
function intervaloDoTexto(txt){
  const s = String(txt ?? "");
  let m = /(\d{1,2})[\/.](\d{1,2})\s*(?:a|até|-|–)\s*(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})/i.exec(s);
  if(m){
    const d1=+m[1], m1=+m[2], d2=+m[3], m2=+m[4], y=+m[5];
    const y1 = m1 > m2 ? y-1 : y;
    return { ini:ymd(y1,m1,d1), fim:ymd(y,m2,d2) };
  }
  m = /(\d{1,2})[\/.](\d{1,2})\s*[-–]\s*(\d{1,2})[\/.](\d{1,2})/.exec(s);
  if(m) return { parcial:{ d1:+m[1], m1:+m[2], d2:+m[3], m2:+m[4] } };
  return null;
}

function lerSop(wb){
  const blocos = [], recusadas = [];
  let intervalo = null;

  for(const nomeAba of wb.SheetNames){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[nomeAba],{header:1,defval:""});
    if(!rows.length) continue;

    /* linha "Esperado" e a linha de dias imediatamente acima dela */
    let iEsp = -1;
    for(let i=0;i<Math.min(rows.length,12);i++){
      if(norm(rows[i][0]) === "ESPERADO"){ iEsp = i; break; }
    }
    if(iEsp < 1) continue;

    let iHdr = -1, dias = null;
    for(let i=iEsp-1;i>=0 && i>=iEsp-3;i--){
      const achados = [];
      for(let c=0;c<(rows[i]||[]).length;c++){
        const m = /^(\d{1,2})\s*[\(\n]/.exec(String(rows[i][c] ?? "").trim());
        if(m) achados.push({ col:c, dia:+m[1] });
      }
      if(achados.length >= 20){ iHdr = i; dias = achados; break; }
    }
    if(iHdr < 0){ recusadas.push(nomeAba+" (não achei a linha de dias)"); continue; }

    /* O intervalo real vem do título da aba; sem ele, do nome dela. */
    const inter = intervaloDoTexto(rows[0][0]) || intervaloDoTexto(nomeAba);
    if(!inter || !inter.ini){ recusadas.push(nomeAba+" (não consegui ler o período)"); continue; }
    if(!intervalo) intervalo = inter;

    /* Resolve cada coluna para uma DATA COMPLETA: os dias correm do 16
       ao fim do mês e recomeçam no 1 do mês seguinte. A virada é onde o
       número do dia cai. */
    const p0 = ymdParts(inter.ini), p1 = ymdParts(inter.fim);
    let y = p0.y, mes = p0.m, anterior = 0;
    const mapa = {};
    let ok = true;
    for(const { col:c, dia } of dias){
      if(dia < anterior){
        if(mes === p1.m && y === p1.y){ ok = false; break; }   // virou duas vezes: layout estranho
        mes = p1.m; y = p1.y;
      }
      anterior = dia;
      const data = ymd(y,mes,dia);
      if(!isValidYmd(data)){ ok = false; break; }
      const v = rows[iEsp][c];
      const n = Number(v);
      mapa[data] = (v === "" || v === null || !isFinite(n)) ? null : n;
    }
    if(!ok){ recusadas.push(nomeAba+" (sequência de dias inconsistente)"); continue; }

    blocos.push({ rotulo:rotuloDaAba(nomeAba), aba:nomeAba, dias:mapa,
      periodo:{ ini:inter.ini, fim:inter.fim } });
  }

  if(!blocos.length){
    return { erro:"Não encontrei nenhuma aba com a linha “Esperado” e uma linha de dias."
      + (recusadas.length ? " Abas descartadas: "+recusadas.join("; ")+"." : "") };
  }
  return { blocos, intervalo, recusadas };
}

/* ================================================================
   FLUXO
   ================================================================ */
function bindDrop(idDrop, idFile, fn){
  const drop = $(idDrop), file = $(idFile);
  if(!drop || !file) return;
  drop.onclick = () => file.click();
  ["dragover","dragenter"].forEach(e => drop.addEventListener(e, ev => {
    ev.preventDefault(); drop.classList.add("over"); }));
  ["dragleave","drop"].forEach(e => drop.addEventListener(e, ev => {
    ev.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", ev => { if(ev.dataTransfer.files[0]) fn(ev.dataTransfer.files[0], drop); });
  file.onchange = () => { if(file.files[0]) fn(file.files[0], drop); };
}

function carregar(qual, file, drop){
  const st = $(qual === "fatura" ? "sm-stF" : "sm-stS");
  $(qual === "fatura" ? "sm-fnF" : "sm-fnS").textContent = file.name;
  st.textContent = "Lendo…"; st.className = "st";
  drop.classList.remove("loaded","err");
  const rd = new FileReader();
  rd.onload = e => {
    try{
      const wb = XLSX.read(e.target.result,{type:"array",cellDates:true});
      const r = qual === "fatura" ? lerFatura(wb) : lerSop(wb);
      if(r.erro){ st.textContent = "✗ "+r.erro; st.className = "st bad"; drop.classList.add("err");
                  S[qual === "fatura" ? "fatura" : "sop"] = null; return pronto(); }
      if(qual === "fatura"){
        S.fatura = r; S.nomeF = file.name;
        st.textContent = "✓ aba "+r.aba+" · "+r.linhas.length+" linha(s)"
          + (r.comp ? " · competência "+String(r.comp.m).padStart(2,"0")+"/"+r.comp.y : "");
      } else {
        S.sop = r; S.nomeS = file.name;
        st.textContent = "✓ "+r.blocos.length+" bloco(s): "+r.blocos.map(b=>b.rotulo).join(" + ")
          + (r.recusadas.length ? " · "+r.recusadas.length+" aba(s) ignorada(s)" : "");
      }
      st.className = "st ok"; drop.classList.add("loaded");
      pronto();
    }catch(ex){ st.textContent = "✗ "+ex.message; st.className = "st bad"; drop.classList.add("err"); }
  };
  rd.readAsArrayBuffer(file);
}

function pronto(){
  const ok = !!(S.fatura && S.sop);
  $("sm-btnRun").disabled = !ok;
  $("sm-msgRun").textContent = ok ? "Pronto para analisar."
    : (!S.fatura && !S.sop) ? "Carregue os dois arquivos."
    : !S.fatura ? "Falta a fatura." : "Falta a planilha operacional.";
}

function analisar(){
  const err = $("sm-err");
  err.classList.add("hidden");
  const blocos = S.sop.blocos;

  /* Validações de compatibilidade ANTES de somar qualquer coisa. */
  const problemas = [];
  const per = blocos[0].periodo;
  for(const b of blocos.slice(1)){
    if(b.periodo.ini !== per.ini || b.periodo.fim !== per.fim){
      problemas.push("as abas "+blocos[0].rotulo+" e "+b.rotulo+" cobrem períodos diferentes ("
        + fmtYmd(per.ini)+"–"+fmtYmd(per.fim)+" contra "+fmtYmd(b.periodo.ini)+"–"+fmtYmd(b.periodo.fim)+")");
    }
    const dA = Object.keys(blocos[0].dias).sort(), dB = Object.keys(b.dias).sort();
    if(dA.join() !== dB.join()){
      problemas.push("as datas de "+blocos[0].rotulo+" e "+b.rotulo+" não coincidem — "
        + "somar as duas por posição de coluna produziria dias trocados");
    }
  }
  if(problemas.length){
    err.innerHTML = "<b>Não dá para somar o S&amp;OP com segurança:</b><br>" + problemas.join(";<br>") + ".";
    err.classList.remove("hidden");
    $("sm-result").classList.add("hidden");
    return;
  }

  /* A competência é a da fatura; sem ela, deduz-se do fim do período. */
  const comp = S.fatura.comp
    ? buildCompetence(S.fatura.comp.y, S.fatura.comp.m)
    : buildCompetence(ymdParts(per.fim).y, ymdParts(per.fim).m);

  const sim = simularRetorno({ labor:S.fatura.linhas, blocos, periodo:per, comp });
  if(sim.erro){ err.textContent = sim.erro; err.classList.remove("hidden"); return; }
  S.sim = sim;
  render();
}

/* ================================================================
   TELA
   ================================================================ */
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const n2 = v => v === null || v === undefined || !isFinite(Number(v))
  ? "—" : Number(Number(v).toFixed(2)).toLocaleString("pt-BR");
const sinal = v => v === null || v === undefined ? "—" : (Number(v) > 0 ? "+" : "") + n2(v);

const COR = { pref:"var(--s-labor)", sop:"var(--s-alvo)", pos:"var(--ok)" };

function render(){
  const sim = S.sim, t = sim.totais;
  $("sm-result").classList.remove("hidden");

  $("sm-avisos").innerHTML = sim.avisos.map(a =>
    '<div class="sm-aviso '+a.tipo+'">'+esc(a.texto)+'</div>').join("");

  const cards = [
    ["PREF total enviado", n2(t.pref), "HC-dia", ""],
    ["S&OP total", n2(t.cliente), "HC-dia", ""],
    ["Q Pós total previsto", n2(t.pos), "HC-dia", "ok"],
    ["HC-dia em risco de correção", n2(t.hcEmRisco), "acima do S&OP", t.hcEmRisco > 0 ? "bad" : "ok"],
    ["HC-dia abaixo do S&OP", n2(t.hcAbaixo), "possível subfaturamento", t.hcAbaixo > 0 ? "warn" : "ok"],
    ["Dias alinhados", t.alinhado, "de "+t.dias, "ok"],
    ["Dias com possível redução", t.reducao, "PREF acima do S&OP", t.reducao ? "bad" : "ok"],
    ["Dias com possível subfaturamento", t.subfaturamento, "PREF abaixo do S&OP", t.subfaturamento ? "warn" : "ok"]
  ];
  if(t.revisao) cards.push(["Dias para revisão", t.revisao, "não reconstruídos", "warn"]);
  $("sm-cards").innerHTML = cards.map(([l,v,s,cls]) =>
    '<div class="sm-card'+(cls?" "+cls:"")+'"><div class="v">'+esc(String(v))+'</div>'
    + '<div class="l">'+esc(l)+'</div><div class="s">'+esc(s)+'</div></div>').join("");

  $("sm-legenda").innerHTML =
    '<span><i style="background:'+COR.pref+'"></i>PREF enviado</span>'
    + '<span><i style="background:'+COR.sop+'"></i>S&amp;OP total ('+sim.blocos.map(esc).join(" + ")+')</span>'
    + '<span><i style="background:'+COR.pos+'"></i>Q Pós previsto</span>';

  desenharGrafico(sim.dias);
  desenharDesvios(sim.dias);
  desenharTabela(sim);
}

/* Gráfico de linhas simples, em SVG inline. */
function desenharGrafico(dias){
  const el = $("sm-chart");
  const pts = dias.filter(d => d.qCliente !== null);
  if(!pts.length){ el.innerHTML = '<p class="sm-msg">Nenhum dia pôde ser reconstruído.</p>'; return; }
  const W=880, H=250, padL=34, padR=16, padT=12, padB=30, n=pts.length;
  const serie = k => pts.map(d => Number(d[k]));
  const v1=serie("pref"), v2=serie("qCliente"), v3=serie("qPos");
  const max = Math.max(1, ...v1, ...v2, ...v3), min = Math.min(0, ...v1, ...v2, ...v3);
  const x = i => n <= 1 ? padL : padL + (i/(n-1))*(W-padL-padR);
  const y = v => H-padB - ((v-min)/(max-min || 1))*(H-padT-padB);
  const grade = [], ticks = [];
  for(let s=0;s<=4;s++){
    const v = min + (max-min)*s/4, yy = y(v);
    grade.push('<line x1="'+padL+'" x2="'+(W-padR)+'" y1="'+yy.toFixed(1)+'" y2="'+yy.toFixed(1)
      + '" stroke="var(--line)" stroke-width="1"/>');
    grade.push('<text x="'+(padL-6)+'" y="'+(yy+3).toFixed(1)+'" text-anchor="end" font-size="9" '
      + 'fill="var(--mut2)" font-family="var(--mono)">'+Math.round(v)+'</text>');
  }
  const passo = Math.max(1, Math.round(n/8));
  for(let i=0;i<n;i+=passo) ticks.push('<text x="'+x(i).toFixed(1)+'" y="'+(H-8)
    + '" text-anchor="middle" font-size="9" fill="var(--mut2)" font-family="var(--mono)">'
    + fmtShort(pts[i].data)+'</text>');
  const linha = (vals,cor,largura) => '<path d="'+vals.map((v,i)=>(i?"L":"M")+x(i).toFixed(1)+","+y(v).toFixed(1)).join(" ")
    + '" fill="none" stroke="'+cor+'" stroke-width="'+largura+'" stroke-linejoin="round" stroke-linecap="round"/>';
  /* Área entre PREF e S&OP só onde o PREF está por cima: é o risco. */
  const risco = [];
  pts.forEach((d,i) => { if(d.gap > 0) risco.push('<line x1="'+x(i).toFixed(1)+'" x2="'+x(i).toFixed(1)
    + '" y1="'+y(d.qCliente).toFixed(1)+'" y2="'+y(d.pref).toFixed(1)+'" stroke="var(--bad)" stroke-width="3" opacity=".35"/>'); });

  el.innerHTML = '<svg viewBox="0 0 '+W+' '+H+'" role="img" aria-label="PREF, S&OP e Q Pós previsto por dia">'
    + grade.join("") + risco.join("")
    + linha(v2, COR.sop, 2) + linha(v3, COR.pos, 3) + linha(v1, COR.pref, 2)
    + ticks.join("") + '</svg>';
}

function desenharDesvios(dias){
  const grupos = [
    ["reducao","Provável redução", dias.filter(d => d.status === "reducao").sort((a,b)=>b.gap-a.gap)],
    ["subfaturamento","Possível subfaturamento", dias.filter(d => d.status === "subfaturamento").sort((a,b)=>a.gap-b.gap)],
    ["revisao","Revisão necessária", dias.filter(d => d.status === "revisao")],
    ["alinhado","Alinhado", dias.filter(d => d.status === "alinhado")]
  ];
  $("sm-desvios").innerHTML = grupos.filter(g => g[2].length).map(([k,titulo,ds]) =>
    '<div class="sm-grupo st-'+k+'">'
    + '<header><b>'+esc(titulo)+'</b><span>'+ds.length+' dia(s)</span></header>'
    + '<div class="sm-chips">'+ds.map(d =>
        '<span class="sm-chipdia" title="'+esc(d.diagnostico)+'">'+fmtShort(d.data)
        + (d.gap === null ? '' : '<b>'+sinal(d.gap)+'</b>')+'</span>').join("")
    + '</div></div>').join("");
}

function desenharTabela(sim){
  const cabBlocos = sim.blocos.map(b => "<th>S&amp;OP "+esc(b)+"</th>").join("");
  const linhas = sim.dias.map(d =>
    '<tr class="st-'+d.status+'">'
    + '<td>'+fmtYmd(d.data)+'</td>'
    + d.blocos.map(b => '<td>'+n2(b.valor)+'</td>').join("")
    + '<td class="forte">'+n2(d.qCliente)+'</td>'
    + '<td class="forte">'+n2(d.pref)+'</td>'
    + '<td class="forte">'+n2(d.qPos)+'</td>'
    + '<td class="'+(d.gap > 0 ? "pos" : d.gap < 0 ? "neg" : "")+'">'+sinal(d.gap)+'</td>'
    + '<td class="'+(d.correcao < 0 ? "neg" : "")+'">'+sinal(d.correcao)+'</td>'
    + '<td><span class="sm-st st-'+d.status+'">'+esc(SIM_STATUS_LABEL[d.status])+'</span></td>'
    + '<td class="diag">'+esc(d.diagnostico)+'</td></tr>').join("");
  $("sm-tabela").innerHTML = '<table><thead><tr><th>Data</th>'+cabBlocos
    + '<th>Q cliente / S&amp;OP</th><th>PREF enviado</th><th>Q Pós previsto</th>'
    + '<th>Gap PREF × S&amp;OP</th><th>Correção prevista</th><th>Status</th><th>Diagnóstico</th>'
    + '</tr></thead><tbody>'+linhas+'</tbody></table>';
}

/* ================================================================
   EXPORTAÇÃO
   ================================================================ */
function exportar(){
  const sim = S.sim;
  if(!sim) return;
  const dt = v => v === null ? "" : ymdToExcelDate(v);
  const wb = XLSX.utils.book_new();

  /* --- aba 1: o comparativo, para leitura humana --- */
  const cab = ["Data", ...sim.blocos.map(b => "S&OP "+b), "Q cliente / S&OP", "PREF enviado",
               "Q Pós previsto", "Gap PREF x S&OP", "Correção prevista", "Status", "Diagnóstico"];
  const aoa = [cab];
  for(const d of sim.dias){
    aoa.push([ dt(d.data), ...d.blocos.map(b => b.valor), d.qCliente, d.pref, d.qPos,
      d.gap, d.correcao, SIM_STATUS_LABEL[d.status], d.diagnostico ]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  formatarDatas(ws, aoa.length, 0);
  ws["!cols"] = [{wch:12}, ...sim.blocos.map(()=>({wch:11})), {wch:17},{wch:14},{wch:16},
                 {wch:16},{wch:17},{wch:24},{wch:110}];
  XLSX.utils.book_append_sheet(wb, ws, "COMPARATIVO");

  /* --- aba 2: o retorno simulado, no formato que a Fusão de Linhas lê ---
     A Fusão localiza esta aba pelo CABEÇALHO, não pelo nome, então o nome
     pode dizer com todas as letras que o arquivo é simulado. */
  const retAoa = [["Data Trab.","Employee_Type","Qtd. PREF","Q Meli (Pós Comp. Diaristas)",
                   "Desvio","Ocorrencia"]];
  for(const l of simLinhasRetorno(sim)){
    retAoa.push([ dt(l.data), l.tipo, l.pref, l.qPos, l.desvio, l.ocorrencia ]);
  }
  const wsR = XLSX.utils.aoa_to_sheet(retAoa);
  formatarDatas(wsR, retAoa.length, 0);
  wsR["!cols"] = [{wch:12},{wch:14},{wch:11},{wch:28},{wch:10},{wch:110}];
  XLSX.utils.book_append_sheet(wb, wsR, "RETORNO SIMULADO");

  /* --- aba 3: metadados, para ninguém confundir com o oficial --- */
  const meta = [
    ["AVISO", SIM_AVISO_METADADOS],
    [],
    ["Fatura de origem", S.nomeF],
    ["Planilha operacional", S.nomeS],
    ["Período", fmtYmd(sim.periodo.ini)+" a "+fmtYmd(sim.periodo.fim)],
    ["Competência", sim.comp ? sim.comp.label : ""],
    ["Blocos de S&OP somados", sim.blocos.join(" + ")],
    [],
    ["Fórmula — Q cliente", "soma do Esperado de cada bloco no dia"],
    ["Fórmula — Q Pós previsto", "MIN(PREF, Q cliente)"],
    ["Fórmula — Gap", "PREF - Q cliente"],
    ["Fórmula — Correção prevista", "Q Pós previsto - PREF"],
    [],
    ["Linhas do LABOR no PREF", sim.linhas.dentro],
    ["Linhas fora do PREF", sim.linhas.fora.length],
    [],
    ["PREF total (HC-dia)", sim.totais.pref],
    ["S&OP total (HC-dia)", sim.totais.cliente],
    ["Q Pós total previsto (HC-dia)", sim.totais.pos],
    ["HC-dia em risco de correção", sim.totais.hcEmRisco],
    ["HC-dia abaixo do S&OP", sim.totais.hcAbaixo]
  ];
  for(const a of sim.avisos) meta.push(["Aviso", a.texto]);
  const wsM = XLSX.utils.aoa_to_sheet(meta);
  wsM["!cols"] = [{wch:32},{wch:120}];
  XLSX.utils.book_append_sheet(wb, wsM, "METADADOS");

  /* --- aba 4: o que ficou fora do PREF, e por quê --- */
  const foraAoa = [["Linha","GROOT","Nome","Cargo","Conta","Início","Fim","Rateio","Motivo","Detalhe"]];
  for(const f of sim.linhas.fora){
    const l = f.linha;
    foraAoa.push([ l.linha, l.groot, l.nome, l.cargo, l.conta,
      isValidYmd(l.inicio)?dt(l.inicio):"", isValidYmd(l.fim)?dt(l.fim):"", l.rateio,
      f.motivo.texto, f.motivo.detalhe || "" ]);
  }
  const wsF = XLSX.utils.aoa_to_sheet(foraAoa);
  formatarDatas(wsF, foraAoa.length, 5); formatarDatas(wsF, foraAoa.length, 6);
  wsF["!cols"] = [{wch:8},{wch:11},{wch:34},{wch:24},{wch:26},{wch:12},{wch:12},{wch:8},{wch:56},{wch:100}];
  XLSX.utils.book_append_sheet(wb, wsF, "FORA DO PREF");

  const base = String(S.nomeF).replace(/\.xlsx?$/i,"");
  /* Sem acento: num link para blob: URL, acento no atributo `download`
     faz o Chromium descartar o nome inteiro. */
  const nome = ("Retorno Simulado - "+base+".xlsx")
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[\\/:*?"<>|]/g,"-");
  const buf = XLSX.write(wb,{ bookType:"xlsx", type:"array" });
  const url = URL.createObjectURL(new Blob([buf],
    { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const link = document.createElement("a");
  link.href = url; link.download = nome; link.click();
  URL.revokeObjectURL(url);
}

function formatarDatas(ws, linhas, coluna){
  for(let r=1;r<linhas;r++){
    const ref = XLSX.utils.encode_cell({ r, c:coluna });
    const cel = ws[ref];
    if(cel && cel.v instanceof Date){ cel.t = "d"; cel.z = "dd/mm/yyyy"; }
  }
}

/* ================================================================
   LIGAÇÕES
   ================================================================ */
bindDrop("sm-dropF","sm-fileF",(f,d) => carregar("fatura",f,d));
bindDrop("sm-dropS","sm-fileS",(f,d) => carregar("sop",f,d));
const btnRun = $("sm-btnRun"); if(btnRun) btnRun.onclick = analisar;
const btnExp = $("sm-btnExport"); if(btnExp) btnExp.onclick = exportar;

})();
