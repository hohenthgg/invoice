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

const S = { fatura:null, sop:null, diar:null, sim:null,
            nomeF:"", nomeS:"", nomeD:"", fonte:"planilha" };

/* De onde vem o S&OP: da planilha operacional, dia a dia por operação,
   ou de um valor fixo que vale para todos os dias do período. O fixo
   dispensa a planilha — é o número que o contrato fecha para o mês. */
function fonteSop(){
  const m = document.querySelector('input[name="sm-fonte"]:checked');
  return m ? m.value : "planilha";
}
function sopFixo(){
  const el = $("sm-sopFixo");
  const n = el ? Number(el.value) : NaN;
  return isFinite(n) && n >= 0 ? n : null;
}
/* Um bloco único, com o mesmo valor em cada dia do período. */
function blocoFixo(periodo, valor){
  const dias = {};
  for(let d = periodo.ini; d <= periodo.fim; d = addDays(d,1)) dias[d] = valor;
  return [{ rotulo:"S&OP fixo", aba:"(informado na tela)", dias, periodo }];
}

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
        iRa=col(c,"% RATEIO","RATEIO"), iP=col(c,"PERIODO"), iMt=col(c,"MATRICULA"),
        iDf=col(c,"DIAS TRABALHADOS X FOLGA"), iEs=col(c,"ESCALA");
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
      matricula: iMt >= 0 ? String(r[iMt] ?? "").trim() : "",
      inicio: parseExcelDate(r[iI]), fim: parseExcelDate(r[iF]),
      rateio: parseRateio(r[iRa]),
      diasFolga: iDf >= 0 ? String(r[iDf] ?? "").trim() : "",
      escala: iEs >= 0 ? String(r[iEs] ?? "").trim() : "",
      /* A linha crua e a posição das colunas ficam guardadas porque a
         exportação devolve o LABOR no layout do próprio template — as
         colunas que o app não lê têm de sair como entraram. */
      bruta: r.slice() });
  }
  if(!linhas.length) return { erro:'A aba "'+nome+'" não tem linhas com dado.' };
  return { aba:nome, linhas, comp, unidade: unidadeDaFatura(wb),
    diarias: lerDiariasDaFatura(wb),
    layout:{ topo: rows.slice(0, h.idx+1), largura: Math.max(h.cels.length,
      ...linhas.map(l => l.bruta.length)),
      cols:{ groot:iG, nome:iN, cargo:iC, conta:iCt, regime:iR, inicio:iI, fim:iF,
             rateio:iRa, matricula:iMt, obs:col(c,"OBSERVACOES","OBSERVAÇÕES"),
             fixas:["FORNECEDOR","PERIODO","PERIODO - MES/ANO","UNIDADE","DESCRICAO CONTA",
                    "CONTA CONTABIL","UNIDADE + UF","EMPRESA"].map(x => col(c,x)) } } };
}

/* ================================================================
   AS DIÁRIAS QUE A FATURA JÁ LANÇA

   Elas vivem na aba DIARISTAS, não no LABOR, e ignorá-las foi o defeito
   que fez o app pedir 23 diaristas num dia em que a fatura já pagava 15
   — com 29 pessoa-dia repetidas entre as duas listas, que é cobrança
   dobrada. Uma vaga ocupada por diária é uma vaga ocupada.
   ================================================================ */
function lerDiariasDaFatura(wb){
  const nome = acharAba(wb,"DIARISTAS","DIARISTA");
  if(!nome) return [];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[nome],{header:1,defval:""});
  const h = acharCabecalho(rows,["GROOT ID","DATA"]);
  if(!h) return [];
  const c = h.cels;
  const iG=col(c,"GROOT ID"), iN=col(c,"NOME"), iC=col(c,"CARGO"), iE=col(c,"ESCALA"),
        iD=col(c,"DATA"), iQ=col(c,"QUANTIDADE");
  if(iG < 0 || iD < 0) return [];
  const out = [];
  for(let i=h.idx+1;i<rows.length;i++){
    const r = rows[i]; if(!r) continue;
    const groot = String(r[iG] ?? "").trim();
    const data = parseExcelDate(r[iD]);
    if(!groot || !isValidYmd(data)) continue;
    const q = iQ >= 0 ? Number(r[iQ]) : 1;
    out.push({ groot, nome:String(r[iN] ?? "").trim(),
      cargo: iC >= 0 ? String(r[iC] ?? "").trim() : "",
      escala: iE >= 0 ? String(r[iE] ?? "").trim() : "",
      data, quantidade: isFinite(q) && q !== 0 ? q : 1 });
  }
  return out;
}

/* O nome da unidade, do RESUMO — é ele que escolhe a aba do SIGO. */
function unidadeDaFatura(wb){
  const resumo = acharAba(wb,"RESUMO");
  if(!resumo) return "";
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[resumo],{header:1,defval:""});
  for(const r of rows.slice(0,20)){
    for(let c=0;c<r.length-1;c++){
      if(norm(r[c]) === "NOME UNIDADE"){
        const v = String(r[c+1] ?? "").trim();
        if(v) return v;
      }
    }
  }
  return "";
}

/* ================================================================
   LEITURA — DIARISTAS (SIGO)

   Uma aba por filial, cabeçalho na linha 4: DATA SOLICITAÇÃO,
   SOLICITANTE, GROOT ID. A aba usada é a que casa com a unidade da
   fatura — somar as filiais todas daria um número que não é desta
   fatura.
   ================================================================ */
function lerDiaristasSigo(wb, unidade){
  const candidatas = [];
  for(const nomeAba of wb.SheetNames){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[nomeAba],{header:1,defval:""});
    const h = acharCabecalho(rows,["GROOT ID","DATA","SOLICITANTE"]);
    if(h) candidatas.push({ aba:nomeAba, rows, h });
  }
  if(!candidatas.length){
    return { erro:"Nenhuma aba no formato SIGO (cabeçalho com DATA SOLICITAÇÃO, SOLICITANTE e GROOT ID)." };
  }
  const alvo = norm(unidade);
  const escolhida = alvo
    ? candidatas.find(c => norm(c.aba) === alvo) || candidatas.find(c => norm(c.aba).includes(alvo))
      || candidatas.find(c => alvo.includes(norm(c.aba)))
    : null;
  if(!escolhida){
    return { erro:'Não achei no SIGO uma aba para a unidade "'+(unidade || "(não identificada na fatura)")
      + '". Abas disponíveis: '+candidatas.map(c => c.aba).join(", ")+"." };
  }

  const c = escolhida.h.cels;
  const iD=col(c,"DATA SOLICITACAO","DATA"), iS=col(c,"SOLICITANTE"), iG=col(c,"GROOT ID"),
        iN=col(c,"NOME"), iE=col(c,"EMPRESA");
  const regs = [];
  for(let i=escolhida.h.idx+1;i<escolhida.rows.length;i++){
    const r = escolhida.rows[i]; if(!r) continue;
    const data = parseExcelDate(r[iD]);
    const groot = String(r[iG] ?? "").replace(/\D/g,"");
    if(!isValidYmd(data) || !groot) continue;
    /* Quem PEDIU decide de quem é o diarista; a agência não importa. */
    const txt = norm(r[iS]);
    const solic = !txt ? "" : (txt.includes("MELI") ? "meli" : "id");
    regs.push({ data, groot, solic, nome:String(r[iN] ?? "").trim(),
                empresa: iE >= 0 ? String(r[iE] ?? "").trim() : "" });
  }
  if(!regs.length) return { erro:'A aba "'+escolhida.aba+'" não tem linha com data e GROOT legíveis.' };
  return { aba:escolhida.aba, regs, abas:candidatas.map(c => c.aba) };
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
  /* Zerar o value depois de usar: sem isso, escolher o MESMO arquivo de
     novo não dispara `change`, e quem corrigiu a ordem de carga fica
     clicando sem resposta. */
  file.onchange = () => { const f = file.files[0]; file.value = ""; if(f) fn(f, drop); };
}

const CAMPOS = { fatura:["sm-stF","sm-fnF"], sop:["sm-stS","sm-fnS"], diar:["sm-stD","sm-fnD"] };

function carregar(qual, file, drop){
  const [idSt, idFn] = CAMPOS[qual];
  const st = $(idSt);
  $(idFn).textContent = file.name;
  st.textContent = "Lendo…"; st.className = "st";
  drop.classList.remove("loaded","err");
  const rd = new FileReader();
  rd.onload = e => {
    try{
      const wb = XLSX.read(e.target.result,{type:"array",cellDates:true});
      const r = qual === "fatura" ? lerFatura(wb)
              : qual === "sop"    ? lerSop(wb)
                                  : lerDiaristasSigo(wb, S.fatura ? S.fatura.unidade : "");
      if(r.erro){ st.textContent = "✗ "+r.erro; st.className = "st bad"; drop.classList.add("err");
                  S[qual] = null; return pronto(); }
      if(qual === "fatura"){
        S.fatura = r; S.nomeF = file.name;
        st.textContent = "✓ aba "+r.aba+" · "+r.linhas.length+" linha(s)"
          + (r.comp ? " · competência "+String(r.comp.m).padStart(2,"0")+"/"+r.comp.y : "")
          + (r.unidade ? " · unidade "+r.unidade : "");
        /* chegou a fatura: o SIGO que estava esperando pode ser lido */
        if(S.diarPendente && S.dropDiar){
          const pend = S.diarPendente; S.diarPendente = null;
          carregar("diar", pend, S.dropDiar);
        }
      } else if(qual === "sop"){
        S.sop = r; S.nomeS = file.name;
        st.textContent = "✓ "+r.blocos.length+" bloco(s): "+r.blocos.map(b=>b.rotulo).join(" + ")
          + (r.recusadas.length ? " · "+r.recusadas.length+" aba(s) ignorada(s)" : "");
      } else {
        S.diar = r; S.nomeD = file.name;
        const nId = r.regs.filter(x => x.solic === "id").length;
        const nMeli = r.regs.filter(x => x.solic === "meli").length;
        const nSem = r.regs.length - nId - nMeli;
        st.textContent = "✓ aba "+r.aba+" · "+r.regs.length+" solicitação(ões) · "
          + nId+" ID · "+nMeli+" cliente"+(nSem ? " · "+nSem+" sem solicitante" : "");
      }
      st.className = "st ok"; drop.classList.add("loaded");
      pronto();
    }catch(ex){ st.textContent = "✗ "+ex.message; st.className = "st bad"; drop.classList.add("err"); }
  };
  rd.readAsArrayBuffer(file);
}

function pronto(){
  const fixo = fonteSop() === "fixo";
  const painel = $("vt-painel-simulacao");
  if(painel) painel.classList.toggle("fonte-fixo", fixo);

  const ok = !!S.fatura && (fixo ? sopFixo() !== null : !!S.sop);
  $("sm-btnRun").disabled = !ok;
  $("sm-msgRun").textContent =
      ok               ? "Pronto para analisar."
    : !S.fatura        ? "Falta a fatura."
    : fixo             ? "Informe o HC por dia."
                       : "Falta a planilha operacional.";
}

function analisar(){
  const err = $("sm-err");
  err.classList.add("hidden");
  const fixo = fonteSop() === "fixo";

  /* Com S&OP fixo não há planilha para conferir: o período vem da
     própria fatura, e o valor informado vale para todos os dias. */
  if(fixo){
    const per = periodoDaFatura();
    if(!per){ return falhar(err, "Não consegui deduzir o período da fatura para aplicar o S&OP fixo."); }
    const v = sopFixo();
    return rodar(blocosFixoDoPeriodo(per, v), per, "fixo", v);
  }

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
    return falhar(err, "<b>Não dá para somar o S&amp;OP com segurança:</b><br>"
      + problemas.join(";<br>") + ".", true);
  }
  rodar(blocos, per, "planilha", null);
}

function falhar(err, msg, html){
  if(html) err.innerHTML = msg; else err.textContent = msg;
  err.classList.remove("hidden");
  $("sm-result").classList.add("hidden");
}

/* O período faturado a partir da competência da própria fatura: o ciclo
   16 do mês anterior → 15 do mês da competência. */
function periodoDaFatura(){
  const c = S.fatura && S.fatura.comp;
  if(!c) return null;
  const ini = c.m === 1 ? ymd(c.y-1, 12, 16) : ymd(c.y, c.m-1, 16);
  return { ini, fim: ymd(c.y, c.m, 15) };
}
const blocosFixoDoPeriodo = (per, valor) => blocoFixo(per, valor);

function rodar(blocos, per, fonte, valorFixo){
  /* A competência é a da fatura; sem ela, deduz-se do fim do período. */
  const comp = S.fatura.comp
    ? buildCompetence(S.fatura.comp.y, S.fatura.comp.m)
    : buildCompetence(ymdParts(per.fim).y, ymdParts(per.fim).m);

  const diaristas = S.diar
    ? S.diar.regs.filter(x => x.data >= per.ini && x.data <= per.fim)
    : [];
  const sim = simularRetorno({ labor:S.fatura.linhas, blocos, periodo:per, comp, diaristas });
  if(sim.erro) return falhar($("sm-err"), sim.erro);
  sim.fonte = fonte;
  sim.valorFixo = valorFixo;
  S.fonte = fonte;
  S.sim = sim;
  render();
}

/* ================================================================
   AJUDA NO HOVER

   Cada número da tela é composto de um jeito, e o jeito não é óbvio:
   "Q Pós previsto" não é o S&OP, "HC-dia em risco" não desconta os
   dias abaixo, e o PREF não é a contagem de linhas do LABOR. Em vez de
   um parágrafo de legenda que ninguém lê, a composição fica atrás do
   cursor, no próprio número.

   O balão é posicionado por JS, e não por CSS, porque a tabela vive
   dentro de um contêiner com `overflow-x: auto` — um ::after seria
   recortado na borda do scroll.
   ================================================================ */
function tipPref(){
  return "Soma dos <b>% RATEIO</b> das linhas do LABOR ativas no dia — ativa é "
    + "<b>início ≤ dia</b> e <b>fim ≥ dia</b>, com DATA FIM vazia valendo até o fim do período. "
    + "Entram só as linhas de <b>LABOR DIRETO</b>, com cargo da lista do PREF e classificadas como "
    + "competência corrente. Ficam de fora liderança e indiretos, a linha de ABS e todo retroativo.";
}
function tipSop(blocos){
  return "Linha <b>Esperado</b> de cada aba de operação da planilha operacional, somada "
    + "<b>por data</b> — nunca por posição de coluna. Neste arquivo: "
    + blocos.map(b => "<b>"+esc(b)+"</b>").join(" + ")+".";
}
const TIP = {
  qPos: "<b>MIN(PREF, S&OP)</b> do dia. É previsão conservadora e assimétrica de propósito: "
      + "o cliente pode <b>cortar</b> o que foi enviado acima do plano, mas não paga o que não foi "
      + "enviado. Com PREF abaixo do S&OP, o previsto é o próprio PREF — não o S&OP.",
  gap:  "<b>PREF − S&OP</b> do dia. Positivo: enviado acima do plano, é o que pode ser cortado. "
      + "Negativo: enviado abaixo do plano, indício de pessoa faturável faltando no Labor.",
  corr: "<b>Q Pós previsto − PREF</b>. Fica negativa quando há risco de corte e <b>zero</b> quando "
      + "o PREF está abaixo do S&OP — o app não presume aumento automático.",
  data: "Cada dia do período faturado, do dia 16 do mês anterior ao dia 15 do mês da competência.",
  status: "Comparação entre PREF e S&OP do dia: acima → possível correção, igual → alinhado, "
      + "abaixo → possível subfaturamento. <b>Revisão necessária</b> quando o dia não pôde ser "
      + "reconstruído e nenhum número foi previsto.",
  diag: "A leitura do dia em texto, já com os números que a produziram.",
  totalPref: "Soma do PREF de <b>todos os dias</b> do período, em HC-dia.",
  totalSop:  "Soma do S&OP de <b>todos os dias</b> do período, em HC-dia.",
  totalPos:  "Soma de <b>MIN(PREF, S&OP)</b> dia a dia. Não é o menor dos dois totais: é a soma "
           + "dos menores de cada dia, que pode ser diferente.",
  risco: "Soma de <b>(PREF − S&OP)</b> apenas nos dias em que o PREF está <b>acima</b>. Os dias "
       + "abaixo <b>não abatem</b> este total — o que sobra num dia não compensa o que faltou "
       + "em outro.",
  abaixo: "Soma de <b>(S&OP − PREF)</b> apenas nos dias em que o PREF está <b>abaixo</b>. "
        + "Não vira receita: é indício de gente faturável faltando no Labor.",
  diasAlinhados: "Dias em que PREF e S&OP são iguais — nenhuma correção prevista.",
  diasReducao: "Dias com PREF acima do S&OP.",
  diasSub: "Dias com PREF abaixo do S&OP.",
  diaristas: "Diaristas <b>solicitados no SIGO</b> naquele dia, contados uma vez por pessoa. "
      + "Não entram os que <b>já constam no LABOR do dia</b> — esses já estão sendo cobrados como "
      + "quadro fixo, e contá-los de novo seria contar duas vezes. Quem tem estorno no LABOR "
      + "cobrindo o dia <b>conta</b>: o fixo foi devolvido justamente para pagar a diária. "
      + "O detalhe separa quem foi pedido pela <b>ID</b> de quem foi pedido pelo <b>cliente</b>.",
  totalDiaristas: "Soma dos diaristas disponíveis em todos os dias, em pessoa-dia.",
  abate: "Quanto da falta os diaristas disponíveis dariam para cobrir: em cada dia, o menor entre "
       + "os disponíveis e a própria falta. Um dia nunca fica positivo por sobra de diarista.",
  diasRevisao: "Dias que não puderam ser reconstruídos — S&OP sem valor para o dia, ou linha do "
             + "Labor que não pôde ser classificada. Saem sem número previsto, de propósito."
};

/* Um único balão para o painel inteiro, movido conforme o cursor. */
function ligarAjuda(){
  const tip = $("sm-tip"), painel = $("vt-painel-simulacao");
  if(!tip || !painel || painel.dataset.ajuda === "1") return;
  painel.dataset.ajuda = "1";
  painel.addEventListener("mouseover", ev => {
    const alvo = ev.target.closest("[data-tip]");
    if(!alvo) return;
    tip.innerHTML = alvo.dataset.tip;
    tip.hidden = false;
    posicionar(alvo, tip, painel);
  });
  painel.addEventListener("mouseout", ev => {
    const alvo = ev.target.closest("[data-tip]");
    if(alvo && !alvo.contains(ev.relatedTarget)) tip.hidden = true;
  });
  painel.addEventListener("scroll", () => { tip.hidden = true; }, true);
}
function posicionar(alvo, tip, painel){
  const a = alvo.getBoundingClientRect(), p = painel.getBoundingClientRect();
  const largura = tip.offsetWidth, altura = tip.offsetHeight;
  let x = a.left - p.left + a.width/2 - largura/2;
  x = Math.max(6, Math.min(x, painel.clientWidth - largura - 6));
  /* Acima do alvo quando cabe, abaixo quando não. O "cabe" é medido na
     JANELA e não no painel: o painel é alto e rola, então um alvo no
     topo da tela tem espaço de sobra em coordenadas do painel e mesmo
     assim jogaria o balão para fora do campo de visão. Cabeçalho de
     tabela é sticky e vive colado no topo — é justamente o caso. */
  const cabeAcima = a.top - altura - 8 > 4;
  const y = cabeAcima ? a.top - p.top - altura - 8 : a.bottom - p.top + 8;
  tip.style.left = x + "px";
  tip.style.top  = y + "px";
}

/* ================================================================
   TELA
   ================================================================ */
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const n2 = v => v === null || v === undefined || !isFinite(Number(v))
  ? "—" : Number(Number(v).toFixed(2)).toLocaleString("pt-BR");
const sinal = v => v === null || v === undefined ? "—" : (Number(v) > 0 ? "+" : "") + n2(v);

function render(){
  const sim = S.sim, t = sim.totais;
  $("sm-result").classList.remove("hidden");

  const selo = $("sm-result").querySelector(".sm-selo");
  if(selo){
    selo.innerHTML = 'Retorno <b>previsto</b> — estimado por MIN(PREF, S&amp;OP). Não é o retorno '
      + 'oficial do cliente. S&amp;OP '
      + (sim.fonte === "fixo"
          ? 'de <b>valor fixo</b>: '+n2(sim.valorFixo)+' HC em todos os dias do período.'
          : 'da <b>planilha operacional</b>, dia a dia ('+sim.blocos.map(esc).join(" + ")+').');
  }

  $("sm-avisos").innerHTML = sim.avisos.map(a =>
    '<div class="sm-aviso '+a.tipo+'">'+esc(a.texto)+'</div>').join("");

  const cards = [
    ["PREF total enviado", n2(t.pref), "HC-dia", "", TIP.totalPref+" "+tipPref()],
    ["S&OP total", n2(t.cliente), "HC-dia", "", TIP.totalSop+" "+tipSop(sim.blocos)],
    ["Q Pós total previsto", n2(t.pos), "HC-dia", "ok", TIP.totalPos+" "+TIP.qPos],
    ["HC-dia em risco de correção", n2(t.hcEmRisco), "acima do S&OP", t.hcEmRisco > 0 ? "bad" : "ok", TIP.risco],
    ["HC-dia abaixo do S&OP", n2(t.hcAbaixo), "possível subfaturamento", t.hcAbaixo > 0 ? "warn" : "ok", TIP.abaixo],
    ["Dias alinhados", t.alinhado, "de "+t.dias, "ok", TIP.diasAlinhados],
    ["Dias com possível redução", t.reducao, "PREF acima do S&OP", t.reducao ? "bad" : "ok", TIP.diasReducao],
    ["Dias com possível subfaturamento", t.subfaturamento, "PREF abaixo do S&OP", t.subfaturamento ? "warn" : "ok", TIP.diasSub]
  ];
  if(t.revisao) cards.push(["Dias para revisão", t.revisao, "não reconstruídos", "warn", TIP.diasRevisao]);
  if(t.diaristasDisp !== undefined){
    cards.push(["Diaristas disponíveis", n2(t.diaristasDisp),
      "pessoa-dia" + (t.diaristasOcupados ? " · "+t.diaristasOcupados+" já no LABOR" : ""),
      "", TIP.totalDiaristas]);
    cards.push(["Falta que os diaristas cobririam", n2(t.abatePossivel), "HC-dia",
      t.abatePossivel > 0 ? "ok" : "", TIP.abate]);
  }
  $("sm-cards").innerHTML = cards.map(([l,v,s,cls,tip]) =>
    '<div class="sm-card'+(cls?" "+cls:"")+'" data-tip="'+esc(tip)+'"><div class="v">'+esc(String(v))+'</div>'
    + '<div class="l">'+esc(l)+'</div><div class="s">'+esc(s)+'</div></div>').join("");

  desenharDesvios(sim.dias);
  desenharTabela(sim);
  desenharEqualizacao();
  ligarAjuda();
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

/* Total do dia, com a origem embaixo. Ganha destaque só quando falta
   gente — é aí que a disponibilidade responde alguma coisa. */
function celulaDiaristas(d){
  const c = d.diaristas;
  if(!c) return '<td>—</td>';
  const util = d.gap !== null && d.gap < 0;
  const partes = [];
  if(c.dispId)   partes.push(c.dispId+" ID");
  if(c.dispMeli) partes.push(c.dispMeli+" cliente");
  if(c.dispSem)  partes.push(c.dispSem+" s/ solic.");
  return '<td class="sm-diar'+(util && c.disp ? " util" : "")+'">'
    + '<b>'+n2(c.disp)+'</b>'
    + (partes.length ? '<small>'+esc(partes.join(" · "))+'</small>' : "")
    + (c.ocupados ? '<small class="ocup">'+c.ocupados+' já no LABOR</small>' : "")
    + '</td>';
}

function desenharTabela(sim){
  const th = (rotulo, tip) => '<th data-tip="'+esc(tip)+'">'+rotulo+'</th>';
  /* A coluna por operação só existe quando há mais de uma: com um bloco
     só — o caso do S&OP fixo — ela repetiria o total logo ao lado. */
  const detalhar = sim.blocos.length > 1;
  const temDiar = sim.dias.some(d => d.diaristas);
  const cabBlocos = detalhar ? sim.blocos.map(b => th(esc(b),
    "Linha <b>Esperado</b> da aba de <b>"+esc(b)+"</b>, na coluna deste dia. "
    + "A coluna é resolvida para a data completa antes de entrar na soma.")).join("") : "";
  const linhas = sim.dias.map(d =>
    '<tr class="st-'+d.status+'">'
    + '<td>'+fmtYmd(d.data)+'</td>'
    + (detalhar ? d.blocos.map(b => '<td>'+n2(b.valor)+'</td>').join("") : "")
    + '<td class="forte">'+n2(d.qCliente)+'</td>'
    + '<td class="forte">'+n2(d.pref)+'</td>'
    + '<td class="forte">'+n2(d.qPos)+'</td>'
    + '<td class="'+(d.gap > 0 ? "pos" : d.gap < 0 ? "neg" : "")+'">'+sinal(d.gap)+'</td>'
    + (temDiar ? celulaDiaristas(d) : "")
    + '<td class="'+(d.correcao < 0 ? "neg" : "")+'">'+sinal(d.correcao)+'</td>'
    + '<td><span class="sm-st st-'+d.status+'">'+esc(SIM_STATUS_LABEL[d.status])+'</span></td>'
    + '<td class="diag">'+esc(d.diagnostico)+'</td></tr>').join("");
  /* Com o valor fixo, a coluna é o QF do cliente — chamá-la de S&OP
     seria dizer que veio da planilha operacional, e não veio. */
  $("sm-tabela").innerHTML = '<table><thead><tr>'+th("Data", TIP.data)+cabBlocos
    + th(sim.fonte === "fixo" ? "QF" : "S&amp;OP", tipSop(sim.blocos))
    + th("PREF", tipPref())
    + th("Q Pós", TIP.qPos)
    + th("Gap", TIP.gap)
    + (temDiar ? th("Diaristas disp.", TIP.diaristas) : "")
    + th("Correção", TIP.corr)
    + th("Status", TIP.status)
    + th("Diagnóstico", TIP.diag)
    + '</tr></thead><tbody>'+linhas+'</tbody></table>';
}

/* ================================================================
   EQUALIZAÇÃO — PREF × QF DO CLIENTE

   Esta seção não calcula nada. Ela chama simPlanoEqualizacao(), que
   chama o motor de js/equalizacao.js — o MESMO que a Fusão de Linhas
   usa para montar o Labor ajustado. Aqui o plano é SUGESTÃO: nada é
   aplicado à fatura, e é por isso que a seção existe separada do
   resto da tela.

   Ela só aparece quando o alvo é o QF do cliente (fonte "valor fixo"),
   porque é contra ele que a equalização faz sentido — o S&OP diário
   responde outra pergunta.
   ================================================================ */
const EQ_ROTULO = {
  retirar:  "Retirar linha",
  adiar:    "Adiar início",
  encurtar: "Antecipar fim",
  pausar:   "Pausar e retomar"
};
const EQ_ORDEM = ["retirar","adiar","encurtar","pausar"];

function opcoesEq(){
  const c = $("sm-eqAdiar"), p = $("sm-eqPausa");
  let pausaDesde = null;
  if(p && p.value){ const [y,m,d] = p.value.split("-").map(Number); pausaDesde = ymd(y,m,d); }
  return { permitirAdiarInicio: c ? c.checked : true, pausaDesde };
}

const faixaTxt = f => f.de === f.ate ? fmtShort(f.de) : fmtShort(f.de)+"–"+fmtShort(f.ate);

/* O que muda na linha, em uma frase. É o que o card mostra fechado. */
function eqMudanca(a){
  if(a.tipo === "retirar")  return "linha sai do Labor";
  if(a.tipo === "adiar")    return "início "+fmtShort(a.linha.inicio)+" → "+fmtShort(a.novoInicio);
  if(a.tipo === "encurtar") return "fim "+(a.linha.fim ? fmtShort(a.linha.fim) : "em aberto")
    + " → "+fmtShort(a.novoFim);
  return a.pausas.map(p => "pausa "+fmtShort(p.fim)+" a "+fmtShort(p.ini)).join(" · ");
}

function eqDetalhe(a){
  const partes = ["<p>"+esc(a.motivo)+"</p>"];
  partes.push('<p><b>Impacto na curva:</b> '+sinal(-a.impacto.hc)+' HC em '
    + a.impacto.dias + ' dia(s) — ' + esc(a.impacto.faixas.map(faixaTxt).join(", ")) + '.</p>');
  partes.push('<p><b>Vigência atual:</b> '+fmtYmd(a.linha.inicio)+' → '
    + (a.linha.fim ? fmtYmd(a.linha.fim) : "em aberto")
    + ' · <b>% rateio</b> '+n2(a.linha.rateio)
    + (a.linha.matricula ? ' · <b>Matrícula</b> '+esc(a.linha.matricula) : "")
    + ' · <b>linha</b> '+a.linha.linha+' do LABOR.</p>');
  if(a.diarista && a.diarista.dias.length){
    partes.push('<p class="sm-eqdiar"><b>Também aparece como diarista</b> em '
      + a.diarista.dias.length + ' dos dias que o plano tira do fixo ('
      + esc(a.diarista.dias.map(fmtShort).join(", ")) + '). No período inteiro são '
      + a.diarista.total + ' diária(s) — ' + a.diarista.id + ' ID, ' + a.diarista.meli + ' cliente. '
      + 'Isso costuma ser transição de fixo para diária, e explica a correção: confirme se não há '
      + 'dupla cobrança nesses dias.</p>');
  } else if(a.diarista){
    partes.push('<p class="sm-eqdiar">Aparece na base de diaristas em '+a.diarista.total
      + ' dia(s) do período, mas nenhum deles coincide com os dias que o plano tira do fixo.</p>');
  }
  return partes.join("");
}

function desenharEqualizacao(){
  const bloco = $("sm-eq-bloco");
  if(!bloco) return;
  const sim = S.sim;
  const alvo = (sim && sim.fonte === "fixo") ? sim.valorFixo : null;
  if(!sim || !alvo){ bloco.hidden = true; return; }
  bloco.hidden = false;

  const diaristas = S.diar
    ? S.diar.regs.filter(x => x.data >= sim.periodo.ini && x.data <= sim.periodo.fim) : [];
  const plano = simPlanoEqualizacao({ labor:S.fatura.linhas, periodo:sim.periodo,
    alvo, diaristas, diarias:S.fatura.diarias, opcoes:opcoesEq() });
  S.plano = plano;
  if(plano.erro){ $("sm-eq").innerHTML = '<div class="sm-aviso">'+esc(plano.erro)+'</div>'; return; }

  const t = plano.totais;
  const quadros = plano.dias.map(d => d.quadro);
  const item = (l,v) => '<div><span>'+l+'</span><b>'+v+'</b></div>';
  /* O quadro do dia é fixo + diária já lançada. Mostrar só o fixo
     esconderia a vaga que a diária ocupa, e foi assim que a falta saiu
     inflada. */
  const resumo = '<div class="sm-eqresumo">'
    + item("Período", fmtShort(plano.periodo.ini)+" → "+fmtShort(plano.periodo.fim))
    + item("QF do cliente", n2(alvo))
    + item("Quadro mínimo", n2(Math.min(...quadros)))
    + item("Quadro máximo", n2(Math.max(...quadros)))
    + (t.diarias ? item("Diárias já na fatura", n2(t.diarias)+" pessoa-dia") : "")
    + item("Excesso no período", n2(t.excessoAntes)+" HC-dia")
    + item("Excesso após o plano", n2(t.excessoDepois)+" HC-dia")
    + '</div>';


  if(!plano.acoes.length && !Object.keys(plano.revisar).length && !Object.keys(plano.falta).length){
    $("sm-eq").innerHTML = resumo + '<div class="sm-eqok"><b>Nada a equalizar.</b> '
      + 'O quadro já bate com o QF em todos os dias do período.</div>';
    return;
  }

  const grupos = EQ_ORDEM.map(tipo => {
    const itens = plano.acoes.filter(a => a.tipo === tipo);
    if(!itens.length) return "";
    return '<div class="sm-eqgrupo">'
      + '<header><span class="sm-eqchip '+tipo+'">'+EQ_ROTULO[tipo]+'</span>'
      + '<span class="n">'+itens.length+' pessoa(s)</span></header>'
      + itens.map(a =>
          '<div class="sm-eqitem" onclick="this.classList.toggle(\'aberto\')">'
          + '<div class="cab"><div class="nm">'+esc(a.linha.nome)
          + '<small>Groot '+esc(a.linha.groot || "—")+' · '+esc(a.linha.cargo)+'</small></div>'
          + '<div class="mud">'+esc(eqMudanca(a))
          + '<small>'+sinal(-a.impacto.hc)+' HC · '+a.impacto.dias+' dia(s)'
          + (a.diarista && a.diarista.dias.length ? ' · também diarista' : "")+'</small></div></div>'
          + '<div class="det">'+eqDetalhe(a)+'</div></div>').join("")
      + '</div>';
  }).join("");

  /* `falta` guarda o tamanho da falta como número positivo; na tela ela
     é uma queda em relação ao QF e sai com sinal negativo. */
  /* As pessoas que faltam, com nome — só existe com a base do SIGO. */
  const inc = plano.inclusoes;
  const grupoInc = (inc && inc.pessoas.length)
    ? '<div class="sm-eqgrupo"><header><span class="sm-eqchip incluir">Incluir</span>'
      + '<span class="n">'+inc.totais.pessoas+' pessoa(s) · '+inc.totais.incluido+' pessoa-dia</span></header>'
      + '<div class="det aberto"><p>Diaristas solicitados no SIGO e <b>sem cobrança naquele dia</b> '
      + '— nem no LABOR, nem como diária já lançada nesta fatura. Primeiro os da <b>ID</b>, até '
      + 'acabar; o do cliente só entra depois '
      + '('+inc.totais.id+' ID · '+inc.totais.meli+' cliente'
      + (inc.totais.sem ? ' · '+inc.totais.sem+' sem solicitante' : "")+'). '
      + (inc.totais.descoberto > 0
          ? 'Ainda ficam <b>'+n2(inc.totais.descoberto)+'</b> pessoa-dia sem cobertura — não havia '
            + 'diarista livre bastante no SIGO.'
          : 'A falta do período fica <b>inteiramente coberta</b>.')
      + ' Os <b>'+inc.totais.verificados+'</b> pessoa-dia escolhidos foram reconferidos contra a '
      + 'base antes de entrar'
      + (inc.totais.recusados ? ', e <b>'+inc.totais.recusados+'</b> foram recusados por não '
          + 'constar no dia' : ' — nenhum entrou num dia em que não foi solicitado')
      + '.</p></div>'
      + inc.pessoas.map(p =>
          '<div class="sm-eqitem" onclick="this.classList.toggle(\'aberto\')">'
          + '<div class="cab"><div class="nm">'+esc(p.nome || "(sem nome no SIGO)")
          + '<small>Groot '+esc(p.groot)+' · diarista '
          + (p.solic === "id" ? "ID" : p.solic === "meli" ? "do cliente" : "sem solicitante")+'</small></div>'
          + '<div class="mud">'+esc(p.faixas.map(faixaTxt).join(", "))
          + '<small>+'+p.total+' dia(s)</small></div></div>'
          + '<div class="det"><p>Sai na aba <b>Diaristas</b> do arquivo exportado, no layout da '
          + 'fatura: '+p.total+' linha(s), uma por dia, cargo <i>Diarista</i> e quantidade 1. '
          + 'Não entra no Labor — quem cobre a falta é diarista, não quadro fixo, e lançá-lo '
          + 'como fixo seria cobrar outra coisa.</p></div></div>').join("")
      + '</div>' : "";

  const listaDias = (o, neg) => Object.keys(o).map(Number).sort((a,b)=>a-b)
    .map(d => fmtShort(d)+" ("+sinal(neg ? -o[d] : o[d])+")").join(" · ");
  const revisar = Object.keys(plano.revisar).length
    ? '<div class="sm-eqgrupo revisar"><header><span class="sm-eqchip revisar">Revisar</span>'
      + '<span class="n">'+Object.keys(plano.revisar).length+' dia(s)</span></header>'
      + '<div class="det aberto"><p>Nestes dias sobra quadro que nenhuma das quatro ações resolve. '
      + 'Onde há diária já lançada na fatura, é dela que vem a sobra — o quadro fixo já está no '
      + 'teto e a equalização não mexe em diária que já aconteceu. Nos demais, zerar exigiria '
      + 'partir um contrato de um jeito que o motor não faz sozinho. Avalie caso a caso.</p>'
      + '<p class="dias">'+listaDias(plano.revisar,false)+'</p></div></div>' : "";
  const falta = (Object.keys(plano.falta).length && !grupoInc)
    ? '<div class="sm-eqgrupo falta"><header><span class="sm-eqchip falta">Falta</span>'
      + '<span class="n">'+Object.keys(plano.falta).length+' dia(s)</span></header>'
      + '<div class="det aberto"><p>Nestes dias o Labor está <b>abaixo</b> do QF. A equalização não '
      + 'mexe neles — inventar pessoa não é trabalho dela. Verifique se falta gente no arquivo.</p>'
      + '<p class="dias">'+listaDias(plano.falta,true)+'</p></div></div>' : "";

  $("sm-eq").innerHTML = resumo
    + '<div class="sm-eqselo">Este plano equaliza <b>matematicamente</b> o Labor ao QF. '
    + 'Confirme se corresponde à movimentação operacional real antes de aplicar — '
    + 'nada é alterado na fatura por aqui.</div>'
    + grupos + grupoInc + revisar + falta
    + '<div class="vt-acoes-fim"><button class="vt-btn" id="sm-btnExportEq">'
    + 'Exportar fatura equalizada (.xlsx)</button></div>';
  const btn = $("sm-btnExportEq");
  if(btn) btn.onclick = exportarEqualizado;
}

/* ================================================================
   EXPORTAÇÃO DO LABOR EQUALIZADO

   O arquivo sai no LAYOUT DA FATURA — as duas abas que o cliente lê,
   `Labor` e `Diaristas`, com as colunas, as larguras, o cabeçalho
   amarelo e as bordas do modelo. Escrito com ExcelJS, porque o SheetJS
   da leitura não carrega estilo.

   `Labor`      o quadro fixo depois do plano: as retiradas saem, os
                inícios e fins ajustados já vêm mudados, e uma pausa
                vira duas linhas da mesma pessoa.
   `Diaristas`  as pessoas escolhidas para cobrir a falta, uma linha
                por pessoa-dia, como a aba de diaristas da fatura faz.

   Por que a inclusão vai para `Diaristas` e não para o `Labor`: quem
   entra para cobrir a falta é diarista — foi solicitado no SIGO para
   aquele dia, não contratado para o quadro. Lançá-lo no Labor seria
   cobrá-lo como fixo, que é outra coisa. A consequência aparece na
   curva e está dita no METADADOS: o Labor exportado fica ABAIXO do QF
   nos dias em que a falta foi coberta por diária, e é assim mesmo.

   As quatro abas de documentação continuam, porque o arquivo tem de
   explicar o que fez: EQUALIZACAO, INCLUSOES, REVISAR e METADADOS.
   ================================================================ */
const bordaFina = cor => ({ top:{style:"thin",color:cor}, left:{style:"thin",color:cor},
                            bottom:{style:"thin",color:cor}, right:{style:"thin",color:cor} });

/* Copiado célula a célula do modelo de fatura, valores inclusive — as
   larguras são as fracionárias do arquivo, não arredondadas, senão a
   coluna sai com um fio de diferença da original. */
const TPL = {
  fonte:      { name:"Calibri", family:2, size:9, color:{ theme:1 } },
  fonteCab:   { name:"Calibri", family:2, size:9, color:{ theme:1 }, bold:true },
  fillCab:    { type:"pattern", pattern:"solid",
                fgColor:{ argb:"FFFFFF00" }, bgColor:{ argb:"FFFFFF00" } },
  fillBranco: { type:"pattern", pattern:"solid", fgColor:{ theme:0 }, bgColor:{ theme:0 } },
  bordaClara: bordaFina({ argb:"FFF3F3F3" }),
  /* O cabeçalho do Labor no modelo não tem borda inferior — quem fecha
     a linha é a borda superior do corpo. Copiado como está. */
  bordaCabSemFundo: (() => { const b = bordaFina({ argb:"FFF3F3F3" }); delete b.bottom; return b; })(),
  bordaCinza: bordaFina({ theme:2, tint:-0.0499893185216834 }),
  centro:     { horizontal:"center", vertical:"middle" },
  centroWrap: { horizontal:"center", vertical:"middle", wrapText:true }
};
const TPL_LABOR = {
  colunas:  ["GROOT ID","NOME","MATRICULA","REGIME DE CONTRATO ","CARGO","DATA DE INÍCIO",
             "DATA FIM","% RATEIO","DIAS TRABALHADOS X FOLGA","ESCALA"],
  larguras: [7.54296875,39.1796875,9.1796875,17.08984375,17.26953125,
             11.6328125,8.6328125,7.453125,21.1796875,17],
  formatos: { 6:"mm-dd-yy", 7:"mm-dd-yy", 8:"0%", 9:"0%" },
  bordaCorpo: "cinza", cabSemFundo: true
};
const TPL_DIARISTAS = {
  colunas:  ["GROOT ID","NOME","CARGO","ESCALA","DATA","QUANTIDADE"],
  larguras: [7.54296875,35.26953125,17.26953125,17,8.6328125,10.1796875],
  formatos: { 5:"mm-dd-yy" },
  cabWrap:  true, fillCorpo: true, bordaCorpo: "clara"
};
/* As abas de documentação usam o mesmo cabeçalho amarelo — é o mesmo
   arquivo —, mas texto à esquerda: são frases, não números. */
const TPL_DOC = { esquerda:true };

const ymdParaData = v => { const p = ymdParts(v); return new Date(Date.UTC(p.y, p.m-1, p.d)); };
const semAcento = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"");
/* GROOT e matrícula são número na fatura; sair como texto muda a cara
   da coluna e quebra ordenação. */
const numSePuder = v => /^\d+$/.test(String(v ?? "").trim()) ? Number(v) : (v ?? "");

function abaEstilizada(wb, nome, tpl, linhas){
  const ws = wb.addWorksheet(nome);
  (tpl.larguras || []).forEach((w,i) => { ws.getColumn(i+1).width = w; });
  const borda = tpl.bordaCorpo === "cinza" ? TPL.bordaCinza : TPL.bordaClara;
  const cab = ws.addRow(tpl.colunas);
  cab.eachCell({ includeEmpty:true }, c => {
    c.font = TPL.fonteCab; c.fill = TPL.fillCab;
    c.border = tpl.cabSemFundo ? TPL.bordaCabSemFundo : TPL.bordaClara;
    c.alignment = tpl.cabWrap ? TPL.centroWrap
      : tpl.esquerda ? { horizontal:"left", vertical:"middle" } : TPL.centro;
  });
  for(const vals of linhas){
    const r = ws.addRow(vals);
    /* `eachCell` pula célula vazia mesmo com includeEmpty quando a linha
       é mais curta que o cabeçalho: percorrer pelo tamanho do cabeçalho
       garante que a última coluna também recebe borda. */
    for(let i=1;i<=tpl.colunas.length;i++){
      const c = r.getCell(i);
      c.font = TPL.fonte;
      c.border = borda;
      c.alignment = tpl.esquerda
        ? { horizontal: typeof vals[i-1] === "number" ? "center" : "left",
            vertical:"top", wrapText:true }
        : TPL.centro;
      if(tpl.fillCorpo) c.fill = TPL.fillBranco;
      if(tpl.formatos && tpl.formatos[i]) c.numFmt = tpl.formatos[i];
      else if(c.value instanceof Date) c.numFmt = "mm-dd-yy";
    }
  }
  return ws;
}

/* Os segmentos [início, fim] de uma linha depois do plano. Uma pausa
   parte a vigência em dois períodos — mesma pessoa, duas linhas. */
function segmentosApos(l, a){
  if(!a) return [{ ini:l.inicio, fim:l.fim }];
  if(a.tipo === "retirar") return [];
  const ini = a.novoInicio != null ? a.novoInicio : l.inicio;
  const fim = a.novoFim != null ? a.novoFim : l.fim;
  const segs = []; let cursor = ini;
  for(const p of (a.pausas || []).slice().sort((x,y) => x.fim - y.fim)){
    segs.push({ ini:cursor, fim:p.fim }); cursor = p.ini;
  }
  segs.push({ ini:cursor, fim });
  return segs;
}

/* A escala do diarista é a da OPERAÇÃO, e a fatura diz a unidade. Nome
   que casa com mais de uma operação — Pouso Alegre tem SVC e XD, com
   horários diferentes — fica em branco: escolher um turno seria chutar
   qual. */
function escalaDoDiarista(unidade){
  const u = norm(unidade);
  if(!u) return "";
  const ops = Object.keys(ESCALA_HORARIO_PADRAO).filter(op => norm(op) === u);
  return ops.length === 1 ? escalaHorarioDe(ops[0]) : "";
}

async function exportarEqualizado(){
  const sim = S.sim, plano = S.plano, fat = S.fatura;
  if(!sim || !plano || plano.erro || !fat) return;
  const dentro = simClassificarLinhas(fat.linhas).dentro;
  const acaoDe = new Map();
  plano.acoes.forEach(a => acaoDe.set(dentro[a.id], a));
  const dt = v => isValidYmd(v) ? ymdParaData(v) : null;

  /* ---- aba Labor: o quadro fixo depois do plano ---- */
  const linhasLabor = [];
  let mantidas = 0, removidas = 0, ajustadas = 0;
  for(const l of fat.linhas){
    const a = acaoDe.get(l);
    const segs = segmentosApos(l, a);
    if(!segs.length){ removidas++; continue; }
    if(a) ajustadas++; else mantidas++;
    for(const seg of segs) linhasLabor.push([
      numSePuder(l.groot), l.nome, numSePuder(l.matricula), l.regime, l.cargo,
      dt(seg.ini), dt(seg.fim),
      (typeof l.rateio === "number" && isFinite(l.rateio)) ? l.rateio : 1,
      l.diasFolga, l.escala ]);
  }

  /* ---- aba Diaristas: quem entra para cobrir a falta ---- */
  const incluir = $("sm-eqIncluir");
  const inc = (incluir && incluir.checked) ? plano.inclusoes : null;
  const escala = escalaDoDiarista(fat.unidade);
  /* As diárias que a fatura JÁ lança saem primeiro, como estão: o
     arquivo é a fatura equalizada inteira, não só o que o app acrescentou.
     Substituí-las apagaria cobrança legítima. */
  const jaNaFatura = (fat.diarias || []).filter(d =>
    d.data >= sim.periodo.ini && d.data <= sim.periodo.fim);
  const linhasDiar = jaNaFatura
    .slice().sort((a,b) => a.data - b.data || String(a.nome).localeCompare(String(b.nome)))
    .map(d => [numSePuder(d.groot), d.nome, d.cargo || "Diarista",
               d.escala || escala, ymdParaData(d.data), d.quantidade]);
  if(inc) for(const p of inc.pessoas) for(const d of p.dias)
    linhasDiar.push([numSePuder(p.groot), p.nome, "Diarista", escala, ymdParaData(d), 1]);

  /* ---- o dossiê do que foi feito ---- */
  const eqLinhas = plano.acoes.map(a => [
    EQ_ROTULO[a.tipo], numSePuder(a.linha.groot), a.linha.nome, a.linha.cargo,
    dt(a.linha.inicio), dt(a.linha.fim), dt(a.novoInicio), dt(a.novoFim),
    a.pausas.map(p => fmtYmd(p.fim)+" a "+fmtYmd(p.ini)).join(" · "),
    -a.impacto.hc, a.impacto.dias, a.motivo]);
  if(!eqLinhas.length) eqLinhas.push(["(nada a ajustar)"]);

  const incLinhas = [];
  if(inc) for(const p of inc.pessoas) for(const f of p.faixas)
    incLinhas.push([numSePuder(p.groot), p.nome,
      p.solic === "id" ? "ID" : p.solic === "meli" ? "Cliente" : "(sem solicitante)",
      dt(f.de), dt(f.ate), p.dias.filter(d => d >= f.de && d <= f.ate).length,
      "Solicitado no SIGO e sem cobrança no LABOR nestes dias"]);
  if(!incLinhas.length) incLinhas.push(["(ninguém a incluir)"]);

  const revLinhas = [];
  const diariasDe = {};
  for(const d of plano.dias) diariasDe[d.data] = d.diarias || 0;
  for(const [d,q] of Object.entries(plano.revisar).sort((a,b) => a[0]-b[0]))
    /* Excesso que cabe dentro das diárias do dia tem outra explicação:
       o quadro fixo já está no teto e o que passa é diária já lançada,
       em que a equalização não mexe. Dizer "partir um contrato" aí
       mandaria o usuário procurar no lugar errado. */
    revLinhas.push(["Excesso sem solução", dt(+d), q,
      diariasDe[+d] >= q
        ? "O quadro fixo já está no teto; o que passa do QF são as "+diariasDe[+d]
          + " diária(s) já lançadas neste dia, em que a equalização não mexe."
        : "Zerar exigiria partir um contrato — avaliar caso a caso."]);
  if(inc) for(const [d,c] of Object.entries(inc.dias).sort((a,b) => a[0]-b[0]))
    { if(c.descoberto > 0) revLinhas.push(["Falta descoberta", dt(+d), c.descoberto,
      "Não havia diarista livre bastante no SIGO neste dia ("+c.disponiveis+" disponível(is))."]); }
  if(!inc) for(const [d,q] of Object.entries(plano.falta).sort((a,b) => a[0]-b[0]))
    revLinhas.push(["Falta", dt(+d), q, "Quadro abaixo do QF; nenhuma inclusão foi gerada."]);
  if(!revLinhas.length) revLinhas.push(["(nada a revisar)"]);

  const metaLinhas = [
    ["Arquivo", "Labor equalizado gerado pela Validação Template"],
    ["Fatura de origem", S.nomeF],
    ["Unidade", fat.unidade || "—"],
    ["Período", fmtYmd(plano.periodo.ini)+" a "+fmtYmd(plano.periodo.fim)],
    ["QF do cliente (alvo)", plano.alvo],
    ["Permitir adiar início", plano.opcoes.permitirAdiarInicio ? "sim" : "não"],
    ["Pausar a partir de", plano.opcoes.pausaDesde ? fmtYmd(plano.opcoes.pausaDesde) : "não"],
    ["Linhas mantidas sem alteração", mantidas],
    ["Linhas ajustadas", ajustadas],
    ["Linhas removidas", removidas],
    ["Linhas na aba Diaristas", linhasDiar.length],
    ["  · diárias que a fatura já tinha", jaNaFatura.length],
    ["  · diárias acrescentadas pelo app", linhasDiar.length - jaNaFatura.length],
    ["Pessoas incluídas", inc ? inc.totais.pessoas : 0],
    ["  · da ID", inc ? inc.totais.id : 0],
    ["  · do cliente", inc ? inc.totais.meli : 0],
    ["  · sem solicitante", inc ? inc.totais.sem : 0],
    ["Falta coberta (pessoa-dia)", inc ? inc.totais.incluido : 0],
    ["Pessoa-dia conferidos contra o SIGO", inc ? inc.totais.verificados : 0],
    ["  · recusados por não estar na base no dia", inc ? inc.totais.recusados : 0],
    ["Falta descoberta (pessoa-dia)", inc ? inc.totais.descoberto : 0],
    ["Excesso antes do plano (HC-dia)", plano.totais.excessoAntes],
    ["Excesso depois do plano (HC-dia)", plano.totais.excessoDepois],
    ["Diárias já lançadas (pessoa-dia)", plano.totais.diarias],
    ["", ""],
    ["AVISO", "As alterações equalizam MATEMATICAMENTE o Labor ao QF. Confirme se "
      + "correspondem à movimentação operacional real antes de enviar."],
    ["Onde está a inclusão", "Quem entra para cobrir a falta é DIARISTA e está na aba "
      + "Diaristas, uma linha por pessoa-dia — não no Labor. Por isso o Labor desta "
      + "exportação fica ABAIXO do QF nos dias cobertos por diária: o quadro fixo é "
      + "esse mesmo, e a diferença é diária."],
    ["Como a inclusão foi escolhida", "Base SIGO: pessoa solicitada no dia e sem cobrança "
      + "naquele dia — nem no LABOR, nem como diária já lançada nesta fatura. Prioridade "
      + "para os diaristas da ID; o do cliente só entra depois de esgotados os internos."],
    ["Conferência das datas", "Cada pessoa-dia acrescentado foi reconferido contra a base do "
      + "SIGO, linha a linha, ANTES de virar linha do arquivo: ninguém entra num dia em que "
      + "não foi solicitado, e dia sem solicitação não é preenchido para emendar intervalo. "
      + "O que não passasse seria descartado e contado acima."],
    ["O que é o quadro do dia", "Quadro fixo do LABOR MAIS as diárias já lançadas na fatura. "
      + "É esse total que vai ao confronto com o QF: a vaga ocupada por diária é vaga "
      + "ocupada, e a falta do dia já vem descontada dela."],
    ["Campos em branco", "Matrícula, regime, dias trabalhados e turno das pessoas incluídas "
      + "não são deduzíveis do SIGO e por isso não são escritos. A ESCALA é a da operação "
      + "da unidade; unidade que casa com mais de uma operação sai em branco."]];

  const wb = new ExcelJS.Workbook();
  wb.creator = "Validação Template";
  abaEstilizada(wb, "Labor", TPL_LABOR, linhasLabor);
  abaEstilizada(wb, "Diaristas", TPL_DIARISTAS, linhasDiar);
  abaEstilizada(wb, "EQUALIZACAO", { ...TPL_DOC,
    colunas:["Ação","GROOT ID","Nome","Cargo","Início atual","Fim atual","Início novo",
             "Fim novo","Pausas","HC","Dias","Motivo"],
    larguras:[16,12,34,26,12,12,12,12,26,7,7,90] }, eqLinhas);
  abaEstilizada(wb, "INCLUSOES", { ...TPL_DOC,
    colunas:["GROOT ID","Nome","Solicitante","De","Até","Dias","Observação"],
    larguras:[12,34,16,12,12,7,52] }, incLinhas);
  abaEstilizada(wb, "REVISAR", { ...TPL_DOC,
    colunas:["Tipo","Data","Quantidade","Observação"],
    larguras:[22,12,12,72] }, revLinhas);
  abaEstilizada(wb, "METADADOS", { ...TPL_DOC,
    colunas:["Campo","Valor"], larguras:[34,110] }, metaLinhas);

  const per = String(fat.comp ? fat.comp.y*100 + fat.comp.m : "");
  /* Acento no `download` de um blob: faz o Chromium descartar o nome. */
  const nome = semAcento("Labor_equalizado_"+(fat.unidade || "fatura")+"_"+per+".xlsx")
    .replace(/[^\w.\-]+/g,"_");
  const buf = await wb.xlsx.writeBuffer();
  const link = document.createElement("a");
  link.href = URL.createObjectURL(new Blob([buf],
    { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  link.download = nome;
  link.click();
  URL.revokeObjectURL(link.href);

  const btn = $("sm-btnExportEq");
  if(btn) btn.textContent = "Baixado ✓ — " + nome;
  return nome;
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
               "Headcount bruto", "Q Pós previsto", "Gap PREF x S&OP", "Correção prevista",
               "Status", "Diagnóstico"];
  const aoa = [cab];
  for(const d of sim.dias){
    aoa.push([ dt(d.data), ...d.blocos.map(b => b.valor), d.qCliente, d.pref, d.bruto, d.qPos,
      d.gap, d.correcao, SIM_STATUS_LABEL[d.status], d.diagnostico ]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  formatarDatas(ws, aoa.length, 0);
  ws["!cols"] = [{wch:12}, ...sim.blocos.map(()=>({wch:11})), {wch:17},{wch:14},{wch:16},
                 {wch:16},{wch:16},{wch:17},{wch:24},{wch:110}];
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
    ["Planilha operacional", sim.fonte === "fixo" ? "(não usada)" : S.nomeS],
    ["Período", fmtYmd(sim.periodo.ini)+" a "+fmtYmd(sim.periodo.fim)],
    ["Fonte do S&OP", sim.fonte === "fixo"
      ? "valor fixo informado na tela: "+sim.valorFixo+" HC em todos os dias"
      : "planilha operacional, dia a dia por operação"],
    ["Competência", sim.comp ? sim.comp.label : ""],
    ["Blocos de S&OP somados", sim.blocos.join(" + ")],
    [],
    ["Fórmula — PREF", "soma dos % RATEIO das linhas ativas no dia, com o sinal — mesma regra da Fusão de Linhas"],
    ["Fórmula — Q cliente", "soma do Esperado de cada bloco no dia"],
    ["Fórmula — Q Pós previsto", "MIN(PREF, Q cliente)"],
    ["Fórmula — Gap", "PREF - Q cliente"],
    ["Fórmula — Correção prevista", "Q Pós previsto - PREF"],
    [],
    ["Linhas do LABOR no PREF", sim.linhas.dentro],
    ["Linhas fora do PREF", sim.linhas.fora.length],
    [],
    ["PREF total (HC-dia, líquido)", sim.totais.pref],
    ["Headcount bruto (HC-dia, só positivos)", sim.totais.bruto],
    ["Linhas de estorno no PREF", sim.totais.estornos],
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
/* O SIGO depende da fatura — é ela que diz de qual unidade ler. Em vez
   de recusar quem soltou na ordem trocada, o arquivo fica guardado e é
   lido assim que a fatura chega. */
bindDrop("sm-dropD","sm-fileD",(f,d) => {
  S.dropDiar = d;
  if(!S.fatura){
    S.diarPendente = f;
    $("sm-fnD").textContent = f.name;
    const st = $("sm-stD");
    st.textContent = "aguardando a fatura — é ela que diz de qual unidade ler o SIGO";
    st.className = "st warn";
    return;
  }
  carregar("diar",f,d);
});
document.querySelectorAll('input[name="sm-fonte"]').forEach(r => r.onchange = pronto);
const campoFixo = $("sm-sopFixo");
if(campoFixo) campoFixo.oninput = pronto;
pronto();   // deixa o passo 3 e o destaque da fonte coerentes já na abertura
[$("sm-eqAdiar"), $("sm-eqPausa"), $("sm-eqIncluir")].forEach(el => { if(el) el.onchange = () => { if(S.sim) desenharEqualizacao(); }; });
const btnRun = $("sm-btnRun"); if(btnRun) btnRun.onclick = analisar;
const btnExp = $("sm-btnExport"); if(btnExp) btnExp.onclick = exportar;

})();
