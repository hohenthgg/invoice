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
            baseG:null, cadastro:null, correcoes:[],
            nomeF:"", nomeS:"", nomeD:"", nomeG:"", fonte:"planilha",
            /* null = ninguém decidiu ainda; "manter" | "reduzir" */
            decisaoDiarias:null };

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
        iDf=col(c,"DIAS TRABALHADOS X FOLGA"), iEs=col(c,"ESCALA"), iU=col(c,"UNIDADE");
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
  const diarias = lerDiariasDaFatura(wb);
  /* O código da unidade (SMG9) sai da própria coluna UNIDADE do LABOR —
     é ele que casa com a coluna Filiais da base oficial. */
  let unidadeCod = "";
  if(iU >= 0) for(const l of linhas){ const v = String(l.bruta[iU] ?? "").trim();
    if(v){ unidadeCod = v; break; } }
  return { aba:nome, linhas, comp, unidade: unidadeDaFatura(wb), unidadeCod,
    diarias, layoutDiarias: lerDiariasDaFatura.layout || null,
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
  lerDiariasDaFatura.layout = null;
  const nome = acharAba(wb,"DIARISTAS","DIARISTA");
  if(!nome) return [];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[nome],{header:1,defval:""});
  const h = acharCabecalho(rows,["GROOT ID","DATA"]);
  if(!h) return [];
  const c = h.cels;
  const iG=col(c,"GROOT ID"), iN=col(c,"NOME"), iC=col(c,"CARGO"), iE=col(c,"ESCALA"),
        iD=col(c,"DATA"), iQ=col(c,"QUANTIDADE"), iO=col(c,"OBSERVACOES","OBSERVAÇÕES");
  if(iG < 0 || iD < 0) return [];
  lerDiariasDaFatura.layout = { aba:nome, topo:h.idx+1, largura:Math.max(c.length,
    ...rows.slice(h.idx+1).map(r => (r||[]).length)),
    cols:{ groot:iG, nome:iN, cargo:iC, escala:iE, data:iD, qtd:iQ, obs:iO,
           fixas:["FORNECEDOR","PERIODO","PERIODO - MES/ANO","UNIDADE","DESCRICAO CONTA",
                  "CONTA CONTABIL","UNIDADE + UF"].map(x => col(c,x)) } };
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
      data, quantidade: isFinite(q) && q !== 0 ? q : 1,
      bruta: r.slice(), linha:i+1 });
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
   LEITURA — BASE OFICIAL DE GROOT

   Groot ID | Nome | CPF | Filiais, em qualquer aba: vale a primeira
   cujo cabeçalho tenha GROOT e NOME. CPF e Filiais são opcionais —
   sem eles a conferência perde força, não a validade.
   ================================================================ */
function lerBaseGroot(wb){
  for(const nomeAba of wb.SheetNames){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[nomeAba],{header:1,defval:""});
    const h = acharCabecalho(rows,["GROOT","NOME"]);
    if(!h) continue;
    const c = h.cels;
    const iG=col(c,"GROOT ID","GROOT"), iN=col(c,"NOME"),
          iC=col(c,"CPF"), iF=col(c,"FILIAIS","FILIAL");
    const regs = [];
    for(let i=h.idx+1;i<rows.length;i++){
      const r = rows[i]; if(!r) continue;
      const groot = String(r[iG] ?? "").replace(/\D/g,"");
      const nomeP = String(r[iN] ?? "").trim();
      if(!groot || !nomeP) continue;
      regs.push({ groot, nome:nomeP,
        cpf: iC >= 0 ? String(r[iC] ?? "").trim() : "",
        filiais: iF >= 0 ? String(r[iF] ?? "").trim() : "" });
    }
    if(regs.length) return { aba:nomeAba, regs,
      comCpf: regs.filter(r => r.cpf).length };
  }
  return { erro:"Nenhuma aba com colunas GROOT ID e NOME — é este o formato da base oficial." };
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

const CAMPOS = { fatura:["sm-stF","sm-fnF"], sop:["sm-stS","sm-fnS"],
                 diar:["sm-stD","sm-fnD"], groot:["sm-stG","sm-fnG"] };

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
              : qual === "groot"  ? lerBaseGroot(wb)
                                  : lerDiaristasSigo(wb, S.fatura ? S.fatura.unidade : "");
      if(r.erro){ st.textContent = "✗ "+r.erro; st.className = "st bad"; drop.classList.add("err");
                  S[qual] = null; return pronto(); }
      if(qual === "fatura"){
        S.fatura = r; S.nomeF = file.name;
        S.decisaoDiarias = null;   // fatura nova, decisão nova
        S.correcoes = []; S.cadastro = null;
        /* O arquivo cru fica guardado: a exportação devolve a FATURA
           INTEIRA com duas abas trocadas, e as outras onze só continuam
           lá se vierem do original. */
        S.fatura.buffer = e.target.result;
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
      } else if(qual === "groot"){
        S.baseG = r; S.nomeG = file.name;
        st.textContent = "✓ aba "+r.aba+" · "+r.regs.length+" pessoa(s) na base"
          + (r.comCpf ? " · "+r.comCpf+" com CPF" : "");
        /* Base carregada depois da análise: reconfere na hora. */
        if(S.sim) desenharCadastro();
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
  const sim = simularRetorno({ labor:S.fatura.linhas, blocos, periodo:per, comp, diaristas,
    diarias: S.fatura.diarias });
  if(sim.erro) return falhar($("sm-err"), sim.erro);
  sim.fonte = fonte;
  sim.valorFixo = valorFixo;
  S.fonte = fonte;
  S.sim = sim;
  render();
}

/* ================================================================
   TELA — O FLUXO DE FECHAMENTO

   Cinco passos, um embaixo do outro, cada um com um resumo de uma
   linha no cabeçalho e o detalhe no corpo. A régua de tudo aqui é a
   pedida na especificação: poucas informações por tela, decisões
   explícitas, detalhe técnico recolhido.

     1  Cadastro × base oficial   (se a base de GROOT foi carregada)
     2  Quadro × S&OP             Data | Quadro | SOP | Diferença
     3  Cobrir faltas             faltavam X → adicionados Y
     4  Corrigir excessos         o motor da Fusão, e a decisão da diária
     5  Revisão e exportação      o resumo curto e o arquivo

   O raciocínio não mora aqui: quadro e plano vêm de js/simulacao.js,
   o motor é o de js/equalizacao.js, o cadastro é js/cadastro.js.
   ================================================================ */
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const n2 = v => v === null || v === undefined || !isFinite(Number(v))
  ? "—" : Number(Number(v).toFixed(2)).toLocaleString("pt-BR");
const sinal = v => v === null || v === undefined ? "—" : (Number(v) > 0 ? "+" : "") + n2(v);

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
  return { permitirAdiarInicio: c ? c.checked : true, pausaDesde,
           cortarDiarias: S.decisaoDiarias === "reduzir" };
}

/* A DECISÃO SOBRE A DIÁRIA JÁ ALOCADA

   O Labor se ajusta sozinho porque é projeção. A diária não: ela já
   foi alocada e já está faturada, e reduzir a quantidade de um dia é
   decisão de quem opera. O app calcula quanto DARIA para cortar,
   mostra, e não corta até alguém escolher — a exportação fica
   bloqueada enquanto isso, para o arquivo nunca sair de um estado que
   ninguém decidiu. */
function decidirDiarias(escolha){
  S.decisaoDiarias = escolha;
  render();
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
      + esc(a.diarista.dias.map(fmtShort).join(", ")) + '). Costuma ser transição de fixo para '
      + 'diária — confirme se não há dupla cobrança nesses dias.</p>');
  }
  return partes.join("");
}

/* ---------------------------------------------------------------
   O plano do período — quadro, faltas, excessos. Uma chamada só, que
   serve os passos 2, 3, 4 e 5. Com S&OP fixo o alvo é um número; com
   a planilha operacional, o alvo é o S&OP de cada dia.
   --------------------------------------------------------------- */
function calcularPlano(){
  const sim = S.sim;
  const diaristas = S.diar
    ? S.diar.regs.filter(x => x.data >= sim.periodo.ini && x.data <= sim.periodo.fim) : [];
  const dados = { labor:S.fatura.linhas, periodo:sim.periodo,
    diaristas, diarias:S.fatura.diarias, opcoes:opcoesEq() };
  if(sim.fonte === "fixo") dados.alvo = sim.valorFixo;
  else {
    dados.alvos = {};
    for(const d of sim.dias) dados.alvos[d.data] = d.qCliente;
  }
  S.plano = simPlanoEqualizacao(dados);
}

function render(){
  $("sm-result").classList.remove("hidden");
  $("sm-avisos").innerHTML = (S.sim.avisos || []).map(a =>
    '<div class="sm-aviso '+a.tipo+'">'+esc(a.texto)+'</div>').join("")
    || '<p class="fx-nada">Nenhum aviso técnico.</p>';
  calcularPlano();
  desenharCadastro();
  desenharQuadro();
  desenharFaltas();
  desenharExcessos();
  desenharFinal();
}

/* ================================================================
   PASSO 1 — CADASTRO × BASE OFICIAL
   ================================================================ */
function pessoasDaFatura(){
  const vistos = new Map();
  const add = (groot, nome, origem) => {
    const chave = simGrootNum(groot)+"|"+norm(nome);
    if(!nome) return;
    if(vistos.has(chave)){ vistos.get(chave).origem.add(origem); return; }
    vistos.set(chave, { groot, nome, origem:new Set([origem]) });
  };
  for(const l of S.fatura.linhas) add(l.groot, l.nome, "LABOR");
  for(const d of (S.fatura.diarias || [])) add(d.groot, d.nome, "DIARISTAS");
  return [...vistos.values()].map(p => ({ groot:p.groot, nome:p.nome,
    origem:[...p.origem].join(" + ") }));
}

/* Aplica um GROOT confirmado à fatura EM MEMÓRIA — linhas do LABOR, da
   aba DIARISTAS e as células cruas que a exportação reescreve. O
   arquivo original só muda na exportação; aqui muda o que o resto do
   fluxo enxerga, e por isso tudo é recalculado em seguida. */
function mutarGroot(deGroot, nome, paraGroot){
  const alvoG = simGrootNum(deGroot), alvoN = norm(nome);
  const trocados = [];
  const C = S.fatura.layout ? S.fatura.layout.cols : {};
  for(const l of S.fatura.linhas){
    if(simGrootNum(l.groot) !== alvoG || norm(l.nome) !== alvoN) continue;
    trocados.push({ tipo:"labor", ref:l, prevGroot:l.groot,
      prevBruta: (l.bruta && C.groot >= 0) ? l.bruta[C.groot] : undefined });
    l.groot = paraGroot;
    if(l.bruta && C.groot >= 0) l.bruta[C.groot] = numSePuder(paraGroot);
  }
  const CD = S.fatura.layoutDiarias ? S.fatura.layoutDiarias.cols : {};
  for(const d of (S.fatura.diarias || [])){
    if(simGrootNum(d.groot) !== alvoG || norm(d.nome) !== alvoN) continue;
    trocados.push({ tipo:"diaria", ref:d, prevGroot:d.groot,
      prevBruta: (d.bruta && CD.groot >= 0) ? d.bruta[CD.groot] : undefined });
    d.groot = paraGroot;
    if(d.bruta && CD.groot >= 0) d.bruta[CD.groot] = numSePuder(paraGroot);
  }
  return trocados;
}

function aplicarCorrecao(nome, deGroot, paraGroot){
  const trocados = mutarGroot(deGroot, nome, paraGroot);
  if(!trocados.length) return;
  S.correcoes.push({ nome, de:simGrootNum(deGroot), para:simGrootNum(paraGroot),
    linhas:trocados.length, trocados });
  reanalisar();
}

function desfazerCorrecao(i){
  const c = S.correcoes[i];
  if(!c) return;
  for(const t of c.trocados){
    t.ref.groot = t.prevGroot;
    if(t.prevBruta !== undefined){
      const cols = t.tipo === "labor" ? S.fatura.layout.cols : S.fatura.layoutDiarias.cols;
      t.ref.bruta[cols.groot] = t.prevBruta;
    }
  }
  S.correcoes.splice(i,1);
  reanalisar();
}

/* Recalcula tudo a partir da fatura em memória — depois de uma
   correção de cadastro o quadro, a ocupação e o plano podem mudar. */
function reanalisar(){ analisar(); }

const CAD_ROTULO = { corrigir:"GROOT incorreto", preencher:"GROOT ausente",
                     ambiguo:"Mais de um candidato", fora:"Fora da base" };

function desenharCadastro(){
  const sec = $("fx-cadastro");
  const aplicadas = S.correcoes.length
    ? '<div class="cad-aplicadas">'
      + S.correcoes.map((c,i) =>
          '<div class="cad-ok"><span>✓</span> '+esc(c.nome)+' · <code>'
          + esc(c.de || "—")+' → '+esc(c.para)+'</code> ('+c.linhas+' linha(s))'
          + ' <button class="cad-undo" data-undo="'+i+'">desfazer</button></div>').join("")
      + '</div>' : "";

  if(!S.baseG){
    sec.hidden = false;
    $("fx-cadResumo").textContent = "sem base carregada — passo pulado";
    $("fx-cadCorpo").innerHTML = aplicadas
      + '<p class="fx-nada">Carregue a <b>base oficial de GROOT</b> no passo de arquivos para '
      + 'conferir e corrigir o cadastro (Groot ID · Nome · CPF · Filiais).</p>';
    ligarCadastro();
    return;
  }

  const r = cadConferir({ pessoas:pessoasDaFatura(), base:S.baseG.regs,
    unidadeCod:S.fatura.unidadeCod, unidadeNome:S.fatura.unidade });
  S.cadastro = r;
  sec.hidden = false;

  const acionaveis = r.achados.filter(a => a.tipo === "corrigir" || a.tipo === "preencher");
  const ambiguos   = r.achados.filter(a => a.tipo === "ambiguo");
  const fora       = r.achados.filter(a => a.tipo === "fora");

  $("fx-cadResumo").textContent =
    (acionaveis.length || ambiguos.length)
      ? acionaveis.length+" sugestão(ões)"
        + (ambiguos.length ? " · "+ambiguos.length+" para decidir" : "")
        + (S.correcoes.length ? " · "+S.correcoes.length+" aplicada(s)" : "")
      : S.correcoes.length
        ? S.correcoes.length+" correção(ões) aplicada(s) — cadastro fechado"
        : "cadastro confere com a base ("+r.totais.ok+" pessoas)";

  const card = (a,i) => {
    const chip = '<span class="cad-chip '+a.tipo+'">'+CAD_ROTULO[a.tipo]+'</span>';
    if(a.tipo === "corrigir" || a.tipo === "preencher"){
      return '<div class="cad-card">'
        + '<div class="cab" onclick="this.parentElement.classList.toggle(\'aberto\')">'+chip
        + '<b>'+esc(a.nome)+'</b><code>'+esc(a.groot || "—")+' → '+esc(a.sugestao.groot)+'</code>'
        + '<span class="conf '+a.confianca+'">'+(a.confianca === "alta" ? "confiança alta" : "confiança média")+'</span>'
        + '<button class="vt-btn mini" data-aplicar="'+i+'">Aplicar</button></div>'
        + '<div class="det"><p>'+esc(a.motivo)+'</p><p class="ref">Na base: <b>'+esc(a.sugestao.nome)+'</b>'
        + (a.sugestao.cpf ? ' · CPF '+esc(a.sugestao.cpf) : "")
        + (a.sugestao.filiais ? ' · '+esc(a.sugestao.filiais) : "")
        + ' · aparece em '+esc(a.origem)+'</p></div></div>';
    }
    if(a.tipo === "ambiguo"){
      return '<div class="cad-card aberto">'
        + '<div class="cab">'+chip+'<b>'+esc(a.nome)+'</b><code>'+esc(a.groot || "—")+'</code>'
        + '<span class="conf">decisão sua</span></div>'
        + '<div class="det"><p>'+esc(a.motivo)+'</p><div class="cad-cands">'
        + a.candidatos.map(c =>
            '<button class="cad-cand" data-nome="'+esc(a.nome)+'" data-de="'+esc(a.groot)+'" '
            + 'data-para="'+esc(c.groot)+'">Usar <b>'+esc(c.groot)+'</b>'
            + (c.filiais ? ' · '+esc(c.filiais) : "")+(c.cpf ? ' · CPF '+esc(c.cpf) : "")
            + '</button>').join("")
        + '</div></div></div>';
    }
    return "";
  };

  const foraBloco = fora.length
    ? '<details class="cad-fora"><summary>'+fora.length+' pessoa(s) fora da base — '
      + 'informação, nada a corrigir</summary>'
      + fora.map(a => '<div class="cad-foraitem"><b>'+esc(a.nome)+'</b> · '
          + esc(a.groot || "sem GROOT")+' · '+esc(a.origem)+'<br><small>'+esc(a.motivo)+'</small></div>').join("")
      + '</details>' : "";

  $("fx-cadCorpo").innerHTML =
    aplicadas
    + (acionaveis.length || ambiguos.length
        ? r.achados.map((a,i) => card(a,i)).join("")
        : '<p class="fx-nada">Todos os GROOTs conferem com a base oficial.</p>')
    + foraBloco;
  ligarCadastro();
}

function ligarCadastro(){
  const corpo = $("fx-cadCorpo");
  if(!corpo) return;
  [...corpo.querySelectorAll("[data-aplicar]")].forEach(b => b.onclick = ev => {
    ev.stopPropagation();
    const a = S.cadastro.achados[+b.dataset.aplicar];
    aplicarCorrecao(a.nome, a.groot, a.sugestao.groot);
  });
  [...corpo.querySelectorAll(".cad-cand")].forEach(b => b.onclick = () =>
    aplicarCorrecao(b.dataset.nome, b.dataset.de, b.dataset.para));
  [...corpo.querySelectorAll("[data-undo]")].forEach(b => b.onclick = () =>
    desfazerCorrecao(+b.dataset.undo));
}

/* ================================================================
   PASSO 2 — QUADRO × S&OP

   A tela mais simples do fluxo, de propósito: Data | Quadro | SOP |
   Diferença. O quadro é Labor ativo + diaristas já lançados na
   fatura; o estorno (rateio ≤ 0) não é pessoa e fica fora da conta —
   e fora do alcance de qualquer ajuste.
   ================================================================ */
function desenharQuadro(){
  const p = S.plano;
  if(p.erro){ $("fx-quadroCorpo").innerHTML = '<div class="sm-aviso">'+esc(p.erro)+'</div>'; return; }
  const acima  = p.dias.filter(d => d.dif > 0).length;
  const abaixo = p.dias.filter(d => d.dif < 0).length;
  const semAlvo = p.dias.filter(d => d.dif == null).length;
  $("fx-quadroResumo").textContent =
    (!acima && !abaixo) ? "todos os dias no alvo"
      : (acima ? acima+" dia(s) acima" : "") + (acima && abaixo ? " · " : "")
        + (abaixo ? abaixo+" dia(s) abaixo" : "")
        + (semAlvo ? " · "+semAlvo+" sem S&OP" : "");

  const linhas = p.dias.map(d =>
    '<tr class="'+(d.dif == null ? "" : d.dif > 0 ? "acima" : d.dif < 0 ? "abaixo" : "ok")+'">'
    + '<td>'+fmtYmd(d.data)+'</td>'
    + '<td class="forte">'+n2(d.quadro)
    + (d.diarias ? '<small>'+n2(d.pref)+' fixo + '+n2(d.diarias)+' diária</small>' : "")+'</td>'
    + '<td>'+(d.alvo == null ? "—" : n2(d.alvo))+'</td>'
    + '<td class="'+(d.dif > 0 ? "pos" : d.dif < 0 ? "neg" : "")+'">'
    + (d.dif == null ? "sem S&OP" : d.dif === 0 ? "✓" : sinal(d.dif))+'</td></tr>').join("");
  $("fx-quadroCorpo").innerHTML =
    '<table class="fx-tab"><thead><tr><th>Data</th><th>Quadro</th>'
    + '<th>'+(S.sim.fonte === "fixo" ? "QF" : "S&amp;OP")+'</th><th>Diferença</th></tr></thead>'
    + '<tbody>'+linhas+'</tbody></table>';
}

/* ================================================================
   PASSO 3 — COBRIR FALTAS COM DIARISTA ID
   ================================================================ */
function desenharFaltas(){
  const p = S.plano;
  if(p.erro){ $("fx-faltasCorpo").innerHTML = ""; $("fx-faltasResumo").textContent = "—"; return; }
  const faltaTotal = Object.values(p.falta || {}).reduce((s,v) => s+v, 0);
  const inc = p.inclusoes;

  if(!faltaTotal){
    $("fx-faltasResumo").textContent = "nenhum dia abaixo do alvo";
    $("fx-faltasCorpo").innerHTML = '<p class="fx-nada">Nenhuma falta a cobrir.</p>';
    return;
  }
  if(!inc){
    $("fx-faltasResumo").textContent = "faltam "+n2(faltaTotal)+" pessoa-dia — carregue o SIGO";
    $("fx-faltasCorpo").innerHTML = '<p class="fx-nada">Há <b>'+n2(faltaTotal)+' pessoa-dia</b> abaixo '
      + 'do alvo. Carregue a base <b>SIGO</b> no passo de arquivos para cobrir com diaristas ID — '
      + 'só entram pessoas solicitadas naquele mesmo dia e ainda não cobradas.</p>';
    return;
  }

  $("fx-faltasResumo").textContent =
    "faltavam "+n2(faltaTotal)+" → adicionados "+n2(inc.totais.incluido)+" diaristas"
    + (inc.totais.descoberto > 0 ? " · "+n2(inc.totais.descoberto)+" sem cobertura" : "");

  const nomes = inc.pessoas.map(pp =>
    '<div class="sm-eqitem" onclick="this.classList.toggle(\'aberto\')">'
    + '<div class="cab"><div class="nm">'+esc(pp.nome || "(sem nome no SIGO)")
    + '<small>Groot '+esc(pp.groot)+' · diarista '
    + (pp.solic === "id" ? "ID" : pp.solic === "meli" ? "do cliente" : "sem solicitante")+'</small></div>'
    + '<div class="mud">'+esc(pp.faixas.map(faixaTxt).join(", "))
    + '<small>+'+pp.total+' dia(s)</small></div></div>'
    + '<div class="det"><p>Entra na aba <b>DIARISTAS</b> do template, no mesmo padrão das linhas '
    + 'originais: uma linha por dia, cargo <i>Diarista</i>, quantidade 1.</p></div></div>').join("");

  $("fx-faltasCorpo").innerHTML =
    '<p class="fx-frase">Faltavam <b>'+n2(faltaTotal)+'</b> pessoa-dia → adicionados '
    + '<b>'+n2(inc.totais.incluido)+'</b> diaristas ('+inc.totais.id+' ID'
    + (inc.totais.meli ? ' · '+inc.totais.meli+' do cliente, só depois de esgotar os da ID' : "")+').'
    + (inc.totais.descoberto > 0
        ? ' Ficam <b>'+n2(inc.totais.descoberto)+'</b> pessoa-dia sem diarista livre no SIGO.' : "")
    + ' Ninguém entra em dia em que já está no Labor, já é diária da fatura ou já foi '
    + 'adicionado pelo próprio processo.</p>'
    + '<details class="fx-exp"><summary>Ver os '+inc.totais.pessoas+' nomes</summary>'+nomes+'</details>';
}

/* ================================================================
   PASSO 4 — CORRIGIR EXCESSOS NO LABOR

   O mesmo motor da Fusão de Linhas, com as mesmas proteções: só mexe
   onde resolve excesso sem criar falta indevida em outro dia. O Labor
   é sempre a primeira camada; a diária já lançada só sai com decisão
   explícita.
   ================================================================ */
function desenharExcessos(){
  const p = S.plano;
  if(p.erro){ $("fx-excessosCorpo").innerHTML = ""; $("fx-excessosResumo").textContent = "—"; return; }
  const t = p.totais;
  const podeCortar = p.corteDisponivel ? p.corteDisponivel.totais.cortado : 0;

  const partes = [];
  if(t.retirar)  partes.push("− "+t.retirar+" linha(s)");
  if(t.adiar)    partes.push("↪ "+t.adiar+" início(s)");
  if(t.encurtar) partes.push("↩ "+t.encurtar+" fim(ns)");
  if(t.pausar)   partes.push("⏸ "+t.pausar+" pausa(s)");
  $("fx-excessosResumo").textContent =
    partes.length ? partes.join(" · ")
      + (podeCortar > 0 && !S.decisaoDiarias ? " · decisão pendente" : "")
    : podeCortar > 0 ? "excesso só em diária — decisão pendente"
    : "nenhum excesso no período";

  const grupos = EQ_ORDEM.map(tipo => {
    const itens = p.acoes.filter(a => a.tipo === tipo);
    if(!itens.length) return "";
    return '<div class="sm-eqgrupo">'
      + '<header><span class="sm-eqchip '+tipo+'">'+EQ_ROTULO[tipo]+'</span>'
      + '<span class="n">'+itens.length+' pessoa(s)</span></header>'
      + '<details class="fx-exp fx-exp-grupo"><summary>Ver os '+itens.length+' nomes</summary>'
      + itens.map(a =>
          '<div class="sm-eqitem" onclick="this.classList.toggle(\'aberto\')">'
          + '<div class="cab"><div class="nm">'+esc(a.linha.nome)
          + '<small>Groot '+esc(a.linha.groot || "—")+' · '+esc(a.linha.cargo)+'</small></div>'
          + '<div class="mud">'+esc(eqMudanca(a))
          + '<small>'+sinal(-a.impacto.hc)+' HC · '+a.impacto.dias+' dia(s)'
          + (a.diarista && a.diarista.dias.length ? ' · também diarista' : "")+'</small></div></div>'
          + '<div class="det">'+eqDetalhe(a)+'</div></div>').join("")
      + '</details></div>';
  }).join("");

  const decisao = (podeCortar > 0) ? (() => {
    const d = S.decisaoDiarias;
    const dias = Object.keys(p.corteDisponivel.dias).length;
    const bt = (chave,rot,desc) =>
      '<button class="sm-eqdec'+(d === chave ? " ativo" : "")+'" data-dec="'+chave+'">'
      + '<b>'+rot+'</b><i>'+desc+'</i></button>';
    return '<div class="sm-eqdecisao'+(d ? " decidido" : "")+'">'
      + '<div class="tit">'+(d ? "Decisão tomada" : "Falta uma decisão sua")+' — '
      + '<b>'+n2(p.corteDisponivel.totais.excesso)+' HC-dia</b> acima do alvo em '+dias+' dia(s), '
      + 'e o Labor já foi ajustado até onde dava</div>'
      + '<p>O que passa do alvo nesses dias é <b>diária já alocada e faturada</b>. Reduzir isso é '
      + 'decisão de quem opera — o app não decide sozinho.</p>'
      + '<div class="sm-eqdecs">'
      + bt("manter","Manter diaristas", "não mexer neles; o excesso residual sai declarado em REVISAR")
      + bt("reduzir","Reduzir diaristas", "retirar só o necessário para chegar ao alvo — "
          + n2(podeCortar)+" pessoa-dia, a interna primeiro")
      + '</div></div>';
  })() : "";

  const cortes = (p.corte && p.corte.cortes.length) ? (() => {
    const tc = p.corte.totais;
    const porDia = new Map();
    for(const c of p.corte.cortes){
      if(!porDia.has(c.data)) porDia.set(c.data, []);
      porDia.get(c.data).push(c);
    }
    return '<div class="sm-eqgrupo"><header><span class="sm-eqchip cortar">Retirar diária</span>'
      + '<span class="n">'+n2(tc.cortado)+' pessoa-dia</span></header>'
      + '<div class="det aberto"><p>Você escolheu <b>reduzir</b>. Sai só o necessário para chegar '
      + 'ao alvo — a interna primeiro, a do cliente por último ('+tc.id+' interna(s) · '
      + tc.meli+' do cliente'+(tc.sem ? ' · '+tc.sem+' sem solicitante' : "")+').</p></div>'
      + [...porDia.entries()].sort((a,b) => a[0]-b[0]).map(([d,cs]) =>
          '<div class="sm-eqitem" onclick="this.classList.toggle(\'aberto\')">'
          + '<div class="cab"><div class="nm">'+fmtYmd(d)
          + '<small>'+cs.length+' diária(s) acima do alvo</small></div>'
          + '<div class="mud">−'+cs.length+'</div></div>'
          + '<div class="det"><p>'+cs.map(c => esc(c.nome || "(sem nome)")
              + ' <b>'+esc(c.groot)+'</b>').join(' · ')+'</p></div></div>').join("")
      + '</div>';
  })() : "";

  const listaDias = (o, neg) => Object.keys(o).map(Number).sort((a,b)=>a-b)
    .map(d => fmtShort(d)+" ("+sinal(neg ? -o[d] : o[d])+")").join(" · ");
  const revisar = Object.keys(p.revisar).length
    ? '<div class="sm-eqgrupo revisar"><header><span class="sm-eqchip revisar">Revisar</span>'
      + '<span class="n">'+Object.keys(p.revisar).length+' dia(s)</span></header>'
      + '<div class="det aberto"><p>Excesso que nenhum ajuste resolve sem criar falta indevida. '
      + 'Avalie caso a caso.</p><p class="dias">'+listaDias(p.revisar,false)+'</p></div></div>' : "";

  $("fx-excessosCorpo").innerHTML =
    (grupos || decisao || cortes || revisar)
      ? '<p class="fx-frase">O <b>motor da Fusão de Linhas</b> analisou o período inteiro. Cada '
        + 'ajuste só entra se resolver excesso <b>sem criar falta em outro dia</b>; estornos '
        + '(rateio ≤ 0) são preservados como estão e nunca são candidatos.</p>'
        + grupos + decisao + cortes + revisar
      : '<p class="fx-nada">Nenhum excesso a corrigir.</p>';
  [...$("fx-excessosCorpo").querySelectorAll(".sm-eqdec")].forEach(b =>
    b.onclick = () => decidirDiarias(b.dataset.dec));
}

/* ================================================================
   PASSO 5 — REVISÃO E EXPORTAÇÃO
   ================================================================ */
function desenharFinal(){
  const p = S.plano;
  if(p.erro){ $("fx-finalCorpo").innerHTML = ""; $("fx-finalResumo").textContent = "—"; return; }
  const t = p.totais;
  const inc = p.inclusoes;
  const podeCortar = p.corteDisponivel ? p.corteDisponivel.totais.cortado : 0;
  const pendente = podeCortar > 0 && !S.decisaoDiarias;
  const revisarN = Object.keys(p.revisar).length
    + (inc && inc.totais.descoberto > 0 ? 1 : 0);

  const linhas = [];
  if(inc && inc.totais.incluido) linhas.push(["+", n2(inc.totais.incluido)+" diárias adicionadas","mais"]);
  if(t.retirar)  linhas.push(["−", t.retirar+" linha(s) retiradas do Labor","menos"]);
  if(t.adiar)    linhas.push(["↪", t.adiar+" data(s) de início ajustadas","aj"]);
  if(t.encurtar) linhas.push(["↩", t.encurtar+" data(s) fim ajustadas","aj"]);
  if(t.pausar)   linhas.push(["⏸", t.pausar+" contrato(s) pausados e retomados","aj"]);
  if(p.corte && p.corte.totais.cortado)
    linhas.push(["−", n2(p.corte.totais.cortado)+" diária(s) retiradas (sua decisão)","menos"]);
  if(S.correcoes.length) linhas.push(["✎", S.correcoes.length+" GROOT(s) corrigidos","aj"]);
  if(revisarN) linhas.push(["⚠", revisarN+" caso(s) para revisar","rev"]);
  if(!linhas.length) linhas.push(["✓","nada a corrigir — a fatura sai como entrou","ok"]);

  $("fx-finalResumo").textContent = pendente ? "aguardando a decisão do passo 4"
    : t.excessoDepois === 0 && (!inc || !inc.totais.descoberto)
      ? "pronta para exportar" : "pronta, com ressalvas em REVISAR";

  $("fx-finalCorpo").innerHTML =
    '<div class="fx-resumofinal">'
    + linhas.map(([ic,tx,cls]) => '<div class="fx-rl '+cls+'"><span>'+ic+'</span>'+esc(tx)+'</div>').join("")
    + '</div>'
    + '<div class="sm-eqselo">As correções equalizam <b>matematicamente</b> a fatura ao alvo. '
    + 'Confirme se correspondem à movimentação operacional real antes de enviar. O arquivo '
    + 'exportado é <b>o mesmo template</b> — estrutura, abas e formatação preservadas — com as '
    + 'abas LABOR e DIARISTAS corrigidas e o memorial em EQUALIZACAO / INCLUSOES / REVISAR / '
    + 'CADASTRO / METADADOS.</div>'
    + '<div class="vt-acoes-fim">'
    + '<button class="vt-btn" id="sm-btnExportEq"'+(pendente ? " disabled" : "")+'>'
    + 'Exportar fatura corrigida (.xlsx)</button>'
    + (pendente ? '<span class="sm-msg">Decida no passo 4 o que fazer com as diárias acima do alvo.</span>' : "")
    + '</div>';
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
/* As abas de documentação usam o mesmo cabeçalho amarelo — é o mesmo
   arquivo —, mas texto à esquerda: são frases, não números. */
const TPL_DOC = { esquerda:true };

const ymdParaData = v => { const p = ymdParts(v); return new Date(Date.UTC(p.y, p.m-1, p.d)); };
/* Data → serial do Excel, pelos componentes UTC: sem fuso, sem deriva. */
const SERIAL0 = Date.UTC(1899,11,30);
const dataParaSerial = d => Math.round(
  (Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - SERIAL0)/864e5);
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

/* A FATURA INTEIRA, COM DUAS ABAS TROCADAS

   O arquivo que sai é o arquivo que entrou: as treze abas do template,
   com os estilos, as fórmulas e tudo o que o app nem lê — e o LABOR e o
   DIARISTAS reescritos com o resultado do plano. Exportar só duas abas
   soltas obrigava a colar de volta à mão, e colar à mão é onde a
   correção volta a virar erro.

   Por isso o arquivo cru fica guardado na carga: ExcelJS o abre inteiro,
   as duas abas são reescritas EM CIMA das próprias, herdando o estilo da
   linha que já estava lá, e as abas de documentação entram no fim. */
function estiloDaLinha(ws, linha, largura){
  const est = [];
  const r = ws.getRow(linha);
  for(let i=1;i<=largura;i++){
    const c = r.getCell(i);
    est.push({ font:c.font, fill:c.fill, border:c.border, alignment:c.alignment, numFmt:c.numFmt });
  }
  return est;
}

/** Troca os dados de uma aba mantendo o cabeçalho e o visual dela.
 *
 *  Escreve NO LUGAR, célula a célula, em vez de apagar as linhas e
 *  recriá-las: `spliceRows` não remove nada nesta fatura — as linhas
 *  carregam fórmula, e a remoção sai silenciosamente sem efeito, o que
 *  fazia o arquivo sair com o Labor DUPLICADO (170 viravam 340). Escrever
 *  por cima ainda tem a vantagem de cada célula guardar o próprio
 *  formato: copiar o formato de uma linha-modelo transformava número em
 *  data em coluna com máscara de data.
 *
 *  Atribuir um valor a uma célula com fórmula troca a fórmula pelo valor,
 *  que é o que se quer — o arquivo tem de sair com o resultado do plano,
 *  não com a conta que produzia o número antigo. */
function reescreverAba(ws, linhasCabecalho, dados, largura){
  const est = estiloDaLinha(ws, linhasCabecalho + 1, largura);
  const antigas = Math.max(0, ws.rowCount - linhasCabecalho);
  const total = Math.max(antigas, dados.length);
  for(let i=0;i<total;i++){
    const linha = ws.getRow(linhasCabecalho + 1 + i);
    const vals = dados[i];
    const nova = i >= antigas;
    for(let c=1;c<=largura;c++){
      const cel = linha.getCell(c);
      const v = vals ? vals[c-1] : null;
      /* Data vira SERIAL, nunca objeto Date: o ExcelJS grava Date pelo
         fuso local — escrito de São Paulo, uma data de meia-noite UTC
         saía como 03:00:28 do mesmo dia, e num fuso a leste sairia no
         dia anterior. O serial não tem fuso, e a célula já traz o
         formato de data do próprio arquivo. */
      cel.value = (v === undefined || v === "") ? null
                : (v instanceof Date ? dataParaSerial(v) : v);
      if(nova){
        const e = est[c-1];
        if(e){
          if(e.font) cel.font = e.font;
          if(e.fill) cel.fill = e.fill;
          if(e.border) cel.border = e.border;
          if(e.alignment) cel.alignment = e.alignment;
          if(e.numFmt) cel.numFmt = e.numFmt;
        }
        if(v instanceof Date && !cel.numFmt) cel.numFmt = "dd/mm/yyyy";
      }
    }
  }
}

async function exportarEqualizado(){
  const sim = S.sim, plano = S.plano, fat = S.fatura;
  if(!sim || !plano || plano.erro || !fat) return;
  if(!fat.buffer){ falhar($("sm-err"), "Recarregue a fatura: o arquivo original é necessário "
    + "para devolver a planilha inteira com as abas trocadas."); return; }
  const lay = fat.layout, C = lay.cols;
  const dentro = simClassificarLinhas(fat.linhas).dentro;
  const acaoDe = new Map();
  plano.acoes.forEach(a => acaoDe.set(dentro[a.id], a));
  const dt = v => isValidYmd(v) ? ymdParaData(v) : null;

  /* ---- LABOR: o quadro fixo depois do plano ---- */
  const linhasLabor = [];
  let mantidas = 0, removidas = 0, ajustadas = 0;
  for(const l of fat.linhas){
    const a = acaoDe.get(l);
    if(!a){ mantidas++; linhasLabor.push(l.bruta.slice()); continue; }
    const segs = segmentosApos(l, a);
    if(!segs.length){ removidas++; continue; }
    ajustadas++;
    for(const seg of segs){
      const row = l.bruta.slice();
      if(C.inicio >= 0) row[C.inicio] = dt(seg.ini) || "";
      if(C.fim >= 0)    row[C.fim]    = dt(seg.fim) || "";
      linhasLabor.push(row);
    }
  }

  /* ---- DIARISTAS: as da fatura menos as cortadas, mais as novas ---- */
  const layD = fat.layoutDiarias;
  const incluir = $("sm-eqIncluir");
  const inc = (incluir && incluir.checked) ? plano.inclusoes : null;
  const escala = escalaDoDiarista(fat.unidade);

  /* Cada corte tira UMA pessoa-dia; casar por (groot, data) e consumir
     um por vez evita apagar duas linhas quando a fatura repete a pessoa
     no mesmo dia. */
  const aCortar = new Map();
  for(const c of (plano.corte ? plano.corte.cortes : []))
    aCortar.set(c.groot+"|"+c.data, (aCortar.get(c.groot+"|"+c.data) || 0) + 1);

  const linhasDiar = [];
  let cortadas = 0;
  for(const d of (fat.diarias || [])){
    const k = simGrootNum(d.groot)+"|"+d.data;
    if(aCortar.get(k) > 0){ aCortar.set(k, aCortar.get(k)-1); cortadas++; continue; }
    linhasDiar.push(d.bruta ? d.bruta.slice() : []);
  }
  let novas = 0;
  if(inc && layD){
    const CD = layD.cols;
    const constante = idx => {
      if(idx < 0) return "";
      let achou;
      for(const d of (fat.diarias || [])){
        const x = d.bruta ? d.bruta[idx] : "";
        if(x === "" || x === null || x === undefined) continue;
        const k = x instanceof Date ? +x : x;
        if(achou === undefined) achou = { k, x };
        else if(achou.k !== k) return "";
      }
      return achou ? achou.x : "";
    };
    const fixas = CD.fixas.map(i => [i, constante(i)]);
    for(const p of inc.pessoas) for(const d of p.dias){
      const row = new Array(layD.largura).fill("");
      fixas.forEach(([i,v]) => { if(i >= 0) row[i] = v; });
      if(CD.groot >= 0)  row[CD.groot]  = numSePuder(p.groot);
      if(CD.nome >= 0)   row[CD.nome]   = p.nome;
      if(CD.cargo >= 0)  row[CD.cargo]  = "Diarista";
      if(CD.escala >= 0) row[CD.escala] = escala;
      if(CD.data >= 0)   row[CD.data]   = ymdParaData(d);
      if(CD.qtd >= 0)    row[CD.qtd]    = 1;
      if(CD.obs >= 0)    row[CD.obs]    = "INCLUIDO PELA EQUALIZACAO - diarista "
        + (p.solic === "id" ? "ID" : p.solic === "meli" ? "do cliente" : "sem solicitante")
        + " no SIGO";
      linhasDiar.push(row); novas++;
    }
  }

  /* ---- o dossiê do que foi feito ---- */
  const eqLinhas = plano.acoes.map(a => [
    EQ_ROTULO[a.tipo], numSePuder(a.linha.groot), a.linha.nome, a.linha.cargo,
    dt(a.linha.inicio), dt(a.linha.fim), dt(a.novoInicio), dt(a.novoFim),
    a.pausas.map(p => fmtYmd(p.fim)+" a "+fmtYmd(p.ini)).join(" · "),
    -a.impacto.hc, a.impacto.dias, a.motivo]);
  for(const c of (plano.corte ? plano.corte.cortes : []))
    eqLinhas.push(["Retirar diária", numSePuder(c.groot), c.nome, "Diarista",
      dt(c.data), dt(c.data), null, null, "", -c.quantidade, 1,
      "Diária lançada acima do QF num dia em que o quadro fixo já estava no teto. "
      + "Sai a interna primeiro; a do cliente é a última ("
      + (c.solic === "id" ? "interna" : c.solic === "meli" ? "do cliente" : "sem solicitante no SIGO")+")."]);
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
    revLinhas.push(["Excesso mantido", dt(+d), q,
      plano.decisaoDiarias === "manter" && diariasDe[+d] >= q
        ? "O quadro fixo já está no teto e o que passa do QF são as "+diariasDe[+d]
          + " diária(s) já alocadas neste dia. Mantidas por decisão do usuário."
        : "Nem as quatro ações sobre o quadro fixo nem a redução de diária resolveram este dia."]);
  if(inc) for(const [d,c] of Object.entries(inc.dias).sort((a,b) => a[0]-b[0]))
    { if(c.descoberto > 0) revLinhas.push(["Falta descoberta", dt(+d), c.descoberto,
      "Não havia diarista livre bastante no SIGO neste dia ("+c.disponiveis+" disponível(is))."]); }
  if(!inc) for(const [d,q] of Object.entries(plano.falta).sort((a,b) => a[0]-b[0]))
    revLinhas.push(["Falta", dt(+d), q, "Quadro abaixo do QF; nenhuma inclusão foi gerada."]);
  if(!revLinhas.length) revLinhas.push(["(nada a revisar)"]);

  const fechados = plano.dias.filter(d => d.difPos === 0).length;
  const metaLinhas = [
    ["Arquivo", "Fatura equalizada pela Validação Template — as abas "+fat.aba+" e "
      + (layD ? layD.aba : "DIARISTAS")+" foram reescritas; as demais saíram como estavam"],
    ["Fatura de origem", S.nomeF],
    ["Unidade", fat.unidade || "—"],
    ["Período", fmtYmd(plano.periodo.ini)+" a "+fmtYmd(plano.periodo.fim)],
    ["Alvo", isFinite(plano.alvo) ? "QF fixo de "+plano.alvo+" HC por dia"
                                   : "S&OP diário da planilha operacional"],
    ["Correções de cadastro aplicadas", S.correcoes.length],
    ["Dias que fecham no QF", fechados+" de "+plano.dias.length],
    ["Permitir adiar início", plano.opcoes.permitirAdiarInicio ? "sim" : "não"],
    ["Pausar a partir de", plano.opcoes.pausaDesde ? fmtYmd(plano.opcoes.pausaDesde) : "não"],
    ["Diárias acima do QF", plano.decisaoDiarias === "reduzir"
      ? "REDUZIR — decisão do usuário; retirado só o necessário para chegar ao QF"
      : "MANTER — decisão do usuário; o excesso residual sai declarado em REVISAR"],
    ["", ""],
    ["Linhas do LABOR mantidas", mantidas],
    ["Linhas do LABOR ajustadas", ajustadas],
    ["Linhas do LABOR removidas", removidas],
    ["Diárias da fatura mantidas", linhasDiar.length - novas],
    ["Diárias retiradas por estar acima do QF", cortadas],
    ["  · internas", plano.corte ? plano.corte.totais.id : 0],
    ["  · do cliente", plano.corte ? plano.corte.totais.meli : 0],
    ["  · sem solicitante no SIGO", plano.corte ? plano.corte.totais.sem : 0],
    ["Diárias acrescentadas para cobrir falta", novas],
    ["  · da ID", inc ? inc.totais.id : 0],
    ["  · do cliente", inc ? inc.totais.meli : 0],
    ["Pessoa-dia conferidos contra o SIGO", inc ? inc.totais.verificados : 0],
    ["  · recusados por não estar na base no dia", inc ? inc.totais.recusados : 0],
    ["Falta descoberta (pessoa-dia)", inc ? inc.totais.descoberto : 0],
    ["Excesso antes do plano (HC-dia)", plano.totais.excessoAntes],
    ["Excesso depois do plano (HC-dia)", plano.totais.excessoDepois],
    ["", ""],
    ["AVISO", "As alterações equalizam MATEMATICAMENTE o quadro ao QF. Confirme se "
      + "correspondem à movimentação operacional real antes de enviar."],
    ["O que é o quadro do dia", "Quadro fixo do LABOR MAIS as diárias da aba de diaristas. "
      + "É esse total que vai ao confronto com o QF."],
    ["Ordem das ações", "Primeiro as quatro do motor sobre o quadro FIXO — retirar linha, adiar "
      + "início, pausar/retomar, antecipar fim —, nenhuma delas criando falta em outro dia. O que "
      + "sobrar de excesso depois disso não é fixo: é diária acima do QF, e sai por último."],
    ["Como a inclusão foi escolhida", "Base SIGO: pessoa solicitada NAQUELE DIA e sem cobrança "
      + "no dia — nem no LABOR, nem como diária já lançada. Prioridade para os diaristas da ID; "
      + "o do cliente só entra depois de esgotados os internos. Cada pessoa-dia é reconferido "
      + "contra a base antes de virar linha."],
    ["Campos em branco", "Matrícula, regime, turno e valores das linhas acrescentadas não são "
      + "deduzíveis do SIGO e por isso não são escritos. A ESCALA é a da operação da unidade."]];

  /* ---- monta o arquivo por cima do original ---- */
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fat.buffer);
  /* O ExcelJS reemite as validações de dados com FAIXAS SOBREPOSTAS —
     `K2:K3000` e `K10:K3000` na mesma coluna, três regras virando cinco.
     Faixa sobreposta é uma das violações que fazem o Excel abrir o
     arquivo pedindo para "recuperar". Acontece no round-trip puro, sem
     edição nenhuma, então não há o que ajustar na escrita: o jeito é
     não levar as validações adiante. Perde-se a lista suspensa de duas
     colunas; ganha-se um arquivo que abre. */
  wb.eachSheet(ws => {
    if(ws.dataValidations) ws.dataValidations.model = {};
    /* E a formatação condicional sai pelo mesmo motivo: as regras do
       original vivem num `extLst` que o ExcelJS não carrega, e ele grava
       o `<conditionalFormatting sqref="A6"/>` VAZIO. Elemento sem uma
       regra dentro é inválido pelo esquema, e o Excel repara o arquivo
       por causa dele. */
    if(ws.conditionalFormattings) ws.conditionalFormattings = [];
  });
  const wsL = wb.getWorksheet(fat.aba);
  if(wsL) reescreverAba(wsL, lay.topo.length, linhasLabor, lay.largura);
  if(layD){
    const wsD = wb.getWorksheet(layD.aba);
    if(wsD) reescreverAba(wsD, layD.topo, linhasDiar, layD.largura);
  }
  for(const nome of ["EQUALIZACAO","INCLUSOES","REVISAR","CADASTRO","METADADOS"])
    if(wb.getWorksheet(nome)) wb.removeWorksheet(wb.getWorksheet(nome).id);
  abaEstilizada(wb, "EQUALIZACAO", { ...TPL_DOC,
    colunas:["Ação","GROOT ID","Nome","Cargo","Início atual","Fim atual","Início novo",
             "Fim novo","Pausas","HC","Dias","Motivo"],
    larguras:[16,12,34,26,12,12,12,12,26,7,7,90] }, eqLinhas);
  abaEstilizada(wb, "INCLUSOES", { ...TPL_DOC,
    colunas:["GROOT ID","Nome","Solicitante","De","Até","Dias","Observação"],
    larguras:[12,34,16,12,12,7,52] }, incLinhas);
  if(S.correcoes.length){
    const cadLinhas = S.correcoes.map(c =>
      [c.nome, c.de || "(sem GROOT)", c.para, c.linhas,
       "Confirmada pelo usuário contra a base oficial de GROOT"]);
    abaEstilizada(wb, "CADASTRO", { ...TPL_DOC,
      colunas:["Nome","GROOT anterior","GROOT corrigido","Linhas alteradas","Observação"],
      larguras:[34,16,16,14,52] }, cadLinhas);
  }
  abaEstilizada(wb, "REVISAR", { ...TPL_DOC,
    colunas:["Tipo","Data","Quantidade","Observação"], larguras:[22,12,12,72] }, revLinhas);
  abaEstilizada(wb, "METADADOS", { ...TPL_DOC,
    colunas:["Campo","Valor"], larguras:[40,110] }, metaLinhas);

  const per = String(fat.comp ? fat.comp.y*100 + fat.comp.m : "");
  /* Acento no `download` de um blob: faz o Chromium descartar o nome. */
  const nome = semAcento("Fatura_equalizada_"+(fat.unidade || "fatura")+"_"+per+".xlsx")
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
bindDrop("sm-dropG","sm-fileG",(f,d) => carregar("groot",f,d));
[$("sm-eqAdiar"), $("sm-eqPausa"), $("sm-eqIncluir")].forEach(el =>
  { if(el) el.onchange = () => { if(S.sim) render(); }; });
const btnRun = $("sm-btnRun"); if(btnRun) btnRun.onclick = analisar;

})();
