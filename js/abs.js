/* Calcular ABS — validador de absenteísmo por filial
   ================================================================

   Confronta o quadro S&OP diário (sem over) com quem esteve presente,
   e compensa o déficit de cada dia com os diaristas SOLICITADOS PELA
   ID LOGISTICS naquele dia — a agência não importa, e o diarista pedido
   pelo MELI fica fora: é do MELI, não cobertura do quadro. O abate
   é limitado ao próprio déficit, de modo que um dia nunca fica
   "positivo" por excesso de diarista.

   Todo o módulo vive dentro de uma IIFE: esta aba define uma função
   `render`, e a aba de Conciliação também — a primeira roda no escopo
   global. Os `id` levam o prefixo `abs-` pelo mesmo motivo das demais
   abas: todas coexistem no mesmo documento.

   Não há handler inline chamando função deste módulo, então nada
   precisa ser publicado em window.
   ================================================================ */
(function(){
"use strict";
/* ============================== estado ============================== */
const S = { mesA:null, mesB:null, sigo:null, sop:{}, sopOrigem:{}, sopArquivo:[],
            sopFiles:[], results:null, compSrc:'id' };
/* Quais diaristas podem compensar o absenteísmo, escolhido no diálogo antes
   de validar. Empty (sem solicitante) fica sempre fora: não dá para atribuir. */
const COMP_LABEL = {id:'ID Logistics', meli:'MELI', ambos:'ID Logistics + MELI'};
function aceitaSolic(s){
  if(S.compSrc==='ambos') return s==='id' || s==='meli';
  return s===S.compSrc;
}
const MESES = {jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12};
const MES_NOME = ['','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const MES_LONGO = {janeiro:1,fevereiro:2,marco:3,abril:4,maio:5,junho:6,julho:7,agosto:8,setembro:9,outubro:10,novembro:11,dezembro:12};
const FILIAIS = {PA:'Pouso Alegre',VG:'Varginha',VRG:'Varginha',PC:'Poços de Caldas',DV:'Divinópolis',PM:'Patos de Minas'};
const OPER = {SVC:'Service',XD:'XD',SD:'Same Day',FULL:'Full'};
const DOW = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
const norm = s => String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();

/* ============================== leitura ============================== */
async function readWB(file){
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  return wb;
}
function cellVal(ws,r,c){ const v = ws.getRow(r).getCell(c).value;
  if (v && typeof v==='object'){ if (v.result!==undefined) return v.result; if (v.richText) return v.richText.map(t=>t.text).join(''); if (v instanceof Date) return v; if (v.text!==undefined) return v.text; }
  return v;
}
/* abas de operação das planilhas mensais. Dois padrões de nome:
     "PAXD Jul"  / "DV Ago"           → um mês (formato antigo)
     "PASVC 16.07-15.08"              → intervalo que cruza dois meses
   O código (PASVC, DV, PM…) é sempre a parte alfabética inicial; o resto diz
   o(s) mês(es). Devolve `meses` como ARRAY porque uma aba pode cobrir dois. */
function parseTabName(name){
  const raw = name.trim();
  const mCode = /^([A-Za-z]+)/.exec(raw);
  if(!mCode) return null;
  let code = mCode[1].toUpperCase(), op=null;
  for(const suf of ['SVC','FULL','XD','SD']){
    if(code.endsWith(suf) && code.length>suf.length){ op=suf; code=code.slice(0,-suf.length); break; }
  }
  if(!FILIAIS[code]) return null;
  const resto = raw.slice(mCode[1].length).trim();
  let meses=null;
  const mMes = /^([A-Za-zç]{3,})/.exec(resto);          // "Jul", "Agosto"
  if(mMes){ const mm=MESES[norm(mMes[1]).slice(0,3)]; if(mm) meses=[mm]; }
  if(!meses){                                            // "16.07-15.08" / "16/07 a 15/08"
    const mr=/(\d{1,2})[.\/](\d{1,2}).*?(\d{1,2})[.\/](\d{1,2})/.exec(resto);
    if(mr){ const m1=+mr[2], m2=+mr[4]; if(m1>=1&&m1<=12&&m2>=1&&m2<=12) meses=m1===m2?[m1]:[m1,m2]; }
  }
  if(!meses) return null;
  return { fil:code, op, meses, tab:name };
}
/* Mês/ano de cada coluna do bloco S&OP, a partir do título:
     "· Julho/2026"            → um mês
     "· 16/07 a 15/08/2026"    → intervalo (o ano do fim; o início recua um ano
                                 se o mês inicial for maior, cruzando dezembro) */
function mesesDoTitulo(titulo){
  const t=norm(titulo);
  let m=/(\d{1,2})[\/.](\d{1,2})\s*(?:a|ate|-)\s*(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})/.exec(t);
  if(m){ const m1=+m[2], m2=+m[4], y=+m[5];
    return {tipo:'range', ini:{m:m1, y:m1<=m2?y:y-1}, fim:{m:m2, y}}; }
  const monMap={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12,...MES_LONGO};
  for(const mm of t.matchAll(/\b([a-z]{3,})[\s\-\/]?(?:de\s*)?(\d{4})\b/g)){
    const mk=Object.keys(monMap).find(k=>mm[1].startsWith(k));
    if(mk) return {tipo:'mes', ini:{m:monMap[mk], y:+mm[2]}, fim:{m:monMap[mk], y:+mm[2]}};
  }
  return null;
}
/* Onde estão as coisas dentro de uma aba de operação.
   ----------------------------------------------------------------
   O modelo novo põe, ANTES da grade de pessoas, um bloco de headcount:

     linha  1  S&OP — HEADCOUNT DIÁRIO · Pouso Alegre · Julho/2026
     linha  2  Métrica  …  1 (Qua) | 2 (Qui) | …
     linha  3  Esperado …      88  |     88  | …      ← o S&OP do dia
     …
     linha 10  COLABORADORES — PRESENÇAS E FALTAS
     linha 11  Mat. | Groot ID | Nome | … | 1 Qua | 2 Qui | …
     linha 12+ as pessoas

   No modelo antigo a grade começa direto na linha 1 e não há S&OP. Por
   isso nada aqui é posição fixa: procura-se a linha "Mat./Nome" e, se
   existir, a linha "Esperado" com o cabeçalho de dias logo acima. */
function localizarBlocos(ws){
  const lim=Math.min(ws.rowCount||1, 30);
  let pessoas=0, sop=0, contratado=0, sopHeader=0;
  for(let r=1;r<=lim;r++){
    const a=norm(cellVal(ws,r,1)), c=norm(cellVal(ws,r,3));
    if(!pessoas && /^mat\.?$/.test(a) && c.startsWith('nome')) pessoas=r;
    if(!sop && a==='esperado') sop=r;
    if(!contratado && a==='contratado') contratado=r;
  }
  if(!pessoas) return null;
  const diasDe=(r,re)=>{
    const cols={};
    for(let c=9;c<=ws.columnCount;c++){
      const m=re.exec(String(cellVal(ws,r,c)??'').trim());
      if(m) cols[+m[1]]=c;
    }
    return cols;
  };
  const diasPessoas=diasDe(pessoas, /^(\d{1,2})(\s|\n|$)/);
  if(Object.keys(diasPessoas).length<3) return null;
  let diasSop={};
  if(sop){
    // o cabeçalho de dias do bloco S&OP é a linha logo acima da "Esperado"
    for(let r=sop-1;r>=1 && !Object.keys(diasSop).length;r--) diasSop=diasDe(r, /^(\d{1,2})\s*\(/);
    if(Object.keys(diasSop).length<3){ sop=0; diasSop={}; }
    else sopHeader=1;
  }
  return { pessoas, diasPessoas, sop, contratado, diasSop, titulo:String(cellVal(ws,1,1)??'') };
}
function validaAba(ws, layout){
  if(!layout) return 'cabeçalho fora do padrão (não achei a linha "Mat./Nome" com colunas de dias a partir da coluna I)';
  return null;
}
function lerAba(ws, layout){
  const dias=layout.diasPessoas, h=layout.pessoas;
  const rows=[];
  ws.eachRow((row,rn)=>{
    if(rn<=h) return;
    const mat=row.getCell(1).value, nome=row.getCell(3).value;
    if(mat==null && nome==null) return;
    const rec={ info:[], dias:{} };
    for(let c=1;c<=8;c++){ let v=row.getCell(c).value; if(v&&typeof v==='object'&&v.richText) v=v.richText.map(t=>t.text).join(''); rec.info.push(v??null); }
    for(const [d,c] of Object.entries(dias)){ let v=row.getCell(+c).value; if(v!=null&&v!=='') rec.dias[+d]=String(v).trim(); }
    rows.push(rec);
  });
  return { rows };
}
/* Lê o bloco de headcount embutido numa aba de operação: a linha "Esperado"
   (o S&OP do dia) e, quando existir, a "Contratado" — dado de origem, não
   algo que dê para derivar da grade de presenças (contar marcas não-DF dava
   179 num dia em que a planilha diz 137).

   Cada COLUNA é resolvida ao seu ano-mês pelo título: numa aba de um mês só
   todas caem no mesmo ym; numa aba de intervalo (16/07 a 15/08) as colunas
   16..31 são de julho e 1..15 de agosto — a virada é detectada quando o dia
   cai em relação à coluna anterior. Devolve um bloco por ym. */
function lerSopDaAba(ws, layout){
  if(!layout || !layout.sop) return null;
  const info=mesesDoTitulo(layout.titulo);
  if(!info) return {erro:'sem mês/ano legível no título (ex.: "· Julho/2026" ou "16/07 a 15/08/2026")'};
  const diasOrdem=Object.keys(layout.diasSop).map(Number).sort((a,b)=>layout.diasSop[a]-layout.diasSop[b]);
  const ymDe=new Map();
  let cur={...info.ini}, prev=-Infinity;
  for(const d of diasOrdem){
    if(d<prev && info.tipo==='range') cur={m:info.fim.m, y:info.fim.y};
    ymDe.set(d, cur.y*100+cur.m); prev=d;
  }
  const linha=r=>{
    const o={};
    for(const d of diasOrdem){
      const v=cellVal(ws, r, layout.diasSop[d]);
      if(typeof v==='number'){ const ym=ymDe.get(d); (o[ym]??={})[d]=v; }
    }
    return o;
  };
  const esp=linha(layout.sop);
  if(!Object.keys(esp).length) return {erro:'linha "Esperado" sem nenhum valor numérico'};
  const con=layout.contratado? linha(layout.contratado) : {};
  return { blocos:Object.keys(esp).map(ym=>({ym:+ym, dias:esp[ym], contratado:con[ym]||{}})) };
}
/* headcount S&OP — flexível. Aceita:
   (a) 1 aba por filial, título "· Agosto/2026", colunas "1 SAB … 31 SEG"  → mês único
   (b) 1 aba por filial+mês, nome "Pouso Alegre Jul-26", título com intervalo
       "16/07/2026 a 31/07/2026", colunas "16 QUI … 15 SEX" (podem cruzar meses)
   Sempre lê apenas a linha "S&OP" (sem over) de cada seção.
   Resolve o mês/ano de cada COLUNA, não da aba, para suportar intervalos que cruzam meses. */
function acharFilial(nome){
  const n=norm(nome);
  return Object.keys(FILIAIS).find(k=>norm(FILIAIS[k])===n)
      ?? Object.keys(FILIAIS).find(k=>n.includes(norm(FILIAIS[k])));
}
/* deriva, para cada dia-do-mês presente nas colunas, qual ano-mês (ym) ele representa,
   a partir de pistas do título/nome da aba. Retorna Map(dia -> ym) e lista ordenada. */
function resolverMeses(titulo, nomeAba, diasOrdem){
  const t=norm(titulo)+' '+norm(nomeAba);
  const ranges=[]; // [{y,m}] na ordem em que aparecem
  // formato dd/mm/yyyy (título com intervalo)
  let mDate=[...t.matchAll(/(\d{1,2})\/(\d{1,2})\/(\d{4})/g)];
  for(const m of mDate) ranges.push({d:+m[1], m:+m[2], y:+m[3]});
  // formato "mmm-aa" ou "mmm/aa" ou "mmm de aaaa" ou "Mês/AAAA"
  const monMap={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12,...MES_LONGO};
  let mMon=[...t.matchAll(/\b([a-z]{3,})[\s\-\/]?(?:de\s*)?(\d{2,4})\b/g)];
  for(const m of mMon){ const mk=Object.keys(monMap).find(k=>m[1].startsWith(k)); if(mk){ let y=+m[2]; if(y<100)y+=2000; ranges.push({m:monMap[mk], y}); } }

  const dias=diasOrdem.slice(); // já em ordem de coluna
  const map=new Map();
  if(ranges.length>=1 && ranges[0].m){
    // percorre os dias em ordem; quando o dia "reinicia" (cai vs anterior) avança para o próximo mês do range
    let ri=0; let prev=-Infinity;
    let cur={y:ranges[0].y, m:ranges[0].m};
    for(const d of dias){
      if(d<prev){ // virou o mês
        ri++;
        if(ranges[ri]&&ranges[ri].m){ cur={y:ranges[ri].y, m:ranges[ri].m}; }
        else { cur={...cur}; cur.m++; if(cur.m>12){cur.m=1;cur.y++;} }
      }
      map.set(d, cur.y*100+cur.m);
      prev=d;
    }
    return map;
  }
  return null; // sem pista de mês
}
function parseHeadcount(wb, fname){
  const out=[]; const erros=[];
  for(const ws of wb.worksheets){
    const filCode=acharFilial(ws.name);
    if(!filCode) continue;
    const titulo=String(cellVal(ws,1,1)??'');
    // localizar a linha de cabeçalho de dias (procura nas primeiras 4 linhas)
    let hRow=0, diasCol={};
    for(let rr=1;rr<=4 && !hRow;rr++){
      const cols={};
      for(let c=2;c<=ws.columnCount;c++){
        const h=String(cellVal(ws,rr,c)??'').trim();
        const dm=/^(\d{1,2})\s+[A-Za-zçÇ]{3}/.exec(h); // "16 QUI", "1 SAB"
        if(dm) cols[+dm[1]]=c;
      }
      if(Object.keys(cols).length>=3){ hRow=rr; diasCol=cols; }
    }
    if(!hRow){ erros.push(`aba "${ws.name}": não achei colunas de dias ("16 QUI, 17 SEX…")`); continue; }
    const diasOrdem=Object.entries(diasCol).sort((a,b)=>a[1]-b[1]).map(([d])=>+d);
    const mesPorDia=resolverMeses(titulo, ws.name, diasOrdem);
    if(!mesPorDia){ erros.push(`aba "${ws.name}": não identifiquei o mês (título sem "Mês/Ano" nem intervalo dd/mm/aaaa, e nome da aba sem "Jul-26")`); continue; }

    let secao='';
    for(let r=hRow+1;r<=ws.rowCount;r++){
      const a=String(cellVal(ws,r,1)??'').trim();
      if(!a) continue;
      const an=norm(a);
      if(an==='s&op'){
        // agrupa valores por ym (a mesma linha pode conter dias de 2 meses)
        const porYm={};
        for(const [d,c] of Object.entries(diasCol)){
          const v=cellVal(ws,r,+c);
          if(typeof v!=='number') continue;
          const ym=mesPorDia.get(+d); if(!ym) continue;
          (porYm[ym]??={})[+d]=v;
        }
        for(const [ym,dias] of Object.entries(porYm))
          out.push({ filCode, secao: secao||'TOTAL', ym:+ym, dias, aba:ws.name });
      } else if(!/^s&op\s*c\//.test(an) && an!=='contratado' && !an.startsWith('dif')){
        secao=a;
      }
    }
  }
  if(!out.length) return { erro: erros[0] || 'nenhuma aba de filial reconhecida' };
  return { blocos: out, avisos: erros };
}

/* Remonta S.sop a partir das duas fontes possíveis, nesta ordem:
   1. o S&OP embutido nas próprias planilhas mensais (modelo novo);
   2. o arquivo de headcount avulso, que sobrepõe — carregá-lo é um ato
      explícito de quem quer justamente aquele número.
   A origem fica registrada por seção para o seletor dizer de onde veio. */
function aplicarSopDosMeses(){
  S.sop={}; S.sopOrigem={}; S.sopContratado={};
  const por=(b,origem)=>{
    const dest=(((S.sop[b.filCode]??={})[b.secao]??={})[b.ym]??={});
    Object.assign(dest, b.dias);
    ((S.sopOrigem[b.filCode]??={})[b.secao]??=new Set()).add(origem);
    if(b.contratado && Object.keys(b.contratado).length){
      const dc=(((S.sopContratado[b.filCode]??={})[b.secao]??={})[b.ym]??={});
      Object.assign(dc, b.contratado);
    }
  };
  for(const slot of ['mesA','mesB']) (S[slot]?.sop??[]).forEach(b=>por(b,'mensal'));
  (S.sopArquivo??[]).forEach(b=>por(b,'arquivo'));
}
/* de onde veio a seção, para rotular o seletor sem mentir */
function rotuloFonte(fil, secao){
  const o=S.sopOrigem?.[fil]?.[secao];
  if(!o) return 'S&OP';
  if(o.has('arquivo') && o.has('mensal')) return 'S&OP (planilha mensal + headcount)';
  return o.has('arquivo')? 'Headcount' : 'S&OP da planilha mensal';
}

/* ============================== upload ============================== */
function bindDrop(dropId, fileId, fn){
  const drop=document.getElementById(dropId), inp=document.getElementById(fileId);
  drop.addEventListener('click',()=>inp.click());
  drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over')});
  drop.addEventListener('dragleave',()=>drop.classList.remove('over'));
  drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('over'); if(e.dataTransfer.files.length) fn(e.dataTransfer.files,drop);});
  inp.addEventListener('change',()=>{ if(inp.files.length) fn(inp.files,drop); });
}
async function handleMes(slot,files,drop){
  const fn=drop.querySelector('.fname'), st=drop.querySelector('.st');
  const file=files[0];
  fn.textContent=file.name; st.textContent='Lendo…'; st.className='st';
  drop.classList.remove('loaded','err');
  try{
    const wb=await readWB(file);
    const tabs=[]; const warns=[]; const sop=[];
    for(const ws of wb.worksheets){
      const p=parseTabName(ws.name); if(!p) continue;
      const layout=localizarBlocos(ws);
      const err=validaAba(ws, layout);
      if(err){ warns.push(ws.name+': '+err); continue; }
      tabs.push({...p, ws, layout});
      /* O S&OP agora mora na própria aba da operação. Cada aba é UMA
         operação, então ela vira uma seção com o nome do turno — e uma
         filial sem turno (DV, PM) vira a seção TOTAL. É a mesma forma que
         o arquivo de headcount produz, então tudo a jusante segue igual. */
      const s=lerSopDaAba(ws, layout);
      if(s && s.erro) warns.push(ws.name+' (S&OP): '+s.erro);
      else if(s) for(const bl of s.blocos) sop.push({filCode:p.fil, secao:p.op??'TOTAL',
                           ym:bl.ym, dias:bl.dias, contratado:bl.contratado, aba:ws.name});
    }
    if(!tabs.length) throw new Error('padrão diferente — nenhuma aba de operação reconhecida (ex.: PAXD Jul, PASVC 16.07-15.08)');
    const meses=[...new Set(tabs.flatMap(t=>t.meses))].sort((a,b)=>a-b);
    S[slot]={ file:file.name, tabs, meses, sop };
    aplicarSopDosMeses();
    st.textContent='✓ '+tabs.length+' operações · mês: '+meses.map(m=>MES_NOME[m]).join(', ')
      + (sop.length? ' · S&OP de '+sop.length+' aba(s)':'')
      + (warns.length? ' · ⚠ '+warns.length+' aviso(s)':'');
    st.className='st '+(warns.length?'warn':'ok'); drop.classList.add('loaded');
    if(warns.length) console.warn('Avisos da leitura:',warns);
  }catch(e){ st.textContent='✗ '+e.message; st.className='st bad'; drop.classList.add('err'); S[slot]=null; }
  refreshOps(); checkReady();
}
async function handleSigo(files,drop){
  const fn=drop.querySelector('.fname'), st=drop.querySelector('.st');
  const file=files[0];
  fn.textContent=file.name; st.textContent='Lendo…'; st.className='st';
  drop.classList.remove('loaded','err');
  try{
    const wb=await readWB(file);
    const abas=wb.worksheets.filter(ws=>
      /groot/i.test(String(cellVal(ws,4,5)??'')) && /data/i.test(String(cellVal(ws,4,2)??'')));
    if(!abas.length) throw new Error('padrão diferente — nenhuma aba de filial no formato SIGO (cabeçalho na linha 4 com DATA SOLICITAÇÃO e GROOT ID)');
    const base={};
    for(const ws of abas){
      const regs=[];
      ws.eachRow((row,rn)=>{
        if(rn<5) return;
        let dt=row.getCell(2).value, id=row.getCell(5).value;
        if(dt&&typeof dt==='object'&&dt.result!==undefined) dt=dt.result;
        if(id&&typeof id==='object'&&id.result!==undefined) id=id.result;
        if(!(dt instanceof Date)) return;
        const idn=String(id??'').replace(/\D/g,'');
        if(!idn) return; // somente ID
        /* Quem pediu decide se o diarista compensa: o abate de ABS só pode
           usar diarista SOLICITADO PELA ID LOGISTICS — o do MELI é do
           MELI, não cobertura do quadro. A agência (MOURA, TSI…) não importa. */
        const solicTxt=norm(row.getCell(3).value);
        const solic=!solicTxt? '' : (solicTxt.includes('meli')? 'meli' : 'id');
        const txt=c=>{ let v=row.getCell(c).value;
          if(v&&typeof v==='object'){ if(v.result!==undefined)v=v.result; else if(v.richText)v=v.richText.map(t=>t.text).join(''); }
          return String(v??'').trim(); };
        // ExcelJS entrega datas em UTC: usar getters UTC para não deslocar 1 dia em UTC-3
        regs.push({ data:new Date(dt.getUTCFullYear(),dt.getUTCMonth(),dt.getUTCDate()), id:idn, solic,
                    solicitante:txt(3), empresa:txt(4), nome:txt(6), cargo:txt(7),
                    escala:txt(8), aba:ws.name });
      });
      base[ws.name]=regs;
    }
    S.sigo={ file:file.name, base };
    const todos=Object.values(base).flat();
    const nId=todos.filter(r=>r.solic==='id').length;
    const nMeli=todos.filter(r=>r.solic==='meli').length;
    const nSem=todos.length-nId-nMeli;
    st.textContent='✓ '+abas.length+' filiais · '+nId+' ID Logistics'
      +(nMeli?' · '+nMeli+' MELI':'')
      +(nSem?' · '+nSem+' sem solicitante':'')
      +' — a fonte da compensação é escolhida ao validar';
    st.className='st ok'; drop.classList.add('loaded');
  }catch(e){ st.textContent='✗ '+e.message; st.className='st bad'; drop.classList.add('err'); S.sigo=null; }
  checkReady();
}
async function handleSop(files,drop){
  const fn=drop.querySelector('.fname'), st=drop.querySelector('.st');
  st.textContent='Lendo…'; st.className='st'; drop.classList.remove('loaded','err');
  try{
    for(const file of files){
      const wb=await readWB(file);
      const r=parseHeadcount(wb,file.name);
      if(r.erro) throw new Error(file.name+': '+r.erro);
      (S.sopArquivo??=[]).push(...r.blocos);
      if(r.avisos&&r.avisos.length) console.warn('Headcount avisos:',r.avisos);
      if(!S.sopFiles.includes(file.name)) S.sopFiles.push(file.name);
    }
    aplicarSopDosMeses();
    const meses=[...new Set(Object.values(S.sop).flatMap(sec=>Object.values(sec).flatMap(m=>Object.keys(m))))]
      .map(Number).sort().map(ym=>MES_NOME[ym%100]+'/'+Math.floor(ym/100));
    fn.textContent=S.sopFiles.join(' · ');
    st.textContent='✓ S&OP de '+Object.keys(S.sop).length+' filiais · '+meses.join(', ');
    st.className='st ok'; drop.classList.add('loaded');
  }catch(e){ st.textContent='✗ '+e.message; st.className='st bad'; drop.classList.add('err'); }
  refreshOps(); checkReady();
}
bindDrop('abs-dropA','abs-fileA',(f,d)=>handleMes('mesA',f,d));
bindDrop('abs-dropB','abs-fileB',(f,d)=>handleMes('mesB',f,d));
bindDrop('abs-dropS','abs-fileS',handleSigo);
bindDrop('abs-dropH','abs-fileH',handleSop);

/* ============================== operações ============================== */
/* Abas de operação disponíveis, unindo os dois slots sem duplicar. Um arquivo
   de intervalo (16/07 a 15/08) já traz tudo num slot só; dois arquivos mensais
   trazem a mesma operação uma vez por mês, mas para LISTAR o cartão basta uma —
   o roster depois lê ambas. Dedup por filial|operação, guardando as abas que
   alimentam cada cartão. */
function blocosCarregados(){ return [S.mesA, S.mesB].filter(Boolean); }
function abasUnicas(){
  const vis=new Map();
  for(const bl of blocosCarregados()) for(const t of bl.tabs){
    const k=t.fil+'|'+(t.op??''); if(!vis.has(k)) vis.set(k, t);
  }
  return [...vis.values()];
}
/* monta a lista de cartões: cada operação + um cartão TOTAL por filial com 2+ operações */
function listarCartoes(){
  const cards=abasUnicas().map(t=>({ fil:t.fil, op:t.op??null, total:false }));
  const porFil={};
  for(const c of cards) (porFil[c.fil]??=[]).push(c);
  for(const [fil,list] of Object.entries(porFil))
    if(list.length>1) cards.push({ fil, op:null, total:true, opsIncluidas:list.map(c=>c.op) });
  return cards;
}
const cardKey=c=> c.total? c.fil+'|__TOTAL__' : c.fil+'|'+(c.op??'');
/* opções de fonte S&OP coerentes com o cartão:
   - operação: a seção equivalente (sugerida); demais seções entram como "outras"; TOTAL não entra
   - TOTAL/filial completa: TOTAL (sugerido); soma AM+PM quando existir; seções como "outras" */
function opcoesFonte(card){
  const secs=Object.keys(S.sop[card.fil]??{});
  const sug=[], outras=[];
  const semTotal=secs.filter(s=>norm(s)!=='total');
  const rot=s=>rotuloFonte(card.fil,s)+' · '+s;
  if(card.total || !card.op){
    const t=secs.find(s=>norm(s)==='total');
    if(t) sug.push({v:'hc:'+t, l:rotuloFonte(card.fil,t)+' · TOTAL filial (sem over)'});
    if(!t && semTotal.length) sug.push({v:'hcsum:'+semTotal.join(','),
      l:rotuloFonte(card.fil,semTotal[0])+' · soma das operações ('+semTotal.join('+')+')'});
    for(const s of semTotal) outras.push({v:'hc:'+s, l:rot(s)});
  } else {
    const m=semTotal.find(s=>norm(s)===norm(card.op));
    if(m) sug.push({v:'hc:'+m, l:rotuloFonte(card.fil,m)+' · '+m+' (sem over)'});
    const am=semTotal.find(s=>norm(s)==='am'), pm=semTotal.find(s=>norm(s)==='pm');
    if(!m && am && pm && norm(card.op)==='svc')
      sug.push({v:'hcsum:'+am+','+pm, l:rotuloFonte(card.fil,am)+' · AM+PM (soma)'});
    for(const s of semTotal) if(s!==m) outras.push({v:'hc:'+s, l:rot(s)});
  }
  return { sug, outras };
}
function refreshOps(){
  const grid=document.getElementById('abs-opsGrid');
  const prev={};
  grid.querySelectorAll('.opchip').forEach(ch=>{
    const sel=ch.querySelector('select');
    prev[ch.dataset.key]={ sel:ch.querySelector('input[type=checkbox]').checked,
      src:sel?.value, man:ch.querySelector('input[type=number]')?.value,
      /* A grade é montada assim que as duas planilhas mensais chegam, antes do
         headcount. Nesse momento "Manual" é a única opção — preservar essa
         escolha depois faria o app ignorar o S&OP recém-carregado e seguir com o
         valor de placeholder. Só a escolha feita COM alternativas disponíveis é
         digna de ser preservada. */
      tinhaAlternativa: !!sel && sel.options.length>1 };
  });
  if(!S.mesA&&!S.mesB){ grid.innerHTML='<span class="msg" style="color:var(--dim)">Carregue a planilha de absenteísmo para listar as operações.</span>'; return; }
  const cards=listarCartoes();
  if(!cards.length){ grid.innerHTML='<span class="msg bad">Nenhuma operação reconhecida nas planilhas.</span>'; return; }
  grid.innerHTML='';
  for(const card of cards){
    const key=cardKey(card);
    const p=prev[key];
    const {sug,outras}=opcoesFonte(card);
    const manterEscolha = p && p.src && (p.tinhaAlternativa || p.src!=='manual');
    const escolha = manterEscolha ? p.src : (sug[0]?.v ?? 'manual');
    const subt = card.total? 'TOTAL ('+card.opsIncluidas.join('+')+')'
               : card.op? OPER[card.op]+' ('+card.op+')' : 'filial completa';
    const div=document.createElement('label'); div.className='opchip'+(p?.sel?' sel':''); div.dataset.key=key;
    const opt=(o)=>`<option value="${o.v}" ${escolha===o.v?'selected':''}>${o.l}</option>`;
    div.innerHTML=`<div class="top"><input type="checkbox" data-key="${key}" ${p?.sel?'checked':''}>
      <span class="nm"><span class="f">${FILIAIS[card.fil]}</span><br><span class="o">${subt}</span></span></div>
      <div class="src"><span class="lb">Fonte S&amp;OP</span>
        <select data-src="${key}" onclick="event.stopPropagation()">
          ${sug.map(opt).join('')}
          ${outras.length?`<optgroup label="Outras seções (conferir)">${outras.map(opt).join('')}</optgroup>`:''}
          <option value="manual" ${escolha==='manual'?'selected':''}>Manual (seg–sex)</option>
        </select>
        <input type="number" min="0" value="${p?.man??33}" data-sop="${key}" onclick="event.stopPropagation()"
          style="display:${escolha==='manual'?'':'none'}">
      </div>`;
    div.querySelector('input[type=checkbox]').addEventListener('change',e=>{div.classList.toggle('sel',e.target.checked);checkReady();});
    div.querySelector('select').addEventListener('change',e=>{
      div.querySelector('input[type=number]').style.display = e.target.value==='manual'?'':'none';
    });
    grid.appendChild(div);
  }
}
function checkReady(){
  const any=[...document.querySelectorAll('#abs-opsGrid input[type=checkbox]')].some(c=>c.checked);
  // basta uma planilha de abs (a de intervalo já cobre tudo) + SIGO + seleção
  document.getElementById('abs-btnRun').disabled=!((S.mesA||S.mesB)&&S.sigo&&any);
}

/* ============================== cálculo ============================== */
function diasPeriodo(ini,fim){
  const out=[]; const d=new Date(ini);
  while(d<=fim){ out.push(new Date(d)); d.setDate(d.getDate()+1); }
  return out;
}
function sigoSheetsFor(fil,op,total){
  const alvoFil=norm(FILIAIS[fil]);
  const nomes=Object.keys(S.sigo.base);
  const doFil=nomes.filter(n=>norm(n).includes(alvoFil));
  if(!doFil.length) return [];
  if(total) return doFil; // TOTAL da filial: todas as abas dela (ex.: PA SVC + PA XD)
  /* Quando a filial marca a operação na coluna ESCALA, quem manda é a ESCALA,
     não a aba: as abas do SIGO se misturam (a aba "Pouso Alegre XD" carrega
     linhas SD e FULL) e o cartão SD nem tinha aba própria — caía na do SVC.
     Nesse caso varremos todas as abas da filial e filtramos por escala. */
  if(op && escalaDaOperacao(doFil,op)) return doFil;
  if(op){ const comOp=doFil.find(n=>norm(n).includes(norm(op))); if(comOp) return [comOp]; }
  const semOp=doFil.find(n=>!['svc','xd','sd','full'].some(s=>norm(n).endsWith(' '+s)));
  return [semOp ?? doFil[0]];
}
/* A filial usa este token de operação na coluna ESCALA? Só então dá para
   filtrar por ela. Varginha e Poços marcam turno (AM/PM) e Divinópolis marca
   horário — nessas, filtrar por "SVC" zeraria a compensação inteira. */
function escalaDaOperacao(abas,op){
  if(!op) return null;
  const alvo=norm(op);
  for(const sh of abas) for(const r of (S.sigo.base[sh]||[]))
    if(norm(r.escala)===alvo) return alvo;
  return null;
}
/* Confere período e cobertura de meses. Devolve {ini,fim,dias} ou null (com a
   mensagem já na tela). Roda ANTES do diálogo, para não oferecer a escolha
   quando o período nem é válido. */
function validarPeriodo(){
  const msg=document.getElementById('abs-msgRun'); msg.textContent=''; msg.className='msg';
  const ini=new Date(document.getElementById('abs-dtIni').value+'T00:00:00');
  const fim=new Date(document.getElementById('abs-dtFim').value+'T00:00:00');
  if(!(ini<fim)){ msg.textContent='Período inválido.'; msg.className='msg bad'; return null; }
  const dias=diasPeriodo(ini,fim);
  const mesesNec=[...new Set(dias.map(d=>d.getMonth()+1))];
  const mesesTem=[...new Set(blocosCarregados().flatMap(b=>b.meses))];
  const faltam=mesesNec.filter(m=>!mesesTem.includes(m));
  const mp=document.getElementById('abs-msgPeriodo');
  if(faltam.length){ mp.textContent='⚠ período pede '+faltam.map(m=>MES_NOME[m]).join(', ')+' e as planilhas de abs trazem '+mesesTem.map(m=>MES_NOME[m]).join(', '); mp.className='msg bad'; return null; }
  mp.textContent='✓ período coberto pelas bases'; mp.className='msg ok';
  return {ini,fim,dias};
}
function processar(){
  const v=validarPeriodo(); if(!v) return;
  const {ini,fim,dias}=v;

  const selecionadas=[...document.querySelectorAll('#abs-opsGrid input[type=checkbox]:checked')].map(c=>c.dataset.key);
  const results=[];
  for(const key of selecionadas){
    const [fil,opRaw]=key.split('|');
    const total = opRaw==='__TOTAL__';
    const op = (total||!opRaw)? null : opRaw;
    const src=document.querySelector(`select[data-src="${key}"]`).value;
    const manual=+document.querySelector(`input[data-sop="${key}"]`).value||0;
    /* abas de abs que compõem o cartão: 1 operação, ou todas da filial no cartão
       TOTAL. Deduplicadas por (fil|op|aba): o mesmo arquivo nos dois slots, ou
       uma aba de intervalo que cobre dois meses, não pode entrar duas vezes. */
    const tabsPorMes={}; // mes -> [tabs]
    const vistas=new Set();
    for(const bloco of blocosCarregados()) for(const t of bloco.tabs){
      if(t.fil!==fil) continue;
      if(!total && (t.op??'')!==(op??'')) continue;
      const idt=t.fil+'|'+(t.op??'')+'|'+t.tab;
      if(vistas.has(idt)) continue; vistas.add(idt);
      for(const m of t.meses) (tabsPorMes[m]??=[]).push(t);
    }
    const porMes={}; // mes -> rows concatenadas das operações do cartão
    for(const [m,ts] of Object.entries(tabsPorMes))
      porMes[m]=ts.flatMap(t=>lerAba(t.ws, t.layout).rows);
    const mesA=Math.min(...Object.keys(porMes).map(Number)), mesB=Math.max(...Object.keys(porMes).map(Number));
    const kOf=r=> r.info[0]!=null?String(r.info[0]).trim():'N:'+String(r.info[2]??'').trim();
    const mapA=new Map(porMes[mesA].map(r=>[kOf(r),r]));
    const mapB=mesB!==mesA? new Map(porMes[mesB].map(r=>[kOf(r),r])) : mapA;
    const chaves=[...new Set([...mapA.keys(),...mapB.keys()])];
    const roster=chaves.map(k=>{
      const a=mapA.get(k), b=mapB.get(k);
      const info=(b??a).info.slice();
      const cel={};
      for(const d of dias){
        const m=d.getMonth()+1, dd=d.getDate();
        const sr=m===mesA?a:(m===mesB?b:null);
        if(sr){ cel[+d]=sr.dias[dd]??null; }
        else if(m===mesB&&!b&&a){ cel[+d]='DF'; }
        else { cel[+d]=null; }
      }
      return { info, cel };
    }).sort((x,y)=>String(x.info[2]??'').localeCompare(String(y.info[2]??''),'pt'));

    /* S&OP diário: hc:SEÇÃO, hcsum:SEÇÕES ou manual */
    let secoes=null;
    if(src.startsWith('hc:')) secoes=[src.slice(3)];
    else if(src.startsWith('hcsum:')) secoes=src.slice(6).split(',');
    const sopDia = d=>{
      if(!secoes) return (d.getDay()>=1&&d.getDay()<=5)? manual : null;
      const ym=d.getFullYear()*100+(d.getMonth()+1);
      let soma=0, achou=false;
      for(const s of secoes){
        const v=S.sop[fil]?.[s]?.[ym]?.[d.getDate()];
        if(typeof v==='number'){ soma+=v; achou=true; }
      }
      return achou? soma : undefined; // undefined = sem dado no headcount
    };
    /* Contratado do dia, da linha "Contratado" do bloco embutido — dado de
       origem. null quando as abas não trazem (modelo antigo). */
    const contratadoDia = d=>{
      if(!secoes) return null;
      const ym=d.getFullYear()*100+(d.getMonth()+1);
      let soma=0, achou=false;
      for(const s of secoes){
        const v=S.sopContratado[fil]?.[s]?.[ym]?.[d.getDate()];
        if(typeof v==='number'){ soma+=v; achou=true; }
      }
      return achou? soma : null;
    };
    const semSop=[];
    /* diaristas da compensação: únicos por dia, com GROOT, e SOLICITADOS PELA
       ID LOGISTICS — os do MELI existem no SIGO, mas não cobrem quadro de ABS */
    const sigoSheets=sigoSheetsFor(fil,op,total);
    /* Escala do cartão: o diarista cobre a operação em que foi solicitado. Um
       SD lançado na aba do XD é do SD, não do XD — a referência confirma:
       em 06/08 a aba XD tem 6 XD + 4 FULL e o abate é 6. */
    const escalaOp = total? null : escalaDaOperacao(sigoSheets,op);
    /* Quando a filial não marca ESTA operação na escala mas marca OUTRAS, o
       pool é "tudo menos o que é das outras" — senão o SD de Poços entraria
       no Same Day e também no Service, contado duas vezes. */
    const deOutrasOps = (total||escalaOp)? []
      : Object.keys(OPER).map(norm).filter(o=>o!==norm(op) && escalaDaOperacao(sigoSheets,o));
    const daOperacao = r => {
      const e=norm(r.escala);
      return escalaOp? e===escalaOp : deOutrasOps.indexOf(e)<0;
    };
    /* Pool por dia, separado por solicitante e com ID em PRIORIDADE: quando a
       escolha é "ambos", um GROOT que aparece nos dois no mesmo dia conta como
       ID; o MELI só entra com quem a ID não trouxe. Assim o MELI é sempre
       fallback — usado apenas depois de esgotada a ID. */
    const idPorDia={}, meliPorDia={}, diarRegs=[];
    let foraDaEscala=0;
    for(const sh of sigoSheets) for(const r of S.sigo.base[sh]){
      if(!aceitaSolic(r.solic)) continue;
      if(r.data<ini||r.data>fim) continue;
      if(!daOperacao(r)){ foraDaEscala++; continue; }
      const idSet=(idPorDia[+r.data]??=new Set()), meliSet=(meliPorDia[+r.data]??=new Set());
      if(r.solic==='id'){ if(!idSet.has(r.id)){ idSet.add(r.id); diarRegs.push(r); } }
      else { if(!idSet.has(r.id) && !meliSet.has(r.id)){ meliSet.add(r.id); diarRegs.push(r); } }
    }
    const daily=dias.map(d=>{
      let esp=sopDia(d);
      if(esp===undefined){ semSop.push(d); esp=null; }
      if(esp===0) esp=null; // sem expectativa no dia
      let cont=0,pres=0;
      for(const r of roster){ const v=r.cel[+d]; if(v&&v!=='DF')cont++; if(v==='P')pres++; }
      const semBase = esp!==null && cont===0;
      const dif=(esp!==null&&!semBase)? pres-esp : null;
      const dispId=(idPorDia[+d]??new Set()).size, dispMeli=(meliPorDia[+d]??new Set()).size;
      const disp=dispId+dispMeli;                                 // solicitados no dia (únicos)
      const falta=(dif!==null&&dif<0)? -dif : 0;
      const usadosId=Math.min(dispId, falta);                     // ID primeiro
      const usadosMeli=Math.min(dispMeli, falta-usadosId);        // MELI só no que sobrou
      const usados=usadosId+usadosMeli;                           // = min(disp, déficit)
      const pos=dif!==null? dif+usados : null;                    // teto do abate é zerar o dia
      return { d, esp, cont, contSop:contratadoDia(d), pres, dif, disp, dispId, dispMeli,
               usados, usadosId, usadosMeli, pos, semBase };
    });
    const validos=daily.filter(x=>x.dif!==null);
    const espT=validos.reduce((a,x)=>a+x.esp,0);
    const faltPre=validos.reduce((a,x)=>a+Math.max(0,-x.dif),0);
    const faltPos=validos.reduce((a,x)=>a+Math.max(0,-x.pos),0);
    const usadosT=validos.reduce((a,x)=>a+x.usados,0);
    const usadosIdT=validos.reduce((a,x)=>a+x.usadosId,0);
    const usadosMeliT=validos.reduce((a,x)=>a+x.usadosMeli,0);
    const dispT=validos.filter(x=>x.dif<0).reduce((a,x)=>a+x.disp,0);
    results.push({ key, fil, op, total, src, secoes, manual, roster, daily, dias, mesA, mesB, sigoSheets,
      escalaOp, deOutrasOps, foraDaEscala, diarRegs,
      // filial que não marca a operação na ESCALA: o pool é o mesmo dos irmãos
      escalaIndistinta: !!op && !total && !escalaOp,
      espT, faltPre, faltPos, usadosT, usadosIdT, usadosMeliT, dispT,
      absPre: espT? faltPre/espT : 0, absPos: espT? faltPos/espT : 0,
      semBase: daily.filter(x=>x.semBase).map(x=>x.d), semSop });
  }
  S.results={ ini, fim, list:results };
  render(results,ini,fim);
}
/* Validar abre o diálogo da fonte de diaristas; só depois de escolher é que
   processa. O período/coberturas são validados antes de abrir, para não
   mostrar o diálogo quando nem dá para rodar. */
const modalBg=document.getElementById('abs-modalBg');
function abrirDialogoCompensacao(){
  const marcado=modalBg.querySelector(`input[name="abs-comp"][value="${S.compSrc}"]`);
  if(marcado) marcado.checked=true;
  modalBg.hidden=false;
  modalBg.querySelector('#abs-modalOk').focus();
}
function fecharDialogo(){ modalBg.hidden=true; }
document.getElementById('abs-btnRun').addEventListener('click',()=>{
  if(!validarPeriodo()) return;   // mensagens de período aparecem sem abrir o diálogo
  abrirDialogoCompensacao();
});
document.getElementById('abs-modalCancel').addEventListener('click',fecharDialogo);
modalBg.addEventListener('click',e=>{ if(e.target===modalBg) fecharDialogo(); });
document.addEventListener('keydown',e=>{ if(e.key==='Escape'&&!modalBg.hidden) fecharDialogo(); });
document.getElementById('abs-modalOk').addEventListener('click',()=>{
  S.compSrc=modalBg.querySelector('input[name="abs-comp"]:checked').value;
  fecharDialogo();
  processar();
});

/* ============================== render ============================== */
const fmtPct=v=>(v*100).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+'%';
const fmtDia=d=>String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0');
function render(list,ini,fim){
  document.getElementById('abs-step4').style.display='';
  document.getElementById('abs-resSub').textContent='Período '+fmtDia(ini)+' a '+fmtDia(fim)
    +' · '+list.length+' seleção(ões) · compensação por diaristas '+COMP_LABEL[S.compSrc]
    +' · S&OP diário sem over · abate limitado ao déficit do dia';
  /* A tela por operação foi ocultada a pedido: o resultado completo — por dia,
     por operação — está na planilha. Aqui fica só a confirmação e os avisos
     que mereceriam atenção antes de abrir o arquivo. */
  const area=document.getElementById('abs-resArea'); area.innerHTML='';
  const avisos=[];
  for(const r of list){
    const nome=FILIAIS[r.fil]+(r.total?' · TOTAL':(r.op?' · '+OPER[r.op]:''));
    if(!r.sigoSheets.length) avisos.push(`<b>${nome}</b>: sem aba do SIGO para esta filial — compensação zerada.`);
    if(r.escalaIndistinta&&list.filter(x=>x.fil===r.fil&&x.op&&!x.total).length>1)
      avisos.push(`<b>${nome}</b>: a coluna ESCALA do SIGO não identifica a operação — os diaristas são os mesmos das outras operações da filial e podem estar contados em duas.`);
    if(r.semSop.length) avisos.push(`<b>${nome}</b>: sem S&OP para ${r.semSop.length} dia(s), excluídos do %.`);
  }
  /* Com "ambos", mostra que o MELI foi só fallback: quanto do abate veio da ID
     e quanto o MELI precisou completar. */
  let fallback='';
  if(S.compSrc==='ambos'){
    const id=list.reduce((a,r)=>a+r.usadosIdT,0), meli=list.reduce((a,r)=>a+r.usadosMeliT,0);
    fallback='<div class="abs-pronto" style="border-left-color:var(--amber);margin-top:8px">'
      +'<b style="color:var(--amber)">Ambos — ID primeiro.</b> Do abate, <b>'+id+'</b> vieram da ID Logistics; '
      +'o MELI completou <b>'+meli+'</b>'+(meli?'':' (a ID cobriu tudo)')+'.</div>';
  }
  area.innerHTML='<div class="abs-pronto"><b>✓ Absenteísmo processado.</b> '
    + 'Baixe a planilha — o resultado por dia e por operação está nela.</div>'
    + fallback
    + (avisos.length? '<ul class="abs-avisos">'+avisos.map(a=>'<li>⚠ '+a+'</li>').join('')+'</ul>' : '');

  /* Um download por filial, além do unificado: é o arquivo que vai para cada
     gerente sem levar junto o resto da rede. Os botões nascem do resultado —
     só aparecem as filiais realmente calculadas. */
  const fils=[...new Set(list.map(r=>r.fil))];
  let bar=document.getElementById('abs-porFilial');
  if(!bar){
    bar=document.createElement('div'); bar.id='abs-porFilial'; bar.className='abs-porfilial';
    document.getElementById('abs-btnXlsx').insertAdjacentElement('afterend', bar);
  }
  bar.innerHTML=fils.length>1? '<span class="lb">ou uma filial por arquivo:</span>' : '';
  if(fils.length>1) for(const fil of fils){
    const b=document.createElement('button');
    b.type='button'; b.className='abs-btnfil'; b.textContent=FILIAIS[fil];
    b.addEventListener('click',()=>exportarFilial(fil));
    bar.appendChild(b);
  }

  document.getElementById('abs-step4').scrollIntoView({behavior:'smooth'});
}

/* ============================== export xlsx (modelo unificado + resumo) ============================== */
const COR={hdr:'FF1F4E78', P:'FFC6EFCE', F:'FFFFC7CE', AM:'FFFFEB9C', DF:'FFD9D9D9', FJ:'FFE2EFDA', DSR:'FFDDEBF7', FO:'FFFCE4D6', dif:'FFF2F2F2', med:'FFFFF2CC', diar:'FFFCE4D6'};
const BORD={top:{style:'thin',color:{argb:'FFBFBFBF'}},bottom:{style:'thin',color:{argb:'FFBFBFBF'}},left:{style:'thin',color:{argb:'FFBFBFBF'}},right:{style:'thin',color:{argb:'FFBFBFBF'}}};
const F10={name:'Arial',size:10}, F10B={name:'Arial',size:10,bold:true};
function colL(n){let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=(n-1-m)/26;}return s;}
/* Data para o Excel: SEMPRE meia-noite UTC.
   ----------------------------------------------------------------
   Internamente as datas são meia-noite LOCAL, e o ExcelJS grava o instante
   UTC. Em UTC-3 isso vira 03:00Z, ou seja o serial 46223,125 em vez de
   46223 — e aí o COUNTIFS do resumo, que compara com DATE(2026,7,20)
   (serial inteiro), não casa com nada: TODO abate recalculava para zero ao
   abrir a planilha. Os valores em cache estavam certos, o que enganava.
   Um ambiente em UTC não reproduz o defeito, então a defesa mora aqui, no
   único ponto em que data vira célula. */
const dataExcel = d => new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));

async function exportar(){
  if(!S.results) return;
  await exportarLista(S.results.list);
}
/* Uma filial por arquivo: o mesmo workbook (Diaristas + Unificado + Resumo por
   cartão), restrito aos cartões da filial. É o que se manda para cada gerente
   sem levar junto o resto da rede. */
async function exportarFilial(fil){
  if(!S.results) return;
  const lista=S.results.list.filter(r=>r.fil===fil);
  if(!lista.length) return;
  await exportarLista(lista);
}
/* Nome do arquivo, padronizado:
     "Conciliação ABS Diaristas - Filial - Operação - Período"
   Filial e Operação vêm da SELEÇÃO: quando ela cobre uma só, sai o nome
   dela; quando cobre várias, sai "Todas as Filiais" / "Todas as Operações".
   Um cartão de filial inteira (Divinópolis, sem sub-operação) é "Geral". */
function rotuloOperacao(r){ return r.total? 'TOTAL' : (r.op? OPER[r.op] : 'Geral'); }
function nomeArquivoABS(list){
  const {ini,fim}=S.results;
  const fils=[...new Set(list.map(r=>r.fil))];
  const ops=[...new Set(list.map(rotuloOperacao))];
  const filPart = fils.length===1 ? FILIAIS[fils[0]] : 'Todas as Filiais';
  const opPart  = ops.length===1  ? ops[0]          : 'Todas as Operações';
  const dia=d=>fmtDia(d).replace('/','-')+'-'+d.getFullYear();
  const bruto=`Conciliação ABS Diaristas - ${filPart} - ${opPart} - ${dia(ini)} a ${dia(fim)}`;
  /* Acento no atributo `download` de um link para blob: URL faz o Chromium
     DESCARTAR o nome inteiro e salvar como "download". Não avisa, não erra —
     só entrega o arquivo sem nome. Por isso o acento sai aqui, e não por
     gosto: "Conciliacao" salva certo, "Conciliação" some. Depois tira o que
     é ilegal em nome de arquivo. */
  return bruto.normalize('NFD').replace(/[\u0300-\u036f]/g,'')
              .replace(/[\\/:*?"<>|]/g,'-')+'.xlsx';
}
async function exportarLista(list){
  const {ini,fim}=S.results;
  const wb=new ExcelJS.Workbook();

  /* Uma aba de diaristas POR SELEÇÃO, no layout do SIGO — é o formato do
     modelo de referência, e é o que permite o COUNTIFS do resumo filtrar por
     escala apontando para a aba da própria operação. Só entra quem conta para
     a compensação: solicitado pela ID Logistics e da escala da operação. */
  const umaSo=list.length===1;
  const abaDiar=r=>'Diaristas '+((umaSo||list.every(x=>x.fil===r.fil))
    ? (r.total?'TOTAL':(r.op??FILIAIS[r.fil])) : r.fil+(r.total?'TOTAL':(r.op??'')));
  const HD=['MÊS\nSOLICITAÇÃO','DATA\nSOLICITAÇÃO','SOLICITANTE','EMPRESA\nDIARISTA','GROOT ID','NOME','CARGO','ESCALA'];
  const MESES_PT=['','janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  for(const r of list){
    const wsD=wb.addWorksheet(abaDiar(r).substring(0,31));
    wsD.addRow(HD).eachCell(c=>{c.font={...F10B,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.hdr}};c.border=BORD;c.alignment={horizontal:'center',vertical:'middle',wrapText:true};});
    wsD.getRow(1).height=28;
    for(const reg of r.diarRegs){
      const row=wsD.addRow([MESES_PT[reg.data.getMonth()+1]+'/'+String(reg.data.getFullYear()).slice(2),
        dataExcel(reg.data), reg.solicitante, reg.empresa, isNaN(Number(reg.id))?reg.id:Number(reg.id),
        reg.nome, reg.cargo, reg.escala]);
      row.eachCell(c=>{c.font=F10;c.border=BORD;});
      row.getCell(2).numFmt='dd/mm/yyyy';
    }
    wsD.columns=[{width:15},{width:15},{width:14},{width:14},{width:12},{width:38},{width:14},{width:12}];
    wsD.views=[{state:'frozen',ySplit:1}];
    wsD.autoFilter={from:'A1',to:'H1'};
  }

  for(const r of list){
    const opTag=r.fil+(r.total?'TOTAL':(r.op??''));
    const rotulo=FILIAIS[r.fil]+(r.total?' TOTAL':(r.op?' '+r.op:''));
    const nDias=r.dias.length;
    const ws=wb.addWorksheet(opTag+' Unificado');
    const hdr=['Mat.','Groot ID','Nome','Cargo','Escala','Turno','Liderança','Vínculo',
      ...r.dias.map(d=>d.getDate()+'\n'+DOW[d.getDay()])];
    ws.addRow(hdr).eachCell(c=>{c.font={...F10B,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.hdr}};c.alignment={horizontal:'center',vertical:'middle',wrapText:true};c.border=BORD;});
    ws.getRow(1).height=28;
    for(const p of r.roster){
      const row=ws.addRow([...p.info, ...r.dias.map(d=>p.cel[+d]??null)]);
      row.eachCell({includeEmpty:true},(c,cn)=>{
        c.font=F10;c.border=BORD;
        c.alignment={horizontal:[3,4,5,7].includes(cn)?'left':'center',vertical:'middle'};
        if(cn>8&&c.value!=null){const f=COR[String(c.value).trim()];if(f)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:f}};}
      });
    }
    ws.columns=[{width:9},{width:13},{width:42},{width:26},{width:38},{width:10},{width:30},{width:12},...r.dias.map(()=>({width:6.5}))];
    ws.views=[{state:'frozen',ySplit:1}];
    ws.autoFilter={from:'A1',to:colL(8+nDias)+String(r.roster.length+1)};

    const wr=wb.addWorksheet(umaSo? 'Resumo Gerencial' : ('Resumo '+opTag));
    const lastRow=r.roster.length+1, uni=`'${opTag} Unificado'`;
    const fonteTxt=r.secoes? 'S&OP headcount '+r.secoes.map(s=>s==='TOTAL'?'TOTAL filial':s).join('+')+' sem over' : 'S&OP manual '+r.manual;
    wr.getCell('A1').value='HEADCOUNT DIÁRIO '+rotulo+' ('+fmtDia(ini)+' a '+fmtDia(fim)+'/'+fim.getFullYear()+') · '+fonteTxt;
    wr.getCell('A1').font={name:'Arial',size:13,bold:true,color:{argb:COR.hdr}};
    wr.mergeCells(1,1,1,1+nDias);
    const hRow=wr.getRow(3);
    hRow.getCell(1).value='Métrica';
    r.dias.forEach((d,i)=>hRow.getCell(2+i).value=d.getDate()+'\n'+DOW[d.getDay()]);
    hRow.getCell(2+nDias).value='Média Período';
    for(let c=1;c<=2+nDias;c++){const cl=hRow.getCell(c);cl.font={...F10B,color:{argb:'FFFFFFFF'}};cl.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.hdr}};cl.alignment={horizontal:'center',vertical:'middle',wrapText:true};cl.border=BORD;}
    hRow.height=28;
    /* O Contratado vem da planilha mensal quando ela traz a linha; contar
       marcas não-DF na grade dava 179 num dia em que a origem diz 137. Só o
       fallback (modelo antigo, sem o bloco) continua derivado — e é rotulado
       como derivado, para ninguém confundir. */
    const temContSop=r.daily.some(x=>x.contSop!==null&&x.contSop!==undefined);
    /* Seis métricas, como o modelo de referência: a linha de diaristas é o
       ABATE (o quanto do déficit foi coberto), não o estoque do dia. */
    const rotulos=['Quadro S&OP',
      temContSop?'Contratado - Escala do dia (planilha mensal)':'Contratado - Escala do dia (derivado)',
      'Presente','Diferença (S&OP-Presente)',
      'Diaristas ID Logistics (abate do déficit)','Abs Pós Compensação Diaristas'];
    rotulos.forEach((t,i)=>{const c=wr.getCell(4+i,1);c.value=t;c.font=F10B;c.border=BORD;
      if(i===3)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.dif}};
      if(i===4)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.diar}};});
    r.dias.forEach((d,i)=>{
      const col=2+i, L=colL(col), pL=colL(9+i), day=r.daily[i];
      const set=(rw,formula,res,fill)=>{const c=wr.getCell(rw,col);
        c.value=formula?{formula,result:res}:(res===null||res===undefined)?null:res;
        c.font=rw===7||rw===10?F10B:F10;c.alignment={horizontal:'center'};c.border=BORD;
        if(fill)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:fill}};};
      set(4,null,day.esp);
      // valor da origem entra como valor; a fórmula viva só faz sentido no derivado
      if(temContSop) set(5,null,day.contSop);
      else set(5,`SUMPRODUCT((${uni}!${pL}2:${pL}${lastRow}<>"")*(${uni}!${pL}2:${pL}${lastRow}<>"DF"))`,day.cont);
      set(6,`COUNTIF(${uni}!${pL}2:${pL}${lastRow},"P")`,day.pres);
      set(7,day.dif!==null?`${L}6-${L}4`:null,day.dif,COR.dif);
      /* Abate: conta os diaristas do dia na aba da própria operação — filtrando
         a ESCALA quando a filial a usa — e limita ao déficit, que é a regra da
         aba: um dia nunca fica positivo por sobra de diarista. */
      const alvo=`'${abaDiar(r)}'`;
      const criterios=`${alvo}!$B:$B,DATE(${d.getFullYear()},${d.getMonth()+1},${d.getDate()})`
        + (r.escalaOp? `,${alvo}!$H:$H,"${r.escalaOp.toUpperCase()}"` : '');
      set(8,day.dif!==null?`IF(${L}7>=0,0,MIN(COUNTIFS(${criterios}),-${L}7))`:null,
          day.dif!==null?day.usados:null,COR.diar);
      set(9,day.dif!==null?`${L}7+${L}8`:null,day.pos);
    });
    const mc=2+nDias, mL=colL(1+nDias);
    const med=(rw,f,res)=>{const c=wr.getCell(rw,mc);c.value={formula:f,result:res};c.font=F10B;c.alignment={horizontal:'center'};c.border=BORD;c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.med}};c.numFmt='0.0';};
    const vs=r.daily.filter(x=>x.dif!==null);
    med(4,`ROUND(AVERAGEIF(B4:${mL}4,"<>0"),1)`, vs.length?+(vs.reduce((a,x)=>a+x.esp,0)/vs.length).toFixed(1):0);
    med(5,`ROUND(AVERAGEIF(B5:${mL}5,"<>0"),1)`, vs.length?+(vs.reduce((a,x)=>a+(temContSop?(x.contSop??0):x.cont),0)/vs.length).toFixed(1):0);
    med(6,`ROUND(AVERAGEIF(B6:${mL}6,"<>0"),1)`, vs.length?+(vs.reduce((a,x)=>a+x.pres,0)/vs.length).toFixed(1):0);
    med(7,`ROUND(SUMIF(B7:${mL}7,"<0"),1)`, -r.faltPre);
    med(8,`SUM(B8:${mL}8)`, r.usadosT);
    med(9,`ROUND(SUMIF(B9:${mL}9,"<0"),1)`, -r.faltPos);
    const bl=[[11,'Abs Operacional ANTES (faltas/S&OP)',r.absPre],[12,'Abs Operacional PÓS diaristas',r.absPos],[13,'Range contratual',0.025]];
    for(const [rw,t,v] of bl){
      wr.getCell(rw,1).value=t; wr.getCell(rw,1).font=F10B;
      const c=wr.getCell(rw,2); c.value=v; c.numFmt='0.00%'; c.font=F10B; c.alignment={horizontal:'center'};
      c.fill={type:'pattern',pattern:'solid',fgColor:{argb: rw===13?COR.med : (v<=0.025?COR.P:COR.F)}};
    }
    wr.getColumn(1).width=34;
    for(let c=2;c<=1+nDias;c++)wr.getColumn(c).width=8;
    wr.getColumn(mc).width=14;
    wr.views=[{state:'frozen',xSplit:1,ySplit:3}];
  }

  const buf=await wb.xlsx.writeBuffer();
  const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=nomeArquivoABS(list);
  a.click(); URL.revokeObjectURL(a.href);
}
document.getElementById('abs-btnXlsx').addEventListener('click',exportar);
})();
