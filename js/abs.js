/* Calcular ABS — validador de absenteísmo por filial
   ================================================================

   Confronta o quadro S&OP diário (sem over) com quem esteve presente,
   e compensa o déficit de cada dia com os diaristas efetivamente
   solicitados naquele dia — o abate é limitado ao próprio déficit, de
   modo que um dia nunca fica "positivo" por excesso de diarista.

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
const S = { mesA:null, mesB:null, sigo:null, sop:{}, sopFiles:[], results:null };
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
/* abas de operação das planilhas mensais: "PAXD Jul", "DV Ago" */
function parseTabName(name){
  const m = /^([A-Za-z]+)\s+([A-Za-zç]{3})/.exec(name.trim());
  if(!m) return null;
  const mes = MESES[norm(m[2]).slice(0,3)];
  if(!mes) return null;
  let code = m[1].toUpperCase(), op=null;
  for(const suf of ['SVC','FULL','XD','SD']){
    if(code.endsWith(suf) && code.length>suf.length){ op=suf; code=code.slice(0,-suf.length); break; }
  }
  if(!FILIAIS[code]) return null;
  return { fil:code, op, mes, tab:name };
}
function validaAba(ws){
  const h1=String(cellVal(ws,1,1)??''), h3=String(cellVal(ws,1,3)??''), h9=String(cellVal(ws,1,9)??'');
  if(!/mat/i.test(h1) || !/nome/i.test(h3)) return 'cabeçalho fora do padrão (esperado Mat./Nome nas colunas A/C)';
  if(!/^\d{1,2}\s*\n?/.test(h9)) return 'colunas de dias não encontradas a partir da coluna I';
  return null;
}
function lerAba(ws){
  const dias={};
  for(let c=9;c<=ws.columnCount;c++){
    const h=String(cellVal(ws,1,c)??''); const m=/^(\d{1,2})/.exec(h.trim());
    if(m) dias[+m[1]]=c;
  }
  const rows=[];
  ws.eachRow((row,rn)=>{
    if(rn===1) return;
    const mat=row.getCell(1).value, nome=row.getCell(3).value;
    if(mat==null && nome==null) return;
    const rec={ info:[], dias:{} };
    for(let c=1;c<=8;c++){ let v=row.getCell(c).value; if(v&&typeof v==='object'&&v.richText) v=v.richText.map(t=>t.text).join(''); rec.info.push(v??null); }
    for(const [d,c] of Object.entries(dias)){ let v=row.getCell(+c).value; if(v!=null&&v!=='') rec.dias[+d]=String(v).trim(); }
    rows.push(rec);
  });
  return { rows };
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
    const tabs=[]; const warns=[];
    for(const ws of wb.worksheets){
      const p=parseTabName(ws.name); if(!p) continue;
      const err=validaAba(ws);
      if(err){ warns.push(ws.name+': '+err); continue; }
      tabs.push({...p, ws});
    }
    if(!tabs.length) throw new Error('padrão diferente — nenhuma aba de operação reconhecida (ex.: PAXD Jul, DV Ago)');
    const meses=[...new Set(tabs.map(t=>t.mes))];
    S[slot]={ file:file.name, tabs, meses };
    st.textContent='✓ '+tabs.length+' operações · mês: '+meses.map(m=>MES_NOME[m]).join(', ')
      + (warns.length? ' · ⚠ '+warns.length+' aba(s) fora do padrão':'');
    st.className='st '+(warns.length?'warn':'ok'); drop.classList.add('loaded');
    if(warns.length) console.warn('Abas ignoradas:',warns);
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
        // ExcelJS entrega datas em UTC: usar getters UTC para não deslocar 1 dia em UTC-3
        regs.push({ data:new Date(dt.getUTCFullYear(),dt.getUTCMonth(),dt.getUTCDate()), id:idn,
                    nome:String(row.getCell(6).value??''), escala:String(row.getCell(8).value??'') });
      });
      base[ws.name]=regs;
    }
    S.sigo={ file:file.name, base };
    st.textContent='✓ '+abas.length+' filiais · '+Object.values(base).reduce((a,b)=>a+b.length,0)+' registros com ID';
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
      for(const b of r.blocos){
        const dest=(((S.sop[b.filCode]??={})[b.secao]??={})[b.ym]??={});
        Object.assign(dest, b.dias); // funde dias (abas por mês contribuem para o mesmo ym)
      }
      if(r.avisos&&r.avisos.length) console.warn('Headcount avisos:',r.avisos);
      if(!S.sopFiles.includes(file.name)) S.sopFiles.push(file.name);
    }
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
/* monta a lista de cartões: cada operação + um cartão TOTAL por filial com 2+ operações */
function listarCartoes(){
  const keysA=new Set(S.mesA.tabs.map(t=>t.fil+'|'+(t.op??'')));
  const ops=S.mesB.tabs.filter(t=>keysA.has(t.fil+'|'+(t.op??'')));
  const cards=ops.map(t=>({ fil:t.fil, op:t.op??null, total:false }));
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
  if(card.total || !card.op){
    const t=secs.find(s=>norm(s)==='total');
    if(t) sug.push({v:'hc:'+t, l:'Headcount · TOTAL filial (sem over)'});
    if(!t && semTotal.length) sug.push({v:'hcsum:'+semTotal.join(','), l:'Headcount · soma das seções'});
    for(const s of semTotal) outras.push({v:'hc:'+s, l:'Headcount · '+s});
  } else {
    const m=semTotal.find(s=>norm(s)===norm(card.op));
    if(m) sug.push({v:'hc:'+m, l:'Headcount · '+m+' (sem over)'});
    const am=semTotal.find(s=>norm(s)==='am'), pm=semTotal.find(s=>norm(s)==='pm');
    if(!m && am && pm && norm(card.op)==='svc')
      sug.push({v:'hcsum:'+am+','+pm, l:'Headcount · AM+PM (soma)'});
    for(const s of semTotal) if(s!==m) outras.push({v:'hc:'+s, l:'Headcount · '+s});
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
  if(!S.mesA||!S.mesB){ grid.innerHTML='<span class="msg" style="color:var(--dim)">Carregue as duas planilhas mensais para listar as operações.</span>'; return; }
  const cards=listarCartoes();
  if(!cards.length){ grid.innerHTML='<span class="msg bad">Nenhuma operação em comum entre as duas planilhas.</span>'; return; }
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
  document.getElementById('abs-btnRun').disabled=!(S.mesA&&S.mesB&&S.sigo&&any);
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
  if(op){ const comOp=doFil.find(n=>norm(n).includes(norm(op))); if(comOp) return [comOp]; }
  const semOp=doFil.find(n=>!['svc','xd','sd','full'].some(s=>norm(n).endsWith(' '+s)));
  return [semOp ?? doFil[0]];
}
function processar(){
  const msg=document.getElementById('abs-msgRun'); msg.textContent=''; msg.className='msg';
  const ini=new Date(document.getElementById('abs-dtIni').value+'T00:00:00');
  const fim=new Date(document.getElementById('abs-dtFim').value+'T00:00:00');
  if(!(ini<fim)){ msg.textContent='Período inválido.'; msg.className='msg bad'; return; }
  const dias=diasPeriodo(ini,fim);
  const mesesNec=[...new Set(dias.map(d=>d.getMonth()+1))];
  const mesesTem=[...new Set([...S.mesA.meses,...S.mesB.meses])];
  const faltam=mesesNec.filter(m=>!mesesTem.includes(m));
  const mp=document.getElementById('abs-msgPeriodo');
  if(faltam.length){ mp.textContent='⚠ período pede '+faltam.map(m=>MES_NOME[m]).join(', ')+' e as planilhas de abs trazem '+mesesTem.map(m=>MES_NOME[m]).join(', '); mp.className='msg bad'; return; }
  mp.textContent='✓ período coberto pelas bases'; mp.className='msg ok';

  const selecionadas=[...document.querySelectorAll('#abs-opsGrid input[type=checkbox]:checked')].map(c=>c.dataset.key);
  const results=[];
  for(const key of selecionadas){
    const [fil,opRaw]=key.split('|');
    const total = opRaw==='__TOTAL__';
    const op = (total||!opRaw)? null : opRaw;
    const src=document.querySelector(`select[data-src="${key}"]`).value;
    const manual=+document.querySelector(`input[data-sop="${key}"]`).value||0;
    /* abas de abs que compõem o cartão: 1 operação, ou todas da filial no cartão TOTAL */
    const tabsPorMes={}; // mes -> [tabs]
    for(const bloco of [S.mesA,S.mesB]) for(const t of bloco.tabs){
      if(t.fil!==fil) continue;
      if(!total && (t.op??'')!==(op??'')) continue;
      (tabsPorMes[t.mes]??=[]).push(t);
    }
    const porMes={}; // mes -> rows concatenadas das operações do cartão
    for(const [m,ts] of Object.entries(tabsPorMes))
      porMes[m]=ts.flatMap(t=>lerAba(t.ws).rows);
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
    const semSop=[];
    /* diaristas (somente ID, únicos por dia, união das abas do SIGO do cartão) */
    const sigoSheets=sigoSheetsFor(fil,op,total);
    const diarPorDia={};
    for(const sh of sigoSheets) for(const r of S.sigo.base[sh]){
      if(r.data>=ini&&r.data<=fim){ (diarPorDia[+r.data]??=new Set()).add(r.id); }
    }
    const daily=dias.map(d=>{
      let esp=sopDia(d);
      if(esp===undefined){ semSop.push(d); esp=null; }
      if(esp===0) esp=null; // sem expectativa no dia
      let cont=0,pres=0;
      for(const r of roster){ const v=r.cel[+d]; if(v&&v!=='DF')cont++; if(v==='P')pres++; }
      const semBase = esp!==null && cont===0;
      const dif=(esp!==null&&!semBase)? pres-esp : null;
      const disp=(diarPorDia[+d]??new Set()).size;               // solicitados no dia (com ID)
      const usados=(dif!==null&&dif<0)? Math.min(disp,-dif) : 0; // abate limitado ao déficit
      const pos=dif!==null? dif+usados : null;                   // teto do abate é zerar o dia
      return { d, esp, cont, pres, dif, disp, usados, pos, semBase };
    });
    const validos=daily.filter(x=>x.dif!==null);
    const espT=validos.reduce((a,x)=>a+x.esp,0);
    const faltPre=validos.reduce((a,x)=>a+Math.max(0,-x.dif),0);
    const faltPos=validos.reduce((a,x)=>a+Math.max(0,-x.pos),0);
    const usadosT=validos.reduce((a,x)=>a+x.usados,0);
    const dispT=validos.filter(x=>x.dif<0).reduce((a,x)=>a+x.disp,0);
    results.push({ key, fil, op, total, src, secoes, manual, roster, daily, dias, mesA, mesB, sigoSheets,
      espT, faltPre, faltPos, usadosT, dispT,
      absPre: espT? faltPre/espT : 0, absPos: espT? faltPos/espT : 0,
      semBase: daily.filter(x=>x.semBase).map(x=>x.d), semSop });
  }
  S.results={ ini, fim, list:results };
  render(results,ini,fim);
}
document.getElementById('abs-btnRun').addEventListener('click',processar);

/* ============================== render ============================== */
const fmtPct=v=>(v*100).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+'%';
const fmtDia=d=>String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0');
function render(list,ini,fim){
  document.getElementById('abs-step4').style.display='';
  document.getElementById('abs-resSub').textContent='Período '+fmtDia(ini)+' a '+fmtDia(fim)+' · '+list.length+' seleção(ões) · S&OP diário sem over · abate limitado ao déficit do dia';
  const area=document.getElementById('abs-resArea'); area.innerHTML='';
  for(const r of list){
    const nome=FILIAIS[r.fil]+(r.total?' · TOTAL':(r.op?' · '+OPER[r.op]:''));
    const fonte=r.secoes? 'headcount · '+r.secoes.map(s=>s==='TOTAL'?'TOTAL filial':s).join('+')+' (sem over)' : 'manual '+r.manual+' (seg–sex)';
    const okPos=r.absPos<=0.025, okPre=r.absPre<=0.025;
    const pin=v=>Math.min(96,Math.max(3,(v/0.06)*100));
    const defRows=r.daily.filter(x=>x.dif!==null&&x.dif<0);
    const el=document.createElement('div'); el.className='opres';
    el.innerHTML=`
      <h3>${nome}
        <span class="badge ${okPre?'ok':'bad'}">antes ${fmtPct(r.absPre)}</span>
        <span class="badge ${okPos?'ok':'bad'}">pós diaristas ${fmtPct(r.absPos)}</span>
      </h3>
      <div class="gauge"><div class="lbl"><span>0%</span><span>6%</span></div>
        <div class="gtrack" style="--cut:${(0.025/0.06)*100}%">
          <div class="gcut"></div>
          <span class="gpin antes" style="left:${pin(r.absPre)}%">A ${fmtPct(r.absPre)}</span>
          <span class="gpin pos ${okPos?'':'above'}" style="left:${pin(r.absPos)}%">P ${fmtPct(r.absPos)}</span>
        </div></div>
      <div class="kpis">
        <div class="kpi"><div class="l">S&amp;OP período</div><div class="v">${r.espT}</div></div>
        <div class="kpi"><div class="l">Faltas vs S&amp;OP</div><div class="v r">${r.faltPre}</div></div>
        <div class="kpi"><div class="l">Diaristas usados</div><div class="v y">${r.usadosT}</div></div>
        <div class="kpi"><div class="l">Disponíveis (dias déficit)</div><div class="v">${r.dispT}</div></div>
        <div class="kpi"><div class="l">Descoberto pós</div><div class="v ${r.faltPos?'r':'g'}">${r.faltPos}</div></div>
        <div class="kpi"><div class="l">Colaboradores</div><div class="v">${r.roster.length}</div></div>
      </div>
      ${defRows.length?`<table class="mini"><thead><tr><th>Dia crítico</th><th>S&amp;OP</th><th>Presente</th><th>Dif</th><th>Disponíveis</th><th>Usados</th><th>Pós</th></tr></thead>
        <tbody>${defRows.map(x=>`<tr><td>${fmtDia(x.d)} (${DOW[x.d.getDay()]})</td><td>${x.esp}</td><td>${x.pres}</td>
          <td class="neg">${x.dif}</td><td class="zer">${x.disp||'—'}</td><td class="${x.usados?'pos':'zer'}">${x.usados||'—'}</td><td class="${x.pos<0?'neg':'zer'}">${x.pos}</td></tr>`).join('')}</tbody></table>`
        :'<div class="note">Nenhum dia abaixo do S&OP no período. Diaristas zerados por regra.</div>'}
      <div class="note">Fonte S&OP: <b style="color:var(--txt)">${fonte}</b>${r.sigoSheets.length?` · Diaristas: aba(s) <b style="color:var(--txt)">${r.sigoSheets.join(' + ')}</b> do SIGO.`:' · <b>⚠</b> sem aba do SIGO para esta filial — compensação zerada.'}</div>
      ${r.semSop.length?`<div class="note"><b>⚠</b> Sem S&OP no headcount para ${r.semSop.length} dia(s) (excluídos do %): ${r.semSop.slice(0,8).map(fmtDia).join(', ')}${r.semSop.length>8?'…':''} — carregue o headcount do mês correspondente.</div>`:''}
      ${r.semBase.length?`<div class="note"><b>⚠</b> Dias sem lançamento na base de abs (excluídos do %): ${r.semBase.map(fmtDia).join(', ')}.</div>`:''}
    `;
    area.appendChild(el);
  }
  document.getElementById('abs-step4').scrollIntoView({behavior:'smooth'});
}

/* ============================== export xlsx (modelo unificado + resumo) ============================== */
const COR={hdr:'FF1F4E78', P:'FFC6EFCE', F:'FFFFC7CE', AM:'FFFFEB9C', DF:'FFD9D9D9', FJ:'FFE2EFDA', DSR:'FFDDEBF7', FO:'FFFCE4D6', dif:'FFF2F2F2', med:'FFFFF2CC', diar:'FFFCE4D6'};
const BORD={top:{style:'thin',color:{argb:'FFBFBFBF'}},bottom:{style:'thin',color:{argb:'FFBFBFBF'}},left:{style:'thin',color:{argb:'FFBFBFBF'}},right:{style:'thin',color:{argb:'FFBFBFBF'}}};
const F10={name:'Arial',size:10}, F10B={name:'Arial',size:10,bold:true};
function colL(n){let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=(n-1-m)/26;}return s;}

async function exportar(){
  if(!S.results) return;
  const {ini,fim,list}=S.results;
  const wb=new ExcelJS.Workbook();

  const wsD=wb.addWorksheet('Diaristas');
  wsD.addRow(['SELEÇÃO','DATA','GROOT ID','NOME','ESCALA','ABA SIGO']).eachCell(c=>{c.font={...F10B,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.hdr}};c.border=BORD;c.alignment={horizontal:'center'};});
  for(const r of list){
    const rotulo=FILIAIS[r.fil]+(r.total?' TOTAL':(r.op?' '+r.op:''));
    const vistos=new Set();
    for(const sh of r.sigoSheets) for(const reg of S.sigo.base[sh]){
      if(reg.data<ini||reg.data>fim) continue;
      const k=+reg.data+'|'+reg.id; if(vistos.has(k))continue; vistos.add(k);
      const row=wsD.addRow([rotulo,reg.data,reg.id,reg.nome,reg.escala,sh]);
      row.eachCell(c=>{c.font=F10;c.border=BORD;});
      row.getCell(2).numFmt='dd/mm/yyyy';
    }
  }
  wsD.columns=[{width:22},{width:13},{width:12},{width:38},{width:14},{width:20}];
  wsD.views=[{state:'frozen',ySplit:1}];

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

    const wr=wb.addWorksheet('Resumo '+opTag);
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
    const rotulos=['Quadro S&OP (sem over)','Contratado - Escala do dia','Presente','Diferença (S&OP-Presente)','Diaristas Disponíveis (com ID)','Diaristas Utilizados (abate)','Saldo Pós Compensação'];
    rotulos.forEach((t,i)=>{const c=wr.getCell(4+i,1);c.value=t;c.font=F10B;c.border=BORD;
      if(i===3)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.dif}};
      if(i===4||i===5)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.diar}};});
    r.dias.forEach((d,i)=>{
      const col=2+i, L=colL(col), pL=colL(9+i), day=r.daily[i];
      const set=(rw,formula,res,fill)=>{const c=wr.getCell(rw,col);
        c.value=formula?{formula,result:res}:(res===null||res===undefined)?null:res;
        c.font=rw===7||rw===10?F10B:F10;c.alignment={horizontal:'center'};c.border=BORD;
        if(fill)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:fill}};};
      set(4,null,day.esp);
      set(5,`SUMPRODUCT((${uni}!${pL}2:${pL}${lastRow}<>"")*(${uni}!${pL}2:${pL}${lastRow}<>"DF"))`,day.cont);
      set(6,`COUNTIF(${uni}!${pL}2:${pL}${lastRow},"P")`,day.pres);
      set(7,day.dif!==null?`${L}6-${L}4`:null,day.dif,COR.dif);
      set(8,day.dif!==null?`COUNTIFS(Diaristas!$B:$B,DATE(${d.getFullYear()},${d.getMonth()+1},${d.getDate()}),Diaristas!$A:$A,"${rotulo}")`:null,day.dif!==null?day.disp:null,COR.diar);
      set(9,day.dif!==null?`IF(${L}7>=0,0,MIN(${L}8,-${L}7))`:null,day.dif!==null?day.usados:null,COR.diar);
      set(10,day.dif!==null?`${L}7+${L}9`:null,day.pos);
    });
    const mc=2+nDias, mL=colL(1+nDias);
    const med=(rw,f,res)=>{const c=wr.getCell(rw,mc);c.value={formula:f,result:res};c.font=F10B;c.alignment={horizontal:'center'};c.border=BORD;c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COR.med}};c.numFmt='0.0';};
    const vs=r.daily.filter(x=>x.dif!==null);
    med(4,`ROUND(AVERAGEIF(B4:${mL}4,"<>0"),1)`, vs.length?+(vs.reduce((a,x)=>a+x.esp,0)/vs.length).toFixed(1):0);
    med(5,`ROUND(AVERAGEIF(B5:${mL}5,"<>0"),1)`, vs.length?+(vs.reduce((a,x)=>a+x.cont,0)/vs.length).toFixed(1):0);
    med(6,`ROUND(AVERAGEIF(B6:${mL}6,"<>0"),1)`, vs.length?+(vs.reduce((a,x)=>a+x.pres,0)/vs.length).toFixed(1):0);
    med(7,`ROUND(SUMIF(B7:${mL}7,"<0"),1)`, -r.faltPre);
    med(8,`SUM(B8:${mL}8)`, vs.reduce((a,x)=>a+(x.dif!==null?x.disp:0),0));
    med(9,`SUM(B9:${mL}9)`, r.usadosT);
    med(10,`ROUND(SUMIF(B10:${mL}10,"<0"),1)`, -r.faltPos);
    const bl=[[12,'Abs Operacional ANTES (faltas/S&OP)',r.absPre],[13,'Abs Operacional PÓS diaristas',r.absPos],[14,'Range contratual',0.025]];
    for(const [rw,t,v] of bl){
      wr.getCell(rw,1).value=t; wr.getCell(rw,1).font=F10B;
      const c=wr.getCell(rw,2); c.value=v; c.numFmt='0.00%'; c.font=F10B; c.alignment={horizontal:'center'};
      c.fill={type:'pattern',pattern:'solid',fgColor:{argb: rw===14?COR.med : (v<=0.025?COR.P:COR.F)}};
    }
    wr.getColumn(1).width=34;
    for(let c=2;c<=1+nDias;c++)wr.getColumn(c).width=8;
    wr.getColumn(mc).width=14;
    wr.views=[{state:'frozen',xSplit:1,ySplit:3}];
  }

  const buf=await wb.xlsx.writeBuffer();
  const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='Absenteismo_Unificado_'+fmtDia(ini).replace('/','-')+'_a_'+fmtDia(fim).replace('/','-')+'.xlsx';
  a.click(); URL.revokeObjectURL(a.href);
}
document.getElementById('abs-btnXlsx').addEventListener('click',exportar);
})();
