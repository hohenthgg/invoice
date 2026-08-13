/* Fusão de Linhas — auditoria Labor × Retorno, plano de equalização e exportações.
   Todo o módulo vive dentro de uma IIFE: a aba de Conciliação Faturas roda no
   escopo global e as duas definem, entre outras, uma função `render`. O que o
   HTML precisa chamar é publicado de forma explícita em `window.Fusao`. */
"use strict";
(function(){
const drop=document.getElementById('fz-drop'),fi=document.getElementById('fz-file');
drop.onclick=()=>fi.click();
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('over')}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('over')}));
drop.addEventListener('drop',ev=>{if(ev.dataTransfer.files[0])handle(ev.dataTransfer.files[0])});
fi.onchange=()=>{if(fi.files[0])handle(fi.files[0])};

function showErr(m){const b=document.getElementById('fz-errbox');b.textContent=m;b.classList.remove('hidden')}
function norm(s){return String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim()}
// Cargos que o MELI considera no PREF (em forma normalizada). Editável.
const CARGOS_MELI=["auxiliar de apoio log i","operador transpaleteira","part time - auxiliar log i (3 dias)"];
function exDate(v){ // serial Excel ou Date → Date UTC-neutra
  if(v==null||v==='')return null;
  if(v instanceof Date)return new Date(Date.UTC(v.getFullYear(),v.getMonth(),v.getDate()));
  if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);return d?new Date(Date.UTC(d.y,d.m-1,d.d)):null}
  const m=String(v).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if(m){let y=+m[3];if(y<100)y+=2000;return new Date(Date.UTC(y,+m[2]-1,+m[1]))}
  const d=new Date(v);return isNaN(d)?null:new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
}
const fmt=d=>d?String(d.getUTCDate()).padStart(2,'0')+'/'+String(d.getUTCMonth()+1).padStart(2,'0'):'—';

function findHeader(rows,keys){ // retorna {idx, map col→indice}
  for(let i=0;i<Math.min(rows.length,8);i++){
    const cells=(rows[i]||[]).map(norm);
    if(keys.every(k=>cells.some(c=>c.includes(k)))){
      return {idx:i,cells};
    }
  }
  return null;
}
const col=(cells,...alts)=>{for(const a of alts){const j=cells.findIndex(c=>c.includes(a));if(j>=0)return j}return -1};

function handle(file){
  document.getElementById('fz-errbox').classList.add('hidden');
  const r=new FileReader();
  r.onload=e=>{
    try{audit(XLSX.read(e.target.result,{type:'array',cellDates:true}))}
    catch(err){showErr('Não consegui ler o arquivo: '+err.message)}
  };
  r.readAsArrayBuffer(file);
}

function audit(wb){
  let labRows=null,retRows=null,labHdr=null,retHdr=null;
  for(const name of wb.SheetNames){
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:''});
    if(!labHdr){const h=findHeader(rows,['groot','inicio']); if(h){labHdr=h;labRows=rows}}
    if(!retHdr){const h=findHeader(rows,['data trab','qtd. pref']); if(h){retHdr=h;retRows=rows}}
  }
  if(!labHdr)return showErr('Não encontrei a aba de Labor (procuro cabeçalho com "GROOT ID" e "DATA DE INÍCIO").');
  if(!retHdr)return showErr('Não encontrei a aba de Retorno MELI (procuro cabeçalho com "Data Trab." e "Qtd. PREF").');

  // ---- Labor ----
  const lc=labHdr.cells;
  const labHeaderRow=labRows[labHdr.idx];   // cabeçalho original (exato)
  const iG=col(lc,'groot'),iN=col(lc,'nome'),iM=col(lc,'matricula'),iC=col(lc,'cargo'),
        iI=col(lc,'inicio'),iF=col(lc,'data fim','fim'),iR=col(lc,'rateio'),iRg=col(lc,'regime');
  const labor=[];
  for(let i=labHdr.idx+1;i<labRows.length;i++){
    const row=labRows[i]; if(!row||row.every(c=>c===''))continue;
    const g=String(row[iG]??'').trim();
    if(norm(g)==='abs'||norm(String(row[iN]??'')).includes('absenteismo'))continue;
    const ini=exDate(row[iI]); if(!ini)continue;
    labor.push({groot:g,nome:String(row[iN]??''),mat:String(row[iM]??''),cargo:String(row[iC]??''),regime:iRg>=0?String(row[iRg]??''):'',
      ini,fim:exDate(row[iF]),rateio:(iR>=0&&row[iR]!==''&&!isNaN(+row[iR]))?+row[iR]:1,
      pt:norm(row[iC]).includes('part time'),elig:CARGOS_MELI.includes(norm(row[iC])),
      rawRow:row.slice()});
  }
  if(!labor.length)return showErr('A aba de Labor foi encontrada mas não tem linhas válidas (verifique a DATA DE INÍCIO).');
  // contexto para o exportador
  window.__labCtx={header:labHeaderRow,iI,iF};

  // ---- Retorno ----
  const rc=retHdr.cells;
  const jD=col(rc,'data trab'),jT=col(rc,'employee'),jP=col(rc,'qtd. pref'),
        jQ=col(rc,'pos comp','pós comp'),jV=col(rc,'desvio'),jO=col(rc,'ocorrencia');
  const ret=[];
  for(let i=retHdr.idx+1;i<retRows.length;i++){
    const row=retRows[i]; const d=exDate(row?.[jD]); if(!d)continue;
    ret.push({d,tipo:norm(row[jT]||'full_time'),pref:+row[jP]||0,
      qpos:jQ>=0?(+row[jQ]||0):null,desv:jV>=0&&row[jV]!==''?+row[jV]:null,
      ocor:jO>=0?String(row[jO]??''):''});
  }
  if(!ret.length)return showErr('A aba de Retorno foi encontrada mas não tem linhas com data válida.');

  // ---- Conciliação ----
  // Dif. = FT no seu Labor - Q Pós Comp.
  const count=(d,pt)=>labor.reduce((s,p)=>s+((p.elig&&p.pt===pt)&&p.ini<=d&&(p.fim==null||p.fim>=d)?p.rateio:0),0);
  const dias=ret.map(r=>{
    const isPT=r.tipo.includes('part');
    const meu=count(r.d,isPT);
    // Diferença principal: FT no Labor menos Q Pós Compensação
    const dif=r.qpos!=null?meu-r.qpos:meu-r.pref;
    const desvRec=r.qpos!=null?r.pref-r.qpos:null;
    let st,diag;
    if(Math.round(dif*100)!==0){
      st='meu';
      diag=dif>0?`Seu Labor tem ${fmtN(Math.abs(dif))} pessoa(s) a MAIS do que a Q Pós Compensação — verifique se há pessoas extras ou ausências não lançadas.`
               :`Seu Labor tem ${fmtN(Math.abs(dif))} pessoa(s) a MENOS do que a Q Pós Compensação — linhas faltando ou DATA FIM cortada cedo demais.`;
    } else if(r.desv!=null&&r.desv>3){
      st='meli'; diag='Labor bate com a Q Pós Compensação. O desvio vem do lado MELI: o Scheduling/Rostering não reconheceu a escala desse dia — acionar o BC/planning do MELI.';
    } else if(r.desv!=null&&r.desv<0){
      st='meli'; diag='Labor bate com a Q Pós Compensação, mas operou mais gente do que o declarado — avalie nova versão do PREF para refletir o quadro real.';
    } else {st='ok';diag='OK — contagem e desvio sob controle.'}
    return {...r,meu,outroPT:count(r.d,!isPT),dif,desvRec,st,diag};
  });

  // ---- Verificações do Labor ----
  const grootZero=labor.filter(p=>p.groot==='0'||p.groot===''),
        fimAntes=labor.filter(p=>p.fim&&p.fim<p.ini);
  const byG={}; labor.forEach(p=>{if(p.groot&&p.groot!=='0'){(byG[p.groot]=byG[p.groot]||[]).push(p)}});
  const overlaps=[];
  Object.values(byG).forEach(g=>{ if(g.length<2)return;
    for(let a=0;a<g.length;a++)for(let b=a+1;b<g.length;b++){
      const fa=g[a].fim??new Date(Date.UTC(2099,0,1)),fb=g[b].fim??new Date(Date.UTC(2099,0,1));
      if(g[a].ini<=fb&&g[b].ini<=fa)overlaps.push(g[a].nome+' ('+g[a].groot+')');
    }});
  const fimPeriodo=ret.reduce((m,r)=>r.d>m?r.d:m,ret[0].d);
  const truncados=[];
  labor.forEach(p=>{ if(!p.fim||p.fim>=fimPeriodo)return;
    const temSucessor=labor.some(q=>q!==p&&(q.groot===p.groot&&p.groot!=='0'||q.mat&&q.mat===p.mat)&&q.ini>p.fim);
    if(!temSucessor)truncados.push({...p}); });
  const desvDiv=dias.filter(d=>d.desv!=null&&d.desvRec!=null&&d.desv!==d.desvRec);

  // Buracos de 1 dia: pessoa ativa no dia anterior e no dia seguinte (do período), mas sem linha
  // no dia do meio. Sinal forte de erro de preenchimento em Labors do tipo "uma linha por dia".
  const diasUnicos=[...new Set(dias.map(d=>+d.d))].sort((a,b)=>a-b).map(t=>new Date(t));
  const ativosNoDia=(d)=>new Set(labor.filter(p=>p.ini<=d&&(p.fim==null||p.fim>=d)).map(p=>p.mat));
  const setsPorDia=diasUnicos.map(ativosNoDia);
  const buracos=[];
  for(let i=1;i<diasUnicos.length-1;i++){
    const antes=setsPorDia[i-1], atual=setsPorDia[i], depois=setsPorDia[i+1];
    antes.forEach(mat=>{
      if(depois.has(mat)&&!atual.has(mat)){
        const p=labor.find(x=>x.mat===mat);
        if(p)buracos.push({data:diasUnicos[i],nome:p.nome,mat,groot:p.groot});
      }
    });
  }

  const chk={grootZero,fimAntes,overlaps:[...new Set(overlaps)],truncados,desvDiv,fimPeriodo,buracos};
  window.__exp={labor,dias,ret,chk};
  recalcularPlano();
}

function recalcularPlano(){
  const {labor,dias,ret,chk}=window.__exp;
  const pausaDesde=getPausaDesde();
  const permitirAdiar=getPermitirAdiarInicio();
  const plano=equalizar(labor,ret,pausaDesde,permitirAdiar);
  // dif pós-plano por dia
  const eff=labor.map((p,i)=>{const a=plano.acoes.get(i)||{};return{...p,
    ini:a.novo_ini||p.ini, fim:a.novo_fim!==undefined?a.novo_fim:p.fim, off:!!a.retirar, pausas:a.pausas||[]}});
  const ativoNoDia=(p,d)=>{
    if(p.off||!(p.ini<=d&&(p.fim==null||p.fim>=d)))return false;
    return !p.pausas.some(pa=>pa.fim<d&&d<pa.ini);
  };
  dias.forEach(d=>{
    const isPT=d.tipo.includes('part');
    const c=eff.reduce((s,p)=>s+((p.elig&&p.pt===isPT&&ativoNoDia(p,d.d))?p.rateio:0),0);
    const alvo=d.qpos!=null?d.qpos:d.pref;
    d.alvoEq=alvo;
    d.ignorado=!(alvo>0);
    d.difPos=d.ignorado?null:(c-alvo);
  });

  render(dias,labor,chk,plano);
  window.__exp.plano=plano;
}

function reequalizar(){
  if(!window.__exp)return;
  recalcularPlano();
}

function getPermitirAdiarInicio(){
  const el=document.getElementById('fz-permitirAdiar');
  return el?el.checked:true;
}

function getPausaDesde(){
  const el=document.getElementById('fz-pausaDesde');
  if(!el||!el.value)return null;
  const [y,m,d]=el.value.split('-').map(Number);
  return new Date(Date.UTC(y,m-1,d));
}

function equalizar(labor,ret,pausaDesde,permitirAdiarInicio){
  if(permitirAdiarInicio===undefined)permitirAdiarInicio=true;
  const INF=new Date(Date.UTC(2099,0,1));
  const acoes=new Map(); // idx -> {retirar, novo_ini, novo_fim, pausas}
  const incluir={};
  const tipos=[...new Set(ret.map(r=>r.tipo.includes('part')))];
  for(const tipoPT of tipos){
    const dias=ret.filter(r=>r.tipo.includes('part')===tipoPT).sort((a,b)=>a.d-b.d);
    // Alvo por dia: Q Pós Comp. quando disponível, senão PREF
    const dlist=dias.map(r=>r.d), alvo=dias.map(r=>r.qpos!=null?r.qpos:r.pref);
    const pool=labor.map((p,i)=>({...p,i})).filter(p=>p.elig&&p.pt===tipoPT);
    const effIni=new Map(pool.map(p=>[p.i,p.ini]));
    const effFim=new Map(pool.map(p=>[p.i,p.fim??INF]));
    const act=(i,d)=>effIni.get(i)<=d&&d<=effFim.get(i);
    // exc>0: sobra gente (cortar); exc<0: falta (reportar); excesso residual = revisar
    // Dias com alvo 0 (escala não publicada / Go-Live) são ignorados: exc=0, nenhum ajuste.
    const ignorado=dlist.map((d,k)=>!(alvo[k]>0));
    const exc=dlist.map((d,k)=>ignorado[k]?0:(pool.reduce((s,p)=>s+(act(p.i,d)?p.rateio:0),0)-alvo[k]));
    const byI=new Map(pool.map(p=>[p.i,p]));
    // ordem para escolher quem cortar: início mais recente primeiro, depois matrícula desc
    const ordem=[...pool].sort((a,b)=>(b.ini-a.ini)||String(b.mat).localeCompare(String(a.mat))).map(p=>p.i);

    // Fase 1 — RETIRAR: processa dia a dia (do mais antigo ao mais novo). Em cada dia com excesso,
    // escolhe candidatos cujo intervalo de dias ativos está INTEIRAMENTE dentro do excesso disponível
    // em cada um desses dias (verificado contra o exc atual, nunca deixando nenhum dia ir a negativo).
    for(let k=0;k<dlist.length;k++){
      while(exc[k]>0){
        let best=null,bg=-1;
        for(const i of ordem){
          if(acoes.has(i)||!act(i,dlist[k]))continue;
          const ad=dlist.map((d,kk)=>act(i,d)?kk:-1).filter(kk=>kk>=0);
          if(ad.some(kk=>exc[kk]<byI.get(i).rateio))continue;  // não cabe em algum dia ativo
          if(ad.length>bg){bg=ad.length;best=i}
        }
        if(best==null)break;
        dlist.forEach((d,kk)=>{if(act(best,d))exc[kk]-=byI.get(best).rateio});
        acoes.set(best,{retirar:true});
      }
    }

    // Fase 2 — ADIAR INÍCIO: corta excesso no começo do contrato (desabilitável)
    if(permitirAdiarInicio){
      for(let k=0;k<dlist.length;k++){
        while(exc[k]>0){
          const i=ordem.find(i=>!acoes.has(i)&&+effIni.get(i)===+dlist[k]&&act(i,dlist[k]));
          if(i==null)break;
          let j=k;
          while(j<dlist.length&&exc[j]>0&&act(i,dlist[j]))j++;
          const novo=j<dlist.length?dlist[j]:new Date(+effFim.get(i)+864e5);
          for(let m=k;m<j;m++)exc[m]-=byI.get(i).rateio;
          if(j<dlist.length){acoes.set(i,{novo_ini:novo});effIni.set(i,novo);}
          else{acoes.set(i,{retirar:true});effIni.set(i,INF);} // empurrado pra fora do período = retirar
        }
      }
    }

    // Fase 3 — PAUSAR/RETOMAR: vales (excesso cercado de dias sem excesso, com retomada de demanda
    // depois) NÃO são tratados com corte definitivo — fecha o contrato no início do vale e reabre
    // no fim do vale, preservando GROOT/matrícula (mesma pessoa, sem duplicar nem recriar).
    const pausado=(i,d)=>{ const a=acoes.get(i); if(!a||!a.pausas)return false;
      return a.pausas.some(p=>p.fim<d&&d<p.ini); };
    const ativoP=(i,d)=>act(i,d)&&!pausado(i,d);
    for(let k=0;k<dlist.length;k++){
      while(exc[k]>0){
        if(!pausaDesde||dlist[k]<pausaDesde)break;   // sem data definida, ou antes dela: não pausa
        const candidatos=ordem.filter(i=>!acoes.get(i)?.retirar&&!acoes.get(i)?.novo_ini&&!acoes.get(i)?.novo_fim&&ativoP(i,dlist[k]));
        let aplicado=false;
        for(const i of candidatos){
          let j=k;
          while(j+1<dlist.length&&exc[j+1]>0&&ativoP(i,dlist[j+1]))j++;
          // só é vale (pausa) se o contrato seguir ativo DEPOIS do excesso E a demanda nesse próximo
          // dia ativo já não estiver em excesso (ou seja, há de fato uma retomada, não um fim disfarçado)
          const temRetomada=effFim.get(i)>dlist[j] && j+1<dlist.length && !ignorado[j+1];
          if(!temRetomada)continue;   // este candidato não serve para pausa — tenta o próximo
          const fechaEm=new Date(+dlist[k]-864e5);
          const reabreEm=new Date(+dlist[j]+864e5);
          for(let m=k;m<=j;m++)exc[m]-=byI.get(i).rateio;
          const a=acoes.get(i)||{};
          a.pausas=a.pausas||[]; a.pausas.push({fim:fechaEm,ini:reabreEm});
          acoes.set(i,a);
          aplicado=true;
          break;
        }
        if(!aplicado)break;   // nenhum candidato deste dia serve para pausa — passa para a Fase 4
      }
    }

    // Fase 4 — ENCURTAR FIM: corta excesso residual que sobra no fim do contrato (sem retomada).
    // Só aceita candidatos cujo FIM ORIGINAL já cai dentro do trecho de excesso sendo resolvido —
    // nunca corta quem teria dias futuros SEM excesso depois do corte (isso criaria falta artificial).
    for(let k=dlist.length-1;k>=0;k--){
      while(exc[k]>0){
        const candidatos=ordem.filter(i=>!acoes.has(i)&&act(i,dlist[k]));
        let aplicado=false;
        for(const i of candidatos){
          let j=k;
          while(j-1>=0&&exc[j-1]>=byI.get(i).rateio&&act(i,dlist[j-1]))j--;
          const novoFim=new Date(+dlist[j]-864e5);
          // o fim ORIGINAL do candidato precisa estar dentro de [dlist[j], dlist[k]] — ou seja,
          // ele já terminaria por ali de qualquer forma; senão, cortar criaria falta em dias futuros.
          const fimOriginal=effFim.get(i);
          if(fimOriginal>dlist[k])continue; // termina depois do trecho — não é candidato seguro a corte
          for(let m=j;m<=k;m++){ exc[m]-=byI.get(i).rateio; }
          if(novoFim<effIni.get(i)){acoes.set(i,{retirar:true});}
          else{acoes.set(i,{novo_fim:novoFim});effFim.set(i,novoFim);}
          aplicado=true;
          break;
        }
        if(!aplicado)break;
      }
    }

    // Falta (exc<0): só reportar. Excesso residual (exc>0): dia-vale sem solução, revisar manualmente.
    const res={}, revisar={};
    dlist.forEach((d,k)=>{ if(exc[k]<0)res[+d]=-exc[k]; else if(exc[k]>0)revisar[+d]=exc[k]; });
    const tipoNome=tipoPT?'PART_TIME':'FULL_TIME';
    if(Object.keys(res).length)incluir[tipoNome]={dias:res};
    if(Object.keys(revisar).length)(incluir.__revisar=incluir.__revisar||{})[tipoNome]=revisar;
  }
  return {acoes,incluir};
}

function renderPlano(plano,labor){
  const G=[];
  const grp=(cls,chip,titulo,hint,rows,extra)=>G.push(
    `<div class="pgroup" onclick="this.classList.toggle('open')">
      <header><span class="chip ${cls}">${chip}</span><b>${titulo}</b><span class="n">${rows.length||extra||0}</span></header>
      ${hint?`<div class="hint">${hint}</div>`:''}
      ${rows.length?`<div class="plist">${rows.join('')}</div>`:''}
    </div>`);
  const ret=[],adi=[],enc=[],pau=[];
  plano.acoes.forEach((a,i)=>{
    const p=labor[i];
    const linha=det=>`<div class="prow"><div class="nm">${p.nome}<small>${p.cargo} · GROOT ${p.groot||'—'} · Mat. ${p.mat||'—'}</small></div><div class="dt">${fmt(p.ini)} – ${p.fim?fmt(p.fim):'em aberto'}<small>${det}</small></div></div>`;
    if(a.retirar)ret.push(linha('linha removida'));
    else if(a.novo_ini)adi.push(linha('início → '+fmt(a.novo_ini)));
    else if(a.novo_fim!==undefined)enc.push(linha('fim → '+(a.novo_fim?fmt(a.novo_fim):'em branco')));
    else if(a.pausas&&a.pausas.length)pau.push(linha(a.pausas.map(pa=>`pausa de ${fmt(pa.fim)} a ${fmt(pa.ini)}`).join(' · ')));
  });
  if(ret.length)grp('ch-ret','Retirar',`Retirado do Labor — excesso sobre o alvo do MELI`,'Estas linhas foram <b>removidas</b> do arquivo ajustado.',ret);
  if(adi.length)grp('ch-adi','Adiar início',`DATA DE INÍCIO movida`,'O início foi <b>adiado</b> para eliminar excesso no começo do contrato.',adi);
  if(enc.length)grp('ch-est','Encurtar fim',`DATA FIM antecipada`,'A DATA FIM foi <b>antecipada</b> para eliminar excesso no fim do contrato.',enc);
  if(pau.length)grp('ch-pau','Pausar/Retomar',`Vale temporário — contrato pausado e retomado`,'A demanda caiu e voltou ao normal. Em vez de retirar a pessoa, o contrato foi <b>fechado</b> no início do vale e <b>reaberto</b> no fim — mesma pessoa, mesmo GROOT/matrícula, duas linhas no arquivo.',pau);
  const inc=Object.entries(plano.incluir).filter(([k])=>k!=='__revisar');
  inc.forEach(([tipo,o])=>{
    const dias_obj=o.dias;
    const totalMax=Math.max(...Object.values(dias_obj));
    const cells=Object.entries(dias_obj).map(([ts,q])=>{const d=new Date(+ts);
      return `<div class="inc-cell"><div class="q">+${fmtN(q)}</div><div class="d">${fmt(d)}</div></div>`}).join('');
    const uid='inc_'+tipo.replace('_','');
    G.push(`<div class="pgroup open" id="grp_${uid}">
      <header onclick="document.getElementById('grp_${uid}').classList.toggle('open')"><span class="chip ch-inc">Incluir</span><b>Pessoas que faltam no arquivo (${tipo.replace('_',' ')})</b><span class="n">${totalMax}</span></header>
      <div class="hint">Estas pessoas <b>não existem no Labor</b>. Suba a base de diaristas (modelo padrão) — o sistema seleciona automaticamente a quantidade necessária por dia e gera o arquivo pronto para importar.</div>
      <div class="inc-grid">${cells}</div>
      <div style="padding:12px 16px 14px;border-top:1px solid var(--line)">
        <div id="diarista_status_${uid}" style="font-size:.75rem;color:var(--mut);margin-bottom:8px">Nenhuma base carregada.</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
          <label class="dlbtn" style="cursor:pointer;background:var(--panel);color:var(--ink);border:1px solid var(--amber);font-size:.8rem;padding:9px 16px;gap:7px">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width:15px;height:15px"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            Carregar base de diaristas
            <input type="file" accept=".xlsx,.xls" style="display:none" onchange="Fusao.carregarDiaristas(event,'${uid}','${tipo}')">
          </label>
          <button id="btn_diarista_${uid}" class="dlbtn" style="display:none;font-size:.8rem;padding:9px 16px;gap:7px" onclick="Fusao.exportarDiaristas('${uid}','${tipo}')">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width:15px;height:15px"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Baixar Labor de Diaristas
          </button>
        </div>
      </div>
    </div>`);
    // guarda os dias necessários por tipo para uso posterior
    window.__incDias=window.__incDias||{};
    window.__incDias[tipo]=dias_obj;
  });
  // dias-vale: excesso que não dá pra cortar sem partir contrato
  const rev=plano.incluir.__revisar;
  if(rev){
    Object.entries(rev).forEach(([tipo,o])=>{
      const cells=Object.entries(o).map(([ts,q])=>{const d=new Date(+ts);
        return `<div class="inc-cell" style="border-color:var(--warn)"><div class="q" style="color:var(--warn)">+${fmtN(q)}</div><div class="d">${fmt(d)}</div></div>`}).join('');
      G.push(`<div class="pgroup open"><header><span class="chip" style="background:var(--warn);color:#1a1305">Revisar</span><b>Revisar manualmente (${tipo.replace('_',' ')})</b><span class="n">${Object.keys(o).length}</span></header>
        <div class="hint">Nestes dias o alvo do MELI <b>cai e volta a subir</b> (vale isolado). Zerar exigiria partir um contrato em dois períodos — o app não faz isso automaticamente. Avalie caso a caso se há alguém que de fato trabalhou só nesse intervalo. Esses dias saem na aba 'REVISAR' do arquivo.</div>
        <div class="inc-grid">${cells}</div></div>`);
    });
  }
  // botão de download
  const btn=document.getElementById('fz-btnExport');
  const temAcao=plano.acoes.size>0||inc.length>0||!!rev;
  if(btn){btn.style.display=temAcao?'inline-flex':'none';btn.classList.remove('done');
    btn.innerHTML=btn.innerHTML.replace('Baixado ✓','Baixar Labor ajustado (.xlsx)');}
  if(!G.length)G.push('<div class="alert green"><b>Nada a ajustar</b><p>Seu Labor já está equalizado com o PREF do MELI em todos os dias.</p></div>');
  document.getElementById('fz-plano').innerHTML=G.join('');
}


const fmtN=n=>Number.isInteger(n)?n:n.toFixed(2);

function render(dias,labor,chk,plano){
  document.getElementById('fz-drop').classList.add('hidden');
  document.getElementById('fz-result').classList.remove('hidden');
  document.getElementById('fz-sec-hcm').classList.remove('hidden');
  const okD=dias.filter(d=>d.st==='ok').length,
        meuD=dias.filter(d=>d.st==='meu').length,
        meliD=dias.filter(d=>d.st==='meli').length;

  const ban=document.getElementById('fz-banner');
  if(meuD===0&&chk.overlaps.length===0&&chk.fimAntes.length===0){
    ban.className='banner ok';
    ban.innerHTML=`<div><div class="big">✔ Seu Labor está correto</div><p>A contagem diária bate com a Qtd. PREF em todos os ${dias.length} dias. ${meliD>0?`Os desvios existentes em ${meliD} dia(s) são do lado MELI (escala não reconhecida no Scheduling/Rostering) — a cobrança é para o BC deles, não ajuste seu.`:'Nenhuma ação necessária.'}</p></div>`;
  } else {
    ban.className='banner '+(meuD>0?'bad':'warn');
    ban.innerHTML=`<div><div class="big">${meuD>0?'✖ Há divergência no seu Labor':'⚠ Labor bate, mas há pontos de atenção'}</div><p>${meuD>0?`Em ${meuD} de ${dias.length} dias a contagem do seu Labor não bate com a Qtd. PREF que o MELI está usando. Veja o diagnóstico dia a dia abaixo — em geral é versão desatualizada ou DATA FIM cortada antes do fim do período (${fmt(chk.fimPeriodo)}).`:'Confira as verificações do arquivo abaixo.'}</p></div>`;
  }

  document.getElementById('fz-cards').innerHTML=`
    <div class="card"><div class="v ok">${okD}</div><div class="l">dias OK</div></div>
    <div class="card"><div class="v ${meuD?'bad':'ok'}">${meuD}</div><div class="l">dias com divergência sua</div></div>
    <div class="card"><div class="v ${meliD?'warn':'ok'}">${meliD}</div><div class="l">dias com desvio lado MELI</div></div>
    <div class="card"><div class="v amber">${labor.length}</div><div class="l">linhas no labor</div></div>`;

  renderPlano(plano,labor);
  renderLineChart(dias);
  renderDonut(okD,meuD,meliD);
  document.getElementById('fz-strip').innerHTML=dias.map(d=>
    `<div class="cell c-${d.st==='ok'?'ok':d.st==='meu'?'meu':'meli'}" title="${fmt(d.d)} — ${d.diag}">${d.d.getUTCDate()}<small>${['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'][d.d.getUTCMonth()]}</small></div>`).join('');

  document.querySelector('#fz-tbl tbody').innerHTML=dias.map(d=>{
    const cls=d.st==='ok'?'okr':d.st==='meu'?'bad':'warnr';
    const difCls=d.dif>0?'pos':d.dif<0?'neg':'zero';
    return `<tr class="${cls}"><td>${fmt(d.d)}</td><td>${d.pref}</td><td>${fmtN(d.meu)}</td>
      <td class="${difCls}">${d.dif>0?'+':''}${fmtN(d.dif)}</td><td>${fmtN(d.outroPT)}</td>
      <td>${d.qpos??'—'}</td><td>${d.desv??'—'}</td>
      <td>${d.desvRec??'—'}${d.desv!=null&&d.desvRec!=null&&d.desv!==d.desvRec?' ⚠':''}</td>
      <td class="${d.difPos==null?'':d.difPos===0?'zero':d.difPos>0?'pos':'neg'}">${d.difPos==null?'<span style="color:var(--mut)">—</span>':(d.difPos>0?'+':'')+fmtN(d.difPos)}</td>
      <td class="diag">${d.diag}</td></tr>`}).join('');

  renderChecks(chk);
}

// =================== GRÁFICOS ===================
const S_LABOR=getComputedStyle(document.documentElement).getPropertyValue('--s-labor').trim()||'#3987e5';
const S_ALVO=getComputedStyle(document.documentElement).getPropertyValue('--s-alvo').trim()||'#c98500';

function renderLineChart(dias){
  const el=document.getElementById('fz-chartLinha');
  if(!el)return;
  if(!dias.length){el.innerHTML='<p style="font-size:.78rem;color:var(--mut)">Sem dados.</p>';return;}
  const W=640,H=230,padL=32,padR=14,padT=12,padB=26;
  const n=dias.length;
  const v1=dias.map(d=>d.meu);
  const v2=dias.map(d=>d.qpos!=null?d.qpos:d.pref);
  const maxV=Math.max(1,...v1,...v2);
  const x=i=>n<=1?padL:padL+(i/(n-1))*(W-padL-padR);
  const y=v=>H-padB-(v/maxV)*(H-padT-padB);
  const steps=4,grid=[],xticks=[];
  for(let s=0;s<=steps;s++){
    const v=maxV*s/steps,yy=y(v);
    grid.push(`<line x1="${padL}" x2="${W-padR}" y1="${yy.toFixed(1)}" y2="${yy.toFixed(1)}" stroke="var(--line)" stroke-width="1"/>`);
    grid.push(`<text x="${(padL-6).toFixed(1)}" y="${(yy+3).toFixed(1)}" text-anchor="end" font-size="9" fill="var(--mut)" font-family="var(--mono)">${Math.round(v)}</text>`);
  }
  const tickEvery=Math.max(1,Math.round(n/6));
  for(let i=0;i<n;i+=tickEvery)xticks.push(`<text x="${x(i).toFixed(1)}" y="${H-6}" text-anchor="middle" font-size="9" fill="var(--mut)" font-family="var(--mono)">${fmt(dias[i].d)}</text>`);
  const pathFor=vals=>vals.map((v,i)=>(i===0?'M':'L')+x(i).toFixed(1)+','+y(v).toFixed(1)).join(' ');
  const endLabel=(vals,color,dy)=>{const i=n-1,vv=vals[i];
    return `<text x="${(x(i)+6).toFixed(1)}" y="${(y(vv)+dy).toFixed(1)}" font-size="10" font-weight="700" font-family="var(--mono)" fill="${color}">${fmtN(vv)}</text>`;};
  const closeEnds=Math.abs(y(v1[n-1])-y(v2[n-1]))<12;
  el.innerHTML=`
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Gráfico: Seu Labor comparado ao alvo MELI por dia">
      ${grid.join('')}
      <path d="${pathFor(v2)}" fill="none" stroke="${S_ALVO}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
      <path d="${pathFor(v1)}" fill="none" stroke="${S_LABOR}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
      ${xticks.join('')}
      ${endLabel(v2,S_ALVO,closeEnds?-8:4)}
      ${endLabel(v1,S_LABOR,closeEnds?10:4)}
      <g id="lc-hover" style="opacity:0">
        <line id="lc-cross" x1="0" x2="0" y1="${padT}" y2="${H-padB}" stroke="var(--mut)" stroke-width="1" stroke-dasharray="3,3"/>
        <circle id="lc-dot2" r="5" fill="${S_ALVO}" stroke="var(--navy2)" stroke-width="2"/>
        <circle id="lc-dot1" r="5" fill="${S_LABOR}" stroke="var(--navy2)" stroke-width="2"/>
      </g>
      <rect id="lc-capture" x="${padL}" y="${padT}" width="${Math.max(1,W-padL-padR)}" height="${H-padT-padB}" fill="transparent"/>
    </svg>
    <div class="chart-tip" id="lc-tip"></div>`;
  const svg=el.querySelector('svg'),cap=el.querySelector('#lc-capture'),hov=el.querySelector('#lc-hover'),
        cross=el.querySelector('#lc-cross'),dot1=el.querySelector('#lc-dot1'),dot2=el.querySelector('#lc-dot2'),
        tip=el.querySelector('#lc-tip');
  const idxAt=svgX=>{const t=(svgX-padL)/(W-padL-padR);return Math.max(0,Math.min(n-1,Math.round(t*(n-1))));};
  const move=e=>{
    const r=svg.getBoundingClientRect();
    const svgX=(e.clientX-r.left)/r.width*W;
    const i=idxAt(svgX);
    const px=x(i),py1=y(v1[i]),py2=y(v2[i]);
    cross.setAttribute('x1',px);cross.setAttribute('x2',px);
    dot1.setAttribute('cx',px);dot1.setAttribute('cy',py1);
    dot2.setAttribute('cx',px);dot2.setAttribute('cy',py2);
    hov.style.opacity='1';
    const wrapR=el.getBoundingClientRect();
    tip.style.left=((px/W)*wrapR.width)+'px';
    tip.style.top=((Math.min(py1,py2)/H)*wrapR.height)+'px';
    tip.innerHTML=`<b>${fmt(dias[i].d)}</b><br><span style="color:${S_LABOR}">●</span> Seu Labor: <b>${fmtN(v1[i])}</b><br><span style="color:${S_ALVO}">●</span> Alvo MELI: <b>${fmtN(v2[i])}</b>`;
    tip.classList.add('show');
  };
  cap.addEventListener('mousemove',move);
  cap.addEventListener('mouseleave',()=>{hov.style.opacity='0';tip.classList.remove('show');});
}

function renderDonut(okD,meuD,meliD){
  const el=document.getElementById('fz-chartDonut');
  if(!el)return;
  const total=okD+meuD+meliD;
  if(!total){el.innerHTML='<p style="font-size:.78rem;color:var(--mut)">Sem dados.</p>';return;}
  const segs=[
    {v:okD,color:'var(--ok)',label:'Dias OK'},
    {v:meuD,color:'var(--bad)',label:'Divergência sua'},
    {v:meliD,color:'var(--warn)',label:'Desvio lado MELI'}
  ];
  const r=48,cx=64,cy=64,C=2*Math.PI*r;
  let acc=0;
  const arcs=segs.filter(s=>s.v>0).map(s=>{
    const frac=s.v/total,len=frac*C,gap=Math.min(2,len*0.06);
    const dash=`${Math.max(0,len-gap).toFixed(2)} ${(C-len+gap).toFixed(2)}`;
    const rot=-90+acc/total*360;
    acc+=s.v;
    return `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${s.color}" stroke-width="18"
      stroke-dasharray="${dash}" stroke-dashoffset="0" transform="rotate(${rot.toFixed(2)} ${cx} ${cy})" stroke-linecap="butt"/>`;
  }).join('');
  const legend=segs.map(s=>`<div class="key"><i class="sw" style="background:${s.color}"></i>${s.label} <b>${s.v}</b> <span style="color:var(--mut)">(${Math.round(s.v/total*100)}%)</span></div>`).join('');
  el.innerHTML=`
    <svg width="128" height="128" viewBox="0 0 128 128" role="img" aria-label="Distribuição dos dias por status">
      <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--line)" stroke-width="18" opacity=".5"/>
      ${arcs}
      <text x="${cx}" y="${cy-3}" text-anchor="middle" font-size="20" font-weight="700" fill="var(--ink)" font-family="var(--mono)">${total}</text>
      <text x="${cx}" y="${cy+13}" text-anchor="middle" font-size="8" fill="var(--mut)" letter-spacing=".5">DIAS</text>
    </svg>
    <div class="donut-legend">${legend}</div>`;
}

// =================== ABAS DE VERIFICAÇÃO ===================
function renderChecks(chk){
  const cats=[
    {key:'overlaps',label:'GROOT duplicado',cls:'red',
      title:n=>`GROOT ID com períodos sobrepostos — ${n} caso(s)`,
      desc:'A mesma pessoa conta duas vezes no mesmo dia (típico de efetivação temporário → CLT com datas cruzadas). Ajuste para o fim de um contrato ser anterior ao início do outro.',
      ok:'Nenhum GROOT ID com períodos sobrepostos.',
      names:()=>chk.overlaps.join(' · ')},
    {key:'grootZero',label:'GROOT zerado/faltante',cls:'',
      title:n=>`GROOT ID zerado — ${n} pessoa(s)`,
      desc:'Risco de deduplicação no sistema do MELI. Solicite o GROOT correto no portal antes do próximo envio.',
      ok:'Nenhuma linha com GROOT ID zerado ou vazio.',
      names:()=>chk.grootZero.map(p=>p.nome).join(' · ')},
    {key:'fimAntes',label:'Datas invertidas',cls:'red',
      title:n=>`DATA FIM anterior à DATA DE INÍCIO — ${n} linha(s)`,
      desc:'Corrigir antes de enviar; o MELI descarta ou zera essas pessoas.',
      ok:'Nenhuma linha com DATA FIM anterior à DATA DE INÍCIO.',
      names:()=>chk.fimAntes.map(p=>p.nome).join(' · ')},
    {key:'buracos',label:'Buracos de 1 dia',cls:'red',
      title:n=>`Buraco de 1 dia — ${n} caso(s)`,
      desc:'A pessoa está ativa no dia anterior e no dia seguinte, mas falta a linha exatamente neste dia — sinal forte de erro de preenchimento (típico de Labor com uma linha por dia trabalhado). Adicione a linha do dia faltante.',
      ok:'Nenhum buraco de 1 dia detectado.',
      names:()=>chk.buracos.slice(0,40).map(b=>`${fmt(b.data)}: ${b.nome}`).join(' · ')+(chk.buracos.length>40?' …':'')},
    {key:'truncados',label:'Cortes de período',cls:'',
      title:n=>`Possível corte de DATA FIM antes do fim do período — ${n} pessoa(s)`,
      desc:`O período do retorno vai até ${fmt(chk.fimPeriodo)}, mas essas pessoas têm DATA FIM anterior e nenhuma linha sucessora. Se seguem trabalhando, deixe a DATA FIM em branco ou use a data real.`,
      ok:'Nenhum corte de período suspeito.',
      names:()=>chk.truncados.slice(0,40).map(p=>`${p.nome} (fim ${fmt(p.fim)})`).join(' · ')+(chk.truncados.length>40?' …':'')},
    {key:'desvDiv',label:'Desvio divergente',cls:'',
      title:n=>`Desvio informado ≠ desvio recalculado — ${n} dia(s)`,
      desc:'Nesses dias o MELI calculou o desvio com outra versão do PREF. Vale pedir ao BC qual versão está vigente.',
      ok:'Nenhuma divergência entre desvio informado e recalculado.',
      names:()=>chk.desvDiv.map(d=>fmt(d.d)).join(' · ')}
  ];
  const counts=cats.map(c=>(chk[c.key]||[]).length);

  const tabsHtml=cats.map((c,i)=>
    `<button class="tab${i===0?' active':''}${counts[i]?' has-issue':''}" data-tab="chk-${c.key}" onclick="Fusao.switchChkTab('${c.key}')">
      ${c.label} <span class="cnt">${counts[i]}</span>
    </button>`).join('');

  const panelsHtml=cats.map((c,i)=>{
    const n=counts[i];
    const body=n
      ? `<div class="alert ${c.cls}"><b>${c.title(n)}</b><p>${c.desc}</p><div class="names">${c.names()}</div></div>`
      : `<div class="alert green"><b>✔ Tudo certo</b><p>${c.ok}</p></div>`;
    return `<div class="tabpanel${i===0?' active':''}" id="chk-${c.key}">${body}</div>`;
  }).join('');

  document.getElementById('fz-chkTabs').innerHTML=tabsHtml;
  document.getElementById('fz-chkPanels').innerHTML=panelsHtml;
}
function switchChkTab(key){
  document.querySelectorAll('#fz-chkTabs .tab').forEach(t=>t.classList.toggle('active',t.dataset.tab==='chk-'+key));
  document.querySelectorAll('#fz-chkPanels .tabpanel').forEach(p=>p.classList.toggle('active',p.id==='chk-'+key));
}

function exportarLabor(){
  const ctx=window.__labCtx, exp=window.__exp;
  if(!ctx||!exp){showErr('Faça a auditoria de um arquivo antes de exportar.');return;}
  const {header,iI,iF}=ctx;
  const {labor,plano}=exp;

  const toSerial=d=>{ if(d==null)return ''; const base=Date.UTC(1899,11,30);
    return Math.round((Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())-base)/864e5); };

  const aoa=[header.slice()];
  labor.forEach((p,i)=>{
    const a=plano.acoes.get(i)||{};
    if(a.retirar)return;                       // RETIRAR: linha sai do arquivo

    // monta os segmentos [início,fim] desta pessoa após adiar/encurtar e pausas
    let ini0=a.novo_ini!=null?a.novo_ini:p.ini;
    let fim0=a.novo_fim!==undefined?a.novo_fim:p.fim;
    const pausas=(a.pausas||[]).slice().sort((x,y)=>x.fim-y.fim);
    const segs=[]; let cursorIni=ini0;
    pausas.forEach(pa=>{ segs.push({ini:cursorIni,fim:pa.fim}); cursorIni=pa.ini; });
    segs.push({ini:cursorIni,fim:fim0});

    segs.forEach(seg=>{
      const row=p.rawRow.slice();
      row[iI]=toSerial(seg.ini);
      row[iF]=seg.fim==null?'':toSerial(seg.fim);
      aoa.push(row);
    });
  });

  const wb=XLSX.utils.book_new();
  const wsL=XLSX.utils.aoa_to_sheet(aoa);
  const range=XLSX.utils.decode_range(wsL['!ref']);
  for(let r=1;r<=range.e.r;r++){
    [iI,iF].forEach(ci=>{ const ref=XLSX.utils.encode_cell({r,c:ci}); const cell=wsL[ref];
      if(cell&&typeof cell.v==='number'){cell.t='n';cell.z='dd/mm/yyyy';} });
  }
  XLSX.utils.book_append_sheet(wb,wsL,'Labor enviado ao MELI');

  const incAoa=[['Tipo','Data','Pessoas a incluir','(preencher na aba Labor: GROOT ID, NOME, MATRICULA, REGIME, CARGO, INÍCIO, FIM, % RATEIO, DIAS, ESCALA)']];
  Object.entries(plano.incluir).filter(([k])=>k!=='__revisar').forEach(([tipo,o])=>{
    Object.entries(o.dias).sort((a,b)=>a[0]-b[0]).forEach(([ts,q])=>{
      incAoa.push([tipo.replace('_',' '),toSerial(new Date(+ts)),q,'']);
    });
  });
  if(incAoa.length===1) incAoa.push(['(nada a incluir)','','','']);
  const wsI=XLSX.utils.aoa_to_sheet(incAoa);
  for(let r=1;r<incAoa.length;r++){ const ref=XLSX.utils.encode_cell({r,c:1}); const cell=wsI[ref];
    if(cell&&typeof cell.v==='number'){cell.t='n';cell.z='dd/mm/yyyy';} }
  wsI['!cols']=[{wch:14},{wch:12},{wch:16},{wch:78}];
  XLSX.utils.book_append_sheet(wb,wsI,'A_INCLUIR');

  // aba REVISAR (dias-vale com excesso residual)
  const rev=plano.incluir.__revisar;
  if(rev){
    const revAoa=[['Tipo','Data','Excesso a revisar','Observação']];
    Object.entries(rev).forEach(([tipo,o])=>{
      Object.entries(o).sort((a,b)=>a[0]-b[0]).forEach(([ts,q])=>{
        revAoa.push([tipo.replace('_',' '),toSerial(new Date(+ts)),q,
          'Alvo do MELI cai e sobe (vale). Zerar exigiria partir contrato — avaliar manualmente.']);
      });
    });
    const wsR=XLSX.utils.aoa_to_sheet(revAoa);
    for(let r=1;r<revAoa.length;r++){ const ref=XLSX.utils.encode_cell({r,c:1}); const cell=wsR[ref];
      if(cell&&typeof cell.v==='number'){cell.t='n';cell.z='dd/mm/yyyy';} }
    wsR['!cols']=[{wch:14},{wch:12},{wch:16},{wch:62}];
    XLSX.utils.book_append_sheet(wb,wsR,'REVISAR');
  }

  const hoje=new Date(), pad=n=>String(n).padStart(2,'0');
  const nome=`Labor_ajustado_${hoje.getFullYear()}${pad(hoje.getMonth()+1)}${pad(hoje.getDate())}.xlsx`;
  XLSX.writeFile(wb,nome);

  const btn=document.getElementById('fz-btnExport');
  if(btn){btn.classList.add('done');
    btn.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg> Baixado ✓';}
}
// =================== BASE DE DIARISTAS ===================

// Armazena {uid -> [{groot, nome, data, cargo, escala, empresa}]}
window.__baseDiaristas={};

function carregarDiaristas(evt, uid, tipo){
  const file=evt.target.files[0]; if(!file)return;
  const status=document.getElementById('diarista_status_'+uid);
  status.textContent='Lendo arquivo…'; status.style.color='var(--mut)';
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
      // aceita qualquer aba; procura cabeçalho com "GROOT" e "DATA"
      let rows=null, hdr=null, hdri=-1;
      for(const nm of wb.SheetNames){
        const r=XLSX.utils.sheet_to_json(wb.Sheets[nm],{header:1,defval:''});
        for(let i=0;i<Math.min(r.length,5);i++){
          const c=r[i].map(v=>String(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim());
          if(c.some(x=>x.includes('groot'))&&c.some(x=>x.includes('nome'))){hdr=c;hdri=i;rows=r;break;}
        }
        if(rows)break;
      }
      if(!rows){status.textContent='❌ Cabeçalho não reconhecido — o arquivo precisa ter colunas GROOT ID e NOME.';status.style.color='var(--bad)';return;}
      const ci=(hdr,...alts)=>{for(const a of alts){const j=hdr.findIndex(c=>c.includes(a));if(j>=0)return j;}return -1;};
      const iG=ci(hdr,'groot'), iN=ci(hdr,'nome'), iD=ci(hdr,'data'), iC=ci(hdr,'cargo'), iE=ci(hdr,'escala','turno'), iEmp=ci(hdr,'empresa');
      const base=[];
      for(let i=hdri+1;i<rows.length;i++){
        const row=rows[i]; if(!row||row.every(v=>v===''))continue;
        const groot=String(row[iG]??'').trim(); if(!groot||groot==='0')continue;
        let data=null;
        if(iD>=0&&row[iD]!==''){
          if(row[iD] instanceof Date) data=new Date(Date.UTC(row[iD].getFullYear(),row[iD].getMonth(),row[iD].getDate()));
          else if(typeof row[iD]==='number'){const d=XLSX.SSF.parse_date_code(row[iD]);if(d)data=new Date(Date.UTC(d.y,d.m-1,d.d));}
          else{const m=String(row[iD]).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);if(m){let y=+m[3];if(y<100)y+=2000;data=new Date(Date.UTC(y,+m[2]-1,+m[1]));}}
        }
        base.push({groot, nome:String(row[iN]??'').trim(),
          data, cargo:iC>=0?String(row[iC]??'').trim():'', escala:iE>=0?String(row[iE]??'').trim():'',
          empresa:iEmp>=0?String(row[iEmp]??'').trim():''});
      }
      window.__baseDiaristas[uid]=base;
      // checar cobertura: quantos dias têm pelo menos 1 diarista disponível
      const dias=window.__incDias&&window.__incDias[tipo]?window.__incDias[tipo]:{};
      const cobertura=Object.entries(dias).filter(([ts,q])=>{
        const d=new Date(+ts);
        return base.some(p=>!p.data||+p.data===+d);
      }).length;
      const total=Object.keys(dias).length;
      status.innerHTML=`✔ <b>${base.length}</b> diaristas carregados — cobertura em <b>${cobertura}/${total}</b> dias com necessidade.`;
      status.style.color='var(--ok)';
      document.getElementById('btn_diarista_'+uid).style.display='inline-flex';
    }catch(err){status.textContent='❌ Erro ao ler o arquivo: '+err.message;status.style.color='var(--bad)';}
  };
  reader.readAsArrayBuffer(file);
}

function exportarDiaristas(uid, tipo){
  const base=window.__baseDiaristas[uid]||[];
  const dias=window.__incDias&&window.__incDias[tipo]?window.__incDias[tipo]:{};
  if(!base.length){alert('Carregue a base de diaristas primeiro.');return;}

  const toSerial=d=>{if(!d)return '';const base0=Date.UTC(1899,11,30);
    return Math.round((Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())-base0)/864e5);};

  // cabeçalho do modelo de extração
  const aoa=[['GROOT ID','NOME','MATRICULA','REGIME DE CONTRATO ','CARGO','DATA DE INÍCIO','DATA FIM','% RATEIO','DIAS TRABALHADOS X FOLGA','ESCALA']];
  const avisos=[];

  // ordena os dias cronologicamente
  const diasOrdenados=Object.entries(dias).sort((a,b)=>+a[0]-+b[0]);

  // rastrea por GROOT quais datas já foram usadas (evita duplicar o mesmo diarista no mesmo dia)
  const usados=new Set(); // "groot_serial"

  for(const [ts, qtd] of diasOrdenados){
    const d=new Date(+ts);
    const serial=toSerial(d);
    // candidatos para este dia: diaristas da base cuja data (se informada) bate, ou sem data (disponível qualquer dia)
    const cands=base.filter(p=>{
      if(usados.has(p.groot+'_'+serial))return false; // já alocado neste dia
      if(!p.data)return true; // sem data específica = disponível qualquer dia
      return +p.data===+d;
    });
    const selecionados=cands.slice(0,qtd);
    const faltam=qtd-selecionados.length;
    selecionados.forEach(p=>{
      usados.add(p.groot+'_'+serial);
      aoa.push([p.groot, p.nome, p.groot, 'TEMPORÁRIO',
        'Auxiliar de Apoio LOG I', serial, serial, 1, '', p.escala||'']);
    });
    if(faltam>0) avisos.push({data:d,faltam,disponiveis:cands.length});
  }

  const wb2=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  // formatar colunas de data
  for(let r=1;r<aoa.length;r++){
    ['E5','F5'].forEach((_,ci)=>{
      const ref=XLSX.utils.encode_cell({r,c:5+ci});
      const cell=ws[ref]; if(cell&&typeof cell.v==='number'){cell.t='n';cell.z='dd/mm/yyyy';}
    });
    // col 5 (DATA DE INÍCIO) e 6 (DATA FIM)
    [5,6].forEach(ci=>{
      const ref=XLSX.utils.encode_cell({r,c:ci});
      const cell=ws[ref]; if(cell&&typeof cell.v==='number'){cell.t='n';cell.z='dd/mm/yyyy';}
    });
  }
  ws['!cols']=[{wch:12},{wch:36},{wch:12},{wch:14},{wch:22},{wch:14},{wch:14},{wch:10},{wch:22},{wch:16}];
  XLSX.utils.book_append_sheet(wb2,ws,'Labor enviado ao MELI');

  // aba de avisos se houver dias sem cobertura completa
  if(avisos.length){
    const avAoa=[['Data','Necessidade no dia','Disponíveis na base','Faltando','Observação']];
    avisos.forEach(a=>{
      avAoa.push([toSerial(a.data), a.faltam+a.disponiveis, a.disponiveis, a.faltam,
        'Incluir manualmente os '+a.faltam+' restante(s)']);
    });
    const wsA=XLSX.utils.aoa_to_sheet(avAoa);
    for(let r=1;r<avAoa.length;r++){const ref=XLSX.utils.encode_cell({r,c:0});const cell=wsA[ref];if(cell&&typeof cell.v==='number'){cell.t='n';cell.z='dd/mm/yyyy';}}
    wsA['!cols']=[{wch:14},{wch:16},{wch:18},{wch:10},{wch:44}];
    XLSX.utils.book_append_sheet(wb2,wsA,'DIAS SEM COBERTURA');
  }

  const hoje=new Date(),pad=n=>String(n).padStart(2,'0');
  const nomearq=`Labor_Diaristas_${hoje.getFullYear()}${pad(hoje.getMonth()+1)}${pad(hoje.getDate())}.xlsx`;
  XLSX.writeFile(wb2,nomearq);

  const btn=document.getElementById('btn_diarista_'+uid);
  if(btn){btn.classList.add('done');
    btn.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg> Baixado ✓';}
  if(avisos.length) setTimeout(()=>alert(`⚠ ${avisos.length} dia(s) sem cobertura completa na base. Confira a aba "DIAS SEM COBERTURA" do arquivo.`),300);
}
// =================== CONFRONTO HCM MELI ===================

window.normGroot=g=>String(g||'').trim().replace(/\.0$/,'');

window.__hcmLinhas = []; // linhas do MELI não cobertas

function carregarHCM(evt){
  const file=evt.target.files[0]; if(!file)return;
  const st=document.getElementById('hcm-status');
  st.textContent='Lendo arquivo…'; st.style.color='var(--mut)';
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
      // procura a aba de detalhe
      const sheetNm=wb.SheetNames.find(n=>n.toLowerCase().includes('detalle')||n.toLowerCase().includes('detalhe'))||wb.SheetNames[0];
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetNm],{header:1,defval:''});
      // achar cabeçalho
      let hdri=-1, hdr=null;
      for(let i=0;i<Math.min(rows.length,5);i++){
        const c=rows[i].map(v=>String(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim());
        if(c.some(x=>x.includes('groot'))&&c.some(x=>x.includes('colaborador')||x.includes('nome'))){
          hdri=i; hdr=c; break;
        }
      }
      if(!hdr){st.textContent='❌ Cabeçalho não reconhecido. Precisa de colunas Groot ID e Colaborador.';st.style.color='var(--bad)';return;}
      const ci=(...alts)=>{for(const a of alts){const j=hdr.findIndex(c=>c.includes(a));if(j>=0)return j;}return -1;};
      const iG=ci('groot'), iN=ci('colaborador','nome'), iI=ci('inicio','start'), iF=ci('fin','fim','end'), iC=ci('contrato','contrat');
      const parseDate=v=>{
        if(!v||v==='')return null;
        if(v instanceof Date)return new Date(Date.UTC(v.getFullYear(),v.getMonth(),v.getDate()));
        if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);return d?new Date(Date.UTC(d.y,d.m-1,d.d)):null;}
        // formatos dd/mm/yyyy ou yyyy-mm-dd
        const m=String(v).match(/(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{1,4})/);
        if(!m)return null;
        const p1=+m[1],p2=+m[2],p3=+m[3];
        if(p1>1000) return new Date(Date.UTC(p1,p2-1,p3));   // yyyy-mm-dd
        return new Date(Date.UTC(p3<100?p3+2000:p3,p2-1,p1)); // dd/mm/yyyy
      };
      const hcm=[];
      for(let i=hdri+1;i<rows.length;i++){
        const row=rows[i]; if(!row||row.every(v=>v===''))continue;
        const groot=window.normGroot(row[iG]); if(!groot||groot==='0')continue;
        hcm.push({groot, nome:String(row[iN]??'').trim(),
          ini:parseDate(row[iI]), fim:parseDate(row[iF]),
          contrato:iC>=0?String(row[iC]??'').trim():''});
      }
      processarHCM(hcm, sheetNm, file.name);
    }catch(err){st.textContent='❌ Erro: '+err.message;st.style.color='var(--bad)';}
  };
  reader.readAsArrayBuffer(file);
}

function processarHCM(hcm, sheetNm, fileName){
  const st=document.getElementById('hcm-status');
  const exp=window.__exp;
  if(!exp){st.textContent='❌ Audite o Labor primeiro antes de carregar o HCM.';st.style.color='var(--bad)';return;}
  const labor=exp.labor||[];

  // Construir conjunto de coberturas já existentes: {groot -> [{ini,fim}]}
  // Fontes: 1) Labor principal  2) todas as bases de diaristas carregadas
  const coberturas={}; // groot -> [{ini, fim}]
  const addCob=(groot,ini,fim)=>{
    const g=window.normGroot(groot); if(!g||g==='0')return;
    if(!coberturas[g])coberturas[g]=[];
    coberturas[g].push({ini:ini||new Date(0), fim:fim||new Date(Date.UTC(2099,0,1))});
  };

  // Labor principal
  labor.forEach(p=>{
    const a=(exp.plano&&exp.plano.acoes)?exp.plano.acoes.get(labor.indexOf(p)):null;
    if(a&&a.retirar)return; // ignorar retirados
    const ini=a&&a.novo_ini?a.novo_ini:p.ini;
    const fim=a&&a.novo_fim!==undefined?a.novo_fim:p.fim;
    // pausas geram múltiplos segmentos
    const pausas=(a&&a.pausas)||[];
    if(!pausas.length){addCob(p.groot,ini,fim);}
    else{
      let cur=ini;
      pausas.forEach(pa=>{addCob(p.groot,cur,pa.fim); cur=pa.ini;});
      addCob(p.groot,cur,fim);
    }
  });

  // Diaristas carregados em todas as bases
  Object.values(window.__baseDiaristas||{}).forEach(base=>{
    base.forEach(p=>{
      const d=p.data;
      addCob(p.groot, d, d); // diarista: ini=fim=dia
    });
  });

  // Verificar cada linha do HCM
  const sobrepoe=(ini1,fim1,ini2,fim2)=>{
    const f1=fim1||new Date(Date.UTC(2099,0,1));
    const f2=fim2||new Date(Date.UTC(2099,0,1));
    return ini1<=f2 && ini2<=f1;
  };

  const ausentes=[], cobertos=[], semDatas=[];
  hcm.forEach(h=>{
    if(!h.ini&&!h.fim){semDatas.push(h);return;}
    const cobs=coberturas[h.groot]||[];
    const coberto=cobs.some(c=>sobrepoe(h.ini,h.fim,c.ini,c.fim));
    if(coberto)cobertos.push(h);
    else ausentes.push(h);
  });

  window.__hcmLinhas=ausentes;

  // Renderizar resultado
  const res=document.getElementById('hcm-result');
  res.classList.remove('hidden');
  st.innerHTML=`✔ <b>${hcm.length}</b> linhas lidas de "${sheetNm}" — <b style="color:var(--bad)">${ausentes.length}</b> ausentes no Labor, <b style="color:var(--ok)">${cobertos.length}</b> já cobertos.`;
  st.style.color='var(--ok)';

  const ban=document.getElementById('hcm-banner');
  if(!ausentes.length){
    ban.className='banner ok';
    ban.innerHTML=`<div><div class="big">✔ Todos cobertos</div><p>Todos os colaboradores do HCM já existem no Labor ou na base de diaristas. Nenhuma ação necessária.</p></div>`;
  }else{
    ban.className='banner bad';
    ban.innerHTML=`<div><div class="big">✖ ${ausentes.length} colaborador(es) não encontrado(s) no Labor</div><p>Estas pessoas constam no relatório HCM do MELI (${fileName}) mas <b>não têm cobertura</b> no Labor atual nem na base de diaristas — para nenhum período que se sobreponha. O RH precisa revisar e incluir.</p></div>`;
  }

  // Cards resumo
  const grootsAusentes=new Set(ausentes.map(a=>a.groot)).size;
  document.getElementById('hcm-cards').innerHTML=`
    <div class="card"><div class="v bad">${ausentes.length}</div><div class="l">Linhas ausentes no Labor</div></div>
    <div class="card"><div class="v bad">${grootsAusentes}</div><div class="l">Groot IDs distintos ausentes</div></div>
    <div class="card"><div class="v ok">${cobertos.length}</div><div class="l">Linhas já cobertas</div></div>
    <div class="card"><div class="v amber">${hcm.length}</div><div class="l">Total no HCM</div></div>`;

  // Tabela
  const fmt2=d=>d?String(d.getUTCDate()).padStart(2,'0')+'/'+String(d.getUTCMonth()+1).padStart(2,'0')+'/'+d.getUTCFullYear():'—';
  const tbody=document.querySelector('#hcm-tbl tbody');
  tbody.innerHTML=ausentes.map(h=>`<tr>
    <td style="font-family:var(--mono)">${h.groot}</td>
    <td style="text-align:left">${h.nome}</td>
    <td>${h.contrato||'—'}</td>
    <td>${fmt2(h.ini)}</td>
    <td>${fmt2(h.fim)}</td>
    <td><span style="font-size:.7rem;color:var(--bad);font-weight:700">AUSENTE NO LABOR</span></td>
  </tr>`).join('');

  if(!ausentes.length) document.getElementById('hcm-btn-export').style.display='none';
}

function exportarHCM(){
  const linhas=window.__hcmLinhas||[];
  if(!linhas.length){alert('Nenhuma linha ausente para exportar.');return;}
  const toSerial=d=>{if(!d)return '';const base=Date.UTC(1899,11,30);
    return Math.round((Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())-base)/864e5);};
  const aoa=[['Groot ID','Colaborador','Contrato','Data de início (MELI)','Data de fim (MELI)','Observação']];
  linhas.forEach(h=>aoa.push([h.groot,h.nome,h.contrato||'',toSerial(h.ini),toSerial(h.fim),'Não encontrado no Labor nem na base de diaristas']));
  const wb2=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  for(let r=1;r<aoa.length;r++){
    [3,4].forEach(ci=>{const ref=XLSX.utils.encode_cell({r,c:ci});const cell=ws[ref];if(cell&&typeof cell.v==='number'){cell.t='n';cell.z='dd/mm/yyyy';}});
  }
  ws['!cols']=[{wch:12},{wch:38},{wch:12},{wch:18},{wch:18},{wch:50}];
  XLSX.utils.book_append_sheet(wb2,ws,'Ausentes no Labor');
  const hoje=new Date(),pad=n=>String(n).padStart(2,'0');
  XLSX.writeFile(wb2,`HCM_Ausentes_${hoje.getFullYear()}${pad(hoje.getMonth()+1)}${pad(hoje.getDate())}.xlsx`);
}

/* ================================================================
   SUPERFÍCIE PÚBLICA — só o que os handlers inline do HTML usam
   ================================================================ */
window.Fusao={reequalizar,exportarLabor,carregarHCM,exportarHCM,
  switchChkTab,carregarDiaristas,exportarDiaristas};
})();
