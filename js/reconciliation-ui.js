/* Ajustes MELI — interface do modo "Conciliar duas faturas"
   ================================================================

   Fica dentro de uma IIFE e publica em window.Recon só o que os
   handlers do HTML chamam. O modo 1 (projetar ajustes) continua no
   escopo global, intocado — os dois convivem na mesma aba.

   O fluxo é sempre: detectar → apontar → explicar → o usuário decide →
   prévia → gerar cópia nova. Nenhuma linha é alterada na importação.
   ================================================================ */
"use strict";
(function(){

const $=id=>document.getElementById(id);
const esc2=s=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/"/g,"&quot;");

/* Estado do modo. Vive enquanto a página estiver aberta; as decisões do
   usuário ficam aqui e não são persistidas em lugar nenhum. */
const R={
  origem:null,      // {employees, comp, fileName, buffer, sheetName, map, confident}
  seguinte:null,
  result:null,
  filtro:"TODOS",
  aberto:null       // id do item expandido
};

/* ================================================================
   1. LEITURA DAS DUAS FATURAS
   ================================================================ */
function lerFatura(data, fileName){
  const wb=XLSX.read(data,{type:"array",cellDates:true});
  const read=readLaborSheet(wb);
  if(read.erro) return {erro:read.erro};
  const employees=parseEmployees(read.rows, read.map);
  if(!employees.length) return {erro:"A aba de Labor não tem linhas legíveis."};

  /* A competência é buscada na ordem de força da evidência — campo da
     planilha, nome do arquivo, sequência com a outra fatura, heurística —
     e a fonte usada fica visível na tela. Ver js/competence-source.js. */
  const retornoHint=readRetornoHint(wb);
  const det=resolveCompetence({wb, employees, fileName, retornoHint});
  return {employees, comp:det.comp, fonte:det.fonte, fonteDetalhe:det.detalhe,
    confianca:det.confianca, fileName,
    buffer:data.slice().buffer, sheetName:read.sheetName, map:read.map, retornoHint};
}

/* PRIORIDADE 3: com as duas faturas na mão, a sequência é evidência mais forte
   que a heurística de datas — mas nunca sobrepõe um campo explícito nem uma
   escolha manual do usuário. */
const FONTE_FRACA=f=>!f||f.id==="HEURISTICA"||f.id==="VAZIO";
function reforcarPorSequencia(){
  if(!R.origem||!R.seguinte) return;
  const forte=q=>R[q]&&!FONTE_FRACA(R[q].fonte);
  if(forte("origem")&&FONTE_FRACA(R.seguinte.fonte)){
    const nm=proximoMes(R.origem.comp);
    R.seguinte.comp=buildCompetence(nm.y,nm.m);
    R.seguinte.fonte=COMP_SOURCE.SEQUENCIA;
    R.seguinte.fonteDetalhe="assumida a partir de "+R.origem.comp.label;
    R.seguinte.confianca=COMP_SOURCE.SEQUENCIA.confianca;
  } else if(forte("seguinte")&&FONTE_FRACA(R.origem.fonte)){
    const ma=mesAnterior(R.seguinte.comp);
    R.origem.comp=buildCompetence(ma.y,ma.m);
    R.origem.fonte=COMP_SOURCE.SEQUENCIA;
    R.origem.fonteDetalhe="assumida a partir de "+R.seguinte.comp.label;
    R.origem.confianca=COMP_SOURCE.SEQUENCIA.confianca;
  }
}

/* O usuário manda na competência: a detecção é um palpite, não um veredito. */
function setComp(qual, y, m){
  if(!R[qual]) return;
  R[qual].comp=buildCompetence(y,m);
  R[qual].fonte=COMP_SOURCE.MANUAL;
  R[qual].fonteDetalhe="escolhida por você";
  R[qual].confianca=COMP_SOURCE.MANUAL.confianca;
  R.result=null;
  renderSlot(qual); renderSequencia(); renderResultado();
}
function assumirSeguinte(){
  if(!R.origem||!R.seguinte) return;
  const c=R.origem.comp;
  const nm=c.m===12?{y:c.y+1,m:1}:{y:c.y,m:c.m+1};
  setComp("seguinte", nm.y, nm.m);
}

function renderSlot(qual){
  const r=R[qual], box=$("rc-"+qual+"-info");
  if(!r){ box.className="rc-slot-info"; box.textContent="Nenhum arquivo carregado."; return; }
  box.className="rc-slot-info ok";
  const CONF={ALTA:"Alta",MEDIA:"Média",REVISAO:"Revisão necessária"};
  box.innerHTML=`<b>${esc2(r.fileName)}</b>
    <span class="rc-linhas">${r.employees.length} linhas de Labor</span>
    <span class="rc-comp-linha">Competência
      <select onchange="Recon.setComp('${qual}',null,this.value)">${
        MONTHS.map((nome,i)=>`<option value="${i+1}"${i+1===r.comp.m?" selected":""}>${nome}</option>`).join("")
      }</select>
      <select onchange="Recon.setComp('${qual}',this.value,null)">${
        Array.from({length:5},(_,k)=>r.comp.y-2+k).map(y=>`<option value="${y}"${y===r.comp.y?" selected":""}>${y}</option>`).join("")
      }</select>
    </span>
    <span class="rc-fonte">
      <span>Fonte: <b>${esc2(r.fonte?r.fonte.label:"—")}</b>${r.fonteDetalhe?` <i>(${esc2(r.fonteDetalhe)})</i>`:""}</span>
      <span class="rc-conf-tag c-${r.confianca}">Confiança: ${esc2(CONF[r.confianca]||r.confianca)}</span>
    </span>`;
}

/* Arrastar e soltar nos dois slots. O <input type=file> continua ali para quem
   prefere clicar; o drop só oferece o mesmo caminho por outra via. */
function initDropZones(){
  ["origem","seguinte"].forEach(qual=>{
    const slot=document.querySelector(`#modo-conciliar .rc-slot[data-slot="${qual}"]`);
    if(!slot) return;
    const zona=slot.querySelector(".rc-slot-drop");
    const parar=ev=>{ ev.preventDefault(); ev.stopPropagation(); };
    ["dragenter","dragover"].forEach(t=>zona.addEventListener(t,ev=>{
      parar(ev); ev.dataTransfer.dropEffect="copy"; zona.classList.add("over"); }));
    ["dragleave","dragend"].forEach(t=>zona.addEventListener(t,ev=>{
      parar(ev); zona.classList.remove("over"); }));
    zona.addEventListener("drop",ev=>{
      parar(ev); zona.classList.remove("over");
      const f=ev.dataTransfer.files&&ev.dataTransfer.files[0];
      if(f) receber(qual,f);
    });
  });
  /* soltar fora da zona não deve fazer o navegador abrir o arquivo */
  ["dragover","drop"].forEach(t=>document.addEventListener(t,ev=>{
    if(ev.target.closest&&ev.target.closest("#modo-conciliar")) return;
    if(ev.type==="drop") ev.preventDefault();
  }));
}

function carregar(qual, ev){
  const f=ev.target.files[0];
  ev.target.value="";
  if(f) receber(qual,f);
}

function receber(qual, f){
  const box=$("rc-"+qual+"-info");
  box.className="rc-slot-info carregando";
  box.textContent="Lendo "+f.name+"…";
  const reader=new FileReader();
  reader.onload=e=>{
    let r;
    try{ r=lerFatura(new Uint8Array(e.target.result), f.name); }
    catch(err){ r={erro:"Falha ao ler o arquivo: "+err.message}; }
    if(r.erro){
      box.className="rc-slot-info erro";
      box.textContent=r.erro;
      R[qual]=null;
    } else {
      R[qual]=r;
      reforcarPorSequencia();
      renderSlot("origem"); renderSlot("seguinte");
    }
    R.result=null;
    renderSequencia();
    renderResultado();
  };
  reader.readAsArrayBuffer(f);
}

/* ================================================================
   2. SEQUÊNCIA DAS COMPETÊNCIAS
   ================================================================ */
function renderSequencia(){
  const el=$("rc-seq");
  const btn=$("rc-btnConciliar");
  if(!R.origem||!R.seguinte){
    el.className="rc-seq";
    el.innerHTML=`<span class="rc-seq-msg">Carregue as duas faturas para conferir a sequência.</span>`;
    btn.disabled=true;
    return;
  }
  const s=checkSequence(R.origem.comp, R.seguinte.comp);
  el.className="rc-seq "+(s.ok?"ok":"warn");
  el.innerHTML=`<span class="rc-seq-par">${esc2(R.origem.comp.label)} <i>→</i> ${esc2(R.seguinte.comp.label)}</span>
    <span class="rc-seq-msg">${s.ok?"✓ ":"⚠ "}${esc2(s.msg)}</span>
    ${s.kind==="invertida"?`<button class="rc-inline-btn" onclick="Recon.trocarOrdem()">Trocar ordem</button>`:""}
    ${s.ok?"":`<button class="rc-inline-btn" onclick="Recon.assumirSeguinte()">Assumir ${esc2(R.origem.comp.nextLabel)} na seguinte</button>`}
    ${s.ok?"":`<span class="rc-seq-nota">Ou corrija a competência em cada fatura. Você pode prosseguir mesmo assim.</span>`}`;
  btn.disabled=false;    // nunca bloqueia: o usuário decide
}

function trocarOrdem(){
  const t=R.origem; R.origem=R.seguinte; R.seguinte=t;
  R.result=null;
  renderSlot("origem"); renderSlot("seguinte");
  renderSequencia();
  renderResultado();
}

/* ================================================================
   3. CONCILIAÇÃO
   ================================================================ */
function conciliar(){
  if(!R.origem||!R.seguinte) return;
  R.result=reconcile({origem:R.origem, seguinte:R.seguinte});
  R.filtro="TODOS"; R.aberto=null;
  renderResultado();
  $("rc-resultado").scrollIntoView({behavior:"smooth",block:"nearest"});
}

/* ================================================================
   4. TELA DE RESULTADO
   ================================================================ */
const FILTROS=[
  {k:"TODOS",      rot:"Todos"},
  {k:"CONCILIADO", rot:"Conciliados"},
  {k:"AUSENTE",    rot:"Ajustes ausentes"},
  {k:"DIVERGENTE", rot:"Divergentes"},
  {k:"DUPLICADO",  rot:"Duplicados"},
  {k:"SEM_ORIGEM", rot:"Retroativos sem origem"},
  {k:"INDETERMINADO", rot:"Indeterminados"},
  {k:"REVISAO",    rot:"Revisão manual"}
];
function passaFiltro(it,f){
  if(f==="TODOS") return true;
  if(f==="DIVERGENTE") return it.alerts.includes(RECON_STATUS.PERIODO)
    ||it.alerts.includes(RECON_STATUS.FTE)||it.alerts.includes(RECON_STATUS.SINAL);
  if(f==="REVISAO") return it.status===RECON_STATUS.REVISAO||it.status===RECON_STATUS.AMBIGUA;
  return it.status===f;
}
function contaFiltro(f){ return R.result.items.filter(it=>passaFiltro(it,f)).length; }

function renderResultado(){
  const el=$("rc-resultado");
  if(!R.result){ el.innerHTML=""; el.classList.add("hidden"); return; }
  el.classList.remove("hidden");
  const {items, summary, contexto}=R.result;

  let html=`<section>
    <h2>Resultado da conciliação</h2>
    <div class="cards">
      <div class="card"><div class="v amber">${summary.total}</div><div class="l">total analisado</div></div>
      <div class="card"><div class="v ok">${summary.conciliados}</div><div class="l">conciliados</div></div>
      <div class="card"><div class="v ${summary.pendencias?"bad":"ok"}">${summary.pendencias}</div><div class="l">pendências</div></div>
      <div class="card"><div class="v ${summary.descontosAusentes?"bad":"ok"}">${summary.descontosAusentes}</div><div class="l">descontos ausentes</div></div>
      <div class="card"><div class="v ${summary.acrescimosAusentes?"bad":"ok"}">${summary.acrescimosAusentes}</div><div class="l">acréscimos ausentes</div></div>
      <div class="card"><div class="v ${summary.divergencias?"warn":"ok"}">${summary.divergencias}</div><div class="l">divergências</div></div>
    </div>
    <p class="rc-contexto">Reconstruídos <b>${contexto.esperados}</b> ajustes devidos a partir de
      ${esc2(contexto.compN.label)} (${contexto.linhasOrigem} linhas) e lidos
      <b>${contexto.retroativosNaSeguinte}</b> lançamentos retroativos em
      ${esc2(contexto.compNext.label)} (${contexto.linhasSeguinte} linhas).
      ${contexto.ignoradas?`<b>${contexto.ignoradas}</b> linha${contexto.ignoradas===1?"":"s"} sem nome,
        sem identificador e sem período — rodapé, subtotal ou formatação — ${contexto.ignoradas===1?"foi ignorada":"foram ignoradas"}.`:""}
      Nada foi alterado — as decisões abaixo são suas.</p>
  </section>`;

  html+=`<section>`+legenda(items)+`<div class="rc-filtros">`+
    FILTROS.map(f=>`<button class="rc-filtro${R.filtro===f.k?" active":""}" onclick="Recon.setFiltro('${f.k}')">
      ${f.rot} <span class="cnt">${contaFiltro(f.k)}</span></button>`).join("")+
    `</div>`;

  const vis=items.filter(it=>passaFiltro(it,R.filtro));
  if(!vis.length){
    html+=`<div class="rc-vazio">Nenhum apontamento neste filtro.</div>`;
  } else {
    html+=`<div class="tablebox"><table class="rc-tabela"><thead><tr>
      <th>Status</th><th>Colaborador</th><th>Esperado</th><th>Encontrado</th>
      <th class="num">Diferença</th><th>Decisão</th>
    </tr></thead><tbody>`;
    vis.forEach(it=>{ html+=linhaItem(it); });
    html+=`</tbody></table></div>`;
  }
  html+=`</section>`;

  html+=blocoPrevia();
  el.innerHTML=html;
}

/* Por que cada status existe e o que ele pede de você. Só aparecem os status
   presentes no resultado — legenda de coisa que não está na tela é ruído. */
const EXPLICA={
  CONCILIADO:"O ajuste devido foi encontrado na fatura seguinte, fechando identidade, competência de origem, período, sinal e FTE. Nada a fazer.",
  AUSENTE:"A fatura de origem obrigava um ajuste e nenhuma linha correspondente foi encontrada na seguinte. É a pendência mais direta: ou o lançamento faltou, ou aconteceu fora destas duas faturas.",
  PERIODO_DIVERGENTE:"O lançamento existe e está no sentido certo, mas cobre um período diferente do devido. O trecho exato que não fechou é apontado no detalhe.",
  FTE_DIVERGENTE:"Período e sentido corretos, rateio diferente do esperado. Só aparece quando o rateio realmente difere — um período menor produzindo FTE menor não conta como divergência.",
  SINAL_INCORRETO:"O período confere, mas o lançamento está invertido: desconto onde caberia acréscimo, ou o contrário.",
  DUPLICADO:"O mesmo período aparece mais de uma vez na fatura seguinte. O app nunca escolhe qual linha remover — isso é decisão sua.",
  SEM_ORIGEM:"Existe um retroativo na fatura seguinte que nenhuma movimentação das duas faturas explica. Não quer dizer que esteja errado: a causa pode estar numa base de RH que não foi carregada. Requer validação.",
  IDENTIDADE_AMBIGUA:"A mesma matrícula aparece ligada a mais de um GROOT entre as duas faturas. Os registros não são unidos automaticamente — confirme quem é quem antes de agir.",
  INDETERMINADO:"A linha tem um período que poderia ser retroativo, mas os sinais se contradizem. Enquanto não for esclarecida, ela fica fora de todas as comparações.",
  REVISAO:"A fatura de origem não traz um dado necessário para calcular o ajuste devido, ou há mais de uma leitura possível para a correspondência. O app prefere avisar a chutar."
};
function legenda(items){
  const presentes=[...new Set(items.map(i=>i.status))]
    .sort((a,b)=>ORDEM_LEGENDA.indexOf(a)-ORDEM_LEGENDA.indexOf(b));
  if(!presentes.length) return "";
  return `<details class="rc-legenda"><summary>O que cada status significa (${presentes.length} em uso)</summary>
    <dl>${presentes.map(st=>{
      const m=RECON_META[st]||{label:st,tone:"info",icon:"·"};
      return `<dt class="t-${m.tone}">${m.icon} ${esc2(m.label)}</dt><dd>${esc2(EXPLICA[st]||"")}</dd>`;
    }).join("")}</dl></details>`;
}
const ORDEM_LEGENDA=["AUSENTE","SINAL_INCORRETO","DUPLICADO","PERIODO_DIVERGENTE","FTE_DIVERGENTE",
  "SEM_ORIGEM","IDENTIDADE_AMBIGUA","INDETERMINADO","REVISAO","CONCILIADO"];

function linhaItem(it){
  const meta=RECON_META[it.status]||{label:it.status,tone:"info",icon:"·"};
  const conf={ALTA:"Alta confiança",MEDIA:"Média confiança",REVISAO:"Revisão necessária"}[it.confianca];
  const dif=it.diffDias!==0
    ? (it.diffDias>0?"+":"")+it.diffDias+" dia"+(Math.abs(it.diffDias)===1?"":"s")
    : (Math.abs(it.diffFte)>1e-9?fmtFteUI(it.diffFte):"0");
  const aberto=R.aberto===it.id;
  const extras=(it.alerts||[]).filter(a=>a!==it.status)
    .map(a=>(RECON_META[a]||{label:a}).label);
  let h=`<tr class="rc-linha t-${meta.tone}${aberto?" aberta":""}" onclick="Recon.toggle(${it.id})">
    <td><span class="rc-status t-${meta.tone}">${meta.icon} ${meta.label}</span>
        ${extras.map(e=>`<span class="rc-alerta-extra">+ ${esc2(e)}</span>`).join("")}
        <span class="rc-conf c-${it.confianca}" title="${esc2(it.confiancaMotivo||"")}">${conf}</span></td>
    <td class="rc-nome">${esc2(it.nome)}<small>${esc2(it.groot?"GROOT "+it.groot:(it.matricula?"Mat. "+it.matricula:"sem identificador"))}</small></td>
    <td class="rc-per">${it.esperado?descreverAjusteUI(it.esperado):"—"}</td>
    <td class="rc-per">${it.achado?descreverAjusteUI(it.achado):"—"}</td>
    <td class="num rc-dif">${esc2(dif)}</td>
    <td onclick="event.stopPropagation()">${seletorDecisao(it)}</td>
  </tr>`;
  if(aberto) h+=detalhe(it);
  return h;
}

const DECISOES=[["MANTER","Manter como está"],["ACEITAR","Aceitar sugestão"],
  ["IGNORAR","Ignorar apontamento"],["REVISAR","Revisar manualmente"]];
function seletorDecisao(it){
  const temSugestao=!!sugerirCorrecao(it,null);
  return `<select class="rc-decisao d-${it.decisao}" onchange="Recon.decidir(${it.id},this.value)">`+
    DECISOES.map(([v,r])=>`<option value="${v}"${it.decisao===v?" selected":""}${v==="ACEITAR"&&!temSugestao?" disabled":""}>${r}</option>`).join("")+
    `</select>`;
}

function detalhe(it){
  return `<tr class="rc-detalhe"><td colspan="6">
    ${timeline(it)}
    <div class="rc-narrativa">
      ${bloco("Cobrança original em "+esc2(it.compOrigem), it.cobrancaOriginal
        ? (it.cobrancaOriginal.cobrado
            ? `<p class="rc-forte">${fmtShort(it.cobrancaOriginal.start)} → ${fmtShort(it.cobrancaOriginal.end)}</p>
               <p>${it.cobrancaOriginal.days} dia${it.cobrancaOriginal.days===1?"":"s"} · FTE ${fmtFteUI(it.cobrancaOriginal.fte)}</p>
               <p class="rc-base">${esc2(it.cobrancaOriginal.base)}</p>`
            : `<p class="rc-mut">Não cobrada nesta competência.</p>
               <p class="rc-base">${esc2(it.cobrancaOriginal.base)}</p>`)
        : `<p class="rc-mut">Sem linha de origem identificada nesta fatura.</p>`)}

      ${bloco("Movimentação inferida", it.movimentacaoInferida
        ? `<p class="rc-forte">${esc2(it.movimentacaoInferida.texto)}</p>
           <p class="rc-base">Base da inferência: ${esc2(it.movimentacaoInferida.base)}</p>`
        : `<p class="rc-mut">Nenhuma movimentação posterior ao corte foi inferida.</p>`)}

      ${bloco("Retroativo esperado em "+esc2(it.compAplicacao), it.esperado
        ? `<p class="rc-forte">${it.esperado.kind}</p>
           <p>${fmtShort(it.esperado.start)} → ${fmtShort(it.esperado.end)} · ${it.esperado.days} dia${it.esperado.days===1?"":"s"}</p>
           <p><b>FTE esperado:</b> ${fmtFteUI(it.esperado.fte)}</p>`
        : `<p class="rc-mut">A regra do corte não prevê ajuste para esta pessoa.</p>`)}

      ${bloco("Retroativo encontrado", it.achado
        ? `<p class="rc-forte">${it.achado.kind}</p>
           <p>${fmtShort(it.achado.start)} → ${fmtShort(it.achado.end)} · ${it.achado.days} dia${it.achado.days===1?"":"s"}</p>
           <p><b>FTE:</b> ${fmtFteUI(it.achado.fte)}${it.achado.srcRow?` · linha ${it.achado.srcRow}`:""}</p>
           ${it.achados.length>1?`<p class="rc-mut">${it.achados.length} lançamentos: ${it.achados.map(x=>"linha "+x.srcRow).join(", ")}</p>`:""}`
        : `<p class="rc-mut">Nenhum lançamento retroativo correspondente.</p>`)}
    </div>

    ${it.achado&&it.achado.classeMotivo?`<div class="rc-classe">
      <b>Linha interpretada como ${esc2((typeof LINE_CLASS_LABEL!=="undefined"&&LINE_CLASS_LABEL[it.achado.classe])||it.achado.classe)}</b>
      <p>${esc2(it.achado.classeMotivo)}</p></div>`:""}
    ${!it.achado&&it.lineClassReason?`<div class="rc-classe">
      <b>Linha interpretada como ${esc2((typeof LINE_CLASS_LABEL!=="undefined"&&LINE_CLASS_LABEL[it.lineClassification])||"Indeterminado")}</b>
      <p>${esc2(it.lineClassReason)}</p></div>`:""}

    ${dimensoes(it)}

    <div class="rc-diag"><b>Diagnóstico</b><p>${esc2(it.diagnostico)}</p></div>
    ${it.impacto?`<div class="rc-impacto">${it.impacto.calculado
      ? `<b>Impacto financeiro</b> · original ${brl(it.impacto.original)} · esperado ${brl(it.impacto.esperado)} · encontrado ${brl(it.impacto.encontrado)} · diferença ${brl(it.impacto.diferenca)}`
      : esc2(it.impacto.motivo)}</div>`:""}
    ${it.decisao==="ACEITAR"?painelAjuste(it):painelSugestaoDisponivel(it)}
  </td></tr>`;
}

function bloco(titulo, corpo){
  return `<div class="rc-col"><h4>${titulo}</h4>${corpo}</div>`;
}
function brl(v){
  return (v===null||v===undefined)?"—":v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
}

/* Checklist das dimensões comparadas — mostra o que fechou e o que não. */
const DIM_ROT={identidade:"Identidade", competencia:"Competência", periodo:"Período",
  sinal:"Sinal", fte:"FTE"};
function dimensoes(it){
  const d=it.dimensoes||{};
  const itens=Object.keys(DIM_ROT).map(k=>{
    const v=d[k];
    const ic=v===true?"✓":v===false?"✗":"–";
    const cls=v===true?"ok":v===false?"bad":"na";
    return `<span class="rc-dim ${cls}"><i>${ic}</i>${DIM_ROT[k]}</span>`;
  }).join("");
  return `<div class="rc-dims">${itens}
    <span class="rc-dim-conf">Confiança: <b>${{ALTA:"Alta",MEDIA:"Média",REVISAO:"Revisão necessária"}[it.confianca]}</b>${it.confiancaMotivo?" — "+esc2(it.confiancaMotivo):""}</span>
  </div>`;
}

/* Linha do tempo da competência de origem: corte, período cobrado, ajuste. */
function timeline(it){
  const first=it.compFirst, last=it.compLast, dias=it.compDias;
  const pos=v=>((ymdParts(v).d-1)/(dias-1))*100;
  const faixa=(a,b,cls,rot)=>{
    const x=pos(Math.max(a,first)), w=Math.max(1.2,pos(Math.min(b,last))-x);
    return `<div class="rc-tl-faixa ${cls}" style="left:${x}%;width:${w}%" title="${rot}"></div>`;
  };
  const marcas=[1,5,10,15,20,25,dias].filter((v,i,a)=>a.indexOf(v)===i);
  const cobr=it.cobrancaOriginal&&it.cobrancaOriginal.cobrado?it.cobrancaOriginal:null;
  const cortePos=pos(it.compSnapshot);
  return `<div class="rc-tl">
    <div class="rc-tl-titulo">${esc2(it.compOrigem)}</div>
    <div class="rc-tl-trilho">
      ${cobr?faixa(cobr.start,cobr.end,"cobr","cobrado originalmente"):""}
      ${it.esperado?faixa(it.esperado.start,it.esperado.end,it.esperado.kind==="DESCONTAR"?"esp-neg":"esp-pos","ajuste esperado"):""}
      ${it.achado?faixa(it.achado.start,it.achado.end,"ach","ajuste encontrado"):""}
      <div class="rc-tl-corte" style="left:${cortePos}%"><i></i><span>corte ${fmtShort(it.compSnapshot)}</span></div>
    </div>
    <div class="rc-tl-eixo">${marcas.map(d=>`<span style="left:${((d-1)/(dias-1))*100}%">${d}</span>`).join("")}</div>
    <div class="rc-tl-legenda">
      ${cobr?`<span><i class="cobr"></i>cobrado em ${esc2(it.compOrigem)}</span>`:""}
      ${it.esperado?`<span><i class="${it.esperado.kind==="DESCONTAR"?"esp-neg":"esp-pos"}"></i>ajuste esperado</span>`:""}
      ${it.achado?`<span><i class="ach"></i>ajuste encontrado</span>`:""}
    </div>
  </div>`;
}

function painelSugestaoDisponivel(it){
  const s=sugerirCorrecao(it,null);
  if(!s) return `<div class="rc-nada">Sem correção automática a propor — este apontamento é informativo.</div>`;
  return `<div class="rc-proposta">
    <b>Sugestão disponível</b>
    <span>${esc2(descreverProposta(s))}</span>
    <span class="rc-mut">Selecione “Aceitar sugestão” para revisar e editar antes de aplicar.</span>
  </div>`;
}

function descreverProposta(s){
  if(s.acao==="INCLUIR")    return "Incluir "+(s.kind==="DESCONTAR"?"desconto":"acréscimo")+" de "
    +fmtShort(s.start)+" a "+fmtShort(s.end)+" ("+s.days+" dia"+(s.days===1?"":"s")+", FTE "+fmtFteUI(s.fte)+")";
  if(s.acao==="SUBSTITUIR") return "Alterar a linha "+s.alvoRow+" para "+fmtShort(s.start)+" a "+fmtShort(s.end)
    +" ("+s.days+" dia"+(s.days===1?"":"s")+", FTE "+fmtFteUI(s.fte)+")";
  if(s.acao==="REMOVER")    return "Remover a(s) linha(s) "+(s.alvoRows||[]).join(", ")+", mantendo a linha "+s.manterRow;
  return "";
}

/* Ajuste fino: o usuário edita a sugestão antes de ela entrar na prévia. */
function painelAjuste(it){
  const s=it.sugestao; if(!s) return "";
  if(s.acao==="REMOVER") return painelRemocao(it,s);
  const podeEscolherModo=it.alerts.includes(RECON_STATUS.PERIODO)&&it.achado;
  return `<div class="rc-ajuste">
    <b>Ajuste fino</b>
    ${podeEscolherModo?`<div class="rc-modos">
      <label><input type="radio" name="modo${it.id}" value="SUBSTITUIR"${s.acao==="SUBSTITUIR"?" checked":""}
        onchange="Recon.modo(${it.id},'SUBSTITUIR')"> Substituir a linha existente pelo período esperado</label>
      <label><input type="radio" name="modo${it.id}" value="COMPLEMENTAR"${s.acao==="INCLUIR"?" checked":""}
        onchange="Recon.modo(${it.id},'COMPLEMENTAR')"> Complementar somente ${it.faltantes.length?esc2(it.faltantes.map(f=>f.days===1?fmtShort(f.start):fmtShort(f.start)+" a "+fmtShort(f.end)).join(" e ")):"a diferença"}</label>
      <label><input type="radio" name="modo${it.id}" value="MANTER"
        onchange="Recon.decidir(${it.id},'MANTER')"> Manter o original, sem alterar nada</label>
      <label><input type="radio" name="modo${it.id}" value="REVISAR"
        onchange="Recon.decidir(${it.id},'REVISAR')"> Revisar manualmente</label>
    </div>`:""}
    <div class="rc-campos">
      <label>Tipo
        <select onchange="Recon.editar(${it.id},'kind',this.value)">
          <option value="DESCONTAR"${s.kind==="DESCONTAR"?" selected":""}>Desconto</option>
          <option value="ACRESCENTAR"${s.kind==="ACRESCENTAR"?" selected":""}>Acréscimo</option>
        </select></label>
      <label>Início <input type="date" value="${isoDe(s.start)}" onchange="Recon.editar(${it.id},'start',this.value)"></label>
      <label>Fim <input type="date" value="${isoDe(s.end)}" onchange="Recon.editar(${it.id},'end',this.value)"></label>
      <label>% Rateio <input type="number" step="0.01" min="0" max="1" value="${s.rateio}" onchange="Recon.editar(${it.id},'rateio',this.value)"></label>
      <label class="rc-obs">Observação <input type="text" value="${esc2(it.observacao)}" placeholder="registro para a aba CONCILIAÇÃO" onchange="Recon.editar(${it.id},'observacao',this.value)"></label>
    </div>
    <div class="rc-recalc">Competência origem <b>${esc2(it.compOrigem)}</b> ·
      <b>${s.days}</b> dia${s.days===1?"":"s"} · FTE <b>${fmtFteUI(s.fte)}</b>
      ${it.esperado?` · diferença para o esperado: <b>${fmtFteUI(s.fte-it.esperado.fte)}</b>`:""}</div>
  </div>`;
}

/* Duplicidade não tem período a editar: o que o usuário escolhe é QUAL
   lançamento fica. O app nunca decide isso sozinho. */
function painelRemocao(it,s){
  return `<div class="rc-ajuste">
    <b>Qual lançamento deve permanecer?</b>
    <div class="rc-modos">
      ${it.achados.map(a=>`<label><input type="radio" name="manter${it.id}" value="${a.srcRow}"
        ${a.srcRow===s.manterRow?" checked":""} onchange="Recon.manter(${it.id},${a.srcRow})">
        Linha ${a.srcRow} — ${fmtShort(a.start)} a ${fmtShort(a.end)} · ${a.days} dia${a.days===1?"":"s"} · FTE ${fmtFteUI(a.fte)}</label>`).join("")}
    </div>
    <div class="rc-campos">
      <label class="rc-obs">Observação <input type="text" value="${esc2(it.observacao)}"
        placeholder="registro para a aba CONCILIAÇÃO" onchange="Recon.editar(${it.id},'observacao',this.value)"></label>
    </div>
    <div class="rc-recalc">Serão removidas as linhas <b>${(s.alvoRows||[]).join(", ")||"—"}</b>,
      mantendo a linha <b>${s.manterRow}</b>.</div>
  </div>`;
}
function manter(id, srcRow){
  const it=R.result.items[id];
  if(!it.sugestao||it.sugestao.acao!=="REMOVER") return;
  it.sugestao.manterRow=srcRow;
  it.sugestao.alvoRows=it.achados.map(a=>a.srcRow).filter(r=>r!==srcRow);
  renderResultado();
}

function isoDe(v){ const p=ymdParts(v);
  return p.y+"-"+String(p.m).padStart(2,"0")+"-"+String(p.d).padStart(2,"0"); }
function ymdDeIso(s){ const m=String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m?ymd(+m[1],+m[2],+m[3]):null; }
function fmtFteUI(v){ return (v>0?"+":"")+v.toFixed(4).replace(".",","); }
function descreverAjusteUI(a){
  const sinal=a.kind==="DESCONTAR"?"−":"+";
  return `<span class="rc-sinal ${a.kind==="DESCONTAR"?"neg":"pos"}">${sinal}</span>${fmtShort(a.start)} → ${fmtShort(a.end)}
    <small>${a.days} dia${a.days===1?"":"s"} · ${fmtFteUI(a.fte)}</small>`;
}

/* ================================================================
   5. DECISÕES DO USUÁRIO
   ================================================================ */
function decidir(id, valor){
  const it=R.result.items[id];
  it.decisao=valor;
  it.sugestao = valor==="ACEITAR" ? sugerirCorrecao(it,null) : null;
  if(valor==="ACEITAR") R.aberto=id;     // abre para o usuário conferir antes de aplicar
  renderResultado();
}
function modo(id, m){
  const it=R.result.items[id];
  it.sugestao=sugerirCorrecao(it,m)||it.sugestao;
  renderResultado();
}
function editar(id, campo, valor){
  const it=R.result.items[id], s=it.sugestao;
  if(campo==="observacao"){ it.observacao=valor; return; }
  if(!s) return;
  if(campo==="kind") s.kind=valor;
  else if(campo==="start"){ const v=ymdDeIso(valor); if(v) s.start=v; }
  else if(campo==="end"){ const v=ymdDeIso(valor); if(v) s.end=v; }
  else if(campo==="rateio"){ const n=parseFloat(String(valor).replace(",",".")); if(isFinite(n)) s.rateio=n; }
  if(s.end<s.start) s.end=s.start;
  s.days=diffDaysInclusive(s.start,s.end);
  s.fte=(s.kind==="DESCONTAR"?-1:1)*(s.days/it.compDias)*s.rateio;
  renderResultado();
}
function toggle(id){ R.aberto=R.aberto===id?null:id; renderResultado(); }
function setFiltro(f){ R.filtro=f; renderResultado(); }

/* ================================================================
   6. PRÉVIA E EXPORTAÇÃO
   ================================================================ */
/** "1 linha será incluída" / "3 linhas serão incluídas" */
function frase(n, verboSing, verboPlur, particípio){
  const p=n===1?"":"s";
  return `<b>${n}</b> linha${p} ${n===1?verboSing:verboPlur} ${particípio}${p}`;
}

function blocoPrevia(){
  const plano=planoDeAlteracoes(R.result.items);
  const res=resumoPrevia(plano, R.seguinte?R.seguinte.employees.length:0);
  if(!plano.total){
    return `<section><h2>Fatura conciliada</h2>
      <div class="rc-previa vazia">Nenhuma decisão de alteração foi tomada.
      A fatura conciliada sairia idêntica à original, apenas com a aba de auditoria.
      <div class="rc-acoes"><button class="rc-btn ghost" onclick="Recon.gerar()">Gerar mesmo assim (só auditoria)</button></div></div>
    </section>`;
  }
  const linhas=[];
  plano.incluir.forEach(({item,sug})=>linhas.push(`<li><span class="rc-tag inc">Incluir</span> ${esc2(item.nome)} — ${esc2(descreverProposta(sug))}</li>`));
  plano.substituir.forEach(({item,sug})=>linhas.push(`<li><span class="rc-tag alt">Alterar</span> ${esc2(item.nome)} — ${esc2(descreverProposta(sug))}</li>`));
  plano.remover.forEach(({item,sug})=>linhas.push(`<li><span class="rc-tag rem">Remover</span> ${esc2(item.nome)} — ${esc2(descreverProposta(sug))}</li>`));

  return `<section><h2>Prévia da fatura conciliada</h2>
    <div class="rc-previa">
      <div class="rc-previa-num">
        <span>${frase(res.incluidas,"será","serão","incluída")}</span>
        <span>${frase(res.alteradas,"será","serão","alterada")}</span>
        <span>${frase(res.removidas,"será","serão","removida")}</span>
        <span class="rc-intactas">${frase(res.intactas,"permanecerá","permanecerão","intacta")}</span>
      </div>
      <ul class="rc-previa-lista">${linhas.join("")}</ul>
      <p class="rc-mut">A exportação parte da fatura ${esc2(R.seguinte?R.seguinte.comp.label:"")} e cria um arquivo novo.
        O original não é alterado. Volte e mude qualquer decisão antes de gerar.</p>
      <div class="rc-acoes">
        <button class="rc-btn primaria" onclick="Recon.gerar()">Gerar Fatura Conciliada</button>
      </div>
    </div>
  </section>`;
}

async function gerar(){
  if(!R.seguinte||!R.result) return;
  const base=(R.seguinte.fileName||"").replace(/\.[^.]+$/,"").replace(/[^\wÀ-ÿ]+/g,"_").slice(0,40)||"SMG3";
  try{
    const r=await gerarFaturaConciliada({
      buffer:R.seguinte.buffer, sheetName:R.seguinte.sheetName, map:R.seguinte.map,
      items:R.result.items, comp:R.seguinte.comp, fileNameBase:base
    });
    const el=document.querySelector(".rc-acoes");
    if(el) el.insertAdjacentHTML("beforeend",
      `<span class="rc-ok">✓ ${esc2(r.arquivo)} gerado</span>`);
  }catch(err){
    alert("Falha ao gerar a fatura conciliada: "+err.message);
  }
}

/* ================================================================
   7. ALTERNÂNCIA ENTRE OS MODOS DA ABA
   ================================================================ */
function setModo(m){
  document.querySelectorAll("#rc-modos .rc-modo").forEach(b=>
    b.classList.toggle("active", b.dataset.modo===m));
  $("modo-projetar").classList.toggle("active", m==="projetar");
  $("modo-conciliar").classList.toggle("active", m==="conciliar");
  document.body.dataset.modoConciliacao=m;
}

if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",initDropZones);
else initDropZones();

window.Recon={carregar, conciliar, trocarOrdem, setFiltro, toggle, decidir, modo, editar,
  gerar, setModo, manter, setComp:(qual,y,m)=>{
    const r=R[qual]; if(!r) return;
    setComp(qual, y===null?r.comp.y:+y, m===null?r.comp.m:+m);
  }, assumirSeguinte};

})();
