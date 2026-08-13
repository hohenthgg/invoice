/* Ajustes MELI — renderização */
"use strict";

/* ================================================================
   RENDERIZAÇÃO
   ================================================================ */
const esc=s=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/"/g,"&quot;");

function runAnalysis(){
  const {adjustments,errors}=analyze(state.employees,state.comp);
  state.adjustments=adjustments;
  state.errors=errors;
  render();
}

function compEditor(){
  let opts="";
  MONTHS.forEach((m,i)=>opts+=`<option value="${i+1}" ${i+1===state.comp.m?"selected":""}>${m}</option>`);
  let yopts="";
  for(let y=state.comp.y-2;y<=state.comp.y+2;y++) yopts+=`<option value="${y}" ${y===state.comp.y?"selected":""}>${y}</option>`;
  return `<div class="comp-edit" id="compEdit">
    <select id="selMonth">${opts}</select><select id="selYear">${yopts}</select>
    <button id="btnApplyComp">Aplicar</button>
  </div>`;
}

function render(){
  const el=document.getElementById("result");
  const a=state.adjustments, c=state.comp;
  const nd=a.filter(r=>r.kind==="DESCONTAR").length;
  const na=a.length-nd;
  let html=`<div class="comp" id="compLine" title="Clique para alterar a competência">${c.label}
      <span class="arrow">→ ${c.nextLabel}</span>
      ${state.confident?"":'<span class="alter" id="btnAlter">Alterar</span>'}
    </div>`
    + compEditor()
    + `<div class="count"><b>${a.length} ajuste${a.length===1?"":"s"} encontrado${a.length===1?"":"s"}</b>`
    + (a.length? ` · ${nd} desconto${nd===1?"":"s"} · ${na} acréscimo${na===1?"":"s"}`:"") + `</div>`;

  if(!a.length){
    html+=`<div class="hint">Nenhuma movimentação após o corte de ${fmtShort(c.snapshot)}. Nada a ajustar.</div>`;
  } else {
    html+=`<table><thead><tr>
      <th>Ajuste</th><th>Colaborador</th><th>Movimentação</th><th>Período</th><th class="num">Dias</th><th></th>
    </tr></thead><tbody>`;
    a.forEach(r=>{
      const cls=r.kind==="DESCONTAR"?"desc":"acr";
      const warn=r.warnings.length?`<span class="warn" title="${esc(r.warnings.join(" • "))}">⚠</span>`:"";
      html+=`<tr class="adj" data-id="${r.id}">
        <td><span class="tag ${cls}">${r.kind} ${r.days} DIA${r.days===1?"":"S"}</span>${warn}</td>
        <td>${esc(r.emp.nome)}</td>
        <td>${esc(r.mov)}</td>
        <td class="period">${fmtShort(r.start)} → ${fmtShort(r.end)}</td>
        <td class="num">${r.days}</td>
        <td class="chk"><input type="checkbox" class="inc" data-id="${r.id}" ${r.include?"checked":""}></td>
      </tr>`;
    });
    html+=`</tbody></table>`;
  }
  if(state.errors.length){
    html+=`<div class="errors" id="errBox">
      ${state.errors.length} registro${state.errors.length===1?" não pôde":"s não puderam"} ser analisado${state.errors.length===1?"":"s"}
      · <span class="toggle" id="errToggle">Ver</span>
      <ul>${state.errors.map(e=>`<li><b>${esc(e.nome)}</b> — ${esc(e.reason)}</li>`).join("")}</ul>
    </div>`;
  }
  el.innerHTML=html;
  document.getElementById("footer").classList.toggle("on",a.length>0);
  updateFooter();
  bindResultEvents();
}

function updateFooter(){
  const total=state.adjustments.length;
  const n=state.adjustments.filter(r=>r.include).length;
  document.getElementById("selCount").textContent=n+" de "+total+" ajuste"+(total===1?"":"s")+" selecionado"+(n===1?"":"s");
  document.getElementById("btnExport").disabled=n===0;
}

function toggleDetail(id){
  const existing=document.querySelector("tr.detail");
  const wasFor=existing?+existing.dataset.for:null;
  if(existing) existing.remove();
  if(wasFor===id) return;
  const row=document.querySelector(`tr.adj[data-id="${id}"]`);
  const r=state.adjustments[id];
  const tr=document.createElement("tr");
  tr.className="detail"; tr.dataset.for=id;
  tr.innerHTML=`<td colspan="6">
    <div class="grid">
      <span><b>GROOT</b> <span class="mono">${esc(r.emp.groot||"—")}</span></span>
      <span><b>Matrícula</b> <span class="mono">${esc(r.emp.matricula||"—")}</span></span>
      <span><b>Rateio</b> ${(r.rateio*100).toLocaleString("pt-BR",{maximumFractionDigits:1})}%</span>
      <span><b>FTE ajuste</b> ${(r.fte>0?"+":"")+r.fte.toFixed(4).replace(".",",")}</span>
    </div>
    ${esc(r.sentence)}${r.warnings.length?"<br>⚠ "+esc(r.warnings.join(" • ")):""}
  </td>`;
  row.after(tr);
}

function bindResultEvents(){
  const alterBtn=document.getElementById("btnAlter");
  const openEdit=()=>document.getElementById("compEdit").classList.add("on");
  if(alterBtn) alterBtn.onclick=e=>{e.stopPropagation();openEdit();};
  document.getElementById("compLine").onclick=openEdit;
  const apply=document.getElementById("btnApplyComp");
  if(apply) apply.onclick=()=>{
    state.comp=buildCompetence(+document.getElementById("selYear").value,+document.getElementById("selMonth").value);
    state.confident=true;
    runAnalysis();
  };
  const errT=document.getElementById("errToggle");
  if(errT) errT.onclick=()=>{
    const box=document.getElementById("errBox");
    box.classList.toggle("open");
    errT.textContent=box.classList.contains("open")?"Ocultar":"Ver";
  };
}
