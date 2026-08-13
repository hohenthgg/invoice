/* Ajustes MELI — inicialização e eventos */
"use strict";

/* ================================================================
   EVENTOS GLOBAIS
   ================================================================ */
function init(){
  document.getElementById("btnImport").onclick=()=>document.getElementById("fileInput").click();
  document.getElementById("fileInput").onchange=ev=>{
    const f=ev.target.files[0]; if(!f) return;
    const reader=new FileReader();
    reader.onload=e=>{ try{ if(importWorkbook(new Uint8Array(e.target.result),f.name)) runAnalysis(); }
      catch(err){ alert("Falha ao ler o arquivo: "+err.message); } };
    reader.readAsArrayBuffer(f);
    ev.target.value="";
  };
  document.getElementById("btnExport").onclick=exportAdjustments;
  document.getElementById("result").addEventListener("click",ev=>{
    const chk=ev.target.closest("input.inc");
    if(chk){ state.adjustments[+chk.dataset.id].include=chk.checked; updateFooter(); ev.stopPropagation(); return; }
    const tr=ev.target.closest("tr.adj");
    if(tr) toggleDetail(+tr.dataset.id);
  });
}
if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",init);
else init();
