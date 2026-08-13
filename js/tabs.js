/* Ajustes MELI — navegação entre as abas principais */
"use strict";

/* ================================================================
   ABAS PRINCIPAIS

   Cada aba é um <div class="apppanel" id="panel-<nome>">. O nome da
   aba ativa também vai para o data-tab do <body>, porque o rodapé de
   seleção é fixo e pertence só à Conciliação — o CSS o esconde nas
   demais sem que a aba 1 precise saber que existem outras.
   ================================================================ */
function initTabs(){
  const tabs=[...document.querySelectorAll("#appTabs .apptab")];
  if(!tabs.length) return;

  function activate(name){
    tabs.forEach(t=>t.classList.toggle("active",t.dataset.panel===name));
    document.querySelectorAll(".apppanel").forEach(p=>
      p.classList.toggle("active",p.id==="panel-"+name));
    document.body.dataset.tab=name;
    if(location.hash.slice(1)!==name) history.replaceState(null,"","#"+name);
  }

  tabs.forEach(t=>t.onclick=()=>activate(t.dataset.panel));

  // permite abrir direto numa aba: index.html#fusao
  const fromHash=location.hash.slice(1);
  if(fromHash && document.getElementById("panel-"+fromHash)) activate(fromHash);
}
if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",initTabs);
else initTabs();
