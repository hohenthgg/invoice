/* Extração · Diarista — filtra as abas de operação do controle de diaristas por
   período e solicitante, trata Groot IDs duplicados e exporta no layout de origem.

   O módulo já nasceu isolado numa IIFE; o que precisou mudar foi o endereçamento:
   os ids dos elementos levam o prefixo `ex-` porque esta aba e a de conciliação
   tinham, as duas, um `#fileInput`. */
(function(){
  "use strict";

  const OPERATIONS = [
    "Pouso Alegre SVC",
    "Pouso Alegre XD",
    "Poços de Caldas",
    "Varginha",
    "Divinópolis",
    "Patos de Minas"
  ];
  const OUT_HEADERS = ["MÊS\nSOLICITAÇÃO","DATA\nSOLICITAÇÃO","SOLICITANTE","EMPRESA\nDIARISTA","GROOT ID","NOME","CARGO","ESCALA"];
  const MESES = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];

  // Estética da planilha original (cabeçalho escuro, texto branco, bordas finas)
  const STY = {
    headerFill: "FF141414",
    headerFont: "FFFFFFFF",
    border: "FF000000"
  };

  const $ = id => document.getElementById(id);
  const drop = $("ex-drop"), fileInput = $("ex-fileInput");
  let workbook = null;
  let parsed = null;   // {op: [{date, id, solic:'id'|'meli'|'', cells:[...]}]}
  let results = null;
  let lastPeriod = null;

  function norm(s){
    return String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ").trim().toLowerCase();
  }
  function dateKey(d){
    return d.getUTCFullYear()+"-"+String(d.getUTCMonth()+1).padStart(2,"0")+"-"+String(d.getUTCDate()).padStart(2,"0");
  }
  function fmtBR(key){ const [y,m,d]=key.split("-"); return d+"/"+m+"/"+y; }
  function mesLabel(key){ const [y,m]=key.split("-"); return MESES[+m-1]+"/"+y.slice(2); }
  function keyToUTCDate(key){ const [y,m,d]=key.split("-").map(Number); return new Date(Date.UTC(y,m-1,d)); }
  function excelSerialToKey(n){ return dateKey(new Date(Math.round((n-25569)*86400*1000))); }
  function solicKind(v){
    const n = norm(v);
    if(!n) return "";
    if(n.includes("meli")) return "meli";
    return "id"; // "ID Logistics" e variações
  }

  // ---------- PASSO 1: leitura ----------
  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("keydown", e => { if(e.key==="Enter"||e.key===" "){e.preventDefault();fileInput.click();} });
  ["dragover","dragenter"].forEach(ev => drop.addEventListener(ev, e => {e.preventDefault();drop.classList.add("over");}));
  ["dragleave","drop"].forEach(ev => drop.addEventListener(ev, e => {e.preventDefault();drop.classList.remove("over");}));
  drop.addEventListener("drop", e => { if(e.dataTransfer.files.length) loadFile(e.dataTransfer.files[0]); });
  fileInput.addEventListener("change", () => { if(fileInput.files.length) loadFile(fileInput.files[0]); });

  function showWarn(el, msg){ el.textContent = msg; el.classList.add("show"); }

  function loadFile(file){
    const warn = $("ex-warn1");
    warn.classList.remove("show");
    const reader = new FileReader();
    reader.onload = e => {
      try{
        workbook = XLSX.read(new Uint8Array(e.target.result), {type:"array", cellDates:true});
        parseWorkbook(file.name);
      }catch(err){
        showWarn(warn, "Não foi possível ler o arquivo. Verifique se é um .xlsx válido. ("+err.message+")");
      }
    };
    reader.onerror = () => showWarn(warn, "Falha ao abrir o arquivo.");
    reader.readAsArrayBuffer(file);
  }

  function findSheet(opName){
    const target = norm(opName);
    return workbook.SheetNames.find(n => norm(n) === target) || null;
  }

  function parseWorkbook(fileName){
    const warn = $("ex-warn1");
    parsed = {};
    const missing = [];
    let globalMin = null, globalMax = null, totalRows = 0;

    for(const op of OPERATIONS){
      const sheetName = findSheet(op);
      if(!sheetName){ missing.push(op); parsed[op] = []; continue; }
      const ws = workbook.Sheets[sheetName];
      const grid = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, blankrows:false});

      let headerIdx = -1, colDate = 1, colId = 4;
      for(let i=0;i<Math.min(grid.length,20);i++){
        const row = grid[i]||[];
        const idCol = row.findIndex(c => norm(c).includes("groot"));
        if(idCol >= 0){
          headerIdx = i; colId = idCol;
          const dCol = row.findIndex(c => norm(c).includes("data"));
          if(dCol >= 0) colDate = dCol;
          break;
        }
      }
      if(headerIdx < 0){ missing.push(op + " (cabeçalho não encontrado)"); parsed[op] = []; continue; }

      const rows = [];
      for(let i=headerIdx+1;i<grid.length;i++){
        const r = grid[i]; if(!r) continue;
        const dv = r[colDate];
        let key = null;
        if(dv instanceof Date && !isNaN(dv)) key = dateKey(dv);
        else if(typeof dv === "number" && dv > 20000 && dv < 80000) key = excelSerialToKey(dv);
        if(!key) continue;

        const idRaw = r[colId];
        const id = (idRaw === null || idRaw === undefined || String(idRaw).trim()==="") ? null : String(idRaw).trim();
        const solicRaw = r[colDate+1];

        rows.push({
          date: key,
          id: id,
          solic: solicKind(solicRaw),
          cells: [
            mesLabel(key),                            // MÊS SOLICITAÇÃO
            key,                                      // DATA (chave; vira Date na exportação)
            solicRaw != null ? solicRaw : "",         // SOLICITANTE
            r[colDate+2] != null ? r[colDate+2] : "", // EMPRESA
            idRaw != null ? idRaw : "",               // GROOT ID
            r[colId+1] != null ? r[colId+1] : "",     // NOME
            r[colId+2] != null ? r[colId+2] : "",     // CARGO
            r[colId+3] != null ? r[colId+3] : ""      // ESCALA
          ]
        });
        if(!globalMin || key < globalMin) globalMin = key;
        if(!globalMax || key > globalMax) globalMax = key;
        totalRows++;
      }
      parsed[op] = rows;
    }

    if(totalRows === 0){
      showWarn(warn, "Nenhuma linha com data encontrada nas abas de operação. Confira se o arquivo segue o padrão SIGO.");
      return;
    }

    $("ex-filechip").textContent = "✔ " + fileName + " — " + totalRows.toLocaleString("pt-BR") + " lançamentos lidos nas abas de operação";
    $("ex-filechip").classList.add("show");
    if(missing.length) showWarn(warn, "Atenção: aba(s) não encontrada(s): " + missing.join(", ") + ". As demais serão processadas normalmente.");
    $("ex-num1").classList.add("done"); $("ex-num1").textContent = "✓";

    $("ex-step2").classList.remove("disabled");
    $("ex-dtIni").value = globalMin;
    $("ex-dtFim").value = globalMax;
    $("ex-dtIni").min = globalMin; $("ex-dtIni").max = globalMax;
    $("ex-dtFim").min = globalMin; $("ex-dtFim").max = globalMax;
    $("ex-rangeNote").textContent = "O arquivo contém lançamentos de " + fmtBR(globalMin) + " até " + fmtBR(globalMax) + ".";

    $("ex-step3").classList.add("disabled");
    $("ex-num2").classList.remove("done"); $("ex-num2").textContent="2";
  }

  // ---------- PASSO 2/3: processamento ----------
  $("ex-btnProcess").addEventListener("click", () => {
    const warn = $("ex-warn2");
    warn.classList.remove("show");
    const ini = $("ex-dtIni").value, fim = $("ex-dtFim").value;
    if(!ini || !fim){ showWarn(warn,"Preencha as duas datas."); return; }
    if(ini > fim){ showWarn(warn,"A data inicial é maior que a final."); return; }
    const mode = document.querySelector('input[name="dedup"]:checked').value;
    const solic = document.querySelector('input[name="solic"]:checked').value;

    results = {};
    lastPeriod = {ini, fim};

    /* Filtra por período e solicitante, mantendo a ordem das operações; a
       deduplicação depois é GLOBAL — ver js/extraction-dedup.js. */
    const entradas = OPERATIONS.map(op => {
      let inPeriod = parsed[op].filter(r => r.date >= ini && r.date <= fim);
      if(solic !== "ambos") inPeriod = inPeriod.filter(r => r.solic === solic);
      return {op, rows: inPeriod};
    });

    const dedup = deduplicarPessoaDia(entradas, mode);
    lastDedup = dedup;
    let tFilt=0, tDup=0, tFinal=0;
    for(const op of OPERATIONS){
      results[op] = dedup.porOperacao[op];
      tFilt += results[op].filtered; tDup += results[op].dups; tFinal += results[op].rows.length;
    }
    renderResults(ini, fim, tFilt, tDup, tFinal, mode === "nao", solic);
  });
  let lastDedup = null;

  function renderResults(ini, fim, tFilt, tDup, tFinal, dedupOff, solic){
    const body = $("ex-resultBody");
    body.innerHTML = "";
    for(const op of OPERATIONS){
      const r = results[op];
      const tr = document.createElement("tr");
      if(r.rows.length === 0) tr.className = "empty";
      const dupCell = dedupOff
        ? '<span class="zero">—</span>'
        : (r.dups ? '<span class="dup">−'+r.dups.toLocaleString("pt-BR")+'</span>' : '<span class="zero">0</span>');
      tr.innerHTML =
        '<td><span class="op">'+op+'</span></td>' +
        '<td class="num">'+r.filtered.toLocaleString("pt-BR")+'</td>' +
        '<td class="num">'+dupCell+'</td>' +
        '<td class="num">'+r.rows.length.toLocaleString("pt-BR")+'</td>' +
        '<td style="text-align:right"></td>';
      const btn = document.createElement("button");
      btn.className = "btn mini amber";
      btn.textContent = "Baixar .xlsx";
      btn.disabled = r.rows.length === 0;
      if(btn.disabled){ btn.style.opacity=".4"; btn.style.cursor="default"; }
      btn.addEventListener("click", () => downloadOne(op));
      tr.lastElementChild.appendChild(btn);
      body.appendChild(tr);
    }
    const solicTxt = solic === "ambos" ? "ID + MELI" : (solic === "id" ? "só ID Logistics" : "só MELI");
    $("ex-periodLabel").textContent = fmtBR(ini) + " → " + fmtBR(fim) + " · " + solicTxt;
    const rs = lastDedup ? lastDedup.resumo : null;
    $("ex-totalBar").innerHTML =
      "Registros encontrados: <b>"+tFilt.toLocaleString("pt-BR")+"</b>&nbsp;·&nbsp;" +
      "Únicos pessoa-dia: <b>"+tFinal.toLocaleString("pt-BR")+"</b>&nbsp;·&nbsp;" +
      "Duplicados removidos: <b>"+(dedupOff ? "—" : tDup.toLocaleString("pt-BR"))+"</b>" +
      (rs && rs.semGroot ? "&nbsp;·&nbsp;GROOT ausente: <b>"+rs.semGroot.toLocaleString("pt-BR")+"</b>" : "");
    renderDetalheDedup(dedupOff);
    $("ex-num2").classList.add("done"); $("ex-num2").textContent="✓";
    $("ex-step3").classList.remove("disabled");
    $("ex-step3").scrollIntoView({behavior:"smooth",block:"nearest"});
  }

  /* Nada é removido em silêncio: quem foi descartado, de onde veio e onde a
     ocorrência mantida ficou. Painel recolhido para não poluir a tela. */
  function renderDetalheDedup(dedupOff){
    const alvo = $("ex-detalheDedup");
    if(!alvo) return;
    if(dedupOff || !lastDedup){ alvo.innerHTML=""; alvo.classList.add("hidden"); return; }
    const {descartados, semGroot, resumo} = lastDedup;
    if(!descartados.length && !semGroot.length){ alvo.innerHTML=""; alvo.classList.add("hidden"); return; }
    alvo.classList.remove("hidden");
    const regra = resumo.modo==="dia"
      ? "Duplicado = mesmo GROOT normalizado + mesma data, valendo entre operações, abas e arquivos."
      : "Duplicado = mesmo GROOT normalizado no período inteiro, valendo entre operações, abas e arquivos.";
    let h = '<details class="ex-dedup"><summary>O que foi removido ('
      + descartados.length + ' duplicado' + (descartados.length===1?'':'s')
      + (semGroot.length? ' · '+semGroot.length+' sem GROOT':'') + ')</summary>'
      + '<p class="regra">'+regra+' Fica sempre a <b>primeira ocorrência encontrada</b>.</p>';
    if(descartados.length){
      h += '<table><thead><tr><th>GROOT</th><th>Data</th><th>Operação mantida</th><th>Operação descartada</th></tr></thead><tbody>'
        + descartados.slice(0,200).map(d=>'<tr><td>'+d.groot+'</td><td>'+d.data
          +'</td><td>'+d.mantidaEm+'</td><td>'+d.descartadaDe+'</td></tr>').join("")
        + '</tbody></table>'
        + (descartados.length>200?'<p class="regra">…e mais '+(descartados.length-200)+'.</p>':'');
    }
    if(semGroot.length){
      h += '<p class="regra alerta"><b>GROOT ausente — revisar:</b> '+semGroot.length
        + ' registro'+(semGroot.length===1?'':'s')+' sem identificador '
        + (semGroot.length===1?'foi mantido':'foram mantidos')
        + ' e não '+(semGroot.length===1?'entrou':'entraram')+' na deduplicação. Sem GROOT não há '
        + 'pessoa-dia, e tratar todos os vazios como a mesma pessoa apagaria gente diferente.</p>';
    }
    alvo.innerHTML = h + '</details>';
  }

  // ---------- exportação com a estética do SIGO ----------
  function thinBorder(){
    const s = {style:"thin", color:{argb:STY.border}};
    return {top:s, left:s, bottom:s, right:s};
  }

  function addStyledSheet(wbx, op){
    const ws = wbx.addWorksheet(op.substring(0,31), {
      views: [{state:"frozen", ySplit:1}]
    });
    ws.columns = [
      {width:15},{width:15},{width:14},{width:13},{width:12},{width:38},{width:12},{width:9}
    ];

    // Cabeçalho: fundo escuro, texto branco, negrito, centralizado (padrão SIGO)
    const head = ws.addRow(OUT_HEADERS);
    head.height = 32;
    head.eachCell(cell => {
      cell.fill = {type:"pattern", pattern:"solid", fgColor:{argb:STY.headerFill}};
      cell.font = {bold:true, color:{argb:STY.headerFont}, size:11, name:"Calibri"};
      cell.alignment = {horizontal:"center", vertical:"middle", wrapText:true};
      cell.border = thinBorder();
    });

    for(const r of results[op].rows){
      const c = r.cells;
      const idNum = (c[4] !== "" && !isNaN(Number(c[4]))) ? Number(c[4]) : c[4];
      const row = ws.addRow([ c[0], keyToUTCDate(c[1]), c[2], c[3], idNum, c[5], c[6], c[7] ]);
      row.eachCell({includeEmpty:true}, (cell, col) => {
        cell.border = thinBorder();
        cell.font = {size:11, name:"Calibri"};
        if(col === 1) cell.alignment = {horizontal:"right", vertical:"middle"};
        else if(col === 2){ cell.numFmt = "dd/mm/yyyy"; cell.alignment = {horizontal:"right", vertical:"middle"}; }
        else if(col === 5) cell.alignment = {horizontal:"right", vertical:"middle"};
        else if(col === 7 || col === 8) cell.alignment = {horizontal:"center", vertical:"middle"};
        else cell.alignment = {horizontal:"left", vertical:"middle"};
      });
    }

    // Filtros no cabeçalho, como na planilha original
    ws.autoFilter = {from:{row:1, column:1}, to:{row:1, column:8}};
    return ws;
  }

  // Padrão de nome: "Filial - Diaristas - Mês.26"
  function mesArquivo(key){
    const [y,m] = key.split("-");
    const nome = MESES[+m-1];
    return nome.charAt(0).toUpperCase() + nome.slice(1) + "." + y.slice(2);
  }
  function periodoNome(){
    const a = mesArquivo(lastPeriod.ini), b = mesArquivo(lastPeriod.fim);
    return a === b ? a : a + " a " + b; // período dentro do mesmo mês: "Agosto.26"; cruzando meses: "Julho.26 a Agosto.26"
  }
  function fileName(filial){
    return filial + " - Diaristas - " + periodoNome() + ".xlsx";
  }
  async function saveWorkbook(wbx, fileName){
    const buf = await wbx.xlsx.writeBuffer();
    const blob = new Blob([buf], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  async function downloadOne(op){
    const wbx = new ExcelJS.Workbook();
    addStyledSheet(wbx, op);
    await saveWorkbook(wbx, fileName(op));
  }
  $("ex-btnAll").addEventListener("click", async () => {
    for(const op of OPERATIONS){
      if(results[op].rows.length === 0) continue;
      await downloadOne(op);
      await new Promise(res => setTimeout(res, 400)); // evita bloqueio de múltiplos downloads
    }
  });
  $("ex-btnCombined").addEventListener("click", async () => {
    const wbx = new ExcelJS.Workbook();
    for(const op of OPERATIONS) addStyledSheet(wbx, op);
    await saveWorkbook(wbx, fileName("6 Operações"));
  });

})();
