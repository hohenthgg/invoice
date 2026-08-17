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
  // conteúdo vindo da planilha vai para o HTML: escapar é obrigatório
  const esc = v => String(v==null?"":v).replace(/[&<>"]/g, c =>
    ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]));
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
  // chave inesperada não deve virar "undefined/undefined/" na tela
  function dataBR(key){ return /^\d{4}-\d{2}-\d{2}$/.test(String(key||"")) ? fmtBR(key) : (key||"—"); }
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
        const txt = v => (v === null || v === undefined) ? "" : String(v).trim();

        rows.push({
          date: key,
          id: id,
          solic: solicKind(solicRaw),
          /* Campos nomeados além de `cells`: quando um registro precisa ser
             mostrado ao usuário (ex.: sem GROOT, para revisão manual), quem
             o exibe não deveria ter de conhecer a posição das colunas. */
          nome:        txt(r[colId+1]),
          cargo:       txt(r[colId+2]),
          escala:      txt(r[colId+3]),
          empresa:     txt(r[colDate+2]),
          solicitante: txt(solicRaw),
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
    /* A auditoria olha as linhas ANTES da dedup: o que interessa é justamente
       o que ela uniria (mesmo GROOT com nomes diferentes) ou deixaria de unir
       (mesmo nome com GROOTs diferentes). */
    lastAudit = auditarIdentidade(entradas);
    lastSolic = solic;
    let tFilt=0, tDup=0, tFinal=0;
    for(const op of OPERATIONS){
      results[op] = dedup.porOperacao[op];
      tFilt += results[op].filtered; tDup += results[op].dups; tFinal += results[op].rows.length;
    }
    renderResults(ini, fim, tFilt, tDup, tFinal, mode === "nao", solic);
  });
  let lastDedup = null, lastAudit = null, lastSolic = "ambos";

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
    const ra = lastAudit ? lastAudit.resumo : null;
    const conflitos = ra ? ra.mesmoNome + ra.mesmoGroot : 0;
    $("ex-totalBar").innerHTML =
      "Registros encontrados: <b>"+tFilt.toLocaleString("pt-BR")+"</b>&nbsp;·&nbsp;" +
      "Únicos pessoa-dia: <b>"+tFinal.toLocaleString("pt-BR")+"</b>&nbsp;·&nbsp;" +
      "Duplicados removidos: <b>"+(dedupOff ? "—" : tDup.toLocaleString("pt-BR"))+"</b>" +
      (ra && ra.semGroot ? "&nbsp;·&nbsp;GROOT ausente: <b>"+ra.semGroot.toLocaleString("pt-BR")+"</b>" : "") +
      (conflitos ? "&nbsp;·&nbsp;Conflitos de identidade: <b>"+conflitos.toLocaleString("pt-BR")+"</b>" : "");
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
    if(!lastDedup){ alvo.innerHTML=""; alvo.classList.add("hidden"); return; }
    const {descartados, resumo} = lastDedup;
    const aud = lastAudit || {semGroot:[], mesmoNome:[], mesmoGroot:[], resumo:{}};
    /* A lista de "sem GROOT" vem da auditoria, não da dedup: com a dedup
       desligada a dedup não olha identificador nenhum, mas os registros sem
       identificador continuam lá e continuam precisando de revisão. */
    const semGroot = aud.semGroot;
    const temAnomalia = aud.mesmoNome.length || aud.mesmoGroot.length;
    /* Com a dedup desligada não há descartados, mas os conflitos de identidade
       continuam existindo — o arquivo é o mesmo. */
    if(!descartados.length && !semGroot.length && !temAnomalia){
      alvo.innerHTML=""; alvo.classList.add("hidden"); return;
    }
    alvo.classList.remove("hidden");
    const regra = resumo.modo==="dia"
      ? "Duplicado = mesmo GROOT normalizado + mesma data, valendo entre operações, abas e arquivos."
      : "Duplicado = mesmo GROOT normalizado no período inteiro, valendo entre operações, abas e arquivos.";
    const partes = [];
    if(!dedupOff) partes.push(descartados.length + ' duplicado' + (descartados.length===1?'':'s'));
    if(semGroot.length) partes.push(semGroot.length + ' sem GROOT');
    if(temAnomalia) partes.push((aud.mesmoNome.length + aud.mesmoGroot.length) + ' conflito'
      + (aud.mesmoNome.length + aud.mesmoGroot.length===1?'':'s') + ' de identidade');
    let h = '<details class="ex-dedup"><summary>Deduplicação e revisão de identidade ('
      + partes.join(' · ') + ')</summary>'
      + '<p class="regra">' + (dedupOff
          ? 'Deduplicação desligada — nada foi removido. Os conflitos de identidade abaixo continuam valendo: eles são do arquivo, não da remoção.'
          : regra + ' Fica sempre a <b>primeira ocorrência encontrada</b>.')
      + ' <button type="button" class="btn mini amber" id="ex-btnRelatorio">Baixar relatório de revisão (.xlsx)</button></p>';
    if(descartados.length && !dedupOff){
      h += '<table><thead><tr><th>GROOT</th><th>Nome</th><th>Data</th><th>Operação mantida</th><th>Operação descartada</th></tr></thead><tbody>'
        + descartados.slice(0,200).map(d=>'<tr><td>'+esc(d.groot)+'</td><td>'+(esc(d.nome)||'<i>—</i>')
          +'</td><td>'+esc(dataBR(d.data))+'</td><td>'+esc(d.mantidaEm)+'</td><td>'+esc(d.descartadaDe)+'</td></tr>').join("")
        + '</tbody></table>'
        + (descartados.length>200?'<p class="regra">…e mais '+(descartados.length-200)+'.</p>':'');
    }
    if(semGroot.length) h += blocoSemGroot(semGroot);
    h += blocoAnomalias(aud);
    alvo.innerHTML = h + '</details>';
    const btn = $("ex-btnRelatorio");
    if(btn) btn.addEventListener("click", baixarRelatorio);
  }

  /* Sem GROOT não há pessoa-dia — mas o registro existe e alguém trabalhou.
     Um número solto não é revisável: para conferir é preciso saber QUEM é a
     pessoa e em qual filial ela está, para então buscar o identificador na
     origem. Daí a lista nominal, recolhida por padrão. */
  const LIMITE_SEM_GROOT = 300;
  function blocoSemGroot(semGroot){
    const n = semGroot.length, um = n===1;
    let h = '<p class="regra alerta"><b>GROOT ausente — revisar:</b> ' + n
      + ' registro'+(um?'':'s')+' sem identificador ' + (um?'foi mantido':'foram mantidos')
      + ' e não '+(um?'entrou':'entraram')+' na deduplicação. Sem GROOT não há '
      + 'pessoa-dia, e tratar todos os vazios como a mesma pessoa apagaria gente diferente. '
      + 'Se algum desses nomes se repetir no mesmo dia, a duplicidade só pode ser '
      + 'resolvida preenchendo o GROOT na origem.</p>';

    // por filial primeiro: normalmente o problema está concentrado em uma operação
    const porOp = new Map();
    semGroot.forEach(s => porOp.set(s.op, (porOp.get(s.op)||0)+1));
    h += '<p class="regra">Por filial: '
      + Array.from(porOp).map(([op,c]) => '<b>'+esc(op)+'</b> '+c).join(' · ') + '</p>';

    h += '<details class="ex-semgroot"><summary>Ver ' + (um?'o registro':'os '+n+' registros')
      + ' sem identificador</summary>'
      + '<table><thead><tr><th>Filial / operação</th><th>Data</th><th>Nome</th>'
      + '<th>Empresa</th><th>Cargo</th><th>Escala</th><th>Solicitante</th></tr></thead><tbody>'
      + semGroot.slice(0, LIMITE_SEM_GROOT).map(s => {
          const cel = v => '<td>' + (esc(v) || '<i>—</i>') + '</td>';
          return '<tr><td>'+esc(s.op)+'</td><td>'+esc(dataBR(s.date))+'</td>'
            + cel(s.nome) + cel(s.empresa) + cel(s.cargo) + cel(s.escala) + cel(s.solicitante)
            + '</tr>';
        }).join("")
      + '</tbody></table>'
      + (n>LIMITE_SEM_GROOT
          ? '<p class="regra">…e mais '+(n-LIMITE_SEM_GROOT)+'. A lista completa está no relatório .xlsx.</p>'
          : '')
      + '</details>';
    return h;
  }

  /* Conflitos de identidade: o identificador e o nome contando histórias
     diferentes. O app só aponta — sem outra fonte, o dado não diz qual das
     versões está certa, e unir por conta própria apagaria gente. */
  const LIMITE_CASOS = 50, LIMITE_OCORR = 8;
  function blocoAnomalias(aud){
    let h = "";
    if(aud.mesmoNome.length){
      h += secaoConflito({
        classe:"ex-mesmonome", casos:aud.mesmoNome,
        titulo:"Mesmo nome, identificadores diferentes",
        intro:"A deduplicação <b>não uniu</b> estes registros — e não deveria: nome não é chave, "
             +"homônimo existe. Mas se for a mesma pessoa cadastrada duas vezes, ela está sendo "
             +"contada em dobro.",
        colVariante:"GROOT"
      });
    }
    if(aud.mesmoGroot.length){
      h += secaoConflito({
        classe:"ex-mesmogroot", casos:aud.mesmoGroot,
        titulo:"Mesmo identificador, nomes diferentes",
        intro:"A deduplicação <b>uniu</b> estes registros, por serem o mesmo GROOT. Se os nomes "
             +"forem de pessoas diferentes, uma diária pode ter sido descartada como se fosse "
             +"repetida — este é o caso mais grave da lista.",
        colVariante:"Nome"
      });
    }
    return h;
  }

  function secaoConflito({classe, casos, titulo, intro, colVariante}){
    const mostrados = casos.slice(0, LIMITE_CASOS);
    let h = '<details class="ex-conflito '+classe+'"><summary>'+esc(titulo)+' — '
      + casos.length + ' caso' + (casos.length===1?'':'s') + '</summary>'
      + '<p class="regra">'+intro+'</p>';
    h += mostrados.map(c => {
      const cabecalho = classe==="ex-mesmonome"
        ? esc(c.rotulo) + ' <span class="ex-tag">'+c.variantes.length+' identificadores</span>'
        : 'GROOT ' + esc(c.rotulo) + ' <span class="ex-tag">'+c.variantes.length+' nomes</span>';
      const linhas = c.variantes.map(v => {
        const vis = v.ocorrencias.slice(0, LIMITE_OCORR);
        const cel = x => '<td>' + (esc(x) || '<i>—</i>') + '</td>';
        return vis.map((o,i) =>
            '<tr>' + (i===0
              ? '<td class="var" rowspan="'+vis.length+'">'+(esc(v.rotulo)||'<i>—</i>')
                + '<span class="vezes">'+v.vezes+'×</span></td>'
              : '')
            + '<td>'+esc(o.op)+'</td><td>'+esc(dataBR(o.date))+'</td>'
            + cel(o.empresa) + cel(o.cargo) + cel(o.escala) + cel(o.solicitante) + '</tr>').join("")
          + (v.ocorrencias.length>LIMITE_OCORR
              ? '<tr><td></td><td colspan="6" class="mais">…e mais '
                + (v.ocorrencias.length-LIMITE_OCORR) + ' ocorrência'
                + (v.ocorrencias.length-LIMITE_OCORR===1?'':'s') + ' — a lista completa está no .xlsx.</td></tr>'
              : '');
      }).join("");
      return '<div class="ex-caso"><div class="ex-caso-ttl">'+cabecalho+'</div>'
        + '<p class="ex-caso-why">'+esc(c.explicacao)+'</p>'
        + '<table><thead><tr><th>'+esc(colVariante)+'</th><th>Filial / operação</th><th>Data</th>'
        + '<th>Empresa</th><th>Cargo</th><th>Escala</th><th>Solicitante</th></tr></thead>'
        + '<tbody>'+linhas+'</tbody></table></div>';
    }).join("");
    if(casos.length > LIMITE_CASOS)
      h += '<p class="regra">…e mais '+(casos.length-LIMITE_CASOS)
         + ' caso'+(casos.length-LIMITE_CASOS===1?'':'s')+'. O relatório .xlsx traz todos.</p>';
    return h + '</details>';
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
  /* ---------- relatório de revisão ----------
     O painel na tela serve para decidir; o .xlsx serve para trabalhar: filtrar,
     ordenar, marcar o que já foi corrigido e devolver para quem cuida do
     cadastro. Uma aba por tipo de problema, sempre com filial, data e os
     campos que permitem achar a pessoa na origem. */
  const COLS_CONTEXTO = ["FILIAL / OPERAÇÃO","DATA","EMPRESA","CARGO","ESCALA","SOLICITANTE"];
  const W_CONTEXTO = [20,12,16,18,10,14];

  function addRelSheet(wbx, nome, headers, widths, linhas){
    const ws = wbx.addWorksheet(nome.substring(0,31), {views:[{state:"frozen", ySplit:1}]});
    ws.columns = widths.map(w => ({width:w}));
    const head = ws.addRow(headers);
    head.height = 28;
    head.eachCell(cell => {
      cell.fill = {type:"pattern", pattern:"solid", fgColor:{argb:STY.headerFill}};
      cell.font = {bold:true, color:{argb:STY.headerFont}, size:11, name:"Calibri"};
      cell.alignment = {horizontal:"center", vertical:"middle", wrapText:true};
      cell.border = thinBorder();
    });
    linhas.forEach(vals => {
      const row = ws.addRow(vals);
      row.eachCell({includeEmpty:true}, (cell, col) => {
        cell.border = thinBorder();
        cell.font = {size:11, name:"Calibri"};
        if(cell.value instanceof Date) cell.numFmt = "dd/mm/yyyy";
        cell.alignment = {horizontal: (typeof cell.value === "number" || cell.value instanceof Date)
          ? "right" : "left", vertical:"middle", wrapText:false};
      });
    });
    if(linhas.length) ws.autoFilter = {from:{row:1, column:1}, to:{row:1, column:headers.length}};
    return ws;
  }

  // data vira Date de verdade para o Excel filtrar e ordenar; chave ilegível fica como texto
  const dataCell = k => /^\d{4}-\d{2}-\d{2}$/.test(String(k||"")) ? keyToUTCDate(k) : (k||"");
  const ctxVals = o => [o.op, dataCell(o.date), o.empresa, o.cargo, o.escala, o.solicitante];

  async function baixarRelatorio(){
    const aud = lastAudit, ded = lastDedup;
    if(!aud || !ded) return;
    const wbx = new ExcelJS.Workbook();
    const solicTxt = lastSolic === "ambos" ? "ID Logistics + MELI"
                   : (lastSolic === "id" ? "só ID Logistics" : "só MELI");
    const modoTxt = ded.resumo.modo === "dia" ? "uma linha por pessoa por dia"
                  : (ded.resumo.modo === "periodo" ? "uma linha por pessoa no período" : "sem remoção");

    // ---- Resumo: o que cada aba quer dizer, para o relatório se explicar sozinho
    addRelSheet(wbx, "Resumo", ["ITEM","VALOR","O QUE SIGNIFICA"], [34,14,96], [
      ["Período", fmtBR(lastPeriod.ini)+" a "+fmtBR(lastPeriod.fim), "Recorte por data de solicitação."],
      ["Solicitante", solicTxt, "Filtro aplicado na extração."],
      ["Deduplicação", modoTxt, "Duplicado = mesmo GROOT normalizado + mesma data, valendo entre operações, abas e arquivos."],
      ["Registros no período", aud.resumo.registros, "Antes da deduplicação."],
      ["Duplicados removidos", ded.resumo.duplicados, "Ficou sempre a primeira ocorrência encontrada. Aba 'Duplicados removidos'."],
      ["Sem identificador", aud.resumo.semGroot, "Mantidos e fora da deduplicação — sem GROOT não existe pessoa-dia. Aba 'Sem identificador'."],
      ["Mesmo nome, GROOTs difer.", aud.resumo.mesmoNome, "A dedup NÃO uniu. Se for a mesma pessoa cadastrada duas vezes, está contada em dobro."],
      ["…destes, ids realmente distintos", aud.resumo.mesmoNomeIdsDistintos, "Os demais diferem só por zeros à esquerda (formatação)."],
      ["Mesmo GROOT, nomes difer.", aud.resumo.mesmoGroot, "A dedup UNIU. Confira antes de aceitar o resultado."],
      ["…destes, nomes realmente distintos", aud.resumo.mesmoGrootNomesDistintos, "O caso mais grave: uma diária pode ter sido descartada como repetida."],
      ["", "", ""],
      ["Nada foi corrigido automaticamente", "—", "O aplicativo aponta; a correção é no cadastro de origem. Sem outra fonte, o dado não diz qual das versões está certa."]
    ]);

    // ---- Sem identificador
    if(aud.semGroot.length){
      addRelSheet(wbx, "Sem identificador",
        ["NOME"].concat(COLS_CONTEXTO).concat(["GROOT (COMO ESTAVA)"]),
        [38].concat(W_CONTEXTO).concat([20]),
        aud.semGroot.map(o => [o.nome].concat(ctxVals(o)).concat([o.grootBruto])));
    }

    // ---- Mesmo nome, GROOTs diferentes
    if(aud.mesmoNome.length){
      const linhas = [];
      aud.mesmoNome.forEach((c, i) => c.variantes.forEach(v => v.ocorrencias.forEach(o =>
        linhas.push([i+1, c.rotulo, v.rotulo, v.vezes, c.explicacao].concat(ctxVals(o))))));
      addRelSheet(wbx, "Mesmo nome · GROOT difere",
        ["CASO","NOME","GROOT","VEZES","DIAGNÓSTICO"].concat(COLS_CONTEXTO),
        [7,38,16,8,80].concat(W_CONTEXTO), linhas);
    }

    // ---- Mesmo GROOT, nomes diferentes
    if(aud.mesmoGroot.length){
      const linhas = [];
      aud.mesmoGroot.forEach((c, i) => c.variantes.forEach(v => v.ocorrencias.forEach(o =>
        linhas.push([i+1, c.rotulo, v.rotulo, v.vezes, c.explicacao].concat(ctxVals(o))))));
      addRelSheet(wbx, "Mesmo GROOT · nome difere",
        ["CASO","GROOT","NOME","VEZES","DIAGNÓSTICO"].concat(COLS_CONTEXTO),
        [7,16,38,8,80].concat(W_CONTEXTO), linhas);
    }

    // ---- Duplicados removidos
    if(ded.descartados.length){
      addRelSheet(wbx, "Duplicados removidos",
        ["GROOT","NOME","DATA","OPERAÇÃO MANTIDA","OPERAÇÃO DESCARTADA"],
        [16,38,12,22,22],
        ded.descartados.map(d => [d.groot, d.nome,
          dataCell(d.data), d.mantidaEm, d.descartadaDe]));
    }

    await saveWorkbook(wbx, "Diaristas - Revisão de identidade - " + periodoNome() + ".xlsx");
  }

  $("ex-btnCombined").addEventListener("click", async () => {
    const wbx = new ExcelJS.Workbook();
    for(const op of OPERATIONS) addStyledSheet(wbx, op);
    await saveWorkbook(wbx, fileName("6 Operações"));
  });

})();
