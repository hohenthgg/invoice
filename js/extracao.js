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
  /* ================================================================
     AS COLUNAS, LOCALIZADAS PELO PRÓPRIO CABEÇALHO
     ----------------------------------------------------------------
     A leitura antiga achava só o GROOT e contava posições a partir dele
     (nome = groot+1, cargo = groot+2…). Em Varginha isso devolveu 119
     linhas com GROOT, NOME, CARGO e ESCALA TODOS vazios: basta a coluna
     estar em outro lugar para tudo escorregar junto, em silêncio.

     Agora cada coluna procura o próprio nome. `nomes` são grafias
     aceitas, em ordem de preferência; a resolução é feita das mais
     específicas para as mais genéricas, para "ESCALA NATURAL" não ser
     abocanhada por "ESCALA".

     `horario` é lido à parte porque Divinópolis já traz o horário na
     origem, sob "ESCALA NATURAL". Quando o arquivo tem, vale o arquivo;
     quando não tem, vale o padrão da operação. */
  const COLUNAS_SIGO = [
    {k:"horario",     nomes:["escala horario","escala natural","horario","horario escala","turno"]},
    {k:"mes",         nomes:["mes solicitacao","mes da solicitacao","mes"]},
    {k:"data",        nomes:["data solicitacao","data da solicitacao","data"]},
    {k:"solicitante", nomes:["solicitante","solicitado por"]},
    {k:"empresa",     nomes:["empresa diarista","empresa","fornecedor"]},
    {k:"groot",       nomes:["groot id","id groot","groot","id"]},
    {k:"nome",        nomes:["nome","nome completo","colaborador"]},
    {k:"cargo",       nomes:["cargo","funcao","função"]},
    {k:"escala",      nomes:["escala"]}
  ];

  /* ================================================================
     O MODELO DA PLANILHA ENTREGUE
     ----------------------------------------------------------------
     Copiado de "Modelo diaristas.xlsx": oito colunas, e a ESCALA é o
     HORÁRIO (`03:00 07:00 08:00 11:20`), não o 6x1 do SIGO. Larguras,
     fontes, tons de cinza do cabeçalho e alinhamentos vêm de lá, célula
     a célula — a ideia é que o arquivo gerado passe por um do modelo.

     `sempre` marca o que a planilha carrega mesmo vazio: sem essas
     colunas o registro não identifica ninguém, e escondê-las esconderia
     o problema. As demais somem quando não têm um único valor. */
  const CINZA_ESCURO = {theme:1, tint:0.249977111117893};
  const CINZA_CLARO  = {theme:1, tint:0.499984740745262};
  const MODELO = [
    {k:"mes",         titulo:"MÊS\nSOLICITAÇÃO",  largura:12.09, fundo:CINZA_ESCURO, sempre:true},
    {k:"data",        titulo:"DATA\nSOLICITAÇÃO", largura:16.36, fundo:CINZA_CLARO,  sempre:true, fmt:"mm-dd-yy"},
    {k:"solicitante", titulo:"SOLICITANTE",       largura:8.36,  fundo:CINZA_CLARO,  tituloAl:"center", al:"center"},
    {k:"empresa",     titulo:"EMPRESA DIARISTA",  largura:8.45,  fundo:CINZA_CLARO,  tituloAl:"center", al:"center"},
    {k:"groot",       titulo:"GROOT ID",          largura:8.82,  fundo:CINZA_CLARO,  tituloAl:"left",   al:"center", sempre:true},
    {k:"nome",        titulo:"NOME",              largura:42.63, fundo:CINZA_CLARO,  sempre:true},
    {k:"cargo",       titulo:"CARGO",             largura:20.82, fundo:CINZA_ESCURO},
    {k:"escala",      titulo:"ESCALA",            largura:17,    fundo:CINZA_CLARO,  al:"center", horario:true, sempre:true}
  ];
  const ORDEM_SAIDA = MODELO.map(c => c.k);
  const COL = k => MODELO.find(c => c.k === k);
  const MESES = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];

  const FONTE = "Aptos Narrow";
  const STY = {border: "FF000000"};

  const $ = id => document.getElementById(id);
  // conteúdo vindo da planilha vai para o HTML: escapar é obrigatório
  const esc = v => String(v==null?"":v).replace(/[&<>"]/g, c =>
    ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]));
  const drop = $("ex-drop"), fileInput = $("ex-fileInput");
  let workbook = null;
  let parsed = null;   // {op: [{date, id, solic:'id'|'meli'|'', cells:[...]}]}
  let leitura = {semData:[], abas:[], colunas:[]};   // o que a leitura do arquivo não aproveitou
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
  // "junho-26", como no modelo — com hífen, não com barra
  function mesLabel(key){ const [y,m]=key.split("-"); return MESES[+m-1]+"-"+y.slice(2); }
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

  /* Localiza o cabeçalho e cada coluna pelo próprio nome.
     Duas defesas contra o que quebrou Varginha:
       · o cabeçalho é a linha que reconhece MAIS colunas, não a primeira que
         contém "groot" — um "GROOT" solto numa legenda não manda mais;
       · entre dois candidatos para a mesma coluna, ganha o que tem dado
         embaixo. Uma coluna certa e vazia é indistinguível de uma errada, e
         escolher a vazia foi exatamente o que aconteceu. */
  function localizarColunas(grid){
    const limite = Math.min(grid.length, 25);
    let melhor = null;
    for(let i=0;i<limite;i++){
      const row = grid[i] || [];
      const col = {}, usadas = new Set();
      COLUNAS_SIGO.forEach(c => {
        const candidatos = [];
        row.forEach((celula, j) => {
          if(usadas.has(j)) return;
          const h = norm(celula);
          if(!h) return;
          const exato = c.nomes.indexOf(h);
          if(exato >= 0){ candidatos.push({j, peso: exato}); return; }
          if(c.nomes.some(n => h.includes(n))) candidatos.push({j, peso: 100});
        });
        if(!candidatos.length) return;
        // desempate: grafia mais exata, depois quem realmente tem dado embaixo
        candidatos.forEach(x => { x.preenchidas = contarPreenchidas(grid, i, x.j); });
        candidatos.sort((a,b) => (b.preenchidas>0) - (a.preenchidas>0)
                              || a.peso - b.peso
                              || b.preenchidas - a.preenchidas);
        col[c.k] = candidatos[0].j;
        usadas.add(candidatos[0].j);
      });
      const achadas = Object.keys(col).length;
      if(col.data !== undefined && col.groot !== undefined
         && (!melhor || achadas > melhor.achadas))
        melhor = {headerIdx:i, col, achadas};
    }
    return melhor;
  }

  function contarPreenchidas(grid, headerIdx, j){
    let n = 0;
    for(let i=headerIdx+1;i<Math.min(grid.length, headerIdx+200);i++){
      const c = grid[i] && grid[i][j];
      if(c !== null && c !== undefined && String(c).trim() !== "") n++;
    }
    return n;
  }

  function parseWorkbook(fileName){
    const warn = $("ex-warn1");
    parsed = {};
    const missing = [];
    /* O que a leitura descarta também é um problema — e era o mais invisível
       de todos: uma linha sem data legível não entrava nem no total de lidos,
       então nem a contagem denunciava a falta. */
    leitura = {semData:[], abas:[], colunas:[]};
    let globalMin = null, globalMax = null, totalRows = 0;

    for(const op of OPERATIONS){
      const sheetName = findSheet(op);
      if(!sheetName){
        missing.push(op); parsed[op] = [];
        leitura.abas.push({op, motivo:"Aba não encontrada no arquivo."});
        continue;
      }
      const ws = workbook.Sheets[sheetName];
      const grid = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, blankrows:false});

      const mapa = localizarColunas(grid);
      if(!mapa){
        missing.push(op + " (cabeçalho não encontrado)"); parsed[op] = [];
        leitura.abas.push({op, motivo:"Aba encontrada, mas sem um cabeçalho reconhecível nas 25 "
          + "primeiras linhas (não achei DATA nem GROOT ID) — a aba inteira ficou de fora."});
        continue;
      }
      // coluna que existe no padrão e não foi achada nesta aba: sai vazia, e isso se diz
      COLUNAS_SIGO.forEach(c => {
        if(c.k !== "horario" && mapa.col[c.k] === undefined)
          leitura.colunas.push({op, coluna:c.nomes[0].toUpperCase()});
      });

      const rows = [];
      const txt = v => (v === null || v === undefined) ? "" : String(v).trim();
      const val = (r, k) => mapa.col[k] === undefined ? "" : txt(r[mapa.col[k]]);
      /* Perda de dado é uma linha sem data legível que carrega ALGO QUE
         IDENTIFICA uma diária: pessoa (groot, nome) ou pelo menos empresa,
         cargo ou escala. Dois falsos alarmes reais ensinaram o limite:
         as tabelas auxiliares ao lado ("Dias/Qtd", cadastro GROOT) enchem
         a linha fora da tabela principal, e a aba XD tem ~9.700 linhas
         futuras com só o SOLICITANTE pré-arrastado ("MELI") — nenhum dos
         dois é uma diária perdida. */
      const colunasQueIdentificam = ["groot","nome","empresa","cargo","escala"]
        .map(k => mapa.col[k]).filter(j => j !== undefined);
      for(let i=mapa.headerIdx+1;i<grid.length;i++){
        const r = grid[i]; if(!r) continue;
        const dv = mapa.col.data === undefined ? null : r[mapa.col.data];
        let key = null;
        if(dv instanceof Date && !isNaN(dv)) key = dateKey(dv);
        else if(typeof dv === "number" && dv > 20000 && dv < 80000) key = excelSerialToKey(dv);
        if(!key){
          if(colunasQueIdentificam.some(j => txt(r[j]) !== ""))
            leitura.semData.push({op, linha:i+1, valor:txt(dv),
                                  nome:val(r,"nome"), groot:val(r,"groot")});
          continue;
        }

        const idRaw = val(r, "groot");
        rows.push({
          date: key,
          id: idRaw === "" ? null : idRaw,   // como estava na célula; a limpeza é na comparação
          solic: solicKind(val(r, "solicitante")),
          /* Campos nomeados: quem exibe ou exporta um registro não precisa
             conhecer a posição das colunas na origem. */
          mes:         mesLabel(key),
          nome:        val(r, "nome"),
          cargo:       val(r, "cargo"),
          escala:      val(r, "escala"),
          empresa:     val(r, "empresa"),
          solicitante: val(r, "solicitante"),
          horario:     val(r, "horario")     // vazio quando a origem não traz
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
    /* Um formato só para a tela e para a planilha: os balões contam estes
       tópicos e o Excel recebe exatamente estas linhas. */
    /* Só vale apontar a falta de horário onde há registro para preencher —
       filial sem diarista no período não é problema nenhum. */
    const semHorario = OPERATIONS
      .filter(op => !escalaHorarioDe(op) && dedup.porOperacao[op].rows.length)
      .map(op => ({op, registros: dedup.porOperacao[op].rows.length}));
    /* A saída é montada aqui, e não só na hora de baixar: o painel precisa
       dizer que coluna saiu e que célula foi preenchida ANTES de a pessoa
       abrir o arquivo. */
    const saida = OPERATIONS.filter(op => dedup.porOperacao[op].rows.length)
      .map(op => Object.assign({op}, montarSaida(op, dedup.porOperacao[op].rows)));
    lastTopicos = listarTopicos(lastAudit, mode === "nao" ? [] : dedup.descartados,
                                leitura, semHorario, saida);
    let tFilt=0, tDup=0, tFinal=0;
    for(const op of OPERATIONS){
      results[op] = dedup.porOperacao[op];
      tFilt += results[op].filtered; tDup += results[op].dups; tFinal += results[op].rows.length;
    }
    renderResults(ini, fim, tFilt, tDup, tFinal, mode === "nao", solic);
  });
  let lastDedup = null, lastAudit = null, lastTopicos = null, lastSolic = "ambos";

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
    /* A barra fica com os números da extração; os problemas viraram balões,
       cada um com seu nome e sua cor, logo abaixo. */
    $("ex-totalBar").innerHTML =
      "Registros encontrados: <b>"+tFilt.toLocaleString("pt-BR")+"</b>&nbsp;·&nbsp;" +
      "Únicos pessoa-dia: <b>"+tFinal.toLocaleString("pt-BR")+"</b>&nbsp;·&nbsp;" +
      "Duplicados removidos: <b>"+(dedupOff ? "—" : tDup.toLocaleString("pt-BR"))+"</b>";
    renderDetalheDedup(dedupOff);
    $("ex-num2").classList.add("done"); $("ex-num2").textContent="✓";
    $("ex-step3").classList.remove("disabled");
    $("ex-step3").scrollIntoView({behavior:"smooth",block:"nearest"});
  }

  /* ================================================================
     PAINEL DE REVISÃO
     ----------------------------------------------------------------
     Uma pilha de listas não diz por onde começar. Os balões respondem
     isso de relance: um por tópico, na ordem de quem precisa de atenção
     antes, com a contagem e a cor da gravidade. Clicar abre o tópico.
     ================================================================ */
  const GRAV_CLASSE = {alta:"g-alta", media:"g-media", baixa:"g-baixa", info:"g-info"};
  const slug = s => "ex-t-" + normalizeNome(s).toLowerCase().replace(/\s+/g, "-");

  function renderDetalheDedup(dedupOff){
    const alvo = $("ex-detalheDedup");
    if(!alvo) return;
    if(!lastDedup || !lastAudit){ alvo.innerHTML=""; alvo.classList.add("hidden"); return; }
    const topicos = lastTopicos || [];
    if(!topicos.length){
      alvo.innerHTML = '<p class="ex-limpo">Nenhum problema encontrado no período: '
        + 'todos os registros têm identificador, nome legível e nenhum conflito entre eles.</p>';
      alvo.classList.remove("hidden");
      return;
    }
    alvo.classList.remove("hidden");
    const baloes = resumirTopicos(topicos);
    const regra = lastDedup.resumo.modo==="dia"
      ? "Duplicado = mesmo GROOT normalizado + mesma data, valendo entre operações, abas e arquivos."
      : "Duplicado = mesmo GROOT normalizado no período inteiro, valendo entre operações, abas e arquivos.";

    let h = '<div class="ex-revisao-top">'
      + '<div><b>Revisão</b> — ' + topicos.length + ' ponto' + (topicos.length===1?'':'s')
      + ' em ' + baloes.length + ' tópico' + (baloes.length===1?'':'s') + '</div>'
      + '<button type="button" class="btn mini amber" id="ex-btnRelatorio">Baixar relatório (.xlsx)</button>'
      + '</div>'
      + '<div class="ex-baloes">'
      + baloes.map(b => '<button type="button" class="ex-balao ' + GRAV_CLASSE[b.gravidade]
          + '" data-alvo="' + slug(b.topico) + '">' + esc(b.topico)
          + '<b>' + b.total.toLocaleString("pt-BR") + '</b></button>').join("")
      + '</div>'
      + '<p class="regra">' + (dedupOff
          ? 'Deduplicação desligada — nada foi removido. Os demais tópicos continuam valendo: são do arquivo, não da remoção.'
          : regra + ' Fica sempre a <b>primeira ocorrência encontrada</b>.') + '</p>';

    // uma seção por tópico, na mesma ordem dos balões
    h += baloes.map(b => {
      const linhas = topicos.filter(t => t.topico === b.topico);
      const corpo = b.topico === "Mesmo nome, GROOTs diferentes"
          ? corpoConflito(lastAudit.mesmoNome, "GROOT",
              "A deduplicação <b>não uniu</b> estes registros — e não deveria: nome não é chave, "
              + "homônimo existe. Mas se for a mesma pessoa cadastrada duas vezes, ela está sendo "
              + "contada em dobro.")
        : b.topico === "Mesmo GROOT, nomes diferentes"
          ? corpoConflito(lastAudit.mesmoGroot, "Nome",
              "A deduplicação <b>uniu</b> estes registros, por serem o mesmo GROOT. Se os nomes "
              + "forem de pessoas diferentes, uma diária pode ter sido descartada como se fosse "
              + "repetida.")
          : corpoLista(linhas);
      return '<details class="ex-topico ' + GRAV_CLASSE[b.gravidade] + '" id="' + slug(b.topico) + '">'
        + '<summary>' + esc(b.topico) + ' <span class="ex-tag">' + b.total + '</span>'
        + '<span class="ex-grav">' + esc(GRAVIDADE_TXT[b.gravidade]) + '</span></summary>'
        + corpo + '</details>';
    }).join("");

    alvo.innerHTML = h;
    const btn = $("ex-btnRelatorio");
    if(btn) btn.addEventListener("click", baixarRelatorio);
    alvo.querySelectorAll(".ex-balao").forEach(b => b.addEventListener("click", () => {
      const d = document.getElementById(b.dataset.alvo);
      if(!d) return;
      d.open = true;
      d.scrollIntoView({behavior:"smooth", block:"nearest"});
    }));
  }

  /* Lista simples de um tópico. Só aparecem as colunas que carregam alguma
     informação: numa aba não lida não existe nome nem cargo, e sete traços
     seguidos não ajudam ninguém. */
  const COLUNAS = [
    {k:"nome",        t:"Nome"},
    {k:"groot",       t:"GROOT"},
    {k:"op",          t:"Filial / operação"},
    {k:"date",        t:"Data", fmt:dataBR},
    {k:"empresa",     t:"Empresa"},
    {k:"cargo",       t:"Cargo"},
    {k:"escala",      t:"Escala"},
    {k:"solicitante", t:"Solicitante"}
  ];
  const LIMITE_LINHAS = 300;

  function corpoLista(linhas){
    const cols = COLUNAS.filter(c => linhas.some(l => String(l[c.k] || "").trim() !== ""));
    const umDiag = linhas.every(l => l.diagnostico === linhas[0].diagnostico);
    let h = '<p class="regra">' + (umDiag ? esc(linhas[0].diagnostico) + " " : "")
      + '<b>O que fazer:</b> ' + esc(linhas[0].acao) + '</p>';

    // por filial: o problema quase sempre está concentrado em uma operação
    const porOp = new Map();
    linhas.forEach(l => { if(l.op) porOp.set(l.op, (porOp.get(l.op)||0)+1); });
    if(porOp.size > 1)
      h += '<p class="regra">Por filial: '
         + Array.from(porOp).map(([op,c]) => '<b>'+esc(op)+'</b> '+c).join(' · ') + '</p>';

    h += '<table class="ex-lista"><thead><tr>'
      + cols.map(c => '<th>'+esc(c.t)+'</th>').join("")
      + (umDiag ? "" : '<th>Diagnóstico</th>') + '</tr></thead><tbody>'
      + linhas.slice(0, LIMITE_LINHAS).map(l => '<tr>'
          + cols.map(c => '<td>' + (esc(c.fmt ? c.fmt(l[c.k]) : l[c.k]) || '<i>—</i>') + '</td>').join("")
          + (umDiag ? "" : '<td class="diag">'+esc(l.diagnostico)+'</td>') + '</tr>').join("")
      + '</tbody></table>';
    if(linhas.length > LIMITE_LINHAS)
      h += '<p class="regra">…e mais ' + (linhas.length-LIMITE_LINHAS)
         + '. O relatório .xlsx traz todos.</p>';
    return h;
  }

  /* Conflitos de identidade: o identificador e o nome contando histórias
     diferentes. O app só aponta — sem outra fonte, o dado não diz qual das
     versões está certa, e unir por conta própria apagaria gente. */
  const LIMITE_CASOS = 50, LIMITE_OCORR = 8;

  function corpoConflito(casos, colVariante, intro){
    const mostrados = casos.slice(0, LIMITE_CASOS);
    let h = '<p class="regra">'+intro+'</p>';
    h += mostrados.map(c => {
      const cabecalho = colVariante==="GROOT"
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
    return h;
  }

  // ---------- exportação com a estética do SIGO ----------
  function thinBorder(){
    const s = {style:"thin", color:{argb:STY.border}};
    return {top:s, left:s, bottom:s, right:s};
  }

  /* Monta os valores de saída de uma operação e decide o que vale a pena
     escrever. Duas regras, ambas sobre não empurrar ruído para quem lê:

       · coluna sem um único valor não vai para a planilha (a ESCALA de
         Patos de Minas eram 169 células em branco com um título em cima);
       · célula vazia numa coluna que só tem UM valor no arquivo inteiro é
         preenchida com ele — 112 dos 169 CARGO de Patos estavam vazios e
         os 57 preenchidos diziam todos "Diarista".

     O preenchimento para onde a certeza acaba: havendo dois valores
     distintos, a célula continua vazia. Escolher entre "Diarista" e
     "Diarista Dom. e Feriados" seria inventar. */
  function montarSaida(op, linhas){
    /* No modelo, ESCALA é o HORÁRIO. A linha do SIGO manda: horário escrito
       passa verbatim (Divinópolis, Patos), turno vira horário pela tabela da
       operação (AM/PM de Varginha, svc/xd de Pouso), vazio recebe o padrão.
       Token sem mapa fica como está e é contado para o painel apontar —
       ver resolverEscala em js/config.js. */
    const semMapa = new Map();
    const dados = linhas.map(r => {
      const e = resolverEscala(op, r.escala, r.horario);
      if(e.origem === "sem_mapa") semMapa.set(e.token, (semMapa.get(e.token)||0)+1);
      return {
        mes: r.mes,
        data: keyToUTCDate(r.date),
        solicitante: r.solicitante,
        empresa: r.empresa,
        groot: grootParaSaida(r.id),
        nome: r.nome,
        cargo: r.cargo,
        escala: e.valor
      };
    });

    const preenchidas = [];
    ORDEM_SAIDA.forEach(k => {
      const valores = dados.map(d => String(d[k] == null ? "" : d[k]).trim());
      const distintos = [...new Set(valores.filter(v => v !== ""))];
      if(distintos.length === 1 && valores.some(v => v === "")){
        const n = valores.filter(v => v === "").length;
        dados.forEach(d => { if(String(d[k]||"").trim() === "") d[k] = distintos[0]; });
        preenchidas.push({coluna:tituloDe(k), valor:distintos[0], celulas:n});
      }
    });

    const colunas = ORDEM_SAIDA.filter(k => COL(k).sempre
      || dados.some(d => String(d[k] == null ? "" : d[k]).trim() !== ""));
    return {dados, colunas, preenchidas,
            escalasSemMapa: Array.from(semMapa, ([token, vezes]) => ({token, vezes})),
            vazias: ORDEM_SAIDA.filter(k => colunas.indexOf(k) < 0).map(tituloDe)};
  }
  const tituloDe = k => COL(k).titulo.replace("\n", " ");

  /* Escreve a aba no formato de "Modelo diaristas.xlsx": mesmas larguras,
     mesmos dois tons de cinza no cabeçalho, mesma fonte, mesmos
     alinhamentos. Sem congelar linha e sem autofiltro, porque o modelo não
     tem — o arquivo gerado tem que passar por um do modelo. */
  function addStyledSheet(wbx, op){
    const ws = wbx.addWorksheet(op.substring(0,31), {
      views: [{state:"normal", showGridLines:true}]
    });
    const saida = montarSaida(op, results[op].rows);
    const cols = saida.colunas.map(COL);
    ws.columns = cols.map(c => ({width:c.largura}));

    const head = ws.addRow(cols.map(c => c.titulo));
    head.height = 39;
    head.eachCell((cell, i) => {
      const c = cols[i-1];
      cell.fill = {type:"pattern", pattern:"solid", fgColor:c.fundo};
      cell.font = {bold:true, size:10, color:{theme:0}, name:FONTE, family:2, scheme:"minor"};
      cell.alignment = {horizontal:c.tituloAl, vertical:"middle", wrapText:true};
      cell.border = thinBorder();
    });

    for(const d of saida.dados){
      const row = ws.addRow(cols.map(c => {
        // GROOT numérico entra como número, como no modelo; alfanumérico fica texto
        if(c.k === "groot") return (d.groot !== "" && !isNaN(Number(d.groot))) ? Number(d.groot) : d.groot;
        return d[c.k];
      }));
      row.eachCell({includeEmpty:true}, (cell, i) => {
        const c = cols[i-1];
        cell.border = thinBorder();
        cell.font = c.horario
          ? {size:9, color:{theme:1}, name:"Calibri", family:2}
          : {size:11, color:{theme:1}, name:FONTE, family:2, scheme:"minor"};
        if(c.fmt) cell.numFmt = c.fmt;
        if(c.horario) cell.fill = {type:"pattern", pattern:"solid", fgColor:{argb:"FFFFFFFF"}};
        if(c.al) cell.alignment = c.horario ? {horizontal:c.al, vertical:"middle"} : {horizontal:c.al};
      });
    }
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
     cadastro.

     A primeira aba é uma lista só, uma linha por problema, com as MESMAS
     colunas para todos os tópicos — é ela que se filtra. As abas seguintes
     são o mesmo conteúdo já separado por tópico, para quem prefere abrir e
     ler. Cada linha diz o que é, quão grave é e o que fazer. */
  const REL_COLS = [
    {t:"GRAVIDADE",         w:14, v:t => GRAVIDADE_TXT[t.gravidade]},
    {t:"TÓPICO",            w:30, v:t => t.topico},
    {t:"NOME",              w:36, v:t => t.nome},
    {t:"GROOT",             w:15, v:t => t.groot},
    {t:"FILIAL / OPERAÇÃO", w:20, v:t => t.op},
    {t:"DATA",              w:12, v:t => dataCell(t.date)},
    {t:"EMPRESA",           w:16, v:t => t.empresa},
    {t:"CARGO",             w:18, v:t => t.cargo},
    {t:"ESCALA",            w:10, v:t => t.escala},
    {t:"SOLICITANTE",       w:14, v:t => t.solicitante},
    {t:"O QUE ACONTECEU",   w:72, v:t => t.diagnostico, wrap:true},
    {t:"O QUE FAZER",       w:56, v:t => t.acao, wrap:true}
  ];
  // a cor diz a gravidade antes de qualquer leitura
  const FILL_GRAV = {alta:"FFF8D7DA", media:"FFFDF0D5", baixa:"FFE8EEF7", info:"FFF0F0F0"};
  const FONT_GRAV = {alta:"FF9B1C24", media:"FF8A5A00", baixa:"FF33507D", info:"FF666666"};

  function estiloLinha(row, gravidade){
    row.eachCell({includeEmpty:true}, (cell, col) => {
      cell.border = thinBorder();
      cell.font = {size:11, name:"Calibri"};
      if(cell.value instanceof Date) cell.numFmt = "dd/mm/yyyy";
      const wrap = !!(REL_COLS[col-1] && REL_COLS[col-1].wrap);
      cell.alignment = {horizontal:(typeof cell.value === "number" || cell.value instanceof Date)
        ? "right" : "left", vertical:"top", wrapText:wrap};
    });
    if(gravidade){
      const c = row.getCell(1);
      c.fill = {type:"pattern", pattern:"solid", fgColor:{argb:FILL_GRAV[gravidade]}};
      c.font = {size:11, name:"Calibri", bold:true, color:{argb:FONT_GRAV[gravidade]}};
      c.alignment = {horizontal:"center", vertical:"top"};
    }
  }

  function addRelSheet(wbx, nome, headers, widths, linhas, gravidades){
    const ws = wbx.addWorksheet(nome.substring(0,31), {views:[{state:"frozen", ySplit:1}]});
    ws.columns = widths.map(w => ({width:w}));
    const head = ws.addRow(headers);
    head.height = 30;
    head.eachCell(cell => {
      cell.fill = {type:"pattern", pattern:"solid", fgColor:{argb:STY.headerFill}};
      cell.font = {bold:true, color:{argb:STY.headerFont}, size:11, name:"Calibri"};
      cell.alignment = {horizontal:"center", vertical:"middle", wrapText:true};
      cell.border = thinBorder();
    });
    linhas.forEach((vals, i) => estiloLinha(ws.addRow(vals), gravidades && gravidades[i]));
    if(linhas.length) ws.autoFilter = {from:{row:1, column:1}, to:{row:1, column:headers.length}};
    return ws;
  }

  // data vira Date de verdade para o Excel filtrar e ordenar; chave ilegível fica como texto
  const dataCell = k => /^\d{4}-\d{2}-\d{2}$/.test(String(k||"")) ? keyToUTCDate(k) : (k||"");

  /* Aba de tópico: as mesmas colunas da lista geral, menos as duas que
     seriam idênticas em toda linha (o próprio tópico) ou vazias. */
  function addAbaTopico(wbx, nome, linhas){
    const cols = REL_COLS.filter(c => c.t !== "TÓPICO"
      && (c.t === "GRAVIDADE" || linhas.some(t => String(c.v(t) === 0 ? "0" : (c.v(t) || "")).trim() !== "")));
    addRelSheet(wbx, nome, cols.map(c => c.t), cols.map(c => c.w),
      linhas.map(t => cols.map(c => c.v(t))), linhas.map(t => t.gravidade));
  }

  // "Mesmo GROOT, nomes diferentes" → "2 Mesmo GROOT nomes difer" (31 caracteres, sem : \ / ? * [ ])
  function nomeAba(i, topico){
    return (i + " " + String(topico).replace(/[:\\\/?*\[\]]/g, " ")).substring(0, 31).trim();
  }

  async function baixarRelatorio(){
    const aud = lastAudit, ded = lastDedup, topicos = lastTopicos || [];
    if(!aud || !ded) return;
    const wbx = new ExcelJS.Workbook();
    const solicTxt = lastSolic === "ambos" ? "ID Logistics + MELI"
                   : (lastSolic === "id" ? "só ID Logistics" : "só MELI");
    const modoTxt = ded.resumo.modo === "dia" ? "uma linha por pessoa por dia"
                  : (ded.resumo.modo === "periodo" ? "uma linha por pessoa no período" : "sem remoção");
    const baloes = resumirTopicos(topicos);

    // ---- 0 · Leia-me: o relatório se explicando sozinho, longe do app
    const leiaMe = [
      ["COMO USAR", "", ""],
      ["Aba 'Todos os problemas'", "1 linha = 1 problema",
       "É a aba de trabalho: filtre por GRAVIDADE ou TÓPICO, marque o que já corrigiu. As demais abas são o mesmo conteúdo separado por tópico."],
      ["Coluna GRAVIDADE", "Grave / Revisar / Provável grafia / Informativo",
       "Grave = pode ter apagado uma diária de verdade. Revisar = impede identificar a pessoa. Provável grafia = quase certamente a mesma pessoa escrita diferente. Informativo = comportamento esperado, listado para conferência."],
      ["Coluna O QUE FAZER", "—", "O próximo passo concreto de cada linha. Quase sempre é corrigir o cadastro de origem."],
      ["Nada foi corrigido automaticamente", "—",
       "O aplicativo aponta; a correção é na origem. Sem outra fonte, o dado não diz qual das versões está certa, e unir por conta própria apagaria gente."],
      ["", "", ""],
      ["EXTRAÇÃO QUE GEROU ESTE RELATÓRIO", "", ""],
      ["Período", fmtBR(lastPeriod.ini)+" a "+fmtBR(lastPeriod.fim), "Recorte por data de solicitação."],
      ["Solicitante", solicTxt, "Filtro aplicado na extração."],
      ["Deduplicação", modoTxt, "Duplicado = mesmo GROOT normalizado + mesma data, valendo entre operações, abas e arquivos."],
      ["Registros no período", aud.resumo.registros, "Antes da deduplicação."],
      ["Únicos pessoa-dia", aud.resumo.registros - ded.resumo.duplicados, "O que a extração entrega."],
      ["", "", ""],
      ["O QUE FOI ENCONTRADO", "", ""]
    ];
    baloes.forEach(b => leiaMe.push([b.topico, b.total,
      GRAVIDADE_TXT[b.gravidade] + " — ver a aba correspondente."]));
    if(!baloes.length) leiaMe.push(["Nenhum problema encontrado", 0,
      "Todos os registros têm identificador, nome legível e nenhum conflito entre eles."]);
    addRelSheet(wbx, "0 Leia-me", ["ITEM","VALOR","O QUE SIGNIFICA"], [36,22,104], leiaMe);
    // esta aba é texto corrido, não tabela: filtro atrapalharia
    wbx.getWorksheet("0 Leia-me").autoFilter = undefined;
    wbx.getWorksheet("0 Leia-me").getColumn(3).alignment = {wrapText:true, vertical:"top"};

    // ---- 1 · Todos os problemas: a aba que se filtra
    if(topicos.length){
      addRelSheet(wbx, "1 Todos os problemas", REL_COLS.map(c => c.t), REL_COLS.map(c => c.w),
        topicos.map(t => REL_COLS.map(c => c.v(t))), topicos.map(t => t.gravidade));
    }

    // ---- 2..n · uma aba por tópico, na mesma ordem dos balões
    baloes.forEach((b, i) => addAbaTopico(wbx, nomeAba(i + 2, b.topico),
      topicos.filter(t => t.topico === b.topico)));

    await saveWorkbook(wbx, "Diaristas - Revisão - " + periodoNome() + ".xlsx");
  }

  $("ex-btnCombined").addEventListener("click", async () => {
    const wbx = new ExcelJS.Workbook();
    for(const op of OPERATIONS) addStyledSheet(wbx, op);
    await saveWorkbook(wbx, fileName("6 Operações"));
  });

})();
