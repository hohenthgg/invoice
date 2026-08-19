/* Retorno Simulado — o que o cliente provavelmente reconhecerá
   ================================================================

   Módulo PURO, sem DOM. Depende de js/dates.js e js/config.js,
   carregados antes no index.html.

   O QUE ESTE MÓDULO RESPONDE

   Três números por dia, e a distância entre eles:

     PREF ENVIADO   o headcount que o Labor apresenta ao cliente
     Q CLIENTE      o que o S&OP do dia suporta (SVC + SD + …)
     Q PÓS PREVISTO o que provavelmente será reconhecido

   A regra da previsão é deliberadamente ASSIMÉTRICA:

       Q PÓS PREVISTO = MIN(PREF, Q CLIENTE)

   O cliente pode CORTAR o que foi apresentado acima do planejamento,
   mas não vai PAGAR o que nem sequer foi enviado. Enviar 130 contra um
   S&OP de 138 não vira 138 — vira 130, e os 8 de folga são um alerta
   de subfaturamento para conferir o Labor, não uma receita a receber.
   Prever `Q PÓS = Q CLIENTE` nos dois sentidos seria otimista no lado
   errado, e é o erro que este módulo existe para não cometer.

   ISTO É UMA PREVISÃO

   Não há retorno oficial nenhum aqui. Tudo o que sai daqui é rotulado
   como simulado, e o dia que não puder ser reconstruído com segurança
   sai como REVISÃO NECESSÁRIA em vez de sair com um número inventado.

   O QUE ENTRA NO PREF

   A regra é a MESMA da Fusão de Linhas (js/fusao.js), e tem de ser:
   as duas reconstroem o mesmo número a partir do mesmo Labor.

     · cargo na lista de CARGOS_PREF (js/config.js) — auxiliar e
       operador; liderança e indiretos ficam fora por não estarem lá;
     · DATA DE INÍCIO legível, e o dia dentro de [início, fim], com fim
       vazio valendo até o fim do período;
     · o % RATEIO entra COM O SINAL QUE TEM.

   O sinal é onde este módulo já esteve errado. O PREF é uma quantidade
   FATURADA — o que a fatura apresenta, líquido dos estornos — e não um
   retrato do turno. Excluir os rateios negativos por serem "ajuste
   financeiro" inflava o PREF: nesta fatura eram 30 linhas de estorno,
   que somavam até 30 HC a mais num único dia de julho.

   As duas leituras continuam disponíveis: `pref` é o líquido, que vai
   ao confronto com o S&OP, e `bruto` conta só os positivos — quanta
   gente estava no turno. Só o primeiro é o PREF.
   ================================================================ */
"use strict";

const SIM_STATUS = {
  REDUCAO:  "reducao",          // PREF acima do S&OP: risco de corte
  ALINHADO: "alinhado",
  SUB:      "subfaturamento",   // PREF abaixo do S&OP
  REVISAO:  "revisao"           // não dá para reconstruir com segurança
};
const SIM_STATUS_LABEL = {
  reducao:"Possível correção", alinhado:"Alinhado",
  subfaturamento:"Possível subfaturamento", revisao:"Revisão necessária"
};

/* Motivos de uma linha do Labor ficar fora do PREF. `avisa` marca os
   que o usuário precisa ver: sair por cargo fora da lista é esperado
   para liderança, mas não para um cargo novo que ninguém declarou. */
const SIM_FORA = {
  ABS:         { chave:"abs",       texto:"linha de absenteísmo do template, não é pessoa" },
  /* Liderança e indiretos também saem por cargo, mas sair é o esperado
     deles — o aviso é para o cargo que ninguém declarou. A conta serve
     só para separar os dois casos, nunca para excluir: quem exclui é a
     lista de cargos, como na Fusão de Linhas. */
  INDIRETO:    { chave:"indireto",  texto:"cargo de liderança/indiretos — não entra no PREF" },
  CARGO:       { chave:"cargo",     texto:"cargo fora da lista do PREF" },
  SEM_INICIO:  { chave:"sem_inicio",texto:"sem DATA DE INÍCIO legível" }
};

const simNorm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"")
  .toLowerCase().replace(/\s+/g," ").trim();
const simEhIndireto = conta => /lideranca|indireto/.test(simNorm(conta));
const simEhAbs = l => simNorm(l.groot) === "abs" || /absenteismo/.test(simNorm(l.nome));

/* ================================================================
   QUEM ENTRA NO PREF

   Exatamente a regra da Fusão de Linhas (js/fusao.js): cargo na lista
   do PREF, DATA DE INÍCIO legível, fora a linha de ABS — e o rateio
   entra COM O SINAL QUE TEM.

   O sinal é o ponto, e é onde esta função já esteve errada. O PREF é
   uma quantidade FATURADA, não um retrato do turno: é o que a fatura
   apresenta ao cliente, líquido dos estornos. Uma linha de rateio -1
   referente a 27/07→31/07 reduz o que está sendo cobrado naqueles
   dias, e ignorá-la fazia o PREF sair alto — 30 linhas de estorno
   inflavam julho inteiro, até 30 HC num único dia.

   O headcount operacional bruto (só os positivos) continua sendo
   calculado e sai como `prefBruto`, porque as duas leituras são úteis:
   uma responde "quanto estou cobrando", a outra "quanta gente estava
   no turno". Só a primeira é o PREF.
   ================================================================ */
function simClassificarLinhas(labor){
  const dentro = [], fora = [];
  for(const l of labor){
    const marca = motivo => fora.push({ linha:l, motivo });
    if(simEhAbs(l)){ marca(SIM_FORA.ABS); continue; }
    if(!isValidYmd(l.inicio)){ marca(SIM_FORA.SEM_INICIO); continue; }
    if(!CARGOS_PREF.includes(simNorm(l.cargo))){
      marca(simEhIndireto(l.conta) ? SIM_FORA.INDIRETO : SIM_FORA.CARGO); continue;
    }
    dentro.push(l);
  }
  return { dentro, fora };
}

/* Rateio ausente ou ilegível vale 1, como na Fusão de Linhas: a linha
   existe e cobra uma pessoa. */
const simRateio = l =>
  (typeof l.rateio === "number" && isFinite(l.rateio)) ? l.rateio : 1;
const simAtivo = (l,d) => l.inicio <= d && (!isValidYmd(l.fim) || l.fim >= d);

/* ================================================================
   DIARISTAS DISPONÍVEIS NO DIA

   Quando o PREF fica abaixo do S&OP, a pergunta seguinte é quem
   poderia cobrir a diferença. O SIGO diz quem foi solicitado em cada
   dia, e por quem — ID Logistics ou o cliente.

   Só que estar solicitado não é o mesmo que estar disponível: se a
   pessoa já aparece no LABOR daquele dia, ela já está sendo cobrada
   como quadro fixo, e contá-la de novo seria contar duas vezes.

   A ocupação é medida pelo LÍQUIDO do dia, não pela existência da
   linha. Uma linha de rateio -1 cobrindo 27/07→31/07 é justamente o
   estorno do fixo para pagar aqueles dias como diária: a pessoa está
   livre, e é assim que os três casos reais desta fatura se comportam.
   ================================================================ */
function simGrootNum(v){ return String(v ?? "").replace(/\D/g,""); }

function simDisponibilidade(diaristas, labor, dias){
  if(!diaristas || !diaristas.length) return null;

  /* Líquido do LABOR por GROOT e por dia — todas as linhas, não só as
     do PREF: a pergunta é se a pessoa já está cobrada, em qualquer
     cargo. */
  const porGroot = new Map();
  for(const l of labor){
    const g = simGrootNum(l.groot);
    if(!g || !isValidYmd(l.inicio)) continue;
    if(!porGroot.has(g)) porGroot.set(g, []);
    porGroot.get(g).push(l);
  }
  const ocupado = (g, d) => {
    const ls = porGroot.get(g);
    if(!ls) return false;
    let liq = 0;
    for(const l of ls) if(simAtivo(l,d)) liq += simRateio(l);
    return liq > 0;
  };

  /* Uma pessoa por dia, e a ID tem precedência: o mesmo GROOT pedido
     pelos dois no mesmo dia é uma pessoa só, contada como ID. */
  const porDia = new Map();
  for(const r of diaristas){
    const g = simGrootNum(r.groot);
    if(!g || !isValidYmd(r.data)) continue;
    if(!porDia.has(r.data)) porDia.set(r.data, new Map());
    const m = porDia.get(r.data);
    const atual = m.get(g);
    if(atual === "id") continue;                       // já é ID, nada supera
    m.set(g, r.solic === "id" ? "id" : (atual || r.solic || ""));
  }

  const mapa = {};
  for(const d of dias){
    const m = porDia.get(d) || new Map();
    const cont = { total:m.size, id:0, meli:0, sem:0,
                   ocupados:0, disp:0, dispId:0, dispMeli:0, dispSem:0 };
    for(const [g, solic] of m){
      const chave = solic === "id" ? "id" : solic === "meli" ? "meli" : "sem";
      cont[chave]++;
      if(ocupado(g,d)){ cont.ocupados++; continue; }
      cont.disp++;
      cont[chave === "id" ? "dispId" : chave === "meli" ? "dispMeli" : "dispSem"]++;
    }
    mapa[d] = cont;
  }
  return mapa;
}

/* ================================================================
   O CONFRONTO, DIA A DIA
   ================================================================ */
function simDiagnostico(dia){
  const n = v => (v === null || v === undefined || !isFinite(Number(v)))
    ? "—" : Number(Number(v).toFixed(2)).toLocaleString("pt-BR");

  /* O texto do dia em revisão não cita valores: é justamente por não
     ter valor confiável que ele está em revisão. */
  if(dia.status === SIM_STATUS.REVISAO){
    return "Não foi possível reconstruir este dia com segurança: "+dia.revisao.join(" ")
      + " Nenhum valor foi previsto para não inventar número.";
  }
  const fontes = dia.blocos.map(b => b.rotulo+" "+n(b.valor)).join(" + ");
  if(dia.status === SIM_STATUS.REDUCAO){
    return "Foram enviados "+n(dia.gap)+" HC acima do S&OP total deste dia ("+fontes
      + " = "+n(dia.qCliente)+"). Caso o cliente limite o reconhecimento ao planejamento, "
      + "há risco de redução de "+n(dia.gap)+" HC.";
  }
  if(dia.status === SIM_STATUS.SUB){
    return "O Labor enviado está "+n(-dia.gap)+" HC abaixo do S&OP disponível ("+fontes
      + " = "+n(dia.qCliente)+"). Não presumir que o cliente aumentará automaticamente a "
      + "cobrança: verifique se existem pessoas faturáveis faltando no Labor."
      + (dia.diaristas
          ? " Havia "+n(dia.diaristas.disp)+" diarista(s) disponível(is) neste dia ("
            + n(dia.diaristas.dispId)+" ID, "+n(dia.diaristas.dispMeli)+" cliente)"
            + (dia.diaristas.ocupados ? ", fora "+n(dia.diaristas.ocupados)
                                        + " que já constam no LABOR do dia" : "")
            + " — dariam para cobrir "+n(dia.abatePossivel)+" HC."
          : "");
  }
  return "PREF e S&OP total batem em "+n(dia.qCliente)+" ("+fontes+"). Sem correção prevista.";
}

function simularRetorno(dados){
  const labor = dados.labor || [], blocos = dados.blocos || [];
  const diaristas = dados.diaristas || [];
  const ini = dados.periodo.ini, fim = dados.periodo.fim;
  const comp = dados.comp;
  const avisos = [];

  if(!blocos.length) return { erro:"Nenhum bloco de S&OP foi identificado na planilha operacional." };

  const { dentro, fora } = simClassificarLinhas(labor);

  /* Cargos fora da lista viram aviso nominal: o usuário precisa saber
     QUAL cargo ficou de fora para decidir se ele deveria entrar. */
  const cargosDuvidosos = [...new Set(fora.filter(f => f.motivo.chave === "cargo")
    .map(f => String(f.linha.cargo || "(cargo em branco)").trim()))];
  if(cargosDuvidosos.length){
    avisos.push({ tipo:"cargo", texto:
      "Cargo requer validação: "+cargosDuvidosos.join(", ")+". "
      + "Estas linhas estão em LABOR DIRETO mas o cargo não consta na lista de cargos do PREF, "
      + "então NÃO foram somadas — o PREF pode estar subestimado nos dias em que elas estão ativas. "
      + "Confirme se o cargo entra no PREF antes de usar a previsão." });
  }
  const nNeg = dentro.filter(l => simRateio(l) < 0).length;
  if(nNeg) avisos.push({ tipo:"estorno", texto:
    nNeg+" linha(s) do PREF têm rateio negativo e entram SUBTRAINDO, como na Fusão de Linhas. "
    + "São estornos e ajustes retroativos, e o PREF é a quantidade faturada — o que a fatura "
    + "apresenta, líquido deles. O headcount bruto do turno, só com os lançamentos positivos, "
    + "aparece separado em cada dia." });
  const nSemInicio = fora.filter(f => f.motivo.chave === "sem_inicio").length;
  if(nSemInicio) avisos.push({ tipo:"indefinida", texto:
    nSemInicio+" linha(s) ficaram fora do PREF por não ter DATA DE INÍCIO legível. Sem a data "
    + "não há como saber em que dias elas estariam ativas." });

  const listaDias = [];
  for(let d = ini; d <= fim; d = addDays(d,1)) listaDias.push(d);
  const disponibilidade = simDisponibilidade(diaristas, labor, listaDias);

  const dias = [];
  let totalPref = 0, totalBruto = 0, totalCliente = 0, totalPos = 0, hcAcima = 0, hcAbaixo = 0;
  let totalDisp = 0, totalAbate = 0, totalOcupados = 0;
  const contagem = { reducao:0, alinhado:0, subfaturamento:0, revisao:0 };

  for(let d = ini; d <= fim; d = addDays(d,1)){
    /* Célula vazia do S&OP é AUSÊNCIA, não zero. `Number(null)` vale 0,
       e deixar isso passar transformaria um dia sem planejamento num
       S&OP zerado — o PREF inteiro daquele dia viraria "risco de corte"
       inventado. Por isso o valor só existe se for número de verdade. */
    const porBloco = blocos.map(b => {
      const bruto = Object.prototype.hasOwnProperty.call(b.dias, d) ? b.dias[d] : null;
      const n = Number(bruto);
      return { rotulo:b.rotulo,
               valor: (bruto === null || bruto === undefined || bruto === "" || !isFinite(n)) ? null : n };
    });
    const faltando = porBloco.filter(b => b.valor === null || !isFinite(b.valor));

    /* PREF é o líquido — a mesma soma da Fusão de Linhas, com o sinal
       do rateio. `bruto` conta só os positivos: é quanta gente estava
       no turno, e serve para explicar a diferença, nunca para comparar
       com o S&OP. */
    let pref = 0, bruto = 0;
    for(const l of dentro){
      if(!simAtivo(l,d)) continue;
      const r = simRateio(l);
      pref += r;
      if(r > 0) bruto += r;
    }
    pref = Math.round(pref*1e6)/1e6;
    bruto = Math.round(bruto*1e6)/1e6;

    const revisao = [];
    if(faltando.length) revisao.push("o S&OP de "+faltando.map(b=>b.rotulo).join(" e ")+" não tem valor para este dia.");

    const dia = { data:d, blocos:porBloco, revisao, pref, bruto };
    if(revisao.length){
      dia.qCliente = null; dia.qPos = null; dia.gap = null; dia.correcao = null;
      dia.status = SIM_STATUS.REVISAO;
    } else {
      const qCliente = porBloco.reduce((s,b) => s + b.valor, 0);
      const qPos = Math.min(pref, qCliente);
      dia.qCliente = qCliente;
      dia.qPos = qPos;
      dia.gap = Math.round((pref - qCliente)*1e6)/1e6;
      dia.correcao = Math.round((qPos - pref)*1e6)/1e6;
      dia.status = dia.gap > 0 ? SIM_STATUS.REDUCAO
                 : dia.gap < 0 ? SIM_STATUS.SUB
                 : SIM_STATUS.ALINHADO;
      totalPref += pref; totalBruto += bruto; totalCliente += qCliente; totalPos += qPos;
      if(dia.gap > 0) hcAcima += dia.gap; else hcAbaixo += -dia.gap;
    }
    /* Quantos diaristas poderiam cobrir a falta do dia — limitado à
       própria falta, como no abate de ABS: um dia nunca fica positivo
       por sobra de diarista. */
    if(disponibilidade){
      const c = disponibilidade[d];
      dia.diaristas = c;
      const falta = (dia.gap !== null && dia.gap < 0) ? -dia.gap : 0;
      dia.abatePossivel = Math.min(c.disp, falta);
      totalDisp += c.disp;
      totalOcupados += c.ocupados;
      totalAbate += dia.abatePossivel;
    }
    dia.diagnostico = simDiagnostico(dia);
    contagem[dia.status]++;
    dias.push(dia);
  }

  return {
    periodo:{ ini, fim }, comp, blocos:blocos.map(b => b.rotulo), dias, avisos,
    linhas:{ dentro:dentro.length, fora },
    totais:{
      pref: Math.round(totalPref*100)/100,
      bruto: Math.round(totalBruto*100)/100,
      estornos: nNeg,
      cliente: Math.round(totalCliente*100)/100,
      pos: Math.round(totalPos*100)/100,
      /* O número que resume o risco: HC-dia que o cliente pode cortar. */
      hcEmRisco: Math.round(hcAcima*100)/100,
      hcAbaixo: Math.round(hcAbaixo*100)/100,
      ...(disponibilidade ? { diaristasDisp: totalDisp, diaristasOcupados: totalOcupados,
                              abatePossivel: totalAbate } : {}),
      dias: dias.length, ...contagem
    }
  };
}

/* ================================================================
   PREF × QF DO CLIENTE — E O QUE FAZER COM O EXCESSO

   Aqui a Validação para de só apontar o excesso e passa a resolvê-lo.
   Nada disso é decidido neste arquivo: quem decide é o motor de
   js/equalizacao.js — o MESMO que a Fusão de Linhas usa. Esta função
   é tradução e explicação.

     · traduz o Labor da fatura (AAAAMMDD, % RATEIO com sinal) para o
       eixo do motor;
     · monta a curva alvo — aqui ela é CONSTANTE: o QF do cliente, o
       mesmo número em todos os dias do período, e não o S&OP diário;
     · chama eqEqualizar();
     · e devolve cada ação já contada em dias e HC, com o motivo, e
       cruzada com a base de diaristas.

   A diferença de vocabulário com a Fusão é só essa: lá a curva alvo
   vem do retorno oficial, dia a dia, e vem separada por FULL_TIME e
   PART_TIME. Aqui o alvo é um número só, então há um grupo só. Mesmo
   Labor + mesma curva ⇒ mesmo plano, porque é a mesma função.

   E o plano é SUGESTÃO. A matemática fecha a curva; quem sabe se a
   movimentação aconteceu de verdade é a operação.
   ================================================================ */
const SIM_GRUPO = "QUADRO";

/* Dias soltos viram faixas: [16,17,18,20] → 16–18 e 20. */
function simFaixas(dias){
  const out = [];
  for(const d of dias){
    const u = out[out.length-1];
    if(u && addDays(u.ate,1) === d) u.ate = d; else out.push({ de:d, ate:d });
  }
  return out;
}

const SIM_EQ_MOTIVO = {
  retirar: "A vigência inteira desta linha cai em dias que continuam acima do QF depois "
    + "de ela sair. Retirá-la zera excesso sem deixar nenhum dia do período descoberto.",
  adiar: "O excesso está no começo da vigência, e a pessoa volta a ser necessária depois. "
    + "Adiar o início preserva a pessoa e corta só os dias em sobra.",
  encurtar: "O excesso está no fim da vigência, e a DATA FIM atual já cai dentro dele — "
    + "antecipar não descobre nenhum dia posterior.",
  pausar: "A demanda cai e volta a subir. Em vez de retirar a pessoa, o contrato fecha no "
    + "início do vale e reabre no fim: mesma pessoa, mesmo GROOT, dois períodos."
};

function simPlanoEqualizacao(dados){
  const labor = dados.labor || [];
  const alvo = Number(dados.alvo);
  const opc = dados.opcoes || {};
  const ini = dados.periodo.ini, fim = dados.periodo.fim;

  if(!isFinite(alvo) || alvo <= 0)
    return { erro:"Informe o QF do cliente para calcular a equalização." };

  const { dentro } = simClassificarLinhas(labor);
  const listaDias = [];
  for(let d = ini; d <= fim; d = addDays(d,1)) listaDias.push(d);
  const eixo = listaDias.map(eqDeYmd);

  const pessoas = dentro.map((l,i) => ({
    id:i, ini:eqDeYmd(l.inicio), fim:isValidYmd(l.fim) ? eqDeYmd(l.fim) : null,
    rateio:simRateio(l), desempate:l.matricula || "", grupo:SIM_GRUPO }));
  const dias = listaDias.map(d => ({ dia:eqDeYmd(d), alvo, grupo:SIM_GRUPO }));

  const plano = eqEqualizar(pessoas, dias, {
    pausaDesde: isValidYmd(opc.pausaDesde) ? eqDeYmd(opc.pausaDesde) : null,
    permitirAdiarInicio: opc.permitirAdiarInicio !== false });

  /* A curva antes e depois do plano. A de depois é a prova: se o plano
     está certo, ela encosta no alvo sem furar para baixo. */
  const ativoAntes = (p,n) => p.ini <= n && n <= (p.fim == null ? EQ_INF : p.fim);
  const linhasDias = listaDias.map((d,k) => {
    let antes = 0, depois = 0;
    for(const p of pessoas){
      if(ativoAntes(p,eixo[k])) antes += p.rateio;
      if(eqAtivoApos(p, plano.acoes.get(p.id), eixo[k])) depois += p.rateio;
    }
    antes = Math.round(antes*1e6)/1e6; depois = Math.round(depois*1e6)/1e6;
    return { data:d, pref:antes, alvo, dif:Math.round((antes-alvo)*1e6)/1e6,
             prefPos:depois, difPos:Math.round((depois-alvo)*1e6)/1e6 };
  });

  /* Quem também aparece no SIGO, e em que dias. Mesma regra de
     precedência do resto do módulo: pedido pelos dois é um pedido só,
     contado como ID. */
  const diarPorGroot = new Map();
  for(const r of (dados.diaristas || [])){
    const g = simGrootNum(r.groot);
    if(!g || !isValidYmd(r.data)) continue;
    if(!diarPorGroot.has(g)) diarPorGroot.set(g, new Map());
    const m = diarPorGroot.get(g);
    if(m.get(r.data) === "id") continue;
    m.set(r.data, r.solic === "id" ? "id" : (m.get(r.data) || r.solic || ""));
  }

  const acoes = [];
  plano.acoes.forEach((a,id) => {
    const p = pessoas[id], l = dentro[id];
    const tipo = eqTipoAcao(a);
    if(!tipo) return;
    /* Impacto: os dias em que a pessoa contava e deixa de contar. */
    const perdidos = listaDias.filter((d,k) =>
      ativoAntes(p,eixo[k]) && !eqAtivoApos(p,a,eixo[k]));
    const hc = Math.round(perdidos.length * p.rateio * 1e6)/1e6;

    /* Cruzamento com o SIGO: a mesma pessoa aparecendo como diarista
       justamente nos dias que o plano tira do fixo é a explicação
       operacional da correção, não uma coincidência. */
    const mapaDiar = diarPorGroot.get(simGrootNum(l.groot));
    let diarista = null;
    if(mapaDiar){
      const nosDias = perdidos.filter(d => mapaDiar.has(d));
      const todos = [...mapaDiar.keys()].filter(d => d >= ini && d <= fim).sort((x,y)=>x-y);
      diarista = { dias:nosDias, total:todos.length,
        id:todos.filter(d => mapaDiar.get(d) === "id").length,
        meli:todos.filter(d => mapaDiar.get(d) === "meli").length };
    }

    acoes.push({ tipo, id,
      linha:{ groot:l.groot, nome:l.nome, cargo:l.cargo, matricula:l.matricula || "",
              inicio:l.inicio, fim:l.fim, rateio:simRateio(l), linha:l.linha },
      novoInicio: a.novo_ini != null ? eqParaYmd(a.novo_ini) : null,
      novoFim:    a.novo_fim !== undefined ? eqParaYmd(a.novo_fim) : null,
      pausas: (a.pausas || []).map(x => ({ fim:eqParaYmd(x.fim), ini:eqParaYmd(x.ini) })),
      impacto:{ dias:perdidos.length, hc, faixas:simFaixas(perdidos) },
      motivo: SIM_EQ_MOTIVO[tipo], diarista });
  });

  /* Ordem de leitura: maior impacto primeiro — é a ordem em que uma
     pessoa conferiria o plano. */
  acoes.sort((a,b) => Math.abs(b.impacto.hc) - Math.abs(a.impacto.hc)
    || a.linha.nome.localeCompare(b.linha.nome));

  const chavesYmd = o => Object.fromEntries(Object.entries(o || {})
    .map(([n,q]) => [eqParaYmd(+n), q]));
  const revisar = chavesYmd(plano.incluir.__revisar?.[SIM_GRUPO]);
  const falta   = chavesYmd(plano.incluir[SIM_GRUPO]?.dias);

  const conta = t => acoes.filter(a => a.tipo === t).length;
  return {
    alvo, periodo:{ ini, fim }, dias:linhasDias, acoes, revisar, falta,
    opcoes:{ permitirAdiarInicio: opc.permitirAdiarInicio !== false,
             pausaDesde: isValidYmd(opc.pausaDesde) ? opc.pausaDesde : null },
    totais:{
      pessoas: pessoas.length,
      retirar: conta(EQ_ACAO.RETIRAR), adiar: conta(EQ_ACAO.ADIAR),
      encurtar: conta(EQ_ACAO.ENCURTAR), pausar: conta(EQ_ACAO.PAUSAR),
      hc: Math.round(acoes.reduce((s,a) => s + a.impacto.hc, 0)*100)/100,
      diasAcima:  linhasDias.filter(d => d.dif > 0).length,
      diasAbaixo: linhasDias.filter(d => d.dif < 0).length,
      excessoAntes: Math.round(linhasDias.reduce((s,d) => s + Math.max(d.dif,0), 0)*100)/100,
      excessoDepois: Math.round(linhasDias.reduce((s,d) => s + Math.max(d.difPos,0), 0)*100)/100,
      /* Se isto for maior que zero o plano criou falta — não deveria
         acontecer nunca, e é o número que denuncia. */
      faltaCriada: Math.round(linhasDias.reduce((s,d) =>
        s + Math.max(0, Math.min(d.dif,0) - Math.min(d.difPos,0)), 0)*100)/100,
      comDiarista: acoes.filter(a => a.diarista && a.diarista.dias.length).length
    }
  };
}

/* ================================================================
   O ARQUIVO SIMULADO

   A aba de retorno é escrita com os MESMOS cabeçalhos que a Fusão de
   Linhas procura, para o arquivo poder ser usado lá como um retorno de
   teste. A Fusão localiza a aba pelo cabeçalho, não pelo nome, então o
   nome pode dizer com todas as letras que o arquivo é simulado.
   ================================================================ */
const SIM_AVISO_METADADOS =
  "ARQUIVO SIMULADO. Q Pós foi estimado por MIN(PREF, soma do S&OP dos blocos). "
  + "Não corresponde ao retorno oficial do cliente.";

function simLinhasRetorno(sim){
  const linhas = [];
  for(const dia of sim.dias){
    if(dia.status === SIM_STATUS.REVISAO){
      linhas.push({ data:dia.data, tipo:"FULL_TIME", pref:dia.pref, qPos:null, desvio:null,
        ocorrencia:"REVISÃO NECESSÁRIA — "+dia.diagnostico });
      continue;
    }
    linhas.push({ data:dia.data, tipo:"FULL_TIME", pref:dia.pref, qPos:dia.qPos,
      desvio: Math.round((dia.pref - dia.qPos)*1e6)/1e6,
      ocorrencia:"PREVISÃO SIMULADA — "+dia.diagnostico });
  }
  return linhas;
}

if(typeof module !== "undefined" && module.exports){
  module.exports = { simularRetorno, simClassificarLinhas, simLinhasRetorno,
                     simPlanoEqualizacao, SIM_GRUPO,
                     SIM_STATUS, SIM_STATUS_LABEL, SIM_AVISO_METADADOS };
}
