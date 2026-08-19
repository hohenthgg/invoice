/* Motor de equalização — js/equalizacao.js
   ================================================================

   Este motor decide o que mexer no Labor quando sobra gente. Ele
   nasceu dentro da Fusão de Linhas e nunca teve teste: rodava só no
   navegador, dentro de uma IIFE que monta a tela na carga. Agora que
   a Validação de Template usa o MESMO motor, ele precisa de dois
   tipos de prova.

   1. AS PROTEÇÕES

   O que separa este motor de um "escolha N nomes do dia com excesso"
   é o que ele se RECUSA a fazer:

     · não retira quem tem um único dia ativo sem excesso — retirar
       criaria falta nesse dia;
     · não antecipa a DATA FIM de quem seguiria trabalhando depois do
       trecho em excesso;
     · não pausa quem não volta a ser necessário — isso é um fim
       disfarçado de vale;
     · não trata dia com alvo 0 como demanda zero — escala não
       publicada não é motivo para cortar ninguém.

   Cada uma dessas recusas é um teste aqui. São elas que fazem o plano
   fechar a curva sem furá-la para baixo, e o teste que morde é sempre
   o do sentido oposto: se a proteção sumir, o motor "resolve" o dia e
   quebra outro.

   2. A PARIDADE

   Duas abas chamam este motor. Se cada uma tivesse a sua cópia, elas
   divergiriam — e a divergência apareceria como duas listas de nomes
   diferentes para a mesma fatura, sem ninguém saber qual valia. O
   teste de paridade monta o MESMO Labor nos dois vocabulários (Date +
   retorno dia a dia, do lado da Fusão; AAAAMMDD + QF constante, do
   lado da Validação), roda os dois caminhos completos — adaptador
   incluído — e exige plano idêntico, ação por ação, data por data.
   ================================================================ */
"use strict";
const { load } = require("./load.js");

const ctx = load(["dates.js","config.js","equalizacao.js","fusao.js","simulacao.js"],
                 ["EQ_INF","EQ_ACAO","CARGOS_PREF"]);
const { eqEqualizar, eqTipoAcao, simPlanoEqualizacao, ymd, EQ_ACAO } = ctx;
const fusaoEqualizar = ctx.window.Fusao.equalizar;

let pass = 0, fail = 0;
function check(ok, label, extra){
  if(ok){ pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}
const secao = t => console.log("\n" + t);

/* ================================================================
   ATALHOS

   O motor trabalha num eixo de inteiros onde d+1 é o dia seguinte.
   Nos testes de unidade os dias são 100, 101, 102… — mais legível que
   data, e é exatamente o que o motor enxerga.
   ================================================================ */
const D0 = 100;
const dias = (n, alvo) => Array.from({length:n}, (_,k) => ({ dia:D0+k, alvo, grupo:"G" }));
const P = (id, ini, fim, extra) => ({ id, ini:D0+ini, fim:fim==null?null:D0+fim,
  rateio:1, desempate:String(id), grupo:"G", ...extra });
/** Quadro por dia depois do plano — a curva que o cliente veria. */
function curva(pessoas, dias, plano){
  return dias.map(d => pessoas.reduce((s,p) =>
    s + (ctx.eqAtivoApos(p, plano.acoes.get(p.id), d.dia) ? p.rateio : 0), 0));
}
const acaoDe = (plano, id) => eqTipoAcao(plano.acoes.get(id));

/* ================================================================
   FASE 1 — RETIRAR
   ================================================================ */
secao("Retirar linha inteira");
{
  /* Alvo 2, e três pessoas o período inteiro: sobra uma, todos os dias. */
  const ds = dias(5, 2);
  const ps = [P(0,0,null), P(1,0,null), P(2,0,null)];
  const plano = eqEqualizar(ps, ds);
  check(plano.acoes.size === 1, "sobra de 1 em todos os dias resolve com 1 retirada",
        plano.acoes.size + " ações");
  check(acaoDe(plano,2) === EQ_ACAO.RETIRAR, "quem entrou por último é o retirado");
  check(curva(ps,ds,plano).every(v => v === 2), "a curva fecha no alvo em todos os dias");
}
{
  /* A mesma sobra, mas o candidato natural tem um dia ativo SEM
     excesso: retirá-lo furaria a curva para baixo naquele dia. */
  const ds = [{dia:100,alvo:2,grupo:"G"},{dia:101,alvo:3,grupo:"G"},{dia:102,alvo:2,grupo:"G"}];
  const ps = [P(0,0,null), P(1,0,null), P(2,0,null)];
  const plano = eqEqualizar(ps, ds);
  const c = curva(ps,ds,plano);
  check(c[1] >= 3, "não retira quem cobre um dia sem excesso — o dia 101 não pode cair", "curva "+c);
  check(acaoDe(plano,2) !== EQ_ACAO.RETIRAR, "a retirada total é descartada por um único dia apertado");
}

/* ================================================================
   FASE 2 — ADIAR INÍCIO
   ================================================================ */
secao("Adiar data de início");
{
  /* 4 pessoas fixas contra alvo 4, e uma quinta que entra no dia 2
     quando o alvo ainda é 4 e sobe para 5 no dia 4. Ela é necessária —
     só não no começo. */
  const ds = [{dia:100,alvo:4,grupo:"G"},{dia:101,alvo:4,grupo:"G"},
              {dia:102,alvo:4,grupo:"G"},{dia:103,alvo:4,grupo:"G"},
              {dia:104,alvo:5,grupo:"G"},{dia:105,alvo:5,grupo:"G"}];
  const ps = [P(0,0,null),P(1,0,null),P(2,0,null),P(3,0,null),P(4,2,null)];
  const plano = eqEqualizar(ps, ds);
  check(acaoDe(plano,4) === EQ_ACAO.ADIAR, "quem entra cedo demais tem o início adiado, não é retirado");
  check(plano.acoes.get(4).novo_ini === 104,
        "o novo início é o primeiro dia em que ela deixa de sobrar",
        "novo_ini " + plano.acoes.get(4)?.novo_ini);
  check(curva(ps,ds,plano).join(",") === "4,4,4,4,5,5", "a curva encosta no alvo dos dois lados");
}
{
  const ds = [{dia:100,alvo:4,grupo:"G"},{dia:101,alvo:4,grupo:"G"},
              {dia:102,alvo:4,grupo:"G"},{dia:103,alvo:4,grupo:"G"},
              {dia:104,alvo:5,grupo:"G"},{dia:105,alvo:5,grupo:"G"}];
  const ps = [P(0,0,null),P(1,0,null),P(2,0,null),P(3,0,null),P(4,2,null)];
  const plano = eqEqualizar(ps, ds, { permitirAdiarInicio:false });
  check(acaoDe(plano,4) !== EQ_ACAO.ADIAR,
        "com adiar desligado o motor não mexe em admissão — sobra vira 'revisar'");
  check(!!plano.incluir.__revisar?.G, "e o excesso que sobrou é reportado, não escondido");
}

/* ================================================================
   FASE 3 — PAUSAR E RETOMAR
   ================================================================ */
secao("Pausar e retomar");
{
  /* O alvo cai no meio e volta a subir: um vale. Quem está ativo o
     período inteiro não deve ser retirado por causa dele. */
  const ds = [{dia:100,alvo:3,grupo:"G"},{dia:101,alvo:2,grupo:"G"},
              {dia:102,alvo:2,grupo:"G"},{dia:103,alvo:3,grupo:"G"},
              {dia:104,alvo:3,grupo:"G"}];
  const ps = [P(0,0,null),P(1,0,null),P(2,0,null)];
  const semPausa = eqEqualizar(ps, ds);
  check(semPausa.acoes.size === 0 && !!semPausa.incluir.__revisar,
        "sem data de pausa, o vale sai para revisão manual — ninguém é cortado");

  const comPausa = eqEqualizar(ps, ds, { pausaDesde:100 });
  const alvo = ds.map(d => d.alvo);
  check(comPausa.acoes.size === 1, "com a pausa liberada, um contrato cobre o vale inteiro");
  const p = [...comPausa.acoes.values()][0];
  check(!!p.pausas && p.pausas.length === 1 && p.pausas[0].fim === 100 && p.pausas[0].ini === 103,
        "fecha no dia anterior ao vale e reabre no dia seguinte",
        JSON.stringify(p.pausas));
  check(curva(ps,ds,comPausa).join(",") === alvo.join(","),
        "a curva bate com o alvo em todos os dias, inclusive na retomada");
}
{
  /* O mesmo formato, mas SEM retomada: o excesso vai até o fim. Pausar
     aqui seria um fim disfarçado. */
  const ds = [{dia:100,alvo:3,grupo:"G"},{dia:101,alvo:2,grupo:"G"},{dia:102,alvo:2,grupo:"G"}];
  const ps = [P(0,0,null),P(1,0,null),P(2,0,null)];
  const plano = eqEqualizar(ps, ds, { pausaDesde:100 });
  const temPausa = [...plano.acoes.values()].some(a => a.pausas && a.pausas.length);
  check(!temPausa, "excesso que vai até o fim do período não é pausa — é fim, e a pausa é recusada");
}

/* ================================================================
   FASE 4 — ENCURTAR FIM
   ================================================================ */
secao("Antecipar data fim");
{
  /* Alvo 2 subindo para 3: quem termina DENTRO do trecho em excesso
     pode ter o fim antecipado. */
  const ds = [{dia:100,alvo:3,grupo:"G"},{dia:101,alvo:3,grupo:"G"},
              {dia:102,alvo:2,grupo:"G"},{dia:103,alvo:2,grupo:"G"}];
  const ps = [P(0,0,null),P(1,0,null),P(2,0,3)];
  const plano = eqEqualizar(ps, ds);
  check(acaoDe(plano,2) === EQ_ACAO.ENCURTAR || acaoDe(plano,2) === EQ_ACAO.RETIRAR,
        "o excesso concentrado no fim é resolvido pelo fim do contrato",
        String(acaoDe(plano,2)));
  check(curva(ps,ds,plano).join(",") === "3,3,2,2", "e a curva fecha exatamente no alvo");
}
{
  /* Agora quem está em excesso no fim do trecho SEGUE ativo depois.
     Antecipar o fim dele criaria falta nos dias seguintes. */
  const ds = [{dia:100,alvo:2,grupo:"G"},{dia:101,alvo:2,grupo:"G"},
              {dia:102,alvo:3,grupo:"G"},{dia:103,alvo:3,grupo:"G"}];
  const ps = [P(0,0,null),P(1,0,null),P(2,0,null)];
  const plano = eqEqualizar(ps, ds, { permitirAdiarInicio:false });
  const c = curva(ps,ds,plano);
  check(c[2] >= 3 && c[3] >= 3, "não antecipa o fim de quem trabalha depois do trecho", "curva "+c);
}

/* ================================================================
   RATEIO NEGATIVO ENTRA NA CURVA, MAS NUNCA VIRA CANDIDATO

   A linha de rateio -1 é estorno: ela REDUZ o que a fatura cobra
   naqueles dias, então tem de contar na curva. O que ela não pode ser
   é escolhida por nenhuma das quatro fases — retirar um -1 SOBE a
   curva, ou seja, resolveria o excesso ao contrário. É o tipo de
   defeito que o total esconde: o plano "fecha" no papel e o arquivo
   gerado sai pior do que entrou.
   ================================================================ */
secao("Rateio ≤ 0 conta na curva, mas não é candidato");
{
  /* Quatro positivos e um estorno: a curva do dia é 3, e o alvo é 2. */
  const ds = dias(4, 2);
  const ps = [P(0,0,null), P(1,0,null), P(2,0,null), P(3,0,null), P(4,0,null,{rateio:-1})];
  const plano = eqEqualizar(ps, ds);
  const byId = new Map(ps.map(p => [p.id, p]));
  const escolhidos = [...plano.acoes.keys()];
  check(escolhidos.every(i => byId.get(i).rateio > 0),
        "nenhuma ação recai sobre linha de rateio ≤ 0",
        "escolhidos " + JSON.stringify(escolhidos.map(i => byId.get(i).rateio)));
  check(plano.acoes.size === 1, "e uma ação bastou — a sobra era de 1");
  check(!plano.acoes.has(4), "o estorno especificamente não foi tocado");
  check(curva(ps,ds,plano).every(v => v === 2),
        "e a curva ainda fecha no alvo — o estorno seguiu contando nela");
}
{
  /* A prova de que ele CONTA na curva: as mesmas três pessoas, com e
     sem o estorno, dão planos diferentes. Ignorá-lo faria o motor ver
     um excesso que a fatura não tem e cortar alguém à toa. */
  const ds = dias(4, 2);
  const comEstorno = eqEqualizar([P(0,0,null),P(1,0,null),P(2,0,null),P(3,0,null,{rateio:-1})], ds);
  const semEstorno = eqEqualizar([P(0,0,null),P(1,0,null),P(2,0,null)], ds);
  check(comEstorno.acoes.size === 0, "com o estorno na curva não há excesso — ninguém é cortado");
  check(semEstorno.acoes.size === 1, "sem ele, o mesmo quadro teria uma pessoa a mais e sairia uma");
}
{
  /* Cada fase, uma a uma, com o estorno como único candidato "óbvio"
     para aquele formato de excesso. Nenhuma pode mordê-lo. */
  const casos = [
    ["retirar", dias(3,1), [P(0,0,null), P(1,0,null,{rateio:-1})], {}],
    ["adiar início", [{dia:100,alvo:1,grupo:"G"},{dia:101,alvo:1,grupo:"G"},{dia:102,alvo:2,grupo:"G"}],
      [P(0,0,null), P(1,0,null,{rateio:-1}), P(2,0,null)], {}],
    ["pausar", [{dia:100,alvo:2,grupo:"G"},{dia:101,alvo:1,grupo:"G"},{dia:102,alvo:2,grupo:"G"}],
      [P(0,0,null), P(1,0,null), P(2,0,null,{rateio:-1})], { pausaDesde:100 }],
    ["antecipar fim", [{dia:100,alvo:2,grupo:"G"},{dia:101,alvo:1,grupo:"G"},{dia:102,alvo:1,grupo:"G"}],
      [P(0,0,null), P(1,0,null), P(2,0,2,{rateio:-1})], { permitirAdiarInicio:false }]
  ];
  for(const [nome, ds, ps, opc] of casos){
    const plano = eqEqualizar(ps, ds, opc);
    const byId = new Map(ps.map(p => [p.id, p]));
    check([...plano.acoes.keys()].every(i => byId.get(i).rateio > 0),
          "fase '" + nome + "' não escolhe linha de rateio ≤ 0",
          JSON.stringify([...plano.acoes.keys()].map(i => byId.get(i).rateio)));
  }
}
{
  /* Rateio 0 sairia como ação que não muda nada — enfeite no plano. */
  const ds = dias(3, 1);
  const plano = eqEqualizar([P(0,0,null), P(1,0,null,{rateio:0}), P(2,0,null)], ds);
  check(!plano.acoes.has(1), "rateio 0 também fica fora: a ação seria um no-op");
}
/* ================================================================
   O QUE ELE NÃO RESOLVE — E DIZ QUE NÃO RESOLVEU
   ================================================================ */
secao("Falta, excesso residual e dias sem alvo");
{
  const ds = dias(3, 5);
  const ps = [P(0,0,null),P(1,0,null),P(2,0,null)];
  const plano = eqEqualizar(ps, ds);
  check(plano.acoes.size === 0, "quando falta gente o motor não mexe em ninguém");
  check(plano.incluir.G && Object.keys(plano.incluir.G.dias).length === 3,
        "a falta é relatada dia a dia, para quem chama decidir");
  check(plano.incluir.G.dias[100] === 2, "e com o tamanho certo: faltam 2");
}
{
  /* Alvo 0 é escala não publicada, não demanda zero. */
  const ds = [{dia:100,alvo:0,grupo:"G"},{dia:101,alvo:3,grupo:"G"},{dia:102,alvo:3,grupo:"G"}];
  const ps = [P(0,0,null),P(1,0,null),P(2,0,null)];
  const plano = eqEqualizar(ps, ds);
  check(plano.acoes.size === 0, "dia com alvo 0 é ignorado, não vira corte de 3 pessoas");
  check(!plano.incluir.G && !plano.incluir.__revisar, "e não vira falta nem excesso");
}

/* ================================================================
   GRUPOS INDEPENDENTES
   ================================================================ */
secao("Grupos");
{
  const ds = [{dia:100,alvo:1,grupo:"FULL_TIME"},{dia:100,alvo:1,grupo:"PART_TIME"}];
  const ps = [{id:0,ini:100,fim:null,rateio:1,desempate:"a",grupo:"FULL_TIME"},
              {id:1,ini:100,fim:null,rateio:1,desempate:"b",grupo:"FULL_TIME"},
              {id:2,ini:100,fim:null,rateio:1,desempate:"c",grupo:"PART_TIME"}];
  const plano = eqEqualizar(ps, ds);
  check(plano.acoes.size === 1 && plano.acoes.has(1),
        "o excesso de um grupo não é resolvido tirando gente do outro");
}

/* ================================================================
   PARIDADE — FUSÃO × VALIDAÇÃO

   Mesmo Labor, mesmo período, mesma curva alvo. A Fusão recebe o alvo
   como um retorno dia a dia com Q Pós = 182; a Validação recebe o QF
   182 e monta a curva constante sozinha. Os dois planos têm de ser o
   mesmo objeto, ação por ação.
   ================================================================ */
secao("Paridade entre a Fusão de Linhas e a Validação Template");

const PER = { ini: ymd(2026,7,16), fim: ymd(2026,8,15) };
const QF = 182;
const dt = v => new Date(Date.UTC(Math.trunc(v/10000), Math.trunc(v/100)%100-1, v%100));
const paraYmd = d => d.getUTCFullYear()*10000 + (d.getUTCMonth()+1)*100 + d.getUTCDate();

/** Uma descrição, dois vocabulários. É esse o ponto do teste: os dois
 *  lados descrevem o MESMO Labor, cada um do seu jeito. */
function fixture(){
  const gente = [];
  const add = (ini, fim, rateio) => gente.push({
    groot:"90"+String(gente.length).padStart(5,"0"),
    nome:"COLABORADOR "+String(gente.length).padStart(3,"0"),
    matricula:"M"+String(gente.length).padStart(4,"0"),
    cargo:"Auxiliar de Apoio Log I", conta:"LABOR DIRETO",
    inicio:ini, fim, rateio, linha:gente.length+2 });

  for(let i=0;i<176;i++) add(ymd(2026,7,16), null, 1);       // quadro estável
  add(ymd(2026,7,16), ymd(2026,7,27), 1);                    // sai no meio do período
  add(ymd(2026,7,16), ymd(2026,8,3), 1);
  add(ymd(2026,7,20), null, 1);                              // entradas escalonadas
  add(ymd(2026,7,22), null, 1);
  add(ymd(2026,7,22), ymd(2026,8,10), 1);
  add(ymd(2026,7,28), null, 1);
  add(ymd(2026,7,28), ymd(2026,8,15), 1);
  add(ymd(2026,8,1),  null, 1);
  add(ymd(2026,8,5),  ymd(2026,8,15), 1);
  add(ymd(2026,7,16), ymd(2026,7,31), -1);                   // estorno: entra subtraindo
  return gente;
}

function planoDaFusao(gente, opc){
  const labor = gente.map(p => ({ groot:p.groot, nome:p.nome, mat:p.matricula, cargo:p.cargo,
    regime:"", ini:dt(p.inicio), fim:p.fim ? dt(p.fim) : null, rateio:p.rateio,
    pt:false, elig:true, rawRow:[] }));
  const ret = [];
  for(let d = PER.ini; d <= PER.fim; d = ctx.addDays(d,1))
    ret.push({ d:dt(d), tipo:"full_time", pref:QF, qpos:QF, desv:null, ocor:"" });
  const plano = fusaoEqualizar(labor, ret, opc.pausaDesde ? dt(opc.pausaDesde) : null,
                               opc.permitirAdiarInicio !== false);
  const out = [];
  plano.acoes.forEach((a,id) => out.push({ id, tipo:eqTipoAcao(a),
    novoInicio: a.novo_ini ? paraYmd(a.novo_ini) : null,
    novoFim: a.novo_fim !== undefined ? paraYmd(a.novo_fim) : null,
    pausas: (a.pausas||[]).map(p => ({ fim:paraYmd(p.fim), ini:paraYmd(p.ini) })) }));
  return out.sort((x,y) => x.id - y.id);
}

function planoDaValidacao(gente, opc){
  const r = simPlanoEqualizacao({ labor:gente, periodo:PER, alvo:QF, opcoes:opc });
  return { r, plano: r.acoes.map(a => ({ id:a.id, tipo:a.tipo,
    novoInicio:a.novoInicio, novoFim:a.novoFim, pausas:a.pausas }))
    .sort((x,y) => x.id - y.id) };
}

for(const opc of [{ permitirAdiarInicio:true },
                  { permitirAdiarInicio:false },
                  { permitirAdiarInicio:true, pausaDesde: PER.ini }]){
  const rotulo = "adiar=" + (opc.permitirAdiarInicio !== false)
    + (opc.pausaDesde ? " · pausa liberada" : "");
  const gente = fixture();
  const f = planoDaFusao(gente, opc);
  const { r, plano: v } = planoDaValidacao(gente, opc);
  check(f.length > 0, "há o que equalizar neste Labor (" + rotulo + ")", f.length + " ações");
  check(JSON.stringify(f) === JSON.stringify(v),
        "Fusão e Validação chegam ao MESMO plano (" + rotulo + ")",
        "fusão " + JSON.stringify(f.slice(0,3)) + " · validação " + JSON.stringify(v.slice(0,3)));
  check(r.totais.faltaCriada === 0, "o plano não cria falta em nenhum dia (" + rotulo + ")",
        "falta criada " + r.totais.faltaCriada);
}

{
  /* A prova de que o teste morde: mude a curva alvo de um lado só e os
     planos têm de divergir. Sem isso, "iguais" poderia ser só duas
     listas vazias. */
  const gente = fixture();
  const a = planoDaValidacao(gente, { permitirAdiarInicio:true }).plano;
  const b = simPlanoEqualizacao({ labor:gente, periodo:PER, alvo:QF-3,
    opcoes:{ permitirAdiarInicio:true } }).acoes.map(x => x.id).sort();
  check(JSON.stringify(a.map(x => x.id).sort()) !== JSON.stringify(b),
        "alvo diferente dá plano diferente — a comparação não é vácua");
}

/* ================================================================
   O QUE A VALIDAÇÃO ACRESCENTA POR CIMA DO MOTOR
   ================================================================ */
secao("Leitura da Validação: impacto, curva e diaristas");

const gente = fixture();
const res = simPlanoEqualizacao({ labor:gente, periodo:PER, alvo:QF,
  opcoes:{ permitirAdiarInicio:true } });

check(res.dias.length === 31, "o período sai dia a dia, 16/07 a 15/08", res.dias.length + " dias");
check(res.dias.every(d => d.alvo === QF), "o alvo é o QF, constante — não o S&OP do dia");
check(res.totais.excessoDepois === 0,
      "depois do plano não sobra excesso em dia nenhum", "sobrou " + res.totais.excessoDepois);
check(res.totais.excessoAntes > 0, "e antes havia — o plano fez algo");
check(res.acoes.every(a => a.impacto.dias > 0 && a.motivo),
      "toda sugestão vem com impacto contado em dias e com o motivo escrito");
check(res.acoes.every(a => a.linha.nome && a.linha.groot),
      "e nomeia a pessoa: nenhuma sugestão sai sem nome e GROOT");
check(res.acoes.every(a => gente.some(p => p.groot === a.linha.groot)),
      "as sugestões saem do Labor — o motor não inventa pessoa");
/* A mesma garantia do motor, cobrada na saída que o usuário vê: nenhuma
   ação do plano da Validação carrega rateio ≤ 0. */
check(res.acoes.every(a => a.linha.rateio > 0),
      "nenhuma ação do plano tem rateio original ≤ 0",
      JSON.stringify(res.acoes.map(a => a.linha.rateio).filter(r => r <= 0)));
check(gente.some(p => p.rateio <= 0),
      "…e o Labor de teste tem estorno, senão a garantia acima seria vácua");
{
  /* Pelo caminho completo da Validação, onde o rateio vem do % RATEIO
     da fatura — e onde o estorno é o caso real que motivou a regra. */
  const comPausa = simPlanoEqualizacao({ labor:gente, periodo:PER, alvo:QF,
    opcoes:{ permitirAdiarInicio:true, pausaDesde:PER.ini } });
  check(gente.some(p => p.rateio <= 0), "o Labor do teste tem mesmo linha de estorno");
  check(res.acoes.every(a => a.linha.rateio > 0) && comPausa.acoes.every(a => a.linha.rateio > 0),
        "nenhuma sugestão da Validação recai sobre linha de rateio ≤ 0",
        JSON.stringify(res.acoes.filter(a => a.linha.rateio <= 0).map(a => a.linha.nome)));
}
{
  const somaDif = res.acoes.reduce((s,a) => s + a.impacto.hc, 0);
  check(Math.abs(somaDif - res.totais.hc) < 1e-6, "o total de HC bate com a soma das ações");
}
{
  const semAlvo = simPlanoEqualizacao({ labor:gente, periodo:PER, alvo:0, opcoes:{} });
  check(!!semAlvo.erro, "sem QF informado o módulo recusa em vez de equalizar contra nada");
}
{
  /* Cruzamento com o SIGO: quem o plano tira do fixo justamente nos
     dias em que aparece como diarista é a explicação operacional da
     correção — e é isso que precisa chegar na tela. */
  const alvoAcao = res.acoes[0];
  const diasTirados = [];
  for(const f of alvoAcao.impacto.faixas)
    for(let d = f.de; d <= f.ate; d = ctx.addDays(d,1)) diasTirados.push(d);
  const diaristas = diasTirados.slice(0,3).map(d =>
    ({ groot:alvoAcao.linha.groot, data:d, solic:"id", nome:alvoAcao.linha.nome, empresa:"" }));
  const comDiar = simPlanoEqualizacao({ labor:gente, periodo:PER, alvo:QF, diaristas,
    opcoes:{ permitirAdiarInicio:true } });
  const mesma = comDiar.acoes.find(a => a.linha.groot === alvoAcao.linha.groot);
  check(!!mesma.diarista && mesma.diarista.dias.length === diasTirados.slice(0,3).length,
        "quem sai do fixo e aparece como diarista nos MESMOS dias é sinalizado",
        JSON.stringify(mesma.diarista));
  check(comDiar.totais.comDiarista === 1, "e o total conta só quem tem essa coincidência");
  check(JSON.stringify(comDiar.acoes.map(a => a.id).sort()) ===
        JSON.stringify(res.acoes.map(a => a.id).sort()),
        "a base de diaristas NÃO muda a matemática — só acrescenta contexto");
}
{
  const copia = JSON.parse(JSON.stringify(gente));
  simPlanoEqualizacao({ labor:gente, periodo:PER, alvo:QF, opcoes:{} });
  check(JSON.stringify(gente) === JSON.stringify(copia),
        "as linhas do Labor saem da equalização exatamente como entraram");
}

/* ================================================================
   QUEM ENTRA PARA COBRIR A FALTA

   O outro lado do plano. Aqui o teste que morde é o da PRIORIDADE: um
   total certo com a repartição errada gasta diarista do cliente com
   interno sobrando, e o total não denuncia nada. É o mesmo defeito que
   tests/abs-prioridade.js protege no abate "ambos", e a mesma regra.
   ================================================================ */
secao("Inclusão de diaristas para cobrir a falta");

const { simInclusoes } = ctx;
const D = ymd(2026,7,20);
const diar = (groot, data, solic, nome) => ({ groot, data, solic, nome:nome || "P"+groot, empresa:"" });
const lab = (groot, ini, fim, rateio) => ({ groot, nome:"FIXO "+groot, cargo:"Auxiliar de Apoio Log I",
  conta:"LABOR DIRETO", matricula:"", inicio:ini, fim, rateio:rateio == null ? 1 : rateio, linha:1 });

{
  const r = simInclusoes({ falta:{ [D]:2 }, labor:[],
    diaristas:[diar("1",D,"meli"), diar("2",D,"id"), diar("3",D,"meli"), diar("4",D,"id")] });
  const solics = r.pessoas.map(p => p.solic).sort();
  check(r.totais.incluido === 2, "cobre exatamente a falta do dia — nem mais, nem menos");
  check(solics.join(",") === "id,id", "com ID sobrando, o do cliente NÃO é usado", solics.join(","));
}
{
  const r = simInclusoes({ falta:{ [D]:3 }, labor:[],
    diaristas:[diar("1",D,"meli"), diar("2",D,"id"), diar("3",D,"meli")] });
  check(r.totais.id === 1 && r.totais.meli === 2,
        "esgotado o ID, o do cliente complementa — e só o que faltou",
        r.totais.id+" ID, "+r.totais.meli+" cliente");
}
{
  /* Já está no LABOR do dia: contá-lo seria cobrar duas vezes. */
  const r = simInclusoes({ falta:{ [D]:2 }, labor:[lab("2",D,D)],
    diaristas:[diar("2",D,"id"), diar("5",D,"id")] });
  check(r.totais.incluido === 1 && r.pessoas[0].groot === "5",
        "quem já está cobrado no LABOR do dia não é incluído de novo");
  check(r.totais.descoberto === 1, "e o que não deu para cobrir sai como descoberto, com o tamanho");
}
{
  /* Rateio ≤ 0 é estorno, e estorno não devolve a pessoa ao mercado: o
     fixo daquele dia continua ativo, e incluí-la de novo como diarista
     seria cobrá-la duas vezes. */
  const r = simInclusoes({ falta:{ [D]:1 }, labor:[lab("2",D,D,1), lab("2",D,D,-1)],
    diaristas:[diar("2",D,"id")] });
  check(r.totais.incluido === 0,
        "o estorno ao lado do fixo não libera a pessoa — ela segue cobrada no dia");
  check(r.totais.descoberto === 1, "e a falta fica declarada em vez de coberta com repetição");
}
{
  /* Só o estorno, sem fixo ativo no dia: aí não há cobrança nenhuma e a
     pessoa está mesmo livre. */
  const r = simInclusoes({ falta:{ [D]:1 }, labor:[lab("2",D,D,-1)],
    diaristas:[diar("2",D,"id")] });
  check(r.totais.incluido === 1, "sem lançamento positivo no dia, a pessoa está livre");
}
{
  const d1 = ymd(2026,7,20), d2 = ymd(2026,7,21), d3 = ymd(2026,7,23);
  const r = simInclusoes({ falta:{ [d1]:1, [d2]:1, [d3]:1 }, labor:[],
    diaristas:[diar("7",d1,"id"), diar("7",d2,"id"), diar("7",d3,"id")] });
  check(r.pessoas.length === 1, "a mesma pessoa em três dias é uma pessoa, não três");
  check(r.pessoas[0].faixas.length === 2,
        "dias seguidos viram uma faixa; o dia solto vira outra",
        JSON.stringify(r.pessoas[0].faixas));
}
{
  const r = simInclusoes({ falta:{ [D]:2.5 }, labor:[],
    diaristas:[diar("1",D,"id"), diar("2",D,"id"), diar("3",D,"id")] });
  check(r.totais.incluido === 2, "falta fracionária arredonda para baixo — passar do alvo é o erro");
}
{
  const r = simInclusoes({ falta:{ [D]:2 }, labor:[], diaristas:[] });
  check(r.totais.incluido === 0 && r.totais.descoberto === 2,
        "sem base de diaristas não se inventa ninguém — a falta fica declarada");
}

/* ================================================================
   NINGUÉM ENTRA NUM DIA EM QUE NÃO FOI SOLICITADO

   A regra mais dura da inclusão: cada pessoa-dia escrito no arquivo
   tem de existir na base do SIGO NAQUELE DIA. O jeito de quebrar isso
   não é escolher a pessoa errada — é preencher o buraco entre dois
   dias soltos dela, que parece arrumação e é diária inventada.
   ================================================================ */
{
  const d1 = ymd(2026,7,20), d2 = ymd(2026,7,21), d3 = ymd(2026,7,22);
  /* Solicitada em 20 e 22, NÃO em 21 — e 21 tem falta e nenhum outro
     candidato. A tentação de "completar o intervalo" mora aqui. */
  const r = simInclusoes({ falta:{ [d1]:1, [d2]:1, [d3]:1 }, labor:[],
    diaristas:[diar("5",d1,"id"), diar("5",d3,"id")] });
  const p = r.pessoas[0];
  check(p.dias.join(",") === d1+","+d3,
        "o dia do buraco não é preenchido: ela só entra onde foi solicitada",
        JSON.stringify(p.dias));
  check(p.faixas.length === 2,
        "…e as faixas mostram os dois trechos, sem emendar por cima do buraco",
        JSON.stringify(p.faixas));
  check(r.dias[d2].incluido === 0 && r.dias[d2].descoberto === 1,
        "o dia sem candidato fica declarado como descoberto");
}
{
  /* A reconferência final roda mesmo quando não tem nada a recusar, e
     é sobre ela que a garantia se apoia — não sobre a construção. */
  const d1 = ymd(2026,7,20);
  const r = simInclusoes({ falta:{ [d1]:2 }, labor:[],
    diaristas:[diar("6",d1,"id"), diar("7",d1,"meli")] });
  check(r.conferencia.verificados === 2 && r.conferencia.recusados.length === 0,
        "toda pessoa-dia escolhida passa pelo índice cru da base antes de virar linha",
        JSON.stringify(r.conferencia));
}
{
  /* Por dentro do caminho completo, com a fatura inteira: nenhuma
     pessoa-dia do plano está fora da base. */
  const gente2 = fixture();
  const base = [];
  for(let d = PER.ini; d <= PER.fim; d = ctx.addDays(d,1))
    if(d % 2 === 0) for(let k=0;k<10;k++) base.push(diar("81"+k, d, "id"));
  const r = simPlanoEqualizacao({ labor:gente2, periodo:PER, alvo:QF, diaristas:base,
    opcoes:{ permitirAdiarInicio:true } });
  const cru = new Set(base.map(x => x.groot+"|"+x.data));
  const fora = r.inclusoes.pessoas.flatMap(p => p.dias
    .filter(d => !cru.has(p.groot+"|"+d)).map(d => p.groot+"|"+d));
  check(fora.length === 0,
        "no plano completo, toda pessoa-dia incluída existe na base naquele dia",
        JSON.stringify(fora.slice(0,5)));
  check(r.inclusoes.conferencia.verificados > 0,
        "…e a conferência de fato rodou sobre alguma coisa — a garantia não é vácua",
        String(r.inclusoes.conferencia.verificados));
  check(r.inclusoes.pessoas.every(p => p.dias.every(d => d % 2 === 0)),
        "e os dias sem solicitação nenhuma continuam vazios, sem emenda");
}
{
  /* Pelo caminho completo: o plano da Validação traz as inclusões
     prontas quando a base do SIGO vem junto. */
  const gente2 = fixture();
  const dias = [];
  for(let d = PER.ini; d <= PER.fim; d = ctx.addDays(d,1)) dias.push(d);
  const base = dias.flatMap(d => [diar("8000001",d,"id"), diar("8000002",d,"meli")]);
  const r = simPlanoEqualizacao({ labor:gente2, periodo:PER, alvo:QF, diaristas:base,
    opcoes:{ permitirAdiarInicio:true } });
  check(!!r.inclusoes && r.inclusoes.pessoas.length > 0,
        "o plano completo traz as inclusões junto com as retiradas");
  check(r.inclusoes.pessoas[0].solic === "id", "e o primeiro da fila é o da ID");
  const semBase = simPlanoEqualizacao({ labor:gente2, periodo:PER, alvo:QF,
    opcoes:{ permitirAdiarInicio:true } });
  check(semBase.inclusoes === null, "sem a base, o plano não finge que sabe quem incluir");
  check(JSON.stringify(r.acoes.map(a => a.id)) === JSON.stringify(semBase.acoes.map(a => a.id)),
        "e as inclusões não mudam nenhuma das retiradas — são lados independentes do plano");
}

/* ================================================================
   A DIÁRIA QUE A FATURA JÁ LANÇA

   O defeito que estes testes travam: a falta era calculada contra o
   LABOR só, e a aba DIARISTAS da fatura nunca era lida. Num dia em que
   a fatura já pagava 15 diárias o app pedia mais 23, e 29 pessoa-dia
   apareciam nas duas listas — cobrança dobrada, com nome e data.
   ================================================================ */
secao("Diárias já lançadas na fatura");
{
  const gente2 = fixture();
  const d1 = ymd(2026,7,20);
  const diarias = Array.from({length:5}, (_,k) => ({ groot:"70000"+k, data:d1, quantidade:1 }));
  const sem = simPlanoEqualizacao({ labor:gente2, periodo:PER, alvo:QF, opcoes:{} });
  const com = simPlanoEqualizacao({ labor:gente2, periodo:PER, alvo:QF, diarias, opcoes:{} });
  const dSem = sem.dias.find(x => x.data === d1), dCom = com.dias.find(x => x.data === d1);
  check(dCom.quadro === dCom.pref + 5,
        "o quadro do dia é o fixo MAIS as diárias já lançadas",
        dCom.pref + " + " + dCom.diarias + " = " + dCom.quadro);
  check(dCom.dif === dSem.dif + 5,
        "e é o quadro, não o fixo, que vai ao confronto com o QF");
  check((com.falta[d1] || 0) === Math.max(0, (sem.falta[d1] || 0) - 5),
        "a falta do dia sai descontada das diárias que a fatura já paga",
        "sem " + (sem.falta[d1]||0) + " · com " + (com.falta[d1]||0));
}
{
  /* A diária conta na curva mas não pode virar ação: quem equaliza mexe
     no quadro fixo, não em diária que já aconteceu. */
  const gente2 = fixture();
  const diarias = [];
  for(let d = PER.ini; d <= PER.fim; d = ctx.addDays(d,1))
    for(let k=0;k<8;k++) diarias.push({ groot:"7100"+k, data:d, quantidade:1 });
  const r = simPlanoEqualizacao({ labor:gente2, periodo:PER, alvo:QF, diarias, opcoes:{} });
  check(r.acoes.every(a => gente2.some(p => p.groot === a.linha.groot)),
        "nenhuma ação do plano cai sobre uma diária — só sobre linha do Labor");
  check(r.totais.diarias === 8 * r.dias.length,
        "e o total de diárias do período é contado", r.totais.diarias);
}
{
  /* O teste que morde de verdade: quem já está lançado como diária
     naquele dia não pode ser escolhido de novo para o mesmo dia. */
  const D2 = ymd(2026,7,20);
  const r = simInclusoes({ falta:{ [D2]:2 }, labor:[],
    diarias:[{ groot:"91", data:D2, quantidade:1 }],
    diaristas:[diar("91",D2,"id"), diar("92",D2,"id")] });
  check(r.totais.incluido === 1 && r.pessoas[0].groot === "92",
        "quem já está lançado como diária no dia não é incluído de novo",
        JSON.stringify(r.pessoas.map(p => p.groot)));
  check(r.totais.descoberto === 1,
        "e a falta que sobrou é declarada em vez de coberta com repetição");
}
{
  const D2 = ymd(2026,7,20);
  const r = simInclusoes({ falta:{ [D2]:1 }, labor:[],
    diarias:[{ groot:"91", data:ymd(2026,7,21), quantidade:1 }],
    diaristas:[diar("91",D2,"id")] });
  check(r.totais.incluido === 1,
        "diária em OUTRO dia não bloqueia a pessoa — a ocupação é por dia");
}

/* ================================================================
   FASE 5 — RETIRAR A DIÁRIA ACIMA DO QF

   As quatro fases do motor mexem no quadro FIXO. Quando elas terminam e
   ainda sobra excesso, o excesso não é mais fixo — é diária lançada por
   cima de um quadro que já está no teto, e cortar mais gente fixa para
   compensar criaria falta em outro dia.

   O que morde aqui é a ORDEM: um total certo com a repartição errada
   devolve a diária do cliente enquanto sobra interna, e o total não
   denuncia. É o mesmo defeito do abate "ambos", pelo avesso.
   ================================================================ */
secao("Retirar diária acima do QF");

const { simCortarDiarias } = ctx;
const dia = (groot, data, extra) => ({ groot, data, quantidade:1, nome:"P"+groot, ...extra });

{
  const D3 = ymd(2026,7,24);
  const r = simCortarDiarias({ residual:{ [D3]:2 },
    diarias:[dia("1",D3), dia("2",D3), dia("3",D3), dia("4",D3)],
    diaristas:[diar("1",D3,"meli"), diar("2",D3,"id"), diar("3",D3,"meli"), diar("4",D3,"id")] });
  check(r.totais.cortado === 2, "corta exatamente o excesso do dia — nem mais, nem menos");
  check(r.totais.id === 2 && r.totais.meli === 0,
        "sai a interna primeiro; a do cliente é a última a sair",
        r.totais.id+" internas, "+r.totais.meli+" do cliente");
}
{
  const D3 = ymd(2026,7,24);
  const r = simCortarDiarias({ residual:{ [D3]:3 },
    diarias:[dia("1",D3), dia("2",D3), dia("3",D3)],
    diaristas:[diar("1",D3,"meli"), diar("2",D3,"id"), diar("3",D3,"meli")] });
  check(r.totais.id === 1 && r.totais.meli === 2,
        "esgotada a interna, a do cliente entra no corte — e só o que faltou");
}
{
  /* Nunca passar do excesso: com 5 de excesso e 2 diárias, saem 2 e o
     resto fica declarado. Cortar quadro fixo aqui criaria falta. */
  const D3 = ymd(2026,7,24);
  const r = simCortarDiarias({ residual:{ [D3]:5 },
    diarias:[dia("1",D3), dia("2",D3)], diaristas:[] });
  check(r.totais.cortado === 2 && r.totais.restante === 3,
        "só há o que cortar até acabar a diária; o resto sai declarado",
        JSON.stringify(r.totais));
}
{
  const D3 = ymd(2026,7,24), D4 = ymd(2026,7,25);
  const r = simCortarDiarias({ residual:{ [D3]:1 },
    diarias:[dia("1",D4), dia("2",D4)], diaristas:[] });
  check(r.totais.cortado === 0 && r.totais.restante === 1,
        "diária de OUTRO dia não serve para resolver este — o corte é por dia");
}
{
  const r = simCortarDiarias({ residual:{}, diarias:[dia("1",ymd(2026,7,24))], diaristas:[] });
  check(r.cortes.length === 0, "dia sem excesso não perde diária nenhuma");
}
{
  /* Pelo caminho completo: o plano corta a diária só depois de as quatro
     fases terem feito o que podiam no quadro fixo. */
  /* Quadro fixo EXATAMENTE no QF o período inteiro: nenhuma das quatro
     fases pode cortar ninguém sem furar outro dia. O excesso dos três
     dias com diária é, por construção, só diária. */
  const gente2 = [];
  for(let i=0;i<QF;i++) gente2.push({ groot:"95"+String(i).padStart(4,"0"),
    nome:"FIXO "+i, matricula:"M"+i, cargo:"Auxiliar de Apoio Log I", conta:"LABOR DIRETO",
    inicio:PER.ini, fim:null, rateio:1, linha:i+2 });
  const diarias = [];
  for(const d of [ymd(2026,7,20), ymd(2026,7,21), ymd(2026,8,1)])
    for(let k=0;k<5;k++) diarias.push({ groot:"92"+k, data:d, quantidade:1 });
  const r = simPlanoEqualizacao({ labor:gente2, periodo:PER, alvo:QF, diarias, opcoes:{} });
  check(r.acoes.length === 0,
        "o motor não mexe no quadro fixo quando cortar criaria falta em outro dia",
        JSON.stringify(r.acoes.map(a => a.tipo)));
  check(r.corte.totais.cortado === 15, "e o excesso sai inteiro em diária: 3 dias × 5",
        String(r.corte.totais.cortado));
  check(r.corte.cortes.every(c => diarias.some(d => d.groot === c.groot && d.data === c.data)),
        "…e toda diária cortada existia mesmo na fatura, naquele dia");
  check(r.dias.every(d => d.difPos <= 0),
        "depois do plano inteiro nenhum dia fica acima do QF",
        JSON.stringify(r.dias.filter(d => d.difPos > 0).map(d => d.data+":"+d.difPos)));
  check(r.totais.excessoDepois === 0, "e o excesso do período zera");
}

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
