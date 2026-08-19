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

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
