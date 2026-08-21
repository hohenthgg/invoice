/* Cadastro — js/cadastro.js
   ================================================================

   O que estes testes protegem: a fronteira entre CORRIGIR e ESCOLHER.

   O módulo pode sugerir correção quando existe UM candidato plausível
   na base — nome que bate, filial que não desmente, CPF que não
   desmente. No instante em que há dois, a resposta certa é parar e
   perguntar: corrigir no palpite é como o erro entrou na planilha da
   primeira vez. Os testes cobrem os dois lados dessa fronteira, e as
   duas forças que a movem — a filial desempata, o CPF derruba.
   ================================================================ */
"use strict";
const { load } = require("./load.js");

const ctx = load(["dates.js","config.js","validacao.js","cadastro.js"], ["CAD_TIPO"]);
const { cadConferir, CAD_TIPO } = ctx;

let pass = 0, fail = 0;
function check(ok, label, extra){
  if(ok){ pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}
const secao = t => console.log("\n" + t);

const B = (groot, nome, extra) => ({ groot, nome, ...extra });
const P = (groot, nome, extra) => ({ groot, nome, origem:"labor", ...extra });
const rodar = (pessoas, base, extra) =>
  cadConferir({ pessoas, base, unidadeCod:"SMG9", unidadeNome:"VARGINHA", ...extra });

/* ================================================================
   O CASO QUE MOTIVOU O MÓDULO
   ================================================================ */
secao("O dígito engolido: 211618 → 2611618");
{
  const r = rodar(
    [P("211618","JOAO DA SILVA")],
    [B("2611618","JOAO DA SILVA",{ filiais:"SMG9", cpf:"11122233344" })]);
  const a = r.achados[0];
  check(a && a.tipo === CAD_TIPO.CORRIGIR, "GROOT que não existe na base, nome que existe: sugestão de correção");
  check(a.sugestao.groot === "2611618", "…apontando o GROOT da base", a && a.sugestao.groot);
  check(a.confianca === "alta", "nome idêntico + filial da fatura = confiança alta", a && a.confianca);
  check(r.totais.corrigir === 1 && r.totais.ok === 0, "e os totais contam a sugestão");
}
{
  const r = rodar(
    [P("2611618","JOAO DA SILVA")],
    [B("2611618","JOAO DA SILVA",{ filiais:"SMG9" })]);
  check(r.achados.length === 0 && r.totais.ok === 1,
        "GROOT e nome batendo com a base não geram achado nenhum");
}
{
  /* Variação de nome não impede o casamento — sobrenome a mais é a
     mesma pessoa, como na auditoria. */
  const r = rodar(
    [P("211618","JOAO DA SILVA")],
    [B("2611618","JOAO DA SILVA SANTOS",{ filiais:"SMG9" })]);
  check(r.achados[0] && r.achados[0].tipo === CAD_TIPO.CORRIGIR
        && r.achados[0].confianca === "media",
        "nome em variação ainda sugere, mas com confiança média",
        JSON.stringify(r.achados[0] && r.achados[0].confianca));
}
{
  /* Primeiro nome igual e o resto diferente é OUTRA pessoa. */
  const r = rodar(
    [P("211618","JOAO PEREIRA MOTA")],
    [B("2611618","JOAO DA SILVA",{ filiais:"SMG9" })]);
  check(r.achados[0] && r.achados[0].tipo === CAD_TIPO.FORA,
        "coincidência parcial de nome NÃO vira sugestão — sai como fora da base");
}

{
  /* O caso que apareceu no primeiro uso real: a fatura cheia de SILVAs
     e SANTOS. Nome repetindo sobrenome não pode arrastar meio arquivo
     para a lista de candidatos — Lorrany só casa com Lorrany. */
  const r = rodar(
    [P("2429028","LORRANY SILVA SANTOS SILVA")],
    [B("22429028","LORRANY SILVA SANTOS SILVA",{ filiais:"SMG9" }),
     B("2429120","WESLEI DOS SANTOS SILVA",{ filiais:"SMG9" }),
     B("2626644","MARCELA SILVA DOS SANTOS",{ filiais:"SMG9" }),
     B("2589084","JULIANA DA SILVA",{ filiais:"SMG9" })]);
  const a = r.achados[0];
  check(a && a.tipo === CAD_TIPO.CORRIGIR && a.sugestao.groot === "22429028",
        "homônimos parciais por sobrenome comum não viram candidatos — sobra o certo",
        JSON.stringify(a && (a.candidatos || a.sugestao)));
}

/* ================================================================
   AMBIGUIDADE É DECISÃO HUMANA
   ================================================================ */
secao("Dois candidatos: parar e perguntar");
{
  const r = rodar(
    [P("999","MARIA SOUZA")],
    [B("111","MARIA SOUZA",{ filiais:"SMG9" }), B("222","MARIA SOUZA",{ filiais:"SMG9" })]);
  const a = r.achados[0];
  check(a && a.tipo === CAD_TIPO.AMBIGUO, "dois candidatos plausíveis = ambíguo, nunca sugestão");
  check(a.candidatos.length === 2, "…com os dois candidatos listados para a escolha");
}
{
  /* A filial desempata: o homônimo de outra filial sai da disputa. */
  const r = rodar(
    [P("999","MARIA SOUZA")],
    [B("111","MARIA SOUZA",{ filiais:"SMG3" }), B("222","MARIA SOUZA",{ filiais:"SMG9" })]);
  const a = r.achados[0];
  check(a && a.tipo === CAD_TIPO.CORRIGIR && a.sugestao.groot === "222",
        "homônimo de outra filial sai da disputa — sobra um, vira sugestão",
        JSON.stringify(a && a.sugestao));
}
{
  /* …mas filial só desempata, não elimina o único candidato. */
  const r = rodar(
    [P("999","MARIA SOUZA")],
    [B("111","MARIA SOUZA",{ filiais:"SMG3" })]);
  check(r.achados[0] && r.achados[0].tipo === CAD_TIPO.CORRIGIR,
        "um único candidato de outra filial ainda é sugestão — a filial não inventa nem apaga");
}

/* ================================================================
   CPF: CONFIRMA FORTE, DERRUBA MAIS FORTE AINDA
   ================================================================ */
secao("CPF quando disponível");
{
  const r = rodar(
    [P("999","MARIA SOUZA",{ cpf:"22233344455" })],
    [B("111","MARIA SOUZA",{ filiais:"SMG9", cpf:"22233344455" }),
     B("222","MARIA SOUZA",{ filiais:"SMG9", cpf:"99988877766" })]);
  const a = r.achados[0];
  check(a && a.tipo === CAD_TIPO.CORRIGIR && a.sugestao.groot === "111" && a.confianca === "alta",
        "CPF igual resolve a ambiguidade e fecha em confiança alta",
        JSON.stringify(a && [a.tipo, a.sugestao && a.sugestao.groot]));
}
{
  const r = rodar(
    [P("999","MARIA SOUZA",{ cpf:"00011122233" })],
    [B("111","MARIA SOUZA",{ filiais:"SMG9", cpf:"22233344455" })]);
  check(r.achados[0] && r.achados[0].tipo === CAD_TIPO.FORA,
        "CPF divergente derruba o candidato por melhor que o nome pareça");
}

/* ================================================================
   OS OUTROS TIPOS
   ================================================================ */
secao("Preencher, conflito e fora da base");
{
  const r = rodar(
    [P("","CARLA DIAS")],
    [B("333","CARLA DIAS",{ filiais:"SMG9" })]);
  check(r.achados[0] && r.achados[0].tipo === CAD_TIPO.PREENCHER
        && r.achados[0].sugestao.groot === "333",
        "fatura sem GROOT e base com: sugestão de preencher");
}
{
  /* O GROOT digitado pertence a OUTRA pessoa: a sugestão diz a quem. */
  const r = rodar(
    [P("444","JOAO DA SILVA")],
    [B("444","PEDRO ALMEIDA COSTA",{ filiais:"SMG9" }),
     B("555","JOAO DA SILVA",{ filiais:"SMG9" })]);
  const a = r.achados[0];
  check(a && a.tipo === CAD_TIPO.CORRIGIR && a.sugestao.groot === "555",
        "GROOT de outra pessoa + nome que existe = correção para o GROOT certo");
  check(a.conflito && /PEDRO/.test(a.conflito.nome),
        "…dizendo de quem é o GROOT digitado", JSON.stringify(a && a.conflito));
}
{
  const r = rodar(
    [P("777","ROBERTO NOVATO ALMEIDA")],
    [B("333","CARLA DIAS",{ filiais:"SMG9" })]);
  check(r.achados[0] && r.achados[0].tipo === CAD_TIPO.FORA,
        "pessoa que não está na base sai como informação, sem correção inventada");
}
{
  const r = rodar([P("1","ABS"), P("2"," ABSENTEISMO")], [B("333","CARLA DIAS")]);
  check(r.achados.length === 0, "a linha de absenteísmo do template não é pessoa e não entra");
}
{
  const r = cadConferir({ pessoas:[P("999","MARIA SOUZA")], base:[] });
  check(r.achados.length === 0 && r.totais.corrigir === 0,
        "base vazia não produz achado — sem referência não há conferência");
}

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
