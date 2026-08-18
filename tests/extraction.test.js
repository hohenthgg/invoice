/* Testes da deduplicação da Extração · Diarista.

   O que estes testes provam:

       MESMO GROOT NORMALIZADO + MESMA DATA = UMA ÚNICA PESSOA-DIA

   independentemente de operação, aba ou arquivo de origem. Um `seen` por
   operação faria `SVC|123456|10/08` e `XD|123456|10/08` sobreviverem os dois,
   produzindo dupla contagem de diarista e compensação indevida de ABS. */
"use strict";
const { load } = require("./load");
const ctx = load(["identity.js", "config.js", "extraction-dedup.js", "extraction-audit.js"],
                 ["DEDUP_MODOS", "AUDIT_MOTIVO", "NOME_PROBLEMA", "GRAVIDADE",
                  "ESCALA_HORARIO_PADRAO"]);
const { normalizeGroot, limparGroot, grootParaSaida, hasGroot, normalizeNome,
        normalizeDateKey, personDayKey,
        deduplicarPessoaDia, DEDUP_MODOS, auditarIdentidade, AUDIT_MOTIVO,
        problemasDoNome, nomeQuebrado, NOME_PROBLEMA,
        listarTopicos, resumirTopicos, GRAVIDADE,
        ESCALA_HORARIO_PADRAO, escalaHorarioDe,
        pareceHorario, resolverEscala } = ctx;

let pass = 0, fail = 0;
function check(ok, label, extra) {
  if (ok) { pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}

/** registro de diarista, como a leitura do SIGO produz */
const reg = (id, date, nome) => ({ id, date, nome: nome === undefined ? "Diarista" : nome, solic: "id",
                                   empresa: "Empresa X", cargo: "Auxiliar", escala: "6x1",
                                   solicitante: "ID Logistics" });
/** roda a dedup sobre operações nomeadas */
function dedup(porOp, modo) {
  return deduplicarPessoaDia(Object.entries(porOp).map(([op, rows]) => ({ op, rows })),
                             modo || DEDUP_MODOS.DIA);
}
const totalMantido = r => Object.values(r.porOperacao).reduce((a, o) => a + o.rows.length, 0);

/* ================================================================ */
console.log("\nNormalização de GROOT");
check(normalizeGroot(123456) === "123456", "número vira texto canônico");
check(normalizeGroot("123456.0") === "123456", "float do Excel perde o .0");
check(normalizeGroot("123456.00") === "123456", "…e o .00 também");
check(normalizeGroot(" 123456 ") === "123456", "espaços em volta somem");
check(normalizeGroot("123456") === normalizeGroot(123456)
      && normalizeGroot(123456) === normalizeGroot("123456.0")
      && normalizeGroot("123456.0") === normalizeGroot(" 123456 "),
      "todas as grafias equivalentes colapsam no mesmo identificador");
check(normalizeGroot("ABC123") === "ABC123", "identificador alfanumérico não é convertido");

/* `{2499441}` — em Pouso Alegre XD, 12 pessoas apareciam como DUAS porque o
   mesmo GROOT vinha ora embrulhado, ora limpo, e a dedup obedecia. */
check(limparGroot("{2499441}") === "2499441" && limparGroot("[2499441]") === "2499441"
      && limparGroot("(2499441)") === "2499441" && limparGroot("/2499441/") === "2499441"
      && limparGroot("\\2499441\\") === "2499441" && limparGroot("<2499441>") === "2499441",
      "chaves, colchetes, parênteses, barras e sinais de maior/menor são embalagem",
      limparGroot("{2499441}"));
check(limparGroot("{[2499441]}") === "2499441", "…inclusive embrulho duplo");
check(normalizeGroot("{2499441}") === normalizeGroot("2499441"),
      "o embrulhado e o limpo passam a ser A MESMA pessoa na comparação");
check(normalizeGroot("{ 2499441 }") === "2499441", "espaço dentro do embrulho também sai");
check(limparGroot("AB-12/34") === "AB-12/34",
      "delimitador NO MEIO do identificador não é embalagem — pode ser parte do id",
      limparGroot("AB-12/34"));
check(limparGroot("{00123456}") === "00123456",
      "…e a limpeza não come zeros à esquerda");
check(grootParaSaida("{2499441}") === "2499441" && grootParaSaida("{abc12}") === "abc12",
      "a saída sai limpa, mas sem forçar caixa alta — é para ler, não para comparar",
      grootParaSaida("{abc12}"));
{
  const r = dedup({ XD: [reg("{2499441}", "2026-08-10"), reg("2499441", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "…e a mesma pessoa-dia embrulhada e limpa vira UM registro, não dois",
        `mantidos=${totalMantido(r)}`);
}
check(normalizeGroot("00123456") === "00123456",
      "zeros à esquerda são preservados — podem ser parte do identificador",
      normalizeGroot("00123456"));
check(normalizeGroot("00123456") !== normalizeGroot("123456"),
      "…e por isso não são unidos silenciosamente a outro id");
check(normalizeGroot("") === "" && normalizeGroot(null) === "" && normalizeGroot(undefined) === "",
      "vazio, nulo e indefinido não identificam ninguém");
check(!hasGroot("0") && !hasGroot("") && hasGroot("123456"),
      "'0' e vazio não contam como GROOT presente");

console.log("\nNormalização de data");
check(normalizeDateKey("10/08/2026") === "2026-08-10", "DD/MM/AAAA");
check(normalizeDateKey("2026-08-10") === "2026-08-10", "AAAA-MM-DD");
check(normalizeDateKey(new Date(Date.UTC(2026, 7, 10))) === "2026-08-10", "objeto Date (UTC)");
check(normalizeDateKey(46244) === "2026-08-10", "serial do Excel", normalizeDateKey(46244));
check(normalizeDateKey(20260810) === "2026-08-10", "inteiro AAAAMMDD interno");
check(personDayKey(123456, "10/08/2026") === "123456|2026-08-10",
      "a chave pessoa-dia é GROOT|AAAA-MM-DD", personDayKey(123456, "10/08/2026"));
check(personDayKey("", "10/08/2026") === null, "sem GROOT não existe chave pessoa-dia");

/* ================================================================ */
console.log("\nDeduplicação pessoa-dia");

// Teste 1 — mesmo GROOT, mesma data, mesma operação
{
  const r = dedup({ SVC: [reg("123456", "2026-08-10"), reg("123456", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "1 · mesmo GROOT, mesma data, mesma operação → 1 registro",
        `mantidos=${totalMantido(r)} dups=${r.resumo.duplicados}`);
}

// Teste 2 — mesmo GROOT, mesma data, OPERAÇÕES DIFERENTES  (o defeito corrigido)
{
  const r = dedup({ SVC: [reg("123456", "2026-08-10")],
                    XD:  [reg("123456", "2026-08-10")],
                    SD:  [reg("123456", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 2,
        "2 · mesmo GROOT e data em SVC, XD e SD → 1 registro válido, 2 duplicados",
        `mantidos=${totalMantido(r)} dups=${r.resumo.duplicados}`);
  check(r.porOperacao.SVC.rows.length === 1 && r.porOperacao.XD.rows.length === 0
        && r.porOperacao.SD.rows.length === 0,
        "…e quem fica é a PRIMEIRA ocorrência encontrada");
  check(r.descartados.length === 2 && r.descartados[0].mantidaEm === "SVC"
        && r.descartados[0].descartadaDe === "XD",
        "…com registro de onde veio a descartada e onde a mantida ficou",
        JSON.stringify(r.descartados[0]));
}

// Teste 3 — mesmo GROOT, datas diferentes
{
  const r = dedup({ SVC: [reg("123456", "2026-08-10"), reg("123456", "2026-08-11")] });
  check(totalMantido(r) === 2 && r.resumo.duplicados === 0,
        "3 · mesmo GROOT em datas diferentes → 2 registros (duas diárias)");
}

// Teste 4 — GROOTs diferentes, mesma data
{
  const r = dedup({ SVC: [reg("123456", "2026-08-10"), reg("654321", "2026-08-10")] });
  check(totalMantido(r) === 2 && r.resumo.duplicados === 0,
        "4 · GROOTs diferentes na mesma data → 2 registros");
}

// Teste 5 — 123456 e 123456.0
{
  const r = dedup({ SVC: [reg(123456, "2026-08-10")], XD: [reg("123456.0", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "5 · 123456 e 123456.0 na mesma data → duplicado");
}

// Teste 6 — " 123456 " e 123456
{
  const r = dedup({ SVC: [reg(" 123456 ", "2026-08-10")], XD: [reg(123456, "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "6 · ' 123456 ' e 123456 na mesma data → duplicado");
}

// Teste 7 — mesma pessoa-dia vinda de arquivos diferentes
{
  // dois arquivos carregados na mesma extração chegam como operações distintas
  const r = dedup({ "Arquivo A · SVC": [reg("123456", "2026-08-10")],
                    "Arquivo B · SVC": [reg("123456", "2026-08-10")] });
  check(totalMantido(r) === 1 && r.resumo.duplicados === 1,
        "7 · mesma pessoa-dia em arquivos diferentes → 1 único registro");
}

// Teste 8 — GROOT vazio nunca é deduplicado
{
  const r = dedup({ SVC: [reg("", "2026-08-10", "Fulano"), reg("", "2026-08-10", "Beltrano"),
                          reg(null, "2026-08-10", "Sicrano")] });
  check(totalMantido(r) === 3 && r.resumo.duplicados === 0,
        "8 · três registros sem GROOT na mesma data → nenhum é removido",
        `mantidos=${totalMantido(r)}`);
  check(r.resumo.semGroot === 3 && r.semGroot.length === 3,
        "…são contados e marcados para revisão", `semGroot=${r.resumo.semGroot}`);

  /* Uma contagem não é revisável. Quem revisa precisa achar a pessoa na
     origem para preencher o GROOT, e para isso precisa do nome, da filial e
     da data — não de "119 registros". */
  check(r.semGroot.every(s => s.op === "SVC" && s.date === "2026-08-10"),
        "…com a filial e a data de cada um", JSON.stringify(r.semGroot[0]));
  check(r.semGroot.map(s => s.nome).join("|") === "Fulano|Beltrano|Sicrano",
        "…e o nome de cada pessoa, na ordem do arquivo",
        r.semGroot.map(s => s.nome).join("|"));
  check(r.semGroot.every(s => s.reg && s.reg.empresa === "Empresa X"
                              && s.reg.cargo === "Auxiliar" && s.reg.escala === "6x1"
                              && s.reg.solicitante === "ID Logistics"),
        "…e o registro inteiro, para exibir empresa, cargo, escala e solicitante",
        JSON.stringify(r.semGroot[0] && r.semGroot[0].reg));
  check(r.semGroot[0].reg === r.porOperacao.SVC.rows[0],
        "…apontando para o MESMO registro que ficou na saída, não para uma cópia");
}
{
  // filiais diferentes: o painel precisa dizer onde o problema está concentrado
  const r = dedup({ SVC: [reg("", "2026-08-10", "Fulano")],
                    XD:  [reg("", "2026-08-10", "Beltrano"), reg(null, "2026-08-11", "Sicrano")] });
  const porOp = r.semGroot.reduce((m, s) => (m[s.op] = (m[s.op] || 0) + 1, m), {});
  check(porOp.SVC === 1 && porOp.XD === 2,
        "…e dá para agrupar por filial a partir da lista", JSON.stringify(porOp));
}

/* ================================================================ */
console.log("\nModos e transparência");
{
  const entrada = { SVC: [reg("1", "2026-08-10"), reg("1", "2026-08-11")],
                    XD:  [reg("1", "2026-08-10")] };
  const porDia = dedup(entrada, DEDUP_MODOS.DIA);
  check(totalMantido(porDia) === 2,
        "modo 'mesma data': 10/08 fica uma vez, 11/08 é outra diária", totalMantido(porDia));

  const porPeriodo = dedup(entrada, DEDUP_MODOS.PERIODO);
  check(totalMantido(porPeriodo) === 1,
        "modo 'período': uma linha por pessoa no período inteiro, entre operações",
        totalMantido(porPeriodo));

  const semDedup = dedup(entrada, DEDUP_MODOS.NAO);
  check(totalMantido(semDedup) === 3 && semDedup.resumo.duplicados === 0,
        "modo 'não remover': tudo permanece como está no arquivo");
}
{
  const r = dedup({ SVC: [reg("1", "2026-08-10"), reg("1", "2026-08-10")],
                    XD:  [reg("", "2026-08-10")] });
  check(r.resumo.encontrados === 3 && r.resumo.unicos === 2
        && r.resumo.duplicados === 1 && r.resumo.semGroot === 1,
        "o resumo bate: encontrados = únicos + duplicados",
        JSON.stringify(r.resumo));
  check(r.resumo.global === true, "o resumo declara que a deduplicação é global");
}
{
  // data ilegível não é motivo para descartar ninguém
  const r = dedup({ SVC: [reg("123456", "data ruim"), reg("123456", "data ruim")] });
  check(totalMantido(r) === 2 && r.resumo.duplicados === 0,
        "sem data legível não há pessoa-dia — os registros são preservados");
}

/* ================================================================
   Auditoria de identidade — o identificador e o nome contam a mesma
   história? A dedup responde "é a mesma pessoa-dia?"; isto responde a
   pergunta anterior, e só APONTA: nada é unido nem separado sozinho. */
console.log("\nAuditoria de identidade");

const auditar = porOp => auditarIdentidade(
  Object.entries(porOp).map(([op, rows]) => ({ op, rows })));

console.log("\n  Normalização de nome (só para comparar, nunca para identificar)");
check(normalizeNome("José  da Silva") === "JOSE DA SILVA", "acento, caixa e espaço duplo somem",
      normalizeNome("José  da Silva"));
check(normalizeNome("MARIA D'ÁVILA") === "MARIA D AVILA", "pontuação vira separador",
      normalizeNome("MARIA D'ÁVILA"));
check(normalizeNome(null) === "" && normalizeNome(undefined) === "", "nulo e indefinido viram vazio");
/* A aba de Pouso XD numera os nomes por posição na lista. Sem tirar isso,
   "1. Douglas" e "5. Douglas" viravam duas pessoas sob o mesmo GROOT. */
check(normalizeNome("1. Douglas David da Silva") === "DOUGLAS DAVID DA SILVA"
      && normalizeNome("1. Douglas David da Silva") === normalizeNome("5. Douglas David da Silva")
      && normalizeNome("23 - Maria José") === "MARIA JOSE",
      "numeração de lista na frente do nome é posição, não identidade",
      normalizeNome("1. Douglas David da Silva"));
check(problemasDoNome("1. Douglas David da Silva").length === 0,
      "…e não marca o nome como quebrado por 'ter número'",
      JSON.stringify(problemasDoNome("1. Douglas David da Silva")));
check(problemasDoNome("JOAO SILVA 2").includes(NOME_PROBLEMA.DIGITOS),
      "…mas número NO MEIO do nome continua sendo defeito");

console.log("\n  Mesmo nome, identificadores diferentes");
{
  const r = auditar({ SVC: [reg("123456", "2026-08-10", "ANA SOUZA"),
                            reg("999999", "2026-08-11", "Ana  Souza")] });
  check(r.mesmoNome.length === 1 && r.mesmoGroot.length === 0,
        "grafias diferentes do mesmo nome com GROOTs distintos viram um caso",
        `mesmoNome=${r.mesmoNome.length} mesmoGroot=${r.mesmoGroot.length}`);
  check(r.mesmoNome[0].variantes.length === 2 && r.mesmoNome[0].motivo === AUDIT_MOTIVO.IDS,
        "…com as duas variantes e o diagnóstico de ids realmente distintos",
        JSON.stringify(r.mesmoNome[0].motivo));
  /* O que varia sob um mesmo nome é o GROOT. Rotular a variante com o nome
     repetiria a mesma informação em toda linha e esconderia o conflito. */
  check(r.mesmoNome[0].rotulo === "ANA SOUZA"
        && r.mesmoNome[0].variantes.map(v => v.rotulo).sort().join("|") === "123456|999999",
        "…e a variante é rotulada pelo GROOT, não pelo nome do grupo",
        r.mesmoNome[0].variantes.map(v => v.rotulo).join("|"));
  check(r.mesmoNome[0].variantes[0].ocorrencias[0].op === "SVC"
        && !!r.mesmoNome[0].variantes[0].ocorrencias[0].date,
        "…e cada variante carrega onde e quando apareceu");
}
{
  const r = auditar({ SVC: [reg("00123456", "2026-08-10", "ANA SOUZA")],
                      XD:  [reg("123456",   "2026-08-11", "ANA SOUZA")] });
  check(r.mesmoNome.length === 1 && r.mesmoNome[0].motivo === AUDIT_MOTIVO.ZEROS,
        "GROOTs que só diferem por zeros à esquerda são diagnosticados como formatação",
        JSON.stringify(r.mesmoNome[0] && r.mesmoNome[0].motivo));
}
{
  const r = auditar({ SVC: [reg("123456", "2026-08-10", "ANA SOUZA"),
                            reg("123456", "2026-08-11", "ANA SOUZA")] });
  check(r.mesmoNome.length === 0, "mesmo nome com o MESMO GROOT não é conflito nenhum");
}

console.log("\n  Mesmo identificador, nomes diferentes");
{
  const r = auditar({ SVC: [reg("123456", "2026-08-10", "ANA SOUZA")],
                      XD:  [reg("123456", "2026-08-11", "CARLOS PEREIRA")] });
  check(r.mesmoGroot.length === 1 && r.mesmoGroot[0].motivo === AUDIT_MOTIVO.NOMES,
        "duas pessoas sob o mesmo GROOT são apontadas como o caso grave",
        JSON.stringify(r.mesmoGroot[0] && r.mesmoGroot[0].motivo));
  check(r.mesmoGroot[0].rotulo === "123456"
        && r.mesmoGroot[0].variantes.map(v => v.rotulo).sort().join("|") === "ANA SOUZA|CARLOS PEREIRA",
        "…identificadas pelo nome como veio da planilha, não pela chave normalizada",
        r.mesmoGroot[0].variantes.map(v => v.rotulo).join("|"));
  check(r.resumo.mesmoGrootNomesDistintos === 1,
        "…e contadas à parte no resumo, separadas das divergências de grafia");
}
{
  const r = auditar({ SVC: [reg("123456", "2026-08-10", "ANA MARIA SOUZA")],
                      XD:  [reg("123456", "2026-08-11", "ANA SOUZA")] });
  check(r.mesmoGroot.length === 1 && r.mesmoGroot[0].motivo === AUDIT_MOTIVO.ABREVIACAO,
        "nome contido no outro é diagnosticado como abreviação, não como duas pessoas",
        JSON.stringify(r.mesmoGroot[0] && r.mesmoGroot[0].motivo));
  check(r.resumo.mesmoGrootNomesDistintos === 0, "…e não infla a contagem do caso grave");
}
{
  const r = auditar({ SVC: [reg("123456", "2026-08-10", "SOUZA ANA")],
                      XD:  [reg("123456", "2026-08-11", "ANA SOUZA")] });
  check(r.mesmoGroot[0].motivo === AUDIT_MOTIVO.ORDEM,
        "mesmos nomes em ordem trocada são diagnosticados como grafia",
        JSON.stringify(r.mesmoGroot[0].motivo));
}
{
  // nome vazio não contradiz nome nenhum — senão todo registro incompleto viraria conflito
  const r = auditar({ SVC: [reg("123456", "2026-08-10", "ANA SOUZA"),
                            reg("123456", "2026-08-11", "")] });
  check(r.mesmoGroot.length === 0, "nome em branco não é um 'nome diferente'");
}
{
  // registro sem GROOT não entra na comparação, mas entra na lista de revisão
  const r = auditar({ SVC: [reg("", "2026-08-10", "ANA SOUZA"), reg("", "2026-08-11", "ANA SOUZA")] });
  check(r.mesmoNome.length === 0 && r.semGroot.length === 2,
        "sem identificador não há o que confrontar — vai para a lista de revisão");
  check(r.semGroot[0].nome === "ANA SOUZA" && r.semGroot[0].op === "SVC"
        && r.semGroot[0].empresa === "Empresa X" && r.semGroot[0].cargo === "Auxiliar"
        && r.semGroot[0].escala === "6x1" && r.semGroot[0].solicitante === "ID Logistics",
        "…com todos os campos que permitem achar a pessoa na origem",
        JSON.stringify(r.semGroot[0]));
  check(r.semGroot[0].grootBruto === "", "…e o identificador exatamente como estava na célula");
}
{
  const r = auditar({ SVC: [reg(0, "2026-08-10", "ANA SOUZA")] });
  check(r.semGroot.length === 1 && r.semGroot[0].grootBruto === "0",
        "GROOT '0' não identifica ninguém, mas o relatório mostra que era '0'",
        JSON.stringify(r.semGroot[0] && r.semGroot[0].grootBruto));
}

console.log("\n  A auditoria não muda nada");
{
  const rows = [reg("123456", "2026-08-10", "ANA SOUZA"), reg("999999", "2026-08-10", "ANA SOUZA")];
  const antes = JSON.stringify(rows);
  const r = auditar({ SVC: rows });
  check(JSON.stringify(rows) === antes && r.mesmoNome.length === 1,
        "os registros de entrada saem intactos — a auditoria só aponta");
  const d = dedup({ SVC: rows });
  check(totalMantido(d) === 2,
        "…e a deduplicação segue tratando GROOTs distintos como pessoas distintas");
}

/* ================================================================
   Nome quebrado — a célula não está vazia, mas também não identifica
   ninguém. É o defeito que passa despercebido justamente por parecer
   preenchido. */
console.log("\nNome quebrado");
const temProblema = (n, p) => problemasDoNome(n).includes(p);

check(temProblema("", NOME_PROBLEMA.VAZIO) && temProblema("   ", NOME_PROBLEMA.VAZIO),
      "célula vazia ou só com espaços");
check(temProblema("#N/D", NOME_PROBLEMA.ERRO_PLANILHA)
      && temProblema("#REF!", NOME_PROBLEMA.ERRO_PLANILHA)
      && temProblema("#VALOR!", NOME_PROBLEMA.ERRO_PLANILHA),
      "erro do Excel no lugar do nome (#N/D, #REF!, #VALOR!)");
check(temProblema("JoÃ£o Silva", NOME_PROBLEMA.CODIFICACAO),
      "acentuação corrompida", JSON.stringify(problemasDoNome("JoÃ£o Silva")));
check(temProblema("A DEFINIR", NOME_PROBLEMA.PLACEHOLDER)
      && temProblema("sem nome", NOME_PROBLEMA.PLACEHOLDER)
      && temProblema("XXXXX", NOME_PROBLEMA.PLACEHOLDER),
      "texto genérico no lugar do nome");
check(temProblema("JOAO SILVA 2", NOME_PROBLEMA.DIGITOS), "números no meio do nome");
check(temProblema("AB", NOME_PROBLEMA.CURTO), "curto demais");
check(temProblema("AAAAA BBBB", NOME_PROBLEMA.REPETICAO), "caractere repetido em sequência");
check(problemasDoNome("JOSÉ CARLOS DA SILVA").length === 0, "nome normal não tem problema nenhum",
      JSON.stringify(problemasDoNome("JOSÉ CARLOS DA SILVA")));
check(problemasDoNome("D'ÁVILA SOUZA-LIMA").length === 0,
      "apóstrofo e hífen são parte de nome de verdade",
      JSON.stringify(problemasDoNome("D'ÁVILA SOUZA-LIMA")));
check(problemasDoNome("MARIA").join() === NOME_PROBLEMA.UM_TOKEN && !nomeQuebrado("MARIA"),
      "nome de uma palavra só é aviso, não defeito — existe gente cadastrada assim");
check(nomeQuebrado("#N/D") && nomeQuebrado("") && !nomeQuebrado("ANA SOUZA"),
      "só o que impede identificar a pessoa conta como quebrado");
{
  const r = auditar({ SVC: [reg("123456", "2026-08-10", "ANA SOUZA")],
                      XD:  [reg("123456", "2026-08-11", "#N/D")] });
  check(r.nomesQuebrados.length === 1 && r.mesmoGroot.length === 0,
        "nome quebrado sai da comparação: '#N/D' não é 'outra pessoa' sob o mesmo GROOT",
        `quebrados=${r.nomesQuebrados.length} mesmoGroot=${r.mesmoGroot.length}`);
  check(r.nomesQuebrados[0].problemasNome.includes(NOME_PROBLEMA.ERRO_PLANILHA),
        "…e o motivo fica registrado no próprio registro");
}

/* ================================================================
   Tópicos — o formato único que alimenta os balões da tela e as abas
   do Excel. Se divergirem, a tela e a planilha contam histórias
   diferentes do mesmo arquivo. */
console.log("\nTópicos");
{
  const r = auditar({
    SVC: [reg("", "2026-08-10", "FULANO SEM ID"),
          reg("555001", "2026-08-11", "RENATO GOMES"),
          reg("555001", "2026-08-12", "LUCIANA DIAS"),
          reg("777002", "2026-08-13", "#N/D")]
  });
  const desc = [{groot:"123456", data:"2026-08-14", nome:"ANA SOUZA",
                 mantidaEm:"SVC", descartadaDe:"XD"}];
  const leitura = {abas:[{op:"Varginha", motivo:"Aba não encontrada no arquivo."}],
                   semData:[{op:"SVC", linha:42, valor:"31/02", nome:"BENEDITA LIMA", groot:"9"}]};
  const t = listarTopicos(r, desc, leitura);

  /* 7 = 1 aba + 1 linha sem data + 1 nome quebrado + 1 sem identificador
         + 2 ocorrências do conflito de GROOT + 1 duplicado.
     O conflito rende UMA LINHA POR OCORRÊNCIA, não uma por caso: na planilha
     é preciso enxergar cada registro afetado, não só que o caso existe. */
  check(t.length === 7, "uma linha por registro afetado, sem perder nem duplicar nenhum",
        `${t.length}: ` + t.map(x => x.topico).join(" | "));
  check(t.filter(x => x.topico === "Mesmo GROOT, nomes diferentes")
         .map(x => x.nome).sort().join("|") === "LUCIANA DIAS|RENATO GOMES",
        "…e as duas pessoas do conflito aparecem nominalmente");
  check(t.every(x => x.topico && x.gravidade && x.diagnostico && x.acao),
        "toda linha diz o que é, quão grave é, o que aconteceu e o que fazer");
  check(t[0].topico === "Aba não lida" && t[0].gravidade === GRAVIDADE.ALTA,
        "a aba não lida vem primeiro — derruba a filial inteira e nenhum outro tópico a enxerga",
        t[0].topico);
  check(t.findIndex(x => x.topico === "Mesmo GROOT, nomes diferentes") <
        t.findIndex(x => x.topico === "Duplicado removido"),
        "o grave vem antes do informativo");

  const b = resumirTopicos(t);
  check(b.reduce((a, x) => a + x.total, 0) === t.length,
        "os balões somam exatamente os problemas — nenhum fica fora da contagem",
        JSON.stringify(b.map(x => x.topico + ":" + x.total)));
  check(b[0].gravidade === GRAVIDADE.ALTA && b[b.length-1].gravidade === GRAVIDADE.INFO,
        "…e vêm ordenados de quem precisa de atenção antes para quem é só informativo");
  check(new Set(b.map(x => x.topico)).size === b.length, "um balão por tópico, sem repetição");

  const semData = t.find(x => x.topico === "Linha sem data legível");
  check(/linha 42/.test(semData.diagnostico) && /31\/02/.test(semData.diagnostico)
        && semData.nome === "BENEDITA LIMA",
        "a linha sem data legível diz qual linha é, o que estava na célula e de quem era",
        semData.diagnostico);
}
{
  // com a deduplicação desligada não há duplicados, mas o resto continua
  const r = auditar({ SVC: [reg("", "2026-08-10", "FULANO")] });
  const t = listarTopicos(r, [], null);
  check(t.length === 1 && t[0].topico === "Sem identificador",
        "sem duplicados e sem problemas de leitura, sobra só o que a auditoria achou");
  check(resumirTopicos([]).length === 0, "nenhum problema, nenhum balão");
}

/* ================================================================
   Escala horário — o SIGO não informa horário. Ele vem do padrão da
   operação, levantado na aba DIARISTAS das faturas 3PL de julho/2026. */
console.log("\nEscala horário por operação");
const OPS = ["Pouso Alegre SVC","Pouso Alegre XD","Poços de Caldas",
             "Varginha","Divinópolis","Patos de Minas"];
check(OPS.every(op => op in ESCALA_HORARIO_PADRAO),
      "toda operação da extração tem entrada na tabela — nenhuma cai no esquecimento",
      OPS.filter(op => !(op in ESCALA_HORARIO_PADRAO)).join());
check(escalaHorarioDe("Pouso Alegre SVC") === "03:00 07:00 08:00 12:48"
      && escalaHorarioDe("Pouso Alegre XD") === "13:00 17:00 18:00 22:45",
      "as duas unidades de Pouso Alegre têm horários diferentes (SMG3 madrugada, BRXMG3 tarde)",
      escalaHorarioDe("Pouso Alegre XD"));
check(escalaHorarioDe("Poços de Caldas") === "03:00 07:00 08:00 12:48"
      && escalaHorarioDe("Varginha") === "03:00 07:00 08:00 11:20"
      && escalaHorarioDe("Divinópolis") === "01:30 05:00 06:00 09:50"
      && escalaHorarioDe("Patos de Minas") === "00:00 05:00 06:00 09:20",
      "…e cada uma das demais tem o horário da sua fatura");
check(OPS.every(op => escalaHorarioDe(op) !== ""),
      "as seis operações estão levantadas — nenhuma sai em branco",
      OPS.filter(op => !escalaHorarioDe(op)).join());
check(escalaHorarioDe("Varginha") !== escalaHorarioDe("Poços de Caldas"),
      "Varginha tem horário próprio (11:20), não o de Poços de Caldas (12:48) — "
      + "os dois começam 03:00 e é fácil confundir",
      escalaHorarioDe("Varginha"));
check(escalaHorarioDe("Filial Inexistente") === "",
      "operação desconhecida não inventa horário");
check(OPS.filter(op => escalaHorarioDe(op))
        .every(op => /^\d{2}:\d{2}( \d{2}:\d{2}){3}$/.test(escalaHorarioDe(op))),
      "todos no formato da fatura: entrada, início e fim do intervalo, saída");
/* A coluna ESCALA do SIGO fala uma língua por filial: Varginha diz AM/PM,
   Divinópolis escreve o horário inteiro, Patos usa "00:30 as 09:18", Pouso
   usa svc/xd. A saída precisa refletir os horários fielmente. */
console.log("\n  Resolução da escala por linha");
check(pareceHorario("01:00 04:00 05:00 09:20") && pareceHorario("00:30 as 09:18")
      && !pareceHorario("AM") && !pareceHorario("XD") && !pareceHorario(""),
      "horário escrito é reconhecido nos dois formatos; turno e vazio não");
{
  const r = resolverEscala("Varginha", "AM", "");
  const s = resolverEscala("Varginha", "PM", "");
  check(r.valor === "03:00 07:00 08:00 11:20" && s.valor === "10:00 15:00 16:00 19:48"
        && r.origem === "token" && s.origem === "token",
        "Varginha: AM e PM viram os dois horários da fatura SMG9",
        r.valor + " · " + s.valor);
  check(resolverEscala("Varginha", "pm", "").valor === "10:00 15:00 16:00 19:48",
        "…sem depender de caixa (pm = PM)");
}
check(resolverEscala("Pouso Alegre SVC", "svc", "").valor === "03:00 07:00 08:00 12:48"
      && resolverEscala("Pouso Alegre SVC", "XD", "").valor === "13:00 17:00 18:00 22:45",
      "Pouso SVC: svc e XD viram os horários das faturas SMG3 e BRXMG3 — a aba mistura os dois turnos");
check(resolverEscala("Divinópolis", "01:00 04:00 05:00 09:20", "").valor === "01:00 04:00 05:00 09:20"
      && resolverEscala("Divinópolis", "01:00 04:00 05:00 09:20", "").origem === "arquivo",
      "Divinópolis: horário escrito na linha passa verbatim, sem conversão");
check(resolverEscala("Patos de Minas", "00:30 as 09:18", "").valor === "00:30 as 09:18",
      "Patos: o formato próprio ('00:30 as 09:18') também passa como está");
check(resolverEscala("Patos de Minas", "", "").valor === "00:00 05:00 06:00 09:20"
      && resolverEscala("Patos de Minas", "", "").origem === "padrao",
      "…e a célula vazia recebe o padrão da operação");
{
  const r = resolverEscala("Poços de Caldas", "SD", "");
  check(r.valor === "SD" && r.origem === "sem_mapa",
        "turno sem horário levantado (SD) fica como está — escrever o horário de outro "
        + "turno seria pôr dado errado com cara de certo", JSON.stringify(r));
}
check(resolverEscala("Varginha", "AM", "07:00 11:00 12:00 16:20").valor === "07:00 11:00 12:00 16:20",
      "a coluna explícita de horário, quando preenchida, manda sobre tudo");

{
  const r = auditar({ SVC: [reg("1", "2026-08-10", "ANA SOUZA")] });
  const t = listarTopicos(r, [], null, [],
    [{op:"Poços de Caldas", escalasSemMapa:[{token:"SD", vezes:17}], vazias:[], preenchidas:[]}]);
  check(t.length === 1 && t[0].topico === "Escala sem horário mapeado"
        && /"SD"/.test(t[0].diagnostico) && /17 registros/.test(t[0].diagnostico),
        "o turno sem mapa vira tópico, dizendo o token e quantos registros saíram sem conversão",
        JSON.stringify(t[0] && t[0].diagnostico));
}

{
  const r = auditar({ SVC: [reg("1", "2026-08-10", "ANA SOUZA")] });
  const t = listarTopicos(r, [], null, [{op:"Varginha", registros:12}]);
  check(t.length === 1 && t[0].topico === "Sem escala horário"
        && t[0].op === "Varginha" && /12 registros/.test(t[0].diagnostico),
        "a filial sem horário vira tópico de revisão, com quantos registros ficam em branco",
        JSON.stringify(t[0] && [t[0].topico, t[0].op]));
  check(listarTopicos(r, [], null, []).length === 0,
        "…e não aparece quando todas as filiais têm horário");
}

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
