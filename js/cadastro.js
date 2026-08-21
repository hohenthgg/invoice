/* Cadastro — a fatura contra a base oficial de GROOT
   ================================================================

   Módulo PURO, sem DOM. Depende de js/validacao.js (valCompararNomes)
   carregado antes.

   A base oficial diz quem cada GROOT É:

       Groot ID | Nome | CPF | Filiais

   e a fatura diz quem ela COBROU. Quando os dois discordam, quase
   sempre é digitação — um dígito engolido transforma 2611618 em
   211618 — e o efeito não é cosmético: o GROOT errado quebra a
   deduplicação, o cruzamento com o SIGO e o retorno do cliente, tudo
   que identifica pessoa por esse número.

   O QUE ESTE MÓDULO DEVOLVE, E O QUE ELE SE RECUSA A FAZER

   Para cada pessoa da fatura cujo GROOT não bate com a base, ele
   procura pelo NOME — com o mesmo comparador da auditoria, que sabe
   separar variação de digitação de gente diferente — e usa a FILIAL
   como contexto: um homônimo de outra filial não é candidato. O CPF,
   quando a fatura o tiver, é confirmação forte: CPF igual fecha a
   questão, CPF diferente derruba o candidato por melhor que o nome
   pareça.

   Com UM candidato plausível, sai uma SUGESTÃO — nunca uma correção
   aplicada: quem confirma é o usuário. Com mais de um, sai AMBÍGUO,
   marcado para decisão humana, porque escolher no palpite é como o
   erro entrou na planilha da primeira vez. Sem nenhum, sai
   "fora da base", que é informação e não defeito: a base pode estar
   desatualizada.
   ================================================================ */
"use strict";

const CAD_TIPO = {
  CORRIGIR:  "corrigir",    // GROOT da fatura não é o da base para este nome
  PREENCHER: "preencher",   // fatura sem GROOT, base tem
  AMBIGUO:   "ambiguo",     // mais de um candidato plausível — decisão humana
  FORA:      "fora"         // pessoa não encontrada na base
};

const cadNorm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"")
  .toUpperCase().replace(/\s+/g," ").trim();
const cadGroot = v => String(v ?? "").replace(/\D/g,"");
const cadCpf = v => String(v ?? "").replace(/\D/g,"");

/* Nome bate? Mesmo juiz da auditoria: idêntico ou variação (sobrenome a
   mais, um caractere trocado). "Parcial" NÃO basta — primeiro nome igual
   com o resto diferente é outra pessoa até prova em contrário. */
function cadNomeBate(a, b){
  const r = valCompararNomes(a, b);
  return r.relacao === "iguais" || r.relacao === "variacao";
}

/* A filial da base ("SMG9", "SMG9, SMG3", "Varginha") contém a unidade
   da fatura? Aceita o código (SMG9) ou o nome (VARGINHA) — a base real
   pode trazer qualquer um dos dois. Filial vazia não elimina: ausência
   de dado não é dado. */
function cadFilialBate(filiais, unidadeCod, unidadeNome){
  const f = cadNorm(filiais);
  if(!f) return true;
  const cod = cadNorm(unidadeCod), nome = cadNorm(unidadeNome);
  if(!cod && !nome) return true;
  return (cod && f.includes(cod)) || (nome && f.includes(nome));
}

/**
 * Confere as pessoas da fatura contra a base oficial.
 *
 * @param {Object} dados
 * @param {Array<{groot,nome,cpf?,origem}>} dados.pessoas  uma entrada por
 *        pessoa distinta da fatura (LABOR e DIARISTAS), com a origem
 * @param {Array<{groot,nome,cpf?,filiais?}>} dados.base   a base oficial
 * @param {string} [dados.unidadeCod]   ex.: "SMG9"
 * @param {string} [dados.unidadeNome]  ex.: "VARGINHA"
 * @returns {{achados:Array, totais:Object}}
 */
function cadConferir(dados){
  const base = (dados.base || []).map(r => ({
    groot: cadGroot(r.groot), nome: String(r.nome ?? "").trim(),
    cpf: cadCpf(r.cpf), filiais: String(r.filiais ?? "").trim()
  })).filter(r => r.groot && r.nome);

  const porGroot = new Map(base.map(r => [r.groot, r]));

  const achados = [];
  const totais = { ok:0, corrigir:0, preencher:0, ambiguo:0, fora:0 };
  /* Sem base não há referência: devolver "fora da base" para todo mundo
     seria ruído vestido de achado. */
  if(!base.length) return { achados, totais };

  for(const p of (dados.pessoas || [])){
    const g = cadGroot(p.groot);
    const nome = String(p.nome ?? "").trim();
    if(!nome) continue;                       // sem nome não há o que casar
    /* A linha de ABS do template não é pessoa — nem pelo nome, nem
       pelo "GROOT" ABS. */
    if(/ABSENTEISMO/.test(cadNorm(nome)) || cadNorm(nome) === "ABS"
       || cadNorm(p.groot) === "ABS") continue;

    const daBase = g ? porGroot.get(g) : null;
    if(daBase && cadNomeBate(nome, daBase.nome)){ totais.ok++; continue; }

    /* O GROOT não resolveu — o nome é a chave. */
    let cand = base.filter(r => cadNomeBate(nome, r.nome));

    /* CPF, quando a fatura tem: igual confirma, diferente elimina. */
    const cpfFat = cadCpf(p.cpf);
    if(cpfFat){
      const exato = cand.filter(r => r.cpf && r.cpf === cpfFat);
      if(exato.length) cand = exato;
      else cand = cand.filter(r => !r.cpf);   // cpf divergente derruba
    }

    /* Filial como contexto: só desempata, nunca inventa. */
    if(cand.length > 1){
      const daFilial = cand.filter(r =>
        cadFilialBate(r.filiais, dados.unidadeCod, dados.unidadeNome));
      if(daFilial.length >= 1) cand = daFilial;
    }
    /* O mesmo GROOT digitado não pode ser candidato de si mesmo. */
    cand = cand.filter(r => r.groot !== g || !g);

    if(cand.length === 1){
      const c = cand[0];
      const confianca =
        (cpfFat && c.cpf === cpfFat) ? "alta"
        : (valCompararNomes(nome, c.nome).relacao === "iguais"
           && cadFilialBate(c.filiais, dados.unidadeCod, dados.unidadeNome)) ? "alta"
        : "media";
      const tipo = g ? CAD_TIPO.CORRIGIR : CAD_TIPO.PREENCHER;
      totais[tipo]++;
      achados.push({ tipo, nome, groot:g, origem:p.origem, confianca,
        sugestao:{ groot:c.groot, nome:c.nome, cpf:c.cpf, filiais:c.filiais },
        /* Se o GROOT digitado pertence a OUTRA pessoa na base, dizer a
           quem: é a diferença entre erro de digitação e troca de gente. */
        conflito: daBase ? { nome: daBase.nome } : null,
        motivo: g
          ? "O nome bate com a base, mas lá esse nome tem o GROOT "+c.groot
            + (daBase ? "; o GROOT "+g+" digitado pertence a "+daBase.nome : "")
            + "."
          : "A fatura não traz GROOT para este nome; na base ele existe com o GROOT "+c.groot+"." });
    } else if(cand.length > 1){
      totais.ambiguo++;
      achados.push({ tipo:CAD_TIPO.AMBIGUO, nome, groot:g, origem:p.origem,
        candidatos: cand.map(c => ({ groot:c.groot, nome:c.nome, cpf:c.cpf, filiais:c.filiais })),
        motivo: cand.length+" pessoas da base têm esse nome — corrigir no palpite é como o erro "
          + "entrou; a escolha é sua." });
    } else {
      totais.fora++;
      achados.push({ tipo:CAD_TIPO.FORA, nome, groot:g, origem:p.origem,
        motivo: "Nome não encontrado na base oficial"
          + (daBase ? "; o GROOT "+g+" pertence a "+daBase.nome+" na base" : "")
          + ". Pode ser admissão recente ou base desatualizada — nada a corrigir automaticamente." });
    }
  }

  /* Sugestões primeiro (são acionáveis), ambíguos depois, "fora" por último. */
  const peso = { corrigir:0, preencher:1, ambiguo:2, fora:3 };
  achados.sort((a,b) => peso[a.tipo]-peso[b.tipo] || a.nome.localeCompare(b.nome));
  return { achados, totais };
}

if(typeof module !== "undefined" && module.exports){
  module.exports = { cadConferir, cadNomeBate, cadFilialBate, CAD_TIPO, cadGroot };
}
