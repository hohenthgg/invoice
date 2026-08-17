/* Extração · Diarista — auditoria de identidade
   ================================================================

   A deduplicação responde "estes dois registros são a mesma pessoa-dia?".
   Esta auditoria responde a pergunta anterior: "o identificador e o nome
   estão contando a mesma história?".

   Três coisas ficam de fora da deduplicação e precisam de olho humano:

     1. SEM IDENTIFICADOR — a pessoa existe, o GROOT não. Nada pode ser
        deduplicado, e a lista é o caminho para preencher na origem.

     2. MESMO NOME, GROOTs DIFERENTES — ou são homônimos (acontece), ou a
        mesma pessoa foi cadastrada duas vezes, ou alguém digitou o id
        errado. A dedup não uniu — e não deveria unir sozinha: nome não é
        chave.

     3. MESMO GROOT, NOMES DIFERENTES — o mais grave. A dedup UNIU esses
        registros por serem o mesmo identificador. Se os nomes divergem de
        verdade, uma diária pode ter sido apagada como se fosse repetida.

   O módulo só APONTA. Nada é unido, separado ou removido a partir daqui:
   sem outra fonte, o dado não diz qual das versões está certa.

   Módulo puro, sem DOM.
   ================================================================ */
"use strict";

/* Relações possíveis entre as variantes de um mesmo grupo. O rótulo importa
   tanto quanto a detecção: "provável formatação" e "duas pessoas diferentes"
   pedem ações opostas de quem revisa. */
const AUDIT_MOTIVO = {
  ZEROS:      "zeros",        // GROOTs que só diferem por zeros à esquerda
  IDS:        "ids",          // GROOTs realmente distintos
  ORDEM:      "ordem",        // mesmos nomes, ordem trocada
  ABREVIACAO: "abreviacao",   // um nome está contido no outro
  NOMES:      "nomes"         // nomes de fato diferentes
};

const AUDIT_MOTIVO_TXT = {
  zeros:      "Os identificadores só diferem por zeros à esquerda — quase certamente formatação, não duas pessoas. Padronize na origem.",
  ids:        "Identificadores realmente distintos para o mesmo nome — ou são homônimos, ou a pessoa foi cadastrada duas vezes, ou um lançamento saiu com o id de outra pessoa.",
  ordem:      "Os mesmos nomes em ordem diferente — provável grafia, mesma pessoa.",
  abreviacao: "Um nome está contido no outro (abreviado ou incompleto) — provável mesma pessoa.",
  nomes:      "Nomes de fato diferentes sob o mesmo identificador. Atenção: a deduplicação UNIU esses registros; se forem pessoas diferentes, uma diária pode ter sido descartada como repetida."
};

/** tokens comparáveis de um nome */
function tokensNome(v){ return normalizeNome(v).split(" ").filter(Boolean); }

/** `00123` e `123` são o mesmo número escrito diferente. */
function semZerosAEsquerda(s){ return String(s).replace(/^0+(?=.)/, ""); }

/** Por que dois ou mais GROOTs aparecem sob o mesmo nome. */
function relacaoEntreGroots(groots){
  const sem = new Set(groots.map(semZerosAEsquerda));
  return sem.size === 1 ? AUDIT_MOTIVO.ZEROS : AUDIT_MOTIVO.IDS;
}

/** Por que dois ou mais nomes aparecem sob o mesmo GROOT. */
function relacaoEntreNomes(nomes){
  const toks = nomes.map(tokensNome);
  if (new Set(toks.map(t => t.slice().sort().join(" "))).size === 1) return AUDIT_MOTIVO.ORDEM;
  for (let i = 0; i < toks.length; i++){
    for (let j = i + 1; j < toks.length; j++){
      const a = new Set(toks[i]), b = new Set(toks[j]);
      const aEmB = toks[i].every(t => b.has(t)), bEmA = toks[j].every(t => a.has(t));
      if (!aEmB && !bEmA) return AUDIT_MOTIVO.NOMES;
    }
  }
  return AUDIT_MOTIVO.ABREVIACAO;
}

/** Achata as entradas em registros comparáveis, já normalizados. */
function planificar(entradas){
  const out = [];
  (entradas || []).forEach(({op, rows}) => {
    (rows || []).forEach(r => {
      out.push({
        op, date: r.date,
        grootBruto: (r.id === null || r.id === undefined) ? "" : String(r.id),
        groot: hasGroot(r.id) ? normalizeGroot(r.id) : "",
        nome: String(r.nome == null ? "" : r.nome).trim(),
        nomeChave: normalizeNome(r.nome),
        empresa: r.empresa || "", cargo: r.cargo || "",
        escala: r.escala || "", solicitante: r.solicitante || ""
      });
    });
  });
  return out;
}

function agrupar(rows, chaveDe){
  const m = new Map();
  rows.forEach(r => {
    const k = chaveDe(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  });
  return m;
}

/** Monta os grupos conflitantes de uma das duas direções.
 *
 *  @param chaveDe          o que é IGUAL dentro do grupo
 *  @param varianteDe       o que DIVERGE dentro do grupo
 *  @param grupoRotuloDe    como o grupo se apresenta a quem lê
 *  @param varianteRotuloDe como cada variante se apresenta — precisa descrever
 *         a VARIANTE, não o grupo: sob um mesmo nome o que varia é o GROOT,
 *         e repetir o nome ali não diria nada a quem revisa. */
function conflitos(registros, chaveDe, varianteDe, grupoRotuloDe, varianteRotuloDe, relacaoDe){
  const saida = [];
  agrupar(registros, chaveDe).forEach((rs, chave) => {
    const porVariante = agrupar(rs, varianteDe);
    if (porVariante.size < 2) return;                 // sem divergência, sem caso
    const variantes = [];
    porVariante.forEach((ocorrencias, valor) => {
      variantes.push({valor, rotulo: varianteRotuloDe(valor, ocorrencias),
                      vezes: ocorrencias.length, ocorrencias});
    });
    variantes.sort((a, b) => b.vezes - a.vezes || String(a.valor).localeCompare(String(b.valor)));
    const motivo = relacaoDe(variantes.map(v => v.valor));
    saida.push({chave, rotulo: grupoRotuloDe(rs[0]), variantes, motivo,
                explicacao: AUDIT_MOTIVO_TXT[motivo], registros: rs.length});
  });
  // mais variantes primeiro: o caso mais confuso é o que precisa de decisão antes
  saida.sort((a, b) => b.variantes.length - a.variantes.length
                    || b.registros - a.registros
                    || String(a.rotulo).localeCompare(String(b.rotulo)));
  return saida;
}

/** Audita a identidade da consolidação inteira, ANTES da deduplicação —
 *  é justamente o que a dedup uniria ou deixaria de unir que interessa.
 *
 *  @param {Array<{op:string, rows:Array}>} entradas
 *  @returns {{semGroot:Array, mesmoNome:Array, mesmoGroot:Array, resumo:Object}}
 */
function auditarIdentidade(entradas){
  const registros = planificar(entradas);
  const semGroot = registros.filter(r => !r.groot);

  /* Só entram na comparação registros que têm as DUAS pontas. Nome vazio não
     contradiz nome nenhum, e sem GROOT não há o que confrontar. */
  const comparaveis = registros.filter(r => r.groot && r.nomeChave);

  // sob um mesmo nome, o que varia é o GROOT
  const mesmoNome  = conflitos(comparaveis, r => r.nomeChave, r => r.groot,
                               r => r.nome, groot => groot, relacaoEntreGroots);
  // sob um mesmo GROOT, o que varia é o nome — mostrado como veio da planilha,
  // não como a chave normalizada de comparação
  const mesmoGroot = conflitos(comparaveis, r => r.groot, r => r.nomeChave,
                               r => r.groot, (chave, ocs) => ocs[0].nome, relacaoEntreNomes);

  return {
    semGroot, mesmoNome, mesmoGroot,
    resumo: {
      registros: registros.length,
      semGroot: semGroot.length,
      mesmoNome: mesmoNome.length,
      mesmoGroot: mesmoGroot.length,
      // o subconjunto que não dá para explicar por grafia — o que dói de verdade
      mesmoGrootNomesDistintos: mesmoGroot.filter(g => g.motivo === AUDIT_MOTIVO.NOMES).length,
      mesmoNomeIdsDistintos:    mesmoNome.filter(g => g.motivo === AUDIT_MOTIVO.IDS).length
    }
  };
}

if (typeof module !== "undefined" && module.exports){
  module.exports = {AUDIT_MOTIVO, AUDIT_MOTIVO_TXT, auditarIdentidade,
                    relacaoEntreGroots, relacaoEntreNomes};
}
