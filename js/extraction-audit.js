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

/* ================================================================
   NOME QUEBRADO
   ----------------------------------------------------------------
   Um nome só serve para achar a pessoa. "#N/D", "A DEFINIR", "JOAO 2"
   e "JoÃ£o" não acham ninguém — e, pior, passam despercebidos porque a
   célula não está vazia. Cada um tem uma origem diferente e um conserto
   diferente, então cada um recebe seu próprio motivo.

   Nada é descartado por causa disso: o registro continua na extração.
   ================================================================ */
const NOME_PROBLEMA = {
  VAZIO:        "vazio",
  ERRO_PLANILHA:"erro_planilha",
  CODIFICACAO:  "codificacao",
  PLACEHOLDER:  "placeholder",
  DIGITOS:      "digitos",
  SIMBOLOS:     "simbolos",
  CURTO:        "curto",
  REPETICAO:    "repeticao",
  UM_TOKEN:     "um_token"
};

const NOME_PROBLEMA_TXT = {
  vazio:         "Sem nome — a célula está vazia.",
  erro_planilha: "A célula traz um erro do Excel (#N/D, #REF!, …) no lugar do nome: a fórmula que buscava o nome não achou nada.",
  codificacao:   "Acentuação corrompida — o arquivo passou por uma conversão que quebrou a codificação.",
  placeholder:   "Texto genérico no lugar do nome (a definir, sem nome, teste…) — o cadastro nunca foi preenchido.",
  digitos:       "Há números no meio do nome — provável matrícula, código ou contador colado na célula.",
  simbolos:      "Há símbolos que não pertencem a um nome.",
  curto:         "Curto demais para ser um nome.",
  repeticao:     "Caractere repetido em sequência (xxxx, aaaa) — preenchimento de rascunho.",
  um_token:      "Só uma palavra, sem sobrenome — pode ser cadastro incompleto."
};

/* Um único termo genérico não identifica pessoa alguma. Comparados já
   normalizados (sem acento, sem pontuação, caixa alta). */
const NOMES_GENERICOS = new Set([
  "A DEFINIR","ADEFINIR","A CONFIRMAR","DEFINIR","SEM NOME","SEMNOME","SEM CADASTRO",
  "NAO INFORMADO","NAO INFORMADA","NAO IDENTIFICADO","NAO IDENTIFICADA","INDEFINIDO",
  "PENDENTE","FALTA","VAGA","VAGO","TESTE","TESTE 1","NOME","SEM","N D","N A","NA","ND",
  "X","XX","XXX","XXXX","XXXXX","NULL","NONE","NULO","DIARISTA","COLABORADOR","FUNCIONARIO"
]);

/** Devolve a lista de problemas de um nome — vazia quando o nome está bem. */
function problemasDoNome(bruto){
  const s = String(bruto === null || bruto === undefined ? "" : bruto).trim();
  const chave = normalizeNome(s);
  if (!s || !chave) return [NOME_PROBLEMA.VAZIO];

  const p = [];
  if (/#(N\/?D|N\/?A|REF|VALOR|VALUE|NOME|NAME|DIV\/0|NUM|NULO|NULL)!?\??/i.test(s))
    p.push(NOME_PROBLEMA.ERRO_PLANILHA);
  // Ã©, Â , ï¿½, � — resíduo de UTF-8 lido como Latin-1 e vice-versa
  if (/[\u00c3\u00c2][\u0080-\u00bf]|\ufffd|\u00ef\u00bf\u00bd/.test(s)) p.push(NOME_PROBLEMA.CODIFICACAO);
  if (NOMES_GENERICOS.has(chave)) p.push(NOME_PROBLEMA.PLACEHOLDER);
  if (/\d/.test(s)) p.push(NOME_PROBLEMA.DIGITOS);
  // letras, espaço e o que aparece em nome de verdade (hífen, apóstrofo, ponto)
  if (/[^\p{L}\s'’.\-]/u.test(s) && !p.includes(NOME_PROBLEMA.ERRO_PLANILHA)
      && !p.includes(NOME_PROBLEMA.DIGITOS)) p.push(NOME_PROBLEMA.SIMBOLOS);
  if (chave.replace(/[^A-Z]/g, "").length < 3) p.push(NOME_PROBLEMA.CURTO);
  if (/(.)\1{3,}/i.test(chave.replace(/\s/g, ""))) p.push(NOME_PROBLEMA.REPETICAO);

  /* Nome de uma palavra só é o único item que não é defeito garantido —
     existe gente cadastrada assim. Vira aviso, e só quando o resto está bem. */
  if (!p.length && tokensNome(s).length === 1) p.push(NOME_PROBLEMA.UM_TOKEN);
  return p;
}

/** Só os problemas que impedem identificar a pessoa — exclui o aviso brando. */
function nomeQuebrado(bruto){
  const p = problemasDoNome(bruto);
  return p.length > 0 && !(p.length === 1 && p[0] === NOME_PROBLEMA.UM_TOKEN);
}

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
      const problemas = problemasDoNome(r.nome);
      out.push({
        op, date: r.date,
        grootBruto: (r.id === null || r.id === undefined) ? "" : String(r.id),
        groot: hasGroot(r.id) ? normalizeGroot(r.id) : "",
        nome: String(r.nome == null ? "" : r.nome).trim(),
        nomeChave: normalizeNome(r.nome),
        problemasNome: problemas,
        nomeQuebrado: problemas.length > 0
                      && !(problemas.length === 1 && problemas[0] === NOME_PROBLEMA.UM_TOKEN),
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

  /* Nome quebrado é um problema por si, com conserto próprio. Deixá-lo na
     comparação faria "ANA SOUZA vs #N/D" ser diagnosticado como duas pessoas
     sob o mesmo identificador — apontando a causa errada. */
  const nomesQuebrados = registros.filter(r => r.nomeQuebrado);

  /* Só entram na comparação registros que têm as DUAS pontas legíveis. Nome
     vazio não contradiz nome nenhum, e sem GROOT não há o que confrontar. */
  const comparaveis = registros.filter(r => r.groot && r.nomeChave && !r.nomeQuebrado);

  // sob um mesmo nome, o que varia é o GROOT
  const mesmoNome  = conflitos(comparaveis, r => r.nomeChave, r => r.groot,
                               r => r.nome, groot => groot, relacaoEntreGroots);
  // sob um mesmo GROOT, o que varia é o nome — mostrado como veio da planilha,
  // não como a chave normalizada de comparação
  const mesmoGroot = conflitos(comparaveis, r => r.groot, r => r.nomeChave,
                               r => r.groot, (chave, ocs) => ocs[0].nome, relacaoEntreNomes);

  return {
    semGroot, nomesQuebrados, mesmoNome, mesmoGroot,
    resumo: {
      registros: registros.length,
      semGroot: semGroot.length,
      nomesQuebrados: nomesQuebrados.length,
      mesmoNome: mesmoNome.length,
      mesmoGroot: mesmoGroot.length,
      // o subconjunto que não dá para explicar por grafia — o que dói de verdade
      mesmoGrootNomesDistintos: mesmoGroot.filter(g => g.motivo === AUDIT_MOTIVO.NOMES).length,
      mesmoNomeIdsDistintos:    mesmoNome.filter(g => g.motivo === AUDIT_MOTIVO.IDS).length
    }
  };
}

/* ================================================================
   TÓPICOS
   ----------------------------------------------------------------
   A auditoria produz estruturas de formatos diferentes — uma lista de
   registros aqui, grupos com variantes ali. Para a tela (os balões) e
   para a planilha (uma linha por problema) o que serve é uma coisa só:
   tópicos, cada um com gravidade, o que significa e o que fazer.
   ================================================================ */
const GRAVIDADE = { ALTA:"alta", MEDIA:"media", BAIXA:"baixa", INFO:"info" };
const GRAVIDADE_TXT = { alta:"Grave", media:"Revisar", baixa:"Provável grafia", info:"Informativo" };
const GRAVIDADE_ORDEM = { alta:0, media:1, baixa:2, info:3 };

/** Uma linha por problema encontrado, pronta para a tela e para o Excel.
 *  @param aud          resultado de auditarIdentidade
 *  @param descartados  os duplicados removidos pela deduplicação
 *  @param leitura      {semData, abas} — o que a leitura do arquivo não aproveitou
 *  @param semHorario   operações sem escala horário levantada, com o total de
 *                      registros afetados */
function listarTopicos(aud, descartados, leitura, semHorario){
  const t = [];
  const ctx = r => ({op:r.op, date:r.date, empresa:r.empresa, cargo:r.cargo,
                     escala:r.escala, solicitante:r.solicitante});
  const vazio = {empresa:"", cargo:"", escala:"", solicitante:""};

  /* Vem primeiro por ser o mais amplo: uma aba fora derruba a filial inteira,
     e nenhum outro tópico consegue apontar o que nem foi lido. */
  ((leitura && leitura.abas) || []).forEach(a => t.push(Object.assign({
    topico: "Aba não lida", gravidade: GRAVIDADE.ALTA,
    nome: "", groot: "", op: a.op, date: "",
    diagnostico: a.motivo + " Nenhum registro desta filial entrou na extração, "
               + "em nenhum período.",
    acao: "Conferir o nome da aba e o cabeçalho no arquivo de origem, e extrair de novo."
  }, vazio)));

  /* O horário não vem do SIGO — vem do padrão da operação, levantado na
     fatura. Sem fatura da unidade não há padrão, e a coluna sai vazia. Chutar
     o horário de outra filial seria pior do que a lacuna. */
  (semHorario || []).forEach(s => t.push(Object.assign({
    topico: "Sem escala horário", gravidade: GRAVIDADE.MEDIA,
    nome: "", groot: "", op: s.op, date: "",
    diagnostico: "Não há escala horário levantada para esta filial, então a coluna "
               + "ESCALA HORÁRIO sai em branco em " + s.registros + " registro"
               + (s.registros === 1 ? "" : "s") + ". O SIGO não informa horário: "
               + "ele vem do padrão da operação, lido na fatura 3PL.",
    acao: "Enviar uma fatura 3PL desta unidade para levantar o horário padrão do diarista."
  }, vazio)));

  ((leitura && leitura.semData) || []).forEach(l => t.push(Object.assign({
    topico: "Linha sem data legível", gravidade: GRAVIDADE.MEDIA,
    nome: l.nome, groot: l.groot, op: l.op, date: "",
    diagnostico: "A linha " + l.linha + " da aba tem conteúdo, mas a data não pôde ser lida"
               + (l.valor ? ' (a célula trazia "' + l.valor + '")' : " (célula vazia)")
               + ". Ficou fora da extração inteira — não é filtro de período.",
    acao: "Corrigir a data na origem. Enquanto isso, esta diária não é contada em lugar nenhum."
  }, vazio)));

  aud.semGroot.forEach(r => t.push(Object.assign({
    topico: "Sem identificador", gravidade: GRAVIDADE.MEDIA,
    nome: r.nome, groot: r.grootBruto,
    diagnostico: "O registro existe, o GROOT não. Sem identificador não há pessoa-dia: "
               + "este registro ficou fora da deduplicação e não pode ser confrontado com nenhum outro.",
    acao: "Preencher o GROOT ID na planilha de origem."
  }, ctx(r))));

  aud.nomesQuebrados.forEach(r => t.push(Object.assign({
    topico: "Nome quebrado", gravidade: GRAVIDADE.MEDIA,
    nome: r.nome, groot: r.grootBruto,
    diagnostico: r.problemasNome.map(p => NOME_PROBLEMA_TXT[p]).join(" "),
    acao: "Corrigir o nome na origem — sem ele não dá para conferir de quem é o registro."
  }, ctx(r))));

  aud.mesmoNome.forEach(g => g.variantes.forEach(v => v.ocorrencias.forEach(o => t.push(Object.assign({
    topico: "Mesmo nome, GROOTs diferentes",
    gravidade: g.motivo === AUDIT_MOTIVO.IDS ? GRAVIDADE.MEDIA : GRAVIDADE.BAIXA,
    nome: g.rotulo, groot: v.rotulo,
    diagnostico: g.explicacao,
    acao: g.motivo === AUDIT_MOTIVO.IDS
      ? "Conferir se é homônimo ou cadastro duplicado. Se for a mesma pessoa, está sendo contada em dobro."
      : "Padronizar a grafia do identificador na origem."
  }, ctx(o))))));

  aud.mesmoGroot.forEach(g => g.variantes.forEach(v => v.ocorrencias.forEach(o => t.push(Object.assign({
    topico: "Mesmo GROOT, nomes diferentes",
    gravidade: g.motivo === AUDIT_MOTIVO.NOMES ? GRAVIDADE.ALTA : GRAVIDADE.BAIXA,
    nome: v.rotulo, groot: g.rotulo,
    diagnostico: g.explicacao,
    acao: g.motivo === AUDIT_MOTIVO.NOMES
      ? "Conferir ANTES de usar o resultado: a deduplicação uniu estes registros."
      : "Padronizar a grafia do nome na origem."
  }, ctx(o))))));

  (descartados || []).forEach(d => t.push({
    topico: "Duplicado removido", gravidade: GRAVIDADE.INFO,
    nome: d.nome, groot: d.groot, op: d.descartadaDe, date: d.data,
    empresa:"", cargo:"", escala:"", solicitante:"",
    diagnostico: "Mesma pessoa-dia já contada em " + d.mantidaEm
               + ". Ficou a primeira ocorrência encontrada.",
    acao: "Nenhuma — remoção esperada. Confira só se as operações estiverem erradas."
  }));

  t.sort((a, b) => GRAVIDADE_ORDEM[a.gravidade] - GRAVIDADE_ORDEM[b.gravidade]
                || String(a.topico).localeCompare(String(b.topico))
                || String(a.date).localeCompare(String(b.date)));
  return t;
}

/** Os balões: um por tópico presente, na ordem de quem precisa de atenção antes. */
function resumirTopicos(topicos){
  const m = new Map();
  topicos.forEach(t => {
    if (!m.has(t.topico)) m.set(t.topico, {topico:t.topico, gravidade:t.gravidade, total:0});
    const b = m.get(t.topico);
    b.total++;
    if (GRAVIDADE_ORDEM[t.gravidade] < GRAVIDADE_ORDEM[b.gravidade]) b.gravidade = t.gravidade;
  });
  return Array.from(m.values())
    .sort((a, b) => GRAVIDADE_ORDEM[a.gravidade] - GRAVIDADE_ORDEM[b.gravidade] || b.total - a.total);
}

if (typeof module !== "undefined" && module.exports){
  module.exports = {AUDIT_MOTIVO, AUDIT_MOTIVO_TXT, NOME_PROBLEMA, NOME_PROBLEMA_TXT,
                    GRAVIDADE, GRAVIDADE_TXT, auditarIdentidade, problemasDoNome, nomeQuebrado,
                    listarTopicos, resumirTopicos, relacaoEntreGroots, relacaoEntreNomes};
}
