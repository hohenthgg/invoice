/* Extração · Diarista — deduplicação por pessoa-dia
   ================================================================

   GROOT representa uma pessoa. Logo:

       pessoa-dia = GROOT normalizado + data normalizada

   Se a mesma combinação aparece em duas operações, duas abas ou dois
   arquivos, não são duas pessoas nem duas diárias: é a mesma pessoa-dia
   repetida. Contá-la duas vezes produz dupla cobrança do diarista e
   compensação indevida de ABS.

   Por isso o controle de duplicidade é GLOBAL — vive na consolidação
   inteira, não dentro de cada operação. Um `seen` por operação faria
   `SVC|123456|10/08` e `XD|123456|10/08` sobreviverem os dois.

   Módulo puro, sem DOM: a regra é testável isoladamente.
   ================================================================ */
"use strict";

const DEDUP_MODOS = {
  DIA:     "dia",       // duplicado só quando repete na mesma data
  PERIODO: "periodo",   // uma linha por pessoa no período inteiro
  NAO:     "nao"        // não remover nada
};

/** Deduplica a consolidação inteira.
 *
 *  @param {Array<{op:string, rows:Array}>} entradas — as linhas já filtradas
 *         por período e solicitante, agrupadas por operação e na ordem em que
 *         devem ser consideradas (a primeira ocorrência é a que fica).
 *  @param {string} modo — DEDUP_MODOS
 *  @returns {{porOperacao:Object, resumo:Object, descartados:Array, semGroot:Array}}
 */
function deduplicarPessoaDia(entradas, modo){
  const porOperacao={}, descartados=[], semGroot=[];
  /* O conjunto vive AQUI, fora do laço de operações. É essa única linha
     que faz a diferença entre uma pessoa-dia e duas. */
  const vistos=new Map();          // chave → {op, date, id}
  let encontrados=0, duplicados=0, unicos=0;

  entradas.forEach(({op, rows})=>{
    const mantidas=[]; let dupsDaOperacao=0;

    (rows||[]).forEach(r=>{
      encontrados++;

      if(modo===DEDUP_MODOS.NAO){ mantidas.push(r); unicos++; return; }

      /* Sem GROOT não existe pessoa-dia. Tratar todos os vazios como a mesma
         chave apagaria pessoas diferentes de uma vez — o oposto do objetivo.
         Ficam todas, sinalizadas para revisão. */
      if(!hasGroot(r.id)){
        mantidas.push(r); unicos++;
        /* `reg` é o registro inteiro: quem revisa precisa saber QUEM é a
           pessoa (nome, empresa, cargo, escala, solicitante), não só que
           existe um vazio. Guardar só a contagem transforma um pedido de
           revisão em um número que ninguém consegue conferir. */
        semGroot.push({op, date:r.date, nome:r.nome||"", reg:r});
        return;
      }

      const chave = modo===DEDUP_MODOS.DIA
        ? personDayKey(r.id, r.date)
        : normalizeGroot(r.id);

      if(chave===null){ mantidas.push(r); unicos++; return; }   // data ilegível

      if(vistos.has(chave)){
        const primeira=vistos.get(chave);
        duplicados++; dupsDaOperacao++;
        descartados.push({
          groot:normalizeGroot(r.id), data:normalizeDateKey(r.date),
          nome:r.nome||"", mantidaEm:primeira.op, descartadaDe:op
        });
        return;                     // mantém a PRIMEIRA ocorrência encontrada
      }
      vistos.set(chave,{op, date:r.date, id:r.id});
      mantidas.push(r); unicos++;
    });

    porOperacao[op]={rows:mantidas, dups:dupsDaOperacao,
      filtered:(rows||[]).length, dedupOff:modo===DEDUP_MODOS.NAO};
  });

  return {porOperacao, descartados, semGroot,
    resumo:{encontrados, unicos, duplicados, semGroot:semGroot.length,
      modo, global:true}};
}

if(typeof module!=="undefined"&&module.exports){
  module.exports={DEDUP_MODOS, deduplicarPessoaDia};
}
