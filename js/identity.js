/* Ajustes MELI — identidade de pessoa, normalizada em um lugar só
   ================================================================

   GROOT ID representa UMA pessoa. Se a mesma combinação de GROOT e dia
   aparece em dois lugares, não são duas pessoas nem duas diárias — é a
   mesma pessoa-dia repetida. Essa é a regra que evita cobrança dupla,
   contagem dupla de diarista e compensação indevida de ABS.

   Toda comparação de GROOT do aplicativo passa por aqui. Duas
   implementações da mesma regra é como se cria inconsistência entre
   abas: a Conciliação achando que são a mesma pessoa e a Extração
   achando que não.
   ================================================================ */
"use strict";

/** Normaliza um GROOT ID para comparação.
 *
 *  O Excel devolve o mesmo identificador de várias formas — número, texto,
 *  float com `.0`, com espaços em volta. Todas essas viram a mesma chave.
 *
 *  O que NÃO é feito de propósito: converter para número e voltar. Isso
 *  destruiria zeros à esquerda, e `00123456` pode ser um identificador
 *  diferente de `123456` — não há como saber pelo dado. Na dúvida, o
 *  conservador é não unir: unir duas pessoas por engano é pior do que
 *  deixar de unir duas grafias da mesma. */
function normalizeGroot(v){
  if(v===null||v===undefined) return "";
  let s=String(v).trim();
  if(!s) return "";

  // 123456.0 / 123456.00 → 123456   (o ".0" é ruído de float, não é o id)
  if(/^-?\d+\.0+$/.test(s)) return s.replace(/\.0+$/,"");

  // 1.23456e5 → 123456   (notação científica de célula numérica larga)
  if(/^-?\d+(\.\d+)?e[+-]?\d+$/i.test(s)){
    const n=Number(s);
    if(Number.isFinite(n)&&Number.isInteger(n)) return String(n);
  }

  // Inteiro puro, alfanumérico, ou com zeros à esquerda: preservado como está.
  return s.toUpperCase();
}

/** Um GROOT que identifica alguém de fato. "0" e vazio não identificam. */
function hasGroot(v){ const s=normalizeGroot(v); return s!=="" && s!=="0"; }

/** Normaliza uma data para a chave `AAAA-MM-DD`, venha ela como for:
 *  objeto Date, serial do Excel, DD/MM/AAAA, AAAA-MM-DD ou inteiro AAAAMMDD. */
function normalizeDateKey(v){
  if(v===null||v===undefined||v==="") return "";
  const iso=(y,m,d)=>y+"-"+String(m).padStart(2,"0")+"-"+String(d).padStart(2,"0");

  if(v instanceof Date){
    if(isNaN(v.getTime())) return "";
    // datas do ExcelJS chegam em UTC; usar getters UTC não desloca um dia
    return iso(v.getUTCFullYear(), v.getUTCMonth()+1, v.getUTCDate());
  }
  if(typeof v==="number"&&isFinite(v)){
    // inteiro AAAAMMDD, o formato interno de dates.js
    if(v>=19900101&&v<=21001231&&Number.isInteger(v)){
      return iso(Math.floor(v/10000), Math.floor(v/100)%100, v%100);
    }
    // serial do Excel
    if(v>0&&v<80000){
      const d=new Date(Date.UTC(1899,11,30)+Math.round(v)*86400000);
      return iso(d.getUTCFullYear(), d.getUTCMonth()+1, d.getUTCDate());
    }
    return "";
  }
  const s=String(v).trim();
  if(!s) return "";
  let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(m) return iso(+m[1],+m[2],+m[3]);
  m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if(m){ let y=+m[3]; if(y<100)y+=2000; return iso(y,+m[2],+m[1]); }
  if(/^\d+(\.\d+)?$/.test(s)) return normalizeDateKey(Number(s));
  return "";
}

/** A chave pessoa-dia: `GROOT|AAAA-MM-DD`.
 *  Devolve null quando não há GROOT — sem identificador não existe
 *  pessoa-dia, e tratar todos os vazios como a mesma pessoa apagaria
 *  gente de verdade. */
function personDayKey(groot, data){
  if(!hasGroot(groot)) return null;
  const d=normalizeDateKey(data);
  if(!d) return null;
  return normalizeGroot(groot)+"|"+d;
}

if(typeof module!=="undefined"&&module.exports){
  module.exports={normalizeGroot, hasGroot, normalizeDateKey, personDayKey};
}
