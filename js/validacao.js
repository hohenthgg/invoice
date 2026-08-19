/* Validação Template — auditoria de uma fatura 3PL no template padrão
   ================================================================

   Módulo PURO, sem DOM: recebe as abas já lidas como objetos simples e
   devolve os achados. Quem lê o .xlsx é js/validacao-ui.js. Essa
   separação existe para os testes rodarem contra o raciocínio de
   verdade, não contra uma cópia dele.

   O QUE ESTE MÓDULO É

   Um auditor, não um validador de campos. A diferença: um validador
   olha uma célula por vez e diz "isto está fora do formato"; um auditor
   junta evidências de linhas diferentes e conclui. As duas linhas do
   mesmo GROOT ID abaixo têm exatamente os mesmos campos preenchidos, e
   uma é erro grave enquanto a outra é rotina:

     GROOT 2110045  MARIANA COSTA DO AMARAL SANTOS  Temporário  30/03→10/08
     GROOT 2110045  MARIANA COSTA DO AMARAL             Efetivo     11/08→aberto

     GROOT 2110046  BRUNO TEIXEIRA LOPES   Efetivo     Líder      23/03→aberto
     GROOT 2110046  SERGIO ALMEIDA CUNHA       Temporário  Auxiliar   03/06→04/08

   No primeiro caso os nomes são o mesmo nome (um traz um sobrenome a
   mais), os períodos se encostam sem sobrepor e o vínculo evolui de
   Temporário para Efetivo: é a efetivação de uma pessoa, e o app
   sugere consolidar. No segundo os nomes não têm relação, os cargos
   são diferentes e os períodos se sobrepõem: duas pessoas cobrando
   pelo mesmo GROOT ID.

   Nenhuma regra isolada distingue os dois. A distinção vem de combinar
   nome + cargo + vínculo + período + sinal do valor — e é isso que
   cada regra aqui faz, registrando em `raciocinio` como chegou lá.

   TRÊS PRINCÍPIOS

   1. CONSERVADOR. Na dúvida, a severidade cai (Revisar em vez de
      Crítico). Falso positivo em fatura custa confiança: quem recebe
      dez alertas errados para de ler o décimo primeiro.
   2. EXPLICÁVEL. Todo achado carrega o raciocínio em texto e os
      registros que o sustentam. Quem lê precisa poder discordar.
   3. NÃO DECIDE. O módulo aponta, explica e sugere. Nenhuma linha da
      fatura é alterada aqui — a decisão é de quem audita.
   ================================================================ */
"use strict";

/* ---------------------------------------------------------------- */
/* Severidades, na ordem em que importam.                            */
const VAL_SEV = {
  CRITICO:  "critico",     // provável erro que gera faturamento incorreto
  REVISAR:  "revisar",     // suspeito, mas pode ser legítimo
  CADASTRO: "cadastro",    // padronização / qualidade de dado
  INFO:     "info"         // encontrado e aparentemente justificável
};
const VAL_SEV_ORDEM = [VAL_SEV.CRITICO, VAL_SEV.REVISAR, VAL_SEV.CADASTRO, VAL_SEV.INFO];
const VAL_SEV_META = {
  critico:  { label:"Crítico",     desc:"Provável erro que pode gerar faturamento incorreto." },
  revisar:  { label:"Revisar",     desc:"Situação suspeita, mas que pode ser legítima." },
  cadastro: { label:"Cadastro",    desc:"Problema de padronização ou qualidade de dado." },
  info:     { label:"Informativo", desc:"Situação encontrada, aparentemente justificável." }
};

/* Cargos em que o GROOT ID é especialmente esperado: são o operacional
   que o cliente concilia pessoa a pessoa. Ausência aqui pesa mais. */
const VAL_CARGOS_CRITICOS = ["auxiliar", "operador"];

/* A linha de absenteísmo do template usa "ABS" no lugar do GROOT ID.
   Não é cadastro incompleto — é um marcador do próprio modelo. */
const VAL_GROOT_RESERVADOS = ["abs"];

const VAL_PARTICULAS = new Set(["de","da","do","dos","das","e","di","du"]);

/* ================================================================
   NORMALIZAÇÃO
   ================================================================ */
function valNorm(s){
  return String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .toUpperCase().replace(/[^A-Z0-9 ]/g," ").replace(/\s+/g," ").trim();
}
function valTokens(nome){
  return valNorm(nome).split(" ").filter(t => t && !VAL_PARTICULAS.has(t.toLowerCase()));
}
/* GROOT como texto comparável: o Excel devolve ora "2110046", ora
   2110046, ora 2110046.0 — os três são a mesma pessoa. */
function valGroot(v){
  if(v === null || v === undefined) return "";
  let s = String(v).trim();
  if(s === "") return "";
  if(/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/,"");        // 2110046.0
  return s;
}
function valTemGroot(v){
  const s = valGroot(v);
  return s !== "" && s !== "0";
}
function valDia(d){ return d instanceof Date ? Math.floor(d.getTime()/864e5) : null; }
function valFmt(d){
  if(!(d instanceof Date)) return "—";
  const p = n => String(n).padStart(2,"0");
  return p(d.getUTCDate()) + "/" + p(d.getUTCMonth()+1) + "/" + d.getUTCFullYear();
}
/* O app não escreve valor monetário em lugar nenhum — nem na tela, nem
   no relatório. O que o raciocínio precisa do lançamento é só o SINAL:
   positivo é cobrança, negativo é estorno, e é essa diferença que
   separa dupla cobrança de ajuste retroativo. A grandeza não muda
   nenhuma conclusão, então não é escrita. */
function valSinal(v){
  const n = Number(v);
  if(!isFinite(n) || n === 0) return "neutro";
  return n > 0 ? "positivo" : "negativo";
}
const valEhNegativo = v => valSinal(v) === "negativo";
const valEhPositivo = v => valSinal(v) === "positivo";

/* Distância de edição, limitada — só precisamos saber se é pequena. */
function valLev(a,b){
  if(a === b) return 0;
  if(Math.abs(a.length-b.length) > 3) return 99;
  const m = a.length, n = b.length;
  let prev = Array.from({length:n+1},(_,j)=>j), cur = new Array(n+1);
  for(let i=1;i<=m;i++){
    cur[0] = i;
    for(let j=1;j<=n;j++){
      cur[j] = Math.min(prev[j]+1, cur[j-1]+1, prev[j-1] + (a[i-1]===b[j-1]?0:1));
    }
    [prev,cur] = [cur,prev];
  }
  return prev[n];
}

/* ================================================================
   COMPARAÇÃO DE NOMES

   Devolve a RELAÇÃO entre dois nomes, nunca uma decisão sobre as
   pessoas. Quem decide é a regra, que também tem período, cargo e
   vínculo na mão.

     iguais    — o mesmo nome depois de normalizar
     variacao  — quase certamente a mesma pessoa: um nome é o outro com
                 um sobrenome a mais, ou difere por erro de digitação
     parcial   — o primeiro nome coincide e os sobrenomes não: pode ser
                 troca de cadastro, pode ser duas pessoas
     distintos — sem relação aparente
   ================================================================ */
function valCompararNomes(a,b){
  const na = valNorm(a), nb = valNorm(b);
  if(na === nb) return { relacao:"iguais", motivo:"nomes idênticos" };
  if(!na || !nb) return { relacao:"distintos", motivo:"um dos nomes está vazio" };

  const ta = valTokens(a), tb = valTokens(b);
  const setB = new Set(tb);
  const comuns = ta.filter(t => setB.has(t));

  /* Contenção: "MARIANA COSTA DO AMARAL" dentro de "MARIANA COSTA DO
     AMARAL SANTOS". Exige 2 tokens em comum para um "MARIA" solto não
     engolir toda Maria da planilha. */
  if(comuns.length >= 2 && (comuns.length === ta.length || comuns.length === tb.length)){
    return { relacao:"variacao",
      motivo:"um dos nomes é o outro com "
        + Math.abs(ta.length-tb.length) + " sobrenome(s) a mais" };
  }

  /* Digitação: mesma estrutura, um único token diferente e próximo.
     "SIMONE APARECIDA CARDOSO" × "SIMONE PARECIDA CARDOSO". */
  if(ta.length === tb.length && ta.length >= 2){
    const difs = [];
    for(let i=0;i<ta.length;i++) if(ta[i] !== tb[i]) difs.push([ta[i],tb[i]]);
    if(difs.length === 1){
      const [x,y] = difs[0], d = valLev(x,y);
      if(d <= 2 && Math.max(x.length,y.length) >= 5){
        return { relacao:"variacao", motivo:'"'+x+'" × "'+y+'" diferem por '+d+' caractere(s)' };
      }
    }
  }

  /* Primeiro nome igual (ou quase) e o resto sem relação. */
  const p0 = ta[0] || "", p1 = tb[0] || "";
  if(p0 && p1 && (p0 === p1 || valLev(p0,p1) <= 1)){
    return { relacao:"parcial",
      motivo: comuns.length
        ? "coincidem "+comuns.length+" nome(s), os demais são diferentes"
        : "o primeiro nome coincide, os sobrenomes são completamente diferentes" };
  }
  return { relacao:"distintos",
    motivo: comuns.length ? "coincide apenas "+comuns.length+" nome(s) soltos"
                          : "nenhum nome em comum" };
}

/* ================================================================
   PERÍODOS
   ================================================================ */
/* Vínculo sem DATA FIM vale até o fim do período faturado — é assim
   que a fatura o cobra, então é assim que ele tem de ser comparado. */
function valFimEfetivo(reg, fimPeriodo){
  return reg.fim instanceof Date ? reg.fim : (fimPeriodo || null);
}
function valSobrepoe(a,b,fimPeriodo){
  const ai = valDia(a.ini), bi = valDia(b.ini);
  if(ai === null || bi === null) return false;
  const af = valDia(valFimEfetivo(a,fimPeriodo)), bf = valDia(valFimEfetivo(b,fimPeriodo));
  const AF = af === null ? Infinity : af, BF = bf === null ? Infinity : bf;
  return ai <= BF && bi <= AF;
}
function valDiasSobrepostos(a,b,fimPeriodo){
  const ai = valDia(a.ini), bi = valDia(b.ini);
  const af = valDia(valFimEfetivo(a,fimPeriodo)), bf = valDia(valFimEfetivo(b,fimPeriodo));
  if(ai === null || bi === null) return 0;
  const AF = af === null ? Infinity : af, BF = bf === null ? Infinity : bf;
  const ini = Math.max(ai,bi), fim = Math.min(AF,BF);
  return fim >= ini && isFinite(fim) ? (fim-ini+1) : (fim >= ini ? Infinity : 0);
}
/* "Encostados": um termina e o outro começa no dia seguinte. É a
   assinatura de uma troca de vínculo, não de duas pessoas. */
function valEncostados(a,b){
  const af = valDia(a.fim), bi = valDia(b.ini);
  return af !== null && bi !== null && bi - af === 1;
}
const valRegimeNorm = r => valNorm(r).toLowerCase();
const valEhTemporario = r => /tempor/.test(valRegimeNorm(r));
const valEhEfetivo    = r => /efetivo|clt|fixo/.test(valRegimeNorm(r));

/* ================================================================
   ACHADOS
   ================================================================ */
function valCriarAchado(o){
  return {
    id: o.id,
    regra: o.regra,
    severidade: o.severidade,
    titulo: o.titulo,
    groot: o.groot ?? "",
    nome: o.nome ?? "",
    cargo: o.cargo ?? "",
    aba: o.aba ?? "",
    datas: o.datas ?? "",
    /* sinal do lançamento, nunca a grandeza */
    lancamento: o.lancamento ?? "",
    /* peso serve só para ordenar os achados dentro da severidade; não é
       exibido em lugar nenhum, e por isso não é menção monetária */
    peso: o.peso ?? 0,
    raciocinio: o.raciocinio,
    sugestao: o.sugestao ?? "",
    opcoes: o.opcoes ?? ["Corrigido","Aceito/justificado","Ignorar"],
    registros: o.registros ?? []
  };
}
/* Um registro do jeito que a tela mostra lado a lado. */
function valRegistro(r){
  return {
    aba: r.aba, linha: r.linha, groot: r.groot, nome: r.nome, cargo: r.cargo,
    regime: r.regime ?? "", ini: r.ini ?? null, fim: r.fim ?? null,
    data: r.data ?? null, qtd: r.qtd ?? null, lancamento: valSinal(r.valor)
  };
}

/* ================================================================
   REGRA 1 — UM GROOT ID, MAIS DE UMA PESSOA

   A regra mais cara da auditoria, e a que mais precisa raciocinar: o
   mesmo GROOT ID aparecendo duas vezes pode ser efetivação (rotina),
   correção cadastral (rotina), estorno (rotina) ou duas pessoas
   cobrando pelo mesmo identificador (erro grave).

   A conclusão sai da combinação, nesta ordem de peso:
     nome (relação)  →  período (sobrepõe? encosta?)
     →  vínculo (Temporário → Efetivo?)  →  sinal do valor  →  cargo
   ================================================================ */
function valRegraGrootCompartilhado(labor, ctx, achados){
  const porGroot = new Map();
  for(const r of labor){
    if(!valTemGroot(r.groot) || VAL_GROOT_RESERVADOS.includes(valNorm(r.groot).toLowerCase())) continue;
    const k = valGroot(r.groot);
    if(!porGroot.has(k)) porGroot.set(k,[]);
    porGroot.get(k).push(r);
  }

  for(const [groot, regs] of porGroot){
    if(regs.length < 2) continue;

    /* Pares, não o grupo inteiro: três linhas do mesmo GROOT podem
       conter uma efetivação legítima E uma invasão de terceiro. */
    for(let i=0;i<regs.length;i++) for(let j=i+1;j<regs.length;j++){
      const a = regs[i], b = regs[j];
      const cmp = valCompararNomes(a.nome, b.nome);
      const sobrepoe = valSobrepoe(a,b,ctx.fim);
      const dias = valDiasSobrepostos(a,b,ctx.fim);
      const encostados = valEncostados(a,b) || valEncostados(b,a);
      const cargosIguais = valNorm(a.cargo) === valNorm(b.cargo);
      const ambosPositivos = valEhPositivo(a.valor) && valEhPositivo(b.valor);
      const algumNegativo = valEhNegativo(a.valor) || valEhNegativo(b.valor);
      const emAberto = [a,b].filter(x => !(x.fim instanceof Date)).length;
      const registros = [valRegistro(a), valRegistro(b)];
      const base = { groot, nome:a.nome+"  ×  "+b.nome, aba:"LABOR",
        datas: valFmt(a.ini)+"→"+valFmt(a.fim)+"  ×  "+valFmt(b.ini)+"→"+valFmt(b.fim),
        lancamento: valSinal(a.valor)+" / "+valSinal(b.valor),
        peso: Math.abs(Number(a.valor)||0) + Math.abs(Number(b.valor)||0), registros };

      /* ---- mesma pessoa ---------------------------------------- */
      if(cmp.relacao === "iguais" || cmp.relacao === "variacao"){
        const transicao = (valEhTemporario(a.regime) && valEhEfetivo(b.regime))
                       || (valEhTemporario(b.regime) && valEhEfetivo(a.regime));

        if(encostados && !sobrepoe && transicao){
          const tmp = valEhTemporario(a.regime) ? a : b, efe = tmp === a ? b : a;
          achados.push(valCriarAchado({ ...base,
            id:"G1-"+groot+"-"+i+"-"+j, regra:"Mudança de vínculo", severidade:VAL_SEV.INFO,
            titulo:"Temporário → Efetivo no mesmo GROOT ID", cargo:a.cargo,
            raciocinio:
              "O GROOT "+groot+" aparece duas vezes, e as evidências apontam para a mesma pessoa: "
              + (cmp.relacao === "iguais" ? "os nomes são idênticos" : "os nomes são o mesmo nome ("+cmp.motivo+")")
              + ". O contrato temporário termina em "+valFmt(tmp.fim)+" e o efetivo começa em "
              + valFmt(efe.ini)+" — dia seguinte, sem sobreposição — e o vínculo evolui de Temporário "
              + "para Efetivo. Isso é a assinatura de uma efetivação, não de duas pessoas: "
              + "não há dia cobrado em duplicidade."
              + (cargosIguais ? " O cargo é o mesmo nos dois trechos, o que reforça a leitura."
                              : " Atenção: o cargo muda de \""+a.cargo+"\" para \""+b.cargo+"\" — "
                                + "promoção junto da efetivação é possível, mas confirme."),
            sugestao:
              "Se a regra de faturamento permitir, consolide numa única linha: apague a DATA FIM de "
              + valFmt(tmp.fim)+" da linha temporária, mantenha a DATA DE INÍCIO original de "
              + valFmt(tmp.ini)+" e atualize o REGIME DE CONTRATO para \""+efe.regime+"\". "
              + "Se o faturamento precisar separar os trechos por tarifa distinta, mantenha as duas linhas.",
            opcoes:["Manter separado","Consolidar","Ignorar alerta"]
          }));
          continue;
        }

        if(sobrepoe){
          achados.push(valCriarAchado({ ...base,
            id:"G2-"+groot+"-"+i+"-"+j, regra:"Períodos sobrepostos", cargo:a.cargo,
            severidade: ambosPositivos ? VAL_SEV.CRITICO : VAL_SEV.REVISAR,
            titulo:"Mesma pessoa com dois vínculos ativos ao mesmo tempo",
            raciocinio:
              "Os dois registros são da mesma pessoa ("+cmp.motivo+"), mas os períodos se sobrepõem em "
              + (isFinite(dias) ? dias+" dia(s)" : "todo o período")+". Enquanto a efetivação encosta um "
              + "período no outro, aqui eles convivem — a pessoa aparece contratada duas vezes nos mesmos dias."
              + (ambosPositivos
                  ? " Os dois lançamentos são positivos, então os dias sobrepostos estão sendo"
                    + " cobrados em dobro."
                  : " Um dos lançamentos é negativo, o que sugere estorno ou ajuste retroativo "
                    + "em vez de cobrança dupla — confirme antes de tratar como erro.")
              + (emAberto === 2 ? " Os dois estão em aberto (sem DATA FIM), então a sobreposição não se fecha sozinha." : ""),
            sugestao:
              "Feche a DATA FIM do vínculo anterior no dia antes do início do seguinte, ou remova a linha "
              + "duplicada se ela foi lançada por engano. Se um dos lançamentos for estorno, marque-o como ajuste.",
            opcoes:["Ajustar DATA FIM","Excluir linha duplicada","Marcar como ajuste retroativo","Ignorar alerta"]
          }));
          continue;
        }

        if(cmp.relacao === "variacao"){
          achados.push(valCriarAchado({ ...base,
            id:"G3-"+groot+"-"+i+"-"+j, regra:"Variação cadastral de nome", severidade:VAL_SEV.CADASTRO,
            titulo:"Mesma pessoa grafada de duas formas", cargo:a.cargo,
            raciocinio:
              "O GROOT "+groot+" traz dois nomes que são o mesmo nome ("+cmp.motivo+"), em períodos que não "
              + "se sobrepõem. Não há risco de cobrança dupla: é qualidade de cadastro, e importa porque "
              + "uma conferência por nome — do lado do cliente ou seu — deixaria de casar as duas linhas.",
            sugestao:"Padronize a grafia nas duas linhas, adotando a forma que consta no cadastro oficial da pessoa.",
            opcoes:["Corrigido","Aceito/justificado","Ignorar"]
          }));
        }
        continue;
      }

      /* ---- pessoas diferentes ---------------------------------- */
      const distintos = cmp.relacao === "distintos";
      let severidade;
      if(sobrepoe && ambosPositivos) severidade = VAL_SEV.CRITICO;
      else if(sobrepoe)              severidade = VAL_SEV.REVISAR;
      else                           severidade = distintos ? VAL_SEV.REVISAR : VAL_SEV.REVISAR;

      const partes = [];
      partes.push("O GROOT "+groot+" aparece para \""+a.nome.trim()+"\" e \""+b.nome.trim()
        +"\", e "+(distintos ? "os nomes não têm relação ("+cmp.motivo+")"
                             : "só há coincidência parcial ("+cmp.motivo+")")+".");
      partes.push(cargosIguais
        ? "O cargo é o mesmo nos dois (\""+a.cargo+"\"), o que não ajuda a separar as pessoas nem a uni-las."
        : "Os cargos também são diferentes — \""+a.cargo+"\" e \""+b.cargo+"\". Cargos distintos não mudam "
          + "necessariamente a regra financeira, mas tornam improvável uma simples correção de cadastro: "
          + "quem corrige um nome não costuma trocar o cargo junto.");
      if(sobrepoe){
        partes.push("Os períodos se sobrepõem em "+(isFinite(dias) ? dias+" dia(s)" : "todo o período")
          + " ("+valFmt(a.ini)+"→"+valFmt(a.fim)+" contra "+valFmt(b.ini)+"→"+valFmt(b.fim)
          + "), então não é o caso de o GROOT ter sido reaproveitado depois que a primeira pessoa saiu.");
      } else {
        partes.push("Os períodos não se sobrepõem ("+valFmt(a.ini)+"→"+valFmt(a.fim)+" e depois "
          + valFmt(b.ini)+"→"+valFmt(b.fim)+"), o que abre a hipótese de o identificador ter sido "
          + "reaproveitado para outra pessoa após o desligamento da primeira — irregular, mas de "
          + "consequência financeira menor do que a cobrança simultânea.");
      }
      if(ambosPositivos && sobrepoe){
        partes.push("Os dois lançamentos são positivos: as duas pessoas estão sendo efetivamente "
          + "cobradas ao mesmo tempo pelo mesmo identificador.");
      } else if(algumNegativo){
        const neg = valEhNegativo(a.valor) ? a : b;
        partes.push("O lançamento de \""+neg.nome.trim()+"\" é negativo, o que sugere estorno ou "
          + "correção já em curso — possivelmente o próprio conserto deste problema. Confirme se o "
          + "estorno cobre exatamente o período que foi cobrado a mais.");
      }
      partes.push(sobrepoe && ambosPositivos
        ? "Improvável ser alteração cadastral. Recomenda-se identificar qual das duas pessoas tem o GROOT correto e corrigir a outra."
        : "Recomenda-se confirmar a titularidade do identificador antes de fechar a fatura.");

      achados.push(valCriarAchado({ ...base,
        id:"G4-"+groot+"-"+i+"-"+j, regra:"GROOT ID com pessoas diferentes", severidade,
        titulo: distintos ? "Mesmo GROOT ID para duas pessoas diferentes"
                          : "Mesmo GROOT ID com nomes parcialmente coincidentes",
        cargo: cargosIguais ? a.cargo : a.cargo+" / "+b.cargo,
        raciocinio: partes.join(" "),
        sugestao:
          "Verifique no cadastro qual pessoa é a titular do GROOT "+groot
          + " e solicite o identificador correto para a outra. Enquanto isso, não feche a fatura com as "
          + "duas linhas positivas no mesmo período.",
        opcoes:["Corrigir GROOT","Excluir linha indevida","Aceito/justificado","Ignorar alerta"]
      }));
    }
  }
}

/* ================================================================
   REGRA 2 — LABOR × DIARISTA NO MESMO PERÍODO

   A mesma pessoa não pode estar no quadro fixo e ser paga como
   diarista no mesmo dia. Mas a conclusão depende do SINAL do LABOR:
   um lançamento negativo é estorno, e estornar o fixo justamente para
   pagar como diarista é ajuste legítimo, não cobrança dupla.
   ================================================================ */
function valRegraLaborDiarista(labor, diaristas, ctx, achados){
  const porGroot = new Map();
  for(const d of diaristas){
    if(!valTemGroot(d.groot)) continue;
    const k = valGroot(d.groot);
    if(!porGroot.has(k)) porGroot.set(k,[]);
    porGroot.get(k).push(d);
  }

  labor.forEach((r,idx) => {
    if(!valTemGroot(r.groot) || VAL_GROOT_RESERVADOS.includes(valNorm(r.groot).toLowerCase())) return;
    const ds = porGroot.get(valGroot(r.groot));
    if(!ds || !(r.ini instanceof Date)) return;

    const fimEf = valFimEfetivo(r, ctx.fim);
    const ini = valDia(r.ini), fim = fimEf ? valDia(fimEf) : Infinity;
    const hits = ds.filter(d => {
      const dd = valDia(d.data);
      return dd !== null && dd >= ini && dd <= fim;
    });
    if(!hits.length) return;

    const negativo = valEhNegativo(r.valor);
    const datas = hits.map(h => valFmt(h.data)).join(", ");
    const semFim = !(r.fim instanceof Date);

    achados.push(valCriarAchado({
      id:"LD-"+valGroot(r.groot)+"-"+idx,
      regra:"LABOR × DIARISTA", severidade: negativo ? VAL_SEV.REVISAR : VAL_SEV.CRITICO,
      titulo: negativo ? "Sobreposição LABOR × DIARISTA com LABOR negativo"
                       : "Mesma pessoa no LABOR e como diarista nos mesmos dias",
      groot: valGroot(r.groot), nome: r.nome, cargo: r.cargo, aba:"LABOR + DIARISTAS",
      datas: valFmt(r.ini)+"→"+valFmt(r.fim)+" · diárias em "+datas,
      lancamento: valSinal(r.valor), peso: Math.abs(Number(r.valor)||0),
      raciocinio:
        "O GROOT "+valGroot(r.groot)+" ("+r.nome.trim()+") tem vínculo LABOR de "+valFmt(r.ini)
        + (semFim ? " em aberto — considerado até o fim do período faturado, "+valFmt(ctx.fim)+" — "
                  : " a "+valFmt(r.fim)+", ")
        + "e "+hits.length+" diária(s) lançada(s) dentro desse intervalo ("+datas+"). "
        + (negativo
            ? "O lançamento LABOR, porém, é negativo. Isso muda a leitura: um lançamento "
              + "negativo costuma representar estorno, desligamento retroativo ou outro ajuste de "
              + "faturamento — inclusive a própria devolução do fixo para pagar os dias como diária. "
              + "Não trate como dupla cobrança sem checar: o que precisa ser confirmado é se o estorno "
              + "cobre exatamente os dias em que a diária foi paga."
            : "O lançamento LABOR é positivo, então nesses dias a pessoa está sendo "
              + "cobrada duas vezes: uma pelo quadro fixo e outra pela diária. É dupla cobrança até prova "
              + "em contrário."),
      sugestao: negativo
        ? "Confirme com o financeiro a natureza do lançamento negativo. Se for estorno do fixo para "
          + "pagamento como diária, marque como ajuste retroativo e mantenha as duas linhas."
        : "Escolha uma das duas cobranças para os dias em conflito: ou encurte a DATA FIM do LABOR, ou "
          + "remova as diárias que caem dentro do vínculo fixo.",
      opcoes:["Manter lançamento","Excluir LABOR","Excluir diária","Marcar como ajuste retroativo","Ignorar alerta"],
      registros:[valRegistro(r), ...hits.map(valRegistro)]
    }));
  });
}

/* ================================================================
   REGRA 3 — GROOT ID AUSENTE
   ================================================================ */
function valRegraGrootAusente(linhas, aba, achados){
  linhas.forEach((r,idx) => {
    if(valTemGroot(r.groot)) return;
    if(VAL_GROOT_RESERVADOS.includes(valNorm(r.groot).toLowerCase())) return;
    if(!String(r.nome ?? "").trim()) return;             // linha vazia do template

    const cargoN = valNorm(r.cargo).toLowerCase();
    const critico = VAL_CARGOS_CRITICOS.some(c => cargoN.includes(c));
    achados.push(valCriarAchado({
      id:"SG-"+aba+"-"+idx,
      regra:"GROOT ID ausente", severidade: critico ? VAL_SEV.CRITICO : VAL_SEV.REVISAR,
      titulo:"Lançamento sem GROOT ID", groot:"", nome:r.nome, cargo:r.cargo, aba,
      datas: r.data instanceof Date ? valFmt(r.data) : valFmt(r.ini)+"→"+valFmt(r.fim),
      lancamento: valSinal(r.valor), peso: Math.abs(Number(r.valor)||0),
      raciocinio:
        "A linha de \""+String(r.nome).trim()+"\" ("+(r.cargo || "cargo em branco")+", aba "+aba
        + ") tem lançamento e nenhum GROOT ID. "
        + (critico
            ? "O cargo é operacional, e é exatamente nesses — auxiliar e operador — que o cliente concilia "
              + "pessoa a pessoa pelo GROOT. Sem o identificador, a linha não tem como ser conferida do "
              + "outro lado e tende a ser glosada."
            : "O cargo não está entre os que o cliente concilia pessoa a pessoa, então a ausência pesa menos "
              + "— mas continua impedindo o rastreio da linha.")
        + (valEhNegativo(r.valor)
            ? " O lançamento é negativo: sendo um estorno, o GROOT é ainda mais necessário para o "
              + "cliente casar a devolução com a cobrança original." : ""),
      sugestao:"Busque o GROOT ID desta pessoa no cadastro e preencha a coluna antes de enviar a fatura.",
      registros:[valRegistro(r)]
    }));
  });
}

/* ================================================================
   REGRA 4 — GROOT ID FORA DO PADRÃO

   O padrão não é declarado em lugar nenhum: é aprendido da própria
   planilha. Se 97% dos identificadores têm 7 dígitos, 7 é o padrão
   desta fatura — e o app diz isso em vez de afirmar que os outros
   estão errados.
   ================================================================ */
function valRegraGrootFormato(todas, achados){
  const contagem = new Map();
  for(const r of todas){
    const g = valGroot(r.groot);
    if(!g || VAL_GROOT_RESERVADOS.includes(g.toLowerCase())) continue;
    if(!/^\d+$/.test(g)) continue;
    contagem.set(g.length, (contagem.get(g.length)||0) + 1);
  }
  let padrao = null, maior = 0, total = 0;
  for(const [len,n] of contagem){ total += n; if(n > maior){ maior = n; padrao = len; } }
  const share = total ? maior/total : 0;

  const vistos = new Set();
  for(const r of todas){
    const bruto = String(r.groot ?? "");
    const g = valGroot(r.groot);
    if(!g || VAL_GROOT_RESERVADOS.includes(g.toLowerCase())) continue;
    if(vistos.has(g)) continue;
    vistos.add(g);

    const problemas = [];
    if(/\s/.test(bruto.trim())) problemas.push("contém espaço no meio");
    if(/[.,]/.test(bruto)) problemas.push("veio com casas decimais ("+bruto.trim()+")");
    if(!/^\d+$/.test(g)) problemas.push("contém caractere não numérico");
    else {
      if(/^0\d/.test(g)) problemas.push("tem zero à esquerda, que o Excel costuma perder");
      if(padrao && g.length !== padrao) problemas.push("tem "+g.length+" dígitos");
    }
    if(!problemas.length) continue;

    const soComprimento = problemas.length === 1 && /dígitos$/.test(problemas[0]);
    achados.push(valCriarAchado({
      id:"FM-"+g, regra:"GROOT ID fora do padrão",
      severidade: soComprimento ? VAL_SEV.CADASTRO : VAL_SEV.REVISAR,
      titulo:"Identificador fora do padrão predominante",
      groot:g, nome:r.nome, cargo:r.cargo, aba:r.aba,
      raciocinio:
        "O GROOT \""+bruto.trim()+"\" "+problemas.join(" e ")+". "
        + (padrao
            ? "Nesta fatura o padrão predominante é de "+padrao+" dígitos ("+Math.round(share*100)
              + "% dos identificadores). "
            : "")
        + (soComprimento
            ? "Isso não significa que esteja errado: identificadores antigos costumam ser mais curtos e "
              + "continuarem válidos. Vale conferir se é cadastro legado ou preenchimento incompleto."
            : "Formatação diferente do resto tende a quebrar a conferência automática do outro lado, "
              + "que compara texto com texto."),
      sugestao: soComprimento
        ? "Confirme no cadastro se o identificador é legado. Sendo válido, marque como justificado para não reaparecer."
        : "Corrija a formatação da célula para número inteiro sem espaços nem casas decimais.",
      registros:[valRegistro(r)]
    }));
  }
}

/* ================================================================
   REGRA 5 — HOMÔNIMOS: mesmo nome, GROOTs diferentes
   ================================================================ */
function valRegraHomonimos(todas, achados){
  const porNome = new Map();
  for(const r of todas){
    const n = valNorm(r.nome);
    if(!n || !valTemGroot(r.groot)) continue;
    if(!porNome.has(n)) porNome.set(n,new Map());
    porNome.get(n).set(valGroot(r.groot), r);
  }
  for(const [nome, mapa] of porNome){
    if(mapa.size < 2) continue;
    const regs = [...mapa.values()], ids = [...mapa.keys()];
    achados.push(valCriarAchado({
      id:"HM-"+nome.replace(/\s+/g,"_"), regra:"Homônimo ou cadastro duplicado",
      severidade:VAL_SEV.REVISAR, titulo:"Mesmo nome com GROOT IDs diferentes",
      groot: ids.join(" / "), nome: regs[0].nome, cargo: regs[0].cargo,
      aba: [...new Set(regs.map(r=>r.aba))].join(" + "),
      raciocinio:
        "\""+String(regs[0].nome).trim()+"\" aparece com "+ids.length+" identificadores diferentes ("
        + ids.join(", ")+"). Ou são duas pessoas de mesmo nome — o que acontece e é legítimo — ou a mesma "
        + "pessoa foi cadastrada duas vezes, e nesse caso pode estar sendo cobrada em dobro. "
        + "O nome sozinho não decide: compare cargo, período e matrícula antes de concluir.",
      sugestao:"Confirme no cadastro se são pessoas distintas. Sendo a mesma, unifique sob um único GROOT ID.",
      registros: regs.map(valRegistro)
    }));
  }
}

/* ================================================================
   REGRA 6 — DUPLICIDADE PESSOA-DIA EM DIARISTAS
   ================================================================ */
function valRegraDiariaDuplicada(diaristas, achados){
  const mapa = new Map();
  for(const d of diaristas){
    if(!valTemGroot(d.groot) || !(d.data instanceof Date)) continue;
    const k = valGroot(d.groot)+"|"+valDia(d.data);
    if(!mapa.has(k)) mapa.set(k,[]);
    mapa.get(k).push(d);
  }
  for(const [k,regs] of mapa){
    if(regs.length < 2) continue;
    const cargosDistintos = new Set(regs.map(r=>valNorm(r.cargo))).size > 1;
    achados.push(valCriarAchado({
      id:"DD-"+k, regra:"Diária duplicada", severidade:VAL_SEV.CRITICO,
      titulo:"Mesma pessoa com mais de uma diária no mesmo dia",
      groot: valGroot(regs[0].groot), nome: regs[0].nome, cargo: regs[0].cargo, aba:"DIARISTAS",
      datas: valFmt(regs[0].data), peso: regs.length,
      raciocinio:
        "O GROOT "+valGroot(regs[0].groot)+" ("+String(regs[0].nome).trim()+") tem "+regs.length
        + " diárias lançadas em "+valFmt(regs[0].data)+". "
        + "Uma pessoa trabalha no máximo uma diária por dia, então uma das linhas é repetição."
        + (cargosDistintos
            ? " As linhas têm tipos de diária diferentes ("
              + [...new Set(regs.map(r=>r.cargo))].join(", ")+"), o que pode indicar reclassificação "
              + "(dia útil virou domingo/feriado) em que a linha antiga não foi apagada."
            : " As linhas são do mesmo tipo, o que reforça a hipótese de lançamento repetido."),
      sugestao:"Mantenha uma única linha para o dia — a de tipo correto — e remova as demais.",
      registros: regs.map(valRegistro)
    }));
  }
}

/* ================================================================
   REGRA 7 — CONSISTÊNCIA ARITMÉTICA E DE DATAS
   ================================================================ */
function valRegraAritmetica(diaristas, achados){
  diaristas.forEach((d,idx) => {
    const q = Number(d.qtd), u = Number(d.unit), v = Number(d.valor);
    if(!isFinite(q) || !isFinite(u) || !isFinite(v)) return;
    const esperado = q*u;
    if(Math.abs(esperado - v) <= 0.01) return;
    achados.push(valCriarAchado({
      id:"AR-"+idx, regra:"Valor final incompatível", severidade:VAL_SEV.CRITICO,
      titulo:"VALOR FINAL diferente de QUANTIDADE × VALOR UNITÁRIO",
      groot: valGroot(d.groot), nome:d.nome, cargo:d.cargo, aba:"DIARISTAS",
      datas: valFmt(d.data), peso: Math.abs(v-esperado),
      raciocinio:
        "Nesta linha o VALOR FINAL não corresponde a QUANTIDADE × VALOR UNITÁRIO — a quantidade "
        + "lançada é "+q+" e o produto das duas colunas não bate com a terceira. Como o total da "
        + "fatura soma o VALOR FINAL, a diferença entra na cobrança do jeito que está. Costuma ser "
        + "fórmula sobrescrita por digitação manual. Confira as três células desta linha na planilha.",
      sugestao:"Restaure a fórmula da coluna VALOR FINAL (quantidade × valor unitário) nesta linha.",
      registros:[valRegistro(d)]
    }));
  });
}

function valRegraDatasInvertidas(labor, achados){
  labor.forEach((r,idx) => {
    if(!(r.ini instanceof Date) || !(r.fim instanceof Date)) return;
    if(valDia(r.fim) >= valDia(r.ini)) return;
    achados.push(valCriarAchado({
      id:"DI-"+idx, regra:"Datas invertidas", severidade:VAL_SEV.CRITICO,
      titulo:"DATA FIM anterior à DATA DE INÍCIO",
      groot: valGroot(r.groot), nome:r.nome, cargo:r.cargo, aba:"LABOR",
      datas: valFmt(r.ini)+"→"+valFmt(r.fim), lancamento: valSinal(r.valor),
      raciocinio:
        "O vínculo começa em "+valFmt(r.ini)+" e termina em "+valFmt(r.fim)+" — antes de começar. "
        + "O período é impossível, então qualquer cálculo de dias sobre esta linha sai errado, "
        + "inclusive o rateio que gera a cobrança dela.",
      sugestao:"Confira qual das duas datas está trocada e corrija antes de recalcular a linha.",
      registros:[valRegistro(r)]
    }));
  });
}

/* Datas fora do período faturado: a linha existe, mas cobra dias que
   não pertencem a esta competência. */
function valRegraForaDoPeriodo(diaristas, ctx, achados){
  if(!(ctx.ini instanceof Date) || !(ctx.fim instanceof Date)) return;
  const ini = valDia(ctx.ini), fim = valDia(ctx.fim);
  diaristas.forEach((d,idx) => {
    const dd = valDia(d.data);
    if(dd === null || (dd >= ini && dd <= fim)) return;
    achados.push(valCriarAchado({
      id:"FP-"+idx, regra:"Data fora do período", severidade:VAL_SEV.REVISAR,
      titulo:"Diária lançada fora do período faturado",
      groot: valGroot(d.groot), nome:d.nome, cargo:d.cargo, aba:"DIARISTAS",
      datas: valFmt(d.data), lancamento: valSinal(d.valor),
      raciocinio:
        "A diária de "+String(d.nome).trim()+" está lançada em "+valFmt(d.data)+", fora do período "
        + "faturado ("+valFmt(ctx.ini)+" a "+valFmt(ctx.fim)+"). Pode ser retroativo legítimo de uma "
        + "competência anterior que não foi cobrada, ou data digitada com mês errado — de um jeito "
        + "ou de outro a linha entra nesta fatura.",
      sugestao:"Confirme se é retroativo aprovado. Não sendo, corrija a data ou mova a linha para a competência correta.",
      registros:[valRegistro(d)]
    }));
  });
}

/* ================================================================
   REGRA 8 — QUALIDADE DE CADASTRO
   ================================================================ */
function valRegraCamposVazios(linhas, aba, achados){
  linhas.forEach((r,idx) => {
    const temValor = isFinite(Number(r.valor)) && Number(r.valor) !== 0;
    const semNome = !String(r.nome ?? "").trim();
    const semCargo = !String(r.cargo ?? "").trim();
    if(!temValor || (!semNome && !semCargo)) return;
    achados.push(valCriarAchado({
      id:"CV-"+aba+"-"+idx, regra:"Campo obrigatório vazio", severidade:VAL_SEV.REVISAR,
      titulo: semNome ? "Lançamento sem nome" : "Lançamento sem cargo",
      groot: valGroot(r.groot), nome:r.nome, cargo:r.cargo, aba,
      datas: r.data instanceof Date ? valFmt(r.data) : valFmt(r.ini)+"→"+valFmt(r.fim),
      lancamento: valSinal(r.valor),
      raciocinio:
        "A linha tem lançamento e está "
        + (semNome ? "sem nome" : "sem cargo")+(semNome && semCargo ? " e sem cargo" : "")+". "
        + "Uma linha que cobra precisa dizer por quem e a que título — sem isso ela não é conferível "
        + "nem por você nem pelo cliente.",
      sugestao:"Preencha o campo em branco ou remova a linha, se ela for resíduo do template.",
      registros:[valRegistro(r)]
    }));
  });
}

/* Grafias diferentes do mesmo regime ("Efetivo" e "efetivo") não mudam
   valor nenhum, mas quebram qualquer agrupamento por texto. */
function valRegraRegimeInconsistente(labor, achados){
  const mapa = new Map();
  for(const r of labor){
    const bruto = String(r.regime ?? "").trim();
    if(!bruto) continue;
    const chave = valRegimeNorm(bruto);
    if(!mapa.has(chave)) mapa.set(chave,new Map());
    const m = mapa.get(chave);
    m.set(bruto,(m.get(bruto)||0)+1);
  }
  for(const [chave,formas] of mapa){
    if(formas.size < 2) continue;
    const lista = [...formas.entries()].sort((a,b)=>b[1]-a[1]);
    achados.push(valCriarAchado({
      id:"RG-"+chave, regra:"Grafia inconsistente", severidade:VAL_SEV.CADASTRO,
      titulo:"Mesmo regime de contrato escrito de formas diferentes",
      groot:"", nome:"", cargo:"", aba:"LABOR",
      raciocinio:
        "O regime aparece como "+lista.map(([f,n]) => "\""+f+"\" ("+n+"x)").join(" e ")
        + ". É o mesmo vínculo escrito de jeitos diferentes: não altera valor, mas qualquer "
        + "agrupamento, filtro ou fórmula que compare texto vai tratar as duas formas como categorias "
        + "distintas — inclusive as deste app.",
      sugestao:"Padronize para a forma predominante: \""+lista[0][0]+"\".",
      registros:[]
    }));
  }
}

/* ================================================================
   REGRA 9 — VALOR UNITÁRIO DESTOANTE PARA O MESMO CARGO

   Compara cada linha com a MEDIANA do seu próprio cargo — média seria
   arrastada pelo próprio outlier que queremos achar.
   ================================================================ */
function valRegraTarifaDestoante(diaristas, achados){
  const porCargo = new Map();
  for(const d of diaristas){
    const u = Number(d.unit);
    if(!isFinite(u) || u <= 0) continue;
    const c = valNorm(d.cargo);
    if(!porCargo.has(c)) porCargo.set(c,[]);
    porCargo.get(c).push(d);
  }
  for(const [cargo,regs] of porCargo){
    if(regs.length < 5) continue;                       // amostra pequena não tem padrão
    const vals = regs.map(r=>Number(r.unit)).sort((a,b)=>a-b);
    const mediana = vals[Math.floor(vals.length/2)];
    if(!mediana) continue;
    for(const d of regs){
      const u = Number(d.unit), desvio = Math.abs(u-mediana)/mediana;
      if(desvio <= 0.10) continue;                      // 10% de folga
      achados.push(valCriarAchado({
        id:"TF-"+valGroot(d.groot)+"-"+valDia(d.data)+"-"+Math.round(u*100),
        regra:"Tarifa destoante", severidade:VAL_SEV.REVISAR,
        titulo:"Valor unitário fora do praticado no mesmo tipo de diária",
        groot: valGroot(d.groot), nome:d.nome, cargo:d.cargo, aba:"DIARISTAS",
        datas: valFmt(d.data), peso: desvio,
        raciocinio:
          "Esta diária de \""+d.cargo+"\" tem VALOR UNITÁRIO "+Math.round(desvio*100)+"% "
          + (u > mediana ? "acima" : "abaixo")+" da mediana das "+regs.length
          + " diárias do mesmo tipo nesta fatura. Pode ser tarifa nova entrando em vigor no meio do "
          + "período, ou célula digitada por cima da fórmula.",
        sugestao:"Confirme a tarifa vigente para a data. Sendo reajuste, marque como justificado.",
        registros:[valRegistro(d)]
      }));
    }
  }
}

/* ================================================================
   REGRA 10 — LINHAS QUASE IDÊNTICAS NO LABOR
   ================================================================ */
function valRegraLinhaDuplicada(labor, achados){
  const mapa = new Map();
  labor.forEach((r,idx) => {
    if(!valTemGroot(r.groot)) return;
    const k = [valGroot(r.groot), valDia(r.ini), valDia(r.fim), valNorm(r.cargo), valRegimeNorm(r.regime)].join("|");
    if(!mapa.has(k)) mapa.set(k,[]);
    mapa.get(k).push({ r, idx });
  });
  for(const [k,itens] of mapa){
    if(itens.length < 2) continue;
    const regs = itens.map(i=>i.r);
    achados.push(valCriarAchado({
      id:"LDup-"+k, regra:"Linha duplicada", severidade:VAL_SEV.CRITICO,
      titulo:"Linhas idênticas no LABOR",
      groot: valGroot(regs[0].groot), nome:regs[0].nome, cargo:regs[0].cargo, aba:"LABOR",
      datas: valFmt(regs[0].ini)+"→"+valFmt(regs[0].fim), peso: itens.length,
      raciocinio:
        itens.length+" linhas do LABOR trazem o mesmo GROOT, o mesmo período, o mesmo cargo e o mesmo "
        + "regime para "+String(regs[0].nome).trim()+". "
        + "Não há informação que as diferencie: uma delas é repetição, e a linha está sendo cobrada "
        + itens.length+" vezes.",
      sugestao:"Mantenha uma única linha e remova as repetições.",
      registros: regs.map(valRegistro)
    }));
  }
}

/* ================================================================
   REGRA 11 — HORA EXTRA SEM VÍNCULO NO LABOR
   ================================================================ */
function valRegraHoraExtraOrfa(horaExtra, labor, achados){
  if(!horaExtra || !horaExtra.length) return;
  const noLabor = new Set(labor.filter(r=>valTemGroot(r.groot)).map(r=>valGroot(r.groot)));
  horaExtra.forEach((h,idx) => {
    if(!valTemGroot(h.groot)) return;
    if(noLabor.has(valGroot(h.groot))) return;
    achados.push(valCriarAchado({
      id:"HE-"+idx, regra:"Hora extra sem vínculo", severidade:VAL_SEV.REVISAR,
      titulo:"Hora extra de quem não está no LABOR",
      groot: valGroot(h.groot), nome:h.nome, cargo:h.cargo, aba:"HORA EXTRA",
      datas:"—", lancamento: valSinal(h.valor),
      raciocinio:
        "O GROOT "+valGroot(h.groot)+" ("+String(h.nome).trim()+") tem hora extra lançada mas não aparece "
        + "em nenhuma linha do LABOR desta fatura. Ou a pessoa saiu do quadro e a hora extra é retroativa, "
        + "ou o vínculo ficou faltando no LABOR — nesse segundo caso a fatura está cobrando o extra de "
        + "alguém que ela própria não declara ter.",
      sugestao:"Confirme se a pessoa pertence ao período. Pertencendo, inclua a linha no LABOR.",
      registros:[valRegistro(h)]
    }));
  });
}

/* ================================================================
   ORQUESTRAÇÃO
   ================================================================ */
function auditarFatura(dados){
  const labor = dados.labor || [], diaristas = dados.diaristas || [], horaExtra = dados.horaExtra || [];
  const ctx = { ini: dados.periodo && dados.periodo.ini || null,
                fim: dados.periodo && dados.periodo.fim || null };
  const achados = [];

  valRegraGrootCompartilhado(labor, ctx, achados);
  valRegraLaborDiarista(labor, diaristas, ctx, achados);
  valRegraGrootAusente(labor, "LABOR", achados);
  valRegraGrootAusente(diaristas, "DIARISTAS", achados);
  valRegraGrootFormato([...labor, ...diaristas], achados);
  valRegraHomonimos([...labor, ...diaristas], achados);
  valRegraDiariaDuplicada(diaristas, achados);
  valRegraAritmetica(diaristas, achados);
  valRegraDatasInvertidas(labor, achados);
  valRegraForaDoPeriodo(diaristas, ctx, achados);
  valRegraCamposVazios(labor, "LABOR", achados);
  valRegraCamposVazios(diaristas, "DIARISTAS", achados);
  valRegraRegimeInconsistente(labor, achados);
  valRegraTarifaDestoante(diaristas, achados);
  valRegraLinhaDuplicada(labor, achados);
  valRegraHoraExtraOrfa(horaExtra, labor, achados);

  /* Crítico primeiro; dentro da severidade, o de maior peso — quantas
     linhas o achado envolve e quão longe do normal ele está. */
  achados.sort((a,b) => {
    const d = VAL_SEV_ORDEM.indexOf(a.severidade) - VAL_SEV_ORDEM.indexOf(b.severidade);
    if(d) return d;
    return (Number(b.peso)||0) - (Number(a.peso)||0);
  });

  const resumo = { critico:0, revisar:0, cadastro:0, info:0 };
  for(const a of achados) resumo[a.severidade]++;

  return {
    periodo: ctx,
    /* Contagem de linhas, e só. Somar os lançamentos daria um total
       monetário, que este app não escreve. */
    totais: { labor: labor.length, diaristas: diaristas.length },
    resumo, achados
  };
}

/* Node (testes) e navegador carregam o mesmo arquivo. */
if(typeof module !== "undefined" && module.exports){
  module.exports = { auditarFatura, valCompararNomes, valNorm, valGroot, valTemGroot,
                     valSobrepoe, valEncostados, valLev, VAL_SEV };
}
