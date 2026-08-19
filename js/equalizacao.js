/* MOTOR DE EQUALIZAÇÃO — o único que decide o que mexer no Labor
   ================================================================

   Este arquivo nasceu de dentro da Fusão de Linhas. A lógica é a que
   já rodava lá desde sempre; o que mudou foi o lugar. A Validação de
   Template precisava responder "estou acima do quadro que o cliente
   reconhece — quem eu tiro?", e a resposta certa não era escrever uma
   segunda heurística parecida: era chamar esta.

                       eqEqualizar()
                            │
                 ┌──────────┴──────────┐
                 ▼                     ▼
          Fusão de Linhas         Validação Template
        (alvo = retorno do        (alvo = QF do cliente,
         cliente, dia a dia)       o mesmo em todos os dias)

   Mesmo Labor + mesmo período + mesma curva alvo ⇒ mesmo plano. Não
   por convenção, mas porque é literalmente a mesma função.

   O EIXO DE DATAS

   As duas abas guardam data de jeitos diferentes — a Fusão usa `Date`
   UTC, a Validação usa inteiro AAAAMMDD. Um motor que aceitasse as
   duas teria dois caminhos, e dois caminhos divergem. Então o motor
   não conhece data nenhuma: ele trabalha num eixo de INTEIROS onde
   `d+1` é o dia seguinte — dias desde a época. As conversões moram
   aqui embaixo, e são as únicas.

   O QUE ELE FAZ

   Recebe a curva do quadro (quem está ativo em cada dia, com o rateio
   COM SINAL) e a curva alvo. Onde sobra gente, procura alterações no
   Labor que aproximem uma curva da outra SEM criar falta em nenhum
   outro dia — e é esse "nenhum outro dia" que separa este motor de um
   "escolha N nomes do dia 22". Quatro fases, nesta ordem:

     1 RETIRAR       a linha inteira cabe no excesso em TODOS os dias
                     em que ela está ativa
     2 ADIAR INÍCIO  o excesso está no começo do contrato
     3 PAUSAR        o excesso é um vale, e a pessoa volta a ser
                     necessária depois — fecha e reabre, mesma pessoa
     4 ENCURTAR FIM  o excesso está no fim, e o fim ORIGINAL já cai
                     dentro dele

   O que não couber em nenhuma das quatro sai como `__revisar`: excesso
   que só sumiria partindo um contrato de um jeito que o motor não faz
   sozinho. Falta (alvo acima do quadro) ele apenas relata — inventar
   pessoa não é trabalho dele.

   Dia com alvo 0 ou nulo é IGNORADO, não zerado: escala não publicada
   não é demanda zero, e tratá-la como zero mandaria o dia inteiro para
   o corte.

   E entrar na curva NÃO é o mesmo que poder ser mexido:

       entra na curva          toda linha, com o sinal do rateio
       pode virar candidato    só rateio > 0 e não `imutavel`

   A linha de rateio negativo é estorno: ela reduz de verdade o que a
   fatura cobra, e por isso conta na curva. Mas retirá-la SOBE a curva
   — resolveria o excesso ao contrário —, e adiar ou encurtar uma
   devolução mexe num acerto já feito. Nenhuma das quatro fases pode
   escolhê-la. `imutavel` marca o mesmo pelo outro motivo: ocupa vaga no
   dia mas não é linha do Labor — a diária já lançada na fatura.
   ================================================================ */
"use strict";

/* Fim do mundo para quem está com DATA FIM em branco. 2100-01-01. */
const EQ_INF = Math.round(Date.UTC(2100,0,1)/864e5);

/* As duas representações do projeto, e só elas, entram pelo eixo. */
const eqDeData   = d => Math.round(Date.UTC(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())/864e5);
const eqParaData = n => new Date(n*864e5);
const eqDeYmd    = v => Math.round(Date.UTC(Math.trunc(v/10000), Math.trunc(v/100)%100-1, v%100)/864e5);
const eqParaYmd  = n => { const d = eqParaData(n);
  return d.getUTCFullYear()*10000 + (d.getUTCMonth()+1)*100 + d.getUTCDate(); };

const EQ_ACAO = { RETIRAR:"retirar", ADIAR:"adiar", ENCURTAR:"encurtar", PAUSAR:"pausar" };

/** Que ação uma entrada do plano representa. A ordem importa: uma
 *  pessoa empurrada para fora do período pela fase 2 vira `retirar`. */
function eqTipoAcao(a){
  if(!a) return null;
  if(a.retirar) return EQ_ACAO.RETIRAR;
  if(a.novo_ini != null) return EQ_ACAO.ADIAR;
  if(a.novo_fim !== undefined) return EQ_ACAO.ENCURTAR;
  if(a.pausas && a.pausas.length) return EQ_ACAO.PAUSAR;
  return null;
}

/** Dias em que a pessoa está ativa DEPOIS do plano — usado por quem
 *  chama para medir o impacto de cada sugestão na curva.
 *  @param {{ini:number,fim:number|null}} p pessoa, no eixo inteiro
 *  @param {object|undefined} a ação do plano para ela */
function eqAtivoApos(p, a, dia){
  a = a || {};
  if(a.retirar) return false;
  const ini = a.novo_ini != null ? a.novo_ini : p.ini;
  const fim = a.novo_fim !== undefined ? a.novo_fim : (p.fim == null ? EQ_INF : p.fim);
  if(!(ini <= dia && dia <= fim)) return false;
  return !(a.pausas || []).some(pa => pa.fim < dia && dia < pa.ini);
}

/**
 * O motor.
 * @param {Array<{id:*,ini:number,fim:number|null,rateio:number,desempate:*,grupo:*,
 *                imutavel?:boolean}>} pessoas
 *        quem já está no Labor e conta no quadro. `fim` nulo = em aberto.
 *        A ORDEM importa: o desempate final da escolha é estável nela.
 * @param {Array<{dia:number,alvo:number|null,grupo:*}>} dias a curva alvo.
 * @param {{pausaDesde:number|null, permitirAdiarInicio:boolean}} [opcoes]
 * @returns {{acoes:Map, incluir:object}} `acoes` mapeia id → ação; `incluir`
 *        traz, por grupo, a falta por dia, e em `__revisar` o excesso que
 *        nenhuma das quatro fases resolveu.
 */
function eqEqualizar(pessoas, dias, opcoes){
  opcoes = opcoes || {};
  const permitirAdiarInicio = opcoes.permitirAdiarInicio !== false;
  const pausaDesde = (opcoes.pausaDesde === undefined) ? null : opcoes.pausaDesde;

  const acoes = new Map();
  const incluir = {};
  const grupos = [...new Set(dias.map(d => d.grupo))];

  for(const grupo of grupos){
    const doGrupo = dias.filter(d => d.grupo === grupo).slice().sort((a,b) => a.dia - b.dia);
    const dlist = doGrupo.map(d => d.dia), alvo = doGrupo.map(d => d.alvo);
    const pool  = pessoas.filter(p => p.grupo === grupo);

    const effIni = new Map(pool.map(p => [p.id, p.ini]));
    const effFim = new Map(pool.map(p => [p.id, p.fim == null ? EQ_INF : p.fim]));
    const act = (i,d) => effIni.get(i) <= d && d <= effFim.get(i);

    /* exc>0: sobra gente (cortar); exc<0: falta (relatar).
       Dia com alvo 0/nulo (escala não publicada, Go-Live) é ignorado. */
    const ignorado = dlist.map((d,k) => !(alvo[k] > 0));
    const exc = dlist.map((d,k) => ignorado[k] ? 0
      : (pool.reduce((s,p) => s + (act(p.id,d) ? p.rateio : 0), 0) - alvo[k]));
    const byI = new Map(pool.map(p => [p.id, p]));

    /* QUEM ENTRA NA CURVA × QUEM PODE SER MEXIDO — não é a mesma coisa.
       Toda linha entra na curva acima, inclusive a de rateio negativo:
       ela é estorno, e reduz de verdade o que a fatura cobra naqueles
       dias. Mas ela não pode ser CANDIDATA a nada. Retirar uma linha -1
       SOBE a curva em vez de baixá-la — resolveria o excesso ao
       contrário —, e adiar ou encurtar uma devolução mexe num acerto
       que já foi feito. Rateio 0 fica fora pelo mesmo caminho: a ação
       seria um no-op enfeitando o plano.

       `imutavel` é a mesma ideia por outro motivo: a pessoa conta na
       curva mas não é linha do Labor que se possa mexer. É o caso da
       DIÁRIA já lançada na fatura — ela ocupa uma vaga do dia, e por
       isso entra na conta, mas quem equaliza mexe no quadro fixo, não
       em diária que já aconteceu.

       A proteção vive AQUI, na origem da lista de candidatos, e não em
       cada fase nem em quem chama: as quatro fases (retirar, adiar
       início, pausar/retomar, antecipar fim) escolhem exclusivamente de
       `ordem`, então filtrar `ordem` cobre as quatro de uma vez e
       cobre também as duas abas. */
    const podeSerCandidato = p => p.rateio > 0 && !p.imutavel;
    /* Ordem de escolha: início mais recente primeiro, depois desempate
       decrescente. Quem entrou por último sai primeiro. */
    const ordem = pool.filter(podeSerCandidato).sort((a,b) =>
      (b.ini - a.ini) || String(b.desempate ?? "").localeCompare(String(a.desempate ?? ""))
    ).map(p => p.id);

    /* Fase 1 — RETIRAR: dia a dia, do mais antigo ao mais novo. Em cada dia com
       excesso, escolhe candidatos cujo intervalo de dias ativos está INTEIRAMENTE
       dentro do excesso disponível em cada um desses dias (verificado contra o exc
       atual, nunca deixando nenhum dia ir a negativo). */
    for(let k=0;k<dlist.length;k++){
      while(exc[k] > 0){
        let best = null, bg = -1;
        for(const i of ordem){
          if(acoes.has(i) || !act(i,dlist[k])) continue;
          const ad = dlist.map((d,kk) => act(i,d) ? kk : -1).filter(kk => kk >= 0);
          if(ad.some(kk => exc[kk] < byI.get(i).rateio)) continue;  // não cabe em algum dia ativo
          if(ad.length > bg){ bg = ad.length; best = i; }
        }
        if(best == null) break;
        dlist.forEach((d,kk) => { if(act(best,d)) exc[kk] -= byI.get(best).rateio; });
        acoes.set(best,{ retirar:true });
      }
    }

    /* Fase 2 — ADIAR INÍCIO: corta excesso no começo do contrato (desabilitável). */
    if(permitirAdiarInicio){
      for(let k=0;k<dlist.length;k++){
        while(exc[k] > 0){
          const i = ordem.find(i => !acoes.has(i) && effIni.get(i) === dlist[k] && act(i,dlist[k]));
          if(i == null) break;
          let j = k;
          while(j < dlist.length && exc[j] > 0 && act(i,dlist[j])) j++;
          const novo = j < dlist.length ? dlist[j] : effFim.get(i) + 1;
          for(let m=k;m<j;m++) exc[m] -= byI.get(i).rateio;
          if(j < dlist.length){ acoes.set(i,{ novo_ini:novo }); effIni.set(i,novo); }
          else { acoes.set(i,{ retirar:true }); effIni.set(i,EQ_INF); } // empurrado para fora = retirar
        }
      }
    }

    /* Fase 3 — PAUSAR/RETOMAR: vale (excesso cercado de dias sem excesso, com
       retomada de demanda depois) NÃO é tratado com corte definitivo — fecha o
       contrato no início do vale e reabre no fim, preservando a identidade da
       pessoa (mesma linha, dois períodos). */
    const pausado = (i,d) => { const a = acoes.get(i); if(!a || !a.pausas) return false;
      return a.pausas.some(p => p.fim < d && d < p.ini); };
    const ativoP = (i,d) => act(i,d) && !pausado(i,d);
    for(let k=0;k<dlist.length;k++){
      while(exc[k] > 0){
        if(pausaDesde == null || dlist[k] < pausaDesde) break;  // sem data, ou antes dela: não pausa
        const candidatos = ordem.filter(i => !acoes.get(i)?.retirar && !acoes.get(i)?.novo_ini
          && !acoes.get(i)?.novo_fim && ativoP(i,dlist[k]));
        let aplicado = false;
        for(const i of candidatos){
          let j = k;
          while(j+1 < dlist.length && exc[j+1] > 0 && ativoP(i,dlist[j+1])) j++;
          /* Só é vale se o contrato seguir ativo DEPOIS do excesso E a demanda no
             próximo dia já não estiver em excesso — ou seja, há retomada de verdade,
             não um fim disfarçado. */
          const temRetomada = effFim.get(i) > dlist[j] && j+1 < dlist.length && !ignorado[j+1];
          if(!temRetomada) continue;
          for(let m=k;m<=j;m++) exc[m] -= byI.get(i).rateio;
          const a = acoes.get(i) || {};
          a.pausas = a.pausas || [];
          a.pausas.push({ fim:dlist[k]-1, ini:dlist[j]+1 });
          acoes.set(i,a);
          aplicado = true;
          break;
        }
        if(!aplicado) break;   // nenhum candidato serve para pausa — vai para a fase 4
      }
    }

    /* Fase 4 — ENCURTAR FIM: excesso residual no fim do contrato (sem retomada).
       Só aceita candidato cujo FIM ORIGINAL já cai dentro do trecho de excesso —
       nunca corta quem teria dias futuros SEM excesso, o que criaria falta. */
    for(let k=dlist.length-1;k>=0;k--){
      while(exc[k] > 0){
        const candidatos = ordem.filter(i => !acoes.has(i) && act(i,dlist[k]));
        let aplicado = false;
        for(const i of candidatos){
          let j = k;
          while(j-1 >= 0 && exc[j-1] >= byI.get(i).rateio && act(i,dlist[j-1])) j--;
          const novoFim = dlist[j] - 1;
          if(effFim.get(i) > dlist[k]) continue;  // termina depois do trecho — corte inseguro
          for(let m=j;m<=k;m++) exc[m] -= byI.get(i).rateio;
          if(novoFim < effIni.get(i)){ acoes.set(i,{ retirar:true }); }
          else { acoes.set(i,{ novo_fim:novoFim }); effFim.set(i,novoFim); }
          aplicado = true;
          break;
        }
        if(!aplicado) break;
      }
    }

    const res = {}, revisar = {};
    dlist.forEach((d,k) => { if(exc[k] < 0) res[d] = -exc[k]; else if(exc[k] > 0) revisar[d] = exc[k]; });
    if(Object.keys(res).length) incluir[grupo] = { dias:res };
    if(Object.keys(revisar).length) (incluir.__revisar = incluir.__revisar || {})[grupo] = revisar;
  }
  return { acoes, incluir };
}

if(typeof module !== "undefined" && module.exports){
  module.exports = { eqEqualizar, eqTipoAcao, eqAtivoApos,
                     eqDeData, eqParaData, eqDeYmd, eqParaYmd, EQ_INF, EQ_ACAO };
}
