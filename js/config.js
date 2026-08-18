/* Ajustes MELI — constantes de configuração */
"use strict";

/* ================================================================
   CONSTANTES
   ================================================================ */
const SHEET_NAME = "Labor enviado ao MELI";
const RETORNO_SHEET = "Retorno MELI";
const CUT_DAY = 15;
const MONTHS = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const MELI_HEADERS = ["GROOT ID","NOME","MATRICULA","REGIME DE CONTRATO","CARGO","DATA DE INÍCIO","DATA FIM","% RATEIO","DIAS TRABALHADOS X FOLGA","ESCALA"];
const COLS = {
  groot:["GROOT ID"], nome:["NOME"], matricula:["MATRICULA","MATRÍCULA"],
  regime:["REGIME DE CONTRATO"], cargo:["CARGO"],
  inicio:["DATA DE INÍCIO","DATA DE INICIO"], fim:["DATA FIM"], rateio:["% RATEIO","RATEIO"],
  diasFolga:["DIAS TRABALHADOS X FOLGA"], escala:["ESCALA"]
};

/* ================================================================
   ESCALA HORÁRIO PADRÃO DO DIARISTA, POR OPERAÇÃO
   ----------------------------------------------------------------
   Na fatura 3PL, a aba DIARISTAS traz uma coluna ESCALA cujo valor é o
   horário no formato

       entrada · início do intervalo · fim do intervalo · saída

   e cada operação tem um horário dominante para diarista. Levantado nas
   seis faturas de Julho/2026, uma por unidade:

     unidade   operação            escala horário            turno (BASE TURNO)   no padrão
     SMG3      Pouso Alegre SVC    03:00 07:00 08:00 12:48   5 - Mista - 3º T     473 / 473
     BRXMG3    Pouso Alegre XD     13:00 17:00 18:00 22:45   11 - Mista - 3º T    110 / 110
     SMG5      Poços de Caldas     03:00 07:00 08:00 12:48   5 - Mista - 3º T      58 / 58
     SMG9      Varginha            03:00 07:00 08:00 11:20   4 - Mista - 3º T     108 / 138
     SMG10     Divinópolis         01:30 05:00 06:00 09:50   65 - Mista - 3º T     31 / 37
     SMG11     Patos de Minas      00:00 05:00 06:00 09:20   54 - Mista - 3º T    109 / 109

   É um PADRÃO, não uma verdade de cada linha. Quatro unidades são
   uniformes; duas não:

     Varginha     30 de 138 em `10:00 15:00 16:00 19:48` ("1º e 2º Turno")
     Divinópolis   6 de  37 em `05:30 12:00 13:00 15:18` e
                              `11:00 14:00 15:00 20:00` (idem)

   A exceção acompanha em parte o cargo "Diarista Dom. e Feriados", mas
   NÃO dá para derivar dele: em Pouso Alegre e Patos de Minas esse mesmo
   cargo usa o horário normal da unidade.
   ================================================================ */
const ESCALA_HORARIO_PADRAO = {
  "Pouso Alegre SVC": "03:00 07:00 08:00 12:48",
  "Pouso Alegre XD":  "13:00 17:00 18:00 22:45",
  "Poços de Caldas":  "03:00 07:00 08:00 12:48",
  "Varginha":         "03:00 07:00 08:00 11:20",
  "Divinópolis":      "01:30 05:00 06:00 09:50",
  "Patos de Minas":   "00:00 05:00 06:00 09:20"
};

/** O horário padrão da operação, ou "" quando não há um levantado. */
function escalaHorarioDe(op){
  return Object.prototype.hasOwnProperty.call(ESCALA_HORARIO_PADRAO, op)
    ? ESCALA_HORARIO_PADRAO[op] : "";
}

/* ================================================================
   O QUE A COLUNA ESCALA DO SIGO REALMENTE CARREGA, POR OPERAÇÃO
   ----------------------------------------------------------------
   O padrão acima só entra quando a linha não diz nada. Mas o SIGO real
   diz bastante — só que cada filial fala uma língua:

     Divinópolis      o HORÁRIO completo, por linha ("01:00 04:00 05:00 09:20")
     Patos de Minas   horário em outro formato ("00:30 as 09:18") — ou vazio
     Varginha         turno: AM / PM
     Poços de Caldas  turno: AM (761), PM (10), SD (17)
     Pouso Alegre SVC turno: svc / SVC / xd / XD / sd — e uma coluna
                      "Escala Horário" própria, hoje vazia
     Pouso Alegre XD  turno: XD, SD, FULL

   Horário escrito passa verbatim — inclusive o formato de Patos: é o que
   consta na origem. Turno vira horário pela tabela abaixo, levantada nas
   faturas 3PL (Varginha: AM = 4-Mista-3ºT, PM = "1º e 2º Turno" da fatura
   SMG9 de julho; Pouso: svc = SMG3, xd = BRXMG3).

   Token que NÃO está na tabela (SD, FULL, PM de Poços…) fica como está e
   vira tópico de revisão: escrever o horário de outro turno no lugar
   seria pôr um dado errado com cara de certo.
   ================================================================ */
const ESCALA_TOKEN_HORARIO = {
  "Pouso Alegre SVC": { "SVC":"03:00 07:00 08:00 12:48", "XD":"13:00 17:00 18:00 22:45" },
  "Pouso Alegre XD":  { "XD":"13:00 17:00 18:00 22:45" },
  "Poços de Caldas":  { "AM":"03:00 07:00 08:00 12:48" },
  "Varginha":         { "AM":"03:00 07:00 08:00 11:20", "PM":"10:00 15:00 16:00 19:48" },
  "Divinópolis":      {},
  "Patos de Minas":   {}
};

/* Quais turnos PERTENCEM a cada operação — regra de negócio, não de horário.
   O SVC mistura-se com SD e FULL (e registra o turno XD quando acontece);
   o XD é APENAS XD. Um SD ou FULL na aba XD não é "turno sem horário": é um
   registro que não pertence à operação — foi lançado na aba errada ou o
   turno foi digitado errado. Só as operações com regra declarada policiam;
   nas demais, token desconhecido continua sendo apenas "sem horário". */
const ESCALA_TURNOS_DA_OPERACAO = {
  "Pouso Alegre SVC": ["SVC","SD","FULL","XD"],
  "Pouso Alegre XD":  ["XD"]
};

/** A célula já é um horário? Duas ou mais marcações HH:MM contam como sim —
 *  cobre "01:00 04:00 05:00 09:20" e também "00:30 as 09:18". */
function pareceHorario(v){
  const m = String(v == null ? "" : v).match(/\d{1,2}[:h]\d{2}/g);
  return !!m && m.length >= 2;
}

/** Resolve o que a coluna ESCALA da saída deve dizer para uma linha.
 *
 *  Prioridade: coluna explícita de horário → horário escrito na própria
 *  ESCALA → turno mapeado → vazio vira o padrão da operação.
 *
 *  Duas saídas de exceção, que o painel aponta em vez de esconder:
 *    · `intruso`  — turno que não pertence à operação (SD numa aba que só
 *                   admite XD). Não é falta de horário: é registro na aba
 *                   errada, e converter seria carimbar o erro de válido.
 *    · `sem_mapa` — turno legítimo, mas sem horário levantado ainda.
 *  Nos dois casos o valor original é preservado.
 *
 *  @returns {{valor:string, origem:string, token?:string}} */
function resolverEscala(op, escalaBruta, horarioExplicito){
  const hx = String(horarioExplicito == null ? "" : horarioExplicito).trim();
  if(hx) return {valor:hx, origem:"coluna"};
  const e = String(escalaBruta == null ? "" : escalaBruta).trim();
  if(pareceHorario(e)) return {valor:e, origem:"arquivo"};
  if(e){
    const token = e.toUpperCase();
    const permitidos = ESCALA_TURNOS_DA_OPERACAO[op];
    if(permitidos && permitidos.indexOf(token) < 0)
      return {valor:e, origem:"intruso", token:e};
    const mapa = ESCALA_TOKEN_HORARIO[op] || {};
    if(Object.prototype.hasOwnProperty.call(mapa, token))
      return {valor:mapa[token], origem:"token", token};
    return {valor:e, origem:"sem_mapa", token:e};
  }
  return {valor:escalaHorarioDe(op), origem:"padrao"};
}

if(typeof module!=="undefined"&&module.exports){
  module.exports={ESCALA_HORARIO_PADRAO, escalaHorarioDe, ESCALA_TOKEN_HORARIO,
                  ESCALA_TURNOS_DA_OPERACAO, pareceHorario, resolverEscala};
}
