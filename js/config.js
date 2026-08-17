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

   e cada operação usa um único horário para diarista. Levantado nas
   faturas de Julho/2026 (uma por unidade):

     unidade   operação            escala horário            turno (BASE TURNO)   diaristas
     SMG3      Pouso Alegre SVC    03:00 07:00 08:00 12:48   5 - Mista - 3º T     473 / 473
     BRXMG3    Pouso Alegre XD     13:00 17:00 18:00 22:45   11 - Mista - 3º T    110 / 110
     SMG5      Poços de Caldas     03:00 07:00 08:00 12:48   5 - Mista - 3º T      58 / 58
     SMG10     Divinópolis         01:30 05:00 06:00 09:50   65 - Mista - 3º T     31 / 37
     SMG11     Patos de Minas      00:00 05:00 06:00 09:20   54 - Mista - 3º T    109 / 109

   É um PADRÃO, não uma verdade de cada linha: em Divinópolis, 6 dos 37
   diaristas de julho saíram em outros horários (`05:30 12:00 13:00 15:18`
   e `11:00 14:00 15:00 20:00`, ambos "1º e 2º Turno"). Por isso a coluna
   é preenchida como sugestão e a exceção continua sendo conferida na
   fatura — o SIGO não informa horário.

   Varginha não tem valor porque não veio fatura da unidade. Fica em
   branco de propósito: chutar o horário de outra filial seria pior do
   que a lacuna, que ao menos se enxerga.
   ================================================================ */
const ESCALA_HORARIO_PADRAO = {
  "Pouso Alegre SVC": "03:00 07:00 08:00 12:48",
  "Pouso Alegre XD":  "13:00 17:00 18:00 22:45",
  "Poços de Caldas":  "03:00 07:00 08:00 12:48",
  "Varginha":         "",
  "Divinópolis":      "01:30 05:00 06:00 09:50",
  "Patos de Minas":   "00:00 05:00 06:00 09:20"
};

/** O horário padrão da operação, ou "" quando não há um levantado. */
function escalaHorarioDe(op){
  return Object.prototype.hasOwnProperty.call(ESCALA_HORARIO_PADRAO, op)
    ? ESCALA_HORARIO_PADRAO[op] : "";
}

if(typeof module!=="undefined"&&module.exports){
  module.exports={ESCALA_HORARIO_PADRAO, escalaHorarioDe};
}
