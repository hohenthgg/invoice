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
