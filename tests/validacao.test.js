/* Auditoria de fatura no template padrão — js/validacao.js
   ================================================================

   O que estes testes protegem:

       A DISTINÇÃO ENTRE ERRO E MOVIMENTAÇÃO LEGÍTIMA

   Duas linhas com o mesmo GROOT ID podem ser a efetivação de uma
   pessoa (rotina) ou duas pessoas cobrando pelo mesmo identificador
   (erro grave). Os campos preenchidos são os mesmos; o que separa os
   casos é a COMBINAÇÃO de nome, período, vínculo e sinal do valor.

   Um módulo que erre esse julgamento falha nos dois sentidos, e os
   dois custam caro: marcar toda efetivação como crítica ensina o
   auditor a ignorar o painel, e deixar passar duas pessoas no mesmo
   GROOT deixa a duplicidade entrar na fatura.

   Os casos abaixo reproduzem a ESTRUTURA de uma fatura real — as
   mesmas relações entre nomes, os mesmos formatos de período e os
   mesmos sinais de valor que motivaram a aba. Nomes e identificadores
   são fictícios: dado de pessoa não vive em repositório.
   ================================================================ */
"use strict";
const { auditarFatura, valCompararNomes, valGroot, valTemGroot,
        valSobrepoe, valEncostados, VAL_CATEGORIAS } = require("../js/validacao.js");

let pass = 0, fail = 0;
function check(ok, label, extra){
  if(ok){ pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (extra ? "  → " + extra : "")); }
}
const d = (a,m,dia) => new Date(Date.UTC(a,m-1,dia));
const PERIODO = { ini:d(2026,7,16), fim:d(2026,8,15) };
const L = o => Object.assign({ aba:"LABOR", linha:0, groot:"", nome:"", regime:"Efetivo",
  cargo:"Auxiliar de Apoio LOG I", ini:null, fim:null, qtd:1, unit:0, valor:0 }, o);
const D = o => Object.assign({ aba:"DIARISTAS", linha:0, groot:"", nome:"", cargo:"Diarista",
  data:null, qtd:1, unit:366.04, valor:366.04 }, o);
const rodar = (labor, diaristas, horaExtra) =>
  auditarFatura({ labor:labor||[], diaristas:diaristas||[], horaExtra:horaExtra||[], periodo:PERIODO });
const acharRegra = (r,regra) => r.achados.filter(a => a.regra === regra);

/* ================================================================ */
console.log("\nComparação de nomes — a base de todo o resto");

/* Token repetido não pode inflar a contenção: "LORRANY SILVA SANTOS
   SILVA" repete SILVA, e contar a repetição fazia "JULIANA DA SILVA" —
   dois tokens, ambos presentes em Lorrany — passar por variação. Gente
   diferente virando variação é o pior erro deste comparador. */
check(valCompararNomes("LORRANY SILVA SANTOS SILVA","JULIANA DA SILVA").relacao !== "variacao",
      "sobrenome repetido não faz outra pessoa virar variação");
check(valCompararNomes("LORRANY SILVA SANTOS SILVA","MARCELA SILVA DOS SANTOS").relacao !== "variacao",
      "dois sobrenomes em comum sem contenção real seguem sendo outra pessoa");
check(valCompararNomes("MARIANA COSTA DO AMARAL SANTOS","MARIANA COSTA DO AMARAL").relacao === "variacao",
      "um sobrenome a mais é a mesma pessoa, não outra");
check(valCompararNomes("SIMONE APARECIDA CARDOSO","SIMONE PARECIDA CARDOSO").relacao === "variacao",
      "erro de digitação em um sobrenome é a mesma pessoa");
check(valCompararNomes("BRUNO TEIXEIRA LOPES","SERGIO ALMEIDA CUNHA").relacao === "distintos",
      "nomes sem nada em comum são pessoas diferentes");
check(valCompararNomes("PAULO CESAR BRASELINO DE SOUZA","PAULO CÉSAR DA COSTA SILVA").relacao === "parcial",
      "primeiro nome igual e sobrenomes diferentes fica em 'parcial', sem concluir");
check(valCompararNomes("MARIA SILVA","MARIA SOUZA").relacao === "parcial",
      "um primeiro nome comum sozinho não une duas pessoas");
/* Uma letra de diferença é ruído num nome com corpo e é o nome inteiro
   num nome de uma letra. Sem piso de tamanho, "A ALPHA" e "B BETA"
   virariam "parcialmente coincidentes" — dois nomes sem nada em comum. */
check(valCompararNomes("A ALPHA","B BETA").relacao === "distintos",
      "iniciais soltas não contam como primeiro nome parecido");
check(valCompararNomes("ADRIEL SOUZA","ADRIELE BATISTA").relacao === "parcial",
      "…mas ADRIEL × ADRIELE continua sendo coincidência parcial");
check(valCompararNomes("JOSE DA SILVA","JOSÉ DA SILVA").relacao === "iguais",
      "acento não cria uma segunda pessoa");
/* Sem o piso de 2 tokens em comum, "ANA" e "ANA PAULA MOREIRA DOS REIS"
   seriam a mesma pessoa por contenção — e toda Ana da planilha viraria uma só. */
check(valCompararNomes("ANA","ANA PAULA GONZAGA").relacao !== "variacao",
      "contenção de um único nome não basta para unir registros");

console.log("\nIdentificadores");
check(valGroot(2110046) === "2110046" && valGroot("2110046.0") === "2110046",
      "número, texto e decimal do Excel viram o mesmo identificador");
check(valTemGroot("0") === false && valTemGroot("") === false && valTemGroot("2110046") === true,
      "zero e vazio não são identificador");

console.log("\nPeríodos");
check(valEncostados({fim:d(2026,8,10)},{ini:d(2026,8,11)}) === true,
      "um termina dia 10 e o outro começa dia 11: encostados");
check(valSobrepoe({ini:d(2026,3,23),fim:null},{ini:d(2026,6,3),fim:d(2026,8,4)}, PERIODO.fim) === true,
      "vínculo em aberto sobrepõe o que corre dentro do período faturado");
check(valSobrepoe({ini:d(2026,3,1),fim:d(2026,5,1)},{ini:d(2026,6,3),fim:d(2026,8,4)}, PERIODO.fim) === false,
      "períodos separados não se sobrepõem");

/* ================================================================
   O par de casos que dá nome ao módulo
   ================================================================ */
console.log("\nMesmo GROOT ID: efetivação × duas pessoas");

const efetivacao = rodar([
  L({ groot:"2110045", nome:"MARIANA COSTA DO AMARAL SANTOS", regime:"Temporário",
      ini:d(2026,3,30), fim:d(2026,8,10), valor:1923.04 }),
  L({ groot:"2110045", nome:"MARIANA COSTA DO AMARAL", regime:"Efetivo",
      ini:d(2026,8,11), fim:null, valor:4038.38 })
]);
const efe = acharRegra(efetivacao,"Mudança de vínculo");
check(efe.length === 1 && efe[0].severidade === "info",
      "Temporário→Efetivo encostado, mesma pessoa: informativo, não erro",
      JSON.stringify(efetivacao.achados.map(a => a.severidade+":"+a.regra)));
check(efe.length === 1 && /consolide/i.test(efe[0].sugestao) && /DATA FIM/.test(efe[0].sugestao),
      "…e a sugestão é consolidar apagando a DATA FIM");
check(efetivacao.resumo.critico === 0,
      "uma efetivação normal não gera nenhum crítico");

const duasPessoas = rodar([
  L({ groot:"2110046", nome:"BRUNO TEIXEIRA LOPES", regime:"Efetivo",
      cargo:"Líder de Operação", ini:d(2026,3,23), fim:null, valor:10303.25 }),
  L({ groot:"2110046", nome:"SERGIO ALMEIDA CUNHA", regime:"Temporário",
      cargo:"Auxiliar de Apoio LOG I", ini:d(2026,6,3), fim:d(2026,8,4), valor:795.14 })
]);
const dois = acharRegra(duasPessoas,"GROOT ID com pessoas diferentes");
check(dois.length === 1 && dois[0].severidade === "critico",
      "mesmo GROOT, nomes sem relação, períodos sobrepostos e ambos positivos: crítico");
check(dois.length === 1 && /Líder de Operação/.test(dois[0].raciocinio)
      && /Auxiliar de Apoio LOG I/.test(dois[0].raciocinio),
      "o raciocínio cita os DOIS cargos — foi o que o usuário pediu para nunca ficar de fora");
check(dois.length === 1 && /sobrep/i.test(dois[0].raciocinio) && /positivos/.test(dois[0].raciocinio),
      "…e diz explicitamente que os períodos se sobrepõem e as duas linhas cobram");

/* O mesmo par, agora SEM sobreposição: o identificador pode ter sido
   reaproveitado depois do desligamento. Continua irregular, mas não é
   cobrança simultânea — e a severidade tem de cair. */
const reaproveitado = rodar([
  L({ groot:"9000001", nome:"ANTONIO PEREIRA LIMA", ini:d(2026,1,10), fim:d(2026,4,30), valor:5000 }),
  L({ groot:"9000001", nome:"BEATRIZ ROCHA CAMPOS", ini:d(2026,6,1), fim:null, valor:5000 })
]);
const reap = acharRegra(reaproveitado,"GROOT ID com pessoas diferentes");
check(reap.length === 1 && reap[0].severidade === "revisar",
      "pessoas diferentes sem sobreposição cai para Revisar — não há cobrança simultânea",
      reap[0] && reap[0].severidade);
check(reap.length === 1 && /reaproveitado/.test(reap[0].raciocinio),
      "…e o raciocínio levanta a hipótese do identificador reaproveitado");

/* Um estorno em curso é indício de conserto, não de fraude: a
   severidade cai mesmo com sobreposição. */
const comEstorno = rodar([
  L({ groot:"2110049", nome:"ADRIEL SOUZA MEIRELES DA SILVA", ini:d(2026,6,18), fim:null, valor:5961.41 }),
  L({ groot:"2110049", nome:"ADRIELE WALLACE DO NASCIMENTO BATISTA", ini:d(2026,7,21), fim:d(2026,7,31), valor:-2186.64 })
]);
/* Os nomes têm o primeiro nome parecido e sobrenomes sem relação: isso
   é "parcialmente coincidente", que é uma regra própria — a confiança
   nela é menor que a de "sem relação nenhuma". */
const est = acharRegra(comEstorno,"GROOT ID com nomes parcialmente coincidentes");
check(est.length === 1 && est[0].severidade === "revisar",
      "com um dos lançamentos negativo, a sobreposição vira Revisar e não Crítico",
      est[0] && est[0].severidade);
check(est.length === 1 && /estorno/i.test(est[0].raciocinio),
      "…e o raciocínio nomeia o estorno como hipótese");

/* Mesma pessoa, mas com os dois vínculos abertos ao mesmo tempo:
   aqui a duplicidade é real. */
const mesmaPessoaSobreposta = rodar([
  L({ groot:"9000002", nome:"CARLA MENDES DE SOUZA", ini:d(2026,5,1), fim:null, valor:5000 }),
  L({ groot:"9000002", nome:"CARLA MENDES DE SOUZA", ini:d(2026,7,1), fim:null, valor:5000 })
]);
const sob = acharRegra(mesmaPessoaSobreposta,"Períodos sobrepostos");
check(sob.length === 1 && sob[0].severidade === "critico",
      "a MESMA pessoa com dois vínculos ativos ao mesmo tempo também é crítico");
check(sob.length === 1 && /dobro/.test(sob[0].raciocinio),
      "…porque os dias sobrepostos estão sendo cobrados em dobro");

/* Nome diferente sem sobreposição e sem transição de vínculo: é só
   qualidade de cadastro. */
const soGrafia = rodar([
  L({ groot:"9000003", nome:"SIMONE APARECIDA CARDOSO", regime:"Efetivo", ini:d(2026,1,1), fim:d(2026,3,31), valor:100 }),
  L({ groot:"9000003", nome:"SIMONE PARECIDA CARDOSO", regime:"Efetivo", ini:d(2026,6,1), fim:d(2026,7,1), valor:100 })
]);
check(acharRegra(soGrafia,"Variação cadastral de nome").length === 1,
      "grafias diferentes da mesma pessoa, sem sobreposição, ficam em Cadastro");

/* ================================================================
   LABOR × DIARISTA — o sinal do valor decide
   ================================================================ */
console.log("\nLABOR × DIARISTA");

const dupla = rodar(
  [L({ groot:"3000001", nome:"PEDRO ALVES", ini:d(2026,8,1), fim:d(2026,8,15), valor:5000 })],
  [D({ groot:"3000001", nome:"PEDRO ALVES", data:d(2026,8,10) })]
);
const dup = acharRegra(dupla,"LABOR × DIARISTA");
check(dup.length === 1 && dup[0].severidade === "critico",
      "LABOR positivo com diária dentro do período: dupla cobrança, crítico");

const estornado = rodar(
  [L({ groot:"2110048", nome:"DIEGO VIANA PEREIRA", ini:d(2026,7,27), fim:d(2026,7,31), valor:-961.52 })],
  [D({ groot:"2110048", nome:"DIEGO VIANA PEREIRA", data:d(2026,7,29) }),
   D({ groot:"2110048", nome:"DIEGO VIANA PEREIRA", data:d(2026,7,31) })]
);
const ret = acharRegra(estornado,"LABOR × DIARISTA");
check(ret.length === 1 && ret[0].severidade === "revisar",
      "o MESMO cruzamento com LABOR negativo cai para Revisar",
      ret[0] && ret[0].severidade);
check(ret.length === 1 && /estorno|retroativ/i.test(ret[0].raciocinio),
      "…e o raciocínio explica que negativo costuma ser estorno ou ajuste retroativo");
check(ret.length === 1 && ret[0].opcoes.some(o => /ajuste retroativo/i.test(o)),
      "…oferecendo marcar como ajuste retroativo em vez de decidir sozinho");

/* Vínculo em aberto vale até o fim do período faturado — é assim que a
   fatura o cobra, então é assim que a sobreposição tem de ser medida. */
const abertoComDiaria = rodar(
  [L({ groot:"3000002", nome:"LUCIA RAMOS", ini:d(2026,7,20), fim:null, valor:5000 })],
  [D({ groot:"3000002", nome:"LUCIA RAMOS", data:d(2026,8,5) })]
);
check(acharRegra(abertoComDiaria,"LABOR × DIARISTA").length === 1,
      "vínculo sem DATA FIM alcança a diária de agosto dentro do período");

/* Diária FORA do vínculo não é conflito nenhum. */
const semConflito = rodar(
  [L({ groot:"3000003", nome:"MARCOS DIAS", ini:d(2026,8,1), fim:d(2026,8,5), valor:5000 })],
  [D({ groot:"3000003", nome:"MARCOS DIAS", data:d(2026,7,20) })]
);
check(acharRegra(semConflito,"LABOR × DIARISTA").length === 0,
      "diária anterior ao vínculo não é dupla cobrança — nada é apontado");

/* ================================================================
   GROOT ausente e fora do padrão
   ================================================================ */
console.log("\nGROOT ID ausente e fora do padrão");

const semGroot = rodar([
  L({ groot:"", nome:"LUZIA BARBOSA", cargo:"Auxiliar de Apoio LOG I", ini:d(2026,7,28), fim:d(2026,7,31), valor:-769.21 }),
  L({ groot:"", nome:"HELENA COSTA", cargo:"Supervisor de Operação", ini:d(2026,7,1), fim:null, valor:9000 })
]);
const aus = acharRegra(semGroot,"GROOT ID ausente");
check(aus.length === 2, "toda linha sem GROOT vira achado", "n="+aus.length);
check(aus.some(a => /LUZIA/.test(a.nome) && a.severidade === "critico"),
      "cargo operacional (auxiliar) sem GROOT é crítico");
check(aus.some(a => /HELENA/.test(a.nome) && a.severidade === "revisar"),
      "cargo não operacional sem GROOT pesa menos e fica em Revisar");
check(aus.every(a => a.raciocinio.includes(a.cargo)),
      "o cargo aparece no raciocínio dos dois — é o que diz se a ausência é grave");

/* O padrão não é declarado: é aprendido da própria planilha. */
const padrao = rodar([
  ...Array.from({length:20},(_,i) => L({ groot:String(2400000+i), nome:"PESSOA "+i, ini:d(2026,7,20), valor:100 })),
  L({ groot:"110052", nome:"OSVALDO PINHEIRO GOMES", ini:d(2026,7,20), valor:100 })
]);
const fp = acharRegra(padrao,"GROOT ID fora do padrão");
check(fp.length === 1 && fp[0].groot === "110052",
      "num universo de 7 dígitos, o de 6 é apontado — e só ele", "n="+fp.length);
check(fp.length === 1 && fp[0].severidade === "cadastro" && /não significa que esteja errado/.test(fp[0].raciocinio),
      "…sem afirmar que está errado: pode ser identificador legado");
check(fp.length === 1 && /7 dígitos/.test(fp[0].raciocinio) && /%/.test(fp[0].raciocinio),
      "…citando qual é o padrão e que fatia da planilha ele cobre");

/* A linha de absenteísmo do template usa "ABS" no lugar do GROOT: é
   marcador do modelo, não cadastro incompleto. */
const comAbs = rodar([
  L({ groot:"ABS", nome:" ABSENTEISMO", cargo:" ABSENTEISMO", regime:"FIXO/CLT", valor:0 }),
  L({ groot:"2110051", nome:"TEREZA NOGUEIRA", ini:d(2026,7,20), valor:100 })
]);
check(comAbs.achados.every(a => a.groot !== "ABS"),
      "a linha ABS do template não vira achado de cadastro");

const formatoSujo = rodar([
  ...Array.from({length:10},(_,i) => L({ groot:String(2400000+i), nome:"P"+i, ini:d(2026,7,20), valor:100 })),
  L({ groot:"24 00099", nome:"COM ESPACO", ini:d(2026,7,20), valor:100 })
]);
check(acharRegra(formatoSujo,"GROOT ID fora do padrão").some(a => /espaço/.test(a.raciocinio)),
      "espaço no meio do identificador é apontado com esse nome");

/* ================================================================
   Demais regras
   ================================================================ */
console.log("\nDuplicidades, aritmética e datas");

const diariaDupla = rodar([], [
  D({ groot:"4000001", nome:"JOANA REIS", data:d(2026,8,3) }),
  D({ groot:"4000001", nome:"JOANA REIS", data:d(2026,8,3), cargo:"Diarista Dom. e Feriados", valor:433.46 })
]);
const dd = acharRegra(diariaDupla,"Diária duplicada");
check(dd.length === 1 && dd[0].severidade === "critico",
      "duas diárias da mesma pessoa no mesmo dia é crítico");
check(dd.length === 1 && /reclassifica/i.test(dd[0].raciocinio),
      "…e o raciocínio levanta reclassificação quando os tipos diferem");

const aritmetica = rodar([], [
  D({ groot:"4000002", nome:"RENATO LIMA", data:d(2026,8,3), qtd:1, unit:366.04, valor:500 })
]);
check(acharRegra(aritmetica,"Valor final incompatível").length === 1,
      "VALOR FINAL diferente de quantidade × unitário é apontado");

const invertida = rodar([
  L({ groot:"4000003", nome:"PAULO NUNES", ini:d(2026,8,10), fim:d(2026,8,1), valor:100 })
]);
check(acharRegra(invertida,"Datas invertidas").length === 1,
      "DATA FIM anterior ao início é apontada");

const fora = rodar([], [D({ groot:"4000004", nome:"SUELI PAZ", data:d(2026,5,3) })]);
check(acharRegra(fora,"Data fora do período").length === 1,
      "diária fora do período faturado é apontada");

const homonimo = rodar([
  L({ groot:"4000005", nome:"CARLOS SILVA", ini:d(2026,7,20), valor:100 }),
  L({ groot:"4000006", nome:"CARLOS SILVA", ini:d(2026,7,20), valor:100 })
]);
const hom = acharRegra(homonimo,"Homônimo ou cadastro duplicado");
check(hom.length === 1 && hom[0].severidade === "revisar",
      "mesmo nome com GROOTs diferentes fica em Revisar — homônimo existe");
check(hom.length === 1 && /sozinho não decide/.test(hom[0].raciocinio),
      "…dizendo explicitamente que o nome sozinho não decide");

const linhaDupla = rodar([
  L({ groot:"4000007", nome:"IVO MARTINS", ini:d(2026,7,1), fim:d(2026,8,1), valor:3000 }),
  L({ groot:"4000007", nome:"IVO MARTINS", ini:d(2026,7,1), fim:d(2026,8,1), valor:3000 })
]);
check(acharRegra(linhaDupla,"Linha duplicada").length === 1,
      "linhas idênticas no LABOR são apontadas como repetição");

const regime = rodar([
  L({ groot:"4000008", nome:"A", regime:"Efetivo", ini:d(2026,7,20), valor:100 }),
  L({ groot:"4000009", nome:"B", regime:"efetivo", ini:d(2026,7,20), valor:100 })
]);
check(acharRegra(regime,"Grafia inconsistente").length === 1,
      '"Efetivo" e "efetivo" são apontados como grafia inconsistente');

const heOrfa = rodar(
  [L({ groot:"5000001", nome:"NO LABOR", ini:d(2026,7,20), valor:100 })],
  [],
  [{ aba:"HORA EXTRA", linha:2, groot:"5000002", nome:"FORA DO LABOR", cargo:"Auxiliar", valor:400 }]
);
check(acharRegra(heOrfa,"Hora extra sem vínculo").length === 1,
      "hora extra de quem não está no LABOR é apontada");

const tarifa = rodar([], [
  ...Array.from({length:8},(_,i) => D({ groot:String(6000000+i), nome:"D"+i, data:d(2026,8,3), unit:366.04, valor:366.04 })),
  D({ groot:"6000099", nome:"FORA DA CURVA", data:d(2026,8,3), unit:900, valor:900 })
]);
const tf = acharRegra(tarifa,"Tarifa destoante");
check(tf.length === 1 && tf[0].groot === "6000099",
      "valor unitário muito acima da mediana do mesmo tipo é apontado", "n="+tf.length);

/* ================================================================
   Comportamento geral
   ================================================================ */
console.log("\nComportamento do conjunto");

const limpa = rodar(
  [L({ groot:"7000001", nome:"PESSOA LIMPA", ini:d(2026,7,20), fim:null, valor:5000 })],
  [D({ groot:"7000002", nome:"OUTRA PESSOA", data:d(2026,8,3) })]
);
check(limpa.achados.length === 0,
      "uma fatura sem problema nenhum não gera achado nenhum",
      JSON.stringify(limpa.achados.map(a => a.regra)));

const ordenada = rodar([
  L({ groot:"8000001", nome:"X", regime:"Efetivo", ini:d(2026,1,1), fim:d(2026,3,1), valor:10 }),
  L({ groot:"8000001", nome:"X", regime:"Temporário", ini:d(2026,3,2), fim:null, valor:10 }),
  L({ groot:"8000002", nome:"A ALPHA", ini:d(2026,5,1), fim:null, valor:9999 }),
  L({ groot:"8000002", nome:"B BETA", ini:d(2026,6,1), fim:null, valor:9999 })
]);
check(ordenada.achados[0].severidade === "critico",
      "os críticos vêm primeiro, antes de qualquer informativo");
check(ordenada.achados.every(a => a.raciocinio && a.raciocinio.length > 60),
      "todo achado carrega um raciocínio em texto, não só um rótulo");
check(ordenada.achados.every(a => Array.isArray(a.registros)),
      "todo achado carrega os registros que o sustentam, para o lado a lado");
check(ordenada.achados.every(a => a.id && typeof a.id === "string"),
      "todo achado tem id estável — é o que guarda a decisão humana na tela");

/* Linhas em branco do template não podem virar apontamento: o modelo
   vem com centenas delas. */
const vazias = rodar([
  L({ groot:"", nome:"", cargo:"", ini:null, fim:null, valor:0 }),
  L({ groot:"", nome:"", cargo:"", ini:null, fim:null, valor:null })
]);
check(vazias.achados.length === 0,
      "linhas em branco do template não geram achado");

/* ================================================================
   AGRUPAMENTO POR CATEGORIA

   A tela não lista trinta cartões em coluna: reúne os achados em
   blocos que se abrem. O agrupamento é responsabilidade do motor, para
   a tela não precisar saber que regra pertence a que assunto.
   ================================================================ */
console.log("\nAgrupamento por categoria");

const variado = rodar([
  /* pessoas diferentes, sobrepostas e ambas positivas → crítico */
  L({ groot:"5100001", nome:"RENATO FARIAS",  ini:d(2026,5,1), fim:null, valor:100 }),
  L({ groot:"5100001", nome:"OLIVIA MENDONCA", ini:d(2026,6,1), fim:null, valor:100 }),
  /* mesmo primeiro nome, sobrenomes distintos → parcialmente coincidentes */
  L({ groot:"5100002", nome:"CARLOS PRIMEIRO SOUZA", ini:d(2026,5,1), fim:null, valor:100 }),
  L({ groot:"5100002", nome:"CARLOS SEGUNDO LIMA",   ini:d(2026,6,1), fim:null, valor:100 }),
  /* mesma pessoa, efetivação */
  L({ groot:"5100003", nome:"DORA MELO", regime:"Temporário", ini:d(2026,5,1), fim:d(2026,7,31), valor:100 }),
  L({ groot:"5100003", nome:"DORA MELO", regime:"Efetivo",    ini:d(2026,8,1), fim:null, valor:100 }),
  /* sem identificador */
  L({ groot:"", nome:"ELI SANTOS", cargo:"Auxiliar de Apoio LOG I", ini:d(2026,7,20), valor:100 })
]);

const chaves = variado.grupos.map(g => g.chave);
check(chaves.includes("pessoas-diferentes") && chaves.includes("parcialmente-diferentes")
   && chaves.includes("mesma-pessoa") && chaves.includes("sem-groot"),
      "cada assunto vira o seu grupo", JSON.stringify(chaves));
check(variado.grupos.every(g => g.total > 0),
      "grupo sem achado não é criado — a tela não mostra bloco vazio");
check(variado.grupos.reduce((n,g) => n + g.total, 0) === variado.achados.length,
      "todo achado cai em exatamente um grupo, nenhum se perde",
      variado.grupos.reduce((n,g) => n + g.total, 0) + " de " + variado.achados.length);
check(variado.achados.every(a => a.categoria && a.categoria !== "outros"),
      "nenhuma regra ficou sem categoria declarada",
      JSON.stringify([...new Set(variado.achados.filter(a => a.categoria === "outros").map(a => a.regra))]));

/* A ordem dos grupos é a declarada, não a ordem em que os achados
   apareceram: o assunto mais caro primeiro. */
const ordemDeclarada = VAL_CATEGORIAS.map(c => c.chave);
check(chaves.every((c,i) => i === 0 || ordemDeclarada.indexOf(chaves[i-1]) < ordemDeclarada.indexOf(c)),
      "os grupos saem na ordem declarada, do mais caro ao mais barato",
      JSON.stringify(chaves));

/* Um grupo com um crítico no meio de informativos não pode se anunciar
   como informativo — a cor do bloco é a do pior caso que ele guarda. */
const misto = rodar([
  L({ groot:"5200001", nome:"FABIO NOGUEIRA", regime:"Temporário", ini:d(2026,5,1), fim:d(2026,7,31), valor:100 }),
  L({ groot:"5200001", nome:"FABIO NOGUEIRA", regime:"Efetivo",    ini:d(2026,8,1), fim:null, valor:100 }),
  L({ groot:"5200002", nome:"GILDA PACHECO", ini:d(2026,5,1), fim:null, valor:100 }),
  L({ groot:"5200002", nome:"GILDA PACHECO", ini:d(2026,6,1), fim:null, valor:100 })
]);
const gMesma = misto.grupos.find(g => g.chave === "mesma-pessoa");
check(!!gMesma && gMesma.total === 2 && gMesma.severidade === "critico",
      "a severidade do grupo é a pior que ele contém, não a média nem a primeira",
      gMesma && gMesma.severidade+" ("+JSON.stringify(gMesma.porSev)+")");
check(!!gMesma && gMesma.porSev.critico === 1 && gMesma.porSev.info === 1,
      "…e o grupo informa quantos de cada severidade guarda");

const limpo2 = rodar([L({ groot:"5300001", nome:"SEM PROBLEMA", ini:d(2026,7,20), valor:100 })]);
check(limpo2.grupos.length === 0,
      "fatura sem achado não gera grupo nenhum");

/* ================================================================
   QUADRO DO PERÍODO

   A auditoria recebe só a fatura, e o quadro é o headcount que dá para
   extrair dela sozinha. A conta é a MESMA da Fusão de Linhas e da
   simulação de retorno — se divergir, o app passa a dizer dois números
   diferentes para a mesma pergunta, em duas telas vizinhas.
   ================================================================ */
console.log("\nQuadro do período");

const { valQuadroDiario } = require("../js/validacao.js");
const quadroLabor = [
  L({ groot:"6100001", cargo:"Auxiliar de Apoio LOG I", ini:d(2026,3,1), fim:null, rateio:1 }),
  L({ groot:"6100002", cargo:"Operador Transpaleteira", ini:d(2026,3,1), fim:null, rateio:1 }),
  /* estorno: reduz o que a fatura apresenta de 27/07 a 31/07 */
  L({ groot:"6100003", cargo:"Auxiliar de Apoio LOG I", ini:d(2026,7,27), fim:d(2026,7,31), rateio:-1 }),
  /* liderança e cargo desconhecido não entram */
  L({ groot:"6100004", cargo:"Supervisor de Operação", ini:d(2026,3,1), fim:null, rateio:1 }),
  L({ groot:"6100005", cargo:"Conferente de Expedição", ini:d(2026,3,1), fim:null, rateio:1 }),
  /* a linha de ABS do template não é pessoa */
  L({ groot:"ABS", nome:" ABSENTEISMO", cargo:" ABSENTEISMO", ini:d(2026,3,1), fim:null, rateio:1 })
];
const q = valQuadroDiario(quadroLabor, PERIODO);
const noDiaQ = dt => q.dias.find(x => +x.data === +dt);

check(!!q && q.dias.length === 31, "um ponto por dia do período", q && q.dias.length);
check(noDiaQ(d(2026,7,20)).liquido === 2,
      "fora do estorno, o quadro conta as duas pessoas elegíveis",
      noDiaQ(d(2026,7,20)).liquido);
check(noDiaQ(d(2026,7,29)).liquido === 1,
      "no período do estorno o quadro cai — o rateio entra com o sinal",
      noDiaQ(d(2026,7,29)).liquido);
check(noDiaQ(d(2026,7,29)).bruto === 2,
      "…e o bruto do mesmo dia continua 2: no turno havia duas pessoas",
      noDiaQ(d(2026,7,29)).bruto);
check(q.pessoas === 2,
      "liderança, cargo desconhecido e a linha de ABS ficam fora da contagem de pessoas",
      q.pessoas);
check(q.estornos === 1, "o número de estornos é informado", q.estornos);
check(q.min === 1 && q.max === 2, "mínimo e máximo do período", q.min+".."+q.max);

/* A mesma expressão de js/fusao.js, para provar que os dois números
   não divergem por acaso. */
const eligQ = l => ["auxiliar de apoio log i","operador transpaleteira",
                    "part time - auxiliar log i (3 dias)"]
  .includes(String(l.cargo).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim());
const contaFusaoQ = dia => quadroLabor.reduce((acc,l) => acc +
  ((eligQ(l) && String(l.groot).toLowerCase() !== "abs"
    && l.ini instanceof Date && l.ini <= dia
    && (!(l.fim instanceof Date) || l.fim >= dia)) ? l.rateio : 0), 0);
check(q.dias.every(x => x.liquido === contaFusaoQ(x.data)),
      "o quadro bate com a conta da Fusão de Linhas em todos os 31 dias");

console.log("\n" + pass + " passaram, " + fail + " falharam\n");
process.exit(fail ? 1 : 0);
