/* Carrega os arquivos de js/ em um contexto compartilhado, do mesmo jeito que o
   navegador faz com várias tags <script>. Assim os testes rodam contra o código
   real da aplicação, sem cópias nem adaptações. */
"use strict";
const fs = require("fs");
const vm = require("vm");
const path = require("path");

const JS_DIR = path.join(__dirname, "..", "js");

/** Declarações `const` de nível superior ficam no escopo léxico do contexto e
 *  não viram propriedade do objeto devolvido — diferente de `function`. Para os
 *  testes alcançarem constantes exportadas (RECON_STATUS e afins), avaliamos os
 *  nomes pedidos dentro do próprio contexto e copiamos o resultado.
 *  @param {string[]} files nomes dos arquivos em js/, na ordem de carga
 *  @param {string[]} [consts] nomes de constantes de nível superior a expor */
function load(files, consts) {
  const ctx = {
    console,
    Date, Math, JSON, Object, Array, Number, String, Map, Set, Boolean, RegExp, Error, isFinite, parseFloat, parseInt,
    // stubs mínimos de navegador — só o que a importação toca
    document: {
      getElementById: () => ({ textContent: "", classList: { add() {}, remove() {}, toggle() {} } }),
      querySelector: () => null,
      querySelectorAll: () => [],
      createElement: () => ({ click() {}, style: {} }),
      addEventListener() {}
    },
    alert: msg => { ctx.lastAlert = msg; },
    lastAlert: null
  };
  vm.createContext(ctx);
  for (const f of files) {
    const src = fs.readFileSync(path.join(JS_DIR, f), "utf8");
    vm.runInContext(src, ctx, { filename: "js/" + f });
  }
  if (consts && consts.length) {
    const expr = "({" + consts.map(n => JSON.stringify(n) + ":" + n).join(",") + "})";
    Object.assign(ctx, vm.runInContext(expr, ctx, { filename: "consts" }));
  }
  return ctx;
}

module.exports = { load };
