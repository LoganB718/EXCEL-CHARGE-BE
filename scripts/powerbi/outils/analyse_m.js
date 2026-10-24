// Analyse syntaxique d'expressions M (Power Query) avec @microsoft/powerquery-parser.
// Entrée (stdin) : JSON [{"nom": "...", "texte": "..."}] ; sortie (stdout) : JSON [{"nom", "ok", "erreur"}].
// Usage : node analyse_m.js <dossier node_modules contenant @microsoft/powerquery-parser>
"use strict";
const path = require("path");
const racine = process.argv[2];
const PQP = require(path.join(racine, "@microsoft", "powerquery-parser"));

function lireEntree() {
    return new Promise((ok) => {
        let s = "";
        process.stdin.setEncoding("utf8");
        process.stdin.on("data", (d) => (s += d));
        process.stdin.on("end", () => ok(s));
    });
}

function messageErreur(tache) {
    const e = tache.error || {};
    let m = (e.message || String(e)).toString();
    const inner = e.innerError || {};
    if (inner.message && !m.includes(inner.message)) m += " : " + inner.message;
    let tok = inner.foundToken || inner.token;
    if (tok && tok.token) tok = tok.token;
    if (tok && tok.positionStart) m += " (ligne " + (tok.positionStart.lineNumber + 1) + ", colonne " + (tok.positionStart.lineCodeUnit + 1) + ")";
    else if (inner.graphemePosition) m += " (ligne " + (inner.graphemePosition.lineNumber + 1) + ")";
    return m;
}

(async () => {
    const entrees = JSON.parse(await lireEntree());
    const sorties = [];
    for (const e of entrees) {
        try {
            const t = await PQP.TaskUtils.tryLexParse(PQP.DefaultSettings, e.texte);
            if (PQP.TaskUtils.isParseStageOk(t)) sorties.push({ nom: e.nom, ok: true });
            else sorties.push({ nom: e.nom, ok: false, erreur: messageErreur(t) });
        } catch (x) {
            sorties.push({ nom: e.nom, ok: false, erreur: "exception : " + x });
        }
    }
    process.stdout.write(JSON.stringify(sorties));
})();
