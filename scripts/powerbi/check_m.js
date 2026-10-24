// Analyse syntaxique des expressions Power Query M avec l'analyseur officiel @microsoft/powerquery-parser.
// Usage : node check_m.js <expressions.json> [noms_globaux.json]
//   expressions.json : objet JSON {nom: expression M}
//   noms_globaux.json (facultatif) : liste des noms de requêtes ; chaque identifiant de chaque expression est alors
//   résolu dans sa portée (let, champs d'enregistrement, paramètres de fonction, each) ou parmi ces noms ;
//   les identifiants pointés (Table.SelectRows...) sont listés « BIBLIOTHÈQUE nom » pour contrôle par l'appelant.
// Sortie : une ligne par expression (OK / ERREUR avec position), code de sortie 1 si au moins une erreur.
"use strict";
const fs = require("fs");
const path = require("path");

// node_modules : variable PBI_NODE_MODULES, sinon node_modules/ à côté du script, sinon ../tools/mparser/node_modules
const racine = process.env.PBI_NODE_MODULES || [path.join(__dirname, "node_modules"),
    path.join(__dirname, "..", "tools", "mparser", "node_modules")].find((d) => fs.existsSync(d))
    || path.join(__dirname, "node_modules");
const PQP = require(path.join(racine, "@microsoft", "powerquery-parser"));

function position(texte, decalage) {
    // ligne et colonne (base 1) d'un décalage UTF-16
    const avant = texte.slice(0, decalage).split("\n");
    return `ligne ${avant.length}, colonne ${avant[avant.length - 1].length + 1}`;
}

function decrire(erreur, texte) {
    const inner = erreur && (erreur.innerError || erreur);
    const msg = (inner && inner.message) || String(erreur);
    // jeton fautif (erreurs de l'analyseur) ou position du lexeur
    let pos = "";
    const jeton = inner && (inner.maybeFoundToken || inner.foundToken || inner.token);
    const tok = jeton && (jeton.token || jeton);
    if (tok && tok.positionStart) {
        const p = tok.positionStart;
        pos = p.codeUnit !== undefined ? position(texte, p.codeUnit) : `ligne ${p.lineNumber + 1}`;
    } else if (inner && inner.graphemePosition) {
        pos = `ligne ${inner.graphemePosition.lineNumber + 1}, colonne ${inner.graphemePosition.lineCodeUnit + 1}`;
    } else if (inner && inner.position) {
        pos = JSON.stringify(inner.position);
    }
    return `${msg}${pos ? " (" + pos + ")" : ""}`;
}

// ---- résolution des identifiants dans leur portée
// constructeurs intrinsèques du langage (#table(...), #date(...)...)
const INTRINSEQUES = new Set(["#table", "#date", "#datetime", "#datetimezone", "#time", "#duration", "#binary",
    "#shared", "#sections", "#infinity", "#nan"]);

function nomIdentifiant(literal) {
    // #"Nom avec espaces" -> Nom avec espaces
    return literal.startsWith('#"') ? literal.slice(2, -1).replace(/""/g, '"') : literal;
}

function csv(tableau) {
    return tableau && tableau.elements ? tableau.elements.map((e) => e.node) : [];
}

function parcourir(noeud, portee, references) {
    if (!noeud || typeof noeud !== "object") return;
    if (Array.isArray(noeud)) {
        noeud.forEach((n) => parcourir(n, portee, references));
        return;
    }
    switch (noeud.kind) {
        case "LetExpression": {
            // les variables d'un let sont visibles entre elles et dans l'expression après in
            const liaisons = csv(noeud.variableList);
            const p = new Set(portee);
            liaisons.forEach((l) => p.add(nomIdentifiant(l.key.literal)));
            liaisons.forEach((l) => parcourir(l.value, p, references));
            parcourir(noeud.expression, p, references);
            return;
        }
        case "RecordExpression": {
            // les champs d'un enregistrement littéral sont visibles entre eux
            const champs = csv(noeud.content);
            const p = new Set(portee);
            champs.forEach((c) => p.add(nomIdentifiant(c.key.literal)));
            champs.forEach((c) => parcourir(c.value, p, references));
            return;
        }
        case "FunctionExpression": {
            const params = csv(noeud.parameters.content);
            const p = new Set(portee);
            params.forEach((x) => p.add(nomIdentifiant(x.name.literal)));
            params.forEach((x) => parcourir(x.parameterType, portee, references));
            parcourir(noeud.functionReturnType, portee, references);
            parcourir(noeud.expression, p, references);
            return;
        }
        case "EachExpression": {
            const p = new Set(portee);
            p.add("_");
            parcourir(noeud.paired, p, references);
            return;
        }
        case "IdentifierExpression":
            references.push({ nom: nomIdentifiant(noeud.identifier.literal), portee });
            return;
        default:
            for (const [cle, valeur] of Object.entries(noeud)) {
                if (cle !== "tokenRange" && valeur && typeof valeur === "object") parcourir(valeur, portee, references);
            }
    }
}

async function main() {
    const fichier = process.argv[2];
    const expressions = JSON.parse(fs.readFileSync(fichier, "utf8"));
    const globaux = process.argv[3] ? new Set(JSON.parse(fs.readFileSync(process.argv[3], "utf8"))) : null;
    const bibliotheque = new Set();
    let erreurs = 0;
    for (const [nom, texte] of Object.entries(expressions)) {
        // lexeur seul puis lexeur + analyseur
        const lex = PQP.TaskUtils.tryLex(PQP.DefaultSettings, texte);
        if (!PQP.TaskUtils.isOk(lex)) {
            erreurs += 1;
            console.log(`ERREUR lexicale  ${nom} : ${decrire(lex.error, texte)}`);
            continue;
        }
        const tache = await PQP.TaskUtils.tryLexParse(PQP.DefaultSettings, texte);
        if (!PQP.TaskUtils.isOk(tache)) {
            erreurs += 1;
            console.log(`ERREUR syntaxe   ${nom} : ${decrire(tache.error, texte)}`);
            continue;
        }
        let inconnus = [];
        if (globaux) {
            const references = [];
            parcourir(tache.ast, new Set(), references);
            for (const r of references) {
                if (r.portee.has(r.nom) || globaux.has(r.nom) || INTRINSEQUES.has(r.nom)) continue;
                if (r.nom.includes(".")) bibliotheque.add(r.nom);
                else inconnus.push(r.nom);
            }
            inconnus = [...new Set(inconnus)];
        }
        if (inconnus.length) {
            erreurs += 1;
            console.log(`ERREUR identifiant inconnu  ${nom} : ${inconnus.join(", ")}`);
        } else {
            console.log(`OK  ${nom}  (${tache.lexerSnapshot.tokens.length} jetons, racine ${tache.ast.kind})`);
        }
    }
    for (const b of [...bibliotheque].sort()) console.log(`BIBLIOTHÈQUE ${b}`);
    console.log(`${Object.keys(expressions).length} expression(s) M analysée(s), ${erreurs} erreur(s)`);
    process.exit(erreurs ? 1 : 0);
}

main().catch((e) => {
    console.error(e);
    process.exit(2);
});
