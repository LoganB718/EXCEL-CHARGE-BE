"""Échantillon des tables calculées par le miroir (pour validate_bim.py --donnees).

Usage : python3 exporter_donnees.py <classeur.xlsx> <sortie.json> [AAAA-MM-JJ]
"""
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mirror  # noqa: E402


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    jour = dt.date.fromisoformat(sys.argv[3]) if len(sys.argv) > 3 else dt.date.today()
    modele = mirror.construire(sys.argv[1], jour, heure=dt.datetime.combine(jour, dt.time(8, 30)))
    with open(sys.argv[2], "w", encoding="utf-8") as f:
        json.dump(modele["tables"], f, ensure_ascii=False, default=str)
    print("échantillon écrit :", sys.argv[2], {k: len(v) for k, v in modele["tables"].items()})
    return 0


if __name__ == "__main__":
    sys.exit(main())
