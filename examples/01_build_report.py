"""
01_build_report.py
===================

Genere le classeur complet de reporting risque (positions, historique VL,
limites, formules Excel natives, graphiques) et l'ecrit dans
``examples/output/reporting_risque.xlsx``. Copie ensuite ce classeur vers
``docs/exemple_reporting_risque.xlsx`` pour qu'il soit commite et
telechargeable directement depuis GitHub.

Usage :
    python examples/01_build_report.py
"""

from __future__ import annotations

import shutil
from pathlib import Path

from riskreporting import data, metrics, validate, workbook

RACINE = Path(__file__).resolve().parent.parent
SORTIE = RACINE / "examples" / "output" / "reporting_risque.xlsx"
COPIE_DOCS = RACINE / "docs" / "exemple_reporting_risque.xlsx"


def main() -> None:
    print("=" * 70)
    print("Generation du reporting risque - Fonds Actions Europe Diversifie")
    print("=" * 70)

    jeu = data.jeu_de_donnees_complet(n_positions=25, n_jours=260)
    positions = jeu["positions"]
    historique = jeu["historique"]
    limites = jeu["limites"]

    SORTIE.parent.mkdir(parents=True, exist_ok=True)
    workbook.construire_classeur(
        chemin_sortie=str(SORTIE),
        positions=positions,
        historique=historique,
        limites=limites,
        nom_portefeuille=jeu["nom_portefeuille"],
        date_valorisation=jeu["date_valorisation"],
        niveau_confiance_var=0.99,
        taux_sans_risque=0.025,
    )
    print(f"\nClasseur ecrit : {SORTIE}")

    # --- Controle de sante post-generation ---
    resultat = validate.valider_classeur(str(SORTIE))
    n_formules = validate.compter_formules(str(SORTIE))
    print(f"Validation structurelle : {'OK' if resultat.ok else 'ECHEC'}")
    if resultat.erreurs:
        for erreur in resultat.erreurs:
            print(f"  - ERREUR : {erreur}")
    for avert in resultat.avertissements:
        print(f"  - avertissement : {avert}")
    print(f"Nombre de formules Excel natives ecrites : {n_formules}")
    resultat.lever_si_invalide()

    # --- Resume des indicateurs cles (calcules cote Python, a titre indicatif -
    # les valeurs "vivantes" du classeur sont dans la feuille Synthese) ---
    rendements_fonds = historique["rendement_fonds"].dropna()
    rendements_bench = historique["rendement_bench"].dropna()
    rendements_fonds_np = rendements_fonds.to_numpy()
    rendements_bench_np = rendements_bench.to_numpy()

    var99 = metrics.var_historique(rendements_fonds_np, niveau=0.99)
    es975 = metrics.expected_shortfall(rendements_fonds_np, niveau=0.975)
    vol_ann = metrics.volatilite_annualisee(rendements_fonds_np)
    te = metrics.tracking_error(rendements_fonds_np, rendements_bench_np)
    beta = metrics.beta_portefeuille(rendements_fonds_np, rendements_bench_np)
    sharpe = metrics.ratio_sharpe(rendements_fonds_np, taux_sans_risque=0.025)
    info_ratio = metrics.ratio_information(rendements_fonds_np, rendements_bench_np)
    mdd = metrics.max_drawdown(historique["vl_fonds"].to_numpy())

    concentration_max = float(positions["exposition_eur"].abs().max() / positions["exposition_eur"].abs().sum())
    expo_brute = positions["exposition_eur"].abs().sum()
    expo_nette = positions["exposition_eur"].sum()
    secteur_max = positions.groupby("secteur")["exposition_eur"].apply(lambda s: s.abs().sum()).idxmax()

    print("\n" + "-" * 70)
    print("Resume des indicateurs cles (portefeuille synthetique)")
    print("-" * 70)
    print(f"  Nombre de positions              : {len(positions)}")
    print(f"  Nombre de jours d'historique      : {len(historique)}")
    print(f"  VaR historique 99% (1 jour)       : {var99:.4%}")
    print(f"  Expected Shortfall 97,5%          : {es975:.4%}")
    print(f"  Volatilite annualisee             : {vol_ann:.4%}")
    print(f"  Tracking Error annualisee         : {te:.4%}")
    print(f"  Beta vs benchmark                 : {beta:.4f}")
    print(f"  Ratio de Sharpe                   : {sharpe:.4f}")
    print(f"  Ratio d'information               : {info_ratio:.4f}")
    print(f"  Max Drawdown                      : {mdd:.4%}")
    print(f"  Concentration emetteur max        : {concentration_max:.4%}")
    print(f"  Exposition brute                  : {expo_brute:,.0f} EUR")
    print(f"  Exposition nette                  : {expo_nette:,.0f} EUR")
    print(f"  Secteur le plus expose             : {secteur_max}")

    # --- Copie vers docs/ pour telechargement direct depuis GitHub ---
    COPIE_DOCS.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(SORTIE, COPIE_DOCS)
    print(f"\nClasseur copie vers : {COPIE_DOCS}")


if __name__ == "__main__":
    main()
