"""Tableau de validation croisee Python <-> VBA/Excel pour chaque indicateur, avec les valeurs de reference pour vba/modTests.bas."""

from __future__ import annotations

from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import numpy as np

matplotlib.use("Agg")

from riskreporting import data, metrics

RACINE = Path(__file__).resolve().parent.parent
DOSSIER_IMG = RACINE / "docs" / "img"

NIVEAU_VAR = 0.99
NIVEAU_ES = 0.975
TAUX_SANS_RISQUE = 0.025


def section(titre: str) -> None:
    print("\n" + "=" * 78)
    print(titre)
    print("=" * 78)


def ligne_validation(nom: str, valeur_python: float, formule_excel: str, ref_vba: str) -> None:
    print(f"{nom:<32} Python={valeur_python: .6f}   Excel: {formule_excel}")
    print(f"{'':<32} -> valeur de reference pour modTests.bas : {ref_vba}")


def main() -> None:
    jeu = data.jeu_de_donnees_complet(n_positions=25, n_jours=260)
    historique = jeu["historique"]
    positions = jeu["positions"]

    rendements_fonds = historique["rendement_fonds"].dropna().to_numpy()
    rendements_bench = historique["rendement_bench"].dropna().to_numpy()
    vl_fonds = historique["vl_fonds"].to_numpy()

    section("Validation croisee Python (riskreporting.metrics) <-> VBA/Excel")
    print("Jeu de donnees : historique synthetique 260 jours, graine deterministe.")
    print("Ces valeurs peuvent etre recopiees dans vba/modTests.bas comme")
    print("references numeriques supplementaires (portefeuille complet, en plus")
    print("des cas analytiques deja presents dans le module).\n")

    var99 = metrics.var_historique(rendements_fonds, niveau=NIVEAU_VAR)
    ligne_validation(
        "VaR historique 99%",
        var99,
        "=VaRHistorique(RendementsFonds;0,99)  /  =-PERCENTILE.INC(RendementsFonds;0,01)",
        f"{var99:.6f}",
    )

    var_param = metrics.var_parametrique(rendements_fonds, niveau=NIVEAU_VAR)
    ligne_validation(
        "VaR parametrique 99%",
        var_param,
        "=VaRParametrique(RendementsFonds;0,99)",
        f"{var_param:.6f}",
    )

    var_cf = metrics.var_cornish_fisher(rendements_fonds, niveau=NIVEAU_VAR)
    ligne_validation(
        "VaR Cornish-Fisher 99%",
        var_cf,
        "=VaRCornishFisher(RendementsFonds;0,99)",
        f"{var_cf:.6f}",
    )

    es975 = metrics.expected_shortfall(rendements_fonds, niveau=NIVEAU_ES)
    ligne_validation(
        "Expected Shortfall 97,5%",
        es975,
        "=ExpectedShortfall(RendementsFonds;0,975)",
        f"{es975:.6f}",
    )

    vol_ann = metrics.volatilite_annualisee(rendements_fonds)
    ligne_validation(
        "Volatilite annualisee",
        vol_ann,
        "=VolatiliteAnnualisee(RendementsFonds;252)  /  =STDEV.S(RendementsFonds)*SQRT(252)",
        f"{vol_ann:.6f}",
    )

    vol_ewma = metrics.volatilite_ewma(rendements_fonds, lam=0.94)
    ligne_validation(
        "Volatilite EWMA (lambda=0,94)",
        vol_ewma,
        "=VolatiliteEWMA(RendementsFonds;0,94)",
        f"{vol_ewma:.6f}",
    )

    te = metrics.tracking_error(rendements_fonds, rendements_bench)
    ligne_validation(
        "Tracking Error annualisee",
        te,
        "=TrackingError(RendementsFonds;RendementsBench;252)  /  =STDEV.S(RendementActif)*SQRT(252)",
        f"{te:.6f}",
    )

    beta = metrics.beta_portefeuille(rendements_fonds, rendements_bench)
    ligne_validation(
        "Beta vs benchmark",
        beta,
        "=BetaPortefeuille(RendementsFonds;RendementsBench)  /  =SLOPE(RendementsFonds;RendementsBench)",
        f"{beta:.6f}",
    )

    sharpe = metrics.ratio_sharpe(rendements_fonds, taux_sans_risque=TAUX_SANS_RISQUE)
    ligne_validation(
        "Ratio de Sharpe",
        sharpe,
        "=RatioSharpe(RendementsFonds;0,025;252)",
        f"{sharpe:.6f}",
    )

    info_ratio = metrics.ratio_information(rendements_fonds, rendements_bench)
    ligne_validation(
        "Ratio d'information",
        info_ratio,
        "=RatioInformation(RendementsFonds;RendementsBench;252)",
        f"{info_ratio:.6f}",
    )

    mdd = metrics.max_drawdown(vl_fonds)
    ligne_validation(
        "Max Drawdown",
        mdd,
        "=MaxDrawdown(VLFonds)  /  =MIN(DrawdownFonds)",
        f"{mdd:.6f}",
    )

    section("Black-Scholes / Grecques (S=100, K=100, T=1, r=5%, q=0%, vol=20%)")
    parametres = dict(S=100, K=100, T=1, r=0.05, q=0, sigma=0.2)
    call = metrics.black_scholes("C", **parametres)
    put = metrics.black_scholes("P", **parametres)
    delta_call = metrics.bs_delta("C", **parametres)
    delta_put = metrics.bs_delta("P", **parametres)
    gamma = metrics.bs_gamma(**parametres)
    vega = metrics.bs_vega(**parametres)

    ligne_validation("BlackScholes Call", call, '=BlackScholes("C";100;100;1;0,05;0;0,2)', f"{call:.6f}")
    ligne_validation("BlackScholes Put", put, '=BlackScholes("P";100;100;1;0,05;0;0,2)', f"{put:.6f}")
    ligne_validation("BSDelta Call", delta_call, '=BSDelta("C";100;100;1;0,05;0;0,2)', f"{delta_call:.6f}")
    ligne_validation("BSDelta Put", delta_put, '=BSDelta("P";100;100;1;0,05;0;0,2)', f"{delta_put:.6f}")
    ligne_validation("BSGamma", gamma, "=BSGamma(100;100;1;0,05;0;0,2)", f"{gamma:.6f}")
    ligne_validation("BSVega (convention /100)", vega, "=BSVega(100;100;1;0,05;0;0,2)", f"{vega:.6f}")

    DOSSIER_IMG.mkdir(parents=True, exist_ok=True)

    drawdowns = metrics.drawdown_series(historique["vl_fonds"])

    fig, (ax_vl, ax_dd) = plt.subplots(
        2, 1, figsize=(10, 6), dpi=130, sharex=True, gridspec_kw={"height_ratios": [2, 1]}
    )
    ax_vl.plot(historique.index, historique["vl_fonds"], label="Fonds", color="#1F4E78", linewidth=1.6)
    ax_vl.plot(
        historique.index, historique["vl_bench"], label="Benchmark", color="#9C5700", linewidth=1.2, linestyle="--"
    )
    ax_vl.set_title("Evolution de la valeur liquidative et zones de drawdown")
    ax_vl.set_ylabel("VL (base 100)")
    ax_vl.legend(loc="upper left")
    ax_vl.grid(alpha=0.3)

    ax_dd.fill_between(historique.index, drawdowns.to_numpy() * 100, 0, color="#9C0006", alpha=0.5)
    ax_dd.set_ylabel("Drawdown (%)")
    ax_dd.set_xlabel("Date")
    ax_dd.grid(alpha=0.3)

    fig.tight_layout()
    chemin_vl = DOSSIER_IMG / "vl_drawdown.png"
    fig.savefig(chemin_vl, dpi=130)
    plt.close(fig)
    print(f"\nGraphique VL + drawdown ecrit : {chemin_vl}")

    expo_par_secteur = positions.groupby("secteur")["exposition_eur"].apply(lambda s: s.abs().sum())
    expo_par_secteur = expo_par_secteur.sort_values(ascending=True)
    contribution_pct = expo_par_secteur / expo_par_secteur.sum() * 100

    fig2, ax = plt.subplots(figsize=(10, 6), dpi=130)
    couleurs = matplotlib.colormaps["Blues"](np.linspace(0.4, 0.9, len(contribution_pct)))
    ax.barh(contribution_pct.index, contribution_pct.to_numpy(), color=couleurs)
    ax.set_title("Contribution au risque par secteur (exposition brute)")
    ax.set_xlabel("Part de l'exposition brute (%)")
    ax.grid(axis="x", alpha=0.3)
    fig2.tight_layout()
    chemin_secteurs = DOSSIER_IMG / "contribution_secteurs.png"
    fig2.savefig(chemin_secteurs, dpi=130)
    plt.close(fig2)
    print(f"Graphique contribution au risque ecrit : {chemin_secteurs}")


if __name__ == "__main__":
    main()
