"""
test_metrics.py
================

Coherence des indicateurs de risque Python (riskreporting.metrics) contre des
valeurs analytiques calculees a la main ou issues de scipy.stats, avec les
memes jeux de donnees que ceux utilises dans vba/modTests.bas (pour permettre
une comparaison croisee Python <-> VBA a l'oeil).
"""

from __future__ import annotations

import numpy as np
import pytest

from riskreporting import metrics

# Serie utilisee dans vba/modTests.TesterRiskMetricsSurFeuille :
# -0,05, -0,04, ..., 0,04 (pas de 0,01), 10 valeurs, moyenne = -0,005.
RENDEMENTS_10 = np.array([-0.05 + i * 0.01 for i in range(10)])


# ---------------------------------------------------------------------------
# VaR historique
# ---------------------------------------------------------------------------


def test_var_historique_quantile_lineaire_connu():
    # rang = (1-0.9)*(10-1) = 0.9 -> index 0 + 0.9*(index1-index0)
    # quantile = -0.05 + 0.9*0.01 = -0.041 -> VaR = 0.041
    var = metrics.var_historique(RENDEMENTS_10, niveau=0.9)
    assert var == pytest.approx(0.041, abs=1e-9)


def test_var_historique_positive_sur_serie_a_pertes():
    assert metrics.var_historique(RENDEMENTS_10, niveau=0.9) > 0


def test_var_historique_niveau_invalide_leve():
    with pytest.raises(ValueError):
        metrics.var_historique(RENDEMENTS_10, niveau=1.5)
    with pytest.raises(ValueError):
        metrics.var_historique(RENDEMENTS_10, niveau=0)


# ---------------------------------------------------------------------------
# VaR parametrique / Cornish-Fisher
# ---------------------------------------------------------------------------


def test_var_parametrique_valeur_manuelle():
    from scipy import stats

    mu = RENDEMENTS_10.mean()
    sigma = RENDEMENTS_10.std(ddof=1)
    z = stats.norm.ppf(0.95)
    attendu = max(0.0, -(mu - z * sigma))
    assert metrics.var_parametrique(RENDEMENTS_10, niveau=0.95) == pytest.approx(attendu, abs=1e-12)


def test_var_cornish_fisher_egale_parametrique_si_skew_kurt_nuls():
    # Serie symetrique autour de 0 avec skewness et kurtosis-exces proches de 0
    # (approximation d'une gaussienne discretisee) : Cornish-Fisher ~ parametrique.
    rng = np.random.default_rng(42)
    x = rng.standard_normal(20000)
    x = (x - x.mean()) / x.std(ddof=1)  # standardise exactement
    var_param = metrics.var_parametrique(x, niveau=0.99)
    var_cf = metrics.var_cornish_fisher(x, niveau=0.99)
    assert var_cf == pytest.approx(var_param, abs=0.05)


def test_var_cornish_fisher_necessite_au_moins_4_observations():
    with pytest.raises(ValueError):
        metrics.var_cornish_fisher(np.array([0.01, -0.01, 0.02]), niveau=0.99)


# ---------------------------------------------------------------------------
# Expected Shortfall
# ---------------------------------------------------------------------------


def test_expected_shortfall_moyenne_de_la_queue_connue():
    # 10 observations, niveau 0.8 -> n_queue = ceil(0.2*10) = 2
    # queue triee = [-0.05, -0.04] -> moyenne = -0.045 -> ES = 0.045
    es = metrics.expected_shortfall(RENDEMENTS_10, niveau=0.8)
    assert es == pytest.approx(0.045, abs=1e-9)


def test_expected_shortfall_superieure_ou_egale_a_var_historique():
    # Propriete theorique : ES(niveau) >= VaR(niveau) (la moyenne de la queue
    # est au moins aussi mauvaise que son bord).
    var = metrics.var_historique(RENDEMENTS_10, niveau=0.9)
    es = metrics.expected_shortfall(RENDEMENTS_10, niveau=0.9)
    assert es >= var - 1e-12


# ---------------------------------------------------------------------------
# Volatilite
# ---------------------------------------------------------------------------


def test_volatilite_annualisee_serie_constante_est_nulle():
    serie_constante = np.full(10, 0.001)
    assert metrics.volatilite_annualisee(serie_constante, frequence=252) == pytest.approx(0.0, abs=1e-12)


def test_volatilite_annualisee_formule_manuelle():
    attendu = RENDEMENTS_10.std(ddof=1) * np.sqrt(252)
    assert metrics.volatilite_annualisee(RENDEMENTS_10, frequence=252) == pytest.approx(attendu, abs=1e-12)


def test_volatilite_ewma_poids_somment_a_un():
    # Verification indirecte : sur une serie constante, la vol EWMA doit valoir
    # exactement |valeur constante| (poids normalises).
    serie = np.full(15, 0.02)
    assert metrics.volatilite_ewma(serie, lam=0.94) == pytest.approx(0.02, abs=1e-9)


# ---------------------------------------------------------------------------
# Tracking error / beta / ratios (cas fonds = benchmark, cf. modTests.bas)
# ---------------------------------------------------------------------------


def test_tracking_error_nulle_si_fonds_replique_benchmark():
    fonds = np.array([-0.02 + i * 0.005 for i in range(10)])
    bench = fonds.copy()
    assert metrics.tracking_error(fonds, bench, frequence=252) == pytest.approx(0.0, abs=1e-9)


def test_beta_unitaire_si_fonds_egale_benchmark():
    fonds = np.array([-0.02 + i * 0.005 for i in range(10)])
    bench = fonds.copy()
    assert metrics.beta_portefeuille(fonds, bench) == pytest.approx(1.0, abs=1e-9)


def test_beta_leve_si_variance_benchmark_nulle():
    fonds = np.array([0.01, -0.02, 0.03, 0.01, -0.01])
    bench = np.full(5, 0.005)
    with pytest.raises(ZeroDivisionError):
        metrics.beta_portefeuille(fonds, bench)


def test_ratio_information_leve_si_te_nulle():
    fonds = np.array([-0.02 + i * 0.005 for i in range(10)])
    bench = fonds.copy()
    with pytest.raises(ZeroDivisionError):
        metrics.ratio_information(fonds, bench)


def test_ratio_sharpe_leve_si_volatilite_nulle():
    # 0.25 est exactement representable en binaire : l'ecart-type est
    # rigoureusement nul (pas de bruit d'arrondi comme avec 0.001).
    serie_constante = np.full(10, 0.25)
    with pytest.raises(ZeroDivisionError):
        metrics.ratio_sharpe(serie_constante)


# ---------------------------------------------------------------------------
# Max Drawdown (meme jeu de donnees que vba/modTests.bas)
# ---------------------------------------------------------------------------


def test_max_drawdown_valeur_connue():
    vl = np.array([100, 110, 121, 90.75, 100])
    assert metrics.max_drawdown(vl) == pytest.approx(-0.25, abs=1e-9)


def test_max_drawdown_est_toujours_negatif_ou_nul():
    vl_croissante = np.array([100, 101, 105, 110, 120])
    dd = metrics.max_drawdown(vl_croissante)
    assert dd <= 0.0
    assert dd == pytest.approx(0.0, abs=1e-12)


def test_max_drawdown_serie_aleatoire_est_negatif_ou_nul():
    rng = np.random.default_rng(7)
    vl = 100 * np.cumprod(1 + rng.normal(0.0001, 0.01, size=500))
    assert metrics.max_drawdown(vl) <= 0.0


# ---------------------------------------------------------------------------
# Contribution au risque
# ---------------------------------------------------------------------------


def test_contribution_au_risque_somme_a_un():
    expositions = np.array([100, -50, 200, 30])
    volatilites = np.array([0.1, 0.2, 0.15, 0.3])
    contrib = metrics.contribution_au_risque(expositions, volatilites)
    assert contrib.sum() == pytest.approx(1.0, abs=1e-9)
    assert (contrib >= 0).all()


# ---------------------------------------------------------------------------
# Black-Scholes / Grecques : valeurs de reference identiques a
# vba/modTests.TesterBlackScholes et TesterGrecques (S=100,K=100,T=1,r=5%,q=0,vol=20%)
# ---------------------------------------------------------------------------


def test_black_scholes_call_reference():
    prix = metrics.black_scholes("C", 100, 100, 1, 0.05, 0, 0.2)
    assert prix == pytest.approx(10.450584, abs=1e-4)


def test_black_scholes_put_reference():
    prix = metrics.black_scholes("P", 100, 100, 1, 0.05, 0, 0.2)
    assert prix == pytest.approx(5.573526, abs=1e-4)


def test_black_scholes_parite_call_put():
    call = metrics.black_scholes("C", 100, 100, 1, 0.05, 0, 0.2)
    put = metrics.black_scholes("P", 100, 100, 1, 0.05, 0, 0.2)
    # C - P = S*exp(-qT) - K*exp(-rT)
    assert (call - put) == pytest.approx(100 * np.exp(0) - 100 * np.exp(-0.05 * 1), abs=1e-6)


def test_black_scholes_a_echeance_est_valeur_intrinseque():
    assert metrics.black_scholes("C", 110, 100, 0, 0.05, 0, 0.2) == pytest.approx(10.0, abs=1e-9)
    assert metrics.black_scholes("C", 90, 100, 0, 0.05, 0, 0.2) == pytest.approx(0.0, abs=1e-9)


def test_black_scholes_domaine_invalide_leve():
    with pytest.raises(ValueError):
        metrics.black_scholes("C", -100, 100, 1, 0.05, 0, 0.2)
    with pytest.raises(ValueError):
        metrics.black_scholes("C", 100, 100, 1, 0.05, 0, -0.2)


def test_bs_delta_call_put_reference():
    assert metrics.bs_delta("C", 100, 100, 1, 0.05, 0, 0.2) == pytest.approx(0.636831, abs=1e-4)
    assert metrics.bs_delta("P", 100, 100, 1, 0.05, 0, 0.2) == pytest.approx(-0.363169, abs=1e-4)


def test_bs_gamma_reference():
    assert metrics.bs_gamma(100, 100, 1, 0.05, 0, 0.2) == pytest.approx(0.018762, abs=1e-4)


def test_bs_vega_reference():
    assert metrics.bs_vega(100, 100, 1, 0.05, 0, 0.2) == pytest.approx(0.375240, abs=1e-4)
