"""Indicateurs de risque de marche, en miroir exact des UDF VBA de vba/modRiskMetrics.bas et vba/modOptions.bas."""

from __future__ import annotations

import numpy as np
import pandas as pd
from scipy import stats


def _as_array(rendements: pd.Series | np.ndarray) -> np.ndarray:
    arr = np.asarray(rendements, dtype=float)
    arr = arr[~np.isnan(arr)]
    return arr


def var_historique(rendements: pd.Series | np.ndarray, niveau: float = 0.99) -> float:
    """VaR historique (non parametrique), miroir de modRiskMetrics.VaRHistorique. Utilise le percentile lineaire "inclusif" (equivalent PERCENTILE.INC / numpy.percentile methode par defaut "linear")."""
    if not (0 < niveau < 1):
        raise ValueError("niveau doit etre strictement compris entre 0 et 1")

    arr = _as_array(rendements)
    if arr.size == 0:
        raise ValueError("serie de rendements vide")

    quantile = np.percentile(arr, (1 - niveau) * 100, method="linear")
    return max(0.0, -float(quantile))


def var_parametrique(
    rendements: pd.Series | np.ndarray, niveau: float = 0.99, horizon: float = 1
) -> float:
    """VaR parametrique gaussienne, miroir de modRiskMetrics.VaRParametrique."""
    if not (0 < niveau < 1):
        raise ValueError("niveau doit etre strictement compris entre 0 et 1")
    if horizon < 1:
        raise ValueError("horizon doit etre >= 1")

    arr = _as_array(rendements)
    mu = arr.mean()
    sigma = arr.std(ddof=1)
    z = stats.norm.ppf(niveau)

    return max(0.0, -(mu * horizon - z * sigma * np.sqrt(horizon)))


def var_cornish_fisher(rendements: pd.Series | np.ndarray, niveau: float = 0.99) -> float:
    """VaR Cornish-Fisher, miroir de modRiskMetrics.VaRCornishFisher."""
    if not (0 < niveau < 1):
        raise ValueError("niveau doit etre strictement compris entre 0 et 1")

    arr = _as_array(rendements)
    if arr.size < 4:
        raise ValueError("au moins 4 observations sont necessaires")

    mu = arr.mean()
    sigma = arr.std(ddof=1)
    if sigma == 0:
        return 0.0

    z_scores = (arr - mu) / sigma
    skew = np.mean(z_scores**3)
    kurt = np.mean(z_scores**4) - 3

    z = stats.norm.ppf(niveau)
    z_cf = (
        z
        + (z**2 - 1) * skew / 6
        + (z**3 - 3 * z) * kurt / 24
        - (2 * z**3 - 5 * z) * (skew**2) / 36
    )

    return max(0.0, -(mu - z_cf * sigma))


def expected_shortfall(rendements: pd.Series | np.ndarray, niveau: float = 0.975) -> float:
    """Expected Shortfall historique, miroir de modRiskMetrics.ExpectedShortfall."""
    if not (0 < niveau < 1):
        raise ValueError("niveau doit etre strictement compris entre 0 et 1")

    arr = _as_array(rendements)
    n = arr.size
    if n < 1:
        raise ValueError("serie de rendements vide")

    triee = np.sort(arr)
    n_queue = int(np.ceil((1 - niveau) * n))
    n_queue = max(1, min(n_queue, n))

    moyenne_queue = triee[:n_queue].mean()
    return max(0.0, -float(moyenne_queue))


def volatilite_annualisee(rendements: pd.Series | np.ndarray, frequence: float = 252) -> float:
    """Volatilite annualisee, miroir de modRiskMetrics.VolatiliteAnnualisee."""
    if frequence <= 0:
        raise ValueError("frequence doit etre positive")
    arr = _as_array(rendements)
    if arr.size < 2:
        raise ValueError("au moins 2 observations sont necessaires")
    return float(arr.std(ddof=1) * np.sqrt(frequence))


def volatilite_ewma(rendements: pd.Series | np.ndarray, lam: float = 0.94) -> float:
    """Volatilite EWMA (RiskMetrics), miroir de modRiskMetrics.VolatiliteEWMA. La derniere valeur de la serie est traitee comme l'observation la plus recente (poids maximal)."""
    if not (0 < lam < 1):
        raise ValueError("lambda doit etre strictement compris entre 0 et 1")

    arr = _as_array(rendements)
    n = arr.size
    if n < 2:
        raise ValueError("au moins 2 observations sont necessaires")

    i = np.arange(n)
    poids = (1 - lam) * (lam**i)
    rendements_recents_en_premier = arr[::-1]

    var_ewma = np.sum(poids * rendements_recents_en_premier**2) / np.sum(poids)
    return float(np.sqrt(var_ewma))


def tracking_error(
    rendements_fonds: pd.Series | np.ndarray,
    rendements_bench: pd.Series | np.ndarray,
    frequence: float = 252,
) -> float:
    """Tracking error annualisee, miroir de modRiskMetrics.TrackingError."""
    if frequence <= 0:
        raise ValueError("frequence doit etre positive")

    f = _as_array(rendements_fonds)
    b = _as_array(rendements_bench)
    if f.size != b.size or f.size < 2:
        raise ValueError("les deux series doivent avoir la meme taille (>= 2)")

    actifs = f - b
    return float(actifs.std(ddof=1) * np.sqrt(frequence))


def beta_portefeuille(
    rendements_fonds: pd.Series | np.ndarray, rendements_bench: pd.Series | np.ndarray
) -> float:
    """Beta du portefeuille vs benchmark, miroir de modRiskMetrics.BetaPortefeuille."""
    f = _as_array(rendements_fonds)
    b = _as_array(rendements_bench)
    if f.size != b.size or f.size < 2:
        raise ValueError("les deux series doivent avoir la meme taille (>= 2)")

    cov = np.cov(f, b, ddof=1)[0, 1]
    var_b = np.var(b, ddof=1)
    if var_b == 0:
        raise ZeroDivisionError("variance du benchmark nulle")

    return float(cov / var_b)


def ratio_sharpe(
    rendements: pd.Series | np.ndarray, taux_sans_risque: float = 0.0, frequence: float = 252
) -> float:
    """Ratio de Sharpe annualise, miroir de modRiskMetrics.RatioSharpe."""
    if frequence <= 0:
        raise ValueError("frequence doit etre positive")
    arr = _as_array(rendements)
    if arr.size < 2:
        raise ValueError("au moins 2 observations sont necessaires")

    rendement_annualise = arr.mean() * frequence
    vol_annualisee = arr.std(ddof=1) * np.sqrt(frequence)
    if vol_annualisee == 0:
        raise ZeroDivisionError("volatilite nulle")

    return float((rendement_annualise - taux_sans_risque) / vol_annualisee)


def ratio_information(
    rendements_fonds: pd.Series | np.ndarray,
    rendements_bench: pd.Series | np.ndarray,
    frequence: float = 252,
) -> float:
    """Ratio d'information, miroir de modRiskMetrics.RatioInformation."""
    if frequence <= 0:
        raise ValueError("frequence doit etre positive")

    f = _as_array(rendements_fonds)
    b = _as_array(rendements_bench)
    if f.size != b.size or f.size < 2:
        raise ValueError("les deux series doivent avoir la meme taille (>= 2)")

    actifs = f - b
    rendement_actif_annualise = actifs.mean() * frequence
    te_annualisee = actifs.std(ddof=1) * np.sqrt(frequence)
    if te_annualisee == 0:
        raise ZeroDivisionError("tracking error nulle")

    return float(rendement_actif_annualise / te_annualisee)


def max_drawdown(niveaux: pd.Series | np.ndarray) -> float:
    """Max Drawdown (valeur negative ou nulle), miroir de modRiskMetrics.MaxDrawdown. ``niveaux`` est une serie de VALEURS (VL), pas de rendements."""
    arr = _as_array(niveaux)
    if arr.size < 2:
        raise ValueError("au moins 2 observations sont necessaires")

    plus_haut_cumule = np.maximum.accumulate(arr)
    drawdowns = np.where(
        plus_haut_cumule > 0, (arr - plus_haut_cumule) / plus_haut_cumule, 0.0
    )
    pire = float(drawdowns.min())
    return min(0.0, pire)


def drawdown_series(niveaux: pd.Series) -> pd.Series:
    """Serie complete de drawdown (pas seulement le maximum), utile pour le graphique "VL + drawdowns". Renvoie une Series alignee sur l'index d'entree."""
    arr = _as_array(niveaux)
    plus_haut_cumule = np.maximum.accumulate(arr)
    drawdowns = np.where(
        plus_haut_cumule > 0, (arr - plus_haut_cumule) / plus_haut_cumule, 0.0
    )
    if isinstance(niveaux, pd.Series):
        return pd.Series(drawdowns, index=niveaux.index, name="drawdown")
    return pd.Series(drawdowns, name="drawdown")


def contribution_au_risque(
    expositions: pd.Series | np.ndarray, volatilites: pd.Series | np.ndarray
) -> np.ndarray:
    """Contribution au risque simplifiee par ligne, sous hypothese de correlation unitaire entre lignes (approximation prudente et lisible pour un graphique de reporting) : contribution_i = |exposition_i| * volatilite_i, normalisee pour sommer a 1."""
    expo = np.asarray(expositions, dtype=float)
    vol = np.asarray(volatilites, dtype=float)
    contrib_brute = np.abs(expo) * vol
    total = contrib_brute.sum()
    if total == 0:
        return np.zeros_like(contrib_brute)
    return contrib_brute / total


def black_scholes(
    type_option: str, S: float, K: float, T: float, r: float, q: float, sigma: float
) -> float:
    """Prix Black-Scholes-Merton, miroir de modOptions.BlackScholes."""
    if S <= 0 or K <= 0 or T < 0 or sigma <= 0:
        raise ValueError("parametres hors domaine (S, K, sigma > 0 et T >= 0 requis)")

    type_option = type_option.upper()
    if type_option not in ("C", "CALL", "P", "PUT"):
        raise ValueError(f"type d'option invalide : {type_option}")
    est_call = type_option in ("C", "CALL")

    if T == 0:
        return max(0.0, S - K) if est_call else max(0.0, K - S)

    d1 = (np.log(S / K) + (r - q + 0.5 * sigma**2) * T) / (sigma * np.sqrt(T))
    d2 = d1 - sigma * np.sqrt(T)

    if est_call:
        return float(S * np.exp(-q * T) * stats.norm.cdf(d1) - K * np.exp(-r * T) * stats.norm.cdf(d2))
    return float(K * np.exp(-r * T) * stats.norm.cdf(-d2) - S * np.exp(-q * T) * stats.norm.cdf(-d1))


def bs_delta(type_option: str, S: float, K: float, T: float, r: float, q: float, sigma: float) -> float:
    """Delta Black-Scholes, miroir de modOptions.BSDelta."""
    type_option = type_option.upper()
    est_call = type_option in ("C", "CALL")
    d1 = (np.log(S / K) + (r - q + 0.5 * sigma**2) * T) / (sigma * np.sqrt(T))
    if est_call:
        return float(np.exp(-q * T) * stats.norm.cdf(d1))
    return float(np.exp(-q * T) * (stats.norm.cdf(d1) - 1))


def bs_gamma(S: float, K: float, T: float, r: float, q: float, sigma: float) -> float:
    """Gamma Black-Scholes, miroir de modOptions.BSGamma."""
    d1 = (np.log(S / K) + (r - q + 0.5 * sigma**2) * T) / (sigma * np.sqrt(T))
    return float(np.exp(-q * T) * stats.norm.pdf(d1) / (S * sigma * np.sqrt(T)))


def bs_vega(S: float, K: float, T: float, r: float, q: float, sigma: float) -> float:
    """Vega Black-Scholes (convention desk : pour 1 point de vol), miroir de modOptions.BSVega."""
    d1 = (np.log(S / K) + (r - q + 0.5 * sigma**2) * T) / (sigma * np.sqrt(T))
    return float(S * np.exp(-q * T) * stats.norm.pdf(d1) * np.sqrt(T) / 100)
