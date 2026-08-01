"""
bloomberg_stub.py
==================

Simulateur local (100% offline) de l'API Bloomberg Excel / ``blpapi``, alimente
par des donnees synthetiques deterministes. Objectif : permettre de developper
et de tester la logique du reporting (feuille Positions, controles de limites,
formules de synthese) SANS terminal Bloomberg ni connexion reseau, tout en
documentant precisement la structure des champs Bloomberg reellement utilises
par le classeur VBA (``vba/modBloomberg.bas``).

Champs Bloomberg documentes (mnemoniques reels) :
    PX_LAST           : dernier cours de cloture ou cours courant
    VOLATILITY_90D     : volatilite historique 90 jours (%, annualisee)
    CRNCY              : devise de cotation de l'instrument (ISO 4217)
    DUR_ADJ_MID        : duration modifiee (obligations), mid-market
    CUR_MKT_CAP        : capitalisation boursiere courante
    NAME               : nom court de l'emetteur / instrument
    GICS_SECTOR_NAME   : secteur GICS de l'emetteur
    RSK_BB_ISSUER_RATING : notation credit composite Bloomberg

Ce module n'appelle JAMAIS de service reseau : ``bdp``/``bdh``/``bds`` lisent
uniquement un jeu de donnees en memoire, genere par ``data.py`` ou fourni par
l'appelant. La signature des fonctions imite volontairement celle des
wrappers ``xbbg``/``blpapi`` les plus repandus pour que le code de reporting
soit trivial a re-brancher sur une vraie session Bloomberg (remplacer
l'import de ce module par ``import xbbg`` par exemple).
"""

from __future__ import annotations

import numpy as np
import pandas as pd

CHAMPS_DOCUMENTES = {
    "PX_LAST": "Dernier cours de cloture / cours courant",
    "VOLATILITY_90D": "Volatilite historique 90 jours (%, annualisee)",
    "CRNCY": "Devise de cotation ISO 4217",
    "DUR_ADJ_MID": "Duration modifiee (obligations), mid-market",
    "CUR_MKT_CAP": "Capitalisation boursiere courante",
    "NAME": "Nom court de l'emetteur / instrument",
    "GICS_SECTOR_NAME": "Secteur GICS de l'emetteur",
    "RSK_BB_ISSUER_RATING": "Notation credit composite Bloomberg",
}


class BloombergIndisponibleError(RuntimeError):
    """Levee si un champ ou ticker demande n'existe pas dans le jeu local."""


def _generer_reference_marche(tickers: list[str], graine: int = 20260101) -> pd.DataFrame:
    """
    Construit un petit referentiel de marche synthetique pour une liste de
    tickers, avec des valeurs plausibles pour chaque champ documente.
    Deterministe (meme graine -> memes valeurs) pour la reproductibilite des
    tests.
    """
    rng = np.random.default_rng(abs(hash(tuple(tickers))) % (2**32) if tickers else graine)
    n = len(tickers)

    devises = rng.choice(["EUR", "USD", "GBP", "CHF"], size=n, p=[0.55, 0.30, 0.10, 0.05])
    secteurs = rng.choice(
        [
            "Technologie",
            "Sante",
            "Energie",
            "Finance",
            "Industrie",
            "Consommation",
            "Telecommunications",
            "Materiaux",
        ],
        size=n,
    )
    notations = rng.choice(["AAA", "AA", "A", "BBB", "BB", "B"], size=n)

    return pd.DataFrame(
        {
            "ticker": tickers,
            "PX_LAST": np.round(rng.uniform(20, 400, size=n), 2),
            "VOLATILITY_90D": np.round(rng.uniform(10, 45, size=n), 1),
            "CRNCY": devises,
            "DUR_ADJ_MID": np.round(rng.uniform(0.5, 9.0, size=n), 2),
            "CUR_MKT_CAP": np.round(rng.uniform(5e8, 5e11, size=n), 0),
            "NAME": [f"Emetteur simule {i + 1}" for i in range(n)],
            "GICS_SECTOR_NAME": secteurs,
            "RSK_BB_ISSUER_RATING": notations,
        }
    ).set_index("ticker")


def bdp(tickers: str | list[str], champs: str | list[str], graine: int = 20260101) -> pd.DataFrame:
    """
    Simule ``BDP`` (Bloomberg Data Point) : donnee statique "point in time"
    pour une liste de tickers et de champs.

    Parameters
    ----------
    tickers : identifiant(s) Bloomberg, ex "AAPL US Equity" ou une liste.
    champs  : mnemonique(s) de champ Bloomberg (doivent figurer dans
              ``CHAMPS_DOCUMENTES``), ex "PX_LAST" ou une liste.
    graine  : graine du generateur, pour reproductibilite.

    Returns
    -------
    DataFrame indexe par ticker, une colonne par champ demande.

    Raises
    ------
    BloombergIndisponibleError si un champ demande n'est pas documente/simule.
    """
    tickers_liste = [tickers] if isinstance(tickers, str) else list(tickers)
    champs_liste = [champs] if isinstance(champs, str) else list(champs)

    inconnus = [c for c in champs_liste if c not in CHAMPS_DOCUMENTES]
    if inconnus:
        raise BloombergIndisponibleError(
            f"Champ(s) Bloomberg non simule(s) par ce stub local : {inconnus}. "
            f"Champs disponibles : {sorted(CHAMPS_DOCUMENTES)}"
        )

    reference = _generer_reference_marche(tickers_liste, graine=graine)
    return reference[champs_liste].copy()


def bdh(
    tickers: str | list[str],
    champs: str | list[str],
    date_debut: str,
    date_fin: str,
    graine: int = 20260101,
) -> pd.DataFrame:
    """
    Simule ``BDH`` (Bloomberg Data History) : serie temporelle d'un champ pour
    un ou plusieurs tickers, entre deux dates (jours ouvres uniquement).

    Returns
    -------
    DataFrame indexe par date, colonnes en MultiIndex (ticker, champ) si
    plusieurs tickers/champs sont demandes, sinon colonnes simples.
    """
    tickers_liste = [tickers] if isinstance(tickers, str) else list(tickers)
    champs_liste = [champs] if isinstance(champs, str) else list(champs)

    inconnus = [c for c in champs_liste if c not in CHAMPS_DOCUMENTES]
    if inconnus:
        raise BloombergIndisponibleError(
            f"Champ(s) Bloomberg non simule(s) par ce stub local : {inconnus}."
        )

    dates = pd.bdate_range(start=date_debut, end=date_fin)
    if len(dates) == 0:
        raise ValueError("Aucun jour ouvre entre date_debut et date_fin.")

    colonnes = {}
    for ticker in tickers_liste:
        rng = np.random.default_rng(abs(hash((ticker, graine))) % (2**32))
        niveau_initial = rng.uniform(50, 300)
        rendements = rng.normal(0.0003, 0.012, size=len(dates))
        px_last = niveau_initial * np.cumprod(1 + rendements)
        vol_glissante = (
            pd.Series(rendements).rolling(20, min_periods=5).std().bfill() * np.sqrt(252) * 100
        )

        for champ in champs_liste:
            if champ == "PX_LAST":
                serie = px_last
            elif champ == "VOLATILITY_90D":
                serie = vol_glissante.to_numpy()
            else:
                # Champs statiques (CRNCY, notation...) repetes sur la periode.
                valeur_statique = bdp(ticker, champ, graine=graine).iloc[0, 0]
                serie = np.repeat(valeur_statique, len(dates))
            colonnes[(ticker, champ)] = serie

    if len(tickers_liste) == 1 and len(champs_liste) == 1:
        cle = (tickers_liste[0], champs_liste[0])
        resultat = pd.DataFrame({champs_liste[0]: colonnes[cle]}, index=dates)
    else:
        resultat = pd.DataFrame(colonnes, index=dates)
        resultat.columns = pd.MultiIndex.from_tuples(resultat.columns, names=["ticker", "champ"])

    resultat.index.name = "date"
    return resultat


def bds(ticker: str, champ: str, graine: int = 20260101, n_lignes: int = 10) -> pd.DataFrame:
    """
    Simule ``BDS`` (Bloomberg Data Set) : donnee "bulk" en liste, par exemple
    la composition ponderee d'un indice (champ ``INDX_MWEIGHT``).

    Ce stub supporte specifiquement le champ ``INDX_MWEIGHT`` (composition
    d'indice pondere) a titre d'exemple ; d'autres champs bulk peuvent etre
    ajoutes en suivant le meme schema.
    """
    if champ != "INDX_MWEIGHT":
        raise BloombergIndisponibleError(
            f"Champ BDS non simule par ce stub local : {champ}. "
            "Seul 'INDX_MWEIGHT' est implemente a titre d'exemple."
        )

    rng = np.random.default_rng(abs(hash((ticker, champ, graine))) % (2**32))
    poids_bruts = rng.dirichlet(np.ones(n_lignes) * 2)

    return pd.DataFrame(
        {
            "Member Ticker and Exchange Code": [f"CONST{i + 1:03d} Equity" for i in range(n_lignes)],
            "Percentage Weight": np.round(poids_bruts * 100, 3),
        }
    )


def documentation_champs() -> pd.DataFrame:
    """Renvoie un DataFrame documentant chaque champ Bloomberg simule (pour la feuille Documentation du classeur)."""
    return pd.DataFrame(
        {"champ": list(CHAMPS_DOCUMENTES.keys()), "description": list(CHAMPS_DOCUMENTES.values())}
    )
