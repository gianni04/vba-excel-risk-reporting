"""Generation d'un portefeuille synthetique et d'un historique de VL / benchmark, deterministe (graine fixe) pour la reproductibilite."""

from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

GRAINE_DEFAUT = 20260101

SECTEURS = [
    "Technologie",
    "Sante",
    "Energie",
    "Finance",
    "Industrie",
    "Consommation",
    "Telecommunications",
    "Materiaux",
]

NOTATIONS = ["AAA", "AA", "A", "BBB", "BB", "B"]

DEVISES = ["EUR", "USD", "GBP", "CHF"]


@dataclass
class Portefeuille:
    """Conteneur simple pour un portefeuille synthetique de positions."""

    nom: str
    devise_base: str
    date_valorisation: pd.Timestamp
    positions: pd.DataFrame = field(repr=False)

    @property
    def valeur_totale(self) -> float:
        return float(self.positions["exposition_eur"].sum())

    @property
    def exposition_brute(self) -> float:
        return float(self.positions["exposition_eur"].abs().sum())


def generer_positions(
    n_positions: int = 25, graine: int = GRAINE_DEFAUT
) -> pd.DataFrame:
    """Genere un tableau de positions synthetiques (ISIN, secteur, exposition, devise, notation, duration, beta individuel) representatif d'un portefeuille actions/obligations diversifie."""
    rng = np.random.default_rng(graine)

    isins = [f"XX{rng.integers(1000000000, 9999999999):010d}"[:12] for _ in range(n_positions)]
    libelles = [f"Emetteur {i + 1:02d}" for i in range(n_positions)]
    secteurs = rng.choice(SECTEURS, size=n_positions)
    devises = rng.choice(DEVISES, size=n_positions, p=[0.55, 0.30, 0.10, 0.05])
    notations = rng.choice(
        NOTATIONS, size=n_positions, p=[0.05, 0.15, 0.30, 0.30, 0.15, 0.05]
    )

    expositions = rng.lognormal(mean=13.5, sigma=0.7, size=n_positions)
    signes = np.where(rng.random(n_positions) < 0.1, -1.0, 1.0)
    expositions = expositions * signes

    durations = np.round(rng.uniform(0.5, 9.0, size=n_positions), 2)
    beta_individuel = np.round(rng.normal(1.0, 0.35, size=n_positions), 2)

    df = pd.DataFrame(
        {
            "isin": isins,
            "libelle": libelles,
            "secteur": secteurs,
            "devise": devises,
            "notation": notations,
            "exposition_eur": np.round(expositions, 2),
            "duration": durations,
            "beta_individuel": beta_individuel,
        }
    )

    valeur_totale = df["exposition_eur"].sum()
    df["poids"] = df["exposition_eur"] / valeur_totale

    return df


def generer_historique_vl(
    n_jours: int = 260,
    date_fin: str | pd.Timestamp = "2026-07-31",
    vol_annuelle_fonds: float = 0.14,
    vol_annuelle_bench: float = 0.16,
    correlation: float = 0.92,
    derive_annuelle_fonds: float = 0.06,
    derive_annuelle_bench: float = 0.05,
    graine: int = GRAINE_DEFAUT,
) -> pd.DataFrame:
    """Genere un historique quotidien de valeur liquidative (VL) du fonds et de son benchmark, par un processus de rendements gaussiens correles (mouvement brownien geometrique discretise), base 100 au premier jour."""
    rng = np.random.default_rng(graine + 1)

    dates = pd.bdate_range(end=pd.Timestamp(date_fin), periods=n_jours)

    freq = 252
    mu_f = derive_annuelle_fonds / freq
    mu_b = derive_annuelle_bench / freq
    sigma_f = vol_annuelle_fonds / np.sqrt(freq)
    sigma_b = vol_annuelle_bench / np.sqrt(freq)

    cov = np.array([[1.0, correlation], [correlation, 1.0]])
    L = np.linalg.cholesky(cov)
    z = rng.standard_normal((n_jours, 2))
    z_correles = z @ L.T

    rendement_fonds = mu_f + sigma_f * z_correles[:, 0]
    rendement_bench = mu_b + sigma_b * z_correles[:, 1]

    vl_fonds = 100.0 * np.cumprod(1.0 + rendement_fonds)
    vl_bench = 100.0 * np.cumprod(1.0 + rendement_bench)

    df = pd.DataFrame(
        {
            "vl_fonds": vl_fonds,
            "vl_bench": vl_bench,
            "rendement_fonds": rendement_fonds,
            "rendement_bench": rendement_bench,
        },
        index=dates,
    )
    df.index.name = "date"
    return df


def generer_referentiel_limites(portefeuille: pd.DataFrame) -> pd.DataFrame:
    """Genere un referentiel de limites de risque coherent avec le portefeuille fourni (limites larges par rapport a l'utilisation reelle courante, pour que le reporting d'exemple affiche un melange realiste de statuts OK / ALERTE / DEPASSEMENT)."""
    expo_brute = float(portefeuille["exposition_eur"].abs().sum())

    lignes = [
        {
            "code": "VAR_MAX",
            "type": "VAR",
            "libelle": "VaR historique 99% (1 jour)",
            "limite": 0.025,
            "seuil_alerte": 0.85,
        },
        {
            "code": "TE_MAX",
            "type": "TE",
            "libelle": "Tracking error annualisee",
            "limite": 0.04,
            "seuil_alerte": 0.85,
        },
        {
            "code": "CONCENTRATION_MAX",
            "type": "CONCENTRATION",
            "libelle": "Concentration emetteur max",
            "limite": 0.12,
            "seuil_alerte": 0.85,
        },
    ]

    for secteur in sorted(portefeuille["secteur"].unique()):
        expo_secteur = float(
            portefeuille.loc[portefeuille["secteur"] == secteur, "exposition_eur"]
            .abs()
            .sum()
        )
        poids_secteur = expo_secteur / expo_brute if expo_brute else 0.0
        limite = max(0.20, poids_secteur * 1.15)
        lignes.append(
            {
                "code": f"SECTEUR_{secteur.upper()}",
                "type": "SECTEUR",
                "libelle": f"Exposition sectorielle max - {secteur}",
                "limite": round(limite, 4),
                "seuil_alerte": 0.85,
            }
        )

    return pd.DataFrame(lignes)


def jeu_de_donnees_complet(
    n_positions: int = 25, n_jours: int = 260, graine: int = GRAINE_DEFAUT
) -> dict:
    """Point d'entree unique produisant l'ensemble des donnees necessaires a la construction du classeur de reporting (positions, historique VL, limites)."""
    positions = generer_positions(n_positions=n_positions, graine=graine)
    historique = generer_historique_vl(n_jours=n_jours, graine=graine)
    limites = generer_referentiel_limites(positions)

    return {
        "positions": positions,
        "historique": historique,
        "limites": limites,
        "nom_portefeuille": "Fonds Actions Europe Diversifie",
        "devise_base": "EUR",
        "date_valorisation": historique.index[-1],
    }
