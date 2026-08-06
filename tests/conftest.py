"""Fixtures partagees : jeu de donnees synthetique et classeur genere une seule fois par session de test."""

from __future__ import annotations

import pytest

from riskreporting import data, workbook


@pytest.fixture(scope="session")
def jeu_donnees():
    """Jeu de donnees synthetique complet, deterministe (graine fixe)."""
    return data.jeu_de_donnees_complet(n_positions=25, n_jours=260)


@pytest.fixture(scope="session")
def classeur_genere(tmp_path_factory, jeu_donnees):
    """Construit une seule fois le classeur de reporting pour la session de test."""
    repertoire = tmp_path_factory.mktemp("classeur")
    chemin = repertoire / "reporting_risque_test.xlsx"

    workbook.construire_classeur(
        chemin_sortie=str(chemin),
        positions=jeu_donnees["positions"],
        historique=jeu_donnees["historique"],
        limites=jeu_donnees["limites"],
        nom_portefeuille=jeu_donnees["nom_portefeuille"],
        date_valorisation=jeu_donnees["date_valorisation"],
    )
    return chemin
