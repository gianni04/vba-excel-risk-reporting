"""
validate.py
============

Relit un classeur de reporting risque produit par ``workbook.py`` avec
openpyxl et verifie que les feuilles, les noms definis et les formules
attendues sont bien presents. Utilise a la fois par les tests automatises
(``tests/test_workbook.py``) et par ``examples/01_build_report.py`` pour un
controle de sante post-generation.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula


def _texte_formule(valeur) -> str | None:
    """
    Extrait le texte d'une formule quel que soit son type de representation
    openpyxl : chaine simple (formule normale) ou ``ArrayFormula`` (cas des
    fonctions dynamiques recentes comme XLOOKUP, ecrites par xlsxwriter avec
    le prefixe interne ``_xlfn.``).
    """
    if isinstance(valeur, ArrayFormula):
        return valeur.text
    if isinstance(valeur, str) and valeur.startswith("="):
        return valeur
    return None

FEUILLES_ATTENDUES = ["Synthese", "Positions", "Historique", "Limites", "Graphiques", "Documentation"]

NOMS_DEFINIS_ATTENDUS = [
    "RendementsFonds",
    "RendementsBench",
    "RendementActif",
    "VLFonds",
    "VLBench",
    "DrawdownFonds",
    "NiveauConfianceVaR",
    "TauxSansRisque",
]

FONCTIONS_EXCEL_ATTENDUES = [
    "SUMPRODUCT",
    "PERCENTILE.INC",
    "STDEV.S",
    "SUMIFS",
    "IFERROR",
    "XLOOKUP",
    "INDEX",
    "MATCH",
]


@dataclass
class ResultatValidation:
    """Resultat structure d'une validation, avec details des erreurs trouvees."""

    ok: bool
    erreurs: list[str] = field(default_factory=list)
    avertissements: list[str] = field(default_factory=list)

    def lever_si_invalide(self) -> None:
        if not self.ok:
            raise AssertionError("Validation du classeur echouee :\n- " + "\n- ".join(self.erreurs))


def valider_classeur(chemin_fichier: str) -> ResultatValidation:
    """
    Ouvre le classeur avec openpyxl (data_only=False pour lire les formules
    telles quelles) et effectue une serie de controles structurels :
        - toutes les feuilles attendues sont presentes
        - tous les noms definis attendus sont presents
        - le tableau structure "TablePositions" existe sur la feuille Positions
        - chaque fonction Excel avancee attendue apparait au moins une fois
        - les formules de la feuille Limites sont syntaxiquement bien formees
          (parentheses equilibrees, commencent par '=')
        - la feuille Limites porte une protection de feuille active
    """
    erreurs: list[str] = []
    avertissements: list[str] = []

    classeur = openpyxl.load_workbook(chemin_fichier, data_only=False)

    # --- Feuilles ---
    feuilles_presentes = set(classeur.sheetnames)
    for nom in FEUILLES_ATTENDUES:
        if nom not in feuilles_presentes:
            erreurs.append(f"Feuille manquante : {nom}")

    # --- Noms definis ---
    noms_presents = set(classeur.defined_names.keys())
    for nom in NOMS_DEFINIS_ATTENDUS:
        if nom not in noms_presents:
            erreurs.append(f"Nom defini manquant : {nom}")

    # --- Tableau structure ---
    if "Positions" in feuilles_presentes:
        ws_positions = classeur["Positions"]
        noms_tables = list(ws_positions.tables.keys()) if hasattr(ws_positions, "tables") else []
        if "TablePositions" not in noms_tables:
            erreurs.append("Tableau structure 'TablePositions' introuvable sur la feuille Positions")

    # --- Fonctions Excel avancees (recherche dans toutes les formules) ---
    toutes_formules: list[str] = []
    for nom_feuille in classeur.sheetnames:
        ws = classeur[nom_feuille]
        for row in ws.iter_rows():
            for cell in row:
                texte = _texte_formule(cell.value)
                if texte is not None:
                    toutes_formules.append(texte.upper())

    texte_formules_concatene = "\n".join(toutes_formules)
    for fonction in FONCTIONS_EXCEL_ATTENDUES:
        if fonction not in texte_formules_concatene:
            erreurs.append(f"Fonction Excel attendue absente de toutes les formules : {fonction}")

    # --- Bonne formation syntaxique des formules (parentheses equilibrees) ---
    n_formules_verifiees = 0
    for formule in toutes_formules:
        n_formules_verifiees += 1
        if formule.count("(") != formule.count(")"):
            erreurs.append(f"Parentheses desequilibrees dans une formule : {formule[:80]}")
        if not formule.startswith("="):
            erreurs.append(f"Formule ne commencant pas par '=' : {formule[:80]}")

    if n_formules_verifiees == 0:
        erreurs.append("Aucune formule trouvee dans le classeur (attendu : plusieurs dizaines)")

    # --- Protection de la feuille Limites ---
    if "Limites" in feuilles_presentes:
        ws_limites = classeur["Limites"]
        if not ws_limites.protection.sheet:
            avertissements.append("La feuille Limites n'est pas marquee comme protegee")

    return ResultatValidation(ok=(len(erreurs) == 0), erreurs=erreurs, avertissements=avertissements)


def compter_formules(chemin_fichier: str) -> int:
    """Utilitaire simple : compte le nombre total de cellules contenant une formule."""
    classeur = openpyxl.load_workbook(chemin_fichier, data_only=False)
    total = 0
    for nom_feuille in classeur.sheetnames:
        ws = classeur[nom_feuille]
        for row in ws.iter_rows():
            for cell in row:
                if _texte_formule(cell.value) is not None:
                    total += 1
    return total
