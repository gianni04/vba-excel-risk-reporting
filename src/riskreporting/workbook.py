"""
workbook.py
===========

Construction du classeur Excel de reporting risque multi-feuilles avec
xlsxwriter : Synthese, Positions, Historique, Limites, Graphiques,
Documentation. Les indicateurs de la feuille Synthese sont calcules par des
FORMULES EXCEL NATIVES ecrites directement dans les cellules (pas des valeurs
pre-calculees en Python) afin que le classeur reste vivant : modifier une
position ou une VL recalcule automatiquement tous les indicateurs, exactement
comme le ferait un vrai reporting middle-office.

Fonctions Excel avancees utilisees : SUMPRODUCT, INDEX/MATCH, XLOOKUP,
PERCENTILE.INC, STDEV.S, SUMIFS, AVERAGEIF, SLOPE, IFERROR, COUNTIF, COUNTA.
"""

from __future__ import annotations

from typing import Any

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_range

# ---------------------------------------------------------------------------
# Constantes de mise en forme
# ---------------------------------------------------------------------------

COULEUR_ENTETE = "#1F4E78"
COULEUR_ENTETE_TEXTE = "#FFFFFF"
COULEUR_OK = "#C6EFCE"
COULEUR_OK_TEXTE = "#006100"
COULEUR_ALERTE = "#FFEB9C"
COULEUR_ALERTE_TEXTE = "#9C5700"
COULEUR_DEPASSEMENT = "#FFC7CE"
COULEUR_DEPASSEMENT_TEXTE = "#9C0006"


def _formats(wb: xlsxwriter.Workbook) -> dict[str, Any]:
    return {
        "titre": wb.add_format({"bold": True, "font_size": 16, "font_color": COULEUR_ENTETE}),
        "sous_titre": wb.add_format({"bold": True, "font_size": 11, "italic": True}),
        "entete": wb.add_format(
            {
                "bold": True,
                "bg_color": COULEUR_ENTETE,
                "font_color": COULEUR_ENTETE_TEXTE,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        ),
        "libelle": wb.add_format({"bold": True}),
        "pourcent": wb.add_format({"num_format": "0.00%"}),
        "pourcent1": wb.add_format({"num_format": "0.0%"}),
        "nombre2": wb.add_format({"num_format": "#,##0.00"}),
        "monnaie": wb.add_format({"num_format": "#,##0 €"}),
        "date": wb.add_format({"num_format": "dd/mm/yyyy"}),
        "ok": wb.add_format({"bg_color": COULEUR_OK, "font_color": COULEUR_OK_TEXTE, "bold": True, "align": "center"}),
        "alerte": wb.add_format(
            {"bg_color": COULEUR_ALERTE, "font_color": COULEUR_ALERTE_TEXTE, "bold": True, "align": "center"}
        ),
        "depassement": wb.add_format(
            {"bg_color": COULEUR_DEPASSEMENT, "font_color": COULEUR_DEPASSEMENT_TEXTE, "bold": True, "align": "center"}
        ),
        "bordure": wb.add_format({"border": 1}),
        "note": wb.add_format({"italic": True, "font_size": 9, "font_color": "#666666"}),
    }


def _col_letter(idx: int) -> str:
    """Convertit un indice de colonne 0-based en lettre(s) de colonne Excel."""
    letters = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def construire_classeur(
    chemin_sortie: str,
    positions: pd.DataFrame,
    historique: pd.DataFrame,
    limites: pd.DataFrame,
    nom_portefeuille: str,
    date_valorisation: pd.Timestamp,
    niveau_confiance_var: float = 0.99,
    taux_sans_risque: float = 0.025,
) -> None:
    """
    Construit le classeur complet de reporting risque et l'ecrit sur disque a
    ``chemin_sortie``. Toutes les feuilles sont construites dans une seule
    passe xlsxwriter (le format ne permet pas la relecture/modification
    incrementale - voir validate.py pour la verification post-generation avec
    openpyxl).
    """
    wb = xlsxwriter.Workbook(chemin_sortie, {"nan_inf_to_errors": True})
    fmt = _formats(wb)

    ws_synthese = wb.add_worksheet("Synthese")
    ws_positions = wb.add_worksheet("Positions")
    ws_historique = wb.add_worksheet("Historique")
    ws_limites = wb.add_worksheet("Limites")
    ws_graphiques = wb.add_worksheet("Graphiques")
    ws_doc = wb.add_worksheet("Documentation")

    n_pos = len(positions)
    n_hist = len(historique)

    # ------------------------------------------------------------------
    # Feuille Positions (tableau structure)
    # ------------------------------------------------------------------
    entetes_positions = [
        "ISIN",
        "Libelle",
        "Secteur",
        "Devise",
        "Notation",
        "Exposition (EUR)",
        "Duration",
        "Beta individuel",
        "Poids",
        "Exposition Abs",
    ]
    ligne_entete_pos = 2  # 0-based -> Excel ligne 3
    ligne_derniere_pos = ligne_entete_pos + n_pos  # derniere ligne de donnees (0-based)

    ws_positions.merge_range(0, 0, 0, 9, f"Positions du portefeuille - {nom_portefeuille}", fmt["titre"])
    ws_positions.write(1, 0, "Base de calcul de l'ensemble des expositions et concentrations du reporting.", fmt["note"])

    col_expo = entetes_positions.index("Exposition (EUR)")
    col_expo_abs = entetes_positions.index("Exposition Abs")
    col_poids = entetes_positions.index("Poids")
    lettre_expo = _col_letter(col_expo)
    lettre_expo_abs = _col_letter(col_expo_abs)

    donnees_table = []
    for _, ligne in positions.iterrows():
        donnees_table.append(
            [
                ligne["isin"],
                ligne["libelle"],
                ligne["secteur"],
                ligne["devise"],
                ligne["notation"],
                float(ligne["exposition_eur"]),
                float(ligne["duration"]),
                float(ligne["beta_individuel"]),
                None,  # Poids : formule, ecrite apres add_table
                None,  # Exposition Abs : formule, ecrite apres add_table
            ]
        )

    ws_positions.add_table(
        ligne_entete_pos,
        0,
        ligne_derniere_pos,
        len(entetes_positions) - 1,
        {
            "name": "TablePositions",
            "style": "Table Style Medium 9",
            "data": donnees_table,
            "columns": [{"header": h} for h in entetes_positions],
        },
    )

    for i in range(n_pos):
        excel_row = ligne_entete_pos + 1 + i  # 0-based
        ws_positions.write_formula(
            excel_row, col_expo_abs, f"=ABS({lettre_expo}{excel_row + 1})", fmt["monnaie"]
        )
        ws_positions.write_formula(
            excel_row,
            col_poids,
            f"={lettre_expo}{excel_row + 1}/SUM(TablePositions[Exposition (EUR)])",
            fmt["pourcent"],
        )

    ws_positions.set_column(0, 0, 14)
    ws_positions.set_column(1, 1, 22)
    ws_positions.set_column(2, 4, 14)
    ws_positions.set_column(5, 9, 15)
    ws_positions.freeze_panes(ligne_entete_pos + 1, 1)

    # ------------------------------------------------------------------
    # Feuille Historique
    # ------------------------------------------------------------------
    entetes_hist = [
        "Date",
        "VL Fonds",
        "VL Bench",
        "Rendement Fonds",
        "Rendement Bench",
        "Rendement Actif",
        "Drawdown Fonds",
    ]
    ligne_entete_hist = 2
    ligne_premiere_donnee_hist = ligne_entete_hist + 1  # 0-based
    ligne_derniere_hist = ligne_entete_hist + n_hist  # 0-based, derniere ligne de donnees

    ws_historique.merge_range(0, 0, 0, 6, "Historique de valeur liquidative (VL)", fmt["titre"])
    ws_historique.write(1, 0, "Rendements et drawdown calcules par formule a partir des VL (colonnes B et C).", fmt["note"])

    for c, h in enumerate(entetes_hist):
        ws_historique.write(ligne_entete_hist, c, h, fmt["entete"])

    for i, (date_idx, ligne) in enumerate(historique.iterrows()):
        excel_row = ligne_premiere_donnee_hist + i  # 0-based
        ws_historique.write_datetime(excel_row, 0, pd.Timestamp(date_idx).to_pydatetime(), fmt["date"])
        ws_historique.write_number(excel_row, 1, float(ligne["vl_fonds"]), fmt["nombre2"])
        ws_historique.write_number(excel_row, 2, float(ligne["vl_bench"]), fmt["nombre2"])

        if i == 0:
            ws_historique.write_formula(excel_row, 3, '=""', fmt["pourcent"])
            ws_historique.write_formula(excel_row, 4, '=""', fmt["pourcent"])
        else:
            ws_historique.write_formula(
                excel_row, 3, f"=IFERROR((B{excel_row + 1}/B{excel_row})-1,\"\")", fmt["pourcent"]
            )
            ws_historique.write_formula(
                excel_row, 4, f"=IFERROR((C{excel_row + 1}/C{excel_row})-1,\"\")", fmt["pourcent"]
            )

        ws_historique.write_formula(
            excel_row, 5, f"=IFERROR(D{excel_row + 1}-E{excel_row + 1},\"\")", fmt["pourcent"]
        )
        ws_historique.write_formula(
            excel_row,
            6,
            f"=B{excel_row + 1}/MAX($B${ligne_premiere_donnee_hist + 1}:B{excel_row + 1})-1",
            fmt["pourcent"],
        )

    ws_historique.set_column(0, 0, 12)
    ws_historique.set_column(1, 6, 15)
    ws_historique.freeze_panes(ligne_premiere_donnee_hist, 1)

    # Plages utiles (references Excel, 1-based) pour noms definis et formules
    plage_vl_fonds = f"Historique!$B${ligne_premiere_donnee_hist + 1}:$B${ligne_derniere_hist}"
    plage_vl_bench = f"Historique!$C${ligne_premiere_donnee_hist + 1}:$C${ligne_derniere_hist}"
    plage_rend_fonds = f"Historique!$D${ligne_premiere_donnee_hist + 1}:$D${ligne_derniere_hist}"
    plage_rend_bench = f"Historique!$E${ligne_premiere_donnee_hist + 1}:$E${ligne_derniere_hist}"
    plage_rend_actif = f"Historique!$F${ligne_premiere_donnee_hist + 1}:$F${ligne_derniere_hist}"
    plage_drawdown = f"Historique!$G${ligne_premiere_donnee_hist + 1}:$G${ligne_derniere_hist}"

    # ------------------------------------------------------------------
    # Noms definis (defined names)
    # ------------------------------------------------------------------
    wb.define_name("RendementsFonds", f"={plage_rend_fonds}")
    wb.define_name("RendementsBench", f"={plage_rend_bench}")
    wb.define_name("RendementActif", f"={plage_rend_actif}")
    wb.define_name("VLFonds", f"={plage_vl_fonds}")
    wb.define_name("VLBench", f"={plage_vl_bench}")
    wb.define_name("DrawdownFonds", f"={plage_drawdown}")
    wb.define_name("NiveauConfianceVaR", "=Synthese!$D$5")
    wb.define_name("TauxSansRisque", "=Synthese!$D$6")

    # ------------------------------------------------------------------
    # Feuille Synthese
    # ------------------------------------------------------------------
    ws_synthese.merge_range(0, 0, 0, 4, f"Reporting Risque Quotidien - {nom_portefeuille}", fmt["titre"])
    ws_synthese.write(1, 0, "Portefeuille :", fmt["libelle"])
    ws_synthese.write(1, 1, nom_portefeuille)
    ws_synthese.write(2, 0, "Date de valorisation :", fmt["libelle"])
    ws_synthese.write_datetime(2, 1, pd.Timestamp(date_valorisation).to_pydatetime(), fmt["date"])
    ws_synthese.write(3, 0, "Nombre de jours d'historique :", fmt["libelle"])
    ws_synthese.write_formula(3, 1, "=COUNT(RendementsFonds)")

    ws_synthese.write(4, 0, "Niveau de confiance VaR :", fmt["libelle"])
    ws_synthese.write_number(4, 3, niveau_confiance_var, fmt["pourcent"])
    ws_synthese.write(5, 0, "Taux sans risque annuel :", fmt["libelle"])
    ws_synthese.write_number(5, 3, taux_sans_risque, fmt["pourcent"])

    ligne_entete_indic = 8  # 0-based -> Excel ligne 9
    ws_synthese.write(ligne_entete_indic, 0, "Code", fmt["entete"])
    ws_synthese.write(ligne_entete_indic, 1, "Indicateur", fmt["entete"])
    ws_synthese.write(ligne_entete_indic, 2, "Valeur", fmt["entete"])
    ws_synthese.write(ligne_entete_indic, 3, "Formule Excel", fmt["entete"])

    indicateurs = [
        (
            "VAR_MAX",
            "VaR historique 99% (1 jour)",
            "=IFERROR(-PERCENTILE.INC(RendementsFonds,1-NiveauConfianceVaR),\"N/D\")",
            fmt["pourcent"],
        ),
        (
            "VAR_PARAM",
            "VaR parametrique 99% (1 jour)",
            "=IFERROR(-(AVERAGE(RendementsFonds)-NORM.S.INV(NiveauConfianceVaR)*STDEV.S(RendementsFonds)),\"N/D\")",
            fmt["pourcent"],
        ),
        (
            "ES_975",
            "Expected Shortfall 97,5%",
            "=IFERROR(-AVERAGEIF(RendementsFonds,\"<=\"&PERCENTILE.INC(RendementsFonds,0.025)),\"N/D\")",
            fmt["pourcent"],
        ),
        ("VOL_ANN", "Volatilite annualisee", "=IFERROR(STDEV.S(RendementsFonds)*SQRT(252),\"N/D\")", fmt["pourcent"]),
        ("TE_MAX", "Tracking Error annualisee", "=IFERROR(STDEV.S(RendementActif)*SQRT(252),\"N/D\")", fmt["pourcent"]),
        ("BETA", "Beta vs benchmark", "=IFERROR(SLOPE(RendementsFonds,RendementsBench),\"N/D\")", fmt["nombre2"]),
        (
            "SHARPE",
            "Ratio de Sharpe",
            "=IFERROR((AVERAGE(RendementsFonds)*252-TauxSansRisque)/(STDEV.S(RendementsFonds)*SQRT(252)),\"N/D\")",
            fmt["nombre2"],
        ),
        (
            "INFO_RATIO",
            "Ratio d'information",
            "=IFERROR((AVERAGE(RendementActif)*252)/(STDEV.S(RendementActif)*SQRT(252)),\"N/D\")",
            fmt["nombre2"],
        ),
        ("MAXDD", "Max Drawdown", "=IFERROR(MIN(DrawdownFonds),\"N/D\")", fmt["pourcent"]),
        (
            "CONCENTRATION_MAX",
            "Concentration emetteur max",
            "=IFERROR(MAX(TablePositions[Exposition Abs])/SUM(TablePositions[Exposition Abs]),\"N/D\")",
            fmt["pourcent"],
        ),
        (
            "EXPO_BRUTE",
            "Exposition brute",
            "=IFERROR(SUMPRODUCT(ABS(TablePositions[Exposition (EUR)])),\"N/D\")",
            fmt["monnaie"],
        ),
        ("EXPO_NETTE", "Exposition nette", "=IFERROR(SUM(TablePositions[Exposition (EUR)]),\"N/D\")", fmt["monnaie"]),
        ("NB_LIGNES", "Nombre de lignes en portefeuille", "=COUNTA(TablePositions[ISIN])", fmt["nombre2"]),
        (
            "SECTEUR_MAX",
            "Secteur le plus expose",
            "=IFERROR(INDEX(TablePositions[Secteur],MATCH(MAX(TablePositions[Exposition Abs]),TablePositions[Exposition Abs],0)),\"N/D\")",
            None,
        ),
    ]

    for offset, (code, libelle, formule, format_valeur) in enumerate(indicateurs):
        r = ligne_entete_indic + 1 + offset
        ws_synthese.write(r, 0, code, fmt["bordure"])
        ws_synthese.write(r, 1, libelle, fmt["bordure"])
        ws_synthese.write_formula(r, 2, formule, format_valeur if format_valeur else fmt["bordure"])
        ws_synthese.write(r, 3, formule, fmt["note"])

    ligne_derniere_indic = ligne_entete_indic + len(indicateurs)  # 0-based, derniere ligne ecrite

    # Nombre de depassements de limites (rempli apres construction de la
    # feuille Limites, mais la formule peut etre ecrite des maintenant car
    # xlsxwriter n'a pas besoin d'ordre d'ecriture particulier entre feuilles).
    r_depassements = ligne_derniere_indic + 2
    ws_synthese.write(r_depassements, 0, "NB_DEPASSEMENTS", fmt["bordure"])
    ws_synthese.write(r_depassements, 1, "Nombre de depassements de limites", fmt["bordure"])
    ws_synthese.write_formula(
        r_depassements,
        2,
        "=IFERROR(COUNTIF(Limites!$H$5:$H$100,\"DEPASSEMENT\"),0)",
        fmt["nombre2"],
    )
    ws_synthese.write(r_depassements, 3, '=COUNTIF(Limites!$H$5:$H$100,"DEPASSEMENT")', fmt["note"])

    ws_synthese.set_column(0, 0, 20)
    ws_synthese.set_column(1, 1, 32)
    ws_synthese.set_column(2, 2, 16)
    ws_synthese.set_column(3, 3, 60)
    ws_synthese.freeze_panes(ligne_entete_indic + 1, 0)

    # Lignes de reference pour Limites!F (Code -> Valeur), utilisees par XLOOKUP
    plage_codes_synthese = f"Synthese!$A${ligne_entete_indic + 2}:$A${ligne_derniere_indic + 1}"
    plage_valeurs_synthese = f"Synthese!$C${ligne_entete_indic + 2}:$C${ligne_derniere_indic + 1}"

    # ------------------------------------------------------------------
    # Feuille Limites
    # ------------------------------------------------------------------
    entetes_limites = ["Code", "Type", "Libelle", "Limite", "Seuil alerte", "Valeur constatee", "Utilisation", "Statut"]
    ligne_entete_lim = 3  # 0-based -> Excel ligne 4
    n_lim = len(limites)

    ws_limites.merge_range(0, 0, 0, 7, "Referentiel et controle des limites de risque", fmt["titre"])
    ws_limites.write(
        1,
        0,
        "Valeur constatee calculee par formule (XLOOKUP vers Synthese pour VAR/TE/CONCENTRATION, "
        "SUMIFS sur Positions pour les limites sectorielles).",
        fmt["note"],
    )

    for c, h in enumerate(entetes_limites):
        ws_limites.write(ligne_entete_lim, c, h, fmt["entete"])

    for i, (_, ligne) in enumerate(limites.iterrows()):
        excel_row = ligne_entete_lim + 1 + i  # 0-based
        r1 = excel_row + 1  # 1-based pour les formules
        ws_limites.write(excel_row, 0, ligne["code"])
        ws_limites.write(excel_row, 1, ligne["type"])
        ws_limites.write(excel_row, 2, ligne["libelle"])
        ws_limites.write_number(excel_row, 3, float(ligne["limite"]), fmt["pourcent"])
        ws_limites.write_number(excel_row, 4, float(ligne["seuil_alerte"]), fmt["pourcent"])

        formule_valeur = (
            f'=IFERROR(IF($B{r1}="SECTEUR",'
            f'SUMIFS(TablePositions[Exposition Abs],TablePositions[Secteur],PROPER(MID($A{r1},9,50)))'
            f"/SUM(TablePositions[Exposition Abs]),"
            f'XLOOKUP($A{r1},{plage_codes_synthese},{plage_valeurs_synthese},"N/D")),"N/D")'
        )
        ws_limites.write_formula(excel_row, 5, formule_valeur, fmt["pourcent"])
        ws_limites.write_formula(
            excel_row, 6, f'=IFERROR(ABS(F{r1})/ABS(D{r1}),"N/D")', fmt["pourcent"]
        )
        ws_limites.write_formula(
            excel_row, 7, f'=IF(G{r1}>=1,"DEPASSEMENT",IF(G{r1}>=E{r1},"ALERTE","OK"))'
        )

    ligne_derniere_lim = ligne_entete_lim + n_lim  # 0-based

    # Mise en forme conditionnelle "feux tricolores" par formule (seuil par ligne)
    plage_statut = xl_range(ligne_entete_lim + 1, 7, ligne_derniere_lim, 7)
    premiere_ligne_formule = ligne_entete_lim + 2  # 1-based de la premiere ligne de donnees
    ws_limites.conditional_format(
        plage_statut,
        {
            "type": "formula",
            "criteria": f"=$G{premiere_ligne_formule}>=1",
            "format": fmt["depassement"],
        },
    )
    ws_limites.conditional_format(
        plage_statut,
        {
            "type": "formula",
            "criteria": f"=AND($G{premiere_ligne_formule}>=$E{premiere_ligne_formule},$G{premiere_ligne_formule}<1)",
            "format": fmt["alerte"],
        },
    )
    ws_limites.conditional_format(
        plage_statut,
        {
            "type": "formula",
            "criteria": f"=$G{premiere_ligne_formule}<$E{premiere_ligne_formule}",
            "format": fmt["ok"],
        },
    )

    ws_limites.set_column(0, 0, 20)
    ws_limites.set_column(1, 1, 14)
    ws_limites.set_column(2, 2, 34)
    ws_limites.set_column(3, 7, 14)
    ws_limites.freeze_panes(ligne_entete_lim + 1, 1)
    ws_limites.protect()  # feuille protegee : seules les formules restent modifiables via l'onglet Revision

    # ------------------------------------------------------------------
    # Feuille Graphiques
    # ------------------------------------------------------------------
    ws_graphiques.merge_range(0, 0, 0, 5, "Graphiques de reporting", fmt["titre"])

    # Tableau d'aide : exposition brute par secteur (pour le graphique en barres)
    secteurs_uniques = sorted(positions["secteur"].unique())
    ligne_entete_secteurs = 2
    ws_graphiques.write(ligne_entete_secteurs, 0, "Secteur", fmt["entete"])
    ws_graphiques.write(ligne_entete_secteurs, 1, "Exposition brute", fmt["entete"])
    for i, secteur in enumerate(secteurs_uniques):
        r = ligne_entete_secteurs + 1 + i
        ws_graphiques.write(r, 0, secteur)
        ws_graphiques.write_formula(
            r,
            1,
            f'=SUMIFS(TablePositions[Exposition Abs],TablePositions[Secteur],"{secteur}")',
            fmt["monnaie"],
        )
    ligne_derniere_secteur = ligne_entete_secteurs + len(secteurs_uniques)

    graphique_vl = wb.add_chart({"type": "line"})
    graphique_vl.add_series(
        {
            "name": "Fonds",
            "categories": f"=Historique!$A${ligne_premiere_donnee_hist + 1}:$A${ligne_derniere_hist}",
            "values": f"=Historique!$B${ligne_premiere_donnee_hist + 1}:$B${ligne_derniere_hist}",
            "line": {"width": 2},
        }
    )
    graphique_vl.add_series(
        {
            "name": "Benchmark",
            "categories": f"=Historique!$A${ligne_premiere_donnee_hist + 1}:$A${ligne_derniere_hist}",
            "values": f"=Historique!$C${ligne_premiere_donnee_hist + 1}:$C${ligne_derniere_hist}",
            "line": {"width": 2, "dash_type": "dash"},
        }
    )
    graphique_vl.set_title({"name": "Evolution de la valeur liquidative"})
    graphique_vl.set_x_axis({"name": "Date"})
    graphique_vl.set_y_axis({"name": "VL (base 100)"})
    graphique_vl.set_size({"width": 640, "height": 360})
    ws_graphiques.insert_chart(ligne_entete_secteurs, 4, graphique_vl)

    graphique_contrib = wb.add_chart({"type": "bar"})
    graphique_contrib.add_series(
        {
            "name": "Exposition brute par secteur",
            "categories": f"=Graphiques!$A${ligne_entete_secteurs + 2}:$A${ligne_derniere_secteur + 1}",
            "values": f"=Graphiques!$B${ligne_entete_secteurs + 2}:$B${ligne_derniere_secteur + 1}",
        }
    )
    graphique_contrib.set_title({"name": "Contribution au risque par secteur (exposition brute)"})
    graphique_contrib.set_size({"width": 640, "height": 360})
    ws_graphiques.insert_chart(ligne_entete_secteurs + 20, 4, graphique_contrib)

    ws_graphiques.set_column(0, 1, 22)

    # ------------------------------------------------------------------
    # Feuille Documentation
    # ------------------------------------------------------------------
    ws_doc.merge_range(0, 0, 0, 3, "Documentation du reporting", fmt["titre"])
    ws_doc.set_column(0, 0, 28)
    ws_doc.set_column(1, 1, 70)

    texte_doc = [
        ("Objet", "Reporting risque quotidien : VaR, ES, TE, beta, drawdown, controle de limites."),
        (
            "Methodologie VaR historique",
            "Percentile empirique des rendements passes (PERCENTILE.INC), sans hypothese de distribution.",
        ),
        (
            "Methodologie VaR parametrique",
            "Hypothese de rendements gaussiens, quantile via NORM.S.INV, echelle racine du temps.",
        ),
        (
            "Expected Shortfall",
            "Moyenne des pertes au-dela du seuil de VaR (moyenne de la queue de distribution).",
        ),
        ("Tracking Error", "Ecart-type annualise des rendements actifs (fonds - benchmark)."),
        ("Beta", "Pente de la regression lineaire des rendements du fonds sur ceux du benchmark (fonction SLOPE)."),
        (
            "Bloomberg",
            "La collecte de donnees de marche en conditions reelles s'appuie sur le Bloomberg Excel "
            "Add-in (fonctions BDP/BDH/BDS, voir vba/modBloomberg.bas). Cet add-in NECESSITE un terminal "
            "Bloomberg actif et une session BBComm ouverte. Le module bloomberg_stub.py de ce depot simule "
            "localement ces fonctions pour permettre de developper et tester le reporting hors terminal.",
        ),
        (
            "Limites de risque",
            "Chaque limite est controlee automatiquement (feuille Limites) : statut OK / ALERTE / "
            "DEPASSEMENT selon le taux d'utilisation par rapport au seuil d'alerte parametre.",
        ),
        (
            "Avertissement",
            "Donnees entierement synthetiques, generees a des fins de demonstration. Ne constituent "
            "en aucun cas des donnees de marche reelles ni un conseil en investissement.",
        ),
    ]
    for i, (titre, texte) in enumerate(texte_doc):
        r = 2 + i
        ws_doc.write(r, 0, titre, fmt["libelle"])
        ws_doc.write(r, 1, texte)

    ligne_champs = 2 + len(texte_doc) + 2
    ws_doc.write(ligne_champs, 0, "Champs Bloomberg documentes", fmt["sous_titre"])
    ws_doc.write(ligne_champs + 1, 0, "Champ", fmt["entete"])
    ws_doc.write(ligne_champs + 1, 1, "Description", fmt["entete"])

    from riskreporting.bloomberg_stub import documentation_champs

    df_champs = documentation_champs()
    for i, (_, ligne) in enumerate(df_champs.iterrows()):
        r = ligne_champs + 2 + i
        ws_doc.write(r, 0, ligne["champ"])
        ws_doc.write(r, 1, ligne["description"])

    wb.close()
