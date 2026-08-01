Attribute VB_Name = "modLimites"
'===============================================================================
' Module    : modLimites
' Objet     : Moteur de controle des limites de risque. Lit un referentiel de
'             limites depuis une feuille (VaR max, TE max, exposition
'             sectorielle max, concentration emetteur max, notation minimale),
'             calcule les taux d'utilisation, detecte les depassements
'             (breach) et les seuils d'alerte (warning, ex : 90% de la
'             limite), et journalise chaque controle dans une feuille d'audit
'             horodatee.
' Auteur    : Gianni Pilotti
'
' Structure attendue du referentiel de limites (feuille "Limites"), a partir
' de la ligne 2 (ligne 1 = en-tetes) :
'   Colonne A : Code limite       (ex "VAR_MAX", "TE_MAX", "SECTEUR_ENERGIE")
'   Colonne B : Type de limite    ("VAR", "TE", "SECTEUR", "CONCENTRATION", "NOTATION")
'   Colonne C : Valeur limite     (ex 0,03 pour 3% ; ou "BBB" pour une notation)
'   Colonne D : Seuil d'alerte (%) (ex 0,9 pour une alerte a 90% d'utilisation)
'   Colonne E : Valeur constatee  (calculee par ailleurs, ou par ce module)
'
' Feuille d'audit ("AuditLimites") : chaque controle ajoute une ligne
' horodatee (Date/Heure, Code limite, Valeur constatee, Valeur limite,
' Utilisation %, Statut).
'===============================================================================

Option Explicit

Public Enum StatutLimite
    StatutOK = 0
    StatutAlerte = 1
    StatutDepassement = 2
End Enum

'-------------------------------------------------------------------------------
' Colonnes du referentiel de limites (constantes pour lisibilite / maintenance).
'-------------------------------------------------------------------------------
Private Const COL_CODE As Long = 1
Private Const COL_TYPE As Long = 2
Private Const COL_VALEUR_LIMITE As Long = 3
Private Const COL_SEUIL_ALERTE As Long = 4
Private Const COL_VALEUR_CONSTATEE As Long = 5
Private Const LIGNE_DEPART_DONNEES As Long = 2

'===============================================================================
' CalculerUtilisation
' Calcule le taux d'utilisation d'une limite = valeur constatee / valeur
' limite (en valeur absolue pour rester coherent lorsque des expositions
' negatives sont possibles). Renvoie -1 si la limite est nulle (division
' impossible) ou non numerique (ex : limite de type "NOTATION").
'===============================================================================
Public Function CalculerUtilisation(ByVal valeurConstatee As Double, ByVal valeurLimite As Double) As Double
    If valeurLimite = 0 Then
        CalculerUtilisation = -1
        Exit Function
    End If
    CalculerUtilisation = Abs(valeurConstatee) / Abs(valeurLimite)
End Function

'===============================================================================
' DeterminerStatut
' Determine le statut d'une limite (OK / Alerte / Depassement) a partir du
' taux d'utilisation et du seuil d'alerte (ex : 0,9 = alerte a partir de 90%
' d'utilisation, depassement a partir de 100%).
'===============================================================================
Public Function DeterminerStatut(ByVal utilisation As Double, ByVal seuilAlerte As Double) As StatutLimite
    If utilisation >= 1 Then
        DeterminerStatut = StatutDepassement
    ElseIf utilisation >= seuilAlerte Then
        DeterminerStatut = StatutAlerte
    Else
        DeterminerStatut = StatutOK
    End If
End Function

'===============================================================================
' LibelleStatut : traduit un StatutLimite en libelle francais pour affichage.
'===============================================================================
Public Function LibelleStatut(ByVal statut As StatutLimite) As String
    Select Case statut
        Case StatutOK: LibelleStatut = "OK"
        Case StatutAlerte: LibelleStatut = "ALERTE"
        Case StatutDepassement: LibelleStatut = "DEPASSEMENT"
        Case Else: LibelleStatut = "INCONNU"
    End Select
End Function

'===============================================================================
' ControlerToutesLesLimites
' Parcourt le referentiel de limites present sur feuilleLimites, calcule
' l'utilisation et le statut de chaque ligne, ecrit le resultat en colonnes F
' (utilisation %) et G (statut), et journalise chaque controle dans
' feuilleAudit. Renvoie le nombre de depassements detectes.
'   feuilleLimites : Worksheet contenant le referentiel (voir en-tete module)
'   feuilleAudit   : Worksheet de journalisation (creee/complementee si besoin)
'===============================================================================
Public Function ControlerToutesLesLimites(ByVal feuilleLimites As Worksheet, ByVal feuilleAudit As Worksheet) As Long
    On Error GoTo GestionErreur
    Dim derniereLigne As Long
    Dim ligne As Long
    Dim codeLimite As String, typeLimite As String
    Dim valeurLimiteCell As Variant, valeurConstateeCell As Variant
    Dim valeurLimite As Double, valeurConstatee As Double
    Dim seuilAlerte As Double
    Dim utilisation As Double
    Dim statut As StatutLimite
    Dim nDepassements As Long

    nDepassements = 0
    derniereLigne = feuilleLimites.Cells(feuilleLimites.Rows.Count, COL_CODE).End(xlUp).Row

    If derniereLigne < LIGNE_DEPART_DONNEES Then
        ControlerToutesLesLimites = 0
        Exit Function
    End If

    PreparerFeuilleAudit feuilleAudit

    For ligne = LIGNE_DEPART_DONNEES To derniereLigne
        codeLimite = CStr(feuilleLimites.Cells(ligne, COL_CODE).Value)
        If Len(Trim$(codeLimite)) = 0 Then GoTo LigneSuivante

        typeLimite = CStr(feuilleLimites.Cells(ligne, COL_TYPE).Value)
        valeurLimiteCell = feuilleLimites.Cells(ligne, COL_VALEUR_LIMITE).Value
        valeurConstateeCell = feuilleLimites.Cells(ligne, COL_VALEUR_CONSTATEE).Value
        seuilAlerte = SeuilAlerteSecurise(feuilleLimites.Cells(ligne, COL_SEUIL_ALERTE).Value)

        ' Cas particulier : limite de type "NOTATION" (non numerique) - pas de
        ' calcul de ratio d'utilisation, le controle se fait en amont
        ' (clsPortefeuille.NombrePositionsSousNotation) et ce module se
        ' contente de journaliser le resultat fourni en colonne E (0 ou 1).
        If UCase$(Trim$(typeLimite)) = "NOTATION" Then
            valeurConstatee = CDbl(valeurConstateeCell)
            If valeurConstatee > 0 Then
                statut = StatutDepassement
                utilisation = 1
            Else
                statut = StatutOK
                utilisation = 0
            End If
        Else
            If Not IsNumeric(valeurLimiteCell) Or Not IsNumeric(valeurConstateeCell) Then
                GoTo LigneSuivante
            End If
            valeurLimite = CDbl(valeurLimiteCell)
            valeurConstatee = CDbl(valeurConstateeCell)
            utilisation = CalculerUtilisation(valeurConstatee, valeurLimite)
            If utilisation < 0 Then GoTo LigneSuivante
            statut = DeterminerStatut(utilisation, seuilAlerte)
        End If

        feuilleLimites.Cells(ligne, 6).Value = utilisation
        feuilleLimites.Cells(ligne, 7).Value = LibelleStatut(statut)
        AppliquerCouleurStatut feuilleLimites.Cells(ligne, 7), statut

        JournaliserControle feuilleAudit, codeLimite, valeurConstatee, valeurLimiteCell, utilisation, statut

        If statut = StatutDepassement Then nDepassements = nDepassements + 1

LigneSuivante:
    Next ligne

    ControlerToutesLesLimites = nDepassements
    Exit Function

GestionErreur:
    Err.Raise Err.Number, "modLimites.ControlerToutesLesLimites", Err.Description
End Function

'-------------------------------------------------------------------------------
' SeuilAlerteSecurise : renvoie 0,9 (90%) par defaut si la cellule de seuil
' d'alerte est vide ou non numerique.
'-------------------------------------------------------------------------------
Private Function SeuilAlerteSecurise(ByVal valeurCell As Variant) As Double
    If IsNumeric(valeurCell) And Not IsEmpty(valeurCell) Then
        SeuilAlerteSecurise = CDbl(valeurCell)
    Else
        SeuilAlerteSecurise = 0.9
    End If
End Function

'-------------------------------------------------------------------------------
' AppliquerCouleurStatut : mise en forme "feux tricolores" par code (vert =
' OK, orange = alerte, rouge = depassement) sur la cellule de statut.
'-------------------------------------------------------------------------------
Private Sub AppliquerCouleurStatut(ByVal cell As Range, ByVal statut As StatutLimite)
    Select Case statut
        Case StatutOK
            cell.Interior.Color = RGB(198, 239, 206)   ' vert clair
            cell.Font.Color = RGB(0, 97, 0)
        Case StatutAlerte
            cell.Interior.Color = RGB(255, 235, 156)   ' orange clair
            cell.Font.Color = RGB(156, 87, 0)
        Case StatutDepassement
            cell.Interior.Color = RGB(255, 199, 206)   ' rouge clair
            cell.Font.Color = RGB(156, 0, 6)
    End Select
    cell.Font.Bold = True
End Sub

'-------------------------------------------------------------------------------
' PreparerFeuilleAudit : cree l'en-tete de la feuille d'audit si absente.
'-------------------------------------------------------------------------------
Private Sub PreparerFeuilleAudit(ByVal feuilleAudit As Worksheet)
    If feuilleAudit.Cells(1, 1).Value = "" Then
        feuilleAudit.Cells(1, 1).Value = "Horodatage"
        feuilleAudit.Cells(1, 2).Value = "Code limite"
        feuilleAudit.Cells(1, 3).Value = "Valeur constatee"
        feuilleAudit.Cells(1, 4).Value = "Valeur limite"
        feuilleAudit.Cells(1, 5).Value = "Utilisation %"
        feuilleAudit.Cells(1, 6).Value = "Statut"
        feuilleAudit.Rows(1).Font.Bold = True
    End If
End Sub

'-------------------------------------------------------------------------------
' JournaliserControle : ajoute une ligne horodatee dans la feuille d'audit.
'-------------------------------------------------------------------------------
Private Sub JournaliserControle(ByVal feuilleAudit As Worksheet, ByVal codeLimite As String, ByVal valeurConstatee As Double, _
    ByVal valeurLimiteCell As Variant, ByVal utilisation As Double, ByVal statut As StatutLimite)

    Dim ligneLibre As Long
    ligneLibre = feuilleAudit.Cells(feuilleAudit.Rows.Count, 1).End(xlUp).Row + 1
    If feuilleAudit.Cells(1, 1).Value = "" Then ligneLibre = 2

    feuilleAudit.Cells(ligneLibre, 1).Value = Now
    feuilleAudit.Cells(ligneLibre, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    feuilleAudit.Cells(ligneLibre, 2).Value = codeLimite
    feuilleAudit.Cells(ligneLibre, 3).Value = valeurConstatee
    feuilleAudit.Cells(ligneLibre, 4).Value = valeurLimiteCell
    feuilleAudit.Cells(ligneLibre, 5).Value = utilisation
    feuilleAudit.Cells(ligneLibre, 5).NumberFormat = "0.0%"
    feuilleAudit.Cells(ligneLibre, 6).Value = LibelleStatut(statut)
End Sub
