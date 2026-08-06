Attribute VB_Name = "modLimites"

Option Explicit

Public Enum StatutLimite
    StatutOK = 0
    StatutAlerte = 1
    StatutDepassement = 2
End Enum

Private Const COL_CODE As Long = 1
Private Const COL_TYPE As Long = 2
Private Const COL_VALEUR_LIMITE As Long = 3
Private Const COL_SEUIL_ALERTE As Long = 4
Private Const COL_VALEUR_CONSTATEE As Long = 5
Private Const LIGNE_DEPART_DONNEES As Long = 2

Public Function CalculerUtilisation(ByVal valeurConstatee As Double, ByVal valeurLimite As Double) As Double
    If valeurLimite = 0 Then
        CalculerUtilisation = -1
        Exit Function
    End If
    CalculerUtilisation = Abs(valeurConstatee) / Abs(valeurLimite)
End Function

Public Function DeterminerStatut(ByVal utilisation As Double, ByVal seuilAlerte As Double) As StatutLimite
    If utilisation >= 1 Then
        DeterminerStatut = StatutDepassement
    ElseIf utilisation >= seuilAlerte Then
        DeterminerStatut = StatutAlerte
    Else
        DeterminerStatut = StatutOK
    End If
End Function

Public Function LibelleStatut(ByVal statut As StatutLimite) As String
    Select Case statut
        Case StatutOK: LibelleStatut = "OK"
        Case StatutAlerte: LibelleStatut = "ALERTE"
        Case StatutDepassement: LibelleStatut = "DEPASSEMENT"
        Case Else: LibelleStatut = "INCONNU"
    End Select
End Function

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

Private Function SeuilAlerteSecurise(ByVal valeurCell As Variant) As Double
    If IsNumeric(valeurCell) And Not IsEmpty(valeurCell) Then
        SeuilAlerteSecurise = CDbl(valeurCell)
    Else
        SeuilAlerteSecurise = 0.9
    End If
End Function

Private Sub AppliquerCouleurStatut(ByVal cell As Range, ByVal statut As StatutLimite)
    Select Case statut
        Case StatutOK
            cell.Interior.Color = RGB(198, 239, 206)
            cell.Font.Color = RGB(0, 97, 0)
        Case StatutAlerte
            cell.Interior.Color = RGB(255, 235, 156)
            cell.Font.Color = RGB(156, 87, 0)
        Case StatutDepassement
            cell.Interior.Color = RGB(255, 199, 206)
            cell.Font.Color = RGB(156, 0, 6)
    End Select
    cell.Font.Bold = True
End Sub

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
