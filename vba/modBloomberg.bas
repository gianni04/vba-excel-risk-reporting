Attribute VB_Name = "modBloomberg"

Option Explicit

Private Const DELAI_VERIF_SECONDES As Double = 0.5

Public Function ConstruireFormuleBDP(ByVal ticker As String, ByVal champ As String, Optional ByVal overrides As String = "") As String
    Dim formule As String

    formule = "=BDP(""" & EchapperGuillemets(ticker) & """,""" & EchapperGuillemets(champ) & """"

    If Len(Trim$(overrides)) > 0 Then
        formule = formule & "," & ConstruireArgumentsOverride(overrides)
    End If

    formule = formule & ")"
    ConstruireFormuleBDP = formule
End Function

Public Function ConstruireFormuleBDH(ByVal ticker As String, ByVal champ As String, ByVal dateDebut As String, ByVal dateFin As String, Optional ByVal overrides As String = "") As String
    Dim formule As String

    formule = "=BDH(""" & EchapperGuillemets(ticker) & """,""" & EchapperGuillemets(champ) & """,""" & dateDebut & """,""" & dateFin & """"

    If Len(Trim$(overrides)) > 0 Then
        formule = formule & "," & ConstruireArgumentsOverride(overrides)
    End If

    formule = formule & ")"
    ConstruireFormuleBDH = formule
End Function

Public Function ConstruireFormuleBDS(ByVal ticker As String, ByVal champ As String, Optional ByVal overrides As String = "") As String
    Dim formule As String

    formule = "=BDS(""" & EchapperGuillemets(ticker) & """,""" & EchapperGuillemets(champ) & """"

    If Len(Trim$(overrides)) > 0 Then
        formule = formule & "," & ConstruireArgumentsOverride(overrides)
    End If

    formule = formule & ")"
    ConstruireFormuleBDS = formule
End Function

Private Function ConstruireArgumentsOverride(ByVal overrides As String) As String
    Dim paires() As String
    Dim i As Long
    Dim cle As String, valeur As String
    Dim resultat As String
    Dim posEgal As Long

    paires = Split(overrides, ";")
    resultat = ""

    For i = LBound(paires) To UBound(paires)
        If Len(Trim$(paires(i))) > 0 Then
            posEgal = InStr(paires(i), "=")
            If posEgal > 0 Then
                cle = Trim$(Left$(paires(i), posEgal - 1))
                valeur = Trim$(Mid$(paires(i), posEgal + 1))
                If Len(resultat) > 0 Then resultat = resultat & ","
                resultat = resultat & """" & EchapperGuillemets(cle) & """,""" & EchapperGuillemets(valeur) & """"
            End If
        End If
    Next i

    ConstruireArgumentsOverride = resultat
End Function

Private Function EchapperGuillemets(ByVal texte As String) As String
    EchapperGuillemets = Replace(texte, """", """""")
End Function

Public Function BloombergDisponible() As Boolean
    On Error GoTo PasDisponible
    Dim i As Long
    Dim nomAddIn As String

    For i = 1 To Application.COMAddIns.Count
        nomAddIn = UCase$(Application.COMAddIns(i).Description & Application.COMAddIns(i).progID)
        If InStr(nomAddIn, "BLOOMBERG") > 0 Then
            If Application.COMAddIns(i).Connect Then
                BloombergDisponible = True
                Exit Function
            End If
        End If
    Next i

    BloombergDisponible = False
    Exit Function

PasDisponible:
    BloombergDisponible = False
End Function

Public Function EstErreurBloomberg(ByVal valeurCellule As Variant) As Boolean
    Dim texte As String

    If IsError(valeurCellule) Then
        EstErreurBloomberg = True
        Exit Function
    End If

    If VarType(valeurCellule) = vbString Then
        texte = UCase$(CStr(valeurCellule))
        If InStr(texte, "#N/A") > 0 Or InStr(texte, "REQUESTING DATA") > 0 _
           Or InStr(texte, "INVALID SECURITY") > 0 Or InStr(texte, "FIELD NOT APPLICABLE") > 0 Then
            EstErreurBloomberg = True
            Exit Function
        End If
    End If

    EstErreurBloomberg = False
End Function

Public Function AttendreRafraichissementBloomberg(ByVal plage As Range, Optional ByVal timeoutSecondes As Double = 15) As Boolean
    On Error GoTo GestionErreur
    Dim tempsDebut As Double
    Dim toutResolu As Boolean
    Dim cell As Range

    tempsDebut = Timer

    Do
        toutResolu = True
        For Each cell In plage.Cells
            If EnCoursDeChargement(cell.Value) Then
                toutResolu = False
                Exit For
            End If
        Next cell

        If toutResolu Then
            AttendreRafraichissementBloomberg = True
            Exit Function
        End If

        If (Timer - tempsDebut) > timeoutSecondes Then
            AttendreRafraichissementBloomberg = False
            Exit Function
        End If

        Application.Wait Now + TimeSerial(0, 0, 0) + (DELAI_VERIF_SECONDES / 86400#)
        DoEvents
    Loop

GestionErreur:
    AttendreRafraichissementBloomberg = False
End Function

Private Function EnCoursDeChargement(ByVal valeurCellule As Variant) As Boolean
    If VarType(valeurCellule) = vbString Then
        EnCoursDeChargement = (InStr(UCase$(CStr(valeurCellule)), "REQUESTING DATA") > 0)
    Else
        EnCoursDeChargement = False
    End If
End Function

Public Sub RafraichirBloomberg(ByVal feuilleCible As Worksheet, ByVal plageBloomberg As Range, Optional ByVal plageRepliLocal As Range = Nothing)
    On Error GoTo GestionErreur
    Dim ok As Boolean

    If Not BloombergDisponible() Then
        LogBloomberg "Bloomberg Add-in non detecte - bascule sur donnees locales."
        If Not (plageRepliLocal Is Nothing) Then
            ChargerDonneesLocales plageRepliLocal
        End If
        Exit Sub
    End If

    Application.CalculateFull
    ok = AttendreRafraichissementBloomberg(plageBloomberg, 15)

    If Not ok Then
        LogBloomberg "Timeout de rafraichissement Bloomberg - bascule sur donnees locales."
        If Not (plageRepliLocal Is Nothing) Then
            ChargerDonneesLocales plageRepliLocal
        End If
    Else
        LogBloomberg "Rafraichissement Bloomberg reussi."
    End If

    Exit Sub

GestionErreur:
    LogBloomberg "Erreur lors du rafraichissement Bloomberg : " & Err.Description
    If Not (plageRepliLocal Is Nothing) Then
        ChargerDonneesLocales plageRepliLocal
    End If
End Sub

Public Sub ChargerDonneesLocales(ByVal plageCible As Range)
    On Error GoTo GestionErreur
    Dim donnees As Variant
    Dim i As Long, j As Long

    donnees = Array( _
        Array("PX_LAST", "VOLATILITY_90D", "CRNCY"), _
        Array(101.25, 18.4, "EUR"), _
        Array(97.8, 22.1, "USD"), _
        Array(54.6, 15.7, "EUR"), _
        Array(210.3, 27.9, "USD"), _
        Array(76.15, 19.2, "EUR") _
    )

    For i = LBound(donnees) To UBound(donnees)
        For j = LBound(donnees(i)) To UBound(donnees(i))
            plageCible.Cells(i - LBound(donnees) + 1, j - LBound(donnees(i)) + 1).Value = donnees(i)(j)
        Next j
    Next i

    If plageCible.Cells(1, 1).Comment Is Nothing Then
        plageCible.Cells(1, 1).AddComment "Donnees de secours locales (fallback hors-Bloomberg) - a ne pas utiliser pour une decision d'investissement."
    End If

    LogBloomberg "Jeu de donnees local charge avec succes (mode fallback)."
    Exit Sub

GestionErreur:
    LogBloomberg "Erreur lors du chargement des donnees locales : " & Err.Description
End Sub

Private Sub LogBloomberg(ByVal message As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:mm:ss") & " [modBloomberg] " & message
End Sub
