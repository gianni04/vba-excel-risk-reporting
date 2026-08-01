Attribute VB_Name = "modBloomberg"
'===============================================================================
' Module    : modBloomberg
' Objet     : Wrappers autour des fonctions Bloomberg Excel Add-in (BDP, BDH,
'             BDS) : construction de formules, rafraichissement controle,
'             detection des erreurs de donnees, et FALLBACK EXPLICITE sur un
'             jeu de donnees local lorsque le terminal Bloomberg n'est pas
'             disponible (poste sans add-in, demo, ou hors connexion Bloomberg).
'
' IMPORTANT - PRE-REQUIS BLOOMBERG :
'   Les fonctions BDP / BDH / BDS elles-memes sont fournies par le complement
'   Excel "Bloomberg Excel Add-in" (installe avec un terminal Bloomberg actif
'   et une session BBComm ouverte). Ce module ne reimplemente PAS ces fonctions
'   (impossible sans le terminal) : il construit les CHAINES DE FORMULE a
'   inserer dans les cellules, controle leur rafraichissement, et detecte les
'   erreurs "#N/A Requesting Data" pendant la latence reseau normale de
'   Bloomberg. Si l'add-in n'est pas present, RafraichirBloomberg bascule
'   automatiquement sur ChargerDonneesLocales pour que le classeur reste
'   utilisable en depannage / demonstration / formation.
'
' Auteur    : Gianni Pilotti
'===============================================================================

Option Explicit

' Delai (en secondes) entre deux verifications du rafraichissement Bloomberg.
Private Const DELAI_VERIF_SECONDES As Double = 0.5

'===============================================================================
' ConstruireFormuleBDP
' Construit la chaine de formule BDP (Bloomberg Data Point - donnee statique
' "point in time") avec overrides optionnels.
'   ticker  : identifiant Bloomberg, ex "AAPL US Equity", "FR0000131104 Equity"
'   champ   : mnemonique de champ Bloomberg, ex "PX_LAST", "CRNCY", "DUR_ADJ_MID"
'   overrides : dictionnaire "CHAMP=VALEUR" separes par ";", ex
'               "EQY_FUND_YEAR=2025;CURRENCY=EUR" (peut etre vide)
' Exemple d'utilisation : Range("B2").Formula = ConstruireFormuleBDP("AAPL US Equity", "PX_LAST", "")
'===============================================================================
Public Function ConstruireFormuleBDP(ByVal ticker As String, ByVal champ As String, Optional ByVal overrides As String = "") As String
    Dim formule As String

    formule = "=BDP(""" & EchapperGuillemets(ticker) & """,""" & EchapperGuillemets(champ) & """"

    If Len(Trim$(overrides)) > 0 Then
        formule = formule & "," & ConstruireArgumentsOverride(overrides)
    End If

    formule = formule & ")"
    ConstruireFormuleBDP = formule
End Function

'===============================================================================
' ConstruireFormuleBDH
' Construit la chaine de formule BDH (Bloomberg Data History) pour une serie
' historique entre deux dates.
'   ticker    : identifiant Bloomberg
'   champ     : mnemonique de champ, ex "PX_LAST"
'   dateDebut : date de debut au format "AAAAMMJJ" (ex "20240101")
'   dateFin   : date de fin au format "AAAAMMJJ" (ex "20241231")
'   overrides : ex "Days=W;Fill=P" (frequence hebdo, remplissage precedent)
' Exemple : Range("B2").Formula = ConstruireFormuleBDH("AAPL US Equity","PX_LAST","20240101","20241231","")
'===============================================================================
Public Function ConstruireFormuleBDH(ByVal ticker As String, ByVal champ As String, ByVal dateDebut As String, ByVal dateFin As String, Optional ByVal overrides As String = "") As String
    Dim formule As String

    formule = "=BDH(""" & EchapperGuillemets(ticker) & """,""" & EchapperGuillemets(champ) & """,""" & dateDebut & """,""" & dateFin & """"

    If Len(Trim$(overrides)) > 0 Then
        formule = formule & "," & ConstruireArgumentsOverride(overrides)
    End If

    formule = formule & ")"
    ConstruireFormuleBDH = formule
End Function

'===============================================================================
' ConstruireFormuleBDS
' Construit la chaine de formule BDS (Bloomberg Data Set) pour des donnees en
' liste (ex : composition d'indice, flux obligataires, actionnariat).
'   ticker : identifiant Bloomberg, ex "SXXP Index"
'   champ  : mnemonique de champ "bulk", ex "INDX_MWEIGHT"
'   overrides : optionnel
' Exemple : Range("B2").Formula = ConstruireFormuleBDS("SXXP Index","INDX_MWEIGHT","")
'===============================================================================
Public Function ConstruireFormuleBDS(ByVal ticker As String, ByVal champ As String, Optional ByVal overrides As String = "") As String
    Dim formule As String

    formule = "=BDS(""" & EchapperGuillemets(ticker) & """,""" & EchapperGuillemets(champ) & """"

    If Len(Trim$(overrides)) > 0 Then
        formule = formule & "," & ConstruireArgumentsOverride(overrides)
    End If

    formule = formule & ")"
    ConstruireFormuleBDS = formule
End Function

'-------------------------------------------------------------------------------
' ConstruireArgumentsOverride
' Transforme une chaine "CHAMP1=VAL1;CHAMP2=VAL2" en arguments d'override
' Bloomberg au format attendu par BDP/BDH/BDS : "CHAMP1","VAL1","CHAMP2","VAL2"
'-------------------------------------------------------------------------------
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

'===============================================================================
' BloombergDisponible
' Detecte si le complement Bloomberg Excel Add-in est charge (donc si les
' fonctions BDP/BDH/BDS sont utilisables). Se base sur la presence du COM
' Add-in dans Application.COMAddIns ; methode robuste qui ne plante pas si
' Bloomberg n'est pas installe du tout.
'===============================================================================
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

'===============================================================================
' EstErreurBloomberg
' Detecte si une valeur de cellule correspond a une erreur de donnee Bloomberg
' typique : "#N/A Requesting Data" (donnee en cours de recuperation), "#N/A
' Field Not Applicable", "#N/A Invalid Security", etc.
'   valeurCellule : Variant contenant la valeur lue dans la cellule
'===============================================================================
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

'===============================================================================
' AttendreRafraichissementBloomberg
' Attend que les cellules Bloomberg de la plage donnee ne renvoient plus
' "#N/A Requesting Data", avec timeout pour ne jamais bloquer indefiniment
' Excel (cas d'un poste sans connexion Bloomberg active).
'   plage         : Range contenant des formules BDP/BDH/BDS
'   timeoutSecondes : duree maximale d'attente (defaut 15 secondes)
' Renvoie True si toutes les cellules sont resolues avant le timeout.
'===============================================================================
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

'===============================================================================
' RafraichirBloomberg
' Routine de rafraichissement controle : force le recalcul des formules
' Bloomberg de la feuille active (ou de la plage fournie), attend leur
' resolution, et bascule AUTOMATIQUEMENT sur le jeu de donnees local
' (ChargerDonneesLocales) si Bloomberg n'est pas disponible ou si le
' rafraichissement time-out. Ainsi le classeur reste utilisable pour la
' production du reporting meme sans terminal Bloomberg (formation, poste de
' secours, weekend sans session BBComm).
'   feuilleCible : feuille contenant les formules Bloomberg et la zone de repli
'   plageBloomberg : plage a rafraichir / verifier
'   plageRepliLocal : plage de destination du jeu de donnees local de secours
'===============================================================================
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

'===============================================================================
' ChargerDonneesLocales
' Fallback explicite : alimente la plage cible avec un jeu de donnees de
' marche local (statique) permettant de continuer a produire le reporting
' quand le terminal Bloomberg est indisponible. Les valeurs sont clairement
' identifiees comme des donnees de secours (pas des cours de marche reels) via
' un commentaire insere en premiere cellule.
'   plageCible : plage de destination (premiere cellule = coin superieur gauche)
'===============================================================================
Public Sub ChargerDonneesLocales(ByVal plageCible As Range)
    On Error GoTo GestionErreur
    Dim donnees As Variant
    Dim i As Long, j As Long

    ' Jeu de donnees de secours : PX_LAST, VOLATILITY_90D, CRNCY par ligne de
    ' position, dans le meme ordre que le referentiel positions du classeur.
    ' A adapter/etendre selon le referentiel reel de positions.
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

'-------------------------------------------------------------------------------
' LogBloomberg : journalisation simple dans la fenetre Immediate (Ctrl+G), pour
' tracer les bascules Bloomberg <-> donnees locales pendant l'exploitation.
'-------------------------------------------------------------------------------
Private Sub LogBloomberg(ByVal message As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:mm:ss") & " [modBloomberg] " & message
End Sub
