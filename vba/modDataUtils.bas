Attribute VB_Name = "modDataUtils"
'===============================================================================
' Module    : modDataUtils
' Objet     : Utilitaires de manipulation de donnees communs au reporting
'             risque : import CSV, tri, recherche, calendrier de jours ouvres
'             et jours feries, conversion plage<->tableau, nettoyage des
'             valeurs manquantes.
' Auteur    : Gianni Pilotti
'===============================================================================

Option Explicit

'===============================================================================
' ImporterCSV
' Import robuste d'un fichier CSV dans une feuille, a partir d'une cellule de
' destination. Gere le separateur configurable, les guillemets englobants, et
' les lignes vides (ignorees). Renvoie le nombre de lignes importees, ou -1 en
' cas d'erreur (fichier introuvable, feuille protegee, etc.).
'   cheminFichier : chemin complet du fichier CSV
'   feuilleDestination : Worksheet cible
'   celluleDepart : Range de la cellule de depart (coin superieur gauche)
'   separateur    : caractere separateur (par defaut ";")
'   premiereLigneEntete : True si la 1re ligne est un en-tete (copiee aussi)
'===============================================================================
Public Function ImporterCSV(ByVal cheminFichier As String, ByVal feuilleDestination As Worksheet, ByVal celluleDepart As Range, _
    Optional ByVal separateur As String = ";", Optional ByVal premiereLigneEntete As Boolean = True) As Long

    On Error GoTo GestionErreur
    Dim numeroFichier As Integer
    Dim ligneTexte As String
    Dim champs() As String
    Dim numeroLigne As Long
    Dim numeroColonne As Long
    Dim ligneEcriture As Long

    If Dir(cheminFichier) = "" Then
        ImporterCSV = -1
        Exit Function
    End If

    numeroFichier = FreeFile
    Open cheminFichier For Input As #numeroFichier

    ligneEcriture = 0
    Do While Not EOF(numeroFichier)
        Line Input #numeroFichier, ligneTexte

        If Len(Trim$(ligneTexte)) > 0 Then
            champs = DecouperLigneCSV(ligneTexte, separateur)

            For numeroColonne = LBound(champs) To UBound(champs)
                feuilleDestination.Cells(celluleDepart.Row + ligneEcriture, celluleDepart.Column + numeroColonne).Value = ConvertirChampCSV(champs(numeroColonne))
            Next numeroColonne

            ligneEcriture = ligneEcriture + 1
        End If
    Loop

    Close #numeroFichier
    ImporterCSV = ligneEcriture
    Exit Function

GestionErreur:
    If numeroFichier <> 0 Then
        On Error Resume Next
        Close #numeroFichier
        On Error GoTo 0
    End If
    ImporterCSV = -1
End Function

'-------------------------------------------------------------------------------
' DecouperLigneCSV : decoupe une ligne CSV en tenant compte des guillemets
' englobants (permet des separateurs a l'interieur d'un champ cite).
'-------------------------------------------------------------------------------
Private Function DecouperLigneCSV(ByVal ligne As String, ByVal separateur As String) As String()
    Dim resultat() As String
    Dim nChamps As Long
    Dim i As Long
    Dim champCourant As String
    Dim dansGuillemets As Boolean
    Dim caractere As String

    ReDim resultat(0 To 0)
    nChamps = 0
    champCourant = ""
    dansGuillemets = False

    For i = 1 To Len(ligne)
        caractere = Mid$(ligne, i, 1)

        If caractere = """" Then
            dansGuillemets = Not dansGuillemets
        ElseIf caractere = separateur And Not dansGuillemets Then
            ReDim Preserve resultat(0 To nChamps)
            resultat(nChamps) = champCourant
            nChamps = nChamps + 1
            champCourant = ""
        Else
            champCourant = champCourant & caractere
        End If
    Next i

    ReDim Preserve resultat(0 To nChamps)
    resultat(nChamps) = champCourant

    DecouperLigneCSV = resultat
End Function

'-------------------------------------------------------------------------------
' ConvertirChampCSV : convertit un champ texte en Double si numerique, en Date
' si format de date reconnu, sinon laisse en texte.
'-------------------------------------------------------------------------------
Private Function ConvertirChampCSV(ByVal champ As String) As Variant
    Dim champNettoye As String
    champNettoye = Trim$(champ)

    If Len(champNettoye) = 0 Then
        ConvertirChampCSV = ""
        Exit Function
    End If

    If IsNumeric(champNettoye) Then
        ConvertirChampCSV = CDbl(champNettoye)
    ElseIf IsDate(champNettoye) Then
        ConvertirChampCSV = CDate(champNettoye)
    Else
        ConvertirChampCSV = champNettoye
    End If
End Function

'===============================================================================
' QuickSort
' Tri rapide generique sur un tableau Variant a une dimension, ordre croissant
' ou decroissant. Fonctionne sur des tableaux numeriques ou texte (comparaison
' native VBA).
'   tableau     : tableau Variant a trier (modifie en place)
'   croissant   : True (defaut) pour un tri croissant
'===============================================================================
Public Sub QuickSort(ByRef tableau As Variant, Optional ByVal croissant As Boolean = True)
    If Not IsArray(tableau) Then Exit Sub
    QuickSortRecursif tableau, LBound(tableau), UBound(tableau), croissant
End Sub

Private Sub QuickSortRecursif(ByRef arr As Variant, ByVal gauche As Long, ByVal droite As Long, ByVal croissant As Boolean)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant

    If gauche >= droite Then Exit Sub

    i = gauche
    j = droite
    pivot = arr((gauche + droite) \ 2)

    Do While i <= j
        If croissant Then
            Do While arr(i) < pivot
                i = i + 1
            Loop
            Do While arr(j) > pivot
                j = j - 1
            Loop
        Else
            Do While arr(i) > pivot
                i = i + 1
            Loop
            Do While arr(j) < pivot
                j = j - 1
            Loop
        End If

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If gauche < j Then QuickSortRecursif arr, gauche, j, croissant
    If i < droite Then QuickSortRecursif arr, i, droite, croissant
End Sub

'===============================================================================
' TrierMatriceParColonne
' Trie les LIGNES d'un tableau 2D (Variant) selon les valeurs d'une colonne
' donnee, par insertion (stable, adapte a des matrices de taille moderee comme
' un tableau de positions). Base 1 attendue (tel qu'issu de PlageVersTableau).
'   matrice        : tableau 2D Variant, modifie en place
'   indiceColonne  : indice de la colonne de tri (base 1)
'   croissant      : True (defaut) pour un tri croissant
'===============================================================================
Public Sub TrierMatriceParColonne(ByRef matrice As Variant, ByVal indiceColonne As Long, Optional ByVal croissant As Boolean = True)
    Dim nLignes As Long, nColonnes As Long
    Dim i As Long, j As Long, k As Long
    Dim ligneTemp() As Variant
    Dim doitEchanger As Boolean

    nLignes = UBound(matrice, 1)
    nColonnes = UBound(matrice, 2)

    For i = LBound(matrice, 1) + 1 To nLignes
        j = i
        Do While j > LBound(matrice, 1)
            If croissant Then
                doitEchanger = (matrice(j - 1, indiceColonne) > matrice(j, indiceColonne))
            Else
                doitEchanger = (matrice(j - 1, indiceColonne) < matrice(j, indiceColonne))
            End If

            If Not doitEchanger Then Exit Do

            ReDim ligneTemp(LBound(matrice, 2) To nColonnes)
            For k = LBound(matrice, 2) To nColonnes
                ligneTemp(k) = matrice(j - 1, k)
                matrice(j - 1, k) = matrice(j, k)
                matrice(j, k) = ligneTemp(k)
            Next k

            j = j - 1
        Loop
    Next i
End Sub

'===============================================================================
' RechercheDichotomique
' Recherche dichotomique (binary search) d'une valeur dans un tableau Double
' TRIE par ordre croissant. Renvoie l'indice (base du tableau) si trouve,
' sinon -1.
'   tableau : tableau Double trie par ordre croissant
'   valeurRecherchee : valeur a rechercher
'===============================================================================
Public Function RechercheDichotomique(ByRef tableau() As Double, ByVal valeurRecherchee As Double) As Long
    Dim gauche As Long, droite As Long, milieu As Long

    gauche = LBound(tableau)
    droite = UBound(tableau)

    Do While gauche <= droite
        milieu = (gauche + droite) \ 2

        If tableau(milieu) = valeurRecherchee Then
            RechercheDichotomique = milieu
            Exit Function
        ElseIf tableau(milieu) < valeurRecherchee Then
            gauche = milieu + 1
        Else
            droite = milieu - 1
        End If
    Loop

    RechercheDichotomique = -1
End Function

'===============================================================================
' EstJourOuvre
' Determine si une date est un jour ouvre (lundi-vendredi et absent du
' calendrier de jours feries fourni).
'   dateTest        : date a tester
'   plageJoursFeries : Range contenant les dates de jours feries (optionnel)
'===============================================================================
Public Function EstJourOuvre(ByVal dateTest As Date, Optional ByVal plageJoursFeries As Range = Nothing) As Boolean
    Dim jourSemaine As Integer
    Dim cell As Range

    jourSemaine = Weekday(dateTest, vbMonday) ' 1 = lundi ... 7 = dimanche

    If jourSemaine >= 6 Then
        EstJourOuvre = False
        Exit Function
    End If

    If Not (plageJoursFeries Is Nothing) Then
        For Each cell In plageJoursFeries.Cells
            If IsDate(cell.Value) Then
                If CDate(cell.Value) = DateSerial(Year(dateTest), Month(dateTest), Day(dateTest)) Then
                    EstJourOuvre = False
                    Exit Function
                End If
            End If
        Next cell
    End If

    EstJourOuvre = True
End Function

'===============================================================================
' ProchainJourOuvre
' Renvoie le prochain jour ouvre a partir d'une date donnee (incluse si elle
' est deja ouvree et bIncluDate=True).
'===============================================================================
Public Function ProchainJourOuvre(ByVal dateDepart As Date, Optional ByVal plageJoursFeries As Range = Nothing, Optional ByVal bIncluDate As Boolean = True) As Date
    Dim dateCourante As Date
    Dim compteurSecurite As Long

    dateCourante = dateDepart
    If Not bIncluDate Then dateCourante = dateCourante + 1

    compteurSecurite = 0
    Do While Not EstJourOuvre(dateCourante, plageJoursFeries)
        dateCourante = dateCourante + 1
        compteurSecurite = compteurSecurite + 1
        If compteurSecurite > 3653 Then Exit Do ' garde-fou : 10 ans max
    Loop

    ProchainJourOuvre = dateCourante
End Function

'===============================================================================
' NombreJoursOuvres
' Compte le nombre de jours ouvres entre deux dates incluses, en excluant les
' jours feries fournis.
'===============================================================================
Public Function NombreJoursOuvres(ByVal dateDebut As Date, ByVal dateFin As Date, Optional ByVal plageJoursFeries As Range = Nothing) As Long
    Dim dateCourante As Date
    Dim compteur As Long

    compteur = 0
    dateCourante = dateDebut

    Do While dateCourante <= dateFin
        If EstJourOuvre(dateCourante, plageJoursFeries) Then
            compteur = compteur + 1
        End If
        dateCourante = dateCourante + 1
    Loop

    NombreJoursOuvres = compteur
End Function

'===============================================================================
' PlageVersTableau
' Convertit une Range en tableau Variant 2D (base 1), pour manipulation rapide
' en memoire (evite les acces cellule-par-cellule couteux en performance).
'===============================================================================
Public Function PlageVersTableau(ByVal plage As Range) As Variant
    If plage.Cells.Count = 1 Then
        Dim resultat(1 To 1, 1 To 1) As Variant
        resultat(1, 1) = plage.Value
        PlageVersTableau = resultat
    Else
        PlageVersTableau = plage.Value
    End If
End Function

'===============================================================================
' TableauVersPlage
' Ecrit un tableau Variant 2D (base 1) dans une plage a partir de sa cellule
' superieure gauche.
'===============================================================================
Public Sub TableauVersPlage(ByRef tableau As Variant, ByVal celluleDepart As Range)
    Dim nLignes As Long, nColonnes As Long
    Dim plageDestination As Range

    If Not IsArray(tableau) Then Exit Sub

    nLignes = UBound(tableau, 1) - LBound(tableau, 1) + 1
    nColonnes = UBound(tableau, 2) - LBound(tableau, 2) + 1

    Set plageDestination = celluleDepart.Resize(nLignes, nColonnes)
    plageDestination.Value = tableau
End Sub

'===============================================================================
' NettoyerValeursManquantes
' Remplace, dans une plage, les cellules vides ou en erreur par une valeur de
' remplacement (ex : 0, ou la derniere valeur connue si methode = "LOCF").
'   plage         : Range a nettoyer (modifiee en place)
'   methode       : "ZERO" (defaut) ou "LOCF" (Last Observation Carried Forward)
'   valeurDefaut  : valeur utilisee si methode = "ZERO" (defaut 0)
' Renvoie le nombre de cellules corrigees.
'===============================================================================
Public Function NettoyerValeursManquantes(ByVal plage As Range, Optional ByVal methode As String = "ZERO", Optional ByVal valeurDefaut As Double = 0) As Long
    On Error GoTo GestionErreur
    Dim cell As Range
    Dim derniereValeurConnue As Variant
    Dim nCorrections As Long
    Dim methodeMaj As String

    methodeMaj = UCase$(methode)
    derniereValeurConnue = valeurDefaut
    nCorrections = 0

    For Each cell In plage.Cells
        If EstValeurManquante(cell) Then
            Select Case methodeMaj
                Case "LOCF"
                    cell.Value = derniereValeurConnue
                Case Else
                    cell.Value = valeurDefaut
            End Select
            nCorrections = nCorrections + 1
        Else
            If IsNumeric(cell.Value) Then derniereValeurConnue = cell.Value
        End If
    Next cell

    NettoyerValeursManquantes = nCorrections
    Exit Function

GestionErreur:
    NettoyerValeursManquantes = -1
End Function

Private Function EstValeurManquante(ByVal cell As Range) As Boolean
    If IsEmpty(cell.Value) Then
        EstValeurManquante = True
    ElseIf IsError(cell.Value) Then
        EstValeurManquante = True
    ElseIf VarType(cell.Value) = vbString And Not IsNumeric(cell.Value) Then
        If Len(Trim$(CStr(cell.Value))) = 0 Then
            EstValeurManquante = True
        Else
            EstValeurManquante = False
        End If
    Else
        EstValeurManquante = False
    End If
End Function
