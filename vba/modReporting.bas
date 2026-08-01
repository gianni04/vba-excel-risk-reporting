Attribute VB_Name = "modReporting"
'===============================================================================
' Module    : modReporting
' Objet     : Generation du reporting risque quotidien : construction de la
'             feuille de synthese, mise en forme conditionnelle par code
'             (feux tricolores sur les depassements de limites), creation de
'             graphiques via ChartObjects, export PDF, et preparation
'             (JAMAIS envoi) d'un brouillon Outlook.
' Auteur    : Gianni Pilotti
'
' SECURITE  : EnvoyerBrouillonOutlook cree systematiquement le message via
'             MailItem.Display et NE CONTIENT AUCUN APPEL A .Send. Ceci est
'             une decision deliberee : aucune automatisation de ce classeur ne
'             doit pouvoir envoyer un email a l'insu de l'utilisateur.
'===============================================================================

Option Explicit

'===============================================================================
' ConstruireFeuilleSynthese
' Construit (ou reinitialise) la feuille de synthese du reporting risque :
' titre, date de valorisation, tableau des indicateurs cles avec mise en forme
' conditionnelle par code sur les statuts de limites.
'   feuilleSynthese : Worksheet cible (videe puis reconstruite)
'   nomPortefeuille : nom affiche en titre
'   dateValorisation: date du reporting
'===============================================================================
Public Sub ConstruireFeuilleSynthese(ByVal feuilleSynthese As Worksheet, ByVal nomPortefeuille As String, ByVal dateValorisation As Date)
    On Error GoTo GestionErreur

    feuilleSynthese.Cells.Clear
    feuilleSynthese.Cells.FormatConditions.Delete

    With feuilleSynthese.Range("A1")
        .Value = "Reporting Risque Quotidien - " & nomPortefeuille
        .Font.Size = 16
        .Font.Bold = True
    End With

    feuilleSynthese.Range("A2").Value = "Date de valorisation :"
    feuilleSynthese.Range("B2").Value = dateValorisation
    feuilleSynthese.Range("B2").NumberFormat = "dd/mm/yyyy"

    ' En-tetes du tableau d'indicateurs (ligne 4)
    Dim entetes As Variant
    entetes = Array("Indicateur", "Valeur", "Limite", "Utilisation", "Statut")

    Dim i As Long
    For i = LBound(entetes) To UBound(entetes)
        With feuilleSynthese.Cells(4, i + 1)
            .Value = entetes(i)
            .Font.Bold = True
            .Interior.Color = RGB(217, 226, 243)
            .Borders.LineStyle = xlContinuous
        End With
    Next i

    feuilleSynthese.Columns("A:E").AutoFit
    feuilleSynthese.Range("A1").Select

    Exit Sub

GestionErreur:
    Err.Raise Err.Number, "modReporting.ConstruireFeuilleSynthese", Err.Description
End Sub

'===============================================================================
' AjouterLigneIndicateur
' Ajoute une ligne d'indicateur dans le tableau de synthese (a partir de la
' ligne 5) et applique la mise en forme conditionnelle "feux tricolores" sur
' la cellule de statut.
'   feuilleSynthese : Worksheet cible
'   ligne           : numero de ligne (>= 5)
'   libelle, valeur, limite, utilisation, statut : contenu de la ligne
'===============================================================================
Public Sub AjouterLigneIndicateur(ByVal feuilleSynthese As Worksheet, ByVal ligne As Long, ByVal libelle As String, _
    ByVal valeur As Double, ByVal limite As Double, ByVal utilisation As Double, ByVal statut As String)

    On Error GoTo GestionErreur

    feuilleSynthese.Cells(ligne, 1).Value = libelle
    feuilleSynthese.Cells(ligne, 2).Value = valeur
    feuilleSynthese.Cells(ligne, 3).Value = limite
    feuilleSynthese.Cells(ligne, 4).Value = utilisation
    feuilleSynthese.Cells(ligne, 4).NumberFormat = "0.0%"
    feuilleSynthese.Cells(ligne, 5).Value = statut

    Dim cellStatut As Range
    Set cellStatut = feuilleSynthese.Cells(ligne, 5)

    Select Case UCase$(statut)
        Case "OK"
            cellStatut.Interior.Color = RGB(198, 239, 206)
            cellStatut.Font.Color = RGB(0, 97, 0)
        Case "ALERTE"
            cellStatut.Interior.Color = RGB(255, 235, 156)
            cellStatut.Font.Color = RGB(156, 87, 0)
        Case "DEPASSEMENT"
            cellStatut.Interior.Color = RGB(255, 199, 206)
            cellStatut.Font.Color = RGB(156, 0, 6)
    End Select
    cellStatut.Font.Bold = True

    feuilleSynthese.Range(feuilleSynthese.Cells(ligne, 1), feuilleSynthese.Cells(ligne, 5)).Borders.LineStyle = xlContinuous
    Exit Sub

GestionErreur:
    Err.Raise Err.Number, "modReporting.AjouterLigneIndicateur", Err.Description
End Sub

'===============================================================================
' AppliquerFeuxTricolores
' Applique une mise en forme conditionnelle native Excel (3 regles) sur une
' plage d'utilisation de limite (valeurs en %) : vert si < seuilAlerte, orange
' si entre seuilAlerte et 100%, rouge si >= 100%. Alternative "par formule"
' aux couleurs fixees en dur par AjouterLigneIndicateur, utile si les valeurs
' de la plage peuvent changer dynamiquement (recalcul de formules).
'   plage       : plage de cellules contenant des utilisations (0 a >1)
'   seuilAlerte : seuil d'alerte, ex 0,9
'===============================================================================
Public Sub AppliquerFeuxTricolores(ByVal plage As Range, Optional ByVal seuilAlerte As Double = 0.9)
    On Error GoTo GestionErreur
    Dim fc As FormatCondition

    plage.FormatConditions.Delete

    ' Regle 1 : rouge si >= 100% (depassement)
    Set fc = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="1")
    fc.Interior.Color = RGB(255, 199, 206)
    fc.Font.Color = RGB(156, 0, 6)

    ' Regle 2 : orange si >= seuilAlerte et < 100%
    Set fc = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:=CStr(seuilAlerte), Formula2:="1")
    fc.Interior.Color = RGB(255, 235, 156)
    fc.Font.Color = RGB(156, 87, 0)

    ' Regle 3 : vert si < seuilAlerte
    Set fc = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=CStr(seuilAlerte))
    fc.Interior.Color = RGB(198, 239, 206)
    fc.Font.Color = RGB(0, 97, 0)

    Exit Sub

GestionErreur:
    Err.Raise Err.Number, "modReporting.AppliquerFeuxTricolores", Err.Description
End Sub

'===============================================================================
' CreerGraphiqueVL
' Cree (ou remplace) un graphique en courbe de l'evolution de la valeur
' liquidative (VL) du fonds vs benchmark, via ChartObjects.
'   feuilleCible : Worksheet ou inserer le graphique
'   plageDates   : Range des dates (axe X)
'   plageVLFonds : Range des VL du fonds (serie 1)
'   plageVLBench : Range des VL du benchmark (serie 2, optionnelle)
'   nomGraphique : nom du ChartObject (pour pouvoir le retrouver/remplacer)
'===============================================================================
Public Sub CreerGraphiqueVL(ByVal feuilleCible As Worksheet, ByVal plageDates As Range, ByVal plageVLFonds As Range, _
    Optional ByVal plageVLBench As Range = Nothing, Optional ByVal nomGraphique As String = "GraphiqueVL")

    On Error GoTo GestionErreur
    Dim graphique As ChartObject
    Dim serie As Series

    SupprimerGraphiqueSiExiste feuilleCible, nomGraphique

    Set graphique = feuilleCible.ChartObjects.Add(Left:=feuilleCible.Range("H2").Left, Top:=feuilleCible.Range("H2").Top, Width:=480, Height:=280)
    graphique.Name = nomGraphique

    With graphique.Chart
        .ChartType = xlLine
        .SetSourceData Source:=plageVLFonds
        .SeriesCollection(1).Name = "Fonds"
        .SeriesCollection(1).XValues = plageDates

        If Not (plageVLBench Is Nothing) Then
            Set serie = .SeriesCollection.NewSeries
            serie.Values = plageVLBench
            serie.XValues = plageDates
            serie.Name = "Benchmark"
        End If

        .HasTitle = True
        .ChartTitle.Text = "Evolution de la valeur liquidative"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "VL"
    End With

    Exit Sub

GestionErreur:
    Err.Raise Err.Number, "modReporting.CreerGraphiqueVL", Err.Description
End Sub

'===============================================================================
' CreerGraphiqueContributionRisque
' Cree un graphique en barres de contribution au risque par ligne / secteur.
'   feuilleCible  : Worksheet ou inserer le graphique
'   plageLibelles : Range des libelles (axe des categories)
'   plageValeurs  : Range des contributions au risque (axe des valeurs)
'   nomGraphique  : nom du ChartObject
'===============================================================================
Public Sub CreerGraphiqueContributionRisque(ByVal feuilleCible As Worksheet, ByVal plageLibelles As Range, ByVal plageValeurs As Range, _
    Optional ByVal nomGraphique As String = "GraphiqueContribRisque")

    On Error GoTo GestionErreur
    Dim graphique As ChartObject

    SupprimerGraphiqueSiExiste feuilleCible, nomGraphique

    Set graphique = feuilleCible.ChartObjects.Add(Left:=feuilleCible.Range("H20").Left, Top:=feuilleCible.Range("H20").Top, Width:=480, Height:=280)
    graphique.Name = nomGraphique

    With graphique.Chart
        .ChartType = xlBarClustered
        .SetSourceData Source:=plageValeurs
        .SeriesCollection(1).Name = "Contribution au risque"
        .SeriesCollection(1).XValues = plageLibelles
        .HasTitle = True
        .ChartTitle.Text = "Contribution au risque par ligne"
    End With

    Exit Sub

GestionErreur:
    Err.Raise Err.Number, "modReporting.CreerGraphiqueContributionRisque", Err.Description
End Sub

Private Sub SupprimerGraphiqueSiExiste(ByVal feuilleCible As Worksheet, ByVal nomGraphique As String)
    On Error Resume Next
    feuilleCible.ChartObjects(nomGraphique).Delete
    On Error GoTo 0
End Sub

'===============================================================================
' ExporterEnPDF
' Exporte une feuille (ou tout le classeur) en PDF via ExportAsFixedFormat.
'   classeurCible : Workbook a exporter
'   cheminSortie  : chemin complet du fichier PDF de destination
'   feuilleUnique : Worksheet a exporter seule (Nothing = tout le classeur)
' Renvoie True si l'export a reussi.
'===============================================================================
Public Function ExporterEnPDF(ByVal classeurCible As Workbook, ByVal cheminSortie As String, Optional ByVal feuilleUnique As Worksheet = Nothing) As Boolean
    On Error GoTo GestionErreur

    If Not (feuilleUnique Is Nothing) Then
        feuilleUnique.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminSortie, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Else
        classeurCible.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminSortie, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End If

    ExporterEnPDF = True
    Exit Function

GestionErreur:
    ExporterEnPDF = False
End Function

'===============================================================================
' EnvoyerBrouillonOutlook
' Prepare un email de reporting via Outlook et l'affiche a l'utilisateur pour
' verification et envoi MANUEL. N'appelle JAMAIS MailItem.Send : la routine
' s'arrete systematiquement a .Display, conformement a la politique de
' securite du classeur (aucun envoi automatique sans validation humaine).
'   destinataires : liste d'adresses separees par ";"
'   objet         : objet de l'email
'   corpsHTML     : corps du message au format HTML
'   cheminPieceJointe : chemin d'un fichier a joindre (ex : le PDF exporte),
'                       optionnel ("" si aucune piece jointe)
' Renvoie True si le brouillon a ete cree et affiche avec succes.
'===============================================================================
Public Function EnvoyerBrouillonOutlook(ByVal destinataires As String, ByVal objet As String, ByVal corpsHTML As String, Optional ByVal cheminPieceJointe As String = "") As Boolean
    On Error GoTo GestionErreur
    Dim applicationOutlook As Object
    Dim messageMail As Object

    ' Late binding (CreateObject) : evite d'exiger une reference au type
    ' library Outlook si le classeur est ouvert sur un poste sans Outlook.
    Set applicationOutlook = CreateObject("Outlook.Application")
    Set messageMail = applicationOutlook.CreateItem(0) ' 0 = olMailItem

    With messageMail
        .To = destinataires
        .Subject = objet
        .HTMLBody = corpsHTML

        If Len(Trim$(cheminPieceJointe)) > 0 Then
            If Dir(cheminPieceJointe) <> "" Then
                .Attachments.Add cheminPieceJointe
            End If
        End If

        ' IMPORTANT : .Display ouvre le brouillon pour relecture humaine.
        ' .Send N'EST JAMAIS APPELE ICI - envoi manuel obligatoire.
        .Display
    End With

    EnvoyerBrouillonOutlook = True
    Exit Function

GestionErreur:
    ' Cas frequent : Outlook non installe sur le poste - on echoue proprement
    ' sans planter le reste du reporting.
    EnvoyerBrouillonOutlook = False
End Function

'===============================================================================
' GenererCorpsEmailReporting
' Construit un corps d'email HTML simple resumant le reporting (nombre de
' depassements, date, lien vers le fichier). Fonction pure, facilement
' testable, separee de EnvoyerBrouillonOutlook.
'===============================================================================
Public Function GenererCorpsEmailReporting(ByVal nomPortefeuille As String, ByVal dateValorisation As Date, ByVal nDepassements As Long) As String
    Dim corps As String

    corps = "<html><body>"
    corps = corps & "<p>Bonjour,</p>"
    corps = corps & "<p>Veuillez trouver ci-joint le reporting risque quotidien du <b>" & Format$(dateValorisation, "dd/mm/yyyy") & "</b> pour le portefeuille <b>" & nomPortefeuille & "</b>.</p>"

    If nDepassements > 0 Then
        corps = corps & "<p style='color:red;'><b>Attention : " & nDepassements & " depassement(s) de limite detecte(s).</b> Merci de vous referer a la feuille Limites du classeur joint.</p>"
    Else
        corps = corps & "<p style='color:green;'>Aucun depassement de limite detecte.</p>"
    End If

    corps = corps & "<p>Cordialement,<br/>Equipe Risque</p>"
    corps = corps & "</body></html>"

    GenererCorpsEmailReporting = corps
End Function
