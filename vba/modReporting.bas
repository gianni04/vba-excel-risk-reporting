Attribute VB_Name = "modReporting"

Option Explicit

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

Public Sub AppliquerFeuxTricolores(ByVal plage As Range, Optional ByVal seuilAlerte As Double = 0.9)
    On Error GoTo GestionErreur
    Dim fc As FormatCondition

    plage.FormatConditions.Delete

    Set fc = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="1")
    fc.Interior.Color = RGB(255, 199, 206)
    fc.Font.Color = RGB(156, 0, 6)

    Set fc = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:=CStr(seuilAlerte), Formula2:="1")
    fc.Interior.Color = RGB(255, 235, 156)
    fc.Font.Color = RGB(156, 87, 0)

    Set fc = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=CStr(seuilAlerte))
    fc.Interior.Color = RGB(198, 239, 206)
    fc.Font.Color = RGB(0, 97, 0)

    Exit Sub

GestionErreur:
    Err.Raise Err.Number, "modReporting.AppliquerFeuxTricolores", Err.Description
End Sub

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

Public Function EnvoyerBrouillonOutlook(ByVal destinataires As String, ByVal objet As String, ByVal corpsHTML As String, Optional ByVal cheminPieceJointe As String = "") As Boolean
    On Error GoTo GestionErreur
    Dim applicationOutlook As Object
    Dim messageMail As Object

    Set applicationOutlook = CreateObject("Outlook.Application")
    Set messageMail = applicationOutlook.CreateItem(0)

    With messageMail
        .To = destinataires
        .Subject = objet
        .HTMLBody = corpsHTML

        If Len(Trim$(cheminPieceJointe)) > 0 Then
            If Dir(cheminPieceJointe) <> "" Then
                .Attachments.Add cheminPieceJointe
            End If
        End If

        .Display
    End With

    EnvoyerBrouillonOutlook = True
    Exit Function

GestionErreur:
    EnvoyerBrouillonOutlook = False
End Function

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
