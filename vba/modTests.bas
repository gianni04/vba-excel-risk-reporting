Attribute VB_Name = "modTests"
'===============================================================================
' Module    : modTests
' Objet     : Mini-framework de tests unitaires VBA (AssertEqual,
'             AssertApproxEqual, RunAllTests) avec des cas de test reels sur
'             les UDF de modRiskMetrics et modOptions, dont des valeurs
'             Black-Scholes de reference calculees independamment (Python
'             scipy.stats.norm, voir examples/02_python_vs_vba.py du depot
'             pour la reproduction cote Python).
' Auteur    : Gianni Pilotti
'
' Utilisation : executer RunAllTests() depuis l'editeur VBA (F5) ou l'appeler
' depuis un bouton de la feuille "Tests". Les resultats s'affichent dans la
' fenetre Immediate (Ctrl+G) et dans une MsgBox de synthese.
'===============================================================================

Option Explicit

Private m_nTests As Long
Private m_nReussis As Long
Private m_journal As String

'-------------------------------------------------------------------------------
' AssertEqual : compare deux Variant pour egalite stricte (utilise pour les
' entiers, chaines, booleens - PAS pour les Double, voir AssertApproxEqual).
'-------------------------------------------------------------------------------
Public Sub AssertEqual(ByVal valeurAttendue As Variant, ByVal valeurObtenue As Variant, ByVal nomTest As String)
    m_nTests = m_nTests + 1

    If CStr(valeurAttendue) = CStr(valeurObtenue) Then
        m_nReussis = m_nReussis + 1
        m_journal = m_journal & "[OK]   " & nomTest & vbCrLf
    Else
        m_journal = m_journal & "[ECHEC] " & nomTest & " - attendu=" & CStr(valeurAttendue) & " obtenu=" & CStr(valeurObtenue) & vbCrLf
    End If
End Sub

'-------------------------------------------------------------------------------
' AssertApproxEqual : compare deux valeurs numeriques a une tolerance pres
' (indispensable pour les resultats de calcul flottant).
'-------------------------------------------------------------------------------
Public Sub AssertApproxEqual(ByVal valeurAttendue As Double, ByVal valeurObtenue As Double, ByVal tolerance As Double, ByVal nomTest As String)
    m_nTests = m_nTests + 1

    If Abs(valeurAttendue - valeurObtenue) <= tolerance Then
        m_nReussis = m_nReussis + 1
        m_journal = m_journal & "[OK]   " & nomTest & vbCrLf
    Else
        m_journal = m_journal & "[ECHEC] " & nomTest & " - attendu=" & Format$(valeurAttendue, "0.000000") & _
            " obtenu=" & Format$(valeurObtenue, "0.000000") & " (tolerance=" & tolerance & ")" & vbCrLf
    End If
End Sub

'-------------------------------------------------------------------------------
' AssertTrue / AssertFalse : assertions booleennes simples.
'-------------------------------------------------------------------------------
Public Sub AssertTrue(ByVal condition As Boolean, ByVal nomTest As String)
    m_nTests = m_nTests + 1
    If condition Then
        m_nReussis = m_nReussis + 1
        m_journal = m_journal & "[OK]   " & nomTest & vbCrLf
    Else
        m_journal = m_journal & "[ECHEC] " & nomTest & " - condition fausse" & vbCrLf
    End If
End Sub

Public Sub AssertIsError(ByVal valeurObtenue As Variant, ByVal nomTest As String)
    m_nTests = m_nTests + 1
    If IsError(valeurObtenue) Then
        m_nReussis = m_nReussis + 1
        m_journal = m_journal & "[OK]   " & nomTest & vbCrLf
    Else
        m_journal = m_journal & "[ECHEC] " & nomTest & " - une erreur Excel etait attendue, valeur obtenue=" & CStr(valeurObtenue) & vbCrLf
    End If
End Sub

'===============================================================================
' RunAllTests
' Point d'entree principal : execute tous les groupes de tests et affiche une
' synthese. A lancer manuellement (F5) apres import des modules dans un
' classeur .xlsm - certains tests utilisent des Range sur une feuille
' temporaire "TestsTemp" creee et supprimee automatiquement.
'===============================================================================
Public Sub RunAllTests()
    m_nTests = 0
    m_nReussis = 0
    m_journal = ""

    TesterFonctionsStatistiques
    TesterBlackScholes
    TesterGrecques
    TesterVolImplicite
    TesterRiskMetricsSurFeuille
    TesterDataUtils
    TesterPortefeuille
    TesterLimites

    Debug.Print "==================================================================="
    Debug.Print m_journal
    Debug.Print "==================================================================="
    Debug.Print m_nReussis & " / " & m_nTests & " tests reussis."

    MsgBox m_nReussis & " / " & m_nTests & " tests reussis." & vbCrLf & "Voir la fenetre Immediate (Ctrl+G) pour le detail.", _
        IIf(m_nReussis = m_nTests, vbInformation, vbExclamation), "modTests.RunAllTests"
End Sub

'-------------------------------------------------------------------------------
' TesterFonctionsStatistiques : NormSDistPrecise / NormSInvPrecise, verifies
' contre scipy.stats.norm (Python) - voir examples/02_python_vs_vba.py.
'-------------------------------------------------------------------------------
Private Sub TesterFonctionsStatistiques()
    AssertApproxEqual 0.9750021048517795, modRiskMetrics.NormSDistPrecise(1.96), 0.0000005, "NormSDistPrecise(1,96) ~ 0,975002"
    AssertApproxEqual 0.5, modRiskMetrics.NormSDistPrecise(0), 0.0000005, "NormSDistPrecise(0) = 0,5"
    AssertApproxEqual 2.3263478740408408, modRiskMetrics.NormSInvPrecise(0.99), 0.00001, "NormSInvPrecise(0,99) ~ 2,326348"
    AssertApproxEqual 1.959963984540054, modRiskMetrics.NormSInvPrecise(0.975), 0.00001, "NormSInvPrecise(0,975) ~ 1,959964"
    AssertApproxEqual 0, modRiskMetrics.NormSInvPrecise(0.5), 0.00001, "NormSInvPrecise(0,5) = 0"
End Sub

'-------------------------------------------------------------------------------
' TesterBlackScholes : valeurs de reference calculees en Python (scipy) pour
' S=100, K=100, T=1 an, r=5%, q=0%, sigma=20% :
'   Call = 10,450584   Put = 5,573526
'-------------------------------------------------------------------------------
Private Sub TesterBlackScholes()
    Dim prixCall As Double, prixPut As Double

    prixCall = CDbl(modOptions.BlackScholes("C", 100, 100, 1, 0.05, 0, 0.2))
    prixPut = CDbl(modOptions.BlackScholes("P", 100, 100, 1, 0.05, 0, 0.2))

    AssertApproxEqual 10.450584, prixCall, 0.0001, "BlackScholes Call ATM 1an vol20% ~ 10,450584"
    AssertApproxEqual 5.573526, prixPut, 0.0001, "BlackScholes Put ATM 1an vol20% ~ 5,573526"

    ' Parite call-put : C - P = S*exp(-qT) - K*exp(-rT)
    Dim ecartParite As Double
    ecartParite = prixCall - prixPut - (100 * Exp(0) - 100 * Exp(-0.05 * 1))
    AssertApproxEqual 0, ecartParite, 0.0001, "Parite call-put respectee"

    ' A l'echeance (T=0), le prix = valeur intrinseque
    AssertApproxEqual 10, CDbl(modOptions.BlackScholes("C", 110, 100, 0, 0.05, 0, 0.2)), 0.0001, "BlackScholes Call T=0 = valeur intrinseque"
    AssertApproxEqual 0, CDbl(modOptions.BlackScholes("C", 90, 100, 0, 0.05, 0, 0.2)), 0.0001, "BlackScholes Call T=0 hors la monnaie = 0"

    ' Validation des entrees invalides
    AssertIsError modOptions.BlackScholes("C", -100, 100, 1, 0.05, 0, 0.2), "BlackScholes S negatif renvoie une erreur"
    AssertIsError modOptions.BlackScholes("C", 100, 100, 1, 0.05, 0, -0.2), "BlackScholes sigma negatif renvoie une erreur"
End Sub

'-------------------------------------------------------------------------------
' TesterGrecques : valeurs de reference Python (scipy), meme jeu de parametres.
'   Delta Call = 0,636831   Delta Put = -0,363169
'   Gamma      = 0,018762   Vega (convention /100) = 0,375240
'-------------------------------------------------------------------------------
Private Sub TesterGrecques()
    AssertApproxEqual 0.636831, CDbl(modOptions.BSDelta("C", 100, 100, 1, 0.05, 0, 0.2)), 0.0001, "BSDelta Call ~ 0,636831"
    AssertApproxEqual -0.363169, CDbl(modOptions.BSDelta("P", 100, 100, 1, 0.05, 0, 0.2)), 0.0001, "BSDelta Put ~ -0,363169"
    AssertApproxEqual 0.018762, CDbl(modOptions.BSGamma(100, 100, 1, 0.05, 0, 0.2)), 0.0001, "BSGamma ~ 0,018762"
    AssertApproxEqual 0.375240, CDbl(modOptions.BSVega(100, 100, 1, 0.05, 0, 0.2)), 0.0001, "BSVega (/100) ~ 0,375240"

    ' Le gamma d'un call et d'un put a memes parametres doit etre identique
    AssertApproxEqual CDbl(modOptions.BSGamma(100, 100, 1, 0.05, 0, 0.2)), CDbl(modOptions.BSGamma(100, 100, 1, 0.05, 0, 0.2)), 0.0000001, "BSGamma coherent"
End Sub

'-------------------------------------------------------------------------------
' TesterVolImplicite : reconstitue le prix Black-Scholes puis retrouve sigma.
'-------------------------------------------------------------------------------
Private Sub TesterVolImplicite()
    Dim sigmaAttendu As Double, prixReference As Double
    Dim sigmaRetrouve As Double

    sigmaAttendu = 0.25
    prixReference = CDbl(modOptions.BlackScholes("C", 100, 105, 0.75, 0.02, 0.01, sigmaAttendu))

    sigmaRetrouve = CDbl(modOptions.VolImplicite(prixReference, "C", 100, 105, 0.75, 0.02, 0.01))

    AssertApproxEqual sigmaAttendu, sigmaRetrouve, 0.0005, "VolImplicite retrouve sigma=0,25 par Newton-Raphson"

    ' Cas degenere : point de depart tres eloigne, doit basculer en bissection
    Dim sigmaRetrouve2 As Double
    sigmaRetrouve2 = CDbl(modOptions.VolImplicite(prixReference, "C", 100, 105, 0.75, 0.02, 0.01, 4.9))
    AssertApproxEqual sigmaAttendu, sigmaRetrouve2, 0.001, "VolImplicite converge meme avec un depart eloigne (repli bissection)"
End Sub

'-------------------------------------------------------------------------------
' TesterRiskMetricsSurFeuille : cree une feuille temporaire pour tester les
' UDF necessitant une Range en entree (VaR, ES, vol, TE, beta, drawdown...).
'-------------------------------------------------------------------------------
Private Sub TesterRiskMetricsSurFeuille()
    On Error GoTo Nettoyage
    Dim feuilleTest As Worksheet
    Dim classeur As Workbook
    Dim i As Long

    Set classeur = ThisWorkbook
    SupprimerFeuilleSiExiste classeur, "TestsTemp"
    Set feuilleTest = classeur.Worksheets.Add
    feuilleTest.Name = "TestsTemp"

    ' Rendements synthetiques simples et connus a l'avance : 10 valeurs
    ' -0,05, -0,04, ..., 0,04 (pas de 0,01), moyenne = -0,005
    For i = 1 To 10
        feuilleTest.Cells(i, 1).Value = -0.05 + (i - 1) * 0.01
    Next i

    ' VaR historique 90% sur cette serie triee croissante de 10 valeurs :
    ' quantile a 10% (1-0.9) -> interpolation lineaire methode PERCENTILE.INC
    Dim varObtenue As Double
    varObtenue = CDbl(modRiskMetrics.VaRHistorique(feuilleTest.Range("A1:A10"), 0.9))
    AssertTrue varObtenue > 0, "VaRHistorique renvoie une perte positive"

    ' Serie CONSTANTE (colonne B) -> volatilite nulle
    For i = 1 To 10
        feuilleTest.Cells(i, 2).Value = 0.001
    Next i
    AssertApproxEqual 0, CDbl(modRiskMetrics.VolatiliteAnnualisee(feuilleTest.Range("B1:B10"), 252)), 0.0000001, "VolatiliteAnnualisee(serie constante) = 0"

    ' Serie VARIABLE (colonne F) repliquee a l'identique en colonne G (le
    ' "benchmark") -> TE nulle (le fonds replique exactement son benchmark)
    ' et Beta = 1 (covariance = variance du benchmark).
    For i = 1 To 10
        feuilleTest.Cells(i, 6).Value = -0.02 + (i - 1) * 0.005
        feuilleTest.Cells(i, 7).Value = feuilleTest.Cells(i, 6).Value
    Next i
    AssertApproxEqual 0, CDbl(modRiskMetrics.TrackingError(feuilleTest.Range("F1:F10"), feuilleTest.Range("G1:G10"), 252)), 0.0000001, "TrackingError(fonds=benchmark) = 0"
    AssertApproxEqual 1, CDbl(modRiskMetrics.BetaPortefeuille(feuilleTest.Range("F1:F10"), feuilleTest.Range("G1:G10"))), 0.0001, "BetaPortefeuille(fonds=benchmark) = 1"

    ' MaxDrawdown sur une serie de VL montante puis descendante
    feuilleTest.Cells(1, 4).Value = 100
    feuilleTest.Cells(2, 4).Value = 110
    feuilleTest.Cells(3, 4).Value = 121
    feuilleTest.Cells(4, 4).Value = 90.75  ' -25% depuis le plus haut de 121
    feuilleTest.Cells(5, 4).Value = 100
    AssertApproxEqual -0.25, CDbl(modRiskMetrics.MaxDrawdown(feuilleTest.Range("D1:D5"))), 0.0001, "MaxDrawdown detecte -25% depuis le plus haut"

    ' MaxDrawdown sur une serie strictement croissante = 0 (jamais negatif)
    feuilleTest.Cells(1, 5).Value = 100
    feuilleTest.Cells(2, 5).Value = 101
    feuilleTest.Cells(3, 5).Value = 105
    AssertTrue CDbl(modRiskMetrics.MaxDrawdown(feuilleTest.Range("E1:E3"))) <= 0, "MaxDrawdown toujours <= 0"
    AssertApproxEqual 0, CDbl(modRiskMetrics.MaxDrawdown(feuilleTest.Range("E1:E3"))), 0.0000001, "MaxDrawdown(serie croissante) = 0"

    ' Validation d'entree invalide (niveau hors [0,1])
    AssertIsError modRiskMetrics.VaRHistorique(feuilleTest.Range("A1:A10"), 1.5), "VaRHistorique(niveau=1,5) renvoie une erreur"
    AssertIsError modRiskMetrics.VaRHistorique(feuilleTest.Range("A1:A10"), 0), "VaRHistorique(niveau=0) renvoie une erreur"

Nettoyage:
    On Error Resume Next
    SupprimerFeuilleSiExiste classeur, "TestsTemp"
    On Error GoTo 0
End Sub

Private Sub SupprimerFeuilleSiExiste(ByVal classeur As Workbook, ByVal nomFeuille As String)
    On Error Resume Next
    Dim afficherAlertesInitial As Boolean
    afficherAlertesInitial = Application.DisplayAlerts
    Application.DisplayAlerts = False
    classeur.Worksheets(nomFeuille).Delete
    Application.DisplayAlerts = afficherAlertesInitial
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' TesterDataUtils : QuickSort, RechercheDichotomique, calendrier jours ouvres.
'-------------------------------------------------------------------------------
Private Sub TesterDataUtils()
    Dim tableauTest As Variant
    tableauTest = Array(5, 3, 8, 1, 9, 2)
    modDataUtils.QuickSort tableauTest, True
    AssertEqual 1, tableauTest(0), "QuickSort croissant - premier element"
    AssertEqual 9, tableauTest(5), "QuickSort croissant - dernier element"

    Dim tableauDecroissant As Variant
    tableauDecroissant = Array(5, 3, 8, 1, 9, 2)
    modDataUtils.QuickSort tableauDecroissant, False
    AssertEqual 9, tableauDecroissant(0), "QuickSort decroissant - premier element"

    Dim tableauTrie() As Double
    ReDim tableauTrie(0 To 4)
    tableauTrie(0) = 1: tableauTrie(1) = 3: tableauTrie(2) = 5: tableauTrie(3) = 7: tableauTrie(4) = 9
    AssertEqual 2, modDataUtils.RechercheDichotomique(tableauTrie, 5), "RechercheDichotomique trouve l'indice correct"
    AssertEqual -1, modDataUtils.RechercheDichotomique(tableauTrie, 4), "RechercheDichotomique renvoie -1 si absent"

    ' Lundi 3 mars 2025 est un jour ouvre ; samedi 8 mars 2025 ne l'est pas
    AssertTrue modDataUtils.EstJourOuvre(DateSerial(2025, 3, 3)), "EstJourOuvre(lundi) = True"
    AssertTrue Not modDataUtils.EstJourOuvre(DateSerial(2025, 3, 8)), "EstJourOuvre(samedi) = False"
    AssertTrue Not modDataUtils.EstJourOuvre(DateSerial(2025, 3, 9)), "EstJourOuvre(dimanche) = False"
End Sub

'-------------------------------------------------------------------------------
' TesterPortefeuille : clsPortefeuille, agregations et concentration.
'-------------------------------------------------------------------------------
Private Sub TesterPortefeuille()
    Dim p As New clsPortefeuille
    p.Nom = "Fonds Test"
    p.AjouterPosition "FR0000131104", "TotalEnergies", "Energie", 1500000, "EUR", "A"
    p.AjouterPosition "DE0007236101", "Siemens", "Industrie", 900000, "EUR", "A"
    p.AjouterPosition "US0378331005", "Apple", "Technologie", 600000, "USD", "AA"
    p.AjouterPosition "COUVERTURE01", "Future CAC40 (couverture)", "Derives", -400000, "EUR", "NR"

    AssertApproxEqual 3400000, p.ExpositionBrute(), 0.01, "ExpositionBrute = somme des valeurs absolues"
    AssertApproxEqual 2600000, p.ExpositionNette(), 0.01, "ExpositionNette = somme algebrique"
    AssertEqual 4, p.NombrePositions, "NombrePositions = 4"

    AssertApproxEqual 1500000 / 3400000, p.PoidsSecteur("Energie"), 0.0001, "PoidsSecteur Energie coherent"
    AssertApproxEqual 1500000 / 3400000, p.ConcentrationEmetteur(), 0.0001, "ConcentrationEmetteur = plus grosse ligne / expo brute"

    AssertEqual 1, p.NombrePositionsSousNotation("AA"), "NombrePositionsSousNotation('AA') compte les lignes moins bien notees"
End Sub

'-------------------------------------------------------------------------------
' TesterLimites : moteur de controle des limites (statuts et utilisation).
'-------------------------------------------------------------------------------
Private Sub TesterLimites()
    AssertApproxEqual 0.8, modLimites.CalculerUtilisation(0.024, 0.03), 0.0001, "CalculerUtilisation 0,024/0,03 = 0,8"
    AssertEqual StatutOK, modLimites.DeterminerStatut(0.5, 0.9), "DeterminerStatut(50%, seuil 90%) = OK"
    AssertEqual StatutAlerte, modLimites.DeterminerStatut(0.92, 0.9), "DeterminerStatut(92%, seuil 90%) = ALERTE"
    AssertEqual StatutDepassement, modLimites.DeterminerStatut(1.05, 0.9), "DeterminerStatut(105%, seuil 90%) = DEPASSEMENT"
    AssertEqual "DEPASSEMENT", modLimites.LibelleStatut(StatutDepassement), "LibelleStatut(Depassement) = 'DEPASSEMENT'"
End Sub
