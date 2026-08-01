Attribute VB_Name = "modRiskMetrics"
'===============================================================================
' Module    : modRiskMetrics
' Objet     : Fonctions Excel definies par l'utilisateur (UDF) pour le calcul
'             des indicateurs de risque de marche standard utilises en
'             reporting risque middle-office (VaR, ES, volatilite, TE, beta...).
' Auteur    : Gianni Pilotti
' Portee    : Toutes les fonctions sont appelables directement depuis une
'             feuille de calcul, par exemple :
'                 =VaRHistorique(B2:B251;0,99)
' Conventions :
'   - Les plages d'entree contiennent des RENDEMENTS PERIODIQUES (pas des prix),
'     sauf mention contraire explicite dans l'en-tete de la fonction.
'   - En cas d'entree invalide, chaque fonction renvoie une erreur Excel propre
'     via CVErr(xlErrValue) ou CVErr(xlErrNum) plutot qu'un plantage VBA.
'   - Application.Volatile n'est utilise que lorsque le resultat depend d'un
'     etat non capture par les arguments (aucun cas ici : toutes les fonctions
'     sont pures par rapport a leurs arguments, donc AUCUNE n'est marquee
'     volatile - ceci est documente explicitement pour chaque fonction).
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' RangeToArray
' Convertit une Range Excel en tableau Double a une dimension (base 1).
' Ignore silencieusement les cellules vides et non numeriques si bIgnoreBlanks=True.
'-------------------------------------------------------------------------------
Private Function RangeToArray(ByVal plage As Range, Optional ByVal bIgnoreBlanks As Boolean = True) As Double()
    Dim cell As Range
    Dim resultat() As Double
    Dim n As Long
    Dim i As Long

    n = plage.Cells.Count
    ReDim resultat(1 To n)
    i = 0

    For Each cell In plage.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            i = i + 1
            resultat(i) = CDbl(cell.Value)
        ElseIf Not bIgnoreBlanks Then
            i = i + 1
            resultat(i) = 0
        End If
    Next cell

    If i = 0 Then
        ReDim resultat(1 To 0)
    Else
        ReDim Preserve resultat(1 To i)
    End If

    RangeToArray = resultat
End Function

'-------------------------------------------------------------------------------
' QuickSortDouble
' Tri rapide (in-place) d'un tableau Double, utilise par les fonctions de VaR
' historique / ES qui necessitent les rendements tries par ordre croissant.
'-------------------------------------------------------------------------------
Private Sub QuickSortDouble(ByRef arr() As Double, ByVal gauche As Long, ByVal droite As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, temp As Double

    If gauche >= droite Then Exit Sub

    i = gauche
    j = droite
    pivot = arr((gauche + droite) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If gauche < j Then QuickSortDouble arr, gauche, j
    If i < droite Then QuickSortDouble arr, i, droite
End Sub

'-------------------------------------------------------------------------------
' Moyenne / EcartType : petites fonctions statistiques internes, evitent la
' dependance a Application.WorksheetFunction pour rester robustes et testables
' depuis modTests.bas.
'-------------------------------------------------------------------------------
Private Function Moyenne(ByRef arr() As Double) As Double
    Dim i As Long, s As Double
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    If n <= 0 Then
        Moyenne = 0
        Exit Function
    End If
    s = 0
    For i = LBound(arr) To UBound(arr)
        s = s + arr(i)
    Next i
    Moyenne = s / n
End Function

Private Function EcartType(ByRef arr() As Double, Optional ByVal bEchantillon As Boolean = True) As Double
    Dim i As Long, s As Double, m As Double
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    If n <= 1 Then
        EcartType = 0
        Exit Function
    End If
    m = Moyenne(arr)
    s = 0
    For i = LBound(arr) To UBound(arr)
        s = s + (arr(i) - m) ^ 2
    Next i
    If bEchantillon Then
        EcartType = Sqr(s / (n - 1))
    Else
        EcartType = Sqr(s / n)
    End If
End Function

'-------------------------------------------------------------------------------
' PercentileLineaire
' Percentile par interpolation lineaire (methode "inclusive", coherente avec
' PERCENTILE.INC d'Excel), sur un tableau NON trie en entree (il est trie ici).
'-------------------------------------------------------------------------------
Private Function PercentileLineaire(ByRef arr() As Double, ByVal p As Double) As Double
    Dim n As Long
    Dim triee() As Double
    Dim rang As Double
    Dim rangInf As Long, rangSup As Long
    Dim frac As Double

    n = UBound(arr) - LBound(arr) + 1
    triee = arr
    QuickSortDouble triee, LBound(triee), UBound(triee)

    If n = 1 Then
        PercentileLineaire = triee(LBound(triee))
        Exit Function
    End If

    rang = p * (n - 1) ' index 0-based
    rangInf = Int(rang)
    rangSup = rangInf + 1
    frac = rang - rangInf

    If rangSup > n - 1 Then
        PercentileLineaire = triee(LBound(triee) + n - 1)
    Else
        PercentileLineaire = triee(LBound(triee) + rangInf) + frac * (triee(LBound(triee) + rangSup) - triee(LBound(triee) + rangInf))
    End If
End Function

'-------------------------------------------------------------------------------
' NiveauValide : validation commune du parametre "niveau de confiance".
'-------------------------------------------------------------------------------
Private Function NiveauValide(ByVal niveau As Double) As Boolean
    NiveauValide = (niveau > 0 And niveau < 1)
End Function

'===============================================================================
' VaRHistorique
' VaR historique (non parametrique) exprimee en PERTE POSITIVE (ex : 0,023 =
' perte potentielle de 2,3% du notionnel au niveau de confiance donne).
'   plage  : Range de rendements periodiques (ex : rendements quotidiens)
'   niveau : niveau de confiance, ex 0,99 pour VaR 99%
' Exemple  : =VaRHistorique(B2:B251;0,99)
'===============================================================================
Public Function VaRHistorique(ByVal plage As Range, ByVal niveau As Double) As Variant
    On Error GoTo GestionErreur
    Dim rendements() As Double
    Dim quantile As Double

    If Not NiveauValide(niveau) Then
        VaRHistorique = CVErr(xlErrValue)
        Exit Function
    End If

    rendements = RangeToArray(plage)
    If UBound(rendements) < LBound(rendements) Then
        VaRHistorique = CVErr(xlErrValue)
        Exit Function
    End If

    ' Le quantile de perte au niveau (1 - niveau) de la distribution des
    ' rendements correspond a la VaR (exprimee en perte positive).
    quantile = PercentileLineaire(rendements, 1 - niveau)
    VaRHistorique = Application.WorksheetFunction.Max(0, -quantile)
    Exit Function

GestionErreur:
    VaRHistorique = CVErr(xlErrValue)
End Function

'===============================================================================
' VaRParametrique
' VaR parametrique (gaussienne), avec mise a l'echelle temporelle par racine du
' temps (hypothese de rendements i.i.d.).
'   plage   : Range de rendements periodiques
'   niveau  : niveau de confiance (ex : 0,99)
'   horizon : horizon en nombre de periodes (ex : 10 jours), doit etre >= 1
' Exemple   : =VaRParametrique(B2:B251;0,99;10)
'===============================================================================
Public Function VaRParametrique(ByVal plage As Range, ByVal niveau As Double, Optional ByVal horizon As Double = 1) As Variant
    On Error GoTo GestionErreur
    Dim rendements() As Double
    Dim mu As Double, sigma As Double
    Dim zScore As Double

    If Not NiveauValide(niveau) Or horizon < 1 Then
        VaRParametrique = CVErr(xlErrValue)
        Exit Function
    End If

    rendements = RangeToArray(plage)
    If UBound(rendements) - LBound(rendements) + 1 < 2 Then
        VaRParametrique = CVErr(xlErrValue)
        Exit Function
    End If

    mu = Moyenne(rendements)
    sigma = EcartType(rendements, True)

    zScore = NormSInvPrecise(niveau)

    VaRParametrique = Application.WorksheetFunction.Max(0, -(mu * horizon - zScore * sigma * Sqr(horizon)))
    Exit Function

GestionErreur:
    VaRParametrique = CVErr(xlErrValue)
End Function

'===============================================================================
' VaRCornishFisher
' VaR parametrique corrigee de l'asymetrie (skewness) et de l'aplatissement
' (excess kurtosis) via le developpement de Cornish-Fisher, plus fidele que la
' VaR gaussienne lorsque la distribution des rendements a des queues epaisses.
'   plage  : Range de rendements periodiques
'   niveau : niveau de confiance (ex : 0,99)
' Exemple  : =VaRCornishFisher(B2:B251;0,99)
'===============================================================================
Public Function VaRCornishFisher(ByVal plage As Range, ByVal niveau As Double) As Variant
    On Error GoTo GestionErreur
    Dim rendements() As Double
    Dim mu As Double, sigma As Double
    Dim skew As Double, kurt As Double
    Dim z As Double, zCF As Double
    Dim n As Long, i As Long
    Dim s3 As Double, s4 As Double

    If Not NiveauValide(niveau) Then
        VaRCornishFisher = CVErr(xlErrValue)
        Exit Function
    End If

    rendements = RangeToArray(plage)
    n = UBound(rendements) - LBound(rendements) + 1
    If n < 4 Then
        VaRCornishFisher = CVErr(xlErrValue)
        Exit Function
    End If

    mu = Moyenne(rendements)
    sigma = EcartType(rendements, True)
    If sigma = 0 Then
        VaRCornishFisher = 0
        Exit Function
    End If

    s3 = 0: s4 = 0
    For i = LBound(rendements) To UBound(rendements)
        s3 = s3 + ((rendements(i) - mu) / sigma) ^ 3
        s4 = s4 + ((rendements(i) - mu) / sigma) ^ 4
    Next i
    skew = s3 / n
    kurt = (s4 / n) - 3 ' excess kurtosis

    z = NormSInvPrecise(niveau)

    ' Expansion de Cornish-Fisher au 2e ordre (skew + kurtosis)
    zCF = z + (z ^ 2 - 1) * skew / 6 _
            + (z ^ 3 - 3 * z) * kurt / 24 _
            - (2 * z ^ 3 - 5 * z) * (skew ^ 2) / 36

    VaRCornishFisher = Application.WorksheetFunction.Max(0, -(mu - zCF * sigma))
    Exit Function

GestionErreur:
    VaRCornishFisher = CVErr(xlErrValue)
End Function

'===============================================================================
' ExpectedShortfall
' Expected Shortfall (CVaR) historique = moyenne des pertes au-dela du seuil de
' VaR historique, exprimee en perte positive.
'   plage  : Range de rendements periodiques
'   niveau : niveau de confiance (ex : 0,975)
' Exemple  : =ExpectedShortfall(B2:B251;0,975)
'===============================================================================
Public Function ExpectedShortfall(ByVal plage As Range, ByVal niveau As Double) As Variant
    On Error GoTo GestionErreur
    Dim rendements() As Double
    Dim triee() As Double
    Dim n As Long, nQueue As Long, i As Long
    Dim somme As Double

    If Not NiveauValide(niveau) Then
        ExpectedShortfall = CVErr(xlErrValue)
        Exit Function
    End If

    rendements = RangeToArray(plage)
    n = UBound(rendements) - LBound(rendements) + 1
    If n < 1 Then
        ExpectedShortfall = CVErr(xlErrValue)
        Exit Function
    End If

    triee = rendements
    QuickSortDouble triee, LBound(triee), UBound(triee)

    ' Nombre d'observations dans la queue de perte (1 - niveau), au moins 1
    nQueue = Application.WorksheetFunction.RoundUp((1 - niveau) * n, 0)
    If nQueue < 1 Then nQueue = 1
    If nQueue > n Then nQueue = n

    somme = 0
    For i = LBound(triee) To LBound(triee) + nQueue - 1
        somme = somme + triee(i)
    Next i

    ExpectedShortfall = Application.WorksheetFunction.Max(0, -(somme / nQueue))
    Exit Function

GestionErreur:
    ExpectedShortfall = CVErr(xlErrValue)
End Function

'===============================================================================
' VolatiliteAnnualisee
' Volatilite annualisee = ecart-type des rendements periodiques * racine(freq).
'   plage     : Range de rendements periodiques
'   frequence : nombre de periodes par an (252 quotidien, 52 hebdo, 12 mensuel)
' Exemple     : =VolatiliteAnnualisee(B2:B251;252)
'===============================================================================
Public Function VolatiliteAnnualisee(ByVal plage As Range, Optional ByVal frequence As Double = 252) As Variant
    On Error GoTo GestionErreur
    Dim rendements() As Double

    If frequence <= 0 Then
        VolatiliteAnnualisee = CVErr(xlErrValue)
        Exit Function
    End If

    rendements = RangeToArray(plage)
    If UBound(rendements) - LBound(rendements) + 1 < 2 Then
        VolatiliteAnnualisee = CVErr(xlErrValue)
        Exit Function
    End If

    VolatiliteAnnualisee = EcartType(rendements, True) * Sqr(frequence)
    Exit Function

GestionErreur:
    VolatiliteAnnualisee = CVErr(xlErrValue)
End Function

'===============================================================================
' VolatiliteEWMA
' Volatilite (non annualisee) par moyenne mobile ponderee exponentiellement
' (methode RiskMetrics). Le poids le plus recent a le plus de poids.
'   plage  : Range de rendements periodiques, ordonnee du plus ANCIEN au plus
'            RECENT (convention standard d'une serie chronologique en colonne)
'   lambda : facteur de decroissance, 0 < lambda < 1 (ex : 0,94 RiskMetrics)
' Exemple  : =VolatiliteEWMA(B2:B251;0,94)
'===============================================================================
Public Function VolatiliteEWMA(ByVal plage As Range, Optional ByVal lambda As Double = 0.94) As Variant
    On Error GoTo GestionErreur
    Dim rendements() As Double
    Dim n As Long, i As Long
    Dim varEWMA As Double
    Dim poidsCumules As Double
    Dim poids As Double

    If lambda <= 0 Or lambda >= 1 Then
        VolatiliteEWMA = CVErr(xlErrValue)
        Exit Function
    End If

    rendements = RangeToArray(plage)
    n = UBound(rendements) - LBound(rendements) + 1
    If n < 2 Then
        VolatiliteEWMA = CVErr(xlErrValue)
        Exit Function
    End If

    ' rendements(UBound) = observation la plus recente (fin de la plage)
    varEWMA = 0
    poidsCumules = 0
    For i = 0 To n - 1
        ' i = 0 correspond a l'observation la plus recente
        poids = (1 - lambda) * (lambda ^ i)
        varEWMA = varEWMA + poids * (rendements(UBound(rendements) - i) ^ 2)
        poidsCumules = poidsCumules + poids
    Next i

    If poidsCumules > 0 Then varEWMA = varEWMA / poidsCumules

    VolatiliteEWMA = Sqr(varEWMA)
    Exit Function

GestionErreur:
    VolatiliteEWMA = CVErr(xlErrValue)
End Function

'===============================================================================
' TrackingError
' Tracking error = ecart-type annualise des rendements actifs (fonds - bench).
'   plageFonds : Range de rendements periodiques du portefeuille
'   plageBench : Range de rendements periodiques du benchmark (meme taille)
'   frequence  : periodes par an pour l'annualisation (defaut 252)
' Exemple      : =TrackingError(B2:B251;C2:C251;252)
'===============================================================================
Public Function TrackingError(ByVal plageFonds As Range, ByVal plageBench As Range, Optional ByVal frequence As Double = 252) As Variant
    On Error GoTo GestionErreur
    Dim fonds() As Double, bench() As Double
    Dim actifs() As Double
    Dim n As Long, i As Long

    If frequence <= 0 Then
        TrackingError = CVErr(xlErrValue)
        Exit Function
    End If

    fonds = RangeToArray(plageFonds)
    bench = RangeToArray(plageBench)

    n = UBound(fonds) - LBound(fonds) + 1
    If n <> (UBound(bench) - LBound(bench) + 1) Or n < 2 Then
        TrackingError = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim actifs(1 To n)
    For i = 1 To n
        actifs(i) = fonds(LBound(fonds) + i - 1) - bench(LBound(bench) + i - 1)
    Next i

    TrackingError = EcartType(actifs, True) * Sqr(frequence)
    Exit Function

GestionErreur:
    TrackingError = CVErr(xlErrValue)
End Function

'===============================================================================
' BetaPortefeuille
' Beta = Covariance(fonds, bench) / Variance(bench).
'   plageFonds, plageBench : Ranges de rendements periodiques, meme taille
' Exemple : =BetaPortefeuille(B2:B251;C2:C251)
'===============================================================================
Public Function BetaPortefeuille(ByVal plageFonds As Range, ByVal plageBench As Range) As Variant
    On Error GoTo GestionErreur
    Dim fonds() As Double, bench() As Double
    Dim n As Long, i As Long
    Dim mFonds As Double, mBench As Double
    Dim cov As Double, varBench As Double

    fonds = RangeToArray(plageFonds)
    bench = RangeToArray(plageBench)

    n = UBound(fonds) - LBound(fonds) + 1
    If n <> (UBound(bench) - LBound(bench) + 1) Or n < 2 Then
        BetaPortefeuille = CVErr(xlErrValue)
        Exit Function
    End If

    mFonds = Moyenne(fonds)
    mBench = Moyenne(bench)

    cov = 0
    varBench = 0
    For i = 1 To n
        cov = cov + (fonds(LBound(fonds) + i - 1) - mFonds) * (bench(LBound(bench) + i - 1) - mBench)
        varBench = varBench + (bench(LBound(bench) + i - 1) - mBench) ^ 2
    Next i

    If varBench = 0 Then
        BetaPortefeuille = CVErr(xlErrDiv0)
        Exit Function
    End If

    BetaPortefeuille = cov / varBench
    Exit Function

GestionErreur:
    BetaPortefeuille = CVErr(xlErrValue)
End Function

'===============================================================================
' RatioSharpe
' Ratio de Sharpe annualise = (rendement moyen annualise - taux sans risque) /
'                              volatilite annualisee.
'   plage        : Range de rendements periodiques du portefeuille
'   tauxSansRisque : taux sans risque ANNUEL (ex : 0,03 pour 3%)
'   frequence    : periodes par an (defaut 252)
' Exemple : =RatioSharpe(B2:B251;0,03;252)
'===============================================================================
Public Function RatioSharpe(ByVal plage As Range, Optional ByVal tauxSansRisque As Double = 0, Optional ByVal frequence As Double = 252) As Variant
    On Error GoTo GestionErreur
    Dim rendements() As Double
    Dim mu As Double, sigma As Double
    Dim rendementAnnualise As Double, volAnnualisee As Double

    If frequence <= 0 Then
        RatioSharpe = CVErr(xlErrValue)
        Exit Function
    End If

    rendements = RangeToArray(plage)
    If UBound(rendements) - LBound(rendements) + 1 < 2 Then
        RatioSharpe = CVErr(xlErrValue)
        Exit Function
    End If

    mu = Moyenne(rendements)
    sigma = EcartType(rendements, True)

    rendementAnnualise = mu * frequence
    volAnnualisee = sigma * Sqr(frequence)

    If volAnnualisee = 0 Then
        RatioSharpe = CVErr(xlErrDiv0)
        Exit Function
    End If

    RatioSharpe = (rendementAnnualise - tauxSansRisque) / volAnnualisee
    Exit Function

GestionErreur:
    RatioSharpe = CVErr(xlErrValue)
End Function

'===============================================================================
' RatioInformation
' Ratio d'information = rendement actif annualise moyen / tracking error.
'   plageFonds, plageBench : Ranges de rendements periodiques, meme taille
'   frequence : periodes par an (defaut 252)
' Exemple : =RatioInformation(B2:B251;C2:C251;252)
'===============================================================================
Public Function RatioInformation(ByVal plageFonds As Range, ByVal plageBench As Range, Optional ByVal frequence As Double = 252) As Variant
    On Error GoTo GestionErreur
    Dim fonds() As Double, bench() As Double
    Dim actifs() As Double
    Dim n As Long, i As Long
    Dim rendementActifAnnualise As Double, teAnnualisee As Double

    If frequence <= 0 Then
        RatioInformation = CVErr(xlErrValue)
        Exit Function
    End If

    fonds = RangeToArray(plageFonds)
    bench = RangeToArray(plageBench)

    n = UBound(fonds) - LBound(fonds) + 1
    If n <> (UBound(bench) - LBound(bench) + 1) Or n < 2 Then
        RatioInformation = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim actifs(1 To n)
    For i = 1 To n
        actifs(i) = fonds(LBound(fonds) + i - 1) - bench(LBound(bench) + i - 1)
    Next i

    rendementActifAnnualise = Moyenne(actifs) * frequence
    teAnnualisee = EcartType(actifs, True) * Sqr(frequence)

    If teAnnualisee = 0 Then
        RatioInformation = CVErr(xlErrDiv0)
        Exit Function
    End If

    RatioInformation = rendementActifAnnualise / teAnnualisee
    Exit Function

GestionErreur:
    RatioInformation = CVErr(xlErrValue)
End Function

'===============================================================================
' MaxDrawdown
' Max Drawdown = plus forte baisse pic-a-creux d'une serie de VALEURS
' LIQUIDATIVES (niveaux, pas rendements), exprime en valeur NEGATIVE ou nulle
' (ex : -0,18 = baisse maximale de 18% depuis un plus haut).
'   plage : Range de niveaux de VL (prix), ordonnee chronologiquement
' Exemple : =MaxDrawdown(B2:B251)
'===============================================================================
Public Function MaxDrawdown(ByVal plage As Range) As Variant
    On Error GoTo GestionErreur
    Dim niveaux() As Double
    Dim i As Long
    Dim plusHaut As Double
    Dim drawdown As Double, pireDrawdown As Double

    niveaux = RangeToArray(plage)
    If UBound(niveaux) - LBound(niveaux) + 1 < 2 Then
        MaxDrawdown = CVErr(xlErrValue)
        Exit Function
    End If

    plusHaut = niveaux(LBound(niveaux))
    pireDrawdown = 0

    For i = LBound(niveaux) To UBound(niveaux)
        If niveaux(i) > plusHaut Then plusHaut = niveaux(i)
        If plusHaut > 0 Then
            drawdown = (niveaux(i) - plusHaut) / plusHaut
            If drawdown < pireDrawdown Then pireDrawdown = drawdown
        End If
    Next i

    MaxDrawdown = pireDrawdown
    Exit Function

GestionErreur:
    MaxDrawdown = CVErr(xlErrValue)
End Function

'===============================================================================
' NormSDistPrecise
' Fonction de repartition de la loi normale centree reduite N(0,1), calculee
' par une approximation de haute precision (Zelen & Severo, erreur < 7.5e-8),
' independante de Application.WorksheetFunction.NormSDist pour rester portable.
'   x : quantile
' Exemple : =NormSDistPrecise(1,96)  -> environ 0,975
'===============================================================================
Public Function NormSDistPrecise(ByVal x As Double) As Double
    Dim t As Double, y As Double
    Dim b1 As Double, b2 As Double, b3 As Double, b4 As Double, b5 As Double, p As Double, c As Double
    Dim absX As Double

    p = 0.2316419
    b1 = 0.319381530
    b2 = -0.356563782
    b3 = 1.781477937
    b4 = -1.821255978
    b5 = 1.330274429
    c = 0.39894228 ' 1 / sqrt(2*pi)

    absX = Abs(x)
    t = 1 / (1 + p * absX)
    y = 1 - c * Exp(-absX * absX / 2) * (b1 * t + b2 * t ^ 2 + b3 * t ^ 3 + b4 * t ^ 4 + b5 * t ^ 5)

    If x >= 0 Then
        NormSDistPrecise = y
    Else
        NormSDistPrecise = 1 - y
    End If
End Function

'===============================================================================
' NormSInvPrecise
' Inverse de la fonction de repartition de la loi normale centree reduite,
' approximation rationnelle de Peter Acklam (erreur relative < 1.15e-9).
' Utilisee par les fonctions de VaR parametrique / Cornish-Fisher et par
' modOptions.VolImplicite.
'   p : probabilite, 0 < p < 1
' Exemple : =NormSInvPrecise(0,99) -> environ 2,3263
'===============================================================================
Public Function NormSInvPrecise(ByVal p As Double) As Double
    ' Coefficients de l'approximation d'Acklam
    Dim a1 As Double, a2 As Double, a3 As Double, a4 As Double, a5 As Double, a6 As Double
    Dim b1 As Double, b2 As Double, b3 As Double, b4 As Double, b5 As Double
    Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double, c5 As Double, c6 As Double
    Dim d1 As Double, d2 As Double, d3 As Double, d4 As Double
    Dim pLow As Double, pHigh As Double
    Dim q As Double, r As Double
    Dim resultat As Double

    If p <= 0 Or p >= 1 Then
        ' Hors domaine : on renvoie une valeur sentinelle plutot que de planter,
        ' les fonctions appelantes valident deja p en amont.
        NormSInvPrecise = 0
        Exit Function
    End If

    a1 = -3.969683028665376E+01: a2 = 2.209460984245205E+02: a3 = -2.759285104469687E+02
    a4 = 1.38357751867269E+02: a5 = -3.066479806614716E+01: a6 = 2.506628277459239

    b1 = -5.447609879822406E+01: b2 = 1.615858368580409E+02: b3 = -1.556989798598866E+02
    b4 = 6.680131188771972E+01: b5 = -1.328068155288572E+01

    c1 = -7.784894002430293E-03: c2 = -3.223964580411365E-01: c3 = -2.400758277161838
    c4 = -2.549732539343734: c5 = 4.374664141464968: c6 = 2.938163982698783

    d1 = 7.784695709041462E-03: d2 = 3.224671290700398E-01: d3 = 2.445134137142996
    d4 = 3.754408661907416

    pLow = 0.02425
    pHigh = 1 - pLow

    If p < pLow Then
        q = Sqr(-2 * Log(p))
        resultat = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                   ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    ElseIf p <= pHigh Then
        q = p - 0.5
        r = q * q
        resultat = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
                   (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
    Else
        q = Sqr(-2 * Log(1 - p))
        resultat = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                    ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    End If

    ' Un pas de raffinement de Halley ameliore encore la precision.
    Dim e As Double, u As Double
    e = 0.5 * WorksheetFunctionErfc(-resultat / Sqr(2)) - p
    u = e * Sqr(2 * 3.14159265358979) * Exp(resultat * resultat / 2)
    resultat = resultat - u / (1 + resultat * u / 2)

    NormSInvPrecise = resultat
End Function

'-------------------------------------------------------------------------------
' WorksheetFunctionErfc
' Fonction d'erreur complementaire erfc(x), utilisee uniquement pour le pas de
' raffinement de Halley dans NormSInvPrecise. Approximation d'Abramowitz &
' Stegun 7.1.26 (erreur max ~1.5e-7), sans dependance externe.
'-------------------------------------------------------------------------------
Private Function WorksheetFunctionErfc(ByVal x As Double) As Double
    Dim t As Double, y As Double
    Dim signeX As Integer
    Dim absX As Double

    If x >= 0 Then
        signeX = 1
    Else
        signeX = -1
    End If
    absX = Abs(x)

    t = 1 / (1 + 0.3275911 * absX)
    y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Exp(-absX * absX)

    If signeX = 1 Then
        WorksheetFunctionErfc = 1 - y
    Else
        WorksheetFunctionErfc = 1 + y
    End If
End Function
