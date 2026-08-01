Attribute VB_Name = "modOptions"
'===============================================================================
' Module    : modOptions
' Objet     : UDF de pricing d'options vanille (Black-Scholes-Merton) et de
'             calcul des sensibilites (Grecques), plus extraction de la
'             volatilite implicite par Newton-Raphson avec repli bissection.
' Auteur    : Gianni Pilotti
' Remarque  : NormSDistPrecise et NormSInvPrecise (module modRiskMetrics) sont
'             utilisees ici au lieu de Application.WorksheetFunction.NormSDist/
'             NormSInv afin de ne pas dependre du moteur de calcul Excel pour
'             ces fonctions statistiques - utile aussi pour les tests unitaires
'             (modTests.bas) qui peuvent tourner sans feuille active.
' Convention: type = "C" ou "CALL" pour un call, "P" ou "PUT" pour un put
'             (insensible a la casse).
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' EstCall : normalise le parametre "type" en booleen, leve une erreur VBA
' (interceptee par l'appelant) si la chaine n'est pas reconnue.
'-------------------------------------------------------------------------------
Private Function EstCall(ByVal typeOption As String) As Boolean
    Dim t As String
    t = UCase$(Trim$(typeOption))
    Select Case t
        Case "C", "CALL"
            EstCall = True
        Case "P", "PUT"
            EstCall = False
        Case Else
            Err.Raise vbObjectError + 513, "modOptions.EstCall", "Type d'option invalide : " & typeOption
    End Select
End Function

'-------------------------------------------------------------------------------
' ValiderParametresBS : validation commune des parametres de Black-Scholes.
' Renvoie True si tous les parametres sont dans un domaine economiquement
' valide (S, K, sigma > 0, T >= 0).
'-------------------------------------------------------------------------------
Private Function ValiderParametresBS(ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal sigma As Double) As Boolean
    ValiderParametresBS = (S > 0 And K > 0 And T >= 0 And sigma > 0)
End Function

'-------------------------------------------------------------------------------
' D1BS / D2BS : termes d1 et d2 du modele de Black-Scholes-Merton avec
' rendement de dividende continu q.
'-------------------------------------------------------------------------------
Private Function D1BS(ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, ByVal sigma As Double) As Double
    D1BS = (Log(S / K) + (r - q + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
End Function

Private Function D2BS(ByVal d1 As Double, ByVal sigma As Double, ByVal T As Double) As Double
    D2BS = d1 - sigma * Sqr(T)
End Function

'===============================================================================
' BlackScholes
' Prix d'une option vanille europeenne par la formule de Black-Scholes-Merton.
'   typeOption : "C" (call) ou "P" (put)
'   S     : cours du sous-jacent
'   K     : prix d'exercice (strike)
'   T     : maturite en annees (ex : 0,5 pour 6 mois)
'   r     : taux sans risque continu annualise (ex : 0,03)
'   q     : rendement de dividende continu annualise (0 si aucun)
'   sigma : volatilite annualisee (ex : 0,20 pour 20%)
' Exemple : =BlackScholes("C";100;100;0,5;0,03;0;0,2)
'===============================================================================
Public Function BlackScholes(ByVal typeOption As String, ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, ByVal sigma As Double) As Variant
    On Error GoTo GestionErreur
    Dim d1 As Double, d2 As Double
    Dim estCallLocal As Boolean

    If Not ValiderParametresBS(S, K, T, sigma) Then
        BlackScholes = CVErr(xlErrValue)
        Exit Function
    End If

    estCallLocal = EstCall(typeOption)

    If T = 0 Then
        ' A l'echeance, le prix est la valeur intrinseque.
        If estCallLocal Then
            BlackScholes = Application.WorksheetFunction.Max(0, S - K)
        Else
            BlackScholes = Application.WorksheetFunction.Max(0, K - S)
        End If
        Exit Function
    End If

    d1 = D1BS(S, K, T, r, q, sigma)
    d2 = D2BS(d1, sigma, T)

    If estCallLocal Then
        BlackScholes = S * Exp(-q * T) * modRiskMetrics.NormSDistPrecise(d1) - K * Exp(-r * T) * modRiskMetrics.NormSDistPrecise(d2)
    Else
        BlackScholes = K * Exp(-r * T) * modRiskMetrics.NormSDistPrecise(-d2) - S * Exp(-q * T) * modRiskMetrics.NormSDistPrecise(-d1)
    End If
    Exit Function

GestionErreur:
    BlackScholes = CVErr(xlErrValue)
End Function

'===============================================================================
' BSDelta
' Delta = sensibilite du prix de l'option a une variation du sous-jacent S.
' Exemple : =BSDelta("C";100;100;0,5;0,03;0;0,2)
'===============================================================================
Public Function BSDelta(ByVal typeOption As String, ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, ByVal sigma As Double) As Variant
    On Error GoTo GestionErreur
    Dim d1 As Double
    Dim estCallLocal As Boolean

    If Not ValiderParametresBS(S, K, T, sigma) Then
        BSDelta = CVErr(xlErrValue)
        Exit Function
    End If

    estCallLocal = EstCall(typeOption)

    If T = 0 Then
        If estCallLocal Then
            BSDelta = IIf(S > K, 1, 0)
        Else
            BSDelta = IIf(S < K, -1, 0)
        End If
        Exit Function
    End If

    d1 = D1BS(S, K, T, r, q, sigma)

    If estCallLocal Then
        BSDelta = Exp(-q * T) * modRiskMetrics.NormSDistPrecise(d1)
    Else
        BSDelta = Exp(-q * T) * (modRiskMetrics.NormSDistPrecise(d1) - 1)
    End If
    Exit Function

GestionErreur:
    BSDelta = CVErr(xlErrValue)
End Function

'===============================================================================
' BSGamma
' Gamma = sensibilite du delta a une variation du sous-jacent (identique pour
' call et put).
' Exemple : =BSGamma(100;100;0,5;0,03;0;0,2)
'===============================================================================
Public Function BSGamma(ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, ByVal sigma As Double) As Variant
    On Error GoTo GestionErreur
    Dim d1 As Double
    Dim phiD1 As Double

    If Not ValiderParametresBS(S, K, T, sigma) Or T = 0 Then
        BSGamma = CVErr(xlErrValue)
        Exit Function
    End If

    d1 = D1BS(S, K, T, r, q, sigma)
    phiD1 = DensiteNormale(d1)

    BSGamma = Exp(-q * T) * phiD1 / (S * sigma * Sqr(T))
    Exit Function

GestionErreur:
    BSGamma = CVErr(xlErrValue)
End Function

'===============================================================================
' BSVega
' Vega = sensibilite du prix a une variation de 1 point de volatilite
' (renvoyee pour 1% de variation de sigma, convention usuelle desk : Vega/100).
' Exemple : =BSVega(100;100;0,5;0,03;0;0,2)
'===============================================================================
Public Function BSVega(ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, ByVal sigma As Double) As Variant
    On Error GoTo GestionErreur
    Dim d1 As Double
    Dim phiD1 As Double

    If Not ValiderParametresBS(S, K, T, sigma) Or T = 0 Then
        BSVega = CVErr(xlErrValue)
        Exit Function
    End If

    d1 = D1BS(S, K, T, r, q, sigma)
    phiD1 = DensiteNormale(d1)

    ' Vega "brut" (pour 1.00 = 100% de variation de sigma), divise par 100
    ' pour obtenir la convention desk (variation de 1 point de vol, ex 20%->21%).
    BSVega = S * Exp(-q * T) * phiD1 * Sqr(T) / 100
    Exit Function

GestionErreur:
    BSVega = CVErr(xlErrValue)
End Function

'===============================================================================
' BSTheta
' Theta = sensibilite du prix au passage d'un jour (theta "par jour calendaire",
' convention desk : Theta annuel / 365).
' Exemple : =BSTheta("C";100;100;0,5;0,03;0;0,2)
'===============================================================================
Public Function BSTheta(ByVal typeOption As String, ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, ByVal sigma As Double) As Variant
    On Error GoTo GestionErreur
    Dim d1 As Double, d2 As Double
    Dim phiD1 As Double
    Dim estCallLocal As Boolean
    Dim thetaAnnuel As Double
    Dim terme1 As Double, terme2 As Double, terme3 As Double

    If Not ValiderParametresBS(S, K, T, sigma) Or T = 0 Then
        BSTheta = CVErr(xlErrValue)
        Exit Function
    End If

    estCallLocal = EstCall(typeOption)
    d1 = D1BS(S, K, T, r, q, sigma)
    d2 = D2BS(d1, sigma, T)
    phiD1 = DensiteNormale(d1)

    terme1 = -(S * Exp(-q * T) * phiD1 * sigma) / (2 * Sqr(T))

    If estCallLocal Then
        terme2 = -r * K * Exp(-r * T) * modRiskMetrics.NormSDistPrecise(d2)
        terme3 = q * S * Exp(-q * T) * modRiskMetrics.NormSDistPrecise(d1)
        thetaAnnuel = terme1 + terme2 + terme3
    Else
        terme2 = r * K * Exp(-r * T) * modRiskMetrics.NormSDistPrecise(-d2)
        terme3 = -q * S * Exp(-q * T) * modRiskMetrics.NormSDistPrecise(-d1)
        thetaAnnuel = terme1 + terme2 + terme3
    End If

    BSTheta = thetaAnnuel / 365
    Exit Function

GestionErreur:
    BSTheta = CVErr(xlErrValue)
End Function

'===============================================================================
' BSRho
' Rho = sensibilite du prix a une variation de 1 point de taux sans risque
' (convention desk : Rho / 100, pour 1% de variation de r).
' Exemple : =BSRho("C";100;100;0,5;0,03;0;0,2)
'===============================================================================
Public Function BSRho(ByVal typeOption As String, ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, ByVal sigma As Double) As Variant
    On Error GoTo GestionErreur
    Dim d2 As Double, d1 As Double
    Dim estCallLocal As Boolean

    If Not ValiderParametresBS(S, K, T, sigma) Or T = 0 Then
        BSRho = CVErr(xlErrValue)
        Exit Function
    End If

    estCallLocal = EstCall(typeOption)
    d1 = D1BS(S, K, T, r, q, sigma)
    d2 = D2BS(d1, sigma, T)

    If estCallLocal Then
        BSRho = (K * T * Exp(-r * T) * modRiskMetrics.NormSDistPrecise(d2)) / 100
    Else
        BSRho = (-K * T * Exp(-r * T) * modRiskMetrics.NormSDistPrecise(-d2)) / 100
    End If
    Exit Function

GestionErreur:
    BSRho = CVErr(xlErrValue)
End Function

'-------------------------------------------------------------------------------
' DensiteNormale : densite de probabilite de la loi normale centree reduite.
'-------------------------------------------------------------------------------
Private Function DensiteNormale(ByVal x As Double) As Double
    DensiteNormale = (1 / Sqr(2 * 3.14159265358979)) * Exp(-0.5 * x ^ 2)
End Function

'===============================================================================
' VolImplicite
' Volatilite implicite retrouvee a partir d'un prix de marche observe, par
' methode de Newton-Raphson (rapide) avec repli automatique sur bissection
' (robuste) si Newton-Raphson diverge ou sort du domaine [0,001 ; 5].
'   prixMarche : prix de marche observe de l'option
'   typeOption : "C" ou "P"
'   S, K, T, r, q : parametres Black-Scholes habituels
'   sigmaInitial   : point de depart de Newton-Raphson (defaut 0,20)
'   tolerance      : tolerance de convergence sur le prix (defaut 1e-6)
'   maxIterations  : nombre maximal d'iterations (defaut 100)
' Exemple : =VolImplicite(10,45;"C";100;100;0,5;0,03;0)
'===============================================================================
Public Function VolImplicite(ByVal prixMarche As Double, ByVal typeOption As String, ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, _
    Optional ByVal sigmaInitial As Double = 0.2, Optional ByVal tolerance As Double = 0.000001, Optional ByVal maxIterations As Long = 100) As Variant

    On Error GoTo GestionErreur
    Dim sigma As Double
    Dim prixCalcule As Double, vega As Double
    Dim i As Long
    Dim converge As Boolean

    If S <= 0 Or K <= 0 Or T <= 0 Or prixMarche <= 0 Then
        VolImplicite = CVErr(xlErrValue)
        Exit Function
    End If

    ' --- Etape 1 : tentative Newton-Raphson ---
    sigma = sigmaInitial
    converge = False

    For i = 1 To maxIterations
        prixCalcule = CDbl(BlackScholes(typeOption, S, K, T, r, q, sigma))
        vega = CDbl(BSVega(S, K, T, r, q, sigma)) * 100 ' vega "brut" (pas la convention /100)

        If Abs(prixCalcule - prixMarche) < tolerance Then
            converge = True
            Exit For
        End If

        If vega < 0.0000001 Then
            Exit For ' vega quasi nul : Newton-Raphson est instable, on bascule en bissection
        End If

        sigma = sigma - (prixCalcule - prixMarche) / vega

        If sigma <= 0 Or sigma > 5 Then
            Exit For ' sortie du domaine plausible : on bascule en bissection
        End If
    Next i

    If converge And sigma > 0 And sigma <= 5 Then
        VolImplicite = sigma
        Exit Function
    End If

    ' --- Etape 2 : repli robuste par bissection sur [0,0001 ; 5] ---
    Dim bInf As Double, bSup As Double, bMid As Double
    Dim prixInf As Double, prixMid As Double
    Dim j As Long

    bInf = 0.0001
    bSup = 5

    prixInf = CDbl(BlackScholes(typeOption, S, K, T, r, q, bInf))
    Dim prixSup As Double
    prixSup = CDbl(BlackScholes(typeOption, S, K, T, r, q, bSup))

    ' Le prix de marche doit se situer entre les prix aux bornes pour que la
    ' bissection soit valide (fonction strictement croissante en sigma).
    If (prixMarche < prixInf) Or (prixMarche > prixSup) Then
        VolImplicite = CVErr(xlErrNum) ' pas de solution plausible (prix incoherent avec les bornes d'arbitrage)
        Exit Function
    End If

    For j = 1 To 200
        bMid = (bInf + bSup) / 2
        prixMid = CDbl(BlackScholes(typeOption, S, K, T, r, q, bMid))

        If Abs(prixMid - prixMarche) < tolerance Then
            VolImplicite = bMid
            Exit Function
        End If

        If prixMid < prixMarche Then
            bInf = bMid
        Else
            bSup = bMid
        End If

        If (bSup - bInf) < 0.0000001 Then Exit For
    Next j

    VolImplicite = (bInf + bSup) / 2
    Exit Function

GestionErreur:
    VolImplicite = CVErr(xlErrValue)
End Function
