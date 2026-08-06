Attribute VB_Name = "modOptions"

Option Explicit

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

Private Function ValiderParametresBS(ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal sigma As Double) As Boolean
    ValiderParametresBS = (S > 0 And K > 0 And T >= 0 And sigma > 0)
End Function

Private Function D1BS(ByVal S As Double, ByVal K As Double, ByVal T As Double, ByVal r As Double, ByVal q As Double, ByVal sigma As Double) As Double
    D1BS = (Log(S / K) + (r - q + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
End Function

Private Function D2BS(ByVal d1 As Double, ByVal sigma As Double, ByVal T As Double) As Double
    D2BS = d1 - sigma * Sqr(T)
End Function

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

    BSVega = S * Exp(-q * T) * phiD1 * Sqr(T) / 100
    Exit Function

GestionErreur:
    BSVega = CVErr(xlErrValue)
End Function

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

Private Function DensiteNormale(ByVal x As Double) As Double
    DensiteNormale = (1 / Sqr(2 * 3.14159265358979)) * Exp(-0.5 * x ^ 2)
End Function

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

    sigma = sigmaInitial
    converge = False

    For i = 1 To maxIterations
        prixCalcule = CDbl(BlackScholes(typeOption, S, K, T, r, q, sigma))
        vega = CDbl(BSVega(S, K, T, r, q, sigma)) * 100

        If Abs(prixCalcule - prixMarche) < tolerance Then
            converge = True
            Exit For
        End If

        If vega < 0.0000001 Then
            Exit For
        End If

        sigma = sigma - (prixCalcule - prixMarche) / vega

        If sigma <= 0 Or sigma > 5 Then
            Exit For
        End If
    Next i

    If converge And sigma > 0 And sigma <= 5 Then
        VolImplicite = sigma
        Exit Function
    End If

    Dim bInf As Double, bSup As Double, bMid As Double
    Dim prixInf As Double, prixMid As Double
    Dim j As Long

    bInf = 0.0001
    bSup = 5

    prixInf = CDbl(BlackScholes(typeOption, S, K, T, r, q, bInf))
    Dim prixSup As Double
    prixSup = CDbl(BlackScholes(typeOption, S, K, T, r, q, bSup))

    If (prixMarche < prixInf) Or (prixMarche > prixSup) Then
        VolImplicite = CVErr(xlErrNum)
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
