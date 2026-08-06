VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPortefeuille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private m_nom As String
Private m_devise As String
Private m_dateValorisation As Date

Private m_isin() As String
Private m_libelle() As String
Private m_secteur() As String
Private m_exposition() As Double
Private m_devisePosition() As String
Private m_notation() As String
Private m_nbPositions As Long
Private Const CAPACITE_INITIALE As Long = 10
Private m_capacite As Long

Private Sub Class_Initialize()
    m_capacite = CAPACITE_INITIALE
    ReDim m_isin(1 To m_capacite)
    ReDim m_libelle(1 To m_capacite)
    ReDim m_secteur(1 To m_capacite)
    ReDim m_exposition(1 To m_capacite)
    ReDim m_devisePosition(1 To m_capacite)
    ReDim m_notation(1 To m_capacite)
    m_nbPositions = 0
    m_devise = "EUR"
    m_dateValorisation = Date
End Sub


Public Property Get Nom() As String
    Nom = m_nom
End Property

Public Property Let Nom(ByVal valeur As String)
    m_nom = valeur
End Property

Public Property Get Devise() As String
    Devise = m_devise
End Property

Public Property Let Devise(ByVal valeur As String)
    If Len(valeur) <> 3 Then
        Err.Raise vbObjectError + 601, "clsPortefeuille.Devise", "Code devise ISO invalide (3 lettres attendues) : " & valeur
    End If
    m_devise = UCase$(valeur)
End Property

Public Property Get DateValorisation() As Date
    DateValorisation = m_dateValorisation
End Property

Public Property Let DateValorisation(ByVal valeur As Date)
    m_dateValorisation = valeur
End Property

Public Property Get NombrePositions() As Long
    NombrePositions = m_nbPositions
End Property

Public Sub AjouterPosition(ByVal isin As String, ByVal libelle As String, ByVal secteur As String, _
    ByVal exposition As Double, Optional ByVal devisePosition As String = "EUR", Optional ByVal notation As String = "NR")

    On Error GoTo GestionErreur

    If Len(Trim$(isin)) = 0 Then
        Err.Raise vbObjectError + 602, "clsPortefeuille.AjouterPosition", "ISIN vide non autorise."
    End If

    If m_nbPositions >= m_capacite Then
        m_capacite = m_capacite * 2
        ReDim Preserve m_isin(1 To m_capacite)
        ReDim Preserve m_libelle(1 To m_capacite)
        ReDim Preserve m_secteur(1 To m_capacite)
        ReDim Preserve m_exposition(1 To m_capacite)
        ReDim Preserve m_devisePosition(1 To m_capacite)
        ReDim Preserve m_notation(1 To m_capacite)
    End If

    m_nbPositions = m_nbPositions + 1
    m_isin(m_nbPositions) = isin
    m_libelle(m_nbPositions) = libelle
    m_secteur(m_nbPositions) = secteur
    m_exposition(m_nbPositions) = exposition
    m_devisePosition(m_nbPositions) = devisePosition
    m_notation(m_nbPositions) = notation
    Exit Sub

GestionErreur:
    Err.Raise Err.Number, "clsPortefeuille.AjouterPosition", Err.Description
End Sub

Public Sub ViderPositions()
    Class_Initialize
End Sub

Public Function ExpositionBrute() As Double
    Dim i As Long, s As Double
    s = 0
    For i = 1 To m_nbPositions
        s = s + Abs(m_exposition(i))
    Next i
    ExpositionBrute = s
End Function

Public Function ExpositionNette() As Double
    Dim i As Long, s As Double
    s = 0
    For i = 1 To m_nbPositions
        s = s + m_exposition(i)
    Next i
    ExpositionNette = s
End Function

Public Function ValeurTotale() As Double
    ValeurTotale = ExpositionNette()
End Function

Public Function PoidsPosition(ByVal isin As String) As Double
    Dim i As Long
    Dim total As Double

    total = ValeurTotale()
    If total = 0 Then
        PoidsPosition = 0
        Exit Function
    End If

    For i = 1 To m_nbPositions
        If StrComp(m_isin(i), isin, vbTextCompare) = 0 Then
            PoidsPosition = m_exposition(i) / total
            Exit Function
        End If
    Next i

    PoidsPosition = 0
End Function

Public Function ExpositionParSecteur(ByVal secteur As String) As Double
    Dim i As Long, s As Double
    s = 0
    For i = 1 To m_nbPositions
        If StrComp(m_secteur(i), secteur, vbTextCompare) = 0 Then
            s = s + Abs(m_exposition(i))
        End If
    Next i
    ExpositionParSecteur = s
End Function

Public Function PoidsSecteur(ByVal secteur As String) As Double
    Dim brut As Double
    brut = ExpositionBrute()
    If brut = 0 Then
        PoidsSecteur = 0
        Exit Function
    End If
    PoidsSecteur = ExpositionParSecteur(secteur) / brut
End Function

Public Function ConcentrationEmetteur() As Double
    Dim i As Long
    Dim maxExpo As Double
    Dim brut As Double

    brut = ExpositionBrute()
    If brut = 0 Then
        ConcentrationEmetteur = 0
        Exit Function
    End If

    maxExpo = 0
    For i = 1 To m_nbPositions
        If Abs(m_exposition(i)) > maxExpo Then maxExpo = Abs(m_exposition(i))
    Next i

    ConcentrationEmetteur = maxExpo / brut
End Function

Public Function NombrePositionsSousNotation(ByVal notationMinimale As String) As Long
    Dim i As Long, n As Long
    Dim rangMin As Long

    rangMin = RangNotation(notationMinimale)
    n = 0
    For i = 1 To m_nbPositions
        If RangNotation(m_notation(i)) > rangMin Then n = n + 1
    Next i

    NombrePositionsSousNotation = n
End Function

Private Function RangNotation(ByVal notation As String) As Long
    Select Case UCase$(Trim$(notation))
        Case "AAA": RangNotation = 1
        Case "AA": RangNotation = 2
        Case "A": RangNotation = 3
        Case "BBB": RangNotation = 4
        Case "BB": RangNotation = 5
        Case "B": RangNotation = 6
        Case "CCC": RangNotation = 7
        Case Else: RangNotation = 8
    End Select
End Function

Public Function ListeISIN() As String()
    Dim resultat() As String
    Dim i As Long

    If m_nbPositions = 0 Then
        ReDim resultat(1 To 0)
        ListeISIN = resultat
        Exit Function
    End If

    ReDim resultat(1 To m_nbPositions)
    For i = 1 To m_nbPositions
        resultat(i) = m_isin(i)
    Next i

    ListeISIN = resultat
End Function

Public Function ExpositionPosition(ByVal isin As String) As Double
    Dim i As Long
    For i = 1 To m_nbPositions
        If StrComp(m_isin(i), isin, vbTextCompare) = 0 Then
            ExpositionPosition = m_exposition(i)
            Exit Function
        End If
    Next i
    ExpositionPosition = 0
End Function
