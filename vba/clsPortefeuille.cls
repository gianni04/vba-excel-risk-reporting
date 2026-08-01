VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPortefeuille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Classe     : clsPortefeuille
' Objet      : Encapsulation orientee objet d'un portefeuille : collection de
'              positions, proprietes Let/Get, methodes d'agregation
'              (exposition brute/nette, poids, exposition sectorielle,
'              concentration emetteur). Demontre la maitrise de la POO en VBA
'              (classe + collection d'objets clsPosition simules par un
'              tableau de structures internes, VBA n'ayant pas de Type dans
'              une collection native facilement manipulable ici).
' Auteur     : Gianni Pilotti
'
' Utilisation typique :
'   Dim p As New clsPortefeuille
'   p.Nom = "Fonds Actions Europe"
'   p.AjouterPosition "FR0000131104", "TotalEnergies", "Energie", 1500000, "EUR", "A"
'   p.AjouterPosition "DE0007236101", "Siemens", "Industrie", 900000, "EUR", "A"
'   Debug.Print p.ExpositionBrute
'   Debug.Print p.PoidsPosition("FR0000131104")
'===============================================================================

Option Explicit

' --- Variables membres privees (encapsulation) ---
Private m_nom As String
Private m_devise As String
Private m_dateValorisation As Date

' Tableaux paralleles representant les positions (evite la dependance a une
' classe clsPosition separee pour garder le module autonome et lisible).
Private m_isin() As String
Private m_libelle() As String
Private m_secteur() As String
Private m_exposition() As Double
Private m_devisePosition() As String
Private m_notation() As String
Private m_nbPositions As Long
Private Const CAPACITE_INITIALE As Long = 10
Private m_capacite As Long

'-------------------------------------------------------------------------------
' Class_Initialize : initialisation de la collection interne de positions.
'-------------------------------------------------------------------------------
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

' --- Proprietes Let/Get : encapsulation classique VBA ---

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

'===============================================================================
' AjouterPosition
' Ajoute une position au portefeuille. Redimensionne automatiquement la
' capacite interne si necessaire (doublement, amorti O(1)).
'   isin        : code ISIN de la ligne
'   libelle     : nom de l'emetteur / instrument
'   secteur     : secteur GICS/ICB simplifie
'   exposition  : exposition en devise du portefeuille (peut etre negative
'                 pour une position courte / instrument de couverture)
'   devisePosition : devise native de la position
'   notation    : notation credit simplifiee (ex : "AAA", "A", "BBB", "NR")
'===============================================================================
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

'===============================================================================
' ViderPositions : reinitialise le portefeuille (supprime toutes les positions).
'===============================================================================
Public Sub ViderPositions()
    Class_Initialize
End Sub

'===============================================================================
' ExpositionBrute : somme des valeurs absolues des expositions (gross exposure).
'===============================================================================
Public Function ExpositionBrute() As Double
    Dim i As Long, s As Double
    s = 0
    For i = 1 To m_nbPositions
        s = s + Abs(m_exposition(i))
    Next i
    ExpositionBrute = s
End Function

'===============================================================================
' ExpositionNette : somme algebrique des expositions (net exposure).
'===============================================================================
Public Function ExpositionNette() As Double
    Dim i As Long, s As Double
    s = 0
    For i = 1 To m_nbPositions
        s = s + m_exposition(i)
    Next i
    ExpositionNette = s
End Function

'===============================================================================
' ValeurTotale : alias de ExpositionNette, utilise comme base de calcul des
' poids (peut differer si l'on souhaite exclure la tresorerie a l'avenir).
'===============================================================================
Public Function ValeurTotale() As Double
    ValeurTotale = ExpositionNette()
End Function

'===============================================================================
' PoidsPosition : poids d'une position (identifiee par ISIN) rapporte a la
' valeur totale nette du portefeuille. Renvoie 0 si l'ISIN est introuvable ou
' si la valeur totale est nulle (division par zero evitee).
'===============================================================================
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

'===============================================================================
' ExpositionParSecteur : exposition brute agregee pour un secteur donne.
'===============================================================================
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

'===============================================================================
' PoidsSecteur : poids d'un secteur rapporte a l'exposition brute totale.
'===============================================================================
Public Function PoidsSecteur(ByVal secteur As String) As Double
    Dim brut As Double
    brut = ExpositionBrute()
    If brut = 0 Then
        PoidsSecteur = 0
        Exit Function
    End If
    PoidsSecteur = ExpositionParSecteur(secteur) / brut
End Function

'===============================================================================
' ConcentrationEmetteur : plus grande exposition brute individuelle rapportee
' a l'exposition brute totale (indicateur de concentration mono-emetteur).
'===============================================================================
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

'===============================================================================
' NombrePositionsSousNotation : compte les positions dont la notation est
' strictement inferieure au seuil (echelle simplifiee : AAA > AA > A > BBB >
' BB > B > NR, NR traite comme le plus bas).
'   notationMinimale : ex "BBB"
'===============================================================================
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

'-------------------------------------------------------------------------------
' RangNotation : convertit une notation credit simplifiee en rang numerique
' (plus petit = meilleure qualite de credit).
'-------------------------------------------------------------------------------
Private Function RangNotation(ByVal notation As String) As Long
    Select Case UCase$(Trim$(notation))
        Case "AAA": RangNotation = 1
        Case "AA": RangNotation = 2
        Case "A": RangNotation = 3
        Case "BBB": RangNotation = 4
        Case "BB": RangNotation = 5
        Case "B": RangNotation = 6
        Case "CCC": RangNotation = 7
        Case Else: RangNotation = 8 ' NR ou inconnu : traite comme le plus risque
    End Select
End Function

'===============================================================================
' ListeISIN : renvoie un tableau des ISIN presents dans le portefeuille (utile
' pour parcourir les positions depuis l'exterieur de la classe).
'===============================================================================
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

'===============================================================================
' ExpositionPosition : exposition d'une position donnee (0 si ISIN introuvable).
'===============================================================================
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
