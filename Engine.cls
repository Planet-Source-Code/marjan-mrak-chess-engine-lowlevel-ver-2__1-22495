VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent FoundMove[(arg1, arg2, ... , argn)]
Public Event FoundMove(ByVal Move As String, ByVal status As String)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent DrawMove[(arg1, arg2, ... , argn)]
Public Event DrawMove(ByVal Move As String)
'local variable(s) to hold property value(s)
Private mvarBlackCanCastleLeft As Boolean 'local copy
Private mvarBlackCanCastleRight As Boolean 'local copy
Private mvarWhiteCanCastleLeft As Boolean 'local copy
Private mvarWhiteCanCastleRight As Boolean 'local copy
Dim pars As Chess.Parser
Dim pars1 As Chess.Parser
Dim kopijapolozajev As String
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent DrawCastle[(arg1, arg2, ... , argn)]
Public Event DrawCastle(ByVal Move As String)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent RestoreBoard[(arg1, arg2, ... , argn)]
Public Event RestoreBoard(ByVal Position As String)
'local variable(s) to hold property value(s)

Private mvarStyleOfPlay As Integer 'local copy
Public Sub Postavi()
tipka = True
End Sub

Public Property Let StyleOfPlay(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StyleOfPlay = 5
    mvarStyleOfPlay = vData
End Property


Public Property Get StyleOfPlay() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StyleOfPlay
    StyleOfPlay = mvarStyleOfPlay
End Property



Public Property Let Levels(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Levels = 5
    mvarLevels = vData
End Property


Public Property Get Levels() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Levels
    Levels = mvarLevels
End Property

Public Property Let WhiteCanCastleRight(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WhiteCanCastleRight = 5
    mvarWhiteCanCastleRight = vData
End Property

Public Property Let WhiteCanCastleLeft(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WhiteCanCastleLeft = 5
    mvarWhiteCanCastleLeft = vData
End Property

Public Property Let BlackCanCastleRight(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlackCanCastleRight = 5
    mvarBlackCanCastleRight = vData
End Property

Public Property Let BlackCanCastleLeft(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlackCanCastleLeft = 5
    mvarBlackCanCastleLeft = vData
End Property

Public Sub think(ByVal Position As String, ByVal numMove As Integer)

DoEvents

Dim mislicrnega As Boolean
Dim narisana As Boolean
Dim moznepoteze As String
Dim moznapoteza As String
Dim najpoteza As New Collection
Dim najpotezasibek As New Collection

Dim n As Integer
Dim kopijapolozajev As String
'!!!!!!!!!!!!!
Dim moccrn As Integer
Dim najmoccrn As Integer
Dim najposcrn As Integer
Dim poscrn As Integer
Dim najdel As Boolean, g As Integer, dodal As Boolean
Dim najmocbel As Integer
Dim mocbel As Integer
Dim zanka As Integer

Set pars = CreateObject("Chess.Parser")
Set pars1 = CreateObject("Chess.Parser")
mislicrnega = True


Ponovi:
moznepoteze = vrniVseMoznePoteze(Position, mislicrnega)
If moznepoteze = "" Then
    MsgBox "Mate!"
    Exit Sub
    End If
For n = 1 To najpoteza.Count
    najpoteza.Remove (1)
    Next

For n = 1 To najpotezasibek.Count
    najpotezasibek.Remove (1)
    Next


najposcrn = izracunajPozicijoNasprotnika(Position, True)
najmoccrn = izracunajNasprotnika(Position, False)

najmocbel = izracunajNasprotnika(Position, True)

'izpraznemo seznam potez
zanka = zanka + 1
For n = 1 To Len(moznepoteze) Step 6 'vsak sedmi znak je zaèetek poteze

'osvežimo položaje
kopijapolozajev = Position

'potezo imamo
dodal = False
moznapoteza = Mid(moznepoteze, n, 5)
izvrsiPotezo moznapoteza, kopijapolozajev



'primerjava pozicijske in kolièinske vrednosti
moccrn = izracunajNasprotnika(kopijapolozajev, False)
poscrn = izracunajPozicijoNasprotnika(kopijapolozajev, True)

mocbel = izracunajNasprotnika(kopijapolozajev, True)


If moccrn >= najmoccrn And zanka = 1 Then
    If poscrn >= najposcrn Then
        'pa poglejmo še eno potezo naprej
        If narisana = True Then RaiseEvent RestoreBoard(Position)
        najdel = True
        RaiseEvent DrawMove(moznapoteza)
        If preveriBelega(kopijapolozajev, najmoccrn) = "" Then
            RaiseEvent RestoreBoard(Position)
            GoTo naprej
            End If
        If AliJeBranjena(kopijapolozajev, moznapoteza, False) = False Then
            GoTo naprej
            End If
            
        If poscrn > najposcrn Then
            For g = 1 To najpoteza.Count
            najpoteza.Remove (1)
            Next
            najpoteza.Add moznapoteza
            End If
        If poscrn = najposcrn Then
            najpoteza.Add moznapoteza
            End If
        dodal = True
        najmoccrn = moccrn
        najposcrn = poscrn
        narisana = True
        End If
    End If




If moccrn >= najmoccrn And zanka = 2 Then
    If poscrn >= najposcrn Then
        'pa poglejmo še eno potezo naprej
        If narisana = True Then RaiseEvent RestoreBoard(Position)
        najdel = True
        RaiseEvent DrawMove(moznapoteza)
        If preveriBelega(kopijapolozajev, najmoccrn) = "" And AliJeBranjena(kopijapolozajev, moznapoteza, False) = False Then
            RaiseEvent RestoreBoard(Position)
            GoTo naprej
            End If
        If poscrn > najposcrn Then
            For g = 1 To najpoteza.Count
            najpoteza.Remove (1)
            Next
            najpoteza.Add moznapoteza
            End If
        If poscrn = najposcrn Then
            najpoteza.Add moznapoteza
            End If
        dodal = True
        najmoccrn = moccrn
        najposcrn = poscrn
        narisana = True
        End If
    End If



naprej:
'kaj pa beli? ali kaj oslabi?
     If mocbel < najmocbel Then
        'o ja, oslabi!
        If AliJeBranjena(kopijapolozajev, moznapoteza, False) = True Then najpotezasibek.Add moznapoteza
        End If
    
    Next
If najpoteza.Count = 0 And najpotezasibek.Count = 0 Then GoTo Ponovi

If najpotezasibek.Count > 0 Then
    'èe že ne moremo umakniti figure, mu naredimo èimveè škode!
    RaiseEvent FoundMove(najpotezasibek(1), "")
    Exit Sub
    End If




'Debug.Print "poteze:" & Chr(13)
'For g = 1 To najpoteza.Count
'Debug.Print najpoteza(g)
'Next
'Debug.Print Chr(13)
'Debug.Print "Najšibek:"
'For g = 1 To najpotezasibek.Count
'Debug.Print najpotezasibek(g)
'Next
'Debug.Print Chr(13)
If najpotezasibek.Count > 0 And najpoteza.Count = 0 Then
    'èe že ne moremo umakniti figure, mu naredimo èimveè škode!
    RaiseEvent FoundMove(najpotezasibek(1), "")
    Exit Sub
    End If

'èe pa ob prvem iskanju ni poteze, je treba najti najbolj pozicijski premik

g = Int(Rnd * najpoteza.Count) + 1

RaiseEvent FoundMove(najpoteza(g), "")
End Sub

Function vrniVseMoznePoteze(ByVal polozaji As String, ByRef crninapotezi As Boolean) As String
'funkcija je loèena, ker se bo klicala rekurzivno
Dim tempVal As Integer
Dim n As Integer, m As Integer
Dim figura As String, spolozaj As String
Dim xoff, yoff

pars.BlackTurn = crninapotezi

'pars.MoveSearch = False

Dim kam As String, g As Integer
Dim evpot As String
    
For m = 1 To 8
    For n = 1 To 8
    spolozaj = vrniZnakLinije(n) & vrniZnakVrste(m)
    figura = vrniFiguro(polozaji, spolozaj)
    'je figura iste barve kot igralec?
        If Trim(figura) = "" Then GoTo naslednja
        If Left(figura, 1) = "B" And crninapotezi = True Then GoTo naslednja
        If Left(figura, 1) = "C" And crninapotezi = False Then GoTo naslednja
        
    'je iste barve
    'kam se lahko premakne?
    'gledali bomo vseh osem smeri
    
      If crnroslahkodesno = True And crnroslahkolevo = True And crninapotezi = True Then test = True
      If belroslahkodesno = True And belroslahkolevo = True And crninapotezi = False Then test = True
        
    'pa kralja, èe lahko izvede rošado
    If Right(figura, 1) <> "K" Then GoTo nikralj
    vrniRosadnePoteze kam, spolozaj, crninapotezi
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        'preverimo še, ali je rošada s potezo onemogoèena...
        pars.MoveSearch = test
        
goon:
                
        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
               
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno
        
        pars.Parse (polozaji)
        
        If (pars.BlackCastled = True Or pars.WhiteCastled = True) And pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            'RaiseEvent DrawCastle(evpot)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            
            End If
        Next
        End If
    
nikralj:
    'nazaj
       
    vrniVsePoteze polozaji, kam, spolozaj, 0, -1, crninapotezi
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno
        
        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
        pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            'RaiseEvent DrawMove(evpot)
            End If
        Next
        
        End If
    
    'naprej
    vrniVsePoteze polozaji, kam, spolozaj, 0, 1, crninapotezi
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno

        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
                pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            'RaiseEvent DrawMove(evpot)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            End If
        Next
        End If
    
    'levo
    vrniVsePoteze polozaji, kam, spolozaj, -1, 0, crninapotezi
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno

        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
            pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            'RaiseEvent DrawMove(evpot)
            End If
        Next
        End If
    
    'desno
    vrniVsePoteze polozaji, kam, spolozaj, 1, 0, crninapotezi
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno

        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
        pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            'RaiseEvent DrawMove(evpot)
            End If
        Next
        End If
    
    'levo gor
    vrniVsePoteze polozaji, kam, spolozaj, -1, 1, crninapotezi
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno

        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
        pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            'RaiseEvent DrawMove(evpot)
            End If
        Next
        End If
    
    'desno gor
    vrniVsePoteze polozaji, kam, spolozaj, 1, 1, crninapotezi
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno
        
        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
        
        pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            'RaiseEvent DrawMove(evpot)
            End If
        Next
        End If
    
    'levo dol
    vrniVsePoteze polozaji, kam, spolozaj, -1, -1, crninapotezi
        If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno

        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
                pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            'RaiseEvent DrawMove(evpot)
            End If
        Next
        End If
    
    'desno dol
    vrniVsePoteze polozaji, kam, spolozaj, 1, -1, crninapotezi
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno

        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
            pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            'RaiseEvent DrawMove(evpot)
            End If
        Next
        End If
    
    'še konje...
    'za vsako figuro naj velja, da je kojn...
    'pa je kojn?
    
    If Right(figura, 1) <> "S" Then GoTo naslednja
    
    vrnivsekonjskepoteze kam, spolozaj
    If Len(kam) > 0 Then
        pars.BlackTurn = crninapotezi
        pars.BlackCanCastleLeft = crnroslahkolevo
        pars.BlackCanCastleRight = crnroslahkodesno
        pars.WhiteCanCastleLeft = belroslahkolevo
        pars.WhiteCanCastleRight = belroslahkodesno
       
        For g = 1 To Len(kam) Step 2
        DoEvents
        pars.Move = spolozaj & "-" & Mid(kam, g, 2)
        pars.Parse (polozaji)
        If pars.ErrorNumber = 0 Then
            'poteza je dovoljena!
            evpot = spolozaj & "-" & Mid(kam, g, 2)
            vrniVseMoznePoteze = vrniVseMoznePoteze & evpot & ";"
            'RaiseEvent DrawMove(evpot)
            End If
        Next
        End If
    
naslednja:
    Next
    Next

End Function


Sub vrniRosadnePoteze(ByRef komb As String, ByVal polozaj As String, ByVal crni As Boolean)
If crni = True Then
    komb = "C8G8"
    Else
    komb = "C1G1"
    End If
    
End Sub



Sub izvrsiPotezo(ByVal poteza As String, ByRef polozaji As String)
Dim figura As String
figura = vrniFiguro(polozaji, Left(poteza, 2))
postaviFiguro polozaji, Right(poteza, 2), figura
postaviFiguro polozaji, Left(poteza, 2), "  "
End Sub

Function izracunajNasprotnika(ByRef polozaji As String, ByVal crni As Boolean) As Integer
Dim n As Integer
Dim m As Integer
Dim figura As String
Dim figbele As Byte
Dim figcrne As Byte

'najprej bomo prešteli figure
prestejFigure polozaji, figbele, figcrne

'obdelamo celo plošèo
For n = 1 To 8  'x
    For m = 1 To 8  'y
    DoEvents
    figura = vrniFiguro(polozaji, vrniZnakLinije(n) & vrniZnakVrste(m))
    'ugotovimo, ali barva nasprotnika ustreza, in ali je na polju sploh figura
    If Trim(figura) = "" Then GoTo naspdalje
    If (Left(figura, 1) = "B" And crni = False) Or (Left(figura, 1) = "C" And crni = True) Then GoTo naspdalje:
    'barva ustreza. Prištejemo ustrezno vrednost
    
'    MsgBox figura
    
    Select Case Right(figura, 1)
    Case "P"
    'kmet ima vrednost 1
    izracunajNasprotnika = izracunajNasprotnika + 1
    
    Case "T"
    'trdnjava ima vrednost 5
    izracunajNasprotnika = izracunajNasprotnika + 5
    
    Case "S"
    'konj ima vrednost 3, èe je figur manj, sicer 4
        Select Case crni
            Case True
            If figcrne < 9 Then
                izracunajNasprotnika = izracunajNasprotnika + 3
                Else
                izracunajNasprotnika = izracunajNasprotnika + 4
                End If
            Case False
                If figbele < 9 Then
                izracunajNasprotnika = izracunajNasprotnika + 3
                Else
                izracunajNasprotnika = izracunajNasprotnika + 4
                End If
            End Select
            
    Case "L"
    'tu je pa ravno nasprotno od konja
        Select Case crni
            Case True
            If figcrne < 9 Then
                izracunajNasprotnika = izracunajNasprotnika + 4
                Else
                izracunajNasprotnika = izracunajNasprotnika + 3
                End If
            Case False
                If figbele < 9 Then
                izracunajNasprotnika = izracunajNasprotnika + 4
                Else
                izracunajNasprotnika = izracunajNasprotnika + 3
                End If
            End Select
    Case "Q"
                izracunajNasprotnika = izracunajNasprotnika + 10
    End Select
    
naspdalje:
    Next
Next


End Function


Sub prestejFigure(ByRef polozaji As String, ByRef figbele, ByRef figcrne)
figbele = 0
figcrne = 0
Dim n  As Integer
Dim m  As Integer
Dim figura As String * 2

For n = 1 To 8  'x
    For m = 1 To 8  'y
    figura = vrniFiguro(polozaji, vrniZnakLinije(n) & vrniZnakVrste(m))
    'ugotovimo, ali barva nasprotnika ustreza
    If Trim(figura) = "" Then GoTo stejdalje
    If Left(figura, 1) = "B" Then
        figbele = figbele + 1
        Else
        figcrne = figcrne + 1
        End If
    
stejdalje:
    Next
Next

End Sub


Function izracunajPozicijoNasprotnika(ByRef posicija As String, ByRef crniigra As Boolean) As Integer
'funkcija vrne seštevek polj (tudi tistih, na katerih so nasprotnikove figure)
Dim n As Integer
Dim polje As String
Dim polja As String


For n = 1 To Len(posicija) Step 6
polje = Mid(posicija, n, 2)

If Trim(Mid(posicija, n + 3, 2)) <> "" Then
    'èe so figure iste barve, nam tega polja ni treba šteti!
    If crniigra = True And Left(vrniFiguro(posicija, polje), 1) = "C" Then GoTo racnapr
    If crniigra = False And Left(vrniFiguro(posicija, polje), 1) = "B" Then GoTo racnapr
    End If

If aliJeSah(polje, Not (crniigra)) = True Then
    polja = polja & polje & " "
    'da, polje je pod sahom
    izracunajPozicijoNasprotnika = izracunajPozicijoNasprotnika + 1
    
    Select Case polje
    'èe so polja sredinska, so vsaj dvakrat vredna!
    Case "E4", "D4", "E5", "D5"
    izracunajPozicijoNasprotnika = izracunajPozicijoNasprotnika + 1
    End Select
    
    End If
racnapr:
Next

'MsgBox "Nadzorujem polja:" & Len(polja) / 2
'MsgBox "Nadzorujem polja:" & polja
End Function


Function chksublevel(ByVal tpozicija As String, ByVal crninapotezi As Boolean, ByVal globina As Integer) As Boolean


Dim moznepoteze As String
Dim n As Integer
Dim moznapoteza As String
Dim najpoteza As String
Dim cfig As Integer
Dim bfig As Integer
Dim najfig As Integer

'èe je 'èrninapotezi' true, se nivo prešteje!



moznepoteze = vrniVseMoznePoteze(tpozicija, crninapotezi)


chksublevel = False
Dim moc As Integer
Dim najmoc As Integer

For n = 1 To Len(moznepoteze) Step 6 'vsak sedmi znak je zaèetek poteze
    DoEvents
    'osvežimo položaje
    kopijapolozajev = tpozicija

    'potezo imamo
    moznapoteza = Mid(moznepoteze, n, 5)


izvrsiPotezo moznapoteza, kopijapolozajev

If globina >= mvarLevels Then GoTo nosublevel



'konec razmisljanja?


If chksublevel(kopijapolozajev, Not (crninapotezi), globina + 1) = False Then GoTo skipajtole
        
    'ko smo iz podnivoja dobili rezultate, jih lahko primerjamo
    
nosublevel:

'MsgBox "racunam crnega"
cfig = izracunajNasprotnika(kopijapolozajev, False)
bfig = izracunajNasprotnika(kopijapolozajev, True)

GoTo racmoc

If crninapotezi = True Then
    'èe je crni na potezi...
If bfig >= cfig Then
    GoTo racmoc
    Else
    najpoteza = moznapoteza
    End If
    
    Else

If bfig <= cfig Then
    'oziroma beli...
    GoTo racmoc
    Else
    najpoteza = moznapoteza
    End If
    
    End If
    
racmoc:
If crninapotezi = True Then moc = izracunajPozicijoNasprotnika(kopijapolozajev, Not crninapotezi)
If crninapotezi = False Then moc = izracunajPozicijoNasprotnika(kopijapolozajev, crninapotezi)

If moc > najmoc Then
        najmoc = moc
        najpoteza = moznapoteza
        End If



skipajtole:
Next

If najpoteza = "" Then
    chksublevel = False
    Else
    Debug.Print "Podnivo "; globina & " izbral:" & najpoteza
    chksublevel = True
    RaiseEvent DrawMove(najpoteza)
    End If



End Function



Function preveriBelega(polo As String, mocCrnega As Integer) As String



DoEvents
Dim najmocbel As Integer, mocbel As Integer
Dim mislicrnega As Boolean
Dim narisana As Boolean
Dim moznepoteze As String
Dim moznapoteza As String
Dim najpoteza As String
Dim n As Integer
Dim kopijapolozajev As String
'!!!!!!!!!!!!!
Dim moccrn As Integer
Dim najmoccrn As Integer
Dim najposcrn As Integer
Dim poscrn As Integer
Dim posda As Boolean
Dim mocda As Boolean
Dim zacposcrni As Integer
Dim zacmoccrni As Integer
Dim najdel As Boolean



'Sedaj igra beli!
mislicrnega = False


moznepoteze = vrniVseMoznePoteze(polo, mislicrnega)
najpoteza = ""
poscrn = 0
najposcrn = 0
moccrn = 0
najmoccrn = 0


'najmoccrn = izracunajNasprotnika(polo, (mislicrnega))
najmoccrn = mocCrnega
najmocbel = izracunajNasprotnika(polo, Not (mislicrnega))
For n = 1 To Len(moznepoteze) Step 6 'vsak sedmi znak je zaèetek poteze


'osvežimo položaje
kopijapolozajev = polo

'potezo imamo
moznapoteza = Mid(moznepoteze, n, 5)
'èe refernece ni, jo naredimo
If najpoteza = "" Then najpoteza = moznapoteza


izvrsiPotezo moznapoteza, kopijapolozajev
'DoEvents

'igramo belega; mislicrnega=false

moccrn = izracunajNasprotnika(kopijapolozajev, (mislicrnega))
mocbel = izracunajNasprotnika(kopijapolozajev, Not (mislicrnega))

'RaiseEvent RestoreBoard(polo)
'RaiseEvent DrawMove(moznapoteza)

If moccrn < najmoccrn Then
        
            'oslabitev!+ kolièinska prevlada belega
            preveriBelega = ""
            Exit Function
        
        End If
Next

If moznepoteze = "" Then
    MsgBox "Mate!"
    End If

preveriBelega = " "
End Function


Function AliJeBranjena(ByVal kopijapolozajev As String, moznapoteza As String, igraCrni As Boolean) As Boolean

'odstraniti moramo figuro...
postaviFiguro kopijapolozajev, Right(moznapoteza, 2), "  "
AliJeBranjena = aliJeSah(Right(moznapoteza, 2), igraCrni)
End Function
