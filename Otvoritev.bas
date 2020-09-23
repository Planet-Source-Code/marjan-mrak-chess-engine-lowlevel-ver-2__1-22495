Attribute VB_Name = "Otvoritev"
Option Explicit


Sub vrniVsePoteze(ByRef polozaji As String, ByRef kam As String, ByVal odkod As String, ByVal Xoffset As Integer, Yoffset As Integer, ByVal blackmoves As Boolean)
Dim x As Integer, y As Integer
Dim tkam As String, tfigura As String
kam = ""
x = vrniVrednostKoordinate(Left(odkod, 1))
y = vrniVrednostKoordinate(Right(odkod, 1))


Do
x = x + Xoffset
If x > 8 Or x < 1 Then Exit Do
y = y + Yoffset
If y < 1 Or y > 8 Then Exit Do

tkam = vrniZnakLinije(x) & vrniZnakVrste(y)

'èe je na polju figura iste barve kot igralec, je poteza nesmiselna

tfigura = Trim(vrniFiguro(polozaji, tkam))
If tfigura <> "" And Left(tfigura, 1) = "C" And blackmoves = False Then
    kam = kam & tkam
    End If
If tfigura <> "" And Left(tfigura, 1) = "B" And blackmoves = True Then
    kam = kam & tkam
    End If
If tfigura = "" Then kam = kam + tkam

Loop
End Sub

Sub vrnivsekonjskepoteze(ByRef seznam, ByVal kje As String)
seznam = ""
Dim x As Integer, y As Integer

x = vrniVrednostKoordinate(Left(kje, 1))
y = vrniVrednostKoordinate(Right(kje, 1))

'naprej levo
If ((x - 1) > 0) And ((y + 2) < 9) Then
    seznam = seznam & vrniZnakLinije(x - 1) & vrniZnakVrste(y + 2)
    End If

'naprej desno
If ((x + 1) < 9) And ((y + 2) < 9) Then
    seznam = seznam & vrniZnakLinije(x + 1) & vrniZnakVrste(y + 2)
    End If

'nazaj levo
If ((x - 1) > 0) And ((y - 2) > 0) Then
    seznam = seznam & vrniZnakLinije(x - 1) & vrniZnakVrste(y - 2)
    End If
'nazaj desno
If ((x + 1) < 9) And ((y - 2) > 0) Then
    seznam = seznam & vrniZnakLinije(x + 1) & vrniZnakVrste(y - 2)
    End If

'levo zgoraj
If ((x - 2) > 0) And ((y + 1) < 9) Then
    seznam = seznam & vrniZnakLinije(x - 2) & vrniZnakVrste(y + 1)
    End If

'levo spodaj
If ((x - 2) > 0) And ((y - 1) > 0) Then
    seznam = seznam & vrniZnakLinije(x - 2) & vrniZnakVrste(y - 1)
    End If

'desno zgoraj
If ((x + 2) < 9) And ((y + 1) < 9) Then
    seznam = seznam & vrniZnakLinije(x + 2) & vrniZnakVrste(y + 1)
    End If

'desno spodaj
If ((x + 2) < 9) And ((y - 1) > 0) Then
    seznam = seznam & vrniZnakLinije(x + 2) & vrniZnakVrste(y - 1)
    End If


End Sub
