Attribute VB_Name = "risanjeEngine"
Option Explicit

Sub minipreracunajKoordinate(ByRef x As Single, ByRef y As Single, ByVal kam As String)
Dim kolicnik As Single

'najprej èrka
kolicnik = Asc(UCase(Left(kam, 1))) - 65
x = x + (kolicnik * minifaktor)

'potem številka
kolicnik = Asc(UCase(Right(kam, 1))) - 48
y = y - (kolicnik * minifaktor) 'dvigamo se

    
End Sub


Sub minidrawPiece(ByVal katero As String, ByVal kje As String)
Dim orgX As Single
Dim orgY As Single
Dim figindex As Integer

orgX = 50
orgY = miniSahovnica.Msahovnica.Height - 80 'dno šahovnice
minipreracunajKoordinate orgX, orgY, kje




Select Case UCase(Right(katero, 1))
    
        Case "P"
        'kmet
        figindex = 0
        
        Case "T"
        'trdnjava
        figindex = 1
                
        Case "S"
        'skakaè
        figindex = 2
               
        Case "L"
        'lovec
        figindex = 3
        
        Case "Q"
        'kraljica
        figindex = 4
        
        Case "K"
        'kralj
        figindex = 5
        End Select
        
    If UCase(Left(katero, 1)) = "C" Then figindex = figindex + 6
       
    miniSahovnica.figurac.Picture = miniSahovnica.figc.GraphicCell(figindex)
    miniSahovnica.figurab.Picture = miniSahovnica.figb.GraphicCell(figindex)
        

'pa jo narišimo
miniSahovnica.PaintPicture miniSahovnica.figurab.Picture, orgX, orgY, , , , , , , vbSrcAnd
miniSahovnica.PaintPicture miniSahovnica.figurac.Picture, orgX, orgY, , , , , , , vbSrcPaint


End Sub

Sub minidisplayEmptyBoard()
miniSahovnica.Top = sahovnica.Height
miniSahovnica.Left = 0
miniSahovnica.Width = miniSahovnica.Msahovnica.Width
miniSahovnica.Height = miniSahovnica.Msahovnica.Height + 320
miniSahovnica.miniPaintPicture sahovnica.Msahovnica.Picture, 0, 0
Dim n As Integer, m As Integer
'izpraznemo šahovnico
minicrnarosadamogocadesno = True
minibelarosadamogocadesno = True
minicrnarosadamogocalevo = True
minibelarosadamogocalevo = True

For n = 1 To 8
For m = 1 To 8
    minipolozaji(n, m) = "  "
    Next
    Next

miniSahovnica.Show
End Sub

Sub miniDisplayCompleteBoard()
Dim n As Integer
miniSahovnica.Width = miniSahovnica.Msahovnica.Width
miniSahovnica.Height = miniSahovnica.Msahovnica.Height + 320
miniSahovnica.Top = sahovnica.Height
miniSahovnica.Left = 0
minicrnarosadamogocadesno = True
minibelarosadamogocadesno = True
minicrnarosadamogocalevo = True
minibelarosadamogocalevo = True

'izpraznemo vsebino
Dim m As Integer
For n = 1 To 8
For m = 1 To 8
    minipolozaji(n, m) = "  "
    Next
    Next

miniSahovnica.PaintPicture miniSahovnica.Msahovnica.Picture, 0, 0
'beli kmetje
minidrawPiece "BP", "A2"
minidrawPiece "BP", "B2"
minidrawPiece "BP", "c2"
minidrawPiece "BP", "d2"
minidrawPiece "BP", "e2"
minidrawPiece "BP", "f2"
minidrawPiece "BP", "g2"
minidrawPiece "BP", "h2"
'še zafilamo tabelo
For n = 1 To 8
    minipolozaji(n, 2) = "BP"
    Next
'èrni kmetje
minidrawPiece "cP", "A7"
minidrawPiece "cP", "B7"
minidrawPiece "cP", "c7"
minidrawPiece "cP", "d7"
minidrawPiece "cP", "e7"
minidrawPiece "cP", "f7"
minidrawPiece "cP", "g7"
minidrawPiece "cP", "h7"
For n = 1 To 8
    minipolozaji(n, 7) = "CP"
    Next


'beli trdnjavi
minidrawPiece "bt", "A1"
minipolozaji(1, 1) = "BT"
minidrawPiece "bt", "h1"
minipolozaji(8, 1) = "BT"

'èrni trdnjavi
minidrawPiece "ct", "A8"
minipolozaji(1, 8) = "CT"
minidrawPiece "ct", "h8"
minipolozaji(8, 8) = "CT"

'bela konja
minidrawPiece "bs", "b1"
minipolozaji(2, 1) = "BS"
minidrawPiece "bs", "g1"
minipolozaji(7, 1) = "BS"

'èrna konja
minidrawPiece "cs", "b8"
minipolozaji(2, 8) = "CS"
minidrawPiece "cs", "g8"
minipolozaji(7, 8) = "CS"

'bela lovca
minidrawPiece "bl", "c1"
minipolozaji(3, 1) = "BL"
minidrawPiece "bl", "f1"
minipolozaji(6, 1) = "BL"

'crna lovca
minidrawPiece "cl", "c8"
minipolozaji(3, 8) = "CL"
minidrawPiece "cl", "f8"
minipolozaji(6, 8) = "CL"


'bela kraljica
minidrawPiece "bq", "d1"
minipolozaji(4, 1) = "BQ"

'crna kraljica
minidrawPiece "cq", "d8"
minipolozaji(4, 8) = "CQ"

'beli kralj
minidrawPiece "bk", "e1"
minipolozaji(5, 1) = "BK"

'èrni kralj
minidrawPiece "ck", "e8"
minipolozaji(5, 8) = "CK"

miniSahovnica.Show
End Sub
Function minisestavipolozaje() As String
'tu moramo iz polozajev sestaviti string:
Dim x As Integer, y As Integer
Dim pol As String
pol = ""

For y = 1 To 8
    For x = 1 To 8
    If Trim(Left(polozaji(x, y), 1)) = Chr(0) Then polozaji(x, y) = "  "
    
    pol = pol & Chr(x + 64) & Trim(Str(y)) & ":" & polozaji(x, y) & "|"
    Next
Next
sestavipolozaje = pol
End Function

Sub minideletePiece(ByVal kje As String)
'izbrišemo zahtevan položaj
Dim orgX As Single
Dim orgY As Single
orgX = 50
orgY = miniSahovnica.Msahovnica.Height - 80 'dno šahovnice
minipreracunajKoordinate orgX, orgY, kje

miniSahovnica.PaintPicture _
miniSahovnica.Msahovnica.Picture, _
orgX, orgY, , , orgX, orgY, _
miniSahovnica.figurab.Width, miniSahovnica.figurab.Height
End Sub

Sub minimovepiece(ByVal poteza As String)
Dim ime As String, prvapot As String, drugapot As String
Dim linija As Integer, vrsta As Integer

statusPoteze = ""

'najprej izbrišemo figuro
prvapot = Left(poteza, 2)
drugapot = Right(poteza, 2)


minideletePiece prvapot 'izpraznemo štart
minideletePiece drugapot 'in cilj
'vzamemo ime figure
linija = Asc(UCase(Left(prvapot, 1))) - 64
vrsta = Asc(UCase(Right(prvapot, 1))) - 48

ime = minipolozaji(linija, vrsta)
minipolozaji(linija, vrsta) = "  " 'izpraznemo polozaj



linija = Asc(UCase(Left(drugapot, 1))) - 64
vrsta = Asc(UCase(Right(drugapot, 1))) - 48

minipolozaji(linija, vrsta) = ime

minidrawPiece ime, drugapot
End Sub







