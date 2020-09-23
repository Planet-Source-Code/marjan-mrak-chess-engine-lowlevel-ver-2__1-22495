Attribute VB_Name = "Risanje"
Option Explicit
Const faktor = 740
Global polozaji(9, 9) As String * 2
Global crnipripravljen As Boolean
Global pozrtih(2) As Integer
Global statusPoteze As String
Global belarosadamogocalevo As Boolean
Global crnarosadamogocalevo As Boolean
Global belarosadamogocadesno As Boolean
Global crnarosadamogocadesno As Boolean



Sub preracunajKoordinate(ByRef x As Single, ByRef y As Single, ByVal kam As String)
Dim kolicnik As Single

'najprej èrka
kolicnik = Asc(UCase(Left(kam, 1))) - 65
x = x + (kolicnik * faktor)

'potem številka
kolicnik = Asc(UCase(Right(kam, 1))) - 48
y = y - (kolicnik * faktor) 'dvigamo se

    
End Sub


Sub drawPiece(ByVal katero As String, ByVal kje As String)
Dim orgX As Single
Dim orgY As Single
Dim figindex As Integer

orgX = 300
orgY = sahovnica.Msahovnica.Height - 200 'dno šahovnice
preracunajKoordinate orgX, orgY, kje




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
       
    sahovnica.figurac.Picture = sahovnica.figc.GraphicCell(figindex)
    sahovnica.figurab.Picture = sahovnica.figb.GraphicCell(figindex)
        

'pa jo narišimo
sahovnica.PaintPicture sahovnica.figurab.Picture, orgX, orgY, , , , , , , vbSrcAnd
sahovnica.PaintPicture sahovnica.figurac.Picture, orgX, orgY, , , , , , , vbSrcPaint
DoEvents

End Sub

Sub displayEmptyBoard()
sahovnica.Width = sahovnica.Msahovnica.Width
sahovnica.Height = sahovnica.Msahovnica.Height + 1000
sahovnica.PaintPicture sahovnica.Msahovnica.Picture, 0, 0
sahovnica.Left = ploscax
sahovnica.Top = ploscay

Dim n As Integer, m As Integer
'izpraznemo šahovnico
crnarosadamogocadesno = True
belarosadamogocadesno = True
crnarosadamogocalevo = True
belarosadamogocalevo = True

For n = 1 To 8
For m = 1 To 8
    polozaji(n, m) = "  "
    minipolozaji(n, m) = "  "
    Next
    Next

sahovnica.Show
DoEvents
End Sub

Sub DisplayCompleteBoard()
Dim n As Integer
sahovnica.Width = sahovnica.Msahovnica.Width
sahovnica.Height = sahovnica.Msahovnica.Height + 1000
sahovnica.Left = ploscax
sahovnica.Top = ploscay
crnarosadamogocadesno = True
belarosadamogocadesno = True
crnarosadamogocalevo = True
belarosadamogocalevo = True

'izpraznemo vsebino
Dim m As Integer
For n = 1 To 8
For m = 1 To 8
    polozaji(n, m) = "  "
    minipolozaji(n, m) = "  "
    Next
    Next

sahovnica.PaintPicture sahovnica.Msahovnica.Picture, 0, 0
'beli kmetje
drawPiece "BP", "A2"
drawPiece "BP", "B2"
drawPiece "BP", "c2"
drawPiece "BP", "d2"
drawPiece "BP", "e2"
drawPiece "BP", "f2"
drawPiece "BP", "g2"
drawPiece "BP", "h2"
'še zafilamo tabelo
For n = 1 To 8
    polozaji(n, 2) = "BP"
    minipolozaji(n, 2) = "BP"
    Next
'èrni kmetje
drawPiece "cP", "A7"
drawPiece "cP", "B7"
drawPiece "cP", "c7"
drawPiece "cP", "d7"
drawPiece "cP", "e7"
drawPiece "cP", "f7"
drawPiece "cP", "g7"
drawPiece "cP", "h7"
For n = 1 To 8
    polozaji(n, 7) = "CP"
    minipolozaji(n, 7) = "CP"
    Next


'beli trdnjavi
drawPiece "bt", "A1"
polozaji(1, 1) = "BT"
minipolozaji(1, 1) = "BT"
drawPiece "bt", "h1"
polozaji(8, 1) = "BT"

'èrni trdnjavi
drawPiece "ct", "A8"
polozaji(1, 8) = "CT"
drawPiece "ct", "h8"
polozaji(8, 8) = "CT"

'bela konja
drawPiece "bs", "b1"
polozaji(2, 1) = "BS"
drawPiece "bs", "g1"
polozaji(7, 1) = "BS"

'èrna konja
drawPiece "cs", "b8"
polozaji(2, 8) = "CS"
drawPiece "cs", "g8"
polozaji(7, 8) = "CS"

'bela lovca
drawPiece "bl", "c1"
polozaji(3, 1) = "BL"
drawPiece "bl", "f1"
polozaji(6, 1) = "BL"

'crna lovca
drawPiece "cl", "c8"
polozaji(3, 8) = "CL"
drawPiece "cl", "f8"
polozaji(6, 8) = "CL"


'bela kraljica
drawPiece "bq", "d1"
polozaji(4, 1) = "BQ"

'crna kraljica
drawPiece "cq", "d8"
polozaji(4, 8) = "CQ"

'beli kralj
drawPiece "bk", "e1"
polozaji(5, 1) = "BK"

'èrni kralj
drawPiece "ck", "e8"
polozaji(5, 8) = "CK"

sahovnica.Show
DoEvents
End Sub

Sub displayWatch()
ura.Top = uray
ura.Left = urax
ura.Show
End Sub

Sub displaymoves()
Poteze.Left = potx
Poteze.Top = poty
Poteze.Show
End Sub

Function sestavipolozaje() As String
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

Sub deletePiece(ByVal kje As String)
'izbrišemo zahtevan položaj
Dim orgX As Single
Dim orgY As Single
orgX = 300
orgY = sahovnica.Msahovnica.Height - 200 'dno šahovnice
preracunajKoordinate orgX, orgY, kje

sahovnica.PaintPicture _
sahovnica.Msahovnica.Picture, _
orgX, orgY, , , orgX, orgY, _
sahovnica.figurab.Width, sahovnica.figurab.Height
DoEvents
End Sub

Sub movepiece(ByVal poteza As String)
Dim ime As String, prvapot As String, drugapot As String
Dim linija As Integer, vrsta As Integer

statusPoteze = ""

'najprej izbrišemo figuro
prvapot = Left(poteza, 2)
drugapot = Right(poteza, 2)


deletePiece prvapot 'izpraznemo štart
deletePiece drugapot 'in cilj
'vzamemo ime figure
linija = Asc(UCase(Left(prvapot, 1))) - 64
vrsta = Asc(UCase(Right(prvapot, 1))) - 48

ime = polozaji(linija, vrsta)
polozaji(linija, vrsta) = "  " 'izpraznemo polozaj



linija = Asc(UCase(Left(drugapot, 1))) - 64
vrsta = Asc(UCase(Right(drugapot, 1))) - 48

'pogledati je treba, ali se na destinaciji nahaja kaka figura!
'èe se , potem jo je treba narisati v 'POŽRTE' oknu
If Trim(polozaji(linija, vrsta)) <> "" Then
    Dim pozrta As String
    
    'katero figuro bomo pojedli?
    pozrta = polozaji(linija, vrsta)
    
    drawcaptured (pozrta)
    
    If crnipripravljen = False Then
        pozrtih(0) = pozrtih(0) + 1
        Else
        pozrtih(1) = pozrtih(1) + 1
        End If
    End If

polozaji(linija, vrsta) = ime

drawPiece ime, drugapot
End Sub

Sub displaycaptured()
pozrte.Left = pozx
pozrte.Top = pozy
pozrte.Show
End Sub

Sub drawcaptured(ByVal kaj As String)
Dim orgX As Single
Dim orgY As Single
Dim figindex As Integer, pozindex As Integer

If crnipripravljen = False Then
    pozindex = 0
    Else
    pozindex = 1
    End If

orgX = pozrtih(pozindex) * sahovnica.figurab.Width

'v bistvu je treba prezrcaliti barvi, kajti figure gredo nasprotniku!

If crnipripravljen = False Then
    orgY = sahovnica.figurab.Height * 2 'èe je figuro vzel beli, jo je treba
    Else        'narisati èrnemu
    orgY = 300
    End If


Select Case UCase(Right(kaj, 1))
    
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
        
    If UCase(Left(kaj, 1)) = "C" Then figindex = figindex + 6
       
    sahovnica.figurac.Picture = sahovnica.figc.GraphicCell(figindex)
    sahovnica.figurab.Picture = sahovnica.figb.GraphicCell(figindex)
        

'pa jo narišimo
pozrte.PaintPicture sahovnica.figurab.Picture, orgX, orgY, , , , , , , vbSrcAnd
pozrte.PaintPicture sahovnica.figurac.Picture, orgX, orgY, , , , , , , vbSrcPaint
End Sub

