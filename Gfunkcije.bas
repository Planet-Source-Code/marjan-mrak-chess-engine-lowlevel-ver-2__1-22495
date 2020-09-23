Attribute VB_Name = "Gfunkcije"
Option Explicit
Global pozx As Integer, pozy As Integer
Global pcx As Integer, pcy As Integer
Global urax As Integer, uray As Integer
Global potx As Integer, poty As Integer
Global ploscax As Integer, ploscay As Integer
Global zapStPot As Integer
Global crniigra As Boolean
Type cas
sekunde As Byte
minute As Byte
ure As Byte
End Type
Global tipka As Boolean



Function vrniVrednostKoordinate(ByVal kje As String) As Integer
'se gre za številko?
If Asc(kje) < 57 Then
    'ja!
    vrniVrednostKoordinate = Asc(kje) - 48
    Else
    'ne, je èrka
    vrniVrednostKoordinate = Asc(kje) - 64
    End If
End Function

Function vrniZnakLinije(ByVal vrednost As Integer) As String
vrniZnakLinije = Chr(vrednost + 64)
End Function
Function vrniZnakVrste(ByVal vrednost As Integer) As String
vrniZnakVrste = Chr(vrednost + 48)
End Function

Sub preberiVrstico(ByVal tekst As String, ByRef x As Integer, ByRef y As Integer, ByRef lahko As Boolean)
Dim pol As Integer
Dim pola As Integer

pol = InStr(tekst, ":")
tekst = Mid(tekst, pol + 1)

pol = InStr(tekst, ",")
x = CInt(Left(tekst, pol - 1))

tekst = Mid(tekst, pol + 1)

'še za drugo vejico
pol = InStr(tekst, ",")
If pol = 0 Then
    y = CInt(tekst)
    Exit Sub
    End If
    
'vejica je bila, se pravi...
y = CInt(Left(tekst, pol - 1))
tekst = Mid(tekst, pol + 1)

'še true/false
tekst = UCase(tekst)

If Left(tekst, 1) = "R" Or Left(tekst, 1) = "T" Then lahko = True
If Left(tekst, 1) = "N" Or Left(tekst, 1) = "F" Then lahko = False
End Sub
