Attribute VB_Name = "Kraljica"
Option Explicit

Function preveriKraljico(ByVal crninavrsti As Boolean, ByVal poteza As String, ByVal figura As String) As Boolean
'kot prvo bi pogledali, ali je premik diagonalen,
'kar pomeni, da morata biti razliki med
'X in Y koordinatama enaki
Dim razX As Integer, razy As Integer

preveriKraljico = True 'zaenkrat vse OK

vrniRazliko poteza, razX, razy

'preveriti premik
'kraljica gre lahko diagonalno ali pravokotno!
If Abs(razX) <> Abs(razy) Then
    'kraljica ni šla diagonalno
    'je šla pravokotno
    If Abs(razX) > 1 And Abs(razy) > 1 Then
        'ne, ni šla pravokotno! napaka
        preveriKraljico = False
        Exit Function
        End If
    End If


    
'sedaj pa težji del
'ugotoviti je treba, ali trdnjavi kaj stoji na poti.

If diagonalaProsta(poteza) = False Then
    'kraljica ne more na ciljno polje; diagonala!
    preveriKraljico = False
    Exit Function
    End If

End Function


