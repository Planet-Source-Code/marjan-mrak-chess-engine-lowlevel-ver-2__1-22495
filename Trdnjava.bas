Attribute VB_Name = "Trdnjava"
Option Explicit
Function preveriTrdnjavo(ByVal crninavrsti As Boolean, ByVal poteza As String, ByVal figura As String) As Boolean
'kot prvo bi pogledali, ali je premik diagonalen,
'kar pomeni, da morata biti razliki med
'X in Y koordinatama enaki
Dim razX As Integer, razy As Integer

preveriTrdnjavo = True 'zaenkrat vse OK

vrniRazliko poteza, razX, razy


If Abs(razX) > 0 And Abs(razy) > 0 Then
    'trdnjava se ni premaknila pravokotno!
    preveriTrdnjavo = False
    Exit Function
    End If
    
'sedaj pa težji del
'ugotoviti je treba, ali trdnjavi kaj stoji na poti.

If diagonalaProsta(poteza) = False Then
    'trdnjava ne more na ciljno polje; diagonala!
    preveriTrdnjavo = False
    Exit Function
    End If

End Function

