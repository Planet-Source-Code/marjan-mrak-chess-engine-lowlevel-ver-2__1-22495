Attribute VB_Name = "Lovec"
Option Explicit
Function preveriLovca(ByVal crninavrsti As Boolean, ByVal poteza As String, ByVal figura As String) As Boolean
'kot prvo bi pogledali, ali je premik diagonalen,
'kar pomeni, da morata biti razliki med
'X in Y koordinatama enaki
Dim razX As Integer, razy As Integer

preveriLovca = True 'zaenkrat vse OK

vrniRazliko poteza, razX, razy


If Abs(razX) <> Abs(razy) Then
    'lovec se ni premaknil diagonalno!
    preveriLovca = False
    Exit Function
    End If
    
'sedaj pa težji del
'ugotoviti je treba, ali lovcu kaj stoji na poti.

If diagonalaProsta(poteza) = False Then
    'lovec ne more na ciljno polje; diagonala!
    preveriLovca = False
    Exit Function
    End If

End Function
