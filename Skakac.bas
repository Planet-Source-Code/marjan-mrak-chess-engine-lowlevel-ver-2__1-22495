Attribute VB_Name = "skakac"
Option Explicit
Function preveriSkakaca(ByVal crninavrsti As Boolean, ByVal poteza As String, ByVal figura As String) As Boolean

'najprej preverimo, èe je konèno mesto prosto
preveriSkakaca = True 'zaenkrat OK

'najprej premik
Dim razX As Integer, razy As Integer

vrniRazliko poteza, razX, razy

If Abs(razX) <> 1 And Abs(razX) <> 2 Then
    'ni 'skoèil'!
    preveriSkakaca = False
    Exit Function
    End If
    
If Abs(razy) <> 1 And Abs(razy) <> 2 Then
    'ni 'skoèil'!
    preveriSkakaca = False
    Exit Function
    End If

razX = Abs(razX)
razy = Abs(razy)

If razX = 1 And razy <> 2 Then
    'nepravilen skok
    preveriSkakaca = False
    Exit Function
    End If

If razX = 2 And razy <> 1 Then
    'nepravilen skok
    preveriSkakaca = False
    Exit Function
    End If


If razy = 1 And razX <> 2 Then
    preveriSkakaca = False
    Exit Function
    End If

If razy = 2 And razX <> 1 Then
    preveriSkakaca = False
    Exit Function
    End If


If Trim(vrniFiguro(polozaji, Mid(poteza, 4, 2))) <> "" Then
    'polje ni prosto
    Exit Function
    End If
End Function
