Attribute VB_Name = "Kralj"
Option Explicit

Function preveriKralja(ByVal crninavrsti As Boolean, ByVal poteza As String, ByVal figura As String) As Boolean
'kot prvo bi pogledali, ali je premik diagonalen,
'kar pomeni, da morata biti razliki med
'X in Y koordinatama enaki
Dim razX As Integer, razy As Integer

preveriKralja = True 'zaenkrat vse OK

vrniRazliko poteza, razX, razy

If Abs(razy) > 1 Then
    'neveljaven premik
    preveriKralja = False
    Exit Function
    End If
    

'preveriti premik
'kralj gre lahko diagonalno ali pravokotno!
If Abs(razX) <> Abs(razy) Then
    'kralj ni �el diagonalno
    'je �el pravokotno?
    If Abs(razX) > 1 And Abs(razy) > 1 Then
        'ne, ni �el pravokotno! napaka
        preveriKralja = False
        Exit Function
        End If
    End If

If Abs(razX) > 2 Then
    'kralj ni rokiral!
    preveriKralja = False
    Exit Function
    End If
    
'�e je kralj �el za dve polji levo ali desno, to pomeni ro�ado!
'najprej preverimo, kateri...
If poteza = "E1-G1" Then
    'poiskus bele ro�ade v levo
        preveriKralja = preveriRosado(polozaji, poteza, crninavrsti)
        If preveriKralja = False Then
        Exit Function
        End If
        GoTo ocrnirosado
        End If
    

If poteza = "E1-C1" Then
    'poiskus bele ro�ade v levo
        preveriKralja = preveriRosado(polozaji, poteza, crninavrsti)
        
        If preveriKralja = False Then
            Exit Function
            End If
            GoTo ocrnirosado
        End If
    

If poteza = "E8-C8" Then
    'poiskus bele ro�ade v levo
        preveriKralja = preveriRosado(polozaji, poteza, crninavrsti)
        If preveriKralja = False Then
            Exit Function
            End If
            GoTo ocrnirosado
        End If
    

If poteza = "E8-G8" Then
    'poiskus bele ro�ade v levo
        preveriKralja = preveriRosado(polozaji, poteza, crninavrsti)
        If preveriKralja = False Then
        Exit Function
        End If
        GoTo ocrnirosado
    End If

'sedaj pa je treba preveriti, ali je kralj �el za eno polje
If Abs(razX) > 1 Or Abs(razy) > 1 Then
        'kralj se je premaknil ve� kot eno polje
        preveriKralja = False
        Exit Function
        End If
        

ocrnirosado:
If test = True Then Exit Function

'ni �el. O�rni mo�nost ro�ade!
If Left(poteza, 2) = "E1" Then
    belroslahkodesno = False
    belroslahkolevo = False
    End If
    
If Left(poteza, 2) = "E8" Then
    crnroslahkodesno = False
    crnroslahkolevo = False
    End If
    
    
    
End Function



