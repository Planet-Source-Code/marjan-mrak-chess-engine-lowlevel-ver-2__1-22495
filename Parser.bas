Attribute VB_Name = "Parser"
Option Explicit
Global crnipripravljen As Boolean
'PARSER!!!!
Function sparsaj(ByVal poteza As String) As String
sparsaj = ""    'po odefoltu ni ni� narobe!
'najprej ugotovimo, ali je sploh kaka figura na polo�aju
'�isto najprej pa izra�un koordinat
Dim xkoord As Integer, ykoord As Integer
Dim znak As String * 1  '�rka oz. �t.

xkoord = Asc(Left(poteza, 1)) - 64
ykoord = Asc(Mid(poteza, 2, 1)) - 48
'pa poglejmo

If Trim(polozaji(xkoord, ykoord)) = "" Then
    'ni nobene figure!
    sparsaj = "Na za�etnem polo�aju ni nobene figure!"
    Exit Function
    End If
    

End Function
