Attribute VB_Name = "Kmet"
Option Explicit


Function preverikmeta(ByVal crninavrsti As Boolean, ByVal poteza As String, ByVal figura As String) As Boolean
Dim bel As Boolean
preverikmeta = True 'zaenkrat nimamo razloga za paniko

'kakšne barve je?
If Left(figura, 1) = "B" Then
    bel = True
    Else
    bel = False
    End If

'ali je razlika med premikom v Y osi 2?
Dim razY As Integer, razX As Integer

vrniRazliko poteza, razX, razY


If razY < 0 And crninavrsti = False Then
    'premik belega nazaj
    preverikmeta = False
    Exit Function
    End If
    
If razY > 0 And crninavrsti = True Then
    'premik èrnega naprej
    preverikmeta = False
    Exit Function
    End If

'za koliko se je premaknil?
If Abs(razX) > 1 Then
    'kmet je zavil veè kot eno polje vstran
    preverikmeta = False
    Exit Function
    End If

If Abs(razY) > 2 Then
    'premik naprej/nazaj za veè kot dve polji. Napaka
    preverikmeta = False
    Exit Function
    End If

'torej, kmet je šel v pravo smer za najveè 2 polji in lahko da je zavil.

'najprej poglejmo, ali je šel naprej za dve polji
'to lahko stori samo, èe štarta iz osnovne vrste
If Abs(razY) = 2 Then
    'èe je beli, je moral iz vrste 2
    If crninavrsti = False And Mid(poteza, 2, 1) <> "2" Then
        'šel je iz druge vrste, kot bi smel pri premiku za dve polji!
        preverikmeta = False
        Exit Function
        End If
        
    'kaj pa  èrni?
    If crninavrsti = True And Mid(poteza, 2, 1) <> "7" Then
        'šel je iz druge vrste kot bi smel!
        preverikmeta = False
        Exit Function
        End If
        
    'pravilno je šel dve polji naprej in je zavil?
    If Abs(razX) > 0 Then
        'zavil je. Ne bi smel!
        preverikmeta = False
        Exit Function
        End If
        
    'je vmes kaka figura?
    'pri crnem?
    If crninavrsti = True And Trim(vrniFiguro(polozaji, Left(poteza, 1) & "6")) <> "" Then
        'figura je vmes!
        preverikmeta = False
        Exit Function
        End If
        
    'pri belem
    If crninavrsti = False And Trim(vrniFiguro(polozaji, Left(poteza, 1) & "3")) <> "" Then
        'figura je vmes!
        preverikmeta = False
        Exit Function
        End If
    
    'ali je konèno polje prosto?
    If Trim(vrniFiguro(polozaji, Mid(poteza, 4, 2))) <> "" Then
        'na konènem polju je figura. Napaka!
        preverikmeta = False
        Exit Function
        End If
    
    End If
    
    'tu je moral iti za eno polje in lahko da je zavil
    
    'je na konènem polju kaka figura
    Dim fig As String
    fig = Trim(vrniFiguro(polozaji, Mid(poteza, 4, 2)))
    
    
    If Abs(razX) = 1 Then
        'zavil je!
        If Abs(razY) <> 1 Then
            'žreti hoèe vstran
            preverikmeta = False
            Exit Function
            End If
        
        
        'naskakuje lastno figuro?
        If crninavrsti = False And Left(fig, 1) = "B" Then
            'požreti hoèe svojo figuro!
            preverikmeta = False
            Exit Function
            End If
            
        If crninavrsti = True And Left(fig, 1) = "C" Then
            'požreti hoèe lastno figuro
            preverikmeta = False
            Exit Function
            End If
        'pa je kaj za požreti?
        
        If fig = "" Then
            'ni nièesar!
            preverikmeta = False
            Exit Function
            End If
        
        End If
        
        
    'tu je lahko šel samo še za eno polje naprej. je tam kaka figura?
    If Abs(razX) = 0 And fig <> "" Then
        'tam je figura
        preverikmeta = False
        Exit Function
        End If
        
    
    
    

End Function
