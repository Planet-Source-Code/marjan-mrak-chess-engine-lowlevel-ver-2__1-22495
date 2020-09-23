VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form sahovnica 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chess board"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8325
   Begin PicClip.PictureClip figb 
      Left            =   2520
      Top             =   1920
      _ExtentX        =   6059
      _ExtentY        =   2249
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "Form1.frx":0000
   End
   Begin PicClip.PictureClip figc 
      Left            =   2760
      Top             =   3480
      _ExtentX        =   6059
      _ExtentY        =   2249
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "Form1.frx":515A
   End
   Begin VB.PictureBox Msahovnica 
      AutoSize        =   -1  'True
      Height          =   6420
      Left            =   240
      Picture         =   "Form1.frx":A2B4
      ScaleHeight     =   6360
      ScaleWidth      =   6360
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "White moves"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Image figurac 
      Height          =   855
      Left            =   480
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image figurab 
      Height          =   975
      Left            =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "sahovnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents mot As Chess.Engine
Attribute mot.VB_VarHelpID = -1



Private Sub Form_Click()
mot.Postavi
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim koda As Integer

If KeyAscii = 0 Then GoTo izvrsipotezo

If crniigra = True Then
    'èrni je na vrsti!
    MsgBox "Not your move!", vbCritical, "Opozorilo"
    Exit Sub
    End If
    
If KeyAscii = 8 Then
        If Label1.Caption = "" Then Exit Sub
                      
        If Len(Label1.Caption) = 3 Then
        Label1.Caption = Left(Label1.Caption, 1)
        Exit Sub
        End If
        
    Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 1)
    Exit Sub
    End If
    

If KeyAscii = 0 Then GoTo izvrsipotezo
If KeyAscii < 32 Then Exit Sub

'prvi in tretji znak MORA biti èrka!

If Len(Label1.Caption) = 0 Or Len(Label1.Caption) = 3 Then
    koda = Asc(UCase(Chr(KeyAscii)))
    If koda < 65 Or koda > 72 Then Exit Sub
    End If
    
If Len(Label1.Caption) = 1 Or Len(Label1.Caption) = 4 Then
    If KeyAscii < 49 Or KeyAscii > 56 Then Exit Sub
    End If



'sprejetje poteze!
izvrsipotezo:
Dim prav As Chess.Parser
Set prav = CreateObject("Chess.Parser")
Dim rez As String

Label1.Caption = Label1.Caption & UCase(Chr(KeyAscii))


If Len(Label1.Caption) = 2 Then
    prav.BlackTurn = crnipripravljen
    prav.Move = Label1.Caption
    rez = sestavipolozaje
    prav.Parse (rez)
    
    If prav.ErrorNumber > 0 Then
            rez = prav.ErrorText
            MsgBox "Error:" & rez, vbCritical, "ERROR!!!"
            Label1.Caption = ""
            GoTo endsub
            End If
        
    Label1.Caption = Label1.Caption & "-"
    End If

If Len(Label1.Caption) = 5 Then
    'poteza! je bil bel?
        'najprej jo je treba sparsati in potem prikazati...
    
        
        prav.BlackTurn = crnipripravljen
        prav.Move = Label1.Caption
        Dim linija As Integer, vrsta As Integer
        
        rez = sestavipolozaje
        prav.BlackCanCastleLeft = crnarosadamogocalevo
        prav.WhiteCanCastleLeft = belarosadamogocalevo
        prav.BlackCanCastleRight = crnarosadamogocadesno
        prav.WhiteCanCastleRight = belarosadamogocadesno
        prav.Parse (rez)
             
        
        
        If prav.ErrorNumber > 0 Then
            rez = prav.ErrorText
            MsgBox "ERROR:" & rez, vbCritical, "ERROR!!!"
            Label1.Caption = ""
            GoTo endsub
            End If
        
        
        
        'opravimo premik
        If zapStPot < 2 Then ura.startajUro
        
        crnarosadamogocadesno = prav.BlackCanCastleRight
        belarosadamogocadesno = prav.WhiteCanCastleRight
        crnarosadamogocalevo = prav.BlackCanCastleLeft
        belarosadamogocalevo = prav.WhiteCanCastleLeft
        
        
        movepiece (Label1.Caption)
        DoEvents
        minimovepiece (Label1.Caption)
        DoEvents
        If prav.WhiteCastled = True Then
            'kam je bila rošada(katero trdnjavo premaknemo)
            If Right(Label1.Caption, 2) = "G1" Then
                movepiece "H1-F1"
                minimovepiece "H1-F1"
                Else
                movepiece "A1-D1"
                minimovepiece "A1-D1"
                End If
            End If
            
        If prav.BlackCastled = True Then
            'kam je bila rošada(katero trdnjavo premaknemo)
            If Right(Label1.Caption, 2) = "G8" Then
                movepiece "H8-F8"
                minimovepiece "H8-F8"
                Else
                movepiece "A8-D8"
                minimovepiece "A8-D8"
                End If
            End If
        
        If prav.Capture = True Then statusPoteze = "X"
        If prav.Check = True Then statusPoteze = statusPoteze & "+"
        
        If crnipripravljen = False Then
            crnipripravljen = True
            Poteze.Seznam = Poteze.Seznam & zapStPot & ". " & Label1.Caption & statusPoteze
            Label1.Caption = ""
            Label2.Caption = "Balck moves"
            GoTo endsub
            Else
            crnipripravljen = False
            Poteze.Seznam = Poteze.Seznam & " ... " & Label1.Caption & statusPoteze & vbCrLf
            Label1.Caption = ""
            Label2.Caption = "White moves"
            zapStPot = zapStPot + 1
            End If
End If
            
endsub:
Set prav = Nothing
If crnipripravljen = True Then
crniigra = True
Crnimisli
End If
End Sub

Private Sub Form_Load()
crnipripravljen = False
zapStPot = 1
pozrtih(0) = 0
pozrtih(1) = 0
End Sub

Sub Crnimisli()
Set mot = CreateObject("Chess.engine")
statusPoteze = ""
mot.BlackCanCastleLeft = crnarosadamogocalevo
mot.BlackCanCastleRight = crnarosadamogocadesno
mot.WhiteCanCastleLeft = belarosadamogocalevo
mot.WhiteCanCastleRight = belarosadamogocadesno
mot.Levels = Nastavitve.stNivojev.Value
mot.StyleOfPlay = Nastavitve.obrNap.Value

mot.think sestavipolozaje(), zapStPot
End Sub

Private Sub mot_FoundMove(ByVal cpoteza As String, ByVal statuspo As String)
If cpoteza = "00-00" Then
    ura.ustaviUro
    Exit Sub  'konec igre!
    End If
    
crniigra = False
Label1.Caption = UCase(cpoteza)
statusPoteze = statuspo
Form_KeyPress (0)
End Sub

Private Sub mot_DrawMove(ByVal cpoteza As String)
Dim sfig As String
Dim cfig As String

minimovepiece cpoteza
End Sub

Private Sub mot_drawcastle(ByVal poteza As String)
If poteza = "E8-G8" Then
    minimovepiece (poteza)
    minimovepiece ("H8-F8")
    End If

If poteza = "E1-G1" Then
    minimovepiece (poteza)
    minimovepiece ("H1-F1")
    End If

If poteza = "E1-C1" Then
    minimovepiece (poteza)
    minimovepiece ("A1-D1")
    End If

If poteza = "E8-C8" Then
    minimovepiece (poteza)
    minimovepiece ("A8-D8")
    End If

'MsgBox "moja ideja!"

If poteza = "E8-G8" Then
    minimovepiece ("G8-E8")
    minimovepiece ("F8-H8")
    End If

If poteza = "E1-G1" Then
    minimovepiece ("G1-E1")
    minimovepiece ("F1-H1")
    End If

If poteza = "E1-C1" Then
    minimovepiece ("C1-E1")
    minimovepiece ("D1-A1")
    End If

If poteza = "E8-C8" Then
    minimovepiece ("C8-E8")
    minimovepiece ("D8-A8")
    End If

End Sub

Private Sub mot_RestoreBoard(ByVal polozajin As String)
'potrebno restavrirati mini plošèo na poteze
Dim n As Integer, m As Integer

minidisplayEmptyBoard   'ozris prazne plošèe

'napolnimo jo!
For n = 1 To 8
For m = 1 To 8
   If polozajin = "" Then
    minidrawPiece polozaji(n, m), vrniZnakLinije(n) & vrniZnakVrste(m)
    minipolozaji(n, m) = polozaji(n, m)
    Else
    minidrawPiece Mid(polozajin, 4, 2), Left(polozajin, 2)
    minipolozaji(m, n) = Mid(polozajin, 4, 2)
    polozajin = Mid(polozajin, 7)
    End If
   Next
   Next
sahovnica.SetFocus
End Sub
