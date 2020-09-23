VERSION 5.00
Begin VB.Form ura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clock"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4980
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Black:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "White:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
   Begin VB.Label uracrni 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label uraBeli 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "ura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim belcas As cas
Dim crncas As cas

Sub startajUro()
Timer1.Enabled = True
End Sub



Private Sub Timer1_Timer()
'kateremu se štejejo sekunde?
If crnipripravljen = True Then
    'prištejemo sekundo
    With crncas
    .sekunde = .sekunde + 1
    If .sekunde = 60 Then
        .sekunde = 0
        .minute = .minute + 1
        End If
    If .minute = 60 Then
        .minute = 0
        .ure = .ure + 1
        End If
    End With
    
    Else
    With belcas
    .sekunde = .sekunde + 1
    If .sekunde = 60 Then
        .sekunde = 0
        .minute = .minute + 1
        End If
    If .minute = 60 Then
        .minute = 0
        .ure = .ure + 1
        End If
    End With
    End If
    
'izpišemo oba èasa
uraBeli.Caption = formatirajCas(belcas.ure, belcas.minute, belcas.sekunde)
uracrni.Caption = formatirajCas(crncas.ure, crncas.minute, crncas.sekunde)
End Sub

Function formatirajCas(ByVal ur As Byte, ByVal minut As Byte, ByVal sekund As Byte) As String
If ur < 10 Then
    formatirajCas = formatirajCas & "0" & ur & ":"
    Else
    formatirajCas = formatirajCas & ur & ":"
    End If
    
If minut < 10 Then
    formatirajCas = formatirajCas & "0" & minut & ":"
    Else
    formatirajCas = formatirajCas & minut & ":"
    End If

If sekund < 10 Then
    formatirajCas = formatirajCas & "0" & sekund
    Else
    formatirajCas = formatirajCas & sekund
    End If
End Function

Sub ustaviUro()
Timer1.Enabled = False
End Sub

Sub resetirajUro()
uraBeli.Caption = "00:00:00"
uracrni.Caption = "00:00:00"
crncas.minute = 0
crncas.ure = 0
crncas.sekunde = 0
belcas.minute = 0
belcas.sekunde = 0
belcas.ure = 0
End Sub
