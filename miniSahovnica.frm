VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form miniSahovnica 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PC thinking"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Msahovnica 
      AutoSize        =   -1  'True
      Height          =   2790
      Left            =   240
      Picture         =   "miniSahovnica.frx":0000
      ScaleHeight     =   2730
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2805
   End
   Begin PicClip.PictureClip figb 
      Left            =   3960
      Top             =   2040
      _ExtentX        =   3651
      _ExtentY        =   1244
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "miniSahovnica.frx":115A
   End
   Begin PicClip.PictureClip figc 
      Left            =   2760
      Top             =   3600
      _ExtentX        =   3651
      _ExtentY        =   1244
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "miniSahovnica.frx":1560
   End
   Begin VB.Image figurab 
      Height          =   975
      Left            =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image figurac 
      Height          =   855
      Left            =   480
      Top             =   120
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   2
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Na potezi je beli"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
End
Attribute VB_Name = "miniSahovnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim minizapStPot As Integer


Private Sub Form_Load()
crnipripravljen = False
minizapStPot = 1
minipozrtih(0) = 0
minipozrtih(1) = 0
End Sub


