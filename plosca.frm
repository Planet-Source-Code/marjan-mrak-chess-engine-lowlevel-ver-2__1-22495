VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   960
      Left            =   5520
      Picture         =   "plosca.frx":0000
      Top             =   5160
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   6360
      Left            =   120
      Picture         =   "plosca.frx":164A
      Top             =   480
      Width           =   6360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Image1.Top = Y
Image1.Left = X
End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
Image2.Top = Y
Image2.Left = X
End Sub

Private Sub Image2_DragDrop(Source As Control, X As Single, Y As Single)
Image2.Top = Image1.Height - Y - Image1.Top - Image2.Height / 2
Image2.Left = Image1.Width - X - Image1.Left

End Sub
