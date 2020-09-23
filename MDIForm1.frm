VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Chess by MAN Soft"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu game_new 
         Caption         =   "New"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu game_over 
         Caption         =   "End"
      End
   End
   Begin VB.Menu setings 
      Caption         =   "Settings"
   End
   Begin VB.Menu view 
      Caption         =   "VIew"
      Begin VB.Menu view_pc 
         Caption         =   "PC"
         Checked         =   -1  'True
      End
      Begin VB.Menu view_moves 
         Caption         =   "Moves"
         Checked         =   -1  'True
      End
      Begin VB.Menu view_captured 
         Caption         =   "Captured"
         Checked         =   -1  'True
      End
      Begin VB.Menu view_clock 
         Caption         =   "Clock"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub game_new_Click()
pozrtih(0) = 0
zapStPot = 1
pozrtih(1) = 0
pozrtih(2) = 0
DisplayCompleteBoard
displaymoves
displaycaptured
miniDisplayCompleteBoard
displayWatch
sahovnica.SetFocus
If view_pc.Checked = False Then miniSahovnica.Hide
If view_moves.Checked = False Then Poteze.Hide
If view_clock.Checked = False Then ura.Hide
If view_captured.Checked = False Then pozrte.Hide
sahovnica.SetFocus
Poteze.Seznam.Text = ""
pozrte.Cls
ura.ustaviUro
ura.resetirajUro
crniigra = False

End Sub

Private Sub game_over_Click()
Open App.Path & "\sah.ini" For Output As #1
'najprej polozaji oken
Print #1, "Sahovnica:" & sahovnica.Left & "," & sahovnica.Top
Print #1, "PC:" & miniSahovnica.Left & "," & miniSahovnica.Top & "," & view_pc.Checked
Print #1, "Poteze:" & Poteze.Left & "," & Poteze.Top & "," & view_moves.Checked
Print #1, "Ura:" & ura.Left & "," & ura.Top & "," & view_clock.Checked
Print #1, "Zajete:" & pozrte.Left & "," & pozrte.Top & "," & view_captured.Checked
Print #1, "Število nivojev:" & Nastavitve.stNivojev.Value
Print #1, "Naèin igre:" & Nastavitve.obrNap.Value
Close
End
End Sub

Private Sub MDIForm_Load()
Dim vrstica As String
Dim lahkoPokaze As Boolean

Open App.Path & "\sah.ini" For Input As #1
Line Input #1, vrstica
preberiVrstico vrstica, ploscax, ploscay, lahkoPokaze

Line Input #1, vrstica
preberiVrstico vrstica, pcx, pcy, lahkoPokaze
view_pc.Checked = lahkoPokaze

Line Input #1, vrstica
preberiVrstico vrstica, potx, poty, lahkoPokaze
view_moves.Checked = lahkoPokaze

Line Input #1, vrstica
preberiVrstico vrstica, urax, uray, lahkoPokaze
view_clock.Checked = lahkoPokaze


Line Input #1, vrstica
preberiVrstico vrstica, pozx, pozy, lahkoPokaze
view_captured.Checked = lahkoPokaze

Line Input #1, vrstica
Nastavitve.stNivojev.Value = CInt(Mid(vrstica, 17))

Line Input #1, vrstica
Nastavitve.obrNap.Value = CInt(Mid(vrstica, 12))

Close


DisplayCompleteBoard
displaymoves
displaycaptured
miniDisplayCompleteBoard
displayWatch
sahovnica.SetFocus
If view_pc.Checked = False Then miniSahovnica.Hide
If view_moves.Checked = False Then Poteze.Hide
If view_clock.Checked = False Then ura.Hide
If view_captured.Checked = False Then pozrte.Hide
End Sub

Private Sub setings_Click()
Nastavitve.Show 1
End Sub

Private Sub view_captured_Click()
If view_captured.Checked = True Then
    pozrte.Hide
    view_captured.Checked = False
    sahovnica.SetFocus
    Else
    pozrte.Show
    view_captured.Checked = True
    sahovnica.SetFocus
    End If
End Sub

Private Sub view_clock_Click()
If view_clock.Checked = True Then
    ura.Hide
    view_clock.Checked = False
    sahovnica.SetFocus
    Else
    ura.Show
    view_clock.Checked = True
    sahovnica.SetFocus
    End If
End Sub

Private Sub view_moves_Click()
If view_moves.Checked = True Then
    Poteze.Hide
    view_moves.Checked = False
    sahovnica.SetFocus
    Else
    Poteze.Show
    view_moves.Checked = True
    sahovnica.SetFocus
    End If

End Sub

Private Sub view_pc_Click()
If view_pc.Checked = True Then
    miniSahovnica.Hide
    view_pc.Checked = False
    sahovnica.SetFocus
    Else
    miniSahovnica.Show
    view_pc.Checked = True
    sahovnica.SetFocus
    End If
End Sub
