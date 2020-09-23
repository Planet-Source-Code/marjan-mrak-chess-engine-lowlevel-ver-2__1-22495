VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sahovnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Sub DisplayCompleteBoard()
Form1.PaintPicture Form1.Msahovnica.Picture, 0, 0
'beli kmetje
drawPiece "BP", "A2"
drawPiece "BP", "B2"
drawPiece "BP", "c2"
drawPiece "BP", "d2"
drawPiece "BP", "e2"
drawPiece "BP", "f2"
drawPiece "BP", "g2"
drawPiece "BP", "h2"

'èrni kmetje
drawPiece "cP", "A7"
drawPiece "cP", "B7"
drawPiece "cP", "c7"
drawPiece "cP", "d7"
drawPiece "cP", "e7"
drawPiece "cP", "f7"
drawPiece "cP", "g7"
drawPiece "cP", "h7"

'beli trdnjavi
drawPiece "bt", "A1"
drawPiece "bt", "h1"

'èrni trdnjavi
drawPiece "ct", "A8"
drawPiece "ct", "h8"

'bela konja
drawPiece "bs", "b1"
drawPiece "bs", "g1"

'èrna konja
drawPiece "cs", "b8"
drawPiece "cs", "g8"

'bela lovca
drawPiece "bl", "c1"
drawPiece "bl", "f1"

'crna lovca
drawPiece "cl", "c8"
drawPiece "cl", "f8"


'bela kraljica
drawPiece "bq", "d1"

'crna kraljica
drawPiece "cq", "d8"


'beli kralj
drawPiece "bk", "e1"

'èrni kralj
drawPiece "ck", "e8"

Form1.Show

End Sub


Public Sub displayEmptyBoard()
Form1.Show
Form1.PaintPicture Sahovnica.Picture, 0, 0
End Sub

