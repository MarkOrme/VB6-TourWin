VERSION 5.00
Begin VB.Form PeakFrm 
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   1470
   ClientTop       =   1890
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   6720
   Begin VB.Menu PeaFilmnu 
      Caption         =   "&File"
      Begin VB.Menu PeaOptmnu 
         Caption         =   "&Options, Peak...."
      End
      Begin VB.Menu PeaSepmnu 
         Caption         =   "-"
      End
      Begin VB.Menu PeaEximnu 
         Caption         =   "E&xit Peak..."
      End
   End
End
Attribute VB_Name = "PeakFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSRibbon1_Click(Value As Integer)

End Sub


