VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   2025
   ClientTop       =   1590
   ClientWidth     =   6690
   Height          =   4545
   Left            =   1965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6690
   Top             =   1245
   Width           =   6810
   Begin VB.PictureBox GraTo_Txt 
      Height          =   2500
      Left            =   360
      ScaleHeight     =   2445
      ScaleWidth      =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   4500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
GraTo_Txt.Cls
GraTo_Txt.Line (104, 110)-(196, 110)
End Sub


Private Sub Form_Load()
GraTo_Txt.ScaleLeft = 100
GraTo_Txt.ScaleTop = 200
GraTo_Txt.ScaleWidth = 100
GraTo_Txt.ScaleHeight = -100

End Sub


