VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Tag             =   "7"
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4320
      Width           =   1215
   End
   Begin Project1.SimpleGrid SimpleGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   5895
      _ExtentX        =   10186
      _ExtentY        =   6588
      CellHeight      =   225
      CellWidth       =   0
      FormatStyle     =   1
      HeaderCaption   =   "uuuuu"
      CellWidth       =   0
      CellHeight      =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oParser     As New CParser

Private Sub Form_Unload(Cancel As Integer)
Set oParser = Nothing
End Sub

Private Sub SimpleGrid1_FetchRows()
On Local Error GoTo FetchRows_Err

Call SimpleGrid1.AddRow("0", "<0>Zero-Contact<0/><1>Ormesher<1/><2>Leanda<2/><3>my@.com<3/>")
Call SimpleGrid1.AddRow("1", "<0>Zero-The Contact<0/><1>Monro<1/><2>Mark<2/><3>your@.com<3/>")


On Local Error GoTo 0
Exit Sub

FetchRows_Err:
Resume Next

End Sub
