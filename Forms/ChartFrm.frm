VERSION 5.00
Begin VB.Form ChartFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Chart"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNamCht 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtColCht 
      Height          =   285
      Index           =   0
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   270
   End
End
Attribute VB_Name = "ChartFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oChart As cChart
Attribute oChart.VB_VarHelpID = -1

Private Sub Form_Load()
Set oChart = ObjChart

DisplayChartColours oChart.info.CurrentChartType
End Sub

Private Sub Form_Unload(Cancel As Integer)
oChart.info.IsFormLoaded = False
End Sub

Function DisplayChartColours(sChartType As String) As Boolean

Dim iLoop As Integer

If Not oChart.info.IsFormLoaded Then

 For iLoop = oChart.LowBoundColour + 1 To oChart.UpBoundColour
    ' Load Colour controls
    With txtColCht(iLoop - 1)
           Load txtColCht(iLoop)
              txtColCht(iLoop).Top = .Top + .Height + 30
              txtColCht(iLoop).Visible = True
    End With
    ' Load Colour Controls
    With txtNamCht(iLoop - 1)
           Load txtNamCht(iLoop)
              txtNamCht(iLoop).Top = .Top + .Height + 30
              txtNamCht(iLoop).Visible = True
    End With
 Next iLoop
End If

For iLoop = oChart.LowBoundColour To oChart.UpBoundColour
    txtColCht(iLoop).BackColor = oChart.GetColour(iLoop, oChart.info.CurrentChartType)
    txtNamCht(iLoop).Text = oChart.GetName(iLoop, oChart.info.CurrentChartType)
Next iLoop

oChart.info.IsFormLoaded = True
ChartFrm.Caption = sChartType & LoadResString(800)

End Function

Private Sub oChart_Modified()
'Update GUI with modified Chart names
DisplayChartColours oChart.info.CurrentChartType
End Sub
