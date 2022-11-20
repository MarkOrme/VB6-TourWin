VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form UpGrades 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Tag             =   "500"
   Begin MSComctlLib.ProgressBar BarupgPrg 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblupgPrg 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Tag             =   "501"
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblupgdesc 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Tag             =   "502"
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "UpGrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
' ------------------------------------------------
' Display to user that database is being upgraded
' ------------------------------------------------
On Local Error GoTo Form_Err
Dim iLoop As Integer, i As Integer
If bDebug Then Handle_Err 0, "Form_Load-UpGrade"
LoadFormResourceString Me
BarupgPrg.Min = 0
BarupgPrg.Max = 100
For iLoop = 1 To 100
    BarupgPrg.Value = iLoop
    DoEvents
    For i = 1 To 20000
    Next
Next iLoop
Exit Sub
Form_Err:
    If bDebug Then Handle_Err Err, "Form_Load-UpGrade"
    Resume Next
End Sub
