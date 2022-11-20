VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form UpGrade 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "500"
   Begin ComctlLib.ProgressBar BarupgPrg 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "UpGrade.frx":0000
   End
   Begin VB.Label lblupgPrg 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Tag             =   "501"
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblupgdesc 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Tag             =   "502"
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "upGrade"
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
Dim iLoop As Integer, I As Integer
Handle_Err 0, "Form_Load-UpGrade"
LoadFormResourceString Me
BarupgPrg.Min = 0
BarupgPrg.Max = 100
For iLoop = 1 To 100
    BarupgPrg.Value = iLoop
    DoEvents
    For I = 1 To 20000
    Next
Next iLoop
Exit Sub
Form_Err:
    Handle_Err Err, "Form_Load-UpGrade"
    Resume Next
End Sub
