VERSION 5.00
Begin VB.Form frmExportOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "####"
   ClientHeight    =   3015
   ClientLeft      =   4050
   ClientTop       =   2610
   ClientWidth     =   4155
   Icon            =   "frmExportOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "575"
   Begin VB.CheckBox chkOverWrite 
      Caption         =   "####"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Tag             =   "587"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame frmList 
      Caption         =   "####"
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Tag             =   "858"
      Top             =   240
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ListBox lstItems 
         Height          =   1425
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblItems 
         Caption         =   "####"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Tag             =   "586"
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "####"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Tag             =   "578"
      Top             =   240
      Width           =   3855
      Begin VB.OptionButton OptType 
         Caption         =   "####"
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Tag             =   "581"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.OptionButton OptType 
         Caption         =   "####"
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Tag             =   "580"
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton OptType 
         Caption         =   "####"
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Tag             =   "579"
         Top             =   360
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "####"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Tag             =   "577"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "####"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Tag             =   "576"
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "frmExportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim iLoop As Long
On Local Error GoTo OK_Err
' Show hourglass
Screen.MousePointer = vbHourglass

' Export options are displayed
If LoadResString(575) = Me.Caption Then

    If OptType(gcEVENT_DAYS).Value Then
        Me.Caption = LoadResString(582)
        cPeakSchedules.GetListOfItems gcEVENT_DAYS, lstItems
        
    ElseIf OptType(gcPEAK_SCHEDULE).Value Then
        Me.Caption = LoadResString(583)
        cPeakSchedules.GetListOfItems gcPEAK_SCHEDULE, lstItems
        
    Else ' Must be Daily Activities
        Me.Caption = LoadResString(584)
        cPeakSchedules.GetListOfItems gcDAILY_ACTIVITIES, lstItems
    
    End If
' =====================
' Hidden Option Frame
' and show ListBox...
' =====================
fraOptions.Visible = False

Me.frmList.Visible = True
Me.chkOverWrite.Visible = True

Else
    If cPeakSchedules.GetFileAttributes(IIf(Me.chkOverWrite.Value = 0, False, True)) Then
    For iLoop = 0 To lstItems.ListCount - 1

        If lstItems.Selected(iLoop) Then
            lstItems.ListIndex = iLoop
            cPeakSchedules.Export gcActive_Type_PeakNames, lstItems.Text
        End If
    Next iLoop
        Unload Me
        cPeakSchedules.DoExport gcActive_Type_PeakNames
    End If

End If
' Restore mouse
Screen.MousePointer = vbDefault
Exit Sub
OK_Err:
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
' Load Resource strings
LoadFormResourceString Me
' Remove the following code to restore
' the options to export, daily, peak and calendar...

Me.Caption = LoadResString(583)
cPeakSchedules.GetListOfItems gcPEAK_SCHEDULE, lstItems

fraOptions.Visible = False
Me.frmList.Visible = True
Me.chkOverWrite.Visible = True

End Sub
