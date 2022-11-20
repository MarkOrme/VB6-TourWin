VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form GuideFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peak Performance Guide Names."
   ClientHeight    =   5520
   ClientLeft      =   1515
   ClientTop       =   1650
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   7395
   Begin TabDlg.SSTab GuiPeaTab 
      Height          =   4575
      Left            =   120
      TabIndex        =   41
      Top             =   360
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   8070
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Peak Name"
      TabPicture(0)   =   "GUIDEFRM.frx":0000
      Tab(0).ControlCount=   20
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GuiColCmd(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "GuiColCmd(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "GuiColCmd(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "GuiColCmd(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "GuiColCmd(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "GuiColCmd(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "GuiColCmd(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "GuiColCmd(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "GuiColCmd(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "GuiColCmd(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Peak(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Peak(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Peak(7)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Peak(6)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Peak(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Peak(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Peak(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Peak(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Peak(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Peak(0)"
      Tab(0).Control(19).Enabled=   0   'False
      TabCaption(1)   =   "Peak &Names"
      TabPicture(1)   =   "GUIDEFRM.frx":001C
      Tab(1).ControlCount=   20
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Peak(19)"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "Peak(18)"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "Peak(17)"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "Peak(16)"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "Peak(15)"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "Peak(14)"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "Peak(13)"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "Peak(12)"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "Peak(11)"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Peak(10)"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "GuiColCmd(19)"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "GuiColCmd(18)"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "GuiColCmd(17)"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "GuiColCmd(16)"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "GuiColCmd(15)"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "GuiColCmd(14)"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "GuiColCmd(13)"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "GuiColCmd(12)"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "GuiColCmd(11)"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "GuiColCmd(10)"
      Tab(1).Control(19).Enabled=   -1  'True
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   19
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   38
         Top             =   3960
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   18
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   36
         Top             =   3600
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   17
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   34
         Top             =   3240
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   16
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   32
         Top             =   2880
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   15
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   30
         Top             =   2520
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   14
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   28
         Top             =   2160
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   13
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   26
         Top             =   1800
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   12
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   24
         Top             =   1440
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   11
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   22
         Top             =   1080
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   10
         Left            =   -74760
         MaxLength       =   100
         TabIndex        =   20
         Top             =   720
         Width           =   6015
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   19
         Left            =   -68520
         TabIndex        =   39
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   18
         Left            =   -68520
         TabIndex        =   37
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   17
         Left            =   -68520
         TabIndex        =   35
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   16
         Left            =   -68520
         TabIndex        =   33
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   15
         Left            =   -68520
         TabIndex        =   31
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   14
         Left            =   -68520
         TabIndex        =   29
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   13
         Left            =   -68520
         TabIndex        =   27
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   12
         Left            =   -68520
         TabIndex        =   25
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   11
         Left            =   -68520
         TabIndex        =   23
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   10
         Left            =   -68520
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   100
         TabIndex        =   0
         Top             =   660
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   1
         Left            =   240
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1020
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   2
         Left            =   240
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1380
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   3
         Left            =   240
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1740
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   4
         Left            =   240
         MaxLength       =   100
         TabIndex        =   8
         Top             =   2100
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   5
         Left            =   240
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2460
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   6
         Left            =   240
         MaxLength       =   100
         TabIndex        =   12
         Top             =   2820
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   7
         Left            =   240
         MaxLength       =   100
         TabIndex        =   14
         Top             =   3180
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   8
         Left            =   240
         MaxLength       =   100
         TabIndex        =   16
         Top             =   3540
         Width           =   6015
      End
      Begin VB.TextBox Peak 
         Height          =   285
         Index           =   9
         Left            =   240
         MaxLength       =   100
         TabIndex        =   18
         Top             =   3900
         Width           =   6015
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   0
         Left            =   6480
         TabIndex        =   1
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   1
         Left            =   6480
         TabIndex        =   3
         Top             =   1020
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   2
         Left            =   6480
         TabIndex        =   5
         Top             =   1380
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   3
         Left            =   6480
         TabIndex        =   7
         Top             =   1740
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   4
         Left            =   6480
         TabIndex        =   9
         Top             =   2100
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   5
         Left            =   6480
         TabIndex        =   11
         Top             =   2460
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   6
         Left            =   6480
         TabIndex        =   13
         Top             =   2820
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   7
         Left            =   6480
         TabIndex        =   15
         Top             =   3180
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   8
         Left            =   6480
         TabIndex        =   17
         Top             =   3540
         Width           =   375
      End
      Begin VB.CommandButton GuiColCmd 
         Height          =   285
         Index           =   9
         Left            =   6480
         TabIndex        =   19
         Top             =   3900
         Width           =   375
      End
   End
   Begin VB.CommandButton PeaCloCmd 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   42
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton PeaSavCmd 
      Caption         =   "&Save and exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   40
      Top             =   5040
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog GuiComDia 
      Left            =   240
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Click and edit desired description."
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "GuideFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Changed As Boolean

Sub Define_GuideFrm_mnu(LoadType As String)
On Local Error GoTo Define_GuideFrm_Err

If Not ObjNew.info.NewUser Then
Select Case LoadType
    Case Unloadmnu:
        MDI!MdiEximnu.Caption = MdiFrm_Exit
        MDI!MdiOptmnu.Caption = MdiFrm_Option
        MDI!MdiNewmnu.Caption = MdiFrm_Newmnu
        MDI!MdiSavmnu.Caption = MdiFrm_Savmnu
        MDI!MdiDelmnu.Caption = MdiFrm_Delmnu
        MDI!MdiOptmnu.Enabled = True
        MDI!MdiNewmnu.Enabled = True
        MDI!MdiSavmnu.Enabled = False
        MDI!MdiDelmnu.Enabled = False
        MDI!MdiPrimnu.Enabled = False
        MDI!MdiGuimnu.Enabled = True
    Case Loadmnu:
        MDI!MdiEximnu.Caption = GuideFrm_Exit
        MDI!MdiOptmnu.Caption = GuideFrm_Option
        MDI!MdiSavmnu.Caption = gcGuideFrm_Savmnu
        MDI!MdiNewmnu.Caption = gcGuideFrm_Newmnu
        MDI!MdiDelmnu.Caption = gcGuideFrm_Delmnu
        MDI!MdiOptmnu.Enabled = False
        MDI!MdiGuimnu.Enabled = False
        MDI!MdiNewmnu.Enabled = False
        MDI!MdiSavmnu.Enabled = True
        MDI!MdiDelmnu.Enabled = False
        MDI!MdiGuimnu.Enabled = False
End Select
End If
Exit Sub
Define_GuideFrm_Err:
    If bDebug Then Handle_Err Err, "Define_GuideFrm_mun-GuideFrm"
    Resume Next
End Sub

Private Sub Form_Activate()
    Define_GuideFrm_mnu Loadmnu    ' Changes Mdi menu to in Focus Form
End Sub

Private Sub Form_Deactivate()
Define_GuideFrm_mnu Unloadmnu
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Me.PeaCloCmd_Click
End Sub

Private Sub Form_Load()
On Local Error GoTo Guide_Err

Dim FieldName As String, RetStr As String, I As Integer
Dim SQL As String

Screen.MousePointer = 11
CentreForm GuideFrm, -1
Me.KeyPreview = True
Define_GuideFrm_mnu Loadmnu         ' Changes Mdi menu to in Focus Form
    'Get Peak description
    
' =======================================
' Load values into memory if not already.
' =======================================
If Not ObjChart.info.IsValuesLoaded Then
    ObjChart.LoadChartColoursAndNames
End If

For I = 0 To ObjChart.UpBoundColour(gcPeakChart)

    ProgressBar "Loading Peak Names...", -1, I * 1.111, -1
    ' Get Peak# Field value
    ' ---------------------
    RetStr = ObjChart.GetName(I, gcPeakChart)
    
    If RetStr <> "No Return" Or RetStr <> "" Then Peak(I) = RetStr
    ' Get Color# Field value
    ' ---------------------
    RetStr = ObjChart.GetColour(I, gcPeakChart)
    If RetStr <> "No Return" Then Peak(I).BackColor = RetStr
    
Next I

'Visible = false for progress bar
ProgressBar "", 0, 0, 0
Screen.MousePointer = 0
Changed = False
Exit Sub
Guide_Err:
    If bDebug Then Handle_Err Err, "Load-GuideFrm"
    Screen.MousePointer = 0
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Define_GuideFrm_mnu Unloadmnu    ' Changes Mdi menu to in Focus Form
End Sub

Private Sub GuiColCmd_Click(Index As Integer)
On Local Error GoTo GuiColCmd_Err
GuiComDia.Action = 3
Peak(Index).BackColor = GuiComDia.Color
Exit Sub
GuiColCmd_Err:
    If Err = 32755 Then  'User Canceled
        Err.Clear
        Exit Sub
    Else
        If bDebug Then Handle_Err Err, "GuiColCmd_Click-GuideFrm"
        Resume Next
    End If
End Sub

Sub PeaCloCmd_Click()
On Local Error GoTo PeaCloCmd_Err
If Changed Then
    If vbYes = MsgBox("Do you wish to save changes?", vbYesNo + vbQuestion, LoadResString(gcTourVersion)) Then
        PeaSavCmd_Click
        Exit Sub
    Else
        Unload GuideFrm
    End If
End If
Unload GuideFrm
Exit Sub
PeaCloCmd_Err:
    If bDebug Then Handle_Err Err, "PeaCloCmd_Click-GuideFrm"
    Resume Next
End Sub

Private Sub Peak_Change(Index As Integer)
Changed = True
End Sub

Sub PeaSavCmd_Click()
    
On Local Error GoTo PeakSave_Err

Dim I As Integer

' ------------------------
' set Default return value
' ------------------------
    
    For I = ObjChart.LowBoundColour To ObjChart.UpBoundColour(gcPeakChart)
    
        If Peak(I).Text <> ObjChart.GetName(I, gcPeakChart) Then ObjChart.SetName I, Peak(I).Text, gcPeakChart
        If Peak(I).BackColor <> ObjChart.GetColour(I, gcPeakChart) Then ObjChart.SetColour I, Peak(I).BackColor, gcPeakChart
                  
    Next I

ObjChart.SaveChartColoursAndNames gcPeakChart

Unload GuideFrm
Exit Sub
PeakSave_Err:
    If bDebug Then Handle_Err Err, "PeaSavCmd-GuideFrm"
    Resume Next
End Sub


