VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "TourWin: Serious Software for Serious Athletes."
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   HelpContextID   =   3
   Icon            =   "mditour.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.TextBox MdiBarTxt 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7800
         TabIndex        =   4
         Top             =   50
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox ProgressBack 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   50
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4950
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11165
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer Mdi_Time 
      Interval        =   3
      Left            =   6960
      Top             =   1800
   End
   Begin VB.PictureBox ProgressBar 
      Align           =   3  'Align Left
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   4320
      Left            =   0
      ScaleHeight     =   4260
      ScaleWidth      =   660
      TabIndex        =   0
      Top             =   630
      Width           =   720
      Begin VB.Image MdiNicDn 
         Height          =   480
         Left            =   120
         Picture         =   "mditour.frx":0442
         Top             =   3000
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiNicUp 
         Height          =   480
         Left            =   120
         Picture         =   "mditour.frx":074C
         Top             =   3960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiNicBut 
         Height          =   495
         Left            =   90
         ToolTipText     =   "Contacts Module"
         Top             =   2490
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image MdiConUp 
         Height          =   480
         Left            =   0
         Picture         =   "mditour.frx":0A56
         Top             =   3360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiConDn 
         Height          =   480
         Left            =   0
         Picture         =   "mditour.frx":0D60
         Top             =   3360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiConBut 
         Height          =   495
         Left            =   90
         ToolTipText     =   "Contacts Module"
         Top             =   1890
         Width           =   495
      End
      Begin VB.Image MdiEveDn 
         Height          =   480
         Left            =   120
         Picture         =   "mditour.frx":106A
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiEveUp 
         Height          =   480
         Left            =   240
         Picture         =   "mditour.frx":1374
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiEveBut 
         Height          =   495
         Left            =   90
         ToolTipText     =   "Calendar Module"
         Top             =   1290
         Width           =   495
      End
      Begin VB.Image MdiDayBut 
         Height          =   480
         Left            =   90
         ToolTipText     =   "Daily Module"
         Top             =   90
         Width           =   480
      End
      Begin VB.Image MdiDayDn 
         Height          =   480
         Left            =   120
         Picture         =   "mditour.frx":167E
         Top             =   3600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiDayUp 
         Height          =   480
         Left            =   120
         Picture         =   "mditour.frx":1988
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiGraDn 
         Height          =   480
         Left            =   240
         Picture         =   "mditour.frx":1C92
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MdiGraBut 
         Height          =   495
         Left            =   90
         ToolTipText     =   "Graph Module"
         Top             =   660
         Width           =   495
      End
      Begin VB.Image MdiGraUp 
         Height          =   480
         Left            =   120
         Picture         =   "mditour.frx":1F9C
         Top             =   3840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image MidNotBmp 
         Height          =   330
         Left            =   -840
         Picture         =   "mditour.frx":22A6
         Top             =   0
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   6960
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":2430
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":2982
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":2ED4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":3426
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":3978
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":3ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":441C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":496E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":4EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":5412
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":5964
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":5EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mditour.frx":6408
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MidFilmnu 
      Caption         =   "&File"
      Begin VB.Menu MdiOptmnu 
         Caption         =   "&Option, StartUp..."
      End
      Begin VB.Menu Mdisepmnu 
         Caption         =   "-"
      End
      Begin VB.Menu MdiPStmnu 
         Caption         =   "Print &SetUp"
      End
      Begin VB.Menu MdiPrimnu 
         Caption         =   "&Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu MdiSe4mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mdiImportmnu 
         Caption         =   "&Import..."
      End
      Begin VB.Menu mdiExportmnu 
         Caption         =   "&Export..."
      End
      Begin VB.Menu MdiEximnu 
         Caption         =   "E&xit TourWin..."
      End
   End
   Begin VB.Menu MdiEdimnu 
      Caption         =   "&Edit"
      Begin VB.Menu MdiCutmnu 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu MdiCopymnu 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu MdiPastemnu 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu MdiDeletemnu 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu Mdisep9mnu 
         Caption         =   "-"
      End
      Begin VB.Menu MdiNewmnu 
         Caption         =   "&New User"
      End
      Begin VB.Menu MdiSavmnu 
         Caption         =   "&Save User"
         Enabled         =   0   'False
      End
      Begin VB.Menu MdiDelmnu 
         Caption         =   "&Delete User"
      End
   End
   Begin VB.Menu MdiDbmnu 
      Caption         =   "&View"
      Begin VB.Menu MdiDiamnu 
         Caption         =   "Da&ily..."
      End
      Begin VB.Menu MdiCalmnu 
         Caption         =   "&Calendar..."
      End
      Begin VB.Menu Mdinicmnu 
         Caption         =   "Concon&i..."
      End
      Begin VB.Menu MdiCntmnu 
         Caption         =   "C&ontacts..."
      End
      Begin VB.Menu MDIWizMnu 
         Caption         =   "Setup Wizard..."
      End
   End
   Begin VB.Menu MdiToomnu 
      Caption         =   "&Tools"
      Begin VB.Menu MdiRepmnu 
         Caption         =   "&Reports"
         Begin VB.Menu MdiDarmnu 
            Caption         =   "&Daily Report"
         End
         Begin VB.Menu MdiWeemnu 
            Caption         =   "&Weekly Report"
            Begin VB.Menu MdiTotmnu 
               Caption         =   "Weekly ( T&otals Report )"
            End
            Begin VB.Menu MdiPermnu 
               Caption         =   "Weekly ( &Percentage )"
            End
         End
         Begin VB.Menu MdiCndmnu 
            Caption         =   "&Calendar"
            Begin VB.Menu MdiCdrmnu 
               Caption         =   "&Daily Calendar Report"
            End
            Begin VB.Menu MdiCermnu 
               Caption         =   "&Event Calendar Report"
            End
            Begin VB.Menu MdiCprmnu 
               Caption         =   "&Peak Calendar Report"
            End
         End
         Begin VB.Menu ConiRepmnu 
            Caption         =   "Concon&i Report"
         End
         Begin VB.Menu MdiCnlmnu 
            Caption         =   "Contact &Listing"
         End
         Begin VB.Menu mnuRepHis 
            Caption         =   "&Historical Events"
         End
      End
      Begin VB.Menu MdiGramnu 
         Caption         =   "&Graph..."
      End
   End
   Begin VB.Menu MdiSetmnu 
      Caption         =   "&SetUp"
      Begin VB.Menu MdiSysmnu 
         Caption         =   "System &Names"
         Begin VB.Menu MdiGuimnu 
            Caption         =   "&Peak Guide..."
         End
         Begin VB.Menu MdiTypmnu 
            Caption         =   "&Event Type..."
         End
         Begin VB.Menu MdiHeamnu 
            Caption         =   "&Heart Rate..."
         End
      End
      Begin VB.Menu MdiSysSch 
         Caption         =   "System &Schedules"
         Begin VB.Menu MdiSysHea 
            Caption         =   "&Heart Rate..."
         End
         Begin VB.Menu MdiCycmnu 
            Caption         =   "&Peak..."
         End
      End
   End
   Begin VB.Menu MdiHlpmnu 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu MdiConmnu 
         Caption         =   "&Contents"
      End
      Begin VB.Menu MdiSeamnu 
         Caption         =   "&Search for help on..."
      End
      Begin VB.Menu MdiHspmnu 
         Caption         =   "-"
      End
      Begin VB.Menu MdiTecmnu 
         Caption         =   "&Technical Support"
      End
      Begin VB.Menu MdiSe3mnu 
         Caption         =   "-"
      End
      Begin VB.Menu MdiAbomnu 
         Caption         =   "&About TourWin..."
      End
   End
   Begin VB.Menu DaiLevmnu 
      Caption         =   "Level"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu DaiLv1mnu 
         Caption         =   "Level - &1"
      End
      Begin VB.Menu DaiLv2mnu 
         Caption         =   "Level - &2 "
      End
      Begin VB.Menu DaiLv3mnu 
         Caption         =   "Level - &3 "
      End
      Begin VB.Menu DaiLv4mnu 
         Caption         =   "Level - &4 "
      End
      Begin VB.Menu DaiLv5mnu 
         Caption         =   "Level - &5"
      End
      Begin VB.Menu DaiShowmnu 
         Caption         =   "&Show all Levels"
      End
   End
   Begin VB.Menu MdiBegmnu 
      Caption         =   "&Begin Report"
      Visible         =   0   'False
   End
   Begin VB.Menu CalPopmnu 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu CalPoemnu 
         Caption         =   "Edit"
      End
      Begin VB.Menu CalPodmnu 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim started As Integer
Dim DesLen As Integer


'Function GetStartUpValues() As Integer
'
'On Local Error GoTo getstart_Err
'
'If bDebug Then Handle_Err 0, "getstartupvalues-MDI"
'
'PassFrm.Show 1
'' --------------------------
'' New User therefore startup
'' Else normal startup.
'' * New user here, is based
'' on fact UserTour.mdb
'' doesn't exist...
'' --------------------------
'If ObjNew.info.NewUser Then
'        GetStartUpValues = 0
'Else
'        GetStartUpValues = objMdi.info.Load
'End If
'Exit Function
'getstart_Err:
'
'    If bDebug Then
'        MsgBox Error$(Err)
'        Handle_Err Err, "getstartupvalues-MDI"
'    End If
'    Resume Next
'End Function


Sub Setup_Desktop(Optional sLoadType As String)
' -------------------------------------
' Sets up Mdi desktop,
' Move Progress bar and other text box
' for different graphic settings.
' -------------------------------------

On Local Error GoTo Desk_Err

If "Program Begin" = sLoadType Then
    MDI.Caption = LoadResString(gcMDICaption)
    MdiGraBut.Picture = MdiGraUp.Picture
    MdiDayBut.Picture = MdiDayUp.Picture
    MdiEveBut.Picture = MdiEveUp.Picture
    MdiConBut.Picture = MdiConUp.Picture
    MdiNicBut.Picture = MdiNicUp.Picture
        ' Enable Icon bar buttons
    MDI!MdiDayBut.Enabled = True
    MDI!MdiGraBut.Enabled = True
    MDI!MdiEveBut.Enabled = True
    MDI!MdiConBut.Enabled = True
    MDI!MdiNicBut.Enabled = True
    MDI.ProgressBar.Enabled = True
    ' ----------------------------------------------
    ' Only show Desktop picture if ShowMeta is true!
    ' ----------------------------------------------
    If objMdi.info.ShowMeta = True Then
            MDI.Picture = LoadPicture(objMdi.info.MetaFile)
            
    End If
    'Make line for middle text box containing Name and dataPath.
    StatusBar1.Panels(2).Text = Trim$(objMdi.info.Name) & "  " & Trim$(objMdi.info.Datapath)
' ---------------------------------------
' Get MDI Form size from Registry and set
' enter into MDIForm
' ---------------------------------------

' Don't show Bad Key
gbSkipRegErrMsg = True
objMdi.info.iWindowState = Val(GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcWindowState)))
If objMdi.info.iWindowState >= 0 And objMdi.info.iWindowState <= 3 Then

            MDI.WindowState = objMdi.info.iWindowState

End If ' Ends LoadType if
End If ' Ends LoadType if

MDI.ProgressBack.left = (MDI.Width * 0.98) - MDI.ProgressBack.Width
MDI.MdiBarTxt.left = MDI.ProgressBack.left  'Bar inside starts at same Position

'picWeb.left = ProgressBar.Width
'picWeb.top = MDI.Toolbar1.Height
'picWeb.Width = MDI.ScaleWidth
'picWeb.Height = MDI.ScaleHeight

Exit Sub
Desk_Err:
    If bDebug Then Handle_Err Err, "Setup_Desktop-MDI"
    Resume Next
End Sub

Private Sub CalPodmnu_Click()
    Calndfrm.CalPodmnu
End Sub

Private Sub CalPoemnu_Click()
    Calndfrm.CalPoemnu
End Sub

Private Sub ConiRepmnu_Click()
Conconi_Report
End Sub

Private Sub DaiLv1mnu_Click()
DailyFrm.DaiLevVSc.Value = 5

End Sub

Private Sub DaiLv2mnu_Click()
DailyFrm.DaiLevVSc.Value = 4

End Sub


Private Sub DaiLv3mnu_Click()
DailyFrm.DaiLevVSc.Value = 3
'MDI.DaiLv1mnu.Checked = False
'MDI.DaiLv2mnu.Checked = False
'MDI.DaiLv3mnu.Checked = True
'MDI.DaiLv4mnu.Checked = False
'MDI.DaiLv5mnu.Checked = False
'
'DailyFrm.Set_Act_Hea_Cap (3)
End Sub


Private Sub DaiLv4mnu_Click()
DailyFrm.DaiLevVSc.Value = 2

End Sub


Private Sub DaiLv5mnu_Click()
DailyFrm.DaiLevVSc.Value = 1

End Sub


Private Sub DaiShowmnu_Click()
Dim i As Integer, allevels As String
Dim LenOfDesc As Integer, TypeNum As String
Dim RetStr As String
If bDebug Then Handle_Err 0, "DaiShowmnu-daily"
allevels = ""
For i = 1 To 9
 TypeNum = "Heart" + Format$(i, "0")  ' Format for InI Key Name.
 RetStr = Get_NameTour_HeartNames(TypeNum)
 If RetStr = "No Return" Then RetStr = "Field not used"
 LenOfDesc = 20 - Len(Trim$(RetStr))
 allevels = allevels & Trim$(RetStr) & String(LenOfDesc, 32) & "= " + aLevel(i) + vbLf
Next i
    MsgBox allevels, 64, "Individual heart rate durations."
End Sub

Private Sub Mdi_Time_Timer()
On Local Error GoTo Time_Err

Dim dRetLng As Long

' Use Windows API and obtain date setting from users machine

StatusBar1.Panels(3).Text = Format$(Now, "mmmm dd, yyyy hh:mm:ss") & "  "

If Not objMdi Is Nothing Then
    StatusBar1.Panels(2).Text = objMdi.info.Name & "  " & objMdi.info.Datapath & "  "
    
    If objMdi.info.RunOnce Then
        LoadStartUpForm objMdi.info.Load
    End If
    objMdi.info.RunOnce = False
End If

On Local Error GoTo 0
Exit Sub

Time_Err:
    If bDebug Then Handle_Err Err, "Time_Timer-Mdi"
    Resume Next
End Sub

Private Sub MdiAbomnu_Click()
Aboutfrm.Show vbModal
End Sub

Public Sub MdiBegmnu_Click()

Select Case MdiBegmnu.Caption
    Case GraphFrm_Begmnu:
            Graphfrm.GraBegmnu
    Case CalndFrm_Begmnu:
            Calndfrm.CalBegmnu  '
End Select
End Sub

Private Sub MdiCalmnu_Click()
Calndfrm.Show
End Sub

Private Sub MdiCdrmnu_Click()
    Daily_Calendar_Report
End Sub

Private Sub MdiCermnu_Click()
Event_Calendar_Report
End Sub

Public Sub MdiCndmnu_Click()
End Sub

Private Sub MdiCnlmnu_Click()
Contact_Report
End Sub

Private Sub MdiCntmnu_Click()
ContFrm.Show
End Sub

Private Sub MdiConBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MdiConBut.Picture = MdiConDn.Picture

End Sub


Private Sub MdiConBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        If X <= 0 Or X > MdiConBut.Width Or Y < 0 Or Y > MdiConBut.Height Then
            MdiConBut.Picture = MdiConUp.Picture
        Else
            MdiConBut.Picture = MdiConDn.Picture
        End If
    End Select


StatusBar1.Panels(1).Text = "Opens Contact Database"
DesLen = 0
End Sub


Private Sub MdiConBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MdiConBut.Picture = MdiConUp.Picture
DoEvents
    Select Case Button
    Case 1
        If X >= 0 And X < MdiConBut.Width And Y >= 0 And Y < MdiConBut.Height Then
            ContFrm.Show
        End If
    End Select

End Sub


Private Sub MdiConmnu_Click()
    #If Win32 Then
        Dim ret As Long
    #Else
        Dim ret As Integer
    #End If
    ret = HTMLHelp(MDI.hWnd, App.Path & "\TourWin.chm", cdlHelpContext, CLng(1))
    
End Sub

Private Sub MdiCopymnu_Click()
On Local Error Resume Next
' Copy selected text onto Clipboard.
If TypeOf ActiveForm.ActiveControl Is TextBox Then
   Clipboard.SetText ActiveForm.ActiveControl.SelText
End If

End Sub

Private Sub MdiCprmnu_Click()
sBuffer = "Report"
Peak_Calendar_Report
End Sub

Private Sub MdiCutmnu_Click()
On Local Error Resume Next
' ActiveForm refers to the active form in the MDI form.
 If TypeOf ActiveForm.ActiveControl Is TextBox Then
    ' Copy selected text onto Clipboard.
    Clipboard.SetText ActiveForm.ActiveControl.SelText
    ' Delete selected text.
    ActiveForm.ActiveControl.SelText = ""
 End If

End Sub

Private Sub MdiCycmnu_Click()
P_SetFrm.Show
End Sub

Private Sub MdiDarmnu_Click()
 Daily_Report
End Sub

Private Sub MdiDayBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo MdiDay_Err
If Button = 1 Then MdiDayBut.Picture = MdiDayDn.Picture
Exit Sub
MdiDay_Err:
    If bDebug Then Handle_Err Err, "MdiDayBut-Mdi"
    Resume Next
End Sub


Private Sub MdiDayBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > MdiDayBut.Width Or Y < 0 Or Y > MdiDayBut.Height Then
            MdiDayBut.Picture = MdiDayUp.Picture
        Else
            MdiDayBut.Picture = MdiDayDn.Picture
        End If
    End Select

StatusBar1.Panels(1).Text = "Opens daily file."
DesLen = 0
End Sub


Private Sub MdiDayBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Call Calendar Form only if
' mouse is release when over
' top of button
    MdiDayBut.Picture = MdiDayUp.Picture
    DoEvents
    Select Case Button
    Case 1
        If X >= 0 And X < MdiDayBut.Width And Y >= 0 And Y < MdiDayBut.Height Then
            DailyFrm.Show
        End If
    End Select


End Sub

Private Sub MdiDeletemnu_Click()
On Local Error Resume Next
If TypeOf ActiveForm.ActiveControl Is TextBox Then
   ActiveForm.ActiveControl.SelText = ""
End If
End Sub

Private Sub MdiDelmnu_Click()


Select Case MdiDelmnu.Caption
        Case DailyFrm_Delmnu:
             DailyFrm.DaiDelmnu
        Case CalndFrm_Delmnu:
             Calndfrm.CalDelmnu
        Case gcContFrm_Delmnu:
'             ContFrm.ConDelmnu_Click
        Case P_SetUp_Delmnu:
             P_SetFrm.PeaDelCmd_Click
        Case gcCONCONI_DELMNU:
            ConcFrm.DeleteMenu_Click
        Case MdiFrm_Delmnu:
            If objMdi.Delete Then
                Set objMdi = Nothing
                ' Connect to database
                Set objMdi = cTour_DB.OpenDBAndLogin()
                
                ' If objmdi is nothing, then user did not successfully
                ' login to database, therefore, end program...
                
                If objMdi Is Nothing Then
                    Unload Me
                    Exit Sub
                End If
                
                objMdi.LoadUserSettings
            End If
            
            
End Select
End Sub

Private Sub MdiDiamnu_Click()
DailyFrm.Show
End Sub


Private Sub MdiEveBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo MdiEve_Err
If Button = 1 Then MdiEveBut.Picture = MdiEveDn.Picture
Exit Sub
MdiEve_Err:
    If bDebug Then Handle_Err Err, "MdiEveBut-Mdi"
    Resume Next

End Sub


Private Sub MdiEveBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > MdiEveBut.Width Or Y < 0 Or Y > MdiEveBut.Height Then
            MdiEveBut.Picture = MdiEveUp.Picture
        Else
            MdiEveBut.Picture = MdiEveDn.Picture
        End If
    End Select

StatusBar1.Panels(1).Text = "Opens Events Calendar"
DesLen = 0

End Sub


Private Sub MdiEveBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' Call Calendar Form only if
' mouse is release when over
' top of button
MdiEveBut.Picture = MdiEveUp.Picture
DoEvents
    Select Case Button
    Case 1
        If X >= 0 And X < MdiEveBut.Width And Y >= 0 And Y < MdiEveBut.Height Then
            Calndfrm.Show
        End If
    End Select
    
End Sub


Private Sub mdiExportmnu_Click()
    
    
    ' Call Export to gather information
    frmExportOptions.Show vbModal
    
End Sub

Private Sub MDIForm_Initialize()

    LoadObjects

End Sub

Private Sub MDIForm_Load()
' =============================================================================
' MDIForm_Load: Purpose, program begins here, set global variables and
'               determine if database exits.
' =============================================================================

' =======================
' Declare local variables
' =======================
Dim LoadFrm As Integer

' ====================
' Set Error trapping &
' Write Debugging Info
' ====================
On Local Error GoTo MdiError

'Initialize local/global variabls
started = -1
iDailyDB = 0
iSearcherDB = 0


' Get Command Line Arguments
bDebug = IIf(UCase$(Command()) = "/DEBUG", True, False)
    
If bDebug Then
    Handle_Err 0, vbCrLf & String$(25, "*")
    Handle_Err 0, "Start Program"
End If


' Connect to database
Set objMdi = cTour_DB.OpenDBAndLogin()

' If objmdi is nothing, then user did not successfully
' login to database, therefore, end program...

If objMdi Is Nothing Then
    Unload Me
    Exit Sub
End If

' Check registration Record
cLicense.GetSystemRegistrationKey

If Not cLicense.IsKeyOK(cLicense.Key_Registration) Then
    Select Case cLicense.ShowTrialNotice
        Case 0: ' Register Later
            ' No steps required...
            
        Case 1: ' Canceled on Expired. End program
          Unload MDI
          Exit Sub
          
        Case 2: ' Registered successfully!
            ' No steps required...
            
        Case 3: ' Fail on Registration
          Unload MDI
          Exit Sub
          
    End Select
End If

' if Install then show setup wizard
If objMdi.info.WelcomeWizard Then 'or ObjNew.info.NewUser Or
   GetStartUpSetting (0)
   MDI.WindowState = 2 ' Maximized
   frmWizard.Show
   objMdi.info.WelcomeWizard = False
Else
    ' Load setup values
    ' ****************************
    ' * FLAG FOR TIMER CONTROL
    ' * TO RUN STARTUP ITEMS ONCE
    ' ****************************
    If 0 <> objMdi.info.Load Then
        objMdi.info.RunOnce = True
    Else
        objMdi.info.RunOnce = False
    End If
    GetStartUpSetting (objMdi.info.Load)
End If
    
App.HELPFILE = App.Path & "\TourWin.chm"

'Unload frmSplash
Screen.MousePointer = 0
'Browser.Navigate "www.velonews.com"

Exit Sub
MdiError:
    Unload frmSplash
    Resume Next
    
End Sub


Private Sub MDIForm_Resize()
On Local Error Resume Next

Setup_Desktop 'Resize Desktop controls

End Sub


Private Sub MDIForm_Unload(Cancel As Integer)

On Local Error GoTo Unload_Error
' -------------------------------
' Free all database connections
' -------------------------------
ObjTour.DBClose 11
Write_Tour_Setting_To_Registry "", "", ""

If bDebug Then
    Handle_Err 0, "End Program"
    Handle_Err 0, vbCrLf & String$(25, "*")
End If

On Error GoTo 0
Exit Sub

Unload_Error:
If bDebug Then Handle_Err Err, "Unload-Mdi"
Resume Next

End Sub


Private Sub MdiGraBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo MdiGra_Err
If Button = 1 Then MdiGraBut.Picture = MdiGraDn.Picture
Exit Sub
MdiGra_Err:
    If bDebug Then Handle_Err Err, "MdiGraBut-Mdi"
    Resume Next
End Sub

Private Sub MdiGraBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the button is pressed, display the up bitmap if the
    ' mouse is dragged outside the button's area, otherwise
    ' display the up bitmap
    Select Case Button
    Case 1
        If X <= 0 Or X > MdiGraBut.Width Or Y < 0 Or Y > MdiGraBut.Height Then
            MdiGraBut.Picture = MdiGraUp.Picture
        Else
            MdiGraBut.Picture = MdiGraDn.Picture
        End If
    End Select

StatusBar1.Panels(1).Text = "Opens Graphic Option."
DesLen = 0
End Sub


Private Sub MdiGraBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Call Calendar Form only if
' mouse is release when over
' top of button
MdiGraBut.Picture = MdiGraUp.Picture
DoEvents
    Select Case Button
    Case 1
        If X >= 0 And X < MdiGraBut.Width And Y >= 0 And Y < MdiGraBut.Height Then
            Screen.MousePointer = 11
            Graphfrm.Show
            Screen.MousePointer = 0
        End If
    End Select

End Sub

Private Sub MdiGramnu_Click()
Screen.MousePointer = 11
Graphfrm.Show

Screen.MousePointer = 0
End Sub

Private Sub MdiGraUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(1).Text = "Load Graph Program!"
End Sub


Private Sub MdiGuimnu_Click()
    cActivityNames.Type_ID = gcActive_Type_PeakNames
    cActivityNames.ShowForm
End Sub


Private Sub MdiHeamnu_Click()
    cActivityNames.Type_ID = gcActive_Type_HeartNames
    cActivityNames.ShowForm
End Sub


'------------------------------------------------------------------------------
' mdiImportmnu - Purpose, to import external information into currently opened
'                tourwin database for the current user.
'------------------------------------------------------------------------------
Private Sub mdiImportmnu_Click()

' Get File to import...
On Local Error GoTo Import_Error
If cPeakSchedules.GetFileAttributes(False) Then

    cPeakSchedules.DoImport ' Attempt to import info.
    cPeakNames.IsUpdated = False    ' This forces a reload of cPeakNames...
    
End If
Exit Sub
Import_Error:
    
    Err.Clear
End Sub

Private Sub MdiNewmnu_Click()

Select Case MdiNewmnu.Caption
    Case DailyFrm_Newmnu:
             DailyFrm.DaiNewmnu
             
    Case MdiFrm_Newmnu:
            objMdi.info.NewUser = True
            MdiOpt.Show
            
    Case gcContFrm_Newmnu:
            ContFrm.ConNewmnu_Click
            
    Case P_SetUp_Newmnu:
            P_SetFrm.cmdNew_Click
            
    Case gcCONCONI_NEWMNU:
            ConcFrm.NewMenu_Click
            
End Select
End Sub

Private Sub MdiNicBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MdiNicBut.Picture = MdiNicDn.Picture

End Sub


Private Sub MdiNicBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        If X <= 0 Or X > MdiNicBut.Width Or Y < 0 Or Y > MdiNicBut.Height Then
            MdiNicBut.Picture = MdiNicUp.Picture
        Else
            MdiNicBut.Picture = MdiNicDn.Picture
        End If
    End Select

StatusBar1.Panels(1).Text = "Opens Conconi Database"
DesLen = 0

End Sub


Private Sub MdiNicBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MdiNicBut.Picture = MdiNicUp.Picture
DoEvents
    Select Case Button
    Case 1
        If X >= 0 And X < MdiNicBut.Width And Y >= 0 And Y < MdiNicBut.Height Then
            ConcFrm.Show
        End If
    End Select
End Sub


Private Sub Mdinicmnu_Click()
ConcFrm.Show
End Sub

Private Sub MdiOptmnu_Click()

On Local Error GoTo MdiOptmnu_Err
Dim SQL As String

    Select Case MdiOptmnu.Caption
        Case DailyFrm_Option:
            DailyOpt.Show vbModal
        Case MdiFrm_Option:
            CallFrom = "MdiOptmnu"
            objMdi.info.NewUser = False
            MdiOpt.Show
        Case gcContFrm_Option:
            ContOpt.Show vbModal
            If Not UserCancel Then
'            SQL = "[" & ObjCont.info.IndxOrdr & "] " & ObjCont.info.SrtOrdr
'            SQL = "Select * FROM Contacts WHERE Id = " & objMdi.info.ID & " Order By " & SQL
'            ContFrm.ConData.RecordSource = SQL
'            ContFrm.ConData.Refresh
            End If
        Case gcCONCONI_OPTION:
            frmConcOpt.Show vbModal
        Case Else
            If bDebug Then Handle_Err Err, "Case Else " & MdiOptmnu.Caption
    End Select
Exit Sub
MdiOptmnu_Err:

    If bDebug Then Handle_Err Err, "MdiOptmnu_Click-Mdi"


Resume Next
End Sub

Private Sub MdiPastemnu_Click()
On Local Error Resume Next
' Paste Content of clipBoard to ActiveControl
If TypeOf ActiveForm.ActiveControl Is TextBox Then
   ActiveForm.ActiveControl.SelText = Clipboard.GetText()
End If

End Sub

Private Sub MdiPermnu_Click()
Weekly_Percentage_Report
End Sub

Private Sub MdiPrimnu_Click()

On Local Error GoTo MdiPrimnu_err
    
    Select Case ActiveForm.Name 'MdiOptmnu.Caption
        Case "DailyFrm":
                Daily_Report
        Case "Graphfrm":
              Graphfrm.PrintForm
        Case "Calndfrm":
              Select Case Calndfrm.EveFilCbo.Text
                    Case gcPeakChart
                            sBuffer = "Report"
                            Peak_Calendar_Report
                    Case gcEventChart
                            Event_Calendar_Report
                    Case gcDailyChart
                            Daily_Calendar_Report
              End Select
        Case "ContFrm":
                Contact_Report
                
        Case "ConcFrm":
                Conconi_Report
                
    End Select

Exit Sub
MdiPrimnu_err:
End Sub

Private Sub MdiPStmnu_Click()
    Printer_Setup
End Sub

Private Sub MdiSavmnu_Click()


Select Case MdiSavmnu.Caption
        Case DailyFrm_Savmnu:
            DailyFrm.DaiSavmnu
        Case MdiFrm_Savmnu:
            'MdiOpt.MdOSavmnu
        Case gcHrtZone_Savmnu:
            HrtZone.HeaSavCmd_Click
        Case gcMdO_Savmnu:
            MdiOpt.MdOSavmnu_Click
        Case P_SetUp_Savmnu:
            P_SetFrm.PeaSavCmd_Click
        Case gcCONCONI_SAVMNU:
            ConcFrm.SaveMenu_Click
End Select
End Sub

Private Sub MdiSeamnu_Click()
    #If Win32 Then
        Dim ret As Long
    #Else
        Dim ret As Integer
    #End If
    ret = HTMLHelp(MDI.hWnd, App.Path & "\" & HELPFILE, HELP_CONTEXT, CLng(0))
End Sub

Private Sub MdiSysHea_Click()
HrtZone.Show
End Sub

Private Sub MdiTecmnu_Click()
TechSupport
End Sub
Private Sub MdiTotmnu_Click()
WeeklyTotal_Report
End Sub

Private Sub MdiTypmnu_Click()
cActivityNames.Type_ID = gcActive_Type_EventNames
cActivityNames.ShowForm

End Sub

Private Sub MdiEximnu_Click()

On Local Error GoTo FileDoesnotExist
Select Case MdiEximnu.Caption
        Case CalndFrm_Exit:
                    Unload Calndfrm
        Case DailyFrm_Exit
                    DailyFrm.DaiEximnu
        Case GraphFrm_Exit:
                    Unload Graphfrm
                                        
        Case gcContFrm_Exit:
                    Unload ContFrm
                                        
        Case HrtZone_Exit:
                    Unload HrtZone
                    
        Case gcMdO_Exit:
                    Unload MdiOpt
                    
        Case P_SetUp_Exit:
                    Unload P_SetFrm
                    
        Case gcCONCONI_EXIT:
                    Unload ConcFrm
                    
        Case MdiFrm_Exit:
                    Unload Me ' End 'End TourWin... End was cause a GPF, so changed to unload
            
End Select
Exit Sub
FileDoesnotExist:
    Select Case Err.Number
            Case 0:
            Case Else:
            If bDebug Then Handle_Err Err, "MdiEximnu_Click-Mdi"
    End Select
    Resume Next
End Sub


Function GetStartUpSetting(iLoadFrm As Integer)


Setup_Desktop "Program Begin"                ' Check graphical settings
' validate database structure
Datapath = objMdi.info.Datapath

'CheckDataBaseStructure

End Function

Private Sub MDIWizMnu_Click()
    frmWizard.Show
End Sub

Private Sub mnuRepHis_Click()
    Historical_Report
End Sub





Private Sub picWeb_Resize()

'Browser.left = 0
'Browser.top = 0
'Browser.Width = picWeb.ScaleWidth
'Browser.Height = picWeb.ScaleHeight

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1:
            ' New Toolbar button
            ' Check if Active Form is Wizard...
            If ActiveForm Is Nothing Then
                MdiNewmnu_Click
                Exit Sub
            End If
            'Debug.Print ActiveForm.Name
            'if "ContFrm" = ActiveForm.Name Then Mdi
            If Not "frmWizard" = ActiveForm.Name Then MdiNewmnu_Click
        Case 2:     ' Save Toolbar button
            If ActiveForm Is Nothing Then Exit Sub
            If Not "frmWizard" = ActiveForm.Name Then MdiSavmnu_Click
        Case 4:     ' Print Toolbar Button
            If ActiveForm Is Nothing Then Exit Sub
            If Not "frmWizard" = ActiveForm.Name Then MdiPrimnu_Click
        Case 6:     ' Cut Toolbar button
            MdiCutmnu_Click
        Case 7:     ' Copy Toolbar Button
            MdiCopymnu_Click
        Case 8:     ' Paste Toolbar Button
            MdiPastemnu_Click

    End Select
End Sub


Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 Then DailyFrm.PopupMenu MDI.DaiLevmnu
End Sub

Public Sub RefreshPicture()

Const SRCCOPY = &HCC0020
Dim picSource         As PictureBox


    If objMdi.info.ShowMeta = True Then
'        picSource.Picture = LoadPicture(objMdi.info.MetaFile)
        Me.Picture = LoadPicture(objMdi.info.MetaFile)
    Else
        Me.Picture = LoadPicture("")   ' Basically clears MDI desktop...
    End If

Dim X As Integer, Y As Integer

'With picSource
'    For x% = 0 To ScaleWidth Step .ScaleWidth
'        For y% = 0 To ScaleHeight Step .ScaleHeight
'            BitBlt MDI.hDC, x%, y%, .ScaleWidth, _
'                     .ScaleHeight, .hDC, 0, 0, SRCCOPY
'        Next y%
'    Next x%
'End With

End Sub


'---------------------------------------------------------------------------------------
' PROCEDURE : LoadStartUpForm
' DATE      : 10/7/04 18:06
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function LoadStartUpForm(ByVal lLoad As Long) As Boolean

On Local Error GoTo LoadStartUpForm_Error
'Declare local variables

    
' Load form ?
 Select Case lLoad
       Case 0:             'Load nothing
       Case 1:
            DailyFrm.Show  'Load Daily...
       Case 2:
            Graphfrm.Show  'Load Graphfrm...
       Case 3:
            Calndfrm.Show  'Load CalndFrm...
       Case 4:
            ConcFrm.Show   'Load Conconi Form
       Case 5:
            ContFrm.Show
       Case Else:
            If bDebug Then Handle_Err Err, "LoadStartUpForm-Mdi --> " & CStr(objMdi.info.Load) & " not supported."
End Select

On Error GoTo 0
Exit Function

LoadStartUpForm_Error:
    If bDebug Then Handle_Err Err, "LoadStartUpForm-MDI"
    Resume Next


End Function
