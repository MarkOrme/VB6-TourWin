VERSION 5.00
Begin VB.Form frmWizard 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "??? Wizard"
   ClientHeight    =   5055
   ClientLeft      =   1965
   ClientTop       =   1815
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Wizard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7155
   Tag             =   "1"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Tracking Daily Activity"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   4
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7245
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   35
         Tag             =   "295"
         Top             =   3840
         Width           =   5055
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   34
         Tag             =   "294"
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Image ImgWizGraph 
         Height          =   495
         Left            =   600
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image ImgWizDaily 
         Height          =   495
         Left            =   600
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   975
         Left            =   1440
         TabIndex        =   33
         Tag             =   "293"
         Top             =   2760
         Width           =   5295
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   600
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblDefWiz 
         Caption         =   "lblDefWiz"
         Height          =   495
         Left            =   600
         TabIndex        =   20
         Tag             =   "291"
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   975
         Index           =   10
         Left            =   1440
         TabIndex        =   17
         Tag             =   "292"
         Top             =   1320
         Width           =   5280
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Peak schedule Setup..."
      Enabled         =   0   'False
      Height          =   4425
      Index           =   3
      Left            =   -10000
      TabIndex        =   14
      Tag             =   "280"
      Top             =   0
      Width           =   7245
      Begin VB.CommandButton cmdScheduleShow 
         Caption         =   "Show Me"
         Height          =   375
         Left            =   5760
         TabIndex        =   32
         Top             =   3720
         Width           =   1092
      End
      Begin VB.CommandButton cmdHeartShow 
         Caption         =   "Show Me"
         Height          =   375
         Left            =   5760
         TabIndex        =   31
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   600
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   495
         Left            =   600
         TabIndex        =   30
         Tag             =   "270"
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label Label4 
         Caption         =   "Heart Rate Zone:"
         Height          =   975
         Left            =   600
         TabIndex        =   19
         Tag             =   "290"
         Top             =   2520
         Width           =   6135
      End
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   855
         Index           =   9
         Left            =   600
         TabIndex        =   15
         Tag             =   "280"
         Top             =   1200
         Width           =   6360
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Setting up"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   2
      Left            =   -10000
      TabIndex        =   12
      Top             =   0
      Width           =   7245
      Begin VB.CommandButton cmdEventShow 
         Caption         =   "Show Me"
         Height          =   375
         Left            =   5760
         TabIndex        =   29
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdPeakShow 
         Caption         =   "Show Me"
         Height          =   375
         Left            =   5760
         TabIndex        =   28
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   600
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   735
         Index           =   12
         Left            =   600
         TabIndex        =   27
         Tag             =   "272"
         Top             =   1200
         Width           =   5640
      End
      Begin VB.Label Label3 
         Caption         =   "Event Name"
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Tag             =   "270"
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   855
         Index           =   8
         Left            =   600
         TabIndex        =   13
         Tag             =   "271"
         Top             =   2520
         Width           =   5640
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Planning Your Year"
      Enabled         =   0   'False
      Height          =   4305
      Index           =   1
      Left            =   -10000
      TabIndex        =   10
      Top             =   120
      Width           =   7245
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   615
         Index           =   11
         Left            =   1440
         TabIndex        =   26
         Tag             =   "253"
         Top             =   3240
         Width           =   4680
      End
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   855
         Index           =   6
         Left            =   1440
         TabIndex        =   25
         Tag             =   "252"
         Top             =   2160
         Width           =   4680
      End
      Begin VB.Image ImgWizCalendar 
         Height          =   495
         Left            =   600
         Top             =   2160
         Width           =   495
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   600
         X2              =   6720
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Image ImgWizContact 
         Height          =   495
         Left            =   600
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   615
         Index           =   4
         Left            =   1440
         TabIndex        =   24
         Tag             =   "251"
         Top             =   1200
         Width           =   4680
      End
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   615
         Index           =   7
         Left            =   600
         TabIndex        =   11
         Tag             =   "250"
         Top             =   240
         Width           =   6120
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Introduction Screen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   0
      Left            =   -10000
      TabIndex        =   6
      Top             =   0
      Width           =   7155
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   600
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   3
         Left            =   600
         TabIndex        =   23
         Tag             =   "263"
         Top             =   3120
         Width           =   5880
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   22
         Tag             =   "262"
         Top             =   2160
         Width           =   5880
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   21
         Tag             =   "261"
         Top             =   1320
         Width           =   5880
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Tag             =   "260"
         Top             =   360
         Width           =   5880
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Finished!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   5
      Left            =   -10000
      TabIndex        =   8
      Tag             =   "3000"
      Top             =   0
      Width           =   7155
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         ForeColor       =   &H80000008&
         Height          =   990
         Index           =   5
         Left            =   2880
         TabIndex        =   9
         Tag             =   "296"
         Top             =   210
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   3075
         Index           =   5
         Left            =   210
         Picture         =   "Wizard.frx":000C
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2430
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   4485
      Width           =   7155
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   312
         Index           =   4
         Left            =   1320
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "213"
         Top             =   120
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   312
         Index           =   3
         Left            =   3840
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "212"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Height          =   312
         Index           =   2
         Left            =   2640
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "211"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   312
         Index           =   1
         Left            =   5880
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Help"
         Height          =   312
         HelpContextID   =   1
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "254"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   108
         X2              =   7012
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   108
         X2              =   7012
         Y1              =   24
         Y2              =   24
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_STEPS = 6

Const RES_ERROR_MSG = 30000

'BASE VALUE FOR HELP FILE FOR THIS WIZARD:
Const HELP_BASE = 1000
Const HELP_FILE = "MYWIZARD.HLP"

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_FINISH = 5

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = gcTourVersion
Const INTRO_KEY = "IntroductionScreen"
Const SHOW_INTRO = "ShowIntro"
Const TOPIC_TEXT = "<TOPIC_TEXT>"
'Frame Names
Const gIntroduction_Screen = 0
Const gUserInformation = 1
Const gEvent = 2
Const gHeart = 3
Const gTraining = 4
Const gFinished = 5


'module level vars
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean
Dim mbFinishOK      As Boolean

Private Sub chkShowIntro_Click()
'    If chkShowIntro.Value Then
'        SaveSetting APP_CATEGORY, WIZARD_NAME, INTRO_KEY, SHOW_INTRO
'    Else
'        SaveSetting APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString
'    End If
End Sub

Private Sub cmdEventShow_Click()
        SendKeys "%{s}ne", False
End Sub

Private Sub cmdHeartShow_Click()
    SendKeys "%{s}nh", False
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    Dim lHelpTopic As Long
    Dim rc As Long
    Dim bValid As Boolean
    Dim SQL As String
    Dim iLoop As Integer
    
    
    ' Flag whether user as correctly inputed data
    bValid = True
    
    Select Case Index
        Case BTN_HELP
            mbHelpStarted = True
            lHelpTopic = HELP_BASE + 10 * (1 + mnCurStep)
            rc = HTMLHelp(MDI.hwnd, App.Path & "\" & HELPFILE, cdlHelpContext, CLng(1))
                 
        Case BTN_CANCEL
                'If vbYes = MsgBox(LoadResString(270), vbYesNo + vbQuestion, LoadResString(gcTourVersion)) Then
                Unload Me
                    'Unload MDI
                    
                'End If
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
                 Select Case mnCurStep
                        Case gIntroduction_Screen
                        Case gUserInformation
                            ' Check the User Name & database is fill in
                        Case gEvent
                        Case gHeart
                        Case gTraining
                        Case gFinished
                End Select
                
            If bValid Then  ' Check if next frame should be displayed
            
                nAltStep = mnCurStep + 1
                SetStep nAltStep, DIR_NEXT
                
            End If
        Case BTN_FINISH
'            'wizard creation code goes here
'            ' Create Database
'            'CreateTourDatabases txtDirWiz, "TourWin.mdb"
'
'            ' Write local Registry
'            objMdi.info.dataPath = txtDirWiz & "\"
'            ' Create Software\TourWin key
'            CreateRegKey LoadResString(gcRegTourKey)
'            WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcTourDBase), objMdi.info.dataPath
'
'            ' Create Default Table Entries
'            ' --------------------------
'            ' Define newly created user
'            ' --------------------------
'
'            ' Open Database and add user
'
'            ' ------------------
'            ' Define SearcherDB
'            ' ------------------
'            SQL = "Select * From " & gcUserTour_UserTbl
'            ObjTour.RstSQL iSearcherDB, SQL
'
'            ObjTour.AddNew iSearcherDB  ' Add new record
'            UpdateFormValueToDatabase
'            ObjTour.Update iSearcherDB  ' Update record
'
'            ' --------------------------
'            ' Define newly created user
'            ' --------------------------
'            SQL = "Name = '" & txtUserWiz & "'"
'            ObjTour.DBFindFirst SQL, iSearcherDB
'            objMdi.info.ID = Int(Val(ObjTour.DBGetField("ID", iSearcherDB)))
'            NewUser.CreateDailyOpt objMdi.info.ID
'            NewUser.CreateContactOpt objMdi.info.ID
'
'            ' -------------------------------
'            ' Add Event entry to Events Table
'            ' -------------------------------
'                SQL = "Select * From " & gcNameTour_Events & " WHERE ID='" & objMdi.info.ID & "'"
'                ObjTour.RstSQL iSearcherDB, SQL
'                If ObjTour.RstRecordCount(iSearcherDB) = 0 Then
'                    ObjTour.AddNew iSearcherDB  ' Add new record
'                    ObjTour.DBSetField "ID", objMdi.info.ID, iSearcherDB
'                Else
'                    ObjTour.Edit iSearcherDB
'                End If
'                ' Update field value
'                ObjTour.DBSetField "Event0", txtEventWiz, iSearcherDB
'                ObjTour.DBSetField "Color0", 255, iSearcherDB
'                ObjTour.Update iSearcherDB  ' Update record
'
'            ' -------------------------------
'            ' Add Heart entry to HeartNames Table
'            ' -------------------------------
'                SQL = "Select * From " & gcNameTour_HeartNames & " WHERE ID='" & objMdi.info.ID & "'"
'                ObjTour.RstSQL iSearcherDB, SQL
'                If ObjTour.RstRecordCount(iSearcherDB) = 0 Then
'                    ObjTour.AddNew iSearcherDB  ' Add new record
'                    ObjTour.DBSetField "ID", objMdi.info.ID, iSearcherDB
'                Else
'                    ObjTour.Edit iSearcherDB
'                End If
'                ' Update field value
'                ObjTour.DBSetField "Heart1", txtHeartWiz, iSearcherDB
'
'                ObjTour.Update iSearcherDB  ' Update record
'
'
'            ' -------------------------------
'            ' Add Peak entry to PeakNames Table
'            ' -------------------------------
'                SQL = "Select * From " & gcNameTour_PeakNames & " WHERE ID='" & objMdi.info.ID & "'"
'                ObjTour.RstSQL iSearcherDB, SQL
'                If ObjTour.RstRecordCount(iSearcherDB) = 0 Then
'                    ObjTour.AddNew iSearcherDB  ' Add new record
'                    ObjTour.DBSetField "ID", objMdi.info.ID, iSearcherDB
'                Else
'                    ObjTour.Edit iSearcherDB
'                End If
'                ' Update field value
'                ObjTour.DBSetField "Peak0", txtPeaWiz, iSearcherDB
'                ObjTour.Update iSearcherDB  ' Update record
'
'
'            ObjNew.info.NewUser = False
'            ' ----------------
'            ' Enable tool bar
'            ' ----------------
'            MDI.Setup_Desktop "Program Begin"
'            ' Load Daily Settings
'            ObjTour.RstSQL iSearcherDB, "SELECT * FROM " & gcUserTour_DailyOpt & " WHERE ID = " & objMdi.info.ID
'            objDai.Load_Daily_Settings
'
'            Unload Me
'
'            'If GetSetting(APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, vbNullString) = vbNullString Then
'             '   frmConfirm.Show vbModal
'            'End If
        
    End Select
End Sub

Private Sub cmdPeakShow_Click()

    SendKeys "%{s}np", False
End Sub

Private Sub cmdscheduleshow_Click()
    SendKeys "%{s}sp", False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmdNav_Click BTN_HELP
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'init all vars
    mbFinishOK = False
    
    For i = 0 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
    
    ' Center form on owner
    CentreForm Me, 0
    'Load All string info for Form
    
    LoadFormResourceString Me
    Me.Caption = Me.Caption & " - Welcome " & objMdi.info.Name
    ImgWizContact.Picture = MDI.MdiConUp.Picture
    ImgWizCalendar.Picture = MDI.MdiEveUp.Picture
    ImgWizDaily.Picture = MDI.MdiDayUp.Picture
    ImgWizGraph.Picture = MDI.MdiGraUp.Picture
    SetStep 0, DIR_NONE

End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
  
    Select Case nStep
        Case STEP_INTRO
      
        Case STEP_1
      
        Case STEP_2
        
        Case STEP_3
      
        Case STEP_4
            mbFinishOK = False
      
        Case STEP_FINISH
            mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
  
End Sub

Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = 0 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

    'Me.Caption = FRM_TITLE & " - " & LoadResString(fraStep(nStep).Tag)

End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error message
    sTmp = LoadResString(RES_ERROR_MSG)
    'get the specific message
    sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
    'see if we need to save the settings
    
    'If mbHelpStarted Then rc = HTMLHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
End Sub

'Function UpdateFormValueToDatabase()
'
'        ' Update PassWord setting
'        '"Id" is autoIncr by jet database...
'        ObjTour.DBSetField "Security", False, iSearcherDB
'        objMdi.info.Security = False
'
'        ObjTour.DBSetField "Load", 0, iSearcherDB
'        objMdi.info.Load = 0
'
'        ObjTour.DBSetField "Name", Trim$(txtUserWiz), iSearcherDB
'        objMdi.info.Name = txtUserWiz
'
'        ObjTour.DBSetField "DataPath", objMdi.info.dataPath, iSearcherDB
'
'        ObjTour.DBSetField "PassWord", " ", iSearcherDB
'        ObjTour.DBSetField "MetaFile", " ", iSearcherDB
'        ObjTour.DBSetField "ShowMeta", False, iSearcherDB
'
'End Function

Private Sub wizEximnu_Click()
    End
End Sub

Private Sub wizHelmnu_Click()
Dim Ret As Long
    Ret = HTMLHelp(MDI.hwnd, App.Path & "\" & HELPFILE, cdlHelpContext, CLng(1))
End Sub

