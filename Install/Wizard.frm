VERSION 5.00
Begin VB.Form frmWizard 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "??? Wizard"
   ClientHeight    =   5730
   ClientLeft      =   1965
   ClientTop       =   1815
   ClientWidth     =   7515
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
   ScaleHeight     =   5730
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7575
      Begin VB.Image Image1 
         Height          =   915
         Left            =   6000
         Picture         =   "Wizard.frx":000C
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblStep"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Tag             =   "200"
         Top             =   240
         Width           =   6240
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
      Height          =   4065
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   7515
      Begin VB.Frame Frame2 
         Caption         =   "License Agreement"
         Height          =   3015
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   7215
         Begin VB.OptionButton OptLicense 
            Caption         =   "I Do NO&T agree license agreement."
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   25
            Top             =   2640
            Width           =   3135
         End
         Begin VB.OptionButton OptLicense 
            Caption         =   "I &agree license agreement."
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   2640
            Width           =   2895
         End
         Begin VB.TextBox txtLicense 
            Height          =   2175
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   23
            Text            =   "Wizard.frx":1946
            ToolTipText     =   "License Agreement"
            Top             =   360
            Width           =   6735
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Tag             =   "718"
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "User Information"
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
      Height          =   4065
      Index           =   1
      Left            =   -10000
      TabIndex        =   14
      Top             =   1080
      Width           =   7515
      Begin VB.Frame Frame1 
         Caption         =   "User Information"
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   7215
         Begin VB.TextBox txtFirstName 
            Height          =   285
            Left            =   1500
            MaxLength       =   15
            TabIndex        =   2
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtLastName 
            Height          =   285
            Left            =   1500
            MaxLength       =   20
            TabIndex        =   4
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtEmail 
            Height          =   285
            Left            =   1500
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "F&irst Name"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "&Last Name"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "&E-mail Address"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   5
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter user information. The first name you enter will be used as the login name the next time you start the application. "
         ForeColor       =   &H80000008&
         Height          =   870
         Index           =   1
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   6240
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Directory"
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
      Height          =   4065
      Index           =   2
      Left            =   -10000
      TabIndex        =   15
      Top             =   1080
      Width           =   7515
      Begin VB.CommandButton cmdChDir 
         Cancel          =   -1  'True
         Caption         =   "&Change Directory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5520
         MaskColor       =   &H00000000&
         TabIndex        =   19
         Top             =   3360
         Width           =   1845
      End
      Begin VB.Frame fraDir 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   4890
         Begin VB.Label lblDestDir 
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4440
         End
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   3015
         Index           =   2
         Left            =   3600
         Picture         =   "Wizard.frx":194C
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   3615
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
      Height          =   4065
      Index           =   3
      Left            =   -10000
      TabIndex        =   16
      Top             =   1080
      Width           =   7515
      Begin VB.Label lblBegin 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   4530
         WordWrap        =   -1  'True
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
      ScaleWidth      =   7515
      TabIndex        =   7
      Top             =   5160
      Width           =   7515
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   312
         Index           =   4
         Left            =   6240
         MaskColor       =   &H00000000&
         TabIndex        =   12
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   312
         Index           =   3
         Left            =   3840
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Height          =   312
         Index           =   2
         Left            =   2520
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Cancel"
         Height          =   312
         Index           =   1
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Help"
         Height          =   312
         Index           =   0
         Left            =   1320
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   105
         X2              =   7320
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   105
         X2              =   7320
         Y1              =   30
         Y2              =   30
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_STEPS = 4

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
Const STEP_FINISH = 3

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = "Blank Wizard"
Const INTRO_KEY = "IntroductionScreen"
Const SHOW_INTRO = "ShowIntro"
Const TOPIC_TEXT = "<TOPIC_TEXT>"

'module level vars
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean

Dim mbFinishOK      As Boolean


Private Sub cmdChDir_Click()
    ShowPathDialog gstrDIR_DEST

    If gfRetVal = gintRET_CONT Then
        lblDestDir.Caption = gstrDestDir

    End If

End Sub

Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    Dim lHelpTopic As Long
    Dim rc As Long
    Dim sTemp As String
    
    Select Case Index
        Case BTN_HELP
            mbHelpStarted = True
            lHelpTopic = HELP_BASE + 10 * (1 + mnCurStep)

        
        Case BTN_CANCEL
               ExitSetup Me, gintRET_EXIT
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            
            ' Check License agreement
            If mnCurStep = STEP_INTRO Then
                'Make sure on is selected and take appropriate action...
                If OptLicense(1).Value = False And OptLicense(0).Value = False Then
                    MsgBox "Please select agreement option.", vbOKOnly + vbInformation, "TourWin"
                    Exit Sub
                End If
                If OptLicense(1).Value = True Then
                    MsgBox "You have disagreed with agreement, installation will end.", vbOKOnly + vbInformation, "TourWin"
                    ExitSetup Me, 10 ' Do show cancel message
                End If
            End If
            ' =======================
            ' Check that First & last
            ' name are filled in
            ' =======================
            If mnCurStep = 1 And (frmWizard.txtLastName = "" Or frmWizard.txtFirstName = "") Then
                 MsgBox "Please enter first and last name.", vbOKOnly + vbInformation, "TourWin Version 1.0"
                 txtFirstName.SetFocus
                 Exit Sub
            End If
            nAltStep = mnCurStep + 1
            SetStep nAltStep, DIR_NEXT
            '
            If mnCurStep = 1 Then txtFirstName.SetFocus
            
        Case BTN_FINISH
            'wizard creation code goes here
            'gbInstallSampleDB = chkSampleDB.Value
            gFirstName = frmWizard.txtFirstName
            gLastName = frmWizard.txtLastName
            gEmail = frmWizard.txtEmail
            gDirectory = frmWizard.lblDestDir
            
            WriteRegistryEntries

            Unload Me
        
    End Select
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
    
    'Load All string info for Form
    LoadResStrings Me
    lblBegin.Caption = ResolveResString(resLBLBEGIN)
    fraDir.Caption = ResolveResString(resFRMDIRECTORY)
    'cmdChDir.Caption = ResolveResString(resBTNCHGDIR)
    lblDestDir.Caption = gstrDestDir
    frmWizard.Caption = "TourWin Version 1.85"
    
    ' Load License Agreement
    LoadLicenseAgreement txtLicense
    'Determine 1st Step:
        SetStep 0, DIR_NONE
    
End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
  
    Select Case nStep
        Case STEP_INTRO ' Intro & License Agreement...
      
        Case STEP_1 ' User Information
      
        Case STEP_2 'Location
          mbFinishOK = False
            
        Case STEP_FINISH
          mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    fraStep(nStep).Top = 1080
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

    Me.Caption = FRM_TITLE & " - " & LoadResString(fraStep(nStep).Tag)
    
    Select Case nStep
        Case 0:
            lblStep(0).Caption = LoadResString(200)
        Case 1:
            lblStep(0).Caption = "User Information"
        Case 2:
            lblStep(0).Caption = "Specify the install location."
        Case 3:
            lblStep(0).Caption = "Finished -  ready to install TourWin!"
    End Select

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If MsgWarning(ResolveResString(resASKEXIT), MB_ICONQUESTION Or MB_YESNO Or MB_DEFBUTTON2, gstrTitle) = IDNO Then
            Cancel = True
        Else
            ExitSetup Me, gintRET_CANCEL
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
    'see if we need to save the settings
    
      
'        SaveSetting APP_CATEGORY, WIZARD_NAME, "OptionName", Option Value
      
    
  
    
End Sub

Sub LoadResStrings(fForm As Form)
Dim oControl As Control
On Local Error Resume Next
For Each oControl In fForm
    If oControl.Tag <> "" Then
    If TypeOf oControl Is Label Then oControl.Caption = LoadResString(oControl.Tag)
    End If
Next
End Sub

Private Sub txtEmail_GotFocus()

With txtEmail
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtFirstName_GotFocus()

With txtFirstName
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtLastName_GotFocus()

With txtLastName
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub LoadLicenseAgreement(ByRef oControl As TextBox)

' Declare local variables...
Dim sLine As String

'Open license file and add to text box...
oControl.Text = ""

If "" <> Dir$(gstrSrcPath & "\Agreement.txt") Then
    Open gstrSrcPath & "\Agreement.txt" For Input As #1
    Do While Not EOF(1)
        Input #1, sLine
        oControl.Text = oControl.Text & sLine & vbCrLf
    Loop
    
    Close 1
End If  ' End check for file...
End Sub

Private Sub WriteRegistryEntries()

Dim sTemp       As String
Dim sRetStr     As String

' Check if current DB is listed in DBase key, if not add!
gbSkipRegErrMsg = True
sRetStr = GetRegStringValue("HKEY_LOCAL_MACHINE\Software\Tourwin", "DBase")

' Check if key exist, if not create
If REG_ERROR = sRetStr Then
    gbSkipRegErrMsg = True
    sTemp = "HKEY_LOCAL_MACHINE\Software\Tourwin\DBase"
    CreateRegKey sTemp
    
    ' Write entries to registry
    sTemp = "HKEY_LOCAL_MACHINE\Software\Tourwin"
    gbSkipRegErrMsg = True
    WriteRegStringValue sTemp, "DBase", frmWizard.lblDestDir
Else
' -------------------------
' Check if current db is
' listed in return string
' if not, append to end.
' -------------------------
    If 0 = InStr(1, sRetStr, frmWizard.lblDestDir, vbTextCompare) Then
        gbSkipRegErrMsg = True
        WriteRegStringValue "HKEY_LOCAL_MACHINE\Software\Tourwin", "DBase", sRetStr & "," & frmWizard.lblDestDir
    End If
End If

' Create Software\TourWin key

sTemp = "HKEY_LOCAL_MACHINE\Software\Tourwin\UserName"
CreateRegKey sTemp
sTemp = "HKEY_LOCAL_MACHINE\Software\Tourwin"
gbSkipRegErrMsg = True
WriteRegStringValue sTemp, "UserName", frmWizard.txtFirstName

sTemp = "HKEY_LOCAL_MACHINE\Software\Tourwin\Tour_Exe"
CreateRegKey sTemp
sTemp = "HKEY_LOCAL_MACHINE\Software\Tourwin"
gbSkipRegErrMsg = True
WriteRegStringValue sTemp, "Tour_Exe", frmWizard.lblDestDir

End Sub
