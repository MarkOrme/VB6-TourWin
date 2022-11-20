VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5370
   ClientLeft      =   1620
   ClientTop       =   1500
   ClientWidth     =   6765
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6765
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   3900
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picWorking 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   1515
      Left            =   510
      ScaleHeight     =   1455
      ScaleWidth      =   5730
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1950
      Width           =   5790
      Begin VB.Label lblWorking 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   120
         TabIndex        =   13
         Top             =   150
         Width           =   5415
      End
   End
   Begin VB.Frame fraPasswords 
      BackColor       =   &H00FF0000&
      Height          =   4740
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6765
      Begin MSFlexGridLib.MSFlexGrid grdPasswords 
         Height          =   1911
         Left            =   260
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2665
         Width           =   6318
         _ExtentX        =   11165
         _ExtentY        =   3360
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   3
         Left            =   5775
         Picture         =   "frmMain.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   25
         Top             =   450
         Width           =   510
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   525
         Picture         =   "frmMain.frx":0614
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   24
         Top             =   450
         Width           =   510
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   225
         TabIndex        =   11
         Top             =   1275
         Width           =   6315
         Begin VB.ComboBox cboSpecial 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2025
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   675
            Width           =   2340
         End
         Begin VB.TextBox txtQuantity 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   300
            MaxLength       =   5
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   675
            Width           =   1365
         End
         Begin VB.ComboBox cboNbrCount 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4875
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   675
            Width           =   1065
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Special Considerations"
            Height          =   240
            Left            =   2250
            TabIndex        =   23
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Number of passwords to create (Max 2500)"
            Height          =   495
            Index           =   1
            Left            =   225
            TabIndex        =   22
            Top             =   225
            Width           =   1530
         End
         Begin VB.Label lblNbrCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "  Number of letters per password to convert"
            Height          =   465
            Left            =   4500
            TabIndex        =   14
            Top             =   225
            Width           =   1590
         End
      End
      Begin VB.Label lblPasswordTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Passwords"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   300
         TabIndex        =   26
         Top             =   300
         Width           =   6090
      End
   End
   Begin VB.Frame fraPhrase 
      BackColor       =   &H0000FF00&
      Height          =   4740
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6765
      Begin VB.TextBox txtPassphrase 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         HideSelection   =   0   'False
         Left            =   225
         MultiLine       =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2625
         Width           =   6315
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   5775
         Picture         =   "frmMain.frx":091E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   19
         Top             =   450
         Width           =   510
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   525
         Picture         =   "frmMain.frx":0C28
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   18
         Top             =   450
         Width           =   510
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   225
         TabIndex        =   9
         Top             =   1275
         Width           =   6315
         Begin VB.ComboBox cboNbrOfWords 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2025
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   675
            Width           =   1440
         End
         Begin VB.ComboBox cboLength 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   675
            Width           =   1440
         End
         Begin VB.ComboBox cboSpecial 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   675
            Width           =   2340
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Number of words per phrase"
            Height          =   420
            Index           =   2
            Left            =   2025
            TabIndex        =   17
            Top             =   225
            Width           =   1440
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Special considerations"
            Height          =   270
            Index           =   1
            Left            =   3900
            TabIndex        =   16
            Top             =   375
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Varying Word lengths"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   375
            Width           =   1590
         End
      End
      Begin VB.Label lblPassPhrase 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Passphrase"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   915
         Left            =   300
         TabIndex        =   20
         Top             =   300
         Width           =   6090
      End
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4875
      Width           =   1000
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4425
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4875
      Width           =   1000
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":0F32
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   60
      TabIndex        =   28
      Top             =   4740
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblMyName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   375
      TabIndex        =   27
      Top             =   4800
      Width           =   2565
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptCreate 
         Caption         =   "What to &create"
         Begin VB.Menu mnuOptPWD 
            Caption         =   "Password"
            Index           =   0
         End
         Begin VB.Menu mnuOptPWD 
            Caption         =   "P&assphrase"
            Index           =   1
         End
      End
      Begin VB.Menu mnuOptLength 
         Caption         =   "&Min word length"
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "3 Chars"
            Index           =   0
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "4 Chars"
            Index           =   1
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "5 Chars"
            Index           =   2
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "6 Chars"
            Index           =   3
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "7 Chars"
            Index           =   4
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "8 Chars"
            Index           =   5
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "9 Chars"
            Index           =   6
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "10 Chars"
            Index           =   7
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "11 Chars"
            Index           =   8
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "12 Chars"
            Index           =   9
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "13 Chars"
            Index           =   10
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "14 Chars"
            Index           =   11
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "15 Chars"
            Index           =   12
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "16 Chars"
            Index           =   13
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "17 Chars"
            Index           =   14
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "18 Chars"
            Index           =   15
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "19 Chars"
            Index           =   16
         End
         Begin VB.Menu mnuOptCharLength 
            Caption         =   "20 Chars"
            Index           =   17
         End
      End
      Begin VB.Menu mnuOptTypeCase 
         Caption         =   "&Type of case"
         Begin VB.Menu mnuOptCase 
            Caption         =   "Lowercase"
            Index           =   0
         End
         Begin VB.Menu mnuOptCase 
            Caption         =   "Uppercase"
            Index           =   1
         End
         Begin VB.Menu mnuOptCase 
            Caption         =   "Propercase"
            Index           =   2
         End
         Begin VB.Menu mnuOptCase 
            Caption         =   "Mixed Case"
            Index           =   3
         End
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptOmit 
         Caption         =   "Om&it These Characters"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***************************************************************************
' Project:       Passphrase.vbp
'
' Module:        frmMain
'
' Description:   This is the main form that has multiple layers.
'                There are four parts:
'                    1.  The password generation.
'                    2.  The passphrase generation.
'                    3.  The passphrase display
'                    4.  The Working message.
'
'                The user can stop this process immediately by pressing the
'                STOP or EXIT button.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-JAN-2000  Kenneth Ives     Module created
' ***************************************************************************
  
' ---------------------------------------------------------------------------
' Define module level variables (prefixed with "m_")
' ---------------------------------------------------------------------------
  ' -- String variables
  Private m_strTextData As String
  Private m_strFilename As String
  
  ' -- Long Integer variables
  Private m_lngNbrCount As Long
  Private m_lngCount    As Long
  Private m_intColumns  As Long
  
  ' -- Boolean variables
  Private m_bAlphabetic As Boolean
  
  ' -- Integer variables
  Private m_intWordLength As Integer
  Private m_intOmitCnt As Integer
  
' ---------------------------------------------------------
' Reduce flicker while loading a control
'
' Lock the window to prevent redrawing
' Syntax:  LockWindowUpdate list1.hWnd
'
' Unlock display
' Syntax:  LockWindowUpdate 0&
' ---------------------------------------------------------
  Private Declare Function LockWindowUpdate Lib "user32" _
          (ByVal hWnd As Long) As Long
  
Private Sub cboLength_Click()

' ---------------------------------------------------------------------------
' Determine the number of characters and range, if any
' ---------------------------------------------------------------------------
  If InStr(cboLength.Text, "All") Then
      g_intMinLength = Val(Right(Trim(cboLength.Text), 2))
      g_intMaxLength = 0
  Else
      g_intMinLength = Val(Left(Trim(cboLength.Text), 2))
      g_intMaxLength = Val(Right(Trim(cboLength.Text), 2))
  End If
  
End Sub

Private Sub cboSpecial_Click(Index As Integer)

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim i As Integer

' ---------------------------------------------------------------------------
' Reset the number of letters to convert
' ---------------------------------------------------------------------------
  m_lngNbrCount = 0

' ---------------------------------------------------------------------------
' determine the number of characters to convert in a password
' ---------------------------------------------------------------------------
  If Index = 1 Then  ' passwords
      Select Case Trim(cboSpecial(Index).Text)
           
           Case "1-Alphabetic Only"
                cboNbrCount.Clear
                cboNbrCount.Enabled = False
                lblNbrCount.Enabled = False
                m_bAlphabetic = True
                
           Case Else
                lblNbrCount.Enabled = True
                cboNbrCount.Enabled = True
                m_bAlphabetic = False
                cboNbrCount.Clear
                
                ' load the combo box with numbers
                ' always one short so the first
                ' character is alphabetic
                For i = 1 To (m_intWordLength - 2)
                    cboNbrCount.AddItem " " & i
                Next
                               
                ' display the first item in the list
                cboNbrCount.ListIndex = 0
      End Select
  End If
  
End Sub

Private Sub cmdChoice_Click(Index As Integer)
  
' ***************************************************************************
' Routine:       cmdChoice_Click
'
' Description:   Perform the commands associated with a command button
'
' Parameters:    Index - which button was pressed
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

  Select Case Index
         Case 0:  ' Start and Stop
              If cmdChoice(0).Caption = "&Start" Then
                  
                  ' if START is displayed on the command button change
                  ' it to STOP and set the cancellation switch
                  g_bStop_Pressed = False
                  cmdChoice(0).Caption = "&Stop"
                  
                  ' empty the output boxes
                  txtPassphrase = ""
                  grdPasswords.Clear
  
                  ' display a message of work in progress
                  With lblWorking
                       .BackColor = &H80FFFF  ' Bright yellow
                       .Caption = "Working"
                  End With
                  With picWorking
                       .Enabled = True
                       .BackColor = &HFF&     ' bright Red border
                       .Visible = True
                  End With
                  DoEvents
                  
                  ' see if we are to create passwords
                  If g_bCreate_Password Then
                      ' Create passwords
                      Create_Passwords
                       
                      ' Format the output display
                      Format_PWD_Display
                  Else
                      ' create a passphrase
                      ' Gather all the information from the screen
                      ' see haw many words are to be used in the
                      ' passphrase
                      g_lngNbrOfWords = CLng(Trim(cboNbrOfWords.Text))
                      Create_Special_Word
                       
                      ' Build the passphrase
                      Create_Phrase
                      
                      ' Display the passphrase
                      Update_Screen 2
                  End If
                      
                  cmdChoice(0).Caption = "&Start"
                  ' Hide the working message
                  With picWorking
                       .Enabled = False
                       .Visible = False
                  End With
                  DoEvents
                  
              ElseIf cmdChoice(0).Caption = "&Stop" Then
                  ' if STOP is displayed on the command button change
                  ' it to START and reset the cancellation switch
                  cmdChoice(0).Caption = "&Start"
                  g_bStop_Pressed = True
                  
                  ' Hide the working message
                  With picWorking
                       .Enabled = False
                       .Visible = False
                  End With
                  DoEvents
              End If
         
         Case 1:  ' Exit application
              ' see Form_QueryUnload event
              Unload Me
  End Select
  
End Sub

Private Sub Update_Screen(Index As Integer)
   
' ***************************************************************************
' Routine:       Update_Screen
'
' Description:   Prepare a screen for a display update.  this can be either
'                the initial input or output.
'
' Parameters:    Index - which screen to update
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Update the display
' ---------------------------------------------------------------------------
  Select Case Index
         Case 0: ' Prepare the password frame
              fraPhrase.Visible = False
              fraPasswords.Visible = True
         
         Case 1: ' Prepare the passphrase frame
              fraPasswords.Visible = False
              fraPhrase.Visible = True
         
         Case 2: ' display the passphrase
              txtPassphrase = g_strPassphrase
  End Select
  
End Sub

Private Sub Form_Initialize()

' ------------------------------------------------------------------------------------------
' Centered on the screen.
' ------------------------------------------------------------------------------------------
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

End Sub

Private Sub Form_Load()
  
' ---------------------------------------------------------------------------
' Initialize combo boxes and variables
' ---------------------------------------------------------------------------
  g_intMinLength = 0
  g_intMaxLength = 0
  m_intColumns = 0
  m_intOmitCnt = 0
  m_strFilename = ""
  g_strDBName = App.Path & "\PWords.dat"
  Erase g_arOmit()     ' Empty the char omission array
  mnuOptPWD_Click 0    ' display the password screen
  mnuOptCase_Click 0   ' default is lowercase display
  cboSpecial_Click 1   ' Default to the passwords
  Fill_Combo_Boxes     ' load the combo boxes
  
' ---------------------------------------------------------------------------
' set up the form to be displayed
' ---------------------------------------------------------------------------
  With frmMain
        .lblMyName = "Freeware by Kenneth Ives" & vbLf & "kenaso@home.com"
        .txtQuantity = "1"
        .Caption = "Password/Passphrase Creator v" & App.Major & "." & App.Minor
        .picWorking.Visible = False
        .grdPasswords.Rows = 0
        .grdPasswords.Cols = 0
        '
        .Show vbModeless
        .Refresh
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' ------------------------------------------------------------------------------------------
' Based on the the unload code the system passes, we determine what to do.
'
' Unloadmode codes
'     0 - Close from the control-menu box or Upper right "X"
'     1 - Unload method from code elsewhere in the application
'     2 - Windows Session is ending
'     3 - Task Manager is closing the application
'     4 - MDI Parent is closing
' ------------------------------------------------------------------------------------------
  Select Case UnloadMode
         Case 0: StopTheProgram
         Case Else: ' Fall thru. Something else is shutting us down.
  End Select

End Sub

Private Sub mnuAbout_Click()

' ---------------------------------------------------------------------------
' Define variables
' ---------------------------------------------------------------------------
  Dim sMsg          As String
  Dim sMsgBoxTitle  As String
  Dim iMsgBoxResp   As Integer
  Dim iResponse     As Integer
 
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  sMsgBoxTitle = App.FileDescription & " v" & App.Major & "." & App.Minor
  sMsg = StrConv(App.EXEName & ".exe", vbProperCase) & vbCrLf
  sMsg = sMsg & "Written by " & App.CompanyName & vbCrLf
  sMsg = sMsg & App.LegalCopyright & vbCrLf
  
  iMsgBoxResp = vbOKOnly + vbInformation + vbApplicationModal
  MsgBox sMsg, iMsgBoxResp, sMsgBoxTitle
   
End Sub

Private Sub mnuOptCase_Click(Index As Integer)
  
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intIndex As Integer
  
' ---------------------------------------------------------------------------
' Set the appropriate check mark
' ---------------------------------------------------------------------------
  For intIndex = 0 To 3
      If intIndex = Index Then
          mnuOptCase(intIndex).Checked = True
          g_intTypeCase = intIndex
      Else
          ' remove all other check marks
          mnuOptCase(intIndex).Checked = False
      End If
  Next
  
End Sub

Private Sub mnuOptCharLength_Click(Index As Integer)
  
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intIndex  As Integer
  Dim intItems  As Integer
  
' ---------------------------------------------------------------------------
' Display the appropriate values
' ---------------------------------------------------------------------------
  If g_bCreate_Password Then
      For intIndex = 6 To 17
          mnuOptCharLength(intIndex).Visible = True
      Next
      intItems = 17  ' number of items on the list
  Else
      For intIndex = 6 To 17
          mnuOptCharLength(intIndex).Checked = False
          mnuOptCharLength(intIndex).Visible = False
      Next
      intItems = 5  ' number of items on the list
  End If
  
' ---------------------------------------------------------------------------
' Set the appropriate check mark
' ---------------------------------------------------------------------------
  For intIndex = 0 To intItems
      If intIndex = Index Then
          mnuOptCharLength(intIndex).Checked = True
          m_intWordLength = intIndex + 3      ' calc the word length
          
          If Not m_bAlphabetic Then
              cboSpecial_Click 1       ' refill letters to convert box
          End If
      Else
          ' remove all other check marks
          mnuOptCharLength(intIndex).Checked = False
      End If
  Next
  
' ---------------------------------------------------------------------------
' fill characters per word into combo box
' ---------------------------------------------------------------------------
  If Not g_bCreate_Password Then
      
      cboLength.Clear      ' empty the combo box
      
      ' reload the combo box
      Select Case m_intWordLength
             Case 3:  ' 3 character length
                    cboLength.AddItem " All 3"
                    cboLength.AddItem " 3 to 4"
                    cboLength.AddItem " 3 to 5"
                    cboLength.AddItem " 3 to 6"
                    cboLength.AddItem " 3 to 7"
                    cboLength.AddItem " 3 to 8"
            
             Case 4:  ' 4 character length
                    cboLength.AddItem " All 4"
                    cboLength.AddItem " 4 to 5"
                    cboLength.AddItem " 4 to 6"
                    cboLength.AddItem " 4 to 7"
                    cboLength.AddItem " 4 to 8"
            
             Case 5:  ' 5 character length
                    cboLength.AddItem " All 5"
                    cboLength.AddItem " 5 to 6"
                    cboLength.AddItem " 5 to 7"
                    cboLength.AddItem " 5 to 8"
            
             Case 6:  ' 6 character length
                    cboLength.AddItem " All 6"
                    cboLength.AddItem " 6 to 7"
                    cboLength.AddItem " 6 to 8"
            
             Case 7:  ' 7 character length
                    cboLength.AddItem " All 7"
                    cboLength.AddItem " 7 to 8"
            
             Case 8:  ' 8 character length
                    cboLength.AddItem " All 8"
      End Select
      
      ' highlight the first item in the combo box
      cboLength.ListIndex = 0
  End If
  
End Sub

Private Sub mnuFileExit_Click()

' ---------------------------------------------------------------------------
' Shutdown this application
' ---------------------------------------------------------------------------
  cmdChoice_Click 1

End Sub

Private Sub mnuOptOmit_Click()

' ***************************************************************************
' Routine:       mnuOptOmit_Click
'
' Description:   display an input box with a list of special characters to
'                be omitted
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strMsg          As String
  Dim strNewList      As String
  Dim strCurrentList  As String
  Dim intIndex1       As Integer
  Dim intIndex2       As Integer
  Dim intTemp         As Integer
  Dim bFoundIt        As Boolean
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strMsg = "Enter the characters to be omitted on the line below.  "
  strMsg = strMsg & "Leaving the line blank will default to none of this characters "
  strMsg = strMsg & "being omitted.  Pressing CANCEL will empty the line." & vbCrLf & vbCrLf
  strMsg = strMsg & Space(5) & " ~ ! @ # $ % ^ & * ( ) _ - + = { [ } ] | \ : ; < , > . ? /"
  strNewList = ""
  strCurrentList = ""
    
' ---------------------------------------------------------------------------
' If there are any chars in the list convert back to ASCII
' display values
' ---------------------------------------------------------------------------
  If m_intOmitCnt > 0 Then
      For intIndex1 = 1 To UBound(g_arOmit)
          strCurrentList = strCurrentList & Chr(g_arOmit(intIndex1)) & " "
      Next
  End If

' ---------------------------------------------------------------------------
' Empty the array
' ---------------------------------------------------------------------------
  Erase g_arOmit()
  m_intOmitCnt = 0
  
' ---------------------------------------------------------------------------
' Display message, title, and default value.
' ---------------------------------------------------------------------------
  strNewList = InputBox(strMsg, "Characters to omit", strCurrentList)
  
' ---------------------------------------------------------------------------
' see if anything was entered on the line.  No duplicates
' allowed.
' ---------------------------------------------------------------------------
  strNewList = Trim(strNewList)
  If Len(strNewList) > 0 Then
      
      For intIndex1 = 1 To Len(strNewList)
      
          ' make sure this is not a blank space
          If Mid(strNewList, intIndex1, 1) <> Chr(32) Then
              
              ' Capture one char at a time and convert to decimal value
              intTemp = Asc(Mid(strNewList, intIndex1, 1))
              
              ' make sure it is a special character
              Select Case intTemp
                     Case 33, 35 To 38, 40 To 47, 58 To 64, 91 To 95, 123 To 126
                          
                          bFoundIt = False
                          
                          ' look for duplicates
                          If g_arOmit(1) > 0 Then            ' anything in the array?
                              For intIndex2 = 1 To UBound(g_arOmit)
                                  If g_arOmit(intIndex2) > 0 Then    ' anything in this element?
                                      ' is this item in this array element?
                                      If intTemp = g_arOmit(intIndex2) Then
                                          bFoundIt = True    ' we have a match
                                          Exit For           ' exit loop
                                      End If
                                  Else
                                      Exit For  ' found empty element, exit loop
                                  End If
                              Next
                          End If
                          
                          ' if not a dupe then add it to the array
                          If Not bFoundIt Then
                              m_intOmitCnt = m_intOmitCnt + 1  ' increment counter
                              g_arOmit(m_intOmitCnt) = intTemp    ' add to the array
                          End If
              End Select
          End If
      Next
  End If

End Sub

Private Sub mnuFileOpen_Click()

' ---------------------------------------------------------------------------
' Set CancelError is True
' ---------------------------------------------------------------------------
  cmDialog.CancelError = True
  
  On Error GoTo ErrHandler
  
' ---------------------------------------------------------------------------
' Setup and display the "FILE OPEN" dialog box
' ---------------------------------------------------------------------------
  With cmDialog
       .Flags = cdlOFNHideReadOnly Or _
                cdlOFNExplorer Or _
                cdlOFNLongNames Or _
                cdlOFNFileMustExist
       .FileName = ""
       ' Set filters
       .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
       .FilterIndex = 1             ' Specify default filter
       .ShowOpen                    ' Display the Open dialog box
  End With
  
' ---------------------------------------------------------------------------
' Capture name of selected file
' ---------------------------------------------------------------------------
  m_strFilename = cmDialog.FileName
  If Len(Trim(m_strFilename)) = 0 Then Exit Sub
  
' ---------------------------------------------------------------------------
' Start notepad to review this file
' ---------------------------------------------------------------------------
  Shell "notepad.exe " & m_strFilename, vbNormalFocus
  Exit Sub
  
ErrHandler:

' ---------------------------------------------------------------------------
' User pressed the Cancel button
' ---------------------------------------------------------------------------
  Exit Sub

End Sub

Private Sub mnuFilePrint_Click()

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTitleLine  As String
  Dim strTmp1       As String
  Dim strTmp2       As String
  Dim intPosition   As Integer
  Dim cPrint        As clsPrint
  
  Set cPrint = New clsPrint
  
' ---------------------------------------------------------------------------
' Capture and test to see if we have any data
' ---------------------------------------------------------------------------
  If g_bCreate_Password Then
      If Len(g_strPassphrase) = 0 Then
          Exit Sub
      Else
          strTitleLine = Format(m_lngCount, "#,0") & " Passwords " & _
                         Format(m_intWordLength, "#0") & " characters long"
          m_strTextData = g_strPassphrase
      End If
  Else
      If Len(Trim(txtPassphrase)) = 0 Then
          Exit Sub
      Else
          strTitleLine = "Passphrase"
          m_strTextData = Trim(txtPassphrase.Text)
      End If
  End If
    
' ---------------------------------------------------------------------------
' See if we have data to print
' ---------------------------------------------------------------------------
  If Len(Trim(m_strTextData)) = 0 Then
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' Set Cancel to True.
' ---------------------------------------------------------------------------
  cmDialog.CancelError = True

On Error GoTo ErrHandler
   
' ---------------------------------------------------------------------------
' Display the "Print" dialog box
' ---------------------------------------------------------------------------
  With cmDialog
       ' default to print all pages, no saving to a file,
       ' and no selective printing
       .Flags = cdlPDAllPages Or _
                cdlPDHidePrintToFile Or _
                cdlPDNoSelection
       .ShowPrinter   ' Display the Print dialog box
  End With
  
' ---------------------------------------------------------------------------
' change the curser to an hourglass
' ---------------------------------------------------------------------------
  Screen.MousePointer = vbHourglass
  
' ---------------------------------------------------------------------------
' Print the data
' ---------------------------------------------------------------------------
  cPrint.PrintText strTitleLine, m_strTextData
  Screen.MousePointer = vbNormal

ErrHandler:

' ---------------------------------------------------------------------------
' User pressed Cancel button.
' ---------------------------------------------------------------------------
  Set cPrint = Nothing
  Exit Sub

End Sub

Private Sub mnuOptPWD_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intIndex As Integer
  
' ---------------------------------------------------------------------------
' Set the appropriate check mark
' ---------------------------------------------------------------------------
  For intIndex = 0 To 1
      If intIndex = Index Then
          mnuOptPWD(intIndex).Checked = True
      Else
          ' remove all other check marks
          mnuOptPWD(intIndex).Checked = False
      End If
  Next
  
' ---------------------------------------------------------------------------
' Set the flag
' ---------------------------------------------------------------------------
  If Index = 0 Then
      g_bCreate_Password = True
      mnuOptCharLength_Click 5
  Else
      g_bCreate_Password = False
      mnuOptCharLength_Click 2
  End If

' ---------------------------------------------------------------------------
' update the screen
' ---------------------------------------------------------------------------
  Update_Screen Index

End Sub

Private Sub mnuFileSaveAs_Click()

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim hFile        As Integer
  Dim intPosition  As Integer
  Dim strName      As String
  Dim strTmpLine   As String
  
' ---------------------------------------------------------------------------
' Capture and test to see if we have any data
' ---------------------------------------------------------------------------
  If g_bCreate_Password Then
      If Len(Trim(g_strPassphrase)) = 0 Then
          Exit Sub
      Else
          m_strTextData = g_strPassphrase
      End If
  Else
      If Len(Trim(txtPassphrase)) = 0 Then
          Exit Sub
      Else
          m_strTextData = txtPassphrase
      End If
  End If

' ---------------------------------------------------------------------------
' See if we have data to save
' ---------------------------------------------------------------------------
  If Len(Trim(m_strTextData)) = 0 Then
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' Set CancelError is True
' ---------------------------------------------------------------------------
  cmDialog.CancelError = True
  
  On Error GoTo ErrHandler
  
' ---------------------------------------------------------------------------
' Display the "FILE SAVE AS" dialog box
' ---------------------------------------------------------------------------
  With cmDialog
       ' Set flags
       .Flags = cdlOFNExplorer Or _
                cdlOFNLongNames Or _
                cdlOFNHideReadOnly Or _
                cdlOFNOverwritePrompt
       
       .FileName = ""               ' empty filename selection
       
       ' Set filters
       .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
       .FilterIndex = 1             ' Specify default filter
       .ShowOpen                    ' Display File Open dialog box
       
       m_strFilename = cmDialog.FileName    ' capture filename
  End With
  
' ---------------------------------------------------------------------------
' see if a file name was entered or selected
' ---------------------------------------------------------------------------
  If Len(Trim(m_strFilename)) = 0 Then
      Exit Sub
  Else
      ' save just the filename
      intPosition = InStrRev(m_strFilename, "\", Len(m_strFilename))
      strName = Mid(m_strFilename, intPosition + 1)
  End If
  
' ---------------------------------------------------------------------------
' save the file with a max line length of 70 (including carriage return and
' linefeed characters + 68 data characters)
' ---------------------------------------------------------------------------
  hFile = FreeFile
  Open m_strFilename For Output As #hFile
  Print #hFile, "Filename:  " & strName
  Print #hFile, "Created:   " & Format(Now(), "dddd  d mmmm yyyy  h:mm ampm")
  Print #hFile, " "
  Print #hFile, "Length of each word:  " & Format(m_intWordLength, "#0")
  Print #hFile, "Number of passwords:  " & Format(m_lngCount, "#,0")
  Print #hFile, String(70, 45)
  
  ' loop thru the long data string and format an output line
  ' with a max length of 65 characters
  Do
      ' if the data string is greater than 65 characters,
      ' it will have to be evaluated by parsing backwards
      ' for the first blank space
      If Len(m_strTextData) > 68 Then
          ' capture the first 69 characters of data
          strTmpLine = Left(m_strTextData, 69)
          
          ' parse backwards looking for the first blank space
          intPosition = InStrRev(strTmpLine, Chr(32), Len(strTmpLine))
              
          ' resize the data string up to that blank space
          strTmpLine = Trim(Left(strTmpLine, intPosition - 1))
            
          ' print to the output file.  By not using a semi-colon
          ' a chr(13)+chr(10) are automatically appended to the
          ' end of each line.
          Print #hFile, strTmpLine
          
          ' Resize the original data string minus the length
          ' of the final output string.  Remove leading blanks.
          m_strTextData = LTrim(Mid(m_strTextData, Len(strTmpLine) + 1))
      Else
          ' print whatever is left and
          ' exit this loop
          Print #hFile, Trim(m_strTextData)
          Exit Do
      End If
  Loop
  
  Close #hFile   ' close the output file
  
ErrHandler:
' ---------------------------------------------------------------------------
' User pressed the Cancel button
' ---------------------------------------------------------------------------
  Exit Sub

End Sub

Private Sub txtPassphrase_KeyDown(KeyCode As Integer, Shift As Integer)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim CtrlDown    As Integer
  Dim PressedKey  As Integer
  
' ---------------------------------------------------------------------------
' Initialize  variables
' ---------------------------------------------------------------------------
  CtrlDown = (Shift And vbCtrlMask) > 0   ' Define control key
  PressedKey = Asc(UCase(Chr(KeyCode)))   ' Convert to uppercase
    
' ---------------------------------------------------------------------------
' Check to see if it is okay to make changes.
' ---------------------------------------------------------------------------
  If CtrlDown And PressedKey = vbKeyX Then
      Edit_Cut         ' Ctrl + X was pressed
  ElseIf CtrlDown And PressedKey = vbKeyA Then
      With txtPassphrase ' Ctrl + A was pressed
           .SelStart = 0
           .SelLength = Len(.Text)
      End With
  ElseIf CtrlDown And PressedKey = vbKeyC Then
      Edit_Copy        ' Ctrl + C was pressed
  ElseIf CtrlDown And PressedKey = vbKeyV Then
      Edit_Paste       ' Ctrl + V was pressed
  ElseIf PressedKey = vbKeyDelete Then
      Edit_Delete      ' Delete key was pressed
  End If

End Sub

Private Sub Fill_Combo_Boxes()

' ***************************************************************************
' Routine:       Fill_Combo_Boxes
'
' Description:   Fill all the combo boxes with data
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim intIndex As Integer
  
' ---------------------------------------------------------------------------
' load all the combo boxes for read only access
' ---------------------------------------------------------------------------
  With frmMain
        
        ' number of words in phrase
        .cboNbrCount.Clear
        For intIndex = 1 To 14
            .cboNbrOfWords.AddItem " " & intIndex
        Next
        .cboNbrOfWords.ListIndex = 4

        ' Special character options
        For intIndex = 0 To 1
            .cboSpecial(intIndex).Clear
            .cboSpecial(intIndex).AddItem " 1-Alphabetic Only"
            .cboSpecial(intIndex).AddItem " 2-Numeric Mix"
            .cboSpecial(intIndex).AddItem " 3-Special Char Mix"
            .cboSpecial(intIndex).AddItem " 4-Numeric and Special"
            .cboSpecial(intIndex).ListIndex = 0
        Next
  End With
    
End Sub

Private Sub Format_PWD_Display()

' ***************************************************************************
' Routine:       Format_PWD_Display
'
' Description:   Formats the output display of the passwords for the screen
'
' Parameters:    lngCount - number of passwords to format
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' 30-NOV-2000  Kenneth Ives     Modified the output display
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim lngPosition    As Long
  Dim lngRowCount    As Long
  Dim lngIndex       As Long
  Dim lngWidth       As Long
  Dim lngWordCounter As Long
  Dim lngPolnger     As Long
  Dim lngChar2Conv   As Long
  Dim intWordsPerRow As Integer
  Dim lngLength      As Long
  Dim strTestWord    As String
  Dim strOutput      As String
  Dim arData()       As String
    
' ---------------------------------------------------------------------------
' initialize variables
' ---------------------------------------------------------------------------
  lngWordCounter = 0
  lngRowCount = 0

' ---------------------------------------------------------------------------
' Test password length
' ---------------------------------------------------------------------------
  If Len(Trim(g_strPassphrase)) = 0 Then
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' Format lngo the proper case
' ---------------------------------------------------------------------------
  Select Case g_intTypeCase
  
         Case 0: ' All Lowercase characters
              g_strPassphrase = StrConv(g_strPassphrase, vbLowerCase)
         
         Case 1: ' All Uppercase characters
              g_strPassphrase = StrConv(g_strPassphrase, vbUpperCase)
         
         Case 2: ' All Propercase (First character is uppercase)
              g_strPassphrase = StrConv(g_strPassphrase, vbProperCase)
         
         Case 3: ' Mixed case
              Erase g_arlngData()                      ' empty the array
              lngLength = Len(g_strPassphrase) - 1  ' get length of string
              g_strPassphrase = StrConv(g_strPassphrase, vbLowerCase)
              
              If lngLength < 5 Then       ' If string length is less than 5
                  lngChar2Conv = 2        '    only 2 chars to be converted
              Else
                  lngChar2Conv = Int(lngLength \ 2) ' estimate half string length
                  If lngChar2Conv = 0 Then
                      lngChar2Conv = 1
                  End If
              End If
                
              ' randomly generate the position in the passphrase string
              ' as to which characters are to be converted to uppercase
              Create_Random_Pointers lngLength, g_arlngData(), True
              
              For lngIndex = 1 To lngChar2Conv
                  
                  lngPolnger = g_arlngData(lngIndex)
                  Mid(g_strPassphrase, lngPolnger, 1) = _
                      StrConv(Mid(g_strPassphrase, lngPolnger, 1), vbUpperCase)
              Next
  End Select
  
' ---------------------------------------------------------------------------
' Sort the data Ascending
' ---------------------------------------------------------------------------
  arData = Sort_Data(m_lngCount)
  g_strPassphrase = ""
  strOutput = ""
  lngLength = UBound(arData) - 1

' ---------------------------------------------------------------------------
' set up number of grid columns.  This was tedious.  There sometimes can be
' a difference in display between IDE and the compiled version.  Be sure to
' thoroughly check both.  Been there and it is not fun.  :-)
' ---------------------------------------------------------------------------
  Select Case m_intWordLength
         Case 3:        m_intColumns = 9
         Case 4:        m_intColumns = 8
         Case 5:        m_intColumns = 6
         Case 6 To 7:   m_intColumns = 5
         Case 8 To 9:   m_intColumns = 4
         Case 10 To 13: m_intColumns = 3
         Case Else:     m_intColumns = 2
  End Select

' ---------------------------------------------------------------------------
' Prepare the grid
' ---------------------------------------------------------------------------
  grdPasswords.Rows = 0                      ' number of current rows
  grdPasswords.Cols = m_intColumns           ' number of columns
  strTestWord = String(m_intWordLength, "M") ' create test word
  lngWidth = TextWidth(strTestWord) + 200    ' calc the column width
  
  For lngIndex = 0 To m_intColumns - 1
      grdPasswords.ColWidth(lngIndex) = 0        ' set min column width
      grdPasswords.ColWidth(lngIndex) = lngWidth ' adjust column width
  Next
  
  grdPasswords.Refresh    ' refresh the grid
  
  ' Temporarily lock the grid control while loading.
  ' This will speed things up and reduce the amount of
  ' flicker
  LockWindowUpdate grdPasswords.hWnd
  
' ---------------------------------------------------------------------------
' loop thru and build the password display output
' ---------------------------------------------------------------------------
  For lngIndex = 0 To lngLength
             
      ' append password to output string and then
      ' append a TAB as a delimiter
      strOutput = strOutput & arData(lngIndex) & vbTab
      
      If Len(Trim(arData(lngIndex))) = 0 Or _
         lngIndex = lngLength Then
            ' get rid of the trailing tab
            strOutput = Trim(strOutput)
            
            ' if there are any passwords on the
            ' output line then place them in the
            ' grid one row at a time
            If Len(strOutput) > 0 Then
                grdPasswords.AddItem strOutput, lngRowCount
            End If
            
            Exit For    ' exit this loop
      End If
          
      lngWordCounter = lngWordCounter + 1
      
      If lngWordCounter = m_intColumns Then
          lngWordCounter = 0                          ' reset word counter
          strOutput = Trim(strOutput)                 ' remove trailing tab
          grdPasswords.AddItem strOutput, lngRowCount ' Place data in the grid
          lngRowCount = lngRowCount + 1               ' update the row position
          
          ' Replace tabs with four spaces, linefeeds with 2
          ' carriage returns with 2.  The latter are always
          ' next to each other
          strOutput = Trim(Replace(strOutput, Chr(9), String(4, 32)))
          strOutput = Trim(Replace(strOutput, Chr(10), String(2, 32)))
          strOutput = Trim(Replace(strOutput, Chr(13), String(2, 32)))
          
          ' save to master string.  User may want to save to a file.
          g_strPassphrase = g_strPassphrase & strOutput & String(4, 32)
          
          ' empty output string
          strOutput = ""
      End If
      
      ' if the stop button was pressed then leave
      DoEvents
      If g_bStop_Pressed Then
          Exit For
      End If
  Next
        
' ---------------------------------------------------------------------------
' unlock the grid control after we have finished loading the data into it.
' ---------------------------------------------------------------------------
  LockWindowUpdate 0&
  
End Sub

Private Sub txtQuantity_GotFocus()

' ---------------------------------------------------------------------------
' Highlight all the text in the box
' ---------------------------------------------------------------------------
  txtQuantity.SelStart = 0
  txtQuantity.SelLength = Len(txtQuantity.Text)

End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)

' ---------------------------------------------------------------------------
' Allow only numbers and backspace to be entered
' ---------------------------------------------------------------------------
  Select Case KeyAscii
         Case 9, 13: KeyAscii = 0
         Case 8, 48 To 57: Exit Sub
         Case Else: KeyAscii = 0
  End Select
  
End Sub

Private Sub txtQuantity_LostFocus()
  
' ---------------------------------------------------------------------------
' Minimum value of one and maximum value of 1000
' ---------------------------------------------------------------------------
  If Val(Trim(txtQuantity)) > 2500 Then
      MsgBox "Maximum value allowed is 2500", vbOKOnly, _
             "Quantity error"
  End If
  
End Sub

Public Sub Create_Special_Word()

' ***************************************************************************
' Routine:       Create_Special_Word
'
' Description:   Build a special word for a passphrase.  It can be all
'                numeric, all special characters, or a mix of both.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex       As Long
  Dim lngLength      As Long
  Dim strChar        As String
  Dim bNumeric_Only  As Boolean
  
' ---------------------------------------------------------------------------
' empty the arrays
' ---------------------------------------------------------------------------
  Erase g_arWords()
  Erase g_arlngData()
  bNumeric_Only = False
                  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  g_strSpecialWord = ""
  
' ---------------------------------------------------------------------------
' Determine the special word length
' ---------------------------------------------------------------------------
  If g_intMaxLength = 0 Then
      lngLength = g_intMinLength ' all the same length selected
  Else
      lngLength = g_intMaxLength - 1 ' multiple lengths
  End If
                  
' ---------------------------------------------------------------------------
' See if the user wants a special word created
' ---------------------------------------------------------------------------
  Select Case Trim(cboSpecial(0).Text)
         
         Case "1-Alphabetic Only"
              g_strSpecialWord = ""
         
         Case "2-Numeric Mix": ' create a maximum length word with numbers
              For lngIndex = 1 To lngLength
                  g_strSpecialWord = g_strSpecialWord & CStr(Int(Rnd * 9))
              Next
         
         Case "3-Special Char Mix":
                  
              ' if the user omitted all the special characters then leave
              If m_intOmitCnt = 29 Then
                  m_lngNbrCount = 0
                  MsgBox "User has omitted all special characters." & vbCrLf & _
                         "Either update the omitted characters list" & vbCrLf & _
                         "or make another selection.", vbInformation + vbOKOnly, _
                         "No data available"
                         
              ' create a maximum length word with special characters
              Else
                  ' all we want are the special characters
                  g_strSpecialWord = Do_Special_Only(False, lngLength)
              End If
              
         Case "4-Numeric and Special":
                  
              ' if the user omitted all the special
              ' characters then use numbers only
              If m_intOmitCnt = 29 Then
                  bNumeric_Only = True
                  MsgBox "User has omitted all special characters." & vbCrLf & _
                         "Only numreic values will be used.", _
                         vbInformation + vbOKOnly, "No data available"
              End If
                  
              If bNumeric_Only Then
                  For lngIndex = 1 To lngLength
                      g_strSpecialWord = g_strSpecialWord & CStr(Int(Rnd * 9))
                  Next
              Else
                  ' One word with numbers and special characters
                  g_strSpecialWord = Do_Special_Only(True, lngLength)
              End If
  End Select
  
' ---------------------------------------------------------------------------
' If a special or numeric word was selected then determine its
' position in the passphrase string
' ---------------------------------------------------------------------------
  If Len(Trim(g_strSpecialWord)) > 0 Then
      Create_Random_Pointers g_lngNbrOfWords, g_arlngData(), True
      g_intPosition = g_arlngData(0) ' one word, one position in the phrase
  End If

End Sub

Public Sub Create_Passwords()

' ***************************************************************************
' Routine:       Create_Passwords
'
' Description:   Builds the passwords
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim intIndex     As Integer
  Dim intPointer   As Integer
  Dim lngIndex     As Long
  Dim lngTotChar   As Long
  Dim lngLength    As Long
  Dim strOnePWord  As String
  Dim strRndData   As String
  Dim strSpecial   As String
  Dim cRND         As clsRndData
  
' ---------------------------------------------------------------------------
' initialize local variables
' ---------------------------------------------------------------------------
  g_strPassphrase = ""
  strRndData = ""
  
' ---------------------------------------------------------------------------
' Evaluate the count
' ---------------------------------------------------------------------------
  If Len(Trim(txtQuantity.Text)) = 0 Or _
     Val(Trim(txtQuantity.Text)) = 0 Then
         txtQuantity.Text = "0"
         Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' read the combo box
' ---------------------------------------------------------------------------
  m_lngCount = Val(Trim(txtQuantity.Text))
  m_lngNbrCount = Val(Trim(cboNbrCount.Text))
  lngTotChar = m_lngCount * m_intWordLength
  
' ---------------------------------------------------------------------------
' Build all the passwords in one long string (all alphabetic).  This is the
' fastest way possible.  Use upper and lower case.  we will convert later.
' ---------------------------------------------------------------------------
  Set cRND = New clsRndData
  strRndData = cRND.Build_Random_Data(lngTotChar, False, True, True)
  Set cRND = Nothing
  
' ---------------------------------------------------------------------------
' if stop button pressed then leave
' ---------------------------------------------------------------------------
  DoEvents
  If g_bStop_Pressed Then
      g_strPassphrase = ""
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' parse the random data string and process one password at a
' time
' ---------------------------------------------------------------------------
  For lngIndex = 1 To lngTotChar Step m_intWordLength
      
      strOnePWord = Left(strRndData, m_intWordLength)   ' grab one password
      strRndData = Mid(strRndData, m_intWordLength + 1) ' resize what is left
      
      Erase g_arlngData()             ' empty the array
      lngLength = m_intWordLength - 1 ' calc 1 short of the password length
          
      ' See if the user wants passwords mixed with
      ' numbers and/or special characters
      Select Case Trim(cboSpecial(1).Text)
         
             Case "1-Alphabetic Only": ' Just fall thru. Nothing to do.
             
             ' build the passwords using numbers and letters
             Case "2-Numeric Mix":
                  
                  If m_lngNbrCount > 0 Then
                      ' randomly generate the position in the passphrase
                      ' string as to which characters are to be converted
                      Create_Random_Pointers lngLength, g_arlngData(), False
                      
                      ' change chars in the password string
                      For intIndex = 1 To m_lngNbrCount
                          strSpecial = Trim(CStr(Int(Rnd * 9))) ' create 0-9
                          intPointer = g_arlngData(intIndex)    ' id the position
                          ' replace the char in the password
                          Mid(strOnePWord, intPointer, 1) = strSpecial
                      Next
                  End If
                  
                  ' if stop button pressed then leave
                  DoEvents
                  If g_bStop_Pressed Then
                      g_strPassphrase = ""
                      Exit For
                  End If
                      
             ' build into the password special characters
             ' except for apostrophes and quotes
             Case "3-Special Char Mix":
                         
                  ' if user omitted all the special characters then leave
                  If m_intOmitCnt = 29 Then
                      m_lngNbrCount = 0
                      
                  ElseIf m_lngNbrCount > 0 Then
                      ' randomly generate the position in the passphrase
                      ' string as to which characters are to be converted
                      Create_Random_Pointers lngLength, g_arlngData(), False
                    
                      strSpecial = Do_Special_Only(False, m_lngNbrCount)  ' Special char only
                      
                      For intIndex = 1 To m_lngNbrCount
                          intPointer = g_arlngData(intIndex) ' ID the position
                          ' replace the char in the password
                          Mid(strOnePWord, intPointer, 1) = Mid(strSpecial, intIndex, 1)
                      Next
                  End If
                  
                  ' if stop button pressed then leave
                  DoEvents
                  If g_bStop_Pressed Then
                      g_strPassphrase = ""
                      Exit For
                  End If
             
             ' build the password using a mix of printable
             ' characters, numbers and alphabetic characters
             Case "4-Numeric and Special":
                 
                  ' if the user omitted all the special
                  ' characters then create numeric only
                  If m_intOmitCnt = 29 Then
                  
                      ' this will be numeric only
                      If m_lngNbrCount > 0 Then
                          ' randomly generate the position in the passphrase
                          ' string as to which characters are to be converted
                          Create_Random_Pointers lngLength, g_arlngData(), False
                      
                          ' change chars in the password string
                          For intIndex = 1 To m_lngNbrCount
                              strSpecial = Trim(CStr(Int(Rnd * 9))) ' create 0-9
                              intPointer = g_arlngData(intIndex)    ' id the position
                              ' replace the char in the password
                              Mid(strOnePWord, intPointer, 1) = strSpecial
                          Next
                      End If
                      
                  ElseIf m_lngNbrCount > 0 Then
                      ' randomly generate the position in the passphrase
                      ' string as to which characters are to be converted
                      Create_Random_Pointers lngLength, g_arlngData(), False
                      
                      strSpecial = Do_Special_Only(True, m_lngNbrCount)  ' Special char/Number
                      
                      For intIndex = 1 To m_lngNbrCount
                          intPointer = g_arlngData(intIndex) ' ID the position
                          ' replace the char in the password
                          Mid(strOnePWord, intPointer, 1) = Mid(strSpecial, intIndex, 1)
                      Next
                  End If
                  
                  ' if stop button pressed then leave
                  DoEvents
                  If g_bStop_Pressed Then
                      g_strPassphrase = ""
                      Exit For
                  End If
      End Select
            
      ' append one password at a time to the output string
      ' separated by a blank space
      g_strPassphrase = g_strPassphrase & strOnePWord & Chr(32)
  
  Next
          
' ---------------------------------------------------------------------------
' if stop button pressed then leave
' ---------------------------------------------------------------------------
  DoEvents
  If g_bStop_Pressed Then
      g_strPassphrase = ""
      Exit Sub
  End If
         
End Sub

