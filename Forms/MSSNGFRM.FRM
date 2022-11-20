VERSION 5.00
Begin VB.Form MssngFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Location"
   ClientHeight    =   2475
   ClientLeft      =   1965
   ClientTop       =   1650
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MSSNGFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2475
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton MssCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton MssOkCmd 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox MssDatTxt 
      Height          =   285
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.DriveListBox MssDrive 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.DirListBox MssDirLst 
      Height          =   1215
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "D&rive:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "&Directory:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Location:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "MssngFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bLoading    As Boolean   ' Used in Activate and change event...

Private Sub Form_Activate()
On Local Error Resume Next
bLoading = True
' first set drive...
MssDrive.Drive = Mid$(MssDatTxt.Text, 1, 2)
MssDirLst.Path = MssDrive.Drive

' Now set path
MssDirLst.Path = MssDatTxt.Text

bLoading = False
End Sub

Private Sub MssCanCmd_Click()
Dim Msg As String, RetInt As Integer

 Unload MssngFrm
 
End Sub


Private Sub MssDatTxt_GotFocus()

With MssDatTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub MssDirLst_Change()

If Not bLoading Then
    MssDatTxt.Text = MssDirLst.Path
End If

End Sub


Private Sub MssDrive_Change()
On Local Error GoTo MssDri_Err

MssDirLst.Path = MssDrive.Drive

Exit Sub
MssDri_Err:
    MssDrive.Drive = MssDirLst.Path
End Sub


Private Sub MssOkCmd_Click()
Dim LenofTxt As Integer
On Local Error GoTo Mss_Err

' ---------------------------
' Must check if path is valid
' as user is allowed to enter
' location on their own
' ---------------------------

On Local Error GoTo Directory_Error
' -------------------------
' remove \ at end of string
' -------------------------
LenofTxt = Len(MssDatTxt)
If Mid$(MssDatTxt, LenofTxt, 1) = "\" Then MssDatTxt = Mid$(MssDatTxt, 1, LenofTxt - 1)
ChDir MssDatTxt
dataPath = MssDatTxt
Unload MssngFrm
Exit Sub
Directory_Error:
    Select Case Err
        Case 76:  ' Error 76 is Directory does not exist
           If vbYes = MsgBox("Invalid directory. Directory does not exist, " _
                   & "do you wish to create directory?", vbYesNo, LoadResString(gcTourVersion)) Then MakeDirectory MssDatTxt
                                                                    ' bug, MkDir creates only one new
                                                                    ' directory, must parse string
                                                                    ' and create individual directories.
            MssDatTxt.SetFocus
            MssDirLst.Path = MssDatTxt.Text
            Exit Sub
        Case Else
            If bDebug Then Handle_Err Err, "MssOkCmd-MssngFrm"
            Resume Next
    End Select
Exit Sub
Mss_Err:
    If bDebug Then Handle_Err Err, "MssOkCmd-MssngFrm"
    Resume Next
End Sub


