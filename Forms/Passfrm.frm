VERSION 5.00
Begin VB.Form PassFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logon - TourWin Version 1.00"
   ClientHeight    =   2475
   ClientLeft      =   2865
   ClientTop       =   2505
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Passfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2475
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   Begin VB.Frame Frame1 
      Caption         =   "User Name and  Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Tag             =   "14"
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cboPasdb 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox PasPasTxt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox PasNamTxt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblPasDB 
         Alignment       =   1  'Right Justify
         Caption         =   "&Database :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label PasPasLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "&Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&User Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   6
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.CommandButton PasCanCmd 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton PasOkCmd 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Tag             =   "576"
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "PassFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const INVALIDNAME = 1
Const INVALIDPWD = 2
Const INVALIDDATAPATH = 3
Const OK = 0
Const BADNAME = 1
Const BADPASSWORD = 2
Const FOUNDANDOPENED = 0
Const CREATEDANDOPENED = 1
Const NOTFOUNDNORCREATED = 2


Event LoginSubmitted(ByVal sName As String, ByVal sPWD As String, ByVal sDatapath As String)

Private Sub Form_Load()
On Local Error GoTo PassForm_Err

    LoadFormResourceString Me

On Local Error GoTo 0
Exit Sub

PassForm_Err:
If bDebug Then
   Handle_Err Err, "Form_Load-PassForm"
End If
Resume Next
End Sub


Private Sub PasCanCmd_Click()

    Unload Me
    
End Sub


Private Sub PasNamTxt_GotFocus()

With PasNamTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub PasNamTxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = Val(vbTab) Then
    KeyAscii = 0
    PasOkCmd.SetFocus
    PasOkCmd_Click
End If
End Sub



Private Sub PasOkCmd_Click()

On Local Error GoTo OK_Error

If bDebug Then Handle_Err 0, "Login ... PasOKCmd-PassForm"

RaiseEvent LoginSubmitted(PasNamTxt.Text, PasPasTxt.Text, cboPasdb.Text)
DoEvents

On Error GoTo 0
Exit Sub

OK_Error:
If bDebug Then
    Handle_Err Err, "PasOKCmd-PassForm"
End If
Resume Next

End Sub


Private Sub PasPasTxt_GotFocus()

With PasPasTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub PasPasTxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    PasOkCmd.SetFocus
    PasOkCmd_Click
End If
End Sub


'---------------------------------------------------------------------------------------
' PROCEDURE : BadLogin
' DATE      : 6/30/04 18:47
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub BadLogin(ByVal lErrorType As Long)

On Local Error GoTo BadLogin_Error
'Declare local variables


Select Case lErrorType

       Case INVALIDNAME:
                MsgBox LoadResString(gcPassFrmUserName), vbOKOnly + vbExclamation, LoadResString(gcTourVersion)
                PasNamTxt.SetFocus

                
       Case INVALIDPWD:
                MsgBox "Incorrect user password.", vbOKOnly + vbExclamation, LoadResString(gcTourVersion)
                PasPasTxt.SetFocus

                
       Case INVALIDDATAPATH:
                MsgBox "Invalid database"
                cboPasdb.SetFocus

End Select


On Error GoTo 0
Exit Sub

BadLogin_Error:
If bDebug Then Handle_Err Err, "BadLogin-PassFrm"
Resume Next

End Sub
