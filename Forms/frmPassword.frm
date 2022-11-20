VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change password"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReEnter 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtNewPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Verify:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   850
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New password:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   370
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Password As String


Public Property Get Password() As String
 Password = m_Password
End Property

Public Property Let Password(ByVal vNewValue As String)
 m_Password = vNewValue
End Property

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
If txtNewPwd.Text = "" Or txtReEnter.Text = "" Then
    MsgBox "Blank password is not support. Please try again.", vbOKOnly + vbInformation, LoadResString(gcTourVersion)
    Exit Sub
End If
If txtNewPwd.Text <> txtReEnter.Text Then
    MsgBox "Passwords do not match. Please try again.", vbOKOnly + vbInformation, LoadResString(gcTourVersion)
    Exit Sub
End If
Me.Password = txtNewPwd.Text
Me.Hide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Me.Hide
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
End Sub

Private Sub txtReEnter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdOK.SetFocus
    cmdOK_Click
End If
End Sub
