VERSION 5.00
Begin VB.Form frmItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "####"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "950"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "####"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Tag             =   "952"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "####"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Tag             =   "951"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "####"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Tag             =   "953"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   240
      MaxLength       =   100
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblDescription 
      Caption         =   "####"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Tag             =   "954"
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_UserCancel As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColor_Click()

On Local Error GoTo ColourSelection_Cancelled

Dim sColor As SelectedColor


sColor = ShowColor(Me.hWnd)

If Not sColor.bCanceled Then
    txtDescription.BackColor = sColor.oSelectedColor
End If
On Error GoTo 0

Exit Sub
ColourSelection_Cancelled:
Err.Clear

End Sub

Private Sub cmdOK_Click()

If "" = txtDescription.Text Then
    MsgBox "Blank description not support!", vbOKOnly + vbInformation
    Exit Sub
End If

Me.USERCANCELLED = False
Me.Hide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Me.Hide
End If
End Sub

Private Sub Form_Load()

' Load Generic Resources
LoadFormResourceString Me
Me.USERCANCELLED = True
Me.KeyPreview = True

End Sub

Public Property Let MaxDescLength(ByVal vNewValue As Long)
    If vNewValue > 0 And vNewValue < 101 Then
        txtDescription.MaxLength = vNewValue
    End If
End Property

Public Property Get ColourButtonEnabled() As Boolean
    ColourButtonEnabled = cmdColor.Enabled
End Property

Public Property Let ColourButtonEnabled(ByVal vNewValue As Boolean)
    cmdColor.Enabled = vNewValue
    cmdColor.Visible = vNewValue
End Property

Public Property Get USERCANCELLED() As Boolean
    USERCANCELLED = m_UserCancel
End Property

Public Property Let USERCANCELLED(ByVal vNewValue As Boolean)
    m_UserCancel = vNewValue
End Property

Private Sub txtDescription_GotFocus()

With txtDescription
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub
