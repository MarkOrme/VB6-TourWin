VERSION 5.00
Begin VB.Form frmTrialNotice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trial Notice"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTrialNotice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNotice 
      Caption         =   "Notice"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.Label lblRemaining 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblUsage 
         Caption         =   "This is a trial version limited to 60 opens. Click the Registration button to register."
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Later"
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Registration"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "frmTrialNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_UsageCount As Long
Private m_Results As Long

Public Property Get UsageCount() As Long
    UsageCount = m_UsageCount
End Property

Public Property Let UsageCount(ByVal lNewValue As Long)
    m_UsageCount = lNewValue
End Property

Private Sub cmdAction_Click(Index As Integer)

Select Case Index
  Case 0: ' Register
    frmRegistration.Show vbModal
    If 2 = cLicense.Licensing_Results Then
        Me.Hide
    End If
  Case 1: ' Later or Cancel
    cLicense.Licensing_Results = IIf(cmdAction(Index).Caption = "&Later", 0, 1)
    Me.Hide
End Select


End Sub

Private Sub Form_Activate()

If Me.UsageCount >= gcTRIALLENGTH Then
    lblUsage.Caption = "This trial version has expired. Click the Registration button to register."
    lblRemaining.Caption = ""
    cmdAction(1).Caption = "E&xit"    'Disable Later option...
Else
    lblRemaining.Caption = "There are " & CStr(gcTRIALLENGTH - Me.UsageCount) & " opens remaining."
End If

End Sub


Public Property Get NoticeResults() As Long
    NoticeResults = m_Results
End Property

Private Property Let NoticeResults(ByVal lNewValue As Long)
    m_Results = lNewValue
End Property
