VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Installation Options"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox chkSampleDB 
         Alignment       =   1  'Right Justify
         Caption         =   "Install Sample Database"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    ExitSetup Me, gintRET_EXIT
End Sub

Private Sub cmdFinish_Click()

gstrUserName = txtFirstName
5 gbInstallSampleDB = chkSampleDB.Value

Unload Me

End Sub

Private Sub Form_Load()

gstrUserName = ""
gbInstallSampleDB = False

cmdFinish.Caption = ResolveResString(resBTNINSTALL)
cmdExit.Caption = ResolveResString(resBTNEXIT)

End Sub
