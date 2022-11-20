VERSION 5.00
Begin VB.Form frmConcOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conconi Options"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConcOpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdaction 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdaction 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Startup Options"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox chkLoadOnStart 
         Caption         =   "Load last sessions settings"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmConcOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaction_Click(Index As Integer)
Const OK = 0
Dim iTemp As Integer

Select Case Index
       Case OK: ' OK Button
       
       objMdi.info.UserOptions.SetValue chkLoadOnStart.Value, BitFlags.Conconi_LoadSettings
       objMdi.SaveUserSettings

End Select

Unload Me

End Sub

Private Sub Form_Load()
On Local Error Resume Next

    chkLoadOnStart.Value = objMdi.info.UserOptions.GetValue(BitFlags.Conconi_LoadSettings)


End Sub
