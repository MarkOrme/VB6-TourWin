VERSION 5.00
Begin VB.Form frmRegistration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   " Registration Number "
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtKey 
         Height          =   285
         Index           =   3
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Index           =   2
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)

' Declare local variables
Dim sKey            As String

Select Case Index
    Case 0: ' OK button
    sKey = Trim(txtKey(0).Text) & Trim(txtKey(1).Text) _
        & Trim(txtKey(2).Text) & Trim(txtKey(3).Text)
            
    If cLicense.IsKeyOK(sKey) Then
        MsgBox "Registration successful, thank you and enjoy!", vbOKOnly & vbInformation, LoadResString(gcTourVersion)
        cLicense.SetSystemRegistrationKey (sKey)
        cLicense.Licensing_Results = 2 ' Successful registration ...
        Unload Me
    Else
        MsgBox "Key is NOT valid, please try again.", vbOKOnly & vbInformation, LoadResString(gcTourVersion)
    End If

    Case 1: ' Cancel button
      cLicense.Licensing_Results = 3 ' Failed on Registration...
      Unload Me
End Select
End Sub

Private Sub txtKey_GotFocus(Index As Integer)

' Highlight field
With txtKey(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub
