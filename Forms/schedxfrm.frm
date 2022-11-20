VERSION 5.00
Begin VB.Form schedxfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repeat Schedule"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "schedxfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOksch 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox lstdate 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Please select the desired start date."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "schedxfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    UserCancel = True
    Unload Me
End Sub

Private Sub cmdOksch_Click()
objMdi.info.iInterations = lstdate.ListIndex + 1
Unload Me
End Sub

Private Sub Form_Load()
Dim iLoop As Integer
Dim iForwardOrBackward As Integer
' Determine if schedule is prior
' to today's date or forward...
iForwardOrBackward = -1
If objMdi.info.bStartOf Then
    iForwardOrBackward = 1
    Label1.Caption = "Please select the desired end date."
End If


CentreForm schedxfrm, 0
For iLoop = 1 To 50
If -1 = iForwardOrBackward Then
lstdate.AddItem "Interation - " & Format$(iLoop, "00") & " --> " & _
                                  Format$(DateAdd("d", (iForwardOrBackward * objMdi.info.iInterationLen * iLoop) + 1, objMdi.info.dInterationDate), "mm-dd-yyyy")
Debug.Print "Interation - " & Format$(iLoop, "00") & " --> " & _
                                  Format$(DateAdd("d", (iForwardOrBackward * objMdi.info.iInterationLen * iLoop) + 1, objMdi.info.dInterationDate), "mm-dd-yyyy")
Else
lstdate.AddItem "Interation - " & Format$(iLoop, "00") & " --> " & _
                                  Format$(DateAdd("d", (iForwardOrBackward * objMdi.info.iInterationLen * iLoop), objMdi.info.dInterationDate), "mm-dd-yyyy")
End If
Next iLoop
If lstdate.ListCount > 0 Then
    'Select the first item in list
    lstdate.ListIndex = 0
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'vbFormControlMenu
If UnloadMode = vbFormControlMenu Then
    UserCancel = True
End If
End Sub
