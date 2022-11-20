VERSION 5.00
Begin VB.Form DailyOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Options."
   ClientHeight    =   2610
   ClientLeft      =   3405
   ClientTop       =   1980
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DAILYOPT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2610
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton DaOCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton DaiSavCmd 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame DaOFrm 
      Caption         =   "Fields to track."
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   518
      Width           =   4215
      Begin VB.CheckBox DaOKmChk 
         Caption         =   "&Milege in Metric"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox DaOWeiChk 
         Caption         =   "&Weight Field"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox DaOdayChk 
         Caption         =   "Da&y Rate Descriptor"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox DaOHeaChk 
         Caption         =   "Resting &Heart Rate."
         Height          =   195
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox DaODesChk 
         Caption         =   "&Description"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox DaOSlpChk 
         Caption         =   "S&leep Field"
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   2040
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select which fields are visible in daily diary form."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "DailyOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bChanged As Boolean


Sub Set_Check_Setting()
On Local Error GoTo Check_Err

    UpdateDailyObject
        
Exit Sub
Check_Err:
    If bDebug Then Handle_Err Err, "Set_Check_Setting-DailyOpt"
    Resume Next
End Sub


Private Sub DaiSavCmd_Click()

If bChanged Then
    Set_Check_Setting
    ' ------------------------
    ' Update Daily Form
    ' to reflect user changes
    ' ------------------------
    DailyFrm.LoadVisible
    
End If
    Unload Me
End Sub

Private Sub DaOKmChk_Click()
bChanged = True
End Sub

Private Sub DaOCanCmd_Click()
    Unload Me
End Sub


Private Sub DaOdayChk_Click()
bChanged = True
End Sub

Private Sub DaODesChk_Click()
bChanged = True
End Sub

Private Sub DaOHeaChk_Click()
bChanged = True
End Sub

Private Sub DaOSlpChk_Click()
bChanged = True
End Sub

Private Sub DaOWeiChk_Click()
bChanged = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then

        If bChanged Then
            If vbYes = MsgBox("Save changes.", vbYesNo + vbInformation + vbDefaultButton2, gcTour_Win) Then
                Set_Check_Setting
                DailyFrm.LoadVisible
            End If
        End If

    Unload Me
End If
    
End Sub

Private Sub Form_Load()

    CentreForm DailyOpt, 0
    
    CheckVisibleControls
    Me.KeyPreview = True
    bChanged = False
    
End Sub

Sub CheckVisibleControls()

' -----------------------------
' Use the Daily Object to set
' DaO control values
' -----------------------------
With objMdi.info.DailyOpt
    DaOSlpChk.Value = .GetValue(BitFlags.Daily_Sleep)
    DaODesChk.Value = .GetValue(BitFlags.Daily_Description)
    DaOHeaChk.Value = .GetValue(BitFlags.Daily_RestHeart)
    DaOdayChk.Value = .GetValue(BitFlags.Daily_DayType)
    DaOWeiChk.Value = .GetValue(BitFlags.Daily_Weight)
    DaOKmChk.Value = .GetValue(BitFlags.Daily_Metric)
End With

End Sub

Sub UpdateDailyObject()

' -------------------------------------------------------
' UpDateDailyObject - Purpose, to update both database
'               and daily object to visiblity changes
' -------------------------------------------------------
With objMdi.info.DailyOpt
    
    ' Description
    .SetValue DaODesChk.Value, BitFlags.Daily_Description
    
    ' Sleep

    .SetValue DaOSlpChk.Value, BitFlags.Daily_Sleep
    
    ' Rest Heart
    .SetValue DaOHeaChk.Value, BitFlags.Daily_RestHeart
    
    ' Date Rate

    .SetValue DaOdayChk.Value, BitFlags.Daily_DayType
    
    ' Weight

    .SetValue DaOWeiChk.Value, BitFlags.Daily_Weight
    
    'Metric

    .SetValue DaOKmChk.Value, BitFlags.Daily_Metric

End With

objMdi.SaveUserSettings

End Sub
