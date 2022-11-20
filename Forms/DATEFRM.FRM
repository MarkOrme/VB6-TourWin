VERSION 5.00
Begin VB.Form DateFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#####"
   ClientHeight    =   1200
   ClientLeft      =   3450
   ClientTop       =   3075
   ClientWidth     =   3225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DATEFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1200
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "302"
   Begin Tourwin2002.UTextBox DatToTxt 
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "End date (mm-dd-yyyy)"
      Top             =   360
      Width           =   1300
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldName   =   ""
      Begin VB.CommandButton cmdCalendar 
         Height          =   255
         Index           =   1
         Left            =   1080
         Picture         =   "DATEFRM.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin Tourwin2002.UTextBox DatFroTxt 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Start date (mm-dd-yyyy)"
      Top             =   360
      Width           =   1300
      _ExtentX        =   2328
      _ExtentY        =   450
      Max             =   10
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldName   =   ""
      Begin VB.CommandButton cmdCalendar 
         Height          =   255
         Index           =   0
         Left            =   1080
         Picture         =   "DATEFRM.frx":052E
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.CommandButton DatOkCmd 
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton DatCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "####"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Tag             =   "301"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "####"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Tag             =   "300"
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "DateFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalendar_Click(Index As Integer)
Dim CMaxDate As New CCalendar
Dim lRt As Long
Dim lTop As Long
Dim lLeft As Long

Const USERCANCELLED = 0
Const USERSELECTEDDATE = 1
Select Case Index
        Case 0: 'From Date
        'Calculate placement of Calendar Dialog...
            lTop = Me.top + DatFroTxt.top + (2 * DatFroTxt.Height)
            lLeft = Me.left + DatFroTxt.left

            CMaxDate.MaxDate = CDate(DatFroTxt.Text)
            
        Case 1:
        'Calculate placement of Calendar Dialog...
            lTop = Me.top + DatToTxt.top + (2 * DatToTxt.Height)
            lLeft = Me.left + DatToTxt.left

            CMaxDate.MaxDate = CDate(DatToTxt.Text)
        
        Case Else
            MsgBox "This Index is not supported - DateFrm " & CStr(Index)
            Exit Sub
End Select

lRt = CMaxDate.ShowCalendar(lTop, lLeft, Me)
Select Case lRt
    Case USERCANCELLED:
    Case USERSELECTEDDATE:
    Select Case Index
            Case 0:
                DatFroTxt.Text = Format$(CMaxDate.MaxDate, "mm-dd-yyyy")
            Case 1:
                DatToTxt.Text = Format$(CMaxDate.MaxDate, "mm-dd-yyyy")
    End Select
End Select
Set CMaxDate = Nothing

End Sub

Private Sub DatCanCmd_Click()

    cmdCalendar(0).Visible = False
    cmdCalendar(1).Visible = False
    UserCancel = True
    DateFrm.Hide
End Sub


Private Sub DatCanCmd_GotFocus()

    cmdCalendar(0).Visible = False
    cmdCalendar(1).Visible = False

End Sub

Private Sub DatFroTxt_GotFocus()
    cmdCalendar(0).Visible = True
    cmdCalendar(1).Visible = False

End Sub

Private Sub DatFroTxt_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then DatToTxt.SetFocus


End Sub


Private Sub DatFroTxt_ToolTip()
MDI.StatusBar1.Panels(1).Text = DatFroTxt.ToolTipText
End Sub

Private Sub DatOkCmd_Click()
On Local Error GoTo DatOk_Err

Dim sRetStr As String

UserCancel = False
' ------------------------
' Check for valid dates,
' From Date is <= End date
' ------------------------
'
' Check From date For validation
If Not IsDate(DatFroTxt.Text) Then
    MsgBox "Invalid date: " & DatFroTxt.Text, 0, "TourWin - Incorrect Date Format"
    DatFroTxt.SetFocus
End If
If Not IsDate(DatToTxt.Text) Then
    MsgBox "Invalid date: " & DatToTxt.Text, 0, "TourWin - Incorrect Date Format"
    DatToTxt.SetFocus
End If

If DateDiff("d", DatFroTxt.Text, DatToTxt.Text) < 0 Then
    MsgBox "From Date is later than To Date!", 0, "Invalid Date Range"
    DatFroTxt.SetFocus
Else
    ' ------------------------------
    ' Write Date values to registry
    ' ------------------------------
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcDateFrmStart), DateFrm.DatFroTxt.Text)
        If REG_ERROR = sRetStr Then
                gbSkipRegErrMsg = True ' Don't show reg error
                CreateRegKey LoadResString(gcRegTourKey) & "\" & LoadResString(gcDateFrmStart)
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcDateFrmStart), DateFrm.DatFroTxt.Text
        End If
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcDateFrmEnd), DateFrm.DatToTxt.Text)
        If REG_ERROR = sRetStr Then
                gbSkipRegErrMsg = True ' Don't show reg error
                CreateRegKey LoadResString(gcRegTourKey) & "\" & LoadResString(gcDateFrmEnd)
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcDateFrmEnd), DateFrm.DatToTxt.Text
        End If

    DateFrm.Hide
    
End If
Exit Sub
DatOk_Err:
    If bDebug Then Handle_Err Err, "DatOkCmd-DateFrm"
    Resume Next
End Sub


Private Sub DatOkCmd_GotFocus()
cmdCalendar(0).Visible = False
cmdCalendar(1).Visible = False
End Sub

Private Sub DatToTxt_GotFocus()
cmdCalendar(0).Visible = False
cmdCalendar(1).Visible = True


End Sub

Private Sub DatToTxt_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then DatOkCmd.SetFocus

End Sub

Private Sub DatToTxt_ToolTip()
    MDI.StatusBar1.Panels(1).Text = DatToTxt.ToolTipText
End Sub

Private Sub Form_Activate()
    DatFroTxt.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If vbKeyEscape = KeyCode Then
    UserCancel = True
    Me.Hide
End If
End Sub

Private Sub Form_Load()
On Local Error GoTo Daily_Load_Error
Dim sRetVal As String
UserCancel = True

Me.KeyPreview = True

LoadFormResourceString Me
'DatFroTxt.SelLength = 1
'DatToTxt.SelLength = 1
CentreForm DateFrm, 0
' ----------------------------------
' Get Registry Start and End Values
' if Start Date = __-__-____
' ----------------------------------
If Not IsDate(DatFroTxt) Then
gbSkipRegErrMsg = True ' Don't show reg error
sRetVal = GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcDateFrmStart))
DatFroTxt.Text = IIf(REG_ERROR = sRetVal, Format$(Now, "mm-dd-yyyy"), sRetVal)
        
        
gbSkipRegErrMsg = True ' Don't show reg error
sRetVal = GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcDateFrmEnd))
DatToTxt.Text = IIf(REG_ERROR = sRetVal, Format$(Now, "mm-dd-yyyy"), sRetVal)
                
End If
DatFroTxt.SetFocus
Exit Sub
Daily_Load_Error:

    If bDebug Then Handle_Err Err, "Form_Load-DailyFrm"
    Resume Next
End Sub


