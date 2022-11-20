VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EventFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event Type Names"
   ClientHeight    =   5115
   ClientLeft      =   3510
   ClientTop       =   675
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5115
   ScaleWidth      =   4455
   Begin VB.Frame Frame1 
      Caption         =   "Event Type Names :"
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3855
      Begin VB.TextBox Eve 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   25
         TabIndex        =   23
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   1
         Left            =   240
         MaxLength       =   25
         TabIndex        =   22
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   2
         Left            =   240
         MaxLength       =   25
         TabIndex        =   21
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   3
         Left            =   240
         MaxLength       =   25
         TabIndex        =   20
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   4
         Left            =   240
         MaxLength       =   25
         TabIndex        =   19
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   5
         Left            =   240
         MaxLength       =   25
         TabIndex        =   18
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   6
         Left            =   240
         MaxLength       =   25
         TabIndex        =   17
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   7
         Left            =   240
         MaxLength       =   25
         TabIndex        =   16
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   8
         Left            =   240
         MaxLength       =   25
         TabIndex        =   15
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox Eve 
         Height          =   285
         Index           =   9
         Left            =   240
         MaxLength       =   25
         TabIndex        =   14
         Top             =   3840
         Width           =   2535
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   12
         Top             =   960
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   1320
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   10
         Top             =   1680
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   9
         Top             =   2040
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   8
         Top             =   2400
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   6
         Left            =   3000
         TabIndex        =   7
         Top             =   2760
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   7
         Left            =   3000
         TabIndex        =   6
         Top             =   3120
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   8
         Left            =   3000
         TabIndex        =   5
         Top             =   3480
         Width           =   285
      End
      Begin VB.CommandButton EveColCmd 
         Height          =   285
         Index           =   9
         Left            =   3000
         TabIndex        =   4
         Top             =   3840
         Width           =   285
      End
      Begin VB.Label Label3 
         Caption         =   "Color "
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton EveCloCmd 
      Caption         =   "Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton EveSavCmd 
      Caption         =   "S&ave"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog EveComDia 
      Left            =   4560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "The name entered is not the event itself but rather a group name to sort event(s) by."
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "EventFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Eve_Changed As Boolean
Private Sub Eve_Change(Index As Integer)
Eve_Changed = True
End Sub

Sub EveCloCmd_Click()
Dim MsgStr As String, MsgRet As Integer
If Eve_Changed Then
    MsgStr = "Changes have been made, save changes."
    If vbYes = MsgBox(MsgStr, vbYesNo + vbQuestion, LoadResString(gcTourVersion)) Then EveSavCmd_Click
    Eve_Changed = False
End If
Unload EventFrm
End Sub

Private Sub EveColCmd_Click(Index As Integer)
On Local Error GoTo EveCol_Err

Dim sColor As SelectedColor

sColor = ShowColor(Me.hWnd)


Eve(Index).BackColor = sColor.oSelectedColor

Exit Sub
EveCol_Err:
    If Err = 32755 Then      'User Clicked Cancel
        Err.Clear
        Exit Sub
    End If
    If bDebug Then Handle_Err Err, "EveColCmd-EventFrm"
    Screen.MousePointer = 0
    Resume Next
End Sub

Sub EveSavCmd_Click()
On Local Error GoTo EveSav_Err

Dim i As Integer
For i = 0 To 9
            ' Validate Data
            If Len(Eve(i)) >= 31 Then
                MsgBox "Data length to long, max length = 30"
                ObjTour.CancelUpdate iSearcherDB
                Eve(i).SetFocus
                Exit Sub
            End If

            ObjChart.SetColour i, Eve(i).BackColor, gcEventChart
            ObjChart.SetName i, Eve(i).Text, gcEventChart
            
Next i
    ObjChart.SaveChartColoursAndNames gcEventChart
    
' ------------------------------------
' Change Flag so QueryUnload does not
' prompt user again...
' ------------------------------------
Eve_Changed = False
Unload EventFrm

Exit Sub
EveSav_Err:
    If bDebug Then Handle_Err Err, "EveSavCmd-EventFrm"
    Resume Next
End Sub

Private Sub Form_Activate()
' Load menu...
Define_Form_menu Me.Name, Loadmnu
End Sub

Private Sub Form_GotFocus()
' Load menu...
Define_Form_menu Me.Name, Loadmnu
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo Event_Err

Dim i As Integer, RetLng As Long
Dim TypeNum As String, RetStr As String, TypeCol As String


Screen.MousePointer = 11            ' Mouse to hour glase
CentreForm EventFrm, -1
Me.KeyPreview = True
' Load menu...
Define_Form_menu Me.Name, Loadmnu

'EventFrm.Height = 5100

' =======================================
' Load values into memory if not already.
' =======================================
If Not ObjChart.info.IsValuesLoaded Then
    ObjChart.LoadChartColoursAndNames
End If

For i = 0 To 9                      'Loop throught the nine type names
    ' Update MDI Progress Bar
    ProgressBar "Loading Type Names...", -1, i * 1.11, -1
    ' GetField Value
    RetStr = ObjChart.GetName(i, gcEventChart)
        If RetStr <> "No Return" Then
                Eve(i).Text = Trim(RetStr)
        Else
                Eve(i).Text = " "
        End If

    ' GetField Value
    RetStr = ObjChart.GetColour(i, gcEventChart)
    'RetStr = ObjTour.DBGetField(TypeCol, iSearcherDB)
        If RetStr <> "No Return" Then
                Eve(i).BackColor = Trim(RetStr)
        End If
Next i

ProgressBar "", 0, 0, 0
Screen.MousePointer = 0             ' Mouse to default
Eve_Changed = False

Exit Sub
                                    ' Error Handler
Event_Err:
     If bDebug Then Handle_Err Err, "load-EventFrm"
     Screen.MousePointer = 0
     Resume Next
End Sub

Private Sub Form_LostFocus()
' Load menu...
Define_Form_menu Me.Name, Unloadmnu
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Msg ' Declare variable.
If True = Eve_Changed Then
    If UnloadMode > 0 Then
            ' If exiting the application.
            Msg = LoadResString(50)
        Else
            ' If just closing the form.
            Msg = LoadResString(51)
    End If
            ' If user clicks the No button, stop QueryUnload.
        If MsgBox(Msg, vbQuestion + vbYesNo, LoadResString(gcTourVersion)) = vbYes Then
            EveSavCmd_Click
        End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Load menu...
Define_Form_menu Me.Name, Unloadmnu
End Sub
