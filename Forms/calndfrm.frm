VERSION 5.00
Begin VB.Form Calndfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Events Calender"
   ClientHeight    =   5550
   ClientLeft      =   90
   ClientTop       =   1365
   ClientWidth     =   9570
   Icon            =   "calndfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5550
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Tag             =   "700"
   Begin Tourwin2002.UTextBox CalPagMsk 
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Calendar Page (Range 1 to 99)"
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Max             =   2
      FieldType       =   1
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
   End
   Begin Tourwin2002.Spin spnYear 
      Height          =   315
      Left            =   8640
      TabIndex        =   37
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
   End
   Begin Tourwin2002.Spin SpnPageNum 
      Height          =   315
      Left            =   6240
      TabIndex        =   36
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
   End
   Begin VB.ComboBox EveFilCbo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "calndfrm.frx":0442
      Left            =   960
      List            =   "calndfrm.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "|-1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox CalTypCob 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "|-1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame fraTypelist 
      Caption         =   $"calndfrm.frx":0446
      Height          =   1455
      Left            =   240
      TabIndex        =   22
      Tag             =   "706"
      Top             =   3840
      Width           =   9135
      Begin VB.ListBox CalEveLis 
         Height          =   840
         ItemData        =   "calndfrm.frx":04EF
         Left            =   240
         List            =   "calndfrm.frx":04F1
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   8655
      End
   End
   Begin VB.TextBox CalYeaTxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   7
      Tag             =   "|-1"
      Text            =   "1997"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblYear 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Event Yea&r:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Tag             =   "704"
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Event T&ype:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label lbldatafile 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Data File:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Tag             =   "701"
      Top             =   165
      Width           =   855
   End
   Begin VB.Label lblPageNum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Page #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Tag             =   "703"
      Top             =   165
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   9495
      X2              =   9495
      Y1              =   600
      Y2              =   3735
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   600
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   9495
      Y1              =   600
      Y2              =   615
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   0
      X2              =   9480
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Dec 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   35
      Tag             =   "12"
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Nov 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   34
      Tag             =   "11"
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Oct 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   33
      Tag             =   "10"
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Sep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   32
      Tag             =   "09"
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Aug 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   31
      Tag             =   "08"
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Jul 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   30
      Tag             =   "07"
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Jun 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   29
      Tag             =   "06"
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label May 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   28
      Tag             =   "05"
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Apr 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   27
      Tag             =   "04"
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Mar 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   26
      Tag             =   "03"
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Feb 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   25
      Tag             =   "02"
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Jan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   24
      Tag             =   "01"
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDec 
      BackColor       =   &H00C000C0&
      Caption         =   "  Dec"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   21
      Tag             =   "718"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblNov 
      BackColor       =   &H00C000C0&
      Caption         =   "  Nov"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   20
      Tag             =   "717"
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C000C0&
      Caption         =   "  Oct"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   19
      Tag             =   "716"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C000C0&
      Caption         =   "  Sept"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   18
      Tag             =   "715"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C000C0&
      Caption         =   "  Aug"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   17
      Tag             =   "714"
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C000C0&
      Caption         =   "  July"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Tag             =   "713"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C000C0&
      Caption         =   "  June"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Tag             =   "712"
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C000C0&
      Caption         =   "  May"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Tag             =   "711"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   -120
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.Label April 
      BackColor       =   &H00C000C0&
      Caption         =   "  April"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Tag             =   "710"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C000C0&
      Caption         =   "  Mar"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Tag             =   "709"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblFeb 
      BackColor       =   &H00C000C0&
      Caption         =   "  Feb"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Tag             =   "708"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblJan 
      BackColor       =   &H00C000C0&
      Caption         =   "  Jan"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Tag             =   "707"
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lbldayofweek 
      BackColor       =   &H00C000C0&
      Caption         =   " mo tu we  th  fr  sa su  mo tu we  th  fr  sa su mo tu we  th  fr  sa su  mo tu we  th  fr  sa su  mo tu we  th  fr  sa su  mo tu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Tag             =   "705"
      Top             =   600
      Width           =   9015
   End
End
Attribute VB_Name = "Calndfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Calndfrm
' DateTime  : 2/9/03 18:29
' Author    : mor
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit
Dim IsLoad As Integer
Dim cFileChange As String
Dim cEventChange As String
Dim cYearChange As String
Dim DayMonthFocus As String * 10
Dim lEventHandle As Long

Sub Apr_Days(WeekDay As String)
On Local Error GoTo Apr_Err

Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Apr(1).left = Apr(1).left + Apr(1).Width
        Case "Wednesday":
                Apr(1).left = Apr(1).left + (Apr(1).Width * 2)
        Case "Thursday":
                Apr(1).left = Apr(1).left + (Apr(1).Width * 3)
        Case "Friday":
                Apr(1).left = Apr(1).left + (Apr(1).Width * 4)
        Case "Saturday":
                Apr(1).left = Apr(1).left + (Apr(1).Width * 5)
        Case "Sunday":
                Apr(1).left = Apr(1).left + (Apr(1).Width * 6)
End Select
Apr(1).BackColor = &HFFFFFF                         ' White BackColor
'------------------
' 31 day in Aprch
'------------------
For i = 2 To 30
    If Apr.Count < i Then Load Apr(i)
        Apr(i).Caption = Format$(i, "00")           'input cal number
        Apr(i).left = Apr(i - 1).left + Apr(i - 1).Width
        Apr(i).BackColor = &HFFFFFF                  'White backColor
        Apr(i).Visible = False
Next i
Exit Sub
Apr_Err:
    If bDebug Then Handle_Err Err, "Apr_Days-Calndfrm"
    Resume Next
End Sub

Sub Aug_Days(WeekDay As String)
On Local Error GoTo Aug_Err

Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Aug(1).left = Aug(1).left + Aug(1).Width
        Case "Wednesday":
                Aug(1).left = Aug(1).left + (Aug(1).Width * 2)
        Case "Thursday":
                Aug(1).left = Aug(1).left + (Aug(1).Width * 3)
        Case "Friday":
                Aug(1).left = Aug(1).left + (Aug(1).Width * 4)
        Case "Saturday":
                Aug(1).left = Aug(1).left + (Aug(1).Width * 5)
        Case "Sunday":
                Aug(1).left = Aug(1).left + (Aug(1).Width * 6)
End Select
Aug(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 31 day in Augch
'------------------
For i = 2 To 31
    If Aug.Count < i Then Load Aug(i)
        Aug(i).Caption = Format$(i, "00")
        Aug(i).left = Aug(i - 1).left + Aug(i - 1).Width
        Aug(i).BackColor = &HFFFFFF                  'White backColor
        Aug(i).Visible = False
Next i
Exit Sub
Aug_Err:
    If bDebug Then Handle_Err Err, "Aug_Days-Calndfrm"
    Resume Next
End Sub

Sub Dec_Days(WeekDay As String)
On Local Error GoTo Dec_Err

Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Dec(1).left = Dec(1).left + Dec(1).Width
        Case "Wednesday":
                Dec(1).left = Dec(1).left + (Dec(1).Width * 2)
        Case "Thursday":
                Dec(1).left = Dec(1).left + (Dec(1).Width * 3)
        Case "Friday":
                Dec(1).left = Dec(1).left + (Dec(1).Width * 4)
        Case "Saturday":
                Dec(1).left = Dec(1).left + (Dec(1).Width * 5)
        Case "Sunday":
                Dec(1).left = Dec(1).left + (Dec(1).Width * 6)
End Select
Dec(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 31 day in Decch
'------------------
For i = 2 To 31
    If Dec.Count < i Then Load Dec(i)
        Dec(i).Caption = Format$(i, "00")
        Dec(i).left = Dec(i - 1).left + Dec(i - 1).Width
        Dec(i).BackColor = &HFFFFFF                  'White backColor
        Dec(i).Visible = False
Next i
Exit Sub
Dec_Err:
    If bDebug Then Handle_Err Err, "Dec_Days-Calndfrm"
    Resume Next
End Sub

Sub DeleteDays(StartD As String, EndD As String)
On Local Error GoTo DeleteDays_Err
Dim SQL As String, LengthOfDelete As String, PageNum As String
' -------------------------
' Make to select statements
' 1) Peak table
' 2) Daily & Event table
' -------------------------
    PageNum = Trim$(CalPagMsk.Text)
    LengthOfDelete = "#" & StartD & "# and #" & EndD & "#"
    If Trim$(EveFilCbo) = "Daily" Then
        SQL = "DELETE * FROM " & Trim$(EveFilCbo) & " WHERE Date BETWEEN " & LengthOfDelete & " AND id = " & objMdi.info.ID
    Else
        SQL = "DELETE * FROM " & Trim$(EveFilCbo) & " WHERE Date BETWEEN " & LengthOfDelete & " And Page = " & PageNum & " AND id = " & objMdi.info.ID
    End If
    ObjTour.DBExecute SQL

Exit Sub
DeleteDays_Err:
    If bDebug Then Handle_Err Err, "DeleteDays-CalndFrm"
    Resume Next
End Sub
Sub DeleteDaysByPeakName(ByVal sName As String)
On Local Error GoTo DeleteDaysByPeakName_Err
Dim SQL As String, LengthOfDelete As String, PageNum As String
' -------------------------
' Make to select statements
' 1) Peak table
' 2) Daily & Event table
' -------------------------


' THIS PROCEDURE NEEDS TO BE TESTED!!!!!!!!!!!!!!

    PageNum = Trim$(CalPagMsk.Text)

    If Trim$(EveFilCbo) = "Daily" Then
        SQL = "DELETE * FROM " & Trim$(EveFilCbo) & " WHERE Name = " & sName & " AND id = " & objMdi.info.ID
    Else
        SQL = "DELETE * FROM " & Trim$(EveFilCbo) & " WHERE Name = " & sName & " And Page = " & PageNum & " And ID = " & objMdi.info.ID
    End If
    ObjTour.DBExecute SQL

Exit Sub
DeleteDaysByPeakName_Err:
    If bDebug Then Handle_Err Err, "DeleteDaysByPeakName-CalndFrm"
    Resume Next
End Sub

Function FormatString_For_ListBox(dDate As Date, sEvent As String, sType As String) As Boolean
On Local Error GoTo FormatString_Err
' -----------------------------------------------------------------
' Purpose: Adds a days activity to CalEveLis Combo Box. There this
' function can be called from anywhere.
' Return value: Boolean, False is unsuccessful
'                        True is successful.
' Format of String: 15 spaces for date, 150 spaces for event, 20
' for type. Event string is centered
' -----------------------------------------------------------------
Dim ToListOfDays As String, LenOfDesc As Byte, Allecate As Byte
FormatString_For_ListBox = False
Allecate = 115
 ' Add Date to String
 'ToListOfDays = Format$(dDate, "mm-dd-yyyy") & vbTab
 ' Add sEvent to String
 'LenOfDesc = Len(sEvent)
' If LenOfDesc >= Allecate Then
'        ToListOfDays = ToListOfDays & Mid$(Trim$(sEvent), 1, Allecate)
' Else
ToListOfDays = Format$(dDate, "mm-dd-yyyy") & vbTab & Trim$(sEvent) & vbTab & Trim$(sType)
 'End If
' Add sType to String
'ToListOfDays = ToListOfDays &
    CalEveLis.AddItem ToListOfDays
FormatString_For_ListBox = True
Exit Function
FormatString_Err:
    If bDebug Then Handle_Err Err, "FormatString_For_ListBox-CalndFrm"
    Exit Function
End Function

Sub Jan_Days(WeekDay As String)
On Local Error GoTo Jan_Err

Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Jan(1).left = Jan(1).left + Jan(1).Width
        Case "Wednesday":
                Jan(1).left = Jan(1).left + (Jan(1).Width * 2)
        Case "Thursday":
                Jan(1).left = Jan(1).left + (Jan(1).Width * 3)
        Case "Friday":
                Jan(1).left = Jan(1).left + (Jan(1).Width * 4)
        Case "Saturday":
                Jan(1).left = Jan(1).left + (Jan(1).Width * 5)
        Case "Sunday":
                Jan(1).left = Jan(1).left + (Jan(1).Width * 6)
End Select
Jan(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 28 day in January
'------------------
For i = 2 To 31
    If Jan.Count < i Then Load Jan(i)
        Jan(i).Caption = Format$(i, "00")
        Jan(i).left = Jan(i - 1).left + Jan(i - 1).Width
        Jan(i).BackColor = &HFFFFFF                  'White backColor
        Jan(i).Visible = False
Next i
Exit Sub

Jan_Err:
     If bDebug Then Handle_Err Err, "Jan_Days-Calndfrm"
    Resume Next
End Sub


Sub Feb_Days(WeekDay As String)
On Local Error GoTo Feb_Err

Dim i As Integer
Dim iDays As Integer

Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Feb(1).left = Feb(1).left + Feb(1).Width
        Case "Wednesday":
                Feb(1).left = Feb(1).left + (Feb(1).Width * 2)
        Case "Thursday":
                Feb(1).left = Feb(1).left + (Feb(1).Width * 3)
        Case "Friday":
                Feb(1).left = Feb(1).left + (Feb(1).Width * 4)
        Case "Saturday":
                Feb(1).left = Feb(1).left + (Feb(1).Width * 5)
        Case "Sunday":
                Feb(1).left = Feb(1).left + (Feb(1).Width * 6)
End Select
Feb(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 28 day in Febuary
'------------------
'iDays = IIf((Year(Calndfrm.CalYeaTxt.Text) Mod 4) = 0, 29, 28)
' Always load 29 days,
' but on show 29 control for leap years
For i = 2 To 29
    If Feb.Count < i Then Load Feb(i)
        Feb(i).Caption = Format$(i, "00")
        Feb(i).left = Feb(i - 1).left + Feb(i - 1).Width
        Feb(i).BackColor = &HFFFFFF                  'White backColor
        Feb(i).Visible = False
Next i
Exit Sub

Feb_Err:
    If bDebug Then Handle_Err Err, "Feb_Days-Calndfrm"
    Resume Next
End Sub


Sub Jul_Days(WeekDay As String)
On Local Error GoTo Jul_Err

Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Jul(1).left = Jul(1).left + Jul(1).Width
        Case "Wednesday":
                Jul(1).left = Jul(1).left + (Jul(1).Width * 2)
        Case "Thursday":
                Jul(1).left = Jul(1).left + (Jul(1).Width * 3)
        Case "Friday":
                Jul(1).left = Jul(1).left + (Jul(1).Width * 4)
        Case "Saturday":
                Jul(1).left = Jul(1).left + (Jul(1).Width * 5)
        Case "Sunday":
                Jul(1).left = Jul(1).left + (Jul(1).Width * 6)
End Select
Jul(1).BackColor = &HFFFFFF                         ' White BackColor
'------------------
' 31 day in Julch
'------------------
For i = 2 To 31
    If Jul.Count < i Then Load Jul(i)
        Jul(i).Caption = Format$(i, "00")
        Jul(i).left = Jul(i - 1).left + Jul(i - 1).Width
        Jul(i).BackColor = &HFFFFFF                  'White backColor
        Jul(i).Visible = False
Next i
Exit Sub
Jul_Err:
     If bDebug Then Handle_Err Err, "Jul_Days-Calndfrm"
    Resume Next

End Sub

Sub Jun_Days(WeekDay As String)
On Local Error GoTo Jun_Err


Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Jun(1).left = Jun(1).left + Jun(1).Width
        Case "Wednesday":
                Jun(1).left = Jun(1).left + (Jun(1).Width * 2)
        Case "Thursday":
                Jun(1).left = Jun(1).left + (Jun(1).Width * 3)
        Case "Friday":
                Jun(1).left = Jun(1).left + (Jun(1).Width * 4)
        Case "Saturday":
                Jun(1).left = Jun(1).left + (Jun(1).Width * 5)
        Case "Sunday":
                Jun(1).left = Jun(1).left + (Jun(1).Width * 6)
End Select
Jun(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 31 day in Junch
'------------------
For i = 2 To 30
    If Jun.Count < i Then Load Jun(i)
        Jun(i).Caption = Format$(i, "00")
        Jun(i).left = Jun(i - 1).left + Jun(i - 1).Width
        Jun(i).BackColor = &HFFFFFF                  'White backColor
        Jun(i).Visible = False
Next i
Exit Sub
Jun_Err:
     If bDebug Then Handle_Err Err, "Jun_Days-Calndfrm"
    Resume Next

End Sub

Sub Label_Visible(Way As Integer)
Dim i As Integer
Dim iDays  As Integer
' -----------------------
' All months that have 30

For i = 1 To Jan.Count
    Jan(i).Visible = Way
    Mar(i).Visible = Way
    May(i).Visible = Way
    Jul(i).Visible = Way
    Aug(i).Visible = Way
    Oct(i).Visible = Way
    Dec(i).Visible = Way
Next i
'----------------------------
'All months that have 31 days
For i = 1 To Sep.Count
    Sep(i).Visible = Way
    Apr(i).Visible = Way
    Jun(i).Visible = Way
    Nov(i).Visible = Way
Next i

'
' Feb incase of leap year,
' process separately
If Way = 0 Then
    For i = 1 To Feb.Count
        Feb(i).Visible = Way
    Next i
Else
    iDays = IIf((Val(Calndfrm.CalYeaTxt.Text) Mod 4) = 0, 29, 28)
    For i = 1 To iDays
        Feb(i).Visible = Way
    Next i
End If
    
   
End Sub

Sub LoadCmdDays(Centry As String)
On Local Error GoTo LoadCmdDays_Err


Dim FirstDayOfMonth As String, UserDate As String
Dim i As Integer, MonCounter As Integer
Dim DateStr As String
'loads each months control days

For MonCounter = 1 To 12       ' Loop through each month
'-------------------------
'Setup to get day name for
'beginning of each month
'-------------------------
ProgressBar "Loading Days...", -1, MonCounter * 0.83, -1
' think is the area which has problems with
' windows 95, must state MM-DD-YYYY

DateStr = Format$(MonCounter, "00") & "-01-" & Centry
UserDate = Format$(DateStr, "MM-DD-YYYY")
        'UserDate = Format$(MonCounter, "00") & "-" & "01-" & Centry
FirstDayOfMonth = Format$(UserDate, "dddd")
Select Case MonCounter
        Case 1:
            Jan_Days (FirstDayOfMonth)
        Case 2:
            Feb_Days (FirstDayOfMonth)
        Case 3:
            Mar_Days (FirstDayOfMonth)
        Case 4:
            Apr_Days (FirstDayOfMonth)
        Case 5:
            May_Days (FirstDayOfMonth)
        Case 6:
            Jun_Days (FirstDayOfMonth)
        Case 7:
            Jul_Days (FirstDayOfMonth)
        Case 8:
            Aug_Days (FirstDayOfMonth)
        Case 9:
            Sep_Days (FirstDayOfMonth)
        Case 10:
            Oct_Days (FirstDayOfMonth)
        Case 11:
            Nov_Days (FirstDayOfMonth)
        Case 12:
            Dec_Days (FirstDayOfMonth)
End Select
Next MonCounter
ProgressBar "", 0, 0, 0
Exit Sub

LoadCmdDays_Err:
    If bDebug Then Handle_Err Err, "LoadCmdDays-Calndfrm"
    Resume Next
End Sub



Sub LoadEventDays(Centry As String, EventType As String, PageNum As Integer, Optional sMonth As String, Optional ByVal sTypeString As String)
' -----------------------------
' This sub obtains a dynaset
' of the current selected year
' and update each months cmd
' and CalYealis box.
' -----------------------------
On Local Error GoTo LoadEventDays_Err

Dim BegOfYear As String, EndOfYear As String, SQL As String
Dim LenghtOfYear As String, ToListOfDays As String, TempStr As String
Dim Records As Integer, LenOfDesc As Integer
Dim i As Integer, Found As Boolean
Dim iGetCBIndexFromString As Long

Found = False
CalEveLis.Visible = 0
BegOfYear = "01-01-" & Centry
EndOfYear = "12-31-" & Centry
LenghtOfYear = "#" & BegOfYear & "# and #" & EndOfYear & "#"

Select Case EventType
    Case gcDailyChart:
        If Not IsMissing(sTypeString) And gcALLTYPES <> sTypeString Then
            SQL = "Select * From " & EventType & " WHERE Date BETWEEN " & LenghtOfYear & " AND Id = " & objMdi.info.ID & " And DayType = '" & sTypeString & "'"
        Else
            SQL = "Select * From " & EventType & " WHERE Date BETWEEN " & LenghtOfYear & " AND Id = " & objMdi.info.ID
        End If
        
    Case gcEventChart
        If Not IsMissing(sTypeString) And gcALLTYPES <> sTypeString Then
            SQL = "Select * From " & EventType & " WHERE Date BETWEEN " & LenghtOfYear & " And Page = " & PageNum & " And EveType = '" & sTypeString & "' AND Id = " & objMdi.info.ID
        Else
            SQL = "Select * From " & EventType & " WHERE Date BETWEEN " & LenghtOfYear & " And Page = " & PageNum & " AND Id = " & objMdi.info.ID
        End If

    Case gcPeakChart:
        If Not IsMissing(sTypeString) And gcALLTYPES <> sTypeString Then
            SQL = "Select * From " & EventType & " WHERE Date BETWEEN " & LenghtOfYear & " And Page = " & PageNum & " And CycleName = '" & sTypeString & "' AND Id = " & objMdi.info.ID
        Else
            SQL = "Select * From " & EventType & " WHERE Date BETWEEN " & LenghtOfYear & " And Page = " & PageNum & " AND Id = " & objMdi.info.ID
        End If
End Select

' Define RecordSet.
ObjTour.RstSQL lEventHandle, SQL

' Check if any records match
If ObjTour.RstRecordCount(lEventHandle) = 0 Then Exit Sub

Records = 0   'Records processed = 0

  Dim tabpos(2) As Long
    Dim x&
    tabpos(0) = 50 ' About 10 characters
    tabpos(1) = 312
    
    'Clear any existing Tab Stops
    CalEveLis.Visible = False
    Call SendMessage(CalEveLis.hWnd, LB_SETTABSTOPS, 0&, 0&)
    Call SendMessage(CalEveLis.hWnd, LB_SETTABSTOPS, 2, tabpos(0))
    
Do Until ObjTour.EOF(lEventHandle)

    ProgressBar "Loading Events...", -1, (Records / ObjTour.RstRecordCount(lEventHandle)) * 10, -1
    
    Select Case Format$(ObjTour.DBGetField(gcDate, lEventHandle), "mmmm")
            Case "January":
                Jan(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "February":
                Feb(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "March":
                Mar(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "April":
                Apr(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "May":
                 May(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "June":
                Jun(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "July":
                Jul(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "August":
                Aug(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "September":
                Sep(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "October":
                Oct(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "November":
                Nov(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
            Case "December":
                Dec(Format$(ObjTour.DBGetField(gcDate, lEventHandle), "dd")).BackColor = ObjTour.DBGetField(gcColor, lEventHandle)
    End Select
    'Max len is 100
    Select Case EveFilCbo.Text
       Case "Peak":
            FormatString_For_ListBox ObjTour.DBGetField(gcDate, lEventHandle), ObjTour.DBGetField(gcPeakTour_Descr, lEventHandle), ObjTour.DBGetField(gcPeakTour_CycleName, lEventHandle)
       Case "Daily":
            FormatString_For_ListBox ObjTour.DBGetField(gcDate, lEventHandle), "", ObjTour.DBGetField(gcDai_Tour_DayType, lEventHandle)
        Case "Event":
            FormatString_For_ListBox ObjTour.DBGetField(gcDate, lEventHandle), ObjTour.DBGetField(gcEve_Tour_Evememo, lEventHandle), ObjTour.DBGetField(gcEve_Tour_EveType, lEventHandle)
    End Select
    ObjTour.DBMoveNext lEventHandle
    Records = Records + 1               'Records processed
    
Loop

If Records > 0 Then
    CalEveLis.Visible = True
    CalEveLis.Refresh
    If CalEveLis.ListCount > 0 Then CalEveLis.ListIndex = 0
End If

' Set ListBox to Closest list for todays date!
iGetCBIndexFromString = FindItemListControl(CalEveLis, Format$(Date, "mm-dd-yyyy"))
If -1 <> iGetCBIndexFromString Then
  ' Push list to bottom, then desired with appear at top of list...
  With CalEveLis
    .ListIndex = .ListCount - 1
    .ListIndex = iGetCBIndexFromString
  End With
End If

ProgressBar "", 0, 0, 0                 'Clear Progress bar



Exit Sub
LoadEventDays_Err:
     MsgBox Err.Description
     If bDebug Then
        Handle_Err Err, "LoadEventDays-Calndfrm"
    End If
    Resume Next
End Sub

Sub LoadFile()
' -----------------------------------------------
' Names are names of tables from Events database
' -----------------------------------------------

On Local Error GoTo LoadFile_Err
' -----------------------------------
' First, load control is values
' and set to default incase calendar
' has now LoadLast values
' -----------------------------------

EveFilCbo.AddItem gcPeakChart
EveFilCbo.AddItem gcEventChart
EveFilCbo.AddItem gcDailyChart

EveFilCbo.ListIndex = 0

CalTypCob.AddItem gcALLTYPES
CalTypCob.ListIndex = 0
CalPagMsk.Text = "01"
SpnPageNum.Value = Int(CalPagMsk.Text)

' Today's year
CalYeaTxt.Text = Format$(Date, "YYYY")
spnYear.Value = Int(CalYeaTxt.Text)
ObjTour.Settings.ReadFormSettingsFromReg Me
'-------------------------
'Load all Event Type Names
'-------------------------
LoadTypeNames EveFilCbo.Text, CalTypCob.Text

ObjTour.Settings.ReadFormSettingsFromReg Me

Exit Sub

LoadFile_Err:
    If bDebug Then Handle_Err Err, "LoadFile-Calndfrm"
    Resume Next
End Sub

Sub LoadTypeNames(ByVal DataType As String, ByVal sLastSelected As String)
' --------------------------
' Load Event Type Names
' from Tour.ini file for the
' appropriate event table...
' --------------------------
On Local Error GoTo LoadTypeNames_Err
Dim i As Integer
Dim EventName As String, EventType As String, SQL As String
Dim iGetCBIndexFromString As Long
Dim lPeakDistinctHandle As Long
Dim LenghtOfYear As String

CalTypCob.Clear
CalTypCob.AddItem gcALLTYPES
CalTypCob.ListIndex = 0

Select Case DataType
    Case "Event":
        ProgressBar "Loading Types...", -1, i * 1.1111, -1
        
        ' Build SQL String
        SQL = "SELECT Distinct " & gcEve_Tour_EveType & " From " & gcEve_Tour_Event & _
              " WHERE Page = " & CalPagMsk.Text & _
              " And Id = " & objMdi.info.ID & _
              " ORDER BY " & gcEve_Tour_EveType & " ASC"
                      
        ' Get a Distinct List of CycleNames...
        ObjTour.RstSQL lPeakDistinctHandle, SQL

        ' Check if any records match
        If ObjTour.RstRecordCount(lPeakDistinctHandle) <> 0 Then
    
            Do Until ObjTour.EOF(lPeakDistinctHandle)
                CalTypCob.AddItem ObjTour.DBGetField(gcEve_Tour_EveType, lPeakDistinctHandle)
                ObjTour.DBMoveNext lPeakDistinctHandle
            Loop
        
            iGetCBIndexFromString = FindItemListControl(CalTypCob, sLastSelected)
            If -1 <> iGetCBIndexFromString Then
                CalTypCob.ListIndex = iGetCBIndexFromString
            Else
                CalTypCob.ListIndex = 0
            End If
            
        End If
        
        ObjTour.FreeHandle lPeakDistinctHandle
        
        ProgressBar "", 0, 0, 0


    Case "Peak":
        ProgressBar "Loading Types...", -1, i * 1.1111, -1
        
        ' Build SQL String
        SQL = "SELECT Distinct CycleName From " & gcEve_Tour_Peak & _
              " WHERE Page = " & CalPagMsk.Text & _
              " And Id = " & objMdi.info.ID & _
              " ORDER BY CycleName ASC"
                      
        ' Get a Distinct List of CycleNames...
        ObjTour.RstSQL lPeakDistinctHandle, SQL

        ' Check if any records match
        If ObjTour.RstRecordCount(lPeakDistinctHandle) <> 0 Then
    
            Do Until ObjTour.EOF(lPeakDistinctHandle)
                CalTypCob.AddItem ObjTour.DBGetField("CycleName", lPeakDistinctHandle)
                ObjTour.DBMoveNext lPeakDistinctHandle
            Loop
        
            iGetCBIndexFromString = FindItemListControl(CalTypCob, sLastSelected)
            If -1 <> iGetCBIndexFromString Then
                CalTypCob.ListIndex = iGetCBIndexFromString
            Else
                CalTypCob.ListIndex = 0
            End If
            
        End If
        
        ObjTour.FreeHandle lPeakDistinctHandle
        
        ProgressBar "", 0, 0, 0
        
    Case "Daily":
        ' ------------------------------
        ' Daily is a view only option
        ' if user wishes to make changes
        ' then do so through the
        ' Daily Form...
        ' ------------------------------
        ProgressBar "Loading Types...", -1, i * 1.1111, -1
        
        ' Build SQL String
        SQL = "SELECT Distinct " & gcDai_Tour_DayType & " From " & gcEve_Tour_Event_Daily & _
              " WHERE Id = " & objMdi.info.ID & _
              " ORDER BY " & gcDai_Tour_DayType & " ASC"
                      
        ' Get a Distinct List of CycleNames...
        ObjTour.RstSQL lPeakDistinctHandle, SQL
        Debug.Print SQL
        ' Check if any records match
        If ObjTour.RstRecordCount(lPeakDistinctHandle) <> 0 Then
    
            Do Until ObjTour.EOF(lPeakDistinctHandle)
                CalTypCob.AddItem ObjTour.DBGetField("DayType", lPeakDistinctHandle)
                ObjTour.DBMoveNext lPeakDistinctHandle
            Loop
        
            iGetCBIndexFromString = FindItemListControl(CalTypCob, sLastSelected)
            If -1 <> iGetCBIndexFromString Then
                CalTypCob.ListIndex = iGetCBIndexFromString
            Else
                CalTypCob.ListIndex = 0
            End If
            
        End If
        
        ObjTour.FreeHandle lPeakDistinctHandle
        
        ProgressBar "", 0, 0, 0
End Select
    CalTypCob.Refresh
Exit Sub
LoadTypeNames_Err:
    Call Handle_Err(Err, "LoadTypeNames-Calndfrm")
    Resume Next
End Sub

Sub Mar_Days(WeekDay As String)
    On Local Error GoTo Mar_Err

Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Mar(1).left = Mar(1).left + Mar(1).Width
        Case "Wednesday":
                Mar(1).left = Mar(1).left + (Mar(1).Width * 2)
        Case "Thursday":
                Mar(1).left = Mar(1).left + (Mar(1).Width * 3)
        Case "Friday":
                Mar(1).left = Mar(1).left + (Mar(1).Width * 4)
        Case "Saturday":
                Mar(1).left = Mar(1).left + (Mar(1).Width * 5)
        Case "Sunday":
                Mar(1).left = Mar(1).left + (Mar(1).Width * 6)
End Select
Mar(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 31 day in March
'------------------
For i = 2 To 31
    If Mar.Count < i Then Load Mar(i)
        Mar(i).Caption = Format$(i, "00")
        Mar(i).left = Mar(i - 1).left + Mar(i - 1).Width
        Mar(i).BackColor = &HFFFFFF                  'White backColor
        Mar(i).Visible = False
Next i
Exit Sub
Mar_Err:
    Call Handle_Err(Err, "Mar_Days-Calndfrm")
    Resume Next
            
End Sub


Sub May_Days(WeekDay As String)
On Local Error GoTo May_Err
Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                May(1).left = May(1).left + May(1).Width
        Case "Wednesday":
                May(1).left = May(1).left + (May(1).Width * 2)
        Case "Thursday":
                May(1).left = May(1).left + (May(1).Width * 3)
        Case "Friday":
                May(1).left = May(1).left + (May(1).Width * 4)
        Case "Saturday":
                May(1).left = May(1).left + (May(1).Width * 5)
        Case "Sunday":
                May(1).left = May(1).left + (May(1).Width * 6)
End Select
May(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 31 day in Maych
'------------------
For i = 2 To 31
    If May.Count < i Then Load May(i)
        May(i).Caption = Format$(i, "00")
        May(i).left = May(i - 1).left + May(i - 1).Width
        May(i).BackColor = &HFFFFFF                  'White backColor
        May(i).Visible = False
Next i
Exit Sub
May_Err:
    Call Handle_Err(Err, "May_Days-Calndfrm")
    Resume Next
End Sub

Sub Month_Clicked(Day As Integer, Month As String)
' -----------------------
' User is either asking
' to create new day or
' wants toupdate existing
' -----------------------
Dim DateStr As String
On Local Error GoTo Mon_Click_Err
Month = Format$(Month, "00;(00)")
DateStr = "Date = " & "#" & Month & "-" & Format$(Day, "00") & "-" & CalYeaTxt.Text & "#"

objEve.info.EventDate = Format$(Mid$(DateStr, 9, 10), "mm-dd,yyyy")
Select Case EveFilCbo.Text
       Case "Peak":
                CreatFrm.CreConchk.Visible = True
                CreatFrm.Show 1
        Case "Daily":
                'CreatFrm.Frame1.Visible = False
                CreatFrm.CreConchk.Visible = False
                CreatFrm.Show 1
        Case "Event":
            objEve.info.EventDate = Format$(Mid$(DateStr, 9, 10), "mm-dd,yyyy")
            CreEveFrm.Caption = Format$(Mid$(DateStr, 9, 10), "mmmm dd, yyyy")
            CreEveFrm.Show 1
End Select
Exit Sub
Mon_Click_Err:


    If bDebug Then Handle_Err Err, "Month_Clicked-CalndFrm"

    Resume Next
End Sub

Sub Nov_Days(WeekDay As String)
On Local Error GoTo Nov_Err
Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Nov(1).left = Nov(1).left + Nov(1).Width
        Case "Wednesday":
                Nov(1).left = Nov(1).left + (Nov(1).Width * 2)
        Case "Thursday":
                Nov(1).left = Nov(1).left + (Nov(1).Width * 3)
        Case "Friday":
                Nov(1).left = Nov(1).left + (Nov(1).Width * 4)
        Case "Saturday":
                Nov(1).left = Nov(1).left + (Nov(1).Width * 5)
        Case "Sunday":
                Nov(1).left = Nov(1).left + (Nov(1).Width * 6)
End Select
Nov(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 31 day in Novch
'------------------
For i = 2 To 30
    If Nov.Count < i Then Load Nov(i)
        Nov(i).Caption = Format$(i, "00")
        Nov(i).left = Nov(i - 1).left + Nov(i - 1).Width
        Nov(i).BackColor = &HFFFFFF                  'White backColor
        Nov(i).Visible = False
Next i
Exit Sub
Nov_Err:
    Call Handle_Err(Err, "Nov_Days-Calndfrm")
    Resume Next
End Sub

Sub Oct_Days(WeekDay As String)
On Local Error GoTo Oct_Err
Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Oct(1).left = Oct(1).left + Oct(1).Width
        Case "Wednesday":
                Oct(1).left = Oct(1).left + (Oct(1).Width * 2)
        Case "Thursday":
                Oct(1).left = Oct(1).left + (Oct(1).Width * 3)
        Case "Friday":
                Oct(1).left = Oct(1).left + (Oct(1).Width * 4)
        Case "Saturday":
                Oct(1).left = Oct(1).left + (Oct(1).Width * 5)
        Case "Sunday":
                Oct(1).left = Oct(1).left + (Oct(1).Width * 6)
End Select
Oct(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 31 day in Octch
'------------------
For i = 2 To 31
    If Oct.Count < i Then Load Oct(i)
        Oct(i).Caption = Format$(i, "00")
        Oct(i).left = Oct(i - 1).left + Oct(i - 1).Width
        Oct(i).BackColor = &HFFFFFF                  'White backColor
        Oct(i).Visible = False
Next i
Exit Sub
Oct_Err:
    Call Handle_Err(Err, "Oct_Days-Calndfrm")
    Resume Next

End Sub
Sub RestCalendar()
' --------------------------------
' Rests Label to starting position
' --------------------------------
On Local Error GoTo Rest_Err
Dim CurrentCentry As String
Dim i As Integer


CurrentCentry = Format$(Val(CalYeaTxt.Text), "0000")

Screen.MousePointer = 11
Label_Visible 0
Jan(1).left = 480
Feb(1).left = 480
Mar(1).left = 480
Apr(1).left = 480
May(1).left = 480
Jun(1).left = 480
Jul(1).left = 480
Aug(1).left = 480
Sep(1).left = 480
Oct(1).left = 480
Nov(1).left = 480
Dec(1).left = 480

CalEveLis.Clear ' clear list of previous events

LoadCmdDays (CurrentCentry)                     'Align Cmd controls

'--------------------------------
'Now load all list all event
'using select year and Event type
'--------------------------------
Call LoadEventDays(CurrentCentry, EveFilCbo.Text, CalPagMsk.Text, , CalTypCob.Text)
Label_Visible -1

If ObjChart.info.IsFormLoaded And ObjChart.info.CurrentChartType <> EveFilCbo.Text Then
    ObjChart.info.CurrentChartType = EveFilCbo.Text
    ChartFrm.DisplayChartColours ObjChart.info.CurrentChartType
End If

Screen.MousePointer = 0
Exit Sub
Rest_Err:
    If bDebug Then Handle_Err Err, "RestCalnder-Calndfrm"
    Resume Next
End Sub




Sub Sep_Days(WeekDay As String)
On Local Error GoTo Sep_Err
Dim i As Integer
Select Case WeekDay
        Case "Monday":
        Case "Tuesday":
                Sep(1).left = Sep(1).left + Sep(1).Width
        Case "Wednesday":
                Sep(1).left = Sep(1).left + (Sep(1).Width * 2)
        Case "Thursday":
                Sep(1).left = Sep(1).left + (Sep(1).Width * 3)
        Case "Friday":
                Sep(1).left = Sep(1).left + (Sep(1).Width * 4)
        Case "Saturday":
                Sep(1).left = Sep(1).left + (Sep(1).Width * 5)
        Case "Sunday":
                Sep(1).left = Sep(1).left + (Sep(1).Width * 6)
End Select
Sep(1).BackColor = &HFFFFFF                         ' White BackColor

'------------------
' 31 day in Sepch
'------------------
For i = 2 To 30
    If Sep.Count < i Then Load Sep(i)
        Sep(i).Caption = Format$(i, "00")
        Sep(i).left = Sep(i - 1).left + Sep(i - 1).Width
        Sep(i).BackColor = &HFFFFFF                  'White backColor
        Sep(i).Visible = False
Next i
Exit Sub
Sep_Err:
    Call Handle_Err(Err, "Sep_Days-Calndfrm")
    Resume Next

End Sub

Private Sub Apr_DblClick(Index As Integer)
On Local Error GoTo Apr_Clicked_Err
Month_Clicked Index, Apr(Index).Tag
Exit Sub
Apr_Clicked_Err:
        If bDebug Then Handle_Err Err, "Apr_Clicked-CalndFrm"
        Resume Next
End Sub


Private Sub Apr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "04-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub


Private Sub Aug_DblClick(Index As Integer)
On Local Error GoTo Aug_Clicked_Err
Month_Clicked Index, Aug(Index).Tag
Exit Sub
Aug_Clicked_Err:
        If bDebug Then Handle_Err Err, "Aug_Clicked-CalndFrm"
        Resume Next
End Sub


Private Sub Aug_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "08-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub

Sub CalBegmnu()
' ----------------------------
' Check Control values first
' before call RestCalendar sub
' ----------------------------
On Local Error Resume Next
If IsLoad Then
    If EveFilCbo.Text = "" Then
        MsgBox "Select data source from " & vbLf & "'Data File' list box.", 64, "TourWin Warning"
        EveFilCbo.SetFocus
        Exit Sub
    End If
    If CalTypCob.Text = "" Then
        MsgBox "Select Event Type source from " & vbLf & "'Event Type' list box.", 64, "TourWin Warning"
        CalTypCob.SetFocus
        Exit Sub
    End If
    If Trim$(EveFilCbo.Text) = "Peak" Then
        If Val(CalPagMsk.Text) > 99 Or Val(CalPagMsk.Text) < 1 Then
        MsgBox "Select Page source from " & vbLf & "'Page #' text box.", 64, "TourWin Warning"
        CalPagMsk.SetFocus
        Exit Sub
        End If
    End If
    If CalYeaTxt.Text = "" Then
        MsgBox "Select year range from " & vbLf & "'Event Year' list box.", 64, "TourWin Warning"
        CalYeaTxt.SetFocus
        Exit Sub
    End If
        MDI.MdiDelmnu.Enabled = True
        MDI.MdiNewmnu.Enabled = True
        
        If ObjChart.info.IsFormLoaded Then
            ObjChart.info.CurrentChartType = EveFilCbo.Text
            ChartFrm.DisplayChartColours ObjChart.info.CurrentChartType
        End If
    RestCalendar
Else
    ' Load EveFilCbo, CalTypCob etc
'    LoadFile
    
End If


    
End Sub


Sub CalDelmnu()
On Local Error GoTo CalDel_Err
Dim StartD As String, EndD As String

objMdi.info.sCurrentActiveWindow = gcCalendarWindow

    UserCancel = False
    PeakOptRpt.Show 1

    objMdi.info.sCurrentActiveWindow = ""
    If UserCancel = True Then
        Exit Sub
    End If
    CalBegmnu
Exit Sub
CalDel_Err:

    If bDebug Then Handle_Err Err, "CalDelmnu-CalndFrm"
    Resume Next
End Sub



Private Sub CalPagMsk_Change()

    If 0 = IsLoad Then Exit Sub
    
    If Val(CalPagMsk.Text) > SpnPageNum.Max _
    Or Val(CalPagMsk.Text) < SpnPageNum.Min Then Exit Sub
    
    SpnPageNum.Value = Val(CalPagMsk.Text)
            
    IsLoad = 0
    RestCalendar
    IsLoad = -1
End Sub


Private Sub CalPagMsk_LostFocus()

With CalPagMsk

If Val(.Text) > gcMAXPAGEVALUE Or Val(.Text) < 1 Then
    .Text = IIf(Val(.Text) > gcMAXPAGEVALUE, gcMAXPAGEVALUE, gcMINPAGEVALUE)
    SpnPageNum.Value = Val(.Text)
End If

End With

End Sub

Sub CalPodmnu()
Dim MsgStr As String, RetInt As Integer
Dim SQL As String
On Local Error GoTo CalPod_Err
' --------------------
' Make sure day exists
' --------------------
SQL = "Date = " & "#" & Format$(DayMonthFocus, "MM-DD-YY") & "#"
ObjTour.DBFindFirst SQL, lEventHandle
'setCal.FindFirst SQL
If ObjTour.NoMatch(lEventHandle) Then
        MsgBox "No activity to delete?", vbOKOnly + vbQuestion, "TourWin Warning"
        Exit Sub
End If
' -------------------------
' Record Exist therefore...
' -------------------------
MsgStr = "Are you sure you want to " & vbLf
MsgStr = MsgStr & "delete : " & Format$(DayMonthFocus, "mmmm")
MsgStr = MsgStr & Mid$(DayMonthFocus, 4, 2) & ", "
MsgStr = MsgStr & Mid$(DayMonthFocus, 7, 4)

RetInt = MsgBox(MsgStr, vbYesNoCancel + 64, "TourWin Warning")
If RetInt = vbYes Then
    DeleteDays DayMonthFocus, DayMonthFocus
    CalBegmnu
End If
Exit Sub
CalPod_Err:
    If bDebug Then Handle_Err Err, "CalPodmnu_Click-Calndfrm"
    Resume Next
End Sub

Sub CalPoemnu()
Dim SQL As String
On Local Error GoTo CalPoe_Err
' --------------------
' Make sure day exists
' --------------------
SQL = "Date = " & "#" & Format$(DayMonthFocus, "MM-DD-YY") & "#"
ObjTour.DBFindFirst SQL, lEventHandle
'setCal.FindFirst SQL
If ObjTour.NoMatch(lEventHandle) Then
        MsgBox "Nothing to edit?", vbOKOnly, LoadResString(gcTourVersion)
        Exit Sub
Else
        Month_Clicked Format$(DayMonthFocus, "dd"), Format$(DayMonthFocus, "MM")
End If
Exit Sub
CalPoe_Err:
    If bDebug Then Handle_Err Err, "CalPoe_Err-Calndfrm"
    Resume Next
End Sub


Private Sub CalPagMsk_ToolTip()
    MDI.StatusBar1.Panels(1).Text = CalPagMsk.ToolTipText
End Sub

Private Sub CalTypCob_Click()

    If 0 = IsLoad Then Exit Sub
        
    IsLoad = 0
    RestCalendar
    IsLoad = -1

End Sub

Private Sub CalYeaTxt_Change()

    If 0 = IsLoad Then Exit Sub
    
    If Val(CalYeaTxt.Text) > spnYear.Max _
    Or Val(CalYeaTxt.Text) < spnYear.Min Then Exit Sub
    
    spnYear.Value = Val(CalYeaTxt.Text)
    
    IsLoad = 0
    RestCalendar
    IsLoad = -1

End Sub

Private Sub CalYeaTxt_GotFocus()

With CalYeaTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub CalYeaTxt_KeyPress(KeyAscii As Integer)
    KeyAscii = TextValidate(KeyAscii, "Int Only", CalYeaTxt)
End Sub

Private Sub CalYeaTxt_LostFocus()

With CalYeaTxt

If Val(.Text) > gcMAXYEAR Or Val(.Text) < 1 Then
    .Text = IIf(Val(.Text) > gcMAXYEAR, gcMAXYEAR, gcMINYEAR)
    spnYear.Value = Val(.Text)
End If

End With


End Sub

Private Sub Dec_DblClick(Index As Integer)
On Local Error GoTo Dec_Clicked_Err
Month_Clicked Index, Dec(Index).Tag
Exit Sub
Dec_Clicked_Err:
        If bDebug Then Handle_Err Err, "Dec_Clicked-CalndFrm"
        Resume Next
End Sub


Private Sub Dec_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "12-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub




Private Sub EveFilCbo_Click()
   If -1 = IsLoad Then
    IsLoad = 0
    LoadTypeNames EveFilCbo.Text, CalTypCob.Text
    RestCalendar
    IsLoad = -1
   End If
End Sub

Private Sub EveFilCbo_KeyPress(KeyAscii As Integer)
Beep
KeyAscii = 0
End Sub


Private Sub EveFilCbo_LostFocus()
On Local Error GoTo EveFil_Err
Dim CurrentCentry As String
' ------------------
' Update CalTypCob with
' appropriate list.
' ------------------
IsLoad = 0
If cFileChange <> EveFilCbo.Text Then
    Select Case EveFilCbo
            Case "Peak":
                CalTypCob.Clear
    End Select
    
    cFileChange = EveFilCbo.Text
    LoadTypeNames EveFilCbo.Text, ""
End If

'LoadTypeNames EveFilCbo.Text, ""
IsLoad = -1
Exit Sub
EveFil_Err:
    If bDebug Then Handle_Err Err, "EveFilCbo-Calndfrm"
    Resume Next
End Sub


Private Sub EveFilCbo_Scroll()
   If -1 <> IsLoad Then
    MDI.MdiBegmnu_Click
   End If
End Sub

Private Sub Feb_DblClick(Index As Integer)
On Local Error GoTo Feb_Clicked_Err
Month_Clicked Index, Feb(Index).Tag
Exit Sub
Feb_Clicked_Err:
        If bDebug Then Handle_Err Err, "Mar_Clicked-CalndFrm"
        Resume Next
End Sub


Private Sub Feb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "02-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub


Private Sub Form_Activate()
    Define_CalndFrm_mnu Loadmnu
End Sub

Private Sub Form_Deactivate()
    Define_CalndFrm_mnu Unloadmnu
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo Form_Err

Calndfrm.KeyPreview = True


'ObjTour.RstSQL lEventHandle, "SELECT * FROM " _
                            & gcEve_Tour_Event_Daily _
                            & " WHERE ID = " & objMdi.info.ID
                               
'lEventHandle = ObjTour.GetHandle
'ObjTour.DBOpen gcEve_Tour, gcEve_Tour_Event_Daily, lEventHandle

Define_CalndFrm_mnu Loadmnu

'-----------------------------------------------------
' Load Database names into EveFilCbo object
' and set default values for list box and text boxes
'-----------------------------------------------------
' Set Spin Control range
SpnPageNum.Min = 1
SpnPageNum.Max = gcMAXPAGEVALUE

spnYear.Min = gcMINYEAR
spnYear.Max = gcMAXYEAR

IsLoad = 0
LoadFile        ' Load Hard code values for EveFilCbo
IsLoad = -1


RestCalendar
'Show companion form
ObjChart.ShowChart EveFilCbo.Text

cFileChange = EveFilCbo.Text
cEventChange = CalTypCob.Text
cYearChange = CalYeaTxt.Text

CentreForm Calndfrm, -1
IsLoad = -1


Exit Sub
Form_Err:
    If bDebug Then Handle_Err Err, "Form_Load-Calndfrm"
    Resume Next
End Sub

Sub Define_CalndFrm_mnu(LoadType As String)
On Local Error GoTo Define_CalndFrm_Err

If Not objMdi.info.NewUser Then
Select Case LoadType
    Case Unloadmnu:
        MDI!MdiEximnu.Caption = MdiFrm_Exit
        MDI!MdiOptmnu.Caption = MdiFrm_Option
        MDI!MdiNewmnu.Caption = MdiFrm_Newmnu
        MDI!MdiSavmnu.Caption = MdiFrm_Savmnu
        MDI!MdiDelmnu.Caption = MdiFrm_Delmnu
        MDI!MdiOptmnu.Enabled = True
        MDI!MdiCalmnu.Enabled = True
        MDI!MdiNewmnu.Enabled = True
        MDI!MdiSavmnu.Enabled = False
        MDI!MdiDelmnu.Enabled = False
'        MDI!MdiBegmnu.Visible = False
        MDI!MdiPrimnu.Enabled = False
        
    Case Loadmnu:
        MDI!MdiEximnu.Caption = CalndFrm_Exit
        MDI!MdiOptmnu.Caption = CalndFrm_Option
        MDI!MdiNewmnu.Caption = CalndFrm_Newmnu
        MDI!MdiSavmnu.Caption = CalndFrm_Savmnu
        MDI!MdiDelmnu.Caption = CalndFrm_Delmnu
        MDI!MdiBegmnu.Caption = CalndFrm_Begmnu
        MDI!MdiOptmnu.Enabled = False
        MDI!MdiCalmnu.Enabled = False
        MDI!MdiNewmnu.Enabled = False
        MDI!MdiSavmnu.Enabled = False
        MDI!MdiDelmnu.Enabled = True
'        MDI!MdiBegmnu.Enabled = True
'        MDI!MdiBegmnu.Visible = True
        MDI!MdiPrimnu.Enabled = True
End Select
End If
Exit Sub
Define_CalndFrm_Err:
    If bDebug Then Handle_Err Err, "Define_CalndFrm_mnu-CalndFrm"
    Resume Next
End Sub



Private Sub Form_Unload(Cancel As Integer)

    ' Write Forms Settings to Registry
    ObjTour.Settings.WriteFormSettingsToReg Me
    
    Define_CalndFrm_mnu Unloadmnu
    
    ' Free DB
    ObjTour.DBClose lEventHandle
    ObjTour.FreeHandle lEventHandle
    
    ObjChart.CloseChart
    Unload Me
End Sub


Private Sub Jan_DblClick(Index As Integer)
On Local Error GoTo Jan_Clicked_Err
Month_Clicked Index, Jan(Index).Tag
If UserCancel = False Then CalBegmnu
Exit Sub
Jan_Clicked_Err:
        If bDebug Then Handle_Err Err, "Jan_Clicked-CalndFrm"
        Resume Next
End Sub


Private Sub Jan_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "01-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub

Private Sub Jul_DblClick(Index As Integer)
On Local Error GoTo Jul_Clicked_Err
Month_Clicked Index, Jul(Index).Tag
Exit Sub
Jul_Clicked_Err:
        If bDebug Then Handle_Err Err, "Jul_Clicked-CalndFrm"
        Resume Next
End Sub


Private Sub Jul_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "07-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub


Private Sub Jun_DblClick(Index As Integer)
On Local Error GoTo Jun_Clicked_Err
Month_Clicked Index, Jun(Index).Tag
Exit Sub
Jun_Clicked_Err:
        If bDebug Then Handle_Err Err, "Jun_Clicked-CalndFrm"
        Resume Next
End Sub


Private Sub Jun_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "06-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub


Private Sub Mar_DblClick(Index As Integer)
On Local Error GoTo Mar_Clicked_Err
Month_Clicked Index, Mar(Index).Tag
Exit Sub
Mar_Clicked_Err:
        If bDebug Then Handle_Err Err, "Mar_Clicked-CalndFrm"
        Resume Next
End Sub



Private Sub Mar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "03-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub

Private Sub May_DblClick(Index As Integer)
On Local Error GoTo May_Clicked_Err
Month_Clicked Index, May(Index).Tag
Exit Sub
May_Clicked_Err:
        If bDebug Then Handle_Err Err, "May_Clicked-CalndFrm"
        Resume Next
End Sub

Private Sub May_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "05-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub


Private Sub Nov_DblClick(Index As Integer)
On Local Error GoTo Nov_Clicked_Err
Month_Clicked Index, Nov(Index).Tag
Exit Sub
Nov_Clicked_Err:
        If bDebug Then Handle_Err Err, "Nov_Clicked-CalndFrm"
        Resume Next

End Sub


Private Sub Nov_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "11-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub


Private Sub Oct_DblClick(Index As Integer)
On Local Error GoTo Oct_Clicked_Err
Month_Clicked Index, Oct(Index).Tag
Exit Sub
Oct_Clicked_Err:
        If bDebug Then Handle_Err Err, "Oct_Clicked-CalndFrm"
        Resume Next

End Sub


Private Sub Oct_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "10-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If

End Sub


Private Sub Sep_DblClick(Index As Integer)
On Local Error GoTo Sep_Clicked_Err
Month_Clicked Index, Sep(Index).Tag
Exit Sub
Sep_Clicked_Err:
        If bDebug Then Handle_Err Err, "Sep_Clicked-CalndFrm"
        Resume Next
End Sub


Private Sub Sep_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
DayMonthFocus = "09-" & Format$(Index, "00") & "-" & CalYeaTxt
Calndfrm.PopupMenu MDI.CalPopmnu
End If
End Sub


Private Sub SpnPageNum_Change()
On Local Error GoTo SpnPage_Err

CalPagMsk.Text = Format$(SpnPageNum.Value, "0#")


On Local Error GoTo 0
Exit Sub
SpnPage_Err:
    Call Handle_Err(Err, "SpnPageNum_Change-Calndfrm")
    Resume Next


End Sub

Private Sub spnYear_Change()
On Local Error GoTo SpinYear_Err

CalYeaTxt.Text = Format$(spnYear.Value, "0000")

If cYearChange <> CalYeaTxt.Text Then
   cYearChange = CalYeaTxt.Text
End If

On Local Error GoTo 0

Exit Sub
SpinYear_Err:
    Call Handle_Err(Err, "SpnYear_Change-Calndfrm")
    Resume Next

End Sub
