VERSION 5.00
Begin VB.Form ConcFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conconi Test Results."
   ClientHeight    =   5175
   ClientLeft      =   1635
   ClientTop       =   1680
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CONCFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5175
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   " Test Data "
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   3135
      Begin Tourwin2002.SimpleGrid SimpleGrid1 
         Height          =   2775
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4895
         CellHeight      =   ""
         CellWidth       =   ""
         FormatStyle     =   ""
         CellWidth       =   ""
         CellHeight      =   ""
         Columns         =   3
         HeaderCaption   =   "<0>Time (min)</0><1>Heart Rate </1><2>Watts</2>"
         RowBuffer       =   9
         SortBy          =   ""
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Evaluation"
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdResults 
         Height          =   255
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin Tourwin2002.UTextBox ConDatTxt 
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         ToolTipText     =   "Date of Conconi test (mm-dd-yyyy)"
         Top             =   240
         Width           =   975
         _ExtentX        =   2143
         _ExtentY        =   661
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
      Begin Tourwin2002.UTextBox mskPedal 
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "Average cadence for text (Range 0 to 999)"
         Top             =   600
         Width           =   615
         _ExtentX        =   2143
         _ExtentY        =   661
         Max             =   3
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
      Begin Tourwin2002.UTextBox mskDuration 
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         ToolTipText     =   "Duration of text (hh:mm:ss)"
         Top             =   960
         Width           =   975
         _ExtentX        =   2143
         _ExtentY        =   661
         Max             =   8
         FieldType       =   3
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
      Begin Tourwin2002.UTextBox mskMaxPulse 
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         ToolTipText     =   "Maximum heart rate obtained (Range 0 to 999)"
         Top             =   1320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Max             =   3
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
      Begin VB.Label Label1 
         Caption         =   "Maximum &pulse rate:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblDuration 
         Caption         =   "T&otal duration:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label CncPedLbl 
         Caption         =   "Pedaling fre&quency:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label CncDatLbl 
         AutoSize        =   -1  'True
         Caption         =   "&Date of Test:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Graphical Results "
      Height          =   4935
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   5535
      Begin VB.CheckBox chkLegend 
         Caption         =   "Display &Legend"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   3000
         TabIndex        =   12
         Top             =   275
         Width           =   1575
      End
      Begin VB.CheckBox chkPlot 
         Caption         =   "Plot Poi&nts"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   1680
         TabIndex        =   11
         Top             =   275
         Width           =   1095
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Dra&w Graph"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1275
      End
      Begin Tourwin2002.GraphLite graConconi 
         Height          =   4215
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   5055
         _ExtentX        =   9128
         _ExtentY        =   7435
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "ConcFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const GRIDROWS = 16
Private Const MINUTES_DATA = 0
Private Const PULSE_DATA = 1
Private Const WATTS_DATA = 2
Dim MyData() As Variant  ' Holds user data
Dim bChanged As Boolean
Dim lConconiHandle As Long

Private mTotalRows& ' Contains the total rows in the set of records
Private UserData() As Variant ' 2-dimensional array containing records
Private Const MAXCOLS = 3 ' Maximum number of fields in record set.

Private bLoading    As Boolean

Enum ConconiLoadType
   FetchFromRecordset = 0
   WriteToRecordset = 1
   RestControls = 2
End Enum

Private Function FillGrid(ByVal lHandle As Long) As Boolean

' 3 Columns, 15 Rows of Data
Dim lCount  As Long
Dim sRow    As String
Dim sKey    As String ' Key is Date + ID + Min

lCount = ObjTour.RstRecordCount(lHandle)

ReDim UserData(0 To 2, 0 To lCount)

ObjTour.DBMoveFirst lHandle
lCount = 0
Do
    
    UserData(MINUTES_DATA, lCount) = ObjTour.DBGetField(gcData_Minute, lHandle)
    UserData(PULSE_DATA, lCount) = ObjTour.DBGetField(gcData_PulseRate, lHandle)
    UserData(WATTS_DATA, lCount) = ObjTour.DBGetField(gcData_Watt, lHandle)
    
    sRow = "<0>" & ObjTour.DBGetField(gcData_Minute, lHandle) & "<0/>"
    sRow = sRow & "<1>" & ObjTour.DBGetField(gcData_PulseRate, lHandle) & "<1/>"
    sRow = sRow & "<2>" & ObjTour.DBGetField(gcData_Watt, lHandle) & "<2/>"
    
    sKey = ObjTour.DBGetField(gcData_Date, lHandle) & vbTab & _
           ObjTour.DBGetField(gcID, lHandle) & vbTab & _
           ObjTour.DBGetField(gcData_Minute, lHandle)
           
SimpleGrid1.AddRow sKey, sRow

    ObjTour.DBMoveNext (lHandle)
    lCount = lCount + 1
Loop While Not ObjTour.EOF(lHandle)

SimpleGrid1.RowCount = lCount
mTotalRows& = lCount

End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : LoadConconiTest
' DATE      : 3/11/03 12:25
' Author    : mor
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function LoadConconiTest(ByVal eType As ConconiLoadType, Optional ByVal lHandle As Long = 0) As Boolean
On Error GoTo UpdateInfo_Error

' Declare Local Variables
Dim SQL As String

Select Case eType
        Case ConconiLoadType.FetchFromRecordset:
                ConDatTxt.Text = Format$(ObjTour.DBGetField(gcDai_Tour_Date, lHandle), "MM/dd/yyyy")
                mskPedal.Text = ObjTour.DBGetField(gcDai_Tour_DayType, lHandle)
                mskDuration.Text = ObjTour.DBGetField(gcDai_Tour_TotalHrs, lHandle)
                ' Max Heart Rate
                mskMaxPulse.Text = ObjTour.DBGetField(gcDai_Tour_Heart, lHandle)
                
                ' Load Test Data
                Call LoadTestData(CDate(ConDatTxt.Text))
                
        Case ConconiLoadType.WriteToRecordset:
                ' Determine if record exists, if not create, else update

                SQL = " SELECT * FROM " & gcDai_Tour_Dai & _
                      " WHERE ID = " & objMdi.info.ID & " AND " & _
                      "      Type = " & gcDAI_CONCONI & " AND " & _
                      "      Date = #" & ConDatTxt.Text & "#"
                      
                ObjTour.RstSQL lHandle, SQL
                If ObjTour.RstRecordCount(lHandle) > 0 Then
                    ObjTour.Edit lHandle
                Else
                    ObjTour.AddNew lHandle
                End If
                    ' Date
                    ObjTour.DBSetField gcDai_Tour_Date, ConDatTxt.Text, lHandle
                    ' Dai record type
                    ObjTour.DBSetField gcDai_Tour_DaiType, gcDAI_CONCONI, lHandle
                    ' Owner ID
                    ObjTour.DBSetField gcDai_Tour_Id, objMdi.info.ID, lHandle
                    ' Pedal
                    ObjTour.DBSetField gcDai_Tour_DayType, mskPedal.Text, lHandle
                    ' Total Duration
                    ObjTour.DBSetField gcDai_Tour_TotalHrs, mskDuration.Text, lHandle
                    ' Max Heart Rate
                    ObjTour.DBSetField gcDai_Tour_Heart, mskMaxPulse.Text, lHandle
                
                ObjTour.Update lHandle
                
        Case ConconiLoadType.RestControls:
                ConDatTxt.Text = Format$(Now, "MM/dd/yyyy")
                mskPedal.Text = "000"
                mskDuration.Text = "00:00:00"
                mskMaxPulse.Text = "000"
                
End Select

LoadConconiTest = True
On Error GoTo 0
Exit Function

UpdateInfo_Error:
    If bDebug Then Handle_Err Err, "UpdateInfo()-Concfrm"
    Resume Next
End Function


'---------------------------------------------------------------------------------------
' PROCEDURE : DeleteMenu_Click
' DATE      : 3/11/03 12:23
' Author    : mor
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub DeleteMenu_Click()
On Local Error GoTo CncDel_Err
Dim MsgStr As String
Dim lRt As Long
Dim SQL As String


MsgStr = "Are you sure you wish to delete this Conconi test"
lRt = MsgBox(MsgStr, vbYesNo + vbQuestion, LoadResString(gcTourVersion))

If lRt = vbYes Then
    SQL = " DELETE * FROM " & gcDai_Tour_Dai & _
          " WHERE ID = " & objMdi.info.ID & " AND " & _
          "      Type = " & gcDAI_CONCONI & " AND " & _
          "      Date = #" & ConDatTxt.Text & "#"
    ObjTour.DBExecute SQL
    
    ' Now, delete all records for the specified date range...
    SQL = " DELETE * FROM " & gcData & _
          " WHERE ID = " & objMdi.info.ID & " AND " & _
          gcData_Date & " = #" & Format$(ConDatTxt.Text, "MM/dd/yyyy") & "#"
    
    ObjTour.DBExecute SQL
' Now find an existing test
    LoadConconiTest ConconiLoadType.RestControls
End If
    
On Error GoTo 0
Exit Sub
CncDel_Err:
    If bDebug Then Handle_Err Err, "CncDelmnu-Concfrm"
    Resume Next
End Sub

Public Sub NewMenu_Click()
On Local Error GoTo CncNewmnu_Err

' Declare local variables


'check if current conconi test has changed
'and needs to be saved?
If bChanged Then
    If vbYes = MsgBox("Save changes", vbYesNo + vbQuestion, LoadResString(gcTourVersion)) Then
        SaveMenu_Click
    End If
End If


LoadConconiTest ConconiLoadType.RestControls
bChanged = False

Exit Sub
CncNewmnu_Err:
    If bDebug Then Handle_Err Err, "CncNewmnu_Click-ConcFrm"
    Resume Next
End Sub
Public Sub SaveMenu_Click()
On Local Error GoTo SaveMenu_Error

    'LoadConconiTest WriteToRecordset, lConconiHandle
    SaveConconiTest

On Local Error GoTo 0
Exit Sub

SaveMenu_Error:
    If bDebug Then Handle_Err Err, "SaveMenu_Click-ConcFrm"
    Resume Next

End Sub

Public Sub OptionMenu_Click()
    MsgBox "Current not avaiable", 0, "TourWin"

End Sub
Private Sub chkLegend_Click()

If bLoading Then Exit Sub

graConconi.DisplayLegend = chkLegend.Value * -1
graConconi.Refresh

End Sub

Private Sub chkPlot_Click()

If bLoading Then Exit Sub

graConconi.PlotPoints = chkPlot.Value * -1
graConconi.Refresh

End Sub

Private Sub cmdDraw_Click()
On Local Error GoTo Draw_Error

' Declare Local variables...
Dim DataPoints As Integer
Dim i       As Integer
Dim cSlope  As Currency  ' Holds slope value...
Dim lLow    As Long
Dim lHigh   As Long

' Initialize local variables
lLow = 150
lHigh = 160
DataPoints = UBound(UserData, 2) - 1


' Make sure there are enough Data Points
' to graph...
If DataPoints < 2 Then
    GoTo InsufficientData
End If
' change back to 2 to add slope...
ReDim MyData(1, DataPoints) As Variant

For i = 0 To UBound(UserData, 2) - 1
   MyData(0, i) = UserData(2, i)    ' Watts
   MyData(1, i) = UserData(1, i)    ' UserData
    If lLow > MyData(1, i) Then lLow = MyData(1, i)
    If lHigh < MyData(1, i) Then lHigh = MyData(1, i)
Next i

graConconi.BackColor = &HFFFFFF 'white
graConconi.RegisterData MyData()
graConconi.SetSeriesOptions 0, vbBlue, "Watts Vs. Heart Rate"
'graConconi.SetSeriesOptions 1, vbRed, "Aerobic Threshold"
graConconi.Title = "Conconi Test"
graConconi.LowScale = lLow
graConconi.HighScale = lHigh
graConconi.VerticalTickInterval = 5
graConconi.HorizontalTickFrequency = 7
graConconi.ChartType = 1

graConconi.Refresh
Exit Sub

Draw_Error:
    If bDebug Then Handle_Err Err, "cmdDraw_Click-ConcFrm"
    If Err.Number = 9 Then GoTo InsufficientData
    Err.Clear
    Resume Next

InsufficientData:
    MsgBox "Insufficient data to draw conconi graph, please ensure 4 more data points.", vbOKOnly + vbInformation, LoadResString(gcTourVersion)
End Sub

Private Sub cmdResults_Click()

Dim SQL As String

' show list of date for recorded Conconi test
frmConconiList.SelectedDate = CDate(Me.ConDatTxt.Text)
frmConconiList.Show vbModal

If frmConconiList Is Nothing Then Exit Sub
If frmConconiList.Conconi_UserCancelled Then Exit Sub

If bChanged Then
    'Prompt user to save changes
    If vbYes = MsgBox("Save Changes?", vbYesNo + vbQuestion, " Save ") Then
        SaveMenu_Click
    End If
End If

'3. Select * from Dai where type = gcDAI_CONCONI
SQL = " SELECT * FROM " & gcDai_Tour_Dai & _
      " WHERE ID = " & objMdi.info.ID & " AND " & _
      "      Type = " & gcDAI_CONCONI & " AND " & _
      "      [Date] = #" & Format$(frmConconiList.SelectedDate, "MM/dd/yyyy") & "#"

ObjTour.RstSQL lConconiHandle, SQL
If ObjTour.RstRecordCount(lConconiHandle) > 0 Then
    Call LoadConconiTest(FetchFromRecordset, lConconiHandle)
    
End If
Unload frmConconiList
cmdDraw_Click

bChanged = False

End Sub

Private Sub ConDatTxt_Change()
bChanged = True
End Sub




Private Sub ConDatTxt_ToolTip()
MDI.StatusBar1.Panels(1).Text = ConDatTxt.ToolTipText
End Sub



Private Sub Form_Activate()
Define_Form_menu Me.Name, Loadmnu
End Sub

Private Sub Form_Deactivate()
Define_Form_menu Me.Name, Unloadmnu
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
    If vbKeyF1 = KeyAscii Then
        HTMLHelp Me.hWnd, App.Path & "\TourWin.chm", HH_DISPLAY_TOPIC, 0
    End If
End Sub

Private Sub Form_Load()

On Local Error GoTo Conc_Load_Err

' Declare local variables
Dim SQL As String

bLoading = True
ConcFrm.KeyPreview = True

' Update MDI menu
Define_Form_menu Me.Name, Loadmnu

' Setup Grid with title for columns
DefineGridHeader

LoadSaveSettings True

'3. Select * from Dai where type = gcDAI_CONCONI
SQL = " SELECT * FROM " & gcDai_Tour_Dai & _
      " WHERE ID = " & objMdi.info.ID & " AND " & _
      "      Type = " & gcDAI_CONCONI & _
      " ORDER BY Date ASC"

ObjTour.RstSQL lConconiHandle, SQL

If ObjTour.RstRecordCount(lConconiHandle) > 0 Then
' Populate with current values
    LoadConconiTest ConconiLoadType.FetchFromRecordset, lConconiHandle
Else
' Put in todays date
    LoadConconiTest ConconiLoadType.RestControls
End If

ObjTour.FreeHandle lConconiHandle

bChanged = False

' I don't know why, but frames backcolor is
' different from the rest...
Frame1.BackColor = Me.BackColor

cmdDraw_Click

bLoading = False

On Error GoTo 0

Exit Sub
Conc_Load_Err:
    If bDebug Then Handle_Err Err, "Form_Load-ConcFrm"
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error GoTo Unload_Error

Define_Form_menu Me.Name, Unloadmnu
If bChanged Then
    'Prompt user to save changes
    If vbYes = MsgBox("Save Changes?", vbYesNo + vbQuestion, " Save ") Then
        SaveMenu_Click
    End If
End If


' See cMDIVar for BitField settings
With objMdi.info.UserOptions
    If .GetBool(BitFlags.Conconi_LoadSettings) Then
        LoadSaveSettings (False)
    End If
End With


On Local Error GoTo 0

Exit Sub

Unload_Error:
If bDebug Then Handle_Err Err, "Form_UnLoad-ConcFrm"
MsgBox Err.Description
Resume Next

End Sub

Private Sub mskDuration_Change()
bChanged = True
End Sub



Private Sub mskMaxPulse_Change()
bChanged = True
End Sub



Private Sub mskPedal_Change()
bChanged = True
End Sub

Private Function LoadTestData(ByVal dDate As Date) As Boolean
On Local Error GoTo LoadTestData_Error

' Declare Local Variables
Dim SQL As String
Dim lDataHandle As Long
' First, clear the existing information

SQL = " SELECT * FROM " & gcData & _
      " WHERE ID = " & objMdi.info.ID & " AND " & _
      gcData_Date & " = #" & Format$(dDate, "MM/dd/yyyy") & "# " & _
      " ORDER BY " & gcData_Minute & " ASC"
      
ObjTour.RstSQL lDataHandle, SQL

If ObjTour.RstRecordCount(lDataHandle) > 0 Then
    ' Loop through each item and add to grid
    Call FillGrid(lDataHandle)
End If

ObjTour.FreeHandle lDataHandle

On Error GoTo 0
Exit Function

LoadTestData_Error:
    If bDebug Then Handle_Err Err, "LoadTestData-ConcFrm"
    Resume Next

End Function

Private Function DefineGridHeader() As Boolean

Dim m_ColumnsWidths     As String

m_ColumnsWidths = "<0>850<0/><1>850<1/><2>700<2/>"

Call SimpleGrid1.GridSettings(3, _
                             "<0>Time (min)<0/><1>Heart Rate<1/><2>Watts<2/>", _
                             CStr(dbText), _
                             m_ColumnsWidths)

End Function

Private Function SaveConconiTest() As Boolean

On Local Error GoTo SaveConconi_Error
'Declare local variables
Dim SQL As String
Dim lHandle As Long


' Determine if record exists, if not create, else update

    SQL = " SELECT * FROM " & gcDai_Tour_Dai & _
          " WHERE ID = " & objMdi.info.ID & " AND " & _
          "      Type = " & gcDAI_CONCONI & " AND " & _
          "      Date = #" & ConDatTxt.Text & "#"
          
    ObjTour.RstSQL lHandle, SQL
    If ObjTour.RstRecordCount(lHandle) > 0 Then
        ObjTour.Edit lHandle
    Else
        ObjTour.AddNew lHandle
    End If
        ' Date
        ObjTour.DBSetField gcDai_Tour_Date, ConDatTxt.Text, lHandle
        ' Dai record type
        ObjTour.DBSetField gcDai_Tour_DaiType, gcDAI_CONCONI, lHandle
        ' Owner ID
        ObjTour.DBSetField gcDai_Tour_Id, objMdi.info.ID, lHandle
        ' Pedal
        ObjTour.DBSetField gcDai_Tour_DayType, mskPedal.Text, lHandle
        ' Total Duration
        ObjTour.DBSetField gcDai_Tour_TotalHrs, mskDuration.Text, lHandle
        ' Max Heart Rate
        ObjTour.DBSetField gcDai_Tour_Heart, mskMaxPulse.Text, lHandle
    
    ObjTour.Update lHandle
    ObjTour.FreeHandle lHandle
    
    SaveTestData (CDate(ConDatTxt.Text))
    
On Error GoTo 0
Exit Function

SaveConconi_Error:
    If bDebug Then Handle_Err Err, "SaveConconiTest-ConcFrm"
    Resume Next

End Function
'---------------------------------------------------------------------------------------
' PROCEDURE : SaveTestData
' DATE      : 3/13/03 09:40
' Author    : Mark Ormesher
' Purpose   : This function first deletes all data records for the current date
'             then inserts the data values one at a time.
'---------------------------------------------------------------------------------------
Private Function SaveTestData(ByVal dDate As Date) As Boolean

On Local Error GoTo SaveTestData_Error

'Declare local variables
Dim SQL As String
Dim lDataHandle As Long
Dim iLoop As Long

' First, initial lDataHandle with a SELECT statement
SQL = " SELECT * FROM " & gcData & _
      " WHERE ID = " & objMdi.info.ID & " AND " & _
      gcData_Date & " = #" & Format$(dDate, "MM/dd/yyyy") & "#"
      
ObjTour.RstSQL lDataHandle, SQL

' Now, delete all records for the specified date range...
SQL = " DELETE * FROM " & gcData & _
      " WHERE ID = " & objMdi.info.ID & " AND " & _
      gcData_Date & " = #" & Format$(dDate, "MM/dd/yyyy") & "#"

ObjTour.DBExecute SQL

' Now, insert each UserData() array element
For iLoop = 0 To UBound(UserData, 2)
If Not IsEmpty(UserData(PULSE_DATA, iLoop)) Or Not IsEmpty(UserData(PULSE_DATA, iLoop)) Then
    SQL = "Insert into " & gcData & "(ID ," & _
                                    "[" & gcData_Date & "], " & _
                                    gcData_Minute & ", " & _
                                    gcData_PulseRate & ", " & _
                                    gcData_Watt & ") Values (" & _
                                    objMdi.info.ID & ", " & _
                                    "#" & Format$(dDate, "MM/dd/yyyy") & "#, " & _
                                    CStr(iLoop) & ", " & _
                                    UserData(PULSE_DATA, iLoop) & ", " & _
                                    UserData(WATTS_DATA, iLoop) & ")"
    ObjTour.DBExecute SQL
End If
Next iLoop

ObjTour.FreeHandle lDataHandle

SaveTestData = True
bChanged = False

On Error GoTo 0
Exit Function

SaveTestData_Error:
    If bDebug Then Handle_Err Err, "SaveTestData-ConcFrm"
    Resume Next
    
End Function


Private Sub LoadSaveSettings(ByVal bLoadSettings As Boolean)

On Local Error GoTo LoadSave_Error

If bLoadSettings Then
    With objMdi.info.UserOptions
    
      chkPlot.Value = .GetValue(BitFlags.Conconi_PlotPoints)
      graConconi.PlotPoints = .GetBool(BitFlags.Conconi_PlotPoints)

      chkLegend.Value = .GetValue(BitFlags.Conconi_DrawLegend)
      graConconi.DisplayLegend = .GetBool(BitFlags.Conconi_DrawLegend)
      
    End With
        
Else 'Save

 With objMdi.info.UserOptions
 
    .SetValue chkLegend.Value, BitFlags.Conconi_DrawLegend
    .SetValue chkPlot.Value, BitFlags.Conconi_PlotPoints
    objMdi.SaveUserSettings
 End With

End If


On Local Error GoTo 0
Exit Sub

LoadSave_Error:
    If bDebug Then Handle_Err Err, "LoadSaveSettings-ConcFrm"
    Resume Next

End Sub

Private Sub mskDuration_ToolTip()
MDI.StatusBar1.Panels(1).Text = mskDuration.ToolTipText
End Sub

Private Sub mskMaxPulse_ToolTip()
MDI.StatusBar1.Panels(1).Text = mskMaxPulse.ToolTipText
End Sub

Private Sub mskPedal_ToolTip()
MDI.StatusBar1.Panels(1).Text = mskPedal.ToolTipText
End Sub

Private Sub SimpleGrid1_DeleteRow(ByVal RowID As Long, ByVal sKey As String)

If DeleteRow(sKey) Then
    
    SimpleGrid1.DeleteRow RowID, sKey
    
End If

End Sub

Private Sub SimpleGrid1_EditRow(ByVal RowID As Long, ByVal sKey As String, ByVal lColumnNo As Long, ByVal sValue As String)

    Call EditRow(sKey, lColumnNo, sValue)
    
End Sub


'---------------------------------------------------------------------------------------
' PROCEDURE : DeleteRow
' DATE      : 9/11/04 14:39
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function DeleteRow(ByVal sKey As String) As Boolean
On Local Error GoTo DeleteRow_Error

Const DateInfo = 0
Const IDInfo = 1
Const MinInfo = 2

'Declare local variables
Dim SQL As String
Dim lDataHandle As Long
Dim vKey    As Variant

'Pesimistic
DeleteRow = False

' Split key into Date, id, and Min
vKey = Split(sKey, vbTab)

' First, clear the existing information
SQL = " SELECT * FROM " & gcData & _
      " WHERE ID = " & vKey(IDInfo) & " AND " & _
        gcData_Date & " = #" & Format$(vKey(DateInfo), "MM/dd/yyyy") & "# " & _
      " AND " & gcData_Minute & " = " & vKey(MinInfo)
      
ObjTour.RstSQL lDataHandle, SQL

If ObjTour.RstRecordCount(lDataHandle) > 0 Then
    ' Loop through each item and add to grid
   ObjTour.Delete lDataHandle
   DeleteRow = True
End If

ObjTour.FreeHandle lDataHandle

On Error GoTo 0
Exit Function

DeleteRow_Error:
    If bDebug Then Handle_Err Err, "DeleteRow-ConcFrm"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : EditRow
' DATE      : 9/11/04 15:00
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function EditRow(ByVal sKey As String, ByVal lColumn As Long, ByVal sValue As String) As Boolean
On Local Error GoTo EditRow_Error

Const DateInfo = 0
Const IDInfo = 1
Const MinInfo = 2

Const TimeInfo = 0
Const HeartInfo = 1
Const WattInfo = 2

'Declare local variables
Dim SQL As String
Dim oRecord     As New CRecord
Dim vKey    As Variant

'Pesimistic
EditRow = False

' Split key into Date, id, and Min
vKey = Split(sKey, vbTab)

' First, clear the existing information
SQL = " SELECT * FROM " & gcData & _
      " WHERE ID = " & vKey(IDInfo) & " AND " & _
        gcData_Date & " = #" & Format$(vKey(DateInfo), "MM/dd/yyyy") & "# " & _
      " AND " & gcData_Minute & " = " & vKey(MinInfo)
      
With oRecord

    .RstSQL SQL
    If .RstRecordCount() > 0 Then
        ' Loop through each item and add to grid
        .Edit
        Select Case lColumn
             Case TimeInfo:
                    .SetField gcData_Minute, sValue
             Case HeartInfo:
                    .SetField gcData_PulseRate, sValue
             Case WattInfo:
                    .SetField gcData_Watt, sValue
             
        End Select
        .Update
        EditRow = True
    End If
End With

On Error GoTo 0
Exit Function

EditRow_Error:
    If bDebug Then Handle_Err Err, "EditRow-ConcFrm"
    Resume Next


End Function

Private Sub SimpleGrid1_NewRow(ByVal RowID As Long, ByVal lColumnNo As Long, ByVal sValue As String)
On Local Error GoTo NewRow_Err


SimpleGrid1.UpdateRowKey RowID, NewRow(lColumnNo, sValue)

On Error GoTo 0
Exit Sub

NewRow_Err:
If bDebug Then Handle_Err Err, "NewRow-ConcFrm"
Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : NewRow
' DATE      : 9/11/04 21:58
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function NewRow(ByVal lColumn As Long, ByVal sValue As String) As String
On Local Error GoTo NewRow_Error
'Declare local variables

Const DateInfo = 0
Const IDInfo = 1
Const MinInfo = 2

Const TimeInfo = 0
Const HeartInfo = 1
Const WattInfo = 2

'Declare local variables
Dim SQL         As String
Dim oRecord     As New CRecord
Dim sKey        As String

'Pesimistic
NewRow = ""

' First, clear the existing information
SQL = " SELECT * FROM " & gcData '& _

With oRecord

    .RstSQL SQL
    .AddNew
    .SetField gcData_Date, ConDatTxt.Text
    .SetField gcID, objMdi.info.ID
        Select Case lColumn
             Case TimeInfo:
                    .SetField gcData_Minute, sValue
             Case HeartInfo:
                    .SetField gcData_PulseRate, sValue
             Case WattInfo:
                    .SetField gcData_Watt, sValue
             
        End Select
    .Update
End With

'Build Key for return
sKey = ConDatTxt.Text & vbTab
sKey = sKey & objMdi.info.ID & vbTab
If TimeInfo = lColumn Then
    sKey = sKey & sValue
Else
    sKey = sKey & ""
End If

NewRow = sKey

On Error GoTo 0
Exit Function

NewRow_Error:
    If bDebug Then Handle_Err Err, "NewRow-ConcFrm"
    Resume Next


End Function
