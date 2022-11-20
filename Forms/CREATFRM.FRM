VERSION 5.00
Begin VB.Form CreatFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "January 1st 1997"
   ClientHeight    =   3435
   ClientLeft      =   1590
   ClientTop       =   1635
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CREATFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3435
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame crelblPeak 
      Caption         =   "Peak Schedule"
      Height          =   2775
      Left            =   225
      TabIndex        =   9
      Top             =   120
      Width           =   5010
      Begin VB.Frame fraDaytype 
         Caption         =   " Daily type "
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   4695
         Begin VB.OptionButton CreDayOpt 
            Caption         =   "&Month"
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.OptionButton CreDayOpt 
            Caption         =   "&Week"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.OptionButton CreDayOpt 
            Caption         =   "&Day"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.TextBox CreMemTxt 
         Height          =   285
         Left            =   120
         MaxLength       =   255
         TabIndex        =   1
         Top             =   480
         Width           =   4575
      End
      Begin VB.CheckBox chkStartof 
         Alignment       =   1  'Right Justify
         Caption         =   "&First day of schedule"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox CreEveCbo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox CreConchk 
         Alignment       =   1  'Right Justify
         Caption         =   "Re&peat"
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox CreNotetxt 
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "CREATFRM.frx":000C
         Top             =   1920
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label lblNote 
         Caption         =   "N&ote:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label CreHidTxt 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblSchedule 
         Caption         =   "Sc&hedule:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "&Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton CreCloCmd 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CreSavCmd 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
End
Attribute VB_Name = "CreatFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim setEve As Dynaset, dbEve As Database
Dim TableType As String, RecExist As Boolean
'Dim dbPea As Database, setPea As Dynaset
Dim SQL As String
Dim CDynamicActivities As CDynamics


'Sub CreateEventDay()
'' --------------------
'' Review this code!!!
'' - setEve should not be used.
'' --------------------
'On Local Error GoTo CreEve_Err
'Dim SQL As String, EndDate As String
'SQL = "Date = " & "#" & Format$(CreatFrm.Caption, "mm-dd-yyyy") & "#"
'setEve.FindFirst SQL
'
'If setEve.RecordCount = 0 Then
'    setEve.AddNew
'        setEve("Date") = Format$(CreatFrm.Caption, "mm-dd-yyyy")
'        setEve("Evemeno") = CreMemTxt.Text
'        setEve("EveType") = CreEveCbo.Text
'        setEve("Color") = 40
'    setEve.Update
'    setEve.Close
'End If
'Exit Sub
'CreEve_Err:
'    If bDebug Then Handle_Err Err, "CreateEventDay-CreatFrm"
'    Resume Next
'End Sub

Function CreatePeakDays(dEndDate As Date) As Integer
'
' This sub creates
'

Dim i           As Integer
Dim iSched      As Integer
Dim SQL         As String
Dim Act         As Integer
Dim dDate       As Date
Dim iLoop       As Integer
Dim iEvent_ID   As Integer        'Track ID of newly created event...

Dim iCurrentPos As Integer
Dim sDeliminator As String
Dim sPosition   As String
Dim Sched       As String
Dim SchedLen    As Integer



On Local Error GoTo Create_Err
' ----------------------------
' Find the select Cycle type
' in Peaktour.mdb database and get
' peak parameters / Template
' -----------------------------

' Set Default values
objMdi.info.bStartOf = IIf(1 = chkStartof.Value, True, False)
objMdi.info.iInterations = 1 ' Note Loop is 0 to x -1
iSched = 1
Act = 0

' Get Schedule from iPeakHandle
SQL = "SELECT * FROM Peaks WHERE P_Nam = '" & Trim$(CreEveCbo.Text) & "' AND Id = " & objMdi.info.ID
ObjTour.RstSQL lPeakHandle, SQL

' Check if iteration is used!
If 1 = CreConchk.Value Then

    'Prompt user for number of interation
    objMdi.info.iInterationLen = ObjTour.DBGetField(PEAKS_LENGTH, lPeakHandle)
    Screen.MousePointer = vbDefault
    UserCancel = False
    schedxfrm.Show vbModal
    
    ' Check if user Cancelled
    If UserCancel Then
        Exit Function
        ObjTour.FreeHandle lPeakHandle
    End If
End If

' Determine if should appear in Daily Form
If CreDayOpt(0).Value Then Act = 1
If CreDayOpt(1).Value Then Act = 2
If CreDayOpt(2).Value Then Act = 3
    
' ------------------------------------------------------------
' First Added new Event to Eve_Tour\Event_Tracker table.
' Retrieve the Event_ID, and add it to each Peak table record.
' ------------------------------------------------------------
SQL = "SELECT * FROM Event_Tracker ORDER BY Event_ID Desc"
ObjTour.RstSQL iSearcherDB, SQL
ObjTour.DBMoveFirst iSearcherDB

iEvent_ID = Val(ObjTour.DBGetField("Event_Id", iSearcherDB)) + 1
' Add Record to EventTracker Table
ObjTour.AddNew iSearcherDB
    ObjTour.DBSetField "ID", objMdi.info.ID, iSearcherDB
    ObjTour.DBSetField "Event_ID", iEvent_ID, iSearcherDB
    ObjTour.DBSetField "Name", Mid(CreatFrm.CreMemTxt.Text, 1, 255), iSearcherDB
ObjTour.Update iSearcherDB
        
' ---------------------------------
' Now Add peak schedule to database
' ----------------------------------
' Peak table is where
' User lPeakHandle handle to Peak table
' and iSearcherdb to

' Add days to iSearcherdb Recordset
SQL = "Select * From Peak Where Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

ProgressBar "Creating Peak Days...", -1, 0, -1
Screen.MousePointer = 11

Sched = ObjTour.DBGetField("P_Sched", lPeakHandle)
SchedLen = ObjTour.DBGetField(PEAKS_LENGTH, lPeakHandle)
CreatePeakDays = SchedLen

' ---------------------
' Create each cycle day
' ---------------------
If objMdi.info.bStartOf Then
    dDate = dEndDate
Else
    dDate = DateAdd("d", -1 * objMdi.info.iInterations * SchedLen, dEndDate)
    dDate = DateAdd("d", 1, dDate)
End If

' Set CDynamicActivities
Set CDynamicActivities = New CDynamics

For iLoop = 0 To objMdi.info.iInterations - 1

' Set up loop variables
iCurrentPos = 1
sDeliminator = ","
sPosition = ""
cActivityNames.Type_ID = gcActive_Type_PeakNames
    
    For i = 1 To SchedLen
        
        ObjTour.AddNew iSearcherDB
            ObjTour.DBSetField "Id", objMdi.info.ID, iSearcherDB
            ObjTour.DBSetField "Date", dDate, iSearcherDB ' Datex
            ' =================================
            ' Set variable for GetPostionValue
            ' =================================
            GetPostionValue Sched, iCurrentPos, sDeliminator, sPosition
            cActivityNames.FindItemByPosition Val(sPosition)
            
            ObjTour.DBSetField "Color", cActivityNames.Colour, iSearcherDB          ' Color
            
            
            ObjTour.DBSetField "Descr", UpdateTag(cActivityNames.Description, sPosition), iSearcherDB          ' Description
            ObjTour.DBSetField "Active", Act, iSearcherDB
            ObjTour.DBSetField "CycleName", CreEveCbo.Text, iSearcherDB
            ObjTour.DBSetField "Page", Calndfrm.CalPagMsk.Text, iSearcherDB
            ObjTour.DBSetField "Event_ID", iEvent_ID, iSearcherDB
            
        ObjTour.Update iSearcherDB
        dDate = DateAdd("d", 1, dDate)
        
    ProgressBar "Creating Peak Days...", -1, (iSched / ObjTour.DBGetField(PEAKS_LENGTH, lPeakHandle)) * 10, -1
    Next i
Next iLoop

ObjTour.FreeHandle lPeakHandle
ProgressBar " ", 0, 0, 0
Screen.MousePointer = 0

Exit Function
Create_Err:
If bDebug Then Handle_Err Err, "CreatePeakDays-CreatFrm"
Resume Next

End Function


Sub Disable_Cre_Controls()
        CreMemTxt.Enabled = False
        CreEveCbo.Enabled = False
        CreConchk.Enabled = False
        chkStartof.Enabled = False
        CreDayOpt(0).Enabled = False
        CreDayOpt(1).Enabled = False
        CreDayOpt(2).Enabled = False
        CreSavCmd.Enabled = False
        CreCloCmd.Caption = "E&xit"
        CreDayOpt(0).Visible = False
        CreDayOpt(1).Visible = False
        CreDayOpt(2).Visible = False
        lblNote.Visible = False
        fraDaytype.Visible = False
End Sub

Sub Load_Active_Peak_Day()
'
' Load_Active_Peak_Day - Purpose is to load Peak days
'                   Description into form
'
' This function is only called if record exist
' and isearcherdb cursor is pointing to it.
' ------------------------------------------------------

On Local Error GoTo Active_Peak_Day

    'CreDeslbl.Caption = "Peak For: " & sRetVal
    CreMemTxt.Text = ObjTour.DBGetField("Name", iSearcherDB)
    CreNotetxt.Visible = True
    CreNotetxt.Text = ObjTour.DBGetField("Descr", iSearcherDB)
    CreEveCbo.Enabled = True
    CreEveCbo.Clear
    CreEveCbo.AddItem ObjTour.DBGetField("CycleName", iSearcherDB)
    CreEveCbo.Enabled = False
    fraDaytype.Visible = False
    If ObjTour.DBGetField("Active", iSearcherDB) = 1 Then CreDayOpt(0) = True
    If ObjTour.DBGetField("Active", iSearcherDB) = 2 Then CreDayOpt(1) = True
    If ObjTour.DBGetField("Active", iSearcherDB) = 3 Then CreDayOpt(2) = True
    
  Disable_Cre_Controls
  
Exit Sub
Active_Peak_Day:
If bDebug Then Handle_Err Err, "Active_Peak_Day"
Resume Next
End Sub

Sub SetUp_Daily(Days_Date As String)
' ---------------------------------
' This set only Lists the occurance
' of a daily activity, is does not
' allow the addition of new records
' or editting or exist records.
' This function is only called when
' record is found, so display.
' ---------------------------------
On Local Error GoTo SetUp_Daily

Dim SearchStr As String

CreMemTxt = "Activity For " & objMdi.info.Name
crelblPeak.Caption = "Daily Information"
Disable_Cre_Controls
lblSchedule.Caption = "Day &Type:"
CreEveCbo.Enabled = True
CreEveCbo.AddItem ObjTour.DBGetField("DayType", iSearcherDB)
CreEveCbo.ListIndex = 0
CreNotetxt.Visible = True
chkStartof.Visible = False
lblNote.Visible = True
' ----------------------------------------
' Get Days memo note and display in form
' ----------------------------------------
SearchStr = "Select memo from " & gcDai_Tour_Dai & " Where date=#" & Days_Date & "# AND Id = " & objMdi.info.ID

ObjTour.RstSQL iSearcherDB, SearchStr

    If 1 = ObjTour.RstRecordCount(iSearcherDB) Then
        CreNotetxt = ObjTour.DBGetField("memo", iSearcherDB)
    End If


        
Exit Sub
SetUp_Daily:
    If bDebug Then Handle_Err Err, "SetUp_Daily"
Resume Next
End Sub

Sub SetUp_Peak()
On Local Error GoTo SetUp_Err
Dim sRetStr As String
            CreDayOpt(0).Visible = -1
            CreDayOpt(1).Visible = -1
            CreDayOpt(2).Visible = -1
            lblNote.Visible = False
            ' Load Cycle Names
                        
            SQL = "SELECT * FROM " & gcPeakTour_Peaks & " WHERE Id = " & objMdi.info.ID
            ObjTour.RstSQL iSearcherDB, SQL
            
            ' ----------------
            ' if Peak schedule
            ' records exist.
            ' ----------------
            If ObjTour.RstRecordCount(iSearcherDB) <> 0 Then
                ObjTour.DBMoveFirst iSearcherDB
                
                Do Until ObjTour.EOF(iSearcherDB) = True
                    sRetStr = ObjTour.DBGetField("P_Nam", iSearcherDB)
                    If IsNull(sRetStr) Or sRetStr = "" Then
                    Else
                    CreEveCbo.AddItem sRetStr
                    End If
                        
                    ObjTour.DBMoveNext iSearcherDB
            
                Loop
                
                 'ObjTour.DBMoveFirst iSearcherDB
                 CreEveCbo.ListIndex = 0
            Else
                        CreEveCbo.Text = "None"
            End If
Exit Sub
SetUp_Err:
    If bDebug Then Handle_Err Err, "SetUP_Peak-CreatFrm"
    Resume Next
End Sub

'Sub UpdateCalender()
'On Local Error GoTo Update_Err
'Dim MnthDy As String
'
'Calndfrm!CalEveLis.AddItem CreMemTxt.Text
'MnthDy = Mid$(CreatFrm.Caption, 1, 3)
'' Select month to update
'Select Case MnthDy
'        Case "Jan":
'            Calndfrm!Jan(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Feb":
'            Calndfrm!Feb(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Mar":
'            Calndfrm!Mar(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Apr":
'            Calndfrm!Apr(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "May":
'            Calndfrm!May(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Jun":
'            Calndfrm!Jun(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Jul":
'            Calndfrm!Jul(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Aug":
'            Calndfrm!Aug(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Sep":
'            Calndfrm!Sep(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Oct":
'            Calndfrm!Oct(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Nov":
'            Calndfrm!Nov(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'        Case "Dec":
'            Calndfrm!Dec(Format$(CreatFrm.Caption, "dd")).BackColor = setEve("Color")
'End Select
'Exit Sub
'Update_Err:
'    If bDebug Then Handle_Err Err, "UpdateCalender-CreatFrm"
'    Resume Next
'
'End Sub

Private Sub CreCloCmd_Click()
Unload CreatFrm
End Sub


Private Sub CreMemTxt_GotFocus()

With CreMemTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub CreSavCmd_Click()

Dim i           As Integer
Dim iTimes      As String 'number of cycles to create
Dim iLenOfCycle As Integer
Dim EndDate     As Date

On Local Error GoTo CreSav_Err

If CreMemTxt.Text = "" Then
    MsgBox "Please fill in Name!", vbOKOnly, LoadResString(gcTourVersion)
    CreMemTxt.SetFocus
    Exit Sub
End If


i = CreEveCbo.ListIndex
iTimes = 1
iLenOfCycle = 0
CreSavCmd.Enabled = 0
EndDate = CreatFrm.Caption
    
Select Case TableType
        Case "Peak":
             objMdi.info.dInterationDate = Format$(EndDate, "mm-dd-yyyy")
             iLenOfCycle = CreatePeakDays(Format$(EndDate, "mm-dd-yyyy"))
                
        Case "Event":
                'CreateEventDay
End Select
' ------------------------------------------
' Update CalndFrm with color and description
' ------------------------------------------
Unload CreatFrm
UserCancel = False
Calndfrm.RestCalendar
Exit Sub
CreSav_Err:
        If bDebug Then Handle_Err Err, "CreSavCmd-CreatFrm"
        Resume Next
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
' ---------------
' Load Type Names
' ---------------
On Local Error GoTo CreatFrm_Load_Err

Dim TypeNum As String, i As Integer, RetStr As String, SQL As String
CentreForm CreatFrm, 0
' ---------------
' Open dynaset to
' events database
' ---------------
CreatFrm.Caption = Format$(objEve.info.EventDate, "mmmm dd, yyyy")
CreHidTxt.Caption = Format$(objEve.info.EventDate, "mm-dd,yyyy")


CreSavCmd.Enabled = -1
UserCancel = True
Me.KeyPreview = True
TableType = Calndfrm!EveFilCbo.Text
' -------------------------------------
' Based on Type take appropriate action
' -------------------------------------
Select Case TableType
    Case "Peak":
            ' ---------------------------------------
            ' one of two steps are taken
            ' If day click has record, then display
            ' else, allow user to create peak
            ' ---------------------------------------
            SQL = "SELECT * FROM " & gcEve_Tour_Peak & " P " & _
                  " INNER JOIN " & gcEve_Tour_Event_Tracker & " E ON (P.ID=E.ID AND P.Event_ID=E.EVENT_ID) " & _
                  " WHERE P.Id = " & objMdi.info.ID & " And P.Page = " & Val(Calndfrm!CalPagMsk.Text) & _
                  " AND P.Date = #" & Format$(objEve.info.EventDate, "mm-dd-yyyy") & "#"
                  
            ObjTour.RstSQL iSearcherDB, SQL
            
            If ObjTour.RstRecordCount(iSearcherDB) > 0 Then
                
                Load_Active_Peak_Day    ' Day Found show description
            Else
            
                SetUp_Peak    'Load Schedule Names into List Box
            End If
            
    Case "Daily":
            ' Define Query
            SQL = "SELECT * FROM " & gcEve_Tour_Event_Daily & _
            " WHERE Id = " & objMdi.info.ID & " AND " & _
            " Date = #" & Format$(objEve.info.EventDate, "mm-dd-yyyy") & "#"
            
            ObjTour.RstSQL iSearcherDB, SQL
            If ObjTour.RstRecordCount(iSearcherDB) <> 0 Then
                SetUp_Daily Format$(objEve.info.EventDate, "mm-dd,yyyy")
            Else
                MsgBox LoadResString(gcNoRecordFound), vbOKOnly, LoadResString(gcTourVersion)
            End If
            
End Select
' -------------------------
' Default combobox to first
' item in list
' -------------------------
If CreEveCbo.ListCount > 0 Then CreEveCbo.ListIndex = 0
                                    ' Update Progress bar
ProgressBar "", 0, 0, 0
Screen.MousePointer = 0             ' Mouse to default
Exit Sub
CreatFrm_Load_Err:
    
    If bDebug Then Handle_Err Err, "CreatFrm_Load-CreatFrm"
    Resume Next
End Sub


Private Sub Form_LostFocus()
CreDayOpt(0).Visible = 0
CreDayOpt(1).Visible = 0
CreDayOpt(2).Visible = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
        CreMemTxt.Enabled = True
        CreEveCbo.Enabled = True
        CreDayOpt(0).Enabled = True
        CreDayOpt(1).Enabled = True
        CreDayOpt(2).Enabled = True
        CreSavCmd.Enabled = True
        CreCloCmd.Caption = "&Cancel"
        CreDayOpt(0).Visible = True
        CreDayOpt(1).Visible = True
        CreDayOpt(2).Visible = True
        CreNotetxt.Visible = False

End Sub



Public Function UpdateTag(ByVal sDesc As String, ByVal sPos As String) As String
On Local Error GoTo Update_Error


Dim iTagEndPos As Integer
Dim iTagBeginPos As Integer
Dim iIndexTag As Integer
Dim iLoop As Integer
Dim CDynam As CDynamic

UpdateTag = sDesc

Set CDynam = CDynamicActivities(sPos)

'Check if Position already exists
If CDynam Is Nothing Then

    'Determine if tag exists
    iTagEndPos = InStr(1, sDesc, "/>", vbTextCompare)
    If 0 = iTagEndPos Then
        CDynamicActivities.Add "", False, 0, 0, sPos
    Else
        'Find begining of tag
        iTagBeginPos = InStr(1, sDesc, "<", vbTextCompare)
        If iTagBeginPos < iTagEndPos Then
            Set CDynam = CDynamicActivities.Add(sDesc, True, iTagBeginPos, iTagEndPos, sPos)
            UpdateTag = CDynam.GetCurrentValue(sDesc)
        Else
            ' Bad format skip
            CDynamicActivities.Add "", False, 0, 0, sPos
        End If
    
    End If
Else
    ' It exist then perform update
    UpdateTag = CDynam.GetNextValue(sDesc)
    
End If

On Error GoTo 0

Exit Function
Update_Error:
Err.Clear

End Function
