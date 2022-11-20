Attribute VB_Name = "TourFunc"
Option Explicit
' ----------------------
' HTML Help Declaration
' ----------------------
Public Const cdlHelpContext = &H1&
Public Const WM_TCARD = &H52&
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HELPFILE = "TourWin.chm"

Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

' ----------------------
' DEBUG File Declaration
' ----------------------
Public Const TOURWINLOGFILE = "TourWin.log"


' This type is used in CreatFrm
Public Type TagReplaceArray
       ActivityPos As Integer
       IndexInActivity As Integer
       LastValue As Long
End Type

Public aTagReplace(1 To 5) As TagReplaceArray

Public Type LevelArray
        Abo As String * 8
        Mid As String * 8
        Bel As String * 8
End Type

Public aLevel(1 To 9) As String * 8

Public UserCancel    As Boolean   ' Used to determine exit of form
Public Datapath      As String    ' Location of Databases
Public ErrorOccurred As Boolean
Public iSearcherDB   As Long      ' Stores Database Handle ID.
Public iDailyDB      As Long
Public lEventHandle  As Long
Public lPeakHandle   As Long
Public bDebug        As Boolean   ' Debug information flag

' Declare Mdi objects
Public objMdiVar     As cMDIVar         'See MdoFrm for visual list of
Public objMdi        As cMDI               'field variables.
Public objDaiVar     As cDaiVar
'Public objDai        As cDai
Public ObjCont       As cCont
Public ObjContVar    As cContVar
Public CallFrom      As String
Public objEve        As cEve
Public objEveVar     As cEveVar
Public ObjChart      As cChart
Public ObjChartVar   As cChartVar
Public ObjTour       As cTourInfo
Public ObjTourVar    As cTourVar
Public cHeartNames   As cHeartNames
Public cActivityNames As cActivity_Names
Public cPeakNames     As cPeak_Names
Public cPeakSchedules As cPeak_Schedules
Public cExportFile    As CExport
Public cLicense       As CRegistration

' ------------------
' FilePointer Handle
' ------------------
Public aHandle(1 To 10) As Long
Public rdoEnv As rdoEnvironment
' Rdo Definitions
Public rdoDb1 As rdoConnection
Public rdoDb2 As rdoConnection
Public rdoDb3 As rdoConnection
Public rdoDb4 As rdoConnection
Public rdoDb5 As rdoConnection
Public rdoDb6 As rdoConnection
Public rdoDb7 As rdoConnection
Public rdoDb8 As rdoConnection
Public rdoDb9 As rdoConnection
' ResultSets
Public rdoRs1 As rdoResultset
Public rdoRs2 As rdoResultset
Public rdoRs3 As rdoResultset
Public rdoRs4 As rdoResultset
Public rdoRs5 As rdoResultset
Public rdoRs6 As rdoResultset
Public rdoRs7 As rdoResultset
Public rdoRs8 As rdoResultset
Public rdoRs9 As rdoResultset

' ------------------
' Database variables
' ------------------

Public cTour_DB     As CDatabase
Public dbsAccess    As CDatabase
Public dbs1         As Database
Public dbs2         As Database
Public dbs3         As Database
Public dbs4         As Database
Public dbs5 As Database
Public dbs6 As Database
Public dbs7 As Database
Public dbs8 As Database
Public dbs9 As Database
Public dbs10 As Database

Public rstTour As Recordset
Public rst1 As Recordset
Public rst2 As Recordset
Public rst3 As Recordset
Public rst4 As Recordset
Public rst5 As Recordset
Public rst6 As Recordset
Public rst7 As Recordset
Public rst8 As Recordset
Public rst9 As Recordset
Public rst10 As Recordset

' ----------------------
' Public Const define
' for single mdiFrm menu
' ----------------------
Public Const Unloadmnu = "Unloadmnu"
Public Const Loadmnu = "Loadmnu"
' GuideFrm menu values
Public Const GuideFrm_Exit = "E&xit Peak Form"
Public Const GuideFrm_Option = "&Option Peak setup"
Public Const gcGuideFrm_Newmnu = "&New Peak"
Public Const gcGuideFrm_Savmnu = "&Save Peaks"
Public Const gcGuideFrm_Delmnu = "&Delete Peak"

' Mdi Form
Public Const MdiFrm_Exit = "E&xit TourWin..."
Public Const MdiFrm_Option = "&Option, StartUp..."
Public Const MdiFrm_Newmnu = "&New User"
Public Const MdiFrm_Savmnu = "&Save User"
Public Const MdiFrm_Delmnu = "&Delete User"
' Calendar Form
Public Const CalndFrm_Exit = "E&xit Calendar.."
Public Const CalndFrm_Option = "&Option, Calendar..."
Public Const CalndFrm_Newmnu = "&New Event"
Public Const CalndFrm_Savmnu = "&Save Event"
Public Const CalndFrm_Delmnu = "&Delete Event"
Public Const CalndFrm_Begmnu = "&Begin Calendar"
Public Const gcCalendarWindow = "CalendarWindow"
' ContFrm Form menus
Public Const gcContFrm_Option = "&Option, Contact..."
Public Const gcContFrm_Exit = "E&xit Contact"
Public Const gcContFrm_Newmnu = "&New Contact"
Public Const gcContFrm_Savmnu = "&Save Contact"
Public Const gcContFrm_Delmnu = "&Delete Contact"
' EventFrm Form menus...
Public Const gcEvent_Exit = "E&xit Event Form"
Public Const gcEvent_Option = "&Option, Event Form..."
Public Const gcEvent_Newmnu = "&New Event"
Public Const gcEvent_Savmnu = "&Save Events"
Public Const gcEvent_Delmnu = "&Delete Event"

' MdiOpt Form menus...
Public Const gcMdO_Exit = "E&xit Options"
Public Const gcMdO_Option = "Start &Options..."
Public Const gcMdO_Newmnu = "&New"
Public Const gcMdO_Savmnu = "&Save Options"
Public Const gcMdO_Delmnu = "&Delete Options"

' Graph Form
Public Const GraphFrm_Exit = "E&xit Graph..."
Public Const GraphFrm_Option = "&Option, Graph..."
Public Const GraphFrm_Newmnu = "&New Graph"
Public Const GraphFrm_Savmnu = "&Save Graph"
Public Const GraphFrm_Delmnu = "&Delete Graph"
Public Const GraphFrm_Begmnu = "&Begin Graph"
' HrtZone Form
Public Const HrtZone_Exit = "E&xit Heart Zone..."
Public Const HrtZone_Option = "&Option, Heart Zones..."
Public Const gcHrtZone_Newmnu = "&New Caption"
Public Const gcHrtZone_Savmnu = "&Save Heart Rates"
Public Const gcHrtZone_Delmnu = "&Delete Caption"

' Daily Form
Public Const DailyFrm_Option = "&Options, Daily..."
Public Const DailyFrm_Exit = "E&xit Daily Diary..."
Public Const DailyFrm_Newmnu = "&New Record"
Public Const DailyFrm_Savmnu = "&Save Record"
Public Const DailyFrm_Delmnu = "&Delete Record"

' P_Set Form
Public Const P_SetUp_Option = "&Options, Peak Setup..."
Public Const P_SetUp_Exit = "E&xit Peak Setup"
Public Const P_SetUp_Newmnu = "&New Schedule"
Public Const P_SetUp_Savmnu = "&Save Schedule"
Public Const P_SetUp_Delmnu = "&Delete Schedule"

' Conconi Form
Public Const gcCONCONI_EXIT = "E&xit Conconi"
Public Const gcCONCONI_OPTION = "&Option, Conconi..."
Public Const gcCONCONI_NEWMNU = "&New Conconi Test"
Public Const gcCONCONI_SAVMNU = "&Save Conconi Test"
Public Const gcCONCONI_DELMNU = "&Delete Conconi Test"

' DataBase Objects
'Public dbsEve_Tour As Database ' Global database object
'
' Chart Form values

Public Const gcDailyChart = "Daily"
Public Const gcEventChart = "Event"
Public Const gcPeakChart = "Peak"
' ------------------
' TourWin Constants
' -----------------
' This values are pointers to Tourwin.rc
Public Const gcTourVersion = 1  ' This only says TourWin...
Public Const gcRegTourKey = 2
Public Const gcRegTourExe = 3

Public Const gcTop_Form = 4
Public Const gcLeft_Form = 5
Public Const gcHeight_Form = 6
Public Const gcWidth_Form = 7
Public Const gcWindowState = 8
Public Const gcFormsRegSubKey = 9
Public Const gcTourWinHelp = 10
Public Const gcTourDBase = 11
Public Const gcTourUserName = 12
Public Const gcTourWelcome = 13
Public Const gcTourLastDB = 15
Public Const gcTourWinVersionNo = 16
Public Const gcLeaMarTech = 17
' ------------------
' Database Structure
' ------------------
Type databaseStructure
    Tablecount As Integer
    Tables(0 To 10) As String
    FieldCount(0 To 10) As String
    Fields(1 To 20) As String
End Type
'
'
Public TypeEve_Tour As databaseStructure
Public TypeNameTour As databaseStructure
Public PeakTour_Tour As databaseStructure

' --------------
' Database names
' --------------
Public bSQLDatabase As Boolean
Public Const gcTOURWIN_PASSWORD = "Tourwin"
Public Const gcTour_Win = "TourWin.mdb"

' ---------------------
' Database Table Names
' ---------------------
Public Const gcID = "ID"    ' Generic const for all tables ID field...

'Activity_Names Table
Public Const gcActivitiesTable = "Activity_Name"
Public Const gcActive_ID = "ID"
Public Const gcActive_Type = "Type_ID"
Public Const gcActive_Pos = "Pos"
Public Const gcActive_Des = "Description"
Public Const gcActive_Colour = "Colour"
' Constants for Type_ID
Public Const gcActive_Type_PeakNames = 0
Public Const gcActive_Type_EventNames = 1
Public Const gcActive_Type_HeartNames = 2

'Eve_tour.mdb
Public Const gcEve_Tour_Peak = "Peak"
Public Const gcEve_Tour_Event_Tracker = "Event_Tracker"
Public Const gcEve_Tour_Event_Daily = "Daily"
Public Const gcEve_Tour_Event = "Event"
Public Const gcPeak_Table = "Peak"
Public Const gcPeaks_Table = "Peaks"

' ContTour.mdb
Public Const gcContTour_Contacts = "Contacts"

' NameTour.mdb
Public Const gcNameTour_Events = "Events"
Public Const gcNameTour_HeartNames = "HeartNames"
Public Const gcNameTour_PeakNames = "PeakNames"

' Dai_Tour.mdb
Public Const gcDai_Tour_Dai = "Dai"
Public Const gcUserTour_ContactOpt = "ContactOpt"
Public Const gcUserTour_DailyOpt = "DailyOpt"

'PeakTour.mdb
Public Const gcPeakTour_Peaks = "Peaks"

' Data Table for Conconi data
Public Const gcData = "Data"
Public Const gcData_Date = "Date"
Public Const gcData_Minute = "iMin"
Public Const gcData_PulseRate = "iPr"
Public Const gcData_Watt = "iWatt"

'lEVEL
Public Const gcLEVELS_TABLE = "Levels"
' -----------
' Field Names
' -----------
'Global Field Names
Public Const gcDate = "Date"
Public Const gcColor = "Color"


Public Const gcEve_Tour_Peak_Fields = ",Date,Id,Descr,Active,Color,CycleName,Page,Event_ID,"

' --------------
' ContTour Table
' --------------

Public Const gcContTour_Contacts_Contact = "Contact"
Public Const gcContTour_Contacts_Last = "Last"
Public Const gcContTour_Contacts_First = "First"
Public Const gcContTour_Contacts_Phone = "Phone"
Public Const gcContTour_Contacts_Fax = "Fax"
Public Const gcContTour_Contacts_Email = "Email"
Public Const gcContact_Count = "Count"

' ----------------
' ContactOpt Table
' ----------------
Public Const gcUserTour_ContactOpt_Id = "Id"
Public Const gcUserTour_ContactOpt_IndexOrder = "IndexOrder"
Public Const gcUserTour_ContactOpt_SortOrder = "SortOrder"
Public Const gcUserTour_ContactOpt_ContactWidth = "ContactWidth"
Public Const gcUserTour_ContactOpt_LastWidth = "LastWidth"
Public Const gcUserTour_ContactOpt_FirstWidth = "FirstWidth"
Public Const gcUserTour_ContactOpt_PhoneWidth = "PhoneWidth"
Public Const gcUserTour_ContactOpt_FaxWidth = "FaxWidth"
Public Const gcUserTour_ContactOpt_E_MailWidth = "E-MailWidth"

' -------------------------
' Peak Table (Eve_Tour.mdb)
' -------------------------
Public Const gcPeakTour_Descr = "Descr"
Public Const gcPeakTour_CycleName = "CycleName"

' ---------------
' Dai_Tour Fields
' ---------------
Public Const gcDai_Tour_Date = "Date"
Public Const gcDai_Tour_DaiType = "Type"    'Long Data Type
Public Const gcDai_Tour_Id = "Id"
Public Const gcDai_Tour_DayType = "DayType"
Public Const gcDai_Tour_Heart = "Heart"
Public Const gcDai_Tour_DayRate = "DayRate"
Public Const gcDai_Tour_Oddometer = "Oddometer"
Public Const gcDai_Tour_Weight = "Weight"
Public Const gcDai_Tour_DaiMile = "DaiMile"
Public Const gcDai_Tour_HeaV1 = "HeaV1"
Public Const gcDai_Tour_HeaV2 = "HeaV2"
Public Const gcDai_Tour_HeaV3 = "HeaV3"
Public Const gcDai_Tour_HeaV4 = "HeaV4"
Public Const gcDai_Tour_HeaV5 = "HeaV5"
Public Const gcDai_Tour_HeaV6 = "HeaV6"
Public Const gcDai_Tour_HeaV7 = "HeaV7"
Public Const gcDai_Tour_HeaV8 = "HeaV8"
Public Const gcDai_Tour_HeaV9 = "HeaV9"
Public Const gcDai_Tour_Peak_Day = "Peak_Day"
Public Const gcDai_Tour_Peak_Wee = "Peak_Wee"
Public Const gcDai_Tour_Peak_Mon = "Peak_Mon"
Public Const gcDai_Tour_Memo = "Memo"
Public Const gcDai_Tour_Sleep = "Sleep"
Public Const gcDai_Tour_Level1 = "Level1"
Public Const gcDai_Tour_Level2 = "Level2"
Public Const gcDai_Tour_Level3 = "Level3"
Public Const gcDai_Tour_Level4 = "Level4"
Public Const gcDai_Tour_Level5 = "Level5"
Public Const gcDai_Tour_TotalHrs = "TotalHrs"
Public Const gcDai_Tour_WeeklySunday = "WeeklySunday"
' --------------------------------------------------------------
' Two types of Dai records - Normal and Historical Significates
' --------------------------------------------------------------
Public Const gcDAI_NORMAL = 1
Public Const gcDAI_HISTORICAL = 2
Public Const gcDAI_CONCONI = 3

' ----------
' Daily Table
' ----------
Public Const gcDaily_Table = "Daily"
' ----------
' Event Tour
' ----------
Public Const gcEve_Tour_Evememo = "Evememo"
Public Const gcEve_Tour_EveType = "EveType"
Public Const gcEve_Tour_PointToCont = "PointToCont"
Public Const gcEve_Tour_Contact = "Contact"
Public Const gcEve_Tour_Id = "Id"
Public Const gcEve_Tour_Date = "Date"
Public Const gcEve_Tour_Color = "Color"
Public Const gcEve_Tour_Page = "Page"

' ---------
' Peaks
' ---------
Public Const PEAKS_LENGTH = "Cycle_Length"
Public Const gcPEAK_NAME = "P_Nam"
Public Const gcPEAK_SCHED = "P_Sched"
Public Const gcPEAK_DATE = "P_Date"
Public Const gcALLTYPES = "All Types"
Public Const gcPEAK_ID = "Id"

' ---------
' UserTour
' ---------
Public Const gcUserTour_UserTbl = "UserTbl"
Public Const gcUserTour_UserTbl_Name = "Name"
Public Const gcUserTour_UserTbl_PassWord = "PassWord"
Public Const gcUserTour_UserTbl_DataPath = "DataPath"
Public Const gcUserTour_UserTbl_Security = "Security"
Public Const gcUserTour_UserTbl_Load = "Load"         ' dbByte
Public Const gcUserTour_UserTbl_Metafile = "Metafile" ' dbText, 100
Public Const gcUserTour_UserTbl_ShowMeta = "ShowMeta" ' dbBoolean
Public Const gcUserTour_UserTbl_BitField = "BitField" ' dbInteger
Public Const gcUserTour_UserTbl_DailyOptions = "DailyOpt" 'dbLong

' --------------
' General Prompts
' --------------
Public Const gcWantToDelete = 832
Public Const gcNoRecordFound = 52
Public Const gcInvalidNumeric = 53
Public Const gcInvalidDate = 54
Public Const gcInvalidFormat = 55
Public Const gcNumericToLarge = 56
Public Const gcPrompt_Delete = 57
Public Const SELECT_ITEM_FROM_LIST = 913
' --------------
' Form constants
' --------------
Public Const gcMDICaption = 100
Public Const gcTourCommentEmail = 101
' --------------------
' Event_Calendar Const
' --------------------
Public Const gcEventCalendarTitle = 5200
Public Const gcEventCalendarSubTitle = 5201
Public Const gcMINPAGEVALUE = 1
Public Const gcMAXPAGEVALUE = 99
Public Const gcMINYEAR = 1931
Public Const gcMAXYEAR = 2030
' --------------
' ContFrm Const
' --------------
Public Const gcSaveColumnChanges = 825

'
' Passfrm
' -------
Public Const gcPassFrmUserName = 200

' -------
' DateFrm
' -------
Public Const gcDateFrmStart = 300
Public Const gcDateFrmEnd = 301
' --------------
' Error Messages
' --------------
Public Const gcFailedEditTryAgain = 1000
Public Const gcOutOfDBHandles = 1001
Public Const gcNoRecords = 1002
Public Const gcFieldNotExist = 1003

' ------------------------------
' Constants for Export Options
' ------------------------------
Public Const gcFILE_EXTENTION = "tor"
' Schedule Table and Fields
Public Const gcEXPORT_SCHEDULE_TABLE = "Schedules"
Public Const gcEXPORT_SCHEDULES_FIELD = "Schedule"
Public Const gcEXPORT_TYPE_ID_FIELD = "Type_ID"
Public Const gcEXPORT_NAME_FIELD = "Name"
Public Const gcEXPORT_CYCLE_L_FIELD = "Cycle_Length"
' Activity Table and Fields
Public Const gcEXPORT_ACTIVITY_TABLE = "Activity"
Public Const gcEXPORT_POS_FIELD = "Position"
Public Const gcEXPORT_DESCRIPTION_FIELD = "Description"
Public Const gcEXPORT_COLOUR_FIELD = "Colour"
Public Const gcEXPORT_DATE_FIELD = "Date"

Public Const gcEXPORT_PASSWORD = "export"
Public Const gcEVENT_DAYS = 0
Public Const gcPEAK_SCHEDULE = 1
Public Const gcDAILY_ACTIVITIES = 2
Public Const EXPORT_HANDLE = 10


' ---------------------
' Progress Bar Messages
' ---------------------
Public Const gcLoadingContactDB = 1500
Public Const gcLoadingDailyDB = 1501
Public Const gcNewContactRec = 1502
Public Const gcLoadingData = 1503
Public Const gcSavingdata = 1504
Public sBuffer As String  ' Global variable used for all purposes...

Public Const ICONBARWIDTH = 720
Public Const MENUANDTOOLBARHEIGTH = 840
Public Const LB_SETHORIZONTALEXTENT = &H194 'List box constant

Public gCrystalReport As CRPEAuto.Application
Public CryReport As CRPEAuto.Report
Public CryFormula As CRPEAuto.FormulaFieldDefinition
Public CryFormulas As CRPEAuto.FormulaFieldDefinitions

Public gActivityForm As Object
Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
(ByVal hWndParent As Long, ByVal fRequest As Long, ByVal lpszDriver _
As String, ByVal lpszAttributes As String) As Long

' Constants and Declarations for Registration
Public Const gcTOURSYSTEM = "TourSystem"    ' Name field of user record
Public Const gcTRIALLENGTH = 60             ' Number of allowable opens

Public Enum eCheckDatabase
    CreateDBaseAfterInstall
    CheckIfExist_LastUserFound
End Enum

Sub Daily_Calendar_Report()

Dim StartD       As String
Dim EndD         As String
Dim sDataSource  As String
Dim sReportRange As String

On Local Error GoTo Daily_Rep_Err

sDataSource = Datapath & gcTour_Win
DateFrm.Show 1

If UserCancel = True Then Exit Sub

' Make sure DSN is pointing to current database
ObjTour.RegisterUpdateDSN objMdi.info.Datapath & gcTour_Win

StartD = Format$(DateFrm.DatFroTxt, "YYYY, MM, DD")
EndD = Format$(DateFrm.DatToTxt, "YYYY, MM, DD")
sReportRange = Format$(StartD, "mmmm dd, yyyy") & " to " & Format$(EndD, "mmmm dd, yyyy")
          
If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing
Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Calnd_Da.rpt")

' Loop Through each formula and update accordingly
Set CryFormulas = CryReport.FormulaFields

For Each CryFormula In CryFormulas
    Select Case CryFormula.Name
        Case "{@DateRange}":
            CryFormula.Text = Chr$(34) & sReportRange & Chr$(34)
    End Select
Next

CryReport.RecordSelectionFormula = "{Daily.Date} in Date (" & StartD & ") to Date (" & EndD & ") AND {Daily.Id} = " & objMdi.info.ID

CryReport.Preview ("Daily Calender Activity Report. " & LoadResString(gcTourVersion))
    
Exit Sub
Daily_Rep_Err:
    Select Case Err
        Case 13:       'File not found
                    MsgBox Error$(Err)
        Case 20533:
                    MsgBox Error$(Err)
        Case 20515:
                    MsgBox Error$(Err)
        Case Else
        If bDebug Then Handle_Err Err, "Daily_Calendar_Report-TourFunc"
        Resume Next
    End Select
End Sub

Sub Peak_Calendar_Report()
' ------------------------------------------
' The procedure calls "peakschd.rpt" after
' first allowing the user to define the
' parameters for the report.
' ------------------------------------------
Dim sDataSource As String

On Local Error GoTo Daily_Rep_Err

sDataSource = Datapath & gcTour_Win
' ---------------------------
' Diplay dialog to retrieve
' peak schedules to report by
' use sBuffer to pass select values
' ---------------------------

PeakOptRpt.Show vbModal

If UserCancel = True Then Exit Sub

' Make sure DSN is pointing to current database
ObjTour.RegisterUpdateDSN objMdi.info.Datapath & gcTour_Win

If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing

' Create Crystal Objects
Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\peakschd.rpt")

CryReport.RecordSelectionFormula = sBuffer

CryReport.Preview "Peak Schedule Report. " & LoadResString(gcTourVersion)

Exit Sub
Daily_Rep_Err:
    Select Case Err
        Case 13:       'File not found
                    MsgBox Error$(Err)
        Case 20533:
                    MsgBox Error$(Err)
        Case 20515:
                    MsgBox Error$(Err)
        Case Else
                MsgBox Error$(Err)
        If bDebug Then Handle_Err Err, "Daily_Calendar_Report-TourFunc"

        Resume Next
    End Select

End Sub

Sub Event_Calendar_Report()

Dim StartD          As String
Dim EndD            As String
Dim sReportRange    As String

On Local Error GoTo Event_Rep_Err

' Prompt user for date range
DateFrm.Show 1
If UserCancel = True Then Exit Sub

StartD = Format$(DateFrm.DatFroTxt, "YYYY, MM, DD")
EndD = Format$(DateFrm.DatToTxt, "YYYY, MM, DD")
sReportRange = Format$(StartD, "mmmm dd, yyyy") & " to " & Format$(EndD, "mmmm dd, yyyy")

' Create Crystal Objects
Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Calnd_Ev.rpt")

CryReport.RecordSelectionFormula = "{Event.Date} in Date (" & StartD & ") to Date (" & EndD & ") AND {Event.Id} = " & objMdi.info.ID

' Loop Through each formula and update accordingly
Set CryFormulas = CryReport.FormulaFields

For Each CryFormula In CryFormulas
    Select Case CryFormula.Name
        Case "{@rpt_date_range}":
            CryFormula.Text = Chr$(34) & sReportRange & Chr$(34)
        Case "{rpt_title}"
            CryFormula.Text = Chr$(34) & LoadResString(gcEventCalendarTitle) & Chr$(34)
        Case "{rpt_sub_title}"
            CryFormula.Text = Chr$(34) & LoadResString(gcEventCalendarSubTitle) & Chr$(34)
    End Select
Next


CryReport.Preview LoadResString(gcEventCalendarTitle) & " " & LoadResString(gcTourVersion)

'        Size_Crystal_Window
Exit Sub

Event_Rep_Err:
    MsgBox Error$(Err)
    Select Case Err
        Case 13:       'File not found
            MsgBox Error$(Err)
        Case 20533:
            MsgBox Error$(Err)
        Case 20515:
            MsgBox Error$(Err)
        Case Else
        If bDebug Then Handle_Err Err, "Event_Calendar_Report-TourFunc"
        Resume Next
    End Select

End Sub


Function Get_NameTour_HeartNames(FieldType As String) As String
' -----------------------------------------
' Purpose:  Is a central and Only location
'           which accesses the HeartNames Table
'           found in the NameTour database.
' -----------------------------------------
On Local Error GoTo Get_NameTour_HeartNames_Err

Dim SQL As String, RetStr As String
Dim lEventHandle As Long

If bDebug Then Handle_Err 0, "Get_NameTour_HeartNames-TourFunc"

' Check if handle was obtained.
lEventHandle = ObjTour.GetHandle
If lEventHandle = 0 Then
        Get_NameTour_HeartNames = ""
        Exit Function
End If

' Open DB Connection
'ObjTour.DBOpen gcNameTour, gcNameTour_HeartNames, lEventHandle

' Set RecordSet
'SQL = "SELECT * FROM " & gcNameTour_HeartNames & " WHERE Id = " & objMdi.info.ID
SQL = "SELECT * FROM " & gcActivitiesTable & " WHERE Id = " & objMdi.info.ID & " AND " & gcActive_Type & " = " & gcActive_Type_HeartNames & " AND " & gcActive_Pos & " = " & Mid$(FieldType, 6, 1)
ObjTour.RstSQL lEventHandle, SQL

' ------------------------
' set Default return value
' ------------------------
Get_NameTour_HeartNames = "No Return"

If ObjTour.RstRecordCount(lEventHandle) <> 0 Then
    ' --------------------------
    ' Set TypeNum = Event# field
    ' --------------------------
    RetStr = IIf("" = ObjTour.DBGetField(gcActive_Des, lEventHandle), "No Return", ObjTour.DBGetField(gcActive_Des, lEventHandle))
        If RetStr <> "No Return" And Trim$(RetStr) <> "" Then
                Get_NameTour_HeartNames = Trim(RetStr)
        End If
End If   'Ends if RecordCount = 0

' Release and close db
ObjTour.FreeHandle lEventHandle

Exit Function
Get_NameTour_HeartNames_Err:
    If bDebug Then Handle_Err Err, "Get_NameTour_HeartNames-TourFunc"
    Err.Clear
    Exit Function
End Function

Sub UpdateCalender(DDMMYYYY As Date, Color As Long)
On Local Error GoTo Update_Err
Dim UpdateMnth As String, UpdateDay As String
UpdateMnth = Mid$(Format$(DDMMYYYY, "MMM"), 1, 3)
' Select month to update
' ----------------------
Select Case UpdateMnth
        Case "Jan":
            Calndfrm!Jan(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Feb":
            Calndfrm!Feb(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Mar":
            Calndfrm!Mar(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Apr":
            Calndfrm!Apr(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "May":
            Calndfrm!May(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Jun":
            Calndfrm!Jun(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Jul":
            Calndfrm!Jul(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Aug":
            Calndfrm!Aug(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Sep":
            Calndfrm!Sep(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Oct":
            Calndfrm!Oct(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Nov":
            Calndfrm!Nov(Format$(DDMMYYYY, "dd")).BackColor = Color
        Case "Dec":
            Calndfrm!Dec(Format$(DDMMYYYY, "dd")).BackColor = Color
End Select
Exit Sub
Update_Err:
    If bDebug Then Handle_Err Err, "UpdateCalender-CreatFrm"
    Resume Next
End Sub
Sub Contact_Report()
On Local Error GoTo Contact_Rep_Err
Dim Data_Source As String

If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing

' Make sure DSN is pointing to current database
ObjTour.RegisterUpdateDSN objMdi.info.Datapath & gcTour_Win

Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Contact.rpt")
CryReport.RecordSelectionFormula = "{Contacts.Id} = " & objMdi.info.ID

CryReport.Preview "Contact Report"
    
Exit Sub
Contact_Rep_Err:
    
    If bDebug Then
        Handle_Err Err, "Contact_Rep_Err-TourFunc"
    End If
    Resume Next
End Sub

' ===========================================================================
' CreateTourDatabase - Purpose of function is to create an Access
'       database for this application
'
'   Database name will be Tourwin.mdb with the following tables
'       UserObject -
'               ContactOpt
'               DailyOpt
'               UserTbl
'       DailyObject:
'               Dai
'       EventObject:
'               Daily
'               Event
'               Event_Tracker
'               Peak
'       NameTourObject:
'               Events
'               HeartNames
'               Levels
'               PeakNames
'       PeakTourObject
'               Peaks
'       ContTourObject:
'               Contacts
'       ConcTourObject:
'               Data
'               Tables
'==============================
' Revisions:
'   April 17th, 2000
'       Changed from creating single independent database to a single
'       database with all table held within.
' ===========================================================================

Function CreateTourDatabases(CreatePath As String, dbName As String) As Boolean
' -----------------------------------------------------------------
' Creates specificed database in CreatePath directory. The
' intended use for this function is for new installs and new users
' -----------------------------------------------------------------
On Local Error GoTo CreateTour_Err
Dim bTemp As Boolean

' -------------------------------
' Assume pestimitic that database
' was not created sucessfully!
' -------------------------------

CreateTourDatabases = False

If "" = Dir$(CreatePath) Then
    MakeDirectory CreatePath
End If
    ' --------------------------------
    ' Start By Creating Tourwin.mdb
    ' and adding User table
    ' --------------------------------

    ' Define UserTbl Table
    If Define_UserTour_UserTbl_Plus_Fields(CreatePath) Then
            CreateTourDatabases = True
    Else
            MsgBox "Unable to Create/Overwrite UserTour-UserTbl Database", vbOKOnly, LoadResString(gcTourVersion)
    End If
    ' Define DailyOpt Table
'    If Define_UserTour_DailyOpt_Plus_Fields(CreatePath) Then
'            CreateTourDatabases = True
'    Else
'             MsgBox "Unable to Create/Overwrite UserTour-DailyOpt Database", vbOKOnly, LoadResString(gcTourVersion)
'    End If
    ' Define ContactOpt Table
    If Define_UserTour_ContactOpt_Plus_Fields(CreatePath) Then
            CreateTourDatabases = True
    Else
            MsgBox "Unable to Create/Overwrite UserTour-ContactOpt Database", vbOKOnly, LoadResString(gcTourVersion)
    End If

    If Define_ConcTour_Data_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite ConcTour-Data Database", vbOKOnly, LoadResString(gcTourVersion)
    End If
        
'     If Define_ConcTour_Tables_Plus_Fields(CreatePath) Then
'        CreateTourDatabases = True
'    Else
'        MsgBox "Unable to Create/Overwrite ConcTour-Tables Database", vbOKOnly, LoadResString(gcTourVersion)
'    End If
    
    ' ---------------------
    ' Create Contact Table
    ' ---------------------
    If Define_ContTour_Contacts_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite PeakTour-Peak Database", vbOKOnly, LoadResString(gcTourVersion)
    End If
        
        
    ' ---------------------
    ' Create Dai Table
    ' ---------------------
    If Define_Dai_Tour_Dai_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite DaiTour.mdb Database", vbOKOnly, LoadResString(gcTourVersion)
    End If
    
    ' ---------------------
    ' Create Daily Table
    ' ---------------------
    If Define_Eve_Tour_Daily_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite Eve_Tour-Daily Database", vbOKOnly, LoadResString(gcTourVersion)
    End If
    
    ' ---------------------
    ' Create Event Table
    ' ---------------------
    If Define_Eve_Tour_Event_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite Eve_Tour-Daily Database", vbOKOnly, LoadResString(gcTourVersion)
    End If
    
    If Define_Eve_Tour_Peak_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite Eve_Tour-Peak Database", vbOKOnly, LoadResString(gcTourVersion)
    End If
    ' --------------------------------------
    ' Create Event Tracker Table with fields
    ' --------------------------------------
    If Define_Eve_Tour_Event_Tracker_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite Event_Tracker-Peak Database", vbOKOnly, LoadResString(gcTourVersion)
    End If
        

    ' ---------------------
    ' Create Daily Table
    ' ---------------------
'    If Define_NameTour_HeartNames_Plus_Fields(CreatePath) Then
'        CreateTourDatabases = True
'    Else
'        MsgBox "Unable to Create/Overwrite NameTour-HeartNames Database", vbOKOnly, LoadResString(gcTourVersion)
'    End If
    If Define_NameTour_Levels_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite NameTour-Levels Database", vbOKOnly, LoadResString(gcTourVersion)
    End If

    ' Create Activities Table
    If Define_Activities_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite Activity Table", vbOKOnly, LoadResString(gcTourVersion)
    End If
    
'    If Define_NameTour_Events_Plus_Fields(CreatePath) Then
'        CreateTourDatabases = True
'    Else
'        MsgBox "Unable to Create/Overwrite NameTour-Events Database", vbOKOnly, LoadResString(gcTourVersion)
'    End If
        
    
    If Define_PeakTour_Peak_Plus_Fields(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite PeakTour-Peak Database", vbOKOnly, LoadResString(gcTourVersion)
    End If

    ' -------------------------------------------
    ' Create DaiPercentage Query Def for reports
    ' -------------------------------------------
    If Define_DaiPercentage_QueryDef(CreatePath) Then
        CreateTourDatabases = True
    Else
        MsgBox "Unable to Create/Overwrite DaiTour.mdb Database", vbOKOnly, LoadResString(gcTourVersion)
    End If

On Local Error GoTo 0
Exit Function
CreateTour_Err:
If bDebug Then Handle_Err Err, "CreateTourDatabases-TourFunc"
Resume Next
End Function

Function Define_ConcTour_Data_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Conc_Data_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index

Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating ConcTour Database ", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("Data")

' Append cDate Field
Set TourField = TourTableDef.CreateField("Date", dbDate)
    TourTableDef.Fields.Append TourField
' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append iMin Field
Set TourField = TourTableDef.CreateField("iMin", dbInteger)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
ProgressBar "Creating ConcTour Database ", -1, 5, -1
' Append iWatt Field
Set TourField = TourTableDef.CreateField(gcData_Watt, dbInteger)
    TourTableDef.Fields.Append TourField
' Append iPR
Set TourField = TourTableDef.CreateField("iPr", dbInteger)
    TourTableDef.Fields.Append TourField
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
ProgressBar "Creating ConcTour Database...", -1, 7, -1

' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_ConcTour_Data_Plus_Fields = True
ProgressBar "Creating ConcTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0
Exit Function
Conc_Data_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_ConcTour_Data_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_ConcTour_Data_Plus_Fields-TourFunc"
    Resume Next
End Function


Function Define_ConcTour_Tables_Plus_Fields(CreatePath As String) As Boolean
On Local Error GoTo Conc_Tables_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index

' --------------------------------------------------
' This function is replace with Dai record Type = 3
' --------------------------------------------------
Exit Function

'Set DefaultWorkspace = DBEngine.Workspaces(0)
'' Open, ConcTour database.
'Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
'' Fill in new database.
'ProgressBar "Creating ConcTour Database...", -1, 2, -1
'' Create new TableDef.
'Set TourTableDef = TourDatabase.CreateTableDef("Tables")
'
'' Append cDate Field
'Set TourField = TourTableDef.CreateField("Date", dbText, 10)
'    TourTableDef.Fields.Append TourField
'' Append Id Field
'Set TourField = TourTableDef.CreateField("Id", dbLong)
'    TourField.Required = True
'    TourTableDef.Fields.Append TourField
'' Append iPedal Field
'Set TourField = TourTableDef.CreateField("iPedal", dbInteger)
'    TourField.Required = True
'    TourTableDef.Fields.Append TourField
'' Append iDuration Field
'Set TourField = TourTableDef.CreateField("iDuration", dbInteger)
'    TourTableDef.Fields.Append TourField
'' Append iMax_P
'Set TourField = TourTableDef.CreateField("iMax_P", dbInteger)
'    TourTableDef.Fields.Append TourField
'ProgressBar "Creating ConcTour Database...", -1, 5, -1
'' Append iMax_W
'Set TourField = TourTableDef.CreateField("iMax_W", dbInteger)
'    TourTableDef.Fields.Append TourField
'' Append iDeflex
'Set TourField = TourTableDef.CreateField("iDeflex", dbInteger)
'    TourTableDef.Fields.Append TourField
'' -----------------------------------------------------------------
'' Save TableDef definition by appending it to TableDefs collection.
'' -----------------------------------------------------------------
'TourDatabase.TableDefs.Append TourTableDef
'ProgressBar "Creating ConcTour Database...", -1, 7, -1
'' ---------------------------------
'' Close newly created TourDatabase.
'' ---------------------------------
'TourDatabase.Close
'Define_ConcTour_Tables_Plus_Fields = True
'ProgressBar "Creating ConcTour Database...", -1, 10, -1
'ProgressBar "", 0, 0, 0
Exit Function
Conc_Tables_Err:
    If bDebug Then Handle_Err Err, "Define_ConcTour_Tables_Plus_Fields-TourFunc"
    Resume Next
End Function

Function Define_ContTour_Contacts_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Cont_Contacts_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating ConcTour Database...", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("Contacts")

' Append Id Field
Set TourField = TourTableDef.CreateField(gcID, dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Contact Field
Set TourField = TourTableDef.CreateField("Contact", dbText, 50)
    TourTableDef.Fields.Append TourField
' Append Last Field
Set TourField = TourTableDef.CreateField("Last", dbText, 35)
    TourTableDef.Fields.Append TourField
' Append First Field
Set TourField = TourTableDef.CreateField("First", dbText, 25)
    TourTableDef.Fields.Append TourField
ProgressBar "Creating ConcTour Database...", -1, 5, -1
' Append Phone Field
Set TourField = TourTableDef.CreateField("Phone", dbText, 25)
    TourTableDef.Fields.Append TourField
' Append Fax Field
Set TourField = TourTableDef.CreateField("Fax", dbText, 25)
    TourTableDef.Fields.Append TourField
' Append E-Mail Field
Set TourField = TourTableDef.CreateField("E-Mail", dbText, 50)
    TourTableDef.Fields.Append TourField
' Append Count Field
Set TourField = TourTableDef.CreateField(gcContact_Count, dbLong)
    TourTableDef.Fields.Append TourField
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
    
' ---------------
' Append DaiIndex
' ---------------
' Create primary index.
Set TourIndex = TourTableDef.CreateIndex("ConIndex")
    TourIndex.Primary = False
    TourIndex.Required = True
    TourIndex.Unique = False
Set TourField = TourTableDef.CreateField("Contact")
    TourIndex.Fields.Append TourField
ProgressBar "Creating ConcTour Database...", -1, 7, -1
    
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' Create Second Index
Set TourIndex = TourTableDef.CreateIndex("CouIndex")
    TourIndex.Primary = True
    TourIndex.Required = True
    TourIndex.Unique = True
Set TourField = TourTableDef.CreateField("Count")
    'TourIndex.Attributes = dbDescending
    TourIndex.Fields.Append TourField
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_ContTour_Contacts_Plus_Fields = True
ProgressBar "Creating ConcTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0

Exit Function
Cont_Contacts_Err:
    If Err = 3204 Then          'Database already exist
        Err = 0
        Define_ContTour_Contacts_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_ContTour_Contacts_Plus_Fields-TourFunc"
    Resume Next


End Function

Function Define_Dai_Tour_Dai_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Dai_Fields_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating DaiTour Database...", -1, 2, -1
' Create new TableDef.
    Set TourTableDef = TourDatabase.CreateTableDef("Dai")

' Append Date Field
Set TourField = TourTableDef.CreateField("Date", dbDate)
    TourTableDef.Fields.Append TourField
' Append Type Field
Set TourField = TourTableDef.CreateField(gcDai_Tour_DaiType, dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append DayType Field
Set TourField = TourTableDef.CreateField("DayType", dbText, 100)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Heart Field
Set TourField = TourTableDef.CreateField("Heart", dbInteger)
    TourTableDef.Fields.Append TourField
' Append DayRate
Set TourField = TourTableDef.CreateField("DayRate", dbInteger)
    TourTableDef.Fields.Append TourField
' Append Oddometer
Set TourField = TourTableDef.CreateField("Oddometer", dbSingle)
    TourTableDef.Fields.Append TourField
' Append Weight
Set TourField = TourTableDef.CreateField("Weight", dbSingle)
    TourTableDef.Fields.Append TourField
' Append DaiMile
Set TourField = TourTableDef.CreateField("DaiMile", dbSingle)
    TourTableDef.Fields.Append TourField
ProgressBar "Creating DaiTour Database...", -1, 5, -1
' Append HeaV1
Set TourField = TourTableDef.CreateField("HeaV1", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append HeaV2
Set TourField = TourTableDef.CreateField("HeaV2", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append HeaV3
Set TourField = TourTableDef.CreateField("HeaV3", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append HeaV4
Set TourField = TourTableDef.CreateField("HeaV4", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append HeaV5
Set TourField = TourTableDef.CreateField("HeaV5", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append HeaV6
Set TourField = TourTableDef.CreateField("HeaV6", dbText, 8)
    TourTableDef.Fields.Append TourField
ProgressBar "Creating DaiTour Database...", -1, 7, -1
' Append HeaV7
Set TourField = TourTableDef.CreateField("HeaV7", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append HeaV8
Set TourField = TourTableDef.CreateField("HeaV8", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append HeaV9
Set TourField = TourTableDef.CreateField("HeaV9", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append Peak_Day
Set TourField = TourTableDef.CreateField("Peak_Day", dbText, 100)
    TourTableDef.Fields.Append TourField
' Append Peak_Wee
Set TourField = TourTableDef.CreateField("Peak_Wee", dbText, 100)
    TourTableDef.Fields.Append TourField
' Append Peak_Mon
Set TourField = TourTableDef.CreateField("Peak_Mon", dbText, 100)
    TourTableDef.Fields.Append TourField
' Append Memo
Set TourField = TourTableDef.CreateField("Memo", dbMemo)
    TourTableDef.Fields.Append TourField
' Append Sleep
Set TourField = TourTableDef.CreateField("Sleep", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append Level
Set TourField = TourTableDef.CreateField("Level", dbText, 1)
    TourTableDef.Fields.Append TourField
' Append Level1
Set TourField = TourTableDef.CreateField("Level1", dbText, 3)
    TourTableDef.Fields.Append TourField
' Append Level2
Set TourField = TourTableDef.CreateField("Level2", dbText, 3)
    TourTableDef.Fields.Append TourField
' Append Level3
Set TourField = TourTableDef.CreateField("Level3", dbText, 3)
    TourTableDef.Fields.Append TourField
' Append Level4
Set TourField = TourTableDef.CreateField("Level4", dbText, 3)
    TourTableDef.Fields.Append TourField
' Append Level5
Set TourField = TourTableDef.CreateField("Level5", dbText, 3)
    TourTableDef.Fields.Append TourField
' Append TotalHrs
Set TourField = TourTableDef.CreateField("TotalHrs", dbText, 8)
    TourTableDef.Fields.Append TourField
' Append WeeklySunday
Set TourField = TourTableDef.CreateField(gcDai_Tour_WeeklySunday, dbDate)
    TourTableDef.Fields.Append TourField
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
' ---------------
' Append DaiIndex
' ---------------
' Create primary index.
Set TourIndex = TourTableDef.CreateIndex("DaiIndex")
    TourIndex.Primary = True
    TourIndex.Required = True
    TourIndex.Unique = True
With TourIndex
    .Fields.Append TourTableDef.CreateField("Date")
    .Fields.Append TourTableDef.CreateField("Id")
    .Fields.Append TourTableDef.CreateField("Type")
End With
ProgressBar "Creating DaiTour Database...", -1, 9, -1
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' Create Second Index
Set TourIndex = TourTableDef.CreateIndex("TypeIndex")
    TourIndex.Primary = False
    TourIndex.Required = False
    TourIndex.Unique = False
    
Set TourField = TourTableDef.CreateField("DayType")
    TourIndex.Fields.Append TourField
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_Dai_Tour_Dai_Plus_Fields = True
ProgressBar "Creating DaiTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0
Exit Function
Dai_Fields_Err:
    If Err = 3204 Then      'Database already exist
        Define_Dai_Tour_Dai_Plus_Fields = False
        Err = 0
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_Dai_Fields-TourFunc"
    Resume Next
End Function


Function Define_Eve_Tour_Daily_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Eve_Daily_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin") ' Fill in new database.
ProgressBar "Creating Eve_Tour Database...", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("Daily")

' Append Date Field
Set TourField = TourTableDef.CreateField("Date", dbDate)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append DayType Field
Set TourField = TourTableDef.CreateField("DayType", dbText, 25)
    TourTableDef.Fields.Append TourField
ProgressBar "Creating Eve_Tour Database...", -1, 5, -1
' Append Color Field
Set TourField = TourTableDef.CreateField("Color", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = False
    TourTableDef.Fields.Append TourField
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef

' Add index
' ---------
        
Set TourIndex = TourTableDef.CreateIndex("DtIndex")
    TourIndex.Primary = True
    TourIndex.Unique = True
    TourIndex.Required = True
With TourIndex
    .Fields.Append TourTableDef.CreateField("Date")
    .Fields.Append TourTableDef.CreateField("Id")
End With
ProgressBar "Creating Eve_Tour Database...", -1, 8, -1
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_Eve_Tour_Daily_Plus_Fields = True
ProgressBar "Creating Eve_Tour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0
Exit Function
Eve_Daily_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_Eve_Tour_Daily_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_Eve_Tour_Daily_Plus_Fields-TourFunc"
    Resume Next
        
End Function

Function Define_Eve_Tour_Event_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Eve_Event_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating Eve_Tour Database...", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("Event")

' Append Date Field
Set TourField = TourTableDef.CreateField("Date", dbDate)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Page Field
Set TourField = TourTableDef.CreateField("Page", dbInteger)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Evememo Field
Set TourField = TourTableDef.CreateField("Evememo", dbText, 100)
    TourTableDef.Fields.Append TourField
ProgressBar "Creating Eve_Tour Database...", -1, 5, -1
' Append EveType Field
Set TourField = TourTableDef.CreateField("EveType", dbText, 20)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color Field
Set TourField = TourTableDef.CreateField("Color", dbLong)
    TourField.DefaultValue = 16777215 ' Set Default value
    TourTableDef.Fields.Append TourField
' Append PointToCont Field
Set TourField = TourTableDef.CreateField("PointToCont", dbInteger)
    TourTableDef.Fields.Append TourField
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
ProgressBar "Creating Eve_Tour Database...", -1, 7, -1

' Add index
' ---------
Set TourIndex = TourTableDef.CreateIndex("DateIndex")
    TourIndex.Primary = True
    TourIndex.Unique = True
    TourIndex.Required = True
With TourIndex
    .Fields.Append TourTableDef.CreateField("Date")
    .Fields.Append TourTableDef.CreateField("Id")
End With
    
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_Eve_Tour_Event_Plus_Fields = True
ProgressBar "Creating Eve_Tour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0

Exit Function
Eve_Event_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_Eve_Tour_Event_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_Eve_Tour_Event_Plus_Fields-TourFunc"
    Resume Next
End Function

Function Define_Eve_Tour_Peak_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Eve_Peak_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating Eve_Tour Database...", -1, 2, -1

' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef(gcPeak_Table)

' Append Date Field
Set TourField = TourTableDef.CreateField("Date", dbDate)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Descr Field
Set TourField = TourTableDef.CreateField("Descr", dbText, 100)
    TourTableDef.Fields.Append TourField
' Append Active Field
Set TourField = TourTableDef.CreateField("Active", dbInteger)
    TourTableDef.Fields.Append TourField
' Append Color Field
Set TourField = TourTableDef.CreateField("Color", dbLong)
    TourTableDef.Fields.Append TourField
ProgressBar "Creating Eve_Tour Database...", -1, 5, -1
' Append CycleName Field
Set TourField = TourTableDef.CreateField("CycleName", dbText, 50)
    TourTableDef.Fields.Append TourField
' Append Page Field
Set TourField = TourTableDef.CreateField("Page", dbInteger)
    TourTableDef.Fields.Append TourField
' Append Event_ID Field
Set TourField = TourTableDef.CreateField("Event_ID", dbInteger)
    TourTableDef.Fields.Append TourField
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
ProgressBar "Creating Eve_Tour Database...", -1, 7, -1

' Add index
' ---------
Set TourIndex = TourTableDef.CreateIndex("ActIndex")
    TourIndex.Primary = True
    TourIndex.Unique = False
    TourIndex.Required = False
Set TourField = TourTableDef.CreateField("Active")
    TourIndex.Fields.Append TourField
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' Add index
' ---------
Set TourIndex = TourTableDef.CreateIndex("DateIndex")
    TourIndex.Primary = False
    TourIndex.Unique = False
    TourIndex.Required = True
With TourIndex
    .Fields.Append TourTableDef.CreateField("Date")
    .Fields.Append TourTableDef.CreateField("Id")
End With
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_Eve_Tour_Peak_Plus_Fields = True
ProgressBar "Creating Eve_Tour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0

Exit Function
Eve_Peak_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_Eve_Tour_Peak_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_Eve_Tour_Peak_Plus_Fields-TourFunc"
    Resume Next


End Function

Function Define_Eve_Tour_Event_Tracker_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Eve_Event_Tracker_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating Eve_Tour Database...", -1, 2, -1

' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("Event_Tracker")

' Append Id Field
Set TourField = TourTableDef.CreateField("ID", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Event_Id Field
Set TourField = TourTableDef.CreateField("Event_ID", dbInteger)
    TourTableDef.Fields.Append TourField
' Append Name Field
Set TourField = TourTableDef.CreateField("Name", dbText, 255)
    TourTableDef.Fields.Append TourField
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
ProgressBar "Creating Eve_Tour Database...", -1, 7, -1

' Add index
' ---------
Set TourIndex = TourTableDef.CreateIndex("Event_Index")
    TourIndex.Primary = True
    TourIndex.Unique = False
    TourIndex.Required = False
Set TourField = TourTableDef.CreateField("Event_ID")
    TourIndex.Fields.Append TourField
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close

Define_Eve_Tour_Event_Tracker_Plus_Fields = True
ProgressBar "Creating Eve_Tour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0

Exit Function
Eve_Event_Tracker_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_Eve_Tour_Event_Tracker_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_Eve_Tour_Event_Tracker_Plus_Fields "
    Resume Next

End Function


Function Define_NameTour_Events_Plus_Fields(CreatePath As String) As Boolean
On Local Error GoTo Name_Events_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating NameTour Database ", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("Events")
' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Event0 Field
Set TourField = TourTableDef.CreateField("Event0", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event1 Field
Set TourField = TourTableDef.CreateField("Event1", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event2 Field
Set TourField = TourTableDef.CreateField("Event2", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event3 Field
Set TourField = TourTableDef.CreateField("Event3", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event4 Field
Set TourField = TourTableDef.CreateField("Event4", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event5 Field
Set TourField = TourTableDef.CreateField("Event5", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event6 Field
Set TourField = TourTableDef.CreateField("Event6", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event7 Field
Set TourField = TourTableDef.CreateField("Event7", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event8 Field
Set TourField = TourTableDef.CreateField("Event8", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Event9 Field
Set TourField = TourTableDef.CreateField("Event9", dbText, 30)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
    
' Append Color0 Field
Set TourField = TourTableDef.CreateField("Color0", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color1 Field
Set TourField = TourTableDef.CreateField("Color1", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color2 Field
Set TourField = TourTableDef.CreateField("Color2", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color3 Field
Set TourField = TourTableDef.CreateField("Color3", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color4 Field
Set TourField = TourTableDef.CreateField("Color4", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color5 Field
Set TourField = TourTableDef.CreateField("Color5", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color6 Field
Set TourField = TourTableDef.CreateField("Color6", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color7 Field
Set TourField = TourTableDef.CreateField("Color7", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color8 Field
Set TourField = TourTableDef.CreateField("Color8", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color9 Field
Set TourField = TourTableDef.CreateField("Color9", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
ProgressBar "Creating NameTour Database...", -1, 7, -1

' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_NameTour_Events_Plus_Fields = True
ProgressBar "Creating NameTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0
Exit Function
Name_Events_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_NameTour_Events_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_NameTour_Events_Plus_Fields-TourFunc"
    Resume Next


End Function

Function Define_NameTour_Levels_Plus_Fields(CreatePath As String) As Boolean
On Local Error GoTo Name_Levels_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating NameTour Database ", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef(gcLEVELS_TABLE)
' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Level1 Field
Set TourField = TourTableDef.CreateField("Level1", dbText, 3)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Level2 Field
Set TourField = TourTableDef.CreateField("Level2", dbText, 3)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Level3 Field
Set TourField = TourTableDef.CreateField("Level3", dbText, 3)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Level4 Field
Set TourField = TourTableDef.CreateField("Level4", dbText, 3)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Level5 Field
Set TourField = TourTableDef.CreateField("Level5", dbText, 3)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
ProgressBar "Creating NameTour Database...", -1, 7, -1

' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_NameTour_Levels_Plus_Fields = True
ProgressBar "Creating NameTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0
Exit Function
Name_Levels_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_NameTour_Levels_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_NameTour_Levels_Plus_Fields-TourFunc"
    Resume Next
End Function
'Function Define_NameTour_HeartNames_Plus_Fields(CreatePath As String) As Boolean
'' ---------------------------
'' Add field(s) to MyTableDef.
'' ---------------------------
'On Local Error GoTo Name_Names_Err
'Dim DefaultWorkspace As Workspace
'Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
'Dim TourIndex As Index, I As Integer, FieldVar As String
'Set DefaultWorkspace = DBEngine.Workspaces(0)
'' Create new, Decrypted database.
'Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
'
'' Fill in new database.
'ProgressBar "Creating NameTour Database ", -1, 2, -1
'' Create new TableDef.
'Set TourTableDef = TourDatabase.CreateTableDef("HeartNames")
'' Append Id Field
'Set TourField = TourTableDef.CreateField("Id", dbLong)
'    TourField.Required = True
'    TourTableDef.Fields.Append TourField
'For I = 0 To 9
'    ' Append Heart1..9 Field
'    FieldVar = "Heart" & Format$(I, "0")
'    Set TourField = TourTableDef.CreateField(FieldVar, dbText, 15)
'        TourField.AllowZeroLength = True
'    TourTableDef.Fields.Append TourField
'Next I
'
'' -----------------------------------------------------------------
'' Save TableDef definition by appending it to TableDefs collection.
'' -----------------------------------------------------------------
'TourDatabase.TableDefs.Append TourTableDef
'ProgressBar "Creating NameTour Database...", -1, 7, -1
'
'' ---------------------------------
'' Close newly created TourDatabase.
'' ---------------------------------
'TourDatabase.Close
'Define_NameTour_HeartNames_Plus_Fields = True
'ProgressBar "Creating NameTour Database...", -1, 10, -1
'ProgressBar "", 0, 0, 0
'Exit Function
'Name_Names_Err:
'    If Err = 3204 Then      'DataBAse already exist
'        Err = 0
'        Define_NameTour_HeartNames_Plus_Fields = False
'        Exit Function
'    End If
'    If bDebug Then Handle_Err Err, "Define_NameTour_HeartNames_Plus_Fields-TourFunc"
'    Resume Next
'End Function
Function Define_NameTour_PeakNames_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Name_PeakNames_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating NameTour Database ", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("PeakNames")
' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Peak0 Field
Set TourField = TourTableDef.CreateField("Peak0", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak1 Field
Set TourField = TourTableDef.CreateField("Peak1", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak2 Field
Set TourField = TourTableDef.CreateField("Peak2", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak3 Field
Set TourField = TourTableDef.CreateField("Peak3", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak4 Field
Set TourField = TourTableDef.CreateField("Peak4", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak5 Field
Set TourField = TourTableDef.CreateField("Peak5", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak6 Field
Set TourField = TourTableDef.CreateField("Peak6", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak7 Field
Set TourField = TourTableDef.CreateField("Peak7", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak8 Field
Set TourField = TourTableDef.CreateField("Peak8", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak9 Field
Set TourField = TourTableDef.CreateField("Peak9", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
Set TourField = TourTableDef.CreateField("Peak10", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak1 Field
Set TourField = TourTableDef.CreateField("Peak11", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak2 Field
Set TourField = TourTableDef.CreateField("Peak12", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak3 Field
Set TourField = TourTableDef.CreateField("Peak13", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak4 Field
Set TourField = TourTableDef.CreateField("Peak14", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak5 Field
Set TourField = TourTableDef.CreateField("Peak15", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak6 Field
Set TourField = TourTableDef.CreateField("Peak16", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak7 Field
Set TourField = TourTableDef.CreateField("Peak17", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak8 Field
Set TourField = TourTableDef.CreateField("Peak18", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Peak19 Field
Set TourField = TourTableDef.CreateField("Peak19", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
    
' Append Color0 Field
Set TourField = TourTableDef.CreateField("Color0", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color1 Field
Set TourField = TourTableDef.CreateField("Color1", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color2 Field
Set TourField = TourTableDef.CreateField("Color2", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color3 Field
Set TourField = TourTableDef.CreateField("Color3", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color4 Field
Set TourField = TourTableDef.CreateField("Color4", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color5 Field
Set TourField = TourTableDef.CreateField("Color5", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color6 Field
Set TourField = TourTableDef.CreateField("Color6", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color7 Field
Set TourField = TourTableDef.CreateField("Color7", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color8 Field
Set TourField = TourTableDef.CreateField("Color8", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color9 Field
Set TourField = TourTableDef.CreateField("Color9", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color0 Field
Set TourField = TourTableDef.CreateField("Color10", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color1 Field
Set TourField = TourTableDef.CreateField("Color11", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color2 Field
Set TourField = TourTableDef.CreateField("Color12", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color3 Field
Set TourField = TourTableDef.CreateField("Color13", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color4 Field
Set TourField = TourTableDef.CreateField("Color14", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color5 Field
Set TourField = TourTableDef.CreateField("Color15", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color6 Field
Set TourField = TourTableDef.CreateField("Color16", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color7 Field
Set TourField = TourTableDef.CreateField("Color17", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' Append Color18 Field
Set TourField = TourTableDef.CreateField("Color18", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
    
Set TourField = TourTableDef.CreateField("Color19", dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
    
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
ProgressBar "Creating NameTour Database...", -1, 7, -1

' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_NameTour_PeakNames_Plus_Fields = True
ProgressBar "Creating NameTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0
Exit Function
Name_PeakNames_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_NameTour_PeakNames_Plus_Fields = False
        Exit Function
    End If
    MsgBox Error$(Err)
    If bDebug Then Handle_Err Err, "Define_NameTour_PeakNames_Plus_Fields-TourFunc"
    Resume Next


End Function


Function Define_PeakTour_Peak_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo Peak_Peak_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")

' Fill in new database.
ProgressBar "Creating PeakTour Database...", -1, 2, -1

' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef(gcPeaks_Table)

' Append Id Field
Set TourField = TourTableDef.CreateField(gcPEAK_ID, dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append P_Date Field
Set TourField = TourTableDef.CreateField(gcPEAK_DATE, dbDate)
    TourTableDef.Fields.Append TourField
' Append P_Sched Field
Set TourField = TourTableDef.CreateField(gcPEAK_SCHED, dbMemo)
    TourTableDef.Fields.Append TourField
' Append P_Nam Field
Set TourField = TourTableDef.CreateField(gcPEAK_NAME, dbText, 35)
    TourTableDef.Fields.Append TourField
' Append Length Field
Set TourField = TourTableDef.CreateField(PEAKS_LENGTH, dbInteger)
    TourTableDef.Fields.Append TourField
    
    
ProgressBar "Creating PeakTour Database...", -1, 7, -1
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
    TourDatabase.TableDefs.Append TourTableDef
    
' ---------------
' Append DaiIndex
' ---------------
' Create Primary Index
Set TourIndex = TourTableDef.CreateIndex("NameIndex")
    TourIndex.Primary = True
    TourIndex.Required = True
    TourIndex.Unique = True
Set TourField = TourTableDef.CreateField(gcPEAK_NAME)
    TourIndex.Fields.Append TourField
    
Set TourField = TourTableDef.CreateField(gcPEAK_ID)
    TourIndex.Fields.Append TourField

' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' Create Second index.
Set TourIndex = TourTableDef.CreateIndex("DateIndex")
    TourIndex.Primary = False
    TourIndex.Required = False
    TourIndex.Unique = False
Set TourField = TourTableDef.CreateField(gcPEAK_DATE)
    TourIndex.Fields.Append TourField
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_PeakTour_Peak_Plus_Fields = True
ProgressBar "Creating PeakTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0

Exit Function
Peak_Peak_Err:
    If Err = 3204 Then      'DataBase already exist
        Err = 0
        Define_PeakTour_Peak_Plus_Fields = False
        Exit Function
    End If
   
    If bDebug Then Handle_Err Err, "Define_PeakTour_Peak_Plus_Fields-TourFunc"
    Resume Next
End Function

Function Define_UserTour_ContactOpt_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo ContactOpt_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating UserTour Database...", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("ContactOpt")

' Append Id Field
Set TourField = TourTableDef.CreateField("Id", dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append IndexOrder Field
Set TourField = TourTableDef.CreateField("IndexOrder", dbText, 30)
    TourTableDef.Fields.Append TourField
ProgressBar "Creating UserTour Database...", -1, 5, -1
' Append ColumnOrder Field
Set TourField = TourTableDef.CreateField("ColumnOrder", dbText, 10)
    TourTableDef.Fields.Append TourField
' Append SortOrder Field
Set TourField = TourTableDef.CreateField("SortOrder", dbText, 4)
    TourTableDef.Fields.Append TourField
' Append ContactWidth Field
Set TourField = TourTableDef.CreateField("ContactWidth", dbSingle)
    TourTableDef.Fields.Append TourField
' Append LastWidth Field
Set TourField = TourTableDef.CreateField("LastWidth", dbSingle)
    TourTableDef.Fields.Append TourField
' Append FirstWidth Field
Set TourField = TourTableDef.CreateField("FirstWidth", dbSingle)
    TourTableDef.Fields.Append TourField
' Append PhoneWidth Field
Set TourField = TourTableDef.CreateField("PhoneWidth", dbSingle)
    TourTableDef.Fields.Append TourField
' Append FaxWidth Field
Set TourField = TourTableDef.CreateField("FaxWidth", dbSingle)
    TourTableDef.Fields.Append TourField
' Append E-MailWidth Field
Set TourField = TourTableDef.CreateField("E-MailWidth", dbSingle)
    TourTableDef.Fields.Append TourField
    
Define_UserTour_ContactOpt_Plus_Fields = True
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
ProgressBar "Creating UserTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0
Exit Function
ContactOpt_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_UserTour_ContactOpt_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_UserTour_ContactOpt_Plus_Fields-TourFunc"
    Resume Next
        
End Function

Function Define_UserTour_UserTbl_Plus_Fields(CreatePath As String) As Boolean
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo UserTbl_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index


Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.CreateDatabase(CreatePath & "\" & gcTour_Win, dbLangGeneral, dbEncrypt + dbVersion30)
    TourDatabase.NewPassword "", gcTOURWIN_PASSWORD
' Fill in new database.
ProgressBar "Creating UserTour Database...", -1, 2, -1
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef("UserTbl")
' Append Id Field
Set TourField = TourTableDef.CreateField(gcUserTour_UserTbl_Name, dbText, 15)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append id Field
Set TourField = TourTableDef.CreateField(gcID, dbLong)
    TourField.Required = True
    TourField.Attributes = dbAutoIncrField
    TourTableDef.Fields.Append TourField
' Append PassWord Field
Set TourField = TourTableDef.CreateField("PassWord", dbText, 15)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
ProgressBar "Creating UserTour Database...", -1, 5, -1
' Append DataPath Field
Set TourField = TourTableDef.CreateField("DataPath", dbText, 255)
    TourTableDef.Fields.Append TourField
'' Append Security Field
'Set TourField = TourTableDef.CreateField("Security", dbBoolean)
'    TourTableDef.Fields.Append TourField
'' Append Load Field
'Set TourField = TourTableDef.CreateField("Load", dbByte)
'    TourTableDef.Fields.Append TourField
' Append MetaFile Field
Set TourField = TourTableDef.CreateField("Metafile", dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
'' Append ShowMeta Field
'Set TourField = TourTableDef.CreateField("ShowMeta", dbBoolean)
'    TourTableDef.Fields.Append TourField
' Append BitField Field
Set TourField = TourTableDef.CreateField("BitField", dbInteger)
    TourField.DefaultValue = 1
    TourTableDef.Fields.Append TourField
    
' Append DailyOpt Field
Set TourField = TourTableDef.CreateField(gcUserTour_UserTbl_DailyOptions, dbLong)
    TourField.DefaultValue = 63     ' Default with all Daily options visible!
    TourTableDef.Fields.Append TourField
    
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef

' Add index
' ---------
Set TourIndex = TourTableDef.CreateIndex("IdIndex")
    TourIndex.Primary = True
    TourIndex.Unique = True
    TourIndex.Required = True
Set TourField = TourTableDef.CreateField("Id")
    TourIndex.Fields.Append TourField
    
ProgressBar "Creating UserTour Database...", -1, 8, -1

Define_UserTour_UserTbl_Plus_Fields = True

' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close

ProgressBar "Creating UserTour Database...", -1, 10, -1
ProgressBar "", 0, 0, 0
Exit Function
UserTbl_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        Define_UserTour_UserTbl_Plus_Fields = False
        Exit Function
    End If
    If bDebug Then Handle_Err Err, "Define_UserTour_UserTbl_Plus_Fields-TourFunc"
    Resume Next
         
End Function
Public Function Define_DaiPercentage_QueryDef(ByVal sPath As String) As Boolean
On Local Error GoTo DaiPercentage_Error
' Declare local variables
Dim qTemp As QueryDef
Dim sTemp As String

sTemp = "SELECT Dai.WeeklySunday, Sum(Dai.DaiMile) AS Distance, Sum((Left(Dai.HeaV1,2)*3600) + Right(Dai.HeaV1,2) + (MID(Dai.HeaV1,4,2) * 60)) AS HeaV1, Sum((Left(Dai.HeaV2,2)*3600) + Right(Dai.HeaV2,2) + (MID(Dai.HeaV2,4,2) * 60)) AS HeaV2, Sum((Left(Dai.HeaV4,2)*3600) + Right(Dai.HeaV4,2) + (MID(Dai.HeaV4,4,2) * 60)) AS HeaV4, Sum((Left(Dai.HeaV3,2)*3600) + Right(Dai.HeaV3,2) + (MID(Dai.HeaV3,4,2) * 60)) AS HeaV3, Sum((Left(Dai.HeaV5,2)*3600) + Right(Dai.HeaV5,2) + (MID(Dai.HeaV5,4,2) * 60)) AS HeaV5, Sum((Left(Dai.HeaV6,2)*3600) + Right(Dai.HeaV6,2) + (MID(Dai.HeaV6,4,2) * 60)) AS HeaV6, Sum((Left(Dai.HeaV7,2)*3600) + Right(Dai.HeaV7,2) + (MID(Dai.HeaV7,4,2) * 60)) AS HeaV7, Sum((Left(Dai.HeaV8,2)*3600) + Right(Dai.HeaV8,2) + (MID(Dai.HeaV8,4,2) * 60)) AS HeaV8, Sum((Left(Dai.HeaV9,2)*3600) + Right(Dai.HeaV9,2) + (MID(Dai.HeaV9,4,2) * 60)) AS HeaV9, Sum((Left(Dai.TotalHrs,2)*3600) + Right(Dai.TotalHrs,2) + (MID(Dai.TotalHrs,4,2) * 60)) AS TotalHrs " & _
        "From Dai WHERE Type = 1 " & _
        "GROUP BY Dai.WeeklySunday"

ObjTour.CreateQueryDefObject sPath, "DaiPercentage", sTemp

Define_DaiPercentage_QueryDef = True
On Local Error GoTo 0
Exit Function

DaiPercentage_Error:
    Define_DaiPercentage_QueryDef = False
    If bDebug Then Handle_Err Err, "Define_DaiPercentage_QueryDef-TourFunc"
    Resume Next

End Function

Sub EditErrorFrm(The_Colprit As String)
' --------------------------------------
' Load the errorfrm with valuable user
' information like the form which caused
' the problem and load ~tourwin.log into
' ErrRecTxt...
' ---------------------------------------
Dim DesTxt As String, RecTxt As String
Dim TourTxt As String
DesTxt = "An error has occurred in " & Trim$(The_Colprit) & ". If problem "
DesTxt = DesTxt & "persists contact manufactor. Choose Techical Support "
DesTxt = DesTxt & "menu under Help for contacting manufactor options."
ErrorFrm.ErrDesTxt = DesTxt

Open App.Path & "\~TourWin.log" For Input As #5
    While Not EOF(5)
        Line Input #5, TourTxt
        RecTxt = RecTxt & TourTxt & vbLf
    Wend
    Close 5
ErrorFrm.ErrRecTxt = RecTxt
End Sub

Function Get_NameTour_Events(FieldType As String) As String

' -----------------------------------------
' Purpose:  Is a central and Only location
'           which accesses the Events Table
'           found in the NameTour database.
' -----------------------------------------
On Local Error GoTo Get_NameTour_Events_Err
Dim SQL As String, RetStr As String
' ------------------------
' set Default return value
' ------------------------
Get_NameTour_Events = "No Return"

If ObjTour.RstRecordCount(iSearcherDB) <> 0 Then
    ' --------------------------
    ' Set TypeNum = Event# field
    ' --------------------------
    RetStr = IIf("" = ObjTour.DBGetField(FieldType, iSearcherDB), "No Return", ObjTour.DBGetField(FieldType, iSearcherDB))
        If RetStr <> "No Return" And Trim$(RetStr) <> "" Then
                Get_NameTour_Events = Trim(RetStr)
        End If
End If  'Ends if RecordCount = 0

Exit Function
Get_NameTour_Events_Err:
    If bDebug Then Handle_Err Err, "Get_NameTour_Events-TourFunc"
    Exit Function
End Function
Function Get_NameTour_PeakNames(FieldType As String) As String
' -----------------------------------------
' Purpose:  Is a central and Only location
'           which accesses the PeakNames table
'           found in the NameTour database.
' -----------------------------------------
On Local Error GoTo Get_NameTour_PeakNames_Err
Dim SQL As String, RetStr As String
' ------------------------
' set Default return value
' ------------------------
Get_NameTour_PeakNames = "No Return"

If ObjTour.RstRecordCount(iSearcherDB) <> 0 Then
    ' --------------------------
    ' Set TypeNum = Event# field
    ' --------------------------
    RetStr = IIf("" = ObjTour.DBGetField(FieldType, iSearcherDB), "No Return", ObjTour.DBGetField(FieldType, iSearcherDB))
        If RetStr <> "No Return" And Trim$(RetStr) <> "" Then
                Get_NameTour_PeakNames = Trim(RetStr)
        End If
End If                              'Ends if RecordCount = 0
Exit Function
Get_NameTour_PeakNames_Err:
    If bDebug Then Handle_Err Err, "Get_NameTour_PeakNames-TourFunc"
    Resume Next

End Function


Sub Printer_Setup()
On Local Error GoTo P_Setup
    Call ShowPrinter(MDI.hWnd)
Exit Sub
P_Setup:
    If bDebug Then Handle_Err Err, "Printer_Setup-TourFunc"
    Resume Next
End Sub

Sub Weekly_Percentage_Report()

Dim StartD      As String
Dim EndD        As String
Dim sDataFiles  As String
Dim cH_Name     As cHeartNameVar
Dim CryFormulas As CRPEAuto.FormulaFieldDefinitions
Dim CryFormula  As CRPEAuto.FormulaFieldDefinition

On Local Error GoTo Rep_Err

sDataFiles = objMdi.info.Datapath & gcTour_Win
DateFrm.Show vbModal

If UserCancel = True Then Exit Sub
' Make sure DSN is pointing to current database
ObjTour.RegisterUpdateDSN objMdi.info.Datapath & gcTour_Win
        
StartD = Format$(DateFrm.DatFroTxt, "YYYY, MM, DD")
EndD = Format$(DateFrm.DatToTxt, "YYYY, MM, DD")
If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing

' Create Crystal Objects
Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Weeklypc.rpt")

CryReport.RecordSelectionFormula = "{Dai.WeeklySunday} in Date (" & StartD & ") to Date (" & EndD & ") AND {Dai.Id} = " & objMdi.info.ID & " And {Dai.Type} = " & gcDAI_NORMAL
' Loop Through each formula and update accordingly
Set CryFormulas = CryReport.FormulaFields

cHeartNames.LoadValues objMdi.info.ID

For Each CryFormula In CryFormulas
    Select Case CryFormula.Name
        Case "{@rpt_date_range}":
            CryFormula.Text = Chr$(34) & LoadResString(5102) & Format$(StartD, "MMMM d, yyyy") & " to " & Format$(EndD, "MMMM d, yyyy") & Chr$(34)
        Case "{@rpt_title}":
            CryFormula.Text = Chr$(34) & LoadResString(5100) & Chr$(34)
        Case "{@rpt_sub_title}":
            CryFormula.Text = Chr$(34) & LoadResString(5101) & Chr$(34)
        Case Else
        If CryFormula.Name Like "{@HeaV#_Title}" Then
        
            Set cH_Name = cHeartNames(Mid$(CryFormula.Name, 7, 1))
            If Not cH_Name Is Nothing Then
                CryFormula.Text = Chr$(34) & cH_Name.Description & Chr$(34)
            End If
            
        End If
    End Select
Next

If Not CryFormulas Is Nothing Then Set CryFormulas = Nothing
If Not CryFormula Is Nothing Then Set CryFormula = Nothing
CryReport.Preview "Weekly Totals Report " & LoadResString(gcTourVersion)

On Error GoTo 0
Exit Sub

Rep_Err:
If Not CryFormulas Is Nothing Then Set CryFormulas = Nothing
If Not CryFormula Is Nothing Then Set CryFormula = Nothing
If Not CryReport Is Nothing Then Set CryReport = Nothing
    Select Case Err
        Case 13:      'File not found
        MsgBox Error$(Err)
        Case Else
        If bDebug Then Handle_Err Err, "TourFunc-Weekly_Percentage_Report"
        Resume Next
    End Select
Err.Clear

End Sub

Sub WeeklyTotal_Report()

Dim StartD      As String
Dim EndD        As String
Dim sDataFiles  As String
Dim cH_Name     As cHeartNameVar
Dim CryFormulas As CRPEAuto.FormulaFieldDefinitions
Dim CryFormula  As CRPEAuto.FormulaFieldDefinition


On Local Error GoTo Rep_Err

sDataFiles = objMdi.info.Datapath & gcTour_Win
DateFrm.Show vbModal
DoEvents

If UserCancel = True Then Exit Sub

' Make sure DSN is pointing to current database
ObjTour.RegisterUpdateDSN objMdi.info.Datapath & gcTour_Win

Screen.MousePointer = vbHourglass
    
StartD = Format$(DateFrm.DatFroTxt, "YYYY, MM, DD")
EndD = Format$(DateFrm.DatToTxt, "YYYY, MM, DD")

If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing

' Create Crystal Objects
Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Weekly.rpt")

CryReport.RecordSelectionFormula = "{Dai.Date} in Date (" & StartD & ") to Date (" & EndD & ") AND {Dai.Id} = " & objMdi.info.ID & " And {Dai.Type} = " & gcDAI_NORMAL
' Loop Through each formula and update accordingly
Set CryFormulas = CryReport.FormulaFields

cHeartNames.LoadValues objMdi.info.ID

For Each CryFormula In CryFormulas
    Select Case CryFormula.Name
        Case "{@rpt_date_range}":
            CryFormula.Text = Chr$(34) & LoadResString(5102) & Format$(StartD, "MMMM d, yyyy") & " to " & Format$(EndD, "MMMM d, yyyy") & Chr$(34)
        Case "{@rpt_title}":
            CryFormula.Text = Chr$(34) & LoadResString(5100) & Chr$(34)
        Case "{@rpt_sub_title}":
            CryFormula.Text = Chr$(34) & LoadResString(5101) & Chr$(34)
        Case Else
        If CryFormula.Name Like "{@HeaV#_Title}" Then
        
            Set cH_Name = cHeartNames(Mid$(CryFormula.Name, 7, 1))
            If Not cH_Name Is Nothing Then
                CryFormula.Text = Chr$(34) & cH_Name.Description & Chr$(34)
            End If
            
        End If
    End Select
Next
Screen.MousePointer = vbDefault
' Release child Crystal Objects
If Not CryFormula Is Nothing Then Set CryFormula = Nothing
If Not CryFormulas Is Nothing Then Set CryFormulas = Nothing

CryReport.Preview "Weekly Totals Report " & LoadResString(gcTourVersion)
    
On Local Error GoTo 0
Exit Sub

Rep_Err:
Screen.MousePointer = vbDefault
    If bDebug Then Handle_Err Err, "TourFunc-WeeklyTotal_Report"
  
    Select Case Err
        Case 13:
        MsgBox Error$(Err)
        Case 20533:
            MsgBox Error$(Err)
        Case 20504:      'File not found
            MsgBox Error$(Err)
        Case Else
        MsgBox Error$(Err)
        If bDebug Then Handle_Err Err, "MdiTotmnu-MDI"
        Exit Sub
    End Select
End Sub

' --------------------------
' Peak event Report function
' --------------------------
Sub Peak_Event_Report()

Dim StartD       As String
Dim EndD         As String
Dim sReportRange As String

On Local Error GoTo Rep_Err

' Prompt for date range...
DateFrm.Show 1

If UserCancel = True Then Exit Sub

StartD = Format$(DateFrm.DatFroTxt, "YYYY, MM, DD")
EndD = Format$(DateFrm.DatToTxt, "YYYY, MM, DD")
sReportRange = Format$(StartD, "mmmm dd, yyyy") & " to " & Format$(EndD, "mmmm dd, yyyy")

' Make sure DSN is pointing to current database
ObjTour.RegisterUpdateDSN objMdi.info.Datapath & gcTour_Win

If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing

Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Daily.rpt")
CryReport.RecordSelectionFormula = _
"{Dai.Date} in Date (" & StartD & ") to Date (" & EndD & ") AND {Dai.Id} = " & objMdi.info.ID

CryReport.Preview "Daily Activity Report. " & LoadResString(gcTourVersion)
'    For I = 1 To 9
'            HeartStr = "Heart" & Format$(I, "0")
'            RetStr = Get_NameTour_HeartNames(HeartStr)
'        If RetStr <> "No Return" Then
'            FormulaStr = "HeaV" & Format$(I, "0") & "='" & RetStr & "'"
'        Else
'            FormulaStr = "HeaV" & Format$(I, "0") & "=''"
'        End If
'        .Formulas(I) = FormulaStr
'    Next I
'    ' ------------------------
'    ' Define DateRange Formula
'    ' ------------------------
'        .Formulas(10) = "DateRange=" & "'" & sReportRange & "'"
'        Size_Crystal_Window
'        .Action = 2
'    End With
Exit Sub
Rep_Err:
    Select Case Err
        Case 13 Or 20533:      'File not found
        MsgBox Error$(Err)
        Case 20515:
                    MsgBox Error$(Err)
        Case Else
        If bDebug Then Handle_Err Err, "MdiTotmnu-MDI"
        Resume Next
    End Select

End Sub
Sub Daily_Report()


Dim StartD        As String
Dim EndD          As String
Dim sDataSource   As String
Dim sReportRange  As String
Dim cH_Name       As cHeartNameVar
Dim CryFormulas   As CRPEAuto.FormulaFieldDefinitions
Dim CryFormula    As CRPEAuto.FormulaFieldDefinition
Dim CryParameters As CRPEAuto.ParameterFieldDefinitions
Dim CryParameter  As CRPEAuto.ParameterFieldDefinition

On Local Error GoTo Rep_Err

If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing
    
sDataSource = objMdi.info.Datapath & gcTour_Win
DateFrm.Show 1

If UserCancel = True Then Exit Sub

Call ObjTour.RegisterUpdateDSN(objMdi.info.Datapath & gcTour_Win)

' Create Selection Formula string
StartD = Format$(DateFrm.DatFroTxt.Text, "YYYY, MM, DD")
EndD = Format$(DateFrm.DatToTxt.Text, "YYYY, MM, DD")
sReportRange = Format$(StartD, "mmmm dd, yyyy") & " to " & Format$(EndD, "mmmm dd, yyyy")
        
' Create Crystal Objects
Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Daily.rpt")
'Set CryFormulas = CryReport.for

CryReport.RecordSelectionFormula = "{Dai.Date} in Date (" & StartD & ") to Date (" & EndD & ") AND {Dai.Id} = " & objMdi.info.ID & " And {Dai.Type} = " & gcDAI_NORMAL

' Loop Through each formula and update accordingly
Set CryFormulas = CryReport.FormulaFields

cHeartNames.LoadValues objMdi.info.ID
For Each CryFormula In CryFormulas
    Select Case CryFormula.Name
        Case "{@DateRange}":
            CryFormula.Text = Chr$(34) & sReportRange & Chr$(34)
        Case Else
        If CryFormula.Name Like "{@HeaV#_Title}" Then
        
            Set cH_Name = cHeartNames(Mid$(CryFormula.Name, 7, 1))
            If Not cH_Name Is Nothing Then
                CryFormula.Text = Chr$(34) & cH_Name.Description & Chr$(34)
            End If
            
        End If
    End Select
Next

' Update Report Parameter
Set CryParameters = CryReport.ParameterFields

For Each CryParameter In CryParameters
    Select Case CryParameter.Name
        Case "{?ID}":
            CryParameter.SetCurrentValue objMdi.info.ID
    End Select
Next

' Release child Crystal Objects
If Not CryFormula Is Nothing Then Set CryFormula = Nothing
If Not CryFormulas Is Nothing Then Set CryFormulas = Nothing
If Not CryParameter Is Nothing Then Set CryParameter = Nothing
If Not CryParameters Is Nothing Then Set CryParameters = Nothing


CryReport.Preview "Daily Activity Report. " & LoadResString(gcTourVersion)

Exit Sub
Rep_Err:
If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing
If bDebug Then Handle_Err Err, "Daily Report"
MsgBox Error$(Err.Number)
    Select Case Err.Number
        Case 13 Or 20533:      'File not found
        MsgBox Error$(Err)
        Case 20515:
                    MsgBox Error$(Err)
        Case Else
        'If bDebug Then Handle_Err Err, "MdiTotmnu-MDI"
        Resume Next
    End Select

End Sub


Sub Handle_Err(iErr%, sCallingProc$)
'----------------------
' Purpose: 1) Tracks user actions and writes actions to ~tourwin.log file.
' If error occurs. Then ~Tourwin.log is copied to TourWin.log. The main
' purpose is for Debug errors. An action is distinguish from an error by
' the value of iErr%, if 0 then action....
'
' Writes to text file rather than
' showing the user ever error that
' occurs.
' CHANGE LOG --------------
'   June 9, 2003 - Make changes to write out ever log as soon as it happens
'                   This should improve the pin-pointing GPF and other anomillies
'----------------------
' Declare local variables

Dim iFile       As Integer    ' Handle to next free file handle                            ' automatically end program.
Static Audit    As String

' -------------------------------
' Prevent crashes in errorhandler
' -------------------------------
On Error Resume Next


If iErr% = 0 Then
    Audit = Format$(Now, "HH:MM:SS MM-DD-YYYY") & "  " & Trim$(sCallingProc$)
Else
    Audit = Format$(Now, "HH:MM:SS MM-DD-YYYY") & "  " & sCallingProc & "    Error!!!   " & " " & Err.Source & "   " & Error$(iErr)
    ' ----------------------
    ' Actual error occurred!
    ' ----------------------
    ErrorOccurred = True
End If

' --------------------
' Audit trial portion.
' --------------------

iFile = FreeFile
If " " = Dir$(App.Path & "\" & TOURWINLOGFILE) Then
   Open App.Path & "\" & TOURWINLOGFILE For Output As #iFile
Else
   Open App.Path & "\" & TOURWINLOGFILE For Append As #iFile
End If

Print #iFile, Audit
Close iFile
Audit = ""

DoEvents
End Sub


Sub LoadObjects()

Set cTour_DB = New CDatabase       ' Global Database Object

' set Cont object
Set ObjCont = New cCont
Set ObjContVar = New cContVar
Set ObjCont.info = ObjContVar
' set cEve object
Set objEve = New cEve
Set objEveVar = New cEveVar
Set objEve.info = objEveVar

' set Chart object form controls
Set ObjChart = New cChart
Set ObjChartVar = New cChartVar
Set ObjChart.info = ObjChartVar

'' Set cTour object to Database Object
Set ObjTour = New cTourInfo
Set ObjTourVar = New cTourVar
Set ObjTour.Settings = ObjTourVar

Set cHeartNames = New cHeartNames
Set cActivityNames = New cActivity_Names
Set cPeakNames = New cPeak_Names
Set cPeakSchedules = New cPeak_Schedules
Set cExportFile = New CExport
Set cLicense = New CRegistration

ErrorOccurred = False

End Sub

Public Sub ProgressBar(LabStr$, LabVis%, TxtWdth%, TxtVis%)

'-----------------------------
' This sub is used to
' provide the user information
' will loading database, file
' etc...
' -----------------------------
On Local Error GoTo ProgressBar_Err
With MDI
    .MdiBarTxt = LabStr
    .ProgressBack.Text = LabStr
    
    .MdiBarTxt.Visible = TxtVis
    .ProgressBack.Visible = TxtVis
    .MdiBarTxt.Width = ((TxtWdth / 10) * .ProgressBack.Width)
End With
DoEvents
Exit Sub
ProgressBar_Err:
      If bDebug Then Handle_Err Err, "ProgressBar MDI"
    Resume Next
End Sub

Sub AboutMsg()
On Local Error GoTo about_Err

Dim Msg As String
    Msg = LoadResString(gcTourVersion) & vbLf
    Msg = Msg & " ©CopyWrite 1997. All rights reserved"
    MsgBox Msg, 64, "TourWin Written by Mark Ormesher."
    
Exit Sub
about_Err:
    If bDebug Then Handle_Err Err, "aboutMsg-Module2"
Resume Next
End Sub


Function Get_Eve_Peak() As String
On Local Error GoTo Get_Eve_Err
'
' Functionallity
'
Exit Function
Get_Eve_Err:
        If bDebug Then Handle_Err Err, "TourFunc-Module2"
        Resume Next
End Function




Sub TechSupport()
On Local Error GoTo TechSupport_Err

MsgBox LoadResString(gcTourCommentEmail), vbOKOnly + vbInformation, "TourWin written by " & LoadResString(gcLeaMarTech)
    
On Error GoTo 0
Exit Sub
TechSupport_Err:
    If bDebug Then Handle_Err Err, "TechSupport-TourFunc"
    Resume Next

End Sub

Function UpdateEventdb(EveTble$, Datte As Date, Descr$, Active%, Colour As Single, Descr2$) As Boolean

Dim SQL         As String
Dim SearchStr   As String

On Local Error GoTo UpdateEvent_Err

If bDebug Then Handle_Err 0, "UpdateEventdb-Module2"       'Audit trial...

SQL = "SELECT * FROM " & EveTble & " WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL lEventHandle, SQL

SearchStr = "Date = #" & Datte & "#"
ObjTour.DBFindFirst SearchStr, lEventHandle

If ObjTour.NoMatch(lEventHandle) Then
    ObjTour.AddNew lEventHandle
    ObjTour.DBSetField "Id", objMdi.info.ID, lEventHandle
Else
    ObjTour.Edit lEventHandle
End If

Select Case EveTble
   Case "Event":
            ObjTour.DBSetField "Date", Datte, lEventHandle
            ObjTour.DBSetField "Evemeno", Descr, lEventHandle
            ObjTour.DBSetField "EveType", Descr2, lEventHandle
            ObjTour.DBSetField "Color", Colour, lEventHandle
        Case "Peak":
            ObjTour.DBSetField "Date", Datte, lEventHandle
            ObjTour.DBSetField "Descr", Descr, lEventHandle
            ObjTour.DBSetField "Active", Active, lEventHandle
            ObjTour.DBSetField "Color", Colour, lEventHandle
            ObjTour.DBSetField "CycleName", Descr2, lEventHandle
        Case "Daily":
            ' -----------------------------------------------------
            ' Find Color that accompanies DayType Description. This
            ' info is stored in DB = NameTour Table = Events
            ' ------------------------------------------------------
        Colour = 12632256  ' If no match, default to grey
        ' ======================
        ' Define RecordSet then
        ' Get Fields
        ' =====================
        ' Define Record Set
        cActivityNames.Type_ID = gcActive_Type_EventNames
        If cActivityNames.FindItemByName(Descr) Then
            ObjTour.DBSetField "Color", cActivityNames.Colour, lEventHandle
        End If
        
        ObjTour.DBSetField "Date", Datte, lEventHandle
        ObjTour.DBSetField "DayType", Descr, lEventHandle

End Select

ObjTour.Update lEventHandle

Exit Function
UpdateEvent_Err:

    If bDebug Then Handle_Err Err, "UpdateEventdb-Module2"
    
    Resume Next
End Function

Public Function Size_Crystal_Window()
If gCrystalReport Is Nothing Then Exit Function
With gCrystalReport
        .WindowTop = (MDI.top + 620 + MDI.Toolbar1.Height) / 15
        .WindowLeft = 10 + (MDI.left / 15)
        .WindowWidth = (MDI.Width - 10) / (15.3)
        .WindowHeight = (MDI.Height - 1710) / 15
End With

End Function
Public Function Write_Tour_Setting_To_Registry(sSubKey As String, sKey As String, sValue As String)
' =======================================================
' Purpose: Checks that TourWin EXE path is specific
'          so in case of upgrades, files can be written
'          correct locaton.
' =======================================================
On Local Error GoTo Write_Setting_Error
Dim sRetStr As String

If bDebug Then
    Handle_Err 0, "Write_Tour_Settings_To_Registry-TourFunc"
End If

gbSkipRegErrMsg = True ' Don't show reg error

If REG_ERROR = GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcRegTourExe)) Then
    If True = CreateRegKey(LoadResString(gcRegTourKey)) Then
        gbSkipRegErrMsg = True
        WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcRegTourExe), App.Path
    End If
End If

If "" <> sValue Then
' --------------------------------------------
' Write Registry setting for passed parameters
' --------------------------------------------
gbSkipRegErrMsg = True ' Don't show reg error
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey) & "\" & sSubKey, sKey, sValue)

        If False = sRetStr Then
                gbSkipRegErrMsg = True ' Don't show reg error
                CreateRegKey LoadResString(gcRegTourKey) & "\" & sSubKey
                WriteRegStringValue LoadResString(gcRegTourKey) & "\" & sSubKey, sKey, sValue
        End If

Else

' ----------------------
' Write MDI Form size
' form Registry and set
' ----------------------

gbSkipRegErrMsg = True ' Don't show reg error

'' WindowState
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcWindowState), Str$(MDI.WindowState))
        If REG_ERROR = sRetStr Then
                gbSkipRegErrMsg = True ' Don't show reg error
                CreateRegKey LoadResString(gcRegTourKey) & "\" & LoadResString(gcWindowState)
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcWindowState), Str$(MDI.WindowState)
        End If

'' Top
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcTop_Form), Str$(Screen.ActiveForm.top))
        If REG_ERROR = sRetStr Then
                gbSkipRegErrMsg = True ' Don't show reg error
                CreateRegKey LoadResString(gcRegTourKey) & "\" & LoadResString(gcTop_Form)
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcTop_Form), Str$(Screen.ActiveForm.top)
        End If
'' Left
gbSkipRegErrMsg = True ' Don't show reg error
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcLeft_Form), Str$(Screen.ActiveForm.left))
        If REG_ERROR = sRetStr Then
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcLeft_Form), Str$(Screen.ActiveForm.left)
        End If
        
        
gbSkipRegErrMsg = True ' Don't show reg error
'' Height
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcHeight_Form), Str$(Screen.ActiveForm.Height))
        If REG_ERROR = sRetStr Then
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcHeight_Form), Str$(Screen.ActiveForm.Height)
        End If
        
        
gbSkipRegErrMsg = True ' Don't show reg error
'' Width
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcWidth_Form), Str$(Screen.ActiveForm.Width))
        If REG_ERROR = sRetStr Then
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcWidth_Form), Str$(Screen.ActiveForm.Height)
        End If
        
gbSkipRegErrMsg = True ' Don't show reg error
'' Width
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcTourUserName), objMdi.info.Name)
        If REG_ERROR = sRetStr Then
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcTourUserName), objMdi.info.Name
        End If
        
gbSkipRegErrMsg = True ' Don't show reg error
' Update Last Database used
sRetStr = WriteRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcTourLastDB), objMdi.info.Datapath)
        If REG_ERROR = sRetStr Then
                CreateRegKey LoadResString(gcRegTourKey) & "\" & LoadResString(gcTourLastDB)
                WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcTourLastDB), objMdi.info.Datapath
        End If
        
End If

' Check if current DB is listed in DBase key, if not add!
gbSkipRegErrMsg = True
sRetStr = GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcTourDBase))

' Check if key exist, if not create
If REG_ERROR = sRetStr Then
    gbSkipRegErrMsg = True
    CreateRegKey LoadResString(gcRegTourKey) & "\" & LoadResString(gcTourDBase)
    gbSkipRegErrMsg = True
    WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcTourDBase), objMdi.info.Datapath
Else
' -------------------------
' Check if current db is
' listed in return string
' if not, append to end.
' -------------------------
    If 0 = InStr(1, sRetStr, objMdi.info.Datapath & ",", vbTextCompare) Then
        gbSkipRegErrMsg = True
        WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcTourDBase), sRetStr & "," & objMdi.info.Datapath & ","
    End If
End If
On Error GoTo 0
Exit Function

Write_Setting_Error:
If bDebug Then
    Handle_Err Err, "Write_Tour_Settings_To_Registry-TourFunc"
End If
Resume Next

End Function

Public Function FindItemListControl(ByVal oControl As Object, ByVal sText As String) As Long
' Constants for API calls
Const LB_SETTABSTOPS = &H192
Const CB_FINDSTRING = &H14C
Const LB_FINDSTRING = &H18F
Const CB_FINDSTRINGEXACT = &H158

Dim sTemp As String

sTemp = sText
If TypeOf oControl Is ListBox Then
   FindItemListControl = SendMessage(oControl.hWnd, LB_FINDSTRING, -1, ByVal sTemp)
ElseIf TypeOf oControl Is ComboBox Then
   FindItemListControl = SendMessage(oControl.hWnd, CB_FINDSTRING, -1, ByVal sTemp)
End If
    
End Function

'
' LockRecordEdit
'Impovements
Function LockRecordEdit(lHandle As Long) As Boolean
On Local Error Resume Next
'    LockRecordEdit = False                  ' Assume pesimitic
'    rstTemp.LockEdits = True                ' By setting Lockedits to true
'    rstTemp.Edit                            ' An error will occur when edit is
 ObjTour.Edit lHandle
 LockRecordEdit = True
'                                            ' requested if record is already locked!
'        If 0 = Err Then
'                LockRecordEdit = True
'        Else
'        Do While True
'            If vbYes = MsgBox(Err.Description & vbCrLf & LoadResString(gcFailedEditTryAgain), vbYesNo, LoadResString(gcTourVersion)) Then
'                rstTemp.Edit
'                    If 0 = Err Then
'                        LockRecordEdit = True
'                        Exit Function
'                    End If
'            Else
'                Exit Function
'            End If
'        Loop
'        End If
        
End Function


Public Function Read_Tour_Setting_From_Registry(ByVal sSubKey As String, ByVal sKey As String, ByRef sValue As String) As Boolean
' =======================================================
' Purpose: Reads passed registry key
' =======================================================
Dim sRetStr As String

' Optimistic
Read_Tour_Setting_From_Registry = True

' --------------------------------------------
' Read Registry setting for passed parameters
' --------------------------------------------
gbSkipRegErrMsg = True ' Don't show reg error

  sValue = GetRegStringValue(LoadResString(gcRegTourKey) & "\" & sSubKey, sKey)

 ' If failed to read key, return false
 If REG_ERROR = sRetStr Then
    Read_Tour_Setting_From_Registry = False
 End If
 
End Function


Function LoadFormResourceString(frmCurrent As Form) As Boolean
' ------------------------------------------------------
' This function is called from every form loads event
' it passes the form in which the resourse string are to
' be loaed into.
' -------------------------------------------------------
Dim oControl As Control
On Local Error GoTo LoadFromRes_Err
' -------------------
' Track function call
' -------------------
If bDebug Then Handle_Err 0, "LoadFormResourceString-TourFunc"

' Assume Pesimistic
LoadFormResourceString = False
' ------------------------
' First Load Forms Caption
' ------------------------
frmCurrent.Caption = LoadResString(Val(frmCurrent.Tag))
For Each oControl In frmCurrent.Controls
    
    If Val(oControl.Tag) > 0 Then
        oControl.Caption = LoadResString(Val(oControl.Tag))
    End If
Next
LoadFormResourceString = True
Exit Function
LoadFromRes_Err:
    If bDebug Then Handle_Err Err, "LoadFormResourceString-TourFunc"
    Resume Next
End Function

Function CheckDataBaseStructure()

Dim wsDefaultWksp As Workspace
Dim dbsCheck As Database
Dim TourField As Field
Dim fldField As Field
Dim TourTableDef As TableDef
Dim iLoop As Integer, iFound As Integer
Dim sTableBuffer As String      ' Hold all NON MSys table name for database
Dim sFieldBuffer As String      ' Hold all Field names for given table...


On Local Error GoTo Check_Err

Set wsDefaultWksp = DBEngine.Workspaces(0)
' ------------------------------------------------------------
' Define Structure of each TourWin Database.
' Then compare the existing database again define,
' if different from definition, then add table(s) or Field(s)
' ------------------------------------------------------------
' Eve_Tour
TypeEve_Tour.Tablecount = 4
TypeEve_Tour.Tables(1) = "Daily"
TypeEve_Tour.Tables(2) = "Event"
TypeEve_Tour.Tables(3) = "Peak"
TypeEve_Tour.Tables(4) = "Event_Tracker"


' NameTour
' ---------
'Table names
TypeNameTour.Tablecount = 4
TypeNameTour.Tables(1) = "HeartNames"
TypeNameTour.Tables(2) = "Events"
TypeNameTour.Tables(3) = "Levels"
TypeNameTour.Tables(4) = "PeakNames"
'Field Name

TypeNameTour.FieldCount(0) = 37
TypeNameTour.Fields(1) = "Id"
TypeNameTour.Fields(2) = "Peak0"
TypeNameTour.Fields(3) = "Peak10"
TypeNameTour.Fields(4) = "Peak18"
TypeNameTour.Fields(5) = "Color0"
TypeNameTour.Fields(6) = "Color10"


' PeakTour
PeakTour_Tour.Tablecount = 1
PeakTour_Tour.Tables(1) = "Peak"
PeakTour_Tour.FieldCount(0) = 20

' ============================================================'
'                                                             '
' Check Eve_Tour.mdb                                          '
'       Purpose, check both table and field structure adding  '
'       fields or tables when missing.                        '
' ============================================================'
'If "" = Dir$(objMdi.info.Datapath & "\" & gcEve_Tour) Then
'    CreateTourDatabases objMdi.info.Datapath, gcEve_Tour
'Else
'    Set dbsCheck = wsDefaultWksp.OpenDatabase(objMdi.info.Datapath & "\" & gcEve_Tour, True, False)
'' Check Table(s)
'        sTableBuffer = ","
'        For iLoop = 0 To dbsCheck.TableDefs.Count - 1
'            sTableBuffer = sTableBuffer + dbsCheck.TableDefs(iLoop).Name & ","
'        Next iLoop
'        For iLoop = 1 To TypeEve_Tour.Tablecount
'            iFound = InStr(1, sTableBuffer, "," & TypeEve_Tour.Tables(iLoop) & ",", vbTextCompare)
'
'            If 0 = iFound Then
'                        UpGrades.Show
'                        Select Case TypeEve_Tour.Tables(iLoop)
'                                Case gcEve_Tour_Event_Tracker:
'                                    Define_Eve_Tour_Event_Tracker_Plus_Fields objMdi.info.Datapath
'                                Case Else:
'                        End Select
'                        Unload UpGrades
'            End If
'        Next iLoop
'' --------------------------
'' Check Fields with table...
'' --------------------------
'    ' ----------------------------------------------
'    ' Loop through sBuffer and make sure all fields
'    ' exist for Event_Tracker table...
'    ' ----------------------------------------------
'    sTableBuffer = ","
'    Set TourTableDef = dbsCheck.TableDefs("Peak")
'    For Each TourField In TourTableDef.Fields '
'            sTableBuffer = sTableBuffer & TourField.Name & ","
'    Next
'        iFound = InStr(1, sTableBuffer, ",Event_ID,", vbTextCompare)
'            If 0 = iFound Then
'                Set TourTableDef = dbsCheck.TableDefs("Peak")
'                    With TourTableDef
'                        .Fields.Append .CreateField("Event_ID", dbInteger)
'                    End With
'            End If
'End If
'
'' -------------------
'' Check NameTour.mdb
'' -------------------
'If "" = Dir$(objMdi.info.Datapath & "\" & gcNameTour) Then
'    CreateTourDatabases objMdi.info.Datapath, gcNameTour
'Else
'    If Not dbsCheck Is Nothing Then Set dbsCheck = Nothing
'    Set dbsCheck = wsDefaultWksp.OpenDatabase(objMdi.info.Datapath & "\" & gcNameTour, False, False)
'
'
'' Check Table(s)
'
'        sTableBuffer = ","
'        For iLoop = 0 To dbsCheck.TableDefs.Count - 1
'            sTableBuffer = sTableBuffer + dbsCheck.TableDefs(iLoop).Name & ","
'        Next iLoop
'        For iLoop = 1 To TypeNameTour.Tablecount
'            iFound = InStr(1, sTableBuffer, "," & TypeNameTour.Tables(iLoop) & ",", vbTextCompare)
'
'            If 0 = iFound Then
'
'                        Select Case TypeNameTour.Tables(iLoop)
'                                Case gcEve_Tour_Event_Tracker:
'                                    Define_Eve_Tour_Event_Tracker_Plus_Fields objMdi.info.Datapath
'                                Case gcNameTour_PeakNames:
'                                    Define_NameTour_PeakNames_Plus_Fields objMdi.info.Datapath
'                                Case Else:
'                        End Select
'
'            End If
'        Next iLoop
'
'End If
'' Check Fields
'' Determine if Peak10 field exists?
'' ---------------------------------
'sFieldBuffer = ","
'For Each fldField In dbsCheck.TableDefs("PeakNames").Fields
'    sFieldBuffer = sFieldBuffer & "," & fldField.Name
'Next fldField
'
'' Check for Peak10
'iFound = InStr(1, sFieldBuffer, "," & TypeNameTour.Fields(3) & ",", vbTextCompare)
'If 0 = iFound Then
'        ' Not Found, add field to table
'        For iLoop = 10 To 18
'            Set TourTableDef = dbsCheck.TableDefs("PeakNames")
'                With TourTableDef
'                    sFieldBuffer = "Peak" & Format$(iLoop, "##")
'
'                    .Fields.Append .CreateField(sFieldBuffer, dbText, 100)
'                    .Fields(sFieldBuffer).AllowZeroLength = True
'
'                    sFieldBuffer = "Color" & Format$(iLoop, "##")
'                    .Fields.Append .CreateField(sFieldBuffer, dbText, 15)
'                    .Fields(sFieldBuffer).AllowZeroLength = True
'                End With
'
'        Next iLoop
'End If
'' -------------------
'' Check PeakTour.mdb
'' -------------------
'If "" = Dir$(objMdi.info.Datapath & "\" & gcPeakTour) Then
'    CreateTourDatabases objMdi.info.Datapath, gcNameTour
'Else
'    Set dbsCheck = wsDefaultWksp.OpenDatabase(objMdi.info.Datapath & "\" & gcPeakTour, True, False)
'
'
'' Check Table(s)
'
'        sTableBuffer = ","
'        For iLoop = 0 To dbsCheck.TableDefs.Count - 1
'            sTableBuffer = sTableBuffer + dbsCheck.TableDefs(iLoop).Name & ","
'        Next iLoop
'        For iLoop = 1 To TypeNameTour.Tablecount
'                         'TypeNameTour
'            iFound = InStr(1, sTableBuffer, "," & TypeNameTour.Tables(iLoop) & ",", vbTextCompare)
'
'            If 0 = iFound Then
'
'                        Select Case TypeNameTour.Tables(iLoop)
'                                Case gcEve_Tour_Event_Tracker:
'                                    Define_Eve_Tour_Event_Tracker_Plus_Fields objMdi.info.Datapath
'                                Case Else:
'                        End Select
'
'            End If
'        Next iLoop
'
'End If



If Not dbsCheck Is Nothing Then Set dbsCheck = Nothing
If Not wsDefaultWksp Is Nothing Then Set wsDefaultWksp = Nothing
Exit Function
Check_Err:

    If bDebug Then Handle_Err Err, "CheckDatabaseStructure-TourFunc"
    Resume Next
End Function

Sub Define_Form_menu(ByVal sFormName As String, ByVal LoadType As String)
On Local Error GoTo Define_Form_menu_Err

If LoadType = Unloadmnu Then

            ' Standard MDI menu
            MDI!MdiEximnu.Caption = MdiFrm_Exit
            MDI!MdiOptmnu.Caption = MdiFrm_Option
            MDI!MdiNewmnu.Caption = MdiFrm_Newmnu
            MDI!MdiSavmnu.Caption = MdiFrm_Savmnu
            MDI!MdiDelmnu.Caption = MdiFrm_Delmnu
            MDI!MdiOptmnu.Enabled = True
            MDI!MdiNewmnu.Enabled = True
            MDI!MdiSavmnu.Enabled = False
            MDI!MdiDelmnu.Enabled = False
            MDI!MdiPrimnu.Enabled = False

    Select Case sFormName
        Case "DailyFrm":
            MDI!MdiDiamnu.Enabled = True
        Case "HrtZone":
            MDI!MdiHeamnu.Enabled = True
        Case "EventFrm":
            MDI!MdiTypmnu.Enabled = True
        Case "MdiOpt":
            MDI!MdiOptmnu.Enabled = True
    End Select
Else
    Select Case sFormName
        'Define Daily Form menus
        Case "DailyFrm":
            MDI!MdiEximnu.Caption = DailyFrm_Exit
            MDI!MdiOptmnu.Caption = DailyFrm_Option
            MDI!MdiNewmnu.Caption = DailyFrm_Newmnu
            MDI!MdiSavmnu.Caption = DailyFrm_Savmnu
            MDI!MdiDelmnu.Caption = DailyFrm_Delmnu
            MDI!MdiOptmnu.Enabled = True
            MDI!MdiDiamnu.Enabled = False
            MDI!MdiNewmnu.Enabled = True
            MDI!MdiSavmnu.Enabled = True
            MDI!MdiDelmnu.Enabled = True
            MDI!MdiPrimnu.Enabled = True
        'HrtZone Form menus
        Case "HrtZone":
            MDI!MdiEximnu.Caption = HrtZone_Exit
            MDI!MdiOptmnu.Caption = HrtZone_Option
            MDI!MdiNewmnu.Caption = gcHrtZone_Newmnu
            MDI!MdiSavmnu.Caption = gcHrtZone_Savmnu
            MDI!MdiDelmnu.Caption = gcHrtZone_Delmnu
            MDI!MdiOptmnu.Enabled = False
            MDI!MdiNewmnu.Enabled = False
            MDI!MdiSavmnu.Enabled = True
            MDI!MdiDelmnu.Enabled = False
            MDI!MdiPrimnu.Enabled = True
            MDI!MdiHeamnu.Enabled = False
        Case "EventFrm":
            MDI!MdiEximnu.Caption = gcEvent_Exit
            MDI!MdiOptmnu.Caption = gcEvent_Option
            MDI!MdiNewmnu.Caption = gcEvent_Newmnu
            MDI!MdiSavmnu.Caption = gcEvent_Savmnu
            MDI!MdiDelmnu.Caption = gcEvent_Delmnu
            MDI!MdiOptmnu.Enabled = False
            MDI!MdiNewmnu.Enabled = False
            MDI!MdiSavmnu.Enabled = True
            MDI!MdiDelmnu.Enabled = False
            MDI!MdiPrimnu.Enabled = False
            MDI!MdiTypmnu.Enabled = False
        Case "MdiOpt":
            MDI!MdiEximnu.Caption = gcMdO_Exit
            MDI!MdiOptmnu.Caption = gcMdO_Option
            MDI!MdiNewmnu.Caption = gcMdO_Newmnu
            MDI!MdiSavmnu.Caption = gcMdO_Savmnu
            MDI!MdiDelmnu.Caption = gcMdO_Delmnu
            MDI!MdiOptmnu.Enabled = False
            MDI!MdiNewmnu.Enabled = False
            MDI!MdiSavmnu.Enabled = True
            MDI!MdiDelmnu.Enabled = False
            MDI!MdiPrimnu.Enabled = False
        Case "ConcFrm":
            MDI!MdiEximnu.Caption = gcCONCONI_EXIT
            MDI!MdiOptmnu.Caption = gcCONCONI_OPTION
            MDI!MdiNewmnu.Caption = gcCONCONI_NEWMNU
            MDI!MdiSavmnu.Caption = gcCONCONI_SAVMNU
            MDI!MdiDelmnu.Caption = gcCONCONI_DELMNU
            MDI!MdiOptmnu.Enabled = True
            MDI!MdiNewmnu.Enabled = True
            MDI!MdiSavmnu.Enabled = True
            MDI!MdiDelmnu.Enabled = True
            MDI!MdiPrimnu.Enabled = True
    End Select
End If
Exit Sub
Define_Form_menu_Err:
    If bDebug Then Handle_Err Err, "Define_Form_menu-TourFunc"
    Resume Next
End Sub



Function Define_Activities_Fields(CreatePath As String) As Boolean
' -------------------------------------------------------------------
' Define_Activity_Fields -  The function creates the Activity Table
'                           which stores Peak, Event and Heart Names
'
' Return Value:     True if successful.
' Revision History:
'   Jan 23, 2001    - Created function
' -------------------------------------------------------------------
On Local Error GoTo Activity_Fields_Err
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index
Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
' Fill in new database.
ProgressBar "Creating Activity Table... ", -1, 2, -1

' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef(gcActivitiesTable)
' Append Id Field
Set TourField = TourTableDef.CreateField(gcActive_ID, dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
    
' Append Type Field
Set TourField = TourTableDef.CreateField(gcActive_Type, dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
    
' Append Type Field
Set TourField = TourTableDef.CreateField(gcActive_Pos, dbLong)
    TourField.Required = True
    TourField.DefaultValue = 0
    TourTableDef.Fields.Append TourField
    
' Append Description Field
Set TourField = TourTableDef.CreateField(gcActive_Des, dbText, 100)
    TourField.DefaultValue = ""
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
    
' Append Colour Field
Set TourField = TourTableDef.CreateField(gcActive_Colour, dbText, 15)
    TourField.DefaultValue = "16777215" ' White BackGround
    TourField.AllowZeroLength = False
    TourTableDef.Fields.Append TourField
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
ProgressBar "Creating Activity Table...", -1, 7, -1

' ---------------
' Append Index
' ---------------
' Create primary index.
Set TourIndex = TourTableDef.CreateIndex("ActivityIndex")
    TourIndex.Primary = True
    TourIndex.Required = True
    TourIndex.Unique = True
Set TourField = TourTableDef.CreateField(gcActive_ID)
    TourIndex.Fields.Append TourField
Set TourField = TourTableDef.CreateField(gcActive_Type)
    TourIndex.Fields.Append TourField
Set TourField = TourTableDef.CreateField(gcActive_Pos)
    TourIndex.Fields.Append TourField
    
ProgressBar "Creating Activity Table...", -1, 7, -1
    
' ------------------------------------------------------------
' Save Index definition by appending it to Indexes collection.
' ------------------------------------------------------------
    TourTableDef.Indexes.Append TourIndex
    
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
Define_Activities_Fields = True
Exit Function
Activity_Fields_Err:
    If bDebug Then Handle_Err Err, "Define_Activity_Fields-TourFunc"
    Resume Next
    
End Function


Public Function Historical_Report()

Dim sDataSource     As String

On Local Error GoTo Rep_Err

If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing
    
' Make sure DSN is pointing to current database
ObjTour.RegisterUpdateDSN objMdi.info.Datapath & gcTour_Win
    
sDataSource = objMdi.info.Datapath & gcTour_Win
        
' Create Crystal Objects
Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Historical.rpt")

' Specify records to retrieve
CryReport.RecordSelectionFormula = "{Dai.Id} = " & objMdi.info.ID & " And {Dai.Type} = " & gcDAI_HISTORICAL

CryReport.Preview "Historical Events Report. " & LoadResString(gcTourVersion)

Exit Function
Rep_Err:
If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing
If bDebug Then Handle_Err Err, "Historical Report"
MsgBox Error$(Err)
    Select Case Err
        Case 13 Or 20533:      'File not found
        MsgBox Error$(Err)
        Case 20515:
                    MsgBox Error$(Err)
        Case Else
        'If bDebug Then Handle_Err Err, "MdiTotmnu-MDI"
        Resume Next
    End Select


End Function

Public Function Conconi_Report()
Dim sDataSource As String
Dim CryParameters As CRPEAuto.ParameterFieldDefinitions
Dim CryParameter As CRPEAuto.ParameterFieldDefinition


On Local Error GoTo Rep_Err

If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing
    
sDataSource = objMdi.info.Datapath & gcTour_Win
        
' Make sure DSN is pointing to current database
ObjTour.RegisterUpdateDSN objMdi.info.Datapath & gcTour_Win
        
' Create Crystal Objects
Set gCrystalReport = CreateObject("Crystal.CRPE.Application")
Set CryReport = gCrystalReport.OpenReport(App.Path & "\Conconi.rpt")

' Update Report Parameter
Set CryParameters = CryReport.ParameterFields

For Each CryParameter In CryParameters
    Select Case CryParameter.Name
        Case "{?ID}":
            CryParameter.SetCurrentValue objMdi.info.ID
    End Select
Next

' Release child Crystal Objects
If Not CryParameter Is Nothing Then Set CryParameter = Nothing
If Not CryParameters Is Nothing Then Set CryParameters = Nothing

CryReport.Preview "Conconi Report. " & LoadResString(gcTourVersion)

Exit Function
Rep_Err:
If Not gCrystalReport Is Nothing Then Set gCrystalReport = Nothing
If bDebug Then Handle_Err Err, "Conconi Report"
MsgBox Error$(Err)
    Select Case Err
        Case 13 Or 20533:      'File not found
        MsgBox Error$(Err)
        Case 20515:
                    MsgBox Error$(Err)
        Case Else
        'If bDebug Then Handle_Err Err, "MdiTotmnu-MDI"
        Resume Next
    End Select


End Function


'---------------------------------------------------------------------------------------
' PROCEDURE : SyntaxSafe
' DATE      : 6/4/03 21:31
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function SyntaxSafe(ByVal sValue As String) As String
On Local Error GoTo SyntaxSafe_Error
'Declare local variables

    SyntaxSafe = Replace(sValue, "'", "''", , , vbTextCompare)

On Error GoTo 0
Exit Function

SyntaxSafe_Error:
    If bDebug Then Handle_Err Err, "SyntaxSafe-TourFunc"
    Resume Next

End Function

