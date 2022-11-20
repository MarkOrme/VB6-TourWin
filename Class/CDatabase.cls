VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------
' CDatabase:    Connection to datasource - DataLayer
' Author:       Mark Ormesher
' Description:  The CDatabase class establish connection to database
'               using the dbs object. In turn CRecord class using the CDatabase
'               object to create recordset object.
' -----------------------------------------------------------------------------
Option Explicit
Private m_AccessDB      As Database
Private m_SQLDB         As rdoConnection
Private mTourVar        As cTourVar
Private m_sPath         As String
Private m_bSQLDatabase  As Boolean
Private m_MDI           As cMDI
Private m_MDIVar        As cMDIVar
Private oLogin          As CLogin
Const CTABLECOUNT = 12


' -----------------------------------------------------------------------------
' DBOpen - This function is only called via the cTour_DB object
' and is used to narrow data access to one function
' Return Values:
'               0 - Database opened successfully.
'               1 - Created new database and then opened successfully.
'               2 - Failed to open database.
' -----------------------------------------------------------------------------

Public Function DBOpen(ByVal sDatapath As String) As Long

On Local Error GoTo DBOpen_Err

Const FOUNDANDOPENED = 0
Const CREATEDANDOPENED = 1
Const NOTFOUNDNORCREATED = 2

Dim lRt     As Long


If bDebug Then Handle_Err 0, "DBOpen-cDatabase - OpenDatabase(" & m_MDI.info.Datapath & gcTour_Win


lRt = Check_Data_Exist(sDatapath)

DBOpen = lRt

If 2 = lRt Then
    Exit Function
End If

m_bSQLDatabase = False

If left(UCase$(sDatapath), 4) = "DSN:" Then
    m_bSQLDatabase = True
End If

    ' 0 or 1, does not matter just open
    Select Case m_bSQLDatabase
        Case False: ' Access Database
        
            Set m_AccessDB = DBEngine.Workspaces(0).OpenDatabase(sDatapath & "\" & gcTour_Win, False, False, ";pwd=" & gcTOURWIN_PASSWORD)
            
            If m_AccessDB Is Nothing Then
                DBOpen = NOTFOUNDNORCREATED
            End If
            
            If CREATEDANDOPENED = lRt Then
                CreateRelationships ' Append relationships
            End If
            
        Case True: ' SQL Database

            Set rdoDb1 = _
                rdoEnv.OpenConnection(m_MDI.info.Datapath, _
                rdDriverNoPrompt, False, "UID=TourAdmin;PWD=tourwin")
                
            If rdoDb1 Is Nothing Then
                DBOpen = NOTFOUNDNORCREATED
            End If
                
    End Select
    

Exit Function
DBOpen_Err:
DBOpen = NOTFOUNDNORCREATED

If bDebug Then
    Handle_Err Err, "DBOpen-cTourInfo"
    MsgBox Err.Description
End If
    Resume Next
End Function

Public Function DBClose() As Boolean

If Not m_bSQLDatabase Then
    
    If Not m_AccessDB Is Nothing Then
        m_AccessDB.Close
        Set m_AccessDB = Nothing
    End If
Else
    If Not m_SQLDB Is Nothing Then
        m_SQLDB.Close
        Set m_SQLDB = Nothing
    End If
End If

End Function
Public Function DBExecute(ByVal sQuery As String) As Boolean

DBExecute = True
If bSQLDatabase Then

    Call m_SQLDB.Execute(sQuery)
    
Else
    Call m_AccessDB.Execute(sQuery)

End If

End Function


Public Function CreateExportDB(ByVal sPath_Name As String) As Boolean
'
' CreateExportDB
' ---------------------------
' Add field(s) to MyTableDef.
' ---------------------------
On Local Error GoTo CreateExportDB_Err

Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database, TourTableDef As TableDef, TourField As Field
Dim TourIndex As Index


Set DefaultWorkspace = DBEngine.Workspaces(0)
' Create new, Decrypted database.
Set TourDatabase = DefaultWorkspace.CreateDatabase(sPath_Name, dbLangGeneral, dbDecrypt)
    TourDatabase.NewPassword "", gcEXPORT_PASSWORD
    
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef(gcEXPORT_SCHEDULE_TABLE)

' Append Type_ID Field
Set TourField = TourTableDef.CreateField(gcEXPORT_SCHEDULE_TABLE, dbText, 255)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
    
' Append Type_ID Field
Set TourField = TourTableDef.CreateField(gcEXPORT_TYPE_ID_FIELD, dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
    
' Append Name Field
Set TourField = TourTableDef.CreateField(gcEXPORT_NAME_FIELD, dbText, 35)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append PassWord Field
Set TourField = TourTableDef.CreateField(gcEXPORT_CYCLE_L_FIELD, dbLong)
    'TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
    
' Create new TableDef.
Set TourTableDef = TourDatabase.CreateTableDef(gcEXPORT_ACTIVITY_TABLE)
' Append Type_ID Field
Set TourField = TourTableDef.CreateField(gcEXPORT_TYPE_ID_FIELD, dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
' Append Name Field
Set TourField = TourTableDef.CreateField(gcEXPORT_POS_FIELD, dbLong)
    TourField.Required = True
    TourTableDef.Fields.Append TourField
    
' Append Description Field
Set TourField = TourTableDef.CreateField(gcEXPORT_DESCRIPTION_FIELD, dbText, 100)
    TourField.AllowZeroLength = True
    TourTableDef.Fields.Append TourField
    
' Append Colour Field
Set TourField = TourTableDef.CreateField(gcEXPORT_COLOUR_FIELD, dbText, 15)
    TourField.AllowZeroLength = False
    TourTableDef.Fields.Append TourField
    
' -----------------------------------------------------------------
' Save TableDef definition by appending it to TableDefs collection.
' -----------------------------------------------------------------
TourDatabase.TableDefs.Append TourTableDef
    
' ---------------------------------
' Close newly created TourDatabase.
' ---------------------------------
TourDatabase.Close
CreateExportDB = True
Exit Function
CreateExportDB_Err:
    If Err = 3204 Then      'DataBAse already exist
        Err = 0
        CreateExportDB = False
        Exit Function
    End If
    MsgBox Err.Description
    If bDebug Then Handle_Err Err, "CreateExportDB_Err-cTourInfo"
    Resume Next

End Function

Private Function Update_Table_Name(ByRef sSQL As String) As Boolean

Dim sTempString As String
Dim iFrom As Integer
Dim iEndOfTableName As String

    ' Find 'FROM'
    iFrom = InStr(1, UCase(sSQL), " FROM ", vbTextCompare)
    iEndOfTableName = InStr(iFrom + 6, sSQL, " ", vbTextCompare)
    
    sTempString = Mid$(sSQL, 1, iFrom + 5) & "[TourAdmin." & Mid$(sSQL, iFrom + 6, iEndOfTableName - (iFrom + 6)) & "]" & Mid$(sSQL, iEndOfTableName)
    sSQL = sTempString
    
End Function

Public Function CreateQueryDefObject(ByVal sName As String, ByVal sSQL As String) As String

On Local Error GoTo Create_Error
Dim qTempQueryDef As QueryDef

Set qTempQueryDef = dbs1.CreateQueryDef(sName)

' Assigning SQL value creates Def right now
qTempQueryDef.SQL = sSQL
On Local Error GoTo 0

Exit Function
Create_Error:
Err.Clear

End Function


Private Sub Class_Initialize()
Set m_MDI = New cMDI
Set m_MDIVar = New cMDIVar
Set m_MDI.info = m_MDIVar

End Sub

Private Sub Class_Terminate()
    If Not m_MDI.info Is Nothing Then Set m_MDI.info = Nothing
    If Not m_MDIVar Is Nothing Then Set m_MDIVar = Nothing
    If Not m_MDI Is Nothing Then Set m_MDI = Nothing
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : Check_Data_Exist
' DATE      : 6/6/03 16:10
' Author    : Mark Ormesher
' Purpose   : CheckType
'---------------------------------------------------------------------------------------
Private Function Check_Data_Exist(DPath As String) As Long
On Local Error GoTo Check_Err

Dim lRt     As Long
Const FOUNDANDOPENED = 0
Const CREATEDANDOPENED = 1
Const NOTFOUNDNORCREATED = 2

Do While "" = Dir(DPath & "\" & gcTour_Win)

'*********************************************************
' If database is not found, then check eCheck
' to determine if database needs to be created after
' installation, if so, do not prompt user, just create...
'*********************************************************

If eCheckDatabase.CreateDBaseAfterInstall = m_MDIVar.eCheckDBType Then

    lRt = CREATEDANDOPENED
    
Else

    'Prompt user to find or create...
    lRt = CLng(InputBox("Database not found! Would you like to search for database or create new? 0 for find, 1 for create, cancel to end program", "Database choose", "0"))
    
End If
 
 
 
Select Case lRt

       Case FOUNDANDOPENED: 'Actually, find and open
            DPath = GetPath("MDB files (*.mdb)" & Chr$(0) & gcTour_Win, gcTour_Win, -1)
            m_MDI.info.Datapath = DPath
            Check_Data_Exist = FOUNDANDOPENED
            
            
       Case CREATEDANDOPENED:
            CreateTourDatabases DPath, gcTour_Win
            m_MDI.info.Datapath = DPath
            m_MDI.info.WelcomeWizard = True
            Check_Data_Exist = CREATEDANDOPENED
                       
       Case Else
            Check_Data_Exist = NOTFOUNDNORCREATED
            Exit Do
End Select
       
Loop


On Error GoTo 0
Exit Function
Check_Err:
    If bDebug Then Handle_Err Err, "Check_Data_Exist-TourFunc"
    Resume Next
End Function

Private Function GetPath(FileType As String, FileName As String, Critical As Integer) As String
On Local Error GoTo GetPath_Err

Dim LenSubtract As Integer, Msg As String
Dim Resp As Integer

' Prompt user to search for data files
Msg = "TourWin cannot find data files!" & vbLf & vbLf
Msg = Msg & "Data files must be keeped together in one" & vbLf
Msg = Msg & "directory, but maybe located separately " & vbLf
Msg = Msg & "from the Tourwin.exe program." & vbLf & vbLf
Msg = Msg & "Options:" & vbLf
Msg = Msg & "   Ok button to search for data path." & vbLf
Msg = Msg & "   Cancel button to terminate program."
Resp = MsgBox(Msg, vbOKCancel + vbCritical, "Searching for 'TourWin' data files.")
If Resp = vbCancel And Critical = -1 Then
        End
ElseIf Resp = vbCancel And Critical = 0 Then
        GetPath = "Canceled"
        Exit Function
End If

LenSubtract = Len(FileName) + 1 ' The plus +1 removes the '\'

Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    
    FileDialog.sFilter = FileType
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    FileDialog.sDlgTitle = "Open"
    sOpen = ShowOpen(MDI.hWnd)
    
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        GetPath = Mid$(FileDialog.sFile, 1, Len(FileDialog.sFile) - LenSubtract)
    Else
        GetPath = ""
    End If

Exit Function
GetPath_Err:
    'User pressed Cancel Button
    GetPath = ""
End Function


'---------------------------------------------------------------------------------------
' PROCEDURE : OpenDBAndLogin
' DATE      : 6/27/04 10:57
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function OpenDBAndLogin() As cMDI

On Local Error GoTo OpenDBAndLogin_Error
'Declare local variables
Const FOUNDANDOPENED = 0
Const CREATEDANDOPENED = 1
Const NOTFOUNDNORCREATED = 2

Dim sRetVal As String

Set oLogin = New CLogin
' Get info from registry

GetLoginInformationFromRegistry

oLogin.Display

    Select Case oLogin.LoginSuccessful
           Case FOUNDANDOPENED:

                With m_MDI.info
                        .Name = oLogin.UserName
                        .Password = oLogin.Password
                        .Datapath = oLogin.Datapath
                End With
    
                m_MDI.LoadUserSettings
                Set OpenDBAndLogin = m_MDI
                
           Case CREATEDANDOPENED:
                 
                With m_MDI.info
                        .Name = oLogin.UserName
                        .Password = oLogin.Password
                        .Datapath = oLogin.Datapath
                End With
                m_MDI.AddUser
                m_MDI.LoadUserSettings
                Set OpenDBAndLogin = m_MDI
           
           Case NOTFOUNDNORCREATED:
                Set OpenDBAndLogin = Nothing
End Select


On Error GoTo 0
Exit Function

OpenDBAndLogin_Error:
    If bDebug Then Handle_Err Err, "OpenDBAndLogin-CDatabase"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : OpenRecordSet
' DATE      : 6/27/04 11:34
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function OpenRecordSet(ByVal Name As String, lType As Long) As Recordset

On Local Error GoTo OpenRecordSet_Error

    Set OpenRecordSet = m_AccessDB.OpenRecordSet(Name, lType)

On Error GoTo 0
Exit Function

OpenRecordSet_Error:
    If bDebug Then Handle_Err Err, "OpenRecordSet-CDatabase"
    Resume Next

End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : GetLoginInformationFromRegistry
' DATE      : 6/30/04 19:17
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub GetLoginInformationFromRegistry()

On Local Error GoTo GetLoginInformationFromRegistry_Error

Dim sRetVal     As String


gbSkipRegErrMsg = True  ' Don't show reg error
sRetVal = Trim$(GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcTourUserName)))
If REG_ERROR <> sRetVal Then
    oLogin.UserName = sRetVal
End If


gbSkipRegErrMsg = True  ' Don't show reg error
sRetVal = Trim$(GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcTourLastDB)))
If REG_ERROR <> sRetVal Then

    oLogin.Datapath = sRetVal
'    m_MDI.info.eCheckDBType = CheckIfExist_LastUserFound
End If

On Error GoTo 0
Exit Sub

GetLoginInformationFromRegistry_Error:
    If bDebug Then Handle_Err Err, "GetLoginInformationFromRegistry-CDatabase"
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : CreateRelationships
' DATE      : 7/3/04 12:31
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function CreateRelationships() As Boolean

On Local Error GoTo CreateRelationships_Error
'Declare local variables
Dim lLoop       As Long
Dim rel         As Relation
Dim fld         As Field
Dim fld2        As Field
Dim sTemp       As String


If m_AccessDB Is Nothing Then
    Exit Function
    CreateRelationships = False
End If


' *************************************************
' * Create Cascade Delete From User Table
' * To all other tables...
' * ------------------------------------------------
For lLoop = 1 To CTABLECOUNT - 1
    
    Select Case lLoop
           Case 0:
                sTemp = gcUserTour_UserTbl
           Case 1:
                sTemp = gcDai_Tour_Dai
           Case 2:
                sTemp = gcDaily_Table
           Case 3:
                sTemp = gcPeak_Table
           Case 4:
                sTemp = gcPeaks_Table
           Case 5:
                sTemp = gcActivitiesTable
           Case 6:
                sTemp = gcContTour_Contacts
           Case 7:
                sTemp = gcUserTour_ContactOpt
           Case 8:
                sTemp = gcLEVELS_TABLE
           Case 9:
                sTemp = gcEve_Tour_Event_Tracker
           Case 10:
                sTemp = gcEve_Tour_Event
           Case 11:
                sTemp = gcData
    End Select
    Set rel = m_AccessDB.CreateRelation("User_To_" & sTemp)
    With rel
        .Table = gcUserTour_UserTbl
        .ForeignTable = sTemp
    End With

    ' ------------------------------------------------
    ' Define fields for relationship
    ' ------------------------------------------------
    ' Define ID to ID
    Set fld = rel.CreateField(gcID)
        fld.ForeignName = gcID
            
    ' ------------------------------------------------
    ' Define attributes and append to database object
    ' ------------------------------------------------
    rel.Attributes = dbRelationDeleteCascade
    
    rel.Fields.Append fld
    m_AccessDB.Relations.Append rel
    Set rel = Nothing
Next lLoop


On Error GoTo 0
Exit Function

CreateRelationships_Error:
    If bDebug Then Handle_Err Err, "CreateRelationships-CDatabase"
    Resume Next


End Function
