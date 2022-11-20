VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cMDI
' DateTime  : 5/1/04 11:07
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
' -------------------------------------------
' This class uses cUserVar's public variables
' -------------------------------------------
' Constants for iBitFlag - CMDIVar
' |||| |||| |||| ||-- Conconi - Plot Points   = "BitField" And 2
' |||| |||| |||| |--- Conconi - Draw Legend   = "BitField" And 4
' |||| |||| |||| ---- Conconi - Draw AT       = "BitField" And 8
' |||| |||| |||------ Conconi - Load Settings = "BitField" And 16

Public Enum BitFlags
      Pos_0 = 0
      Conconi_PlotPoints = 1
      Conconi_DrawLegend = 2
      Conconi_DrawAT = 3
      Conconi_LoadSettings = 4
      User_Security = 5
      User_ShowMeta = 6
      User_Load1 = 7
      User_Load2 = 8
      User_Load3 = 9
      Pos_10 = 10
      Pos_11 = 11
      Pos_12 = 12
      Pos_13 = 13
      Pos_14 = 14
      Daily_RestHeart = 0
      Daily_DayType = 1
      Daily_Weight = 2
      Daily_Sleep = 3
      Daily_Description = 4
      Daily_Metric = 5
      Contact_SaveColumnWidth = 0
End Enum

Private mMDIVar As cMDIVar

Public Property Set info(pMDIVar As cMDIVar)
    Set mMDIVar = pMDIVar
End Property

Public Property Get info() As cMDIVar

On Local Error Resume Next
If mMDIVar Is Nothing Then Exit Property
Set info = mMDIVar
If info.Datapath <> "" Then
    ' Make sure "\" is the last character
    If "\" <> Right$(info.Datapath, 1) Then
        info.Datapath = info.Datapath & "\"
    End If
End If

End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SaveUserSettings
' DATE      : 2/20/03 12:41
' Author    : mor
' Purpose   :
'---------------------------------------------------------------------------------------

Public Function SaveUserSettings() As Long

' Declare Local Variables
Dim SQL As String   ' Query holder
Dim oUser       As CRecord

On Error GoTo SaveUserSettings_Error

Set oUser = New CRecord

'
SQL = "SELECT * FROM " & gcUserTour_UserTbl & " WHERE ID = " & objMdi.info.ID

With oUser
    .RstSQL SQL

    If .NoMatch() Then
    
    Else
      .Edit

    '
    '    'Both UserOptions & DailyOpt properties are
    '    info.UserOptions.SetValue (CLng(.GetField(gcUserTour_UserTbl_BitField)))
    '    info.DailyOpt.SetValue (.GetField(gcUserTour_UserTbl_DailyOptions))
    '    info.Password = .GetField(gcUserTour_UserTbl_PassWord)
    '
    '    With info.UserOptions
    '        info.Load = (4 * .GetValue(BitFlags.User_Load3)) + _
    '                    (2 * .GetValue(BitFlags.User_Load2)) + _
    '                    (.GetValue(BitFlags.User_Load1))
    '    End With
    '
    '    info.Security = info.UserOptions.GetBool(BitFlags.User_Security)
    '    info.ShowMeta = info.UserOptions.GetBool(BitFlags.User_ShowMeta)
    '
    '    info.MetaFile = .GetField(gcUserTour_UserTbl_Metafile)
        ' Move ObjMdi Values to RecordSet...

        .SetField "PassWord", mMDIVar.Password
        .SetField "MetaFile", IIf(IsNull(mMDIVar.MetaFile), " ", mMDIVar.MetaFile)
        .SetField "ShowMeta", mMDIVar.ShowMeta
        .SetField "Security", mMDIVar.Security
        .SetField "BitField", mMDIVar.UserOptions.GetValue
        .SetField gcUserTour_UserTbl_DailyOptions, mMDIVar.DailyOpt.GetValue
    
        ' Save
      .Update
    End If
End With

Set oUser = Nothing

SaveUserSettings = lResult
On Error GoTo 0
Exit Function

SaveUserSettings_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveUserSettings of Class Module cMDI"

End Function

' When this function is called, it raises the
' BackGroundImage change event...
Public Sub MetaImage(ByVal sMetaFile As String, ByVal bShow As Boolean)

    Me.info.MetaFile = sMetaFile
    Me.info.ShowMeta = bShow
    
End Sub


''---------------------------------------------------------------------------------------
'' PROCEDURE : AttemptUserLogin
'' DATE      : 6/26/04 20:14
'' Author    : Mark Ormesher
'' Purpose   :
''---------------------------------------------------------------------------------------
'Public Function Login() As Long
'
'On Local Error GoTo Login_Error
'
''Declare local variables
'Const OK = 0
'Const BADNAME = 1
'Const BADPASSWORD = 2
'
'Dim Search      As String
'Dim DirFound    As String
'Dim SQL         As String
'Dim lSearcher   As Long
'
'
'' User is allow three attempts
'' before program automatically ends..
'
'' ----------------------------
'' Load Searcher object and
'' Determine if any users exist
'' ----------------------------
'SQL = "SELECT * FROM " & gcUserTour_UserTbl
'
'ObjTour.RstSQL lSearcher, SQL
'
'ObjTour.DBMoveLast lSearcher
'
'Search = "Name = '" & Trim$(mMDIVar.Name) & "'"
'ObjTour.DBMoveFirst lSearcher
'
'ObjTour.DBFindFirst Search, lSearcher
'
'' ------------------------------------
'' If no match to user name.
'' ------------------------------------
'If ObjTour.NoMatch(lSearcher) Then
'
'    MsgBox LoadResString(gcPassFrmUserName), vbOKOnly + vbExclamation, LoadResString(gcTourVersion)
'    Login = BADNAME
'    Exit Function
'
'End If
'
''Is security on?
'If ObjTour.DBGetField("PassWord", lSearcher) <> mMDIVar.Password Then
'
'    MsgBox "Incorrect user password.", vbOKOnly + vbExclamation, LoadResString(gcTourVersion)
'    Login = BADPASSWORD
'    Exit Function
'
'End If
'
'Login = OK
'LoadUserSettings lSearcher
'ObjTour.FreeHandle lSearcher
'
'On Error GoTo 0
'Exit Function

'    ' Check if any record exist, if not, then
'    ' new install and user should be added...
'    If 0 = ObjTour.RstRecordCount(iSearcherDB) Then
'
'        With ObjTour
'
'            .AddNew (iSearcherDB)
'           Call .DBSetField(gcUserTour_UserTbl_Name, Trim$(mMDIVar.Name), iSearcherDB)
'           Call .DBSetField(gcUserTour_UserTbl_DataPath, mMDIVar.dataPath, iSearcherDB)
'           Call .DBSetField(gcUserTour_UserTbl_Load, 0, iSearcherDB)
'            .Update (iSearcherDB)
'
'        End With
'
'    ' Move cursor to newly created record and load values
'        SQL = "SELECT * FROM " & gcUserTour_UserTbl & " WHERE Name = '" & Trim$(mMDIVar.Name) & "'"
'
'        ObjTour.RstSQL iSearcherDB, SQL
'        ObjTour.DBMoveFirst iSearcherDB
'
'        'Load all run-time values
''        Check_User
'
'        ' Now allow flow to continue
''        GetContactOpt
'        Unload Me 'PassFrm.Hide
'    Else
'        Beep
'        MsgBox LoadResString(gcPassFrmUserName), vbOKOnly + vbExclamation, LoadResString(gcTourVersion)
'        PasNamTxt.SetFocus
'        PasOkCmd.Enabled = True
'    End If
'Else
'' -------------------
'' set all user values
'' -------------------
'
'If Check_User Then
'' Check datapath
'    DirFound = Check_Data_Exist(objMdi.info.dataPath)
'        If DirFound <> ObjTour.DBGetField("DataPath", iSearcherDB) Then
'            dataPath = DirFound
'
'            ObjTour.Edit iSearcherDB           'setUser.Edit
'            ObjTour.DBSetField "DataPath", DirFound, iSearcherDB
'            ObjTour.Update iSearcherDB        'setUser.Update
'        End If
'
'    GetContactOpt
'    Unload Me 'PassFrm.Hide
'
'End If ' End Check_User if
''Unload frmSplash
'End If ' End NoMatch if
'
'    Login = bResult
'
'On Error GoTo 0
'Exit Function

'Login_Error:
'    If bDebug Then Handle_Err Err, "Login-cMDI"
'    Resume Next
'
'
'End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : LoadUserSettings
' DATE      : 6/28/04 19:38
' Author    : Mark Ormesher
' Purpose   : Loads user information based on Name property...
'---------------------------------------------------------------------------------------
Public Function LoadUserSettings() As Boolean
On Local Error GoTo LoadUserSettings_Error

Dim oUser       As CRecord

Set oUser = New CRecord

If oUser Is Nothing Then
    LoadUserSettings = False
    Exit Function
End If

oUser.RstSQL "SELECT * FROM " & gcUserTour_UserTbl & " WHERE NAME = '" & mMDIVar.Name & "'"

If 0 = oUser.RstRecordCount Then
    LoadUserSettings = False
    Set oUser = Nothing
    Exit Function
End If

With oUser
    
    mMDIVar.ID = .GetField(gcID)
    

    mMDIVar.Password = .GetField(gcUserTour_UserTbl_PassWord)
    mMDIVar.MetaFile = .GetField(gcUserTour_UserTbl_Metafile)
    
    'Both UserOptions & DailyOpt properties are bit fields
    mMDIVar.UserOptions.SetValue (CLng(.GetField(gcUserTour_UserTbl_BitField)))
    mMDIVar.DailyOpt.SetValue (.GetField(gcUserTour_UserTbl_DailyOptions))
    
    ' Get individual items from UserOptions bit field
    With mMDIVar.UserOptions
        mMDIVar.Load = (4 * .GetValue(BitFlags.User_Load3)) + _
                    (2 * .GetValue(BitFlags.User_Load2)) + _
                    (.GetValue(BitFlags.User_Load1))
    End With
    
    mMDIVar.Security = mMDIVar.UserOptions.GetBool(BitFlags.User_Security)
    mMDIVar.ShowMeta = mMDIVar.UserOptions.GetBool(BitFlags.User_ShowMeta)

    
End With

Set oUser = Nothing
LoadUserSettings = True

On Error GoTo 0
Exit Function

LoadUserSettings_Error:
    If bDebug Then Handle_Err Err, "LoadUserSettings-cMDI"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : AddUser
' DATE      : 6/28/04 22:13
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function AddUser() As Boolean

On Local Error GoTo AddUser_Error
'Declare local variables
Dim lQuery      As Long
Dim SQL         As String

lQuery = ObjTour.GetHandle()
' Point to UserTbl
' Move cursor to newly created record and load values
SQL = "SELECT * FROM " & gcUserTour_UserTbl
ObjTour.RstSQL lQuery, SQL
ObjTour.DBMoveFirst lQuery


With ObjTour

    .AddNew (lQuery)
        Call .DBSetField(gcUserTour_UserTbl_Name, Trim$(mMDIVar.Name), lQuery)
        Call .DBSetField(gcUserTour_UserTbl_DataPath, mMDIVar.Datapath, lQuery)
        Call .DBSetField(gcUserTour_UserTbl_Load, 0, lQuery)
    .Update (lQuery)

End With

'' Move cursor to newly created record and load values
'SQL = "SELECT * FROM " & gcUserTour_UserTbl & " WHERE Name = '" & Trim$(mMDIVar.Name) & "'"
'
'ObjTour.RstSQL lQuery, SQL
'ObjTour.DBMoveFirst lQuery

ObjTour.FreeHandle lQuery

On Error GoTo 0
Exit Function

AddUser_Error:
    If bDebug Then Handle_Err Err, "AddUser-cMDI"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : Delete
' DATE      : 7/1/04 22:15
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function Delete() As Boolean

On Local Error GoTo Delete_Error
'Declare local variables
Dim oUser   As CRecord

If vbYes = MsgBox(LoadResString(gcPrompt_Delete) & " the user " & info.Name, vbYesNo + vbCritical + vbDefaultButton2, gcTour_Win) Then
    Set oUser = New CRecord
    'find user
    oUser.RstSQL "SELECT * FROM " & gcUserTour_UserTbl & " WHERE ID = " & info.ID
    If oUser.RstRecordCount = 1 Then
        oUser.Delete gcUserTour_UserTbl, mMDIVar.ID, vbNullString
        Delete = True
    Else
        Delete = False
    End If
Else
    Delete = False
End If

If Not oUser Is Nothing Then
    Set oUser = Nothing
End If

On Error GoTo 0
Exit Function

Delete_Error:
    If bDebug Then Handle_Err Err, "Delete-cMDI"
    Resume Next

End Function
