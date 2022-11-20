VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTourInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mTourVar        As cTourVar
Private mRecord         As Recordset

Public Property Set Settings(pTourVar As cTourVar)
Set mTourVar = pTourVar
End Property

Public Property Get Settings() As cTourVar
    Set Settings = mTourVar
End Property

Public Function DBOpen(ByVal sDataSource As String, sTableName As String, lDBHandle As Long) As Boolean
' --------------------------------------------------
' This function is only called via the cTour object
' and is used to narrow data access to one function
' only...
' --------------------------------------------------
On Local Error GoTo DBOpen_Err
' This is need for Export...
If 10 <> lDBHandle Then
sDataSource = gcTour_Win
End If
If bDebug Then
    Handle_Err 0, "DBOpen-cTourInfo - OpenDatabase(" & objMdi.info.Datapath & gcTour_Win & ")"
End If
DBOpen = False

If Not bSQLDatabase Then
Select Case lDBHandle
    Case 1:

         Set dbs1 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        
        If sTableName <> "" Then
        Set rst1 = dbs1.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst1.RecordCount <> 0 Then
                rst1.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
        End If
    Case 2:
        Set dbs2 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        Set rst2 = dbs2.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst2.RecordCount <> 0 Then
                rst2.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
    Case 3:

        Set dbs3 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        Set rst3 = dbs3.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst3.RecordCount <> 0 Then
                rst3.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
    Case 4:
        Set dbs4 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        Set rst4 = dbs4.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst4.RecordCount <> 0 Then
                rst4.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
    Case 5:
        Set dbs5 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        Set rst5 = dbs5.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst5.RecordCount <> 0 Then
                rst5.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
    Case 6:
        Set dbs6 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        Set rst6 = dbs6.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst6.RecordCount <> 0 Then
                rst6.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
    Case 7:
        Set dbs7 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        Set rst7 = dbs7.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst7.RecordCount <> 0 Then
                rst7.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
    Case 8:
        Set dbs8 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        Set rst8 = dbs8.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst8.RecordCount <> 0 Then
                rst8.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
    Case 9:
        Set dbs9 = DBEngine.Workspaces(0).OpenDatabase(objMdi.info.Datapath & sDataSource, False, False, ";pwd=Tourwin")
        Set rst9 = dbs9.OpenRecordSet("Select * From " & sTableName & " Where ID = " & objMdi.info.ID, dbOpenDynaset)
            If rst9.RecordCount <> 0 Then
                rst9.MoveFirst
                DBOpen = True
            Else
                DBOpen = False
            End If
    Case 10:
        ' DB and Rst are reserved for Export...
        Set dbs10 = DBEngine.Workspaces(0).OpenDatabase(sDataSource, False, False, ";pwd=" & gcEXPORT_PASSWORD)
        Set rst10 = dbs10.OpenRecordSet("Select * From " & sTableName, dbOpenDynaset)
            If rst10.RecordCount <> 0 Then
                rst10.MoveFirst
            End If
            DBOpen = True
End Select
Else
' SQL Database
If rdoEnv Is Nothing Then
    ' Mark uncomment...
    'Set rdoEnv = rdoEnvironments(0)
End If

Select Case lDBHandle
    Case 1:
    Set rdoDb1 = _
            rdoEnv.OpenConnection("TourWin_SQL", _
            rdDriverNoPrompt, False, "UID=TourAdmin;PWD=tourwin")
    Case 2:
    Set rdoDb2 = _
            rdoEnv.OpenConnection("TourWin_SQL", _
            rdDriverNoPrompt, False, "UID=TourAdmin;PWD=tourwin")
    Case 3:

    Set rdoDb3 = _
            rdoEnv.OpenConnection("TourWin_SQL", _
            rdDriverNoPrompt, False, "UID=TourAdmin;PWD=tourwin")
    Case 4:
    Set rdoDb4 = _
            rdoEnv.OpenConnection("TourWin_SQL", _
            rdDriverNoPrompt, False, "UID=TourAdmin;PWD=tourwin")
    Case 5:
    Set rdoDb5 = _
            rdoEnv.OpenConnection("TourWin_SQL", _
            rdDriverNoPrompt, False, "UID=TourAdmin;PWD=tourwin")
    Case 6:
    Set rdoDb6 = _
            rdoEnv.OpenConnection("TourWin_SQL", _
            rdDriverNoPrompt, False, "UID=TourAdmin;PWD=tourwin")
    Case 7:
    Case 8:
    Case 9:
End Select


End If
Exit Function
DBOpen_Err:
    MsgBox Err.Description
    If bDebug Then Handle_Err Err, "DBOpen-cTourInfo"
    Resume Next
End Function

Public Function DBMoveNext(lDBHandle As Long) As Boolean

On Local Error GoTo DBMove_Err
'Assume optimistic
DBMoveNext = True
If Not bSQLDatabase Then

Select Case lDBHandle
        Case 1:
                If False = rst1.EOF Then
                        rst1.MoveNext
                        If rst1.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If
        Case 2:
                If False = rst2.EOF Then
                        rst2.MoveNext
                        If rst2.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 3:
                If False = rst3.EOF Then
                        rst3.MoveNext
                        If rst3.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 4:
                If False = rst4.EOF Then
                        rst4.MoveNext
                        If rst4.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 5:
                If False = rst5.EOF Then
                        rst5.MoveNext
                        If rst5.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 6:
                If False = rst6.EOF Then
                        rst6.MoveNext
                        If rst6.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 7:
                If False = rst7.EOF Then
                        rst7.MoveNext
                        If rst7.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 8:
                If False = rst8.EOF Then
                        rst8.MoveNext
                        If rst8.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 9:
                If False = rst9.EOF Then
                        rst9.MoveNext
                        If rst9.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 10:
                If False = rst10.EOF Then
                        rst10.MoveNext
                        If rst10.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If
        
End Select
Else
Select Case lDBHandle
        Case 1:
                If False = rdoRs1.EOF Then
                        rdoRs1.MoveNext
                        If rdoRs1.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If
        Case 2:
                If False = rdoRs2.EOF Then
                        rdoRs2.MoveNext
                        If rdoRs2.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 3:
                If False = rdoRs3.EOF Then
                        rdoRs3.MoveNext
                        If rdoRs3.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 4:
                If False = rdoRs4.EOF Then
                        rdoRs4.MoveNext
                        If rdoRs4.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 5:
                If False = rdoRs5.EOF Then
                        rdoRs5.MoveNext
                        If rdoRs5.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 6:
                If False = rdoRs6.EOF Then
                        rdoRs6.MoveNext
                        If rdoRs6.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 7:
                If False = rdoRs7.EOF Then
                        rdoRs7.MoveNext
                        If rdoRs7.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 8:
                If False = rdoRs8.EOF Then
                        rdoRs8.MoveNext
                        If rdoRs8.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

        Case 9:
                If False = rdoRs9.EOF Then
                        rdoRs9.MoveNext
                        If rdoRs9.EOF Then DBMoveNext = False
                Else
                        DBMoveNext = False
                End If

End Select
End If
Exit Function
DBMove_Err:
    If bDebug Then Handle_Err Err, "DBMoveNext-cTourInfo"
    Resume Next
End Function
Public Function DBMovePrevious(lDBHandle As Long) As Boolean

On Local Error GoTo DBMovePrevious_Err
'Assume optimistic
        DBMovePrevious = True
If Not bSQLDatabase Then

Select Case lDBHandle
        Case 1:
                If False = rst1.BOF Then
                        rst1.MovePrevious
                Else
                        DBMovePrevious = False
                End If
        Case 2:
                If False = rst2.BOF Then
                        rst2.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 3:
                If False = rst3.BOF Then
                        rst3.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 4:
                If False = rst4.BOF Then
                        rst4.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 5:
                If False = rst5.BOF Then
                        rst5.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 6:
                If False = rst6.BOF Then
                        rst6.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 7:
                If False = rst7.BOF Then
                        rst7.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 8:
                If False = rst8.BOF Then
                        rst8.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 9:
                If False = rst9.BOF Then
                        rst9.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 10:
                If False = rst10.BOF Then
                        rst10.MovePrevious
                Else
                        DBMovePrevious = False
                End If
        
End Select
Else

Select Case lDBHandle
        Case 1:
                If False = rdoRs1.BOF Then
                        rdoRs1.MovePrevious
                Else
                        DBMovePrevious = False
                End If
        Case 2:
                If False = rdoRs2.BOF Then
                        rdoRs2.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 3:
                If False = rdoRs3.BOF Then
                        rdoRs3.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 4:
                If False = rdoRs4.BOF Then
                        rdoRs4.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 5:
                If False = rdoRs5.BOF Then
                        rdoRs5.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 6:
                If False = rdoRs6.BOF Then
                        rdoRs6.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 7:
                If False = rdoRs7.BOF Then
                        rdoRs7.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 8:
                If False = rdoRs8.BOF Then
                        rdoRs8.MovePrevious
                Else
                        DBMovePrevious = False
                End If

        Case 9:
                If False = rdoRs9.BOF Then
                        rdoRs9.MovePrevious
                Else
                        DBMovePrevious = False
                End If


End Select
End If
Exit Function
DBMovePrevious_Err:
    If bDebug Then Handle_Err Err, "DBMovePrevious-cTourInfo"
    Resume Next
End Function

'====================================================
' DBGetField - Purpose, returns the field value for
' the give record set.
'
' Return Value - Value of field or "" if Null
' ====================================================
Public Function DBGetField(sFieldName As String, lDBHandle As Long) As String
On Local Error GoTo DBGetField_Err

' ----------------------------------
' If no field given exit procedure
' -----------------------------------
If "" = sFieldName Then
    DBGetField = ""
    Exit Function
End If


If Not bSQLDatabase Then
Select Case lDBHandle
        Case 1:
            DBGetField = IIf(IsNull(rst1(sFieldName)), "", rst1(sFieldName))
        Case 2:
            DBGetField = IIf(IsNull(rst2(sFieldName)), "", rst2(sFieldName))
        Case 3:
            DBGetField = IIf(IsNull(rst3(sFieldName)), "", rst3(sFieldName))
        Case 4:
            DBGetField = IIf(IsNull(rst4(sFieldName)), "", rst4(sFieldName))
        Case 5:
            DBGetField = IIf(IsNull(rst5(sFieldName)), "", rst5(sFieldName))
        Case 6:
            DBGetField = IIf(IsNull(rst6(sFieldName)), "", rst6(sFieldName))
        Case 7:
            DBGetField = IIf(IsNull(rst7(sFieldName)), "", rst7(sFieldName))
        Case 8:
            DBGetField = IIf(IsNull(rst8(sFieldName)), "", rst8(sFieldName))
        Case 9:
            DBGetField = IIf(IsNull(rst9(sFieldName)), "", rst9(sFieldName))
        Case 10:
            DBGetField = IIf(IsNull(rst10(sFieldName)), "", rst10(sFieldName))
End Select
Else
Select Case lDBHandle
        Case 1:
            DBGetField = IIf(IsNull(rdoRs1(sFieldName)), "", rdoRs1(sFieldName))
        Case 2:
            DBGetField = IIf(IsNull(rdoRs2(sFieldName)), "", rdoRs2(sFieldName))
        Case 3:
            DBGetField = IIf(IsNull(rdoRs3(sFieldName)), "", rdoRs3(sFieldName))
        Case 4:
            DBGetField = IIf(IsNull(rdoRs4(sFieldName)), "", rdoRs4(sFieldName))
        Case 5:
            DBGetField = IIf(IsNull(rdoRs5(sFieldName)), "", rdoRs5(sFieldName))
        Case 6:
            DBGetField = IIf(IsNull(rdoRs6(sFieldName)), "", rdoRs6(sFieldName))
        Case 7:
            DBGetField = IIf(IsNull(rdoRs7(sFieldName)), "", rdoRs7(sFieldName))
        Case 8:
            DBGetField = IIf(IsNull(rdoRs8(sFieldName)), "", rdoRs8(sFieldName))
        Case 9:
            DBGetField = IIf(IsNull(rdoRs9(sFieldName)), "", rdoRs9(sFieldName))
End Select

End If
Exit Function
DBGetField_Err:
        If bDebug Then
            Handle_Err Err, "DBGetField-cTourInfo"
        End If
        DBGetField = "" 'gcFieldNotExist
        Resume Next
End Function
Public Function DBSetField(sFieldName As String, sFieldValue As Variant, lDBHandle As Long) As String
On Local Error GoTo DBSetField_Err
If Not bSQLDatabase Then

Select Case lDBHandle
    ' ----------------------------------------------
    ' Think I'll have to add some code that first
    ' determines the field data type the change
    ' the passed sFieldValue to appropriate type...
    ' ----------------------------------------------
        Case 1:
            rst1(sFieldName) = sFieldValue
        Case 2:
            rst2(sFieldName) = sFieldValue
        Case 3:
            rst3(sFieldName) = sFieldValue
        Case 4:
            rst4(sFieldName) = sFieldValue
        Case 5:
            rst5(sFieldName) = sFieldValue
        Case 6:
            rst6(sFieldName) = sFieldValue
        Case 7:
            rst7(sFieldName) = sFieldValue
        Case 8:
            rst8(sFieldName) = sFieldValue
        Case 9:
            rst9(sFieldName) = sFieldValue
        Case 10:
            rst10(sFieldName) = sFieldValue
End Select
Else

Select Case lDBHandle
    ' ----------------------------------------------
    ' Think I'll have to add some code that first
    ' determines the field data type the change
    ' the passed sFieldValue to appropriate type...
    ' ----------------------------------------------
        Case 1:
            rdoRs1(sFieldName) = sFieldValue
        Case 2:
            rdoRs2(sFieldName) = sFieldValue
        Case 3:
            rdoRs3(sFieldName) = sFieldValue
        Case 4:
            rdoRs4(sFieldName) = sFieldValue
        Case 5:
            rdoRs5(sFieldName) = sFieldValue
        Case 6:
            rdoRs6(sFieldName) = sFieldValue
        Case 7:
            rdoRs7(sFieldName) = sFieldValue
        Case 8:
            rdoRs8(sFieldName) = sFieldValue
        Case 9:
            rdoRs9(sFieldName) = sFieldValue
End Select

End If
Exit Function
DBSetField_Err:

       
        If bDebug Then Handle_Err Err, "DBSetField-cTourInfo"

  DBSetField = True
  Resume Next
End Function

Public Function DBClose(lDBHandle As Long) As Boolean
If bDebug Then
    Handle_Err 0, "DBClose-cTourInfo (Release rst=" & CStr(lDBHandle) & ")"
End If
Select Case lDBHandle
        Case 1:
            If Not rst1 Is Nothing Then Set rst1 = Nothing
            If Not dbs1 Is Nothing Then Set dbs1 = Nothing

        Case 2:
            If Not rst2 Is Nothing Then Set rst2 = Nothing
        Case 3:
            If Not rst3 Is Nothing Then Set rst3 = Nothing
            
        Case 4:
            If Not rst4 Is Nothing Then Set rst4 = Nothing
            
        Case 5:
            If Not rst5 Is Nothing Then Set rst5 = Nothing

        Case 6:
            If Not rst6 Is Nothing Then Set rst6 = Nothing

        Case 7:
            If Not rst7 Is Nothing Then Set rst7 = Nothing

        Case 8:
            If Not rst8 Is Nothing Then Set rst8 = Nothing

        Case 9:
            If Not rst9 Is Nothing Then Set rst9 = Nothing

        Case 10:
        
            If Not rst10 Is Nothing Then Set rst10 = Nothing
            If Not dbs10 Is Nothing Then Set dbs10 = Nothing
        Case 11:
        
            If Not rst1 Is Nothing Then Set rst1 = Nothing
            If Not rst2 Is Nothing Then Set rst2 = Nothing
            If Not rst3 Is Nothing Then Set rst3 = Nothing
            If Not rst4 Is Nothing Then Set rst4 = Nothing
            If Not rst5 Is Nothing Then Set rst5 = Nothing
            If Not rst6 Is Nothing Then Set rst6 = Nothing
            If Not rst7 Is Nothing Then Set rst7 = Nothing
            If Not rst8 Is Nothing Then Set rst8 = Nothing
            If Not rst9 Is Nothing Then Set rst9 = Nothing
            If Not rst10 Is Nothing Then Set rst10 = Nothing
            
            If Not dbs1 Is Nothing Then Set dbs1 = Nothing
        
End Select

End Function
Public Function DBExecute(sQuery As String) As Boolean

    cTour_DB.DBExecute sQuery
' dbs1.Execute sQuery

End Function

Public Function DBFindFirst(sFind As String, lDBHandle As Long) As String
On Local Error GoTo DBFindFirst_Err

If bDebug Then Handle_Err 0, "DBFindFirst-cTourInfo"

Select Case lDBHandle
    ' ----------------------------------------------
    ' Think I'll have to add some code that first
    ' determines the field data type the change
    ' the passed sFieldValue to appropriate type...
    ' ----------------------------------------------
        Case 1:
            rst1.FindFirst sFind
        Case 2:
            rst2.FindFirst sFind
        Case 3:
            rst3.FindFirst sFind
        Case 4:
            rst4.FindFirst sFind
        Case 5:
            rst5.FindFirst sFind
        Case 6:
            rst6.FindFirst sFind
        Case 7:
            rst7.FindFirst sFind
        Case 8:
            rst8.FindFirst sFind
        Case 9:
            rst9.FindFirst sFind
        Case 10:
            rst10.FindFirst sFind

End Select
Exit Function
DBFindFirst_Err:
    If bDebug Then Handle_Err Err, "DBFindFirst-cTourInfo"
End Function

Public Function DBFindPrevious(sFind As String, lDBHandle As Long) As String
On Local Error GoTo DBFindPrevious_Err

If bDebug Then Handle_Err 0, "DBFindPrevious-cTourInfo"
Select Case lDBHandle
    ' ----------------------------------------------
    ' Think I'll have to add some code that first
    ' determines the field data type the change
    ' the passed sFieldValue to appropriate type...
    ' ----------------------------------------------
        Case 1:
            rst1.FindPrevious sFind
        Case 2:
            rst2.FindPrevious sFind
        Case 3:
            rst3.FindPrevious sFind
        Case 4:
            rst4.FindPrevious sFind
        Case 5:
            rst5.FindPrevious sFind
        Case 6:
            rst6.FindPrevious sFind
        Case 7:
            rst7.FindPrevious sFind
        Case 8:
            rst8.FindPrevious sFind
        Case 9:
            rst9.FindPrevious sFind
        Case 10:
            rst10.FindPrevious sFind
End Select
Exit Function
DBFindPrevious_Err:
    If bDebug Then Handle_Err Err, "DBFindPrevious-cTourInfo"
End Function
Public Function DBFindNext(sFind As String, lDBHandle As Long) As String
On Local Error GoTo DBFindNext_Err

If bDebug Then Handle_Err 0, "DBFindNext-cTourInfo"
Select Case lDBHandle
    ' ----------------------------------------------
    ' Think I'll have to add some code that first
    ' determines the field data type the change
    ' the passed sFieldValue to appropriate type...
    ' ----------------------------------------------
        Case 1:
            rst1.FindNext sFind
        Case 2:
            rst2.FindNext sFind
        Case 3:
            rst3.FindNext sFind
        Case 4:
            rst4.FindNext sFind
        Case 5:
            rst5.FindNext sFind
        Case 6:
            rst6.FindNext sFind
        Case 7:
            rst7.FindNext sFind
        Case 8:
            rst8.FindNext sFind
        Case 9:
            rst9.FindNext sFind
        Case 10:
            rst10.FindNext sFind
End Select
Exit Function
DBFindNext_Err:
    If bDebug Then Handle_Err Err, "DBFindNext-cTourInfo"
End Function

Public Function GetHandle() As Long
Dim lLoop As Long
' ---------------------------------------------------
' If GetHandle returns 0, then no handles are
' available and user should be prompted to close
' some dbHook, also a good way to check that this
' program is releasing all database handles
' ---------------------------------------------------
GetHandle = 0
For lLoop = LBound(aHandle) To UBound(aHandle)
    If 0 = aHandle(lLoop) Then
            aHandle(lLoop) = lLoop
            GetHandle = lLoop
            Exit For
    End If
Next

If 0 = GetHandle Then MsgBox "Low resource, unable to create Database Handle", vbOKOnly, gcTourVersion

If bDebug Then
    Handle_Err 0, "GetHandle-cTourInfo (Handle=" & CStr(GetHandle) & ")"
End If
End Function

Public Function FreeHandle(lHandle As Long) As Boolean

If bDebug Then
    Handle_Err 0, "FreeHandle-cTourInfo (Handle=" & CStr(lHandle) & ")"
End If

    If lHandle <= UBound(aHandle) And lHandle >= LBound(aHandle) Then
        aHandle(lHandle) = 0
        FreeHandle = True
    End If
    
End Function

Public Function RstRecordCount(lHandle As Long) As Long
On Local Error GoTo RstRecordCount_Err

If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            RstRecordCount = rst1.RecordCount
        Case 2:
            RstRecordCount = rst2.RecordCount
        Case 3:
            RstRecordCount = rst3.RecordCount
        Case 4:
            RstRecordCount = rst4.RecordCount
        Case 5:
            RstRecordCount = rst5.RecordCount
        Case 6:
            RstRecordCount = rst6.RecordCount
        Case 7:
            RstRecordCount = rst7.RecordCount
        Case 8:
            RstRecordCount = rst8.RecordCount
        Case 9:
            RstRecordCount = rst9.RecordCount
        Case 10:
            RstRecordCount = rst10.RecordCount
End Select
Else
Select Case lHandle
        Case 1:
            RstRecordCount = rdoRs1.RowCount
        Case 2:
            RstRecordCount = rdoRs2.RowCount
        Case 3:
            RstRecordCount = rdoRs3.RowCount
        Case 4:
            RstRecordCount = rdoRs4.RowCount
        Case 5:
            RstRecordCount = rdoRs5.RowCount
        Case 6:
            RstRecordCount = rdoRs6.RowCount
        Case 7:
            RstRecordCount = rdoRs7.RowCount
        Case 8:
            RstRecordCount = rdoRs8.RowCount
        Case 9:
            RstRecordCount = rdoRs9.RowCount
End Select

End If

Exit Function
RstRecordCount_Err:
        MsgBox Err.Description
        If bDebug Then Handle_Err Err, "RstRecordCount_Err-cTourInfo"
        RstRecordCount = 0
        Resume Next

End Function
Public Function AddNew(lHandle As Long) As Boolean
On Local Error GoTo AddNew_Err
' Optimistic
AddNew = True
If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            rst1.AddNew
        Case 2:
            rst2.AddNew
        Case 3:
            rst3.AddNew
        Case 4:
            rst4.AddNew
        Case 5:
            rst5.AddNew
        Case 6:
            rst6.AddNew
        Case 7:
            rst7.AddNew
        Case 8:
            rst8.AddNew
        Case 9:
            rst9.AddNew
        Case 10:
            rst10.AddNew
End Select
Else
Select Case lHandle
        Case 1:
            rdoRs1.AddNew
        Case 2:
            rdoRs2.AddNew
        Case 3:
            rdoRs3.AddNew
        Case 4:
            rdoRs4.AddNew
        Case 5:
            rdoRs5.AddNew
        Case 6:
            rdoRs6.AddNew
        Case 7:
            rdoRs7.AddNew
        Case 8:
            rdoRs8.AddNew
        Case 9:
            rdoRs9.AddNew
End Select

End If
Exit Function

AddNew_Err:
        If bDebug Then Handle_Err Err, "AddNew_Err-cTourInfo"
        AddNew = False
        Resume Next

End Function

Public Function Delete(lHandle As Long) As Boolean
On Local Error GoTo Delete_Err
' Optimistic
Delete = True
If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            rst1.Delete
        Case 2:
            rst2.Delete
        Case 3:
            rst3.Delete
        Case 4:
            rst4.Delete
        Case 5:
            rst5.Delete
        Case 6:
            rst6.Delete
        Case 7:
            rst7.Delete
        Case 8:
            rst8.Delete
        Case 9:
            rst9.Delete
        Case 10:
            rst10.Delete
End Select
Else
Select Case lHandle
        Case 1:
            rdoRs1.Delete
        Case 2:
            rdoRs2.Delete
        Case 3:
            rdoRs3.Delete
        Case 4:
            rdoRs4.Delete
        Case 5:
            rdoRs5.Delete
        Case 6:
            rdoRs6.Delete
        Case 7:
            rdoRs7.Delete
        Case 8:
            rdoRs8.Delete
        Case 9:
            rdoRs9.Delete
End Select

End If
Exit Function
Delete_Err:
        If bDebug Then Handle_Err Err, "Delete_Err-cTourInfo"
        Delete = False
        Resume Next

End Function

Public Function Update(lHandle As Long) As Boolean
On Local Error GoTo Update_Err
' Optimistic
Update = True
If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            rst1.Update
        Case 2:
            rst2.Update
        Case 3:
            rst3.Update
        Case 4:
            rst4.Update
        Case 5:
            rst5.Update
        Case 6:
            rst6.Update
        Case 7:
            rst7.Update
        Case 8:
            rst8.Update
        Case 9:
            rst9.Update
        Case 10:
            rst10.Update
End Select
Else
Select Case lHandle
        Case 1:
            rdoRs1.Update
        Case 2:
            rdoRs2.Update
        Case 3:
            rdoRs3.Update
        Case 4:
            rdoRs4.Update
        Case 5:
            rdoRs5.Update
        Case 6:
            rdoRs6.Update
        Case 7:
            rdoRs7.Update
        Case 8:
            rdoRs8.Update
        Case 9:
            rdoRs9.Update
End Select
End If
Exit Function
Update_Err:

        If bDebug Then
        MsgBox Error$(Err)
        Handle_Err Err, "Update_Err-cTourInfo"
        End If
        Update = False
        Resume Next

End Function

Public Function Edit(lHandle As Long) As Boolean
On Local Error GoTo Edit_Err
' Optimistic
Edit = True

If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            rst1.Edit
        Case 2:
            rst2.Edit
        Case 3:
            rst3.Edit
        Case 4:
            rst4.Edit
        Case 5:
            rst5.Edit
        Case 6:
            rst6.Edit
        Case 7:
            rst7.Edit
        Case 8:
            rst8.Edit
        Case 9:
            rst9.Edit
        Case 10:
            rst10.Edit
End Select
Else
Select Case lHandle
        Case 1:
            rdoRs1.Edit
        Case 2:
            rdoRs2.Edit
        Case 3:
            rdoRs3.Edit
        Case 4:
            rdoRs4.Edit
        Case 5:
            rdoRs5.Edit
        Case 6:
            rdoRs6.Edit
        Case 7:
            rdoRs7.Edit
        Case 8:
            rdoRs8.Edit
        Case 9:
            rdoRs9.Edit
End Select
End If
Exit Function
Edit_Err:
        If bDebug Then Handle_Err Err, "Edit_Err-cTourInfo"
        Edit = False
        Resume Next
End Function

Public Function CancelUpdate(lHandle As Long) As Boolean
On Local Error GoTo CancelUpdate_Err
' Optimistic
CancelUpdate = True
Select Case lHandle
        Case 1:
            rst1.CancelUpdate
        Case 2:
            rst2.CancelUpdate
        Case 3:
            rst3.CancelUpdate
        Case 4:
            rst4.CancelUpdate
        Case 5:
            rst5.CancelUpdate
        Case 6:
            rst6.CancelUpdate
        Case 7:
            rst7.CancelUpdate
        Case 8:
            rst8.CancelUpdate
        Case 9:
            rst9.CancelUpdate
        Case 10:
            rst10.CancelUpdate
End Select
Exit Function
CancelUpdate_Err:
        If bDebug Then Handle_Err Err, "CancelUpdate_Err-cTourInfo"
        CancelUpdate = False
        Resume Next
End Function

Public Function RstSQL(lHandle As Long, sSQL As String) As Boolean
On Local Error GoTo RstSQL_Err
'Assume optimistic
RstSQL = True
If lHandle = 0 Then
    lHandle = ObjTour.GetHandle
End If

If bDebug Then
    Handle_Err 0, "RstSQL-cTourInfo (Handle=" & CStr(lHandle) & ")" & vbCrLf & sSQL
End If

If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            Set rst1 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 2:
            Set rst2 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 3:
            Set rst3 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 4:
            Set rst4 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 5:
            Set rst5 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 6:
            Set rst6 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 7:
            Set rst7 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 8:
            Set rst8 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 9:
            Set rst9 = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
        Case 10: ' export
            Set rst10 = dbs10.OpenRecordSet(sSQL, dbOpenDynaset)
End Select
Else
Select Case lHandle
        Case 1:
            If rdoDb1 Is Nothing Then
                ObjTour.DBOpen objMdi.info.Datapath, "", lHandle
            End If
            Set rdoRs1 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
        Case 2:
            Set rdoRs2 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
        Case 3:
            Set rdoRs3 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
        Case 4:
            Set rdoRs4 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
        Case 5:
            Set rdoRs5 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
        Case 6:
            Set rdoRs6 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
        Case 7:
            Set rdoRs7 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
        Case 8:
            Set rdoRs8 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
        Case 9:
            Set rdoRs9 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
End Select
End If
Exit Function
RstSQL_Err:
    
        If bDebug Then
            MsgBox Error$(Err)
            Handle_Err Err, "RstSQL_Err-cTourInfo"
        End If
        RstSQL = False
        Resume Next
End Function
Public Function NoMatch(lHandle As Long) As Boolean
On Local Error GoTo NoMatch_Err
'Assume optimistic
NoMatch = True
Select Case lHandle
        Case 1:
            NoMatch = rst1.NoMatch
        Case 2:
            NoMatch = rst2.NoMatch
        Case 3:
            NoMatch = rst3.NoMatch
        Case 4:
            NoMatch = rst4.NoMatch
        Case 5:
            NoMatch = rst5.NoMatch
        Case 6:
            NoMatch = rst6.NoMatch
        Case 7:
            NoMatch = rst7.NoMatch
        Case 8:
            NoMatch = rst8.NoMatch
        Case 9:
            NoMatch = rst9.NoMatch
        Case 10:
            NoMatch = rst10.NoMatch
End Select
Exit Function
NoMatch_Err:

        If bDebug Then Handle_Err Err, "NoMatch_Err-cTourInfo"
        NoMatch = False
        Resume Next
End Function

Public Function DBMoveLast(lHandle As Long) As Boolean
On Local Error GoTo DBMoveLast_Err
' Optimistic
DBMoveLast = True
If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            rst1.MoveLast
        Case 2:
            rst2.MoveLast
        Case 3:
            rst3.MoveLast
        Case 4:
            rst4.MoveLast
        Case 5:
            rst5.MoveLast
        Case 6:
            rst6.MoveLast
        Case 7:
            rst7.MoveLast
        Case 8:
            rst8.MoveLast
        Case 9:
            rst9.MoveLast
        Case 10:
            rst10.MoveLast
End Select
Else
Select Case lHandle
        Case 1:
            rdoRs1.MoveLast
        Case 2:
            rdoRs2.MoveLast
        Case 3:
            rdoRs3.MoveLast
        Case 4:
            rdoRs4.MoveLast
        Case 5:
            rdoRs5.MoveLast
        Case 6:
            rdoRs6.MoveLast
        Case 7:
            rdoRs7.MoveLast
        Case 8:
            rdoRs8.MoveLast
        Case 9:
            rdoRs9.MoveLast
End Select

End If
Exit Function
DBMoveLast_Err:
        If bDebug Then Handle_Err Err, "DBMoveLast_Err-cTourInfo (Handle=" & CStr(lHandle) & ")"
        DBMoveLast = False
        Resume Next
End Function
Public Function DBMoveFirst(lHandle As Long) As Boolean
On Local Error GoTo DBMoveFirst_Err
' Optimistic
DBMoveFirst = True
If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            rst1.MoveFirst
        Case 2:
            rst2.MoveFirst
        Case 3:
            rst3.MoveFirst
        Case 4:
            rst4.MoveFirst
        Case 5:
            rst5.MoveFirst
        Case 6:
            rst6.MoveFirst
        Case 7:
            rst7.MoveFirst
        Case 8:
            rst8.MoveFirst
        Case 9:
            rst9.MoveFirst
        Case 10:
            rst10.MoveFirst
End Select
Else
Select Case lHandle
        Case 1:
            rdoRs1.MoveFirst
        Case 2:
            rdoRs2.MoveFirst
        Case 3:
            rdoRs3.MoveFirst
        Case 4:
            rdoRs4.MoveFirst
        Case 5:
            rdoRs5.MoveFirst
        Case 6:
            rdoRs6.MoveFirst
        Case 7:
            rdoRs7.MoveFirst
        Case 8:
            rdoRs8.MoveFirst
        Case 9:
            rdoRs9.MoveFirst
End Select

End If
Exit Function
DBMoveFirst_Err:
        If bDebug Then Handle_Err Err, "DBMoveFirst_Err-cTourInfo"
        DBMoveFirst = False
        Resume Next
End Function

Public Function EOF(lHandle As Long) As Boolean
On Local Error GoTo EOF_Err
' Optimistic
If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            EOF = rst1.EOF
        Case 2:
            EOF = rst2.EOF
        Case 3:
            EOF = rst3.EOF
        Case 4:
            EOF = rst4.EOF
        Case 5:
            EOF = rst5.EOF
        Case 6:
            EOF = rst6.EOF
        Case 7:
            EOF = rst7.EOF
        Case 8:
            EOF = rst8.EOF
        Case 9:
            EOF = rst9.EOF
        Case 10:
            EOF = rst10.EOF
End Select
Else
Select Case lHandle
        Case 1:
            EOF = rdoRs1.EOF
        Case 2:
            EOF = rdoRs2.EOF
        Case 3:
            EOF = rdoRs3.EOF
        Case 4:
            EOF = rdoRs4.EOF
        Case 5:
            EOF = rdoRs5.EOF
        Case 6:
            EOF = rdoRs6.EOF
        Case 7:
            EOF = rdoRs7.EOF
        Case 8:
            EOF = rdoRs8.EOF
        Case 9:
            EOF = rdoRs9.EOF
End Select

End If
Exit Function

EOF_Err:
        If bDebug Then Handle_Err Err, "EOF_Err-cTourInfo"
        EOF = True
        Resume Next
End Function
Public Function BOF(lHandle As Long) As Boolean
On Local Error GoTo BOF_Err
' Optimistic
If Not bSQLDatabase Then
Select Case lHandle
        Case 1:
            BOF = rst1.BOF
        Case 2:
            BOF = rst2.BOF
        Case 3:
            BOF = rst3.BOF
        Case 4:
            BOF = rst4.BOF
        Case 5:
            BOF = rst5.BOF
        Case 6:
            BOF = rst6.BOF
        Case 7:
            BOF = rst7.BOF
        Case 8:
            BOF = rst8.BOF
        Case 9:
            BOF = rst9.BOF
        Case 10:
            BOF = rst10.BOF
    End Select
Else
    Select Case lHandle
        Case 1:
            BOF = rdoRs1.BOF
        Case 2:
            BOF = rdoRs2.BOF
        Case 3:
            BOF = rdoRs3.BOF
        Case 4:
            BOF = rdoRs4.BOF
        Case 5:
            BOF = rdoRs5.BOF
        Case 6:
            BOF = rdoRs6.BOF
        Case 7:
            BOF = rdoRs7.BOF
        Case 8:
            BOF = rdoRs8.BOF
        Case 9:
            BOF = rdoRs9.BOF
    End Select
End If

Exit Function
BOF_Err:
        If bDebug Then Handle_Err Err, "BOF_Err-cTourInfo"
        BOF = True
        Resume Next
End Function

Public Function BookMark(lHandle As Long, Optional vBookMark As Variant) As Variant
On Local Error GoTo BookMark_Err
' Optimistic
If Not IsNull(vBookMark) And vBookMark <> "" Then
    Select Case lHandle
        Case 1:
            BookMark = rst1.BookMark
        Case 2:
            BookMark = rst2.BookMark
        Case 3:
            BookMark = rst3.BookMark
        Case 4:
            BookMark = rst4.BookMark
        Case 5:
            BookMark = rst5.BookMark
        Case 6:
            BookMark = rst6.BookMark
        Case 7:
            BookMark = rst7.BookMark
        Case 8:
            BookMark = rst8.BookMark
        Case 9:
            BookMark = rst9.BookMark
        Case 10:
            BookMark = rst10.BookMark
    End Select
Else
    Select Case lHandle
        Case 1:
            rst1.BookMark = vBookMark
        Case 2:
            rst2.BookMark = vBookMark
        Case 3:
            rst3.BookMark = vBookMark
        Case 4:
            rst4.BookMark = vBookMark
        Case 5:
            rst5.BookMark = vBookMark
        Case 6:
            rst6.BookMark = vBookMark
        Case 7:
            rst7.BookMark = vBookMark
        Case 8:
            rst8.BookMark = vBookMark
        Case 9:
            rst9.BookMark = vBookMark
        Case 10:
            rst10.BookMark = vBookMark
    End Select

End If
Exit Function
BookMark_Err:

        If bDebug Then Handle_Err Err, "BookMark_Err-cTourInfo"
        BookMark = 0
        Resume Next
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

Public Function CreateQueryDefObject(ByVal CreatePath As String, ByVal sName As String, ByVal sSQL As String) As String
On Local Error GoTo Create_Error

' Declare local variables
Dim DefaultWorkspace As Workspace
Dim TourDatabase As Database
Dim qTempQueryDef As QueryDef

' Initialize objects
Set DefaultWorkspace = DBEngine.Workspaces(0)
Set TourDatabase = DefaultWorkspace.OpenDatabase(CreatePath & "\" & gcTour_Win, False, False, ";pwd=Tourwin")
Set qTempQueryDef = TourDatabase.CreateQueryDef(sName)

' Assigning SQL value creates Def right now
qTempQueryDef.SQL = sSQL

' close and release objects
qTempQueryDef.Close
Set qTempQueryDef = Nothing

TourDatabase.Close
Set TourDatabase = Nothing

DefaultWorkspace.Close
Set DefaultWorkspace = Nothing

On Local Error GoTo 0
Exit Function
Create_Error:
    If bDebug Then Handle_Err Err, "CreateQueryDefObject-cTourInfo"
    Resume Next
End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : RegisterUpdateDSN
' DATE      : 4/1/03 20:07
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function RegisterUpdateDSN(ByVal sPath As String) As Boolean

On Local Error GoTo RegisterUpdateDSN_Error
'Declarations Used to Generate DSN
Const ODBC_ADD_DSN = 1 ' Add data source
Const ODBC_CONFIG_DSN = 2 ' Configure (edit) data source
Const ODBC_REMOVE_DSN = 3 ' Remove data source
Const vbAPINull As Long = 0& ' NULL Pointer

'Declare local variables
Dim lRet As Long
Dim sDriver As String
Dim sAttributes As String

sDriver = "Microsoft Access Driver (*.mdb)"

sAttributes = "Description=DSN for TourWin reports" & _
vbCr & "dbq=" & sPath & _
vbCr & "UID=" & _
vbCr & "PWD=Tourwin" & _
vbCr & "server=Microsoft Access Driver (*.mdb)" '& _
'vbCr & "ReadOnly=1"
    
DBEngine.RegisterDatabase "Tourwin", "Microsoft Access Driver (*.mdb)", True, sAttributes
lRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, sDriver, sAttributes)

On Error GoTo 0
Exit Function

RegisterUpdateDSN_Error:
    MsgBox Err.Description
    If bDebug Then Handle_Err Err, "RegisterUpdateDSN-cTourInfo"
    Resume Next

End Function
