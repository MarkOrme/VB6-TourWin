VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------
' CRecord:      Used to retrieve rows from CDatabase class
' Author:       Mark Ormesher
' Description:  Interface from DataLayer to business layer
' Revision:
' -----------------------------------------------------------------------------
Option Explicit
Private m_RstAccess         As Recordset    ' RecordSet for Access
Private m_RstSQL            As rdoResultset ' Recordset for SQL

Public Function AddNew() As Boolean

On Local Error GoTo AddNew_Err
' Optimistic
AddNew = True

If Not bSQLDatabase Then
    m_RstAccess.AddNew
Else
    m_RstSQL.AddNew
End If


On Local Error GoTo 0

Exit Function
AddNew_Err:
AddNew = False
If bDebug Then Handle_Err Err, "AddNew_Err-cTourInfo"
AddNew = False
Resume Next

End Function
Public Function BOF() As Boolean
On Local Error GoTo BOF_Err

If bSQLDatabase Then
    BOF = m_RstSQL.BOF
Else
    BOF = m_RstAccess.BOF
End If

On Local Error GoTo 0
Exit Function
BOF_Err:
        If bDebug Then Handle_Err Err, "BOF_Err-cTourInfo"
        BOF = True
        Resume Next
End Function

Public Function EOF() As Boolean
On Local Error GoTo EOF_Err

If bSQLDatabase Then
    EOF = m_RstSQL.EOF
Else
    EOF = m_RstAccess.EOF
End If

    
On Local Error GoTo 0
Exit Function
EOF_Err:
If bDebug Then Handle_Err Err, "EOF_Err-cTourInfo"
EOF = True
Resume Next

End Function

Public Function BookMark(lHandle As Long, Optional vBookMark As Variant) As Variant
On Local Error GoTo BookMark_Err

' Optimistic
If Not IsNull(vBookMark) And vBookMark <> "" Then
    BookMark = IIf(bSQLDatabase, m_RstSQL.BookMark, m_RstAccess.BookMark)
End If

On Local Error GoTo 0
Exit Function
BookMark_Err:

If bDebug Then Handle_Err Err, "BookMark_Err-cTourInfo"
BookMark = 0
Resume Next
End Function

Public Function MoveNext() As Boolean
On Local Error GoTo DBMove_Err

'Assume optimistic
MoveNext = False

If Not Me.EOF Then
    Select Case bSQLDatabase
            Case False: ' Access
                m_RstAccess.MoveNext
                If Not Me.EOF Then MoveNext = True
                
            Case True: ' SQL
                m_RstSQL.MoveNext
                If Not Me.EOF Then MoveNext = True
    End Select
End If

On Local Error GoTo 0
Exit Function
DBMove_Err:
If bDebug Then Handle_Err Err, "DBMoveNext-cTourInfo"
Resume Next
End Function
Public Function MovePrevious() As Boolean

On Local Error GoTo DBMovePrevious_Err
'Assume optimistic
MovePrevious = False

If Not Me.BOF Then
    Select Case bSQLDatabase
            Case False: ' Access
                m_RstAccess.MovePrevious
                If Not Me.BOF Then MovePrevious = True
                
            Case True: ' SQL
                m_RstSQL.MovePrevious
                If Not Me.BOF Then MovePrevious = True
    End Select
End If

On Local Error GoTo 0
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
Public Function GetField(ByVal sFieldName As String) As String

On Local Error GoTo DBGetField_Err
If Not bSQLDatabase Then
    GetField = CStr(IIf(IsNull(m_RstAccess(sFieldName)), "", m_RstAccess(sFieldName)))
Else
    GetField = CStr(IIf(IsNull(m_RstSQL(sFieldName)), "", m_RstSQL(sFieldName)))
End If

On Local Error GoTo 0
Exit Function

DBGetField_Err:
If bDebug Then
    Handle_Err Err, "GetField-cRecord"
End If
GetField = ""
Resume Next

End Function
Public Function SetField(ByVal sFieldName As String, ByVal sFieldValue As Variant) As Boolean
On Local Error GoTo DBSetField_Err

' ----------------------------------------------
' Think I'll have to add some code that first
' determines the field data type the change
' the passed sFieldValue to appropriate type...
' ----------------------------------------------

SetField = True

If Not bSQLDatabase Then
   m_RstAccess(sFieldName) = sFieldValue
Else
   m_RstSQL(sFieldName) = sFieldValue
End If

On Local Error GoTo 0
Exit Function

DBSetField_Err:
If bDebug Then Handle_Err Err, "SetField-cRecord"
SetField = False

Resume Next
End Function

Public Function FindFirst(sFind As String) As Boolean
On Local Error GoTo DBFindFirst_Err

If bDebug Then Handle_Err 0, "DBFindFirst-cTourInfo"

FindFirst = False
If Not bSQLDatabase Then
    m_RstAccess.FindFirst sFind
    FindFirst = Not m_RstAccess.NoMatch
Else
    'm_RstSQL.f
End If
        
On Local Error GoTo 0
Exit Function
DBFindFirst_Err:
FindFirst = False
If bDebug Then Handle_Err Err, "DBFindFirst-cTourInfo"

End Function

Public Function FindPrevious(sFind As String) As Boolean
On Local Error GoTo DBFindPrevious_Err

FindPrevious = False

If bDebug Then Handle_Err 0, "DBFindPrevious-cTourInfo"

If Not bSQLDatabase Then
    m_RstAccess.FindPrevious sFind
    FindPrevious = Not m_RstAccess.NoMatch
Else
    'm_RstSQL.fin
End If

On Local Error GoTo 0
Exit Function
FindPrevious = False
DBFindPrevious_Err:
If bDebug Then Handle_Err Err, "DBFindPrevious-cTourInfo"
End Function
Public Function FindNext(sFind As String) As Boolean
On Local Error GoTo DBFindNext_Err
If bDebug Then Handle_Err 0, "DBFindNext-cTourInfo"

FindNext = False
If Not bSQLDatabase Then
    m_RstAccess.FindNext sFind
    FindNext = Not m_RstAccess.NoMatch
Else
    '    m_RstSQL.fi
End If

On Local Error GoTo 0
Exit Function

FindNext = False
DBFindNext_Err:
If bDebug Then Handle_Err Err, "DBFindNext-cTourInfo"
End Function

Public Function Delete(ByVal sTableName As String, ByVal lID As Long, Optional sDeleteQuery As String = "") As Boolean
On Local Error GoTo Delete_Err

' Optimistic
Delete = True

' First perform delete...
' if OK then Delete appropriate records from child tables
'If "" = sDeleteQuery Then
'    If Not m_RstAccess.Delete The\\\\\\\

If sTableName = gcUserTour_UserTbl Then GoTo User_Table


User_Table:
Dai_Table:

If Not bSQLDatabase Then
    m_RstAccess.Delete
Else
    m_RstSQL.Delete
End If

On Local Error GoTo 0
Exit Function
Delete_Err:

If bDebug Then Handle_Err Err, "Delete_Err-cTourInfo"
Delete = False
Resume Next

End Function

Public Function Update() As Boolean
On Local Error GoTo Update_Err
' Optimistic
Update = True

If Not bSQLDatabase Then
    m_RstAccess.Update
Else
    m_RstSQL.Update
End If

On Local Error GoTo 0
Exit Function
Update_Err:

Update = False

If bDebug Then
MsgBox Error$(Err)
Handle_Err Err, "Update_Err-cTourInfo"
End If

Resume Next

End Function

Public Function Edit() As Boolean
On Local Error GoTo Edit_Err
' Optimistic
Edit = True

If Not bSQLDatabase Then
    m_RstAccess.Edit
Else
    m_RstSQL.Edit
End If

On Local Error GoTo 0
Exit Function
Edit_Err:
Edit = False
If bDebug Then Handle_Err Err, "Edit_Err-cTourInfo"
Resume Next

End Function

Public Function RstRecordCount() As Long
On Local Error GoTo RstRecordCount_Err

If Not bSQLDatabase Then
    RstRecordCount = m_RstAccess.RecordCount
Else
    RstRecordCount = m_RstSQL.RowCount
End If

On Local Error GoTo 0
Exit Function
RstRecordCount_Err:
RstRecordCount = 0

MsgBox Err.Description
If bDebug Then Handle_Err Err, "RstRecordCount_Err-cTourInfo"

Resume Next
End Function

Public Function RstSQL(ByVal sSQL As String) As Boolean
On Local Error GoTo RstSQL_Err

'Assume optimistic
RstSQL = True

If Not bSQLDatabase Then
    
    Set m_RstAccess = cTour_DB.OpenRecordSet(sSQL, dbOpenDynaset)
    If m_RstAccess Is Nothing Then RstSQL = False
Else
    'Set rdoRs9 = rdoDb1.OpenResultset(sSQL, rdOpenDynamic, rdConcurRowVer, 0)
End If

On Local Error GoTo 0
Exit Function
RstSQL_Err:
    
If bDebug Then
    MsgBox Error$(Err)
    Handle_Err Err, "RstSQL_Err-cTourInfo"
End If
RstSQL = False
Resume Next
End Function
Public Function NoMatch() As Boolean
On Local Error GoTo NoMatch_Err
'Assume optimistic
NoMatch = True

If Not bSQLDatabase Then
    NoMatch = m_RstAccess.NoMatch
Else
    'NoMatch = m_RstSQL.no
End If

On Local Error GoTo 0
Exit Function
NoMatch_Err:

NoMatch = False
If bDebug Then Handle_Err Err, "NoMatch_Err-cTourInfo"

Resume Next
End Function

Public Function MoveLast() As Boolean

On Local Error GoTo DBMoveLast_Err
' Optimistic
MoveLast = True
If Not bSQLDatabase Then
    m_RstAccess.MoveLast
Else
    m_RstSQL.MoveLast
End If

On Local Error GoTo 0
Exit Function
DBMoveLast_Err:
MoveLast = False

If bDebug Then Handle_Err Err, "DBMoveLast_Err-cTourInfo"
Resume Next
End Function
Public Function MoveFirst() As Boolean
On Local Error GoTo DBMoveFirst_Err
' Optimistic
MoveFirst = True
If Not bSQLDatabase Then
    m_RstAccess.MoveFirst
Else
    m_RstSQL.MoveFirst
End If

On Local Error GoTo 0
Exit Function
DBMoveFirst_Err:

MoveFirst = False
If bDebug Then Handle_Err Err, "DBMoveFirst_Err-cTourInfo"

Resume Next
End Function

Public Function CancelUpdate() As Boolean
On Local Error GoTo CancelUpdate_Err
' Optimistic
CancelUpdate = True

If Not bSQLDatabase Then
    m_RstAccess.CancelUpdate
Else
    m_RstSQL.CancelUpdate
End If
    
On Local Error GoTo 0
Exit Function
CancelUpdate_Err:
CancelUpdate = False

If bDebug Then Handle_Err Err, "CancelUpdate_Err-cTourInfo"

Resume Next
End Function

Private Sub Class_Terminate()

If Not m_RstAccess Is Nothing Then
    m_RstAccess.Close
    Set m_RstAccess = Nothing
End If


End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : Execute
' DATE      : 8/25/04 18:22
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal sQuery As String) As Boolean

On Local Error GoTo Execute_Error
'Declare local variables

    
       Execute = cTour_DB.DBExecute(sQuery)

On Error GoTo 0
Exit Function

Execute_Error:
    If bDebug Then Handle_Err Err, "Execute-CRecord"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : AbsolutePosition
' DATE      : 8/29/04 15:43
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function AbsolutePosition(ByVal lPosition As Long) As Boolean
On Local Error GoTo AbsolutePosition_Error
'Declare local variables
Dim bResult As Boolean

If Not bSQLDatabase Then
    m_RstAccess.AbsolutePosition = lPosition
Else
    m_RstSQL.AbsolutePosition = lPosition
End If

On Error GoTo 0
Exit Function

AbsolutePosition_Error:
    If bDebug Then Handle_Err Err, "AbsolutePosition-CRecord"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : GetRows
' DATE      : 8/29/04 15:48
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function GetRows(ByVal lCount As Long) As Variant
On Local Error GoTo GetRows_Error
'Declare local variables
Dim vResult As Variant

If Not bSQLDatabase Then
    GetRows = m_RstAccess.GetRows(lCount)
Else
    GetRows = m_RstSQL.GetRows(lCount)
End If

On Error GoTo 0
Exit Function

GetRows_Error:
    If bDebug Then Handle_Err Err, "GetRows-CRecord"
    Resume Next


End Function
