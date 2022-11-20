VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------
' CBase:        Base class for all business objects
' Author:       Mark Ormesher
' Description:  Implements CRecordSet and delegates
' Revision:
' -----------------------------------------------------------------------------
Option Explicit

Implements CRecord                      ' Inherit CRecord Class
Private m_CRecord   As CRecord          ' Private Object to delegate

Private Sub Class_Initialize()
    Set m_CRecord = New CRecord         ' Create
End Sub

Private Sub Class_Terminate()
    Set m_CRecord = Nothing             ' Release
End Sub

Public Function CRecord_AddNew() As Boolean
        m_CRecord.AddNew                ' Delegate
End Function

Public Function CRecord_BOF() As Boolean
    CRecord_BOF = m_CRecord.BOF
End Function

Public Function CRecord_BookMark(lHandle As Long, Optional vBookMark As Variant) As Variant

End Function

Public Function CRecord_CancelUpdate() As Boolean
    CRecord_CancelUpdate = m_CRecord.CancelUpdate
End Function

Public Function CRecord_Delete() As Boolean
    CRecord_Delete = m_CRecord.Delete
End Function

Public Function CRecord_Edit() As Boolean
    CRecord_Edit = m_CRecord.Edit
End Function

Public Function CRecord_EOF() As Boolean
    CRecord_EOF = m_CRecord.EOF
End Function

Public Function CRecord_FindFirst(sFind As String) As Boolean
    CRecord_FindFirst = m_CRecord.FindFirst(sFind)
End Function

Public Function CRecord_FindNext(sFind As String) As Boolean
    CRecord_FindNext = m_CRecord.FindNext(sFind)
End Function

Public Function CRecord_FindPrevious(sFind As String) As Boolean
    CRecord_FindPrevious = m_CRecord.FindPrevious(sFind)
End Function

Public Function CRecord_GetField(ByVal sFieldName As String) As String
    CRecord_GetField = m_CRecord.GetField(sFieldName)
End Function

Public Function CRecord_MoveFirst() As Boolean
    CRecord_MoveFirst = m_CRecord.MoveFirst
End Function

Public Function CRecord_MoveLast() As Boolean
    CRecord_MoveLast = m_CRecord.MoveLast
End Function

Public Function CRecord_MoveNext() As Boolean
    CRecord_MoveNext = m_CRecord.MoveNext
End Function

Public Function CRecord_MovePrevious() As Boolean
    CRecord_MovePrevious = m_CRecord.MovePrevious
End Function

Public Function CRecord_NoMatch() As Boolean
    CRecord_NoMatch = m_CRecord.NoMatch
End Function

Public Function CRecord_RstRecordCount() As Long
    CRecord_RstRecordCount = m_CRecord.RstRecordCount
End Function

Public Function CRecord_RstSQL(ByVal sSQL As String) As Boolean
    CRecord_RstSQL = m_CRecord.RstSQL(sSQL)
End Function

Public Function CRecord_SetField(ByVal sFieldName As String, ByVal sFieldValue As Variant) As Boolean
    CRecord_SetField = m_CRecord.SetField(sFieldName, sFieldValue)
End Function

Public Function CRecord_Update() As Boolean
    CRecord_Update = m_CRecord.Update
End Function
