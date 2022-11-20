VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDailyBusiness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------
' CDailyBusiness:       Business rules for Dai table
' Author:               Mark Ormesher
' Description:          Implements CBase and delagates accordingly
' Revision:
' -----------------------------------------------------------------------------
Option Explicit
Implements CRecord                      ' Inherit CRecord Class
Private m_CRecord   As CRecord          ' Private Object to delegate
Private mDaiVar As cDaiVar

Private Sub Class_Initialize()
    Set m_CRecord = New CRecord         ' Create
End Sub

Private Sub Class_Terminate()
    Set m_CRecord = Nothing             ' Release
End Sub


' =============================================================================
' THE FOLLOWING METHOD ARE DEFAULT METHODS AND PROPERTIES FOR CDAILYBUSINESS.
' =============================================================================
Public Property Set info(pDaiVar As cDaiVar)
    Set mDaiVar = pDaiVar
End Property

Public Property Get info() As cDaiVar
    Set info = mDaiVar
End Property

Public Function Load_Daily_Settings() As Boolean
On Local Error GoTo Load_Daily_Error

Load_Daily_Settings = False

    objDai.info.Desc = IIf(ObjTour.DBGetField("Description", iSearcherDB) = "", False, ObjTour.DBGetField("Description", iSearcherDB))
    objDai.info.RestH = IIf(ObjTour.DBGetField("Rest Heart", iSearcherDB) = "", False, ObjTour.DBGetField("Rest Heart", iSearcherDB))
    objDai.info.Sleep = IIf(ObjTour.DBGetField("Sleep", iSearcherDB) = "", False, ObjTour.DBGetField("Sleep", iSearcherDB))
    objDai.info.DayR = IIf(ObjTour.DBGetField("DayInt", iSearcherDB) = "", False, ObjTour.DBGetField("DayInt", iSearcherDB))
    objDai.info.Weight = IIf(ObjTour.DBGetField("Weight", iSearcherDB) = "", False, ObjTour.DBGetField("Weight", iSearcherDB))
    objDai.info.Metric = IIf(ObjTour.DBGetField("Metric", iSearcherDB) = "", False, ObjTour.DBGetField("Metric", iSearcherDB))
    
Load_Daily_Settings = True
Exit Function
Load_Daily_Error:
     If bDebug Then
        MsgBox Error$(Err)
        Handle_Err Err, "Load_Daily_Settings-cDai"
     End If
     Exit Function
End Function

' =============================================================================
' THE FOLLOWING METHOD ARE INHERTED FROM THE BASE CLASS (CRECORD)
' =============================================================================
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
Dim sQuery        As String
' Set Recordset
sQuery = " SELECT * " & _
       " FROM " & gcDai_Tour_Dai & _
       " WHERE Id = " & objMdi.info.ID & " AND Type = " & CStr(gcDAI_NORMAL) & _
       " ORDER BY Date ASC"

    CRecord_RstSQL = m_CRecord.RstSQL(sQuery)
End Function

Public Function CRecord_SetField(ByVal sFieldName As String, ByVal sFieldValue As Variant) As Boolean
    CRecord_SetField = m_CRecord.SetField(sFieldName, sFieldValue)
End Function

Public Function CRecord_Update() As Boolean
    CRecord_Update = m_CRecord.Update
End Function

