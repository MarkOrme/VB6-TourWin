VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHistorical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_Handle As Long


Public Function DoesHistoryRecordExist(ByVal sDate As String) As Boolean
On Local Error GoTo DoesHistoryRecordExist_Err

Dim lHandle As Long
Dim sSQL As String

'Pesimistic
DoesHistoryRecordExist = False

' Set Recordset
sSQL = " SELECT * " & _
       " FROM " & gcDai_Tour_Dai & _
       " WHERE Id = " & objMdi.info.ID & _
       "   AND Type = " & CStr(gcDAI_HISTORICAL) & _
       "   AND Date = #" & sDate & "# "
       
lHandle = Me.DBHandle
ObjTour.RstSQL lHandle, sSQL
Me.DBHandle = lHandle
If ObjTour.RstRecordCount(Me.DBHandle) > 0 Then DoesHistoryRecordExist = True
On Error GoTo 0
Exit Function

DoesHistoryRecordExist_Err:
Err.Clear
End Function

Public Property Get DBHandle() As Long
    DBHandle = m_Handle
End Property

Public Property Let DBHandle(ByVal vNewValue As Long)
    m_Handle = vNewValue
End Property

Public Function GetHistoricalField(ByVal sFieldName As String, ByRef sValue As String) As Boolean
On Error GoTo GetHistoricalField_Err
    'Optimistic
    GetHistoricalField = True
    sValue = ObjTour.DBGetField(sFieldName, Me.DBHandle)
On Error GoTo 0
Exit Function

GetHistoricalField_Err:
GetHistoricalField = False
End Function

Public Function ExecuteQuery(ByVal sQuery As String) As Boolean
    ObjTour.DBExecute sQuery
End Function

Private Sub Class_Terminate()
    ObjTour.FreeHandle Me.DBHandle
End Sub
