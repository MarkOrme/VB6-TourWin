VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitReprentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cBitReprentation
' DateTime  : 6/26/04 07:08
' Author    : Mark Ormesher
' Purpose   : Encapsulations a long data type and provides methods
'             for retrieving individual bit value.
'---------------------------------------------------------------------------------------
Option Explicit
Const BITRANGE = 16
Dim m_lValue        As Long

'---------------------------------------------------------------------------------------
' PROCEDURE : GetBool
' DATE      : 6/26/04 07:11
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function GetBool(ByVal Bit As Long) As Boolean

On Local Error GoTo GetBool_Error
    If Bit < BITRANGE Then ' Make sure it is an allowable bit value
        GetBool = IsBitSet(m_lValue, Bit)
    Else
        GetBool = False
    End If

On Error GoTo 0
Exit Function

GetBool_Error:
    If bDebug Then Handle_Err Err, "GetBool-cBitReprentation"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : GetValue
' DATE      : 6/26/04 07:11
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function GetValue(Optional ByVal Bit As Variant) As Long

On Local Error GoTo GetValue_Error

If IsMissing(Bit) Then
    GetValue = m_lValue
Else
      If CLng(Bit) < BITRANGE Then ' Make sure it is an allowable bit value
        GetValue = GetBit(m_lValue, CLng(Bit))
      Else
        GetValue = 0
      End If
End If

On Error GoTo 0
Exit Function

GetValue_Error:
    If bDebug Then Handle_Err Err, "GetValue-cBitReprentation"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : SetValue
' DATE      : 6/26/04 07:12
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function SetValue(ByVal lValue As Long, Optional ByVal Bit As Variant) As Long

On Local Error GoTo SetValue_Error
'Declare local variables

If IsMissing(Bit) Then
    m_lValue = lValue
Else
  If CLng(Bit) < BITRANGE Then ' Make sure it is an allowable bit value
    If lValue = 0 Then
        m_lValue = ClearBit(m_lValue, CLng(Bit))
    Else
        m_lValue = SetBit(m_lValue, CLng(Bit))
    End If
  End If
End If

SetValue = m_lValue

On Error GoTo 0
Exit Function

SetValue_Error:
    If bDebug Then Handle_Err Err, "SetValue-cBitReprentation"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : SetBool
' DATE      : 6/26/04 07:12
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function SetBool(ByVal bValue As Boolean, ByVal Bit As Long) As Long

On Local Error GoTo SetBool_Error

If bValue = False Then
    m_lValue = SetValue(0, Bit)
Else
    m_lValue = SetValue(1, Bit)
End If

    SetBool = m_lValue

On Error GoTo 0
Exit Function

SetBool_Error:
    If bDebug Then Handle_Err Err, "SetBool-cBitReprentation"
    Resume Next


End Function


' --------------------------------------------
' Following functions are for bit manipulation
' specifically for iBitFlags...
' --------------------------------------------
Private Function SetBit(InByte As Long, Bit As Long) As Long
'Set het n'de bit of van InInteger

SetBit = InByte Or (2 ^ Bit)  'Set het n'de Bit

End Function
Private Function ClearBit(InByte As Long, Bit As Long) As Long
'Clear het n'de bit of van InByte
   
ClearBit = InByte And Not (2 ^ Bit) 'Clear het n'de Bit

End Function

Private Function IsBitSet(InByte As Long, Bit As Long) As Boolean
'Is het n'de bit van InByte gezet of niet?

IsBitSet = ((InByte And (2 ^ Bit)) > 0)

End Function
Public Function ToggleBit(InByte As Long, Bit As Long) As Integer
'Toggle'ed het n'de van InByte

ToggleBit = InByte Xor (2 ^ Bit)

End Function

Private Function GetBit(ByVal InByte As Long, ByVal Bit As Long) As Long

    GetBit = IIf(IsBitSet(InByte, Bit), 1, 0)

End Function

Private Function MakeBit(ByRef InByte As Long, ByVal Bit As Long, ByVal Value As Long) As Long


    Select Case Value
            Case 0:
                MakeBit = ClearBit(InByte, Bit)
            Case 1:
                MakeBit = SetBit(InByte, Bit)
    End Select
    
End Function

