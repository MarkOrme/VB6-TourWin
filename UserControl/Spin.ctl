VERSION 5.00
Begin VB.UserControl Spin 
   BackColor       =   &H00000000&
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   ScaleHeight     =   510
   ScaleWidth      =   360
   ToolboxBitmap   =   "Spin.ctx":0000
   Begin VB.VScrollBar varrow 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Spin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Constants & Structures
Const SYSTEMERRORRANGE = 512
Const INVALIDVALUE = vbObjectError + SYSTEMERRORRANGE + 1
Const INVALIDVALUETEXT = "Value does not fall within Min Max property range."

' Variables
Dim m_lValue        As Long
Dim m_Min           As Long
Dim m_Max           As Long
Dim m_Value         As Long
Dim bInitializing   As Boolean

' Events
Public Event Change()

Private Sub UserControl_Initialize()

bInitializing = True
varrow.Value = 1000
m_lValue = 1000  ' Init
bInitializing = False
End Sub

Private Sub varrow_Change()

On Error Resume Next

Dim lTemp       As Long

If bInitializing Then Exit Sub
lTemp = m_lValue - varrow.Value

If (m_Value + lTemp) > m_Max Or _
   (m_Value + lTemp) < m_Min Then
Else
    m_Value = m_Value + lTemp
    RaiseEvent Change
End If

m_lValue = varrow.Value
    
End Sub

Private Sub UserControl_Resize()

On Error Resume Next

' Resize controls
varrow.Move 0, 0, ScaleWidth, ScaleHeight

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : Value
' DATE      : 5/15/04 12:54
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Value() As Long

On Local Error GoTo Value_Error
'Declare local variables

    Value = m_Value

On Error GoTo 0
Exit Property

Value_Error:
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Value
' DATE      : 5/15/04 12:54
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Value(ByVal lValue As Long)

On Local Error GoTo Value_Error
'Declare local variables

    ' ---------------------------------
    ' First, ensure value falls within
    ' specified Min Max range!!!
    ' ---------------------------------
    If lValue > m_Max Or lValue < m_Min Then
        On Error GoTo 0
        Err.Raise INVALIDVALUE, "Spin Control", INVALIDVALUETEXT
        Exit Property
    End If
    
    
    m_Value = lValue
        
On Error GoTo 0
Exit Property

Value_Error:
    Resume Next

End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Min
' DATE      : 5/15/04 13:02
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Min() As Integer

On Local Error GoTo Min_Error
'Declare local variables

    Min = m_Min

On Error GoTo 0
Exit Property

Min_Error:
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Min
' DATE      : 5/15/04 13:02
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Min(ByVal iMin As Integer)

On Local Error GoTo Min_Error
'Declare local variables

  m_Min = iMin

On Error GoTo 0
Exit Property

Min_Error:
    Resume Next

End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Max
' DATE      : 5/15/04 13:03
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Max() As Integer

On Local Error GoTo Max_Error
'Declare local variables

    Max = m_Max

On Error GoTo 0
Exit Property

Max_Error:
    Resume Next
End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Max
' DATE      : 5/15/04 13:03
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Max(ByVal iMax As Integer)

On Local Error GoTo Max_Error
'Declare local variables

    m_Max = iMax

On Error GoTo 0
Exit Property

Max_Error:
    Resume Next

End Property
