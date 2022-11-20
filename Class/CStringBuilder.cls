VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Improves string concatenation by keeping track of the lenght of the current string."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias _
"RtlMoveMemory" (Destination As Any, Source As Any, _
ByVal Length As Long)

'local variable(s) to hold property value(s)
Private m_Bump As Long 'local copy
Private m_Len As Long 'local copy
Private m_BufferLen As Long
Private m_Buffer() As Byte

Private Const defBufferBump As Long = 2048

Public Function ToString() As String

    If m_Len > 0 Then
        ToString = Space$(m_Len \ 2)
        
        Call CopyMemory(ByVal StrPtr(ToString), _
        m_Buffer(0), m_Len)

    End If
End Function

Public Function EnsureCapacity(ByVal NewVal As Long) As Long
If NewVal > m_BufferLen Then
    Me.Capacity = NewVal
End If
End Function

Public Function Append(NewText As String) As CStringBuilder

Static LenNewText As Long

LenNewText = LenB(NewText)
If LenNewText Then
    Me.EnsureCapacity LenNewText + m_Len
    
    Call CopyMemory(m_Buffer(m_Len), ByVal _
    StrPtr(NewText), LenNewText)
    
    m_Len = m_Len + LenNewText
End If

Set Append = Me
End Function

Public Property Let Length(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Length = 5
If vData > 0 Then
    If (vData * 2) > m_BufferLen Then
        Me.Capacity = vData * 2

    End If
    
    m_Len = vData * 2

End If
End Property



Public Property Get Length() As Long
    Length = m_Len \ 2
End Property



Public Property Let Capacity(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Capacity = 5
Dim x As Long
    If vData > 0 Then
        x = Int(vData / defBufferBump) + 1
        ReDim Preserve m_Buffer(0 To ((x * defBufferBump) - 1)) As Byte
        m_BufferLen = UBound(m_Buffer) + 1
    End If
End Property




Public Property Get Capacity() As Long
Capacity = m_BufferLen

End Property



Public Property Let BufferBump(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BufferBump = 5
If vData > 0 Then
    If vData Mod 2 Then
        vData = vData + 1
    End If
    
    m_Bump = vData
End If
End Property


Public Property Get BufferBump() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BufferBump
    BufferBump = m_Bump
End Property



Private Sub Class_Initialize()
m_Bump = defBufferBump
m_Len = 0
ReDim m_Buffer(0 To m_Bump - 1) As Byte
m_BufferLen = UBound(m_Buffer) + 1
End Sub
