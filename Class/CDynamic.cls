VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDynamic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The key for this class is Position ID."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarContainsTag As Boolean 'local copy
Private mvarTagBegin As Integer 'local copy
Private mvarTagEnd As Integer 'local copy
Private mvarLastValue As Currency 'local copy
Private mvarOperation As String 'local copy
Private m_GrowthValue As Currency

Public Function GetNextValue(ByVal sDesc As String) As String
Dim sNewValue As Currency
Dim sTemp As String

If Not ContainsTag Then
    GetNextValue = sDesc
Else
    If "Add" = Operation Then
        LastValue = LastValue + GrowthValue
    Else
        LastValue = LastValue + (LastValue * (GrowthValue / 100))
    End If

    sTemp = Mid$(sDesc, 1, TagBegin - 1) & CStr(LastValue) & Mid$(sDesc, TagEnd + 2)
    GetNextValue = sTemp
    
End If
End Function

Public Function GetCurrentValue(ByVal sDesc As String) As String
Dim sTemp As String

If Not ContainsTag Then
    GetCurrentValue = sDesc
Else

sTemp = Mid$(sDesc, 1, TagBegin - 1) & CStr(LastValue) & Mid$(sDesc, TagEnd + 2)
GetCurrentValue = sTemp
    
End If
End Function

Public Property Let Operation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Operation = 5
    mvarOperation = vData
End Property


Public Property Get Operation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Operation
    Operation = mvarOperation
End Property



Public Property Let LastValue(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastValue = 5
    mvarLastValue = vData
End Property


Public Property Get LastValue() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastValue
    LastValue = mvarLastValue
End Property



Public Property Let TagEnd(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TagEnd = 5
    mvarTagEnd = vData
End Property


Public Property Get TagEnd() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TagEnd
    TagEnd = mvarTagEnd
End Property



Public Property Let TagBegin(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TagBegin = 5
    mvarTagBegin = vData
End Property


Public Property Get TagBegin() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TagBegin
    TagBegin = mvarTagBegin
End Property



Public Property Let ContainsTag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContainsTag = 5
    mvarContainsTag = vData
End Property


Public Property Get ContainsTag() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContainsTag
    ContainsTag = mvarContainsTag
End Property


Public Property Get GrowthValue() As Currency
    GrowthValue = m_GrowthValue

End Property

Public Property Let GrowthValue(ByVal vNewValue As Currency)
    m_GrowthValue = vNewValue
End Property
