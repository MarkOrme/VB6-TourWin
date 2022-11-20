VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPeak_Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Peak_Schedule"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarPeak_Name As String 'local copy
Private mvarPeak_Schedule As String 'local copy
Private mvarPeak_Length As Long 'local copy
Private mvarType_ID As Long 'local copy
Private mvarMarkedForExport As Boolean

Public Property Let Marked_For_Export(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Type_ID = 5
    mvarMarkedForExport = vData
End Property


Public Property Get Marked_For_Export() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Type_ID
    Marked_For_Export = mvarMarkedForExport
End Property

Public Property Let Type_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Type_ID = 5
    mvarType_ID = vData
End Property


Public Property Get Type_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Type_ID
    Type_ID = mvarType_ID
End Property



Public Property Let Peak_Length(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Peak_Length = 5
    mvarPeak_Length = vData
End Property


Public Property Get Peak_Length() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Peak_Length
    Peak_Length = mvarPeak_Length
End Property



Public Property Let Peak_Schedule(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Peak_Schedule = 5
    mvarPeak_Schedule = vData
End Property


Public Property Get Peak_Schedule() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Peak_Schedule
    Peak_Schedule = mvarPeak_Schedule
End Property



Public Property Let Peak_Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Peak_Name = 5
    mvarPeak_Name = vData
End Property


Public Property Get Peak_Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Peak_Name
    Peak_Name = mvarPeak_Name
End Property



