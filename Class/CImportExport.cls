VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CInOut"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarTourWin_Pos As Long 'local copy
Private mvarExport_Pos As Long 'local copy
Private m_Found As Boolean ' Flag
Private mvarCInOut As CInOut


Public Property Get CInOut() As CInOut
    If mvarCInOut Is Nothing Then
        Set mvarCInOut = New CInOut
    End If

    Set CInOut = mvarCInOut
End Property


Public Property Set CInOut(vData As CInOut)
    Set mvarCInOut = vData
End Property
Private Sub Class_Terminate()
    Set mvarCInOut = Nothing
End Sub

Public Property Let Export_Pos(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Export_Pos = 5
    mvarExport_Pos = vData
End Property


Public Property Get Export_Pos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Export_Pos
    Export_Pos = mvarExport_Pos
End Property

Public Property Let TourWin_Pos(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TourWin_Pos = 5
    mvarTourWin_Pos = vData
End Property

Public Property Get TourWin_Pos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TourWin_Pos
    TourWin_Pos = mvarTourWin_Pos
End Property

Public Property Let AlreadyFound(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TourWin_Pos = 5
    m_Found = vData
End Property

Public Property Get AlreadyFound() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TourWin_Pos
    AlreadyFound = m_Found
End Property


