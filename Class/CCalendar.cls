VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarMaxDate As Date 'local copy
Private m_Parent As Long

Public Function ShowCalendar(ByVal lTop As Long, ByVal lLeft As Long, ByRef frmParent As Form) As Long

    frmDate.MaxDate = Me.MaxDate
    frmDate.UpdateCalendar Me.MaxDate
    frmDate.Top = lTop
    frmDate.Left = lLeft
    frmDate.Show vbModal
    ShowCalendar = frmDate.SelectionResult
    MaxDate = frmDate.MaxDate
    Unload frmDate
End Function

Public Property Let MaxDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxDate = 5
    mvarMaxDate = vData
End Property


Public Property Get MaxDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxDate
    MaxDate = mvarMaxDate
End Property



