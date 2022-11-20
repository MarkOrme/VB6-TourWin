VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHeartNameVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cHeartNames"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarFieldName As String 'local copy
Private mvarDescription As String 'local copy
Private mvarcHeartNames As cHeartNames
'local variable(s) to hold property value(s)
Private mvarID As Integer 'local copy
Private mvarReport_Name As String 'local copy
Public Property Let Report_Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Report_Name = 5
    mvarReport_Name = vData
End Property


Public Property Get Report_Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Report_Name
    Report_Name = mvarReport_Name
End Property



Public Property Let ID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


Public Property Get ID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property





Public Property Get cHeartNames() As cHeartNames
    If mvarcHeartNames Is Nothing Then
        Set mvarcHeartNames = New cHeartNames
    End If


    Set cHeartNames = mvarcHeartNames
End Property


Public Property Set cHeartNames(vData As cHeartNames)
    Set mvarcHeartNames = vData
End Property
Private Sub Class_Terminate()
    Set mvarcHeartNames = Nothing
End Sub



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
Attribute Description.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let FieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldName = 5
    mvarFieldName = vData
End Property


Public Property Get FieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldName
    FieldName = mvarFieldName
End Property



