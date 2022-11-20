VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExportActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'local variable(s) to hold property value(s)
Private mvarType_ID As Long 'local copy
Private mvarPosition As Long 'local copy
Private mvarDescription As String 'local copy
Private mvarColour As String 'local copy
Private mvarDate As Date 'local copy
Public Property Let ActivityDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Date = 5
    mvarDate = vData
End Property


Public Property Get ActivityDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Date
    Date = mvarDate
End Property



Public Property Let Colour(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Colour = 5
    mvarColour = vData
End Property


Public Property Get Colour() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Colour
    Colour = mvarColour
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let Position(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Position = 5
    mvarPosition = vData
End Property


Public Property Get Position() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Position
    Position = mvarPosition
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



