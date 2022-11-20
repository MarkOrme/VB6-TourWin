VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PeakName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copy
Private mvarColour As String 'local copy
Private mvarPosition As Long 'local copy
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



