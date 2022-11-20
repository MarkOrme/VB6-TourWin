VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPeak_Names"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cPeakName"
Attribute VB_Ext_KEY = "Member0" ,"cPeakName"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarIsUpdated As Boolean 'local copy
Public Property Let IsUpdated(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsUpdated = 5
    mvarIsUpdated = vData
End Property


Public Property Get IsUpdated() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsUpdated
    IsUpdated = mvarIsUpdated
End Property



Public Function Add(Description As String, Colour As String, Position As Long, Type_ID As Long, Optional sKey As String) As cPeakName
    'create a new object
    Dim objNewMember As cPeakName
    Set objNewMember = New cPeakName


    'set the properties passed into the method
    objNewMember.Description = Description
    objNewMember.Colour = Colour
    objNewMember.Position = Position
    objNewMember.Type_ID = Type_ID


    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cPeakName
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    mvarIsUpdated = False
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Public Function Load_Collection()
' -----------------------------------------------------------------
' Load_Collection - Loop through Activity table and load
'                   each Peak Activity Name. These are the
'                   item that make up a schedule...
'
' -----------------------------------------------------------------

cActivityNames.Type_ID = gcActive_Type_PeakNames
' ----------------------------------------------
' Just in case item are already loaded, refresh
' the collection by setting to NOTHING
' ----------------------------------------------
' This code was added because after an import,
' the updated flag is set to false, which forces
' Load_Collection to be called. If collection is
' is already load, duplicates are found... thus
' reset collection by Nothing then re-adding.
' ----------------------------------------------
Set mCol = Nothing
Set mCol = New Collection

If cActivityNames.StartSearch(gcActive_Type_PeakNames) Then
    Do
        With cActivityNames
            Add .Description, .Colour, .Position, .Type_ID, .Position
        End With
    Loop While cActivityNames.GetNext
End If

    mvarIsUpdated = True
    
End Function
