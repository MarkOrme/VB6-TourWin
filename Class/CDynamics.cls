VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDynamics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contains a collection of CDynamic Class objects"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CDynamic"
Attribute VB_Ext_KEY = "Member0" ,"CDynamic"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private Const cPERCENTAGE = "Percentage"
Private Const cADD = "Add"

Public Function Add(ByVal sDesc As String, ByVal ContainsTag As Boolean, TagBegin As Integer, TagEnd As Integer, Optional sKey As String) As CDynamic
    'create a new object
    Dim objNewMember As CDynamic
    Dim sRetStr As String
    
    Set objNewMember = New CDynamic
    
If Not ContainsTag Then
    
    objNewMember.ContainsTag = ContainsTag
    objNewMember.TagBegin = TagBegin
    objNewMember.TagEnd = TagEnd
    objNewMember.LastValue = 0
    objNewMember.Operation = "None"
    objNewMember.GrowthValue = 0
        
Else
    ' Get Base Value
    sRetStr = ""
    Do While "" = sRetStr
        sRetStr = InputBox$("Enter numeric starting base value for: " & vbCrLf & sDesc, "Base value for peak schedule", "100")
        
        If Not IsNumeric(sRetStr) Then sRetStr = ""
            
    Loop
    
    'set the properties passed into the method
    objNewMember.ContainsTag = ContainsTag
    objNewMember.TagBegin = TagBegin
    objNewMember.TagEnd = TagEnd
    objNewMember.LastValue = Val(sRetStr)
    
    
    If "%" = Mid$(sDesc, TagEnd - 1, 1) Then
        objNewMember.Operation = cPERCENTAGE
    Else
        objNewMember.Operation = cADD
    End If
    
    'Get growth value - if percentage omit %
    If cPERCENTAGE = objNewMember.Operation Then
        objNewMember.GrowthValue = Val(Mid$(sDesc, TagBegin + 1, (TagEnd - 1) - (TagBegin + 1)))
    Else
        objNewMember.GrowthValue = Val(Mid$(sDesc, TagBegin + 1, (TagEnd - 1) - (TagBegin)))
    End If
End If

If Len(sKey) = 0 Then
    mCol.Add objNewMember
Else
    mCol.Add objNewMember, sKey
End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CDynamic
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
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

