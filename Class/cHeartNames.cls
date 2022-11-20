VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHeartNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cHeartNameVar"
Attribute VB_Ext_KEY = "Member0" ,"cHeartNameVar"
'local variable to hold collection
Private mCol As Collection
Public Function LoadValues(ByVal iID As Integer) As Boolean

Dim iLoop As Integer, HeartStr As String, RetStr As String

'Be Optimistic
LoadValues = True

If mCol.Count <> 0 Then Exit Function

    For iLoop = 1 To 9
        HeartStr = "Heart" & Format$(iLoop, "0")
        RetStr = Get_NameTour_HeartNames(HeartStr)
        
            If RetStr <> "No Return" Then
                Add HeartStr, RetStr, "HeaV" & Format$(iLoop, "0") & "_Title='" & RetStr & "'", iID, iLoop
            Else
                Add HeartStr, "", "HeaV" & Format$(iLoop, "0") & "_Title='" & RetStr & "'", iID, iLoop
            End If
    
    Next iLoop
    
End Function

Public Function Add(ByVal FieldName As String, ByVal Description As String, ByVal sReportName As String, ByVal iID As Integer, Optional ByVal sKey As String) As cHeartNameVar
    'create a new object
    Dim objNewMember As cHeartNameVar
    Set objNewMember = New cHeartNameVar


    'set the properties passed into the method
    objNewMember.FieldName = FieldName
    objNewMember.Description = Description
    objNewMember.Report_Name = sReportName
    objNewMember.ID = iID
    Set objNewMember.cHeartNames = cHeartNames



    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cHeartNameVar
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  On Local Error Resume Next
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

