VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CImportExport"
Attribute VB_Ext_KEY = "Member0" ,"CImportExport"
'local variable to hold collection
Private mCol As Collection
Public Function Add(TourWin_Pos As Long, Export_Pos As Long, Optional sKey As String) As CImportExport
    'create a new object
    Dim objNewMember As CImportExport
    Set objNewMember = New CImportExport


    'set the properties passed into the method
    objNewMember.TourWin_Pos = TourWin_Pos
    objNewMember.Export_Pos = Export_Pos


    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CImportExport
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


'---------------------------------------------------------------------------------------
' PROCEDURE : ReMapSchedule
' DATE      : 4/13/03 17:03
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function ReMapSchedule(ByVal sExportSchedule As String, ByVal lSchedLen As Long) As String
On Local Error GoTo ReMapSchedule_Error
'Declare local variables
Dim iSched As Integer
Dim sSchedNam As String
Dim iCurrentPos As Integer
Dim sDeliminator As String
Dim sPosition As String
Dim oMap As CImportExport
Dim i As Long

' Start by setting Found flag to false
For Each oMap In Me
    oMap.AlreadyFound = False
Next

' Do search and replace on each schedule value
sExportSchedule = "," & sExportSchedule
Sched = sExportSchedule
iCurrentPos = 1
sDeliminator = ","

For i = 0 To lSchedLen ' Need an extra loop because the first attempt always fails!
                        ' therefore 0 to Length....
  ' Get Number Column Value
  If GetPostionValue(Sched, iCurrentPos, sDeliminator, sPosition) Then
    ' Try and find Pos value in me
    Set oMap = Me.Item(sPosition)
    If Not oMap Is Nothing Then
      'Check if this position was already updated...
      If Not oMap.AlreadyFound Then
        Do
          sSchedNam = sExportSchedule
          sExportSchedule = Replace(sExportSchedule, "," & sPosition & ",", ",!" & oMap.TourWin_Pos & ",", , vbTextCompare)
        Loop While sSchedNam <> sExportSchedule
        oMap.AlreadyFound = True
      End If
    End If ' End oMap is nothing
  End If ' End GetPos...
Next i

' ---------------------
' Remove ",!" with ","
' ---------------------
Do
    sSchedNam = sExportSchedule
    sExportSchedule = Replace(sExportSchedule, ",!", ",")
Loop While sSchedNam <> sExportSchedule

' ---------------------
' Remove first ","
' ---------------------
ReMapSchedule = Mid$(sExportSchedule, 2)

On Error GoTo 0
Exit Function

ReMapSchedule_Error:
    If bDebug Then Handle_Err Err, "ReMapSchedule-CInOut"
    Resume Next


End Function
