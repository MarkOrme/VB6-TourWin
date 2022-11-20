VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CExportActivity"
Attribute VB_Ext_KEY = "Member0" ,"CExportActivity"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Private mvarDBHandle As Long
Public Function Add(Type_ID As Long, Position As Long, Description As String, Colour As String, ExportDate As Date, Optional sKey As String) As CExportActivity
    'create a new object
    Dim objNewMember As CExportActivity
    Set objNewMember = New CExportActivity


    'set the properties passed into the method
    objNewMember.Type_ID = Type_ID
    objNewMember.Position = Position
    objNewMember.Description = Description
    objNewMember.Colour = Colour
    objNewMember.ActivityDate = Date




    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function
Public Function LoadExportActivities() As Boolean

' ==================================================
' AddSchedulesToExport - This function adds the actual
'                       main schedule to export file
'
' dbs10 must already be set...
' ==================================================

On Local Error GoTo LoadExportActivities_Error
Dim sSQL As String
Dim sPosition As String
Dim sType As String

sSQL = "Select * FROM " & gcEXPORT_ACTIVITY_TABLE
ObjTour.RstSQL EXPORT_HANDLE, sSQL
With ObjTour
    Do While Not .EOF(EXPORT_HANDLE)
    
        sType = .DBGetField(gcEXPORT_TYPE_ID_FIELD, EXPORT_HANDLE)
        sPosition = .DBGetField(gcEXPORT_POS_FIELD, EXPORT_HANDLE)
        
        Add CLng(sType), CLng(sPosition), .DBGetField(gcEXPORT_DESCRIPTION_FIELD, EXPORT_HANDLE), .DBGetField(gcEXPORT_COLOUR_FIELD, EXPORT_HANDLE), CDate("01-01-2000"), Trim$(sType) & "!" & Trim$(sPosition)
        .DBMoveNext EXPORT_HANDLE
    Loop
End With
Exit Function
LoadExportActivities_Error:
    Resume Next
End Function


Public Property Get Item(vntIndexKey As Variant) As CExportActivity
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

Public Function Add_Activity_To_Export(ByVal cActivity As cActivity_Names, Optional lPos As Long) As Boolean
On Local Error GoTo AddSchedulesToExport_Error
Dim sSQL As String
sSQL = "Select * FROM " & gcEXPORT_ACTIVITY_TABLE

With ObjTour
    .RstSQL EXPORT_HANDLE, sSQL

    .AddNew EXPORT_HANDLE
    .DBSetField gcEXPORT_DESCRIPTION_FIELD, cActivity.Description, EXPORT_HANDLE
    .DBSetField gcEXPORT_TYPE_ID_FIELD, cActivity.Type_ID, EXPORT_HANDLE
    If lPos > 0 Then
        .DBSetField gcEXPORT_POS_FIELD, lPos, EXPORT_HANDLE
    Else
        .DBSetField gcEXPORT_POS_FIELD, cActivity.Position, EXPORT_HANDLE
    End If

    .DBSetField gcEXPORT_COLOUR_FIELD, cActivity.Colour, EXPORT_HANDLE
    .Update EXPORT_HANDLE
End With

    Add_Activity_To_Export = True
    
Exit Function
AddSchedulesToExport_Error:
    Add_Activity_To_Export = False
    Resume Next

End Function
Public Function OpenExportDB(ByVal sLocation As String, ByVal sTable As String, ByVal lDBHandle As Long) As Boolean
    If ObjTour.DBOpen(sLocation, sTable, lDBHandle) Then
        OpenExportDB = True
    Else
        OpenExportDB = False
    End If
End Function

Public Function CloseExportDB(ByVal lDBHandle As Long) As Boolean
    If ObjTour.DBClose(lDBHandle) Then
        CloseExportDB = True
    Else
        CloseExportDB = False
    End If
End Function
Public Function AddSchedulesToExport(ByVal cSchedule As cPeak_Schedule, Optional ByVal sUpdateSchedule As String) As Boolean
' ==================================================
' AddSchedulesToExport - This function adds the actual
'                       main schedule to export file
' ==================================================

On Local Error GoTo AddSchedulesToExport_Error
Dim sSQL As String

sSQL = "Select * FROM " & gcEXPORT_SCHEDULE_TABLE

With ObjTour
    .RstSQL EXPORT_HANDLE, sSQL

    .AddNew EXPORT_HANDLE
    If sUpdateSchedule <> "" Then
        .DBSetField gcEXPORT_SCHEDULE_TABLE, sUpdateSchedule, EXPORT_HANDLE
    Else
        .DBSetField gcEXPORT_SCHEDULE_TABLE, cSchedule.Peak_Schedule, EXPORT_HANDLE
    End If
    .DBSetField gcEXPORT_TYPE_ID_FIELD, CLng(gcActive_Type_PeakNames), EXPORT_HANDLE
    .DBSetField gcEXPORT_NAME_FIELD, cSchedule.Peak_Name, EXPORT_HANDLE
    .DBSetField gcEXPORT_CYCLE_L_FIELD, cSchedule.Peak_Length, EXPORT_HANDLE
    .Update EXPORT_HANDLE
End With
    AddSchedulesToExport = True
    
Exit Function
AddSchedulesToExport_Error:
    Resume Next
End Function

Public Function AddEachScheduleToExport(CScheduleClass As cPeak_Schedule, ByRef sReturnString As String) As Boolean
' ==================================================
' AddEachScheduleToExport - This function adds the
'                       individual schedule activities
'                       to Export database.
'
' Business Rule:
'       Before adding item to Activity Table check for
'       Position and Name match, if no match then
'       Check if Position value is available.
' ==================================================
Dim iLoop As Long
Dim sDeliminator As String
Dim sPosition As String
Dim iCurrentPos As Integer
Dim CTempExport As CExportActivity
Dim CTempActivity As CExportActivity
Dim lMaxPos As Long
Dim CPositions As CInOut
Dim CPosition As CImportExport

' Set default values
sDeliminator = ","
sPosition = ""
iCurrentPos = 1
sReturnString = ""
Set CPositions = New CInOut

If cExportFile.Count = 0 Then
    cExportFile.LoadExportActivities
End If
' ---------------------------------------
' Loop through schedules and add each
' schedule description and name...
' ---------------------------------------
For iLoop = 1 To CScheduleClass.Peak_Length

' Get Number Column Value
 GetPostionValue CScheduleClass.Peak_Schedule, iCurrentPos, sDeliminator, sPosition
 
Set CPosition = CPositions(sPosition)

If Not CPosition Is Nothing Then
' -------------------------------
' If CPosition is set, then
' Position has already been set
' -------------------------------
' Make up Schedule string
 sReturnString = sReturnString & Trim$(Str$(CPosition.Export_Pos)) & sDeliminator
Else
' ----------------------------------
' Find Activity record in Tourwin DB
' ----------------------------------
cActivityNames.Type_ID = gcActive_Type_PeakNames
If cActivityNames.FindItemByPosition(CLng(sPosition)) Then

  ' -------------------------------------------------
  ' Check if Export Activity already has this record
  ' -------------------------------------------------
  Set CTempActivity = cExportFile(Trim$(cActivityNames.Type_ID) & "!" & Trim$(cActivityNames.Position))
  ' ----------------------------------
  ' If CTempActivity is set, then Position
  ' was found in Export Activity File
  ' ----------------------------------
  If CTempActivity Is Nothing Then
      cExportFile.Add_Activity_To_Export cActivityNames
      
        With cActivityNames
            cExportFile.Add CLng(.Type_ID), CLng(.Position), .Description, .Colour, CDate("01-01-2000"), Trim$(.Type_ID) & "!" & Trim$(.Position)
            CPositions.Add .Position, .Position, .Position
            sReturnString = sReturnString & Trim$(Str$(.Position)) & sDeliminator
        End With
  Else
      ' ------------------------------------------
      ' If Description match, then no need
      ' to add Activity to Export->Activity table
      ' ------------------------------------------
      If CTempActivity.Description = cActivityNames.Description Then
        ' This Activity already exists
        With cActivityNames
            CPositions.Add .Position, .Position, .Position
            sReturnString = sReturnString & Trim$(Str$(.Position)) & sDeliminator
        End With
      Else
        'If cExportFile.FindItemByName(cActivityNames.Description) Then
        ' ---------------------------------------------------
        ' Position and Type_ID exist, but
        ' description is different. This
        ' occurs if the export file is being
        ' update after being created by another
        ' tourwin user.
        ' 1. Need to find first available Position number
        ' 2. Update schedule with new position values
        ' ---------------------------------------------------
        'GetMaxPosition
        lMaxPos = cExportFile.GetMaxPosition(cActivityNames.Type_ID)
        CPositions.Add cActivityNames.Position, lMaxPos, cActivityNames.Position
        
        sReturnString = sReturnString & Trim$(Str$(lMaxPos)) & sDeliminator
        
        cExportFile.Add_Activity_To_Export cActivityNames, lMaxPos
        
      End If

  End If
End If
End If
Next iLoop
If Not CPosition Is Nothing Then Set CPosition = Nothing
If Not CPositions Is Nothing Then Set CPositions = Nothing
End Function

Public Function GetMaxPosition(ByVal lType As Long) As Long
Dim sSQL As String
sSQL = "Select Max(Position) as MaxPosition FROM " & gcEXPORT_ACTIVITY_TABLE & _
       " WHERE " & gcEXPORT_TYPE_ID_FIELD & " = " & Str$(lType)

ObjTour.RstSQL EXPORT_HANDLE, sSQL

GetMaxPosition = Val(ObjTour.DBGetField("MaxPosition", EXPORT_HANDLE)) + 1

End Function
