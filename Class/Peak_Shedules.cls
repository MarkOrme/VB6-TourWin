VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPeak_Schedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Peak_Schedule"
Attribute VB_Ext_KEY = "Member0" ,"Peak_Schedule"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarOverWrite As Boolean 'local copy
Private mvarFileLocationAndName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarDBHandle As Long 'local copy

Public Property Get FileLocationAndName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileLocationAndName
    FileLocationAndName = mvarFileLocationAndName
End Property


Public Property Get OverWrite() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OverWrite
    OverWrite = mvarOverWrite
End Property

Public Function Create_Export_File() As Boolean
End Function

Public Function FileExist() As Boolean
    If "" = Dir$(mvarFileLocationAndName) Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function
Public Function GetFileAttributes(ByVal bOverWrite As Boolean) As Boolean
' ------------------------------------------------------------------
' Get Export File Location
' Return fail if user cancels
' ------------------------------------------------------------------
On Local Error GoTo GetFileAttributes_Error

' -----

Dim sOpen           As SelectedFile
Dim Count           As Integer
Dim FileList        As String
      
    
' See Standard CommonDialog Flags for all options
FileDialog.sFilter = "Export (*." & gcFILE_EXTENTION & ")" & Chr$(0) & "*." & gcFILE_EXTENTION
FileDialog.sFile = "*." & gcFILE_EXTENTION
FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
FileDialog.sDlgTitle = "Open"
FileDialog.sInitDir = App.Path & "\"

sOpen = ShowOpen(MDI.hWnd)

If Err.Number <> 32755 And sOpen.bCanceled = False Then
    FileAttributes bOverWrite, FileDialog.sFile
    GetFileAttributes = True
End If
        
    
Exit Function
GetFileAttributes_Error:
If 20477 = Err.Number Then
    FileDialog.sFile = ""
End If
GetFileAttributes = False
End Function
Public Function FileAttributes(ByVal bOverWrite As Boolean, ByVal sFileDirAndName As String) As Boolean
' =======================================================================
' FileAttributes - Make sure that mvarFileLocation is valid
'
' =======================================================================
Dim sFile As String
Dim lfound As Long
Dim bFound As Boolean
    mvarOverWrite = bOverWrite
    ' -------------------------------------------
    ' Will assume that the only check that needs
    ' to be performed is file name, directory
    ' should be good.
    ' -------------------------------------------
    ' Strip of directory and work with file name
    For lfound = Len(sFileDirAndName) To 1 Step -1
            If "\" = Mid(sFileDirAndName, lfound, 1) Then
                sFile = Mid(sFileDirAndName, lfound + 1)
                mvarFileLocationAndName = Mid(sFileDirAndName, 1, lfound)
                Exit For
            End If
    Next lfound
    
    ' Make sure file has correct extension
    bFound = False
    For lfound = 1 To Len(sFile)
            If "." = Mid(sFile, lfound, 1) Then
                sFile = Mid(sFile, 1, lfound) & gcFILE_EXTENTION
                bFound = True
                Exit For
            End If
    Next lfound
    If Not bFound Then
        sFile = sFile & "." & gcFILE_EXTENTION
    End If
    mvarFileLocationAndName = mvarFileLocationAndName & sFile

End Function

Public Function Export(ByVal lType_Id As Long, ByVal sName As String) As Boolean
Dim iLoop As Integer
Dim cTempPeakSched As cPeak_Schedule
' ------------------------------------------------------------------
' Export - This function is used to export flag with schedule are
'          to be exported...
' ------------------------------------------------------------------
Select Case lType_Id
       Case gcActive_Type_PeakNames:
       
        Set cTempPeakSched = Item(sName)
        If Not cTempPeakSched Is Nothing Then
            cTempPeakSched.Marked_For_Export = True
            Set cTempPeakSched = Nothing
            Export = True
        End If

       Case gcActive_Type_EventNames:
       Case gcActive_Type_HeartNames:
End Select
End Function

Public Function GetListOfItems(ByVal lType_Id As Long, oControl As Object) As Boolean
' ========================================================
' GetListOfItems - Check if collection is loaded, if not
'                   then load collection and added
'                   items to listbox
' ========================================================
Dim cTempClass As cPeak_Schedule

If cPeakSchedules.Count = 0 Then
Dim iTempHandle As Long
Dim sSQL As String
    iTempHandle = ObjTour.GetHandle
    sSQL = "SELECT * FROM " & gcPeakTour_Peaks & " WHERE ID = " & objMdi.info.ID
    ObjTour.RstSQL iTempHandle, sSQL
    ObjTour.DBMoveFirst iTempHandle
    With ObjTour
        Do While Not .EOF(iTempHandle)
            ' Add Schedule to collection
            Add .DBGetField(gcPEAK_NAME, iTempHandle), .DBGetField(gcPEAK_SCHED, iTempHandle), .DBGetField(PEAKS_LENGTH, iTempHandle), lType_Id, .DBGetField(gcPEAK_NAME, iTempHandle)
            .DBMoveNext iTempHandle
        Loop
        .FreeHandle iTempHandle
    End With
End If

    ' Add items to list box
    If Not oControl Is Nothing Then
        oControl.Clear
        For Each cTempClass In cPeakSchedules
            oControl.AddItem cTempClass.Peak_Name
        Next
    End If
End Function

Public Function Add(Peak_Name As String, Peak_Schedule As String, Peak_Length As Long, Type_ID As Long, Optional sKey As String) As cPeak_Schedule
    'create a new object
    Dim objNewMember As cPeak_Schedule
    Set objNewMember = New cPeak_Schedule


    'set the properties passed into the method
    objNewMember.Peak_Name = Peak_Name
    objNewMember.Peak_Schedule = Peak_Schedule
    objNewMember.Peak_Length = Peak_Length
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

Public Property Get Item(vntIndexKey As Variant) As cPeak_Schedule
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


Public Function DoExport(ByVal lType As Long) As Boolean
' -------------------------------------------
' DoExport of Selected items
' -------------------------------------------
' Check if File exist, create if not
' ----------------------------------
Dim sUpdatedSchedule As String  ' Holds update schedule is changed when adding to export

On Local Error GoTo DoExport_Error
If Not FileExist Then
    If Not ObjTour.CreateExportDB(mvarFileLocationAndName) Then
        DoExport = False
        Exit Function
    End If
End If

    ' Reserved rst for Export...
    'EXPORT_HANDLE
    If Not cExportFile.OpenExportDB(mvarFileLocationAndName, gcEXPORT_SCHEDULE_TABLE, EXPORT_HANDLE) Then
        DoExport = False
        Exit Function
    End If

Select Case lType

    Case gcActive_Type_PeakNames:
        Dim cTempClass As cPeak_Schedule
            ' --------------------------------------------
            ' Loop thru each cPeakSchedule, and for each
            ' item that is marked for export, do so.
            ' --------------------------------------------
            For Each cTempClass In cPeakSchedules
                If cTempClass.Marked_For_Export Then
                   cExportFile.AddEachScheduleToExport cTempClass, sUpdatedSchedule
                   ' Pass the return from above
                    cExportFile.AddSchedulesToExport cTempClass, sUpdatedSchedule
                End If
            Next
    Case gcActive_Type_EventNames:
    
    Case gcActive_Type_HeartNames:
End Select


    ' Close Export File
    cExportFile.CloseExportDB EXPORT_HANDLE

    DoExport = True
    
Exit Function
DoExport_Error:
    'MsgBox Err.Description
End Function
'---------------------------------------------------------------------------------------
' PROCEDURE : DoImport
' DATE      : 4/11/03 13:36
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function DoImport() As Boolean
On Local Error GoTo DoImport_Error

'Declare local variables
Dim bResult As Boolean
Dim CImportFile As CExport
Dim CPosMapping As CInOut
Dim sSQL As String
Dim lCount As Long
Dim lCurrentRecPos As Long
Dim Criteria As String
Dim lSchedLen As Long
Dim sTemp As String

'Attempt to open Import File ...
Set CImportFile = New CExport
Set CPosMapping = New CInOut

If CImportFile.OpenExportDB(Me.FileLocationAndName, gcEXPORT_ACTIVITY_TABLE, EXPORT_HANDLE) Then
    ' -------------------
    ' Load Peak Schedule
    ' -------------------
    sSQL = "SELECT * FROM " & gcEXPORT_ACTIVITY_TABLE & _
           " WHERE " & gcEXPORT_TYPE_ID_FIELD & " = " & gcActive_Type_PeakNames & " ORDER BY " & gcEXPORT_POS_FIELD & " ASC"
           
    ObjTour.RstSQL EXPORT_HANDLE, sSQL
    
  If ObjTour.RstRecordCount(EXPORT_HANDLE) <> 0 Then
    ' ------------------
    ' Setup progress bar
    ' ------------------
    ObjTour.DBMoveLast (EXPORT_HANDLE)
    lCount = ObjTour.RstRecordCount(EXPORT_HANDLE)
    lCurrentRecPos = 1
    ProgressBar "Loading New Descriptions...", -1, (lCurrentRecPos / lCount) * 10, -1
    
    ObjTour.DBMoveFirst (EXPORT_HANDLE)
    ' Set ActivityNames to the correct TypeID
    cActivityNames.Type_ID = gcActive_Type_PeakNames
    cActivityNames.StartSearch gcActive_Type_PeakNames
    
      Do
        ' Try and find if description already exists...
        If Not cActivityNames.FindItemByName(ObjTour.DBGetField(gcEXPORT_DESCRIPTION_FIELD, EXPORT_HANDLE)) Then
                
          cActivityNames.Add gcActive_Type_PeakNames, _
                         ObjTour.DBGetField(gcEXPORT_DESCRIPTION_FIELD, EXPORT_HANDLE), _
                         ObjTour.DBGetField(gcEXPORT_COLOUR_FIELD, EXPORT_HANDLE)
                         
        End If ' End Find by name
                
        ' Map new Position with export value
        CPosMapping.Add cActivityNames.Position, _
                        ObjTour.DBGetField(gcEXPORT_POS_FIELD, EXPORT_HANDLE), _
                        ObjTour.DBGetField(gcEXPORT_POS_FIELD, EXPORT_HANDLE)
                        
        ObjTour.DBMoveNext (EXPORT_HANDLE)
        lCurrentRecPos = lCurrentRecPos + 1
        ProgressBar "Loading New Descriptions...", -1, (lCurrentRecPos / lCount) * 10, -1
        
    Loop While Not ObjTour.EOF(EXPORT_HANDLE)
  End If
    ' Hide Progress Bar...
    ProgressBar "", 0, 0, 0
    
    ' Now copy the schedule names over to Peaks Table
    ' updating the schedule name fields.
    sSQL = "SELECT * FROM " & gcEXPORT_SCHEDULE_TABLE

    ObjTour.RstSQL EXPORT_HANDLE, sSQL

    If ObjTour.RstRecordCount(EXPORT_HANDLE) <> 0 Then
    ' ------------------
    ' Setup progress bar
    ' ------------------
    ObjTour.DBMoveLast (EXPORT_HANDLE)
    lCount = ObjTour.RstRecordCount(EXPORT_HANDLE)
    lCurrentRecPos = 1
    ProgressBar "Loading New Schedules...", -1, (lCurrentRecPos / lCount) * 10, -1

    ObjTour.DBMoveFirst (EXPORT_HANDLE)
    Do
      ' Determine if the Schedule name already exists!!!
      ' --------------------------------------------------
      Criteria = "P_Nam = '" & Trim$(ObjTour.DBGetField(gcEXPORT_NAME_FIELD, EXPORT_HANDLE)) & "'"
      
      ' This checks the current open user database
      sSQL = "SELECT * FROM " & gcPeakTour_Peaks & " WHERE Id = " & objMdi.info.ID
      With ObjTour
        .RstSQL iSearcherDB, sSQL

        .DBFindFirst Criteria, iSearcherDB

    If .NoMatch(iSearcherDB) Then     ' Check if record is found.
        
          .AddNew iSearcherDB 'Connection to User Database
          .DBSetField "Id", objMdi.info.ID, iSearcherDB
          .DBSetField "P_Date", Format$(Now, "mm-dd-yyyy"), iSearcherDB
          .DBSetField "P_Nam", Mid$(.DBGetField(gcEXPORT_NAME_FIELD, EXPORT_HANDLE), 1, 35), iSearcherDB
    Else
        'Prompt user to overwrite
        If vbYes = MsgBox("The import file contains the following peak schedule " & _
                          Chr$(34) & Trim$(ObjTour.DBGetField(gcEXPORT_NAME_FIELD, EXPORT_HANDLE)) & Chr$(34) & vbCrLf & _
                          " which matches an existing schedule. Would you like to update?", vbYesNo, LoadResString(gcTourVersion) & " - Import") Then
         .Edit iSearcherDB
        Else
          GoTo NextSchedule
        End If
    End If
            
    lSchedLen = .DBGetField(gcEXPORT_CYCLE_L_FIELD, EXPORT_HANDLE)
    .DBSetField PEAKS_LENGTH, lSchedLen, iSearcherDB
    
    ' Use CMapping Object to update P_Sched
                        ' Use this constant, the underlying string is correct
      sTemp = .DBGetField(gcEXPORT_SCHEDULE_TABLE, EXPORT_HANDLE)
     .DBSetField "P_Sched", CPosMapping.ReMapSchedule(sTemp, lSchedLen), iSearcherDB
     
     .Update iSearcherDB
NextSchedule:
     .DBMoveNext EXPORT_HANDLE
      lCurrentRecPos = lCurrentRecPos + 1
      ProgressBar "Loading New Schedules...", -1, (lCurrentRecPos / lCount) * 10, -1
    End With
    Loop While Not ObjTour.EOF(EXPORT_HANDLE)
    
 End If ' Ends if recordcount = 0
  
    ObjTour.DBClose (EXPORT_HANDLE)
End If

' Hide Progress Bar...
ProgressBar "", 0, 0, 0

Set CImportFile = Nothing
Set CPosMapping = Nothing

DoImport = bResult

On Error GoTo 0
Exit Function

DoImport_Error:
    If bDebug Then Handle_Err Err, "DoImport-cPeak_Schedules"
    Resume Next


End Function
