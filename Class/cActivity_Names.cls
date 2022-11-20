VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActivity_Names"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Activity Names class tracks Peak Names, Event Names and Heart Level Names."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"PeakNames"
'local variable(s) to hold property value(s)
Private mvarType_ID As Long 'local copy
Private mvarActivityHandle As Long
'local variable(s) to hold property value(s)
Private mvarPeakNames_Form As Object 'local copy
Private mvarEventName_Form As Object 'local copy
Private mvarHeartNames_Form As Object 'local copy
'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copy
Private mvarColour As String 'local copy
Private mvarPosition As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarMaxPosition As Long 'local copy
Public Property Let MaxPosition(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxPosition = 5
    mvarMaxPosition = vData
End Property


Public Property Get MaxPosition() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxPosition
    MaxPosition = mvarMaxPosition
End Property

Public Property Let Type_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Type_ID = 5
    mvarType_ID = vData
End Property
Public Property Get Type_ID() As Long
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Type_ID = 5
    Type_ID = mvarType_ID
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




Public Property Set HeartNames(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.HeartNames = Form1
    Set mvarHeartNames = vData
End Property


Public Property Get HeartNames() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeartNames
    Set HeartNames = mvarHeartNames
End Property



Public Property Set EventNames(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.EventNames = Form1
    Set mvarEventNames = vData
End Property


Public Property Get EventNames() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EventNames
    Set EventNames = mvarEventNames
End Property



Public Property Set PeakNames(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.PeakNames = Form1
    Set mvarPeakNames = vData
End Property


Public Property Get PeakNames() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PeakNames
    Set PeakNames = mvarPeakNames
End Property



Public Property Set HeartNames_Form(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.HeartNames_Form = Form1
    Set mvarHeartNames_Form = vData
End Property


Public Property Get HeartNames_Form() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeartNames_Form
    Set HeartNames_Form = mvarHeartNames_Form
End Property



Public Property Set EventName_Form(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.EventName_Form = Form1
    Set mvarEventName_Form = vData
End Property


Public Property Get EventName_Form() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EventName_Form
    Set EventName_Form = mvarEventName_Form
End Property



Public Property Set PeakNames_Form(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.PeakNames_Form = Form1
    Set mvarPeakNames_Form = vData
End Property


Public Property Get PeakNames_Form() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PeakNames_Form
    Set PeakNames_Form = mvarPeakNames_Form
End Property



Public Function ShowForm() As Boolean
' ------------------------------
' If form is not already loaded,
' then load form
' ------------------------------
On Local Error GoTo ShowForm_Error

Select Case mvarType_ID
    Case gcActive_Type_PeakNames:
    
        If mvarPeakNames_Form Is Nothing Then
            Set mvarPeakNames_Form = New frmActivity
        End If
        
        mvarPeakNames_Form.Show vbModal, MDI
        Set mvarPeakNames_Form = Nothing

    Case gcActive_Type_EventNames:
    
        If mvarEventName_Form Is Nothing Then
            Set mvarEventName_Form = New frmActivity
        End If
        
        mvarEventName_Form.Show vbModal, MDI
        Set mvarEventName_Form = Nothing
        
    Case gcActive_Type_HeartNames:
    
        If mvarHeartNames_Form Is Nothing Then
            Set mvarHeartNames_Form = New frmActivity
        End If
        
        mvarHeartNames_Form.Show vbModal, MDI
        Set mvarHeartNames_Form = Nothing
            
End Select
Exit Function
ShowForm_Error:
    If bDebug Then Handle_Err Err, "ShowForm-cActivity_Names"
    Resume Next
End Function

Public Function Delete(ByVal lType_Id As Long, ByVal sDecription As String) As Boolean

On Local Error GoTo Delete_Error
Delete = True
    If FindItemByName(sDecription) Then
        ObjTour.Delete mvarActivityHandle
    End If
Exit Function
Delete_Error:
    Err.Clear
    Delete = False
End Function

Public Function Add(ByVal lType_Id As Long, ByVal sDescription As String, ByVal sColour As String) As Boolean
' -------------------------------------------
' Find current Maximum Position value then
' insert record with Max +1
' --------------------------------------------
On Local Error GoTo Add_Error
ObjTour.AddNew mvarActivityHandle
    ObjTour.DBSetField gcActive_ID, objMdi.info.ID, mvarActivityHandle
    ObjTour.DBSetField gcActive_Type, lType_Id, mvarActivityHandle
    ObjTour.DBSetField gcActive_Des, Left(sDescription, 100), mvarActivityHandle
    ObjTour.DBSetField gcActive_Colour, Left(sColour, 15), mvarActivityHandle
    If lType_Id = gcActive_Type_HeartNames Then
        ObjTour.DBSetField gcActive_Pos, cActivityNames.NextAvailablePosition(gcActive_Type_HeartNames), mvarActivityHandle
    Else
        ObjTour.DBSetField gcActive_Pos, cActivityNames.MaxPosition + 1, mvarActivityHandle
    End If
    
ObjTour.Update mvarActivityHandle

    ' Update Properties
    cActivityNames.Description = Left(sDescription, 100)
    cActivityNames.Colour = Left(sColor, 15)
    ' Update MaxPosition
    cActivityNames.MaxPosition = cActivityNames.MaxPosition + 1
    cActivityNames.Position = cActivityNames.MaxPosition
    
Add = True
Exit Function
Add_Error:
    If bDebug Then Handle_Err Err, "Add-cActivity_Names"
    Resume Next
    
End Function

Public Function GetNext() As Boolean
' ------------------------------------------------------
' GetNext - Move to next record and update properties.
' ------------------------------------------------------
On Local Error GoTo GetNext_Error
    ObjTour.DBMoveNext mvarActivityHandle
    
    If Not ObjTour.EOF(mvarActivityHandle) Then
        mvarDescription = ObjTour.DBGetField(gcActive_Des, mvarActivityHandle)
        mvarColour = ObjTour.DBGetField(gcActive_Colour, mvarActivityHandle)
        mvarPosition = ObjTour.DBGetField(gcActive_Pos, mvarActivityHandle)
        GetNext = True
    Else
        GetNext = False
    End If
Exit Function
GetNext_Error:
    If bDebug Then Handle_Err Err, "GetNext-cActivity_Names"
    Resume Next

End Function

Public Function StartSearch(ByVal lType_Id As Long) As Boolean

Dim sSQL As String
On Local Error GoTo StartSearch_Error

' Get Handle
If mvarActivityHandle = 0 Then
    mvarActivityHandle = ObjTour.GetHandle
End If

' Create SQL Query
sSQL = "SELECT * FROM " & gcActivitiesTable & " WHERE " & gcActive_ID & " = " & _
       objMdi.info.ID & " AND " & gcActive_Type & " = " & lType_Id & " Order By " & gcActive_Pos

' Define recordset
ObjTour.RstSQL mvarActivityHandle, sSQL
ObjTour.DBMoveLast mvarActivityHandle

' Are there any records
If ObjTour.RstRecordCount(mvarActivityHandle) = 0 Then
    StartSearch = False
    mvarMaxPosition = 0
Else
    ' On Last record, set MaxPosition
    mvarMaxPosition = Val(ObjTour.DBGetField(gcActive_Pos, mvarActivityHandle))

    ' Move to first record
    ObjTour.DBMoveFirst mvarActivityHandle
    StartSearch = True
    mvarDescription = ObjTour.DBGetField(gcActive_Des, mvarActivityHandle)
    mvarColour = ObjTour.DBGetField(gcActive_Colour, mvarActivityHandle)
    mvarPosition = ObjTour.DBGetField(gcActive_Pos, mvarActivityHandle)

End If

Exit Function
StartSearch_Error:
    If bDebug Then Handle_Err Err, "StartSearch-cActivity_Names"
    Resume Next

End Function

Public Function Count(ByVal lType_Id As Long) As Long

Dim sSQL As String
On Local Error GoTo Count_Error

' Get Handle
If mvarActivityHandle = 0 Then
    mvarActivityHandle = ObjTour.GetHandle
End If

' Create SQL Query
sSQL = "SELECT * FROM " & gcActivitiesTable & " WHERE " & gcActive_ID & " = " & _
       objMdi.info.ID & " AND " & gcActive_Type & " = " & lType_Id

Debug.Print sSQL

' Define recordset
ObjTour.RstSQL mvarActivityHandle, sSQL
' Return count
Count = ObjTour.RstRecordCount(mvarActivityHandle)

Exit Function
Count_Error:
    If bDebug Then Handle_Err Err, "Count-cActivity_Names"
    Resume Next
    
End Function


Public Function Modify(ByVal lType_Id As Long, ByVal sDescription As String, ByVal sColour As String) As Boolean
Dim sSQL As String

If FindItemByPosition(mvarPosition) Then

    ObjTour.Edit mvarActivityHandle
     ObjTour.DBSetField gcActive_Des, sDescription, mvarActivityHandle
     ObjTour.DBSetField gcActive_Colour, sColour, mvarActivityHandle
    ObjTour.Update mvarActivityHandle
    
    mvarDescription = sDescription
    mvarColour = sColour
    
    Modify = True
    
End If

End Function

Public Function FindItemByName(ByVal sDescription As String) As Boolean

Dim sSQL As String
    
On Local Error Resume Next
sSQL = "SELECT * FROM " & gcActivitiesTable & " WHERE " & gcActive_ID & " = " & _
       objMdi.info.ID & " AND " & gcActive_Type & " = " & mvarType_ID & " AND " & _
       gcActive_Des & " = '" & SyntaxSafe(sDescription) & "'"
Debug.Print sSQL

    ObjTour.RstSQL mvarActivityHandle, sSQL
    If ObjTour.RstRecordCount(mvarActivityHandle) > 0 Then
        mvarDescription = ObjTour.DBGetField(gcActive_Des, mvarActivityHandle)
        mvarColour = ObjTour.DBGetField(gcActive_Colour, mvarActivityHandle)
        mvarPosition = ObjTour.DBGetField(gcActive_Pos, mvarActivityHandle)
        FindItemByName = True
    Else
        FindItemByName = False
    End If
    
End Function

Public Function FindItemByPosition(ByVal lPosition As Long) As Boolean

Dim sSQL As String
    
On Local Error Resume Next
sSQL = "SELECT * FROM " & gcActivitiesTable & " WHERE " & gcActive_ID & " = " & _
       objMdi.info.ID & " AND " & gcActive_Type & " = " & mvarType_ID & " AND " & _
       gcActive_Pos & " = " & lPosition
Debug.Print sSQL

    ObjTour.RstSQL mvarActivityHandle, sSQL
    If ObjTour.RstRecordCount(mvarActivityHandle) > 0 Then
        mvarDescription = ObjTour.DBGetField(gcActive_Des, mvarActivityHandle)
        mvarColour = ObjTour.DBGetField(gcActive_Colour, mvarActivityHandle)
        mvarPosition = ObjTour.DBGetField(gcActive_Pos, mvarActivityHandle)
        mvarType_ID = ObjTour.DBGetField(gcActive_Type, mvarActivityHandle)
        FindItemByPosition = True
    Else
        FindItemByPosition = False
    End If
    
End Function



'---------------------------------------------------------------------------------------
' PROCEDURE : NextAvailablePosition
' DATE      : 4/9/03 19:46
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function NextAvailablePosition(ByVal lTypeID As Long) As Long

On Local Error GoTo NextAvailablePosition_Error
'Declare local variables
Dim lRt As Long
Dim lDBHandle As Long
Dim sSQL As String

'Initialize return value
lRt = 1 ' Start with 1

' Create SQL Query
sSQL = "SELECT * FROM " & gcActivitiesTable & " WHERE " & gcActive_ID & " = " & _
       objMdi.info.ID & " AND " & gcActive_Type & " = " & lTypeID & " Order By " & gcActive_Pos & " ASC"

' Define recordset
ObjTour.RstSQL lDBHandle, sSQL

' Are there any records
If ObjTour.RstRecordCount(lDBHandle) <> 0 Then

' Loop thru recordset until EOF or gap is found
Do
    If lRt <> ObjTour.DBGetField(gcActive_Pos, lDBHandle) Then Exit Do
    lRt = lRt + 1
Loop While ObjTour.DBMoveNext(lDBHandle)

End If
ObjTour.FreeHandle lDBHandle
lDBHandle = 0

NextAvailablePosition = lRt

On Error GoTo 0
Exit Function

NextAvailablePosition_Error:
    If bDebug Then Handle_Err Err, "NextAvailablePosition-cActivity_Names"
    Resume Next

End Function
