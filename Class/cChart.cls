VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mChartVar As cChartVar
Private Const cPeak = "Peak"
Private Const cEvent = "Event"

Private ChartColour(0 To 9) As Long  ' Stores Colour value
Private ChartName(0 To 9) As String  ' Store Description

Private DailyColour(0 To 9) As Long  ' Store Daily Colour
Private DailyName(0 To 9) As String  '

Private PeakColour(0 To 19) As Long  ' Store Daily Colour
Private PeakName(0 To 19) As String  '

Private EventColour(0 To 9) As Long  ' Store Daily Colour
Private EventName(0 To 9) As String  ' Store Event Name Text

Private IsLoaded As Boolean
Public Event Modified()


Public Property Set info(pChartVar As cChartVar)
Set mChartVar = pChartVar
End Property

Public Property Get info() As cChartVar
Set info = mChartVar
End Property

Public Sub ShowChart(sChartType As String)

' First is Arrays are not loaded
' then load them...
If Not ObjChart.info.IsValuesLoaded Then LoadChartColoursAndNames

' Second Check if Form is loaded
If Not ObjChart.info.IsFormLoaded Then
    'Set chart Type
        ObjChart.info.CurrentChartType = sChartType
        ChartFrm.Show
Else
    ' if sEventType = CurrentChartType then exit
    If sChartType = ObjChart.info.CurrentChartType Then Exit Sub

    ' Call ChartFrm function DisplayChart
    ' to refresh currently displayed
    ObjChart.info.CurrentChartType = sChartType
    ChartFrm.DisplayChartColours (ObjChart.info.CurrentChartType)
    
End If
End Sub
Public Sub CloseChart(Optional sEventType As String)
    
    ObjChart.info.IsFormLoaded = False
    Unload ChartFrm
End Sub

Public Function LowBoundColour() As Long
    LowBoundColour = 0
End Function

Public Function UpBoundColour(Optional ByVal sArrayType As String) As Long
    If cPeak = sArrayType Then
        UpBoundColour = UBound(PeakName)
    Else
        UpBoundColour = UBound(EventName)
    End If
End Function
Public Function SetName(ByVal iNameIndex As Integer, ByVal sCaption As String, sChartType As String) As Boolean
On Local Error Resume Next

    If IsNumeric(iNameIndex) Then
            Select Case sChartType
                    Case gcDailyChart:
                        DailyName(iNameIndex) = sCaption
                    Case gcEventChart:
                        EventName(iNameIndex) = sCaption
                    Case gcPeakChart:
                        PeakName(iNameIndex) = sCaption
            End Select
    End If
End Function
Public Function SetColour(iColourIndex As Integer, lColour As Long, sChartType As String) As Boolean
        If IsNumeric(iColourIndex) And IsNumeric(lColour) Then
                Select Case sChartType
                        Case gcDailyChart:
                            DailyColour(iColourIndex) = lColour
                        Case gcEventChart:
                            EventColour(iColourIndex) = lColour
                        Case gcPeakChart:
                            PeakColour(iColourIndex) = lColour
                End Select
        End If
End Function
Public Function GetColour(iColourIndex As Integer, sChartType As String) As Long

        If IsNumeric(iColourIndex) Then
                Select Case sChartType
                    Case gcDailyChart:
                            GetColour = EventColour(iColourIndex)
                    Case gcEventChart:
                            GetColour = EventColour(iColourIndex)
                    Case gcPeakChart:
                            GetColour = PeakColour(iColourIndex)
                End Select
        End If
        
End Function
Public Function GetName(iColourIndex As Integer, sChartType As String) As String
        If IsNumeric(iColourIndex) Then

                Select Case sChartType
                    Case gcDailyChart:
                            GetName = EventName(iColourIndex)
                    Case gcEventChart:
                            GetName = EventName(iColourIndex)
                    Case gcPeakChart:
                            GetName = PeakName(iColourIndex)
                End Select
        End If
End Function

Public Function SaveChartColoursAndNames(ByVal sChartType As String) As Boolean
On Local Error GoTo SaveChart_Error
Dim FieldVar As String, sFieldName As String
Dim iLoop As Integer
' --------------------------------
' Determines if new records is
' created or edit existing record.
' --------------------------------
Select Case sChartType
    Case gcEventChart:
        ObjTour.RstSQL iSearcherDB, "SELECT * FROM " & gcNameTour_Events & " WHERE ID = " & objMdi.info.ID
        sFieldName = "Event"
    Case gcDailyChart:
        ObjTour.RstSQL iSearcherDB, "SELECT * FROM " & gcNameTour_Events & " WHERE ID = " & objMdi.info.ID
        sFieldName = "Daily"
    
    Case gcPeakChart:
        ObjTour.RstSQL iSearcherDB, "SELECT * FROM " & gcNameTour_PeakNames & " WHERE ID = " & objMdi.info.ID
        sFieldName = "Peak"
End Select

If ObjTour.RstRecordCount(iSearcherDB) = 0 Then
    ObjTour.AddNew (iSearcherDB)
    ObjTour.DBSetField "ID", objMdi.info.ID, iSearcherDB
Else
    ObjTour.Edit iSearcherDB
End If
    For iLoop = ObjChart.LowBoundColour To ObjChart.UpBoundColour(sChartType)
        ' Write description
        FieldVar = sFieldName & Format$(iLoop, "0")
        
        Select Case sChartType
            Case gcEventChart:
                ObjTour.DBSetField FieldVar, EventName(iLoop), iSearcherDB
                ' Write Color
                FieldVar = "Color" & Format$(iLoop, "0")
                ObjTour.DBSetField FieldVar, EventColour(iLoop), iSearcherDB
                
            Case gcDailyChart:
                ObjTour.DBSetField FieldVar, DailyName(iLoop), iSearcherDB
                ' Write Color
                FieldVar = "Color" & Format$(iLoop, "0")
                ObjTour.DBSetField FieldVar, DailyColour(iLoop), iSearcherDB
                
            Case gcPeakChart:
                ObjTour.DBSetField FieldVar, PeakName(iLoop), iSearcherDB
                FieldVar = "Color" & Format$(iLoop, "0")
                ObjTour.DBSetField FieldVar, PeakColour(iLoop), iSearcherDB
                
        End Select

    Next iLoop
    
  ObjTour.Update iSearcherDB
  ' if success, raise event
  ' so other forms can update themselfs
  RaiseEvent Modified
  
Exit Function
SaveChart_Error:
    If bDebug Then Handle_Err Err, "cChart-SaveChartColourAndNames"
    Resume Next
End Function

Public Sub LoadChartColoursAndNames()

' Declare local variables
Dim i As Integer

' Default all the colours to white...
For i = ObjChart.LowBoundColour To ObjChart.UpBoundColour(gcPeakChart)
 PeakColour(i) = CLng(&H80000005)
Next i

For i = ObjChart.LowBoundColour To ObjChart.UpBoundColour(gcEventChart)
    EventColour(i) = CLng(&H80000005)
Next i

' ----------------
' Load Peak Array
' ----------------
' Returns true is records exist
If cActivityNames.StartSearch(gcActive_Type_PeakNames) Then

    For i = ObjChart.LowBoundColour To ObjChart.UpBoundColour(gcPeakChart)
            
        PeakName(i) = cActivityNames.Description
        PeakColour(i) = cActivityNames.Colour
            
        If Not cActivityNames.GetNext Then Exit For ' exit for if no more records
    Next i
End If

' ----------------
' Load Event Array
' ----------------
If cActivityNames.StartSearch(gcActive_Type_EventNames) Then

    For i = ObjChart.LowBoundColour To ObjChart.UpBoundColour(gcEventChart)
            
        EventName(i) = cActivityNames.Description
        EventColour(i) = cActivityNames.Colour
            
        If Not cActivityNames.GetNext Then Exit For ' exit for if no more records
    Next i
End If

ObjChart.info.IsValuesLoaded = True

End Sub

