VERSION 5.00
Begin VB.Form PeakOptRpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPeakOptRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "400"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "403"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdPeakRun 
      Caption         =   "Delete "
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Tag             =   "401"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ListBox LstOption 
      Height          =   2985
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Tag             =   "403"
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label lblPeakmulti 
      Caption         =   "Select Peak event(s) to delete, Press CTRL key for multi-select"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Tag             =   "402"
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "PeakOptRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------
' Multi-dimensional array,
' 0 to 100 matchs that of lstOption
' index and tracks Event_Tracker Peak.ID = Event.ID
' --------------------------------------
Dim aListIndex(0 To 100, 1) As Integer

Private Sub cmdCancel_Click()
UserCancel = True
Unload Me
End Sub

Private Sub cmdPeakRun_Click()
Dim iLoop As Integer
Dim bFlag As Boolean
Dim lPeakHandle As Long
Dim sSQL As String      ' Buffer string

' Loop through list to determine if
' any item is selected
For iLoop = 0 To LstOption.ListCount - 1
    If LstOption.Selected(iLoop) Then
        bFlag = True
        Exit For
    End If
Next iLoop

' if nothing selected, prompt user...
If Not bFlag Then
    MsgBox "Please select an item", vbOKOnly, LoadResString(gcTourVersion)
    Exit Sub
End If

bFlag = False
' --------------------------------------------
' This is a delete request from Calendar Form
' Perform Delete of selected Peak Names
' --------------------------------------------
If gcCalendarWindow = objMdi.info.sCurrentActiveWindow Then
'Select distinct Name From Event_Tracker inner join peak on
'( Event_Tracker.id=Peak.id) where peak.page=1 and Event_Tracker.id=1
lPeakHandle = ObjTour.GetHandle
'ObjTour.DBOpen gcEve_Tour, gcEve_Tour_Event_Tracker, lPeakHandle

For iLoop = 0 To LstOption.ListCount - 1
    If True = LstOption.Selected(iLoop) Then
    ' Create Query to delete records form Peak table
    sSQL = "DELETE * FROM " & gcEve_Tour_Peak _
        & " WHERE Page=" & Calndfrm.CalPagMsk.Text _
        & " AND Event_ID=" & Str$(aListIndex(iLoop, 1)) _
        & " AND ID=" & objMdi.info.ID
        Debug.Print sSQL
        
        ObjTour.DBExecute sSQL
  
  
    sSQL = "DELETE * FROM " & gcEve_Tour_Event_Tracker _
        & " WHERE Event_ID=" & Str$(aListIndex(iLoop, 1)) _
        & " AND ID=" & objMdi.info.ID

         ObjTour.DBExecute sSQL
 
    End If
Next iLoop
ObjTour.DBClose lPeakHandle

Else
' -------------------------------------------
' Determine if any list items are selected
' while creating SelectionFormula for Crystal
' -------------------------------------------
sBuffer = "{Peak.Event_ID} IN ["
For iLoop = 0 To LstOption.ListCount - 1
        If True = LstOption.Selected(iLoop) Then
            If False = bFlag Then
                    sBuffer = sBuffer & Str$(aListIndex(iLoop, 1))
                    bFlag = True
            Else
                    sBuffer = sBuffer & " ," & Str$(aListIndex(iLoop, 1))
            End If
            
        End If
Next iLoop
If LstOption.ListCount = 0 Then
    sBuffer = "{Peak.Event_ID} IN [0]"
Else
    sBuffer = sBuffer & "]"
End If
End If
UserCancel = False
Unload Me
End Sub

Private Sub Form_Load()
Dim sRetVal As String
Dim PeakDBHandle As Long    'DB Handle
Dim sSQL As String          ' Buffer for calendar query statement.
CentreForm PeakOptRpt, 0
' -------------------------------
' Assume pesimitic and flag
' that user cancel report request
' -------------------------------
UserCancel = True
' ------------------------------------
' Only Load resource string if
' called from Report menu.
' this is because this form is shared
' ------------------------------------
If "Report" = sBuffer Then
    LoadFormResourceString PeakOptRpt
Else
    PeakOptRpt.Caption = "Peak Delete."
End If

sBuffer = ""

' --------------------------------------------
' Check if form loaded for report or calendar!
' if loaded from calendar then Redefine SQL
' --------------------------------------------
If gcCalendarWindow = objMdi.info.sCurrentActiveWindow Then

    sSQL = "SELECT DISTINCT Name, " & gcEve_Tour_Event_Tracker & ".Event_ID From " _
         & gcEve_Tour_Event_Tracker & " INNER JOIN " & gcEve_Tour_Peak & " ON (" _
         & gcEve_Tour_Event_Tracker & ".ID=" & gcEve_Tour_Peak & ".id) Where " _
         & gcEve_Tour_Event_Tracker & ".id=" & objMdi.info.ID & " and " & gcEve_Tour_Peak & ".Page=" & Calndfrm.CalPagMsk.Text
         
Else

    sSQL = "SELECT * FROM " & gcEve_Tour_Event_Tracker _
        & " WHERE ID = " & objMdi.info.ID _
        & " ORDER BY Name ASC"
    
End If
' -------------------
' Create Recordset...
' -------------------

ObjTour.RstSQL PeakDBHandle, sSQL

Do
    sRetVal = ObjTour.DBGetField("Name", PeakDBHandle)
    
    If Val(sRetVal) <> gcFieldNotExist And sRetVal <> "" Then
        LstOption.AddItem sRetVal
        aListIndex(Val(LstOption.ListCount) - 1, 1) = Val(ObjTour.DBGetField("Event_ID", PeakDBHandle))
        
    End If
    
    ObjTour.DBMoveNext PeakDBHandle
        
Loop While Val(sRetVal) <> gcFieldNotExist And sRetVal <> ""
    
ObjTour.FreeHandle PeakDBHandle


End Sub
