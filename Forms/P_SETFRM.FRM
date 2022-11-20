VERSION 5.00
Begin VB.Form P_SetFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Peak Setup."
   ClientHeight    =   4680
   ClientLeft      =   1575
   ClientTop       =   1410
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "P_SETFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4680
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSchedule 
      Caption         =   " Schedule "
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9135
      Begin VB.ComboBox PeaNamCbo 
         DataSource      =   "PeaData"
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboPeaLen 
         Height          =   315
         ItemData        =   "P_SETFRM.frx":000C
         Left            =   6480
         List            =   "P_SETFRM.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "&Length:"
         Height          =   255
         Left            =   5760
         TabIndex        =   2
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.ListBox PeaLisCbo 
      BackColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   5880
      TabIndex        =   10
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Frame PeaFra 
      Caption         =   "Description"
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   9135
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Remove >>"
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   8
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Re&move >"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "< &Add"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.PictureBox PeaDayGrd 
         BackColor       =   &H00FFFFFF&
         Height          =   2595
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   4275
         TabIndex        =   5
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "C&ycle:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "&Items:"
         Height          =   255
         Index           =   0
         Left            =   5760
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton PeaCloCmd 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "P_SetFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Changed As Boolean      'used to track changes to fields values
Dim PeakCol(0 To 19) As Long   ' snapshot of colors from ini.
Dim PeakDis(0 To 19) As String ' snapshot of discription.
Dim GrdRow As Integer, GrdCol As Integer
Dim PatchArray(0 To 1000) As Long
Dim cTempName As cPeakName      ' Class structure
Dim bSaveMenuEnabled As Boolean
Dim bDeleteMenuEnable As Boolean
Dim bLoading As Boolean
Dim sLastPeakName As String

Sub FillGrid(ByVal NoUsed$, ByVal Sched$, ByVal DayLen%)
'------------------------------------------------
' If QueryTxt = " " then
' just fill in grid titles and set length
' to Daylen.
' else
' access peaktour and find field name
' -----------------------------------------------

On Local Error GoTo FillError

' Declare Local Variables
Dim i As Integer, CellWidth As Integer, RowNum As Integer
Dim iSched As Integer, cSchedNam As String
Dim iCurrentPos As Integer
Dim sDeliminator As String
Dim sPosition As String
Dim cPeak_Name As cPeakName

' ------------------
' set default values if found correct below
' ------------------
'bLoading = True
'    'Initialize the recordset
'
'    If DayLen <> 0 Then
'        ' define default grid size
'        PeaDayGrd.Rows = DayLen + 1     ' + 1 for Title row
'        PeaDayGrd.Cols = 3
'        ' define colwidths
'        PeaDayGrd.ColWidth(0) = 500
'        PeaDayGrd.ColWidth(1) = 4500
'        PeaDayGrd.ColWidth(2) = 500
'
'        ' move to 0,0
'        PeaDayGrd.Row = 0
'        PeaDayGrd.Col = 0
'        ' name title names
'        PeaDayGrd.Text = "Day"
'        PeaDayGrd.Col = 1
'        PeaDayGrd.Text = "Days Activity Type"
'        PeaDayGrd.Col = 2
'        PeaDayGrd.Text = "Number"
'        PeaDayGrd.Col = 0
'        For i = 1 To DayLen
'            PeaDayGrd.Row = i
'            PeaDayGrd.Col = 0
'            PeaDayGrd.Text = Format$(i, "000")
'            PeaDayGrd.Col = 1
'            PeaDayGrd.Text = ""
'            PeaDayGrd.Col = 2
'            PeaDayGrd.Text = ""
'        Next i
'    End If
'
'If Sched <> " " Then ' fill grid  with peaktour record.
'' ListIndex starts at 0 so DayLen -1
'If DayLen > 0 Then cboPeaLen.ListIndex = DayLen - 1
'    ' ----------------------------
'    ' Loop throw each Sched number
'    ' and assign the correct dis
'    ' to grid.text cell
'    ' -----------------------------    iCurrentPos = 1
'    sDeliminator = ","
'    sPosition = ""
'
'    For i = 1 To DayLen
'        PeaDayGrd.Row = i
'            PeaDayGrd.Col = 2       ' Number columun
'            ' Get Number Column Value
'            If GetPostionValue(Sched, iCurrentPos, sDeliminator, sPosition) Then
'
'                PeaDayGrd.Text = Format$(sPosition, "0")
'            End If
'            If 65 <= Asc(PeaDayGrd.Text) Then
'                PeaDayGrd.Text = Asc(PeaDayGrd.Text) + 10 - 65
'            End If
'
'            PeaDayGrd.Col = PeaDayGrd.Col - 1
'            ' Get Description
'            Set cPeak_Name = cPeakNames.Item(sPosition)
'            If cPeak_Name Is Nothing Then
'                PeaDayGrd.Text = "No Entry. "
'            Else
'                PeaDayGrd.Text = cPeak_Name.Description
'            End If
'    Next i
'
'    PeaDayGrd.Row = 1   ' Set focus to first row
'    PeaDayGrd.Col = 1
'End If ' End Sched if
'bLoading = False
Exit Sub

FillError:
        If bDebug Then Handle_Err Err, "Fillgrid-P_SetFrm"
        Resume Next
End Sub
Private Sub cboPeaLen_Click()
On Local Error Resume Next

' Declare local variables
Dim lCurrentRows As Long
Dim i As Long

If Not bLoading Then

    ' Check if the number of rows is
    ' 1. Equal
    ' 2. Greater than
    ' 3. Less than
    ' ----------------
    ' Check for Equal
    ' ----------------
'    If Val(cboPeaLen.Text) + 1 = PeaDayGrd.Rows Then
'        ' Do nothing
'        Exit Sub
'    End If
'
'    Changed = True
'
'    ' --------------------
'    ' Check for less than
'    ' --------------------
'    If Val(cboPeaLen.Text) + 1 < PeaDayGrd.Rows Then
'       PeaDayGrd.Rows = Val(cboPeaLen.Text) + 1      ' + 1 for Title row
'        Exit Sub
'    End If
'
'    ' ----------------
'    ' Must be more...
'    ' ----------------
'    lCurrentRows = PeaDayGrd.Rows
'
'       PeaDayGrd.Rows = Val(cboPeaLen.Text) + 1
'        For i = lCurrentRows To PeaDayGrd.Rows - 1
'            PeaDayGrd.Row = i
'            PeaDayGrd.Col = 0
'            PeaDayGrd.Text = Format$(i, "000")
'        Next i
End If  ' End bLoading....

End Sub

Private Sub cmdAction_Click(Index As Integer)

On Local Error GoTo Action_Error

Dim iLoop As Integer
Changed = True

'Select Case Index
'    Case 0: 'Add Button
'        If -1 = PeaLisCbo.ListIndex Then
'            MsgBox "Select Activity to add", vbOKOnly + vbInformation
'            Exit Sub
'        End If
'
''        If PeaDayGrd.Row > PeaDayGrd.Rows Or 0 = PeaDayGrd.Row Then
''            MsgBox "Select schedule row to add activity too", vbOKOnly + vbInformation
''            Exit Sub
''        End If
'
'        PeaDayGrd.Col = 2
'        PeaDayGrd.Text = Format$(PatchArray(PeaLisCbo.ListIndex), "0")
'
'        PeaDayGrd.Col = 1
'        PeaDayGrd.Text = PeaLisCbo.Text
'
'        ' Move down 1 row if possible
'        If PeaDayGrd.Row < (PeaDayGrd.Rows - 1) Then
'            PeaDayGrd.Row = PeaDayGrd.Row + 1
'        End If
'    Case 1: ' Remove Button
'
'        PeaDayGrd.Col = 1
'        PeaDayGrd.Text = " "
'
'        PeaDayGrd.Col = 2
'        PeaDayGrd.Text = " "
'
'    Case 2: ' Remove All Button
'        For iLoop = 1 To PeaDayGrd.Rows - 1
'            PeaDayGrd.Row = iLoop
'            PeaDayGrd.Col = 1
'            PeaDayGrd.Text = " "
'
'            PeaDayGrd.Col = 2
'            PeaDayGrd.Text = " "
'        Next iLoop
'    Case Else
'        MsgBox "No case for cmdAction click " & CStr(Index)
'End Select
On Error GoTo 0

Exit Sub
Action_Error:
    Err.Clear
End Sub

Public Sub cmdNew_Click()
On Local Error GoTo New_Err

Dim iGetCBIndexFromString As Long

frmItems.ColourButtonEnabled = False
frmItems.MaxDescLength = 35 ' Max Length of Peak Name
frmItems.Show vbModal

If Not frmItems.USERCANCELLED Then
    
    PeaNamCbo.AddItem frmItems.txtDescription.Text
    ' Set comboBox to the New value
    iGetCBIndexFromString = FindItemListControl(PeaNamCbo, frmItems.txtDescription.Text)
    If -1 <> iGetCBIndexFromString Then
            PeaNamCbo.ItemData(iGetCBIndexFromString) = -1
            PeaNamCbo.ListIndex = iGetCBIndexFromString
    End If
    
    cboPeaLen.ListIndex = 6
    FillGrid " ", " ", cboPeaLen.Text
    If Not bSaveMenuEnabled Then
        bSaveMenuEnabled = True
        MDI.MdiSavmnu.Enabled = bSaveMenuEnabled
    End If
    Changed = True  'Set flag that this should be saved...
    'cboPeaLen.SetFocus
    
End If
Unload frmItems

On Local Error GoTo 0
Exit Sub

New_Err:
If bDebug Then Handle_Err Err, "cmdNew_Click-P_SetFrm"
Resume Next

End Sub

Private Sub Form_Activate()
    Define_P_Set_mnu Loadmnu
End Sub

Private Sub Form_Deactivate()
    Define_P_Set_mnu Unloadmnu
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
' --------------
' setup
' --------------
On Local Error GoTo Form_Err
Const MAXCYCLELENGTH = 180
Dim SQL As String
Dim iTemp As Integer

bLoading = True
CentreForm P_SetFrm, -1
SendMessage PeaLisCbo.hWnd, LB_SETHORIZONTALEXTENT, 550, 0
Me.KeyPreview = True

' ========================================
' Load PeakNames collection if necessary
' ========================================
If Not cPeakNames.IsUpdated Then
    cPeakNames.Load_Collection
End If

' ========================================
' Load Peak List Combobox with Peak Names
' ========================================
PeaLisCbo.Clear
For Each cTempName In cPeakNames
    PeaLisCbo.AddItem cTempName.Description
    
    ' PatchArray is a link between collection and List Box
    PatchArray(PeaLisCbo.ListCount - 1) = cTempName.Position
Next
If PeaLisCbo.ListCount > 0 Then PeaLisCbo.ListIndex = 0

' Load Cycle Length value

For iTemp = 1 To MAXCYCLELENGTH
    cboPeaLen.AddItem CStr(iTemp)
Next iTemp

' -----------------------------
' Open database and set dynaset
' -----------------------------

SQL = "SELECT * FROM " & gcPeakTour_Peaks & " WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

If ObjTour.RstRecordCount(iSearcherDB) <> 0 Then
    ObjTour.DBMoveFirst iSearcherDB
    
     Do Until ObjTour.EOF(iSearcherDB)
        If "" <> ObjTour.DBGetField("P_Nam", iSearcherDB) Then
            PeaNamCbo.AddItem ObjTour.DBGetField("P_Nam", iSearcherDB)
        End If
        
       ObjTour.DBMoveNext iSearcherDB
           
    Loop
    
       ObjTour.DBMoveFirst iSearcherDB

       PeaNamCbo.ListIndex = 0
       FillGrid " ", ObjTour.DBGetField("P_Sched", iSearcherDB), Val(ObjTour.DBGetField(PEAKS_LENGTH, iSearcherDB))

Else
        If cboPeaLen.ListCount > 0 Then
            cboPeaLen.ListIndex = 0
            FillGrid " ", " ", cboPeaLen.Text
        End If

        bDeleteMenuEnable = False
        MDI.MdiDelmnu.Enabled = bDeleteMenuEnable
        
End If

If PeaNamCbo.ListCount > 0 Then PeaNamCbo.Text = PeaNamCbo.List(0)
    ProgressBar " ", 0, 0, 0
    Changed = False
    bLoading = False
Exit Sub
' Helps keep track of current PeakName
sLastPeakName = PeaNamCbo.Text
Form_Err:
    If bDebug Then Handle_Err Err, "Form_Load-P_SetFrm"
    Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
If bDebug Then Handle_Err 0, "FormUnload-HrtZone"
    If Changed Then
        If vbYes = MsgBox("Save changes to " & sLastPeakName & "?", vbYesNo, LoadResString(gcTourVersion)) Then
            PeaSavCmd_Click
        End If
End If

    Changed = False
    Define_P_Set_mnu Unloadmnu
End Sub

Private Sub P_SAbomnu_Click()
AboutMsg
End Sub

Private Sub P_SDelmnu_Click()
PeaDelCmd_Click
End Sub

Private Sub P_SEximnu_Click()
    P_SetFrm.Hide
End Sub

Private Sub P_SSavmnu_Click()
PeaSavCmd_Click
End Sub

Private Sub PeaCloCmd_Click()
Unload P_SetFrm
End Sub


Private Sub PeaDayGrd_Click()
'
' Declare local variable
'
Dim GrdLeft As Integer
Dim GrdTop As Integer
Dim GrdCount As Integer

On Local Error GoTo PeaDay_Err
'If PeaDayGrd.Col = 2 And PeaDayGrd.Row > 0 Then
'    'Calculate location so Cbo
'    'covers grid cell .
'    GrdLeft = PeaDayGrd.Left + PeaDayGrd.ColWidth(0) + PeaDayGrd.ColWidth(1) + 15
'    GrdCount = PeaDayGrd.RowPos(PeaDayGrd.Row) / 240
'    Select Case GrdCount
'            Case 1 To 6:
'            ' stops errors for grids  with < 7 rows
'                If GrdCount >= Val(PeaDayGrd.Rows) - 1 Then
'                    GrdTop = PeaDayGrd.Top + PeaDayGrd.RowPos(PeaDayGrd.Row - 5)
'                Else
'                    GrdTop = PeaDayGrd.Top + PeaDayGrd.RowPos(PeaDayGrd.Row + 1)
'                End If
'            Case Else
'                  GrdTop = PeaDayGrd.Top + PeaDayGrd.RowPos(PeaDayGrd.Row - 5)
'    End Select
'    PeaLisCbo.Left = GrdLeft
'    PeaLisCbo.Top = GrdTop
'        PeaLisCbo.Visible = -1
'        PeaLisCbo.SetFocus
'End If
Exit Sub
PeaDay_Err:
        If bDebug Then Handle_Err Err, "PeaDayGrd_Click-P_SetFrm"
        Resume Next
End Sub

Private Sub PeaDayGrd_LostFocus()
'GrdRow = PeaDayGrd.Row
'GrdCol = PeaDayGrd.Col
End Sub




Public Sub PeaDelCmd_Click()
On Local Error GoTo PeaDel_Err
Dim Deleted_Descr As String, MsgStr As String, SearchStr As String
' ---------------------------------------------
' 1.) Prompt user before deleting
'     then Remove Item form ListBox...
' 2.) Find with Criteria search and then delete
' 3.) set List Box to first possibe and Load
' ---------------------------------------------
Deleted_Descr = Trim$(PeaNamCbo.Text)
MsgStr = "Are you sure you wish to delete" & vbLf
MsgStr = MsgStr & "'" & Deleted_Descr & "'"

If vbNo = MsgBox(MsgStr, vbYesNo + vbQuestion, LoadResString(gcTourVersion)) Then Exit Sub
' ---
' 1.)
' ---
PeaNamCbo.RemoveItem PeaNamCbo.ListIndex
'
' 2.)
' ---
SearchStr = "P_Nam = '" & Deleted_Descr & "'"
ObjTour.DBFindFirst SearchStr, iSearcherDB
'setPea.FindFirst SearchStr
If Not ObjTour.NoMatch(iSearcherDB) Then ObjTour.Delete iSearcherDB
' ---
' 3.)
' ---
       PeaNamCbo.ListIndex = 0
       If ObjTour.RstRecordCount(iSearcherDB) <> 0 Then
            ObjTour.DBMoveFirst iSearcherDB
            'setPea.MoveFirst
            FillGrid " ", ObjTour.DBGetField("P_Sched", iSearcherDB), ObjTour.DBGetField(PEAKS_LENGTH, iSearcherDB)
       Else
            PeaNamCbo.Text = "None"
            cboPeaLen.ListIndex = 0
            FillGrid " ", " ", Val(cboPeaLen.Text)
            bDeleteMenuEnable = False
            
       End If
Exit Sub
PeaDel_Err:
    If bDebug Then Handle_Err Err, "PeaDelCmd-P_SetFrm"
    Resume Next
End Sub

Private Sub PeaLisCbo_DblClick()
On Local Error GoTo PeaLisDbl_Err
'' ---------------
'' Set focus back
'' to orginal cell
'' ---------------
'PeaDayGrd.Row = GrdRow
'PeaDayGrd.Col = GrdCol
'PeaDayGrd.Text = PeaLisCbo.Text
'' -------------------
'' Move to Number cell
'' -------------------
'PeaDayGrd.Col = PeaDayGrd.Col - 1
'PeaDayGrd.Text = Format$(PatchArray(PeaLisCbo.ListIndex), "0")
'PeaLisCbo.Visible = 0
'PeaSavCmd.SetFocus
Exit Sub
PeaLisDbl_Err:
    If bDebug Then Handle_Err Err, "PeaLisCbo-P_SetFrm"
    Resume Next
End Sub

Private Sub PeaLisCbo_GotFocus()
'PeaLisCbo.ListIndex = 0
End Sub
Private Sub PeaNamCbo_Change()
    Changed = True
End Sub

Private Sub PeaNamCbo_Click()
    Update_Window
    sLastPeakName = PeaNamCbo.Text
End Sub

Private Sub PeaNamCbo_LostFocus()

If sLastPeakName = PeaNamCbo.Text Then Exit Sub

    If PeaNamCbo.ItemData(PeaNamCbo.ListIndex) = -1 Then
        PeaNamCbo.ItemData(PeaNamCbo.ListIndex) = 0
    Else
        Update_Window
    End If
End Sub

Public Sub PeaSavCmd_Click()
On Local Error GoTo PeaSav_Err
' ---------------
' Create New Peak
' ---------------
 Dim setSched As String
 Dim i As Integer, n As Integer
 Dim TypeCol As String
 Dim setCol As String, setDis As String
 Dim Criteria As String, SQL As String

'Criteria = "P_Nam = '" & Trim$(PeaNamCbo) & "'"
Criteria = "P_Nam = '" & Trim$(sLastPeakName) & "'"

' ======================
' Define RecordSet then
' Get Fields
' =====================
SQL = "SELECT * FROM " & gcPeakTour_Peaks & " WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

ObjTour.DBFindFirst Criteria, iSearcherDB
'setPea.FindFirst Criteria
If ObjTour.NoMatch(iSearcherDB) Then     ' Check if record is found.

ObjTour.AddNew iSearcherDB
    ObjTour.DBSetField "Id", objMdi.info.ID, iSearcherDB
    ObjTour.DBSetField "P_Date", Format$(Now, "mm-dd-yyyy"), iSearcherDB
    ObjTour.DBSetField "P_Nam", Mid$(PeaNamCbo.Text, 1, 35), iSearcherDB
    ObjTour.DBSetField PEAKS_LENGTH, Val(cboPeaLen.Text), iSearcherDB
Else
    ObjTour.Edit iSearcherDB
End If
' --------------------------
' Loop through grid for Number col and
' make string of each number
' --------------------------
'PeaDayGrd.Col = 2
'setSched = ""
'For n = 1 To Val(cboPeaLen.Text)
'    PeaDayGrd.Row = n
'    If setSched = "" Then
'        setSched = PeaDayGrd.Text
'    Else
'        setSched = setSched & "," & PeaDayGrd.Text
'
'    End If
'Next n
setSched = setSched & ","
ObjTour.DBSetField "P_Sched", Trim$(setSched), iSearcherDB
ObjTour.DBSetField PEAKS_LENGTH, Val(cboPeaLen.Text), iSearcherDB
ObjTour.Update iSearcherDB
If Not bDeleteMenuEnable Then
    bDeleteMenuEnable = True
    MDI.MdiDelmnu.Enabled = bDeleteMenuEnable
End If
Changed = False

Exit Sub
PeaSav_Err:
    If bDebug Then Handle_Err Err, "PeaSavCmd-P_SetFrm"
    Resume Next
End Sub


Public Function Update_Window()
On Local Error GoTo PeaNam_Err
' ------------------------------
' Find setPea Name which is
' the same as name in PeaNamCbo.
' If !Found then New Entry
' ------------------------------
Dim Criteria As String, SQL As String
Dim iGetCBIndexFromString As Long

If Changed Then
    If vbYes = MsgBox("Save changes to " & sLastPeakName & "?", vbYesNo, LoadResString(gcTourVersion)) Then
       PeaSavCmd_Click
    Else
        'Determine if it is necessary to remove item from list!
        ' this would only happen after Add Schedule without saving...
        ' Get Index value
        iGetCBIndexFromString = FindItemListControl(PeaNamCbo, sLastPeakName)
        
        ' If found, check itemdata value...
        If -1 <> iGetCBIndexFromString Then
        
            ' If new item without save, itemdata will have a -1...
            If -1 = PeaNamCbo.ItemData(iGetCBIndexFromString) Then
                PeaNamCbo.RemoveItem iGetCBIndexFromString
            End If
            
        End If
        
    End If
    
    Changed = False
End If

Criteria = "P_Nam = '" & Trim$(PeaNamCbo.Text) & "'"

' ---------------------------------
' Define RecordSet then Get Fields
' ---------------------------------
SQL = "SELECT * FROM " & gcPeakTour_Peaks & " WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

ObjTour.DBFindFirst Criteria, iSearcherDB

If ObjTour.NoMatch(iSearcherDB) Then    ' Check if record is found.
    FillGrid " ", " ", Val(cboPeaLen.Text)
Else
    FillGrid " ", ObjTour.DBGetField("P_Sched", iSearcherDB), Val(ObjTour.DBGetField(PEAKS_LENGTH, iSearcherDB))
End If
Changed = False
Exit Function
PeaNam_Err:
    If bDebug Then Handle_Err Err, "Update_Window-P_SetFrm"
    Resume Next

End Function

Sub Define_P_Set_mnu(LoadType As String)
On Local Error GoTo Define_P_Set_mnu
' ------------------------------
' Turn Key Preview on
' to trap esc key press among
' other key presses
' ------------------------------
Me.KeyPreview = True
Select Case LoadType
    Case Unloadmnu:
        MDI!MdiEximnu.Caption = MdiFrm_Exit
        MDI!MdiOptmnu.Caption = MdiFrm_Option
        MDI!MdiNewmnu.Caption = MdiFrm_Newmnu
        MDI!MdiSavmnu.Caption = MdiFrm_Savmnu
        MDI!MdiDelmnu.Caption = MdiFrm_Delmnu
        MDI!MdiOptmnu.Enabled = True
        MDI!MdiDiamnu.Enabled = True
        MDI!MdiNewmnu.Enabled = True
        MDI!MdiSavmnu.Enabled = False
        MDI!MdiDelmnu.Enabled = False
        MDI!MdiPrimnu.Enabled = False
    Case Loadmnu:
    
        MDI!MdiEximnu.Caption = P_SetUp_Exit
        MDI!MdiOptmnu.Caption = P_SetUp_Option
        MDI!MdiNewmnu.Caption = P_SetUp_Newmnu
        MDI!MdiSavmnu.Caption = P_SetUp_Savmnu
        MDI!MdiDelmnu.Caption = P_SetUp_Delmnu
        MDI!MdiOptmnu.Enabled = False
        MDI!MdiDiamnu.Enabled = False
        MDI!MdiNewmnu.Enabled = True
        MDI!MdiSavmnu.Enabled = bSaveMenuEnabled
        MDI!MdiDelmnu.Enabled = True
        MDI!MdiPrimnu.Enabled = True
        
End Select

Exit Sub
Define_P_Set_mnu:
    If bDebug Then Handle_Err Err, "Define_P_Set_mnu-P_SetFrm"
    Resume Next
End Sub

