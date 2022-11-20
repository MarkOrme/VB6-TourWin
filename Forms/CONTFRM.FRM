VERSION 5.00
Begin VB.Form ContFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contact Database"
   ClientHeight    =   4860
   ClientLeft      =   2610
   ClientTop       =   2865
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CONTFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4860
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Tag             =   "824"
   Begin Tourwin2002.SimpleGrid SimpleGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      CellHeight      =   "50"
      CellWidth       =   "75"
      FormatStyle     =   ""
      CellWidth       =   "75"
      CellHeight      =   "50"
      Columns         =   6
      HeaderCaption   =   ""
      RowBuffer       =   10
      SortBy          =   ""
   End
   Begin VB.ListBox ConSorLst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton ConExiCmd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Tag             =   "830"
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "      Quick Find Buttons...   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1175
      Left            =   1335
      TabIndex        =   2
      Tag             =   "831"
      Top             =   3600
      Width           =   6495
      Begin VB.CommandButton ConQuiBut 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "ContFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fDeleteOptions              As frmContactDelete
Attribute fDeleteOptions.VB_VarHelpID = -1
Dim oEventTbl                   As CRecord
Dim oContacts                   As CRecord
Dim m_ColumnsWidths             As String
Dim m_FieldType_MaxLen          As String


Private mTotalRows              As Long ' Contains the total rows in the set of records

Private Const MAXCOLS = 6 ' Maximum number of fields in record set.

Sub CheckColumnWidth()

On Local Error GoTo CheckCol_Err
Const SaveSettings = 1


If objMdi.info.UserOptions.GetBool(BitFlags.Pos_0) Then
    LoadContactSettings SaveSettings
End If


On Error GoTo 0
Exit Sub

CheckCol_Err:
If bDebug Then Handle_Err Err, "CheckColumnWidth-ContFrm"
Resume Next

End Sub

Private Sub ConExiCmd_Click()
ConEximnu_Click
End Sub

Private Sub ConEximnu_Click()
CheckColumnWidth
ContFrm.Hide
Define_ContFrm_mnu Unloadmnu
End Sub

Sub ConNewmnu_Click()
On Local Error GoTo MdiNew_Err
'Dim dbCont As Database, setCont As Recordset,
Dim SQL As String, lNewHandle As Long

'lNewHandle = ObjTour.GetHandle
'If 0 = lNewHandle Then Exit Sub
'
''----------------------------------------------
''Open Secondary data Object to find count value
''----------------------------------------------
'SQL = "SELECT Count FROM Contacts ORDER BY Count DESC"
'ObjTour.RstSQL lNewHandle, SQL
'ObjTour.DBMoveLast lNewHandle
''Set setCont = dbCont.OpenRecordset(SQL, dbOpenDynaset)
'
'ProgressBar LoadResString(gcNewContactRec), -1, 1, -1
'
'ConData.Recordset.AddNew
'ConData.Recordset("Contact") = " "
'ConData.Recordset("Id") = objMdi.info.ID
'
'If ObjTour.RstRecordCount(lNewHandle) > 0 Then     ' Check for empty database
'    ObjTour.DBMoveFirst lNewHandle
'    ' Increment the counter for the New contact record...
'    ConData.Recordset("Count") = ObjTour.DBGetField("Count", lNewHandle) + 1
'Else
'    ConData.Recordset("Count") = 1
'End If
''ObjTour.DBClose lNewHandle
'ObjTour.FreeHandle lNewHandle
'ConData.Recordset.Update
'ConData.Refresh
'ConDBGrid.Refresh
'' Give User progress information
'ProgressBar LoadResString(gcNewContactRec), -1, 5, -1
'
'ConDBGrid.Row = 0
'ConDBGrid.Col = 0
'ProgressBar "", 0, 0, 0
Exit Sub
MdiNew_Err:
    If bDebug Then Handle_Err Err, "MdiNewmnu-ContFrm"
    Resume Next
End Sub

Private Sub ConQuiBut_Click(Index As Integer)
Dim iLoop       As Integer     ' Loop variable
Dim bFound      As Boolean
Dim lRt         As Long

On Local Error GoTo ConQuiBut_Err
bFound = False

lRt = FindItemListControl(ConSorLst, Chr$(65 + Index))
If -1 = lRt Then
'    ConDBGrid.Row = 0
Else
'    With ConDBGrid
'        Call .Scroll(0, .VisibleRows)
'        .Row = (lRt - .VisibleRows)
'    End With
'
End If
Exit Sub

For iLoop = 0 To ConSorLst.ListCount - 1
    ConSorLst.ListIndex = iLoop
    If Chr$(65 + Index) <= UCase$(Mid$(ConSorLst.Text, 1, 1)) Then
           If iLoop = 0 Or iLoop > (ConSorLst.ListCount - 1) Then
                    If iLoop = 0 Then
'                        ConDBGrid.Row = 0
                        bFound = True
                        Exit For
                    Else
'                        ConDBGrid.Row = ConSorLst.ListCount - 1
                        bFound = True
                        Exit For
                    End If
            End If
'           ConDBGrid.Row = iLoop '- 1
            bFound = True
           Exit For
    End If
Next iLoop
'If Not bFound Then ConDBGrid.Row = ConSorLst.ListCount = 1
Exit Sub
ConQuiBut_Err:
    If bDebug Then Handle_Err Err, "ConRefmnu-ContFrm"

End Sub

Private Sub ConRefmnu_Click()
On Local Error GoTo ConRef_Err


'    ConDBGrid.Refresh
'    LoadListOfSortedField

Exit Sub
ConRef_Err:

    If bDebug Then Handle_Err Err, "ConRefmnu-ContFrm"
    Resume Next
End Sub

Sub Define_ContFrm_mnu(LoadType As String)
On Local Error GoTo Define_ContFrm_Err

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
    
        MDI!MdiEximnu.Caption = gcContFrm_Exit
        MDI!MdiOptmnu.Caption = gcContFrm_Option
        MDI!MdiNewmnu.Caption = gcContFrm_Newmnu
        MDI!MdiSavmnu.Caption = gcContFrm_Savmnu
        MDI!MdiDelmnu.Caption = gcContFrm_Delmnu
        MDI!MdiOptmnu.Enabled = True
        'MDI!MdiDiamnu.Enabled = False
        MDI!MdiNewmnu.Enabled = True
        MDI!MdiSavmnu.Enabled = True
        MDI!MdiDelmnu.Enabled = True
        MDI!MdiPrimnu.Enabled = True
        
End Select

Exit Sub
Define_ContFrm_Err:
    If bDebug Then Handle_Err Err, "Define_ContFrm_mun-ContFrm"
    Resume Next
End Sub

Private Function DeleteRow_Check(ByVal sKey As String) As Boolean

On Local Error GoTo DeleteRow_Check_Error

Dim MsgStr          As String
Dim lCountVal       As Long
Dim sSQL            As String
Dim lAssociationCnt As Long
Dim lRt             As Long

Const CANCEL_BUTTON = -1
Const ARCHIVE_OPT = 0
Const DISPLAY_OPT = 1

' Move record set to appropriate row

oContacts.RstSQL GetSQLString

oContacts.FindFirst gcContact_Count & " = " & CStr(sKey)
If oContacts.NoMatch Then
    MsgBox "Requested delete could not be performed.", vbOKOnly & vbInformation, LoadResString(gcTour_Win)
    DeleteRow_Check = False
    Exit Function
End If

' ----------------------------------------------------
' Get Count field of current contact and determine if
' contact is associated with any events.
' ----------------------------------------------------
Set oEventTbl = New CRecord
sSQL = "SELECT * FROM " & gcEve_Tour_Event & " WHERE " & gcEve_Tour_PointToCont & " = " & CStr(sKey)

oEventTbl.RstSQL sSQL
oEventTbl.MoveLast
oEventTbl.MoveFirst

lAssociationCnt = oEventTbl.RstRecordCount
    
If 0 = lAssociationCnt Then
    
'    ' ----------------------------
'    ' Prompt user for verify
'    ' before deleting record
'    ' ----------------------------
    MsgStr = LoadResString(gcWantToDelete) + vbLf
    MsgStr = MsgStr & "Contact  = " & " '" & Trim$(oContacts.GetField("Contact")) & "'"

    If vbYes = MsgBox(MsgStr, vbYesNoCancel, LoadResString(gcTourVersion)) Then
        oContacts.Delete gcContTour_Contacts, objMdi.info.ID
        DeleteRow_Check = True
    Else
        DeleteRow_Check = False
    End If
    
    
    Set oEventTbl = Nothing
    Exit Function
Else
' Display Association with events dialog...
    Set fDeleteOptions = New frmContactDelete
    fDeleteOptions.AssociatedCount = lAssociationCnt
    
    
    lRt = fDeleteOptions.Display
    
    Select Case lRt
           Case CANCEL_BUTTON:
                'do nothing
                DeleteRow_Check = False
                
           Case ARCHIVE_OPT:
            ' Change Count to negative, update associated events and
            ' refresh grid to give the illusion of a deletion...

                DeleteRow_Check = True
                sSQL = " UPDATE " & gcEve_Tour_Event & _
                       " SET " & gcEve_Tour_PointToCont & " = " & CStr((-1 * Int(sKey))) & _
                       " WHERE " & gcEve_Tour_PointToCont & " = " & CStr(sKey) & " AND ID = " & objMdi.info.ID
                       
                oEventTbl.Execute sSQL
                
                ' Now delete (hide) contact by changing Count to negative
                sSQL = " UPDATE " & gcContTour_Contacts & _
                       " SET " & gcContact_Count & " = " & CStr((-1 * Int(sKey))) & _
                       " WHERE " & gcContact_Count & " = " & CStr(sKey) & " AND ID = " & objMdi.info.ID
                       
                oEventTbl.Execute sSQL
                
           Case DISPLAY_OPT:
            ' Display list, update event pointToCont field and delete
            ' selected contact...
                MsgBox " Display Option"
                DeleteRow_Check = False
                
    End Select
    
End If

If Not oEventTbl Is Nothing Then
    Set oEventTbl = Nothing
End If

On Local Error GoTo 0
Exit Function

DeleteRow_Check_Error:
If bDebug Then Handle_Err Err, "DeleteRow_Check_Error-ContFrm"
Resume Next


End Function

Private Sub Form_Activate()
Const GetSettings = 0
    'Reload menus
    Define_ContFrm_mnu Loadmnu
    
End Sub

Private Sub Form_Initialize()
Const GetSettings = 0
Dim sSQL        As String

Set oContacts = New CRecord

sSQL = GetSQLString
oContacts.RstSQL sSQL

oContacts.MoveLast
oContacts.MoveFirst

mTotalRows = oContacts.RstRecordCount


LoadContactSettings GetSettings

Call SimpleGrid1.GridSettings(6, _
                             "<0>Contact<0/><1>Last Name<1/><2>First Name<2/><3>Phone<3/><4>Fax<4/><5>Email<5/>", _
                             m_FieldType_MaxLen, _
                             m_ColumnsWidths)



End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then
    Unload Me
    Exit Sub
End If

End Sub

Private Sub Form_Load()

On Local Error GoTo Load_Err

ProgressBar LoadResString(gcLoadingContactDB), -1, 1, -1
' Load Contact Menus
Define_ContFrm_mnu Loadmnu

Me.KeyPreview = True

CentreForm ContFrm, -1
' -----------------------------------
' Load Resource string for this form
' -----------------------------------
LoadFormResourceString ContFrm

' ---------------------------
' Use ObjCon info to find out
' Active index and sort order
'----------------------------
'DESC or ASC
    ProgressBar LoadResString(gcLoadingContactDB), -1, 5, -1

' --------------------------
' Current Frame is InVisible
' and will update later.....
' --------------------------
LoadQuickbuttons

SimpleGrid1.Refresh

ProgressBar LoadResString(gcLoadingContactDB), -1, 10, -1
ProgressBar "", 0, 0, 0

Exit Sub
Load_Err:
    If bDebug Then Handle_Err Err, "Form_Load-ContFrm"
    Resume Next
End Sub

Private Sub Form_LostFocus()
CheckColumnWidth
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    Define_ContFrm_mnu Unloadmnu
    Cancel = -1
    Me.Hide
End If

End Sub

Private Sub Form_Terminate()

On Local Error Resume Next
    Set oContacts = Nothing
    Set oEventTbl = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
CheckColumnWidth
Define_ContFrm_mnu Unloadmnu
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : GetSQLString
' DATE      : 7/25/04 15:42
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Function GetSQLString() As String
On Local Error GoTo GetSQLString_Error
'Declare local variables
Dim sResult As String


GetSQLString = " SELECT " & gcContTour_Contacts_Contact & ", " & _
                 gcContTour_Contacts_Last & ", " & gcContTour_Contacts_First & ", " & _
                 gcContTour_Contacts_Phone & ", " & gcContTour_Contacts_Fax & ", " & _
                 gcContTour_Contacts_Email & ", " & _
                 gcContact_Count & _
               " FROM " & gcContTour_Contacts & _
               " WHERE ID = " & objMdi.info.ID & " AND " & gcContact_Count & " >=0" & _
               " ORDER BY " & ObjCont.info.IndxOrdr & " " & ObjCont.info.SrtOrdr


On Error GoTo 0
Exit Function

GetSQLString_Error:
    If bDebug Then Handle_Err Err, "GetSQLString-ContFrm"
    Resume Next


End Function

Sub LoadContactSettings(Object_Update As Integer)
On Local Error GoTo LContactSettings_Err

Dim SQL         As String
Dim lUserHandle As Long
Dim oParser     As New CParser



 lUserHandle = ObjTour.GetHandle
  If lUserHandle = 0 Then
     Unload Me
     Exit Sub
  End If
  
 ObjTour.RstSQL lUserHandle, " SELECT * FROM " & gcUserTour_ContactOpt & _
                             " WHERE ID = " & objMdi.info.ID


If Object_Update = 1 Then
' ---------------------------
' Check whether record exists
' ---------------------------
If ObjTour.RstRecordCount(lUserHandle) = 0 Then
    ObjTour.AddNew lUserHandle
    ObjTour.DBSetField (gcUserTour_ContactOpt_Id), objMdi.info.ID, lUserHandle
    ObjTour.DBSetField (gcUserTour_ContactOpt_IndexOrder), "Contact", lUserHandle
    ObjTour.DBSetField (gcUserTour_ContactOpt_SortOrder), "DESC", lUserHandle
Else
    ObjTour.Edit lUserHandle
End If

    oParser.TheString = SimpleGrid1.CellWidth
    
    ObjCont.info.ContactWidth = CSng(oParser.GetElement(0))
    ObjCont.info.LastWidth = CSng(oParser.GetElement(1))
    ObjCont.info.FirstWidth = CSng(oParser.GetElement(2))
    ObjCont.info.PhoneWidth = CSng(oParser.GetElement(3))
    ObjCont.info.FaxWidth = CSng(oParser.GetElement(4))
    ObjCont.info.EMailWidth = CSng(oParser.GetElement(5))
   
   
    ObjTour.DBSetField (gcUserTour_ContactOpt_ContactWidth), ObjCont.info.ContactWidth, lUserHandle
    ObjTour.DBSetField (gcUserTour_ContactOpt_LastWidth), ObjCont.info.LastWidth, lUserHandle
    ObjTour.DBSetField (gcUserTour_ContactOpt_FirstWidth), ObjCont.info.FirstWidth, lUserHandle
    ObjTour.DBSetField (gcUserTour_ContactOpt_PhoneWidth), ObjCont.info.PhoneWidth, lUserHandle
    ObjTour.DBSetField (gcUserTour_ContactOpt_FaxWidth), ObjCont.info.FaxWidth, lUserHandle
    ObjTour.DBSetField (gcUserTour_ContactOpt_E_MailWidth), ObjCont.info.EMailWidth, lUserHandle
   
    ObjTour.Update lUserHandle
    ObjTour.FreeHandle lUserHandle
        
Else '******************* initial if else block
' Format ColumnWidth string
ObjCont.info.ContactWidth = ObjTour.DBGetField(gcUserTour_ContactOpt_ContactWidth, lUserHandle)
ObjCont.info.LastWidth = ObjTour.DBGetField(gcUserTour_ContactOpt_LastWidth, lUserHandle)
ObjCont.info.FirstWidth = ObjTour.DBGetField(gcUserTour_ContactOpt_FirstWidth, lUserHandle)
ObjCont.info.PhoneWidth = ObjTour.DBGetField(gcUserTour_ContactOpt_PhoneWidth, lUserHandle)
ObjCont.info.FaxWidth = ObjTour.DBGetField(gcUserTour_ContactOpt_FaxWidth, lUserHandle)
ObjCont.info.EMailWidth = ObjTour.DBGetField(gcUserTour_ContactOpt_E_MailWidth, lUserHandle)

m_ColumnsWidths = "<0>" & ObjCont.info.ContactWidth & "<0/>"
m_ColumnsWidths = m_ColumnsWidths & "<1>" & ObjCont.info.LastWidth & "<1/>"
m_ColumnsWidths = m_ColumnsWidths & "<2>" & ObjCont.info.FirstWidth & "<2/>"
m_ColumnsWidths = m_ColumnsWidths & "<3>" & ObjCont.info.PhoneWidth & "<3/>"
m_ColumnsWidths = m_ColumnsWidths & "<4>" & ObjCont.info.FaxWidth & "<4/>"
m_ColumnsWidths = m_ColumnsWidths & "<5>" & ObjCont.info.EMailWidth & "<5/>"

m_FieldType_MaxLen = "<0>" & dbText & vbTab & "50<0/>"
m_FieldType_MaxLen = m_FieldType_MaxLen & "<1>" & dbText & vbTab & "35<1/>"
m_FieldType_MaxLen = m_FieldType_MaxLen & "<2>" & dbText & vbTab & "25<2/>"
m_FieldType_MaxLen = m_FieldType_MaxLen & "<3>12" & vbTab & "25<3/>" ' Phone Format
m_FieldType_MaxLen = m_FieldType_MaxLen & "<4>12" & vbTab & "25<4/>" ' Phone Format, for fax field
m_FieldType_MaxLen = m_FieldType_MaxLen & "<5>" & dbText & vbTab & "50<5/>" ' Phone Format
    
End If

On Local Error GoTo 0
Exit Sub

LContactSettings_Err:
If bDebug Then Handle_Err Err, "LoadContactSettings-ContFrm"
Resume Next

End Sub

Function LoadListOfSortedField()

Dim sSorted As String
Dim lHandle As Long     'fp

ConSorLst.Clear
ContFrm.Frame1.Enabled = True

sSorted = ObjCont.info.IndxOrdr
    
ObjTour.RstSQL lHandle, GetSQLString
' check if any rows exist.
If ObjTour.RstRecordCount(lHandle) > 0 Then
    Do
        ConSorLst.AddItem ObjTour.DBGetField(sSorted, lHandle)
        Debug.Print ObjTour.DBGetField(sSorted, lHandle)
    Loop While ObjTour.DBMoveNext(lHandle)
End If
ObjTour.FreeHandle lHandle
    
End Function

Sub LoadQuickbuttons()
Dim i As Integer
On Local Error GoTo LoadQuick_Err
For i = 1 To 12
    Load ConQuiBut(i)
        ConQuiBut(i).left = ConQuiBut(i - 1).left + 100 + ConQuiBut(i - 1).Width
        ConQuiBut(i).Caption = Chr$(65 + i)
        ConQuiBut(i).Visible = True
Next i
For i = 13 To 25
    Load ConQuiBut(i)
    If i = 13 Then
        ConQuiBut(i).left = ConQuiBut(0).left
    Else
        ConQuiBut(i).left = ConQuiBut(i - 1).left + 100 + ConQuiBut(i - 1).Width
    End If
        ConQuiBut(i).top = ConQuiBut(0).top + ConQuiBut(0).Height + 50
        ConQuiBut(i).Caption = Chr$(65 + i)
        ConQuiBut(i).Visible = True
Next i
Exit Sub
LoadQuick_Err:
    If bDebug Then Handle_Err Err, "LoadQuickButtons-ContFrm"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : MaxCount
' DATE      : 7/24/04 10:43
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function MaxCount() As Long

On Local Error GoTo MaxCount_Error
'Declare local variables
Dim lResult As Long
Dim SQL     As String


SQL = "SELECT Count FROM Contacts ORDER BY Count DESC"
With oContacts
    .RstSQL SQL
    .MoveLast

    If .RstRecordCount > 0 Then     ' Check for empty database
        .MoveFirst
        ' Increment the counter for the New contact record...
        lResult = .GetField(gcContact_Count) + 1
    Else
        lResult = 1
    End If
End With


MaxCount = lResult

On Error GoTo 0
Exit Function

MaxCount_Error:
If bDebug Then Handle_Err Err, "MaxCount-ContFrm"
Resume Next

End Function

Private Sub SimpleGrid1_DeleteRow(ByVal RowID As Long, ByVal sKey As String)

If DeleteRow_Check(sKey) Then
    '
    SimpleGrid1.DeleteRow RowID, sKey
    
End If

End Sub

Private Sub SimpleGrid1_EditRow(ByVal RowID As Long, ByVal sKey As String, ByVal lColumnNo As Long, ByVal sValue As String)
Dim sSQL    As String

oContacts.FindFirst "Count = " & sKey
If Not oContacts.NoMatch Then
    oContacts.Edit
    Select Case lColumnNo
           Case 0: 'Contact
            Call oContacts.SetField(gcContTour_Contacts_Contact, sValue)
           Case 1: 'Last Name
            Call oContacts.SetField(gcContTour_Contacts_Last, sValue)
           Case 2: 'First Name
            Call oContacts.SetField(gcContTour_Contacts_First, sValue)
           Case 3: 'Phone
            Call oContacts.SetField(gcContTour_Contacts_Phone, sValue)
           Case 4: 'Fax
            Call oContacts.SetField(gcContTour_Contacts_Fax, sValue)
           Case 5: 'Email
            Call oContacts.SetField(gcContTour_Contacts_Email, sValue)
            Case Else
                If bDebug Then Handle_Err "Case not handled", "SimpleGrid1_EditRow-ContFrm"
    End Select
    oContacts.Update
End If

End Sub

'Private Sub SimpleGrid1_EditRow(ByVal RowID As Long, ByVal sKey As String, ByVal sValue As String)
'
'MsgBox "Row No. " & CStr(RowID) & " was changed to " & sValue
'End Sub

Private Sub SimpleGrid1_FetchRows(ByVal lCount As Long, _
                                  ByVal lStartPos As Long, _
                                  ByVal sKey As String, _
                                  ByVal eDirection As eMoveDirection)
                                  
On Local Error GoTo FetchRows_Err
Dim lFetched        As Long
Dim sValue          As String
Dim vData           As Variant
Dim lRows           As Long
Static bAnotherCall As Boolean

' **********************************
' If sKey is null, then search by
' absoluteposition (lStartPos)
' **********************************

oContacts.RstSQL GetSQLString
' Determine where to start fetching rows from...

If vbNullString = sKey Then
    oContacts.MoveLast
    oContacts.MoveFirst
    
    If 0 = oContacts.RstRecordCount Then Exit Sub
    Call oContacts.AbsolutePosition(lStartPos)
    vData = oContacts.GetRows(lCount)
        
Else

    oContacts.FindFirst "Count = " & sKey
    If eDirection = Down Then
       oContacts.MoveNext
       
    ElseIf eDirection = Up Then
       oContacts.MovePrevious
    End If
    vData = oContacts.GetRows(lCount)
End If


For lRows = LBound(vData, 1) To UBound(vData, 2)

sValue = "<0>" & vData(0, lRows) & "<0/>"
sValue = sValue & "<1>" & vData(1, lRows) & "<1/>"
sValue = sValue & "<2>" & vData(2, lRows) & "<2/>"
sValue = sValue & "<3>" & vData(3, lRows) & "<3/>"
sValue = sValue & "<4>" & vData(4, lRows) & "<4/>"
sValue = sValue & "<5>" & vData(5, lRows) & "<5/>"
    
If bAnotherCall Then Exit For
SimpleGrid1.AddRow vData(6, lRows), sValue

Next lRows


SimpleGrid1.RowCount = oContacts.RstRecordCount

On Local Error GoTo 0
Exit Sub

FetchRows_Err:
If bDebug Then Handle_Err Err, "FetchRows-ContFrm"
Resume Next

End Sub

Private Sub SimpleGrid1_LoadSettings()

Call SimpleGrid1.GridSettings(6, "<0>Contact<0/><1>Last Name<1/><2>First Name<2/><3>Phone<3/><4>Fax<4/><5>Email<5/>", CStr(dbText), "1200")

End Sub

Private Sub SimpleGrid1_NewRow(ByVal RowID As Long, ByVal lColumnNo As Long, ByVal sValue As String)
On Local Error GoTo NewRow_Err

Dim lMaxCount       As Long

' Must Call MaxCount before Addnew...
lMaxCount = MaxCount

' Set Required fields
oContacts.RstSQL "SELECT * FROM " & gcContTour_Contacts
oContacts.AddNew

oContacts.SetField gcID, objMdi.info.ID ' set id of record
oContacts.SetField gcContact_Count, lMaxCount
    
    Select Case lColumnNo
           Case 0: 'Contact
            Call oContacts.SetField(gcContTour_Contacts_Contact, sValue)
           Case 1: 'Last Name
            Call oContacts.SetField(gcContTour_Contacts_Last, sValue)
           Case 2: 'First Name
            Call oContacts.SetField(gcContTour_Contacts_First, sValue)
           Case 3: 'Phone
            Call oContacts.SetField(gcContTour_Contacts_Phone, sValue)
           Case 4: 'Fax
            Call oContacts.SetField(gcContTour_Contacts_Fax, sValue)
           Case 5: 'Email
            Call oContacts.SetField(gcContTour_Contacts_Email, sValue)
            Case Else
                If bDebug Then Handle_Err "Case not handled", "SimpleGrid1_EditRow-ContFrm"
    End Select
If oContacts.Update Then
    Call SimpleGrid1.UpdateRowKey(RowID, CStr(lMaxCount))
End If
On Local Error GoTo 0
Exit Sub

NewRow_Err:
If bDebug Then Handle_Err Err, "NewRow-ContFrm"
Resume Next


End Sub

Private Sub SimpleGrid1_SortChange(ByVal sOrder As String)
On Local Error GoTo SortChange_Err

Dim lSpacePos   As Long

lSpacePos = InStr(1, sOrder, " ", vbTextCompare)

' -------------------------
' Update Desc / Asc setting
' -------------------------

ObjCont.info.SrtOrdr = Trim$(Mid$(sOrder, lSpacePos))


' -------------------------
' Update Field to sort by
' -------------------------

Select Case Int(Mid(sOrder, 1, lSpacePos - 1))
       Case 0: 'Sort By Contact Field
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Contact
       Case 1:
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Last
       Case 2:
        ObjCont.info.IndxOrdr = gcContTour_Contacts_First
       Case 3:
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Phone
       Case 4:
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Fax
       Case 5:
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Email
End Select

On Local Error GoTo 0
Exit Sub


SortChange_Err:
If bDebug Then Handle_Err Err, "SortChange-ContFrm"
Resume Next

End Sub
