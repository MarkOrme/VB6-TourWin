VERSION 5.00
Begin VB.UserControl SimpleGrid 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "SimpleGrid.ctx":0000
   ScaleHeight     =   2505
   ScaleWidth      =   6930
   Begin VB.PictureBox imgResize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1000
      Left            =   4440
      ScaleHeight     =   975
      ScaleWidth      =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   15
   End
   Begin Tourwin2002.Header Header1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      ColumnSizes     =   "<0>1000<0/><1>900<1/><2>800<2/><3>700<3/><4>600<4/><5>500<5/><6>400<6/>"
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Left            =   6600
      Max             =   5
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   5
      TabIndex        =   0
      Top             =   2160
      Width           =   5415
   End
   Begin Tourwin2002.Row Row1 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
   End
End
Attribute VB_Name = "SimpleGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"

Option Explicit
'Constants
Const LINETHICKNESS = 50
Const M_CELLHEIGHT = 225
Const M_BUTTONWIDTH = 325
Const CELLWIDTH_ = 1300

' Enums and Type (structures)
Enum FieldFormat
    NoFormat
    SystemFormat
End Enum

Enum ResizeDirection
    Vertical
    Horizontally
End Enum


Private Type GridSettings
    Key_Top         As String
    Key_Bottom      As String
    MaxRows         As String
    RecordCount     As Long
    RowCurrent      As Long
    VisibleRow      As Long
    VScrollValue    As Long
    MoveDirection   As eMoveDirection
End Type


'Events
Public Event FetchRows(ByVal lCount As Long, ByVal lStartPos As Long, ByVal sKey As String, ByVal eDirection As eMoveDirection)
Public Event DeleteRow(ByVal RowID As Long, ByVal sKey As String)
Public Event SortChange(ByVal sOrder As String)
Public Event LoadSettings()
Public Event EditRow(ByVal RowID As Long, ByVal sKey As String, ByVal lColumnNo As Long, ByVal sValue As String)
Public Event NewRow(ByVal RowID As Long, ByVal lColumnNo As Long, ByVal sValue As String)

'Property variables
Private m_lColumns          As Long
Private m_lRows             As Long
Private m_PreviousValue     As Long
Private m_FieldFormats      As String
Private m_sHeaderCaption    As String
Private m_sCellWidth        As String
Private m_sCellHeight       As String
Private m_lHorizontalStart  As Long
Private m_lRowBuffer        As Long
Private m_GridSettings      As GridSettings

'Class
Private m_Rows              As CRows
Private m_Cols              As CColumns
Private m_sSortBy As String



'---------------------------------------------------------------------------------------
' PROCEDURE : Columns
' DATE      : 7/31/04 09:26
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Columns() As Long
Attribute Columns.VB_ProcData.VB_Invoke_Property = "ppGrid"

On Local Error GoTo Columns_Error
'Declare local variables

    Columns = m_lColumns

On Error GoTo 0
Exit Property

Columns_Error:
If bDebug Then Handle_Err Err, "Columns-SimpleGrid"
Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Columns
' DATE      : 7/31/04 09:26
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Columns(ByVal lColumns As Long)

On Local Error GoTo Columns_Error
'Declare local variables
Dim lTemp       As Long

    m_lColumns = lColumns

    Call UserControl.PropertyChanged("Columns")
    
    Header1.Columns = m_lColumns
    For lTemp = Row1.LBound To Row1.UBound
        Row1(lTemp).Columns = m_lColumns
    Next lTemp
    
On Error GoTo 0
Exit Property

Columns_Error:
    'If bDebug Then Handle_Err Err, "Columns-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : FormatStyle
' DATE      : 7/31/04 09:33
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get FormatStyle() As String

On Local Error GoTo Format_Error
'Declare local variables

    FormatStyle = m_FieldFormats

On Error GoTo 0
Exit Property

Format_Error:
If bDebug Then Handle_Err Err, "Format-SimpleGrid"
Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : FormatStyle
' DATE      : 7/31/04 09:33
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let FormatStyle(ByVal sFormat As String)

On Local Error GoTo Format_Error
'Declare local variables

    m_FieldFormats = sFormat
    Call UserControl.PropertyChanged("FormatStyle")

On Error GoTo 0
Exit Property

Format_Error:
If bDebug Then Handle_Err Err, "Format-SimpleGrid"
Resume Next


End Property


'---------------------------------------------------------------------------------------
' PROCEDURE : HeaderCaption
' DATE      : 7/31/04 09:40
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get HeaderCaption() As String
Attribute HeaderCaption.VB_ProcData.VB_Invoke_Property = "ppGrid"

On Local Error GoTo HeaderCaption_Error
'Declare local variables

    HeaderCaption = m_sHeaderCaption

On Error GoTo 0
Exit Property

HeaderCaption_Error:
    'If bDebug Then Handle_Err Err, "HeaderCaption-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : HeaderCaption
' DATE      : 7/31/04 09:40
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let HeaderCaption(ByVal sHeaderCaption As String)

On Local Error GoTo HeaderCaption_Error
'Declare local variables
Dim lLoop       As Long
Dim oParser     As New CParser

    m_sHeaderCaption = sHeaderCaption
    
    Call UserControl.PropertyChanged("HeaderCaption")

    Header1.HeaderCaptions = m_sHeaderCaption

On Error GoTo 0
Exit Property

HeaderCaption_Error:
    'If bDebug Then Handle_Err Err, "HeaderCaption-SimpleGrid"
    Resume Next


End Property


'---------------------------------------------------------------------------------------
' PROCEDURE : CellWidth
' DATE      : 7/31/04 09:48
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get CellWidth() As String

On Local Error GoTo CellWidth_Error
'Declare local variables

    CellWidth = m_sCellWidth

On Error GoTo 0
Exit Property

CellWidth_Error:
    If bDebug Then Handle_Err Err, "CellWidth-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : CellWidth
' DATE      : 7/31/04 09:48
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let CellWidth(ByVal sCellWidth As String)

On Local Error GoTo CellWidth_Error
'Declare local variables

    m_sCellWidth = sCellWidth

    Call UserControl.PropertyChanged("CellWidth")

On Error GoTo 0
Exit Property

CellWidth_Error:
If bDebug Then Handle_Err Err, "CellWidth-SimpleGrid"
Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : CellHeight
' DATE      : 7/31/04 09:48
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Get CellHeight() As String

On Local Error GoTo CellHeight_Error
'Declare local variables

    CellHeight = m_sCellHeight

On Error GoTo 0
Exit Property

CellHeight_Error:
If bDebug Then Handle_Err Err, "CellHeight-SimpleGrid"
Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : CellHeight
' DATE      : 7/31/04 09:48
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Let CellHeight(ByVal sCellHeight As String)

On Local Error GoTo CellHeight_Error
'Declare local variables

    m_sCellHeight = sCellHeight
    Call UserControl.PropertyChanged("CellHeight")

On Error GoTo 0
Exit Property

CellHeight_Error:
    If bDebug Then Handle_Err Err, "CellHeight-SimpleGrid"
    Resume Next


End Property

Private Sub Cell_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Dim lLoop As Long


CellHeight = CellHeight + (Y - CLng(Source.Tag))
Source.Tag = CStr(Y)
Source.ZOrder 0

'Call ResizeGrid(Vertical)

End Sub


Private Sub cmdRowSelector_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

CellHeight = CellHeight + (Y - CLng(Source.Tag))
Source.Tag = CStr(Y)
Source.ZOrder 0

'ResizeGrid Vertical

End Sub


Private Sub Header1_ColumnResize(ByVal X As Single, ByVal Y As Single)
    With imgResize
        .top = 0
        .Height = UserControl.ScaleHeight
        .left = X
    End With
End Sub

Private Sub Header1_ColumnResizeEnd()

    With imgResize
        .top = 0
        .Height = 0
        .left = 0
        .Visible = False
    End With

End Sub

Private Sub Header1_ColumnResizeStart()

    With imgResize
        .top = 0
        .Height = UserControl.ScaleHeight
        .left = 0
        .Width = 5
        .Visible = True
        .ZOrder
    End With

End Sub

Private Sub Header1_ResizeColumns(ByVal lStartPos As Long, ByVal sColumnSizes As String)
Dim lLoop   As Long

CellWidth = sColumnSizes

For lLoop = Row1.LBound To Row1.UBound
    On Local Error Resume Next
    Call Row1(lLoop).ResizeColumns(lStartPos, sColumnSizes)
Next lLoop
    
End Sub

Private Sub Header1_SortChange(ByVal sSortDescription As String)
Dim lRow            As Long
Dim lIsHighlighted  As Long

' Check which row is currently highlighted and sort based on
' its sSortDescription field...

RaiseEvent SortChange(sSortDescription)

lIsHighlighted = -1
For lRow = Row1.LBound To m_GridSettings.VisibleRow - 1
    If Row1(lRow).Highlited Then
        lIsHighlighted = lRow

        Exit For
    End If
Next lRow

For lRow = Row1.LBound To m_GridSettings.VisibleRow - 1
     Row1(lRow).Visible = False
     Row1(lRow).RowType = NewRow
Next lRow
m_GridSettings.VisibleRow = 0

If -1 = lIsHighlighted Then
    RaiseEvent FetchRows(m_GridSettings.MaxRows, 0, m_GridSettings.Key_Top, Neutral)
Else
    RaiseEvent FetchRows(m_GridSettings.MaxRows, 0, Row1(lIsHighlighted).Key, Neutral)
End If


End Sub

Private Sub HScroll1_Change()

'Header1.MoveColumn HScroll1.Value

End Sub


Private Sub Row1_Changed(Index As Integer, ByVal lIndex As Long, ByVal sValue As String)

Dim lLoop       As Long

    If Row1(Index).RowType <> NewRow Then
        RaiseEvent EditRow(Index, Row1(Index).Key, lIndex, sValue)
    Else
        'Before raising NewRow event, update grid so
        'so info is now standard row.
        If m_GridSettings.MaxRows > m_GridSettings.VisibleRow Then
        
            ' add new row below...
            Row1(Index).RowType = Standard
            If Row1.UBound <= Index Then
                Load Row1(Index + 1)
                 Call Row1(Index + 1).ResizeColumns(0, CellWidth)
            End If
            
            With Row1(Index + 1)
                .Columns = Columns
                .ZOrder 1
                .left = 0
                .RowType = NewRow
                .Visible = True
                .top = Row1(Index).top + Row1(Index).Height
                .Refresh
            End With
            m_GridSettings.VisibleRow = m_GridSettings.VisibleRow + 1
                    
            RaiseEvent NewRow(Index, lIndex, sValue)
        Else
            ' Move rows up ...
            ' In this case move info up one row, leave new row but
            ' clear it's data...
            For lLoop = 0 To (m_GridSettings.MaxRows - 2)
                Row1(lLoop).MoveRow Up, Row1(lLoop + 1)
            Next lLoop
            
            RaiseEvent NewRow(Index - 1, lIndex, sValue)
        End If
        

    End If
    
End Sub

Private Sub Row1_DeleteRow(Index As Integer, ByVal sKey As String)
    RaiseEvent DeleteRow(Index, sKey)
End Sub

Private Sub Row1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If TypeOf Source Is PictureBox Then
    MsgBox "Drop @ " & CStr(X)
End If
End Sub

Private Sub Row1_KeyUp(Index As Integer, ByVal lColumn As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
' check for arrow keys
If Index > 0 And vbKeyUp = KeyCode Then
    Row1(Index - 1).SetCellFocus (lColumn)
    Exit Sub
End If

If Index < (m_GridSettings.VisibleRow - 1) And vbKeyDown = KeyCode Then
    Row1(Index + 1).SetCellFocus (lColumn)
    Exit Sub
End If

If lColumn < Columns And vbKeyRight = KeyCode Then
    SendKeys "{TAB}"
    Exit Sub
End If

If lColumn > 0 And vbKeyLeft = KeyCode Then
    SendKeys "+{TAB}"
    Exit Sub
End If



End Sub

Private Sub Row1_RowGotFocus(Index As Integer)

Dim lLoop       As Long

For lLoop = Row1.LBound To Row1.UBound
    If Index <> lLoop Then
        Row1(lLoop).UnHighliteRow
    End If
Next lLoop


End Sub



Private Sub UserControl_DragDrop(Source As Control, X As Single, Y As Single)



CellHeight = CellHeight + (Y - CLng(Source.Tag))
Source.Tag = CStr(Y)
Source.ZOrder 0

End Sub

Private Sub UserControl_Initialize()
Set m_Rows = New CRows
Set m_Cols = New CColumns

m_GridSettings.MaxRows = 5
m_PreviousValue = 0 ' VScroll helper

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Local Error Resume Next

    CellHeight = PropBag.ReadProperty("CellHeight", 50)
    CellWidth = PropBag.ReadProperty("CellWidth", 75)
    Me.FormatStyle = PropBag.ReadProperty("FormatStyle")
    Columns = PropBag.ReadProperty("Columns")
    HeaderCaption = PropBag.ReadProperty("HeaderCaption")
    RowBuffer = PropBag.ReadProperty("RowBuffer", 10)
    HorizontalStart = CellHeight
    SortBy = PropBag.ReadProperty("SortBy")
    
    
End Sub

Private Sub UserControl_Resize()

HScroll1.Width = UserControl.ScaleWidth - (VScroll1.Width)
HScroll1.top = UserControl.Height - (HScroll1.Height + 65)
VScroll1.left = UserControl.Width - (VScroll1.Width + 50)
VScroll1.top = 0
VScroll1.Height = HScroll1.top

End Sub

Private Sub UserControl_Show()
    Row1(0).Columns = Columns
    '' Prototype
    '' FetchRows(ByVal lCount As Long, ByVal sKey As String, ByVal eDirection As eMoveDirection)
    ' Calculate number of rows and can visibly fit into grid.
    
    RowBuffer = (UserControl.Height - Header1.Height - HScroll1.Height) / Row1(0).Height
    m_GridSettings.VScrollValue = 0
    m_GridSettings.MoveDirection = Down
    m_GridSettings.MaxRows = RowBuffer
    Row1(0).RowType = NewRow
    m_GridSettings.VisibleRow = 1
    RaiseEvent FetchRows(RowBuffer, 0, vbNullString, eMoveDirection.Down)
    DoEvents

End Sub

Private Sub UserControl_Terminate()

    Set m_Rows = Nothing
    Set m_Cols = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "CellHeight", CellHeight, 75
    PropBag.WriteProperty "CellWidth", CellWidth, 50
    PropBag.WriteProperty "FormatStyle", FormatStyle, 0
    PropBag.WriteProperty "CellWidth", CellWidth, 1300
    PropBag.WriteProperty "CellHeight", CellHeight, 480
    PropBag.WriteProperty "Columns", Columns, 0
    PropBag.WriteProperty "HeaderCaption", HeaderCaption
    PropBag.WriteProperty "RowBuffer", RowBuffer
    PropBag.WriteProperty "SortBy", SortBy
    
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : DrawGrid
' DATE      : 8/4/04 13:40
' Author    : Mark Ormesher
' Purpose   : This function create a empty grid.
'             the grid has two rows: Header row and new record row.
'             The first control of each column is a button control followed
'             by text boxes...
'---------------------------------------------------------------------------------------
Private Function DrawGrid() As Boolean
On Local Error GoTo DrawGrid_Error
''Declare local variables
'Dim lRow        As Long
'Dim lCol        As Long
'Dim lItems      As Long
'Dim lTop        As Long
'
'' -----------------------------------------
'' Determine the number of textboxes to add.
'' -----------------------------------------
'lItems = (2 * Columns)
'
'
'' ------------------
'' Load Text boxes...
'' ------------------
'For lRow = (Cell.Count) To lItems
'    Load Cell(lRow)
'Next
'
'' -------------------------------
'' Load horizontal image controls
'' -------------------------------
'Load Hsplitter(1)
'
'' -------------------------------
'' Load button controls
'' -------------------------------
'Load cmdRowSelector(1)
'
'' ----------------------------
'' Initialize loop veriables...
'' ----------------------------
'lTop = 0
'
'' ------------
'' Build Grid -
'' ------------
'For lRow = 0 To 1
'     ' -------------------------------
'     ' Put Button at beginning of row
'     ' -------------------------------
'      With cmdRowSelector(lRow)
'        .Left = 0
'        .Width = M_BUTTONWIDTH
'        .Top = lTop
'        .Height = M_CELLHEIGHT
'        .Visible = True
'      End With
'
'   For lCol = 0 To Columns - 1
'
'     lItems = (lRow * Columns) + lCol
'        '-----------------------------------------
'        ' Define Column properties for given cell
'        '-----------------------------------------
'
'        With Cell(lItems)
'          If 0 = lCol Then
'            .Left = cmdRowSelector(lRow).Left + M_BUTTONWIDTH + LINETHICKNESS
'          Else
'            .Left = Cell(lItems - 1).Left + (Cell(lItems - 1).Width + LINETHICKNESS)
'            ' Appropriate for horizontal moves
'          End If
'
'
'         .BackColor = vbWhite
'         .Enabled = True
'         .Width = CELLWIDTH_
'         .Height = CellHeight
'         .Top = lTop
'         .Visible = True
'
'        End With
'
'        '-----------------------------------------
'        ' Define Row properties for given cell
'        '-----------------------------------------
'        Select Case lRow
'               Case 0: '  First Row, add header information
'                    Select Case lCol
'                           Case 0: ' First Caption
'                            Cell(lItems).Text = "Contact"
'                           Case 1: ' First Caption
'                            Cell(lItems).Text = "First Name"
'                           Case 2: ' First Caption
'                            Cell(lItems).Text = "Last Name"
'                           Case 3: ' First Caption
'                            Cell(lItems).Text = "E-mail"
'                    End Select
'                    Cell(lItems).BackColor = &HC0C0C0
'                    Cell(lItems).Enabled = False
'                Case Else:
'
'        End Select
'
'
'    Next lCol
'    ' Position Horizontal splitter just below added row
'
'    With Hsplitter(lRow)
'        .Left = 0
'        .Top = lTop + M_CELLHEIGHT
'        .Width = Cell(lItems).Left + CELLWIDTH_
'        .Visible = True
'        .Tag = .Top
'        .ZOrder 0
'    End With
'
'    ' Update Top position for next row...
'     lTop = M_CELLHEIGHT + LINETHICKNESS
'Next lRow
'
'' Fill Grid
'CellHeight = M_CELLHEIGHT
On Local Error GoTo 0
Exit Function

DrawGrid_Error:
    'If bDebug Then Handle_Err Err, "DrawGrid-SimpleGrid"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : HorizontalStart
' DATE      : 8/6/04 15:49
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Get HorizontalStart() As Long

On Local Error GoTo HorizontalStart_Error
'Declare local variables

    HorizontalStart = m_lHorizontalStart

On Error GoTo 0
Exit Property

HorizontalStart_Error:
'    If bDebug Then Handle_Err Err, "HorizontalStart-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : HorizontalStart
' DATE      : 8/6/04 15:49
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Let HorizontalStart(ByVal lHorizontalStart As Long)

On Local Error GoTo HorizontalStart_Error
'Declare local variables

    m_lHorizontalStart = Abs(lHorizontalStart)

On Error GoTo 0
Exit Property

HorizontalStart_Error:
'    If bDebug Then Handle_Err Err, "HorizontalStart-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : GetRowCollection
' DATE      : 8/8/04 10:22
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get GetRowCollection() As CRows

On Local Error GoTo GetRowCollection_Error
'Declare local variables

    Set GetRowCollection = m_Rows

On Error GoTo 0
Exit Property

GetRowCollection_Error:
    'If bDebug Then Handle_Err Err, "GetRowCollection-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : GetRowCollection
' DATE      : 8/8/04 10:22
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Set GetRowCollection(colGetRowCollection As CRows)

On Local Error GoTo GetRowCollection_Error
'Declare local variables

    Set m_Rows = colGetRowCollection

On Error GoTo 0
Exit Property

GetRowCollection_Error:
    'If bDebug Then Handle_Err Err, "GetRowCollection-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : AddRow
' DATE      : 8/8/04 18:24
' Author    : Mark Ormesher
' Purpose   : AddRow by the client when a FetchRow event is fired
'             by this control. The senerio that call for a FetchRow event are:
'             1. Grid is completely full and the user clicks VScroll control to
'                request a single row change. Scroll up or down is determined by
'                m_GridSettings.MoveDirection value.
'
'             2. User clicks VScroll control to request large row change.
'
'---------------------------------------------------------------------------------------
Public Function AddRow(ByVal sKey As String, ByVal sValue As String) As Boolean

On Local Error GoTo AddRow_Error

'Declare local variables
Dim oParser     As CParser
Dim lItem       As Long ' Cell item
Dim lCount      As Long ' Count of cells
Dim lTop        As Long ' Position of current last rows
Dim lLoop       As Long ' Number of columns
Dim lRowbutton  As Long

' The NEWRow row is always available... so
' change this row to standard, add requested
' information (i.e. sKey + sValue) and then
' make new row at bottom of grid
' ********************************
' * DETERMINE WHICH ROW WILL RECEIVE
' * THE PASSED INFO...
' ********************************
' Senerio #1.
If CInt(m_GridSettings.MaxRows) = m_GridSettings.VisibleRow Then
        ' Move each row up one and then
        ' update the last or first row
        ' with passed info.
    If m_GridSettings.MoveDirection = Down Then
            
        For lLoop = 0 To Row1.UBound - 2
           Row1(lLoop).MoveRow Up, Row1(lLoop + 1)
        Next lLoop
        
        Row1(lLoop).Value = sValue
        Row1(lLoop).Key = sKey
        Row1(lLoop).Visible = True
        m_GridSettings.Key_Top = Row1(0).Key
        m_GridSettings.Key_Bottom = Row1(Row1.UBound - 1).Key
        
    Else
        For lLoop = Row1.UBound To 1 Step -1
            Row1(lLoop).MoveRow Down, Row1(lLoop - 1)
        Next lLoop
        
        Row1(0).Value = sValue
        Row1(0).Key = sKey
        
        m_GridSettings.Key_Top = Row1(0).Key
        m_GridSettings.Key_Bottom = Row1(Row1.UBound - 1).Key
        
    End If
  Exit Function
End If


' Store upper bound control index in variable
lItem = m_GridSettings.VisibleRow - 1
If lItem < 0 Then lItem = 0

    ' Add visible row to grid
  If Row1(lItem).RowType = NewRow Then
        With Row1(lItem)
                .RowType = Standard ' Very important that this property
                                    ' is changed first
                .Tag = CLng(lItem)
                .Value = sValue
                .Key = sKey
                .LeftMostColumn = HScroll1.Value
                .Visible = True
        End With
   End If

    ' Add new row
    lItem = m_GridSettings.VisibleRow
    
    ' Check if invisible or not loaded...
    If lItem >= (Row1.UBound + 1) Then
    
        Load UserControl.Row1(lItem)
        With Row1(lItem)
            .Columns = Columns

            .ZOrder 1
            .Refresh
        End With
    End If
        

    Row1(lItem).left = 0
    
      If 0 = lItem Then
         Row1(lItem).top = Header1.top + Header1.Height
      Else
         Row1(lItem).top = Row1(lItem - 1).top + Row1(lItem - 1).Height
      End If
      
    Row1(lItem).RowType = NewRow
    Row1(lItem).Visible = True
    DoEvents

    m_GridSettings.Key_Top = Row1(0).Key
    m_GridSettings.Key_Bottom = Row1(lItem - 1).Key
    m_GridSettings.VisibleRow = m_GridSettings.VisibleRow + 1
    
On Error GoTo 0
Exit Function

AddRow_Error:
    'If bDebug Then Handle_Err Err, "AddRow-SimpleGrid"
    Resume Next
    MsgBox Err.Description


End Function


'---------------------------------------------------------------------------------------
' PROCEDURE : ResizeGrid
' DATE      : 8/9/04 17:11
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function ResizeGrid(eDirection As ResizeDirection) As Boolean
On Local Error GoTo ResizeGrid_Error
Dim lLoop       As Long
Dim lRow        As Long
Dim lCol        As Long
Dim lItem       As Long
Dim lTop        As Long


'Declare local variables

'If eDirection = Vertical Then
'    lRow = cmdRowSelector.Count - 1
'
'    ' Don't resize the first row...
'    For lLoop = 1 To lRow
'        ' For each row change size of row...
'        lTop = cmdRowSelector(lLoop - 1).Top + cmdRowSelector(lLoop - 1).Height + LINETHICKNESS
'
'        cmdRowSelector(lLoop).Top = lTop
'        cmdRowSelector(lLoop).Height = CellHeight
'
'        ' Now move Horizontal image below newly
'        ' resized textboxs
'        Hsplitter(lRow).Top = lTop + CellHeight
'        Hsplitter(lRow).Tag = Hsplitter(lRow).Top
'        Hsplitter(lRow).ZOrder 0
'
'
'        For lCol = 0 To Columns - 1
'
'            lItem = (lRow * lCol) + Columns
'            Cell(lItem).Top = lTop + LINETHICKNESS
'            Cell(lItem).Height = CellHeight
'            Cell(lItem).Refresh
'        Next lCol
'    Next lLoop
'Else
'End If


On Error GoTo 0
Exit Function

ResizeGrid_Error:
    'If bDebug Then Handle_Err Err, "ResizeGrid-SimpleGrid"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : DeleteRow
' DATE      : 8/14/04 12:23
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function DeleteRow(ByVal RowID As Long, ByVal RecordId As String) As Boolean
On Local Error GoTo DeleteRow_Error
'Declare local variables
Dim lRow        As Long
Dim sKey        As String
Dim lLastRow    As Long


'Starting at RowID move information up...
lRow = RowID

Do While lRow + 1 < (m_GridSettings.VisibleRow - 1)

    Call Row1.Item(lRow).MoveRow(Up, Row1(lRow + 1))
    lRow = lRow + 1
    
Loop
 
 
' Edit the last two rows
'lRow = m_GridSettings.VisibleRow

With Row1(lRow + 1) ' hide previous New Row
   .RowType = Standard
   .Visible = False
End With

With Row1(lRow)  ' new row
   .RowType = NewRow
   .Visible = True
End With
 
 
m_GridSettings.Key_Top = Row1(0).Key
m_GridSettings.Key_Bottom = Row1(lRow - 1).Key
sKey = m_GridSettings.Key_Bottom

m_GridSettings.VisibleRow = m_GridSettings.VisibleRow - 1
 
' Now ask for record to replace deleted one...

If (m_GridSettings.RecordCount - 1) > (m_GridSettings.VisibleRow - 1) Then
    RaiseEvent FetchRows(1, 0, sKey, Down)
End If

DeleteRow = True

On Error GoTo 0
Exit Function

DeleteRow_Error:
    If bDebug Then Handle_Err Err, "DeleteRow-SimpleGrid"
    Resume Next


End Function


'---------------------------------------------------------------------------------------
' PROCEDURE : Refresh
' DATE      : 8/15/04 13:18
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub Refresh()

On Local Error GoTo Refresh_Error
'Declare local variables
Dim lRow        As Long


Header1.Columns = Columns
Header1.HeaderCaptions = HeaderCaption

For lRow = Row1.LBound To Row1.UBound
    Row1(lRow).Refresh
Next

On Error GoTo 0
Exit Sub

Refresh_Error:
    If bDebug Then Handle_Err Err, "Refresh-SimpleGrid"
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : RowBuffer
' DATE      : 8/15/04 14:10
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get RowBuffer() As Long

On Local Error GoTo RowBuffer_Error
'Declare local variables
If m_lRowBuffer <= 0 Then m_lRowBuffer = 1

    RowBuffer = m_lRowBuffer

On Error GoTo 0
Exit Property

RowBuffer_Error:
    If bDebug Then Handle_Err Err, "RowBuffer-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : RowBuffer
' DATE      : 8/15/04 14:10
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let RowBuffer(ByVal lRowBuffer As Long)

On Local Error GoTo RowBuffer_Error
'Declare local variables

    m_lRowBuffer = lRowBuffer

    Call UserControl.PropertyChanged("RowBuffer")

On Error GoTo 0
Exit Property

RowBuffer_Error:
    If bDebug Then Handle_Err Err, "RowBuffer-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : RowCount
' DATE      : 8/15/04 15:28
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get RowCount() As Long

On Local Error GoTo RowCount_Error
'Declare local variables

    RowCount = m_GridSettings.RecordCount

On Error GoTo 0
Exit Property

RowCount_Error:
    If bDebug Then Handle_Err Err, "RowCount-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : RowCount
' DATE      : 8/15/04 15:28
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let RowCount(ByVal lRowCount As Long)

On Local Error GoTo RowCount_Error
'Declare local variables

m_GridSettings.RecordCount = lRowCount
    
VScroll1.Max = m_GridSettings.RecordCount
    
    
If m_GridSettings.RecordCount < m_GridSettings.MaxRows Then
    VScroll1.Visible = False
End If

' If sorting, allow scrolling up or down
' to find records before or after
If m_GridSettings.RecordCount > m_GridSettings.VisibleRow Then
    VScroll1.Visible = True
End If
    
If m_GridSettings.RecordCount > 50 Then
    VScroll1.LargeChange = Int(VScroll1.Max / 25)
Else
    VScroll1.LargeChange = 2
End If

    
    

    
On Error GoTo 0
Exit Property

RowCount_Error:
    If bDebug Then Handle_Err Err, "RowCount-SimpleGrid"
    Resume Next


End Property

Private Sub VScroll1_Change()

If VScroll1.Value > m_GridSettings.VScrollValue Then
    RaiseEvent FetchRows(m_GridSettings.MaxRows, VScroll1.Value, vbNullString, Down)
Else
    RaiseEvent FetchRows(m_GridSettings.MaxRows, VScroll1.Value, vbNullString, Up)
End If

m_GridSettings.VScrollValue = VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()

If VScroll1.Value > m_GridSettings.VScrollValue Then
    RaiseEvent FetchRows(m_GridSettings.MaxRows, VScroll1.Value, vbNullString, Down)
Else
    RaiseEvent FetchRows(m_GridSettings.MaxRows, VScroll1.Value, vbNullString, Up)
End If

m_GridSettings.VScrollValue = VScroll1.Value
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : SortBy
' DATE      : 8/29/04 16:51
' Author    : Mark Ormesher
' Purpose   : SortBy stores the column that the grid is sorted by.
'             the sort string is in SQL format, example:
'             "Field ASC"     or "LastName DESC"
'---------------------------------------------------------------------------------------
Public Property Get SortBy() As String

On Local Error GoTo SortBy_Error
'Declare local variables

    SortBy = m_sSortBy

On Error GoTo 0
Exit Property

SortBy_Error:
    If bDebug Then Handle_Err Err, "SortBy-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SortBy
' DATE      : 8/29/04 16:51
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let SortBy(ByVal sSortBy As String)
On Local Error GoTo SortBy_Error
'Declare local variables

    m_sSortBy = sSortBy

    Call UserControl.PropertyChanged("SortBy")

On Error GoTo 0
Exit Property

SortBy_Error:
    If bDebug Then Handle_Err Err, "SortBy-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : GetColumns
' DATE      : 8/30/04 19:01
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get GetColumns() As CColumns

On Local Error GoTo GetColumns_Error
'Declare local variables

    Set GetColumns = m_Cols

On Error GoTo 0
Exit Property

GetColumns_Error:
    If bDebug Then Handle_Err Err, "GetColumns-SimpleGrid"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : GetColumns
' DATE      : 8/30/04 19:01
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Set GetColumns(colGetColumns As CColumns)

On Local Error GoTo GetColumns_Error
'Declare local variables

    Set m_Cols = colGetColumns

On Error GoTo 0
Exit Property

GetColumns_Error:
    If bDebug Then Handle_Err Err, "GetColumns-SimpleGrid"
    Resume Next


End Property

Public Function GridSettings(ByVal lColumns As Long, _
                             ByVal sHeaderCaptions As String, _
                             ByVal sHeaderDBFields As String, _
                             ByVal sCellWidth As String) As Boolean
                             
On Local Error GoTo Settings_Err

Dim bResult As Boolean

FormatStyle = sHeaderDBFields

Columns = lColumns
Header1.HeaderCaptions = sHeaderCaptions ' This will update both
                                         ' Row & Header.
                                         
                                         
' This will in turn update row widths
Header1.ColumnSizes = sCellWidth
Header1.ResizeColumns 0
        
On Local Error GoTo 0
Exit Function


Settings_Err:
If bDebug Then Handle_Err Err, "GetColumns-SimpleGrid"
Resume Next

End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : UpdateRowKey
' DATE      : 9/7/04 20:30
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function UpdateRowKey(ByVal lRowID As Long, ByVal sNewKey As String) As Boolean
On Local Error GoTo UpdateRowKey_Error
'Declare local variables

If m_GridSettings.Key_Top = Row1(lRowID).Key Then m_GridSettings.Key_Top = sNewKey
If m_GridSettings.Key_Bottom = Row1(lRowID).Key Then m_GridSettings.Key_Bottom = sNewKey

Row1(lRowID).Key = sNewKey


UpdateRowKey = True

On Error GoTo 0
Exit Function

UpdateRowKey_Error:

If bDebug Then Handle_Err Err, "UpdateRowKey-SimpleGrid"
UpdateRowKey = False
Resume Next


End Function
