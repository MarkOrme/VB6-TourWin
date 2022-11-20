VERSION 5.00
Begin VB.UserControl Row 
   BackColor       =   &H00808080&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   ScaleHeight     =   255
   ScaleWidth      =   4230
   Begin VB.CommandButton cmdAction 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   325
   End
   Begin Tourwin2002.UTextBox Cell 
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   397
      FieldType       =   2
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataFieldName   =   ""
      BorderStyle     =   0
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   15
      Left            =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Row"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const LINEGAP = 25
Const VLINEGAP = 15

Enum eRowType
    Header
    Standard
    NewRow
End Enum

Enum eMoveDirection
    Up
    Down
    Neutral
End Enum

Event RowGotFocus()
Event DeleteRow(ByVal sKey As String)
Event Changed(ByVal lIndex As Long, ByVal sValue As String)
Event KeyUp(ByVal lColumn As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)


Private m_lColumns          As Long
Private m_bHighlited        As Boolean
Private m_eRowType          As eRowType
Private m_sValue            As String
Private m_sKey              As String
Private oParser             As CParser
Private m_lLeftMostColumn   As Long
Private m_sCellType         As String
Private m_sCellWidth        As String
Private m_NewRowChanged     As Boolean


'---------------------------------------------------------------------------------------
' PROCEDURE : Columns
' DATE      : 8/9/04 21:15
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Columns() As Long

On Local Error GoTo Columns_Error
'Declare local variables

    Columns = m_lColumns

On Error GoTo 0
Exit Property

Columns_Error:
    'If bDebug Then Handle_Err Err, "Columns-Row"
    Resume Next

End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Columns
' DATE      : 8/9/04 21:15
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Columns(ByVal lColumns As Long)

On Local Error GoTo Columns_Error
'Declare local variables
Dim lCell       As Long
Dim lTotalLen   As Long
Dim lLoop       As Long
Dim oColumns    As CColumns
Dim oColumn     As CColumn

 m_lColumns = lColumns
 Refresh

On Error GoTo 0
Exit Property

Columns_Error:
If bDebug Then Handle_Err Err, "Columns-Row"
Resume Next

End Property



Private Sub Cell_Changed(Index As Integer, ByVal sValue As String)

If Cell(Index).Text <> Cell(Index).Tag Then
    RaiseEvent Changed(Index, sValue)
End If

End Sub

Private Sub Cell_GotFocus(Index As Integer)
On Local Error GoTo Focus_Err

RaiseEvent RowGotFocus


On Local Error GoTo 0
Exit Sub
Focus_Err:
If bDebug Then Handle_Err Err, "GotFocus-Row"
Resume Next
End Sub

Private Sub Cell_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)

    RaiseEvent KeyUp(Index, KeyCode, Shift)
    
End Sub

Private Sub Cell_LostFocus(Index As Integer)
If Highlited Then
    HighliteCell Index
End If
End Sub

Private Sub cmdAction_GotFocus()
On Local Error GoTo cmdAction_Err
RaiseEvent RowGotFocus
HighliteRow
Highlited = True


On Local Error GoTo 0
Exit Sub
cmdAction_Err:
If bDebug Then Handle_Err Err, "cmdAction_GotFocus-Row"
Resume Next

End Sub


Private Sub cmdAction_KeyUp(KeyCode As Integer, Shift As Integer)
On Local Error GoTo cmdAction_KeyUp_Err

If vbKeyDelete = KeyCode Then
    RaiseEvent DeleteRow(Me.Key)
End If

On Local Error GoTo 0
Exit Sub

cmdAction_KeyUp_Err:
If bDebug Then Handle_Err Err, "cmdAction_KeyUp-Row"
Resume Next

End Sub



Private Sub UserControl_GotFocus()
    m_NewRowChanged = False
End Sub

Private Sub UserControl_Initialize()
On Local Error GoTo UserControl_Init_Err

Dim lLoop       As Long

Me.RowType = NewRow

On Local Error GoTo 0
Exit Sub

UserControl_Init_Err:
If bDebug Then Handle_Err Err, "UserControl_Init-Row"
Resume Next

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Local Error GoTo ReadProperties_Err

Dim MyControl As Control
Dim AControl As Object

' ------------------------
' Retrieve columns widths
' from header control
' ------------------------
For Each AControl In ParentControls
  If TypeOf AControl Is Header Then
    Set MyControl = AControl
    CellWidth = MyControl.ColumnSizes
    Exit For
  End If
Next


' -----------------------------
' Retrieve field descriptions
' from UserControl control
' ------------------------
For Each AControl In ParentControls
  If TypeOf AControl Is SimpleGrid Then
    Set MyControl = AControl
    CellType = MyControl.FormatStyle
    Exit For
  End If
Next


On Local Error GoTo 0
Exit Sub

ReadProperties_Err:
If bDebug Then Handle_Err Err, "ReadProperties-Row"
Resume Next


End Sub

Private Sub UserControl_Resize()
On Local Error GoTo Resize_Err

Dim lLoop       As Long

' -------------------------------
' Correct for vertical spacing...
' -------------------------------

cmdAction.left = 0
cmdAction.top = 0
cmdAction.Height = UserControl.Height - LINEGAP

For lLoop = Cell.LBound To Cell.UBound
    With Cell(lLoop)
        .top = 0
        .Height = cmdAction.Height - LINEGAP
    End With
Next lLoop

On Local Error GoTo 0

Exit Sub
Resize_Err:
If bDebug Then Handle_Err Err, "UserControl_Resize-Row"
Resume Next


End Sub


'---------------------------------------------------------------------------------------
' PROCEDURE : MoveColumn
' DATE      : 8/9/04 21:34
' Author    : Mark Ormesher
' Purpose   : lDirection is the value of the column index (starts
'             with 0) that appears left most...
'---------------------------------------------------------------------------------------
Public Function MoveColumn(ByVal eDir As eMoveDirection, oRow As Row) As Long

On Local Error GoTo MoveColumn_Error
'Declare local variables
Dim lLoop   As Long
 
If eDir = Up Then
    If oRow.RowType = Standard Then
    
        Me.Value = oRow.Value
        Me.Key = oRow.Key
        
    ElseIf oRow.RowType = NewRow Then
    
        Me.Value = ""
        Me.RowType = NewRow
        Me.Key = vbNullString
        
    End If
End If

MoveColumn = True

On Error GoTo 0
Exit Function

MoveColumn_Error:
If bDebug Then Handle_Err Err, "MoveColumn-Row"
Resume Next


End Function
Public Function MoveRow(ByVal eDir As eMoveDirection, oRow As Row) As Long

On Local Error GoTo MoveRow_Error
'Declare local variables
Dim lLoop   As Long
 
  Me.Value = oRow.Value
  Me.Key = oRow.Key
 
 
If eDir = Down Then Me.RowType = Standard
           
MoveRow = 0

On Error GoTo 0
Exit Function

MoveRow_Error:
    If bDebug Then Handle_Err Err, "MoveRow-Row"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : HighliteRow
' DATE      : 8/10/04 17:56
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function HighliteRow() As Boolean
On Local Error GoTo HighliteRow_Error
'Declare local variables
Dim lCell       As Long


    For lCell = Cell.LBound To Cell.UBound
        Call HighliteCell(lCell)
    Next
    
    
HighliteRow = True

On Error GoTo 0
Exit Function

HighliteRow_Error:
    'If bDebug Then Handle_Err Err, "HighliteRow-Row"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : UnHighliteRow
' DATE      : 8/10/04 17:56
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function UnHighliteRow() As Boolean
On Local Error GoTo UnHighliteRow_Error
'Declare local variables
Dim lCell       As Long

For lCell = Cell.LBound To Cell.UBound
    Cell(lCell).BackColor = &HFFFFFF
    Cell(lCell).ForeColor = &H0&
Next

UnHighliteRow = True

On Error GoTo 0
Exit Function

UnHighliteRow_Error:
    'If bDebug Then Handle_Err Err, "UnHighliteRow-Row"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : Highlited
' DATE      : 8/10/04 17:57
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Highlited() As Boolean

On Local Error GoTo Highlited_Error
'Declare local variables

    Highlited = m_bHighlited

On Error GoTo 0
Exit Property

Highlited_Error:
    'If bDebug Then Handle_Err Err, "Highlited-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Highlited
' DATE      : 8/10/04 17:57
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Highlited(ByVal bHighlited As Boolean)

On Local Error GoTo Highlited_Error
'Declare local variables

    m_bHighlited = bHighlited

On Error GoTo 0
Exit Property

Highlited_Error:
    'If bDebug Then Handle_Err Err, "Highlited-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : RowType
' DATE      : 8/10/04 18:43
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get RowType() As eRowType

On Local Error GoTo RowType_Error
'Declare local variables

    RowType = m_eRowType

On Error GoTo 0
Exit Property

RowType_Error:
    'If bDebug Then Handle_Err Err, "RowType-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : RowType
' DATE      : 8/10/04 18:43
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let RowType(ByVal eRowType As eRowType)

On Local Error GoTo RowType_Error
'Declare local variables

m_eRowType = eRowType

If m_eRowType = NewRow Then
    cmdAction.Caption = "*"
    Me.Value = ""
    
Else
    cmdAction.Caption = ""
End If

On Error GoTo 0
Exit Property

RowType_Error:
    'If bDebug Then Handle_Err Err, "RowType-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Value
' DATE      : 8/11/04 16:49
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Value() As String

On Local Error GoTo Value_Error

    Value = m_sValue


On Error GoTo 0
Exit Property

Value_Error:
'    If bDebug Then Handle_Err Err, "Value-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Value
' DATE      : 8/11/04 16:49
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Value(ByVal sValue As String)

On Local Error GoTo Value_Error
'Declare local variables
Dim lLoop       As Long
Dim oParser     As New CParser


m_sValue = sValue
    
    
oParser.TheString = m_sValue

For lLoop = Cell.LBound To Cell.UBound
    Cell(lLoop).Text = oParser.GetElement(lLoop)
    Cell(lLoop).Tag = Cell(lLoop).Text
Next lLoop
    
On Error GoTo 0
Exit Property

Value_Error:
If bDebug Then Handle_Err Err, "Value-Row"
Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Key
' DATE      : 8/11/04 16:49
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Key() As String

On Local Error GoTo Key_Error
'Declare local variables

    Key = m_sKey

On Error GoTo 0
Exit Property

Key_Error:
'    If bDebug Then Handle_Err Err, "Key-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Key
' DATE      : 8/11/04 16:49
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Key(ByVal sKey As String)

On Local Error GoTo Key_Error
'Declare local variables

    m_sKey = sKey

On Error GoTo 0
Exit Property

Key_Error:
    'If bDebug Then Handle_Err Err, "Key-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : LeftMostColumn
' DATE      : 8/11/04 17:17
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get LeftMostColumn() As Long

On Local Error GoTo LeftMostColumn_Error
'Declare local variables

    LeftMostColumn = m_lLeftMostColumn

On Error GoTo 0
Exit Property

LeftMostColumn_Error:
'    If bDebug Then Handle_Err Err, "LeftMostColumn-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : LeftMostColumn
' DATE      : 8/11/04 17:17
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let LeftMostColumn(ByVal lLeftMostColumn As Long)

On Local Error GoTo LeftMostColumn_Error
'Declare local variables

    m_lLeftMostColumn = lLeftMostColumn


On Error GoTo 0
Exit Property

LeftMostColumn_Error:
    'If bDebug Then Handle_Err Err, "LeftMostColumn-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SelLength
' DATE      : 8/29/04 21:06
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get SelLength() As Long

On Local Error GoTo SelLength_Error
'Declare local variables

    SelLength = ActiveControl.SelLength
    

On Error GoTo 0
Exit Property

SelLength_Error:
    If bDebug Then Handle_Err Err, "SelLength-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SelLength
' DATE      : 8/29/04 21:06
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let SelLength(ByVal lSelLength As Long)

On Local Error GoTo SelLength_Error
'Declare local variables

    ActiveControl.SelLength = lSelLength


On Error GoTo 0
Exit Property

SelLength_Error:
    If bDebug Then Handle_Err Err, "SelLength-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SelStart
' DATE      : 8/29/04 21:06
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get SelStart() As Long

On Local Error GoTo SelStart_Error
'Declare local variables

    SelStart = ActiveControl.SelStart

On Error GoTo 0
Exit Property

SelStart_Error:
    If bDebug Then Handle_Err Err, "SelStart-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SelStart
' DATE      : 8/29/04 21:06
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let SelStart(ByVal lSelStart As Long)

On Local Error GoTo SelStart_Error
'Declare local variables

     ActiveControl.SelStart = lSelStart


On Error GoTo 0
Exit Property

SelStart_Error:
    If bDebug Then Handle_Err Err, "SelStart-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : HighliteCell
' DATE      : 8/31/04 18:41
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Function HighliteCell(ByVal lItem As Long) As Boolean

On Local Error GoTo HighliteCell_Error
'Declare local variables
HighliteCell = True

        Cell(lItem).BackColor = &HFF0000
        Cell(lItem).ForeColor = &H80000005
    

On Error GoTo 0
Exit Function

HighliteCell_Error:
If bDebug Then Handle_Err Err, "HighliteCell-Row"
Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : CellType
' DATE      : 8/31/04 18:45
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Get CellType() As String

On Local Error GoTo CellType_Error
'Declare local variables

    CellType = m_sCellType

On Error GoTo 0
Exit Property

CellType_Error:
    If bDebug Then Handle_Err Err, "CellType-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : CellType
' DATE      : 8/31/04 18:45
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Let CellType(ByVal sCellType As String)

On Local Error GoTo CellType_Error
'Declare local variables

    m_sCellType = sCellType

    Call UserControl.PropertyChanged("CellType")

On Error GoTo 0
Exit Property

CellType_Error:
    If bDebug Then Handle_Err Err, "CellType-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : CellWidth
' DATE      : 8/31/04 18:46
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Get CellWidth() As String

On Local Error GoTo CellWidth_Error
'Declare local variables

    CellWidth = m_sCellWidth

On Error GoTo 0
Exit Property

CellWidth_Error:
    If bDebug Then Handle_Err Err, "CellWidth-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : CellWidth
' DATE      : 8/31/04 18:46
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Let CellWidth(ByVal sCellWidth As String)

On Local Error GoTo CellWidth_Error
'Declare local variables

    m_sCellWidth = sCellWidth

    Call UserControl.PropertyChanged("CellWidth")

On Error GoTo 0
Exit Property

CellWidth_Error:
    If bDebug Then Handle_Err Err, "CellWidth-Row"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Refresh
' DATE      : 8/31/04 18:51
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub Refresh()

On Local Error GoTo Refresh_Error
Const cType = 0
Const cMaxLen = 1

'Declare local variables
Dim lCell           As Long
Dim lWidth          As Long
Dim oParser         As New CParser
Dim oWidthParser    As New CParser
Dim oFormatParser   As New CParser
Dim vFormat         As Variant

' Update look of row control...
Set oParser = New CParser
oParser.TheString = Value
oWidthParser.TheString = CellWidth
oFormatParser.TheString = CellType

If Cell.UBound > m_lColumns Then
    For lCell = Cell.UBound To m_lColumns Step -1
        Unload Cell(lCell)
    Next lCell
    
Else
    For lCell = Cell.UBound To m_lColumns - 1
      If Cell.UBound < lCell Then Load Cell(lCell)
        
        With Cell(lCell)
          If 0 = lCell Then
            .left = cmdAction.Width + VLINEGAP
          Else
            .left = Cell(lCell - 1).left + Cell(lCell - 1).Width + VLINEGAP
          End If
          
        .Text = oParser.GetElement(lCell)

          If IsNumeric(oWidthParser.GetElement(lCell)) Then
            .Width = oWidthParser.GetElement(lCell)
          End If
          
        ' Define Format Attributes
         If "" <> oFormatParser.TheString Then
          vFormat = Split(oFormatParser.GetElement(lCell), vbTab)
          .Max = vFormat(cMaxLen)
          .FieldType = vFormat(cType)
         End If
          .BorderStyle = None
          
            If Highlited Then
             .BackColor = &HFF0000
             .ForeColor = &HFFFFFF
            Else
             .BackColor = &HFFFFFF
             .ForeColor = &H0&
            End If
    
        .Visible = True
        
        End With
    
Next lCell
        
Call ResizeColumns(0, CellWidth)
End If

        
UserControl.Width = Cell(lCell - 1).left + Cell(lCell - 1).Width + LINEGAP



On Error GoTo 0
Exit Sub

Refresh_Error:
    If bDebug Then Handle_Err Err, "Refresh-Row"
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : ResizeColumns
' DATE      : 8/9/04 21:34
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function ResizeColumns(ByVal lStartColumn As Long, sColumnSizes As String) As Long

On Local Error GoTo ResizeColumns_Error

'Declare local variables
Dim lLoop           As Long
Dim oParseWidth     As New CParser
Dim lWidth          As Long

CellWidth = sColumnSizes
oParseWidth.TheString = sColumnSizes


For lLoop = lStartColumn To Cell.UBound

    With Cell(lLoop)
    
        If 0 = lLoop Then
           .left = cmdAction.Width + LINEGAP
        Else
           .left = Cell(lLoop - 1).left + Cell(lLoop - 1).Width + LINEGAP
        End If
        
        ' Check for valid width value
        If IsNumeric(oParseWidth.GetElement(lLoop)) Then
        lWidth = CSng(oParseWidth.GetElement(lLoop))
        If lWidth > 50 Then
        .Width = lWidth
        End If
        End If
    End With
    
Next lLoop
    

UserControl.Width = Cell(lLoop - 1).left + Cell(lLoop - 1).Width + VLINEGAP

On Error GoTo 0
Exit Function

ResizeColumns_Error:
    If bDebug Then Handle_Err Err, "ResizeColumns-Row"
    Resume Next


End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : SetCellFocus
' DATE      : 9/11/04 07:27
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub SetCellFocus(ByVal lColumn As Long)

On Local Error GoTo SetCellFocus_Error

'Declare local variables
If lColumn >= Cell.LBound And lColumn <= Cell.UBound Then
    Cell(lColumn).SetFocus
End If

On Error GoTo 0
Exit Sub

SetCellFocus_Error:
    If bDebug Then Handle_Err Err, "SetCellFocus-Row"
    Resume Next

End Sub
