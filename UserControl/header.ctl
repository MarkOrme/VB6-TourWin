VERSION 5.00
Begin VB.UserControl Header 
   BackColor       =   &H00808080&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   PropertyPages   =   "header.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   4230
   Begin VB.CommandButton cmdColumn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Height          =   225
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   325
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
Attribute VB_Name = "Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Const LINEGAP = 25
Const VLINEGAP = 5
Event SortChange(ByVal sSortDescription As String)
Event ColumnResizeStart()
Event ColumnResizeEnd()
Event ColumnResize(ByVal X As Single, ByVal Y As Single)
Event ResizeColumns(ByVal lStartPos As Long, ByVal sColumnSizes As String)


Private Type tColumnResize
            Source  As Long
            Active  As Boolean
            EndPos  As Single
        End Type
        
Private m_ColumnResize      As tColumnResize
Private m_HeaderCaption     As String
Private m_lColumns          As Long
Private m_bHighlited        As Boolean
Private m_sColumnSizes As String

'---------------------------------------------------------------------------------------
' PROCEDURE : Columns
' DATE      : 8/9/04 21:15
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Columns() As Long
Attribute Columns.VB_ProcData.VB_Invoke_Property = "Grid"

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
Dim lLoop           As Long
Dim oParser         As New CParser
Dim oParserWidths   As New CParser

m_lColumns = lColumns

oParser.TheString = HeaderCaptions  ' String of caption values
oParserWidths.TheString = ColumnSizes

If cmdColumn.UBound > (m_lColumns - 1) Then
    'Unload columns
    For lLoop = cmdColumn.UBound To (m_lColumns - 1) Step -1
        Unload cmdColumn(lLoop)
    Next lLoop
Else
    'Load columns
    For lLoop = 0 To (m_lColumns - 1)
        ' Does control need to be loaded?
        If cmdColumn.UBound < lLoop Then
            Load cmdColumn(lLoop)
        End If
        
        
        With cmdColumn(lLoop)
         If 0 = lLoop Then
            .left = cmdAction.Width + LINEGAP
         Else
            .left = cmdColumn(lLoop - 1).left + cmdColumn(lLoop - 1).Width + LINEGAP
         End If
         .Width = oParserWidths.GetElement(lLoop)
         .Caption = oParser.GetElement(lLoop)
         .Visible = True
        End With
        
    Next lLoop
End If
    
UserControl.Width = cmdColumn(lLoop - 1).left + cmdColumn(lLoop - 1).Width + VLINEGAP + 9

On Error GoTo 0
Exit Property

Columns_Error:
    'If bDebug Then Handle_Err Err, "Columns-Row"
    Resume Next


End Property


Private Sub cmdColumn_Click(Index As Integer)
Static lDirection  As Long

If 0 = lDirection Then
    RaiseEvent SortChange(CStr(Index) & " ASC")
    lDirection = 1
Else
    RaiseEvent SortChange(CStr(Index) & " DESC")
        lDirection = 0
End If
    
    
End Sub

Private Sub cmdColumn_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If TypeOf Source Is PictureBox Then
    m_ColumnResize.EndPos = X
    RaiseEvent ColumnResizeEnd
    
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lCell       As Long

' Determine which column to resize
' To determine source (column),


m_ColumnResize.Source = -1  ' Initiate to no source

For lCell = cmdColumn.LBound To cmdColumn.UBound - 1

    If X > cmdColumn(lCell).left And X < cmdColumn(lCell + 1).left Then
        m_ColumnResize.Source = lCell
        Exit For
    End If
    
Next lCell

' If not found, then very end of usercontrol
If m_ColumnResize.Source = -1 And X > cmdColumn(cmdColumn.UBound).left Then
    m_ColumnResize.Source = cmdColumn.UBound
End If

If m_ColumnResize.Source > -1 Then
    RaiseEvent ColumnResizeStart
    m_ColumnResize.Active = True
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_ColumnResize.Active Then
        RaiseEvent ColumnResize(X, Y)
        MousePointer = 9
    End If

    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo MouseUp_Error

Dim sTemp       As String
Dim lNewWidth   As Long
Dim oParser     As New CParser

If m_ColumnResize.Active Then

    MousePointer = 0
    RaiseEvent ColumnResizeEnd
    
    m_ColumnResize.EndPos = X
    m_ColumnResize.Active = False
    
    ' Now update ColumnSizes property
    ' and then call ResizeColumns method
    lNewWidth = m_ColumnResize.EndPos - cmdColumn(m_ColumnResize.Source).left
    oParser.TheString = ColumnSizes
    
    ' Replace old column size with new one...
    oParser.ChangeElement m_ColumnResize.Source, CStr(lNewWidth)
    ColumnSizes = oParser.TheString
        
    ' Resize columns
    ResizeColumns m_ColumnResize.Source

    
End If
    
On Error GoTo 0
Exit Sub

MouseUp_Error:
If bDebug Then Handle_Err Err, "UserControl_MouseUp-Header"
Resume Next

    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ColumnSizes = PropBag.ReadProperty("ColumnSizes")
End Sub

Private Sub UserControl_Resize()
On Local Error GoTo Resize_Err

Dim lLoop As Long

cmdAction.left = 0
cmdAction.top = 0
cmdAction.Height = UserControl.Height - LINEGAP

For lLoop = cmdColumn.LBound To cmdColumn.UBound
    With cmdColumn(lLoop)
        If 0 = lLoop Then
            .left = cmdAction.Width + LINEGAP
        Else
            .left = cmdColumn(lLoop - 1).left + cmdColumn(lLoop - 1).Width + LINEGAP
        End If
    End With
Next lLoop

On Local Error GoTo 0
Exit Sub
Resize_Err:
Resume Next


End Sub

Private Sub UserControl_Show()
On Local Error GoTo Show_Err
Dim lCell       As Long
Dim lTotalLen   As Long


For lCell = 0 To Columns - 1

    With cmdColumn(lCell)
        .left = cmdColumn(lCell - 1).left + cmdColumn(lCell - 1).Width + LINEGAP
        .Caption = CStr(lCell)
        .Visible = True
    End With
    
Next lCell

UserControl.Width = cmdColumn(lCell - 1).left + cmdColumn(lCell - 1).Width + LINEGAP

On Local Error GoTo 0
Exit Sub
Show_Err:
Resume Next
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : ResizeColumns
' DATE      : 8/9/04 21:34
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function ResizeColumns(ByVal lStartColumn As Long) As Long

On Local Error GoTo ResizeColumns_Error
'Declare local variables
Dim lLoop           As Long
Dim oParseWidth     As New CParser
 
oParseWidth.TheString = ColumnSizes


For lLoop = lStartColumn To cmdColumn.UBound

    With cmdColumn(lLoop)
    
        If 0 = lLoop Then
           .left = cmdAction.Width + LINEGAP
        Else
           .left = cmdColumn(lLoop - 1).left + cmdColumn(lLoop - 1).Width + LINEGAP
        End If
        .Width = oParseWidth.GetElement(lLoop)
        
    End With
    
Next lLoop
    
' Raise Event to SimpleGrid which in turn calls
' row control...
RaiseEvent ResizeColumns(lStartColumn, ColumnSizes)

UserControl.Width = cmdColumn(lLoop - 1).left + cmdColumn(lLoop - 1).Width + VLINEGAP

On Error GoTo 0
Exit Function

ResizeColumns_Error:
    If bDebug Then Handle_Err Err, "ResizeColumns-Row"
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
'Dim lCell       As Long
'    For lCell = Cell.LBound To Cell.UBound
'        Cell(lCell).BackColor = &HFF0000
'        Cell(lCell).ForeColor = &H80000005
'    Next
    
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

'For lCell = Cell.LBound To Cell.UBound
'    Cell(lCell).BackColor = &HFFFFFF
'    Cell(lCell).ForeColor = &H0&
'Next

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

Public Property Get HeaderCaptions() As String
    HeaderCaptions = m_HeaderCaption
End Property

Public Property Let HeaderCaptions(ByVal sNewValue As String)
On Local Error GoTo Let_Caption_Err
Dim lLoop   As Long
Dim oParser As New CParser

    m_HeaderCaption = sNewValue
    
oParser.TheString = m_HeaderCaption
For lLoop = cmdColumn.LBound To cmdColumn.UBound
    cmdColumn(lLoop).Caption = oParser.GetElement(lLoop)
Next lLoop
    
On Local Error GoTo 0
Exit Property
Let_Caption_Err:
Resume Next
End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : ColumnSizes
' DATE      : 9/4/04 19:35
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get ColumnSizes() As String

On Local Error GoTo ColumnSizes_Error
'Declare local variables

    ColumnSizes = m_sColumnSizes

On Error GoTo 0
Exit Property

ColumnSizes_Error:
    If bDebug Then Handle_Err Err, "ColumnSizes-Header"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : ColumnSizes
' DATE      : 9/4/04 19:35
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let ColumnSizes(ByVal sColumnSizes As String)

On Local Error GoTo ColumnSizes_Error
'Declare local variables

    m_sColumnSizes = sColumnSizes

    Call UserControl.PropertyChanged("ColumnSizes")

On Error GoTo 0
Exit Property

ColumnSizes_Error:
    If bDebug Then Handle_Err Err, "ColumnSizes-Header"
    Resume Next


End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ColumnSizes", ColumnSizes)
    
End Sub
