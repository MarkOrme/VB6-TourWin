VERSION 5.00
Begin VB.UserControl UTextBox 
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   270
   ScaleWidth      =   1215
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "UTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event ToolTip()
Event Changed(ByVal sValue As String)
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)


' Enums
Enum eTextType
        Date_Field = dbDate
        Numeric_Field = dbNumeric
        String_Field = dbText
        Time_Field = dbTime
        Phone_Field = 12
End Enum

Enum eBorderStyle
    None
    Fixed_Single
End Enum

Private m_eFieldType As eTextType
Private m_SavedText     As String
Private m_SavedSelStart As Long
Private m_SavedSelLength As Long
Private m_sDataFieldName As String




Private Sub Text1_Change()
On Local Error GoTo Change_Err

Static IsActive     As Boolean
Static PreviousLen  As Long


'Prevent recursion
If IsActive Then Exit Sub

With Text1
Select Case FieldType
        Case 0: ' Date
          If Not (.Text Like Mid$("##-##-####", 1, Len(.Text))) And _
             Not (.Text Like Mid$("##/##/####", 1, Len(.Text))) Then GoTo Restore
          If Len(.Text) >= PreviousLen Then
          If Len(.Text) = 2 Or Len(.Text) = 5 Then
            IsActive = True
                .Text = .Text & "-"
                .SelStart = Len(.Text)
                .SelLength = 0
                 SaveValues
            IsActive = False
          End If
          End If
        Case 1, dbNumeric: ' Numeric
'            If Len(.Text) >= PreviousLen Then
                If (.Text Like "*[!0-9.]*") Or _
                   (.Text Like "*.*.*") Then GoTo Restore
                   
                  IsActive = True
                      .Text = .Text
                      .SelStart = Len(.Text)
                      .SelLength = 0
                       SaveValues
                  IsActive = False

'            End If
          
        Case 2: ' String
        Case 3, dbTime:   ' Time
          If Not (.Text Like Mid$("##:##:####", 1, Len(.Text))) Then GoTo Restore
          If Len(.Text) >= PreviousLen Then
          If Len(.Text) = 2 Or Len(.Text) = 5 Then
            IsActive = True
                .Text = .Text & ":"
                .SelStart = Len(.Text)
                .SelLength = 0
                 SaveValues
            IsActive = False
          End If
          End If
        Case 4: 'Phone
          If Len(.Text) >= PreviousLen Then
          
          If Not (.Text Like Mid$("+1 (###) ###-####", 1, Len(.Text))) And _
             Not (.Text Like Mid$("### ###-####", 1, Len(.Text))) And _
             Not (.Text Like Mid$("(###) ###-####", 1, Len(.Text))) Then GoTo Restore

                If (.Text Like Mid$("+1 (###) ###-####", 1, Len(.Text))) Then
                  ' Check for ")"
                  If Len(.Text) = 7 Then
                    IsActive = True
                        .Text = .Text & ") "
                        .SelStart = Len(.Text)
                        .SelLength = 0
                         SaveValues
                    IsActive = False
                 End If
                 
                  If Len(.Text) = 12 Then
                  IsActive = True
                      .Text = .Text & "-"
                      .SelStart = Len(.Text)
                      .SelLength = 0
                       SaveValues
                  IsActive = False
                  End If
                 
                End If
              End If
        Case Else:
End Select

PreviousLen = Len(.Text)

End With
On Local Error GoTo 0
Exit Sub

Restore:
    Beep
    IsActive = True
    Text1.Text = m_SavedText
    Text1.SelStart = m_SavedSelStart
    Text1.SelLength = m_SavedSelLength
    IsActive = False

On Local Error GoTo 0
Exit Sub

Change_Err:
Resume Next

End Sub

Private Sub Text1_GotFocus()

RaiseEvent ToolTip

With Text1
    .BackColor = vbYellow
    .ForeColor = vbBlack
    .Tag = .Text
    .SelStart = 0
    .SelLength = Len(.Text)
    If Phone_Field = FieldType Then
        If Len(.Text) = 0 Then
            .Text = "+1 ("
            .SelStart = 5
            .SelLength = 0
        End If
    End If
End With

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    SaveValues
    
    RaiseEvent KeyUp(KeyCode, Shift)
    
End Sub

Private Sub Text1_LostFocus()

    FormatValue Text1.Text
    Text1.BackColor = vbWhite
    Text1.ForeColor = vbBlack
    
    If Text1.Tag <> Text1.Text Then
        RaiseEvent Changed(Text1.Text)
    End If
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SaveValues
End Sub

Private Sub UserControl_InitProperties()

    Set Text1.Font = UserControl.Font

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Local Error Resume Next
    Max = PropBag.ReadProperty("Max")
    FieldType = PropBag.ReadProperty("FieldType")
    Text = PropBag.ReadProperty("Text")
    Text1.BackColor = PropBag.ReadProperty("BackColor")
    Set Font = PropBag.ReadProperty("Font")
    DataFieldName = PropBag.ReadProperty("DataFieldName", "")
    BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    
End Sub

Private Sub UserControl_Resize()
With Text1
    .left = 0
    .top = 0
    .Height = UserControl.ScaleHeight
    .Width = UserControl.ScaleWidth
End With
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : Max
' DATE      : 8/19/04 16:22
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Max() As Long
On Local Error GoTo Max_Error
'Declare local variables

    Max = Text1.MaxLength

On Error GoTo 0
Exit Property

Max_Error:
'    If bDebug Then Handle_Err Err, "Max-UTextBox"
    Resume Next

End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Max
' DATE      : 8/19/04 16:22
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Max(ByVal lMax As Long)


On Local Error GoTo Max_Error
'Declare local variables
    Text1.MaxLength = lMax
    Call UserControl.PropertyChanged("Max")

On Error GoTo 0
Exit Property

Max_Error:
'    If bDebug Then Handle_Err Err, "Max-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : FieldType
' DATE      : 8/19/04 16:24
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get FieldType() As eTextType

On Local Error GoTo FieldType_Error
'Declare local variables

    FieldType = m_eFieldType

On Error GoTo 0
Exit Property

FieldType_Error:
    'If bDebug Then Handle_Err Err, "FieldType-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : FieldType
' DATE      : 8/19/04 16:24
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let FieldType(ByVal eFieldType As eTextType)

On Local Error GoTo FieldType_Error
'Declare local variables

    m_eFieldType = eFieldType
    Call UserControl.PropertyChanged("FieldType")
On Error GoTo 0
Exit Property

FieldType_Error:
    'If bDebug Then Handle_Err Err, "FieldType-UTextBox"
    Resume Next


End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Local Error Resume Next

    Call PropBag.WriteProperty("Max", Max, 0)
    Call PropBag.WriteProperty("FieldType", FieldType, 0) 'Default to string
    Call PropBag.WriteProperty("Text", Text, "")
    Call PropBag.WriteProperty("BackColor", Text1.BackColor)
    Call PropBag.WriteProperty("Font", Font)
    Call PropBag.WriteProperty("DataFieldName", DataFieldName)
    Call PropBag.WriteProperty("BorderStyle", BorderStyle, 1)
    
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : SaveValues
' DATE      : 8/19/04 16:34
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub SaveValues()

On Local Error GoTo SaveValues_Error
'Declare local variables

m_SavedText = Text1.Text
m_SavedSelStart = Text1.SelStart
m_SavedSelLength = Text1.SelLength

On Error GoTo 0
Exit Sub

SaveValues_Error:
'    If bDebug Then Handle_Err Err, "SaveValues-UTextBox"
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : FormatValue
' DATE      : 8/19/04 20:07
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub FormatValue(ByVal sValue As String)
On Local Error GoTo FormatValue_Error
'Declare local variables
    
    Select Case m_eFieldType
            Case 0: ' Date
                Max = 10
                If Len(Text1.Text) = 10 Then
                  If Not IsDate(sValue) Then
                    MsgBox "Invalid date.", vbOKOnly + vbInformation, "TourWin Verions 1.0"
                    Text1.SelStart = 1
                    Text1.SelLength = Len(sValue)
                    Text1.SetFocus
                    
                  End If
                End If
            
            Case 1: ' Numeric
            Case 2: ' String
            Case 3: ' Time
            Max = 8
            If Len(sValue) = 8 Then
              If Not IsDate(sValue) Then
                MsgBox "Invalid Time.", vbOKOnly + vbInformation, "TourWin Verions 1.0"
                Text1.SelStart = 1
                Text1.SelLength = Len(sValue)
                Text1.SetFocus
              End If
            End If

    End Select




On Error GoTo 0
Exit Sub

FormatValue_Error:
'    If bDebug Then Handle_Err Err, "FormatValue-UTextBox"
    
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : Text
' DATE      : 8/22/04 13:24
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "200"
On Local Error GoTo Text_Error
'Declare local variables

    Text = Text1.Text

On Error GoTo 0
Exit Property

Text_Error:
    If bDebug Then Handle_Err Err, "Text-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Text
' DATE      : 8/22/04 13:24
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Text(ByVal sText As String)

On Local Error GoTo Text_Error
'Declare local variables

    Text1.Text = sText

    Call UserControl.PropertyChanged("Text")

On Error GoTo 0
Exit Property

Text_Error:
    If bDebug Then Handle_Err Err, "Text-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : BackColor
' DATE      : 8/22/04 16:14
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR

On Local Error GoTo BackColor_Error
'Declare local variables

    BackColor = Text1.BackColor

On Error GoTo 0
Exit Property

BackColor_Error:
    If bDebug Then Handle_Err Err, "BackColor-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : BackColor
' DATE      : 8/22/04 16:14
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)

On Local Error GoTo BackColor_Error
'Declare local variables

    Text1.BackColor = NewColor
    Call UserControl.PropertyChanged("BackColor")

On Error GoTo 0
Exit Property

BackColor_Error:
    If bDebug Then Handle_Err Err, "BackColor-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Font
' DATE      : 8/22/04 16:21
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Font() As Font

On Local Error GoTo Font_Error
'Declare local variables

    Set Font = UserControl.Font

On Error GoTo 0
Exit Property

Font_Error:
    If bDebug Then Handle_Err Err, "Font-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Font
' DATE      : 8/22/04 16:21
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Font(ByVal NewFont As Font)

On Local Error GoTo Font_Error
'Declare local variables
Dim oControl    As Object

    Set UserControl.Font = NewFont
    
    For Each oControl In Controls
        If TypeOf oControl Is TextBox Then
            Set oControl.Font = UserControl.Font
        End If
    Next
    
    Call UserControl.PropertyChanged("Font")

On Error GoTo 0
Exit Property

Font_Error:
    If bDebug Then Handle_Err Err, "Font-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : DataFieldName
' DATE      : 8/22/04 16:30
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get DataFieldName() As String

On Local Error GoTo DataFieldName_Error
'Declare local variables

    DataFieldName = m_sDataFieldName

On Error GoTo 0
Exit Property

DataFieldName_Error:
    If bDebug Then Handle_Err Err, "DataFieldName-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : DataFieldName
' DATE      : 8/22/04 16:30
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let DataFieldName(ByVal sDataFieldName As String)

On Local Error GoTo DataFieldName_Error
'Declare local variables

    m_sDataFieldName = sDataFieldName

    Call UserControl.PropertyChanged("DataFieldName")

On Error GoTo 0
Exit Property

DataFieldName_Error:
    If bDebug Then Handle_Err Err, "DataFieldName-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SelLength
' DATE      : 8/29/04 21:13
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get SelLength() As Long

On Local Error GoTo SelLength_Error
'Declare local variables

    SelLength = Text1.SelLength

On Error GoTo 0
Exit Property

SelLength_Error:
    If bDebug Then Handle_Err Err, "SelLength-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SelLength
' DATE      : 8/29/04 21:13
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let SelLength(ByVal lSelLength As Long)

On Local Error GoTo SelLength_Error
'Declare local variables

    Text1.SelLength = lSelLength

On Error GoTo 0
Exit Property

SelLength_Error:
    If bDebug Then Handle_Err Err, "SelLength-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SelStart
' DATE      : 8/29/04 21:13
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get SelStart() As Long

On Local Error GoTo SelStart_Error
'Declare local variables

    SelStart = Text1.SelStart

On Error GoTo 0
Exit Property

SelStart_Error:
    If bDebug Then Handle_Err Err, "SelStart-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : SelStart
' DATE      : 8/29/04 21:13
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let SelStart(ByVal lSelStart As Long)

On Local Error GoTo SelStart_Error
'Declare local variables

    Text1.SelStart = lSelStart


On Error GoTo 0
Exit Property

SelStart_Error:
    If bDebug Then Handle_Err Err, "SelStart-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : ForeColor
' DATE      : 8/29/04 21:19
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get ForeColor() As Long

On Local Error GoTo ForeColor_Error
'Declare local variables

    ForeColor = Text1.ForeColor

On Error GoTo 0
Exit Property

ForeColor_Error:
    If bDebug Then Handle_Err Err, "ForeColor-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : ForeColor
' DATE      : 8/29/04 21:19
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let ForeColor(ByVal lForeColor As Long)

On Local Error GoTo ForeColor_Error
'Declare local variables

    Text1.ForeColor = lForeColor

On Error GoTo 0
Exit Property

ForeColor_Error:
    If bDebug Then Handle_Err Err, "ForeColor-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : BorderStyle
' DATE      : 8/29/04 21:38
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get BorderStyle() As eBorderStyle

On Local Error GoTo BorderStyle_Error
'Declare local variables

    BorderStyle = Text1.BorderStyle

On Error GoTo 0
Exit Property

BorderStyle_Error:
    If bDebug Then Handle_Err Err, "BorderStyle-UTextBox"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : BorderStyle
' DATE      : 8/29/04 21:38
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let BorderStyle(ByVal eBorderStyle As eBorderStyle)

On Local Error GoTo BorderStyle_Error
'Declare local variables

    Text1.BorderStyle = eBorderStyle

    Call UserControl.PropertyChanged("BorderStyle")

On Error GoTo 0
Exit Property

BorderStyle_Error:
    If bDebug Then Handle_Err Err, "BorderStyle-UTextBox"
    Resume Next


End Property
