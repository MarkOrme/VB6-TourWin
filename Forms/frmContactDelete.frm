VERSION 5.00
Begin VB.Form frmContactDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contact Deletion"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContactDelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Deletion Options "
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton optDelete 
         Caption         =   "Option1"
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3615
      End
      Begin VB.OptionButton optDelete 
         Caption         =   "Option1"
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Label lblInformation 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmContactDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MESSAGESTRING_1 = "This contact is associated with "
Private Const MESSAGESTRING_2 = " event(s). " & vbCrLf & "When deleting the contact, take the following step:"
Private Const OPTION_1 = "Archive contact information for exiting Event association(s)."
Private Const OPTION_2 = "Display a list of contacts to change all effected Event(s)."
Private Const ARCHIVE_NO = 0
Private Const DISPLAY_NO = 1
Private Const CANCELLED = -1

Private m_lAssociatedCount As Long
Private m_lResult As Long

Private Sub cmdAction_Click(Index As Integer)


If optDelete(ARCHIVE_NO).Value Then
    Result = ARCHIVE_NO
Else
    Result = DISPLAY_NO
End If

Me.Hide

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : AssociatedCount
' DATE      : 7/15/04 20:02
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get AssociatedCount() As Long

On Local Error GoTo AssociatedCount_Error
'Declare local variables

    AssociatedCount = m_lAssociatedCount

On Error GoTo 0
Exit Property

AssociatedCount_Error:
    If bDebug Then Handle_Err Err, "AssociatedCount-frmContactDelete"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : AssociatedCount
' DATE      : 7/15/04 20:02
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let AssociatedCount(ByVal lAssociatedCount As Long)

On Local Error GoTo AssociatedCount_Error
'Declare local variables

    m_lAssociatedCount = lAssociatedCount

On Error GoTo 0
Exit Property

AssociatedCount_Error:
    If bDebug Then Handle_Err Err, "AssociatedCount-frmContactDelete"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Display
' DATE      : 7/15/04 20:04
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function Display()

On Local Error GoTo Display_Error
'Declare local variables


' Pesimistic
Result = CANCELLED

lblInformation.Caption = MESSAGESTRING_1 & AssociatedCount & MESSAGESTRING_2
optDelete(ARCHIVE_NO).Caption = OPTION_1
optDelete(DISPLAY_NO).Caption = OPTION_2

Me.Show vbModal

Display = Result

On Error GoTo 0
Exit Function

Display_Error:
    If bDebug Then Handle_Err Err, "Display-frmContactDelete"
    Resume Next

End Function


'---------------------------------------------------------------------------------------
' PROCEDURE : CloseForm
' DATE      : 7/15/04 20:05
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub CloseForm()

On Local Error GoTo CloseForm_Error

    Unload Me

On Error GoTo 0
Exit Sub

CloseForm_Error:
    If bDebug Then Handle_Err Err, "CloseForm-frmContactDelete"
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : Result
' DATE      : 7/19/04 20:47
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Result() As Long

On Local Error GoTo Result_Error
'Declare local variables

    Result = m_lResult

On Error GoTo 0
Exit Property

Result_Error:
    If bDebug Then Handle_Err Err, "Result-frmContactDelete"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Result
' DATE      : 7/19/04 20:47
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Result(ByVal lResult As Long)

On Local Error GoTo Result_Error
'Declare local variables

    m_lResult = lResult

On Error GoTo 0
Exit Property

Result_Error:
    If bDebug Then Handle_Err Err, "Result-frmContactDelete"
    Resume Next


End Property

