VERSION 5.00
Begin VB.Form frmConconiList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Conconi Test Results"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "ConconiList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDates 
      Height          =   840
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
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
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
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
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmConconiList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_UserCancelled As Boolean
Dim m_date As Date

Public Property Get Conconi_UserCancelled() As Boolean
    Conconi_UserCancelled = m_UserCancelled
End Property

Public Property Let Conconi_UserCancelled(ByVal bNewValue As Boolean)
    m_UserCancelled = bNewValue
End Property

Public Property Get SelectedDate() As Date
    SelectedDate = m_date
End Property

Public Property Let SelectedDate(ByVal dNewValue As Date)
    m_date = dNewValue
End Property

Private Sub cmdAction_Click(Index As Integer)
Select Case Index
        Case 0: ' OK Button
            If -1 <> lstDates.ListIndex Then
                SelectedDate = CDate(lstDates.Text)
                Conconi_UserCancelled = False
                Me.Hide
            Else
                MsgBox "Please select list box item", vbOKOnly, LoadResString(gcTourVersion)
                If lstDates.ListCount > 0 Then
                    lstDates.ListIndex = 0
                    lstDates.SetFocus
                End If
            End If
        Case 1: ' Cancel Button
            Conconi_UserCancelled = True
            Me.Hide
End Select

End Sub

Private Sub Form_Activate()
Dim SQL As String
Dim lTempHandle As Long
Dim lRt         As Long

lstDates.Clear

SQL = " SELECT " & gcDai_Tour_Date & " FROM " & gcDai_Tour_Dai & _
      " WHERE ID = " & objMdi.info.ID & " AND " & _
      "      Type = " & gcDAI_CONCONI & _
      " ORDER BY Date ASC"
      
ObjTour.RstSQL lTempHandle, SQL

If ObjTour.RstRecordCount(lTempHandle) > 0 Then
    Do
        lstDates.AddItem Format$(ObjTour.DBGetField(gcDai_Tour_Date, lTempHandle), "MMMM dd, yyyy")
        ObjTour.DBMoveNext lTempHandle
    Loop While Not ObjTour.EOF(lTempHandle)
    
End If

ObjTour.FreeHandle lTempHandle

If lstDates.ListCount > 0 Then
    ' Be pesitic
    lstDates.ListIndex = 0
    
    lRt = FindItemListControl(lstDates, Format$(SelectedDate, "MMMM dd, yyyy"))
    If -1 <> lRt Then
        lstDates.ListIndex = lRt
    End If

End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Conconi_UserCancelled = True
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
End Sub

Private Sub lstDates_DblClick()
    cmdAction_Click (0) 'OK Click
End Sub
