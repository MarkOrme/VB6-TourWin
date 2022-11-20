VERSION 5.00
Begin VB.Form frmHistorical 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historical Significates"
   ClientHeight    =   2955
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHistorical.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtComment 
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmHistorical.frx":000C
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Historical Significates"
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cboHistorical 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtDistance 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin Tourwin2002.UTextBox txtTime 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         ToolTipText     =   "Time for event (hh:mm:ss)"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Max             =   8
         FieldType       =   22
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
      End
      Begin VB.Label Label1 
         Caption         =   "Comment:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Distance:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Time:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmHistorical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_UserCancelled As Boolean
Dim m_IsDirty As Boolean
Dim m_Name As String
Dim m_Time As String
Dim m_Distance As Single
Dim m_Comment As String


Private Sub cboHistorical_Click()
    Me.IsDirty = True
End Sub

Private Sub cmdCancel_Click()

    Me.UserCancelled = True
    Me.Tag = "Cancel"
    Me.Hide
    
End Sub


Private Sub cmdDelete_Click()
Me.Tag = "Delete"
Me.Hide
End Sub

Private Sub cmdOK_Click()

Me.UserCancelled = False
Me.HistoricalComment = txtComment.Text
Me.HistoricalDistance = txtDistance.Text
Me.HistoricalName = cboHistorical.Text
Me.HistoricalTime = txtTime.Text
Me.Tag = "OK"
Me.Hide

End Sub

Private Sub Form_Activate()
Dim lPosition As Long
txtComment.Text = Me.HistoricalComment
txtDistance.Text = Me.HistoricalDistance

' Default list to 0

lPosition = FindItemListControl(cboHistorical, Me.HistoricalName)
If -1 <> lPosition Then
    cboHistorical.ListIndex = lPosition
Else
    If cboHistorical.ListCount > 0 Then cboHistorical.ListIndex = 0
End If

txtTime.Text = Me.HistoricalTime
Me.UserCancelled = False
Me.IsDirty = False

End Sub

Private Sub Form_Initialize()
Me.HistoricalComment = " "
Me.HistoricalDistance = "0.0"
Me.HistoricalTime = "00:00:00"
Me.HistoricalName = " "
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If vbKeyEscape = KeyAscii Then
    Me.UserCancelled = True
    Me.Hide
End If

End Sub


Public Property Get UserCancelled() As Boolean
UserCancelled = m_UserCancelled
End Property

Public Property Let UserCancelled(ByVal vNewValue As Boolean)
m_UserCancelled = vNewValue
End Property

Public Property Get IsDirty() As Boolean

IsDirty = m_IsDirty

End Property

Public Property Let IsDirty(ByVal vNewValue As Boolean)

m_IsDirty = vNewValue

End Property

Private Sub Form_Load()
Me.KeyPreview = True
End Sub

Private Sub txtComment_Change()
    Me.IsDirty = True
End Sub

Private Sub txtDistance_Change()
    Me.IsDirty = True
End Sub

Private Sub txtDistance_GotFocus()

With txtDistance
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtTime_Change()
    Me.IsDirty = True
End Sub

Public Property Get HistoricalName() As String
    HistoricalName = m_Name
End Property

Public Property Let HistoricalName(ByVal vNewValue As String)
    m_Name = vNewValue
End Property

Public Property Get HistoricalTime() As String
    HistoricalTime = m_Time
End Property

Public Property Let HistoricalTime(ByVal vNewValue As String)
    m_Time = vNewValue
End Property

Public Property Get HistoricalDistance() As Single
    HistoricalDistance = m_Distance
End Property

Public Property Let HistoricalDistance(ByVal vNewValue As Single)
    m_Distance = vNewValue
End Property

Public Property Get HistoricalComment() As String
    HistoricalComment = m_Comment
End Property

Public Property Let HistoricalComment(ByVal vNewValue As String)
    m_Comment = vNewValue
End Property


Private Sub txtTime_ToolTip()
    MDI.StatusBar1.Panels(1).Text = txtTime.ToolTipText
End Sub
