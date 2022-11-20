VERSION 5.00
Begin VB.Form HrtZone 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Heart Zone SetUp."
   ClientHeight    =   2865
   ClientLeft      =   1575
   ClientTop       =   2700
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HRTZONE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Heart Rate Zones"
      Height          =   2175
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox HeaLisCbo 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   10
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   11
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   12
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   13
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   14
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   15
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   16
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   17
         Left            =   2760
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   18
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   19
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   20
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   21
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   22
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   23
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox HeaBouTxt 
         Height          =   285
         Index           =   24
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Level &1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   0
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Level &2 :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Level &3 :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Level &4 :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Level &5 :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.CommandButton HeaSavCmd 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton HeaCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "HrtZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IndexFocus As Integer           'TextBox which had focus last
Dim cLevels(1 To 5) As String * 3   'Stores Level setting...
Dim bHeartChanged As Boolean
Dim bIsListBoxLoaded As Boolean
Dim lHasFocus As Long

Private Sub Form_Activate()
    Define_Form_menu Me.Name, Loadmnu
End Sub

Private Sub Form_Deactivate()
    Define_Form_menu Me.Name, Unloadmnu
End Sub


Private Sub Form_GotFocus()
CentreForm HrtZone, -1
HrtZone.WindowState = 0     'Normal
Define_Form_menu Me.Name, Loadmnu
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
' -------------------------------------------------------------------
' 1.)Loop Through HeartNames Table and assign Descr to HeaBouTxt(1..9)
' 2.)Then go back and add to HeaLisCbo
' 3.)Loop throught Levels Table and assign Levels To HeaBouTxt(10..24)
' -------------------------------------------------------------------
On Local Error GoTo FormLoad_Err
Dim i As Integer, TypeNum As String, m As Integer, Q As Integer
'Dim dbHeart As Database, setHeart As Recordset,
Dim SQL As String
Dim LevelNum As String


CentreForm HrtZone, 0

Define_Form_menu Me.Name, Loadmnu
Me.KeyPreview = True

' Initialize Connect to Activity Table
cActivityNames.Type_ID = gcActive_Type_HeartNames
' -------------------------------------------------
' Based on Record count show appropriate controls
' and load the first 10 descriptions.
' -------------------------------------------------
If cActivityNames.StartSearch(gcActive_Type_HeartNames) Then
    Do
    
        HeaLisCbo.AddItem cActivityNames.Description
        HeaLisCbo.ItemData(HeaLisCbo.ListCount - 1) = cActivityNames.Position
        
    Loop While cActivityNames.GetNext
End If

' --------------------------------
' 3.) Load Levels and assign HeaBouTxt(10..24) Descriptions.
' --------------------------------
SQL = "SELECT * FROM Levels WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

If ObjTour.RstRecordCount(iSearcherDB) <> 0 Then
    ObjTour.DBMoveFirst iSearcherDB

    For i = 1 To 5
        LevelNum = "Level" & Format$(i, "0")
        cLevels(i) = Format$(ObjTour.DBGetField(LevelNum, iSearcherDB), "000")
        For m = 1 To 3
            Select Case i
                Case 1:
                    Q = Val(Mid$(ObjTour.DBGetField("Level1", iSearcherDB), m, 1))
                    If cActivityNames.FindItemByPosition(Q) Then
                    HeaBouTxt(10 + (m - 1)) = cActivityNames.Description
                    HeaBouTxt(10 + (m - 1)).Tag = Format$(Q, "0")
                    End If
                Case 2:
                    Q = Val(Mid$(ObjTour.DBGetField("Level2", iSearcherDB), m, 1))
                    If cActivityNames.FindItemByPosition(Q) Then
                    HeaBouTxt(13 + (m - 1)) = cActivityNames.Description
                    HeaBouTxt(13 + (m - 1)).Tag = Format$(Q, "0")
                    End If
                Case 3:
                    Q = Val(Mid$(ObjTour.DBGetField("Level3", iSearcherDB), m, 1))
                    If cActivityNames.FindItemByPosition(Q) Then
                    HeaBouTxt(16 + (m - 1)) = cActivityNames.Description
                    HeaBouTxt(16 + (m - 1)).Tag = Format$(Q, "0")
                    End If
                Case 4:
                    Q = Val(Mid$(ObjTour.DBGetField("Level4", iSearcherDB), m, 1))
                    If cActivityNames.FindItemByPosition(Q) Then
                    HeaBouTxt(19 + (m - 1)) = cActivityNames.Description
                    HeaBouTxt(19 + (m - 1)).Tag = Format$(Q, "0")
                    End If
                Case 5:
                    Q = Val(Mid$(ObjTour.DBGetField("Level5", iSearcherDB), m, 1))
                    If cActivityNames.FindItemByPosition(Q) Then
                    HeaBouTxt(22 + (m - 1)) = cActivityNames.Description
                    HeaBouTxt(22 + (m - 1)).Tag = Format$(Q, "0")
                    End If
            End Select
        Next m
    Next i
Else
    For i = 1 To 5
        cLevels(i) = "000"
    Next i
End If
ProgressBar "", 0, 0, 0
bHeartChanged = False
bIsListBoxLoaded = False
Exit Sub
FormLoad_Err:

    If bDebug Then Handle_Err Err, "FormLoad-HrtZone"
    Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

' -----------------
' Check for changes
' -----------------
If bDebug Then Handle_Err 0, "FormUnload-HrtZone"
    If bHeartChanged Then
        If vbYes = MsgBox("Save Changes?", vbYesNo, LoadResString(gcTourVersion)) Then HeaSavCmd_Click
End If

Define_Form_menu Me.Name, Unloadmnu

End Sub


Private Sub HeaBouTxt_Change(Index As Integer)
    bHeartChanged = True
End Sub

Private Sub HeaBouTxt_GotFocus(Index As Integer)
On Local Error GoTo HeaBouGot_Err
' Declare local variables
Dim lRt As Long

'check which textbox was clicked
If Index < 10 Then Exit Sub
    
IndexFocus = Index

' Set this list box to display the Heaboutxt Text value.
lRt = FindItemListControl(HeaLisCbo, HeaBouTxt(Index).Text)
If -1 <> lRt Then
    HeaLisCbo.ListIndex = lRt
End If

HeaLisCbo.Top = HeaBouTxt(Index).Top
HeaLisCbo.Left = HeaBouTxt(Index).Left
HeaLisCbo.TabIndex = HeaBouTxt(Index).TabIndex
HeaLisCbo.Visible = True
HeaLisCbo.SetFocus
'HeaBouTxt(Index).Visible = False

Exit Sub
HeaBouGot_Err:
    If bDebug Then Handle_Err Err, "HeaBouClick-HrtZone"
    Resume Next

End Sub

Sub HeaCanCmd_Click()
    Unload HrtZone
End Sub


Private Sub HeaLisCbo_Click()

HeaBouTxt(IndexFocus).Text = HeaLisCbo.Text
HeaLisCbo.Visible = False

End Sub

Private Sub HeaLisCbo_LostFocus()
On Local Error GoTo HeaLis_Err
Dim i As Integer


If HeaLisCbo.Text = "" Then
    HeaBouTxt(IndexFocus).Visible = True
Exit Sub
End If

HeaBouTxt(IndexFocus).Text = HeaLisCbo.Text
HeaLisCbo.Visible = False

    HeaBouTxt(IndexFocus).Tag = Format$(Str$(HeaLisCbo.ItemData(HeaLisCbo.ListIndex)), "0")
    Select Case HeaBouTxt(IndexFocus).Index
        Case 10 To 12:
            cLevels(1) = HeaBouTxt(10).Tag & HeaBouTxt(11).Tag & HeaBouTxt(12).Tag
        Case 13 To 15:
            cLevels(2) = HeaBouTxt(13).Tag & HeaBouTxt(14).Tag & HeaBouTxt(15).Tag
        Case 16 To 18:
            cLevels(3) = HeaBouTxt(16).Tag & HeaBouTxt(17).Tag & HeaBouTxt(18).Tag
        Case 19 To 21:
            cLevels(4) = HeaBouTxt(19).Tag & HeaBouTxt(20).Tag & HeaBouTxt(21).Tag
        Case 22 To 24:
            cLevels(5) = HeaBouTxt(22).Tag & HeaBouTxt(23).Tag & HeaBouTxt(24).Tag
    End Select
    
'    If IndexFocus < HeaBouTxt.UBound Then
'        HeaBouTxt(IndexFocus + 1).SetFocus
'    End If

Exit Sub
HeaLis_Err:
    If bDebug Then Handle_Err Err, "HeaLisCbo-HrtZone"
    Resume Next
End Sub


Sub HeaSavCmd_Click()
' -------------------------
' Write results to Heart Table
' -------------------------
On Local Error GoTo HeaSav_Err

Dim SQL As String
Dim i As Integer, m As Integer
Dim HeartNum As String, LevelNum As String

'SQL = "SELECT * FROM HeartNames WHERE Id = " & objMdi.info.ID
'ObjTour.RstSQL iSearcherDB, SQL
'' -------------------------------------
'' Write Type description to Names Table
'' -------------------------------------
'If ObjTour.RstRecordCount(iSearcherDB) = 0 Then
'   ObjTour.AddNew iSearcherDB
'   ObjTour.DBSetField "Id", objMdi.info.ID, iSearcherDB
'Else
'   ObjTour.Edit iSearcherDB
'
'End If
'
' 'setHeart("Heart0") = "Field not Used"
' ObjTour.DBSetField "Heart0", "Field not Used", iSearcherDB
'For I = 1 To 9
'    HeartNum = "Heart" & Format$(I, "0")
'    ObjTour.DBSetField HeartNum, HeaBouTxt(I).Text, iSearcherDB
'
'Next I

'ObjTour.Update iSearcherDB
 
' ---------------------------------------
' Write Type description to Levels Table
' ---------------------------------------
SQL = "SELECT * FROM Levels WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

LevelNum = 0

If ObjTour.RstRecordCount(iSearcherDB) = 0 Then
   ObjTour.AddNew iSearcherDB
   ObjTour.DBSetField "Id", objMdi.info.ID, iSearcherDB
Else
   ObjTour.Edit iSearcherDB
End If

For i = 1 To 5

    LevelNum = "Level" & Format$(i, "0")
    ObjTour.DBSetField LevelNum, cLevels(i), iSearcherDB
    
Next i

ObjTour.Update iSearcherDB

bHeartChanged = False
Unload HrtZone

Exit Sub
HeaSav_Err:
    MsgBox Error$(Err)
    If bDebug Then Handle_Err Err, "HeaSavCmd-HrtZone"
    Resume Next
End Sub


Private Sub HrtAbomnu_Click()
AboutMsg
End Sub


Private Sub HrtOptmnu_Click()

End Sub

Private Sub MidEximnu_Click()
HrtZone.Hide
End Sub


