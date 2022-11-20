VERSION 5.00
Begin VB.Form Graphfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graphical representation of user data."
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   1515
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GRAPHFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5250
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin Tourwin2002.UTextBox GraToTxt 
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
      Begin VB.CommandButton cmdCalendar 
         Height          =   255
         Index           =   1
         Left            =   960
         Picture         =   "GRAPHFRM.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin Tourwin2002.UTextBox GraFroTxt 
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
      Begin VB.CommandButton cmdCalendar 
         Height          =   255
         Index           =   0
         Left            =   960
         Picture         =   "GRAPHFRM.frx":052E
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox GraGraph 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   9315
      TabIndex        =   8
      Top             =   120
      Width           =   9375
   End
   Begin VB.Frame GraFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Graph Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   9375
      Begin VB.ComboBox GraGraCbo 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "|-1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox GraFieCbo 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "|-1"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Field Types:"
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GraphType:"
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date Range:"
         ForeColor       =   &H00400040&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   915
      End
   End
End
Attribute VB_Name = "Graphfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub LoadGraphCbo()
GraGraCbo.Clear
GraGraCbo.AddItem "2D Pie"
GraGraCbo.AddItem "3D Pie"
GraGraCbo.AddItem "2D Bar"
GraGraCbo.AddItem "3D Bar"
GraGraCbo.AddItem "Gantt"
GraGraCbo.AddItem "Line"
GraGraCbo.AddItem "Log/Lin"
GraGraCbo.AddItem "Area"
GraGraCbo.AddItem "Scatter"
GraGraCbo.AddItem "Polar"
GraGraCbo.AddItem "HLC"
End Sub

Sub RunGraph(DrawType As Integer)
On Local Error GoTo RunGraph_Err
Dim i As Integer, J As Integer, SQL As String
'Static StrtDt As String, EndDt As String
'Dim FirstDate As Long, RecCount As Integer
'Dim Msg As String, DateStr As String
'Dim lDBHandle As Long   ' Handle to database & recordset...
'
'GraGraph.DrawMode = 2
'GraGraph.GraphStyle = 2
'GraGraph.YAxisMin = 120
'GraGraph.YAxisMax = 160
'' Only if date range changes
'' does query need to be update.
'' ----------
'' ProgressBar
'' ---------------------------------------------
'
'If StrtDt <> GraToTxt.Text Or EndDt <> GraFroTxt.Text Then
'    StrtDt = Format$(GraFroTxt.Text, "MM-DD-YYYY")
'    EndDt = Format$(GraToTxt.Text, "MM-DD-YYYY")
'
'    SQL = "SELECT * FROM Dai WHERE Id = " & objMdi.info.ID & " and Date  BETWEEN #" & StrtDt & "# AND #" & EndDt & "#"
'
'    ObjTour.RstSQL lDBHandle, SQL
'    If lDBHandle = 0 Then
'        MsgBox LoadResString(gcOutOfDBHandles), vbOKOnly, LoadResString(gcTourVersion)
'        ObjTour.FreeHandle lDBHandle
'        Unload Me
'    End If
'
'End If
'' -------------------------
'' check for empty recordset
'' -------------------------
'' Move Last
'ObjTour.DBMoveLast (lDBHandle)
'
'If ObjTour.RstRecordCount(lDBHandle) <= 1 Then  'Check for records...
'        MsgBox LoadResString(gcNoRecords), vbOKOnly + vbCritical, LoadResString(gcTourVersion)
'        ObjTour.FreeHandle lDBHandle
'    Exit Sub
'End If
'' Move Last
'ObjTour.DBMoveLast (lDBHandle)
'
'' Check Record Count
'RecCount = ObjTour.RstRecordCount(lDBHandle) - 1
'
'GraGraph.NumPoints = RecCount
'GraGraph.GraphType = GraGraCbo.ListIndex + 1
'
'GraGraph.AutoInc = 1
''Let NumPoint 1 = x axis
'
'' Move First record
'ObjTour.DBMoveFirst lDBHandle
'
''Return date value
'FirstDate = CDate(ObjTour.DBGetField("Date", lDBHandle))
'GraGraph.AutoInc = 1
'
'' ----------------------------
'' Loop thru Daily recordset
'' and enter data into graph..
'' ----------------------------
'GraGraph.GraphTitle = GraFieCbo.Text & " data for " & StrtDt & " to " & EndDt
'For i = 1 To RecCount
'    If Val(ObjTour.DBGetField(GraFieCbo.Text, lDBHandle)) <> 0 Then
'        GraGraph.GraphData = Val(ObjTour.DBGetField(GraFieCbo.Text, lDBHandle))    'setDaily(GraFieCbo.Text)
'    Else
'        GraGraph.NumPoints = IIf(GraGraph.NumPoints - 1 >= 2, GraGraph.NumPoints - 1, 2)
'    End If
'
'    ' Move to next Record...
'    ObjTour.DBMoveNext lDBHandle
'
'    ' Update progessbar...
'    ProgressBar "Processing graph data...", -1, (i / RecCount) * 10, -1
'Next i
'
'ProgressBar "", 0, 0, 0
'GraGraph.DrawMode = DrawType        ' 2 = Screen, 5 = Printer,
'ObjTour.FreeHandle lDBHandle
'Close
Exit Sub
RunGraph_Err:
        'MsgBox Error$(Err), vbOKOnly, LoadResString(gcTourVersion)
        If bDebug Then Handle_Err Err, "RunGraph-Graphfrm"
        Err.Clear
        Exit Sub
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    CallFrom = "GraphFrm"
    
    DateFrm.Show vbModal
    If Not UserCancel Then
        GraToTxt.Text = DateFrm.DatToTxt.Text
        GraFroTxt.Text = DateFrm.DatFroTxt.Text
    End If

End Sub

Private Sub Form_Activate()
    Define_GraphFrm_mnu Loadmnu
End Sub

Private Sub Form_Deactivate()
    Define_GraphFrm_mnu Unloadmnu
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
' ------------------------------------
' Open database connect and wait for
' user to set graph parameters
' ------------------------------------

If bDebug Then Handle_Err 0, "Loading Graph Form."

' Update progessbar
ProgressBar "Loading Graph...", -1, 1, -1
CentreForm Graphfrm, -1
Me.KeyPreview = True
Define_GraphFrm_mnu Loadmnu

LoadGraphCbo
GraGraCbo.Text = GraGraCbo.List(0)

' Update progress bar
ProgressBar "Loading Graph...", -1, 3, -1

' ----------------------------
' Add Field item to graph by
' to list box...
' ----------------------------
GraFieCbo.AddItem "Heart"
GraFieCbo.AddItem "DaiMile"
ProgressBar "Loading Graph...", -1, 5, -1
GraFieCbo.AddItem "Weight"
ProgressBar "Loading Graph...", -1, 9, -1
GraFieCbo.Text = GraFieCbo.List(0)

' ------------------------
' Set date to last entered
' Start and End date
' ------------------------
If Not IsDate(GraFroTxt.Text) Then
gbSkipRegErrMsg = True ' Don't show reg error
GraFroTxt.Text = GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcDateFrmStart))
        If REG_ERROR = GraFroTxt.Text Then GraFroTxt.Text = Format$(Now, "MM-dd-yyyy")
        
gbSkipRegErrMsg = True ' Don't show reg error
GraToTxt.Text = GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcDateFrmEnd))
        If REG_ERROR = GraToTxt.Text Then GraToTxt.Text = Format$(Now, "MM-dd-yyyy")
        
End If

' Load setting from registry
ObjTour.Settings.ReadFormSettingsFromReg Me

ProgressBar "", 0, 0, 0
End Sub

Sub Define_GraphFrm_mnu(LoadType As String)
On Local Error GoTo Define_GraphFrm_Err

If Not objMdi.info.NewUser Then
Select Case LoadType
    Case Unloadmnu:
        MDI!MdiEximnu.Caption = MdiFrm_Exit
        MDI!MdiOptmnu.Caption = MdiFrm_Option
        MDI!MdiNewmnu.Caption = MdiFrm_Newmnu
        MDI!MdiSavmnu.Caption = MdiFrm_Savmnu
        MDI!MdiDelmnu.Caption = MdiFrm_Delmnu
        MDI!MdiOptmnu.Enabled = True
        MDI!MdiGramnu.Enabled = True
        MDI!MdiNewmnu.Enabled = True
        MDI!MdiSavmnu.Enabled = False
        MDI!MdiDelmnu.Enabled = False
        MDI!MdiBegmnu.Visible = False
        MDI!MdiPrimnu.Enabled = False
        
    Case Loadmnu:
        MDI!MdiEximnu.Caption = GraphFrm_Exit
        MDI!MdiOptmnu.Caption = GraphFrm_Option
        MDI!MdiNewmnu.Caption = GraphFrm_Newmnu
        MDI!MdiSavmnu.Caption = GraphFrm_Savmnu
        MDI!MdiDelmnu.Caption = GraphFrm_Delmnu
        MDI!MdiBegmnu.Caption = GraphFrm_Begmnu
        MDI!MdiOptmnu.Enabled = False
        MDI!MdiGramnu.Enabled = False
        MDI!MdiNewmnu.Enabled = False
        MDI!MdiSavmnu.Enabled = False
        MDI!MdiDelmnu.Enabled = False
        MDI!MdiBegmnu.Enabled = True
        MDI!MdiBegmnu.Visible = True
        MDI!MdiPrimnu.Enabled = True
        
End Select
End If
Exit Sub
Define_GraphFrm_Err:
    If bDebug Then Handle_Err Err, "Define_GraphFrm_mnu-GraphFrm"
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)

Define_GraphFrm_mnu Unloadmnu
ObjTour.Settings.WriteFormSettingsToReg Me

End Sub

Sub GraBegmnu()
' -------------------------------------
' Check for proper setup Ie. Date range
' Field selection, etc...
' -------------------------------------
RunGraph 2          '2 = Screen
End Sub


Private Sub GraFieCbo_GotFocus()
cmdCalendar(0).Visible = False
cmdCalendar(1).Visible = False

End Sub

Private Sub GraFroTxt_GotFocus()
cmdCalendar(0).Visible = True
cmdCalendar(1).Visible = False

End Sub

Private Sub GraFroTxt_LostFocus()

If Not IsDate(GraFroTxt.Text) Then
    MsgBox LoadResString(gcInvalidDate) & " '" & GraFroTxt.Text & "'", vbOKOnly + vbExclamation, LoadResString(gcInvalidFormat)
    GraFroTxt.SetFocus
End If

End Sub

Private Sub GraFroTxt_ToolTip()
MDI.StatusBar1.Panels(1).Text = GraFroTxt.ToolTipText
End Sub

Private Sub GraGraCbo_GotFocus()
cmdCalendar(0).Visible = False
cmdCalendar(1).Visible = False
End Sub

Private Sub GraToTxt_GotFocus()
cmdCalendar(0).Visible = False
cmdCalendar(1).Visible = True



End Sub

Private Sub GraToTxt_LostFocus()
If Not IsDate(GraToTxt.Text) Then
    MsgBox LoadResString(gcInvalidDate) & " '" & GraToTxt.Text & "'", vbOKOnly + vbExclamation, LoadResString(gcInvalidFormat)
    GraToTxt.SetFocus
End If

End Sub

Private Sub GraToTxt_ToolTip()
MDI.StatusBar1.Panels(1).Text = GraToTxt.ToolTipText
End Sub

Private Sub Label1_DblClick()
    CallFrom = "GraphFrm"
    
    DateFrm.Show vbModal
    If Not UserCancel Then
        GraToTxt.Text = DateFrm.DatToTxt.Text
        GraFroTxt.Text = DateFrm.DatFroTxt.Text
    End If
    
End Sub

Public Function PrintGraph() As Boolean
' /////////////////////////////////////////////////
' // Send current graph to printer...
' //
' ////////////////////////////////////////////////
    Me.PrintForm
    
End Function
