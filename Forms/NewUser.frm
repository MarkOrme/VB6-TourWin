VERSION 5.00
Begin VB.Form NewUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New User."
   ClientHeight    =   3795
   ClientLeft      =   1860
   ClientTop       =   1950
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3795
   ScaleWidth      =   7125
   Begin VB.Frame Frame4 
      Caption         =   "Security"
      Height          =   1455
      Left            =   3840
      TabIndex        =   17
      Top             =   240
      Width           =   3015
      Begin VB.TextBox NewPasTxt 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton NewChaCmd 
         Caption         =   "Change Pass&Word"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox NewPasChk 
         Caption         =   "&PassWord Protected."
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "PassWord:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      Height          =   1455
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   3495
      Begin VB.TextBox NewUseTxt 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "User Name."
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton NewFndCmd 
         Caption         =   "Fin&d"
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox NewDatTxt 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data location:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "User Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DeskTop"
      Height          =   1335
      Left            =   3840
      TabIndex        =   12
      Top             =   1800
      Width           =   3015
      Begin VB.CheckBox NewMetOpt 
         Caption         =   "Show metafile"
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton NewFinCmd 
         Caption         =   "&Find"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox NewMetTxt 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Metfile for DeskTop"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Start Up Options."
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   3500
      Begin VB.OptionButton NewNotOpt 
         Caption         =   "Load nothing on StartUp."
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton NewCalOpt 
         Caption         =   "Load Calendar Form on StartUp."
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton NewGraOpt 
         Caption         =   "Load Graph Form on StartUp."
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton NewDaiOpt 
         Caption         =   "Load Daily Form on StartUp."
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.CommandButton NewCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton NewSavCmd 
      Caption         =   "&Save and Exit"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Menu MiOFilmnu 
      Caption         =   "&File"
      Begin VB.Menu NewOptmnu 
         Caption         =   "Option, StartUp..."
      End
      Begin VB.Menu NewSepmnu 
         Caption         =   "-"
      End
      Begin VB.Menu NewPStmnu 
         Caption         =   "&Print SetUp"
      End
      Begin VB.Menu NewPrimnu 
         Caption         =   "&Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu NewSe4mnu 
         Caption         =   "-"
      End
      Begin VB.Menu NewdEximnu 
         Caption         =   "E&xit Options..."
      End
   End
   Begin VB.Menu NewEdimnu 
      Caption         =   "&Edit"
      Begin VB.Menu NewNewmnu 
         Caption         =   "&New User"
      End
      Begin VB.Menu NewSavmnu 
         Caption         =   "&Save User"
      End
      Begin VB.Menu NewDelmnu 
         Caption         =   "&Delete User"
      End
   End
   Begin VB.Menu NewDaimnu 
      Caption         =   "&DataBase"
      Begin VB.Menu NewDiamnu 
         Caption         =   "Da&ily Databases"
      End
      Begin VB.Menu NewCalmnu 
         Caption         =   "&Calendar Database"
      End
      Begin VB.Menu Newnicmnu 
         Caption         =   "Concon&i Database"
         Visible         =   0   'False
      End
      Begin VB.Menu NewCntmnu 
         Caption         =   "&Contacts Database"
      End
   End
   Begin VB.Menu NewToomnu 
      Caption         =   "&Tools"
      Begin VB.Menu NewRepmnu 
         Caption         =   "&Reports"
         Begin VB.Menu NewDarmnu 
            Caption         =   "&Daily Report"
         End
         Begin VB.Menu NewWeemnu 
            Caption         =   "&Weekly Report"
            Begin VB.Menu MdoTotmmu 
               Caption         =   "Weekly ( T&otals Report )"
            End
            Begin VB.Menu NewPermnu 
               Caption         =   "Weekly ( &Percentage )"
            End
         End
         Begin VB.Menu NewCndmnu 
            Caption         =   "&Calendar"
         End
         Begin VB.Menu NewCnlmnu 
            Caption         =   "Contact &Listing"
         End
      End
      Begin VB.Menu NewGramnu 
         Caption         =   "&Graph"
      End
      Begin VB.Menu MdoMaimnu 
         Caption         =   "Database &Maintenance"
      End
   End
   Begin VB.Menu NewSetmnu 
      Caption         =   "&SetUp"
      Begin VB.Menu NewSysmnu 
         Caption         =   "System &Names"
         Begin VB.Menu NewGuimnu 
            Caption         =   "&Peak Guide Names"
         End
         Begin VB.Menu NewTypmnu 
            Caption         =   "&Event Type Names"
         End
         Begin VB.Menu NewHeamnu 
            Caption         =   "&Heart Rate Zone"
         End
      End
      Begin VB.Menu NewCycmnu 
         Caption         =   "System Peak S&chedules"
      End
   End
   Begin VB.Menu NewHlpmnu 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu NewConmnu 
         Caption         =   "&Contents"
      End
      Begin VB.Menu NewSeamnu 
         Caption         =   "&Search for help on..."
      End
      Begin VB.Menu NewHspmnu 
         Caption         =   "-"
      End
      Begin VB.Menu NewTecmnu 
         Caption         =   "&Technical Support"
      End
      Begin VB.Menu NewSe3mnu 
         Caption         =   "-"
      End
      Begin VB.Menu NewAbomnu 
         Caption         =   "&About TourWin cycling Program"
      End
   End
End
Attribute VB_Name = "NewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Changed As Integer
Dim dbUser As Database, setUser As Recordset
Dim CurrentName As String
Dim bAddUser As Boolean   ' Flags when new user menu has been choosen.

Sub AddUserNamesToList()
On Local Error Resume Next
setUser.MoveFirst
While Not setUser.EOF
    NewUseTxt.AddItem setUser("Name")
    setUser.MoveNext
Wend
End Sub

Function CheckInput()
' ----------------------------
' Function is call to ensure
' all fields have valid values
' ----------------------------
CheckInput = -1
If NewDatTxt.Text = "" Then
    MsgBox "Data location: Not allow zero length!", vbCritical, LoadResString(gcTourVersion)
    NewDatTxt.SetFocus
    CheckInput = 0
    Exit Function
End If
'If NewPasTxt.Text = "" Then
'    MsgBox "PassWord: Not allow zero length!", vbCritical, LoadResString(gcTourVersion)
'    NewPasTxt.SetFocus
'    CheckInput = 0
'    Exit Function
'End If
End Function

Sub ClearControls()
NewPasChk.Value = 0
NewNotOpt.Value = True
NewDatTxt = " "
NewUseTxt = " "
NewPasTxt = " "
NewMetTxt = " "
End Sub

Sub CreateContactOpt(lID As Long)

On Local Error GoTo CreateContactOpt_Err

If bDebug Then Handle_Err 0, "CreateContactOpt-MdiOpt"

Dim SQL As String
' ----------------------
' Create default record
' ----------------------
SQL = "INSERT INTO " & gcUserTour_ContactOpt & " (ID, IndexOrder,ColumnOrder,SortOrder,ContactWidth," & _
      "LastWidth,FirstWidth,PhoneWidth,FaxWidth,[E-MailWidth]) VALUES (" & Str$(lID) & ",'Contact','0123456','ASC',1020,1020,1020,1020,1020,1020)"
      
ObjTour.DBExecute SQL

Exit Sub
CreateContactOpt_Err:
    If bDebug Then
        MsgBox Error$(Err)
        If bDebug Then Handle_Err Err, "CreateContactOpt-MdiOpt"
    End If
    Resume Next
    
End Sub

Sub CreateDailyOpt(lID As Long)
On Local Error GoTo CreateDailyOpt_Err
If bDebug Then Handle_Err 0, "Set_Check_Setting-DailyOpt"

Dim SQL As String

' ----------------------
' Create default record
' ----------------------
SQL = "INSERT INTO " & gcUserTour_DailyOpt & " (ID, Description,[Rest Heart],Sleep,DayInt," & _
      "Weight,Metric) VALUES (" & Str$(lID) & ",True,True,True,True,True,True)"

ObjTour.DBExecute SQL

Exit Sub
CreateDailyOpt_Err:

    If bDebug Then Handle_Err Err, "CreateDailyOpt-MdiOpt"

    Resume Next

End Sub

Function Determine_New_Or_Existing() As String
On Local Error GoTo Determine_Err
Dim MsgStr As String
    
Determine_New_Or_Existing = "Existing User"

MsgStr = " Yes for NEw,No for existing 'Name' " & CurrentName
If vbYes = MsgBox(MsgStr, vbYesNo + vbCritical, LoadResString(gcTourVersion)) Then
    Determine_New_Or_Existing = "New User"
End If
Exit Function
Determine_Err:
    If bDebug Then Handle_Err Err, "Determine_New_Or_Existing"
    Resume Next
End Function

Sub User_Scr_to_Data()
On Local Error GoTo FirstUser_Err

        ' Update PassWord setting
        '"Id" is autoIncr by jet database...
        ObjTour.DBSetField "Security", IIf(NewPasChk.Value = 1, True, False), iSearcherDB
        If NewNotOpt.Value = True Then
           ObjTour.DBSetField "Load", 0, iSearcherDB
        ElseIf NewDaiOpt.Value = True Then
           ObjTour.DBSetField "Load", 1, iSearcherDB
        ElseIf NewGraOpt.Value = True Then
           ObjTour.DBSetField "Load", 2, iSearcherDB
        ElseIf NewCalOpt.Value = True Then
           ObjTour.DBSetField "Load", 3, iSearcherDB
        End If
        
        ObjTour.DBSetField "Name", Trim$(NewUseTxt), iSearcherDB

        ObjTour.DBSetField "DataPath", Trim$(NewDatTxt), iSearcherDB
        ObjTour.DBSetField "PassWord", NewPasTxt, iSearcherDB
        ObjTour.DBSetField "MetaFile", IIf(IsNull(NewMetTxt), " ", NewMetTxt), iSearcherDB
        ObjTour.DBSetField "ShowMeta", NewMetOpt.Value, iSearcherDB
        
Exit Sub
FirstUser_Err:
    If bDebug Then Handle_Err Err, "User_Scr_To_Data-MdiOpt"
    Resume Next
End Sub

Sub Setformvalues(NewOrOld As Integer)
' -------------------
' Take user value and
' Setup controls
' --------------------
If NewOrOld = 0 Then
NewPasChk.Value = 0
If setUser("Security") = True Then NewPasChk.Value = 1
NewNotOpt.Value = IIf(setUser("Load") = 0, True, False)
NewDaiOpt.Value = IIf(setUser("Load") = 1, True, False)
NewGraOpt.Value = IIf(setUser("Load") = 2, True, False)
NewCalOpt.Value = IIf(setUser("Load") = 3, True, False)
NewDatTxt = setUser("DataPath")
NewUseTxt = setUser("Name")
NewPasTxt = setUser("PassWord")
NewMetTxt = setUser("MetaFile")
NewMetOpt.Value = setUser("ShowMeta")
Else
NewPasChk.Value = 0
NewNotOpt.Value = True
NewDatTxt = ""
NewUseTxt = ""
NewPasTxt = ""
NewMetTxt = " "
NewMetOpt.Value = False
End If
End Sub

Sub Setup_UserTour_tables()
'
'
On Local Error GoTo SetUp_UserTour_Err
CreateTourDatabases NewDatTxt, gcDai_Tour
CreateTourDatabases NewDatTxt, "ContTour.mdb"
CreateTourDatabases NewDatTxt, "ConcTour.mdb"
CreateTourDatabases NewDatTxt, "NameTour.mdb"
CreateTourDatabases NewDatTxt, gcPeakTour
CreateTourDatabases NewDatTxt, "Eve_Tour.mdb"
Exit Sub
SetUp_UserTour_Err:
    If bDebug Then Handle_Err Err, "Setup_UserTour_Tables-MdiOpt"
    Resume Next
End Sub

Function UpdateTourUser_UserTbl() As Integer
' --------------------------------------------------
' Purpose: To destiguish between three possiblities
' 1.) First User, Must create database....
' 2.) Second New User, .AddNew....
' 3.) Three, Update exist user, .Edit...
' --------------------------------------------------
On Local Error GoTo Update_Err
Dim SQL As String, ProgressType As String, MsgStr As String
Dim RetInt As Integer, DPath As String, NotResolved As Boolean
Dim RetStr As String

UpdateTourUser_UserTbl = 0 ' Assume pesimistic
' ---------------------------
' Check required field before
' letting user continue.
' ---------------------------
If Not CheckInput Then Exit Function
    
' ------------------
' Define SearcherDB
' ------------------
SQL = "Select * From " & gcUserTour_UserTbl
ObjTour.RstSQL iSearcherDB, SQL

' ------------------------
' Try and find exist user
' with same name
' ------------------------
SQL = "Name = '" & NewUser.NewUseTxt & "'"
ObjTour.DBFindFirst SQL, iSearcherDB
            
' ---------------------------------
' Determine if it is save
' to create new user or
' does user name already exist?
' ---------------------------------
If ObjTour.NoMatch(iSearcherDB) Then
    
    ObjTour.AddNew iSearcherDB  ' Add new record
    User_Scr_to_Data            ' Update Dialog fields database
    ObjTour.Update iSearcherDB  ' Update record

    ' --------------------------
    ' Define newly created user
    ' --------------------------
    SQL = "Name = '" & NewUser.NewUseTxt & "'"
    ObjTour.DBFindFirst SQL, iSearcherDB
    CreateDailyOpt ObjTour.DBGetField("ID", iSearcherDB)
    CreateContactOpt ObjTour.DBGetField("ID", iSearcherDB)
Else
    MsgBox "Current user already exist! Either edit existing enter by going to File opions or " & vbCrLf & _
    "change User name.", vbOKOnly, LoadResString(gcTourVersion)
    Exit Function
End If


' Create default setting for First User
' ---------

    If ObjNew.info.NewUser Then
        ' First User...
      setUser.MoveFirst
      'CreateDailyOpt
      'CreateContactOpt
      objMdi.info.Name = Trim$(NewUseTxt)
      objMdi.info.ID = setUser("Id")
      objMdi.info.Datapath = Trim$(NewDatTxt)
'      PassFrm.GetDailyOpt
'      PassFrm.GetContactOpt
      Setup_UserTour_tables       'Create Daily,Peak... Databases...
    End If
    
'Prompt User

    MsgStr = "Do you wish to load new user now?"
    RetInt = MsgBox(MsgStr, vbYesNo + vbQuestion, LoadResString(gcTourVersion))
            If RetInt = vbYes Then
                'Include all obj settings
                objMdi.info.Name = ObjTour.DBGetField("Name", iSearcherDB)
                objMdi.info.ID = ObjTour.DBGetField("Id", iSearcherDB)
                objMdi.info.Datapath = Trim$(NewDatTxt)
            End If

If bDebug Then Handle_Err 0, "UpdateTourUser-MdiOpt"
ProgressBar "", 0, 0, 0
UpdateTourUser_UserTbl = -1
Exit Function
Update_Err:
    If bDebug Then Handle_Err Err, "UpdateTourUser-MdiOpt"
    Resume Next
    
End Function

Private Sub Form_GotFocus()
MdiOpt.WindowState = 0
CentreForm MdiOpt, -1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo MdiLoad_Err
Dim SQL As String, ProgressType As String
' -----------------------------------------------
' This Form replace the dual purpose of MdoFrm
' NewUser is allow called from startup therefore
' many assumetino are allowed.
' -----------------------------------------------
CentreForm NewUser, -1
bAddUser = False  'Flags new menu currently not selected...
Me.KeyPreview = True
Changed = 0
Exit Sub
MdiLoad_Err:
    If bDebug Then Handle_Err Err, "Form_Load-NewUser"
    Resume Next
End Sub


Private Sub NewAbomnu_Click()
AboutMsg
End Sub

Private Sub NewCalOpt_Click()
Changed = -1
End Sub

Private Sub NewCanCmd_Click()
On Local Error GoTo NewCanCmd_Err
Dim RetInt As Integer
' ------------------------------------------
' If    NewUser = true then
' Kill UserTour.mdb so password form will
' not appear as no user really exist yet!!!
' ------------------------------------------
If Changed Then
  RetInt = MsgBox("Do you wish to save changes.", vbYesNo + vbQuestion, LoadResString(gcTourVersion))
  If RetInt = vbYes Then NewSavCmd_Click
End If
  Datapath = objMdi.info.Datapath

setUser.Close
dbUser.Close
Unload NewUser
Exit Sub
NewCanCmd_Err:
    If bDebug Then Handle_Err Err, "NewCanCmd_Err-MdiOpt"
    Resume Next
End Sub


Private Sub NewChaCmd_Click()
NewPasTxt.Enabled = -1
NewPasTxt.SetFocus
End Sub

Private Sub NewDaiOpt_Click()
Changed = -1
End Sub

Private Sub NewDatTxt_Change()
Changed = -1
End Sub

Private Sub NewDelmnu_Click()
On Local Error GoTo NewDelMnu_Err
Dim MsgStr As String, RetInt As Integer, SQL As String

MsgStr = "Are you sure you wish " & vbLf
MsgStr = MsgStr & "to delete: " & NewUseTxt
RetInt = MsgBox(MsgStr, vbYesNo + vbQuestion, LoadResString(gcTourVersion))
If RetInt = vbYes Then
    ' Remove from List, TableUser,DailyOpt,Contopt...
    SQL = "DELETE * FROM DailyOpt WHERE Name = '" & Trim$(NewUseTxt) & "'"
    dbUser.Execute SQL
    SQL = "DELETE * FROM ContactOpt WHERE Name = '" & Trim$(NewUseTxt) & "'"
    dbUser.Execute SQL
    SQL = "DELETE * FROM UserTbl WHERE Name = '" & Trim$(NewUseTxt) & "'"
    dbUser.Execute SQL
    
    CurrentName = Trim$(NewUseTxt)
    SQL = "Name = '" & NewUseTxt & "'"
    setUser.FindFirst SQL
    CurrentName = Trim$(NewUseTxt)
    Setformvalues 0

    
End If
Exit Sub
NewDelMnu_Err:
    If bDebug Then Handle_Err Err, "NewDelmnu-MdiOpt"
    Resume Next
End Sub

Private Sub NewdEximnu_Click()
Unload MdiOpt
dbUser.Close
setUser.Close
End Sub

Private Sub NewFinCmd_Click()
On Local Error GoTo NewFin_Err
If bDebug Then Handle_Err 0, "NewFinCmd_Click-MdiOpt"


Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    
    FileDialog.sFilter = "Metafiles (*.wmf)" & Chr$(0) & "*.wmf" & Chr$(0) & "Bitmap (*.bmp)" & Chr$(0) & "*.bmp"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    FileDialog.sDlgTitle = "Open"
    FileDialog.sInitDir = NewMetTxt.Text
    sOpen = ShowOpen(Me.hWnd)
    
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        NewMetTxt = FileDialog.sFile
    End If


Exit Sub
NewFin_Err:
    Select Case Err
        Case 32755:
            Exit Sub     'User Cancel
        Case Else
            If bDebug Then Handle_Err Err, "NewFin-MdiOpt"
    End Select
End Sub

Private Sub NewFndCmd_Click()
On Local Error GoTo NewFin_Err
Dim RetStr As String
If bDebug Then Handle_Err 0, "NewFndCmd_Click-MdiOpt"
MssngFrm.MssDatTxt.Text = CurDir
'MssngFrm.MssUseTxt = NewUseTxt
MssngFrm.Show 1

'If Not ObjNew.info.NewUser Then RetStr = Check_Data_Exist(Datapath)

If RetStr = "Canceled" Then Exit Sub
NewDatTxt = Datapath
Exit Sub
NewFin_Err:
    Select Case Err
        Case 32755:
            Exit Sub     'User Cancel
        Case Else
            If bDebug Then Handle_Err Err, "NewFin-MdiOpt"
    End Select
End Sub

Private Sub NewGraOpt_Click()
Changed = -1
End Sub
'Private Sub NewMetOpt_Click(Value As Integer)
'Changed = -1
'End Sub

Private Sub NewMetOpt_Click()
Changed = -1
End Sub

Private Sub NewMetTxt_Change()
Changed = -1
End Sub

Private Sub NewPasTxt_GotFocus()
NewPasTxt.SelStart = 0
NewPasTxt.SelLength = Len(NewPasTxt.Text)
End Sub

Private Sub NewUseTxt_KeyPress(KeyAscii As Integer)
' ------------------------
' Allows only 15 character
' length user name.
' ------------------------
If Len(NewUseTxt.Text) >= 15 Then
    Beep
    MsgBox "Maximun user name exceeded.", vbOKOnly, LoadResString(gcTourVersion)
    KeyAscii = 0
End If
End Sub

Private Sub NewNewmnu_Click()
' ----------------------------------------
' Check for change and prompt user to save
' ----------------------------------------

If bDebug Then Handle_Err 0, "NewNewmnu-MdiOpt"
On Local Error GoTo NewErr

If Changed Then
    If vbYes = MsgBox("Save changes first?", vbYesNo + vbQuestion, "Save changes") Then UpdateTourUser_UserTbl
End If
ClearControls

bAddUser = True  ' Flags new menu choosen...
NewUseTxt.Visible = True
NewUseTxt.Visible = False

Exit Sub
NewErr:
    If bDebug Then Handle_Err Err, "NewNewmnu-MdiOpt"
    Resume Next
End Sub

Private Sub NewNotOpt_Click()
Changed = -1
End Sub

Private Sub NewPasChk_Click()
Changed = -1
End Sub

Private Sub NewPasTxt_Change()
Changed = -1
End Sub

Private Sub NewSavCmd_Click()

On Local Error GoTo NewSavCmd_Err

If bDebug Then Handle_Err 0, "NewSavCmd_Click-MdiOpt"
    ' Test if all fields were entered...
    If Not UpdateTourUser_UserTbl Then
        Exit Sub
    End If
    ObjNew.info.NewUser = False  ' New user has been created.

bAddUser = False   ' Set new menu clicked flag to false
Unload Me

Exit Sub
NewSavCmd_Err:
    If bDebug Then Handle_Err Err, "NewSavCmd_Err-NewUser"
    Resume Next
End Sub


Private Sub NewSavmnu_Click()
On Local Error GoTo NewSav_Err
' -----------------------
' Update TourUser/UserTbl
' -----------------------
If Changed Then UpdateTourUser_UserTbl

bAddUser = False  ' Set new menu clicked flag to false
NewUseTxt.Visible = False
NewUseTxt.Visible = True

Exit Sub
NewSav_Err:
    If bDebug Then Handle_Err Err, "NewSavmnu-MdiOpt"
    Resume Next
End Sub

Private Sub NewTecmnu_Click()
TechSupport
End Sub


Private Sub NewUseTxt_GotFocus()
NewUseTxt.SelStart = 0
NewUseTxt.SelLength = Len(NewUseTxt.Text)
End Sub


