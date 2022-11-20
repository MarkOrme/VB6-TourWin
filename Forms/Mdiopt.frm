VERSION 5.00
Begin VB.Form MdiOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Setup Options."
   ClientHeight    =   3780
   ClientLeft      =   2505
   ClientTop       =   3000
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mdiopt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      Height          =   975
      Left            =   255
      TabIndex        =   14
      Top             =   240
      Width           =   6615
      Begin VB.CommandButton MdOChaCmd 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         ToolTipText     =   "Change password"
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox MdOPasChk 
         Alignment       =   1  'Right Justify
         Caption         =   "Securit&y ON"
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox MdONamTxt 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "User Name"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblSetPWD 
         Caption         =   "Set Passwo&rd"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "&Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " User Settings "
      Height          =   1815
      Left            =   255
      TabIndex        =   13
      Top             =   1320
      Width           =   6615
      Begin VB.CheckBox MdoMetOpt 
         Alignment       =   1  'Right Justify
         Caption         =   "Sho&w metfile"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Show image"
         Top             =   1080
         Width           =   1875
      End
      Begin VB.CommandButton MdOFinCmd 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   10
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox MdOMetTxt 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         ToolTipText     =   "Location of image"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cboMDOStartup 
         Height          =   315
         ItemData        =   "Mdiopt.frx":000C
         Left            =   1920
         List            =   "Mdiopt.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Startup option"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblMetafile 
         Caption         =   "Des&ktop image :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "At start&up:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton MdOCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   3250
      Width           =   855
   End
   Begin VB.CommandButton MdOSavCmd 
      Caption         =   "S&ave"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   3250
      Width           =   855
   End
End
Attribute VB_Name = "MdiOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Changed     As Integer
Dim dbUser      As Database
Dim setUser     As Recordset
Dim CurrentName As String
Dim bAddUser    As Boolean   ' Flags when new user menu has been choosen.
Private m_sPWD  As String

Sub AddUserNamesToList()
On Local Error Resume Next

ObjTour.DBMoveFirst iSearcherDB

While Not ObjTour.EOF(iSearcherDB)

    MdONamTxt.AddItem ObjTour.DBGetField("Name", iSearcherDB)
    ObjTour.DBMoveNext iSearcherDB
Wend

End Sub

Function CheckInput()
' ----------------------------
' Function is call to ensure
' all fields have valid values
' ----------------------------
CheckInput = -1
'If MdoDatTxt.Text = "" Then
'    MsgBox "Data location: Not allow zero length!", vbCritical, LoadResString(gcTourVersion)
'    MdoDatTxt.SetFocus
'    CheckInput = 0
'    Exit Function
'End If

End Function

Sub ClearControls()
MdOPasChk.Value = 0
cboMDOStartup.ListIndex = 0
MdONamTxt = " "
MdOMetTxt = " "
End Sub

Sub CreateContactOpt(lID As Long)

On Local Error GoTo CreateContactOpt_Err
If bDebug Then Handle_Err 0, "CreateContactOpt-MdiOpt"

Dim SQL As String

SQL = "SELECT * FROM ContactOpt"
ObjTour.RstSQL iSearcherDB, SQL

' ------------------------------------
' Create new record in ContactOpt table
' with default values.
' ------------------------------------

ObjTour.AddNew iSearcherDB
ObjTour.DBSetField "Id", lID, iSearcherDB
ObjTour.DBSetField "IndexOrder", "Contact", iSearcherDB
ObjTour.DBSetField "ColumnOrder", "0123456", iSearcherDB
ObjTour.DBSetField "SortOrder", "ASC", iSearcherDB
ObjTour.DBSetField "ContactWidth", 1020, iSearcherDB
ObjTour.DBSetField "LastWidth", 1020, iSearcherDB
ObjTour.DBSetField "FirstWidth", 1020, iSearcherDB
ObjTour.DBSetField "PhoneWidth", 1020, iSearcherDB
ObjTour.DBSetField "FaxWidth", 1020, iSearcherDB
ObjTour.DBSetField "E-MailWidth", 1020, iSearcherDB
ObjTour.Update iSearcherDB

Exit Sub
CreateContactOpt_Err:
    If bDebug Then Handle_Err Err, "CreateContactOpt-MdiOpt"
    Resume Next
    
End Sub

Sub CreateDailyOpt(lID As Long)
On Local Error GoTo CreateDailyOpt_Err

If bDebug Then Handle_Err 0, "Set_Check_Setting-DailyOpt"

Dim SQL As String


SQL = "SELECT * FROM DailyOpt"
ObjTour.RstSQL iSearcherDB, SQL

' ------------------------------------
' If no match to user name, try again.
' ------------------------------------
ObjTour.AddNew iSearcherDB

' Figure out how to get Id number set by AutoIncr

    ObjTour.DBSetField "Id", lID, iSearcherDB
    ObjTour.DBSetField "Description", True, iSearcherDB
    ObjTour.DBSetField "Rest Heart", True, iSearcherDB
    ObjTour.DBSetField "Sleep", True, iSearcherDB
    ObjTour.DBSetField "DayInt", True, iSearcherDB
    ObjTour.DBSetField "Weight", True, iSearcherDB
    ObjTour.DBSetField "Metric", True, iSearcherDB
    
    ObjTour.Update iSearcherDB
    
On Local Error GoTo 0
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
    ObjTour.DBSetField "Security", IIf(MdOPasChk.Value = 1, True, False), iSearcherDB
    
    ObjTour.DBSetField "Load", cboMDOStartup.ListIndex, iSearcherDB
    ObjTour.DBSetField "Name", Trim$(MdONamTxt), iSearcherDB
    ObjTour.DBSetField "PassWord", PWD, iSearcherDB
    ObjTour.DBSetField "MetaFile", IIf(IsNull(MdOMetTxt), " ", MdOMetTxt), iSearcherDB
    ObjTour.DBSetField "ShowMeta", MdoMetOpt.Value, iSearcherDB
    ObjTour.DBSetField "BitField", CLng(1), iSearcherDB
    
On Local Error GoTo 0
Exit Sub
FirstUser_Err:
    If bDebug Then Handle_Err Err, "User_Scr_To_Data-MdiOpt"
    Resume Next
End Sub

Sub Setformvalues(lNewOrOld As Long)
' -------------------
' Take user value and
' Setup controls
' --------------------
On Local Error GoTo SetFormValues_Err

Const CURRENTUSER = 0

If CURRENTUSER = lNewOrOld Then
    
    
    ' ----------------------------------------------
    ' toggle on and off security to set accompanying
    ' controls to proper enabled state.
    ' ----------------------------------------------
    MdOPasChk.Value = 1
    MdOPasChk.Value = Abs(CLng(objMdi.info.Security))
    
    
    cboMDOStartup.ListIndex = objMdi.info.Load
    MdONamTxt = objMdi.info.Name
    PWD = objMdi.info.Password  ' Load Password into Property
    MdOMetTxt = objMdi.info.MetaFile
    MdoMetOpt.Value = Abs(CLng(objMdi.info.ShowMeta))
    

Else

    ' Toggle to set appropriate enable state...
    MdOPasChk.Value = 1
    MdOPasChk.Value = 0
    
    MdoMetOpt.Value = 1
    MdoMetOpt.Value = 0
    
    cboMDOStartup.ListIndex = 0
    MdONamTxt = "<Enter Name>"
    MdOMetTxt = " "

    Me.Caption = "New user setup"
    Me.MdONamTxt.Locked = False
    
End If

On Error GoTo 0
Exit Sub

SetFormValues_Err:
If bDebug Then
    If bDebug Then Handle_Err Err, "MDIOpt-SetFormValues"
End If
Resume Next

End Sub


Function UpdateTourUser_UserTbl() As Boolean
' --------------------------------------------------
' Purpose: To destingish between three possiblities
' 1.) First User, Must create database....
' 2.) Second New User, .AddNew....
' 3.) Three, Update exist user, .Edit...
' --------------------------------------------------
On Local Error GoTo Update_Err
Dim lLoop       As Long
Dim sTemp       As String
Dim sRetStr     As String

UpdateTourUser_UserTbl = False ' Assume pesimistic
' ---------------------------------------------------
' Check required field before letting user continue.
' ---------------------------------------------------
If Not CheckInput Then Exit Function

If objMdi.info.NewUser Then
    ' First determine if Tourwin.mdb database exists
    ' in specified folder.
'    If "" = Dir$(MdiOpt.MdoDatTxt.Text & "\Tourwin.mdb") Then
'        ' Create New Database
'        CreateTourDatabases MdiOpt.MdoDatTxt.Text, "TourWin.mdb"
'
'        ' Release all current connects to current db
'        For lLoop = 10 To 1 Step -1
'            ObjTour.FreeHandle lLoop
'        Next lLoop
'
'        ' 11 will close all recordsets and db connects.
'        ObjTour.DBClose 11
'
'    End If
    
    ' Change global datapath so new user info
    ' can be written. Based on CreateNewUser return value
    ' global datapath value may need to be restored.
    sTemp = objMdi.info.Datapath
'    objMdi.info.Datapath = MdiOpt.MdoDatTxt.Text
        
    ' Write new datapath to Registry...
    ' Check if current DB is listed in DBase key, if not add!
    gbSkipRegErrMsg = True
    sRetStr = GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcTourDBase))
    
    ' Check if key exist, if not create
    If REG_ERROR <> sRetStr Then
        ' -------------------------
        ' Check if current db is listed in return string if not, append to end.
        ' -------------------------
        If 0 = InStr(1, sRetStr, objMdi.info.Datapath, vbTextCompare) Then
            gbSkipRegErrMsg = True
            WriteRegStringValue LoadResString(gcRegTourKey), LoadResString(gcTourDBase), sRetStr & "," & objMdi.info.Datapath
        End If
    End If
    
    ' False means do not load new user, stay with existing...
    If False = CreateNewUser() Then
        objMdi.info.Datapath = sTemp
    Else
        cPeakNames.IsUpdated = False
        objMdi.info.WelcomeWizard = True
    End If

    
    UpdateTourUser_UserTbl = True
Else
    UpdateTourUser_UserTbl = UpdateExistingUser()
End If

If bDebug Then Handle_Err 0, "UpdateTourUser-MdiOpt"
ProgressBar "", 0, 0, 0

On Local Error GoTo 0
Exit Function

Update_Err:
If bDebug Then
    MsgBox Error$(Err)
    If bDebug Then Handle_Err Err, "UpdateTourUser-MdiOpt"
End If
    Resume Next
    
End Function

Private Sub cboMDOStartup_Change()
Changed = -1
End Sub

Private Sub cboMDOStartup_Click()
Changed = -1
End Sub



Private Sub Form_Activate()
    Define_Form_menu Me.Name, Loadmnu
End Sub

Private Sub Form_GotFocus()
Define_Form_menu Me.Name, Loadmnu
MdiOpt.WindowState = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo MdiLoad_Err

Me.KeyPreview = True
CurrentName = Trim$(objMdi.info.Name)
CentreForm Me, 0
' Check if existing or new user info...
If objMdi.info.NewUser Then
    Setformvalues 1 ' New user...
Else
    Setformvalues 0 ' Existing user...
End If

Changed = 0
On Error GoTo 0
Exit Sub

MdiLoad_Err:
    If bDebug Then Handle_Err Err, "Form_Load-MdiOpt"
    Resume Next
End Sub


Private Sub Form_Terminate()
On Local Error GoTo Form_Error

On Local Error GoTo 0
Exit Sub
Form_Error:
    If bDebug Then Handle_Err Err, "Form_Terminate_Err-MdiOpt"
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Changed = -1 Then MdOCanCmd_Click
        
    Define_Form_menu Me.Name, Unloadmnu
End Sub

Sub MdOCanCmd_Click()
On Local Error GoTo MdOCanCmd_Err
Dim RetInt As Integer
' ------------------------------------------
' If    NewUser = true then
' Kill UserTour.mdb so password form will
' not appear as no user really exist yet!!!
' ------------------------------------------
If Changed Then
  RetInt = MsgBox("Do you wish to save changes.", vbYesNo + vbQuestion, LoadResString(gcTourVersion))
  If RetInt = vbYes Then
    MdOSavCmd_Click
  Else
    Changed = 0
  End If
End If
  Datapath = objMdi.info.Datapath

Unload MdiOpt
Exit Sub
MdOCanCmd_Err:
    If bDebug Then Handle_Err Err, "MdOCanCmd_Err-MdiOpt"
    Resume Next
End Sub


Private Sub MdOChaCmd_Click()

' Show Password dialog
With frmPassword
    .Password = ""
    .Show vbModal
    If .Password <> "" Then
        PWD = .Password
    End If
End With
Unload frmPassword

End Sub

Private Sub MdoDatTxt_Change()
Changed = -1
End Sub

Sub MdODelmnu_Click()
On Local Error GoTo MdODelMnu_Err
Dim MsgStr As String, RetInt As Integer, SQL As String

MsgStr = "Are you sure you wish " & vbLf
MsgStr = MsgStr & "to delete: " & MdONamTxt
RetInt = MsgBox(MsgStr, vbYesNo + vbQuestion, LoadResString(gcTourVersion))
If RetInt = vbYes Then
    ' Remove from List, TableUser,DailyOpt,Contopt...
    SQL = "DELETE * FROM DailyOpt WHERE Name = '" & Trim$(MdONamTxt) & "'"
    dbUser.Execute SQL
    SQL = "DELETE * FROM ContactOpt WHERE Name = '" & Trim$(MdONamTxt) & "'"
    dbUser.Execute SQL
    SQL = "DELETE * FROM UserTbl WHERE Name = '" & Trim$(MdONamTxt) & "'"
    dbUser.Execute SQL
    
    CurrentName = Trim$(MdONamTxt)
    SQL = "Name = '" & MdONamTxt & "'"
    
    PassFrm.Show
    
End If
Exit Sub
MdODelMnu_Err:
    If bDebug Then Handle_Err Err, "MdODelmnu-MdiOpt"
    Resume Next
End Sub

Sub MdOdEximnu_Click()
Unload MdiOpt
dbUser.Close
setUser.Close
End Sub


Private Sub MdOFinCmd_Click()
On Local Error GoTo MdoFin_Err
If bDebug Then Handle_Err 0, "MdOFinCmd_Click-MdiOpt"

Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    
    FileDialog.sFilter = "Metafiles (*.wmf)" & Chr$(0) & "*.wmf" & Chr$(0) & "Bitmap (*.bmp)" & Chr$(0) & "*.bmp"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.sInitDir = MdOMetTxt.Text
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    FileDialog.sInitDir = MdOMetTxt.Text
    FileDialog.sDlgTitle = "Open"
    sOpen = ShowOpen(Me.hWnd)
    
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        MdOMetTxt = FileDialog.sFile
    End If


Exit Sub
MdoFin_Err:
    Select Case Err
        Case 32755:
            Exit Sub     'User Cancel
        Case Else
            If bDebug Then Handle_Err Err, "MdoFin-MdiOpt"
    End Select
End Sub

Private Sub MdOFndCmd_Click()
On Local Error GoTo MdoFin_Err
Dim RetStr As String
If bDebug Then Handle_Err 0, "MdOFndCmd_Click-MdiOpt"

MssngFrm.MssDatTxt.Text = objMdi.info.Datapath

MssngFrm.Show 1

'If Not ObjNew.info.NewUser Then RetStr = Check_Data_Exist(dataPath)

If RetStr = "Canceled" Then Exit Sub
MsgBox "Review modification at this point"
'MdoDatTxt = Datapath

Exit Sub
MdoFin_Err:
    Select Case Err
        Case 32755:
            Exit Sub     'User Cancel
        Case Else
            If bDebug Then Handle_Err Err, "MdoFin-MdiOpt"
    End Select
End Sub

Private Sub MdoMetOpt_Click()
Changed = -1

' ------------------------------------
' Disable or enable Metafile controls
' based on show meta file option...
' ------------------------------------
MdOFinCmd.Enabled = IIf(MdoMetOpt.Value = 1, True, False)
MdOMetTxt.Enabled = IIf(MdoMetOpt.Value = 1, True, False)
lblMetafile.Enabled = IIf(MdoMetOpt.Value = 1, True, False)

End Sub

Private Sub MdOMetTxt_Change()
Changed = -1
End Sub

Private Sub MdOMetTxt_GotFocus()

With MdOMetTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub MdONamTxt_Change()
    Changed = -1
End Sub

Private Sub MdONamTxt_GotFocus()

With MdONamTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub MdONamTxt_KeyPress(KeyAscii As Integer)
' ------------------------
' Allows only 15 character
' length user name.
' ------------------------
If Len(MdONamTxt.Text) >= 15 Then
    Beep
    MsgBox "Maximun user name exceeded.", vbOKOnly, LoadResString(gcTourVersion)
    KeyAscii = 0
End If
End Sub



Sub MdONewmnu_Click()
' ----------------------------------------
' Check for change and prompt user to save
' ----------------------------------------

If bDebug Then Handle_Err 0, "MdoNewmnu-MdiOpt"
On Local Error GoTo NewErr

If Changed Then
    If vbYes = MsgBox("Save changes first?", vbYesNo + vbQuestion, "Save changes") Then UpdateTourUser_UserTbl
End If
ClearControls

bAddUser = True  ' Flags new menu choosen...

Exit Sub
NewErr:
    If bDebug Then Handle_Err Err, "MdONewmnu-MdiOpt"
    Resume Next
End Sub

Private Sub MdOPasChk_Click()
Changed = -1
MdOChaCmd.Enabled = CBool(MdOPasChk.Value)
End Sub

Private Sub MdOPasTxt_Change()
Changed = -1
End Sub


Private Sub MdOSavCmd_Click()
On Local Error GoTo MdOSavCmd_Err
' Declare local variables
Dim sngl As String, RetStr As String

If bDebug Then Handle_Err 0, "MdoSavCmd_Click-MdiOpt"

MdOSavCmd.Enabled = False
Screen.MousePointer = vbHourglass

' Move information to OBJMDI object
UpdateExistingUser

If objMdi.info.NewUser Then
    objMdi.AddUser
Else
    objMdi.SaveUserSettings
End If


bAddUser = False   ' Set new menu clicked flag to false
' Check if new background picture should be
' loaded or removed....

MDI.RefreshPicture  ' Information is already save
                    ' just prompt MDI window to check if image
                    ' should be updated.

MdOSavCmd.Enabled = True
Screen.MousePointer = vbDefault
Changed = 0
Unload MdiOpt
If objMdi.info.WelcomeWizard Then
    objMdi.info.WelcomeWizard = False
    frmWizard.Show
End If
On Local Error GoTo 0
Exit Sub

MdOSavCmd_Err:
    If bDebug Then Handle_Err Err, "MdOSavCmd_Err-MdiOpt"
    Resume Next
End Sub


Sub MdOSavmnu_Click()
On Local Error GoTo MdOSav_Err
' -----------------------
' Update TourUser/UserTbl
' -----------------------
If Changed Then
    MdOSavCmd_Click ' UpdateTourUser_UserTbl
Else
    Unload MdiOpt
End If

Exit Sub
MdOSav_Err:
    If bDebug Then Handle_Err Err, "MdOSavmnu-MdiOpt"
    Resume Next
End Sub


Private Function CreateNewUser() As Boolean
On Local Error GoTo CreateUser_Err
' Declare local variables
Dim SQL As String, ProgressType As String, MsgStr As String
Dim RetInt As Integer, DPath As String, NotResolved As Boolean
Dim RetStr As String, lIdNum As Long
Dim lRt As Long

' Initialize variable
CreateNewUser = False   'False means do not load new user, true means load new user...

' ------------------
' Define SearcherDB
' ------------------
SQL = "Select * From " & gcUserTour_UserTbl
ObjTour.RstSQL iSearcherDB, SQL

' ------------------------
' Try and find exist user
' with same name
' ------------------------
SQL = "Name = '" & MdiOpt.MdONamTxt.Text & "'"
ObjTour.DBFindFirst SQL, iSearcherDB
            
' ----------------------------------------------------------------------------
' Determine if it is save to create new user or does user name already exist?
' ----------------------------------------------------------------------------
If ObjTour.NoMatch(iSearcherDB) Then
    
    ObjTour.AddNew iSearcherDB  ' Add new record
    User_Scr_to_Data            ' Update Dialog fields database
    ObjTour.Update iSearcherDB  ' Update record

    ' --------------------------
    ' Define newly created user
    ' --------------------------
    SQL = "Name = '" & MdONamTxt.Text & "'"
    ObjTour.DBFindFirst SQL, iSearcherDB
    
    lRt = ObjTour.DBGetField("ID", iSearcherDB)
    Call CreateDailyOpt(lRt)
    Call CreateContactOpt(lRt)
Else
    MsgBox LoadResString(1100) & vbCrLf & LoadResString(1101), vbOKOnly, LoadResString(gcTourVersion)
    CreateNewUser = False
    Exit Function
End If
    
' Prompt User
MsgStr = "Do you wish to load new user now?"
RetInt = MsgBox(MsgStr, vbYesNo + vbQuestion, LoadResString(gcTourVersion))

If RetInt = vbYes Then
    'Include all obj settings
    objMdi.info.Name = MdONamTxt.Text
    objMdi.info.ID = lRt
    MsgBox "Review change to following line"
    objMdi.info.Datapath = objMdi.info.Datapath
    CreateNewUser = True
End If

On Local Error GoTo 0
Exit Function

CreateUser_Err:
If bDebug Then Handle_Err Err, "CreateUser-MdiOpt"
Resume Next

End Function
Private Function UpdateExistingUser() As Boolean
On Local Error GoTo UpdateUser_Err

Dim lBitValue       As Long
    
' ---------------------------------
' Update MDI Object with in values
' ---------------------------------
objMdi.info.Name = Trim$(MdONamTxt)

objMdi.info.Security = IIf(MdOPasChk.Value = 1, True, False)
If objMdi.info.Security Then

    Call objMdi.info.UserOptions.SetBool(True, BitFlags.User_Security)
Else
    Call objMdi.info.UserOptions.SetBool(False, BitFlags.User_Security)
End If

' -----------------
' Update Load Type
' -----------------
' 0 = nothing
' 1 = Daily
' 2 = Calendar
' 3 = Contact
' 4 = Conconi
' 5 = Contacts
objMdi.info.Load = cboMDOStartup.ListIndex

'With objMdi.info.UserOptions
Select Case objMdi.info.Load
        Case 0: 'All three bit should be set to 0
            ' Load nothing
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load1)
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load2)
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load3)
            ' Load Daily
        Case 1:
            Call objMdi.info.UserOptions.SetValue(1, BitFlags.User_Load1)
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load2)
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load3)
        
        Case 2: 'Calendar
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load1)
            Call objMdi.info.UserOptions.SetValue(1, BitFlags.User_Load2)
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load3)
            
        Case 3:
            Call objMdi.info.UserOptions.SetValue(1, BitFlags.User_Load1)
            Call objMdi.info.UserOptions.SetValue(1, BitFlags.User_Load2)
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load3)
                
        Case 4:
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load1)
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load2)
            Call objMdi.info.UserOptions.SetValue(1, BitFlags.User_Load3)
            
        Case 5:
            Call objMdi.info.UserOptions.SetValue(1, BitFlags.User_Load1)
            Call objMdi.info.UserOptions.SetValue(0, BitFlags.User_Load2)
            Call objMdi.info.UserOptions.SetValue(1, BitFlags.User_Load3)
        
        Case Else
            Handle_Err "Case does not handle " & CStr(objMdi.info.Load), "UpdateExistingUser-MdiOpt"
    End Select
    
    
objMdi.info.Password = PWD
objMdi.info.MetaFile = IIf(IsNull(MdOMetTxt), " ", MdOMetTxt)

objMdi.info.ShowMeta = CBool(MdoMetOpt.Value)
If objMdi.info.ShowMeta Then
    Call objMdi.info.UserOptions.SetBool(True, BitFlags.User_ShowMeta)
Else
    Call objMdi.info.UserOptions.SetBool(True, BitFlags.User_ShowMeta)
End If

UpdateExistingUser = True

On Local Error GoTo 0
Exit Function

UpdateUser_Err:
If bDebug Then Handle_Err Err, "UpdateExistingUser-MdiOpt"
Resume Next
End Function

'---------------------------------------------------------------------------------------
' PROCEDURE : PWD
' DATE      : 7/3/04 22:29
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Get PWD() As String

On Local Error GoTo PWD_Error
'Declare local variables

    PWD = m_sPWD

On Error GoTo 0
Exit Property

PWD_Error:
    If bDebug Then Handle_Err Err, "PWD-MdiOpt"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : PWD
' DATE      : 7/3/04 22:29
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Private Property Let PWD(ByVal sPWD As String)

On Local Error GoTo PWD_Error
'Declare local variables

    m_sPWD = sPWD

On Error GoTo 0
Exit Property

PWD_Error:
    If bDebug Then Handle_Err Err, "PWD-MdiOpt"
    Resume Next


End Property
