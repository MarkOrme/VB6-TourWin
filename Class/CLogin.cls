VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : CLogin
' DateTime  : 6/30/04 16:48
' Author    : Mark Ormesher
' Purpose   : Interface for displaying and validating user
'             login attempt
'---------------------------------------------------------------------------------------
Option Explicit

'Declare local variables
Const INVALIDNAME = 0
Const INVALIDPWD = 1
Const INVALIDDATAPATH = 2

Const OK = 0
Const BADNAME = 1
Const BADPASSWORD = 2
Const FOUNDANDOPENED = 0
Const CREATEDANDOPENED = 1
Const NOTFOUNDNORCREATED = 2

Private WithEvents m_LoginForm      As PassFrm
Attribute m_LoginForm.VB_VarHelpID = -1
Private m_lLoginSuccessful          As Long
Private m_sUserName                 As String
Private m_sPassword                 As String
Private m_sDatapath                 As String

'---------------------------------------------------------------------------------------
' PROCEDURE : UserName
' DATE      : 6/30/04 17:04
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get UserName() As String

On Local Error GoTo UserName_Error
'Declare local variables

    UserName = m_sUserName

On Error GoTo 0
Exit Property

UserName_Error:
    If bDebug Then Handle_Err Err, "UserName-CLogin"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : UserName
' DATE      : 6/30/04 17:04
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let UserName(ByVal sUserName As String)

On Local Error GoTo UserName_Error
'Declare local variables

    m_sUserName = sUserName

On Error GoTo 0
Exit Property

UserName_Error:
    If bDebug Then Handle_Err Err, "UserName-CLogin"
    Resume Next


End Property


'---------------------------------------------------------------------------------------
' PROCEDURE : Display
' DATE      : 6/30/04 16:50
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub Display()

On Local Error GoTo Display_Error
'Declare local variables
Set m_LoginForm = New PassFrm

' Set initial values
If Len(Me.UserName) > 0 Then m_LoginForm.PasNamTxt.Text = Me.UserName
If Len(Me.Password) > 0 Then m_LoginForm.PasPasTxt.Text = Me.Password
If Len(Me.Datapath) > 0 Then
    GetDatabaseList m_LoginForm.cboPasdb, Me.Datapath
Else
    GetDatabaseList m_LoginForm.cboPasdb
End If

' Set loginSuccessful for cancel button...
LoginSuccessful = NOTFOUNDNORCREATED

m_LoginForm.Show vbModal


Set m_LoginForm = Nothing

On Error GoTo 0
Exit Sub

Display_Error:
    If bDebug Then Handle_Err Err, "Display-CLogin"
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : LoginSuccessful
' DATE      : 6/30/04 16:51
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get LoginSuccessful() As Long

On Local Error GoTo LoginSuccessful_Error
'Declare local variables

    LoginSuccessful = m_lLoginSuccessful

On Error GoTo 0
Exit Property

LoginSuccessful_Error:
    If bDebug Then Handle_Err Err, "LoginSuccessful-CLogin"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : LoginSuccessful
' DATE      : 6/30/04 16:51
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let LoginSuccessful(ByVal lLoginSuccessful As Long)

On Local Error GoTo LoginSuccessful_Error
'Declare local variables

    m_lLoginSuccessful = lLoginSuccessful

On Error GoTo 0
Exit Property

LoginSuccessful_Error:
    If bDebug Then Handle_Err Err, "LoginSuccessful-CLogin"
    Resume Next


End Property

Private Sub m_LoginForm_LoginSubmitted(ByVal sName As String, ByVal sPWD As String, ByVal sDatapath As String)

On Local Error GoTo Login_Error
Dim lRt         As Long

Me.UserName = sName
Me.Password = sPWD
Me.Datapath = sDatapath


' Validate ...
' First try and open the database

Select Case cTour_DB.DBOpen(sDatapath)
       Case FOUNDANDOPENED:  'Found and opened successfully

             lRt = ValidateLogin

                Select Case lRt

                       Case OK:
                            Me.LoginSuccessful = FOUNDANDOPENED
                            Unload m_LoginForm

                       Case BADNAME, BADPASSWORD:
                            m_LoginForm.BadLogin lRt
                            Exit Sub
                                
                End Select


       Case CREATEDANDOPENED:  'Add new user to new DB

                Me.LoginSuccessful = CREATEDANDOPENED
                Unload m_LoginForm

       Case NOTFOUNDNORCREATED:
                Me.LoginSuccessful = NOTFOUNDNORCREATED

End Select

DoEvents
On Error GoTo 0
Exit Sub

Login_Error:
If bDebug Then Handle_Err Err, "Login_Error-CLogin"
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : Password
' DATE      : 6/30/04 17:05
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Password() As String

On Local Error GoTo Password_Error
'Declare local variables

    Password = m_sPassword

On Error GoTo 0
Exit Property

Password_Error:
If bDebug Then Handle_Err Err, "Password-CLogin"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Password
' DATE      : 6/30/04 17:05
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Password(ByVal sPassword As String)

On Local Error GoTo Password_Error
'Declare local variables

    m_sPassword = sPassword

On Error GoTo 0
Exit Property

Password_Error:
    If bDebug Then Handle_Err Err, "Password-CLogin"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Datapath
' DATE      : 6/30/04 17:05
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get Datapath() As String

On Local Error GoTo Datapath_Error
'Declare local variables

    Datapath = m_sDatapath

On Error GoTo 0
Exit Property

Datapath_Error:
    If bDebug Then Handle_Err Err, "Datapath-CLogin"
    Resume Next


End Property

'---------------------------------------------------------------------------------------
' PROCEDURE : Datapath
' DATE      : 6/30/04 17:05
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let Datapath(ByVal sDatapath As String)

On Local Error GoTo Datapath_Error
'Declare local variables

    m_sDatapath = sDatapath

On Error GoTo 0
Exit Property

Datapath_Error:
    If bDebug Then Handle_Err Err, "Datapath-CLogin"
    Resume Next


End Property

Private Sub GetDatabaseList(ByRef oControl As ComboBox, Optional ByVal sSetAsSelected As String)

    ' Read list of db name from registry.
Dim bEnd As Boolean
Dim iStartAt As Integer
Dim iPos As Integer
Dim sDBList As String
Dim iGetCBIndexFromString As Integer

gbSkipRegErrMsg = True
sDBList = GetRegStringValue(LoadResString(gcRegTourKey), LoadResString(gcTourDBase))

If REG_ERROR = sDBList Then sDBList = sSetAsSelected
    
If Right$(sDBList, 1) <> "," Then
     sDBList = sDBList & ","
End If

iStartAt = 1
bEnd = False

Do
    iPos = InStr(iStartAt, sDBList, ",", vbBinaryCompare)
    oControl.AddItem Mid(sDBList, iStartAt, iPos - iStartAt)
    If iPos >= Len(sDBList) Then Exit Do
    iStartAt = iPos + 1
Loop While Not bEnd

If oControl.ListCount > 0 Then
    If "" = sSetAsSelected Then ' Argument is missing
        oControl.ListIndex = 0
    Else
        ' Try and find sSetAsSelected
        iGetCBIndexFromString = FindItemListControl(oControl, sSetAsSelected)
        If -1 <> iGetCBIndexFromString Then
            oControl.ListIndex = iGetCBIndexFromString
        Else
            oControl.ListIndex = 0  ' Make sure something is set a default
        End If
    End If
End If
    
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : Login
' DATE      : 6/26/04 20:14
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function ValidateLogin() As Long

On Local Error GoTo Login_Error

'Declare local variables
Const OK = 0
Const BADNAME = 1
Const BADPASSWORD = 2

Dim Search      As String
Dim DirFound    As String
Dim SQL         As String
Dim lSearcher   As Long
Dim mMDIVar     As cMDIVar

' User is allow three attempts
' before program automatically ends..

Set mMDIVar = New cMDIVar

' ----------------------------
' Load Searcher object and
' Determine if any users exist
' ----------------------------
SQL = "SELECT * FROM " & gcUserTour_UserTbl

ObjTour.RstSQL lSearcher, SQL

ObjTour.DBMoveLast lSearcher

Search = "Name = '" & Trim$(UserName) & "'"
ObjTour.DBMoveFirst lSearcher

ObjTour.DBFindFirst Search, lSearcher

' ------------------------------------
' If no match to user name.
' ------------------------------------
If ObjTour.NoMatch(lSearcher) Then

    ValidateLogin = BADNAME
    ObjTour.FreeHandle lSearcher
    Exit Function
    
End If

' Is security flag on?
mMDIVar.UserOptions.SetValue (ObjTour.DBGetField(gcUserTour_UserTbl_BitField, lSearcher))

If mMDIVar.UserOptions.GetBool(BitFlags.User_Security) Then
    If ObjTour.DBGetField("PassWord", lSearcher) <> Me.Password Then
    
        ValidateLogin = BADPASSWORD
        ObjTour.FreeHandle lSearcher
        Exit Function
    
    End If
End If

ValidateLogin = OK
Set mMDIVar = Nothing
ObjTour.FreeHandle lSearcher

On Error GoTo 0
Exit Function

Login_Error:

End Function
