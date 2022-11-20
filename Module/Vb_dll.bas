Attribute VB_Name = "Dll_Module"
'API Declarations
Global Const HELP_CONTEXT = &H1
Global Const HELP_CONTENTS = &H3
Global Const HELP_SETCONTENTS = &H5
Global Const HELP_PARTIALKEY = &H105
Global Const HELP_KEY = &H101
Global Const HELP_COMMAND = &H102
Global Const HELP_HELPONHELP = &H4
Global Const HELP_CONTEXTPOPUP = &H8
Global Const HELP_FORCEFILE = &H9
Global Const WM_USER = &H400
Global Const LB_FINDSTRING = (WM_USER + 16)

' Constants for GetLocaleInfo function
Public Const LOCALE_SDATE = &H1D            '  date separator
Public Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Public Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Public Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Public Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Public Const LOCALE_SSHORTDATE = &H1F       '  short date format string
Public Const LOCALE_SLONGDATE = &H20        '  long date format string
Public Const LOCALE_USER_DEFAULT = &H400

#If Win32 Then
Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpString$, ByVal lplFileName$)
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Declare Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function PESetSelectionFormula Lib "CRPE.DLL" (ByVal Printjob As Integer, ByVal FormulaString As String) As Integer
Declare Function PEGetSelectionFormula Lib "CRPE.DLL" (ByVal Printjob As Integer, textHandle As Integer, textLenght As Integer) As Integer
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Declare Function GetPrivateProfileString% Lib "kernel.dll" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString$, ByVal nSize%, ByVal lplFileName$)
Declare Function WritePrivateProfileString% Lib "kernel.dll" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName$)
Declare Function GetPrivateProfileInt% Lib "kernel.dll" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName$)
Declare Function WinHelp Lib "User.Exe" (ByVal hwnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer
Declare Sub SetCaretPos Lib "User.Exe" (ByVal x As Integer, ByVal Y As Integer)
Declare Function PostMessage Lib "User.Exe" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
Declare Function SendMessage Lib "User.Exe" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function PESetSelectionFormula Lib "CRPE.DLL" (ByVal Printjob As Integer, ByVal FormulaString As String) As Integer
Declare Function PEGetSelectionFormula Lib "CRPE.DLL" (ByVal Printjob As Integer, textHandle As Integer, textLenght As Integer) As Integer
Declare Function GetLocaleInfo Lib "kernel.dll" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#End If

Type POINTAPI
    x As Integer
    Y As Integer
End Type
Dim lpPoint As POINTAPI
Declare Function GetFocus Lib "User.Exe" () As Integer
Declare Sub GetCaretPos Lib "User.Exe" (lpPoint As POINTAPI)

Public Const LB_SETTABSTOPS = &H192

Global PeakNames(0 To 9) As String
Global IniPath As String                ' setup by password for use by multi
Global UserPath As String               ' user, ie different directories.
                                        
Function GetIntFromINI(AppName$, KeyName$, FileName$) As Integer
#If Win32 Then
    Dim RetVal As Long
#Else
    Dim RetVal As Integer
#End If

GetIntFromINI = Int(GetPrivateProfileInt(ByVal AppName$, ByVal KeyName$, RetInt, ByVal FileName$))
End Function

Function GetStrFromINI(AppName$, KeyName$, FileName$) As String
On Local Error GoTo GetStr_Err
#If Win32 Then
    Dim RetVal As Long
#Else
    Dim RetVal As Integer
#End If

Dim RetStr As String
RetStr = String(255, Chr(0))
If bDebug Then Handle_Err 0, "GetStrFromInI-Module1"
RetVal = GetPrivateProfileString(AppName, ByVal KeyName, "No Return", RetStr, Len(RetStr), FileName)
'Finds first Null character elimantes.
GetStrFromINI = Mid$(RetStr, 1, InStr(1, RetStr, Chr$(0), 1) - 1)
Exit Function
GetStr_Err:
    If bDebug Then Handle_Err Err, "GetStrFromInI-Module1"
    Resume Next
End Function

Function WriteStrToINI(AppName$, KeyName$, lpString$, FileName$) As String
#If Win32 Then
    Dim RetVal As Long
#Else
    Dim RetVal As Integer
#End If
RetVal = WritePrivateProfileString(AppName, KeyName, lpString, FileName)
WriteStrToINI = RetVal
End Function

Function fGetSel&()
Dim location As Long, ending As Long, starting As Long
Const EM_GETSEL = &H400 + 0
location = SendMessage(GetFocus(), EM_GETSEL, 0, 0&)
ending = location \ &H10000
starting = location And &H7FFF
fGetSel& = starting
End Function

Function fSetSel&(Pos&)
Dim location As Long
Const EM_SETSEL = &H400 + 1
location = Pos& * 2 ^ 16 + Pos&
Pos = SendMessage(GetFocus(), EM_SETSEL, 0, location)
End Function




