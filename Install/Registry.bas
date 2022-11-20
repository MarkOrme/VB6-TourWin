Attribute VB_Name = "basRegistry"
 ' Disclaimer of Warranty:

' This software and the accompanying files are provided "as is"
' and without warranties as to performance of the software and
' the accompanying files, or any other warranties whether expressed
' or implied.  No warranty of fitness for a particular purpose
' is offered.
'
' You MAY NOT sell this software or it's source code.
' You MAY use this code in any way you find useful.

'Only 4 of the functions in this module are public.  These are the
'CreateRegKey, DeleteRegKey, GetRegStringValue, and WriteRegStringValue
'functions.  The 3 remaining user-defined functions and the Registry
'API functions are private.  There should be no need to have to call
'these from outside this module unless you wanted to re-write it
'and do things differently from what I have written, which you
'are perfectly welcome to do.  As is, this module only supports
'querying string values; however, the constants needed to query other
'types of values are included.  It should not be too difficult to add
'support for these other types.   This module is NOT intended to be used
'for an editor-type program.  All Registry API functions (except
'less-functional counter-parts), constants, etc., have been declared
'for completeness, but they are not all used.

Option Explicit

'Required for RegEnumKey and RegQueryInfoKey
Private Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type

Private Declare Function RegOpenKeyEx& Lib "advapi32.dll" _
Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpszSubKey$, _
dwOptions&, ByVal samDesired&, lpHKey&)

Private Declare Function RegCreateKey& Lib "advapi32" Alias _
"RegCreateKeyA" (ByVal hKey&, ByVal lpszSubKey$, phkResult&)

Private Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias _
"RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, _
ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, _
lpSecurityAttributes&, phkResult&, lpdwDisposition&)

Private Declare Function RegDeleteKey& Lib "advapi32" Alias _
"RegDeleteKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String)

Private Declare Function RegCloseKey& Lib "advapi32.dll" _
(ByVal hKey&)

Private Declare Function RegQueryValueEx& Lib "advapi32.dll" _
Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpszValueName$, _
ByVal lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)

Private Declare Function RegSetValueEx& Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal dwRes&, _
ByVal dwType&, lpDataBuff As Any, ByVal nSize&)

Private Declare Function RegConnectRegistry& Lib "advapi32.dll" _
(ByVal lpMachineName$, ByVal hKey&, phkResult&)

Private Declare Function RegFlushKey& Lib "advapi32.dll" (ByVal hKey&)

Private Declare Function RegEnumKeyEx& Lib "advapi32.dll" Alias _
"RegEnumKeyExA" (ByVal hKey&, ByVal dwIndex&, ByVal lpName$, _
lpcbName&, ByVal lpReserved&, ByVal lpClass$, lpcbClass&, _
lpftLastWriteTime As FILETIME)

Private Declare Function RegEnumValue& Lib "advapi32.dll" Alias _
"RegEnumValueA" (ByVal hKey&, ByVal dwIndex&, ByVal lpName$, _
lpcbName&, ByVal lpReserved&, lpdwType&, lpValue As Any, lpcbValue&)

Private Declare Function RegQueryInfoKey& Lib "advapi32.dll" Alias _
"RegQueryInfoKeyA" (ByVal hKey&, ByVal lpClass$, lpcbClass&, _
ByVal lpReserved&, lpcSubKeys&, lpcbMaxSubKeyLen&, lpcbMaxClassLen&, _
lpcValues&, lpcbMaxValueNameLen&, lpcbMaxValueLen&, _
lpcbSecurityDescriptor&, lpftLastWriteTime As FILETIME)

'Error codes returned from Registry functions.
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

'Registry value types
Const REG_NONE = 0&                        ' No value type
Const REG_SZ = 1&                          ' Unicode nul terminated string
Const REG_EXPAND_SZ = 2&                   ' Unicode nul terminated string
                                           ' (with environment variable references)
Const REG_BINARY = 3&                      ' Free form binary
Const REG_DWORD = 4&                       ' 32-bit number
Const REG_DWORD_LITTLE_ENDIAN = 4&         ' 32-bit number (same as REG_DWORD)
Const REG_DWORD_BIG_ENDIAN = 5&            ' 32-bit number
Const REG_LINK = 6&                        ' Symbolic Link (unicode)
Const REG_MULTI_SZ = 7&                    ' Multiple Unicode strings
Const REG_RESOURCE_LIST = 8&               ' Resource list in the resource map
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&    ' Resource list in the hardware description
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

'Registry Read/Write permissions:
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or _
KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

'Set this flag to True if you do not want to display a message
'if a Registry error occurs.  There are valid reasons for doing this.
'For example, querying a value for a subkey that doesn't exist
'in the Registry.  In this case, you will get a 'Bad Key Name' error,
'but you may not want this message displayed.  You should set this
'flag immediately before the call to any of the Public Registry
'functions in this module.  The functions will automatically reset
'it back to False regardless of whether it was successful or not.
Public gbSkipRegErrMsg As Boolean

Public Const REG_ERROR = "REGISTRY_ERROR"

Public Function DeleteRegKey(sKeyName As String) As Boolean

'This functions deletes the passed key.  It returns True if the
'key is deleted and False if an error occurred or the key didn't
'exist to begin with.  No error message is displayed (unless there
'is a parsing error) because there is really no point in it.  The only
'likely error would be because the key does not exist and we wanted
'to delete it anyway.  So, so what if it doesn't exist.

Dim hKey As Long, lRtn As Long
Dim lMainKeyHandle As Long

DeleteRegKey = False

Call ParseKey(sKeyName, lMainKeyHandle)

If lMainKeyHandle Then
    'Open the key
    lRtn = RegOpenKeyEx(lMainKeyHandle, sKeyName, 0&, KEY_WRITE, hKey)
    If lRtn = ERROR_SUCCESS Then
        lRtn = RegDeleteKey(hKey, sKeyName)
        lRtn = RegCloseKey(hKey)
        DeleteRegKey = True
    End If
End If

'Even though error messages are not displayed within this function,
'we should be thorough and reset the flag anyway, just in case it
'was set prior to calling the function.
gbSkipRegErrMsg = False

End Function

Private Function GetMainKeyHandle(sMainKeyName As String) As Long

'Returns the handle of the main key (the constants below)

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case sMainKeyName
    Case "HKEY_CLASSES_ROOT"
        GetMainKeyHandle = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetMainKeyHandle = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetMainKeyHandle = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        GetMainKeyHandle = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
        GetMainKeyHandle = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
        GetMainKeyHandle = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Private Function GetRegError(lErrorCode As Long) As String
    
'This function returns the error string associated with error
'codes returned by the Registry API functions

Select Case lErrorCode
    Case 1009, 1015
        'As Doomsday* would say "We're in trouble now!"
        GetRegError = "The Registry Database is corrupt!"
    Case 2, 1010
        GetRegError = "Bad Key Name"
    Case 1011
        GetRegError = "Can't Open Key"
    Case 4, 1012
        GetRegError = "Can't Read Key"
    Case 5
        GetRegError = "Access to this key is denied"
    Case 1013
        GetRegError = "Can't Write Key"
    Case 8, 14
        GetRegError = "Out of memory"
    Case 87
        GetRegError = "Invalid Parameter"
    Case 234
        GetRegError = "There is more data than the buffer has been allocated to hold."
    Case Else
        GetRegError = "Undefined Error Code:  " & str$(lErrorCode)
End Select

'* - Haven't you ever played Wing Commander?  :)
End Function

Private Sub ParseKey(sKeyName As String, lKeyHandle As Long)
    
'This sub parses the passed keyname, separates the main HKEY from
'its subkeys, and returns the main key handle and subkey.
'Pass the full path to the key in this format:
'HKEY_CURRENT_USER\Control Panel\Desktop

Dim nBackSlash As Integer

nBackSlash = InStr(sKeyName, "\")

If Left(sKeyName, 5) <> "HKEY_" Or Right(sKeyName, 1) = "\" Then
    MsgBox "Incorrect Format:" & vbCrLf & vbCrLf & sKeyName
    Exit Sub
End If

If nBackSlash = 0 Then
    'Only the main key was specified; get its handle
    lKeyHandle = GetMainKeyHandle(sKeyName)
    'No other keyname to return
    sKeyName = ""
Else
    'Get the handle to the main key
    lKeyHandle = GetMainKeyHandle(Left(sKeyName, nBackSlash - 1))
    'Strip the main key string and return the rest
    sKeyName = Right(sKeyName, Len(sKeyName) - nBackSlash)
End If

'Make sure the handle is valid
If lKeyHandle < &H80000000 Or lKeyHandle > &H80000006 Then
    MsgBox "Not a valid main key handle"
End If

End Sub
Public Function GetRegStringValue(sSubKey As String, sEntry As String) As String

'Returns the string value for the passed key.

'Because it's possible virtually any string value could be returned,
'we cannot use a "0" or a null string to indicate that an error
'occurred, so return the string defined by the constant REG_ERROR.
'Even so, it's still possible that for some obscure reason, an
'application even wrote that string as an entry, but it's not likely.

Dim hKey As Long, lMainKeyHandle As Long
Dim lRtn As Long, sBuffer As String
Dim lBufferSize As Long, lType As Long
Dim sErrMsg As String


lType = REG_SZ
'lType = REG_DWORD
GetRegStringValue = REG_ERROR

Call ParseKey(sSubKey, lMainKeyHandle)

If lMainKeyHandle Then
    'Open the key
    lRtn = RegOpenKeyEx(lMainKeyHandle, sSubKey, 0&, KEY_READ, hKey)
    If lRtn = ERROR_SUCCESS Then
        'Query for the value
        sBuffer = Space(255)
        lBufferSize = Len(sBuffer)
        lRtn = RegQueryValueEx(hKey, sEntry, 0&, lType, sBuffer, lBufferSize)
        If lRtn = ERROR_SUCCESS Then
            lRtn = RegCloseKey(hKey)
            'Remove spaces and eliminate the terminating character
            If lBufferSize > 1 Then
                GetRegStringValue = Left(sBuffer, lBufferSize - 1)
            End If
        Else
            If gbSkipRegErrMsg = False Then
                sErrMsg = GetRegError(lRtn)
                MsgBox sErrMsg, vbCritical, "Registry Query Error"
            End If
        End If
    Else
        If Not gbSkipRegErrMsg Then
            sErrMsg = GetRegError(lRtn)
            MsgBox sErrMsg, vbCritical, "Registry Open Error"
        End If
    End If
End If

gbSkipRegErrMsg = False

End Function

Public Function WriteRegStringValue(sSubKey As String, sEntry As String, sValue As String) As Boolean

'Returns True if successful; otherwise False

Dim hKey As Long, lMainKeyHandle As Long
Dim lRtn As Long, lDataSize As Long
Dim lType As Long, sErrMsg As String

WriteRegStringValue = False

lType = REG_SZ

Call ParseKey(sSubKey, lMainKeyHandle)

If lMainKeyHandle Then
    'Open the key
    lRtn = RegOpenKeyEx(lMainKeyHandle, sSubKey, 0&, KEY_WRITE, hKey)
    If lRtn = ERROR_SUCCESS Then
        'Write the value
        lDataSize = Len(sValue)
        lRtn = RegSetValueEx(hKey, sEntry, 0&, lType, ByVal sValue, lDataSize)
        If lRtn = ERROR_SUCCESS Then
            WriteRegStringValue = True
        Else
            If Not gbSkipRegErrMsg Then
                sErrMsg = GetRegError(lRtn)
                MsgBox sErrMsg, vbCritical, "Registry Write Error"
            End If
        End If
        lRtn = RegCloseKey(hKey)
    Else
        If Not gbSkipRegErrMsg Then
            sErrMsg = GetRegError(lRtn)
            MsgBox sErrMsg, vbCritical, "Registry Open Error"
        End If
    End If
End If

gbSkipRegErrMsg = False

End Function

Public Function CreateRegKey(sSubKey As String) As Boolean

'This function creates the passed key. It returns True if successful
'or False if an error occurred.

'Creating a key also opens it.  The rest of the functions are
'written with the assumption that the key is closed, so close it
'after creating it.  While you could change this to return the
'handle of the created key, you would also have to modify the
'rest of the functions to pass them the handle.  Besides, it's
'much safer to NOT have the keys open any longer than absolutely
'necesssary.  If keys are open and Windows crashes, there's a power
'outage, etc., the entire Registry could get corrupted.

Dim lMainKeyHandle As Long, hKey As Long
Dim sErrMsg As String, lRtn As Long

CreateRegKey = False

Call ParseKey(sSubKey, lMainKeyHandle)

'Create the subkey if we have a handle to the main key
If lMainKeyHandle Then
    lRtn = RegCreateKey(lMainKeyHandle, sSubKey, hKey)
    If lRtn = ERROR_SUCCESS Then
        'key successfully created
        lRtn = RegCloseKey(hKey)
        CreateRegKey = True
    Else
        If Not gbSkipRegErrMsg Then
            sErrMsg = GetRegError(lRtn)
            MsgBox sErrMsg, vbCritical
        End If
    End If
End If

gbSkipRegErrMsg = False

End Function
