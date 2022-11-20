Attribute VB_Name = "Module3"
Option Explicit
Function TextValidate(KeyAscii As Integer, TxtType, Cntrl As Control) As Integer
'
' Allow carrot movement keys
'
Select Case KeyAscii
        Case 8:
        Cntrl.SelLength = 0
        TextValidate = KeyAscii
        Exit Function
End Select

Select Case TxtType
        Case "Date":        ' Allow 0 - 9 and "-" or "/" Char
            If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 45 Or KeyAscii = 47 Then
                'Allow
                Exit Function
            Else
                KeyAscii = 0
            End If
        Case "Int Only":
            If KeyAscii <= 47 Or KeyAscii > 57 Then KeyAscii = 0
        Case "Time Only":
            If KeyAscii <= 47 Or KeyAscii > 58 Then KeyAscii = 0
        Case "Real Only":
            If KeyAscii <= 45 Or KeyAscii > 57 And KeyAscii <> 47 Then KeyAscii = 0
End Select
TextValidate = KeyAscii
End Function

Sub CentreForm(NameFrm As Form, TopHeight As Integer)
On Local Error GoTo Centre_Err
' ----------------------
' Define Left coordinate
' ----------------------
If "PassFrm" = NameFrm.Name Then
    NameFrm.Left = (Screen.Width / 2) - (NameFrm.Width / 2)
Else
    NameFrm.Left = (MDI.Width / 2) - (NameFrm.Width / 2)
End If
' -------------------------------
' If Left is negative then make 0
' -------------------------------
If 0 > Val(NameFrm.Left) Then NameFrm.Left = 0

' ----------------------
' Define Top coordinate
' ----------------------

If TopHeight = -1 Then
    NameFrm.Top = 0
Else
    If "PassFrm" = NameFrm.Name Then
        NameFrm.Top = (Screen.Height / 2) * 0.85 - (NameFrm.Height / 2)
    Else
        NameFrm.Top = (MDI.Height / 2) - (NameFrm.Height / 2)
    End If
End If

' -------------------------------
' If Left is negative then make 0
' -------------------------------
If 0 > Val(NameFrm.Top) Then NameFrm.Top = 0

Exit Sub
Centre_Err:
        If bDebug Then Handle_Err Err, "CentreForm-Generic"
        Resume Next
End Sub

Function GetDateRange() As Boolean
' ---------------------------
' This function calls DateFrm
' and return True is user
' entered valid date
' Make Sure Project has DateFrm.Frm loaded
' ----------------------------

On Local Error GoTo GetDateRange_Err
DateFrm.Show 1

Exit Function
GetDateRange_Err:
    If bDebug Then Handle_Err Err, "GetDateRange-Module2"
    Resume Next
End Function
Function MakeDirectory(sPathToCreate As String) As Boolean
'
' Separate the sPathToCreate string by '\'
'
On Local Error Resume Next
Dim iBackSlash As Integer
Dim iStart As Integer
Dim sPath As String
iStart = 1
iBackSlash = InStr(iStart, sPathToCreate, "\", vbTextCompare)

Do
    ChDir Mid$(sPathToCreate, 1, iBackSlash - 1)
    If 0 <> Err Then
        Err.Clear
        Debug.Print Mid$(sPathToCreate, 0, iBackSlash - 1)
        sPath = Mid$(sPathToCreate, 1, iBackSlash - 1)
        MkDir sPath
    End If
    iStart = iBackSlash + 1
    iBackSlash = InStr(iStart, sPathToCreate, "\", vbTextCompare)
Loop While iBackSlash <> 0
MkDir sPathToCreate
End Function

Function GetPostionValue(ByVal Sched As String, ByRef iCurrentPos As Integer, ByVal sDeliminator As String, ByRef sPosition As String) As Boolean
' ===============================================
' GetPostionValue - Find the first occurance of sDeliminator
' and return value found within
' ===============================================
    Dim iStartingPoint As Integer
    On Local Error GoTo GetPosition_Error
    iStartingPoint = iCurrentPos
    If iStartingPoint <= 0 Then iStartingPoint = 1
    iCurrentPos = InStr(iStartingPoint, Sched, sDeliminator, vbTextCompare)
    If iCurrentPos > 0 Then
        sPosition = Mid$(Sched, iStartingPoint, (iCurrentPos - iStartingPoint))
        GetPostionValue = True
        iCurrentPos = iCurrentPos + 1
    End If
Exit Function
GetPosition_Error:
    MsgBox Err.Description
End Function
