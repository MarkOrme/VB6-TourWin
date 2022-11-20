VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       CRegistration
' FILENAME:     C:\My Code\vb\Registration\CRegistration.cls
' AUTHOR:       Phil Fresle
' CREATED:      06-Sep-2000
' COPYRIGHT:    Copyright 2000 Frez Systems Limited.
'
' DESCRIPTION:
' This class demonstrates how you can create semi-random license keys for
' your applications and test their validity. You should change the algorithm
' used to calculate validity to ensure uniqueness to yourself.
'
' The 12 character random key gives around 1 billion different combinations but
' a user would only have to try an average of MOD_NUMBER keys to randomly find
' a valid value. This works for me as I am pretty sure to get a unique key
' whilst at least discouraging a random attack. By storing the keys in a
' database along with who they have been issued to I can track key use.
'
' If this class was put into a DLL you would not want to expose the Generate
' method on the default interface or anyone could use it.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
'
' MODIFICATION HISTORY:
' 1.0       06-Sep-2000
'           Phil Fresle
'           Initial Version
'
' 1.1       07-Sep-2000
'           Phil Fresle
'           Allowed for smarter/faster generation of keys and easier
'           customisation through the use of constants
'*******************************************************************************
Option Explicit
Private m_Key                   As String ' Value of data location field of TourSystem user record.
Private m_UsageCount            As Long   ' Value of BitField
Private m_Results               As Long   ' Result of registration


Private Const PART_LENGTH       As Long = 4
Private Const PART1_LENGTH      As Long = 5
Private Const KEY_LENGTH        As Long = PART1_LENGTH + (3 * PART_LENGTH)
Private Const SLEEP_INTERVAL    As Long = 1000

' The value of (9 * (POS_17_MULTIPLIER + POS_16_MULTIPLIER + POS_15_MULTIPLIER))
' must be at least MOD_NUMBER to allow quick generation of keys.
'
' Combinations of 0 to 9 times each of POS_17_MULTIPLIER, POS_16_MULTIPLIER
' and POS_15_MULTIPLIER must allow selection of every number between 1 and
' MOD_NUMBER.
'
' POS_17_MULTIPLIER must be 1, POS_16_MULTIPLIER must be in the range 2-10, and
' POS_15_MULTIPLIER must be greater than POS_16_MULTIPLIER to allow the fast
' calculation as it is currently implemented to work.
'
' As long as the above rules are followed the numbers may be changed to
' implement your own unique checksum calculation
Private Const POS_01_MULTIPLIER As Long = 17
Private Const POS_02_MULTIPLIER As Long = 12
Private Const POS_03_MULTIPLIER As Long = 10
Private Const POS_04_MULTIPLIER As Long = 8
Private Const POS_05_MULTIPLIER As Long = 4
Private Const POS_06_MULTIPLIER As Long = 15
Private Const POS_07_MULTIPLIER As Long = 7
Private Const POS_08_MULTIPLIER As Long = 5
Private Const POS_09_MULTIPLIER As Long = 3
Private Const POS_10_MULTIPLIER As Long = 13
Private Const POS_11_MULTIPLIER As Long = 11
Private Const POS_12_MULTIPLIER As Long = 6
Private Const POS_13_MULTIPLIER As Long = 2
Private Const POS_14_MULTIPLIER As Long = 16
Private Const POS_15_MULTIPLIER As Long = 14
Private Const POS_16_MULTIPLIER As Long = 9
Private Const POS_17_MULTIPLIER As Long = 1

' The MOD_NUMBER cannot be greater than around
' (9 * (SUM(POS_06_MULTIPLIER..POS_17_MULTIPLIER) - 3))
' and should be significantly less than that value (say a quarter)
' 211 was chosen as it is Prime and will give 'random enough'
' numbers, of course a brute force attack will get a valid
' number very quickly
Private Const MOD_NUMBER As Long = 211

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'*******************************************************************************
' PartOK (FUNCTION)
'
' PARAMETERS:
' (In) - sPart - String - Part to check
'
' RETURN VALUE:
' Boolean - True if OK
'
' DESCRIPTION:
' Makes sure part of the key is not easily guessed
'*******************************************************************************
Private Function PartOK(ByVal sPart As String) As Boolean
    sPart = Trim(sPart)
    
    If Len(sPart) <> PART_LENGTH Then
        PartOK = False
    ElseIf Not IsNumeric(sPart) Then
        PartOK = False
    ElseIf sPart = "0000" Or sPart = "1111" Or sPart = "2222" _
        Or sPart = "3333" Or sPart = "4444" Or sPart = "5555" _
        Or sPart = "6666" Or sPart = "7777" Or sPart = "8888" _
        Or sPart = "9999" Or sPart = "1234" Or sPart = "4321" Then
            PartOK = False
    Else
        PartOK = True
    End If
End Function

'*******************************************************************************
' GenerateKey (FUNCTION)
'
' PARAMETERS:
' (In) - sPart1 - String - Application and version identifier
'
' RETURN VALUE:
' String - The new key
'
' DESCRIPTION:
' Generate a new application key. NOTE: sPart1 is used to identify the
' application, and application version. Generates most of the key randomly then
' uses the last three numbers to make sure that the checksum will be a modulus
' of MOD_NUMBER.
'*******************************************************************************
Public Function GenerateKey(ByVal sPart1 As String) As String
    Dim sKey        As String
    Dim sPart2      As String
    Dim sPart3      As String
    Dim sPart4      As String
    Dim lCount      As Long
    Dim lModulus    As Long
    Dim lNextNumber As Long
    
    sPart1 = Trim(sPart1)
    
    If Len(sPart1) <> PART1_LENGTH Then
        Err.Raise vbObjectError, , "Part 1 must be " & PART1_LENGTH & " characters"
    ElseIf Not IsNumeric(sPart1) Then
        Err.Raise vbObjectError, , "Part 1 must be numeric"
    Else
        ' Could be a decimal or negative number, so make sure
        For lCount = 1 To Len(sPart1)
            If Not IsNumeric(Mid(sPart1, lCount, 1)) Then
                Err.Raise vbObjectError, , "Part 1 must be a whole positive numeric number"
            End If
        Next
    End If
    
    ' We want a random key
    Randomize
    
    ' Loop till we get a valid key
    Do
        ' Save time by pre-valiating each part
        Do
            sPart2 = ""
            ' Put together four random numbers
            For lCount = 1 To PART_LENGTH
                sPart2 = sPart2 & CStr(Int(10 * Rnd))
            Next
        Loop Until PartOK(sPart2)
    
        Do
            sPart3 = ""
            ' Put together four random numbers
            For lCount = 1 To PART_LENGTH
                sPart3 = sPart3 & CStr(Int(10 * Rnd))
            Next
        Loop Until PartOK(sPart3)
    
        ' PMF - 07-Sep-2000 15:29
        ' This is where changes were made for version 1.1 so that keys are
        ' more quickly created than simply generating over and over until a valid
        ' key is found
        
        ' Start with one random number
        sPart4 = CStr(Int(10 * Rnd))
        
        ' This part of the code makes sure that the last three numbers when added
        ' to the existing numbers will ensure that the full key is modulus MOD_NUMBER
        lModulus = MOD_NUMBER - _
            (GetCheckSum(sPart1 & sPart2 & sPart3 & sPart4) Mod MOD_NUMBER)
        
        lNextNumber = lModulus \ POS_15_MULTIPLIER
        If lNextNumber > 9 Then
            lNextNumber = 9
        End If
        sPart4 = sPart4 & CStr(lNextNumber)
        lModulus = lModulus - (lNextNumber * POS_15_MULTIPLIER)
        
        lNextNumber = lModulus \ POS_16_MULTIPLIER
        If lNextNumber > 9 Then
            lNextNumber = 9
        End If
        sPart4 = sPart4 & CStr(lNextNumber)
        lModulus = lModulus - (lNextNumber * POS_16_MULTIPLIER)
        
        lNextNumber = lModulus \ POS_17_MULTIPLIER
        sPart4 = sPart4 & CStr(lNextNumber)
        
        ' Put together the full key for checking
        sKey = sPart1 & sPart2 & sPart3 & sPart4
    Loop Until KeyOK(sKey)
    
    GenerateKey = sKey
End Function

'*******************************************************************************
' IsKeyOK (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sKey - String - Key to check
'
' RETURN VALUE:
' Boolean - Whether key was valid
'
' DESCRIPTION:
' Public interface to check a key, makes sure the app sleeps for a while to
' discourage brut force checking
'*******************************************************************************
Public Function IsKeyOK(sKey As String) As Boolean
    ' This sleep interval is optional code. I put it in to discourage brute force
    ' discovery of a valid key by trying different numbers over and over. You may
    ' not require this or decide on a different interval (change the
    ' SLEEP_INTERVAL const).
    
    'GetSystemRegistrationKey
    
    IsKeyOK = KeyOK(sKey)
    
End Function

'*******************************************************************************
' KeyOK (FUNCTION)
'
' PARAMETERS:
' (In) - sKey - String - Key to check
'
' RETURN VALUE:
' Boolean - True if OK
'
' DESCRIPTION:
' Checks if the entire key is OK by a customised sum of the key, the way this
' is calculated should be unique to you, so I suggest you change the algorithm.
'*******************************************************************************
Private Function KeyOK(ByVal sKey As String) As Boolean
    Dim lChecksum   As Long
    Dim sPart2      As String
    Dim sPart3      As String
    Dim sPart4      As String
    Dim lCount      As Long
    
    sKey = Trim(sKey)
    
    ' Make sure key is correct length
    If Len(sKey) <> KEY_LENGTH Then
        KeyOK = False
        Exit Function
    ElseIf Not IsNumeric(sKey) Then
        KeyOK = False
        Exit Function
    Else
        For lCount = 1 To Len(sKey)
            If Not IsNumeric(Mid(sKey, lCount, 1)) Then
                KeyOK = False
                Exit Function
            End If
        Next
    End If
    
    ' Make sure key parts are valid
    sPart2 = Mid(sKey, PART1_LENGTH + 1, PART_LENGTH)
    sPart3 = Mid(sKey, PART1_LENGTH + PART_LENGTH + 1, PART_LENGTH)
    sPart4 = Mid(sKey, PART1_LENGTH + (2 * PART_LENGTH) + 1, PART_LENGTH)
    If Not PartOK(sPart4) Then
        KeyOK = False
        Exit Function
    ElseIf Not PartOK(sPart3) Then
        KeyOK = False
        Exit Function
    ElseIf Not PartOK(sPart2) Then
        KeyOK = False
        Exit Function
    End If
    
    ' Compute checksum
    lChecksum = GetCheckSum(sKey)
            
    ' If the checksum is modulus MOD_NUMBER then it is valid
    If (lChecksum Mod MOD_NUMBER) = 0 Then
        KeyOK = True
    Else
        KeyOK = False
    End If
End Function

'*******************************************************************************
' GetCheckSum (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sKey - String - Key to compute checksum with.
'
' RETURN VALUE:
' Lone - Value of checksum.
'
' DESCRIPTION:
' This function calculates the custom checksum. It may be called with a full
' key from the validation routine, or a partial key from the key generation
' routine.
'
' NOTE: This function was added for version 1.1
'*******************************************************************************
Private Function GetCheckSum(sKey As String) As Long
    Dim lChecksum As Long
    
    lChecksum = 0
    
    lChecksum = lChecksum + (POS_01_MULTIPLIER * CLng(Mid(sKey, 1, 1)))
    lChecksum = lChecksum + (POS_02_MULTIPLIER * CLng(Mid(sKey, 2, 1)))
    lChecksum = lChecksum + (POS_03_MULTIPLIER * CLng(Mid(sKey, 3, 1)))
    lChecksum = lChecksum + (POS_04_MULTIPLIER * CLng(Mid(sKey, 4, 1)))
    lChecksum = lChecksum + (POS_05_MULTIPLIER * CLng(Mid(sKey, 5, 1)))
    lChecksum = lChecksum + (POS_06_MULTIPLIER * CLng(Mid(sKey, 6, 1)))
    lChecksum = lChecksum + (POS_07_MULTIPLIER * CLng(Mid(sKey, 7, 1)))
    lChecksum = lChecksum + (POS_08_MULTIPLIER * CLng(Mid(sKey, 8, 1)))
    lChecksum = lChecksum + (POS_09_MULTIPLIER * CLng(Mid(sKey, 9, 1)))
    lChecksum = lChecksum + (POS_10_MULTIPLIER * CLng(Mid(sKey, 10, 1)))
    lChecksum = lChecksum + (POS_11_MULTIPLIER * CLng(Mid(sKey, 11, 1)))
    lChecksum = lChecksum + (POS_12_MULTIPLIER * CLng(Mid(sKey, 12, 1)))
    lChecksum = lChecksum + (POS_13_MULTIPLIER * CLng(Mid(sKey, 13, 1)))
    lChecksum = lChecksum + (POS_14_MULTIPLIER * CLng(Mid(sKey, 14, 1)))
    
    If Len(sKey) = KEY_LENGTH Then
        lChecksum = lChecksum + (POS_15_MULTIPLIER * CLng(Mid(sKey, 15, 1)))
        lChecksum = lChecksum + (POS_16_MULTIPLIER * CLng(Mid(sKey, 16, 1)))
        lChecksum = lChecksum + (POS_17_MULTIPLIER * CLng(Mid(sKey, 17, 1)))
    End If
    
    GetCheckSum = lChecksum
End Function

Public Sub GetSystemRegistrationKey()
Dim sSQL        As String
Dim sVersion    As String

sSQL = "SELECT * FROM " & gcUserTour_UserTbl & " WHERE NAME = '" & gcTOURSYSTEM & "'"

ObjTour.RstSQL iSearcherDB, sSQL
If 0 = ObjTour.RstRecordCount(iSearcherDB) Then
' Add record
  With ObjTour
    .AddNew (iSearcherDB)
    .DBSetField gcUserTour_UserTbl_Name, gcTOURSYSTEM, iSearcherDB
    
     m_UsageCount = 1
    .DBSetField gcUserTour_UserTbl_BitField, m_UsageCount, iSearcherDB
    
     m_Key = "12345222233334444"
    .DBSetField gcUserTour_UserTbl_DataPath, m_Key, iSearcherDB
    
    'Store the Database version number with System record - Password field
    sVersion = CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision)
    .DBSetField gcUserTour_UserTbl_PassWord, sVersion, iSearcherDB
    
    .Update iSearcherDB
  End With
Else
' Read Record
m_Key = ObjTour.DBGetField(gcUserTour_UserTbl_DataPath, iSearcherDB)
m_UsageCount = CLng(ObjTour.DBGetField(gcUserTour_UserTbl_BitField, iSearcherDB)) + 1

With ObjTour
  .Edit (iSearcherDB)
  .DBSetField gcUserTour_UserTbl_BitField, m_UsageCount, iSearcherDB
  .Update iSearcherDB
End With

End If

End Sub

Public Property Get UsageCount() As Long
    UsageCount = m_UsageCount
End Property

Public Function ShowTrialNotice() As Long

    ' Set default Licensing result
    If Me.UsageCount >= gcTRIALLENGTH Then
        cLicense.Licensing_Results = 1 ' Cancel on Expire
    Else
        cLicense.Licensing_Results = 0 ' Later...
    End If
    
    
    frmTrialNotice.UsageCount = Me.UsageCount
    frmTrialNotice.Show vbModal

    ShowTrialNotice = cLicense.Licensing_Results

    Unload frmTrialNotice
    
End Function

Public Sub SetSystemRegistrationKey(ByVal sKey As String)

Dim sSQL As String

sSQL = "SELECT * FROM " & gcUserTour_UserTbl & " WHERE NAME = '" & gcTOURSYSTEM & "'"

ObjTour.RstSQL iSearcherDB, sSQL
If 0 = ObjTour.RstRecordCount(iSearcherDB) Then
' Add record
  With ObjTour
    .AddNew (iSearcherDB)
    .DBSetField gcUserTour_UserTbl_Name, gcTOURSYSTEM, iSearcherDB
    
     m_UsageCount = 1
    .DBSetField gcUserTour_UserTbl_BitField, m_UsageCount, iSearcherDB
    
     m_Key = sKey
    .DBSetField gcUserTour_UserTbl_DataPath, m_Key, iSearcherDB
    
    .Update iSearcherDB
  End With
Else
' Edit record
  With ObjTour
    .Edit iSearcherDB
     m_Key = sKey
    .DBSetField gcUserTour_UserTbl_DataPath, m_Key, iSearcherDB
    .Update iSearcherDB
  End With

End If

End Sub


Public Property Get Key_Registration() As String
    Key_Registration = m_Key
End Property


Public Property Get Licensing_Results() As Long
    Licensing_Results = m_Results
End Property

Public Property Let Licensing_Results(ByVal lNewValue As Long)
    m_Results = lNewValue
End Property
