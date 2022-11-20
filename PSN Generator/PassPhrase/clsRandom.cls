VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRndData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class module designed to manipulate the Visual Basic random number generator and the data produced by it."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ***************************************************************************
' Module:        clsRndData
'
' Description:   Build a random generated data string using my unique
'                gathering methods.  There are 3 ways to create a seed value
'                for the random number generator.  See these routines:
'
'                      Seed1_Random_Generator
'                      Seed2_Random_Generator
'                      Seed3_Random_Generator
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Modified to use byte arrays and changed the logic flow.
' ***************************************************************************

' ---------------------------------------------------------------------------
' NOTE:  You can use either the GetTickCoount() or GetTickcount() API to
'        perform the same function.  Get the current number of milliseconds
'        since the last reboot.
' ---------------------------------------------------------------------------
' The timeGetTime() API will capture the time in milliseconds.  The counter
' overflows after 596 hours (24.8 days) from the last reboot.
' ---------------------------------------------------------------------------
  Private Declare Function timeGetTime Lib "winmm.dll" () As Long

' ---------------------------------------------------------------------------
' The GetTickCount() API will capture the time in milliseconds.  The counter
' overflows after 1192.8 hours (49.7 days) from the last reboot.
' ---------------------------------------------------------------------------
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  
' ---------------------------------------------------------------------------
' API function for obtaining volume information
' ---------------------------------------------------------------------------
  Private Declare Function GetVolumeInformation Lib "kernel32" _
          Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
          ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
          lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
          lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
          ByVal nFileSystemNameSize As Long) As Long

Public Function Rnd2(sngLow As Single, sngHigh As Single) As Single
  
' ***************************************************************************
' Routine:       Rnd2
'
' Description:   Create a random value between two values.  We are assuming
'                that the random number generator has already been seeded.
'
' Parameters:    sngLow  - Low end value
'                sngHign - High end value
'
' Return Values: A random generated value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-JUN-1999  The VB2TheMax Team  fbalena@vb2themax.com
'              Routine created
' 03-OCT-2000  Kenneth Ives  kenaso@home.com
'              Modified and documented
' ***************************************************************************
  DoEvents        ' gives the system clock time to advance a few milliseconds
  Randomize (CDbl(Now()) + Timer)    ' Reseed the VB random number generator
  Rnd2 = (Rnd * (sngHigh - sngLow)) + sngLow

End Function

Private Function Create_Char_Data(bKeyboardChars As Boolean, _
                   Optional lngDataLength As Long = 50, _
                   Optional bUse_Letters_Only As Boolean = False) As Variant

' ***************************************************************************
' Routine:       Create_Char_Data
'
' Description:   This routine will build a string of random generated data
'                based on type of data desired and the length of the string.
'
' Parameters:    bKeyboardChars - TRUE = Printable keyboard characters only.
'                     ASCII values 33 to 126
'                                 FALSE = All ASCII values (0 to 255)
'                lngDataLength - amount of data to create
'                bUse_Letters_Only - TRUE/FALSE on using just alphanetic chars
'
' Returns:       byte array of random generated data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Modified to use byte arrays and changed the logic flow.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex  As Long
  Dim intChar   As Integer
  Dim arByte()  As Byte
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  Erase arByte()
  ReDim arByte(lngDataLength)
  
' ---------------------------------------------------------------------------
' Build a byte array of randomly generated ASCII values
' ---------------------------------------------------------------------------
  For lngIndex = 0 To lngDataLength - 1
  
       ' Printable keyboard characters only
      If bKeyboardChars And bUse_Letters_Only Then
          ' Loop until we get an alphabetic character (A-Z, a-z)
          Do
              intChar = Int(Rnd2(65, 122))
              
              Select Case intChar
                     Case 91 To 96: intChar = 0
              End Select
              
          Loop Until (intChar > 32 And intChar < 127)
      ElseIf bKeyboardChars And Not bUse_Letters_Only Then
                ' Loop until we get a printable character
          ' value of 33 to 126 (0x21 to 0x7E)
          ' should only be one iteration
          Do
              intChar = Int(Rnd2(33, 126))
          Loop Until (intChar > 32 And intChar < 127)
      
      Else
          ' generate integer value 0 to 255
          intChar = Int(Rnd * 255)
      End If
      
      ' save the ASCII decimal value
      arByte(lngIndex) = intChar
  Next
  
' ---------------------------------------------------------------------------
' Return the array of data
' ---------------------------------------------------------------------------
  Create_Char_Data = arByte()
  
' ---------------------------------------------------------------------------
' Empty the variables for security reasons
' ---------------------------------------------------------------------------
  intChar = 0
  Erase arByte()    ' empty the array
  ReDim arByte(0)   ' resize to the smallest element
  
End Function

Public Function Mix_The_Data(arByte() As Byte) As Variant

' ***************************************************************************
' Routine:       Mix_The_Data
'
' Description:   This function will accept an incoming string and will
'                perform an XoR binary compare on each byte to create a new
'                ASCII decimal value.
'
' Parameters:    arByte() - byte array of data
'
' Returns:       Reshuffled byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 11-NOV-2000  Kenneth Ives  kenaso@home.com
'              Created routine
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Modified to use byte arrays and changed the logic flow.
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex1      As Long
  Dim lngIndex2      As Long
  Dim lngNewIndex    As Long
  Dim lngMax         As Long
  Dim bytTemp        As Byte
  Dim intOuterLoop   As Integer
  
' ---------------------------------------------------------------------------
' Determine the maximum number of elements in array
' ---------------------------------------------------------------------------
  lngMax = UBound(arByte) - 1
  
' ---------------------------------------------------------------------------
' See if anything was passed that has to be mixed
' ---------------------------------------------------------------------------
  If lngMax <= 0 Then
      Mix_The_Data = arByte()
      Exit Function
  End If

' ---------------------------------------------------------------------------
' Create a number between 5 to 10
' ---------------------------------------------------------------------------
  intOuterLoop = Int(Rnd2(5, 10))
  
' ---------------------------------------------------------------------------
' The array will now undergo multiple mixing operations in a series of
' loops (5 to 10 times).
' ---------------------------------------------------------------------------
  For lngIndex1 = 1 To intOuterLoop
  
      Randomize CDbl(Now()) + Timer ' reseed random generator
         
      ' go thru the input array and randomly pick an element
      ' to move to another position within the array.
      For lngIndex2 = 0 To lngMax
      
          ' generate a new index 0 to the max number of elements
          Do
              lngNewIndex = Int(Rnd * lngMax)
          Loop Until lngNewIndex <> lngIndex2
          
          ' swap the the data
          bytTemp = arByte(lngIndex2)
          arByte(lngIndex2) = arByte(lngNewIndex)
          arByte(lngNewIndex) = bytTemp
      Next
  Next

' ---------------------------------------------------------------------------
' Return the reshuffled data
' ---------------------------------------------------------------------------
  Mix_The_Data = arByte()

End Function

Public Function Build_Random_Data(lngDataLength As Long, _
                    Optional bUse_Hex_Chars As Boolean = False, _
                    Optional bUse_Keyboard_Chars As Boolean = False, _
                    Optional bUse_Letters_Only As Boolean = False) As String

' ***************************************************************************
' Routine:       Build_Random_Data
'
' Description:   Build a random data string.  Keep it within reason so as
'                not to clog your memory processes.  The maximum value I
'                have ever used was 32768 (1024 * 32).
'
'
' Parameters:    lngDataLength = Length of the data string
'                bUse_Hex_Chars - (Optional) (Default) FALSE - do not
'                      convert to hex
'                                 TRUE  - convert each byte to its 2-char hex
'                      representation before returning (Ex:  15  becomes  0F)
'                bUse_Keyboard_Chars (Optional) (Default) FALSE - Use all 256
'                      ASCII decimal values (0-255)
'                                 TRUE  - Use only printable visible ASCII
'                      decimal values (33-126)
'                bUse_Letters_Only (Optional) (Default) FALSE-Use any char
'                      generated.
'                                 TRUE - Return only letters of the alphabet
'
' Return Values: String of random generated data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Modified to use byte arrays and changed the logic flow.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex    As Long
  Dim lngMax      As Long
  Dim arByte()    As Byte
  Dim arString()  As String
  Dim strTmp      As String
  Dim intChoice   As Integer
  
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Erase arByte()
  Erase arString()
  strTmp = ""
  
' ---------------------------------------------------------------------------
' If length of the desired data string equal zero then leave
' ---------------------------------------------------------------------------
  If lngDataLength = 0 Then
      Exit Function
  Else
      ReDim arByte(lngDataLength)
  End If
  
' ---------------------------------------------------------------------------
' Loop 10 times while genrating a number 1 to 3.  We will save the last
' iteration to determine which seeding process to use.
' ---------------------------------------------------------------------------
  For lngIndex = 1 To 10
      intChoice = CInt(Rnd2(1, 3))
  Next
  
  Select Case intChoice
         Case 1:    Seed1_Random_Generator
         Case 2:    Seed2_Random_Generator
         Case 3:    Seed3_Random_Generator
         Case Else: Seed1_Random_Generator   ' safety net
  End Select
  
' ---------------------------------------------------------------------------
' Build a string of randomly generated characters
' ---------------------------------------------------------------------------
  arByte = Create_Char_Data(bUse_Keyboard_Chars, lngDataLength, bUse_Letters_Only)
  
' ---------------------------------------------------------------------------
' Mix up the random generated data some more
' ---------------------------------------------------------------------------
  arByte = Mix_The_Data(arByte())

' ---------------------------------------------------------------------------
' See if the data is to be converted to hex format
' ---------------------------------------------------------------------------
  If bUse_Hex_Chars Then
      
      lngMax = lngDataLength * 2  ' double the size due to conversion
      ReDim arString(lngMax)      ' size the array before loading
      
      For lngIndex = 0 To lngDataLength - 1
          ' convert byte data to 2-char hex
          arString(lngIndex) = Right("00" & Hex(arByte(lngIndex)), 2)
      Next
  Else
      ReDim arString(lngDataLength)   ' resize the array
      
      ' convert byte data to string data
      For lngIndex = 0 To lngDataLength - 1
          arString(lngIndex) = Chr(arByte(lngIndex))
      Next
  End If
  
' ---------------------------------------------------------------------------
' Build an output string before leaving
' ---------------------------------------------------------------------------
  Build_Random_Data = Build_Output_String(arString())
  
' ---------------------------------------------------------------------------
' Empty variables for security purposes
' ---------------------------------------------------------------------------
  strTmp = vbNullString
  
  Erase arByte()    ' empty the array
  ReDim arByte(0)   ' resize to smallest size
  
  Erase arString()  ' empty the array
  ReDim arString(0) ' resize to smallest size
  
End Function

Private Function Build_Output_String(arstrInput() As String) As String

' ***************************************************************************
' Routine:       Build_Output_String
'
' Description:   This routine will build the final output string
'
' Parameters:    arstrInput() - input string array
'
' Returnss:      String of data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Created routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define variables
' ---------------------------------------------------------------------------
  Dim lngIndex   As Long
  Dim lngMax     As Long
  Dim strTmp     As String
  Dim strOutput  As String
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strOutput = ""
  strTmp = ""
  lngMax = UBound(arstrInput) - 1
  
' ---------------------------------------------------------------------------
' If no data then leave
' ---------------------------------------------------------------------------
  If lngMax < 0 Then
      Build_Output_String = ""
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' build the output string in 1024 byte chunks.
' ---------------------------------------------------------------------------
  For lngIndex = 0 To lngMax
      
      ' append the data to a temp string and check its size.
      ' if it is 1024 bytes long then append it to the output
      ' string.  Empty the temp string and start over.
      strTmp = strTmp & arstrInput(lngIndex)
      
      If Len(strTmp) Mod 1024 = 0 Then
          strOutput = strOutput & strTmp
          strTmp = ""
      End If
  Next
  
' ---------------------------------------------------------------------------
' Return the completed string.  Append the Temp string in case there was
' some data left over.
' ---------------------------------------------------------------------------
  Build_Output_String = strOutput & strTmp

' ---------------------------------------------------------------------------
' Empty variables for security reasons
' ---------------------------------------------------------------------------
  strOutput = vbNullString
  strTmp = vbNullString
  
End Function

Public Sub Seed1_Random_Generator()

' ***************************************************************************
' Routine:       Seed1_Random_Generator
'
' Description:   Generates a 64-bit (8-byte word) double precision number
'                as the new seed for the random generator.  The random
'                generator takes this number and then creates a 32-bit
'                (4-byte) single precision number.  Reseed as often as you
'                want.
'
'                This type of random generation is needed for gaming,
'                password creation, encryption, etc.  A calculated value
'                that cannot be easily compromised.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Removed obsolete variables
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define variables
' ---------------------------------------------------------------------------
  Dim bTestMode   As Boolean
  Dim lngRetVal   As Long
  Dim lngVolSn    As Long
  Dim lngTickCnt  As Long
  
  ' 64-bit(8-byte) floating-point number variables
  Dim Pi          As Double
  Dim varRndSeed  As Variant
  Dim varDecimal  As Variant
  
Top_Of_Test_Loop:
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  lngVolSn = 0
  lngTickCnt = 0
  varRndSeed = 0
  varDecimal = 0
  bTestMode = False
  
' ---------------------------------------------------------------------------
' Calculate Pi to 14 decimal places which should be 3.14159265358979
' ---------------------------------------------------------------------------
  Pi = Atn(1) * 4
  
' ---------------------------------------------------------------------------
' get number of milliseconds since the last reboot
' ---------------------------------------------------------------------------
  lngTickCnt = GetTickCount()
  
' ---------------------------------------------------------------------------
' Pre-Initialize the VB random generator
' ---------------------------------------------------------------------------
  Randomize lngTickCnt + Pi
  
' ---------------------------------------------------------------------------
' Use API call to get volume serial number of drive C:.  Assumes there will
' always be a drive C: since we are using Windows.
' ---------------------------------------------------------------------------
  lngRetVal = GetVolumeInformation("C:\", "", 0, lngVolSn, 0, 0, "", 0)
   
' ---------------------------------------------------------------------------
' Create a new seed value based on the following:
'
'  1.  System data/time stamp in double precision
'  2.  Add the Random generated value from Pi
'  3.  Add the Random generated value of the tickcount
'  4.  Add the Volume serial number of drive C:
'  5.  Add Pi to the running total
'  6.  Capture just the decimal portion of the running total
'  7.  Add Pi to the decimal value
'  8.  Multiply new decimal value by random value of 1
'  9.  Convert this decimal value using CDec() function of VB
' 10.  Add running total and decimal value together
' ---------------------------------------------------------------------------
  varRndSeed = CDbl(Now()) + (Rnd * Pi)
  varRndSeed = varRndSeed + CDbl((Rnd * lngTickCnt) + lngVolSn)
  varRndSeed = varRndSeed + (Rnd * lngVolSn) + Pi
  
  ' save decimal value only
  varDecimal = Mid(varRndSeed, InStr(varRndSeed, "."))
  varDecimal = CDec((varDecimal + Pi) * (Rnd * 1))
  varRndSeed = varRndSeed + varDecimal
  
' ---------------------------------------------------------------------------
' For debugging only.
' To turn on debugging, uncomment the next line to set
' bTestMode = True
'
' Sample Output:    238688296.59674506039113906
'                   260981561.7276916851491254
'                   142932264.06364296356874
'                   209517593.4024467258857
'                   250310710.85324095735983
'                   165968569.997226489607152
' ---------------------------------------------------------------------------
  'bTestMode = True
  If bTestMode Then
      Debug.Print varRndSeed    ' Display the output
      Stop                      ' press F5 to continue
      GoTo Top_Of_Test_Loop     ' Start over
  End If
  
' ---------------------------------------------------------------------------
' Initialize the random generator using the new calculated value (64-bit
' value) as the seed.  Randomize returns a data type that stores a 32-bit
' (4-byte) single-precision floating-point number.
' ---------------------------------------------------------------------------
  Randomize varRndSeed
  
' ---------------------------------------------------------------------------
' Empty all variables for security purposes
' ---------------------------------------------------------------------------
  lngVolSn = 0
  varRndSeed = Empty
  varDecimal = Empty
  
End Sub

Public Sub Seed2_Random_Generator()

' ***************************************************************************
' Routine:       Seed2_Random_Generator
'
' Description:   Generates a 64-bit (8-byte word) double precision number
'                as the new seed for the random generator.  The random
'                generator takes this number and then creates a 32-bit
'                (4-byte) single precision number.  Reseed as often as you
'                want.
'
'                This type of random generation is needed for gaming,
'                password creation, encryption, etc.  A calculated value
'                that cannot be easily compromised.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine
' 19-MAY-2000  Kenneth Ives  kenaso@home.com
'              Modified starting postion calculation
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Removed obsolete variables
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim sngMultiplier  As Single
  Dim lngFinish      As Long
  Dim lngIndex       As Long
  Dim lngRetVal      As Long
  Dim lngVolSn       As Long
  Dim lngTickCnt     As Long
  Dim bTestMode      As Boolean
    
  ' 64-bit(8-byte) floating-point number variables
  Dim Pi             As Double
  Dim varRndSeed     As Variant
  Dim varDecimal     As Variant
  
  
Top_Of_Test_Loop:
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  bTestMode = False
  sngMultiplier = 1.321        ' Pick your own number. I chose 1.321
  varRndSeed = 0
  varDecimal = 0
  lngVolSn = 0
  
' ---------------------------------------------------------------------------
' Calculate Pi to 14 decimal places which should be 3.14159265358979
' ---------------------------------------------------------------------------
  Pi = Atn(1) * 4
  
' ---------------------------------------------------------------------------
' get number of milliseconds since the last reboot
' ---------------------------------------------------------------------------
  lngTickCnt = GetTickCount()
  
' ---------------------------------------------------------------------------
' Pre-Initialize the VB random generator
' ---------------------------------------------------------------------------
  Randomize lngTickCnt + Pi
  
' ---------------------------------------------------------------------------
' Use API call to get volume serial number of drive C:.  Assumes there will
' always be a drive C: since we are using Windows.
' ---------------------------------------------------------------------------
  lngRetVal = GetVolumeInformation("C:\", "", 0, lngVolSn, 0, 0, "", 0)
   
' ---------------------------------------------------------------------------
' Generate the start and end points for the next loop
' ---------------------------------------------------------------------------
  lngFinish = Int(Rnd2(5, 15))  ' create a value of 5 to 15
  
' ---------------------------------------------------------------------------
' Calculate the new seed value in a loop.  Save the last iteration value
' ---------------------------------------------------------------------------
  For lngIndex = 1 To lngFinish
  
      ' initialize with the date/time stamp plus the
      ' current number of milliseconds since the last
      ' reboot
      varRndSeed = CDbl(Now()) + lngTickCnt
          
      ' Now we start making a unique number
      varRndSeed = Sqr(varRndSeed) + Pi + lngVolSn
      varRndSeed = CDbl((Timer + varRndSeed) * sngMultiplier)
      varRndSeed = varRndSeed + lngTickCnt + lngVolSn
      varRndSeed = (varRndSeed / (lngIndex + 1))
      
      ' if the looping index is divided equally by three
      ' then perform some calculation changes in one
      ' direction else go in another direction.  Helps
      ' us from falling into a pattern.
      If lngIndex Mod 3 = 0 Then
          sngMultiplier = sngMultiplier * 0.5
          varRndSeed = varRndSeed * Pi
      Else
          sngMultiplier = sngMultiplier + 0.5
          varRndSeed = varRndSeed * sngMultiplier
      End If
  
      ' save just decimal value
      varDecimal = Mid(varRndSeed, InStr(varRndSeed, "."))
      varDecimal = CDec(varDecimal + Pi)
      varRndSeed = varRndSeed + varDecimal + (Rnd * 1)
      
      ' gives the system clock time to advance a few milliseconds
      DoEvents
  Next

' ---------------------------------------------------------------------------
' For debugging only.  To turn on debugging, uncomment the next line to set
' bTestMode = True
'
' Sample output:   29595994.47378014152651
'                  45337025.66965503496078
'                 113368419.68224317930845
'                  27484260.41284549482284
'                  42173945.8184387370566633
'                 139529402.7532063705843
' ---------------------------------------------------------------------------
 'bTestMode = True
  If bTestMode Then
      Debug.Print varRndSeed    ' Display the output
      Stop                      ' press F5 to continue
      GoTo Top_Of_Test_Loop     ' Start over
  End If
  
' ---------------------------------------------------------------------------
' Initialize the random generator using the new calculated value (64-bit
' value) as the seed.  Randomize returns a data type that stores a 32-bit
' (4-byte) single-precision floating-point number.
' ---------------------------------------------------------------------------
  Randomize varRndSeed

' ---------------------------------------------------------------------------
' Empty variables for security reasons
' ---------------------------------------------------------------------------
  lngVolSn = 0
  sngMultiplier = 0
  varRndSeed = Empty
  varDecimal = Empty

End Sub

Public Sub Seed3_Random_Generator()

' ***************************************************************************
' Routine:       Seed3_Random_Generator
'
' Description:   Generates a 64-bit (8-byte word) double precision number
'                as the new seed for the random generator.  The random
'                generator takes this number and then creates a 32-bit
'                (4-byte) single precision number.  Reseed as often as you
'                want.
'
'                This type of random generation is needed for gaming,
'                password creation, encryption, etc.  A calculated value
'                that cannot be easily compromised.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@home.com
'              Created routine
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Removed obsolete variables
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim varRndSeed     As Variant
  Dim intIndex       As Integer
  Dim intMax         As Integer
  Dim Pi             As Double
  Dim lngLowerLimit  As Long
  Dim lngUpperLimit  As Long
  Dim lngTickCnt     As Long
  Dim bTestMode      As Boolean
  
Top_Of_Test_Loop:

' ---------------------------------------------------------------------------
' Initialize the VB random number generator seed
' ---------------------------------------------------------------------------
  bTestMode = False
  lngLowerLimit = 0
  lngUpperLimit = 0
  varRndSeed = Empty
  
' ---------------------------------------------------------------------------
' Calculate Pi to 14 decimal places which should be Pi = 3.14159265358979
' ---------------------------------------------------------------------------
  Pi = Atn(1) * 4
  
' ---------------------------------------------------------------------------
' get number of milliseconds since the last reboot
' ---------------------------------------------------------------------------
  lngTickCnt = GetTickCount()
  
' ---------------------------------------------------------------------------
' Pre-Initialize the VB random generator
' ---------------------------------------------------------------------------
  Randomize lngTickCnt + Pi
  
' ---------------------------------------------------------------------------
' Initialize the lower and upper limits.  Should only be one iteration.
' ---------------------------------------------------------------------------
  While lngLowerLimit >= lngUpperLimit
      lngLowerLimit = Int(Rnd * 8192)         '    0 to 8192
      lngUpperLimit = Int(Rnd2(8193, 32767))  ' 8193 to 32767
  Wend
  
' ---------------------------------------------------------------------------
' Initialize the new seed.  Should only be one iteration.
' ---------------------------------------------------------------------------
  While varRndSeed = 0
      varRndSeed = ((lngUpperLimit - lngLowerLimit + 1) * Rnd)
      varRndSeed = varRndSeed + lngLowerLimit + CDbl(Now)
  Wend
 
' ---------------------------------------------------------------------------
' Create a number 5 to 15
' ---------------------------------------------------------------------------
  intMax = CInt(Rnd2(5, 15))
  
' ---------------------------------------------------------------------------
' Loop thru and recalculate the value of the new seed
' ---------------------------------------------------------------------------
  For intIndex = 0 To intMax
      varRndSeed = ((lngUpperLimit / lngLowerLimit) + (varRndSeed / 25)) * lngLowerLimit
      varRndSeed = Mid(varRndSeed, InStr(varRndSeed, "."))  ' save decimal value only
      varRndSeed = CDec(varRndSeed)
  Next

' ---------------------------------------------------------------------------
' For debugging only.  To turn on debugging, uncomment the
' next line to set bTestMode = True
'
' Sample Output:    0.585557684654190418663810265
'                   0.0135991812904
'                   0.990122159816596628591411226
'                   0.23230047005405142543520764
'                   0.475468645882074791294757949
' ---------------------------------------------------------------------------
  'bTestMode = True
  If bTestMode Then
      Debug.Print varRndSeed    ' display output
      Stop                      ' press F5 to continue
      GoTo Top_Of_Test_Loop     ' Start over
  End If
  
' ---------------------------------------------------------------------------
' Initialize the random generator using the new calculated value (64-bit
' value) as the seed.  Randomize returns a data type that stores a 32-bit
' (4-byte) single-precision floating-point number.
' ---------------------------------------------------------------------------
  Randomize varRndSeed
 
' ---------------------------------------------------------------------------
' Empty the variables for security reasons
' ---------------------------------------------------------------------------
  varRndSeed = Empty
  lngUpperLimit = 0
  lngLowerLimit = 0
  
End Sub

