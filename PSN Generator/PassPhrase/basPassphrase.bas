Attribute VB_Name = "basPassphrase"
Option Explicit

' ---------------------------------------------------------------------------
' Define variables to be used throughout this form. Variables prefixed with
' "m_" are module level.  Prefixed with "g_" are used throughout the
' application.
' ---------------------------------------------------------------------------
  ' -- String variables
  Public g_strDBName        As String
  Public g_strSpecialWord   As String
  Public g_strPassphrase    As String
  Private m_strTargetTitle  As String
  
  ' -- Boolean variables
  Public g_bStop_Pressed    As Boolean
  Public g_bCreate_Password As Boolean
  Private m_bFoundApp       As Boolean
  
  ' -- Long Integer variables
  Public g_lngNbrOfWords    As Long
  
  ' -- Integer variables
  Public g_intPosition      As Integer
  Public g_intMinLength     As Integer
  Public g_intMaxLength     As Integer
  Public g_intTypeCase      As Integer
  
  ' -- Arrays
  Public g_arOmit(1 To 35)  As Integer
  Public g_arWords()        As String
  Public g_arlngData()      As Long

Public Function Build_Using_Multiple_Lengths() As String
  
' ***************************************************************************
' Routine:       Build_Using_Multiple_Lengths
'
' Description:   Build a passphrase using varying length words
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim strTmp     As String
  Dim strTable   As String
  Dim lngIndex1  As Long
  Dim lngIndex2  As Long
  Dim lngIndex3  As Long
  Dim arMulti()  As Variant
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strTmp = ""
  lngIndex3 = 1
  Erase arMulti()
  
' ---------------------------------------------------------------------------
' start looping thru the tables
' ---------------------------------------------------------------------------
  For lngIndex1 = g_intMinLength To g_intMaxLength
      
      Erase g_arWords()
      strTable = ""
      
      Select Case lngIndex1
             Case 3: strTable = "Char_3"
             Case 4: strTable = "Char_4"
             Case 5: strTable = "Char_5"
             Case 6: strTable = "Char_6"
             Case 7: strTable = "Char_7"
             Case 8: strTable = "Char_8"
      End Select
        
      ' Read the dataabse to get the words
      If Not ReadDatabase(strTable, g_lngNbrOfWords) Then
          Build_Using_Multiple_Lengths = ""
          Exit Function
      Else
          For lngIndex2 = 1 To g_lngNbrOfWords
              ReDim Preserve arMulti(lngIndex3)
              arMulti(lngIndex3) = g_arWords(lngIndex2)
              lngIndex3 = lngIndex3 + 1
              
              ' if stop button pressed then leave
              DoEvents
              If g_bStop_Pressed Then
                  strTmp = ""
                  GoTo Normal_Exit
              End If
          Next
      End If
  Next
  
  On Error GoTo Multi_Length_Error
   
' ---------------------------------------------------------------------------
' how many words are there?
' ---------------------------------------------------------------------------
  lngIndex3 = UBound(arMulti)
  
' ---------------------------------------------------------------------------
' Build a random list before building the passphrase
' ---------------------------------------------------------------------------
  Create_Random_Pointers lngIndex3, g_arlngData(), True
  
' ---------------------------------------------------------------------------
' Build the passphrase.  See if we are to use special words
' ---------------------------------------------------------------------------
  If Len(Trim(g_strSpecialWord)) <> 0 Then
      ' if only one word was chosen, ignore
      ' all other options
      If g_lngNbrOfWords = 1 Then
          strTmp = arMulti(1)
      Else
          ' build the passphrase with the multiple words
          lngIndex1 = 0
          lngIndex3 = 0
          strTmp = ""
          
          Do
              lngIndex1 = lngIndex1 + 1
          
              ' get index for the word list
              lngIndex3 = g_arlngData(lngIndex1)
              
              ' determine the position of the special word
              ' in the passphrase string
              If lngIndex1 = g_intPosition Then
                  strTmp = strTmp & g_strSpecialWord & " "
              Else
                  ' Load the passphrase with words
                  strTmp = strTmp & arMulti(lngIndex3) & " "
              End If
              
              ' if stop button pressed then leave
              DoEvents
              If g_bStop_Pressed Then
                  strTmp = ""
                  GoTo Normal_Exit
              End If
          Loop Until lngIndex1 = g_lngNbrOfWords
      End If
  Else
      ' if no special words requested then build the passphrase
      For lngIndex1 = 1 To g_lngNbrOfWords
          
          ' get index for the word list
          lngIndex3 = g_arlngData(lngIndex1)
              
          ' build the passphrase
          strTmp = strTmp & arMulti(lngIndex3) & " "
              
          ' if stop button pressed then leave
          DoEvents
          If g_bStop_Pressed Then
              strTmp = ""
              GoTo Normal_Exit
          End If
      Next
  End If
    
Normal_Exit:
' ---------------------------------------------------------------------------
' remove all leading and trailing blanks
' ---------------------------------------------------------------------------
  Build_Using_Multiple_Lengths = Trim(strTmp)
  Exit Function
  
Multi_Length_Error:
' ---------------------------------------------------------------------------
' Display an error message
' ---------------------------------------------------------------------------
  MsgBox "Error" & CStr(Err.Number) & Err.Description & vbLf & _
         "Routine:  Build_Using_Multiple_Lengths", vbOKOnly, _
         "Building Passphrase"
  Build_Using_Multiple_Lengths = ""
 
End Function

Public Function Build_Using_Same_Length(sTable As String) As String
  
' ***************************************************************************
' Routine:       Build_Using_Same_Length
'
' Description:   Build a passphrase using the same length words
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim strTmp    As String
  Dim intIndex  As Integer
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strTmp = ""
  
' ---------------------------------------------------------------------------
' Read the dataabse to get the words
' ---------------------------------------------------------------------------
  If Not ReadDatabase(sTable, g_lngNbrOfWords) Then
      Build_Using_Same_Length = ""
      Exit Function
  End If
   
  On Error GoTo Same_Length_Error
  
' ---------------------------------------------------------------------------
' Build the passphrase.  See if we are to use special words
' ---------------------------------------------------------------------------
  If Len(Trim(g_strSpecialWord)) <> 0 Then
      ' if only one word was chosen, ignore
      ' all other options
      If g_lngNbrOfWords = 1 Then
          strTmp = g_arWords(1)
      Else
          ' build the passphrase with the multiple words
          For intIndex = 1 To g_lngNbrOfWords
              ' this determines the position of the special word
              ' in the passphrase string
              If intIndex = g_intPosition Then
                  strTmp = strTmp & g_strSpecialWord & " "
              Else
                  ' Load the passphrase with words
                  strTmp = strTmp & g_arWords(intIndex) & " "
              End If
              
              ' if stop button pressed then leave
              DoEvents
              If g_bStop_Pressed Then
                  strTmp = ""
                  GoTo Normal_Exit
              End If
          Next
      End If
  Else
      ' if no special words requested then build the passphrase
      For intIndex = 1 To g_lngNbrOfWords
          strTmp = strTmp & g_arWords(intIndex) & " "
              
          ' if stop button pressed then leave
          DoEvents
          If g_bStop_Pressed Then
              strTmp = ""
              GoTo Normal_Exit
          End If
      Next
  End If
    
Normal_Exit:
' ---------------------------------------------------------------------------
' remove all leading and trailing blanks
' ---------------------------------------------------------------------------
  Build_Using_Same_Length = Trim(strTmp)
  Exit Function
  
Same_Length_Error:
' ---------------------------------------------------------------------------
' Display an error message
' ---------------------------------------------------------------------------
  MsgBox "Error" & CStr(Err.Number) & Err.Description & vbLf & _
         "Routine:  Build_Using_Same_Length", vbOKOnly, "Building Passphrase"
  Build_Using_Same_Length = ""
 
End Function
   
Public Function ReadDatabase(sTable As String, lWordCount As Long) As Boolean

' ***************************************************************************
' Routine:       ReadDatabase
'
' Description:   Reads the PWord.mdb database and retrieves the number of
'                words requested.
'
' Parameters:    sTable - one of the tables in the database to be read
'                lWordCount - Number of words to extract from the table
'
' Return Values: True or False
'
' Special Logic: We will try to find a word three times.  If we cannot locate
'                it, the program will terminate.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-JAN-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************
  On Error GoTo ReadDatabase_Error

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim connPWords   As ADODB.Connection  'Connect to the ADO Data Type
  Dim connRS       As ADODB.Recordset   'Record Source Name
  Dim SQLstmt      As String            'SQL Statement String(s)
  Dim lngRecCount  As Long
  Dim intBadCount  As Integer
  Dim lngIndex     As Long
  Dim strCriteria  As String
  Dim bNoMoreData  As Boolean
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  intBadCount = 0
  Erase g_arlngData()
  SQLstmt = "SELECT * FROM [" & sTable & "]"
  
' ---------------------------------------------------------------------------
' Establish the Connection
' ---------------------------------------------------------------------------
  Set connPWords = New ADODB.Connection
  connPWords.CursorLocation = adUseClient
  connPWords.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & g_strDBName & ";" & "Persist Security Info=False"

' ---------------------------------------------------------------------------
' Open the Connection
' ---------------------------------------------------------------------------
  connPWords.Open

' ---------------------------------------------------------------------------
' Get the Records
' ---------------------------------------------------------------------------
  Set connRS = New ADODB.Recordset
  connRS.Open SQLstmt, connPWords, adOpenStatic, adLockOptimistic, adCmdText
  lngRecCount = connRS.RecordCount  ' save the reocrd count
  
' ---------------------------------------------------------------------------
' see if there is data in the table
' ---------------------------------------------------------------------------
  If lngRecCount < 1 Then
      GoTo ReadDatabase_Error
  End If
  
' ---------------------------------------------------------------------------
' Generate a list of non-repeating numbers based on the record count from
' the table
' ---------------------------------------------------------------------------
  Create_Random_Pointers lngRecCount, g_arlngData(), True

' ---------------------------------------------------------------------------
' We will try to extract selected words from the table based on the numbers
' generated in the above statement.  If we are not successful after three
' tries then we must assume the table is corrupted.
' ---------------------------------------------------------------------------
  Do
       ' Initialization
       lngIndex = 0
       bNoMoreData = False
       Erase g_arWords()
       ReDim g_arWords(1 To g_lngNbrOfWords)
      
      ' read the list of generated numbers and use them
      ' as pointer in the database
      Do
          strCriteria = "WrdNbr = " & g_arlngData(lngIndex)
         
          ' find the first occurance of this
          ' number in the table
          connRS.MoveFirst
          connRS.Find strCriteria
         
          If connRS.EOF Then
              bNoMoreData = True
              Exit Do
          End If
        
          lngIndex = lngIndex + 1
          g_arWords(lngIndex) = connRS!Word
         
      Loop Until lngIndex = lWordCount

      ' if we hit the end of the table, this
      ' flag will be TRUE
      If bNoMoreData Then
          intBadCount = intBadCount + 1  ' increment the bad counter
      Else
          Exit Do  ' success.  Exit the loop.
      End If
      
  Loop Until intBadCount = 3  ' We should never get to three
  
' ---------------------------------------------------------------------------
' see if we had to make three tries at the database.  If not then we were
' successful.
' ---------------------------------------------------------------------------
  If intBadCount = 3 Then
      MsgBox "Have tried three times to find three different words." & vbCrLf & _
             "The table may be corrupted.  Replace the PWords.dat" & vbCrLf & _
             "file with another copy.", vbCritical + vbOKOnly, "Reading Database"
      ReadDatabase = False
  Else
      ReadDatabase = True
  End If

  
CleanUp:
  connRS.Close        ' close the recordset
  connPWords.Close    ' close the database
  

Normal_Exit:
' ---------------------------------------------------------------------------
' free objects form memory
' ---------------------------------------------------------------------------
  Set connRS = Nothing
  Set connPWords = Nothing
  Exit Function


ReadDatabase_Error:
' ---------------------------------------------------------------------------
' Display an error message
' ---------------------------------------------------------------------------
  MsgBox "Error:  " & CStr(Err.Number) & " " & Err.Description & vbLf & _
         "Table is corrupted or empty.", vbOKOnly, "Reading Database"
  ReadDatabase = False
  GoTo Normal_Exit
  
End Function
Public Function Create_Random_Pointers(ByVal lMaxNbrs As Long, _
                                   arData() As Long, _
                                   Optional bUseNumberOne As Boolean = True)

' ***************************************************************************
' Routine:       Create_Random_Pointers
'
' Description:   This routine will receive a maximum value of numbers that
'                are to be used as pointers in an array.  An array is loaded
'                from 1 to this value.  The array is then mixed and the
'                values in each element are swapped with other elements.
'                This rearranges the pointers in the array because when the
'                array is retruned, it is read sequentially.
'
' Parameters:    lMaxNbrs - Total count of numbers to be returned
'
'                arData() - Return the numbers in this array.
'
'                bUseNumberOne - True or False.  Whether to a allow a one to
'                           be generated.  Default is TRUE.
'
' Return Values: An array of rearranged numbers
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 02-FEB-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************
    
' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex As Long
  
' ---------------------------------------------------------------------------
' resize output array to match the max size
' ---------------------------------------------------------------------------
  Erase arData()          ' Empty the array
  ReDim arData(lMaxNbrs)  ' resize the array
      
' ---------------------------------------------------------------------------
' preload the array starting with zero
' ---------------------------------------------------------------------------
  If g_bCreate_Password Then
      ' Lowest possible number is a 2
      For lngIndex = 0 To (lMaxNbrs - 1)
          arData(lngIndex) = lngIndex + 2
      Next
  Else
      ' Lowest possible number is a 1
      For lngIndex = 0 To (lMaxNbrs - 1)
          arData(lngIndex) = lngIndex + 1
      Next
  End If
  
' ---------------------------------------------------------------------------
' This array will now undergo several mixing operations in another routine
' ---------------------------------------------------------------------------
  arData = Mix_The_Data(arData())
  
End Function

Private Function Mix_The_Data(arData() As Long) As Variant

' ***************************************************************************
' Routine:       Mix_The_Data
'
' Description:   This function will accept an incoming string and will
'                swap array positions
'
' Parameters:    arData() - array of data
'
' Returns:       Rearranged data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-NOV-2000  Kenneth Ives  kenaso@home.com
'              Created routine
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex1      As Long
  Dim lngIndex2      As Long
  Dim lngNewIndex    As Long
  Dim lngMax         As Long
  Dim lngTemp        As Long
  
' ---------------------------------------------------------------------------
' Determine the maximum number of elements in array
' ---------------------------------------------------------------------------
  lngMax = UBound(arData) - 1
  
' ---------------------------------------------------------------------------
' See if anything was passed that has to be mixed
' ---------------------------------------------------------------------------
  If lngMax < 1 Then
      Mix_The_Data = arData()
      Exit Function
  End If

' ---------------------------------------------------------------------------
' The array will now undergo three mixing operations
' ---------------------------------------------------------------------------
  For lngIndex1 = 1 To 3
  
      Randomize CDbl(Now()) + Timer ' reseed random generator
         
      ' go thru the input array and randomly pick an element
      ' to move to another position within the array.
      For lngIndex2 = 0 To lngMax
      
          ' generate a new index 0 to the max number of elements.
          ' make sure the new index is not the same as the current
          ' index
          Do
              lngNewIndex = CLng(Rnd * lngMax)
          Loop Until lngNewIndex <> lngIndex2
          
          ' swap the the data
          lngTemp = arData(lngIndex2)
          arData(lngIndex2) = arData(lngNewIndex)
          arData(lngNewIndex) = lngTemp
      Next
  Next

' ---------------------------------------------------------------------------
' Return the rearranged data
' ---------------------------------------------------------------------------
  Mix_The_Data = arData()

End Function

Public Sub Create_Phrase()

' ***************************************************************************
' Routine:       Create_Phrase
'
' Description:   Build the passphrase and then format the final output
'                display
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim strTable     As String
  Dim strTmp1      As String
  Dim strTmp2      As String
  Dim strTmp3      As String
  Dim lngIndex     As Long
  Dim lngPointer   As Long
  Dim lngNbr2Conv  As Long
  Dim lngLength    As Long
  
' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  strTable = ""
  g_strPassphrase = ""
  strTmp1 = ""
  strTmp2 = ""
  strTmp3 = ""
  
' ---------------------------------------------------------------------------
' Determine which table to read if ALL words are the same length
' ---------------------------------------------------------------------------
  If g_intMaxLength = 0 Then
      Select Case g_intMinLength
             Case 3: strTable = "Char_3"
             Case 4: strTable = "Char_4"
             Case 5: strTable = "Char_5"
             Case 6: strTable = "Char_6"
             Case 7: strTable = "Char_7"
             Case 8: strTable = "Char_8"
      End Select
      
      g_strPassphrase = Build_Using_Same_Length(strTable)
      If Len(Trim(g_strPassphrase)) = 0 Then Exit Sub
  Else
       ' if multple word lengths is selected
      g_strPassphrase = Build_Using_Multiple_Lengths()
      If Len(Trim(g_strPassphrase)) = 0 Then Exit Sub
  End If
   
' ---------------------------------------------------------------------------
' convert the passphrase to propercase
' ---------------------------------------------------------------------------
  Select Case g_intTypeCase
         
         Case 0: ' All lowercase letters
              g_strPassphrase = StrConv(g_strPassphrase, vbLowerCase)
  
         Case 1: ' All Uppercase letters
              g_strPassphrase = StrConv(g_strPassphrase, vbUpperCase)
  
         Case 2: ' All Propercase letters (first character uppercase)
              g_strPassphrase = StrConv(g_strPassphrase, vbProperCase)
  
         Case 3: ' Random mixed case letters
              Erase g_arlngData()                   ' empty the array
              lngLength = Len(g_strPassphrase) - 1  ' get length of string
                
              If lngLength < 5 Then                 ' If string length is less than 5
                  lngNbr2Conv = 2                   '  only 2 chars to be converted
              Else
                  lngNbr2Conv = Int(lngLength / 2)  ' estimate half string length
                  If lngNbr2Conv = 0 Then
                      lngNbr2Conv = 1
                  End If
              End If
                
              ' randomly generate the position in the passphrase string
              ' as to which characters are to be converted to uppercase
              Create_Random_Pointers lngLength, g_arlngData(), True
              
              For lngIndex = 1 To lngNbr2Conv
                  lngPointer = g_arlngData(lngIndex)
                  Mid(g_strPassphrase, lngPointer, 1) = _
                      StrConv(Mid(g_strPassphrase, lngPointer, 1), vbUpperCase)
              Next
  End Select
  
End Sub

Public Function Do_Special_Only(bAlsoNumeric As Boolean, _
                                lngCount As Long) As String

' ***************************************************************************
' Routine:       Do_Special_Only
'
' Description:   Generate an integer that meets the decimal value of a
'                Special keyboard character
'
' Parameters:    bAlsoNumeric - flag to designate if we are to include
'                    numbers along with the special characters
'                lngCount - Quantity requested
'
' Returns:       Special character/numberic string
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intChar  As Integer
  Dim strTmp   As String
  Dim cRND     As clsRndData
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  strTmp = ""
  Set cRND = New clsRndData
  
' ---------------------------------------------------------------------------
' Loop until a specific integer is generated
' ---------------------------------------------------------------------------
  Do
      intChar = Int(cRND.Rnd2(33, 126))
            
      ' see if this is one of the omitted characters
      Check_For_Omitted intChar
         
      ' see if this numeric and special character
      '  or special character only
      If bAlsoNumeric Then
          ' determine the decimal value
          Select Case intChar
                 Case 33, 35 To 38, 40 To 64, 91 To 95, 123 To 126
                      ' append to the output string
                      strTmp = strTmp & Chr(intChar)
                 Case Else
                      intChar = 0
          End Select
      Else
          ' determine the decimal value
          Select Case intChar
                 Case 33, 35 To 38, 40 To 47, 58 To 64, 91 To 95, 123 To 126
                      ' append to the output string
                      strTmp = strTmp & Chr(intChar)
                 Case Else
                      intChar = 0
          End Select
      End If
      
  Loop Until Len(strTmp) = lngCount

' ---------------------------------------------------------------------------
' Return the generated value string
' ---------------------------------------------------------------------------
  Set cRND = Nothing
  Do_Special_Only = strTmp
  
End Function

Public Function ReadFromFile(ByVal sFile_Name As String) As String

' ---------------------------------------------------------------------------
' Read a file into a string. This function will raise an error if the file
' is not found.
' ---------------------------------------------------------------------------
  
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim hFile As Integer

' ---------------------------------------------------------------------------
' Get first free file handle
' ---------------------------------------------------------------------------
  hFile = FreeFile
  
' ---------------------------------------------------------------------------
' Open the text file and dump the contents into a string then close the file.
' ---------------------------------------------------------------------------
  Open sFile_Name For Input As #hFile
  ReadFromFile = Input(LOF(hFile), #hFile)
  Close #hFile
  
End Function

Public Sub Main()

' ---------------------------------------------------------------------------
' Set up the path where all of the mail processing
' will take place.
' ---------------------------------------------------------------------------
  ChDrive App.Path
  ChDir App.Path
      
' ---------------------------------------------------------------------------
' if there is another instance of this program running then leave
' ---------------------------------------------------------------------------
  If App.PrevInstance Then End
  
' ---------------------------------------------------------------------------
' Load the main form
' ---------------------------------------------------------------------------
  Load frmMain
  
End Sub
Public Function WordWrap(sInText As String, lLength As Long) As String

' -----------------------------------------------------------------------------
' This function converts raw text into CRLF delimited lines.
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
' Define local variables
' -----------------------------------------------------------------------------
  Dim sNextLine As String
  Dim sTmpStr As String
  Dim lLen As Long
  Dim iBlank As Integer
  Dim iLineFeed As Integer
  Dim bDoneOnce As Boolean
  
' -----------------------------------------------------------------------------
' Initialize local variables
' -----------------------------------------------------------------------------
  lLength = lLength + 1
  sInText = Trim(sInText)
  bDoneOnce = False
  
' -----------------------------------------------------------------------------
' Loop thru the text and insert the necessary breaks
' -----------------------------------------------------------------------------
Do
    ' get the length of the remaining amount of text
    lLen = Len(sNextLine)
    ' Find the first blank space
    iBlank = InStr(sInText, " ")
    ' find the first linefeed character
    iLineFeed = InStr(sInText, vbLf)

    ' If we found a linefeed character then we
    ' know that we have reached a break point
    ' and will shorten the text string accordingly
    If iLineFeed Then
        If lLen + iLineFeed <= lLength Then
            sTmpStr = sTmpStr & sNextLine & Left(sInText, iLineFeed)
            sNextLine = ""
            sInText = Mid(sInText, iLineFeed + 1)
            GoTo LoopHere
        End If
    End If
    
    ' If we found a blank space
    If iBlank Then
        ' if the length of the text plus the position of the
        ' blank space is less than or equal to  the length
        ' of the text string readjust and move to the next blank
        If lLen + iBlank <= lLength Then
            bDoneOnce = True
            sNextLine = sNextLine & Left(sInText, iBlank)
            sInText = Mid(sInText, iBlank + 1)
        ' if the location of the blank space is greater than
        ' the length of the text string then this is a good
        ' breaking point.
        ElseIf iBlank > lLength Then
            sTmpStr = sTmpStr & vbCrLf & Left(sInText, lLength)
            sInText = Mid(sInText, lLength + 1)
        ' append the rest of the text to the newly formatted
        ' string
        Else
            sTmpStr = sTmpStr & sNextLine & vbCrLf
            sNextLine = ""
        End If
    Else
        ' if the length is not equal zero
        If lLen Then
            If lLen + Len(sInText) > lLength Then
                sTmpStr = sTmpStr & sNextLine & vbCrLf & sInText & vbCrLf
            Else
                sTmpStr = sTmpStr & sNextLine & sInText & vbCrLf
            End If
        Else
            sTmpStr = sTmpStr & sInText & vbCrLf
        End If
        '
        ' We are finished.  time to leave.
        Exit Do
        
    End If

LoopHere:
  Loop

' -----------------------------------------------------------------------------
' Return the newly formatted string
' -----------------------------------------------------------------------------
  WordWrap = sTmpStr
  
End Function
Public Function Check_For_Omitted(intChar As Integer)

' ***************************************************************************
' Routine:       Check_For_Omitted
'
' Description:   Compare an integer value against the list of possible
'                omitted Special keyboard characters
'
' Parameters:    intChar - integer to compare
'
' Return Values: integer
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 31-JAN-2000  Kenneth Ives     Module created kenaso@home.com
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intIndex As Integer
  
' ---------------------------------------------------------------------------
' If the array is empty then leave
' ---------------------------------------------------------------------------
  If Len(g_arOmit(1)) = 0 Then
      Exit Function
  End If
  
' ---------------------------------------------------------------------------
' Loop thru the omit array and look for a match
' ---------------------------------------------------------------------------
  For intIndex = 1 To UBound(g_arOmit)
      If intChar = g_arOmit(intIndex) Then
          intChar = 0
          Exit Function
      End If
  Next
  
End Function

Public Function Sort_Data(lngCount As Long) As Variant

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex      As Long
  Dim intPosition   As Integer
  Dim arTemp()      As String
  Dim strTmp        As String
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  ReDim arTemp(lngCount)
  intPosition = 1
  lngIndex = -1
  strTmp = g_strPassphrase
  
' ---------------------------------------------------------------------------
' See if there are enough elements to sort
' ---------------------------------------------------------------------------
  If lngCount < 2 Then
      arTemp(0) = strTmp
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' Load temp array with passwords
' ---------------------------------------------------------------------------
  Do While intPosition > 0
      intPosition = InStr(1, strTmp, Chr(32))
      
      If intPosition > 0 Then
          lngIndex = lngIndex + 1
          arTemp(lngIndex) = Trim(Left(strTmp, intPosition))
          strTmp = Mid(strTmp, intPosition + 1)
      End If
  Loop
  
' ---------------------------------------------------------------------------
' If more than one number is requested then sort the return array using a
' Sort algorithym in ascending order before returning.
' ---------------------------------------------------------------------------
  
  Select Case lngCount
         Case 1 To 25:         ' do a bubble sort
              For lngIndex = 0 To (lngIndex - 1)
                  If StrComp(arTemp(lngIndex), arTemp(lngIndex + 1)) > -1 Then
                      Swap_Data arTemp(lngIndex), arTemp(lngIndex + 1)
                      lngIndex = -1        ' Reset the index and start over
                  End If
              Next
           
         Case Is > 25:         ' do a quicksort
              lngIndex = lngCount
              QuickSort arTemp(), 0, lngIndex - 1
                
         Case Else:  ' fall thru and do nothing to the array
  End Select
  
Normal_Exit:
' ---------------------------------------------------------------------------
' Return the sorted data
' ---------------------------------------------------------------------------
  Sort_Data = arTemp()
  
End Function

Public Sub QuickSort(arString() As String, lngLow As Long, lngHi As Long)

' ***************************************************************************
' Routine:       QuickSort
'
' Description:   This routine will accept and sort in ascending order a
'                string array of data.  This routine is used when the data
'                to be sorted is known to be ALL string.
'
' Parameters:    arString() - Array to be sorted
'                lngLow     - Minimum number of elements in the array
'                lngHi      - Maximum numer of elements in the array
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 30-APR-2000  Kenneth Ives     Routine created kenaso@home.com
' ***************************************************************************

' ----------------------------------------------------------------------
' Define local striables
' ----------------------------------------------------------------------
  Dim strMidPoint  As String    ' midpoint of the array to be sorted
  Dim strHold      As String    ' Temp hold area for swapping values
  Dim lngTmpLow    As Long      ' Index pointer
  Dim lngTmpHi     As Long      ' Index pointer
   
' ----------------------------------------------------------------------
' See if this is an empty array by checking to see if there is data in
' the first element.  If not, then leave.
' ----------------------------------------------------------------------
  If Len(Trim(arString(0))) = 0 Then
      Exit Sub
  End If
  
' ----------------------------------------------------------------------
' Leave if there is nothing to sort
' ----------------------------------------------------------------------
  If lngLow >= lngHi Then
      Exit Sub
  End If

' ----------------------------------------------------------------------
' Save the count of the minimum and maximum number of elements in the
' array to be sorted.
' ----------------------------------------------------------------------
  lngTmpLow = lngLow
  lngTmpHi = lngHi
   
' ----------------------------------------------------------------------
' Calculate the midpoint of the array
' ----------------------------------------------------------------------
  strMidPoint = arString((lngLow + lngHi) / 2)

' ----------------------------------------------------------------------
' Start the sorting process
' ----------------------------------------------------------------------
  While (lngTmpLow <= lngTmpHi)
       
      ' Always process the low end first.  Loop as long the array data
      ' element is LESS than the data in the temporary holding area
      ' and the temporary low value is LESS than the maximum number of
      ' array elements.  To make an accurate sort for string data, we
      ' temporarily convert the data to Uppercase to make the comparisons.
      ' This is because Uppercase will collate before Lowercase.
      While (StrConv(arString(lngTmpLow), vbUpperCase) < StrConv(strMidPoint, vbUpperCase) And lngTmpLow < lngHi)
          lngTmpLow = lngTmpLow + 1  ' Increment the temp low counter
      Wend
   
      ' Now, we will process the high end.  Loop as long the data in the
      ' temporary holding area is LESS than the array data element
      ' and the temporary high value is GREATER than the minimum number
      ' of array elements.  To make an accurate sort for string data, we
      ' temporarily convert the data to Uppercase to make the comparisons.
      ' This is because Uppercase will collate before Lowercase.
      While (StrConv(strMidPoint, vbUpperCase) < StrConv(arString(lngTmpHi), vbUpperCase) And lngTmpHi > lngLow)
          lngTmpHi = lngTmpHi - 1    ' Decrement the temp high counter
      Wend

      ' if the temp low end is LESS than or equal to the temp high end,
      ' then swap places
      If (lngTmpLow <= lngTmpHi) Then
          Swap_Data arString(lngTmpLow), arString(lngTmpHi)
          lngTmpLow = lngTmpLow + 1                ' Increment the temp low counter
          lngTmpHi = lngTmpHi - 1                  ' Decrement the temp high counter
      End If
 Wend
    
' ----------------------------------------------------------------------
' If the minimum number of elements in the array is LESS than the temp
' high end, then make a recursive call to this routine.  Always sort
' the low end of the array first.  This gives you a solid base.
' ----------------------------------------------------------------------
  If (lngLow < lngTmpHi) Then
      QuickSort arString(), lngLow, lngTmpHi
  End If
   
' ----------------------------------------------------------------------
' If the temp low end is LESS than the maximum number of elements in
' the array, then make a recursive call to this routine.  The high end
' is always sorted last.
' ----------------------------------------------------------------------
  If (lngTmpLow < lngHi) Then
      QuickSort arString(), lngTmpLow, lngHi
  End If

End Sub

Private Function Swap_Data(strValue1 As String, strValue2 As String)

' ***************************************************************************
' Routine:       Swap_Data
'
' Description:   Swap data with each other.  I wrote this function since
'                BASIC stopped having its own SWAP function.
'
' Parameters:    strValue1 - string data to be swapped with strValue2
'                strValue2 - string data to be swapped with strValue1
'
' Return Values: Swapped data
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 15-JUN-2000  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************
    
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strTmp As String
  
' ---------------------------------------------------------------------------
' Swap the values with each other
' ---------------------------------------------------------------------------
  strTmp = strValue1
  strValue1 = strValue2
  strValue2 = strTmp
  
End Function
Public Sub StopTheProgram()

' ---------------------------------------------------------------------------
' Upload all forms from memory and terminate this application
' ---------------------------------------------------------------------------
  Unload_All_Forms
  End
  
End Sub

Public Sub Unload_All_Forms()

' ---------------------------------------------------------------------------
' Unload all forms before terminating an application.  The module that calls
' this routine, usually executes "END" when it returns.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim frm As Form
  
' ---------------------------------------------------------------------------
' For each form associated with this application, we will first hide the
' form, deactivate it, and then remove the form object from memory.  This
' greatly reduces the risk of leaving fragments of a dead application in
' memory.
' ---------------------------------------------------------------------------
  For Each frm In Forms
      frm.Hide
      Unload frm
      Set frm = Nothing
  Next
  
End Sub

