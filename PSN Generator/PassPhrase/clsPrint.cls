VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------
' Do not remember where I found this piece of code.  If it is yours, let me
' know.  I will place your name and URL here.  Thank you.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Margin values
' ---------------------------------------------------------------------------
  Private Const MARGIN_LEFT = 1440
  Private Const MARGIN_RIGHT = 1440
  Private Const MARGIN_TOP = 1440
  Private Const MARGIN_BOTTOM = 1440

Private Sub DoNewPage(sTitleLine As String, bEjectPage As Boolean)
    
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim sBuffer As String

' ---------------------------------------------------------------------------
' Start new page if requested
' ---------------------------------------------------------------------------
  If bEjectPage Then Printer.NewPage

' ---------------------------------------------------------------------------
' Define printer font for header/footer
' ---------------------------------------------------------------------------
  Printer.Print
  Printer.Font = "Times New Roman"
  Printer.FontSize = 12
  Printer.FontBold = True

' ---------------------------------------------------------------------------
' Print page header
' ---------------------------------------------------------------------------
  Printer.CurrentY = (MARGIN_TOP - Printer.TextHeight(sTitleLine)) / 2
  Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(sTitleLine)) / 2
  Printer.Print sTitleLine;

' ---------------------------------------------------------------------------
' Print page footer
' ---------------------------------------------------------------------------
  sBuffer = StrConv(Format(Now(), "d mmm yyyy  hh:mm:ss"), vbProperCase)
  Printer.CurrentX = MARGIN_LEFT
  Printer.CurrentY = Printer.ScaleHeight - (MARGIN_BOTTOM / 2)
  Printer.Print sBuffer;

  sBuffer = "Page  " & Printer.Page
  Printer.CurrentX = Printer.ScaleWidth - (Printer.TextWidth(sBuffer) + MARGIN_RIGHT)
  Printer.CurrentY = Printer.ScaleHeight - (MARGIN_BOTTOM / 2)
  Printer.Print sBuffer;

' ---------------------------------------------------------------------------
' Reset position to top of page and redefine font for page body print.
' ---------------------------------------------------------------------------
  Printer.Print
  Printer.Font = "Courier New"
  Printer.FontSize = 10
  Printer.FontBold = True
  Printer.CurrentX = MARGIN_LEFT
  Printer.CurrentY = MARGIN_TOP

End Sub

Public Sub PrintText(sTitleLine As String, sDataToPrint As String)

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim intPosition   As Integer   ' index counter
  Dim intMaxLength  As Integer   ' max output record length
  Dim strPrintLine  As String    ' output print line

' ---------------------------------------------------------------------------
' Initialize first page
' ---------------------------------------------------------------------------
  Call DoNewPage(sTitleLine, False)         ' print page header/footer
  sDataToPrint = sDataToPrint & Chr(32)     ' add 1 trailing blank
  intMaxLength = 78
  
' ---------------------------------------------------------------------------
' Build the output data prior to dumping to the printer
' ---------------------------------------------------------------------------
  Do

      ' see if there is anything left to print
      If Len(Trim(sDataToPrint)) = 0 Then
          Exit Do
      End If
      
      'Get next string of data
      strPrintLine = Left(sDataToPrint, intMaxLength)
              
      ' verify we have something to print
      If Len(Trim(strPrintLine)) = 0 Then
          Exit Do
      End If
      
      intPosition = InStrRev(strPrintLine, Chr(32), Len(strPrintLine))
      strPrintLine = Trim(Left(strPrintLine, intPosition))
      
      'Check if word will fit on this line
      If Len(strPrintLine) <= intMaxLength Then
          
          ' Resize input data string
          sDataToPrint = Trim(Mid(sDataToPrint, intPosition)) & Chr(32)
          
          Printer.Print     'Send carriage-return line-feed to printer
          
          'Check if we need to start a new page
          If Printer.CurrentY > (Printer.ScaleHeight - MARGIN_BOTTOM) Then
              Call DoNewPage(sTitleLine, True) ' print page header/footer
          Else
              Printer.CurrentX = MARGIN_LEFT   ' set position to left margin
          End If
                  
          Printer.Print strPrintLine;                   ' Print this line on the page
          
          
      ElseIf Len(strPrintLine) = 0 Then
          Exit Do
      End If
  Loop

' ---------------------------------------------------------------------------
' Notify printer we are finished
' ---------------------------------------------------------------------------
  Printer.EndDoc

End Sub

