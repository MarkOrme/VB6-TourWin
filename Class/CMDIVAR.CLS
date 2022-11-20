VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMDIVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' --------------------------------------------
' Global class variable for user definiation
' like form options, passwords, user name etc.
' --------------------------------------------
Const cME = 2
Public Name                 As String
Public ID                   As Integer
Public Password             As String
Public Datapath             As String
Public Security             As Boolean
Public Load                 As Integer
Public MetaFile             As String
Public ShowMeta             As Boolean
Public iInterations         As Integer ' Used by create form to determine number of interation
                               ' to create.
Public iWindowState         As Integer
Public iInterationLen       As Integer
Public dInterationDate      As Date
Public bStartOf             As Boolean  ' If the Peak schedule starts with selected date
                            ' as apposide to ending on specified date.
Public NewUser              As Boolean
Public RunOnce              As Boolean
Public sCurrentActiveWindow As String 'set to the current active child window.
Public UserOptions         As cBitReprentation   ' Stores several settings...
Public cSHORTDATEFORMAT    As String

' iBitFlag allocations
' 0000 0000 0000 0000
' ^^^^ ^^^^ ^^^^ ^^^^
' |||| |||| |||| |||| Save Contact ColumnWidth = 0
' |||| |||| |||| |||- Load Wizard             = "BitField" And 1
' |||| |||| |||| ||-- Conconi - Plot Points   = "BitField" And 2
' |||| |||| |||| |--- Conconi - Draw Legend   = "BitField" And 4
' |||| |||| |||| ---- Conconi - Draw AT       = "BitField" And 8
' |||| |||| |||------ Conconi - Load Settings = "BitField" And 16
' |||| |||| ||------- User    - Security      - 0 no  , 1 yes
' |||| |||| |-------- User    - Meta          - 0 hide, 1 show
' |||| |||| --------- User    - Load1         - 0 hide, 1 show
' |||| |||----------- User    - Load2         - 0 hide, 1 show
' |||| ||------------ User    - Load3         - 0 hide, 1 show
'


Public DailyOpt            As cBitReprentation
' Bit mapping for lDailyOpt
' 0000 0000 0000 0000
' ^^^^ ^^^^ ^^^^ ^^^^
' |||| |||| |||| ||||
' |||| |||| |||| |||--- Resting Heart Rate  - 0 hide, 1 show
' |||| |||| |||| ||---- Day Type            - 0 hide, 1 show
' |||| |||| |||| |----- Weight              - 0 hide, 1 show
' |||| |||| |||| ------ Sleep               - 0 hide, 1 show
' |||| |||| |||-------- Description         - 0 hide, 1 show
' |||| |||| ||--------- Metric              - 0 hide, 1 show




Public eCheckDBType         As eCheckDatabase

'local variable(s) to hold property value(s)
Private mvarWelcomeWizard As Boolean 'local copy
Public Property Let WelcomeWizard(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WelcomeWizard = 5
    mvarWelcomeWizard = vData
    
End Property


Public Property Get WelcomeWizard() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WelcomeWizard
    WelcomeWizard = mvarWelcomeWizard
    
End Property

Private Sub Class_Initialize()

Dim sShortDate As String
Dim sNewDate As String
Dim sBuffer As String * 100
Dim dl As Long

On Local Error Resume Next

Set UserOptions = New cBitReprentation
Set DailyOpt = New cBitReprentation
    
dl = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer, 99)
cSHORTDATEFORMAT = Mid$(sBuffer, 1, dl - 1)

End Sub

Private Sub Class_Terminate()
    Set UserOptions = Nothing
    Set DailyOpt = Nothing
    
End Sub
