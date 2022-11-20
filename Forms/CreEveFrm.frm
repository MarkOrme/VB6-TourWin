VERSION 5.00
Begin VB.Form CreEveFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3180
   ClientLeft      =   2145
   ClientTop       =   2055
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CreEveFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Contact Information. "
      Height          =   1695
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   7095
      Begin VB.TextBox CevEmaTxt 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox CevFaxTxt 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox CevPhoTxt 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox CevFirTxt 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox CevLasTxt 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox CevConCbo 
         Height          =   315
         ItemData        =   "CreEveFrm.frx":000C
         Left            =   1080
         List            =   "CreEveFrm.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "E-&Mail:"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Fa&x:"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "&Phone:"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "&First Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "&Last Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "C&ontact:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Event Information."
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox CevLisCbo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox CevDesTxt 
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "&Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "&Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CevSavCmd 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CevCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "CreEveFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lEventHandle            As Long
Dim lContHandle             As Long

Dim EventColor(0 To 100)    As String
Dim RecordExist             As Boolean


Function Check_CreEve_Input() As Boolean
' -----------------------------
' Checks that certain Text boxs
' are filled in by user
' -----------------------------
On Local Error GoTo CreEve_Input_Err
Check_CreEve_Input = True
If Trim$(CevDesTxt.Text) = "" Then
    Beep
    MsgBox "Blank 'Event Description' not allowed!", vbOKOnly + vbCritical, LoadResString(gcTourVersion)
    Check_CreEve_Input = False
    CevDesTxt.SetFocus
    Exit Function
End If
Exit Function
CreEve_Input_Err:
    If bDebug Then Handle_Err Err, "Check_CreEve_Input-CreEveFrm"
    Resume Next
End Function

Sub setCont_To_Data(ContName As String)

On Local Error GoTo setCont_To_Err

Dim SearchField As String

' Search for desired contact.
SearchField = "Contact = '" & ContName & "'"
ObjTour.DBFindFirst SearchField, lContHandle

' Check if contact is found
If Not ObjTour.EOF(lContHandle) Then

    CevLasTxt.Text = ObjTour.DBGetField(gcContTour_Contacts_Last, lContHandle) 'setCont("Last")
    CevFirTxt.Text = ObjTour.DBGetField(gcContTour_Contacts_First, lContHandle)
    CevPhoTxt.Text = ObjTour.DBGetField(gcContTour_Contacts_Phone, lContHandle)
    CevFaxTxt.Text = IIf("" = ObjTour.DBGetField(gcContTour_Contacts_Fax, lContHandle), " ", ObjTour.DBGetField("Fax", lContHandle)) ' Error
    CevEmaTxt.Text = IIf("" = ObjTour.DBGetField(gcContTour_Contacts_Email, lContHandle), " ", ObjTour.DBGetField("E-Mail", lContHandle))
    
End If

Exit Sub
setCont_To_Err:
    If bDebug Then Handle_Err Err, "setCont_To_Data-CreEveFrm"
    Resume Next
End Sub

Private Sub CevCanCmd_Click()

Unload CreEveFrm

End Sub
Private Sub CevConCbo_Click()
        setCont_To_Data CevConCbo.Text
End Sub

Private Sub CevConCbo_LostFocus()
setCont_To_Data CevConCbo.Text
End Sub


Private Sub CevDesTxt_GotFocus()
With CevDesTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub CevDesTxt_KeyPress(KeyAscii As Integer)
If Len(Trim$(CevDesTxt)) >= 100 Then
    MsgBox "Max length of 100 characters reached", vbOKOnly, LoadResString(gcTourVersion)
    KeyAscii = 0
End If

End Sub
Private Sub CevEmaTxt_GotFocus()

With CevEmaTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub CevFaxTxt_GotFocus()

With CevFaxTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub CevFirTxt_GotFocus()

With CevFirTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub CevLasTxt_GotFocus()

With CevLasTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub
Private Sub CevPhoTxt_GotFocus()

With CevPhoTxt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub CevSavCmd_Click()
On Local Error GoTo CevSavCmd_Err
Dim RetBoo As Boolean
' ------------------------
' Check invalid user input
' ------------------------
If Not Check_CreEve_Input Then Exit Sub

If RecordExist Then
    ObjTour.Edit lEventHandle
Else
    ObjTour.AddNew lEventHandle
    
    ObjTour.DBSetField gcEve_Tour_Id, objMdi.info.ID, lEventHandle
    ObjTour.DBSetField gcEve_Tour_Date, objEve.info.EventDate, lEventHandle
    If CevLisCbo.ListIndex >= 0 Then
    ObjTour.DBSetField gcEve_Tour_Color, CLng(EventColor(CevLisCbo.ListIndex)), lEventHandle
    End If
    ObjTour.DBSetField gcEve_Tour_Page, Val(Calndfrm!CalPagMsk.Text), lEventHandle
    
End If

    ObjTour.DBSetField gcEve_Tour_Evememo, CevDesTxt.Text, lEventHandle
    ObjTour.DBSetField gcEve_Tour_EveType, CevLisCbo.Text, lEventHandle
    ObjTour.DBSetField gcEve_Tour_PointToCont, ObjTour.DBGetField("Count", lContHandle), lEventHandle
    
    ' Make sure the above line is correct...
    ObjTour.Update lEventHandle
    
    
   Calndfrm.RestCalendar
   
   Unload CreEveFrm ' Release DB objects is found in Unload function
   
Exit Sub

CevSavCmd_Err:
    If bDebug Then Handle_Err Err, "CevSavCmd-CreEveFrm"
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo CreEve_Err


Dim SQL As String, SearchStr As String, i As Integer
Dim TypeNum As String, TypeCol As String, RetStr As String


RecordExist = False
CentreForm CreEveFrm, 0
Me.KeyPreview = True

SQL = "SELECT * FROM " & gcContTour_Contacts & " WHERE Id = " & objMdi.info.ID

ObjTour.RstSQL lContHandle, SQL


If ObjTour.RstRecordCount(lContHandle) <> 0 Then
    ObjTour.DBMoveFirst lContHandle

    ' ----------------------
    ' Load Contact Name into
    ' CevConCbo list box....
    ' ----------------------
    While Not ObjTour.EOF(lContHandle)  'setCont.EOF
        CevConCbo.AddItem ObjTour.DBGetField(gcContTour_Contacts_Contact, lContHandle) 'setCont("Contact")
        ObjTour.DBMoveNext lContHandle

    Wend
    
    CevConCbo.ListIndex = 0
    
    setCont_To_Data CevConCbo.Text
    
End If

' =====================
' Define Record Set
' =====================
SQL = "SELECT * FROM " & gcNameTour_Events & " WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

' Load list box
CevLisCbo.Clear
cActivityNames.Type_ID = gcActive_Type_EventNames
' -------------------------------------------------
' Based on Record count show appropriate controls
' and load the first 10 descriptions.
' -------------------------------------------------
If cActivityNames.StartSearch(gcActive_Type_EventNames) Then
    i = 0
    Do
        ProgressBar "Loading Type Names...", -1, i * 1.11, -1
        CevLisCbo.AddItem cActivityNames.Description
        EventColor(i) = cActivityNames.Colour
        i = i + 1
    Loop While cActivityNames.GetNext
End If


     ProgressBar "", 0, 0, 0

' --------------------
' Check to find if
' Event exist for date
' --------------------

SQL = "SELECT * FROM " & gcEve_Tour_Event & " WHERE Id = " & objMdi.info.ID & " and Date = #" & objEve.info.EventDate & "# AND Page = " & Val(Calndfrm!CalPagMsk.Text)

ObjTour.RstSQL lEventHandle, SQL

' Check if records exist.
If ObjTour.RstRecordCount(lEventHandle) <> 0 Then
    
    CevDesTxt = ObjTour.DBGetField(gcEve_Tour_Evememo, lEventHandle)
    CevLisCbo.Text = ObjTour.DBGetField(gcEve_Tour_EveType, lEventHandle)
    SearchStr = "Count = " & ObjTour.DBGetField(gcEve_Tour_PointToCont, lEventHandle)
    
    ObjTour.DBFindFirst SearchStr, lContHandle
    'setCont.FindFirst SearchStr
    
        If Not ObjTour.EOF(lContHandle) Then
            RecordExist = True
            
            CevConCbo.Text = ObjTour.DBGetField(gcEve_Tour_Contact, lContHandle)
            setCont_To_Data ObjTour.DBGetField(gcEve_Tour_Contact, lContHandle)
        End If
End If
'
' Load Events
     Exit Sub
CreEve_Err:
    If bDebug Then Handle_Err Err, "Form_Load-CreEveFrm"
    Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)

ObjTour.DBClose lContHandle
ObjTour.FreeHandle lContHandle
lContHandle = 0

ObjTour.DBClose lEventHandle
ObjTour.FreeHandle lEventHandle
lEventHandle = 0

End Sub
