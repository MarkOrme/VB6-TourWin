VERSION 5.00
Begin VB.Form ContOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contact database options."
   ClientHeight    =   2355
   ClientLeft      =   2835
   ClientTop       =   1890
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Contopt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "851"
   Begin VB.CheckBox chkColumnwidth 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Tag             =   "864"
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton CoOCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Tag             =   "855"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sort type"
      Height          =   975
      Left            =   3360
      TabIndex        =   12
      Tag             =   "853"
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton CoODesOpt 
         Caption         =   "Descen&ding"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "863"
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton CoOAscOpt 
         Caption         =   "Ascendin&g"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "862"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton CoOSavCmd 
      Caption         =   "S&ave"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Tag             =   "854"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Active Index."
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Tag             =   "852"
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton CoOEmaOpt 
         Caption         =   "E&mail Address."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Tag             =   "861"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton CoOFaxOpt 
         Caption         =   "&Fax Number."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Tag             =   "860"
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton CoOPhoOpt 
         Caption         =   "&Phone Number."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   1
         Tag             =   "859"
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton CoOFirOpt 
         Caption         =   "Firs&t Name."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Tag             =   "858"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton CoOLasOpt 
         Caption         =   "Last &Name."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Tag             =   "857"
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton CoOConOpt 
         Caption         =   "C&ontact."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Tag             =   "856"
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "ContOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Load_Settings(lLoadOptions As Long)
' ---------------------------------------------------------
' Update TourUser file with DaiOpt table information
' On Local Error GoTo Check_Err
' ---------------------------------------------------------
Const LOADOPTIONS = 0
Const UPDATEOPTIONS = 1

Dim oContactOptions     As CRecord
Dim SQL                 As String


If bDebug Then
    Handle_Err 0, "Load_Settings-ContOpt"
End If

' Create Record object
Set oContactOptions = New CRecord
If oContactOptions Is Nothing Then
    Unload Me
End If


If lLoadOptions = UPDATEOPTIONS Then

    With oContactOptions
    
        .RstSQL "SELECT * FROM " & gcUserTour_ContactOpt & " WHERE ID = " & objMdi.info.ID
                            
        'Verify that record exist for current user
       If .RstRecordCount > 0 Then
          .Edit
       Else
        .AddNew
    
        .SetField gcUserTour_ContactOpt_Id, objMdi.info.ID
        .SetField gcUserTour_ContactOpt_IndexOrder, 1020
        .SetField gcUserTour_ContactOpt_SortOrder, 1020
        .SetField gcUserTour_ContactOpt_ContactWidth, 1020
        .SetField gcUserTour_ContactOpt_LastWidth, 1020
        .SetField gcUserTour_ContactOpt_FirstWidth, 1020
        .SetField gcUserTour_ContactOpt_PhoneWidth, 1020
        .SetField gcUserTour_ContactOpt_FaxWidth, 1020
        .SetField gcUserTour_ContactOpt_E_MailWidth, 1020
       
   
                ' The SaveColumnWidth is saved with the user record (BitField field)
            '    If objMdi.info.UserOptions.GetBool(BitFlags.Pos_0) Then
            '        objMdi.info.iBitFlag = objMdi.info.iBitFlag + 1
            '        objMdi.SaveUserSettings
            '    End If
      End If

   End With




ProgressBar LoadResString(gcSavingdata), -1, 5, -1

If CoOConOpt.Value = -1 Then             'Sort By Contact Field
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Contact
        oContactOptions.SetField "IndexOrder", gcContTour_Contacts_Contact
ElseIf CoOLasOpt.Value = -1 Then         'Sort by Last Name Field
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Last
        oContactOptions.SetField "IndexOrder", gcContTour_Contacts_Last
ElseIf CoOFirOpt.Value = -1 Then         'Sort by First Name Field
        ObjCont.info.IndxOrdr = gcContTour_Contacts_First
        oContactOptions.SetField "IndexOrder", gcContTour_Contacts_First
ElseIf CoOPhoOpt.Value = -1 Then         'Sort by Phone Name Field
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Phone
        oContactOptions.SetField "IndexOrder", gcContTour_Contacts_Phone
ElseIf CoOFaxOpt.Value = -1 Then         'Sort by Fax Name Field
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Fax
        oContactOptions.SetField "IndexOrder", gcContTour_Contacts_Fax
ElseIf CoOEmaOpt.Value = -1 Then         'Sort by E-Mail Name Field
        ObjCont.info.IndxOrdr = gcContTour_Contacts_Email
        oContactOptions.SetField "IndexOrder", gcContTour_Contacts_Email
End If
    ProgressBar LoadResString(gcSavingdata), -1, 7, -1
' -------------------------
' Update Desc / Asc setting
' -------------------------
If CoOAscOpt.Value = -1 Then
    ObjCont.info.SrtOrdr = "ASC"
    oContactOptions.SetField "SortOrder", "ASC"
Else
    ObjCont.info.SrtOrdr = "DESC"
    oContactOptions.SetField "SortOrder", "DESC"
End If

' Write Save Width to User Object (ObjMdi)
objMdi.info.UserOptions.SetValue ContOpt.chkColumnwidth.Value, BitFlags.Pos_0
objMdi.SaveUserSettings


'Update Database
oContactOptions.Update

Set oContactOptions = Nothing
' ----------------------------
' Load ConSorLst With list of
' records for currently sorted
' field.
' ----------------------------
ContFrm.LoadListOfSortedField

ProgressBar LoadResString(gcSavingdata), -1, 10, -1
ProgressBar "", 0, 0, 0
' Update Form with current
' object values

Else
'Active Index Option
    Select Case ObjCont.info.IndxOrdr
            Case gcContTour_Contacts_Contact:
                CoOConOpt.Value = -1
            Case gcContTour_Contacts_Last:
                CoOLasOpt.Value = -1
            Case gcContTour_Contacts_First:
                CoOFirOpt.Value = -1
            Case gcContTour_Contacts_Phone:
                CoOPhoOpt.Value = -1
            Case gcContTour_Contacts_Fax:
                CoOFaxOpt.Value = -1
            Case gcContTour_Contacts_Email:
                CoOEmaOpt.Value = -1
    End Select
ProgressBar LoadResString(gcLoadingData), -1, 5, -1
    
    'Sort Order Option
    If Trim$(ObjCont.info.SrtOrdr) = "ASC" Then
        CoOAscOpt.Value = 1
        CoODesOpt.Value = 0
    Else
        CoOAscOpt.Value = 0
        CoODesOpt.Value = 1
    End If
    
    chkColumnwidth.Value = objMdi.info.UserOptions.GetValue(BitFlags.Contact_SaveColumnWidth)
    
End If
ProgressBar LoadResString(gcLoadingData), -1, 10, -1
ProgressBar "", 0, 0, 0
End Sub

Private Sub CoOCanCmd_Click()
Unload ContOpt
UserCancel = True
End Sub


Private Sub CoOSavCmd_Click()
Const LOADOPTIONS = 0
Const UPDATEOPTIONS = 1

ProgressBar LoadResString(gcSavingdata), -1, 1, -1
    Load_Settings UPDATEOPTIONS
    Unload Me
UserCancel = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Const LOADOPTIONS = 0
Const UPDATEOPTIONS = 1

    ProgressBar LoadResString(gcLoadingData), -1, 1, -1
    LoadFormResourceString ContOpt
    Me.KeyPreview = True
    CentreForm ContOpt, -1
        
        Load_Settings LOADOPTIONS
        
UserCancel = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
UserCancel = True
End Sub


