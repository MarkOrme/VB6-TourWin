VERSION 5.00
Begin VB.Form frmActivity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "####"
   ClientHeight    =   3885
   ClientLeft      =   3510
   ClientTop       =   1620
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmActivity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "914"
   Begin VB.CommandButton cmdModify 
      Caption         =   "####"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Tag             =   "906"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "####"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Tag             =   "904"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame fraActivities 
      Caption         =   "####"
      Height          =   3545
      Left            =   120
      TabIndex        =   5
      Tag             =   "902"
      Top             =   240
      Width           =   3375
      Begin VB.ListBox lstItems 
         Height          =   3180
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblDes 
         Caption         =   "####"
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
         Left            =   240
         TabIndex        =   6
         Tag             =   "905"
         Top             =   -240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "####"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Tag             =   "901"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "####"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Tag             =   "900"
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lType_Id As Long

Private Sub cmdAdd_Click()
Dim lRt As Long
' --------------------------------------------------
' Set cActivityNames properties to ""
' so frmItems will no its an Add action
' --------------------------------------------------

' Set up dialog before showing
' Heart Names has some restrictions
' -----------------------------------
If cActivityNames.Type_ID = gcActive_Type_HeartNames Then
    If lstItems.ListCount >= 9 Then
        MsgBox "Maximum number if items is 9. Either modify or delete existing item(s).", vbOKOnly, LoadResString(gcTourVersion)
        cmdAdd.Enabled = False
        Exit Sub
    End If
    frmItems.MaxDescLength = 15
    frmItems.ColourButtonEnabled = False
Else
    
    frmItems.MaxDescLength = 100
    frmItems.ColourButtonEnabled = True
    
End If

frmItems.Show vbModal

If Not frmItems.USERCANCELLED Then

cActivityNames.Description = frmItems.txtDescription.Text
cActivityNames.Colour = frmItems.txtDescription.BackColor

If cActivityNames.Add(lType_Id, cActivityNames.Description, cActivityNames.Colour) Then
    lstItems.AddItem cActivityNames.Description
    
    ' Highlight newly added item
    lRt = FindItemListControl(lstItems, cActivityNames.Description)
    If -1 = lRt Then
        lstItems.ListIndex = lstItems.ListCount - 1
    Else
        lstItems.ListIndex = lRt
    End If
    
    ' Enable Delete and Modify button
    If cmdDelete.Enabled = False Then
        cmdDelete.Enabled = True
        cmdModify.Enabled = True
    End If
    If gcActive_Type_PeakNames = lType_Id Then
        cPeakNames.IsUpdated = False    ' Set flag to false, force refresh...
    End If
Else
    MsgBox "Add Item Failed"
End If
End If
Unload frmItems

End Sub

Private Sub CmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdDelete_Click()
Dim iListIndexPos As Integer

If lstItems.ListIndex = -1 Then
    MsgBox LoadResString(SELECT_ITEM_FROM_LIST), vbOKOnly + vbInformation, LoadResString(gcTourVersion)
Else
    If vbYes = MsgBox(LoadResString(gcWantToDelete) & " '" & lstItems.Text & "' ", vbYesNo + vbCritical, LoadResString(gcTourVersion)) Then
        iListIndexPos = lstItems.ListIndex
        If cActivityNames.Delete(lType_Id, lstItems.Text) Then
            lstItems.RemoveItem lstItems.ListIndex
            ' Set focus to the item that now
            ' takes the place of ListIndex
            On Error Resume Next
            If (lstItems.ListCount - 1) >= iListIndexPos Then
                lstItems.ListIndex = iListIndexPos
            Else

                If lstItems.ListCount = 1 Then
                    lstItems.ListIndex = 0
                Else
                    If lstItems.ListCount = 0 Then
                        cmdDelete.Enabled = False
                        cmdModify.Enabled = False
                    End If
                End If
            End If
            cmdAdd.Enabled = True
            lstItems.SetFocus
            
            ' Refresh cPeakNames collection...
            If gcActive_Type_PeakNames = lType_Id Then
                cPeakNames.IsUpdated = False    ' Set flag to false, force refresh...
            End If

        End If
    End If
End If
End Sub

Private Sub cmdModify_Click()

If lstItems.ListIndex = -1 Then
        MsgBox LoadResString(SELECT_ITEM_FROM_LIST), vbOKOnly + vbInformation, LoadResString(gcTourVersion)
        Exit Sub
Else
' --------------------------------------------------
' Set cActivityNames properties to ""
' so frmItems will no its an Add action
' --------------------------------------------------

cActivityNames.Description = lstItems.Text

' Get Modify Description
cActivityNames.FindItemByName cActivityNames.Description
frmItems.txtDescription.Text = cActivityNames.Description


If cActivityNames.Description = "" Then Exit Sub
    If cActivityNames.Type_ID = gcActive_Type_HeartNames Then
        frmItems.MaxDescLength = 15
        frmItems.ColourButtonEnabled = False
        frmItems.txtDescription.BackColor = &H80000005
    Else
        frmItems.MaxDescLength = 100
        frmItems.ColourButtonEnabled = True
        
        If cActivityNames.Colour <> "" Then
            frmItems.txtDescription.BackColor = cActivityNames.Colour
        End If
    End If
End If


frmItems.Show vbModal
    
If Not frmItems.USERCANCELLED Then
    cActivityNames.Description = frmItems.txtDescription.Text
    cActivityNames.Colour = frmItems.txtDescription.BackColor
    
        If cActivityNames.Modify(lType_Id, cActivityNames.Description, cActivityNames.Colour) Then
            ' ---------------------------
            ' Remove current item
            ' and add update to exist
            ' position
            ' ---------------------------

            Dim lIndex As Long
                lIndex = lstItems.ListIndex
                lstItems.RemoveItem lIndex
                lstItems.AddItem cActivityNames.Description, lIndex
                lstItems.ListIndex = lIndex
                
                ' Refresh cPeakNames collection...
                If gcActive_Type_PeakNames = lType_Id Then
                    cPeakNames.IsUpdated = False    ' Set flag to false, force refresh...
                End If
        End If
    End If
    
    Unload frmItems


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If vbKeyEscape = KeyAscii Then
    Me.Hide
End If
End Sub

Private Sub Form_Load()

Me.KeyPreview = True
CentreForm Me, 0
' Set horizontal scrollbar to allow viewing of entry...
SendMessage lstItems.hWnd, LB_SETHORIZONTALEXTENT, 550, 0

lType_Id = cActivityNames.Type_ID

' Load Generic Resources
LoadFormResourceString Me

' ---------------------------
' Based on lType_ID, Define
' Form and frame captions
' ---------------------------
Select Case lType_Id
       Case gcActive_Type_PeakNames:
            Me.Caption = LoadResString(910)
       Case gcActive_Type_EventNames:
            Me.Caption = LoadResString(911)
       Case gcActive_Type_HeartNames:
            Me.Caption = LoadResString(912)
End Select

' -------------------------------------------------
' Based on Record count show appropriate controls
' and load the first 10 descriptions.
' -------------------------------------------------
If cActivityNames.StartSearch(lType_Id) Then
    Do
        lstItems.AddItem cActivityNames.Description
        
    Loop While cActivityNames.GetNext
End If

If lstItems.ListCount = 0 Then
    cmdDelete.Enabled = False
    cmdModify.Enabled = False
Else
    lstItems.ListIndex = 0
End If
End Sub


Private Sub lstItems_DblClick()
    cmdModify_Click
End Sub
