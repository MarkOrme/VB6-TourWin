VERSION 5.00
Begin VB.Form frmCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   1830
   ClientLeft      =   -45
   ClientTop       =   255
   ClientWidth     =   5910
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "copy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5910
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      FillColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   168
      ScaleHeight     =   330
      ScaleWidth      =   5535
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   708
      Width           =   5592
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "#"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2085
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   1275
      Width           =   1665
   End
   Begin VB.Label lblDestFile 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   168
      TabIndex        =   1
      Top             =   300
      Width           =   5640
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   0
      Width           =   120
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' frmCopy
Private Sub cmdExit_Click()
    ExitSetup Me, gintRET_EXIT
End Sub

Private Sub Form_Load()
Dim lHeight     As Long
Dim lWidth      As Long

SetFormFont Me
cmdExit.Caption = ResolveResString(resBTNCANCEL)
lblCopy.Caption = ResolveResString(resLBLDESTFILE)
lblDestFile.Caption = gstrNULL
    
frmCopy.Caption = gstrTitle

fGetTaskBarDimensions lHeight, lWidth
' Set form in button right hand corner
frmCopy.Top = Screen.Height - (frmCopy.Height + lHeight + 500)
frmCopy.Left = Screen.Width - frmCopy.Width - 10
        
            
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        ExitSetup Me, gintRET_EXIT
        Cancel = 1
    End If
End Sub

