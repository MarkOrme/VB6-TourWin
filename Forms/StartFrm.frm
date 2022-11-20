VERSION 5.00
Begin VB.Form StartFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TourWin Setup Option..."
   ClientHeight    =   4725
   ClientLeft      =   1155
   ClientTop       =   1755
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4725
   ScaleWidth      =   6930
   Begin VB.Frame fraInfsta 
      Height          =   3735
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6615
      Begin VB.Frame fraInfsta 
         Height          =   3735
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6615
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   1335
         Left            =   240
         TabIndex        =   6
         Tag             =   "550"
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdNexsta 
      Caption         =   "&Next"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdBacSta 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton StaCanCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton StaSavCmd 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Menu StaFilmnu 
      Caption         =   "&File"
      Begin VB.Menu StaEximnu 
         Caption         =   "E&xit TourWin"
      End
   End
   Begin VB.Menu StaEdimnu 
      Caption         =   "&Edit"
      Enabled         =   0   'False
   End
   Begin VB.Menu StaDatmnu 
      Caption         =   "&Database"
      Enabled         =   0   'False
   End
   Begin VB.Menu StaToomnu 
      Caption         =   "&Tools"
      Enabled         =   0   'False
   End
   Begin VB.Menu StaSetmnu 
      Caption         =   "&Setup"
      Enabled         =   0   'False
   End
   Begin VB.Menu StaHelmnu 
      Caption         =   "&Help"
      Begin VB.Menu StaConmnu 
         Caption         =   "&Contents"
      End
      Begin VB.Menu StaSeamnu 
         Caption         =   "&Search for Help on ..."
      End
      Begin VB.Menu StaTecmnu 
         Caption         =   "&Technical Support"
      End
      Begin VB.Menu StaAbomnu 
         Caption         =   "&About TourWin cycling Program"
      End
   End
End
Attribute VB_Name = "StartFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cStartFrame = 0
Const cUserFrame = 1
Const cTrainFrame = 2
Const cHeartFrame = 3
Const cEventFrame = 4
Const cFinishFrame = 5


Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
CentreForm StartFrm, -1
' ----------------
' Disable tool bar
' ----------------
MDI!MdiDayBut.Enabled = False
MDI!MdiGraBut.Enabled = False
MDI!MdiEveBut.Enabled = False
MDI!MdiConBut.Enabled = False
MDI!MdiNicBut.Enabled = False
'
' Make frames invisiable

End Sub


Private Sub Form_Unload(Cancel As Integer)
' ------------------------------------------
' This function will only be called if user
' click Finished / Saved button, so load
' Startup values
' ----------------------------------------
    MDI.GetStartUpSetting (0)
    
End Sub

Private Sub StaAbomnu_Click()
Aboutfrm.Show vbModal
End Sub

Private Sub StaCanCmd_Click()
' -------------------------
' Prompt user before ending
' program.
' -------------------------
On Local Error GoTo StaCanCmd_Err
Dim MsgStr As String, MsgRet As Integer
    MsgStr = "Canceling now will end program " & vbLf
    MsgStr = MsgStr & "as setup is not complete. Are " & vbLf
    MsgStr = MsgStr & "you sure you wish to cancel?"
    
    If vbYes = MsgBox(MsgStr, vbYesNo + vbCritical, LoadResString(gcTourVersion)) Then
    
        If "" <> Dir$(App.Path & "\Usertour.mdb") Then Kill App.Path & "\Usertour.mdb"
            End
    End If
Exit Sub
StaCanCmd_Err:
    Err.Clear
    Resume Next
End Sub

Private Sub StaConmnu_Click()
    #If Win32 Then
        Dim ret As Long
    #Else
        Dim ret As Integer
    #End If
    ret = HTMLHelp(MDI.hWnd, App.Path & "\" & HELPFILE, HELP_CONTENTS, CLng(0))
End Sub


Private Sub StaEximnu_Click()
StaCanCmd_Click
End Sub


Private Sub StaSavCmd_Click()
' -------------------------
' Check that all is okay...
' -------------------------
ObjNew.info.NewUser = False
' ----------------
' Disable tool bar
' ----------------
MDI!MdiDayBut.Enabled = True
MDI!MdiGraBut.Enabled = True
MDI!MdiEveBut.Enabled = True
MDI!MdiConBut.Enabled = True
MDI!MdiNicBut.Enabled = True
Unload StartFrm
End Sub


Private Sub StaSeamnu_Click()
    #If Win32 Then
        Dim ret As Long
    #Else
        Dim ret As Integer
    #End If
    ret = HTMLHelp(MDI.hWnd, App.Path & "\" & HELPFILE, HELP_CONTENTS, CLng(0))
End Sub

Private Sub StaTecmnu_Click()
TechSupport
End Sub


