VERSION 5.00
Begin VB.Form ErrorFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Error Dialog Box."
   ClientHeight    =   3495
   ClientLeft      =   1425
   ClientTop       =   1410
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ErrDesTxt 
      Height          =   465
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "ERRORFRM.frx":0000
      Top             =   310
      Width           =   5655
   End
   Begin VB.TextBox ErrRecTxt 
      Height          =   1575
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save data and &End Application."
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton ErrSenCmd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Send Back To Application."
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   4
      X1              =   360
      X2              =   6240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   3
      X1              =   6240
      X2              =   6240
      Y1              =   2640
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   2
      X1              =   360
      X2              =   6240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   360
      X2              =   360
      Y1              =   240
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   360
      X2              =   6240
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "ErrorFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub


Private Sub Command1_Click()
Close All
End
End Sub


Private Sub ErrSenCmd_Click()
ErrorFrm.Hide
End Sub


Private Sub Form_Load()
Screen.MousePointer = 1
CentreForm ErrorFrm, -1
End Sub


