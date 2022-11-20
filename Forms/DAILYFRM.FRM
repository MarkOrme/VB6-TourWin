VERSION 5.00
Begin VB.Form DailyFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Diary"
   ClientHeight    =   5040
   ClientLeft      =   2145
   ClientTop       =   1695
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   4
   Icon            =   "DAILYFRM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5040
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton DaiNexCmd 
      Caption         =   "&Next Record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   28
      Top             =   4620
      Width           =   1455
   End
   Begin VB.CommandButton DaiPreCmd 
      Caption         =   "&Previous Record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   4620
      Width           =   1455
   End
   Begin VB.Frame DaiButFra 
      Caption         =   "    Daily Note and Historical Significates     "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   44
      Top             =   1920
      Width           =   3615
      Begin VB.CommandButton cmdHistorical 
         Caption         =   "Historical Signifi&cates"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton DaiMemCmd 
         Caption         =   "Edit &Memo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "    Daily Ratings and Measurements    "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   42
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdDate 
         Height          =   255
         Left            =   1200
         Picture         =   "DAILYFRM.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Date Search"
         Top             =   480
         Width           =   255
      End
      Begin Tourwin2002.UTextBox DaiWeiTxt 
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         ToolTipText     =   "Body weight (range 0 to 999.99)"
         Top             =   1320
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         Max             =   6
         FieldType       =   1
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin Tourwin2002.UTextBox DaiRatTxt 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         ToolTipText     =   "Overall health and energy rating (range 0 to 9)"
         Top             =   1320
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         Max             =   1
         FieldType       =   1
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin Tourwin2002.UTextBox DaiHeaTxt 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Tag             =   "gcDai_Tour_Heart"
         ToolTipText     =   "Resting heart rate (000)"
         Top             =   840
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         Max             =   3
         FieldType       =   1
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin VB.ComboBox DaiDayCbo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   4
         Tag             =   "gcDai_Tour_DayType"
         Text            =   "DaiDayCbo"
         Top             =   480
         Width           =   1695
      End
      Begin Tourwin2002.UTextBox DaiSlpTxt 
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         ToolTipText     =   "Amount of rest (00:00:00)"
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Max             =   8
         FieldType       =   3
         Text            =   "00:00:00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   "Sleep"
      End
      Begin Tourwin2002.UTextBox DaiDatTxt 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Date (mm-dd-yyyy)"
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin VB.Label DaiWeiLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Weight:"
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
         Left            =   1950
         TabIndex        =   11
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label DaiRatLbl 
         AutoSize        =   -1  'True
         Caption         =   "Day &Rate:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label DaiSlpLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "&Sleep:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1905
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label DaiHeaLbl 
         Caption         =   "Rest&ing Heart: Rate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label DaiDaiLbl 
         AutoSize        =   -1  'True
         Caption         =   "Day T&ype:"
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
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   750
      End
      Begin VB.Label DaiDatLbl 
         AutoSize        =   -1  'True
         Caption         =   "Saturday"
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
         Left            =   720
         TabIndex        =   43
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Date:"
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
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame DaiMilFrm 
      Caption         =   "  Daily and Annual mileage.  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4080
      TabIndex        =   40
      Top             =   1920
      Width           =   3135
      Begin Tourwin2002.UTextBox DaiOddTxt 
         Height          =   285
         Left            =   1200
         TabIndex        =   24
         ToolTipText     =   "Ongoing mileage total (0 to 100,000)"
         Top             =   360
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         Max             =   6
         FieldType       =   1
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin Tourwin2002.UTextBox DaiMilTxt 
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         ToolTipText     =   "Days training mileage (range 0 to 999.99)"
         Top             =   720
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         Max             =   6
         FieldType       =   1
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin VB.Label DaiMilLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Milea&ge:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   345
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "&Oddometer:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   225
         TabIndex        =   23
         Top             =   375
         Width           =   975
      End
      Begin VB.Label DaiKmLbl 
         AutoSize        =   -1  'True
         Caption         =   "Km"
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
         Left            =   2160
         TabIndex        =   41
         Top             =   375
         Width           =   210
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Training Heart Rate Duration.  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      TabIndex        =   39
      ToolTipText     =   "Level No. 1"
      Top             =   120
      Width           =   3135
      Begin VB.VScrollBar DaiLevVSc 
         Height          =   975
         LargeChange     =   5
         Left            =   2640
         Max             =   5
         Min             =   1
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Value           =   5
         Width           =   255
      End
      Begin VB.TextBox DaiTotTxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Text            =   "00:00:00"
         Top             =   1320
         Width           =   855
      End
      Begin Tourwin2002.UTextBox DaiAboTxt 
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         ToolTipText     =   "Duration for given heart rate (hh:mm:ss)"
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Max             =   8
         FieldType       =   3
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin Tourwin2002.UTextBox DaiMidTxt 
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         ToolTipText     =   "Duration for given heart rate (hh:mm:ss)"
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         FieldType       =   3
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin Tourwin2002.UTextBox DaiBelTxt 
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         ToolTipText     =   "Duration for given heart rate (hh:mm:ss)"
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         FieldType       =   3
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldName   =   ""
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total time: "
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
         Left            =   840
         TabIndex        =   19
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label DaiAboLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Above :"
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
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label DaiMidLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Between: "
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
         Left            =   555
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label DaiBelLbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Below 140:"
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
         Left            =   255
         TabIndex        =   17
         Top             =   960
         Width           =   1395
      End
   End
   Begin VB.Frame DaiFrame 
      Caption         =   "                             Days Predefined Peak Performance Schedule.                             "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   30
      Top             =   3120
      Width           =   6855
      Begin VB.TextBox DaiPWeTxt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Week"
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox DaiPMoTxt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Month"
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox DaiPDaTxt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Day"
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label15 
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   975
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Week"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   625
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Week"
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
         Index           =   1
         Left            =   -480
         TabIndex        =   34
         Top             =   600
         Width           =   435
      End
   End
   Begin VB.TextBox DaiMemTxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Text            =   "DAILYFRM.frx":052E
      Top             =   2400
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label DaiEdiLab 
      AutoSize        =   -1  'True
      Caption         =   "Daily Editor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3000
      TabIndex        =   35
      Top             =   2040
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "DailyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_loading               As Boolean      ' Flag for loading form...
Dim Changed                 As Integer           ' Flag if user makes field change
Dim cLevel(1 To 5)          As String * 3 ' All level Heart Vector definitions
Dim StartLvl As String * 1       ' Start Level

Sub Data_To_Scr(AddType As Integer)
' ------------------------------------
' Either clears screen, if AddType = 0
' or put SetDai value into Controls
' ------------------------------------
On Local Error GoTo To_Scr_Err

' Declare Local variables
'Dim HeaV As String
Dim i As Integer
Dim TypeNum As String, RetStr As String * 255
Dim oControl    As Control

Dim HeaV As CStringBuilder
' Initialize Local Variables
Set HeaV = New CStringBuilder

If bDebug Then Handle_Err 0, "Data_To_Scr-DailyFrm"
' Always Clear Heart Rate values
DaiAboTxt.Text = "00:00:00"
DaiMidTxt.Text = "00:00:00"
DaiBelTxt.Text = "00:00:00"

If AddType = 0 Then
        ProgressBar "Loading default values...", -1, 1, -1
        DaiDatTxt.Text = Format$(Date, objMdi.info.cSHORTDATEFORMAT)
        DaiDatLbl.Caption = Format$(Date, "dddd")
        If DaiDayCbo.ListCount > 0 Then DaiDayCbo.ListIndex = 0
        'DaiDayCbo.Text = "Regular"
        DaiHeaTxt.Text = "0"
        DaiSlpTxt.Text = "00:00:00"
        DaiRatTxt.Text = "0"
        DaiOddTxt.Text = "0"
        ProgressBar "Loading default values...", -1, 5, -1
        DaiWeiTxt.Text = "0"
        DaiMilTxt.Text = "0"
        DaiTotTxt.Text = "00:00:00"
        ProgressBar "Loading default values...", -1, 8, -1
        DaiMemTxt.Text = " "
        DaiPDaTxt.Text = PeakNames(0)
        DaiPWeTxt.Text = PeakNames(0)
        DaiPMoTxt.Text = PeakNames(0)
        For i = 1 To 9 Step 3
            aLevel(i) = "00:00:00"
            aLevel(i + 1) = "00:00:00"
            aLevel(i + 2) = "00:00:00"
        Next i
        StartLvl = 2
        Set_Act_Hea_Cap 1
        
        ProgressBar "Loading default values...", -1, 10, -1
        ProgressBar "", 0, 0, 0
        ' ---------
        ' Day exist
        ' ---------
Else

' *************************
' * LATE BINDING TOO SLOW *
' *************************

'    For Each oControl In DailyFrm.Controls
'        If TypeOf oControl Is UTextBox Then
'            If "" <> oControl.DataFieldName Then
'                oControl.Text = ObjTour.DBGetField(oControl.DataFieldName, iDailyDB)
'            End If
'        End If
'    Next
    
    ProgressBar "Loading default values...", -1, 1, -1
    
    DaiDayCbo.Text = IIf("" = ObjTour.DBGetField("DayType", iDailyDB), " ", ObjTour.DBGetField("DayType", iDailyDB)) 'error
    DaiHeaTxt.Text = ObjTour.DBGetField("Heart", iDailyDB)
    DaiSlpTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_Sleep, iDailyDB), "00:00:00", ObjTour.DBGetField(gcDai_Tour_Sleep, iDailyDB))
    DaiRatTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_DayRate, iDailyDB), "0", ObjTour.DBGetField(gcDai_Tour_DayRate, iDailyDB))
    DaiOddTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_Oddometer, iDailyDB), "0", ObjTour.DBGetField(gcDai_Tour_Oddometer, iDailyDB))
    ProgressBar "Loading default values...", -1, 2, -1
    DaiWeiTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_Weight, iDailyDB), "000.00", ObjTour.DBGetField(gcDai_Tour_Weight, iDailyDB))
    DaiMilTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_DaiMile, iDailyDB), 0, ObjTour.DBGetField(gcDai_Tour_DaiMile, iDailyDB))
    DaiTotTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_TotalHrs, iDailyDB), "00:00:00", ObjTour.DBGetField(gcDai_Tour_TotalHrs, iDailyDB))

    DaiDatTxt.Text = Format$(ObjTour.DBGetField(gcDai_Tour_Date, iDailyDB), "mm-dd-yyyy") 'objMdi.info.cSHORTDATEFORMAT) 'Read only?
    DaiDatLbl.Caption = Format$(ObjTour.DBGetField(gcDai_Tour_Date, iDailyDB), "dddd")
        ' ----------------------
        ' look to see which 3
        ' Vectors are being used
        ' for active heart rate!
        ' ----------------------
        StartLvl = IIf("" = ObjTour.DBGetField("Level", iDailyDB), 1, ObjTour.DBGetField("Level", iDailyDB))
        ' =====================================
        ' A little fudge to get the
        ' Heart rate label to update correctly
        ' StartLvl is updated in DaiLevSc
        ' with it's original value.
        ' ====================================
        Dim iTemp As Integer
        'iTemp = StartLvl
        'StartLvl = 0
        DaiLevVSc.Value = (5 - StartLvl) + 1
        
        For i = 1 To 5
            TypeNum = "Level" & Format$(i, "0")
            cLevel(i) = IIf("" = ObjTour.DBGetField(TypeNum, iDailyDB), "123", ObjTour.DBGetField(TypeNum, iDailyDB))
        Next i
        ProgressBar "Loading default values...", -1, 5, -1
        For i = 1 To 9
            TypeNum = "HeaV" & Format(i, "0")
            aLevel(i) = ObjTour.DBGetField(TypeNum, iDailyDB)
        Next i
        ProgressBar "Loading default values...", -1, 7, -1
        Set_Act_Hea_Cap (0) 'RemoveThis
        ' -------------------
        ' Peak Schedule types
        ' -------------------
        
        DaiPDaTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_Peak_Day, iDailyDB), "No Peak Schedule", ObjTour.DBGetField(gcDai_Tour_Peak_Day, iDailyDB))
        DaiPWeTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_Peak_Wee, iDailyDB), "No Peak Schedule", ObjTour.DBGetField(gcDai_Tour_Peak_Wee, iDailyDB))
        DaiPMoTxt.Text = IIf("" = ObjTour.DBGetField(gcDai_Tour_Peak_Mon, iDailyDB), "No Peak Schedule", ObjTour.DBGetField(gcDai_Tour_Peak_Mon, iDailyDB))
        ProgressBar "Loading default values...", -1, 8, -1
        If "" = ObjTour.DBGetField(gcDai_Tour_Memo, iDailyDB) Then
            DaiMemTxt.Text = " "
        Else
            DaiMemTxt.Text = ObjTour.DBGetField(gcDai_Tour_Memo, iDailyDB)
        End If
        ProgressBar "Loading default values...", -1, 10, -1
        ProgressBar "", 0, 0, 0
End If
Exit Sub
To_Scr_Err:
    
    If bDebug Then Handle_Err Err, "Data_To_Scr-DailyFrm"
    Resume Next
End Sub


Sub Data_To_Set(NewRec%)
' ---------------
' Check whether record already
' exist or if new record!
' ---------------
On Local Error GoTo To_Set_Err

' Declare local variables

Dim SQL As String, CurDate As String, HeaV As String
Dim i As Integer, TypeNum As String
Dim lEventHandle As Long
If bDebug Then Handle_Err 0, "Data_To_Set-Daily"

' Get a DB Handle
lEventHandle = ObjTour.GetHandle

' Open Database
'ObjTour.DBOpen gcEve_Tour, gcEve_Tour_Peak, lEventHandle

' Define RecordSet
CurDate = "#" & Format$(DaiDatTxt.Text, objMdi.info.cSHORTDATEFORMAT) & "#"
SQL = "SELECT * FROM " & gcEve_Tour_Peak & " WHERE Date like " & CurDate & " And  Active BETWEEN 1 AND 3 AND Id = " & objMdi.info.ID
ObjTour.RstSQL lEventHandle, SQL

ProgressBar "Saving data...", -1, 1, -1

ObjTour.DBSetField gcDai_Tour_Date, DaiDatTxt.Text, iDailyDB
ObjTour.DBSetField gcDai_Tour_DaiType, gcDAI_NORMAL, iDailyDB   'Normal Daily Type

ObjTour.DBSetField gcDai_Tour_DayType, Mid$(DaiDayCbo.Text, 1, 30), iDailyDB
ObjTour.DBSetField gcDai_Tour_Heart, DaiHeaTxt.Text, iDailyDB
ObjTour.DBSetField gcDai_Tour_Sleep, DaiSlpTxt.Text, iDailyDB
ObjTour.DBSetField gcDai_Tour_DayRate, DaiRatTxt.Text, iDailyDB
ObjTour.DBSetField gcDai_Tour_Oddometer, DaiOddTxt.Text, iDailyDB
ObjTour.DBSetField gcDai_Tour_Weight, DaiWeiTxt.Text, iDailyDB
ObjTour.DBSetField gcDai_Tour_DaiMile, DaiMilTxt.Text, iDailyDB
ObjTour.DBSetField gcDai_Tour_TotalHrs, DaiTotTxt, iDailyDB
ObjTour.DBSetField gcDai_Tour_WeeklySunday, DateAdd("d", (-1 * WeekDay(DaiDatTxt.Text, vbSunday)) + 1, DaiDatTxt.Text), iDailyDB

' ----------------------
' Use Level to determine
' which vectors hours go
' into
' ----------------------
ProgressBar "Saving data...", -1, 2, -1
If NewRec Then              ' Format all fields then
                            ' update with edit below.
    ObjTour.DBSetField gcDai_Tour_Id, objMdi.info.ID, iDailyDB
    For i = 1 To 9
        HeaV = "HeaV" & Format$(i, "0")
            ObjTour.DBSetField HeaV, "00:00:00", iDailyDB
    Next i
        'Set Daily schedule
            ObjTour.DBSetField gcDai_Tour_Peak_Day, "No Current Peak Schedule", iDailyDB
            ObjTour.DBSetField gcDai_Tour_Peak_Wee, "No Current Peak Schedule", iDailyDB
            ObjTour.DBSetField gcDai_Tour_Peak_Mon, "No Current Peak Schedule", iDailyDB
    ' Load levels
For i = 1 To 5                      'Loop throught the nine type names
    TypeNum = "Level" & Format$(i, "0")
    ObjTour.DBSetField TypeNum, cLevel(i), iDailyDB
Next i

End If

For i = 1 To 9             ' Update fields with user edits
    HeaV = "HeaV" & Format$(i, "0")
    ObjTour.DBSetField HeaV, IIf(IsNull(aLevel(i)), "00:00:00", aLevel(i)), iDailyDB
Next i
ProgressBar "Saving data...", -1, 3, -1

ObjTour.DBSetField "Level", StartLvl, iDailyDB
ObjTour.DBSetField "Memo", DaiMemTxt.Text, iDailyDB

' ----------------------
' Need to calculate Peak
' Value. Default to Zero
' ----------------------
ProgressBar "Saving data...", -1, 4, -1

If NewRec And ObjTour.RstRecordCount(lEventHandle) <> 0 Then ' Seek for peak events occuring on this day
    'setEve.MoveFirst
    ObjTour.DBMoveFirst lEventHandle
    
Do While Not ObjTour.EOF(lEventHandle)
'    If setEve("Active") = 1 Then
'        SetDai("Peak_Day") = setEve("Descr")
'    ElseIf setEve("Active") = 2 Then
'        SetDai("Peak_Wee") = setEve("Descr")
'    ElseIf setEve("Active") = 3 Then
'        SetDai("Peak_Mon") = setEve("Descr")
'    End If

    If ObjTour.DBGetField("Active", lEventHandle) = 1 Then
            ObjTour.DBSetField gcDai_Tour_Peak_Day, ObjTour.DBGetField("Descr", lEventHandle), iDailyDB
    ElseIf ObjTour.DBGetField("Active", lEventHandle) = 2 Then
            ObjTour.DBSetField gcDai_Tour_Peak_Wee, ObjTour.DBGetField("Descr", lEventHandle), iDailyDB
    ElseIf ObjTour.DBGetField("Active", lEventHandle) = 3 Then
            ObjTour.DBSetField gcDai_Tour_Peak_Mon, ObjTour.DBGetField("Descr", lEventHandle), iDailyDB
    End If

    ObjTour.DBMoveNext lEventHandle
    ProgressBar "Saving data...", -1, 7, -1
Loop

    Changed = -1
End If              ' end If NewRec.

ObjTour.DBClose lEventHandle
ObjTour.FreeHandle lEventHandle
ProgressBar "", 0, 0, 0
Exit Sub
To_Set_Err:
        If bDebug Then Handle_Err Err, "Data_To_Set-DailyFrm"
        Resume Next
End Sub


Sub GetPeakNames()
'-------------------------------
' Retrieves Peak day description
' form tour ini file
' ------------------------------
On Local Error GoTo GetPeak_Err
If bDebug Then Handle_Err 0, "GetPeakName-DailyFrm"
Dim i As Integer
Dim PeakKey As String, SQL As String

ProgressBar "Loading Peak...", -1, 1, -1
' ======================
' Define RecordSet then
' Get Fields
' =====================
SQL = "SELECT * FROM " & gcNameTour_PeakNames & " WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

For i = 0 To 9                          'Loop through ini file
    PeakKey = "Peak" & Format$(i, "0")
    PeakNames(i) = Get_NameTour_PeakNames(PeakKey)
    ProgressBar "Loading Peak...", -1, i * 1.11, -1
Next i
ProgressBar " ", 0, 0, 0
Exit Sub
GetPeak_Err:
    If bDebug Then Handle_Err Err, "GetPeakNames-DailyFrm"
    Resume Next
End Sub

Sub Load_Levels(FromLevel As Integer)
On Local Error GoTo Load_Lev_Err

Dim i As Integer, TypeNum As String
Dim SQL As String
Dim dbLevel As Database, setLevel As Recordset
' --------------------------------------------
' if FromIni = -1 then load clevels with
' ini file setting, else get from daitour file
' --------------------------------------------
If bDebug Then Handle_Err 0, "Load_Levels-DailyFrm"

ProgressBar "Loading Levels...", -1, 1, -1

SQL = "SELECT * FROM Levels WHERE Id = " & objMdi.info.ID
ObjTour.RstSQL iSearcherDB, SQL

'If FromLevel Then
' --------------------
' Get Level Definetion
' From NameTour/Levels
' --------------------
    If ObjTour.RstRecordCount(iSearcherDB) <> 0 Then
        For i = 1 To 5
            TypeNum = "Level" + Format$(i, "0")  ' Format for Field Name.
            cLevel(i) = ObjTour.DBGetField(TypeNum, iSearcherDB)
            ProgressBar "Loading Daily...", -1, i * 1.8, -1
        Next i
    Else
        For i = 1 To 5
            cLevel(i) = "000"
        Next i
    End If
'Else    ' Get Level from Daily file...

'For i = 1 To 5
'    TypeNum = "Level" + Format$(i, "0")  ' Format for InI Key Name.
'    cLevel(i) = ObjTour.DBGetField(TypeNum, iSearcherDB) ' SetDai(TypeNum)
'    ProgressBar "Loading Daily...", -1, i * 1.8, -1
'Next i
'
'End If

ProgressBar "", 0, 0, 0

Exit Sub
Load_Lev_Err:
    If bDebug Then Handle_Err Err, "Load_Levels-DailyFrm"
    Resume Next
End Sub

Sub LoadControls()
' --------------------------------
' Purpose: Load user select fields
' --------------------------------
Dim SQL As String
Dim lUserHandle As Long

'Impovement: On load of program read to setting into memory ...
lUserHandle = ObjTour.GetHandle

'ObjTour.DBOpen App.Path & "\" & gcUserTour, gcUserTour_DailyOpt, lUserHandle

' define recordset
SQL = "Select * " & gcUserTour_DailyOpt & " WHERE Name = '" & objMdi.info.Name & "'"
ObjTour.RstSQL lUserHandle, SQL

If ObjTour.NoMatch(lUserHandle) Then
    MsgBox "User is not define in DailyOpt table"
Exit Sub
End If
DaiSlpTxt.Visible = IIf(ObjTour.DBGetField("Sleep", lUserHandle), -1, 0)
DaiSlpLbl.Visible = IIf(ObjTour.DBGetField("Sleep", lUserHandle), -1, 0)

ObjTour.DBClose lUserHandle
ObjTour.FreeHandle lUserHandle

End Sub

Sub LoadDayType()
On Local Error GoTo LoadDayType_Err
Dim i As Integer, RetStr As String, EventType As String
Dim SQL As String

If bDebug Then Handle_Err 0, "LoadDayType-DailyFrm"

Screen.MousePointer = vbHourglass ' Mouse to hour glass
DaiDayCbo.Clear

cActivityNames.Type_ID = gcActive_Type_EventNames
' -------------------------------------------------
' Based on Record count show appropriate controls
' and load the first 10 descriptions.
' -------------------------------------------------
If cActivityNames.StartSearch(gcActive_Type_EventNames) Then

    Do
        DaiDayCbo.AddItem cActivityNames.Description
        
    Loop While cActivityNames.GetNext
    
End If

ProgressBar " ", 0, 0, 0

Exit Sub
LoadDayType_Err:
     If bDebug Then Handle_Err Err, "LoadDayType-DailyFrm"
    Resume Next
End Sub

Sub LoadVisible()
'
' Sets visible properites based on user settings
' ----------------------------------------------
Dim TrueFalse As Integer

TrueFalse = 0
With objMdi.info

    DaiSlpTxt.Visible = .DailyOpt.GetBool(BitFlags.Daily_Sleep)
    DaiSlpLbl.Visible = .DailyOpt.GetBool(BitFlags.Daily_Sleep)


    DaiDayCbo.Visible = .DailyOpt.GetBool(BitFlags.Daily_DayType)
    DaiDaiLbl.Visible = .DailyOpt.GetBool(BitFlags.Daily_DayType)


    DaiHeaTxt.Visible = .DailyOpt.GetBool(BitFlags.Daily_RestHeart)
    DaiHeaLbl.Visible = .DailyOpt.GetBool(BitFlags.Daily_RestHeart)
    

    DaiRatTxt.Visible = .DailyOpt.GetBool(BitFlags.Daily_Description)
    DaiRatLbl.Visible = .DailyOpt.GetBool(BitFlags.Daily_Description)


    DaiWeiTxt.Visible = .DailyOpt.GetBool(BitFlags.Daily_Weight)
    DaiWeiLbl.Visible = .DailyOpt.GetBool(BitFlags.Daily_Weight)
    
    If .DailyOpt.GetBool(BitFlags.Daily_Metric) Then
        DaiKmLbl.Caption = "Km"
    Else
        DaiKmLbl.Caption = "Miles"
    End If
End With


End Sub

Sub M_Next()
Dim DateStr As String
Dim varCurrentPos As Variant
' -----------------
' Check for changes
' -----------------
If bDebug Then Handle_Err 0, "DaiNexCmdMouUP-DailyFrm"

If Changed Then
    If vbYes = MsgBox("Save Changes?", vbYesNo + vbQuestion, LoadResString(gcTourVersion)) Then DaiSavmnu
End If
      
    DateStr = "Date > " & "#" & Format$(DaiDatTxt.Text, objMdi.info.cSHORTDATEFORMAT) & "#"
    
   ' varCurrentPos = ObjTour.BookMark(iDailyDB)

    ObjTour.DBFindNext DateStr, iDailyDB 'SetDai.FindNext DateStr
    
    If Not ObjTour.NoMatch(iDailyDB) Then
        DaiPreCmd.Enabled = -1
        Data_To_Scr (1)
    Else
        ' Use book mark and restore orginal position
        'ObjTour.BookMark iDailyDB, varCurrentPos
        DaiNexCmd.Enabled = 0
    End If
Changed = 0
End Sub

Sub M_Previous()
' -----------------
' Check for changes
' -----------------
Dim DateStr As String
Dim varCurrentPos As Variant

If bDebug Then Handle_Err 0, "DaiPreCmdMouUp-DailyFrm"

On Local Error GoTo M_Previous_Err

If Changed Then
    If vbYes = MsgBox("Save Changes?", vbYesNo + vbQuestion, LoadResString(gcTourVersion)) Then DaiSavmnu
End If
            ' -----------------------
            ' Move to Previous Record
            ' Check current location
            ' -----------------------
        
            'varCurrentPos = ObjTour.BookMark(iDailyDB)
            DateStr = "Date < " & "#" & Format$(DaiDatTxt.Text, objMdi.info.cSHORTDATEFORMAT) & "# AND ID=" & objMdi.info.ID
            ObjTour.DBFindPrevious DateStr, iDailyDB
            
            If Not ObjTour.NoMatch(iDailyDB) Then
            
                DaiPreCmd.Enabled = -1
                DaiNexCmd.Enabled = -1
            Else
                ' User book mark a restore orginal position
                'ObjTour.BookMark iDailyDB, varCurrentPos
                DaiPreCmd.Enabled = 0
                
            End If

                Data_To_Scr (1)
                Changed = 0
Exit Sub
M_Previous_Err:
    If bDebug Then Handle_Err Err, "DaiPreCmdMouUp-DailyFrm"
    Resume Next
End Sub

Sub Move_To(SearchDate As String)
On Local Error GoTo Move_To_Err
Dim DateStr As String

    DateStr = "Date = " & "#" & Format$(SearchDate, "mm-dd-yyyy") & "#"
    DaiDatTxt.Text = Format$(SearchDate, objMdi.info.cSHORTDATEFORMAT)
    ObjTour.DBFindFirst DateStr, iDailyDB
    
    If Not ObjTour.NoMatch(iDailyDB) Then
        Data_To_Scr (1)
        Changed = 0
    Else
        Data_To_Scr (0)
        Changed = 0
    End If
Exit Sub
Move_To_Err:
    If bDebug Then Handle_Err Err, "Move_to_Err-DailyFrm"
    Resume Next
End Sub

Sub Set_Act_Hea_Cap(Index%)
'------------------------------------------------------------------
' Data Dictionary
' StartLvl: contains the Level value user is changing from
' CurVec  : number inwhich text value is written to ie Heart Vector
' Index   : number of level number user selected.
' cLevel  : range is 0 - 8 thus + 1 for correct vector number
' NewVec  : Update text box with heart Vector number
' This Sub is Called for three reasons:
'   --------------------------------------
'   1) Daily form is being loaded
'   2) Single Heart Zone has been modified
'   3) User selects different zone set.
' -----------------------------------------------------------------
On Local Error GoTo Act_Hea_Err

Dim RetStr As String
Dim i As Integer
Dim CurVec As String * 1, NewVec As String * 1
Dim cTempHeart As cHeartNameVar

cHeartNames.LoadValues objMdi.info.ID ' No harm in calling even if already loaded...

'Reason 1)
If Index% = 0 Then
If bDebug Then Handle_Err 0, "Set_Act_Hea_Cap-DailyFrm Reason 1"

For i = 1 To 3
    ' Get New description for InI
    'TypeNum = "Heart" & Format$(Mid$(cLevel(StartLvl), i, 1), "0") ' Format for InI Key Name.
        
    Set cTempHeart = cHeartNames(Mid$(cLevel(StartLvl), i, 1))
    If Not cTempHeart Is Nothing Then
        RetStr = cTempHeart.Description
    Else
        RetStr = "No Return"
    End If
                        
    
    CurVec = Mid$(cLevel(StartLvl), i, 1)
            If Mid$(RetStr, 1, 9) = "No Return" Then RetStr = "0 BPM"
            Select Case i
                    Case 1:
                    If CurVec <> 0 Then DaiBelTxt.Text = aLevel(CurVec)
                    DaiBelLbl.Caption = Trim(RetStr)
                    Case 2:
                    If CurVec <> 0 Then DaiMidTxt.Text = aLevel(CurVec)
                    DaiMidLbl.Caption = Trim(RetStr)
                    Case 3:
                    If CurVec <> 0 Then DaiAboTxt.Text = aLevel(CurVec)
                    DaiAboLbl.Caption = Trim(RetStr)
            End Select
Next i
Exit Sub

End If      ' End Reason 1 IF

'  Reason 2)
If StartLvl = Index Then   'Text box lost focus, update user
If bDebug Then Handle_Err 0, "Set_Act_Hea_Cap-Daily Reason 2"                            ' change to alevel
For i = 1 To 3
    ' Get New description for InI
    CurVec = Mid$(cLevel(StartLvl), i, 1)
            ' -----------------------------
            ' Store existing heart value then
            ' insert new heart time value.
            ' -----------------------------
            If CurVec <> 0 And CurVec <> "" Then
            Select Case i
                    Case 0: ' not used
                    Case 1:
                    If aLevel(CurVec) <> DaiBelTxt.Text Then Changed = -1
                    aLevel(CurVec) = DaiBelTxt.Text
                    Case 2:
                    If aLevel(CurVec) <> DaiMidTxt.Text Then Changed = -1
                    aLevel(CurVec) = DaiMidTxt.Text
                    Case 3:
                    If aLevel(CurVec) <> DaiAboTxt.Text Then Changed = -1
                    aLevel(CurVec) = DaiAboTxt.Text
            End Select
            End If
Next i
    
Exit Sub
End If  ' Lost Focus


If bDebug Then Handle_Err 0, "Set_Act_Hea_Cap-Daily Reason 3"

For i = 1 To 3
    ' Get New description for InI           StartLvl
    'TypeNum = "Heart" + Format$(Mid$(cLevel(Index), i, 1), "0") ' Format for InI Key Name.
    
    Set cTempHeart = cHeartNames(Mid$(cLevel(Index), i, 1))
    If Not cTempHeart Is Nothing Then
        RetStr = cTempHeart.Description
    Else
        RetStr = "No Return"
    End If
    
    CurVec = Mid$(cLevel(StartLvl), i, 1)
    NewVec = Mid$(cLevel(Index), i, 1)
    
    If RetStr = "No Return" Then RetStr = "Field Not Used"
    ' -----------------------------
    ' Store existing heart value then
    ' insert new heart time value.
    ' -----------------------------
    
            Select Case i
                    Case 1:
                    If CurVec <> 0 Then aLevel(CurVec) = DaiBelTxt.Text
                    DaiBelLbl.Caption = Trim(RetStr)
                    If NewVec = 0 Then
                        DaiBelTxt.Text = "00:00:00"
                    Else
                        DaiBelTxt.Text = aLevel(NewVec)
                    End If
                    Case 2:
                    If CurVec <> 0 Then aLevel(CurVec) = DaiMidTxt.Text
                    DaiMidLbl.Caption = Trim(RetStr)
                     If NewVec = 0 Then
                        DaiMidTxt.Text = "00:00:00"
                    Else
                        DaiMidTxt.Text = aLevel(NewVec)
                    End If
                    Case 3:
                    If CurVec <> 0 Then aLevel(CurVec) = DaiAboTxt.Text
                    DaiAboLbl.Caption = Trim(RetStr)
                     If NewVec = 0 Then
                        DaiAboTxt.Text = "00:00:00"
                    Else
                        DaiAboTxt.Text = aLevel(NewVec)
                    End If
            End Select
Next i

StartLvl = Index
Exit Sub
Act_Hea_Err:
    If bDebug Then Handle_Err Err, "Set_Act_Hea_Cap-DailyFrm"
    Resume Next
End Sub


Function UpdateTotalText()
' -----------------------------------------------------------------------
' Purpose: Updates DaiTotTxt text box with total hours from each of the 9
' different heart rate levels.
' -----------------------------------------------------------------------
Dim Hr As Integer, Mn As Integer, Sc As Integer, i As Integer
Dim Tot As String


Sc = Val(Mid$(aLevel(1), 7, 2)) + Val(Mid$(aLevel(2), 7, 2)) + Val(Mid$(aLevel(3), 7, 2))
Sc = Sc + Val(Mid$(aLevel(4), 7, 2)) + Val(Mid$(aLevel(5), 7, 2)) + Val(Mid$(aLevel(6), 7, 2))
Sc = Sc + Val(Mid$(aLevel(7), 7, 2)) + Val(Mid$(aLevel(8), 7, 2)) + Val(Mid$(aLevel(9), 7, 2))
Select Case Int(Sc / 60)
        Case 0:
        Case 1 To 9:
            Mn = Int(Sc / 60)
            Sc = Sc - (Int(Sc / 60) * 60)
End Select 'Ends second formating

Mn = Mn + Val(Mid$(aLevel(1), 4, 2)) + Val(Mid$(aLevel(2), 4, 2)) + Val(Mid$(aLevel(3), 4, 2))
Mn = Mn + Val(Mid$(aLevel(4), 4, 2)) + Val(Mid$(aLevel(5), 4, 2)) + Val(Mid$(aLevel(6), 4, 2))
Mn = Mn + Val(Mid$(aLevel(7), 4, 2)) + Val(Mid$(aLevel(8), 4, 2)) + Val(Mid$(aLevel(9), 4, 2))

Select Case Int(Mn / 60)
    Case 0:
    Case 1 To 9:
        Hr = Int(Mn / 60)
        Mn = Mn - (Int(Mn / 60) * 60)
End Select

Hr = Hr + Val(Mid$(aLevel(1), 1, 2)) + Val(Mid$(aLevel(2), 1, 2)) + Val(Mid$(aLevel(3), 1, 2))
Hr = Hr + Val(Mid$(aLevel(4), 1, 2)) + Val(Mid$(aLevel(5), 1, 2)) + Val(Mid$(aLevel(6), 1, 2))
Hr = Hr + Val(Mid$(aLevel(7), 1, 2)) + Val(Mid$(aLevel(8), 1, 2)) + Val(Mid$(aLevel(9), 1, 2))

DaiTotTxt = Format$(Hr, "00") & ":" & Format$(Mn, "00") & ":" & Format$(Sc, "00")
End Function

Private Sub cmdDate_Click()
Dim CMaxDate As New CCalendar
Dim lRt As Long
Dim lTop As Long
Dim lLeft As Long

Const USERCANCELLED = 0
Const USERSELECTEDDATE = 1

'Calculate placement of Calendar Dialog...
lTop = Me.top + DaiDatTxt.top + (3 * DaiDatTxt.Height) + MENUANDTOOLBARHEIGTH
lLeft = Me.left + DaiDatTxt.left + ICONBARWIDTH + Frame2.left

CMaxDate.MaxDate = CDate(DaiDatTxt.Text)
lRt = CMaxDate.ShowCalendar(lTop, lLeft, Me)
Select Case lRt
    Case USERCANCELLED:
    Case USERSELECTEDDATE:
        Call Move_To(Format$(CMaxDate.MaxDate, objMdi.info.cSHORTDATEFORMAT))
End Select
Set CMaxDate = Nothing

End Sub

Private Sub cmdHistorical_Click()

Dim bExist As Boolean
Dim cHistory As New CHistorical
Dim sTemp As String
Dim lHistoryHandle As Long
Dim i As Long



    If cHistory.DoesHistoryRecordExist(DaiDatTxt.Text) Then
        'Comment Field
        cHistory.GetHistoricalField gcDai_Tour_Memo, sTemp
        frmHistorical.HistoricalComment = sTemp
        'Name Field
        cHistory.GetHistoricalField gcDai_Tour_DayType, sTemp
        frmHistorical.HistoricalName = sTemp
        'Time Field
        cHistory.GetHistoricalField gcDai_Tour_TotalHrs, sTemp
        frmHistorical.HistoricalTime = sTemp
        'Distance Field
        cHistory.GetHistoricalField gcDai_Tour_DaiMile, sTemp
        frmHistorical.HistoricalDistance = sTemp
        frmHistorical.cmdDelete.Enabled = True
        bExist = True
    Else
        ' If no record is found, then default the values to
        ' the current daily values
            'Comment Field
        frmHistorical.HistoricalComment = Me.DaiMemTxt.Text
        'Name Field
        frmHistorical.HistoricalName = Me.DaiDayCbo.Text
        'Time Field
        frmHistorical.HistoricalTime = Me.DaiTotTxt.Text
        'Distance Field
        frmHistorical.HistoricalDistance = Me.DaiMilTxt.Text
        
        frmHistorical.cmdDelete.Enabled = False
        
        bExist = False
    End If
    
' Clear ComboBox
frmHistorical.cboHistorical.Clear
' Load combo box
cActivityNames.Type_ID = gcActive_Type_EventNames
' -------------------------------------------------
' Based on Record count show appropriate controls
' and load the first 10 descriptions.
' -------------------------------------------------
If cActivityNames.StartSearch(gcActive_Type_EventNames) Then
    Do
        frmHistorical.cboHistorical.AddItem cActivityNames.Description
    Loop While cActivityNames.GetNext
End If
    
'    sTemp = "SELECT * FROM " & gcNameTour_Events & " WHERE Id = " & objMdi.info.ID
'    ObjTour.RstSQL lHistoryHandle, sTemp
'
'    For I = 0 To 9                      'Loop throught the nine type names
'        sTemp = "Event" + Format$(I, "0")
'        sTemp = ObjTour.DBGetField(sTemp, lHistoryHandle)
'        If sTemp <> "No Return" Then
'                frmHistorical.cboHistorical.AddItem sTemp
'        End If
'    Next I
'    ObjTour.FreeHandle lHistoryHandle
    
    ' Find if Record exists...
    frmHistorical.Show vbModal

    Select Case LCase(frmHistorical.Tag)
            Case "ok":
            If bExist Then
            ' Update
                ObjTour.Edit cHistory.DBHandle
            Else
            ' Insert
                With frmHistorical
                    ObjTour.AddNew cHistory.DBHandle
                    ObjTour.DBSetField gcDai_Tour_Id, objMdi.info.ID, cHistory.DBHandle
                    ObjTour.DBSetField gcDai_Tour_DaiType, gcDAI_HISTORICAL, cHistory.DBHandle
                    ObjTour.DBSetField gcDai_Tour_Date, DaiDatTxt.Text, cHistory.DBHandle
              End With
            End If
        
            With frmHistorical
            ObjTour.DBSetField gcDai_Tour_DayType, .HistoricalName, cHistory.DBHandle
            ObjTour.DBSetField gcDai_Tour_TotalHrs, .HistoricalTime, cHistory.DBHandle
            ObjTour.DBSetField gcDai_Tour_DaiMile, .HistoricalDistance, cHistory.DBHandle
            ObjTour.DBSetField gcDai_Tour_Memo, .HistoricalComment, cHistory.DBHandle
            End With
            
            ObjTour.Update cHistory.DBHandle
                
            Case "delete":
                If bExist Then
                    ObjTour.Delete cHistory.DBHandle
                End If
            Case "cancel":
                ' do nothing
    End Select
    
    Unload frmHistorical
    Set cHistory = Nothing
    
End Sub


'Private Sub DaiAboTxt_KeyPress(KeyAscii As Integer)
'DaiAboTxt.SelLength = IIf(KeyAscii = 8, 0, 1)
'End Sub

Private Sub DaiAboTxt_LostFocus()
Set_Act_Hea_Cap (StartLvl)
UpdateTotalText
End Sub




Private Sub DaiAboTxt_ToolTip()
    Call ShowToolTipInStatusBar(DaiAboTxt.ToolTipText)
End Sub

'Private Sub DaiBelTxt_KeyPress(KeyAscii As Integer)
'DaiBelTxt.SelLength = IIf(KeyAscii = 8, 0, 1)
'End Sub

Private Sub DaiBelTxt_LostFocus()
Set_Act_Hea_Cap (StartLvl)
UpdateTotalText
End Sub


Private Sub DaiDatCnt_Validate(Action As Integer, Save As Integer)
On Local Error GoTo DaiDatCnt_Err

Select Case Action
        Case vbDataActionMovePrevious:
            M_Previous
        Case vbDataActionMoveNext:
            M_Next
End Select
Exit Sub
DaiDatCnt_Err:
    If bDebug Then Handle_Err Err, "DaiDatCnt_Validate-DailyFrm"
    Resume Next
End Sub



Private Sub DaiBelTxt_ToolTip()
Call ShowToolTipInStatusBar(DaiBelTxt.ToolTipText)
End Sub

'Private Sub DaiDatTxt_KeyPress(KeyAscii As Integer)
'    DaiDatTxt.SelLength = IIf(KeyAscii = 8, 0, 1)
'End Sub

Private Sub DaiDatTxt_LostFocus()
' --------------------
' check for valid date
' will
' --------------------
If bDebug Then Handle_Err 0, "DaiDatTxt-DailyFrm"
If m_loading Then Exit Sub

On Local Error GoTo DatTxt_Err
Dim RetInt As Integer

If Not IsDate(DaiDatTxt.Text) Then
    MsgBox LoadResString(gcInvalidDate) & " '" & DaiDatTxt.Text & "'", vbOKOnly + vbExclamation, LoadResString(gcInvalidFormat)
    DaiDatTxt.SetFocus
Else
    ' Update Day text with Current date value
        DaiDatLbl.Caption = Format$(DaiDatTxt.Text, "dddd")
End If

Exit Sub
DatTxt_Err:
    If bDebug Then Handle_Err Err, "DaiDatTxt-DailyFrm"
    Resume Next
End Sub


Private Sub DaiDatTxt_ToolTip()
Call ShowToolTipInStatusBar(DaiDatTxt.ToolTipText)
End Sub

Private Sub DaiDayCbo_Change()
Changed = -1
End Sub

Private Sub DaiDayCbo_KeyPress(KeyAscii As Integer)
If Len(DaiDayCbo.Text) > 30 Then
    'Don't Allow and Visible key strokes
    
    If KeyAscii < 49 Then Exit Sub
    
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub DaiDayCbo_LostFocus()
'
' check if change was made
' -------------------------
On Local Error Resume Next
If DaiDayCbo.Text <> ObjTour.DBGetField("DayType", iDailyDB) Then Changed = -1
End Sub


Sub DaiDelmnu()
'
'
Dim DateStr As String, sDeleteStr As String
Dim lEventHandle As Long


If bDebug Then Handle_Err 0, "DaiDelmnu-DailyFrm"


DateStr = "Date = " & "#" & Format$(DaiDatTxt, "mm-dd-yyyy") & "#"

ObjTour.DBFindFirst DateStr, iDailyDB
'SetDai.FindFirst DateStr
DoEvents
If ObjTour.NoMatch(iDailyDB) Then    ' Check if record is found.
   MsgBox Format$(DaiDatTxt, "mm-dd-yyyy") & " is not a saved record. No record to delete.", 32, "Warning"
   Exit Sub
Else
   ' ----------------------------
   ' Delete Daily Database Record
   ' ----------------------------
   ObjTour.Delete iDailyDB
   'SetDai.Delete
   ' ----------------------------
   ' Delete Daily Event record
   ' ----------------------------
   sDeleteStr = "DELETE * FROM " & gcEve_Tour_Event_Daily & " WHERE " & DateStr
   ' Get Database handle
'   lEventHandle = ObjTour.GetHandle
'   ObjTour.DBOpen gcEve_Tour, gcEve_Tour_Event_Daily, lEventHandle
   
   ' Delete event table daily records
   ObjTour.DBExecute sDeleteStr
   
   ' Release Event Handle
'   ObjTour.DBClose lEventHandle
'   ObjTour.FreeHandle lEventHandle
   
   'dbsEve_Tour.Execute sDeleteStr
   ' ----------------------------
   ' set the dynaset to first date
   ' before just deleted record
   ' ----------------------------
   'If SetDai.RecordCount > 0 Then       ' check record exist
   If ObjTour.RstRecordCount(iDailyDB) > 0 Then
        If Not ObjTour.BOF(iDailyDB) Then
            ObjTour.DBMoveNext iDailyDB
            'SetDai.MovePrevious
            If ObjTour.EOF(iDailyDB) Then ObjTour.DBMoveNext iDailyDB  'SetDai.MoveNext
        Else
            'SetDai.MoveNext
            ObjTour.DBMoveNext iDailyDB
            If ObjTour.EOF(iDailyDB) Then ObjTour.DBMovePrevious iDailyDB  'SetDai.MovePrevious
        End If
        Data_To_Scr 1
    Else
            MsgBox "All records deleted!", 32, "Empty database"
            MDI!MdiDelmnu.Enabled = False
    End If
    
End If

End Sub

Sub DaiEximnu()
If bDebug Then Handle_Err 0, "DaiPreCmdMouUp-DailyFrm"
On Local Error GoTo DaiExi_Err
If Changed = -1 Then
    If vbYes = MsgBox("Save Changes?", vbYesNo, " Save ") Then DaiSavmnu
End If
Unload DailyFrm
Exit Sub
DaiExi_Err:
    If bDebug Then Handle_Err Err, "DaiEximnu-DailyFrm"
    Resume Next
End Sub



Private Sub DaiHeaTxt_KeyPress(KeyAscii As Integer)
KeyAscii = TextValidate(KeyAscii, "Int Only", DaiHeaTxt)
DaiHeaTxt.SelLength = 1
End Sub

Private Sub DaiHeaTxt_LostFocus()
'
' check if change was made
'
On Local Error Resume Next
If Val(DaiHeaTxt) <> ObjTour.DBGetField("Heart", iDailyDB) Then Changed = -1
    End Sub

Private Sub DaiHeaTxt_ToolTip()
    Call ShowToolTipInStatusBar(DaiHeaTxt.ToolTipText)
End Sub

Private Sub DaiLevVSc_Change()

Set_Act_Hea_Cap (5 - DaiLevVSc.Value) + 1
Changed = -1
MDI.DaiLv1mnu.Checked = False
MDI.DaiLv2mnu.Checked = False
MDI.DaiLv3mnu.Checked = False
MDI.DaiLv4mnu.Checked = False
MDI.DaiLv5mnu.Checked = False

Select Case (5 - DaiLevVSc.Value) + 1
    Case 1:
        MDI.DaiLv1mnu.Checked = True
    Case 2:
        MDI.DaiLv2mnu.Checked = True
    Case 3:
        MDI.DaiLv3mnu.Checked = True
    Case 4:
        MDI.DaiLv4mnu.Checked = True
    Case 5:
        MDI.DaiLv5mnu.Checked = True
End Select

Frame1.ToolTipText = "Level No. " & CStr((5 - DaiLevVSc.Value) + 1)
'UpdateTotalText
    
End Sub

Private Sub DaiMemCmd_Click()
If bDebug Then Handle_Err 0, "DaiMemCmd-DailyFrm"
DaiButFra.Visible = False
DaiMemTxt.Visible = -1
DaiEdiLab.Visible = -1
DaiMemCmd.Visible = 0
DaiFrame.Visible = 0
DaiMilTxt.Visible = 0
DaiOddTxt.Visible = 0
DaiMilFrm.Visible = 0
DaiMemTxt.SetFocus
End Sub

Private Sub DaiMemTxt_Change()
Changed = -1
End Sub

Private Sub DaiMemTxt_LostFocus()
On Local Error GoTo DaiMemTxt_Err
If bDebug Then Handle_Err 0, "DaiMemTxtLostFocus-Daily"
DaiButFra.Visible = True
DaiMemTxt.Visible = 0
DaiEdiLab.Visible = 0
DaiMemCmd.Visible = -1
DaiPreCmd.Visible = -1
DaiFrame.Visible = -1
DaiMilTxt.Visible = -1
DaiOddTxt.Visible = -1
DaiMilFrm.Visible = -1
'
' check if change was made
'
If Trim$(DaiMemTxt) <> Trim$(ObjTour.DBGetField("Memo", iDailyDB)) Then Changed = -1
Exit Sub
DaiMemTxt_Err:
    If bDebug Then Handle_Err Err, "DaiMemTxt_LostFocus-DailyFrm"
    Resume Next
End Sub



'Private Sub DaiMidTxt_KeyPress(KeyAscii As Integer)
'DaiMidTxt.SelLength = IIf(KeyAscii = 8, 0, 1)
'End Sub

Private Sub DaiMidTxt_LostFocus()
Set_Act_Hea_Cap (StartLvl)
UpdateTotalText
End Sub


Private Sub DaiMidTxt_ToolTip()
    ShowToolTipInStatusBar DaiMidTxt.ToolTipText
End Sub



Private Sub DaiMilTxt_KeyPress(KeyAscii As Integer)
KeyAscii = TextValidate(KeyAscii, "Real Only", DaiMilTxt)
End Sub

Private Sub DaiMilTxt_LostFocus()
'
' check if change was made
'
On Local Error Resume Next
If Val(DaiMilTxt) <> ObjTour.DBGetField("DaiMile", iDailyDB) Then Changed = -1

End Sub


Sub DaiNewmnu()
' -----------------
' Check for changes
' -----------------
Const uYes = 6
Dim Resp As Integer
If bDebug Then Handle_Err 0, "DaiNewmnu-DailyFrm"
If Changed Then
    Resp = MsgBox("Save Changes?", 36, " Save ")
        If Resp = uYes Then
            DaiSavmnu
        End If
End If
Data_To_Scr (0)
' ---------------------------------------------------------
' This line of code was added. The problem was if the user
' opened the daily form before adding HeartLevels,
' the levels would always be stuck on "000"
' ---------------------------------------------------------
Load_Levels (-1)
MDI!MdiSavmnu.Enabled = True
DaiDatTxt.SetFocus
End Sub

Private Sub DaiMilTxt_ToolTip()
    ShowToolTipInStatusBar DaiMilTxt.ToolTipText
End Sub

Private Sub DaiNexCmd_Click()
 M_Next
End Sub


Private Sub DaiOddTxt_KeyPress(KeyAscii As Integer)
KeyAscii = TextValidate(KeyAscii, "Real Only", DaiOddTxt)
End Sub

Private Sub DaiOddTxt_LostFocus()
'
' check if change was made
'
On Local Error Resume Next
If Val(DaiOddTxt) <> ObjTour.DBGetField("Oddometer", iDailyDB) Then Changed = -1
If 0 = Val(DaiOddTxt.Text) Then Exit Sub

If "" = DaiMilTxt.Text Or 0 = Val(DaiMilTxt.Text) Then
Dim lTempHandle As Long
Dim sSQL As String
' Calculate Mileage value
' from last entered Oddometer entry.

' Set Recordset
sSQL = "SELECT Top 1 " & gcDai_Tour_Oddometer & ", " & gcDai_Tour_Date & _
       " FROM " & gcDai_Tour_Dai & _
       " WHERE ID = " & objMdi.info.ID & _
       " AND " & gcDai_Tour_Date & " < #" & DailyFrm.DaiDatTxt.Text & "#" & _
       " AND " & gcDai_Tour_Oddometer & " > 0 " & _
       " ORDER BY " & gcDai_Tour_Date & " Desc"
                             
Debug.Print sSQL
' If handle was not avaiable, exit sub
If Not ObjTour.RstSQL(lTempHandle, sSQL) Then
    ObjTour.FreeHandle lTempHandle
    Exit Sub
End If

If ObjTour.RstRecordCount(lTempHandle) = 0 Then
    ObjTour.FreeHandle lTempHandle
    Exit Sub
End If

DaiMilTxt = CStr(Val(DaiOddTxt.Text) - Val(ObjTour.DBGetField(gcDai_Tour_Oddometer, lTempHandle)))
ObjTour.FreeHandle lTempHandle

End If


End Sub

Private Sub DaiOddTxt_ToolTip()
    ShowToolTipInStatusBar DaiOddTxt.ToolTipText
End Sub

Private Sub DaiPreCmd_Click()
    M_Previous
End Sub


Private Sub DaiRatTxt_KeyPress(KeyAscii As Integer)
KeyAscii = TextValidate(KeyAscii, "Int Only", DaiRatTxt)
End Sub

Private Sub DaiRatTxt_LostFocus()
'
' check if change was made
'
On Local Error Resume Next
If Val(DaiRatTxt) <> ObjTour.DBGetField("DayRate", iDailyDB) Then
    Changed = -1
End If

End Sub

Sub DaiSavmnu()
' ---------------------------------
' If all control values are
' valid for transfer to set dynaset
' the determine if record already
' exist or new record!
' ---------------------------------
Dim DateStr As String

If bDebug Then Handle_Err 0, "DaiSavmnu-Daily"
On Local Error GoTo SaveError
Screen.MousePointer = vbHourglass
'
' Check for valid date format
'
If Not IsDate(DaiDatTxt.Text) Then
    MsgBox LoadResString(gcInvalidDate) & " '" & DaiDatTxt.Text & "'", 0, LoadResString(gcInvalidFormat)
    DaiDatTxt.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If

If Not IsNumeric(DaiMilTxt.Text) Then
    MsgBox LoadResString(gcInvalidNumeric) & " '" & DaiMilTxt.Text & "'", 0, LoadResString(gcInvalidFormat)
    DaiMilTxt.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If

If Not IsNumeric(DaiOddTxt.Text) Then
    MsgBox LoadResString(gcInvalidNumeric) & " '" & DaiOddTxt.Text & "'", 0, LoadResString(gcInvalidFormat)
    DaiOddTxt.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If


DateStr = "Date = " & "#" & Format$(DaiDatTxt.Text, objMdi.info.cSHORTDATEFORMAT) & "#"
ObjTour.DBFindFirst DateStr, iDailyDB
Set_Act_Hea_Cap Val(StartLvl)

If ObjTour.NoMatch(iDailyDB) Then    ' Check if record is found.
        ObjTour.AddNew iDailyDB
        Data_To_Set -1       ' moves screen values to dynaset.
        DateStr = "Date = " & "#" & Format$(ObjTour.DBGetField("Date", iDailyDB), "mm-dd-yyyy") & "#"
        ObjTour.Update iDailyDB
        'SetDai.Update   ' When update is issued with 'Addnew'
                        ' first record has focus!!!
                        ' Query for new record and set active
                        ' day to newly created record.

        ObjTour.RstSQL iDailyDB, " SELECT * " & _
                                 " FROM " & gcDai_Tour_Dai & _
                                 " WHERE Id = " & objMdi.info.ID & " AND Type = " & CStr(gcDAI_NORMAL) & _
                                 " ORDER BY Date ASC"
        
        ObjTour.DBFindFirst DateStr, iDailyDB

        UpdateEventdb "Daily", ObjTour.DBGetField("Date", iDailyDB), ObjTour.DBGetField("DayType", iDailyDB), 0, 0, ""
        ' -----------------
        ' Check Next and
        ' previous settings
        ' -----------------
        If Not ObjTour.BOF(iDailyDB) Then DaiPreCmd.Enabled = True
        
        If Not ObjTour.EOF(iDailyDB) Then DaiNexCmd.Enabled = True
        
Else
        If LockRecordEdit(iDailyDB) Then
        'SetDai.Edit
            Data_To_Set 0        ' moves screen values to dynaset.
            UpdateEventdb "Daily", ObjTour.DBGetField("Date", iDailyDB), ObjTour.DBGetField("DayType", iDailyDB), 0, 0, ""
            ObjTour.Update iDailyDB
        End If
End If

 Changed = 0
' --------------------
' Update screen incase
' of peak change
' --------------------
        Data_To_Scr -1
' --------------------------------
' Check Previous and
' next button for enabled setting
' --------------------------------
DateStr = "Date = " & "#" & Format$(DaiDatTxt.Text, objMdi.info.cSHORTDATEFORMAT) & "#"
ObjTour.DBFindFirst DateStr, iDailyDB

Screen.MousePointer = 0
Changed = 0
Exit Sub
SaveError:
    Select Case Err
        Case 3022:
            
        Case Else:
            MsgBox Error$(Err)
    End Select
    Resume Next
End Sub

Private Sub DaiRatTxt_ToolTip()
    
    ShowToolTipInStatusBar DaiRatTxt.ToolTipText

End Sub

Private Sub DaiSlpTxt_ToolTip()

    ShowToolTipInStatusBar DaiSlpTxt.ToolTipText

End Sub


Private Sub DaiWeiTxt_KeyPress(KeyAscii As Integer)
    KeyAscii = TextValidate(KeyAscii, "Real Only", DaiWeiTxt)
End Sub

Private Sub DaiWeiTxt_LostFocus()
'
' check if change was made
'
On Local Error Resume Next
If Val(DaiWeiTxt) <> ObjTour.DBGetField("Weight", iDailyDB) Then Changed = -1

End Sub


Private Sub Data1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Button = " & CStr(Button)
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
MsgBox "Action = " & CStr(Action)
End Sub

Private Sub DaiWeiTxt_ToolTip()
    ShowToolTipInStatusBar DaiWeiTxt.ToolTipText
End Sub

Private Sub Form_Activate()
    Define_DailyFrm_mnu Loadmnu
End Sub

Private Sub Form_Deactivate()
    Define_DailyFrm_mnu Unloadmnu
End Sub


Private Sub Form_GotFocus()
If bDebug Then Handle_Err 0, "FormGotFocus-Daily"

DailyFrm.WindowState = 0
CentreForm DailyFrm, -1
LoadVisible

End Sub

Sub Define_DailyFrm_mnu(LoadType As String)
On Local Error GoTo Define_DailyFrm_Err
' ------------------------------
' Turn Key Preview on
' to trap esc key press among
' other key presses
' ------------------------------
DailyFrm.KeyPreview = True
If Not objMdi.info.NewUser Then
Select Case LoadType
    Case Unloadmnu:
        MDI!MdiEximnu.Caption = MdiFrm_Exit
        MDI!MdiOptmnu.Caption = MdiFrm_Option
        MDI!MdiNewmnu.Caption = MdiFrm_Newmnu
        MDI!MdiSavmnu.Caption = MdiFrm_Savmnu
        MDI!MdiDelmnu.Caption = MdiFrm_Delmnu
        MDI!MdiOptmnu.Enabled = True
        MDI!MdiDiamnu.Enabled = True
        MDI!MdiNewmnu.Enabled = True
        MDI!MdiSavmnu.Enabled = False
        MDI!MdiDelmnu.Enabled = False
        MDI!MdiPrimnu.Enabled = False
    Case Loadmnu:
        MDI!MdiEximnu.Caption = DailyFrm_Exit
        MDI!MdiOptmnu.Caption = DailyFrm_Option
        MDI!MdiNewmnu.Caption = DailyFrm_Newmnu
        MDI!MdiSavmnu.Caption = DailyFrm_Savmnu
        MDI!MdiDelmnu.Caption = DailyFrm_Delmnu
        MDI!MdiOptmnu.Enabled = True
        MDI!MdiDiamnu.Enabled = False
        MDI!MdiNewmnu.Enabled = True
        MDI!MdiSavmnu.Enabled = True
        MDI!MdiDelmnu.Enabled = True
        MDI!MdiPrimnu.Enabled = True
        
End Select
End If
Exit Sub
Define_DailyFrm_Err:
    If bDebug Then Handle_Err Err, "Define_DailyFrm_mun-DailyFrm"
    Resume Next
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    ' ------------------------------------------
    ' If user presses ESC key, exit sub.
    ' Dailyfrm.unload will first check if
    ' a change has been made and prompt user
    ' to save changes before unload daily form.
    ' ------------------------------------------
    If KeyAscii = vbKeyEscape Then Unload Me
    
End Sub

Private Sub Form_Load()
'----------------------------------------
' When this form is loaded
' the current date will always
' be current system date.
'----------------------------------------
On Local Error GoTo DaiError
Dim DateStr As String
Dim sSQL As String

Changed = 0
m_loading = True

If bDebug Then Handle_Err 0, "FormLoad-Daily"
Define_DailyFrm_mnu Loadmnu
CentreForm DailyFrm, -1
LoadVisible
Screen.MousePointer = 11                        'changes mouse to hourglass

' Set Recordset
sSQL = " SELECT * " & _
       " FROM " & gcDai_Tour_Dai & _
       " WHERE Id = " & objMdi.info.ID & " AND Type = " & CStr(gcDAI_NORMAL) & _
       " ORDER BY Date ASC"
       
ObjTour.RstSQL iDailyDB, sSQL

' Set FindFirst Statement
DateStr = "Date = " & "#" & Format$(Date, "mm-dd-yyyy") & "#"
ObjTour.DBFindFirst DateStr, iDailyDB

DoEvents
LoadDayType

' Any record found in FindFirst
If ObjTour.NoMatch(iDailyDB) Then    ' Check if record is found.
    Load_Levels -1              ' 0 = day dosen't exist.
    Data_To_Scr (0)             ' prompt for change...
    If ObjTour.RstRecordCount(iDailyDB) = 0 Then DaiPreCmd.Enabled = 0
    DaiNexCmd.Enabled = 0
Else
    Load_Levels 0
    Data_To_Scr (1)             ' 1 = day exist.
End If

' ------------------------
' Reset runtime variables
' ------------------------
Screen.MousePointer = 0                         'changes mouse to default
m_loading = False
Changed = 0

Exit Sub
DaiError:
     If bDebug Then Handle_Err Err, "Load-DailyFrm"
    Screen.MousePointer = 0
    Resume Next
End Sub


Private Sub Form_LostFocus()

'dbDai.Close

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then DailyFrm.PopupMenu MDI.DaiLevmnu
End Sub


Private Sub Form_Resize()
On Error Resume Next
If MDI.Width <= 7664 Then MDI.Width = 7665
If MDI.Height <= 5565 Then MDI.Height = 5565
End Sub


Private Sub Form_Unload(Cancel As Integer)
' -----------------
' Check for changes
' -----------------
If bDebug Then Handle_Err 0, "FormunLoad-DailyFrm"
    If Changed Then
        If vbYes = MsgBox("Save Changes?", 36, " Save ") Then DaiSavmnu
    End If
Define_DailyFrm_mnu Unloadmnu
Unload DailyFrm

End Sub


Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then DailyFrm.PopupMenu MDI.DaiLevmnu
End Sub

'---------------------------------------------------------------------------------------
' PROCEDURE : ShowToolTipInStatusBar
' DATE      : 9/4/04 08:38
' Author    : Mark Ormesher
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub ShowToolTipInStatusBar(ByVal sToolTip As String)

On Local Error GoTo ShowToolTipInStatusBar_Error
'Declare local variables

    MDI.StatusBar1.Panels(1).Text = sToolTip

On Error GoTo 0
Exit Sub

ShowToolTipInStatusBar_Error:
If bDebug Then Handle_Err Err, "ShowToolTipInStatusBar-DailyFrm"
Resume Next

End Sub
