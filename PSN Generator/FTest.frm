VERSION 5.00
Begin VB.Form FTest 
   Caption         =   "Test Registration"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test License Key"
      Height          =   495
      Left            =   300
      TabIndex        =   7
      Top             =   1560
      Width           =   2115
   End
   Begin VB.TextBox txtPart4 
      Height          =   315
      Left            =   2580
      TabIndex        =   6
      Top             =   1140
      Width           =   615
   End
   Begin VB.TextBox txtPart3 
      Height          =   315
      Left            =   1860
      TabIndex        =   5
      Top             =   1140
      Width           =   615
   End
   Begin VB.TextBox txtPart2 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   1140
      Width           =   615
   End
   Begin VB.TextBox txtPart1 
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   1140
      Width           =   615
   End
   Begin VB.TextBox txtUserSpecifiedPart1 
      Height          =   315
      Left            =   3180
      TabIndex        =   1
      Text            =   "00101"
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate License Key"
      Height          =   495
      Left            =   300
      TabIndex        =   2
      Top             =   540
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Part 1 of Key (5 Numeric Characters):"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FTest
' FILENAME:     C:\My Code\vb\Registration\FTest.frm
' AUTHOR:       Phil Fresle
' CREATED:      06-Sep-2000
' COPYRIGHT:    Copyright 2000 Frez Systems Limited.
'
' DESCRIPTION:
' Used to test class for generating and testing license keys.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
'
' MODIFICATION HISTORY:
' 1.0       06-Sep-2000
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

'*******************************************************************************
' cmdGenerate_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Test key generation
'*******************************************************************************
Private Sub cmdGenerate_Click()

Dim oRegistration   As CRegistration
Dim sKey            As String
Dim lLoop           As Long
Dim sPSN            As String

Set oRegistration = New CRegistration

For lLoop = 1 To 25
sKey = oRegistration.GenerateKey(txtUserSpecifiedPart1.Text)

sPSN = sPSN & Left(sKey, 5)
sPSN = sPSN & Mid(sKey, 6, 4)
sPSN = sPSN & Mid(sKey, 10, 4)
sPSN = sPSN & Mid(sKey, 14, 4) & vbCrLf

Next lLoop

'Show last PSN creation
txtPart1.Text = Left(sKey, 5)
txtPart2.Text = Mid(sKey, 6, 4)
txtPart3.Text = Mid(sKey, 10, 4)
txtPart4.Text = Mid(sKey, 14, 4)

'Append to clipboard
Clipboard.Clear
Clipboard.SetText sPSN
MsgBox "PSN are in clipboard"

Set oRegistration = Nothing
End Sub

'*******************************************************************************
' cmdTest_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Test key validation
'*******************************************************************************
Private Sub cmdTest_Click()
    Dim oRegistration   As CRegistration
    Dim sKey            As String
    
    sKey = Trim(txtPart1.Text) & Trim(txtPart2.Text) _
        & Trim(txtPart3.Text) & Trim(txtPart4.Text)
        
    Set oRegistration = New CRegistration
    
    If oRegistration.IsKeyOK(sKey) Then
        MsgBox "Key is valid"
    Else
        MsgBox "Key is NOT valid"
    End If
    
    Set oRegistration = Nothing
End Sub
