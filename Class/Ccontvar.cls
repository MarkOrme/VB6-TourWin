VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cContVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IndxOrdr As String
Public ClmnOrdr As String
Public SrtOrdr As String
Public ContactWidth As Single
Public LastWidth As Single
Public FirstWidth As Single
Public PhoneWidth As Single
Public FaxWidth As Single
Public EMailWidth As Single


Private Sub Class_Initialize()
' -----------------------------
' Set the default width values
' -----------------------------
On Local Error GoTo Init_Err

Me.ContactWidth = 1000
Me.LastWidth = 1000
Me.FirstWidth = 1000
Me.PhoneWidth = 1000
Me.FaxWidth = 1000
Me.EMailWidth = 1000
Me.IndxOrdr = "Contact"
Me.ClmnOrdr = "0123456"
Me.SrtOrdr = "ASC"

Exit Sub
Init_Err:
    MsgBox Error$(Err)

End Sub
