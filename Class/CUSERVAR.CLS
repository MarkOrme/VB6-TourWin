VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUserVar"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' --------------------------------------------
' Global class variable for user definiation
' like form options, passwords, user name etc.
' --------------------------------------------

Public Name As String
Public PassWord As String
Public DataPath As String
Public Security As Boolean
Public Load As Integer
Public MetaFile As String

