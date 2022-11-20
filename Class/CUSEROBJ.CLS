VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUserObj"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' -------------------------------------------
' This class uses cUserVar's public variables
' -------------------------------------------
Private mUserVar As cUserVar

Public Property Set UserMdi(pUserMdi As cUserVar)
Set mUserVar = pUserMdi
End Property

Public Property Get UserMdi() As cUserVar
Set UserMdi = mUserVar
End Property


