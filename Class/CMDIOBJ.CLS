VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMDI"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
' -------------------------------------------
' This class uses cUserVar's public variables
' -------------------------------------------
Private mMDIVar As cMDIVar

Public Property Set Info(pMDIVar As cMDIVar)
Set mMDIVar = pMDIVar
End Property

Public Property Get Info() As cMDIVar
Set Info = mMDIVar
End Property


