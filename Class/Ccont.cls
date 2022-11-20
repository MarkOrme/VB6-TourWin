VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mContVar As cContVar

Public Property Set info(pContVar As cContVar)
Set mContVar = pContVar
End Property

Public Property Get info() As cContVar
Set info = mContVar
End Property

