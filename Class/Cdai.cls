VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mDaiVar As cDaiVar
Public Property Set info(pDaiVar As cDaiVar)
Set mDaiVar = pDaiVar
End Property


Public Property Get info() As cDaiVar
Set info = mDaiVar
End Property


Public Function Load_Daily_Settings() As Boolean
On Local Error GoTo Load_Daily_Error
Dim lDailyOptions   As Long

Load_Daily_Settings = False

lDailyOptions = objMdi.info.DaiOption

With objDai.info
    .Desc = IsBitSet(lDailyOptions, BitFlags.Daily_Description)
    .RestH = IsBitSet(lDailyOptions, BitFlags.Daily_RestHeart)
    .Sleep = IsBitSet(lDailyOptions, BitFlags.Daily_Sleep)
    .DayR = IsBitSet(lDailyOptions, BitFlags.Daily_RestHeart)
    .Weight = IsBitSet(lDailyOptions, BitFlags.Daily_Weight)
    .Metric = IsBitSet(lDailyOptions, BitFlags.Daily_Metric)
End With

Load_Daily_Settings = True

Exit Function
Load_Daily_Error:
     If bDebug Then
        MsgBox Error$(Err)
        Handle_Err Err, "Load_Daily_Settings-cDai"
     End If
     Exit Function
End Function
