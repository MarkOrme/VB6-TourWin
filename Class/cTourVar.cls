VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTourVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mcSaveValueToReg = 2  '

Public Function WriteFormSettingsToReg(frmForm As Form) As Boolean
' ======================================================================
' WriteFormSettingToReg -
' Purpose, write form controls values to registry. So next time form is
' loaded a companion method - ReadFormSettingsFromReg - can be used to
' define form control values as there where when the user last left.
'
' This method loops through each control on passed form, for each
' control it's tag property is read. If the value after the first
' '|' pipe character is -1, then controls text value is written
' to registry.
'
' Registry Location:
'    gcRegTourKey\gcFormsRegSubKey\FormName\controls="Value"
' ======================================================================
Dim oControl As Control
Dim sFormName As String
On Local Error GoTo Settings_Err

' Get Form Name
sFormName = frmForm.Name

' Loop through each control
For Each oControl In frmForm

  If 0 <> InStr(1, oControl.Tag, "|", vbTextCompare) Then
    
  ' check if value should be saved to Reg...
    If -1 = Val(GetTagsSettingValue(oControl.Tag, mcSaveValueToReg)) Then
    
         Write_Tour_Setting_To_Registry sFormName, oControl.Name, oControl.Text
         
    End If
  End If
Next oControl
Exit Function

Settings_Err:
    If bDebug Then Handle_Err Err, "WriteFormSettingsToReg-cTourVar"
    Err.Clear
    Resume Next
    
End Function

Private Function GetTagsSettingValue(sTagValue As String, iTagType As Integer)
' =========================================================
' GetFormsSettingValue -
'   Purpose to parse the passed sTagValue for the word after
'   the first | character.
' ==========================================================
Dim iPipePos As Integer

' Check if '|' character exists
iPipePos = InStr(1, sTagValue, "|", vbTextCompare)

If 0 <> iPipePos Then

 If iTagType = mcSaveValueToReg Then
    GetTagsSettingValue = Mid(sTagValue, iPipePos + 1)
 End If
    
End If

End Function
Public Function ReadFormSettingsFromReg(frmForm As Form) As Boolean
' ======================================================================
' WriteFormSettingToReg -
' Purpose, write form controls values to registry. So next time form is
' loaded a companion method - ReadFormSettingsFromReg - can be used to
' define form control values as there where when the user last left.
'
' This method loops through each control on passed form, for each
' control it's tag property is read. If the value after the first
' '|' pipe character is -1, then controls text value is written
' to registry.
'
' Registry Location:
'    gcRegTourKey\gcFormsRegSubKey\FormName\controls="Value"
' ======================================================================
Dim oControl As Control
Dim sFormName As String
Dim sKeyValue As String
Dim iGetCBIndexFromString As Long

On Local Error GoTo ReadSettings_Err

' Get Form Name
sFormName = frmForm.Name

' Loop through each control
For Each oControl In frmForm

  If 0 <> InStr(1, oControl.Tag, "|", vbTextCompare) Then
    
  ' check if value should be saved to Reg...
    If -1 = Val(GetTagsSettingValue(oControl.Tag, mcSaveValueToReg)) Then
    
       If Read_Tour_Setting_From_Registry(sFormName, oControl.Name, sKeyValue) Then
            If REG_ERROR <> sKeyValue Then
                Debug.Print "Control Name: " & oControl.Name & " Value: " & sKeyValue
                If TypeOf oControl Is ListBox Or TypeOf oControl Is ComboBox Then
                    ' Try and find sSetAsSelected
                    iGetCBIndexFromString = FindItemListControl(oControl, sKeyValue)
                    If -1 <> iGetCBIndexFromString Then
                        oControl.ListIndex = iGetCBIndexFromString
                    End If
                    
                Else
                    oControl.Text = sKeyValue
                End If
                
            End If
            
       End If
    End If
  End If
Next oControl

Exit Function

ReadSettings_Err:
    If bDebug Then Handle_Err Err, "ReadFormSettingsToReg-cTourVar"
    Err.Clear
    Resume Next
    
End Function

