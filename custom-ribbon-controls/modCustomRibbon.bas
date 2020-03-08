Attribute VB_Name = "modCustomRibbon"
Option Explicit

Private m_Ribbon As IRibbonUI
Private m_BoolValue As Boolean


Public Sub customUI_onLoad(objRibbon As IRibbonUI)
    Set m_Ribbon = objRibbon
End Sub


Public Sub button1_onAction(control As IRibbonControl)
    m_BoolValue = Not m_BoolValue
    
    If (Not m_Ribbon Is Nothing) Then
        m_Ribbon.InvalidateControl "button2"
        m_Ribbon.InvalidateControl "checkBox2"
        m_Ribbon.InvalidateControl "toggleButton1"
    End If
End Sub
Public Sub button2_getLabel(control As IRibbonControl, ByRef returnedVal)
    If m_BoolValue Then
        returnedVal = "Enabled Button 2"
    Else
        returnedVal = "Disabled Button 2"
    End If
End Sub
Public Sub button2_getSize(control As IRibbonControl, ByRef returnedVal)
    If m_BoolValue Then
        returnedVal = 1 'large
    Else
        returnedVal = 0 'normal
    End If
End Sub
Public Sub button2_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If m_BoolValue Then
        returnedVal = 1 'Enabled
    Else
        returnedVal = 0 'Disabled
    End If
End Sub


Public Sub toggleButton1_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = m_BoolValue
End Sub
Public Sub toggleButton1_onAction(control As IRibbonControl, ByVal toggleButton1Value As Boolean)
    Debug.Print CStr(toggleButton1Value)
End Sub


Public Sub checkBox1_onAction(control As IRibbonControl, ByVal checkboxValue As Boolean)
    m_BoolValue = checkboxValue
    
    If (Not m_Ribbon Is Nothing) Then
        m_Ribbon.InvalidateControl "button2"
        m_Ribbon.InvalidateControl "checkBox2"
        m_Ribbon.InvalidateControl "toggleButton1"
    End If
End Sub
Public Sub checkBox2_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = m_BoolValue
End Sub


Public Sub editBox1_getText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "editBox_getText"
End Sub
Public Sub editBox1_onChange(control As IRibbonControl, ByVal editBox1Value)
    Debug.Print CStr(editBox1Value)
End Sub


Public Sub comboBox2_getItemCount(control As IRibbonControl, ByRef returnedVal)
    returnedVal = 5
End Sub
Public Sub comboBox2_getItemLabel(control As IRibbonControl, ByVal itemIndex As Integer, ByRef returnedVal)
    returnedVal = "ComboItem" & CStr(itemIndex)
End Sub
Public Sub comboBox2_getText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Default Text"
End Sub
Public Sub comboBox2_onChange(control As IRibbonControl, ByVal comboBox2Value As String)
    Debug.Print comboBox2Value
End Sub


Public Sub dropDown2_getItemCount(control As IRibbonControl, ByRef returnedVal)
    returnedVal = 5
End Sub
Public Sub dropDown2_getItemLabel(control As IRibbonControl, ByVal itemIndex As Integer, ByRef returnedVal)
    returnedVal = "DropDownItem" & CStr(itemIndex)
End Sub
Public Sub dropDown2_getSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = 1
End Sub
Public Sub dropDown2_onAction(control As IRibbonControl, ByVal controlId As String, ByVal itemIndex As Integer)
    Debug.Print controlId & " " & CStr(itemIndex)
End Sub
Public Sub dropDown2_getItemImage(control As IRibbonControl, ByVal itemIndex As Integer, ByRef returnedVal)
    Select Case itemIndex
    Case 0
        returnedVal = "Cut"
    Case 1
        returnedVal = "Copy"
    Case 2
        returnedVal = "Paste"
    Case 3
        returnedVal = "Delete"
    Case 4
        returnedVal = "SelectAll"
    End Select
End Sub


