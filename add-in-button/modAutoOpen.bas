Attribute VB_Name = "modAutoOpen"
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Adds 3 temporary buttons to ribbon in "Add-ins" tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Private Const gToolBarName As String = "add-in-button"

Sub Auto_Open()
    Call SetupToolBar
End Sub

'Sub Auto_Close()
'    Call DeleteCustomCommandBar
'End Sub

Sub SetupToolBar()

Dim ScheduleCommandBar As CommandBar
Dim cmdBarCntrl As CommandBarControl

    'Call DeleteCustomCommandBar
    'Set ScheduleCommandBar = Application.CommandBars.Add(Name:=gToolBarName & ThisWorkbook.FullName, Temporary:=True)
    Set ScheduleCommandBar = Application.CommandBars.Add(Name:=ThisWorkbook.FullName, Temporary:=True)
    
    ScheduleCommandBar.Visible = True
    
    Set cmdBarCntrl = ScheduleCommandBar.Controls.Add(1, 1953, , , True)
    cmdBarCntrl.Caption = "Init Sheet"
    cmdBarCntrl.TooltipText = cmdBarCntrl.Caption & ". Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna."
    cmdBarCntrl.OnAction = "Method1"
    cmdBarCntrl.Style = MsoButtonStyle.msoButtonIconAndCaption

    Set cmdBarCntrl = ScheduleCommandBar.Controls.Add(1, 1950, , , True)
    cmdBarCntrl.Caption = "Calculate"
    cmdBarCntrl.TooltipText = cmdBarCntrl.Caption & ". Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna."
    cmdBarCntrl.OnAction = "Method2"
    cmdBarCntrl.Style = MsoButtonStyle.msoButtonIconAndCaption

    Set cmdBarCntrl = ScheduleCommandBar.Controls.Add(1, 1154, , , True)
    cmdBarCntrl.Caption = "Export"
    cmdBarCntrl.TooltipText = cmdBarCntrl.Caption & ". Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna."
    cmdBarCntrl.OnAction = "Method3"
    cmdBarCntrl.Style = MsoButtonStyle.msoButtonIconAndCaption
End Sub

'Sub DeleteCustomCommandBar()
'    On Error Resume Next
'    Application.CommandBars(gToolBarName).Delete
'    On Error GoTo 0
'End Sub

Public Sub Method1()
    MsgBox "Init Sheet"
End Sub

Public Sub Method2()
    MsgBox "Calculate"
End Sub

Public Sub Method3()
    MsgBox "Export"
End Sub

