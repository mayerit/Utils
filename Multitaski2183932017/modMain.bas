Attribute VB_Name = "modMain"
#Const modMain = True
Option Explicit

Option Private Module

'################################
'the memory order of allocation in
'this first module is very sensitive
'to this properly moving the memory
'according to the main() and it's
'memory of variables allocated.
'#############################

Public Sub Main()
   
    Initialize
    ReadyInput
    InsertProgram
    Load frmMain
    frmMain.Show
        
    Do
        
        PrintPrograms
        ProcessOrdered
        LoopbackInput
        WindowsYield
        
    Loop Until EndOfInput Or (Forms.Count = 0)
    
    Do While (Forms.Count > 0)
        Unload Forms(0)
    Loop

    ClearPrograms
    Terminate

End Sub

Public Sub LoopbackInput()

    Static wmTimer As Single
    If wmTimer = 0 Or (Timer - wmTimer) >= 1 Then
        wmTimer = Timer
        AddMessage WM_TIMER, wmTimer, 0, 0
    End If
    
    If InputLoop Then
        Static lstMsg As ProcInfo
        If Toggled(VK_NUMPAD0) Then
            If Not (lstMsg.uMsg = WM_USER + WM_MOUSEMOVE And lstMsg.wArg = MouseXY.x And lstMsg.lArg = MouseXY.y And lstMsg.addr = Val(AddressOf CustomWindowProc)) Then
                AddMessage WM_USER + WM_MOUSEMOVE, MouseXY.x, MouseXY.y, AddressOf CustomWindowProc
                lstMsg.uMsg = WM_USER + WM_MOUSEMOVE
                lstMsg.wArg = MouseXY.x
                lstMsg.lArg = MouseXY.y
                lstMsg.addr = Val(AddressOf CustomWindowProc)
            End If
        Else
            If Not (lstMsg.uMsg = WM_USER + WM_MOUSEMOVE And lstMsg.wArg = MouseXY.x And lstMsg.lArg = MouseXY.y And lstMsg.addr = 0) Then
                AddMessage WM_USER + WM_MOUSEMOVE, MouseXY.x, MouseXY.y, 0
                lstMsg.uMsg = WM_USER + WM_MOUSEMOVE
                lstMsg.wArg = MouseXY.x
                lstMsg.lArg = MouseXY.y
                lstMsg.addr = 0
            End If
        End If
    End If
    
    If frmMain.Visible Then
    
        If Pressed(VK_NUMPAD1) Then
            ReturnScenario = 1
        End If
    
        If Pressed(VK_NUMPAD2) Then
            ReturnScenario = 2
        End If
    
        If Pressed(VK_NUMPAD3) Then
            ReturnScenario = 3
        End If
    
        If Pressed(VK_NUMPAD4) Then
            ReturnScenario = 4
        End If
        
        If Toggled(VK_RIGHT) Then
            ShiftPrograms
        ElseIf Pressed(VK_LEFT) Then
            ShiftPrograms
        End If
        
        If Pressed(VK_INSERT) Then
            InsertProgram
        End If
        If Pressed(VK_DELETE) Then
            DeleteProgram
        End If
        If Pressed(VK_RETURN) Then
            ClearPrograms
        End If
    Else
        Pressed(VK_ESCAPE) = False
        Toggled(VK_ESCAPE) = False
    End If
    
    If Pressed(VK_F12) Then
        If Not frmMain.Visible Then
            frmMain.Show
        Else
            frmMain.Hide
        End If
    End If
        
End Sub

