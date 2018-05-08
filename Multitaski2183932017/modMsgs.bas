Attribute VB_Name = "modMsgs"
#Const modMsgs = True
Option Explicit

Option Private Module

Public Type ProcInfo
    Proc As Long
    hwnd As Long
    uMsg As Long
    wArg As Long
    lArg As Long
    addr As Long
End Type

Private MessagePeak As Long
Public MessageCount As Long
Public MessageQueue() As ProcInfo
Public ReturnScenario As Long

Public Sub WindowsYield()
    Static wMsg As msg
    If PeekMessage(wMsg, 0, 0, 0, PM_NOREMOVE) Then
        Do
            TranslateMessage wMsg
            DispatchMessage wMsg
        Loop While PeekMessage(wMsg, 0, 0, 0, PM_REMOVE + PM_NOYIELD)
    Else
        Sleep 1
    End If
End Sub

Public Sub PrintMessage(ByVal txt As String, Optional ByVal newline As Boolean = True)
    If frmMain.Visible Then
        frmMain.PrintMessage txt, newline
    End If
End Sub

Public Sub AddMessage(ByVal uMsg As Long, ByVal wArg As Long, ByVal lArg As Long, ByVal addr As Long)
    Static lstMsg As ProcInfo
    If Not Programs Is Nothing Then
        Dim StopHandle As Long
        StopHandle = Programs.Handle
        If Not (lstMsg.uMsg = uMsg And lstMsg.wArg = wArg And lstMsg.lArg = lArg And lstMsg.addr = addr) Then
            Do
                If Programs.Accepts(uMsg) Then
                    
                    If Not (lstMsg.hwnd = Programs.Handle And lstMsg.uMsg = uMsg And lstMsg.wArg = wArg And lstMsg.lArg = lArg And lstMsg.addr = addr) Then
                        MessageCount = MessageCount + 1
                        ReDim Preserve MessageQueue(1 To MessageCount) As ProcInfo
                        MessageQueue(MessageCount).Proc = MessagePeak
                        MessageQueue(MessageCount).hwnd = Programs.Handle
                        MessageQueue(MessageCount).uMsg = uMsg
                        MessageQueue(MessageCount).wArg = wArg
                        MessageQueue(MessageCount).lArg = lArg
                        MessageQueue(MessageCount).addr = addr
                        lstMsg.Proc = MessagePeak
                        lstMsg.hwnd = Programs.Handle
                        lstMsg.uMsg = uMsg
                        lstMsg.wArg = wArg
                        lstMsg.lArg = lArg
                        lstMsg.addr = addr
                    End If
                    
                End If
                ShiftPrograms
            Loop Until Programs.Handle = StopHandle
        End If
    End If
End Sub

Public Sub DelMessage(ByVal Number As Long)
    If (MessageCount > 0) Then
        Dim i As Long
        If (Number <= (MessageCount - 2)) And (Number > 0) Then
            For i = Number To MessageCount - 1
                MessageQueue(i) = MessageQueue(i + 1)
            Next
        End If
        MessageCount = MessageCount - 1
        If MessageCount = 0 Then
            Erase MessageQueue
        Else
            ReDim Preserve MessageQueue(1 To MessageCount) As ProcInfo
        End If
        
    End If
End Sub

Public Sub Initialize()
    ReturnScenario = 4
    MessagePeak = 1
End Sub

Public Sub Terminate()
    Erase MessageQueue
End Sub

Public Sub ProcessOrdered()
    Static i As Long
    If Not Programs Is Nothing Then
        i = i + MessagePeak
        If (i > MessageCount) Then i = 1
        Dim StopHandle As Long
        StopHandle = Programs.Handle
        Do
'            i = 1
'            Do Until (i > MessageCount)
                If (i <= MessageCount) And (i > 0) Then
                   ' If MessageQueue(i).hwnd = Programs.Handle Then
                        Select Case MessageQueue(i).Proc
                            Case 0
                                DelMessage i
                            Case Else
                                MessageQueue(i).Proc = HandleWindowProc(MessageQueue(i).Proc, MessageQueue(i).hwnd, _
                                    MessageQueue(i).uMsg, MessageQueue(i).wArg, MessageQueue(i).lArg, MessageQueue(i).addr)
                                'i = i + MessagePeak
                        End Select
'                    Else
'                        i = i + MessagePeak
'                    End If
                Else
                    'i = i + MessagePeak
                End If
                i = i + MessagePeak
'            Loop
            ShiftPrograms
        Loop Until Programs.Handle = StopHandle
    End If
End Sub
Public Function HandleWindowProc(ByVal Proc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wArg As Long, ByVal lArg As Long, ByVal addr As Long) As Long
    If (addr <= 4) Then addr = Val(AddressOf DefaultWindowProc)
    RtlMoveMemory ByVal VarPtr(HandleWindowProc), CallWindowProc(addr, hwnd, uMsg, wArg, lArg), 4&
End Function

Public Function DefaultWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wArg As Long, ByVal lArg As Long) As Long
    'return 1 to specify failure; return 0 to specify success; return -1 to specity unhandle and use default;
    Select Case ReturnScenario
        Case 3
            DefaultWindowProc = 0
        Case 4
            DefaultWindowProc = Round(Rnd, 0)
        Case Else
            DefaultWindowProc = 1
    End Select
    Select Case DefaultWindowProc
        Case 1
            PrintMessage "DefaultHandler(" & hwnd & ", " & uMsg & ", " & wArg & ", " & lArg & ") = Failure"
        Case Else
            PrintMessage "DefaultHandler(" & hwnd & ", " & uMsg & ", " & wArg & ", " & lArg & ") = Success"
    End Select
End Function

Public Function CustomWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wArg As Long, ByVal lArg As Long) As Long
    'return 1 to specify failure; return 0 to specify success; return -1 to specity unhandle and use default;
    Select Case ReturnScenario
        Case 1
            CustomWindowProc = -1
        Case 2
            CustomWindowProc = 1
        Case 3
            CustomWindowProc = 0
        Case 4
            CustomWindowProc = Round(Rnd, 0)
    End Select
    Select Case CustomWindowProc
        Case 1
            PrintMessage "CustomHandler(" & hwnd & ", " & uMsg & ", " & wArg & ", " & lArg & ") = Failure"
        Case 0
            PrintMessage "CustomHandler(" & hwnd & ", " & uMsg & ", " & wArg & ", " & lArg & ") = Success"
        Case -1
            PrintMessage "CustomHandler(" & hwnd & ", " & uMsg & ", " & wArg & ", " & lArg & ") = Default"
    End Select
End Function


