Attribute VB_Name = "Module1"

Public Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long
    
Public Declare Function OpenProcess Lib "Kernel32.dll" _
    (ByVal dwDesiredAccessas As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcId As Long) As Long

Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Sub TerminateProc(ByVal TargetID As Long)
On Error Resume Next

Dim TerminateProcHandle As Long
Dim RetVal              As Long

    TerminateProcHandle = OpenProcess(PROCESS_ALL_ACCESS, _
    True, _
    TargetID)

    RetVal = TerminateProcess(TerminateProcHandle, _
    0&)

End Sub

