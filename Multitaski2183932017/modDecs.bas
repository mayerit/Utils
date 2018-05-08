Attribute VB_Name = "modDecs"
Option Explicit

Option Private Module

Public Const WM_USER = &H400
Public Const WM_TIMER = &H113
Public Const WM_MOUSEMOVE = &H200

Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Public Const PM_NOYIELD = &H2

Public Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function VarPtrArray Lib "MSVBVM60.DLL" Alias "VarPtr" (Var() As Any) As Long

Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Property Let BitLong(ByRef Word As Long, ByRef Bit As Long, ByRef Value As Boolean)
    If (Word And (Bit)) And (Not Value) Then
        Word = Word - (Bit)
    ElseIf (Not (Word And (Bit))) And Value Then
        Word = Word Or (Bit)
    End If
End Property

Public Property Get BitLong(ByRef Word As Long, ByRef Bit As Long) As Boolean
    BitLong = (Word And (Bit))
End Property

Public Sub SwapObjects(ByRef obj1 As Object, ByRef obj2 As Object)
    Static obj3 As Object
    vbaObjSetAddref obj3, ObjPtr(obj1)
    vbaObjSetAddref obj1, ObjPtr(obj2)
    vbaObjSetAddref obj2, ObjPtr(obj3)
    vbaObjSet obj3, ObjPtr(Nothing)
End Sub
