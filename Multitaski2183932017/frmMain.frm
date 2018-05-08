VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Debug and Statistics"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   709
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rows As Long
Private msg As String
Public Registers As String
Public Sub PrintPrograms()
 
    frmMain.Cls
    Dim str As String
    If Not Programs Is Nothing Then
        
        Dim StopHandle As Long
        StopHandle = Programs.Handle
        Do
            str = str & " " & Format(Programs.Handle, "0000")
            ShiftPrograms
        Loop Until Programs.Handle = StopHandle
        
    End If
    
    str = "This simulates a multitasking message handling system with test response handlers by creating" & vbCrLf & _
            """Program"" Objects, much like a ""Window"" with subclassing, handling timer and mouse message events." & vbCrLf & vbCrLf & _
            "Esc = Exits (if this is open), F12 = Hide/Show, Insert/Delete = Change Handles, Enter = Clear" & vbCrLf & _
            "Num0 = Function, Num1 = Default, Num2 = Failure, Num3 = Success, Num4 = Random; " & vbCrLf & vbCrLf & _
            "Messages: " & Format(MessageCount, "000000") & "; Handles:" & str & ";" & vbCrLf & vbCrLf

    Dim h As Single
    Dim w As Single
    rows = 0
    
    h = 0
    Do Until str = ""
        w = 0
        Do While (w < frmMain.ScaleWidth) And (Not (str = ""))
            w = w + frmMain.TextWidth(Left(str, 1))
            If Left(str, 2) = vbCrLf Then
                str = Mid(str, 3)
                Exit Do
            Else
                Me.Print Left(str, 1);
                str = Mid(str, 2)
            End If
            
        Loop
        Me.Print
        h = h + frmMain.TextHeight("X")
    Loop
    If str <> "" Then
        Me.Print str
        h = h + frmMain.TextHeight("X")
        str = ""
    End If

    Do While h < frmMain.ScaleHeight
        h = h + frmMain.TextHeight("X")
        rows = rows + 1
    Loop

    Do While CountWord(msg, vbCrLf) >= rows
        RemoveNextArg msg, vbCrLf
    Loop
    
    Me.Print msg

End Sub

Public Sub PrintMessage(ByVal txt As String, Optional ByVal newline As Boolean = True)

    msg = msg & txt & IIf(newline, vbCrLf, "")
       
End Sub

Public Function CountWord(ByVal Text As String, ByVal Word As String, Optional ByVal Exact As Boolean = True) As Long
    Dim cnt As Long
    cnt = UBound(Split(Text, Word, , IIf(Exact, vbBinaryCompare, vbTextCompare)))
    If cnt > 0 Then CountWord = cnt
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
    If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
        RemoveNextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
        TheParams = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator)))
    Else
        RemoveNextArg = Trim(TheParams)
        TheParams = ""
    End If
End Function
