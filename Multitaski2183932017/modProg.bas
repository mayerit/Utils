Attribute VB_Name = "modProg"
Option Explicit

Option Private Module

Public Reference As Long
Public Programs As Program
Public References As VBA.Collection

Public Property Get ProgramCount() As Long
    If Not References Is Nothing Then
        ProgramCount = References.Count
    Else
        ProgramCount = 0
    End If
End Property

Public Sub ClearPrograms()
    Do Until Programs Is Nothing
        DeleteProgram
    Loop
End Sub

Public Sub InsertProgram()
    If References Is Nothing Then Set References = New Collection
    
    Dim tmpObj As Program
    If Programs Is Nothing Then
        Set Programs = New Program
        Set tmpObj = Programs
    Else
        Set tmpObj = New Program
    End If
    References.Add ObjPtr(tmpObj), "h" & tmpObj.Handle
    Set tmpObj.Prior = Programs.Prior
    Set Programs.Prior.Forth = tmpObj
    Set tmpObj.Forth = Programs
    Set Programs.Prior = tmpObj
End Sub

Public Sub DeleteProgram()
    If Not Programs Is Nothing Then
        Dim tmpObj As Program
        References.Remove "h" & Programs.Handle
        If (Programs.Prior.Handle = Programs.Handle) Then
            Set Programs = Nothing
        Else
            Set Programs.Prior.Forth = Programs.Forth
            Set Programs.Forth.Prior = Programs.Prior
            Set Programs = Programs.Prior.Forth
        End If
    End If
End Sub

Public Static Sub ShiftPrograms()
    If Not Programs Is Nothing Then
        Set Programs = Programs.Forth
    End If
End Sub

Public Sub PrintPrograms()
    If frmMain.Visible Then
        frmMain.PrintPrograms
    End If
End Sub


