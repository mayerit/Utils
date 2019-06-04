Attribute VB_Name = "GeneralModViceri"
Option Explicit

Public Sub RaiseMyError()
    MsgBox Err.Number & "-" & Err.Description & "-" & Err.Source, vbInformation
End Sub

Public Sub SendTabOnReturn(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If
End Sub

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub

Public Sub ResizeGrid(pGrid As MSFlexGrid, pForm As Form)
    Dim intRow As Integer
    Dim intCol As Integer
    
    With pGrid
        For intCol = 0 To .Cols - 1
            For intRow = 0 To .Rows - 1
                If .ColWidth(intCol) < pForm.TextWidth(.TextMatrix(intRow, intCol)) + 100 Then
                   .ColWidth(intCol) = pForm.TextWidth(.TextMatrix(intRow, intCol)) + 100
                End If
            Next
        Next
    End With
End Sub
