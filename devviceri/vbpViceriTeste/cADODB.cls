VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private cDB As ADODB.Connection

Dim strSQL As String

'local variable(s) to hold property value(s)
Private mvarConnectionString As String 'local copy

Public Function ExecSQL(pstrSQL As String) As Boolean
ExecSQL = False
'If OpenConnection Then
    cDB.Execute pstrSQL
'End If
ExecSQL = True
End Function

Public Function OpenRSet(pstrSQL As String) As ADODB.Recordset
'Set cDB = New ADODB.Connection
On Error GoTo ADOERROR__
Dim x As New ADODB.Recordset
'If OpenConnection Then
    x.Open pstrSQL, cDB, adOpenStatic, adLockReadOnly
    Set OpenRSet = x
'End If
    Exit Function

ADOERROR__:
Dim strErr As String
    Select Case Err
        Case adErrObjectOpen
            strErr = "Error #" & Err.Number & ": " & Err.Description & vbCrLf
            strErr = strErr & "Error reported by: " & Err.Source & vbCrLf
            strErr = strErr & "Help File: " & Err.HelpFile & vbCrLf
            strErr = strErr & "Topic ID: " & Err.HelpContext
            MsgBox strErr
            'Debug.Print strErr
            Err.Clear
            Resume Next
        ' If some other error occurs that
        ' has nothing to do with ADO, show
        ' the number and description and exit.
        Case Else
            strErr = "Error #" & Err.Number & ": " & Err.Description & vbCrLf
            MsgBox strErr
            'Debug.Print strErr
            'Unload Me
    End Select
    
'    Dim Errs1 As cDB.Errors
'    Set Errs1 = cDB.Errors
'     For Each errLoop In Errs1
'          With errLoop
'            StrTmp = StrTmp & vbCrLf & "Error #" & i & ":"
'            StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
'            StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
'            StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
'            i = i + 1
'       End With
'    Next

    
    
End Function

Public Function OpenConnection() As Boolean
OpenConnection = False
Set cDB = New ADODB.Connection

cDB.ConnectionString = "Provider=SQLNCLI11;Server=LAPTOP-UPBU2BQP;Database=DBViceri;Trusted_Connection=yes;timeout=30;"
cDB.Open
cDB.IsolationLevel = adXactSerializable
strSQL = "SET DATEFORMAT DMY"
cDB.Execute strSQL
OpenConnection = True

End Function

Public Property Let ConnectionString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnectionString = 5
    mvarConnectionString = vData
End Property


Public Property Get ConnectionString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.ConnectionString
    ConnectionString = mvarConnectionString
End Property



Private Sub Class_Initialize()
Dim k As Boolean
k = OpenConnection()
End Sub
