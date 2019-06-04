VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCorretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarIdCorretor As Double 'local copy
Private mvarCodigo As String 'local copy
Private mvarNome As String 'local copy
Private mvarCPF As String 'local copy

Private Function ExisteCPF(pADODB As cADODB, pCPF As String) As Boolean
Dim strSQL As String

ExisteCPF = False
strSQL = "select cpf from cliente where cpf = '" & pCPF & "' union select cpf from corretor where cpf = '" & pCPF & "'"
Dim oRSET As ADODB.Recordset
Set oRSET = pADODB.OpenRSet(strSQL)
If Not oRSET.EOF Then
    ExisteCPF = True
End If

End Function


Public Function AddCorretor(pADODB As cADODB) As Double
Dim strSQL As String
Dim dblLastID As Double
Dim oRSET As ADODB.Recordset

'*** valida se cpf ja existe em cliente ou corretor
If ExisteCPF(pADODB, mvarCPF) Then
    MsgBox "CPF: " & Format(mvarCPF, "@@@.@@@.@@@-@@") & " ja existe em Cliente ou Corretor.", vbInformation
    AddCorretor = -1
    Exit Function
End If


strSQL = "INSERT INTO dbo.Corretor (Codigo, Nome, CPF) VALUES ("
    strSQL = strSQL & "'" & mvarCodigo & "',"
    strSQL = strSQL & "'" & mvarNome & "',"
    strSQL = strSQL & "'" & mvarCPF & "')"
    pADODB.ExecSQL strSQL
    Set oRSET = pADODB.OpenRSet("select @@IDENTITY")
    If Not oRSET.EOF Then
        dblLastID = oRSET(0).Value
        mvarIdCorretor = oRSET(0).Value
        AddCorretor = oRSET(0).Value
    End If
    
End Function

Public Function ToList(pADODB As cADODB) As ADODB.Recordset
Dim strSQL As String
Dim oRSET As ADODB.Recordset

strSQL = "select idcorretor, codigo, nome, cpf from corretor order by nome  "
Set ToList = pADODB.OpenRSet(strSQL)

End Function






Public Property Let CPF(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPF = 5
    mvarCPF = vData
End Property


Public Property Get CPF() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.CPF
    CPF = mvarCPF
End Property



Public Property Let Nome(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nome = 5
    mvarNome = vData
End Property


Public Property Get Nome() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.Nome
    Nome = mvarNome
End Property



Public Property Let Codigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Codigo = 5
    mvarCodigo = vData
End Property


Public Property Get Codigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.Codigo
    Codigo = mvarCodigo
End Property



Public Property Let IdCorretor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IdCorretor = 5
    mvarIdCorretor = vData
End Property


Public Property Get IdCorretor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.IdCorretor
    IdCorretor = mvarIdCorretor
End Property



