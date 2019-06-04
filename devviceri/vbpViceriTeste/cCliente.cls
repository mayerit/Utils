VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarIdCliente As Double 'local copy
Private mvarNome As String 'local copy
Private mvarCPF As String 'local copy
Private mvarEndereco As String 'local copy
Private mvarAtivo As Integer 'local copy
Private mvarCidadeID As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarCorretorID As Double 'local copy

Const MyContextID = 1010407    ' Define a constant for contextID.
Public Sub Teste(pSTR As String)
    On Error Resume Next
    Err.Raise vbObjectError + 1, App.Title, "This is a test of raising an error"
End Sub

Public Property Let CorretorID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CorretorID = 5
    mvarCorretorID = vData
End Property


Public Property Get CorretorID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.CorretorID
    CorretorID = mvarCorretorID
End Property


Public Function ToList(pADODB As cADODB, pFiltro As String) As ADODB.Recordset
Dim strSQL As String
Dim oRSET As ADODB.Recordset

strSQL = "select cliente.Nome, dbo.formatarCNPJCPF(cliente.CPF) as CPF, Ativo = CASE cliente.Ativo"
strSQL = strSQL & " WHEN 1 THEN 'Sim' "
strSQL = strSQL & " WHEN 0 THEN 'Não' "
strSQL = strSQL & " End, "
strSQL = strSQL & " Corretor.Nome as 'Nome Corretor', Corretor.Codigo as 'Código Corretor', UF.Nome as 'UF', Cidade.Nome as 'Cidade', 'x'"
    strSQL = strSQL & " from cliente  "
strSQL = strSQL & " inner join ClienteCorretor on cliente.IdCliente = ClienteCorretor.IdCliente "
strSQL = strSQL & " inner join Corretor on ClienteCorretor.IdCorretor = Corretor.IdCorretor "
strSQL = strSQL & " inner join Cidade on cliente.CidadeID = Cidade.ID "
strSQL = strSQL & " inner join UF on Cidade.IDUF = UF.ID"


If pFiltro <> vbNullString Then
    strSQL = strSQL & " where " & pFiltro
End If

Set ToList = pADODB.OpenRSet(strSQL)
End Function


Public Property Let CidadeID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CidadeID = 5
    mvarCidadeID = vData
End Property


Public Property Get CidadeID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.CidadeID
    CidadeID = mvarCidadeID
End Property



Public Property Let Ativo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ativo = 5
    mvarAtivo = vData
End Property


Public Property Get Ativo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.Ativo
    Ativo = mvarAtivo
End Property

Public Property Let Endereco(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Endereco = 5
    mvarEndereco = vData
End Property

Public Property Get Endereco() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.Endereco
    Let Endereco = mvarEndereco
End Property

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

Public Property Let IdCliente(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IdCliente = 5
    mvarIdCliente = vData
End Property

Public Property Get IdCliente() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.IdCliente
    IdCliente = mvarIdCliente
End Property

Public Function ExisteCPF(pADODB As cADODB, pCPF As String) As Boolean
Dim strSQL As String

ExisteCPF = False
strSQL = "select cpf from cliente where cpf = '" & pCPF & "' union select cpf from corretor where cpf = '" & pCPF & "'"
Dim oRSET As ADODB.Recordset
Set oRSET = pADODB.OpenRSet(strSQL)
If Not oRSET.EOF Then
    ExisteCPF = True
End If

End Function


Public Function AddCliente(pADODB As cADODB) As Double
Dim strSQL As String
Dim dblLastID As Double
Dim oRSET As ADODB.Recordset

'*** valida se cpf ja existe em cliente ou corretor
If ExisteCPF(pADODB, mvarCPF) Then
    MsgBox "CPF: " & Format(mvarCPF, "@@@.@@@.@@@-@@") & " ja existe em Cliente ou Corretor.", vbInformation
    AddCliente = -1
    Exit Function
End If



strSQL = "INSERT INTO dbo.Cliente (Nome, CPF, Endereco, Ativo, CidadeID) VALUES ("
    strSQL = strSQL & "'" & mvarNome & "',"
    strSQL = strSQL & "'" & mvarCPF & "',"
    strSQL = strSQL & "'" & mvarEndereco & "',"
    strSQL = strSQL & mvarAtivo & ","
    strSQL = strSQL & mvarCidadeID & ")"
    pADODB.ExecSQL strSQL
    Set oRSET = pADODB.OpenRSet("select @@IDENTITY")
    If Not oRSET.EOF Then
        dblLastID = oRSET(0).Value
        mvarIdCliente = oRSET(0).Value
        AddCliente = oRSET(0).Value
    End If
    '****************************
    strSQL = "INSERT INTO dbo.ClienteCorretor (IdCorretor, IdCliente) values("
    strSQL = strSQL & mvarCorretorID & "," & dblLastID & ")"
    pADODB.ExecSQL strSQL
    
End Function

Public Function DelCliente(pADODB As cADODB, pCPF As String) As Boolean
DelCliente = False
Dim strSQL As String

strSQL = "delete from ClienteCorretor where IdCliente in (select IdCliente from Cliente Where cpf = '" & pCPF & "')"
pADODB.ExecSQL strSQL

strSQL = "delete from cliente where cpf = '" & pCPF & "'"
pADODB.ExecSQL strSQL
 
DelCliente = True
End Function
