VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarIDCidade As Double 'local copy
Private mvarNome As String 'local copy
'local variable(s) to hold property value(s)
Private mvarUFID As Double 'local copy

Public Function ToList(pADODB As cADODB, pIDUF As Double) As ADODB.Recordset
Dim strSQL As String
Dim oRSET As ADODB.Recordset

strSQL = "SELECT ID, NOME, IDUF FROM CIDADE WHERE IDUF = " & pIDUF & " ORDER BY NOME"
Set ToList = pADODB.OpenRSet(strSQL)
End Function


Public Property Let UFID(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UFID = 5
    mvarUFID = vData
End Property


Public Property Get UFID() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.UFID
    UFID = mvarUFID
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



Public Property Let IDCidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDCidade = 5
    mvarIDCidade = vData
End Property


Public Property Get IDCidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: 'Debug.Print X.IDCidade
    IDCidade = mvarIDCidade
End Property



