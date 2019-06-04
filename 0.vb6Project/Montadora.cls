VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Montadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MontadoraCOLL"
Public MontadoraID As String
Attribute MontadoraID.VB_VarUserMemId = 0
Attribute MontadoraID.VB_VarDescription = "description da montadorID XXXX"

Private mvarNome As String 'local copy
Private mvarMontadoraCOLL As MontadoraCOLL


Public Property Get MontadoraCOLL() As MontadoraCOLL
    If mvarMontadoraCOLL Is Nothing Then
        Set mvarMontadoraCOLL = New MontadoraCOLL
    End If


    Set MontadoraCOLL = mvarMontadoraCOLL
End Property


Public Property Set MontadoraCOLL(vData As MontadoraCOLL)
    Set mvarMontadoraCOLL = vData
End Property
Private Sub Class_Terminate()
    Set mvarMontadoraCOLL = Nothing
End Sub



Public Property Let Nome(ByVal vData As String)
Attribute Nome.VB_Description = "description da montadora nome"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nome = 5
    mvarNome = vData
End Property


Public Property Get Nome() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nome
    Nome = mvarNome
End Property



