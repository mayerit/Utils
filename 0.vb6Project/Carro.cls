VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Carro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)

Public CarroID As String
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent Evento1[(arg1, arg2, ... , argn)]
Public Event Evento1(a1 As Variant, a2 As String)
'local variable(s) to hold property value(s)
Private mvarCarroAno As Integer 'local copy

Private mvarCarroNome As String 'local copy
Public Property Let CarroAno(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CarroAno = 5
    mvarCarroAno = vData
End Property


Public Property Get CarroAno() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CarroAno
    CarroAno = mvarCarroAno
End Property






Public Function DelCarro(CarroID As String) As Boolean
End Function


Public Function AddCarro(CarroID As String, CarroNome As String) As Boolean
End Function


Public Property Let CarroNome(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CarroNome = 5
    mvarCarroNome = vData
End Property


Public Property Get CarroNome() As String
Attribute CarroNome.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CarroNome
    CarroNome = mvarCarroNome
End Property



