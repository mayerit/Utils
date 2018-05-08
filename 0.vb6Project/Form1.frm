VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   2325
   ClientTop       =   2475
   ClientWidth     =   16545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   16545
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   4320
      TabIndex        =   36
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   35
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   34
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   33
      Text            =   "VERDANA"
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   32
      Text            =   "VERDANA"
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdInsertOracle 
      Caption         =   "Insert Oracle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   29
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   15735
      Begin VB.CommandButton Command2 
         Caption         =   "Insert 1000 x "
         Height          =   345
         Left            =   7800
         TabIndex        =   28
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CommandButton cmdInsertSP 
         Caption         =   "Insert SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   27
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtGuid 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   25
         Top             =   3480
         Width           =   5115
      End
      Begin VB.CommandButton cmdGuid 
         Caption         =   "Guid"
         Height          =   465
         Left            =   14280
         TabIndex        =   24
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdRoda1000 
         Caption         =   "Insert 1000 x "
         Height          =   345
         Left            =   7800
         TabIndex        =   23
         Top             =   270
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDateValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   20
         Text            =   "09/10/2015"
         Top             =   2970
         Width           =   3435
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6510
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDecValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   16
         Text            =   "123456789.78"
         Top             =   2520
         Width           =   3435
      End
      Begin VB.TextBox txtIntValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   14
         Text            =   "999"
         Top             =   2070
         Width           =   3435
      End
      Begin VB.TextBox txtDesignation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   12
         Text            =   "designation teste"
         Top             =   1620
         Width           =   3435
      End
      Begin VB.TextBox txtLastName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   10
         Text            =   "last name teste"
         Top             =   1170
         Width           =   3435
      End
      Begin VB.TextBox txtFirstName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   8
         Text            =   "fisrt name teste"
         Top             =   720
         Width           =   3435
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   6
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Guid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   3510
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Date Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dec Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   2550
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Int Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   2130
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   1230
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdSQLServerADO 
      Caption         =   "&SQL Server ADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdSQLServer 
      Caption         =   "&SQL Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salvar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "LUIZ FERNANDO MARTINS MAYER"
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9840
      TabIndex        =   31
      Top             =   960
      Width           =   630
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9840
      TabIndex        =   30
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   645
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5640
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://docs.oracle.com/cd/E51173_01/win.122/e18594/using.htm#OLEDB168
'http://www.vb6.us/tutorials/oracle-and-visual-basic-using-ado
'https://www.connectionstrings.com/oracle-provider-for-ole-db-oraoledb/
'http://www.c-sharpcorner.com/UploadFile/nipuntomar/connection-strings-for-oracle/
'http://www.macoratti.net/vb6_msfg.htm

Private Sub cmdQuery_Click()

End Sub

Private Sub cmdDelete_Click()
Dim con As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String

con.ConnectionString = "Provider=SQLOLEDB;" _
         & "Server=ROOT-18826FD16E;" _
         & "Database=POC;" _
         & "Integrated Security=SSPI;" _
         & "DataTypeCompatibility=80;"

con.Open
strSQL = "SET DATEFORMAT DMY"
con.Execute strSQL
'******************************************************
strSQL = ""
strSQL = "delete from employee where id = " & txtID.Text
con.Execute strSQL

con.Close
Set con = Nothing

End Sub

Private Sub cmdGuid_Click()
    MsgBox GetGUID
End Sub

Private Sub cmdInsert_Click()

On Error GoTo Error_cmdInsert_Click

Dim con As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String, intID As Integer

'con.ConnectionString = "Provider=SQLOLEDB;" _
         & "Server=(localdb)\MSSQLLocalDB;" _
         & "Database=POC;" _
         & "Integrated Security=SSPI;" _
         & "DataTypeCompatibility=80;"


con.ConnectionString = "Provider=SQLNCLI11;Server=(localdb)\MSSQLLocalDB;Database=POC;Trusted_Connection=yes;timeout=30;"

con.Open
con.IsolationLevel = adXactSerializable
strSQL = "SET DATEFORMAT DMY"
con.Execute strSQL
'******************************************************
strSQL = ""
'strSQL = "insert into employee (firstname, lastname, designation, intvalue, decvalue, datevalue) values ('" & _
    txtFirstName.Text & "','" & txtLastName.Text & "','" & txtDesignation.Text & "'," & txtIntValue.Text & "," & txtDecValue.Text & ",'" & txtDateValue.Text & "')"
strSQL = "insert into employee (firstname, lastname, designation, intvalue, decvalue, datevalue, id) values ('" & _
    txtFirstName.Text & "','" & txtLastName.Text & "','" & txtDesignation.Text & "'," & txtIntValue.Text & "," & txtDecValue.Text & ",GETDATE(), '" & GetGUID & "')"
'con.BeginTrans
con.Execute strSQL
'con.CommitTrans

'Dim rst As New ADODB.Recordset
Set rst.ActiveConnection = con
rst.Source = "select SCOPE_IDENTITY()"
rst.Open

If Not rst.EOF Then
    'MsgBox rst.Fields(0).Value
    txtID.Text = rst.Fields(0).value
    intID = rst.Fields(0).value
End If
rst.Close
'**********************************************
Set rst.ActiveConnection = con
rst.Source = "select guid from employee where id = " & intID
rst.Open

If Not rst.EOF Then
    'MsgBox rst.Fields(0).Value
    txtGuid.Text = rst.Fields(0).value
End If
rst.Close

con.Close
Set con = Nothing
Exit Sub

Error_cmdInsert_Click:
'con.RollbackTrans

End Sub

Private Sub cmdInsertOracle_Click()
On Error GoTo Error_cmdInsert_Click

Dim con As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String, intID As Integer

'Provider=OraOLEDB.Oracle;dbq=localhost:1521/XE;Database=myDataBase;User Id=myUsername;Password=myPassword;

'con.ConnectionString = "Provider=OraOLEDB.Oracle;dbq=192.168.68.31:1521/XE;Database=myDataBase;User Id=myUsername;Password=myPassword;"
'con.ConnectionString = "Provider=OraOLEDB.Oracle;dbq=192.168.68.31:1521;User Id=scott;Password=tiger;"
con.Provider = "OraOLEDB.Oracle"
con.Properties("Data Source") = "ORCL_31"
con.Properties("User Id") = "scott"
con.Properties("Password") = "tiger"
con.Open

'con.Open
'con.IsolationLevel = adXactSerializable
'strSQL = "SET DATEFORMAT DMY"
'con.Execute strSQL
'******************************************************
strSQL = ""
strSQL = "SELECT SYSTIMESTAMP FROM DUAL"
Set rst.ActiveConnection = con
rst.Source = strSQL
rst.Open

If Not rst.EOF Then
    MsgBox rst.Fields(0).value
    'Exit Sub
End If

rst.Close

strSQL = "SELECT * FROM SCOTT.DEPT"
Set rst.ActiveConnection = con
rst.Source = strSQL
rst.Open

Do While Not rst.EOF
    MsgBox rst.Fields(0).value & "-" & rst.Fields(1).value
    rst.MoveNext
    
Loop
Exit Sub
'strSQL = "insert into employee (firstname, lastname, designation, intvalue, decvalue, datevalue) values ('" & _
    txtFirstName.Text & "','" & txtLastName.Text & "','" & txtDesignation.Text & "'," & txtIntValue.Text & "," & txtDecValue.Text & ",'" & txtDateValue.Text & "')"
strSQL = "insert into employee (firstname, lastname, designation, intvalue, decvalue, datevalue, guid) values ('" & _
    txtFirstName.Text & "','" & txtLastName.Text & "','" & txtDesignation.Text & "'," & txtIntValue.Text & "," & txtDecValue.Text & ",GETDATE(), '" & GetGUID & "')"
con.BeginTrans
con.Execute strSQL
con.CommitTrans

'Dim rst As New ADODB.Recordset
Set rst.ActiveConnection = con
rst.Source = "select SCOPE_IDENTITY()"
rst.Open

If Not rst.EOF Then
    'MsgBox rst.Fields(0).Value
    txtID.Text = rst.Fields(0).value
    intID = rst.Fields(0).value
End If
rst.Close
'**********************************************
Set rst.ActiveConnection = con
rst.Source = "select guid from employee where id = " & intID
rst.Open

If Not rst.EOF Then
    'MsgBox rst.Fields(0).Value
    txtGuid.Text = rst.Fields(0).value
End If
rst.Close

con.Close
Set con = Nothing
Exit Sub

Error_cmdInsert_Click:
con.RollbackTrans

End Sub

Private Sub cmdInsertSP_Click()
'http://www.w3schools.com/asp/met_comm_createparameter.asp

Dim con As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String, intID As Integer, strGUID As String

con.ConnectionString = "Provider=SQLOLEDB;" _
         & "Server=vs3550\SQLEXPRESS;" _
         & "Database=POCVB6;" _
         & "Integrated Security=SSPI;" _
         & "DataTypeCompatibility=80;"

con.Open
con.IsolationLevel = adXactSerializable
strSQL = "SET DATEFORMAT DMY"
con.Execute strSQL
DoEvents
'***********************************************************************
Set cmd = New ADODB.Command
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "InsertEmployee"

cmd.Parameters.Append cmd.CreateParameter("FirtsName", adVarChar, adParamInput, 50, txtFirstName.Text)
cmd.Parameters.Append cmd.CreateParameter("LastName", adVarChar, adParamInput, 50, txtLastName.Text)
cmd.Parameters.Append cmd.CreateParameter("Designation", adVarChar, adParamInput, 50, txtDesignation.Text)
cmd.Parameters.Append cmd.CreateParameter("IntValue", adInteger, adParamInput, 3, txtIntValue.Text)

'cmd.Parameters.Append cmd.CreateParameter("DecValue", adDecimal, adParamInput, 18, txtDecValue.Text)
'cmd("DecValue").Precision = 18
'cmd("DecValue").NumericScale = 2
cmd.Parameters.Append cmd.CreateParameter("DecValue", adVarChar, adParamInput, 18, txtDecValue.Text)

cmd.Parameters.Append cmd.CreateParameter("DateValue", adVarChar, adParamInput, 10, txtDateValue.Text)
strGUID = GetGUID
txtGuid.Text = strGUID
cmd.Parameters.Append cmd.CreateParameter("Guid", adVarChar, adParamInput, 37, GetGUID)
cmd.Parameters.Append cmd.CreateParameter("result", adInteger, adParamOutput)

cmd.Execute
res = cmd("result")
txtID.Text = res

Set cmd = Nothing
'*****************************
con.Close
Set con = Nothing
 


End Sub

Private Sub cmdRoda1000_Click()
Dim i As Integer

i = 1
Do While (i <= 1000)
    cmdInsert_Click
    i = i + 1
Loop
End Sub

Private Sub cmdSQLServer_Click()
'-- Here we want to open the database
 Dim sConnectionString As String
 Dim strSQLStmt As String

 '-- Build the connection string
 sConnectionString = "PROVIDER = MSDASQL;driver={SQL Server};database=POC ;server=ROOT-18826FD16E;uid=sa;pwd=su79000;"


 strSQLStmt = "SELECT * from employee "

'DB WORK
Dim db As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim result As String

db.ConnectionString = sConnectionString
db.Open 'open connection

With cmd
  .ActiveConnection = db
  .CommandText = strSQLStmt
  .CommandType = adCmdText
End With

With rs
  .CursorType = adOpenStatic
  .CursorLocation = adUseClient
  .LockType = adLockOptimistic
  .Open cmd
End With

If rs.EOF = False Then
    rs.MoveFirst
    Let result = rs.Fields(0)
End If
'close conns
rs.Close
db.Close
Set db = Nothing
Set cmd = Nothing
Set rs = Nothing


'set local box

' TextBox1.Text = strSQLStmt
TextBox1.Text = result

End Sub

Private Sub cmdSQLServerADO_Click()
'http://www.sqlshack.com/creating-using-crud-stored-procedures/
'https://www.youtube.com/watch?v=W_IhNL9lAGI
'http://www.sqlshack.com/creating-using-crud-stored-procedures/
'Operação com data
'http://forum.imasters.com.br/topic/224454-manipulando-data-no-sql-server-conteudo-alterado/
'http://blog.sqlauthority.com/2007/09/28/sql-server-introduction-and-example-for-dateformat-command/

Dim con As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim sXMLResult As String

con.ConnectionString = "Provider=SQLOLEDB;" _
         & "Server=ROOT-18826FD16E;" _
         & "Database=POC;" _
         & "Integrated Security=SSPI;" _
         & "DataTypeCompatibility=80;"

con.Open

' Get the xml data as a recordset.
Set rst.ActiveConnection = con
rst.Source = "SELECT * from employee"
rst.Open

' Display the data in the recordset.
Do While (Not rst.EOF)
   sXMLResult = rst.Fields(0).value
   'Debug.Print (sXMLResult)
   rst.MoveNext
Loop

con.Close
Set con = Nothing
End Sub

Private Sub cmdUpdate_Click()
Dim con As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String

con.ConnectionString = "Provider=SQLOLEDB;" _
         & "Server=ROOT-18826FD16E;" _
         & "Database=POC;" _
         & "Integrated Security=SSPI;" _
         & "DataTypeCompatibility=80;"

con.Open
strSQL = "SET DATEFORMAT DMY"
con.Execute strSQL
'******************************************************
strSQL = ""
strSQL = "update employee set firstname = '" & txtFirstName.Text & "', lastname = '" & txtLastName.Text & "', designation = '" & txtDesignation.Text & _
    "', intvalue = " & txtIntValue.Text & ", decvalue = " & txtDecValue.Text & ", datevalue = GETDATE() where id = " & txtID.Text
    
con.Execute strSQL

con.Close
Set con = Nothing

End Sub

Private Sub Command1_Click()
Dim con   As New ADODB.Connection
Dim Rst1  As New ADODB.Recordset
Dim Rst2  As New ADODB.Recordset
Dim Rst3  As New ADODB.Recordset
Dim cmd   As New ADODB.Command
Dim Prm1  As New ADODB.Parameter
Dim Prm2  As New ADODB.Parameter

con.Provider = "OraOLEDB.Oracle"

con.ConnectionString = "Data Source=(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.11)(PORT=1521)))(CONNECT_DATA=(SID=ORCL)(SERVER=DEDICATED)));" & _
                       "User ID=scott;Password=tiger;"
con.Open
cmd.ActiveConnection = con

' Although Employees.GetEmpRecords() takes four parameters, only
' two need to be bound because Ref cursor parameters are automatically
' bound by the provider.

Set Prm1 = cmd.CreateParameter("Prm1", adSmallInt, adParamInput, , 20)
cmd.Parameters.Append Prm1
Set Prm2 = cmd.CreateParameter("Prm2", adSmallInt, adParamOutput)
cmd.Parameters.Append Prm2

' Enable PLSQLRSet property
cmd.Properties("PLSQLRSet") = True

' Stored Procedures returning resultsets must be called using the
' ODBC escape sequence for calling stored procedures.
cmd.CommandText = "{CALL Employees.GetEmpRecords(?, ?)}"

' Get the first recordset
Set Rst1 = cmd.Execute

'Set DataGrid1.DataSource = Rst1
'DataGrid1.Refresh
Debug.Print "----------- R1 ----------- "
Do While Not Rst1.EOF
    'MsgBox Rst1.Fields("EMPNO").Value & "-" & Rst1.Fields("ENAME").Value & "-" & Rst1.Fields("DEPTNO").Value
    Debug.Print Rst1.Fields("EMPNO").value & "-" & Rst1.Fields("ENAME").value & "-" & Rst1.Fields("DEPTNO").value
    Rst1.MoveNext
Loop

' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False

' Get the second recordset
Debug.Print "----------- R2 ----------- "
Set Rst2 = Rst1.NextRecordset
Do While Not Rst2.EOF
    'MsgBox Rst2.Fields("EMPNO").Value
    Debug.Print Rst2.Fields("EMPNO").value
    Rst2.MoveNext
Loop

' Just as in a stored procedure, the REF CURSOR return value must
' not be bound in a stored function.
Prm1.value = 7839
Prm2.value = 0

' Enable PLSQLRSet property
cmd.Properties("PLSQLRSet") = True

' Stored Functions returning resultsets must be called using the
' ODBC escape sequence for calling stored functions.
cmd.CommandText = "{CALL Employees.GetDept(?, ?)}"

' Get the rowset
Set Rst3 = cmd.Execute
Debug.Print "----------- R3 ----------- "
Do While Not Rst3.EOF
    Debug.Print Rst3.Fields("DEPTNO").value
    Rst3.MoveNext
Loop

' Disable PLSQLRSet
cmd.Properties("PLSQLRSet") = False

' Clean up
Rst1.Close
Rst2.Close
Rst3.Close

End Sub


Private Sub Command2_Click()
Dim i As Integer

i = 1
Do While (i <= 1000)
    cmdInsertSP_Click
    i = i + 1
Loop
End Sub

Private Sub Command5_Click()
Dim x As String
x = RandomString(1000)
MsgBox x
End Sub
