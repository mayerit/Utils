VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viceri Teste"
   ClientHeight    =   7620
   ClientLeft      =   1200
   ClientTop       =   2835
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   14055
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   11520
      TabIndex        =   17
      Top             =   120
      Width           =   2300
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   11520
      TabIndex        =   16
      Top             =   1920
      Width           =   2300
   End
   Begin VB.CommandButton cmdCadCliente 
      Caption         =   "Cadastrar Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   11520
      TabIndex        =   15
      Top             =   1320
      Width           =   2300
   End
   Begin VB.CommandButton cmdCadCorretor 
      Caption         =   "Cadastrar Corretor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   11520
      TabIndex        =   14
      Top             =   720
      Width           =   2300
   End
   Begin VB.ComboBox cboCidade 
      DataField       =   "cidade.id;=;int"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   2895
   End
   Begin VB.ComboBox cboUF 
      DataField       =   "uf.id;=;int"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox chkInativoCliente 
      Caption         =   "Ativo"
      DataField       =   "cliente.ativo;=;int"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtNomeCliente 
      DataField       =   "cliente.nome;like;str"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Text            =   "Nome do Cliente XXXXXX"
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox txtNomeCorretor 
      DataField       =   "corretor.nome;like;str"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "Nome do Corretor XXXXXX"
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox txtCodigoCorretor 
      DataField       =   "corretor.codigo;like;str"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Text            =   "1212"
      Top             =   240
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5055
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   8916
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskCPFCliente 
      DataField       =   "cliente.cpf;=;str"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "999.999.999-99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
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
      Left            =   6720
      TabIndex        =   13
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "UF"
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
      Left            =   6720
      TabIndex        =   12
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CPF Cliente"
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
      Left            =   480
      TabIndex        =   11
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nome Clliente"
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
      Left            =   480
      TabIndex        =   10
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome Corretor"
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
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código Corretor"
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
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub PreecheComboCidades()
Dim oCidade As New cCidade
Dim lll As New cADODB
Dim oRSET As ADODB.Recordset

    If cboUF.ListIndex = -1 Then Exit Sub
    Set oRSET = oCidade.ToList(lll, cboUF.ItemData(cboUF.ListIndex))
    cboCidade.Clear
    Do While Not oRSET.EOF
        ''Debug.Print oRSET("ID") & oRSET("Nome") & oRSET("IDUF")
        cboCidade.AddItem oRSET("Nome")
        cboCidade.ItemData(cboCidade.NewIndex) = (oRSET("ID"))
        oRSET.MoveNext
    Loop

End Sub

Private Sub cboCidade_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub cboUF_Change()
    cboCidade.ListIndex = -1
End Sub

Private Sub cboUF_Click()
    'cboCidade.ListIndex = -1
    PreecheComboCidades
End Sub

Private Sub cboUF_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub chkInativoCliente_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub cmdCadCliente_Click()
    frmCadCliente.Show vbModal
    cmdLimpar_Click
    cmdPesquisar_Click
    
    

End Sub

Private Sub cmdCadCorretor_Click()
    frmCadCorretor.Show vbModal
End Sub

Private Sub cmdLimpar_Click()
    Dim sTemp As String

    txtCodigoCorretor.text = ""
    txtNomeCorretor.text = ""
    txtNomeCliente.text = ""
    sTemp = mskCPFCliente.Mask
    mskCPFCliente.Mask = ""
    mskCPFCliente.text = ""
    mskCPFCliente.Mask = sTemp
    chkInativoCliente.Value = vbChecked
    cboUF.ListIndex = -1
    cboCidade.ListIndex = -1
    
    Dim oUF As New cUF
Dim lll As New cADODB
Dim oRSET As ADODB.Recordset


    Set oRSET = oUF.ToList(lll)
    cboUF.Clear
    If oRSET Is Nothing Then Exit Sub
    cboUF.AddItem ""
    Do While Not oRSET.EOF
        'Debug.Print oRSET("ID") & oRSET("Nome")
        cboUF.AddItem oRSET("Nome")
        cboUF.ItemData(cboUF.NewIndex) = (oRSET("ID"))
        oRSET.MoveNext
    Loop
    
On Error Resume Next
txtCodigoCorretor.SetFocus
End Sub

Private Sub cmdPesquisar_Click()

Dim oCli As New cCliente
Dim lll As New cADODB
Dim oRSET As New ADODB.Recordset
'**************************************************************************************
'*** Construção do Filtro
'Dim x As String
Dim k
Dim Filtro As String
Dim FiltroSQL As New Collection
Dim Where As String
Dim x As Object
For Each x In Me.Controls
    If (TypeOf x Is TextBox) Or (TypeOf x Is MaskEdBox) Or (TypeOf x Is CheckBox) Or (TypeOf x Is ComboBox) Then
        'Debug.Print x.Name & x.DataField
        If x.DataField <> "" Then
            k = Split(x.DataField, ";")
            If (TypeOf x Is TextBox) Then
                If x.text <> "" Then
                    Filtro = ""
                    Filtro = k(0) & " " & k(1) & " " & _
                        IIf(k(2) = "str", "'", "") & _
                        x.text & IIf(k(1) = "like", "%", "") & _
                        IIf(k(2) = "str", "'", "")
                    FiltroSQL.Add Filtro
                End If
            End If
            
            If (TypeOf x Is MaskEdBox) Then
                If x.ClipText <> "" Then
                    Filtro = ""
                    Filtro = k(0) & " " & k(1) & " " & _
                        IIf(k(2) = "str", "'", "") & _
                        x.ClipText & IIf(k(1) = "like", "%", "") & _
                        IIf(k(2) = "str", "'", "")
                    FiltroSQL.Add Filtro
                End If
            End If
            
            If (TypeOf x Is CheckBox) Then
                Filtro = ""
                Filtro = k(0) & " " & k(1) & " " & IIf(x.Value = vbChecked, "1", "0")
                    FiltroSQL.Add Filtro
            End If
            
            If (TypeOf x Is ComboBox) Then
                If x.text <> "" Then
                    Filtro = ""
                    Filtro = k(0) & " " & k(1) & " " & x.ItemData(x.ListIndex)
                    FiltroSQL.Add Filtro
                End If
            End If
            
        End If
    End If
Next

Where = ""
Dim kk As Variant
For Each kk In FiltroSQL
    Where = Where & kk & " and "
Next
Where = Mid(Where, 1, Len(Where) - 5)
''Debug.Print Where
'**************************************************************************************
Set oRSET = oCli.ToList(lll, Where)
FillGrid oRSET

'ResizeGrid MSFlexGrid1, Me
MSFlexGrid1.ColWidth(0) = 3000
MSFlexGrid1.ColAlignment(0) = flexAlignLeftCenter
MSFlexGrid1.ColWidth(1) = 1800
MSFlexGrid1.ColAlignment(1) = flexAlignLeftCenter
MSFlexGrid1.ColWidth(2) = 600
MSFlexGrid1.ColAlignment(2) = flexAlignLeftCenter
MSFlexGrid1.ColWidth(3) = 3000
MSFlexGrid1.ColAlignment(3) = flexAlignLeftCenter
MSFlexGrid1.ColWidth(4) = 1500
MSFlexGrid1.ColAlignment(4) = flexAlignLeftCenter
MSFlexGrid1.ColWidth(5) = 400
MSFlexGrid1.ColAlignment(5) = flexAlignLeftCenter
MSFlexGrid1.ColWidth(6) = 3000
MSFlexGrid1.ColAlignment(6) = flexAlignLeftCenter
MSFlexGrid1.ColWidth(7) = 400
MSFlexGrid1.ColAlignment(7) = flexAlignCenterCenter


End Sub




Public Sub FillGrid(ByRef pRSet As ADODB.Recordset)

MSFlexGrid1.Rows = pRSet.RecordCount + 1
MSFlexGrid1.Cols = pRSet.Fields.Count '+ 1 ' - 1
MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0

Dim x As Object
Dim ColValue As Integer

ColValue = 0
For Each x In pRSet.Fields
    'Debug.Print pRSet.Fields(ColValue).Name & " " & MSFlexGrid1.Col
    MSFlexGrid1.Col = ColValue
    MSFlexGrid1.text = x.Name
     'MSFlexGrid1.Col '+ 1
    ColValue = ColValue + 1
Next

MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1 ' - 1
'MSFlexGrid1.Col = MSFlexGrid1.Col + 1

MSFlexGrid1.Cols = MSFlexGrid1.Cols - 1

If MSFlexGrid1.Rows = 1 Then Exit Sub
MSFlexGrid1.Row = 1
MSFlexGrid1.Col = 0
MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1



MSFlexGrid1.Clip = pRSet.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
MSFlexGrid1.Row = 1

Dim i As Integer
For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Col = 7
    MSFlexGrid1.Row = i
    MSFlexGrid1.CellForeColor = vbRed
Next
End Sub



Private Sub Form_Load()
    cmdLimpar_Click
End Sub

Private Sub MSFlexGrid1_Click()
'If MSFlexGrid1.Col < 1 Or MSFlexGrid1.Row < 1 Then Exit Sub
 '   MsgBox "Row:" & MSFlexGrid1.Row & " Col:" & MSFlexGrid1.Col
 ''Debug.Print MSFlexGrid1.Row & MSFlexGrid1.Col
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim strMensagem As String
Dim lstrCPF As String
Dim lll As New cADODB
Dim k As New cCliente

If y < MSFlexGrid1.RowHeight(0) Then Exit Sub



    If MSFlexGrid1.MouseCol = 7 Then
        lstrCPF = MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, 1)
        strMensagem = "Excluir Cliente: " & MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, 0) & _
            " CPF: " & MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, 1)
        If MsgBox(strMensagem, vbYesNo + vbQuestion) = vbYes Then
            If k.DelCliente(lll, Replace(Replace(lstrCPF, ".", ""), "-", "")) Then
                MsgBox "Cliente Excluído com Sucesso.", vbInformation
                cmdPesquisar_Click
            End If
        End If
    End If
    'MsgBox "Row:" & MSFlexGrid1.MouseRow & " Col:" & MSFlexGrid1.MouseCol & " X:" & x & " Y:" & y

    
    
End Sub


Private Function calculacpf(CPF As String) As Boolean
    calculacpf = True
End Function

Public Function ValidaCGC(CGC As String) As Boolean
    ValidaCGC = True
End Function

Private Sub mskCPFCliente_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub mskCPFCliente_Validate(Cancel As Boolean)
    If Len(mskCPFCliente.ClipText) <> 11 And Not (mskCPFCliente.ClipText = vbNullString) Then
        MsgBox "CPF inválido.", vbInformation
        mskCPFCliente.SetFocus
        GeneralModViceri.Sendkeys "{HOME}"
        GeneralModViceri.Sendkeys "+{END}"
        Cancel = True
    End If
End Sub

Private Sub txtCodigoCorretor_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub
Private Sub txtNomeCliente_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub txtNomeCorretor_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub
