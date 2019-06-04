VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro Cliente"
   ClientHeight    =   3540
   ClientLeft      =   10290
   ClientTop       =   4155
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   2300
   End
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
      Left            =   2520
      TabIndex        =   8
      Top             =   2880
      Width           =   2300
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sai&r"
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
      Left            =   4920
      TabIndex        =   9
      Top             =   2880
      Width           =   2300
   End
   Begin VB.ComboBox cboCorretor 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox txtNome 
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
      Left            =   1320
      TabIndex        =   1
      Text            =   "LUIZ FERNANDO MARTINS MAYER"
      Top             =   480
      Width           =   4095
   End
   Begin VB.CheckBox chkInativo 
      Caption         =   "Ativo"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cboCidade 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2025
      Width           =   4095
   End
   Begin VB.ComboBox cboUF 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1650
      Width           =   1215
   End
   Begin VB.TextBox txtEndereco 
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
      Left            =   1320
      TabIndex        =   3
      Text            =   "RUA CAMPANULAS, 27"
      Top             =   1260
      Width           =   4095
   End
   Begin MSMask.MaskEdBox mskCPF 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   870
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
      Caption         =   "Corretor"
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
      TabIndex        =   15
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
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
      TabIndex        =   14
      Top             =   540
      Width           =   555
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   13
      Top             =   2124
      Width           =   675
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   12
      Top             =   1728
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Endereço"
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
      TabIndex        =   11
      Top             =   1332
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CPF"
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
      TabIndex        =   10
      Top             =   936
      Width           =   390
   End
End
Attribute VB_Name = "frmCadCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub PReencheComboCorretor()
Dim oCorretor As New cCorretor
Dim lll As New cADODB
Dim oRSET As ADODB.Recordset

Set oRSET = oCorretor.ToList(lll)
    cboCorretor.Clear
    Do While Not oRSET.EOF
        ''Debug.Print oRSET("ID") & oRSET("Nome") & oRSET("IDUF")
        cboCorretor.AddItem oRSET("Nome")
        cboCorretor.ItemData(cboCorretor.NewIndex) = (oRSET("IDCORRETOR"))
        oRSET.MoveNext
    Loop

End Sub



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

Private Sub cboCidade_Change()
    'cboCidade.ListIndex = -1
End Sub

Private Sub cboCidade_Click()
    'PreecheComboCidades
End Sub

Private Sub cboCidade_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub cboCorretor_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub cboUF_Change()
    cboCidade.ListIndex = -1
End Sub

Private Sub cboUF_Click()
    PreecheComboCidades
End Sub

Private Sub cboUF_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub cboUF_LostFocus()
If cboUF.ListIndex = -1 Then
    'MsgBox "Seleção inválida para UF"
End If
    'MsgBox cboUF.ItemData(cboUF.ListIndex)
End Sub




Private Sub Combo1_Change()

End Sub

Private Sub Combo1_GotFocus()
    
End Sub

Private Sub cmdLimpar_Click()
Dim sTemp As String

    chkInativo.Value = vbChecked
    txtNome.text = ""
    txtEndereco.text = ""
    cboUF.ListIndex = -1
    cboCidade.ListIndex = -1
    cboCorretor.ListIndex = -1
    sTemp = mskCPF.Mask
    mskCPF.Mask = ""
    mskCPF.text = ""
    mskCPF.Mask = sTemp
    
    PReencheComboCorretor
    
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

    
    'txtNome.SetFocus
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
Dim oCliente As New cCliente
Dim lll As New cADODB
Dim lId As Double

    oCliente.Nome = txtNome.text
    oCliente.CPF = mskCPF.ClipText
    oCliente.Endereco = txtEndereco.text
    oCliente.CidadeID = cboCidade.ItemData(cboCidade.ListIndex)
    oCliente.Ativo = IIf(chkInativo.Value = vbChecked, 1, 0)
    oCliente.CorretorID = cboCorretor.ItemData(cboCorretor.ListIndex)
    
    lId = oCliente.AddCliente(lll)
    If lId > 0 Then
        MsgBox "Cliente Gravado com Sucesso!" '& lId
        'MsgBox oCliente.IdCliente
    End If


End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
    cmdLimpar_Click
End Sub

Private Sub mskCPF_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub mskCPF_Validate(Cancel As Boolean)
    If Len(mskCPF.ClipText) <> 11 And Not (mskCPF.ClipText = vbNullString) Then
        MsgBox "CPF inválido.", vbInformation
        mskCPF.SetFocus
        GeneralModViceri.Sendkeys "{HOME}"
        GeneralModViceri.Sendkeys "+{END}"
        Cancel = True
    End If
End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub
