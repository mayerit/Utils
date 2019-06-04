VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadCorretor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro Corretor"
   ClientHeight    =   2055
   ClientLeft      =   6105
   ClientTop       =   5100
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   5
      Top             =   1320
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
      TabIndex        =   4
      Top             =   1320
      Width           =   2300
   End
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
      TabIndex        =   3
      Top             =   1320
      Width           =   2300
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   1920
      TabIndex        =   0
      Text            =   "1212"
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox txtNome 
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
      Left            =   1920
      TabIndex        =   1
      Text            =   "Nome do Corretor XXXXXX"
      Top             =   420
      Width           =   4095
   End
   Begin MSMask.MaskEdBox mskCPF 
      DataField       =   "cliente.cpf;=;str"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   780
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1515
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
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   1395
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
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   1125
   End
End
Attribute VB_Name = "frmCadCorretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimpar_Click()
Dim sTemp As String

txtCodigo.text = ""
txtNome.text = ""
sTemp = mskCPF.Mask
    mskCPF.Mask = ""
    mskCPF.text = ""
    mskCPF.Mask = sTemp

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
Dim oCorretor As New cCorretor
Dim lll As New cADODB
Dim lId As Double

    oCorretor.Codigo = txtCodigo.text
    oCorretor.Nome = txtNome.text
    oCorretor.CPF = mskCPF.ClipText
    
    lId = oCorretor.AddCorretor(lll)
    If lId > 0 Then
        MsgBox "Correto Gravado com Sucesso!" '& lId
        'MsgBox oCorretor.IdCorretor
    End If

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

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    SendTabOnReturn (KeyAscii)
End Sub
