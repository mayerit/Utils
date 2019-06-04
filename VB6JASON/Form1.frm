VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Some JSON to RS to FlexGrid"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2010
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1875
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   3307
      _Version        =   393216
      BackColorBkg    =   -2147483636
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With New ADODB.Stream
        .Type = adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile "json.txt"
        Set MSHFlexGrid1.DataSource = Transform.JsonToRecordset(.ReadText(adReadAll))
        .Close
    End With
    With MSHFlexGrid1
        .ColWidth(0) = 300
        .ColWidth(1) = 600
        .ColWidth(2) = 1200
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
    End With
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        MSHFlexGrid1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub
