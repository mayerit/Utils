VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2085
   ClientLeft      =   8310
   ClientTop       =   5325
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4125
   Begin VB.CommandButton Command1 
      Caption         =   "Terminate Process"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Process ID"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Call TerminateProc(Text1)


End Sub
