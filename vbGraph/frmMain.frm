VERSION 5.00
Object = "*\AvbGraph.vbp"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin vbGraph.Graph Graph1 
      Height          =   945
      Left            =   120
      TabIndex        =   14
      Top             =   90
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   1667
      State           =   "frmMain.frx":0000
   End
   Begin VB.Timer tmrAddPoint 
      Left            =   180
      Top             =   0
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Height          =   3075
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   7725
      TabIndex        =   0
      Top             =   1365
      Width           =   7785
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Point"
         Height          =   465
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1365
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   285
         LargeChange     =   50
         Left            =   120
         Max             =   500
         Min             =   1
         TabIndex        =   11
         Top             =   2610
         Value           =   500
         Width           =   3735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Grid"
         Height          =   465
         Left            =   120
         TabIndex        =   10
         Top             =   630
         Width           =   1365
      End
      Begin VB.CommandButton cmdLoadGrid 
         Caption         =   "Load Grid"
         Height          =   465
         Left            =   1530
         TabIndex        =   9
         Top             =   630
         Width           =   1365
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Point"
         Height          =   465
         Left            =   1530
         TabIndex        =   8
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "Music Style"
         Height          =   465
         Index           =   0
         Left            =   4500
         TabIndex        =   7
         Top             =   150
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "Taskman Style"
         Height          =   465
         Index           =   1
         Left            =   4500
         TabIndex        =   6
         Top             =   690
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "Graph Style"
         Height          =   465
         Index           =   2
         Left            =   4500
         TabIndex        =   5
         Top             =   1230
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "Custom Style"
         Height          =   465
         Index           =   3
         Left            =   5940
         TabIndex        =   4
         Top             =   150
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "Bar Style"
         Height          =   465
         Index           =   4
         Left            =   5940
         TabIndex        =   3
         Top             =   690
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "Line Style"
         Height          =   465
         Index           =   5
         Left            =   5940
         TabIndex        =   2
         Top             =   1230
         Width           =   1365
      End
      Begin VB.CommandButton cmdStyle 
         Caption         =   "Progress Bar"
         Height          =   465
         Index           =   6
         Left            =   4500
         TabIndex        =   1
         Top             =   1740
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Speed"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   2340
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FILENAME As String = "C:\Woof.dat"
Private Const FILENAME_PIC As String = "C:\Woof.bmp"

Private Const STYLE_MUSIC As Long = 0
Private Const STYLE_TASKMAN As Long = 1
Private Const STYLE_GRAPH As Long = 2
Private Const STYLE_CUSTOM As Long = 3
Private Const STYLE_BAR As Long = 4
Private Const STYLE_LINE As Long = 5
Private Const STYLE_PROGRESS As Long = 6

Private mlngStyle   As Long

Private Sub cmdAdd_Click()
    AddPoint
End Sub

Private Sub cmdLoadGrid_Click()
    tmrAddPoint.Enabled = False
    HScroll1.Value = HScroll1.Max
    Graph1.Loadsettings FILENAME
    tmrAddPoint.Enabled = True
End Sub

Private Sub cmdRemove_Click()
    If Graph1.Points.Count > 0 Then
        Graph1.Points.Remove Graph1.Points.Count
    End If
End Sub

Private Sub cmdSave_Click()
    Graph1.Savesettings FILENAME
End Sub

Private Sub cmdStyle_Click(Index As Integer)
    SetStyle Index
End Sub

Private Sub Form_Load()
    Randomize
    SetStyle STYLE_TASKMAN
End Sub

Private Sub Form_Resize()
Dim lngGap  As Long
On Error Resume Next
    lngGap = 150
    Graph1.Move lngGap, lngGap, Me.ScaleWidth - (lngGap * 2), Me.ScaleHeight - picButtons.Height - (lngGap * 2)
End Sub

Private Sub HScroll1_Change()
    tmrAddPoint.Interval = IIf(HScroll1.Value = 500, 0, HScroll1.Value)
End Sub

Private Sub tmrAddPoint_Timer()
    AddPoint
End Sub

Private Sub AddPoint()
Dim lngIndex    As Long
Dim lngValue    As Long
Dim blnRedraw   As Boolean

    With Graph1
        blnRedraw = .Redraw
        .Redraw = False
        Select Case mlngStyle
            Case STYLE_MUSIC
                .Points.Clear
                For lngIndex = 1 To IIf(Graph1.FixedPoints = 0, Graph1.Points.Count, Graph1.FixedPoints)
                    lngValue = (Rnd * 80) + 50
                    .Points.Add lngValue
                Next lngIndex
            Case Else
                If mlngStyle = STYLE_PROGRESS Then
                    .Points.Add .MaxValue
                    .Redraw = True
                    If .Points.Count = .FixedPoints Then
                        .Points.Clear
                    End If
                Else
                    lngValue = (Rnd * (.MaxValue - .MinValue)) + .MinValue
                    .Points.Add lngValue
                End If
        End Select
        .Redraw = blnRedraw
    End With

End Sub

Private Sub SetStyle(ByVal plngIndex As Long)
Dim lngX    As Long
    With Graph1
        mlngStyle = plngIndex
        .Redraw = False
        .Points.Clear
        Select Case plngIndex
            Case STYLE_MUSIC
                .FixedPoints = 60
                .ShowAxis = False
                .ShowGrid = False
                .ShowPoints = False
                .ShowBars = True
                .ShowLines = False
                .FadeIn = False
                .MaxValue = 200
                .MinValue = 0
                .YGridInc = 20
                .xGridInc = 1
                .BarWidth = 1
                .BackColor = &H404040
                .BarColor = &HFF8080
                .LineColor = &HFF00&
                .GridColor = &H808080
                HScroll1.Value = 150
            Case STYLE_TASKMAN
                .FixedPoints = 60
                .BarWidth = 0.8
                .ShowAxis = False
                .ShowGrid = True
                .ShowPoints = False
                .ShowLines = True
                .ShowBars = False
                .FadeIn = False
                .MaxValue = 100
                .MinValue = 0
                .YGridInc = 20
                .xGridInc = 1
                .BackColor = RGB(0, 0, 0)
                .LineColor = &HFF00&
                .GridColor = &H808080
                HScroll1.Value = 100
            Case STYLE_GRAPH
                .FadeIn = False
                .FixedPoints = 0
                .MaxValue = 1.2
                .BarWidth = 0.8
                .MinValue = -1.2
                .YGridInc = 0.5
                .xGridInc = 90
                .AxisColor = 0
                .BackColor = RGB(255, 255, 255)
                .LineColor = RGB(200, 50, 255)
                .GridColor = RGB(200, 200, 200)
                .PointColor = RGB(255, 100, 100)
                .ShowAxis = True
                .ShowGrid = True
                .ShowLines = True
                .ShowPoints = True
                .ShowBars = False
                For lngX = 1 To 720
                    .Points.Add Sin(CDbl(CDbl(3.14) * CDbl(2) * CDbl((lngX / 360))))
                Next lngX
                HScroll1.Value = HScroll1.Max
            Case STYLE_CUSTOM
                .BarWidth = 0.8
                .FixedPoints = 50
                .MaxValue = 70
                .MinValue = -30
                .xGridInc = 1
                .YGridInc = 10
                .ShowAxis = True
                .ShowGrid = True
                .ShowPoints = True
                .ShowLines = True
                .ShowBars = True
                .FadeIn = True
                .ShowBars = True
                .BackColor = RGB(121, 145, 200)
                .GridColor = RGB(110, 135, 190)
                .LineColor = RGB(255, 255, 255)
                .PointColor = RGB(255, 0, 0)
                .BarColor = &HC0FFFF
                HScroll1.Value = 100
            Case STYLE_BAR
                .BarWidth = 0.8
                .FixedPoints = 0
                .MaxValue = 100
                .MinValue = 0
                .xGridInc = 1
                .YGridInc = 10
                .ShowAxis = False
                .ShowGrid = False
                .ShowPoints = False
                .ShowLines = False
                .ShowBars = True
                .FadeIn = False
                .BackColor = RGB(121, 145, 200)
                .GridColor = RGB(110, 135, 190)
                .LineColor = RGB(255, 255, 255)
                .PointColor = RGB(255, 255, 255)
                .BarColor = &HECD2BF
                Do
                    AddPoint
                Loop Until .Points.Count = 10
                HScroll1.Value = HScroll1.Max
            Case STYLE_LINE
                .BarWidth = 0.8
                .FixedPoints = 0
                .MaxValue = 100
                .MinValue = 0
                .xGridInc = 1
                .YGridInc = 10
                .ShowAxis = False
                .ShowGrid = True
                .ShowPoints = True
                .ShowLines = True
                .ShowBars = False
                .FadeIn = False
                .BackColor = RGB(255, 255, 255)
                .GridColor = RGB(200, 200, 200)
                .LineColor = &HA98354
                .PointColor = RGB(255, 0, 0)
                .BarColor = &HECD2BF
                Do
                    AddPoint
                Loop Until .Points.Count = 60
                HScroll1.Value = HScroll1.Max
            Case STYLE_PROGRESS
                .BarWidth = 1
                .FixedPoints = 80
                .MaxValue = 1
                .MinValue = 0
                .xGridInc = 1
                .YGridInc = 1
                .ShowAxis = False
                .ShowGrid = False
                .ShowPoints = False
                .ShowLines = False
                .ShowBars = True
                .FadeIn = False
                .BackColor = RGB(255, 255, 255)
                .GridColor = RGB(200, 200, 200)
                .LineColor = &HA98354
                .PointColor = RGB(255, 0, 0)
                .BarColor = &HECD2BF
                HScroll1.Value = 100
        End Select
        .Redraw = True
    End With
End Sub

