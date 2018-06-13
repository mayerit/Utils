VERSION 5.00
Begin VB.UserControl Graph 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   3510
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
   End
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020

Private WithEvents mobjPoints      As Points
Attribute mobjPoints.VB_VarHelpID = -1

Private mudtControlProps    As gtypControlProps
Private mudtGraphProps      As gtypGraphProps

Private mblnDesignMode  As Boolean

Public Enum eBorderStyle
   egrNone = 0
   egrFixedSingle = 1
End Enum

Public Enum eAppearance
   egrFlat = 0
   egr3D = 1
End Enum

Private Type mtypPOINT
    X   As Long
    Y   As Long
End Type

Private Type mtypRECT
    Left    As Long
    Right   As Long
    Top     As Long
    Bottom  As Long
End Type

Private Sub UserControl_Initialize()
    picDraw.FillStyle = vbFSSolid
    Set mobjPoints = New Points
End Sub

Private Sub UserControl_InitProperties()
    InitProperties
    mblnDesignMode = Not UserControl.Ambient.UserMode
    DrawControl
    PropertyChanged PB_STATE
End Sub

Private Sub UserControl_Paint()
    DrawGraph
End Sub

Private Sub UserControl_Terminate()
    Set mobjPoints = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    State = PropBag.ReadProperty(PB_STATE, State)
    mblnDesignMode = Not UserControl.Ambient.UserMode
    DrawControl
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PB_STATE, State
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    With UserControl
        picDraw.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
    Refresh
End Sub

Private Sub mobjPoints_Changed()
Static blnWorking As Boolean
    If Not blnWorking Then
        blnWorking = True
        RemovePoints
        If Not mblnDesignMode Then
            DrawGraph
        End If
        blnWorking = False
    End If
End Sub

Private Property Let GraphState(ByRef Value() As Byte)
Dim udtData     As gtypGraphData
    udtData.Data = Value
    LSet mudtGraphProps = udtData
End Property

Private Property Get GraphState() As Byte()
Dim udtData     As gtypGraphData
    LSet udtData = mudtGraphProps
    GraphState = udtData.Data
End Property

Friend Property Let ControlState(ByRef Value() As Byte)
Dim udtData     As gtypControlData
    udtData.Data = Value
    LSet mudtControlProps = udtData
End Property

Friend Property Get ControlState() As Byte()
Dim udtData     As gtypControlData
    LSet udtData = mudtControlProps
    ControlState = udtData.Data
End Property

Private Property Let State(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        ControlState = .ReadProperty(PB_CONTROL)
        GraphState = .ReadProperty(PB_GRAPH)
    End With
    Set objPB = Nothing
End Property

Private Property Get State() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_CONTROL, ControlState
        .WriteProperty PB_GRAPH, GraphState
        State = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Property Let SuperState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        State = .ReadProperty(PB_STATE, State)
        mobjPoints.SuperState = .ReadProperty(PB_POINTS, mobjPoints.SuperState)
    End With
    Set objPB = Nothing
End Property

Friend Property Get SuperState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_STATE, State
        .WriteProperty PB_POINTS, mobjPoints.SuperState
        SuperState = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Property Let FileState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        GraphState = .ReadProperty(PB_GRAPH, GraphState)
        mobjPoints.SuperState = .ReadProperty(PB_POINTS, mobjPoints.SuperState)
    End With
    Set objPB = Nothing
End Property

Friend Property Get FileState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_GRAPH, GraphState
        .WriteProperty PB_POINTS, mobjPoints.SuperState
        FileState = .Contents
    End With
    Set objPB = Nothing
End Property


Private Sub InitProperties()
    With mudtGraphProps
        .BackColor = RGB(255, 255, 255)
        .LineColor = RGB(84, 131, 169)
        .BarColor = &HDEA68D
        .PointColor = &HFFFFC0
        .AxisColor = RGB(0, 0, 0)
        .GridColor = RGB(223, 223, 223)
        .FixedPoints = 20
        .XGridInc = 1
        .YGridInc = 10
        .MaxValue = 100
        .MinValue = 0
        .FadeIn = False
        .ShowGrid = True
        .ShowAxis = False
        .ShowLines = True
        .ShowPoints = True
        .ShowBars = True
        .BarWidth = 0.8
    End With
    With mudtControlProps
        .Redraw = True
        .BorderStyle = eBorderStyle.egrFixedSingle
        .Appearance = eAppearance.egr3D
    End With
End Sub

Public Property Get Points() As Points
    Set Points = mobjPoints
End Property

Public Property Let Redraw(ByVal Value As Boolean)
    mudtControlProps.Redraw = Value
    If Value Then
        Refresh
    End If
End Property

Public Property Get Redraw() As Boolean
    Redraw = mudtControlProps.Redraw
End Property

Public Property Let Appearance(ByVal Value As eAppearance)
    mudtControlProps.Appearance = Value
    UserControl.Appearance = Value
    DrawControl
    PropertyChanged PB_STATE
End Property

Public Property Get Appearance() As eAppearance
    Appearance = mudtControlProps.Appearance
End Property

Public Property Let BorderStyle(ByVal Value As eBorderStyle)
    mudtControlProps.BorderStyle = Value
    UserControl.BorderStyle = Value
    DrawControl
    PropertyChanged PB_STATE
End Property

Public Property Get BorderStyle() As eBorderStyle
    BorderStyle = mudtControlProps.BorderStyle
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.BackColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mudtGraphProps.BackColor
End Property

Public Property Let LineColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.LineColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = mudtGraphProps.LineColor
End Property

Public Property Let BarColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.BarColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BarColor() As OLE_COLOR
    BarColor = mudtGraphProps.BarColor
End Property

Public Property Let PointColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.PointColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get PointColor() As OLE_COLOR
    PointColor = mudtGraphProps.PointColor
End Property

Public Property Let AxisColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.AxisColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get AxisColor() As OLE_COLOR
    AxisColor = mudtGraphProps.AxisColor
End Property

Public Property Let GridColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.GridColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get GridColor() As OLE_COLOR
    GridColor = mudtGraphProps.GridColor
End Property

Public Property Let FixedPoints(ByVal Value As Long)
    mudtGraphProps.FixedPoints = Value
    RemovePoints
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get FixedPoints() As Long
    FixedPoints = mudtGraphProps.FixedPoints
End Property

Public Property Let XGridInc(ByVal Value As Long)
    mudtGraphProps.XGridInc = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get XGridInc() As Long
    XGridInc = mudtGraphProps.XGridInc
End Property

Public Property Let YGridInc(ByVal Value As Double)
    mudtGraphProps.YGridInc = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get YGridInc() As Double
    YGridInc = mudtGraphProps.YGridInc
End Property

Public Property Let MaxValue(ByVal Value As Double)
    mudtGraphProps.MaxValue = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get MaxValue() As Double
    MaxValue = mudtGraphProps.MaxValue
End Property

Public Property Let MinValue(ByVal Value As Double)
    mudtGraphProps.MinValue = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get MinValue() As Double
    MinValue = mudtGraphProps.MinValue
End Property

Public Property Let ShowGrid(ByVal Value As Boolean)
    mudtGraphProps.ShowGrid = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowGrid() As Boolean
    ShowGrid = mudtGraphProps.ShowGrid
End Property

Public Property Let ShowAxis(ByVal Value As Boolean)
    mudtGraphProps.ShowAxis = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowAxis() As Boolean
    ShowAxis = mudtGraphProps.ShowAxis
End Property

Public Property Let ShowLines(ByVal Value As Boolean)
    mudtGraphProps.ShowLines = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowLines() As Boolean
    ShowLines = mudtGraphProps.ShowLines
End Property

Public Property Let ShowBars(ByVal Value As Boolean)
    mudtGraphProps.ShowBars = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowBars() As Boolean
    ShowBars = mudtGraphProps.ShowBars
End Property

Public Property Let ShowPoints(ByVal Value As Boolean)
    mudtGraphProps.ShowPoints = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowPoints() As Boolean
    ShowPoints = mudtGraphProps.ShowPoints
End Property

Public Property Let FadeIn(ByVal Value As Boolean)
    mudtGraphProps.FadeIn = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get FadeIn() As Boolean
    FadeIn = mudtGraphProps.FadeIn
End Property

Public Property Let BarWidth(ByVal Value As Single)
    mudtGraphProps.BarWidth = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BarWidth() As Single
    BarWidth = mudtGraphProps.BarWidth
End Property

Private Sub AddDefaultPoints()
    mobjPoints.Clear
    AddDefaultPoint 80
    AddDefaultPoint 10
    AddDefaultPoint 70
    AddDefaultPoint 25
    AddDefaultPoint 50
    AddDefaultPoint 45
    AddDefaultPoint 15
    AddDefaultPoint 85
    AddDefaultPoint 5
    AddDefaultPoint 75
    AddDefaultPoint 65


End Sub

Private Sub AddDefaultPoint(ByVal plngPercent As Long)
    mobjPoints.Add (plngPercent / 100) * (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) + mudtGraphProps.MinValue
End Sub

Private Sub RemovePoints()
    Do While mudtGraphProps.FixedPoints > 0 And mobjPoints.Count > mudtGraphProps.FixedPoints
        mobjPoints.Remove 1
    Loop
End Sub

Public Sub Refresh()
    DrawGraph
End Sub

Public Sub DrawControl()
    With UserControl
        .Appearance = mudtControlProps.Appearance
        .BorderStyle = mudtControlProps.BorderStyle
    End With
End Sub

Private Sub DrawGraph()
Dim lngX        As Long
Dim lngY        As Long
Dim lngCount    As Long
Dim lngStepX    As Long
Dim lngStepY    As Long
Dim lngWidth    As Long
Dim lngHeight   As Long
Dim lngIndex    As Long
Dim udtPoints() As mtypPOINT
Dim lngYAxis    As Long
Dim lngBarWidth As Long
Dim lngFixedCount   As Long
Dim udtBar      As mtypRECT
Dim udtGrid     As mtypRECT
    If UserControl.Height > 0 And UserControl.Width > 0 Then
    If mudtControlProps.Redraw Or mblnDesignMode Then
        If mblnDesignMode Then
            AddDefaultPoints
        End If
        With picDraw
            .Cls
            .BackColor = mudtGraphProps.BackColor

            lngWidth = .ScaleWidth - 15
            lngHeight = .ScaleHeight - 15

            'draw grid
            lngCount = mobjPoints.Count
            If mudtGraphProps.FixedPoints = 0 Or lngCount < mudtGraphProps.FixedPoints Then
                If mudtGraphProps.FixedPoints > 0 Then
                    lngFixedCount = mudtGraphProps.FixedPoints
                Else
                    lngFixedCount = lngCount
                End If
            Else
                lngFixedCount = mudtGraphProps.FixedPoints
            End If
            If lngFixedCount > 0 Then
                If mudtGraphProps.ShowBars Then
                    If lngCount > lngFixedCount Or Not mudtGraphProps.FadeIn Then
                        lngBarWidth = CLng((lngWidth / lngFixedCount) * mudtGraphProps.BarWidth)
                    ElseIf lngCount > 0 Then
                        lngBarWidth = CLng((lngWidth / lngCount) * mudtGraphProps.BarWidth)
                    End If
                End If
            End If

            udtPoints = GetPoints(0, lngWidth, 0, lngHeight, lngBarWidth)

            With udtGrid
                .Left = 0
                .Top = 0
                .Right = lngWidth
                .Bottom = lngHeight
            End With
            DrawGrid udtGrid, mudtGraphProps.GridColor, lngBarWidth

            If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) > 0 Then
                lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
            End If

            'drawlines and bars
            If lngCount > 0 And (mudtGraphProps.ShowLines Or mudtGraphProps.ShowBars) Then
                For lngIndex = 1 To UBound(udtPoints)
                    If mudtGraphProps.ShowBars Then
                        udtBar.Left = udtPoints(lngIndex).X - (lngBarWidth / 2)
                        udtBar.Right = udtPoints(lngIndex).X + (lngBarWidth / 2)

                        udtBar.Top = udtPoints(lngIndex).Y
                        udtBar.Bottom = lngYAxis
                        DrawBar udtBar, mudtGraphProps.BarColor
                    End If
                    If mudtGraphProps.ShowLines And lngIndex > 1 Then
                        DrawLine udtPoints(lngIndex - 1), udtPoints(lngIndex), mudtGraphProps.LineColor
                  End If
                Next lngIndex
            End If

            'draw axis
            If mudtGraphProps.ShowAxis Then
                picDraw.Line (0, 0)-(0, lngHeight), mudtGraphProps.AxisColor
                If mudtGraphProps.MaxValue <= 0 Then
                    picDraw.Line (0, 0)-(lngWidth, 0), mudtGraphProps.AxisColor
                ElseIf mudtGraphProps.MinValue < 0 Then
                    picDraw.Line (0, (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue)-(lngWidth, (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue), mudtGraphProps.AxisColor
                Else
                    picDraw.Line (0, lngHeight)-(lngWidth, lngHeight), mudtGraphProps.AxisColor
                End If
            End If

            'draw points
            If lngCount > 0 And mudtGraphProps.ShowPoints Then
                For lngIndex = 1 To UBound(udtPoints)
                    If lngIndex Mod mudtGraphProps.XGridInc = 0 Or lngIndex = 1 Then
                        DrawPoint udtPoints(lngIndex), mudtGraphProps.PointColor, "Woof"
                    End If
                Next lngIndex
            End If

            'copy picture to usercontrol
            BitBlt UserControl.hDC, 0, 0, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, SRCCOPY
        End With
    End If
    End If
End Sub

Private Function GetPoints(ByVal plngLeft As Long, ByVal plngRight As Long, ByVal plngTop As Long, ByVal plngBottom As Long, ByVal plngBarWidth As Long) As mtypPOINT()
Dim udtPoints() As mtypPOINT
Dim lngCount    As Long
Dim lngIndex    As Long
Dim objPoint    As Point
Dim lngX        As Long
Dim lngPtCount  As Long
Dim lngFixedCount   As Long
Dim lngYAxis    As Long
    lngCount = mobjPoints.Count

    If lngCount > 0 Then
        If mudtGraphProps.FixedPoints = 0 Or lngCount < mudtGraphProps.FixedPoints Then
            lngPtCount = lngCount
            If mudtGraphProps.FixedPoints > 0 Then
                lngFixedCount = mudtGraphProps.FixedPoints
            Else
                lngFixedCount = lngCount
            End If
        Else
            lngPtCount = mudtGraphProps.FixedPoints
            lngFixedCount = mudtGraphProps.FixedPoints
        End If
        ReDim udtPoints(lngPtCount) As mtypPOINT

        For Each objPoint In mobjPoints
            lngIndex = lngIndex + 1
            If mudtGraphProps.FixedPoints > 0 And lngIndex > mudtGraphProps.FixedPoints Then
                Set objPoint = Nothing
                Exit For
            End If

            If lngIndex = 1 Then
                If lngFixedCount = 1 Then
                    lngX = plngLeft + (((plngRight - plngLeft)) / 2)
                Else
                    lngX = plngLeft + (plngBarWidth / 2)
                End If
            ElseIf lngIndex = lngFixedCount Then
                lngX = plngRight - (plngBarWidth / 2)
            Else
                lngX = (lngIndex - 1) * (((plngRight - plngLeft) - plngBarWidth) / (lngFixedCount - 1)) + (plngBarWidth / 2)
            End If

            udtPoints(lngIndex).X = lngX
            If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) <> 0 Then
                lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
                udtPoints(lngIndex).Y = lngYAxis - objPoint.Value * ((plngBottom - plngTop) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue))
            End If
        Next objPoint
    End If
    GetPoints = udtPoints
End Function

Private Sub DrawLine(ByRef pudtPt1 As mtypPOINT, ByRef pudtPt2 As mtypPOINT, ByVal plngColor As String)
    picDraw.Line (pudtPt1.X, pudtPt1.Y)-(pudtPt2.X, pudtPt2.Y), plngColor
End Sub

Private Sub DrawPoint(ByRef pudtPt As mtypPOINT, ByVal plngColor As Long, ByVal pstrCaption As String)
    picDraw.FillColor = plngColor
    picDraw.Circle (pudtPt.X, pudtPt.Y), 40, 0 'plngColor
End Sub

Private Sub DrawBar(ByRef pudtRect As mtypRECT, ByVal plngColor As Long)
    picDraw.FillColor = plngColor
    With pudtRect
        picDraw.Line (.Left, .Top)-(.Right, .Bottom), 0, B
    End With
End Sub

Private Sub DrawGrid(ByRef pudtRect As mtypRECT, ByVal plngColor As Long, ByVal plngBarWidth As Long)
Dim lngCount    As Long
Dim lngIndex    As Long
Dim lngX        As Long
Dim lngY        As Long
Dim lngFixedCount   As Long
Dim lngYAxis    As Long
Dim lngStepY    As Long
Dim lngHeight   As Long
    lngCount = mobjPoints.Count
    If lngCount > 0 And mudtGraphProps.ShowGrid Then

        lngHeight = picDraw.ScaleHeight - 15

        If mudtGraphProps.FixedPoints = 0 Or lngCount < mudtGraphProps.FixedPoints Then
            If mudtGraphProps.FixedPoints > 0 Then
                lngFixedCount = mudtGraphProps.FixedPoints
            Else
                lngFixedCount = lngCount
            End If
        Else
            lngFixedCount = mudtGraphProps.FixedPoints
        End If
        For lngIndex = 1 To lngFixedCount
            If lngIndex = 1 Then
                If lngFixedCount = 1 Then
                    lngX = pudtRect.Left + (((pudtRect.Right - pudtRect.Left)) / 2)
                Else
                    lngX = pudtRect.Left + (plngBarWidth / 2)
                End If
            ElseIf lngIndex = lngFixedCount Then
                lngX = pudtRect.Right - (plngBarWidth / 2)
            Else
                lngX = (lngIndex - 1) * (((pudtRect.Right - pudtRect.Left) - plngBarWidth) / (lngFixedCount - 1)) + (plngBarWidth / 2)
            End If
            If lngIndex Mod mudtGraphProps.XGridInc = 0 Then
                picDraw.Line (lngX, 0)-(lngX, lngHeight), plngColor
            End If
        Next lngIndex

        'draw horizontal lines
        If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) > 0 Then
            lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
            lngStepY = (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.YGridInc
            For lngY = lngYAxis To 0 Step -lngStepY
                picDraw.Line (0, lngHeight - lngY)-(picDraw.ScaleWidth - 15, lngHeight - lngY), mudtGraphProps.GridColor
            Next lngY

            For lngY = lngYAxis To lngHeight Step lngStepY
                picDraw.Line (0, lngHeight - lngY)-(picDraw.ScaleWidth - 15, lngHeight - lngY), mudtGraphProps.GridColor
            Next lngY
        End If
    End If
End Sub

Public Sub SaveSettings(ByVal Filename As String)
    If Len(Filename) > 0 Then
        If Dir(Filename) <> vbNullString Then
            Kill Filename
        End If
    End If
    SaveFile Filename, FileState
End Sub

Public Sub LoadSettings(ByVal Filename As String)
    FileState = GetFile(Filename)
    Refresh
End Sub

