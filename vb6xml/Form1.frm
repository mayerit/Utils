VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SaveValuesIndented"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtStreet 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Street"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Zip"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "State"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "City"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_AppPath As String

' Add a new node to the indicated parent node.
Private Sub CreateNode(ByVal indent As Integer, ByVal parent As IXMLDOMNode, ByVal node_name As String, ByVal node_value As String)
Dim new_node As IXMLDOMNode

    ' Indent.
    parent.appendChild parent.ownerDocument.createTextNode(Space$(indent))

    ' Create the new node.
    Set new_node = parent.ownerDocument.createElement(node_name)

    ' Set the node's text value.
    new_node.Text = node_value

    ' Add the node to the parent.
    parent.appendChild new_node

    ' Add a newline.
    parent.appendChild parent.ownerDocument.createTextNode(vbCrLf)
End Sub
' Return the node's value.
Private Function GetNodeValue(ByVal start_at_node As IXMLDOMNode, ByVal node_name As String, Optional ByVal default_value As String = "") As String
Dim value_node As IXMLDOMNode

    Set value_node = start_at_node.selectSingleNode(".//" & node_name)
    If value_node Is Nothing Then
        GetNodeValue = default_value
    Else
        GetNodeValue = value_node.Text
    End If
End Function

' Load saved values from XML.
Private Sub LoadValues()
Dim xml_document As DOMDocument
Dim values_node As IXMLDOMNode

    ' Load the document.
    Set xml_document = New DOMDocument
    xml_document.Load m_AppPath & "Values.xml"

    ' If the file doesn't exist, then
    ' xml_document.documentElement is Nothing.
    If xml_document.documentElement Is Nothing Then
        ' The file doesn't exist. Do nothing.
        Exit Sub
    End If

    ' Find the Values section.
    Set values_node = xml_document.selectSingleNode("Values")

    ' Read the saved values.
    txtFirstName.Text = GetNodeValue(values_node, "FirstName", "???")
    txtLastName.Text = GetNodeValue(values_node, "LastName", "???")
    txtStreet.Text = GetNodeValue(values_node, "Street", "???")
    txtCity.Text = GetNodeValue(values_node, "City", "???")
    txtState.Text = GetNodeValue(values_node, "State", "???")
    txtZip.Text = GetNodeValue(values_node, "Zip", "???")
End Sub
' Save the current values.
Private Sub SaveValues()
Dim xml_document As DOMDocument
Dim values_node As IXMLDOMNode

    ' Create the XML document.
    Set xml_document = New DOMDocument

    ' Create the Values section node.
    Set values_node = xml_document.createElement("Values")

    ' Add a newline.
    values_node.appendChild xml_document.createTextNode(vbCrLf)

    ' Add the Values section node to the document.
    xml_document.appendChild values_node

    ' Create nodes for the values inside the
    ' Values section node.
    CreateNode 4, values_node, "FirstName", txtFirstName.Text
    CreateNode 4, values_node, "LastName", txtLastName.Text
    CreateNode 4, values_node, "Street", txtStreet.Text
    CreateNode 4, values_node, "City", txtCity.Text
    CreateNode 4, values_node, "State", txtState.Text
    CreateNode 4, values_node, "Zip", txtZip.Text

    ' Save the XML document.
    xml_document.save m_AppPath & "Values.xml"
End Sub
Private Sub Form_Load()
    ' Get the application's startup path.
    m_AppPath = App.Path
    If Right$(m_AppPath, 1) <> "\" Then m_AppPath = m_AppPath & "\"

    ' Load the values.
    LoadValues
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Save the current values.
    SaveValues
End Sub

