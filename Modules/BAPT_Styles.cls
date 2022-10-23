VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Styles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'### Members

Private m_Default As BAPT_Style

Private m_Table As BAPT_Style

Private m_Val As BAPT_Style
Private m_Total As BAPT_Style
Private m_GrandTotal As BAPT_Style

Private m_AxisLabel As BAPT_Style
Private m_RowLabel As BAPT_Style
Private m_ColLabel As BAPT_Style
Private m_TotalLabel As BAPT_Style

Private m_Attr As BAPT_Style


'### Properties

Public Property Get Default() As BAPT_Style
    Set Default = m_Default
End Property

Public Property Get Table() As BAPT_Style
    Set Table = m_Table
End Property

Public Property Get Val() As BAPT_Style
    Set Val = m_Val
End Property

Public Property Get Total() As BAPT_Style
    Set Total = m_Total
End Property

Public Property Get GrandTotal() As BAPT_Style
    Set GrandTotal = m_GrandTotal
End Property

Public Property Get AxisLabel() As BAPT_Style
    Set AxisLabel = m_AxisLabel
End Property

Public Property Get RowLabel() As BAPT_Style
    Set RowLabel = m_RowLabel
End Property

Public Property Get ColLabel() As BAPT_Style
    Set ColLabel = m_ColLabel
End Property

Public Property Get TotalLabel() As BAPT_Style
    Set TotalLabel = m_TotalLabel
End Property

Public Property Get Attr() As BAPT_Style
    Set Attr = m_Attr
End Property

'### Instanzing

Private Sub Class_Initialize()
    
    'Create sub classes
    Set m_Default = New BAPT_Style
    
    Set m_Table = New BAPT_Style
    Set m_Val = New BAPT_Style
    Set m_Total = New BAPT_Style
    Set m_GrandTotal = New BAPT_Style
    
    Set m_AxisLabel = New BAPT_Style
    Set m_RowLabel = New BAPT_Style
    Set m_ColLabel = New BAPT_Style
    Set m_TotalLabel = New BAPT_Style
    
    Set m_Attr = New BAPT_Style
    
    'Set default values
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

'### Generator Function

Friend Function ToCSS() As String
    Dim strText As New BAPT_StringBuilder
    
    strText.AddLine ".pvtUi { " & Default.ToCSS & " }"
    
    strText.AddLine "table.pvtTable { " & Table.ToCSS & " }"
    strText.AddLine "table.pvtTable .pvtVal { " & Val.ToCSS & " }"
    strText.AddLine "table.pvtTable .pvtTotal { " & Total.ToCSS & " }"
    strText.AddLine "table.pvtTable .pvtGrandTotal { " & GrandTotal.ToCSS & " }"
    
    strText.AddLine "table.pvtTable .pvtAxisLabel { " & AxisLabel.ToCSS & " }"
    strText.AddLine "table.pvtTable .pvtRowLabel  { " & RowLabel.ToCSS & " }"
    strText.AddLine "table.pvtTable .pvtColLabel { " & ColLabel.ToCSS & " }"
    strText.AddLine "table.pvtTable .pvtTotalLabel { " & TotalLabel.ToCSS & " }"
    
    strText.AddLine "table.pvtTable .pvtAttr { " & Attr.ToCSS & " }"
    
    ToCSS = strText.ToString
End Function
