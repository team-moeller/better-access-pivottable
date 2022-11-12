VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Styles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021, 2022 Thomas Möller, supported by K.D.Gundermann                         #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.61.02  published: 12.11.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Default As BAPT_Style
Private m_Table As BAPT_Style
Private m_Vals As BAPT_Style    ' better name for this ??
Private m_Val As BAPT_Style
Private m_Total As BAPT_Style
Private m_GrandTotal As BAPT_Style
Private m_AxisContainer As BAPT_Style
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

Public Property Get Vals() As BAPT_Style
    Set Vals = m_Vals
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

Public Property Get AxisContainer() As BAPT_Style
    Set AxisContainer = m_AxisContainer
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
    Set m_Vals = New BAPT_Style
    Set m_Val = New BAPT_Style
    Set m_Total = New BAPT_Style
    Set m_GrandTotal = New BAPT_Style
    Set m_AxisContainer = New BAPT_Style
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


'### Public Functions


'### Generator Function

Friend Function ToCSS() As String

    'Declarations
    Dim strText As New BAPT_StringBuilder
    
    strText.AddLine getCSS(".pvtUi", Default)
    strText.AddLine getCSS("table.pvtTable tbody tr td", Table, "important!")
    strText.AddLine getCSS("table .pvtVals", Vals)
    strText.AddLine getCSS("table.pvtTable .pvtVal", Val)
    strText.AddLine getCSS("table.pvtTable .pvtTotal", Total)
    strText.AddLine getCSS("table.pvtTable .pvtGrandTotal", GrandTotal)
    strText.AddLine getCSS("table.pvtUi .pvtAxisContainer", AxisContainer)
    strText.AddLine getCSS("table.pvtTable .pvtAxisLabel", AxisLabel)
    strText.AddLine getCSS("table.pvtTable .pvtRowLabel", RowLabel)
    strText.AddLine getCSS("table.pvtTable .pvtColLabel", ColLabel)
    strText.AddLine getCSS("table.pvtTable .pvtTotalLabel", TotalLabel)
    strText.AddLine getCSS(".pvtUi .pvtAttr", Attr)
    
    ToCSS = strText.ToString
    
End Function

Private Function getCSS(ByVal Scope As String, ByVal Style As BAPT_Style, Optional ByVal Annotation As String) As String

    'Declarations
    Dim strCSS As String

    strCSS = Style.ToCSS
    If strCSS <> vbNullString Then
        getCSS = Space(8) & Scope & " { " & strCSS
        If Len(Annotation) > 0 Then
            getCSS = getCSS & " " + Annotation & " }"
        Else
            getCSS = getCSS & " }"
        End If
    End If
    
End Function
