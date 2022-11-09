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
'# Version 1.58.38  published: 09.11.2022                                                      #
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
    
    If Len(Me.Default.ToCSS) > 0 Then
        strText.AddLine "        .pvtUi { " & Me.Default.ToCSS & " }"
    End If
    If Len(Me.Vals.ToCSS) > 0 Then
        strText.AddLine "        table .pvtVals { " & Me.Vals.ToCSS & " }"
    End If
    If Len(Me.Table.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable tbody tr td { " & Me.Table.ToCSS & " important! }"
    End If
    If Len(Me.Val.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable .pvtVal { " & Me.Val.ToCSS & " }"
    End If
    If Len(Me.Total.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable .pvtTotal { " & Me.Total.ToCSS & " }"
    End If
    If Len(Me.GrandTotal.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable .pvtGrandTotal { " & Me.GrandTotal.ToCSS & " }"
    End If
    If Len(Me.AxisContainer.ToCSS) > 0 Then
        strText.AddLine "        table.pvtUi .pvtAxisContainer { " & Me.AxisContainer.ToCSS & " }"
    End If
    If Len(Me.AxisLabel.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable .pvtAxisLabel { " & Me.AxisLabel.ToCSS & " }"
    End If
    If Len(Me.RowLabel.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable .pvtRowLabel  { " & Me.RowLabel.ToCSS & " }"
    End If
    If Len(Me.ColLabel.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable .pvtColLabel { " & Me.ColLabel.ToCSS & " }"
    End If
    If Len(Me.TotalLabel.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable .pvtTotalLabel { " & Me.TotalLabel.ToCSS & " }"
    End If
    If Len(Me.Attr.ToCSS) > 0 Then
        strText.AddLine "        table.pvtTable .pvtAttr { " & Me.Attr.ToCSS & " }"
    End If
    
    ToCSS = strText.ToString
    
End Function
