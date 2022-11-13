VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Columns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.62.03  published: 13.11.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Fields As String
Private m_ShowTotals As Boolean
Private m_Order As ptOrder


'### Properties

Public Property Let Fields(ByVal newValue As String)
    newValue = "'" & newValue & "'"
    newValue = Replace(newValue, ",", "','")
    newValue = Replace(newValue, ",' ", ",'")
    m_Fields = newValue
End Property
Public Property Get Fields() As String
    Fields = m_Fields
End Property

Public Property Let ShowTotals(ByVal newValue As Boolean)
    m_ShowTotals = newValue
End Property
Public Property Get ShowTotals() As Boolean
    ShowTotals = m_ShowTotals
End Property

Public Property Let Order(ByVal newValue As ptOrder)
    m_Order = newValue
End Property
Public Property Get Order() As ptOrder
    Order = m_Order
End Property


'### Instanzing

Private Sub Class_Initialize()
    
    'Set default values
    Me.ShowTotals = True
    Me.Order = ptOrder.key_a_to_z
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


' ### Public Functions


' ### Generator Function

Friend Function GetCssToHideTotals() As String

    GetCssToHideTotals = "        tbody > tr > th.pvtTotalLabel, .colTotal, .pvtGrandTotal { display: none; }"

End Function
