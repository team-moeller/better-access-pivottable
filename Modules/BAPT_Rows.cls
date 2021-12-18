VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Rows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.10.11  published: 18.12.2021                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Fields As String
Private m_ShowTotals As Boolean


'### Properties

Public Property Let Fields(newValue As String)
    newValue = "'" & newValue & "'"
    newValue = Replace(newValue, ",", "','")
    newValue = Replace(newValue, ",' ", ",'")
    m_Fields = newValue
End Property
Public Property Get Fields() As String
    Fields = m_Fields
End Property

Public Property Let ShowTotals(newValue As Boolean)
    m_ShowTotals = newValue
End Property
Public Property Get ShowTotals() As Boolean
    ShowTotals = m_ShowTotals
End Property


'### Instanzing

Private Sub Class_Initialize()
    
    'Set default values
    Me.ShowTotals = True
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


' ### Public Functions


' ### Generator Function

Friend Function GetCssToHideTotals() As String

    GetCssToHideTotals = "    <style>thead > tr > th.pvtTotalLabel, .rowTotal, .pvtGrandTotal { display: none; }</style>"

End Function


