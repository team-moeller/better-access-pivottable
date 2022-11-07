VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT__Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###############################################################################################
'# Copyright (c) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.56.05  published: 07.11.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Helper  As BAPT_Helper
Private m_Scripts As BAPT_Scripts


'### Instancing

Private Sub Class_Initialize()
    
    Set m_Helper = New BAPT_Helper
    Set m_Scripts = New BAPT_Scripts

End Sub

    
'### Initializer

Public Function PivotTable(Optional ByVal TheControl As WebBrowserControl) As BAPT_PivotTable
  
    Set PivotTable = New BAPT_PivotTable
    PivotTable.Init TheControl
    
End Function


'### Public Methods

Friend Function Generator(ByVal PivotTableOptions As BAPT_PivotTable) As BAPT_PivotTableGenerator

    Set Generator = New BAPT_PivotTableGenerator
    Generator.Init PivotTableOptions

End Function

Friend Function Helper() As BAPT_Helper

    Set Helper = m_Helper

End Function

Public Function Scripts() As BAPT_Scripts
    
    Set Scripts = m_Scripts

End Function


