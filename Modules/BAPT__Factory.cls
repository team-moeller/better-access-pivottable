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
'# Copyright (c) 2021 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.07.05  published: 08.12.2021                                                      #
'###############################################################################################

Option Compare Database
Option Explicit

Private m_Helper  As BAPT_Helper
Private m_Scripts As BAPT_Scripts

Public Function PivotTable(Optional ByVal TheControl As WebBrowserControl) As BAPT_PivotTable
  
    Set PivotTable = New BAPT_PivotTable
    PivotTable.Init TheControl
    
End Function

Friend Function Helper() As BAPT_Helper

    Set Helper = m_Helper

End Function

Public Function Scripts() As BAPT_Scripts
    
    Set Scripts = m_Scripts

End Function

Private Sub Class_Initialize()
    
    Set m_Helper = New BAPT_Helper
    Set m_Scripts = New BAPT_Scripts

End Sub



