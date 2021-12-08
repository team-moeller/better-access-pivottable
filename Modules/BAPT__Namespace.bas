Attribute VB_Name = "BAPT__Namespace"
'###############################################################################################
'# Copyright (c) 2021 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.07.05  published: 08.12.2021                                                      #
'###############################################################################################

Option Compare Database
Option Explicit

Private m_BetterAccessPivotTable As BAPT__Factory

Public Property Get BetterAccessPivotTable() As BAPT__Factory

    If m_BetterAccessPivotTable Is Nothing Then
        Set m_BetterAccessPivotTable = New BAPT__Factory
    End If
    
    Set BetterAccessPivotTable = m_BetterAccessPivotTable

End Property

Public Property Get BAPT() As BAPT__Factory

    Set BAPT = BetterAccessPivotTable

End Property


