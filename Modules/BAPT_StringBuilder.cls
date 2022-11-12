VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###############################################################################################
'# Copyright (C) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.61.02  published: 12.11.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


Private m_String As String

Public Sub Add(ByVal TheString As String)
    m_String = m_String & TheString
End Sub

Public Sub AddLine(ByVal TheString As String)
    If m_String = vbNullString Then
        m_String = TheString
    ElseIf TheString = vbNullString Then
        m_String = m_String
    Else
        m_String = m_String & vbCrLf & TheString
    End If
End Sub

Public Sub CutTrailingComma()
    m_String = BAPT.Helper.CutTrailingComma(m_String)
End Sub

Public Function ToString() As String
    ToString = m_String
End Function

