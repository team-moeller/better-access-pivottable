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
'# Copyright (c) 2021 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.08.16  published: 11.12.2021                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Fields As String


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
