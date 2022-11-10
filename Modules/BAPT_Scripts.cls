VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Scripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.59.03  published: 10.11.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_PivotJS       As BAPT_PivotJS
Private m_PivotCSS      As BAPT_PivotCSS
Private m_JQueryJS      As BAPT_JQueryJS
Private m_JQueryUI      As BAPT_JQueryUI


'### Properties

Public Property Get PivotJS() As BAPT_PivotJS
    Set PivotJS = m_PivotJS
End Property

Public Property Get PivotCSS() As BAPT_PivotCSS
    Set PivotCSS = m_PivotCSS
End Property

Public Property Get JQueryJS() As BAPT_JQueryJS
    Set JQueryJS = m_JQueryJS
End Property

Public Property Get JQueryUI() As BAPT_JQueryUI
    Set JQueryUI = m_JQueryUI
End Property


'### Instanzing

Private Sub Class_Initialize()

    Set m_PivotJS = New BAPT_PivotJS
    Set m_PivotCSS = New BAPT_PivotCSS
    Set m_JQueryJS = New BAPT_JQueryJS
    Set m_JQueryUI = New BAPT_JQueryUI
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods


