VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021, 2023 Thomas Möller, supported by K.D.Gundermann                         #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 2.03.02  published: 28.03.2023                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Background As BAPT_Background
Private m_Font As BAPT_Font


'### Properties

Public Property Get Background() As BAPT_Background
    Set Background = m_Background
End Property

Public Property Get Font() As BAPT_Font
    Set Font = m_Font
End Property


'### Instanzing

Private Sub Class_Initialize()
    
    'Create sub classes
    Set m_Background = New BAPT_Background
    Set m_Font = New BAPT_Font
    
End Sub


'### Public Functions


'### Generator Function

Friend Function ToCSS() As String

    'Declarations
    Dim strText As New BAPT_StringBuilder
    
    strText.Add Me.Background.ToCSS
    strText.Add Me.Font.ToCSS
    
    ToCSS = strText.ToString
    
End Function
