VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Background"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021, 2022 Thomas Möller, supported by K.D.Gundermann                         #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.63.04  published: 24.11.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Color As String
Private m_Opacity As Double


'### Properties

Public Property Let Color(ByVal newValue As String)
    m_Color = newValue
End Property
Public Property Get Color() As String
    Color = m_Color
End Property

Public Property Let Opacity(ByVal newValue As Double)
    m_Opacity = newValue
End Property
Public Property Get Opacity() As Double
    Opacity = m_Opacity
End Property


'### Instanzing


'### Public Functions


'### Generator Function

Friend Property Get ToCSS() As String

    'Declarations
    Dim strText As New BAPT_StringBuilder
    
    If Color <> vbNullString Then
        strText.Add "background-color: " & Me.Color & "; "
    End If
    
    If Opacity > 0 Then
        strText.Add "opacity: " & Str(Me.Opacity) & "; "
    End If
    
    ToCSS = strText.ToString

End Property

