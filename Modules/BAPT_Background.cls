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
'# Copyright (c) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.50.08  published: 01.11.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit

Private m_Color As String
Private m_Opacity As Double

' ### Properties

Public Property Let Color(ByVal This_Color As String)
    m_Color = This_Color
End Property
Public Property Get Color() As String
    Color = m_Color
End Property

Public Property Let Opacity(ByVal This_Value As Double)
    m_Opacity = This_Value
End Property
Public Property Get Opacity() As Double
    Opacity = m_Opacity
End Property


' ### Functions

Friend Property Get ToCSS() As String
    Dim strText As New BAPT_StringBuilder
    
    If Color <> vbNullString Then
        strText.Add "background-color: " & Color & ";"
    End If
    If Opacity > 0 Then
        strText.Add "opacity: " & Str(Opacity) & ";"
    End If
    
    ToCSS = strText.ToString

End Property
