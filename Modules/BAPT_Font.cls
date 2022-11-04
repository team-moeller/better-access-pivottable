VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021, 2022 Thomas Möller, supported by K.D.Gundermann                         #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.53.04  published: 04.11.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Color As String
Private m_Family As String
Private m_Size As Integer
Private m_IsItalic As Boolean
Private m_IsBold As Boolean


'### Properties

Public Property Let Color(ByVal newValue As String)
    m_Color = newValue
End Property
Public Property Get Color() As String
    Color = m_Color
End Property

Public Property Let Family(ByVal newValue As String)
    m_Family = newValue
End Property
Public Property Get Family() As String
    Family = m_Family
End Property

Public Property Let Size(ByVal newValue As Integer)
    m_Size = newValue
End Property
Public Property Get Size() As Integer
    Size = m_Size
End Property

Public Property Let IsItalic(ByVal newValue As Boolean)
    m_IsItalic = newValue
End Property
Public Property Get IsItalic() As Boolean
    IsItalic = m_IsItalic
End Property

Public Property Let IsBold(ByVal newValue As Boolean)
    m_IsBold = newValue
End Property
Public Property Get IsBold() As Boolean
    IsBold = m_IsBold
End Property


'### Instanzing


'### Public Functions


'### Generator Function

Friend Property Get ToCSS() As String

    'Declarations
    Dim strText As New BAPT_StringBuilder
    
    If Color <> vbNullString Then
        strText.Add "color: " & Color & ";"
    End If
    If Family <> vbNullString Then
        strText.Add "font-family: " & Family & ";"
    End If
    If Size > 0 Then
        strText.Add "font-size: " & Size & "px;"
    End If
    If IsItalic Then
        strText.Add "font-style: italic;"
    End If
    If IsBold Then
        strText.Add "font-weight: bold;"
    End If
    
    ToCSS = strText.ToString

End Property
