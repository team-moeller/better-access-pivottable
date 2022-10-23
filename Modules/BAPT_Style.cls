VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'### Members
Private m_Font As BAPT_Font

'### Properties
Public Property Get Font() As BAPT_Font
    Set Font = m_Font
End Property

'### Instanzing


Private Sub Class_Initialize()
    
    'Create sub classes
    Set m_Font = New BAPT_Font
End Sub


'### Generator Function

Friend Function ToCSS() As String
    Dim str As New BAPT_StringBuilder
    
    str.AddLine Font.ToCSS
    
    ToCSS = str.ToString
End Function
