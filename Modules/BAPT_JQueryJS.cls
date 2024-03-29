VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_JQueryJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021, 2023 Thomas M�ller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 2.05.05  published: 16.09.2023                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Source As ptScriptSource
Private m_CDN As String
Private m_PathFilename As String


' ### Properties

Public Property Let Source(ByVal newValue As ptScriptSource)
    m_Source = newValue
End Property
Public Property Get Source() As ptScriptSource
    Source = m_Source
End Property

Public Property Let CDN(ByVal newValue As String)
    m_CDN = newValue
End Property
Public Property Get CDN() As String
    CDN = m_CDN
End Property

Public Property Let PathFilename(ByVal newValue As String)
    m_PathFilename = newValue
End Property
Public Property Get PathFilename() As String
    PathFilename = m_PathFilename
End Property


'### Instanzing

Private Sub Class_Initialize()

    Me.Source = ptScriptSource.CDN
    Me.CDN = "https://cdnjs.cloudflare.com/ajax/libs/jquery/1.11.2/jquery.min.js"
    Me.PathFilename = CurrentProject.Path & "\jquery.min.js"
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


' ### Public Functions

Public Function SaveToDisk() As Boolean

    SaveToDisk = BAPT.Helper.SaveFileToDisk("jquery.min.js", CurrentProject.Path)
    
End Function


' ### Generator Function

Friend Function getLocation() As String

    'Declarations
    Dim strText As String
    
    Select Case Me.Source
        Case ptScriptSource.CDN
            strText = "'" & Me.CDN & "'"
        Case ptScriptSource.LocalFile
            strText = "'" & Me.PathFilename & "'"
        Case Else
            'Do nothing
    End Select
            
    getLocation = strText

End Function
