VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_JQueryUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.04.39  published: 30.11.2021                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Source As ptScriptSource
Private m_CDN As String
Private m_PathFilename As String


' ### Properties

Public Property Let Source(ByVal This_Source As ptScriptSource)
    m_Source = This_Source
End Property
Public Property Get Source() As ptScriptSource
    Source = m_Source
End Property

Public Property Let CDN(ByVal This_CDN As String)
    m_CDN = This_CDN
End Property
Public Property Get CDN() As String
    CDN = m_CDN
End Property

Public Property Let PathFilename(ByVal This_PathFilename As String)
    m_PathFilename = This_PathFilename
End Property
Public Property Get PathFilename() As String
    PathFilename = m_PathFilename
End Property


'### Instanzing

Private Sub Class_Initialize()
    Me.Source = ptScriptSource.CDN
    Me.CDN = "https://cdnjs.cloudflare.com/ajax/libs/jqueryui/1.11.4/jquery-ui.min.js"
    Me.PathFilename = CurrentProject.Path & "\jquery-ui.min.js"
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


' ### Public Functions

Public Function SaveToDisk() As Boolean

    SaveToDisk = BAPT.Helper.SaveFileToDisk("jquery-ui.min.js", CurrentProject.Path)
    
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
