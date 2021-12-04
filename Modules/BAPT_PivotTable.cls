VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_PivotTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.05.04  published: 04.12.2021                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum ptScriptSource
    CDN = 1
    LocalFile = 2
End Enum


'### Members

Private WithEvents m_Control As WebBrowserControl
Attribute m_Control.VB_VarHelpID = -1
Private m_IsInteractive As Boolean
Private m_DataSourceObjectName As String
Private m_Columns As String
Private m_Rows As String


'### Properties

Public Property Set Control(This_Control As WebBrowserControl)
    Set m_Control = This_Control
End Property
Public Property Get Control() As WebBrowserControl
    Set Control = m_Control
End Property

Public Property Let IsInteractive(newValue As Boolean)
    m_IsInteractive = newValue
End Property
Public Property Get IsInteractive() As Boolean
    IsInteractive = m_IsInteractive
End Property

Public Property Let DataSourceObjectName(newValue As String)
    m_DataSourceObjectName = newValue
End Property
Public Property Get DataSourceObjectName() As String
    DataSourceObjectName = m_DataSourceObjectName
End Property

Public Property Let Columns(newValue As String)
    newValue = "'" & newValue & "'"
    newValue = Replace(newValue, ",", "','")
    newValue = Replace(newValue, ",' ", ",'")
    m_Columns = newValue
End Property
Public Property Get Columns() As String
    Columns = m_Columns
End Property

Public Property Let Rows(newValue As String)
    newValue = "'" & newValue & "'"
    newValue = Replace(newValue, ",", "','")
    newValue = Replace(newValue, ",' ", ",'")
    m_Rows = newValue
End Property
Public Property Get Rows() As String
    Rows = m_Rows
End Property


'### Instanzing

Private Sub Class_Initialize()
    IsInteractive = True
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub ShowPivot()

    Dim Datei As String
    
    CreateHtmlFile
       
    Datei = CurrentProject.Path & "\BetterAccessPivotTable.html"
    
    Control.Object.Navigate Datei
    
End Sub

    
'### Initializer

Public Sub Init(Optional ByVal TheControl As WebBrowserControl)
      Set Me.Control = TheControl
End Sub


'### Private Methods

Private Sub CreateHtmlFile()

    'Variablen deklarieren
    Dim d As Integer
    Dim strPath As String
    Dim strFilename As String
    Dim strContent As String
    Const ForWriting As Long = 2
    
    strPath = CurrentProject.Path
    strFilename = "\BetterAccessPivotTable.html"
    
    strContent = strContent & getHead
    strContent = strContent & getData
    strContent = strContent & getScript
    strContent = strContent & getFoot
 
    d = FreeFile()
    Open strPath & strFilename For Output As d
    Print #d, strContent;
 
    Close d
  
End Sub

Private Function getHead() As String

    'Variablen deklarieren
    Dim strText As String
    
    strText = strText & "<!DOCTYPE html>" & vbCrLf
    strText = strText & "<!-- saved from url=(0014)about:internet -->" & vbCrLf
    strText = strText & "<!-- saved from url=(0016)http://localhost -->" & vbCrLf
    strText = strText & "<html>" & vbCrLf
    strText = strText & "<head>" & vbCrLf
    strText = strText & "    <meta charset='utf-8'>" & vbCrLf
    strText = strText & "    <meta http-equiv='X-UA-Compatible' content='IE=Edge'/>" & vbCrLf
    strText = strText & "    <title>PivotTable</title>" & vbCrLf
    strText = strText & "    <style>body{margin:0;padding:0;}</style>" & vbCrLf
    strText = strText & "    <link rel='stylesheet' href=" & BAPT.Scripts.PivotCSS.getLocation & ">" & vbCrLf
    strText = strText & "    <script src=" & BAPT.Scripts.JQueryJS.getLocation & "></script>" & vbCrLf
    strText = strText & "    <script src=" & BAPT.Scripts.JQueryUI.getLocation & "></script>" & vbCrLf
    strText = strText & "    <script src=" & BAPT.Scripts.PivotJS.getLocation & "></script>" & vbCrLf
    strText = strText & "</head>" & vbCrLf
    strText = strText & "<body>" & vbCrLf
    strText = strText & "    <div id='output' style='margin: 30px;'></div>" & vbCrLf
    
    getHead = strText
    
End Function

Private Function getData() As String

    'Variablen deklarieren
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim strText As String
    Dim strLine As String
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset(Me.DataSourceObjectName)

    strText = strText & "    <script>" & vbCrLf
    strText = strText & "    var data = [" & vbCrLf
    Do While Not rst.EOF
        strLine = "                {"
        For Each fld In rst.Fields
            strLine = strLine & "'" & fld.Name & "': '" & rst(fld.Name) & "', "
        Next fld
        strLine = Left$(strLine, Len(strLine) - 2)
        strText = strText & strLine & "}," & vbCrLf
        rst.MoveNext
    Loop
    strText = Left$(strText, Len(strText) - 3) & vbCrLf
    strText = strText & "               ]" & vbCrLf
    
    getData = strText

End Function

Private Function getScript() As String

    'Variablen deklarieren
    Dim strText As String
    
    strText = strText & "    $(function(){" & vbCrLf
    If Me.IsInteractive = True Then
        strText = strText & "        $('#output').pivotUI(" & vbCrLf
    Else
        strText = strText & "        $('#output').pivot(" & vbCrLf
    End If
    strText = strText & "            data," & vbCrLf
    strText = strText & "            {" & vbCrLf
    strText = strText & "                rows: [" & Me.Rows & "]," & vbCrLf
    strText = strText & "                cols: [" & Me.Columns & "]" & vbCrLf
    strText = strText & "            }" & vbCrLf
    
    getScript = strText

End Function

Private Function getFoot() As String

    'Variablen deklarieren
    Dim strText As String
    
    strText = strText & "        );" & vbCrLf
    strText = strText & "     });" & vbCrLf
    strText = strText & "    </script>" & vbCrLf
    strText = strText & "</body>" & vbCrLf
    strText = strText & "</html>"
    
    getFoot = strText
    
End Function
