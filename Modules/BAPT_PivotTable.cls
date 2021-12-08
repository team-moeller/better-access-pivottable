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
'# Version 1.07.05  published: 08.12.2021                                                      #
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
Private m_SaveHtmlFileForDebugging As Boolean
Private m_ShowInBrowser As Boolean


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

Public Property Let SaveHtmlFileForDebugging(ByVal This_SaveHtmlFileForDebugging As Boolean)
    m_SaveHtmlFileForDebugging = This_SaveHtmlFileForDebugging
End Property
Public Property Get SaveHtmlFileForDebugging() As Boolean
    SaveHtmlFileForDebugging = m_SaveHtmlFileForDebugging
End Property

Public Property Let ShowInBrowser(newValue As Boolean)
    m_ShowInBrowser = newValue
End Property
Public Property Get ShowInBrowser() As Boolean
    ShowInBrowser = m_ShowInBrowser
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

    'Declarations
    Dim m_IE As Object
    Dim Doc As Object
    
    If Me.ShowInBrowser = True Then
        Set m_IE = CreateObject("InternetExplorer.Application")
        m_IE.Navigate "about:"
        m_IE.Visible = True
        Do While m_IE.ReadyState <> 4: DoEvents: Loop
        Set Doc = m_IE.Document
    Else
        Set Doc = Control.Object.Document
    End If
    
    With Doc
        .Open
        .Write getHtmlString
        .Close
    End With
    
End Sub

    
'### Initializer

Public Sub Init(Optional ByVal TheControl As WebBrowserControl)
      Set Me.Control = TheControl
End Sub


'### Private Methods

Private Function getHtmlString()

    'Declarations
    Dim strHtml As New BAPT_StringBuilder
    
    With strHtml
        .AddLine getHead
        .AddLine getData
        .AddLine getScript
        .AddLine getFoot
    End With
 
    If m_SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(strHtml.ToString)
    End If
 
    getHtmlString = strHtml.ToString
      
End Function

Private Function getHead() As String

    'Declarations
    Dim strText As New BAPT_StringBuilder
    
    With strText
        .AddLine "<!DOCTYPE html>"
        .AddLine "<!-- saved from url=(0014)about:internet -->"
        .AddLine "<!-- saved from url=(0016)http://localhost -->"
        .AddLine "<html>"
        .AddLine "<head>"
        .AddLine "    <meta charset='utf-8'>"
        .AddLine "    <meta http-equiv='X-UA-Compatible' content='IE=Edge'/>"
        .AddLine "    <title>PivotTable</title>"
        .AddLine "    <style>body{margin:0;padding:0;}</style>"
        .AddLine "    <link rel='stylesheet' href=" & BAPT.Scripts.PivotCSS.getLocation & ">"
        .AddLine "    <script src=" & BAPT.Scripts.JQueryJS.getLocation & "></script>"
        .AddLine "    <script src=" & BAPT.Scripts.JQueryUI.getLocation & "></script>"
        .AddLine "    <script src=" & BAPT.Scripts.PivotJS.getLocation & "></script>"
        .AddLine "</head>"
        .AddLine "<body>"
        .AddLine "    <div id='output' style='margin: 30px;'></div>"
    End With
    
    getHead = strText.ToString
    
End Function

Private Function getData() As String

    'Declarations
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim strText As New BAPT_StringBuilder
    Dim strLine As String
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset(Me.DataSourceObjectName)

    With strText
        .AddLine "    <script>"
        .AddLine "    var data = ["
        Do While Not rst.EOF
            strLine = "                {"
            For Each fld In rst.Fields
                strLine = strLine & "'" & fld.Name & "': '" & rst(fld.Name) & "', "
            Next fld
            strLine = Left$(strLine, Len(strLine) - 2)
            .AddLine strLine & "},"
            rst.MoveNext
        Loop
        .CutTrailingComma
        .AddLine "               ]"
    End With
    
    getData = strText.ToString

End Function

Private Function getScript() As String

    'Declarations
    Dim strText As New BAPT_StringBuilder
    
    With strText
        .AddLine "    $(function(){"
        If Me.IsInteractive = True Then
            .AddLine "        $('#output').pivotUI("
        Else
            .AddLine "        $('#output').pivot("
        End If
        .AddLine "            data,"
        .AddLine "            {"
        .AddLine "                rows: [" & Me.Rows & "],"
        .AddLine "                cols: [" & Me.Columns & "]"
        .AddLine "            }"
    End With
    
    getScript = strText.ToString

End Function

Private Function getFoot() As String

    'Declarations
    Dim strText As New BAPT_StringBuilder
    
    With strText
        .AddLine "        );"
        .AddLine "     });"
        .AddLine "    </script>"
        .AddLine "</body>"
        .AddLine "</html>"
    End With
    
    getFoot = strText.ToString
    
End Function


'### Private Helper Methods

Private Sub saveHtmlFile(ByVal strContent As String)

    'Declarations
    Dim strPath As String
    Dim strFileName As String
    Dim f As Integer
    Const HtmlFileName As String = "BetterAccessChart_{}.html"
    
    strPath = CurrentProject.Path & "\"
    strFileName = Replace(HtmlFileName, "{}", m_Control.Name)
    
    f = FreeFile()
    Open strPath & strFileName For Output As f
    Print #f, strContent;
    Close f
    
End Sub

