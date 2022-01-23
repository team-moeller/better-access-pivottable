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
'# Copyright (c) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.17.09  published: 23.01.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum ptScriptSource
    CDN = 1
    LocalFile = 2
End Enum

Public Enum ptOrder
    key_a_to_z = 1
    value_a_to_z = 2
    value_z_to_a = 3
End Enum


'### Constants

Private Const HtmlFileName As String = "BetterAccessPivotTable_{}.html"


'### Members

Private WithEvents m_Control As WebBrowserControl
Attribute m_Control.VB_VarHelpID = -1
Private m_IsInteractive As Boolean
Private m_DataSourceObjectName As String
Private m_SaveHtmlFileForDebugging As Boolean
Private m_Aggregator As String
Private m_AggregatorField As String
Private m_RendererName As String

Private m_Rows As BAPT_Rows
Private m_Columns As BAPT_Columns
    
    
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

Public Property Let SaveHtmlFileForDebugging(ByVal This_SaveHtmlFileForDebugging As Boolean)
    m_SaveHtmlFileForDebugging = This_SaveHtmlFileForDebugging
End Property
Public Property Get SaveHtmlFileForDebugging() As Boolean
    SaveHtmlFileForDebugging = m_SaveHtmlFileForDebugging
End Property

Public Property Let Aggregator(newValue As String)
    m_Aggregator = "'" & newValue & "'"
End Property
Public Property Get Aggregator() As String
    Aggregator = m_Aggregator
End Property

Public Property Let AggregatorField(newValue As String)
    m_AggregatorField = "'" & newValue & "'"
End Property
Public Property Get AggregatorField() As String
    AggregatorField = m_AggregatorField
End Property

Public Property Let RendererName(newValue As String)
    m_RendererName = "'" & newValue & "'"
End Property
Public Property Get RendererName() As String
    RendererName = m_RendererName
End Property


'### Object Properties

Public Property Get Rows() As BAPT_Rows
    Set Rows = m_Rows
End Property

Public Property Get Columns() As BAPT_Columns
    Set Columns = m_Columns
End Property


'### Instanzing

Private Sub Class_Initialize()
    
    'Create sub classes
    Set m_Rows = New BAPT_Rows
    Set m_Columns = New BAPT_Columns
    
    'Set default values
    Me.IsInteractive = True
    Me.Aggregator = "Count"
    Me.RendererName = "Table"
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub ShowPivot()
    
    CreateHtmlFile
        
    Control.Object.Navigate getHTMLFileName
    
End Sub

Public Sub ShowInBrowser()
    
    CreateHtmlFile
        
    BAPT.Helper.ShellExecute getHTMLFileName
    
End Sub

    
'### Initializer

Public Sub Init(Optional ByVal TheControl As WebBrowserControl)
      Set Me.Control = TheControl
End Sub


'### Private Methods

Private Sub CreateHtmlFile()

    'Declarations
    Dim objFS As Object
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
        
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHtml.ToString
    objFS.SaveToFile getHTMLFileName, 2   '2: Create Or Update
    objFS.Close
      
End Sub

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
        If Me.Columns.ShowTotals = False Then
            .AddLine Me.Columns.GetCssToHideTotals
        End If
        If Me.Rows.ShowTotals = False Then
            .AddLine Me.Rows.GetCssToHideTotals
        End If
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
        .AddLine "        $('#output').pivotUI("
        .AddLine "            data,"
        .AddLine "            {"
        .AddLine "                rows: [" & Me.Rows.Fields & "],"
        .AddLine "                rowOrder: " & BAPT.Helper.getOrderString(Me.Rows.Order) & ","
        .AddLine "                cols: [" & Me.Columns.Fields & "],"
        .AddLine "                colOrder: " & BAPT.Helper.getOrderString(Me.Columns.Order) & ","
        .AddLine "                aggregatorName: " & Me.Aggregator & ","
        .AddLine "                vals: [" & Me.AggregatorField & "],"
        .AddLine "                rendererName: [" & Me.RendererName & "],"
        If Me.IsInteractive = False Then
            .AddLine "                showUI: false,"
        End If
        .CutTrailingComma
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
    Dim strFilename As String
    Dim f As Integer
    Const HtmlFileName As String = "BetterAccessPivotTable_{}.html"
    
    strPath = CurrentProject.Path & "\"
    strFilename = Replace(HtmlFileName, "{}", m_Control.Name)
    
    f = FreeFile()
    Open strPath & strFilename For Output As f
    Print #f, strContent;
    Close f
    
End Sub

Private Function getHTMLFileName() As String
    
    'Declarations
    Dim strPath As String
    Dim strFilename As String
    
    strPath = BAPT.Helper.getTempFolder
    strFilename = Replace(HtmlFileName, "{}", m_Control.Name)

    getHTMLFileName = strPath & strFilename
    
End Function
