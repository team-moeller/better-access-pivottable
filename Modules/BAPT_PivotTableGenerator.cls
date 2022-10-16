VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_PivotTableGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###############################################################################################
'# Copyright (c) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.26.02  published: 16.10.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Options  As BAPT_PivotTable


' ### Properties

Private Property Get Options() As BAPT_PivotTable
    Set Options = m_Options
End Property


'### Initializer

Friend Sub Init(ByVal PivotTableOptions As BAPT_PivotTable)

    Set m_Options = PivotTableOptions
    
End Sub



'### Public Methods

Friend Function GenerateHTML() As String

    'Declarations
    Dim strHTML As New BAPT_StringBuilder

    With strHTML
        .AddLine getHead
        .AddLine Options.DataSource.getData
        .AddLine getScript
        .AddLine getFoot
    End With
    GenerateHTML = strHTML.ToString()
 
End Function


'### Private Methods / Functions

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
        If Options.Localization.Locale <> "en" Then
            .AddLine "    <script src=" & Options.Localization.getLocation & "></script>"
        End If
        If Options.Columns.ShowTotals = False Then
            .AddLine Options.Columns.GetCssToHideTotals
        End If
        If Options.Rows.ShowTotals = False Then
            .AddLine Options.Rows.GetCssToHideTotals
        End If
        .AddLine "</head>"
        .AddLine "<body>"
        .AddLine "    <div id='output' style='margin: 30px;'></div>"
    End With
    
    getHead = strText.ToString
    
End Function

Private Function getScript() As String

    'Declarations
    Dim strText As New BAPT_StringBuilder
    
    With strText
        .AddLine "    $(function(){"
        .AddLine "        $('#output').pivotUI("
        .AddLine "            data,"
        .AddLine "            {"
        .AddLine "                rows: [" & Options.Rows.Fields & "],"
        .AddLine "                rowOrder: " & BAPT.Helper.getOrderString(Options.Rows.Order) & ","
        .AddLine "                cols: [" & Options.Columns.Fields & "],"
        .AddLine "                colOrder: " & BAPT.Helper.getOrderString(Options.Columns.Order) & ","
        .AddLine "                aggregatorName: " & Options.Aggregator & ","
        .AddLine "                vals: [" & Options.AggregatorField & "],"
        .AddLine "                rendererName: [" & Options.RendererName & "],"
        If Options.IsInteractive = False Then
            .AddLine "                showUI: false,"
        End If
        .CutTrailingComma
        .AddLine "            },"
        If Options.Localization.Locale <> "en" Then
            .AddLine "            false,"
            .AddLine "            '" & Options.Localization.Locale & "'"
        End If
        .CutTrailingComma
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
