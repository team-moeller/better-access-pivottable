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
'# Version 1.20.07  published: 19.03.2022                                                      #
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

Public Enum chDisplayIn
    chWebBrowserControl = 1
    chSystemBrowser = 2
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
Private m_DisplayIn As chDisplayIn

Private m_Rows As BAPT_Rows
Private m_Columns As BAPT_Columns
    
    
'### Properties

Public Property Set Control(ByRef This_Control As WebBrowserControl)
    Set m_Control = This_Control
    m_DisplayIn = chWebBrowserControl
End Property
Public Property Get Control() As WebBrowserControl
    Set Control = m_Control
End Property

Public Property Let IsInteractive(ByVal newValue As Boolean)
    m_IsInteractive = newValue
End Property
Public Property Get IsInteractive() As Boolean
    IsInteractive = m_IsInteractive
End Property

Public Property Let DataSourceObjectName(ByVal newValue As String)
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

Public Property Let Aggregator(ByVal newValue As String)
    m_Aggregator = "'" & newValue & "'"
End Property
Public Property Get Aggregator() As String
    Aggregator = m_Aggregator
End Property

Public Property Let AggregatorField(ByVal newValue As String)
    m_AggregatorField = "'" & newValue & "'"
End Property
Public Property Get AggregatorField() As String
    AggregatorField = m_AggregatorField
End Property

Public Property Let RendererName(ByVal newValue As String)
    m_RendererName = "'" & newValue & "'"
End Property
Public Property Get RendererName() As String
    RendererName = m_RendererName
End Property

Public Property Let DisplayIn(ByVal newValue As chDisplayIn)
    m_DisplayIn = newValue
End Property
Public Property Get DisplayIn() As chDisplayIn
    DisplayIn = m_DisplayIn
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
    Me.DisplayIn = chWebBrowserControl
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub ShowPivot()

    CreateHtmlFile
    
    Select Case Me.DisplayIn
        Case chDisplayIn.chWebBrowserControl
            Control.Object.Navigate getHTMLFileName
        Case chDisplayIn.chSystemBrowser
            BAPT.Helper.ShellExecute getHTMLFileName
        Case Else
            'Do nothing
    End Select
    
End Sub
    
    
'### Initializer

Public Sub Init(Optional ByVal TheControl As WebBrowserControl)
      Set Me.Control = TheControl
End Sub


'### Private Methods

Private Sub CreateHtmlFile()

    'Declarations
    Dim objFS As Object
    Dim strHtml As String
 
    With BAPT.Generator(Me)
        strHtml = .GenerateHTML
    End With
 
    If m_SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(strHtml)
    End If
        
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHtml
    objFS.SaveToFile getHTMLFileName, 2   '2: Create Or Update
    objFS.Close
      
End Sub


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
