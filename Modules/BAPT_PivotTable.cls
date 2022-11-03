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
'# Version 1.52.01  published: 03.11.2022                                                      #
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
    chWebBrowserActiveX = 2
    chSystemBrowser = 3
End Enum


'### Members

Private m_WebBrowserControl As Access.WebBrowserControl
Private m_WebBrowserActiveX As SHDocVw.WebBrowser
Private m_IsInteractive As Boolean
Private m_SaveHtmlFileForDebugging As Boolean
Private m_Aggregator As String
Private m_AggregatorField As String
Private m_RendererName As String
Private m_DisplayIn As chDisplayIn

Private m_Rows As BAPT_Rows
Private m_Columns As BAPT_Columns
Private m_DataSource As BAPT_DataSource
Private m_Localization As BAPT_Localization
Private m_Styles As BAPT_Styles
    
    
'### Properties

Public Property Set Control(ByVal New_Control As Object)
    If TypeName(New_Control) = "WebBrowserControl" Then
        Set Me.WebBrowserControl = New_Control
    ElseIf TypeName(New_Control.Object) = "WebBrowser" Then
        Set Me.WebBrowserActiveX = New_Control.Object
    Else
        Err.Raise vbObjectError, , "Type of control not supported!"
    End If
End Property

Public Property Set WebBrowserControl(ByVal New_Control As Access.WebBrowserControl)
    Set m_WebBrowserControl = New_Control
    m_DisplayIn = chWebBrowserControl
End Property
Public Property Get WebBrowserControl() As WebBrowserControl
    Set WebBrowserControl = m_WebBrowserControl
End Property

Public Property Set WebBrowserActiveX(ByVal New_Control As SHDocVw.WebBrowser)
    Set m_WebBrowserActiveX = New_Control
    m_DisplayIn = chWebBrowserActiveX
End Property
Public Property Get WebBrowserActiveX() As SHDocVw.WebBrowser
    Set WebBrowserActiveX = m_WebBrowserActiveX
End Property

Public Property Let IsInteractive(ByVal newValue As Boolean)
    m_IsInteractive = newValue
End Property
Public Property Get IsInteractive() As Boolean
    IsInteractive = m_IsInteractive
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

Public Property Get DataSource() As BAPT_DataSource
    Set DataSource = m_DataSource
End Property

Public Property Get Localization() As BAPT_Localization
    Set Localization = m_Localization
End Property

Public Property Get Styles() As BAPT_Styles
    Set Styles = m_Styles
End Property


'### Instanzing

Private Sub Class_Initialize()
    
    'Create sub classes
    Set m_Rows = New BAPT_Rows
    Set m_Columns = New BAPT_Columns
    Set m_DataSource = New BAPT_DataSource
    Set m_Localization = New BAPT_Localization
    Set m_Styles = New BAPT_Styles
    
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
            Me.WebBrowserControl.Object.Navigate getHTMLFileName
        Case chDisplayIn.chWebBrowserActiveX
            Me.WebBrowserActiveX.Navigate getHTMLFileName
        Case chDisplayIn.chSystemBrowser
            BAPT.Helper.ShellExecute getHTMLFileName
        Case Else
            'Do nothing
    End Select
    
End Sub
    
    
'### Initializer

Public Sub Init(Optional ByVal TheControl As Object)
      Set Me.Control = TheControl
End Sub


'### Private Methods

Private Sub CreateHtmlFile()

    'Declarations
    Dim strHTML As String
    
    With BAPT.Generator(Me)
        strHTML = .GenerateHTML
    End With
 
    If Me.SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(CurrentProject.Path & "\", strHTML)
    End If
    
    Call saveHtmlFile(BAPT.Helper.getTempFolder, strHTML)
    
End Sub


'### Private Helper Methods

Private Sub saveHtmlFile(ByVal Folder As String, ByVal strHTML As String)

    'Declarations
    Dim objFS As Object
    Dim FileName As String
       
    FileName = getHTMLFileName(Folder)
    
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHTML
    objFS.SaveToFile FileName, 2   '2: Create Or Update
    objFS.Close
     
End Sub

Private Function getHTMLFileName(Optional Path As String = vbNullString) As String
    
    'Declarations
    Dim strFilename As String
    Const HtmlFileName As String = "BetterAccessPivotTable_{}.html"

    If Path = vbNullString Then
        Path = BAPT.Helper.getTempFolder
    End If
    strFilename = Replace(HtmlFileName, "{}", ObjPtr(Me))

    getHTMLFileName = Path & strFilename
    
End Function
