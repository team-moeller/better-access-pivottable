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
'# Copyright (c) 2021, 2023 Thomas M�ller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 2.05.05  published: 16.09.2023                                                      #
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

Private m_WebBrowserControl As Access.Edge
Private m_WebBrowserActiveX As SHDocVw.WebBrowser
Private m_IsInteractive As Boolean
Private m_SaveHtmlFileForDebugging As Boolean
Private m_Aggregator As String
Private m_AggregatorField As String
Private m_RendererName As String
Private m_DisplayIn As chDisplayIn
Private m_Margin As Long

Private m_Rows As BAPT_Rows
Private m_Columns As BAPT_Columns
Private m_DataSource As BAPT_DataSource
Private m_Localization As BAPT_Localization
Private m_Styles As BAPT_Styles
    
    
'### Properties

Public Property Set Control(ByVal New_Control As Object)
    If TypeName(New_Control) = "Edge" Then
        Set Me.WebBrowserControl = New_Control
    ElseIf TypeName(New_Control.Object) = "WebBrowser" Then
        Set Me.WebBrowserActiveX = New_Control.Object
    Else
        Err.Raise vbObjectError, , "Type of control not supported!"
    End If
End Property

Public Property Set WebBrowserControl(ByVal New_Control As Access.Edge)
    Set m_WebBrowserControl = New_Control
    m_DisplayIn = chWebBrowserControl
End Property
Public Property Get WebBrowserControl() As Edge
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

Public Property Let Margin(ByVal newValue As Long)
    m_Margin = newValue
End Property
Public Property Get Margin() As Long
    Margin = m_Margin
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
    Me.Margin = 30
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub ShowPivot()

    'Declarations
    Dim strHtml As String
    Dim js As String
    
    Select Case Me.DisplayIn
        Case chDisplayIn.chWebBrowserControl
            If Me.WebBrowserControl.LocationURL = vbNullString Then
                CreateEmptyHtmlFile
                Me.WebBrowserControl.Navigate "https://msaccess/" & getHTMLFileName(, True)
                Do While Me.WebBrowserControl.ReadyState <> 4
                    DoEvents
                Loop
            End If
            strHtml = BAPT.Generator(Me).GenerateHTML
            If Me.SaveHtmlFileForDebugging = True Then
                Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
            End If
            strHtml = Replace(strHtml, vbCrLf, "")
            js = js & "document.open();"
            js = js & "document.write(""" & strHtml & """);"
            js = js & "document.close();"
            Me.WebBrowserControl.ExecuteJavascript js
        Case chDisplayIn.chWebBrowserActiveX
            CreateEmptyHtmlFile
            Me.WebBrowserActiveX.Navigate getHTMLFileName(, True)
            Do While Me.WebBrowserActiveX.Busy Or Me.WebBrowserActiveX.ReadyState <> 4
                DoEvents
            Loop
            With Me.WebBrowserActiveX.Document
                .Open
                strHtml = BAPT.Generator(Me).GenerateHTML
                .Write strHtml
                .Close
            End With
            If Me.SaveHtmlFileForDebugging = True Then
                Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
            End If
        Case chDisplayIn.chSystemBrowser
            CreateHtmlFile
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

Private Sub CreateEmptyHtmlFile()

    'Decalarations
    Dim strHtml As New BAPT_StringBuilder
    
    With strHtml
        .AddLine "<html>"
        .AddLine "<head>"
        .AddLine "</head>"
        .AddLine "<body>"
        .AddLine "</body>"
        .AddLine "</html>"
    End With
    
    Call saveHtmlFile(vbNullString, strHtml.ToString, True)
    
End Sub

Private Sub CreateHtmlFile()

    'Declarations
    Dim strHtml As String
    
    With BAPT.Generator(Me)
        strHtml = .GenerateHTML
    End With
 
    If Me.SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
    End If
    
    Call saveHtmlFile(BAPT.Helper.getTempFolder, strHtml)
    
End Sub


'### Private Helper Methods

Private Sub saveHtmlFile(ByVal Folder As String, ByVal strHtml As String, Optional ByVal fEmpty As Boolean = False)

    'Declarations
    Dim objFS As Object
    Dim FileName As String
       
    FileName = getHTMLFileName(Folder, fEmpty)
    
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHtml
    objFS.SaveToFile FileName, 2   '2: Create Or Update
    objFS.Close
    
End Sub

Private Function getHTMLFileName(Optional Path As String = vbNullString, Optional fEmpty As Boolean = False) As String
    
    'Declarations
    Dim strFilename As String
    Const HtmlFileName As String = "BetterAccessPivotTable_{}.html"

    If Path = vbNullString Then
        Path = BAPT.Helper.getTempFolder
    End If
    If fEmpty = True Then
        strFilename = "Empty.html"
    Else
        strFilename = Replace(HtmlFileName, "{}", ObjPtr(Me))
    End If

    getHTMLFileName = Path & strFilename
    
End Function

