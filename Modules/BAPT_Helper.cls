VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###############################################################################################
'# Copyright (c) 2021, 2023 Thomas M�ller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 2.05.05  published: 16.09.2023                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Windows API Declarations

#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As LongPtr, ByVal lpbuffer As String) As Long
    Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
    Private Declare PtrSafe Function ShellExecuteA Lib "Shell32" _
        (ByVal hWnd As LongPtr, ByVal lpOperation As String, _
         ByVal lpFile As String, ByVal lpParameters As String, _
         ByVal lpDirectory As String, ByVal nShowCmd As eShowCmd) As Long
#Else
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, ByVal lpbuffer As String) As Long
    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
    Private Declare Function ShellExecuteA Lib "Shell32" _
        (ByVal hWnd As Long, ByVal lpOperation As String, _
         ByVal lpFile As String, ByVal lpParameters As String, _
         ByVal lpDirectory As String, ByVal nShowCmd As eShowCmd) As Long
#End If


'### Enumerations

Public Enum eShowCmd
    SW_HIDE = 0             'Blendet Fenster aus und aktiviert ein anderes Fenster
    SW_MAXIMIZE = 3         'Maximiert Fenster und aktiviert es
    SW_MINIMIZE = 6         'Minimiert Fenster und aktiviert ein anderes Fenster
    SW_RESTORE = 9          'Aktiviert bisher minimiertes Fenster in aktueller Gr��e
    SW_SHOW = 5             'Aktiviert Fenster in aktueller Gr��e
    SW_SHOWDEFAULT = 10     'Aktiviert Fenster in urspr�nglicher Gr��e
    SW_SHOWMINIMIZED = 2    'Minimiert Fenster und aktiviert es
    SW_SHOWMININOACTIVE = 7 'Minimiert Fenster. Das aktive Fenster bleibt aktiv
    SW_SHOWNA = 8           'Fenster in aktueller Gr��e. Aktives Fenster bleibt aktiv
    SW_SHOWNOACTIVATE = 4   'Fenster in neuer Gr��e. Das aktive Fenster bleibt aktiv
    SW_SHOWNORMAL = 1       'Aktiviert Fenster. Falls schon min/maximiert, normale Gr��e
End Enum


'### Functions

Friend Function getOrderString(ByVal ThisOrder As ptOrder) As String

    Select Case ThisOrder
        Case ptOrder.key_a_to_z
            getOrderString = "'key_a_to_z'"
        Case ptOrder.value_a_to_z
            getOrderString = "'value_a_to_z'"
        Case ptOrder.value_z_to_a
            getOrderString = "'value_z_to_a'"
        Case Else
            getOrderString = "'key_a_to_z'"
    End Select

End Function

Friend Function CutTrailingComma(ByVal strText As String) As String

    If Right$(strText, 1) = "," Then
        CutTrailingComma = Left$(strText, Len(strText) - 1)
    ElseIf Right$(strText, 3) = "," & vbCrLf Then
        CutTrailingComma = Left$(strText, Len(strText) - 3) & vbCrLf
    Else
        CutTrailingComma = strText
    End If

End Function

Friend Function SaveFileToDisk(ByVal FileName As String, ByVal Path As String) As Boolean
On Error GoTo Handle_Error

    'Declarations
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim FileID As Long
    Dim Buffer() As Byte
    Dim FileLen As Long
    Dim Success As Boolean

    Set cnn = CurrentProject.Connection
    Set rst = New ADODB.Recordset
    rst.Open "Select FileData FROM USys_FileData Where FileName='" & FileName & "'", _
        cnn, adOpenDynamic, adLockOptimistic

    FileID = FreeFile
    FileLen = Nz(LenB(rst!FileData), 0)

    If FileLen > 0 Then
        ReDim Buffer(FileLen)
        MakeSureDirectoryPathExists (Path & "\")
        Open Path & "\" & FileName For Binary Access Write As FileID
        Buffer = rst!FileData.GetChunk(FileLen)
        Put FileID, , Buffer
        Close FileID
    End If
    Success = True

Exit_Here:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set cnn = Nothing
    SaveFileToDisk = Success
    Exit Function

Handle_Error:
    Select Case Err.Number
        Case 0
            Resume
        Case Else
            MsgBox Err.Description, vbExclamation, Err.Number
            Resume Exit_Here
    End Select

End Function

Friend Function getTempFolder() As String
    
    'Declaration
    Dim RetVal As Long
    Dim TempDir As String
 
    TempDir = Space$(256)
    RetVal = GetTempPath(Len(TempDir), TempDir)
 
    If RetVal <> 0 Then
        getTempFolder = Left$(TempDir, RetVal)
    End If
    
End Function

Friend Sub ShellExecute(ByVal FileName As String)

    Call ShellExecuteA(Application.hWndAccessApp, "open", FileName, vbNullString, CurDir$, SW_SHOW)

End Sub


