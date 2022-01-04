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
'# Copyright (c) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.15.05  published: 04.01.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Windows API Declarations
#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" _
            Alias "GetTempPathA" (ByVal nBufferLength As LongPtr, _
                                  ByVal lpbuffer As String) As Long
#Else
    Private Declare Function GetTempPath Lib "kernel32" _
            Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                  ByVal lpbuffer As String) As Long
#End If
    

#If VBA7 Then
    Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#Else
    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#End If

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
    
    'Declarations
    Dim strTempPath As String
    Dim lngTempPath As Long
    
    strTempPath = String(144, vbNullChar)
    lngTempPath = Len(strTempPath)
    If (GetTempPath(lngTempPath, strTempPath) > 0) Then
        getTempFolder = TrimNull(strTempPath)
    Else
        getTempFolder = vbNullString
    End If
    
End Function

Private Function TrimNull(ByVal strString As String) As String

    'Declarations
    Dim intPos As Integer

    intPos = VBA.InStr(strString, vbNullChar)
    If intPos > 0 Then
        TrimNull = VBA.Left$(strString, intPos - 1)
    Else
        TrimNull = strString
    End If

End Function

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

