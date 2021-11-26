Attribute VB_Name = "mdl_Helpers"
'###############################################################################################
'# Copyright (c) 2021 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.03.05  published: 26.11.2021                                                      #
'###############################################################################################

Option Compare Database
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#Else
    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#End If

Public Function IsFormOpen(ByVal strFormName As String) As Boolean

    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0 Then
        If Forms(strFormName).CurrentView <> 0 Then
            IsFormOpen = True
        End If
    End If
    
End Function


Public Sub PrepareAndExportModules()

    'Declarations
    Dim Version As String
    Dim CodeLine As String
    Dim vbc As Object
    
    MakeSureDirectoryPathExists CurrentProject.Path & "\Modules\"
    Version = DLast("V_Number", "tbl_VersionHistory")
    CodeLine = "'# Version " & Version & "  published: " & Format$(Date, "dd.mm.yyyy") & "                                                      #"
    
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        If vbc.Type = 1 Or vbc.Type = 2 Then
            Application.VBE.ActiveVBProject.VBComponents(vbc.Name).CodeModule.InsertLines 4, CodeLine
            Application.VBE.ActiveVBProject.VBComponents(vbc.Name).CodeModule.DeleteLines 5, 1
    
            Application.VBE.ActiveVBProject.VBComponents(vbc.Name).Export CurrentProject.Path & "\Modules\" & vbc.Name & IIf(vbc.Type = 2, ".cls", ".bas")
        End If
    Next
    Application.DoCmd.RunCommand (acCmdCompileAndSaveAllModules)
    
    MsgBox "Export done", vbInformation, "Better Access Charts"

End Sub

Public Sub ImportModules()

    'Declarations
    Dim strFile As String
    Dim vbc As Object
    
    strFile = Dir(CurrentProject.Path & "\Modules\")
    Do While Len(strFile) > 0
        On Error Resume Next
        Set vbc = Application.VBE.ActiveVBProject.VBComponents(strFile)
        Application.VBE.ActiveVBProject.VBComponents.Remove vbc
        On Error GoTo 0
        Application.VBE.ActiveVBProject.VBComponents.Import CurrentProject.Path & "\Modules\" & strFile
        Debug.Print strFile
        strFile = Dir
    Loop
    Application.DoCmd.RunCommand (acCmdCompileAndSaveAllModules)
    
    MsgBox "Import done", vbInformation, "Better Access Charts"

End Sub

