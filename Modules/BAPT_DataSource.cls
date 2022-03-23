VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAPT_DataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2021, 2022 Thomas Möller                                                      #
'# MIT License  => https://github.com/team-moeller/better-access-pivottable/blob/main/LICENSE  #
'# Version 1.23.17  published: 23.03.2022                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum chDataSourceType
    dstDemo
    dstTableName
    dstQueryName
    dstSQLStament
    dstRecordset
End Enum


'### Members

Private m_Type As chDataSourceType
Private m_ObjectName As String
Private m_SQLStatement As String
Private m_Recordset As Object


' ### Properties

Public Property Let ObjectName(ByVal This_ObjectName As String)
    m_Type = dstTableName
    m_ObjectName = This_ObjectName
End Property
Public Property Get ObjectName() As String
    ObjectName = m_ObjectName
End Property

Public Property Let SQLStatement(ByVal This_SQLStatement As String)
    m_Type = dstSQLStament
    m_SQLStatement = This_SQLStatement
End Property
Public Property Get SQLStatement() As String
    SQLStatement = m_SQLStatement
End Property

Public Property Set Recordset(ByVal This_Recordset As Object)  ' ADODB/DAO.Recordset
    m_Type = dstRecordset
    Set m_Recordset = This_Recordset
End Property
Public Property Get Recordset() As Object   ' ADODB/DAO.Recordset
    Set Recordset = m_Recordset
End Property


' ### Initializer

Friend Sub Init(Optional ByVal DataSourceType As chDataSourceType = dstDemo, _
                Optional ByVal DataSource As Variant)
  
    Select Case DataSourceType
        Case dstDemo:       ' will be created later
        Case dstTableName:  Me.ObjectName = DataSource
        Case dstQueryName:  Me.ObjectName = DataSource
        Case dstSQLStament: Me.SQLStatement = DataSource
        Case dstRecordset:  Set Me.Recordset = DataSource
    End Select
  
End Sub


' ### Generator Function

Friend Function getData() As String
On Error GoTo Handle_Error

    'Declarations
    Dim db As DAO.Database
    Dim rst As Object           'ADO/DAO Recordset
    
    Set db = CurrentDb()
    
    'Detect data source
    Select Case m_Type
        Case dstTableName, dstQueryName:
            Set rst = db.OpenRecordset(ObjectName)
        Case dstSQLStament:
            Set rst = db.OpenRecordset(SQLStatement)
        Case dstRecordset:
            Set rst = Me.Recordset
        Case dstDemo:
            ' TODO
    End Select
    
    getData = getDataFromRS(rst)
    
Exit_Here:
    If Not (rst Is Nothing) Then
        rst.Close
        Set rst = Nothing
    End If
    Set db = Nothing
    Exit Function

Handle_Error:
    Select Case Err.Number
        Case Else
            MsgBox Err.Description, vbExclamation, Err.Number
    End Select
    Resume Exit_Here
    Resume Next

End Function

Private Function getDataFromRS(ByVal rst As Object) As String

    'Declarations
    Dim fld As Object           'ADODB/DAO.Field
    Dim strText As New BAPT_StringBuilder
    Dim strLine As String

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
    
    getDataFromRS = strText.ToString
    
End Function

