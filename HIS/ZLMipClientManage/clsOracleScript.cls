VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOracleScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum SQLTYPE
    enumNone = 0
    enumDefine = 1
    enumSelect = 2
    enumInsert
    enumUpdate
    enumDelete
    enumIndex
    enumConstraint
    enumTable
    enumSequence
    enumGrant
    enumBlock
    enumProcedure
    enumFunction
    enumPackage
    enumDrop
    enumAlter
    enumSynonym
    enumCommit
End Enum

Private mcolSQL As Collection
Private mobjFile As TextStream                          '数据流对象
Private mobjFso As New FileSystemObject                 '用来进行文件操作
Private mstrScriptFile As String
Private mstrLastError As String

Public Event BeforeAnalyseLine(ByVal Line As Long, ByVal Lines As Long)
Public Event AfterAnalyseLine(ByVal Line As Long, ByVal Lines As Long)

Public Function OpenScriptFile(ByVal strScriptFile As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '入参:
    '返回:
    '******************************************************************************************************************
    mstrScriptFile = strScriptFile
    OpenScriptFile = AnalyseScript
    
End Function

Public Function ExecuteSQL(ByRef cnOracle As ADODB.Connection, ByVal strSQL As String, Optional ByVal bytSQLType As SQLTYPE = enumSelect) As Boolean
    '******************************************************************************************************************
    '功能:
    '入参:
    '返回:
    '******************************************************************************************************************
    
'    On Error GoTo errHand
    
    mstrLastError = ""
        
    Select Case bytSQLType
    '--------------------------------------------------------------------------------------------------------------
    Case enumInsert, enumUpdate, enumDelete, enumGrant, enumIndex, enumConstraint, _
        enumTable, enumDrop, enumAlter, enumSequence, enumDefine, enumCommit, enumSynonym, enumFunction, enumProcedure, enumPackage
        
        Call cnOracle.Execute(strSQL)
        
    '--------------------------------------------------------------------------------------------------------------
    Case enumBlock
    
        
    Case Else
    
        Call cnOracle.Execute(strSQL)
        
    End Select
    
    Exit Function
'
'errHand:
'    mstrLastError = Err.Description
End Function

Public Property Get LastError() As String
    LastError = mstrLastError
End Property

Public Property Get SQLCount() As Long
    SQLCount = mcolSQL.Count
End Property

Public Property Get SQL(ByVal Index As Long) As String
    SQL = mcolSQL.Item(Index)
End Property

Private Function AnalyseScript() As Boolean
    '******************************************************************************************************************
    '功能:
    '入参:
    '返回:
    '******************************************************************************************************************
    Dim bytSQLType As SQLTYPE
    Dim strLine As String
    Dim strLine1 As String
    Dim strSQL As String
    Dim intCount As Integer
    Dim strLastLine As String
    Dim lngLine As Long
    Dim lngLines As Long
    
    On Error GoTo errHand
    
    Set mcolSQL = New Collection
    
    bytSQLType = enumNone
    Set mobjFile = mobjFso.OpenTextFile(mstrScriptFile, ForReading)
    If mobjFile.AtEndOfStream = False Then
        mobjFile.ReadAll
        lngLines = mobjFile.Line
    Else
        mobjFile.Close
        Exit Function
    End If
    mobjFile.Close
    
    Set mobjFile = mobjFso.OpenTextFile(mstrScriptFile, ForReading)
    Do While Not mobjFile.AtEndOfStream
        strLine = mobjFile.ReadLine
        
        lngLine = lngLine + 1
        RaiseEvent BeforeAnalyseLine(lngLine, lngLines)
        
        '去掉前后的特殊字符
        strLine1 = TrimSpecialChar(strLine)
        
        'SQL的开始标记
        '--------------------------------------------------------------------------------------------------------------
        If InStr(UCase(strLine1), UCase("Define")) = 1 Then
            
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumDefine
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Select")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumSelect
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Insert")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumInsert
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Delete")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumDelete
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Update")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumUpdate
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Declare")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumBlock
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Grant")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumGrant
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(Trim(strLastLine & " " & strLine1)), UCase("Create")) = 1 Then
            If UCase(strLine1) = UCase("Create") Then
                '单独
                strLastLine = strLine1
                strSQL = strLine
            '----------------------------------------------------------------------------------------------------------
            ElseIf InStr(UCase(Trim(strLastLine & " " & strLine1)), UCase("Create Table")) = 1 Then
                If bytSQLType = enumNone Then
                    strSQL = Trim(strLastLine & " " & strLine)
                    bytSQLType = enumTable
                Else
                    strSQL = strSQL & vbCrLf & strLine
                End If
                strLastLine = ""
            '----------------------------------------------------------------------------------------------------------
            ElseIf InStr(UCase(Trim(strLastLine & " " & strLine1)), UCase("Create Sequence")) = 1 Then
                If bytSQLType = enumNone Then
                    strSQL = Trim(strLastLine & " " & strLine)
                    bytSQLType = enumSequence
                Else
                    strSQL = strSQL & vbCrLf & strLine
                End If
                strLastLine = ""
            '----------------------------------------------------------------------------------------------------------
            ElseIf InStr(UCase(Trim(strLastLine & " " & strLine1)), UCase("Create Index")) = 1 Then
                If bytSQLType = enumNone Then
                    strSQL = Trim(strLastLine & " " & strLine)
                    bytSQLType = enumIndex
                Else
                    strSQL = strSQL & vbCrLf & strLine
                End If
                strLastLine = ""
                
            '----------------------------------------------------------------------------------------------------------
            ElseIf InStr(UCase(Trim(strLastLine & " " & strLine1)), UCase("Create Public Synonym")) = 1 Then
                If bytSQLType = enumNone Then
                    strSQL = Trim(strLastLine & " " & strLine)
                    bytSQLType = enumIndex
                Else
                    strSQL = strSQL & vbCrLf & strLine
                End If
                strLastLine = ""
                
            '----------------------------------------------------------------------------------------------------------
            ElseIf InStr(UCase(Trim(strLastLine & " " & strLine1)), UCase("Create Or Replace Procedure")) = 1 Then
                If bytSQLType = enumNone Then
                    strSQL = Trim(strLastLine & " " & strLine)
                    bytSQLType = enumProcedure
                Else
                    strSQL = strSQL & vbCrLf & strLine
                End If
                strLastLine = ""
            '----------------------------------------------------------------------------------------------------------
            ElseIf InStr(UCase(Trim(strLastLine & " " & strLine1)), UCase("Create Or Replace Function")) = 1 Then
                If bytSQLType = enumNone Then
                    strSQL = Trim(strLastLine & " " & strLine)
                    bytSQLType = enumFunction
                Else
                    strSQL = strSQL & vbCrLf & strLine
                End If
                strLastLine = ""
            '----------------------------------------------------------------------------------------------------------
            ElseIf InStr(UCase(Trim(strLastLine & " " & strLine1)), UCase("Create Or Replace Package")) = 1 Then
                If bytSQLType = enumNone Then
                    strSQL = Trim(strLastLine & " " & strLine)
                    bytSQLType = enumPackage
                Else
                    strSQL = strSQL & vbCrLf & strLine
                End If
                strLastLine = ""
            Else
                strLastLine = strLastLine & " " & strLine1
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Drop")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumDrop
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Alter")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumAlter
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        ElseIf InStr(UCase(strLine1), UCase("Commit")) = 1 Then
            If bytSQLType = enumNone Then
                strSQL = strLine
                bytSQLType = enumCommit
            Else
                strSQL = strSQL & vbCrLf & strLine
            End If
        '--------------------------------------------------------------------------------------------------------------
        Else
            If bytSQLType <> enumNone Then strSQL = strSQL & vbCrLf & strLine
        End If
        
        'SQL的结束标记
        '--------------------------------------------------------------------------------------------------------------
        Select Case bytSQLType
        '--------------------------------------------------------------------------------------------------------------
        Case enumSelect, enumInsert, enumUpdate, enumDelete, enumGrant, enumIndex, enumConstraint, _
            enumTable, enumDrop, enumAlter, enumSequence, enumDefine, enumCommit, enumSynonym
            
            If Right(strLine1, 1) = ";" Or strLine1 = "/" Then
                intCount = intCount + 1
                strSQL = TrimSpecialChar(strSQL)
                strSQL = TrimSpecialChar(Mid(strSQL, 1, Len(strSQL) - 1))
                mcolSQL.Add strSQL, "SQL_" & intCount
                strSQL = ""
                bytSQLType = enumNone
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case enumBlock, enumFunction, enumProcedure, enumPackage
        
            If strLine1 = "/" Then
                intCount = intCount + 1
                strSQL = TrimSpecialChar(strSQL)
                strSQL = TrimSpecialChar(Mid(strSQL, 1, Len(strSQL) - 1))
                
'                If Right(strSQL, 1) = ";" Then
'                    strSQL = TrimSpecialChar(Mid(strSQL, 1, Len(strSQL) - 1))
'                End If
                
                mcolSQL.Add strSQL, "SQL_" & intCount
                strSQL = ""
                bytSQLType = enumNone
            End If
            
        End Select
        
        RaiseEvent AfterAnalyseLine(lngLine, lngLines)
        
    Loop
    mobjFile.Close
    
    AnalyseScript = True
    
    Exit Function
errHand:
    
End Function

Private Function TrimSpecialChar(ByVal strText As String) As String
    '******************************************************************************************************************
    '功能:
    '入参:
    '返回:
    '******************************************************************************************************************
    Dim strNewText As String
    Dim lngLoop As Long
        
    '去除前后空格
    strNewText = Trim(strText)
    
    '判断前导字符
    For lngLoop = 1 To Len(strNewText)
        Select Case Asc(Mid(strNewText, lngLoop, 1))
        Case 9, 32, 13, 10       'tab,空格
        Case Else
            strNewText = Mid(strNewText, lngLoop)
            Exit For
        End Select
    Next
    
    '判断后导字符
    For lngLoop = Len(strNewText) To 1 Step -1
        Select Case Asc(Mid(strNewText, lngLoop, 1))
        Case 9, 32, 13, 10       'tab,空格
        Case Else
            strNewText = Mid(strNewText, 1, lngLoop)
            Exit For
        End Select
    Next
            
    '内容中间有连续多个空格的换算为一个空格
    Do While InStr(strNewText, "  ") > 0
        strNewText = Replace(strNewText, "  ", " ")
    Loop
    
    TrimSpecialChar = strNewText
End Function


Private Sub Class_Initialize()
    Set mobjFso = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set mobjFso = Nothing
    Set mobjFile = Nothing
    Set mcolSQL = Nothing
End Sub



