VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义
Private mstrScriptPath As String
Private mstrScriptIni As String
Private mstrOwner As String
Private mstrUserName As String
Private mobjFile As FileSystemObject
Private mobjText As TextStream
Private mstrSysName As String
Private mstrSQL As String
Private mblnDBAUser As Boolean
Private mfrmMain As Object
Private mintSysCode As Integer
Private mstrSystem As String

Private mcllBakErr As Collection
Private mblnHaveBak As Boolean
Private mstrDbLinkName As String
Private mstrBakOwner As String
Private mrsErrTable As ADODB.Recordset
Private mstrSplit As String
Private mrsSynonym As New ADODB.Recordset
Private mclsRunScript As New clsRunScript

'事件定义
Public Event BeforeProgress(ByVal Title As String, ByVal Max As Long)
Public Event Progressing(ByVal Progress As Long)
Public Event AfterProgress()
Public Event BeforeObjectCheck()
Public Event AfterObjectCheck()
Public Event Exception()

'######################################################################################################################
Public Function InitCheck(ByVal frmMain As Object, ByVal strOwner As String, ByVal strScriptIni As String, ByVal strUserName As String, ByVal strSysName As String, ByVal blnDBAUser As Boolean, ByVal intSysCode As Integer, ByVal strSystem As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set mfrmMain = frmMain
    mstrOwner = strOwner
    mstrScriptIni = strScriptIni
    mstrUserName = strUserName
    mstrSysName = strSysName
    mblnDBAUser = blnDBAUser
    mintSysCode = intSysCode
    mstrSystem = strSystem
    
    mstrScriptPath = Mid(mstrScriptIni, 1, Len(mstrScriptIni) - 11)
    
    frmAppChkRpt.hgdReport.Rows = 1
    frmAppChkRpt.lblWarn.Visible = False
            
    If CheckValid = False Then Exit Function
    
    
    InitCheck = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Public Function InitCheckManageTool(ByVal frmMain As Object, ByVal strScriptIni As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set mfrmMain = frmMain
    mstrScriptIni = strScriptIni
    
    frmAppChkRpt.hgdReport.Rows = 1
    frmAppChkRpt.lblWarn.Visible = False
    
    InitCheckManageTool = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Public Function CheckToolsObject() As Boolean
    '******************************************************************************************************************
    '功能：检查Tools对象是否正确
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim strSeverScrip As String 'Server脚本
    Dim cnTools As ADODB.Connection
    Dim cllErrTable As New Collection
    Dim rsConstraint As New ADODB.Recordset, cllErrConstraint As New Collection
    Dim rsIndex As New ADODB.Recordset, cllErrIndex As New Collection
    Dim rsSequence As New ADODB.Recordset, cllErrSequence As New Collection
    Dim rsTypes As New ADODB.Recordset, cllErrTypes As New Collection
    Dim rsProgram As New ADODB.Recordset, cllErrProgram As New Collection
    Dim rsUserObject As New ADODB.Recordset
    Dim lngCount As Long
    Dim strSQL As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim strTemp As String
    Dim strType As String
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    
    RaiseEvent BeforeObjectCheck
    
    strSeverScrip = mstrScriptIni
                
    If Not mclsRunScript.OpenFile(strSeverScrip) Or mclsRunScript.EOF Then
        RaiseEvent Exception
        Exit Function
    End If

    '------------------------------------------------------------------------------------------------------------------
    Set cnTools = GetConnection("ZLTOOLS")
    If cnTools Is Nothing Then
        MsgBox "打开管理工具失败,请检查!", vbInformation, gstrSysName
        Exit Function
    End If
    On Error GoTo errHand
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from USER_CONSTRAINTS"
    Call OpenRecordset(rsConstraint, strSQL, mfrmMain.Caption, , , cnTools)
    strSQL = "select INDEX_NAME,STATUS from USER_INDEXES where TABLE_OWNER='ZLTOOLS'"
    Call OpenRecordset(rsIndex, strSQL, mfrmMain.Caption, , , cnTools)
    
    strSQL = "select SEQUENCE_NAME from USER_SEQUENCES"
    Call OpenRecordset(rsSequence, strSQL, mfrmMain.Caption, , , cnTools)

    strSQL = "Select object_Name,status From USER_OBJECTS Where object_type='TYPE'"
    Call OpenRecordset(rsTypes, strSQL, mfrmMain.Caption, , , cnTools)
    
    strSQL = "select OBJECT_TYPE,OBJECT_NAME,STATUS from USER_OBJECTS where OBJECT_TYPE in('PROCEDURE','FUNCTION','PACKAGE')"
    Call OpenRecordset(rsProgram, strSQL, mfrmMain.Caption, , , cnTools)
    
    '检查公共同义词
    Set mrsSynonym = New ADODB.Recordset
    gstrSQL = " Select table_name From All_Synonyms Where Table_Owner = 'ZLTOOLS' And OWNER='PUBLIC'"
    Call OpenRecordset(mrsSynonym, gstrSQL, mfrmMain.Caption, , , cnTools)
    
    
    
    Dim objCollectionDB As New Collection
    Dim rsColumns As New ADODB.Recordset
    '列数据
    '------------------------------------------------------------------------------------------------------------------

    strTemp = "SELECT TABLE_NAME,COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT From User_TAB_COLUMNS Where TABLE_NAME NOT LIKE 'BIN$%'"
    Call OpenRecordset(rsColumns, strTemp, mfrmMain.Caption, , , cnTools)
    If rsColumns.BOF = False Then
        Do While Not rsColumns.EOF
            strTemp = rsColumns("TABLE_NAME").value & "'" & rsColumns("COLUMN_NAME").value & "'" & rsColumns("DATA_TYPE").value & "'" & rsColumns("DATA_LENGTH").value & "'" & rsColumns("DATA_PRECISION").value & "'" & rsColumns("DATA_SCALE").value & "'" & rsColumns("DATA_DEFAULT").value
            objCollectionDB.Add strTemp, "K" & rsColumns("TABLE_NAME").value & "'" & rsColumns("COLUMN_NAME").value
            rsColumns.MoveNext
        Loop
    End If
    
    
    Dim objClsTable As New clsCollectionRecordset
    '------------------------------------------------------------------------------------------------------------------
    
    strTemp = "Select TABLE_NAME From USER_TABLES Where  TABLE_NAME NOT LIKE 'BIN$%'"
    Call OpenRecordset(rsColumns, strTemp, mfrmMain.Caption, , , cnTools)
    Call objClsTable.LoadData(rsColumns, Array(), Array("TABLE_NAME"))
    
    
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    RaiseEvent BeforeProgress("正在检查服务器工具对象", 100)
    
    Do While Not mclsRunScript.EOF
        strSQL = UCase(TrimEx(mclsRunScript.SQLInfo.SQL))
        strTemp = mclsRunScript.SQLInfo.PartSQL
'        If mclsRunScript.Line >= 805 Then Stop
        If Not mclsRunScript.SQLInfo.Block Then
            If strTemp Like "CREATE TABLE*" Or strTemp Like "CREATE GLOBAL TEMPORARY TABLE*" Then
                '检查表
                Call CheckToolsTables(cnTools, objClsTable, objCollectionDB, strSQL, cllErrTable)
            ElseIf strTemp Like "ALTER TABLE*" Then
                If InStr(1, strTemp, "CONSTRAINT") > 0 Then
                    Call CheckToolsConstraint(cnTools, rsConstraint, strSQL, cllErrConstraint)
                End If
            ElseIf strTemp Like "CREATE INDEX*" Then
                '检查索引
                Call CheckToolsIndex(cnTools, rsIndex, strSQL, cllErrIndex)
            ElseIf strTemp Like "CREATE SEQUENCE*" Then
                '检查序列
                Call CheckToolsSequeces(cnTools, rsSequence, strSQL, cllErrSequence)
            End If
        Else
            strType = mclsRunScript.SQLInfo.BlockType
            If strType = "TYPE" Then
                '检查自定义类型
                Call CheckToolsTypes(cnTools, rsTypes, strSQL, cllErrTypes)
            ElseIf strType Like "*PROCEDURE*" Or _
                strType Like "*FUNCTION*" Or _
                (strType Like "*PACKAGE*" And Not strType Like "*PACKAGE BODY*") Then
                '检查过程与函数的合法性
                Call CheckToolsPrograms(cnTools, rsProgram, strSQL, cllErrProgram)
            End If
        End If
        intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
        If intSvrPercent <> intPercent Then
            intSvrPercent = intPercent
            RaiseEvent Progressing(intSvrPercent)
        End If
        Call mclsRunScript.ReadNextSQL
    Loop

    '------------------------------------------------------------------------------------------------------------------
    
    '填充相关的错语数据
    For lngCount = 1 To cllErrTable.Count
        Call InputErrRpt(CStr(cllErrTable(lngCount)(0)), CStr(cllErrTable(lngCount)(1)), CStr(cllErrTable(lngCount)(2)), CStr(cllErrTable(lngCount)(3)))
    Next
    Set cllErrTable = Nothing
    
    For lngCount = 1 To cllErrConstraint.Count
        Call InputErrRpt(CStr(cllErrConstraint(lngCount)(0)), CStr(cllErrConstraint(lngCount)(1)), CStr(cllErrConstraint(lngCount)(2)), CStr(cllErrConstraint(lngCount)(3)))
    Next
    For lngCount = 1 To cllErrIndex.Count
        Call InputErrRpt(CStr(cllErrIndex(lngCount)(0)), CStr(cllErrIndex(lngCount)(1)), CStr(cllErrIndex(lngCount)(2)), CStr(cllErrIndex(lngCount)(3)))
    Next
    For lngCount = 1 To cllErrSequence.Count
        Call InputErrRpt(CStr(cllErrSequence(lngCount)(0)), CStr(cllErrSequence(lngCount)(1)), CStr(cllErrSequence(lngCount)(2)), CStr(cllErrSequence(lngCount)(3)))
    Next
    For lngCount = 1 To cllErrTypes.Count
        Call InputErrRpt(CStr(cllErrTypes(lngCount)(0)), CStr(cllErrTypes(lngCount)(1)), CStr(cllErrTypes(lngCount)(2)), CStr(cllErrTypes(lngCount)(3)))
    Next
    For lngCount = 1 To cllErrProgram.Count
        Call InputErrRpt(CStr(cllErrProgram(lngCount)(0)), CStr(cllErrProgram(lngCount)(1)), CStr(cllErrProgram(lngCount)(2)), CStr(cllErrProgram(lngCount)(3)))
    Next
 
    If rsConstraint.State = 1 Then rsConstraint.Close
    If rsIndex.State = 1 Then rsIndex.Close
    If rsSequence.State = 1 Then rsSequence.Close
    If rsTypes.State = 1 Then rsTypes.Close
    If rsProgram.State = 1 Then rsProgram.Close
    
    '检查公共同义词
    If mrsSynonym.State = 1 Then
        mrsSynonym.Filter = 0
        mrsSynonym.Close
    End If
    
    RaiseEvent AfterProgress
    
    
    RaiseEvent AfterObjectCheck
    
    CheckToolsObject = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Function CheckToolsTypes(ByVal cnTools As ADODB.Connection, rsObject As ADODB.Recordset, strSQL As String, ByRef cllTemp As Collection) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查Tools工具中指定的自定义类型
    '参数:cnTools-连接到服务器管理工具中的连接
    '     rsObject-用户对象
    '     strSQL-创建表的SQL
    '出参:返回错误信息
    '返回:合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strObjName As String, arySql As Variant, strTableName As String, strColumns As String
    Dim rsColumns As New ADODB.Recordset
    
    strSQL = UCase(Replace(Replace(Trim(strSQL), vbTab, " "), vbCrLf, " "))
    arySql = Split(strSQL, " TYPE ")
    strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
    strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
    With rsObject
        .Filter = " object_Name='" & strObjName & "'"
        If .EOF Then
           cllTemp.Add Array("自定义类型", strObjName, "不存在", "严重：部分功能不能正常运行")
        ElseIf Nvl(!Status) <> "VALID" Then
           cllTemp.Add Array("自定义类型", strObjName, "失效", "严重：部分功能不能正常运行")
        End If
        If .EOF = False Then
            '检查同义试词是否存在
             mrsSynonym.Filter = "table_name='" & strObjName & "'"
             If mrsSynonym.EOF Then
                 cllTemp.Add Array("序列", strObjName, "无此公共同义词", "较重：可能导致应用系统运行不正常！")
             End If
        End If
    End With
    CheckToolsTypes = True
End Function

Private Function CheckToolsPrograms(ByVal cnTools As ADODB.Connection, rsObject As ADODB.Recordset, strSQL As String, ByRef cllTemp As Collection) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查Tools工具中指定表的过程和函数以及包
    '参数:cnTools-连接到服务器管理工具中的连接
    '     rsObject-用户对象
    '     strSQL-创建表的SQL
    '出参:返回错误信息
    '返回:合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strObjName As String, arySql As Variant, strTableName As String, strColumns As String
    Dim rsColumns As New ADODB.Recordset
    
    strSQL = UCase(Replace(Replace(Trim(strSQL), vbTab, " "), vbCrLf, " "))
    With rsObject
        If InStr(1, strSQL, UCase(" PACKAGE ")) > 0 And InStr(1, strSQL, "CREATE ") > 0 Then
            arySql = Split(strSQL, " PACKAGE ")
            strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
            If InStr(strSQL, " ") <> 0 Then
                strObjName = Trim(Left(strSQL, InStr(1, strSQL, " ") - 1))
            End If
            
            If InStr(1, strObjName, "(") <> 0 Then
                strObjName = Trim(Left(strObjName, InStr(1, strObjName, "(") - 1))
            End If
            .Filter = "OBJECT_NAME='" & strObjName & "'"
            If .EOF Then
               cllTemp.Add Array("过程包", strObjName, "不存在", "严重：部分功能不能正常运行")
            ElseIf .Fields("STATUS").value <> "VALID" Then
               cllTemp.Add Array("过程包", strObjName, "处于无效状态", "严重：部分功能不能正常运行")
            End If
            If .EOF = False Then
                '检查同义试词是否存在
                 mrsSynonym.Filter = "table_name='" & strObjName & "'"
                 If mrsSynonym.EOF Then
                     cllTemp.Add Array("序列", strObjName, "无此公共同义词", "较重：可能导致应用系统运行不正常！")
                 End If
            End If
        ElseIf InStr(1, strSQL, " PROCEDURE ") > 0 And InStr(1, strSQL, "CREATE ") > 0 Then
            arySql = Split(strSQL, " PROCEDURE ")
            strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
            If InStr(1, strSQL, " ") <> 0 Then
                strObjName = Trim(Left(strSQL, InStr(1, strSQL, " ") - 1))
            End If
            If InStr(1, strObjName, "(") <> 0 Then
                strObjName = Trim(Left(strObjName, InStr(1, strObjName, "(") - 1))
            End If
            .Filter = "OBJECT_NAME='" & strObjName & "'"
            If .EOF Then
               cllTemp.Add Array("过程/函数", strObjName, "不存在", "严重：部分功能不能正常运行")
            ElseIf .Fields("STATUS").value <> "VALID" Then
               cllTemp.Add Array("过程/函数", strObjName, "处于无效状态", "严重：部分功能不能正常运行")
            End If
            If .EOF = False Then
                '检查同义试词是否存在
                 mrsSynonym.Filter = "table_name='" & strObjName & "'"
                 If mrsSynonym.EOF Then
                     cllTemp.Add Array("序列", strObjName, "无此公共同义词", "较重：可能导致应用系统运行不正常！")
                 End If
            End If
        ElseIf InStr(1, strSQL, " FUNCTION ") > 0 And InStr(1, strSQL, "CREATE ") > 0 Then
            arySql = Split(strSQL, " FUNCTION ")
            strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
            If InStr(strSQL, " ") <> 0 Then
                strObjName = Trim(Left(strSQL, InStr(1, strSQL, " ") - 1))
            End If
            If InStr(1, strObjName, "(") <> 0 Then
                strObjName = Trim(Left(strObjName, InStr(1, strObjName, "(") - 1))
            End If
            .Filter = "OBJECT_NAME='" & strObjName & "'"
            If .EOF Then
               cllTemp.Add Array("过程/函数", strObjName, "不存在", "严重：部分功能不能正常运行")
            ElseIf .Fields("STATUS").value <> "VALID" Then
               cllTemp.Add Array("过程/函数", strObjName, "处于无效状态", "严重：部分功能不能正常运行")
            End If
            If .EOF = False Then
                '检查同义试词是否存在
                 mrsSynonym.Filter = "table_name='" & strObjName & "'"
                 If mrsSynonym.EOF Then
                     cllTemp.Add Array("序列", strObjName, "无此公共同义词", "较重：可能导致应用系统运行不正常！")
                 End If
            End If
        End If
    End With
End Function

Private Function CheckToolsSequeces(ByVal cnTools As ADODB.Connection, rsObject As ADODB.Recordset, strSQL As String, ByRef cllTemp As Collection) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查Tools工具中指定表的相关序列
    '参数:cnTools-连接到服务器管理工具中的连接
    '     rsObject-用户对象
    '     strSQL-创建表的SQL
    '出参:返回错误信息
    '返回:合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strObjName As String, arySql As Variant, strTableName As String, strColumns As String
    Dim rsColumns As New ADODB.Recordset
    
    strSQL = UCase(Replace(Replace(Trim(strSQL), vbTab, " "), vbCrLf, " "))
    arySql = Split(strSQL, " SEQUENCE ")
    strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
    strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
    With rsObject
        .Filter = "SEQUENCE_NAME='" & strObjName & "'"
        If .EOF Then
           cllTemp.Add Array("序列", strObjName, "不存在", "严重：部分功能不能正常运行")
        Else
               '检查同义试词是否存在
                mrsSynonym.Filter = "table_name='" & strObjName & "'"
                If mrsSynonym.EOF Then
                    cllTemp.Add Array("序列", strObjName, "无此公共同义词", "较重：可能导致应用系统运行不正常！")
                End If
        End If
        
    End With
    CheckToolsSequeces = True
End Function


Private Function CheckToolsIndex(ByVal cnTools As ADODB.Connection, rsObject As ADODB.Recordset, strSQL As String, ByRef cllTemp As Collection) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查Tools工具中指定表的相关索引
    '参数:cnTools-连接到服务器管理工具中的连接
    '     rsObject-用户对象
    '     strSQL-创建表的SQL
    '出参:返回错误信息
    '返回:合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strObjName As String, arySql As Variant, strColumns As String
    Dim rsColumns As New ADODB.Recordset, strTablenName As String
    Dim strTemp As String
    
    strSQL = UCase(Replace(Replace(Trim(strSQL), vbTab, " "), vbCrLf, " "))
    arySql = Split(strSQL, " INDEX ")
    If UBound(arySql) > 0 Then
        strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
        strTablenName = Trim(Split(Split(strSQL, "ON")(1), "(")(0))
        strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
        With rsObject
            .Filter = "INDEX_NAME='" & strObjName & "'"
            If .EOF Then
               cllTemp.Add Array("索引", strObjName, "不存在", "较重：可能影响系统运行速度")
            ElseIf .Fields("STATUS").value <> "VALID" Then
               cllTemp.Add Array("索引", strObjName, "当前处于无效状态")
            Else
                strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                        " from USER_IND_COLUMNS" & _
                        " where INDEX_NAME='" & strObjName & "'" & _
                        " order by COLUMN_POSITION"
                Call OpenRecordset(rsColumns, strTemp, mfrmMain.Caption, , , cnTools)
                With rsColumns
                    Do While Not .EOF
                        If .AbsolutePosition = 1 Then
                            strColumns = !Table_Name & "(" & !Column_Name
                        Else
                            strColumns = strColumns & "," & !Column_Name
                        End If
                        .MoveNext
                    Loop
                        strColumns = strColumns & ")"
                End With
                arySql = Split(strSQL, " ON ")
                strTemp = Replace(Left(arySql(1), InStr(1, arySql(1), ")")), " ", "")
                If strColumns <> strTemp Then
                   cllTemp.Add Array("索引", strObjName, "索引列错误，应为“" & strTemp & "”，现为“" & strColumns & "”", "较重：可能影响系统运行速度")
                End If
            End If
        End With
    End If
    CheckToolsIndex = True
End Function

Private Function CheckToolsConstraint(ByVal cnTools As ADODB.Connection, rsObject As ADODB.Recordset, strSQL As String, ByRef cllTemp As Collection) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查Tools工具中指定表的相关约束
    '参数:cnTools-连接到服务器管理工具中的连接
    '     rsObject-用户对象
    '     strSQL-创建表的SQL
    '出参:返回错误信息
    '返回:合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strObjName As String, arySql As Variant, strTableName As String, strColumns As String
    Dim rsColumns As New ADODB.Recordset
    Dim strTemp As String
    
    strSQL = UCase(Replace(Replace(Trim(strSQL), vbTab, " "), vbCrLf, " "))
    arySql = Split(strSQL, " CONSTRAINT ")
    If UBound(arySql) > 0 Then
        '刘兴宏加入:
        strTableName = Trim(Split(Trim(arySql(0)), "TABLE")(1))
        strTableName = Split(strTableName, " ")(0)
        
        strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
        strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
        
        With rsObject
                .Filter = "CONSTRAINT_NAME='" & strObjName & "'"
                If .EOF Then
                    If InStr(1, strSQL, " CHECK") > 0 Then
                        cllTemp.Add Array("约束", strObjName, "不存在", "较轻：基本不影响系统运行")
                    Else
                        cllTemp.Add Array("约束", strObjName, "不存在", "较重：可能导致数据不一致，影响运行速度")
                    End If
                ElseIf .Fields("STATUS").value <> "ENABLED" Then
                   cllTemp.Add Array("约束", strObjName, "当前处于禁止状态", "较重：可能系统已经存在问题")
                ElseIf !VALIDATED <> "VALIDATED" Then
                   cllTemp.Add Array("约束", strObjName, "当前处于无效状态", "较重：可能数据一致性已被破坏")
                ElseIf Not IsNull(!BAD) Then
                   cllTemp.Add Array("约束", strObjName, "约束被意外损坏", "严重：可能存在硬件错误")
                Else
                    strTemp = "" & _
                    " Select COLUMN_NAME" & _
                    " From USER_CONS_COLUMNS" & _
                    " Where CONSTRAINT_NAME='" & strObjName & "'" & _
                    " order by POSITION"
                    OpenRecordset rsColumns, strTemp, mfrmMain.Caption, , , cnTools
                    
                    strColumns = ""
                    With rsColumns
                        Do While Not .EOF
                            strColumns = strColumns & "," & !Column_Name
                            .MoveNext
                        Loop
                    End With
                    If InStr(1, strSQL, " PRIMARY ") > 0 Then
                        If rsObject!constraint_type <> "P" Then
                           cllTemp.Add Array("约束", strObjName, "约束类型错误，应为主键约束", "严重：可能影响系统运行")
                        Else
                            arySql = Split(strSQL, " PRIMARY ")
                            strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                               cllTemp.Add Array("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "较重：可能影响系统运行")
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " UNIQUE") > 0 Then
                        If !constraint_type <> "U" Then
                           cllTemp.Add Array("约束", strObjName, "约束类型错误，应为唯一约束", "较重：可能影响系统运行")
                        Else
                            arySql = Split(strSQL, " UNIQUE ")
                            If UBound(arySql) = 0 Then arySql = Split(strSQL, " UNIQUE(")
                            strTemp = Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                               cllTemp.Add Array("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "较重：可能影响系统运行")
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " FOREIGN ") > 0 Then
                        If !constraint_type <> "R" Then
                           cllTemp.Add Array("约束", strObjName, "约束类型错误，应为外键约束", "严重：可能影响系统运行")
                        Else
                            arySql = Split(strSQL, " FOREIGN ")
                            strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                            If strColumns <> "," & strTemp Then
                               cllTemp.Add Array("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "严重：可能影响数据一致性")
                            End If
                        End If
                    ElseIf InStr(1, strSQL, " CHECK") > 0 Then
                        If !constraint_type <> "C" Then
                           cllTemp.Add Array("约束", strObjName, "约束类型错误，应为检查约束", "严重：可能影响系统运行")
                        End If
                    End If
                End If
        End With
    End If
    CheckToolsConstraint = True
End Function

Private Function CheckToolsTables(ByVal cnTools As ADODB.Connection, ByVal objClsTable As clsCollectionRecordset, ByVal objCollectionDB As Collection, strSQL As String, ByRef cllTemp As Collection) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查Tools工具中指定表的结构是否正确
    '参数:cnTools-连接到服务器管理工具中的连接
    '     rsObject-用户对象
    '     strSQL-创建表的SQL
    '出参:cllTemp:返回错误信息
    '返回:合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strObjName As String, arySql As Variant
    Dim rsColumns As New ADODB.Recordset
    Dim strTemp As String
    Dim objCollection As New Collection
    Dim intLoop As Integer
    Dim strScriptFieldInfo As String
    Dim varScriptFieldInfo As Variant
    Dim strDBFieldInfo As String
    Dim varDBFieldInfo As Variant
    
    On Error GoTo errHand
    
    arySql = Split(strSQL, " TABLE ")
    strSQL = Trim(arySql(1)) '已经去掉Oracle关键字
    If InStr(1, strSQL, " ") > 0 And InStr(strSQL, " ") < InStr(strSQL, "(") Then
        strObjName = Trim(Left(strSQL, InStr(strSQL, " ")))
    Else
        strObjName = Trim(Left(strSQL, InStr(strSQL, "(") - 1))
    End If
    strObjName = Replace(strObjName, vbCrLf, "")
    
    If UCase(strObjName) = UCase("zlregfile") Then
        strObjName = UCase("zlregfile")
    End If
    '去掉所有者
    If InStr(strObjName, ".") > 0 Then
        strObjName = Mid(strObjName, InStr(strObjName, ".") + 1)
    End If

    If objClsTable.Item(Array(strObjName)) = False Then
    
        cllTemp.Add Array("数据表", strObjName, "不存在", "严重：部分功能不能正常运行")
    Else
    
            'new:从脚本中分析出字段、类型、长度
        '--------------------------------------------------------------------------------------------------
        Set objCollection = New Collection
        If AnsyTableScript(strObjName, strSQL, objCollection) = True Then
            '开始比较
            For intLoop = 1 To objCollection.Count
                strScriptFieldInfo = objCollection.Item(intLoop)
                varScriptFieldInfo = Split(strScriptFieldInfo, "'")

                On Error Resume Next
                strDBFieldInfo = ""
                strDBFieldInfo = objCollectionDB.Item("K" & strObjName & "'" & varScriptFieldInfo(0))
                varDBFieldInfo = Split(strDBFieldInfo, "'")
                On Error GoTo errHand

                '按脚本中的字段数和当前数据库中的字段进行比较，找出差字段、字段类型、字段长度等不一样的情况
                If CompareFieldTool(strObjName, varDBFieldInfo, varScriptFieldInfo, cllTemp) = True Then

                End If
                
                '检查同义试词是否存在
                mrsSynonym.Filter = "table_name='" & strObjName & "'"
    
                If mrsSynonym.EOF Then
                    cllTemp.Add Array("数据表", strObjName, "无此公共同义词", "较重：可能导致应用系统运行不正常！")
                End If

            Next
        End If
    End If
    CheckToolsTables = True
    
    Exit Function
errHand:
    MsgBox err.Description
    
End Function


Public Function CheckObject() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim objCollectionConstraintInTab As New Collection
    Dim objClsBakTable As New clsCollectionRecordset
    Dim rsTemp As New Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    RaiseEvent BeforeObjectCheck
    
    Set mcllBakErr = New Collection
    
    Call CheckHistoryLinkView
    Call CheckMustData(gcnOracle, mstrScriptPath, mintSysCode, objClsBakTable)
    Call CheckSequence(mstrScriptPath & "zlSequence.sql")
    Call CheckTable(mstrScriptPath & "zlTable.sql", objCollectionConstraintInTab, objClsBakTable)
    Call CheckConstraint(mstrScriptPath & "zlConstraint.sql", objCollectionConstraintInTab)
    Call CheckIndex(mstrScriptPath & "zlIndex.sql")
    
'    Call CheckConstraintIndex(mstrScriptPath & "zlConstraint.sql", objCollectionConstraintInTab)
    
    Call CheckView(mstrScriptPath & "zlView.sql")
    Call CheckProgram(mstrScriptPath & "zlProgram.sql")
    Call CheckPackage(mstrScriptPath & "zlPackage.sql")
    

    '加载错误信息:刘兴宏
    '出错,给出错信息赋值,以便显示
    If mcllBakErr.Count > 0 Then
        For intLoop = 1 To mcllBakErr.Count
            Call InputErrRpt("历史数据空间-" & CStr(mcllBakErr(intLoop)(0)), CStr(mcllBakErr(intLoop)(1)), CStr(mcllBakErr(intLoop)(2)), CStr(mcllBakErr(intLoop)(3)))
        Next
    End If
    
    RaiseEvent AfterObjectCheck
            
    CheckObject = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Public Sub ShowReport()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：无
    '******************************************************************************************************************
    If frmAppChkRpt.hgdReport.Rows > 1 Then
        If MsgBox("检查过程发现部分对象存在问题" & vbCr & "查看检查报告吗？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            frmAppChkRpt.hgdReport.FixedRows = 1
            frmAppChkRpt.Show 1
        End If
    Else
        MsgBox "对象检查完毕，未发现与安装文件不符的情况", vbInformation, gstrSysName
    End If

    Unload frmAppChkRpt
    
End Sub

'######################################################################################################################
Private Function CheckValid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If mstrUserName <> mstrOwner Then
        If MsgBox("你不是所有者登陆，可能历史数据空间不能正常检查，是否继续?", vbQuestion + vbYesNo + vbDefaultButton1, mstrSysName) = vbNo Then Exit Function
    End If

    If CheckIniFile(mstrScriptIni, True) = False Then
        RaiseEvent Exception
        Exit Function
    End If
    
    frmAppChkRpt.hgdReport.Rows = 1
    frmAppChkRpt.lblWarn.Visible = False

    CheckValid = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Function CheckIniFile(ByVal strScript As String, ByVal blnMsg As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim strTemp As String
    
    Dim intDefSysCode As Integer
    Dim strDefSysName As String
    Dim strDefVersion As String
    Dim strDefSpace As String
    
    On Error GoTo errHand
        
    err = 0
    On Error Resume Next
    
'    strIniPath = Mid(strScript, 1, Len(strScript) - 11)
    '相关文件匹配性检查
    strTemp = ""
    If Dir(mstrScriptPath & "zlSequence.sql") = "" Then strTemp = strTemp & vbCr & "序列文件" & mstrScriptPath & "zlSequence.sql"
    If Dir(mstrScriptPath & "zlTable.sql") = "" Then strTemp = strTemp & vbCr & "数据表文件" & mstrScriptPath & "zlTable.sql"
    If Dir(mstrScriptPath & "zlConstraint.sql") = "" Then strTemp = strTemp & vbCr & "约束文件" & mstrScriptPath & "zlConstraint.sql"
    If Dir(mstrScriptPath & "zlIndex.sql") = "" Then strTemp = strTemp & vbCr & "索引文件" & mstrScriptPath & "zlIndex.sql"
    If Dir(mstrScriptPath & "zlView.sql") = "" Then strTemp = strTemp & vbCr & "视图文件" & mstrScriptPath & "zlView.sql"
    If Dir(mstrScriptPath & "zlProgram.sql") = "" Then strTemp = strTemp & vbCr & "函数过程文件" & mstrScriptPath & "zlProgram.sql"
    
    '不检查,因为9系统没有此文件
    'If Dir(mstrScriptPath & "zlPackage.sql") = "" Then strTemp = strTemp & vbCr & "包文件" & mstrScriptPath & "zlPackage.sql"
    
    If Dir(mstrScriptPath & "zlManData.sql") = "" Then strTemp = strTemp & vbCr & "管理数据文件" & mstrScriptPath & "zlManData.sql"
    If Dir(mstrScriptPath & "zlAppData.sql") = "" Then strTemp = strTemp & vbCr & "应用数据文件" & mstrScriptPath & "zlAppData.sql"
    If strTemp <> "" Then
        If blnMsg Then MsgBox "以下服务器安装的相关文件丢失，不能继续，包括：" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '配置文件正确性检查
    Set mobjText = mobjFile.OpenTextFile(strScript)
    strTemp = Trim(mobjText.ReadLine)
    If Left(strTemp, 5) = "[系统号]" Then
        intDefSysCode = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(mobjText.ReadLine)
    If Left(strTemp, 5) = "[系统名]" Then
        strDefSysName = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(mobjText.ReadLine)
    If Left(strTemp, 5) = "[版本号]" Then
        strDefVersion = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(mobjText.ReadLine)
    If Left(strTemp, 5) = "[表空间]" Then
        strDefSpace = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    mobjText.Close
    
    If err <> 0 Then
        CheckIniFile = False
        If blnMsg Then MsgBox "安装配置文件不正确", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '配置文件符合性检查
    If intDefSysCode <> mintSysCode \ 100 Then
        err.Raise 10
        If blnMsg Then MsgBox "选择文件不是该系统的安装配置文件", vbExclamation, gstrSysName
    ElseIf InStr(1, mstrSystem, Trim(strDefVersion)) = 0 Then
        err.Raise 10
        If blnMsg Then MsgBox "选择文件与该系统版本不符", vbExclamation, gstrSysName
    End If
    
    If err = 0 Then
        CheckIniFile = True
    Else
        CheckIniFile = False
    End If

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function


Private Function CheckHistoryLinkView() As Boolean
    '******************************************************************************************************************
    '功能：检查历史数据空间
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    On Error GoTo errHand

    '分两个检查：
    '   1.zlBakTable中各个转储表对应视图的有效性(也就是检查了当前历史空间DB_LINK的有效性)。
    '   2.当前历史空间中结构的正确性(为了提高检查性能，此步分布在相关步骤同时进行(主要是索引、表、约束）步骤中进行)
    Dim rsTemp As New ADODB.Recordset
    Dim lngSys As Long, i As Long
    
    On Error GoTo errHand
    
    Set mcllBakErr = New Collection

    lngSys = mintSysCode
    
    mblnHaveBak = False
    If mstrUserName <> mstrOwner Then Exit Function
    
    gstrSQL = "Select 1 from zlbakTables where 系统=" & lngSys
    Call OpenRecordset(rsTemp, gstrSQL, "检查历史数据空间")
    If rsTemp.EOF Then
        '不存在历史数据空间，不检查
        Exit Function
    End If
    
    gstrSQL = "Select 系统, 编号, 名称, 所有者, db连接, 当前, 只读 From zltools.zlbakspaces where 当前=1 and 系统=" & lngSys
    Call OpenRecordset(rsTemp, gstrSQL, "检查历史数据空间")
    
    If rsTemp.EOF Then
        '不存在,则不必检查.
        Call InputErrRpt("分布数据", "历史数据空间", "不存在", "严重：历史数据转出或查询不能正常运行")
        mblnHaveBak = False
        Exit Function
    End If
    
    mstrDbLinkName = Nvl(rsTemp!db连接)
    mstrBakOwner = UCase(Nvl(rsTemp!所有者))
    If mstrBakOwner = "" Then '
        Call InputErrRpt("历史数据空间", mstrDbLinkName, "无历史数据空间所有者", "严重：历史数据空间将不能正常运行")
    End If
    
    If mstrDbLinkName <> "" Then
        err = 0
        On Error Resume Next
        gstrSQL = "Select 1 from dual@" & mstrDbLinkName
        Call OpenRecordset(rsTemp, gstrSQL, "连接验证", , , gcnOracle)
        If err <> 0 Then
            Call InputErrRpt("历史数据空间-远程连接", mstrDbLinkName, "连接不正常", "严重：历史数据空间将不能正常运行")
        End If
        err = 0
        On Error GoTo 0
    End If
    
    mblnHaveBak = True

    CheckHistoryLinkView = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function CheckSequence(ByVal strFile As String) As Boolean
    '******************************************************************************************************************
    '功能：检查序列对象，判断是否存在
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim rsObjects As New ADODB.Recordset
    Dim lngCount As Long
    Dim strTemp As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim objClsSequences As New clsCollectionRecordset
    Dim arySql() As String
    Dim strObjName As String
    
    On Error GoTo errHand
    
    RaiseEvent BeforeProgress("正在检查序列", 100)
    
    '--------------------------------------------------------------------------------------------------------------
    If Not mclsRunScript.OpenFile(strFile) Or mclsRunScript.EOF Then Exit Function
    '--------------------------------------------------------------------------------------------------------------
    
    If mblnDBAUser Then
        mstrSQL = "select SEQUENCE_NAME from DBA_SEQUENCES where SEQUENCE_OWNER='" & mstrOwner & "'"
    Else
        mstrSQL = "select SEQUENCE_NAME from USER_SEQUENCES"
    End If
    If rsObjects.State = adStateOpen Then rsObjects.Close
    rsObjects.Open mstrSQL, gcnOracle, adOpenKeyset
    Call objClsSequences.LoadData(rsObjects, Array(), Array("SEQUENCE_NAME"))
    
    '--------------------------------------------------------------------------------------------------------------
    Do While Not mclsRunScript.EOF
        mstrSQL = mclsRunScript.SQLInfo.PartSQL
        arySql = Split(mstrSQL, " SEQUENCE ")
        mstrSQL = Trim(arySql(1)) '已经去掉Oracle关键字
        strObjName = Trim(Left(mstrSQL, InStr(mstrSQL, " ")))
        
        If objClsSequences.Item(Array(strObjName)) = False Then
            Call InputErrRpt("序列", strObjName, "不存在", "严重：部分功能不能正常运行")
        End If
        
        intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
        If intSvrPercent <> intPercent Then
            intSvrPercent = intPercent
            RaiseEvent Progressing(intSvrPercent)
        End If
        Call mclsRunScript.ReadNextSQL
    Loop

    RaiseEvent AfterProgress
    
    CheckSequence = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Function GetFileLines(ByVal strFile As String) As Long

    err = 0
    On Error Resume Next
    Set mobjText = mobjFile.OpenTextFile(strFile)
    If err.Number <> 0 Then Exit Function
    
    Do While Not mobjText.AtEndOfStream
        mobjText.ReadLine
    Loop
    GetFileLines = mobjText.Line
        
End Function

Private Function CheckTable(ByVal strFile As String, ByRef objCollectionConstraintInTab As Collection, ByVal objClsBakTable As clsCollectionRecordset, Optional ByVal blnAddToRsTable As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能:检查数据表，同时判断数据表的列是否正确
    '入参:FileName-脚本文件名(包含具体路径)
    '     blnAddToRsTable-是否将相关的错误信息写在记录集中
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim strSQL1 As String
    Dim arySql() As String, strObjName As String
    Dim intVer As Integer
    Dim rsColumns As New ADODB.Recordset
    Dim lngCount As Long
    Dim strTemp As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim objCollection As New Collection
    Dim objCollectionDB As New Collection
    Dim objCollectionBakDB As New Collection
    Dim blnExist As Boolean
    Dim intLoop As Integer
    Dim strScriptFieldInfo As String
    Dim varScriptFieldInfo As Variant
    Dim strDBFieldInfo As String
    Dim varDBFieldInfo As Variant
    Dim objClsTable As New clsCollectionRecordset
    Dim objClsBakTableList As New clsCollectionRecordset
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    RaiseEvent BeforeProgress("正在检查数据表", 100)
    
    If Not mclsRunScript.OpenFile(strFile) Or mclsRunScript.EOF Then Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
    If mblnDBAUser Then
        mstrSQL = "Select TABLE_NAME From DBA_TABLES Where OWNER='" & mstrOwner & "'"
    Else
        mstrSQL = "Select TABLE_NAME From USER_TABLES"
    End If
    If rsColumns.State = adStateOpen Then rsColumns.Close
    rsColumns.Open mstrSQL, gcnOracle, adOpenKeyset
    Call objClsTable.LoadData(rsColumns, Array(), Array("TABLE_NAME"))
                        
    '列数据
    '------------------------------------------------------------------------------------------------------------------
    If mblnDBAUser Then
        strTemp = "SELECT TABLE_NAME,COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT From DBA_TAB_COLUMNS WHERE OWNER='" & mstrOwner & "' And Table_Name Not Like 'H%'"
    Else
        strTemp = "SELECT TABLE_NAME,COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT From USER_TAB_COLUMNS Where Table_Name Not Like 'H%'"
    End If
    If rsColumns.State = adStateOpen Then rsColumns.Close
    rsColumns.Open strTemp, gcnOracle, adOpenKeyset
    If rsColumns.BOF = False Then
        Do While Not rsColumns.EOF
            strTemp = rsColumns("TABLE_NAME").value & "'" & rsColumns("COLUMN_NAME").value & "'" & rsColumns("DATA_TYPE").value & "'" & rsColumns("DATA_LENGTH").value & "'" & rsColumns("DATA_PRECISION").value & "'" & rsColumns("DATA_SCALE").value & "'" & rsColumns("DATA_DEFAULT").value
            objCollectionDB.Add strTemp, "K" & rsColumns("TABLE_NAME").value & "'" & rsColumns("COLUMN_NAME").value
            rsColumns.MoveNext
        Loop
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Do While Not mclsRunScript.EOF
        mstrSQL = UCase(TrimEx(mclsRunScript.SQLInfo.SQL))
        strSQL1 = mstrSQL
        arySql = Split(mstrSQL, " TABLE ")
        mstrSQL = Trim(arySql(1)) '已经去掉Oracle关键字
        If InStr(1, mstrSQL, " ") > 0 And InStr(mstrSQL, " ") < InStr(mstrSQL, "(") Then
            strObjName = Trim(Left(mstrSQL, InStr(mstrSQL, " ")))
        Else
            strObjName = Trim(Left(mstrSQL, InStr(mstrSQL, "(") - 1))
        End If
        strObjName = Replace(strObjName, vbCrLf, "")
        
        If objClsTable.Item(Array(strObjName)) = False Then
            If blnAddToRsTable Then
                '1-存在对象,2-不存在对象,3-失效
                 Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "数据表", 2, False, strSQL1, "不存在", "严重：部分功能不能正常运行")
            Else
                Call InputErrRpt("数据表", strObjName, "不存在", "严重：部分功能不能正常运行")
            End If
        Else
            'new:从脚本中分析出字段、类型、长度
            '--------------------------------------------------------------------------------------------------
            Set objCollection = New Collection
            If AnsyTableScript(strObjName, mstrSQL, objCollection, objCollectionConstraintInTab) = True Then
                '开始比较
                For intLoop = 1 To objCollection.Count
                    strScriptFieldInfo = objCollection.Item(intLoop)
                    varScriptFieldInfo = Split(strScriptFieldInfo, "'")
                    
                    On Error Resume Next
                    strDBFieldInfo = ""
                    strDBFieldInfo = objCollectionDB.Item("K" & strObjName & "'" & varScriptFieldInfo(0))
                    varDBFieldInfo = Split(strDBFieldInfo, "'")
                    On Error GoTo errHand
                    
                    '按脚本中的字段数和当前数据库中的字段进行比较，找出差字段、字段类型、字段长度等不一样的情况
                    If CompareField(strObjName, varDBFieldInfo, varScriptFieldInfo, blnAddToRsTable) = True Then
                    
                    End If

                Next
            End If
            '--------------------------------------------------------------------------------------------------
        End If
        intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
        If intSvrPercent <> intPercent Then
            intSvrPercent = intPercent
            RaiseEvent Progressing(intSvrPercent)
        End If
        Call mclsRunScript.ReadNextSQL
    Loop

    RaiseEvent AfterProgress
    
    CheckTable = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function AnsyTableScript(ByVal strObjName As String, ByVal strTableScript As String, ByRef objCollection As Collection, Optional ByRef objCollectionConstraintInTab As Collection) As Boolean
    '******************************************************************************************************************
    '功能：分析
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngCount As Long
    Dim lngPostion As Long
    Dim varLine As Variant
    Dim varItem As Variant
    Dim strLine As String
    Dim strTemp As String
    Dim strFieldName As String
    Dim strFieldType As String
    Dim strFieldLength As String
    Dim strFieldDefault As String
    Dim blnField As Boolean
    Dim intLoop As Integer
    Dim intCol As Integer
    
    On Error GoTo errHand
    
    '去掉表名
    lngPostion = InStr(strTableScript, "(")
    If lngPostion > 0 Then
        strTableScript = Trim(Mid(strTableScript, lngPostion + 1))
        
        If Left(strTableScript, 2) = vbCrLf Then strTableScript = Trim(Mid(strTableScript, 3))
                    
        '去掉表名
        lngPostion = InStr(strTableScript, "TABLESPACE")
        If lngPostion = 0 Then
            lngPostion = InStr(strTableScript, "PCTFREE")
            If lngPostion = 0 Then
                lngPostion = InStr(UCase(strTableScript), UCase("Commit"))
            End If
        End If
        
        If lngPostion > 0 Then
            strTableScript = Trim(Mid(strTableScript, 1, lngPostion - 1))
            If Right(strTableScript, 2) = vbCrLf Then strTableScript = Trim(Mid(strTableScript, 1, Len(strTableScript) - 2))
            If Right(strTableScript, 1) = ")" Then strTableScript = Trim(Mid(strTableScript, 1, Len(strTableScript) - 1))
            
            varLine = Split(strTableScript, vbCrLf)
            
            lngCount = UBound(varLine)
            For lngLoop = 0 To lngCount
                strFieldDefault = ""
                strLine = Trim(varLine(lngLoop))
                
                lngPostion = InStr(strLine, "--")
                If lngPostion > 0 Then strLine = Trim(Mid(strLine, 1, lngPostion - 1))
                
                blnField = (strLine <> "" And strLine <> ")")
                If blnField Then
                    
                    If Not (UCase(strLine) Like "*NUMBER*" Or _
                        UCase(strLine) Like "*VARCHAR*" Or _
                        UCase(strLine) Like "*DATE*" Or _
                        UCase(strLine) Like "*NUMERIC*" Or _
                        UCase(strLine) Like "*BLOB*" Or _
                        UCase(strLine) Like "*CLOB*" Or _
                        UCase(strLine) Like "*LONG*") Then
                        
                        blnField = False
                        
                    End If
                    
                    If blnField Then blnField = Not (InStr(UCase(strLine), UCase("Store")) > 0 And InStr(UCase(strLine), UCase("Cache")) > 0)
                    If blnField Then blnField = Not (InStr(UCase(strLine), UCase("ORGANIZATION")) > 0 And InStr(UCase(strLine), UCase("INDEX")) > 0)
                    
                    If InStr(UCase(strLine), UCase("CONSTRAINT")) > 0 And InStr(UCase(strLine), UCase("PRIMARY")) > 0 Then
                        '写在表中的约束，需要先记录下来
                        
                        strTemp = "Alter Table " & strObjName & " Add " & strLine
                        
                        objCollectionConstraintInTab.Add strTemp, "K" & strObjName
                        
                        blnField = False
                    End If
                    
                End If
                
                If blnField = True Then
                  If Right(strLine, 1) = "," Then strLine = Mid(strLine, 1, Len(strLine) - 1)
                  strLine = Trim(Replace(strLine, vbTab, " "))
                  varItem = Split(strLine, " ")
                  
                  strFieldName = varItem(0)
                  For intLoop = 1 To UBound(varItem)
                      If Trim(varItem(intLoop)) <> "" Then
                          strTemp = ""
                          For intCol = intLoop To UBound(varItem)
                              If UCase(varItem(intCol)) <> "DEFAULT" Then
                                    Select Case UCase(varItem(intCol))
                                    Case "NOT", "NULL"
                                        
                                    Case Else
                                        strTemp = strTemp & varItem(intCol)
                                    End Select
                              Else
                                    strFieldDefault = varItem(intCol + 1)
                                Exit For
                              End If
                          Next
                          Exit For
                      End If
                  Next
    
                  strFieldLength = "0"
                  lngPostion = InStr(strTemp, "(")
                  If lngPostion > 0 Then
                      strFieldType = Mid(strTemp, 1, lngPostion - 1)
                      strTemp = Mid(strTemp, lngPostion + 1)
                      lngPostion = InStr(strTemp, ")")
                      If lngPostion > 0 Then
                          strFieldLength = Trim(Mid(strTemp, 1, lngPostion - 1))
                      End If
                  Else
                      If Right(strTemp, 1) = ")" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
                      strFieldType = strTemp
                  End If
                  
                    If UCase(strFieldType) = UCase("longraw") Then strFieldType = "LONG RAW"
                    If UCase(strFieldType) = UCase("varchar") Then strFieldType = "VARCHAR2"
                    If UCase(strFieldType) = UCase("numeric") Then strFieldType = "NUMBER"
        
                  objCollection.Add strFieldName & "'" & strFieldType & "'" & strFieldLength & "'" & strFieldDefault, "K" & strFieldName
                End If
            Next
            
        End If
    End If
    
    AnsyTableScript = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function CompareField(ByVal strTableName As String, ByVal varDBField As Variant, ByVal varScriptField As Variant, Optional ByVal blnAddToRsTable As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：检查
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim strFieldName As String
    Dim strFieldType As String
    Dim strFieldLength As String
    Dim strFieldDefault As String
    Dim str原字段精度 As String
    Dim str现字段精度 As String
    Dim strTemp As String
    Dim strSQL As String
    Dim strSQL1 As String
    Dim intPrecison As Integer
    Dim intScale As Integer
    
    Dim strDBFieldName As String
    Dim strDBFieldType As String
    Dim intDBFieldLength As Integer
    Dim intDBFieldPrecision As Integer
    Dim intDBFieldScale As Integer
    Dim strDBFieldDefault As String
    
    On Error GoTo errHand
        
    If strTableName = "病历范文目录" Then
        strTableName = "病历范文目录"
    End If
    
    If UBound(varDBField) <> -1 Then
        strDBFieldName = varDBField(1)
        strDBFieldType = varDBField(2)
        intDBFieldLength = Val(varDBField(3))
        intDBFieldPrecision = Val(varDBField(4))
        intDBFieldScale = Val(varDBField(5))
        strDBFieldDefault = varDBField(6)
    Else
        strDBFieldName = ""
    End If
                                    
    strFieldName = varScriptField(0)
    strFieldType = varScriptField(1)
    strFieldLength = varScriptField(2)
    strFieldDefault = varScriptField(3)
    
    If InStr(strFieldLength, ",") Then
        intPrecison = Val(Mid(strFieldLength, 1, InStr(strFieldLength, ",") - 1))
        intScale = Val(Mid(strFieldLength, InStr(strFieldLength, ",") + 1))
    Else
        intPrecison = Val(strFieldLength)
        intScale = 0
        strFieldLength = intPrecison
    End If
    
    If strDBFieldName = "" Then
        '缺少的字段
        Select Case UCase(strFieldType)
        Case "NUMBER"
            strTemp = strFieldName & " NUMBER(" & strFieldLength & ")"
            If strFieldDefault <> "" Then strTemp = strTemp & " DEFAULT " & strFieldDefault
            str现字段精度 = strFieldLength
        Case "VARCHAR2"
            strTemp = strFieldName & " VARCHAR2(" & strFieldLength & ")"
            If strFieldDefault <> "" Then strTemp = strTemp & " DEFAULT " & strFieldDefault
             str现字段精度 = strFieldLength
        Case Else
            strTemp = strFieldName & " " & strFieldType
        End Select
        
        If blnAddToRsTable Then
            '1-存在对象,2-不存在对象,3-失效,4-缺少列
             strSQL = " Alter Table " & strTableName & " Add(" & strTemp & ")"
             Call zlInsertRecData(mrsErrTable, strTableName, strTableName, "数据表", 4, False, strSQL, "缺少列 " & strTemp, "严重：部分功能不能正常运行", strFieldName, "", strFieldType, str原字段精度, str现字段精度)
        Else
            Call InputErrRpt("数据表", strTableName, "缺少列 " & strTemp, "严重：部分功能不能正常运行")
        End If
    Else
        '长度不等
        Select Case UCase(strFieldType)
        '----------------------------------------------------------------------------------
        Case "NUMBER"
            If intPrecison > intDBFieldPrecision Or intScale > intDBFieldScale Then
                If intDBFieldScale = 0 Then
                    strTemp = strFieldName & "列长度小于规定值：应为“" & "NUMBER(" & strFieldLength & ")”" & " 现为“" & "NUMBER(" & intDBFieldPrecision & ")”"
                Else
                    strTemp = strFieldName & "列长度小于规定值：应为“" & "NUMBER(" & strFieldLength & ")”" & " 现为“" & "NUMBER(" & intDBFieldPrecision & "," & intDBFieldScale & ")”"
                End If
                
                str原字段精度 = intDBFieldPrecision & "," & intDBFieldScale
                str现字段精度 = strFieldLength
                
                If blnAddToRsTable Then
                    If Nvl(strDBFieldType) = strFieldType Then
                        '类型相同的情况下，才处理
                        '1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度问题
                         
                         strSQL1 = " Alter Table " & strTableName & " Modify(" & strFieldName & " NUMBER(" & strFieldLength & "))"
                         
                        If intPrecison > intDBFieldPrecision Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strTableName, "数据表", 5, False, strSQL1, "精度过小 ", "严重：较大的数据将无法正确存储：" & strTemp, Nvl(strFieldName), strDBFieldType, strFieldType, str原字段精度, str现字段精度)
                        ElseIf intScale > intDBFieldScale Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strTableName, "数据表", 5, False, strSQL1, "精度过小 ", "严重：可能导致数据精度不足：" & strTemp, Nvl(strFieldName), strDBFieldType, strFieldType, str原字段精度, str现字段精度)
                        Else
                            Call InputErrRpt("数据表", strTableName, strTemp, "较轻：基本不影响运行")
                            Call zlInsertRecData(mrsErrTable, strTableName, strTableName, "数据表", 5, False, strSQL1, "精度过大 ", "较轻：基本不影响运行：" & strTemp, Nvl(strFieldName), strDBFieldType, strFieldType, str原字段精度, str现字段精度)
                        End If
                    End If
                Else
                    If intPrecison > intDBFieldPrecision Then
                        Call InputErrRpt("数据表", strTableName, strTemp, "较重：较大的数据将无法正确存储")
                    ElseIf intScale > intDBFieldScale Then
                        Call InputErrRpt("数据表", strTableName, strTemp, "较重：可能导致数据精度不足")
                    Else
                        Call InputErrRpt("数据表", strTableName, strTemp, "较轻：基本不影响运行")
                    End If
                End If
                         
            End If
        '----------------------------------------------------------------------------------
        Case "VARCHAR2"
            If intPrecison <> intDBFieldLength Then
                strTemp = strFieldName & "列长度小于规定值：应为“" & "VARCHAR2(" & intPrecison & ")”" & " 现为“" & "VARCHAR2(" & intDBFieldLength & ")”"
                str原字段精度 = intDBFieldLength
                str现字段精度 = intPrecison
                If blnAddToRsTable Then
                    If Nvl(strDBFieldType) = strFieldType Then
                        '类型相同的情况下，才处理
                        '1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度问题
                         strSQL1 = " Alter Table " & strTableName & " Modify(" & strFieldName & " VARCHAR2(" & intPrecison & ")" & ")"
                        
                        If intPrecison > intDBFieldLength Then
                            Call zlInsertRecData(mrsErrTable, strTableName, strTableName, "数据表", 5, False, strSQL1, "精度过小", "较重：可能导致较长文本无法存储：" & strTemp, Nvl(strFieldName), strDBFieldType, strFieldType, str原字段精度, str现字段精度)
                        Else
                            Call zlInsertRecData(mrsErrTable, strTableName, strTableName, "数据表", 5, False, strSQL1, "精度过大", "较轻：基本不影响运行：" & strTemp, Nvl(strFieldName), strDBFieldType, strFieldType, str原字段精度, str现字段精度)
                        End If
                    End If
                Else
                    If intPrecison > intDBFieldLength Then
                        Call InputErrRpt("数据表", strTableName, strTemp, "较重：可能导致较长文本无法存储")
                    Else
                        Call InputErrRpt("数据表", strTableName, strTemp, "较轻：基本不影响运行")
                    End If
                End If
            End If
        '----------------------------------------------------------------------------------
        Case Else
            
        End Select
        
        If strDBFieldType <> strFieldType Then
             strTemp = strFieldName & "列类型不一至：应为“" & strFieldType & "” 现为：“" & strDBFieldType & "”"
             If blnAddToRsTable Then
                Call zlInsertRecData(mrsErrTable, strTableName, strTableName, "数据表", 5, False, "", "字段类型不对", "较重：可能导致数据不能存储!" & strTemp, strFieldName, strDBFieldType, strFieldType, "", "")
             Else
                Call InputErrRpt("数据表", strTableName, strTemp, "较重：可能导致数据不能存储!")
             End If
        End If
                                    
    End If
        
    CompareField = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function CompareTableColumn(ByVal strTableName As String, ByVal varSourceField As Variant, ByVal varTargeField As Variant, ByRef colResult As Collection) As Boolean
    '******************************************************************************************************************
    '功能：检查
    '参数：
    '返回：成功返回true，否则返回false
    '*****************************************************************************************************************
    Dim strFieldName As String
    Dim strFieldType As String
    Dim strFieldLength As String
    Dim strFieldDefault As String
    Dim str原字段精度 As String
    Dim str现字段精度 As String
    Dim strTemp As String
    Dim strSQL As String
    Dim strSQL1 As String
    Dim intPrecison As Integer
    Dim intScale As Integer
    
    Dim strDBFieldName As String
    Dim strDBFieldType As String
    Dim intDBFieldLength As Integer
    Dim intDBFieldPrecision As Integer
    Dim intDBFieldScale As Integer
    Dim strDBFieldDefault As String
    
    On Error GoTo errHand
        
    If strTableName = "病历范文目录" Then
        strTableName = "病历范文目录"
    End If
    
    If UBound(varTargeField) <> -1 Then
        strDBFieldName = varTargeField(1)
        strDBFieldType = varTargeField(2)
        intDBFieldLength = Val(varTargeField(3))
        intDBFieldPrecision = Val(varTargeField(4))
        intDBFieldScale = Val(varTargeField(5))
        strDBFieldDefault = varTargeField(6)
    Else
        strDBFieldName = ""
    End If
                                    
    strFieldName = varSourceField(0)
    strFieldType = varSourceField(1)
    strFieldLength = varSourceField(2)
    strFieldDefault = varSourceField(3)
    
    If InStr(strFieldLength, ",") Then
        intPrecison = Val(Mid(strFieldLength, 1, InStr(strFieldLength, ",") - 1))
        intScale = Val(Mid(strFieldLength, InStr(strFieldLength, ",") + 1))
    Else
        intPrecison = Val(strFieldLength)
        intScale = 0
        strFieldLength = intPrecison
    End If

    If strDBFieldName = "" Then

        '缺少的字段
        Select Case UCase(strFieldType)
        Case "NUMBER"
            strTemp = strFieldName & " NUMBER(" & strFieldLength & ")"
            If strFieldDefault <> "" Then strTemp = strTemp & " DEFAULT " & strFieldDefault
'            str现字段精度 = strFieldLength
        Case "VARCHAR2"
            strTemp = strFieldName & " VARCHAR2(" & strFieldLength & ")"
            If strFieldDefault <> "" Then strTemp = strTemp & " DEFAULT " & strFieldDefault
'             str现字段精度 = strFieldLength
        Case Else
            strTemp = strFieldName & " " & strFieldType
        End Select
        
        colResult.Add Array("数据表", strTableName, "缺少列 " & strTemp, "严重：部分功能不能正常运行")
    Else

        '长度不等
        Select Case UCase(strFieldType)
        '----------------------------------------------------------------------------------
        Case "NUMBER"
            If intPrecison > intDBFieldPrecision Or intScale > intDBFieldScale Then
                If intDBFieldScale = 0 Then
                    strTemp = strFieldName & "列长度小于规定值：应为“" & "NUMBER(" & strFieldLength & ")”" & " 现为“" & "NUMBER(" & intDBFieldPrecision & ")”"
                Else
                    strTemp = strFieldName & "列长度小于规定值：应为“" & "NUMBER(" & strFieldLength & ")”" & " 现为“" & "NUMBER(" & intDBFieldPrecision & "," & intDBFieldScale & ")”"
                End If

                If intPrecison > intDBFieldPrecision Then
                    colResult.Add Array("数据表", strTableName, strTemp, "较重：较大的数据将无法正确存储")
                ElseIf intScale > intDBFieldScale Then
                    colResult.Add Array("数据表", strTableName, strTemp, "较重：可能导致数据精度不足")
                Else
                    colResult.Add Array("数据表", strTableName, strTemp, "较轻：基本不影响运行")
                End If
                         
            End If
        '----------------------------------------------------------------------------------
        Case "VARCHAR2"
            If intPrecison <> intDBFieldLength Then
                strTemp = strFieldName & "列长度小于规定值：应为“" & "VARCHAR2(" & intPrecison & ")”" & " 现为“" & "VARCHAR2(" & intDBFieldLength & ")”"
                
                If intPrecison > intDBFieldLength Then
                    colResult.Add Array("数据表", strTableName, strTemp, "较重：可能导致较长文本无法存储")
                Else
                    colResult.Add Array("数据表", strTableName, strTemp, "较轻：基本不影响运行")
                End If

            End If
        '----------------------------------------------------------------------------------
        Case Else
            
        End Select
        
        If strDBFieldType <> strFieldType Then
             strTemp = strFieldName & "列类型不一至：应为“" & strFieldType & "” 现为：“" & strDBFieldType & "”"
             colResult.Add Array("数据表", strTableName, strTemp, "较重：可能导致数据不能存储!")
        End If
        
    End If
        
    CompareTableColumn = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function CompareFieldTool(ByVal strTableName As String, ByVal varDBField As Variant, ByVal varScriptField As Variant, ByRef cllTemp As Collection) As Boolean
    '******************************************************************************************************************
    '功能：检查
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim strFieldName As String
    Dim strFieldType As String
    Dim strFieldLength As String
    Dim strFieldDefault As String
    Dim str原字段精度 As String
    Dim str现字段精度 As String
    Dim strTemp As String
    Dim strSQL As String
    Dim strSQL1 As String
    Dim intPrecison As Integer
    Dim intScale As Integer
    
    Dim strDBFieldName As String
    Dim strDBFieldType As String
    Dim intDBFieldLength As Integer
    Dim intDBFieldPrecision As Integer
    Dim intDBFieldScale As Integer
    Dim strDBFieldDefault As String
    
    On Error GoTo errHand
        
    If strTableName = "病历范文目录" Then
        strTableName = "病历范文目录"
    End If
    
    If UBound(varDBField) <> -1 Then
        strDBFieldName = varDBField(1)
        strDBFieldType = varDBField(2)
        intDBFieldLength = Val(varDBField(3))
        intDBFieldPrecision = Val(varDBField(4))
        intDBFieldScale = Val(varDBField(5))
        strDBFieldDefault = varDBField(6)
    Else
        strDBFieldName = ""
    End If
                                    
    strFieldName = varScriptField(0)
    strFieldType = varScriptField(1)
    strFieldLength = varScriptField(2)
    strFieldDefault = varScriptField(3)
    
    If InStr(strFieldLength, ",") Then
        intPrecison = Val(Mid(strFieldLength, 1, InStr(strFieldLength, ",") - 1))
        intScale = Val(Mid(strFieldLength, InStr(strFieldLength, ",") + 1))
    Else
        intPrecison = Val(strFieldLength)
        intScale = 0
        strFieldLength = intPrecison
    End If

    If strDBFieldName = "" Then

        '缺少的字段
        Select Case UCase(strFieldType)
        Case "NUMBER"
            strTemp = strFieldName & " NUMBER(" & strFieldLength & ")"
            If strFieldDefault <> "" Then strTemp = strTemp & " DEFAULT " & strFieldDefault
'            str现字段精度 = strFieldLength
        Case "VARCHAR2"
            strTemp = strFieldName & " VARCHAR2(" & strFieldLength & ")"
            If strFieldDefault <> "" Then strTemp = strTemp & " DEFAULT " & strFieldDefault
'             str现字段精度 = strFieldLength
        Case Else
            strTemp = strFieldName & " " & strFieldType
        End Select
        
        cllTemp.Add Array("数据表", strTableName, "缺少列 " & strTemp, "严重：部分功能不能正常运行")
    Else

        '长度不等
        Select Case UCase(strFieldType)
        '----------------------------------------------------------------------------------
        Case "NUMBER"
            If intPrecison > intDBFieldPrecision Or intScale > intDBFieldScale Then
                If intDBFieldScale = 0 Then
                    strTemp = strFieldName & "列长度小于规定值：应为“" & "NUMBER(" & strFieldLength & ")”" & " 现为“" & "NUMBER(" & intDBFieldPrecision & ")”"
                Else
                    strTemp = strFieldName & "列长度小于规定值：应为“" & "NUMBER(" & strFieldLength & ")”" & " 现为“" & "NUMBER(" & intDBFieldPrecision & "," & intDBFieldScale & ")”"
                End If

                If intPrecison > intDBFieldPrecision Then
                    cllTemp.Add Array("数据表", strTableName, strTemp, "较重：较大的数据将无法正确存储")
                ElseIf intScale > intDBFieldScale Then
                    cllTemp.Add Array("数据表", strTableName, strTemp, "较重：可能导致数据精度不足")
                Else
                    cllTemp.Add Array("数据表", strTableName, strTemp, "较轻：基本不影响运行")
                End If
                         
            End If
        '----------------------------------------------------------------------------------
        Case "VARCHAR2"
            If intPrecison <> intDBFieldLength Then
                strTemp = strFieldName & "列长度小于规定值：应为“" & "VARCHAR2(" & intPrecison & ")”" & " 现为“" & "VARCHAR2(" & intDBFieldLength & ")”"
                
                If intPrecison > intDBFieldLength Then
                    cllTemp.Add Array("数据表", strTableName, strTemp, "较重：可能导致较长文本无法存储")
                Else
                    cllTemp.Add Array("数据表", strTableName, strTemp, "较轻：基本不影响运行")
                End If

            End If
        '----------------------------------------------------------------------------------
        Case Else
            
        End Select
        
        If strDBFieldType <> strFieldType Then
             strTemp = strFieldName & "列类型不一至：应为“" & strFieldType & "” 现为：“" & strDBFieldType & "”"
             cllTemp.Add Array("数据表", strTableName, strTemp, "较重：可能导致数据不能存储!")
        End If
        
    End If
        
    CompareFieldTool = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function


Private Function CheckConstraint(ByVal strFile As String, ByVal objCollectionConstraintInTab As Collection, Optional ByVal blnAddToRsTable As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：检查对象约束，判断是否有效存在
    '参数：FileName-脚本文件名(包含具体路径)
    '       blnAddToRsTable-是否将相关的错误信息写在记录集中
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim strSQL1 As String, strTemp1 As String
    Dim arySql() As String, strObjName As String, strColumns As String, strTableName As String
    Dim rsColumns As New ADODB.Recordset
    Dim rsConstraints As New ADODB.Recordset
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim strTemp As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim objCollectionColumn As New Collection
    Dim objCollectionColumnBak As New Collection
    Dim objClsConstraintsBak As New clsCollectionRecordset
    Dim objClsConstraints  As New clsCollectionRecordset
    Dim objClsBakTable  As New clsCollectionRecordset
    Dim strSvrKey As String
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
        
    RaiseEvent BeforeProgress("正在检查约束", 100)
    If Not mclsRunScript.OpenFile(strFile) Or mclsRunScript.EOF Then
        '如果脚本文件为空或者之前表结构脚本中也不存在约束，则直接退出约束对象检查
        If objCollectionConstraintInTab.Count = 0 Then Exit Function
    End If
    '收集数据库中的约束，并记录到集合中
    '------------------------------------------------------------------------------------------------------------------
    If mblnDBAUser Then
        mstrSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD,Search_Condition from DBA_CONSTRAINTS where OWNER='" & mstrOwner & "'"
    Else
        mstrSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD,Search_Condition from USER_CONSTRAINTS"
    End If
    If rsColumns.State = adStateOpen Then rsColumns.Close
    rsColumns.Open mstrSQL, gcnOracle, adOpenKeyset
    Call objClsConstraints.LoadData(rsColumns, Array("CONSTRAINT_TYPE", "STATUS", "VALIDATED", "BAD", "Search_Condition"), Array("CONSTRAINT_NAME"))
    
    
    '收集数据库中的约束列项，并记录到集合中
    '------------------------------------------------------------------------------------------------------------------
    strColumns = ""
    If mblnDBAUser Then
        strTemp = "Select Distinct CONSTRAINT_NAME,POSITION,COLUMN_NAME from DBA_CONS_COLUMNS where OWNER='" & mstrOwner & "' order by CONSTRAINT_NAME,POSITION"
    Else
        strTemp = "Select Distinct CONSTRAINT_NAME,POSITION,COLUMN_NAME  from USER_CONS_COLUMNS Order by CONSTRAINT_NAME,POSITION"
    End If
    If rsColumns.State = adStateOpen Then rsColumns.Close
    rsColumns.Open strTemp, gcnOracle, adOpenKeyset
    Do While Not rsColumns.EOF
        If strSvrKey <> rsColumns!Constraint_Name Then
            If strSvrKey <> "" Then
                objCollectionColumn.Add strColumns, "K" & strSvrKey
            End If
            strSvrKey = rsColumns!Constraint_Name
            strColumns = ""
        End If
        
        strColumns = strColumns & "," & rsColumns!Column_Name
        rsColumns.MoveNext
    Loop
    If strColumns <> "" Then objCollectionColumn.Add strColumns, "K" & strSvrKey
    
    
    '收集数据库中的关联外键约束，并记录到集合中
    '------------------------------------------------------------------------------------------------------------------
    If mblnDBAUser Then
        strTemp = "Select R_Constraint_Name,table_Name,Constraint_Name,OWNER,R_OWNER,R_Constraint_Name,DELETE_RULE from DBA_CONSTRAINTS Where  Constraint_Type='R'"
    Else
        strTemp = "Select R_Constraint_Name,table_Name,Constraint_Name,OWNER,R_OWNER,R_Constraint_Name,DELETE_RULE From USER_CONSTRAINTS Where Constraint_Type='R'"
    End If
    If rsConstraints.State = adStateOpen Then rsConstraints.Close
    rsConstraints.Open strTemp, gcnOracle, adOpenKeyset
    
    
    '先检查表结构脚本中的组织索引表中的PK约束
    '--------------------------------------------------------------------------------------------------------------
    If objCollectionConstraintInTab.Count > 0 Then
        For lngLoop = 1 To objCollectionConstraintInTab.Count
            strTemp = objCollectionConstraintInTab.Item(lngLoop)
            Call CheckConstraintItem(objClsConstraints, objCollectionColumn, objClsBakTable, objCollectionColumnBak, strTemp, rsConstraints, blnAddToRsTable, objClsConstraintsBak)
        Next
    End If
    
    '检查约束脚本中的约束
    '--------------------------------------------------------------------------------------------------------------

    On Error GoTo 0
    
    Do While Not mclsRunScript.EOF
        Call CheckConstraintItem(objClsConstraints, objCollectionColumn, objClsBakTable, objCollectionColumnBak, mclsRunScript.SQLInfo.SQL, rsConstraints, blnAddToRsTable, objClsConstraintsBak)
        intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
        If intSvrPercent <> intPercent Then
            intSvrPercent = intPercent
            RaiseEvent Progressing(intSvrPercent)
        End If
        Call mclsRunScript.ReadNextSQL
    Loop

    RaiseEvent AfterProgress
    
    CheckConstraint = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function CheckForeignIndex(ByVal strFile As String, ByVal objCollectionConstraintInTab As Collection, Optional ByVal blnAddToRsTable As Boolean = False) As Boolean
    
    '检查多字段外键上的索引情况(只限于检查脚本本身的设计是否正确）
    
    
    
    '外键字段(a,b,c)|a'b''a,b
    
    
    
    '外键字段仅有一部分字段存在索引
    '------------------------------------------------------------------------------------------------------------------
    
    
    
    
    '外键字段不是组合索引中的前导字段
    '------------------------------------------------------------------------------------------------------------------
        
    
    
End Function

Private Function CheckConstraintItem(ByVal objClsConstraints As clsCollectionRecordset, _
                                    ByVal objCollectionColumn As Collection, _
                                    ByVal objClsBakTable As clsCollectionRecordset, _
                                    ByVal objCollectionColumnBak As Collection, _
                                    ByVal strConstraintSQL As String, _
                                    ByVal rsConstraints As ADODB.Recordset, _
                                    Optional ByVal blnAddToRsTable As Boolean = False, _
                                    Optional ByVal objClsConstraintsBak As clsCollectionRecordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    
    Dim strColumns As String
    Dim arySql() As String
    Dim strTemp As String
    Dim strTemp1 As String
    Dim strTableName As String
    Dim strObjName As String
    Dim strSQL1 As String
'    Dim objCls As New clsCollectionRecordset
    
    strConstraintSQL = UCase(TrimEx(strConstraintSQL, True))
    strSQL1 = strConstraintSQL
    arySql = Split(strConstraintSQL, " CONSTRAINT ")
    
    If UBound(arySql) > 0 Then
        '刘兴宏加入:
        strTableName = Trim(Split(Trim(arySql(0)), "TABLE")(1))
        strTableName = Split(strTableName, " ")(0)
        strConstraintSQL = Trim(arySql(1)) '已经去掉Oracle关键字
        strObjName = Trim(Left(strConstraintSQL, InStr(strConstraintSQL, " ")))
                        
        '------------------------------------------------------------------------------------------------------------------
        If objClsConstraints.Item(Array(strObjName)) = False Then
    
            If blnAddToRsTable Then
                '1-存在对象,2-不存在对象,3-失效
                '类型:主键,唯一,外键,约束,索引,视图,...
                If InStr(1, UCase(Replace(strConstraintSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "外键", 2, False, strSQL1, "不存在", "较重：可能导致数据不一致，影响运行速度")
                ElseIf InStr(1, strConstraintSQL, " CHECK") > 0 Then
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 2, False, strSQL1, "不存在", "较轻：基本不影响系统运行")
                Else
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 2, False, strSQL1, "不存在", "较重：可能导致数据不一致，影响运行速度")
                End If
            Else
                If InStr(1, strConstraintSQL, " CHECK") > 0 Then
                    Call InputErrRpt("约束", strObjName, "不存在", "较轻：基本不影响系统运行")
                Else
                    Call InputErrRpt("约束", strObjName, "不存在", "较重：可能导致数据不一致，影响运行速度")
                End If
            End If
        '------------------------------------------------------------------------------------------------------------------
        ElseIf objClsConstraints.ItemValue("STATUS") <> "ENABLED" Then
            If blnAddToRsTable Then
                '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态
                '类型:主键,唯一,外键,约束,索引,视图,...
                If InStr(1, UCase(Replace(strConstraintSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "外键", 6, False, strSQL1, "当前处于禁止状态", "较重：可能系统已经存在问题")
                Else
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 6, False, strSQL1, "当前处于禁止状态", "较重：可能系统已经存在问题")
                End If
            Else
                Call InputErrRpt("约束", strObjName, "当前处于禁止状态", "较重：可能系统已经存在问题")
            End If
        '------------------------------------------------------------------------------------------------------------------
        ElseIf objClsConstraints.ItemValue("VALIDATED") <> "VALIDATED" Then
            If blnAddToRsTable Then
                '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态
                '类型:主键,唯一,外键,约束,索引,视图,...
                If InStr(1, UCase(Replace(strConstraintSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "外键", 3, False, strSQL1, "当前处于无效状态", "较重：可能数据一致性已被破坏")
                Else
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 3, False, strSQL1, "当前处于无效状态", "较重：可能数据一致性已被破坏")
                End If
            Else
                Call InputErrRpt("约束", strObjName, "当前处于无效状态", "较重：可能数据一致性已被破坏")
            End If
        '------------------------------------------------------------------------------------------------------------------
        ElseIf objClsConstraints.ItemValue("BAD") <> "" Then
            If blnAddToRsTable Then
                '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态(损坏)
                '类型:主键,唯一,外键,约束,索引,视图,...
                If InStr(1, UCase(Replace(strConstraintSQL, " ", "")), UCase("ForeignKey")) > 0 Then
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "外键", 6, False, strSQL1, "约束被意外损坏", "严重：可能存在硬件错误")
                Else
                    Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 6, False, strSQL1, "约束被意外损坏", "严重：可能存在硬件错误")
                End If
            Else
                Call InputErrRpt("约束", strObjName, "约束被意外损坏", "严重：可能存在硬件错误")
            End If
        '------------------------------------------------------------------------------------------------------------------
        Else
            
            On Error Resume Next
            strColumns = ""
            strColumns = objCollectionColumn.Item("K" & strObjName)
            On Error GoTo errHand
            '--------------------------------------------------------------------------------------------------------------
            If InStr(1, strConstraintSQL, " PRIMARY ") > 0 Then
    
                If objClsConstraints.ItemValue("constraint_type") <> "P" Then
                    If blnAddToRsTable Then
                        '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束类型错误", "严重：可能影响系统运行 ,应为主键约束")
                        If objClsConstraints.ItemValue("constraint_type") = "U" Then
                            '同时，需要检查相关的级联外键
                            Call zl获取关联外键(rsConstraints, strObjName, mrsErrTable)
                        End If
                    Else
                        Call InputErrRpt("约束", strObjName, "约束类型错误，应为主键约束", "严重：可能影响系统运行")
                    End If
                Else
                    arySql = Split(strConstraintSQL, " PRIMARY ")
                    strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                    If strColumns <> "," & strTemp Then
                        If blnAddToRsTable Then
                            '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                            '类型:主键,唯一,外键,约束,索引,视图,...
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 4, False, strSQL1, "约束列错误", "较重：可能影响系统运行，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")")
                            '同时，需要检查相关的级联外键
                            Call zl获取关联外键(rsConstraints, strObjName, mrsErrTable)
                            
                        Else
                            Call InputErrRpt("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "较重：可能影响系统运行")
                        End If
                    End If
                End If
            '--------------------------------------------------------------------------------------------------------------
            ElseIf InStr(1, strConstraintSQL, " UNIQUE") > 0 Then
    
                If objClsConstraints.ItemValue("constraint_type") <> "U" Then
                    If blnAddToRsTable Then
                        '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束类型错误", "较重：可能影响系统运行 ,应为唯一约束")
    
                        If objClsConstraints.ItemValue("constraint_type") = "P" Then
                            '同时，需要检查相关的级联外键
                            Call zl获取关联外键(rsConstraints, strObjName, mrsErrTable)
                        End If
                    Else
                        Call InputErrRpt("约束", strObjName, "约束类型错误，应为唯一约束", "较重：可能影响系统运行")
                    End If
                Else
                    arySql = Split(strConstraintSQL, " UNIQUE ")
                    If UBound(arySql) = 0 Then arySql = Split(strConstraintSQL, " UNIQUE(")
                    strTemp = Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "(", ""), " ", "")
                    If strColumns <> "," & strTemp Then
                        If blnAddToRsTable Then
                            '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                            '类型:主键,唯一,外键,约束,索引,视图,...
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 4, False, strSQL1, "约束列错误", "较重：可能影响系统运行，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")")
                            '同时，需要检查相关的级联外键
                            Call zl获取关联外键(rsConstraints, strObjName, mrsErrTable)
                        Else
                            Call InputErrRpt("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "较重：可能影响系统运行")
                        End If
                    End If
                End If
            '--------------------------------------------------------------------------------------------------------------
            ElseIf InStr(1, strConstraintSQL, " FOREIGN ") > 0 Then
    
                If objClsConstraints.ItemValue("constraint_type") <> "R" Then
                    If blnAddToRsTable Then
                        '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束类型错误", "较重：可能影响系统运行，应为外键约束")
                    Else
                        Call InputErrRpt("约束", strObjName, "约束类型错误，应为外键约束", "严重：可能影响系统运行")
                    End If
                Else
                    arySql = Split(strConstraintSQL, " FOREIGN ")
                    strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                    If strColumns <> "," & strTemp Then
                        If blnAddToRsTable Then
                            '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                            '类型:主键,唯一,外键,约束,索引,视图,...
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 4, False, strSQL1, "约束列错误", "较重：可能影响系统运行，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")")
                        Else
                            Call InputErrRpt("约束", strObjName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "严重：可能影响数据一致性")
                        End If
                    End If
                End If
            '--------------------------------------------------------------------------------------------------------------
            ElseIf InStr(1, strConstraintSQL, " CHECK") > 0 Then
    
                If objClsConstraints.ItemValue("constraint_type") <> "C" Then
                    If blnAddToRsTable Then
                        '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束类型错误", "较重：可能影响系统运行，应为检查约束")
                    Else
                        Call InputErrRpt("约束", strObjName, "约束类型错误，应为检查约束", "严重：可能影响系统运行")
                    End If
                Else
                    '25047:刘兴洪增加
                    arySql = Split(strConstraintSQL, " CHECK")
                    strTemp = Replace(UCase(Replace(Replace(Replace(arySql(1), " ", ""), vbTab, ""), vbCrLf, "")), ";", "")
                    strTemp1 = "(" & Replace(UCase(Replace(Replace(Replace(Nvl(objClsConstraints.ItemValue("Search_Condition")), " ", ""), vbTab, ""), vbCrLf, "")), ";", "") & ")"
                    If strTemp <> strTemp1 Then
                        '检查约束检查是否一致
                        strTemp = Trim(Replace(Replace(Replace(arySql(1), vbCrLf, " "), vbTab, " "), ";", ""))
                        strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
                        
                        If blnAddToRsTable Then
                            '状态:'1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致
                            '类型:主键,唯一,外键,约束,索引,视图,...
                            Call zlInsertRecData(mrsErrTable, strTableName, strObjName, "约束", 7, False, strSQL1, "约束内容不一致", "较重：可能影响系统运行,导致数据不一致!,应为(" & strTemp & "),现在为(" & Nvl(objClsConstraints.ItemValue("Search_Condition")) & ")")
                        Else
                            Call InputErrRpt("约束", strObjName, "约束内容不一致,应为(" & strTemp & "),现在为(" & Nvl(objClsConstraints.ItemValue("Search_Condition")) & ")", "较重：可能影响系统运行,导致数据不一致!")
                        End If
                    End If
                End If
                
            
            End If
        End If
    End If
                            
    CheckConstraintItem = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
    
errHand:
    MsgBox err.Description
End Function

Private Function CheckIndex(ByVal strFile As String, Optional ByVal blnAddToRsTable As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能:检查对象索引，判断是否有效存在
    '入参:FileName-脚本文件名(包含具体路径)
    '     blnAddToRsTable-是否将相关的错误信息写在记录集中
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    On Error GoTo errHand

    Dim str原字段精度 As String, str现字段精度 As String, strSQL1 As String
    Dim arySql() As String, strObjName As String, strColumns As String
    Dim strTablenName As String
    
    Dim rsColumns As New ADODB.Recordset
    Dim lngCount As Long
    Dim strTemp As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim strSvrKey As String
    Dim objCollection As Collection
    Dim objCollectionBak As Collection
    
    Dim objClsIndex As New clsCollectionRecordset
    Dim objClsBakIndex As New clsCollectionRecordset
    Dim objClsBakTable  As New clsCollectionRecordset
    
    '------------------------------------------------------------------------------------------------------------------
        
    RaiseEvent BeforeProgress("正在检查索引", 100)
    
    If Not mclsRunScript.OpenFile(strFile) Or mclsRunScript.EOF Then Exit Function
        
    '------------------------------------------------------------------------------------------------------------------

    If mblnDBAUser Then
        mstrSQL = "select INDEX_NAME,STATUS from DBA_INDEXES where OWNER='" & mstrOwner & "' and TABLE_OWNER='" & mstrOwner & "'"
    Else
        mstrSQL = "select INDEX_NAME,STATUS from USER_INDEXES where TABLE_OWNER='" & mstrOwner & "'"
    End If
    If rsColumns.State = adStateOpen Then rsColumns.Close
    rsColumns.Open mstrSQL, gcnOracle, adOpenKeyset
    Call objClsIndex.LoadData(rsColumns, Array("STATUS"), Array("INDEX_NAME"))
    
    '------------------------------------------------------------------------------------------------------------------
    strColumns = ""
    If mblnDBAUser = True Then
        strTemp = "Select Distinct INDEX_NAME,COLUMN_POSITION,TABLE_NAME,COLUMN_NAME from DBA_IND_COLUMNS where INDEX_OWNER='" & mstrOwner & "' order by INDEX_NAME,COLUMN_POSITION"
    Else
        strTemp = "Select Distinct INDEX_NAME,COLUMN_POSITION,TABLE_NAME,COLUMN_NAME from USER_IND_COLUMNS Order by INDEX_NAME,COLUMN_POSITION"
    End If
    If rsColumns.State = adStateOpen Then rsColumns.Close
    rsColumns.Open strTemp, gcnOracle, adOpenKeyset
    
    Set objCollection = New Collection
    Do While Not rsColumns.EOF
        If strSvrKey <> rsColumns!Index_Name Then
            If strSvrKey <> "" Then
                strColumns = strColumns & ")"
                objCollection.Add strColumns, "K" & strSvrKey
            End If
            strSvrKey = rsColumns!Index_Name
            strColumns = rsColumns!Table_Name & "(" & rsColumns!Column_Name
        Else
            strColumns = strColumns & "," & rsColumns!Column_Name
        End If
        
        rsColumns.MoveNext
    Loop
    If strColumns <> "" Then objCollection.Add strColumns, "K" & strSvrKey
    
    '------------------------------------------------------------------------------------------------------------------

    '------------------------------------------------------------------------------------------------------------------

    On Error GoTo 0
    Do While Not mclsRunScript.EOF
        mstrSQL = UCase(TrimEx(mclsRunScript.SQLInfo.SQL, True))
        strSQL1 = mstrSQL
        arySql = Split(mstrSQL, " INDEX ")
        If UBound(arySql) > 0 Then
        
            mstrSQL = Trim(arySql(1)) '已经去掉Oracle关键字
            strTablenName = Trim(Split(Split(mstrSQL, "ON")(1), "(")(0))
            strObjName = Trim(Left(mstrSQL, InStr(mstrSQL, " ")))
            
            If objClsIndex.Item(Array(strObjName)) = False Then

                If blnAddToRsTable Then
                    '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态
                    '类型:主键,唯一,外键,约束,索引,视图,...
                    Call zlInsertRecData(mrsErrTable, strTablenName, strObjName, "索引", 2, False, strSQL1, "不存在", "较重：可能影响系统运行速度")
                Else
                    Call InputErrRpt("索引", strObjName, "不存在", "较重：可能影响系统运行速度")
                End If
            
            ElseIf objClsIndex.ItemValue("STATUS") <> "VALID" Then
                Call InputErrRpt("索引", strObjName, "当前处于无效状态")
            Else

                On Error Resume Next
                strColumns = ""
                strColumns = objCollection.Item("K" & strObjName)
                On Error GoTo errHand
                
                arySql = Split(mstrSQL, " ON ")
                strTemp = Replace(Left(arySql(1), InStr(1, arySql(1), ")")), " ", "")
                If strColumns <> strTemp Then
                    If blnAddToRsTable Then
                        '状态:1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态
                        '类型:主键,唯一,外键,约束,索引,视图,...
                        Call zlInsertRecData(mrsErrTable, strTablenName, strObjName, "索引", 4, False, strSQL1, "索引列错误", "较重：可能影响系统运行速度,应为“" & strTemp & "”，现为“" & strColumns & "”")
                    Else
                        Call InputErrRpt("索引", strObjName, "索引列错误，应为“" & strTemp & "”，现为“" & strColumns & "”", "较重：可能影响系统运行速度")
                    End If
                End If
            End If
        End If
        intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
        If intSvrPercent <> intPercent Then
            intSvrPercent = intPercent
            RaiseEvent Progressing(intSvrPercent)
        End If
        Call mclsRunScript.ReadNextSQL
    Loop
    
    RaiseEvent AfterProgress
    
    CheckIndex = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function CheckView(ByVal strFile As String) As Boolean
    '******************************************************************************************************************
    '功能：检查视图对象，判断是否存在
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim rsObjects As New ADODB.Recordset
    Dim rsColumns As New ADODB.Recordset
    Dim lngCount As Long
    Dim strTemp As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim arySql() As String
    Dim strObjName As String
    Dim objClsView As New clsCollectionRecordset
    
    On Error GoTo errHand
    
        
    RaiseEvent BeforeProgress("正在检查视图", 100)
    
    If Not mclsRunScript.OpenFile(strFile) Or mclsRunScript.EOF Then Exit Function
    '------------------------------------------------------------------------------------------------------------------
    If mblnDBAUser Then
        mstrSQL = "Select A.Object_Name View_Name From Dba_Objects a Where a.Owner = '" & mstrOwner & "' And A.Object_Type Like '%VIEW'"
    Else
        mstrSQL = "Select A.Object_Name View_Name From User_Objects A Where A.Object_Type Like '%VIEW'"
    End If
    If rsObjects.State = adStateOpen Then rsObjects.Close
    rsObjects.Open mstrSQL, gcnOracle, adOpenKeyset
    Call objClsView.LoadData(rsObjects, Array(), Array("VIEW_NAME"))
    
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo 0
    Do While Not mclsRunScript.EOF
        
        mstrSQL = UCase(TrimEx(mclsRunScript.SQLInfo.SQL, True))
        arySql = Split(mstrSQL, " VIEW ")
        mstrSQL = Trim(arySql(1)) '已经去掉Oracle关键字
        strObjName = Trim(Left(mstrSQL, InStr(mstrSQL, " ")))
        
        If objClsView.Item(Array(strObjName)) = False Then
            Call InputErrRpt("视图", strObjName, "不存在", "严重：部分功能不能正常运行")
        End If
        
        intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
        If intSvrPercent <> intPercent Then
            intSvrPercent = intPercent
            RaiseEvent Progressing(intSvrPercent)
        End If
        Call mclsRunScript.ReadNextSQL
    Loop

    RaiseEvent AfterProgress
    
    CheckView = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Function CheckProgram(ByVal strFile As String, Optional ByVal blnAddToRsTable As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能:检查函数和过程对象，判断是否存在
    '入参:FileName-脚本文件名(包含具体路径)
    '     blnAddToRsTable-是否将相关的错误信息写在记录集中
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim rsObjects As New ADODB.Recordset
    Dim rsColumns As New ADODB.Recordset
    Dim lngCount As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim objCollectionScriptPara As Collection
    Dim objCollectionDBPara As Collection
    Dim objCollectionAllDBPara As Collection
    Dim strSvrProcedureName As String
    Dim objCollectionTableFieldType As New Collection
    Dim intLoop As Integer
    Dim arySql() As String, strObjName As String, strSQL1 As String
    Dim blnExist As Boolean
    Dim strStatus As String
    Dim strType As String
    Dim objClsProcedure As New clsCollectionRecordset
    
    On Error GoTo errHand
        
    RaiseEvent BeforeProgress("正在检查函数/过程", 100)
        
    '--------------------------------------------------------------------------------------------------------------
    If Not mclsRunScript.OpenFile(strFile) Or mclsRunScript.EOF Then Exit Function
            
    '--------------------------------------------------------------------------------------------------------------
'    If mblnDBAUser Then
        mstrSQL = "select TABLE_NAME,COLUMN_NAME,DATA_TYPE from ALL_tab_columns where OWNER In ('" & mstrOwner & "','ZLTOOLS')"
'    Else
'        mstrSQL = "select TABLE_NAME,COLUMN_NAME,DATA_TYPE from user_tab_columns"
'    End If
    If rsObjects.State = adStateOpen Then rsObjects.Close
    rsObjects.Open mstrSQL, gcnOracle, adOpenKeyset
    Set objCollectionTableFieldType = New Collection
    Do While Not rsObjects.EOF
        objCollectionTableFieldType.Add rsObjects("DATA_TYPE").value, "K" & UCase(rsObjects("TABLE_NAME").value) & "'" & UCase(rsObjects("COLUMN_NAME").value)
        rsObjects.MoveNext
    Loop
    
    '读取所有过程的参数
    '--------------------------------------------------------------------------------------------------------------
    'Nvl(Type_Name, Data_Type),自定义类型以及对象类型，名字存放在Type_Name中，Data_Type标识整体类型
    'PACKAGE_NAME Is Null=除去包参数
    'DATA_LEVEL=0 ,参数类型名，而不是类型所属类型
    mstrSQL = "Select Object_Name, NVL(Argument_Name,'[返回值]') Argument_Name, Position, Sequence, Nvl(Type_Name, Data_Type) Data_Type, Default_Value, In_Out," & vbNewLine & _
                "       Data_Length, Data_Precision, Data_Scale" & vbNewLine & _
                "From All_Arguments A" & vbNewLine & _
                "Where Owner = '" & mstrOwner & "'  And PACKAGE_NAME Is Null And DATA_LEVEL=0 " & vbNewLine & _
                "Order By Object_Name, Decode(A.Argument_Name,NUll,999,Sequence),Position"
    
'    mstrSQL = "" & _
'    "   Select object_name,argument_name,Position,Sequence,data_type,default_value,In_Out,Data_length,Data_Precision,Data_scale " & _
'    "   From all_arguments  " & _
'    "   Where owner='" & mstrOwner & "' " & _
'    "   Order By object_name,Position"
    If rsObjects.State = adStateOpen Then rsObjects.Close
    rsObjects.Open mstrSQL, gcnOracle, adOpenKeyset
    
    strTemp = "": strSvrProcedureName = ""
    Set objCollectionAllDBPara = New Collection
    Do While Not rsObjects.EOF
        If strSvrProcedureName <> UCase(rsObjects("object_name").value) Then
            If strTemp <> "" And strSvrProcedureName <> "" Then
                objCollectionAllDBPara.Add strTemp, "K" & UCase(strSvrProcedureName)
            End If
            strSvrProcedureName = UCase(rsObjects("object_name").value)
            strTemp = rsObjects("argument_name").value & "'" & rsObjects("data_type").value & "'" & rsObjects("default_value").value & "'" & rsObjects("In_Out").value
        Else
            strTemp = strTemp & "`" & rsObjects("argument_name").value & "'" & rsObjects("data_type").value & "'" & rsObjects("default_value").value & "'" & rsObjects("In_Out").value
        End If
        rsObjects.MoveNext
    Loop
    If strTemp <> "" And strSvrProcedureName <> "" Then
        objCollectionAllDBPara.Add strTemp, "K" & UCase(strSvrProcedureName)
    End If
    
    '读取过程清单
    '--------------------------------------------------------------------------------------------------------------

    If mblnDBAUser Then
        mstrSQL = "select OBJECT_TYPE,OBJECT_NAME,STATUS from DBA_OBJECTS where OBJECT_TYPE in('PROCEDURE','FUNCTION') and OWNER='" & mstrOwner & "'"
    Else
        mstrSQL = "select OBJECT_TYPE,OBJECT_NAME,STATUS from USER_OBJECTS where OBJECT_TYPE in('PROCEDURE','FUNCTION')"
    End If
    If rsObjects.State = adStateOpen Then rsObjects.Close
    rsObjects.Open mstrSQL, gcnOracle, adOpenKeyset
    Call objClsProcedure.LoadData(rsObjects, Array("OBJECT_TYPE", "STATUS"), Array("OBJECT_NAME"))
    
    '--------------------------------------------------------------------------------------------------------------
    err = 0

    On Error GoTo errHand
    Do While Not mclsRunScript.EOF
        strSQL1 = mclsRunScript.SQLInfo.SQL
        strObjName = mclsRunScript.SQLInfo.BlockName
        strType = mclsRunScript.SQLInfo.BlockType
        '------------------------------------------------------------------------------------------------------
        If InStr(1, strType, "PROCEDURE") > 0 Then

            If objClsProcedure.Item(Array(strObjName)) = False Then
                If blnAddToRsTable Then
                    '1-存在对象,2-不存在对象,3-失效
                     Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "过程/函数", 2, False, strSQL1, "不存在", "严重：部分功能不能正常运行")
                Else
                    Call InputErrRpt("过程/函数", strObjName, "不存在", "严重：部分功能不能正常运行")
                End If

            ElseIf objClsProcedure.ItemValue("STATUS") <> "VALID" Then
                If blnAddToRsTable Then
                    '1-存在对象,2-不存在对象,3-失效
                     Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "过程/函数", 3, False, strSQL1, "处于无效状态", "严重：部分功能不能正常运行")
                Else
                    Call InputErrRpt("过程/函数", strObjName, "处于无效状态", "严重：部分功能不能正常运行")
                End If
            Else
                
'                    If UCase(strObjName) = UCase("Zl_病人结帐记录_Reprint") Then
'                        strObjName = "Zl_病人结帐记录_Reprint"
'                    End If

                Set objCollectionScriptPara = New Collection
                If AnsyProcedurePara(strObjName, strSQL1, objCollectionScriptPara, objCollectionTableFieldType) = True Then
                    
                    Set objCollectionDBPara = New Collection
                    On Error Resume Next
                    strTemp = ""
                    strTemp = objCollectionAllDBPara.Item("K" & UCase(strObjName))
                    On Error GoTo errHand
                    If strTemp <> "" Then
                        varTemp = Split(strTemp, "`")
                        For intLoop = 0 To UBound(varTemp)
                            objCollectionDBPara.Add varTemp(intLoop), "K" & UCase(strObjName) & "'" & Split(varTemp(intLoop), "'")(0)
                            On Error GoTo errHand
                        Next
                    End If
                    Call CompareProcedurePara(strSQL1, strObjName, objCollectionDBPara, objCollectionScriptPara, blnAddToRsTable)
                End If
            End If
        '------------------------------------------------------------------------------------------------------
        ElseIf InStr(1, strType, "FUNCTION") > 0 Then
            If objClsProcedure.Item(Array(strObjName)) = False Then
                If blnAddToRsTable Then
                    '1-存在对象,2-不存在对象,3-失效
                     Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "过程/函数", 2, False, strSQL1, "不存在", "严重：部分功能不能正常运行")
                Else
                    Call InputErrRpt("过程/函数", strObjName, "不存在", "严重：部分功能不能正常运行")
                End If

            ElseIf objClsProcedure.ItemValue("STATUS") <> "VALID" Then
                If blnAddToRsTable Then
                    '1-存在对象,2-不存在对象,3-失效
                     Call zlInsertRecData(mrsErrTable, strObjName, strObjName, "过程/函数", 3, False, strSQL1, "处于无效状态", "严重：部分功能不能正常运行")
                Else
                    Call InputErrRpt("过程/函数", strObjName, "处于无效状态", "严重：部分功能不能正常运行")
                End If
            Else
                Set objCollectionScriptPara = New Collection
                If AnsyProcedurePara(strObjName, strSQL1, objCollectionScriptPara, objCollectionTableFieldType) = True Then
                    Set objCollectionDBPara = New Collection
                    On Error Resume Next
                    strTemp = ""
                    strTemp = objCollectionAllDBPara.Item("K" & UCase(strObjName))
                    On Error GoTo errHand
                    If strTemp <> "" Then
                        varTemp = Split(strTemp, "`")
                        For intLoop = 0 To UBound(varTemp)
                            objCollectionDBPara.Add varTemp(intLoop), "K" & UCase(strObjName) & "'" & Split(varTemp(intLoop), "'")(0)
                            On Error GoTo errHand
                        Next
                        Call CompareProcedurePara(strSQL1, strObjName, objCollectionDBPara, objCollectionScriptPara, blnAddToRsTable)
                    End If
                End If
            End If
        End If
        intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
        If intSvrPercent <> intPercent Then
            intSvrPercent = intPercent
            RaiseEvent Progressing(intSvrPercent)
        End If
        Call mclsRunScript.ReadNextSQL
    Loop

    RaiseEvent AfterProgress
    
    CheckProgram = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description
'    Resume
End Function

Private Function CheckPackage(ByVal strFile As String) As Boolean
    '******************************************************************************************************************
    '功能：检查包对象，判断是否存在
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim rsObjects As New ADODB.Recordset
    Dim rsColumns As New ADODB.Recordset
    Dim lngCount As Long
    Dim strTemp As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim arySql() As String, strObjName As String
    
    On Error GoTo errHand
    
    With rsObjects
    
        RaiseEvent BeforeProgress("正在检查包", 100)
        If Not gobjFile.FileExists(strFile) Then Exit Function
        If Not mclsRunScript.OpenFile(strFile) Or mclsRunScript.EOF Then Exit Function
        
        If mblnDBAUser Then
            mstrSQL = "Select OBJECT_TYPE,OBJECT_NAME,STATUS from DBA_OBJECTS where OBJECT_TYPE in('PACKAGE') and OWNER='" & mstrOwner & "'"
        Else
            mstrSQL = "Select OBJECT_TYPE,OBJECT_NAME,STATUS from USER_OBJECTS where OBJECT_TYPE in('PACKAGE')"
        End If
        If .State = adStateOpen Then .Close
        .Open mstrSQL, gcnOracle, adOpenKeyset

        Do While Not mclsRunScript.EOF
            strObjName = mclsRunScript.SQLInfo.BlockName
            If InStr(1, mclsRunScript.SQLInfo.BlockType, "PACKAGE BODY") > 0 Then
                .Filter = "OBJECT_NAME='" & strObjName & "'"
                If .EOF Then
                    Call InputErrRpt("包", strObjName, "不存在", "严重：部分功能不能正常运行")
                ElseIf .Fields("STATUS").value <> "VALID" Then
                    Call InputErrRpt("包", strObjName, "处于无效状态", "严重：部分功能不能正常运行")
                End If
            End If
            intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
            Call mclsRunScript.ReadNextSQL
        Loop
    End With
    RaiseEvent AfterProgress
    
    CheckPackage = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Sub CheckMustData(ByVal cnOracle As ADODB.Connection, ByVal strInitPath As String, ByVal lngSys As Long, ByRef objClsBakTable As clsCollectionRecordset)
    '------------------------------------------------------------------------------------------------------------------
    '功能:检查数据安装的必需数据:包含如下:zlprogFuncs和zlprogPrivs,zlProgRelas
    '     需执行两个脚本:zlmandata.sql和zlreports.sql
    '参数:cnOracle-检查的数据库
    '     strInitPath-安装脚本路径
    '     lngSys-系统号
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim strFile As String, blnHaveReport As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim lngPostion As Long
    Dim strLine As String
    Dim lngSysCode As Long
    Dim lngMoudal As Long
    Dim varTemp As Variant
    Dim strFunc As String
    Dim strOwner As String
    Dim strObject As String
    Dim strObjectPrivs As String
    Dim strFieldSerial As String
    Dim intLoop As Integer
    Dim intPosSysCode As Integer
    Dim intPosModual As Integer
    Dim intPosFunc As Integer
    Dim intPosObject As Integer
    Dim intPosObjectPriv As Integer
    Dim objClsProgram As New clsCollectionRecordset
    Dim objClsFunc As New clsCollectionRecordset
    Dim objClsPriv As New clsCollectionRecordset
    
    On Error GoTo errHand
    
                
    strFile = strInitPath & "zlManData.sql"
    If mobjFile.FileExists(strFile) = False Then
        RaiseEvent Exception
        Exit Sub
    End If
    
    RaiseEvent BeforeProgress("正在检查权限", 100)
    
    If Not mclsRunScript.OpenFile(strFile) Or mclsRunScript.EOF Then Exit Sub
    
    
    gstrSQL = "Select 序号,标题 From zlPrograms a Where a.系统=" & lngSys
    Call OpenRecordset(rsTemp, gstrSQL, mfrmMain.Caption, , , cnOracle)
    Call objClsProgram.LoadData(rsTemp, Array("标题"), Array("序号"))
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select Distinct a.系统,a.序号,Upper(a.功能) As 功能," & _
                "'功能权限' as A1,c.标题||'：'||a.功能  as A2,'不存在此权限' as A3 ,'严重：不能正常授权(zlProgFuncs)' as A4 " & _
                "From zlProgFuncs a ,zlprograms c Where a.系统=c.系统 And a.序号=c.序号 And a.系统=" & lngSys
        
    Call OpenRecordset(rsTemp, gstrSQL, mfrmMain.Caption, , , cnOracle)
    Call objClsFunc.LoadData(rsTemp, Array(), Array("系统", "序号", "功能"))
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select Distinct a.系统,a.序号,Upper(a.功能) As 功能,Upper(a.对象) As 对象,Upper(a.权限) As 权限," & _
                "'功能权限' as A1,c.标题||'\'||Upper(a.功能)||'\'||Upper(a.对象)||'：'||Upper(a.权限)  as A2,'不存在此权限' as A3 ,'严重：可能应用程序不能正常访问此权限(zlProgPrivs)' as A4 " & _
                "From zlProgPrivs a ,zlprograms c Where a.系统=c.系统 And a.序号=c.序号 and a.系统=" & lngSys & " Order by a.系统,a.序号"
    
    Call OpenRecordset(rsTemp, gstrSQL, mfrmMain.Caption, , , cnOracle)
    Call objClsPriv.LoadData(rsTemp, Array(), Array("系统", "序号", "功能", "对象", "权限"))
        
    lngRow = 0

    '------------------------------------------------------------------------------------------------------------------
    Do While Not mclsRunScript.EOF
'        If mclsRunScript.Line >= 214 Then Stop
        Call CheckMustDataItem(mclsRunScript.SQLInfo.SQL, objClsProgram, objClsFunc, objClsPriv, objClsBakTable)

        intPercent = 100 * mclsRunScript.Line / mclsRunScript.LinesCount
        If intSvrPercent <> intPercent Then
            intSvrPercent = intPercent
            RaiseEvent Progressing(intSvrPercent)
        End If
        Call mclsRunScript.ReadNextSQL
    Loop

    RaiseEvent AfterProgress
    
    Set objClsFunc = Nothing
    Set objClsPriv = Nothing
    
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
    Resume
End Sub

Private Sub CheckMustDataItem(ByVal strPrivilegeSQL As String, ByVal objClsProgram As clsCollectionRecordset, ByVal objClsFunc As clsCollectionRecordset, ByVal objClsPriv As clsCollectionRecordset, ByVal objClsBakTable As clsCollectionRecordset)
    Dim strOwner As String
    Dim strObject As String
    Dim strObjectPrivs As String
    Dim strTemp As String
    Dim strLine As String
    Dim lngPostion As Long
    Dim strFieldSerial As String
    Dim varTemp As Variant
    Dim intLoop As Integer
    Dim intPosSysCode As Integer
    Dim intPosModual As Integer
    Dim intPosFunc As Integer
    Dim intPosTitle As Integer
    Dim intPosObject As Integer
    Dim intPosObjectPriv As Integer
    Dim lngSysCode As Long
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim strObjectBak As String
    Dim rsTemp As New ADODB.Recordset
    Dim varLine As Variant
    
    On Error GoTo errHand
    
    strObject = ""
    strObjectPrivs = ""
 
    strPrivilegeSQL = UCase(TrimEx(strPrivilegeSQL))
    strTemp = Mid(Replace(strPrivilegeSQL, " ", ""), 1, 80)
    strTemp = Replace(strTemp, vbCrLf, "")
    strTemp = Replace(strTemp, Chr(10), "")
    strTemp = Replace(strTemp, Chr(13), "")
    '----------------------------------------------------------------------------------------------------------
    
    If strTemp Like "INSERTINTOZLBAKTABLES(*" Then 'UCase("InsertIntozlBakTables(*")
        
        strLine = strPrivilegeSQL
                        
        varLine = Split(strLine, vbCrLf)
        For intLoop = 0 To UBound(varLine)
            
            If intLoop = 0 Then
                strLine = "Select 0 As 系统,0 as 组号,'' As 表名,0 as 序号,0 as 直接转出,0 停用触发器 From Dual where 1=0 " & vbCrLf
                strLine = strLine & "Union All "
            Else
                strLine = strLine & varLine(intLoop) & vbCrLf
            End If
        Next
        
        rsTemp.Open strLine, gcnOracle
        Call objClsBakTable.LoadData(rsTemp, Array(), Array("表名"))
        
        strLine = ""
    '----------------------------------------------------------------------------------------------------------
    ElseIf strTemp Like "INSERTINTOZLPROGFUNCS(*" Then 'UCase("InsertIntozlProgFuncs(*")

        '分解出系统，模块号，功能
        strLine = strPrivilegeSQL
        lngPostion = InStr(strLine, "VALUES")
        If lngPostion > 0 Then

            strFieldSerial = UCase(Trim(Mid(strLine, 1, lngPostion - 1)))
            strLine = Trim(Mid(strLine, lngPostion + 6))

            lngPostion = InStr(strFieldSerial, "ZLPROGFUNCS") 'UCase("zlProgFuncs")
            If lngPostion > 0 Then
                strFieldSerial = Trim(Mid(strFieldSerial, lngPostion + 11))
                If Left(strFieldSerial, 1) = "(" Then strFieldSerial = Mid(strFieldSerial, 2)
                If Right(strFieldSerial, 1) = ")" Then strFieldSerial = Mid(strFieldSerial, 1, Len(strFieldSerial) - 1)
                varTemp = Split(strFieldSerial, ",")
                For intLoop = 0 To UBound(varTemp)
                    Select Case Trim(varTemp(intLoop))
                    Case "系统"
                        intPosSysCode = intLoop
                    Case "序号"
                        intPosModual = intLoop
                    Case "功能"
                        intPosFunc = intLoop
                    End Select
                Next
            End If
            
            If Left(strLine, 1) = "(" Then strLine = Mid(strLine, 2)
            If Right(strLine, 1) = ")" Then strLine = Mid(strLine, 1, Len(strLine) - 1)

            varTemp = Split(strLine, ",")
            lngSysCode = Val(varTemp(intPosSysCode))
            lngMoudal = Val(varTemp(intPosModual))
            strFunc = Trim(varTemp(intPosFunc))
        Else
            strLine = ""
        End If
    '----------------------------------------------------------------------------------------------------------
    ElseIf strTemp Like "INSERTINTOZLPROGPRIVS(*" Then 'UCase("InsertIntozlProgPrivs(*")

        '分解出系统，模块号，功能，所有者，对象，权限

        strLine = strPrivilegeSQL
        lngPostion = InStr(UCase(strLine), "VALUES")
        If lngPostion > 0 Then
            strFieldSerial = UCase(Trim(Mid(strLine, 1, lngPostion - 1)))
            strLine = Trim(Mid(strLine, lngPostion + 6))

            lngPostion = InStr(strFieldSerial, "ZLPROGPRIVS") 'UCase("zlProgPrivs")
            If lngPostion > 0 Then
                strFieldSerial = Trim(Mid(strFieldSerial, lngPostion + 11))
                If Left(strFieldSerial, 1) = "(" Then strFieldSerial = Mid(strFieldSerial, 2)
                If Right(strFieldSerial, 1) = ")" Then strFieldSerial = Mid(strFieldSerial, 1, Len(strFieldSerial) - 1)
                varTemp = Split(strFieldSerial, ",")
                For intLoop = 0 To UBound(varTemp)
                    Select Case Trim(varTemp(intLoop))
                    Case "系统"
                        intPosSysCode = intLoop
                    Case "序号"
                        intPosModual = intLoop
                    Case "功能"
                        intPosFunc = intLoop
                    Case "对象"
                        intPosObject = intLoop
                    Case "权限"
                        intPosObjectPriv = intLoop
                    End Select
                Next
            End If

            If Left(strLine, 1) = "(" Then strLine = Mid(strLine, 2)
            If Right(strLine, 1) = ")" Then strLine = Mid(strLine, 1, Len(strLine) - 1)

            varTemp = Split(strLine, ",")
            lngSysCode = Val(varTemp(intPosSysCode))
            lngMoudal = Val(varTemp(intPosModual))
            strFunc = UCase(Trim(varTemp(intPosFunc)))
            strOwner = mstrOwner
            strObject = UCase(Trim(varTemp(intPosObject)))
            strObjectPrivs = UCase(Trim(varTemp(intPosObjectPriv)))
            strObjectBak = ""
            
            If strObject <> "" Then strObject = Mid(strObject, 2, Len(strObject) - 2)
            If strObjectPrivs <> "" Then strObjectPrivs = Mid(strObjectPrivs, 2, Len(strObjectPrivs) - 2)
            If UCase(strObjectPrivs) = "SELECT" Then
                If objClsBakTable.Item(Array(strObject)) = True Then
                    strObjectBak = "H" & strObject
                End If
            End If
'                    If lngMoudal = 1000 Then
'                        lngMoudal = 1000
'                    End If
        Else
            strLine = ""
        End If

    Else
        strLine = ""
    End If


    '----------------------------------------------------------------------------------------------------------
    If strLine <> "" Then

        strFunc = Mid(strFunc, 2, Len(strFunc) - 2)
        If strObject = "" Then

'                   If lngMoudal = 1539 Then
'                        lngMoudal = 1539
'                    End If

            If objClsProgram.Item(Array(lngMoudal)) = True Then
                If objClsFunc.Item(Array(lngSysCode, lngMoudal, strFunc)) = False Then
                    Call InputErrRpt("功能权限", objClsProgram.ItemValue("标题") & "：" & strFunc, "不存在此权限", "严重：不能正常授权(zlProgFuncs)")
                End If
            End If

        Else

           If lngMoudal = 1539 Then
                lngMoudal = 1539
            End If

            If objClsProgram.Item(Array(lngMoudal)) = True Then
            
                If objClsPriv.Item(Array(lngSysCode, lngMoudal, strFunc, strObject, strObjectPrivs)) = False Then
                    Call InputErrRpt("功能权限", objClsProgram.ItemValue("标题") & "\" & strFunc & "\" & strObject & "：" & strObjectPrivs, "不存在此权限", "严重：可能应用程序不能正常访问此权限(zlProgPrivs)")
                End If
                    
                If strObjectBak <> "" And lngMoudal <> 1000 Then
                    If objClsPriv.Item(Array(lngSysCode, lngMoudal, strFunc, strObjectBak, strObjectPrivs)) = False Then
                        Call InputErrRpt("功能权限", objClsProgram.ItemValue("标题") & "\" & strFunc & "\" & strObjectBak & "：" & strObjectPrivs, "不存在此权限", "严重：可能应用程序不能正常访问此权限(zlProgPrivs)")
                    End If
                End If
            End If

        End If

    End If
    
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Sub

Private Sub InputErrRpt(ObjType As String, ObjName As String, ErrInfo As String, Optional Advice As String)
    '----------------------------------------------------
    '填写一行错误报告
    '----------------------------------------------------
    With frmAppChkRpt.hgdReport
        .Rows = .Rows + 1
        If .Tag <> ObjType Then
            .TextMatrix(.Rows - 1, 0) = "------< " & ObjType & "检查情况 >------"
            .TextMatrix(.Rows - 1, 1) = "------< " & ObjType & "检查情况 >------"
            .MergeRow(.Rows - 1) = True
            .Rows = .Rows + 1
        End If
        .Tag = ObjType
        .TextMatrix(.Rows - 1, 0) = ObjName
        If .ColData(0) < mfrmMain.TextWidth(ObjName) Then
            .ColData(0) = mfrmMain.TextWidth(ObjName)
        End If
        .TextMatrix(.Rows - 1, 1) = ErrInfo
        If .ColData(1) < mfrmMain.TextWidth(ErrInfo) Then
            .ColData(1) = mfrmMain.TextWidth(ErrInfo)
        End If
        .TextMatrix(.Rows - 1, 2) = Advice
        If .ColData(2) < mfrmMain.TextWidth(Advice) Then
            .ColData(2) = mfrmMain.TextWidth(Advice)
        End If
            
    End With
        
    '显示提醒标签
    If InStr(Advice, "严重") > 0 Or InStr(Advice, "较重") > 0 Then
         frmAppChkRpt.lblWarn.Visible = True
    End If
End Sub

Private Function zl获取关联外键(ByVal rsTemp As ADODB.Recordset, ByVal str对象名 As String, ByRef rsData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定主键的关联外键
    '入参:str对象名-对象名称
    '出参:rsData
    '返回:
    '编制:刘兴洪
    '日期:2009-08-20 11:30:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTemp As String
        
    With rsTemp
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        
        .Filter = "R_Constraint_Name='" & str对象名 & "'"
        If .RecordCount = 0 Then
            zl获取关联外键 = True
            Exit Function
        End If
        
        .Filter = "R_Constraint_Name='" & str对象名 & "' And OWNER<>R_OWNER AND R_OWNER='" & mstrOwner & "'"
                
        If .RecordCount <> 0 Then
            '存在其他系统关联,当项主键只能手工更新
            If Not rsData.EOF Then
                If UCase(Nvl(rsData!对象名称)) <> UCase(str对象名) Then
                    rsData.Filter = "对象名称='" & UCase(str对象名) & "'"
                End If
            Else
                rsData.Filter = "对象名称='" & UCase(str对象名) & "'"
            End If
            
            If rsData.EOF Then
                rsData.Filter = 0
                zl获取关联外键 = True
                Exit Function
            End If
            
            .MoveFirst
            strTemp = ""
            Do While Not .EOF
                strTemp = strTemp & "," & Nvl(!Table_Name) & "(" & Nvl(!Owner) & ")"
                .MoveNext
            Loop
            
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            If strTemp <> "" Then strTemp = Substr("被其他系统关联，不能有效修正！如表:" & strTemp, 1, 500)
            '更新标志
            rsData!修正标志 = 4
            rsData!修正说明 = strTemp
            rsData.Update
            .Close
            zl获取关联外键 = True
            Exit Function
        End If
        
        .Filter = 0
        .MoveFirst
        
        '需要删除相关的级联
        Do While Not .EOF
            '标志: '1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致,8-处理主键时，需要先处理外键
            Call zlInsertRecData(rsData, Nvl(!Table_Name), Nvl(!Constraint_Name), "外键", 8, False, "", "级联处理", "在处理主键或唯一键时，需要先处理外键！")
            .MoveNext
        Loop
        .Close
        zl获取关联外键 = True
        Exit Function
    End With
End Function

Private Function AnsyProcedurePara(ByVal strProcedureName As String, ByVal strProcedureSQL As String, ByRef objCollectionPara As Collection, ByRef objCollectionTableFieldType As Collection) As Boolean
    '******************************************************************************************************************
    '功能：检查
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    
    Dim lngLoop As Long
    Dim lngCount As Long
    Dim strProcedurePara As String
    Dim strLine As String
    Dim varLine As Variant
    Dim strTemp As String
    Dim varWord As Variant
    Dim intLoop As Integer
    Dim lngPostion As Long
    Dim lngPostion2 As Long
    Dim strProcedureParaTypeTable As String
    Dim strProcedureParaTypeField As String
    
    Dim strProcedureParaName As String
    Dim strProcedureParaType As String
    Dim strProcedureParaDefault As String
    Dim strProcedureParaOut As String
    Dim intStage As Integer '-1=字段前，0-解析字段，1-字段解析完毕
    Dim strReturn As String, blnReturn As Boolean
    On Error GoTo errHand

'    If strProcedureName = "ZL_电子病历时机_INSERT" Then
'        strProcedureName = "ZL_电子病历时机_INSERT"
'    End If
        
    '先提取出参数块脚本
    '------------------------------------------------------------------------------------------------------------------
    varLine = Split(strProcedureSQL, vbCrLf)
    lngCount = UBound(varLine)
    intStage = -1
    For lngLoop = 0 To lngCount
        
        strLine = Trim(varLine(lngLoop))
        strLine = Replace(strLine, vbTab, " ")
        '取掉--注释
        lngPostion = InStr(strLine, "--")
        If lngPostion > 0 Then strLine = Mid(strLine, 1, lngPostion - 1)
        If UCase(strLine) Like "* RETURN *" Or UCase(strLine) Like "*)RETURN *" Or UCase(strLine) Like "*RETURN *" Then
            lngPostion = InStr(UCase(strLine), "RETURN")
            If lngPostion > 1 Then
                strLine = Trim(Mid(strLine, 1, lngPostion - 2) & " ,[返回值] " & Mid(strLine, lngPostion + Len("RETURN")))
            Else
                strLine = ",[返回值] " & Mid(strLine, lngPostion + Len("RETURN"))
            End If
        End If
        If strLine Like ")*" Then strLine = Trim(Mid(strLine, 2))
        If strLine Like "*)" Then
            If Len(Replace(strLine, "(", "")) > Len(Replace(strLine, ")", "")) Then
                strLine = Trim(Mid(strLine, 1, Len(strLine) - 1))
            End If
        End If
        strTemp = ""
        If Trim(strLine) <> "" And Trim(strLine) <> "(" And Trim(strLine) <> ")" Then
            varWord = Split(UCase(strLine), " ")
            For intLoop = 0 To UBound(varWord)
                Select Case UCase(varWord(intLoop))
                    Case "IS", ")IS", "AS", ")AS", "PIPELINED"
                    '找到函数或存储过程的参数结尾处
                        intStage = 1: Exit For
                    Case UCase(strProcedureName), UCase(strProcedureName) & "("
                        '过程名，过程名(
                        intStage = 0
                    Case Else
                        If varWord(intLoop) <> "" And intStage = 0 Then
                            If varWord(intLoop) = ",[返回值]" Then blnReturn = True
                            If blnReturn Then
                                strReturn = strReturn & " " & varWord(intLoop)
                            Else
                                If Left(UCase(varWord(intLoop)), Len(strProcedureName)) = UCase(strProcedureName) Then
                                    strTemp = strTemp & " " & Mid(varWord(intLoop), Len(strProcedureName) + 1)
                                Else
                                    strTemp = strTemp & " " & varWord(intLoop)
                                End If
                            End If
                        End If
                End Select
            Next
        End If
        If strTemp <> "" Then strProcedurePara = strProcedurePara & strTemp & vbCrLf
        If intStage = 1 Then Exit For
    Next

    '取消/*注释
    lngPostion = InStr(strProcedurePara, "/*")
    Do While lngPostion > 0
        lngPostion2 = InStr(strProcedurePara, "*/")
        If lngPostion2 <= 0 Then Exit Do
        strProcedurePara = Mid(strProcedurePara, 1, lngPostion - 1) & Mid(strProcedurePara, lngPostion2 + 2)
        lngPostion = InStr(strProcedurePara, "/*")
    Loop
    strProcedurePara = Trim(Replace(strProcedurePara, vbTab, " "))
    strProcedurePara = Trim(Replace(strProcedurePara, vbCrLf, ""))
    If Trim(strProcedurePara) = "" And strReturn <> "" Then
        strReturn = Mid(Trim(strReturn), 2)
    End If
    If Left(strProcedurePara, 1) = "(" Then strProcedurePara = Trim(Mid(strProcedurePara, 2))
    If Right(strProcedurePara, 1) = ")" Then strProcedurePara = Trim(Mid(strProcedurePara, 1, Len(strProcedurePara) - 1))
    If Left(strReturn, 1) = "(" Then strReturn = Trim(Mid(strReturn, 2))
    If Right(strReturn, 1) = ")" Then strReturn = Trim(Mid(strReturn, 1, Len(strReturn) - 1))
    If strReturn <> "" Then
        strProcedurePara = strProcedurePara & strReturn
    End If
        
    '开始正式分析参数
    '------------------------------------------------------------------------------------------------------------------
    If strProcedurePara <> "" Then
        varLine = Split(strProcedurePara, ",")
        lngCount = UBound(varLine)
        
        For lngLoop = 0 To lngCount
            
            strLine = Trim(varLine(lngLoop))
            strLine = Trim(Replace(strLine, vbTab, " "))
            varWord = Split(strLine, " ")
            
            strProcedureParaName = ""
            strProcedureParaType = ""
            strProcedureParaDefault = ""
            strProcedureParaOut = ""
            strTemp = ""
            
            For intLoop = 0 To UBound(varWord)
                If varWord(intLoop) <> "" Then
                    If strProcedureParaName = "" Then
                        strProcedureParaName = varWord(intLoop)
                    ElseIf UCase(varWord(intLoop)) = "IN" Then
                        If strProcedureParaOut = "OUT" Then
                            strProcedureParaOut = "IN/OUT"
                        Else
                            strProcedureParaOut = "IN"
                        End If
                    ElseIf UCase(varWord(intLoop)) = "OUT" Then
                        If strProcedureParaOut = "IN" Then
                            strProcedureParaOut = "IN/OUT"
                        Else
                             strProcedureParaOut = "OUT"
                        End If
                    Else
                        strTemp = strTemp & " " & varWord(intLoop)
                    End If
                End If
            Next
            If strProcedureParaOut = "" Then
                If strProcedureParaName = "[返回值]" Then
                    strProcedureParaOut = "OUT"
                Else
                    strProcedureParaOut = "IN"
                End If
            End If
            lngPostion = InStr(strTemp, ":")
            If lngPostion > 0 Then
                '有缺省值
                strProcedureParaType = Trim(Mid(strTemp, 1, lngPostion - 1))
                strTemp = Trim(Mid(strTemp, lngPostion + 1))
                lngPostion = InStr(strTemp, "=")
                If lngPostion > 0 Then
                    strProcedureParaDefault = Trim(Mid(strTemp, lngPostion + 1))
                End If
            Else
                strProcedureParaType = strTemp
            End If
            
            lngPostion = InStr(strProcedureParaType, "%")
            If lngPostion > 0 Then
                strProcedureParaType = Mid(strProcedureParaType, 1, lngPostion - 1)
                lngPostion = InStr(strProcedureParaType, ".")
                If lngPostion > 0 Then
                    If InStr(lngPostion + 1, strProcedureParaType, ".") > 0 Then
                        '去掉对象所有者
                        strProcedureParaType = Trim(Mid(strProcedureParaType, lngPostion + 1))
                        lngPostion = InStr(strProcedureParaType, ".")
                    End If
                    strProcedureParaTypeTable = Trim(Mid(strProcedureParaType, 1, lngPostion - 1))
                    strProcedureParaTypeField = Trim(Mid(strProcedureParaType, lngPostion + 1))
                    
                    On Error Resume Next
                    err.Clear
                    err = 0
                    strProcedureParaType = Trim(UCase(objCollectionTableFieldType.Item("K" & UCase(strProcedureParaTypeTable) & "'" & UCase(strProcedureParaTypeField))))
                    On Error GoTo errHand
                End If
            End If
        
            Select Case UCase(Trim(strProcedureParaType))
                Case "INTEGER"
                    strProcedureParaType = "NUMBER"
                End Select
        
            objCollectionPara.Add strProcedureParaName & mstrSplit & strProcedureParaType & mstrSplit & strProcedureParaDefault & mstrSplit & strProcedureParaOut, "K" & strProcedureParaName
        Next
        AnsyProcedurePara = True
    Else
         AnsyProcedurePara = False
    End If

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function CompareProcedurePara(ByVal strExcuteSQL As String, ByVal strObjectName As String, ByVal objCollectionDB As Collection, ByVal objCollectionScript As Collection, Optional blnAddRec As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：检查
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************

    Dim strTemp As String
    Dim varParaDB As Variant
    Dim varParaScript As Variant
    Dim intLoop As Integer
    Dim strDBRetrn As String, strScriptReturn As String
    
'    If UCase(strObjectName) = UCase("Zl_挂号安排_Delete") Then
'        strObjectName = "Zl_挂号安排_Delete"
'    End If
    '取返回值，在该过程结束时检查
    On Error Resume Next
        strDBRetrn = objCollectionDB.Item("K" & UCase(strObjectName) & "'[返回值]")
        If err.Number <> 0 Then
            strDBRetrn = "[没有返回值]": err.Clear
        Else '删掉返回值
            objCollectionDB.Remove ("K" & UCase(strObjectName) & "'[返回值]")
        End If
        strScriptReturn = objCollectionScript.Item("K[返回值]")
        If err.Number <> 0 Then
            strScriptReturn = "[没有返回值]": err.Clear
        Else '删掉返回值
            objCollectionScript.Remove ("K[返回值]")
        End If
    On Error GoTo errHand
    '检查参数个数是否正确
    '--------------------------------------------------------------------------------------------------------------
    If objCollectionDB.Count <> objCollectionScript.Count Then
        '参数个数不对,返回
        If blnAddRec Then
            '1-存在对象,2-不存在对象,3-失效
             Call zlInsertRecData(mrsErrTable, strObjectName, strObjectName, "过程/函数", 1, False, strExcuteSQL, "参数个数不对", "严重：部分功能不能正常运行")
        Else
            Call InputErrRpt("过程/函数", strObjectName, "参数个数不对", "严重：部分功能不能正常运行")
        End If
        Exit Function
    End If
        
    For intLoop = 1 To objCollectionDB.Count
        
        '分析出参数名、类型、缺省值
        
        strTemp = objCollectionDB.Item(intLoop)
        varParaDB = Split(strTemp, "'")
        
        strTemp = objCollectionScript.Item(intLoop)
        varParaScript = Split(strTemp, mstrSplit)
                
        '检查参数名称是否正确
        '--------------------------------------------------------------------------------------------------------------
        If UCase(Trim(varParaDB(0))) <> UCase(Trim(varParaScript(0))) Then
            If blnAddRec Then
                '1-存在对象,2-不存在对象,3-失效
                 Call zlInsertRecData(mrsErrTable, strObjectName, strObjectName, "过程/函数", 1, False, strExcuteSQL, "参数名不对", "严重：可能造成数据混乱,原来为:“" & varParaDB(0) & "”现为:" & varParaScript(0))
            Else
                Call InputErrRpt("过程/函数", strObjectName, "参数名不对,原来为:“" & varParaDB(0) & "”现为:" & varParaScript(0), "严重：可能造成数据混乱")
            End If
            Exit Function
        End If
        
        '检查参数类型是否正确
        '--------------------------------------------------------------------------------------------------------------
        If UCase(Trim(varParaDB(1))) <> UCase(Trim(varParaScript(1))) Then
            If blnAddRec Then
                '1-存在对象,2-不存在对象,3-失效
                 Call zlInsertRecData(mrsErrTable, strObjectName, strObjectName, "过程/函数", 1, False, strExcuteSQL, "参数名不对", "严重：部分功能不能正常运行,“" & varParaDB(0) & "”原来为:" & varParaDB(1) & "现为:" & varParaScript(1))
            Else
                Call InputErrRpt("过程/函数", strObjectName, "参数类型不对,“" & varParaDB(0) & "”原来为:" & varParaDB(1) & "现为:" & varParaScript(1), "严重：部分功能不能正常运行")
            End If
            Exit Function
        End If
        
        '检查缺省值，由于oralce中的缺省值字段是Long型，所以暂时不能读出来（以前也是一样的，根本没有起作用）
        '--------------------------------------------------------------------------------------------------------------
        If UCase(Trim(varParaDB(2))) <> UCase(Trim(varParaScript(2))) Then
'            If blnAddRec Then
'                '1-存在对象,2-不存在对象,3-失效
'                 Call zlInsertRecData(mrsErrTable, strObjectName, strObjectName, "过程/函数", 1, False, strExcuteSQL, "缺省值不对", "严重：可能造成数据混乱,“" & varParaDB(0) & "”原来为:" & varParaDB(2) & "现为:" & varParaScript(2))
'            Else
'                Call InputErrRpt("过程/函数", strObjectName, "缺省值不对,“" & varParaDB(0) & "”原来为:" & varParaDB(2) & "现为:" & varParaScript(2), "严重：可能造成数据混乱")
'            End If
'            Exit Function
        End If
        
        '检查参数入出参
        '--------------------------------------------------------------------------------------------------------------
        If UCase(Trim(varParaDB(3))) <> UCase(Trim(varParaScript(3))) Then
            If blnAddRec Then
                '1-存在对象,2-不存在对象,3-失效
                 Call zlInsertRecData(mrsErrTable, strObjectName, strObjectName, "过程/函数", 1, False, strExcuteSQL, "入出参错误", "较轻：部分功能不能正常运行,“" & varParaDB(0) & "”原来为:" & varParaDB(3) & "现为:" & varParaScript(3))
            Else
                Call InputErrRpt("过程/函数", strObjectName, "入出参错误,“" & varParaDB(0) & "”原来为:" & varParaDB(3) & "现为:" & varParaScript(3), "较轻：部分功能不能正常运行")
            End If
            Exit Function
        End If
        
    Next
    
    If strDBRetrn <> Replace(strScriptReturn, mstrSplit, "'") Then
        varParaDB = Split(strDBRetrn, "'")
        varParaScript = Split(strScriptReturn, mstrSplit)
        If strDBRetrn = "" Then
            If blnAddRec Then
                '1-存在对象,2-不存在对象,3-失效
                 Call zlInsertRecData(mrsErrTable, strObjectName, strObjectName, "过程/函数", 1, False, strExcuteSQL, "返回值错误", "严重：部分功能不能正常运行，原来没有返回值，现为:" & varParaScript(1))
            Else
                Call InputErrRpt("过程/函数", strObjectName, "返回值类型错误,原来没有返回值，现为:" & varParaScript(1), "严重：部分功能不能正常运行")
            End If
            Exit Function
        ElseIf Replace(strScriptReturn, mstrSplit, "'") = "" Then
            If blnAddRec Then
                '1-存在对象,2-不存在对象,3-失效
                 Call zlInsertRecData(mrsErrTable, strObjectName, strObjectName, "过程/函数", 1, False, strExcuteSQL, "返回值错误", "严重：部分功能不能正常运行，现为没有返回值，原来:" & varParaDB(1))
            Else
                Call InputErrRpt("过程/函数", strObjectName, "返回值类型错误，现为没有返回值，原来:" & varParaDB(1), "严重：部分功能不能正常运行")
            End If
            Exit Function
        Else
            If UCase(Trim(varParaDB(1))) <> UCase(Trim(varParaScript(1))) Then
                If blnAddRec Then
                    '1-存在对象,2-不存在对象,3-失效
                     Call zlInsertRecData(mrsErrTable, strObjectName, strObjectName, "过程/函数", 1, False, strExcuteSQL, "返回值类型不对", "严重：部分功能不能正常运行,返回值原来为:" & varParaDB(1) & "现为:" & varParaScript(1))
                Else
                    Call InputErrRpt("过程/函数", strObjectName, "返回值类型错误,返回值原来为:" & varParaDB(1) & "现为:" & varParaScript(1), "严重：部分功能不能正常运行")
                End If
                Exit Function
            End If
        End If
    End If
    CompareProcedurePara = True
    Exit Function
errHand:
    MsgBox err.Description
End Function

Private Sub Class_Initialize()
    Set mobjFile = New FileSystemObject
    mstrSplit = "IAMSPLITESTRING"
End Sub

Private Sub Class_Terminate()
    Set mobjFile = Nothing
    Set mclsRunScript = Nothing
End Sub
