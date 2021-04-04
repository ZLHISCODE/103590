Attribute VB_Name = "mdlProc"
Option Explicit
'模块说明:处理过程类模块
'返回的过程集合定义为
'       "P_Name", adVarChar, 32 过程名称
'       "P_Define", adLongVarChar, 9999999#  过程定义(完整的过程文本)
'       "P_System", adVarChar, 20   系统名称
'       "P_SysNum", adInteger, 5 系统编号
'       "P_Owner", adVarChar, 20   系统所有者
'       "P_Ver", adVarChar, 20  脚本文件版本

Public gstrBCode As New clsStringBulider

Public Sub GetProceduresByFile(ByVal strFile As String, rsProcedure As ADODB.Recordset, _
                                            Optional ByVal strFileVer As String, Optional ByVal lngSysNum As Long, _
                                            Optional ByVal strSysName As String, Optional ByVal strOwner As String)
    '根据传入的文件名称,返回记录集
    '参数:strVer 文件对应版本
    Dim objTxt As TextStream
    Dim arrTxt() As String, dblRow As Double
    Dim strLine As String, strFMT As String
    Dim blnBegin As Boolean, strPName As String
    Dim arrDelete() As String, strProcName As String
    Dim i As Long
    
    On Error GoTo errH
    If Not gobjFile.FileExists(strFile) Then Exit Sub
    If gobjFile.GetFile(strFile).Size = 0 Then Exit Sub '文件为空
    
    If strFileVer = "" Then
        strFileVer = Mid(strFile, InStrRev(strFile, "\") + 1)
    End If
    If rsProcedure Is Nothing Then
        Set rsProcedure = New ADODB.Recordset
        With rsProcedure
            .Fields.Append "P_Name", adVarChar, 32 '过程名称
            .Fields.Append "P_Define", adLongVarChar, 9999999#  '过程定义
            
            .Fields.Append "P_System", adVarChar, 20   '系统名称
            .Fields.Append "P_SysNum", adInteger, 5 '系统编号
            .Fields.Append "P_Owner", adVarChar, 20   '系统所有者
            .Fields.Append "P_Ver", adVarChar, 50  '脚本文件版本
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    End If
    
    '一次将文本文件中的数据都读取出来,存在数组arrTxt中
    Set objTxt = gobjFile.OpenTextFile(strFile)
    arrTxt = Split(objTxt.ReadAll, vbNewLine)
    objTxt.Close
    
    gstrBCode.Clear
    '循环,将每一段的过程名称和过程定义保存到记录集中
    ReDim arrDelete(0)
    For dblRow = 0 To UBound(arrTxt)
        strLine = RTrim(arrTxt(dblRow))
        strFMT = UCase(TrimComment(TrimEx(strLine)))
        
        '如果该行含有Drop Procedure语句 ,就把过程名称记录下来,后续从记录中把该过程删除
        If InStr(1, strFMT, "DROP PROCEDURE") > 0 Then
            strProcName = Mid(strFMT, InStr(1, strFMT, "DROP PROCEDURE") + Len("DROP PROCEDURE "))  '截取
            strProcName = Split(strProcName, " ")(0)    '取第一个空格之前
            If InStr(1, strProcName, ".") > 0 Then strProcName = Split(strProcName, ".")(1) '判断是否有所有者
            If InStr(1, strProcName, ";") > 0 Then strProcName = Left(strProcName, Len(strProcName) - 1) '如果是分号结尾,应该去掉分号
            arrDelete(UBound(arrDelete)) = strProcName
            ReDim Preserve arrDelete(UBound(arrDelete) + 1)
        End If
        
        '开始记录过程
        If strFMT Like "CREATE*PROCEDURE *" Or strFMT Like "CREATE*FUNCTION *" Then
            strPName = Split(strFMT, " ")(4)
            If InStr(1, strPName, "(") > 0 Then strPName = Left(strPName, InStr(1, strPName, "(") - 1)
            If InStr(1, strPName, ".") > 0 Then strPName = Split(strPName, ".")(1)  '有可能脚本中的过程名前有 所有者. 如: zltools.zl_xxx

            blnBegin = True
            
            gstrBCode.Append Replace(strLine, """", "") '过程名称两侧可能有" 应该去掉
        Else
            '结束记录过程
            If (strFMT = "/" Or UBound(arrTxt) = dblRow) And blnBegin Then
                    rsProcedure.Filter = "P_Name = '" & strPName & "'"
                    If rsProcedure.RecordCount = 0 Then
                        rsProcedure.AddNew
                        rsProcedure!P_Name = strPName
                    End If
                
                    rsProcedure!P_Define = gstrBCode.ToString
                    rsProcedure!P_Ver = strFileVer
                    
                    
                    If lngSysNum <> 0 Then
                        rsProcedure!P_SysNum = lngSysNum
                    End If
                    If strSysName <> "" Then
                        rsProcedure!P_System = strSysName
                    End If
                    If strOwner <> "" Then
                        rsProcedure!P_Owner = strOwner
                    End If
                    
                    rsProcedure.Update
                    
                    blnBegin = False
                    gstrBCode.Clear
            ElseIf blnBegin Then
                gstrBCode.Append vbNewLine
                gstrBCode.Append Left(strLine, 4000)
            End If
        End If
    Next
    
    '如果脚本里有Drop Procedure语句 ,就从记录集中把过程删除
    For i = 0 To UBound(arrDelete)
        rsProcedure.Filter = "P_Name  = '" & arrDelete(i) & "'"
        If rsProcedure.RecordCount <> 0 Then
            rsProcedure.Delete
        End If
    Next
    
    rsProcedure.Filter = 0
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "获取文件失败" & err.Description, , gstrSysName
End Sub

Public Function LoadBaseProcs(ByVal strProcName As String) As String
    '功能：加载数据库存储过程
    Dim rsSource As ADODB.Recordset, strSQL As String
    Dim strProcText As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    '存储过程收集，收集数据库作为基本存储过程
    strSQL = "Select Name, Type, Text, Line 序号 From User_Source Where Type In ('PROCEDURE', 'FUNCTION') And Name =[1] Order By  Line"
    Set rsSource = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取数据库过程源码", strProcName)
    
    gstrBCode.Clear
    
    If Not rsSource.EOF Then
        Do While Not rsSource.EOF
            strTmp = rsSource!Text
            strTmp = Replace(strTmp, vbCr, "")
            strTmp = Replace(strTmp, vbLf, "")
            strTmp = Replace(strTmp, vbNewLine, "")
            
            If rsSource!序号 = 1 Then
                '数据库源码没有CREATE OR REPLACE
                gstrBCode.Append "CREATE OR REPLACE "
            Else
                gstrBCode.Append vbNewLine
            End If
            
            If UCase(strTmp) Like "*" & """" & UCase(strProcName) & """" & "*" Then
                    strTmp = Replace(UCase(strTmp), """" & UCase(strProcName) & """", strProcName)
            End If
            
            gstrBCode.Append strTmp
            rsSource.MoveNext
        Loop
    End If
    LoadBaseProcs = gstrBCode.ToString
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function UpdateProc2DB(rsProc As ADODB.Recordset, intType As Integer, Optional strErr As String) As Boolean
    '将过程集合保存至数据库
    '参数:rsProc-过程集合  intType-过程类型(1-变动过程 2-升级后被修改的过程)
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim lngID As Long
    Dim arrTxt() As String, i As Long
    Dim lngSysNum As Long, strIDs As String, arrIds As Variant
    
    On Error GoTo errH
    strErr = ""
    If rsProc Is Nothing Then
        UpdateProc2DB = True
        Exit Function
    End If
    If rsProc.RecordCount = 0 Then
        UpdateProc2DB = True
        Exit Function
    End If
    
    With rsProc
        .Filter = 0
        
        Do While Not .EOF
            lngID = GetProcIdByName(!P_Name)
            gcnOracle.BeginTrans
            '更新数据至zlProcedure
            If lngID = 0 Then
                If intType = 1 Then
                    strSQL = "Insert Into Zlprocedure (ID, 类型, 名称, 状态, 所有者, 系统编号, 升级前版本) Values" & vbNewLine & _
                                 "(Zlprocedure_Id.Nextval,1,'" & !P_Name & "',1,'" & !P_Owner & "'," & !P_SysNum & ",'" & !P_Ver & "')"
                Else
                    strSQL = "Insert Into Zlprocedure (ID, 类型, 名称, 状态, 所有者, 系统编号, 升级后版本) Values" & vbNewLine & _
                                 "(Zlprocedure_Id.Nextval,1,'" & !P_Name & "',1,'" & !P_Owner & "'," & !P_SysNum & ",'" & !P_Ver & "')"
                End If
            Else
                '删除已转出的内容
                gcnOracle.Execute "Delete from zlProcedureText where 性质=3 and 过程ID = (Select ID From zlProcedure where 状态 = 4 And ID = " & lngID & ")"
                gcnOracle.Execute "Update zlProcedure Set 状态 = 1 Where 状态 = 4 And ID = " & lngID    '只修改已转出过程的状态
                
                '更新数据
                If intType = 1 Then
                    strSQL = "Update zlProcedure Set 类型 = 1,所有者='" & !P_Owner & "',系统编号=" & !P_SysNum & ",升级前版本='" & !P_Ver & "'" & vbNewLine & _
                                 "Where Id = " & lngID
                Else
                    strSQL = "Update zlProcedure Set 类型 = 1,所有者='" & !P_Owner & "',系统编号=" & !P_SysNum & ",升级后版本='" & !P_Ver & "'" & vbNewLine & _
                                 "Where Id = " & lngID
                End If
            End If
            gcnOracle.Execute strSQL
            
            '删除zlProcedureText中的数据
            If lngID = 0 Then
                lngID = GetProcIdByName(!P_Name)
            End If
            
            If intType = 1 Then
                gcnOracle.Execute "Delete from zlProcedureText where 性质=1 and 过程ID = " & lngID
            Else
                gcnOracle.Execute "Delete from zlProcedureText where 性质=4 and 过程ID = " & lngID
            End If
            
            '插入过程定义到zlProcedureText
            arrTxt = Split(!P_Define, vbNewLine)
            strSQL = "Insert Into zlProcedureText(过程ID,性质,序号,内容) "
            For i = 0 To UBound(arrTxt)
                If i = UBound(arrTxt) Then
                    strSQL = strSQL & vbNewLine & "Select " & lngID & "," & IIf(intType = 1, "1", "4") & "," & (i + 1) & ",'" & Replace(arrTxt(i), "'", "''") & "' From Dual "
                Else
                    strSQL = strSQL & vbNewLine & "Select " & lngID & "," & IIf(intType = 1, "1", "4") & "," & (i + 1) & ",'" & Replace(arrTxt(i), "'", "''") & "' From Dual Union All "
                End If
            Next
            gcnOracle.Execute strSQL
            
            
            If strIDs = "" Then
                lngSysNum = !P_SysNum
                strIDs = lngID
            Else
                strIDs = strIDs & "," & lngID '拼接所有ID
            End If
            
            gcnOracle.CommitTrans
            .MoveNext
        Loop
    End With
    
    '删除非该系统的其他数据,因为有的库zlProcedureText表外键不是级联删除,因此要先删除子表
    If intType = 1 Then
        gcnOracle.BeginTrans
        arrIds = TranStr2Var(strIDs, ",", 2000) '防止字符超长
        For i = 0 To UBound(arrIds)
            strSQL = "Delete From zlProcedureText Where 过程ID In  " & vbNewLine & _
                        "(Select ID from Zlprocedure Where 类型 = 1 And 系统编号 = " & lngSysNum & " And  ID Not In (Select Column_Value From Table(f_Str2list('" & arrIds(i) & "', ','))))"
            gcnOracle.Execute strSQL
        
            strSQL = "Delete From zlProcedure Where 类型 = 1 And 系统编号 = " & lngSysNum & " And  ID Not In (Select Column_Value From Table(f_Str2list('" & arrIds(i) & "', ',')))"
            gcnOracle.Execute strSQL
        Next
        
        gcnOracle.CommitTrans
    End If
    UpdateProc2DB = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    strErr = err.Description
End Function


Public Function GetProcIdByName(ByVal strName As String, Optional ByVal intProp As Integer, Optional ByVal intStat As Integer) As Long
    '根据名称返回过程ID
    '参数说明:
    'strName -名称
    'intPorc-类型-1-用户变动过程;2-空白过程;3-用户过程
    'intStat-状态:1-待调整;2-已自动调整;3-已人工调整;4-已导出
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim lngID As Long
    
    On Error GoTo errH
    strSQL = "Select Id From zlProcedure Where 名称 = [1]" & IIf(intProp = 0, "", " And 类型 = [2]") & IIf(intStat = 0, "", "And 状态 = [3]")
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取ID", strName, intProp, intStat)
    
    
    If rsTmp.RecordCount = 0 Then
        lngID = 0
    Else
        lngID = rsTmp!id
    End If
    
    GetProcIdByName = lngID
    Exit Function
errH:
    MsgBox "获取过程ID出错" & vbNewLine & err.Description, , gstrSysName
End Function

Public Function GetPorcTxtByName(ByVal strName As String, ByVal intType As Integer) As String
    '根据过程名称和文本类型返回过程文本
    'strName:过程名称  intType:文本类型 1-上次定义过程;2-上次标准过程;3-本次自定过程;4-本次标准过程
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strResult As String
    
    On Error GoTo errH
    
    strSQL = "Select 内容  From zlProcedureText Where 性质 = [2]  And 过程ID = (Select ID From zlProcedure Where 名称=[1] ) Order by 序号 "
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取过程文本", strName, intType)
    
    If rsTmp.RecordCount = 0 Then
        Exit Function
    End If
    
    Do While Not rsTmp.EOF
        If strResult = "" Then
            strResult = rsTmp!内容
        Else
            strResult = strResult & vbNewLine & rsTmp!内容
        End If
        rsTmp.MoveNext
    Loop
    
    GetPorcTxtByName = strResult
    Exit Function
errH:
    MsgBox "获取过程文本出现错误." & vbNewLine & err.Description, , "错误"
End Function


Public Function CheckProcManage() As Boolean
    '功能:检查用户变动过程管理模块是否已经加载
    '说明:用户变动过程是在升级前使用的功能,不能通过脚本来提交,所以要在程序中进行判断后临时添加\修改
    '需要添加\修改的部分:1.管理工具中模块的添加;2.zlProcedure\zlProcedureText表结构的修改
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    On Error Resume Next
    
    '1.添加模块
    strSQL = "Select 1 From zlSvrTools Where 上级 = '01' And 标题 In ('变动过程升级管理','变动过程日常管理')"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "检查变动过程模块")
    
    If rsTmp.RecordCount <> 2 Then
        gcnOracle.Execute "Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0106','01','变动过程升级管理','B',Null,16)"
        gcnOracle.Execute "Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0107','01','变动过程日常管理','U',Null,17)"
    End If
    
    '2.修改结构zlProcedure表增加了三个字段  升级前版本\升级后版本\系统编号
    err.Clear
    strSQL = "Select 升级前版本,升级后版本,系统编号 From zlTools.zlProcedure where 1=0"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "检查变动过程结构")
    
    '如果出现错误,就添加字段
    If err.Number <> 0 Then
        
        If gcnTools Is Nothing Then
            Set gcnTools = GetConnection("ZLTOOLS")
        End If
        
        gcnTools.Execute "Alter Table Zltools.Zlprocedure Add 升级前版本 Varchar2(50)"
        gcnTools.Execute "Alter Table Zltools.Zlprocedure Add 升级后版本 Varchar2(50)"
        gcnTools.Execute "Alter Table Zltools.zlProcedure Add 系统编号 Number(5)"
        gcnTools.Execute "Alter Table Zltools.Zlprocedure Modify 说明 Varchar2(2000)"
    End If
    
    CheckProcManage = True
End Function

Public Function ConvertStr(ByVal strSource As String) As String
    '功能:去掉字符串的空格\换行符,并转换为大写
    
    strSource = UCase(strSource)
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, vbNewLine, "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    strSource = Replace(strSource, vbTab, "")
    strSource = Replace(strSource, vbBack, "")
    ConvertStr = strSource
End Function

Public Function GetSqlColor() As String
    '公共方法:获取语法控件的SQL语法高亮显示设置
    '获取后直接将语法控件的SyntaxScheme属性设为返回值即可
    Dim strColor As String, strPath As String
    
    If Not gblnInIDE Then '增加多环境支持
        strPath = App.Path & "\PUBLIC\_sql.schclass"
    Else
        strPath = gobjFSO.GetParentFolderName(GetSetting("ZLSOFT", "公共全局", "程序路径")) & "\PUBLIC\_sql.schclass"
    End If
    If Not gobjFSO.FileExists(strPath) Then
        strPath = "C:\Appsoft\PUBLIC\_sql.schclass"
    End If
    
    If gobjFSO.FileExists(strPath) Then
        strColor = ReadFileToString(strPath)
    End If
    GetSqlColor = strColor
End Function
