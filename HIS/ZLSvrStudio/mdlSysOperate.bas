Attribute VB_Name = "mdlSysOperate"
Option Explicit
'文件类型,该类型顺序与文件执行顺序相同
'X.X.X可能为4位版本号X.X.X.X,此时为特殊SP脚本。
Public Enum FileType
    'FT_Before 脚本与FT_DBA脚本执行执行顺序可以互换
    FT_DBA = 0 '需要DBA用户执行的脚本(System用户):ZLUPgradeX.X.X_DBA.sql,ZL*_X.X.X_DBA.sql
    FT_Before = 1 '提前执行脚本：ZLUPgradeX.X.X_Before.sql.sql(管理工具）,ZL*_X.X.X_History_Before.sql (应用系统历史库)ZL*_X.X.X_Before.sql(应用系统在线库) *代表系统号\100
    FT_Standard = 2 '普通升级脚本：ZLUPgradeX.X.X.sql,ZLUPgradeX.X.X(补充).sql,ZL*_X.X.X.sql ,ZL*_X.X.X(补充).sql,ZL*_X.X.X_History.sql
    FT_Optional = 3 '可选执行脚本:ZLUPgradeX.X.X_Optional.sql,ZL*_X.X.X_Optional.sql，ZL*_X.X.X__HISTORY_Optional.sql
    FT_Deferred = 4 '延迟执行脚本:ZL*_X.X.X_Deferred.sql,ZL*_X.X.X__HISTORY_DEFERRED
End Enum
'文件所属系统
Public Enum SysType
    ST_Tools = 0 '管理工具脚本,具有文件类型：FT_Before,FT_DBA,FT_Standard,FT_Optional
    ST_App = 1 '应用系统在线库,具有文件类型：FT_Before,FT_DBA,FT_Standard,FT_Optional，FT_Deferred
    ST_History = 2 '应用系统历史库，具有文件类型：FT_Before,FT_Standard,FT_Deferred，FT_Optional
End Enum
'版本类型
Public Enum VersionType
    VT_Normal = 0 '正常版本
    VT_Supple = 1 '补充发布版本，下一个大版本发布后，前一个版本新发布的SP就是补充版本
End Enum

Public Enum UserCheckType
    UCT_ZLTOOLS = 0 '管理工具用户验证
    UCT_DBAUser = 1 'DBA用户验证
    '以前该类的序号为1，现在调整为2，主要后面连续这几种类型都是通过直接调用窗体来使用的
    UCT_CurZLBAK = 2 '当前历史库验证
    UCT_NormalUser = 3 '普通用户验证
    UCT_SysOwner = 4 '管理员登录验证
    UCT_RACInsUser = 5 'RAC实例用户验证
    UCT_AuditLog = 6   '记录重要日志
End Enum

Public gcllMustObj As Collection '必要对象检查
Public gobjLog As TextStream
Private mstrStSysOwner As String '标准版所有者
Public Function CheckAndAdjustMustTable(ByVal strTable As String, Optional ByVal strColumn As String, Optional ByVal blnMsg As Boolean, Optional ByVal strOwner As String = "ZLTOOLS", Optional ByVal blnCache As Boolean = True) As Boolean
'功能：检查并修正必要的数据结构
'参数：strTable=表名
'         strColumn=列名
'         blnMsg=检查并修复失败是否提示
'         strOwner=对象所有者
'         blnCache=判断是否缓存对象检查结果，一些特殊对象需要缓存，其余普通对象不缓存，防止程序缓存数据较多
'返回：检查并修复是否成功
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnHaveTable As Boolean, blnHaveColumn As Boolean
    Dim strFMT As String
    Dim objGetTmp As New clsObjectInfo, objParent As clsObjectInfo, objCurent As clsObjectInfo
    Dim blnHaveData As Boolean
    
    strTable = UCase(strTable): strColumn = UCase(strColumn): strOwner = UCase(strOwner)
    '加载内置对象检查修复方案
    If gcllMustObj Is Nothing Then
        Set gcllMustObj = New Collection
        'ZLUpgrade表检查以及提前列检查
        Set objParent = objGetTmp.GetObject("ZLUPGRADE", OT_Table, _
                                        "CREATE TABLE ZLTOOLS.zlUpgrade(系统 NUMBER(5),原始版本 VARCHAR2(10),目标版本 VARCHAR2(10),升迁时间 DATE,升迁结果 NUMBER(1)" & _
                                        ",结果版本 VARCHAR2(10),中止语句 VARCHAR2(200),提前执行 number(1))PCTFREE 5|" & _
                                        "ALTER TABLE ZLTOOLS.zlUpgrade ADD CONSTRAINT  zlUpgrade_UQ_升迁时间 Unique (系统,升迁时间)   USING INDEX PCTFREE 5|" & _
                                        "ALTER TABLE ZLTOOLS.zlUpgrade ADD CONSTRAINT  zlUpgrade_FK_系统 FOREIGN KEY (系统) REFERENCES zlSystems(编号) ON DELETE CASCADE")
        Set objCurent = objGetTmp.GetObject("提前执行", OT_Column, "alter Table ZLTOOLS.ZLUPGRADE add 提前执行 number(1)", , objParent)
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLUPGRADE|提前执行")
        'ZLBAKTABLES表检查
        Set objCurent = objGetTmp.GetObject("ZLBAKTABLES", OT_Table, _
                                        "Create Table ZLTOOLS.zlBakTables(系统 Number(5),表名 Varchar2(30),组号 Number(2),序号 Number(3),直接转出 Number(1),停用触发器 number(1))|" & _
                                        "Alter Table ZLTOOLS.zlBakTables    Add Constraint zlBakTables_PK Primary Key (系统,表名) USING INDEX PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlBakTables Add Constraint zlBakTables_FK_系统 Foreign Key (系统) References zlSystems(编号) On Delete Cascade")
        
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLBAKTABLES|")
        'ZLBAKSPACES表检查
        Set objCurent = objGetTmp.GetObject("ZLBAKSPACES", OT_Table, _
                                        "Create Table ZLTOOLS.zlBakSpaces(系统 Number(5),编号 Number(18),名称 Varchar2(30),所有者 Varchar2(30),DB连接 Varchar2(128),当前 Number(1),只读 Number(1))PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlBakSpaces Add Constraint zlBakSpaces_PK Primary Key (系统,编号) USING INDEX PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlBakSpaces    Add Constraint zlBakSpaces_UQ_名称 Unique (系统,名称) USING INDEX PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlBakSpaces Add Constraint zlBakSpaces_FK_系统 Foreign Key (系统) References zlSystems(编号) On Delete Cascade")
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLBAKSPACES|")
        'zlUpgradeConfig表检查
        Set objCurent = objGetTmp.GetObject("zlUpgradeConfig", OT_Table, _
                                        "Create Table ZLTOOLS.zlUpgradeConfig(项目 varchar2(50),内容 varchar2(4000))PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.zlUpgradeConfig Add Constraint zlUpgradeConfig_PK Primary Key (项目) USING INDEX PCTFREE 5|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('客户端状态',1)|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('用户状态',1)|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('后台作业状态',1)|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('触发器状态',1)|" & _
                                        "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('禁用的系统调度',Null)")
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|zlUpgradeConfig|")
        'ZLTriggers表检查
        Set objCurent = objGetTmp.GetObject("ZLTriggers", OT_Table, _
                                        "Create Table ZLTOOLS.ZLTriggers(名称 varChar2(100),所有者 varChar2(100))PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.ZLTriggers Add Constraint ZLTriggers_UQ_名称 Unique (名称,所有者) USING INDEX PCTFREE 5|" & _
                                        "Alter Table ZLTOOLS.ZLTriggers Modify 名称  constraint ZLTriggers_NN_名称   not  null")
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLTriggers|")
        'ZLClient表检查以及系统升级禁用列检查
        Set objCurent = objGetTmp.GetObject("系统升级禁用", OT_Column, "alter Table ZLTOOLS.ZLCLIENTS add 系统升级禁用 number(1)", , objGetTmp.GetObject("ZLCLIENTS", OT_Table))
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLCLIENTS|系统升级禁用")
        'ZLAutoJob表检查以及系统升级停用列检查
        Set objCurent = objGetTmp.GetObject("系统升级停用", OT_Column, "alter Table ZLTOOLS.ZLAutoJobs add 系统升级停用 number(1)", , objGetTmp.GetObject("ZLAutoJobs", OT_Table))
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|ZLAutoJobs|系统升级停用")
        '上机人员表表检查以及系统升级停用列检查
        Set objCurent = objGetTmp.GetObject("系统升级锁定", OT_Column, "alter Table " & gstrUserName & ".上机人员表 add 系统升级锁定 number(1)", gstrUserName, objGetTmp.GetObject("上机人员表", OT_Table, , gstrUserName))
        gcllMustObj.Add objCurent, UCase(gstrUserName & "|上机人员表|系统升级锁定")
        'Zlsvrtools表检查以及次序列检查
        Set objCurent = objGetTmp.GetObject("次序", OT_Column, "alter Table ZLTOOLS.Zlsvrtools add 次序 number(3)", , objGetTmp.GetObject("Zlsvrtools", OT_Table))
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|Zlsvrtools|次序")
        'zlParameters表检查以及部门列检查
        Set objCurent = objGetTmp.GetObject("部门", OT_Column, "alter table Zltools.zlParameters add 部门 NUMBER(1)", , objGetTmp.GetObject("zlParameters", OT_Table))
        gcllMustObj.Add objCurent, UCase("ZLTOOLS|zlParameters|部门")
    End If
    
    On Error Resume Next
    Set objCurent = gcllMustObj(strOwner & "|" & strTable & "|" & strColumn)
    If err.Number <> 0 Then
        err.Clear
        If strColumn = "" Then
            Set objCurent = objGetTmp.GetObject(strTable, OT_Table, , strOwner)
        Else
            Set objParent = gcllMustObj(strOwner & "|" & strTable & "|")
            If err.Number <> 0 Then
                err.Clear
                Set objParent = objGetTmp.GetObject(strTable, OT_Table, , strOwner)
            Else
                If blnCache Then gcllMustObj.Remove strOwner & "|" & strTable & "|" '合并检查对象
            End If
            Set objCurent = objGetTmp.GetObject(strColumn, OT_Column, , strOwner, objParent)
        End If
        '缓存对象检查
        If blnCache Then gcllMustObj.Add objCurent, UCase(strOwner & "|" & strTable & "|" & strColumn)
    End If
    If Not objCurent.ObjectCheck(blnMsg) Then
        Exit Function
    End If
    CheckAndAdjustMustTable = True
    On Error GoTo errh
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function GetConnection(ByVal strUserName As String, Optional ByVal blnValidate As Boolean = True) As ADODB.Connection
'功能：获取连接
'参数：strUserName=ZLTOOLS,管理工具，DBA-DBA用户确认，其他，其他用户确认
'          blnValidate=True:如果不能获取连接则需要输入密码，产生连接,False=如果不能获取连接则直接退出
    Dim uctType As UserCheckType
    Dim cnTmp As ADODB.Connection
    Dim blnNew As Boolean, blnViladate As Boolean
    
    Select Case UCase(strUserName)
        Case "ZLTOOLS"
            If gcnTools Is Nothing Then
                blnNew = True
            ElseIf gcnTools.State = adStateClosed Then
                blnNew = True
            End If
            If blnNew Then
                Set gcnTools = gobjRegister.GetConnection(gstrServer, "ZLTOOLS", IIf(gstrToolsPwd = "", "ZLTOOLS", gstrToolsPwd), False, MSODBC, "", False)
                If gcnTools.State = adStateOpen Then
                    Call SetSQLTrace(gstrServer, "ZLTOOLS", gcnTools)
                    gstrToolsPwd = IIf(gstrToolsPwd = "", "ZLTOOLS", gstrToolsPwd)
                    Set GetConnection = gcnTools: Exit Function
                ElseIf gstrToolsPwd = "" Then
                    Set gcnTools = gobjRegister.GetConnection(gstrServer, "ZLTOOLS", "ZLSOFT", False, MSODBC, "", False)
                    If gcnTools.State = adStateOpen Then
                        Call SetSQLTrace(gstrServer, "ZLTOOLS", gcnTools)
                        gstrToolsPwd = "ZLSOFT"
                        Set GetConnection = gcnTools: Exit Function
                    End If
                End If
            Else
                Set GetConnection = gcnTools: Exit Function
            End If
            uctType = UCT_ZLTOOLS
        Case "DBA", "SYSTEM", "SYS"
            If gcnSystem Is Nothing Then
                blnNew = True
            ElseIf gcnSystem.State = adStateClosed Then
                blnNew = True
            End If
            If gstrSysPwd <> "" And blnNew Then
                Set gcnSystem = gobjRegister.GetConnection(gstrServer, gstrSysUser, gstrSysPwd, False, MSODBC, "", False)
                If gcnSystem.State = adStateOpen Then
                    Call SetSQLTrace(gstrServer, gstrSysUser, gcnSystem)
                    Set GetConnection = gcnSystem: Exit Function
                End If
            ElseIf Not blnNew Then
                Set GetConnection = gcnSystem: Exit Function
            End If
            uctType = UCT_DBAUser
            If UCase(strUserName) = "DBA" Then strUserName = "SYSTEM"
        Case Else
            uctType = UCT_NormalUser
    End Select
    If blnValidate Then
        If Not frmUserCheckLogin.ShowLogin(uctType, cnTmp, strUserName) Then Exit Function
        
        Call SetSQLTrace(gstrServer, strUserName, cnTmp)
        Set GetConnection = cnTmp
        If uctType = UCT_ZLTOOLS Then
            Set gcnTools = cnTmp
        ElseIf uctType = UCT_DBAUser Then
            Set gcnSystem = cnTmp
        End If
    End If
End Function

Public Sub RecToLog(ByVal rsInput As ADODB.Recordset, Optional ByVal strSort As String, Optional ByVal strName As String)
'将记录集转换为字符串，用来跟踪记录日志
    Dim i As Long
    Dim lngShort As Long
    Dim strLine As String
    
    If Not gblnTrace Then Exit Sub
    If rsInput Is Nothing Then
        WriteTraceLog "===============" & strName & "========================="
        WriteTraceLog "Nothing"
    End If
    rsInput.Filter = ""
    rsInput.Sort = strSort
    '对列名进行日志
    
   WriteTraceLog "===============" & strName & "========================="
    For i = 0 To rsInput.Fields.Count - 1
        strLine = strLine & RPAD(rsInput.Fields(i).name, 12)
    Next
    WriteTraceLog strLine
    Do While Not rsInput.EOF
        lngShort = 0: strLine = ""
        For i = 0 To rsInput.Fields.Count - 1
            If Len(rsInput.Fields(i).value & "") < 9 And lngShort <> 0 Then
                strLine = strLine & RPAD(rsInput.Fields(i).value & "", 9)
                lngShort = IIf(lngShort - 3 <= 0, 0, lngShort - 3)
            ElseIf Len(rsInput.Fields(i).value & "") > 12 Then
                strLine = strLine & RPAD(rsInput.Fields(i).value & "", 12)
                lngShort = lngShort + Len(rsInput.Fields(i).value & "") - 12
            Else
                strLine = strLine & RPAD(rsInput.Fields(i).value & "", 12)
            End If
        Next
        WriteTraceLog strLine
        rsInput.MoveNext
    Loop
End Sub

Public Function GetToolsVersion() As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strVer As String, blnUpdate As Boolean
    
    On Error GoTo errh
    '读取管理工具版本
    strSQL = "Select 内容 From Zlreginfo Where 项目 = '版本号'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
    If Not rsTmp.EOF Then strVer = Trim(rsTmp!内容 & "")
    '若管理工具版本是无效版本，则自动修正
    If strVer = "" Then
        blnUpdate = Not rsTmp.EOF '需要跟新版本
        On Error Resume Next
        strSQL = "Select 结果版本 From Zlupgrade Where 系统 Is Null And Nvl(提前执行, 0) = 0 Order By 升迁时间 Desc"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
        If Not rsTmp.EOF Then strVer = Trim(rsTmp!结果版本 & "")
        If err.Number <> 0 Then err.Clear
        On Error GoTo errh
        If strVer <> "" Then
            If blnUpdate Then
                gcnOracle.Execute "Update ZLreginfo set 内容='" & strVer & "' where 项目='版本号'"
            Else
                gcnOracle.Execute "Insert Into zlRegInfo(项目,行号,内容) Values('版本号',1,'" & strVer & "')"
            End If
        End If
    End If
    GetToolsVersion = strVer
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'功能：将指定INI配置文件的内容读取到记录集中
'返回：Nothing或包含"项目,内容"的记录集,其中同一项目可能有多行内容
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "项目", adVarChar, 100
    rsTmp.Fields.Append "内容", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = Null
                rsTmp.Update
            End If
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))

            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!项目 = strItem
            rsTmp!内容 = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!项目 = strItem
        rsTmp!内容 = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Public Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'功能：检查对应的配置文件格式是否正确
'参数：rsINI=存放配置文件内容的记录集，包含"项目,内容"字段
'      strItem=配置文件中必须要求有内容的项目串,如"项目1|项目2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "项目='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If rsINI!内容 & "" = "" Then Exit Function
        If arrItem(i) Like "*版本号" Then
            If Not IsVerSion(rsINI!内容) Then Exit Function
        End If
    Next
    CheckINIValid = True
End Function

Public Function SplitLine(ByVal strSQL As String) As Variant
'功能：对SQL进行换行拆分，同时记录换行符
    Dim arrLine As Variant, arrReturn() As Variant
    Dim i As Long, j As Long, lngStart As Long, lngEx As Long, lngCur As Long
    Dim strTmp As String
    arrReturn = Array()
    If strSQL = "" Then SplitLine = arrReturn: Exit Function
    arrLine = Split(Replace(Replace(strSQL, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ReDim Preserve arrReturn(UBound(arrLine) * 2)
    lngStart = 1
    For i = LBound(arrLine) To UBound(arrLine)
        If i <> 0 Then
            strTmp = Mid(strSQL, lngStart, 2)
            If strTmp = vbCrLf Then
                arrReturn(i * 2 - 1) = vbCrLf
                lngStart = lngStart + 2
            Else
                arrReturn(i * 2 - 1) = Mid(strSQL, lngStart, 1)
                lngStart = lngStart + 1
            End If
        End If
        arrReturn(i * 2) = arrLine(i)
        lngStart = lngStart + Len(arrLine(i))
    Next
    SplitLine = arrReturn
End Function

Public Function TrimCommentLossless(ByVal strSQL As String) As String
'功能：无损去掉注释，与TrimComment比较，该算法不会损害真实数据。
    Dim arrLine As Variant, arrTmp As Variant
    Dim i As Long, j As Long
    Dim blnStr As Boolean, blnMultiCom As Boolean
    Dim lngPos1 As Long, lngPos2 As Long, lngPos3 As Long
    Dim blnAddLine As Boolean
    Dim strTmp As String, strFMT As String
    
    On Error GoTo errh
    '去除多行注释。
    arrTmp = Split(strSQL, "'")
    strFMT = "": blnStr = False: blnMultiCom = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            arrLine = SplitLine(arrTmp(i))
            blnAddLine = True
            For j = LBound(arrLine) To UBound(arrLine) Step 2
                strTmp = arrLine(j)
                blnAddLine = j <> UBound(arrLine)
                If blnMultiCom Then '已经处于多行注释范围，则优先查找结束符
                    lngPos2 = InStr(strTmp, "*/")
                    If lngPos2 > 0 Then
                        strTmp = Mid(strTmp, lngPos2 + 2)
                        blnMultiCom = False
                    Else
                        strTmp = "": blnAddLine = False
                    End If
                End If
                If Not blnMultiCom Then '针对/* -- */ 与/*   */--处理
                    lngPos2 = InStr(strTmp, "/*")
                    lngPos1 = InStr(strTmp, "--")
                    '去掉有效的多行注释内容'/* --*/ ,/* */ 代码段 --/* */
                    '1、存在--,但是--在多行开始符之后
                    '2、不存在--，存在多行开始符
                    Do While Not blnMultiCom And (lngPos2 > 0 And lngPos2 < lngPos1 Or lngPos1 = 0 And lngPos2 > 0)
                        lngPos3 = InStr(lngPos2, strTmp, "*/")
                        If lngPos3 > 0 Then
                            strTmp = Left(strTmp, lngPos2 - 1) & Mid(strTmp, lngPos3 + 2)
                        Else
                            strTmp = Left(strTmp, lngPos2 - 1)
                            blnMultiCom = True
                        End If
                        lngPos2 = InStr(strTmp, "/*")
                        lngPos1 = InStr(strTmp, "--")
                    Loop
                End If
                '注释中的空行，则不做处理
                If blnAddLine Then
                    strFMT = strFMT & strTmp & arrLine(j + 1)
                Else
                    strFMT = strFMT & strTmp
                End If
            Next
        Else
            strTmp = ""
            '针对 "'B''C''D'"该类字符串进行识别
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '存在空串，则为单引号字符
                        strTmp = strTmp & "''"
                    Else '不存在，则该处为字符的最后一段
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            strFMT = strFMT & "'" & strTmp & "'"
        End If
        If Not blnMultiCom Then '非多行注释，则调整字符串边界
            blnStr = Not blnStr '开始进入字符串边界
        End If
    Next
    
    '去除单行注释
    arrTmp = Split(strFMT, "'")
    strFMT = "": blnStr = False: blnMultiCom = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            arrLine = SplitLine(arrTmp(i))
'            blnMultiCom = False
            For j = LBound(arrLine) To UBound(arrLine) Step 2
                strTmp = arrLine(j)
                If j = LBound(arrLine) And blnMultiCom Then
                    blnMultiCom = UBound(arrLine) = 0
                Else
                    blnAddLine = j <> UBound(arrLine)
                    lngPos1 = InStr(strTmp, "--")
                    If lngPos1 > 0 Then
                        strTmp = Left(strTmp, lngPos1 - 1)
                        blnMultiCom = UBound(arrLine) = j
                    End If
                    If blnAddLine Then
                        strFMT = strFMT & strTmp & arrLine(j + 1)
                    Else
                        strFMT = strFMT & strTmp
                    End If
                End If
            Next
        Else
            strTmp = ""
            '针对 "'B''C''D'"该类字符串进行识别
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '存在空串，则为单引号字符
                        strTmp = strTmp & "''"
                    Else '不存在，则该处为字符的最后一段
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            strFMT = strFMT & "'" & strTmp & "'"
        End If
        If Not blnMultiCom Then
            blnStr = Not blnStr '开始进入字符串边界
        End If
    Next
    TrimCommentLossless = strFMT
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function GetFMTSQLStr(ByVal strSQL As String, ByRef cllStrs As Collection) As String
'功能：获取SQL中的字符串，并用占位符占位，返回格式化的SQL
    Dim arrTmp As Variant
    Dim i As Long, j As Long, intIndex As Integer
    Dim strFMT As String, strTmp As String
    Dim blnStr As Boolean
    
    Set cllStrs = New Collection
    arrTmp = Split(strSQL, "'")
    strFMT = "": blnStr = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            strFMT = strFMT & arrTmp(i)
        Else
            strTmp = ""
            '针对 "'B''C''D'"该类字符串进行识别
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '存在空串，则为单引号字符
                        strTmp = strTmp & "''"
                    Else '不存在，则该处为字符的最后一段
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            intIndex = intIndex + 1
            '标记字符串
            strFMT = strFMT & "[S" & intIndex & "]"
            cllStrs.Add strTmp, "S" & intIndex
        End If
        blnStr = Not blnStr '开始进入字符串边界
    Next
    arrTmp = SplitLine(strFMT)
    strFMT = "": blnStr = False
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        strTmp = TrimEx(arrTmp(i))
        If strTmp <> "" Then
            If Right(strTmp, 1) = ";" And i <> UBound(arrTmp) Then
                strFMT = strFMT & " " & strTmp & vbCrLf
            Else
                strFMT = strFMT & " " & strTmp
            End If
        End If
    Next
    '去掉操作符中的空格
    arrTmp = SplitLine(strFMT)
    strFMT = ""
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        strTmp = TrimEx(TrimBesideOperator(arrTmp(i)))
        If strTmp <> "" Then
            If Right(strTmp, 1) = ";" And i <> UBound(arrTmp) Then
                strFMT = strFMT & " " & strTmp & vbCrLf
            Else
                strFMT = strFMT & " " & strTmp
            End If
        End If
    Next
    GetFMTSQLStr = UCase(strFMT)
End Function

Public Function TrimBesideOperator(ByVal strText As String) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    Dim i As Long
    
    strText = Replace(Replace(strText, " :", ":"), ": ", ":")
    strText = Replace(Replace(strText, " =", "="), "= ", "=")
    strText = Replace(Replace(strText, " .", "."), ". ", ".")
    strText = Replace(Replace(strText, " )", ")"), ") ", ")")
    strText = Replace(Replace(strText, " (", "("), "( ", "(")
    strText = Replace(Replace(strText, " %", "("), "% ", "%")
    strText = Replace(Replace(strText, " \", "\"), "\ ", "\")
    TrimBesideOperator = strText
End Function

Public Function GetInfoInsideBracket(ByVal strInfo As String, Optional ByVal strLeftChar As String, Optional ByVal strRightChar As String) As String
'从括号里面取内容
'返回括号里面的内容，只取最外层
    Dim lngSart As Long, lngEnd As Long
    If strRightChar = "" Then strRightChar = ")"
    If strLeftChar = "" Then strLeftChar = "("
    lngEnd = InStrRev(strInfo, strRightChar) - Len(strRightChar) + 1 '算头不算尾，所以不减一
    lngSart = InStr(strInfo, strLeftChar) + Len(strLeftChar)
    If lngEnd < lngSart Then
        GetInfoInsideBracket = ""
    Else
        GetInfoInsideBracket = Mid(strInfo, lngSart, lngEnd - lngSart)
    End If
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    TrimEx = strText
End Function

Public Function TrimComment(ByVal strSQL As String) As String
'功能：去掉写在单行strSQL语句后面的"--"注释
'说明：主要是RunSQLFile的子函数
    Dim blnStr As Boolean
    Dim i As Long, K As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                K = i: Exit For
            End If
        Next
        If K > 0 Then strSQL = RTrim(Left(strSQL, K - 1))
    End If
    TrimComment = strSQL
End Function

Public Function SplitSQL(ByVal strSQL As String) As String
'功能：取";"结尾前面的的SQL语句,可能";"号后有"--"注释。
'说明：主要是RunSQLFile的子函数
    Dim i As Long, K As Long
    
    '先去掉注释部份
    strSQL = TrimComment(strSQL)
    
    For i = Len(strSQL) To 1 Step -1
        If Mid(strSQL, i, 1) = ";" Then
            K = i: Exit For
        End If
    Next
    If K > 0 Then strSQL = Left(strSQL, K - 1)
    
    SplitSQL = strSQL
End Function

Public Function RemoveMark(ByVal strText As String) As String
'功能：去除一段文字中的前导"--"注释标记
    Dim arrText As Variant, strTemp As String, i As Long
    
    arrText = Split(strText, vbCrLf)
    
    strText = ""
    For i = 0 To UBound(arrText)
        strTemp = arrText(i)
        If Left(strTemp, 2) = "--" And Replace(strTemp, "-", "") <> "" Then
            strText = strText & vbCrLf & Mid(strTemp, 3)
        End If
    Next
    RemoveMark = Mid(strText, 3)
End Function

Public Function GetLogSQL(objSQL As clsSQLInfo) As String
'功能：获取简要SQL语句，用于填写日志
    Dim strSQL As String
    
    If objSQL.Block Then
        If objSQL.BlockName <> "" Then
            strSQL = Trim(Split(objSQL.SQL, vbCrLf)(0))
            If InStr(strSQL, "(") > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(strSQL, "(") - 1))
            End If
            If InStr(1, strSQL, " as", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " is", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " is", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " Return", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " Return", vbTextCompare) - 1))
            End If
        Else '匿名块
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 255)
        End If
    ElseIf UCase(LTrim(objSQL.SQL)) Like "CREATE * VIEW" Then
        '视图特殊处理
        strSQL = Split(objSQL.SQL, vbCrLf)(0)
        If InStr(1, strSQL, " as", vbTextCompare) > 0 Then '视图只能用as
            strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
        End If
    Else
        If InStr(objSQL.SQL, vbCrLf) > 0 Then
            '多行SQL
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 255)
        Else
            strSQL = ActualStr(objSQL.SQL, 255)
        End If
    End If
    GetLogSQL = strSQL
End Function

Public Function CheckInitFile(ByVal lngSys As Long, ByVal strFile As String, Optional ByVal blnOnlyCheck As Boolean, Optional ByRef rsReturnINI As ADODB.Recordset, Optional ByVal blnUpgradeCheck As Boolean = True) As Boolean
'参数：blnUpgradeCheck=检查升迁检查文件
   Dim strSysPath As String, strTmp As String
   Dim rsINI As ADODB.Recordset
   If Not gobjFile.FileExists(strFile) Then
        If Not blnOnlyCheck Then MsgBox "安装配置文件""" & strFile & """不存在。", vbExclamation, gstrSysName
        Exit Function
    End If
    If UCase(gobjFile.GetFileName(strFile)) <> IIf(lngSys = 0, "ZLSERVER.SQL", "ZLSETUP.INI") Then
        If Not blnOnlyCheck Then MsgBox "安装配置文件名不正确。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If lngSys = 0 Then '管理工具
        '检查管理工具升级检查函数文件是否存在。
        If blnUpgradeCheck Then
            strSysPath = gobjFile.GetParentFolderName(strFile)
            strTmp = strSysPath & "\zlUpgradeCheck.sql"
            If Not gobjFile.FileExists(strTmp) Then
                If Not blnOnlyCheck Then MsgBox "管理工具升级检查文件""" & strTmp & """不存在。", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    Else '应用系统
        Set rsINI = ReadINIToRec(strFile)
        If Not CheckINIValid(rsINI, "系统号|版本号|表空间|管理工具版本号") Then
            If Not blnOnlyCheck Then MsgBox "安装配置文件格式不正确。", vbExclamation, gstrSysName
            Exit Function
        End If
        '配置文件系统号不匹配
        rsINI.Filter = "项目='系统号'"
        If Val(rsINI!内容) <> lngSys \ 100 Then
            If Not blnOnlyCheck Then MsgBox "所选配置文件不是本系统的安装配置文件。", vbExclamation, gstrSysName
            Exit Function
        End If
        strSysPath = gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(strFile))
        '系统升迁目录检查
        If Not gobjFile.FolderExists(strSysPath & "\升级脚本") Then
            If Not blnOnlyCheck Then MsgBox "系统升迁目录""" & strSysPath & "\升级脚本""不存在。", vbExclamation, gstrSysName
            Exit Function
        End If
        If blnUpgradeCheck Then
            '检查应用系统升级检查函数文件是否存在。
            strTmp = strSysPath & "\升级脚本\zl" & lngSys \ 100 & "_UpgradeCheck.sql"
            If Not gobjFile.FileExists(strTmp) Then
                If Not blnOnlyCheck Then MsgBox "系统升级检查文件""" & strTmp & """不存在。", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
        '对应的安装脚本文件是否存在,不需要检查，因为已经取消了可选脚本执行
    End If
    Set rsReturnINI = rsINI
    CheckInitFile = True
End Function

Public Function GetUpgradeFiles(ByVal rsUpgradeFiles As ADODB.Recordset, ByVal lngSys As Long, ByVal strCurVer As String, ByVal strIniPath As String, _
                                                        Optional ByVal strNoramlBreak As String, Optional ByVal strBeforeBreak As String, _
                                                        Optional ByRef strMaxVer As String, Optional ByRef strCurMaxVer As String, Optional ByVal strBakDB As String, _
                                                        Optional ByVal blnReadByMax As Boolean) As ADODB.Recordset
'功能：获取升级要执行的文件
'参数：rsUpgradeFiles=升级文件记录集，可能是多个系统的升级文件记录集
'          lngSys=系统号,=-1表示只初始化记录集
'          strIniPath=安装配置文件
'          strBreakVers=升迁配置文件的断点版本
'          strBakDB=历史库用名
'          strMaxVer=最大的版本
'          strCurMaxVer=本次升迁的目标版本
'          blnReadByMax=根据最大版本strMaxVer读取脚本（主要用于系统安装时管理工具版本较低管理工具单独升级时使用）
'                                   该参数为True时，不会进行断点处理，其余和正常应用系统脚本处理一致
'返回:升级文件记录
'        strMaxVer=最终目标版本,即当前脚本所能升迁到的最大打版本
'        strCurMaxVer=本次升迁的目标版本，系统升迁可能由于某些版本不能连续升迁，可能需要分多次升迁在能到最终目标版本。
'                               没有不能连续升迁的版本时,该版本与strMaxVer相同
'说明：
'        strBakDB="":读取所有脚本。此时如下参数含义
'                            strNoramlBreak：在线库（lngSys=0是为管理工具）常规升级中止信息
'                            strBeforeBreak:在线库（lngSys=0是为管理工具）提前升级中止信息
'                            strMaxVer:仅用来返回升迁的最终目标版本
'                            strCurMaxVer:仅用来返回升迁的本次目标版本
'                            返回的文件记录集中高于本次升迁目标版本的脚本全部剔除。
'        strBakDB<>"":读取大于strCurVer并且不大于strMaxVer的脚本。并于生成历史库的脚本文件记录集。
'                             在历史库非单独升迁时，生成的脚本文件记录集，要包括大于应用系统当前版本与应用系统本次目标版本之间的历史库脚本
'                             此时如下参数含义：
'                            strNoramlBreak：历史库常规升级中止信息
'                            strBeforeBreak:历史库提前升级中止信息
'                            strMaxVer:在线库的当前版本
    Dim rsCurFiles As ADODB.Recordset, arrFields As Variant, blnNew As Boolean
    Dim strCurPriFull As String, strCurFull As String, strMaxFull As String, strMaxPriFull As String
    Dim cllFolder As New Collection, objFolder As Folder, objFile As File
    Dim strBreak As String, strTmp As String, arrTmp As Variant, strFilter As String
    Dim strFileVer As String, stFile As SysType, ftFile As FileType, vtFile As VersionType, strSetupVer As String, blnSpecial As Boolean
    Dim strFileNameRule As String, stJudge As SysType
    Dim cllSuppleVers As New Collection, Item As Variant
    Dim i As Long
    Dim strFirstBreak As String, strSecdBreak As String
    Dim strBaseSupple As String
    
    On Error GoTo errh
    
    strCurPriFull = VerFull(GetPrimaryVer(strCurVer))
    strCurFull = VerFull(strCurVer)
    strMaxFull = VerFull(strMaxVer, True) '空串会生成9999.9999.9999.9999
    strMaxPriFull = VerFull(GetPrimaryVer(strMaxFull)) '防止空串生成失败，因此不用strMaxVer生成
    If rsUpgradeFiles Is Nothing Then
        blnNew = True
    ElseIf rsUpgradeFiles.State = adStateClosed Then
        blnNew = True
    End If
    
    If blnNew Or lngSys = -1 Then
        '配置版本:对提前执行脚本为最低要求版本，对应应用系统在线库普通升级脚本为对应管理工具脚本
        Set rsUpgradeFiles = CopyNewRec(Nothing, True, , _
                                                                Array("系统编号", adInteger, 5, Empty, "所有者", adVarChar, 100, Empty, "SysType", adInteger, 1, Empty, _
                                                                        "FileName", adVarChar, 50, Empty, "FilePath", adVarChar, 1000, Empty, "FileType", adInteger, 1, Empty, _
                                                                        "SPVer", adVarChar, 20, Empty, "FullSPVer", adVarChar, 20, Empty, "VerType", adInteger, 1, Empty, _
                                                                        "Optional", adVarChar, 2000, Empty, "AbortLine", adInteger, 10, Empty, "Special", adInteger, 1, Empty, _
                                                                        "配置版本", adVarChar, 20, Empty, "断点", adInteger, 1, Empty))
    End If
    If lngSys = -1 Then Set GetUpgradeFiles = rsUpgradeFiles: Exit Function
    '读取当前系统的脚本
    rsUpgradeFiles.Filter = "系统编号=" & lngSys & IIf(strBakDB <> "", " And 所有者='" & UCase(strBakDB) & "'", "")
    '脚本已经存在，则不用重新读取。
    '历史库读取，必须最大版本不为空。因为历史库单独升迁的目标版本为在线库当前版本。非单独升级时，在线库当前版本之上的历史脚本已经读取
    If Not rsUpgradeFiles.EOF Or strBakDB <> "" And strMaxVer = "" Then Set GetUpgradeFiles = rsUpgradeFiles: Exit Function
    Set rsCurFiles = CopyNewRec(rsUpgradeFiles, strBakDB = "")
    '////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////          1、升迁文件读取            ///////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    '获取需要搜集脚本的文件夹
    If lngSys = 0 Then
        cllFolder.Add gobjFile.GetFile(strIniPath).ParentFolder
        strFileNameRule = "ZLUPGRADE*.*.*.SQL"
    Else
        strFileNameRule = "ZL" & lngSys \ 100 & "_*.*.*.SQL"
        For Each objFolder In gobjFile.GetFolder(gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(strIniPath)) & "\升级脚本\").SubFolders
            If IsVerSion(objFolder.name) And objFolder.name Like "*.*.0" Then
                If VerFull(objFolder.name) >= strCurPriFull And VerFull(objFolder.name) <= strMaxPriFull Then
                    cllFolder.Add objFolder
                End If
            End If
        Next
    End If
    arrFields = Array("系统编号", "SysType", "FileName", "FilePath", "FileType", "SPVer", "FullSPVer", "VerType", "Special", "配置版本")
    '遍历,提取文件
    For Each objFolder In cllFolder
        If lngSys <> 0 And strBakDB = "" Then '获取zlUpgrade.ini
            '获取有效的断点版本
            strTmp = GetUpgradeIniBreak(objFolder.Path & "\zlUpgrade.ini", IIf(VerFull(objFolder.name) >= strCurPriFull, strCurVer, objFolder.name), GetPrimaryVer(objFolder.name, True))
            If strTmp <> "" Then
                strBreak = strBreak & "," & strTmp
            End If
        End If
        '获取文件
        For Each objFile In objFolder.Files
            If UCase(objFile.name) Like strFileNameRule Then '符合文件的规则的才进行名称解析
                If AnalysisFileName(objFile.name, lngSys, strFileVer, ftFile, stFile, vtFile, blnSpecial) Then
                    If VerFull(strFileVer) > strCurFull And VerFull(strFileVer) <= strMaxFull Then
                        If vtFile = VT_Supple Then
                            On Error Resume Next
                            '确认该大版本已经标记的补充版本
                            strBaseSupple = cllSuppleVers("K_" & GetPrimaryVer(strFileVer))
                            If err.Number <> 0 Then
                                err.Clear
                                cllSuppleVers.Add strFileVer, "K_" & GetPrimaryVer(strFileVer)
                            '已经标记的补充版本小于当前版本，则讲标记修改为当前版本
                            ElseIf VerFull(strBaseSupple) > VerFull(strFileVer) Then
                                cllSuppleVers.Remove "K_" & GetPrimaryVer(strFileVer)
                                cllSuppleVers.Add strFileVer, "K_" & GetPrimaryVer(strFileVer)
                            End If
                            On Error GoTo errh
                        End If
                        '获取配置版本
                        If ftFile = FT_Before Or ftFile = FT_Standard And stFile = ST_App And VerFull(strFileVer) > VerFull("10.32.0") Then
                            arrTmp = Split(GetUpgradeCtrolInfo(objFile.Path, ftFile = FT_Before) & "|", "|")
                            strSetupVer = VerFull(arrTmp(IIf(ftFile = FT_Before, 0, 1))) '扩充为标准版本，方便比较;    提前执行返回：最低要求版本，常规升级脚本返回：连续升级|对应管理工具版本
                            '10.34.0之后，管理工具，应用系统版本已经一一对应，且没有脚本的版本用空文件放置
                            If ftFile = FT_Standard Then
                                 If VerFull(strFileVer) >= VerFull("10.34.0") Then
                                    strSetupVer = VerFull(strFileVer) '扩充为标准版本，方便比较
                                ElseIf strSetupVer = VerFull("0") Then  '读取应用对应工具版本失败，则自动生成一个
                                    strSetupVer = VerFull(GetContractVersion(strFileVer, True))
                                End If
                            End If
                            If Val(arrTmp(0)) <> 1 And ftFile = FT_Standard And strBakDB = "" Then strBreak = strBreak & "," & strFileVer
                        Else
                            strSetupVer = ""
                        End If
                        rsCurFiles.AddNew arrFields, Array(lngSys, stFile, objFile.name, objFile.Path, ftFile, strFileVer, VerFull(strFileVer), vtFile, IIf(blnSpecial, 1, 0), strSetupVer)
                    End If
                End If
            End If
        Next
    Next
    '////////////////////////////////////////////////////////////////////////////////////
    '////////////////////   2.上次升迁信息的剔除，补充版本断点标记  ///////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    '标记补充版本
    For Each Item In cllSuppleVers
        '大于该大版本的最小的补充版本，且小余下一个版本
        Call RecUpdate(rsCurFiles, "FullSPVer>='" & VerFull(Item) & "' And FullSPVer<'" & VerFull(GetPrimaryVer(Item, True)) & "'", "VerType", VT_Supple)
    Next
    stJudge = IIf(lngSys = 0, ST_Tools, IIf(strBakDB = "", ST_App, ST_History))
    strFilter = "SysType=" & stJudge & " And FileType<>" & FT_Deferred
    '剔除提前中止语句之前的文件
    arrTmp = Split(strBeforeBreak & "||", "|")
    '没有中止文件，则小于等于中止版本的提前执行脚本都要删除，否则，只删除小于中止版本的提前脚本
    Call RecDelete(rsCurFiles, strFilter & " And FileType=" & FT_Before & " And FullSPVer<" & IIf(arrTmp(1) = "", "=", "") & "'" & VerFull(arrTmp(0)) & "'")
    If arrTmp(1) <> "" Then '有中止文件，记录中止点
        Call RecUpdate(rsCurFiles, strFilter & "And FileType=" & FT_Before & " And SPVer='" & arrTmp(0) & "'", "AbortLine", Val(arrTmp(2)))
    End If
    arrTmp = Split(strNoramlBreak & "||", "|")
    '剔除正常中止语句之前的文件
    Call RecDelete(rsCurFiles, strFilter & " And FullSPVer<" & IIf(arrTmp(1) = "", "=", "") & "'" & VerFull(arrTmp(0)) & "'")
    If arrTmp(1) <> "" Then '有中止文件
        '删除中止中止版本中执行顺序在中止文件之前的文件
        Call RecDelete(rsCurFiles, strFilter & " And SPVer='" & arrTmp(0) & "' And FileType<" & Val(arrTmp(1)))
        '记录中止点
        Call RecUpdate(rsCurFiles, strFilter & " And SPVer='" & arrTmp(0) & "' And FileType=" & Val(arrTmp(1)), "AbortLine", Val(arrTmp(2)))
    End If
    '不能连续升迁版本的标记
    strBreak = Mid(strBreak, 2): arrTmp = Split(strBreak, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        Call RecUpdate(rsCurFiles, "SPVer='" & arrTmp(i) & "'", "断点", 1)
    Next
    '测试脚本执行，则不剔除补充版本与不需要的特殊SP
    If Not gblnTestUpgrade Then
        '剔除补充版本。按版本排序，第一个非补充版本之前的所有补充版本全部删掉。
        rsCurFiles.Filter = "VerType=" & VT_Normal: rsCurFiles.Sort = "FullSPVer Desc"
        If Not rsCurFiles.EOF Then Call RecDelete(rsCurFiles, "VerType=" & VT_Supple & " And FullSPVer<'" & rsCurFiles!FullSPVer & "'")
    
        '剔除特殊SP脚本。按版本排序，第一个非特殊SP版本之前的所有特殊SP全部删掉
        '这种判断有个问题，可能一个版本没有正式脚本，但是有特殊SP脚本，因此不按这种处理。
        rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
        If Not rsCurFiles.EOF Then
            strTmp = VerFull(VerSpecialNormal(rsCurFiles!SPVer))
            Call RecDelete(rsCurFiles, "Special=1 And FullSPVer<'" & strTmp & "'")
        End If
    End If
    '////////////////////////////////////////////////////////////////////////////////////
    '/////////////// 3、最终目标版本、本次目标版本、以及历史库脚本的读取 ////////////
    '///////////////////////////////////////////////////////////////////////////////////
    If strBakDB = "" Then
        If blnReadByMax Then '根据最大版本读取
            '获取实际可以升级到的最大版本
            rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
            strCurMaxVer = ""
            If Not rsCurFiles.EOF Then
                strCurMaxVer = rsCurFiles!SPVer & ""
            End If
        Else
            '获取最终目标版本以及本次目标版本
            rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
            strMaxVer = "": strCurMaxVer = ""
            If Not rsCurFiles.EOF Then
                strMaxVer = rsCurFiles!SPVer & ""
                rsCurFiles.Filter = "断点=1": rsCurFiles.Sort = "FullSPVer"
                If Not rsCurFiles.EOF Then
                    strFirstBreak = rsCurFiles!SPVer
                    If rsCurFiles.RecordCount > 1 Then
                        rsCurFiles.MoveNext: strSecdBreak = rsCurFiles!SPVer
                    End If
                    rsCurFiles.Filter = "FullSPVer<'" & VerFull(strFirstBreak) & "'"
                    strCurMaxVer = IIf(rsCurFiles.EOF, strSecdBreak, strFirstBreak)
                End If
            End If
            If strCurMaxVer = "" Then
                strCurMaxVer = strMaxVer
            Else '删除不需要本次升迁不需要执行的脚本
                Call RecDelete(rsCurFiles, "FullSPVer>'" & VerFull(strCurMaxVer) & "'")
            End If
        End If
    Else
    '获取历史库升迁记录
        '删除小于历史库当前版本的脚本（历史库版本可能高于在线库，因此需要这样处理）
        Call RecDelete(rsCurFiles, "FullSPVer<='" & VerFull(strCurVer) & "'")
        '删除在线库脚本
        Call RecDelete(rsCurFiles, "SysType<>" & ST_History)
        '更新文件记录集的所有者
        Call RecUpdate(rsCurFiles, "", "所有者", UCase(strBakDB))
    End If
    '合并记录集，将本次读取的文件合并到所有记录集中
    rsCurFiles.Filter = ""
    Call RecDataAppend(rsUpgradeFiles, rsCurFiles)
    Set GetUpgradeFiles = rsUpgradeFiles
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function FormatUpgradeBreak(ByVal lngSys As Long, ByVal strResultVer As String, Optional ByVal strUpgradeBreak As String) As String
'功能：解析中止信息，将中止语句标准化 格式：文件版本|文件类型|出错行号
'参数：
'     strResultVer:ZLUpgrade中的结果版本
'     strUpgradeBreak=升迁中止语句
'返回：文件的不带路径的文件名
    Dim arrTmp As Variant
    Dim lngPos As Long
    Dim strTmp As String
    Dim strFileName As String
    Dim lngAbort As Long
    Dim strFileVer As String '从文件名上读取的版本信息
    Dim ftReturn As FileType
    Dim strReturn As String
    
    strReturn = strResultVer & "||"
    If strUpgradeBreak <> "" Then
        '历史库的中止语句可能为版本号
        If Not IsVerSion(strUpgradeBreak) Then
            strUpgradeBreak = strUpgradeBreak & "||"
            arrTmp = Split(strUpgradeBreak, "|")
            If gobjFile.FileExists(arrTmp(0)) Then
                strFileName = gobjFile.GetFileName(arrTmp(0))
            Else '可能是补充版本已经删掉了
                strTmp = StrReverse(arrTmp(0))
                lngPos = InStr(strTmp, "\")
                '截取最后一个\后的内容
                If lngPos <> 0 Then
                    strFileName = StrReverse(Mid(strTmp, lngPos - 1))
                Else
                    strFileName = ""
                End If
            End If
            lngAbort = Val(arrTmp(1))
            If strFileName <> "" Then
                If AnalysisFileName(strFileName, lngSys, strFileVer, ftReturn) Then
                    strReturn = strFileVer & "|" & ftReturn & "|" & lngAbort
                End If
            End If
        Else '历史库提前升级存放的是版本号
            strReturn = strUpgradeBreak & "||"
        End If
    End If
    FormatUpgradeBreak = strReturn
End Function

Public Function GetUpgradeIniBreak(ByVal strFile As String, Optional ByVal strMinVer As String, Optional ByVal strMaxVer As String)
'功能：获取升迁配置文件的断点
'参数：strFile=升迁配置文件路径
'          strMinVer=升迁配置文件目标版本的最小值
'          strMaxVer=升迁配置文件目标版本的最大值
    Dim rsSub As ADODB.Recordset
    Dim strBreakVer As String
    
    If Not gobjFile.FileExists(strFile) Then Exit Function
    Set rsSub = ReadINIToRec(strFile)
    If rsSub Is Nothing Then Exit Function
    rsSub.Filter = "项目='连续升级'" '升级配置文件的目标版本是否能连续升级
    If rsSub.EOF Then Exit Function
    If Val(rsSub!内容 & "") = 1 Then Exit Function '连续升级不用处理
    rsSub.Filter = "项目='目标版本'" '升级配置文件的目标版本
    If rsSub.EOF Then Exit Function
    strBreakVer = Trim(rsSub!内容 & "")
    If Not IsVerSion(strBreakVer) Then Exit Function
    If strMinVer <> "" Then '小于最小版本，则该断点无效
        If VerFull(strBreakVer) <= VerFull(strMinVer) Then Exit Function
    End If
    If strMaxVer <> "" Then '大于最小版本，则该断点无效
        If VerFull(strBreakVer) > VerFull(strMaxVer) Then Exit Function
    End If
    GetUpgradeIniBreak = strBreakVer
End Function

Public Function GetUpgradeCtrolInfo(ByVal strFile As String, Optional ByVal blnBefore As Boolean) As String
'功能：获取文件中的控制信息
'      strFile=进行判断的脚本文件路径
'      blnBefore=文件是否是提起执行脚本
'返回: blnBefore=false: 连续升级|管理工具版本号
'        blnBefore=True: 最低版本号

    Dim objStream As Scripting.TextStream
    Dim strLine As String, arrFind() As Variant, i As Long, strTmp As String, arrTmp As Variant
    Dim strContinue As String, strToolVer As String, strBreakVer As String, strReqVer As String
    Dim rsSub As ADODB.Recordset
    
    On Error GoTo errh
    
    Set objStream = gobjFile.OpenTextFile(strFile, ForReading)
    If blnBefore Then
        arrFind = Array("[[]最低版本号[]]")
    Else
        arrFind = Array("[[]连续升级[]]", "[[]管理工具版本号[]]")
    End If
    Do While Not objStream.AtEndOfStream
        strLine = TrimEx(objStream.ReadLine, True)
        If strLine Like "--" & arrFind(i) & "*" Then
            strTmp = Trim(Mid(strLine, Len("--" & arrFind(i)) - 4 + 1))
            If Not blnBefore Then
                If i = 0 Then
                    strContinue = strTmp
                Else
                    strToolVer = strTmp
                End If
            Else
                strReqVer = strTmp
            End If
        End If
        If i = UBound(arrFind) Then Exit Do
        i = i + 1
    Loop
    objStream.Close
    
    If blnBefore Then
        GetUpgradeCtrolInfo = Trim(strReqVer)
    Else
        If Trim(strContinue) = "" Then strContinue = "1"
        GetUpgradeCtrolInfo = Trim(strContinue) & "|" & Trim(strToolVer)
    End If
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
'    Debug.Print err.Source & "\" & Me.name & "\GetCtrolInfo:" & err.Description
End Function


Public Function AnalysisFileName(ByVal strFileName As String, ByVal lngSys As Long, Optional ByRef strVersion As String, Optional ByRef ftReturn As FileType, _
                                                        Optional ByRef stReturn As SysType, Optional ByRef vtReturn As VersionType = VT_Normal, Optional ByRef blnSpecial As Boolean) As Boolean
'功能:t通过文件名获取文件信息
'参数：
'   strFile=不包含路径的文件名,带扩展名
'   lngSys=系统号
'返回:
'       True=成功获取，False=获取失败（文件不是系统升级脚本）
'       strVerReturn=文件版本
'       ftReturn=文件类型
'       stReturn=系统类型
'       vtReturn=版本类型
    Dim strSysString As String, strSuffix As String
    Dim arrVer As Variant
    vtReturn = VT_Normal
    blnSpecial = False
    strVersion = ""
    ftReturn = FT_Before
    stReturn = ST_Tools
    If Not UCase(strFileName) Like "*.SQL" Then Exit Function
    strFileName = UCase(Left(strFileName, Len(strFileName) - 4))
    arrVer = Split(strFileName, ".")
    '版本文件的文件名仅有2个句点号(特殊SP包含3个）
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    '获取脚本系统前缀
    If arrVer(0) Like "ZLUPGRADE*" Then
        strSysString = "ZLUPGRADE"
        stReturn = ST_Tools
    ElseIf arrVer(0) Like "ZL" & lngSys \ 100 & "_*" Then
        strSysString = "ZL" & lngSys \ 100 & "_"
        stReturn = ST_App
    Else
        Exit Function '没有系统标识前缀，不是系统脚本
    End If
    '系统标识后面紧跟的是版本
    arrVer(0) = Mid(arrVer(0), Len(strSysString) + 1) '获取主板本
    arrVer(UBound(arrVer)) = GetPrefixNumber(arrVer(UBound(arrVer)), strSuffix) '获取次级版本
    '获取的主板本，大版本以及次级版本若不为数字，则退出
    If Not IsNumeric(arrVer(0)) Or Not IsNumeric(arrVer(1)) Or Not IsNumeric(arrVer(2)) Or Not IsNumeric(arrVer(UBound(arrVer))) Then Exit Function
    strVersion = arrVer(0) & "." & arrVer(1) & "." & arrVer(2) & IIf(UBound(arrVer) = 2, "", "." & arrVer(UBound(arrVer)))
    If Not IsVerSion(strVersion) Then Exit Function
    '四位版本号就是特殊SP
    blnSpecial = strVersion Like "*.*.*.*"
    '版本后是文件类型信息
    If stReturn = ST_App And strSuffix Like "_HISTORY*" Then
        stReturn = ST_History
        strSuffix = Mid(strSuffix, Len("_HISTORY") + 1)
    End If
    If strSuffix Like "*(补充)" Then
        vtReturn = VT_Supple
        strSuffix = Replace(strSuffix, "(补充)", "") '防止补充信息位置不固定
    End If
    Select Case strSuffix
        Case ""
            ftReturn = FT_Standard
        Case "_DBA"
            If stReturn = ST_History Then Exit Function '历史库不支持DBA脚本
            ftReturn = FT_DBA
        Case "_OPTIONAL"
            ftReturn = FT_Optional
        Case "_BEFORE"
            ftReturn = FT_Before
        Case "_DEFERRED"
            If stReturn = ST_Tools Then Exit Function '管理工具不支持延迟执行脚本
            ftReturn = FT_Deferred
        Case Else '不再命名规则范围内的，则获取失败
            Exit Function
    End Select
    AnalysisFileName = True
End Function

Public Function GetPrefixNumber(ByVal strInput As String, Optional ByRef strOther As String) As String
'功能：获取一个字符串的数字前缀，以及剩余部分
'参数：strInput=输入的字符串
'          strOther =去掉数字前缀的剩余部分
    Dim i As Long
    
    For i = 1 To Len(strInput)
        If Not IsNumeric(Mid(strInput, i, 1)) Then
            Exit For
        End If
    Next
    strOther = Mid(strInput, i)
    GetPrefixNumber = Mid(strInput, 1, i - 1)
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'功能：返回VB最大支持的版本号形式:9999.9999.9999.9999,最小版本号0000.0000.0000.0000
'参数：strVer=当前版本号
'           blnMax=True,若果为空，则返回最大支持版本，False=若果为空，则返回最小支持版本
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '增加一段，以兼容特殊SP版本号
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

Public Function VerPAD(ByVal strVer As String) As String
'功能：使版本号的主版本号左填充为4位，保证主版本后原点可以与其他版本号对齐
'参数：strVer=当前版本号
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        Exit Function
    End If
    arrVer = Split(strVer & ".", ".")
    VerPAD = RPAD(Lpad(arrVer(0), 2) & "." & arrVer(1) & "." & arrVer(2) & IIf(Val(arrVer(3)) = 0, "", "." & Format(Val(arrVer(3)), "0000")), 20)
End Function

Public Function GetPrimaryVer(ByVal strVer As String, Optional ByVal blnNext As Boolean)
'功能：获取一个版本的主版本
'参数：strVer=当前版本
'          blnNext=是否获取下一个主版本
'返回：主版本
    Dim arrVer As Variant
    
    arrVer = Split(strVer & "..", ".")
    If blnNext Then
        GetPrimaryVer = Val(arrVer(0)) & "." & (Val(arrVer(1)) + 1) & "." & 0
        '管理工具没有9.45.0，直接和应用系统同一编号，为10.34.0
        If GetPrimaryVer = "9.45.0" Then GetPrimaryVer = "10.34.0"
    Else
        GetPrimaryVer = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & 0
    End If
End Function

Public Function GetContractVersion(ByVal strVer As String, Optional ByVal blnGetTools As Boolean = True)
'功能：获取应用系统对应管理工具的主版本，或者管理工具对应应用系统版本，需要推算
'参数：strVer=当前应用系统版本
'          blnGetTools=True-获取对应的管理工具版本,False-获取对应的应用系统版本
'返回：对应版本，应用系统10.34.0之前，只求对应大版本，不具体到SP版本
'                          管理工具10.34.0之前，只求对应大版本，不具体到SP版本
    Dim arrVer As Variant
    Dim lngDistance As Long
    If strVer = "" Then strVer = "9.1.0"
    If blnGetTools Then
        If VerFull(strVer) >= VerFull("10.34.0") Then '10.34.0  以后管理工具和应用系统版本统一
            GetContractVersion = strVer
        Else
            arrVer = Split(strVer & "...", ".")
            lngDistance = 33 - Val(arrVer(1)) '获取应用系统与10.33.0版本的大版本间隔
            '管理工具9.44.0减去相应大版本间隔就为对应管理工具版本
            GetContractVersion = "9." & (44 - lngDistance) & ".0"
        End If
    Else
        If VerFull(strVer) >= VerFull("10.34.0") Then  '  以后管理工具和应用系统版本统一
            GetContractVersion = strVer
        Else
            arrVer = Split(strVer & "...", ".")
            lngDistance = 44 - Val(arrVer(1)) '获取管理工具与9.44.0版本的大版本间隔
            '应用系统10.33.0减去相应大版本间隔就为对应应用系统的版本
            GetContractVersion = "10." & (33 - lngDistance) & ".0"
        End If
    End If
End Function

Public Function VerNormal(ByVal strVer As String) As String
'功能：将VB最大支持的版本号形式:9999.9999.9999转换为常见版本虚形式，如0010.0034.0000.0000，转换为10.34.0
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then Exit Function
    arrVer = Split(strVer & ".", ".")
    VerNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2)) & IIf(Val(arrVer(3)) = 0, "", "." & Format(Val(arrVer(3)), "0000"))
End Function

Public Function VerSpecialNormal(ByVal strVer As String) As String
'获取一个特殊sp对应的正式版本，如果是一个正式版本，则返回其自身
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then Exit Function
    arrVer = Split(strVer & ".", ".")
    VerSpecialNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2))
End Function

Public Function IsVerSion(ByVal strVer As String) As Boolean
'功能：判断字符串是否是版本号
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If i = 3 Then
            If Format(Val(arrVer(i)), "0000") <> Format(Trim(arrVer(i)), "0000") Then Exit Function
        Else
            If Val(arrVer(i)) & "" <> Trim(arrVer(i)) Then Exit Function
        End If
    Next
    
    IsVerSion = True
End Function

Public Function ReadHisUpgrade(ByVal cnHistory As ADODB.Connection, ByVal strOwner As String, Optional ByVal blnMsg As Boolean, Optional ByVal lngSys As Long, Optional ByVal blnDB_LINK As Boolean) As ADODB.Recordset
'功能:获取历史表空间的各系统信息升迁信息
'参数： cnHistory=历史库连接
'           strOwner=所有者
'           lngSys=系统编号=0：获取该历史库的所有系统，<>0:仅获得该系统历史库
'           blnDB_LINK=是否是DBLINK连接
    Dim rsReturn As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim objTmp As New clsObjectInfo, objParent As clsObjectInfo, objCur As clsObjectInfo
    Dim strSQL As String
    
    On Error GoTo errh
    Set rsReturn = CopyNewRec(Nothing, True, , Array("系统编号", adInteger, Empty, Empty, "当前版本", adVarChar, 20, Empty, "中止信息", adVarChar, 2000, Empty, _
                                                                                    "提前中止信息", adVarChar, 2000, Empty, "提前执行", adInteger, 1, Empty))
    'zlbakinfo检查
    Set objParent = objTmp.GetObject("zlbakinfo", OT_Table, , strOwner, , cnHistory)
    '提前执行列检查
    Set objCur = objTmp.GetObject("提前执行", OT_Column, "alter Table zlbakinfo add 提前执行 number(1)", strOwner, objParent, cnHistory)
    If Not objCur.ObjectCheck(blnMsg) Then
        GoTo ExitCode
    End If
    '提前中止语句列检查
    Set objCur = objTmp.GetObject("提前中止语句", OT_Column, "alter Table zlbakinfo add 提前中止语句 VarChar2(500)", strOwner, objParent, cnHistory)
    If Not objCur.ObjectCheck(blnMsg) Then
        GoTo ExitCode
    End If
'    '创建ZLBAKInfo视图
'    strSQL = "create or replace view " & strOwner & ".zlbakinfo as" & vbNewLine & _
'        "Select ""系统"",""版本号"",""更新日期"",""最后转储日期"",""最后复制日期"",""中止语句"",""提前执行"",""提前中止语句"" From " & strOwner & ".ZLBAKINFO"
'    cnHistory.Execute strSQL
    '授权
    If Not blnDB_LINK Then
        strSQL = "Grant Select On  " & strOwner & ".zlbakinfo To " & gstrUserName
        cnHistory.Execute strSQL
    End If
    '生成记录集信息
    strSQL = "Select 系统,版本号,中止语句,提前执行,提前中止语句  from zlbakinfo " & IIf(lngSys = 0, "", "Where 系统=" & lngSys) & " order by 系统"
    Set rsTmp = gclsBase.OpenSQLRecord(cnHistory, strSQL, "获取历史库中系统信息")
    Do While Not rsTmp.EOF
        rsReturn.AddNew Array("系统编号", "当前版本", "中止信息", "提前中止信息", "提前执行"), _
                                    Array(rsTmp!系统, rsTmp!版本号, FormatUpgradeBreak(rsTmp!系统, rsTmp!版本号 & "", rsTmp!中止语句 & ""), FormatUpgradeBreak(rsTmp!系统, rsTmp!版本号 & "", rsTmp!提前中止语句 & ""), rsTmp!提前执行)
        rsTmp.MoveNext
    Loop
    Set ReadHisUpgrade = rsReturn
    Exit Function
errh:
    Set ReadHisUpgrade = rsReturn
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
    Exit Function
ExitCode:
    If 0 = 1 Then
        Resume
    End If
    Set ReadHisUpgrade = rsReturn
End Function

Public Function CheckHavHistory(ByVal lngSys As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:检查是否需要创建历史空间（即存在待转出表）
    '参数:lngSys-系统号
    '返回:需要创建,返true,否则False
    '--------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    strSQL = "Select 1 from zltools.zlbakTables where 系统=[1] and rownum<=1"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取bak数据", lngSys)
    If rsTmp.EOF Then
       '返回False,表示该系统没有历史数据空间,没有要处理历史数据空间
       Exit Function
    End If
    CheckHavHistory = True
End Function

Public Function GrantBakToUser(ByVal cnOracle As ADODB.Connection, ByVal strToOwner As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    '功能:检查表是否存在
    '参数:strTableName-表名
    '     cnoracle-数据库连接名
    '     strOwNer-所有者
    '返回:存在该表返回true,否则False
    '-----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    err = 0: On Error GoTo ErrHand:
    strSQL = "Select TABLE_NAME from user_all_tables" & _
            " Union All Select View_Name From User_Views"
    Call OpenRecordset(rsTemp, strSQL, "重新授权", , , cnOracle)
    With rsTemp
        Do While Not .EOF
            strSQL = "Grant ALL on " & Nvl(!Table_Name) & " to " & strToOwner & " With Grant Option"
            cnOracle.Execute strSQL
            .MoveNext
        Loop
    End With
    GrantBakToUser = True
    Exit Function
ErrHand:
    If MsgBox("在授权时出现如下错误,请检查!" & vbCrLf & " (" & err.Number & ") " & err.Description, vbRetryCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbRetry Then
        Resume
    End If
    GrantBakToUser = False
End Function

Public Sub ReGrantToRole(ByVal cnOracle As ADODB.Connection, ByVal strRoleNames As String, ByVal blnGrantBase As Boolean, strOwners() As String, Optional ByRef objProcess As Object, Optional ByRef objlblPer As Object, Optional ByRef lngRoleCount As Long)
'功能：对角色进行重新授权。
'参数：cnOracle=连接
'      strRoleNames=重新授权的角色。若为空，则为所有角色重新授权。不为空时，多个角色以都好分割，角色不超过15个。
'      blnGrantBase=是否授予字典管理工具权限
'      strOwners=授权的系统的所有者
'      objProcess=进度
    Dim rsPrivs As ADODB.Recordset, rsRoles As ADODB.Recordset
    Dim strRolePars As String
    Dim strSQL As String, i As Long
    Dim lngMax As Long, lngCur As Long
    Dim blnProcess As Boolean
    
    On Error GoTo errh
    blnProcess = Not objProcess Is Nothing
    If strRoleNames = "" Then
    
        '以前的SQL经测试查询到47811条数据,现在经优化，查询到3057条数据，加Distinct主要是一个角色有多个功能，功能访问的表之间有重叠
        '角色数：31，版本：10.29.30，优化前整个角色授权耗时138秒，优化后19秒
        strSQL = "Select 权限, 对象, 所有者, f_List2str(Cast(Collect(角色) As t_Strlist)) As 角色" & vbNewLine & _
                "From (Select 对象, 所有者, 权限, 角色, Floor(Row_Number() Over(Partition By 权限, 对象, 所有者 Order By 角色) / 10) Rn" & vbNewLine & _
                "       From (Select Distinct Upper(p.对象) 对象, p.所有者, Upper(p.权限) 权限, r.角色" & vbNewLine & _
                "              From zlProgPrivs P, zlRoleGrant R, User_Role_Privs U" & vbNewLine & _
                "              Where Nvl(P.系统, 0) = Nvl(R.系统, 0) And P.所有者 = User And P.序号 = R.序号 And P.功能 = R.功能 And R.角色 = U.Granted_Role))" & vbNewLine & _
                "Group By 权限, 对象, 所有者, Rn"
    '    '原来SQL
    '    strSQL = "Select P.对象, P.所有者, P.权限, R.角色" & vbNewLine & _
    '            "From zlProgPrivs P, zlRoleGrant R, User_Role_Privs U" & vbNewLine & _
    '            "Where Nvl(P.系统, 0) = Nvl(R.系统, 0) And P.所有者 = User And P.序号 = R.序号 And P.功能 = R.功能 And R.角色 = U.Granted_Role"
        Set rsPrivs = gclsBase.OpenSQLRecord(cnOracle, strSQL, "角色授权")
        
        strSQL = "Select F_List2str(Cast(Collect(角色) As T_Strlist)) 角色, 角色数" & vbNewLine & _
                "From (Select Floor(Rownum / 10) Rn, 角色, Count(1) Over(Partition By 计数) 角色数" & vbNewLine & _
                "       From (Select Distinct R.角色, 1 计数" & vbNewLine & _
                "              From zlRoleGrant R" & vbNewLine & _
                "              Where Exists" & vbNewLine & _
                "               (Select 1" & vbNewLine & _
                "                     From zlProgPrivs P" & vbNewLine & _
                "                     Where Nvl(P.系统, 0) = Nvl(R.系统, 0) And P.所有者 = User And P.序号 = R.序号 And P.功能 = R.功能) And Exists" & vbNewLine & _
                "               (Select 1 From User_Role_Privs U Where R.角色 = U.Granted_Role)" & vbNewLine & _
                "              Order By R.角色))" & vbNewLine & _
                "Group By Rn, 角色数"
'    '原来SQL
'    strSQL = "Select Distinct R.角色" & vbNewLine & _
'            "From zlProgPrivs P, zlRoleGrant R, User_Role_Privs U" & vbNewLine & _
'            "Where Nvl(P.系统, 0) = Nvl(R.系统, 0) And P.所有者 = User And P.序号 = R.序号 And P.功能 = R.功能 And R.角色 = U.Granted_Role" & vbNewLine & _
'            "Order By R.角色"
        Set rsRoles = gclsBase.OpenSQLRecord(cnOracle, strSQL, "角色授权")
    Else
        strRolePars = "'" & Replace(UCase(strRoleNames), ",", "','") & "'"
        strSQL = "Select 权限, 对象, 所有者, f_List2str(Cast(Collect(角色) As t_Strlist)) As 角色" & vbNewLine & _
                "From (Select 对象, 所有者, 权限, 角色, Floor(Row_Number() Over(Partition By 权限, 对象, 所有者 Order By 角色) / 10) Rn" & vbNewLine & _
                "       From (Select Distinct Upper(p.对象) 对象, p.所有者, Upper(p.权限) 权限, r.角色" & vbNewLine & _
                "              From Zlprogprivs p, Zlrolegrant r" & vbNewLine & _
                "              Where Nvl(p.系统, 0) = Nvl(r.系统, 0) And p.所有者 = User And p.序号 = r.序号 And p.功能 = r.功能 And r.角色 in(" & strRolePars & ")))" & vbNewLine & _
                "Group By 权限, 对象, 所有者, Rn"
        Set rsPrivs = gclsBase.OpenSQLRecord(cnOracle, strSQL, "角色授权")
        
        strSQL = "Select f_List2str(Cast(Collect(角色) As t_Strlist)) 角色, 角色数" & vbNewLine & _
                "From (Select Floor(Rownum / 10) Rn, 角色, Count(1) Over(Partition By 计数) 角色数" & vbNewLine & _
                "       From (Select Distinct r.角色, 1 计数" & vbNewLine & _
                "              From Zlrolegrant r" & vbNewLine & _
                "              Where Exists (Select 1" & vbNewLine & _
                "                     From Zlprogprivs p" & vbNewLine & _
                "                     Where Nvl(p.系统, 0) = Nvl(r.系统, 0) And p.所有者 = User And p.序号 = r.序号 And p.功能 = r.功能) And" & vbNewLine & _
                "                    r.角色 In (" & strRolePars & ")" & vbNewLine & _
                "              Order By r.角色))" & vbNewLine & _
                "Group By Rn, 角色数"
        Set rsRoles = gclsBase.OpenSQLRecord(cnOracle, strSQL, "角色授权")
    End If
    On Error Resume Next
    lngMax = rsPrivs.RecordCount + 25 * rsRoles.RecordCount
    For lngCur = 1 To rsPrivs.RecordCount
        If blnProcess Then
            objProcess.value = lngCur / lngMax * 100
            objlblPer.Caption = Format(objProcess.value / 100, "0%")
        End If
        DoEvents
        cnOracle.Execute "Grant " & rsPrivs!权限 & " on " & rsPrivs!所有者 & "." & rsPrivs!对象 & " to " & rsPrivs!角色
        rsPrivs.MoveNext
    Next
    If rsRoles.RecordCount <> 0 Then
        lngRoleCount = Val(rsRoles!角色数 & "")
    End If
    For i = 1 To rsRoles.RecordCount
        lngCur = i * 25
        If blnProcess Then
            objProcess.value = lngCur / lngMax * 100
            objlblPer.Caption = Format(objProcess.value / 100, "0%")
        End If
        DoEvents
        Call GrantSpecialToRole(cnOracle, rsRoles!角色 & "", True, strOwners, True)
        rsRoles.MoveNext
    Next
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub



Public Sub ReGrantForTools(ByVal cnTools As ADODB.Connection, Optional ByVal strSysOwner As String, Optional ByVal blnSysGrant As Boolean)
    '----------------------------------------------------------------------------------------------------------
    '功能:对管理工具的对象进行重新授权并创建同义词
    '参数:cnTools：管理工具连接。strSysOwner为空时，可以传应用系统连接，此时为应用系统转授权限。
    '     strSysOwner:应用系统所有者。为空是服务器创建调用，只创建公共同义词以及对Public授权，非空时对该用户进行ZLTOOLS对象权限授予
    '     blnSysGrant:系统所有者转授管理工具权限，需加前缀ZLTOOLS.,当strSysOwner为空且该参数为True时对所有系统所有者进行授权
    '返回:
    '----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim strSysSQL As String
    
    On Error Resume Next
    '修正同义词缺失的对象，默认此种对象的权限也缺失
    strSQL = "Select Object_Name" & vbNewLine & _
                    "From ((Select Object_Name" & vbNewLine & _
                    "        From User_Objects" & vbNewLine & _
                    "        Where Object_Type In ('FUNCTION', 'PROCEDURE', 'TYPE', 'PACKAGE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
                    "              Instr(Object_Name, 'BIN$') <= 0) Minus" & vbNewLine & _
                    "       (Select Synonym_Name From All_Synonyms Where Owner = 'PUBLIC' And Table_Owner = 'ZLTOOLS'))"
    Call OpenRecordset(rsTemp, strSQL, "管理工具公共同义词缺失修正", , , cnTools)
    For i = 1 To rsTemp.RecordCount
        cnTools.Execute "Create Public Synonym " & rsTemp!Object_Name & " For ZLTOOLS." & rsTemp!Object_Name
        If err.Number <> 0 Then '可能存在其他用户的同义词，必须优先ZLtools,因此删掉。
            err.Clear
            cnTools.Execute "Drop Public Synonym " & rsTemp!Object_Name
            cnTools.Execute "Drop  Synonym " & rsTemp!Object_Name
            cnTools.Execute "Create Public Synonym " & rsTemp!Object_Name & " For ZLTOOLS." & rsTemp!Object_Name
            If err.Number <> 0 Then err.Clear
        End If
        rsTemp.MoveNext
    Next
    '修正Public权限缺失
    strSQL = "Select Object_Name,Privilege" & vbNewLine & _
                    "From ((Select Object_Name," & vbNewLine & _
                    "               Decode(Object_Type, 'SEQUENCE', 'SELECT', 'TABLE', 'SELECT', 'VIEW', 'SELECT', 'EXECUTE') Privilege" & vbNewLine & _
                    "        From User_Objects" & vbNewLine & _
                    "        Where Object_Type In ('FUNCTION', 'PROCEDURE', 'TYPE', 'PACKAGE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
                    "              Instr(Object_Name, 'BIN$') <= 0 And" & vbNewLine & _
                    "              Object_Name Not In ('B_ROLEGROUPMGR', 'ZL_ZLROLEGRANT_BATCHDELETE', 'ZL_ZLROLEGRANT_BATCHINSERT')) Minus" & vbNewLine & _
                    "       (Select Table_Name, Privilege" & vbNewLine & _
                    "        From User_Tab_Privs" & vbNewLine & _
                    "        Where Grantee = 'PUBLIC' And Grantor = 'ZLTOOLS' And Instr(Table_Name, 'BIN$') <= 0))"
    Call OpenRecordset(rsTemp, strSQL, "管理工具Public权限缺失修正", , , cnTools)
    For i = 1 To rsTemp.RecordCount
        cnTools.Execute "Grant " & rsTemp!Privilege & " On ZLTOOLS." & rsTemp!Object_Name & " To Public"
        rsTemp.MoveNext
    Next
    
    If err.Number <> 0 Then err.Clear
    
    If strSysOwner <> "" Then
        strSysSQL = "Select '" & Trim(UCase(strSysOwner)) & "' 所有者 FROM Dual"
    ElseIf blnSysGrant Then
        '管理工具对象重新授权并创建公共同义词,对应用系统和历史库都授权
        strSysSQL = "Select Distinct 所有者" & vbNewLine & _
                "From Zlsystems" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select Distinct 所有者" & vbNewLine & _
                "From Zlbakspaces a" & vbNewLine & _
                "Where Exists (Select 1 From All_Users b Where b.Username = Upper(a.所有者))"
    End If
    If strSysSQL <> "" Then
        '应用系统所有者缺失对象权限，简化处理
        strSQL = "Select b.所有者 Grantee, a.Object_Name, 'ALL' Privilege" & vbNewLine & _
                        "From User_Objects a, (" & strSysSQL & ") b" & vbNewLine & _
                        "Where a.Object_Type In ('FUNCTION', 'PROCEDURE', 'TYPE', 'PACKAGE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
                        "      Instr(a.Object_Name, 'BIN$') <= 0" & vbNewLine & _
                        "Minus" & vbNewLine & _
                        "Select Grantee, Table_Name, 'ALL' Privilege" & vbNewLine & _
                        "From User_Tab_Privs a, (" & strSysSQL & ") b" & vbNewLine & _
                        "Where Grantee =b.所有者  And Grantor = 'ZLTOOLS' And Grantable = 'YES' And Instr(Table_Name, 'BIN$') <= 0 And" & vbNewLine & _
                        "      Privilege In ('EXECUTE', 'INSERT')"
        Call OpenRecordset(rsTemp, strSQL, "应用系统所有者对象权限修正", , , cnTools)
        
        For i = 1 To rsTemp.RecordCount
            cnTools.Execute "Grant " & rsTemp!Privilege & " On ZLTOOLS." & rsTemp!Object_Name & " To " & rsTemp!Grantee & " With Grant Option"
            rsTemp.MoveNext
        Next
    End If
    If err.Number <> 0 Then err.Clear: cnTools.Errors.Clear
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Public Function GrantSpecialToRole(ByVal cnOracle As ADODB.Connection, ByVal strRoleNames As String, ByVal blnGrantBase As Boolean, strOwners() As String, Optional ByVal blnCreateRole As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '功能:对管理工具的对象或应用程序一些对象进行授权（特殊的对象）
    '参数:cnOracle：应用系统连接
    '     strRoleNames:被授权的角色，多个角色以逗号分割，一般不超过15个角色
    '     blnGrantBase:是否对应用系统基础表进行授权
    '     strOwners：应用系统所遇者
    '     blnCreateRole=是否创建角色，创建角色才授予公共基础表，否则导致SQL重新解析，所有用户业务停滞
    '返回:
    '----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnsysSt As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errh
    strSQL = "Select 所有者 From zlSystems Where Floor(编号 / 100) = 1"
    If UBound(strOwners) <> -1 And mstrStSysOwner = "" Then
        OpenRecordset rsTmp, strSQL, "获取标准版系统所有者", , , cnOracle
        Do While Not rsTmp.EOF
            mstrStSysOwner = mstrStSysOwner & "," & rsTmp!所有者
            rsTmp.MoveNext
        Loop
        If mstrStSysOwner <> "" Then mstrStSysOwner = mstrStSysOwner & ","
    End If
    On Error Resume Next
    For i = LBound(strOwners) To UBound(strOwners)
        If strOwners(i) <> "" Then
            If blnCreateRole Then
                cnOracle.Execute "grant select on " & strOwners(i) & ".部门表 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".人员表 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".部门人员 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".上机人员表 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".人员性质说明 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".人员性质分类 to " & strRoleNames
            End If
            
            If InStr(mstrStSysOwner, "," & strOwners(i) & ",") > 0 Then
                If blnCreateRole Then
                    '消息平台对象
                    cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息类型 to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息清单 to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息提醒部门 to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息提醒人员 to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息状态 to " & strRoleNames
                    cnOracle.Execute "grant select on " & strOwners(i) & ".三方服务配置目录 to " & strRoleNames
                    cnOracle.Execute "grant execute on " & strOwners(i) & ".Zlpub_业务消息清单_insert to " & strRoleNames
                    cnOracle.Execute "grant execute on " & strOwners(i) & ".Zl_业务消息清单_insert to " & strRoleNames
                    cnOracle.Execute "grant execute on " & strOwners(i) & ".Zl_业务消息清单_read to " & strRoleNames
                End If
            End If
            If blnGrantBase Then
                cnOracle.Execute "grant execute on " & strOwners(i) & ".zl_字典管理_execute to " & strRoleNames
            End If
        End If
    Next
    If err.Number <> 0 Then err.Clear
    On Error GoTo errh
    If blnCreateRole Then
        '对服务器的几个表进行特殊授权
        '------------------------------------------------------------------------------------------------------------------
        '清除本机界面异常
        cnOracle.Execute "grant delete                on ZLTOOLS.zluserparas to " & strRoleNames
        '客户端升级容错处理
        cnOracle.Execute "grant update                on ZLTOOLS.zlclients   to " & strRoleNames
        cnOracle.Execute "grant insert,update         on ZLTOOLS.zlDiaryLog to " & strRoleNames
        cnOracle.Execute "grant insert                on ZLTOOLS.zlErrorLog to " & strRoleNames
        cnOracle.Execute "grant update,delete         on ZLTOOLS.zlMessages to " & strRoleNames
        cnOracle.Execute "grant update,delete         on ZLTOOLS.zlMsgState to " & strRoleNames
        cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientScheme to " & strRoleNames
        cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientParaSet to " & strRoleNames
        cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientparaList to " & strRoleNames
        cnOracle.Execute "grant Select on sys.dba_role_privs to " & strRoleNames
    End If
    GrantSpecialToRole = True
    Exit Function
errh:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function CheckCBOPars() As Boolean
'功能：检查成本计算参数并提供修改功能
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strMsg As String
    Dim cnTmp As ADODB.Connection
    
    On Error GoTo errh
    strSQL = "Select Name,Value,Decode(Name,'optimizer_index_cost_adj','20','80') suggestivevalue" & vbNewLine & _
                    "From V$parameter" & vbNewLine & _
                    "Where Name = 'optimizer_index_cost_adj' And Value = '100' Or Name = 'optimizer_index_caching' And Value = '0'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "成本计算参数检查")
    
    If rsTmp.RecordCount <> 0 Then
        If rsTmp.RecordCount = 1 Then
            strMsg = "数据库参数""" & rsTmp!name & """初始值为" & rsTmp!value & "，可能会" & vbNewLine & _
                            "导致产品性能问题。 建议修改为" & rsTmp!suggestivevalue & "，是否修改？"
        Else
            strMsg = "以下两个数据库参数的初始值可能会导致的产品性能问题：" & vbNewLine & _
                            "   参数""" & rsTmp!name & """初始值为" & rsTmp!value & "，建议修改为" & rsTmp!suggestivevalue & "，"
            rsTmp.MoveNext
            strMsg = strMsg & vbNewLine & _
                            "   参数""" & rsTmp!name & """初始值为" & rsTmp!value & "， 建议修改为" & rsTmp!suggestivevalue & "，" & vbNewLine & _
                            "   是否修改？"
        End If
        If MsgBox(strMsg, vbInformation + vbYesNo, App.Title) = vbYes Then
            If Not gcnSystem Is Nothing Then
                Set cnTmp = gcnSystem
            ElseIf gblnDBA Then
                Set cnTmp = gcnOracle
            Else
                Set cnTmp = GetConnection("SYSTEM")
            End If
            '修正参数
            If Not cnTmp Is Nothing Then
                rsTmp.MoveFirst
                Do While Not rsTmp.EOF
                    strSQL = "alter system set " & rsTmp!name & "= " & rsTmp!suggestivevalue
                    cnTmp.Execute strSQL
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
    CheckCBOPars = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Function

Public Sub CompileAllInvalidObject(ByRef cnThis As ADODB.Connection, ByRef strErrInfor As String, ByRef objPanel As Panel, ByRef objProgressBar As ProgressBar)
'功能：编译指定连接所有者的无效对象
'参数：cnThis=所有者连接,本函数可针对不同所有者调用
'      objPanel=用于显示当前编译的对象名称
'      objProgressBar=用于显示编译进度
    Dim rsObjects As New ADODB.Recordset
    Dim rsDepends As New ADODB.Recordset
    Dim arrObjects As Variant, strCompile As String
    Dim strSQL As String, i As Long
    Dim strUser As String
    
    On Error GoTo errHandle
    strErrInfor = ""
  
    strSQL = _
        "Select User, Object_Name, Object_Type" & vbNewLine & _
        "From User_Objects" & vbNewLine & _
        "Where Object_Type In" & vbNewLine & _
        "      ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
        "      Object_Name Not Like 'BIN$%' And Status = 'INVALID'" & vbNewLine & _
        "Order By Object_Type, Object_Name"

    rsObjects.CursorLocation = adUseClient
    rsObjects.Open strSQL, cnThis, adOpenKeyset '可以看见其他用户更改的数据
    
    objProgressBar.Max = 100
    objProgressBar.value = 0
      
    If Not rsObjects.EOF Then
        strUser = rsObjects!User
        strSQL = _
            "Select Name, Type, Referenced_Name, Referenced_Type" & vbNewLine & _
            "From User_Dependencies" & vbNewLine & _
            "Where Referenced_Owner = User And Type In ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE'," & vbNewLine & _
            "       'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
            "      Referenced_Type In" & vbNewLine & _
            "      ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
            "      Not(Name=Referenced_Name And Type=Referenced_Type) And" & vbNewLine & _
            "      Name Not Like 'BIN$%' And Referenced_Name Not Like 'BIN$%'"
        rsDepends.CursorLocation = adUseClient
        rsDepends.Open strSQL, cnThis, adOpenKeyset '可以看见其他用户更改的数据
        
        ReDim arrObjects(rsObjects.RecordCount - 1) As String
        For i = 1 To rsObjects.RecordCount
            arrObjects(i - 1) = rsObjects!Object_Name & "," & rsObjects!Object_Type
            rsObjects.MoveNext
        Next
        
        '编译无效对象
        DoEvents
        For i = 0 To UBound(arrObjects)
            objPanel.Text = Split(arrObjects(i), ",")(0)    '显示当前对象名称
            objProgressBar.value = (i + 1) / (UBound(arrObjects) + 1) * 100
            DoEvents    '为了刷新进度
            Call CompileInvalidObject(cnThis, Split(arrObjects(i), ",")(0), Split(arrObjects(i), ",")(1), rsObjects, rsDepends, strCompile, strErrInfor)
        Next
    End If
    If strErrInfor <> "" Then strErrInfor = "以下无效对象编译出错:" & vbCrLf & strErrInfor
    
    Exit Sub
    
errHandle: '函数内部的其他未知异常
    If MsgBox(err.Description, vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then Resume
End Sub

Private Sub CompileInvalidObject(ByRef cnThis As ADODB.Connection, ByVal strName As String, ByVal strType As String, _
    ByRef rsObjects As ADODB.Recordset, ByRef rsDepends As ADODB.Recordset, ByRef strCompile As String, ByRef strErrInfor As String)
'功能：编译指定的无效对象
'参数：strCompile=已经编译的对象名串
'说明：CompileAllnvalidObject的子函数
    Dim arrObjRef As Variant, strErr As String
    Dim strSQL As String, i As Long
        
    If InStr(strCompile & ",", "," & strName & ",") > 0 Then Exit Sub
    
    '递归编译当前对象所引用的对象
    rsDepends.Filter = "Name='" & strName & "' And Type='" & strType & "'" '不加类型可能引起递归溢出(同名BODY)
    If Not rsDepends.EOF Then
        ReDim arrObjRef(rsDepends.RecordCount - 1) As String
        For i = 1 To rsDepends.RecordCount
            arrObjRef(i - 1) = rsDepends!Referenced_Name & "," & rsDepends!Referenced_Type
            rsDepends.MoveNext
        Next
        For i = 0 To UBound(arrObjRef)
            rsObjects.Filter = "Object_Name='" & Split(arrObjRef(i), ",")(0) & "' And Object_Type='" & Split(arrObjRef(i), ",")(1) & "'"
            If Not rsObjects.EOF Then '引用对象也是无效对象时
                Call CompileInvalidObject(cnThis, Split(arrObjRef(i), ",")(0), Split(arrObjRef(i), ",")(1), rsObjects, rsDepends, strCompile, strErrInfor)
            End If
        Next
    End If
    '编译当前对象
    Select Case strType
    Case "PROCEDURE"
        strSQL = "ALTER PROCEDURE " & strName & " COMPILE"
    Case "FUNCTION"
        strSQL = "ALTER FUNCTION " & strName & " COMPILE"
    Case "VIEW"
        strSQL = "ALTER VIEW " & strName & " COMPILE"
    Case "MATERIALIZED VIEW"
        strSQL = "ALTER MATERIALIZED VIEW " & strName & " COMPILE"
    Case "TRIGGER"
        strSQL = "ALTER TRIGGER " & strName & " COMPILE"
    Case "PACKAGE"
        strSQL = "ALTER PACKAGE " & strName & " COMPILE"
    Case "PACKAGE BODY"
        strSQL = "ALTER PACKAGE " & strName & " COMPILE BODY"
    Case "TYPE"
        strSQL = "ALTER TYPE " & strName & " COMPILE"
    Case "TYPE BODY"
        strSQL = "ALTER TYPE " & strName & " COMPILE BODY"
    End Select
    If strSQL <> "" Then
        strErr = ""
        err.Clear: On Error Resume Next
        cnThis.Execute strSQL
        If cnThis.Errors.Count > 0 Then
            '特殊情况(未出错):Err.Number=0,NativeError=0
            '[Microsoft][ODBC driver for Oracle]创建的过程或软件包带有编译错误
            '没有更多的结果。
            If Not (cnThis.Errors(0).NativeError = 0 And cnThis.Errors.Count = 1) Then
                If cnThis.Errors(0).NativeError <> 0 Then
                    strErr = cnThis.Errors(0).Description
                    strErrInfor = strErrInfor & vbCrLf & strName & ":" & strErr
                Else
                    strErrInfor = strErrInfor & vbCrLf & strName
                End If
            End If
        End If
        err.Clear: On Error GoTo 0
        strCompile = strCompile & "," & strName
    End If
End Sub

Public Function GetDetailParas(ByVal lngParID As Long, Optional ByRef rsSysTems As ADODB.Recordset, Optional ByRef int部门 As Integer, Optional ByRef int本机 As Integer, Optional ByRef int私有 As Integer, Optional ByRef strOwner As String) As Recordset
'功能：获取参数的详细参数
'参数：lngParID=参数ID
'          rsSystems=系统列表，主要包含字段（系统、所有者）
'返回：参数id, 站点, 部门id,部门, 部门简码,用户名, 人员id,人员,人员简码,机器名,机器名简码,参数值
    Dim rsParInfo  As ADODB.Recordset, rsDetailParas As ADODB.Recordset
    Dim strSQL As String
    Dim lngSys As Long
    Dim StrDefaultSQL As String
    
    On Error GoTo errh
    strSQL = "Select Nvl(系统, 0) 系统, Nvl(私有, 0) 私有, Nvl(本机, 0) 本机, Nvl(部门, 0) 部门 From zlParameters A Where ID =[1]"
    Set rsParInfo = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetDetailParas", lngParID)
    If Not rsParInfo.EOF Then
        int部门 = Val(rsParInfo!部门)
        int本机 = Val(rsParInfo!本机)
        int私有 = Val(rsParInfo!私有)
        lngSys = Val(rsParInfo!系统)
    End If
    If Not (int部门 = 0 And int本机 = 0 And int私有 = 0) And rsSysTems Is Nothing Then
    '公共模块和公共全局,不存在详细参数数据，因此不用获取系统
        Set rsSysTems = New ADODB.Recordset
        Set rsSysTems = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    End If
    If lngSys <> 0 Then '非管理工具参数
        rsSysTems.Filter = "编号=" & lngSys
        If Not rsSysTems.EOF Then strOwner = rsSysTems!所有者
    Else '管理工具参数为私有全局或公共全局
        rsSysTems.Filter = "编号=100"
        If Not rsSysTems.EOF Then
            strOwner = rsSysTems!所有者
        Else
             rsSysTems.Filter = ""
             rsSysTems.Sort = "编号"
             If Not rsSysTems.EOF Then strOwner = rsSysTems!所有者
        End If
    End If
    If int部门 = 1 Then '部门参数
        strSQL = "Select a.参数id, b.站点, a.部门id, b.名称 部门, zlSpellCode(b.名称) 部门简码, Null 用户名, Null 人员id, Null 人员, Null 人员简码, Null 机器名," & vbNewLine & _
                    "       Null 机器名简码, a.参数值" & vbNewLine & _
                    "From Zldeptparas A, " & strOwner & ".部门表 B" & vbNewLine & _
                    "Where a.参数id =[1] And a.部门id = b.Id"
        StrDefaultSQL = "Select a.参数id, Null 站点, a.部门id, Null 部门, Null 部门简码, Null 用户名, Null 人员id, Null 人员, Null 人员简码, Null 机器名, Null 机器名简码, a.参数值" & vbNewLine & _
                                    "From Zldeptparas A" & vbNewLine & _
                                    "Where a.参数id = [1]"
    Else
        If int本机 = 1 Then '本机类型参数
            StrDefaultSQL = "Select a.参数id, Null 站点, Null 部门id, Null 部门, Null 部门简码, a.用户名, Null 人员id, Null 人员, Null 人员简码, a.机器名, Null 机器名简码, a.参数值" & vbNewLine & _
                                    "From zlUserParas A" & vbNewLine & _
                                    "Where a.参数id = [1]"
            If int私有 = 1 Then '本机私有模块
                strSQL = "Select a.参数id, d.站点, e.Id 部门id, e.名称 部门, zlSpellCode(e.名称) 部门简码, a.用户名, b.人员id, c.姓名 人员, zlSpellCode(c.姓名) 人员简码, a.机器名," & vbNewLine & _
                                "       zlSpellCode(a.机器名) 机器名简码, a.参数值" & vbNewLine & _
                                "From zlUserParas A, " & strOwner & ".上机人员表 B, " & strOwner & ".人员表 C, zlClients D, " & strOwner & ".部门表 E" & vbNewLine & _
                                "Where a.参数id =[1] And a.用户名 = b.用户名(+) And a.机器名 = d.工作站(+) And b.人员id = c.Id(+) And d.部门 = e.名称(+) And" & vbNewLine & _
                                "      a.用户名 Is Not Null And a.机器名 Is Not Null"
            Else '本机公共模块
                strSQL = "Select a.参数id, d.站点, e.Id 部门id, e.名称 部门, zlSpellCode(e.名称) 部门简码, a.用户名, Null 人员id, Null 人员, Null 人员简码, a.机器名," & vbNewLine & _
                            "       zlSpellCode(a.机器名) 机器名简码, a.参数值" & vbNewLine & _
                            "From zlUserParas A, zlClients D, " & strOwner & ".部门表 E" & vbNewLine & _
                            "Where a.参数id =[1] And a.机器名 = d.工作站(+) And d.部门 = e.名称(+) And a.用户名 Is Null And a.机器名 Is Not Null"
            End If
        Else
            If int私有 = 1 Then '私有模块或私有全局
                strSQL = "Select a.参数id, e.站点, e.Id 部门id, e.名称 部门, zlSpellCode(e.名称) 部门简码, a.用户名, b.人员id, c.姓名 人员, zlSpellCode(c.姓名) 人员简码, a.机器名," & vbNewLine & _
                            "       zlSpellCode(a.机器名) 机器名简码, a.参数值" & vbNewLine & _
                            "From zlUserParas A, " & strOwner & ".上机人员表 B, " & strOwner & ".人员表 C, " & strOwner & ".部门表 E, " & strOwner & ".部门人员 F" & vbNewLine & _
                            "Where a.参数id =[1] And a.用户名 = b.用户名(+) And b.人员id = c.Id(+) And c.Id = f.人员id(+) And f.缺省 = 1 And f.部门id = e.Id(+) And" & vbNewLine & _
                            "      a.用户名 Is Not Null And a.机器名 Is Null"
            Else '公共模块和公共全局,不存在相关的数据
                strSQL = "Select a.参数id, Null 站点, Null 部门id, Null 部门, Null 部门简码, a.用户名, Null 人员id, Null 人员, Null 人员简码, a.机器名," & vbNewLine & _
                            "       zlSpellCode(a.机器名) As 机器名简码, a.参数值" & vbNewLine & _
                            "From zlUserParas A" & vbNewLine & _
                            "Where 参数id = [1] And 1 = 2"
            End If
        End If
    End If
    If strOwner = "" Then '没有安装任何应用系统
        Set rsDetailParas = gclsBase.OpenSQLRecord(gcnOracle, StrDefaultSQL, "GetDetailParas", lngParID)
    Else
        On Error Resume Next '可能该系统没有部门人员等基础表
        Set rsDetailParas = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetDetailParas", lngParID)
        If err.Number <> 0 Then
            err.Clear
            On Error GoTo errh
            strOwner = ""
            Set rsDetailParas = gclsBase.OpenSQLRecord(gcnOracle, StrDefaultSQL, "GetDetailParas", lngParID)
        Else
            On Error GoTo errh
        End If
    End If
    Set GetDetailParas = rsDetailParas
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function


Public Function KillSessions(Optional ByVal strSesionInfo As String, Optional ByRef cllRacConn As Collection) As Boolean
'功能：获取杀掉会话的SQL
'blnKill=是否直接执行
'strSesionInfo=杀掉制定回话ID,一般为"Sid,Serial"
'cllRacConn=返回RAC环境下所需要的连接，为空，则不需要进行Rac特殊处理
    Dim rsTmp As ADODB.Recordset, strSQL As String, strAdjustSQL As String, strKillProcess As String
    Dim strTmp As String, bln10g As Boolean, strPre As String
    Dim rsIns As ADODB.Recordset, cnnTmp As ADODB.Connection
    Dim strUser As String, rsSysTems As ADODB.Recordset
    
    On Error GoTo errh
    '可能此时gstrUserName尚未获取
    If gstrUserName <> "" Then
        strUser = gstrUserName
    Else
        Set rsSysTems = New ADODB.Recordset
        Set rsSysTems = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
        If Not rsSysTems.EOF Then
            rsSysTems.Sort = "编号"
            strUser = rsSysTems!所有者
        End If
    End If
    
    If strSesionInfo <> "" Then
        gcnOracle.Execute "alter system kill session '" & strSesionInfo & "' immediate"
        KillSessions = True
        Exit Function
    End If
    '获取数据库版本
    bln10g = GetOracleVersion(True, True) < 11
    '如果直接杀掉之前，需要禁用客户端
    strSQL = "Select 项目, 内容 From Zlupgradeconfig Where 项目 =[1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, "客户端状态")
    If rsTmp.EOF Then
        strAdjustSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('客户端状态',1)"
    ElseIf Val(rsTmp!内容 & "") = 0 Then
        strAdjustSQL = "Update ZLTOOLS.zlUpgradeConfig Set  内容='1'  Where 项目 ='客户端状态'"
    End If
    If strAdjustSQL <> "" Then '标记已经使用禁用客户端并杀掉会话功能，可以在升迁管理中点击按钮启用
        gcnOracle.Execute strAdjustSQL
    End If
    On Error Resume Next
    '该语句跟踪出来在PL/SQL中执行得不到想要的数据
    strSQL = "Select Distinct A.Username, A.Program, A.Audsid, B.Ip, B.工作站" & vbNewLine & _
                    "From v$session a, Zlclients b" & vbNewLine & _
                    "Where A.Terminal = B.工作站 And Upper(A.Program) =Upper( [1] ) And A.Audsid = Userenv('SessionID') And" & vbNewLine & _
                    "      B.Ip = Sys_Context('USERENV', 'IP_ADDRESS')"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, IIf(gblnInIDE, "vb6.exe", App.EXEName & ".exe"))
    If Not rsTmp.EOF Then strTmp = rsTmp!工作站 & ""
    If err.Number <> 0 Then err.Clear
    On Error GoTo errh
    '禁用客户端
    strAdjustSQL = "Update Zlclients Set 禁止使用 = 1, 系统升级禁用 = 1 Where Nvl(禁止使用, 0) = 0 " & IIf(strTmp <> "", "  And 工作站 <> '" & strTmp & "'", "")
    gcnOracle.Execute strAdjustSQL
    '判断是否存在ZLkillProcess表
    If CheckAndAdjustMustTable("zlkillprocess", , False) Then
        strKillProcess = "zlkillprocess"
        On Error Resume Next
        If err.Number <> 0 Then err.Clear
        strSQL = "Select Count(1) 计数 From Zltools.Zlkillprocess Where Rownum < 2"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "zlkillprocess数据判断")
        If rsTmp!计数 = 0 Then
            strKillProcess = ""
        End If
        '可能没有查询权限
        If err.Number <> 0 Then
            err.Clear
            strKillProcess = ""
        End If
        On Error GoTo errh
    End If
    If strKillProcess <> "" Then
        strKillProcess = "Select Upper(名称) From Zltools.Zlkillprocess Union All" & vbNewLine & _
                        "Select 'VB6.EXE' From Zltools.Zlkillprocess"
    Else
        strKillProcess = "'ZL9LABPRINTSVR.EXE','ZL9LABRECEIV.EXE','ZL9LABTCPSVR.EXE','ZL9LISCOMM.EXE'," & _
                        "'ZL9WIZARDMAIN.EXE','ZLACTMAIN.EXE','ZLHIS+.EXE','ZLHISCRUST.EXE','ZLLISRECEIVESEND.EXE'," & _
                        "'ZLNEWQUERY.EXE','ZLORCLCONFIG.EXE','ZLPACSBROWSERSTATION.EXE','ZLPACSSRV.EXE'," & _
                        "'ZLPEISAUTOANALYSE.EXE','ZLRPTSQLADJUST.EXE','ZLRUNAS.EXE','ZLSVRNOTICE.EXE'," & _
                        "'ZLSVRSTUDIO.EXE','ZLWIZARDSTART.EXE','VB6.EXE'"
    End If
    strTmp = ""
    
    '11gR2可以直接杀会话ALTER system KILL SESSION '73,15625,@1'
    '10g需要登录到对应的Rac实例
    strSQL = "Select 'alter system kill session ' || Chr(39) || a.Sid || ',' || a.Serial# || " & IIf(bln10g, "", "',@' || a.INST_ID || ") & " Chr(39) || ' immediate' SQL," & vbNewLine & _
            "       a.Program, b.Ip," & IIf(bln10g, " a.INST_ID,  Decode(INST_ID, userenv('instance'), 1, 0) 当前标志", "userenv('instance') INST_ID,1 当前标志") & vbNewLine & _
            "  From Gv$session a, Zlclients b" & vbNewLine & _
                "Where a.Terminal = b.工作站 And Upper(a.Program) In" & vbNewLine & _
                "(" & strKillProcess & ") And" & vbNewLine & _
                "      (a.Terminal <> userenv('terminal') Or" & vbNewLine & _
                "      a.Terminal= userenv('terminal')  And Upper(a.Program) Not In ('VB6.EXE', 'ZLSVRSTUDIO.EXE'))" & vbNewLine & _
                "Order By a.Terminal,a.Program"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取会话")
    
    Set cllRacConn = New Collection
    If bln10g Then
        rsTmp.Filter = "当前标志=0"
        If Not rsTmp.EOF Then
            strSQL = "select a.inst_ID, a.Instance_Name, a.Host_name, b.NAME, b.DBID" & vbNewLine & _
                    "  from gv$instance a, gv$database b" & vbNewLine & _
                    " where a.INST_ID = b.INST_ID" & vbNewLine & _
                    "   and a.INST_ID <> userenv('instance')" & vbNewLine & _
                    "   and a.STATUS = 'OPEN'"
            Set rsIns = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取实例信息")
            '不遍历rsTmp，主要是为了解约时间
            Do While Not rsIns.EOF
                strTmp = rsIns!INST_ID & "," & rsIns!DBID & "," & rsIns!Instance_Name & "(" & rsIns!name & ")"
                If frmUserCheckLogin.ShowLogin(UCT_RACInsUser, cnnTmp, strUser, "", "", strTmp) Then
                    cllRacConn.Add cnnTmp, "K_" & rsIns!INST_ID
                End If
                rsIns.MoveNext
            Loop
        End If
        rsTmp.Filter = ""
    End If
    rsTmp.Sort = "当前标志 desc,INST_ID"
    On Error Resume Next
    strTmp = "": strPre = ""
    Do While Not rsTmp.EOF
        If rsTmp!当前标志 = 0 Then
            If strPre <> rsTmp!INST_ID Then
                strPre = rsTmp!INST_ID & ""
                Set cnnTmp = cllRacConn("K_" & rsIns!INST_ID)
            End If
            cnnTmp.Execute rsTmp!SQL
        Else
            gcnOracle.Execute rsTmp!SQL
        End If
        If err.Number <> 0 Then
            strTmp = strTmp & vbNewLine & rsTmp!Program & "(" & rsTmp!IP & ")"
            err.Clear
        End If
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then
        MsgBox "如下程序的会话无法中止：" & strTmp & "。请手工处理！", vbInformation, gstrSysName
    End If
    KillSessions = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function LockAppUser() As Boolean
'功能：禁用应用系统用户
'blnKill=是否直接执行
'strSesionInfo=杀掉制定回话ID,一般为"Sid,Serial"
'cllRacConn=返回RAC环境下所需要的连接，为空，则不需要进行Rac特殊处理
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strUser As String, rsSysTems As ADODB.Recordset
    
    On Error Resume Next
    If Not CheckAndAdjustMustTable("Zlclients", "系统升级禁用", True) Then
        MsgBox "没有必要的结构支持，无法锁定用户，请手工处理！", vbInformation, gstrSysName
        Exit Function
    End If
    '可能此时gstrUserName尚未获取
    If gstrUserName <> "" Then
        strUser = gstrUserName
    Else
        Set rsSysTems = New ADODB.Recordset
        Set rsSysTems = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
        If Not rsSysTems.EOF Then
            rsSysTems.Sort = "编号"
            strUser = rsSysTems!所有者
        End If
    End If
    '标记锁定的用户
    strSQL = "Update " & strUser & ".上机人员表 b" & vbNewLine & _
                        "Set 系统升级锁定 = 1" & vbNewLine & _
                        "Where Exists (Select 1 From Dba_Users a Where Account_Status = 'OPEN' And A.Username = Upper(B.用户名)) And Upper(B.用户名)<>'" & UCase(strUser) & "'"
    gcnOracle.Execute strSQL
    '标记已经禁用功能，可以在升迁管理中点击按钮启用
    strSQL = "Select 项目, 内容 From Zlupgradeconfig Where 项目 =[1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, "用户状态")
    strSQL = ""
    If rsTmp.EOF Then
        strSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('用户状态',1)"
    ElseIf Val(rsTmp!内容 & "") = 0 Then
        strSQL = "Update ZLTOOLS.zlUpgradeConfig Set  内容='1'  Where 项目 ='用户状态'"
    End If
    If strSQL <> "" Then
        gcnOracle.Execute strSQL
    End If
    '锁定用户账户
    strSQL = "Select 'alter user ' || 用户名 || ' account lock '  SQL" & vbNewLine & _
                "From " & strUser & ".上机人员表 b" & vbNewLine & _
                "Where Exists (Select 1 From Dba_Users a Where Account_Status = 'OPEN' And A.Username = Upper(B.用户名)) And Upper(B.用户名)<>[1] " & vbNewLine & _
                "Order By 用户名"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取应用系统用户", UCase(strUser))
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            gcnOracle.Execute rsTmp!SQL
            If err.Number <> 0 Then err.Clear
            rsTmp.MoveNext
        Loop
    End If
    LockAppUser = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function CreateTable(ByVal cnCurrentSelect As ADODB.Connection, ByVal strOwner As String, ByVal strTableSpaces As String, ByVal strBakName As String, ByVal strTable As String, ByVal strTbsNameLob As String, Optional ByVal cnDBACreate As ADODB.Connection) As String
    '--------------------------------------------------------------------------------------------------------------------
    '功能:构建table.sql
    '参数:strTableSpaces-表空间,strTbsNameLob-大对象表空间
    '     strBakName-空间所有者
    '     strOwner-表的所有者
    '     strTableName-表名
    '     cnDBACreate=进行表结构创建的连接
    '     cnCurrentSelect=进行表结构查询的连接
    '返回：成功返回true,否则返回False
    '--------------------------------------------------------------------------------------------------------------------
    Dim rsTable As New ADODB.Recordset
    Dim rsColumn As New ADODB.Recordset
    Dim strTemp As String, strLobs As String
    Dim strSQL As String, blnHaveLob As Boolean
    
    On Error GoTo ErrHand
    strSQL = "Select a.Table_Name,a.Column_Name, " & _
             "               a.Data_Type, " & _
             "               a.Data_Length, " & _
             "               a.Data_Precision, " & _
             "               a.Data_Scale, " & _
             "               a.Nullable, " & _
             "               a.Data_Default " & _
             "   From Sys.all_Tab_Columns a  " & _
             "   Where a.Owner = [1] and table_Name=[2]" & _
             "   Order By a.Table_Name,a.Column_Id"
    Set rsColumn = gclsBase.OpenSQLRecord(cnCurrentSelect, strSQL, "CreateTable", strOwner, strTable)
    
    strSQL = "select table_name, tablespace_name, pct_free, TEMPORARY,DURATION 周期 " & _
             " from sys.all_tables where owner = [1] and table_name=[2]"
    Set rsTable = gclsBase.OpenSQLRecord(cnCurrentSelect, strSQL, "CreateTable", strOwner, strTable)
    
    strSQL = ""
    With rsTable
        Do While Not .EOF
            rsColumn.Filter = "Table_Name='" & !Table_Name & "'"
            If rsColumn.EOF Then
                If Not cnDBACreate Is Nothing Then
                    MsgBox "表:" & !Table_Name & "的列名不存在!", vbInformation + vbDefaultButton1
                End If
                Exit Function
            End If
            If Nvl(!Temporary) = "Y" Then
                    strSQL = "CREATE GLOBAL TEMPORARY TABLE " & strBakName & "." & !Table_Name & "("
            Else
                   strSQL = "CREATE TABLE " & strBakName & "." & !Table_Name & "("
            End If
            
            strLobs = ""
            Do While Not rsColumn.EOF
                Select Case rsColumn!DATA_TYPE
                Case "NUMBER"
                    strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE & _
                            IIf(Nvl(rsColumn!Data_Precision) = "", "", "(" & Nvl(rsColumn!Data_Precision) & IIf(Val(Nvl(rsColumn!Data_Scale)) = 0, "", "," & Val(Nvl(rsColumn!Data_Scale))) & ")")
                Case "DATE", "FLOAT", "TIMESTAMP(6)" 'TIMESTAMP(6)：精确定小数秒
                    strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE
                    
                Case "BLOB", "CLOB", "BFILE", "XMLTYPE"
                    strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE
                    blnHaveLob = True
                    
                    If rsColumn!DATA_TYPE = "BLOB" Or rsColumn!DATA_TYPE = "CLOB" Then
                        strLobs = IIf(strLobs = "", "", strLobs & ",") & rsColumn!Column_Name
                    End If
                Case Else
                    If Val(Nvl(rsColumn!Data_Length)) = 0 Then
                        strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE
                    Else
                        strTemp = RPAD(rsColumn!Column_Name, 15, " ") & " " & rsColumn!DATA_TYPE & "(" & Nvl(rsColumn!Data_Length) & ")"
                    End If
                End Select
                If rsColumn.AbsolutePosition = rsColumn.RecordCount Then
                    strSQL = strSQL & " " & strTemp & IIf(Nvl(rsColumn!DATA_DEFAULT) = "", "", " DEFAULT " & Replace(Nvl(rsColumn!DATA_DEFAULT), Chr(10), "")) & ")"
                Else
                    strSQL = strSQL & " " & strTemp & IIf(Nvl(rsColumn!DATA_DEFAULT) = "", "", " DEFAULT " & Replace(Nvl(rsColumn!DATA_DEFAULT), Chr(10), "")) & ","
                End If
                
                rsColumn.MoveNext
            Loop
            
            If Nvl(!Temporary) = "Y" Then
                If InStr(1, Nvl(!周期), "TRANSACTION") > 0 Then
                    strSQL = strSQL & " ON COMMIT DELETE ROWS;"
                Else
                    strSQL = strSQL & " ON COMMIT PRESERVE ROWS;"
                End If
            Else
                If strLobs <> "" Then
                    If GetOracleVersion(True, True) > 10 Then
                        strLobs = " Lob(" & strLobs & ") Store as Securefile(NOCache LOGGING)"
                    Else
                        strLobs = " Lob(" & strLobs & ") Store as (Cache)"  '测试表明Basefile方式下Cache LOGGING的写入最快,10G不支持Cache LOGGING关键字，缺省是LOGGING
                    End If
                End If
                strSQL = strSQL & strLobs & " TABLESPACE " & IIf(blnHaveLob, strTbsNameLob, strTableSpaces)
                If Nvl(!pct_free) <> "" Then
                    '由于是只读数据，为提高存储效率，固定pctfree为5
                    strSQL = strSQL & " PCTFREE 5"
                End If
            End If
            .MoveNext
        Loop
    End With
    If Not cnDBACreate Is Nothing Then
        cnDBACreate.Execute strSQL
    End If
    CreateTable = strSQL
    Exit Function
ErrHand:
    If Not cnDBACreate Is Nothing Then
        If MsgBox("在获取相关的表结构时发生错误，详细错误如下:" & vbCrLf & err.Description & _
            vbCrLf & strSQL & vbCrLf & "是否跳过此表的创建?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            CreateTable = IIf(strSQL = "", "Skip", strSQL)
        End If
    End If
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub CreateTempTabForBakTable(ByVal strTable As String)
'功能：为历史转出表创建临时表及同义词
'参数：
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTableH As String
    
    On Error GoTo ErrHand
    strTableH = "H" & strTable
    
    '如果已存在临时表，则不必重建
    strSQL = "Select 1 From User_Tables Where Table_Name = [1] And Temporary = 'Y'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "检查历史表视图", strTableH)
    If rsTmp.RecordCount = 0 Then

        '以前创建的公共同义词虽然指向的是H视图，现在改为H表后，不影响使用，它只是通过名称关联的，没有管对象类型，所以不用删除后重建
        '1.删除以前指向历史表空间的视图
        '如果是10.35.70之前的版本，才有视图
        strSQL = "Select 1 From User_Views Where View_Name = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "检查历史表视图", strTableH)
        If rsTmp.RecordCount > 0 Then
            strSQL = "Drop View " & strTableH
            gcnOracle.Execute strSQL
        End If
        
        '2.创建临时表
        strSQL = "Create Global Temporary Table " & strTableH & " On Commit Delete Rows as select * from " & strTable & " where 1=0"
        gcnOracle.Execute strSQL
    End If
    
    Exit Sub
ErrHand:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub DropTempTabForBakTable(ByVal strTable As String)
'功能：检查并删除临时历史表
'参数：
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTableH As String
    
    On Error GoTo ErrHand
    strTableH = "H" & strTable
    
    strSQL = "Select 1 From User_Tables Where Table_Name = [1] And Temporary = 'Y'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "检查历史表视图", strTableH)
    If rsTmp.RecordCount > 0 Then
        '由Dblink历史库切换为本地历史库时，需删除同名的临时表后才能创建视图
        strSQL = "Drop Table " & strTableH
        gcnOracle.Execute strSQL
    End If
    
    Exit Sub
ErrHand:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Public Function CreateAppView(ByVal strOwner As String, ByVal strBakOwner As String, _
    ByVal lng系统 As Long, ByVal strDbLink As String, _
    Optional ByRef pgbState As ProgressBar, Optional ByRef clsScript As clsRunScript) As Boolean
    '----------------------------------------------------------------------------------------------------------------------
    '功能:创建H表的视图
    '参数:gcnOracle-当前应用系统的数据连接
    '     lng系统-系统编号
    '     strOwner-应用系统的所有者
    '     strBakOwner-历史数据空间的所有者
    '     strDbLink-数据库链接名称，第一个字符是@
    '     pgbState-进度条控件
    '     clsScript-升级日志输出类
    '返回:创建成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------------------
    Dim rsTables As ADODB.Recordset
    Dim strBakTableName As String
    Dim i As Long
    
    On Error GoTo ErrHand
    
    gstrSQL = "Select 表名 from zlBakTables where 系统=[1]"
    Set rsTables = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "读取转出表", lng系统)
    
    If Not pgbState Is Nothing Then
        pgbState.Max = 100
        pgbState.value = 0
        DoEvents
    End If
    
    On Error Resume Next
    With rsTables
        Do While Not .EOF
            If strDbLink = "" Then '检查之前是否存在因远程历史库而创建的临时H表
                Call DropTempTabForBakTable(!表名)
            End If
            
            strBakTableName = strBakOwner & "." & !表名 & strDbLink
            
            gstrSQL = "Create or replace view  " & strOwner & ".H" & !表名 & " as Select * From " & strBakTableName
            gcnOracle.Execute gstrSQL
            
            '含有LOB字段的表无法创建指向远程服务器的视图，报错：ORA-22992: 无法使用从远程表选择的 LOB 定位器
            '改为创建临时表，读取时向临时表插入数据,因为通过dblink访问lob，支持insert into ...select 方式
            If strDbLink <> "" Then
                If InStr(err.Description, "ORA-22992") > 0 Then
                    err.Clear
                    Call CreateTempTabForBakTable(!表名)
                End If
            End If
            
            If err.Number <> 0 Then
                If Not clsScript Is Nothing Then
                    clsScript.ErrCount = clsScript.ErrCount + 1
                    clsScript.WriteLog Format(Now, "HH:mm:ss") & "，" & RPAD(strOwner & "." & "H" & !表名 & "创建失败", 30) & "，错误：" & err.Description
                Else
                    If MsgBox("创建视图""H" & !表名 & """出错," & vbCrLf & " 错误信息:(" & err.Number & ") " & err.Description & vbCrLf & "是否忽略该错，继续执行？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
                err.Clear
            End If
            If Not pgbState Is Nothing Then
                i = i + 1
                pgbState.value = i / rsTables.RecordCount * 100
                DoEvents
            End If
            .MoveNext
        Loop
    End With
    
    '单独创建zlbakInfo视图
    If rsTables.RecordCount <> 0 Then
        strBakTableName = strBakOwner & ".ZLBAKINFO" & strDbLink
        gstrSQL = "Create or replace view  " & strOwner & "." & "ZLBAKINFO as Select * From  " & strBakTableName
        gcnOracle.Execute gstrSQL
        
        If err.Number <> 0 Then
            If Not clsScript Is Nothing Then
                clsScript.ErrCount = clsScript.ErrCount + 1
                clsScript.WriteLog Format(Now, "HH:mm:ss") & "，" & RPAD(strOwner & "." & "ZLBAKINFO" & "创建失败", 30) & "，错误：" & err.Description & _
                        IIf(strDbLink = "", "", vbCrLf & "远程连接可能不可用(" & strDbLink & ")")
            Else
                MsgBox strOwner & "." & "ZLBAKINFO" & "创建失败." & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
    End If
    
    CreateAppView = True
    Exit Function
ErrHand:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function IsCanInstallPLJson(ByVal strToolsFloder As String, Optional ByRef blnInstallRemain As Boolean) As Boolean
'功能：判定是否可以安装PLJSON。1、已经安装，则不用安装。2、没有安装，检查PLJSON脚本是否存在，存在则可以安装
'参数：strToolsFolder=APPSOFT\TOOLS\目录位置
'
'返回：True-可以安装PLJSON，False-不能安装PLJSON
'      blnInstallRemain=是否存在安装残留
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errh
    'TYPE BODY,TYPE,SYNONYM
    strSQL = "Select Count(1) 数量" & vbNewLine & _
            "From All_Objects a" & vbNewLine & _
            "Where a.Object_Name In ('JSON', 'JSON_VALUE','JSON_VALUE_ARRAY', 'JSON_LIST','JSON_HELPER','JSON_PARSER','JSON_EXT','JSON_AC')"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "是否安装PLJSON")
    blnInstallRemain = False
    '可能存在私有同义词导致数量大于等于23，JSON_VALUE_ARRAY没有TYPE BODY单独定义
    If rsTmp!数量 < 23 Then
        If gobjFile.FileExists(strToolsFloder & "\PLJSON1.0.6install.SQL") And gobjFile.FileExists(strToolsFloder & "\PLJSON1.0.6uninstall.SQL") Then
            IsCanInstallPLJson = True
        End If
        blnInstallRemain = rsTmp!数量 <> 0
    End If
    Exit Function
errh:
    Call WriteTraceLog("IsConInstallPLJson:" & err.Description)
    err.Clear
End Function

Public Function InstallPLJSON(ByVal cnDBA As ADODB.Connection, ByVal strToolsFloder As String, ByRef objRunScript As clsRunScript, Optional ByVal blnUninstallFirst As Boolean) As Boolean
'功能：安装PLJSON,安装失败，则自动反安装
'参数：strToolsFolder=APPSOFT\TOOLS\目录位置
'      cnDBA=安装所需的DBA连接
'      objRunScript=脚本文件解析对象
'      blnUninstallFirst=是否先进行反安装，可能安装中断，导致部分对象存在
'返回：True-安装成功，False-安装失败
    Dim blnInstallErr As Boolean

    On Error GoTo errh
    '若存在安装残留，则先进行反安装
    If blnUninstallFirst Then
        Call UninstallPLJSON(cnDBA, strToolsFloder, objRunScript)
    End If
    If Not objRunScript.OpenFile(strToolsFloder & "\PLJSON1.0.6install.SQL") Then
        objRunScript.WriteLog String(9, " ") & "结果：PLJSON安装失败"
        Exit Function
    End If
    blnInstallErr = True
    Do While Not objRunScript.EOF
        cnDBA.Execute objRunScript.SQLInfo.SQL
        objRunScript.ReadNextSQL
    Loop
    Exit Function
errh:
    If blnInstallErr Then objRunScript.WriteLog Format(objRunScript.SQLInfo.FileLine, "0000000") & ":" & GetLogSQL(objRunScript.SQLInfo), 2
    objRunScript.WriteLog String(17, " ") & "错误：" & err.Description
    err.Clear
    If blnInstallErr Then
        objRunScript.WriteLog String(17, " ") & "处理：自动终止安装并进行反安装"
        objRunScript.WriteLog String(9, " ") & "结果：PLJSON安装失败"
        Call UninstallPLJSON(cnDBA, strToolsFloder, objRunScript)
    Else
        objRunScript.WriteLog String(17, " ") & "处理：自动终止安装"
        objRunScript.WriteLog String(9, " ") & "结果：PLJSON安装失败"
    End If
End Function

Public Function UninstallPLJSON(ByVal cnDBA As ADODB.Connection, ByVal strToolsFloder As String, ByRef objRunScript As clsRunScript) As Boolean
'功能：反安装PLJSON
'参数：strToolsFolder=APPSOFT\TOOLS\目录位置
'      cnDBA=反安装所需的DBA连接
'      objRunScript=脚本文件解析对象
'返回：True-反安装成功，False-反安装失败
    Dim blnUninstallErr As Boolean

    On Error GoTo errh
    If Not objRunScript.OpenFile(strToolsFloder & "\PLJSON1.0.6uninstall.SQL") Then
        objRunScript.WriteLog String(9, " ") & "结果：PLJSON反安装失败"
        Exit Function
    End If
    blnUninstallErr = True
    Do While Not objRunScript.EOF
        cnDBA.Execute objRunScript.SQLInfo.SQL
        objRunScript.ReadNextSQL
    Loop
    Exit Function
errh:
    If blnUninstallErr Then objRunScript.WriteLog Format(objRunScript.SQLInfo.FileLine, "0000000") & ":" & GetLogSQL(objRunScript.SQLInfo), 2
    objRunScript.WriteLog String(17, " ") & "错误：" & err.Description
    objRunScript.WriteLog String(17, " ") & "处理：自动终止反安装"
    err.Clear
End Function

