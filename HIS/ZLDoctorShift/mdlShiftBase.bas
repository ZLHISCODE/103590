Attribute VB_Name = "mdlShiftBase"
Option Explicit

Public gobjPublicAdvice As Object           '临床公共部件
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function GetShiftType(ByVal bytType As Byte, ByVal strDeptID As String) As ADODB.Recordset
'不同科室可能有相同名称的值班班次
    
    On Error GoTo errH
    Select Case bytType
        Case 1 '获取所有的值班班次信息
            gstrSQL = "Select b.名称 科室, a.值班班次 班次名称, To_Char(a.开始时间, 'hh24:mi') 开始时间, To_Char(a.结束时间, 'hh24:mi') 结束时间" & vbNewLine & _
                    "From 医生值班班次 a, 部门表 b Where a.科室id = b.Id And b.id In(Select * From Table(f_Str2list([1]))) Order By a.开始时间"
        Case 2 '只获取所有的值班班次名称
            gstrSQL = "Select Distinct a.值班班次 班次名称 From 医生值班班次 a, 部门表 b Where a.科室id = b.Id " & vbNewLine & _
                "And b.id In(Select * From Table(f_Str2list([1])))"
    End Select
    Set GetShiftType = zlDatabase.OpenSQLRecord(gstrSQL, "获取班次信息", strDeptID)
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "获取班次信息"
End Function

Public Function GetDeptName(ByVal strDeptID As String) As ADODB.Recordset
    
    On Error GoTo errH
    gstrSQL = "Select 编码 ||'-' || 名称 as 名称, Id,编码 From 部门表 Where Id In (Select * From Table(f_Str2list([1]))) Order by 编码"
    Set GetDeptName = zlDatabase.OpenSQLRecord(gstrSQL, "获取班次信息", strDeptID)
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "获取部门名称"
End Function

Public Function GetPatientType() As ADODB.Recordset

    On Error GoTo errH
    gstrSQL = "Select 简称, 名称, 顺序,提取SQL From 医生交接班病人类型 Where 是否停用 = 0 And 提取sql Is Not Null Order By 顺序"
    Set GetPatientType = zlDatabase.OpenSQLRecord(gstrSQL, "获取班次信息")
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "获取病人类型信息"
End Function

Public Function GetTimeRangePati(ByVal dtBegin As Date, ByVal dtEnd As Date, ByVal strDeptID As String) As ADODB.Recordset
'获取一定时间范围内的患者信息
    Dim rsTemp As ADODB.Recordset
        
    Set rsTemp = GetPatiType
    gstrSQL = ""
    Do While Not rsTemp.EOF
        gstrSQL = IIf(gstrSQL = "", "", gstrSQL & vbNewLine & "Union All ") & rsTemp!提取SQL
        rsTemp.MoveNext
    Loop
    gstrSQL = UCase(gstrSQL)
    gstrSQL = Replace(gstrSQL, "[开始时间]", zlStr.To_Date(dtBegin))
    gstrSQL = Replace(gstrSQL, "[结束时间]", zlStr.To_Date(dtEnd))
    gstrSQL = Replace(gstrSQL, "[科室ID]", "[1]")
    Set GetTimeRangePati = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息", strDeptID)
End Function

Public Function GetPatiType() As ADODB.Recordset
'获取病人类型信息

    On Error GoTo errH
    gstrSQL = "Select 简称,名称,顺序, 提取sql From 医生交接班病人类型 Where 提取sql Is Not Null"
    gstrSQL = gstrSQL & vbNewLine & "Union All Select '出院','出院患者',98,'Select Distinct ''出院'' 类型, a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.出院病床 床号, a.住院号 标识号, a.入院日期 入院时间, a.入院方式, a.出院科室id" & vbNewLine & _
        "From 病案主页 a" & vbNewLine & _
        "Where a.出院日期 >  [开始时间] And" & vbNewLine & _
        "      a.出院日期 <=  [结束时间] And a.病人性质 In(0,2) And a.出院科室id In (Select /*+cardinality(a,10)*/* From Table(f_Str2list([科室ID] )) a) ' From Dual"
    gstrSQL = gstrSQL & " Order By 顺序"
    Set GetPatiType = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息SQL")
    Exit Function
errH:
    MsgBox Err.Description, vbInformation, "获取病人类型信息"
End Function

Public Sub ReadSignSource(ByVal lng记录ID As Long, strSource As String)
'功能：获取交接班管理用于电子签名/验证的源文内容
'参数：
'lng记录ID 交接班的记录id；dtBegin记录的开始时间
'返回：签名/验证签名的源文生成规则
'      strSource=签名/验证签名的交接班源文
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim arrField As Variant, strField As String
    Dim strLine As String
    
    On Error GoTo errH
    
    gstrSQL = "Select a.记录Id, a.科室id, a.交班医生, a.交班班次, a.交班开始时间, a.交班结束时间, a.接班医生, a.接班班次, a.接班开始时间, a.接班结束时间, a.记录人, b.内容id, b.序号, b.病人类型," & vbNewLine & _
        "       b.病人id, b.主页id, b.姓名, b.性别, b.年龄, b.床号, b.标识号, b.入院时间, b.入院方式, b.交班描述" & vbNewLine & _
        "From 医生交接班记录 a, 医生交接班内容 b" & vbNewLine & _
        "Where a.记录Id = b.记录id And a.记录id = [1]" & vbNewLine & _
        "Order By 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlShiftBase", lng记录ID)

    strField = "记录ID,科室ID,交班医生,交班班次,交班开始时间,交班结束时间,接班医生,接班班次,接班开始时间,接班结束时间,记录人," & _
        "内容ID,序号,病人类型,病人ID,主页ID,姓名,性别,年龄,床号,标识号,入院时间,入院方式,交班描述"
    arrField = Split(strField, ",")
        
    '生成医嘱签名源文
    Do While Not rsTmp.EOF
        strLine = ""
        For i = 0 To UBound(arrField)
            If IsDate(rsTmp.Fields(arrField(i)).Value) Then
                strLine = strLine & vbTab & Format(rsTmp.Fields(arrField(i)).Value, "yyyy-MM-dd HH:mm:ss")
            Else
                strLine = strLine & vbTab & rsTmp.Fields(arrField(i)).Value & ""
            End If
        Next
        strSource = strSource & vbCrLf & Mid(strLine, 2)
        rsTmp.MoveNext
    Loop
    
    strSource = Mid(strSource, 3)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetUserInfo(ByVal strName As String) As ADODB.Recordset
'根据姓名获其用户信息
    
    On Error GoTo errH
    gstrSQL = "Select b.人员id, b.用户名, a.编号, a.姓名 ,c.部门id From 人员表 a, 上机人员表 b, 部门人员 c Where a.Id = b.人员id And a.Id = c.人员id And a.姓名 = [1]"
    
    Set GetUserInfo = zlDatabase.OpenSQLRecord(gstrSQL, "mdlShiftBase", strName)
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "获取用户信息"
End Function

Public Function GetCA(ByVal strName As String) As Boolean
'根据用户名判断是否启用电子签名，通过姓名获取部门ID
    Dim rsTemp As ADODB.Recordset
    Dim lngCA As Long
    
    On Error GoTo errH
    Set rsTemp = GetUserInfo(strName)
    If rsTemp.RecordCount = 0 Then Exit Function
    gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) 电子签名 From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "电子签名启用部门", 1, rsTemp!部门ID)
    If rsTemp.RecordCount > 0 Then
        lngCA = Val(NVL(rsTemp!电子签名, 0))
    Else
        lngCA = 0
    End If
    GetCA = IIf(lngCA = 1, True, False)
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "是否启用电子签名息"
End Function

Public Function Get主诉(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    '提取病人主诉
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim lngTmp As Long
    
    '主诉，读取老版病历
    strSQL = "Select a.id From 电子病历记录 A, 电子病历格式 B" & vbNewLine & _
            "Where a.Id = b.文件id and a.病人id=[1] and a.主页id=[2] And (a.病历名称 like '%入院记录' or a.病历名称 like '%入院病历' or a.病历名称 like '%入出院记录' or a.病历名称 like '%入院死亡记录')" & vbNewLine & _
            "And b.文本内容 Is Not Null order by 完成时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取主诉", lng病人ID, lng主页ID)
    
    Do While Not rsTmp.EOF
        strTmp = Sys.ReadLobV2("电子病历格式", "文本内容", "文件id=[1]", "", Val(rsTmp!id & ""))
        If strTmp <> "" Then
            strTmp = Replace(Replace(strTmp, Chr(10), ""), Chr(13), "")
            lngTmp = InStr(strTmp, "【主诉】")
            If lngTmp > 0 Then
                lngTmp = lngTmp + 4
                strTmp = Mid(strTmp, lngTmp, InStr(lngTmp, strTmp, "【") - lngTmp)
                strTmp = Replace(strTmp, "主  诉：", "")
                strTmp = Replace(strTmp, "主  诉", "")
                If strTmp <> "" Then
                    Exit Do
                End If
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    '老版没取到，取新版病历
    If strTmp = "" Then
        strTmp = GetItemAppendByEmr("病人主诉", lng病人ID, lng主页ID)
    End If
    
    
    '判断主诉不能大于50个字符
    If strTmp <> "" And zlCommFun.ActualLen(strTmp) > 50 Then
        strTmp = Mid(strTmp, 1, 25)
    End If
    
    Get主诉 = strTmp
End Function

Public Function GetItemAppendByEmr(ByVal str中文名 As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：读取指定病人的指定提纲在病历填写的信息，例如：主诉，诊断等。从病历中获取附项值
    Dim strText As String
    Dim intType As Integer
    Dim lng就诊ID As Long
    
    On Error Resume Next
    
    If gobjEmr Is Nothing Then Exit Function
    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing: Exit Function
 
    If Not gobjEmr Is Nothing Then
        strText = gobjEmr.GetOrderInspectInfoEx(2, lng病人ID, lng主页ID, str中文名)
        If Err.Number <> 0 Then
            strText = gobjEmr.GetOrderInspectInfo(lng病人ID, str中文名)
        End If
    End If
    
    Err.Clear
    GetItemAppendByEmr = strText
End Function

Public Function GetAdviceDiag(ByVal lng医嘱ID As Long, Optional ByRef str诊断 As String) As String
'功能：获得医嘱对应的诊断信息
'参数：str诊断=关联诊断的诊断名称字符串
'返回：关联诊断的ID，逗号分隔
    Dim rsTmp As Recordset, strSQL As String
    Dim strReturn As String
    
    strSQL = "Select  A.ID,a.诊断描述" & vbNewLine & _
            "From 病人诊断记录 A, 病人诊断医嘱 B" & vbNewLine & _
            "Where b.诊断id=a.id And  b.医嘱ID=[1]" & vbNewLine & _
            "Order By b.rowID"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取医嘱相关诊断", lng医嘱ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            str诊断 = str诊断 & "," & rsTmp!诊断描述
            strReturn = strReturn & "," & rsTmp!id
            rsTmp.MoveNext
        Loop
        str诊断 = Mid(str诊断, 2)
        strReturn = Mid(strReturn, 2)
    End If
    GetAdviceDiag = strReturn
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetALL诊断(intType As Integer, ByVal lngPatiID As Long, ByVal lngPageID As Long, dtEnd As Date) As String
    '提取病人最新诊断
    ' intType：病人类型  '1,'新入',2,'抢救',3,'一级护理',4,'术后',5,'术前',6,'死亡',7,'输血',8,'危',9,'其他',10,'危/重',11,'特检',12,'留观'
    Dim rsTmp As ADODB.Recordset, rsTmp1 As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim str诊断 As String

    If intType <> 5 Then
        '取主要诊断
        strSQL = "Select a.病人id, a.诊断类型, a.诊断描述" & vbNewLine & _
                "From 病人诊断记录 A" & vbNewLine & _
                "Where a.诊断类型 In (1, 2, 3, 11, 12, 13) And Nvl(a.编码序号, 1) = 1 And a.诊断次序 = 1 And" & vbNewLine & _
                "      a.病人id=[1] and a.主页id=[2]  And a.取消时间 Is Null" & vbNewLine & _
                "Order By a.病人id Asc, a.记录来源 Desc, a.诊断类型 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取诊断", lngPatiID, lngPageID)
        Do While Not rsTmp.EOF
            If InStr(strTmp, rsTmp!诊断描述) = 0 Then
                strTmp = strTmp & "," & rsTmp!诊断描述
            End If
            rsTmp.MoveNext
        Loop
        strTmp = Mid(strTmp, 2)
        
        '取不到则取最新保存的诊断
        If strTmp = "" Then
            strSQL = "Select a.记录来源, a.诊断次序, a.诊断类型, a.疾病id, a.诊断id, a.诊断描述, a.记录日期, a.记录人 From 病人诊断记录 A Where a.病人id = [1] And a.主页id =[2] And Nvl(a.编码序号, 1) = 1  And A.取消时间 is Null And" & vbNewLine & _
                    "a.记录日期 = (Select Max(a.记录日期) From 病人诊断记录 A Where a.病人id = [1] And a.主页id =[2] And A.取消时间 is Null And Nvl(a.编码序号, 1) = 1)"
    
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取诊断", lngPatiID, lngPageID)
            
            Do While Not rsTmp.EOF
                If InStr(strTmp, rsTmp!诊断描述) = 0 Then
                    strTmp = strTmp & "," & rsTmp!诊断描述
                End If
                rsTmp.MoveNext
            Loop
            strTmp = Mid(strTmp, 2)
        End If

        str诊断 = strTmp
    Else
        strSQL = "Select a.Id, a.相关id, a.诊疗项目id, a.收费细目id, a.医嘱内容, a.医嘱期效, a.医嘱状态, a.开始执行时间, a.诊疗类别, b.操作类型,b.名称 as 项目名称, a.校对护士, a.校对时间,a.标本部位 From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
                "Where a.病人id = [1] And a.主页id = [2] And a.诊疗项目id = b.Id(+) And Nvl(a.医嘱状态, 0) Not In (-1,1,2, 4) And a.诊疗类别='F' and a.手术时间>[3] and a.校对时间 is not null" & vbNewLine & _
                "Order By a.Id, a.序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取诊断", lngPatiID, lngPageID, CDate(dtEnd))
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                '从附项中获取术前诊断如果附项中有以附项为准
                strSQL = "select 内容 from 病人医嘱附件 where 医嘱ID=[1] and 项目='申请单诊断'"
                Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, "获取诊断", Val(rsTmp!id & ""))
                If Not rsTmp1.EOF Then
                    str诊断 = rsTmp1!内容 & ""
                Else
                    '读取术前诊断
                    Call GetAdviceDiag(Val(rsTmp!id & ""), strTmp)
                    str诊断 = strTmp
                End If
                
            End If
        End If
    End If
    GetALL诊断 = str诊断
End Function

Public Function GetNextId(strTable As String, Optional strFild As String) As Long
    '------------------------------------------------------------------------------------
    '功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
    '参数：
    '   strTable：表名称;strFild字段名，序列名称不一定是ID，例如记录ID
    '返回：
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strtab As String
    
    '不能用错误错处理,原因是序列失效和没有序列时,应该返回错误,不然返回零,就有问题!
    '31730
    'On Error GoTo errH
    strtab = Trim(strTable)
    If strtab = "门诊费用记录" Or strtab = "住院费用记录" Then strtab = "病人费用记录"
    If strFild <> "" Then
        strSQL = "Select " & strtab & "_" & strFild & ".Nextval From Dual"
    Else
        strSQL = "Select " & strtab & "_ID.Nextval From Dual"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取序列")
    GetNextId = rsTmp.Fields(0).Value
'    Exit Function
'errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function InitObjPublicAdvice() As Boolean
'功能：初始临床公共部件
    If gobjPublicAdvice Is Nothing Then
        On Error Resume Next
        Set gobjPublicAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
        If Not gobjPublicAdvice Is Nothing Then
            Call gobjPublicAdvice.InitCommon(gcnOracle, glngSys, , , , , , gobjEmr)
        End If
        Err.Clear: On Error GoTo 0
    End If
    InitObjPublicAdvice = Not gobjPublicAdvice Is Nothing
End Function


