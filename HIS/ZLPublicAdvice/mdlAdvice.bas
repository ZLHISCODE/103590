Attribute VB_Name = "mdlAdvice"
Option Explicit

Private mobjVBA As Object
Private mobjScript As clsScript
Private mrsDefine As Recordset

Public Enum Enum_Inside_Program
    p住院记帐操作 = 1150
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p临床路径应用 = 1256
    p辅诊记录管理 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    p观片工具管理 = 1289
    p输液配置中心 = 1345
End Enum

Public Function Get诊疗项目记录(ByVal lngID As Long, Optional ByVal strIDs As String) As ADODB.Recordset
'功能：读取指定诊疗项目ID的记录
'参数：
    Dim StrSQL As String
    
    StrSQL = "Select /*+ rule*/ 计算规则,站点,类别,分类ID,ID,编码,名称,标本部位,计算单位,计算方式,执行频率,适用性别,单独应用,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,参考目录ID,人员ID,建档时间,撤档时间,录入限量,试管编码,执行分类,执行标记" & _
            " From 诊疗项目目录 Where ID"
    On Error GoTo errH
    If strIDs <> "" Then
        StrSQL = StrSQL & " IN(Select Column_Value From Table(f_Num2list([1])))"
        Set Get诊疗项目记录 = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", strIDs)
    Else
        StrSQL = StrSQL & " = [1]"
        Set Get诊疗项目记录 = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lngID)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetMaxAdviceNO(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng婴儿 As Long) As Long
'功能：获取当前病人的最大医嘱序号
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    If lng主页ID = 0 Then
        StrSQL = "Select Nvl(Max(序号),1) as 序号 From 病人医嘱记录 Where 病人ID=[1] And 主页ID Is Null"
    Else
        StrSQL = "Select Nvl(Max(序号),1) as 序号 From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] And Nvl(婴儿,0)=[3]"
    End If
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng病人ID, lng主页ID, lng婴儿)
    If Not rsTmp.EOF Then GetMaxAdviceNO = rsTmp!序号

    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InitAdviceDefine() As Recordset
'功能：读取医嘱内容定义记录集
'参数：blnNew-是否创建objVBA和objScript对象
'说明：
    Dim StrSQL As String
    Dim rsDefine As Recordset
    

    On Error GoTo errH
    StrSQL = "Select 诊疗类别,医嘱内容 From 医嘱内容定义 Order by 诊疗类别"
    Set rsDefine = New ADODB.Recordset
    Call gobjComlib.zlDatabase.OpenRecordset(rsDefine, StrSQL, "InitAdviceDefine")
    Set InitAdviceDefine = rsDefine
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function FormatExamineAdvice(ByVal strAdvicePro As String, _
    ByVal strAdvicePart As String, ByVal lngExeType As Long) As String
'格式化医嘱检查内容
    Dim strReturn As String
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        Err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    If mrsDefine Is Nothing Then Set mrsDefine = InitAdviceDefine
    mrsDefine.Filter = "诊疗类别='D'"
    If mrsDefine.RecordCount > 0 Then
        strReturn = mrsDefine!医嘱内容 & ""
    End If

    If strReturn = "" Then
        strReturn = strAdvicePro & "," & _
                            Decode(lngExeType, 1, ",床旁执行", 2, ",术中执行", "") & IIF(strAdvicePart <> "", ":" & get部位方法(strAdvicePart), "")
    Else
        If InStr(strReturn, "[检查项目]") > 0 Then
            strReturn = Replace(strReturn, "[检查项目]", _
                                            """" & strAdvicePro & Decode(lngExeType, 1, ",床旁执行", 2, ",术中执行", "") & _
                                            """")
        End If

        '替换部位方法
        If InStr(strReturn, "[检查部位]") > 0 Then
            strReturn = Replace(strReturn, "[检查部位]", _
                                            """" & get部位方法(strAdvicePart) & """")
        End If

        strReturn = mobjVBA.Eval(strReturn)
    End If

    FormatExamineAdvice = strReturn
End Function

Public Function FormatInspectionAdvice(ByVal str检验 As String, ByVal str采集 As String, ByVal str标本 As String) As String
'功能：产生检验医嘱的医嘱内容
    Dim i As Long, strText As String, strField As String, blnDefine As Boolean
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        Err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    If mrsDefine Is Nothing Then Set mrsDefine = InitAdviceDefine
               
    '确定是否定义
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "诊疗类别='C'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(Nvl(mrsDefine!医嘱内容)) = "" Then
            blnDefine = False
        End If
    End If
    
    If Not blnDefine Then
        strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
    Else
        strText = mrsDefine!医嘱内容
        If InStr(strText, "[检验项目]") > 0 Then
            strField = str检验
            strText = Replace(strText, "[检验项目]", """" & strField & """")
        End If
        If InStr(strText, "[检验标本]") > 0 Then
            strField = str标本
            strText = Replace(strText, "[检验标本]", """" & strField & """")
        End If
        If InStr(strText, "[采集方法]") > 0 Then
            strField = str采集
            strText = Replace(strText, "[采集方法]", """" & strField & """")
        End If
        
        '计算医嘱内容
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
        End If
        Err.Clear: On Error GoTo 0
    End If
        
    FormatInspectionAdvice = strText
End Function

Public Function FormatOperationAdvice(ByVal str手术 As String, ByVal str麻醉 As String, ByVal str附术 As String, ByVal str手术时间 As String, ByVal str手术部位 As String) As String
'功能：产生检验医嘱的医嘱内容
    Dim i As Long, strText As String, strField As String, blnDefine As Boolean
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        Err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    If mrsDefine Is Nothing Then Set mrsDefine = InitAdviceDefine
               
    '确定是否定义
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "诊疗类别='F'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(Nvl(mrsDefine!医嘱内容)) = "" Then
            blnDefine = False
        End If
    End If
    If Not blnDefine Then
        strText = Format(str手术时间, "MM月dd日HH:mm")
        If str麻醉 <> "" Then
            strText = strText & IIF(str麻醉 <> "", " 在 " & str麻醉 & " 下行 ", " 行 ")
        End If
        strText = strText & str手术 & IIF(str手术部位 = "", "", "(部位:" & str手术部位 & ")")
        If str附术 <> "" Then
            strText = strText & " 及 " & str附术
        End If
    Else
        strText = mrsDefine!医嘱内容
        If InStr(strText, "[手术时间]") > 0 Then
            strField = str手术时间
            strText = Replace(strText, "[手术时间]", """" & strField & """")
        End If
        If InStr(strText, "[主要手术]") > 0 Then
            strField = str手术 & IIF(str手术部位 = "", "", "(部位:" & str手术部位 & ")")
            strText = Replace(strText, "[主要手术]", """" & strField & """")
        End If
        If InStr(strText, "[附加手术]") > 0 Then
            strField = str附术
            strText = Replace(strText, "[附加手术]", """" & strField & """")
        End If
        If InStr(strText, "[麻醉方法]") > 0 Then
            strField = str麻醉
            strText = Replace(strText, "[麻醉方法]", """" & strField & """")
        End If
        '计算医嘱内容
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            strText = Format(str手术时间, "MM月dd日HH:mm")
            If str麻醉 <> "" Then
                strText = strText & IIF(str麻醉 <> "", " 在 " & str麻醉 & " 下行 ", " 行 ")
            End If
            strText = strText & str手术 & IIF(str手术部位 = "", "", "(部位:" & str手术部位 & ")")
            If str附术 <> "" Then
                strText = strText & " 及 " & str附术
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
            
    FormatOperationAdvice = strText
End Function

Private Function get部位方法(ByVal strExtData As String) As String
'入:部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中
'出:部位名1(方法名1,方法名2),部位名2(方法名1,方法名2)-----
Dim i As Integer, strReturn As String, Arr部位
    If strExtData = "" Then Exit Function
    Arr部位 = Split(Split(strExtData, Chr(9))(0), "|")

    For i = 0 To UBound(Arr部位)
        strReturn = strReturn & "," & Split(Arr部位(i), ";")(0) & "(" & Split(Arr部位(i), ";")(1) & ")"
    Next

    get部位方法 = Mid(strReturn, 2)
End Function

Public Function Get执行内容(ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal str类别 As String _
       , ByVal str医嘱内容 As String, ByVal blnMove As Boolean) As String
'功能：根据指定的医嘱ID,返回执行医嘱内容供显示
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, strTmp As String
    Dim bln给药途径 As Boolean, i As Integer
    Dim str皮试结果 As String

    On Error GoTo errH
    
    '读取医嘱内容
    If (str类别 = "C" And lng相关ID <> 0) Or str类别 = "D" Then
        strTmp = str医嘱内容
        
    ElseIf str类别 <> "E" Or lng相关ID <> 0 Then
        '配方煎法,手术麻醉,输血途径,或其它医嘱,直接显示医嘱内容
        StrSQL = "Select 医嘱内容 From 病人医嘱记录 Where ID=[1]"
        If blnMove Then
            StrSQL = Replace(StrSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Get执行内容", IIF(str类别 = "E", lng相关ID, lng医嘱ID))
        If Not rsTmp.EOF Then strTmp = rsTmp!医嘱内容 & ""
    Else
        '类别为E,且相关ID=0
        StrSQL = "Select A.ID,A.相关ID,A.诊疗类别,A.医嘱内容,A.皮试结果,A.单次用量,B.计算单位,B.操作类型,A.执行频次,A.执行时间方案,B.名称" & _
            " From 病人医嘱记录 A,诊疗项目目录 B" & _
            " Where Not (A.诊疗类别='E' And 相关ID is Not NULL) And A.诊疗项目ID=B.ID" & _
            " And (A.相关ID=[1] Or A.ID=[1])" & _
            " Order by A.序号"
        If blnMove Then
            StrSQL = Replace(StrSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Get执行内容", lng医嘱ID)
        rsTmp.Filter = "相关ID=" & lng医嘱ID
        If Not rsTmp.EOF Then bln给药途径 = InStr(",5,6,", rsTmp!诊疗类别) > 0
        
        If Not bln给药途径 Then
            '一般治疗项目或中药用法，或采集方法
            rsTmp.Filter = 0
            If Not rsTmp.EOF Then
                If rsTmp!诊疗类别 = "E" And rsTmp!操作类型 = "1" Then
                    str皮试结果 = "，皮试结果：" & rsTmp!皮试结果
                    
                    StrSQL = "Select b.过敏反应, b.过敏时间 From 病人医嘱记录 A, 病人过敏记录 B, 诊疗项目目录 C, 诊疗用法用量 D" & _
                        " Where a.病人id = b.病人id And a.诊疗项目id = d.用法id And d.项目id = c.Id And c.类别 In ('5', '6') And d.项目id = b.药物id And" & _
                        " Nvl(d.性质, 0) = 0 And b.记录时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = a.id And 操作类型 = 10) And a.Id = [1] And RowNum<2"

                    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Get执行内容", lng医嘱ID)
                    
                    If Not rsTmp.EOF Then
                        str皮试结果 = str皮试结果 & ",过敏时间：" & Format(rsTmp!过敏时间, "yyyy-MM-dd") & IIF(rsTmp!过敏反应 & "" = "", "", ",过敏反应：" & rsTmp!过敏反应)
                    End If
                End If
            End If
            
            StrSQL = "Select 医嘱内容 From 病人医嘱记录 Where ID=[1]"
            If blnMove Then
                StrSQL = Replace(StrSQL, "病人医嘱记录", "H病人医嘱记录")
            End If
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Get执行内容", lng医嘱ID)
            If Not rsTmp.EOF Then strTmp = rsTmp!医嘱内容 & ""
        Else
            '给药途径
            For i = 1 To rsTmp.RecordCount
                strTmp = strTmp & vbCrLf & IIF(i = rsTmp.RecordCount, "┗", "┣") & rsTmp!医嘱内容 & IIF(Not IsNull(rsTmp!单次用量), " " & FormatEx(rsTmp!单次用量, 5) & rsTmp!计算单位, "")
                rsTmp.MoveNext
            Next
            rsTmp.Filter = "ID=" & lng医嘱ID
            strTmp = rsTmp!名称 & "," & rsTmp!执行频次 & "(" & rsTmp!执行时间方案 & "):每" & rsTmp!计算单位 & " " & Mid(strTmp, 2)
        End If
    End If
    
    Get执行内容 = strTmp & str皮试结果
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Sub GetAdvicePartSaveSql(ByVal lngAdviceID As Long, ByVal strProjectName As String, _
    ByRef curAdviceInf As clsExamineAdvice, ByRef arySql As Variant, ByVal lng序号 As Long, ByVal lng申请序号 As Long)
'获取部位医嘱的保存sql
'参数：lng序号=主医嘱记录的序号
    Dim i As Long, j As Long
    Dim str部位 As String
    Dim strTmp方法 As String
    Dim str方法 As String
    Dim lng医嘱序号 As Long
    Dim lngTmpID As Long
    Dim rsData As ADODB.Recordset
    Dim StrSQL As String

    lng医嘱序号 = lng序号

    StrSQL = "select id from 病人医嘱记录 where 相关id=[1]"
    Set rsData = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "查询部位医嘱", lngAdviceID)

    If rsData.RecordCount > 0 Then
        While Not rsData.EOF
            ReDim Preserve arySql(UBound(arySql) + 1)

            arySql(UBound(arySql)) = " ZL_病人医嘱记录_Delete(" & Val(Nvl(rsData!ID)) & ")"

            Call rsData.MoveNext
        Wend
    End If


    '组织部位插入语句
    For i = 0 To UBound(Split(curAdviceInf.部位方法, "|")) '部位1;方法1,方法2,方法3|部位n;方法1,方法2,方法3---

        str部位 = Split(Split(curAdviceInf.部位方法, "|")(i), ";")(0)
        strTmp方法 = Split(Split(curAdviceInf.部位方法, "|")(i), ";")(1)

        For j = 0 To UBound(Split(strTmp方法, ","))
            lng医嘱序号 = lng医嘱序号 + 1     '病人医嘱记录.序号，递增
            str方法 = Split(strTmp方法, ",")(j)
            lngTmpID = gobjComlib.zlDatabase.GetNextID("病人医嘱记录")

            ReDim Preserve arySql(UBound(arySql) + 1)

            arySql(UBound(arySql)) = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                 lng医嘱序号 & "," & curAdviceInf.病人来源 & "," & curAdviceInf.病人ID & "," & IIF(curAdviceInf.主页ID = 0, "NULL", curAdviceInf.主页ID) & "," & _
                 curAdviceInf.婴儿 & ",1,1,'D'," & curAdviceInf.检查项目ID & ",NULL,NULL,NULL,1," & _
                 "'" & strProjectName & "',NULL," & _
                 "'" & str部位 & "','一次性',NULL,NULL,NULL,NULL,0," & _
                 curAdviceInf.执行科室ID & "," & IIF(curAdviceInf.执行科室ID <= 0, "5", curAdviceInf.执行科室性质) & "," & curAdviceInf.紧急标志 & ",to_date('" & Format(curAdviceInf.开始时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),NULL," & _
                 curAdviceInf.病人科室ID & "," & curAdviceInf.开单科室ID & _
                 ",'" & curAdviceInf.开嘱医生 & "',to_date('" & Format(curAdviceInf.开嘱时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),'" & curAdviceInf.挂号单 & "',Null,'" & str方法 & "'," & curAdviceInf.执行类型 & ",NULL,NULL,'',NULL,NULL,NULL,NULL," & lng申请序号 & ")"
        Next
    Next

End Sub

Public Sub GetAdviceAffixSaveSql(ByVal lngAdviceID As Long, ByRef arrSQL As Variant, ByVal str附项 As String)
'获取医嘱附件的存储sql
    Dim arrAppend As Variant
    Dim j As Long
    
    arrAppend = Array()
    If str附项 <> "" Then
        arrAppend = Split(str附项, "<Split1>")
        For j = 0 To UBound(arrAppend)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & lngAdviceID & "," & _
                "'" & Split(arrAppend(j), "<Split2>")(0) & "'," & Val(Split(arrAppend(j), "<Split2>")(1)) & "," & _
                j + 1 & "," & ZVal(Split(arrAppend(j), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(j), "<Split2>")(3), "'", "''") & "'" & _
                IIF(j = 0, ",1", "") & ")"
        Next
    End If
End Sub

Public Function Check上班安排(ByVal bln药房 As Boolean) As Boolean
'功能：检查医院的科室是否使用了上班安排
'参数：bln药房=是检查药房上班还是其它科室
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    Static bln药房Load As Boolean
    Static bln药房Last As Boolean
    Static bln非药Load As Boolean
    Static bln非药Last As Boolean
    
    If bln药房 Then '是否有安排只需读取一次
        If bln药房Load Then Check上班安排 = bln药房Last: Exit Function
    Else
        If bln非药Load Then Check上班安排 = bln非药Last: Exit Function
    End If
    
    On Error GoTo errH
    
    If bln药房 Then
        StrSQL = "Select 1 From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 IN('西药房','成药房','中药房') And Rownum<2"
    Else
        StrSQL = "Select 1 From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 Not IN('西药房','成药房','中药房') And Rownum<2"
    End If
    Call gobjComlib.zlDatabase.OpenRecordset(rsTmp, StrSQL, "Check上班安排")
    Check上班安排 = rsTmp.RecordCount > 0
    
    If bln药房 Then
        bln药房Load = True: bln药房Last = Check上班安排
    Else
        bln非药Load = True: bln非药Last = Check上班安排
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get操作员部门ID(ByVal int服务对象 As Integer) As Long
'功能：取操作员所属服务对指定对象的部门，缺省部门优先
    Static rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    
    If blnNew Then
        StrSQL = "Select Distinct B.部门ID,Nvl(B.缺省,0) as 缺省,C.服务对象 From 部门人员 B,部门性质说明 C" & _
            " Where B.人员ID = [1] And B.部门ID=C.部门ID" & _
            " Order by 缺省 Desc"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", UserInfo.ID)
    End If
    rsTmp.Filter = "服务对象 = 3 or 服务对象 = " & int服务对象
    
    If Not rsTmp.EOF Then
        Get操作员部门ID = rsTmp!部门ID
    Else
        Get操作员部门ID = UserInfo.部门ID
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
'功能：读取当前登录人员或指定人员的人员性质
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
        
    On Error GoTo errH
    If str姓名 <> "" Then
        StrSQL = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", str姓名)
    Else
        StrSQL = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckPatiDataMoved(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：判断指定病人的数据是否已转出
    Dim rsTmp As ADODB.Recordset, StrSQL As String
 
    StrSQL = "Select 数据转出 From 病案主页 Where 病人ID = [1] And 主页ID = [2]"
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "检查转出", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        CheckPatiDataMoved = Val("" & rsTmp!数据转出) = 1
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub InitSendRecordset(rsExec As Recordset, rsBill As Recordset, rsRXKey As Recordset, rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset, _
    rsNumber As ADODB.Recordset, rsMoneyNow As ADODB.Recordset, rsItems As ADODB.Recordset)
'功能：初始化医嘱发送所需的动态记录集
    '初始化医嘱计价记录集
    Set rsExec = New ADODB.Recordset
    
    rsExec.Fields.Append "医嘱ID", adBigInt
    rsExec.Fields.Append "发送号", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "要求时间", adDate, , adFldIsNullable
    rsExec.Fields.Append "数量", adDouble, , adFldIsNullable
    rsExec.Fields.Append "费用性质", adInteger, , adFldIsNullable
    
    rsExec.CursorLocation = adUseClient
    rsExec.LockType = adLockOptimistic
    rsExec.CursorType = adOpenStatic
    rsExec.Open
    
    '初始化医嘱记帐单据生成记录集
    Set rsBill = New ADODB.Recordset
    
    rsBill.Fields.Append "Key", adVarChar, 100
    rsBill.Fields.Append "NO", adVarChar, 30
    rsBill.Fields.Append "费用序号", adBigInt
    rsBill.Fields.Append "发送序号", adBigInt
    rsBill.CursorLocation = adUseClient
    rsBill.LockType = adLockOptimistic
    rsBill.CursorType = adOpenStatic
    rsBill.Open
        
    Set rsRXKey = New ADODB.Recordset
    rsRXKey.Fields.Append "Key", adVarChar, 200
    rsRXKey.Fields.Append "医嘱ID", adVarChar, 200
    rsRXKey.Fields.Append "条数", adBigInt
    rsRXKey.Fields.Append "张数", adBigInt
    rsRXKey.CursorLocation = adUseClient
    rsRXKey.LockType = adLockOptimistic
    rsRXKey.CursorType = adOpenStatic
    rsRXKey.Open
    
    'SQL记录集
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "类型", adInteger '1-计价,2-签名,3-校对,4-发送,5-费用,6-发料
    rsSQL.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsSQL.Fields.Append "项目ID", adBigInt '收费细目ID
    rsSQL.Fields.Append "序号", adBigInt '用于排序
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '用于NO替换处理时排序
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '计价数量累计记录集
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsTotal.Fields.Append "项目ID", adBigInt
    rsTotal.Fields.Append "库房ID", adBigInt
    rsTotal.Fields.Append "数量", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    '医保上传记帐单
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsUpload.Fields.Append "NO", adVarChar, 30
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
    
    '计录试管编码
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "管码", adVarChar, 18
    rsNumber.Fields.Append "相关ID", adBigInt
    rsNumber.Fields.Append "样本条码", adVarChar, 18
    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
    rsNumber.Fields.Append "婴儿", adBigInt
    rsNumber.Fields.Append "紧急标志", adBigInt
    rsNumber.Fields.Append "标本", adVarChar, 18
    rsNumber.Fields.Append "采集科室ID", adBigInt
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    '当前病人本次要发送的费用
    Set rsMoneyNow = New ADODB.Recordset
    rsMoneyNow.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsMoneyNow.Fields.Append "诊疗项目ID", adBigInt
    rsMoneyNow.Fields.Append "收费项目ID", adBigInt
    rsMoneyNow.Fields.Append "试管编码", adVarChar, 18, adFldIsNullable
    rsMoneyNow.Fields.Append "收费方式", adInteger
    rsMoneyNow.Fields.Append "收费时间", adVarChar, 10
    rsMoneyNow.Fields.Append "执行部门ID", adBigInt
    rsMoneyNow.CursorLocation = adUseClient
    rsMoneyNow.LockType = adLockOptimistic
    rsMoneyNow.CursorType = adOpenStatic
    rsMoneyNow.Open
    
    '当前病人本次发送的费用项目汇总
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "病人ID", adBigInt
    rsItems.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "医嘱ID", adBigInt
    rsItems.Fields.Append "收费类别", adVarChar, 1
    rsItems.Fields.Append "收费细目ID", adBigInt
    rsItems.Fields.Append "数量", adDouble
    rsItems.Fields.Append "单价", adDouble
    rsItems.Fields.Append "实收金额", adDouble
    rsItems.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
    
End Sub

Public Function GetTubeMaterial(ByVal str试管编码 As String) As Long
'功能：根据管码获取对应的试管材料ID
    Dim StrSQL As String, rsTube As Recordset
    
    On Error GoTo errH
    
    StrSQL = "Select 编码,材料ID From 采血管类型 Where 材料ID is Not NULL and 编码=[1]"
    Set rsTube = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "GetTubeMaterial", str试管编码)
    
    If Not rsTube.EOF Then GetTubeMaterial = Nvl(rsTube!材料ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get收费执行科室ID(ByVal lng病人ID As Long, lng主页ID As Long, _
    ByVal str类别 As String, ByVal lng项目ID As Long, ByVal int执行科室 As Integer, _
    ByVal lng病人科室ID As Long, ByVal lng开单科室ID As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal lng执行科室ID As Long, _
    Optional ByVal bytMode As Byte, Optional ByVal bytCallBy As Byte, _
    Optional ByVal int调用场合 As Integer = 1, _
    Optional lng成套缺省执行科室 As Long = 0) As Long
'功能：获取非药收费项目的执行科室
'参数：int范围=1.门诊,2-住院
'      lng执行科室ID=指定的缺省执行科室ID(用于药品和卫材)
'      bytMode=1-要返回缺省值,0-其它
'      bytCallBy=0-医嘱程序调用,1-附费程序调用
'      int调用场合=1-门诊,2-住院
'      lng成套缺省执行科室-缺省执行科室ID
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Integer
    Dim str药房 As String, lng药房 As Long
    Dim lng病人病区ID As Long, bytDay As Byte
    
    On Error GoTo errH
    
    If str类别 = "4" Then
        lng药房 = Val(gobjComlib.zlDatabase.GetPara(IIF(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省发料部门", glngSys, _
            IIF(bytCallBy = 1, p医嘱附费管理, IIF(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达))))
        
        '有执行科室设置时
        StrSQL = _
            " Select Distinct" & _
            "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", int范围, lng病人科室ID, lng项目ID)
        If Not rsTmp.EOF Then
            If bytMode = 1 Then Get收费执行科室ID = rsTmp!执行科室ID  '如果都没有，则返回第一个可用的执行科室
            
            '1:缺省为指定的(医嘱的)执行科室,不管是否服务于病人科室
            rsTmp.Filter = "执行科室ID=" & lng执行科室ID
            
            '2.缺省为参数指定的缺省科室
            If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng药房
            
            '3:其它可服务于病人科室的执行科室
            If rsTmp.EOF Then
                '2.0 如果成套中存在缺省的执行科室,则缺省为成套指定的缺省科室
                If lng成套缺省执行科室 <> 0 Then
                    rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                    If Not rsTmp.EOF Then
                            Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                    End If
                End If
                '2.1:尝试缺省为病人科室
                If lng执行科室ID <> lng病人科室ID And lng药房 <> lng病人科室ID Then
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人科室ID
                End If
                '3.2:尝试缺省为病人病区
                If rsTmp.EOF And lng主页ID <> 0 Then
                    lng病人病区ID = GetPatiUnitID(lng病人ID, lng主页ID)
                    If lng病人病区ID <> 0 And lng病人病区ID <> lng病人科室ID And lng病人病区ID <> lng执行科室ID And lng病人病区ID <> lng药房 Then
                        rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人病区ID
                    End If
                End If
            End If
            '3.3:可服务于病人科室的一个执行科室
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            
            '3.4可服务于所有科室的当前病人科室执行
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0 And 执行科室ID=" & lng病人科室ID
            
            '4:如果都没有，则返回0用于检查
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        If str类别 = "5" Then
            str药房 = "西药房"
            lng药房 = Val(gobjComlib.zlDatabase.GetPara(IIF(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省西药房", glngSys, _
                IIF(bytCallBy = 1, p医嘱附费管理, IIF(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达)), , , , , lng病人科室ID))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(gobjComlib.zlDatabase.GetPara(IIF(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省成药房", glngSys, _
                IIF(bytCallBy = 1, p医嘱附费管理, IIF(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达)), , , , , lng病人科室ID))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(gobjComlib.zlDatabase.GetPara(IIF(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省中药房", glngSys, _
                IIF(bytCallBy = 1, p医嘱附费管理, IIF(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达)), , , , , lng病人科室ID))
        End If
        
        '药品从系统指定的储备药房中找
        If Not Check上班安排(True) Then
            StrSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                " And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(gobjComlib.zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            StrSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                " And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", str药房, int范围, lng病人科室ID, lng项目ID, bytDay)
        If Not rsTmp.EOF Then
            If lng成套缺省执行科室 <> 0 Then
                rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                If Not rsTmp.EOF Then
                        Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                End If
            End If
            Get收费执行科室ID = rsTmp!执行科室ID
            rsTmp.Filter = "执行科室ID=" & lng执行科室ID
            If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng药房
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    Else
        Select Case int执行科室
            Case 0 '0-无明确科室
                Get收费执行科室ID = Get操作员部门ID(int范围)
            Case 1 '1-病人所在科室
                Get收费执行科室ID = lng病人科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get收费执行科室ID = lng病人科室ID
                Else
                    Get收费执行科室ID = GetPatiUnitID(lng病人ID, lng主页ID)
                End If
            Case 3 '3-操作员所在科室
                Get收费执行科室ID = Get操作员部门ID(int范围)
            Case 4 '4-指定科室
                StrSQL = "Select Distinct Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID,Decode(A.病人来源,Null,2,1) as 排序" & _
                    " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                    " Where A.收费细目ID=[1] And A.执行科室ID=B.部门ID" & _
                    " And B.服务对象 IN([2],3) And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                    " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                    " And A.执行科室ID=C.ID And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                    " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " Order by 排序" '默认科室优先
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng项目ID, int范围, lng病人科室ID)
                If Not rsTmp.EOF Then
                    If lng成套缺省执行科室 <> 0 Then
                         rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                         If Not rsTmp.EOF Then
                                 Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                         End If
                     End If
                    Get收费执行科室ID = rsTmp!执行科室ID
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                    If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
                End If
            Case 6 '6-开单人所在科室
                Get收费执行科室ID = lng开单科室ID
        End Select
        If Get收费执行科室ID = 0 Then Get收费执行科室ID = Get操作员部门ID(int范围)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function CalcDrugPrice(ByVal lng药品ID As Long, lng药房ID As Long, ByVal dbl数量 As Double, _
    Optional ByVal str费别 As String, Optional ByVal blnNone加班加价 As Boolean, Optional ByVal str药品价格等级 As String, Optional ByVal str卫材价格等级 As String, Optional ByVal str普通项目价格等级 As String) As Double
'功能：计算药品实价(即然要计算实价,药品则肯定为变价)，传入费别时，则计算实收金额
'参数：dbl数量=售价数量,按费别打折时计算的是实收金额
'      str费别=是否按费别计算打折的价格,主要在直接计算药品的金额而不显示单价时用
'      gbln加班加价=发送时计算才有用,其它地方都为False
'      blnNone加班加价=为真时,"gbln加班加价"无效
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long
    Dim dbl总数量 As Double, dbl当前数量 As Double
    Dim dbl总金额 As Double, dbl时价 As Double
    Dim dbl加班加价率 As Double
    Dim dbl首批时价 As Double, intCount As Integer
        
    If dbl数量 = 0 Then Exit Function
    
    On Error GoTo errH
    
    StrSQL = _
        " Select Nvl(批次,0) as 批次,Nvl(可用数量,0) as 库存," & _
        " Nvl(零售价,Nvl(Decode(Nvl(实际数量,0),0,0,实际金额/实际数量),0)) as 时价" & _
        " From 药品库存" & _
        " Where 库房ID=[1] And 药品ID=[2] And Nvl(可用数量,0)>0" & _
        " And 性质=1 And (Nvl(批次,0)=0 Or 效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Order by " & IIF(gbytMediOutMode = 1, "效期,", "") & "Nvl(批次,0)"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "CalcDrugPrice", lng药房ID, lng药品ID)
    
    dbl总金额 = 0: dbl总数量 = dbl数量
    For i = 1 To rsTmp.RecordCount
        '第一个批次的时价
        intCount = intCount + 1
        If intCount = 1 Then
            dbl首批时价 = Format(rsTmp!时价, gstrDecPrice)
        End If
        If dbl总数量 = 0 Then Exit For '为了始终取到首批时价
        
        If dbl总数量 <= rsTmp!库存 Then
            dbl当前数量 = dbl总数量
        Else
            dbl当前数量 = rsTmp!库存
        End If
        dbl总金额 = dbl总金额 + Format(dbl当前数量 * Format(rsTmp!时价, gstrDecPrice), gstrDec)
        dbl总数量 = Val(dbl总数量) - Val(dbl当前数量)
        If dbl总数量 = 0 Then Exit For
        
        rsTmp.MoveNext
    Next
    
    If dbl总数量 <> 0 Then
        '库存不够,只涉及一个批次时以首批时价为准，否则以第一批或者平均价都不合适
        dbl时价 = IIF(intCount = 1, dbl首批时价, 0)
    Else
        dbl时价 = IIF(intCount = 1, dbl首批时价, Format(dbl总金额 / dbl数量, gstrDecPrice))
        
        '当有费别参数时，是结合数量计算打折实收金额
        If str费别 <> "" Then
            dbl时价 = Format(dbl时价 * dbl数量, gstrDec)
            
            StrSQL = _
                " Select A.屏蔽费别,B.收入项目ID" & _
                " From 收费项目目录 A,收费价目 B" & _
                " Where A.ID=B.收费细目ID And A.ID=[1]" & _
                GetPriceGradeSQL(str药品价格等级, str卫材价格等级, str普通项目价格等级, "A", "B", "2", "3", "4") & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "CalcDrugPrice", lng药品ID, str药品价格等级, str卫材价格等级, str普通项目价格等级)
            If rsTmp.EOF Then Exit Function
            
            '根据费别重新计算实收金额
            If Not (Nvl(rsTmp!屏蔽费别, 0) = 1) Then
                dbl时价 = ActualMoney(str费别 & IIF(gstr动态费别 <> "", "," & gstr动态费别, ""), rsTmp!收入项目ID, dbl时价, lng药品ID, lng药房ID, dbl数量)
            End If
        End If
    End If
    CalcDrugPrice = dbl时价
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Calc次数分解时间(lng次数 As Long, ByVal dat开始时间 As Date, dat终止时间 As Date, strPause As String, _
    ByVal str执行时间 As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, _
    Optional ByVal dat首日日期 As Date) As String
'功能：按次数计算各次的分解执行时间,要求<=终止时间及不在暂停时间段内
'参数：dat开始时间=医嘱的开始执行时间
'      dat终止时间=医嘱的执行终止时间,没有时传入"3000-01-01"
'      strPause=医嘱的暂停时间段
'      dat首日日期=用于首日时间计算参照
'返回：1."时间1,时间2,...."(yyyy-MM-dd HH:mm:ss)
'      2.lng次数=实际能够分解的次数
'说明：1.因为终止时间的限制,因此分解出来的时间个数可能小于要分解的次数
'      2.本函数是假定在执行时间及频率性质完全正确的情况下计算。
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrFirst As Variant, arrNormal As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(str执行时间, ",") > 0 Then
        arrNormal = Split(Split(str执行时间, ",")(1), "-")
        arrFirst = Split(Split(str执行时间, ",")(0), "-")
    Else
        arrNormal = Split(str执行时间, "-")
        arrFirst = Array()
    End If
    
    vCurTime = dat开始时间
    
    If str间隔单位 = "周" Then
        vCurTime = gobjComlib.ZLCommFun.GetWeekBase(dat开始时间)
        
        Do While lng次数 > 0
            blnFirst = (gobjComlib.ZLCommFun.GetWeekBase(vCurTime) = gobjComlib.ZLCommFun.GetWeekBase(dat首日日期)) And dat首日日期 <> Empty And UBound(arrFirst) <> -1
            arrTime = IIF(blnFirst, arrFirst, arrNormal)

            '1/8:00-3/15:00-5/9:00
            For i = 1 To int频率次数
                If i - 1 <= UBound(arrTime) Then '首周可能次数不足
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime > dat终止时间 Then
                        Exit Do
                    ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                        Exit Do
                    ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        lng次数 = lng次数 - 1
                        If lng次数 = 0 Then Exit Do
                    End If
                End If
            Next
            vCurTime = vCurTime + 7
        Loop
    ElseIf str间隔单位 = "天" Then
        Do While lng次数 > 0
            blnFirst = (Int(vCurTime) = Int(dat首日日期)) And dat首日日期 <> Empty And UBound(arrFirst) <> -1
            arrTime = IIF(blnFirst, arrFirst, arrNormal)
        
            If int频率间隔 = 1 Then
                '8:00-12:00-14:00；8-12-14
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        
                        If vTmpTime > dat终止时间 Then
                            Exit Do
                        ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                            Exit Do
                        ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            lng次数 = lng次数 - 1
                            If lng次数 = 0 Then Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime > dat终止时间 Then
                            Exit Do
                        ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                            Exit Do
                        ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            lng次数 = lng次数 - 1
                            If lng次数 = 0 Then Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = vCurTime + int频率间隔
        Loop
    ElseIf str间隔单位 = "小时" Then
        '10:00-20:00-40:00；10-20-40；02:30
        arrTime = arrNormal
        Do While lng次数 > 0
            For i = 1 To int频率次数
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime > dat终止时间 Then
                    Exit Do
                ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                    Exit Do
                ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    lng次数 = lng次数 - 1
                    If lng次数 = 0 Then Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + int频率间隔 / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str间隔单位 = "分钟" Then
        '无执行时间
        Do While lng次数 > 0
            vTmpTime = vCurTime
            
            If vTmpTime > dat终止时间 Then
                Exit Do
            ElseIf TimeisLastPause(vTmpTime, strPause) And dat终止时间 = CDate("3000-01-01") Then
                Exit Do
            ElseIf vTmpTime >= dat开始时间 And Not TimeIsPause(vTmpTime, strPause) Then
                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                lng次数 = lng次数 - 1
                If lng次数 = 0 Then Exit Do
            End If

            vCurTime = Format(vCurTime + int频率间隔 / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If

    lng次数 = UBound(Split(Mid(strDetailTime, 2), ",")) + 1
    Calc次数分解时间 = Mid(strDetailTime, 2)
End Function

Public Function AdviceMoneyMake(ByVal lng病人ID As Long, ByVal lng主页ID As Long, rsMoneyNow As Recordset, rsMoneyDay As ADODB.Recordset, _
    ByVal lng医嘱ID As Long, ByVal lng诊疗项目ID, ByVal lng收费项目ID As Long, ByVal lng执行部门id As Long, ByVal str试管编码 As String, _
    ByVal str收费类别 As String, ByVal int收费方式 As Integer, ByVal str分解时间 As String, ByVal byt来源 As Byte, ByRef lng费用次数 As Long, ByVal dbl总量 As Double, _
    Optional ByVal lng当前医嘱ID As Long, Optional ByVal lng发送号 As Long, Optional ByVal dbl计价数量 As Double, Optional rsExec As Recordset, _
    Optional ByVal lng计算方式 As Long, Optional ByVal str频率 As String, Optional ByVal dbl单量 As Double, Optional ByVal int期效 As Integer = 1, _
    Optional ByVal int费用性质 As Integer, Optional ByVal str诊疗类别 As String, Optional ByVal str样本条码 As String) As Boolean
'功能：判断指定的医嘱费用是否应该产生
'参数：lng主页ID=住院病人才使用，门诊病人传入0不分具体挂号
'      rsMoneyNow=当前病人本次要发送的费用,动态记录集(收费方式=-1,表示首次不收时，一天只收一次的项目的记录)
'      rsMoneyDay=当前病人当天已发送的费用,静态记录集
'      lng医嘱ID=一组医嘱的ID
'      str分解时间=本次发送的执行时间串，以逗号分隔，并且排除了暂停的时间点
'      byt来源:1-门诊，2-住院
'      dbl计价数量=收费项目的计价数量
'      其他=当前行发送医嘱及费用信息
'      lng当前医嘱ID=当前行医嘱id
'      str样本条码=检验医嘱传入样本条码
'以下是计量计时医嘱的数量组织规则
'1、长嘱可选频率、持续性、必要时和不定时以单量作为数次。
'2、临嘱一次性和需要时频率的医嘱取总量作为数次。
'3、临嘱可选频率取单量作为数次，最后一次取总量除以单量取末作为数次，例如红外照射治疗，总量80、单量25，每天4次，那么执行登记时，供执行四次，前三次本次数次为25，第四次为80除以25取模=5。
'4、批量执行登记页面医嘱清单单量后新增列：本次数次，用于显示本次数次。
'5、医嘱编辑时不允许录入首次用量。
'返回：
'      lng费用次数=一天只收一次时（3,4,5,6,7），返回本次发送要收取的次数
'      dbl总量=总的发送次数或数量
'      rsExec=医嘱执行计价的内容
    Dim lng材料ID As Long, blnMakeMoney As Boolean
    Dim rsDays As ADODB.Recordset, i As Long
    Dim arrTmp As Variant
    Dim dbl数量 As Double
    Dim strDate As String
    Dim dbl总量Tmp As Double
    Dim StrSQL As String, rsTmp As Recordset, strTmp As String
    
    blnMakeMoney = True
    lng费用次数 = 1
    
    If int收费方式 = 9 Then
        '自定义
        On Error GoTo errH
        
        StrSQL = "Select zl_fun_CustomExpenses([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17]) as 返回结果 From Dual"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "AdviceMoneyMake", lng病人ID, lng主页ID, byt来源, lng当前医嘱ID, lng医嘱ID, int期效, str频率, lng诊疗项目ID, lng收费项目ID, _
                                            lng执行部门id, str诊疗类别, str收费类别, dbl总量, dbl单量, dbl计价数量, int费用性质, lng计算方式)
        If rsTmp.RecordCount > 0 Then
            strTmp = rsTmp!返回结果 & ""
            
            If Val(Split(strTmp, ":")(0)) = 0 Then
                '不收取
                blnMakeMoney = False
            Else
                '要收取
                If InStr(strTmp, ":") > 0 Then
                    If Val(Split(strTmp, ":")(1)) > 0 Then lng费用次数 = Val(Split(strTmp, ":")(1))
                End If
            End If
        End If
    End If
    
    If int收费方式 = 0 Then
        '正常收费的，检查在本次发送中、本医嘱中是否被排斥
        rsMoneyNow.Filter = "(医嘱ID=" & lng医嘱ID & " And 诊疗项目ID=" & lng诊疗项目ID & " And 收费方式=5)" & _
            " Or (医嘱ID=" & lng医嘱ID & " And 诊疗项目ID=" & lng诊疗项目ID & " And 收费方式=6)" 'Or的使用
        If Not rsMoneyNow.EOF Then blnMakeMoney = False
    ElseIf int收费方式 = 1 Then '检验试管费用(一次发送只收取一次)
        If str试管编码 <> "" Then
            '相同条码(试管)只收取一次
            rsMoneyNow.Filter = "试管编码='" & str试管编码 & "' And 样本条码='" & str样本条码 & "' And 收费项目ID=" & lng收费项目ID & " And 收费方式<>-1"
            If Not rsMoneyNow.EOF Then blnMakeMoney = False
            
            '只收取试管对应的卫材费用
            If blnMakeMoney And str收费类别 = "4" Then
                lng材料ID = GetTubeMaterial(str试管编码)
                If lng材料ID <> 0 And lng收费项目ID <> lng材料ID Then blnMakeMoney = False
            End If
        End If
    ElseIf int收费方式 = 2 Then '一次发送只收取一次
        rsMoneyNow.Filter = "诊疗项目ID=" & lng诊疗项目ID & " And 收费项目ID=" & lng收费项目ID & " And 收费方式<>-1"
        If Not rsMoneyNow.EOF Then blnMakeMoney = False
    ElseIf InStr(",3,4,5,6,7,", int收费方式) > 0 Then
        '3-当天只收取一次；4-当天未执行收取一次；5-当天只收取一次，排斥其他项目；6-当天未执行收取一次，排斥其他项目
        
        '正常收费的，检查在本次发送中、本医嘱中是否被排斥
        If int收费方式 = 7 Then
            rsMoneyNow.Filter = "(医嘱ID=" & lng医嘱ID & " And 诊疗项目ID=" & lng诊疗项目ID & " And 收费方式=5)" & _
                " Or (医嘱ID=" & lng医嘱ID & " And 诊疗项目ID=" & lng诊疗项目ID & " And 收费方式=6)" 'Or的使用
            If Not rsMoneyNow.EOF Then blnMakeMoney = False
        End If
        
        If blnMakeMoney Then
            Set rsDays = GetExecDays(str分解时间)
                        
            '先从本次发送中的找(频率为一天一次且没有收的，判断时当成已收取,以便后续的其他医嘱"首次不收"时不再认为有首次)
            For i = 1 To rsDays.RecordCount
                rsMoneyNow.Filter = "收费时间='" & rsDays!收费时间 & "' And 诊疗项目ID=" & lng诊疗项目ID & " And 收费项目ID=" & lng收费项目ID & _
                    IIF(int收费方式 = 7, "", " And 收费方式<>-1") & _
                    IIF((int收费方式 = 4 Or int收费方式 = 6) And lng执行部门id <> 0, " And 执行部门ID=" & lng执行部门id, "")
                If rsMoneyNow.RecordCount > 0 Then rsDays!存在 = 1
                rsDays.MoveNext
            Next
            '再从已发送中的找(当天及将来执行的)
            rsDays.Filter = "存在=0"
            For i = 1 To rsDays.RecordCount
                If i = 1 Then
                    If rsMoneyDay Is Nothing Then
                        Call GetPatiDayMoneyDetail(rsMoneyDay, lng病人ID, lng主页ID, byt来源, CDate(rsDays!收费时间 & ""))
                    End If
                End If
                rsMoneyDay.Filter = "收费时间='" & rsDays!收费时间 & "' And 诊疗项目ID=" & lng诊疗项目ID & " And 收费项目ID=" & lng收费项目ID & _
                    IIF(int收费方式 = 7, "", " And 收费方式<>-1") & _
                    IIF((int收费方式 = 4 Or int收费方式 = 6) And lng执行部门id <> 0, " And 执行否=0 And 执行部门ID=" & lng执行部门id, "")
                If rsMoneyDay.RecordCount > 0 Then rsDays!存在 = 1
                rsDays.MoveNext
            Next
        End If
    End If
                            
    '记录到本次发送明细项目记录中
    If InStr(",3,4,5,6,7,", int收费方式) > 0 Then
        If int收费方式 = 7 Then
            If blnMakeMoney Then
                rsDays.Filter = "存在=0"    '没收过的那些天(频率为一天一次但未收的当成收过了)，首次不收
                lng费用次数 = dbl总量 - rsDays.RecordCount
                blnMakeMoney = lng费用次数 > 0
            End If
        Else
            rsDays.Filter = "存在=0"
            blnMakeMoney = rsDays.RecordCount > 0
            lng费用次数 = rsDays.RecordCount    '一天一次，有多少天要收就有多少次
        End If
        If blnMakeMoney Or int收费方式 = 7 And lng费用次数 = 0 Then
            For i = 1 To rsDays.RecordCount
                rsMoneyNow.AddNew
                rsMoneyNow!医嘱ID = lng医嘱ID
                rsMoneyNow!诊疗项目ID = lng诊疗项目ID
                rsMoneyNow!收费项目ID = lng收费项目ID
                rsMoneyNow!试管编码 = str试管编码
                rsMoneyNow!样本条码 = str样本条码
                
                '首次不收时，如果频率为一天一次，则计算后的费用次数为0,为了让本次后续发送的其他医嘱正确计算首是否收取，需要产生记录，但收费方式特殊记录为-1
                rsMoneyNow!收费方式 = IIF(int收费方式 = 7 And lng费用次数 = 0, -1, int收费方式)
                rsMoneyNow!收费时间 = rsDays!收费时间
                rsMoneyNow!执行部门ID = lng执行部门id
                rsMoneyNow.Update
            
                rsDays.MoveNext
            Next
        End If
    ElseIf blnMakeMoney Then
        rsMoneyNow.AddNew
        rsMoneyNow!医嘱ID = lng医嘱ID
        rsMoneyNow!诊疗项目ID = lng诊疗项目ID
        rsMoneyNow!收费项目ID = lng收费项目ID
        rsMoneyNow!试管编码 = str试管编码
        rsMoneyNow!样本条码 = str样本条码
        rsMoneyNow!收费方式 = int收费方式
        If str分解时间 <> "" Then
            rsMoneyNow!收费时间 = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd")  '此时间暂时没有用处
        Else
            rsMoneyNow!收费时间 = ""
        End If
        rsMoneyNow!执行部门ID = lng执行部门id
        rsMoneyNow.Update
    End If
    '读取医嘱执行计价(除药品卫材医嘱外的才存储)
    If InStr(",5,6,7,", "," & str诊疗类别 & ",") = 0 Then
        If str分解时间 <> "" And Not rsExec Is Nothing Then
            arrTmp = Split(str分解时间, ",")
            dbl总量Tmp = dbl总量
            For i = 0 To UBound(arrTmp)
                rsExec.AddNew
                rsExec!医嘱ID = lng当前医嘱ID
                rsExec!发送号 = lng发送号
                rsExec!要求时间 = Format(arrTmp(i), "yyyy-MM-dd HH:mm:ss")
                rsExec!收费细目ID = lng收费项目ID
                rsExec!费用性质 = int费用性质
                If blnMakeMoney Then
                    '卫材也可以输入单量总量
                    If str频率 <> "" And (lng计算方式 = 0 And dbl总量 > 0 Or lng计算方式 = 1 Or lng计算方式 = 2 Or str诊疗类别 = "4") Then
                        '计量和计时的需要乘以数次
                        If int期效 = 0 Then
                            '1、长嘱可选频率、持续性、必要时和不定时以单量作为数次。
                            dbl数量 = dbl计价数量 * dbl单量
                        ElseIf InStr("一次性,需要时", str频率) Then
                            '2、临嘱一次性和需要时频率的医嘱取总量作为数次。
                            dbl数量 = dbl计价数量 * dbl总量
                        Else
                            '3、临嘱可选频率取单量作为数次，最后一次剩余的数量，例如红外照射治疗，总量80、单量25，每天4次，那么执行登记时，供执行四次，前三次本次数次为25，第四次为80除以25取模=5。
                            '门诊有可能没有录入执行时间,分解时间就只有一个，按总量作为次数
                            If UBound(arrTmp) = 0 Then
                                dbl数量 = dbl计价数量 * dbl总量
                            Else
                                If i = UBound(arrTmp) Then
                                    dbl数量 = dbl总量Tmp
                                Else
                                    If dbl总量Tmp >= dbl单量 Then
                                        dbl数量 = dbl计价数量 * dbl单量
                                    Else
                                        dbl数量 = dbl总量Tmp
                                    End If
                                    dbl总量Tmp = dbl总量Tmp - dbl数量
                                End If
                            End If
                        End If
                    Else
                        dbl数量 = dbl计价数量
                    End If
                    If i <> 0 Then
                        strDate = Format(arrTmp(i - 1), "yyyy-MM-dd")
                    End If
                    '一次发送收取一次，则只有第一次收取
                    If InStr(",1,2,", int收费方式) > 0 Then
                        If i <> 0 Then dbl数量 = 0
                    ElseIf InStr(",3,4,5,6,", int收费方式) > 0 Then
                        '3456当天只收取一次的，存在=0的收取，默认第一次有数量
                        rsDays.Filter = "存在=0 And 收费时间='" & Format(arrTmp(i), "yyyy-MM-dd") & "'"
                        If Not (rsDays.RecordCount > 0 And Format(arrTmp(i), "yyyy-MM-dd") <> strDate) Then
                            dbl数量 = 0
                        End If
                    ElseIf int收费方式 = 7 Then
                        '当天首次不收取的，存在=1就收取，存在=0的为首次
                        rsDays.Filter = "存在=1 And 收费时间='" & Format(arrTmp(i), "yyyy-MM-dd") & "'"
                        If rsDays.RecordCount = 0 And Format(arrTmp(i), "yyyy-MM-dd") <> strDate Then
                            dbl数量 = 0
                        End If
                    End If
                Else
                    '如果不收取，则设置为0
                    dbl数量 = 0
                End If
                rsExec!数量 = dbl数量
                rsExec.Update
            Next
        End If
    End If
    AdviceMoneyMake = blnMakeMoney
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub GetCurBillSet(rsBill As Recordset, ByVal strKey As String, strNO As String, lng费用序号 As Long, lng发送序号 As Long, bln收费单 As Boolean, ByRef lngNOSequence As Long)
'功能：获取当前费用单据的NO及序号
'参数：lng费用序号=费用记录中的序号,为-1时表示不取费用序号
'      lng发送序号=发送记录中的序号,为-1时表示不取发送序号
'说明：strKey=根据记帐单据生成规则定的唯一关键字
'1.中西成药按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
'2.一个配方中的所有草药分配一个独立单据号
'3.材料医嘱与成药分号规则相同。
'4.其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法)
'5.检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
'6.一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
    rsBill.Filter = "Key='" & strKey & "'"
    If rsBill.EOF Then
        rsBill.AddNew
        rsBill!Key = strKey
        
        '取单据号
        'rsBill!NO = gobjComlib.zldatabase.GetNextNo(IIF(bln收费单, 13, 14)),门诊记帐单也是14
        lngNOSequence = lngNOSequence + 1
        rsBill!NO = "TemporaryNO=" & IIF(bln收费单, 13, 14) & Format(lngNOSequence, "00000")
        
        rsBill!费用序号 = IIF(lng费用序号 = -1, 0, 1)
        rsBill!发送序号 = IIF(lng发送序号 = -1, 0, 1)
        rsBill.Update
    Else
        If lng费用序号 <> -1 Then
            rsBill!费用序号 = rsBill!费用序号 + 1
        End If
        If lng发送序号 <> -1 Then
            rsBill!发送序号 = rsBill!发送序号 + 1
        End If
        rsBill.Update
    End If
    strNO = rsBill!NO
    If lng费用序号 <> -1 Then lng费用序号 = rsBill!费用序号
    If lng发送序号 <> -1 Then lng发送序号 = rsBill!发送序号
End Sub

Public Function GetAuditName(ByVal strName As String) As String
'功能：从"审核医生/实习医生"中取审核医生名
    GetAuditName = Mid(strName, 1, IIF(InStr(strName, "/") > 0, InStr(strName, "/") - 1, Len(strName)))
End Function

Public Function GetPatiUnitID(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Long
'功能：根据病人获取对应的病区ID
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = "Select 当前病区ID as 病区ID From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    GetPatiUnitID = Nvl(rsTmp!病区ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Load动态费别(lng科室id As Long) As String
'功能：权限指定科室读取当前有效的动态费别(目前只用于门诊)
'返回：费别串="三八节,五一节"
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, strTmp As String
    
    On Error GoTo errH
    
    StrSQL = _
        " Select 编码,简码,名称 From 费别" & _
        " Where Nvl(属性,1)=2 And Nvl(适用科室,1)=1 And Nvl(服务对象,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(有效开始,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(有效结束,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Union ALL" & _
        " Select Distinct A.编码,A.简码,A.名称" & _
        " From 费别 A,费别适用科室 B" & _
        " Where A.名称=B.费别 And B.科室ID=[1]" & _
        " And Nvl(A.属性,1)=2 And Nvl(A.适用科室,1)=2 And Nvl(A.服务对象,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(A.有效开始,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(A.有效结束,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by 编码"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Load动态费别", lng科室id)
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    Load动态费别 = Mid(strTmp, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function ActualMoney(str费别 As String, ByVal lng收入项目ID As Long, ByVal cur应收金额 As Currency, _
    Optional ByVal lng收费细目ID As Long, Optional ByVal lng库房ID As Long, Optional ByVal dbl数量 As Double, Optional ByVal dbl加班加价率 As Double) As Currency
'功能：根据收费细目ID或收入项目ID(前者优先),应收金额,按费别设置的分段比例打折规则计算实收金额；
'       或对药品按成本加收比例规则计算实收金额
'参数：str费别=病人费别；如果是按动态费别,传入格式为"病人费别,动态费别1,动态费别2,..."
'      lng库房ID,dbl数量,对药品类项目按成本价加收打折时才需要传入
'      dbl数量=包含付数在内的售价数量
'      dbl加班加价率=小数比率,传入的应收金额已按加班加价计算时需要，用于还原及重算
'返回：按打折规则和比例计算的实收金额,如果是动态费别,则"str费别"返回最优惠费别(注意如果未打折计算,可能原样返回,也可能返回第一个)
'说明：
'按成本价加收比例打折的两种计算方法(实际是一种)：
'1.打折金额 = 成本金额 * (1 + 加收比例)
'2.打折金额 = 成本价 * (1 + 加收比例) * 零售数量
'相关的计算公式：
'      成本价 = 药品售价 * (1 - 差价率)
'      成本金额 = 售价金额 * (1 - 差价率) = 成本价 * 零售数量
'      有库存金额时:差价率 = 库存差价 / 库存金额,否则:差价率 = 指导差价率
'      对于分批药品，应每个出库批次分别计算成本价和成本金额
'      对于时价分批，"药品售价=Nvl(零售价,实际金额/实际数量)"；分批或时价药品库存不足时，不予打折计算。
    Dim rsTmp As ADODB.Recordset, StrSQL As String
    
    On Error GoTo errH
    StrSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, App.ProductName, str费别, lng收费细目ID, lng收入项目ID, cur应收金额 / (1 + dbl加班加价率), dbl数量, lng库房ID)
        
    str费别 = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl加班加价率), gstrDec)
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function TimeisLastPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个时间是否在最后一次暂停的时间内,且最后一次暂停没有启用
'说明：因为这种情况下,如果长嘱没有终止时间,某些计算会死循环
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    
    For i = UBound(arrPause) To 0 Step -1
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then
            strEnd = "3000-01-01 00:00:00"
            If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
                TimeisLastPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个时间是否在暂停的时间段中
'参数：strPause="暂停时间,开始时间;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '可能尚未启用或暂停的时候被停止
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function

Public Function GetExecDays(ByVal str分解时间 As String) As ADODB.Recordset
'功能：根据当前医嘱的执行时间串返回不重复的执行天数记录集
    Dim rsTmp As ADODB.Recordset
    Dim arrTmp As Variant, i As Long, strTmp As String
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "收费时间", adVarChar, 10
    rsTmp.Fields.Append "存在", adInteger '用于决定是否加入已存在的列表
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    arrTmp = Split(str分解时间, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = Format(arrTmp(i), "yyyy-MM-dd")
        rsTmp.Filter = "收费时间='" & strTmp & "'"
        If rsTmp.EOF Then
            rsTmp.AddNew
            rsTmp!收费时间 = strTmp
            rsTmp!存在 = 0
            rsTmp.Update
        End If
    Next
    rsTmp.Filter = ""
    Set GetExecDays = rsTmp
End Function

Private Function GetPatiDayMoneyDetail(rsMoneyDay As ADODB.Recordset, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt来源 As Byte, _
         Optional ByVal lng诊疗项目ID As Long, Optional ByVal lng收费细目ID As Long, Optional ByVal date首日不收取 As Date) As Boolean
'功能：获取指定病人当天及之后医嘱产生的费用项目明细
'参数：lng主页ID=住院病人才使用
'      byt来源:1-门诊(含住院临嘱发送到门诊)，2-住院
'      str首次时间=本次医嘱发送，首次执行的时间
'      date首日不收取=用于添加首日不收取的项目天数，但频率又不是每天一次的，实际是每天一次的，例如隔日一次，每24小时一次等
'返回：rsMoneyDay，包含"诊疗项目ID,收费项目ID,执行部门ID,执行否,收费时间"字段
'      如果是发送当天之前的医嘱，则本过程暂时没有考虑这种情况，检查当天是否已执行时会检查不到
    Dim StrSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strToDay As String, strDay As String
        
    On Error GoTo errH
    
    If lng诊疗项目ID = 0 Then
        Set rsMoneyDay = New ADODB.Recordset '用于清除Filter属性
        strToDay = Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd")
        '执行判断：
        '1.传入的是将填定到费用记录中的执行部门，因此也以费用记录中的执行部门为准判断。
        '2.除和跟踪卫材外，医嘱费用的执行科室与医嘱执行科室相同；以后如果不同了，该函数也可以适应
        '3.医嘱执行时，对应费用的执行状态也会同步标记。
        '4.首次不收的项目，如果频率是一天只收一次，则没有产生费用记录（但有医嘱发送记录）,需要读出来当成已生成的，以便其他首次不收的项目判断
        If byt来源 = 1 Then
            StrSQL = "Select A.诊疗项目ID,C.收费细目ID as 收费项目ID,C.执行部门ID,Decode(Nvl(C.执行状态,0),0,0,1) as 执行否,To_Char(C.发生时间,'yyyy-mm-dd') as 收费时间,0 as 收费方式" & _
                " From 病人医嘱记录 A,病人医嘱发送 B,门诊费用记录 C" & _
                " Where A.病人ID=[1] And Nvl(A.主页ID,0) = [2] And a.医嘱期效 = 1 And A.ID=B.医嘱ID And B.记录性质=C.记录性质 And B.NO=C.NO" & _
                " And B.医嘱ID=C.医嘱序号 And C.记录状态 IN(0,1) And C.发生时间>=[3]" & _
                " Union " & _
                " Select A.诊疗项目ID,D.收费细目id,D.执行科室ID as 执行部门ID,0 as 执行否,To_Char(B.首次时间,'yyyy-mm-dd') as 收费时间,-1 as 收费方式" & _
                " From 病人医嘱记录 A,病人医嘱发送 B,病人医嘱计价 D" & _
                " Where A.病人ID=[1] And Nvl(A.主页ID,0) = [2] And a.医嘱期效 = 1 " & _
                " And A.ID=B.医嘱ID And NVL(B.首次时间,a.开始执行时间)>=[3] And A.ID=D.医嘱ID And D.收费方式=7" & vbNewLine & _
                " And Not Exists (Select 1 From 门诊费用记录 C Where c.收费细目id=d.收费细目id  And b.记录性质 = c.记录性质 And b.No = c.No And a.Id = c.医嘱序号)" & vbNewLine & _
                " Order by 诊疗项目ID,收费项目ID"
            Set rsMoneyDay = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "读取当天及后续的医嘱", lng病人ID, lng主页ID, CDate(strToDay))
            Set rsMoneyDay = gobjComlib.zlDatabase.CopyNewRec(rsMoneyDay)
        Else
            '临嘱：病人医嘱记录.上次执行时间为空
            '长嘱，其他医嘱的相同费用，可能不同时间多次发送,Union去除了重复记录
            '首次不收的项目，如果频率是一天只收一次，则没有产生费用记录（但有医嘱发送记录）,需要读出来当成已生成的，以便其他首次不收的项目判断
            StrSQL = "Select a.诊疗项目id, c.收费细目id As 收费项目id, c.执行部门id, Decode(Nvl(c.执行状态, 0), 0, 0, 1) As 执行否," & vbNewLine & _
                "     Decode(a.医嘱期效, 0, b.首次时间, c.发生时间) As 首次时间, Decode(b.首次时间,null, 1,Trunc(b.末次时间) - Trunc(b.首次时间) + 1) As 天数,0 as 收费方式" & vbNewLine & _
                "From 病人医嘱记录 A, 病人医嘱发送 B, 住院费用记录 C" & vbNewLine & _
                "Where a.病人id = [1] And a.主页id = [2] And a.Id = b.医嘱id And b.记录性质 = c.记录性质 And b.No = c.No And b.医嘱id = c.医嘱序号 And" & vbNewLine & _
                "      c.记录状态 In (0, 1) And ((b.首次时间 > [3] Or b.末次时间 > [3]) Or a.医嘱期效 = 1 And C.发生时间 >= [3])" & vbNewLine & _
                " Union " & vbNewLine & _
                "Select a.诊疗项目id, D.收费细目id, D.执行科室ID as 执行部门id, 0 As 执行否," & vbNewLine & _
                "     b.首次时间, Decode(a.医嘱期效, 0, Trunc(b.末次时间) - Trunc(b.首次时间) + 1, 1) As 天数,-1 as 收费方式" & vbNewLine & _
                "From 病人医嘱记录 A, 病人医嘱发送 B, 病人医嘱计价 D" & vbNewLine & _
                "Where a.病人id = [1] And a.主页id = [2]" & vbNewLine & _
                "   And a.Id = b.医嘱id And ((b.首次时间 > [3] Or b.末次时间 > [3]) Or (a.医嘱期效 = 1 And b.首次时间 is null and a.开始执行时间 >= [3]))" & vbNewLine & _
                "   And A.ID=D.医嘱ID And D.收费方式=7" & vbNewLine & _
                " And Not Exists (Select 1 From 住院费用记录 C Where c.收费细目id=d.收费细目id  And b.记录性质 = c.记录性质 And b.No = c.No And a.Id = c.医嘱序号)" & vbNewLine & _
                "Order By 诊疗项目id, 收费项目id"
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "读取当天及后续的医嘱", lng病人ID, lng主页ID, CDate(strToDay))
            '根据开始时间和天数，将记录集按执行时间分成多条记录
            Set rsMoneyDay = InitPatiExecDays
                    
            For i = 1 To rsTmp.RecordCount
                For j = 1 To rsTmp!天数
                    If j = 1 Then
                        strDay = Format(rsTmp!首次时间, "yyyy-MM-dd")
                    Else
                        strDay = Format(DateAdd("d", j - 1, CDate(rsTmp!首次时间)), "yyyy-MM-dd")
                    End If
                    If strDay >= strToDay Then
                        rsMoneyDay.Filter = "诊疗项目ID=" & Val("" & rsTmp!诊疗项目ID) & " And 收费项目ID=" & Val("" & rsTmp!收费项目ID) & _
                                            " And 收费时间='" & strDay & "' And 执行否=" & Val("" & rsTmp!执行否) & " And 收费方式=" & Val("" & rsTmp!收费方式)
                        If rsMoneyDay.RecordCount = 0 Then
                            rsMoneyDay.AddNew
                            rsMoneyDay!诊疗项目ID = Val("" & rsTmp!诊疗项目ID)
                            rsMoneyDay!收费项目ID = Val("" & rsTmp!收费项目ID)
                            rsMoneyDay!执行部门ID = Val("" & rsTmp!执行部门ID)
                            rsMoneyDay!执行否 = Val("" & rsTmp!执行否)
                            rsMoneyDay!收费方式 = Val("" & rsTmp!收费方式)
                            rsMoneyDay!收费时间 = strDay
                            rsMoneyDay.Update
                        End If
                    End If
                Next
                rsTmp.MoveNext
            Next
            rsMoneyDay.Filter = ""
        End If
    Else
        '门诊发送时用于判断每天首次不收取的项目当天是否执行次数=1,如果=1且没有收费，说明当天首次已经没有收取了
        StrSQL = "Select d.执行科室id As 执行部门id" & vbNewLine & _
                "From 病人医嘱记录 A,病人医嘱发送 B, 病人医嘱计价 D" & vbNewLine & _
                "Where A.病人ID=[1] And Nvl(A.主页ID,0) = [2] And a.Id = b.医嘱id And A.id = d.医嘱id And A.诊疗项目ID = [6] And d.收费方式 = 7 And d.收费细目id = [3] And Not Exists" & vbNewLine & _
                " (Select 1" & vbNewLine & _
                "       From " & IIF(byt来源 = 1, "门诊费用记录", "住院费用记录") & " C" & vbNewLine & _
                "       Where c.收费细目id = d.收费细目id And b.记录性质 = c.记录性质 And b.No = c.No And d.医嘱id = c.医嘱序号) And" & vbNewLine & _
                "      Zl_Adviceexecount(d.医嘱id, [4], [5],1) = 1"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "读取当天及后续的医嘱", lng病人ID, lng主页ID, lng收费细目ID, CDate(Format(date首日不收取, "yyyy-MM-dd")), CDate(Format(date首日不收取, "yyyy-MM-dd 23:59:59")), lng诊疗项目ID)
        If rsTmp.RecordCount > 0 Then
            rsMoneyDay.Filter = "诊疗项目ID=" & lng诊疗项目ID & " And 收费项目ID=" & lng收费细目ID & _
                                " And 收费时间='" & Format(date首日不收取, "yyyy-MM-dd") & "' And 执行否=0" & " And 收费方式=-1"
            If rsMoneyDay.RecordCount = 0 Then
                rsMoneyDay.AddNew
                rsMoneyDay!诊疗项目ID = lng诊疗项目ID
                rsMoneyDay!收费项目ID = lng收费细目ID
                rsMoneyDay!执行部门ID = Val("" & rsTmp!执行部门ID)
                rsMoneyDay!执行否 = 0
                rsMoneyDay!收费方式 = -1
                rsMoneyDay!收费时间 = Format(date首日不收取, "yyyy-MM-dd")
                rsMoneyDay.Update
            End If
        End If
    End If
    
    GetPatiDayMoneyDetail = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function InitPatiExecDays() As ADODB.Recordset
'功能：初始化医嘱相关费用执行的记录集
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "诊疗项目ID", adBigInt
    rsTmp.Fields.Append "收费项目ID", adBigInt
    rsTmp.Fields.Append "执行部门ID", adBigInt
    rsTmp.Fields.Append "收费方式", adInteger
    rsTmp.Fields.Append "执行否", adInteger
    rsTmp.Fields.Append "收费时间", adVarChar, 10
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set InitPatiExecDays = rsTmp
End Function

Public Function GetCuvetteNumber(rsNumber As ADODB.Recordset, ByVal str管码 As String, ByVal lng医嘱ID As Long, _
    ByVal lng相关ID As Long, ByVal str类别 As String, ByVal int操作类型 As Integer, ByVal lng执行科室ID As Long, _
    ByVal int婴儿 As Integer, ByVal lng诊疗项目ID As Long, ByVal int紧急 As Integer, ByVal str标本 As String, ByVal lng采集科室ID As Long) As String
    '功能：对检验医嘱生成样本条码
    '      1.一并采集的同一检验医嘱使用相同的样本条码
    '      2.相同管码的检验使用相同的样本条码
    '      3.校本条码规则:12位的"管码+医嘱ID"
    '参数：rsNumber=动态记录集，具有"管码、相关ID、样本条码"等字段
    Dim strTmp管码 As String, strTmp条码 As String
    
    If str类别 = "C" And str管码 <> "" Then '检验项目才有管码
        rsNumber.Filter = "相关ID=" & lng相关ID
        If rsNumber.EOF Then
            rsNumber.Filter = "诊疗项目id=" & lng诊疗项目ID
            If rsNumber.EOF Then
                rsNumber.Filter = "管码='" & str管码 & "' And 执行科室ID=" & lng执行科室ID & " And 婴儿=" & int婴儿 & _
                    " And 紧急标志=" & int紧急 & " And 标本='" & str标本 & "' And 采集科室ID=" & lng采集科室ID
                If rsNumber.EOF Then
                    '生成新的条码
                    rsNumber.AddNew
                    rsNumber!管码 = str管码
                    rsNumber!相关ID = lng相关ID
'                    rsNumber!样本条码 = str管码 & Format(lng医嘱ID, Replace(Space(12 - Len(str管码)), " ", "0"))
                    rsNumber!样本条码 = gobjComlib.zlDatabase.GetNextNo(125, lng医嘱ID)
                    rsNumber!诊疗项目ID = lng诊疗项目ID
                    rsNumber!执行科室ID = lng执行科室ID
                    rsNumber!婴儿 = int婴儿
                    rsNumber!紧急标志 = int紧急
                    rsNumber!标本 = str标本
                    rsNumber!采集科室ID = lng采集科室ID
                    rsNumber.Update
                    
                    strTmp条码 = rsNumber!样本条码
                Else
                    '相同管码、执行科室、婴儿的检验使用相同的样本条码
                    strTmp管码 = Nvl(rsNumber!管码)
                    strTmp条码 = Nvl(rsNumber!样本条码)
                    
                    rsNumber.AddNew
                    rsNumber!管码 = strTmp管码
                    rsNumber!相关ID = lng相关ID
                    rsNumber!样本条码 = strTmp条码
                    rsNumber!诊疗项目ID = lng诊疗项目ID
                    rsNumber!执行科室ID = lng执行科室ID
                    rsNumber!婴儿 = int婴儿
                    rsNumber!紧急标志 = int紧急
                    rsNumber!标本 = str标本
                    rsNumber!采集科室ID = lng采集科室ID
                    rsNumber.Update
                End If
            Else
                '生成新的条码：相同检验的医嘱使用"不同的"条码
                rsNumber.AddNew
                rsNumber!管码 = str管码
                rsNumber!相关ID = lng相关ID
'                rsNumber!样本条码 = str管码 & Format(lng医嘱ID, Replace(Space(12 - Len(str管码)), " ", "0"))
                rsNumber!样本条码 = gobjComlib.zlDatabase.GetNextNo(125, lng医嘱ID)
                rsNumber!诊疗项目ID = lng诊疗项目ID
                rsNumber!执行科室ID = lng执行科室ID
                rsNumber!婴儿 = int婴儿
                rsNumber!紧急标志 = int紧急
                rsNumber!标本 = str标本
                rsNumber!采集科室ID = lng采集科室ID
                rsNumber.Update
                
                strTmp条码 = rsNumber!样本条码
            End If
        Else
            '一并采集的检验项目使用相同的条码
            strTmp管码 = Nvl(rsNumber!管码)
            strTmp条码 = Nvl(rsNumber!样本条码)
            
            rsNumber.AddNew
            rsNumber!管码 = strTmp管码
            rsNumber!相关ID = lng相关ID
            rsNumber!样本条码 = strTmp条码
            rsNumber!诊疗项目ID = lng诊疗项目ID
            rsNumber!执行科室ID = lng执行科室ID
            rsNumber!婴儿 = int婴儿
            rsNumber!紧急标志 = int紧急
            rsNumber!标本 = str标本
            rsNumber!采集科室ID = lng采集科室ID
            rsNumber.Update
        End If
        ElseIf str类别 = "E" And int操作类型 = 6 Then
        '采集方式使用与医嘱相同(最近)的条码
        If Not rsNumber.EOF Then
            If Nvl(rsNumber!相关ID, 0) = lng医嘱ID Then
                strTmp条码 = Nvl(rsNumber!样本条码)
            End If
        End If
    End If
    
    GetCuvetteNumber = strTmp条码
End Function

Public Function GetAuditRecord(lng病人ID As Long, lng主页ID As Long) As ADODB.Recordset
'功能：获取指定病人的费用审批项目
    Dim StrSQL As String
    
    On Error GoTo errH
    StrSQL = "Select 项目Id,使用限量,已用数量,使用限量-已用数量 可用数量 From 病人审批项目 Where 病人ID=[1] And 主页ID=[2]"
    Set GetAuditRecord = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlInExse", lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get实收金额(ByVal StrSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    StrSQL = Mid(StrSQL, InStr(StrSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    StrSQL = Left(StrSQL, InStr(StrSQL, strMatch) - 1)
    Get实收金额 = CCur(StrSQL)
End Function

Public Function Set实收金额(ByVal StrSQL As String, ByVal cur金额 As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(StrSQL, 1, InStr(StrSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(StrSQL, InStr(StrSQL, strMatch) + Len(strMatch))
    
    Set实收金额 = strLeft & cur金额 & strRight
End Function


Public Function Get病人诊断记录(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str类型 As String) As ADODB.Recordset
'功能：获取病人诊断记录
'参数：lng就诊ID：门诊病人传挂号ID，住院病人传主页ID
'       诊断类型-1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;
'        11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断
'       记录来源:1-病历；2-入院登记；3-首页整理(门诊医生站,诊断摘要);
    Dim StrSQL As String

    On Error GoTo errH
    StrSQL = "Select a.疾病id, a.诊断id, a.诊断描述, a.诊断次序, Nvl(b.编码, c.编码) As 编码, Nvl(b.名称, c.名称) 名称" & vbNewLine & _
             "From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C" & vbNewLine & _
             "Where a.病人id = [1] And a.主页id = [2] And NVL(A.编码序号,1) = 1  And 取消时间 Is Null And 记录来源 IN (1, 3) And Instr(',' ||[3]|| ',', ',' || 诊断类型 || ',') > 0 And a.疾病id = b.Id(+) And" & vbNewLine & _
             "      a.诊断id = c.Id(+)" & vbNewLine & _
             "Order By 记录来源, 诊断类型, 诊断次序"
    Set Get病人诊断记录 = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlPublic", lng病人ID, lng就诊ID, str类型)

    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub ReplaceTrueNO(rsSQL As ADODB.Recordset, rsUpload As ADODB.Recordset)
'功能：将临时产生的NO替换成最终保存的真实NO
    Dim strNO As String, strCur As String, strPre As String
    
    rsSQL.Filter = 0
    rsSQL.Sort = "NO"
    Do While Not rsSQL.EOF
        If Not IsNull(rsSQL!NO) Then
            strCur = Split(rsSQL!NO, "=")(1)
            If strCur <> strPre Then
                strPre = strCur
                strNO = gobjComlib.zlDatabase.GetNextNo(Val(Left(strCur, 2)))
                            
                'rsUpload中一个NO只有一条记录
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "NO='" & rsSQL!NO & "'"
                    If Not rsUpload.EOF Then
                        rsUpload!NO = strNO
                        rsUpload.Update
                    End If
                End If
            End If
            
            rsSQL!sql = Replace(rsSQL!sql, rsSQL!NO, strNO)
            'rsSQL!NO = strNO '这个不更新，避免导致Sort后顺序紊乱
            rsSQL.Update
        End If
        rsSQL.MoveNext
    Loop
End Sub

Public Function Get输液配置中心() As String
'功能：获取输液配置中心的科室IDs
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Integer
    Dim strReturn As String
    
    On Error GoTo errH

    StrSQL = "Select 部门id From 部门性质说明 Where 工作性质 = '配制中心' Order by 部门id"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Get输液配置中心")
    
    For i = 1 To rsTmp.RecordCount
        strReturn = strReturn & "," & rsTmp!部门ID
        rsTmp.MoveNext
    Next
    Get输液配置中心 = Mid(strReturn, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub GetTestLabel(ByVal strScript As String, ByVal strSelect As String, intResult As Integer)
'功能：获取皮试标注和结果
'参数：strScript=皮试结果描述串，如"阳性(+),大阳性(++);阴性(-)"
'      strSelect=所选择的皮试结果
'返回：strLabel = 皮试结果标注，如"(+)"
'      intResult=皮试结果：0-阴性，1-阳性
    Dim arr阳性 As Variant, arr阴性 As Variant
    Dim i As Integer
    
    intResult = 0
    
    arr阳性 = Split(Split(strScript, ";")(0), ",")
    arr阴性 = Split(Split(strScript, ";")(1), ",")
    
    For i = 0 To UBound(arr阳性)
        If arr阳性(i) Like "*" & strSelect & "*" Then
            intResult = 1: Exit Sub
        End If
    Next
    For i = 0 To UBound(arr阴性)
        If arr阴性(i) Like "*" & strSelect & "*" Then
            intResult = 0: Exit Sub
        End If
    Next
End Sub

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'功能：获取药品或卫材出库检查的集合
'参数：bytType:0-药品，1-卫材
    Dim rsTmp As ADODB.Recordset, StrSQL As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '避免出错
    
    StrSQL = _
        " Select Distinct A.ID,C.检查方式" & _
        " From 部门表 A,部门性质说明 B," & IIF(bytType = 0, "药品出库检查", "材料出库检查") & " C" & _
        " Where B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 " & IIF(bytType = 0, "IN('中药房','西药房','成药房')", "='发料部门'") & _
        " And C.库房ID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "GetStockCheck")
    For i = 1 To rsTmp.RecordCount
        colStock.Add Nvl(rsTmp!检查方式, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'功能：判断是否存在指定处方单据类型的入出类别
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = "Select 类别ID From 药品单据性质 Where 单据=[1]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!类别ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetOutPatiInfo(rsPati As Recordset, ByVal lng病人ID As Long, ByVal lng挂号id As Long) As Boolean
'功能：读取病人信息
    Dim StrSQL As String
    
    On Error GoTo errH
    
    '执行部门(号别科室)即病人科室
    StrSQL = "Select 病人ID,预交余额,费用余额 From 病人余额 Where 性质=1 And 类型 = 1 And 病人ID=[1]"
    StrSQL = "Select Decode(A.合同单位ID,NULL,NULL,Nvl(A.工作单位,D.名称)) as 单位,Nvl(c.姓名,A.姓名) 姓名,Nvl(c.性别,A.性别) 性别 ,Nvl(c.年龄,A.年龄) 年龄 ,A.门诊号,C.No as 挂号单," & _
        " A.费别,A.险类,A.结算模式,zl_PatiWarnScheme(A.病人ID) as 适用病人,A.担保额,Nvl(B.预交余额,0)-Nvl(B.费用余额,0) as 剩余款" & _
        " From 病人信息 A,(" & StrSQL & ") B,病人挂号记录 C,合约单位 D" & _
        " Where A.病人ID=B.病人ID(+) And A.合同单位ID=D.ID(+)" & _
        " And A.病人id = C.病人id(+) And A.门诊号 = C.门诊号(+) " & _
        " And A.病人ID=[1] And c.id(+)=[2]"
    'Set mrsPati = New ADODB.Recordset
    Set rsPati = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "GetOutPatiInfo", lng病人ID, lng挂号id)

    GetOutPatiInfo = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetMergeIDs(ByRef vsAdvice As VSFlexGrid, ByVal lngRow As Long, ByVal COL_相关ID As Long, ByVal COL_ID As Long) As String
'功能：获取指定一并给药的医嘱ID串(非一并给药返回当前医嘱ID)
'参数：lngRow=一并给药的开始药品行
    Dim lng相关ID As Long, i As Long
    Dim str医嘱ID As String
    
    With vsAdvice
        lng相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
        For i = lngRow To .Rows - 1
            If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                str医嘱ID = str医嘱ID & "," & Val(.TextMatrix(i, COL_ID))
            Else
                Exit For
            End If
        Next
    End With
    
    GetMergeIDs = Mid(str医嘱ID, 2)
End Function

Public Function GetRXKey(ByRef rsRXKey As ADODB.Recordset, ByVal strKey As String, ByVal str医嘱ID As String) As String
'功能：返回药品处方条数限制关键字,用于处方NO分配
'参数：strKey=当前处方NO的Key,不包含处方条数限制Key部份
'      str医嘱ID=当前药品的医嘱ID串，一并给药包含多个ID，"ID1,ID2,..."
'                一并给药开始行或独立药品行才传入,一并给药中间行传入空
    Dim intNextCount As Integer
    Dim strNextID As String
    
    rsRXKey.Filter = "Key='" & strKey & "'"
    If rsRXKey.EOF Then
        strNextID = gobjComlib.zlStr.Listminus(str医嘱ID, "")
        intNextCount = UBound(Split(strNextID, ",")) + 1
        
        rsRXKey.AddNew
        rsRXKey!Key = strKey
        rsRXKey!医嘱ID = strNextID
        rsRXKey!条数 = intNextCount
        rsRXKey!张数 = 1
        rsRXKey.Update
    ElseIf str医嘱ID <> "" Then
        strNextID = gobjComlib.zlStr.Listminus(str医嘱ID, rsRXKey!医嘱ID)
        intNextCount = UBound(Split(strNextID, ",")) + 1
        
        rsRXKey!医嘱ID = rsRXKey!医嘱ID & "," & strNextID
        rsRXKey!条数 = rsRXKey!条数 + intNextCount
        rsRXKey.Update
    
        If rsRXKey!条数 > gintRXCount Then
            strNextID = gobjComlib.zlStr.Listminus(str医嘱ID, "")
            intNextCount = UBound(Split(strNextID, ",")) + 1
            
            rsRXKey!张数 = rsRXKey!张数 + 1
            rsRXKey!医嘱ID = strNextID
            rsRXKey!条数 = intNextCount
            rsRXKey.Update
        End If
    ElseIf str医嘱ID = "" Then
        '一并给药中间行,保持第一行的关键字
    End If

    GetRXKey = rsRXKey!张数
End Function

Public Function GetClinicBillID(ByVal lng项目ID As Long, ByVal int场合 As Integer) As Long
'功能：获取诊疗项目对应的诊疗单据(不管附项,用于生成发送NO)
'参数：int场合=1-门诊,2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = "Select 病历文件ID From 病历单据应用 Where 诊疗项目ID=[1] And 应用场合=[2]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng项目ID, int场合)
    If Not rsTmp.EOF Then GetClinicBillID = Nvl(rsTmp!病历文件ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetStock(ByVal lng药品ID As Long, Optional ByVal lng库房ID As Long, Optional ByVal int范围 As Integer = 2, _
        Optional ByVal strDepartments As String, Optional ByVal lng总量 As Double) As Double
'功能：获取指定库房指定药品不分批库存(以门诊或住院单位)
'参数：int范围=1-门诊,2-住院(缺省),0-表示按售价
'      strDepartments可用执行科室字符串，用于批量查询库存
'      lng总量 如果lng总量不为空，则查询是否有库存大于这个总量
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, strTmp As String
    
    On Error GoTo errH
    '获取药品库存(不分批或分批药品),药房不分批药品不管效期
    If int范围 = 0 Or int范围 = 3 Then
        StrSQL = _
            " Select Nvl(Sum(A.可用数量),0) as 库存" & _
            " From 药品库存 A" & _
            " Where A.性质=1" & _
            " And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
            " And A.药品ID=[1] And Instr([2],',' || a.库房id || ',')>0 Group By A.库房ID"
    Else
        strTmp = IIF(int范围 = 1, "门诊", "住院")
        StrSQL = _
            " Select Nvl(Sum(A.可用数量),0)/Nvl(B." & strTmp & "包装,1) as 库存" & _
            " From 药品库存 A,药品规格 B" & _
            " Where A.药品ID=B.药品ID(+) And A.性质=1" & _
            " And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
            " And A.药品ID=[1] And Instr([2],',' || a.库房id || ',')>0" & _
            " Group by Nvl(B." & strTmp & "包装,1),A.库房ID"
    End If
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng药品ID, IIF(strDepartments = "", "," & lng库房ID & ",", "," & strDepartments & ","))
    
    Do While Not rsTmp.EOF
    
        If strDepartments = "" Then
            GetStock = Format(rsTmp!库存, "0.00000")
            Exit Function
        Else
            If Val(rsTmp!库存) & "" > lng总量 Then
                GetStock = Format(rsTmp!库存, "0.00000")
                Exit Function
            End If
        End If
        rsTmp.MoveNext
    
    Loop
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Trim分解时间(ByVal lng次数 As Long, ByVal str分解时间 As String) As String
'功能：将医嘱执行的分解时间按次数进行截断
    Dim arrTime() As String, strTmp As String, i As Long
    
    arrTime = Split(str分解时间, ",")
    For i = 0 To lng次数 - 1
        strTmp = strTmp & "," & arrTime(i)
    Next
    Trim分解时间 = Mid(strTmp, 2)
End Function

Public Function GetPriceGradeSQL(ByVal str药品价格等级 As String, ByVal str卫材价格等级 As String, ByVal str普通项目价格等级 As String, ByVal strTableTmpA As String, ByVal strTableTmpB As String, _
           ByVal strParNum药品 As String, ByVal strParNum卫材 As String, ByVal strParNum普通项目 As String) As String
'功能：病人价格等级获得批量获取价格的SQL
'参数：str药品价格等级  '病人的药品价格等级
'      str卫材价格等级  '病人的卫材价格等级
'      str普通项目价格等级  '病人的普通项目价格等级
'     strTableTmpA   收费项目目录 表的as 标志,strTableTmpB  收费价目表 的As标志；
'     strParNum药品  药品价格等级SQL参数序号,strParNum卫材  卫材价格等级SQL参数序号,strParNum普通项目  普通项目价格等级SQL参数序号
    Dim StrSQL As String
    
    If str药品价格等级 = "" And str卫材价格等级 = "" And str普通项目价格等级 = "" Then
        StrSQL = " And " & strTableTmpB & ".价格等级 is Null "
    Else
        StrSQL = " And" & vbNewLine & _
                "      ((Instr(';5;6;7;', ';' || " & strTableTmpA & ".类别 || ';') > 0 And " & strTableTmpB & ".价格等级 = [" & strParNum药品 & "]) Or" & vbNewLine & _
                "      (Instr(';4;', ';' || " & strTableTmpA & ".类别 || ';') > 0 And " & strTableTmpB & ".价格等级 = [" & strParNum卫材 & "]) Or" & vbNewLine & _
                "      (Instr(';4;5;6;7;', ';' || " & strTableTmpA & ".类别 || ';') = 0 And " & strTableTmpB & ".价格等级 = [" & strParNum普通项目 & "]) Or" & vbNewLine & _
                "      (" & strTableTmpB & ".价格等级 Is Null And Not Exists" & vbNewLine & _
                "       (Select 1" & vbNewLine & _
                "         From 收费价目" & vbNewLine & _
                "         Where " & strTableTmpA & ".Id = 收费细目id  And" & vbNewLine & _
                "               ((Instr(';5;6;7;', ';' || " & strTableTmpA & ".类别 || ';') > 0 And 价格等级 = [" & strParNum药品 & "]) Or" & vbNewLine & _
                "               (Instr(';4;', ';' || " & strTableTmpA & ".类别 || ';') > 0 And 价格等级 = [" & strParNum卫材 & "]) Or" & vbNewLine & _
                "               (Instr(';4;5;6;7;', ';' || " & strTableTmpA & ".类别 || ';') = 0 And 价格等级 = [" & strParNum普通项目 & "]))))) "

    End If
    
    GetPriceGradeSQL = StrSQL
End Function