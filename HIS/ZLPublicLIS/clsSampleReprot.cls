VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSampleReprot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InitSampleReprot(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngModule As Long, strErr As String) As Boolean
    '功能：接口部件初始化
    '参数
    'cnMain          主连接
    'lngSys          系统号
    'lngModule       系统模块
    'strErr          返回错误日志
    '返回值           true 成功，false 失败
    On Error GoTo errH
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    glngSys = lngSys
    glngModule = lngModule
    If GetUserInfo = False Then
        strErr = "未找到对应的用户"
        Exit Function
    End If
    InitSampleReprot = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function binInit() As Boolean
    If gcnOracle = "" Then
        binInit = False
        MsgBox "请初始化LIS公共接口!", vbInformation, "LIS公共接口初始化"
        Exit Function
    Else
        binInit = True
    End If
End Function


Public Function GetLaboratoryReportList(ByVal lngPatientID As Long, ByVal lngMainID As Long, Optional lngApplyID As Long) As String
    '功能：提取病人的检验报告列表
    '参数
    'lngPatientID           病人id
    'lngMainID              主页id
    'lngApplyID             申请id
    '返回                   XML格式的字串
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objXML As New zl9ComLib.clsXML
    Dim strXMLOldLIS As String
    Dim strXMLNewLIS As String
    Dim i As Long
    On Error GoTo errH
    If binInit = False Then Exit Function
     InitObjLis
     Call gobjLISInsideComm.GetLabNewReportList(lngPatientID, lngMainID, strXMLNewLIS, lngApplyID)
     If strXMLNewLIS = "" Then
        strSQL = "Select id 检验报告id,医嘱id 申请id,紧急 紧急标志,检验项目,标本序号,微生物标本 是否微生物,报告结果 检验次数,检验人,审核人,审核时间,申请时间 " & vbNewLine & _
                 "  from 检验标本记录 where 病人id = [1] and 主页id =[2] and 医嘱id is not null"
                 
        If lngApplyID > 0 Then
            strSQL = strSQL & " and 医嘱id =[3]"
        End If
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "病人检验报告列表", lngPatientID, lngMainID, lngApplyID)
        If rsTmp.RecordCount > 0 Then
            With objXML
                .ClearXmlText
                .AppendNode "检验报告列表" ', True '父节点[检验报告列表]
                For i = 1 To rsTmp.RecordCount
                    .AppendData "检验报告id", rsTmp!检验报告id '<检验报告id>类型：
                    .AppendData "申请id", rsTmp!申请id '<申请id>类型：
                    .AppendData "紧急标志", rsTmp!紧急标志 & "" '<紧急标志>类型：
                    .AppendData "检验项目", rsTmp!检验项目 & ""  '<检验项目>类型：
                    .AppendData "标本序号", rsTmp!标本序号 '<标本序号>类型：
                    .AppendData "是否微生物", rsTmp!是否微生物 & ""  '<微生物标本>类型：
                    .AppendData "检验次数", Val(rsTmp!检验次数 & "") '< 报告结果 > 类型:
                    .AppendData "检验人", rsTmp!检验人 & "" '<检验人>类型：
                    .AppendData "审核人", rsTmp!审核人 & ""  '<审核人>类型：
                    .AppendData "审核时间", rsTmp!审核时间 & "" '<审核时间>类型：
                    .AppendData "申请时间", rsTmp!申请时间 & ""  '<申请时间>类型：
                    rsTmp.MoveNext
                Next
                .AppendNode "检验报告列表", True
                If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
            End With
        End If
        GetLaboratoryReportList = strXMLOldLIS
    Else
        GetLaboratoryReportList = strXMLNewLIS
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function GetLaboratoryReportResultList(ByVal lngRepottID As Long) As String
    '功能：提取病人的检验报告结果
    '参数
    'lngRepottID            报告id
    '返回                   XML格式的字串
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsNewTmp As ADODB.Recordset
    Dim objXML As New zl9ComLib.clsXML
    Dim strXMLOldLIS As String
    Dim strXMLNewLIS As String
    Dim strBH As String
    Dim i As Long
    On Error GoTo errH
    If binInit = False Then Exit Function
    InitObjLis
    Call gobjLISInsideComm.GetLabNewReportResultList(lngRepottID, strXMLNewLIS)
    If strXMLNewLIS = "" Then
        '新版没有数据，查看老版是否有记录
        strSQL = "select  id,微生物标本 from 检验标本记录 where id = [1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "病人检验结果列表", lngRepottID)
        If rsTmp.RecordCount > 0 Then
            If Val(rsTmp("微生物标本") & "") = 1 Then
                strSQL = "Select distinct a.细菌id, b.中文名 细菌名, a.培养描述 描述, a.耐药机制, e.中文名 抗生素, c.结果 抗生素结果, c.结果类型 耐药性, c.药敏方法, e.用法用量1, e.用法用量2, e.血药浓度1," & vbNewLine & _
                         "          e.血药浓度2 , e.尿药浓度1, e.尿药浓度2" & vbNewLine & _
                         "   From 检验普通结果 A, 检验细菌 B, 检验药敏结果 C, 检验用抗生素 E" & vbNewLine & _
                         "   Where a.细菌id = b.Id And c.细菌结果id = a.Id And c.抗生素id = e.Id and a.检验标本id =[1] order by b.中文名"
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "病人检验结果列表", lngRepottID)
                If rsTmp.RecordCount > 0 Then
                    With objXML
                        .ClearXmlText
                        .AppendNode "微生物项目" ', True '父节点[普通项目]
                        For i = 1 To rsTmp.RecordCount
                            If strBH <> rsTmp!细菌名 & "" Then
                                If strBH <> "" Then
                                    .AppendNode "抗生素结果列表", True
                                End If
                                strBH = rsTmp!细菌名 & ""
                                .AppendData "细菌id", rsTmp!细菌id & "" '<细菌id>类型：
                                .AppendData "细菌名", rsTmp!细菌名 & "" '<细菌名>类型：
                                .AppendData "描述", rsTmp!描述 & "" '<描述>类型：
                                .AppendData "耐药机制", rsTmp!耐药机制 & ""  '<耐药机制>类型：
                                .AppendNode "抗生素结果列表" ', True '父节点[指标内容]
                            End If
                        
                            .AppendData "抗生素", rsTmp!抗生素 & ""  '<抗生素>类型：
                            .AppendData "抗生素结果", rsTmp!抗生素结果 & ""  '<抗生素结果>类型：
                            .AppendData "耐药性", rsTmp!耐药性 & "" '<耐药性>类型：
                            .AppendData "药敏方法", rsTmp!药敏方法 & ""  '<药敏方法>类型：
                            .AppendData "用法用量1", rsTmp!用法用量1 & "" '<用法用量1>类型：
                            .AppendData "用法用量2", rsTmp!用法用量2 & ""  '<用法用量2>类型：
                            .AppendData "血药浓度1", rsTmp!血药浓度1 & "" '< 血药浓度1 > 类型:
                            .AppendData "血药浓度2", rsTmp!血药浓度2 & "" '<血药浓度2>类型：
                            .AppendData "尿药浓度1", rsTmp!尿药浓度1 & ""  '<尿药浓度1>类型：
                            .AppendData "尿药浓度2", rsTmp!尿药浓度2 & ""  '<尿药浓度2>类型：
                            rsTmp.MoveNext
                        Next
                        .AppendNode "抗生素结果列表", True
                        .AppendNode "微生物项目", True
                        If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
                    End With
                End If
            Else
                strSQL = "Select a.检验项目id 指标id, b.英文名 指标代码, b.英文名 指标英文名, b.中文名 指标中文名, a.检验结果, a.结果标志, a.结果参考, a.排列序号, c.隐私项目" & vbNewLine & _
                         "   From 检验普通结果 A, 诊治所见项目 B, 检验项目 C" & vbNewLine & _
                         "   Where a.检验项目id = b.Id And b.Id = c.诊治项目id and  a.检验标本id =[1] "
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "病人检验结果列表", lngRepottID)
                If rsTmp.RecordCount > 0 Then
                    With objXML
                        .ClearXmlText
                        .AppendNode "普通项目" ', True '父节点[普通项目]
                        .AppendNode "指标内容" ', True '父节点[指标内容]
                        For i = 1 To rsTmp.RecordCount
                            .AppendData "指标id", rsTmp!指标id & "" '<指标id>类型：
                            .AppendData "指标代码", rsTmp!指标代码 & "" '<指标代码>类型：
                            .AppendData "指标英文名", rsTmp!指标英文名 & "" '<指标英文名>类型：
                            .AppendData "指标中文名", rsTmp!指标中文名 & ""  '<指标中文名>类型：
                            .AppendData "检验结果", rsTmp!检验结果 & "" '<检验结果>类型：
                            .AppendData "结果标志", rsTmp!结果标志 & ""  '<结果标志>类型：
                            .AppendData "结果参考", rsTmp!结果参考 & "" '< 结果参考 > 类型:
                            .AppendData "排列序号", rsTmp!排列序号 & "" '<排列序号>类型：
                            .AppendData "隐私项目", rsTmp!隐私项目 & ""  '<隐私项目>类型：
                            rsTmp.MoveNext
                        Next
                        .AppendNode "指标内容", True
                        .AppendNode "普通项目", True
                        If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
                    End With
                End If
            End If
        End If
        If strXMLOldLIS <> "" Then GetLaboratoryReportResultList = strXMLOldLIS
    Else
        GetLaboratoryReportResultList = strXMLNewLIS
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function GetBloodBankItem(ByVal lngApplyID As Long) As String
    '功能：提取病人的检验报告结果
    '参数
    'lngApplyID            相关id ,医嘱id
    '返回                   XML格式的字串

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsNewTmp As ADODB.Recordset
    Dim objXML As New zl9ComLib.clsXML
    Dim strXMLOldLIS As String
    Dim strXMLNewLIS As String
    Dim strBH As String
    Dim i As Long
    On Error GoTo errH
    If binInit = False Then Exit Function
    InitObjLis
    Call gobjLISInsideComm.GetNewBloodBankItem(lngApplyID, strXMLNewLIS)
    If strXMLNewLIS = "" Then
        '查看老版是否有记录
        strSQL = "Select a.检验项目id 指标id, b.英文名 指标代码, b.英文名 指标英文名, b.中文名 指标中文名, a.检验结果, a.结果标志, a.结果参考" & vbNewLine & _
                 "   From 检验普通结果 A, 诊治所见项目 B,检验标本记录 c" & vbNewLine & _
                 "   Where a.检验项目id = b.Id  and a.检验标本id =  c.id and   c.医嘱id =[1] "
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "病人检验结果列表", lngApplyID)
        If rsTmp.RecordCount > 0 Then
            With objXML
                .ClearXmlText
                .AppendNode "普通项目" ', True '父节点[普通项目]
                .AppendNode "指标内容" ', True '父节点[指标内容]
                For i = 1 To rsTmp.RecordCount
                    .AppendData "指标id", rsTmp!指标id & "" '<指标id>类型：
                    .AppendData "指标代码", rsTmp!指标代码 & "" '<指标代码>类型：
                    .AppendData "指标英文名", rsTmp!指标英文名 & "" '<指标英文名>类型：
                    .AppendData "指标中文名", rsTmp!指标中文名 & ""  '<指标中文名>类型：
                    .AppendData "检验结果", rsTmp!检验结果 & "" '<检验结果>类型：
                    .AppendData "结果标志", rsTmp!结果标志 & ""  '<结果标志>类型：
                    .AppendData "结果参考", rsTmp!结果参考 & "" '<结果参考> 类型:
                    rsTmp.MoveNext
                Next
                .AppendNode "指标内容", True
                .AppendNode "普通项目", True
                If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
            End With
        End If
        If strXMLOldLIS <> "" Then GetBloodBankItem = strXMLOldLIS
    Else
        GetBloodBankItem = strXMLNewLIS
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function GetTransfusionApplyFor(strItemCodeing As String, lngPatientID As Long, intPatientType As Integer, lngHomePageID As Long, Optional strRegistrationBill As String, _
                                        Optional intBaby As Integer, Optional intType As Integer, Optional ByVal intDay As Integer) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                  根据传入医嘱ID返回结果
    '参数
    '                       strItemCodeing 诊疗项目编码（可传入多个，使用逗号分隔）
    '                       lngPatientID 病人ID
    '                       intPatientType 病人来源 1-门诊，2-住院
    '                       lngHomePageID 主页ID （病人来源=2时查询)
    '                       lngRegistrationBill 挂号单NO（病人来源<>2时查询本次就诊）
    '                       intBaby           是否婴儿
    '                       intType           那种方式，1=再此查7天内的。0 = 不查询 2=指定查询天数（intDay参数）其他= 暂定
    '                       intDay            当intType=2时，此参数才有效，指定要查询多少天的数据
    '
    '标本组成格式
    '
    '                   指标1<split1>诊疗编码1<split1>单位1<split1>隐私项目1<split1>指标代码1<split1>中文名1<split1>英文名1<split1>取值序列1<split1>
    '                       检验结果1<split2>结果标志1<split2>结果参考1<split2>排列序号1<split2>标本类型1<split3>
    '                   指标2<split1>诊疗编码2<split1>隐私项目2<split1>指标代码2<split1>中文名2<split1>英文名2<split1>取值序列2<split1>
    '                       检验结果2<split2>结果标志2<split2>结果参考2<split2>排列序号2<split2>标本类型2<split3>
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim rsTmpRuest As New ADODB.Recordset
    Dim lngID As Long
    Dim lngSampleId As Long
    Dim lngSampleGroup As Long
    Dim lngMicroID As Long
    Dim strSampleOne As String
    Dim strSampleTwo As String
    Dim strErr As String
    Dim varItemCodeing As Variant
    Dim strItemcodeOne As String
    Dim i As Integer
    Dim strStartTime As String
    Dim strEndTime As String
    Dim strTmp As String
    
    strEndTime = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    If intType = 1 Then
        strStartTime = Format(zlDatabase.Currentdate - 7, "yyyy-mm-dd 00:00:00")
    ElseIf intType = 2 Then
        strStartTime = Format(zlDatabase.Currentdate - intDay, "yyyy-mm-dd 00:00:00")
    End If
    
    '分隔的常量
    Const conSplit1 As String = "<split1>"                        '用于分隔项目信息,使用“<split1>”分隔
    Const conSplit2 As String = "<split2>"                        '用于分隔多个项目,使用“<split2>”分隔
    Const conSplit3 As String = "<split3>"                        '用于分隔项目信息,使用“<split1>”分隔
    Const conSplit4 As String = "<split4>"                        '用于分隔多个项目,使用“<split2>”分隔
        
    '只读普通项目
    On Error GoTo errH
    If binInit = False Then Exit Function
    strTmp = ""
    InitObjLis
    If Not gobjLISInsideComm Is Nothing Then
        GetTransfusionApplyFor = gobjLISInsideComm.GetNewTransFusionApplyFor(strItemCodeing, lngPatientID, intPatientType, lngHomePageID, strRegistrationBill, intBaby, intType, intDay)
        strTmp = Replace(GetTransfusionApplyFor, "<split1>", "")
        strTmp = Replace(strTmp, "<split2>", "")
        strTmp = Replace(strTmp, "<split3>", "")
        strTmp = Replace(strTmp, "<split4>", "")
        strTmp = Trim(strTmp)
    End If
    If strTmp = "" Then
        varItemCodeing = Split(strItemCodeing, ",")
        For i = LBound(varItemCodeing) To UBound(varItemCodeing)
            strItemcodeOne = varItemCodeing(i)
            strSQL = "  Select d.id 指标id, d.中文名 || '(' || d.英文名 || ')' 指标, b.单位, b.隐私项目, d.编码 指标代码, d.中文名, d.英文名, b.取值序列" & vbNewLine & _
                     "   From 检验报告项目 A, 检验项目 B, 诊疗项目目录 C, 诊治所见项目 D" & vbNewLine & _
                     "   Where b.诊治项目id = d.Id And a.诊疗项目id = c.Id And  a.报告项目id = d.Id And c.编码 = [1]"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "读取结果", strItemcodeOne)
            Do Until rsTmp.EOF
                strSampleOne = strSampleOne & conSplit3 & rsTmp("指标") & conSplit1 & strItemcodeOne & conSplit1 & rsTmp("单位") & _
                                conSplit1 & rsTmp("隐私项目") & conSplit1 & rsTmp("指标代码") & _
                                conSplit1 & rsTmp("中文名") & conSplit1 & rsTmp("英文名") & conSplit1 & rsTmp("取值序列") & conSplit1
                strSQL = " Select *" & vbNewLine & _
                         "   From (Select  审核时间, c.检验结果, Decode(c.结果标志, 1, '', 2, '↓', 3, '↑', 4, '异常', 5, '↓↓', 6, '↑↑', '') 结果标志, c.结果参考," & vbNewLine & _
                         "                 c.排列序号 , a.标本类型" & vbNewLine & _
                         "          From 检验标本记录 A, 检验项目 B, 检验普通结果 C, 诊治所见项目 D" & vbNewLine & _
                         "          Where a.Id = c.检验标本id And b.诊治项目id = d.Id And c.检验项目id = d.Id " & vbNewLine & _
                         "          [条件]  and a.审核时间 is not null and a.病人来源=[4] and d.id= [5]" & vbNewLine & _
                         "          Order By 审核时间 desc ) E" & vbNewLine & _
                         "   Where Rownum = 1"
                If intPatientType = 2 Then
                    If intBaby <> 0 Then
                        strSQL = Replace(strSQL, "[条件]", " and A.病人ID = [1] and a.主页id=[2]  and a.婴儿=[6] ")
                    Else
                        strSQL = Replace(strSQL, "[条件]", " and A.病人ID = [1] and a.主页id=[2]  and nvl(a.婴儿,0)= 0 ")
                    End If
                    Set rsTmpRuest = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "读取结果", lngPatientID, lngHomePageID, strRegistrationBill, intPatientType, Val(rsTmp("指标id")), intBaby)
                    If rsTmpRuest.RecordCount > 0 Then
                        strSampleOne = strSampleOne & rsTmpRuest("检验结果") & conSplit2 & rsTmpRuest("结果标志") & conSplit2 & rsTmpRuest("结果参考") & conSplit2 & rsTmpRuest("排列序号") & conSplit2 & rsTmpRuest("标本类型") & conSplit2 & rsTmpRuest("审核时间")
                    Else
                        If intType = 1 Or intType = 2 Then
                            strSQL = " Select *" & vbNewLine & _
                                    "   From (Select  审核时间, c.检验结果, Decode(c.结果标志, 1, '', 2, '↓', 3, '↑', 4, '异常', 5, '↓↓', 6, '↑↑', '') 结果标志, c.结果参考," & vbNewLine & _
                                    "                 c.排列序号 , a.标本类型" & vbNewLine & _
                                    "          From 检验标本记录 A, 检验项目 B, 检验普通结果 C, 诊治所见项目 D" & vbNewLine & _
                                    "          Where a.Id = c.检验标本id And b.诊治项目id = d.Id And c.检验项目id = d.Id " & vbNewLine & _
                                    "          [条件]  and a.审核时间 is not null  and d.id= [4]" & vbNewLine & _
                                    "          Order By 审核时间 desc ) E" & vbNewLine & _
                                    "   Where Rownum = 1"
                            If intBaby <> 0 Then
                                strSQL = Replace(strSQL, "[条件]", " and A.病人ID = [1] and  a.审核时间 between [2] and [3]  and a.婴儿=[5]  ")
                            Else
                                strSQL = Replace(strSQL, "[条件]", " and A.病人ID = [1] and  a.审核时间 between [2] and [3]  and nvl(a.婴儿,0)= 0  ")
                            End If
                            Set rsTmpRuest = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "读取结果", lngPatientID, CDate(strStartTime), CDate(strEndTime), Val(rsTmp("指标id")), intBaby)
                            If rsTmpRuest.RecordCount > 0 Then
                                strSampleOne = strSampleOne & rsTmpRuest("检验结果") & conSplit2 & rsTmpRuest("结果标志") & conSplit2 & rsTmpRuest("结果参考") & conSplit2 & rsTmpRuest("排列序号") & conSplit2 & rsTmpRuest("标本类型") & conSplit2 & rsTmpRuest("审核时间")
                            End If
                        End If
                    End If
                Else
                    strSQL = Replace(strSQL, "[条件]", " and A.病人ID = [1] and  a.挂号单=[3] ")
                    Set rsTmpRuest = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "读取结果", lngPatientID, lngHomePageID, strRegistrationBill, intPatientType, Val(rsTmp("指标id")))
                    If rsTmpRuest.RecordCount > 0 Then
                        strSampleOne = strSampleOne & rsTmpRuest("检验结果") & conSplit2 & rsTmpRuest("结果标志") & conSplit2 & rsTmpRuest("结果参考") & conSplit2 & rsTmpRuest("排列序号") & conSplit2 & rsTmpRuest("标本类型") & conSplit2 & rsTmpRuest("审核时间")
                    End If
                End If

                rsTmp.MoveNext
            Loop
        Next
        If strSampleOne <> "" Then
            strSampleOne = Mid(strSampleOne, Len(conSplit3) + 1)
        End If
        GetTransfusionApplyFor = strSampleOne
    End If
    Exit Function
errH:
    strErr = "读取结果出错：" & Err.Number & " " & Err.Description
End Function

Public Function ReviseDate(ByVal strDate As String) As String
'功能：将时间转化为统一的24小时制时间
    ReviseDate = Format(strDate, "yyyy-mm-dd hh:mm:ss")
End Function

Private Sub Class_Initialize()
    If gobjComLib Is Nothing Then Set gobjComLib = CreateObject("zl9Comlib.clsComlib")
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/10
'功    能:初始化通讯参数
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function Init() As Boolean
    Init = funInit
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/24
'功    能:使用winsock向其他部件发送消息
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub SendMessage(ByVal strSend As String)
    If gintStart = 0 Then Exit Sub  '如果未启用通讯，则直接退出
    Call frmComm.funSendMessage(strSend)
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/24
'功    能:向外部提供获取参数的接口
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function GetIniValue(ByVal strKey As String, Optional ByVal strDefault As String) As String
    GetIniValue = funGetIniValue(strKey, strDefault)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/24
'功    能:外部打开参数设置界面
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub ShowParaSet(objFrm As Object)
    frmParaSet.Show 1, objFrm
End Sub

