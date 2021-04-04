Attribute VB_Name = "mdlPass"
Option Explicit

Public Function OutAdviceCheckWarn_MK( _
    ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional ByRef blnNoSave As Boolean, Optional ByRef rsOut As ADODB.Recordset) As Long
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'功能：调用Pass系统中对医嘱进行合理用药审查等相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        1/33-保存自动审查(住院/门诊),2/34-提交自动审查(住院/门诊),3-手工调用审查
'        6-单药警告,12-用药研究,22-病生状态/过敏史管理(编辑)
'      lngRow=当前药品医嘱的行号，lngCmd=0,6时需要
'   lngRow=当前行
'出参
'   blnNoSave=用于标记是否保存（用于界面保存按钮可用性控制）
'   rsOut=禁忌药品说明
'返回：本次审核返回的最高级别警示值,为-1,-2,-3表示没有进行审查
'      检测PASS菜单时，返回>=0表示可以弹出菜单
' rsOut=医嘱相关内容
'说明：用药审查：涉及当天下的临嘱(包括已执行)，和未停止的长嘱
'      用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, rsPatiInfo As New ADODB.Recordset
    Dim rs中药 As ADODB.Recordset
    Dim str药品 As String, str用法 As String, str频率 As String, str用法ID As String, str间隔单位 As String
    Dim str诊断编码 As String, str诊断描述 As String, strTmp As String, strPre用法 As String
    Dim str医嘱ID As String, str序号 As String, str单量 As String, str单量单位 As String
    Dim str急诊标识 As String, str门诊总量 As String, str门诊单位 As String, str用药目的 As String, str嘱托 As String
    Dim str相关ID As String, str中药组IDs As String
    Dim lngMaxWarn As Long, strOld As String
    Dim strSQL As String, blnDo As Boolean
    Dim lngCount As Long, curDate As Date
    Dim lngTmp As Long
    Dim arrLevel(0 To 4) As Long
    Dim i As Long, k As Long, j As Long
    Dim lng中药组ID As Long, lngLight As Long
    
    Dim strType As String
    Dim str身高 As String, str体重 As String
    Dim strCurrentDate As String
    Dim arrLight(0 To 4) As String
    Dim str开嘱科室 As String, str开嘱医生 As String
    Dim int频率次数 As Integer, int频率间隔 As Integer
    Dim objDiag As clsDiagItem
    Dim rs规格 As ADODB.Recordset
    Dim str药品ID As String, str科室ID As String
    Dim str中药配方 As String
    
    Dim arrSQL As Variant
    
    lngMaxWarn = -1
    OutAdviceCheckWarn_MK = lngMaxWarn

    On Error GoTo errH
    Screen.MousePointer = 11

    '美康3.0
    '检验PASS可用状态
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '114036同一个病人多次审查时病人信息每次都要传入
    '-------------------------------------------------------------
    strSQL = "Select B.ID as 就诊ID,B.姓名,B.性别,A.出生日期," & _
             " C.编码 as 科室码,C.名称 as 科室名,E.编号 as 医生码,E.姓名 as 医生名" & _
             " From 病人信息 A,病人挂号记录 B,部门表 C,人员表 E" & _
             " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID" & _
             " And B.执行人=E.姓名(+) And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng病人ID, gobjPati.str挂号单)
    If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
    '附加信息
    strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                    "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                    "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
    Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng病人ID, gobjPati.lng挂号ID)
    rsPatiInfo.Filter = "项目名称='身高'"
    If rsPatiInfo.RecordCount <> 0 Then str身高 = NVL(rsPatiInfo!记录内容)
    rsPatiInfo.Filter = "项目名称='体重'"
    If rsPatiInfo.RecordCount <> 0 Then str体重 = NVL(rsPatiInfo!记录内容)

    Call PassSetPatientInfo(gobjPati.lng病人ID, rsTmp!就诊Id, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), str体重, str身高, _
                            rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), "")

    '传人病人过敏史
    '-------------------------------------------------------
    Set rsTmp = Get病人过敏记录(gobjPati.lng病人ID, 0)

    For i = 1 To rsTmp.RecordCount
        Call PassSetAllergenInfo(i, rsTmp!药物ID & "", rsTmp!药物名 & "", "DrugName", "")
        rsTmp.MoveNext
    Next

    '传人病生状态
    '------------------------------------------------------------------
    
    '* 诊断信息
    strCurrentDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If glngModel = PM_门诊编辑 Then
        If Not gobjDiags Is Nothing Then
            With gobjDiags
                For i = 1 To .Count
                    If .Item(i).str诊断描述 <> "" Then
                        str诊断编码 = IIf(.Item(i).str疾病编码 <> "", .Item(i).str疾病编码, .Item(i).str诊断编码)
                        str诊断描述 = .Item(i).str诊断描述
                        Call PassSetMedCond(i & "", str诊断编码, str诊断描述, "User", strCurrentDate, strCurrentDate)
                    End If
                Next
            End With
        End If
    Else
        Set rsTmp = Get病人诊断记录(gobjPati.lng病人ID, gobjPati.lng挂号ID, "1,11")
        For i = 1 To rsTmp.RecordCount
            Call PassSetMedCond(i & "", rsTmp!编码 & "", rsTmp!名称 & "", "User", strCurrentDate, strCurrentDate)
            rsTmp.MoveNext
        Next
    End If
    
    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With gobjAdvice
            If IIf(glngModel = PM_门诊编辑, .RowData(lngRow) <> 0, True) And InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)) <> 0 Then
                '取药品名称
                If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) > 0 Then
                    str药品 = .TextMatrix(lngRow, gobjCOL.intCOL药品名称)
                Else
                    str药品 = .TextMatrix(lngRow, gobjCOL.intCOL医嘱内容)  '中药名称
                End If

                '取药品给药途径(当前可见行不会是中草药) ,单量单位
                str用法 = ""
                If glngModel = PM_门诊编辑 Then
                    k = .FindRow(CLng(.TextMatrix(lngRow, gobjCOL.intCOL相关ID)), lngRow + 1)
                    If k <> -1 Then str用法 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                    strTmp = .TextMatrix(lngRow, gobjCOL.intCOL单量单位)
                Else
                    str用法 = .TextMatrix(lngRow, gobjCOL.intCOL用法)
                    If InStr(str用法, ",") > 0 Then str用法 = Left(str用法, InStr(str用法, ",") - 1)
                    strTmp = .TextMatrix(lngRow, gobjCOL.intCOL单量)
                    If Mid(strTmp, 1, 2) = "0." Then '单量有小数点的特殊处理
                        strTmp = Replace(strTmp, Format(Val(strTmp) & "", "0.####"), "") '门诊清单单量（“单量” & “单量单位”）
                    Else
                        strTmp = Replace(strTmp, Val(strTmp) & "", "")    '
                    End If
                End If
                '传入查询药品信息
                Call PassSetQueryDrug(.TextMatrix(lngRow, gobjCOL.intCOL收费细目ID), str药品, strTmp, str用法)

                '设置菜单可用状态
                
                OutAdviceCheckWarn_MK = 1    '表示可以弹出菜单
            ElseIf glngModel = PM_门诊医嘱清单 And .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(lngRow, gobjCOL.intCol操作类型) = "4" Then
                OutAdviceCheckWarn_MK = 1    '表示可以弹出菜单
            End If
        End With
        Screen.MousePointer = 0: Exit Function
    End If

    '过敏史/病生状态编辑
    '-------------------------------------------------------------
    If lngCmd = 22 Then
        'lngCmd=21-只读,22-非强制编辑,23-强制编辑
        If PassDoCommand(lngCmd) = 2 Then
            '如果返回值为2表示"过敏史/病生状态编辑"管理发生变化，需要重新自动审查
            lngCmd = 34    '转为自动调用审查,继续执行
        Else
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    '启用了禁忌药品说明参数  且场合为门诊编辑审查功能
    If (lngCmd = 33 Or lngCmd = 34 Or lngCmd = 3) And glngModel = PM_门诊编辑 And gbytReason = 1 Then
        Set rsOut = InitAdviceRS(FUN_输出内容)
    End If
    '传入病人医嘱信息
    '-------------------------------------------------------------
    With gobjAdvice
        If lngCmd = 6 Then
            If glngModel = PM_门诊编辑 Then
                strTmp = .RowData(lngRow)
            Else
                strTmp = .TextMatrix(lngRow, gobjCOL.intCOLID)
            End If
            Call PassSetWarnDrug(strTmp)   '单药警告(已警告的医嘱唯一码)
        Else
            '用药审核或用药研究
            lngCount = 0
            curDate = zlDatabase.Currentdate
            str药品 = "": str用法 = "": str频率 = ""
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_门诊编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                    blnDo = blnDo And (lngCmd = 12 Or Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                Else
                    blnDo = (InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                    Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4"))
                    blnDo = blnDo And (lngCmd = 12 Or Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                End If
                If blnDo Then
                    If glngModel = PM_门诊医嘱清单 And .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4" Then
                        '获取中药医嘱组ID
                        str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                    Else
                        '取药品名称
                        If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                            str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                        Else
                            str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                        End If
                       
                        '取药品给药途径
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str用法 = ""    '一并给药不重复取
                        If str用法 = "" Then
                            If glngModel = PM_门诊编辑 Then
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                If k <> -1 Then
                                    If .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                                        str用法 = .TextMatrix(k, gobjCOL.intCOL用法)
                                    Else
                                        str用法 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                    End If
                                End If
                            Else
                                If Trim(.TextMatrix(i, gobjCOL.intCOL用法)) = "" Then
                                    str用法 = strPre用法
                                Else
                                    str用法 = Split(.TextMatrix(i, gobjCOL.intCOL用法), ",")(0)
                                End If
                                str用法ID = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOL相关ID)), "诊疗项目ID")   '传代码
                                strPre用法 = str用法
                            End If
                        End If
    
                        '取用药频率(次/天),都为整数四舍五入
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str频率 = ""    '一并给药不重复取
                        If str频率 = "" Then
                            If glngModel = PM_门诊编辑 Then
                                str频率 = GetFrequency(.TextMatrix(i, gobjCOL.intCOL间隔单位), .TextMatrix(i, gobjCOL.intCOL频率次数), .TextMatrix(i, gobjCOL.intCOL频率间隔))
                            Else
                                Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), int频率次数, int频率间隔, str间隔单位, IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), "")
                                str频率 = GetFrequency(str间隔单位, int频率次数 & "", int频率间隔 & "")
                            End If
                            str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                            If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                            str开嘱医生 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名") & "/" & str开嘱医生
                        End If
    
                        '传入医嘱信息
                        If glngModel = PM_门诊编辑 Then
                            Call PassSetRecipeInfo(.RowData(i), .TextMatrix(i, gobjCOL.intCOL收费细目ID), str药品, _
                                                   .TextMatrix(i, gobjCOL.intCOL单量), .TextMatrix(i, gobjCOL.intCOL单量单位), str频率, _
                                                   Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd"), "", str用法, _
                                                   .TextMatrix(i, gobjCOL.intCOL相关ID), 1, str开嘱医生)
                            If Not rsOut Is Nothing Then
                                If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                                '西药,中成药
                                    rsOut.AddNew
                                    rsOut!医嘱ID = CLng(.RowData(i) & "")
                                    rsOut!药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                                    rsOut!状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                                    rsOut!禁忌药品说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                                    rsOut.Update
                                ElseIf Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                                '中药配方  禁忌说明保存在用药服法上
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                    If k <> -1 Then
                                        rsOut.AddNew
                                        rsOut!医嘱ID = CLng(.RowData(k) & "")
                                        rsOut!药品名称 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                        rsOut!状态 = .TextMatrix(k, gobjCOL.intCOL状态)
                                        rsOut!禁忌药品说明 = .TextMatrix(k, gobjCOL.intCol禁忌药品说明)
                                        rsOut.Update
                                    End If
                                End If
                            End If
                        Else
                            strTmp = .TextMatrix(i, gobjCOL.intCOL单量)
                            If Mid(strTmp, 1, 2) = "0." Then
                                strTmp = "0" & Val(strTmp)
                            Else
                                strTmp = Val(strTmp)
                            End If
                            
                            Call PassSetRecipeInfo(.TextMatrix(i, gobjCOL.intCOLID), .TextMatrix(i, gobjCOL.intCOL收费细目ID), str药品, _
                                 strTmp, Replace(.TextMatrix(i, gobjCOL.intCOL单量), strTmp, ""), str频率, _
                                 Format(.TextMatrix(i, gobjCOL.intCOL开始时间), "yyyy-MM-dd"), "", str用法, _
                                 .TextMatrix(i, gobjCOL.intCOL相关ID), 1, str开嘱医生)
                        End If
                        lngCount = lngCount + 1
                    End If
                End If
            Next
            '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
            If glngModel = PM_门诊医嘱清单 Then
                If str中药组IDs <> "" Then
                    Set rs中药 = Get中药配方(str中药组IDs)
                    With rs中药
                        For i = 1 To .RecordCount
                            If !相关ID & "" <> str相关ID Then
                                str开嘱医生 = !开嘱医生
                                If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                                str开嘱医生 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名") & "/" & str开嘱医生
                                str频率 = GetFrequency(!间隔单位 & "", !频率次数 & "", !频率间隔 & "")
                                str相关ID = !相关ID & ""
                            End If
                            Call PassSetRecipeInfo(!id, !药品ID & "", !药品名称 & "", !单次用量 & "", !单量单位 & "", str频率, Format(!开始时间 & "", "yyyy-MM-dd"), _
                            "", !用法 & "", !相关ID & "", IIf(!医嘱期效 & "" = "0", "0", "1"), str开嘱医生)
                            
                            lngCount = lngCount + 1
                            .MoveNext
                        Next
                    End With
                End If
            End If
            '无可审查的药品
            If (lngCmd = 33 Or lngCmd = 34 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End With

    '执行相应的命令
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)

    '获取医嘱审查结果,并填写警示灯
    '-------------------------------------------------------------
    If lngCmd = 33 Or lngCmd = 34 Or lngCmd = 3 Then
        arrSQL = Array()
        '返回值顺：0-蓝灯,1-黄灯,2-红灯,3-黑灯,4-橙灯
        '警示级顺：0-蓝灯,1-黄灯,4-橙灯,2-红灯,3-黑灯(因为PASS升级的原因)
        arrLevel(0) = 0: arrLevel(1) = 1: arrLevel(2) = 3: arrLevel(3) = 4: arrLevel(4) = 2
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_门诊编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                    blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                    And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                End If
                If blnDo Then
                    If glngModel = PM_门诊编辑 Then
                        str医嘱ID = .RowData(i)
                    Else
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                    End If
                    k = PassGetWarn(str医嘱ID)
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        
                        '设置警示灯
                        If k >= 0 And k <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(k)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(k + 1).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        
                        If glngModel = PM_门诊编辑 Then
                            '标记审查结果变化,以备更新数据库
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                blnNoSave = True    '标记为未保存
                            End If
                            '记录下禁忌药品 K=3代表黑灯 且 只针对未校对医嘱进行禁忌药品说明原因的标记,已经校对发送的医嘱不处理
                            If k = 3 And Not rsOut Is Nothing Then
                                rsOut.Filter = "医嘱ID = " & str医嘱ID & " And 状态 < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                            End If
                        ElseIf PM_门诊医嘱清单 = glngModel Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                            End If
                        End If
                    Else
                        '中药配方
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                            lng中药组ID = .TextMatrix(i, gobjCOL.intCOL相关ID)          '中药配方组ID
                            lngLight = -1 '初始化
                        End If
                        '设置警示灯 取草药中最大警示值
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If arrLevel(k) > arrLevel(lngLight) Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                    End If
                    
                    '记录最高级别警示值
                    If k >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(k) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = k
                            End If
                        Else
                            lngMaxWarn = k
                        End If
                    End If
                Else
                    If glngModel = PM_门诊编辑 Then
                        '中药警示灯单独设置
                        If .RowData(i) = lng中药组ID And .RowData(i) <> 0 Then
                            strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                            '设置警示灯
                            If lngLight >= 0 And lngLight <= 4 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(lngLight + 1).Picture
                            Else
                                .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                            End If
                            
                            If glngModel = PM_门诊编辑 Then
                                '标记审查结果变化,以备更新数据库
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                    blnNoSave = True    '标记为未保存
                                End If
                                '记录下禁忌药品 K=3代表黑灯
                                If lngLight = 3 And Not rsOut Is Nothing Then
                                    rsOut.Filter = "医嘱ID = " & lng中药组ID & " And 状态 < 3 "
                                    If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                                End If
                            End If
                            lng中药组ID = 0
                            lngLight = -1
                        End If
                    End If
                End If
            Next
            '医嘱清单中药配方警示灯处理
            If glngModel = PM_门诊医嘱清单 And Not rs中药 Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    '中药服法
                    If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        lngLight = -1
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                        rs中药.Filter = "相关ID=" & str医嘱ID
                        
                        For j = 1 To rs中药.RecordCount
                            k = PassGetWarn(rs中药!id & "")
                            '设置警示灯 取草药中最大警示值
                            If k >= 0 Then
                                If lngLight >= 0 Then
                                    If arrLevel(k) > arrLevel(lngLight) Then
                                        lngLight = k
                                    End If
                                Else
                                    lngLight = k
                                End If
                            End If
                            rs中药.MoveNext
                        Next
                        
                        '设置警示灯
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(lngLight + 1).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        '警示灯更新到数据库
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(lngLight >= 0 And lngLight <= 4, lngLight, "NULL") & ")"
                        End If
                            
                        '记录最高级别警示值
                        If lngLight >= 0 Then
                            If lngMaxWarn >= 0 Then
                                If arrLevel(lngLight) > arrLevel(lngMaxWarn) Then
                                    lngMaxWarn = lngLight
                                End If
                            Else
                                lngMaxWarn = lngLight
                            End If
                        End If
                    End If
                Next
            End If
        End With
        
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
        Next
        
    End If
    '返回审查结果
    OutAdviceCheckWarn_MK = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function OperateAdviceCheckWarn_MK(ByVal lngCmd As Long, ByVal lngRow As Long) As Long
'功能：调用Pass系统相关功能,护士站校对时调用
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        21-病生状态/过敏史管理(只读)
'      lngRow=当前药品医嘱的行号:lngCmd=0时需要,多病人批量操作时需要当前病人行
'返回：检测PASS菜单时，返回>=0表示可以弹出菜单,其它返回-1
'说明：用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset
    Dim str药品 As String, str用法 As String
    Dim str药品ID As String
    Dim lng病人ID As Long, lng主页ID As Long
    Dim strSQL As String, i As Long, k As Long
    Dim strCurrentDate As String

    OperateAdviceCheckWarn_MK = -1
    If Not (lngRow >= gobjAdvice.FixedRows) Then Exit Function    '必须要确定病人所在行

    On Error GoTo errH
    Screen.MousePointer = 11
    If gstrVersion = "3.0" Then
    '美康3.0
    
        '检验PASS可用状态
        '-------------------------------------------------------------
        If PassGetState("PassEnable") = 0 Then
            MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
            Screen.MousePointer = 0: Exit Function
        End If
    
        '114036同一个病人多次审查时病人信息每次都要传入
        '-------------------------------------------------------------
        lng病人ID = Val(gobjAdvice.TextMatrix(lngRow, gobjCOL.intCOL病人ID))
        lng主页ID = Val(gobjAdvice.TextMatrix(lngRow, gobjCOL.intCOL主页ID))
      
        strSQL = _
        " Select NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,A.出生日期,B.身高,B.体重,B.入院日期,B.出院日期," & _
                 " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
                 " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
                 " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
                 " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng病人ID, lng主页ID)
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

        Call PassSetPatientInfo(lng病人ID, lng主页ID, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), rsTmp!体重 & "", rsTmp!身高 & "", _
                                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), _
                                IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))

        '传人病人过敏史
        '-------------------------------------------------------
        Set rsTmp = Get病人过敏记录(lng病人ID, lng主页ID)

        For i = 1 To rsTmp.RecordCount
            Call PassSetAllergenInfo(i, rsTmp!药物ID & "", rsTmp!药物名 & "", "DrugName", "")
            rsTmp.MoveNext
        Next

        '传人病生状态
        '------------------------------------------------------------------
        Set rsTmp = Get病人诊断记录(lng病人ID, lng主页ID, "2,12")
        strCurrentDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")

        For i = 1 To rsTmp.RecordCount
            Call PassSetMedCond(i & "", rsTmp!编码 & "", rsTmp!名称 & "", "User", strCurrentDate, strCurrentDate)
            rsTmp.MoveNext
        Next
     
    
        'PASS自定义菜单检测
        '-------------------------------------------------------------
        If lngCmd = 0 Then
            With gobjAdvice
                If Val(.TextMatrix(lngRow, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) > 0 Then
                    '取药品名称
                    If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) > 0 Then
                        str药品 = .TextMatrix(lngRow, gobjCOL.intCOL药品名称)
                    Else
                        str药品 = .TextMatrix(lngRow, gobjCOL.intCOL医嘱内容) '中药名称
                    End If
                        
                    '取药品给药途径(当前可见行不会是中草药)
                    str用法 = .TextMatrix(lngRow, gobjCOL.intCOL用法)
                    
                    '药品长期医嘱按品种下达,传任意药品ID
                    If Val(.TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)) = 0 Then
                        str药品ID = GetDrugID(.TextMatrix(lngRow, gobjCOL.intCOL诊疗项目ID))
                    Else
                        str药品ID = .TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)
                    End If
                    
                    '传入查询药品信息
                    Call PassSetQueryDrug(str药品ID, str药品, .TextMatrix(lngRow, gobjCOL.intCOL单量单位), str用法)
                    
                    OperateAdviceCheckWarn_MK = 1    '表示可以弹出菜单
                End If
            End With
            Screen.MousePointer = 0: Exit Function
        End If
    
        '执行相应的命令
        '-------------------------------------------------------------
        Call PassDoCommand(lngCmd)
    ElseIf gstrVersion = "4.0" Then
    '美康4.0
        With gobjAdvice
            Select Case lngCmd
            
            Case MK4_检测PASS菜单状态
               
                If Val(.TextMatrix(lngRow, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) > 0 Then
                    OperateAdviceCheckWarn_MK = 1    '表示可以弹出菜单
                End If
                Screen.MousePointer = 0: Exit Function
            Case 1
            
            End Select
       End With
    End If
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function OutAdviceCheckWarn_DT() As Boolean
'功能：调用大通用药监测系统对医嘱进行合理用药审查等相关功能
    Dim xmlbase As dt_base, xmlpre As dt_Pres
    Dim strTmp As String, arrTmp As Variant, curDate As Date
    Dim rsTmp As Recordset
    Dim i As Long, k As Long, blnDo As Boolean
    Dim str药品 As String, str给药途径 As String, str频率编码 As String, strXML As String
    Dim arrDiagName(1 To 3) As String, arrDiagCode(1 To 3) As String
    Dim strRetXML As String
    Dim str单量 As String
    Dim str单量单位 As String
    Dim lng医嘱ID As Long
    
    On Error GoTo errH

    curDate = zlDatabase.Currentdate
    With xmlbase
        .dDoctCode = UserInfo.用户名
        .dDoctName = UserInfo.姓名
        .dDoctType = UserInfo.专业技术职务
        .dDeptCode = UserInfo.部门ID
        .dDeptName = UserInfo.部门名
        .dInHosCode = ""
        .dBedNo = ""
        .mPresDate = curDate
        .pCaseID = gobjPati.lng病人ID
        .pOutID = gobjPati.str挂号单
        .pWeight = ""
        .pHeight = ""
        .pBirthday = NVL(gobjPati.dat出生日期, vbNull)
        .pPatiName = gobjPati.str姓名
        .pSex = gobjPati.str性别
        .pStatms = ""
        .pEffect = ""
        .pBloodPress = ""
        .pLiverClean = ""
        
        '* 过敏源
        .pCaseCode1 = ""
        .pCaseName1 = ""
        .pCaseCode2 = ""
        .pCaseName2 = ""
        .pCaseCode3 = ""
        .pCaseName3 = ""
        Set rsTmp = Get病人过敏记录(gobjPati.lng病人ID, 0)
        If rsTmp.RecordCount > 0 Then
            .pCaseCode1 = "" & rsTmp!药物ID
            .pCaseName1 = rsTmp!药物名
            rsTmp.MoveNext
            
            If Not rsTmp.EOF Then
                .pCaseCode2 = "" & rsTmp!药物ID
                .pCaseName2 = rsTmp!药物名
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pCaseCode3 = "" & rsTmp!药物ID
                    .pCaseName3 = rsTmp!药物名
                End If
            End If
        End If
        
        '* 诊断信息
        .pDiagnose1 = ""
        .pDiagnose2 = ""
        .pDiagnose3 = ""
        .pDiagnoseName1 = ""
        .pDiagnoseName2 = ""
        .pDiagnoseName3 = ""
        If glngModel = PM_门诊编辑 Then
            k = 1
            If Not gobjDiags Is Nothing Then
                With gobjDiags
                    For i = 1 To .Count
                        If .Item(i).str诊断描述 <> "" Then
                            arrDiagCode(i) = IIf(.Item(i).str疾病编码 <> "", .Item(i).str疾病编码, .Item(i).str诊断编码)
                            arrDiagName(i) = .Item(i).str诊断描述
                            If k = 3 Then Exit For
                            k = k + 1
                        End If
                    Next
                End With
            End If
            .pDiagnose1 = arrDiagCode(1)
            .pDiagnose2 = arrDiagCode(2)
            .pDiagnose3 = arrDiagCode(3)
            .pDiagnoseName1 = arrDiagName(1)
            .pDiagnoseName2 = arrDiagName(2)
            .pDiagnoseName3 = arrDiagName(3)
        ElseIf glngModel = PM_门诊医嘱清单 Then
            Set rsTmp = Get病人诊断记录(gobjPati.lng病人ID, gobjPati.lng挂号ID, "1")
            If rsTmp.RecordCount > 0 Then
                .pDiagnose1 = "" & rsTmp!编码
                .pDiagnoseName1 = "" & rsTmp!名称
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pDiagnose2 = "" & rsTmp!编码
                    .pDiagnoseName2 = "" & rsTmp!名称
                    rsTmp.MoveNext
                    If Not rsTmp.EOF Then
                        .pDiagnose3 = "" & rsTmp!编码
                        .pDiagnoseName3 = "" & rsTmp!名称
                    End If
                End If
            End If
        End If
        
        '* 病生理状态
        .pBsl1 = ""
        .pBsl2 = ""
        .pBsl3 = ""
        strTmp = Get病人病生理情况(gobjPati.lng病人ID, 0)
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            .pBsl1 = arrTmp(0)
            If UBound(arrTmp) > 0 Then .pBsl2 = arrTmp(1)
            If UBound(arrTmp) > 1 Then .pBsl3 = arrTmp(2)
        End If
    End With
        
    arrTmp = Array()
    With gobjAdvice
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_门诊编辑 Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                    And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                    And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            ElseIf glngModel = PM_门诊医嘱清单 Then
                blnDo = Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                    And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            End If

            If blnDo Then
                '取药品名称
                If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                    str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                Else
                    str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                End If
                
                '取药品给药途径
                If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str给药途径 = "" '一并给药不重复取
                If str给药途径 = "" Then
                    If glngModel = PM_门诊编辑 Then
                        k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                        If k <> -1 Then str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                    ElseIf glngModel = PM_门诊医嘱清单 Then
                        str给药途径 = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOL相关ID)), "诊疗项目ID")  '传代码
                    End If
                End If
                
                Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), 0, 0, "", IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), str频率编码)
                
                If glngModel = PM_门诊编辑 Then
                    str单量 = StrToXML(.TextMatrix(i, gobjCOL.intCOL单量))
                    str单量单位 = StrToXML(.TextMatrix(i, gobjCOL.intCOL单量单位))
                    lng医嘱ID = .RowData(i)
                ElseIf glngModel = PM_门诊医嘱清单 Then
                    lng医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                    str单量 = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL单量)))
                    If Mid(str单量, 1, 2) = "0." Then
                        str单量 = "0" & Val(str单量)
                    Else
                        str单量 = Val(str单量)
                    End If
                    str单量单位 = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL单量)))
                    If Mid(str单量单位, 1, 2) = "0." Then '单量有小数点的特殊处理
                        str单量单位 = Replace(str单量单位, Format(Val(str单量单位) & "", "0.####"), "") '门诊清单单量（“单量” & “单量单位”）
                    Else
                        str单量单位 = Replace(str单量单位, Val(str单量单位) & "", "")    '
                    End If
                End If
                
                xmlpre.PresID = lng医嘱ID
                xmlpre.PresType = "mz"
                xmlpre.Current = 1
                xmlpre.GeneralName = StrToXML(Sys.RowValue("诊疗项目目录", Val(.TextMatrix(i, gobjCOL.intCOL诊疗项目ID)), "名称"))
                xmlpre.HosMediCode = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                xmlpre.MediName = StrToXML(str药品)
                
                xmlpre.DCL = str单量
                xmlpre.PCDM = StrToXML(str频率编码)
                xmlpre.Days = StrToXML(.TextMatrix(i, gobjCOL.intCOL天数))
                
                xmlpre.Unit = str单量单位
                xmlpre.GYTJ = str给药途径
                xmlpre.GroupNum = Val(.TextMatrix(i, gobjCOL.intCOL相关ID))
                
                strXML = MakePresXML(xmlpre, 0)
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = strXML
            End If
        Next
    End With
    
        
    OutAdviceCheckWarn_DT = True
    If UBound(arrTmp) >= 0 Then
        strXML = MakeXML(xmlbase, arrTmp, 0)
        WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strXML
        
        If gbytSuperVolume = 0 Then
            strTmp = dtywzxUI2(4, 0, strXML, strRetXML)
            WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strTmp
            strRetXML = GetAlertFromXml(strRetXML)
            If InStr(strRetXML, ";CJLJJ;") > 0 Then
                MsgBox "用药监测系统发现当前医嘱存在超极量禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                OutAdviceCheckWarn_DT = False
            End If
            strRetXML = ""
        Else
            strTmp = dtywzxUI(4, 0, strXML)
            WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strTmp
        End If
        '
        If glngModel = PM_门诊编辑 Then
            If strTmp = "2" And gbytBlackLamp = 0 Then
                MsgBox "用药监测系统发现当前医嘱存在禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                OutAdviceCheckWarn_DT = False
            ElseIf strTmp = "1" Or strTmp = "2" And gbytBlackLamp = 1 Then
                If MsgBox("用药监测系统发现当前医嘱存在禁忌用药，是否继续?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then OutAdviceCheckWarn_DT = False
            End If
            If OutAdviceCheckWarn_DT Then
                If gbytSuperVolume = 0 Then
                    strTmp = dtywzxUI2(13, 0, strXML, strRetXML)
                    WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strTmp
                    strRetXML = GetAlertFromXml(strRetXML)
                    If InStr(strRetXML, ";CJLJJ;") > 0 Then
                        MsgBox "用药监测系统发现当前医嘱存在超极量禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                        OutAdviceCheckWarn_DT = False
                    End If
                    strRetXML = ""
                Else
                    strTmp = dtywzxUI(13, 0, strXML)
                    WriteLog "" & glngModel, "OutAdviceCheckWarn_DT", strTmp
                End If
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    OutAdviceCheckWarn_DT = False
End Function

Public Function OutAdviceCheckWarn_TYT(ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional ByRef blnNoSave As Boolean, _
    Optional ByRef rsOut As ADODB.Recordset) As Long
'功能：调用太元通系统中对医嘱进行合理用药审查等相关功能
'参数：lngCmd=
'       0-用药规范
'       1-获取医嘱审查结果,并填写警示灯
'       2-药品提示
'       3-医药知识库，4-系统配置;5-获取警示详情
'
'      lngRow=当前药品医嘱的行号，lngCmd=2时需要
'出参：
'      rsOut-禁忌说明
'返回值：医嘱保存调用，需要用返回值判断是否存在禁忌用药
    Dim strDrugCode As String, str医生编码 As String, str开嘱医生 As String, strDescription As String
    Dim str单量 As String, str单量单位 As String
    Dim str医嘱序号 As String
    Dim strSQL As String, strOrderInfo As String, str频率编码 As String, str频率 As String
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim str给药途径 As String, str药品 As String, str中药组IDs As String
    Dim str开嘱科室ID As String, str相关ID As String, str医嘱ID As String
    
    Dim blnDo As Boolean
    Dim curDate As Date
    Dim rsPati As ADODB.Recordset, rs中药 As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim udtPatiOrder As PatientOrder
    Dim udtDrug As PatDrug, udtPatiDiag As PatDiagnosis
    Dim udtPatiSensitive As PatDrugSensitive, UdtPatiSymptom As PatSymptom
    Dim udtAuditResult As AuditResult

    Dim i As Long, k As Long, j As Long, lngMaxWarn As Long, lng医嘱ID As Long
    Dim lng中药组ID As Long, lngLight As Long
    
    Dim strTmp As String, strOld As String

    Dim arrTmp As Variant, colAuditResult As Collection
    Dim arrLight(1 To 3) As String
    
    On Error GoTo errH
    Screen.MousePointer = 11

    With gobjAdvice
        Select Case lngCmd

        Case 0  '0-用药规范

            gobjPass.getPdssPrescription

        Case 1  '1-获取医嘱审查结果,并填写警示灯
        
            If glngModel = PM_门诊编辑 And gbytReason = 1 Then
                Set rsOut = InitAdviceRS(FUN_输出内容)
            End If
                
            If glngModel = PM_门诊医嘱清单 Then
                Set rsTmp = ReadPatient(gobjPati.lng病人ID, gobjPati.str挂号单)
                gobjPati.str姓名 = rsTmp!姓名 & ""
                gobjPati.str性别 = rsTmp!性别 & ""
                gobjPati.dat出生日期 = CDate(rsTmp!出生日期 & "")
                gobjPati.lng挂号ID = rsTmp!就诊Id
            End If
            
            '病人信息
            With udtPatiOrder
                '传人病人信息:病人ID,姓名,性别 1-女, 0-男, 2-不详，病人出生日期，格式 YYYY-MM-DD 不为空（必填）
                
                .PatientID = gobjPati.lng病人ID & ""
                .Pname = gobjPati.str姓名
                .pSex = IIf(gobjPati.str性别 = "男", "0", IIf(gobjPati.str性别 = "女", "1", "2"))
                .pdateOfBirth = Format(gobjPati.dat出生日期, "yyyy-MM-dd")

                '附加信息
                strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                        "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                        "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng病人ID, gobjPati.lng挂号ID)
                rsTmp.Filter = "项目名称='身高'"
                If rsTmp.RecordCount <> 0 Then .pHeight = IIf(Val(rsTmp!记录内容 & "") = 0, "", rsTmp!记录内容 & "")
                rsTmp.Filter = "项目名称='体重'"
                If rsTmp.RecordCount <> 0 Then .pWeight = IIf(Val(rsTmp!记录内容 & "") = 0, "", rsTmp!记录内容 & "")

                .PvisitID = gobjPati.dbl标识号 & ""

                '传人病人生理情况
                strTmp = Get病人病生理情况(gobjPati.lng病人ID, 0)
                .isLact = IIf(InStr(strTmp, "哺乳期") > 0, "1", "0")    '是否哺乳，是为1，否为0 不为空
                .isPregnant = IIf(InStr(strTmp, "孕妇") > 0, "1", "0")    '是否孕妇，是为1 ，否为0 不为空
                .isLiverWhole = IIf(InStr(strTmp, "肝功能异常") > 0, "1", "0") '是否肝功异常 1-异常，0-正常 不为空
                .isKidneyWhole = IIf(InStr(strTmp, "肾功能异常") > 0, "1", "0") '是否肾功异常 1-异常，0-正常 不为空

                '登录医生信息
                .DoctDeptID = UserInfo.部门ID & ""
                .DoctDeptName = UserInfo.部门名 & ""
                .DoctID = UserInfo.编号 & ""
                .DoctName = UserInfo.姓名 & ""
                .DoctTitleID = GetDoctorTitleType(UserInfo.专业技术职务)
                .DoctTitleName = IIf(UserInfo.专业技术职务 = "", "其他职务", UserInfo.专业技术职务)
                .SysFlag = "1"  '2-住院医生站，1-门诊医生站
            End With

            '药品信息
            curDate = zlDatabase.Currentdate
            arrTmp = Array()
            With gobjAdvice

                For i = .FixedRows To .Rows - 1
                    If glngModel = PM_门诊编辑 Then
                        blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                                And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                        blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                        blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL状态)) <> 4  '临时医嘱作废不审查
                    ElseIf glngModel = PM_门诊医嘱清单 Then
                        blnDo = (Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                                And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0) Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4")
                        blnDo = blnDo And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                        blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL状态)) <> 4  '临时医嘱作废不审查
                    End If
                    
                    If blnDo Then
                    
                        If glngModel = PM_门诊医嘱清单 And .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4" Then
                             '获取中药医嘱组ID
                            str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                        Else
                            '取药品名称
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                                str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                            Else
                                str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                            End If
    
                            If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then  '一并给药不重复取
                                '给药途径
                                If glngModel = PM_门诊编辑 Then
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                    If k <> -1 Then str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                                ElseIf glngModel = PM_门诊医嘱清单 Then
                                    str给药途径 = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOL相关ID)), "诊疗项目ID")  '传代码
                                End If
                                '取用药频率(次/天),都为整数四舍五入
                                Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), int频率次数, int频率间隔, str间隔单位, IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), str频率编码)
    
                                If str间隔单位 = "天" Then
                                    str频率 = int频率次数 & "/" & int频率间隔
                                ElseIf str间隔单位 = "周" Then
                                    str频率 = int频率次数 & "/7"
                                ElseIf str间隔单位 = "小时" Then
                                    If int频率间隔 <= 24 Then
                                        str频率 = Format(24 / int频率间隔 * int频率次数, "0") & "/1"
                                    Else
                                        str频率 = int频率次数 & "/" & Format(int频率间隔 / 24, "0")
                                    End If
                                ElseIf str间隔单位 = "分钟" Then
                                    str频率 = Format((24 * 60) / int频率间隔 * int频率次数, "0") & "/1"
                                End If
    
                                str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                                If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                                str医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
                            End If
                            
                            If glngModel = PM_门诊编辑 Then
                                lng医嘱ID = Val(.RowData(i) & "")
                                str单量 = .TextMatrix(i, gobjCOL.intCOL单量)
                                str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                                str开嘱科室ID = .TextMatrix(i, gobjCOL.intCOL开嘱科室ID)
                                str医嘱序号 = .TextMatrix(i, gobjCOL.intCOL序号)
                            ElseIf glngModel = PM_门诊医嘱清单 Then
                                str单量 = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL单量)))
                                If Mid(str单量, 1, 2) = "0." Then
                                    str单量 = "0" & Val(str单量)
                                Else
                                    str单量 = Val(str单量)
                                End If
                                
                                str单量单位 = Trim(.TextMatrix(i, gobjCOL.intCOL单量))
                                If Mid(str单量单位, 1, 2) = "0." Then '单量有小数点的特殊处理
                                    str单量单位 = Replace(str单量单位, Format(Val(str单量单位) & "", "0.####"), "") '门诊清单单量（“单量” & “单量单位”）
                                Else
                                    str单量单位 = Replace(str单量单位, Val(str单量单位) & "", "")    '
                                End If
                                Set rsTmp = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOLID)))
                                str开嘱科室ID = rsTmp!开嘱科室id & ""
                                str医嘱序号 = rsTmp!序号 & ""
                            End If
                            udtDrug.drugID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)    'his 系统的药品代码不为空
                            udtDrug.DrugName = StrToXML(str药品)               'his 系统的药品名称不为空
                            udtDrug.recMainNo = .TextMatrix(i, gobjCOL.intCOL相关ID)     'his 系统的医嘱组号，在一次就诊/住院中唯
                            udtDrug.recSubNo = str医嘱序号       'his 系统的医嘱序号，在一次就诊/住院中唯
                            udtDrug.dosage = str单量      'his 系统的医嘱药品使用剂量不为空
    
                            udtDrug.doseUnits = str单量单位    'his 系统的医嘱药品剂量单位不为空
                            udtDrug.administrationID = str给药途径              'his 系统的医嘱途径代码不为空
                            udtDrug.performFreqDictID = StrToXML(str频率编码)   'his 系统的医嘱频次代码不为空
                            udtDrug.performFreqDictText = str频率               'his 系统的医嘱执行频率描述不为空
    
                            udtDrug.startDateTime = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:mm:ss")    'his 系统的医嘱开始时间,格式 YYYY-MM-DDHH: MM: SS 不为空
                            udtDrug.stopDateTime = ""                           'his 系统的医嘱结束时间,格式 YYYY-MM-DD HH: MM: SS
                            
                            udtDrug.doctorDept = str开嘱科室ID   'his 系统的开医嘱医生所在科室代码
                            udtDrug.DoctorID = str医生编码                          'his 系统的开医嘱医生编码
                            udtDrug.Doctor = str开嘱医生                         'his 系统的开医嘱医生姓名,
                            If glngModel = PM_门诊编辑 Then
                                udtDrug.isNew = IIf(.TextMatrix(i, gobjCOL.intCOLEDIT) = "1", "1", "0")    '新增医嘱值为1；否则为0
                            Else
                                udtDrug.isNew = "0"
                            End If
                            
                            If Not rsOut Is Nothing Then
                                If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                                    rsOut.AddNew
                                    rsOut!医嘱ID = lng医嘱ID
                                    rsOut!药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                                    rsOut!状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                                    rsOut!禁忌药品说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                                    rsOut.Update
                                ElseIf Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                                    '中药配方  禁忌说明保存在用药服法上
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                    If k <> -1 Then
                                        rsOut.AddNew
                                        rsOut!医嘱ID = CLng(.RowData(k) & "")
                                        rsOut!诊疗类别 = .TextMatrix(k, gobjCOL.intCOL诊疗类别)
                                        rsOut!药品名称 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                        rsOut!状态 = .TextMatrix(k, gobjCOL.intCOL状态)
                                        rsOut!禁忌药品说明 = .TextMatrix(k, gobjCOL.intCol禁忌药品说明)
                                        rsOut.Update
                                    End If
                                End If
                            End If
                            
                            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                            arrTmp(UBound(arrTmp)) = udtDrug
                        End If
                    End If
                Next
                
                '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
                If glngModel = PM_门诊医嘱清单 Then
                    If str中药组IDs <> "" Then
                        Set rs中药 = Get中药配方(str中药组IDs)
                        With rs中药
                            For i = 1 To .RecordCount
                                If !相关ID & "" <> str相关ID Then
                                    str开嘱医生 = !开嘱医生
                                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                                    str医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
                                    str频率 = GetFrequency(!间隔单位 & "", !频率次数 & "", !频率间隔 & "")
                                    Call Get频率信息_名称(!频率 & "", CInt(!频率次数 & ""), CInt(!频率间隔 & ""), !间隔单位 & "", 2, str频率编码)
                                    str相关ID = !相关ID
                                End If
                                udtDrug.drugID = !药品ID & ""    'his 系统的药品代码不为空
                                udtDrug.DrugName = !药品名称 & ""              'his 系统的药品名称不为空
                                udtDrug.recMainNo = !相关ID & ""     'his 系统的医嘱组号，在一次就诊/住院中唯
                                udtDrug.recSubNo = !序号 & ""       'his 系统的医嘱序号，在一次就诊/住院中唯
                                udtDrug.dosage = !单次用量 & ""      'his 系统的医嘱药品使用剂量不为空
        
                                udtDrug.doseUnits = !单量单位 & ""    'his 系统的医嘱药品剂量单位不为空
                                udtDrug.administrationID = !用法ID & ""              'his 系统的医嘱途径代码不为空
                                udtDrug.performFreqDictID = str频率编码    'his 系统的医嘱频次代码不为空
                                udtDrug.performFreqDictText = str频率               'his 系统的医嘱执行频率描述不为空
        
                                udtDrug.startDateTime = Format(!开始时间, "yyyy-MM-dd HH:mm:ss")    'his 系统的医嘱开始时间,格式 YYYY-MM-DDHH: MM: SS 不为空
                                udtDrug.stopDateTime = ""                           'his 系统的医嘱结束时间,格式 YYYY-MM-DD HH: MM: SS
                                
                                udtDrug.doctorDept = !开嘱科室id & ""  'his 系统的开医嘱医生所在科室代码
                                udtDrug.DoctorID = str医生编码                          'his 系统的开医嘱医生编码
                                udtDrug.Doctor = str开嘱医生                         'his 系统的开医嘱医生姓名,
                        
                                udtDrug.isNew = "0"
                                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                arrTmp(UBound(arrTmp)) = udtDrug
                                .MoveNext
                            Next
                        End With
                    End If
                End If
            End With
            If UBound(arrTmp) = -1 Then
                Screen.MousePointer = 0: Exit Function
            End If
            udtPatiOrder.PatDrugs = arrTmp

            '诊断
            arrTmp = Array()
            If glngModel = PM_门诊编辑 Then
                If Not gobjDiags Is Nothing Then
                    With gobjDiags
                        For i = 1 To .Count
                            If .Item(i).str诊断描述 <> "" Then
                                udtPatiDiag.diagnosisID = IIf(.Item(i).str疾病编码 <> "", .Item(i).str疾病编码, .Item(i).str诊断编码) 'his 系统的诊断编码
                                udtPatiDiag.diagnosisName = .Item(i).str诊断描述           'his 系统的诊断名称
                                udtPatiDiag.diagnosisType = "门诊诊断"                     '系统的诊断类型，如门诊诊断、入院诊断等
                                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                arrTmp(UBound(arrTmp)) = udtPatiDiag
                            End If
                        Next
                    End With
                End If
            Else
                Set rsTmp = Get病人诊断记录(gobjPati.lng病人ID, gobjPati.lng挂号ID, "1,11")
                For i = 1 To rsTmp.RecordCount
                    udtPatiDiag.diagnosisID = rsTmp!编码 'his 系统的诊断编码
                    udtPatiDiag.diagnosisName = rsTmp!名称          'his 系统的诊断名称
                    udtPatiDiag.diagnosisType = "门诊诊断"                     '系统的诊断类型，如门诊诊断、入院诊断等
                    ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                    arrTmp(UBound(arrTmp)) = udtPatiDiag
                    rsTmp.MoveNext
                Next
            End If
            
            udtPatiOrder.PatDiagnoses = arrTmp
            '过敏
            arrTmp = Array()
            Set rsTmp = Get病人过敏记录(gobjPati.lng病人ID, 0)
            For i = 0 To rsTmp.RecordCount - 1
                udtPatiSensitive.patOrderDrugSensitiveID = "0"          '固定值
                udtPatiSensitive.drugAllergenID = rsTmp!过敏源编码 & ""    '系统的过敏编码
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = udtPatiSensitive
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatDrugSensitives = arrTmp
            '症状
            arrTmp = Array()
            Set rsTmp = GetPatiSymptom(gobjPati.lng病人ID, gobjPati.lng挂号ID)
            For i = 0 To rsTmp.RecordCount - 1
                UdtPatiSymptom.symptomID = rsTmp!编码 & ""    'his 系统的症状编码
                UdtPatiSymptom.symptomName = rsTmp!名称 & ""  'his 系统的症状名称
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = UdtPatiSymptom
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatSymptoms = arrTmp

            strOrderInfo = MakePatientOrderXml(udtPatiOrder)

            '医嘱信息审查接口调用"

            strDescription = gobjPass.checkDrugSecurityWS(strOrderInfo, "1")

            '审查结果处理
            '返回值顺及警示级别：1— 禁忌（建议显示红色警示灯）；2— 慎用（建议显示黄色警示灯示）；3— 提示（建议显示蓝色警示灯）
            '图标颜色frmIcons.imgpassTYT ：1-红，2-黄，3-蓝
            arrLight(1) = "红": arrLight(2) = "黄": arrLight(3) = "蓝"
            lngMaxWarn = 4
            If glngModel = PM_门诊医嘱清单 Then arrTmp = Array()
            If strDescription = "" Then
                MsgBox "药嘱审查功能未执行，请检查太元通接口配置是否有误！", vbInformation + vbOKOnly, G_STR_PASS
                Screen.MousePointer = 0: Exit Function
            ElseIf strDescription = "-101" Then
                '-101：表示用户可以忽略该返回值，不做业务处理。
            Else
                Set colAuditResult = AnalyzeReturnXml(strDescription)
                With gobjAdvice
                    For i = .FixedRows To .Rows - 1
                        If glngModel = PM_门诊编辑 Then
                            blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                                    And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                            blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL状态)) <> 4 And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                        Else
                            blnDo = Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                                   And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                            blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL状态)) <> 4 And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                        End If
                        
                        If blnDo Then

                            '获取警示灯
                            If glngModel = PM_门诊编辑 Then
                                strTmp = .TextMatrix(i, gobjCOL.intCOL相关ID) & "_" & .TextMatrix(i, gobjCOL.intCOL序号)   '关键字格式:组医嘱号_医嘱序号
                                lng医嘱ID = CLng(.RowData(i) & "")
                            Else
                                str医嘱序号 = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOLID)), "序号")
                                strTmp = .TextMatrix(i, gobjCOL.intCOL相关ID) & "_" & str医嘱序号   '关键字格式:组医嘱号_医嘱序号
                            End If
                            On Error Resume Next
                            udtAuditResult = colAuditResult(strTmp)
                            If Err.Number > 0 Then
                                strTmp = "未找到"
                            End If
                            Err.Clear: On Error GoTo 0
                            If strTmp <> "未找到" Then  '找到审核警示灯
                                k = Val(udtAuditResult.alertLevel)
                            Else
                                k = 0
                            End If
                            
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                                '设置警示灯
                                strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                                If k >= 1 And k <= 3 Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(k)
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                                Else
                                    .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                                End If
                                
                                If glngModel = PM_门诊编辑 Then
                                    '标记审查结果变化,以备更新数据库
                                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                        .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                        blnNoSave = True    '标记为未保存
                                    End If
                                ElseIf glngModel = PM_门诊医嘱清单 Then
                                     '警示灯更新到数据库
                                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                        ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                        arrTmp(UBound(arrTmp)) = "ZL_病人医嘱记录_更新审查(" & .TextMatrix(i, gobjCOL.intCOLID) & "," & IIf(k >= 1 And k <= 3, k, "NULL") & ")"
                                    End If
                                End If
                                
                                '记录下禁忌药品 K=1代表红色警示灯
                                If gbytReason = 1 And k = 1 And Not rsOut Is Nothing Then
                                    rsOut.Filter = "医嘱ID = " & lng医嘱ID & " And 状态 < 3 "
                                    If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                                End If
                            Else
                                 '中药配方
                                If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                                    lng中药组ID = .TextMatrix(i, gobjCOL.intCOL相关ID)          '中药配方组ID
                                    lngLight = 4 '初始化
                                End If
                                If k > 0 Then
                                    If lngLight > k Then
                                        lngLight = k
                                    End If
                                End If
                            End If
                            '记录最高级别警示值
                            If k > 0 Then
                                If lngMaxWarn > k Then
                                    lngMaxWarn = k
                                End If
                            End If
                            
                        Else
                            If glngModel = PM_门诊编辑 Then
                                '中药警示灯单独设置
                                If .RowData(i) = lng中药组ID And .RowData(i) <> 0 Then
                                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                                    '设置警示灯
                                    If lngLight >= 1 And lngLight <= 3 Then
                                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                                    Else
                                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                                    End If
                                    
                                    If glngModel = PM_门诊编辑 Then
                                        '标记审查结果变化,以备更新数据库
                                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                            .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                            blnNoSave = True    '标记为未保存
                                        End If
                                        '记录下禁忌药品 K=3代表黑灯
                                        If lngLight = 1 And Not rsOut Is Nothing Then
                                            rsOut.Filter = "医嘱ID = " & lng中药组ID & " And 状态 < 3 "
                                            If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                                        End If
                                    End If
                                    
                                    lng中药组ID = 0
                                    lngLight = 4
                                End If
                            End If
                        End If
                    Next
                    '医嘱清单中药配方警示灯处理
                    If glngModel = PM_门诊医嘱清单 And Not rs中药 Is Nothing Then
                        For i = .FixedRows To .Rows - 1
                            '中药服法
                            If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                                strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                                lngLight = 4
                                str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                                rs中药.Filter = "相关ID=" & str医嘱ID
                                
                                For j = 1 To rs中药.RecordCount
                                    strTmp = rs中药!相关ID & "_" & rs中药!序号  '关键字格式:组医嘱号_医嘱序号
                                    On Error Resume Next
                                    udtAuditResult = colAuditResult(strTmp)
                                    If Err.Number > 0 Then
                                        strTmp = "未找到"
                                    End If
                                    Err.Clear: On Error GoTo 0
                                    If strTmp <> "未找到" Then  '找到审核警示灯
                                        k = Val(udtAuditResult.alertLevel)
                                    Else
                                        k = 0
                                    End If
                                    '设置警示灯 取草药中最大警示值
                                    '记录最高级别警示值
                                    If k > 0 Then
                                        If lngLight > k Then
                                            lngLight = k
                                        End If
                                    End If
                                    rs中药.MoveNext
                                Next
                                
                                '设置警示灯
                                If lngLight >= 1 And lngLight <= 3 Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                                Else
                                    .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                                End If
                                '警示灯更新到数据库
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                    ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                    arrTmp(UBound(arrTmp)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(lngLight >= 1 And lngLight <= 3, lngLight, "NULL") & ")"
                                End If

                                '记录最高级别警示值
                                If lngLight > 0 Then
                                    If lngMaxWarn > lngLight Then
                                        lngMaxWarn = lngLight
                                    End If
                                End If
                            End If
                        Next
                    End If
                    
                End With
                '数据提交,不开启事务
                If glngModel = PM_门诊医嘱清单 Then
                    For i = 0 To UBound(arrTmp)
                        Call zlDatabase.ExecuteProcedure(CStr(arrTmp(i)), "合理用药监测")
                    Next
                End If
            End If
        Case 2    ' 2-药品提示
            If Val(.TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)) <> 0 Then
                '获取所选医嘱的药品编码
                strDrugCode = .TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)
                '调用药品提示接口
                gobjPass.getDrugExplain (strDrugCode)
            Else
                MsgBox "当前选中的医嘱不是按规格下达的药品医嘱。", vbInformation + vbOKOnly, "合理用药监测"
            End If
        Case 3    '3-在线医药知识库
            '调用在线医药知识库
            gobjPass.accessIFMI ("0")  '传入值固定为:"0",无返回值

        Case 4  '4-系统配置
            gobjPass.sysConfig

        Case 5    '5-获取警示详情
            gobjPass.getDrugAlertDetail

        End Select
    End With

    OutAdviceCheckWarn_TYT = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InAdviceCheckWarn_MK4(Optional ByVal bytShow As Byte = 0, Optional ByVal bytSubmit As Byte = 0, _
        Optional blnIsHaveOut As Boolean, Optional ByRef blnNoSave As Boolean, Optional ByRef rsOut As ADODB.Recordset, _
        Optional ByRef lngResult As Long = 1) As Long
'功能：调用Pass系统中对医嘱进行合理用药审查等相关功能
'参数：lngCmd=
'
'出参：
'      rsOut=禁忌药品说明
'      返回：blnIsHaveOut=是否存在离院带药的药品
'           lngResult-药师干预系统 0-不通过；1-通过
'返回：本次审核返回的最高级别警示值,为-1,-2,-3表示没有进行审查
'      检测PASS菜单时，返回>=0表示可以弹出菜单
'说明：用药审查：涉及当天下的临嘱(包括已执行)，和未停止的长嘱
'      用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim rs规格 As ADODB.Recordset
    Dim rs中药 As ADODB.Recordset

    Dim str医嘱ID As String, str相关ID As String, str医嘱期效 As String, str医嘱序号 As String, str医嘱状态 As String
    Dim str开嘱科室ID As String, str开嘱科室 As String, str开嘱科室IDTag As String
    Dim str药品名称 As String, str药品ID As String, str频率 As String, strPre用法 As String
    Dim str用法 As String, str用法ID As String, str离院带药 As String
    Dim str单次用量 As String, str单量单位 As String
    Dim str医生嘱托 As String, str滴速 As String
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim str开嘱医生 As String, str医生编码 As String, str开嘱医生Tag As String
    Dim str总量 As String, str总量单位 As String, str用药目的 As String, str嘱托 As String
    Dim str开嘱时间 As String, str结束时间 As String, str开始时间 As String
    Dim str中药组IDs As String, strGroupIDs As String
    Dim str执行科室ID As String
    Dim str医嘱IDs As String
    
    Dim str警示 As String
    Dim str警示值 As String
    Dim str状态 As String
    Dim str诊疗项目ID As String
    
    Dim lngMaxWarn As Long, strOld As String
    Dim strSQL As String
    Dim lngCount As Long, curDate As Date
    Dim arrLevel(0 To 4) As Long
    Dim arrLight(0 To 4) As String
    Dim strCurrentDate As String
    Dim i As Long, k As Long, j As Long, lng中药组ID As Long, lngLight As Long
    Dim lngBegin As Long, lngEnd As Long
    
    Dim blnOK As Boolean, blnDo As Boolean
    
    Dim strAdvicesIds As String
   
    
    Dim arrSQL As Variant
    Dim arrTmp As Variant
    
    lngMaxWarn = -1
    InAdviceCheckWarn_MK4 = lngMaxWarn

    On Error GoTo errH
    Screen.MousePointer = 11
    
    With gobjAdvice

        'PASS增加一条用药清单记录（多条重复调用）MDC_AddScreenDrug
        lngCount = 0
        curDate = zlDatabase.Currentdate
        '初始化药嘱信息
        Set rsAdvice = InitAdviceRS(FUN_医嘱信息)
        '批量提取规格
        For i = .FixedRows To .Rows - 1
            If InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) = 0 Then
                str诊疗项目ID = str诊疗项目ID & "," & .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
            End If
        Next
        If str诊疗项目ID <> "" Then
            Set rs规格 = GetDrugID(str诊疗项目ID) '一条记录也要求必须返回记录集合
        End If
                
        '启用了禁忌药品说明参数;场合为住院编辑;审查功能
        If glngModel = PM_住院编辑 And gbytReason = 1 Then
            Set rsOut = InitAdviceRS(FUN_输出内容)
        End If
        
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_住院编辑 Then
                '住院编辑界面加载医嘱时已经屏蔽掉作废医嘱及停止和确认停止的长嘱
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
            Else
                blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4")
                
                If blnDo Then
                    '一并给药，只在首行显示期效,其余行擦除（见vsAdvice_DrawCell）
                    '一并给药，期效取首行期效
                    If RowIn一并给药(i, lngBegin, lngEnd) Then
                        str医嘱期效 = .TextMatrix(lngBegin, gobjCOL.intCOL期效)
                    Else
                        str医嘱期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                    End If
                    '1-作废医嘱（7天内作废的）,
                    '2-当天未停用的长期医嘱(1-新开2-疑问3-校对5-已重整,6-已暂停,7-已启用;（8-停止,9-确认停止）只传停止日期大于当天日期 ),
                    '3-当天临时医嘱
                    str状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                    str结束时间 = Format(.TextMatrix(i, gobjCOL.intCOL终止时间), "yyyy-mm-dd")
                    blnDo = blnDo And (str状态 = "4" Or _
                        (str医嘱期效 = "长嘱" And (InStr(",8,9,", str状态) > 0 And str结束时间 > Format(curDate, "yyyy-MM-dd") Or InStr(",1,2,3,5,6,7,", str状态) > 0) Or _
                        str医嘱期效 = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")))
                End If
            End If
            
            If blnDo Then
                '获取中药医嘱组ID
                If glngModel = PM_住院医嘱清单 And (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                    str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    '取药品名称
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                        str药品名称 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                    Else
                        str药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                    End If
                    If glngModel = PM_住院编辑 Then
                        '判断是否是院外执行的药品
                        str离院带药 = ""
                        If Val(.TextMatrix(i, gobjCOL.intCOL执行性质)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))), gobjCOL.intCOL执行性质)) = 5 Then
                            blnIsHaveOut = True: str离院带药 = "离院带药"
                        End If
    
                         '取药品给药途径和中药用法
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str用法 = ""    '一并给药不重复取
                        If str用法 = "" Then
                            str滴速 = ""
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                            If k <> -1 Then
                                If .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                                    str用法 = .TextMatrix(k, gobjCOL.intCOL用法)
                                Else
                                    str用法 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                    If InStr(.TextMatrix(k, gobjCOL.intcol医嘱嘱托), "滴/分钟") > 0 Or InStr(.TextMatrix(k, gobjCOL.intcol医嘱嘱托), "毫升/小时") > 0 Then
                                        str滴速 = .TextMatrix(k, gobjCOL.intcol医嘱嘱托)
                                    End If
                                End If
                                str用法ID = .TextMatrix(k, gobjCOL.intCOL诊疗项目ID)
                            End If
                        End If
                    Else
                        '取药品给药途径和中药用法
                        If Trim(.TextMatrix(i, gobjCOL.intCOL用法)) = "" Then
                            str用法 = strPre用法
                        Else
                            str用法 = Split(.TextMatrix(i, gobjCOL.intCOL用法), ",")(0)
                        End If
                        
                        strPre用法 = str用法
                    End If
                    
                    '取用药频率(次/天),都为整数四舍五入
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str频率 = ""    '一并给药不重复取
                    If str频率 = "" Then
                        str频率 = .TextMatrix(i, gobjCOL.intCOL频率)
                        
                        str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                        If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                        
                        If str开嘱医生Tag <> str开嘱医生 And str开嘱医生 <> "" Then
                            str医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
                            str开嘱医生Tag = str开嘱医生
                        End If
                       
                    End If
                    
                    '长期医嘱按品种下达时,任意取一个药品Id
                    If Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) = 0 Then
                        rs规格.Filter = "药名ID =" & .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                        If Not rs规格.EOF Then str药品ID = rs规格!药品ID & ""
                    Else
                        str药品ID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                    End If
                    '开嘱科室名称
                    str开嘱科室ID = .TextMatrix(i, gobjCOL.intCOL开嘱科室ID)
                    If .TextMatrix(i, gobjCOL.intCOL开嘱科室ID) <> str开嘱科室IDTag And Val(str开嘱科室ID) <> 0 Then
                        str开嘱科室 = Sys.RowValue("部门表", Val(.TextMatrix(i, gobjCOL.intCOL开嘱科室ID)), "名称")
                        str开嘱科室IDTag = .TextMatrix(i, gobjCOL.intCOL开嘱科室ID)
                    End If
                    
                    If glngModel = PM_住院编辑 Then
                        str医嘱ID = .RowData(i)
                        str医嘱期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                        
                        str单次用量 = .TextMatrix(i, gobjCOL.intCOL单量)
                        str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                        
                        str总量 = .TextMatrix(i, gobjCOL.intCOL总量)
                        str总量单位 = .TextMatrix(i, gobjCOL.intcol总量单位)
                        str开始时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:mm:ss")
                        str结束时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL终止时间), "yyyy-MM-dd HH:mm:ss")
                        str开嘱时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:mm:ss")
                        str执行科室ID = .TextMatrix(i, gobjCOL.intCol执行科室ID)
                        If str医嘱期效 = "1" Then
                            str结束时间 = str开始时间
                        End If
                        
                        If Not rsOut Is Nothing Then
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                                '西药,中成药
                                rsOut.AddNew
                                rsOut!医嘱ID = CLng(str医嘱ID)
                                rsOut!禁忌药品说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                                rsOut!药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                                rsOut!状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                                rsOut.Update
                            ElseIf Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                            '中药配方
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                If k <> -1 Then
                                    rsOut.AddNew
                                    rsOut!医嘱ID = CLng(.RowData(k) & "")
                                    rsOut!禁忌药品说明 = .TextMatrix(k, gobjCOL.intCol禁忌药品说明)
                                    rsOut!药品名称 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                    rsOut!状态 = .TextMatrix(k, gobjCOL.intCOL状态)
                                    rsOut.Update
                                End If
                            End If
                        End If
                    Else
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                        str医嘱IDs = str医嘱IDs & "," & str医嘱ID
                        str单次用量 = Val(.TextMatrix(i, gobjCOL.intCOL单量))
                        str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量)
             
                        str单次用量 = FormatEx(str单次用量, 5)
                        str单量单位 = Replace(str单量单位, str单次用量, "")
                        
                        str总量 = Val(.TextMatrix(i, gobjCOL.intCOL总量))
                        str总量单位 = .TextMatrix(i, gobjCOL.intCOL总量)
                        str总量 = FormatEx(str总量, 5)
                        str总量单位 = Replace(str总量单位, str总量, "")
                        
                        
                        str开嘱时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:mm:ss")
                        str结束时间 = Format(.TextMatrix(i, gobjCOL.intCOL终止时间), "yyyy-MM-dd HH:mm:ss")
                        str开始时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:mm:ss")
                        str执行科室ID = ""
                        If str医嘱期效 = "临嘱" Then
                            str结束时间 = str开嘱时间
                        End If
                    End If
                    
                    If str结束时间 & "" = "" Then str结束时间 = " "
                    str医嘱期效 = IIf(str医嘱期效 = "临嘱", 1, 0)
                    str相关ID = .TextMatrix(i, gobjCOL.intCOL相关ID)
                    str医生嘱托 = .TextMatrix(i, gobjCOL.intcol医嘱嘱托)
                    str用药目的 = .TextMatrix(i, gobjCOL.intcol用药目的)
                    
                    If str用药目的 = "1" Then
                        str用药目的 = "3"
                    ElseIf str用药目的 = "2" Then
                        str用药目的 = "4"
                    Else
                        str用药目的 = "0"
                    End If
                    '
                    str医嘱状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                    '"0"-在用（默认）；"1"-已作废；"2"-已停嘱；"3"-离院带药（根据系统设置参与审查）
                    
                    If glngModel = PM_住院编辑 Then
                        blnOK = str离院带药 = "离院带药"
                    Else
                        blnOK = .TextMatrix(i, gobjCOL.intCOL执行性质) = "离院带药"
                        If InStr("," & strGroupIDs & ",", "," & str相关ID & ",") = 0 And InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                            strGroupIDs = strGroupIDs & "," & str相关ID
                        End If
                    End If
                    If blnOK Then
                        str医嘱状态 = "3"
                    ElseIf str医嘱状态 = "4" Then
                        str医嘱状态 = "1"
                    Else
                        str医嘱状态 = "0"
                        If glngModel = PM_住院编辑 Then
                            str医嘱状态 = IIf(InStr(",1,2,", "," & .TextMatrix(i, gobjCOL.intCOLEDIT) & ",") > 0, "9", "0")
                        End If
                    End If
                    '----------------------------------------------------------
                    rsAdvice.AddNew
                    rsAdvice!医嘱ID = str医嘱ID
                    rsAdvice!相关ID = str相关ID
                    rsAdvice!医嘱期效 = str医嘱期效
                    rsAdvice!医嘱序号 = lngCount + 1
                    rsAdvice!医嘱状态 = str医嘱状态
                    rsAdvice!开嘱科室 = str开嘱科室
                    rsAdvice!开嘱科室id = str开嘱科室ID
                    rsAdvice!开嘱医生编码 = str医生编码
                    rsAdvice!开嘱医生 = str开嘱医生
                    rsAdvice!药品ID = str药品ID
                    rsAdvice!药品名称 = str药品名称
                    rsAdvice!单次用量 = str单次用量
                    
                    rsAdvice!单量单位 = str单量单位
                    rsAdvice!频率 = str频率
                    rsAdvice!用法 = str用法
                    rsAdvice!用法ID = str用法ID
                    rsAdvice!开嘱时间 = str开嘱时间
                    rsAdvice!开始时间 = str开始时间
                    rsAdvice!结束时间 = str结束时间
                    
                    rsAdvice!总量 = str总量
                    rsAdvice!总量单位 = str总量单位
                    rsAdvice!用药目的 = str用药目的
                    rsAdvice!医生嘱托 = str医生嘱托
                    rsAdvice!滴速 = str滴速
                    rsAdvice!执行科室ID = str执行科室ID
                    rsAdvice.Update
                    '----------------------------------------------------------------------------
                    
                    lngCount = lngCount + 1
                End If
            End If
        Next
        strGroupIDs = Mid(strGroupIDs, 2)
        
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        
        If glngModel = PM_住院医嘱清单 Then
            '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
            str相关ID = ""
            If str中药组IDs <> "" Then
                Set rs中药 = Get中药配方(str中药组IDs)
                With rs中药
                    For i = 1 To .RecordCount
                        If !相关ID & "" <> str相关ID Then
                            str开嘱医生 = !开嘱医生
                            If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                            str开嘱医生 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名") & "/" & str开嘱医生
                            str开嘱科室 = Sys.RowValue("部门表", Val(!开嘱科室id & ""), "名称")
                            
                            str开嘱时间 = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")
                            str结束时间 = Format(!终止时间 & "", "yyyy-MM-dd HH:mm:ss")
                            str开始时间 = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")
                            If !医嘱期效 & "" = "1" Then
                                str结束时间 = str开嘱时间
                            End If
                            
                            If !用药目的 & "" = "1" Then
                                str用药目的 = "3"
                            ElseIf !用药目的 & "" = "2" Then
                                str用药目的 = "4"
                            Else
                                str用药目的 = "0"
                            End If
                            
                            If !组执行性质 & "" = "5" And !执行性质 <> "5" Then
                                str医嘱状态 = "3"
                            ElseIf !医嘱状态 & "" = "4" Then
                                str医嘱状态 = "1"
                            Else
                                str医嘱状态 = "0"
                            End If
                            str相关ID = !相关ID & ""
                        End If
                        '----------------------------------------------------------
                        rsAdvice.AddNew
                        rsAdvice!医嘱ID = !id
                        rsAdvice!相关ID = !相关ID & ""
                        rsAdvice!医嘱期效 = !医嘱期效 & ""
                        rsAdvice!医嘱序号 = lngCount + 1
                        rsAdvice!医嘱状态 = str医嘱状态
                        rsAdvice!开嘱科室 = str开嘱科室
                        rsAdvice!开嘱科室id = !开嘱科室id & ""
                        rsAdvice!开嘱医生编码 = str医生编码
                        rsAdvice!开嘱医生 = str开嘱医生
                        rsAdvice!药品ID = !药品ID & ""
                        rsAdvice!药品名称 = !药品名称 & ""
                        rsAdvice!单次用量 = !单次用量 & ""
                        
                        rsAdvice!单量单位 = !单量单位 & ""
                        rsAdvice!频率 = !频率 & ""
                        rsAdvice!用法 = !用法 & ""
                        rsAdvice!用法ID = !用法ID & ""
                        rsAdvice!开嘱时间 = str开嘱时间
                        rsAdvice!开始时间 = str开始时间
                        rsAdvice!结束时间 = str结束时间
                        
                        rsAdvice!总量 = !总给予量 & ""
                        rsAdvice!总量单位 = !总量单位 & ""
                        rsAdvice!用药目的 = str用药目的
                        rsAdvice!医生嘱托 = !医生嘱托 & ""
                        rsAdvice!执行科室ID = !执行科室ID & ""
                        rsAdvice.Update
                        '----------------------------------------------------------------------------
                        lngCount = lngCount + 1
                        .MoveNext
                    Next
                End With
            End If

            '从数据库提取作废的医嘱
            ' 只传人七天内作废
            If strAdvicesIds <> "" Then
                strAdvicesIds = strAdvicesIds & ","
            End If
            strSQL = "Select a.Id As 医嘱id, a.相关id, a.序号 As 医嘱序号, a.诊疗类别, a.医嘱期效, a.医嘱状态, a.诊疗项目id, NVL(a.收费细目id,f.药品ID) as 药品ID , Decode(a.诊疗类别||'','7',a.医嘱内容,a.标本部位) as 药品名称, a.执行频次 as 频率, a.单次用量, a.总给予量," & vbNewLine & _
                "       a.执行标记, a.开始执行时间, a.开嘱时间,a.开始执行时间 as 开始时间,a.执行终止时间 as 结束时间, a.医生嘱托, a.开嘱科室id, e.名称 As 开嘱科室, a.开嘱医生, a.用药目的, a.执行科室ID, b.计算单位 as 单量单位, c.住院单位 as 总量单位," & vbNewLine & _
                "       a.用药目的,a.医生嘱托,d.医嘱内容 As 用法, d.诊疗项目id As 用法id " & vbNewLine & _
                "From 病人医嘱记录 A, 诊疗项目目录 B, 药品规格 C, 病人医嘱记录 D, 部门表 E,药品规格 F " & vbNewLine & _
                "Where a.病人id = [1] And a.主页id = [2] And a.诊疗项目id = b.Id(+) And a.收费细目id = c.药品id(+) And a.诊疗项目ID = f.药名ID(+) And Nvl(a.相关id, 0) = d.Id(+) And" & vbNewLine & _
                "      a.开嘱科室id = e.Id(+) And a.诊疗类别 In ('5', '6', '7') And Nvl(a.执行标记, 0) <> -1 And" & vbNewLine & _
                "      (a.医嘱状态 = 4 And a.开嘱时间 Between Trunc(Sysdate) - 7 And Trunc(Sysdate + 1) Or" & vbNewLine & _
                "      (a.医嘱状态 In (8, 9) And Trunc(a.执行终止时间) > Trunc(Sysdate))) And Not Instr([3], ',' || a.Id || ',') > 0"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng病人ID, gobjPati.lng主页ID, strAdvicesIds)
            For i = 1 To rsTmp.RecordCount
                '----------------------------------------------------------
                str医嘱ID = rsTmp!医嘱ID & ""
                str相关ID = rsTmp!相关ID & ""
                str医嘱期效 = rsTmp!医嘱期效 & ""
                str医嘱序号 = rsTmp!医嘱序号 & ""
                str医嘱状态 = rsTmp!医嘱状态 & ""
                str开嘱科室 = rsTmp!开嘱科室 & ""
                str开嘱科室ID = rsTmp!开嘱科室id & ""
                str开嘱医生 = rsTmp!开嘱医生 & ""
                str药品ID = rsTmp!药品ID & ""
                str药品名称 = rsTmp!药品名称 & ""
                str单次用量 = rsTmp!单次用量 & ""
                
                str单量单位 = rsTmp!单量单位 & ""
                str频率 = rsTmp!频率 & ""
                str用法 = rsTmp!用法 & ""
                str用法ID = rsTmp!用法ID & ""
                str开嘱时间 = Format(rsTmp!开始时间 & "", "yyyy-mm-dd HH:MM:ss")
                str开始时间 = Format(rsTmp!开始时间 & "", "yyyy-mm-dd HH:MM:ss")
                str结束时间 = Format(rsTmp!结束时间 & "", "yyyy-mm-dd HH:MM:ss")
                
                str总量 = rsTmp!总给予量 & ""
                str总量单位 = rsTmp!总量单位 & ""
                str用药目的 = rsTmp!用药目的 & ""
                str医生嘱托 = rsTmp!医生嘱托 & ""
                str执行科室ID = rsTmp!执行科室ID & ""
                '-------------------------
                
                If str开嘱医生Tag <> str开嘱医生 And str开嘱医生 <> "" Then
                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                    str医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
                    str开嘱医生Tag = str开嘱医生
                End If
                
                If str医嘱期效 = "1" Then
                    str结束时间 = str开始时间
                End If
                
                If str结束时间 & "" = "" Then str结束时间 = " "
                '"0"-在用（默认）；"1"-已作废；"2"-已停嘱；"3"-离院带药（根据系统设置参与审查）
                If str医嘱状态 = "4" Then
                    str医嘱状态 = "1"
                ElseIf str医嘱状态 = "8" Or str医嘱状态 = "9" Then
                    str医嘱状态 = "2"
                Else
                    str医嘱状态 = "0"
                End If
                
                If str用药目的 = "1" Then
                    str用药目的 = "3"
                ElseIf str用药目的 = "2" Then
                    str用药目的 = "4"
                Else
                    str用药目的 = "0"
                End If
                
                '----------------------------------------------------------
                rsAdvice.AddNew
                rsAdvice!医嘱ID = str医嘱ID
                rsAdvice!相关ID = str相关ID
                rsAdvice!医嘱期效 = str医嘱期效
                rsAdvice!医嘱序号 = str医嘱序号
                rsAdvice!医嘱状态 = str医嘱状态
                rsAdvice!开嘱科室 = str开嘱科室
                rsAdvice!开嘱科室id = str开嘱科室ID
                rsAdvice!开嘱医生编码 = str医生编码
                rsAdvice!开嘱医生 = str开嘱医生
                rsAdvice!药品ID = str药品ID
                rsAdvice!药品名称 = str药品名称
                rsAdvice!单次用量 = str单次用量
                
                rsAdvice!单量单位 = str单量单位
                rsAdvice!频率 = str频率
                rsAdvice!用法 = str用法
                rsAdvice!用法ID = str用法ID
                rsAdvice!开嘱时间 = str开嘱时间
                rsAdvice!开始时间 = str开始时间
                rsAdvice!结束时间 = str结束时间
                
                rsAdvice!总量 = str总量
                rsAdvice!总量单位 = str总量单位
                rsAdvice!用药目的 = str用药目的
                rsAdvice!医生嘱托 = str医生嘱托
                rsAdvice!执行科室ID = str执行科室ID
                rsAdvice.Update
                lngCount = lngCount + 1
                '-------------------------
                rsTmp.MoveNext
            Next
            
            If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
            '取执行科室ID
            If str医嘱IDs <> "" Then
                Set rsTmp = GetDrugInfo_MK4("", str医嘱IDs, gobjPati.lng病人ID, gobjPati.lng主页ID)
                rsAdvice.Filter = ""
                For i = 1 To rsAdvice.RecordCount
                    rsTmp.Filter = "ID=" & rsAdvice!医嘱ID
                    If Not rsTmp.EOF Then
                        rsAdvice!执行科室ID = rsTmp!执行科室ID & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
            '获取滴速
            If strGroupIDs <> "" Then
                Set rsTmp = Get滴速("," & strGroupIDs & ",")
                For i = 1 To rsTmp.RecordCount
                    rsAdvice.Filter = "相关ID =" & rsTmp!id
                    Do While Not rsAdvice.EOF
                        rsAdvice!滴速 = rsTmp!医生嘱托 & ""
                        rsAdvice.MoveNext
                    Loop
                    rsTmp.MoveNext
                Next
                rsAdvice.Filter = ""
            End If
        End If
        '无可审查的药品l
        If lngCount = 0 Then
            Screen.MousePointer = 0: Exit Function
        End If
        
        'PASS审查函数MDC_DoCheck
        Call AdviceCheckWarn_MK4(gobjPati.lng病人ID, "", gobjPati.lng主页ID, bytShow, bytSubmit, rsAdvice, str警示, lngResult)
        
        arrSQL = Array()
        '获取医嘱审查结果,并填写警示灯
        '-------------------------------------------------------------
        '返回值顺：0-蓝灯,1-黑灯,2-红灯,3-橙灯,4-黄灯
        '警示级顺：0-蓝灯,4-黄灯,3-橙灯,2-红灯,1-黑灯(因为PASS升级的原因)
        arrLevel(0) = 0: arrLevel(1) = 4: arrLevel(2) = 3: arrLevel(3) = 2: arrLevel(4) = 1
        arrLight(0) = "蓝_4": arrLight(1) = "黑_4": arrLight(2) = "红_4": arrLight(3) = "橙_4": arrLight(4) = "黄_4"
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_住院编辑 Then
                '住院编辑界面加载医嘱时已经屏蔽掉作废医嘱及停止和确认停止的长嘱
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
            Else
                blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4")
                
                If blnDo Then
                    '一并给药，只在首行显示期效,其余行擦除（见vsAdvice_DrawCell）
                    '一并给药，期效取首行期效
                    If RowIn一并给药(i, lngBegin, lngEnd) Then
                        str医嘱期效 = .TextMatrix(lngBegin, gobjCOL.intCOL期效)
                    Else
                        str医嘱期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                    End If
                    '1-作废医嘱（7天内作废的）,
                    '2-当天未停用的长期医嘱(1-新开2-疑问3-校对5-已重整,6-已暂停,7-已启用;（8-停止,9-确认停止）只传停止日期大于当天日期 ),
                    '3-当天临时医嘱
                    str状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                    str结束时间 = Format(.TextMatrix(i, gobjCOL.intCOL终止时间), "yyyy-mm-dd")
                    blnDo = blnDo And (str状态 = "4" Or _
                        (str医嘱期效 = "长嘱" And (InStr(",8,9,", str状态) > 0 And str结束时间 > Format(curDate, "yyyy-MM-dd") Or InStr(",1,2,3,5,6,7,", str状态) > 0) Or _
                        str医嘱期效 = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")))
                End If
            End If
            If blnDo Then
                If glngModel = PM_住院编辑 Then
                    str医嘱ID = .RowData(i) & ""
                Else
                    str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID) & ""
                End If
                rsAdvice.Filter = "医嘱ID='" & str医嘱ID & "'"
                
                If rsAdvice.RecordCount > 0 Then
                    k = CLng(rsAdvice!警示 & "")
                Else
                    k = -1
                End If
                
                If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                    '西药、西成药'设置警示灯
                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                    If k >= 0 And k <= 4 Then
                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(k)
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                    Else
                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                    End If
                    
                    If glngModel = PM_住院编辑 Then
                        '标记审查结果变化,以备更新数据库
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                            .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                            blnNoSave = True    '标记为未保存
                        End If
                        
                        If Not rsOut Is Nothing And k = 1 Then
                            rsOut.Filter = "医嘱ID=" & CLng(str医嘱ID) & " And 状态 < 3 "
                            If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                        End If
                    ElseIf PM_住院医嘱清单 = glngModel Then
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                        End If
                    End If
                ElseIf .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                    '中药配方
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                        lng中药组ID = .TextMatrix(i, gobjCOL.intCOL相关ID)          '中药配方组ID
                        lngLight = -1 '初始化
                    End If
                    '设置警示灯 取草药中最大警示值
                    If k >= 0 Then
                        If lngLight >= 0 Then
                            If arrLevel(k) > arrLevel(lngLight) Then
                                lngLight = k
                            End If
                        Else
                            lngLight = k
                        End If
                    End If
                End If
    
                 '记录最高级别警示值
                If k >= 0 Then
                    If lngMaxWarn >= 0 Then
                        If arrLevel(k) > arrLevel(lngMaxWarn) Then
                            lngMaxWarn = k
                        End If
                    Else
                        lngMaxWarn = k
                    End If
                End If
            Else
                If glngModel = PM_住院编辑 Then
                    If .RowData(i) = lng中药组ID And .RowData(i) <> 0 Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        '设置警示灯
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        
                        If glngModel = PM_住院编辑 Then
                            '标记审查结果变化,以备更新数据库
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                blnNoSave = True    '标记为未保存
                            End If
                            
                            If Not rsOut Is Nothing And lngLight = 1 Then
                                rsOut.Filter = "医嘱ID=" & lng中药组ID & " And 状态 < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                            End If
                        End If
                        lng中药组ID = 0
                        lngLight = -1
                    End If
                End If
            
            End If
        Next
        '医嘱清单中药配方警示灯处理
        If glngModel = PM_住院医嘱清单 And Not rs中药 Is Nothing Then
            For i = .FixedRows To .Rows - 1
                '中药服法
                If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                    lngLight = -1
                    str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                    rs中药.Filter = "相关ID=" & str医嘱ID
                    
                    For j = 1 To rs中药.RecordCount
                        rsAdvice.Filter = "医嘱ID='" & rs中药!id & "'"
                        If rsAdvice.RecordCount > 0 Then
                            k = CLng(rsAdvice!警示 & "")
                        Else
                            k = -1
                        End If
                        '设置警示灯 取草药中最大警示值
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If arrLevel(k) > arrLevel(lngLight) Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                        rs中药.MoveNext
                    Next
                    
                    '设置警示灯
                    If lngLight >= 0 And lngLight <= 4 Then
                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                    Else
                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                    End If
                    
                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(lngLight >= 0 And lngLight <= 4, lngLight, "NULL") & ")"
                    End If
                    
                    '记录最高级别警示值
                    If lngLight >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(lngLight) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = lngLight
                            End If
                        Else
                            lngMaxWarn = lngLight
                        End If
                    End If
                End If
            Next
        End If
            
'        '对于界面上找不到的医嘱,通过SQL强制刷新
'        If strAdvicesIds <> "" Then
'            strAdvicesIds = Mid(strAdvicesIds, 2)
'            arrTmp = Split(strAdvicesIds, ",")
'            For i = LBound(arrTmp) To UBound(arrTmp)
'                str医嘱ID = Split(arrTmp(i), ":")(0)
'                str警示值 = Split(arrTmp(i), ":")(1)
'                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
'            Next
'        End If
'
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
        Next
            
    End With

    '返回审查结果
    InAdviceCheckWarn_MK4 = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function InAdviceCheckWarn_MK(ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional blnIsHaveOut As Boolean, Optional ByRef blnNoSave As Boolean, _
    Optional ByVal bytFunc As Byte = 0, Optional ByRef rsOut As ADODB.Recordset) As Long
'功能：调用Pass系统中对医嘱进行合理用药审查等相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        1/33-保存自动审查(住院/门诊),2/34-提交自动审查(住院/门诊),3-手工调用审查
'        6-单药警告,12-用药研究,22-病生状态/过敏史管理(编辑)
'      lngRow=当前药品医嘱的行号，lngCmd=0,6时需要
'      返回：blnIsHaveOut=是否存在离院带药的药品
'出参:
'   rsOut=禁忌说明
'返回：本次审核返回的最高级别警示值,为-1,-2,-3表示没有进行审查
'      检测PASS菜单时，返回>=0表示可以弹出菜单
'说明：用药审查：涉及当天下的临嘱(包括已执行)，和未停止的长嘱
'      用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset
    Dim rs中药 As ADODB.Recordset
    Dim str药品 As String, str用法 As String, str频率 As String, strPre用法 As String, str期效 As String
    Dim str药品ID As String, str用法ID As String, strTmp As String, strType As String
    Dim str中药组IDs As String
    Dim str相关ID As String
    Dim lngMaxWarn As Long, strOld As String, lng中药组ID As Long
    Dim strSQL As String, blnDo As Boolean, blnLight As Boolean
    Dim lngCount As Long, curDate As Date
    Dim arrLevel(0 To 4) As Long
    Dim arrLight(0 To 4) As String
    Dim strCurrentDate As String
    Dim i As Long, k As Long, j As Long, lngLight As Long
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim str开嘱医生 As String, str医生编码 As String, str开嘱科室 As String, str医嘱ID As String
    Dim str单量 As String, str单量单位 As String
    Dim str住院总量 As String, str住院单位 As String, str用药目的 As String, str嘱托 As String
    Dim str开嘱时间 As String, str终止时间 As String, str执行时间 As String
    Dim lngBegin As Long, lngEnd As Long, lngGroupMax As Long
    Dim rsAdvice As ADODB.Recordset
    Dim rs规格 As ADODB.Recordset
    Dim strAdvicesIds As String, strAll As String, strFaceID As String
    Dim str执行时间方案 As String
    Dim str停止点 As String
    
    Dim arrSQL As Variant
    Dim arrTmp As Variant
    
    lngMaxWarn = -1
    InAdviceCheckWarn_MK = lngMaxWarn

    On Error GoTo errH
    Screen.MousePointer = 11
    
    '美康3.0
    '检验PASS可用状态
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '114036同一个病人多次审查时病人信息每次都要传入
    '-------------------------------------------------------------
    strSQL = _
    " Select Nvl(B.姓名,A.姓名) 姓名,Nvl(B.性别,A.性别) 性别,A.出生日期,B.身高,B.体重,B.入院日期,B.出院日期," & _
             " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
             " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
             " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
             " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng病人ID, gobjPati.lng主页ID)
    If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

    Call PassSetPatientInfo(gobjPati.lng病人ID, gobjPati.lng主页ID, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), rsTmp!体重 & "", rsTmp!身高 & "", _
                            rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), _
                            IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))

    '传人病人过敏史
    '-------------------------------------------------------
    Set rsTmp = Get病人过敏记录(gobjPati.lng病人ID, gobjPati.lng主页ID)

    For i = 1 To rsTmp.RecordCount
        Call PassSetAllergenInfo(i, rsTmp!药物ID & "", rsTmp!药物名 & "", "DrugName", "")
        rsTmp.MoveNext
    Next

    '传人病生状态
    '------------------------------------------------------------------
    Set rsTmp = Get病人诊断记录(gobjPati.lng病人ID, gobjPati.lng主页ID, "2,12")
    strCurrentDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")

    For i = 1 To rsTmp.RecordCount
        Call PassSetMedCond(i & "", rsTmp!编码 & "", rsTmp!名称 & "", "User", strCurrentDate, strCurrentDate)
        rsTmp.MoveNext
    Next

    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With gobjAdvice
            If IIf(glngModel = PM_住院编辑, .RowData(lngRow) <> 0, True) And InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) > 0 Then
                '取药品名称
                If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) > 0 Then
                    str药品 = .TextMatrix(lngRow, gobjCOL.intCOL药品名称)
                Else
                    str药品 = .TextMatrix(lngRow, gobjCOL.intCOL医嘱内容) '中药名称
                End If
                
                '取药品给药途径(当前可见行不会是中草药)
                If glngModel = PM_住院编辑 Then
                    str用法 = ""
                    k = .FindRow(CLng(.TextMatrix(lngRow, gobjCOL.intCOL相关ID)), lngRow + 1)
                    If k <> -1 Then str用法 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                Else
                    str用法 = .TextMatrix(lngRow, gobjCOL.intCOL用法)
                    If InStr(str用法, ",") > 0 Then str用法 = Left(str用法, InStr(str用法, ",") - 1)
                End If
                
                '药品长期医嘱按品种下达时,收费细目ID为空,传任意药品ID
                If Val(.TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)) = 0 Then
                    str药品ID = GetDrugID(.TextMatrix(lngRow, gobjCOL.intCOL诊疗项目ID))
                Else
                    str药品ID = .TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)
                End If
                
                '传入查询药品信息
                Call PassSetQueryDrug(str药品ID, str药品, .TextMatrix(lngRow, gobjCOL.intCOL单量单位), str用法)
                
                '设置菜单可用状态，在zlPASSPopupCommandBars中设置
                InAdviceCheckWarn_MK = 1    '表示可以弹出菜单
            ElseIf glngModel = PM_住院医嘱清单 And .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(lngRow, gobjCOL.intCol操作类型) = "4" Then
                 InAdviceCheckWarn_MK = 1    '表示可以弹出菜单
            End If
        End With
        Screen.MousePointer = 0: Exit Function
    End If
    If glngModel = PM_住院编辑 Then
        '过敏史/病生状态编辑
        '-------------------------------------------------------------
        If lngCmd = 22 Then
            'lngCmd=21-只读,22-非强制编辑,23-强制编辑
            If PassDoCommand(lngCmd) = 2 Then
                '如果返回值为2表示"过敏史/病生状态编辑"管理发生变化，需要重新自动审查
                lngCmd = 2    '转为自动调用审查,继续执行
            Else
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    '启用了禁忌药品说明参数  且场合为住院编辑审查功能
    If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And glngModel = PM_住院编辑 And gbytReason = 1 Then
        Set rsOut = InitAdviceRS(FUN_输出内容)
    End If
    
    '传入病人医嘱信息
    '-------------------------------------------------------------
    With gobjAdvice
        If lngCmd = 6 Then
            If glngModel = PM_住院编辑 Then
                strTmp = .RowData(lngRow)
            Else
                strTmp = .TextMatrix(lngRow, gobjCOL.intCOLID)
            End If
            Call PassSetWarnDrug(strTmp)    '单药警告(已警告的医嘱唯一码)
        Else
            '用药审核或用药研究
            lngCount = 0
            curDate = zlDatabase.Currentdate
            str药品 = "": str用法 = "": str频率 = ""
            '提起获取任意的诊疗药品ID
            For i = .FixedRows To .Rows - 1
                If InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) = 0 Then
                    str药品 = str药品 & "," & .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                End If
            Next
            If str药品 <> "" Then
                Set rs规格 = GetDrugID(str药品) '一条记录也要求必须返回记录集合
                str药品 = ""
            End If
            
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_住院编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                    blnDo = blnDo And (lngCmd = 12 Or .TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" _
                            Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                Else
                    blnDo = InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4")
                    If blnDo Then
                        '一并给药，只在首行显示期效,其余行擦除（见vsAdvice_DrawCell）
                        '一并给药，期效取首行期效
                        If RowIn一并给药(i, lngBegin, lngEnd) Then
                            str期效 = .TextMatrix(lngBegin, gobjCOL.intCOL期效)
                        Else
                            str期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                        End If
                        '已停止的长嘱也要传入
                        blnDo = (lngCmd = 12 Or .TextMatrix(i, gobjCOL.intCOL状态) <> "4" And _
                                 (str期效 = "长嘱" Or str期效 = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")))
                        '不含已作废的医嘱,含当天的临嘱
                    End If
                End If
                
                If blnDo Then
                    If glngModel = PM_住院医嘱清单 And (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                        '获取中药医嘱组ID
                        str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                    Else
                        '取药品名称
                        If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 1 Then
                            str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                        Else
                            str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                        End If
                        
                        If glngModel = PM_住院编辑 Then
                            '判断是否是院外执行的药品
                            If Val(.TextMatrix(i, gobjCOL.intCOL执行性质)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))), gobjCOL.intCOL执行性质)) = 5 Then
                                blnIsHaveOut = True
                            End If
    
                            '取药品给药途径和中药用法
                            If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str用法 = ""    '一并给药不重复取
                            If str用法 = "" Then
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                If k <> -1 Then
                                    If .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                                        str用法 = .TextMatrix(k, gobjCOL.intCOL用法)
                                    Else
                                        str用法 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                    End If
                                End If
                            End If
                        Else
                            If Trim(.TextMatrix(i, gobjCOL.intCOL用法)) = "" Then
                                str用法 = strPre用法
                            Else
                                str用法 = Split(.TextMatrix(i, gobjCOL.intCOL用法), ",")(0)
                            End If
                            strPre用法 = str用法
                        End If
                        '取用药频率(次/天),都为整数四舍五入
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str频率 = ""    '一并给药不重复取
                        If str频率 = "" Then
                            If glngModel = PM_住院编辑 Then
                                str频率 = GetFrequency(.TextMatrix(i, gobjCOL.intCOL间隔单位), .TextMatrix(i, gobjCOL.intCOL频率次数), .TextMatrix(i, gobjCOL.intCOL频率间隔))
                            Else
                                Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), int频率次数, int频率间隔, str间隔单位, IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), "")
    
                                str频率 = GetFrequency(str间隔单位, int频率次数 & "", int频率间隔 & "")
                            End If
                            str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                            If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                            str开嘱医生 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名") & "/" & str开嘱医生
                        End If
                        '长期医嘱按品种下达时,任意传一个药品ID
                        If Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) = 0 Then
                            rs规格.Filter = "药名ID =" & .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                            If Not rs规格.EOF Then str药品ID = rs规格!药品ID & ""
                        Else
                            str药品ID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                        End If
                        '传入医嘱信息
                        If glngModel = PM_住院编辑 Then
                            str医嘱ID = CStr(.RowData(i))
                        Else
                            str医嘱ID = CStr(.TextMatrix(i, gobjCOL.intCOLID))
                        End If
                        '单量，单量单位
                        str单量 = .TextMatrix(i, gobjCOL.intCOL单量)
                        str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                        str单量 = Replace(str单量, str单量单位, "")
                        
                        Call PassSetRecipeInfo(str医嘱ID, str药品ID, str药品, _
                                             str单量, str单量单位, str频率, _
                                              Format(IIf(glngModel = PM_住院编辑, .Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), .TextMatrix(i, gobjCOL.intCOL开嘱时间)), "yyyy-MM-dd"), _
                                              Format(IIf(glngModel = PM_住院编辑, .Cell(flexcpData, i, gobjCOL.intCOL终止时间), .TextMatrix(i, gobjCOL.intCOL终止时间)), "yyyy-MM-dd"), _
                                              str用法, .TextMatrix(i, gobjCOL.intCOL相关ID), IIf(glngModel = PM_住院编辑, IIf(.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱", 0, 1), IIf(str期效 = "长嘱", 0, 1)), str开嘱医生)
                        If Not rsOut Is Nothing Then
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                                '西药,中成药
                                rsOut.AddNew
                                rsOut!医嘱ID = CLng(str医嘱ID)
                                rsOut!禁忌药品说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                                rsOut!药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                                rsOut!状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                                rsOut.Update
                            ElseIf Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                            '中药配方
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                If k <> -1 Then
                                    rsOut.AddNew
                                    rsOut!医嘱ID = CLng(.RowData(k) & "")
                                    rsOut!禁忌药品说明 = .TextMatrix(k, gobjCOL.intCol禁忌药品说明)
                                    rsOut!药品名称 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                    rsOut!状态 = .TextMatrix(k, gobjCOL.intCOL状态)
                                    rsOut.Update
                                End If
                            End If
                        End If
                        
                        lngCount = lngCount + 1
                    End If
                End If
            Next
            '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
            If glngModel = PM_住院医嘱清单 Then
                If str中药组IDs <> "" Then
                    Set rs中药 = Get中药配方(str中药组IDs)
                    With rs中药
                        For i = 1 To .RecordCount
                            If !相关ID & "" <> str相关ID Then
                                str开嘱医生 = !开嘱医生
                                If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                                str开嘱医生 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名") & "/" & str开嘱医生
                                str频率 = GetFrequency(!间隔单位 & "", !频率次数 & "", !频率间隔 & "")
                                str相关ID = !相关ID & ""
                            End If
                            Call PassSetRecipeInfo(!id, !药品ID & "", !药品名称 & "", !单次用量 & "", !单量单位 & "", str频率, Format(!开嘱时间 & "", "yyyy-MM-dd"), _
                            Format(!停嘱时间 & "", "yyyy-MM-dd"), !用法 & "", !相关ID & "", IIf(!医嘱期效 & "" = "0", "0", "1"), str开嘱医生)
                            
                            lngCount = lngCount + 1
                            .MoveNext
                        Next
                    End With
                End If
            End If
            '无可审查的药品
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End With

    '执行相应的命令
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)

    '获取医嘱审查结果,并填写警示灯
    '-------------------------------------------------------------
    If lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3 Then
        arrSQL = Array()
        '返回值顺：0-蓝灯,1-黄灯,2-红灯,3-黑灯,4-橙灯
        '警示级顺：0-蓝灯,1-黄灯,4-橙灯,2-红灯,3-黑灯(因为PASS升级的原因)
        arrLevel(0) = 0: arrLevel(1) = 1: arrLevel(2) = 3: arrLevel(3) = 4: arrLevel(4) = 2
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_住院编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 1 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 _
                            And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                    blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" _
                            Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 1
                
                    If blnDo Then
                        '一并给药，只在首行显示期效,其余行擦除（vsAdvice_DrawCell）
                        '一并给药，期效取首行期效
                        If RowIn一并给药(i, lngBegin, lngEnd) Then
                            str期效 = .TextMatrix(lngBegin, gobjCOL.intCOL期效)
                        Else
                            str期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                        End If
                        '已停止的长嘱也要传入'不含已作废的医嘱,含当天的临嘱
                        blnDo = .TextMatrix(i, gobjCOL.intCOL状态) <> "4" And (str期效 = "长嘱" _
                               Or str期效 = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                    End If
                End If
                
                If blnDo Then
                    If glngModel = PM_住院编辑 Then
                        str医嘱ID = .RowData(i) & ""
                    Else
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID) & ""
                    End If

                    k = PassGetWarn(str医嘱ID)
                    
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 1 Then
                        '西药、西成药'设置警示灯
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        If k >= 0 And k <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(k)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(k + 1).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        
                        If glngModel = PM_住院编辑 Then
                            '标记审查结果变化,以备更新数据库
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                blnNoSave = True    '标记为未保存
                            End If
                            
                            If Not rsOut Is Nothing And k = 3 Then
                                rsOut.Filter = "医嘱ID=" & CLng(str医嘱ID) & " And 状态 < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                            End If
                        ElseIf glngModel = PM_住院医嘱清单 Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                            End If
                        End If
                        
                    Else
                        '中药配方
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                            lng中药组ID = .TextMatrix(i, gobjCOL.intCOL相关ID)          '中药配方组ID
                            lngLight = -1 '初始化
                        End If
                        '设置警示灯 取草药中最大警示值
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If arrLevel(k) > arrLevel(lngLight) Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                    End If
                    '记录最高级别警示值
                    If k >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(k) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = k
                            End If
                        Else
                            lngMaxWarn = k
                        End If
                    End If
                Else
                    If glngModel = PM_住院编辑 Then
                        If .RowData(i) = lng中药组ID And .RowData(i) <> 0 Then
                            strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                            '设置警示灯
                            If lngLight >= 0 And lngLight <= 4 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(lngLight + 1).Picture
                            Else
                                .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                            End If
                            
                            If glngModel = PM_住院编辑 Then
                                '标记审查结果变化,以备更新数据库
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                    blnNoSave = True    '标记为未保存
                                End If
                                
                                If Not rsOut Is Nothing And lngLight = 3 Then
                                    rsOut.Filter = "医嘱ID=" & lng中药组ID & " And 状态 < 3 "
                                    If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                                End If
                            End If

                            lng中药组ID = 0
                            lngLight = -1
                        End If
                        
                    End If
                End If
            Next
            '医嘱清单中药配方警示灯处理
            If glngModel = PM_住院医嘱清单 And Not rs中药 Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    '中药服法
                    If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        lngLight = -1
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                        rs中药.Filter = "相关ID=" & str医嘱ID
                        
                        For j = 1 To rs中药.RecordCount
                            k = PassGetWarn(rs中药!id & "")
                            '设置警示灯 取草药中最大警示值
                            If k >= 0 Then
                                If lngLight >= 0 Then
                                    If arrLevel(k) > arrLevel(lngLight) Then
                                        lngLight = k
                                    End If
                                Else
                                    lngLight = k
                                End If
                            End If
                            rs中药.MoveNext
                        Next
                        
                        '设置警示灯
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(lngLight + 1).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(lngLight >= 0 And lngLight <= 4, lngLight, "NULL") & ")"
                        End If
                        
                        '记录最高级别警示值
                        If lngLight >= 0 Then
                            If lngMaxWarn >= 0 Then
                                If arrLevel(lngLight) > arrLevel(lngMaxWarn) Then
                                    lngMaxWarn = lngLight
                                End If
                            Else
                                lngMaxWarn = lngLight
                            End If
                        End If
                    End If
                Next
            End If
        End With
        
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
        Next
    End If
    '返回审查结果
    InAdviceCheckWarn_MK = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InAdviceCheckWarn_DT() As Boolean
'功能：调用大通用药监测系统对医嘱进行合理用药审查等相关功能
    Dim xmlbase As dt_base, xmlpre As dt_Pres
    Dim strTmp As String, arrTmp As Variant, curDate As Date
    Dim rsTmp As Recordset
    Dim i As Long, k As Long, blnDo As Boolean
    Dim str药品 As String, str给药途径 As String, str频率编码 As String, strXML As String
    Dim rsPati As ADODB.Recordset
    Dim strRetXML As String
    Dim blnIsHaveOut As Boolean '判断是否存在院外执行的药品

    Set rsPati = GetPatiInfo(gobjPati.lng病人ID, gobjPati.lng主页ID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    curDate = zlDatabase.Currentdate
    With xmlbase
        .dDoctCode = UserInfo.用户名
        .dDoctName = UserInfo.姓名
        .dDoctType = UserInfo.专业技术职务
        .dDeptCode = UserInfo.部门ID
        .dDeptName = UserInfo.部门名
        .dInHosCode = rsPati!住院号 & ""
        .dBedNo = "" & rsPati!当前床号
        .mPresDate = curDate
        .pCaseID = gobjPati.lng病人ID
        .pWeight = ""
        .pHeight = ""
        .pBirthday = NVL(rsPati!出生日期, vbNull)
        .pPatiName = rsPati!姓名
        .pSex = rsPati!性别
        .pStatms = ""
        .pEffect = ""
        .pBloodPress = ""
        .pLiverClean = ""
        
        '* 过敏源
        .pCaseCode1 = ""
        .pCaseName1 = ""
        .pCaseCode2 = ""
        .pCaseName2 = ""
        .pCaseCode3 = ""
        .pCaseName3 = ""
        Set rsTmp = Get病人过敏记录(gobjPati.lng病人ID, gobjPati.lng主页ID)
        If rsTmp.RecordCount > 0 Then
            .pCaseCode1 = "" & rsTmp!药物ID
            .pCaseName1 = rsTmp!药物名
            rsTmp.MoveNext
            
            If Not rsTmp.EOF Then
                .pCaseCode2 = "" & rsTmp!药物ID
                .pCaseName2 = rsTmp!药物名
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pCaseCode3 = "" & rsTmp!药物ID
                    .pCaseName3 = rsTmp!药物名
                End If
            End If
        End If
        
        '* 诊断信息
        .pDiagnose1 = ""
        .pDiagnose2 = ""
        .pDiagnose3 = ""
        .pDiagnoseName1 = ""
        .pDiagnoseName2 = ""
        .pDiagnoseName3 = ""
        Set rsTmp = Get病人诊断记录(gobjPati.lng病人ID, gobjPati.lng主页ID, "2")
        If rsTmp.RecordCount > 0 Then
            .pDiagnose1 = "" & rsTmp!编码
            .pDiagnoseName1 = "" & rsTmp!名称
            rsTmp.MoveNext
            If Not rsTmp.EOF Then
                .pDiagnose2 = "" & rsTmp!编码
                .pDiagnoseName2 = "" & rsTmp!名称
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pDiagnose3 = "" & rsTmp!编码
                    .pDiagnoseName3 = "" & rsTmp!名称
                End If
            End If
        End If
        
        '* 病生理状态
        .pBsl1 = ""
        .pBsl2 = ""
        .pBsl3 = ""
        strTmp = Get病人病生理情况(gobjPati.lng病人ID, gobjPati.lng主页ID)
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            .pBsl1 = arrTmp(0)
            If UBound(arrTmp) > 0 Then .pBsl2 = arrTmp(1)
            If UBound(arrTmp) > 1 Then .pBsl3 = arrTmp(2)
        End If
    End With
        
    arrTmp = Array()
    With gobjAdvice
        For i = .FixedRows To .Rows - 1
           If glngModel = PM_住院编辑 Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" _
                        Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
            ElseIf glngModel = PM_住院医嘱清单 Then
                blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                '不含已作废的医嘱,停止和确认停止的长嘱;包含当天的临嘱
                If blnDo Then
                    blnDo = .TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0 _
                            Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                            And Val(.TextMatrix(i, gobjCOL.intCOL状态)) <> 4
                End If
            End If
            
            If blnDo Then
                '取药品名称
                If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                    str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                Else
                    str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                End If

                '取药品给药途径
                If glngModel = PM_住院编辑 Then
                    '判断是否是院外执行的药品
                    If Val(.TextMatrix(i, gobjCOL.intCOL执行性质)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))), gobjCOL.intCOL执行性质)) = 5 Then
                        blnIsHaveOut = True
                    End If
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str给药途径 = "" '一并给药不重复取
                    If str给药途径 = "" Then
                        k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                        If k <> -1 Then str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                    End If
                ElseIf glngModel = PM_住院医嘱清单 Then
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then  '一并给药不重复取
                        str给药途径 = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOL相关ID)), "诊疗项目ID")  '传代码
                    End If
                End If
                Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), 0, 0, "", IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), str频率编码)
            
                xmlpre.PresID = gobjPati.lng病人ID  '没有医嘱ID传病人ID
                xmlpre.PresType = IIf(.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱", "L", "T")
                xmlpre.GeneralName = StrToXML(Sys.RowValue("诊疗项目目录", Val(.TextMatrix(i, gobjCOL.intCOL诊疗项目ID)), "名称"))
                xmlpre.HosMediCode = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                xmlpre.MediName = StrToXML(str药品)
                xmlpre.DCL = FormatEx(Val(.TextMatrix(i, gobjCOL.intCOL单量)), 5)
                xmlpre.PCDM = StrToXML(str频率编码)
                xmlpre.Unit = StrToXML(.TextMatrix(i, gobjCOL.intCOL单量单位))
                xmlpre.GYTJ = str给药途径
                xmlpre.GroupNum = Val(.TextMatrix(i, gobjCOL.intCOL相关ID))
                xmlpre.BTime = Format(IIf(glngModel = PM_住院编辑, .TextMatrix(i, gobjCOL.intCOL开始时间), .Cell(flexcpData, i, gobjCOL.intCOL开始时间)), "yyyy-MM-dd HH:mm:ss")
                
                xmlpre.ETime = Format(.TextMatrix(i, gobjCOL.intCOL终止时间), "yyyy-MM-dd HH:mm:ss")
                xmlpre.PresTime = Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd HH:mm:ss")
                
                strXML = MakePresXML(xmlpre, 1)
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = strXML
            End If
        Next
    End With
    
        
    InAdviceCheckWarn_DT = True
    If UBound(arrTmp) >= 0 Then
        On Error GoTo errH
        strXML = MakeXML(xmlbase, arrTmp, 1)
        WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strXML
        If gbytSuperVolume = 0 Then
            strTmp = dtywzxUI2(28676, 1, strXML, strRetXML)
            WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strTmp
            strRetXML = GetAlertFromXml(strRetXML)
            If InStr(strRetXML, ";CJLJJ;") > 0 Then
                MsgBox "用药监测系统发现当前医嘱存在超极量禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                InAdviceCheckWarn_DT = False: Exit Function
            End If
            strRetXML = ""
        Else
            strTmp = dtywzxUI(28676, 1, strXML) '分析处方
            WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strTmp
        End If
        
        If glngModel = PM_住院编辑 Then
            If strTmp = "2" And gbytBlackLamp = 0 Then
                If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                    If MsgBox("用药监测系统发现有院外执行的药品存在禁忌用药，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        InAdviceCheckWarn_DT = False
                    End If
                Else
                    MsgBox "用药监测系统发现当前医嘱存在禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                    InAdviceCheckWarn_DT = False: Exit Function
                End If
            ElseIf strTmp = "1" Or strTmp = "2" And gbytBlackLamp = 1 Then
                If MsgBox("用药监测系统发现当前医嘱存在禁忌用药，是否继续?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then InAdviceCheckWarn_DT = False
            End If
            If InAdviceCheckWarn_DT Then
                If gbytSuperVolume = 0 Then
                    strTmp = dtywzxUI2(28685, 1, strXML, strRetXML)
                    WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strTmp
                    strRetXML = GetAlertFromXml(strRetXML)
                    If InStr(strRetXML, ";CJLJJ;") > 0 Then
                        MsgBox "用药监测系统发现当前医嘱存在超极量禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                        InAdviceCheckWarn_DT = False
                        Exit Function
                    End If
                    strRetXML = ""
                Else
                    strTmp = dtywzxUI(28685, 1, strXML)
                    WriteLog "" & glngModel, "InAdviceCheckWarn_DT", strTmp
                End If
            End If
        Else
            If strTmp = "2" Then
                'MsgBox "用药监测系统发现当前医嘱存在严重问题，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                InAdviceCheckWarn_DT = False
            ElseIf strTmp = "1" Then
                'If MsgBox("用药监测系统发现当前医嘱存在一般问题，是否继续?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then InAdviceCheckWarn_DT = False
            End If
            '不调用保存处方接口28685
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    InAdviceCheckWarn_DT = False
End Function

Public Function InAdviceCheckWarn_TYT(ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional blnIsHaveOut As Boolean, _
    Optional ByRef blnNoSave As Boolean, Optional ByRef rsOut As ADODB.Recordset) As Long
'功能：调用太元通系统中对医嘱进行合理用药审查等相关功能
'参数：lngCmd=
'       0-用药规范;1-获取医嘱审查结果,并填写警示灯
'       2-药品提示
'       3-医药知识库;4-系统配置;5-点击警示灯，获取警示详情
'      lngRow=当前药品医嘱的行号，lngCmd=2时需要
'出参:
'   rsOut-禁忌药品说明
'      返回：blnIsHaveOut=是否存在离院带药的药品
'返回值：医嘱保存调用，需要用返回值判断是否存在禁忌用药
    Dim strDrugCode As String, str医生编码 As String, str开嘱医生 As String, strDescription As String
    Dim strSQL As String, strOrderInfo As String, str频率编码 As String, str频率 As String
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String, str相关ID As String
    Dim str给药途径 As String, str药品 As String, str期效 As String, str中药组IDs As String
    Dim str医嘱ID As String
    Dim blnDo As Boolean
    Dim curDate As Date
    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, rs中药 As ADODB.Recordset
    Dim udtPatiOrder As PatientOrder
    Dim udtDrug As PatDrug, udtPatiDiag As PatDiagnosis
    Dim udtPatiSensitive As PatDrugSensitive, UdtPatiSymptom As PatSymptom
    Dim udtAuditResult As AuditResult

    Dim i As Long, k As Long, j As Long, lngMaxWarn As Long, lng中药组ID As Long
    Dim lngBegin As Long, lngEnd As Long, lngLight As Long
    Dim strTmp As String, strOld As String
    Dim arrTmp As Variant, colAuditResult As Collection
    Dim arrLight(1 To 3) As String

    On Error GoTo errH
    Screen.MousePointer = 11

    With gobjAdvice
        Select Case lngCmd
        Case 0   '0-用药规范

            gobjPass.getPdssPrescription

        Case 1  '1-获取医嘱审查结果,并填写警示灯
            If gbytReason = 1 And glngModel = PM_住院编辑 Then
                Set rsOut = InitAdviceRS(FUN_输出内容)
            End If
            strSQL = _
            " Select A.住院号,Nvl(B.姓名,A.姓名) 姓名,Nvl(B.性别,A.性别) 性别 ,A.出生日期,B.身高,B.体重  " & _
                     " From 病人信息 A,病案主页 B" & _
                     " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng病人ID, gobjPati.lng主页ID)
            If rsPati.EOF Then Screen.MousePointer = 0: Exit Function

            '病人信息
            With udtPatiOrder
                '传人病人信息:病人ID,姓名,性别 1-女, 0-男, 2-不详，病人出生日期，格式 YYYY-MM-DD 不为空（必填）

                .PatientID = gobjPati.lng病人ID & ""
                .Pname = rsPati!姓名 & ""
                .pSex = IIf(rsPati!性别 & "" = "男", "0", IIf(rsPati!性别 & "" = "女", "1", "2"))
                .pdateOfBirth = Format(rsPati!出生日期, "yyyy-MM-dd")
                .pHeight = IIf(Val(rsPati!身高 & "") = 0, "", rsPati!身高 & "")
                .pWeight = IIf(Val(rsPati!体重 & "") = 0, "", rsPati!体重 & "")
                .PvisitID = rsPati!住院号 & ""

                '传人病人生理情况
                strTmp = Get病人病生理情况(gobjPati.lng病人ID, gobjPati.lng主页ID)
                .isLact = IIf(InStr(strTmp, "哺乳期") > 0, "1", "0")    '是否哺乳，是为1，否为0 不为空
                .isPregnant = IIf(InStr(strTmp, "孕妇") > 0, "1", "0")    '是否孕妇，是为1 ，否为0 不为空
                .isLiverWhole = IIf(InStr(strTmp, "肝功能异常") > 0, "1", "0") '是否肝功异常 1-异常，0-正常 不为空
                .isKidneyWhole = IIf(InStr(strTmp, "肾功能异常") > 0, "1", "0") '是否肾功异常 1-异常，0-正常 不为空

                '登录医生信息
                .DoctDeptID = UserInfo.部门ID & ""
                .DoctDeptName = UserInfo.部门名 & ""
                .DoctID = UserInfo.编号 & ""
                .DoctName = UserInfo.姓名 & ""
                .DoctTitleID = GetDoctorTitleType(UserInfo.专业技术职务)
                .DoctTitleName = IIf(UserInfo.专业技术职务 = "", "其他职务", UserInfo.专业技术职务)
                .SysFlag = "2"  '2-住院医生站，1-门诊医生站
            End With

            '药品信息
            curDate = zlDatabase.Currentdate
            arrTmp = Array()
            With gobjAdvice

                For i = .FixedRows To .Rows - 1
                    If glngModel = PM_住院医嘱清单 Then
                        blnDo = (InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0) Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4")
                        '一并给药，期效取首行期效
                        If RowIn一并给药(i, lngBegin, lngEnd) Then
                            str期效 = .TextMatrix(lngBegin, gobjCOL.intCOL期效)
                        Else
                            str期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                        End If
                        If blnDo Then
                            '不含已作废的，已停止的，确认停止的医嘱;包含当天的临嘱
                            blnDo = (str期效 = "长嘱" Or str期效 = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")) _
                                    And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0
                        End If
                    Else
                        '药嘱状态：作废、停止、确认停止 不做审查
                        blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 And InStr("4,8,9", .TextMatrix(i, gobjCOL.intCOL状态)) = 0
    
                        If blnDo Then
                            blnDo = .RowData(i) <> 0 And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                                    And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                            blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" _
                                               Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                        End If
                    End If
                    
                    If blnDo Then
                        If glngModel = PM_住院医嘱清单 And (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                            '获取中药医嘱组ID
                            str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                        Else
                            '取药品名称
                            If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                                str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                            Else
                                str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                            End If
    
                            If glngModel = PM_住院编辑 Then
                                '判断是否是院外执行的药品
                                If Val(.TextMatrix(i, gobjCOL.intCOL执行性质)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))), gobjCOL.intCOL执行性质)) = 5 Then
                                    blnIsHaveOut = True
                                End If
                            End If
                            
                            If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then  '一并给药不重复取
                                '给药途径
                                If glngModel = PM_住院编辑 Then
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                    If k <> -1 Then str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                                Else
                                    str给药途径 = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOL相关ID)), "诊疗项目ID") '传代码
                                End If
                                '取用药频率(次/天),都为整数四舍五入
                                Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), int频率次数, int频率间隔, str间隔单位, IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), str频率编码)
    
                                If str间隔单位 = "天" Then
                                    str频率 = int频率次数 & "/" & int频率间隔
                                ElseIf str间隔单位 = "周" Then
                                    str频率 = int频率次数 & "/7"
                                ElseIf str间隔单位 = "小时" Then
                                    If int频率间隔 <= 24 Then
                                        str频率 = Format(24 / int频率间隔 * int频率次数, "0") & "/1"
                                    Else
                                        str频率 = int频率次数 & "/" & Format(int频率间隔 / 24, "0")
                                    End If
                                ElseIf str间隔单位 = "分钟" Then
                                    str频率 = Format((24 * 60) / int频率间隔 * int频率次数, "0") & "/1"
                                End If
    
                                str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                                If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                                str医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
                            End If
    
                            udtDrug.drugID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)    'his 系统的药品代码不为空
                            udtDrug.DrugName = StrToXML(str药品)               'his 系统的药品名称不为空
                            udtDrug.recMainNo = .TextMatrix(i, gobjCOL.intCOL相关ID)     'his 系统的医嘱组号，在一次就诊/住院中唯
                            udtDrug.recSubNo = .TextMatrix(i, gobjCOL.intCOL序号)        'his 系统的医嘱序号，在一次就诊/住院中唯
                            udtDrug.dosage = Val(.TextMatrix(i, gobjCOL.intCOL单量))     'his 系统的医嘱药品使用剂量不为空
    
                            udtDrug.doseUnits = .TextMatrix(i, gobjCOL.intCOL单量单位)    'his 系统的医嘱药品剂量单位不为空
                            udtDrug.administrationID = str给药途径              'his 系统的医嘱途径代码不为空
                            udtDrug.performFreqDictID = StrToXML(str频率编码)   'his 系统的医嘱频次代码不为空
                            udtDrug.performFreqDictText = str频率               'his 系统的医嘱执行频率描述不为空
    
                            udtDrug.startDateTime = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:mm:ss")    'his 系统的医嘱开始时间,格式 YYYY-MM-DDHH: MM: SS 不为空
                            udtDrug.stopDateTime = Format(.TextMatrix(i, gobjCOL.intCOL终止时间), "yyyy-MM-dd HH:mm:ss")    'his 系统的医嘱结束时间,格式 YYYY-MM-DD HH: MM: SS
                            udtDrug.doctorDept = .TextMatrix(i, gobjCOL.intCOL开嘱科室ID)                 'his 系统的开医嘱医生所在科室代码
                            udtDrug.DoctorID = str医生编码                          'his 系统的开医嘱医生编码
                            udtDrug.Doctor = str开嘱医生                         'his 系统的开医嘱医生姓名,
                            If glngModel = PM_住院医嘱清单 Then
                                udtDrug.isNew = "0"                             '新增医嘱值为1；否则为0
                            Else
                                udtDrug.isNew = IIf(.TextMatrix(i, gobjCOL.intCOLEDIT) = "1", "1", "0")
                            End If
                            '
                            If Not rsOut Is Nothing Then
                                If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                                    '西药,中成药
                                    rsOut.AddNew
                                    rsOut!医嘱ID = CLng(CStr(.RowData(i)))
                                    rsOut!禁忌药品说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                                    rsOut!药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                                    rsOut!状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                                    rsOut.Update
                                ElseIf Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                                '中药配方
                                    k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                    If k <> -1 Then
                                        rsOut.AddNew
                                        rsOut!医嘱ID = CLng(CStr(.RowData(k)))
                                        rsOut!禁忌药品说明 = .TextMatrix(k, gobjCOL.intCol禁忌药品说明)
                                        rsOut!药品名称 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                        rsOut!状态 = .TextMatrix(k, gobjCOL.intCOL状态)
                                        rsOut.Update
                                    End If
                                End If
                                
                            End If
                            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                            arrTmp(UBound(arrTmp)) = udtDrug
                        End If
                    End If
                Next
                
                '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
                If glngModel = PM_住院医嘱清单 Then
                    If str中药组IDs <> "" Then
                        Set rs中药 = Get中药配方(str中药组IDs)
                        With rs中药
                            For i = 1 To .RecordCount
                                If !相关ID & "" <> str相关ID Then
                                    str开嘱医生 = !开嘱医生
                                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                                    str医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
                                    str频率 = GetFrequency(!间隔单位 & "", !频率次数 & "", !频率间隔 & "")
                                    Call Get频率信息_名称(!频率 & "", Val(!频率次数 & ""), Val(!频率间隔 & ""), !间隔单位 & "", 2, str频率编码)
                                    str相关ID = !相关ID & ""
                                End If
        
                                udtDrug.drugID = !药品ID & ""                      'his 系统的药品代码不为空
                                udtDrug.DrugName = !药品名称 & ""             'his 系统的药品名称不为空
                                udtDrug.recMainNo = !相关ID & ""             'his 系统的医嘱组号，在一次就诊/住院中唯
                                udtDrug.recSubNo = !序号 & ""      'his 系统的医嘱序号，在一次就诊/住院中唯
                                udtDrug.dosage = !单次用量 & ""     'his 系统的医嘱药品使用剂量不为空
        
                                udtDrug.doseUnits = !单量单位 & ""     'his 系统的医嘱药品剂量单位不为空
                                udtDrug.administrationID = !用法ID & ""              'his 系统的医嘱途径代码不为空
                                udtDrug.performFreqDictID = StrToXML(str频率编码)   'his 系统的医嘱频次代码不为空
                                udtDrug.performFreqDictText = str频率               'his 系统的医嘱执行频率描述不为空
         
                                udtDrug.startDateTime = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")     'his 系统的医嘱开始时间,格式 YYYY-MM-DDHH: MM: SS 不为空
                                udtDrug.stopDateTime = Format(!终止时间 & "", "yyyy-MM-dd HH:mm:ss")    'his 系统的医嘱结束时间,格式 YYYY-MM-DD HH: MM: SS
                                udtDrug.doctorDept = !开嘱科室id & ""               'his 系统的开医嘱医生所在科室代码
                                udtDrug.DoctorID = str医生编码                          'his 系统的开医嘱医生编码
                                udtDrug.Doctor = str开嘱医生                         'his 系统的开医嘱医生姓名,
                                udtDrug.isNew = "0"
                                
                                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                arrTmp(UBound(arrTmp)) = udtDrug
                                .MoveNext
                            Next
                        End With
                    End If
                End If
            End With
           
            If UBound(arrTmp) = -1 Then
                Screen.MousePointer = 0: Exit Function
            End If
            udtPatiOrder.PatDrugs = arrTmp

            '诊断
            arrTmp = Array()
            Set rsTmp = Get病人诊断记录(gobjPati.lng病人ID, gobjPati.lng主页ID, "2,12")   '西医住院，中医住院

            For i = 0 To rsTmp.RecordCount - 1
                udtPatiDiag.diagnosisID = rsTmp!编码 & ""       'his 系统的诊断编码
                udtPatiDiag.diagnosisName = rsTmp!名称 & ""     'his 系统的诊断名称
                udtPatiDiag.diagnosisType = "入院诊断"          '系统的诊断类型，如门诊诊断、入院诊断等
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = udtPatiDiag
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatDiagnoses = arrTmp
            '过敏
            arrTmp = Array()
            Set rsTmp = Get病人过敏记录(gobjPati.lng病人ID, gobjPati.lng主页ID)
            For i = 0 To rsTmp.RecordCount - 1
                udtPatiSensitive.patOrderDrugSensitiveID = "0"          '固定值
                udtPatiSensitive.drugAllergenID = rsTmp!过敏源编码 & ""    '系统的过敏编码
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = udtPatiSensitive
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatDrugSensitives = arrTmp

            '症状
            arrTmp = Array()
            Set rsTmp = GetPatiSymptom(gobjPati.lng病人ID, gobjPati.lng主页ID)
            For i = 0 To rsTmp.RecordCount - 1
                UdtPatiSymptom.symptomID = rsTmp!编码 & ""              'his 系统的症状编码
                UdtPatiSymptom.symptomName = rsTmp!名称 & ""            'his 系统的症状名称

                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = UdtPatiSymptom
                rsTmp.MoveNext
            Next
            udtPatiOrder.PatSymptoms = arrTmp

            strOrderInfo = MakePatientOrderXml(udtPatiOrder)

            '医嘱信息审查接口调用"

            strDescription = gobjPass.checkDrugSecurityWS(strOrderInfo, "1")

            '审查结果处理
            '返回值顺及警示级别(高到低)：1— 禁忌（建议显示红色警示灯）；2— 慎用（建议显示黄色警示灯示）；3— 提示（建议显示蓝色警示灯）
            lngMaxWarn = 4
            If strDescription = "" Then
                MsgBox "药嘱审查功能未执行，请检查太元通接口配置是否有误！", vbInformation + vbOKOnly, G_STR_PASS
                Screen.MousePointer = 0: Exit Function

            ElseIf strDescription = "-101" Then
                '-101：表示用户可以忽略该返回值，不做业务处理。
            Else
                Set colAuditResult = AnalyzeReturnXml(strDescription)
                
                If glngModel = PM_住院医嘱清单 Then arrTmp = Array()
                
                With gobjAdvice
                    '获取警示灯
                    '图标颜色frmIcons.imgpass ：1-红，2-黄，3-蓝
                    arrLight(1) = "红": arrLight(2) = "黄": arrLight(3) = "蓝"
                    For i = .FixedRows To .Rows - 1
                        If glngModel = PM_住院医嘱清单 Then
                            blnDo = InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                            '一并给药，只在首行显示期效,其余行擦除（vsAdvice_DrawCell）
                            '一并给药，期效取首行期效
                            If RowIn一并给药(i, lngBegin, lngEnd) Then
                                str期效 = .TextMatrix(lngBegin, gobjCOL.intCOL期效)
                            Else
                                str期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                            End If
                            '不含已作废的，停止的，确认停止的医嘱,含当天的临嘱
                            blnDo = InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0 And (str期效 = "长嘱" _
                                    Or str期效 = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                        Else
                            '作废，停止，确认停止的不审查
                            blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                                    And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                                    And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                            blnDo = blnDo And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0 And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" _
                                    Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                        End If

                        If blnDo Then
                            strTmp = .TextMatrix(i, gobjCOL.intCOL相关ID) & "_" & .TextMatrix(i, gobjCOL.intCOL序号)   '关键字格式:组医嘱号_医嘱序号
                            On Error Resume Next
                            udtAuditResult = colAuditResult(strTmp)
                            If Err.Number > 0 Then
                                strTmp = "未找到"
                            End If
                            Err.Clear: On Error GoTo 0
                            If strTmp <> "未找到" Then  '找到审核警示灯
                                k = Val(udtAuditResult.alertLevel)
                            Else
                                k = 0
                            End If
                            
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                                If Not rsOut Is Nothing And glngModel = PM_住院编辑 And k = 1 Then
                                    rsOut.Filter = "医嘱ID=" & CLng(.RowData(i) & "") & " And 状态 < 3 "
                                    If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                                End If
                                '设置警示灯
                                strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                                If k >= 1 And k <= 3 Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(k)
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                                Else
                                    .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                                End If
    
                                '标记审查结果变化,以备更新数据库
                                If glngModel = PM_住院医嘱清单 Then
                                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                        ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                        arrTmp(UBound(arrTmp)) = "ZL_病人医嘱记录_更新审查(" & .TextMatrix(i, gobjCOL.intCOLID) & "," & _
                                                                 IIf(CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) = "", "NULL", Val(.Cell(flexcpData, i, gobjCOL.intCOL警示))) & ")"
                                    End If
                                Else
                                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                        .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                        blnNoSave = True    '标记为未保存
                                    End If
                                End If
                            ElseIf .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                                If glngModel = PM_住院编辑 Then
                                    '中药配方
                                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                                        lng中药组ID = CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))          '中药配方组ID
                                        lngLight = 4 '初始化
                                    End If
                                    '设置警示灯 取草药中最大警示值
                                    If k > 0 Then
                                        If lngLight > k Then
                                            lngLight = k
                                        End If
                                    End If
                                End If
                            End If
                            
                            '记录最高级别警示值 (警示值越小警示级越高)
                            If k > 0 Then
                                If lngMaxWarn > k Then
                                    lngMaxWarn = k
                                End If
                            End If
                        Else
                            If glngModel = PM_住院编辑 Then
                                If .RowData(i) = lng中药组ID And .RowData(i) <> 0 Then
                                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                                    '设置警示灯
                                    If lngLight >= 1 And lngLight <= 3 Then
                                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                                    Else
                                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                                    End If
                                    
                                    If glngModel = PM_住院编辑 Then
                                        '标记审查结果变化,以备更新数据库
                                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                            .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                            blnNoSave = True    '标记为未保存
                                        End If
                                        
                                        If Not rsOut Is Nothing And lngLight = 1 Then
                                            rsOut.Filter = "医嘱ID=" & lng中药组ID & " And 状态 < 3 "
                                            If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                                        End If
                                    End If
                                    lng中药组ID = 0
                                    lngLight = 4
                                End If
                            End If
                        End If
                    Next
                    
                    '医嘱清单中药配方警示灯处理
                    If glngModel = PM_住院医嘱清单 And Not rs中药 Is Nothing Then
                        For i = .FixedRows To .Rows - 1
                            '中药服法
                            If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                                strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                                lngLight = 4
                                str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                                rs中药.Filter = "相关ID=" & str医嘱ID
                                
                                For j = 1 To rs中药.RecordCount
                                    strTmp = rs中药!相关ID & "_" & rs中药!序号   '关键字格式:组医嘱号_医嘱序号
                                    On Error Resume Next
                                    udtAuditResult = colAuditResult(strTmp)
                                    If Err.Number > 0 Then
                                        strTmp = "未找到"
                                    End If
                                    Err.Clear: On Error GoTo 0
                                    If strTmp <> "未找到" Then  '找到审核警示灯
                                        k = Val(udtAuditResult.alertLevel)
                                    Else
                                        k = 0
                                    End If
                                    If k > 0 Then
                                        If lngLight > k Then
                                            lngLight = k
                                        End If
                                    End If
                                    
                                    rs中药.MoveNext
                                Next
                                
                                '设置警示灯
                                If lngLight >= 1 And lngLight <= 3 Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                                Else
                                    .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                    Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                                End If
                                
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                    ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                                    arrTmp(UBound(arrTmp)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(lngLight >= 1 And lngLight <= 3, lngLight, "NULL") & ")"
                                End If
                                
                                '记录最高级别警示值
                                If lngLight > 0 Then
                                    If lngMaxWarn > lngLight Then
                                        lngMaxWarn = lngLight
                                    End If
                                End If
                            End If
                        Next
                    End If
                    
                End With
                 '数据提交,不开启事务
                If glngModel = PM_住院医嘱清单 Then
                    For i = 0 To UBound(arrTmp)
                        Call zlDatabase.ExecuteProcedure(CStr(arrTmp(i)), "合理用药监测")
                    Next
                End If
            End If

        Case 2    ' 2-药品提示
            If InStr(",5,6,7,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)) <> 0 Then
                '获取所选医嘱的药品编码
                strDrugCode = .TextMatrix(lngRow, gobjCOL.intCOL收费细目ID)
                '调用药品提示接口
                gobjPass.getDrugExplain (strDrugCode)
            Else
                MsgBox "当前选中的医嘱不是按规格下达的药品医嘱。", vbInformation + vbOKOnly, "合理用药监测"
            End If
        Case 3    '3-在线医药知识库
            '调用在线医药知识库
            gobjPass.accessIFMI ("0")  '传入值固定为:"0",无返回值
        Case 4  '4-系统配置
            gobjPass.sysConfig
        Case 5    '5-获取警示详情
            gobjPass.getDrugAlertDetail
        End Select
    End With
    InAdviceCheckWarn_TYT = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PassInitialize() As Boolean
'功能：对PASS接口进行注册和初始化，同时检查PASS接口DLL是否正确安装
    Dim lngTmp As Long
    Dim strRet As String
    Dim strCheckMode As String
    Dim strDetails As String
    Dim udtBase As YWS_BASE
    Dim udtDTBSBase As DTBS_BASE
    
    
    On Error GoTo errH
    
    If gbytPass = UNPASS Then Exit Function   '83970 PASSMap调用引起的
    
    If gbytPass = MK Then
        If gstrVersion = "3.0" Then
            'PASS功能函数注册(共享客户端模式)
            If RegisterServer <> 0 Then
                MsgBox "PASS客户端注册失败，当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
                Exit Function
            End If
            
            'PASS初始化
            If PassInit(UserInfo.编号 & "/" & UserInfo.用户名, UserInfo.部门码 & "/" & UserInfo.部门名, 10) <> 1 Then
                MsgBox "PASS系统初始化失败，当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
                Exit Function
            End If
                    
            'PASS是否可用检测初始化成功后,不必再调用状态接口 美康工程师提出
            If PassGetState("PassEnable") = 0 Then
                MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
                Call PassQuit: Exit Function
            End If
            
            'PASS应用模式设置(用默认值)
            Call PassSetControlParam(1, 1, 0, 2, 1)    '113198
            
        ElseIf gstrVersion = "4.0" Then
            If glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
                strCheckMode = "mz"
            ElseIf glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Or glngModel = PM_护士校对 Then
                strCheckMode = "zy"
            ElseIf glngModel = PM_PIVA管理 Then
                strCheckMode = "pivas"
            ElseIf glngModel = PM_处方发药 Then
                strCheckMode = "mzyf"
            ElseIf glngModel = PM_部门发药 Then
                strCheckMode = "zyyf"
            ElseIf glngModel = PM_门诊处方审查 Then
                strCheckMode = "mzsc"
            ElseIf glngModel = PM_住院药嘱审查 Then
                strCheckMode = "zysc"
            End If
            '参数未设置医院编码时取站点号
            If gstrHOSCODE = "" Then gstrHOSCODE = IIf(gstrNodeNo = "-", "0", gstrNodeNo)
            lngTmp = MDC_Init(strCheckMode, gstrHOSCODE, UserInfo.编号)
            
            If lngTmp <= 0 Then   'YWJ需要传人模式
                MsgBox "PASS4.0系统初始化失败，当前合理用药监测系统不可用，请检查相关配置是否正确。" & vbCrLf & _
                        "返回值:" & lngTmp & vbTab & _
                        "错误信息:【" & MDC_GetLastError() & "】", vbInformation, gstrSysName
                Exit Function
            End If
            
        End If
    ElseIf gbytPass = DT Then
        If gstrVersion = "3.0" Then  'CS版
            lngTmp = dtywzxUI(0, 0, "") '初始化接口,打开大通程序
            lngTmp = dtywzxUI(768, 0, UserInfo.编号) '传入医生工号
        ElseIf gstrVersion = "4.0" Then 'BS版
            With udtDTBSBase
                .strHIS = "HIS"
                .str医院编码 = gstrHOSCODE
                .str医生代码 = UserInfo.编号
                .str医生级别代码 = UserInfo.专业技术编码   '要求必填
                .str医生级别名称 = UserInfo.专业技术职务
                .str医生名称 = UserInfo.姓名
                .str科室代码 = UserInfo.部门码
                .str科室名称 = UserInfo.部门名
            End With
            gstrBaseXml = DTBS_MakeBASEXML(udtDTBSBase)
            strDetails = DTBS_MakeDetailXML(DTBS_登录, "")
            WriteLog "" & glngModel, "PassInitialize", gstrBaseXml
            WriteLog "" & glngModel, "PassInitialize", strDetails
            lngTmp = CRMS_UI(DTBS_登录, gstrBaseXml, strDetails, "")
            WriteLog "" & glngModel, "PassInitialize", "登录接口返回值:" & lngTmp
        End If
    ElseIf gbytPass = TYT Then
        On Error Resume Next
        Set gobjPass = GetObject(, "Midlayer.ComInterface")
        Err.Clear: On Error GoTo 0
        On Error Resume Next
        If gobjPass Is Nothing Then Set gobjPass = CreateObject("Midlayer.ComInterface")
        Err.Clear: On Error GoTo 0
        If gobjPass Is Nothing Then
            MsgBox "太元通接口初始化失败,可能合理用药监测系统未正确安装或配置。" & _
                   vbCrLf & "在正确安装和配置合理用药监测系统之前，相应的功能不能使用。", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf gbytPass = YWS Then '广州保进 药卫士
        On Error Resume Next
        Set gobjPass = GetObject(, "YWSUI.YWS")
        Err.Clear: On Error GoTo 0
        On Error Resume Next
        If gobjPass Is Nothing Then Set gobjPass = CreateObject("YWSUI.YWS")
        If Err.Number <> 0 Then
            MsgBox "药卫士登录失败,可能合理用药监测系统未正确安装或配置。" & _
                   vbCrLf & "在正确安装和配置合理用药监测系统之前，相应的功能不能使用。", vbInformation, gstrSysName
            WriteLog "" & glngModel, "PassInitialize", "动态创建组件——错误号:" & Err.Number & "|错误描述:" & Err.Description
            Exit Function
        End If
        Err.Clear: On Error GoTo 0
        With udtBase
            .strHIS = "HIS"
            .str医院编码 = ""
            .str医生代码 = UserInfo.编号
            .str医生级别代码 = UserInfo.专业技术编码   '要求必填
            .str医生级别名称 = UserInfo.专业技术职务
            .str医生名称 = UserInfo.姓名
            .str科室代码 = UserInfo.部门码
            .str科室名称 = UserInfo.部门名
        End With
        gstrBaseXml = YWS_MakeBASEXML(udtBase)
        WriteLog "" & glngModel, "PassInitialize", gstrBaseXml
        strRet = gobjPass.YWS_UI(YWS_登录, gstrBaseXml, "", "")
        WriteLog "" & glngModel, "PassInitialize", "登录接口返回值:" & lngTmp
    ElseIf gbytPass = HZYY Then
        '测试地址:"http://118.31.246.211:8080/zlcx/data_detail.action?webHisId=11221&hospitalCode=cqzl123"
    ElseIf gbytPass = ZL Then
        Call ZLShowWindow
    End If
    
    PassInitialize = True
    Exit Function
errH:
    If Err.Number = 53 And InStr(UCase(Err.Description), UCase("ShellRunAs")) > 0 Then
        MsgBox "PASS接口文件 ShellRunAs.dll 不存在,可能合理用药监测系统未正确安装或配置。" & _
            vbCrLf & "在正确安装和配置合理用药监测系统之前，相应的功能不能使用。", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("DIFPassDll")) > 0 Then
        MsgBox "PASS接口文件 DIFPassDll.dll 不存在,可能是因为以下原因：" & vbCrLf & _
            vbCrLf & "1.PASS客户端是第一次登录，请退出之后再重新登录即可正常使用。" & _
            vbCrLf & "2.合理用药监测系统未正确安装或配置，请仔细检查后再登录重试。", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("dtywzxUI")) > 0 Then
        MsgBox "PASS接口文件 dtywzxUI.dll 不存在,可能合理用药监测系统未正确安装或配置。" & _
            vbCrLf & "在正确安装和配置之前，相应的功能不能使用。", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("CRMS_UI")) > 0 Then
        MsgBox "PASS接口文件 CRMS_UI.dll 不存在,可能合理用药监测系统未正确安装或配置。" & _
            vbCrLf & "在正确安装和配置之前，相应的功能不能使用。", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("PASS4Invoke")) > 0 Then
          MsgBox "PASS接口文件 PASS4Invoke.dll 不存在,可能合理用药监测系统未正确安装或配置。" & _
            vbCrLf & "在正确安装和配置之前，相应的功能不能使用。", vbInformation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function

Public Function InAdviceCheckWarn_YWS(Optional ByRef blnNoSave As Boolean) As Boolean
'功能：调用保进药卫士用药监测系统对医嘱进行合理用药审查等相关功能
    Dim udtDetail As YWS_DETAILS
    Dim udtPati As YWS_PATIENT
    Dim colTmp As Collection
    Dim udt过敏源 As YWS_ALLERGIC
    Dim udt诊断 As YWS_DIAGNOSE
    Dim udtPres As YWS_PRESCRIPTION
    Dim udtMedic As YWS_MEDICINE   '药品信息
    Dim i As Long, j As Long, k As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lng中药组ID As Long, lngLight As Long
    Dim arrTmp As Variant, arrSQL As Variant
    Dim curDate As Date
    Dim rsTmp As Recordset, rsRet As ADODB.Recordset, rsPati As ADODB.Recordset, rs中药 As ADODB.Recordset
    Dim strTmp As String, str单量 As String, str单量单位 As String
    Dim str医嘱ID As String, strRetXML As String, str相关ID As String
    Dim str药品 As String, str给药途径 As String, str频率编码 As String, strXML As String
    Dim str期效 As String, strOld As String, str中药组IDs As String
    Dim arrLight(0 To 4) As String
    
    Dim blnDo As Boolean, blnIsHaveOut As Boolean  '判断是否存在院外执行的药品

    '病人信息
    Set rsPati = GetPatiInfo(gobjPati.lng病人ID, gobjPati.lng主页ID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    With udtPati
        .str姓名 = rsPati!姓名
        .str出生日期 = rsPati!出生日期
        .str性别 = rsPati!性别
        .str体重 = rsPati!身高 & ""
        .str身高 = rsPati!体重 & ""
        .str身份证号 = rsPati!身份证号 & ""
        .str病历卡号 = rsPati!住院号 & ""
        .str卡号 = ""
        .str卡类型 = ""
        .str怀孕时间 = ""
        .str怀孕时间单位 = ""
        '过敏源
        Set colTmp = New Collection
        Set rsTmp = Get病人过敏记录(gobjPati.lng病人ID, gobjPati.lng主页ID, 1)
        For i = 1 To rsTmp.RecordCount
            If "" & rsTmp!药物ID <> "" Then
                With udt过敏源
                    .str过敏类型 = "5"   '1=药卫士药品大类 2=药卫士药品成份
                    .str过敏源名称 = rsTmp!药物名
                    .str过敏源代码 = "" & rsTmp!药物ID
                End With
                colTmp.Add udt过敏源, "_" & i
            End If
            rsTmp.MoveNext
        Next
        Set .col过敏源s = colTmp
        
        '诊断记录
        Set colTmp = New Collection
        Set rsTmp = Get病人诊断记录(gobjPati.lng病人ID, gobjPati.lng主页ID, "2,12")
        For i = 1 To rsTmp.RecordCount
            With udt诊断
                If rsTmp!疾病ID & "" <> "" Then
                     .str诊断类型 = "2" '2=IDC10代码
                Else
                    .str诊断类型 = "0" '0=其他
                End If
                .str诊断代码 = "" & rsTmp!编码
                .str诊断名称 = "" & rsTmp!名称
            End With
            colTmp.Add udt诊断, "_" & colTmp.Count + 1
            rsTmp.MoveNext
        Next
        '病生理
        strTmp = Get病人病生理情况(gobjPati.lng病人ID, gobjPati.lng主页ID)
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                With udt诊断
                    .str诊断类型 = "1" '1=病生理状态
                    .str诊断代码 = Sys.RowValue("病生理情况", arrTmp(i), "编码", "名称")
                    .str诊断名称 = arrTmp(i)
                End With
                colTmp.Add udt诊断, "_" & colTmp.Count + 1
            Next
           
        End If
        Set .col诊断s = colTmp
    End With
    
    curDate = zlDatabase.Currentdate
    
    With udtDetail
        .strHIS系统时间 = Format(curDate, "YYYY-MM-DD HH:MM:SS")
        .str门诊住院标识 = "ip"    '住院标识

        .str就诊类型 = YWS_GetTreatType(2, gobjPati.lng病人ID, gobjPati.lng主页ID)
        .str就诊号 = rsPati!住院号 & ""
        .str床位号 = "" & rsPati!当前床号
        '病人信息
        .udt病人信息 = udtPati
        '处方信息
        .udt处方信息 = udtPres
    End With
    '药品信息
    Set colTmp = New Collection
    
    With gobjAdvice
        For i = .FixedRows To .Rows - 1
           If glngModel = PM_住院编辑 Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                        And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0 _
                        Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                        And Val(.TextMatrix(i, gobjCOL.intCOL状态)) <> 4)
                        
            ElseIf glngModel = PM_住院医嘱清单 Then
                blnDo = ((InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0) Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4"))
                
                If blnDo Then
                    '一并给药，只在首行显示期效,其余行擦除（见vsAdvice_DrawCell）
                    '一并给药，期效取首行期效
                    If RowIn一并给药(i, lngBegin, lngEnd) Then
                        str期效 = .TextMatrix(lngBegin, gobjCOL.intCOL期效)
                    Else
                        str期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                    End If
                    '不含已作废的医嘱,停止和确认停止的长嘱;包含当天的临嘱
                    blnDo = str期效 = "长嘱" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0 _
                            Or str期效 = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                            And .TextMatrix(i, gobjCOL.intCOL状态) <> "4"
                End If
            End If

            If blnDo Then
                '获取中药医嘱组ID
                If glngModel = PM_住院医嘱清单 And (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                    str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    '取药品名称
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                        str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                    Else
                        str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                    End If
    
                    '取药品给药途径
                    If glngModel = PM_住院编辑 Then
                         str医嘱ID = CStr(.RowData(i))
                        '判断是否是院外执行的药品
                        If Val(.TextMatrix(i, gobjCOL.intCOL执行性质)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))), gobjCOL.intCOL执行性质)) = 5 Then
                            blnIsHaveOut = True
                        End If
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str给药途径 = "" '一并给药不重复取
                        If str给药途径 = "" Then
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                            If k <> -1 Then str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                        End If
                    ElseIf glngModel = PM_住院医嘱清单 Then
                        str医嘱ID = CStr(.TextMatrix(i, gobjCOL.intCOLID))
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then  '一并给药不重复取
                            str给药途径 = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOL相关ID)), "诊疗项目ID")   '传代码
                        End If
                    End If
                    
                    Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), 0, 0, "", IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), str频率编码)
                    
                    udtMedic.str药品类型 = YWS_GetDrugType(.TextMatrix(i, gobjCOL.intCOL诊疗类别))
                    udtMedic.str处方号 = str医嘱ID    '传医嘱ID
                    udtMedic.Str医嘱类型 = IIf(.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱", "L", "T")
                    udtMedic.str处方时间 = Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd HH:mm:ss")          '处方时间（YYYY-MM-DD HH:mm:SS）
                    udtMedic.str商品名 = YWS_StrToXML(Sys.RowValue("诊疗项目目录", Val(.TextMatrix(i, gobjCOL.intCOL诊疗项目ID)), "名称"))
                    udtMedic.str医院药品代码 = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                    udtMedic.str医保代码 = ""
                    udtMedic.str批准文号 = ""
                    udtMedic.str用药开始时间 = Format(IIf(glngModel = PM_住院编辑, .TextMatrix(i, gobjCOL.intCOL开始时间), .Cell(flexcpData, i, gobjCOL.intCOL开始时间)), "yyyy-MM-dd HH:mm:ss")
                    udtMedic.str用药结束时间 = Format(.TextMatrix(i, gobjCOL.intCOL终止时间), "yyyy-MM-dd HH:mm:ss")
                    udtMedic.str规格 = ""
                    udtMedic.str组号 = .TextMatrix(i, gobjCOL.intCOL相关ID)
                    udtMedic.str用药理由 = ""
                    '单量，单量单位
                    str单量 = .TextMatrix(i, gobjCOL.intCOL单量)
                    str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                    str单量 = Replace(str单量, str单量单位, "")
                    
                    udtMedic.str单次量单位 = str单量单位
                    udtMedic.str单次量 = str单量
                    udtMedic.str频次代码 = str频率编码
                    udtMedic.str给药途径代码 = str给药途径
                    udtMedic.str服药天数 = .TextMatrix(i, gobjCOL.intCOL天数)   'OP 门诊处方有效
                    
                    colTmp.Add udtMedic, "_" & colTmp.Count + 1
                End If
            End If
        Next
        '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
        If glngModel = PM_住院医嘱清单 Then
            If str中药组IDs <> "" Then
                Set rs中药 = Get中药配方(str中药组IDs)
                With rs中药
                    For i = 1 To .RecordCount
                        If !相关ID & "" <> str相关ID Then
                            Call Get频率信息_名称(!频率 & "", 0, 0, "", IIf(!诊疗类别 & "" = "7", 2, 1), str频率编码)
                            str相关ID = !相关ID & ""
                        End If
                        udtMedic.str药品类型 = YWS_GetDrugType(!诊疗类别 & "")
                        udtMedic.str处方号 = !id & ""    '传医嘱ID
                        udtMedic.Str医嘱类型 = IIf(!医嘱期效 & "" = "0", "L", "T")
                        udtMedic.str处方时间 = Format(!开嘱时间 & "", "yyyy-MM-dd HH:mm:ss")
                        udtMedic.str商品名 = !药品名称 & ""
                        udtMedic.str医院药品代码 = !药品ID & ""
                        udtMedic.str医保代码 = ""
                        udtMedic.str批准文号 = ""
                        udtMedic.str用药开始时间 = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")
                        udtMedic.str用药结束时间 = Format(!终止时间 & "", "yyyy-MM-dd HH:mm:ss")
                        udtMedic.str规格 = ""
                        udtMedic.str组号 = !相关ID & ""
                        udtMedic.str用药理由 = ""
                        udtMedic.str单次量单位 = !单量单位 & ""
                        udtMedic.str单次量 = !单次用量 & ""
                        udtMedic.str频次代码 = str频率编码
                        udtMedic.str给药途径代码 = !用法ID & ""
                        udtMedic.str服药天数 = !天数 & ""   'OP 门诊处方有效
                        colTmp.Add udtMedic, "_" & colTmp.Count + 1
                        .MoveNext
                    Next
                End With
            End If
        End If
        With udtPres
            Set .col药品信息 = colTmp
            .str处方号 = "0"
            .str处方理由 = ""
            .str处方时间 = Format(curDate, "YYYY-MM-DD HH:MM:SS")
            .str是否当前处方 = "1" '0 历史处方 1 当前处方（现是默认当前处方，以后扩充）
            .Str医嘱类型 = "L"
        End With
    End With
    
    udtDetail.udt处方信息 = udtPres
    
    InAdviceCheckWarn_YWS = True
    If udtPres.col药品信息.Count > 0 Then
        On Error GoTo errH
        
        strXML = YWS_MakePresXML(udtDetail)
        WriteLog "" & glngModel, "InAdviceCheckWarn_YWS", strXML
        strTmp = gobjPass.YWS_UI(YWS_处方分析, gstrBaseXml, strXML, strRetXML)
        WriteLog "" & glngModel, "InAdviceCheckWarn_YWS", "返回值:" & strTmp & vbCrLf & strRetXML
        Set rsRet = YWS_ReturnRS(strRetXML)
        '设置警示灯
        With gobjAdvice
            arrSQL = Array()
            '图片下标：1-蓝灯,2-黄灯,5-橙灯,3-红灯
            '警示级顺：0-蓝灯,1-黄灯,2-橙灯,3-红灯
            arrLight(0) = "蓝": arrLight(1) = "黄": arrLight(2) = "橙": arrLight(3) = "红"
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_住院编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                            And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                    blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" _
                            Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                ElseIf glngModel = PM_住院医嘱清单 Then
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0

                    If blnDo Then
                        '一并给药，只在首行显示期效,其余行擦除（见vsAdvice_DrawCell）
                        '一并给药，期效取首行期效
                        If RowIn一并给药(i, lngBegin, lngEnd) Then
                            str期效 = .TextMatrix(lngBegin, gobjCOL.intCOL期效)
                        Else
                            str期效 = .TextMatrix(i, gobjCOL.intCOL期效)
                        End If
                        '不含已作废的医嘱,停止和确认停止的长嘱;包含当天的临嘱
                        blnDo = str期效 = "长嘱" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0 _
                                Or str期效 = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                                And .TextMatrix(i, gobjCOL.intCOL状态) <> "4"
                    End If
                End If
                If blnDo Then
                    If glngModel = PM_住院编辑 Then
                        str医嘱ID = .RowData(i) & ""
                    Else
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID) & ""
                    End If

                    rsRet.Filter = "医嘱ID ='" & str医嘱ID & "'"
                    If rsRet.RecordCount > 0 Then
                        k = CLng(rsRet!警示值 & "")
                    Else
                        k = 0
                    End If
               
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 1 Then
                        '西药、西成药'设置警示灯
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        If k >= 0 And k <= 3 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(k)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If

                        If glngModel = PM_住院编辑 Then
                            '标记审查结果变化,以备更新数据库
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                blnNoSave = True    '标记为未保存
                            End If
                        ElseIf glngModel = PM_住院医嘱清单 Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 3, k, "NULL") & ")"
                            End If
                        End If
                    Else
                        '中药配方
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                            lng中药组ID = .TextMatrix(i, gobjCOL.intCOL相关ID)          '中药配方组ID
                            lngLight = -1 '初始化
                        End If
                        '设置警示灯 取草药中最大警示值
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If k > lngLight Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                    End If
                Else
                    If glngModel = PM_住院编辑 Then
                        If .RowData(i) = lng中药组ID And .RowData(i) <> 0 Then
                            strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                            '设置警示灯
                            If lngLight >= 0 And lngLight <= 3 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                            Else
                                .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                            End If

                            If glngModel = PM_住院编辑 Then
                                '标记审查结果变化,以备更新数据库
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                    blnNoSave = True    '标记为未保存
                                End If
                            End If

                            lng中药组ID = 0
                            lngLight = -1
                        End If
                    End If
                End If
            Next
            '医嘱清单中药配方警示灯处理
            If glngModel = PM_住院医嘱清单 And Not rs中药 Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    '中药服法
                    If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        lngLight = -1
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                        rs中药.Filter = "相关ID=" & str医嘱ID

                        For j = 1 To rs中药.RecordCount
                            rsRet.Filter = "医嘱ID ='" & rs中药!id & "" & "'"
                            If rsRet.RecordCount > 0 Then
                                k = CLng(rsRet!警示值 & "")
                            Else
                                k = 0
                            End If
                            '设置警示灯 取草药中最大警示值
                            If k >= 0 Then
                                If lngLight >= 0 Then
                                    If k > lngLight Then
                                        lngLight = k
                                    End If
                                Else
                                    lngLight = k
                                End If
                            End If
                            rs中药.MoveNext
                        Next
                        '设置警示灯
                        If lngLight >= 0 And lngLight <= 3 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(lngLight >= 0 And lngLight <= 3, lngLight, "NULL") & ")"
                        End If
                    End If
                Next
            End If
            For i = LBound(arrSQL) To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
            Next
        End With
        
        If glngModel = PM_住院编辑 Then
            'YWS_处方分析=6 处方分析返回值：0、1、2，3，8分别代表没有问题，其他问题，一般问题，严重问题，有写申诉理由 保存前调用
            '1和2 悬浮界面有灯提示,点击灯可以查看详细信息
            '8-暂时没有提供
            If strTmp = "3" And gbytBlackLamp = 0 Then
                If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                    If MsgBox("用药监测系统发现有院外执行的药品存在禁忌用药，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        InAdviceCheckWarn_YWS = False
                        Exit Function
                    End If
                Else
                    MsgBox "用药监测系统发现当前医嘱存在禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                    InAdviceCheckWarn_YWS = False
                    Exit Function
                End If
            ElseIf strTmp = "3" And gbytBlackLamp = 1 Then
                If MsgBox("用药监测系统发现当前医嘱存在禁忌用药，是否继续?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then InAdviceCheckWarn_YWS = False: Exit Function
            End If
            
            '保存处方
            If strTmp <> "0" Then
                strTmp = gobjPass.YWS_UI(YWS_上传处方, gstrBaseXml, strXML, strRetXML)
                WriteLog "" & glngModel, "InAdviceCheckWarn_YWS", "保存处方:返回值" & strTmp
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    InAdviceCheckWarn_YWS = False
End Function

Public Function OutAdviceCheckWarn_YWS(Optional ByRef blnNoSave As Boolean) As Boolean
'功能：调用保进药卫士用药监测系统对医嘱进行合理用药审查等相关功能
    Dim udtDetail As YWS_DETAILS
    Dim udtPati As YWS_PATIENT
    Dim colTmp As Collection
    Dim udt过敏源 As YWS_ALLERGIC
    Dim udt诊断 As YWS_DIAGNOSE
    Dim udtPres As YWS_PRESCRIPTION
    Dim udtMedic As YWS_MEDICINE   '药品信息
    Dim curDate As Date
    Dim rsTmp As Recordset, rsRet As Recordset, rsPati As Recordset, rsPatiInfo As Recordset
    Dim rs中药 As ADODB.Recordset
    Dim arrTmp As Variant, arrSQL As Variant
    Dim i As Long, j As Long, k As Long
    Dim lng中药组ID As Long, lngLight As Long
    
    Dim str药品 As String, str给药途径 As String, str频率编码 As String, strXML As String
    Dim str单量 As String, str单量单位 As String, str身高 As String, str体重 As String
    Dim str医嘱ID As String, str开嘱时间 As String, strOld As String, str中药组IDs As String
    Dim strRetXML As String, strSQL As String, strTmp As String, str相关ID As String
    Dim arrLight(0 To 4) As String
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    Set rsPati = ReadPatient(gobjPati.lng病人ID, gobjPati.str挂号单)
    If rsPati.EOF Then Screen.MousePointer = 0: Exit Function
    
    If glngModel = PM_门诊医嘱清单 Then
        gobjPati.lng挂号ID = rsPati!就诊Id
    End If
    
    '附加信息
    strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                    "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                    "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
    Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, gobjPati.lng病人ID, gobjPati.lng挂号ID)
    rsPatiInfo.Filter = "项目名称='身高'"
    If rsPatiInfo.RecordCount <> 0 Then str身高 = NVL(rsPatiInfo!记录内容)
    rsPatiInfo.Filter = "项目名称='体重'"
    If rsPatiInfo.RecordCount <> 0 Then str体重 = NVL(rsPatiInfo!记录内容)
            
    With udtPati
        .str姓名 = rsPati!姓名
        .str出生日期 = rsPati!出生日期 & ""
        .str性别 = rsPati!性别 & ""
        .str体重 = str身高
        .str身高 = str体重
        .str身份证号 = rsPati!身份证号 & ""
        .str病历卡号 = gobjPati.lng挂号ID
        .str卡号 = ""
        .str卡类型 = ""
        .str怀孕时间 = ""
        .str怀孕时间单位 = ""
        '过敏源
        Set colTmp = New Collection
        Set rsTmp = Get病人过敏记录(gobjPati.lng病人ID, 0, 1)
        For i = 1 To rsTmp.RecordCount
            If "" & rsTmp!药物ID <> "" Then
                With udt过敏源
                    .str过敏类型 = "5"   '1=药卫士药品大类 2=药卫士药品成份 5-传his药品
                    .str过敏源名称 = rsTmp!药物名
                    .str过敏源代码 = "" & rsTmp!药物ID
                End With
                colTmp.Add udt过敏源, "_" & i
            End If
            rsTmp.MoveNext
        Next
        Set .col过敏源s = colTmp
        
        '诊断记录
        Set colTmp = New Collection
        If glngModel = PM_门诊编辑 Then
            If Not gobjDiags Is Nothing Then
                For i = 1 To gobjDiags.Count
                    With udt诊断
                        If gobjDiags.Item(i).str诊断描述 <> "" Then
                            If gobjDiags.Item(i).str疾病编码 <> "" Then
                                .str诊断类型 = "2" '2=IDC10代码
                                .str诊断代码 = gobjDiags.Item(i).str疾病编码
                            Else
                                .str诊断类型 = "0"
                                .str诊断代码 = gobjDiags.Item(i).str诊断编码
                            End If
                            .str诊断名称 = gobjDiags.Item(i).str诊断描述
                        End If
                    End With
                    colTmp.Add udt诊断, "_" & colTmp.Count + 1
                Next
            End If
        Else
            Set rsTmp = Get病人诊断记录(gobjPati.lng病人ID, gobjPati.lng挂号ID, "1,11")
            For i = 1 To rsTmp.RecordCount
                With udt诊断
                    If rsTmp!疾病ID & "" <> "" Then
                         .str诊断类型 = "2" '2=IDC10代码
                    Else
                        .str诊断类型 = "0" '0=其他
                    End If
                    .str诊断代码 = "" & rsTmp!编码
                    .str诊断名称 = "" & rsTmp!名称
                End With
                colTmp.Add udt诊断, "_" & colTmp.Count + 1
                rsTmp.MoveNext
            Next
        End If
        '病生理
        strTmp = Get病人病生理情况(gobjPati.lng病人ID, 0)
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                With udt诊断
                    .str诊断类型 = "1" '1=病生理状态
                    .str诊断代码 = Sys.RowValue("病生理情况", arrTmp(i), "编码", "名称")
                    .str诊断名称 = arrTmp(i)
                End With
                colTmp.Add udt诊断, "_" & colTmp.Count + 1
            Next
           
        End If
        Set .col诊断s = colTmp
    End With
    
    curDate = zlDatabase.Currentdate
    With udtDetail
        .strHIS系统时间 = Format(curDate, "YYYY-MM-DD HH:MM:SS")
        .str门诊住院标识 = "op"    '住院标识
    
        .str就诊类型 = YWS_GetTreatType(1, gobjPati.lng挂号ID)
        .str就诊号 = gobjPati.lng挂号ID
        .str床位号 = ""
        '病人信息
        .udt病人信息 = udtPati
        '处方信息
        .udt处方信息 = udtPres
    End With
    '药品信息
    Set colTmp = New Collection
    
    arrTmp = Array()
    With gobjAdvice
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_门诊编辑 Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                    And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                    And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                    
            Else
                blnDo = (Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                    Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4")) _
                    And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            End If
            blnDo = blnDo And Val(.TextMatrix(i, gobjCOL.intCOL状态)) <> 4 '作废的医嘱不传入
            If blnDo Then
                If glngModel = PM_门诊医嘱清单 And .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4" Then
                    '获取中药医嘱组ID
                    str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    '取药品名称
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                        str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                    Else
                        str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                    End If
                    '取药品给药途径
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str给药途径 = "" '一并给药不重复取
                    If str给药途径 = "" Then
                        If glngModel = PM_门诊编辑 Then
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                            If k <> -1 Then str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                        Else
                            str给药途径 = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOL相关ID)), "诊疗项目ID")  '传代码
                        End If
                    End If
                    
                    Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), 0, 0, "", IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), str频率编码)
                    
                    If glngModel = PM_门诊编辑 Then
                        str医嘱ID = .RowData(i)
                         '单量，单量单位
                        str单量 = .TextMatrix(i, gobjCOL.intCOL单量)
                        str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                        str单量 = Replace(str单量, str单量单位, "")
                        str开嘱时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd HH:mm:ss")        '处方时间（YYYY-MM-DD HH:mm:SS）
                    Else
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                        str单量 = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL单量)))
                        If Mid(str单量, 1, 2) = "0." Then
                            str单量 = "0" & Val(str单量)
                        Else
                            str单量 = Val(str单量)
                        End If
                        str单量单位 = Trim(StrToXML(.TextMatrix(i, gobjCOL.intCOL单量)))
                        If Mid(str单量单位, 1, 2) = "0." Then '单量有小数点的特殊处理
                            str单量单位 = Replace(str单量单位, Format(Val(str单量单位) & "", "0.####"), "") '门诊清单单量（“单量” & “单量单位”）
                        Else
                            str单量单位 = Replace(str单量单位, Val(str单量单位) & "", "")    '
                        End If
                        
                        str开嘱时间 = Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd HH:mm:ss")          '处方时间（YYYY-MM-DD HH:mm:SS）
                    End If
                    udtMedic.str药品类型 = YWS_GetDrugType(.TextMatrix(i, gobjCOL.intCOL诊疗类别))       '西药/中成药/草药
                    udtMedic.str处方号 = str医嘱ID    '传医嘱ID
                    udtMedic.Str医嘱类型 = "T"
                    udtMedic.str处方时间 = str开嘱时间
                    udtMedic.str商品名 = YWS_StrToXML(Sys.RowValue("诊疗项目目录", Val(.TextMatrix(i, gobjCOL.intCOL诊疗项目ID)), "名称"))
                    udtMedic.str医院药品代码 = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                    udtMedic.str医保代码 = ""
                    udtMedic.str批准文号 = ""
                    udtMedic.str用药开始时间 = ""
                    udtMedic.str用药结束时间 = ""
                    udtMedic.str规格 = ""
                    udtMedic.str组号 = .TextMatrix(i, gobjCOL.intCOL相关ID)
                    udtMedic.str用药理由 = ""
                    udtMedic.str单次量单位 = str单量单位
                    udtMedic.str单次量 = str单量
                    udtMedic.str频次代码 = str频率编码
                    udtMedic.str给药途径代码 = str给药途径
                    
                    udtMedic.str服药天数 = .TextMatrix(i, gobjCOL.intCOL天数)   'OP 门诊处方有效
            
                    colTmp.Add udtMedic, "_" & colTmp.Count + 1
                End If
            End If
        Next
        '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
        If glngModel = PM_门诊医嘱清单 Then
            If str中药组IDs <> "" Then
                Set rs中药 = Get中药配方(str中药组IDs)
                With rs中药
                    For i = 1 To .RecordCount
                        If !相关ID & "" <> str相关ID Then
                            Call Get频率信息_名称(!频率 & "", 0, 0, "", IIf(!诊疗类别 & "" = "7", 2, 1), str频率编码)
                            str相关ID = !相关ID & ""
                        End If
                        udtMedic.str药品类型 = YWS_GetDrugType(!诊疗类别 & "")
                        udtMedic.str处方号 = !id & ""    '传医嘱ID
                        udtMedic.Str医嘱类型 = "T"
                        udtMedic.str处方时间 = Format(!开嘱时间 & "", "yyyy-MM-dd HH:mm:ss")
                        udtMedic.str商品名 = !药品名称 & ""
                        udtMedic.str医院药品代码 = !药品ID & ""
                        udtMedic.str医保代码 = ""
                        udtMedic.str批准文号 = ""
                        udtMedic.str用药开始时间 = ""
                        udtMedic.str用药结束时间 = ""
                        udtMedic.str规格 = ""
                        udtMedic.str组号 = !相关ID & ""
                        udtMedic.str用药理由 = ""
                        udtMedic.str单次量单位 = !单量单位 & ""
                        udtMedic.str单次量 = !单次用量 & ""
                        udtMedic.str频次代码 = str频率编码
                        udtMedic.str给药途径代码 = !用法ID & ""
                        udtMedic.str服药天数 = !天数 & ""   'OP 门诊处方有效
                
                        colTmp.Add udtMedic, "_" & colTmp.Count + 1
                        .MoveNext
                    Next
                End With
            End If
        End If
        With udtPres
            Set .col药品信息 = colTmp
            .str处方号 = "0"
            .str处方理由 = ""
            .str处方时间 = Format(curDate, "YYYY-MM-DD HH:MM:SS")
            .str是否当前处方 = "1" '0 历史处方 1 当前处方（现是默认当前处方，以后扩充）
            .Str医嘱类型 = "T"  '门诊临时
        End With
    End With
    
    udtDetail.udt处方信息 = udtPres
    
    OutAdviceCheckWarn_YWS = True
    If udtPres.col药品信息.Count > 0 Then
        strXML = YWS_MakePresXML(udtDetail)
        WriteLog "" & glngModel, "OutAdviceCheckWarn_YWS", strXML
        'YWS_处方分析=6 处方分析返回值：0、1、2，3，8分别代表没有问题，其他问题，一般问题，严重问题，有写申诉理由 保存前调用
        strTmp = gobjPass.YWS_UI(YWS_处方分析, gstrBaseXml, strXML, strRetXML)
        WriteLog "" & glngModel, "OutAdviceCheckWarn_YWS", "返回值:" & strTmp & vbCrLf & strRetXML
        '设置警示灯
        Set rsRet = YWS_ReturnRS(strRetXML)
        With gobjAdvice
            arrSQL = Array()
            '图片下标：1-蓝灯,2-黄灯,5-橙灯,3-红灯
            '警示级顺：0-蓝灯,1-黄灯,2-橙灯,3-红灯
            arrLight(0) = "蓝": arrLight(1) = "黄": arrLight(2) = "橙": arrLight(3) = "红"
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_门诊编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                    blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                    And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                End If
                If blnDo Then
                    If glngModel = PM_门诊编辑 Then
                        str医嘱ID = .RowData(i)
                    Else
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                    End If
                    '取药嘱最大警示值
                    rsRet.Filter = "医嘱ID ='" & str医嘱ID & "'"
                    If rsRet.RecordCount > 0 Then
                        k = CLng(rsRet!警示值 & "")
                    Else
                        k = 0
                    End If
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        '设置警示灯
                        If k >= 0 And k <= 3 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(k)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        
                        If glngModel = PM_门诊编辑 Then
                            '标记审查结果变化,以备更新数据库
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                blnNoSave = True    '标记为未保存
                            End If
                        ElseIf PM_门诊医嘱清单 = glngModel Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 3, k, "NULL") & ")"
                            End If
                        End If
                    Else
                        '中药配方
                        If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                            lng中药组ID = .TextMatrix(i, gobjCOL.intCOL相关ID)          '中药配方组ID
                            lngLight = -1 '初始化
                        End If
                        '设置警示灯 取草药中最大警示值
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If k > lngLight Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                    End If
                Else
                    If glngModel = PM_门诊编辑 Then
                        '中药警示灯单独设置
                        If .RowData(i) = lng中药组ID And .RowData(i) <> 0 Then
                            strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                            '设置警示灯
                            If lngLight >= 0 And lngLight <= 4 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                            Else
                                .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                                Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                            End If
                            
                            If glngModel = PM_门诊编辑 Then
                                '标记审查结果变化,以备更新数据库
                                If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                    .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                    blnNoSave = True    '标记为未保存
                                End If
                            End If
                            lng中药组ID = 0
                            lngLight = -1
                        End If
                    End If
                End If
            Next
            '医嘱清单中药配方警示灯处理
            If glngModel = PM_门诊医嘱清单 And Not rs中药 Is Nothing Then
                For i = .FixedRows To .Rows - 1
                    '中药服法
                    If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        lngLight = -1
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                        rs中药.Filter = "相关ID=" & str医嘱ID
                        
                        For j = 1 To rs中药.RecordCount
                            '取药嘱最大警示值
                            rsRet.Filter = "医嘱ID ='" & rs中药!id & "'"
                            If rsRet.RecordCount > 0 Then
                                k = CLng(rsRet!警示值 & "")
                            Else
                                k = 0
                            End If
                            '设置警示灯 取草药中最大警示值
                            If k >= 0 Then
                                If lngLight >= 0 Then
                                    If k > lngLight Then
                                        lngLight = k
                                    End If
                                Else
                                    lngLight = k
                                End If
                            End If
                            rs中药.MoveNext
                        Next
                        
                        '设置警示灯
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        '警示灯更新到数据库
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(lngLight >= 0 And lngLight <= 3, lngLight, "NULL") & ")"
                        End If
                    End If
                Next
            End If
            For i = LBound(arrSQL) To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
            Next
        End With

        If glngModel = PM_门诊编辑 Then
            If strTmp = "3" And gbytBlackLamp = 0 Then
                MsgBox "用药监测系统发现当前医嘱存在禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                OutAdviceCheckWarn_YWS = False
                Exit Function
            ElseIf strTmp = "3" And gbytBlackLamp = 1 Then
                If MsgBox("用药监测系统发现当前医嘱存在禁忌用药，是否继续?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then OutAdviceCheckWarn_YWS = False: Exit Function
            End If
            
            '上传处方
            If strTmp <> "0" Then
                strTmp = gobjPass.YWS_UI(YWS_上传处方, gstrBaseXml, strXML, strRetXML)
                WriteLog "" & glngModel, "OutAdviceCheckWarn_YWS", "保存处方:返回值" & strTmp
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    OutAdviceCheckWarn_YWS = False
End Function


Private Function ReadPatient(ByVal lngPatiID As Long, ByVal strNo As String) As ADODB.Recordset
    Dim strSQL As String
    strSQL = "Select b.ID as 就诊ID,B.姓名,B.性别,A.出生日期,A.年龄,A.身份证号 " & _
         " From 病人信息 A,病人挂号记录 B " & _
         " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
    On Error GoTo errH
    Set ReadPatient = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, strNo)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Public Function AdviceCheckWarn_MK_YF(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, _
            ByVal lngCmd As Long, Optional ByVal lngCurrID As Long, Optional ByVal str医嘱IDs As String, _
            Optional str警示 As String) As Long
'功能：调用Pass系统相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        21-病生状态/过敏史管理(只读)
'        1/33-保存自动审查(住院/门诊),2/34-提交自动审查(住院/门诊),3-手工调用审查
'        6=单药警告
'      lngCurrID=当前药品医嘱的行号，lngCmd=0时需要
'      str医嘱IDs 医嘱ID串：ID1,ID2,ID3...
'      str警示-医嘱ID:警示值,医嘱ID2:警示值2
'返回：检测PASS菜单时，返回>=0表示可以弹出菜单,其它返回-1
'说明：用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset, rs规格 As ADODB.Recordset, rs开嘱医生 As ADODB.Recordset
    Dim str药品 As String, str用法 As String, str单量单位 As String, str频率 As String
    Dim str开嘱医生 As String, str开嘱医生串 As String, str药品ID As String
    Dim strSQL As String, i As Long, k As Long
    Dim lng标识号  As Long
    Dim lngCount As Long
    Dim blnDo As Boolean
    Dim strCurrentDate As String

    AdviceCheckWarn_MK_YF = -1

    On Error GoTo errH
    Screen.MousePointer = 11

    '检验PASS可用状态
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '114036同一个病人多次审查时病人信息每次都要传入
    '-------------------------------------------------------------
    Set rsTmp = GetPatiInfo_YF(lng病人ID, str挂号单, lng主页ID)
    If str挂号单 <> "" Then               '门诊病人
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        Call PassSetPatientInfo(lng病人ID, rsTmp!就诊Id, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
            rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), "")
        lng标识号 = NVL(rsTmp!就诊Id, 0)
    Else                                    '住院病人
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        Call PassSetPatientInfo(lng病人ID, lng主页ID, rsTmp!姓名, NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
            rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), NVL(rsTmp!医生码) & "/" & NVL(rsTmp!医生名), ""), _
            IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))
        lng标识号 = lng主页ID
    End If
    '传人病人过敏史
    '-------------------------------------------------------
    Set rsTmp = Get病人过敏记录(lng病人ID, IIf(str挂号单 <> "", 0, lng主页ID))

    For i = 1 To rsTmp.RecordCount
        Call PassSetAllergenInfo(i, rsTmp!药物ID & "", rsTmp!药物名 & "", "DrugName", "")
        rsTmp.MoveNext
    Next

    '传人病生状态
    '------------------------------------------------------------------
    Set rsTmp = Get病人诊断记录(lng病人ID, lng标识号, IIf(str挂号单 <> "", "1,11", "2,12"))
    strCurrentDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")

    For i = 1 To rsTmp.RecordCount
        Call PassSetMedCond(i & "", rsTmp!编码 & "", rsTmp!名称 & "", "User", strCurrentDate, strCurrentDate)
        rsTmp.MoveNext
    Next

    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = MK_检测PASS菜单状态 Then
        If lngCurrID = 0 Then: Exit Function
        strSQL = "Select Nvl(a.标本部位, a.医嘱内容) As 药品名称,a.诊疗项目id,a.收费细目id As 药品id, c.计算单位 As 单量单位, b.医嘱内容 As 用法" & vbNewLine & _
                "From 病人医嘱记录 A, 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "Where a.Id = [1] And a.诊疗项目id = c.Id And a.相关id = b.Id(+)"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngCurrID)
        
        If rsTmp.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
        '取药品名称
        str药品 = rsTmp!药品名称 & ""
        str药品ID = rsTmp!药品ID & ""
        str单量单位 = rsTmp!单量单位 & ""
        If str药品ID = "" Then str药品ID = GetDrugID(rsTmp!诊疗项目ID & "")
        '取药品给药途径
        str用法 = rsTmp!用法 & ""
        '传入查询药品信息
        Call PassSetQueryDrug(str药品ID, str药品, str单量单位, str用法)
    
        AdviceCheckWarn_MK_YF = 1 '表示可以弹出菜单

        Screen.MousePointer = 0: Exit Function
    ElseIf lngCmd = MK_单药警告 Then
        Call PassSetWarnDrug(lngCurrID)    '单药警告(已警告的医嘱唯一码)
    ElseIf lngCmd = MK_病生状态过敏史查看 Then
        Call PassDoCommand(lngCmd)  '21-查看过敏史
        Screen.MousePointer = 0
        Exit Function
    Else
        Set rsTmp = GetAdviceInfo_YF(lng病人ID, lng主页ID, str挂号单, str医嘱IDs)
        If rsTmp.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
        
        '用药审核或用药研究
        With rsTmp
            lngCount = 0
            str药品 = "": str开嘱医生串 = ""
            For i = 1 To .RecordCount
                If Val(!收费细目id & "") = 0 Then
                    str药品 = str药品 & "," & !诊疗项目ID
                End If
                '获取开嘱医生
                If NVL(!开嘱医生) <> "" Then
                    str开嘱医生 = NVL(!开嘱医生)
                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                    If InStr("," & str开嘱医生串 & ",", "," & str开嘱医生 & ",") = 0 And str开嘱医生 <> "" Then
                        str开嘱医生串 = str开嘱医生串 & "," & str开嘱医生
                    End If
                End If
                .MoveNext
            Next
            
            If str药品 <> "" Then
                Set rs规格 = GetDrugID(str药品)
            End If
            
            If str开嘱医生串 <> "" Then
                str开嘱医生串 = Mid(str开嘱医生串, 2)
                Set rs开嘱医生 = Sys.RowValue("人员表", str开嘱医生串, "编号,姓名", "姓名")
            End If
            
            str开嘱医生串 = "": str药品ID = ""
            .MoveFirst
            
            For i = 1 To .RecordCount
                '取用药频率(次/天),都为整数四舍五入
                str频率 = GetFrequency(!间隔单位 & "", !频率次数 & "", !频率间隔 & "")
            
                '长期医嘱按品种下达时,取任意药品ID
                If Val(!收费细目id & "") = 0 Then
                    rs规格.Filter = "药名ID =" & !诊疗项目ID
                    If Not rs规格.EOF Then str药品ID = rs规格!药品ID & ""
                Else
                    str药品ID = !收费细目id & ""
                End If
                '开嘱医生
                str开嘱医生 = NVL(!开嘱医生)
                If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                
                If str开嘱医生 <> Mid(str开嘱医生串, InStr(str开嘱医生串, "/") + 1) Then
                    If Not rs开嘱医生 Is Nothing Then
                        rs开嘱医生.Filter = "姓名='" & str开嘱医生 & "'"
                        If Not rs开嘱医生.EOF Then str开嘱医生串 = rs开嘱医生!编号 & "/" & str开嘱医生
                    End If
                End If
                
                '传入医嘱信息
                Call PassSetRecipeInfo(!医嘱ID & "", str药品ID, !药品名称 & "", !单次用量 & "", !单量单位 & "", str频率, _
                    Format(!开始时间 & "", "yyyy-MM-dd"), Format(!结束时间 & "", "yyyy-MM-dd"), !用法 & "", _
                    !相关ID & "", !医嘱期效 & "", str开嘱医生串)

                lngCount = lngCount + 1

                .MoveNext
            Next
    
            '无可审查的药品
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End With
    End If

    '执行相应的命令
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    
    If str警示 <> "-1" And lngCount > 0 Then
        str警示 = ""
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            k = PassGetWarn(rsTmp!医嘱ID & "")
            str警示 = str警示 & "," & rsTmp!医嘱ID & ":" & k
            rsTmp.MoveNext
        Next
        If str警示 <> "" Then str警示 = Mid(str警示, 2)
    End If
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AdviceCheckWarn_MK4_YF(ByVal lngPatiID As Long, ByVal lng主页ID As Long, Optional ByVal str挂号单 As String, _
                Optional ByVal str医嘱IDs As String, Optional str警示 As String, Optional rsRet As ADODB.Recordset) As Long
'功能：调用Pass4系统审查功能
'参数:
'    str医嘱IDs-需要审查的医嘱ID
'    str警示-医嘱ID:警示值,医嘱ID2:警示值2
'返回：暂无意义
'   rsRet-门诊发送时返回审核状态
'说明：
'
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim strSQL As String, i As Long, k As Long
    Dim str开嘱医生 As String
    Dim str开嘱医生编码 As String
    Dim str医生 As String
    Dim str用药目的 As String
    Dim bytSubmit As Byte
    
    AdviceCheckWarn_MK4_YF = -1

    On Error GoTo errH
    Screen.MousePointer = 11
    '药嘱信息提取
    If glngModel = PM_门诊发送 Then
        Set rsTmp = GetAdviceInfo_YF(lngPatiID, lng主页ID, str挂号单, str医嘱IDs, 2)
        bytSubmit = 1
    Else
        Set rsTmp = GetAdviceInfo_YF(lngPatiID, lng主页ID, str挂号单, str医嘱IDs)
        bytSubmit = 0
    End If
            
    If rsTmp.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
    '
    Set rsAdvice = InitAdviceRS(FUN_医嘱信息)
    
    With rsTmp
        For i = 1 To .RecordCount
            rsAdvice.AddNew
            rsAdvice!医嘱ID = !医嘱ID & ""
            rsAdvice!相关ID = !相关ID & ""
            rsAdvice!医嘱期效 = !医嘱期效 & ""
            rsAdvice!医嘱序号 = !医嘱序号 & ""
            rsAdvice!天数 = !天数 & ""
            If glngModel = PM_处方发药 Or glngModel = PM_部门发药 Or glngModel = PM_PIVA管理 Then
                '"0"-在用（默认）；"1"-已作废；"2"-已停嘱；"3"-离院带药（根据系统设置参与审查）
                rsAdvice!医嘱状态 = "0"
            ElseIf glngModel = PM_门诊发送 Then
                rsAdvice!审核状态 = 1
                rsAdvice!医嘱状态 = IIf(!医嘱状态 & "" = 1, "0", "-1")
            End If
            '
            str开嘱医生 = !开嘱医生 & ""
            If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
            If str医生 <> str开嘱医生 And str开嘱医生 <> "" Then
                str医生 = str开嘱医生   '多条医嘱同一个开嘱医生,只需访问一次
                str开嘱医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
            End If
            
            rsAdvice!开嘱科室 = !开嘱科室 & ""
            rsAdvice!开嘱科室id = !开嘱科室id & ""
            
            rsAdvice!开嘱医生编码 = str开嘱医生编码
            rsAdvice!开嘱医生 = str开嘱医生
            rsAdvice!药品ID = !收费细目id & ""
            
            rsAdvice!药品名称 = !药品名称 & ""
            rsAdvice!单次用量 = FormatEx(NVL(!单次用量), 5)
            rsAdvice!单量单位 = !单量单位 & ""
            rsAdvice!频率 = !频率 & ""
            rsAdvice!用法 = !用法 & ""
            rsAdvice!用法ID = !用法ID & ""    '美康4.0用法ID传用法名称
            '临时医嘱开始时间和结束时间相同
            If !医嘱期效 & "" = "0" Then  '长嘱
                rsAdvice!结束时间 = Format(!结束时间 & "", "YYYY-MM-dd hh:mm:ss")
            Else '临时医嘱
                rsAdvice!结束时间 = Format(!开嘱时间 & "", "YYYY-MM-dd hh:mm:ss")
            End If
            
            rsAdvice!开嘱时间 = Format(!开嘱时间 & "", "YYYY-MM-dd hh:mm:ss")
            rsAdvice!开始时间 = Format(!开始时间 & "", "YYYY-MM-dd hh:mm:ss")
            
            If str挂号单 <> "" Then
                '总量
                If InStr(",5,6,", "," & !诊疗类别 & ",") > 0 Then
                    '成药临嘱有总量,以零售单位存放,门诊单位显示
                    If Not IsNull(!总量) And Not IsNull(!门诊包装) Then
                        rsAdvice!总量 = FormatEx(!总量 / !门诊包装, 5)
                    End If
                End If
                rsAdvice!总量单位 = !门诊单位 & ""
            Else
                rsAdvice!总量 = !总量 & ""
                rsAdvice!总量单位 = !住院单位 & ""
            End If
            
            '用药目的(0默认, 1可能预防，2可能治疗，3预防，4治疗，5预防+治疗)
            str用药目的 = !用药目的 & ""
            If str用药目的 = "1" Then
                str用药目的 = "3"
            ElseIf str用药目的 = "2" Then
                str用药目的 = "4"
            Else
                str用药目的 = "0"
            End If
            rsAdvice!用药目的 = str用药目的
            rsAdvice!医生嘱托 = !医生嘱托 & ""
            rsAdvice!诊疗类别 = !诊疗类别 & ""
            rsAdvice!处方号 = !处方号 & ""
            rsAdvice!执行科室ID = !执行科室ID & ""
            rsAdvice.Update
            .MoveNext
        Next
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    End With
    Call AdviceCheckWarn_MK4(lngPatiID, str挂号单, lng主页ID, 1, bytSubmit, rsAdvice, str警示)   '显示审查界面,不采集数据
    If glngModel = PM_门诊发送 Then
        Set rsRet = rsAdvice
    End If
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetFrequency(ByVal str间隔单位 As String, ByVal str频率次数 As String, ByVal str频率间隔 As String) As String
'功能:构造频率字符串
    Dim str频率 As String
    
    If str间隔单位 = "天" Then
        str频率 = str频率次数 & "/" & str频率间隔
    ElseIf str间隔单位 = "周" Then
        str频率 = str频率次数 & "/7"
    ElseIf str间隔单位 = "小时" Then
        If Val(str频率间隔) <= 24 Then
            str频率 = Format(24 / Val(str频率间隔) * Val(str频率次数), "0") & "/1"
        Else
            str频率 = Val(str频率次数) & "/" & Format(Val(str频率间隔) / 24, "0")
        End If
    ElseIf str间隔单位 = "分钟" Then
        str频率 = Format((24 * 60) / Val(str频率间隔) * Val(str频率次数), "0") & "/1"
    End If
    GetFrequency = str频率
End Function

Public Function AdviceCheckWarn_YWS_YF(ByVal lngPatiID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, _
            Optional ByVal str医嘱IDs As String, Optional str警示 As String) As Boolean
'功能：调用保进药卫士用药监测系统对医嘱进行合理用药审查等相关功能
    Dim udtDetail As YWS_DETAILS
    Dim udtPati As YWS_PATIENT
    Dim colTmp As Collection
    Dim udt过敏源 As YWS_ALLERGIC
    Dim udt诊断 As YWS_DIAGNOSE
    Dim udtPres As YWS_PRESCRIPTION
    Dim udtMedic As YWS_MEDICINE   '药品信息
    
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Long, lngFunc As Long
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim arrTmp As Variant, curDate As Date
    
    Dim rsTmp As ADODB.Recordset
    Dim rsPati As ADODB.Recordset, rsRet As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset

    Dim str身高 As String, str体重 As String, str频率编码 As String, strXML As String
    Dim str期效 As String
    Dim lng挂号ID As String
    Dim str病历卡号 As String
    Dim str就诊类型 As String
    

    Dim k As Long, blnDo As Boolean

    Dim strRetXML As String
    Dim blnIsHaveOut As Boolean '判断是否存在院外执行的药品
    '获取病人信息
    If str挂号单 <> "" Then
        Set rsPati = ReadPatient(lngPatiID, str挂号单)
        If rsPati.RecordCount = 0 Then Exit Function
        lng挂号ID = Val(rsPati!就诊Id & "")
        strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                        "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                        "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
                        
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, lng挂号ID)
        rsPatiInfo.Filter = "项目名称='身高'"
        If rsPatiInfo.RecordCount <> 0 Then str身高 = NVL(rsPatiInfo!记录内容)
        rsPatiInfo.Filter = "项目名称='体重'"
        If rsPatiInfo.RecordCount <> 0 Then str体重 = NVL(rsPatiInfo!记录内容)
        str病历卡号 = lng挂号ID & ""
    Else
        Set rsPati = GetPatiInfo(lngPatiID, lng主页ID)
        If rsPati.RecordCount = 0 Then Exit Function
        str身高 = rsPati!身高 & ""
        str体重 = rsPati!体重 & ""
        str病历卡号 = rsPati!住院号 & ""
    End If
    
    With udtPati
        .str姓名 = rsPati!姓名
        .str出生日期 = rsPati!出生日期 & ""
        .str性别 = rsPati!性别 & ""
        .str体重 = str身高
        .str身高 = str体重
        .str身份证号 = rsPati!身份证号 & ""
        .str病历卡号 = str病历卡号
        .str卡号 = ""
        .str卡类型 = ""
        .str怀孕时间 = ""
        .str怀孕时间单位 = ""
        '过敏源
        Set colTmp = New Collection
        Set rsTmp = Get病人过敏记录(lngPatiID, lng主页ID, 1)
        For i = 1 To rsTmp.RecordCount
            If "" & rsTmp!药物ID <> "" Then
                With udt过敏源
                    .str过敏类型 = "5"   '1=药卫士药品大类 2=药卫士药品成份
                    .str过敏源名称 = rsTmp!药物名
                    .str过敏源代码 = "" & rsTmp!药物ID
                End With
                colTmp.Add udt过敏源, "_" & i
            End If
            rsTmp.MoveNext
        Next
        Set .col过敏源s = colTmp
        
        '诊断记录
        Set colTmp = New Collection
        Set rsTmp = Get病人诊断记录(lngPatiID, IIf(str挂号单 <> "", lng挂号ID, lng主页ID), IIf(str挂号单 <> "", "1,11", "2,12"))
        For i = 1 To rsTmp.RecordCount
            With udt诊断
                If rsTmp!疾病ID & "" <> "" Then
                     .str诊断类型 = "2" '2=IDC10代码
                Else
                    .str诊断类型 = "0" '0=其他
                End If
                .str诊断代码 = "" & rsTmp!编码
                .str诊断名称 = "" & rsTmp!名称
            End With
            colTmp.Add udt诊断, "_" & colTmp.Count + 1
            rsTmp.MoveNext
        Next
        '病生理
        strTmp = Get病人病生理情况(lngPatiID, IIf(str挂号单 <> "", 0, lng主页ID))
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                With udt诊断
                    .str诊断类型 = "1" '1=病生理状态
                    .str诊断代码 = Sys.RowValue("病生理情况", arrTmp(i), "编码", "名称")
                    .str诊断名称 = arrTmp(i)
                End With
                colTmp.Add udt诊断, "_" & colTmp.Count + 1
            Next
           
        End If
        Set .col诊断s = colTmp
    End With
    
    curDate = zlDatabase.Currentdate
    
    With udtDetail
        .strHIS系统时间 = Format(curDate, "YYYY-MM-DD HH:MM:SS")
        If str挂号单 <> "" Then
            .str门诊住院标识 = "op"  '门诊住院标识
            .str就诊类型 = YWS_GetTreatType(1, lng挂号ID)
            .str就诊号 = lng挂号ID & ""
            .str床位号 = ""
        Else
            .str门诊住院标识 = "ip" '门诊住院标识
            .str就诊类型 = YWS_GetTreatType(2, lngPatiID, lng主页ID)
            .str就诊号 = rsPati!住院号 & ""
            .str床位号 = "" & rsPati!当前床号
        End If
        
        '病人信息
        .udt病人信息 = udtPati
        '处方信息
        .udt处方信息 = udtPres
    End With
    '药品信息
    Set colTmp = New Collection
    
    Set rsTmp = GetAdviceInfo_YF(lngPatiID, lng主页ID, str挂号单, str医嘱IDs)
    If rsTmp.RecordCount = 0 Then Exit Function
    With rsTmp
        For i = 1 To rsTmp.RecordCount

            Call Get频率信息_名称(!频率 & "", 0, 0, "", IIf(!诊疗类别 & "" = "7", 2, 1), str频率编码)
            udtMedic.str药品类型 = YWS_GetDrugType(!诊疗类别 & "")
            udtMedic.str处方号 = !医嘱ID & ""    '传医嘱ID
            udtMedic.Str医嘱类型 = IIf(!医嘱期效 & "" = "长嘱", "L", "T")
            udtMedic.str处方时间 = Format(!开嘱时间 & "", "yyyy-MM-dd HH:mm:ss")          '处方时间（YYYY-MM-DD HH:mm:SS）
            udtMedic.str商品名 = YWS_StrToXML(!药品名称 & "")
            udtMedic.str医院药品代码 = !收费细目id & ""
            udtMedic.str医保代码 = ""
            udtMedic.str批准文号 = ""
            udtMedic.str用药开始时间 = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")
            udtMedic.str用药结束时间 = Format(!结束时间 & "", "yyyy-MM-dd HH:mm:ss")
            udtMedic.str规格 = ""
            udtMedic.str组号 = !相关ID & ""
            udtMedic.str用药理由 = ""
            '单量，单量单位
            udtMedic.str单次量单位 = !单量单位 & ""
            udtMedic.str单次量 = !单次用量 & ""
            udtMedic.str频次代码 = str频率编码
            udtMedic.str给药途径代码 = !用法ID & ""
            udtMedic.str服药天数 = !天数 & ""   'OP 门诊处方有效
            
            colTmp.Add udtMedic, "_" & colTmp.Count + 1
            .MoveNext
        Next
        
        With udtPres
            Set .col药品信息 = colTmp
            .str处方号 = "0"
            .str处方理由 = ""
            .str处方时间 = Format(curDate, "YYYY-MM-DD HH:MM:SS")
            .str是否当前处方 = "1" '0 历史处方 1 当前处方（现是默认当前处方，以后扩充）
            .Str医嘱类型 = IIf(str挂号单 <> "", "T", "L")
            
        End With
    End With
    
    udtDetail.udt处方信息 = udtPres
    
    AdviceCheckWarn_YWS_YF = True
    If udtPres.col药品信息.Count > 0 Then
        On Error GoTo errH
        strXML = YWS_MakePresXML(udtDetail)
        WriteLog "" & glngModel, "AdviceCheckWarn_YWS_YF", strXML
        If glngModel = PM_PIVA管理 And str医嘱IDs <> "" Then
            lngFunc = YWS_处方分析仅亮灯
        Else
            lngFunc = YWS_处方分析
        End If
        strTmp = gobjPass.YWS_UI(lngFunc, gstrBaseXml, strXML, strRetXML)
        WriteLog "" & glngModel, "AdviceCheckWarn_YWS_YF", "返回值:" & strTmp & vbCrLf & strRetXML
        If lngFunc = YWS_处方分析仅亮灯 Then
            '不管问题级别都允许弹出问题详情窗口
            Set rsRet = YWS_ReturnRS(strRetXML, 1)
            If rsRet.RecordCount > 0 Then
                frmPassResultYWS.ShowMe rsRet
            End If
        End If
        If str警示 <> "-1" Then
            Set rsRet = YWS_ReturnRS(strRetXML)
            str警示 = ""
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rsRet.Filter = "医嘱ID ='" & rsTmp!医嘱ID & "" & "'"
                If rsRet.RecordCount > 0 Then
                    k = CLng(rsRet!警示值 & "")
                Else
                    k = 0
                End If
                str警示 = str警示 & "," & rsTmp!医嘱ID & ":" & k
                rsTmp.MoveNext
            Next
            If str警示 <> "" Then str警示 = Mid(str警示, 2)
        Else
            str警示 = ""
        End If
        
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    AdviceCheckWarn_YWS_YF = False
End Function

Private Function CheckAdvice_YF(ByVal strNo As String, ByVal int单据 As Integer, lngPatiID As Long, str挂号单 As String, lng主页ID As Long) As Boolean
'功能：检查病人是否允许
'参数:
'返回：T-返回病人ID,挂号单,主页ID (找到医嘱);F-未找医嘱
'
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnRet As Boolean
    
     '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就退出
    strSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] " & _
        " Union All " & _
        " Select distinct B.病人id,0 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,门诊费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] "
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, strNo, int单据)

    If rsTmp.RecordCount = 0 Then
        blnRet = False
    Else
        lngPatiID = rsTmp!病人ID
        str挂号单 = NVL(rsTmp!挂号单)
        lng主页ID = rsTmp!主页ID
        blnRet = True
    End If

    CheckAdvice_YF = blnRet
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiInfo_YF(ByVal lngPatiID As Long, ByVal str挂号单 As String, ByVal lng主页ID As Long) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If gbytPass = MK And gstrVersion = "3.0" Then
        If str挂号单 <> "" Then               '门诊病人
            strSQL = "Select B.ID as 就诊ID,B.姓名,B.性别,A.出生日期," & _
                     " C.编码 as 科室码,C.名称 as 科室名,E.编号 as 医生码,E.姓名 as 医生名" & _
                     " From 病人信息 A,病人挂号记录 B,部门表 C,人员表 E" & _
                     " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID" & _
                     " And B.执行人=E.姓名(+) And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
        Else                                    '住院病人
            strSQL = _
                " Select A.姓名,A.性别,A.出生日期,B.入院日期,B.出院日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
                " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
                " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
                " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[3]"
           
        End If
        
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        If str挂号单 <> "" Then
            strSQL = "Select A.门诊号,B.ID as 就诊ID,B.急诊,B.姓名,B.性别,A.出生日期," & _
                 " C.ID As 科室ID,C.名称 as 科室名,E.编号 as 医生码,E.姓名 as 医生名" & _
                 " From 病人信息 A,病人挂号记录 B,部门表 C,人员表 E" & _
                 " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID" & _
                 " And B.执行人=E.姓名(+) And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
        Else
            strSQL = _
                " Select Nvl(B.姓名,A.姓名) 姓名,Nvl(B.性别,A.性别) 性别,A.出生日期,A.住院号,B.身高,B.体重,B.入院日期,B.出院日期," & _
                         " C.ID as 科室ID,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
                         " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
                         " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID" & _
                         " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[3]"
        End If
    ElseIf gbytPass = TYT Then
        If str挂号单 <> "" Then
            strSQL = "Select b.ID as 就诊ID,A.门诊号,B.姓名,B.性别,A.出生日期,A.年龄,A.身份证号 " & _
            " From 病人信息 A,病人挂号记录 B " & _
            " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
        Else
            strSQL = _
            " Select A.住院号,Nvl(B.姓名,A.姓名) 姓名,Nvl(B.性别,A.性别) 性别 ,A.出生日期,B.身高,B.体重  " & _
                     " From 病人信息 A,病案主页 B" & _
                     " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[3]"
        End If
    ElseIf gbytPass = DT Then
        If str挂号单 <> "" Then
            strSQL = "Select b.ID as 就诊ID,A.门诊号,B.姓名,B.性别,A.出生日期,A.年龄,A.身份证号 " & _
            " From 病人信息 A,病人挂号记录 B " & _
            " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
        Else
            strSQL = "Select A.住院号, A.当前床号, A.出生日期, Nvl(B.姓名, A.姓名) 姓名, Nvl(B.性别, A.性别) 性别, Nvl(B.年龄, A.年龄) 年龄, A.门诊号, A.健康号,A.身份证号,B.身高,B.体重" & vbNewLine & _
                "From 病人信息 A, 病案主页 B" & vbNewLine & _
                "Where A.病人id = B.病人id And A.病人id = [1] And B.主页id = [3]"
        End If
    ElseIf gbytPass = HZYY Then
        If str挂号单 <> "" Then
            strSQL = "Select b.ID as 就诊ID,A.门诊号,B.急诊,B.姓名,B.性别,A.出生日期,A.年龄,A.身份证号,A.籍贯,A.民族,A.职业,A.婚姻状况,B.执行部门ID As 挂号科室ID,C.名称 As 挂号科室" & _
            ",A.手机号,A.家庭地址,D.编码 As 医疗付款方式, B.执行时间 As 就诊时间 " & _
            " From 病人信息 A,病人挂号记录 B,部门表 C, 医疗付款方式 D " & _
            " Where A.病人ID=B.病人ID And B.执行部门ID =C.ID(+) And B.医疗付款方式 = D.名称(+) And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
        Else
            strSQL = "Select A.住院号, A.当前床号, A.出生日期, Nvl(B.姓名, A.姓名) 姓名, Nvl(B.性别, A.性别) 性别, Nvl(B.年龄, A.年龄) As 年龄" & vbNewLine & _
                ", A.门诊号, A.健康号,A.身份证号,A.籍贯,A.民族,B.身高,B.体重,B.入院病区ID,D.名称 As 入院病区,NVL(B.入院病床,0) As 入院病床," & vbNewLine & _
                "B.入院科室ID,C.名称 As 入院科室,B.入院日期, E.编码 AS 医疗付款方式 " & vbNewLine & _
                "From 病人信息 A, 病案主页 B, 部门表 C, 部门表 D,医疗付款方式 E " & vbNewLine & _
                "Where A.病人id = B.病人id  And B.入院科室ID =C.ID And B.入院病区ID =D.ID(+) And B.医疗付款方式 = E.名称(+) And A.病人id = [1] And B.主页id = [3]"
        End If
    ElseIf gbytPass = ZL Then
        If str挂号单 <> "" Then
            strSQL = "Select b.ID as 就诊ID,A.门诊号,A.入院时间,A.就诊时间,A.婚姻状况,B.姓名,B.性别,A.出生日期,A.年龄,A.职业,Decode(A.出生日期, Null, '', Round(Sysdate - A.出生日期, 2))as 年龄数字 " & _
            "   ,A.当前床号,B.执行部门ID AS 当前科室ID,B.执行人,C.名称 AS 当前科室" & vbNewLine & _
            " From 病人信息 A,病人挂号记录 B, 部门表 C " & _
            " Where A.病人ID=B.病人ID And B.执行部门ID =C.ID And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
        Else
            strSQL = "Select A.住院号,A.出生日期,A.入院时间,A.婚姻状况, Nvl(B.姓名, A.姓名) 姓名, Nvl(B.性别, A.性别) 性别,Nvl(B.年龄, A.年龄) As 年龄,A.职业," & _
                "Decode(A.出生日期, Null, '', Round(Sysdate - A.出生日期, 2))as 年龄数字,B.身高,B.体重 " & vbNewLine & _
                "   ,A.当前床号,A.当前科室ID,B.当前病区ID,C.名称 AS 当前科室,D.名称 AS 当前病区 " & vbNewLine & _
                "From 病人信息 A, 病案主页 B,部门表 C,部门表 D" & vbNewLine & _
                "Where B.病人id = A.病人id And A.当前科室ID =C.ID(+) And B.当前病区ID =D.ID(+) And B.病人id = [1] And B.主页id = [3]"
        End If
    End If
    Set GetPatiInfo_YF = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, str挂号单, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAdviceInfo_YF(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, _
    Optional ByVal str医嘱IDs As String, Optional ByVal bytFunc As Byte = 0) As ADODB.Recordset
'功能:获取药嘱信息
    Dim strSQL As String
    
    If str挂号单 <> "" Then
        strSQL = " And a.挂号单 = [3]"
    Else
        strSQL = " And  a.病人ID =[1] And a.主页ID = [2]"
    End If
    If bytFunc = 0 Then
        If str医嘱IDs <> "" Then
            str医嘱IDs = "," & str医嘱IDs & ","
            strSQL = "Select a.Id As 医嘱id, a.相关id,a.医嘱期效, a.序号 As 医嘱序号,a.紧急标志 as 标志,a.开嘱医生, a.医嘱状态, a.开嘱时间, a.开始执行时间 As 开始时间, a.执行终止时间 As 结束时间, a.诊疗类别," & vbNewLine & _
            "       a.诊疗项目id, a.收费细目id, a.执行频次 As 频率, a.间隔单位, a.频率次数, a.频率间隔, a.单次用量, a.总给予量 As 总量, e.门诊包装, e.门诊单位, e.住院单位, a.天数," & vbNewLine & _
            "       b.名称 As 药品名称, b.计算单位 As 单量单位, b.执行频率, g.名称 As 用法, c.诊疗项目id As 用法id, a.开嘱科室id, a.执行科室id, d.名称 As 开嘱科室, a.用药目的,a.用药理由,a.医生嘱托,F.规格, " & vbNewLine & _
            "  Decode(G.类别||'_'||G.操作类型||'_'||G.执行分类,'E_2_1',C.医生嘱托,'') As 滴速,NULL As 处方号" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱记录 C, 部门表 D, 药品规格 E,收费项目目录 F, 诊疗项目目录 G " & vbNewLine & _
            "Where a.诊疗项目id = b.Id And a.相关id = c.Id(+) And a.开嘱科室id = d.Id(+) And a.收费细目id = e.药品id(+) And a.收费细目id=F.ID(+) And " & vbNewLine & _
            "      c.诊疗项目id = g.Id(+) And a.诊疗类别 In ('5', '6', '7') " & strSQL & " And inStr([4],','|| a.ID ||',')>0 " & vbNewLine & _
            "Order By a.序号"
        Else
            strSQL = "Select a.Id As 医嘱id, a.相关id, a.医嘱期效, a.序号 As 医嘱序号,a.紧急标志 as 标志,a.开嘱医生, a.医嘱状态, a.开嘱时间, a.开始执行时间 As 开始时间, a.执行终止时间 As 结束时间, a.诊疗类别," & vbNewLine & _
                     "       a.诊疗项目id, a.收费细目id, a.执行频次 As 频率, a.间隔单位, a.频率次数, a.频率间隔, a.单次用量, a.总给予量 As 总量, e.门诊包装, e.门诊单位, e.住院单位, a.天数," & vbNewLine & _
                     "       b.名称 As 药品名称, b.计算单位 As 单量单位, b.执行频率, g.名称 As 用法, c.诊疗项目id As 用法id, a.开嘱科室id, a.执行科室id, d.名称 As 开嘱科室, a.用药目的,a.用药理由,a.医生嘱托,F.规格,Decode(G.类别||'_'||G.操作类型||'_'||G.执行分类,'E_2_1',C.医生嘱托,'') As 滴速, NULL As 处方号" & vbNewLine & _
                     "From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱记录 C, 部门表 D, 药品规格 E,收费项目目录 F, 诊疗项目目录 G " & vbNewLine & _
                     "Where a.诊疗项目id = b.Id And a.相关id = c.Id(+) And a.开嘱科室id = d.Id(+) And a.收费细目id = e.药品id(+) And a.收费细目id=F.ID(+) And " & vbNewLine & _
                     "      c.诊疗项目id = g.Id(+) And a.诊疗类别 In ('5', '6', '7') " & strSQL & vbNewLine & _
                     "      And ((a.医嘱期效 = 1 And Trunc(a.开嘱时间) = Trunc(Sysdate) And a.医嘱状态 = 8) Or" & vbNewLine & _
                     "      (a.医嘱期效 = 0 And (a.医嘱状态 In (8, 9) And a.执行终止时间 >= Sysdate Or a.医嘱状态 In (3, 5, 7))))" & vbNewLine & _
                     "Order By a.序号"
        End If
    ElseIf bytFunc = 1 Then
        strSQL = "Select a.Id As 医嘱id, a.相关id, a.医嘱期效, a.序号 As 医嘱序号,a.紧急标志 as 标志,a.开嘱医生, a.医嘱状态, a.开嘱时间, a.开始执行时间 As 开始时间, a.执行终止时间 As 结束时间, a.诊疗类别," & vbNewLine & _
        "       a.诊疗项目id, a.收费细目id, a.执行频次 As 频率, a.间隔单位, a.频率次数, a.频率间隔, a.单次用量, a.总给予量 As 总量, e.门诊包装, e.门诊单位, e.住院单位, a.天数," & vbNewLine & _
        "       b.名称 As 药品名称, b.计算单位 As 单量单位, b.执行频率, g.名称 As 用法, c.诊疗项目id As 用法id, a.开嘱科室id, a.执行科室id, d.名称 As 开嘱科室, a.用药目的, a.用药理由, a.医生嘱托,F.规格," & vbNewLine & _
        "       Decode(G.类别||'_'||G.操作类型||'_'||G.执行分类,'E_2_1',C.医生嘱托,'') As 滴速,NULL As 处方号,a.处方序号 AS 处方ID,A.执行性质 AS A执行性质,C.执行性质 AS B执行性质,G.类别,G.操作类型,G.执行分类  " & vbNewLine & _
        "From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱记录 C, 部门表 D, 药品规格 E,收费项目目录 F, 诊疗项目目录 G " & vbNewLine & _
        "Where a.诊疗项目id = b.Id And a.相关id = c.Id(+) And a.开嘱科室id = d.Id(+) And a.收费细目id = e.药品id(+) And a.收费细目id=F.ID(+) And " & vbNewLine & _
        "      c.诊疗项目id = g.Id(+) And a.诊疗类别 In ('5', '6', '7') And Nvl(A.执行标记,0)<>-1 " & strSQL & vbNewLine & _
        "Order By a.序号"
    ElseIf bytFunc = 2 Then
        strSQL = "Select a.病人id, a.挂号单, a.Id As 医嘱id, a.相关id, a.医嘱期效, a.序号 As 医嘱序号, a.紧急标志 As 标志, a.开嘱医生, a.医嘱状态, a.开嘱时间, a.开始执行时间 As 开始时间," & vbNewLine & _
            "       a.执行终止时间 As 结束时间, a.诊疗类别, a.诊疗项目id, a.收费细目id, a.执行频次 As 频率, a.间隔单位, a.频率次数, a.频率间隔, a.单次用量, a.总给予量 As 总量, e.门诊包装," & vbNewLine & _
            "       e.门诊单位, e.住院单位, a.天数, b.名称 As 药品名称, b.计算单位 As 单量单位, b.执行频率, g.名称 As 用法, c.诊疗项目id As 用法id, a.开嘱科室id, a.执行科室id," & vbNewLine & _
            "       d.名称 As 开嘱科室, a.用药目的, a.用药理由, a.医生嘱托, f.规格," & vbNewLine & _
            "       Decode(g.类别 || '_' || g.操作类型 || '_' || g.执行分类, 'E_2_1', c.医生嘱托, '') As 滴速, Nvl(a.处方序号, 0) As 处方号" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱记录 C, 部门表 D, 药品规格 E, 收费项目目录 F, 诊疗项目目录 G" & vbNewLine & _
            "Where a.诊疗项目id = b.Id And a.相关id = c.Id(+) And a.开嘱科室id = d.Id(+) And a.收费细目id = e.药品id(+) And a.收费细目id = f.Id(+) And" & vbNewLine & _
            "      c.诊疗项目id = g.Id(+) And a.诊疗类别 In ('5', '6', '7') And" & vbNewLine & _
            "      (a.医嘱状态 = 1 And Instr([4], ',' || a.相关id || ',') > 0 Or a.医嘱状态 = 8) " & strSQL & vbNewLine & _
            "Order By a.序号"
    End If
    On Error GoTo errH
    Set GetAdviceInfo_YF = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng病人ID, lng主页ID, str挂号单, "," & str医嘱IDs & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function AdviceCheckWarn_MK4(ByVal lngPatiID As Long, ByVal str挂号单 As String, ByVal lng主页ID As Long, ByVal bytShow As Byte, _
    ByVal bytSubmit As Byte, ByRef rsAdvice As ADODB.Recordset, Optional str警示 As String = "-1", Optional lngResult As Long = 1) As String
'功能：美康4.0自定义接口
'参数:
'   lngCmd=0 MK4_检测PASS菜单状态
'       1-手动审查
'       2-保存审查
'   bytShow-0-不显示界面,1-显示界面
'   bytSubmit-0-不采集数据,1-采集数据
'   rsAdvice-传人医嘱记录(引用类型,便于返回警示信息)
'   lngResult 表示药师干预结果：1-通过，0-不能通过
'返回

'str警示-返回警示串：格式：医嘱ID1:警示值1,医嘱ID2:警示值2    （缺省不返回警示值）
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngCount As Long, strInHospNo As String, strVisitCode As String
    Dim i As Long, k As Long
    Dim str身高 As String, str体重 As String, str急诊标识 As String
    Dim lng妊娠 As Long, lng哺乳 As Long, lng肝功 As Long, lng肾功 As Long
    Dim str妊娠日期 As String, strTmp As String, strJson As String
    Dim lng挂号ID As Long
    Dim str药品ID As String
    Dim rs规格 As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset
    Dim str药房IDs As String
    Dim strPharmacyName As String
    
    Dim colTemp As New Collection
    
    On Error GoTo errH
    Screen.MousePointer = 11
    '美康4.0

     '传入病人就诊信息(PASS需要的基本内容,同一病人可不重复传入)
     '-------------------------------------------------------------
     Set rsTmp = GetPatiInfo_YF(lngPatiID, str挂号单, lng主页ID)
     If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
    '病人生理情况
    strTmp = Get病人病生理情况(lngPatiID, IIf(str挂号单 <> "", 0, lng主页ID))
    Call PASS4病生理情况(strTmp, lng哺乳, lng妊娠, lng肝功, lng肾功, str妊娠日期)
    '直接传lng哺乳,lng妊娠传入美康接口，美康内部接收时值会被转换成极大值
    colTemp.Add -1, "K" & "-1"
    colTemp.Add 0, "K" & "0"
    colTemp.Add 1, "K" & "1"
    colTemp.Add 2, "K" & "2"
    colTemp.Add 3, "K" & "3"
    colTemp.Add 4, "K" & "4"
     
    'PASS增加一个病人的基本信息MDC_SetPatient
    If str挂号单 <> "" Then
        lng挂号ID = rsTmp!就诊Id
        '附加信息
        strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                        "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                        "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngPatiID, lng挂号ID)
        rsPatiInfo.Filter = "项目名称='身高'"
        If rsPatiInfo.RecordCount > 0 Then str身高 = NVL(rsPatiInfo!记录内容)
        rsPatiInfo.Filter = "项目名称='体重'"
        If rsPatiInfo.RecordCount > 0 Then str体重 = NVL(rsPatiInfo!记录内容)
        
        str急诊标识 = IIf(NVL(rsTmp!急诊, 0) = 0, 2, 3)
        '药师干预系统
        strInHospNo = rsTmp!门诊号 & "/" & str挂号单
        strVisitCode = rsTmp!门诊号 & ""
         'A.PASS增加一个病人的基本信息MDC_SetPatient
        Call MDC_SetPatient(lngPatiID, strInHospNo, strVisitCode, rsTmp!姓名 & "", NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), _
                     str身高, str体重, rsTmp!科室ID & "", rsTmp!科室名 & "", rsTmp!医生码 & "", rsTmp!医生名 & "", str急诊标识, CLng(colTemp("K" & lng哺乳)), CLng(colTemp("K" & lng妊娠)), str妊娠日期, CLng(colTemp("K" & lng肝功)), CLng(colTemp("K" & lng肾功)))
     
     Else
        strInHospNo = rsTmp!住院号 & ""
        strVisitCode = lng主页ID & ""
        Call MDC_SetPatient(lngPatiID & "", rsTmp!住院号 & "", lng主页ID & "", rsTmp!姓名 & "", rsTmp!性别 & "", _
             Format(rsTmp!出生日期, "yyyy-MM-dd"), rsTmp!身高 & "", rsTmp!体重 & "", _
             rsTmp!科室ID & "", rsTmp!科室名 & "", rsTmp!医生码 & "", rsTmp!医生名 & "", 1, CLng(colTemp("K" & lng哺乳)), CLng(colTemp("K" & lng妊娠)), str妊娠日期, CLng(colTemp("K" & lng肝功)), CLng(colTemp("K" & lng肾功)))
     End If
     
     '传人病人过敏史
     '-------------------------------------------------------
     Set rsTmp = Get病人过敏记录(lngPatiID, IIf(str挂号单 <> "", 0, lng主页ID))
     
     'PASS增加一条标准化的过敏记录（多条重复调用）MDC_AddAller
     '任意取一个药品ID传人
     For i = 1 To rsTmp.RecordCount
         str药品ID = ""
         If rsTmp!药物ID & "" <> "" Then
             strSQL = "select 药品ID from 药品规格 where 药名id=[1] and rownum <2"
             Set rs规格 = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, rsTmp!药物ID)
             If Not rs规格.EOF Then str药品ID = rs规格!药品ID & ""
         End If
         Call MDC_AddAller(i, str药品ID, rsTmp!药物名 & "", rsTmp!过敏反应 & "")
         rsTmp.MoveNext
     Next

     '传人病人诊断
     '------------------------------------------------------------------
     'PASS增加一条标准化的诊断记录（多条重复调用）MDC_AddMedCond
    If glngModel = PM_门诊编辑 Then
        If Not gobjDiags Is Nothing Then
            With gobjDiags
                For i = 1 To .Count
                    If .Item(i).str诊断描述 <> "" Then
                        Call MDC_AddMedCond(i & "", IIf(.Item(i).str疾病编码 <> "", .Item(i).str疾病编码, .Item(i).str诊断编码), .Item(i).str诊断描述, "")
                    End If
                Next
            End With
        End If
    Else
        Set rsTmp = Get病人诊断记录(lngPatiID, IIf(str挂号单 <> "", lng挂号ID, lng主页ID), IIf(str挂号单 <> "", "1,11", "2,12"))
        If lng主页ID <> 0 Then
           Set rsTmp = zlDatabase.CopyNewRec(rsTmp, , "编码,名称") '复制为可编辑的记录集117045
           If CreatePlugInOK Then
               On Error Resume Next
               Call gobjPlugIn.SetPassDiag(lngPatiID, lng主页ID, rsTmp)
               Call zlPlugInErrH(Err, "SetPassDiag")
               If Not rsTmp Is Nothing Then rsTmp.Filter = ""
               Err.Clear: On Error GoTo 0
           End If
        End If
        For i = 1 To rsTmp.RecordCount
            Call MDC_AddMedCond(i & "", rsTmp!编码 & "", rsTmp!名称 & "", "")
            rsTmp.MoveNext
        Next
    End If
    '传入病人手术记录MDC_AddOperation
    Set rsTmp = GetPatiOperation(lngPatiID, lng主页ID, str挂号单)
    
    On Error Resume Next
    For i = 1 To rsTmp.RecordCount
        Call MDC_AddOperation(rsTmp!id & "", rsTmp!编码 & "", rsTmp!名称 & "", "", Format(rsTmp!手术时间 & "", "YYYY-MM-DD HH:MM:SS"), "")
        rsTmp.MoveNext
    Next
    Err.Clear: On Error GoTo 0
    
    On Error GoTo errH
     'PASS增加一条用药清单记录（多条重复调用）MDC_AddScreenDrug
     lngCount = 0
     
     With rsAdvice
         str药房IDs = ""
         For i = 1 To .RecordCount
            If InStr("," & str药房IDs & ",", "," & !执行科室ID & ",") = 0 Then
                str药房IDs = str药房IDs & "," & !执行科室ID
            End If
            .MoveNext
         Next
         If str药房IDs <> "" Then Set rsTmp = GetRS("部门表", "ID,名称", str药房IDs)
         If .RecordCount > 0 Then .MoveFirst
         For i = 1 To .RecordCount
             '传入医嘱信息
             '医嘱ID,相关ID,医嘱期效,医嘱序号,医嘱状态,开嘱科室,开嘱科室ID,开嘱医生编码,开嘱医生,药品ID,药品名称,单次用量,单量单位,频率,用法,用法ID,开嘱时间,开始时间,结束时间,总量,总量单位,用药目的,医生嘱托
             If glngModel = PM_门诊发送 Then
                strTmp = !处方号
             Else
                strTmp = ""
             End If
            If Val(!医嘱状态 & "") = -1 Then
                '门诊发送传入历史医嘱
                strJson = FuncGetOtherRecipInfo(!医嘱ID, strTmp, !药品ID, !药品名称, !用法, !频率, !单量单位, !单次用量, !总量, !总量单位, !天数)
                If strJson <> "" Then Call MDC_AddJsonInfo(strJson)
            Else
                Call MDC_AddScreenDrug(!医嘱ID, !医嘱序号, !药品ID, !药品名称 & "", !单次用量 & "", !单量单位 & "", !频率 & "", !用法 & "", !用法 & "", !开嘱时间 & "", _
                        !结束时间 & "", !开始时间 & "", !相关ID & "", !医嘱期效 & "", !医嘱状态 & "", !开嘱科室id & "", !开嘱科室 & "", !开嘱医生编码 & "", _
                        !开嘱医生 & "", strTmp, !总量 & "", !总量单位 & "", !用药目的 & "", "", "", !医生嘱托 & "")
            End If
            '滴速\执行科室
            rsTmp.Filter = "ID=" & Val(!执行科室ID): strPharmacyName = ""
            If Not rsTmp.EOF Then strPharmacyName = rsTmp!名称 & ""
            strJson = FuncGetDripInfo(!医嘱ID & "", !滴速 & "", Val(!执行科室ID), strPharmacyName, !天数)
            If strJson <> "" Then Call MDC_AddJsonInfo(strJson)
            lngCount = lngCount + 1
            
             .MoveNext
         Next
     End With
     
     
     '无可审查的药品l
     If lngCount = 0 Then
         Screen.MousePointer = 0: Exit Function
     End If
     
    If gblnTEST Then bytShow = 0
     'PASS审查函数MDC_DoCheck
    If bytShow = 0 And bytSubmit = 0 Then
        Call MDC_DoCheck(G_INT_MODEL_0, G_INT_MODEL_0)  '不显示界面,不采集
    ElseIf bytShow = 0 And bytSubmit = 1 Then
        Call MDC_DoCheck(G_INT_MODEL_0, G_INT_MODEL_1) '不显示界面,要采集
    ElseIf bytShow = 1 And bytSubmit = 0 Then
        Call MDC_DoCheck(G_INT_MODEL_1, G_INT_MODEL_0) '显示界面,不要采集
    ElseIf bytShow = 1 And bytSubmit = 1 Then
        Call MDC_DoCheck(G_INT_MODEL_1, G_INT_MODEL_1) '显示界面,要采集
    End If
    
    If gblnPharmReview And glngModel = PM_住院编辑 Then
       On Error Resume Next
       lngResult = MDC_GetTaskStatus(lngPatiID, strInHospNo, strVisitCode, "", 1)  '返回值:1-通过
       WriteLog "" & glngModel, "AdviceCheckWarn_MK4", "MDC_GetTaskStatus 返回值:" & lngResult
       Err.Clear: On Error GoTo 0
    ElseIf gblnPharmReview And glngModel = PM_门诊发送 Then
        With rsAdvice
            .Filter = "医嘱状态='0'": strTmp = "": lngResult = 0
            For i = 1 To .RecordCount
                If strTmp <> !处方号 & "" Then
                    strTmp = !处方号 & ""
                    On Error Resume Next
                    lngResult = MDC_GetTaskStatus(lngPatiID, strInHospNo, strVisitCode, strTmp, 2)
                    WriteLog "" & glngModel, "AdviceCheckWarn_MK4", "MDC_GetTaskStatus 返回值:" & lngResult
                    Err.Clear: On Error GoTo 0
                    !审核状态 = lngResult
                Else
                    !审核状态 = lngResult
                End If
                .MoveNext
            Next
            .Filter = ""
        End With
    Else
       lngResult = 1
    End If

     'PASS审查警示值更新
    If str警示 <> "-1" And lngCount > 0 Then
        str警示 = ""
        rsAdvice.MoveFirst
        For i = 1 To rsAdvice.RecordCount
            k = MDC_GetWarningCode(rsAdvice!医嘱ID & "")
            str警示 = str警示 & "," & rsAdvice!医嘱ID & ":" & k
            rsAdvice!警示 = k
            rsAdvice.MoveNext
        Next
        If str警示 <> "" Then str警示 = Mid(str警示, 2)
        
    Else
        str警示 = ""
    End If
            
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function OutAdviceCheckWarn_MK4(Optional ByVal bytShow As Byte = 0, Optional ByVal bytSubmit As Byte = 0, _
        Optional blnIsHaveOut As Boolean, Optional ByRef blnNoSave As Boolean, Optional ByRef rsOut As ADODB.Recordset, _
        Optional ByRef lngResult As Long = 1) As Long
'功能：调用Pass系统中对医嘱进行合理用药审查等相关功能
'参数：bytShow=0-不显示审查结果界面,1-显示审查结果界面
'       0-检测菜单可用性，1-审查接口
'       0-检测设置PASS菜单状态,1-审查接口
'       bytSubmit=0-无需上传数据,1-上传数据
'出参：
'       rsOut-返回禁忌药品说明
'       lngResult-药师干预系统 0-不通过；1-通过
'返回：本次审核返回的最高级别警示值,为-1,-2,-3表示没有进行审查
'      检测PASS菜单时，返回>=0表示可以弹出菜单
'说明：用药审查：涉及当天下的临嘱(包括已执行)，和未停止的长嘱
'      用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset, rsPatiInfo As New ADODB.Recordset
    Dim rs诊断 As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim rs中药 As ADODB.Recordset
    
    Dim str药品名称 As String, str用法 As String, str频率 As String, str用法ID As String, str间隔单位 As String
    Dim str诊断编码 As String, str诊断描述 As String, strTmp As String, strPre用法 As String, str诊疗类别 As String
    Dim str医嘱ID As String, str相关ID As String, str医嘱序号 As String, str单次用量 As String, str单量单位 As String
    Dim str急诊标识 As String, str总量 As String, str总量单位 As String, str用药目的 As String, str医生嘱托 As String
    Dim str医嘱期效 As String, str医嘱状态 As String, str医生编码 As String
    Dim str药品ID As String, str开嘱科室ID As String, str开嘱科室Tag As String
    Dim str开嘱科室 As String, str开嘱医生 As String, str开嘱医生Tag As String
    Dim str开嘱时间 As String, str开始时间 As String, str结束时间 As String
    Dim str警示 As String, str警示值 As String, str滴速 As String
    Dim str中药组IDs As String, strGroupIDs As String
    Dim str医嘱IDs As String
    Dim str执行科室ID As String
    
    Dim lngMaxWarn As Long, strOld As String
    Dim strSQL As String, blnDo As Boolean
    Dim lngCount As Long, curDate As Date
    Dim lngTmp As Long, lng中药组ID As Long, lngLight As Long
    Dim arrLevel(0 To 4) As Long
    Dim i As Long, k As Long, j As Long
    Dim arrTmp As Variant
    
    Dim strType As String
    Dim str身高 As String, str体重 As String
    Dim arrLight(0 To 4) As String
    
    Dim int频率次数 As Integer, int频率间隔 As Integer
    Dim objDiag As clsDiagItem
    Dim rs规格 As ADODB.Recordset
    
    Dim arrSQL As Variant
    
    lngMaxWarn = -1
    OutAdviceCheckWarn_MK4 = lngMaxWarn

    On Error GoTo errH
    Screen.MousePointer = 11
    
     '传入病人医嘱信息
    '-------------------------------------------------------------
    '启用了禁忌药品说明参数;场合为门诊编辑;审查功能
    If glngModel = PM_门诊编辑 And gbytReason = 1 Then
        Set rsOut = InitAdviceRS(FUN_输出内容)
    End If
    
    With gobjAdvice
        lngCount = 0
        curDate = zlDatabase.Currentdate
        '初始化药嘱信息
        Set rsAdvice = InitAdviceRS(FUN_医嘱信息)
        
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_门诊编辑 Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            Else
                blnDo = ((InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0) _
                Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4"))
                blnDo = blnDo And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            End If
            
            If blnDo Then
                If glngModel = PM_门诊医嘱清单 And (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                    '获取中药医嘱组ID
                    str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    '取药品名称
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                        str药品名称 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                    Else
                        str药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                    End If
                    str药品ID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                    '取药品给药途径
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str用法 = ""    '一并给药不重复取
                    
                    If str用法 = "" Then
                        str滴速 = ""
                        If glngModel = PM_门诊编辑 Then
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                            If k <> -1 Then
                                If .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                                    str用法 = .TextMatrix(k, gobjCOL.intCOL用法)
                                Else
                                    str用法 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                    If InStr(.TextMatrix(k, gobjCOL.intcol医嘱嘱托), "滴/分钟") > 0 Or InStr(.TextMatrix(k, gobjCOL.intcol医嘱嘱托), "毫升/小时") > 0 Then
                                        str滴速 = .TextMatrix(k, gobjCOL.intcol医嘱嘱托)
                                    End If
                                End If
                            End If
                        Else
                            str用法 = Sys.RowValue("病人医嘱记录", Val(.TextMatrix(i, gobjCOL.intCOL相关ID)), "医嘱内容")
                        End If
                    End If
    
                    '取用药频率(次/天),都为整数四舍五入
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str频率 = ""    '一并给药不重复取
                    If str频率 = "" Then
                        str频率 = .TextMatrix(i, gobjCOL.intCOL频率)
                    End If
    
                    '开嘱科室名称
                    str开嘱科室ID = .TextMatrix(i, gobjCOL.intCOL开嘱科室ID)
                    If str开嘱科室ID <> str开嘱科室Tag And str开嘱科室ID <> "" Then
                        str开嘱科室 = Sys.RowValue("部门表", Val(str开嘱科室ID), "名称")
                        str开嘱科室Tag = str开嘱科室ID
                    End If
                    
                    '开嘱医生
                    str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                    
                    If str开嘱医生Tag <> str开嘱医生 And str开嘱医生 <> "" Then
                        str医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
                        str开嘱医生Tag = str开嘱医生
                    End If
    
                    strType = .TextMatrix(i, gobjCOL.intCOL状态)
                    '"0"-在用（默认）；"1"-已作废；"2"-已停嘱；"3"-离院带药（根据系统设置参与审查）
                    If strType = "4" Then '4-作废
                        str医嘱状态 = "1"
                    Else
                        str医嘱状态 = "0"
                    End If
                    
                    'PASS增加一条用药清单记录（多条重复调用）MDC_AddScreenDrug
                    If glngModel = PM_门诊编辑 Then
                        str医嘱ID = .RowData(i)
                        str医嘱序号 = .TextMatrix(i, gobjCOL.intCOL序号)
                        str单次用量 = .TextMatrix(i, gobjCOL.intCOL单量)
                        str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                        str总量 = .TextMatrix(i, gobjCOL.intCOL总量)
                        str总量单位 = .TextMatrix(i, gobjCOL.intcol总量单位)
                        
                        str开嘱时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:MM:SS")
                        str开始时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:MM:SS")
                        str结束时间 = "" '门诊医生工作站，门诊药房 传空值，就可以审查出重复用药
                        str执行科室ID = .TextMatrix(i, gobjCOL.intCol执行科室ID)
                        
                        If Not rsOut Is Nothing Then
                            If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                            '西药,中成药
                                rsOut.AddNew
                                rsOut!医嘱ID = CLng(.RowData(i) & "")
                                rsOut!禁忌药品说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                                rsOut!状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                                rsOut!药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                                rsOut.Update
                            ElseIf Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                            '中药配方  禁忌说明保存在用药服法上
                                k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                                If k <> -1 Then
                                    rsOut.AddNew
                                    rsOut!医嘱ID = CLng(.RowData(k) & "")
                                    rsOut!禁忌药品说明 = .TextMatrix(k, gobjCOL.intCol禁忌药品说明)
                                    rsOut!状态 = .TextMatrix(k, gobjCOL.intCOL状态)
                                    rsOut!药品名称 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                    rsOut.Update
                                End If
                            End If
                        End If
                            
                    Else
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                        str医嘱IDs = str医嘱IDs & "," & str医嘱ID
                        str医嘱序号 = "-1"  '系统自动编号
                        str单次用量 = Val(.TextMatrix(i, gobjCOL.intCOL单量))
                        str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量)
                        str单次用量 = FormatEx(str单次用量, 5)
                        
                        str单量单位 = Replace(str单量单位, str单次用量, "")
                        
                        str总量 = Val(.TextMatrix(i, gobjCOL.intCOL总量))
                        str总量单位 = .TextMatrix(i, gobjCOL.intCOL总量)
                        str总量 = FormatEx(str总量, 5)
                        str总量单位 = Replace(str总量单位, str总量, "")
                        
                        str开嘱时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:MM:SS")
                        str开始时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:MM:SS")
                        str结束时间 = ""  '门诊医生工作站，门诊药房 传空值，就可以审查出重复用药
                        str执行科室ID = ""
                        If InStr(strGroupIDs & ",", "," & .TextMatrix(i, gobjCOL.intCOL相关ID) & ",") = 0 Then
                            strGroupIDs = strGroupIDs & "," & .TextMatrix(i, gobjCOL.intCOL相关ID)
                        End If
                    End If
                    str相关ID = .TextMatrix(i, gobjCOL.intCOL相关ID)
                    str医生嘱托 = .TextMatrix(i, gobjCOL.intcol医嘱嘱托)
                    str用药目的 = .TextMatrix(i, gobjCOL.intcol用药目的)
                    str诊疗类别 = .TextMatrix(i, gobjCOL.intCOL诊疗类别)
                    If str用药目的 = "1" Then
                        str用药目的 = "3"
                    ElseIf str用药目的 = "2" Then
                        str用药目的 = "4"
                    Else
                        str用药目的 = "0"
                    End If
                    
                    '----------------------------------------------------------
                    rsAdvice.AddNew
                    rsAdvice!医嘱ID = str医嘱ID
                    rsAdvice!相关ID = str相关ID
                    rsAdvice!医嘱期效 = "1" '门诊
                    rsAdvice!医嘱序号 = str医嘱序号
                    rsAdvice!医嘱状态 = str医嘱状态
                    rsAdvice!开嘱科室 = str开嘱科室
                    rsAdvice!开嘱科室id = str开嘱科室ID
                    rsAdvice!开嘱医生编码 = str医生编码
                    rsAdvice!开嘱医生 = str开嘱医生
                    rsAdvice!药品ID = str药品ID
                    rsAdvice!药品名称 = str药品名称
                    rsAdvice!单次用量 = str单次用量
                    
                    rsAdvice!单量单位 = str单量单位
                    rsAdvice!频率 = str频率
                    rsAdvice!用法 = str用法
                    rsAdvice!用法ID = ""
                    rsAdvice!开嘱时间 = str开嘱时间
                    rsAdvice!开始时间 = str开始时间
                    rsAdvice!结束时间 = str结束时间
                    
                    rsAdvice!总量 = str总量
                    rsAdvice!总量单位 = str总量单位
                    rsAdvice!用药目的 = str用药目的
                    rsAdvice!医生嘱托 = str医生嘱托
                    rsAdvice!诊疗类别 = str诊疗类别
                    rsAdvice!滴速 = str滴速
                    rsAdvice!执行科室ID = str执行科室ID
                    rsAdvice.Update
                    '---------------------------------------------------------------------------
                    lngCount = lngCount + 1
                End If
            End If
        Next
        '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
        If glngModel = PM_门诊医嘱清单 Then
            If str中药组IDs <> "" Then
                Set rs中药 = Get中药配方(str中药组IDs)
                With rs中药
                    For i = 1 To .RecordCount
                        If !相关ID & "" <> str相关ID Then
                            str开嘱医生 = !开嘱医生 & ""
                            If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                            str开嘱医生 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名") & "/" & str开嘱医生
                            str开嘱科室 = Sys.RowValue("部门表", Val(!开嘱科室id & ""), "名称")

                            str开嘱时间 = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")
                            str结束时间 = ""
                            str开始时间 = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")
                            If !医嘱期效 & "" = "1" Then
                                str结束时间 = str开嘱时间
                            End If
                            
                            If !用药目的 & "" = "1" Then
                                str用药目的 = "3"
                            ElseIf !用药目的 & "" = "2" Then
                                str用药目的 = "4"
                            Else
                                str用药目的 = "0"
                            End If
                            
                            If !医嘱状态 & "" = "4" Then '作废
                                str医嘱状态 = "1"
                            Else
                                str医嘱状态 = "0"
                            End If
                            str相关ID = !相关ID & ""
                        End If
                        '----------------------------------------------------------
                        rsAdvice.AddNew
                        rsAdvice!医嘱ID = !id
                        str医嘱IDs = str医嘱IDs & "," & !id
                        rsAdvice!相关ID = !相关ID & ""
                        rsAdvice!医嘱期效 = "1"
                        rsAdvice!医嘱序号 = lngCount + 1
                        rsAdvice!医嘱状态 = str医嘱状态
                        rsAdvice!开嘱科室 = str开嘱科室
                        rsAdvice!开嘱科室id = !开嘱科室id & ""
                        rsAdvice!开嘱医生编码 = str医生编码
                        rsAdvice!开嘱医生 = str开嘱医生
                        rsAdvice!药品ID = !药品ID & ""
                        rsAdvice!药品名称 = !药品名称 & ""
                        rsAdvice!单次用量 = !单次用量 & ""
                        
                        rsAdvice!单量单位 = !单量单位 & ""
                        rsAdvice!频率 = !频率 & ""
                        rsAdvice!用法 = !用法 & ""
                        rsAdvice!用法ID = ""
                        rsAdvice!开嘱时间 = str开嘱时间
                        rsAdvice!开始时间 = str开始时间
                        rsAdvice!结束时间 = str结束时间
                        
                        rsAdvice!总量 = !总给予量 & ""
                        rsAdvice!总量单位 = !门诊单位 & ""
                        rsAdvice!用药目的 = str用药目的
                        rsAdvice!医生嘱托 = !医生嘱托 & ""
                        rsAdvice!诊疗类别 = !诊疗类别 & ""
                        rsAdvice.Update
                        '----------------------------------------------------------------------------
                        lngCount = lngCount + 1
                        .MoveNext
                    Next
                End With
            End If
            If str医嘱IDs <> "" Then
                Set rsTmp = GetDrugInfo_MK4(gobjPati.str挂号单, str医嘱IDs)
                rsAdvice.Filter = ""
                For i = 1 To rsAdvice.RecordCount
                    rsTmp.Filter = "ID=" & rsAdvice!医嘱ID
                    If Not rsTmp.EOF Then
                        rsAdvice!处方序号 = rsTmp!处方序号 & ""
                        rsAdvice!用法 = rsTmp!用法 & ""
                        rsAdvice!执行科室ID = rsTmp!执行科室ID & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
            'Drip 滴速
            If Mid(strGroupIDs, 2) <> "" Then
                Set rsTmp = Get滴速(strGroupIDs)
                For i = 1 To rsTmp.RecordCount
                    rsAdvice.Filter = "相关ID =" & rsTmp!id
                    Do While Not rsAdvice.EOF
                        rsAdvice!滴速 = rsTmp!医生嘱托 & ""
                        rsAdvice.MoveNext
                    Loop
                    rsTmp.MoveNext
                Next
                rsAdvice.Filter = ""
            End If
        End If
        '无可审查的药品
        If lngCount = 0 Then
            Screen.MousePointer = 0: Exit Function
        End If
        
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        
        Call AdviceCheckWarn_MK4(gobjPati.lng病人ID, gobjPati.str挂号单, 0, bytShow, bytSubmit, rsAdvice, str警示, lngResult)
        
        arrSQL = Array()
        '获取警示级别
        '返回值顺：0-蓝灯,1-黑灯,2-红灯,3-橙灯,4-黄灯
        '警示级顺：0-蓝灯,4-黄灯,3-橙灯,2-红灯,1-黑灯(因为PASS升级的原因)
        arrLevel(0) = 0: arrLevel(1) = 4: arrLevel(2) = 3: arrLevel(3) = 2: arrLevel(4) = 1
        arrLight(0) = "蓝_4": arrLight(1) = "黑_4": arrLight(2) = "红_4": arrLight(3) = "橙_4": arrLight(4) = "黄_4"
        For i = .FixedRows To .Rows - 1
            If glngModel = PM_门诊编辑 Then
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                        And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            Else
                blnDo = ((InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0) _
                Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4"))
                blnDo = blnDo And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
            End If
                
            If blnDo Then
                If glngModel = PM_门诊医嘱清单 Then
                    str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                Else
                    str医嘱ID = .RowData(i)
                End If
                rsAdvice.Filter = "医嘱ID = '" & str医嘱ID & "'"
                If rsAdvice.RecordCount > 0 Then
                     k = CLng(rsAdvice!警示 & "")
                Else
                     k = -1 '医嘱清单中药配方
                End If
               
                If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                    If k >= 0 And k <= 4 Then
                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = k
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(k)).Picture
                    Else
                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                    End If

                    If PM_门诊编辑 = glngModel Then
                        If strOld <> CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) Then
                            .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                            blnNoSave = True    '标记为未保存
                        End If
                        '记录下禁忌药品 K=1 代表黑灯 且 只针对未校对医嘱进行禁忌药品说明原因的标记,已经校对发送的医嘱不处理
                        If k = 1 And Not rsOut Is Nothing Then
                            rsOut.Filter = "医嘱ID = " & str医嘱ID & " And 状态 < 3 "
                            If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                        End If
                    ElseIf PM_门诊医嘱清单 = glngModel Then
                        If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                        End If

                    End If
                ElseIf .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                    '中药配方
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                        lng中药组ID = .TextMatrix(i, gobjCOL.intCOL相关ID)          '中药配方组ID
                        lngLight = -1 '初始化
                    End If
                    '设置警示灯 取草药中最大警示值
                    If k >= 0 Then
                        If lngLight >= 0 Then
                            If arrLevel(k) > arrLevel(lngLight) Then
                                lngLight = k
                            End If
                        Else
                            lngLight = k
                        End If
                    End If
                End If
                '记录最高级别警示值
                If k >= 0 Then
                    If lngMaxWarn >= 0 Then
                        If arrLevel(k) > arrLevel(lngMaxWarn) Then
                            lngMaxWarn = k
                        End If
                    Else
                        lngMaxWarn = k
                    End If
                End If
            Else
                If glngModel = PM_门诊编辑 Then
                    '中药警示灯单独设置
                    If .RowData(i) = lng中药组ID And .RowData(i) <> 0 Then
                        strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                        '设置警示灯
                        If lngLight >= 0 And lngLight <= 4 Then
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                        Else
                            .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                            Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                        End If
                        
                        If glngModel = PM_门诊编辑 Then
                            '标记审查结果变化,以备更新数据库
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                blnNoSave = True    '标记为未保存
                            End If
                            '记录下禁忌药品 =1代表黑灯
                            If lngLight = 1 And Not rsOut Is Nothing Then
                                rsOut.Filter = "医嘱ID = " & lng中药组ID & " And 状态 < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                            End If
                        End If
                        lng中药组ID = 0
                        lngLight = -1
                    End If
                End If
            End If
        Next
        '医嘱清单中药配方警示灯处理
        If glngModel = PM_门诊医嘱清单 And Not rs中药 Is Nothing Then
            For i = .FixedRows To .Rows - 1
                '中药服法
                If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                    strOld = .Cell(flexcpData, i, gobjCOL.intCOL警示)
                    lngLight = -1
                    str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                    rs中药.Filter = "相关ID=" & str医嘱ID
                    
                    For j = 1 To rs中药.RecordCount
                        rsAdvice.Filter = "医嘱ID = '" & rs中药!id & "'"
                        If rsAdvice.RecordCount > 0 Then
                             k = CLng(rsAdvice!警示 & "")
                        Else
                             k = -1 '医嘱清单中药配方
                        End If
                        '设置警示灯 取草药中最大警示值
                        If k >= 0 Then
                            If lngLight >= 0 Then
                                If arrLevel(k) > arrLevel(lngLight) Then
                                    lngLight = k
                                End If
                            Else
                                lngLight = k
                            End If
                        End If
                        rs中药.MoveNext
                    Next
                    
                    '设置警示灯
                    If lngLight >= 0 And lngLight <= 4 Then
                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = CStr(lngLight)
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = frmIcons.imgPass.ListImages(arrLight(lngLight)).Picture
                    Else
                        .Cell(flexcpData, i, gobjCOL.intCOL警示) = ""
                        Set .Cell(flexcpPicture, i, gobjCOL.intCOL警示) = Nothing
                    End If
                    '警示灯更新到数据库
                    If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(lngLight >= 0 And lngLight <= 4, lngLight, "NULL") & ")"
                    End If
                        
                    '记录最高级别警示值
                    If lngLight >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(lngLight) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = lngLight
                            End If
                        Else
                            lngMaxWarn = lngLight
                        End If
                    End If
                    
                End If
            Next
        End If
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
        Next
        
    End With
    
    '返回审查结果
    OutAdviceCheckWarn_MK4 = lngMaxWarn
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function AdviceCheckWarn_DT(ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal str挂号单 As String, _
                Optional ByVal str医嘱IDs As String) As Boolean
'功能：调用大通用药监测系统对医嘱进行合理用药审查等相关功能
    Dim xmlbase As dt_base, xmlpre As dt_Pres
    Dim strTmp As String, arrTmp As Variant, curDate As Date
    Dim rsTmp As Recordset
    Dim i As Long, k As Long, blnDo As Boolean
    Dim str药品 As String, str给药途径 As String, str频率编码 As String, strXML As String
    Dim rsPati As ADODB.Recordset
    Dim strRetXML As String
    Dim blnIsHaveOut As Boolean '判断是否存在院外执行的药品
    Dim lng挂号ID As Long
    
    Set rsPati = GetPatiInfo_YF(lng病人ID, str挂号单, lng主页ID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    curDate = zlDatabase.Currentdate
    With xmlbase
        If str挂号单 = "" Then
            .dInHosCode = rsPati!住院号 & ""
            .dBedNo = "" & rsPati!当前床号
        Else
            .dInHosCode = ""
            .dBedNo = ""
            .pOutID = str挂号单
            lng挂号ID = NVL(rsPati!就诊Id, 0)
        End If
        .pCaseID = lng病人ID
        .dDoctCode = UserInfo.用户名
        .dDoctName = UserInfo.姓名
        .dDoctType = UserInfo.专业技术职务
        .dDeptCode = UserInfo.部门ID
        .dDeptName = UserInfo.部门名
        .mPresDate = curDate
        .pWeight = ""
        .pHeight = ""
        .pBirthday = NVL(rsPati!出生日期, vbNull)
        .pPatiName = rsPati!姓名
        .pSex = rsPati!性别
        .pStatms = ""
        .pEffect = ""
        .pBloodPress = ""
        .pLiverClean = ""
            
        '* 过敏源
        .pCaseCode1 = ""
        .pCaseName1 = ""
        .pCaseCode2 = ""
        .pCaseName2 = ""
        .pCaseCode3 = ""
        .pCaseName3 = ""
        
        If str挂号单 <> "" Then
            Set rsTmp = Get病人过敏记录(lng病人ID, 0)
        Else
            Set rsTmp = Get病人过敏记录(lng病人ID, lng主页ID)
        End If
        If rsTmp.RecordCount > 0 Then
            .pCaseCode1 = "" & rsTmp!药物ID
            .pCaseName1 = rsTmp!药物名
            rsTmp.MoveNext
            
            If Not rsTmp.EOF Then
                .pCaseCode2 = "" & rsTmp!药物ID
                .pCaseName2 = rsTmp!药物名
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pCaseCode3 = "" & rsTmp!药物ID
                    .pCaseName3 = rsTmp!药物名
                End If
            End If
        End If
        
        '* 诊断信息
        .pDiagnose1 = ""
        .pDiagnose2 = ""
        .pDiagnose3 = ""
        .pDiagnoseName1 = ""
        .pDiagnoseName2 = ""
        .pDiagnoseName3 = ""
        If str挂号单 <> "" Then
            Set rsTmp = Get病人诊断记录(lng病人ID, lng挂号ID, "1,11")
        Else
            Set rsTmp = Get病人诊断记录(lng病人ID, lng主页ID, "2,12")
        End If
        If rsTmp.RecordCount > 0 Then
            .pDiagnose1 = "" & rsTmp!编码
            .pDiagnoseName1 = "" & rsTmp!名称
            rsTmp.MoveNext
            If Not rsTmp.EOF Then
                .pDiagnose2 = "" & rsTmp!编码
                .pDiagnoseName2 = "" & rsTmp!名称
                rsTmp.MoveNext
                If Not rsTmp.EOF Then
                    .pDiagnose3 = "" & rsTmp!编码
                    .pDiagnoseName3 = "" & rsTmp!名称
                End If
            End If
        End If
        
        '* 病生理状态
        .pBsl1 = ""
        .pBsl2 = ""
        .pBsl3 = ""
        If str挂号单 <> "" Then
            strTmp = Get病人病生理情况(lng病人ID, 0)
        Else
            strTmp = Get病人病生理情况(lng病人ID, lng主页ID)
        End If
        
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            .pBsl1 = arrTmp(0)
            If UBound(arrTmp) > 0 Then .pBsl2 = arrTmp(1)
            If UBound(arrTmp) > 1 Then .pBsl3 = arrTmp(2)
        End If
    End With
        
    arrTmp = Array()
    Set rsTmp = GetAdviceInfo_YF(lng病人ID, lng主页ID, str挂号单, str医嘱IDs)
    With rsTmp
        For i = 1 To rsTmp.RecordCount
            Call Get频率信息_名称(rsTmp!频率 & "", 0, 0, "", IIf(rsTmp!诊疗类别 & "" = "7", 2, 1), str频率编码)
        
            xmlpre.PresID = rsTmp!医嘱ID & ""  '没有医嘱ID传病人ID
            If str挂号单 <> "" Then
                xmlpre.PresType = "mz"
                xmlpre.Current = 1
                xmlpre.Days = StrToXML(rsTmp!天数 & "")
            Else
                xmlpre.PresType = IIf(rsTmp!医嘱期效 & "" = "0", "L", "T")
                xmlpre.BTime = Format(rsTmp!开始时间 & "", "yyyy-MM-dd HH:mm:ss")
                xmlpre.ETime = Format(rsTmp!结束时间 & "", "yyyy-MM-dd HH:mm:ss")
                xmlpre.PresTime = Format(rsTmp!开嘱时间 & "", "yyyy-MM-dd HH:mm:ss")
            
            End If
            
            xmlpre.GeneralName = StrToXML(rsTmp!药品名称 & "")
            xmlpre.HosMediCode = rsTmp!收费细目id & ""
            xmlpre.MediName = StrToXML(rsTmp!药品名称 & "")
            xmlpre.DCL = FormatEx(rsTmp!单次用量 & "", 5)
            xmlpre.PCDM = StrToXML(str频率编码)
            xmlpre.Unit = StrToXML(rsTmp!单量单位 & "")
            xmlpre.GYTJ = rsTmp!用法ID & ""
            xmlpre.GroupNum = rsTmp!相关ID & ""
            
            strXML = MakePresXML(xmlpre, 1)
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = strXML
            .MoveNext
        Next
    End With
    
    If UBound(arrTmp) >= 0 Then
        On Error GoTo errH
        strXML = MakeXML(xmlbase, arrTmp, 1)
        WriteLog "" & glngModel, "AdviceCheckWarn_DT", strXML
        
        strTmp = dtywzxUI(28676, 1, strXML) '分析处方
        WriteLog "" & glngModel, "AdviceCheckWarn_DT", strTmp

    End If
    AdviceCheckWarn_DT = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    AdviceCheckWarn_DT = False
End Function

Public Function AdviceCheckWarn_TYT_YF(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, _
                    ByVal lngCmd As Long, Optional ByVal lngCurrAdviceID As Long, Optional str警示 As String = "", Optional ByVal str医嘱IDs As String) As Long
'功能：调用太元通系统中对医嘱进行合理用药审查等相关功能
'参数：lngCmd=
'       0-用药规范;1-获取医嘱审查结果,并填写警示灯
'       2-药品提示
'       3-医药知识库;4-系统配置;5-点击警示灯，获取警示详情
'
'返回
'str警示-返回警示串：格式：医嘱ID1:警示值1,医嘱ID2:警示值2
    Dim str医生编码 As String, str开嘱医生 As String, strDescription As String
    Dim strSQL As String, strOrderInfo As String, str频率编码 As String
    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim udtPatiOrder As PatientOrder
    Dim udtDrug As PatDrug, udtPatiDiag As PatDiagnosis
    Dim udtPatiSensitive As PatDrugSensitive, UdtPatiSymptom As PatSymptom
    Dim udtAuditResult As AuditResult

    Dim i As Long, k As Long
    Dim lng挂号ID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strTmp As String, strOld As String
    Dim arrTmp As Variant, colAuditResult As Collection
    
    On Error GoTo errH
    Screen.MousePointer = 11


    Select Case lngCmd
    Case 0   '0-用药规范

        gobjPass.getPdssPrescription

    Case 1  '1-获取医嘱审查结果,并填写警示灯
        Set rsPati = GetPatiInfo_YF(lng病人ID, str挂号单, lng主页ID)
        If rsPati.EOF Then Screen.MousePointer = 0: Exit Function
        '病人信息
        With udtPatiOrder
            '传人病人信息:病人ID,姓名,性别 1-女, 0-男, 2-不详，病人出生日期，格式 YYYY-MM-DD 不为空（必填）
            .PatientID = lng病人ID & ""
            .Pname = rsPati!姓名 & ""
            .pSex = IIf(rsPati!性别 & "" = "男", "0", IIf(rsPati!性别 & "" = "女", "1", "2"))
            .pdateOfBirth = Format(rsPati!出生日期, "yyyy-MM-dd")
            
            If str挂号单 <> "" Then
                lng挂号ID = NVL(rsPati!就诊Id, 0)
                '附加信息
                strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                        "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                        "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
    
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng病人ID, lng挂号ID)
                rsTmp.Filter = "项目名称='身高'"
                If rsTmp.RecordCount <> 0 Then .pHeight = IIf(Val(rsTmp!记录内容 & "") = 0, "", rsTmp!记录内容 & "")
                rsTmp.Filter = "项目名称='体重'"
                If rsTmp.RecordCount <> 0 Then .pWeight = IIf(Val(rsTmp!记录内容 & "") = 0, "", rsTmp!记录内容 & "")
                .PvisitID = rsPati!门诊号 & ""
                .SysFlag = "1"  '2-住院医生站，1-门诊医生站
            Else
                .pHeight = IIf(Val(rsPati!身高 & "") = 0, "", rsPati!身高 & "")
                .pWeight = IIf(Val(rsPati!体重 & "") = 0, "", rsPati!体重 & "")
                .PvisitID = rsPati!住院号 & ""
                .SysFlag = "2"  '2-住院医生站，1-门诊医生站
            End If
            
             '传人病人生理情况
            strTmp = Get病人病生理情况(lng病人ID, IIf(str挂号单 <> "", 0, lng主页ID))
            .isLact = IIf(InStr(strTmp, "哺乳期") > 0, "1", "0")    '是否哺乳，是为1，否为0 不为空
            .isPregnant = IIf(InStr(strTmp, "孕妇") > 0, "1", "0")    '是否孕妇，是为1 ，否为0 不为空
            .isLiverWhole = IIf(InStr(strTmp, "肝功能异常") > 0, "1", "0") '是否肝功异常 1-异常，0-正常 不为空
            .isKidneyWhole = IIf(InStr(strTmp, "肾功能异常") > 0, "1", "0") '是否肾功异常 1-异常，0-正常 不为空
                
            '登录医生信息
            .DoctDeptID = UserInfo.部门ID & ""
            .DoctDeptName = UserInfo.部门名 & ""
            .DoctID = UserInfo.编号 & ""
            .DoctName = UserInfo.姓名 & ""
            .DoctTitleID = GetDoctorTitleType(UserInfo.专业技术职务)
            .DoctTitleName = IIf(UserInfo.专业技术职务 = "", "其他职务", UserInfo.专业技术职务)
           
        End With

        '药品信息
        arrTmp = Array()
        
        Set rsAdvice = GetAdviceInfo_YF(lng病人ID, lng主页ID, str挂号单, str医嘱IDs)
        If rsAdvice.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
        With rsAdvice
            If NVL(!开嘱医生) <> "" Then
                str开嘱医生 = NVL(!开嘱医生)
                If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                str医生编码 = Sys.RowValue("人员表", str开嘱医生, "编号", "姓名")
            End If
            
            For i = 1 To .RecordCount
                udtDrug.drugID = !收费细目id & ""    'his 系统的药品代码不为空
                udtDrug.DrugName = StrToXML(!药品名称 & "")               'his 系统的药品名称不为空
                udtDrug.recMainNo = !相关ID & ""     'his 系统的医嘱组号，在一次就诊/住院中唯
                udtDrug.recSubNo = !医嘱序号 & ""      'his 系统的医嘱序号，在一次就诊/住院中唯
                udtDrug.dosage = !单次用量 & ""     'his 系统的医嘱药品使用剂量不为空
    
                udtDrug.doseUnits = !单量单位 & ""    'his 系统的医嘱药品剂量单位不为空
                udtDrug.administrationID = !用法ID & ""               'his 系统的医嘱途径代码不为空
                str频率编码 = GetFrequency(!间隔单位 & "", !频率次数 & "", !频率间隔 & "")
                udtDrug.performFreqDictID = StrToXML(str频率编码)   'his 系统的医嘱频次代码不为空
                udtDrug.performFreqDictText = !频率 & ""               'his 系统的医嘱执行频率描述不为空
    
                udtDrug.startDateTime = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")    'his 系统的医嘱开始时间,格式 YYYY-MM-DDHH: MM: SS 不为空
                udtDrug.stopDateTime = Format(!结束时间 & "", "yyyy-MM-dd HH:mm:ss")    'his 系统的医嘱结束时间,格式 YYYY-MM-DD HH: MM: SS
                udtDrug.doctorDept = !开嘱科室id & ""               'his 系统的开医嘱医生所在科室代码
                udtDrug.DoctorID = str医生编码                          'his 系统的开医嘱医生编码
                udtDrug.Doctor = str开嘱医生                         'his 系统的开医嘱医生姓名,
                udtDrug.isNew = "0"                             '新增医嘱值为1；否则为0
               
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                arrTmp(UBound(arrTmp)) = udtDrug
                .MoveNext
            Next
        End With
           
        If UBound(arrTmp) = -1 Then
            Screen.MousePointer = 0: Exit Function
        End If
        udtPatiOrder.PatDrugs = arrTmp

        '诊断
        arrTmp = Array()
  
        If str挂号单 <> "" Then
            Set rsTmp = Get病人诊断记录(lng病人ID, lng挂号ID, "1,11")
            strTmp = "门诊诊断"
        Else
            Set rsTmp = Get病人诊断记录(lng病人ID, lng主页ID, "2,12")   '西医住院，中医住院
            strTmp = "入院诊断"
        End If
        
        For i = 0 To rsTmp.RecordCount - 1
            udtPatiDiag.diagnosisID = rsTmp!编码 & ""       'his 系统的诊断编码
            udtPatiDiag.diagnosisName = rsTmp!名称 & ""     'his 系统的诊断名称
            udtPatiDiag.diagnosisType = strTmp      '系统的诊断类型，如门诊诊断、入院诊断等
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = udtPatiDiag
            rsTmp.MoveNext
        Next
        udtPatiOrder.PatDiagnoses = arrTmp
        
        
        '过敏
        arrTmp = Array()
        If str挂号单 <> "" Then
            Set rsTmp = Get病人过敏记录(lng病人ID, 0)
        Else
            Set rsTmp = Get病人过敏记录(lng病人ID, lng主页ID)
        End If
        For i = 0 To rsTmp.RecordCount - 1
            udtPatiSensitive.patOrderDrugSensitiveID = "0"          '固定值
            udtPatiSensitive.drugAllergenID = rsTmp!过敏源编码 & ""    '系统的过敏编码
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = udtPatiSensitive
            rsTmp.MoveNext
        Next
        udtPatiOrder.PatDrugSensitives = arrTmp

        '症状
        arrTmp = Array()
        If str挂号单 <> "" Then
            Set rsTmp = GetPatiSymptom(lng病人ID, lng挂号ID)
        Else
            Set rsTmp = GetPatiSymptom(lng病人ID, lng主页ID)
        End If
        For i = 0 To rsTmp.RecordCount - 1
            UdtPatiSymptom.symptomID = rsTmp!编码 & ""              'his 系统的症状编码
            UdtPatiSymptom.symptomName = rsTmp!名称 & ""            'his 系统的症状名称

            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
            arrTmp(UBound(arrTmp)) = UdtPatiSymptom
            rsTmp.MoveNext
        Next
        udtPatiOrder.PatSymptoms = arrTmp

        strOrderInfo = MakePatientOrderXml(udtPatiOrder)

        '医嘱信息审查接口调用"

        strDescription = gobjPass.checkDrugSecurityWS(strOrderInfo, "1")

        '审查结果处理
        '返回值顺及警示级别(高到低)：1— 禁忌（建议显示红色警示灯）；2— 慎用（建议显示黄色警示灯示）；3— 提示（建议显示蓝色警示灯）
        If strDescription = "" Then
            MsgBox "药嘱审查功能未执行，请检查太元通接口配置是否有误！", vbInformation + vbOKOnly, G_STR_PASS
            Screen.MousePointer = 0: Exit Function

        ElseIf strDescription = "-101" Then
            '-101：表示用户可以忽略该返回值，不做业务处理。
        Else
            If str警示 <> "-1" Then
                Set colAuditResult = AnalyzeReturnXml(strDescription)
                With rsAdvice
                    .MoveFirst
                    str警示 = ""
                    For i = 1 To rsAdvice.RecordCount
                        '获取警示灯
                        strTmp = !相关ID & "_" & !医嘱序号  '关键字格式:组医嘱号_医嘱序号
                        On Error Resume Next
                        udtAuditResult = colAuditResult(strTmp)
                        If Err.Number > 0 Then
                            strTmp = "未找到"
                        End If
                        Err.Clear: On Error GoTo 0
                        If strTmp <> "未找到" Then  '找到审核警示灯
                            str警示 = str警示 & "," & !医嘱ID & ":" & Val(udtAuditResult.alertLevel)
                        End If

                        .MoveNext
                    Next
                    If str警示 <> "" Then str警示 = Mid(str警示, 2)
                End With
            Else
                str警示 = ""
            End If
        End If

    Case 2    ' 2-药品提示

        '调用药品提示接口
        strSQL = "Select 收费细目id From 病人医嘱记录 Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lngCurrAdviceID)
        If rsTmp.RecordCount = 0 Then Exit Function
        gobjPass.getDrugExplain (rsTmp!收费细目id & "")
      
    Case 3    '3-在线医药知识库
        '调用在线医药知识库
        gobjPass.accessIFMI ("0")  '传入值固定为:"0",无返回值
    Case 4  '4-系统配置
        gobjPass.sysConfig
    Case 5    '5-获取警示详情
        gobjPass.getDrugAlertDetail
    End Select

    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub PASS4病生理情况(ByVal strData As String, ByRef lng哺乳 As Long, ByRef lng妊娠 As Long, _
        ByRef lng肝功 As Long, ByRef lng肾功 As Long, ByRef str妊娠日期 As String)
'功能: 获取病生理情况
'哺乳状态，取值：-1无法获取哺乳状态（默认）;0不是;1是
'妊娠状态，取值：-1-无法获取妊娠状态（默认）;0-不是;1-是
'妊娠开始日期，格式为yyyy-mm-dd。
'病人肝损害程度，取值： -1-不确定（默认）；0-无肝损害；1-肝功能不全；2-轻度肝损害；3-中度肝损害；4-重度肝损害*/
'病人肾损害程度，取值： -1-不确定（默认）；0-无肾损害；1-肾功能不全；2-轻度肾损害；3-中度肾损害；4-重度肾损害*/
    Dim i  As Integer
    
    lng哺乳 = 0: lng肝功 = 0: lng肾功 = 0: lng妊娠 = 0: str妊娠日期 = ""
    If strData = "" Then Exit Sub
    For i = LBound(Split(strData, ",")) To UBound(Split(strData, ","))
        If Split(strData, ",")(i) = "哺乳" Then
            lng哺乳 = 1
        ElseIf Split(strData, ",")(i) = "妊娠" Then
            lng妊娠 = 1
        ElseIf InStr("无肾损害,肾功能不全,轻度肾损害,中度肾损害,重度肾损害", Split(strData, ",")(i)) > 0 Then
            If Split(strData, ",")(i) = "无肾损害" Then
                lng肾功 = 0
            ElseIf Split(strData, ",")(i) = "肾功能不全" Then
                lng肾功 = 1
            ElseIf Split(strData, ",")(i) = "轻度肾损害" Then
                lng肾功 = 2
            ElseIf Split(strData, ",")(i) = "中度肾损害" Then
                lng肾功 = 3
            ElseIf Split(strData, ",")(i) = "重度肾损害" Then
                lng肾功 = 4
            End If
        ElseIf InStr("无肝损害,肝功能不全,轻度肝损害,中度肝损害,重度肝损害", Split(strData, ",")(i)) > 0 Then
            If Split(strData, ",")(i) = "无肝损害" Then
                lng肝功 = 0
            ElseIf Split(strData, ",")(i) = "肝功能不全" Then
                lng肝功 = 1
            ElseIf Split(strData, ",")(i) = "轻度肝损害" Then
                lng肝功 = 2
            ElseIf Split(strData, ",")(i) = "中度肝损害" Then
                lng肝功 = 3
            ElseIf Split(strData, ",")(i) = "重度肝损害" Then
                lng肝功 = 4
            End If
        ElseIf InStr(Split(strData, ",")(i), "妊娠日期|") > 0 Then
            str妊娠日期 = Split(Split(strData, ",")(i), "|")(1)
        End If
    Next
End Sub

Public Function AdviceCheckWarn_DTBS(ByVal bytFunc As Byte, Optional ByVal blnUpLoad As Boolean, Optional ByRef rsOut As ADODB.Recordset, _
    Optional ByRef objMap As clsPassMap, Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, _
    Optional ByVal str挂号单 As String, Optional ByVal str医嘱IDs As String) As Boolean
'功能：调用大通用药监测系统(BS版)对医嘱进行合理用药审查等相关功能
'
'参数：
'bytFunc=1-医生工作站;2-药房
'blnUpLoad:是否上传 T-是;F-否
'
'出参：
'      rsOut=禁忌药品说明
    Dim udtDetail As DTBS_DETAILS, xmlpre As dt_Pres
    Dim udtPati As DTBS_PATIENT
    Dim udt诊断 As DTBS_DIAGNOSE
    Dim udt过敏源 As DTBS_ALLERGIC
    Dim udtPres As DTBS_PRESCRIPTION
    Dim udtMedic As DTBS_MEDICINE
    
    Dim colTmp As Collection, colPres As Collection
    Dim str身高 As String, str体重 As String
    Dim strTmp As String, arrTmp As Variant, curDate As Date

    Dim i As Long, j As Long, blnDo As Boolean
    Dim lngTmp As Long, lngPos As Long
    Dim str药品 As String, str给药途径 As String, strXML As String
    Dim str相关ID As String
    Dim strSQL As String
    
    Dim rsPati As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsSub As ADODB.Recordset
    Dim rsRet As ADODB.Recordset
    
    Dim strRetXML As String
    Dim blnIsHaveOut As Boolean '判断是否存在院外执行的药品
    Dim lng挂号ID As Long
    Dim byt场合 As Byte
    
    If bytFunc = 1 Then
        lng病人ID = gobjPati.lng病人ID
        lng主页ID = gobjPati.lng主页ID
        str挂号单 = gobjPati.str挂号单
    End If
    
    Set rsPati = GetPatiInfo_YF(lng病人ID, str挂号单, lng主页ID)
    If rsPati Is Nothing Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    If str挂号单 <> "" Then
        lng挂号ID = Val(rsPati!就诊Id & "")
        strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                        "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                        "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
                        
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng病人ID, lng挂号ID)
        rsPatiInfo.Filter = "项目名称='身高'"
        If rsPatiInfo.RecordCount <> 0 Then str身高 = NVL(rsPatiInfo!记录内容)
        rsPatiInfo.Filter = "项目名称='体重'"
        If rsPatiInfo.RecordCount <> 0 Then str体重 = NVL(rsPatiInfo!记录内容)
    Else
        str身高 = rsPati!身高 & ""
        str体重 = rsPati!体重 & ""
    End If
    
    curDate = zlDatabase.Currentdate
    
    With udtPati
        .str姓名 = rsPati!姓名 & ""
        .str是否婴儿 = 0  '0:  非婴幼儿， 1： 是婴幼儿
        .str出生日期 = rsPati!出生日期 & ""
        .str性别 = rsPati!性别 & ""
        .str体重 = str体重
        .str身高 = str身高
        .str身份证号 = rsPati!身份证号 & ""
        .str卡类型 = ""
        .str卡号 = ""
        .str怀孕时间单位 = ""
        .str怀孕时间 = ""
        
        '过敏源
        Set colTmp = New Collection
        Set rsTmp = Get病人过敏记录(lng病人ID, lng主页ID)
        For i = 1 To rsTmp.RecordCount
            If "" & rsTmp!药物ID <> "" Then
                With udt过敏源
                    .str过敏类型 = "5"   '1=大通药品大类 2=大通药品成份 5-HIS药品代码
                    .str过敏源名称 = rsTmp!药物名
                    .str过敏源代码 = "" & rsTmp!药物ID
                End With
                colTmp.Add udt过敏源, "_" & i
            End If
            rsTmp.MoveNext
        Next
        Set .col过敏源s = colTmp
        
        '诊断记录
        Set colTmp = New Collection
        Select Case glngModel
        Case PM_门诊编辑
            If Not gobjDiags Is Nothing Then
                For i = 1 To gobjDiags.Count
                    With udt诊断
                        If gobjDiags.Item(i).str诊断描述 <> "" Then
                            If gobjDiags.Item(i).str疾病编码 <> "" Then
                                .str诊断类型 = "2" '2=IDC10代码
                                .str诊断代码 = gobjDiags.Item(i).str疾病编码
                            Else
                                .str诊断类型 = "0"
                                .str诊断代码 = gobjDiags.Item(i).str诊断编码
                            End If
                            .str诊断名称 = gobjDiags.Item(i).str诊断描述
                        End If
                    End With
                    colTmp.Add udt诊断, "_" & colTmp.Count + 1
                Next
            End If
        Case Else
            Set rsTmp = Get病人诊断记录(lng病人ID, IIf(str挂号单 <> "", lng挂号ID, lng主页ID), IIf(str挂号单 <> "", "1,11", "2,12"))
            For i = 1 To rsTmp.RecordCount
                With udt诊断
                    If rsTmp!疾病ID & "" <> "" Then
                         .str诊断类型 = "2" '2=IDC10代码
                    Else
                        .str诊断类型 = "0" '0=其他
                    End If
                    .str诊断代码 = "" & rsTmp!编码
                    .str诊断名称 = "" & rsTmp!名称
                End With
                colTmp.Add udt诊断, "_" & colTmp.Count + 1
                rsTmp.MoveNext
            Next
            If colTmp.Count = 0 Then
                colTmp.Add udt诊断, "_" & colTmp.Count + 1
            End If
        End Select
        '病生理
        strTmp = Get病人病生理情况(lng病人ID, IIf(str挂号单 <> "", 0, lng主页ID))
        If strTmp <> "" Then
            Set rsSub = GetRS("病生理情况", "编码,名称", strTmp, "名称", 0, 1)
            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                With udt诊断
                    .str诊断类型 = "1" '1=病生理状态
                    .str诊断名称 = arrTmp(i)
                    rsSub.Filter = "名称='" & arrTmp(i) & "'"
                    If Not rsSub.EOF Then .str诊断代码 = rsSub!编码 & ""
                End With
                colTmp.Add udt诊断, "_" & colTmp.Count + 1
            Next
           
        End If
        Set .col诊断s = colTmp
        '检验检测单节点
        'Set colTmp = New Collection
        'strTmp = ""
    End With
    
    With udtDetail
        .str是否上传 = IIf(blnUpLoad, "1", "0")  '默认不上传处方分析
        .strHIS系统时间 = Format(curDate, "YYYY-MM-dd hh:mm:ss")
        If str挂号单 <> "" Then
            .str门诊住院标识 = "op"
            .str就诊类型 = DTBS_GetTreatType(1, lng挂号ID)
            .str就诊号 = lng挂号ID & ""
        Else
            .str门诊住院标识 = "ip"
            .str就诊类型 = DTBS_GetTreatType(2, lng病人ID, lng主页ID)
            .str就诊号 = rsPati!住院号 & ""
        End If
        .udt病人信息 = udtPati
    End With

    '药品信息
    Select Case glngModel
    Case PM_门诊编辑, PM_门诊医嘱清单, PM_住院编辑, PM_住院医嘱清单
        Set rsTmp = CreateAdviceRS(rsOut)
    Case PM_部门发药, PM_处方发药, PM_PIVA管理
        Set rsTmp = CreateAdviceRS(, lng病人ID, lng主页ID, str挂号单, str医嘱IDs)
        byt场合 = 1
    End Select
    
    If rsTmp.RecordCount = 0 Then AdviceCheckWarn_DTBS = True: Exit Function    '医嘱下达界面没有下达药品时允许保存
    
    With rsTmp
        Set colPres = New Collection
        .MoveFirst
        For i = 1 To rsTmp.RecordCount
            udtPres.str处方号 = !医嘱ID & ""
            udtPres.str处方理由 = ""
            udtPres.str开嘱医生代码 = !开嘱医生编码 & ""
            udtPres.str开嘱医生姓名 = !开嘱医生 & ""
            udtPres.str开嘱科室代码 = !开嘱科室id & ""
            udtPres.str开嘱科室名称 = !开嘱科室 & ""
            udtPres.str处方时间 = Format(!开嘱时间 & "", "YYYY-MM-DD HH:MM:SS")
            udtPres.str是否紧急处方 = IIf(!标志 & "" = "1", "1", "0")
            udtPres.str是否新开处方 = IIf(Val(!医嘱状态 & "") < 2, "1", "0") '暂存,新开 都传1
            udtPres.str是否当前处方 = IIf(byt场合 = 1, 1, IIf(Val(!医嘱状态 & "") < 2, "1", "0")) '0 历史处方 1 当前处方新开处方(=1时处方分析才会返回审查详情)
            udtPres.Str医嘱类型 = IIf(!医嘱期效 & "" = "1", "T", "L") '住院处方有效(hosp_flag = ip)L:长期医嘱 T: 临时医嘱
            
            Set colTmp = New Collection
            udtMedic.str商品名 = DTBS_StrToXML(!药品名称 & "")
            udtMedic.str医院药品代码 = !药品ID & ""
            udtMedic.str配液单号 = ""
            udtMedic.str配液单组号 = ""
            udtMedic.str医保代码 = ""
            udtMedic.str规格 = !规格 & ""
            udtMedic.str组号 = !相关ID & ""
            udtMedic.str用药理由 = !用药理由 & ""   '非抗菌药物为空
            '单量，单量单位
            udtMedic.str单次量单位 = !单量单位 & ""
            udtMedic.str单次量 = !单次用量 & ""
            udtMedic.str频次代码 = !频率编码 & ""
            udtMedic.str给药途径代码 = !用法ID & ""
            udtMedic.str用药开始时间 = Format(!开始时间 & "", "yyyy-MM-dd HH:mm:ss")
            udtMedic.str用药结束时间 = Format(!结束时间 & "", "yyyy-MM-dd HH:mm:ss")
            udtMedic.str服药天数 = !天数 & ""   'OP 门诊处方有效
            udtMedic.str是否预防用药 = IIf(!用药目的 & "" = "1", 1, 0)
            udtMedic.str手术单号 = ""
            udtMedic.str签名医师工号 = ""
            udtMedic.str授权时间 = ""
            udtMedic.str允许用药时间 = ""
            udtMedic.str允许用药次数 = ""
            colTmp.Add udtMedic, "_" & colTmp.Count

            Set udtPres.col药品信息 = colTmp
            colPres.Add udtPres, "_" & colPres.Count + 1
            .MoveNext
        Next
        .Filter = "离院带药=1"
        blnIsHaveOut = .RecordCount > 0
    End With
    
    Set udtDetail.col处方信息 = colPres
    
    If udtPres.col药品信息.Count > 0 Then
        On Error GoTo errH
        strXML = DTBS_MakePresXML(udtDetail)
        WriteLog "" & glngModel, "AdviceCheckWarn_DTBS", "功能号:" & DTBS_处方分析
        WriteLog "" & glngModel, "AdviceCheckWarn_DTBS", "DetailXML:" & strXML
    
        lngTmp = CRMS_UI(DTBS_处方分析, gstrBaseXml, strXML, strRetXML)
        strRetXML = StrConvToNormal(strRetXML)
        WriteLog "" & glngModel, "AdviceCheckWarn_DTBS", "处方分析返回值:" & lngTmp & vbCrLf & "RetXML:" & strRetXML
      
        If blnUpLoad Then AdviceCheckWarn_DTBS = True: Exit Function
       
        If glngModel = PM_门诊编辑 Then
            '内容分为三种：0、1、2，3分别代表没有问题，其他问题，一般问题和严重问题
            '1、 保存处方信息的过程中，从大通用药安全监测系统得到的返回值如果是0或者1，表示当前处方中没有严重问题。（0表示没有问题；1表示其他问题，其他问题都是对医生的提示）
            '2、 得到的返回值如果是2，3，表示当前处方有严重问题，需要对该处方进行拦截。
           If lngTmp = 3 And gbytBlackLamp = 0 Then
                MsgBox "用药监测系统发现当前医嘱存在禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                Exit Function
            ElseIf ((lngTmp = 2 Or lngTmp = 3) And gbytBlackLamp = 1) Then
                If gbytReason = 1 Then
                    '记录下禁忌药品 警示值=2\3代表禁忌 且 只针对未校对医嘱进行禁忌药品说明原因的标记,已经校对发送的医嘱不处理
                    Set rsRet = ReadXML(strRetXML)
                    If Not rsRet Is Nothing Then
                        For i = 1 To rsRet.RecordCount
                            If Not rsOut Is Nothing Then
                                If rsRet!警示值 >= 2 Then
                                    rsOut.Filter = "医嘱ID = " & rsRet!医嘱ID & " And 状态 < 3 "
                                    If rsOut.RecordCount = 1 Then
                                        rsOut!是否禁忌 = 1
                                    End If
                                End If
                            End If
                            rsRet.MoveNext
                        Next
                    End If
                    If Not AddDrugReason(objMap, rsOut) Then Exit Function
                Else
                    If MsgBox("用药监测系统发现当前医嘱存在禁忌用药，是否继续?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf lngTmp = 1 Then
                If MsgBox("用药监测系统发现当前医嘱存在其他问题，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        ElseIf glngModel = PM_住院编辑 Then
            If lngTmp = 3 And gbytBlackLamp = 0 Then
                If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                    If MsgBox("用药监测系统发现有院外执行的药品存在禁忌用药，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "用药监测系统发现当前医嘱存在禁忌用药，操作不能继续!", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf (lngTmp = 2 Or lngTmp = 3) And gbytBlackLamp = 1 Then
                If gbytReason = 1 Then
                    Set rsRet = ReadXML(strRetXML)
                    If Not rsRet Is Nothing Then
                        For i = 1 To rsRet.RecordCount
                            If Not rsOut Is Nothing Then
                                If rsRet!警示值 >= 2 Then
                                    rsOut.Filter = "医嘱ID = " & rsRet!医嘱ID & " And 状态 < 3 "
                                    If rsOut.RecordCount = 1 Then
                                        rsOut!是否禁忌 = 1
                                    End If
                                End If
                            End If
                            rsRet.MoveNext
                        Next
                    End If
                    If Not AddDrugReason(objMap, rsOut) Then Exit Function
                Else
                    If MsgBox("用药监测系统发现当前医嘱存在禁忌用药，是否继续?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf lngTmp = 1 Then
                If MsgBox("用药监测系统发现当前医嘱存在其他问题，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If

    AdviceCheckWarn_DTBS = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    AdviceCheckWarn_DTBS = False
End Function

Public Function CreateAdviceRS(Optional ByRef rsOut As ADODB.Recordset, Optional ByVal lng病人ID As String, _
    Optional ByVal lng主页ID As String, Optional ByVal str挂号单 As String, _
    Optional ByVal str医嘱IDs As String) As ADODB.Recordset
'功能;构造医嘱记录集
    Dim i As Long, k As Long, lngCount As Long, lngPos As Long
    Dim blnDo As Boolean, blnIsHaveOut As Boolean
    Dim str药品 As String, str医嘱ID As String, str相关ID As String
    Dim str开嘱时间 As String
    Dim str期效 As String, str单量 As String, str单量单位 As String, str频率 As String
    Dim str给药途径 As String, str频率编码 As String, str用法 As String, str用法ID As String, str开始时间 As String, str结束时间 As String
    Dim str开嘱科室Tag As String, str开嘱科室ID As String, str诊疗项目IDs As String, str药品ID As String
    Dim str开嘱医生Tag As String, str开嘱医生 As String
    Dim str总量 As String, str总量单位 As String, str状态 As String
    Dim str诊疗ID, str收费细目ID As String
    
    Dim rsAdvice As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rs频率 As ADODB.Recordset
    Dim rs开嘱医生 As ADODB.Recordset
    Dim rs开嘱科室  As ADODB.Recordset
    Dim rs规格 As ADODB.Recordset
    Dim rs药品 As ADODB.Recordset
    
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    Set rsAdvice = InitAdviceRS(FUN_医嘱信息_DTBS)
    
    Select Case glngModel
    Case PM_门诊编辑, PM_住院编辑
        '启用了禁忌药品说明参数;场合为门诊编辑\住院编辑;审查功能
        If (glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑) And gbytReason = 1 Then
            Set rsOut = InitAdviceRS(FUN_输出内容)
        End If
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_门诊编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                            And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-DD") = Format(curDate, "yyyy-MM-DD")
                ElseIf glngModel = PM_住院编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                    blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" _
                            Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                End If
                
                If blnDo Then
                    str诊疗ID = .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                    If InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) = 0 Then
                        If InStr("," & str诊疗项目IDs & ",", "," & str诊疗ID & ",") = 0 Then
                            str诊疗项目IDs = str诊疗项目IDs & "," & str诊疗ID
                        End If
                    End If
                    str医嘱ID = CStr(.RowData(i))
                    
                    '取药品名称
                    If InStr(",5,6,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 Then
                        str药品 = .TextMatrix(i, gobjCOL.intCOL药品名称)
                    Else
                        str药品 = .TextMatrix(i, gobjCOL.intCOL医嘱内容) '中药名称
                    End If
                    
                    '取药品给药途径
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str用法 = ""    '一并给药不重复取
                    If str用法 = "" Then
                        k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                        If k <> -1 Then
                            If .TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7" Then
                                str用法 = .TextMatrix(k, gobjCOL.intCOL用法)
                            Else
                                str用法 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                            End If
                            str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                        End If
                    End If
    
                    '开嘱科室名称
                    str开嘱科室ID = .TextMatrix(i, gobjCOL.intCOL开嘱科室ID)
                    If InStr("," & str开嘱科室Tag & ",", "," & str开嘱科室ID & ",") = 0 Then
                        str开嘱科室Tag = str开嘱科室Tag & "," & str开嘱科室ID
                    End If
                   
                    '开嘱医生
                    str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                    If InStr("," & str开嘱医生Tag & ",", "," & str开嘱医生 & ",") = 0 Then
                        str开嘱医生Tag = str开嘱医生Tag & "," & str开嘱医生
                    End If
                    
                    str开始时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd HH:MM:SS")
'
                    str开嘱时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd HH:mm:ss")         '处方时间（YYYY-MM-DD HH:mm:SS）
                    '单量，单量单位
                    str单量 = .TextMatrix(i, gobjCOL.intCOL单量)
                    str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                    str总量 = .TextMatrix(i, gobjCOL.intCOL总量)
                    str总量单位 = .TextMatrix(i, gobjCOL.intcol总量单位)
                    
                    str药品ID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                    
                    If glngModel = PM_门诊编辑 Then
                        str结束时间 = ""
                        str期效 = "1"
                    ElseIf glngModel = PM_住院编辑 Then
                        str结束时间 = Format(.Cell(flexcpData, i, gobjCOL.intCOL终止时间), "yyyy-MM-dd HH:MM:SS")
                        '判断是否是院外执行的药品
                        If Val(.TextMatrix(i, gobjCOL.intCOL执行性质)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))), gobjCOL.intCOL执行性质)) = 5 Then
                            blnIsHaveOut = True
                        End If
                        str期效 = IIf(.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱", 0, 1)
                    End If
                    
                    If InStr(";" & str频率 & ";", ";" & .TextMatrix(i, gobjCOL.intCOL频率) & "," & IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1) & ";") = 0 Then
                        str频率 = str频率 & ";" & .TextMatrix(i, gobjCOL.intCOL频率) & "," & IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1)
                    End If
                    
                    '禁忌说明
                    If Not rsOut Is Nothing Then
                        If InStr(",5,6,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 Then
                        '西药,中成药
                            rsOut.AddNew
                            rsOut!医嘱ID = CLng(str医嘱ID)
                            rsOut!禁忌药品说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                            rsOut!状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                            rsOut!药品名称 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                            rsOut.Update
                        ElseIf Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then
                        '中药配方  禁忌说明保存在用药服法上
                            k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                            If k <> -1 Then
                                rsOut.AddNew
                                rsOut!医嘱ID = CLng(.RowData(k) & "")
                                rsOut!禁忌药品说明 = .TextMatrix(k, gobjCOL.intCol禁忌药品说明)
                                rsOut!状态 = .TextMatrix(k, gobjCOL.intCOL状态)
                                rsOut!药品名称 = .TextMatrix(k, gobjCOL.intCOL医嘱内容)
                                rsOut.Update
                            End If
                        End If
                    End If
                        
                    '----------------------------------------------------------
                    rsAdvice.AddNew
                    rsAdvice!医嘱ID = str医嘱ID
                    rsAdvice!相关ID = .TextMatrix(i, gobjCOL.intCOL相关ID)
                    rsAdvice!医嘱期效 = str期效
                    rsAdvice!医嘱序号 = lngCount + 1
                    rsAdvice!开嘱科室id = str开嘱科室ID
                    rsAdvice!开嘱医生 = str开嘱医生
                    rsAdvice!诊疗项目ID = str诊疗ID
                    rsAdvice!药品ID = str药品ID
                    rsAdvice!药品名称 = str药品
                    rsAdvice!医嘱状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                    rsAdvice!单次用量 = str单量
                    rsAdvice!单量单位 = str单量单位
                    rsAdvice!频率 = .TextMatrix(i, gobjCOL.intCOL频率)
                    rsAdvice!用法 = str用法
                    rsAdvice!用法ID = str给药途径
                    rsAdvice!开嘱时间 = str开嘱时间
                    rsAdvice!开始时间 = str开始时间
                    rsAdvice!结束时间 = str结束时间
                    rsAdvice!总量 = str总量
                    rsAdvice!总量单位 = str总量单位
                    rsAdvice!天数 = .TextMatrix(i, gobjCOL.intCOL天数)
                    rsAdvice!医生嘱托 = .TextMatrix(i, gobjCOL.intcol医嘱嘱托)
                    rsAdvice!用药目的 = .TextMatrix(i, gobjCOL.intcol用药目的)
                    rsAdvice!用药理由 = .TextMatrix(i, gobjCOL.intcol用药理由)
                    rsAdvice!诊疗类别 = .TextMatrix(i, gobjCOL.intCOL诊疗类别)
                    rsAdvice!标志 = .TextMatrix(i, gobjCOL.intCol标志)
                    rsAdvice!离院带药 = IIf(blnIsHaveOut, 1, 0)
                    rsAdvice.Update
                    '----------------------------------------------------------------------------
                End If
            Next
        End With
    Case PM_门诊医嘱清单, PM_住院医嘱清单
        Set rsTmp = GetAdviceInfo_YF(gobjPati.lng病人ID, gobjPati.lng主页ID, gobjPati.str挂号单, , 1)
        With rsTmp
            If rsTmp.RecordCount = 0 Then Set CreateAdviceRS = rsAdvice: Exit Function
            For i = 1 To .RecordCount
                If glngModel = PM_门诊医嘱清单 Then
                    blnDo = InStr(",5,6,7,", "," & !诊疗类别 & ",") > 0 And Val(!收费细目id & "") <> 0 And Format(!开嘱时间 & "", "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                ElseIf glngModel = PM_住院医嘱清单 Then
                    blnDo = InStr(",5,6,7,", "," & !诊疗类别 & ",") > 0 And Not InStr(",4,8,9,", "," & !医嘱状态 & ",") > 0
                End If
                If blnDo Then
              
                    If InStr(",5,6,7,", "," & !诊疗类别 & ",") > 0 And Not InStr(",4,8,9,", "," & !医嘱状态 & ",") > 0 And Val(!收费细目id & "") = 0 Then
                        If InStr("," & str诊疗项目IDs & ",", "," & !诊疗项目ID & ",") = 0 Then
                            str诊疗项目IDs = str诊疗项目IDs & "," & !诊疗项目ID
                        End If
                    End If
                    '开嘱医生
                    str开嘱医生 = !开嘱医生 & ""
                    If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                    If InStr("," & str开嘱医生Tag & ",", "," & str开嘱医生 & ",") = 0 Then
                        str开嘱医生Tag = str开嘱医生Tag & "," & str开嘱医生
                    End If
                    
                    If gobjPati.str挂号单 <> "" Then
                        str总量单位 = !门诊单位 & ""
                    Else
                        str总量单位 = !住院单位 & ""
                    End If
                    
                    If InStr(";" & str频率 & ";", ";" & !频率 & "," & IIf(!诊疗类别 & "" = "7", 2, 1) & ";") = 0 Then
                        str频率 = str频率 & ";" & !频率 & "," & IIf(!诊疗类别 = "7", 2, 1)
                    End If
                    
                    rsAdvice.AddNew
                    rsAdvice!医嘱ID = !医嘱ID & ""
                    rsAdvice!相关ID = !相关ID & ""
                    rsAdvice!医嘱期效 = !医嘱期效 & ""
                    rsAdvice!医嘱序号 = lngCount + 1
                    rsAdvice!开嘱科室id = !开嘱科室id & ""
                    rsAdvice!开嘱科室 = !开嘱科室 & ""
                    rsAdvice!开嘱医生 = str开嘱医生
                    rsAdvice!诊疗项目ID = !诊疗项目ID & ""
                    rsAdvice!药品ID = !收费细目id & ""
                    rsAdvice!药品名称 = !药品名称 & ""
                    rsAdvice!医嘱状态 = !医嘱状态 & ""
                    rsAdvice!单次用量 = !单次用量 & ""
                    rsAdvice!单量单位 = !单量单位 & ""
                    rsAdvice!频率 = !频率 & ""
                    rsAdvice!用法 = !用法 & ""
                    rsAdvice!用法ID = !用法ID & ""
                    rsAdvice!开嘱时间 = !开嘱时间 & ""
                    rsAdvice!开始时间 = !开始时间 & ""
                    rsAdvice!结束时间 = !结束时间 & ""
                    rsAdvice!总量 = !总量 & ""
                    rsAdvice!总量单位 = str总量单位
                    rsAdvice!天数 = !天数 & ""
                    rsAdvice!医生嘱托 = !医生嘱托 & ""
                    rsAdvice!用药目的 = !用药目的 & ""
                    rsAdvice!用药理由 = !用药理由 & ""
                    rsAdvice!诊疗类别 = !诊疗类别 & ""
                    rsAdvice!规格 = !规格 & ""
                    rsAdvice!标志 = !标志 & ""
                    rsAdvice.Update
                End If
                .MoveNext
            Next
        End With
    Case PM_PIVA管理, PM_部门发药, PM_处方发药
        Set rsTmp = GetAdviceInfo_YF(lng病人ID, lng主页ID, str挂号单)
        With rsTmp
            If rsTmp.RecordCount = 0 Then Set CreateAdviceRS = rsAdvice: Exit Function
            For i = 1 To .RecordCount
            
                If Val(!收费细目id & "") = 0 Then
                    If InStr("," & str诊疗项目IDs & ",", "," & !诊疗项目ID & ",") = 0 Then
                        str诊疗项目IDs = str诊疗项目IDs & "," & !诊疗项目ID
                    End If
                End If
                '开嘱医生
                str开嘱医生 = !开嘱医生 & ""
                If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                If InStr("," & str开嘱医生Tag & ",", "," & str开嘱医生 & ",") = 0 Then
                    str开嘱医生Tag = str开嘱医生Tag & "," & str开嘱医生
                End If
                
                If str挂号单 <> "" Then
                    str总量单位 = !门诊单位 & ""
                Else
                    str总量单位 = !住院单位 & ""
                End If
                
                If InStr(";" & str频率 & ";", ";" & !频率 & "," & IIf(!诊疗类别 & "" = "7", 2, 1) & ";") = 0 Then
                    str频率 = str频率 & ";" & !频率 & "," & IIf(!诊疗类别 = "7", 2, 1)
                End If
                
                rsAdvice.AddNew
                rsAdvice!医嘱ID = !医嘱ID & ""
                rsAdvice!相关ID = !相关ID & ""
                rsAdvice!医嘱期效 = !医嘱期效 & ""
                rsAdvice!医嘱序号 = lngCount + 1
                rsAdvice!开嘱科室id = !开嘱科室id & ""
                rsAdvice!开嘱科室 = !开嘱科室 & ""
                rsAdvice!开嘱医生 = str开嘱医生
                rsAdvice!诊疗项目ID = !诊疗项目ID & ""
                rsAdvice!药品ID = !收费细目id & ""
                rsAdvice!药品名称 = !药品名称 & ""
                rsAdvice!医嘱状态 = !医嘱状态 & ""
                rsAdvice!单次用量 = !单次用量 & ""
                rsAdvice!单量单位 = !单量单位 & ""
                rsAdvice!频率 = !频率 & ""
                rsAdvice!用法 = !用法 & ""
                rsAdvice!用法ID = !用法ID & ""
                rsAdvice!开嘱时间 = !开嘱时间 & ""
                rsAdvice!开始时间 = !开始时间 & ""
                rsAdvice!结束时间 = !结束时间 & ""
                rsAdvice!总量 = !总量 & ""
                rsAdvice!总量单位 = str总量单位
                rsAdvice!天数 = !天数 & ""
                rsAdvice!医生嘱托 = !医生嘱托 & ""
                rsAdvice!用药目的 = !用药目的 & ""
                rsAdvice!用药理由 = !用药理由 & ""
                rsAdvice!诊疗类别 = !诊疗类别 & ""
                rsAdvice!规格 = !规格 & ""
                rsAdvice!标志 = !标志 & ""
                rsAdvice.Update

                .MoveNext
            Next
        End With
    End Select
    
    '附加数据提取
    If rsAdvice.RecordCount > 0 Then
        
        rsAdvice.MoveFirst
        Select Case glngModel
        
        Case PM_门诊编辑, PM_门诊医嘱清单, PM_住院编辑, PM_住院医嘱清单, PM_PIVA管理, PM_部门发药, PM_处方发药
            If str诊疗项目IDs <> "" Then
                str诊疗项目IDs = Mid(str诊疗项目IDs, 2)
                Set rs药品 = GetRS("药品规格", "药名id,药品id", str诊疗项目IDs, "药名id")
            End If
            If str频率 <> "" Then Set rs频率 = GetRS("诊疗频率项目", "编码, 名称, 适用范围", str频率, "名称, 适用范围", 1, 2)
            If str开嘱科室Tag <> "" Then Set rs开嘱科室 = GetRS("部门表", "ID,名称", str开嘱科室Tag)
            If str开嘱医生Tag <> "" Then Set rs开嘱医生 = GetRS("人员表", "编号,姓名", str开嘱医生Tag, "姓名", 0, 1)
            For i = 1 To rsAdvice.RecordCount
                 '长期医嘱按品种下达时,任意取一个药品Id
                If Val(rsAdvice!药品ID & "") = 0 And Val(rsAdvice!医嘱期效 & "") = 0 Then
                    If Not rs药品 Is Nothing Then
                        rs药品.Filter = "药名ID =" & rsAdvice!诊疗项目ID
                        If Not rs药品.EOF Then rsAdvice!药品ID = rs药品!药品ID & ""
                    End If
                End If
                
                If InStr("," & str收费细目ID & ",", "," & rsAdvice!药品ID & ",") = 0 Then
                    str收费细目ID = str收费细目ID & "," & rsAdvice!药品ID
                End If
                
                If Not rs频率 Is Nothing Then
                    rs频率.Filter = "名称 ='" & rsAdvice!频率 & "' And 适用范围=" & IIf(rsAdvice!诊疗类别 & "" = "7", 2, 1)
                    If Not rs频率.EOF Then rsAdvice!频率编码 = rs频率!编码 & ""
                End If
                
                If Not rs开嘱医生 Is Nothing Then
                    rs开嘱医生.Filter = "姓名='" & rsAdvice!开嘱医生 & "'"
                    If Not rs开嘱医生.EOF Then rsAdvice!开嘱医生编码 = rs开嘱医生!编号 & ""
                End If
                If Not rs开嘱科室 Is Nothing Then
                    rs开嘱科室.Filter = "ID =" & rsAdvice!开嘱科室id
                    If Not rs开嘱科室.EOF Then rsAdvice!开嘱科室 = rs开嘱科室!名称 & ""
                End If
                
                rsAdvice.MoveNext
            Next
            
            If str收费细目ID <> "" Then
                str收费细目ID = Mid(str收费细目ID, 2)
                Set rs规格 = GetRS("收费项目目录", "ID,规格", str收费细目ID)
                rsAdvice.MoveFirst
                For i = 1 To rsAdvice.RecordCount
                    If Not rs规格 Is Nothing Then
                        rs规格.Filter = "ID =" & rsAdvice!药品ID
                        If Not rs规格.EOF Then rsAdvice!规格 = rs规格!规格 & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
        End Select
        rsAdvice.MoveFirst
    End If
    Set CreateAdviceRS = rsAdvice
End Function

Private Function GetDrugInfo_MK4(ByVal str挂号单 As String, ByVal strAdvice As String, Optional ByVal lngPatiID As Long, Optional ByVal lng主页ID As Long) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.相关id, a.处方序号, d.名称 As 用法,a.执行科室ID " & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱记录 B, 诊疗项目目录 D" & vbNewLine & _
            "Where a.相关id = b.Id(+) And b.诊疗项目id = d.Id(+) " & IIf(lngPatiID = 0, "And A. 挂号单 = [1]", "And A.病人ID = [3] And A.主页ID =[4]") & " And a.相关id <> 0 And Instr([2], ',' || a.Id || ',') > 0" & vbNewLine & _
            "Order By a.序号"
    Set GetDrugInfo_MK4 = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", str挂号单, "," & strAdvice & ",", lngPatiID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
