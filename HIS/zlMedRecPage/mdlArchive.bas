Attribute VB_Name = "mdlArchive"
Option Explicit
'变量
Public gblnCheck As Boolean '控制CheckBox不能勾选
Public gOldwinproc As Long '原始消息句柄
Public Enum GeneralCtrlArchive
    '基本信息
    GCA_付款方式 = 0
    GCA_健康卡号 = 1
    GCA_住院次数 = 2
    GCA_病案号 = 3
    GCA_姓名 = 4
    GCA_性别 = 5
    GCA_出生日期 = 6
    GCA_年龄 = 7
    GCA_国籍 = 8
    GCA_体重 = 9
    GCA_身高 = 10
    GCA_不足周岁年龄 = 11
    GCA_新生儿体重 = 12
    GCA_新生儿入院体重 = 13
    GCA_出生地点 = 14
    GCA_籍贯 = 15
    GCA_民族 = 16
    GCA_身份证号 = 17
    GCA_职业 = 18
    GCA_婚姻 = 19
    GCA_家庭地址 = 20
    GCA_家庭电话 = 21
    GCA_家庭邮编 = 22
    GCA_户口地址 = 23
    GCA_户口邮编 = 24
    GCA_单位地址 = 25
    GCA_单位电话 = 26
    GCA_单位邮编 = 27
    GCA_联系人姓名 = 28
    GCA_联系人关系 = 29
    GCA_联系人电话 = 30
    GCA_联系人地址 = 31
    GCA_区域 = 32
    GCA_入院途径 = 33
    GCA_入院时间 = 34
    GCA_入院科室 = 35
    GCA_入院病室 = 36
    GCA_转科1 = 38
    GCA_转科2 = 39
    GCA_转科3 = 40
    GCA_出院时间 = 41
    GCA_出院科室 = 42
    GCA_出院病室 = 43
    GCA_住院天数 = 44
    '西医诊断
    GCA_入院情况 = 45
    GCA_确诊日期 = 46
    GCA_病理号 = 47
    GCA_分化程度 = 48
    GCA_最高诊断依据 = 49
    GCA_放射与病理 = 50
    GCA_门诊与出院XY = 51
    GCA_入院与出院XY = 52
    GCA_门诊与入院 = 53
    GCA_临床与病理 = 54
    GCA_临床与尸检 = 55
    GCA_术前与术后 = 56
    GCA_死亡时间 = 57
    GCA_死亡原因 = 58
    GCA_医院感染病原学诊断 = 59
    GCA_抢救次数 = 60
    GCA_成功次数 = 61
    GCA_抢救原因 = 62
    '中医诊断
    GCA_门诊与出院ZY = 63
    GCA_入院与出院ZY = 64
    GCA_辨证 = 65
    GCA_治法 = 66
    GCA_方药 = 67
    GCA_治疗类别 = 68
    GCA_抢救方法 = 69
    GCA_自制中药 = 70
    GCA_中医设备 = 71
    GCA_中医技术 = 72
    GCA_辨证施护 = 73
    '住院情况
    GCA_病例分型 = 74
    GCA_HBsAg = 75
    GCA_输血前9项检查 = 76
    GCA_血型 = 77 '基本信息 MZ
    GCA_HCVAb = 78
    GCA_发病时间 = 79 '就诊信息 MZ
    GCA_RH = 80 '基本信息 MZ
    GCA_HIVAb = 81
    GCA_生育状况 = 82 '基本信息 MZ
    GCA_输液反应 = 83
    GCA_输血反应 = 84
    GCA_输红细胞 = 85
    GCA_输血小板 = 86
    GCA_输血浆 = 87
    GCA_输全血 = 88
    GCA_输其他 = 89
    GCA_自体回收 = 90
    GCA_医学警示 = 91 '基本信息 MZ
    GCA_其他医学警示 = 92 '基本信息 MZ
    GCA_出院方式 = 93
    GCA_转入机构 = 94
    GCA_入院前天 = 95
    GCA_入院前小时 = 96
    GCA_入院前分钟 = 97
    GCA_入院后天 = 98
    GCA_入院后小时 = 99
    GCA_入院后分钟 = 100
    GCA_再入院天数 = 101
    GCA_31天目的 = 102
    GCA_呼吸机小时 = 103
    GCA_随诊期限 = 104
    GCA_门诊医师 = 105
    GCA_科主任 = 106
    GCA_主任医师 = 107
    GCA_主治医师 = 108
    GCA_住院医师 = 109
    GCA_进修医师 = 110
    GCA_研究生医师 = 111
    GCA_主诊医师 = 111
    GCA_实习医师 = 112
    GCA_质控医师 = 113
    GCA_责任护士 = 114
    GCA_质控护士 = 115
    GCA_质控日期 = 116
    GCA_病案质量 = 117
    '其他
    GCA_压疮发生期间 = 118 '附页1 YN
    GCA_压疮分期 = 119 '附页1 YN
    GCA_跌倒或坠床伤害 = 120 '附页1 YN
    GCA_跌倒或坠床原因 = 121 '附页1 YN
    GCA_重症监护 = 123 'HN
    GCA_重症监护天数 = 124 'HN
    GCA_重症监护小时 = 125 'HN
    GCA_肿瘤分期 = 126 'HN
    GCA_肿瘤分期T = 127 'HN
    GCA_肿瘤分期M = 128 'HN
    GCA_肿瘤分期N = 129 'HN
    GCA_Apgar = 130 'HN
    GCA_临床路径管理 = 131 'HN
    GCA_传染病 = 132 'HN
    GCA_DrGs管理 = 133 'HN
    GCA_引发药物 = 134 'SC
    GCA_临床表现 = 135 'SC
    GCA_离院透析尿素氮值 = 136 'SC
    '基本信息 ST HN
    GCA_其他证件 = 122
    GCA_Email = 137
    GCA_QQ = 138
    '西医诊断  ST
    GCA_感染部位 = 145
    GCA_感染与死亡 = 146
    '住院情况 SC
    GCA_输白蛋白 = 139
    GCA_退出原因 = 140 '附页1 YN
    GCA_变异原因 = 141 '附页1 YN
    GCA_院内会诊次数 = 142
    GCA_外院会诊次数 = 143
    GCA_其他会诊情况 = 144
    '附页1
    GCA_重症监护室 = 147
    GCA_重返间隔时间 = 148
    GCA_约束总时间 = 149
    GCA_约束方式 = 150
    GCA_约束工具 = 151
    GCA_约束原因 = 152
    GCA_新生儿离院方式 = 153
    '就诊信息 MZ
    GCA_门诊号 = 154
    GCA_监护人 = 155
    GCA_文化程度 = 156
    GCA_就诊摘要 = 157
    GCA_去向 = 158
    GCA_发病地址 = 159
    GCA_体温 = 160
    GCA_脉搏 = 161
    GCA_呼吸 = 162
    GCA_血压 = 163
    '基本信息 ZY
    GCA_住院号 = 164
    GCA_入院转入 = 165
    GCA_其他关系 = 166
    'SC:附页2
    GCA_距上一次住本院时间 = 167
    '住院情况
    GCA_死亡患者尸检 = 168
    GCA_监护人身份证号 = 172
End Enum

Public Enum CheckCtrlArchive
    '基本信息
    CHKA_再入院 = 0
    CHKA_入院前经外院治疗 = 1
    '西医诊断
    CHKA_是否确诊 = 2
    CHKA_医院感染作病原学检查 = 3
'    CHKA_死亡患者尸检 = 4
    CHKA_新发肿瘤 = 5
    '中医诊断
    CHKA_危重 = 6
    CHKA_急症 = 7
    CHKA_疑难 = 8
    '住院情况
    CHKA_疑难病例 = 9 '西医诊断 SC
    CHKA_示教病案 = 10
    CHKA_科研病案 = 11
    CHKA_随诊 = 12
    '其他
    CHKA_CT = 13
    CHKA_MRI = 14
    CHKA_彩色多普勒 = 15
    CHKA_细菌培养标本送检 = 16 'HN
    CHKA_单病种 = 17 'HN
    '西医诊断 SC
    CHKA_住院期间告病重或病危 = 18 '附页1 YN
    '住院情况 SC
    CHKA_进入路径 = 19 '附页1 YN
    CHKA_完成路径 = 20 '附页1 YN
    CHKA_变异 = 21 '附页1 YN
    CHKA_会诊情况 = 22
    '其他 SC
    CHKA_住院期间身体约束 = 23 'YN 住院期间使用物理约束
    '过敏与手术 YN
    CHKA_围术期死亡 = 24
    CHKA_术后猝死 = 25
    '附页1 YN
    CHKA_人工气道脱出 = 26
    CHKA_重返重症医学科 = 27
    'MZ 就诊信息
    CHKA_复诊 = 28
    CHKA_传染病上传 = 29
    'SC:附页2
    CHKA_是否因同一疾病 = 30
    '就诊信息（OM)
    CHKA_无过敏记录 = 31
End Enum

Public Function ArchivezlRefresh() As Boolean
'功能：刷新或清除医嘱清单
    On Error GoTo errH
    Call ClearPageContent
    If gclsPros.病人ID <> 0 Then
        Set gclsPros.PatiInfo = GetPatiMainInfoData(gclsPros.病人ID, gclsPros.主页ID, IIf(gclsPros.MedPageSandard = ST_门诊首页, "NULL", "")) '病案主页以及病人信息
        If gclsPros.PatiInfo.EOF Then Exit Function
        If gclsPros.MedPageSandard = ST_门诊首页 Then
            gclsPros.出院科室ID = Val(gclsPros.PatiInfo!科室id & "")
        Else
            gclsPros.出院科室ID = Val(gclsPros.PatiInfo!出院科室ID & "")
        End If
        If Not ArchiveInitEnv Then Exit Function
        Call ArchiveLoadPageData(gclsPros.病人ID, gclsPros.主页ID, IIf(gclsPros.MedPageSandard = ST_门诊首页, "NULL", ""))
    End If
    Call ArchiveSetPageHeight
    Call ArchiveFormResize
    ArchivezlRefresh = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ArchiveInitEnv() As Boolean
'功能：设置界面，并初始化列
'         blnFormLoad是否是FormLoad调用
    '表格设置
    If Not InitTableAller Then Exit Function
    If Not InitTableDiag Then Exit Function
    If gclsPros.MedPageSandard <> ST_门诊首页 Then
        If Not InitTableOPS Then Exit Function
        If Not InitTableKSS Then Exit Function
        If Not InitTableFlxAddICU Then Exit Function
        If Not InitTablefMain Then Exit Function
        If gclsPros.ReadPages Then
            If Not InitTableSpirit Then Exit Function
            If Not InitTableChemoth Then Exit Function
            If Not InitTableRadioth Then Exit Function
        End If
        If Not InitTableICUInstruments Then Exit Function
        If Not InitTableInfect Then Exit Function
        If Not InitTableSample Then Exit Function
        If Not InitTableTSJC Then Exit Function
    End If
    ArchiveInitEnv = True
End Function

Private Function ArchiveLoadPageData(ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long = 1, Optional ByVal str挂号单 As String) As Boolean
'功能：电子病案查阅数据加载
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrTmp As Variant
    On Error GoTo errH
    gblnCheck = True
    With gclsPros.CurrentForm
        Set gclsPros.AuxiInfo = GetPatiAuxiInfoData(gclsPros.病人ID, gclsPros.主页ID, str挂号单) '从表信息
        '加载病人信息
        If Not gclsPros.PatiInfo.EOF Then
            For i = 0 To gclsPros.PatiInfo.Fields.Count - 1
                 Call ArchiveSetCtrlValues(UCase(gclsPros.PatiInfo.Fields(i).Name & ""), gclsPros.PatiInfo.Fields(i).Value & "")
            Next
        End If
        '病案主页从表，或病人信息从表加载
        If Not gclsPros.AuxiInfo.EOF Then
            gclsPros.AuxiInfo.MoveFirst
            For i = 1 To gclsPros.AuxiInfo.RecordCount
                 Call ArchiveSetCtrlValues(gclsPros.AuxiInfo!信息名 & "", gclsPros.AuxiInfo!信息值 & "", gclsPros.AuxiInfo!编码 & "")
                 gclsPros.AuxiInfo.MoveNext
            Next
        End If
        '生命体征信息加载
        If gclsPros.MedPageSandard = ST_门诊首页 Then
            '血压信息需要重新组装
            If .txtInfo(GCA_血压).Tag Like "*|*" Then
                arrTmp = Split(.txtInfo(GCA_血压).Tag, "|")
                .txtInfo(GCA_血压).Text = arrTmp(0) & " / " & arrTmp(1) & " " & IIf(.lblInfo(GCA_血压).Tag = "", "mmHg", .lblInfo(GCA_血压).Tag)
            ElseIf .txtInfo(GCA_血压).Tag <> "" Then
                .txtInfo(GCA_血压).Text = " / " & .txtInfo(GCA_血压).Tag & " " & IIf(.lblInfo(GCA_血压).Tag = "", "mmHg", .lblInfo(GCA_血压).Tag)
            End If
            .txtInfo(GCA_血压).Tag = "": .lblInfo(GCA_血压).Tag = ""
             '因为66029问题部分数据存储位置发生变化需重新读取几项数据
            Set rsTmp = GetCareData(gclsPros.病人ID, gclsPros.主页ID)
            rsTmp.Filter = "信息名='身高'"
            If Not rsTmp.EOF Then .txtInfo(GCA_身高).Text = rsTmp!信息值 & " " & rsTmp!单位
            rsTmp.Filter = "信息名='体重'"
            If Not rsTmp.EOF Then .txtInfo(GCA_体重).Text = rsTmp!信息值 & " " & rsTmp!单位
            rsTmp.Filter = "信息名='体温'"
            If Not rsTmp.EOF Then .txtInfo(GCA_体温).Text = rsTmp!信息值 & " " & rsTmp!单位
            rsTmp.Filter = "信息名='收缩压'"
            If Not rsTmp.EOF Then .txtInfo(GCA_血压).Text = IIf(NVL(rsTmp!信息值) = "", "   ", NVL(rsTmp!信息值))
            rsTmp.Filter = "信息名='舒张压'"
            If Not rsTmp.EOF Then .txtInfo(GCA_血压).Text = .txtInfo(GCA_血压).Text & " / " & IIf(NVL(rsTmp!信息值) = "", "   ", NVL(rsTmp!信息值)) & " " & rsTmp!单位
            rsTmp.Filter = "信息名='呼吸'"
            If Not rsTmp.EOF Then .txtInfo(GCA_呼吸).Text = rsTmp!信息值 & IIf(rsTmp!信息值 & "" = "", "", " 次/分")
            rsTmp.Filter = "信息名='脉搏'"
            If Not rsTmp.EOF Then .txtInfo(GCA_脉搏).Text = rsTmp!信息值 & IIf(rsTmp!信息值 & "" = "", "", " 次/分")
        Else
            '留观病人无住院号
            If Val(gclsPros.PatiInfo!病人性质 & "") <> 0 Then
                .lblInfo(GCA_健康卡号).Visible = False
                .txtInfo(GCA_健康卡号).Visible = False
                .lblInfo(GCA_住院号).Visible = False
                .txtInfo(GCA_住院号).Visible = False
            End If
            '住院天数计算
            If Not IsNull(gclsPros.PatiInfo!出院日期) Then
                .txtInfo(GCA_住院天数).Text = DateDiff("d", gclsPros.PatiInfo!入院日期, gclsPros.PatiInfo!出院日期)
            Else
                .txtInfo(GCA_住院天数).Text = DateDiff("d", gclsPros.PatiInfo!入院日期, zlDatabase.Currentdate)
            End If
            If Val(.txtInfo(GCA_住院天数).Text) = 0 Then .txtInfo(GCA_住院天数).Text = "1"

            '自动提取转科科室及入出病室(房间号)
            '---------------------------------------------------------------
            If .txtInfo(GCA_转科1).Text = "" And .txtInfo(GCA_转科2).Text = "" And .txtInfo(GCA_转科3).Text = "" Then
                Set rsTmp = GetPatiTransfer(gclsPros.病人ID, gclsPros.主页ID)
                For i = 1 To rsTmp.RecordCount
                    If i = 1 Then
                        .txtInfo(GCA_转科1).Text = rsTmp!科室名称
                    ElseIf i = 2 Then
                        .txtInfo(GCA_转科2).Text = rsTmp!科室名称
                    ElseIf i = 3 Then
                        .txtInfo(GCA_转科3).Text = rsTmp!科室名称
                        Exit For
                    End If
                    rsTmp.MoveNext
                Next
            End If
            If .txtInfo(GCA_入院病室).Text = "" Or .txtInfo(GCA_出院病室).Text = "" Then
                Set rsTmp = GetPatiRoom(gclsPros.病人ID, gclsPros.主页ID)
                If .txtInfo(GCA_入院病室).Text = "" Then .txtInfo(GCA_入院病室).Text = rsTmp!入院病房 & ""
                If .txtInfo(GCA_出院病室).Text = "" Then .txtInfo(GCA_出院病室).Text = rsTmp!出院病房 & ""
            End If
        End If
        '读取诊断
        Set rsTmp = GetPatiDiagData(gclsPros.病人ID, gclsPros.主页ID, IIf(gclsPros.MedPageSandard = ST_门诊首页, 0, 1), , , gclsPros.Moved)
        If gclsPros.MedPageSandard <> ST_门诊首页 Then
            '加载病原学诊断
            Call FilterDiagByType(rsTmp, DT_病原学诊断)
            If Not rsTmp.EOF Then
                .txtInfo(GCA_医院感染病原学诊断).Text = rsTmp!诊断描述 & ""
            End If
        End If
        '加载诊断
        Call ArchiveLoadVsDiagData(.vsDiagXY, rsTmp, IIf(gclsPros.MedPageSandard <> ST_门诊首页, "1,2,3,5,6,7,10", "1"))
        If gclsPros.Have中医 Then
            Call ArchiveLoadVsDiagData(.vsDiagZY, rsTmp, IIf(gclsPros.MedPageSandard <> ST_门诊首页, "11,12,13", "11"))
        End If
        '过敏信息加载
        If gclsPros.MedPageSandard <> ST_门诊首页 Then
            Call ArchiveLoadAller(.vsAller, GetAllerData(gclsPros.病人ID, gclsPros.主页ID))
        ElseIf .chkInfo(CHKA_无过敏记录).Value = 0 Then '勾选无过敏记录，则不加载过敏记录
            Call ArchiveLoadAller(.vsAller, GetAllerData(gclsPros.病人ID, gclsPros.主页ID))
        End If
        If gclsPros.MedPageSandard <> ST_门诊首页 Then
            '加载手术信息
            Call ArchiveLoadOPS(.vsOPS, GetOPSData(gclsPros.病人ID, gclsPros.主页ID, , gclsPros.Moved))
            '诊断符合情况加载
            Set rsTmp = GetDiagMatchData(gclsPros.病人ID, gclsPros.主页ID)
            Do While Not rsTmp.EOF
                Select Case rsTmp!符合类型
                    Case 1 '门诊与出院
                        .txtInfo(GCA_门诊与出院XY).Text = decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
                    Case 2 '入院与出院
                        .txtInfo(GCA_入院与出院XY).Text = decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
                    Case 3 '放射与病理
                        .txtInfo(GCA_放射与病理).Text = decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
                    Case 4 '临床与病理
                        .txtInfo(GCA_临床与病理).Text = decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
                    Case 5 '临床与尸检
                        .txtInfo(GCA_临床与尸检).Text = decode(NVL(rsTmp!符合情况, 0), 0, "未做", 1, "符合", 2, "不符合", 3, "不肯定", "-")
                    Case 6 '术前与术后
                        .txtInfo(GCA_术前与术后).Text = decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
                    Case 7 '门诊与入院
                        .txtInfo(GCA_门诊与入院).Text = decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
                    Case 11 '中医门诊与出院
                        .txtInfo(GCA_门诊与出院ZY).Text = decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
                    Case 12 '中医入院与出院
                        .txtInfo(GCA_入院与出院ZY).Text = decode(NVL(rsTmp!符合情况, 0), 1, "符合", 2, "不符合", 3, "不肯定", "")
                    Case 13 '中医辨证
                        .txtInfo(GCA_辨证).Text = decode(NVL(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
                    Case 14 '中医治法
                        .txtInfo(GCA_治法).Text = decode(NVL(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
                    Case 15 '中医方药
                        .txtInfo(GCA_方药).Text = decode(NVL(rsTmp!符合情况, 0), 1, "准确", 2, "基本准确", 3, "重大缺陷", 4, "错误", "")
                End Select
                rsTmp.MoveNext
            Loop
            '抗菌药物
            Call ArchiveLoadKSS(.vsKSS, GetKSSData(gclsPros.病人ID, gclsPros.主页ID))
            '病案项目
            If gclsPros.ReadPages Then
                Call ArchiveLoadPageMedRec(gclsPros.病人ID, gclsPros.主页ID)
            End If
            Call ArchiveLoadOtherInfo(gclsPros.病人ID, gclsPros.主页ID)
        End If
    End With
    gblnCheck = False
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ArchiveSetPageHeight()
'功能：根据页面收缩与展开状态设置界面尺寸
'说明：Tag=1表示收缩
    Dim i As Long, intCurIdx As Integer
    With gclsPros.CurrentForm
        For i = 0 To .fraMain.UBound
            If Val(.picSize(i).Tag) = 0 Then
                .fraMain(i).Height = Val(.fraMain(i).Tag)
                Set .picSize(i).Picture = .imgSize.ListImages("-").Picture
            Else
                .fraMain(i).Height = 225
                Set .picSize(i).Picture = .imgSize.ListImages("+").Picture
            End If
        Next
        
        intCurIdx = 0
        For i = 1 To .fraMain.UBound
            If .fraMain(i).Enabled Then
                .fraMain(i).Top = .fraMain(intCurIdx).Top + .fraMain(intCurIdx).Height + 100
                intCurIdx = i
            End If
        Next
       .fraBack.Height = .fraMain(intCurIdx).Top + .fraMain(intCurIdx).Height + .fraMain(0).Top
        Call ArchiveSetScrollbar
    End With
End Sub

Public Sub ArchiveSetScrollbar()
'功能：根据当前窗体尺寸设置滚动条可见性及相关属性
    With gclsPros.CurrentForm
        If .fraBack.Width + IIf(.vsc.Visible, .vsc.Width, 0) <= .picBack.ScaleWidth Then
            .hsc.Visible = False
        Else
            .hsc.Min = 0
            .hsc.SmallChange = 5
            .hsc.LargeChange = 50
            If Not .hsc.Visible Then .hsc.Value = 0
            .hsc.Visible = True
        End If
        
        If .fraBack.Height + IIf(.hsc.Visible, .hsc.Height, 0) <= .picBack.ScaleHeight Then
            .vsc.Visible = False
        Else
            .vsc.Min = 0
            .vsc.SmallChange = 5
            .vsc.LargeChange = 50
            If Not .vsc.Visible Then .vsc.Value = 0
            .vsc.Visible = True
        End If
        .hsc.Max = (.picBack.ScaleWidth - .fraBack.Width - IIf(.vsc.Visible, .vsc.Width, 0)) / Screen.TwipsPerPixelX
        .vsc.Max = (.picBack.ScaleHeight - .fraBack.Height - IIf(.hsc.Visible, .hsc.Height, 0)) / Screen.TwipsPerPixelY
        .fraVH.Visible = .vsc.Visible And .hsc.Visible
    End With
End Sub

Public Sub ArchiveFormResize()
    On Error Resume Next
    With gclsPros.CurrentForm
        .picBack.Left = 0
        .picBack.Top = 0
        .picBack.Width = .ScaleWidth
        .picBack.Height = .ScaleHeight
        .hsc.Left = 0
        .hsc.Top = .picBack.ScaleHeight - .hsc.Height
        .hsc.Width = .picBack.ScaleWidth - IIf(.vsc.Visible, .vsc.Width, 0)
        .vsc.Top = 0
        .vsc.Left = .picBack.ScaleWidth - .vsc.Width
        .vsc.Height = .picBack.ScaleHeight - IIf(.hsc.Visible, .hsc.Height, 0)
        If .fraVH.Visible Then
            .fraVH.Left = .vsc.Left
            .fraVH.Top = .hsc.Top
            .fraVH.Refresh
        End If
        Call ArchiveSetScrollbar
    End With
End Sub

Public Sub ArchiveFormKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = gclsPros.CurrentForm.vsc.Value
    lngMin = gclsPros.CurrentForm.vsc.Min
    lngMax = gclsPros.CurrentForm.vsc.Max
    If lngMax <= lngMin Then '垂直滚动条未隐藏
        If intKeyCode = vbKeyPageDown Then '下
            If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
                gclsPros.CurrentForm.vsc.Value = lngCur + (lngMax - lngMin) / 10
            Else
                gclsPros.CurrentForm.vsc.Value = lngMax
            End If
        Else '上
            If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
                gclsPros.CurrentForm.vsc.Value = lngCur - (lngMax - lngMin) / 10
            Else
                gclsPros.CurrentForm.vsc.Value = lngMin
            End If
        End If
    End If
End Sub

Public Function ArchiveFormLoad() As Boolean
    With gclsPros.CurrentForm
        '滚动条尺寸
        .vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
        .hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
        .fraVH.Width = .vsc.Width: .fraVH.Height = .hsc.Height
        .fraBack.Left = 0: .fraBack.Top = 0
        .picBack.BackColor = .fraBack.BackColor
    End With
End Function

Public Sub ArchivepicSizeClick(ByRef intIndex As Integer)
    With gclsPros.CurrentForm
        .picSize(intIndex).Tag = IIf(Val(.picSize(intIndex).Tag) = 0, 1, 0)
        Call ArchiveSetPageHeight
        Call ArchiveFormResize
        If Not .vsc.Visible Then .fraBack.Top = 0
        If Not .hsc.Visible Then .fraBack.Left = 0
    End With
End Sub

Public Sub ArchivechkInfoClick(ByRef intIndex As Integer)
    If Not gblnCheck Then
        gblnCheck = True
        gclsPros.CurrentForm.chkInfo(intIndex).Value = IIf(gclsPros.CurrentForm.chkInfo(intIndex).Value = 1, 0, 1)
        gblnCheck = False
    End If
End Sub

Public Function ArchiveSetCtrlValues(ByVal strInfoName As String, ByVal strInfoValue As String, Optional ByVal str附加编码 As String) As Boolean
'功能：设置控件值
'参数  strInfoName=信息名
'      strInfoValue=信息值
'      str附加编码=病案附加项目编码判断
    Dim str控件名 As String
    Dim lngCount As Long, i As Long, j As Long, LngRow As Long
    Dim arrTmp As Variant, strTmp As String
    Dim vsTmp As VSFlexGrid, lstTmp As ListBox
    Dim intIndex As Integer, intIndexTmp As Integer
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Select Case strInfoName
        Case "特殊检查4", "特殊检查5", "特殊检查6"
            gclsPros.MainInfoRec.Filter = "信息名='特殊检查'"
        Case "CT", "PETCT", "双源CT", "X片", "B超", "超声心动图", "MRI", "同位素检查"
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                gclsPros.MainInfoRec.Filter = "信息名='特殊检查'"
            Else
                gclsPros.MainInfoRec.Filter = "信息名='" & strInfoName & "'"
            End If
        Case "收缩压", "舒张压", "血压单位"
            gclsPros.MainInfoRec.Filter = "信息名='血压'"
        Case Else
            gclsPros.MainInfoRec.Filter = "信息名='" & strInfoName & "'"
    End Select
    '信息未在记录集中注册，可能是数据加载扩展类型，如：病案附加项目，老抗生素记录
    If gclsPros.MainInfoRec.EOF Then
        If strInfoValue = "" Then Exit Function
        '多个抗生素名称,老版抗生素信息
        If strInfoName Like "抗生素*" And IsNumeric(Mid(strInfoName, 4)) Then
            Set vsTmp = gclsPros.CurrentForm.vsKSS
            With vsTmp
                LngRow = -1
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, KI_抗菌药物名) = "" Then LngRow = i: Exit For
                Next
                If i > .Rows - 1 Then .AddItem "": LngRow = i
                '兼容老数据，在主页从表里先读数据
                .RowData(LngRow) = GetKSSID(strInfoValue)
                If Val(.RowData(LngRow) & "") <> 0 Then
                    .TextMatrix(LngRow, KI_抗菌药物名) = strInfoValue
                    .Cell(flexcpData, LngRow, KI_抗菌药物名) = .TextMatrix(LngRow, KI_抗菌药物名)
                End If
                Call SetKSSSerial
            End With
        '附加项目，可能与界面上原本存在的病案主页从表信息存在名称冲突，
        '因此在找不到界面上的从表信息时才加载病案附加项目
        ElseIf str附加编码 <> "" Then
            Set vsTmp = gclsPros.CurrentForm.vsfMain
            With vsTmp
                For i = 0 To 3 Step 3
                    LngRow = -1: LngRow = .FindRow(strInfoName, , i)
                    If LngRow >= 0 Then
                        If .TextMatrix(LngRow, i + 2) = "是否" Then
                            .Cell(flexcpChecked, LngRow, i + 1) = IIf(Val(strInfoValue) = 0, 2, 1)
                            Exit For
                        Else
                            .TextMatrix(LngRow, i + 1) = strInfoValue
                            Exit For
                        End If
                    End If
                Next
            End With
        End If
    Else
        str控件名 = gclsPros.MainInfoRec!控件名 & ""
        With gclsPros.CurrentForm
            '根据信息扩展状态
            If gclsPros.MainInfoRec!ExpState = 0 Then
                intIndex = Val(gclsPros.MainInfoRec!Index & "")
                Select Case str控件名
                    Case "txtInfo"
                        .txtInfo(intIndex).Text = strInfoValue
                        Select Case intIndex
                            Case GCA_发病时间, GCA_出生日期
                                strInfoValue = Format(strInfoValue, IIf(Format(strInfoValue, "HH:mm") <> "00:00", "yyyy-MM-dd HH:mm", "yyyy-MM-dd"))
                            Case GCA_出院时间
                                strInfoValue = Format(strInfoValue, "yyyy-MM-dd HH:mm")
                            Case GCA_身高, GCA_体重
                                If gclsPros.MedPageSandard = ST_门诊首页 And strInfoValue <> "" Then strInfoValue = strInfoValue & " " & decode(intIndex, GCA_身高, "cm", GCA_体重, "Kg")
                            Case GCA_体温, GCA_呼吸, GCA_脉搏, GCA_新生儿体重, GCA_新生儿入院体重
                                 If strInfoValue <> "" Then strInfoValue = strInfoValue & " " & decode(intIndex, GCA_体温, "℃", GCA_呼吸, "次/分", GCA_脉搏, "次/分", GCA_新生儿体重, "克", GCA_新生儿入院体重, "克")
                            Case GCA_再入院天数
                                .lblInfo(intIndex).Caption = "出院" & IIf(Val(strInfoValue & "") = 0, 31, 7) & "天内再入院计划"
                            Case GCA_31天目的
                                .txtInfo(GCA_再入院天数).Text = IIf(strInfoValue <> "", "有", "无")
                            Case GCA_重症监护天数, GCA_重症监护小时
                                If strInfoValue <> "" Then .txtInfo(GCA_重症监护) = "是"
                            Case GCA_病例分型
                                If strInfoValue <> "" Then strInfoValue = GetNameByCode("病例分型", strInfoValue)
                            Case GCA_确诊日期 '因为确诊标志在确诊日期前加载，因此可以这样做
                                If .chkInfo(CHKA_是否确诊).Value = 0 Then
                                    strInfoValue = ""
                                Else
                                    strInfoValue = Format(strInfoValue, "yyyy-MM-dd HH:mm")
                                End If
                            Case GCA_成功次数
                                If Val(.txtInfo(GCA_抢救次数).Text) = 0 Then strInfoValue = ""
                            Case GCA_随诊期限
                                If .chkInfo(CHKA_随诊).Value = 0 Then
                                    strInfoValue = ""
                                Else
                                    strInfoValue = IIf(Val(gclsPros.PatiInfo!随诊标志 & "") = 9, "", Val(strInfoValue & "")) & _
                                                decode(Val(gclsPros.PatiInfo!随诊标志 & ""), 1, "月", 2, "年", 3, "周", 4, "天", 9, "终身")
                                End If
                            Case GCA_感染部位
                                If strInfoValue <> "" Then
                                    Set rsTmp = GetBaseCode(strInfoName)
                                    strTmp = ""
                                    If InStr(strInfoValue, "|") > 0 Then
                                        strInfoValue = Replace(strInfoValue, "|", ",") '将“|”分割符号转换为逗号
                                    End If
                                    Set rsTmp = GetBaseCode(strInfoName)
                                    For i = 1 To rsTmp.RecordCount
                                        If InStr("," & strInfoValue & ",", "," & rsTmp!编码 & ",") > 0 Then
                                            strTmp = strTmp & "," & NVL(rsTmp!名称)
                                        End If
                                        rsTmp.MoveNext
                                    Next
                                    If strTmp <> "" Then
                                        strInfoValue = Mid(strInfoValue, 2)
                                    Else
                                        strInfoValue = ""
                                    End If
                                End If
                            Case GCA_退出原因, GCA_变异原因
                                If intIndex = GCA_退出原因 Then
                                    .chkInfo(CHKA_完成路径).Value = IIf(strInfoValue = "1", 1, 0)
                                    If strInfoValue = "1" Then strInfoValue = ""
                                Else
                                    .chkInfo(CHKA_变异).Value = IIf(strInfoValue <> "", 1, 0)
                                    If strInfoValue = "1" Then strInfoValue = ""
                                End If
                            Case GCA_其他会诊情况
                                .chkInfo(CHKA_会诊情况).Value = IIf(strInfoValue <> "0", 1, 0)
                                If strInfoValue = "0" Then strInfoValue = ""
                            Case GCA_院内会诊次数, GCA_外院会诊次数
                                .chkInfo(CHKA_会诊情况).Value = 1
                            Case GCA_血压
                                Select Case strInfoName
                                    Case "收缩压"
                                        .txtInfo(intIndex).Tag = strInfoValue & "|" & .txtInfo(intIndex).Tag
                                    Case "舒张压"
                                        .txtInfo(intIndex).Tag = .txtInfo(intIndex).Tag & strInfoValue
                                    Case "血压单位"
                                        .lblInfo(intIndex).Tag = strInfoValue
                                End Select
                            Case GCA_生育状况
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 0, "未生育", 1, "生育1胎", 2, "生育2胎及以上", 4, "不详")
                            Case GCA_输血反应
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 0, "无", 1, "有", 2, "未输")
                            Case GCA_临床路径管理
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 1, "未进入", 2, "变异退出", 3, "完成")
                            Case GCA_DrGs管理
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 1, "无", 2, "按病种", 3, "按费用", 4, "两者都有")
                            Case GCA_传染病
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 1, "甲类", 2, "乙类", 3, "丙类")
                            Case GCA_肿瘤分期
                                If IsNumeric(strInfoValue) Then strInfoValue = decode(Val(strInfoValue), 1, "0期", 2, "I期", 3, "Ⅱ期", 4, "Ⅲ期", 5, "Ⅳ期", 6, "不详")
                            Case GCA_死亡患者尸检
                                strInfoValue = IIf(strInfoValue = "1", "有", " ")
                                '当出院方式为死亡时，死亡尸检无值，展示为无
                                If .txtInfo(intIndex).Tag = "1" And strInfoValue = "" Then
                                    strInfoValue = "无"
                                End If
                                If strInfoValue = "" Then .txtInfo(intIndex).Tag = "0"
                            Case GCA_出院方式
                                If strInfoValue = "死亡" And .txtInfo(GCA_死亡患者尸检).Tag = "0" Then
                                    .txtInfo(GCA_死亡患者尸检) = "无"
                                Else
                                    .txtInfo(GCA_死亡患者尸检).Tag = "1"
                                End If
                            Case GCA_身份证号
                                If zlStr.ActualLen(strInfoValue) > 12 And gclsPros.IsMaskID Then   '生成身份证号掩码
                                    strInfoValue = Mid(strInfoValue, 1, 12) & String(Len(Mid(strInfoValue, 13, 2)), "*") & Mid(strInfoValue, 15)
                                End If
                        End Select
                        .txtInfo(intIndex).Text = strInfoValue
                    Case "chkInfo"
                        .chkInfo(intIndex).Value = IIf(Val(strInfoValue) = 0, 0, 1)
                    Case "lstInfection", "lstAdvEvent"
                        If strInfoName = "感染因素" Then
                            Set lstTmp = .lstInfection
                        ElseIf strInfoName = "不良事件" Then
                            Set lstTmp = .lstAdvEvent
                        End If
                        If InStr(strInfoValue, "|") > 0 Then
                            strInfoValue = Replace(strInfoValue, "|", ",") '将“|”分割符号转换为逗号
                        End If
                        Set rsTmp = GetBaseCode(strInfoName)
                        For i = 1 To rsTmp.RecordCount
                            If InStr("," & strInfoValue & ",", "," & rsTmp!编码 & ",") > 0 Then
                                lstTmp.AddItem NVL(rsTmp!名称)
                            End If
                            rsTmp.MoveNext
                        Next
                End Select
            ElseIf gclsPros.MainInfoRec!ExpState = 1 Then
                If str控件名 <> "vsTSJC" Then
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                    gclsPros.SecdInfoRec.Sort = "Sort"
                End If
                Select Case strInfoName
                    Case "昏迷时间", "转科记录"
                        '保存格式:入院前(天，小时,分钟)|入院后(天，小时,分钟)
                        If strInfoName = "昏迷时间" Then
                            strTmp = Replace(strInfoValue, "|", ",")
                            strTmp = strTmp & ",,,,,"
                        Else
                            strTmp = strInfoValue & ",,,"
                        End If
                        arrTmp = Split(strTmp, ",")
                        For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                            .txtInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = arrTmp(i)
                            gclsPros.SecdInfoRec.MoveNext
                        Next
                    Case Else
                        If str控件名 = "vsTSJC" Then
                            If strInfoName Like "特殊检查*" And gclsPros.MedPageSandard <> ST_四川省标准 Then
                                intIndex = Val(Mid(strInfoName, 5, 1)) - 4
                            Else
                                intIndex = decode(strInfoName, "CT", TR_CT, "PETCT", TR_PETCT, "双源CT", TR_双源CT, _
                                            "X片", TR_X片, "B超", TR_B超, "超声心动图", TR_超声心动图, "MRI", TR_MRI, "同位素检查", TR_同位素检查, -1)
                                strInfoValue = decode(Val(strInfoValue), 1, "1-阳性", 2, "2-阴性", 3, "3-未做", "")
                            End If
                            If intIndex <> -1 Then
                                .vsTSJC.TextMatrix(intIndex, 1) = strInfoValue
                                .vsTSJC.Cell(flexcpData, intIndex, 1) = strInfoValue
                            End If
                        End If
                End Select
            ElseIf gclsPros.MainInfoRec!ExpState = 2 Then
            '加载时处理
            End If
        End With
    End If
    ArchiveSetCtrlValues = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ArchiveLoadVsDiagData(ByRef vsDiagInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal strDiagType As String)
'功能：将诊断加载到表格
'参数：vsDiagInput=需要加载诊断的表格
'      rsInput=读取的诊断记录集
'      strDiagType=诊断类型字符串，各类型以逗号分割
'说明：ArchiveLoadMedPageData的子函数

    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, j As Long, LngRow As Long
    Dim bln分化程度 As Boolean
    Dim bln西医 As Boolean
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String

    On Error GoTo errH
    With vsDiagInput
        bln西医 = vsDiagInput.Name = "vsDiagXY"
        '加载诊断
        arrTmp = Split(strDiagType, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            Call FilterDiagByType(rsInput, Val(arrTmp(i))) '过滤诊断
            Do While Not rsInput.EOF
                If rsInput!编码序号 = 1 Then
                    '确定当前显示行
                    LngRow = .FindRow(arrTmp(i), , DI_诊断分类, , True)
                    For j = LngRow To .Rows - 1
                        If Val(.TextMatrix(j, DI_诊断分类)) = Val(arrTmp(i)) Then
                            LngRow = j
                            If .TextMatrix(j, DI_诊断描述) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    '新增行
                    If .TextMatrix(LngRow, DI_诊断描述) <> "" Then
                        LngRow = LngRow + 1: .AddItem "", LngRow
                        .TextMatrix(LngRow, DI_诊断分类) = arrTmp(i)
                        If gclsPros.MedPageSandard = ST_门诊首页 Then .TextMatrix(LngRow, DI_诊断类型) = IIf(Val(arrTmp(i)) = DT_门诊诊断XY, "西医", "中医")
                    End If
                    
                    strTmp = rsInput!诊断描述 & ""
                    '读取诊断编码，诊断描述为(编码)描述，或(编码)描述(证候) 类型的可以获取诊断描述
                    If strTmp Like "(?*)?*" Then
                        lngPos = InStr(1, strTmp, ")")
                        .TextMatrix(LngRow, DI_诊断编码) = Mid(strTmp, 2, lngPos - 2)
                        strTmp = Mid(strTmp, lngPos + 1)
                    End If
                    If .TextMatrix(LngRow, DI_诊断编码) = "" And Not (IsNull(rsInput!诊断ID) And IsNull(rsInput!疾病id)) Then
                        '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                        .TextMatrix(LngRow, DI_诊断编码) = IIf(Not IsNull(rsInput!疾病id), rsInput!疾病id & "", rsInput!诊断ID & "")
                    End If
                    '获取中医证候，由于诊断描述可能会增加前后缀，前后缀包含括号，所以反向截取字符串
                    If strTmp Like "?*(?*)" And Not bln西医 Then
                        strTmp = StrReverse(strTmp)
                        lngPos = InStr(1, strTmp, "(")
                        .TextMatrix(LngRow, DI_中医证候) = StrReverse(Mid(strTmp, 2, lngPos - 2))
                        strTmp = StrReverse(Mid(strTmp, lngPos + 1))
                    End If
                    
                    '取诊断描述
                    .TextMatrix(LngRow, DI_诊断描述) = strTmp
                    '诊断描述的备份数据
                    If Not (IsNull(rsInput!诊断ID) And IsNull(rsInput!疾病id)) Then
                        .Cell(flexcpData, LngRow, DI_诊断描述) = IIf(Not IsNull(rsInput!疾病id), rsInput!疾病名称 & "", rsInput!诊断名称 & "")
                    Else
                        .Cell(flexcpData, LngRow, DI_诊断描述) = .TextMatrix(LngRow, DI_诊断描述)
                    End If
                    '其他列数据加载
                    .TextMatrix(LngRow, DI_是否疑诊) = IIf(Val(rsInput!是否疑诊 & "") = 1, "？", "")
                    .TextMatrix(LngRow, DI_诊断ID) = rsInput!诊断ID & ""
                    .TextMatrix(LngRow, DI_疾病ID) = rsInput!疾病id & ""
                    .TextMatrix(LngRow, DI_证候ID) = rsInput!证候ID & ""
                    '.TextMatrix(LngRow, DI_ICD附码) = rsInput!附码 & ""
                    .TextMatrix(LngRow, DI_医嘱IDs) = rsInput!医嘱ID & ""
                    .TextMatrix(LngRow, DI_诊断来源) = Val(rsInput!记录来源 & "") '保存记录来源，以便保存时，保存为首页或病案来源
                    If gclsPros.MedPageSandard <> ST_门诊首页 Then
                        .TextMatrix(LngRow, DI_备注) = rsInput!备注 & ""
                        .TextMatrix(LngRow, DI_出院情况) = rsInput!出院情况 & ""
                        .TextMatrix(LngRow, DI_入院病情) = rsInput!入院病情 & ""
                        .TextMatrix(LngRow, DI_是否未治) = IIf(Val(rsInput!是否未治 & "") = 1, "√", "")
                    Else
                        .TextMatrix(LngRow, DI_发病时间) = Format(rsInput!发病时间 & "", "YYYY-MM-DD HH:mm")
                    End If
                    .RowData(LngRow) = Val(rsInput!ID & "")
                Else
                    .TextMatrix(LngRow, DI_附码ID) = rsInput!疾病id & ""
                    .TextMatrix(LngRow, DI_ICD附码) = rsInput!疾病编码 & ""
                    .Cell(flexcpData, LngRow, DI_ICD附码) = .TextMatrix(LngRow, DI_ICD附码)
                End If
                rsInput.MoveNext
            Loop
        Next
        .Row = .FixedRows: .Col = DI_诊断描述
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'
Public Sub ArchiveLoadAller(ByVal vsAller As VSFlexGrid, ByVal rsInput As ADODB.Recordset)
'功能：电子病案查阅加载过敏药物
    Dim i As Long, LngRow As Long

    rsInput.Filter = "记录来源=3" '首页本身填写的
    If rsInput.EOF Then rsInput.Filter = "记录来源<>3" '其它来源的作为缺省显示
    With vsAller
        .Rows = rsInput.RecordCount + 1 '固定行+新行
        For i = 1 To rsInput.RecordCount
            '其它来源的可能有重复
            LngRow = -1
            If Not IsNull(rsInput!药物ID) Then
                LngRow = .FindRow(CLng(rsInput!药物ID))
            ElseIf Not IsNull(rsInput!药物名) Then
                LngRow = .FindRow(CStr(rsInput!药物名), , AI_过敏药物)
            End If
            If LngRow = -1 Then
                .RowData(i) = CLng(NVL(rsInput!药物ID, 0))
                .TextMatrix(i, AI_过敏时间) = Format(rsInput!过敏时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, AI_过敏药物) = NVL(rsInput!药物名)
                .TextMatrix(i, AI_过敏反应) = NVL(rsInput!过敏反应)
            End If
            rsInput.MoveNext
        Next
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = AI_过敏药物
    End With
End Sub

Public Sub ArchiveLoadOPS(ByVal vsOPS As VSFlexGrid, ByVal rsInput As ADODB.Recordset)
'功能：电子病案查阅加载手术情况
    Dim i As Long

    With vsOPS
        .Rows = .FixedRows: .Rows = .FixedRows + rsInput.RecordCount + 1
        For i = 1 To rsInput.RecordCount
            .TextMatrix(i, PI_手术日期) = Format(rsInput!手术开始时间 & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_结束日期) = Format(rsInput!手术结束时间 & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_手术编码) = rsInput!手术编码 & ""
            .TextMatrix(i, PI_手术名称) = rsInput!手术名称 & ""
            .TextMatrix(i, PI_主刀医师) = rsInput!主刀医师 & ""
            .TextMatrix(i, PI_助产护士) = rsInput!助产护士 & ""
            .TextMatrix(i, PI_助手1) = rsInput!第一助手 & ""
            .TextMatrix(i, PI_助手2) = rsInput!第二助手 & ""
            .TextMatrix(i, PI_麻醉方式) = rsInput!麻醉方式 & ""
            .TextMatrix(i, PI_麻醉医师) = rsInput!麻醉医师 & ""
            If rsInput!切口 & rsInput!愈合 & "" <> "" Then
                .TextMatrix(i, PI_切口愈合) = rsInput!切口 & "/" & rsInput!愈合
            End If
            .TextMatrix(i, PI_手术操作ID) = Val(rsInput!手术操作ID & "")
            .TextMatrix(i, PI_诊疗项目ID) = Val(rsInput!诊疗项目id & "")
            .TextMatrix(i, PI_麻醉ID) = Val(rsInput!麻醉ID & "")
            .TextMatrix(i, PI_麻醉类型) = rsInput!麻醉类型 & ""
            .TextMatrix(i, PI_手术情况) = rsInput!手术情况 & ""
            .TextMatrix(i, PI_ASA分级) = rsInput!asa分级 & ""
            .TextMatrix(i, PI_NNIS分级) = rsInput!NNIS分级 & ""
            .TextMatrix(i, PI_手术级别) = rsInput!手术级别 & ""
            .TextMatrix(i, PI_再次手术) = IIf(Val(rsInput!再次手术 & "") = 1, -1, 0)
            .TextMatrix(i, PI_准备天数) = Val(rsInput!准备天数 & "")
            .TextMatrix(i, PI_抗菌用药时间) = Format(rsInput!抗菌用药时间 & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_麻醉开始时间) = Format(rsInput!麻醉开始时间 & "", "yyyy-MM-dd HH:mm")
            .TextMatrix(i, PI_切口部位) = rsInput!切口部位 & ""
            .TextMatrix(i, PI_重返手术室目的) = rsInput!重返目的 & ""
            .TextMatrix(i, PI_重返手术室计划) = IIf(Val(rsInput!重返计划 & "") = 1, -1, 0)
            .TextMatrix(i, PI_切口感染) = IIf(Val(rsInput!切口感染 & "") = 1, -1, 0)
            .TextMatrix(i, PI_并发症) = IIf(Val(rsInput!并发症 & "") = 1, -1, 0)
            .Cell(flexcpData, i, PI_手术名称) = rsInput!手术原名 & ""
            .RowData(i) = Val(rsInput!ID & "")
            rsInput.MoveNext
        Next
    End With
End Sub

Public Sub ArchiveLoadKSS(ByVal vsKSS As VSFlexGrid, ByVal rsInput As ADODB.Recordset)
'功能：电子病案查阅加载手术情况
    Dim LngRow As Long, i As Long

    With vsKSS
        Do While Not rsInput.EOF
           For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, KI_抗菌药物名) = "" Or .RowData(i) = Val(rsInput!药名id & "") Then
                    LngRow = i: Exit For
                End If
            Next
            If i > .Rows - 1 Then
                .AddItem "": LngRow = i
            End If
            '装入数据
            .RowData(LngRow) = Val(rsInput!药名id & "")
            If .RowData(LngRow) <> 0 Then
                .TextMatrix(LngRow, KI_抗菌药物名) = rsInput!名称 & ""
                .Cell(flexcpData, LngRow, KI_抗菌药物名) = .TextMatrix(LngRow, KI_抗菌药物名)
                .TextMatrix(LngRow, KI_用药目的) = rsInput!用药目的 & ""
                .TextMatrix(LngRow, KI_使用阶段) = rsInput!使用阶段 & ""
                .TextMatrix(LngRow, KI_使用天数) = IIf(Val(rsInput!使用天数 & "") = 0, "", Val(rsInput!使用天数 & ""))
                .Cell(flexcpChecked, LngRow, KI_一类切口预防用) = Val(rsInput!一类切口预防用 & "")
                .TextMatrix(LngRow, KI_DDD数) = FormatEx(Val(rsInput!DDD数 & ""), 2)
                .TextMatrix(LngRow, KI_联合用药) = rsInput!联合用药 & ""
            End If
            rsInput.MoveNext
        Loop
    End With
End Sub

Public Function ArchiveLoadPageMedRec(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载放疗与化疗信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-21 15:55:27
    '问题:13999
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    Set rsTmp = GetChemothData(lng病人ID, lng主页ID)
    Set vsTmp = gclsPros.CurrentForm.vsChemoth
    With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!序号 & "")
                .TextMatrix(i, CI_化学治疗编码) = NVL(rsTmp!疾病信息)
                .TextMatrix(i, CI_开始日期) = Format(rsTmp!开始日期, "yyyy-MM-dd")
                .TextMatrix(i, CI_结束日期) = Format(rsTmp!结束日期, "yyyy-MM-dd")
                .TextMatrix(i, CI_疗程数) = Format(Val(rsTmp!疗程数 & ""), "###;-###;;")
                .TextMatrix(i, CI_总量) = Format(Val(rsTmp!总量 & ""), "###;-###;;")
                .TextMatrix(i, CI_化疗方案) = rsTmp!化疗方案 & ""
                .TextMatrix(i, CI_化疗效果) = rsTmp!化疗效果 & ""
                .TextMatrix(i, CI_疾病ID) = rsTmp!疾病id & ""
                rsTmp.MoveNext
            Next
    End With
    Set rsTmp = GetRadiothData(lng病人ID, lng主页ID)
    Set vsTmp = gclsPros.CurrentForm.vsRadioth
    With vsTmp
        .Rows = rsTmp.RecordCount + .FixedRows
        For i = 1 To rsTmp.RecordCount
            .RowData(i) = Val(rsTmp!序号 & "")
            .TextMatrix(i, RI_放射治疗编码) = NVL(rsTmp!疾病信息)
            .TextMatrix(i, RI_开始日期) = Format(rsTmp!开始日期, "yyyy-MM-dd")
            .TextMatrix(i, RI_结束日期) = Format(rsTmp!结束日期, "yyyy-MM-dd")
            .TextMatrix(i, RI_放射剂量) = Format(Val(rsTmp!放射剂量 & ""), "###;-###;;")
            .TextMatrix(i, RI_累计量) = Format(Val(rsTmp!累计量 & ""), "###;-###;;")
            .TextMatrix(i, RI_设野部位) = rsTmp!设野部位 & ""
            .TextMatrix(i, RI_放疗效果) = rsTmp!放疗效果 & ""
            .TextMatrix(i, RI_疾病ID) = rsTmp!疾病id & ""
            rsTmp.MoveNext
        Next
    End With
    If gclsPros.MedPageSandard = ST_卫生部标准 Then
        Set rsTmp = GetSpiritData(lng病人ID, lng主页ID)
        Set vsTmp = gclsPros.CurrentForm.vsSpirit
        With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!序号 & "")
                .TextMatrix(i, SI_药物名称) = rsTmp!药物名称 & ""
                .TextMatrix(i, SI_疗程) = rsTmp!疗程 & ""
                .TextMatrix(i, SI_最高日量) = rsTmp!最高日量 & ""
                .TextMatrix(i, SI_特殊反应) = rsTmp!特殊反应 & ""
                .TextMatrix(i, SI_疗效) = rsTmp!疗效 & ""
                .TextMatrix(i, SI_药品ID) = rsTmp!药品ID & ""
                rsTmp.MoveNext
            Next
        End With
    End If
    ArchiveLoadPageMedRec = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ArchiveLoadOtherInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:加载附页内容
    '参数:lng病人id-病人id
    '     lng主页id -主页id
    '返回:加载成功,返回true,否则返回False
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim vsTmp As VSFlexGrid
    Err = 0: On Error GoTo Errhand
    '重症监护情况
    If gclsPros.MedPageSandard <> ST_湖南省标准 Then
        Set rsTmp = GetICUData(lng病人ID, lng主页ID)
        If gclsPros.MedPageSandard <> ST_云南省标准 Then
            Set vsTmp = gclsPros.CurrentForm.vsFlxAddICU
            With vsTmp
                .Rows = rsTmp.RecordCount + .FixedRows
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(i, UI_监护室名称) = rsTmp!监护室名称 & ""
                    .TextMatrix(i, UI_进入时间) = rsTmp!进入时间 & ""
                    .TextMatrix(i, UI_退出时间) = rsTmp!退出时间 & ""
                    If gclsPros.MedPageSandard = ST_四川省标准 Then
                        .TextMatrix(i, UI_序号) = Val(rsTmp!序号 & "")
                        .Cell(flexcpChecked, i, UI_再入住计划) = Val(rsTmp!再入住计划 & "")
                         .TextMatrix(i, UI_再入住原因) = rsTmp!再入住原因 & ""
                    End If
                    .RowData(i) = Val(rsTmp!序号 & "")
                    rsTmp.MoveNext
                Next
            End With
        Else
            '云南版，没有表格
            rsTmp.Sort = "序号"
            If Not rsTmp.EOF Then
                For i = 0 To rsTmp.Fields.Count - 1
                    Call ArchiveSetCtrlValues(rsTmp.Fields(i).Name, rsTmp.Fields(i).Value & "")
                Next
            End If
        End If
    End If
    '器械导管情况，医院感染，标本送检
    If gclsPros.MedPageSandard = ST_四川省标准 Then
        Set rsTmp = GetICUInstrumentsData(lng病人ID, lng主页ID)
        Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
        With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, TI_ICU类型) = rsTmp!监护室名称 & ""
                .TextMatrix(i, TI_器械及导管) = rsTmp!器械及导管 & ""
                .TextMatrix(i, TI_开始时间) = rsTmp!开始使用时间 & ""
                .TextMatrix(i, TI_结束时间) = rsTmp!结束使用时间 & ""
                .TextMatrix(i, TI_感染累计小时) = rsTmp!感染累计时间 & ""
                .RowData(i) = Val(rsTmp!序号 & "")
                rsTmp.MoveNext
            Next
        End With
        
        Set rsTmp = GetInfectData(lng病人ID, lng主页ID)
        Set vsTmp = gclsPros.CurrentForm.vsInfect
        With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, FI_确诊日期) = rsTmp!确诊日期 & ""
                .TextMatrix(i, FI_感染部位) = rsTmp!感染部位 & ""
                .TextMatrix(i, FI_医院感染名称) = rsTmp!医院感染名称 & ""
                .RowData(i) = Val(rsTmp!序号 & "")
                rsTmp.MoveNext
            Next
        End With
        
        Set rsTmp = GetSampleData(lng病人ID, lng主页ID)
        Set vsTmp = gclsPros.CurrentForm.vsSample
        With vsTmp
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, MI_标本) = rsTmp!标本 & ""
                .TextMatrix(i, MI_病原学代码及名称) = rsTmp!病原学代码 & ""
                .TextMatrix(i, MI_送检日期) = rsTmp!送检日期 & ""
                .RowData(i) = Val(rsTmp!序号 & "")
                rsTmp.MoveNext
            Next
        End With
    End If
    ArchiveLoadOtherInfo = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long
'支持滚轮的滚动
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '向下滚
            zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '向上滚
            zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(gOldwinproc, hwnd, wMsg, wParam, lParam)
End Function


