Attribute VB_Name = "mdlBusiness"
Option Explicit

Public Function InitTableDiag() As Boolean
'功能：设置诊断表格列
    Dim strHeadXY As String, strHeadZY As String
    Dim strRowsXY As String, strRowsZY As String
    Dim intFixedRowsXY As Integer, intFixedRowsZY As Integer

    Dim intFixedColsXY As Integer, intFixedColsZY As Integer
    Dim vsTmp As VSFlexGrid
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    intFixedRowsXY = 1: intFixedColsXY = 1
    intFixedRowsZY = 1: intFixedColsZY = 1
    Select Case gclsPros.FuncType
        Case f诊断选择
            If gclsPros.PatiType = PF_门诊 Then
                '显示列：诊断类型,关联(诊断选择器),诊断编码,诊断描述,中医证候(中医诊断),发病时间,疑诊,增加,删除
                strHeadXY = "诊断类型,900,4;关联,450,4,11;诊断编码,900,4;诊断描述,4000,1;中医证候;发病时间,2200,1;备注;入院病情;出院情况;ICD附码;未治;疑诊,450,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strHeadZY = "诊断类型,900,4;关联,450,4,11;诊断编码,900,4;诊断描述,2900,1;中医证候,1500,1;发病时间,1800,1;备注;入院病情;出院情况;ICD附码;未治;疑诊,450,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strRowsXY = DI_诊断类型 & ",西医," & DI_诊断分类 & "," & DT_门诊诊断XY
                strRowsZY = DI_诊断类型 & ",中医," & DI_诊断分类 & "," & DT_门诊诊断ZY
            Else
                '显示列：诊断类型,关联(诊断选择器),诊断编码,诊断描述,中医证候(中医诊断),备注,入院病情;出院情况,未治,疑诊,增加,删除
                strHeadXY = "诊断类型设置宽,1450,4;关联,450,4,11;诊断编码,850,4;诊断描述,2500,1;中医证候;发病时间;备注,1000,1;入院病情,850,1;出院情况,850,1;ICD附码,700,1;未治,450,4;疑诊,450,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strHeadZY = "诊断类型设置宽,1450,4;关联,450,4,11;诊断编码,850,4;诊断描述,1900,1;中医证候,1400,1;发病时间;备注,900,1;入院病情,900,1;出院情况,900,1;ICD附码;未治;疑诊,450,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strRowsXY = DI_诊断类型 & ",门（急）诊诊断 ," & DI_诊断分类 & "," & DT_门诊诊断XY & ";" & _
                                    DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断XY & ";" & _
                                    DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                    DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                    DI_诊断类型 & ",院内感染," & DI_诊断分类 & "," & DT_院内感染 & ";" & _
                                    DI_诊断类型 & ", 并 发 症 ," & DI_诊断分类 & "," & DT_并发症 & ";" & _
                                    DI_诊断类型 & ",病理诊断," & DI_诊断分类 & "," & DT_病理诊断 & ";" & _
                                    DI_诊断类型 & ",损伤中毒," & DI_诊断分类 & "," & DT_损伤中毒码
                strRowsZY = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断ZY & ";" & _
                                    DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断ZY & ";" & _
                                    DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断ZY & ";" & _
                                    DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断ZY
            End If

        Case f医生首页
            '显示列：诊断类型,关联(诊断选择器),诊断编码,诊断描述,中医证候(中医诊断),备注,入院病情;出院情况,未治,疑诊,增加,删除
            strHeadXY = "诊断类型设置宽,1350,4;关联;诊断编码,1000,4;诊断描述,3200,1;中医证候;发病时间;备注,1000,1;入院病情,850,1;出院情况,850,1;ICD附码,800,1;未治,450,4;疑诊,450,4;" & _
                                    ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
            strHeadZY = "诊断类型设置宽,1350,4;关联;诊断编码,1000,4;诊断描述,2700,1;中医证候,1500,1;发病时间;备注,1300,1;入院病情,850,1;出院情况,850,1;ICD附码;未治;疑诊,450,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
            strRowsXY = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断XY & ";" & _
                                DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断XY & ";" & _
                                DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                DI_诊断类型 & ",院内感染," & DI_诊断分类 & "," & DT_院内感染 & ";" & _
                                DI_诊断类型 & ", 并 发 症 ," & DI_诊断分类 & "," & DT_并发症 & ";" & _
                                DI_诊断类型 & ",病理诊断," & DI_诊断分类 & "," & DT_病理诊断 & ";" & _
                                DI_诊断类型 & ",损伤中毒," & DI_诊断分类 & "," & DT_损伤中毒码
            strRowsZY = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断ZY & ";" & _
                                DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断ZY & ";" & _
                                DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断ZY & ";" & _
                                DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断ZY
        Case f病案首页
                '显示列：诊断类型,诊断编码,诊断描述,中医证候(中医诊断),备注,入院病情;出院情况,ICD附码,未治,疑诊,增加,删除
                strHeadXY = "诊断类型设置宽,1250,4;关联;诊断编码,900,4;诊断描述,3200,1;中医证候;发病时间;备注,1200,1;入院病情,850,1;出院情况,850,1;ICD附码,800,1;未治,350,4;疑诊,350,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strHeadZY = "诊断类型设置宽,1250,4;关联;诊断编码,900,4;诊断描述,3000,1;中医证候,1500,1;发病时间;备注,1100,1;入院病情,850,1;出院情况,850,1;ICD附码;未治;疑诊,350,4;" & _
                                        ",270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strRowsXY = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断XY & ";" & _
                                    DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断XY & ";" & _
                                    DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                    DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                    DI_诊断类型 & ",院内感染," & DI_诊断分类 & "," & DT_院内感染 & ";" & _
                                    DI_诊断类型 & ", 并 发 症 ," & DI_诊断分类 & "," & DT_并发症 & ";" & _
                                    DI_诊断类型 & ",病理诊断," & DI_诊断分类 & "," & DT_病理诊断 & ";" & _
                                    DI_诊断类型 & ",损伤中毒," & DI_诊断分类 & "," & DT_损伤中毒码
                strRowsZY = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断ZY & ";" & _
                                    DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断ZY & ";" & _
                                    DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断ZY & ";" & _
                                    DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断ZY
        Case f电子病案
            If gclsPros.MedPageSandard = ST_门诊首页 Then
                '显示列：诊断类型,诊断编码,诊断描述,中医证候(中医诊断),发病时间,疑诊
                strHeadXY = ",450,4;关联;诊断编码,900,4;诊断描述,3000,1;中医证候;发病时间,1500,1;备注;入院病情;出院情况;ICD附码,800,1;未治;疑诊,450,4;" & _
                                        "增加;删除;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strHeadZY = ",450,4;关联;诊断编码,900,4;诊断描述,3000,1;中医证候,1500,1;发病时间,1500,1;备注;入院病情;出院情况;ICD附码;未治;疑诊,450,4;" & _
                                        "增加;删除;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                intFixedRowsZY = 0
                strRowsXY = DI_诊断类型 & ",西医," & DI_诊断分类 & "," & DT_门诊诊断XY
                strRowsZY = DI_诊断类型 & ",中医," & DI_诊断分类 & "," & DT_门诊诊断ZY
            Else
                '显示列：诊断类型,诊断编码,诊断描述,中医证候(中医诊断),备注,入院病情;出院情况,未治,疑诊,增加,删除
                strHeadXY = "诊断类型设置宽,1350,4;关联;诊断编码,810,4;诊断描述,2700,1;中医证候;发病时间;备注,800,1;入院病情,1000,1;出院情况,810,1;ICD附码,800,1;未治,450,4;疑诊,450,4;" & _
                                        "增加;删除;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strHeadZY = "诊断类型设置宽,1350,4;关联;诊断编码,810,4;诊断描述,2500,1;中医证候,1050,1;发病时间;备注,800,1;入院病情,1000,1;出院情况,810,1;ICD附码;未治,450,4;疑诊,450,4;" & _
                                        "增加;删除;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;是否病人;疗效限制;分娩信息;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
                strRowsXY = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断XY & ";" & _
                                    DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断XY & ";" & _
                                    DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                    DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断XY & ";" & _
                                    DI_诊断类型 & ",院内感染," & DI_诊断分类 & "," & DT_院内感染 & ";" & _
                                    DI_诊断类型 & ", 并 发 症 ," & DI_诊断分类 & "," & DT_并发症 & ";" & _
                                    DI_诊断类型 & ",病理诊断," & DI_诊断分类 & "," & DT_病理诊断 & ";" & _
                                    DI_诊断类型 & ",损伤中毒," & DI_诊断分类 & "," & DT_损伤中毒码
                strRowsZY = DI_诊断类型 & ",门（急）诊诊断," & DI_诊断分类 & "," & DT_门诊诊断ZY & ";" & _
                                    DI_诊断类型 & ",入院诊断," & DI_诊断分类 & "," & DT_入院诊断ZY & ";" & _
                                    DI_诊断类型 & ",出院诊断," & DI_诊断分类 & "," & DT_出院诊断ZY & ";" & _
                                    DI_诊断类型 & ",其他诊断," & DI_诊断分类 & "," & DT_出院诊断ZY
            End If
    End Select

    Set vsTmp = gclsPros.CurrentForm.vsDiagXY
    Call Grid.Init(vsTmp, strHeadXY, strRowsXY, intFixedColsXY, intFixedRowsXY)
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then
            If Not .ColHidden(DI_入院病情) Then .ColData(DI_出院情况) = "有|临床未确定|情况不明|无"
            If Not .ColHidden(DI_出院情况) Then
                Set rsTmp = GetBaseCode("治疗结果")
                If Not rsTmp.EOF Then
                    strTmp = Rec.ToComboList(rsTmp, "[0]-[1]|", "编码", "名称")
                    '用Chr(10)代替空白项是为了实现发送空格弹出下拉列表
                    .ColData(DI_出院情况) = Chr(10) & "|" & strTmp
                Else
                    .ColData(DI_出院情况) = Chr(10) & "|1-治愈|2-好转|3-未愈|4-死亡|5-其他"
                End If
            End If
        End If
        If .Font.Size <> gclsPros.FontSize Then
            .Font.Size = gclsPros.FontSize
            Call Grid.AdjustCols(vsTmp, "," & DI_Del & "," & DI_增加 & ",")
        End If
        If .TextMatrix(0, DI_诊断类型) = "诊断类型设置宽" Then .TextMatrix(0, DI_诊断类型) = "诊断类型" '恢复列头
    End With

    Set vsTmp = gclsPros.CurrentForm.vsDiagZY
    Call Grid.Init(vsTmp, strHeadZY, strRowsZY, intFixedColsZY, intFixedRowsZY)
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then
            If Not .ColHidden(DI_入院病情) Then .ColData(DI_出院情况) = "有|临床未确定|情况不明|无"
            If Not .ColHidden(DI_出院情况) Then
                If strTmp <> "" Then
                    '用Chr(10)代替空白项是为了实现发送空格弹出下拉列表
                    .ColData(DI_出院情况) = Chr(10) & "|" & strTmp
                Else
                    .ColData(DI_出院情况) = Chr(10) & "|1-治愈|2-好转|3-未愈|4-死亡|5-其他"
                End If
            End If
        End If
          If .Font.Size <> gclsPros.FontSize Then
             .Font.Size = gclsPros.FontSize
            Call Grid.AdjustCols(vsTmp, "," & DI_Del & "," & DI_增加 & ",")
          End If
        If .TextMatrix(0, DI_诊断类型) = "诊断类型设置宽" Then .TextMatrix(0, DI_诊断类型) = "诊断类型" '恢复列头
    End With
    InitTableDiag = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableOPS() As Boolean
'功能：设置手术情况表格列
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    Select Case gclsPros.MedPageSandard
        Case ST_卫生部标准
            strHead = ",300,4;" & IIf(gclsPros.UseOPSEndTime, "手术开始时间,1850,4;手术结束时间,1850,4", "手术及操作日期,1850,4;手术结束时间") & ";术前预防性抗菌用药时间;手术情况,875,1;准备天数;手术及操作编码,1500,1;手术及操作名称,2800,1;再次手术,850,4,11;术者,850,1;助产护士,850,1;第Ⅰ助手,850,1;第Ⅱ助手,850,1;" & _
                            "麻醉开始时间;麻醉方式,850,1;ASA分级,850,1;NNIS分级,850,1;手术级别,850,1;麻醉医师,850,1;切口愈合等级,1400,1;切口部位,850,1;重返手术室计划;重返手术室目的;切口感染;并发症;" & _
                            "术前0.5-2小时预防用抗菌药;清洁手术围术期预防用抗菌药天数;非预期的二次手术;麻醉并发症;术中异物遗留;手术并发症;术后出血或血肿;手术伤口裂开;术后深静脉血栓;术后生理/代谢紊乱;术后呼吸衰竭;" & _
                            "术后肺栓塞;术后败血症;术后髋关节骨折;手术操作ID;诊疗项目ID;麻醉ID;麻醉类型;手麻来源"
        Case ST_湖南省标准
            strHead = ",300,4;" & IIf(gclsPros.UseOPSEndTime, "手术开始时间,1850,4;手术结束时间,1850,4", "手术及操作日期,1850,4;手术结束时间") & ";术前预防性抗菌用药时间;手术情况,875,1;准备天数;手术及操作编码,1500,1;手术及操作名称,2800,1;再次手术,850,4,11;术者,850,1;助产护士,850,1;第Ⅰ助手,850,1;第Ⅱ助手,850,1;" & _
                            "麻醉开始时间;麻醉方式,850,1;ASA分级,850,1;NNIS分级,850,1;手术级别,850,1;麻醉医师,850,1;切口愈合等级,1400,1;切口部位;重返手术室计划;重返手术室目的;切口感染;并发症;" & _
                            "术前0.5-2小时预防用抗菌药;清洁手术围术期预防用抗菌药天数;非预期的二次手术;麻醉并发症;术中异物遗留;手术并发症;术后出血或血肿;手术伤口裂开;术后深静脉血栓;术后生理/代谢紊乱;术后呼吸衰竭;" & _
                            "术后肺栓塞;术后败血症;术后髋关节骨折;手术操作ID;诊疗项目ID;麻醉ID;麻醉类型;手麻来源"
        Case ST_四川省标准
            strHead = ",300,4;" & "开始日期,1850,4;结束日期,1850,4;术前预防性抗菌用药时间,2150,4;手术情况,875,1;准备天数,850,7;手术编码,1500,1;手术名称,2800,1;再次手术,850,4,11;主刀医师,850,1;助产护士,850,1;第Ⅰ助手,850,1;第Ⅱ助手,850,1;" & _
                            "麻醉开始时间,1550,4;麻醉方式,850,1;ASA分级,850,1;NNIS分级,850,1;手术分级,850,1;麻醉医师,850,1;切口/愈合,1400,1;切口部位,850,1;重返手术室计划,1400,4,11;重返手术室目的,1400,1;切口感染,850,4,11;并发症,720,4,11;" & _
                            "术前0.5-2小时预防用抗菌药;清洁手术围术期预防用抗菌药天数;非预期的二次手术;麻醉并发症;术中异物遗留;手术并发症;术后出血或血肿;手术伤口裂开;术后深静脉血栓;术后生理/代谢紊乱;术后呼吸衰竭;" & _
                            "术后肺栓塞;术后败血症;术后髋关节骨折;手术操作ID;诊疗项目ID;麻醉ID;麻醉类型;手麻来源"
        Case ST_云南省标准
            strHead = ",300,4;" & "手术日期,1850,4;结束日期;术前预防性抗菌用药时间;手术情况,875,1;准备天数;手术编码,1500,1;手术名称,2800,1;再次手术,850,4,11;主刀医师,850,1;助产护士,850,1;第Ⅰ助手,850,1;第Ⅱ助手,850,1;" & _
                            "麻醉开始时间;麻醉方式,850,1;ASA分级,850,1;NNIS分级,850,1;手术分级,850,1;麻醉医师,850,1;切口/愈合,1400,1;切口部位;重返手术室计划;重返手术室目的;切口感染;并发症;" & _
                            "术前0.5-2小时预防用抗菌药,2400,4,11;清洁手术围术期预防用抗菌药天数,2850,7;非预期的二次手术,1600,4,11;麻醉并发症,1000,4,11;术中异物遗留,1200,4,11;手术并发症,1000,4,11;" & _
                            "术后出血或血肿,1450,4,11;手术伤口裂开,1200,4,11;术后深静脉血栓,1450,4,11;术后生理/代谢紊乱,1700,4,11;术后呼吸衰竭,1200,4,11;术后肺栓塞,1000,4,11;术后败血症,1000,4,11;" & _
                            "术后髋关节骨折,1450,4,11;手术操作ID;诊疗项目ID;麻醉ID;麻醉类型;手麻来源"
    End Select
    Set vsTmp = gclsPros.CurrentForm.vsOPS
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        .Font.Size = 9
        If gclsPros.FuncType <> f电子病案 Then
            .ColComboList(PI_手术情况) = " |择期|急诊|限期"
            .ColComboList(PI_ASA分级) = " |P1|P2|P3|P4|P5|P6"
            .ColComboList(PI_NNIS分级) = " |NNIS0级|NNIS1级|NNIS2级|NNIS3级"
            .ColComboList(PI_手术级别) = " |无|一级手术|二级手术|三级手术|四级手术"
            '切口愈合
            Set rsTmp = GetBaseCode("手术切口愈合")
            If Not rsTmp.EOF Then
                strTmp = " |" & Rec.ToComboList(rsTmp, "[0]-[1]|", "编码", "名称")
            Else
                strTmp = " |0-0 / |1-Ⅰ/甲|2-Ⅰ/乙|3-Ⅰ/丙|4-Ⅰ/其他|5-Ⅱ/甲|6-Ⅱ/乙|7-Ⅱ/丙|8-Ⅱ/其他|9-Ⅲ/甲|10-Ⅲ/乙|11-Ⅲ/丙|12-Ⅲ/其他|13-IV/甲|14-IV/乙|15-IV/丙|16-IV/其他"
            End If
            .ColData(PI_切口愈合) = strTmp
            '麻醉类型
            Set rsTmp = GetBaseCode("诊疗麻醉类型")
            If Not rsTmp.EOF Then
                strTmp = " |" & Rec.ToComboList(rsTmp, "[0]-[1]|", "简码", "名称")
            Else
                strTmp = " |JM-局麻|QM-全麻|CY-持硬|QT-其他|JM-静脉|BC-臂丛|JC-颈丛"
            End If
            .ColData(PI_麻醉类型) = strTmp
        End If
        If gclsPros.FontSize <> 9 Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableOPS = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableAller() As Boolean
'功能：设置过敏情况表格列
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.FuncType = f电子病案 Then
        strHead = "过敏物,4500,1;过敏反应,2000,1;过敏时间,1500,4;过敏源编码;药物ID;过敏来源 "
    Else
        strHead = "过敏物,4500,1;过敏反应,4500,1;过敏时间,1500,4;过敏源编码;药物ID;过敏来源 "
    End If
    Set vsTmp = gclsPros.CurrentForm.vsAller
    Call Grid.Init(gclsPros.CurrentForm.vsAller, strHead)

    If vsTmp.Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    InitTableAller = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableKSS() As Boolean
'功能：设置抗生素使用情况表格列
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.FuncType = f电子病案 Then
        strHead = ",300,4;抗菌药物名称,2250,1;用药目的,2000,1;使用阶段,1100,1;使用天数,850,7;Ⅰ类切口预防用,1400,4,11;DDD数,800,7;联合用药,900,1"
    Else
        strHead = ",300,4;抗菌药物名称,3000,1;用药目的,950,1;使用阶段,900,1;使用天数,950,7;Ⅰ类切口预防用,1400,4,11;DDD数,1000,7;联合用药,900,1"
    End If
    Set vsTmp = gclsPros.CurrentForm.vsKSS
    Call Grid.Init(vsTmp, strHead, , 1)
    With vsTmp
        .Font.Size = 9
        If gclsPros.FuncType <> f电子病案 Then
            .ColComboList(KI_抗菌药物名) = "..."
            .ColComboList(KI_使用阶段) = " |术前|术中|术后|围手术期"
            .ColComboList(KI_联合用药) = "Ⅰ种|Ⅱ联|Ⅲ联|Ⅳ联|>Ⅳ联"
            .ColComboList(KI_用药目的) = " |预防|治疗"
        End If
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableKSS = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTablefMain() As Boolean
'功能：设置病案附加项目表格列
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim LngCol As Long, LngRow As Long, i As Long, lngCount As Long

    On Error GoTo errH
    If gclsPros.FuncType = f电子病案 Then
         '四川版3列显示，其他版本2列显示
        If gclsPros.MedPageSandard = ST_四川省标准 Then
            strHead = "项目,1500,4;内容,1800,1;值域;项目,1500,4;内容,1800,1;值域;项目,1500,4;内容,1800,1;值域"
        Else
            strHead = "项目,1500,4;内容,1250,1;值域;项目,1500,4;内容,1250,1;值域"
        End If
    ElseIf gclsPros.FuncType = f病案首页 Then
        strHead = "项目,1620,4;内容,2210,1;值域;项目,1620,4;内容,2210,1;值域;项目,1620,4;内容,2210,1;值域"
    ElseIf gclsPros.FuncType = f医生首页 Then
        strHead = "项目,1600,4;内容,2030,1;值域;项目,1600,4;内容,2030,1;值域;项目,1600,4;内容,2030,1;值域"
    End If

    
    Set vsTmp = gclsPros.CurrentForm.vsfMain
    Call Grid.Init(vsTmp, strHead)
    strSql = "Select Rownum 序号, 名称, 内容 From (Select 名称, 内容 From 病案项目 Order By 编码)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption)
    With vsTmp
        If rsTmp.RecordCount = 0 Then
            .Rows = .FixedRows
        Else
            If gclsPros.FuncType = f电子病案 Then
                lngCount = IIf(gclsPros.MedPageSandard = ST_四川省标准, 3, 2)
            Else
                lngCount = 3
            End If
            .Rows = rsTmp.RecordCount \ lngCount + 1 + IIf(rsTmp.RecordCount Mod lngCount = 0, 0, 1)
            For i = 0 To .Cols - 1 Step 3
                .Cell(flexcpBackColor, 1, i, .Rows - 1, i) = &HFCE7D8
                .FixedAlignment(i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = flexAlignLeftCenter
            Next
            Do While Not rsTmp.EOF
                i = Val(rsTmp!序号 & "")
                LngRow = .FixedRows + ((i - 1) \ lngCount): LngCol = ((i - 1) Mod lngCount) * 3
                .TextMatrix(LngRow, LngCol) = rsTmp!名称
                .TextMatrix(LngRow, LngCol + 2) = rsTmp!内容 & ""
                If rsTmp!内容 & "" = "是否" Then
                    .TextMatrix(LngRow, LngCol + 1) = "是"
                    .Cell(flexcpChecked, LngRow, LngCol + 1) = 2
                    .Cell(flexcpAlignment, LngRow, LngCol + 1) = flexAlignCenterCenter
                Else
                    .Cell(flexcpAlignment, LngRow, LngCol + 1) = flexAlignLeftCenter
                End If
                rsTmp.MoveNext
            Loop
            If vsTmp.Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
        End If
    End With
    InitTablefMain = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableChemoth() As Boolean
'功能：设置化疗项目表格列
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    If Not gclsPros.ReadPages Then InitTableChemoth = True: Exit Function
    If gclsPros.FuncType = f电子病案 Then
        strHead = "化学治疗编码,2750,1;开始日期,1000,4;结束日期,1000,4;疗程数,900,7;化疗方案,2000,1;总量,900,7;化疗效果" & vbNewLine & "(CR PR NC PD),900,4;疾病ID"
    ElseIf gclsPros.FuncType = f病案首页 Then
        strHead = "化学治疗编码,3400,1;开始日期,1400,4;结束日期,1400,4;疗程数,700,7;化疗方案,2500,1;总量,800,7;化疗效果" & vbNewLine & "(CR PR NC PD),500,4;疾病ID"
    ElseIf gclsPros.FuncType = f医生首页 Then
        strHead = "化学治疗编码,3400,1;开始日期,1200,4;结束日期,1200,4;疗程数,700,7;化疗方案,2500,1;总量,800,7;化疗效果" & vbNewLine & "(CR PR NC PD),500,4;疾病ID"
    End If

    Set vsTmp = gclsPros.CurrentForm.vsChemoth
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then
            .ColComboList(CI_化疗效果) = "CR|PR|NC|PD"
            strTmp = zlDatabase.GetPara("化疗项目", 300, 200, "")
            If strTmp <> "" Then
                '说明:化疗项目信息，以疾病编码为准,格式为:疾病编码,缺省标志;疾病编码1,缺省标志1;...
                strSql = "Select /*+ Rule*/" & vbNewLine & _
                        " a.Id, a.编码, a.编码 || '-' || a.名称 As 疾病信息, b.缺省标志,a.序号" & vbNewLine & _
                        "From 疾病编码目录 A," & vbNewLine & _
                        "     (Select C1 编码, (Case Instr(C2, ',') When 0 Then C2 Else Substr(C2, 1, Instr(C2, ',') - 1) end) As 缺省标志," & vbNewLine & _
                        "             (Case Instr(C2, ',') When 0 Then '1' Else Substr(C2, Instr(C2, ',') +1) end) As 序号" & vbNewLine & _
                        "       From Table(f_Str2list2([1], ';', ','))) B" & vbNewLine & _
                        "Where a.编码 = b.编码 And A.序号=B.序号 And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取化疗项目信息", strTmp)
                    .ColComboList(CI_化学治疗编码) = .BuildComboList(rsTmp, "疾病信息", "ID")
                    If rsTmp.RecordCount = 1 Then
                        .ColData(CI_化学治疗编码) = NVL(rsTmp!ID) & ";" & NVL(rsTmp!疾病信息)
                    ElseIf rsTmp.RecordCount > 1 Then
                        rsTmp.Filter = "缺省标志 like '1*'"
                        If rsTmp.EOF = False Then
                            .ColData(CI_化学治疗编码) = NVL(rsTmp!ID) & ";" & NVL(rsTmp!疾病信息)
                        End If
                    Else
                        .ColData(CI_化学治疗编码) = ";"
                        gclsPros.CurrentForm.lblEdit(0).Caption = "没有可用的化疗治疗编码，请到病案系统中设置。"
                        gclsPros.CurrentForm.lblEdit(0).Visible = True
                        .Editable = flexEDNone
                    End If
            Else
                .ColData(CI_化学治疗编码) = ";"
                gclsPros.CurrentForm.lblEdit(0).Caption = "没有可用的化疗治疗编码，请到病案系统中设置。"
                gclsPros.CurrentForm.lblEdit(0).Visible = True
                .Editable = flexEDNone
            End If
        End If
    End With
    If vsTmp.Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    InitTableChemoth = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableRadioth() As Boolean
'功能：设置放疗项目表格列
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    If Not gclsPros.ReadPages Then InitTableRadioth = True: Exit Function
    If gclsPros.FuncType = f电子病案 Then
        strHead = "放射治疗编码,2750,1;开始日期,1000,4;结束日期,1000,4;设野部位,2300,1;放射剂量,900,7;累计量,1000,7;放疗效果,900,4;疾病ID"
    ElseIf gclsPros.FuncType = f病案首页 Then
        strHead = "放射治疗编码,3400,1;开始日期,1400,4;结束日期,1400,4;设野部位,2300,1;放射剂量,900,7;累计量,1000,7;放疗效果,600,4;疾病ID"
    ElseIf gclsPros.FuncType = f医生首页 Then
        strHead = "放射治疗编码,3400,1;开始日期,1200,4;结束日期,1200,4;设野部位,2300,1;放射剂量,900,7;累计量,1000,7;放疗效果,600,4;疾病ID"
    End If
    
    Set vsTmp = gclsPros.CurrentForm.vsRadioth
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then
            .ColComboList(RI_放疗效果) = "CR|PR|NC|PD"
            strTmp = zlDatabase.GetPara("放疗项目", 300, 200, "")
            If strTmp <> "" Then
                '说明:放疗项目信息，以疾病编码为准,格式为:疾病编码,缺省标志;疾病编码1,缺省标志1;...
                strSql = "Select /*+ Rule*/" & vbNewLine & _
                        " a.Id, a.编码, a.编码 || '-' || a.名称 As 疾病信息, b.缺省标志,a.序号" & vbNewLine & _
                        "From 疾病编码目录 A," & vbNewLine & _
                        "     (Select C1 编码, (Case Instr(C2, ',') When 0 Then C2 Else Substr(C2, 1, Instr(C2, ',') - 1) end) As 缺省标志," & vbNewLine & _
                        "             (Case Instr(C2, ',') When 0 Then '1' Else Substr(C2, Instr(C2, ',') +1) end) As 序号" & vbNewLine & _
                        "       From Table(f_Str2list2([1], ';', ','))) B" & vbNewLine & _
                        "Where a.编码 = b.编码 And A.序号=B.序号 And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取放疗项目信息", strTmp)
                    .ColComboList(RI_放射治疗编码) = .BuildComboList(rsTmp, "疾病信息", "ID")
                    If rsTmp.RecordCount = 1 Then
                        .ColData(RI_放射治疗编码) = NVL(rsTmp!ID) & ";" & NVL(rsTmp!疾病信息)
                    ElseIf rsTmp.RecordCount > 1 Then
                        rsTmp.Filter = "缺省标志 like '1*'"
                        If rsTmp.EOF = False Then
                            .ColData(RI_放射治疗编码) = NVL(rsTmp!ID) & ";" & NVL(rsTmp!疾病信息)
                        End If
                    Else
                        .ColData(RI_放射治疗编码) = ";"
                        gclsPros.CurrentForm.lblEdit(1).Caption = "没有可用的放疗治疗编码，请到病案系统中设置。"
                        gclsPros.CurrentForm.lblEdit(1).Visible = True
                        .Editable = flexEDNone
                    End If
            Else
                .ColData(RI_放射治疗编码) = ";"
                gclsPros.CurrentForm.lblEdit(1).Caption = "没有可用的放疗治疗编码，请到病案系统中设置。"
                gclsPros.CurrentForm.lblEdit(1).Visible = True
                .Editable = flexEDNone
            End If
        End If
    End With
    If vsTmp.Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    InitTableRadioth = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableFlxAddICU() As Boolean
'功能：设置ICU入住情况表格列
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard = ST_卫生部标准 Then
        If gclsPros.FuncType = f医生首页 Then
            strHead = "序号;重症监护室名称,7000,4;进入时间(_年_月_日 时_分_),2000,1,4,9999-99-99 99:99;退出时间(_年_月_日 时_分_),2000,1,4,9999-99-99 99:99;再入住计划;再入住原因"
        ElseIf gclsPros.FuncType = f病案首页 Then
            strHead = "序号;重症监护室名称,6500,4;进入时间(_年_月_日 时_分_),2500,1,4,9999-99-99 99:99;退出时间(_年_月_日 时_分_),2500,1,4,9999-99-99 99:99;再入住计划;再入住原因"
        Else
            strHead = "序号;重症监护室名称,3000,4;进入时间(_年_月_日 时_分_),2800,1,4,9999-99-99 99:99;退出时间(_年_月_日 时_分_),2800,1,4,9999-99-99 99:99;再入住计划;再入住原因"
        End If
    ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
        strHead = "序号,450,7;ICU类型,3100,1;入住时间,2100,4,7,9999-99-99 99:99;转出时间,2100,4,7,9999-99-99 99:99;再入住计划,1200,4,11;再入住原因,800,1"
    Else
        InitTableFlxAddICU = True: Exit Function
    End If
    Set vsTmp = gclsPros.CurrentForm.vsFlxAddICU
    Call Grid.Init(vsTmp, strHead, , IIf(gclsPros.MedPageSandard = ST_四川省标准, 1, 0))
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then
            If gclsPros.MedPageSandard = ST_卫生部标准 Then
                .ColComboList(UI_监护室名称) = "..."
            Else
                .ColComboList(UI_监护室名称) = Rec.ToComboList(GetBaseCode("ICU类型"), "[0].[1]|", "编码", "名称")
            End If
        End If
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableFlxAddICU = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableSpirit() As Boolean
'功能：设置精神药品使用情况表格列
    Dim strHead As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard <> ST_卫生部标准 Then InitTableSpirit = True: Exit Function
    If Not gclsPros.ReadPages Then InitTableSpirit = True: Exit Function
    strHead = "药物名称,2500,1;疗程,2000,1;最高日量,1500,7;特殊反应,2000,1;疗效,2000,1;药品id"
    Set vsTmp = gclsPros.CurrentForm.vsSpirit
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableSpirit = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableTSJC() As Boolean
'功能：设置精神药品使用情况表格列
    Dim strHead As String
    Dim strRows As String
    Dim intFixedRows As Integer, intFixedCols As Integer
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    
    If gclsPros.FuncType = f电子病案 Then
        strHead = ",1000,1;,2600,1"
    Else
        strHead = ",1250,1;,2600,1"
    End If
     
    If gclsPros.MedPageSandard = ST_四川省标准 Then
        strRows = "0,CT;0,PETCT;0,双源CT;0,X片;0,B超;0,超声心动图;0,MRI;0,同位素检查"
    Else
        strRows = "0,特殊检查4;0,特殊检查5;0,特殊检查6"
    End If
    Set vsTmp = gclsPros.CurrentForm.vsTSJC
    Call Grid.Init(vsTmp, strHead, strRows, 1, 0)
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                .ColComboList(1) = "1-阳性|2-阴性|3-未做"
            End If
        End If
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableTSJC = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableICUInstruments() As Boolean
'功能：设置精神药品使用情况表格列
    Dim strHead As String, strRows As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard <> ST_四川省标准 Then InitTableICUInstruments = True: Exit Function
    strHead = "ICU类型,3100,1;器械或导管类型,2400,1;开始使用时间,1600,4,7,9999-99-99 99:99;结束使用时间,1600,4,7,9999-99-99 99:99;感染累计时间(小时:分钟),1100,7,,9999:99;序号"
    Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then .ColComboList(TI_器械及导管) = Rec.ToComboList(GetBaseCode("器械导管目录"), "[0].[1]|", "编码", "名称")
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableICUInstruments = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableInfect() As Boolean
'功能：设置医院感染情况表格列
    Dim strHead As String, strRows As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard <> ST_四川省标准 Then InitTableInfect = True: Exit Function
    strHead = "确诊日期,1400,4,,9999-99-99;感染部位,1400,1;医院感染名称,1000,1;医院感染编码"
    Set vsTmp = gclsPros.CurrentForm.vsInfect
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then .ColComboList(FI_感染部位) = Rec.ToComboList(GetBaseCode("感染部位"), "[0].[1]|", "编码", "名称")
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableInfect = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableSample() As Boolean
'功能：设置标本来源表格列
    Dim strHead As String, strRows As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    If gclsPros.MedPageSandard <> ST_四川省标准 Then InitTableSample = True: Exit Function
    strHead = "标本,1400,1;病原学代码及名称,2800,1;送检日期,1200,4,7,9999-99-99"
    Set vsTmp = gclsPros.CurrentForm.vsSample
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.FuncType <> f电子病案 Then
            .ColComboList(MI_标本) = "1.血液|2.尿液|3.粪便|4.痰液|5.其他分泌物"
            .ColComboList(MI_病原学代码及名称) = Rec.ToComboList(GetBaseCode("病原学目录"), "[0]-[1]|", "编码", "名称")
        End If
        If .Font.Size <> gclsPros.FontSize Then Call zlControl.VSFSetFontSize(vsTmp, gclsPros.FontSize)
    End With
    InitTableSample = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitTableFees() As Boolean
'功能：设置费用统计表格列
    Dim strHead As String, strRows As String
    Dim vsTmp As VSFlexGrid
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim strTmp As String

    On Error GoTo errH
    If gclsPros.FuncType <> f病案首页 Then InitTableFees = True: Exit Function
    strHead = "费用名,2820,1;费用金额,1000,7;费用名,2820,1;费用金额,1000,7;费用名,2820,1;费用金额,1000,7"
    Set vsTmp = gclsPros.CurrentForm.vsFees
    Call Grid.Init(vsTmp, strHead)
    With vsTmp
        If gclsPros.OpenMode <> EM_查阅 Then
            '查询方式下不需要进行初始化
            '57638:刘鹏飞,2013-05-02,费用需要显示下级
            strSql = "Select 上级 || Decode(NVL(上级,''),'','', '_') || 编码 编码,名称 From 病案费目  START WITH 上级 IS NULL CONNECT BY PRIOR 编码 = 上级 ORDER BY 上级 || 编码"
            Call zlDatabase.OpenRecordset(rsTmp, strSql, gclsPros.CurrentForm.Caption)
            strTmp = Rec.ToComboList(rsTmp, "[0].[1]|", "编码", "名称")
            If strTmp <> "" Then
                .ColComboList(0) = strTmp
                .ColComboList(2) = strTmp
                .ColComboList(4) = strTmp
            End If
        End If
    End With
    InitTableFees = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function LoadMedPageData(ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long = 1, Optional ByVal str挂号单 As String, Optional blnNotReRead As Boolean = False, Optional ByVal bln编目 As Boolean) As Boolean
    '----------------------------------------------------------------------------------------------
    '功能:将首页数据加载到界面上
    '入参:lng病人ID=病人ID
    '     lng主页ID=病案主页ID
    '     blnNotReRead=是否是初始数据加载,Fasle=不是初始加载，True=是初始加载
    '     bln编目=是否获取编目的数据，对病案系统有效
    '出参:
    '返回:返回费用记录集
    '编制:刘硕
    '日期:2013-12-26 10:43:02
    '----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    Dim strCode As String
    Dim intMaxDiagSource As Integer
    Dim strTmp As String
    On Error GoTo errH
    Screen.MousePointer = 11
    With gclsPros.CurrentForm
        Set gclsPros.PatiInfo = GetPatiMainInfoData(lng病人ID, lng主页ID, IIf(gclsPros.PatiType = PF_门诊, gclsPros.RegistNo, "")) '病案主页以及病人信息
        '门诊病人可能（门诊首页）只传了挂号单，因此需要重新修正病人ID与主页ID参数
        If gclsPros.PatiType = PF_门诊 Then
            lng病人ID = gclsPros.病人ID
            lng主页ID = gclsPros.主页ID
        End If
        Set gclsPros.AuxiInfo = GetPatiAuxiInfoData(lng病人ID, lng主页ID, IIf(gclsPros.PatiType = PF_门诊, gclsPros.RegistNo, "")) '从表信息
        If gclsPros.FuncType = f病案首页 Then
            '初始化分娩信息记录集
            Set grsDeliceryInfo = zlDatabase.CopyNewRec(gclsPros.AuxiInfo, True, "信息名,信息值,信息值 信息现值", Array("类型", adInteger, 1, Empty, "记录性质", adInteger, 1, Empty))
            Set grsBabyDiag = zlDatabase.CopyNewRec(GetBabyDiagData(lng病人ID, lng主页ID), , , Array("记录性质", adInteger, 1, Empty))
            Set grsBabyInfo = zlDatabase.CopyNewRec(GetBabyInfoData(lng病人ID, lng主页ID), , , Array("记录性质", adInteger, 1, Empty))
        End If
        
        '加载病人信息
        If Not gclsPros.PatiInfo.EOF Then
            For i = 0 To gclsPros.PatiInfo.Fields.Count - 1
                 Call SetCtrlValues(UCase(gclsPros.PatiInfo.Fields(i).Name & ""), gclsPros.PatiInfo.Fields(i).Value & "", , True)
            Next
        End If
        
        '加载病人从表信息和病案附加项目
        If Not gclsPros.AuxiInfo.EOF Then
            gclsPros.AuxiInfo.MoveFirst
            For i = 1 To gclsPros.AuxiInfo.RecordCount
                Call SetCtrlValues(gclsPros.AuxiInfo!信息名 & "", gclsPros.AuxiInfo!信息值 & "", gclsPros.AuxiInfo!编码 & "")
                gclsPros.AuxiInfo.MoveNext
            Next
        End If
        
        '104684问题修改，入院时间取值入科时间
        strTmp = GetInDeptTime(lng病人ID, lng主页ID, "____-__-__ __:__")
        If IsDate(strTmp) Then
            .mskDateInfo(DC_入院时间).Text = strTmp
            .txtDateInfo(DC_入院时间).Text = .mskDateInfo(DC_入院时间).Text
            gclsPros.InTime = Format(strTmp, "yyyy-MM-dd hh:mm:ss")
            strTmp = ""
        End If
        
        If gclsPros.FuncType = f病案首页 Then '病案首页
            '病案首页住院号，病案号，档案号等的生成
            If gclsPros.OpenMode = EM_新增首页 Or gclsPros.OpenMode = EM_新增病案 Then
                strCode = gclsPros.PatiInfo!出院科室编码 & ""
                If strCode = "" Then strCode = gclsPros.PatiInfo!最后科室编码 & ""
                '住院号获取
                If IsNull(gclsPros.PatiInfo!住院号) Then
                    gclsPros.InNo = NVL(GetNextNo(2))
                ElseIf gclsPros.NewInNo And IsHavePageNos(CT_住院号, Not gclsPros.OpenMode = EM_编辑 Or gclsPros.Is编目, gclsPros.PatiInfo!住院号 & "", gclsPros.病人ID) Then
                    gclsPros.InNo = NVL(GetNextNo(2))
                Else
                    gclsPros.InNo = gclsPros.PatiInfo!住院号 & ""
                End If
                .txtSpecificInfo(SLC_住院号).Text = gclsPros.InNo
                '病案号获取
                If IsNull(gclsPros.PatiInfo!病案号) Then
                    If gclsPros.NewInNo Or Not gclsPros.SinPageNo And IsNull(gclsPros.PatiInfo!最后病案号) Then
                        '如果是使用新的住院号,病案号强制默认为住院号
                        '如果不存在住院病案号 , 则病案号 = 当前住院号
                        .txtInfo(GC_病案号).Text = .txtSpecificInfo(SLC_住院号).Text
                    ElseIf gclsPros.SinPageNo Then
                        .txtInfo(GC_病案号).Text = NVL(GetNextNo(4, , strCode))
                    ElseIf Not IsNull(gclsPros.PatiInfo!最后病案号) Then
                        '如果当前次数不存在住院病案号,则取最后一次整理病案的病案号
                        .txtInfo(GC_病案号).Text = gclsPros.PatiInfo!最后病案号 & ""
                    End If
                Else
                    .txtInfo(GC_病案号).Text = gclsPros.PatiInfo!病案号 & ""
                End If
            End If
            If gclsPros.OpenMode <> EM_查阅 Then
                If IsNull(gclsPros.PatiInfo!最后档案号) And gclsPros.UseFileRules Then
                    .txtInfo(GC_档案号).Text = NVL(GetNextNo(5, , strCode))
                Else
                    .txtInfo(GC_档案号).Text = gclsPros.PatiInfo!最后档案号 & ""
                End If
            End If
            If gclsPros.Is编目 Then
            '读取未编目的数据，只能从系统信息中整合出所需要的数据
                '问题26071 by lesfeng 2009-11-29 关联血库管理系统获取输血信息
                Call GetBloodValue(lng病人ID, lng主页ID)
                If gclsPros.OnLine Then
                    '功能:取护理信息:2009-02-03 14:51:23:14878
                    Call GetCareValue(lng病人ID, lng主页ID)
                    '住院转科信息
                    Call LoadTransferData(GetPatiTransfer(lng病人ID, lng主页ID))
                End If
            End If
            '加载费用信息
            Call CacheLoadVsFreesData(.vsFees, GetFreeData(lng病人ID, lng主页ID, Not gclsPros.Is编目), , Not gclsPros.Is编目)
        Else
             If gclsPros.PatiType = PF_门诊 Then '门诊首页信息设置
                '生命体征信息加载
                Call .UCPatiVitalSigns.LoadPatiVitalSigns(lng病人ID, lng主页ID)
                '病人照片加载
                Call ReadPatPricture(lng病人ID, .imgPatient, strTmp)
                gclsPros.PictureFile = strTmp
                gclsPros.CurrentForm.picPatient.Tag = strTmp
             Else '住院首页
                '临床路径相关信息获取
                 Call GetPatiPathInfo
                
                '住院转科信息
                Call LoadTransferData(GetPatiTransfer(lng病人ID, lng主页ID))
                '自动提取转科科室及入出病室(房间 号)
                If .txtInfo(GC_入院病房).Text = "" Or .txtInfo(GC_出院病房).Text = "" Then
                    Set rsTmp = GetPatiRoom(lng病人ID, lng主页ID)
                    If .txtInfo(GC_入院病房).Text = "" Then .txtInfo(GC_入院病房).Text = rsTmp!入院病房 & ""
                    If .txtInfo(GC_出院病房).Text = "" Then .txtInfo(GC_出院病房).Text = rsTmp!出院病房 & ""
                End If
             End If
        End If
        '病案首页，住院首页共有信息设置
        If gclsPros.PatiType <> PF_门诊 Then
            '多信息交互处理，多个信息相互影响的情况下信息的处理
            '有抢救次数才有成功次数,已经加载，这里需要清空
            If Val(gclsPros.PatiInfo!抢救次数 & "") = 0 Then
                .txtSpecificInfo(SLC_抢救次数).Text = ""
                .txtSpecificInfo(SLC_成功次数).Text = ""
            End If
            '随诊时，设置随诊期限
            If Val(gclsPros.PatiInfo!随诊标志 & "") <> 0 Then
                .cboSpecificInfo(SLC_随诊期限).Text = decode(Val(gclsPros.PatiInfo!随诊标志 & ""), 1, "月", 2, "年", 3, "周", 4, "天", 9, "终身", -1)
                .txtSpecificInfo(SLC_随诊期限).Text = decode(Val(gclsPros.PatiInfo!随诊标志 & ""), 0, "", 9, "", NVL(gclsPros.PatiInfo!随诊期限, 0))
                Call CboSpecificInfoClick(SLC_随诊期限)
            End If
        End If
        '将住院科室ID,出院科室ID,合同单位ID保存在界面控件上,保存时可能会用到
        .txtAdressInfo(ADRC_单位地址).Tag = gclsPros.PatiInfo!合同单位id & ""
        If gclsPros.PatiType = PF_住院 Then
            .txtInfo(GC_出院科室).Tag = gclsPros.PatiInfo!出院科室ID & ""
            .txtInfo(GC_入院科室).Tag = gclsPros.PatiInfo!入院科室ID & ""
        End If
         '过敏信息加载
         If gclsPros.MedPageSandard <> ST_门诊首页 Then
            If .chkInfo(CHK_无过敏记录).Value = 0 Then
                Set rsTmp = GetAllerData(lng病人ID, lng主页ID)
                Call CacheLoadVsAllerData(.vsAller, rsTmp)
            End If
        ElseIf .chkInfo(CHK_无过敏记录).Value = 0 Then '勾选无过敏记录，则不加载过敏记录
            Set rsTmp = GetAllerData(lng病人ID, lng主页ID)
            Call CacheLoadVsAllerData(.vsAller, rsTmp)
        End If
        '读取诊断
        Set rsTmp = GetPatiDiagData(lng病人ID, lng主页ID, IIf(gclsPros.PatiType <> PF_门诊, 1, 0), , Not gclsPros.Is编目, gclsPros.Moved)
        rsTmp.Filter = "记录来源=" & IIf(gclsPros.FuncType = f病案首页, 4, 3)
        intMaxDiagSource = IIf(gclsPros.FuncType = f病案首页, 4, -1)
        If gclsPros.FuncType = f病案首页 And rsTmp.EOF Then
            intMaxDiagSource = 3
            rsTmp.Filter = "记录来源=3"
            If rsTmp.EOF Then intMaxDiagSource = 2
        End If
        If Not gclsPros.Is复诊 Or gclsPros.Is复诊 And rsTmp.RecordCount = 0 Then
            '解决修改多次入院病人病案的西医诊断时，出现诊断错乱的问题
            gclsPros.MainInfoRec.Filter = "信息名='西医诊断' or 信息名='中医诊断'"
            gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
            If gclsPros.SecdInfoRec.RecordCount > 0 Then
                gclsPros.SecdInfoRec.MoveFirst
                For i = 1 To gclsPros.SecdInfoRec.RecordCount
                    gclsPros.SecdInfoRec.Delete
                    gclsPros.SecdInfoRec.MoveNext
                Next
            End If
            '2、加载西医诊断
            '   1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
            Call CacheLoadVsDiagData(.vsDiagXY, rsTmp, IIf(gclsPros.PatiType <> PF_门诊, "1,2,3,5,6,7,10", "1"), , intMaxDiagSource)
            '3、加载中医诊断
            '   11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断(主要诊断、其它诊断)
            If gclsPros.Have中医 Then
                Call CacheLoadVsDiagData(.vsDiagZY, rsTmp, IIf(gclsPros.PatiType <> PF_门诊, "11,12,13", "11"), , intMaxDiagSource)
            End If
        End If
        '住院首页病案首页表格加载
        If gclsPros.PatiType <> PF_门诊 Then
            '加载病原学诊断
            Call FilterDiagByType(rsTmp, DT_病原学诊断, intMaxDiagSource)
            If Not rsTmp.EOF Then
                .txtInfo(GC_病原学诊断).Text = rsTmp!诊断描述 & ""
                .cmdInfo(GC_病原学诊断).Tag = Val(rsTmp!疾病id & "")
                Call UpdateCacheRecInfo(0, "病原学诊断", rsTmp!诊断描述 & rsTmp!疾病id & "", , , Val(rsTmp!记录来源 & ""))
            End If

            Set rsTmp = GetOPSData(lng病人ID, lng主页ID, Not gclsPros.Is编目, gclsPros.Moved)
            rsTmp.Filter = "记录来源=" & IIf(gclsPros.FuncType = f病案首页, 4, 3)
            If gclsPros.FuncType = f病案首页 And rsTmp.EOF Then
                rsTmp.Filter = "记录来源=3"
                If rsTmp.EOF Then
                    rsTmp.Filter = "记录来源=1"
                End If
            End If
            '手术加载
            Call CacheLoadVsOPSData(.vsOPS, rsTmp)
            '诊断符合情况加载（诊断符合情况与手术，诊断，尸检标志有关，因此放在这里）
            Call CacheLoadDiagMatchData(GetDiagMatchData(lng病人ID, lng主页ID))
            '抗菌药使用情况加载(从表信息里也存在抗菌药情况的数据（老数据），这里与从表信息综合加载)
            Call CacheLoadVsKSSData(.vsKSS, GetKSSData(lng病人ID, lng主页ID))
            '重症监护使用情况加载
            If gclsPros.MedPageSandard = ST_云南省标准 Then
                Call CacheLoadVsFlxAddICUData(, GetICUData(lng病人ID, lng主页ID))
            ElseIf gclsPros.MedPageSandard = ST_卫生部标准 Or gclsPros.MedPageSandard = ST_四川省标准 Then
                Call CacheLoadVsFlxAddICUData(.vsFlxAddICU, GetICUData(lng病人ID, lng主页ID))
            End If
            '重症监护器械使用、医院感染、标本情况
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                Call CacheLoadVsICUInstrumentsData(.vsICUInstruments, GetICUInstrumentsData(lng病人ID, lng主页ID))
                Call CacheLoadvsInfectData(.vsInfect, GetInfectData(lng病人ID, lng主页ID))
                Call CacheLoadvsSampleData(.vsSample, GetSampleData(lng病人ID, lng主页ID))
            End If
            '放疗、化疗、精神药品加载(这三种数据在病案与标准版系统共享时才加载)
            If gclsPros.ReadPages Then
                Call CacheLoadVsChemothData(.vsChemoth, GetChemothData(lng病人ID, lng主页ID))
                Call CacheLoadVsRadiothData(.vsRadioth, GetRadiothData(lng病人ID, lng主页ID))
                If gclsPros.MedPageSandard = ST_卫生部标准 Then
                    Call CacheLoadVsSpiritData(.vsSpirit, GetSpiritData(lng病人ID, lng主页ID))
                End If
            End If
            Call GetDaysFromLast
        End If
        Call SetAllVSF
        
         
    '检查是否开启外挂部件加载首页
    Call CreatePlugInOK(gclsPros.Module)
    If Not gobjPlugIn Is Nothing Then
        Err.Clear: On Error Resume Next
        If gobjPlugIn.gblnLoadMec = True Then
            '调用病案自定义加载接口
            If Err.Number = 0 Then
                Set gColCtl = CtlAdd
                Call gobjPlugIn.LoadMecInfo(gclsPros.SysNo, gclsPros.Module, lng病人ID, lng主页ID, gclsPros.PatiType, gColCtl)
            End If
            Call zlPlugInErrH(Err, "LoadMecInfo")
            Err.Clear: On Error GoTo 0
        End If
    End If
    

    '调用病案首页外挂部件加载自定义附页数据
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            Err.Clear: On Error Resume Next
            Call gfrmMecCol(i).LoadPlugMec(gclsPros.SysNo, gclsPros.Module, lng病人ID, lng主页ID, gclsPros.PatiType)
            Call zlPlugInErrH(Err, "LoadPlugMec")
            Err.Clear: On Error GoTo 0
        Next
    End If
    
    End With
    Screen.MousePointer = 0
    LoadMedPageData = True
    Exit Function
errH:
    Debug.Print "LoadMedPageData:" & Err.Source & "===" & Err.Description
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'初始化界面控件数据
Public Function InitMedRecEnv(Optional ByVal blnAfterLoadData As Boolean, Optional ByVal blnReLoad As Boolean) As Boolean
'功能：初始化首页编辑时所需要的一些数据
'参数：blnAfterLoadData=是否在数据加载之后初始化，True-在数据加载之后初始化，False-在数据加载之前初始化
'      blnReLoad=是否是重新初始化
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strSql As String, i As Long
    Dim vsTmp As VSFlexGrid, LngCol As Long, LngRow As Long
    Dim objTextBox As TextBox, objCmd As CommandButton, objPadr As PatiAddress
    Dim bln参数设置 As Boolean
    Dim datCur As Date

    On Error GoTo errH
    If Not blnAfterLoadData Then
        With gclsPros.CurrentForm
            Screen.MousePointer = 11
            '设置部分下拉框的高度的宽度
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_职业).hwnd, .cboBaseInfo(BCC_职业).Width + 2800)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_国籍).hwnd, .cboBaseInfo(BCC_国籍).Width + 1600)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_民族).hwnd, .cboBaseInfo(BCC_民族).Width + 800)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_关系).hwnd, .cboBaseInfo(BCC_关系).Width + 1000)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_分化程度).hwnd, .cboBaseInfo(BCC_分化程度).Width + 500)
            Call zlControl.CboSetWidth(.cboBaseInfo(BCC_最高诊断依据).hwnd, .cboBaseInfo(BCC_最高诊断依据).Width + 1200)
            
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_民族), .cboBaseInfo(BCC_民族).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_国籍), .cboBaseInfo(BCC_国籍).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_婚姻), .cboBaseInfo(BCC_婚姻).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_职业), .cboBaseInfo(BCC_职业).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_付款方式), .cboBaseInfo(BCC_付款方式).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_关系), .cboBaseInfo(BCC_关系).Height * 16)
            Call zlControl.CboSetHeight(.cboBaseInfo(BCC_分化程度), .cboBaseInfo(BCC_分化程度).Height * 16)
            
            If gclsPros.PatiType <> PF_门诊 Then
                Call zlControl.CboSetHeight(.cboManInfo(MC_门诊医师), .cboManInfo(MC_门诊医师).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_门诊医师).hwnd, .cboManInfo(MC_门诊医师).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_科主任), .cboManInfo(MC_科主任).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_科主任).hwnd, .cboManInfo(MC_科主任).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_主任或副主任), .cboManInfo(MC_主任或副主任).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_主任或副主任).hwnd, .cboManInfo(MC_主任或副主任).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_主治医师), .cboManInfo(MC_主治医师).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_主治医师).hwnd, .cboManInfo(MC_主治医师).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_住院医师), .cboManInfo(MC_住院医师).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_住院医师).hwnd, .cboManInfo(MC_住院医师).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_进修医师), .cboManInfo(MC_进修医师).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_进修医师).hwnd, .cboManInfo(MC_进修医师).Width + 1600)
                If gclsPros.MedPageSandard = ST_四川省标准 Then
                    Call zlControl.CboSetHeight(.cboManInfo(MC_主诊医师), .cboManInfo(MC_主诊医师).Height * 16)
                    Call zlControl.CboSetWidth(.cboManInfo(MC_主诊医师).hwnd, .cboManInfo(MC_主诊医师).Width + 1600)
                Else
                    Call zlControl.CboSetHeight(.cboManInfo(MC_研究生医师), .cboManInfo(MC_研究生医师).Height * 16)
                    Call zlControl.CboSetWidth(.cboManInfo(MC_研究生医师).hwnd, .cboManInfo(MC_研究生医师).Width + 1600)
                End If
                
                If gclsPros.FuncType = f病案首页 Then
                    Call zlControl.CboSetHeight(.cboManInfo(MC_编目员), .cboManInfo(MC_编目员).Height * 16)
                    Call zlControl.CboSetWidth(.cboManInfo(MC_编目员).hwnd, .cboManInfo(MC_编目员).Width + 1600)
                End If
                Call zlControl.CboSetHeight(.cboManInfo(MC_实习医师), .cboManInfo(MC_实习医师).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_实习医师).hwnd, .cboManInfo(MC_实习医师).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_质控医师), .cboManInfo(MC_质控医师).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_质控医师).hwnd, .cboManInfo(MC_质控医师).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_质控护士), .cboManInfo(MC_质控护士).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_质控护士).hwnd, .cboManInfo(MC_质控护士).Width + 1600)
                Call zlControl.CboSetHeight(.cboManInfo(MC_责任护士), .cboManInfo(MC_责任护士).Height * 16)
                Call zlControl.CboSetWidth(.cboManInfo(MC_责任护士).hwnd, .cboManInfo(MC_责任护士).Width + 1600)
            End If

'            部分固定内容的下拉框
            Call SetCboFromList(Array("", "0-未生育", "1-生育1胎", "2-生育2胎及以上", "4-不详"), Array(.cboBaseInfo(BCC_生育状况)), 0)
            Call SetCboFromList(Array("-", "0-未查", "1-阴", "2-阳", "3-不详"), Array(.cboBaseInfo(BCC_RH)))
            Call SetCboFromList(Array("岁", "月", "天", "小时", "分钟"), Array(.cboSpecificInfo(SLC_年龄)), 0) '添加项目时请注意cboInfo(cbo年龄单位).listIndex<3的判断
            If gclsPros.PatiType <> PF_门诊 Then
                Call SetCboFromList(Array("天", "周", "月", "年", "终身"), Array(.cboSpecificInfo(SLC_随诊期限)), 0)
                Call SetCboFromList(Array("1.1-中", "1.2-民族", "2-中西", "3-西"), Array(.cboBaseInfo(BCC_治疗类别), .cboBaseInfo(BCC_抢救方法)))
                Call SetCboFromList(Array("0-未知", "1-有", "2-无"), Array(.cboBaseInfo(BCC_自制中药制剂)))
                Call SetCboFromList(Array(" ", "1-是", "2-否"), Array(.cboBaseInfo(BCC_中医诊疗设备), .cboBaseInfo(BCC_中医诊疗技术), .cboBaseInfo(BCC_辨证施护)))
                Call SetCboFromList(Array("0-未做", "1-准确", "2-基本准确", "3-重大缺陷", "4-错误"), Array(.cboBaseInfo(BCC_辩证), .cboBaseInfo(BCC_治法), .cboBaseInfo(BCC_方药)))
                Call SetCboFromList(Array("1-有", "2-无", "3-未输", "4-不确定"), Array(.cboBaseInfo(BCC_输液反应)))
                Call SetCboFromList(Array("0-无", "1-有", "2-未输", "3-不确定"), Array(.cboBaseInfo(BCC_输血反应)))
                Call SetCboFromList(Array("-"), Array(.cboBaseInfo(BCC_死亡患者尸检)), 0)
                Call SetCboFromList(Array("-"), Array(.cboBaseInfo(BCC_临床与尸检)), 0)
                Call SetCboFromList(Array("1-是", "2-否", "3-部分"), Array(.cboBaseInfo(BCC_输血前9项检查)))
                Call SetCboFromList(Array("0-未做", "1-符合", "2-不符合", "3-不肯定"), Array(.cboBaseInfo(BCC_门诊与出院XY), .cboBaseInfo(BCC_门诊与入院), .cboBaseInfo(BCC_入院与出院XY), .cboBaseInfo(BCC_放射与病理), .cboBaseInfo(BCC_临床与病理), _
                                                                                            .cboBaseInfo(BCC_术前与术后), .cboBaseInfo(BCC_门诊与出院ZY), .cboBaseInfo(BCC_入院与出院ZY)))
                Call SetCboFromList(Array(" ", "0-院外", "1-住院期间"), Array(.cboBaseInfo(BCC_压疮发生期间)))
                Call SetCboFromList(Array(" ", "1期", "2期", "3期", "4期", "5期", "6期"), Array(.cboBaseInfo(BCC_压疮分期)))
                Call SetCboFromList(Array(" ", "一级", "二级", "三级", "未造成伤害"), Array(.cboBaseInfo(BCC_跌倒或坠床伤害)))
                Call SetCboFromList(Array(" ", "健康原因", "治疗、药物、麻醉原因", "环境因素", "其他原因"), Array(.cboBaseInfo(BCC_跌倒或坠床原因)))
                Call SetCboFromList(Array("月", "天", "小时", "分钟"), Array(.cboSpecificInfo(SLC_婴幼儿年龄)), 0) '添加项目时请注意cboInfo(cbo年龄单位).listIndex<3的判断
                Call SetCboFromList(Array("31天内再住院计划", "7天内再住院计划"), Array(.cboBaseInfo(BCC_再入院计划天数)), 0)
                Call SetCboFromList(Array("", "1-甲", "2-乙", "3-丙"), Array(.cboBaseInfo(BCC_病案质量)))
                Call SetCboFromList(Array("", "0-直接", "1-间接", "2-无"), Array(.cboBaseInfo(BCC_感染与死亡关系)), 0)

                If gclsPros.MedPageSandard <> ST_四川省标准 Then
                    Call SetCboFromList(Array("0-未做", "1-阴性", "2-阳性", "3-弱阳性"), Array(.cboBaseInfo(BCC_HBsAg)))
                    Call SetCboFromList(Array("0-未做", "1-阴性", "2-阳性", "3-不确定"), Array(.cboBaseInfo(BCC_HCVAb), .cboBaseInfo(BCC_HIVAb)))
                End If

                If gclsPros.MedPageSandard = ST_云南省标准 Then
                    If gclsPros.FuncType = f病案首页 Then
                        Call SetCboFromList(Array("第一次住本院", "当天", "2-15天", "16-31天", "＞31天"), Array(.cboBaseInfo(BCC_距上次住院时间)))
                    End If
                    Call SetCboFromList(Array("非重返", "24h内", "24-48h", "＞48h"), Array(.cboBaseInfo(BCC_重返间隔时间)))
                    Call SetCboFromList(Array("", "一处", "两处", "三处", "其他"), Array(.cboBaseInfo(BCC_约束方式)))
                    Call SetCboFromList(Array("", "软式管", "硬式管", "背心", "老人椅", "约束带", "其他"), Array(.cboBaseInfo(BCC_约束工具)))
                    Call SetCboFromList(Array("", "认知障碍", "可能跌倒", "行为紊乱", "治疗需要", "躁动", "医疗限制", "其他"), Array(.cboBaseInfo(BCC_约束原因)))
                    Call SetCboFromList(Array("", "医嘱出院", "转儿科", "转院", "非医嘱出院", "死亡"), Array(.cboBaseInfo(BCC_新生儿离院方式)))
                ElseIf gclsPros.MedPageSandard = ST_湖南省标准 Then
                    Call SetCboFromList(Array("", "1-未进入", "2-变异退出", "3-完成"), Array(.cboBaseInfo(BCC_临床路径管理)), 0)
                    Call SetCboFromList(Array("", "1-无", "2-按病种", "3-按费用", "4-两者都有"), Array(.cboBaseInfo(BCC_实施DGRS管理)), 0)
                    Call SetCboFromList(Array("", "1-甲类", "2-乙类", "3-丙类"), Array(.cboBaseInfo(BCC_法定传染病)), 0)
                    Call SetCboFromList(Array("", "1-0期", "2-I期", "3-Ⅱ期", "4-Ⅲ期", "5-Ⅳ期", "6-不详"), Array(.cboBaseInfo(BCC_肿瘤分期)), 0)
                ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                    Call SetCboFromList(Array("", "1-0期", "2-I期", "3-Ⅱ期", "4-Ⅲ期", "5-Ⅳ期", "6-不详"), Array(.cboBaseInfo(BCC_肿瘤分期)), 0)
                End If
            End If
            '根据一些字典设置下拉框内容
            Call SetCboFromRec(Array(BCC_付款方式, BCC_性别, BCC_婚姻, BCC_职业, BCC_民族, BCC_国籍, BCC_血型, BCC_身份证), 0)
            If gclsPros.PatiType <> PF_门诊 Then
                Call SetCboFromRec(Array(BCC_病例分型), 0, "")
                Call SetCboFromRec(Array(BCC_关系, BCC_入院情况, BCC_入院途径, BCC_分化程度, BCC_最高诊断依据, BCC_出院方式), 0)
                '感染部位，感染因素，不良事件等ListBOX的加载
                Call SetLstBoxFromRec(IIf(gclsPros.MedPageSandard = ST_四川省标准, "感染部位,不良事件", "感染部位,感染因素,不良事件"))
            Else
                Call SetCboFromRec(Array(BCC_去向), 0, " ")
                Call SetCboFromRec(Array(BCC_文化程度), 0)
            End If
            Call SetCboFromRec(Array(BCC_死亡期间), 0)
            If gclsPros.FuncType = f病案首页 Then
                '得到默认出生地
                strSql = "select A.编码,A.名称 from 地区 a where a.缺省标志=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption)
                If rsTmp.RecordCount > 0 Then
                    Call SetPatiAddress(ADRC_出生地点, "出生地点", rsTmp!名称, True)
                    If gclsPros.DefautADD Then
                        Call SetPatiAddress(ADRC_联系人地址, "联系人地址", rsTmp!名称, True)
                        Call SetPatiAddress(ADRC_现住址, "家庭地址", rsTmp!名称, True)
                        .txtSpecificInfo(SLC_家庭邮编).Text = rsTmp!编码 & ""
                    End If
                End If
                '问题:13557
                strSql = "select A.编码,A.名称 from 区域 a where a.缺省标志=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption)
                If rsTmp.RecordCount > 0 Then
                    Call SetPatiAddress(ADRC_病人区域, "区域", rsTmp!名称, True)
                End If
                datCur = zlDatabase.Currentdate
                '设置默认值
                .mskDateInfo(DC_编目日期).Text = Format(datCur, GetFormat(.mskDateInfo(DC_编目日期).Tag))
                .mskDateInfo(DC_收回日期).Text = Format(datCur, GetFormat(.mskDateInfo(DC_收回日期).Tag))
                .mskDateInfo(DC_出生日期).Text = Format(datCur, GetFormat(.mskDateInfo(DC_出生日期).Tag))
                .mskDateInfo(DC_入院时间).Text = Format(datCur, GetFormat(.mskDateInfo(DC_入院时间).Tag))
                .mskDateInfo(DC_出院时间).Text = Format(datCur, GetFormat(.mskDateInfo(DC_出院时间).Tag))
                .mskDateInfo(DC_质控日期).Text = Format(datCur, GetFormat(.mskDateInfo(DC_质控日期).Tag))

                .txtDateInfo(DC_编目日期).Text = .mskDateInfo(DC_编目日期).Text
                .txtDateInfo(DC_收回日期).Text = .mskDateInfo(DC_收回日期).Text
                .txtDateInfo(DC_出生日期).Text = .mskDateInfo(DC_出生日期).Text
                .txtDateInfo(DC_入院时间).Text = .mskDateInfo(DC_入院时间).Text
                .txtDateInfo(DC_出院时间).Text = .mskDateInfo(DC_出院时间).Text
                .txtDateInfo(DC_质控日期).Text = .mskDateInfo(DC_质控日期).Text

                gclsPros.InTime = .mskDateInfo(DC_入院时间).Text
                gclsPros.OutTime = .mskDateInfo(DC_出院时间).Text
                '病案界面设置
                .cmdFeeEdit.Visible = Not gclsPros.OnLine And (gclsPros.OpenMode = EM_新增病案 Or gclsPros.OpenMode = EM_新增首页) And gclsPros.OutFile <> ""
            End If
            If gclsPros.PatiType <> PF_门诊 Then
                '根据相关参数以及变量，初始化界面
                '结构化地址
                On Error Resume Next
                For Each objTextBox In .txtAdressInfo
                    Set objPadr = .padrInfo(objTextBox.Index)
                    strTmp = objPadr.Name
                    If Err.Number = 0 Then  '存在地址控件，则TextBox与CommandButton（可能不存在）需要隐藏
                        objTextBox.Visible = Not gclsPros.IsStructAdress
                        Set objCmd = .cmdAdressInfo(objTextBox.Index) '可能不存在，因为错误忽略，所以直接设置属性
                        objCmd.Visible = Not gclsPros.IsStructAdress
                        objPadr.Visible = gclsPros.IsStructAdress
                        objPadr.ShowTown = gclsPros.IsShowTown
                    Else
                        Err.Clear
                    End If
                Next
                If gclsPros.MedPageSandard = ST_四川省标准 Or gclsPros.MedPageSandard = ST_云南省标准 Then
                    Call SetCboFromRec(Array(BCC_变异原因), 0)
                    .cboBaseInfo(BCC_变异原因).Visible = gclsPros.PathVCauses
                    .fraCbo(0).Visible = gclsPros.PathVCauses
                    .txtInfo(GC_变异原因).Visible = Not gclsPros.PathVCauses
                End If

                On Error GoTo errH
                '四川版获取上次诊断、病案云南获取上次诊断
                If gclsPros.MedPageSandard = ST_四川省标准 Or gclsPros.MedPageSandard = ST_云南省标准 And gclsPros.FuncType = f病案首页 Then
                    .cmdLastDiag.Visible = gclsPros.主页ID > 1
                End If
            End If
        End With
        If Not InitTableDiag Then Exit Function
        If Not InitTableAller Then Exit Function
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
            If Not InitTableFees Then Exit Function
        End If
    Else
        '数据加载后界面的设置
        '中西医诊断加载后界面设置
        Set vsTmp = gclsPros.CurrentForm.vsDiagXY
        With vsTmp
            .Cell(flexcpForeColor, 1, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
            .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
            If gclsPros.PatiType <> PF_门诊 Then
                LngRow = FindDiagRow(DT_出院诊断XY)
                .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
                .Row = .FixedRows: .Col = DI_诊断描述
                Call DiagAfterRowColChange(vsTmp, -1, -1, .Row, .Col)
            Else
                .Cell(flexcpText, .FixedRows, DI_诊断类型, .Rows - 1, DI_诊断类型) = "西医"
            End If
        End With

        Set vsTmp = gclsPros.CurrentForm.vsDiagZY
        With vsTmp
            .Cell(flexcpForeColor, .FixedRows, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
            .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
            If gclsPros.PatiType <> PF_门诊 Then
                LngRow = FindDiagRow(DT_出院诊断ZY)
                .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
                Call DiagAfterRowColChange(vsTmp, -1, -1, .Row, .Col)
            Else
                .Cell(flexcpText, .FixedRows, DI_诊断类型, .Rows - 1, DI_诊断类型) = "中医"
            End If
        End With
        '数据加载后设置抗菌药表格的序号
        If gclsPros.PatiType <> PF_门诊 Then Call SetKSSSerial
        With gclsPros.CurrentForm
            If gclsPros.FuncType = f医生首页 And gclsPros.PatiType <> PF_门诊 Then
                '留观病人无住院号
                If Val(gclsPros.PatiInfo!病人性质 & "") <> 0 Then
                    .lblSpecificInfo(SLC_住院号).Visible = False
                    .txtSpecificInfo(SLC_住院号).Visible = False
                    .txtSpecificInfo(SLC_住院号).Enabled = False '标志为不检查
                    .PicInNum.Visible = False
                End If
            End If
            '门诊首页仅有西医诊断时隐藏中医诊断表格
            If Not gclsPros.Have中医 And gclsPros.PatiType = PF_门诊 Then
                .vsDiagZY.Visible = False
                .vsDiagXY.Height = .vsDiagZY.Top + .vsDiagZY.Height - .vsDiagXY.Top
               .vsDiagXY.ColHidden(DI_诊断分类) = True
                .vsDiagXY.ColWidth(DI_诊断编码) = .vsDiagXY.ColWidth(DI_诊断编码) + .vsDiagXY.ColWidth(DI_诊断分类)
            End If
            '病案首页(出院科室或入院科室有产科性质)产科才有助产护士
            If gclsPros.PatiType <> PF_门诊 Then .vsOPS.ColHidden(PI_助产护士) = Not gclsPros.Is产科
        End With
    End If
    Screen.MousePointer = 0
    InitMedRecEnv = True
    Exit Function
errH:
    Debug.Print "InitMedRecEnv:" & Err.Source & "===" & Err.Description
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckDiagData(ByVal curDate As Date, Optional ByRef blnHaveSel As Boolean) As Boolean
    Dim vsTmp As VSFlexGrid
    Dim lngSize As Long
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim blnHaveDaig As Boolean
    Dim lng中医NotHave As Long '中医缺失诊断
    Dim lng西医NotHave As Long '西医缺失诊断
    Dim lngSame As Long '中医相同诊断不同证候的条数
    Dim lngSameType As Long
    
    On Error GoTo errH
    
    gclsPros.DiagNames = "": gclsPros.DiagRowIDs = ""
    gclsPros.DiseaseIDs = "": gclsPros.DiagIDs = ""
    '病案首页：中医或西医至少一种诊断的，门诊、入院、出院诊断齐全
    If gclsPros.FuncType = f病案首页 Then
        If gclsPros.Have中医 Then
            Set vsTmp = gclsPros.CurrentForm.vsDiagZY
            If vsTmp.TextMatrix(FindDiagRow(DT_门诊诊断ZY), DI_诊断描述) = "" Then
                lng中医NotHave = DT_门诊诊断ZY
            End If
            If lng中医NotHave = 0 Then
                If vsTmp.TextMatrix(FindDiagRow(DT_入院诊断ZY), DI_诊断描述) = "" Then
                    lng中医NotHave = DT_入院诊断ZY
                End If
            End If
            If lng中医NotHave = 0 Then
                If vsTmp.TextMatrix(FindDiagRow(DT_出院诊断ZY), DI_诊断描述) = "" Then
                    lng中医NotHave = DT_出院诊断ZY
                End If
            End If
        End If
        Set vsTmp = gclsPros.CurrentForm.vsDiagXY
         If vsTmp.TextMatrix(FindDiagRow(DT_门诊诊断XY), DI_诊断描述) = "" Then
             lng西医NotHave = DT_门诊诊断XY
         End If
         If lng西医NotHave = 0 Then
             If vsTmp.TextMatrix(FindDiagRow(DT_入院诊断XY), DI_诊断描述) = "" Then
                 lng西医NotHave = DT_入院诊断XY
             End If
         End If
         If lng西医NotHave = 0 Then
             If vsTmp.TextMatrix(FindDiagRow(DT_出院诊断XY), DI_诊断描述) = "" Then
                 lng西医NotHave = DT_出院诊断XY
             End If
         End If
         If lng西医NotHave <> 0 And (lng中医NotHave <> 0 And gclsPros.Have中医 Or Not gclsPros.Have中医) Then
             If gclsPros.Have中医 Then
                Set vsTmp = gclsPros.CurrentForm.vsDiagZY
                vsTmp.Row = FindDiagRow(lng中医NotHave): vsTmp.Col = DI_诊断描述
                If gclsPros.FuncType = f诊断选择 Then
                    Call ShowMessage(vsTmp, "中医诊断的" & decode(lng中医NotHave, DT_门诊诊断ZY, "门（急）诊诊断", DT_入院诊断ZY, "入院诊断", "出院诊断") & "的首要诊断不能为空。")
                    Exit Function
                Else
                    Call AddErrInfo("中医诊断的" & decode(lng中医NotHave, DT_门诊诊断ZY, "门（急）诊诊断", DT_入院诊断ZY, "入院诊断", "出院诊断") & "的首要诊断不能为空。", 0, vsTmp)
                End If
             Else
                Set vsTmp = gclsPros.CurrentForm.vsDiagXY
                vsTmp.Row = FindDiagRow(lng西医NotHave): vsTmp.Col = DI_诊断描述
                If gclsPros.FuncType = f诊断选择 Then
                    Call ShowMessage(vsTmp, "西医诊断的" & decode(lng西医NotHave, DT_门诊诊断XY, "门（急）诊诊断", DT_入院诊断XY, "入院诊断", "出院诊断") & "的首要诊断不能为空。")
                    Exit Function
                Else
                    Call AddErrInfo("西医诊断的" & decode(lng西医NotHave, DT_门诊诊断XY, "门（急）诊诊断", DT_入院诊断XY, "入院诊断", "出院诊断") & "的首要诊断不能为空。", 0, vsTmp)
                End If
             End If
         End If
    End If
    Set vsTmp = gclsPros.CurrentForm.vsDiagXY
    'gclsPros.InsureType = 920 And gclsPros.Module = p门诊医生站,原来的注释是北京医保的无理要求(不关我的事)
    lngSize = IIf(gclsPros.InsureType = 920 And gclsPros.PatiType = PF_门诊, 82, 200)
    With vsTmp
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, DI_诊断描述)) <> "" Then
                blnHaveDaig = True
                If i <> .Rows - 1 Then '检查是否存在类型相同且诊断相同的两行诊断
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, DI_诊断分类)) = Val(.TextMatrix(i, DI_诊断分类)) And Val(.TextMatrix(i, DI_诊断分类)) <> DT_病理诊断 Then
                            If .TextMatrix(j, DI_诊断描述) <> "" Then
                                If .TextMatrix(j, DI_诊断描述) = .TextMatrix(i, DI_诊断描述) Then
                                    .Row = i: .Col = DI_诊断描述
                                    If gclsPros.FuncType = f诊断选择 Then
                                        Call ShowMessage(vsTmp, "发现存在两行相同的诊断信息。")
                                        Exit Function
                                    Else
                                        If lngSameType = Val(.TextMatrix(i, DI_诊断分类)) Then
                                            Exit For
                                        Else
                                            Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "中发现存在相同的诊断信息。", 0, vsTmp)
                                            lngSameType = Val(.TextMatrix(i, DI_诊断分类))
                                            Exit For
                                        End If
                                    End If
                                ElseIf Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                                    If Val(.TextMatrix(j, DI_疾病ID)) = Val(.TextMatrix(i, DI_疾病ID)) Then
                                        .Row = i: .Col = DI_诊断描述
                                        If gclsPros.FuncType = f诊断选择 Then
                                            Call ShowMessage(vsTmp, "发现存在两行相同的诊断信息。")
                                            Exit Function
                                        Else
                                            If lngSameType = Val(.TextMatrix(i, DI_诊断分类)) Then
                                                Exit For
                                            Else
                                                Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "中发现存在相同的诊断信息。", 0, vsTmp)
                                                lngSameType = Val(.TextMatrix(i, DI_诊断分类))
                                                Exit For
                                            End If
                                        End If
                                    End If
                                ElseIf Val(.TextMatrix(i, DI_诊断ID)) <> 0 Then
                                    If Val(.TextMatrix(j, DI_诊断ID)) = Val(.TextMatrix(i, DI_诊断ID)) Then
                                        .Row = i: .Col = DI_诊断描述
                                        If gclsPros.FuncType = f诊断选择 Then
                                            Call ShowMessage(vsTmp, "发现存在两行相同的诊断信息。")
                                            Exit Function
                                        Else
                                            If lngSameType = Val(.TextMatrix(i, DI_诊断分类)) Then
                                                Exit For
                                            Else
                                                Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "中发现存在相同的诊断信息。", 0, vsTmp)
                                                lngSameType = Val(.TextMatrix(i, DI_诊断分类))
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                If Val(.TextMatrix(i, DI_诊断分类)) = DT_病理诊断 Then
                    If gclsPros.FuncType = f病案首页 Or gclsPros.FuncType = f医生首页 Then
                        If Not gclsPros.AddPathologic Then
                            If gclsPros.CurrentForm.txtInfo(GC_病理号).Text = "" Then
                                Call AddErrInfo("该诊断必须填写病理号，请填写病理号。", 0, gclsPros.CurrentForm.txtInfo(GC_病理号))
                            End If
                        End If
                        If gclsPros.CurrentForm.cboBaseInfo(BCC_临床与病理).Text = "" Then
                            Call AddErrInfo("该诊断必须填写临床与病理，请填写临床与病理。", 0, gclsPros.CurrentForm.cboBaseInfo(BCC_临床与病理))
                        End If
                        If gclsPros.CurrentForm.cboBaseInfo(BCC_放射与病理).Text = "" Then
                            Call AddErrInfo("该诊断必须填写放射与病理，请填写放射与病理。", 0, gclsPros.CurrentForm.cboBaseInfo(BCC_放射与病理))
                        End If
                    End If
                End If
                If .TextMatrix(i - 1, DI_诊断描述) = "" And Val(.TextMatrix(i, DI_诊断分类)) = Val(.TextMatrix(i - 1, DI_诊断分类)) Then
                    .Row = i - 1: .Col = DI_诊断描述
                    If gclsPros.FuncType = f诊断选择 Then
                        Call ShowMessage(vsTmp, "请依次输入诊断信息。")
                        Exit Function
                    Else
                        Call AddErrInfo("请依次输入诊断信息。", 0, vsTmp)
                    End If
                End If
                
                If zlCommFun.ActualLen(.TextMatrix(i, DI_诊断描述)) > lngSize Then
                    .Row = i: .Col = DI_诊断描述
                    If gclsPros.FuncType = f诊断选择 Then
                        Call ShowMessage(vsTmp, .TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "内容太长，只允许" & lngSize & "个字符或" & lngSize / 2 & "个汉字。")
                        Exit Function
                    Else
                        Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "内容太长，只允许" & lngSize & "个字符或" & lngSize / 2 & "个汉字。", 0, vsTmp)
                    End If
                End If
                If gclsPros.PatiType = PF_门诊 Then
                    If .TextMatrix(i, DI_发病时间) <> "" Then
                        If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, DI_发病时间), "YYYY-MM-DD HH:mm") Then
                             .Row = i: .Col = DI_发病时间
                            If gclsPros.FuncType = f诊断选择 Then
                                Call ShowMessage(vsTmp, "发病时间应该早于当前时间。")
                                Exit Function
                            Else
                                Call AddErrInfo("发病时间应该早于当前时间。", 0, vsTmp)
                            End If
                        End If
                    End If
                Else
                    If zlCommFun.ActualLen(.TextMatrix(i, DI_备注)) > 200 Then
                        .Row = i: .Col = DI_备注
                        If gclsPros.FuncType = f诊断选择 Then
                            Call ShowMessage(vsTmp, """" & .TextMatrix(i, DI_诊断描述) & """的备注内容太长，只允许200个字符或100个汉字。")
                            Exit Function
                        Else
                            Call AddErrInfo("""" & .TextMatrix(i, DI_诊断描述) & """的备注内容太长，只允许200个字符或100个汉字。", 0, vsTmp)
                        End If
                    End If
                    If gclsPros.FuncType = f病案首页 Then
                        If .TextMatrix(i, DI_附码ID) = .TextMatrix(i, DI_疾病ID) And Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                            .Row = i: .Col = DI_诊断编码
                            If gclsPros.FuncType = f诊断选择 Then
                                Call ShowMessage(vsTmp, "同一行诊断的主码与附码不能相同。")
                                Exit Function
                            Else
                                Call AddErrInfo("同一行诊断的主码与附码不能相同。", 0, vsTmp)
                            End If
                        End If
                        If .TextMatrix(i, DI_诊断编码) = "" And (Not gclsPros.CNIndent Or Val(.TextMatrix(i, DI_诊断分类)) <> DT_出院诊断XY) Then
                            If Val(.TextMatrix(i, DI_诊断分类)) <> DT_病理诊断 Then
                                .Row = i: .Col = DI_诊断编码
                                If gclsPros.FuncType = f诊断选择 Then
                                    Call ShowMessage(vsTmp, "该诊断必须填写诊断编码，请输入有编码的诊断或诊断编码。")
                                    Exit Function
                                Else
                                    Call AddErrInfo("该诊断必须填写诊断编码，请输入有编码的诊断或诊断编码。", 0, vsTmp)
                                End If
                            End If
                        End If
                    End If
                    
                    If .TextMatrix(i, DI_疗效限制) <> "" And InStr(.TextMatrix(i, DI_出院情况), .TextMatrix(i, DI_疗效限制)) > 0 Then
                        If gclsPros.FuncType = f诊断选择 Then
                            If ShowMessage(vsTmp, "“" & .TextMatrix(i, DI_诊断描述) & "”疾病的出院情况为“" & .TextMatrix(i, DI_出院情况) & "”" & _
                                vbCrLf & "是否确认？", True) = vbNo Then
                                .Row = i: .Col = DI_出院情况
                                Exit Function
                            End If
                        Else
                            .Row = i: .Col = DI_出院情况
                            Call AddErrInfo("“" & .TextMatrix(i, DI_诊断描述) & "”疾病的出院情况为“" & .TextMatrix(i, DI_出院情况) & "”是否确认？", 1, vsTmp)
                        End If
                    End If
                    If Val(.TextMatrix(i, DI_诊断分类)) = DT_院内感染 Then
                        If .TextMatrix(i, DI_出院情况) = "" Then
                            If gclsPros.FuncType = f病案首页 Then
                                If Not gclsPros.Null出院情况 Then
                                    .Row = i: .Col = DI_出院情况
                                    If gclsPros.FuncType = f诊断选择 Then
                                        Call ShowMessage(vsTmp, "请填写院内感染的出院情况。")
                                        Exit Function
                                    Else
                                        Call AddErrInfo("请填写院内感染的出院情况。", 0, vsTmp)
                                    End If
                                End If
                            Else
                                .Row = i: .Col = DI_出院情况
                                If gclsPros.FuncType = f诊断选择 Then
                                    If ShowMessage(vsTmp, "院内感染的出院情况没有填写，是否继续？", True) = vbNo Then Exit Function
                                Else
                                    Call AddErrInfo("院内感染的出院情况没有填写，是否继续？", 1, vsTmp)
                                End If
                            End If
                        End If
                    ElseIf Val(.TextMatrix(i, DI_诊断分类)) = DT_出院诊断XY Then
                        If .TextMatrix(i, DI_入院病情) = "" And DiagCellEditable(vsTmp, i, DI_入院病情) Then
                            .Row = i: .Col = DI_入院病情
                            If gclsPros.FuncType = f诊断选择 Then
                                Call ShowMessage(vsTmp, "请填写入院病情。")
                                Exit Function
                            Else
                                Call AddErrInfo("请填写入院病情。", 0, vsTmp)
                            End If
                        End If
                        
                        If .TextMatrix(i, DI_出院情况) = "" Then
                            If gclsPros.FuncType = f病案首页 Then
                                If Not gclsPros.Null出院情况 Then
                                    .Row = i: .Col = DI_出院情况
                                    If gclsPros.FuncType = f诊断选择 Then
                                        Call ShowMessage(vsTmp, "请填出院情况。")
                                        Exit Function
                                    Else
                                        Call AddErrInfo("请填出院情况。", 0, vsTmp)
                                    End If
                                End If
                            Else
                                .Row = i: .Col = DI_出院情况
                                If gclsPros.FuncType = f诊断选择 Then
                                    Call ShowMessage(vsTmp, "请填写出院情况。")
                                    Exit Function
                                Else
                                    Call AddErrInfo("请填出院情况。", 0, vsTmp)
                                End If
                            End If
                        End If
                        If .TextMatrix(i, DI_诊断类型) <> "出院诊断" Then
                            If InStr(.TextMatrix(FindDiagRow(DT_出院诊断XY), DI_出院情况), "死亡") = 0 And InStr(.TextMatrix(i, DI_出院情况), "死亡") > 0 Then
                                .Row = i: .Col = DI_出院情况
                                If gclsPros.FuncType = f诊断选择 Then
                                    If InStr(gclsPros.CurrentForm.txtInfo(GC_出院科室), "产科") = 0 Then
                                        Call ShowMessage(vsTmp, "主要诊断的出院情况不为死亡，但其它的诊断的出院情况却为死亡。")
                                        Exit Function
                                    End If
                                Else
                                    If InStr(gclsPros.CurrentForm.txtInfo(GC_出院科室), "产科") = 0 Then
                                        Call AddErrInfo("主要诊断的出院情况不为死亡，但其它的诊断的出院情况却为死亡。", 0, vsTmp)
                                    End If
                                End If
                            End If
                        Else '首要出院诊断
                            If InStr(.TextMatrix(i, DI_出院情况), "其他") > 0 And gclsPros.Have手术 Then
                                .Row = i: .Col = DI_出院情况
                                If gclsPros.FuncType = f诊断选择 Then
                                    If ShowMessage(vsTmp, "该病人进行了手术，但出院情况选择为其他。是否继续？", True) = vbNo Then Exit Function
                                Else
                                    Call AddErrInfo("该病人进行了手术，但出院情况选择为其他。是否继续？", 1, vsTmp)
                                End If
                            End If
                            If gclsPros.FuncType <> f诊断选择 Then
                                If InStr(.TextMatrix(i, DI_出院情况), "治愈") > 0 And Val(gclsPros.CurrentForm.txtSpecificInfo(SLC_住院天数).Text) < 3 Then
                                    .Row = i: .Col = DI_出院情况
                                    If gclsPros.FuncType = f诊断选择 Then
                                        If ShowMessage(vsTmp, "该病人住院天院为 " & Val(gclsPros.CurrentForm.txtSpecificInfo(SLC_住院天数).Text) & " 天，出院情况却为治愈，是否继续？", True) = vbNo Then Exit Function
                                    Else
                                        Call AddErrInfo("该病人住院天院为 " & Val(gclsPros.CurrentForm.txtSpecificInfo(SLC_住院天数).Text) & " 天，出院情况却为治愈，是否继续？", 1, vsTmp)
                                    End If
                                End If
                            End If
                            If gclsPros.Check损伤中毒 <> 0 Then
                                '主要诊断需要有损伤的外部原因
                                If InStr("ST", Left(.TextMatrix(i, DI_诊断编码), 1)) > 0 And Left(.TextMatrix(i, DI_诊断编码), 1) <> "" Then
                                    '需要损伤中毒外部原因
                                    If .TextMatrix(FindDiagRow(DT_损伤中毒码), DI_诊断描述) = "" Then
                                        .Row = FindDiagRow(DT_损伤中毒码): .Col = DI_诊断描述
                                        If gclsPros.Check损伤中毒 = 1 Then
                                            If gclsPros.FuncType = f诊断选择 Then
                                                Call ShowMessage(vsTmp, "请填写损伤中毒的原因。")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("请填写损伤中毒的原因。", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f诊断选择 Then
                                                If ShowMessage(vsTmp, "没有填写损伤中毒的原因,是否继续？", True) = vbNo Then Exit Function
                                            Else
                                                Call AddErrInfo("没有填写损伤中毒的原因,是否继续？", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                Else
                                    If .TextMatrix(FindDiagRow(DT_损伤中毒码), DI_诊断描述) <> "" Then
                                        .Row = FindDiagRow(DT_损伤中毒码): .Col = DI_诊断描述
                                        If gclsPros.Check损伤中毒 = 1 Then
                                            If gclsPros.FuncType = f诊断选择 Then
                                                Call ShowMessage(vsTmp, "不能填写损伤中毒的原因。")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("不能填写损伤中毒的原因。", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f诊断选择 Then
                                                If ShowMessage(vsTmp, "出院诊断与损伤中毒的原因不符,是否继续？", True) = vbNo Then Exit Function
                                            Else
                                                Call AddErrInfo("出院诊断与损伤中毒的原因不符,是否继续？", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            If gclsPros.Check病理诊断 <> 0 Then
                                '主要诊断需要填写病理诊断的外部原因
                                If (InStr("C", Left(.TextMatrix(i, DI_诊断编码), 1)) > 0 Or (InStr("D", Left(.TextMatrix(i, DI_诊断编码), 1)) > 0 And Val(Mid(.TextMatrix(i, DI_诊断编码), 2, 2)) <= 48)) And Left(.TextMatrix(i, DI_诊断编码), 1) <> "" Then
                                    '需要病理诊断的外部原因
                                    If .TextMatrix(FindDiagRow(DT_病理诊断), DI_诊断描述) = "" Then
                                        .Row = FindDiagRow(DT_病理诊断): .Col = DI_诊断描述
                                        If gclsPros.Check病理诊断 = 1 Then
                                            If gclsPros.FuncType = f诊断选择 Then
                                                Call ShowMessage(vsTmp, "出院诊断为肿瘤诊断,请填写病理诊断。")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("出院诊断为肿瘤诊断,请填写病理诊断。", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f诊断选择 Then
                                                If ShowMessage(vsTmp, "出院诊断为肿瘤诊断,没有填写病理诊断,是否继续？", True) = vbNo Then Exit Function
                                            Else
                                                Call AddErrInfo("出院诊断为肿瘤诊断,没有填写病理诊断,是否继续？", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                Else
                                    If .TextMatrix(FindDiagRow(DT_病理诊断), DI_诊断描述) <> "" Then
                                        .Row = FindDiagRow(DT_病理诊断): .Col = DI_诊断描述
                                        If gclsPros.Check病理诊断 = 1 Then
                                            If gclsPros.FuncType = f诊断选择 Then
                                                Call ShowMessage(vsTmp, "不能填写病理诊断。")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("不能填写病理诊断。", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f诊断选择 Then
                                                If ShowMessage(vsTmp, "出院诊断与病理诊断不符,是否继续？", True) = vbNo Then Exit Function
                                            Else
                                                 Call AddErrInfo("出院诊断与病理诊断不符,是否继续？", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                        End If
                        If (InStr("C", Left(.TextMatrix(i, DI_诊断编码), 1)) > 0 Or (InStr("D", Left(.TextMatrix(i, DI_诊断编码), 1)) > 0 And Val(Mid(.TextMatrix(i, DI_诊断编码), 2, 2)) <= 48)) And Left(.TextMatrix(i, DI_诊断编码), 1) <> "" Then
                            'ICD附码是否必须填写
                            If gclsPros.CheckICD附码 <> 0 Then
                                .Row = i: .Col = DI_ICD附码
                                If .TextMatrix(i, DI_ICD附码) = "" Then
                                    If gclsPros.CheckICD附码 = 1 Then
                                        If gclsPros.FuncType = f诊断选择 Then
                                            Call ShowMessage(vsTmp, "当前诊断为肿瘤诊断,请填写肿瘤形态学编码。")
                                            Exit Function
                                        Else
                                            Call AddErrInfo("当前诊断为肿瘤诊断,请填写肿瘤形态学编码。", 0, vsTmp)
                                        End If
                                    Else
                                        If gclsPros.FuncType = f诊断选择 Then
                                            If ShowMessage(vsTmp, "当前诊断为肿瘤诊断,没有填写肿瘤形态学编码,是否继续？", True) = vbNo Then Exit Function
                                        Else
                                            Call AddErrInfo("当前诊断为肿瘤诊断,没有填写肿瘤形态学编码,是否继续？", 1, vsTmp)
                                        End If
                                    End If
                                Else
                                    If Left(.TextMatrix(i, DI_ICD附码), 1) <> "M" Then

                                        If gclsPros.CheckICD附码 = 1 Then
                                            'ICD附码必须是M开头的
                                            If gclsPros.FuncType = f诊断选择 Then
                                                Call ShowMessage(vsTmp, "当前诊断为肿瘤诊断,只允许填写肿瘤形态学编码(M)。")
                                                Exit Function
                                            Else
                                                Call AddErrInfo("当前诊断为肿瘤诊断,只允许填写肿瘤形态学编码(M)。", 0, vsTmp)
                                            End If
                                        Else
                                            If gclsPros.FuncType = f诊断选择 Then
                                                If ShowMessage(vsTmp, "当前诊断为肿瘤诊断,只允许填写肿瘤形态学编码(M),是否继续？", True) = vbNo Then Exit Function
                                            Else
                                                Call AddErrInfo("当前诊断为肿瘤诊断,只允许填写肿瘤形态学编码(M),是否继续？", 1, vsTmp)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                            
                            
  
                If Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then gclsPros.DiseaseIDs = gclsPros.DiseaseIDs & "," & Val(.TextMatrix(i, DI_疾病ID))
                If Val(.TextMatrix(i, DI_诊断ID)) <> 0 Then gclsPros.DiagIDs = gclsPros.DiagIDs & "," & Val(.TextMatrix(i, DI_诊断ID))
                '是否输入了要求的诊断类型
                If gclsPros.PatiType = PF_门诊 Then
                    gclsPros.IsDiagInput = True
                Else
                    If InStr("," & gclsPros.MustDiagType & ",", "," & Val(.TextMatrix(i, DI_诊断分类)) & ",") > 0 Then
                        gclsPros.IsDiagInput = True
                    End If
                End If
                If Val(.TextMatrix(i, DI_关联)) <> 0 Then
                    gclsPros.DiagRowIDs = gclsPros.DiagRowIDs & IIf(gclsPros.DiagRowIDs <> "", ",", "") & .RowData(i)
                    strTmp = IIf(Trim(.TextMatrix(i, DI_诊断编码)) = "", "", "(" & .TextMatrix(i, DI_诊断编码) & ")") & .TextMatrix(i, DI_诊断描述)
                    gclsPros.DiagNames = gclsPros.DiagNames & IIf(gclsPros.DiagNames <> "", ",", "") & strTmp
                    .Cell(flexcpData, i, DI_关联) = ""
                    blnHaveSel = True
                End If
            End If
        Next
    End With

    If gclsPros.Have中医 Then
        Set vsTmp = gclsPros.CurrentForm.vsDiagZY
        With vsTmp
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, DI_诊断描述) <> "" Then
                    blnHaveDaig = True
                    lngSame = 0
                    lngSameType = 0
                    If i <> .Rows - 1 Then '检查是否存在类型相同且诊断相同的两行诊断
                        For j = i + 1 To .Rows - 1
                            If Val(.TextMatrix(j, DI_诊断分类)) = Val(.TextMatrix(i, DI_诊断分类)) Then
                                If Trim(.TextMatrix(j, DI_诊断描述)) <> "" Then
                                    If .TextMatrix(j, DI_诊断描述) & "|" & .TextMatrix(j, DI_中医证候) = .TextMatrix(i, DI_诊断描述) & "|" & .TextMatrix(i, DI_中医证候) Then
                                        .Row = i: .Col = DI_诊断描述
                                        If gclsPros.FuncType = f诊断选择 Then
                                            Call ShowMessage(vsTmp, "发现存在两行相同的诊断信息。")
                                            Exit Function
                                        Else
                                            If lngSameType = Val(.TextMatrix(i, DI_诊断分类)) Then
                                                Exit For
                                            Else
                                                Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "中发现存在相同的诊断信息。", 0, vsTmp)
                                                lngSameType = Val(.TextMatrix(i, DI_诊断分类))
                                                Exit For
                                            End If
                                        End If
                                    ElseIf Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                                        If Val(.TextMatrix(j, DI_疾病ID)) & "|" & .TextMatrix(j, DI_中医证候) = Val(.TextMatrix(i, DI_疾病ID)) & "|" & .TextMatrix(i, DI_中医证候) Then
                                            .Row = i: .Col = DI_诊断描述
                                            If gclsPros.FuncType = f诊断选择 Then
                                                Call ShowMessage(vsTmp, "发现存在两行相同的诊断信息。")
                                                Exit Function
                                            Else
                                                If lngSameType = Val(.TextMatrix(i, DI_诊断分类)) Then
                                                    Exit For
                                                Else
                                                    Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "中发现存在相同的诊断信息。", 0, vsTmp)
                                                    lngSameType = Val(.TextMatrix(i, DI_诊断分类))
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                    If .TextMatrix(j, DI_诊断描述) = .TextMatrix(i, DI_诊断描述) Then
                                        lngSame = lngSame + 1
                                    ElseIf Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                                        If Val(.TextMatrix(j, DI_疾病ID)) = Val(.TextMatrix(i, DI_疾病ID)) Then
                                            lngSame = lngSame + 1
                                        End If
                                    End If
                                    If lngSame >= 2 Then
                                        .Row = i: .Col = DI_诊断描述
                                        If gclsPros.FuncType = f诊断选择 Then
                                            Call ShowMessage(vsTmp, "存在两条以上的诊断相同且证候不同的诊断，诊断不明确。")
                                            Exit Function
                                        Else
                                            If lngSameType = Val(.TextMatrix(i, DI_诊断分类)) Then
                                                Exit For
                                            Else
'                                                Call AddErrInfo("存在两条以上的诊断相同且证候不同的诊断，诊断不明确。", 0, vsTmp)
                                                Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "中存在两条以上的诊断相同且证候不同的诊断，诊断不明确。", 0, vsTmp)
                                                lngSameType = Val(.TextMatrix(i, DI_诊断分类))
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If
                    If i <> 0 Then
                        If .TextMatrix(i - 1, DI_诊断描述) = "" And Val(.TextMatrix(i, DI_诊断分类)) = Val(.TextMatrix(i - 1, DI_诊断分类)) Then
                            .Row = i - 1: .Col = DI_诊断描述
                            If gclsPros.FuncType = f诊断选择 Then
                                Call ShowMessage(vsTmp, "请依次输入诊断信息。")
                                Exit Function
                            Else
                                Call AddErrInfo("请依次输入诊断信息。", 0, vsTmp)
                            End If
                        End If
                    End If
                    
                    If zlCommFun.ActualLen(.TextMatrix(i, DI_诊断描述)) + zlCommFun.ActualLen(.TextMatrix(i, DI_中医证候)) > lngSize Then
                        .Row = i: .Col = DI_诊断描述
                        If gclsPros.FuncType = f诊断选择 Then
                            Call ShowMessage(vsTmp, .TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "的诊断描述或中医证候内容太长，诊断描述和中医证候加起来只允许" & lngSize & "个字符或" & lngSize / 2 & "个汉字。")
                            Exit Function
                        Else
                            Call AddErrInfo(.TextMatrix(IIf(.TextMatrix(i, DI_诊断类型) = "", FindDiagRow(Val(.TextMatrix(i, DI_诊断分类))), i), DI_诊断类型) & "的诊断描述或中医证候内容太长，诊断描述和中医证候加起来只允许" & lngSize & "个字符或" & lngSize / 2 & "个汉字。", 0, vsTmp)
                        End If
                    End If
                    
                    If gclsPros.PatiType = PF_门诊 Then
                        If .TextMatrix(i, DI_发病时间) <> "" Then
                            If Format(curDate, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, DI_发病时间), "YYYY-MM-DD HH:mm") Then
                                 .Row = i: .Col = DI_发病时间
                                If gclsPros.FuncType = f诊断选择 Then
                                    Call ShowMessage(vsTmp, "发病时间应该早于当前时间。")
                                    Exit Function
                                Else
                                    Call AddErrInfo("发病时间应该早于当前时间。", 0, vsTmp)
                                End If
                            End If
                        End If
                        
                         '中医诊断和西医诊断的自由录入医嘱不能存在相同的
                        If .TextMatrix(i, DI_诊断编码) = "" Then
                            For j = gclsPros.CurrentForm.vsDiagXY.FixedRows To gclsPros.CurrentForm.vsDiagXY.Rows - 1
                                If gclsPros.CurrentForm.vsDiagXY.TextMatrix(j, DI_诊断编码) = "" Then
                                    If gclsPros.CurrentForm.vsDiagXY.TextMatrix(j, DI_诊断描述) = .TextMatrix(i, DI_诊断描述) Then
                                        .Row = i: .Col = DI_诊断描述
                                        If gclsPros.FuncType = f诊断选择 Then
                                            Call ShowMessage(vsTmp, "发现存在两行相同的自由录入诊断信息(中医诊断与西医诊断)。")
                                            Exit Function
                                        Else
                                            Call AddErrInfo("发现存在两行相同的自由录入诊断信息(中医诊断与西医诊断)。", 0, vsTmp)
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Else
                        If zlCommFun.ActualLen(.TextMatrix(i, DI_备注)) > 200 Then
                            .Row = i: .Col = DI_备注
                            If gclsPros.FuncType = f诊断选择 Then
                                Call ShowMessage(vsTmp, """" & .TextMatrix(i, DI_诊断描述) & """的备注内容太长，只允许200个字符或100个汉字。")
                                Exit Function
                            Else
                                Call AddErrInfo("""" & .TextMatrix(i, DI_诊断描述) & """的备注内容太长，只允许200个字符或100个汉字。", 0, vsTmp)
                            End If
                        End If
                        If Val(.TextMatrix(i, DI_诊断分类)) = DT_出院诊断ZY Then
                            If .TextMatrix(i, DI_入院病情) = "" And DiagCellEditable(vsTmp, i, DI_入院病情) Then
                                .Row = i: .Col = DI_入院病情
                                If gclsPros.FuncType = f诊断选择 Then
                                    Call ShowMessage(vsTmp, "请填写入院病情。")
                                    Exit Function
                                Else
                                    Call AddErrInfo("请填写入院病情。", 0, vsTmp)
                                End If
                            End If
                            If .TextMatrix(i, DI_出院情况) = "" Then
                                .Row = i: .Col = DI_出院情况
                                If gclsPros.FuncType = f诊断选择 Then
                                    Call ShowMessage(vsTmp, "请填写出院诊断的出院情况。")
                                    Exit Function
                                Else
                                    Call AddErrInfo("请填写出院诊断的出院情况。", 0, vsTmp)
                                End If
                            End If
                            
                            If Val(.TextMatrix(i - 1, DI_诊断分类)) = DT_出院诊断ZY And InStr(.TextMatrix(FindDiagRow(DT_出院诊断ZY), DI_出院情况), "死亡") = 0 And InStr(.TextMatrix(i, DI_出院情况), "死亡") > 0 Then
                                .Row = i: .Col = DI_出院情况
                                If gclsPros.FuncType = f诊断选择 Then
                                    Call ShowMessage(vsTmp, "主要诊断的出院情况不为死亡，但其它诊断的出院情况却为死亡。")
                                    Exit Function
                                Else
                                    Call AddErrInfo("主要诊断的出院情况不为死亡，但其它诊断的出院情况却为死亡。", 0, vsTmp)
                                End If
                            End If
                        End If
                    End If
                    If Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then gclsPros.DiseaseIDs = gclsPros.DiseaseIDs & "," & Val(.TextMatrix(i, DI_疾病ID))
                    If Val(.TextMatrix(i, DI_诊断ID)) <> 0 Then gclsPros.DiagIDs = gclsPros.DiagIDs & "," & Val(.TextMatrix(i, DI_诊断ID))
                    '是否输入了要求的诊断类型
                    If gclsPros.PatiType = PF_门诊 Then
                        gclsPros.IsDiagInput = True
                    Else
                        If InStr("," & gclsPros.MustDiagType & ",", "," & Val(.TextMatrix(i, DI_诊断分类)) & ",") > 0 Then
                            gclsPros.IsDiagInput = True
                        End If
                    End If
                    If Val(.TextMatrix(i, DI_关联)) <> 0 Then
                        gclsPros.DiagRowIDs = gclsPros.DiagRowIDs & IIf(gclsPros.DiagRowIDs <> "", ",", "") & .RowData(i)
                        strTmp = IIf(Trim(.TextMatrix(i, DI_诊断编码)) = "", "", "(" & .TextMatrix(i, DI_诊断编码) & ")") & .TextMatrix(i, DI_诊断描述) & IIf(.TextMatrix(i, DI_中医证候) <> "", "(" & .TextMatrix(i, DI_中医证候) & ")", "") & IIf(.TextMatrix(i, DI_中医证候) <> "", "(" & .TextMatrix(i, DI_中医证候) & ")", "")
                        gclsPros.DiagNames = gclsPros.DiagNames & IIf(gclsPros.DiagNames <> "", ",", "") & strTmp
                        .Cell(flexcpData, i, DI_关联) = ""
                        blnHaveSel = True
                    End If
                End If
            Next
        End With
    End If
    If gclsPros.FuncType = f病案首页 And Not blnHaveDaig Then
        If gclsPros.FuncType = f诊断选择 Then
            Call ShowMessage(gclsPros.CurrentForm.vsDiagXY, "西医诊断和中医诊断都没有输入,请检查!")
            Exit Function
        Else
            Call AddErrInfo("西医诊断和中医诊断都没有输入,请检查!", 0, gclsPros.CurrentForm.vsDiagXY)
        End If
    End If
    If gclsPros.DiseaseIDs <> "" Then gclsPros.DiseaseIDs = Mid(gclsPros.DiseaseIDs, 2)
    If gclsPros.DiagIDs <> "" Then gclsPros.DiagIDs = Mid(gclsPros.DiagIDs, 2)
    CheckDiagData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Check留观(Optional blnCheck As Boolean) As Boolean
    Dim objCbo As ComboBox, objMSK As MaskEdBox
    Dim curDate As Date
    Dim blnJudge As Boolean
    Dim colErr As New Collection
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    Call ClearErrCol
    Set colErrTmp = colErr  '清空集合
    curDate = zlDatabase.Currentdate
    With gclsPros.CurrentForm
        If .txtSpecificInfo(SLC_年龄).Enabled And Not .txtSpecificInfo(SLC_年龄).Locked Then
            '项目输入长度检查
            If .txtSpecificInfo(SLC_年龄).MaxLength <> 0 And .txtSpecificInfo(SLC_年龄).Text <> "" Then
                strTmp = .txtSpecificInfo(SLC_年龄).Text
                strTmp = strTmp & .cboSpecificInfo(SLC_年龄).Text
                If zlCommFun.ActualLen(strTmp) > .txtSpecificInfo(SLC_年龄).MaxLength Then
                    Call AddErrInfo("输入内容过长。(该项目最多允许 " & .txtSpecificInfo(SLC_年龄).MaxLength & " 个字符或 " & .txtSpecificInfo(SLC_年龄).MaxLength / 2 & " 个汉字)", 0, .txtSpecificInfo(SLC_年龄))
                End If
            End If
        End If
        If .txtInfo(GC_姓名).Enabled And Not .txtInfo(GC_姓名).Locked Then
            If .txtInfo(GC_姓名).Text = "" Then
                Call AddErrInfo("病人的姓名不能为空，请输病人的姓名。", 0, .txtInfo(GC_姓名))
            End If
        End If
        If .txtSpecificInfo(SLC_年龄).Text = "" Then
            Call AddErrInfo("病人的年龄不能为空，请输入病人的年龄。", 0, .txtSpecificInfo(SLC_年龄))
        Else
            '小时为单位，不能大于30天即720小时
            '分钟为单位，不能大于24小时即1440分钟
            If .cboSpecificInfo(SLC_年龄).Visible Then
                strTmp = .cboSpecificInfo(SLC_年龄).Text
                i = decode(strTmp, "岁", 200, "月", 2400, "天", 73000, "小时", 720, "分钟", 1440, 0)
                If Val(.txtSpecificInfo(SLC_年龄).Text) > i And strTmp <> "" Then
                    Call AddErrInfo("年龄值超过最大限制" & i & strTmp & IIf(i = 1440 Or i = 720, "，请使用合适的年龄单位。", "。"), 0, .txtSpecificInfo(SLC_年龄), .cboSpecificInfo(SLC_年龄))
                ElseIf Val(.txtSpecificInfo(SLC_年龄).Text) < 0 Then
                    Call AddErrInfo("年龄值不能为负数。", 0, .txtSpecificInfo(SLC_年龄), .cboSpecificInfo(SLC_年龄))
                End If
            End If
        End If
        strTmp = ""
        
        '必须要输入的内容检查
        If gclsPros.PatiType = PF_门诊 Then
            If InStr(GetInsidePrivs(p病人信息公共部件), "基本信息调整") > 0 Then
                If .cboBaseInfo(BCC_付款方式).Enabled And Not .cboBaseInfo(BCC_付款方式).Locked Then
                    strTmp = "付款方式"
                    If .cboBaseInfo(BCC_付款方式).ListIndex = -1 Then
                       Call AddErrInfo("请输入病人的" & strTmp & "。", 0, .cboBaseInfo(BCC_付款方式))
                    End If
                 End If
            End If
        Else
            If .cboBaseInfo(BCC_付款方式).Enabled And Not .cboBaseInfo(BCC_付款方式).Locked Then
                strTmp = "付款方式"
                If .cboBaseInfo(BCC_付款方式).ListIndex = -1 Then
                    Call AddErrInfo("请输入病人的" & strTmp & "。", 0, .cboBaseInfo(BCC_付款方式))
                End If
            End If
            If .cboBaseInfo(BCC_国籍).Enabled And Not .cboBaseInfo(BCC_国籍).Locked Then
                strTmp = "国籍"
                If .cboBaseInfo(BCC_国籍).ListIndex = -1 Then
                     Call AddErrInfo("请输入病人的" & strTmp & "。", 0, .cboBaseInfo(BCC_国籍))
                End If
            End If
            If .cboBaseInfo(BCC_性别).Enabled And Not .cboBaseInfo(BCC_性别).Locked Then
                strTmp = "BCC_性别"
                If .cboBaseInfo(BCC_性别).ListIndex = -1 Then
                    Call AddErrInfo("请输入病人的" & strTmp & "。", 0, .cboBaseInfo(BCC_性别))
                End If
            End If
        End If
        If .mskDateInfo(DC_出生日期).Enabled Then
            blnJudge = .mskDateInfo(DC_出生日期).Text = Replace(.mskDateInfo(DC_出生日期).Mask, "#", "_")
            If blnJudge Then
                Call AddErrInfo("请输入病人的出生日期。", 0, .mskDateInfo(DC_出生日期))
            End If
            If Not IsDate(.mskDateInfo(DC_出生日期).Text) Then
                Call AddErrInfo("出生日期不是有效的日期格式。", 0, .mskDateInfo(DC_出生日期))
            End If
            If Format(.mskDateInfo(DC_出生日期).Text, "yyyy-MM-dd hh:mm") > Format(gclsPros.InTime, "yyyy-MM-dd hh:mm") Then
                Call AddErrInfo("出生日期在入院时间之后。", 0, .mskDateInfo(DC_出生日期), .mskDateInfo(DC_入院时间))
            End If
        End If
        Call CheckDiagData(curDate)
    End With
    If gColErr.Count > 0 Or gColWarn.Count > 0 Then
        Call LoadVsErrData
        If Not blnCheck Then
            If gColErr.Count = 0 And gColWarn.Count > 0 Then
                If MsgBox("检查出" & CStr(gColWarn.Count) & "个警告，是否忽略全部警告，继续操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                Else
                    Call ClearErrCol
                    Check留观 = True
                End If
            ElseIf gColErr.Count > 0 Then
                Check留观 = False
                Exit Function
            End If
        End If
    End If
    If Not CheckMedPageChange Then
        gclsPros.InfosChange = False
        gclsPros.IsCheckData = False
        Exit Function
    Else
        gclsPros.IsCheckData = False
    End If
    Check留观 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function CheckMedPageData(ByRef blnDiagnose As Boolean, Optional blnCheck As Boolean) As Boolean
'功能：检查首页输入数据合法性
'返回：blnDiagnose=是否填写了诊断
'参数：
    Dim objTextBox As TextBox, objCbo As ComboBox, objMSK As MaskEdBox, objChk As CheckBox, vsTmp As VSFlexGrid
    Dim blnJudge As Boolean, strTmp As String
    Dim curDate As Date
    Dim lngSize As Long
    Dim str疾病IDs As String, str诊断IDs As String
    Dim i As Long, j As Long
    Dim strSql As String, rsTmp As Recordset
    Dim sgnTmp As Single, blnDo As Boolean, blnDoEx As Boolean
    Dim strBirthday As String, strAge As String, strSex As String, strErrIfno As String, str年龄 As String
    Dim objTmp As Object
    Dim blnDateIsNull As Boolean
    Dim blnBaseInfo As Boolean, strBaseInfo As String
    Dim strMask As String, arrTmp As Variant
    Dim str入住时间 As String, str转出时间 As String
    Dim strMsg As String
    Dim colErr As New Collection
    

    blnDiagnose = False
    gclsPros.IsCheckData = True
'    gclsPros.InfosChange = False
    '清理掉之前的警告错误
    Call ClearErrCol
    Set colErrTmp = colErr  '清空集合
    '并发检查病案是否编目或首页处于锁定状态
    If gclsPros.PatiType <> PF_门诊 And gclsPros.FuncType <> f病案首页 Then
        If Not CheckMecRed(gclsPros.病人ID, gclsPros.主页ID, gclsPros.CurrentForm.Caption, "修改首页") Then Exit Function
    End If

    curDate = zlDatabase.Currentdate
    With gclsPros.CurrentForm
        'txtInfo控件相关检查
        For Each objTextBox In .txtInfo
            If objTextBox.Enabled And Not objTextBox.Locked Then
                '项目输入长度检查
                If objTextBox.MaxLength <> 0 And objTextBox.Text <> "" Then
                    If zlCommFun.ActualLen(objTextBox.Text) > objTextBox.MaxLength Then
                        Call AddErrInfo("输入内容过长。(该项目最多允许 " & objTextBox.MaxLength & " 个字符或 " & objTextBox.MaxLength \ 2 & " 个汉字)", 0, objTextBox)
                    End If
                End If
                Select Case objTextBox.Index
                    Case GC_监护人身份证号
                        strTmp = objTextBox.Text
                        If strTmp <> "" Then
                            If Trim(zlCommFun.GetNeedName(.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                                If zlCommFun.ActualLen(strTmp) = Len(strTmp) Then
                                    If gclsPros.IsMaskID Then strTmp = objTextBox.Tag
                                        '初始化病人信息接口
                                        If gobjPatient Is Nothing Then
                                            On Error Resume Next
                                            Set gobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                                            Err.Clear: On Error GoTo errH
                                            Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
                                        End If
                                        If gobjPatient Is Nothing Then
                                            Call AddErrInfo("创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败，不能进行病人身份证信息检查！是否继续？", 1, objCbo)
                                        End If
                                        If Not gobjPatient.CheckPatiIdcard(strTmp, strBirthday, strAge, strSex, strErrIfno, CDate(gclsPros.InTime)) Then '身份证合法则检查是否匹配
                                            '身份证不合法则退出
                                            Call AddErrInfo(strErrIfno, 0, objTextBox)
                                        End If
                                ElseIf zlCommFun.ActualLen(strTmp) > 18 Then
                                    Call AddErrInfo("身份证号不能超过9个汉字或18个英文字符的长度，请检查。", 0, objTextBox)
                                End If
                            End If
                        End If
                End Select
                '必须要输入的内容检查
                If objTextBox.Index = GC_Email Then
                     If objTextBox.Text <> "" Then
                        If InStr(objTextBox.Text, "@") <= 1 Or InStr(objTextBox.Text, ".") <= 3 Or InStr(objTextBox.Text, "@") > InStr(objTextBox.Text, ".") Then
                            Call AddErrInfo("输入的Email的格式不正确，正确格式：""XXX@XX.XX""。", 0, objTextBox)
                        End If
                    End If
                ElseIf objTextBox.Index = GC_转科2 Then
                     If objTextBox.Text <> "" Then
                        If .txtInfo(GC_转科1).Text = "" Then
                            Call AddErrInfo("没有依次输入转科科室，请依次输入。", 0, .txtInfo(GC_转科1), objTextBox)
                        ElseIf .txtInfo(GC_转科1).Text = objTextBox.Text Or .txtInfo(GC_转科3).Text = objTextBox.Text Then
                            Call AddErrInfo("转科的两个科室不应该相同。", 0, objTextBox, .txtInfo(IIf(.txtInfo(GC_转科1).Text = objTextBox.Text, GC_转科1, GC_转科3)))
                        End If
                    Else
                         If .txtInfo(GC_转科3).Text <> "" Then
                            Call AddErrInfo("没有依次输入转科科室，请依次输入。", 0, objTextBox, .txtInfo(GC_转科3))
                        End If
                    End If
                ElseIf objTextBox.Index = GC_31天内再住院 Then
                    If Trim(objTextBox.Text) = "" Then
                        Call AddErrInfo(.cboBaseInfo(BCC_再入院计划天数).Text & "的目的没有填写。", 0, objTextBox)
                    End If
                ElseIf objTextBox.Index = BCC_死亡期间 Then
                    If objTextBox.Text <> "" Then
                        If .cboBaseInfo(BCC_死亡期间).Text = "" And objTextBox.Index = GC_死亡原因 Then
                            Call AddErrInfo("输入了死亡原因，但没有录入死亡期间，是否继续？", 1, .cboBaseInfo(BCC_死亡期间), objTextBox)
                        End If
                    End If
                ElseIf gclsPros.FuncType = f病案首页 Then
                    If objTextBox.Text = "" Then
                        If objTextBox.Index = GC_姓名 Then
                            Call AddErrInfo("病人的姓名不能为空，请输病人的姓名。", 0, objTextBox)
                        ElseIf objTextBox.Index = GC_入院科室 Then
                            Call AddErrInfo("病人的入院科室不能为空，请输病人的入院科室。", 0, objTextBox)
                        ElseIf objTextBox.Index = GC_出院科室 Then
                            Call AddErrInfo("病人的出院科室不能为空，请输病人的出院科室。", 0, objTextBox)
                        End If
                    End If
                End If
            End If
        Next

        '获取上次出院时间以及下次入院时间
        If gclsPros.FuncType = f病案首页 Then
            If gclsPros.InTime = "" Then
                Call AddErrInfo("病人的入院日期不能为空，请输入病人的入院日期。", 0, .mskDateInfo(DC_入院时间))
            End If

            If gclsPros.OutTime = "" Then
                Call AddErrInfo("病人的出院日期不能为空，请输入病人的出院日期。", 0, .mskDateInfo(DC_出院时间))
            End If
            If gclsPros.InTime > Format(curDate, "yyyy-MM-dd hh:mm:ss") Then
                Call AddErrInfo("病人的入院日期不能晚于于当前时间。", 0, .mskDateInfo(DC_入院时间))
            End If
            If gclsPros.OutTime > Format(curDate, "yyyy-MM-dd hh:mm:ss") Then
                Call AddErrInfo("病人的出院日期不能晚于于当前时间。", 0, .mskDateInfo(DC_出院时间))
            End If
            If gclsPros.InTime > gclsPros.OutTime Then
                Call AddErrInfo("病人的出院日期不能早于入院时间。", 0, .mskDateInfo(DC_入院时间), .mskDateInfo(DC_出院时间))
            End If

            strSql = "Select 主页id, 下次入院, 上次出院" & vbNewLine & _
                    "From (Select 主页id, Lead(入院日期, 1, Null) Over(Order By 主页id) 下次入院, Lag(出院日期, 1, Null) Over(Order By 主页id) 上次出院" & vbNewLine & _
                    "       From 病案主页" & vbNewLine & _
                    "       Where 病人id = [1])" & vbNewLine & _
                    "Where 主页id = [2]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取临近主页ID的入出院时间", gclsPros.病人ID, gclsPros.主页ID)
            If rsTmp.RecordCount > 0 Then
                If Not IsNull(rsTmp!上次出院) Then
                    strTmp = Format(rsTmp!上次出院 & "", "yyyy-MM-dd hh:mm")
                    '本次入出院时间检查
                    If Format(gclsPros.InTime, "yyyy-MM-dd hh:mm") < strTmp Then
                        Call AddErrInfo("本次入院日期早于了上次的出院日期(" & strTmp & ")。", 0, .mskDateInfo(DC_入院时间))
                    End If
                End If
                If Not IsNull(rsTmp!下次入院) Then
                    strTmp = Format(rsTmp!下次入院 & "", "yyyy-MM-dd hh:mm")
                    If Format(gclsPros.OutTime, "yyyy-MM-dd hh:mm") > strTmp Then
                        Call AddErrInfo("本次出院日期晚于了下一次的入院日期(" & strTmp & ")。", 0, .mskDateInfo(DC_出院时间))
                    End If
                End If
            End If
        End If

        'txtSpecificInfo控件相关检查
        For Each objTextBox In .txtSpecificInfo
            If objTextBox.Enabled And Not objTextBox.Locked Then
                '项目输入长度检查
                If objTextBox.MaxLength <> 0 And objTextBox.Text <> "" Then
                    strTmp = objTextBox.Text
                    If objTextBox.Index = SLC_年龄 Then
                        strTmp = strTmp & .cboSpecificInfo(SLC_年龄).Text
                    End If
                    If zlCommFun.ActualLen(strTmp) > objTextBox.MaxLength Then
                        Call AddErrInfo("输入内容过长。(该项目最多允许 " & objTextBox.MaxLength & " 个字符或 " & objTextBox.MaxLength / 2 & " 个汉字)", 0, objTextBox)
                    End If
                End If

                Select Case objTextBox.Index
                    Case SLC_家庭电话, SLC_单位电话, SLC_联系人电话
                        strMask = "1234567890-()"
                        For i = 1 To Len(objTextBox.Text)
                            If InStr(strMask, Mid(objTextBox.Text, i, 1)) = 0 Then
                                Call AddErrInfo("输入的内容出现了非法字符，该项允许输入的字符为(" & strMask & ")", 0, objTextBox)
                                Exit For
                            End If
                        Next
                    Case SLC_住院号, SLC_户口邮编, SLC_家庭邮编, SLC_单位邮编, SLC_抢救次数, SLC_成功次数, _
                        SLC_随诊期限, SLC_昏迷时间入院前_小时, SLC_昏迷时间入院前_分钟, SLC_昏迷时间入院后_分钟, SLC_昏迷时间入院后_小时, _
                        SLC_昏迷时间入院前_天, SLC_昏迷时间入院后_天, SLC_呼吸机使用, SLC_重症监护天, SLC_重症监护小时, SLC_QQ, SLC_外院会诊, SLC_院内会诊
                        strMask = "1234567890"
                        If objTextBox.Text <> "" Then
                            If Not IsNumeric(objTextBox.Text) Then
                                Call AddErrInfo("输入的内容有误，该项只能够输入整数。", 0, objTextBox)
                            End If
                        End If
                    Case SLC_输红细胞, SLC_输血小板, SLC_输血浆, SLC_输全血, SLC_输白蛋白, SLC_自体回收, SLC_ICU, _
                         SLC_CCU, SLC_一级护理, SLC_二级护理, SLC_三级护理, SLC_特护, SLC_约束总时间, SLC_身高, SLC_体重, SLC_距上次住院时间
                        strMask = "1234567890."
                        If objTextBox.Text <> "" Then
                            If Not IsNumeric(objTextBox.Text) Then
                                Call AddErrInfo("输入的内容有误，该项只能够输入数字。", 0, objTextBox)
                            End If
                        End If
                    Case SLC_新生儿出生体重, SLC_新生儿入院体重
                        If objTextBox.Text <> "" Then
                            If InStr(objTextBox.Text, ";") > 0 Then
                                arrTmp = Split(objTextBox.Text, ";")
                                For i = LBound(arrTmp) To UBound(arrTmp)
                                    If Not IsNumeric(arrTmp(i)) Then
                                        Call AddErrInfo("输入的内容有误,体重只能够是数字，如有多个新生儿，各个新生儿的体重请用“;”分隔开。", 0, objTextBox)
                                        Exit For
                                    End If
                                Next
                            Else
                                 If Not IsNumeric(objTextBox.Text) Then
                                    Call AddErrInfo("输入的内容有误,体重只能够是数字，如有多个新生儿，各个新生儿的体重请用“;”分隔开。", 0, objTextBox)
                                End If
                            End If
                        End If
                    Case SLC_Apgar
                        If objTextBox.Text <> "" Then
                            If Not IsNumeric(objTextBox.Text) Then
                                Call AddErrInfo("输入的内容有误，该项只能够输入整数。", 0, objTextBox)
                            ElseIf Val(objTextBox.Text) > 10 Then
                                Call AddErrInfo("输入的值只能在0-10 之间。", 0, objTextBox)
                            End If
                        End If
                End Select

                '必须要输入的内容检查
                Select Case objTextBox.Index
                    Case SLC_住院号
                        If objTextBox.Text = "" Then
                            Call AddErrInfo("病人的住院号不能为空，请输入病人的住院号。", 0, objTextBox)
                        Else
                            If gclsPros.InNo <> "0" Then
                                If Trim(objTextBox.Text) <> gclsPros.InNo Then
                                    Call AddErrInfo("该病人的住院号已发生改变，请检查。", 0, objTextBox)
                                End If
                            End If
                        End If
                    Case SLC_年龄
                        If objTextBox.Text = "" Then
                            Call AddErrInfo("病人的年龄不能为空，请输入病人的年龄。", 0, objTextBox)
                        Else
                            '小时为单位，不能大于30天即720小时
                            '分钟为单位，不能大于24小时即1440分钟
                            If .cboSpecificInfo(SLC_年龄).Visible Then
                                strTmp = .cboSpecificInfo(SLC_年龄).Text
                                i = decode(strTmp, "岁", 200, "月", 2400, "天", 73000, "小时", 720, "分钟", 1440, 0)
                                If Val(objTextBox.Text) > i And strTmp <> "" Then
                                    Call AddErrInfo("年龄值超过最大限制" & i & strTmp & IIf(i = 1440 Or i = 720, "，请使用合适的年龄单位。", "。"), 0, objTextBox, .cboSpecificInfo(SLC_年龄))
                                ElseIf Val(objTextBox.Text) < 0 Then
                                    Call AddErrInfo("年龄值不能为负数。", 0, objTextBox, .cboSpecificInfo(SLC_年龄))
                                End If
                            End If
                        End If
                    Case SLC_婴幼儿年龄
                        If objTextBox <> "" Then
                            '小时为单位，不能大于30天即720小时
                            '分钟为单位，不能大于24小时即1440分钟
                            If .cboSpecificInfo(SLC_婴幼儿年龄).Visible Then
                                strTmp = .cboSpecificInfo(SLC_婴幼儿年龄).Text
                                i = decode(strTmp, "月", 12, "天", 365, "小时", 720, "分钟", 1440, 0)
                                If Val(objTextBox.Text) > i Then
                                    Call AddErrInfo("婴儿年龄值超过最大限制" & i & strTmp & IIf(i = 1440 Or i = 720, "，请使用合适的年龄单位。", "。"), 0, objTextBox, .cboSpecificInfo(SLC_婴幼儿年龄))
                                ElseIf Val(objTextBox.Text) < 0 Then
                                    Call AddErrInfo("婴儿年龄值不能为负数。", 0, objTextBox, .cboSpecificInfo(SLC_婴幼儿年龄))
                                End If
                            End If
                        End If
                    Case SLC_婴幼儿年龄_DAY
                        If objTextBox.Visible Then
                            '月为单位，不能大于30天也不能小于0
                            strTmp = Trim(objTextBox.Text)
                            If strTmp = "" Then
                                If Trim(.txtSpecificInfo(SLC_婴幼儿年龄).Text) <> "" Then
                                    Call AddErrInfo("婴儿年龄不足1个月的天数不允许为空。", 0, objTextBox, .cboSpecificInfo(SLC_婴幼儿年龄))
                                End If
                            Else
                                If Trim(.txtSpecificInfo(SLC_婴幼儿年龄).Text) = "" Then
                                    Call AddErrInfo("婴儿年龄的月龄不允许为空。", 0, .txtSpecificInfo(SLC_婴幼儿年龄), .cboSpecificInfo(SLC_婴幼儿年龄))
                                ElseIf strTmp Like "0*" And Len(strTmp) > 1 Then
                                    Call AddErrInfo("婴儿年龄不足1个月的天数，不是有效的整数。", 0, objTextBox, .cboSpecificInfo(SLC_婴幼儿年龄))
                                ElseIf Val(strTmp) >= 30 Or Val(strTmp) < 0 Then
                                    Call AddErrInfo("婴儿年龄不足1个月的天数，正常取值范围为大于等于0且小于30的整数。", 0, objTextBox, .cboSpecificInfo(SLC_婴幼儿年龄))
                                End If
                            End If
                        End If
                    Case SLC_抢救次数
                        '成功次数不能超过抢救次数,病人出院情况为死亡的时候，成功次数可以等于抢救次数，因为可能病人没有抢救就死了
                        If Val(.txtSpecificInfo(SLC_成功次数).Text) > Val(objTextBox.Text) Then
                            Call AddErrInfo("成功次数不能超过抢救次数。", 0, .txtSpecificInfo(SLC_成功次数), .txtSpecificInfo(SLC_抢救次数))
                        End If

                        If objTextBox.Text <> "" Then
                            strTmp = .vsDiagXY.TextMatrix(FindDiagRow(DT_出院诊断XY), DI_出院情况)
                        End If
                    Case SLC_随诊期限
                        If Val(objTextBox.Text) <= 0 Then
                            Call AddErrInfo("请输入正确的随诊期限。", 0, objTextBox)
                        End If
                    Case SLC_住院天数
                        i = DateDiff("d", CDate(gclsPros.InTime), CDate(gclsPros.OutTime))
                        If i = 0 Then i = 1
                        If i <> Val(objTextBox.Text) Then
                            Call AddErrInfo("病人的住院天数不正确，请检查出院时间是否正确。", 0, objTextBox, .mskDateInfo(DC_出院时间))
                        End If
                End Select
            End If
        Next
        'txtAdressInfo控件相关检查
        On Error Resume Next
        For Each objTextBox In .txtAdressInfo
            If objTextBox.Enabled And Not objTextBox.Locked Then
                strTmp = decode(objTextBox.Index, ADRC_出生地点, "出生地点", ADRC_籍贯, "籍贯", ADRC_现住址, "现住址", ADRC_户口地址, "户口地址", ADRC_联系人地址, "联系人地址", ADRC_病人区域, "区域")
                '项目输入长度检查
                blnJudge = .padrInfo(objTextBox.Index).MaxLength '判断控件是否存在
                blnJudge = Err.Number = 0: Err.Clear
                If blnJudge And gclsPros.IsStructAdress Then    '需要检查地址控件的内容
                    If .padrInfo(objTextBox.Index).CheckNullValue() <> "" Then
                        Call AddErrInfo(strTmp & "的" & .padrInfo(objTextBox.Index).CheckNullValue() & "尚未输入，请检查。", 0, .padrInfo(objTextBox.Index))
                    End If
                    If .padrInfo(objTextBox.Index).MaxLength > 0 Then
                        If zlCommFun.ActualLen(.padrInfo(objTextBox.Index).Value) > .padrInfo(objTextBox.Index).MaxLength Then
                            Call AddErrInfo(strTmp & "的内容太长，请检查。(该项目最多允许 " & .padrInfo(objTextBox.Index).MaxLength & " 个字符或 " & .padrInfo(objTextBox.Index).MaxLength \ 2 & " 个汉字)", 0, .padrInfo(objTextBox.Index))
                        End If
                    End If
                Else '需要检查TextBox的内容
                    If objTextBox.MaxLength <> 0 And objTextBox.Text <> "" Then
                        If zlCommFun.ActualLen(objTextBox.Text) > objTextBox.MaxLength Then
                            Call AddErrInfo(strTmp & "的内容过长，请检查。(该项目最多允许 " & objTextBox.MaxLength & " 个字符或 " & objTextBox.MaxLength \ 2 & " 个汉字)", 0, objTextBox)
                        End If
                    End If
                     '必须要输入的内容检查
                    If objTextBox.Index = ADRC_病人区域 Then
                        If objTextBox.Text = "" Then
                            If gclsPros.FuncType = f病案首页 Then
                                Call AddErrInfo("请输入病人的" & strTmp & "。", 0, objTextBox)
                            Else
                                If gclsPros.Check区域 = 1 Then
                                    Call AddErrInfo("请输入病人的" & strTmp & "。", 0, objTextBox)
                                ElseIf gclsPros.Check区域 = 2 Then
                                    Call AddErrInfo("没有输入病人的" & strTmp & ",是否继续？", 1, objTextBox)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next

        strTmp = ""
        'cboBaseInfo检查
        For Each objCbo In .cboBaseInfo
            strTmp = ""
            If objCbo.Enabled And Not objCbo.Locked Then
                '必须要输入的内容检查
                If gclsPros.PatiType = PF_门诊 Then
                    If InStr(GetInsidePrivs(p病人信息公共部件), "基本信息调整") > 0 And objCbo.Index = BCC_付款方式 Then strTmp = "付款方式"
                Else
                    strTmp = decode(objCbo.Index, BCC_付款方式, "付款方式", BCC_国籍, "国籍", BCC_民族, "民族", BCC_职业, "职业", BCC_入院情况, "入院病情", BCC_性别, "性别", "")
                End If
                If strTmp <> "" Then
                    If objCbo.ListIndex = -1 Then
                        Call AddErrInfo("请输入病人的" & strTmp & "。", 0, objCbo)
                    End If
                End If
                '输入内容的有效性检查
                Select Case objCbo.Index
                    Case BCC_血型
                        If objCbo.Text = "" Then
                            Call AddErrInfo("必须填写该病人的血型.", 0, objCbo)
                        End If
                    Case BCC_RH
                        If objCbo.Text = "" Then
                            Call AddErrInfo("必须填写该病人的RH", 0, objCbo)
                        End If
                    Case BCC_婚姻 '15岁以下应为未婚
                        If objCbo.Text <> "" And objCbo.ListIndex <> -1 Then
                            If InStr(objCbo.Text, "未婚") = 0 And InStr(objCbo.Text, "其他") = 0 Then
                                If IsDate(.mskDateInfo(DC_出生日期).Text) Then
                                    If DateDiff("yyyy", CDate(.mskDateInfo(DC_出生日期).Text), curDate) < 15 Then
                                        Call AddErrInfo("该病人年龄小于15岁，婚姻状况应该写为未婚或其他，是否继续？", 1, objCbo)
                                    End If
                                End If
                            End If
                        End If
                    Case BCC_入院情况 '入院病情为危时需要进行抢救
                        If InStr(objCbo.Text, "危") > 0 And Val(.txtSpecificInfo(SLC_抢救次数).Text) = 0 Then
                            Call AddErrInfo("该病人入院病情为危，但没有进行抢救，是否继续？", 1, .txtSpecificInfo(SLC_抢救次数), objCbo)
                        End If
                    Case BCC_术前与术后 '填写了术前与术后，必须填写手术情况
                        If .vsOPS.TextMatrix(1, PI_手术名称) = "" And objCbo.ListIndex > 0 Then
                            Call AddErrInfo("没有填写手术情况,术前与术后只能选择""未做""。", 0, objCbo)
                        End If
                    Case BCC_出院方式 '如果出院方式是死亡，则检查出院诊断是否为死亡
                        If InStr(objCbo.Text, "死亡") > 0 Then
                            strTmp = .vsDiagXY.TextMatrix(FindDiagRow(DT_出院诊断XY), DI_出院情况)
                            If strTmp = "" And gclsPros.IsTCM Then
                                strTmp = .vsDiagZY.TextMatrix(FindDiagRow(DT_出院诊断ZY), DI_出院情况)
                            End If
                            If strTmp <> "" Then
                                If InStr(strTmp, "死亡") = 0 Then
                                    Call AddErrInfo("病人诊断情况不为死亡，出院方式为死亡。", 0, objCbo)
                                End If
                            End If
                        End If
                    Case BCC_身份证
                        If objCbo.Enabled And Not objCbo.Locked Then
                            '对身份证号进行验证
                            If objCbo.Index = BCC_身份证 Then
                                strTmp = objCbo.Text
                                If strTmp <> "" Then
                                    If Trim(zlCommFun.GetNeedName(.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                                        If zlCommFun.ActualLen(strTmp) = Len(strTmp) Then
                                            If gclsPros.IsMaskID Then strTmp = objCbo.Tag
                                                '初始化病人信息接口
                                                If gobjPatient Is Nothing Then
                                                    On Error Resume Next
                                                    Set gobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
                                                    Err.Clear: On Error GoTo errH
                                                    Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
                                                End If
                                                If gobjPatient Is Nothing Then
                                                    Call AddErrInfo("创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败，不能进行病人身份证信息检查！是否继续？", 1, objCbo)
                                                End If
                                                If gobjPatient.CheckPatiIdcard(strTmp, strBirthday, strAge, strSex, strErrIfno, CDate(gclsPros.InTime)) Then '身份证合法则检查是否匹配
                                                    
                                                                                                        If Val(zlDatabase.GetPara(279, 100)) = 1 Then
                                                        strSql = "select 1 from 病人信息 a where a.身份证号=[1] and a.病人id<>[2] and rownum<2"
                                                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, strTmp, gclsPros.病人ID)
                                                        If Not rsTmp.EOF Then Call AddErrInfo("该身份证号已经建档，同一身份证只能对应一个建档病人!", 0, objCbo)
                                                    End If

                                                    strBaseInfo = "": Set objTmp = Nothing
                                                    If Not Trim(.txtSpecificInfo(SLC_年龄).Text) Like "约*" Or Trim(.txtSpecificInfo(SLC_年龄).Text) = "不详" Then
                                                        If Format(strBirthday, "yyyy-MM-dd") <> Format(.mskDateInfo(DC_出生日期).Text, "yyyy-MM-dd") Then
                                                            strBaseInfo = "出生日期"
                                                            Set objTmp = objCbo
                                                        End If
                                                                If Format(.mskDateInfo(DC_出生日期).Text, "HH:MM") <> "00:00" Then
                                                                        strBirthday = strBirthday & " " & Format(.mskDateInfo(DC_出生日期).Text, "HH:MM")
                                                                End If
                                                        If gclsPros.Sex <> strSex Then
                                                            strBaseInfo = strBaseInfo & IIf(strBaseInfo <> "", "、", "") & "性别"
                                                            Set objTmp = .cboBaseInfo(BCC_性别)
                                                        End If
                                                        str年龄 = .txtSpecificInfo(SLC_年龄).Text & IIf(.cboSpecificInfo(SLC_年龄).Visible, .cboSpecificInfo(SLC_年龄).Text, "")
                                                        If strAge <> str年龄 Then
                                                            strBaseInfo = strBaseInfo & IIf(strBaseInfo <> "", "、", "") & "年龄"
                                                            Set objTmp = .txtSpecificInfo(SLC_年龄)
                                                                        If Trim(str年龄) Like "*小时*分钟" Or Trim(str年龄) Like "*分钟" Or Trim(str年龄) Like "*天*小时" Or Trim(str年龄) Like "*小时" Then
                                                                        strAge = .txtSpecificInfo(SLC_年龄).Text & IIf(.cboSpecificInfo(SLC_年龄).Visible, .cboSpecificInfo(SLC_年龄).Text, "")
                                                                        End If
                                                        End If
                                                    End If
                                                    If strBaseInfo <> "" Then
                                                        If InStr(GetInsidePrivs(p病人信息公共部件), "基本信息调整") = 0 Or gclsPros.FuncType = f病案首页 Then
                                                            Call AddErrInfo("身份证号码获取的" & strBaseInfo & "与当前界面的" & strBaseInfo & "不相符，是否继续？", 1, objTmp)
                                                        Else
                                                            Call AddErrInfo("身份证号码获取的" & strBaseInfo & "与当前界面的" & strBaseInfo & "不相符，是否继续？继续则将自动更新界面上的" & strBaseInfo & "。", 1, objTmp)
                                                            blnBaseInfo = True
                                                        End If
                                                    End If
                                            Else '身份证不合法则退出
                                                Call AddErrInfo(strErrIfno, 0, objCbo)
                                            End If
                                        ElseIf zlCommFun.ActualLen(strTmp) > 18 Then
                                            Call AddErrInfo("身份证号不能超过9个汉字或18个英文字符的长度，请检查。", 0, objCbo)
                                        End If
                                    End If
                                ElseIf gclsPros.FuncType = f病案首页 Then
                                    Call AddErrInfo("身份证号码没有输入，是否继续？", 1, objCbo)
                                End If
                            End If
                        End If
                End Select
            End If
        Next
        If gclsPros.PatiType <> PF_门诊 Then
            'cboManInfo检查
            strTmp = ""
            For Each objCbo In .cboManInfo
                If Not objCbo.Locked Then
                    Select Case objCbo.Index
                        Case MC_科主任
                            If objCbo.Text = "" Then strTmp = strTmp & ";科主任"
                        Case MC_主任或副主任
                            If objCbo.Text = "" Then strTmp = strTmp & ";主任医师"
                        Case MC_主治医师
                            If objCbo.Text = "" Then strTmp = strTmp & ";主治医师"
                        Case MC_住院医师
                            If objCbo.Text = "" Then strTmp = strTmp & ";住院医师"
                        Case MC_编目员
                            If gclsPros.FuncType = f病案首页 Then
                                If objCbo.Text = "" Then
                                    Call AddErrInfo("请输入编目员。", 0, objCbo)
                                End If
                            End If
                    End Select
                End If
            Next
            If UBound(Split(strTmp, ";")) = 4 Then '必须要输入的内容检查
                Call AddErrInfo("请在科主任、主任医师、主治医师和住院医师之间至少选择一位。", 0, .cboManInfo(MC_科主任), .cboManInfo(MC_主任或副主任), .cboManInfo(MC_主治医师), .cboManInfo(MC_住院医师))
            End If
            strTmp = ""
        End If
        'mskDateInfo检查
        For Each objMSK In .mskDateInfo
            If objMSK.Enabled Then
                blnJudge = objMSK.Text = Replace(objMSK.Mask, "#", "_")
                '必须要输入的内容检查
                Select Case objMSK.Index
                    Case DC_确诊日期
                        If Not blnJudge Then
                            If Not IsDate(objMSK.Text) Then
                                Call AddErrInfo("确诊日期不是有效的日期格式。", 0, objMSK)
                            End If
                            If gclsPros.FuncType = f病案首页 Then
                                If Not Between(Format(objMSK.Text, "yyyy-MM-dd"), Format(gclsPros.InTime, "yyyy-MM-dd"), _
                                    Format(IIf(gclsPros.OutTime = "", curDate, gclsPros.OutTime), "yyyy-MM-dd")) Then
                                    Call AddErrInfo("确诊日期必须在入院时间和出院时间之间。", 0, objMSK)
                                End If
                            Else
                                If objMSK.Mask = "####-##-##" Then
                                    If Not Between(Format(objMSK.Text, "yyyy-MM-dd"), Format(gclsPros.InTime, "yyyy-MM-dd"), _
                                        Format(IIf(gclsPros.OutTime = "", curDate, gclsPros.OutTime), "yyyy-MM-dd")) Then
                                        Call AddErrInfo("确诊日期必须在入院时间和出院时间之间。", 0, objMSK)
                                    End If
                                Else
                                    If Not Between(Format(objMSK.Text, "yyyy-MM-dd hh:mm"), Format(gclsPros.InTime, "yyyy-MM-dd hh:mm"), _
                                        Format(IIf(gclsPros.OutTime = "", curDate, gclsPros.OutTime), "yyyy-MM-dd hh:mm")) Then
                                        Call AddErrInfo("确诊日期必须在入院时间和出院时间之间。", 0, objMSK)
                                    End If
                                End If
                            End If
                        ElseIf .chkInfo(CHK_是否确诊).Value = 1 Then
                            Call AddErrInfo("请输入确诊日期。", 0, objMSK)
                        End If
                    Case DC_出生日期
                        If blnJudge Then
                            Call AddErrInfo("请输入病人的出生日期。", 0, objMSK)
                        End If
                        If Not IsDate(objMSK.Text) Then
                            Call AddErrInfo("出生日期不是有效的日期格式。", 0, objMSK)
                        End If
                        If Format(objMSK.Text, "yyyy-MM-dd hh:mm") > Format(gclsPros.InTime, "yyyy-MM-dd hh:mm") Then
                            Call AddErrInfo("出生日期在入院时间之后。", 0, objMSK, .mskDateInfo(DC_入院时间))
                        End If
                    Case DC_发病日期
                        If Not blnJudge Then
                            If Not IsDate(objMSK.Text) Then
                                Call AddErrInfo("请输入正确的发病日期。", 0, objMSK)
                            Else
                                If Not IsDate(.mskDateInfo(DC_发病时间).Text) And .mskDateInfo(DC_发病时间).Text <> "__:__" Then
                                    Call AddErrInfo("请输入正确的发病时间。", 0, .mskDateInfo(DC_发病时间))
                                End If
                                strTmp = IIf(IsDate(.mskDateInfo(DC_发病时间).Text), " " & .mskDateInfo(DC_发病时间).Text, "")
                                If CDate(objMSK.Text & strTmp) >= CDate(Format(curDate, GetFormat(objMSK.Tag) & IIf(strTmp = "", "", " HH:mm"))) Then
                                    Call AddErrInfo("发病时间应该早于当前时间。", 0, objMSK)
                                End If
                            End If
                        End If
                    Case DC_死亡时间
                        If Not blnJudge Then
                            If Not IsDate(objMSK.Text) Then
                                Call AddErrInfo("死亡时间不是有效的日期格式。", 0, objMSK)
                            End If
                            If Format(objMSK.Text, "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                Call AddErrInfo("死亡时间应比入院时间晚。", 0, objMSK)
                            End If
                        End If
                    Case DC_编目日期
                        If blnJudge Then
                            Call AddErrInfo("请输入编目日期。", 0, objMSK)
                        End If
                        If Not IsDate(objMSK.Text) Then
                            Call AddErrInfo("编目日期不是有效的日期格式。", 0, objMSK)
                        End If
                    Case DC_质控日期
                        If Not blnJudge Then
                            If Not IsDate(objMSK.Text) Then
                                Call AddErrInfo("质控日期不是有效的日期格式。", 0, objMSK)
                            ElseIf Format(objMSK.Text, "yyyy-MM-dd") < Format(gclsPros.InTime, "yyyy-MM-dd") Then
                                Call AddErrInfo("质控日期不能小于入院日期。", 0, objMSK)
                            End If
                        End If
                End Select
            End If
        Next
        '出院病人清单检查
        If gclsPros.FuncType = f病案首页 And gclsPros.InputOutList Then
            strSql = "Select A.姓名, A.日期, B.名称, B.Id" & vbNewLine & _
                    "From 出院病人清单 A, 部门表 B" & vbNewLine & _
                    "Where A.科室id = B.Id And A.住院号 = [1]" & vbNewLine & _
                    "Order By A.日期 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, .txtSpecificInfo(SLC_住院号).Text)
            If rsTmp.RecordCount = 0 Then
                If gclsPros.OpenMode = EM_编辑 Then
                    Call AddErrInfo("住院号为" & .txtSpecificInfo(SLC_住院号).Text & "的病人还没有被录入住院日报的出院病人清单中，是否继续？", 1, .txtSpecificInfo(SLC_住院号))
                Else
                    Call AddErrInfo("住院号为" & .txtSpecificInfo(SLC_住院号).Text & "的病人还没有被录入住院日报的出院病人清单中。", 0, .txtSpecificInfo(SLC_住院号))
                End If
            Else
                If Not IsNull(rsTmp!日期) Then
                    If zlCommFun.TruncateDate(gclsPros.OutTime) <> zlCommFun.TruncateDate(rsTmp!日期 & "") Then
                        Call AddErrInfo("在住院日报中该病人的出院日期(" & Format(rsTmp!日期 & "", "yyyy-MM-dd") & _
                                ")与当前填写的不符，是否继续？", 1, .mskDateInfo(DC_出院时间))
                    End If
                End If
                If gclsPros.出院科室ID <> Val(rsTmp!ID) And gclsPros.出院科室ID <> 0 Then
                    Call AddErrInfo("在住院日报中该病人的出院科室(" & .txtInfo(GC_出院科室).Text & ")与当前填写的不符，是否继续？", 1, .txtInfo(GC_出院科室))
                End If
            End If
        End If
    End With

    If Not gclsPros.Is护士站 Then
        '表格的检查
        '----------------------------------------------------------------------------------------
        If gclsPros.FuncType = f病案首页 Then
            Set vsTmp = gclsPros.CurrentForm.vsTransfer
            With vsTmp
                For i = .FixedCols To .Cols - 1
                    If .TextMatrix(DR_转科科室, i) <> "" Then
                        If .TextMatrix(DR_转科科室, i) = .TextMatrix(DR_转科科室, i - 1) Then
                            .Row = DR_转科科室: .Col = i
                            Call AddErrInfo("第" & i & "列与第" & i - 1 & "列转入科室相同,请检查!如要插入其他科室请按Insert键。", 0, vsTmp)
                        End If
                    ElseIf i <> .Cols - 1 Then '必须以及输入转科科室
                        If .TextMatrix(DR_转科科室, i + 1) <> "" Then
                            .Row = DR_转科科室: .Col = i
                            Call AddErrInfo("第" & i & "列没有转入科室,但第" & i + 1 & "列存在转入科室,请检查!如要删除该行请按Delete键。", 0, vsTmp)
                        End If
                    End If
                    If .TextMatrix(DR_转科科室, i) = "" And Trim(.TextMatrix(DR_转科科室, i)) <> "" Then
                        .Row = DR_转科科室: .Col = i
                        Call AddErrInfo("第" & i & "列有转科时间但未填转入科室!如要删除该行请按Delete键。", 0, vsTmp)
                    End If
                Next
            End With
        End If

        gclsPros.Have手术 = False
        If gclsPros.PatiType <> PF_门诊 Then
            Set vsTmp = gclsPros.CurrentForm.vsOPS
            With vsTmp
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, PI_手术名称)) <> "" Then
                        gclsPros.Have手术 = True
                    End If
                Next
            End With
        End If

        Call CheckDiagData(curDate)
        '过敏药物表格检查
        blnJudge = True
        Set vsTmp = gclsPros.CurrentForm.vsAller
        With vsTmp
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, AI_过敏药物)) <> "" Then
                    If blnJudge Then
                        If gclsPros.CurrentForm.chkInfo(CHK_无过敏记录).Value = 1 And Trim(.TextMatrix(i, AI_过敏药物)) <> "—" Then
                            Call AddErrInfo("该病人存在过敏记录，不能勾选无过敏记录。", 0, gclsPros.CurrentForm.chkInfo(CHK_无过敏记录))
                        End If
                    End If
                    blnJudge = False
                    If zlCommFun.ActualLen(.TextMatrix(i, AI_过敏药物)) > 60 Then
                        .Row = i: .Col = AI_过敏药物
                        Call AddErrInfo("过敏药物名太长，只允许60个字符或30个汉字。", 0, vsTmp)
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(i, AI_过敏反应)) > 100 Then
                        .Row = i: .Col = AI_过敏反应
                        Call AddErrInfo("过敏反应内容太长，只允许100个字符或50个汉字。", 0, vsTmp)
                    End If
                    For j = i + 1 To .Rows - 1
                        If Trim(.TextMatrix(j, AI_过敏药物)) <> "" And Format(.TextMatrix(i, AI_过敏时间), "yyyy-mm-dd") = Format(.TextMatrix(j, AI_过敏时间), "yyyy-mm-dd") Then
                            blnDateIsNull = False
                            If .TextMatrix(j, AI_过敏时间) = "" Then blnDateIsNull = True
                            If .TextMatrix(j, AI_过敏药物) = .TextMatrix(i, AI_过敏药物) Then
                                .Row = i: .Col = AI_过敏药物
                                Call AddErrInfo("发现" & IIf(blnDateIsNull, "存在过敏时间为空的相同的过敏药物记录。", Format(.TextMatrix(j, AI_过敏时间), "yyyy年mm月dd日") & "内存在相同的过敏药物记录。"), 0, vsTmp)
                            ElseIf Val(.TextMatrix(i, AI_药物ID)) <> 0 And .TextMatrix(i, AI_药物ID) = .TextMatrix(j, AI_药物ID) Then
                                .Row = i: .Col = AI_过敏药物
                                Call AddErrInfo("发现" & IIf(blnDateIsNull, "存在过敏时间为空的相同的过敏药物记录。", Format(.TextMatrix(j, AI_过敏时间), "yyyy年mm月dd日") & "内存在相同的过敏药物记录。"), 0, vsTmp)
                            ElseIf .TextMatrix(i, AI_过敏源编码) <> "" And .TextMatrix(i, AI_过敏源编码) = .TextMatrix(j, AI_过敏源编码) Then
                                .Row = i: .Col = AI_过敏药物
                                Call AddErrInfo("发现" & IIf(blnDateIsNull, "存在过敏时间为空的相同的过敏药物记录。", Format(.TextMatrix(j, AI_过敏时间), "yyyy年mm月dd日") & "内存在相同的过敏药物记录。"), 0, vsTmp)
                            End If
                        End If
                    Next
                End If
            Next
        End With
        If blnJudge Then
            If gclsPros.CurrentForm.chkInfo(CHK_无过敏记录).Value = 0 Then
                Call AddErrInfo("该病人不存在过敏记录，但没有勾选无过敏记录，是否继续？", 1, gclsPros.CurrentForm.chkInfo(CHK_无过敏记录))
            End If
        End If
        blnJudge = False

        If gclsPros.PatiType <> PF_门诊 Then
            Set vsTmp = gclsPros.CurrentForm.vsOPS
            With vsTmp
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, PI_手术名称)) <> "" Then
                        If .TextMatrix(i, PI_手术编码) = "" And gclsPros.FuncType = f病案首页 Then
                            .Row = i: .Col = PI_手术编码
                            Call AddErrInfo("请输入手术编码。", 0, vsTmp)
                        End If
                        If Not IsDate(.TextMatrix(i, PI_手术日期)) Then
                            .Row = i: .Col = PI_手术日期
                            Call AddErrInfo("手术日期输入不正确。", 0, vsTmp)
                        ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, PI_手术日期), "yyyy-MM-dd") > Format(gclsPros.OutTime, "yyyy-MM-dd") Or _
                            Format(.TextMatrix(i, PI_手术日期), "yyyy-MM-dd") < Format(gclsPros.InTime, "yyyy-MM-dd") Then
                            .Row = i: .Col = PI_手术日期    '手术日期没有精确到时间
                            Call AddErrInfo("手术日期不在入出院日期范围内。", 0, vsTmp)
                        End If

                        If gclsPros.UseOPSEndTime Then
                            If Not IsDate(.TextMatrix(i, PI_结束日期)) Then
                                .Row = i: .Col = PI_结束日期
                                Call AddErrInfo("手术结束时间输入不正确。", 0, vsTmp)
                            ElseIf Format(.TextMatrix(i, PI_结束日期), "yyyy-MM-dd HH:mm") < Format(.TextMatrix(i, PI_手术日期), "yyyy-MM-dd HH:mm") Then
                                .Row = i: .Col = PI_结束日期
                                Call AddErrInfo("手术结束时间必须大于手术开始时间。", 0, vsTmp)
                            ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, PI_结束日期), "yyyy-MM-dd HH:mm") > Format(gclsPros.OutTime, "yyyy-MM-dd HH:mm") Or _
                                Format(.TextMatrix(i, PI_结束日期), "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                .Row = i: .Col = PI_结束日期
                                Call AddErrInfo("手术结束时间不在入出院日期范围内。", 0, vsTmp)
                            End If
                        End If
                        If gclsPros.MedPageSandard = ST_四川省标准 Then
                            If Not IsDate(.TextMatrix(i, PI_麻醉开始时间)) Then
                                If .TextMatrix(i, PI_麻醉开始时间) <> "" Then
                                    .Row = i: .Col = PI_麻醉开始时间
                                    Call AddErrInfo("麻醉开始时间输入不正确。", 0, vsTmp)
                                End If
                            ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, PI_麻醉开始时间), "yyyy-MM-dd HH:mm") > Format(gclsPros.OutTime, "yyyy-MM-dd HH:mm") Or _
                                Format(.TextMatrix(i, PI_麻醉开始时间), "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                .Row = i: .Col = PI_结束日期
                                Call AddErrInfo("麻醉开始时间不在入出院日期范围内。", 0, vsTmp)
                            End If

                            If Not IsDate(.TextMatrix(i, PI_抗菌用药时间)) And .TextMatrix(i, PI_抗菌用药时间) <> "" Then
                                .Row = i: .Col = PI_抗菌用药时间
                                Call AddErrInfo("抗菌用药时间输入不正确。", 0, vsTmp)
                            End If
                            strSql = "Select 准备天数 From 病人手麻记录 Where Rownum = 1"
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取字段长度")
                            If Len(Trim(.TextMatrix(i, PI_准备天数))) > 3 Then
                                Call AddErrInfo("准备天数输入不正确,输入的值超过了最大限制长度。", 0, vsTmp)
                            End If
                        End If
                        If gclsPros.MedPageSandard = ST_云南省标准 Then
                            If Len(Trim(.TextMatrix(i, PI_抗菌药天数))) > 5 Then
                                Call AddErrInfo("抗菌用药天数输入不正确,输入的值超过了最大限制长度。", 0, vsTmp)
                            End If
                        End If
                        If zlCommFun.ActualLen(.TextMatrix(i, PI_手术名称)) > 300 Then
                            .Row = i: .Col = PI_手术名称
                            Call AddErrInfo("手术名称内容太长，只允许300个字符或150个汉字。", 0, vsTmp)
                        End If

                        If .ColHidden(PI_助产护士) Then
                            If .TextMatrix(i, PI_主刀医师) = "" Then
                                .Row = i: .Col = PI_主刀医师
                                Call AddErrInfo("请输入主刀医师。", 0, vsTmp)
                            End If
                        Else
                            If .TextMatrix(i, PI_主刀医师) = "" And .TextMatrix(i, PI_助产护士) = "" Then
                                .Row = i: .Col = PI_主刀医师
                                Call AddErrInfo("请输入主刀医师或助产护士。", 0, vsTmp)
                            End If
                        End If
                        For j = i + 1 To .Rows - 1
                            If Trim(.TextMatrix(j, PI_手术名称)) <> "" Then
                            
                                '处理手术结束日期在界面表格中不显示时，结束日期就等于手术日期
                                If gclsPros.MedPageSandard = ST_卫生部标准 Or gclsPros.MedPageSandard = ST_湖南省标准 Then
                                    If Not gclsPros.UseOPSEndTime Then
                                        .TextMatrix(i, PI_结束日期) = .TextMatrix(i, PI_手术日期)
                                        .TextMatrix(j, PI_结束日期) = .TextMatrix(j, PI_手术日期)
                                    End If
                                ElseIf gclsPros.MedPageSandard = ST_云南省标准 Then
                                    .TextMatrix(i, PI_结束日期) = .TextMatrix(i, PI_手术日期)
                                    .TextMatrix(j, PI_结束日期) = .TextMatrix(j, PI_手术日期)
                                End If
                            
                                If gclsPros.MedPageSandard = ST_卫生部标准 Then
                                    If .TextMatrix(i, PI_手术日期) & "|" & .TextMatrix(i, PI_结束日期) & "|" & .TextMatrix(i, PI_手术编码) & "|" & .TextMatrix(i, PI_手术名称) & "|" & .TextMatrix(i, PI_手术操作ID) & "|" & .TextMatrix(i, PI_诊疗项目ID) & "|" & .TextMatrix(i, PI_切口部位) = .TextMatrix(j, PI_手术日期) & "|" & .TextMatrix(j, PI_结束日期) & "|" & .TextMatrix(j, PI_手术编码) & "|" & .Cell(flexcpData, j, PI_手术名称) & "|" & .TextMatrix(j, PI_手术操作ID) & "|" & .TextMatrix(j, PI_诊疗项目ID) & "|" & .TextMatrix(j, PI_切口部位) Then
                                        .Row = j: .Col = PI_手术编码
                                        Call AddErrInfo("发现存在两条手术日期、手术操作和切口部位都相同的手术记录。", 0, vsTmp)
                                    End If
                                Else
                                    If .TextMatrix(i, PI_手术日期) & "|" & .TextMatrix(i, PI_结束日期) & "|" & .TextMatrix(i, PI_手术编码) & "|" & .TextMatrix(i, PI_手术名称) & "|" & .TextMatrix(i, PI_手术操作ID) & "|" & .TextMatrix(i, PI_诊疗项目ID) = .TextMatrix(j, PI_手术日期) & "|" & .TextMatrix(j, PI_结束日期) & "|" & .TextMatrix(j, PI_手术编码) & "|" & .Cell(flexcpData, j, PI_手术名称) & "|" & .TextMatrix(j, PI_手术操作ID) & "|" & .TextMatrix(j, PI_诊疗项目ID) Then
                                        .Row = j: .Col = PI_手术编码
                                        Call AddErrInfo("发现存在两条手术日期与手术操作都相同的手术记录。", 0, vsTmp)
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
            End With
        End If

        If gclsPros.FuncType = f病案首页 Then
            grsBabyInfo.Filter = "分娩方式<>'正常分娩'"
            If grsBabyInfo.RecordCount > 0 Then
                If Not gclsPros.Have手术 Then
                    Call AddErrInfo("胎儿的分娩方式存在不是正常分娩，但未输入相关的手术情况，是否继续？", 1, gclsPros.CurrentForm.vsOPS)
                End If
            Else
                grsBabyInfo.Filter = "分娩方式='正常分娩'"
                If grsBabyInfo.RecordCount > 0 Then
                    If gclsPros.Have手术 Then
                        Call AddErrInfo("胎儿的分娩方式是正常分娩，不能输入相关的手术情况，是否继续？", 1, gclsPros.CurrentForm.vsOPS)
                    End If
                End If
            End If

            Set vsTmp = gclsPros.CurrentForm.vsFees
            With vsTmp
                i = 3
                Do
                    On Error Resume Next
                    sgnTmp = Fix(Val(.TextMatrix(i \ 3, (i Mod 3) * 2 + 1)))
                    If Err.Number <> 0 Or Len(sgnTmp) > 11 Then
                        .Row = i \ 3: .Col = (i Mod 3) * 2 + 1
                        Call AddErrInfo("费用金额数值太大。", 0, vsTmp)
                        Err.Clear: On Error GoTo 0
                    End If
                    If Val(.TextMatrix(i \ 3, (i Mod 3) * 2 + 1)) <> 0 Then
                        If .TextMatrix(i \ 3, (i Mod 3) * 2) Like "*手术*" And Not .TextMatrix(i \ 3, (i Mod 3) * 2) Like "*非手术*" And Not gclsPros.Have手术 Then
                            gclsPros.CurrentForm.vsOPS.Row = gclsPros.CurrentForm.vsOPS.FixedRows: gclsPros.CurrentForm.vsOPS.Col = PI_手术日期
                            Call AddErrInfo("该病人住院费用中含有手术费，但没有录入手术信息，是否继续？", 1, gclsPros.CurrentForm.vsOPS)
                        ElseIf .TextMatrix(i \ 3, (i Mod 3) * 2) Like "*输血*" Or .TextMatrix(i \ 3, (i Mod 3) * 2) Like "*血费*" Then
                            If gclsPros.CurrentForm.cboBaseInfo(BCC_血型).Text = "" And gclsPros.CurrentForm.cboBaseInfo(BCC_RH).Text = "" And _
                                gclsPros.CurrentForm.txtSpecificInfo(SLC_输红细胞).Text = "" And gclsPros.CurrentForm.txtSpecificInfo(SLC_输血小板).Text = "" And _
                                gclsPros.CurrentForm.txtSpecificInfo(SLC_输血浆).Text = "" And gclsPros.CurrentForm.txtSpecificInfo(SLC_输全血).Text = "" And _
                                gclsPros.CurrentForm.txtInfo(GC_输其他).Text = "" Then
                                Call AddErrInfo("该病人存在输血费，请选择血型、Rh或者输入红细胞、" & vbCrLf & "输血小板、输血浆、输全血、输其他中相应的项。", 0, gclsPros.CurrentForm.cboBaseInfo(BCC_血型), gclsPros.CurrentForm.cboBaseInfo(BCC_RH))
                            End If
                        End If
                    End If
                    j = i + 1
                    If j <= .Rows * 3 - 1 Then
                        Do
                            If .TextMatrix(j \ 3, (j Mod 3) * 2) <> "" Then
                                If GetTextByDot(.TextMatrix(i \ 3, (i Mod 3) * 2)) = GetTextByDot(.TextMatrix(j \ 3, (j Mod 3) * 2)) Then
                                    If Not gclsPros.SameName Then
                                        .Row = j \ 3: .Col = (j Mod 3) * 2
                                        Call AddErrInfo("分类费用表中“" & GetTextByDot(.TextMatrix(i \ 3, (i Mod 3) * 2)) & "”费用输入了多次。", 0, vsTmp)
                                    Else '合并重名费用
                                        .TextMatrix(i \ 3, (i Mod 3) * 2) = Format(Val(.TextMatrix(i \ 3, (i Mod 3) * 2)) + Val(.TextMatrix(j \ 3, (j Mod 3) * 2)), gclsPros.FreeFormat)
                                        Call AddOrDelFreeCols(vsTmp, .TextMatrix(j \ 3, (j Mod 3) * 2), .TextMatrix(j \ 3, (j Mod 3) * 2 + 1), False)
                                    End If
                                End If
                            End If
                            If j < .Rows * 3 - 1 Then
                                j = j + 1: blnDoEx = True
                            Else
                                blnDoEx = False
                            End If
                        Loop While blnDoEx
                    End If
                    If i < .Rows * 3 - 1 Then
                        i = i + 1: blnDo = True
                    Else
                        blnDo = False
                    End If
                Loop While blnDo
            End With
        End If
        If gclsPros.PatiType <> PF_门诊 Then
            Set vsTmp = gclsPros.CurrentForm.vsKSS
            With vsTmp
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, KI_抗菌药物名) <> "" Then
                        If Trim(.TextMatrix(i - 1, KI_抗菌药物名)) = "" Then
                            .Row = i - 1: .Col = KI_抗菌药物名
                            Call AddErrInfo("请依次输入抗菌药物内容。", 0, vsTmp)
                        End If
                        If (Len(.TextMatrix(i, KI_使用天数)) > 18 Or Val(.TextMatrix(i, KI_使用天数)) = 0) And Trim(.TextMatrix(i, KI_使用天数)) <> "" Then
                            .Row = i: .Col = KI_使用天数
                            Call AddErrInfo("请填写十八位数以内的数字天数。", 0, vsTmp)
                        End If
                        If zlCommFun.ActualLen(.TextMatrix(i, KI_用药目的)) > 200 And Trim(.TextMatrix(i, KI_用药目的)) <> "" Then
                            .Row = i: .Col = KI_用药目的
                            Call AddErrInfo("请填写100个汉字以内的用药目的。", 0, vsTmp)
                        End If

                        For j = .FixedRows To i - 1
                            If Trim(.TextMatrix(j, KI_抗菌药物名)) = Trim(.TextMatrix(i, KI_抗菌药物名)) And Trim(.TextMatrix(j, KI_用药目的)) = Trim(.TextMatrix(i, KI_用药目的)) And Trim(.TextMatrix(j, KI_使用阶段)) = Trim(.TextMatrix(i, KI_使用阶段)) Then
                                .Row = j: .Col = KI_抗菌药物名
                                Call AddErrInfo("发现存在两行相同的抗菌药物信息。", 0, vsTmp)
                            End If
                        Next
                    End If
                Next
            End With
            If gclsPros.MedPageSandard <> ST_四川省标准 Then
                Set vsTmp = gclsPros.CurrentForm.vsTSJC
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If Trim(.TextMatrix(i, 1)) <> "" Then
                            If i > .FixedRows Then
                                If Trim(.TextMatrix(i - 1, 1)) = "" Then
                                    .Row = i - 1: .Col = 1
                                    Call AddErrInfo("请依次输入特殊检查内容。", 0, vsTmp)
                                End If
                            End If

                            For j = .FixedRows To i - 1
                                If Trim(.TextMatrix(j, 1)) = Trim(.TextMatrix(i, 1)) Then
                                    .Row = j: .Col = 1
                                    Call AddErrInfo("发现存在两行相同的特殊检查信息。", 0, vsTmp)
                                End If
                            Next
                        End If
                    Next
                End With
            End If
            If gclsPros.MedPageSandard = ST_卫生部标准 Or gclsPros.MedPageSandard = ST_四川省标准 Then
                Set vsTmp = gclsPros.CurrentForm.vsFlxAddICU
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If Trim(.TextMatrix(i, UI_监护室名称)) <> "" Then
                            If zlCommFun.ActualLen(.TextMatrix(i, UI_监护室名称)) > 100 Then
                                .Row = i: .Col = UI_监护室名称
                                Call AddErrInfo("重症监护室名称输入内容太长，只允许100个字符或50个汉字。", 0, vsTmp)
                            End If
                            If Trim(.TextMatrix(i, UI_进入时间)) <> "____-__-__ __:__" Then
                                If Not IsDate(.TextMatrix(i, UI_进入时间)) Then
                                     .Row = i: .Col = UI_进入时间
                                    If gclsPros.MedPageSandard = ST_卫生部标准 Then
                                         Call AddErrInfo("进入时间输入不正确。", 0, vsTmp)
                                    ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                                         Call AddErrInfo("入住时间输入不正确。", 0, vsTmp)
                                    End If
                                End If
                            ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, UI_进入时间), "yyyy-MM-dd HH:mm") > Format(gclsPros.OutTime, "yyyy-MM-dd HH:mm") Or _
                                     Format(.TextMatrix(i, UI_进入时间), "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                    .Row = i: .Col = UI_进入时间
                                    If gclsPros.MedPageSandard = ST_卫生部标准 Then
                                        Call AddErrInfo("进入时间不在入出院日期范围内。", 0, vsTmp)
                                    ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                                        Call AddErrInfo("入住时间不在入出院日期范围内。", 0, vsTmp)
                                    End If
                            End If
                            If Trim(.TextMatrix(i, UI_退出时间)) <> "____-__-__ __:__" Then
                                If Not IsDate(.TextMatrix(i, UI_退出时间)) Then
                                    .Row = i: .Col = UI_退出时间
                                    If gclsPros.MedPageSandard = ST_卫生部标准 Then
                                        Call AddErrInfo("退出时间输入不正确。", 0, vsTmp)
                                    ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                                        Call AddErrInfo("转出时间输入不正确。", 0, vsTmp)
                                    End If
                                End If
                            ElseIf gclsPros.OutTime <> "" And Format(.TextMatrix(i, UI_退出时间), "yyyy-MM-dd HH:mm") > Format(gclsPros.OutTime, "yyyy-MM-dd HH:mm") Or _
                                     Format(.TextMatrix(i, UI_退出时间), "yyyy-MM-dd HH:mm") < Format(gclsPros.InTime, "yyyy-MM-dd HH:mm") Then
                                    .Row = i: .Col = UI_退出时间
                                    If gclsPros.MedPageSandard = ST_卫生部标准 Then
                                        Call AddErrInfo("退出时间不在入出院日期范围内。", 0, vsTmp)
                                    ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                                        Call AddErrInfo("转出时间不在入出院日期范围内。", 0, vsTmp)
                                    End If
                            End If
                            If Trim(.TextMatrix(i, UI_退出时间)) <> "" And Trim(.TextMatrix(i, UI_进入时间)) <> "" And CDate(Trim(.TextMatrix(i, UI_退出时间))) < CDate(Trim(.TextMatrix(i, UI_进入时间))) Then
                                .Row = i: .Col = UI_进入时间
                                If gclsPros.MedPageSandard = ST_卫生部标准 Then
                                    Call AddErrInfo("进入ICU的时间必须小于退出ICU的时间。", 0, vsTmp)
                                ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                                    Call AddErrInfo("入住ICU的时间必须小于转出ICU的时间。", 0, vsTmp)
                                End If
                            End If
                        End If
                    Next
                End With
            End If

            '重症监护器械、医院感染、标本情况检查，仅四川版有
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, TI_ICU类型) <> "" And .TextMatrix(i, TI_器械及导管) <> "" Then
                            j = Val(.Cell(flexcpData, i, TI_ICU类型))
                            str入住时间 = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(j, UI_进入时间))
                            str转出时间 = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(j, UI_退出时间))
                            If zlCommFun.ActualLen(.TextMatrix(i, TI_ICU类型)) > 50 Then
                                .Row = i: .Col = TI_ICU类型
                                Call AddErrInfo("ICU类型的内容的最多只允许50个字符/25个汉字。", 0, vsTmp)
                            End If
                            If Not IsDate(.TextMatrix(i, TI_开始时间)) Then
                                .Row = i: .Col = TI_开始时间
                                Call AddErrInfo("请输入正确的开始使用时间。", 0, vsTmp)
                            Else
                                If IsDate(str入住时间) Then
                                    If CDate(.TextMatrix(i, TI_开始时间)) < CDate(str入住时间) Then
                                        .Row = i: .Col = TI_开始时间
                                        Call AddErrInfo("开始使用时间小于了重症监护病房的入住时间,请检查", 0, vsTmp)
                                    End If
                                End If
                            End If
                            If Not IsDate(.TextMatrix(i, TI_结束时间)) Then
                                .Row = i: .Col = TI_结束时间
                                Call AddErrInfo("请输入正确的开始使用时间。", 0, vsTmp)
                            Else
                                If IsDate(str转出时间) Then
                                    If CDate(.TextMatrix(i, TI_结束时间)) > CDate(str转出时间) Then
                                        .Row = i: .Col = TI_结束时间
                                        Call AddErrInfo("结束使用时间大于了重症监护病房的转出时间,请检查", 0, vsTmp)
                                    End If
                                End If
                            End If
                            If IsDate(.TextMatrix(i, TI_开始时间)) And IsDate(.TextMatrix(i, TI_结束时间)) Then
                                If CDate(.TextMatrix(i, TI_开始时间)) > CDate(.TextMatrix(i, TI_结束时间)) Then
                                    .Row = i: .Col = TI_开始时间
                                    Call AddErrInfo("开始使用时间大于了结束使用时间，请检查。", 0, vsTmp)
                                End If
                            End If
                        End If
                    Next
                End With

                Set vsTmp = gclsPros.CurrentForm.vsInfect
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If IsDate(.TextMatrix(i, FI_确诊日期)) Then
                            If gclsPros.OutTime <> "" And Format(.TextMatrix(i, FI_确诊日期), "yyyy-MM-dd") > Format(gclsPros.OutTime, "yyyy-MM-dd") Or _
                                Format(.TextMatrix(i, FI_确诊日期), "yyyy-MM-dd") < Format(gclsPros.InTime, "yyyy-MM-dd") Then
                                .Row = i: .Col = FI_确诊日期
                                Call AddErrInfo("确诊时间不在入出院日期范围内。", 0, vsTmp)
                            End If
                            If .TextMatrix(i, FI_感染部位) = "" Then
                                .Row = i: .Col = FI_感染部位
                                Call AddErrInfo("请输入感染部位。", 0, vsTmp)
                            End If
                            If .TextMatrix(i, FI_医院感染名称) = "" Then
                                .Row = i: .Col = FI_医院感染名称
                                Call AddErrInfo("请输入医院感染名称。", 0, vsTmp)
                            End If
                        ElseIf Trim(.TextMatrix(i, FI_确诊日期)) <> "" And Not IsDate(.TextMatrix(i, FI_确诊日期)) Then
                            .Row = i: .Col = FI_确诊日期
                            Call AddErrInfo("请输入正确的确诊日期。", 0, vsTmp)
                        End If
                    Next
                End With

                Set vsTmp = gclsPros.CurrentForm.vsSample
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, MI_标本) <> "" Then
                            If .TextMatrix(i, MI_病原学代码及名称) = "" Then
                                .Row = i: .Col = MI_病原学代码及名称
                                Call AddErrInfo("请输入病原学代码及名称。", 0, vsTmp)
                            End If
                            If Not IsDate(.TextMatrix(i, MI_送检日期)) Then
                                .Row = i: .Col = MI_送检日期
                                Call AddErrInfo("请输入正确的送检日期。", 0, vsTmp)
                            End If
                        End If
                    Next
                End With

            End If

            If gclsPros.ReadPages And gclsPros.MedPageSandard = ST_卫生部标准 Then
                Set vsTmp = gclsPros.CurrentForm.vsSpirit
                With vsTmp
                    For i = .FixedRows To .Rows - 1
                        If Trim(.TextMatrix(i, SI_药物名称)) <> "" Then
                            '不进行单引号输入检查，因为Form_KeyDown事件已经控制
                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_药物名称))) > 200 Then
                                .Row = i: .Col = SI_药物名称
                                Call AddErrInfo("药物名称输入内容太长，只允许200个字符或100个汉字。", 0, vsTmp)
                            End If
                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_疗程))) > 50 Then
                                .Row = i: .Col = SI_疗程
                                 Call AddErrInfo("疗程输入内容太长，只允许50个字符或25个汉字。", 0, vsTmp)
                            End If
                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_疗效))) > 50 Then
                                .Row = i: .Col = SI_疗效
                                Call AddErrInfo("疗效输入内容太长，只允许50个字符或25个汉字。", 0, vsTmp)
                            End If
                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_特殊反应))) > 100 Then
                                .Row = i: .Col = SI_特殊反应
                                Call AddErrInfo("特殊反应输入内容太长，只允许100个字符或50个汉字。", 0, vsTmp)
                            End If

                            If zlCommFun.ActualLen(Trim(.TextMatrix(i, SI_最高日量))) > 50 Then
                                .Row = i: .Col = SI_最高日量
                                Call AddErrInfo("最高日量输入内容太长，只允许50个字符或25个汉字。", 0, vsTmp)
                            End If
                        End If
                    Next
                End With
            End If
        End If
    End If
    
    
   '检查是否开启外挂部件检查首页
    Call CreatePlugInOK(gclsPros.Module)
    If Not gobjPlugIn Is Nothing Then
        Err.Clear: On Error Resume Next
        If gobjPlugIn.gblnmec = True Then
            '调用病案审查接口
            If Err.Number = 0 Then
                Set gColCtl = CtlAdd
                strMsg = ""
                If gobjPlugIn.CheckMecInfo(gclsPros.SysNo, gclsPros.Module, gclsPros.病人ID, gclsPros.主页ID, gColCtl, strMsg) = False Then
                    If strMsg <> "" And Err.Number = 0 Then
                        Call ErrDw(strMsg)
                    End If
                End If
            End If
            Call zlPlugInErrH(Err, "CheckMecInfo")
            Err.Clear: On Error GoTo 0
        End If
    End If
    
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            Err.Clear: On Error Resume Next
            Call gfrmMecCol(i).CheckPlugMec(gclsPros.SysNo, gclsPros.Module, gclsPros.病人ID, gclsPros.主页ID, colErr)
            Call zlPlugInErrH(Err, "CheckPlugMec")
        Next
        If colErr.Count > 0 And Err.Number = 0 Then
            Call ErrMec(colErr)
            Set colErrTmp = colErr
        End If
        Err.Clear: On Error GoTo 0
    End If
     
'    加载错误和警告到界面上
    If gColErr.Count > 0 Or gColWarn.Count > 0 Then
        Call LoadVsErrData
        If Not blnCheck Then
            If gColErr.Count = 0 And gColWarn.Count > 0 Then
                If MsgBox("检查出" & CStr(gColWarn.Count) & "个警告，是否忽略全部警告，继续操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                Else
                    Call ClearErrCol
                    If blnBaseInfo Then
                        With gclsPros.CurrentForm
                            If Not gobjPatient.SavePatiBaseInfo(gclsPros.病人ID, gclsPros.主页ID, .txtInfo(GC_姓名).Text, strSex, strAge, strBirthday, IIf(gclsPros.PatiType = PF_门诊, "门诊首页", "住院首页"), gclsPros.PatiType, strErrIfno) Then
                                If InStr(strErrIfno, "性别") > 0 Then
                                    Set objTmp = .cboBaseInfo(BCC_性别)
                                ElseIf InStr(strErrIfno, "出生日期") > 0 Then
                                    Set objTmp = .mskDateInfo(DC_出生日期)
                                ElseIf InStr(strErrIfno, "年龄") > 0 Then
                                    Set objTmp = .txtSpecificInfo(SLC_年龄)
                                End If
                                If ShowMessage(objTmp, "身份证号码获取的" & strBaseInfo & "与当前界面的" & strBaseInfo & "不相符，自动更新界面上的" & strBaseInfo & "失败，失败原因：" & strErrIfno & ",是否继续？", True) = vbNo Then Exit Function
                            End If
                            Call SetCtrlValues("性别", strSex)
                            Call SetCtrlValues("年龄", strAge)
                            Call SetCtrlValues("出生日期", strBirthday)
                        End With
                    End If
                    CheckMedPageData = True
                End If
            ElseIf gColErr.Count > 0 Then
                CheckMedPageData = False
                Exit Function
            End If
        End If
    End If
    
    If Not CheckMedPageChange Then
        gclsPros.InfosChange = False
        gclsPros.IsCheckData = False
        Exit Function
    Else
        gclsPros.IsCheckData = False
    End If
    
    CheckMedPageData = True
    Exit Function
errH:
    Debug.Print "CheckMedPageData:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckMedPageChange() As Boolean
'功能：检查首页信息是否发生变化
'缓存界面数据
    On Error GoTo errH
'    gclsPros.InfosChange = False
    With gclsPros.CurrentForm
        Call gclsPros.RollBackCacheRecInfo '回滚信息缓存
        '病案主页，病案主页从表信息缓存
        Call CacheCtrlValues
        '诊断缓存
        '1、缓存病原学诊断
        If gclsPros.PatiType <> PF_门诊 Then Call UpdateCacheRecInfo(1, "病原学诊断", .txtInfo(GC_病原学诊断).Text & .cmdInfo(GC_病原学诊断).Tag)
        '2、缓存西医诊断
        Call CacheLoadVsDiagData(.vsDiagXY, , , True)
        '3、缓存中医诊断
        If gclsPros.Have中医 Then
            Call CacheLoadVsDiagData(.vsDiagZY, , , True)
        End If
        '过敏信息缓存
        Call CacheLoadVsAllerData(.vsAller, , True)
        If gclsPros.PatiType <> PF_门诊 Then
            '手术缓存
            Call CacheLoadVsOPSData(.vsOPS, , True)
            '诊断符合情况缓存
            Call CacheLoadDiagMatchData(, True)
            '病人费用信息缓存
            If gclsPros.FuncType = f病案首页 Then
                Call CacheLoadVsFreesData(.vsFees, , True)
            End If
            '抗菌药使用情况缓存
            Call CacheLoadVsKSSData(.vsKSS, , True)
            '重症监护使用情况缓存
            If gclsPros.MedPageSandard <> ST_湖南省标准 Then
                If gclsPros.MedPageSandard <> ST_云南省标准 Then
                    Call CacheLoadVsFlxAddICUData(.vsFlxAddICU, , True)
                Else
                    Call CacheLoadVsFlxAddICUData(, , True)
                End If
            End If
            '放疗、化疗、精神药品缓存
            If gclsPros.ReadPages Then
                Call CacheLoadVsChemothData(.vsChemoth, , True)
                Call CacheLoadVsRadiothData(.vsRadioth, , True)
                If gclsPros.MedPageSandard = ST_卫生部标准 Then Call CacheLoadVsSpiritData(.vsSpirit, , True)
            End If
            '重症监护器械，医院感染、标本情况缓存
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                Call CacheLoadVsICUInstrumentsData(.vsICUInstruments, , True)
                Call CacheLoadvsInfectData(.vsInfect, , True)
                Call CacheLoadvsSampleData(.vsSample, , True)
            End If
        End If
        '整体检查控件值变化
        Call UpdateCacheRecInfo(2)
        '并发检查病案是否编目或首页处于锁定状态
        If gclsPros.PatiType <> PF_门诊 And gclsPros.FuncType <> f病案首页 Then
            If Not CheckMecRed(gclsPros.病人ID, gclsPros.主页ID, .Caption, "修改首页") Then Exit Function
        End If
    End With
    CheckMedPageChange = True
    Exit Function
errH:
    Debug.Print "CheckMedPageChange:" & Err.Source & "===" & Err.Description
    Call gclsPros.RollBackCacheRecInfo '回滚信息缓存
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SaveMedPageData() As Boolean
'功能：保存病人首页数据
    Dim arrSQL() As Variant
    Dim i As Long
    Dim blnTrans As Boolean
    Dim datCur As Date, strMsg As String
    Dim blnChange As Boolean
    
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            blnChange = blnChange Or gfrmMecCol(i).gblnchange
        Next
        blnChange = blnChange Or gclsPros.InfosChange
    Else
        blnChange = gclsPros.InfosChange
    End If
    If blnChange Then
        datCur = zlDatabase.Currentdate
        arrSQL = Array()
        '现进行病案主页以及病人信息的更新，因为ZL_病案主页_首页整理对原始病案主页从表中的“主治，住院，主任，门诊，医师”信息进行读取
        '若先保存病案主页从表，则会导致，读取到更新后的信息，导致ZL_病案主页_首页整理检查失败
        gclsPros.MainInfoRec.Filter = "是否改变=1"
        If Not gclsPros.Is护士站 And gclsPros.MainInfoRec.RecordCount <> 0 Then
            Call PopPatiMainSQL(arrSQL)
        End If
        '从表信息保存
        Call PopPatiAuxiSQL(arrSQL, gclsPros.Is护士站)
        '医生与护士分添首页,医生站不保存的从表项目
        If Not gclsPros.SeparateEdit And gclsPros.PatiType = PF_住院 Then
            Call PopPatiAuxiSQL(arrSQL, True)
        End If
        If Not gclsPros.Is护士站 Then
            gclsPros.MainInfoRec.Filter = "是否改变=1"
            If gclsPros.MainInfoRec.RecordCount <> 0 Then
                If gclsPros.PatiType = PF_住院 Then
                    '结构化地址保存
                    If gclsPros.IsStructAdress Then
                        Call PopStructAdressSQL(arrSQL)
                    End If
                    '诊断符合情况保存
                    Call PopDiagMatchSQL(arrSQL)
                    '手麻保存
                    Call PopOPSSQL(arrSQL)
                    '抗菌药保存
                    Call PopKSSSQL(arrSQL)
                    If gclsPros.MedPageSandard <> ST_湖南省标准 Then
                        '重症监护情况保存
                        Call PopICUSQL(arrSQL)
                        If gclsPros.MedPageSandard = ST_四川省标准 Then Call PopOtherSQL(arrSQL)
                    End If
                    '病人费用信息保存
                    If gclsPros.FuncType = f病案首页 Then
                        Call PopFeeSQL(arrSQL)
                    End If
                    '放疗，化疗，精神药保存
                    If gclsPros.ReadPages Then
                        Call PopShareInfoSQL(arrSQL)
                    End If
                End If
                '过敏保存
                Call PopAllerSQL(arrSQL)
                '诊断保存
                Call PopPatiDiagSQL(arrSQL, datCur)
            End If
            '分娩信息保存
            If gclsPros.FuncType = f病案首页 Then Call PopDelicerySQL(arrSQL)
        End If
        Screen.MousePointer = 11
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), gclsPros.CurrentForm.Caption)
        Next
        
        '外挂附页数据保存
        If gBlnNew And (Not gfrmMecCol Is Nothing) Then
            For i = 1 To gfrmMecCol.Count
                If gfrmMecCol(i).gblnchange Then
                    Err.Clear: On Error Resume Next
                    If gfrmMecCol(i).savePlugMec(gclsPros.SysNo, gclsPros.Module, gclsPros.病人ID, gclsPros.主页ID) = False Then
                        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                    End If
                    Call zlPlugInErrH(Err, "SavePlugMec")
                    Err.Clear: On Error GoTo 0
                End If
            Next
        End If


        If gclsPros.FuncType = f医生首页 Then
            If gclsPros.PatiType = PF_门诊 Then
                '社区档案同步
                If Not gobjCommunity Is Nothing And gclsPros.CommunityID <> 0 Then
                    If Not gobjCommunity.UpdateInfo(gclsPros.SysNo, p门诊医生站, gclsPros.CommunityID, gclsPros.CommunityNO, gclsPros.病人ID, gclsPros.主页ID) Then
                        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                    End If
                End If
            Else
                '调用医保病人信息修改接口
                If gclsPros.InsureType <> 0 And Not gclsInsure Is Nothing Then
                    If Not gclsInsure.ModiPatiSwap(gclsPros.病人ID, gclsPros.主页ID, gclsPros.InsureType, "2") Then
                        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                    End If
                End If
            End If
            
            If gobjPlugIn Is Nothing Then
                Call CreatePlugInOK(IIf(gclsPros.PatiType = PF_门诊, p门诊医生站, p住院医生站))
            End If
            If Not gobjPlugIn Is Nothing And gclsPros.PatiType = PF_住院 Then
                Err.Clear: On Error Resume Next
                If gobjPlugIn.EMPI_ModifyPatiInfo(gclsPros.SysNo, p住院医生站, gclsPros.病人ID, gclsPros.主页ID, 0, strMsg) = 0 Then
                    If Err.Number = 0 Then
                        gcnOracle.RollbackTrans
                        Screen.MousePointer = 0
                        MsgBox "当前启用了EMPI系统接口，但EMPI系统接口(EMPI_ModifyPatiInfo)未调用成功:" & strMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                If Err.Number <> 0 And Err.Number <> 438 Then
                    gcnOracle.RollbackTrans
                    Screen.MousePointer = 0
                    Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo")
                    Exit Function
                End If
                Err.Clear: On Error GoTo 0
            End If
        End If
        gcnOracle.CommitTrans: blnTrans = False
        '消息触发
        If gclsPros.FuncType <> f病案首页 Then Call SendMsgDiag(datCur)
        '新网接口调用
        If HaveRIS Then
            Call gobjRis.HISModPati(gclsPros.PatiType, gclsPros.病人ID, gclsPros.主页ID)
        End If
    End If
    '缓存信息记录集，更新，将现值赋值给原值,并初始化改变状态
    '目的：首页存在打印预览功能，保存后可以编辑，编辑后又可以保存
    
    Call gclsPros.InitCacheRecInfo(True)
    
    If gclsPros.OpenMode <> EM_新增病案 And gclsPros.OpenMode <> EM_新增首页 Then
        On Error Resume Next
        Call LoadDiagAndAllerFData
    End If

    On Error GoTo errH
    Screen.MousePointer = 0
    gclsPros.InfosChange = False
    SaveMedPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SendMsgDiag(ByVal datCur As Date) As Boolean
'功能：发送诊断消息
    Dim i As Long
    Dim arrTmp As Variant
    Dim strFilter As String
    On Error GoTo errH
    If gclsMipModule Is Nothing Then SendMsgDiag = True: Exit Function
    gclsPros.MainInfoRec.Filter = "(信息名='西医诊断' And 是否改变=1) OR (信息名='中医诊断' And 是否改变=1) "
    For i = 1 To gclsPros.MainInfoRec.RecordCount
        gclsPros.SecdInfoRec.Filter = "改变状态<>" & CS_未改变 & " And 改变状态<>" & CS_更新行 & "   And 序号=" & gclsPros.MainInfoRec!序号
        Do While Not gclsPros.SecdInfoRec.EOF
            arrTmp = Split(gclsPros.SecdInfoRec!信息原值 & "", "|")
            If gclsPros.SecdInfoRec!改变状态 <> CS_新增行 Then '删除行与替换行先触发删除诊断消息
                Call ZLHIS_CIS_011(gclsMipModule, gclsPros.病人ID, gclsPros.PatiName, gclsPros.PatiType, gclsPros.主页ID, gclsPros.出院科室ID, gclsPros.SecdInfoRec!ID, arrTmp(DMP_诊断编码), arrTmp(DMP_疾病编码))
            End If
            arrTmp = Split(gclsPros.SecdInfoRec!信息现值 & "", "|")
            If gclsPros.SecdInfoRec!改变状态 <> CS_删除行 Then  '新增行与替换行触发下达诊断消息
                Call ZLHIS_CIS_010(gclsMipModule, gclsPros.病人ID, gclsPros.PatiName, gclsPros.PatiType, gclsPros.主页ID, gclsPros.出院科室ID, Val(gclsPros.SecdInfoRec!Tag & ""), arrTmp(DMP_诊断类型), arrTmp(DMP_是否疑诊), arrTmp(DMP_诊断次序), arrTmp(DMP_诊断编码), arrTmp(DMP_疾病编码), arrTmp(DMP_疾病附码), arrTmp(DMP_疾病类别), arrTmp(DMP_证候编码), arrTmp(DMP_证候名称), datCur, UserInfo.姓名)
            End If
            gclsPros.SecdInfoRec.MoveNext
        Loop
    Next
    SendMsgDiag = True
    '病原学诊断不触发消息
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function PopPatiMainSQL(ByRef arrSQL As Variant)
'功能：将获取的病案主页或病人信息的SQL放入SQL数组中
'SaveMedPageData的子函数
    Dim lngCallProc As Long
    Dim arrField As Variant, arrTmp As Variant
    Dim i As Long, j As Long
    arrField = Array()
      '病案主页，病人信息，以及病案主页从表
    '1、病案主页以及病人信息的保存
    '先判断，需要调用哪些存储过程
    '病案主页与病人信息共有项目检查
    gclsPros.MainInfoRec.Filter = "是否改变=1"
    If gclsPros.PatiType = PF_住院 Then
        arrField = Array("年龄", "国籍", "区域", "职业", "婚姻状况", "医疗付款方式", "家庭地址", "家庭电话", "家庭地址邮编", "单位地址", _
                        "单位电话", "单位邮编", "联系人姓名", "联系人关系", "联系人电话", "联系人地址", "户口地址", "户口地址邮编")
        If gclsPros.FuncType = f病案首页 Then
             ReDim Preserve arrField(UBound(arrField) + 3)
             arrField(UBound(arrField) - 2) = "住院号"
             arrField(UBound(arrField) - 1) = "入院日期"
             arrField(UBound(arrField)) = "出院日期"
        End If

        For i = LBound(arrField) To UBound(arrField)
            gclsPros.MainInfoRec.MoveFirst
            For j = 1 To gclsPros.MainInfoRec.RecordCount
                If arrField(i) = gclsPros.MainInfoRec!信息名 Then
                    lngCallProc = 1: Exit For
                End If
                gclsPros.MainInfoRec.MoveNext
            Next
            If lngCallProc = 1 Then Exit For
        Next
    End If
    If lngCallProc <> 1 Then
        If gclsPros.PatiType = PF_住院 Then
            '病人信息特有项目检查
            arrField = Array("姓名", "性别", "民族", "籍贯", "出生日期", "出生地点", "身份证号", "其他证件")
            If gclsPros.FuncType <> f病案首页 Then
                ReDim Preserve arrField(UBound(arrField) + 1)
                arrField(UBound(arrField)) = "住院号"
            End If
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                ReDim Preserve arrField(UBound(arrField) + 2)
                arrField(UBound(arrField)) = "Qq"
                arrField(UBound(arrField) - 1) = "Email"
            End If
        Else
             arrField = Array("门诊号", "姓名", "性别", "年龄", "民族", "国籍", "区域", "籍贯", "职业", "出生日期", "出生地点", "身份证号", _
                     "其他证件", "婚姻状况", "医疗付款方式", "家庭地址", "家庭电话", "家庭地址邮编", "户口地址", "户口地址邮编", "合同单位ID", "单位地址", _
                     "单位电话", "单位邮编", "监护人", "复诊", "摘要", "传染病上传", "发病时间", "发病地址")
        End If
        For i = LBound(arrField) To UBound(arrField)
            gclsPros.MainInfoRec.MoveFirst
            For j = 1 To gclsPros.MainInfoRec.RecordCount
                If arrField(i) = gclsPros.MainInfoRec!信息名 Then
                    lngCallProc = 2: Exit For
                End If
                gclsPros.MainInfoRec.MoveNext
            Next
            If lngCallProc = 2 Then Exit For
        Next

        If gclsPros.PatiType = PF_住院 Then
            '病案主页特有项目检查
            arrField = Array("主页id", "身高", "体重", "血型", "入院病况", "入院方式", "出院方式", "再入院", "是否确诊", "确诊日期", "尸检标志", "随诊标志", "随诊期限", _
                    "新发肿瘤", "中医治疗类别", "抢救次数", "成功次数", "门诊医师", "住院医师", "责任护士")
            If gclsPros.FuncType = f病案首页 Then
                arrTmp = Split("病案号,档案号,入院科室ID,出院科室ID,住院天数,费用和,编目员姓名,编目日期", ",")
            Else
                arrTmp = Split("主治医师,主任医师,操作员编号,操作员姓名", ",")
            End If
            ReDim Preserve arrField(UBound(arrField) + UBound(arrTmp) + 1)
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrField(UBound(arrField) - UBound(arrTmp) + i) = arrTmp(i)
            Next
            For i = LBound(arrField) To UBound(arrField)
                gclsPros.MainInfoRec.MoveFirst
                For j = 1 To gclsPros.MainInfoRec.RecordCount
                    If arrField(i) = gclsPros.MainInfoRec!信息名 Then
                        lngCallProc = IIf(lngCallProc = 0, 3, 1): Exit For
                    End If
                    gclsPros.MainInfoRec.MoveNext
                Next
                If lngCallProc = 1 Or lngCallProc = 3 Then Exit For
            Next
        End If
    End If
    If gclsPros.FuncType = f病案首页 Then
        '病案新增病案模式，一定会同时更新病案主页以及病人信息，新增病案模式一定会更新病人信息
        lngCallProc = DecodeEx(gclsPros.OpenMode = EM_新增病案, 1, gclsPros.OpenMode = EM_新增首页 And lngCallProc = 2, 1, lngCallProc)
    End If
    If lngCallProc <> 0 Then
        'ZL_病人信息_首页整理调用
        If lngCallProc <> 3 Then
            arrField = Array("病人ID", IIf(gclsPros.PatiType = PF_门诊, "门诊号", "住院号"), "姓名", "性别", "年龄", "民族", "国籍", "区域", "籍贯", "职业", "出生日期", "出生地点", "身份证号", _
                        "其他证件", "婚姻状况", "医疗付款方式", "家庭地址", "家庭电话", "家庭地址邮编", "户口地址", "户口地址邮编", "合同单位ID", "单位地址", "单位电话", "单位邮编")
            If gclsPros.FuncType = f病案首页 Then
                arrTmp = Split("联系人姓名,联系人关系, 联系人电话, 联系人地址," & IIf(gclsPros.MedPageSandard = ST_四川省标准, "Email,QQ,", ",,") & "入院日期,出院日期,住院次数,主页ID", ",")
            ElseIf gclsPros.PatiType = PF_住院 Then
                arrTmp = Split("联系人姓名,联系人关系, 联系人电话, 联系人地址," & IIf(gclsPros.MedPageSandard = ST_四川省标准, "Email,QQ", ",,"), ",")
            Else
                arrTmp = Split(", , , ,,,监护人,NO,复诊,摘要,传染病上传,发病时间,发病地址", ",")
            End If
            ReDim Preserve arrField(UBound(arrField) + UBound(arrTmp) + 1)
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrField(UBound(arrField) - UBound(arrTmp) + i) = arrTmp(i)
            Next
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = Get首页整理SQL(0, arrField) '获取ZL_病人信息_首页整理的调用SQL
        End If

        'ZL_病案主页_首页整理调用
        If lngCallProc <> 2 Then
            If gclsPros.FuncType = f病案首页 Then
                arrField = Array("病人ID", "主页ID", "住院号", "病案号", "档案号", "年龄", "国籍", "区域", "职业", "身高", "体重", "血型", "婚姻状况", "医疗付款方式", _
                    "家庭地址", "家庭电话", "家庭地址邮编", "户口地址", "户口地址邮编", "单位地址", "单位电话", "单位邮编", "联系人姓名", "联系人关系", _
                    "联系人电话", "联系人地址", "入院病况", "入院方式", "入院科室ID", "入院日期", "出院方式", "出院科室ID", "出院日期", "再入院", _
                    "是否确诊", "确诊日期", "尸检标志", "随诊标志", "随诊期限", "新发肿瘤", "中医治疗类别", "抢救次数", "成功次数", "住院天数", _
                    "费用和", "门诊医师", "住院医师", "责任护士", "编目员姓名", "编目日期")
            Else
                arrField = Array("病人ID", "主页ID", "年龄", "国籍", "区域", "职业", "身高", "体重", "血型", "婚姻状况", "医疗付款方式", _
                    "家庭地址", "家庭电话", "家庭地址邮编", "户口地址", "户口地址邮编", "单位地址", "单位电话", "单位邮编", "联系人姓名", "联系人关系", _
                    "联系人电话", "联系人地址", "入院病况", "入院方式", "出院方式", "再入院", "是否确诊", "确诊日期", "尸检标志", "随诊标志", "随诊期限", _
                    "新发肿瘤", "中医治疗类别", "抢救次数", "成功次数", "门诊医师", "住院医师", "主治医师", "主任医师", "责任护士", "操作员编号", "操作员姓名")
            End If

            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = Get首页整理SQL(1, arrField) '获取ZL_病案主页_首页整理的调用SQL
        End If
    End If
End Function

Public Sub PopPatiAuxiSQL(ByRef arrSQL As Variant, Optional ByVal bln护士站 As Boolean)
'功能：获取一般的从表信息SQL，并放入SQL数组中
'SaveMedPageData的子函数
    Dim arrField As Variant, arrTmp As Variant
    Dim i As Long, j As Long, LngRow As Long, LngCol As Long
    Dim strTmp As String, arrTag As Variant

    arrField = Array()
    If bln护士站 Then
        gclsPros.MainInfoRec.Filter = "是否改变=1"
        strTmp = ",压疮发生期间,压疮分期,跌倒或坠床伤害,跌倒或坠床原因,不良事件,"
        If gclsPros.MedPageSandard = ST_云南省标准 Then
            strTmp = strTmp & "身体约束,约束总时间,约束方式,约束工具,约束原因,"
        ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
            strTmp = strTmp & "输液药物,输液表现,身体约束,透析尿素氮值,输液反应,"
        End If
        '护士站编辑的从表信息
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            gclsPros.MainInfoRec.MoveFirst
            For i = 1 To gclsPros.MainInfoRec.RecordCount
                If strTmp Like "*," & gclsPros.MainInfoRec!信息名 & ",*" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'" & gclsPros.MainInfoRec!信息名 & "','" & gclsPros.MainInfoRec!信息现值 & "')"
                End If
                gclsPros.MainInfoRec.MoveNext
            Next
        End If
    Else
        If gclsPros.PatiType = PF_住院 Then
        '病案附加项目的保存，因为病案附加项目可能会与一般的从表信息名相同，会产生覆盖
        '若果病案附加项目名称与一般从表信息名称相同，则以一般从表信息值为准，因此，先保存病案附加项目
        '病案附加项目的保存
            gclsPros.MainInfoRec.Filter = "信息名='病案项目' And 是否改变=1"
            If gclsPros.MainInfoRec.RecordCount <> 0 Then
                gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态<>0 "
                gclsPros.SecdInfoRec.Sort = "Sort"
                With gclsPros.CurrentForm.vsfMain
                    For i = 1 To gclsPros.SecdInfoRec.RecordCount
                        arrTag = Split(gclsPros.SecdInfoRec!Tag & "", ";")
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'" & arrTag(0) & "','" & gclsPros.SecdInfoRec!信息现值 & "')"
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                End With
            End If
        End If

        gclsPros.MainInfoRec.Filter = "是否改变=1"
        If gclsPros.MainInfoRec.RecordCount <= 0 Then Exit Sub
        '3、一般的病案主页从表信息的保存
        If gclsPros.PatiType = PF_门诊 Then
            '门诊首页从表信息保存
            '生命体征信息保存
            strTmp = gclsPros.CurrentForm.UCPatiVitalSigns.GetSaveSQL(gclsPros.病人ID, gclsPros.主页ID)
            If strTmp <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strTmp
            End If
            gclsPros.MainInfoRec.Filter = "是否改变=1"
            arrField = Array("文化程度", "生育状况", "去向", "RH", "血型", "医学警示", "其他医学警示", "身份证号状态", "无过敏记录", "外籍身份证号", "监护人身份证号")
            For i = LBound(arrField) To UBound(arrField)
                gclsPros.MainInfoRec.Sort = "序号,信息名"
                For j = 1 To gclsPros.MainInfoRec.RecordCount
                    If arrField(i) = gclsPros.MainInfoRec!信息名 Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        If arrField(i) = "医学警示" Or arrField(i) = "其他医学警示" Or arrField(i) = "RH" Or arrField(i) = "血型" Or arrField(i) = "身份证号状态" Or arrField(i) = "外籍身份证号" Then
                            arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & gclsPros.病人ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!信息现值 & "')"
                            If arrField(i) = "RH" Or arrField(i) = "血型" Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & gclsPros.病人ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!信息现值 & "'," & gclsPros.主页ID & ")"
                            End If
                        Else
                            arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & gclsPros.病人ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!信息现值 & "'," & gclsPros.主页ID & ")"
                        End If
                    End If
                    gclsPros.MainInfoRec.MoveNext
                Next
            Next
        Else
            arrField = Array("入院病室", "出院病室", "转科记录", "中医危重", "中医急症", "中医疑难", "中医抢救方法", _
                "自制中药制剂", "死亡根本原因", "死亡时间", "入院前经外院治疗", "示教病案", "科研病案", "疑难病历", "RH", _
                "输血反应", "输红细胞", "输血小板", "输血浆", "输全血", "输其他", "科主任", "主任医师", "主治医师", "进修医师", _
                IIf(gclsPros.MedPageSandard = ST_四川省标准, "主诊医师", "研究生实习医师"), "实习医师", "质控医师", "质控护士", "病原学检查", "输血检查", "彩色多普勒", "特殊检查", _
                "病例分型", "感染因素", "出院转入", "再入院计划天数", "31天内再住院", "不足周岁年龄", "新生儿出生体重", _
                "新生儿入院体重", "呼吸机使用时间", "昏迷时间", "抢救病因", "自体回收", "分化程度", "最高诊断依据", _
                "中医设备", "中医技术", "辨证施护", "病理号", "病案质量", "主页质量日期", "生育状况", "发病时间", _
                "感染与死亡关系", "感染部位", "籍贯", "入院转入", "入院方式", "联系人附加信息", "身份证号状态", "无过敏记录", "住院死亡期间", "外籍身份证号", "监护人身份证号")
            '医学警示用于病人健康卡，病案不需要
            strTmp = IIf(gclsPros.FuncType <> f病案首页, "医学警示,其他医学警示", "收回日期,医保号,主页X线号,特级护理天数,一级护理天数,二级护理天数,三级护理天数,ICU天数,CCU天数,转科时间")
            '四川的输液由护士填写,CT，MRI信息放在特殊检查表格中，在判断到特殊检查时，再保存
            strTmp = strTmp & IIf(gclsPros.MedPageSandard = ST_四川省标准, ",院内会诊,外院会诊,会诊情况,输白蛋白,肿瘤分期,距上一次住本院时间,是否因同一疾病", ",HBSAG,HCV-AB,HIV-AB,输液反应,CT,MRI")
            '临床路径信息，标准版与湖南版没有
            strTmp = strTmp & IIf(gclsPros.MedPageSandard = ST_卫生部标准 Or gclsPros.MedPageSandard = ST_湖南省标准, "", ",临床路径,退出原因,变异原因,告病重病危")
            '湖南版独有肿瘤分期等信息
            strTmp = strTmp & IIf(gclsPros.MedPageSandard <> ST_湖南省标准, "", ",重症监护天数,重症监护小时,单病种,临床路径,肿瘤分期,肿瘤T,肿瘤M,肿瘤N,标本送检,传染病,APGAR,DRGS")
            '云南版独有新生儿离院方式
            strTmp = strTmp & IIf(gclsPros.MedPageSandard <> ST_云南省标准, "", ",新生儿离院方式,围术期死亡,术后猝死")
            '病案云南独有：距上一次住本院时间,是否因同一疾病
            strTmp = strTmp & IIf(gclsPros.MedPageSandard = ST_云南省标准 And gclsPros.FuncType = f病案首页, ",距上一次住本院时间,是否因同一疾病", "")
            arrTmp = Split(strTmp, ",")
            ReDim Preserve arrField(UBound(arrField) + UBound(arrTmp) + 1)
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrField(UBound(arrField) - UBound(arrTmp) + i) = Trim(arrTmp(i))
            Next
            
            For i = LBound(arrField) To UBound(arrField)
                gclsPros.MainInfoRec.MoveFirst
                For j = 1 To gclsPros.MainInfoRec.RecordCount
                    If arrField(i) = gclsPros.MainInfoRec!信息名 Then
                        If arrField(i) = "特殊检查" Then
                            gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态<>" & CS_未改变
                            Do While Not gclsPros.SecdInfoRec.EOF
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                If gclsPros.MedPageSandard <> ST_四川省标准 Then
                                    strTmp = "'特殊检查" & (Val(gclsPros.SecdInfoRec!IndexEx & "") + 4) & "','" & gclsPros.CurrentForm.vsTSJC.TextMatrix(gclsPros.SecdInfoRec!IndexEx, 1) & "'"
                                Else
                                    strTmp = "'" & decode(Val(gclsPros.SecdInfoRec!IndexEx & ""), TR_CT, "CT", TR_PETCT, "PETCT", TR_双源CT, "双源CT", _
                                                TR_X片, "X片", TR_B超, "B超", TR_超声心动图, "超声心动图", TR_MRI, "MRI", TR_同位素检查, "同位素检查") & "','" & Mid(gclsPros.CurrentForm.vsTSJC.TextMatrix(Val(gclsPros.SecdInfoRec!IndexEx & ""), 1), 1, 1) & "'"
                                End If
                                arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & strTmp & ")"
                                gclsPros.SecdInfoRec.MoveNext
                            Loop
                        Else
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!信息现值 & "')"
                            If gclsPros.FuncType <> f病案首页 Then '病案首页，不进行病人信息从表的保存
                                '病人信息从表信息
                                If arrField(i) = "血型" Or arrField(i) = "RH" Or arrField(i) = "医学警示" Or arrField(i) = "其他医学警示" Or arrField(i) = "联系人附加信息" Or arrField(i) = "身份证号状态" Or arrField(i) = "外籍身份证号" Or arrField(i) = "监护人身份证号" Then
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = "zl_病人信息从表_Update(" & gclsPros.病人ID & ",'" & arrField(i) & "','" & gclsPros.MainInfoRec!信息现值 & "')"
                                End If
                            End If
                        End If
                    End If
                    gclsPros.MainInfoRec.MoveNext
                Next
            Next
        End If
    End If
End Sub

Public Sub PopPatiDiagSQL(ByRef arrSQL As Variant, ByVal datCur As Date)
'功能：将诊断SQL加入SQL数组中
'SaveMedPageData的子函数
'西医诊断, 中医诊断,病原学诊断
    Dim k As Integer, LngRow As Long, j As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String
    Dim lngID As Long
    Dim strDiagRowIDs As String, strDiagNames As String
    
    On Error GoTo errH
    
    If gclsPros.FuncType <> f病案首页 Then
        Call MsgDis(gclsPros.DiseaseIDs, gclsPros.DiagIDs)
    End If
    For k = 0 To 1
        gclsPros.MainInfoRec.Filter = "信息名='" & IIf(k = 0, "西医诊断", "中医诊断") & "' And 是否改变=1"
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            Set vsTmp = IIf(k = 0, gclsPros.CurrentForm.vsDiagXY, gclsPros.CurrentForm.vsDiagZY)
            With vsTmp
                '删除行以及主信息改变行需要调用删除方法
                gclsPros.SecdInfoRec.Filter = "(改变状态=" & CS_删除行 & " And 序号=" & gclsPros.MainInfoRec!序号 & ") OR (改变状态=" & CS_替换行 & " And 序号=" & gclsPros.MainInfoRec!序号 & ")": gclsPros.SecdInfoRec.Sort = "Sort": strTmp = ""
                Do While Not gclsPros.SecdInfoRec.EOF
                    strTmp = strTmp & "," & gclsPros.SecdInfoRec!ID
                    gclsPros.SecdInfoRec.MoveNext
                Loop
                If strTmp <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    '病案系统存储过程需加系统编号
                    arrSQL(UBound(arrSQL)) = "Zl" & IIf(gclsPros.FuncType = f病案首页, 3, "") & "_病人诊断记录_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & IIf(gclsPros.FuncType = f病案首页, 4, 3) & ",NULL,NUll,'" & Mid(strTmp, 2) & "')"
                End If
                '主信息改变以及新增行需要调用插入过程
                '次级信息改变，调用更新过程
                gclsPros.SecdInfoRec.Filter = "改变状态>" & CS_未改变 & " And 序号=" & gclsPros.MainInfoRec!序号: gclsPros.SecdInfoRec.Sort = "Sort"
                Do While Not gclsPros.SecdInfoRec.EOF
                    LngRow = gclsPros.SecdInfoRec!IndexEx: j = Val(Mid(gclsPros.SecdInfoRec!信息现值, 1, InStr(gclsPros.SecdInfoRec!信息现值, "|") - 1))
                    If Trim(.TextMatrix(LngRow, DI_诊断编码)) = "" Then
                        strTmp = .TextMatrix(LngRow, DI_诊断描述) & IIf(.TextMatrix(LngRow, DI_中医证候) <> "", "(" & .TextMatrix(LngRow, DI_中医证候) & ")", "")
                    Else
                        strTmp = "(" & .TextMatrix(LngRow, DI_诊断编码) & ")" & .TextMatrix(LngRow, DI_诊断描述) & IIf(.TextMatrix(LngRow, DI_中医证候) <> "", "(" & .TextMatrix(LngRow, DI_中医证候) & ")", "")
                    End If

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If gclsPros.SecdInfoRec!改变状态 <> CS_更新行 Then
                        If gclsPros.FuncType = f病案首页 Then
                            arrSQL(UBound(arrSQL)) = "ZL3_病人诊断记录_INSERT(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & .TextMatrix(LngRow, DI_诊断分类) & "," & _
                                    ZVal(.TextMatrix(LngRow, DI_疾病ID)) & "," & ZVal(.TextMatrix(LngRow, DI_诊断ID)) & "," & ZVal(.TextMatrix(LngRow, DI_证候ID)) & ",'" & _
                                    strTmp & "','" & zlStr.NeedName(.TextMatrix(LngRow, DI_出院情况)) & "'," & IIf(.TextMatrix(LngRow, DI_是否未治) = "", 0, 1) & "," & _
                                    IIf(.TextMatrix(LngRow, DI_是否疑诊) = "", 0, 1) & "," & j & ",'" & .TextMatrix(LngRow, DI_备注) & "','" & _
                                    .TextMatrix(LngRow, DI_入院病情) & "'," & ZVal(.TextMatrix(LngRow, DI_附码ID)) & ")"
                            gclsPros.AddDiag = True
                        Else
                            lngID = zlDatabase.GetNextId("病人诊断记录")
                            gclsPros.SecdInfoRec.Update "Tag", lngID '保存新ID
                            .RowData(LngRow) = lngID
                            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",3,NULL," & .TextMatrix(LngRow, DI_诊断分类) & "," & _
                                                ZVal(.TextMatrix(LngRow, DI_疾病ID)) & "," & ZVal(.TextMatrix(LngRow, DI_诊断ID)) & "," & ZVal(.TextMatrix(LngRow, DI_证候ID)) & ",'" & _
                                                strTmp & "','" & zlStr.NeedName(.TextMatrix(LngRow, DI_出院情况)) & "'," & IIf(.TextMatrix(LngRow, DI_是否未治) = "", 0, 1) & "," & _
                                                IIf(.TextMatrix(LngRow, DI_是否疑诊) = "", 0, 1) & "," & zlStr.To_Date(datCur, "ymdhms") & ",'" & .TextMatrix(LngRow, DI_医嘱IDs) & "' ," & j & ",'" & .TextMatrix(LngRow, DI_备注) & "','" & _
                                                .TextMatrix(LngRow, DI_入院病情) & "'," & zlStr.To_Date(.TextMatrix(LngRow, DI_发病时间), "ymdhm") & ",Null," & lngID & "," & ZVal(.TextMatrix(LngRow, DI_附码ID)) & ")"
                            gclsPros.AddDiag = True
                        End If
                    Else
                        If gclsPros.FuncType = f病案首页 Then
                            arrSQL(UBound(arrSQL)) = "Zl3_病人诊断记录_Update(" & gclsPros.SecdInfoRec!ID & "," & .TextMatrix(LngRow, DI_诊断分类) & "," _
                                                & ZVal(.TextMatrix(LngRow, DI_疾病ID)) & "," & ZVal(.TextMatrix(LngRow, DI_诊断ID)) & "," & ZVal(.TextMatrix(LngRow, DI_证候ID)) & ",'" & _
                                                strTmp & "','" & zlStr.NeedName(.TextMatrix(LngRow, DI_出院情况)) & "'," & IIf(.TextMatrix(LngRow, DI_是否未治) = "", 0, 1) & "," _
                                                & IIf(.TextMatrix(LngRow, DI_是否疑诊) = "", 0, 1) & "," & j & ",'" & .TextMatrix(LngRow, DI_备注) & "','" & _
                                                .TextMatrix(LngRow, DI_入院病情) & "'," & ZVal(.TextMatrix(LngRow, DI_附码ID)) & ")"
                        Else
                            arrSQL(UBound(arrSQL)) = "Zl_病人诊断记录_Update(" & gclsPros.SecdInfoRec!ID & "," & gclsPros.病人ID & "," & gclsPros.主页ID & ",3," & .TextMatrix(LngRow, DI_诊断分类) & "," _
                                                & ZVal(.TextMatrix(LngRow, DI_疾病ID)) & "," & ZVal(.TextMatrix(LngRow, DI_诊断ID)) & "," & ZVal(.TextMatrix(LngRow, DI_证候ID)) & ",'" & _
                                                strTmp & "','" & zlStr.NeedName(.TextMatrix(LngRow, DI_出院情况)) & "'," & IIf(.TextMatrix(LngRow, DI_是否未治) = "", 0, 1) & "," _
                                                & IIf(.TextMatrix(LngRow, DI_是否疑诊) = "", 0, 1) & "," & j & ",'" & .TextMatrix(LngRow, DI_备注) & "','" & _
                                                .TextMatrix(LngRow, DI_入院病情) & "'," & zlStr.To_Date(.TextMatrix(LngRow, DI_发病时间), "ymdhm") & "," & ZVal(.TextMatrix(LngRow, DI_附码ID)) & ")"
                        End If
                    End If
                    gclsPros.SecdInfoRec.MoveNext
                Loop
            End With
        End If
    Next
    
    '诊断选择组织返回值
    If gclsPros.FuncType = f诊断选择 Then
        For k = 0 To 1
            Set vsTmp = IIf(k = 0, gclsPros.CurrentForm.vsDiagXY, gclsPros.CurrentForm.vsDiagZY)
            With vsTmp
                For LngRow = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(LngRow, DI_关联)) <> 0 Then
                        If Trim(.TextMatrix(LngRow, DI_诊断编码)) = "" Then
                            strTmp = .TextMatrix(LngRow, DI_诊断描述) & IIf(.TextMatrix(LngRow, DI_中医证候) <> "", "(" & .TextMatrix(LngRow, DI_中医证候) & ")", "")
                        Else
                            strTmp = "(" & .TextMatrix(LngRow, DI_诊断编码) & ")" & .TextMatrix(LngRow, DI_诊断描述) & IIf(.TextMatrix(LngRow, DI_中医证候) <> "", "(" & .TextMatrix(LngRow, DI_中医证候) & ")", "")
                        End If
                        strDiagRowIDs = strDiagRowIDs & "," & Val(.RowData(LngRow))
                        strDiagNames = strDiagNames & "," & strTmp
                    End If
                Next
            End With
        Next
        gclsPros.DiagRowIDs = Mid(strDiagRowIDs, 2)
        gclsPros.DiagNames = Mid(strDiagNames, 2)
    End If

    If gclsPros.PatiType = PF_住院 And gclsPros.FuncType <> f诊断选择 Then
        gclsPros.MainInfoRec.Filter = "信息名='病原学诊断' And 是否改变=1"
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_DELETE(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & IIf(gclsPros.FuncType = f病案首页, 4, 3) & ",NULL,'21')"
            If Not NVL(gclsPros.MainInfoRec!信息现值) = "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If gclsPros.FuncType = f病案首页 Then
                    arrSQL(UBound(arrSQL)) = "ZL3_病人诊断记录_INSERT(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",21," & _
                        ZVal(gclsPros.CurrentForm.cmdInfo(GC_病原学诊断).Tag) & ",NULL,NULL,'" & gclsPros.CurrentForm.txtInfo(GC_病原学诊断).Text & "',NULL,NULL,NULL)"
                Else
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",3,NULL,21," & _
                            ZVal(gclsPros.CurrentForm.cmdInfo(GC_病原学诊断).Tag) & ",NULL,NULL,'" & gclsPros.CurrentForm.txtInfo(GC_病原学诊断).Text & "',NULL,NULL,NULL," & _
                            zlStr.To_Date(datCur, "ymdhms") & ",Null,1,Null,Null)"
                End If
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub PopKSSSQL(ByRef arrSQL As Variant)
'功能：将抗生素SQL加入数组
'SaveMedPageData的子函数
    Dim LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String, arrTmp As Variant

     '使用抗生素的记录
    gclsPros.MainInfoRec.Filter = "信息名='病人抗生素记录' And 是否改变=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsKSS
        With vsTmp
            gclsPros.SecdInfoRec.Filter = "改变状态<>" & CS_未改变 & " And 序号=" & gclsPros.MainInfoRec!序号: gclsPros.SecdInfoRec.Sort = "改变状态,Sort"
            Do While Not gclsPros.SecdInfoRec.EOF
                '删除行以及主信息改变行需要调用过程传入功能2（删除）
                If gclsPros.SecdInfoRec!改变状态 = CS_删除行 Or gclsPros.SecdInfoRec!改变状态 = CS_替换行 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrTmp = Split(gclsPros.SecdInfoRec!信息原值, "|")
                    arrSQL(UBound(arrSQL)) = "Zl_病人抗生素记录_Update(" & _
                            "2," & gclsPros.病人ID & "," & gclsPros.主页ID & "," & Val(arrTmp(0)) & ",'" & arrTmp(1) & "','" & Trim(arrTmp(2)) & "','" & Trim(arrTmp(3)) & "')"
                End If
                '主信息改变以及新增行需要调用过程传入功能0（新增）
                '次级信息改变，调用过程传入功能1（修改）
                If gclsPros.SecdInfoRec!改变状态 > CS_未改变 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    LngRow = gclsPros.SecdInfoRec!IndexEx
                    arrSQL(UBound(arrSQL)) = "Zl_病人抗生素记录_Update(" & _
                                decode(gclsPros.SecdInfoRec!改变状态, 1, 1, 2, 0, 3, 0) & "," & gclsPros.病人ID & "," & gclsPros.主页ID & "," & Val(.RowData(LngRow) & "") & ",'" & .TextMatrix(LngRow, KI_抗菌药物名) & "','" & _
                                Trim(.TextMatrix(LngRow, KI_用药目的)) & "','" & Trim(.TextMatrix(LngRow, KI_使用阶段)) & "'," & Val(.TextMatrix(LngRow, KI_使用天数)) & ",'" & UserInfo.姓名 & "',Sysdate," & _
                                ZVal(.Cell(flexcpChecked, LngRow, KI_一类切口预防用)) & "," & ZVal(.TextMatrix(LngRow, KI_DDD数)) & ",'" & .TextMatrix(LngRow, KI_联合用药) & "')"
                End If
                gclsPros.SecdInfoRec.MoveNext
            Loop
        End With
        '将老数据-病案主页从表的数据一起删去
        gclsPros.AuxiInfo.Filter = "信息名 Like '抗生素*'": gclsPros.AuxiInfo.Sort = "信息名"
        If gclsPros.AuxiInfo.RecordCount > 0 Then
            Do While Not gclsPros.AuxiInfo.EOF
                If IsNumeric(Mid(gclsPros.AuxiInfo!信息名, 4)) Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'" & gclsPros.AuxiInfo!信息名 & "',NULL)"
                End If
                gclsPros.AuxiInfo.MoveNext
            Loop
        End If
    End If

End Sub

Private Sub PopOPSSQL(ByRef arrSQL As Variant)
'功能：将手麻SQL加入数组
'SaveMedPageData的子函数
    Dim LngRow As Long, lngOrder As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String, arrTmp As Variant
    Dim lngID As Long
    Dim OPSid As Long
    
      '手麻情况
    gclsPros.MainInfoRec.Filter = "信息名='手麻情况' And 是否改变=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsOPS
        With vsTmp
            '删除行以及主信息改变行需要调用删除方法
            gclsPros.SecdInfoRec.Filter = "(改变状态=" & CS_删除行 & " And 序号=" & gclsPros.MainInfoRec!序号 & ") OR (改变状态=" & CS_替换行 & " And 序号=" & gclsPros.MainInfoRec!序号 & ")": gclsPros.SecdInfoRec.Sort = "Sort": strTmp = ""
            Do While Not gclsPros.SecdInfoRec.EOF
                strTmp = strTmp & "," & gclsPros.SecdInfoRec!ID
                gclsPros.SecdInfoRec.MoveNext
            Loop
            If strTmp <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人手麻记录_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & IIf(gclsPros.FuncType = f病案首页, 4, 3) & ",Null,'" & Mid(strTmp, 2) & "')"
            End If
            
            '主信息改变以及新增行需要调用插入过程
            '次级信息改变，调用更新过程
            gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态>" & CS_未改变: gclsPros.SecdInfoRec.Sort = "Sort"
            Do While Not gclsPros.SecdInfoRec.EOF
                LngRow = gclsPros.SecdInfoRec!IndexEx: lngOrder = GetOPSOrder(vsTmp, LngRow)
                strTmp = Trim(.TextMatrix(LngRow, PI_切口愈合))
                If strTmp = "" Then strTmp = "/"
                strTmp = strTmp & "/" & decode(.TextMatrix(LngRow, PI_手术级别), "一级手术", 1, "二级手术", 2, "三级手术", 3, "四级手术", 4, "无", 9, 0)
                arrTmp = Split(strTmp, "/")
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                
                '处理手术结束日期在界面表格中不显示时，结束日期就等于手术日期
                If gclsPros.MedPageSandard = ST_卫生部标准 Or gclsPros.MedPageSandard = ST_湖南省标准 Then
                    If Not gclsPros.UseOPSEndTime Then
                        .TextMatrix(LngRow, PI_结束日期) = .TextMatrix(LngRow, PI_手术日期)
                    End If
                ElseIf gclsPros.MedPageSandard = ST_云南省标准 Then
                    .TextMatrix(LngRow, PI_结束日期) = .TextMatrix(LngRow, PI_手术日期)
                End If

                If gclsPros.SecdInfoRec!改变状态 <> CS_更新行 Then
                    lngID = zlDatabase.GetNextId("病人手麻记录")
                    arrSQL(UBound(arrSQL)) = "ZL_病人手麻记录_Insert(" & lngID & "," & gclsPros.病人ID & "," & gclsPros.主页ID & "," & IIf(gclsPros.FuncType = f病案首页, 4, 3) & "," & lngOrder & "," & _
                            zlStr.To_Date(.TextMatrix(LngRow, PI_手术日期), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_手术日期), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_结束日期), "ymdhm") & "," & _
                            "NULL," & ZVal(.TextMatrix(LngRow, PI_手术操作ID)) & "," & ZVal(.TextMatrix(LngRow, PI_诊疗项目ID)) & ",'" & .TextMatrix(LngRow, PI_手术名称) & "','" & .TextMatrix(LngRow, PI_主刀医师) & "','" & _
                            .TextMatrix(LngRow, PI_助产护士) & "','" & .TextMatrix(LngRow, PI_助手1) & "','" & .TextMatrix(LngRow, PI_助手2) & "',NULL," & zlStr.To_Date(.TextMatrix(LngRow, PI_麻醉开始时间), "ymdhm") & ",NULL," & ZVal(.TextMatrix(LngRow, PI_麻醉ID)) & ",'" & _
                            .TextMatrix(LngRow, PI_麻醉类型) & "',NULL,NULL,'" & .TextMatrix(LngRow, PI_麻醉医师) & "',NULL,NULL,'" & arrTmp(0) & "','" & arrTmp(1) & "',Sysdate,'" & .TextMatrix(LngRow, PI_手术情况) & "','" & _
                            .TextMatrix(LngRow, PI_ASA分级) & "'," & Abs(Val(.TextMatrix(LngRow, PI_再次手术))) & ",'" & .TextMatrix(LngRow, PI_NNIS分级) & "'," & arrTmp(2) & "," & ZVal(.TextMatrix(LngRow, PI_准备天数)) & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_抗菌用药时间), "ymdhm") & ",'" & _
                            .TextMatrix(LngRow, PI_切口部位) & "'," & .Cell(flexcpChecked, LngRow, PI_切口感染) & "," & .Cell(flexcpChecked, LngRow, PI_并发症) & "," & .Cell(flexcpChecked, LngRow, PI_重返手术室计划) & ",'" & .TextMatrix(LngRow, PI_重返手术室目的) & "'," & _
                            .Cell(flexcpChecked, LngRow, PI_预防用抗菌药) & "," & Val(.TextMatrix(LngRow, PI_抗菌药天数)) & "," & .Cell(flexcpChecked, LngRow, PI_非预期的二次手术) & "," & .Cell(flexcpChecked, LngRow, PI_麻醉并发症) & "," & .Cell(flexcpChecked, LngRow, PI_术中异物遗留) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_手术并发症) & "," & .Cell(flexcpChecked, LngRow, PI_术后出血或血肿) & "," & .Cell(flexcpChecked, LngRow, PI_手术伤口裂开) & "," & .Cell(flexcpChecked, LngRow, PI_术后深静脉血栓) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_术后生理代谢紊乱) & "," & .Cell(flexcpChecked, LngRow, PI_术后呼吸衰竭) & "," & .Cell(flexcpChecked, LngRow, PI_术后肺栓塞) & "," & .Cell(flexcpChecked, LngRow, PI_术后败血症) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_术后髋关节骨折) & ")"
                    .RowData(LngRow) = lngID
                    gclsPros.SecdInfoRec!ID = lngID
                Else
                    OPSid = IIf(.RowData(LngRow) <> "", .RowData(LngRow), Val(gclsPros.SecdInfoRec!ID & ""))
                    arrSQL(UBound(arrSQL)) = "ZL_病人手麻记录_Update(" & OPSid & "," & gclsPros.病人ID & "," & gclsPros.主页ID & "," & IIf(gclsPros.FuncType = f病案首页, 4, 3) & "," & lngOrder & "," & _
                            zlStr.To_Date(.TextMatrix(LngRow, PI_手术日期), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_手术日期), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_结束日期), "ymdhm") & "," & _
                            "NULL," & ZVal(.TextMatrix(LngRow, PI_手术操作ID)) & "," & ZVal(.TextMatrix(LngRow, PI_诊疗项目ID)) & ",'" & .TextMatrix(LngRow, PI_手术名称) & "','" & .TextMatrix(LngRow, PI_主刀医师) & "','" & _
                            .TextMatrix(LngRow, PI_助产护士) & "','" & .TextMatrix(LngRow, PI_助手1) & "','" & .TextMatrix(LngRow, PI_助手2) & "',NULL," & zlStr.To_Date(.TextMatrix(LngRow, PI_麻醉开始时间), "ymdhm") & ",NULL," & ZVal(.TextMatrix(LngRow, PI_麻醉ID)) & ",'" & _
                            .TextMatrix(LngRow, PI_麻醉类型) & "',NULL,NULL,'" & .TextMatrix(LngRow, PI_麻醉医师) & "',NULL,NULL,'" & arrTmp(0) & "','" & arrTmp(1) & "','" & .TextMatrix(LngRow, PI_手术情况) & "','" & _
                            .TextMatrix(LngRow, PI_ASA分级) & "'," & Abs(Val(.TextMatrix(LngRow, PI_再次手术))) & ",'" & .TextMatrix(LngRow, PI_NNIS分级) & "'," & arrTmp(2) & "," & ZVal(.TextMatrix(LngRow, PI_准备天数)) & "," & zlStr.To_Date(.TextMatrix(LngRow, PI_抗菌用药时间), "ymdhm") & ",'" & _
                            .TextMatrix(LngRow, PI_切口部位) & "'," & .Cell(flexcpChecked, LngRow, PI_切口感染) & "," & .Cell(flexcpChecked, LngRow, PI_并发症) & "," & .Cell(flexcpChecked, LngRow, PI_重返手术室计划) & ",'" & .TextMatrix(LngRow, PI_重返手术室目的) & "'," & _
                            .Cell(flexcpChecked, LngRow, PI_预防用抗菌药) & "," & Val(.TextMatrix(LngRow, PI_抗菌药天数)) & "," & .Cell(flexcpChecked, LngRow, PI_非预期的二次手术) & "," & .Cell(flexcpChecked, LngRow, PI_麻醉并发症) & "," & .Cell(flexcpChecked, LngRow, PI_术中异物遗留) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_手术并发症) & "," & .Cell(flexcpChecked, LngRow, PI_术后出血或血肿) & "," & .Cell(flexcpChecked, LngRow, PI_手术伤口裂开) & "," & .Cell(flexcpChecked, LngRow, PI_术后深静脉血栓) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_术后生理代谢紊乱) & "," & .Cell(flexcpChecked, LngRow, PI_术后呼吸衰竭) & "," & .Cell(flexcpChecked, LngRow, PI_术后肺栓塞) & "," & .Cell(flexcpChecked, LngRow, PI_术后败血症) & "," & _
                            .Cell(flexcpChecked, LngRow, PI_术后髋关节骨折) & ")"
                End If
                gclsPros.SecdInfoRec.MoveNext
            Loop
        End With
    End If
End Sub

Private Function GetOPSOrder(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long) As Long
'功能：获取指定行手麻记录的次序
    Dim i As Long, lngOrder As Long
    
    With vsOPS
        For i = .FixedRows To LngRow
            If .TextMatrix(i, PI_手术名称) <> "" Then
                lngOrder = lngOrder + 1
            End If
        Next
    End With
    GetOPSOrder = lngOrder
End Function

Private Sub PopAllerSQL(ByRef arrSQL As Variant)
'功能：将过敏信息SQL加入数组
'SaveMedPageData的子函数
    Dim k As Integer, LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String
    
    '过敏信息
    gclsPros.MainInfoRec.Filter = "信息名='过敏药物' And 是否改变=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsAller
        With vsTmp
            '删除行以及主信息改变行需要调用删除方法
            gclsPros.SecdInfoRec.Filter = "(改变状态=" & CS_删除行 & " And 序号=" & gclsPros.MainInfoRec!序号 & ") OR (改变状态=" & CS_替换行 & " And 序号=" & gclsPros.MainInfoRec!序号 & ")": gclsPros.SecdInfoRec.Sort = "Sort": strTmp = ""
            Do While Not gclsPros.SecdInfoRec.EOF
                strTmp = strTmp & "," & gclsPros.SecdInfoRec!ID
                gclsPros.SecdInfoRec.MoveNext
            Loop
            If strTmp <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人过敏记录_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & IIf(gclsPros.FuncType = f病案首页, 4, 3) & ",'" & Mid(strTmp, 2) & "')"
            End If
            '主信息改变以及新增行需要调用插入过程
            '次级信息改变，调用更新过程
            gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态>" & CS_未改变: gclsPros.SecdInfoRec.Sort = "Sort"
            Do While Not gclsPros.SecdInfoRec.EOF
                LngRow = gclsPros.SecdInfoRec!IndexEx
                If .TextMatrix(LngRow, AI_过敏药物) <> "—" Then
    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If gclsPros.SecdInfoRec!改变状态 <> CS_更新行 Then
                        arrSQL(UBound(arrSQL)) = "zl_病人过敏记录_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & _
                                IIf(gclsPros.FuncType = f病案首页, 4, 3) & "," & ZVal(.TextMatrix(LngRow, AI_药物ID)) & ",'" & .TextMatrix(LngRow, AI_过敏药物) & "',1," & _
                                zlStr.To_Date(.TextMatrix(LngRow, AI_过敏时间), "ymd") & ",SysDate,'" & _
                                .TextMatrix(LngRow, AI_过敏反应) & "','" & .TextMatrix(LngRow, AI_过敏源编码) & "')"
                        gclsPros.AddAller = True
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_病人过敏记录_Update(" & gclsPros.SecdInfoRec!ID & "," & gclsPros.病人ID & "," & gclsPros.主页ID & "," & _
                                IIf(gclsPros.FuncType = f病案首页, 4, 3) & "," & ZVal(.TextMatrix(LngRow, AI_药物ID)) & ",'" & .TextMatrix(LngRow, AI_过敏药物) & "',1," & _
                                zlStr.To_Date(.TextMatrix(LngRow, AI_过敏时间), "ymd") & ",'" & _
                                .TextMatrix(LngRow, AI_过敏反应) & "','" & .TextMatrix(LngRow, AI_过敏源编码) & "')"
                    End If
                End If
                gclsPros.SecdInfoRec.MoveNext
            Loop
        End With
    End If
End Sub

Private Sub PopICUSQL(ByRef arrSQL As Variant)
'功能：将重症监护表格产生的SQL加入数组
'SaveMedPageData的子函数
    Dim LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim arrFields As Variant
    
    If gclsPros.MedPageSandard = ST_云南省标准 Then
        gclsPros.MainInfoRec.Filter = "信息名='监护室名称' OR 信息名='人工气道脱出' OR 信息名='重返重症医学科' OR 信息名='重返间隔时间'"
        Set rsTmp = Rec.FilterNew(gclsPros.MainInfoRec)
        rsTmp.Filter = "是否改变=1"
        If rsTmp.RecordCount > 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病案重症监护情况_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病案重症监护情况_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",1"
            arrFields = Array("监护室名称", "", "", "", "", "人工气道脱出", "重返重症医学科", "重返间隔时间")
            For LngRow = LBound(arrFields) To UBound(arrFields)
                If arrFields(LngRow) = "" Then
                    arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & ",Null"
                Else
                    rsTmp.Filter = "信息名='" & arrFields(LngRow) & "'"
                    If arrFields(LngRow) = "监护室名称" Or arrFields(LngRow) = "重返间隔时间" Then
                        arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & ",'" & rsTmp!信息现值 & "'"
                    Else
                         arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & "," & ZVal(rsTmp!信息现值 & "")
                    End If
                End If
            Next
            arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & ")"
        End If
    Else
        '重症监护记录
        gclsPros.MainInfoRec.Filter = "信息名='病案重症监护情况' And 是否改变=1"
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            Set vsTmp = gclsPros.CurrentForm.vsFlxAddICU
            With vsTmp
                '因此存在外键，只有集体删除
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病案重症监护情况_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ")"
                For LngRow = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(LngRow, UI_监护室名称)) <> "" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_病案重症监护情况_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & LngRow & ",'" & Trim(.TextMatrix(LngRow, UI_监护室名称)) & "'," & _
                                    zlStr.To_Date(.TextMatrix(LngRow, UI_进入时间), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, UI_退出时间), "ymdhm") & "," & ZVal(.TextMatrix(LngRow, UI_再入住计划)) & ",'" & .TextMatrix(LngRow, UI_再入住原因) & "')"
                    End If
                Next
            End With
        End If
    End If
End Sub

Public Sub PopDiagMatchSQL(ByRef arrSQL As Variant)
'功能：将诊断符合情况的SQL加入数组
'SaveMedPageData的子函数
    Dim arrField As Variant, arrFieldEx As Variant
    Dim i As Long
    
     '诊断符合情况
    gclsPros.MainInfoRec.Filter = "信息名='诊断符合情况' And 是否改变=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        gclsPros.SecdInfoRec.Filter = "改变状态<>" & CS_未改变 & " And 序号=" & gclsPros.MainInfoRec!序号: gclsPros.SecdInfoRec.Sort = "Sort"
        If gclsPros.SecdInfoRec.RecordCount > 0 Then
            arrField = Array(BCC_门诊与出院XY, BCC_入院与出院XY, BCC_放射与病理, BCC_临床与病理, BCC_临床与尸检, BCC_术前与术后, BCC_门诊与入院, _
                    BCC_门诊与出院ZY, BCC_入院与出院ZY, BCC_辩证, BCC_治法, BCC_方药)
            arrFieldEx = Array(1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15)
            For i = LBound(arrField) To UBound(arrField)
                gclsPros.SecdInfoRec.MoveFirst
                Do While Not gclsPros.SecdInfoRec.EOF
                    If arrField(i) = gclsPros.SecdInfoRec!IndexEx Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_诊断符合情况_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & arrFieldEx(i) & "," & IIf(gclsPros.SecdInfoRec!信息现值 & "" = "", "Null", gclsPros.SecdInfoRec!信息现值) & ")"
                    End If
                    gclsPros.SecdInfoRec.MoveNext
                Loop
            Next
        End If
    End If
End Sub

Private Sub PopStructAdressSQL(ByRef arrSQL As Variant)
'功能：将结构化地址的SQL加入数组
'SaveMedPageData的子函数
    Dim arrField As Variant
    Dim strTmp As String
    Dim i As Long
    Dim blnAdd As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngType As Long
    
    With gclsPros.CurrentForm
        If gclsPros.MedPageSandard = ST_四川省标准 Then
            arrField = Array("出生地点", "籍贯", "家庭地址", "户口地址", "联系人地址", "单位地址")
        Else
            arrField = Array("出生地点", "籍贯", "家庭地址", "户口地址", "联系人地址")
        End If
        For i = LBound(arrField) To UBound(arrField)
            gclsPros.MainInfoRec.Filter = "信息名='" & arrField(i) & "'"
            blnAdd = False
            If Not gclsPros.MainInfoRec.EOF Then
                strTmp = .padrInfo(gclsPros.MainInfoRec!Index).value省 & "," & .padrInfo(gclsPros.MainInfoRec!Index).value市 & "," & .padrInfo(gclsPros.MainInfoRec!Index).value区县
                If .padrInfo(gclsPros.MainInfoRec!Index).Items > 3 Then
                    strTmp = strTmp & "," & IIf(.padrInfo(gclsPros.MainInfoRec!Index).Items = 4, "," & .padrInfo(gclsPros.MainInfoRec!Index).value乡镇, .padrInfo(gclsPros.MainInfoRec!Index).value乡镇 & ",") & .padrInfo(gclsPros.MainInfoRec!Index).value详细地址 & "," & .padrInfo(gclsPros.MainInfoRec!Index).Code
                Else
                    strTmp = strTmp & ",,," & .padrInfo(gclsPros.MainInfoRec!Index).Code
                End If
                If Trim(Replace(strTmp, ",", "")) = "" Then
                    strTmp = ""
                Else
                    strTmp = Replace(strTmp, ",", "','")
                End If
                If gclsPros.MainInfoRec!是否改变 = 0 Then '未发生改变，且没有结构地址信息，则插入数据
                    Set rsTmp = GetStrucAddress(gclsPros.病人ID, gclsPros.主页ID, arrField(i))
                    blnAdd = rsTmp.EOF
                Else
                    blnAdd = True
                End If
                lngType = IIf(strTmp = "", 2, 1)
                If blnAdd Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(" & lngType & "," & gclsPros.病人ID & "," & gclsPros.主页ID & "," & (i + 1) & ",'" & strTmp & "')"
                End If
            End If
        Next
    End With
End Sub

Private Sub PopFeeSQL(ByRef arrSQL As Variant)
'功能：将病人费用表格产生的SQL加入数组
'SaveMedPageData的子函数
    Dim i As Long, LngRow As Long, LngCol As Long
    Dim vsTmp As VSFlexGrid

    gclsPros.MainInfoRec.Filter = "信息名='病人费用' And 是否改变=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人费用_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ")"
        Set vsTmp = gclsPros.CurrentForm.vsFees
        With vsTmp
            For i = .FixedRows * 3 To .Rows * 3 - 1
                LngRow = i \ 3: LngCol = (i Mod 3) * 2
                If .TextMatrix(LngRow, LngCol) <> "" Then '费用名非空
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人费用_insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & _
                        ",'" & GetTextByDot(.TextMatrix(LngRow, LngCol)) & "'," & Val(.TextMatrix(LngRow, LngCol + 1)) & ")"
                End If
            Next
        End With
    End If
End Sub

Public Sub PopDelicerySQL(ByRef arrSQL As Variant)
'功能：将分娩信息的SQL加入数组
'SaveMedPageData的子函数
'   1-SaveMedPageData的子函数
'   2-新生儿登记时允许分娩信息录入
        If grsDeliceryInfo Is Nothing Then Exit Sub
        '保存分娩从表信息
        grsDeliceryInfo.Filter = "类型=0 And 记录性质=1": grsDeliceryInfo.Sort = "信息名"
        Do While Not grsDeliceryInfo.EOF
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病案主页从表_首页整理(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'" & grsDeliceryInfo!信息名 & "','" & grsDeliceryInfo!信息现值 & "')"
            grsDeliceryInfo.MoveNext
        Loop
        grsBabyInfo.Filter = "记录性质=1"
        grsBabyDiag.Filter = "记录性质=1"
        '保存新生儿诊断信息
        If Not grsBabyInfo.EOF Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_分娩信息_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",0)"
        ElseIf Not grsBabyDiag.EOF Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_分娩信息_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",1)"
        End If

        Do While Not grsBabyInfo.EOF
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人分娩信息_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & grsBabyInfo!胎儿次序 & ",'" & grsBabyInfo!分娩方式 & "','" & _
                        grsBabyInfo!出生胎位 & "','" & grsBabyInfo!分娩情况 & "', '" & IIf((grsBabyInfo!出生缺陷 & "") = 0 And grsBabyInfo!出生缺陷 & "" <> "有", 0, 1) & "', '" & grsBabyInfo!婴儿性别 & "','" & grsBabyInfo!婴儿体重 & "', '" & grsBabyInfo!Apgar评分 & "',to_Date('" & grsBabyInfo!分娩时间 & "','YYYY-MM-DD HH24:MI:SS')" & ")"
            grsBabyInfo.MoveNext
        Loop
        Do While Not grsBabyDiag.EOF
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_新生儿诊断记录_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & grsBabyDiag!胎儿次序 & "," & grsBabyDiag!诊断次序 & ", " & grsBabyDiag!疾病id & ",'" & grsBabyDiag!描述信息 & "')"
            grsBabyDiag.MoveNext
        Loop
End Sub

Private Sub PopShareInfoSQL(ByRef arrSQL As Variant)
'功能：将放疗，化疗，精神药品信息的SQL加入数组，该组信息属于病案系统，住院首页在病案共享时才保存
'SaveMedPageData的子函数
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, j As Long, LngRow As Long
    Dim vsTmp As VSFlexGrid
    
    strTmp = "病案精神治疗,病案化疗记录,病案放疗记录"
    arrTmp = Split(strTmp, ",")
    strTmp = ""
    For i = LBound(arrTmp) To UBound(arrTmp)
        gclsPros.MainInfoRec.Filter = "信息名='" & arrTmp(i) & "' And 是否改变=1"
        If gclsPros.MainInfoRec.RecordCount > 0 Then
            If arrTmp(i) = "病案精神治疗" Then
                Set vsTmp = gclsPros.CurrentForm.vsSpirit
            ElseIf arrTmp(i) = "病案化疗记录" Then
                Set vsTmp = gclsPros.CurrentForm.vsChemoth
            Else
                Set vsTmp = gclsPros.CurrentForm.vsRadioth
            End If
            With vsTmp
                '删除行以及主信息改变行需要调用删除方法
                gclsPros.SecdInfoRec.Filter = "(改变状态=" & CS_删除行 & " And 序号=" & gclsPros.MainInfoRec!序号 & ") OR (改变状态=" & CS_替换行 & " And 序号=" & gclsPros.MainInfoRec!序号 & ")": gclsPros.SecdInfoRec.Sort = "Sort": strTmp = ""
                Do While Not gclsPros.SecdInfoRec.EOF
                    strTmp = strTmp & "," & gclsPros.SecdInfoRec!ID
                    gclsPros.SecdInfoRec.MoveNext
                Loop
                If strTmp <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_" & arrTmp(i) & "_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'" & Mid(strTmp, 2) & "')"
                End If
                '主信息改变以及新增行需要调用插入过程
                '次级信息改变，调用更新过程
                gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态>" & CS_未改变: gclsPros.SecdInfoRec.Sort = "Sort"
                Do While Not gclsPros.SecdInfoRec.EOF
                    LngRow = gclsPros.SecdInfoRec!IndexEx
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If gclsPros.SecdInfoRec!改变状态 <> CS_更新行 Then
                        arrSQL(UBound(arrSQL)) = "ZL_" & arrTmp(i) & "_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & LngRow
                        gclsPros.SecdInfoRec!ID = LngRow
                    Else
                        arrSQL(UBound(arrSQL)) = "ZL_" & arrTmp(i) & "_Update(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & LngRow
                    End If
                    For j = .FixedCols To .Cols - 2 '有一隐藏列因此-2
                        Select Case j
                            Case RI_放射治疗编码 'CI_化学治疗编码, SI_药物名称
                                If i = 0 Then
                                    strTmp = ZVal(.TextMatrix(LngRow, SI_药品ID)) & ",'" & .TextMatrix(LngRow, j) & "'"
                                Else
                                    strTmp = ZVal(.TextMatrix(LngRow, IIf(i = 1, CI_疾病ID, RI_疾病ID)))
                                End If
                            Case RI_开始日期, RI_结束日期 'CI_开始日期, CI_结束日期; SI_疗程 ,SI_最高日量
                                If i = 0 Then '精神药品
                                    strTmp = "'" & .TextMatrix(LngRow, j) & "'"
                                Else '放疗化疗
                                    strTmp = zlStr.To_Date(.TextMatrix(LngRow, j), "ymd")
                                End If
                            Case RI_设野部位, RI_放射剂量 'CI_疗程数,SI_特殊反应;CI_化疗方案, SI_疗效
                                '放疗放射剂量与化疗疗程数都为数字型
                                strTmp = IIf(i = 2 And j = RI_放射剂量 Or i = 1 And j = CI_疗程数, Val(.TextMatrix(LngRow, j)), "'" & .TextMatrix(LngRow, j) & "'")
                            Case RI_累计量 'CI_总量
                                strTmp = Val(.TextMatrix(LngRow, j)) & ""
                            Case Else
                                strTmp = "'" & .TextMatrix(LngRow, j) & "'"
                        End Select
                        arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & "," & strTmp
                    Next
                    arrSQL(UBound(arrSQL)) = arrSQL(UBound(arrSQL)) & ")"
                    gclsPros.SecdInfoRec.MoveNext
                Loop
            End With
        End If
    Next
End Sub
Private Sub PopOtherSQL(ByRef arrSQL As Variant)
'功能：将重症器械、医院感染、标本情况加入数组
'SaveMedPageData的子函数
    Dim LngRow As Long
    Dim vsTmp As VSFlexGrid
    Dim strTmp As String, arrTmp As Variant
    '重症器械导管使用情况
    gclsPros.MainInfoRec.Filter = "信息名='器械导管使用情况' And 是否改变=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
        With vsTmp
            '因为存在外键，只有集体删除
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_器械导管使用情况_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ")"
            For LngRow = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(LngRow, TI_器械及导管)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_器械导管使用情况_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & Val(.Cell(flexcpData, LngRow, TI_ICU类型)) & ",'" & GetTextByDot(Trim(.TextMatrix(LngRow, TI_ICU类型)), , "-") & "','" & GetTextByDot(Trim(.TextMatrix(LngRow, TI_器械及导管)), True) & "'," & _
                                                zlStr.To_Date(.TextMatrix(LngRow, TI_开始时间), "ymdhm") & "," & zlStr.To_Date(.TextMatrix(LngRow, TI_结束时间), "ymdhm") & ",'" & .TextMatrix(LngRow, TI_感染累计小时) & "')"
                End If
            Next
        End With
    End If
    '病人感染记录
    gclsPros.MainInfoRec.Filter = "信息名='病人感染记录' And 是否改变=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsInfect
        With vsTmp
            '由于序号作为主键，不采取删除新增方法会造成序号混乱
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
             arrSQL(UBound(arrSQL)) = "Zl_病人感染记录_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ")"
            For LngRow = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(LngRow, FI_感染部位)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人感染记录_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & LngRow & "," & zlStr.To_Date(.TextMatrix(LngRow, FI_确诊日期), "ymdhm") & ",'" & GetTextByDot(Trim(.TextMatrix(LngRow, FI_感染部位))) & "','" & .TextMatrix(LngRow, FI_医院感染编码) & "')"
                End If
            Next
        End With
    End If
    
    '病人病原学检查
    gclsPros.MainInfoRec.Filter = "信息名='病人病原学检查' And 是否改变=1"
    If gclsPros.MainInfoRec.RecordCount > 0 Then
        Set vsTmp = gclsPros.CurrentForm.vsSample
        With vsTmp
            '由于序号作为主键，不采取删除新增方法会造成序号混乱
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
             arrSQL(UBound(arrSQL)) = "Zl_病人病原学检查_Delete(" & gclsPros.病人ID & "," & gclsPros.主页ID & ")"
            For LngRow = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(LngRow, MI_标本)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人病原学检查_Insert(" & gclsPros.病人ID & "," & gclsPros.主页ID & "," & LngRow & ",'" & GetTextByDot(Trim(.TextMatrix(LngRow, MI_标本)), True) & "','" & GetTextByDot(.TextMatrix(LngRow, MI_病原学代码及名称), True, "-") & "'," & zlStr.To_Date(.TextMatrix(LngRow, MI_送检日期), "ymdhm") & ")"
                End If
            Next
        End With
    End If
End Sub

Private Function Get首页整理SQL(ByVal intType As Integer, ByVal arrFilds As Variant) As String
'功能：获取首页整理的存储过程调用SQL
'参数：intType=0-病人信息首页整理SQL,1-病案主页首页整理SQL
'      arrFilds=字段名的数据
'返回：SQL
    Dim strSql As String, strValue As String
    Dim i As Long, lngIdex As Long
    Dim lngMedType As Long
    
    If intType = 0 Then
        If Not gclsPros.OnlyPatiInfo And gclsPros.NoType = IT_New Or Not gclsPros.IsExistPati Then
            lngMedType = 1 '增加病人信息
        End If
        strSql = "ZL" & IIf(gclsPros.FuncType = f病案首页, 3, "") & "_病人信息_首页整理(" & IIf(gclsPros.FuncType = f病案首页, lngMedType & ",", "")
    Else
        If gclsPros.OpenMode <> EM_编辑 And Not gclsPros.Is编目 Then
            lngMedType = 1 '增加病案主页
        End If
        If Not gclsPros.OnlyPatiInfo And gclsPros.NoType = IT_New Or Not gclsPros.IsExistPati Then
            lngMedType = 1 '增加病案主页
        End If
        strSql = "ZL" & IIf(gclsPros.FuncType = f病案首页, 3, "") & "_病案主页_首页整理(" & IIf(gclsPros.FuncType = f病案首页, lngMedType & ",", "")
    End If
    
    For i = LBound(arrFilds) To UBound(arrFilds)
        strValue = ""
        Select Case Trim(arrFilds(i))
            Case ""
                strValue = ",Null"
            Case "操作员编号", "操作员姓名"
                strValue = ",'" & IIf(arrFilds(i) = "操作员编号", UserInfo.编号, UserInfo.姓名) & "'"
            Case "病人ID"
                strValue = gclsPros.病人ID & ""
            Case "主页ID", "出院科室ID", "入院科室ID"
                'arrFilds(i) & "",不知道什么原因，Decode发现第一个参数为一串数字导致Decode出错，因此拼接空串
                strValue = "," & decode(arrFilds(i) & "", "主页ID", gclsPros.主页ID, "出院科室ID", ZVal(gclsPros.出院科室ID), "入院科室ID", gclsPros.入院科室ID)
            Case "NO"
                    strValue = IIf(gclsPros.RegistNo = "", ",Null", ",'" & gclsPros.RegistNo & "'")
            Case "合同单位ID"
                If Trim(gclsPros.CurrentForm.txtAdressInfo(ADRC_单位地址).Text) <> "" Then
                    strValue = Val(gclsPros.CurrentForm.txtAdressInfo(ADRC_单位地址).Tag)
                End If
                strValue = "," & ZVal(strValue)
            Case Else
                gclsPros.MainInfoRec.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                strValue = gclsPros.MainInfoRec!信息现值 & ""
                Select Case arrFilds(i)
                    Case "出生日期", "发病时间"
                        strValue = "," & zlStr.To_Date(strValue, "ymdhm")
                    Case "入院日期", "出院日期", "确诊日期"
                        strValue = "," & zlStr.To_Date(strValue, "ymdhms")
                    Case "编目日期"
                        strValue = "," & zlStr.To_Date(strValue, "ymd")
                    Case "身高", "体重", "再入院"
                        strValue = "," & ZVal(strValue)
                    Case "复诊", "费用和", "住院天数", "新发肿瘤", "是否确诊"
                        strValue = "," & Val(strValue)
                    Case "尸检标志"
                        strValue = IIf(gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).Text = "-", ",Null", "," & Val(strValue))
                    Case "成功次数", "抢救次数", "随诊期限", "随诊标志"
                        strValue = IIf(strValue = "", ",Null", "," & Val(strValue))
                    Case "血型"
                        strValue = IIf(gclsPros.CurrentForm.cboBaseInfo(BCC_血型).Text = "-", ",Null", ",'" & strValue & "'")
                    Case "RH"
                        strValue = IIf(gclsPros.CurrentForm.cboBaseInfo(BCC_RH).Text = "-", ",Null", ",'" & strValue & "'")
                    Case Else
                        strValue = IIf(strValue = "", ",Null", ",'" & strValue & "'")
                End Select
        End Select
        If i = UBound(arrFilds) Then strValue = IIf(strValue = "", "Null", strValue) & ")"
        strSql = strSql & strValue
    Next
    Get首页整理SQL = strSql
    
End Function

Public Function CheckDateRange(ByVal strDate As String, Optional ByVal blnCheckData As Boolean) As Boolean
'功能：检查录入日期是否在入出院日期范围
'参数：strDate=待检查的日期
'      blnCheckData=true:只检查日期范围，不检查时间范围，false:检查具体时间范围
'返回：True=成功，在入出院期间 ； false=失败，不在入出院之间
'说明：入院日期为空，返回false,出院日期为空则处理为当前时间
    
    Dim DateStart As Date, dateEnd As Date
    Dim str入院时间 As String, str出院时间 As String
    Dim strFMT As String
    On Error GoTo errH
    
    CheckDateRange = False
    If Not IsDate(strDate) Then Exit Function
    Select Case Len(strDate)
        Case 10
            strFMT = "yyyy-MM-dd"
        Case 16
            strFMT = "yyyy-MM-dd hh:mm"
        Case 19
            strFMT = "yyyy-MM-dd hh:mm:ss"
        Case Else
            strFMT = "yyyy-MM-dd hh:mm"
    End Select
    '获取默认的入出院时间
    If Not IsDate(gclsPros.InTime) Then
        str入院时间 = "0"
    Else
        str入院时间 = Format(gclsPros.InTime, strFMT)
    End If
    If Not IsDate(gclsPros.OutTime) Then
        str出院时间 = "0"
    Else
        str出院时间 = Format(gclsPros.OutTime, strFMT)
    End If

    '起始时间获取
    DateStart = CDate(str入院时间)
    If DateStart = CDate(0) Then DateStart = zlDatabase.Currentdate
    '中止时间获取
    dateEnd = CDate(str出院时间)
    If dateEnd = CDate(0) Then dateEnd = zlDatabase.Currentdate
    
    '时间检查
    If blnCheckData Then
        strDate = Format(strDate, "yyyy-MM-dd")
        If CDate(strDate) >= CDate(Format(DateStart, "yyyy-MM-dd")) And CDate(strDate) <= CDate(Format(dateEnd, "yyyy-MM-dd")) Then
            CheckDateRange = True
        End If
    Else
        If CDate(strDate) >= DateStart And CDate(strDate) <= dateEnd Then
            CheckDateRange = True
        End If
    End If
    
    Exit Function
errH:
    CheckDateRange = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetCboFromRec(ByVal arrIndex As Variant, Optional ByVal intInfoType As Integer, Optional ByVal strAddBeginItems As String = "NULL", Optional ByVal strAddEndItems As String = "NULL")
'功能：将指定数据源中的数据装入指定索引的一个或多个ComboBox
'参数：arrIndex=ComboBox的Index数组，或者指定的信息名（此时为字符串，为了方便扩展该方法)
'      intInfoType=0-基础字典包，1-人员表
'      strDeFault=如果列表中存在默认标志项，则该默认值不生效
'      strAddBeginItems=是否在列表开头加入新的项，多个项目间以";"分割，默认值NULL标识不添加
'      strAddEndItems=是否在列表结尾加入新的项，多个项目间以";"分割，默认值NULL标识不添加
    '数组的处理
    Dim arrItem As Variant
    Dim i As Long, j As Long
    Dim objCboTmp As ComboBox
    Dim rsTmp As ADODB.Recordset
    '将直接传值的转变为数组
    If TypeName(arrIndex) <> "Variant()" Then
        arrIndex = Array(arrIndex)
    End If
    If TypeName(arrIndex) = "Variant()" Then
        For i = LBound(arrIndex) To UBound(arrIndex)
            If intInfoType = 0 Then
                Set objCboTmp = gclsPros.CurrentForm.cboBaseInfo(arrIndex(i))
            Else
                Set objCboTmp = gclsPros.CurrentForm.cboManInfo(arrIndex(i))
            End If
            '添加缓存记录集
            If intInfoType = 0 Then
                Set rsTmp = GetBaseCode(arrIndex(i))
            ElseIf intInfoType = 1 Then
                Set rsTmp = GetManData(arrIndex(i))
            End If
            '清除原有数据
            objCboTmp.Clear
            '添加额外添加数据
            If strAddBeginItems <> "NULL" Then
                arrItem = Split(strAddBeginItems, ",")
                For j = LBound(arrItem) To UBound(arrItem)
                    objCboTmp.AddItem arrItem(j)
                Next
            End If
            '装入数据
            If Not rsTmp.EOF Then
                If objCboTmp.Index = BCC_血型 Then
                    objCboTmp.AddItem "-"
                End If
                For j = 1 To rsTmp.RecordCount
                    If IsNull(rsTmp!编码) Then
                        objCboTmp.AddItem rsTmp!名称
                    Else
                        objCboTmp.AddItem rsTmp!编码 & "-" & Chr(13) & rsTmp!名称
                    End If
                    objCboTmp.ItemData(objCboTmp.NewIndex) = NVL(rsTmp!ID, 0)
                    If Val(rsTmp!缺省 & "") = 1 Then
                        Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                        objCboTmp.Tag = objCboTmp.NewIndex
                    End If
                    rsTmp.MoveNext
                Next
            End If
        Next
        '添加额外添加数据
        If strAddEndItems <> "NULL" Then
            arrItem = Split(strAddEndItems, ",")
            For j = LBound(arrItem) To UBound(arrItem)
                objCboTmp.AddItem arrItem(j)
                If intInfoType = 1 Then
                    objCboTmp.ItemData(objCboTmp.NewIndex) = -1
                End If
            Next
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub SetLstBoxFromRec(ByVal strlstInfos As String)
'功能：初始化LVW
'参数：strlvwInfo代表的信息命名,多个信息名字之间以逗号分割
    Dim objlstBox As ListBox
    Dim rsTmp As ADODB.Recordset
    Dim arrTmp As Variant
    Dim i As Long
    Dim blnDo As Boolean
    
    arrTmp = Split(strlstInfos, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        Select Case arrTmp(i)
            Case "感染部位"
                Set objlstBox = gclsPros.CurrentForm.lstInfectParts
            Case "感染因素"
                Set objlstBox = gclsPros.CurrentForm.lstInfection
            Case "不良事件"
                Set objlstBox = gclsPros.CurrentForm.lstAdvEvent
        End Select
        Set rsTmp = GetBaseCode(arrTmp(i))
        objlstBox.Clear
        rsTmp.Sort = "编码,名称"
        blnDo = arrTmp(i) = "不良事件" And gclsPros.Is产科
        Do While Not rsTmp.EOF
            If (rsTmp!名称 & "" = "新生儿产伤" Or rsTmp!名称 & "" = "阴道分娩产妇产伤") Then
               If blnDo Then
                   objlstBox.AddItem rsTmp!名称
                   objlstBox.ItemData(objlstBox.NewIndex) = Val(rsTmp!编码)
               End If
            Else
               objlstBox.AddItem rsTmp!名称
               objlstBox.ItemData(objlstBox.NewIndex) = Val(rsTmp!编码)
            End If
            rsTmp.MoveNext
        Loop
    Next
    objlstBox.ListIndex = -1
End Sub

Public Function SetInputRoot(ByVal intType As Integer, ByVal intSysPara As Integer, ByRef intModPara As Integer, ParamArray arrControls() As Variant) As Boolean
'说明：该函数用于系统参数与模块参数共同控制一组单选按钮，系统参数值一般为A(0或1),A+1,A+2....,模块参数为B,B+1,....系统参数为A时，模块参数起作用,且有以下条件
'           模块参数=B(系统参数=A)产生的业务效果与系统参数=A+1相同
'           模块参数=B+1(系统参数=A)产生的业务效果与系统参数=A+2相同
'功能：设置来源，可以设置模块变量西医诊断来源，中医诊断来源，过敏输入来源
'参数：intType=0-西医诊断来源设置，1-中医诊断来源，2-过敏诊断来源
'      intSysPara=系统参数，参数值为A(0或1),A+1,A+2，..，值为A时模块参数起作用
'      intModPara=模块参数
'返回：是否成功
'      intModPara=实际参数值。如系统参数为，0，1，2，模块为0，1 ，系统为0时模块起作用，此时模块参数实际值=模块参数值，当系统参数<>0，如1，模块参数实际值=系统参数-1

    Dim blnVisual As Boolean, blnEnable As Boolean
    Dim i As Long
    Dim blnAller As Boolean

    On Error GoTo errH
    '过敏输入来源，当不启用太元通时控件不可见,其余情况可见
    blnVisual = intType = 2 And gclsPros.PassType = 3 Or intType <> 2
    blnEnable = intSysPara = IIf(intType <> 2, 1, 0)
    If Not blnVisual Then intModPara = 0
    If Not blnEnable Then intModPara = intSysPara - IIf(intType <> 2, 2, 1)
    '设置控件的值以及可用性
    For i = LBound(arrControls) To UBound(arrControls)
        arrControls(i).Visible = blnVisual
        If blnVisual Then
            arrControls(i).Enabled = blnEnable And arrControls(i).Enabled
            '实际模块参数值与控件数组下标起始值一样，顺序一样
            If i = intModPara Then
                arrControls(i).Value = 1
            Else
                arrControls(i).Value = 0
            End If
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SetCtrlValues(ByVal strInfoName As String, ByVal strInfoValue As String, Optional ByVal str附加编码 As String, Optional ByVal blnMain As Boolean) As Boolean
'功能：设置控件值
'参数  strInfoName=信息名
'      strInfoValue=信息值
'      str附加编码=病案附加项目编码判断
    Dim str控件名 As String, strFMT As String
    Dim lngCount As Long, i As Long, j As Long, LngRow As Long
    Dim arrTmp As Variant, strTmp As String
    Dim vsTmp As VSFlexGrid, lstTmp As ListBox
    Dim intIndex As Integer, intIndexTmp As Integer
    Dim LngCols As Long

    On Error GoTo errH
    '问题95480
    '附加项目，可能与界面上原本存在的病案主页从表信息存在名称冲突，
    '因此在先加载病案附加项目，然后再到界面查找是否有该项信息
    If str附加编码 <> "" And gclsPros.FuncType <> f诊断选择 Then
        Set vsTmp = gclsPros.CurrentForm.vsfMain
        LngCols = 6
        With vsTmp
            For i = 0 To LngCols Step 3
                LngRow = -1: LngRow = .FindRow(strInfoName, , i)
                If LngRow >= 0 Then
                    If .TextMatrix(LngRow, i + 2) = "是否" Then
                        .Cell(flexcpChecked, LngRow, i + 1) = IIf(Val(strInfoValue) = 0, 2, 1)
                    Else
                        .TextMatrix(LngRow, i + 1) = strInfoValue
                    End If
                    Call UpdateCacheRecInfo(0, "病案项目", strInfoValue, strInfoValue, LngRow, , .TextMatrix(LngRow, i) & ";" & LngRow & ";" & i)
                    Exit For
                End If
            Next
        End With
    ElseIf str附加编码 = "" Then
        Select Case strInfoName
            Case "分娩时间", "产检次数", "胎次", "胎数", "产程时间1", "产程时间2", "产程时间3", _
                    "总产程时间", "产后出血量", "产科并发症", "会阴Ⅲ度裂伤"
                If grsDeliceryInfo Is Nothing Then Exit Function  '只需要保存
                grsDeliceryInfo.AddNew Array("信息名", "信息值", "信息现值", "类型"), Array(strInfoName, strInfoValue, strInfoValue, 0)
                grsDeliceryInfo.Update
                Exit Function  '只需要保存
            Case "特殊检查4", "特殊检查5", "特殊检查6"
                gclsPros.MainInfoRec.Filter = "信息名='特殊检查'"
            Case "CT", "PETCT", "双源CT", "X片", "B超", "超声心动图", "MRI", "同位素检查"
                If gclsPros.MedPageSandard = ST_四川省标准 Then
                    gclsPros.MainInfoRec.Filter = "信息名='特殊检查'"
                Else
                    gclsPros.MainInfoRec.Filter = "信息名='" & strInfoName & "'"
                End If
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
            End If
        Else
            If strInfoName = "住院次数" And gclsPros.FuncType = f病案首页 And gclsPros.OpenMode <> EM_新增病案 Then Exit Function
            str控件名 = gclsPros.MainInfoRec!控件名 & ""
            With gclsPros.CurrentForm
                '根据信息扩展状态
                If gclsPros.MainInfoRec!ExpState = 0 Then
                    intIndex = Val(gclsPros.MainInfoRec!Index & "")
                    Select Case str控件名
                        Case "txtSpecificInfo"
                            Select Case intIndex
                                Case SLC_婴幼儿年龄, SLC_年龄
                                    Call LoadOldData(strInfoValue, intIndex)
                                Case SLC_重症监护天, SLC_重症监护小时
                                    .txtSpecificInfo(intIndex).Text = strInfoValue
                                    .optInput(OP_ICU有).Value = 1
                                Case SLC_院内会诊, SLC_外院会诊
                                    .txtSpecificInfo(intIndex).Text = Val(strInfoValue)
                                    .chkInfo(CHK_会诊情况).Value = 1
                                Case Else
                                    .txtSpecificInfo(intIndex).Text = strInfoValue
                            End Select
                        Case "cboBaseInfo"
                            '兼容老数据
                            If blnMain And strInfoValue = "" Then '病案主页病人信息中信息为空，则设置默认值（默认值列表加载时已经设置）
                                If strInfoName = "尸检标志" Then
                                    strInfoValue = "-"
                                    If gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).ListCount >= 1 Then
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).Clear
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).AddItem "-"
                                    Else
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).Clear
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).AddItem "0-无"
                                        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).AddItem "1-有"
                                    End If
                                    Call Cbo.SeekIndex(.cboBaseInfo(intIndex), strInfoValue)
                                End If
                            Else
                                If strInfoName = "血型" Then
                                    If strInfoValue = "" Then
                                        strInfoValue = "-"
                                    ElseIf strInfoValue = "未知" Then
                                        strInfoValue = "不详" '未知 读为 不
                                    Else
                                        strInfoValue = strInfoValue
                                    End If
                                ElseIf strInfoName = "RH" Then
                                    If strInfoValue = "" Then
                                        strInfoValue = "-"
                                    ElseIf strInfoValue = "未做" Then
                                        strInfoValue = "未查" '未做 改为 未查
                                    Else
                                        strInfoValue = strInfoValue
                                    End If
                                End If
                                If strInfoName = "再入院计划天数" Or strInfoName = "尸检标志" Then
                                    .cboBaseInfo(intIndex).ListIndex = Val(strInfoValue)
                                    If Val(strInfoValue) = 0 Then strInfoValue = ""
                                Else
                                    '设置Index或控件值
                                    Call Cbo.SeekIndex(.cboBaseInfo(intIndex), strInfoValue)
                                    If .cboBaseInfo(intIndex).ListIndex = -1 And strInfoValue <> "" Then
                                        If .cboBaseInfo(intIndex).Style = 0 Then
                                            .cboBaseInfo(intIndex).Text = strInfoValue
                                        Else
                                            '病案系统以前可能定义有不规范的值
    '                                        If strInfoName = "病例分型" Or strInfoName = "去向" Then
                                            Call SetCboFromName(strInfoValue, .cboBaseInfo(intIndex), , True)
                                        End If
                                    End If
                                End If
                            End If
                            If intIndex = BCC_身份证 Then '身份证控件存储两个信息名
                                If zlCommFun.ActualLen(strInfoValue) = Len(strInfoValue) Then
                                    If Trim(zlCommFun.GetNeedName(.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                                        strInfoValue = IIf(strInfoName = "身份证号状态", "", strInfoValue)
                                        If zlStr.ActualLen(strInfoValue) > 12 And gclsPros.IsMaskID Then   '生成身份证号掩码
                                            .cboBaseInfo(intIndex).Tag = "不触发Change事件" '标记不处罚Change事件
                                            .cboBaseInfo(intIndex).Text = Mid(strInfoValue, 1, 12) & String(Len(Mid(strInfoValue, 13, 2)), "*") & Mid(strInfoValue, 15)
                                            .cboBaseInfo(intIndex).Tag = strInfoValue
                                        End If
                                    Else
                                         strInfoValue = IIf(strInfoName = "外籍身份证号", strInfoValue, "")
                                    End If
                                Else '包含中文，则为身份证号状态
                                    strInfoValue = IIf(strInfoName = "身份证号状态", zlCommFun.GetNeedName(strInfoValue), "")
                                End If
                            End If
    
                        Case "txtInfo"
                            .txtInfo(intIndex).Text = strInfoValue
                            intIndexTmp = decode(strInfoName, "退出原因", CHK_完成路径, "变异原因", CHK_变异, "会诊情况", CHK_会诊情况, -1)
                            If intIndexTmp <> -1 Then
                                If strInfoValue = IIf(strInfoName <> "会诊情况", "1", "0") Then
                                    .chkInfo(intIndexTmp).Value = Val(strInfoValue)
                                    .txtInfo(intIndex).Text = ""
                                Else
                                    If strInfoValue <> "" And strInfoName = "变异原因" Then
                                        .chkInfo(intIndexTmp).Value = 1
                                        If gclsPros.PathVCauses Then
                                            Call Cbo.SeekIndex(.cboBaseInfo(BCC_变异原因), strInfoValue)
                                            If .cboBaseInfo(BCC_变异原因).ListIndex = -1 Then
                                                .cboBaseInfo(BCC_变异原因).AddItem strInfoValue
                                                .cboBaseInfo(BCC_变异原因).ListIndex = .cboBaseInfo(BCC_变异原因).NewIndex
                                            End If
                                        End If
                                    Else
                                        .chkInfo(intIndexTmp).Value = IIf(strInfoName = "退出原因", 0, 1)
                                    End If
                                End If
                            End If
                        Case "chkInfo"
                            .chkInfo(intIndex).Value = IIf(Val(strInfoValue) = 0, 0, 1)
                            Call chkInfoClick(intIndex)
                        Case "cboManInfo"
                            If strInfoName = "编目员姓名" And (strInfoValue = "" Or gclsPros.OpenMode = EM_新增病案 Or gclsPros.OpenMode = EM_新增首页) Then
                                .cboManInfo(intIndex).Text = UserInfo.姓名
                            Else
                                .cboManInfo(intIndex).Text = strInfoValue
                            End If
                        Case "mskDateInfo"
                            '时间按控件的MASK值初始化
                            strFMT = .mskDateInfo(intIndex).Mask
                            If gclsPros.FuncType = f医生首页 And intIndex = DC_出生日期 Then
                                If Format(strInfoValue, "HH:MM") = "00:00" Then
                                    .mskDateInfo(intIndex).Mask = "####-##-##"
                                    .mskDateInfo(intIndex).Tag = "####-##-##"
                                    strFMT = .mskDateInfo(intIndex).Mask
                                 Else
                                    .mskDateInfo(intIndex).Mask = "####-##-## ##:##"
                                    .mskDateInfo(intIndex).Tag = "####-##-## ##:##"
                                    strFMT = .mskDateInfo(intIndex).Mask
                                End If
                            End If
                            If IsDate(strInfoValue) Then
                                strInfoValue = Format(strInfoValue, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                            Else
                                strInfoValue = Replace(strFMT, "#", "_")
                            End If
                            .mskDateInfo(intIndex).Text = strInfoValue
                            If Not IsDate(strInfoValue) Then strInfoValue = ""
                            If strInfoValue = "" And intIndex = DC_编目日期 Then
                                .mskDateInfo(intIndex).Text = Format(zlDatabase.Currentdate, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                            End If
                            .txtDateInfo(intIndex).Text = .mskDateInfo(intIndex).Text
                            If intIndex = DC_出生日期 And gclsPros.FuncType = f病案首页 Then
                                If Trim(.mskDateInfo(intIndex).Text) = "____-__-__ __:__" Then
                                    Call SetCtrlLocked(.mskDateInfo(intIndex), True)
                                    Call SetCtrlLocked(.txtDateInfo(intIndex), True)
                                End If
                            End If
                        Case "txtAdressInfo"
                            Call SetPatiAddress(intIndex, strInfoName, strInfoValue)
                        Case "cboSpecificInfo"
                            Call Cbo.SeekIndex(.cboSpecificInfo(intIndex), strInfoValue)
                            If .cboSpecificInfo(intIndex).Style = 0 And .cboSpecificInfo(intIndex).ListIndex = -1 Then
                               .cboSpecificInfo(intIndex).Text = strInfoValue
                            End If
                        Case "lstInfection", "lstAdvEvent", "lstInfectParts"
                            If strInfoName = "感染因素" Then
                                Set lstTmp = .lstInfection
                            ElseIf strInfoName = "不良事件" Then
                                Set lstTmp = .lstAdvEvent
                            ElseIf strInfoName = "感染部位" Then
                                Set lstTmp = .lstInfectParts
                            End If
                            If InStr(strInfoValue, ",") > 0 Then
                                strInfoValue = Replace(strInfoValue, ",", "|") '将逗号分割符号转换为“|”
                            End If
                            arrTmp = Split(strInfoValue, "|")
                            For j = 0 To lstTmp.ListCount - 1
                                For i = LBound(arrTmp) To UBound(arrTmp)
                                    If lstTmp.ItemData(j) = arrTmp(i) Then
                                        lstTmp.Selected(j) = True: Exit For
                                    End If
                                Next
                            Next
                            lstTmp.ListIndex = -1
                    End Select
                    gclsPros.MainInfoRec.Update "信息原值", strInfoValue
                ElseIf gclsPros.MainInfoRec!ExpState = ES_初始扩展 Then
                    If str控件名 <> "vsTSJC" Then
                        gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                        gclsPros.SecdInfoRec.Sort = "Sort"
                    End If
                    Select Case strInfoName
                        Case "昏迷时间"
                            '保存格式:入院前(天，小时,分钟)|入院后(天，小时,分钟)
                            strTmp = Replace(strInfoValue, "|", ",")
                            strTmp = strTmp & ",,,,,"
                            arrTmp = Split(strTmp, ",")
                            For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                .txtSpecificInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = arrTmp(i)
                                gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(arrTmp(i), arrTmp(i))
                                gclsPros.SecdInfoRec.MoveNext
                            Next
                            gclsPros.MainInfoRec.Update "信息原值", strInfoValue
                        Case "转科记录"
                            strTmp = strInfoValue & ",,,,,,"
                            arrTmp = Split(strTmp, ",")
                            If str控件名 = "txtInfo" Then
                                For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                    .txtInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = arrTmp(i)
                                    gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(arrTmp(i), arrTmp(i))
                                    gclsPros.SecdInfoRec.MoveNext
                                Next
                            Else
                                For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                    .vsTransfer.TextMatrix(DR_转科科室, Val(gclsPros.SecdInfoRec!IndexEx & "")) = arrTmp(i)
                                    gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(arrTmp(i), arrTmp(i))
                                    gclsPros.SecdInfoRec.MoveNext
                                Next
                            End If
                            gclsPros.MainInfoRec.Update "信息原值", strInfoValue
                        Case "转科时间"
                            strTmp = strInfoValue & ",,,,,,"
                            arrTmp = Split(strTmp, ",")
                            If str控件名 = "txtInfo" Then
                                For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                    .txtInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = zlStr.FullDate(arrTmp(i), True)
                                    gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(arrTmp(i), arrTmp(i))
                                    gclsPros.SecdInfoRec.MoveNext
                                Next
                            Else
                                For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                    .vsTransfer.TextMatrix(DR_转科时间, Val(gclsPros.SecdInfoRec!IndexEx & "")) = zlStr.FullDate(arrTmp(i), True)
                                    gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(arrTmp(i), arrTmp(i))
                                    gclsPros.SecdInfoRec.MoveNext
                                Next
                            End If
                            gclsPros.MainInfoRec.Update "信息原值", strInfoValue
                        Case "发病时间"
                            For i = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                                strFMT = .mskDateInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Mask
                                If IsDate(strInfoValue) Then
                                    strTmp = Format(strInfoValue, decode(strFMT, "####-##-##", "yyyy-MM-dd", "##:##", "HH:mm"))
                                    If strTmp = "00:00" Then strTmp = Replace(strFMT, "#", "_")
                                Else
                                    strTmp = Replace(strFMT, "#", "_")
                                End If
                                .mskDateInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = strTmp
                                .txtDateInfo(Val(gclsPros.SecdInfoRec!IndexEx & "")).Text = strTmp
                                If Not IsDate(strTmp) Then strTmp = ""
                                gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(strTmp, strTmp)
                                gclsPros.SecdInfoRec.MoveNext
                            Next
                            gclsPros.MainInfoRec.Update "信息原值", strInfoValue
                        Case "31天内再住院"
                            .optInput(OP_再住院无).Value = strInfoValue = ""
                            .optInput(OP_再住院有).Value = strInfoValue <> ""
                            .txtInfo(GC_31天内再住院).Text = strInfoValue
                            Call SetCtrlLocked(.txtInfo(GC_31天内再住院), strInfoValue = "", True)
                            For i = 1 To gclsPros.SecdInfoRec.RecordCount
                                strTmp = decode(Val(gclsPros.SecdInfoRec!IndexEx & ""), OP_再住院无, IIf(strInfoValue = "", 1, 0), OP_再住院有, IIf(strInfoValue = "", 0, 1), strInfoValue)
                                gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(strTmp, strTmp)
                                gclsPros.SecdInfoRec.MoveNext
                            Next
                            gclsPros.MainInfoRec.Update "信息原值", strInfoValue
                        Case "复诊"
                            .optState(OP_初诊).Value = Val(strInfoValue) = 0
                            .optState(OP_复诊).Value = Val(strInfoValue) <> 0
                            For i = 1 To gclsPros.SecdInfoRec.RecordCount
                                strTmp = IIf(Val(gclsPros.SecdInfoRec!IndexEx & "") = OP_复诊, IIf(Val(strInfoValue) <> 0, 1, 0), IIf(Val(strInfoValue) <> 0, 0, 1))
                                gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(strTmp, strTmp)
                                gclsPros.SecdInfoRec.MoveNext
                            Next
                        Case Else
                            If str控件名 = "vsTSJC" Then
                                If strInfoName Like "特殊检查*" And gclsPros.MedPageSandard <> ST_四川省标准 Then
                                    intIndex = Val(Mid(strInfoName, 5, 1)) - 4
                                    strTmp = strInfoValue
                                Else
                                    intIndex = decode(strInfoName, "CT", TR_CT, "PETCT", TR_PETCT, "双源CT", TR_双源CT, _
                                                "X片", TR_X片, "B超", TR_B超, "超声心动图", TR_超声心动图, "MRI", TR_MRI, "同位素检查", TR_同位素检查, -1)
                                    strTmp = decode(Val(strInfoValue), 1, "1-阳性", 2, "2-阴性", 3, "3-未做", "")
                                End If
                                If intIndex <> -1 Then
                                strInfoValue = strTmp
                                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And IndexEx=" & intIndex
                                    .vsTSJC.TextMatrix(intIndex, 1) = strTmp
                                    .vsTSJC.Cell(flexcpData, intIndex, 1) = strTmp
                                    gclsPros.SecdInfoRec.Update Array("信息原值", "主信息原值"), Array(strInfoValue, strInfoValue)
                                End If
                            End If
                    End Select
                    If str控件名 <> "vsTSJC" Then
                        gclsPros.MainInfoRec.Update "信息原值", strInfoValue
                    End If
                ElseIf gclsPros.MainInfoRec!ExpState = 2 Then
                '加载时处理
                End If
            End With
        End If
    End If
    SetCtrlValues = True
    Exit Function
errH:
    Debug.Print "SetCtrlValues:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CacheCtrlValues()
'功能：缓存界面控件值，控件对应信息一般不扩展或次级扩展
    Dim str控件名 As String, strInfoValue As String, strInfoName As String
    Dim blnEnable As Boolean
    Dim lstTmp As ListBox
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim vsTmp As VSFlexGrid
    Dim intIndex As Integer
    Dim LngCols As Long

    On Error GoTo errH
    With gclsPros.CurrentForm
        '不扩展信息搜集缓存
        gclsPros.MainInfoRec.Filter = "ExpState=0": gclsPros.MainInfoRec.Sort = "序号"
        For i = 1 To gclsPros.MainInfoRec.RecordCount
            strInfoValue = ""
            str控件名 = gclsPros.MainInfoRec!控件名
            strInfoName = gclsPros.MainInfoRec!信息名
            Select Case str控件名
                Case "txtInfo"
                    Select Case strInfoName
                        Case "退出原因", "变异原因"
                            If .chkInfo(CHK_进入路径).Value = 1 Then
                                intIndex = IIf(strInfoName = "退出原因", CHK_完成路径, CHK_变异)
                                If .chkInfo(intIndex).Value = 1 Then
                                    strInfoValue = IIf(strInfoName = "退出原因", "1", .txtInfo(gclsPros.MainInfoRec!Index).Text)
                                    '没有填写变异原因，则保存1
                                    If strInfoValue = "" And strInfoName = "变异原因" Then strInfoValue = "1"
                                Else
                                    strInfoValue = IIf(strInfoName = "退出原因", .txtInfo(gclsPros.MainInfoRec!Index).Text, "")
                                End If
                            End If
                        Case "会诊情况"
                            If .chkInfo(CHK_会诊情况).Value = 1 Then
                                strInfoValue = .txtInfo(gclsPros.MainInfoRec!Index).Text
                            Else
                                strInfoValue = "0"
                            End If
                        Case "联系人附加信息"
                            If .txtInfo(gclsPros.MainInfoRec!Index).Visible Then
                                strInfoValue = .txtInfo(gclsPros.MainInfoRec!Index).Text
                            End If
                        Case Else
                            strInfoValue = .txtInfo(gclsPros.MainInfoRec!Index).Text
                    End Select
                Case "txtSpecificInfo"
                    strInfoValue = .txtSpecificInfo(gclsPros.MainInfoRec!Index).Text
                    Select Case strInfoName
                        Case "年龄"
                            '不知为什么用户环境下在西医诊断页判断.cboSpecificInfo(gclsPros.MainInfoRec!Index).Visible为False
                            If strInfoValue <> "" Then strInfoValue = strInfoValue & IIf(IsNumeric(strInfoValue), .cboSpecificInfo(gclsPros.MainInfoRec!Index).Text, "")
                        Case "不足周岁年龄"
                            If strInfoValue <> "" Then
                                If .cboSpecificInfo(gclsPros.MainInfoRec!Index).Text = "月" Then
                                    strInfoValue = strInfoValue & IIf(IsNumeric(strInfoValue), "月" & Trim(.txtSpecificInfo(SLC_婴幼儿年龄_DAY)) & "天", "")
                                Else
                                    strInfoValue = strInfoValue & IIf(IsNumeric(strInfoValue), .cboSpecificInfo(gclsPros.MainInfoRec!Index).Text, "")
                                End If
                            End If
                        Case "随诊期限"
                            If .chkInfo(CHK_随诊).Value = 1 Then
                                strInfoValue = IIf(Val(strInfoValue) <> 0, Val(strInfoValue), "")
                            Else
                                strInfoValue = ""
                            End If
                        Case "院内会诊", "外院会诊"
                            If .chkInfo(CHK_会诊情况).Value = 0 Then
                                strInfoValue = ""
                            End If
                    End Select
                Case "chkInfo"
                    strInfoValue = .chkInfo(gclsPros.MainInfoRec!Index).Value
                    If strInfoValue = "0" Then strInfoValue = ""
                    If strInfoName = "随诊标志" Then
                        If strInfoValue = "1" Then
                            strInfoValue = decode(zlStr.NeedName(.cboSpecificInfo(SLC_随诊期限).Text), "月", 1, "年", 2, "周", 3, "天", 4, "终身", 9)
                        End If
                    End If
                Case "cboBaseInfo"
                    strInfoValue = .cboBaseInfo(gclsPros.MainInfoRec!Index).Text
                    If strInfoName = "病例分型" Or strInfoName = "生育状况" Then
                        If InStr(strInfoValue, "-") > 0 Then  '如果是规范的数据，则只存编码
                            strInfoValue = Mid(strInfoValue, 1, InStr(strInfoValue, "-") - 1)
                        End If
                    ElseIf strInfoName = "输血反应" Then
                        strInfoValue = IIf(.cboBaseInfo(gclsPros.MainInfoRec!Index).ListIndex = -1, "", .cboBaseInfo(gclsPros.MainInfoRec!Index).ListIndex)
                    ElseIf strInfoName = "再入院计划天数" Or strInfoName = "尸检标志" Then
                        If .cboBaseInfo(gclsPros.MainInfoRec!Index).ListIndex > 0 Then
                            strInfoValue = .cboBaseInfo(gclsPros.MainInfoRec!Index).ListIndex
                        Else
                            strInfoValue = ""
                        End If
                    ElseIf Val(gclsPros.MainInfoRec!Index & "") = BCC_身份证 Then '身份证控件存储两个信息名
                        If zlCommFun.ActualLen(strInfoValue) = Len(strInfoValue) Then
                            If Trim(zlCommFun.GetNeedName(.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                                If gclsPros.FuncType = f医生首页 And strInfoValue <> "" Then '由于掩码的存在，所以取Tag
                                    strInfoValue = .cboBaseInfo(gclsPros.MainInfoRec!Index).Tag
                                End If
                                strInfoValue = IIf(strInfoName = "身份证号", strInfoValue, "")
                            Else
                                strInfoValue = IIf(strInfoName = "外籍身份证号", strInfoValue, "")
                            End If
                        Else '包含中文，则为身份证号状态
                            strInfoValue = IIf(strInfoName = "身份证号状态", zlCommFun.GetNeedName(strInfoValue), "")
                        End If
                    Else
                        strInfoValue = zlStr.NeedName(strInfoValue)
                    End If
                Case "cboManInfo"
                    strInfoValue = zlStr.NeedName(.cboManInfo(gclsPros.MainInfoRec!Index).Text)
                Case "txtAdressInfo"
                    On Error Resume Next
                    strInfoValue = .padrInfo(gclsPros.MainInfoRec!Index).Value
                    If Err.Number <> 0 Then
                        Err.Clear: strInfoValue = .txtAdressInfo(gclsPros.MainInfoRec!Index).Text
                    Else
                        If Not .padrInfo(gclsPros.MainInfoRec!Index).Visible Then
                            strInfoValue = .txtAdressInfo(gclsPros.MainInfoRec!Index).Text
                        End If
                    End If
                    On Error GoTo errH
                Case "mskDateInfo"
                    strInfoValue = .mskDateInfo(gclsPros.MainInfoRec!Index).Text
                    If Not IsDate(strInfoValue) Then strInfoValue = ""
                Case "cboSpecificInfo"
                    strInfoValue = .cboSpecificInfo(gclsPros.MainInfoRec!Index).Text
                Case "lstInfection", "lstAdvEvent", "lstInfectParts"
                    If strInfoName = "感染因素" Then
                        Set lstTmp = .lstInfection
                    ElseIf strInfoName = "不良事件" Then
                        Set lstTmp = .lstAdvEvent
                    ElseIf strInfoName = "感染部位" Then
                        Set lstTmp = .lstInfectParts
                    End If
                    For j = 0 To lstTmp.ListCount - 1
                        If lstTmp.Selected(j) = True Then
                            strInfoValue = strInfoValue & "|" & lstTmp.ItemData(j)
                        End If
                    Next
                    If strInfoValue <> "" Then
                        strInfoValue = Mid(strInfoValue, 2)
                    End If
            End Select
            gclsPros.MainInfoRec.Update "信息现值", strInfoValue
            gclsPros.MainInfoRec.MoveNext
        Next
        '次级扩展信息搜集缓存
        gclsPros.MainInfoRec.Filter = "ExpState=1": gclsPros.MainInfoRec.Sort = "序号"
        For i = 1 To gclsPros.MainInfoRec.RecordCount
            str控件名 = gclsPros.MainInfoRec!控件名 & ""
            strInfoName = gclsPros.MainInfoRec!信息名
            strInfoValue = ""
            Select Case strInfoName
                Case "昏迷时间"
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 1 To gclsPros.SecdInfoRec.RecordCount
                        strTmp = .txtSpecificInfo(gclsPros.SecdInfoRec!IndexEx).Text
                        Call gclsPros.SecdInfoRec.Update(Array("信息现值", "主信息现值"), Array(strTmp, strTmp))
                        strInfoValue = strInfoValue & IIf(j = 4, "|", ",") & strTmp
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    strInfoValue = Mid(strInfoValue, 2)
                    gclsPros.MainInfoRec.Update "信息现值", strInfoValue
                Case "转科记录"
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 1 To gclsPros.SecdInfoRec.RecordCount
                        If str控件名 = "txtInfo" Then
                            strTmp = .txtInfo(gclsPros.SecdInfoRec!IndexEx).Text
                        Else
                            strTmp = .vsTransfer.TextMatrix(DR_转科科室, gclsPros.SecdInfoRec!IndexEx)
                        End If
                        Call gclsPros.SecdInfoRec.Update(Array("信息现值", "主信息现值"), Array(strTmp, strTmp))
                        strInfoValue = strInfoValue & "," & strTmp
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    If strInfoValue <> "" Then strInfoValue = Mid(strInfoValue, 2)
                    gclsPros.MainInfoRec.Update "信息现值", strInfoValue
                Case "转科时间"
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 1 To gclsPros.SecdInfoRec.RecordCount
                        If str控件名 = "txtInfo" Then
                            strTmp = Format(.txtInfo(gclsPros.SecdInfoRec!IndexEx).Text, "yyyyMMddHHmm")
                        Else
                            strTmp = Format(.vsTransfer.TextMatrix(DR_转科时间, gclsPros.SecdInfoRec!IndexEx), "yyyyMMddHHmm")
                        End If
                        Call gclsPros.SecdInfoRec.Update(Array("信息现值", "主信息现值"), Array(strTmp, strTmp))
                        strInfoValue = strInfoValue & "," & strTmp
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    If strInfoValue <> "" Then strInfoValue = Mid(strInfoValue, 2)
                    gclsPros.MainInfoRec.Update "信息现值", strInfoValue
                Case "发病时间"
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 0 To gclsPros.SecdInfoRec.RecordCount - 1
                        If j = 0 Then
                            strInfoValue = .mskDateInfo(gclsPros.SecdInfoRec!IndexEx).Text
                            If Not IsDate(strInfoValue) Then strInfoValue = ""
                            Call gclsPros.SecdInfoRec.Update(Array("信息现值", "主信息现值"), Array(strInfoValue, strInfoValue))
                        Else
                            If strInfoValue <> "" Then
                                strTmp = .mskDateInfo(gclsPros.SecdInfoRec!IndexEx).Text
                                If IsDate(strTmp) Then
                                    Call gclsPros.SecdInfoRec.Update(Array("信息现值", "主信息现值"), Array(strTmp, strTmp))
                                    strInfoValue = strInfoValue & " " & strTmp
                                End If
                            End If
                        End If
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    gclsPros.MainInfoRec.Update "信息现值", strInfoValue
                Case "31天内再住院"
                        gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                        gclsPros.SecdInfoRec.Sort = "Sort"
                        strInfoValue = .txtInfo(GC_31天内再住院).Text
                        If .optInput(OP_再住院无).Value = 1 Then strInfoValue = ""
                        For j = 1 To gclsPros.SecdInfoRec.RecordCount
                            strTmp = decode(Val(gclsPros.SecdInfoRec!IndexEx & ""), OP_再住院无, IIf(strInfoValue = "", 1, 0), OP_再住院有, IIf(strInfoValue = "", 0, 1), strInfoValue)
                            Call gclsPros.SecdInfoRec.Update(Array("信息现值", "主信息现值"), Array(strTmp, strTmp))
                            gclsPros.SecdInfoRec.MoveNext
                        Next
                        gclsPros.MainInfoRec.Update "信息现值", strInfoValue
                Case "复诊"
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 1 To gclsPros.SecdInfoRec.RecordCount
                        strInfoValue = IIf(Val(gclsPros.SecdInfoRec!IndexEx & "") = OP_复诊, IIf(.optState(OP_复诊).Value, 1, 0), IIf(.optState(OP_初诊).Value, 1, 0))
                        gclsPros.SecdInfoRec.Update Array("信息现值", "主信息现值"), Array(strInfoValue, strInfoValue)
                        gclsPros.SecdInfoRec.MoveNext
                    Next
                    gclsPros.MainInfoRec.Update "信息现值", IIf(.optState(OP_复诊).Value, 1, 0)
                Case "特殊检查"
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号
                    gclsPros.SecdInfoRec.Sort = "Sort"
                    For j = 0 To .vsTSJC.Rows - 1
                        strInfoValue = .vsTSJC.TextMatrix(gclsPros.SecdInfoRec!IndexEx, 1)
                        Call gclsPros.SecdInfoRec.Update(Array("信息现值", "主信息现值"), Array(strInfoValue, strInfoValue))
                        gclsPros.SecdInfoRec.MoveNext
                    Next
            End Select
            gclsPros.MainInfoRec.MoveNext
        Next
    End With

    '扩展信息搜集
    gclsPros.MainInfoRec.Filter = "信息名='病案项目'"
    If Not gclsPros.MainInfoRec.EOF Then
        Set vsTmp = gclsPros.CurrentForm.vsfMain
        LngCols = 6
        With vsTmp
            For i = .FixedRows To .Rows - 1
                For j = 0 To LngCols Step 3
                    If .TextMatrix(i, j) <> "" Then
                        If .TextMatrix(i, j + 2) = "是否" Then
                            strInfoValue = IIf(.Cell(flexcpChecked, i, j + 1) = 2, "", 1)
                        Else
                            strInfoValue = .TextMatrix(i, j + 1)
                        End If
                        Call UpdateCacheRecInfo(1, "病案项目", strInfoValue, strInfoValue, i, , .TextMatrix(i, j) & ";" & i & ";" & j)
                    End If
                Next
            Next
        End With
    End If
    Exit Sub
errH:
    Debug.Print "CacheCtrlValues:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub LoadOldData(ByVal strOld As String, Optional ByVal intIndex As Integer)
'功能:将数据库中保存的年龄按规范的格式加载到界面,不规范的原样显示
'参数：strOld =年龄字符串
'      intIndex=加载年龄的控件索引值值
    Dim strTmp As String, lngIdx As Long
    Dim objTxt As TextBox, objCbo As ComboBox
    Dim arrTmp As Variant
    Dim i As Long

    If Trim(strOld) = "" Then Exit Sub
    If intIndex = SLC_年龄 Then
        strTmp = "岁,月,天,小时,分钟"
    ElseIf intIndex = SLC_婴幼儿年龄 Then
        strTmp = "月,天,小时,分钟"
    End If
    arrTmp = Split(strTmp, ",")

    strTmp = strOld
    For i = LBound(arrTmp) To UBound(arrTmp)
        If intIndex = SLC_婴幼儿年龄 And strTmp Like "*月*天" Then
            strTmp = Replace(strTmp, "天", "")
            gclsPros.CurrentForm.txtSpecificInfo(SLC_婴幼儿年龄_DAY).Text = Split(strTmp, "月")(1)
            strTmp = Split(strTmp, "月")(0)
            lngIdx = i
            Exit For
        ElseIf InStr(strOld, arrTmp(i)) > 0 Then
            If InStr(strOld, arrTmp(i)) + Len(arrTmp(i)) - 1 = Len(strOld) Then
                strTmp = Mid(strOld, 1, InStr(strOld, arrTmp(i)) - 1)
                lngIdx = i
            End If
            Exit For
        End If
    Next

    'IsNumeric("")=False,因此连接上字符串1
    If Not IsNumeric(strTmp & "1") Then
        lngIdx = -1
        strTmp = strOld
    End If
    
    '不知为什么用户环境下在西医诊断页判断objCbo.Visible为False
    Set objTxt = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    Set objCbo = gclsPros.CurrentForm.cboSpecificInfo(intIndex)
    objTxt.Text = strTmp

    If lngIdx = -1 Then
        objCbo.Visible = False
        objCbo.Tag = "年龄"
        objCbo.ListIndex = -1
        If objCbo.Container.Name = "fraCbo" Then
            objCbo.Container.Visible = False
        End If
        
        If intIndex = SLC_年龄 Then
            If gclsPros.FuncType = f病案首页 Then
                objTxt.Width = 1250
            Else
                objTxt.Width = 1150
            End If
        ElseIf intIndex = SLC_婴幼儿年龄 Then
            objTxt.Width = 1250
        End If
    Else
        If objCbo.Visible = False Then
            objCbo.Visible = True
            objCbo.Tag = ""
            If objCbo.Container.Name = "fraCbo" Then
                objCbo.Container.Visible = True
            End If
            
            If intIndex = SLC_年龄 Then
                If gclsPros.FuncType = f病案首页 Then
                    objTxt.Width = 450
                Else
                    objTxt.Width = 360
                End If
            ElseIf intIndex = SLC_婴幼儿年龄 Then
                objTxt.Width = 360
            End If
        End If
        objCbo.ListIndex = lngIdx
    End If
End Sub

Public Sub SetPageVisible()
'功能：设置页面可见性
    Dim i As Long
    With gclsPros.CurrentForm
        '非中医科不显示中医诊断
        If gclsPros.PatiType = PF_住院 Then
            For i = .PicPage.LBound To .PicPage.UBound
                .PicPage(i).Tag = "true"
            Next
             '没有编辑中医权限的不显示中医
            .PicPage(PIC_中医诊断).Tag = IIf(gclsPros.Have中医, "true", "false")
            .PicPage(PIC_中医诊断情况).Tag = IIf(gclsPros.Have中医, "true", "false")
            Select Case gclsPros.MedPageSandard
                Case ST_湖南省标准
                    .PicPage(PIC_抗精神病).Tag = "false"
                    .PicPage(PIC_重症监护).Tag = "false"
                Case ST_云南省标准
                    .PicPage(PIC_抗精神病).Tag = "false"
                    .PicPage(PIC_重症监护).Tag = "false"
                Case ST_四川省标准
                    .PicPage(PIC_抗精神病).Tag = "false"
            End Select
            If gclsPros.FuncType = f医生首页 Then
                .PicPage(PIC_住院费用).Tag = "false"
            End If
            If Not gclsPros.ReadPages Then  '病案不共享安装时，不显示放疗与化疗以及精神药品
                .PicPage(PIC_化疗信息).Tag = "false"
                .PicPage(PIC_放疗记录).Tag = "false"
                If gclsPros.MedPageSandard = ST_卫生部标准 Then
                    .PicPage(PIC_抗精神病).Tag = "false"
                End If
            End If
            For i = .PicPage.LBound To .PicPage.UBound
                If .PicPage(i).Tag = "true" Then
                    .PicPage(i).Visible = True
                ElseIf .PicPage(i).Tag = "false" Then
                    .PicPage(i).Visible = False
                End If
            Next
            '设置导航目录
            Call SetMainDirectory
        End If
    End With
End Sub

Public Function SetSignature() As Boolean
'功能：根据当前病人的医师及签名情况，确定签名及界面数据的可编辑性
'返回：界面是否已签名只读不能编辑
    Static rsTmp As ADODB.Recordset
    Dim intCurr As Integer, intHave As Integer
    Dim strSql As String, blnReadOnly As Boolean
    Dim i As Integer, j As Integer
    Dim strTmp As String
    '说明：arrInfos，arrManIdxs，arrSgnIdxs三个数组的元素一一对应，人员级别从低到高
    Dim arrInfos() As Variant '各类签名的信息名
    Dim arrManIdxs() As Variant '签名人员下拉列表的Index
    Dim arrSgnIdxs() As Variant '签名按钮的Index
    '初始化签名相关界面
    blnReadOnly = False: intCurr = -1: intHave = -1
    arrInfos = Array("住院医师签名", "主治医师签名", "主任医师签名", "科主任签名")
    arrManIdxs = Array(MC_住院医师, MC_主治医师, MC_主任或副主任, MC_科主任)
    arrSgnIdxs = Array(SL_住院医师, SL_主治医师, SL_主任医师, SL_科主任)

    On Error GoTo errH
    With gclsPros.CurrentForm
        For i = LBound(arrManIdxs) To UBound(arrManIdxs)
            .cboManInfo(arrManIdxs(i)).ForeColor = .ForeColor: .lblManInfo(arrManIdxs(i)).ForeColor = .ForeColor
            .cboManInfo(arrManIdxs(i)).Locked = False:  .cboManInfo(arrManIdxs(i)).BackColor = vbWindowBackground
            .cmdSign(arrSgnIdxs(i)).Caption = "签名"
            If zlStr.NeedName(.cboManInfo(arrManIdxs(i)).Text) = UserInfo.姓名 Then
                intCurr = i
                .cmdSign(arrSgnIdxs(i)).Enabled = Not gclsPros.Is护士站
            Else
                .cmdSign(arrSgnIdxs(i)).Enabled = False
            End If
            gclsPros.AuxiInfo.Filter = "信息名='" & arrInfos(i) & "'"
            If Not gclsPros.AuxiInfo.EOF Then
                intHave = i
                '已签名用蓝色字表示
                .cboManInfo(arrManIdxs(i)).ForeColor = vbBlue: .lblManInfo(arrManIdxs(i)).ForeColor = vbBlue
                .cmdSign(arrSgnIdxs(i)).Caption = "取消"
                '签名按钮可操作状态
                If gclsPros.AuxiInfo!信息值 & "" = UserInfo.姓名 Then
                    .cmdSign(arrSgnIdxs(i)).Enabled = Not gclsPros.Is护士站
                Else '非自已的签名不能取消
                    .cmdSign(arrSgnIdxs(i)).Enabled = False
                End If
            End If
        Next

        If intHave >= 0 Then
            '涉及签名的项都不允许再更改,不然权限混乱
            For i = LBound(arrManIdxs) To UBound(arrManIdxs)
                .cboManInfo(arrManIdxs(i)).Locked = True: .cboManInfo(arrManIdxs(i)).BackColor = vbButtonFace
                '低级别签名不能变更
                If i < intHave Then
                    .cmdSign(arrSgnIdxs(i)).Enabled = False
                End If
            Next
        End If

        '如果当前人员签名级别不高于已签名级别，则不可编辑
        If intCurr <= intHave And intHave >= 0 Then
            blnReadOnly = True
        End If
    End With
    SetSignature = blnReadOnly

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CacheLoadVsDiagData(ByRef vsDiagInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal strDiagType As String, Optional ByVal blnOnlyCache As Boolean, Optional ByVal intMaxDiagSource As Integer)
'功能：将诊断加载到表格中并且缓存
'参数：vsDiagInput=需要加载诊断的表格
'      rsInput=读取的诊断记录集
'      strDiagType=诊断类型字符串，各类型以逗号分割
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数

    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, j As Long, k As Long, LngRow As Long
    Dim bln分化程度 As Boolean
    Dim bln西医 As Boolean
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim blnFreeDiag As Boolean
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim blnGet附码 As Boolean

    blnGet附码 = gclsPros.GetExtraCode
    On Error GoTo errH
    With vsDiagInput
        bln西医 = vsDiagInput.Name = "vsDiagXY"
        '加载诊断
        If Not blnOnlyCache Then
            arrTmp = Split(strDiagType, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                Call FilterDiagByType(rsInput, Val(arrTmp(i)), intMaxDiagSource) '过滤诊断
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
                                                        If .TextMatrix(LngRow, DI_诊断类型) <> "出院诊断" And Val(.TextMatrix(LngRow, DI_诊断分类)) = 3 Then
                                .Cell(flexcpData, LngRow, DI_诊断类型) = "其他诊断"
                            End If
                        End If

                        If gclsPros.FuncType = f诊断选择 Then
                            If InStr("," & gclsPros.DiagRowIDs & ",", "," & rsInput!ID & ",") > 0 Then
                                .TextMatrix(LngRow, DI_关联) = 1
                            End If
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
                            .TextMatrix(LngRow, DI_诊断编码) = IIf(Not IsNull(rsInput!疾病id), rsInput!疾病编码 & "", rsInput!诊断编码 & "")
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
                        If gclsPros.FuncType = f病案首页 Then
                            .TextMatrix(LngRow, DI_诊断编码) = rsInput!疾病编码 & ""
                            If Not IsNull(rsInput!疾病id) Then '用来判断诊断编码与诊断名称的一致性
                                .Cell(flexcpData, LngRow, DI_诊断描述) = rsInput!疾病名称 & ""
                                If Not gclsPros.CNIndent Or .TextMatrix(LngRow, DI_诊断描述) = "" Then
                                    .TextMatrix(LngRow, DI_诊断描述) = rsInput!疾病名称 & ""
                                End If
                            End If
                        Else
                            If Not (IsNull(rsInput!诊断ID) And IsNull(rsInput!疾病id)) Then
                                .Cell(flexcpData, LngRow, DI_诊断描述) = IIf(Not IsNull(rsInput!疾病id), rsInput!疾病名称 & "", rsInput!诊断名称 & "")
                            Else
                                .Cell(flexcpData, LngRow, DI_诊断描述) = .TextMatrix(LngRow, DI_诊断描述)
                            End If
                        End If
                        If Val(rsInput!证候ID & "") <> 0 And .TextMatrix(LngRow, DI_中医证候) = "" Then
                            .TextMatrix(LngRow, DI_中医证候) = rsInput!证候名称 & ""
                        End If
                        .Cell(flexcpData, LngRow, DI_诊断编码) = .TextMatrix(LngRow, DI_诊断编码)
                        .Cell(flexcpData, LngRow, DI_中医证候) = .TextMatrix(LngRow, DI_中医证候)
                        If .TextMatrix(LngRow, DI_诊断描述) <> "" Then
                            .AutoSize DI_诊断编码, DI_诊断描述
                        End If
                        If .ColWidth(DI_诊断描述) < 3200 Then
                            .ColWidth(DI_诊断描述) = 3200
                        End If
                        '其他列数据加
                        .TextMatrix(LngRow, DI_发病时间) = Format(rsInput!发病时间 & "", "YYYY-MM-DD HH:mm")
                        .TextMatrix(LngRow, DI_备注) = rsInput!备注 & ""
                        .TextMatrix(LngRow, DI_出院情况) = rsInput!出院情况 & ""
                        .TextMatrix(LngRow, DI_入院病情) = rsInput!入院病情 & ""
                        If blnGet附码 Then
                            .TextMatrix(LngRow, DI_ICD附码) = rsInput!附码 & ""
                        End If
                        .TextMatrix(LngRow, DI_是否未治) = IIf(Val(rsInput!是否未治 & "") = 1, "√", "")
                        .TextMatrix(LngRow, DI_是否疑诊) = IIf(Val(rsInput!是否疑诊 & "") = 1, "？", "")
                        If gclsPros.FuncType <> f病案首页 Then
                            .TextMatrix(LngRow, DI_诊断ID) = rsInput!诊断ID & ""
                        End If
                        .TextMatrix(LngRow, DI_疾病ID) = rsInput!疾病id & ""
                        .TextMatrix(LngRow, DI_证候ID) = rsInput!证候ID & ""
                        .TextMatrix(LngRow, DI_医嘱IDs) = rsInput!医嘱ID & ""
                        If gclsPros.FuncType = f病案首页 Then
                            If (arrTmp(i) = DT_出院诊断XY Or arrTmp(i) = DT_出院诊断ZY Or arrTmp(i) = DT_院内感染 Or arrTmp(i) = DT_并发症) Then
'                                .TextMatrix(LngRow, DI_固定附码) = IIf(IsNull(rsInput!附码), "", "1")
                                .TextMatrix(LngRow, DI_是否病人) = IIf(Val(rsInput!是否病人 & "") = 1, "1", "")
                            End If
                        End If
                        .TextMatrix(LngRow, DI_疗效限制) = rsInput!疗效限制 & ""
                        .TextMatrix(LngRow, DI_分娩信息) = IIf(IsNull(rsInput!分娩), "0", "1")
                        .TextMatrix(LngRow, DI_诊断来源) = Val(rsInput!记录来源 & "") '保存记录来源，以便保存时，保存为首页或病案来源
                        .TextMatrix(LngRow, DI_疾病编码) = rsInput!疾病编码 & ""
                        .TextMatrix(LngRow, DI_疾病类别) = rsInput!疾病类别 & ""
                        .TextMatrix(LngRow, DI_证候编码) = rsInput!证候编码 & ""
                        .TextMatrix(LngRow, DI_记录日期) = Format(rsInput!记录日期 & "", "YYYY-MM-DD HH:mm")
                        .TextMatrix(LngRow, DI_记录人员) = rsInput!记录人 & ""
                        .RowData(LngRow) = Val(rsInput!ID & "")
                    Else
                        .TextMatrix(LngRow, DI_附码ID) = rsInput!疾病id & ""
                        .TextMatrix(LngRow, DI_ICD附码) = rsInput!疾病编码 & ""
                        .Cell(flexcpData, LngRow, DI_ICD附码) = .TextMatrix(LngRow, DI_ICD附码)
                    End If
                    rsInput.MoveNext
                Loop
            Next
            '设置诊断相关信息
            If gclsPros.FuncType = f病案首页 Then Call SetDeliceryInfo(vsDiagInput)
        End If

        '数据缓存
        strTmp = ""
        arrMain = Array(DI_诊断编码, DI_诊断分类, DI_诊断ID, DI_疾病ID, DI_证候ID, DI_中医证候)
        arrWhole = Array(DI_诊断分类, DI_疾病编码, DI_诊断编码, DI_ICD附码, DI_疾病类别, DI_证候编码, DI_中医证候, DI_是否疑诊, DI_证候ID, DI_诊断ID, DI_疾病ID, DI_诊断描述, DI_备注, DI_发病时间, DI_入院病情, DI_出院情况, DI_是否未治)
        For i = .FixedRows To .Rows - 1
            blnFreeDiag = Val(.TextMatrix(i, DI_诊断ID)) = 0 And Val(.TextMatrix(i, DI_疾病ID)) = 0 '自由录入诊断
            If .TextMatrix(i, DI_诊断描述) <> "" Then
                If strTmp <> .TextMatrix(i, DI_诊断分类) Then
                    j = 1: strTmp = .TextMatrix(i, DI_诊断分类)
                Else
                    j = j + 1
                End If
                strInfo = j: strMainInfo = j
                For k = LBound(arrWhole) To UBound(arrWhole)
                    strInfo = strInfo & "|" & .TextMatrix(i, arrWhole(k))
                Next
                For k = LBound(arrMain) To UBound(arrMain)
                    If strMainInfo = "" Then
                        strMainInfo = .TextMatrix(i, arrMain(k))
                    Else
                        strMainInfo = strMainInfo & "|" & .TextMatrix(i, arrMain(k))
                    End If
                Next
                If blnFreeDiag Then strMainInfo = strMainInfo & "|" & .TextMatrix(i, DI_诊断描述) '自由录入诊断加上诊断描述
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), IIf(bln西医, "西医诊断", "中医诊断"), strInfo, strMainInfo, i, Val(.RowData(i)), IIf(.TextMatrix(i, DI_诊断来源) = "", IIf(gclsPros.FuncType = f病案首页, "4", "3"), .TextMatrix(i, DI_诊断来源)))
                '防止二次保存，修改来源
                If blnOnlyCache Then .TextMatrix(i, DI_诊断来源) = IIf(gclsPros.FuncType = f病案首页, "4", "3")
            End If
        Next
        '加载附码ID
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, DI_ICD附码)) <> "" And Trim(.TextMatrix(i, DI_附码ID)) = "" Then
                Set rsTmp = GetDiagExtraID(Trim(.TextMatrix(i, DI_ICD附码)))
                If rsTmp.RecordCount > 0 Then
                    .TextMatrix(LngRow, DI_附码ID) = rsTmp!ID & ""
                End If
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CacheLoadVsAllerData(ByRef vsAllerInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：将过敏信息加载到表格中并且缓存
'参数：vsAllerInput=需要加载过敏信息的表格
'      rsInput=过敏信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long, j As Long
    Dim strInfo As String, strMainInfo As String
    On Error GoTo errH
    With vsAllerInput
        If Not blnOnlyCache Then
            .Rows = .FixedRows
            For i = 1 To rsInput.RecordCount
                '其它来源的可能有重复
                LngRow = -1
                If Not IsNull(rsInput!药物ID) Then
                    LngRow = .FindRow(rsInput!药物ID & "", , AI_药物ID, , True)
                ElseIf Not IsNull(rsInput!药物名) Then
                    LngRow = .FindRow(rsInput!药物名 & "", , AI_过敏药物, , True)
                End If
                If LngRow = -1 Then
                    For j = .FixedRows To .Rows - 1
                        If .TextMatrix(j, AI_过敏药物) = "" Then
                            LngRow = j
                        End If
                    Next
                    If LngRow = -1 Then .Rows = .Rows + 1: LngRow = .Rows - 1
                    .TextMatrix(LngRow, AI_过敏时间) = Format(rsInput!过敏时间, "yyyy-MM-dd")
                    .TextMatrix(LngRow, AI_过敏药物) = NVL(rsInput!药物名)
                    .TextMatrix(LngRow, AI_过敏反应) = NVL(rsInput!过敏反应)
                    .TextMatrix(LngRow, AI_过敏源编码) = NVL(rsInput!过敏源编码)
                    .TextMatrix(LngRow, AI_药物ID) = rsInput!药物ID & ""
                    .TextMatrix(LngRow, AI_过敏来源) = rsInput!记录来源 & ""
                    '数据备份存储
                    .Cell(flexcpData, LngRow, AI_过敏时间) = .TextMatrix(LngRow, AI_过敏时间)
                    .Cell(flexcpData, LngRow, AI_过敏药物) = .TextMatrix(LngRow, AI_过敏药物)
                    .Cell(flexcpData, LngRow, AI_过敏反应) = .TextMatrix(LngRow, AI_过敏反应)
                    .Cell(flexcpData, LngRow, AI_过敏源编码) = .TextMatrix(LngRow, AI_过敏源编码)
                    .Cell(flexcpData, LngRow, AI_药物ID) = .TextMatrix(LngRow, AI_药物ID)
                    .RowData(LngRow) = Val(rsInput!ID & "")
                End If
                rsInput.MoveNext
            Next
            .Rows = .Rows + 1 '增加一行空行
            .Row = .FixedRows: .Col = AI_过敏药物
        End If

        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, AI_过敏药物) <> "" Then
                strInfo = .TextMatrix(i, AI_过敏时间) & "|" & .TextMatrix(i, AI_过敏药物) & "|" & .TextMatrix(i, AI_过敏反应) & "|" & .TextMatrix(i, AI_过敏源编码) & "|" & .TextMatrix(i, AI_药物ID) & "|" & .RowData(i)
                strMainInfo = .TextMatrix(i, AI_过敏时间) & "|" & .TextMatrix(i, AI_过敏源编码) & "|" & .TextMatrix(i, AI_药物ID) & "|" & .TextMatrix(i, AI_过敏药物)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "过敏药物", strInfo, strMainInfo, i, Val(.RowData(i)), IIf(.TextMatrix(i, AI_过敏来源) = "", IIf(gclsPros.FuncType = f病案首页, "4", "3"), .TextMatrix(i, AI_过敏来源)))
                '防止二次保存，修改来源
                If blnOnlyCache Then .TextMatrix(i, AI_过敏来源) = IIf(gclsPros.FuncType = f病案首页, "4", "3")
            End If
        Next
    End With
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CacheLoadVsOPSData(ByRef vsOPSInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载病人手麻数据并缓存
'参数：vsOPSInput=需要加载病人手麻信息的表格
'      rsInput=病人手麻信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long, j As Long
    Dim strInfo As String, strMainInfo As String
    Dim lngOrder As Long
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    With vsOPSInput
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '固定行+新行
            For i = 1 To rsInput.RecordCount
                .TextMatrix(i, PI_手术日期) = Format(NVL(rsInput!手术开始时间, rsInput!手术日期) & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, PI_结束日期) = Format(NVL(rsInput!手术结束时间, rsInput!手术日期) & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, PI_手术编码) = rsInput!手术编码 & ""
                .TextMatrix(i, PI_手术名称) = rsInput!手术名称 & ""
                If (Not gclsPros.CNIndent And gclsPros.FuncType = f病案首页) Or .TextMatrix(i, PI_手术名称) = "" Then
                    .TextMatrix(i, PI_手术名称) = rsInput!手术原名 & ""
                    If .TextMatrix(i, PI_手术名称) = "" Then
                        .TextMatrix(i, PI_手术名称) = rsInput!手术名称 & ""
                    End If
                End If
                If .TextMatrix(i, PI_手术名称) <> "" Then
                    .AutoSize PI_手术编码, PI_手术名称
                End If
                .TextMatrix(i, PI_主刀医师) = rsInput!主刀医师 & ""
                .TextMatrix(i, PI_助产护士) = rsInput!助产护士 & ""
                .TextMatrix(i, PI_助手1) = rsInput!第一助手 & ""
                .TextMatrix(i, PI_助手2) = rsInput!第二助手 & ""
                .TextMatrix(i, PI_麻醉方式) = rsInput!麻醉方式 & ""
                .TextMatrix(i, PI_麻醉医师) = rsInput!麻醉医师 & ""
                If rsInput!切口 & rsInput!愈合 & "" <> "" Then
                    .TextMatrix(i, PI_切口愈合) = rsInput!切口 & "/" & rsInput!愈合
                End If
                .TextMatrix(i, PI_手术操作ID) = rsInput!手术操作ID & ""
                .TextMatrix(i, PI_诊疗项目ID) = rsInput!诊疗项目id & ""
                .TextMatrix(i, PI_麻醉ID) = rsInput!麻醉ID & ""
                .TextMatrix(i, PI_麻醉类型) = rsInput!麻醉类型 & ""
                .TextMatrix(i, PI_手术情况) = rsInput!手术情况 & ""
                .TextMatrix(i, PI_ASA分级) = rsInput!asa分级 & ""
                .TextMatrix(i, PI_NNIS分级) = rsInput!NNIS分级 & ""
                .TextMatrix(i, PI_手术级别) = rsInput!手术级别 & ""
                .TextMatrix(i, PI_再次手术) = IIf(Val(rsInput!再次手术 & "") = 1, -1, 0)
                .TextMatrix(i, PI_准备天数) = IIf(Val(rsInput!准备天数 & "") = 0, "", Val(rsInput!准备天数 & ""))
                .TextMatrix(i, PI_抗菌用药时间) = Format(rsInput!抗菌用药时间 & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, PI_麻醉开始时间) = Format(rsInput!麻醉开始时间 & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, PI_切口部位) = rsInput!切口部位 & ""
                .TextMatrix(i, PI_重返手术室目的) = rsInput!重返目的 & ""
                .Cell(flexcpChecked, i, PI_重返手术室计划) = Val(rsInput!重返计划 & "")
                .Cell(flexcpChecked, i, PI_切口感染) = Val(rsInput!切口感染 & "")
                .Cell(flexcpChecked, i, PI_并发症) = Val(rsInput!并发症 & "")
                '10.34.10新增
                .TextMatrix(i, PI_抗菌药天数) = IIf(Val(rsInput!抗菌用药天数 & "") = 0, "", Val(rsInput!抗菌用药天数 & ""))
                .Cell(flexcpChecked, i, PI_预防用抗菌药) = Val(rsInput!术前抗菌用药 & "")
                .Cell(flexcpChecked, i, PI_非预期的二次手术) = Val(rsInput!非预期的二次手术 & "")
                .Cell(flexcpChecked, i, PI_麻醉并发症) = Val(rsInput!麻醉并发症 & "")
                .Cell(flexcpChecked, i, PI_术中异物遗留) = Val(rsInput!术中异物遗留 & "")
                .Cell(flexcpChecked, i, PI_手术并发症) = Val(rsInput!手术并发症 & "")
                .Cell(flexcpChecked, i, PI_术后出血或血肿) = Val(rsInput!术后出血或血肿 & "")
                .Cell(flexcpChecked, i, PI_手术伤口裂开) = Val(rsInput!手术伤口裂开 & "")
                .Cell(flexcpChecked, i, PI_术后深静脉血栓) = Val(rsInput!术后深静脉血栓 & "")
                .Cell(flexcpChecked, i, PI_术后生理代谢紊乱) = Val(rsInput!术后生理代谢紊乱 & "")
                .Cell(flexcpChecked, i, PI_术后呼吸衰竭) = Val(rsInput!术后呼吸衰竭 & "")
                .Cell(flexcpChecked, i, PI_术后肺栓塞) = Val(rsInput!术后肺栓塞 & "")
                .Cell(flexcpChecked, i, PI_术后败血症) = Val(rsInput!术后败血症 & "")
                .Cell(flexcpChecked, i, PI_术后髋关节骨折) = Val(rsInput!术后髋关节骨折 & "")
                .Cell(flexcpData, i, PI_手术名称) = rsInput!手术原名 & ""
                .TextMatrix(i, PI_手麻来源) = rsInput!记录来源 & ""
                .RowData(i) = Val(rsInput!ID & "")
                '记录用于编辑恢复
                For j = 0 To .Cols - 1
                    If j = PI_手术名称 And .TextMatrix(i, PI_手术编码) <> "" Then
                        If .Cell(flexcpData, i, j) = "" Then
                            .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                        End If
                    Else
                        .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                    End If
                Next

                If Trim(.TextMatrix(i, PI_手术级别)) <> "" And rsInput!原手术级别 & "" <> "" Then
                    .Cell(flexcpData, i, PI_手术级别) = 1
                End If
                rsInput.MoveNext
            Next
        End If
        '数据缓存
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, PI_手术名称) <> "" Then
                lngOrder = lngOrder + 1
                strInfo = .TextMatrix(i, PI_手术日期) & "|" & .TextMatrix(i, PI_结束日期) & "|" & .TextMatrix(i, PI_手术编码) & "|" & .TextMatrix(i, PI_手术名称) & "|" & .TextMatrix(i, PI_主刀医师) & "|" & .TextMatrix(i, PI_助产护士) & "|" & _
                        .TextMatrix(i, PI_助手1) & "|" & .TextMatrix(i, PI_助手2) & "|" & .TextMatrix(i, PI_麻醉方式) & "|" & .TextMatrix(i, PI_麻醉医师) & "|" & .TextMatrix(i, PI_切口愈合) & "|" & .TextMatrix(i, PI_手术操作ID) & "|" & _
                        .TextMatrix(i, PI_诊疗项目ID) & "|" & .TextMatrix(i, PI_麻醉ID) & "|" & .TextMatrix(i, PI_麻醉类型) & "|" & .TextMatrix(i, PI_手术情况) & "|" & .TextMatrix(i, PI_ASA分级) & "|" & .TextMatrix(i, PI_NNIS分级) & "|" & _
                        .TextMatrix(i, PI_手术级别) & "|" & .TextMatrix(i, PI_再次手术) & "|" & .TextMatrix(i, PI_准备天数) & "|" & .TextMatrix(i, PI_抗菌用药时间) & "|" & .TextMatrix(i, PI_麻醉开始时间) & "|" & .TextMatrix(i, PI_切口部位) & "|" & _
                        .TextMatrix(i, PI_重返手术室目的) & "|" & .Cell(flexcpChecked, i, PI_重返手术室计划) & "|" & .TextMatrix(i, PI_切口感染) & "|" & .Cell(flexcpChecked, i, PI_并发症) & "|" & .Cell(flexcpChecked, i, PI_预防用抗菌药) & "|" & _
                        .TextMatrix(i, PI_抗菌药天数) & "|" & .Cell(flexcpChecked, i, PI_非预期的二次手术) & "|" & .Cell(flexcpChecked, i, PI_麻醉并发症) & "|" & .Cell(flexcpChecked, i, PI_术中异物遗留) & "|" & .Cell(flexcpChecked, i, PI_手术并发症) & "|" & _
                        .Cell(flexcpChecked, i, PI_术后出血或血肿) & "|" & .Cell(flexcpChecked, i, PI_手术伤口裂开) & "|" & .Cell(flexcpChecked, i, PI_术后深静脉血栓) & "|" & .Cell(flexcpChecked, i, PI_术后生理代谢紊乱) & "|" & .Cell(flexcpChecked, i, PI_术后呼吸衰竭) & "|" & _
                        .Cell(flexcpChecked, i, PI_术后肺栓塞) & "|" & .Cell(flexcpChecked, i, PI_术后败血症) & "|" & .Cell(flexcpChecked, i, PI_术后髋关节骨折) & "|" & .RowData(i) & "|" & lngOrder
                If gclsPros.MedPageSandard = ST_卫生部标准 Then
                    strMainInfo = .TextMatrix(i, PI_手术日期) & "|" & .TextMatrix(i, PI_结束日期) & "|" & .TextMatrix(i, PI_手术编码) & "|" & .TextMatrix(i, PI_手术名称) & "|" & .TextMatrix(i, PI_手术操作ID) & "|" & .TextMatrix(i, PI_诊疗项目ID) & "|" & .TextMatrix(i, PI_切口部位)
                Else
                    strMainInfo = .TextMatrix(i, PI_手术日期) & "|" & .TextMatrix(i, PI_结束日期) & "|" & .TextMatrix(i, PI_手术编码) & "|" & .TextMatrix(i, PI_手术名称) & "|" & .TextMatrix(i, PI_手术操作ID) & "|" & .TextMatrix(i, PI_诊疗项目ID)
                End If
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "手麻情况", strInfo, strMainInfo, i, Val(.RowData(i)), IIf(.TextMatrix(i, PI_手麻来源) = "", IIf(gclsPros.FuncType = f病案首页, "4", "3"), .TextMatrix(i, PI_手麻来源)))
                '防止二次保存，修改来源
                If blnOnlyCache Then .TextMatrix(i, PI_手麻来源) = IIf(gclsPros.FuncType = f病案首页, "4", "3")
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
End Sub

Private Sub CacheLoadVsChemothData(ByRef vsChemothInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载化疗数据并缓存
'参数：vsChemothInput=需要加载化疗信息的表格
'      rsInput=化疗信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsChemothInput
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '固定行+新行
            For i = 1 To rsInput.RecordCount
                .RowData(i) = Val(rsInput!序号 & "")
                .TextMatrix(i, CI_化学治疗编码) = NVL(rsInput!疾病信息)
                .TextMatrix(i, CI_开始日期) = Format(rsInput!开始日期, "yyyy-MM-dd")
                .TextMatrix(i, CI_结束日期) = Format(rsInput!结束日期, "yyyy-MM-dd")
                .TextMatrix(i, CI_疗程数) = Format(Val(rsInput!疗程数 & ""), "###;-###;;")
                .TextMatrix(i, CI_总量) = Format(Val(rsInput!总量 & ""), "###;-###;;")
                .TextMatrix(i, CI_化疗方案) = rsInput!化疗方案 & ""
                .TextMatrix(i, CI_化疗效果) = rsInput!化疗效果 & ""
                .TextMatrix(i, CI_疾病ID) = rsInput!疾病id & ""
                rsInput.MoveNext
            Next
        End If
        '数据缓存
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, CI_化学治疗编码) <> "" Then
                strInfo = .TextMatrix(i, CI_化学治疗编码) & "|" & .TextMatrix(i, CI_开始日期) & "|" & .TextMatrix(i, CI_结束日期) & "|" & .TextMatrix(i, CI_疗程数) & "|" & .TextMatrix(i, CI_总量) & "|" & .TextMatrix(i, CI_化疗方案) & "|" & .TextMatrix(i, CI_化疗效果) & "|" & .TextMatrix(i, CI_疾病ID) & "|" & .RowData(i)
                strMainInfo = .RowData(i) & "|" & .TextMatrix(i, CI_化学治疗编码)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "病案化疗记录", strInfo, strMainInfo, i, IIf(blnOnlyCache, i, Val(.RowData(i))))
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsRadiothData(ByRef vsRadiothInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载放疗数据并缓存
'参数：vsRadiothInput=需要加载放疗信息的表格
'      rsInput=放疗信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsRadiothInput
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '固定行+新行
            For i = 1 To rsInput.RecordCount
                .RowData(i) = Val(rsInput!序号 & "")
                .TextMatrix(i, RI_放射治疗编码) = NVL(rsInput!疾病信息)
                .TextMatrix(i, RI_开始日期) = Format(rsInput!开始日期, "yyyy-MM-dd")
                .TextMatrix(i, RI_结束日期) = Format(rsInput!结束日期, "yyyy-MM-dd")
                .TextMatrix(i, RI_放射剂量) = Format(Val(rsInput!放射剂量 & ""), "###;-###;;")
                .TextMatrix(i, RI_累计量) = Format(Val(rsInput!累计量 & ""), "###;-###;;")
                .TextMatrix(i, RI_设野部位) = rsInput!设野部位 & ""
                .TextMatrix(i, RI_放疗效果) = rsInput!放疗效果 & ""
                .TextMatrix(i, RI_疾病ID) = rsInput!疾病id & ""
                rsInput.MoveNext
            Next
        End If
        '数据缓存
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, RI_放射治疗编码) <> "" Then
                strInfo = .TextMatrix(i, RI_放射治疗编码) & "|" & .TextMatrix(i, RI_开始日期) & "|" & .TextMatrix(i, RI_结束日期) & "|" & .TextMatrix(i, RI_放射剂量) & "|" & .TextMatrix(i, RI_累计量) & "|" & .TextMatrix(i, RI_设野部位) & "|" & .TextMatrix(i, RI_放疗效果) & "|" & .TextMatrix(i, RI_疾病ID) & "|" & .RowData(i)
                strMainInfo = .RowData(i) & "|" & .TextMatrix(i, RI_放射治疗编码)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "病案放疗记录", strInfo, strMainInfo, i, IIf(blnOnlyCache, i, Val(.RowData(i))))
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsSpiritData(ByRef vsSpiritInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载精神药品使用数据并缓存
'参数：vsSpiritInput=需要加载精神药品信息的表格
'      rsInput=精神药品信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsSpiritInput
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '固定行+新行
            For i = 1 To rsInput.RecordCount
                .RowData(i) = Val(rsInput!序号 & "")
                .TextMatrix(i, SI_药物名称) = rsInput!药物名称 & ""
                .TextMatrix(i, SI_疗程) = rsInput!疗程 & ""
                .TextMatrix(i, SI_最高日量) = rsInput!最高日量 & ""
                .TextMatrix(i, SI_特殊反应) = rsInput!特殊反应 & ""
                .TextMatrix(i, SI_疗效) = rsInput!疗效 & ""
                .TextMatrix(i, SI_药品ID) = rsInput!药品ID & ""
                rsInput.MoveNext
            Next
        End If
        '数据缓存
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, SI_药物名称) <> "" Then
                strInfo = .TextMatrix(i, SI_药物名称) & "|" & .TextMatrix(i, SI_疗程) & "|" & .TextMatrix(i, SI_最高日量) & "|" & .TextMatrix(i, SI_特殊反应) & "|" & .TextMatrix(i, SI_疗效) & "|" & .TextMatrix(i, SI_药品ID) & "|" & .RowData(i)
                strMainInfo = .RowData(i) & "|" & .TextMatrix(i, SI_药物名称)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "病案精神治疗", strInfo, strMainInfo, i, IIf(blnOnlyCache, i, Val(.RowData(i))))
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsKSSData(ByRef vsKSSInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载抗菌药使用数据并缓存
'参数：vsKSSInput=需要加载抗菌药信息的表格
'      rsInput=抗菌药信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsKSSInput
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then Exit Sub
            Do While Not rsInput.EOF
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, KI_抗菌药物名) = "" Or (.RowData(i) = Val(rsInput!药名id & "") And .TextMatrix(i, KI_用药目的) = rsInput!用药目的 & "" And .TextMatrix(i, KI_使用阶段) = rsInput!使用阶段 & "") Then
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
        End If
        '数据缓存
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, KI_抗菌药物名) <> "" Then
                strInfo = .RowData(i) & "|" & .TextMatrix(i, KI_抗菌药物名) & "|" & .TextMatrix(i, KI_用药目的) & "|" & .TextMatrix(i, KI_使用阶段) & "|" & .TextMatrix(i, KI_使用天数) & "|" & .Cell(flexcpChecked, i, KI_一类切口预防用) & "|" & .TextMatrix(i, KI_DDD数) & "|" & .TextMatrix(i, KI_联合用药) & "|" & .RowData(i)
                strMainInfo = .RowData(i) & "|" & .TextMatrix(i, KI_抗菌药物名) & "|" & .TextMatrix(i, KI_用药目的) & "|" & .TextMatrix(i, KI_使用阶段)
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "病人抗生素记录", strInfo, strMainInfo, i)
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsFlxAddICUData(Optional ByRef vsFlxAddICUInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载重症监护使用数据并缓存
'参数：vsFlxAddICUInput=需要加载重症监护信息的表格
'      rsInput=重症监护信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String
    Dim strList As String
    Dim blnLocked As Boolean

    If Not vsFlxAddICUInput Is Nothing Then
        With vsFlxAddICUInput
            '数据加载
            If Not blnOnlyCache Then
                If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
                .Rows = rsInput.RecordCount + 2 '固定行+新行
                For i = 1 To rsInput.RecordCount
                    .TextMatrix(i, UI_监护室名称) = rsInput!监护室名称 & ""
                    .TextMatrix(i, UI_进入时间) = rsInput!进入时间 & ""
                    .TextMatrix(i, UI_退出时间) = rsInput!退出时间 & ""
                    If gclsPros.MedPageSandard = ST_四川省标准 Then
                        .TextMatrix(i, UI_序号) = i
                        .Cell(flexcpChecked, i, UI_再入住计划) = Val(rsInput!再入住计划 & "")
                        .TextMatrix(i, UI_再入住原因) = rsInput!再入住原因 & ""
                    Else
                        .TextMatrix(i, UI_序号) = Val(rsInput!序号 & "")
                    End If
                    .RowData(i) = Val(rsInput!序号 & "")
                    rsInput.MoveNext
                Next
            End If
            '数据缓存
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, UI_监护室名称) <> "" Then
                    strList = strList & "|" & .TextMatrix(i, UI_序号) & "-" & .TextMatrix(i, UI_监护室名称)
                    strInfo = .TextMatrix(i, UI_序号) & "|" & .TextMatrix(i, UI_监护室名称) & "|" & .TextMatrix(i, UI_进入时间) & "|" & .TextMatrix(i, UI_退出时间) & "|" & .Cell(flexcpChecked, i, UI_再入住计划) & "|" & .TextMatrix(i, UI_再入住原因) & "|" & .RowData(i)
                    strMainInfo = .TextMatrix(i, UI_监护室名称) & "|" & .RowData(i)
                    Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "病案重症监护情况", strInfo, strMainInfo, i, IIf(blnOnlyCache, i, Val(.RowData(i))))
                End If
            Next
            strList = Mid(strList, 2)
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                gclsPros.CurrentForm.vsICUInstruments.ColComboList(TI_ICU类型) = strList
                gclsPros.CurrentForm.vsICUInstruments.Editable = IIf(strList <> "", flexEDKbdMouse, flexEDNone)
            End If
        End With
    Else
        '云南版，没有表格
        If Not rsInput Is Nothing Then
            rsInput.Sort = "序号"
            If Not rsInput.EOF Then
                rsInput.MoveFirst
                For i = 0 To rsInput.Fields.Count - 1
                    Call SetCtrlValues(rsInput.Fields(i).Name, rsInput.Fields(i).Value & "")
                Next
            End If
        End If
        blnLocked = gclsPros.CurrentForm.txtInfo(GC_重症监护室名称).Text = ""
        Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_人工气道脱出), blnLocked, True)
        Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_重返重症医学科), blnLocked, True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_重返间隔时间), blnLocked, True)
    End If
End Sub

Private Sub CacheLoadVsICUInstrumentsData(ByRef vsICUInstruments As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载器械导管使用情况数据并缓存
'参数：vsICUInstruments=需要加载器械导管使用情况信息的表格
'      rsInput=器械导管使用情况信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsICUInstruments
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '固定行+新行
            For i = 1 To rsInput.RecordCount
                .TextMatrix(i, TI_ICU类型) = rsInput!监护室名称 & ""
                .Cell(flexcpData, i, TI_ICU类型) = Val(rsInput!序号 & "")
                .TextMatrix(i, TI_器械及导管) = rsInput!器械及导管 & ""
                .TextMatrix(i, TI_开始时间) = rsInput!开始使用时间 & ""
                .TextMatrix(i, TI_结束时间) = rsInput!结束使用时间 & ""
                .TextMatrix(i, TI_感染累计小时) = rsInput!感染累计时间 & ""
                .RowData(i) = Val(rsInput!序号 & "")
                rsInput.MoveNext
            Next
        End If
        '数据缓存
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, TI_器械及导管) <> "" Then
                strInfo = .TextMatrix(i, TI_ICU类型) & "|" & .TextMatrix(i, TI_器械及导管) & "|" & .TextMatrix(i, TI_开始时间) & "|" & .TextMatrix(i, TI_结束时间) & "|" & .TextMatrix(i, TI_感染累计小时) & "|" & i
                strMainInfo = .TextMatrix(i, TI_ICU类型) & "|" & .TextMatrix(i, TI_器械及导管) & "|" & i
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "器械导管使用情况", strInfo, strMainInfo, i)
            End If
        Next
    End With
End Sub

Private Sub CacheLoadvsInfectData(ByRef vsInfect As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载病人感染记录数据并缓存
'参数：vsInfect=需要加载病人感染记录信息的表格
'      rsInput=病人感染记录信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsInfect
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '固定行+新行
            For i = 1 To rsInput.RecordCount
                .TextMatrix(i, FI_确诊日期) = rsInput!确诊日期 & ""
                .TextMatrix(i, FI_感染部位) = rsInput!感染部位 & ""
                .TextMatrix(i, FI_医院感染名称) = rsInput!医院感染名称 & ""
                .TextMatrix(i, FI_医院感染编码) = rsInput!医院感染编码 & ""
                .RowData(i) = Val(rsInput!序号 & "")
                rsInput.MoveNext
            Next
        End If
        '数据缓存
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, FI_感染部位) <> "" Then
                strInfo = .TextMatrix(i, FI_确诊日期) & "|" & .TextMatrix(i, FI_感染部位) & "|" & .TextMatrix(i, FI_医院感染名称) & "|" & i
                strMainInfo = .TextMatrix(i, FI_感染部位) & "|" & .TextMatrix(i, FI_医院感染名称) & "|" & i
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "病人感染记录", strInfo, strMainInfo, i)
            End If
        Next
    End With
End Sub

Private Sub CacheLoadvsSampleData(ByRef vsSample As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载病人病原学检查数据并缓存
'参数：vsSample=需要加载病人病原学检查信息的表格
'      rsInput=病人病原学检查信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数
    Dim i As Long, LngRow As Long
    Dim strInfo As String, strMainInfo As String

    With vsSample
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 1: Exit Sub
            .Rows = rsInput.RecordCount + 2 '固定行+新行
            For i = 1 To rsInput.RecordCount
                .TextMatrix(i, MI_标本) = decode(Val(rsInput!标本 & ""), 1, "1.血液", 2, "2.尿液", 3, "3.粪便", 4, "4.痰液", 5, "5.其他分泌物")
                .TextMatrix(i, MI_病原学代码及名称) = rsInput!病原学代码 & ""
                .TextMatrix(i, MI_送检日期) = rsInput!送检日期 & ""
                .RowData(i) = Val(rsInput!序号 & "")
                rsInput.MoveNext
            Next
        End If
        '数据缓存
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, MI_标本) <> "" Then
                strInfo = .TextMatrix(i, MI_标本) & "|" & .TextMatrix(i, MI_病原学代码及名称) & "|" & .TextMatrix(i, MI_送检日期) & "|" & i
                strMainInfo = .TextMatrix(i, MI_标本) & "|" & i
                Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "病人病原学检查", strInfo, strMainInfo, i)
            End If
        Next
    End With
End Sub

Private Sub CacheLoadVsFreesData(ByRef vsFees As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean, Optional ByVal bln编目 As Boolean)
'功能：加载病人住院费用数据并缓存
'参数：vsFees=需要加载病人住院费用的表格
'      rsInput=病人住院费用信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'      bln编目=数据加载时费用是否来自已经编目信息
'说明：LoadMedPageData的子函数
    Dim i As Long, LngCol As Long, LngRow As Long, lng序号 As Long
    Dim strInfo As String, strMainInfo As String
    Dim dbl婴儿费 As Double, dblSum As Double
    Dim blnHave As Boolean, bln婴儿 As Boolean
    On Error GoTo errH
    With vsFees
        '数据加载
        If Not blnOnlyCache Then
            If rsInput Is Nothing Then .Rows = .FixedRows + 2: Exit Sub
            If Not bln编目 Then
                rsInput.Filter = "婴儿费<>0"
                Do While Not rsInput.EOF
                    dbl婴儿费 = dbl婴儿费 + Val(rsInput!金额 & "")
                    rsInput.MoveNext
                Loop
                rsInput.Filter = "婴儿费=0"
            End If
            .Rows = .FixedRows + rsInput.RecordCount \ 3 + 1

            For i = 0 To rsInput.RecordCount - 1
                If i Mod 3 = 0 Then LngRow = LngRow + 1 '3栏填满则定位到下一行
                LngCol = (i Mod 3) * 2 '定位列
                bln婴儿 = rsInput!费目名称 = "婴儿费"
                If bln婴儿 Then blnHave = True
                .TextMatrix(LngRow, LngCol) = rsInput!名称
                .TextMatrix(LngRow, LngCol + 1) = Format(Val(rsInput!金额 & "") + IIf(bln婴儿, dbl婴儿费, 0), gclsPros.FreeFormat)
                rsInput.MoveNext
            Next

            If dbl婴儿费 <> 0 And Not blnHave Then
                If LngCol = 4 Then
                    LngCol = 0: LngRow = LngRow + 1 '写满了，移动到下一行
                Else
                    LngCol = LngCol + 2 '移动到下一栏
                End If
                .TextMatrix(LngRow, LngCol) = "婴儿费"
                .TextMatrix(LngRow, LngCol + 1) = Format(dbl婴儿费, gclsPros.FreeFormat)
                If LngCol = 4 Then .Rows = .Rows + 1 '如果婴儿费写的是第三栏，则新增空行
            End If
            Call SumAndSetFrees
        End If
        '数据缓存
        '加载未编目数据，第一次不缓存,这些数据会当作新增数据
        If bln编目 Or blnOnlyCache Then
            For i = 3 To .Rows * 3 - 1
                LngRow = i \ 3: LngCol = (i Mod 3) * 2
                If .TextMatrix(LngRow, LngCol) <> "" Then
                    strInfo = .TextMatrix(LngRow, LngCol)
                    strMainInfo = .TextMatrix(LngRow, LngCol) & "|" & .TextMatrix(LngRow, LngCol + 1)
                    Call UpdateCacheRecInfo(IIf(blnOnlyCache, 1, 0), "病人费用", strMainInfo, strInfo, , , LngRow & "," & LngCol)
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Sub LoadTransferData(Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean, Optional ByVal bln编目 As Boolean)
'功能：加载病人住院转科信息并缓存
'参数：
'      rsInput=病人住院转科信息记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'      bln编目=数据加载时是否来自已经编目信息
'说明：LoadMedPageData的子函数
    Dim i As Long, LngCol As Long, LngRow As Long
    Dim vsTranfer As VSFlexGrid

    Dim strInfo As String, strMainInfo As String
    If gclsPros.FuncType <> f病案首页 Then
        With gclsPros.CurrentForm
            If .txtInfo(GC_转科1).Text = "" And .txtInfo(GC_转科2).Text = "" And .txtInfo(GC_转科3).Text = "" Then
                For i = 1 To rsInput.RecordCount
                    If i = 1 Then
                        .txtInfo(GC_转科1).Text = rsInput!科室名称 & ""
                    ElseIf i = 2 Then
                        .txtInfo(GC_转科2).Text = rsInput!科室名称 & ""
                    ElseIf i = 3 Then
                        .txtInfo(GC_转科3).Text = rsInput!科室名称 & ""
                        Exit For
                    End If
                    rsInput.MoveNext
                Next
            End If
        End With
    Else
        Set vsTranfer = gclsPros.CurrentForm.vsTransfer
        With vsTranfer
            For i = 1 To rsInput.RecordCount
                .TextMatrix(0, i) = rsInput!科室名称
                .TextMatrix(1, i) = Format(rsInput!开始时间, "YYYY-MM-DD")
                If i = 6 Then Exit For
                rsInput.MoveNext
            Next
        End With
    End If
End Sub

Public Function FreeHaveLowLevel(ByVal LngRow As Long, ByVal LngCol As Long) As Boolean
'功能:判断费用级别是否包含下级费用
    Dim strCode As String
    Dim lngPos As Long, i As Integer
    Dim vsTmp As VSFlexGrid

    Set vsTmp = gclsPros.CurrentForm.vsFees
    With vsTmp
        If .TextMatrix(LngRow, LngCol) = "" Then Exit Function
        strCode = GetFreeCode(.TextMatrix(LngRow, LngCol), True)
        If strCode = "" Then FreeHaveLowLevel = True: Exit Function
        lngPos = InStr(1, strCode, "_") '费用编码格式：父级编码_编码
        '判断是否有有下一级费用，则取当前编码
        If lngPos > 0 Then strCode = Mid(strCode, lngPos + 1)

        For i = 3 To (.Rows * 3) - 1
            LngRow = i \ 3: LngCol = (i Mod 3) * 2 '定位行列
            If .TextMatrix(LngRow, LngCol) Like strCode & "_*.*" Then
                FreeHaveLowLevel = True: Exit Function  '存在子级判断退出
            End If
        Next
    End With
End Function

Public Sub SumAndSetFrees()
'功能:计算各类费用累计，并设置单元格的值以及总费用
    Dim strCode As String, strFathCode As String
    Dim dblSum As Double, lngPos As Long, i As Long, j As Long
    Dim LngRow As Long, LngCol As Long
    Dim vsTmp As VSFlexGrid
    Dim intID As Integer
    Dim blnDo As Boolean, strFee As String

    Dim rsFeeList As New ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim lngSort As Long

    On Error GoTo errH
    rsFeeList.Fields.Append "ID", adInteger, , adFldKeyColumn           '主键
    rsFeeList.Fields.Append "Row", adInteger                            '行号
    rsFeeList.Fields.Append "Col", adInteger                            '列号
    rsFeeList.Fields.Append "Code", adVarChar, 200                      '编码
    rsFeeList.Fields.Append "PID", adInteger, , adFldIsNullable         '父级ID
    rsFeeList.Fields.Append "Fee", adVarChar, 200, adFldIsNullable      '费用字符串
    rsFeeList.Fields.Append "Sort", adInteger, , adFldIsNullable        '排序
    rsFeeList.CursorLocation = adUseClient
    rsFeeList.LockType = adLockOptimistic
    rsFeeList.CursorType = adOpenStatic
    rsFeeList.Open

    Set vsTmp = gclsPros.CurrentForm.vsFees
    With vsTmp
        For i = 3 To (.Rows * 3) - 1
            LngRow = i \ 3: LngCol = (i Mod 3) * 2
            strCode = GetFreeCode(.TextMatrix(LngRow, LngCol))
            If strCode <> "" Then
                rsFeeList.AddNew Array("ID", "Row", "Col", "Code", "Sort"), Array(Identity(lngSort), LngRow, LngCol, strCode, 0)
            End If
        Next
        rsFeeList.Filter = "": rsFeeList.Sort = "Code,ID"
        Set rsTmp = zlDatabase.CopyNewRec(rsFeeList) '备份记录集，用来搜索父ID
        rsTmp.Filter = "": rsTmp.Sort = "Code,ID"
        For i = 1 To rsTmp.RecordCount
            strCode = rsTmp!Code & ""
            lngPos = InStr(strCode, "_")
            If lngPos > 0 Then '获取当前编码作为父编码
                strFathCode = Mid(strCode, lngPos + 1)
            Else
                strFathCode = strCode
            End If
            '搜索子级
            Call Rec.Update(rsFeeList, "Code Like '" & strFathCode & "_*'", "PID", rsTmp!ID)
            rsTmp.MoveNext
        Next
        '设置、计算各个费用之和
        rsFeeList.Filter = "": rsFeeList.Sort = "ID"
        Set rsTmp = zlDatabase.CopyNewRec(rsFeeList) '备份记录集，用来搜索父ID
        Do While JudeSet(rsFeeList)
            intID = Val(rsFeeList!ID & "")
            rsTmp.Filter = "PID=" & intID
            blnDo = False
            If rsTmp.EOF Then '当前费用是最底级
                blnDo = True
            Else '当前费用不是最低级
                rsTmp.Filter = "PID=" & intID & " And Fee=Null"
                dblSum = 0
                If rsTmp.EOF Then '费用的子级均读取到了,则当前费用子级费用求和
                    rsTmp.Filter = "PID=" & intID
                    Do While Not rsTmp.EOF
                        dblSum = dblSum + Val(rsTmp!Fee & "")
                        rsTmp.MoveNext
                    Loop
                    .TextMatrix(rsFeeList!Row, rsFeeList!Col + 1) = Format(dblSum, gclsPros.FreeFormat) '将费用设置在表格上
                    blnDo = True
                End If
            End If
            If blnDo Then '将表格费用填写在记录集上
                rsTmp.Filter = "ID=" & intID
                strFee = Format(Val(.TextMatrix(rsFeeList!Row, rsFeeList!Col + 1)), gclsPros.FreeFormat)
                rsFeeList.Update "Fee", strFee
                rsTmp.Update "Fee", strFee
            Else
                rsFeeList.Update "Sort", Val(rsFeeList!Sort & "") + 1 '排序字段增加，顺序排在后面
            End If
        Loop
        '计算并设置总费用
        rsFeeList.Filter = "PID=Null"
        dblSum = 0
        Do While Not rsFeeList.EOF
            dblSum = dblSum + Val(rsFeeList!Fee & "")
            rsFeeList.MoveNext
        Loop
        gclsPros.CurrentForm.txtSpecificInfo(SLC_费用和).Text = Format(dblSum, gclsPros.FreeFormat)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function JudeSet(ByRef rsInput As ADODB.Recordset) As Boolean
'功能：判断是否有未读取到的费用
    rsInput.Filter = "Fee=Null"
    rsInput.Sort = "Sort,ID"
    JudeSet = Not rsInput.EOF
End Function
Public Function GetFreeCode(ByVal strFreeName As String, Optional ByVal blnMustHave As Boolean) As String
'功能：获取费用编码
'参数：strFreeName=费用名
'      blnMustHave=是否必须有编码，没有编码返回空
'返回：费用编码
'说明：SetHlevelFreeSum的子函数
    Dim strCode As String
    Dim lngPos As Long

    lngPos = InStr(strFreeName, ".")
    If lngPos > 0 Then strCode = Mid(strFreeName, 1, lngPos - 1)
    If blnMustHave And lngPos <= 0 Then strCode = ""
    GetFreeCode = strCode
End Function

Public Sub FilterDiagByType(ByRef rsInput As ADODB.Recordset, ByVal intDiagType As Integer, Optional ByVal intMaxDiagSource As Integer = -1)
'功能：获取制定类型的诊断
'参数：rsInput=需要过滤的诊断记录集
'      intDiagType-诊断类型
'      intMaxDiagSource=最大的记录来源：病案首页使用,病案诊断提取首页时，按诊断来源从大到小依次提取
'返回：rsInput=过滤了的诊断记录集
'说明：LoadMedPageData的子函数
    Dim blnDo As Boolean
    '非病案首页，每一类别按诊断来源优先级获取，因此需要按类别判断是否全部保存
    Select Case intMaxDiagSource
        Case -1, 1, 2
            blnDo = True
        Case Else
            rsInput.Filter = "记录来源=" & intMaxDiagSource & " And 诊断类型=" & intDiagType
    End Select

    If blnDo Then
        If intMaxDiagSource > 0 Then '病案首页
            If Val(intDiagType) <> 21 Then
                If rsInput.EOF Then
                    rsInput.Filter = "记录来源=2 And 诊断类型=" & intDiagType
                End If
                If rsInput.EOF Then
                    rsInput.Filter = "记录来源=1 And 诊断类型=" & intDiagType
                End If
            End If
        Else '住院首页
            rsInput.Filter = "记录来源=3 And 诊断类型=" & intDiagType
            If Val(intDiagType) <> 21 Then
                If rsInput.EOF Then
                    rsInput.Filter = "记录来源=2 And 诊断类型=" & intDiagType
                End If
                If rsInput.EOF Then
                    rsInput.Filter = "记录来源=1 And 诊断类型=" & intDiagType
                End If
            End If
            If rsInput.EOF Then
                rsInput.Filter = "记录来源=4 And 诊断类型=" & intDiagType
            End If
        End If
    End If
End Sub

Public Function FindDiagRow(ByVal dtInput As DiagType) As Long
'功能：获取指定类型诊断的行号
'参数：dtInput=诊断类型
    Dim bln中医Diag  As Boolean
    Dim i As Long, LngRow As Long

    If dtInput <= DT_出院诊断ZY And dtInput > DT_并发症 Then bln中医Diag = True
    With IIf(bln中医Diag, gclsPros.CurrentForm.vsDiagZY, gclsPros.CurrentForm.vsDiagXY)
        LngRow = .FindRow(dtInput & "", , DI_诊断分类)
        FindDiagRow = LngRow
    End With
End Function

Public Sub CacheLoadDiagMatchData(Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnOnlyCache As Boolean)
'功能：加载诊断符合情况数据并缓存
'参数：rsInput=诊断符合情况记录集
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
    Dim arrCtrlIdxs() As Variant
    Dim arrInfoIdxs() As Variant
    Dim i As Long
    Dim objCboTmp As ComboBox
    Dim strTmp As String

    On Error GoTo errH
    If gclsPros.FuncType = f诊断选择 And gclsPros.PatiType = PF_住院 Then
        If gclsPros.IsTCM Then
            arrCtrlIdxs = Array(BCC_门诊与出院XY, BCC_入院与出院XY, BCC_放射与病理, BCC_临床与病理, BCC_临床与尸检, BCC_门诊与入院, _
                            BCC_门诊与出院ZY, BCC_入院与出院ZY)
            arrInfoIdxs = Array(1, 2, 3, 4, 5, 7, 11, 12)
        Else
            arrCtrlIdxs = Array(BCC_门诊与出院XY, BCC_入院与出院XY, BCC_放射与病理, BCC_临床与病理, BCC_临床与尸检, BCC_门诊与入院)
            arrInfoIdxs = Array(1, 2, 3, 4, 5, 7)
        End If
    Else
        If gclsPros.IsTCM Then
            arrCtrlIdxs = Array(BCC_门诊与出院XY, BCC_入院与出院XY, BCC_放射与病理, BCC_临床与病理, BCC_临床与尸检, BCC_术前与术后, BCC_门诊与入院, _
                            BCC_门诊与出院ZY, BCC_入院与出院ZY, BCC_辩证, BCC_治法, BCC_方药)
            arrInfoIdxs = Array(1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15)
        Else
            arrCtrlIdxs = Array(BCC_门诊与出院XY, BCC_入院与出院XY, BCC_放射与病理, BCC_临床与病理, BCC_临床与尸检, BCC_术前与术后, BCC_门诊与入院)
            arrInfoIdxs = Array(1, 2, 3, 4, 5, 6, 7)
        End If
    End If
    If Not blnOnlyCache Then
        For i = LBound(arrCtrlIdxs) To UBound(arrCtrlIdxs)
            '处理诊断符合情况缺省值
            Call SetDiagMatchInfo(arrCtrlIdxs(i))
        Next
    End If
    For i = LBound(arrCtrlIdxs) To UBound(arrCtrlIdxs)
        Set objCboTmp = gclsPros.CurrentForm.cboBaseInfo(arrCtrlIdxs(i))
        If Not blnOnlyCache Then
            If Not rsInput Is Nothing Then
                '加载诊断符合情况
                strTmp = ""
                rsInput.Filter = "符合类型=" & arrInfoIdxs(i)
                If Not rsInput.EOF Then
                    If Val(rsInput!符合情况 & "") >= 0 Then
                        Call zlControl.CboSetIndex(objCboTmp.hwnd, rsInput!符合情况)
                        strTmp = IIf(Not objCboTmp.Locked, rsInput!符合情况 & "", "")
                    End If
                End If
                Call UpdateCacheRecInfo(0, "诊断符合情况", strTmp, strTmp, arrCtrlIdxs(i))
            End If
        Else
            strTmp = ""
            If Not objCboTmp.Locked Then
                strTmp = IIf(objCboTmp.ListIndex = -1, "", objCboTmp.ListIndex)
            Else
                If arrCtrlIdxs(i) = BCC_临床与尸检 Then
                    strTmp = IIf(objCboTmp.ListIndex = -1, "", 4)
                End If
            End If
            Call UpdateCacheRecInfo(1, "诊断符合情况", strTmp, strTmp, arrCtrlIdxs(i))
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CompareDiag(ByVal strTmp1 As String, ByVal strTmp2 As String) As Boolean
'功能：对两种诊断进行对比，有一个诊断是相同的就返回 true
    Dim arrTmp1() As String
    Dim arrTmp2() As String
    Dim i As Long, j As Long

    arrTmp1 = Split(strTmp1, Chr(10))
    arrTmp2 = Split(strTmp2, Chr(10))

    For i = LBound(arrTmp1) To UBound(arrTmp1)
        For j = LBound(arrTmp2) To UBound(arrTmp2)
            If arrTmp1(i) = arrTmp2(j) Then
                CompareDiag = True
                Exit Function
            End If
        Next
    Next
End Function

Private Function GetStrDiagS(ByVal intIdx As Integer) As String
'功能：获取某一类型的所有诊断，用Chr(10)进行分隔
'参数：intIdx=诊断类型
    Dim i As Long
    Dim lngRow1 As Long, lngRow2 As Long
    Dim strTmp1 As String, strTmp2 As String
    Dim bln中医 As Boolean
    Dim vsDiag As VSFlexGrid

    If intIdx = DT_门诊诊断XY Then
        lngRow1 = FindDiagRow(DT_门诊诊断XY)
        lngRow2 = FindDiagRow(DT_入院诊断XY) - 1
    ElseIf intIdx = DT_入院诊断XY Then
        lngRow1 = FindDiagRow(DT_入院诊断XY)
        lngRow2 = FindDiagRow(DT_出院诊断XY) - 1
    ElseIf intIdx = DT_出院诊断XY Then
        lngRow1 = FindDiagRow(DT_出院诊断XY)
        lngRow2 = FindDiagRow(DT_院内感染) - 1
    ElseIf intIdx = DT_门诊诊断ZY Then
        bln中医 = True
        lngRow1 = FindDiagRow(DT_门诊诊断ZY)
        lngRow2 = FindDiagRow(DT_入院诊断ZY) - 1
    ElseIf intIdx = DT_入院诊断ZY Then
        bln中医 = True
        lngRow1 = FindDiagRow(DT_入院诊断ZY)
        lngRow2 = FindDiagRow(DT_出院诊断ZY) - 1
    ElseIf intIdx = DT_出院诊断ZY Then
        bln中医 = True
        lngRow1 = FindDiagRow(DT_出院诊断ZY)
        lngRow2 = gclsPros.CurrentForm.vsDiagZY.Rows - 1
    End If

    Set vsDiag = IIf(bln中医, gclsPros.CurrentForm.vsDiagZY, gclsPros.CurrentForm.vsDiagXY)
    For i = lngRow1 To lngRow2
        If Trim(vsDiag.TextMatrix(i, DI_诊断描述)) <> "" Then
            strTmp1 = strTmp1 & Chr(10) & Trim(vsDiag.TextMatrix(i, DI_诊断描述))
            strTmp2 = strTmp2 & Trim(vsDiag.TextMatrix(i, DI_诊断描述))
        End If
    Next

    If strTmp2 = "" Then
        strTmp1 = ""
    Else
        strTmp1 = Mid(strTmp1, InStr(strTmp1, Chr(10)) + 1)
    End If
    GetStrDiagS = strTmp1
End Function

Public Sub SetDiagMatchInfo(ByVal intIdx As Integer, Optional ByVal blnJustState As Boolean)
'功能：对诊断符合情况进行缺省值设置以及检查是否可以输入
'参数：intIdx=要设置的符合情况控件
'      blnJustState=只设置符合情况状态
    Dim i As Long
    Dim objCboTmp As ComboBox
    Dim strTmp1 As String, strTmp2 As String
    Dim lngHwnd As Long, lngIndex As Long
    Dim blnNotOther As Boolean
    Dim blnMedRecChange As Boolean

    With gclsPros.CurrentForm
        Set objCboTmp = .cboBaseInfo(intIdx)
        blnNotOther = True: lngIndex = -1
         '门诊与出院：门诊诊断和出院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        If intIdx = BCC_门诊与出院XY Then
            strTmp1 = GetStrDiagS(DT_门诊诊断XY)
            strTmp2 = GetStrDiagS(DT_出院诊断XY)
        '入院与出院：入院诊断和出院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        ElseIf intIdx = BCC_入院与出院XY Then
            strTmp1 = GetStrDiagS(DT_入院诊断XY)
            strTmp2 = GetStrDiagS(DT_出院诊断XY)
        '门诊与入院：门诊诊断和入院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        ElseIf intIdx = BCC_门诊与入院 Then
            strTmp1 = GetStrDiagS(DT_门诊诊断XY)
            strTmp2 = GetStrDiagS(DT_入院诊断XY)
        '中医门诊与出院：门诊诊断和出院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        ElseIf intIdx = BCC_门诊与出院ZY Then
            strTmp1 = GetStrDiagS(DT_门诊诊断ZY)
            strTmp2 = GetStrDiagS(DT_出院诊断ZY)
        '中医入院与出院：入院诊断和出院诊断相同时"符合"；其中一个不输入时"不肯定"；不同时"不符合"
        ElseIf intIdx = BCC_入院与出院ZY Then
            strTmp1 = GetStrDiagS(DT_入院诊断ZY)
            strTmp2 = GetStrDiagS(DT_出院诊断ZY)
        Else
           blnNotOther = False
        End If

        If blnNotOther Then
            If strTmp1 & strTmp2 = "" Then
                lngIndex = 0
            ElseIf strTmp1 = "" Or strTmp2 = "" Then
                lngIndex = 3
            Else
                lngIndex = IIf(CompareDiag(strTmp1, strTmp2), 1, 2)
            End If
        Else
            '放射与病理、临床与病理：录入病理诊断后可以录入，缺省为符合。
            If intIdx = BCC_放射与病理 Or intIdx = BCC_临床与病理 Then
                strTmp1 = .vsDiagXY.TextMatrix(FindDiagRow(DT_病理诊断), DI_诊断描述)
                Call SetCtrlLocked(objCboTmp, strTmp1 = "")
                If strTmp1 <> "" Then
                    lngIndex = 1
                    objCboTmp.BackColor = vbWindowBackground
                Else
                    lngIndex = 0
                    objCboTmp.BackColor = vbButtonFace
                End If
            '临床与尸检：勾选尸检后可以录入，缺省为符合。
            ElseIf intIdx = BCC_临床与尸检 Then
                Call SetCtrlLocked(objCboTmp, .cboBaseInfo(BCC_死亡患者尸检).ListIndex <= 0)
                If .cboBaseInfo(BCC_死亡患者尸检).ListIndex = 1 Then
                    lngIndex = 1
                    objCboTmp.BackColor = vbWindowBackground
                Else
                    lngIndex = 0
                    objCboTmp.BackColor = vbButtonFace
                End If
                blnMedRecChange = True
            '术前与术后：输入手术情况后可以录入，缺省为符合。
            ElseIf intIdx = BCC_术前与术后 Then
                For i = .vsOPS.FixedRows To .vsOPS.Rows - 1
                    If Trim(.vsOPS.TextMatrix(i, PI_手术名称)) <> "" Then Exit For
                Next
                If i > .vsOPS.Rows - 1 Then
                    lngIndex = 0                '不可以改时缺省为未做
                Else
                    lngIndex = 1
                End If
                blnMedRecChange = True
            End If
            If blnJustState Then
                lngIndex = objCboTmp.ListIndex
            End If
        End If
        
        If lngIndex > -1 Then
            If gclsPros.FuncType = f病案首页 Then
                If blnMedRecChange Then
                    Call zlControl.CboSetIndex(objCboTmp.hwnd, lngIndex)
                End If
            Else
                Call zlControl.CboSetIndex(objCboTmp.hwnd, lngIndex)
            End If
        End If
    End With
End Sub

Public Sub UpdateCacheRecInfo(Optional ByRef intType As Integer, Optional ByVal strInfoName As String, Optional ByVal strWholeInfo As String, Optional ByVal strMainInfo As String, Optional ByVal lngRowNo As Long = -1, Optional ByVal lngID As Long, Optional ByRef strTag As String)
'功能：更新或加载信息记录集，一般应用于表格
'参数：intType=0-初始化更新加载，1-保存前检查通过后的加载或更新,2-整体信息记录集更新，用于检查控件的改变状态
'      strInfoName=信息名或控件名
'      strWholeInfo=信息内容
'      strMainInfo=主信息值
'      lngRowNo=行号或控件序号
'      lngId=ID号
'      strTag=辅助定位信息,intType=0时填写，intType=1时过滤
'注意事项：扩展信息即次级信息记录集有数据的，strWholeInfo与strMainInfo都要传，两者可以相同
    Dim lngSort As Long, lng状态 As Long
    Dim lng序号 As Long
    Dim i As Long
    Dim rsTmp As ADODB.Recordset, strFilter As String
    Dim strTmp As String, arrTmp As Variant

    On Error GoTo errH
    strWholeInfo = Trim(strWholeInfo)
    strMainInfo = Trim(strMainInfo)
    strTag = Trim(strTag)
    If intType <> 2 Then
        '先依靠信息名寻找寻找，寻找不到时，再按控件名寻找
        gclsPros.MainInfoRec.Filter = "信息名='" & strInfoName & "'"
        If gclsPros.MainInfoRec.EOF Then gclsPros.MainInfoRec.Filter = "控件名='" & strInfoName & "'" & IIf(lngRowNo = -1, "", " And Index=" & lngRowNo)
        If Not gclsPros.MainInfoRec.EOF Then
            Select Case gclsPros.MainInfoRec!ExpState
                Case ES_不用扩展
                    Call gclsPros.MainInfoRec.Update(IIf(intType = 0, "信息原值", "信息现值"), strWholeInfo)
                Case ES_初始扩展
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And IndexEx=" & lngRowNo & IIf(strTag <> "" And intType = 1, " And Tag='" & strTag & "'", "")
                    If intType = 0 Then
                        Call gclsPros.SecdInfoRec.Update(Array("信息原值", "主信息原值", "Tag"), Array(strWholeInfo, strMainInfo, strTag))
                    Else
                        Call gclsPros.SecdInfoRec.Update(Array("信息现值", "主信息现值", "Tag"), Array(strWholeInfo, strMainInfo, strTag))
                    End If
                Case ES_加载扩展
                    gclsPros.SecdInfoRec.Filter = "": gclsPros.SecdInfoRec.Sort = "Sort"
                    If Not gclsPros.SecdInfoRec.EOF Then gclsPros.SecdInfoRec.MoveLast: lngSort = gclsPros.SecdInfoRec!Sort
                    If intType = 0 Then
                        gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And IndexEx=" & lngRowNo & IIf(strTag = "", "", " And Tag='" & strTag & "'")
                        If gclsPros.SecdInfoRec.EOF Then
                            Call gclsPros.SecdInfoRec.AddNew(Array("Sort", "序号", "改变状态", "ID", "页码", "控件名", "IndexEx", "信息原值", "主信息原值", "Tag"), Array(Identity(lngSort), gclsPros.MainInfoRec!序号, 0, IIf(lngID = 0, Null, lngID), gclsPros.MainInfoRec!页码, gclsPros.MainInfoRec!控件名, lngRowNo, strWholeInfo, strMainInfo, strTag))
                        End If
                    Else
                        'Tag为空的以主信息以过滤条件，否则以Tag为过滤条件，Tag过滤主要应用于病案附加项目，此时存储"行,列"
                        If strInfoName = "病案项目" Then
                            gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And Tag='" & Trim(strTag) & "'"
                        Else
                            gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And 主信息原值=" & IIf(strMainInfo = "", "Null", "'" & strMainInfo & "'") & IIf(strTag <> "", " And Tag='" & Trim(strTag) & "'", "")
                        End If
                        If Not gclsPros.SecdInfoRec.EOF Then
                            If gclsPros.SecdInfoRec.RecordCount > 1 Then
                                gclsPros.SecdInfoRec.MoveFirst
                                gclsPros.SecdInfoRec.Filter = "ID=" & gclsPros.SecdInfoRec!ID
                            End If
                            Call gclsPros.SecdInfoRec.Update(Array("IndexEx", "信息现值", "主信息现值", "Tag"), Array(lngRowNo, strWholeInfo, strMainInfo, IIf(strTag = "", gclsPros.SecdInfoRec!Tag, strTag)))
                        Else
                            Call gclsPros.SecdInfoRec.AddNew(Array("Sort", "序号", "改变状态", "页码", "控件名", "IndexEx", "信息现值", "主信息现值", "Tag", "改变状态"), Array(Identity(lngSort), gclsPros.MainInfoRec!序号, 0, gclsPros.MainInfoRec!页码, gclsPros.MainInfoRec!控件名, lngRowNo, strWholeInfo, strMainInfo, strTag, CS_新增行))
                        End If
                    End If
            End Select
        End If
    Else
        '分娩信息，直接过滤，在保存时使用
        If Not grsDeliceryInfo Is Nothing Then
            grsDeliceryInfo.Filter = "类型=0"
            For i = 1 To grsDeliceryInfo.RecordCount
                If grsDeliceryInfo!信息值 <> grsDeliceryInfo!信息现值 Then
                    grsDeliceryInfo.Update "记录性质", 1
                End If
                grsDeliceryInfo.Update
            Next
            grsDeliceryInfo.Filter = "记录性质=1": grsDeliceryInfo.Sort = "信息名"
            grsBabyInfo.Filter = "记录性质=1"
            grsBabyDiag.Filter = "记录性质=1"
        End If
        '更新次级信息记录集并判断主信息记录集
        gclsPros.SecdInfoRec.Filter = ""
        gclsPros.SecdInfoRec.Sort = "Sort"
        For i = 1 To gclsPros.SecdInfoRec.RecordCount
            lng状态 = CS_未改变
            If gclsPros.SecdInfoRec!信息原值 & "" <> gclsPros.SecdInfoRec!信息现值 & "" Then
                lng状态 = CS_更新行
            End If
            If lng状态 = CS_更新行 And IsNull(gclsPros.SecdInfoRec!信息原值) Then
                lng状态 = CS_新增行
            End If
            If lng状态 = CS_更新行 And IsNull(gclsPros.SecdInfoRec!信息现值) Then
                lng状态 = CS_删除行
            End If
            If lng状态 = CS_更新行 And gclsPros.SecdInfoRec!主信息原值 & "" <> gclsPros.SecdInfoRec!主信息现值 & "" Then
                lng状态 = CS_替换行
            End If
            If lng序号 <> gclsPros.SecdInfoRec!序号 And lng状态 <> CS_未改变 Then
                Call Rec.Update(gclsPros.MainInfoRec, "序号=" & gclsPros.SecdInfoRec!序号, "是否改变", 1)
                lng序号 = gclsPros.SecdInfoRec!序号
            End If
            gclsPros.SecdInfoRec.Update "改变状态", lng状态
            gclsPros.SecdInfoRec.MoveNext
        Next
        
        '更新主信息记录集
        gclsPros.MainInfoRec.Filter = "是否改变=0 And ExpState=" & ES_不用扩展
        gclsPros.MainInfoRec.Sort = "序号"
        For i = 1 To gclsPros.MainInfoRec.RecordCount
            If gclsPros.MainInfoRec!信息原值 & "" <> gclsPros.MainInfoRec!信息现值 & "" Then
                gclsPros.MainInfoRec.Update "是否改变", 1
            End If
            gclsPros.MainInfoRec.MoveNext
        Next

        gclsPros.MainInfoRec.Filter = "是否改变=1"
        If gclsPros.PatiType = PF_门诊 And gclsPros.FuncType = f医生首页 Then
             gclsPros.InfosChange = Not gclsPros.MainInfoRec.EOF Or gclsPros.CurrentForm.UCPatiVitalSigns.GetSaveSQL(gclsPros.病人ID, gclsPros.主页ID) <> "" Or gclsPros.IsLastDiag
        ElseIf gclsPros.FuncType = f病案首页 Then
            gclsPros.InfosChange = Not gclsPros.MainInfoRec.EOF
            If Not grsDeliceryInfo Is Nothing Then
                gclsPros.InfosChange = gclsPros.InfosChange Or Not grsDeliceryInfo.EOF
            End If
            If Not grsBabyInfo Is Nothing Then
                gclsPros.InfosChange = gclsPros.InfosChange Or Not grsBabyInfo.EOF
                If Not grsBabyDiag Is Nothing Then
                    gclsPros.InfosChange = gclsPros.InfosChange Or Not grsBabyDiag.EOF
                End If
            End If
        Else
            gclsPros.InfosChange = Not gclsPros.MainInfoRec.EOF
        End If
        '诊断信息，手术信息,过敏信息可能是从其他来源读取胡，若是，且界面改变，则需要全表格保存
        If gclsPros.InfosChange Or gclsPros.DiagSel Then
            strTmp = "西医诊断;中医诊断;过敏药物;手麻情况"
            arrTmp = Split(strTmp, ";")
            For i = LBound(arrTmp) To UBound(arrTmp)
                gclsPros.MainInfoRec.Filter = "信息名='" & arrTmp(i) & "'"
                If Not gclsPros.MainInfoRec.EOF Then
                    '新增，修改或者替换，未改变的其他来源的诊断行均视作新增
                    Call Rec.Update(gclsPros.SecdInfoRec, "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态>=0 And Tag<> '" & IIf(gclsPros.FuncType = f病案首页, 4, 3) & "'", "改变状态", CS_新增行)
                    '删除的其他来源的诊断行均视作新增
                    Call Rec.Update(gclsPros.SecdInfoRec, "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态<0 And Tag<> '" & IIf(gclsPros.FuncType = f病案首页, 4, 3) & "'", "改变状态", CS_未改变)
                    If gclsPros.IsLastDiag And arrTmp(i) Like "*诊断" Then
                        '新增，修改或者替换，未改变的行均视作新增
                        Call Rec.Update(gclsPros.SecdInfoRec, "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态>=0", "改变状态", CS_新增行)
                         '上次挂号诊断，删除行不用处理
                        Call Rec.Update(gclsPros.SecdInfoRec, "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态<0", "改变状态", CS_未改变)
                    End If
                    gclsPros.SecdInfoRec.Filter = "序号=" & gclsPros.MainInfoRec!序号 & " And 改变状态<>0"
                    If Not gclsPros.SecdInfoRec.EOF Then
                        gclsPros.MainInfoRec.Update "是否改变", 1
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
errH:
    Debug.Print "UpdateCacheRecInfo:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetAllerInput(ByVal LngRow As Long, Optional rsInput As ADODB.Recordset, Optional ByVal strTYTInput As String)
'功能：处理过敏药物的输入
'参数：strTYTInput=太元通合理用药接口返回的字符串
'    Dim strSql As String, curDate As Date
    Dim arrTmp As Variant
    Dim strAllerOld As String, strAllerNew As String

    With gclsPros.CurrentForm.vsAller

        strAllerOld = .Cell(flexcpData, LngRow, AI_过敏药物) & ";" & .TextMatrix(LngRow, AI_过敏源编码)

        If gclsPros.UseTYT Then
            arrTmp = Split(strTYTInput, ";")

            If UBound(arrTmp) < 1 Then Exit Sub
            If strAllerOld <> strTYTInput Or Val(.RowData(LngRow) & "") <> 0 Then
                .TextMatrix(LngRow, AI_过敏药物) = arrTmp(1)
                .TextMatrix(LngRow, AI_过敏源编码) = arrTmp(0)
                .RowData(LngRow) = 0
            End If
        Else
            If gclsPros.FuncType <> f病案首页 Then
                If gclsPros.CurrentForm.optAller(PC_按药品目录输入).Value Then
                    If Not rsInput Is Nothing Then
                        .RowData(LngRow) = CLng(rsInput!ID)
                        .TextMatrix(LngRow, AI_过敏药物) = NVL(rsInput!名称)
                    Else
                        .RowData(LngRow) = 0
                        .TextMatrix(LngRow, AI_过敏药物) = .EditText
                    End If
    
                    strAllerNew = .TextMatrix(LngRow, AI_过敏药物) & ";" & .TextMatrix(LngRow, AI_过敏源编码)
    
                    If strAllerOld <> strAllerNew Or Val(.RowData(LngRow) & "") <> 0 Then
                        .TextMatrix(LngRow, AI_过敏源编码) = ""
                    End If
                Else
                    If Not rsInput Is Nothing Then
                        .TextMatrix(LngRow, AI_过敏药物) = rsInput!名称 & ""
                        .TextMatrix(LngRow, AI_过敏源编码) = rsInput!编码 & ""
                        .RowData(LngRow) = 0
                    Else
                        .RowData(LngRow) = 0
                        .TextMatrix(LngRow, AI_过敏药物) = .EditText
                    End If
                End If
            Else
                If Not rsInput Is Nothing Then
                    .RowData(LngRow) = CLng(rsInput!ID)
                    .TextMatrix(LngRow, AI_过敏药物) = NVL(rsInput!名称)
                Else
                    .RowData(LngRow) = 0
                    .TextMatrix(LngRow, AI_过敏药物) = .EditText
                End If

                strAllerNew = .TextMatrix(LngRow, AI_过敏药物) & ";" & .TextMatrix(LngRow, AI_过敏源编码)

                If strAllerOld <> strAllerNew Or Val(.RowData(LngRow) & "") <> 0 Then
                    .TextMatrix(LngRow, AI_过敏源编码) = ""
                End If
            End If
        End If

        .Cell(flexcpData, LngRow, AI_过敏药物) = .TextMatrix(LngRow, AI_过敏药物)
        .TextMatrix(LngRow, AI_药物ID) = Val(.RowData(LngRow) & "")
'        If .Cell(flexcpData, LngRow, AI_过敏时间) = "" Then
'            curDate = zlDatabase.Currentdate
'            .TextMatrix(LngRow, AI_过敏时间) = Format(curDate, "yyyy-MM-dd")
'            .Cell(flexcpData, LngRow, AI_过敏时间) = Format(curDate, "yyyy-MM-dd")
'        End If
        '始终保持一空行
        If LngRow = .Rows - 1 Then
            .AddItem "", LngRow + 1
            Call ChangeVSFHeight(gclsPros.CurrentForm.vsAller, True, 0)
        End If
    End With
End Sub

Public Sub AllerEnterNextCell()
    Dim i As Long, j As Long

    With gclsPros.CurrentForm.vsAller
        If .Col = AI_过敏时间 Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = AI_过敏药物
                 Call .Select(.Row, .Col)
                 '用下面这个方法的话，内容有时会出现滚动的情况
'                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Public Sub zlVsGridRowChange(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngOldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：行列改变时,设置相关的颜色
    '入参：CustomColor-自定义颜色
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-23 11:22:38
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    '行改变时
    Err = 0: On Error Resume Next
    If lngOldRow = lngNewRow Then
        vsGrid.Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, vsGrid.Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
        Exit Sub
    End If
    With vsGrid
        .Cell(flexcpBackColor, lngOldRow, vsGrid.FixedCols, lngOldRow, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, .Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
    End With
End Sub

Public Sub zlVsGridGotFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：进入网格控件时选择的颜色
    '入参：CustomColor-自定颜色
    '编制：刘兴洪
    '日期：2010-03-23 10:52:23
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    '进入控件
    With vsGrid
         If CustomColor <> -1 Then
             .FocusRect = flexFocusSolid
             .HighLight = flexHighlightNever
             If .Row >= .FixedRows Then
                If .Rows - 1 > .FixedRows Then  '清除选择颜色
                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
                End If
                 .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
             End If
         Else
            .FocusRect = flexFocusSolid 'IIf(vsGrid.Editable = flexEDNone, flexFocusNone, flexFocusSolid)
            .HighLight = flexHighlightNever
            .BackColorSel = GRD_GOTFOCUS_COLORSEL
        End If
    End With
    Call zlVsGridRowChange(vsGrid, vsGrid.Row, vsGrid.Row, 0, 0)
End Sub

Public Sub zlVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng主例 As Long = -1, Optional lng尾列 As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef LngRow As Long = -1)
    '-----------------------------------------------------------------------------------------------------------
    '功能:移动单元格的列
    '入参:blnEdit-当前正处于编辑状态,允许新增行
    '     lng主例-主列,如果<0,则主列为0列,否则为指定的列
    '     lng尾列-尾列,如果<0,则主列为.cols-1,否则为指定的列
    '出参:lngRow-如果存在插入行,则返回被插入的行号,否则返回-1
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------
    Dim LngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long

    Err = 0: On Error GoTo Errhand:

    'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
    If lng主例 <> -1 Then
        LngCol = lng主例
    Else
        LngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If LngCol = -1 Then LngCol = 0
    lngLastCol = IIf(lng尾列 < 0, vsGrid.Cols - 1, lng尾列)
    LngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = LngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, LngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        LngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                arrSplit = Split(.ColData(i) & "||", "||")
                If .ColHidden(i) Or Val(arrSplit(1)) >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = LngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, LngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    LngRow = .Row
                                End If
                            End If
                            .Col = LngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
Errhand:
End Sub

Public Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal LngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '功能:插入行
    '参数:vsGrid-插入行的网格格件
    '     lngRow-当前行
    '     blnBefor-在lngrow之间或之后.true:之间,false-之后
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        If blnBefor Then
            .AddItem "", LngRow
        Else
            .AddItem "", LngRow + 1
        End If
        Call ChangeVSFHeight(vsGrid, True)
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = LngRow
            Else
                .Row = LngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub VsFlxGridCheckKeyPress(ByVal objCtl As Object, ByRef LngRow As Long, ByRef LngCol As Long, ByRef intKeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '功能:只能输入数字和回车及退格
    '参数:
    '   objctl:Vsgrid8.0控件
    '   intKeyascii:
    '           Keyascii:8 (退格)
    '   Row-当前行
    '   Col-当前列
    '   TextType:(0-文本式;1-数字式;2-金额式)
    '返回:一个KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:

    If TextType = m文本式 Then
        If intKeyAscii = Asc("'") Then
            intKeyAscii = 0
        End If
        Exit Sub
    End If

    If intKeyAscii < Asc("0") Or intKeyAscii > Asc("9") Then
        Select Case intKeyAscii
        Case vbKeyReturn       '回车

        Case 8                 '退格

        Case Asc(".")
            If TextType = m金额式 Or TextType = m负金额式 Then
                If InStr(objCtl.EditText, ".") <> 0 Then     '只能存在一个小数点
                    intKeyAscii = 0
                End If
            Else
                intKeyAscii = 0
            End If
        Case Asc("-")          '负数
            Dim iRow As Long
            Dim icol As Long
            If Trim(objCtl.EditText) = "" Then Exit Sub
            If TextType <> m负金额式 Then intKeyAscii = 0: Exit Sub
            If objCtl.EditSelStart <> 0 Then intKeyAscii = 0: Exit Sub      '光标不存第一位,不能输入负数
            If InStr(1, objCtl.EditText, "-") <> 0 Then   '只能存在一个负数
                intKeyAscii = 0
            End If
        Case Else
            intKeyAscii = 0
        End Select
    End If
    Exit Sub
Errhand:
    intKeyAscii = 0
End Sub


Public Sub TSJCSetDiagInput(ByVal LngRow As Long, rsInput As ADODB.Recordset)
'功能：处理特殊检查项目的输入
    With gclsPros.CurrentForm.vsTSJC
        If Not rsInput Is Nothing Then
            .TextMatrix(LngRow, 1) = NVL(rsInput!名称)
        Else
            .TextMatrix(LngRow, 1) = .EditText
        End If
        .Cell(flexcpData, LngRow, 1) = .TextMatrix(LngRow, 1)
    End With
End Sub

Public Sub TSJCEnterNextCell()
    With gclsPros.CurrentForm.vsTSJC
        If .Row = .Rows - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            If .Row + 1 > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Public Function DiagCellEditable(ByRef vsDiagTmp As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long) As Boolean
    Dim bln西医 As Boolean
    Dim blnJudge As Boolean
    Dim dtTmp As DiagType
    Dim lng出院Row As Long

    With vsDiagTmp
        bln西医 = .Name = "vsDiagXY"
        '隐藏列不可编辑
        If .ColHidden(LngCol) Then Exit Function
        '首页已经签名的诊断选择器中不允许修改诊断
        If gclsPros.FuncType = f诊断选择 And gclsPros.IsSigned And LngCol <> DI_关联 Then Exit Function
        '必须先输入诊断描述，才能输入其他列(部分逻辑属于新增）
        If .TextMatrix(LngRow, DI_诊断描述) = "" Then
            If Not (gclsPros.FuncType = f病案首页 And LngCol = DI_诊断编码 Or LngCol = DI_Del Or LngCol = DI_诊断描述) Then Exit Function
        ElseIf gclsPros.FuncType <> f病案首页 Then
            If LngCol <> DI_关联 And LngCol <> DI_Del And LngCol <> DI_增加 Then
                '关联医嘱不可编辑
                If gclsPros.FuncType = f诊断选择 Then '需要排除当前医嘱或申请单
                    If GetAdviceIDByDiag(.TextMatrix(LngRow, DI_医嘱IDs), Val(.RowData(LngRow))) <> "" Then Exit Function
                Else
                    If .TextMatrix(LngRow, DI_医嘱IDs) <> "" Then Exit Function
                End If
            End If
        End If
        Select Case LngCol
            Case DI_诊断描述
                If gclsPros.FuncType <> f病案首页 Then
                    '临床路径诊断不允许改
                    If gclsPros.PathState = PS_执行中 Or gclsPros.PathState = PS_正常结束 Then
                        If Not CheckMergePath(gclsPros.病人ID, gclsPros.主页ID, Val(.TextMatrix(LngRow, DI_诊断分类)), Val(.TextMatrix(LngRow, DI_疾病ID))) Then Exit Function
                    End If
                    '两条路径以上，导入诊断不允许改
                    If gclsPros.PathDiag <> "" And gclsPros.PathState > PS_不符合导入 Then
                        If InStr("," & gclsPros.PathDiag & ",", "," & .TextMatrix(.Row, DI_诊断分类) & "|" & Val(.TextMatrix(.Row, DI_疾病ID)) & "|" & Val(.TextMatrix(.Row, DI_诊断ID)) & ",") > 0 Then
                            Exit Function
                        End If
                    End If
                    '正常完成的出院诊断不允许改
                    If gclsPros.PathState = PS_正常结束 And gclsPros.PathOutTime Then
                        If bln西医 Then
                            blnJudge = .TextMatrix(.Row, DI_诊断类型) = "出院诊断" And gclsPros.InPath <= DT_入院诊断XY
                        Else
                            blnJudge = .TextMatrix(.Row, DI_诊断类型) = "出院诊断" And gclsPros.InPath >= DT_门诊诊断ZY
                        End If
                        If blnJudge Then Exit Function
                    End If
                End If
            Case DI_诊断编码
                '病案首页诊断编码与描述相互独立，可以输入诊断编码（为了保证首页中提取的自由录入诊断具有诊断编码的问题）
                '病案可以输入编码，为了输入的时候查看诊断描述
                If gclsPros.FuncType <> f病案首页 Then
                    Exit Function
                End If
            Case DI_ICD附码
                '病案首页，门诊与入院不能输入附码,如果主码存在固定附码，同样不能输入附码
                If Not bln西医 Then Exit Function
                If .TextMatrix(.Row, DI_固定附码) = "1" Or Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_门诊诊断XY Or _
                    Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_入院诊断XY Then
                    Exit Function
                End If
            Case DI_出院情况
                If bln西医 Then
                    '出院诊断和院内感染允许输入出院情况(因为可能院内感染在出院时已经好转或治愈了)
                    blnJudge = Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_出院诊断XY Or Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_院内感染 Or Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_并发症
                Else
                    '非出院诊断时不允许输入
                    blnJudge = Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_出院诊断ZY
                End If
                If Not blnJudge Then Exit Function
                If gclsPros.FuncType = f病案首页 Then
                    If .TextMatrix(LngRow, DI_是否病人) <> "1" Then Exit Function
                End If
            Case DI_入院病情
                '入院病情只能在出院诊断和其他诊断行填写,西医的并发症与院内感染也可以填写
                If bln西医 Then
                    If Val(.TextMatrix(LngRow, DI_诊断分类)) <> DT_出院诊断XY And Val(.TextMatrix(LngRow, DI_诊断分类)) <> DT_并发症 And Val(.TextMatrix(LngRow, DI_诊断分类)) <> DT_院内感染 Then Exit Function
                Else
                    If Val(.TextMatrix(LngRow, DI_诊断分类)) <> DT_出院诊断ZY Then Exit Function
                End If
            Case DI_是否未治 '中医未治列隐藏
                blnJudge = Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_出院诊断XY Or Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_院内感染 Or Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_并发症
                '出院诊断和院内感染允许输入是否未治(因为可能院内感染在出院时已经好转或治愈了)
                If Not blnJudge Then Exit Function
                '出院情况为"其他"时才可以设置是否未治
                If .TextMatrix(LngRow, DI_出院情况) <> "其他" Then Exit Function
            Case DI_增加
                '出院主要诊断不允许增加
                If bln西医 Then
                    blnJudge = .TextMatrix(LngRow, DI_诊断类型) = "出院诊断" And Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_出院诊断XY
                Else
                    blnJudge = .TextMatrix(LngRow, DI_诊断类型) = "主要诊断" And Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_出院诊断ZY
                End If
                If blnJudge Then Exit Function
                '同类型下一行诊断为空，则不允许增加
                If LngRow <> .Rows - 1 Then
                    blnJudge = .TextMatrix(LngRow, DI_诊断分类) = .TextMatrix(LngRow + 1, DI_诊断分类) And .TextMatrix(LngRow, DI_诊断描述) <> "" And .TextMatrix(LngRow + 1, DI_诊断描述) = ""
                    If blnJudge Then Exit Function
                End If
        End Select

        '出院诊断必须依次输入(尚未输入时)
        dtTmp = IIf(bln西医, DT_出院诊断XY, DT_出院诊断ZY)
        If .TextMatrix(LngRow, DI_诊断描述) = "" And Val(.TextMatrix(LngRow, DI_诊断分类)) = dtTmp Then
            If .TextMatrix(LngRow - 1, DI_诊断描述) = "" And Val(.TextMatrix(LngRow - 1, DI_诊断分类)) = dtTmp Then
                Exit Function
            End If
        End If

        If gclsPros.FuncType = f病案首页 And bln西医 Then
            lng出院Row = FindDiagRow(DT_出院诊断XY)
            If Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_损伤中毒码 Then
                If (.TextMatrix(lng出院Row, DI_诊断编码) = "" And .TextMatrix(lng出院Row, DI_诊断描述) = "") Or InStr("ST", Left(.TextMatrix(lng出院Row, DI_诊断编码), 1)) = 0 Then

                ElseIf .TextMatrix(LngRow, DI_固定附码) = "1" And InStr("VWXY", Left(.TextMatrix(LngRow, DI_ICD附码), 1)) > 0 Then
                '固定附码不允许修改
                    Exit Function
                End If
            End If
        End If
        DiagCellEditable = True
    End With
End Function

Public Function GetMedInputSQL(ByVal intType As Integer, ByVal strInput As String, ByRef str性别 As String, Optional ByVal strOtherInfo As String) As String
'功能：获得查询首页输入查询的SQL
'参数：intType:获取的SQL类型,0-西医诊断，1-中医诊断，2-手术操作
'    strInput-查询条件，str性别--病人的性别
'   strOtherInfo:中医诊断-疾病编码种类；西医诊断-诊断类型
'返回：strsql--查询诊断的SQL
    Dim strSql As String

    If gclsPros.Sex Like "*男*" Then
        str性别 = "男"
    ElseIf gclsPros.Sex Like "*女*" Then
        str性别 = "女"
    End If

    Select Case intType
        Case 0, 1 '西医诊断,中医诊断
            If intType = 0 And gclsPros.DiagInputXY = 0 Or intType = 1 And gclsPros.DiagInputZY = 0 And strOtherInfo <> "Z" Then
            '按诊断输入:一个诊断可能属于多个分类
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "B.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSql = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If
                strSql = "Select A.Id, A.Id 项目ID, A.编码, Null 序号, Null 附码, Null 附码id, Null 附码名称, A.名称, A.说明, A.编者, B.简码, 0 疗效限制, 0 分娩," & vbNewLine & _
                                "              0 是否病人, Max(D.疾病id) 疾病id, A.Id 诊断id" & vbNewLine & _
                                "       From 疾病诊断目录 A, 疾病诊断别名 B, 疾病诊断对照 D" & vbNewLine & _
                                " Where A.ID=B.诊断ID And A.ID=D.诊断ID(+) And A.类别=" & IIf(intType = 0, 1, 2) & vbNewLine & _
                                " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                                " And B.码类=[5] And (" & strSql & ")" & vbNewLine & _
                                "Group By A.Id, A.编码, A.名称, A.说明, A.编者,B.简码"
                '读取诊断对应疾病编码附码
                strSql = "Select distinct A.ID,A.项目ID, A.编码, B.序号, B.附码, Null 附码id, Null 附码名称, A.名称, A.说明, Null 编者,A.简码, A.疗效限制, A.分娩, A.是否病人," & vbNewLine & _
                                "       B.编码 疾病编码, B.Id 疾病id, B.类别 疾病类别, A.诊断id," & vbNewLine & _
                                "                 Decode(a.名称, [6], 1, Decode(A.简码,[6],1,decode(A.编码,[6],1,NULL))) As 排序1ID,Decode(d.诊断id, Null, Decode(c.诊断id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                                "                 Decode(Substr(A.名称, 1, Length([6])), [6], 1, Decode(Substr(A.简码, 1, Length([6])),[6],1,decode(Substr(a.编码, 1, Length([6])),[6],1,NULL))) As 排序3ID" & _
                                " From (" & strSql & ") A, 疾病编码目录 B, 疾病诊断科室 C, 疾病诊断科室 D" & vbNewLine & _
                                " Where A.疾病id = B.Id(+)" & vbNewLine & _
                                " And c.诊断id(+) = a.Id And d.诊断id(+) = a.Id And c.科室id(+)=[8]  And d.人员id(+) = [7]" & _
                                " Order By 排序1ID, 排序2ID, 排序3ID, A.编码"
            Else
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "A.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSql = "A.编码 Like [1] Or A.名称 Like [2] Or " & IIf(gclsPros.BriefCode = 0, "A.简码", "A.五笔码") & " Like [2]"
                End If
                If gclsPros.FuncType = f病案首页 Then
                    strSql = _
                        "Select A.Id, A.Id 项目ID,A.编码, A.序号, A.附码,Null 附码ID, Null 附码名称, A.名称, A.说明, Null 编者,A.分类id, " & IIf(gclsPros.BriefCode = 0, "A.简码", "A.五笔码") & " as 简码,  A.疗效限制, A.分娩, C.是否病人,A.编码 疾病编码, A.Id 疾病id,A.类别 疾病类别, Null 诊断id" & vbNewLine & _
                        "From 疾病编码目录 A, 疾病编码分类 C" & vbNewLine & _
                        "Where A.分类id = C.Id(+) And Instr([3],A.类别)>0 And (" & strSql & ")" & _
                        IIf(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is NULL)", "") & _
                        " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by A.编码"
                Else
                    strSql = _
                        "Select A.Id,A.Id 项目ID, A.编码, A.序号, A.附码,Null 附码ID, Null 附码名称, A.名称, A.说明, Null 编者, A.分类id, " & IIf(gclsPros.BriefCode = 0, "A.简码", "A.五笔码") & " as 简码,  A.疗效限制, A.分娩, C.是否病人,A.编码 疾病编码, A.Id 疾病id,A.类别 疾病类别," & vbNewLine & _
                        "       Max(B.诊断id) 诊断id" & vbNewLine & _
                        "From 疾病编码目录 A, 疾病诊断对照 B, 疾病编码分类 C " & vbNewLine & _
                        "Where A.Id = B.疾病id(+) And A.分类id = C.Id(+)  And" & vbNewLine & _
                        " Instr([3],A.类别)>0 And (" & strSql & ")" & _
                        IIf(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is NULL)", "") & _
                        " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        "Group By A.Id, A.编码, A.序号, A.附码, A.名称, A.说明, A.分类id, " & IIf(gclsPros.BriefCode = 0, "A.简码", "A.五笔码") & ", A.疗效限制, A.分娩, A.类别,C.是否病人"
                End If
                strSql = "Select distinct A.Id,A.项目ID, A.编码, A.序号, A.附码,A.附码ID, A.附码名称, A.名称, A.说明, A.编者, A.分类id, A.简码,  A.疗效限制, A.分娩, A.是否病人,A.疾病编码, A.疾病id,A.疾病类别,A.诊断id, " & _
                        " Decode(a.名称, [6], 1, Decode(A.简码,[6],1,decode(a.编码,[6],1,NULL))) As 排序1ID," & vbNewLine & _
                "                Decode(d.疾病id, Null, Decode(c.疾病id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                "                Decode(Substr(a.名称, 1, Length([6])), [6], 1, Decode(Substr(A.简码, 1, Length([6])),[6],1,decode(Substr(a.编码, 1, Length([6])),[6],1,NULL))) As 排序3ID" & vbNewLine & _
                        " From (" & strSql & ") A, 疾病编码科室 C, 疾病编码科室 D " & _
                        " Where  c.疾病id(+) = a.Id And d.疾病id(+) = a.Id And c.科室id(+)=[8]  And d.人员id(+) = [7] " & _
                        " Order By" & IIf(strOtherInfo = "'M,D'", " 疾病类别 desc , ", "") & " 排序1ID, 排序2ID, 排序3ID, A.编码"
            End If
        Case 2 '手术操作
            If gclsPros.OPSInput = 0 And gclsPros.FuncType <> f病案首页 Then
                '按诊疗项目输入
                strSql = "Select distinct A.ID,A.编码,A.名称,A.操作类型 as 规模" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B" & _
                    " Where A.类别='F' And A.服务对象 IN(2,3) And A.ID=B.诊疗项目ID" & _
                    IIf(str性别 <> "", " And Nvl(A.适用性别,0) IN(0,[4])", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.编码 Like [1] Or A.名称 Like [2] Or B.简码 Like [2] Or B.名称 Like [2])" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编码"
            Else
                '按ICD9-CM3输入
                strSql = " Select distinct ID,编码,附码,名称,简码,说明" & _
                    " From 疾病编码目录 Where 类别='S'" & _
                    IIf(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And (编码 Like [1] Or 名称 Like [2] Or 简码 Like [2])" & _
                    " Order by 编码"
            End If
    End Select
    GetMedInputSQL = strSql
End Function

Public Function OPSCellEditable(ByVal LngRow As Long, ByVal LngCol As Long) As Boolean
    Dim vsTemp As VSFlexGrid
    Set vsTemp = gclsPros.CurrentForm.vsOPS

    With vsTemp
        If .ColHidden(LngCol) Then Exit Function

        '必须先输入手术日期,手术名称
        If Not IsDate(.TextMatrix(LngRow, PI_手术日期)) Then
            If LngCol > PI_手术日期 Then Exit Function
        End If
        If .TextMatrix(LngRow, PI_手术名称) = "" Then
            If LngCol > PI_手术名称 Then Exit Function
        End If

        '必须先输入主刀医师
        If .TextMatrix(LngRow, PI_主刀医师) = "" Then
            If LngCol = PI_助手1 Or LngCol = PI_助手2 Then Exit Function
        End If

        '必须先输入第1助手
        If .TextMatrix(LngRow, PI_助手1) = "" Then
            If LngCol = PI_助手2 Then Exit Function
        End If

        '必须先输入麻醉类型
        If Trim(.TextMatrix(LngRow, PI_麻醉类型)) = "" Then
            If LngCol = PI_麻醉医师 Then Exit Function
        End If
        If gclsPros.FuncType <> f病案首页 Then
            '手术名称不能输入
            If LngCol = PI_手术名称 And gclsPros.CurrentForm.chkParaOPSInfo(PC_未找到时自由录入).Value = 0 Then Exit Function
        Else
            If LngCol = PI_手术名称 And Not gclsPros.CNIndent Then Exit Function
        End If

        '提取的手术级别不允许更改
        If LngCol = PI_手术级别 Then
            If gclsPros.Module = p住院医生站 Then
                If .Cell(flexcpData, LngRow, LngCol) = 1 And InStr(GetInsidePrivs(p住院医生站), "修改手术等级") = 0 Then Exit Function
            Else
                If .Cell(flexcpData, LngRow, LngCol) = 1 Then Exit Function
            End If
        End If
        
    End With
    OPSCellEditable = True
End Function

Public Function CheckIsDate(ByVal strKEY As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '功能:检查是否合法的日期型,可以为:(20070101或2007-01-01)或则(01-01或0101)或则(01<01-31>)
    '参数:strKey-需要检查的关建字
    '返回:合法的日期,返回标准格式(yyyy-mm-dd),否则返回""
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKEY) = 4 And InStr(1, strKEY, "-") = 0 Then
        '0101,需要再前面加年
        strKEY = Year(Now) & strKEY
    ElseIf Len(Replace(strKEY, "-", "")) = 4 And InStr(1, strKEY, "-") > 0 Then
        '01-01形式,需要补零
        strKEY = Year(Now) & Replace(strKEY, "-", "")
    ElseIf Len(strKEY) <= 2 And IsNumeric(strKEY) Then
        '指是日
        strKEY = Format(Now, "YYYYMM") & IIf(Len(strKEY) = 2, strKEY, "0" & strKEY)
    End If
    If Len(strKEY) = 8 And InStr(1, strKEY, "-") = 0 Then
        strKEY = TranNumToDate(strKEY)
        If strKEY = "" Then
            MsgBox strTittle & "必须为日期型,请检查！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Not IsDate(strKEY) Then
        MsgBox strTittle & "必须为日期型如(2000-10-10) 或（20001010）,请检查！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckIsDate = strKEY
End Function

Public Function Check日期有效性(ByVal strDate As String, ByVal strTittle As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查日期的有效性
    '入参:strDate-当前日期
    '     strTittle-标题:如:放疗在第几行
    '出参:
    '返回:有效或strDate="",返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-21 17:03:30
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strCurDate As String

    If strDate = "" Then Check日期有效性 = True: Exit Function
    '检查日期是否合法
    If IsDate(strDate) = False Or IsNumeric(strDate) Then
        MsgBox strTittle & "不是一个有效的日期范围,请检查!", vbInformation, gstrSysName
        Exit Function
    End If

    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    If CDate(strDate) > CDate(strCurDate) Then
        MsgBox strTittle & "比当前日期还要大,请检查!", vbInformation, gstrSysName
        Exit Function
    End If

    If CDate(strDate) < CDate(gclsPros.InTime) Then
        MsgBox strTittle & "比入院日期还要小,请检查!", vbInformation, gstrSysName
        Exit Function
    End If

    If gclsPros.OutTime <> "" Then
        If CDate(gclsPros.OutTime) < CDate(strDate) Then
            MsgBox strTittle & "比出院日期还要大,请检查!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Check日期有效性 = True
End Function

Public Function DblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hwnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     blnNegative     是否进行负数检查
    '     blnZero         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------

    Dim dblValue As Double

    If blnZero = True Then
        If strInput = "" Then
            MsgBox str项目 & "未输入，请检查!", vbInformation, gstrSysName
            If hwnd <> 0 Then SetFocusHwnd hwnd
            Exit Function
        End If
    End If
    If strInput = "" Then DblIsValid = True: Exit Function
    If IsNumeric(strInput) = False Then
        MsgBox str项目 & "不是有效的数字格式。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If

    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str项目 & "数值过大，不能超过" & 10 ^ intMax - 1 & "。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    If blnNegative = True And dblValue < 0 Then
        MsgBox str项目 & "不能输入负数。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If

    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str项目 & "数值过小，不能小于-" & 10 ^ intMax - 1 & "位。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If

    If blnZero = True And dblValue = 0 Then
        MsgBox str项目 & "不能输入零。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    DblIsValid = True
End Function

Public Function CheckInPutIsDate(ByVal vsObj As Object, LngRow As Long, LngCol As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能:检查所输入的日期是否合法
    '参数:lngRow -行,lngCol -列
    '返回:日期合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/05/21
    '---------------------------------------------------------------------------------------------------------
    Dim strKEY As String
    Dim str进入时间 As String, str退出时间 As String

    strKEY = Trim(vsObj.EditText)
    strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
    strKEY = Replace(strKEY, Chr(10), "")
    strKEY = zlStr.FullDate(strKEY, , gclsPros.InTime, gclsPros.OutTime)
    If strKEY <> "" Then
        If Not IsDate(strKEY) Then
            MsgBox vsObj.TextMatrix(0, LngCol) & "必须为日期型,请重新输入！", vbInformation + vbDefaultButton1, gstrSysName
             vsObj.EditSelStart = 0
             vsObj.EditSelLength = 1000
            Exit Function
        End If

        Select Case LngCol
           Case UI_进入时间
               str进入时间 = strKEY
               str退出时间 = Trim(vsObj.TextMatrix(LngRow, UI_退出时间))
               If str退出时间 <> "" And str进入时间 >= str退出时间 Then
                   MsgBox "注:" & vbCrLf & "  进入时间大于了退出时间,请检查！", vbInformation + vbDefaultButton1, gstrSysName
                   Exit Function
               End If
           Case UI_退出时间
               str进入时间 = Trim(vsObj.TextMatrix(LngRow, UI_进入时间))
               str退出时间 = strKEY

               If str进入时间 <> "" And CDate(str进入时间) >= CDate(str退出时间) Then
                   MsgBox "注:" & vbCrLf & "  退出时间小于了进入时间,请检查！", vbInformation + vbDefaultButton1, gstrSysName
                   Exit Function
               End If
        End Select
    End If
    CheckInPutIsDate = True
End Function

Public Sub ChangePage(Optional ByVal blnForWord As Boolean = True, Optional ByVal lngPage As Long = -1, Optional ByRef objTmp As Object, Optional blnLocation As Boolean = True)
'功能：选项卡定位到下一页
'参数：blnForWord=是否向前翻页，在最后一页时，定位到第一页，false=向后翻页
'      lngPage=指定页，不等于-1时，不进行翻页，直接定位页面
    Dim lngCurPage As Long, i As Long
    Dim lngHeight As Long
    Dim lngMin As Long, lngMax As Long
    Dim blnCur As Boolean
  
    lngMin = -1
    With gclsPros.CurrentForm
        For i = .PicPage.LBound To .PicPage.UBound
            If .PicPage(i).Tag = "true" Then
                If lngMin = -1 Then lngMin = i
                If i > lngMax Then lngMax = i
                If Not blnCur Then
                    lngHeight = lngHeight + .PicPage(i).Height
                    If lngHeight > 500 + Abs(.picMain.Top) Then
                        lngCurPage = i
                        blnCur = True
                    End If
                End If
            End If
        Next
        If lngPage = -1 Then
            If blnForWord Then
                For i = .PicPage.LBound To .PicPage.UBound
                    If .PicPage(i).Tag = "true" Then
                        If i > lngCurPage Then
                            lngPage = i
                            Exit For
                        End If
                    End If
                Next
                If lngPage = -1 Then lngPage = lngMax
            Else
                For i = .PicPage.UBound To .PicPage.LBound Step -1
                    If .PicPage(i).Tag = "true" Then
                        If i < lngCurPage Then
                            lngPage = i
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        
        If lngPage < lngMin Then
            lngPage = lngMin
        ElseIf lngPage > lngMax Then
            lngPage = lngMax
        End If
       
        lngHeight = 0
        For i = .PicPage.LBound To lngPage - 1
            If .PicPage(i).Tag = "true" Then
                lngHeight = lngHeight + .PicPage(i).Height
            End If
        Next
    
         i = Abs((-500 - lngHeight - .PicPage(0).ScaleTop) / ((.picMain.Height + 1100 - .ScaleHeight)) * 1000)
        .vsbMain.Value = IIf(i > 1000, 1000, i)
        If Not objTmp Is Nothing Then
             zlControl.ControlSetFocus objTmp
             Exit Sub
        End If
        
        If blnLocation Then                     '是否定位到控件，滚动鼠标的时候不定位
            Select Case lngPage
                Case PIC_住院首页
                    If gclsPros.FuncType = f医生首页 Then
                        zlControl.ControlSetFocus .cboBaseInfo(BCC_付款方式)
                    ElseIf gclsPros.FuncType = f病案首页 Then
                        If Not .txtSpecificInfo(SLC_住院号).Locked Then
                            zlControl.ControlSetFocus .txtSpecificInfo(SLC_住院号)
                        Else
                            zlControl.ControlSetFocus .cboBaseInfo(BCC_付款方式)
                        End If
                    End If
                Case PIC_基本信息
                        If gclsPros.FuncType = f医生首页 Then
                            zlControl.ControlSetFocus .cboBaseInfo(BCC_国籍)
                        ElseIf gclsPros.FuncType = f病案首页 Then
                            If .txtInfo(GC_姓名).Locked Then
                                zlControl.ControlSetFocus .cboBaseInfo(BCC_国籍)
                            Else
                                zlControl.ControlSetFocus .txtInfo(GC_姓名)
                            End If
                        End If
                Case PIC_西医诊断
                    Call LocateVSFRowCol(.vsDiagXY, 1, .vsDiagXY.Rows - 1, DI_诊断编码, DI_Del, 1, DI_诊断描述)
                    zlControl.ControlSetFocus .vsDiagXY
                Case PIC_西医诊断情况
                    zlControl.ControlSetFocus .cboBaseInfo(BCC_入院情况)
                Case PIC_中医诊断
                    Call LocateVSFRowCol(.vsDiagZY, 1, .vsDiagZY.Rows - 1, DI_诊断编码, DI_Del, 1, DI_诊断描述)
                    zlControl.ControlSetFocus .vsDiagZY
                Case PIC_中医诊断情况
                    zlControl.ControlSetFocus .cboBaseInfo(BCC_门诊与出院ZY)
                Case PIC_药物过敏
                    Call LocateVSFRowCol(.vsAller, 1, .vsAller.Rows - 1, AI_过敏药物, AI_过敏时间, 1, AI_过敏药物)
                    zlControl.ControlSetFocus .vsAller
                Case PIC_输血信息
                    zlControl.ControlSetFocus .cboBaseInfo(BCC_血型)
                Case PIC_签名信息
                    zlControl.ControlSetFocus .cboManInfo(MC_科主任)
                Case PIC_手术记录
                    Call LocateVSFRowCol(.vsOPS, 1, .vsOPS.Rows - 1, PI_手术日期, PI_切口愈合, 1, PI_手术日期)
                    zlControl.ControlSetFocus .vsOPS
                Case PIC_住院费用
                    zlControl.ControlSetFocus .chkFeeEdit
                Case PIC_住院情况
                    zlControl.ControlSetFocus .cboBaseInfo(BCC_病例分型)
                Case PIC_化疗信息
                    zlControl.ControlSetFocus .vsChemoth
                Case PIC_放疗记录
                    zlControl.ControlSetFocus .vsRadioth
                Case PIC_抗精神病
                    zlControl.ControlSetFocus .vsSpirit
                Case PIC_抗菌药物
                    zlControl.ControlSetFocus .vsKSS
                Case PIC_重症监护
                    zlControl.ControlSetFocus .vsFlxAddICU
                Case PIC_病案附加
                    zlControl.ControlSetFocus .vsfMain
                Case PIC_附页1
                    If gclsPros.MedPageSandard = ST_卫生部标准 Then
                         zlControl.ControlSetFocus .lstInfection
                    ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                        zlControl.ControlSetFocus .vsInfect
                    ElseIf gclsPros.MedPageSandard = ST_云南省标准 Then
                        zlControl.ControlSetFocus .txtInfo(GC_重症监护室名称)
                    ElseIf gclsPros.MedPageSandard = ST_湖南省标准 Then
                        If .optInput(OP_ICU无).Value Then
                            zlControl.ControlSetFocus .optInput(OP_ICU无)
                        Else
                            zlControl.ControlSetFocus .optInput(OP_ICU有)
                        End If
                    End If
                Case PIC_附页2
                    If gIntPic + 1 <> PIC_附页2 Then
                        If gclsPros.MedPageSandard = ST_四川省标准 Then
                            zlControl.ControlSetFocus .chkInfo(CHK_进入路径)
                        End If
                    Else
                        If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                            zlControl.ControlSetFocus gfrmMecCol(lngPage - gIntPic)
                        End If
                    End If
                Case Else
                    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                            zlControl.ControlSetFocus gfrmMecCol(lngPage - gIntPic)
                    End If
            End Select
        End If
    End With
End Sub

Public Sub PrintInMedRec(ByVal mopType As MedRec_Operate, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, _
                        Optional ByVal intPage As Integer, Optional ByRef objReport As Object, Optional ByRef objForm As Object)
'功能：首页打印，预览
'参数：intType=2（打印），=1（预览）0=设置
'     lng科室ID-病人科室
'     intPage=1-4打印的页数（格式）=5打印正面+附页1，=6打印反面+附页2
    Dim strName As String
    Dim lngPage As Long
    Dim objReportTmp As clsReport
    Dim objFormTmp As Object
    Dim bln中医 As Boolean

    If lng病人ID <> 0 Then
        If gobjReport Is Nothing Then Set gobjReport = New clsReport
        Set objReportTmp = IIf(objReport Is Nothing, gobjReport, objReport)
        Set objFormTmp = IIf(objForm Is Nothing, gclsPros.CurrentForm, objForm)
        bln中医 = sys.DeptHaveProperty(lng科室ID, "中医科")
        '病案系统打印封面
        If gclsPros.SysNo \ 100 = 3 Then
            strName = "ZL3_BILL_200"
            intPage = 0
            mopType = MOP_打印
        Else
            Select Case gclsPros.MedPageSandard
                Case ST_卫生部标准 '卫生部标准
                    If bln中医 Then
                        strName = "ZL1_INSIDE_1261_4"
                    Else
                        strName = "ZL1_INSIDE_1261_1"
                    End If
                Case ST_四川省标准    '四川省标准
                    If bln中医 Then
                        strName = "ZL1_INSIDE_1261_6"
                    Else
                        strName = "ZL1_INSIDE_1261_5"
                    End If
                Case ST_云南省标准    '云南省标准
                    If bln中医 Then
                        strName = "ZL1_INSIDE_1261_8"
                    Else
                        strName = "ZL1_INSIDE_1261_7"
                    End If
                Case ST_湖南省标准    '湖南省标准
                    If bln中医 Then
                        strName = "ZL1_INSIDE_1261_10"
                    Else
                        strName = "ZL1_INSIDE_1261_9"
                    End If
            End Select

            If GetSetting("ZLSOFT", "私有模块\" & UserInfo.DBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 0) = 0 And intPage = 0 Then
                Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.DBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 1)
            End If
        End If
        
        If mopType = MOP_设置 Then
            Call ReportPrintSet(gcnOracle, gclsPros.SysNo, strName, objFormTmp)
        Else
            If intPage = 5 Then
                lngPage = 1
            ElseIf intPage = 6 Then
                lngPage = 2
            Else
                lngPage = intPage
            End If
            Call objReportTmp.ReportOpen(gcnOracle, gclsPros.SysNo, strName, objFormTmp, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), mopType)
            If intPage > 4 Then
                Call objReportTmp.ReportOpen(gcnOracle, gclsPros.SysNo, strName, objFormTmp, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), mopType)
            End If
        End If
    End If
End Sub

Public Sub SetKSSSerial()
'功能：设置抗菌药表格的行序号
    Dim i As Long

    With gclsPros.CurrentForm.vsKSS
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, KI_序号) = i
        Next
    End With
End Sub

Public Sub SetPatiAddress(ByVal lngIndex As Long, ByVal strInfoName As String, ByVal strInfoValue As String, Optional ByVal blnDefault As Boolean)
'功能：设置某个地址相关控件的值
'参数:lngIndex=地址相关控件Index
'     strInfoName=信息名
'     strInfoVale信息值
'     blnDefault=是否是设置默认值
    Dim rsTmp As ADODB.Recordset
    Dim blnHavePadr As Boolean '是否有病人地址控件
    Dim strTmp As String

    On Error GoTo errH
    With gclsPros.CurrentForm
        On Error Resume Next
        Err.Clear: strTmp = .padrInfo(lngIndex).Value
        blnHavePadr = Err.Number = 0: Err.Clear
        On Error GoTo errH
        If gclsPros.IsStructAdress And blnHavePadr Then
            Set rsTmp = GetStrucAddress(gclsPros.病人ID, gclsPros.主页ID, strInfoName)
            If rsTmp.RecordCount > 0 Then
                Call .padrInfo(lngIndex).LoadStructAdress(rsTmp!省 & "", rsTmp!市 & "", rsTmp!县 & "", rsTmp!乡镇 & "", rsTmp!其他 & "")
                If blnDefault Then .padrInfo(lngIndex).Tag = rsTmp!省 & "" & rsTmp!市 & "" & rsTmp!县 & "" & rsTmp!乡镇 & "" & rsTmp!其他
            Else
                .padrInfo(lngIndex).Value = strInfoValue
                If blnDefault Then .padrInfo(lngIndex).Tag = strInfoValue
            End If
        Else
            .txtAdressInfo(lngIndex).Text = strInfoValue
            If blnDefault Then .txtAdressInfo(lngIndex).Tag = strInfoValue
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ValidateAge(ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox, Optional ByVal bytIndex As Byte = 0) As Boolean
'功能：检查年龄输入值的有效性
'返回：
'61454:刘鹏飞,2013-05-14,添加对婴幼儿年龄的校对
'bytIndex 0 病人年龄 1 婴幼儿年龄
    If Not IsNumeric(txt年龄.Text) Then ValidateAge = True: Exit Function

    If bytIndex = 0 Then
        Select Case cbo年龄单位.Text
            Case "岁"
                If Val(txt年龄.Text) > 200 Then
                    MsgBox "年龄值超过了最大限制200岁，请检查输入是否正确。", vbInformation, gstrSysName
                    If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "月"
                If Val(txt年龄.Text) > 2400 Then
                    MsgBox "年龄值超过了最大限制2400月，请检查输入是否正确。", vbInformation, gstrSysName
                    If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "天"
                If Val(txt年龄.Text) > 73000 Then
                    MsgBox "年龄值超过最大限制73000天，请检查输入是否正确。", vbInformation, gstrSysName
                    If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "小时" '不能大于30天即720小时
                If Val(txt年龄.Text) > 720 Then
                    MsgBox "年龄值超过了最大限制720小时，请使用合适的年龄单位。", vbInformation, gstrSysName
                    If cbo年龄单位.Enabled And cbo年龄单位.Visible Then cbo年龄单位.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "分钟" '不能大于24小时即1440分钟
                If Val(txt年龄.Text) > 1440 Then
                    MsgBox "年龄值超过了最大限制1440分钟，请使用合适的年龄单位。", vbInformation, gstrSysName
                    If cbo年龄单位.Enabled And cbo年龄单位.Visible Then cbo年龄单位.SetFocus
                    ValidateAge = False: Exit Function
                End If
        End Select
    Else
        Select Case cbo年龄单位.Text
            Case "月"
                If Val(txt年龄.Text) > 12 Then
                    MsgBox "婴儿年龄值超过了最大限制12月，请检查输入是否正确。", vbInformation, gstrSysName
                    If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "天"
                If Val(txt年龄.Text) > 365 Then
                    MsgBox "婴儿年龄值超过了最大限制365天，请检查输入是否正确。", vbInformation, gstrSysName
                    If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "小时"
                If Val(txt年龄.Text) > 720 Then
                    MsgBox "婴儿年龄值超过了最大限制720小时，请使用合适的年龄单位。", vbInformation, gstrSysName
                    If cbo年龄单位.Enabled And cbo年龄单位.Visible Then cbo年龄单位.SetFocus
                    ValidateAge = False: Exit Function
                End If
            Case "分钟"
                If Val(txt年龄.Text) > 1440 Then
                    MsgBox "婴儿年龄值超过了最大限制1440分钟，请使用合适的年龄单位。", vbInformation, gstrSysName
                    If cbo年龄单位.Enabled And cbo年龄单位.Visible Then cbo年龄单位.SetFocus
                    ValidateAge = False: Exit Function
                End If
        End Select
    End If
    ValidateAge = True
End Function

Public Function GetKSSUseStage(ByVal DateUseBegin As Date, ByVal DateUseEnd As Date, ByVal DateSs As Date) As String
'功能：获得抗生素使用阶段
'参数：DateUseBegin 使用时间,DateUseEnd -结束时间  DateSs-手术时间,strTime 上一次的使用阶段
    Dim strTimeTmp As String

    '如果没有手术，则返回空
    If DateSs <> CDate(0) Then
        If DateUseBegin < DateSs And DateUseEnd < DateSs Then
            strTimeTmp = "术前"
        ElseIf DateUseBegin > DateSs And DateUseEnd > DateSs Then
            strTimeTmp = "术后"
        ElseIf DateUseBegin = DateSs And DateUseEnd = DateSs Then
            strTimeTmp = "术中"
        End If
        If strTimeTmp = "" Then strTimeTmp = "围手术期"
    End If
    GetKSSUseStage = strTimeTmp
End Function

Public Function GetKSSUseDay(ByVal AdviceID As Long, ByVal lng药品ID As Long, ByVal str执行时间方案 As String, ByVal Date开始执行时间 As Date, _
            ByVal Date结束时间 As Date, ByVal lng频率次数 As Long, ByVal lng频率间隔 As Long, ByVal str间隔单位 As String, ByVal str用药目的 As String, _
            ByRef rsTime As ADODB.Recordset) As Long
'功能：获取抗生素的使用天数
    Dim blnNew As Boolean
    Dim strPause As String
    Dim j As Long
    Dim StrDecTime As String, arrDecTime As Variant
    Dim DateStart As String
    Dim strTmp As String

    '记录集没实例化，则重新定义
    blnNew = rsTime Is Nothing
    If Not blnNew Then
        blnNew = rsTime.Fields.Count <> 3
        If Not blnNew Then blnNew = rsTime.Fields(0).Name <> "收费时间" Or rsTime.Fields(1).Name <> "药品ID" Or rsTime.Fields(2).Name <> "用药目的"
    End If

    If blnNew Then
        Set rsTime = New ADODB.Recordset
        rsTime.Fields.Append "收费时间", adVarChar, 10
        rsTime.Fields.Append "药品ID", adBigInt
        rsTime.Fields.Append "用药目的", adVarChar, 100
        rsTime.CursorLocation = adUseClient
        rsTime.LockType = adLockOptimistic
        rsTime.CursorType = adOpenStatic
        rsTime.Open
    End If

    strPause = GetAdvicePause(AdviceID)

    If str执行时间方案 <> "" Then
        StrDecTime = Calc段内分解时间(Date开始执行时间, Date结束时间, strPause, str执行时间方案, lng频率次数, lng频率间隔, str间隔单位)
        arrDecTime = Split(StrDecTime, ",")
        For j = 0 To UBound(arrDecTime)
            strTmp = Format(arrDecTime(j), "yyyy-MM-dd")
            rsTime.Filter = "收费时间='" & strTmp & "' And " & "药品id=" & lng药品ID & " And 用药目的='" & str用药目的 & "'"
            If rsTime.EOF Then
                rsTime.AddNew
                rsTime!收费时间 = strTmp
                rsTime!药品ID = lng药品ID
                rsTime!用药目的 = str用药目的
                rsTime.Update
            End If
        Next
    Else
        DateStart = CDate(Format(Date开始执行时间 & "", "yyyy-MM-dd"))
        Do While DateStart <= CDate(Format(Date结束时间 & "", "yyyy-MM-dd"))
            rsTime.Filter = "收费时间='" & Format(CStr(DateStart), "yyyy-MM-dd") & "' And " & "药品id=" & lng药品ID & " And 用药目的='" & str用药目的 & "'"
            If rsTime.EOF Then
                rsTime.AddNew
                rsTime!收费时间 = Format(CStr(DateStart), "yyyy-MM-dd")
                rsTime!药品ID = lng药品ID
                rsTime!用药目的 = str用药目的
                rsTime.Update
            End If
            DateStart = CDate(DateStart) + 1
        Loop
    End If
    rsTime.Filter = "药品id=" & lng药品ID & " And 用药目的='" & str用药目的 & "'"
    GetKSSUseDay = rsTime.RecordCount
End Function

Public Function ClearPageContent() As Boolean
'功能：清除表格的内容
    Dim ctlTmp As Control
    Dim i As Long, j As Long
    Dim rsTemp As New ADODB.Recordset
    '控件变量定义，方便属性查看
    Dim vsTmp As VSFlexGrid, txtTmp As TextBox, paTmp As PatiAddress
    Dim chkTmp As CheckBox, lstTmp As ListBox, cboTmp As ComboBox
    Dim lvwTmp As ListView, mskTmp As MaskEdBox, optTmp As OptionButton
    Dim vsbTmp As VScrollBar, hsbTmp As HScrollBar
    Dim arrTmp As Variant
    On Error GoTo errH
    If gclsPros.FuncType = f电子病案 Then
        gblnCheck = True
        For Each ctlTmp In gclsPros.CurrentForm.Controls
            Select Case TypeName(ctlTmp)
                Case "TextBox" '约120-140个
                    Set txtTmp = ctlTmp
                    If txtTmp.Index = GCA_死亡患者尸检 Then
                        txtTmp.Tag = ""
                        txtTmp.Text = ""
                    Else
                        txtTmp.Text = txtTmp.Tag '可能有默认值
                    End If
                Case "CheckBox" '复选控件大约在15-30个之间
                    Set chkTmp = ctlTmp
                    chkTmp.Value = 0
                Case "VSFlexGrid" '表格清空
                    Set vsTmp = ctlTmp
                    vsTmp.Clear
                Case "ListBox" '有3个
                    Set lstTmp = ctlTmp
                    lstTmp.Clear
            End Select
        Next
        gblnCheck = False
    Else
        For Each ctlTmp In gclsPros.CurrentForm.Controls
            'case语句排列的先后顺序：将最多的放在最前面，表格类控件较少，放后面
            Select Case TypeName(ctlTmp)
                Case "Label", "Frame"
                    'lbl不做处理，留在这里占位置，因为lbl控件比较多，所以放第一位
                Case "TextBox" '约50-60个
                    Set txtTmp = ctlTmp
                    txtTmp.Text = ""
                    '恢复默认值
                    If txtTmp.Name = "txtAdressInfo" Then
                        txtTmp.Text = txtTmp.Tag
                    Else
                         txtTmp.Tag = ""
                    End If
                Case "ComboBox" '约40-50个
                    Set cboTmp = ctlTmp
                    If cboTmp.Style = 0 Then
                        cboTmp.Text = "" '可以输入的下拉列表清空输入内容
                        cboTmp.Tag = ""
                    End If
                    If cboTmp.Tag <> "" Then '恢复默认值
                        cboTmp.ListIndex = Val(cboTmp.Tag)
                    Else
                        cboTmp.ListIndex = -1
                    End If
                    '清除手工添加的人员
                    If gclsPros.FuncType = f病案首页 And cboTmp.Name = "cboManInfo" And cboTmp.ListCount > 0 Then
                        For i = cboTmp.ListCount - 1 To 0 Step -1
                            If cboTmp.ItemData(i) = -999 Then
                                cboTmp.RemoveItem i
                            Else '因为是从后面添加，因此不是-999则退出循环
                                Exit For
                            End If
                        Next
                    End If
                Case "CheckBox" '复选控件大约在10-20个之间
                    Set chkTmp = ctlTmp
                    '恢复默认值
                    chkTmp.Value = 0
'                    If chkTmp.Index = CHK_是否确诊 Then
'                        chkTmp.Value = 1
'                    Else
'                        chkTmp.Value = 0
'                    End If
                Case "VSFlexGrid" '表格清空
                    Set vsTmp = ctlTmp
                    '固定行不等于0的只有诊断表格，转科记录，特殊检查，这些只需要清除单元格内容即可
                    If vsTmp.FixedCols <> 0 Then
                        '不清空诊断类型列的数据
                        '删除诊断类别为空的行，清空诊断类别不为空的行的单元格数据
                        If vsTmp.Name = "vsDiagXY" Or vsTmp.Name = "vsDiagZY" Then
                            '清空数据
                            vsTmp.Cell(flexcpData, vsTmp.FixedRows, vsTmp.FixedCols, vsTmp.Rows - 1, vsTmp.Cols - 1) = Empty
                            vsTmp.Cell(flexcpText, vsTmp.FixedRows, vsTmp.FixedCols, vsTmp.Rows - 1, DI_诊断分类 - 1) = ""
                            vsTmp.Cell(flexcpText, vsTmp.FixedRows, DI_诊断分类 + 1, vsTmp.Rows - 1, vsTmp.Cols - 1) = ""
                            i = vsTmp.FixedRows: j = vsTmp.Rows - 1
                            Do While i <= j
                               If vsTmp.TextMatrix(i, DI_诊断类型) = "" Then
                                    vsTmp.RemoveItem i
                                     j = vsTmp.Rows - 1
                                Else
                                    vsTmp.RowData(i) = 0
                                    i = i + 1
                                End If
                            Loop
                            '设置控件初始的可用状态
                            Call SetDiagReletedInfo(vsTmp)
                            Call ChangeOutInfo
                        '除诊断外具有固定列的表为纵向表，这些纵向表只需清空数据即可
                        Else
                            vsTmp.Cell(flexcpData, vsTmp.FixedRows, vsTmp.FixedCols, vsTmp.Rows - 1, vsTmp.Cols - 1) = Empty
                            vsTmp.Cell(flexcpText, vsTmp.FixedRows, vsTmp.FixedCols, vsTmp.Rows - 1, vsTmp.Cols - 1) = ""
                            For i = vsTmp.FixedRows To vsTmp.Rows - 1
                                vsTmp.RowData(i) = 0
                            Next
                        End If
                    '固定列等于0，固定行不等于0的表格有，过敏，费用，化疗，放疗，抗菌药，重症监护，病案附加项目，精神药
                    '这些表格，只需要删除所有的行，再新增一行即可，重症监护不能这样做，重症监护需要清空内容列
                    ElseIf vsTmp.FixedRows <> 0 Then
                        If vsTmp.Name = "vsfMain" Then
                            For i = vsTmp.FixedRows To vsTmp.Rows - 1
                                For j = 0 To vsTmp.Cols - 1 Step 3
                                    If vsTmp.TextMatrix(i, j + 2) = "是否" Then
                                        vsTmp.Cell(flexcpChecked, i, j + 1) = 2
                                    Else
                                        vsTmp.TextMatrix(i, j + 1) = ""
                                    End If
                                Next
                            Next
                        Else
                            '其他的横向表，删掉所有的行，再新增一行
                            vsTmp.Rows = vsTmp.FixedRows
                            vsTmp.Rows = vsTmp.Rows + 1
                        End If
                    End If
                Case "PatiAddress" '地址控件比较少只有4个放在后面
                    Set paTmp = ctlTmp
                    '恢复默认值
                    paTmp.Value = paTmp.Tag
                Case "ListBox" '有3个
                    Set lstTmp = ctlTmp
                    For i = 0 To lstTmp.ListCount - 1
                        lstTmp.Selected(i) = False
                    Next
                Case "ListView" 'lvwFees
                    Set lvwTmp = ctlTmp
                    For i = 0 To lvwTmp.ListItems.Count - 1
                        lvwTmp.ListItems(i).Checked = False
                    Next
                Case "MaskEdBox"
                    Set mskTmp = ctlTmp
                    If mskTmp.Index = DC_收回日期 And gclsPros.OpenMode <> EM_查阅 Then
                    '病案首页编目不作处理，使用相同的收回日期
                    Else
                        mskTmp.Text = Replace(mskTmp.Tag, "#", "_")
                    End If
                Case "OptionButton"
                    Set optTmp = ctlTmp
                    optTmp.Value = (optTmp.Tag = "1")
                Case "VScrollBar"
                    Set vsbTmp = ctlTmp
                    vsbTmp.Value = 0
                Case "HScrollBar"
                    Set hsbTmp = ctlTmp
                    hsbTmp.Value = 0
            End Select
        Next
        Call SetFaceInit(True, True)
    End If
    gclsPros.IsOK = False
    If Not gclsPros.PatiInfo Is Nothing Then Set gclsPros.PatiInfo = zlDatabase.CopyNewRec(gclsPros.PatiInfo, True)
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetFaceInit(Optional ByVal blnUnlock As Boolean, Optional ByVal blnReSetDefault As Boolean)
'功能：将界面回复到初始状态
' 功能：blnUnlock=是否是签名解锁
'          blnReSetDefault=是否重新设置默认值
    Dim objControl As Object
    Dim i As Long
    Dim LngRow As Long
    Dim strTmp As String, blnTmp As Boolean
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim datCur  As Date
'    If blnReSetDefault Then Stop
    On Error GoTo errH
    With gclsPros.CurrentForm
        '解锁所有控件
        If blnUnlock Then
            For Each objControl In .Controls
                If InStr(",Timer,CommonDialog,Menu,Label,Subclass,", "," & TypeName(objControl) & ",") = 0 Then
                    If Not objControl.Container Is Nothing Then
                        If TypeName(objControl.Container) = "PictureBox" Or TypeName(objControl.Container) = "Frame" Then
                            If Not (objControl.Name = "cmdSign") Then
                                Call SetCtrlLocked(objControl, False)
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If gclsPros.FuncType <> f病案首页 Then
            '设置特定控件的状态
            Call SetCtrlLocked(.txtInfo(GC_姓名), True)
            Call SetCtrlLocked(.cboBaseInfo(BCC_性别), True)
            Call SetCtrlLocked(.txtSpecificInfo(SLC_年龄), True)
            Call SetCtrlLocked(.mskDateInfo(DC_出生日期), True)
        ElseIf gclsPros.FuncType = f病案首页 Then
            Call SetCtrlLocked(.txtInfo(GC_病案号), True)
        End If
        If gclsPros.PatiType = PF_住院 Then
            strSql = "Select * From 病人变动记录 Where 病人ID=[1] And 主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病人变动记录", gclsPros.病人ID, gclsPros.主页ID)
            Call SetCtrlLocked(.txtInfo(BCC_健康卡号), True)
            
            Call SetCtrlLocked(.mskDateInfo(DC_发病时间), Not IsDate(.mskDateInfo(DC_发病日期).Text))
            If gclsPros.FuncType <> f病案首页 Then
                Call SetCtrlLocked(.mskDateInfo(DC_入院时间), True)
                Call SetCtrlLocked(.mskDateInfo(DC_出院时间), True)
                Call SetCtrlLocked(.txtInfo(GC_入院科室), True)
                Call SetCtrlLocked(.txtInfo(GC_出院科室), True)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_住院号), True)
            Else
                '与ZLHIS系统联机参数勾选,不能联机修改的项目，字体为蓝色
                Call SetCtrlLocked(.mskDateInfo(DC_入院时间), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.mskDateInfo(DC_出院时间), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.txtInfo(GC_入院科室), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.txtInfo(GC_出院科室), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.cmdDateInfo(DC_入院时间), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.cmdDateInfo(DC_出院时间), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.cmdInfo(GC_入院科室), IIf(rsTmp.RecordCount > 0, True, False), , True)
                Call SetCtrlLocked(.cmdInfo(GC_出院科室), IIf(rsTmp.RecordCount > 0, True, False), , True)
            End If
            Call SetCtrlLocked(.txtSpecificInfo(SLC_住院天数), True)
            Call SetCtrlLocked(.txtSpecificInfo(SLC_入院次数), True)
            Call SetDiagMatchInfo(BCC_放射与病理, True)
            Call SetDiagMatchInfo(BCC_临床与病理, True)
            Call SetDiagMatchInfo(BCC_临床与尸检, True)
            Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_病理号), False)
            Call SetCtrlLocked(.chkInfo(CHK_病原学检查), .vsDiagXY.TextMatrix(FindDiagRow(DT_院内感染), DI_诊断描述) = "")
            LngRow = FindDiagRow(DT_出院诊断XY)
            strTmp = UCase(Trim(.vsDiagXY.TextMatrix(LngRow, DI_诊断编码)))
            blnTmp = strTmp Like "C*" Or strTmp Like "D0*" Or strTmp Like "D32.*" Or strTmp Like "D33.*"
            Call SetCtrlLocked(.cboBaseInfo(BCC_分化程度), Not blnTmp, Not blnTmp)
            Call SetCtrlLocked(.cboBaseInfo(BCC_最高诊断依据), Not blnTmp, Not blnTmp)

            strTmp = zlStr.NeedName(.cboBaseInfo(BCC_出院方式).Text)
            Call SetCtrlLocked(.mskDateInfo(DC_死亡时间), strTmp <> "死亡")
            Call SetCtrlLocked(.txtInfo(GC_死亡原因), strTmp <> "死亡")
            Call SetCtrlLocked(.cmdInfo(GC_死亡原因), strTmp <> "死亡")
            Call SetCtrlLocked(.cboBaseInfo(BCC_死亡患者尸检), strTmp <> "死亡")
            If .cboBaseInfo(BCC_死亡患者尸检).ListIndex = -1 Then .cboBaseInfo(BCC_死亡患者尸检).ListIndex = 0
            Call SetCtrlLocked(.chkInfo(CHK_随诊), strTmp = "死亡")
            Call SetCtrlLocked(.cboBaseInfo(BCC_死亡期间), strTmp <> "死亡")
            If gclsPros.FuncType = f病案首页 Then
                .cmdDeliceryInfo.Visible = False
                .cmdDeliceryInfo.Enabled = False
                .cmdDeliceryInfo.Tag = ""
                For i = LngRow To .vsDiagXY.Rows - 1
                    If Val(.vsDiagXY.TextMatrix(i, DI_诊断分类)) = DT_出院诊断XY Then
                        If .vsDiagXY.TextMatrix(i, DI_分娩信息) = "1" Then
                            .cmdDeliceryInfo.Visible = True
                            .cmdDeliceryInfo.Enabled = True
                            .cmdDeliceryInfo.Tag = "1"
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If

            Call chkInfoClick(CHK_病原学检查)
            Call chkInfoClick(CHK_随诊)
            Call chkInfoClick(CHK_是否确诊)

            Call SetCtrlLocked(.txtInfo(GC_31天内再住院), .optInput(OP_再住院无).Value)
            Call SetCtrlLocked(.txtInfo(GC_抢救病因), Val(.txtSpecificInfo(SLC_抢救次数).Text) = 0)
            Call SetCtrlLocked(.cmdInfo(GC_抢救病因), Val(.txtSpecificInfo(SLC_抢救次数).Text) = 0)
            Call SetCtrlLocked(.txtSpecificInfo(SLC_成功次数), Val(.txtSpecificInfo(SLC_抢救次数).Text) = 0)
            For i = 0 To .lstAdvEvent.ListCount - 1
                If .lstAdvEvent.List(i) = "压疮" Then
                    Call SetCtrlLocked(.cboBaseInfo(BCC_压疮发生期间), Not .lstAdvEvent.Selected(i))
                    Call SetCtrlLocked(.cboBaseInfo(BCC_压疮分期), Not .lstAdvEvent.Selected(i))
                ElseIf .lstAdvEvent.List(i) = "医院内跌倒/坠床" Then
                    Call SetCtrlLocked(.cboBaseInfo(BCC_跌倒或坠床伤害), Not .lstAdvEvent.Selected(i))
                    Call SetCtrlLocked(.cboBaseInfo(BCC_跌倒或坠床原因), Not .lstAdvEvent.Selected(i))
                End If
            Next

            strTmp = zlStr.NeedName(.cboBaseInfo(BCC_出院方式).Text)
            blnTmp = Not (strTmp Like "*转院*" Or strTmp Like "*转社区*")
            Call SetCtrlLocked(.txtInfo(GC_转入医疗机构), blnTmp)
            Call SetCtrlLocked(.cmdInfo(GC_转入医疗机构), blnTmp)

            strTmp = zlStr.NeedName(.cboBaseInfo(BCC_入院途径).Text)
            blnTmp = Not (strTmp Like "*转入*" And Not strTmp Like "*非转入*")
            Call SetCtrlLocked(.txtInfo(GC_入院转入), blnTmp)
            Call SetCtrlLocked(.cmdInfo(GC_入院转入), blnTmp)

            If gclsPros.MedPageSandard = ST_四川省标准 Then
                blnTmp = zlStr.NeedName(.cboBaseInfo(BCC_输液反应).Text) <> "有"
                Call SetCtrlLocked(.txtInfo(GC_引发药物), blnTmp)
                Call SetCtrlLocked(.txtInfo(GC_临床表现), blnTmp)
                Call chkInfoClick(CHK_会诊情况)
            ElseIf gclsPros.MedPageSandard = ST_云南省标准 Then
                blnTmp = .txtInfo(GC_重症监护室名称).Text = ""
                Call SetCtrlLocked(.chkInfo(CHK_人工气道脱出), blnTmp)
                Call SetCtrlLocked(.chkInfo(CHK_重返重症医学科), blnTmp)
                Call SetCtrlLocked(.cboBaseInfo(BCC_重返间隔时间), blnTmp)
                Call chkInfoClick(CHK_住院物理约束)
            ElseIf gclsPros.MedPageSandard = ST_湖南省标准 Then
                Call SetCtrlLocked(.txtSpecificInfo(SLC_重症监护天), .optInput(OP_ICU无).Value)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_重症监护小时), .optInput(OP_ICU无).Value)
            End If
            If gclsPros.MedPageSandard = ST_云南省标准 Or gclsPros.MedPageSandard = ST_四川省标准 Then
                Call chkInfoClick(CHK_进入路径)
                Call chkInfoClick(CHK_变异)
                Call chkInfoClick(CHK_完成路径)
            End If
            If gclsPros.FuncType = f病案首页 Then
                '病案首页快速编辑设置
                .lblNote.Visible = Not gclsPros.EditUnrecive
                .txtInfo(GC_病案号).TabStop = gclsPros.EditPageNo And gclsPros.OpenMode = EM_新增病案
                .txtInfo(GC_档案号).TabStop = gclsPros.TabFileNo
                .cboBaseInfo(BCC_付款方式).TabStop = gclsPros.TabPayType
                .cboSpecificInfo(SLC_年龄).TabStop = gclsPros.TabAgeUnit
                .cboBaseInfo(BCC_国籍).TabStop = gclsPros.TabNation
                .txtInfo(GC_病案号).Locked = Not gclsPros.EditPageNo
                .chkInfo(CHK_再入院).TabStop = gclsPros.TabReadm
                .txtInfo(GC_X线号).TabStop = gclsPros.TabXRaysNo

                If gclsPros.OpenMode = EM_查阅 Or gclsPros.OpenMode = EM_编辑 Then
                    frmMain.cbsMain.FindControl(, conMenu_Manage_Up, True).Enabled = Get主页IDByCur(gclsPros.主页ID, False) <> 0
                    frmMain.cbsMain.FindControl(, conMenu_Manage_Down, True).Enabled = Get主页IDByCur(gclsPros.主页ID, True) <> 0
                End If

                Call SetCtrlLocked(.txtSpecificInfo(SLC_住院号), gclsPros.OpenMode <> EM_新增病案 And gclsPros.OpenMode <> EM_新增首页)
                Call SetCtrlLocked(.vsFees, True)
                Call SetCtrlLocked(.cboManInfo(MC_编目员), Not gclsPros.Change编码员)
            Else
                Call SetCtrlLocked(.cboBaseInfo(BCC_付款方式), InStr(gclsPros.Privs, "修改医疗付款方式") = 0)
            End If
            '四川版获取上次诊断、病案云南获取上次诊断
            If gclsPros.MedPageSandard = ST_四川省标准 Or gclsPros.MedPageSandard = ST_云南省标准 And gclsPros.FuncType = f病案首页 Then
                .cmdLastDiag.Visible = gclsPros.主页ID > 1
                .lblDiagInfo.Caption = ""
                .lblDiagInfo.Visible = False
            End If
        Else
            '如果填写了发病时间，则下面的发病时间则不允许填写了
            blnTmp = IsDate(.vsDiagXY.TextMatrix(.vsDiagXY.FixedRows, DI_发病时间)) Or IsDate(.vsDiagZY.TextMatrix(.vsDiagZY.FixedRows, DI_发病时间))
            Call SetCtrlLocked(.mskDateInfo(DC_发病日期), blnTmp)
            Call SetCtrlLocked(.mskDateInfo(DC_发病时间), blnTmp)
            Call SetCtrlLocked(.cboBaseInfo(BCC_付款方式), InStr(GetInsidePrivs(p病人信息公共部件), "基本信息调整") = 0)
            '设置单位名称的可输入性
            blnTmp = InStr(gclsPros.Privs, "合约病人登记") = 0 And Not IsNull(gclsPros.PatiInfo!合同单位id)
            Call SetCtrlLocked(.txtAdressInfo(ADRC_单位地址), blnTmp)
        End If
        If .cboBaseInfo(BCC_血型).ListIndex = -1 Then .cboBaseInfo(BCC_血型).ListIndex = 0
        If .cboBaseInfo(BCC_RH).ListIndex = -1 Then .cboBaseInfo(BCC_RH).ListIndex = 0
        '恢复默认值
        If blnReSetDefault Then
            Call SetCboDefault(.cboBaseInfo(BCC_身份证), -1)
            Call SetCboDefault(.cboBaseInfo(BCC_生育状况), 0)
            Call SetCboDefault(.cboSpecificInfo(SLC_年龄), 0)
            If gclsPros.PatiType <> PF_门诊 Then
                Call SetCboDefault(.cboSpecificInfo(SLC_随诊期限), 0)
                Call SetCboDefault(.cboSpecificInfo(SLC_婴幼儿年龄), 0)
                Call SetCboDefault(.cboBaseInfo(BCC_再入院计划天数), 0)
                Call SetCboDefault(.cboBaseInfo(BCC_感染与死亡关系), 0)
                Call SetCboDefault(.cboBaseInfo(BCC_死亡患者尸检), 0)
                If gclsPros.MedPageSandard = ST_湖南省标准 Then
                    Call SetCboDefault(.cboBaseInfo(BCC_临床路径管理), 0)
                    Call SetCboDefault(.cboBaseInfo(BCC_实施DGRS管理), 0)
                    Call SetCboDefault(.cboBaseInfo(BCC_法定传染病), 0)
                    Call SetCboDefault(.cboBaseInfo(BCC_肿瘤分期), 0)
                End If
            End If
            '根据一些字典设置下拉框内容
            Call SetCboDefaultByRec(Array(BCC_付款方式, BCC_性别, BCC_婚姻, BCC_职业, BCC_民族, BCC_国籍, BCC_血型))
            If gclsPros.PatiType <> PF_门诊 Then
                Call SetCboDefaultByRec(Array(BCC_病例分型, BCC_关系, BCC_入院情况, BCC_入院途径, BCC_分化程度, BCC_最高诊断依据, BCC_出院方式))
            Else
                Call SetCboDefaultByRec(Array(BCC_去向, BCC_文化程度))
            End If
            Call SetCboDefaultByRec(Array(BCC_死亡期间))
            If gclsPros.FuncType = f病案首页 Then
                '得到默认出生地
                strSql = "select A.编码,A.名称 from 地区 a where a.缺省标志=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption)
                If rsTmp.RecordCount > 0 Then
                    Call SetPatiAddress(ADRC_出生地点, "出生地点", rsTmp!名称, True)
                    If gclsPros.DefautADD Then
                        Call SetPatiAddress(ADRC_联系人地址, "联系人地址", rsTmp!名称, True)
                        Call SetPatiAddress(ADRC_现住址, "家庭地址", rsTmp!名称, True)
                        .txtSpecificInfo(SLC_家庭邮编).Text = rsTmp!编码 & ""
                    End If
                End If
                '问题:13557
                strSql = "select A.编码,A.名称 from 区域 a where a.缺省标志=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption)
                If rsTmp.RecordCount > 0 Then
                    Call SetPatiAddress(ADRC_病人区域, "区域", rsTmp!名称, True)
                End If
                datCur = zlDatabase.Currentdate
                '设置默认值
                .mskDateInfo(DC_编目日期).Text = Format(datCur, GetFormat(.mskDateInfo(DC_编目日期).Tag))
                If Not IsDate(.mskDateInfo(DC_收回日期).Text) Then
                    .mskDateInfo(DC_收回日期).Text = Format(datCur, GetFormat(.mskDateInfo(DC_收回日期).Tag))
                End If
                .mskDateInfo(DC_出生日期).Text = Format(datCur, GetFormat(.mskDateInfo(DC_出生日期).Tag))
                .mskDateInfo(DC_入院时间).Text = Format(datCur, GetFormat(.mskDateInfo(DC_入院时间).Tag))
                .mskDateInfo(DC_出院时间).Text = Format(datCur, GetFormat(.mskDateInfo(DC_出院时间).Tag))
                .mskDateInfo(DC_质控日期).Text = Format(datCur, GetFormat(.mskDateInfo(DC_质控日期).Tag))

                .txtDateInfo(DC_编目日期).Text = .mskDateInfo(DC_编目日期).Text
                .txtDateInfo(DC_收回日期).Text = .mskDateInfo(DC_收回日期).Text
                .txtDateInfo(DC_出生日期).Text = .mskDateInfo(DC_出生日期).Text
                .txtDateInfo(DC_入院时间).Text = .mskDateInfo(DC_入院时间).Text
                .txtDateInfo(DC_出院时间).Text = .mskDateInfo(DC_出院时间).Text
                .txtDateInfo(DC_质控日期).Text = .mskDateInfo(DC_质控日期).Text

                .cboManInfo(MC_编目员).Text = UserInfo.姓名
                gclsPros.InTime = .mskDateInfo(DC_入院时间).Text
                gclsPros.OutTime = .mskDateInfo(DC_出院时间).Text
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetFaceEditable(ByVal blnSign As Boolean)
'功能：设置界面控件可用性
'参数：blnSign=是否是签名设置界面，True-根据签名状态设置，False-不考虑签名状态
    Dim objControl As Object
    Dim blnEnable As Boolean
    Dim bln基本信息 As Boolean
    Dim lngTabIndex入院 As Long
    Dim strTypeName As String
    Dim lngMainTab As Long
    Dim blnSet As Boolean

    On Error GoTo errH
    With gclsPros.CurrentForm
        '查阅状态，界面控件均不可用
        If gclsPros.OpenMode = EM_查阅 Then
            For Each objControl In .Controls
                If InStr(",Timer,CommonDialog,Menu,Label,Subclass,VScrollBar,HScrollBar,", "," & TypeName(objControl) & ",") = 0 Then
                    If Not objControl.Container Is Nothing Then
                        If TypeName(objControl.Container) = "PictureBox" Or TypeName(objControl.Container) = "Frame" Then
                            If objControl.Name = "cmdDeliceryInfo" Then '分娩查看按钮
                                Call SetCtrlLocked(objControl, objControl.Tag = "")
                            Else
                                Call SetCtrlLocked(objControl, True)
                            End If
                        End If
                    End If
                End If
            Next
        '编辑状态根据具体参数情况设置
        Else
            '具有基本信息的权限的，入院时间之前的控件若没有填写，则可以填写，否则不可编辑（病人胡姓名、性别、年龄、出生日期不可编辑，只有在入出等级以及病人基本信息修改中修改）
            If gclsPros.FuncType = f医生首页 And gclsPros.PatiType = PF_住院 And Not gclsPros.Is护士站 Then
                lngTabIndex入院 = .lblBaseInfo(BCC_入院途径).TabIndex
                bln基本信息 = InStr(";" & gclsPros.Privs & ";", ";首页基本信息;") > 0
            Else
                lngTabIndex入院 = 0
                bln基本信息 = True
            End If
            For Each objControl In .Controls
                strTypeName = TypeName(objControl): blnSet = True: blnEnable = Not blnSign
                '需要判断的情况
                If InStr(",Timer,CommonDialog,Menu,Frame,Label,Line,Subclass,", "," & TypeName(objControl) & ",") = 0 Then
                    If Not objControl.Container Is Nothing Then
                        If TypeName(objControl.Container) = "PictureBox" And InStr(",PicPage,PicMain,", "," & objControl.Name & ",") = 0 Or TypeName(objControl.Container) = "Frame" Then
                            If blnEnable Then
                                '医生与护士分填首页的控制
                                If objControl.Container.Name = "PicAdvEvent" Or objControl.Container.Name = "PicRestrain" Or objControl.Container.Name = "PicCareInfo" Then
                                    blnEnable = Not gclsPros.SeparateEdit Or gclsPros.Is护士站 And gclsPros.SeparateEdit
                                Else
                                    blnEnable = Not gclsPros.SeparateEdit Or Not gclsPros.Is护士站 And gclsPros.SeparateEdit
                                End If
                                '首页基本信息控制
                                If blnEnable Then
                                    If gclsPros.FuncType = f医生首页 And objControl.TabIndex < lngTabIndex入院 Then
                                        blnSet = Not ControlIsLocked(objControl)
                                        blnEnable = bln基本信息 Or Not ControlHaveValue(objControl)
                                    ElseIf gclsPros.PatiType = PF_住院 And blnEnable Then
                                        '具有病人入院或出院科室具有中医科性质，并且参数"中医科室不使用西医病案首页项目"=True。
                                        '病例分型、输液反应、输血反应、输红细胞、输血小板、输血浆、输全血、自体回收、输其他、输血前的9项检查、
                                        'HBsAg、HCV-Ab、HIV-Ab、示教病案、科研病案、随诊、随诊期限、呼吸机使用、研究生医师.设置不可用
                                        If gclsPros.Have中医 And gclsPros.NotUseXYItems Then
                                            Select Case objControl.Name
                                                Case "cboBaseInfo"
                                                    blnEnable = Not (objControl.Index = BCC_病例分型 Or objControl.Index = BCC_输液反应 Or objControl.Index = BCC_输血反应 Or _
                                                                objControl.Index = BCC_输血前9项检查 Or objControl.Index = BCC_HBsAg Or objControl.Index = BCC_HCVAb Or _
                                                                objControl.Index = BCC_HIVAb)
                                                Case "txtSpecificInfo"
                                                    blnEnable = Not (objControl.Index = SLC_输红细胞 Or objControl.Index = SLC_输全血 Or objControl.Index = SLC_输血浆 Or _
                                                                objControl.Index = SLC_自体回收 Or objControl.Index = SLC_输血小板 Or objControl.Index = SLC_呼吸机使用 Or _
                                                                objControl.Index = SLC_随诊期限 Or objControl.Index = SLC_输白蛋白)
                                                Case "cboManInfo"
                                                    blnEnable = objControl.Index <> MC_研究生医师
                                                Case "chkInfo"
                                                    blnEnable = Not (objControl.Index = CHK_科研病案 Or objControl.Index = CHK_示教病案 Or objControl.Index = CHK_随诊)
                                                Case "txtInfo"
                                                    blnEnable = objControl.Index <> GC_输其他
                                                Case "cboSpecificInfo"
                                                    blnEnable = objControl.Index <> SLC_随诊期限
                                            End Select
                                        End If
                                        blnSet = Not ControlIsLocked(objControl)
                                    Else
                                        blnSet = Not ControlIsLocked(objControl)
                                    End If
                                End If
                            End If
                            If objControl.Name = "cmdSign" Or objControl.Name = "cmdUnSign" Then
                                blnSet = gclsPros.Is护士站
                            End If
                            If blnSet Then
                                Call SetCtrlLocked(objControl, Not blnEnable)
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetMedSt(ByVal strCurFormType As String) As MedPage_Standard
'功能：获取首页标准
    Dim strFixed As String, strOther As String
    If strCurFormType = "" Then Exit Function
    strFixed = decode(gclsPros.FuncType, f医生首页, "frmInMedRecEdit", f病案首页, "frmPageMedRecEdit", f电子病案, "frmArchiveInMedRec", "")
    strOther = Replace(strCurFormType, strFixed, "")
    Select Case strOther
        Case ""
            GetMedSt = ST_卫生部标准
        Case "_SC"
            GetMedSt = ST_四川省标准
        Case "_YN"
            GetMedSt = ST_云南省标准
        Case "_HN"
            GetMedSt = ST_湖南省标准
        Case "frmOutMedRecEdit", "frmArchiveOutMedRec"
            GetMedSt = ST_门诊首页
    End Select
End Function

Public Sub SavePatPicture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:lng病人ID - 病人ID
    '74421,刘鹏飞,2014-07-04,读取病人照片信息
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New Recordset
    Dim strFile As String, strSql As String

    On Error GoTo Errhand
    '没有改变图片不保存
    If gclsPros.CurrentForm.picPatient.Tag = gclsPros.PictureFile Then gclsPros.PictureFile = "0": Exit Sub
    gclsPros.PictureFile = ""
    '图片没有被清除，则重新插入图片
    If gclsPros.CurrentForm.picPatient.Tag <> "" Then
        strFile = gclsPros.CurrentForm.picPatient.Tag
        If sys.SaveLob(gclsPros.SysNo, 27, lng病人ID, strFile) = False Then
            MsgBox "保存照片有误,请确认文件是否被删除!", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    gclsPros.PictureFile = gclsPros.CurrentForm.picPatient.Tag
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub VsGriedFocuesMove(ByRef vsBill As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal KeyCode As Integer, _
        Optional lngFiexCol As Long = 0, Optional lngFiexCol1 As Long = -1)
    '------------------------------------------------------------------------------------------------------------
    '功能:按一定规则移动单元格
    '参数:vsBill-表格控件
    '       lngRow-当前行
    '       lngCol-当前列
    '       KeyCode-按键
    '       lngFiexCol-判断是否移到或加入行的固定列
    '       lngFiexCol1-判断是否移到或加入行的固定列(但同时要满足lngFiexCol列)
    '编制:刘兴宏
    '日期:2007/05/18
    '------------------------------------------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim strCurrValue As String, strTmp As String, LngCols As Long, i As Long
    If LngCol = lngFiexCol Then
        strCurrValue = vsBill.EditText
    Else
        strCurrValue = ""
    End If

    With vsBill
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then LngCols = LngCols + 1
        Next
        Select Case LngCol
        Case 0
            If Trim(.TextMatrix(LngRow, lngFiexCol)) = "" And strCurrValue = "" And vsBill.Name <> "vsFlxAddICU" Then
                zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            .Col = LngCol + 1
            GoTo ShowCell:
        Case Else
            If LngCol >= LngCols - 1 Then
                If LngRow < .Rows - 1 Then
                    .Row = LngRow + 1
                    .Col = .FixedCols
                    GoTo ShowCell:
                    Exit Sub
                End If
                If vsBill.Name = "vsFlxAddICU" Then
                    strTmp = Trim(.TextMatrix(LngRow, lngFiexCol))
                    strTmp = Replace(strTmp, ":", "")
                    strTmp = Replace(strTmp, "-", "")
                    strTmp = Replace(strTmp, "_", "")
                    strTmp = Replace(strTmp, " ", "")
                Else
                    strTmp = Trim(.TextMatrix(LngRow, lngFiexCol))
                End If
                If strTmp <> "" Then
                    If lngFiexCol1 > 0 Then
                        If Trim(.TextMatrix(LngRow, lngFiexCol1)) <> "" Then
                            .Rows = .Rows + 1
                            Call ChangeVSFHeight(vsBill, True)
                            .Row = .Rows - 1
                            .Col = .FixedCols
                        End If
                    Else
                        .Rows = .Rows + 1
                        Call ChangeVSFHeight(vsBill, True)
                        .Row = .Rows - 1
                        .Col = .FixedCols
                    End If
                Else
                    zlCommFun.PressKey vbKeyTab
                    Exit Sub
                End If
                GoTo ShowCell:
                Exit Sub
            End If
            .Col = LngCol + 1
         End Select
ShowCell:
        .ShowCell .Row, .Col
    End With
End Sub

Public Function LoadPatiByInNo(ByVal str住院号 As String, Optional ByVal lng主页ID As Long, Optional ByVal str病案号 As String) As Boolean
'根据当前住院号加载病人信息
    Dim blnOut As Boolean '是否外部文件读取
    Dim lng次数 As Long, lng上次次数 As Long, blnNoCheck As Boolean
    Dim blnOrderAdd As Boolean '预留变量:病案是否只能安顺序添加(目前缺省False,以便以后扩展使用)
    Dim lngTemp As Long
    Dim strTemp As String
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim rs病人主页 As ADODB.Recordset
    Dim intSel主页id As Integer
    Dim blnNO主页ID As Boolean
    If str住院号 <> gclsPros.InNo Then
        Call ClearPageContent
    End If
    gclsPros.IsExistPati = False
    gclsPros.Is编目 = False
    If str住院号 = "" Then
        '63725:刘鹏飞,2013-08-06
        If Not gclsPros.EditUnrecive Then Exit Function
        If gclsPros.OnLine Then
            '如果联机且不能新增，则只能从已有的出院病人中得到
            If Not gclsPros.OnLineNew Then Exit Function
            MsgBox "当前正在增加一位在收费系统不存在的病人。", vbInformation, gstrSysName
            '与ZLHIS系统联机参数勾选,不能联机修改的项目，字体为蓝色
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_入院时间), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_出院时间), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_入院科室), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_出院科室), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_入院时间), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_出院时间), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_入院科室), False)
            Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_出院科室), False)
        End If
        gclsPros.OpenMode = EM_新增病案
        '获取新增的病人的各种号码
        Call ValidatePageNos
        If Not ExistInList(gclsPros.InNo, True) Then Exit Function
    Else
        gclsPros.InNo = str住院号
        If Not ExistInList(gclsPros.InNo, True) Then Exit Function
        If Not IsHavePageNos(CT_住院号, False, gclsPros.InNo) Then
            If Not gclsPros.EditUnrecive Then
                MsgBox "该住院号在收费系统中不存在,不能继续。", vbInformation, gstrSysName
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                Exit Function
            End If
            If gclsPros.OnLine Then
                '如果联机且不能新增，则只能从已有的出院病人中得到
                If Not gclsPros.OnLineNew Then
                    MsgBox "该住院号在收费系统中不存,不能继续。", vbInformation, gstrSysName
                    zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                    Exit Function
                Else
                    If Not gclsPros.NewInNo Then
                        '病人信息中是否存在该住院号
                        If IsHavePageNos(CT_住院号ex, False, gclsPros.InNo) Then
                            If MsgBox("住院号为" & gclsPros.InNo & "的病人在病案首页不存在信息，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                                Exit Function
                            End If
                        Else
                            If MsgBox("住院号为" & str住院号 & "的病人在系统中不存在任何信息，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                                Exit Function
                            End If
                        End If
                        Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_入院时间), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_出院时间), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_入院科室), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_出院科室), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_入院时间), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_出院时间), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_入院科室), False)
                        Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_出院科室), False)
                    End If
                End If
            End If
            '通过文件从外部数据库得到病人信息
            If gclsPros.OutFile <> "" Then
                gclsPros.PatiOut.Filter = "住院号= " & IIf(str住院号 = "", 0, str住院号) & IIf(lng主页ID = 0, "", " and 住院次数=" & lng主页ID)
                If gclsPros.PatiOut.EOF Then
                    If MsgBox("住院号为" & str住院号 & "且住院次数为1的病人在外部文件中没找到，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                        zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                        gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                        Exit Function
                    End If
                    blnOut = False
                Else
                    blnOut = True
                End If
            End If
            If blnOut Then
                gclsPros.主页ID = lng主页ID
                If gclsPros.IsSelPati Then Call LoadDataFromOutFile(str住院号)
            End If
            gclsPros.NoType = IT_New
            If gclsPros.NewInNo Then
                If gclsPros.OpenMode <> EM_新增首页 Then
                    strTemp = zlCommFun.ShowMsgbox(gclsPros.CurrentForm.Caption, "住院号为“" & str住院号 & "”的病人未找到，请确定操作方式？", "!新增病案(&A),新增首页(&N)", gclsPros.CurrentForm, vbQuestion)
                Else
                    strTemp = "新增病案"
                    gclsPros.OpenMode = EM_新增病案
                End If
                If strTemp = "新增病案" Then
                    gclsPros.OpenMode = EM_新增病案
                    gclsPros.CurrentForm.txtInfo(GC_病案号).Text = str住院号
                Else
                    gclsPros.OpenMode = EM_新增首页
                    gclsPros.NoType = IT_NewMed
                    If gclsPros.EditPageNo Then
                        gclsPros.CurrentForm.txtInfo(GC_病案号).Text = gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text
                    End If
                End If
                If gclsPros.OnLineNew Then
                    '与ZLHIS系统联机参数勾选,不能联机修改的项目，字体为蓝色
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_入院时间), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_出院时间), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_入院科室), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_出院科室), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_入院时间), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_出院时间), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_入院科室), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_出院科室), False)
                End If
            End If
            gclsPros.OnlyPatiInfo = False
            If Not blnOut Then
                strSql = "select 病人id from 病人信息 where 住院号=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, str住院号)
                If rsTmp.RecordCount > 0 Then
                    gclsPros.病人ID = Val(rsTmp!病人ID & "") '采用倒序，第一条记录就是最后一次
                    gclsPros.主页ID = 1
                    gclsPros.OnlyPatiInfo = True
                    lng主页ID = IIf(gclsPros.SinPageNo, lng主页ID, 0)
                    Call LoadMedPageData(gclsPros.病人ID, IIf(gclsPros.SinPageNo, gclsPros.主页ID, 0), , , gclsPros.Is编目)
                End If
            End If
            If gclsPros.NoType = IT_New And Not gclsPros.OnlyPatiInfo Then
                '新增病案,需重新确定病人ID
                gclsPros.病人ID = NVL(GetNextNo(1))    '住院使用用户自己输入的，但病人ID需要自动产生
                If Not blnOut Then gclsPros.主页ID = 1
            End If
            If blnOut Then
                If gclsPros.OutFile <> "" And Not gclsPros.IsSelPati Then
                    gclsPros.主页ID = Select外部主页id(str住院号, Val(lng主页ID))
                End If
            End If
            gclsPros.CurrentForm.txtSpecificInfo(SLC_入院次数) = gclsPros.主页ID
            If Not gclsPros.EditPageNo Or Trim(gclsPros.CurrentForm.txtInfo(GC_病案号).Text) = "" Then
                '不能编辑 或者 病案号为空，这时自动替换
                gclsPros.CurrentForm.txtInfo(GC_病案号).Text = gclsPros.InNo
            End If
            If gclsPros.EditPageNo Then
                '允许编辑病案号，允许停
                gclsPros.CurrentForm.txtInfo(GC_病案号).TabStop = True
            End If
            '53638:刘鹏飞,2013-05-10,新增档案号编号规则
            If gclsPros.UseFileRules Then
                gclsPros.CurrentForm.txtInfo(GC_档案号).Text = NVL(GetNextNo(CT_档案号, , GetDeptCode(gclsPros.出院科室ID)))
            End If

            If gclsPros.NoType = IT_New Then
                gclsPros.OpenMode = EM_新增病案
            Else
                gclsPros.OpenMode = EM_新增首页
            End If
            gclsPros.Is编目 = False
            LoadPatiByInNo = True
            Exit Function
        Else
            gclsPros.IsExistPati = True
        End If
        strSql = "select 病人id from 病案主页 where 住院号=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, str住院号)
        gclsPros.病人ID = Val(rsTmp!病人ID & "")
        '得出病人在病案处的主页信息
        strSql = "select 主页ID from 病案主页 where 病人ID=[1] and nvl(病人性质,0)=0 and 编目日期 is not null  order by 主页ID Desc"
        Set rs病人主页 = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, Val(gclsPros.病人ID & ""))
        If rs病人主页.RecordCount > 0 Then
            lng次数 = rs病人主页("主页ID") '采用倒序，第一条记录就是最后一次
            lng上次次数 = lng次数
        End If

        If Not gclsPros.IsSelPati Then
            '78747:取消首页只能依次往后添加的限制
            If gclsPros.OutFile = "" Then
                intSel主页id = Select主页ID(gclsPros.病人ID, IIf(blnOrderAdd = True, lng次数, 0))
            Else
                intSel主页id = Select外部主页id(str住院号, IIf(blnOrderAdd = True, lng次数, 0))
            End If
            If lng次数 > intSel主页id And intSel主页id <> 0 And blnOrderAdd = True Then
                ' 刘兴宏:主页id与实际的住院次数不一至,因为实际的住院次数不包含留观病人
                ' 2007/05/10
                Call Get住院次数Or主页id(gclsPros.病人ID, lng次数, lngTemp, False)
                lngTemp = IIf(lngTemp = 0, lng次数, lngTemp)
               '当住院次数大于所选择的住院次数时退出 刘飞2005-8-22
                MsgBox "请选择该病人在第" & lngTemp & "次入院以后的信息！", vbInformation, gstrSysName
                LoadPatiByInNo = False
                Exit Function
            ElseIf intSel主页id <> 0 Then
                '所选择的住院次数>建立病案的住院次数时建立所选择的信息建立病案
                lng次数 = intSel主页id
                blnNoCheck = lng次数 > intSel主页id
            ElseIf intSel主页id = 0 Then
                '63725:刘鹏飞,2013-08-06
                '刘兴宏:因为不存在需要编制的病案,可以退出了.
                If (Not gclsPros.OnLineNew And gclsPros.OnLine And gclsPros.OutFile = "") Or Not gclsPros.EditUnrecive Then
                    Call Get住院次数Or主页id(gclsPros.病人ID, lng次数, lngTemp, False)
                    lngTemp = IIf(lngTemp = 0, lng次数, lngTemp)
                    MsgBox "该病人总共" & lngTemp & "次住院,并且已经建立了病案,不能继续！", vbInformation, gstrSysName
                    LoadPatiByInNo = False
                    Exit Function
                End If
            End If
        Else
           If gclsPros.OutFile <> "" Then
                If lng次数 > lng主页ID Then
                    ' 刘兴宏:主页id与实际的住院次数不一至,因为实际的住院次数不包含留观病人
                    ' 2007/05/10
                    Call Get住院次数Or主页id(gclsPros.病人ID, lng次数, lngTemp, False)
                    lngTemp = IIf(lngTemp = 0, lng次数, lngTemp)
                    MsgBox "请选择该病人在第" & lngTemp & "次入院以后的信息！", vbInformation, gstrSysName
                    gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text = "": gclsPros.InNo = ""
                    LoadPatiByInNo = False
                    Exit Function
                End If
           End If
           lng次数 = lng主页ID
        End If
        '判断上次的出院情况
        If lng上次次数 > 0 And blnNoCheck = False Then
            strSql = "" & _
                "   Select B.出院情况 " & _
                "   From 病案主页 A,病人诊断记录 B " & _
                "   Where A.病人ID=[1] and A.主页ID=[2] " & _
                "           and A.病人ID = B.病人ID And A.主页ID = B.主页ID And B.诊断类型 = 3 And B.诊断次序 = 1 And B.编码序号 = 1 " & _
                "           and a.编目日期 is not null "
            Set rs病人主页 = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, lng上次次数)

            If rs病人主页.EOF = False Then
                If rs病人主页("出院情况") = "死亡" Then
                    MsgBox "该病人第" & lng上次次数 & "次出院情况已经填成死亡，不能再新增首页了。", vbInformation, gstrSysName
                    zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                    gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                    Exit Function
                End If
            End If
        End If
        If gclsPros.OutFile <> "" Then
             If intSel主页id <> 0 Or gclsPros.IsSelPati Then
                gclsPros.主页ID = lng次数
             Else
                lng次数 = lng次数 + 1
             End If
        Else
            '假如该已经有5份主页编码，那这将是他的第6次住院
            strSql = "" & _
                "   SELECT MIN(主页ID) as 住院次数 " & _
                "   FROM 病案主页 " & _
                "   WHERE 病人ID=[1] AND 主页ID> =[2]" & _
                "  AND 编目日期 IS NULL AND nvl(病人性质,0)=0" '未编目，且为正常住院

            Set rs病人主页 = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, lng次数)
            If IsNull(rs病人主页("住院次数")) Then

                If gclsPros.OnLine Then
                    '检查是否存在留观病人,如果存在,取最大的留观病人的主页ID+1
                    strSql = "" & _
                        "   SELECT max(主页ID) as 住院次数 " & _
                        "   FROM 病案主页 " & _
                        "   WHERE 病人ID=[1] AND 主页ID> =[2]" & _
                        "           AND nvl(病人性质,0)<>0" '最后编目日期的留观病人的主页ID
                    Set rs病人主页 = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, lng次数)

                    If rs病人主页.EOF Then
                        '证明不存在此住院次数,则取最大的住院次数
                        lng次数 = lng次数 + 1
                    Else
                        If Val(NVL(rs病人主页("住院次数"))) = 0 Then
                            lng次数 = lng次数 + 1
                        Else
                            '证明存在留观病人,因此需要以最大的留观病人的主页id+1
                            lng次数 = Val(NVL(rs病人主页("住院次数"))) + 1
                        End If
                    End If
                Else
                    lng次数 = lng次数 + 1
                End If
            Else
                lng次数 = rs病人主页("住院次数") '可能中间间隔了留观病人，所以不能直接取最大值+1
            End If
        End If

        '得出病人在院收费处的主页信息
        strSql = "Select 出院日期, 编目日期 from 病案主页 where 病人ID=[1] and 主页ID= [2] And nvl(病人性质,0)=0"
        Set rs病人主页 = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, lng次数)
        If rs病人主页.RecordCount <> 0 Then
            If IsNull(rs病人主页("出院日期")) Then
                MsgBox "该病人仍然在院，不能填写主页。", vbInformation, gstrSysName
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                Exit Function
            End If
        End If
        '63725:刘鹏飞,2013-08-06
        If Not gclsPros.EditUnrecive Then
            strSql = "Select ID from 病案接收记录 Where 病人ID=[1] and 主页ID= [2] And 接收时间 IS NOT NULL"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, lng次数)
            If rsTmp.RecordCount = 0 Then
                MsgBox "当前要编目的是病人第" & lng次数 & "次住院，但病案室还没有接收，不能进行编目操作!", vbInformation, gstrSysName
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                Exit Function
            End If
        End If

        If lng次数 > 1 Then
            ' 刘兴宏:主页id与实际的住院次数不一至,因为实际的住院次数不包含留观病人
            ' 2007/05/10
            lngTemp = 0
            If Get住院次数Or主页id(gclsPros.病人ID, lng次数, lngTemp, False) = False Then
                MsgBox "获取指定主页的次数失败,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
            lngTemp = IIf(lngTemp = 0, lng次数, lngTemp)
            If MsgBox("您正在输入病人的第" & lngTemp & "份首页，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                Exit Function
            End If
        End If
        blnNO主页ID = False
        If rs病人主页.RecordCount = 0 Then
            If gclsPros.OnLine Then
                 ' 刘兴宏:主页id与实际的住院次数不一至,因为实际的住院次数不包含留观病人
                ' 2007/05/10
                Call Get住院次数Or主页id(gclsPros.病人ID, lng次数, lngTemp, False)
                lngTemp = IIf(lngTemp = 0, lng次数, lngTemp)
                '问题32713 by lesfeng 2010-09-13
                If Not gclsPros.OnLineNew Then
                     MsgBox "该病人的第" & lngTemp & "次住院信息在收费系统中没找到，不能继续。", vbInformation, gstrSysName
                    zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                    gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                    Exit Function
                Else
                    MsgBox "该病人的第" & lngTemp & "次住院信息在收费系统中没找到，目前是新增收费系统中不存在的首页。", vbInformation, gstrSysName
                    '与ZLHIS系统联机参数勾选,不能联机修改的项目，字体为蓝色
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_入院时间), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_出院时间), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_入院科室), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_出院科室), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_入院时间), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdDateInfo(DC_出院时间), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_入院科室), False)
                    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_出院科室), False)
                End If
            Else
                If gclsPros.OutFile <> "" Then
                    '通过文件从外部数据库得到病人信息
                    gclsPros.PatiOut.Filter = "住院号= " & IIf(str住院号 = "", 0, str住院号) & IIf(lng次数 = 0, "", " and 住院次数=" & lng次数)
                    If gclsPros.PatiOut.EOF Then
                        If MsgBox("住院号为" & str住院号 & "的病人在外部文件中没找到，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbNo Then
                            zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                            gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                            Exit Function
                        End If
                    End If
                    blnOut = True
                End If
                blnNO主页ID = True
            End If
            '改为增加主页模式
            gclsPros.OpenMode = EM_新增首页
            If gclsPros.NewInNo Then
                '刘兴宏:需要新产生一个住院号
                gclsPros.InNo = NVL(GetNextNo(2))
                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text = gclsPros.InNo
            End If
            Call LoadDataFromPaitiInfo(gclsPros.病人ID, lng次数)
            If blnOut Then
                '刘兴宏
                gclsPros.主页ID = lng次数
                Call LoadDataFromOutFile(str住院号)
            End If
            gclsPros.Is编目 = False
        Else
            If IsNull(rs病人主页("出院日期")) Then
                MsgBox "该病人仍然在院，不能填写主页。", vbInformation, gstrSysName
                zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).SetFocus
                Exit Function
            End If
            Call LoadDataFromPaitiInfo(gclsPros.病人ID, lng次数)
            strSql = "Select 1 From 病人诊断记录 Where 病人ID=[1] And 主页ID=[2] And 记录来源=4 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, lng次数)
            '改为增加主页模式，且只是对已有主页进行编目
            gclsPros.OpenMode = EM_新增首页
            gclsPros.Is编目 = True
            Call LoadMedPageData(gclsPros.病人ID, lng次数, , , rsTmp.RecordCount > 0)
            If gclsPros.MedPageSandard = ST_四川省标准 Or gclsPros.MedPageSandard = ST_云南省标准 And gclsPros.FuncType = f病案首页 Then
                gclsPros.CurrentForm.cmdLastDiag.Visible = lng次数 > 1
                gclsPros.CurrentForm.lblDiagInfo.Caption = ""
                gclsPros.CurrentForm.lblDiagInfo.Visible = False
            End If
            Call SetPageVisible
            Call SetPicPosition(True)
        End If
        '0-当前病人的住院号是没经过验证的；1-住院号是新的；2-住院号是以前的:3-住院号是新增首页的
        gclsPros.NoType = IT_Old
        gclsPros.主页ID = lng次数
        Call ValidatePageNos
        lngTemp = 0
        If blnNO主页ID Then
        Else
            ' 刘兴宏:主页id与实际的住院次数不一至,因为实际的住院次数不包含留观病人
            ' 2007/05/10
            Call Get住院次数Or主页id(gclsPros.病人ID, gclsPros.主页ID, lngTemp, False)
            If intSel主页id = 0 And gclsPros.主页ID <> 1 And Not gclsPros.IsSelPati Then
                lngTemp = lngTemp + 1
            End If
        End If
        gclsPros.CurrentForm.txtSpecificInfo(SLC_入院次数).Text = IIf(lngTemp = 0, lng次数, lngTemp)
    End If
    LoadPatiByInNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub AfterLoadPatiByNo()
    Dim blnEditInfo As Boolean

    Call SetFaceEditable(gclsPros.IsSigned)
    If gclsPros.OpenMode <> EM_编辑 And Not gclsPros.Is编目 Then
        blnEditInfo = True '增加病案主页
    End If
    If gclsPros.NoType = IT_New Then
        blnEditInfo = True '增加病案主页
    End If
    With gclsPros.CurrentForm
        '与ZLHIS系统联机参数勾选,不能联机修改的项目，字体为蓝色
        Call SetCtrlLocked(.mskDateInfo(DC_入院时间), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.mskDateInfo(DC_出院时间), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.txtInfo(GC_入院科室), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.txtInfo(GC_出院科室), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.cmdDateInfo(DC_入院时间), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.cmdDateInfo(DC_出院时间), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.cmdInfo(GC_入院科室), Not blnEditInfo, , Not blnEditInfo)
        Call SetCtrlLocked(.cmdInfo(GC_出院科室), Not blnEditInfo, , Not blnEditInfo)
        '病案首页快速编辑设置
        .lblNote.Visible = Not gclsPros.EditUnrecive
        .cboSpecificInfo(SLC_年龄).TabStop = gclsPros.TabAgeUnit
        .cboBaseInfo(BCC_国籍).TabStop = gclsPros.TabNation
        .txtInfo(GC_档案号).TabStop = gclsPros.TabFileNo
        .txtInfo(GC_病案号).TabStop = gclsPros.EditPageNo And gclsPros.OpenMode = EM_新增病案
        .txtInfo(GC_病案号).Locked = Not gclsPros.EditPageNo
        .cboBaseInfo(BCC_付款方式).TabStop = gclsPros.TabPayType
        .chkInfo(CHK_再入院).TabStop = gclsPros.TabReadm
        .txtInfo(GC_X线号).TabStop = gclsPros.TabXRaysNo
        If Not gclsPros.EditPayType Then
            If .txtInfo(GC_病案号).Locked = False Then
                .txtInfo(GC_病案号).SetFocus
            ElseIf .txtSpecificInfo(SLC_住院号).Locked Then
                .txtInfo(GC_姓名).SetFocus
            Else
                .txtSpecificInfo(SLC_住院号).SetFocus
            End If
        Else
            '需要首先编码医疗付款方式
            .cboBaseInfo(BCC_付款方式).SetFocus
        End If
        
        .vsOPS.ColHidden(PI_助产护士) = Not gclsPros.Is产科
    End With
End Sub

Public Sub LoadDataFromOutFile(ByVal str住院号 As String)
'此处不传送空值给，因为有的系统根本就没传这些值过来
    Dim arrFileds As Variant, i As Long, strName As String

    On Error Resume Next
    arrFileds = Array("姓名", "出生日期", "出生地", "身份证号", "性别", "血型", "职业", "国籍", "民族", "婚姻状况", "联系人关系", "单位电话", "单位邮编", "单位地址", _
                                "家庭地址", "户口地址", "户口地址邮编", "家庭地址邮编", "联系人电话", "联系人地址", "联系人姓名", "医疗付款方式", "入院病情", "其他证件", _
                                "入院日期", "出院日期", "住院医师", "责任护士")
    For i = LBound(arrFileds) To UBound(arrFileds)
        strName = IIf(arrFileds(i) = "出生地", "出生地点", arrFileds(i))
        If Not IsNull(gclsPros.PatiOut(arrFileds(i))) Then
            Call SetCtrlValues(UCase(strName), gclsPros.PatiOut(arrFileds(i)) & "", , True)
        End If
    Next
    On Error GoTo errH
    '刘兴宏:20040812更改的
    gclsPros.FeesOut.Filter = "住院号 = " & IIf(str住院号 = "", 0, str住院号) & " and 住院次数=" & gclsPros.主页ID
    Call CacheLoadVsFreesData(gclsPros.CurrentForm.vsFees, gclsPros.FeesOut, , gclsPros.Is编目)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadDataFromPaitiInfo(ByVal lng病人ID As Long, Optional ByVal int住院次数 As Integer = 1) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id和主页id ,获取相关的病人信息,并填充到相关的控件中
    '参数:lng病人ID-病人id
    '     int住院次数-主页ID
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴宏
    '修改:2007/08/29
    '----------------------------------------------------------------------------------------------------------------------------
    Dim rs病人信息 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lng主页ID As Long
    Dim strSql As String
    Dim i As Long
    Dim strCode As String

    lng主页ID = IIf(gclsPros.SinPageNo, int住院次数, 0)
    Set gclsPros.PatiInfo = GetPatiMainInfoData(lng病人ID, lng主页ID)
    '加载病人信息
    If Not gclsPros.PatiInfo.EOF Then
        For i = 0 To gclsPros.PatiInfo.Fields.Count - 1
            If Not IsNull(gclsPros.PatiInfo.Fields(i).Value) Then
                Call SetCtrlValues(UCase(gclsPros.PatiInfo.Fields(i).Name & ""), gclsPros.PatiInfo.Fields(i).Value & "", , True)
            End If
        Next
    End If
    Err = 0: On Error GoTo errH
    '病案首页住院号，病案号，档案号等的生成
    strCode = gclsPros.PatiInfo!出院科室编码 & ""
    If strCode = "" Then strCode = gclsPros.PatiInfo!最后科室编码 & ""
    '住院号获取
    If IsNull(gclsPros.PatiInfo!住院号) Then
        gclsPros.InNo = NVL(GetNextNo(2))
    ElseIf gclsPros.NewInNo And IsHavePageNos(CT_住院号, Not gclsPros.OpenMode = EM_编辑 Or gclsPros.Is编目, gclsPros.PatiInfo!住院号 & "", gclsPros.病人ID) Then
        gclsPros.InNo = NVL(GetNextNo(2))
    Else
        gclsPros.InNo = gclsPros.PatiInfo!住院号 & ""
    End If
    gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text = gclsPros.InNo
    '病案号获取
    If IsNull(gclsPros.PatiInfo!病案号) Then
        If gclsPros.NewInNo Or Not gclsPros.SinPageNo And IsNull(gclsPros.PatiInfo!最后病案号) Then
            '如果是使用新的住院号,病案号强制默认为住院号
            '如果不存在住院病案号 , 则病案号 = 当前住院号
            gclsPros.CurrentForm.txtInfo(GC_病案号).Text = gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text
        ElseIf gclsPros.SinPageNo Then
            gclsPros.CurrentForm.txtInfo(GC_病案号).Text = NVL(GetNextNo(4, , strCode))
        ElseIf Not IsNull(gclsPros.PatiInfo!最后病案号) Then
            '如果当前次数不存在住院病案号,则取最后一次整理病案的病案号
            gclsPros.CurrentForm.txtInfo(GC_病案号).Text = gclsPros.PatiInfo!最后病案号 & ""
        End If
    Else
        gclsPros.CurrentForm.txtInfo(GC_病案号).Text = gclsPros.PatiInfo!病案号 & ""
    End If
    '53638:刘鹏飞,2013-05-10,新增档案号编号规则
    If IsNull(gclsPros.PatiInfo!最后档案号) And gclsPros.UseFileRules Then
        gclsPros.CurrentForm.txtInfo(GC_档案号).Text = NVL(GetNextNo(5, , strCode))
    Else
        gclsPros.CurrentForm.txtInfo(GC_档案号).Text = gclsPros.PatiInfo!最后档案号 & ""
    End If
    Call CacheLoadVsAllerData(gclsPros.CurrentForm.vsAller, GetAllerData(lng病人ID, lng主页ID))
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function SetCboDefaultValue(ByVal lngIndex As Long) As Boolean
'----------------------------------------------------------------------------------------------------------------------------
'功能:根据字典表重新设置某一ComboBox 的值和默认值
'参数:lngIndex: cboBaseInfo 控件的索引
'返回:成功,返回true,否则返回False
'----------------------------------------------------------------------------------------------------------------------------
    Dim j As Long
    Dim rsTmp As ADODB.Recordset
    Dim objCboTmp As ComboBox
    On Error GoTo errH
    Set objCboTmp = gclsPros.CurrentForm.cboBaseInfo(lngIndex)
    Set rsTmp = GetBaseCode(lngIndex)
    '清除原有数据
    objCboTmp.Clear
    objCboTmp.Tag = ""
    '装入数据
    If Not rsTmp.EOF Then
        For j = 1 To rsTmp.RecordCount
            If IsNull(rsTmp!编码) Then
                objCboTmp.AddItem rsTmp!名称
            Else
                objCboTmp.AddItem rsTmp!编码 & "-" & Chr(13) & rsTmp!名称
            End If
            objCboTmp.ItemData(objCboTmp.NewIndex) = NVL(rsTmp!ID, 0)
            If Val(rsTmp!缺省 & "") = 1 Then
                Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                objCboTmp.Tag = objCboTmp.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    SetCboDefaultValue = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetMainDirectory()
'功能:根据页面设置导航目录
    Dim myNod As Node
    Dim strTmp As String
    Dim strTitle() As String
    Dim i As Long, j As Long
    
    If gclsPros.MedPageSandard = ST_四川省标准 Then
        strTmp = "标题,基本信息,西医诊断,西医诊断情况,中医诊断,中医诊断情况,药物过敏,输血信息,签名信息,手术记录,住院费用,住院情况,化疗信息,放疗信息,抗菌药物使用情况,抗精神病治疗情况,重症监护情况,病案附加项,附页1,附页2"
    Else
        strTmp = "标题,基本信息,西医诊断,西医诊断情况,中医诊断,中医诊断情况,药物过敏,输血信息,签名信息,手术记录,住院费用,住院情况,化疗信息,放疗信息,抗菌药物使用情况,抗精神病治疗情况,重症监护情况,病案附加项,附页"
    End If
   
    '加载外挂附页目录
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            strTmp = strTmp & "," & gfrmMecCol(i).Caption
        Next
    End If
    
    j = 1
    strTitle = Split(strTmp, ",")
    
    frmMain.tvDirectory.Nodes.Clear
    frmMain.tvDirectory.LineStyle = tvwRootLines
    frmMain.tvDirectory.Indentation = 200
    
    With gclsPros.CurrentForm
        For i = .PicPage.LBound To .PicPage.UBound
            If .PicPage(i).Tag = "true" Then
                Set myNod = frmMain.tvDirectory.Nodes.Add(, , "key-" & i, j & ". " & strTitle(i))
                myNod.Expanded = True
                j = j + 1
            End If
        Next
    End With
    
End Sub

Public Function GetReplaceObject(ByVal vsfTmp As VSFlexGrid) As TextBox
'功能: 根据传入的VSFlexGrid在点击的行列出设置一个隐藏的TextBox控件
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim i As Long
    Dim vPoint As POINTAPI
    
    With gclsPros.CurrentForm
        If .txtInfo.UBound <> 9999 Then
            Load .txtInfo(9999)
        End If
        vPoint = GetCoordPos(vsfTmp.hwnd, vsfTmp.CellLeft, vsfTmp.CellTop)
        .txtInfo(9999).Visible = False
        Set .txtInfo(9999).Container = vsfTmp.Container
        lngLeft = vPoint.X - frmMain.Left - .picMain.Left - vsfTmp.Container.Left - frmMain.PicDirectory.Width - 200
        lngTop = vPoint.Y - frmMain.Top - frmMain.PicForm.Top - .picMain.Top - vsfTmp.Container.Top - .Top - 80
        .txtInfo(9999).Move lngLeft, lngTop, vsfTmp.ColWidth(vsfTmp.Col), vsfTmp.RowHeight(vsfTmp.Row)

        Set GetReplaceObject = .txtInfo(9999)
    End With
End Function


Private Sub LoadDiagAndAllerFData()
'功能：保存数据之后重新加载诊断，过敏信息
    Dim rsTmp As ADODB.Recordset
    Dim lng病人ID As Long, lng主页ID As Long
    Dim intMaxDiagSource As Integer
    Dim LngRow As Long, vsTmp As VSFlexGrid
    
    With gclsPros.CurrentForm
        lng病人ID = gclsPros.病人ID
        lng主页ID = gclsPros.主页ID
        
        '过敏信息加载
        If gclsPros.AddAller Then
            Call DeleteCacheRecInfo("过敏药物")
            Set rsTmp = GetAllerData(lng病人ID, lng主页ID)
            Call CacheLoadVsAllerData(.vsAller, rsTmp)
            gclsPros.AddAller = False
        End If
    
        '读取诊断
        If gclsPros.AddDiag Then
            Set rsTmp = GetPatiDiagData(lng病人ID, lng主页ID, IIf(gclsPros.PatiType <> PF_门诊, 1, 0), , Not gclsPros.Is编目, gclsPros.Moved)
            rsTmp.Filter = "记录来源=" & IIf(gclsPros.FuncType = f病案首页, 4, 3)
            intMaxDiagSource = IIf(gclsPros.FuncType = f病案首页, 4, -1)
            If gclsPros.FuncType = f病案首页 And rsTmp.EOF Then
                intMaxDiagSource = 3
                rsTmp.Filter = "记录来源=3"
                If rsTmp.EOF Then intMaxDiagSource = 2
            End If
            If Not gclsPros.Is复诊 Or gclsPros.Is复诊 And rsTmp.RecordCount = 0 Then
                '2、加载西医诊断
                Call DeleteCacheRecInfo("西医诊断")
                Call InitTableDiag
                Call CacheLoadVsDiagData(.vsDiagXY, rsTmp, IIf(gclsPros.PatiType <> PF_门诊, "1,2,3,5,6,7,10", "1"), , intMaxDiagSource)
                '3、加载中医诊断
                If gclsPros.Have中医 Then
                    Call DeleteCacheRecInfo("中医诊断")
                    Call CacheLoadVsDiagData(.vsDiagZY, rsTmp, IIf(gclsPros.PatiType <> PF_门诊, "11,12,13", "11"), , intMaxDiagSource)
                End If
                gclsPros.AddDiag = False
            End If
            
            Set vsTmp = .vsDiagXY
            With vsTmp
                .Cell(flexcpForeColor, 1, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
                .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
                If gclsPros.PatiType <> PF_门诊 Then
                    LngRow = FindDiagRow(DT_出院诊断XY)
                    .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
                    .Row = .FixedRows: .Col = DI_诊断描述
                    Call DiagAfterRowColChange(vsTmp, -1, -1, .Row, .Col)
                Else
                    .Cell(flexcpText, .FixedRows, DI_诊断类型, .Rows - 1, DI_诊断类型) = "西医"
                End If
            End With
    
            Set vsTmp = .vsDiagZY
            With vsTmp
                .Cell(flexcpForeColor, .FixedRows, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
                .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
                If gclsPros.PatiType <> PF_门诊 Then
                    LngRow = FindDiagRow(DT_出院诊断ZY)
                    .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
                    Call DiagAfterRowColChange(vsTmp, -1, -1, .Row, .Col)
                Else
                    .Cell(flexcpText, .FixedRows, DI_诊断类型, .Rows - 1, DI_诊断类型) = "中医"
                End If
            End With
            
        End If
    End With
End Sub

Public Sub DeleteCacheRecInfo(ByVal strInfoName As String)
'功能：删除信息记录集，一般应用于表格
'参数：strInfoName=信息名或控件名
    On Error GoTo errH
    '先依靠信息名寻找寻找，寻找不到时，再按控件名寻找
    gclsPros.MainInfoRec.Filter = "信息名='" & strInfoName & "'"
    If gclsPros.MainInfoRec.EOF Then gclsPros.MainInfoRec.Filter = "控件名='" & strInfoName & "'"
    If Not gclsPros.MainInfoRec.EOF Then
        Select Case gclsPros.MainInfoRec!ExpState
            Case ES_加载扩展
                Call Rec.Delete(gclsPros.SecdInfoRec, "序号=" & gclsPros.MainInfoRec!序号)
        End Select
    End If
    Exit Sub
errH:
    Debug.Print "DeleteCacheRecInfo:" & Err.Source & "===" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckValueChange(Optional ByRef objTmp As Object) As Boolean
'功能：检查首页控件的值是否发生变化
    Dim strOlsInfo As String
    Dim strCurInfo As String
    Dim strCboName As String
    Dim cboTmp As ComboBox
    Dim lngIndex As Long
    Dim blnFind As Boolean
    
    If gclsPros.InfosChange Then Exit Function
    If Not gclsPros.LoadFinish Then Exit Function
    If gclsPros.FuncType <> f医生首页 And gclsPros.FuncType <> f病案首页 Then Exit Function
    If gclsPros.OpenMode = EM_查阅 Then Exit Function
    On Error GoTo errH
    If frmMain.stbThis.Panels(2).Text <> "" Then
        frmMain.stbThis.Panels(2).Text = ""
    End If
    If objTmp Is Nothing Then
        gclsPros.InfosChange = True
        Exit Function
    End If
    If TypeName(objTmp) = "ComboBox" Then
        Set cboTmp = objTmp
        strCurInfo = cboTmp.Text
        strCboName = cboTmp.Name
        lngIndex = cboTmp.Index
    Else
        gclsPros.InfosChange = True
        Exit Function
    End If
    
    If strCboName = "cboBaseInfo" Or strCboName = "cboManInfo" Then
        gclsPros.MainInfoRec.Filter = "控件名='" & strCboName & "'" & "And Index=" & lngIndex
        If Not gclsPros.MainInfoRec.EOF Then
            strOlsInfo = NVL(gclsPros.MainInfoRec!信息原值)
            blnFind = True
        Else
            gclsPros.SecdInfoRec.Filter = "控件名='" & strCboName & "'" & "And IndexEx=" & lngIndex
            If Not gclsPros.SecdInfoRec.EOF Then
                strOlsInfo = NVL(gclsPros.SecdInfoRec!信息原值)
                blnFind = True
            End If
        End If
        If blnFind Then
            If strCurInfo <> strOlsInfo And blnFind Then
                gclsPros.InfosChange = True
            End If
        Else
            gclsPros.InfosChange = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub LocateVSFRowCol(ByRef vsfTmp As VSFlexGrid, ByVal lngMinRow As Long, ByVal lngMaxRow As Long, ByVal lngMinCol As Long, ByVal lngMaxCol As Long, ByVal LngRow As Long, ByVal LngCol As Long)
    If Not Between(vsfTmp.Row, lngMinRow, lngMaxRow) Then vsfTmp.Row = LngRow
    If Not Between(vsfTmp.Col, lngMinCol, lngMaxCol) Then vsfTmp.Col = LngCol
    If vsfTmp.ColHidden(vsfTmp.Col) = True Then vsfTmp.Col = LngCol
End Sub

Private Sub SetDeliceryInfo(ByRef vsDiagTmp As VSFlexGrid)
'功能：根据诊断相关信息，设置“分娩信息”按钮的可见性
    Dim bln西医 As Boolean
    Dim lngTmpRow As Long, i As Long, j As Long

    On Error GoTo errH
    With vsDiagTmp
        bln西医 = .Name = "vsDiagXY"
        If gclsPros.PatiType <> PF_门诊 Then
            If gclsPros.FuncType = f病案首页 Then
                If bln西医 Then                             '分娩设置
                    lngTmpRow = FindDiagRow(DT_病理诊断)
                    i = FindDiagRow(DT_出院诊断XY)
                    
                    gclsPros.CurrentForm.cmdDeliceryInfo.Visible = False
                    gclsPros.CurrentForm.cmdDeliceryInfo.Enabled = False
                    gclsPros.CurrentForm.cmdDeliceryInfo.Tag = ""
                    
                    For j = i To lngTmpRow - 1
                        If .TextMatrix(j, DI_分娩信息) = "1" Then
                            gclsPros.CurrentForm.cmdDeliceryInfo.Visible = True
                            gclsPros.CurrentForm.cmdDeliceryInfo.Enabled = True
                            gclsPros.CurrentForm.cmdDeliceryInfo.Tag = "1"
                            Exit For
                        End If
                    Next
                End If
            End If
        Else
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'解析字符串加入错误提示
Public Sub ErrDw(strMsg As String)
    'strMsg解析 例: 提示消息-(提醒:1/禁止:0)|控件Key值-提示消息-(提醒:1/禁止:0)|表格控件Key值-提示消息-(提醒:1/禁止:0)-Row-Col
    Dim i As Long
    Dim arrTmp() As String
    
    On Error GoTo errH
    With gclsPros.CurrentForm
        If strMsg <> "" Then
            ReDim Preserve arrTmp(UBound(Split(strMsg, "|")))
            arrTmp = Split(strMsg, "|")
            For i = 0 To UBound(arrTmp)
                Select Case UBound(Split(arrTmp(i), "-"))
                    Case 1 '只提示消息不绑定控件
                         Call AddErrInfo(Split(arrTmp(i), "-")(0), Val(Split(arrTmp(i), "-")(1)))
                    Case 2 '绑定控件提示消息
                         Call AddErrInfo(Split(arrTmp(i), "-")(1), Val(Split(arrTmp(i), "-")(2)), gColCtl.Item((Split(arrTmp(i), "-")(0))))
                    Case 4 '绑定表格控件提示消息
                        gColCtl.Item((Split(arrTmp(i), "-")(0))).Row = Val((Split(arrTmp(i), "-")(3)))
                        gColCtl.Item((Split(arrTmp(i), "-")(0))).Col = Val((Split(arrTmp(i), "-")(4)))
                        Call AddErrInfo(Split(arrTmp(i), "-")(1), Val(Split(arrTmp(i), "-")(2)), gColCtl.Item((Split(arrTmp(i), "-")(0))))
                End Select
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'解析集合信息加入错误提示
Public Sub ErrMec(colMsg As Collection)
    'colMsg解析 例: 控件对象:colMsg.item,colMsg.item.tag((提醒:1/禁止:0)|提示消息|PicPage的Index)
    Dim i As Long
    Dim arrTmp() As String
    
    On Error GoTo errH
    With gclsPros.CurrentForm
        For i = 1 To colMsg.Count
            ReDim Preserve arrTmp(UBound(Split(colMsg.Item(i).Tag, "|")))
            arrTmp = Split(colMsg.Item(i).Tag, "|")
            If UBound(arrTmp) <> 0 Then
                Call AddErrInfo(arrTmp(1), Val(arrTmp(0)), colMsg.Item(i))
            End If
            colMsg.Item(i).Tag = ""
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function CtlAdd() As Collection
    Dim colCtl As New Collection
    Dim i As Long
    On Error Resume Next
    With gclsPros.CurrentForm
        colCtl.Add .cboManInfo(MC_门诊医师), "门诊医师"
        colCtl.Add .cboManInfo(MC_科主任), "科主任"
        colCtl.Add .cboManInfo(MC_主任或副主任), "主任或副主任"
        colCtl.Add .cboManInfo(MC_进修医师), "进修医师"
        colCtl.Add .cboManInfo(MC_主治医师), "主治医师"
        colCtl.Add .cboManInfo(MC_住院医师), "住院医师"
        colCtl.Add .cboManInfo(MC_研究生医师), "研究生医师"
        colCtl.Add .cboManInfo(MC_实习医师), "实习医师"
        colCtl.Add .cboManInfo(MC_质控医师), "质控医师"
        colCtl.Add .cboManInfo(MC_质控护士), "质控护士"
        colCtl.Add .cboManInfo(MC_责任护士), "责任护士"
        colCtl.Add .cboManInfo(MC_编目员), "编目员"
        colCtl.Add .cboManInfo(MC_主诊医师), "主诊医师"
        colCtl.Add .mskDateInfo(DC_出生日期), "出生日期"
        colCtl.Add .mskDateInfo(DC_入院时间), "入院时间"
        colCtl.Add .mskDateInfo(DC_出院时间), "出院时间"
        colCtl.Add .mskDateInfo(DC_确诊日期), "确诊日期"
        colCtl.Add .mskDateInfo(DC_死亡时间), "死亡时间"
        colCtl.Add .mskDateInfo(DC_发病日期), "发病日期"
        colCtl.Add .mskDateInfo(DC_发病时间), "发病时间"
        colCtl.Add .mskDateInfo(DC_质控日期), "质控日期"
        colCtl.Add .mskDateInfo(DC_编目日期), "编目日期"
        colCtl.Add .mskDateInfo(DC_收回日期), "收回日期"
        colCtl.Add .txtAdressInfo(ADRC_出生地点), "出生地点"
        colCtl.Add .txtAdressInfo(ADRC_籍贯), "籍贯"
        colCtl.Add .txtAdressInfo(ADRC_现住址), "现住址"
        colCtl.Add .txtAdressInfo(ADRC_户口地址), "户口地址"
        colCtl.Add .txtAdressInfo(ADRC_联系人地址), "联系人地址"
        colCtl.Add .txtAdressInfo(ADRC_病人区域), "病人区域"
        colCtl.Add .txtAdressInfo(ADRC_单位地址), "单位地址"
        colCtl.Add .padrInfo(ADRC_出生地点), "出生地点(结构化)"
        colCtl.Add .padrInfo(ADRC_籍贯), "籍贯(结构化)"
        colCtl.Add .padrInfo(ADRC_现住址), "现住址(结构化)"
        colCtl.Add .padrInfo(ADRC_户口地址), "户口地址(结构化)"
        colCtl.Add .padrInfo(ADRC_联系人地址), "联系人地址(结构化)"
        colCtl.Add .cboBaseInfo(BCC_付款方式), "付款方式"
        colCtl.Add .cboBaseInfo(BCC_性别), "性别"
        colCtl.Add .cboBaseInfo(BCC_婚姻), "婚姻"
        colCtl.Add .cboBaseInfo(BCC_职业), "职业"
        colCtl.Add .cboBaseInfo(BCC_国籍), "国籍"
        colCtl.Add .cboBaseInfo(BCC_民族), "民族"
        colCtl.Add .cboBaseInfo(BCC_关系), "关系"
        colCtl.Add .cboBaseInfo(BCC_入院途径), "入院途径"
        colCtl.Add .cboBaseInfo(BCC_文化程度), "文化程度"
        colCtl.Add .cboBaseInfo(BCC_去向), "去向"
        colCtl.Add .cboBaseInfo(BCC_感染与死亡关系), "感染与死亡关系"
        colCtl.Add .cboBaseInfo(BCC_入院情况), "入院情况"
        colCtl.Add .cboBaseInfo(BCC_分化程度), "分化程度"
        colCtl.Add .cboBaseInfo(BCC_最高诊断依据), "最高诊断依据"
        colCtl.Add .cboBaseInfo(BCC_门诊与出院XY), "门诊与出院XY"
        colCtl.Add .cboBaseInfo(BCC_入院与出院XY), "入院与出院XY"
        colCtl.Add .cboBaseInfo(BCC_门诊与入院), "门诊与入院"
        colCtl.Add .cboBaseInfo(BCC_术前与术后), "术前与术后"
        colCtl.Add .cboBaseInfo(BCC_放射与病理), "放射与病理"
        colCtl.Add .cboBaseInfo(BCC_临床与病理), "临床与病理"
        colCtl.Add .cboBaseInfo(BCC_死亡期间), "死亡期间"
        colCtl.Add .cboBaseInfo(BCC_临床与尸检), "临床与尸检"
        colCtl.Add .cboBaseInfo(BCC_门诊与出院ZY), "门诊与出院ZY"
        colCtl.Add .cboBaseInfo(BCC_入院与出院ZY), "入院与出院ZY"
        colCtl.Add .cboBaseInfo(BCC_辩证), "辩证"
        colCtl.Add .cboBaseInfo(BCC_治法), "治法"
        colCtl.Add .cboBaseInfo(BCC_方药), "方药"
        colCtl.Add .cboBaseInfo(BCC_治疗类别), "治疗类别"
        colCtl.Add .cboBaseInfo(BCC_中医诊疗设备), "中医诊疗设备"
        colCtl.Add .cboBaseInfo(BCC_抢救方法), "抢救方法"
        colCtl.Add .cboBaseInfo(BCC_中医诊疗技术), "中医诊疗技术"
        colCtl.Add .cboBaseInfo(BCC_自制中药制剂), "自制中药制剂"
        colCtl.Add .cboBaseInfo(BCC_辨证施护), "辨证施护"
        colCtl.Add .cboBaseInfo(BCC_病案质量), "病案质量"
        colCtl.Add .cboBaseInfo(BCC_病例分型), "病例分型"
        colCtl.Add .cboBaseInfo(BCC_HBsAg), "HBsAg"
        colCtl.Add .cboBaseInfo(BCC_血型), "血型"
        colCtl.Add .cboBaseInfo(BCC_HCVAb), "HCVAb"
        colCtl.Add .cboBaseInfo(BCC_RH), "RH"
        colCtl.Add .cboBaseInfo(BCC_HIVAb), "HIVAb"
        colCtl.Add .cboBaseInfo(BCC_输液反应), "输液反应"
        colCtl.Add .cboBaseInfo(BCC_输血反应), "输血反应"
        colCtl.Add .cboBaseInfo(BCC_输血前9项检查), "输血前9项检查"
        colCtl.Add .cboBaseInfo(BCC_生育状况), "生育状况"
        colCtl.Add .cboBaseInfo(BCC_出院方式), "出院方式"
        colCtl.Add .cboBaseInfo(BCC_再入院计划天数), "再入院计划天数"
        colCtl.Add .cboBaseInfo(BCC_压疮发生期间), "压疮发生期间"
        colCtl.Add .cboBaseInfo(BCC_压疮分期), "压疮分期"
        colCtl.Add .cboBaseInfo(BCC_跌倒或坠床伤害), "跌倒或坠床伤害"
        colCtl.Add .cboBaseInfo(BCC_跌倒或坠床原因), "跌倒或坠床原因"
        colCtl.Add .cboBaseInfo(BCC_距上次住院时间), "距上次住院时间(下拉)"
        colCtl.Add .cboBaseInfo(BCC_重返间隔时间), "重返间隔时间"
        colCtl.Add .cboBaseInfo(BCC_约束方式), "约束方式"
        colCtl.Add .cboBaseInfo(BCC_约束工具), "约束工具"
        colCtl.Add .cboBaseInfo(BCC_约束原因), "约束原因"
        colCtl.Add .cboBaseInfo(BCC_新生儿离院方式), "新生儿离院方式"
        colCtl.Add .cboBaseInfo(BCC_肿瘤分期), "肿瘤分期"
        colCtl.Add .cboBaseInfo(BCC_临床路径管理), "临床路径管理"
        colCtl.Add .cboBaseInfo(BCC_法定传染病), "法定传染病"
        colCtl.Add .cboBaseInfo(BCC_实施DGRS管理), "实施DGRS管理"
        colCtl.Add .cboBaseInfo(BCC_死亡患者尸检), "死亡患者尸检"
        colCtl.Add .cboBaseInfo(BCC_身份证), "身份证"
        colCtl.Add .cboBaseInfo(BCC_变异原因), "变异原因(下拉)"
        colCtl.Add .cboBaseInfo(BCC_健康卡号), "健康卡号"
        colCtl.Add .chkInfo(CHK_再入院), "再入院"
        colCtl.Add .chkInfo(CHK_入院前外院治疗), "入院前外院治疗"
        colCtl.Add .chkInfo(CHK_是否确诊), "是否确诊"
        colCtl.Add .chkInfo(CHK_病原学检查), "病原学检查"
        colCtl.Add .chkInfo(CHK_新发肿瘤), "新发肿瘤"
        colCtl.Add .chkInfo(CHK_危重), "危重"
        colCtl.Add .chkInfo(CHK_急症), "急症"
        colCtl.Add .chkInfo(CHK_疑难), "疑难"
        colCtl.Add .chkInfo(CHK_示教病案), "示教病案"
        colCtl.Add .chkInfo(CHK_科研病案), "科研病案"
        colCtl.Add .chkInfo(CHK_疑难病例), "疑难病例"
        colCtl.Add .chkInfo(CHK_随诊), "随诊"
        colCtl.Add .chkInfo(CHK_CT), "CT"
        colCtl.Add .chkInfo(CHK_MRI), "MRI"
        colCtl.Add .chkInfo(CHK_彩色多普勒), "彩色多普勒"
        colCtl.Add .chkInfo(CHK_传染病上传), "传染病上传"
        colCtl.Add .chkInfo(CHK_围术期死亡), "围术期死亡"
        colCtl.Add .chkInfo(CHK_术后猝死), "术后猝死"
        colCtl.Add .chkInfo(CHK_进入路径), "进入路径"
        colCtl.Add .chkInfo(CHK_完成路径), "完成路径"
        colCtl.Add .chkInfo(CHK_变异), "变异"
        colCtl.Add .chkInfo(CHK_住院出现危重), "住院出现危重"
        colCtl.Add .chkInfo(CHK_是否同一疾病), "是否同一疾病"
        colCtl.Add .chkInfo(CHK_人工气道脱出), "人工气道脱出"
        colCtl.Add .chkInfo(CHK_重返重症医学科), "重返重症医学科"
        colCtl.Add .chkInfo(CHK_住院物理约束), "住院物理约束"
        colCtl.Add .chkInfo(CHK_单病种管理), "单病种管理"
        colCtl.Add .chkInfo(CHK_细菌标本送检), "细菌标本送检"
        colCtl.Add .chkInfo(CHK_会诊情况), "会诊情况"
        colCtl.Add .chkInfo(CHK_无过敏记录), "无过敏记录"
        colCtl.Add .txtInfo(GC_病案号), "病案号"
        colCtl.Add .txtInfo(GC_档案号), "档案号"
        colCtl.Add .txtInfo(GC_X线号), "X线号"
        colCtl.Add .txtInfo(GC_姓名), "姓名"
        colCtl.Add .txtInfo(GC_其他证件), "其他证件"
        colCtl.Add .txtInfo(GC_联系人姓名), "联系人姓名"
        colCtl.Add .txtInfo(GC_入院科室), "入院科室"
        colCtl.Add .txtInfo(GC_入院病房), "入院病房"
        colCtl.Add .txtInfo(GC_出院科室), "出院科室"
        colCtl.Add .txtInfo(GC_出院病房), "出院病房"
        colCtl.Add .txtInfo(GC_医保号), "医保号"
        colCtl.Add .txtInfo(GC_摘要), "摘要"
        colCtl.Add .txtInfo(GC_门诊号), "门诊号"
        colCtl.Add .txtInfo(GC_监护人), "监护人"
        colCtl.Add .txtInfo(GC_发病地址), "发病地址"
        colCtl.Add .txtInfo(GC_医学警示), "医学警示"
        colCtl.Add .txtInfo(GC_其他医学警示), "其他医学警示"
        colCtl.Add .txtInfo(GC_病理号), "病理号"
        colCtl.Add .txtInfo(GC_死亡原因), "死亡原因"
        colCtl.Add .txtInfo(GC_病原学诊断), "病原学诊断"
        colCtl.Add .txtInfo(GC_抢救病因), "抢救病因"
        colCtl.Add .txtInfo(GC_输其他), "输其他"
        colCtl.Add .txtInfo(GC_转入医疗机构), "转入医疗机构"
        colCtl.Add .txtInfo(GC_31天内再住院), "31天内再住院"
        colCtl.Add .txtInfo(GC_转科1), "转科1"
        colCtl.Add .txtInfo(GC_转科2), "转科2"
        colCtl.Add .txtInfo(GC_转科3), "转科3"
        colCtl.Add .txtInfo(GC_退出原因), "退出原因"
        colCtl.Add .txtInfo(GC_变异原因), "变异原因"
        colCtl.Add .txtInfo(GC_重症监护室名称), "重症监护室名称"
        colCtl.Add .txtInfo(GC_肿瘤T), "肿瘤T"
        colCtl.Add .txtInfo(GC_肿瘤N), "肿瘤N"
        colCtl.Add .txtInfo(GC_肿瘤M), "肿瘤M"
        colCtl.Add .txtInfo(GC_Email), "Email"
        colCtl.Add .txtInfo(GC_其他会诊), "其他会诊"
        colCtl.Add .txtInfo(GC_引发药物), "引发药物"
        colCtl.Add .txtInfo(GC_临床表现), "临床表现"
        colCtl.Add .txtInfo(GC_透析尿素氮值), "透析尿素氮值"
        colCtl.Add .txtInfo(GC_其他关系), "其他关系"
        colCtl.Add .txtInfo(GC_入院转入), "入院转入"
        colCtl.Add .optInput(OP_再住院无), "再住院无"
        colCtl.Add .optInput(OP_再住院有), "再住院有"
        colCtl.Add .optInput(OP_初诊), "初诊"
        colCtl.Add .optInput(OP_复诊), "复诊"
        colCtl.Add .optInput(OP_ICU无), "ICU无"
        colCtl.Add .optInput(OP_ICU有), "ICU有"
        colCtl.Add .optDiag(PC_XY按诊断输入), "XY按诊断输入"
        colCtl.Add .optDiag(PC_XY按疾病编码输入), "XY按疾病编码输入"
        colCtl.Add .optDiag(PC_ZY按诊断输入), "ZY按诊断输入"
        colCtl.Add .optDiag(PC_ZY按疾病编码输入), "ZY按疾病编码输入"
        colCtl.Add .optDiag(PC_按诊断输入), "按诊断输入"
        colCtl.Add .optDiag(PC_按疾病编码输入), "按疾病编码输入"
        colCtl.Add .optAller(PC_按药品目录输入), "按药品目录输入"
        colCtl.Add .optAller(PC_按过敏源输入), "按过敏源输入"
        colCtl.Add .OptParaOPSInfo(PC_按诊疗项目输入), "按诊疗项目输入"
        colCtl.Add .OptParaOPSInfo(PC_按ICDCM9编码输入), "按ICDCM9编码输入"
        colCtl.Add .chkParaOPSInfo(PC_未找到时自由录入), "未找到时自由录入"
        colCtl.Add .cboSpecificInfo(SLC_年龄), "年龄(下拉)"
        colCtl.Add .cboSpecificInfo(SLC_婴幼儿年龄), "婴幼儿年龄(下拉)"
        colCtl.Add .cboSpecificInfo(SLC_随诊期限), "随诊期限(下拉)"
        colCtl.Add .txtSpecificInfo(SLC_单位电话), "单位电话"
        colCtl.Add .txtSpecificInfo(SLC_单位邮编), "单位邮编"
        colCtl.Add .txtSpecificInfo(SLC_家庭电话), "家庭电话"
        colCtl.Add .txtSpecificInfo(SLC_家庭邮编), "家庭邮编"
        colCtl.Add .txtSpecificInfo(SLC_户口邮编), "户口邮编"
        colCtl.Add .txtSpecificInfo(SLC_身高), "身高"
        colCtl.Add .txtSpecificInfo(SLC_身高单位), "身高单位"
        colCtl.Add .txtSpecificInfo(SLC_体重), "体重"
        colCtl.Add .txtSpecificInfo(SLC_体重单位), "体重单位"
        colCtl.Add .txtSpecificInfo(SLC_体温), "体温"
        colCtl.Add .txtSpecificInfo(SLC_入院次数), "入院次数"
        colCtl.Add .txtSpecificInfo(SLC_收缩压), "收缩压"
        colCtl.Add .txtSpecificInfo(SLC_舒张压), "舒张压"
        colCtl.Add .txtSpecificInfo(SLC_联系人电话), "联系人电话"
        colCtl.Add .txtSpecificInfo(SLC_年龄), "年龄"
        colCtl.Add .txtSpecificInfo(SLC_婴幼儿年龄), "婴幼儿年龄"
        colCtl.Add .txtSpecificInfo(SLC_新生儿出生体重), "新生儿出生体重"
        colCtl.Add .txtSpecificInfo(SLC_新生儿入院体重), "新生儿入院体重"
        colCtl.Add .txtSpecificInfo(SLC_住院天数), "住院天数"
        colCtl.Add .txtSpecificInfo(SLC_住院号), "住院号"
        colCtl.Add .txtSpecificInfo(SLC_抢救次数), "抢救次数"
        colCtl.Add .txtSpecificInfo(SLC_成功次数), "成功次数"
        colCtl.Add .txtSpecificInfo(SLC_特护), "特护"
        colCtl.Add .txtSpecificInfo(SLC_一级护理), "一级护理"
        colCtl.Add .txtSpecificInfo(SLC_二级护理), "二级护理"
        colCtl.Add .txtSpecificInfo(SLC_三级护理), "三级护理"
        colCtl.Add .txtSpecificInfo(SLC_ICU), "ICU"
        colCtl.Add .txtSpecificInfo(SLC_CCU), "CCU"
        colCtl.Add .txtSpecificInfo(SLC_输红细胞), "输红细胞"
        colCtl.Add .txtSpecificInfo(SLC_输血小板), "输血小板"
        colCtl.Add .txtSpecificInfo(SLC_输血浆), "输血浆"
        colCtl.Add .txtSpecificInfo(SLC_输全血), "输全血"
        colCtl.Add .txtSpecificInfo(SLC_自体回收), "自体回收"
        colCtl.Add .txtSpecificInfo(SLC_呼吸机使用), "呼吸机使用"
        colCtl.Add .txtSpecificInfo(SLC_昏迷时间入院前_天), "昏迷时间入院前_天"
        colCtl.Add .txtSpecificInfo(SLC_昏迷时间入院前_小时), "昏迷时间入院前_小时"
        colCtl.Add .txtSpecificInfo(SLC_昏迷时间入院前_分钟), "昏迷时间入院前_分钟"
        colCtl.Add .txtSpecificInfo(SLC_昏迷时间入院后_天), "昏迷时间入院后_天"
        colCtl.Add .txtSpecificInfo(SLC_昏迷时间入院后_小时), "昏迷时间入院后_小时"
        colCtl.Add .txtSpecificInfo(SLC_昏迷时间入院后_分钟), "昏迷时间入院后_分钟"
        colCtl.Add .txtSpecificInfo(SLC_随诊期限), "随诊期限"
        colCtl.Add .txtSpecificInfo(SLC_费用和), "费用和"
        colCtl.Add .txtSpecificInfo(SLC_约束总时间), "约束总时间"
        colCtl.Add .txtSpecificInfo(SLC_重症监护天), "重症监护天"
        colCtl.Add .txtSpecificInfo(SLC_重症监护小时), "重症监护小时"
        colCtl.Add .txtSpecificInfo(SLC_Apgar), "Apgar"
        colCtl.Add .txtSpecificInfo(SLC_QQ), "QQ"
        colCtl.Add .txtSpecificInfo(SLC_输白蛋白), "输白蛋白"
        colCtl.Add .txtSpecificInfo(SLC_院内会诊), "院内会诊"
        colCtl.Add .txtSpecificInfo(SLC_外院会诊), "外院会诊"
        colCtl.Add .txtSpecificInfo(SLC_距上次住院时间), "距上次住院时间"
        colCtl.Add .vsTransfer, "转科情况(表格)"
        colCtl.Add .vsDiagXY, "西医诊断(表格)"
        colCtl.Add .vsDiagZY, "中医诊断(表格)"
        colCtl.Add .vsAller, "过敏信息(表格)"
        colCtl.Add .vsOPS, "手术记录(表格)"
        colCtl.Add .vsFees, "住院费用(表格)"
        colCtl.Add .vsChemoth, "化疗记录信息(表格)"
        colCtl.Add .vsRadioth, "放疗记录信息(表格)"
        colCtl.Add .vsKSS, "抗菌药物使用情况(表格)"
        colCtl.Add .vsSpirit, "抗精神病治疗情况(表格)"
        colCtl.Add .vsFlxAddICU, "重症监护情况(表格)"
        colCtl.Add .vsfMain, "病案附加项目(表格)"
        colCtl.Add .vsTSJC, "特殊检查情况(表格)"
        colCtl.Add .lstAdvEvent, "不良事件(表格)"
        colCtl.Add .lstInfection, "感染因素(表格)"
        colCtl.Add .lvwFee, "住院费用(树形表)"
        colCtl.Add .padrInfo(ADRC_单位地址), "单位地址(结构化)"
        colCtl.Add .txtSpecificInfo(SLC_婴幼儿年龄_DAY), "婴幼儿年龄_DAY"
        colCtl.Add .vsICUInstruments, "器械导管使用情况(表格)"
        colCtl.Add .vsInfect, "病人感染记录(表格)"
        colCtl.Add .lstInfectParts, "感染部位(表格)"
        colCtl.Add .vsSample, "病人病原学检查(表格)"
    End With
    Set CtlAdd = colCtl
    On Error GoTo 0
End Function

Private Sub MsgDis(str疾病IDs As String, str诊断IDs As String)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSql As String
    On Error GoTo Errhand
    '判断当前病人是否填写传染病报告卡
    strSql = "Select 文件ID From 电子病历记录 Where 病人ID=[1] And 主页ID=[2] And 病历种类=5 and 创建人=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "MsgDis", gclsPros.病人ID, gclsPros.主页ID, UserInfo.姓名)
    If rsTmp.RecordCount > 0 Then
        '判断用户是否修改或删除诊断
        strSql = ""
        If str疾病IDs <> "" Then
            strSql = " Union Select 疾病id,诊断id From 疾病报告前提 Where 疾病ID IN (Select Column_Value From Table(f_Num2list([3])))"
        End If
        If str诊断IDs <> "" Then
            strSql = strSql & " Union Select 疾病id,诊断id From 疾病报告前提 Where 诊断ID IN (Select Column_Value From Table(f_Num2list([4])))"
        End If
        strSql = "Select a.疾病id, a.诊断id From 病人诊断记录 A, 疾病报告前提 B Where a.病人id = [1] And a.主页id = [2] And a.编码序号 = 1 And (a.疾病id = b.疾病id Or a.诊断id = b.诊断id) " & IIf(strSql = "", "", "Minus (" & Mid(strSql, 8) & ") ")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "MsgDis", gclsPros.病人ID, gclsPros.主页ID, str疾病IDs, str诊断IDs)
        If rsTmp.RecordCount > 0 Then
            MsgBox "当前病人传染病诊断数据发生了改变,请修改传染病报告卡！", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub


