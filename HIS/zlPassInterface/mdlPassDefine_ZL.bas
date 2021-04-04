Attribute VB_Name = "mdlPassDefine_ZL"
Option Explicit

Public gobjFrm As frmPass      '悬浮主窗体
Public grsRet       As ADODB.Recordset   '缓存当前病人最近一次审查结果,切换病人清空;处方上传时提供警示级别
Public Const conMenu_EditPopup = 3    '编辑
Public Const conMenu_Drug_View = 30821 '查看药品说明书
Public Const conMenu_Drug_Match = 5 '配伍禁忌
Public Const conMenu_PAR_SET = 6    '参数设置
Public Const conMenu_FRM_VISIBLE = 7 '隐藏\显示主界面
Public Const conCOLOR_BULE As Long = &HD48A00
Public Const conCOLOR_TITLE_BAR As Long = 16298544 '16298544 rgb(48,178,248); 14392064 'RGB(0, 155, 219)
Public Const conCOLOR_BULELIGHT As Long = &HE4B440

Public Const conSTR_Key_Tip     As String = "适应症,用法用量,不良反应,禁忌症,注意事项,孕妇用药,儿童用药,老年人用药,相互作用,药物过量"
Public gstrParaTip As String

'Public gobjAir As zl9ComLib.clsAirBubble      'zl9ComLib.clsAirBubble
Public gobjAir As Object        '气泡提示

Private mstrPharmDept   As String    '中联药师审方启用科室
Private mstrPassDept    As String    '合理用药监测审查功能启用科室
Private mstrHosName     As String    '医院名称

Public Sub ZLShowWindow()
    If gobjFrm Is Nothing Then Set gobjFrm = New frmPass
    Call gobjFrm.Show
End Sub

Public Sub ZLCloseWindow()
    Unload gobjFrm
    Set gobjFrm = Nothing
End Sub
 
Public Function ZLGetDrugCode(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal str挂号单 As String, _
    Optional ByRef rsAdvice As ADODB.Recordset, Optional ByRef blnIsHaveOut As Boolean, _
    Optional ByRef rsOut As ADODB.Recordset, Optional ByVal bytFunc As Byte = 0) As String
'功能:返回需要审查的药品本位码
'   bytFunc=3 上传处方审查
    Dim i As Long, k As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim int频率次数 As Integer, int频率间隔 As Integer
    Dim blnDo As Boolean, blnAsk As Boolean
    
    Dim str收费细目IDs As String
    Dim str诊疗项目IDs As String
    Dim str给药途径    As String, str期效 As String
    Dim str频率编码    As String, str间隔单位 As String
    Dim str中药组IDs    As String, str相关ID As String
    Dim str医嘱ID       As String, str单量 As String, str单量单位 As String
    Dim str医嘱IDs      As String
    
    Dim str开嘱医生 As String, str开嘱医生Tag As String
    Dim str诊断相关ID As String   '记录相关ID便于门诊取对应诊断
    Dim rsDoct As ADODB.Recordset
    Dim rsDrug As ADODB.Recordset
    
    Dim curDate As Date

    On Error GoTo errH
    
    curDate = zlDatabase.Currentdate
    Set rsAdvice = InitAdviceRS(FUN_医嘱信息_ZL)
    '启用了禁忌药品说明参数  且场合为门诊编辑审查功能
    If (glngModel = PM_住院编辑 Or glngModel = PM_门诊编辑) And (gbytReason = 1 Or gbytReason = 0 And InStr("," & mstrPharmDept & ",", "," & gobjPati.lngDeptID & ",") > 0) Then
        Set rsOut = InitAdviceRS(FUN_输出内容)
    End If
    Select Case glngModel
    Case PM_门诊编辑, PM_住院编辑
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_门诊编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                            And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-DD") = Format(curDate, "yyyy-MM-DD")
                ElseIf glngModel = PM_住院编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                    If blnDo Then
                        blnDo = (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" And InStr(",4,8,9,", .TextMatrix(i, gobjCOL.intCOL状态)) = 0 _
                                Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") _
                                And .TextMatrix(i, gobjCOL.intCOL状态) <> "4")
                    End If
                End If
                '反向问诊需要传入诊疗项目ID 排除药品
                If gstrIP <> "" And bytFunc = 0 Then
                    blnAsk = Not InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL诊疗项目ID)) <> 0 _
                    And Val(.TextMatrix(i, gobjCOL.intCOLEDIT)) = 1
                Else
                    blnAsk = False
                End If
                If blnDo Then
                    If Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 Then
                        str收费细目IDs = str收费细目IDs & IIf(str收费细目IDs = "", "", ",") & .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                    ElseIf Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) = 0 Then
                        If InStr("," & str诊疗项目IDs & ",", "," & .TextMatrix(i, gobjCOL.intCOL诊疗项目ID) & ",") = 0 Then
                            str诊疗项目IDs = str诊疗项目IDs & IIf(str诊疗项目IDs = "", "", ",") & .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                        End If
                    End If
                    If glngModel = PM_住院编辑 Then
                        If Val(.TextMatrix(i, gobjCOL.intCOL执行性质)) <> 5 And Val(.TextMatrix(.FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID))), gobjCOL.intCOL执行性质)) = 5 Then
                            blnIsHaveOut = True
                        End If
                    End If
                    '取药品给药途径
                    If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) <> Val(.TextMatrix(i - 1, gobjCOL.intCOL相关ID)) Then str给药途径 = "" '一并给药不重复取
                    If str给药途径 = "" Then
                        k = .FindRow(CLng(.TextMatrix(i, gobjCOL.intCOL相关ID)), i + 1)
                        If k <> -1 Then str给药途径 = Val(.TextMatrix(k, gobjCOL.intCOL诊疗项目ID))   '传代码
                    End If
                    Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), int频率次数, int频率间隔, str间隔单位, IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), str频率编码)
                    
                    rsAdvice.AddNew
                    rsAdvice!医嘱ID = .RowData(i)
                    rsAdvice!单次量 = .TextMatrix(i, gobjCOL.intCOL单量)
                    rsAdvice!计量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                    rsAdvice!诊疗项目ID = .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                    rsAdvice!药品ID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                    rsAdvice!输液组号 = .TextMatrix(i, gobjCOL.intCOL相关ID)
                    rsAdvice!给药频次 = str频率编码
                    rsAdvice!给药频次名称 = .TextMatrix(i, gobjCOL.intCOL频率)
                    rsAdvice!给药途径 = str给药途径
                    rsAdvice!每日量 = Get每日量(.TextMatrix(i, gobjCOL.intCOL单量), str间隔单位, int频率次数, int频率间隔, .TextMatrix(i, gobjCOL.intCOL频率))
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
                    If bytFunc = 3 Then
                        str医嘱IDs = str医嘱IDs & "," & .RowData(i)
                        rsAdvice!开嘱时间 = Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd HH:MM:SS")
                        rsAdvice!紧急标志 = IIf(.TextMatrix(i, gobjCOL.intCol标志) = "1", "1", "0") '0-普通,1-紧急，2-补录
                        rsAdvice!医嘱状态 = .TextMatrix(i, gobjCOL.intCOL状态)
                        
                        '开嘱医生
                        str开嘱医生 = .TextMatrix(i, gobjCOL.intCOL开嘱医生)
                        If InStr(str开嘱医生, "/") > 0 Then str开嘱医生 = Mid(str开嘱医生, 1, InStr(str开嘱医生, "/") - 1)
                        If InStr("," & str开嘱医生Tag & ",", "," & str开嘱医生 & ",") = 0 Then
                            str开嘱医生Tag = str开嘱医生Tag & "," & str开嘱医生
                        End If
                        rsAdvice!开嘱医生 = str开嘱医生
                        rsAdvice!医嘱内容 = .TextMatrix(i, gobjCOL.intCOL医嘱内容)
                        If glngModel = PM_门诊编辑 Then
                            rsAdvice!用药天数 = .TextMatrix(i, gobjCOL.intCOL天数)   'OP 门诊处方有效
                            '处方诊断
                            If InStr("," & str诊断相关ID & ",", "," & .TextMatrix(i, gobjCOL.intCOL相关ID) & ",") = 0 Then
                                str诊断相关ID = str诊断相关ID & "," & .TextMatrix(i, gobjCOL.intCOL相关ID)
                            End If
                            rsAdvice!医嘱期效 = "1"
                        Else
                            rsAdvice!医嘱期效 = IIf(.TextMatrix(i, gobjCOL.intCOL期效) = "临嘱", "1", "0")
                        End If

                        '抗菌药物的预防、治疗
                        If .TextMatrix(i, gobjCOL.intcol用药目的) = "1" Then
                            rsAdvice!用药目的 = "预防"
                        ElseIf .TextMatrix(i, gobjCOL.intcol用药目的) = "2" Then
                            rsAdvice!用药目的 = "治疗"
                        Else
                            rsAdvice!用药目的 = ""
                        End If
                        '上次审查时缓存数据
                        If Not grsRet Is Nothing Then
                            grsRet.Filter = "OrderId =" & .RowData(i)
                            If Not grsRet.EOF Then
                                rsAdvice!药品禁忌等级 = IIf(grsRet!Level & "" = "", "慎用", grsRet!Level & "")
                                rsAdvice!药品禁忌类型 = grsRet!Type & ""
                            End If
                        End If
                        rsAdvice!药品禁忌说明 = .TextMatrix(i, gobjCOL.intCol禁忌药品说明)
                    End If
                ElseIf blnAsk Then
                    rsAdvice.AddNew
                    rsAdvice!医嘱ID = .RowData(i)
                    rsAdvice!诊疗项目ID = .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                    rsAdvice!性质 = 1 '标识反向问诊
                End If
            Next
            If rsAdvice.RecordCount > 0 Then rsAdvice.UpdateBatch
        End With
    Case PM_门诊医嘱清单, PM_住院医嘱清单
        With gobjAdvice
            For i = .FixedRows To .Rows - 1
               If glngModel = PM_门诊医嘱清单 Then
                    blnDo = (Val(.TextMatrix(i, gobjCOL.intCOLID)) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 _
                        Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4")) _
                        And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                            
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
                    If (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4") Then
                        str中药组IDs = str中药组IDs & "," & .TextMatrix(i, gobjCOL.intCOLID)
                    Else
        
                        Call Get频率信息_名称(.TextMatrix(i, gobjCOL.intCOL频率), int频率次数, int频率间隔, str间隔单位, IIf(.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "7", 2, 1), str频率编码)
                        If Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0 Then
                            str收费细目IDs = str收费细目IDs & IIf(str收费细目IDs = "", "", ",") & .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                        End If
                        '单量，单量单位
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                        str医嘱IDs = str医嘱IDs & "," & str医嘱ID
                        If glngModel = PM_门诊医嘱清单 Then
                            str单量 = Val(.TextMatrix(i, gobjCOL.intCOL单量))
                            str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量)
                            str单量 = FormatEx(str单量, 5)
                            str单量单位 = Replace(str单量单位, str单量, "")
                        Else
                            str单量 = .TextMatrix(i, gobjCOL.intCOL单量)
                            str单量单位 = .TextMatrix(i, gobjCOL.intCOL单量单位)
                            str单量 = Replace(str单量, str单量单位, "")
                        End If
                        rsAdvice.AddNew
                        rsAdvice!医嘱ID = str医嘱ID
                        rsAdvice!单次量 = str单量
                        rsAdvice!计量单位 = str单量单位
                        rsAdvice!诊疗项目ID = .TextMatrix(i, gobjCOL.intCOL诊疗项目ID)
                        rsAdvice!药品ID = .TextMatrix(i, gobjCOL.intCOL收费细目ID)
                        rsAdvice!输液组号 = .TextMatrix(i, gobjCOL.intCOL相关ID)
                        rsAdvice!给药频次 = str频率编码
                        rsAdvice!给药频次名称 = .TextMatrix(i, gobjCOL.intCOL频率)
                        rsAdvice!给药途径 = ""
                        rsAdvice!每日量 = Get每日量(str单量, str间隔单位, int频率次数, int频率间隔, .TextMatrix(i, gobjCOL.intCOL频率))
                        rsAdvice.Update
                        
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
                    End If
                End If
            Next
            '由于医嘱清单配方的特殊性,需要从数据库提取中药名称
            If glngModel = PM_住院医嘱清单 Or glngModel = PM_门诊医嘱清单 Then
                If str中药组IDs <> "" Then
                    Set rsDrug = Get中药配方(str中药组IDs)
                    With rsDrug
                        For i = 1 To .RecordCount
                            If !相关ID & "" <> str相关ID Then
                                Call Get频率信息_名称(!频率 & "", 0, 0, "", IIf(!诊疗类别 & "" = "7", 2, 1), str频率编码)
                                str相关ID = !相关ID & ""
                            End If
                            rsAdvice.AddNew
                            rsAdvice!医嘱ID = !id & ""
                            rsAdvice!单次量 = !单次用量 & ""
                            rsAdvice!计量单位 = !单量单位 & ""
                            rsAdvice!诊疗项目ID = !诊疗项目ID & ""
                            rsAdvice!药品ID = !药品ID & ""
                            rsAdvice!输液组号 = !相关ID & ""
                            rsAdvice!给药频次 = str频率编码
                            rsAdvice!给药途径 = !用法ID & ""
                            rsAdvice!每日量 = Get每日量(str单量, str间隔单位, int频率次数, int频率间隔, !频率 & "")
                            rsAdvice.Update
                            .MoveNext
                        Next
                    End With
                End If
            End If
     
        End With
    End Select
    If str医嘱IDs <> "" Then
        Set rsDrug = GetDrugPlus(lng病人ID, lng主页ID, str挂号单, str医嘱IDs)
        rsAdvice.Filter = "性质=0"
        For i = 1 To rsAdvice.RecordCount
            rsDrug.Filter = "ID=" & rsAdvice!医嘱ID
            If Not rsDrug.EOF Then
                rsAdvice!给药途径 = rsDrug!给药途径ID & ""
                If bytFunc = 3 Then
                    rsAdvice!给药途径名称 = rsDrug!给药途径名称 & ""
                    rsAdvice!剂型 = rsDrug!药品剂型 & ""
                    rsAdvice!毒理分类 = rsDrug!毒理分类 & ""
                    rsAdvice!超量说明 = rsDrug!超量说明 & ""
                    rsAdvice!药品抗菌药物等级 = rsDrug!抗生素 & ""
                End If
            End If
            rsAdvice.MoveNext
        Next
    End If
    If str收费细目IDs <> "" Then
        Set rsDrug = GetRS("药品规格 A", "A.药品ID,A.本位码", str收费细目IDs, "A.药品ID")
        rsAdvice.Filter = "性质=0"
        For i = 1 To rsAdvice.RecordCount
            rsDrug.Filter = "药品ID=" & Val(rsAdvice!药品ID)
            If Not rsDrug.EOF Then
                rsAdvice!本位码 = rsDrug!本位码 & ""
            End If
            rsAdvice.MoveNext
        Next
    End If
    If str诊疗项目IDs <> "" Then
        Set rsDrug = GetRS("药品规格 A,收费项目目录 B", "A.药名ID,A.本位码", str诊疗项目IDs, "A.药品ID=B.ID And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 IS NULL) And A.药名ID")
        rsAdvice.Filter = "性质=0"
        For i = 1 To rsAdvice.RecordCount
            If rsAdvice!本位码 = "" Then
                rsDrug.Filter = "药名ID=" & Val(rsAdvice!诊疗项目ID)
                If Not rsDrug.EOF Then
                    rsAdvice!本位码 = rsDrug!本位码 & ""
                End If
            End If
            rsAdvice.MoveNext
        Next
    End If
    
    If str开嘱医生Tag <> "" And bytFunc = 3 Then
        str开嘱医生Tag = Mid(str开嘱医生Tag, 2)
        Set rsDoct = GetDoctorInfo(str开嘱医生Tag, IIf(glngModel = PM_门诊编辑, 2, 1))
        rsAdvice.Filter = "性质=0"
        For i = 1 To rsAdvice.RecordCount
            rsDoct.Filter = "姓名='" & rsAdvice!开嘱医生 & "'"
            If Not rsDoct.EOF Then
                rsAdvice!医生职称 = rsDoct!聘任技术职务 & ""
                rsAdvice!医生抗菌药物等级 = rsDoct!级别 & ""
            End If
            rsAdvice.MoveNext
        Next
    End If
    
    If glngModel = PM_门诊编辑 And bytFunc = 3 And str诊断相关ID <> "" Then
        str诊断相关ID = Mid(str诊断相关ID, 2)
        Set rsDoct = GetOutAdviceDiagsInfo(str诊断相关ID)
        rsAdvice.Filter = "性质=0"
        For i = 1 To rsAdvice.RecordCount
            rsDoct.Filter = "医嘱ID=" & rsAdvice!输液组号
            Do While Not rsDoct.EOF
                rsAdvice!处方诊断 = rsAdvice!处方诊断 & "|" & rsDoct!诊断描述
                rsDoct.MoveNext
            Loop
            rsAdvice!处方诊断 = Mid(rsAdvice!处方诊断 & "", 2)
            rsAdvice.MoveNext
        Next
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function AdviceCheckWarn_ZL(ByVal lngPatiID As Long, ByVal str挂号单 As String, ByVal lng主页ID As String, _
     Optional ByVal bytFunc As Byte, Optional rsOut As ADODB.Recordset, Optional ByRef objMap As clsPassMap) As Boolean
'功能：调用中联用药监测系统对医嘱进行合理用药审查等相关功能
'
'参数：bytFunc=0 编辑界面审查;1-医生站审查;3-医嘱下达界面保存医嘱后调用,将处方上传处方审查系统
'返回值:
    Dim strRet As String
    Dim strPara As String, strParaAsk As String
    Dim strUrl  As String
    Dim str医嘱期效 As String, str状态 As String, str结束时间 As String, str医嘱ID As String
    Dim strOld As String
    
    Dim rsAdvice    As ADODB.Recordset
    
    Dim bytRet      As Byte
    Dim i As Long, k As Long
    Dim lngBegin As Long, lngEnd As Long
    
    Dim blnIsHaveOut As Boolean
    Dim blnDo As Boolean, blnNoSave As Boolean
    Dim arrSQL As Variant
    Dim arrLight(0 To 4) As String
    Dim datCurr As Date
    
    On Error GoTo errH
    If gblnBreak Then Exit Function
    strPara = ZL_MakeDetailXML(lngPatiID, lng主页ID, str挂号单, rsAdvice, rsOut, blnIsHaveOut, bytFunc)
    If strPara = "" Then AdviceCheckWarn_ZL = True: Exit Function
    '反向问诊
LineAsk:
    If (glngModel = PM_门诊编辑 Or glngModel = PM_住院编辑) And gstrIP <> "" And bytFunc = 0 Then
        strParaAsk = strPara
        Call AskPatiStatus(rsAdvice, strPara, lngPatiID)
    End If
    '合理用药审查
    strUrl = "http://" & gstrDrugIP & ":" & gstrDrugPort & "/DrugCorrect/CheckContent"
    WriteLog "" & glngModel, "AdviceCheckWarn_ZL", "审查URL:" & strUrl & ",审查XML:" & strPara
    strRet = HttpPost(strUrl, strPara, responseText, "text/plain", , , gblnBreak)
    strRet = Replace(strRet, """", "")
    '<errormsg>错误信息</errormsg>
    WriteLog "" & glngModel, "AdviceCheckWarn_ZL", "审查结果:" & strRet
    If strRet = "" Or InStr(strRet, "<errormsg>") > 0 Or gblnBreak Then
        gobjFrm.SetNotifyIcon
        gsngCheckLinkTime = Timer
        AdviceCheckWarn_ZL = True '异常断开,允许保存医嘱
        Exit Function
    End If
    '上传审查系统
    If bytFunc = 3 Then Exit Function
    
    Set grsRet = ZL_ParseXML(strRet)
    If grsRet.RecordCount > 0 Then
        Call frmPassResultZL.ShowMe(gfrmMain, grsRet, bytFunc, bytRet, blnIsHaveOut)
        If bytRet = 3 Then
            strPara = strParaAsk
            GoTo LineAsk
        End If
    End If

    If bytFunc > 1 Then Exit Function
    With gobjAdvice
        arrSQL = Array()
        datCurr = zlDatabase.Currentdate
        '获取医嘱审查结果,并填写警示灯
        '-------------------------------------------------------------
        '返回值顺：0-蓝灯(默认),1-橙灯(慎用 或 空),2-红灯(禁用),3-黄灯(注意),4-黑灯(禁止)
        '警示级顺：0-蓝灯,3-黄灯,1-橙灯,2-红灯,4-黑灯
        arrLight(0) = "蓝_4":    arrLight(1) = "橙_4":  arrLight(2) = "红_4": arrLight(3) = "黄_4": arrLight(4) = "黑_4"
        If glngModel = PM_门诊编辑 Or glngModel = PM_门诊医嘱清单 Then
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_门诊编辑 Then
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0
                    blnDo = blnDo And Format(.Cell(flexcpData, i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(datCurr, "yyyy-MM-dd")
                Else
                    blnDo = ((InStr(",5,6,7,", "," & .TextMatrix(i, gobjCOL.intCOL诊疗类别) & ",") > 0 And Val(.TextMatrix(i, gobjCOL.intCOL收费细目ID)) <> 0) _
                    Or (.TextMatrix(i, gobjCOL.intCOL诊疗类别) = "E" And .TextMatrix(i, gobjCOL.intCol操作类型) = "4"))
                    blnDo = blnDo And Format(.TextMatrix(i, gobjCOL.intCOL开嘱时间), "yyyy-MM-dd") = Format(datCurr, "yyyy-MM-dd")
                End If
                    
                If blnDo Then
                    If glngModel = PM_门诊医嘱清单 Then
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID)
                    Else
                        str医嘱ID = .RowData(i)
                    End If
                    grsRet.Filter = "OrderId = '" & str医嘱ID & "'"
                    grsRet.Sort = "WarnLevel DESC"
                    If grsRet.RecordCount > 0 Then
                         k = CLng(grsRet!Light & "")
                    Else
                         k = 0 '医嘱清单中药配方
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
                            If strOld <> CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) Or Val(.TextMatrix(i, gobjCOL.intCOLEDIT)) = 1 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                blnNoSave = True    '标记为未保存
                            End If
                            '记录下禁忌药品 K=2 代表红灯 且 只针对未校对医嘱进行禁忌药品说明原因的标记,已经校对发送的医嘱不处理
                            If k = 2 And Not rsOut Is Nothing Then
                                rsOut.Filter = "医嘱ID = " & str医嘱ID & " And 状态 < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                            End If
                        ElseIf PM_门诊医嘱清单 = glngModel Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                            End If
                        End If
                    End If
                End If
            Next
            For i = LBound(arrSQL) To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
            Next
        ElseIf glngModel = PM_住院编辑 Or glngModel = PM_住院医嘱清单 Then
            For i = .FixedRows To .Rows - 1
                If glngModel = PM_住院编辑 Then
                    '住院编辑界面加载医嘱时已经屏蔽掉作废医嘱及停止和确认停止的长嘱
                    blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, gobjCOL.intCOL诊疗类别)) > 0 _
                            And Val(.TextMatrix(i, gobjCOL.intCOL婴儿)) = gobjPati.int婴儿 _
                            And (gbytUseType <> 1 Or (gbytUseType = 1 And .Cell(flexcpChecked, i, gobjCOL.intCOL选择) <> 2))
                    blnDo = blnDo And (.TextMatrix(i, gobjCOL.intCOL期效) = "长嘱" Or .TextMatrix(i, gobjCOL.intCOL期效) = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd") = Format(datCurr, "yyyy-MM-dd"))
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
                            (str医嘱期效 = "长嘱" And (InStr(",8,9,", str状态) > 0 And str结束时间 > Format(datCurr, "yyyy-MM-dd") Or InStr(",1,2,3,5,6,7,", str状态) > 0) Or _
                            str医嘱期效 = "临嘱" And Format(.Cell(flexcpData, i, gobjCOL.intCOL开始时间), "yyyy-MM-dd") = Format(datCurr, "yyyy-MM-dd")))
                    End If
                End If
                If blnDo Then
                    If glngModel = PM_住院编辑 Then
                        str医嘱ID = .RowData(i) & ""
                    Else
                        str医嘱ID = .TextMatrix(i, gobjCOL.intCOLID) & ""
                    End If
                    grsRet.Filter = "OrderId='" & str医嘱ID & "'"
                    grsRet.Sort = "WarnLevel DESC"
                    If grsRet.RecordCount > 0 Then
                        k = CLng(grsRet!Light & "")
                    Else
                        k = 0
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
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Or Val(.TextMatrix(i, gobjCOL.intCOLEDIT)) = 1 Then
                                .Cell(flexcpData, i, gobjCOL.intCOL序号) = 1
                                blnNoSave = True    '标记为未保存
                            End If
                            
                            If Not rsOut Is Nothing And k = 2 Then
                                rsOut.Filter = "医嘱ID=" & CLng(str医嘱ID) & " And 状态 < 3 "
                                If rsOut.RecordCount = 1 Then rsOut!是否禁忌 = 1
                            End If
                        ElseIf PM_住院医嘱清单 = glngModel Then
                            If CStr(.Cell(flexcpData, i, gobjCOL.intCOL警示)) <> strOld Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & str医嘱ID & "," & IIf(k >= 0 And k <= 4, k, "NULL") & ")"
                            End If
                        End If
                    End If
                End If
            Next
            For i = LBound(arrSQL) To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), G_STR_PASS)
            Next
        End If
    End With

    If bytRet = 1 Then  '修改处方
        Exit Function
    ElseIf bytRet = 2 Then '允许保存
        If bytFunc = 0 Then
            grsRet.Filter = "Light = 2"
            If grsRet.RecordCount > 0 Then
                If gbytBlackLamp = 1 Then
                    If (gbytReason = 1 Or gbytReason = 0 And InStr("," & mstrPharmDept & ",", "," & gobjPati.lngDeptID & ",") > 0) Then
                        If Not AddDrugReason(objMap, rsOut) Then Exit Function
                    Else
                        If MsgBox("审查发现禁忌用药，您确定要忽略吗？", vbOKCancel + vbQuestion + vbDefaultButton2, gstrSysName) = vbCancel Then
                            Exit Function
                        End If
                    End If
                Else
                    If blnIsHaveOut And gbytOutBlackLamp = 1 Then
                        If MsgBox("存在院外执行的药品审查发现禁忌用药，您确定要忽略吗？", vbOKCancel + vbQuestion + vbDefaultButton2, gstrSysName) = vbCancel Then
                            Exit Function
                        End If
                    End If
                End If
            Else
                If MsgBox("审查发现慎用药品，您确定要忽略吗？", vbOKCancel + vbQuestion + vbDefaultButton2, gstrSysName) = vbCancel Then
                    Exit Function
                End If
            End If
        End If
    End If
    AdviceCheckWarn_ZL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Function JSONParse(ByVal strJSONPath As String, ByVal strJSONData As String) As Variant
    Dim objJSON As Object
    Dim strValue As String
    
    On Error GoTo errH
    Set objJSON = CreateObject("MSScriptControl.ScriptControl")
    objJSON.Language = "JScript"
    strValue = NVL(objJSON.eval("JSON=" & strJSONData & ";JSON." & strJSONPath & ";"))
    JSONParse = JSONReplace(strValue)
    Set objJSON = Nothing
    Exit Function
errH:
    MsgBox "JSONParse 错误号:" & Err.Number & "错误描述:" & Err.Description, vbOKOnly, gstrSysName
End Function

Public Function JSONReplace(ByVal strJson As String)
'功能:JSON中特殊字符串转换
    If strJson <> "" Then
        strJson = Replace(strJson, "\n", vbLf)
        strJson = Replace(strJson, "\r", vbCr)
        strJson = Replace(strJson, "\t", vbTab)
        strJson = Trim(strJson)
    End If
    JSONReplace = strJson
End Function

Public Function ZL_GetPara() As Boolean
        Dim arrList As Variant
        Dim strPara As String
        
10      On Error GoTo errH
20      strPara = zlDatabase.GetPara(90001, glngSys, , "") '读取URLs 固定读取ZLHIS 系统默认100
        '格式服务器IP&&服务器端口号
30      If strPara = "" Then Exit Function
40      arrList = Split(strPara, G_STR_SPLIT)
50      If UBound(arrList) >= 3 Then
60          gstrDrugIP = arrList(0)
70          gstrDrugPort = arrList(1)
80          If Val(arrList(2)) > 10 Then
90              gsngWaitTime = 10
100         ElseIf Val(arrList(2)) < 1 Then
110             gsngWaitTime = 1   '访问等待3s
120         Else
130             gsngWaitTime = Val(arrList(2))
140         End If
150         If Val(arrList(3)) > 10 Then
160             gsngAutoLinkTime = 10   '10分钟
170         ElseIf Val(arrList(3)) < 0 Then
180             gsngAutoLinkTime = 1   ' 1分钟
190         Else
200             gsngAutoLinkTime = Val(arrList(3))
210         End If
220     Else
230         gstrDrugIP = ""
240         gstrDrugPort = ""
250         gsngWaitTime = 3
260         gsngAutoLinkTime = 5
270         Exit Function
280     End If
290     mstrHosName = zlRegInfo("单位名称", , 0)
300     gstrIP = GetParaURL("知识库", "反向问诊")
310     gstrStatusEdit = GetParaURL("知识库", "病人状态编辑")
320     gstrStatusGet = GetParaURL("知识库", "病人状态查询")
330     gstrStatusSave = GetParaURL("知识库", "病人状态保存")
340     gstrParaTip = zlDatabase.GetPara(299, glngSys)
350     strPara = GetParaURL("药师处方审查", "启用科室查询")
360     If strPara <> "" Then
370         strPara = HttpGet(strPara, responseText, 1)
            '{"items":[{"dept_ids":"168,143,159,149,151,156,148,158,"}],"first":{"$ref":"http://192.168.0.231:8080/ords/zlrecipe/recipe/getenabledept"}}
380         WriteLog "" & glngModel, "ZL_GetPara", "启用科室查询:" & strPara
390         If strPara <> "" Then
400             mstrPharmDept = JSONParse("items[0].dept_ids", strPara)
410         End If
420     End If
430     strPara = GetParaURL("知识库", "审查科室查询")
440     If strPara <> "" Then
450         strPara = HttpGet(strPara, responseText, 1)
            '{"items":[{"dept_ids":"132,122,433,473,138,148,149,129,151,168,515,147,144,143,159,514,152,146,157,141,135,145,155,557,150,163,154,106,219,235,236,226,230,237,228,238,513,224,229,231,234,"}],"hasMore":false,"limit":1000,"offset":0,"count":1,"links":[{"rel":"self","href":"http://192.168.0.231:8080/ords/rudrug/para/getenableddeptlist"},{"rel":"describedby","href":"http://192.168.0.231:8080/ords/rudrug/metadata-catalog/para/item"},{"rel":"first","href":"http://192.168.0.231:8080/ords/rudrug/para/getenableddeptlist"}]}
460         WriteLog "" & glngModel, "ZL_GetPara", "审查科室查询:" & strPara
470         If strPara <> "" Then
480             mstrPassDept = JSONParse("items[0].dept_ids", strPara)
490         End If
500     End If
510     ZL_GetPara = True
520     Exit Function
errH:
530     MsgBox "读取参数失败！" & vbNewLine & "ZL_GetPara:第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function ZL_SetPara() As String
    ZL_SetPara = IIf(gstrDrugIP = "", "192.168.6.17", gstrDrugIP) & G_STR_SPLIT & IIf(gstrDrugPort = "", "80", gstrDrugPort) & _
        G_STR_SPLIT & IIf(gsngWaitTime = 0, 3, gsngWaitTime) & G_STR_SPLIT & IIf(gsngAutoLinkTime = 0, 5, gsngAutoLinkTime)
End Function

Public Function ZL_MakeDetailXML(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, _
    Optional ByRef rsAdvice As ADODB.Recordset, Optional ByRef rsOut As ADODB.Recordset, Optional ByRef blnIsHaveOut As Boolean, _
    Optional ByVal bytFunc As Byte) As String
'功能：构造details XML字符串
    Dim strXML As String
    Dim strTmp As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset
    
    Dim colPati  As Collection
    Dim lng挂号ID As Long
    Dim i As Long
    Dim blnTran As Boolean
    
    On Error GoTo errH
    
    Set rsTmp = GetPatiInfo_YF(lng病人ID, str挂号单, lng主页ID)
    If rsTmp.EOF Then Exit Function
    
    gobjPati.lngDeptID = Val(rsTmp!当前科室ID & "")
    If mstrPassDept <> "" And InStr("," & mstrPassDept & ",", "," & gobjPati.lngDeptID & ",") = 0 Then Exit Function
    
    Set colPati = New Collection
    If str挂号单 <> "" Then
        lng挂号ID = rsTmp!就诊Id
        '附加信息
        strSQL = "Select b.项目名称, b.记录内容" & vbNewLine & _
                        "From 病人护理记录 A, 病人护理内容 B" & vbNewLine & _
                        "Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng病人ID, lng挂号ID)
        rsPatiInfo.Filter = "项目名称='身高'"
        If rsPatiInfo.RecordCount > 0 Then
            colPati.Add rsPatiInfo!记录内容 & "", "身高"
        Else
            colPati.Add "", "身高"
        End If
        rsPatiInfo.Filter = "项目名称='体重'"
        If rsPatiInfo.RecordCount > 0 Then
            colPati.Add rsPatiInfo!记录内容 & "", "体重"
        Else
            colPati.Add "", "体重"
        End If
    Else
        colPati.Add rsTmp!身高 & "", "身高"
        colPati.Add rsTmp!体重 & "", "体重"
    End If
    
    colPati.Add rsTmp!年龄 & "", "年龄"
    colPati.Add rsTmp!年龄数字 & "", "年龄数字"
    colPati.Add Get年龄周期(Val(rsTmp!年龄数字 & "")), "年龄周期"
    colPati.Add NVL(rsTmp!性别), "性别"
    colPati.Add NVL(rsTmp!职业), "职业"

    '病人生理情况
    strTmp = "," & Get病人病生理情况(lng病人ID, IIf(str挂号单 <> "", 0, lng主页ID)) & ","
    If InStr(strTmp, ",妊娠,") > 0 Then
        colPati.Add "1", "妊娠"
    Else
        colPati.Add "0", "妊娠"
    End If
    If InStr(strTmp, ",哺乳,") > 0 Then
        colPati.Add "1", "哺乳"
    Else
        colPati.Add "0", "哺乳"
    End If
    If InStr(strTmp, ",肝功能不全,") > 0 Then
        colPati.Add "1", "肝功能不全"
    Else
        colPati.Add "0", "肝功能不全"
    End If
    If InStr(strTmp, ",严重肝功能不全,") > 0 Then
        colPati.Add "1", "严重肝功能不全"
    Else
        colPati.Add "0", "严重肝功能不全"
    End If
    If InStr(strTmp, ",肾功能不全,") > 0 Then
        colPati.Add "1", "肾功能不全"
    Else
        colPati.Add "0", "肾功能不全"
    End If
    If InStr(strTmp, ",严重肾功能不全,") > 0 Then
        colPati.Add "1", "严重肾功能不全"
    Else
        colPati.Add "0", "严重肾功能不全"
    End If
    
    If bytFunc = 3 Then
        '处方审查信息
        colPati.Add "2", "提交类型"
        'colPati.Add lng病人ID, "病人ID"
        If str挂号单 <> "" Then
            'colPati.Add lng挂号ID, "就诊ID"
            colPati.Add rsTmp!门诊号 & "", "门诊号"
            colPati.Add "1", "病人来源"   '1-门诊;2-住院
        Else
            'colPati.Add lng主页ID, "就诊ID"
            colPati.Add rsTmp!住院号 & "", "住院号"
            colPati.Add rsTmp!当前病区ID & "", "就诊病区ID"
            colPati.Add rsTmp!当前病区 & "", "就诊病区"
            colPati.Add "2", "病人来源"   '1-门诊;2-住院
        End If
        colPati.Add rsTmp!姓名 & "", "姓名"
        colPati.Add Format(NVL(rsTmp!出生日期), "YYYY-MM-DD HH:MM:SS"), "出生日期"
        colPati.Add Format(NVL(rsTmp!入院时间), "YYYY-MM-DD HH:MM:SS"), "入院日期"
        colPati.Add rsTmp!当前床号 & "", "当前床号"
        colPati.Add rsTmp!当前科室 & "", "就诊科室"
        colPati.Add rsTmp!当前科室ID & "", "就诊科室ID"
        colPati.Add "0", "婴儿"     '1-婴儿;0-非婴儿
        colPati.Add "100", "HIS_NO"
    Else
        colPati.Add "1", "提交类型"
    End If
    
    '诊断信息
    Set rsTmp = Get病人诊断记录(lng病人ID, IIf(str挂号单 <> "", lng挂号ID, lng主页ID), IIf(str挂号单 <> "", "1,11", "2,12"))
    strTmp = "": strXML = ""
    For i = 1 To rsTmp.RecordCount
        strTmp = strTmp & IIf(i = 1, "", ",") & rsTmp!编码
        strXML = strXML & IIf(i = 1, "", ",") & rsTmp!名称
        rsTmp.MoveNext
    Next
    colPati.Add strXML, "诊断名称"
    colPati.Add strTmp, "诊断"
    
    '医嘱信息
    Call ZLGetDrugCode(lng病人ID, lng主页ID, str挂号单, rsAdvice, blnIsHaveOut, rsOut, bytFunc)
    strXML = ZL_GET_Details(colPati, rsAdvice, IIf(str挂号单 <> "", 1, 2), bytFunc)
    WriteLog "" & glngModel, "ZL_MakeDetailXML", "传入临时表XML:" & strXML
    
    gcnOracle.BeginTrans: blnTran = True
    'XML写入临时表:中联合理用药参数
    Call sys.SaveLob(glngSys, 30, "", strXML, 1)
    '调用过程:对写入临时表的参数内容进行更新
    strSQL = "Zl_中联合理用药参数_Update(" & lng病人ID & "," & lng主页ID & "," & lng挂号ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, G_STR_PASS)
    '获取临时表数据:中联合理用药参数
    strXML = ReadLobForPASS()
    gcnOracle.CommitTrans: blnTran = False
    
    WriteLog "" & glngModel, "ZL_MakeDetailXML", "返回临时表XML:" & strXML
    strXML = "<root><场景ID_IN></场景ID_IN><医院名称_IN>" & mstrHosName & "</医院名称_IN>" & strXML & "</root>"
    WriteLog "" & glngModel, "ZL_MakeDetailXML", "传入审查接口XML:" & strXML
    ZL_MakeDetailXML = strXML
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ZL_GET_PatiInfo(ByVal colPati As Collection, ByVal byt场合 As Byte) As String
 
    Dim strXML As String
'    <patient_info>
'        <info name=“提交类型” value=“1”/> --1-新开，2-保存
'        <info name="年龄数字" value="28114.45"/> --出生日期-sysdate
'        <info name="年龄周期" value="成人"/>
'        <info name="性别" value="女"/>
'        <info name="职业" value="运动员"/>
'        <info name="妊娠" value="1"/>
'        <info name="哺乳" value="1"/>
'        <info name="肝功能不全" value="1">
'        <info name="严重肝功能不全" value="1">
'        <info name="肾功能不全" value="1">
'        <info name="严重肾功能不全" value="1">
'        <info name="诊断" value="J18.000"/> --诊断传编码，多个诊断以逗号分隔
'    </patient_info>
    strXML = _
    "<patient_info>" & vbNewLine & _
            "    <info name=""提交类型"" value=""" & colPati("提交类型") & """/>" & vbNewLine & _
            "    <info name=""年龄数字"" value=""" & colPati("年龄数字") & """  unit=""天""/>" & vbNewLine & _
            "    <info name=""年龄周期"" value=""" & colPati("年龄周期") & """/>" & vbNewLine & _
            "    <info name=""年龄"" value=""" & colPati("年龄") & """/>" & vbNewLine & _
            "    <info name=""身高"" value=""" & colPati("身高") & """/>" & vbNewLine & _
            "    <info name=""体重"" value=""" & colPati("体重") & """/>" & vbNewLine & _
            "    <info name=""性别"" value=""" & colPati("性别") & """/>" & vbNewLine & _
            "    <info name=""职业"" value=""" & colPati("职业") & """/>" & vbNewLine & _
            "    <info name=""妊娠"" value=""" & colPati("妊娠") & """/>" & vbNewLine & _
            "    <info name=""哺乳"" value=""" & colPati("哺乳") & """/>" & vbNewLine & _
            "    <info name=""肝功能不全"" value=""" & colPati("肝功能不全") & """/>" & vbNewLine & _
            "    <info name=""严重肝功能不全"" value=""" & colPati("严重肝功能不全") & """/>" & vbNewLine & _
            "    <info name=""肾功能不全"" value=""" & colPati("肾功能不全") & """/>" & vbNewLine & _
            "    <info name=""严重肾功能不全"" value=""" & colPati("严重肾功能不全") & """/>" & vbNewLine & _
            "    <info name=""诊断名称"" value=""" & colPati("诊断名称") & """/>" & vbNewLine & _
            "    <info name=""诊断"" value=""" & colPati("诊断") & """/>"
        If colPati("提交类型") = "2" Then
             strXML = strXML & vbNewLine & _
                "    <info name=""姓名"" value=""" & colPati("姓名") & """/>" & vbNewLine & _
                "    <info name=""出生日期"" value=""" & colPati("出生日期") & """/>" & vbNewLine & _
                "    <info name=""入院日期"" value=""" & colPati("入院日期") & """/>" & vbNewLine & _
                "    <info name=""当前床号"" value=""" & colPati("当前床号") & """/>" & vbNewLine & _
                "    <info name=""就诊科室"" value=""" & colPati("就诊科室") & """/>" & vbNewLine & _
                "    <info name=""就诊科室ID"" value=""" & colPati("就诊科室ID") & """/>" & vbNewLine & _
                "    <info name=""婴儿"" value=""" & colPati("婴儿") & """/>" & vbNewLine & _
               "    <info name=""HIS_NO"" value=""" & colPati("HIS_NO") & """/>" & vbNewLine
                If byt场合 = 1 Then
                    strXML = strXML & "    <info name=""门诊号"" value=""" & colPati("门诊号") & """/>" & vbNewLine & _
                                "    <info name=""病人来源"" value=""" & colPati("病人来源") & """/>" & vbNewLine
                    
                Else
                    strXML = strXML & "    <info name=""住院号"" value=""" & colPati("住院号") & """/>" & vbNewLine & _
                                    "    <info name=""就诊病区ID"" value=""" & colPati("就诊病区ID") & """/>" & vbNewLine & _
                                    "    <info name=""就诊病区"" value=""" & colPati("就诊病区") & """/>" & vbNewLine & _
                                    "    <info name=""病人来源"" value=""" & colPati("病人来源") & """/>" & vbNewLine
                End If
        End If
    strXML = strXML & "</patient_info>"
    ZL_GET_PatiInfo = strXML
End Function

Public Function ZL_GET_Medicine(ByVal rsAdvice As ADODB.Recordset, ByVal bytFunc As Byte) As String
    Dim strXML As String
    '功能:bytFunc=3 增加审查信息节点
    '<medicine_info>
    '  <medicine>
    '    <info name="医嘱ID" value="2"/>
    '    <info name="本位码" value="86903291000301" main="46d64420-8319-4768-9a11-f4b0f5e4ce7a"/>
    '    <info name="诊疗项目ID" value="67231" main="4e19df1c-c1b9-4a43-a83d-0741a19961ab"/>
    '    <info name="输液组号" value="1"/>
    '    <info name="计量单位" value="ml"/>
    '    <info name="单次量" value="60"/>
    '    <info name="单次量-按体重" value="1.25"/>  '单次用量除以病人体重
    '    <info name="单次量-按体表" value="40.87"/> '单次量-按体表=trunc(单次用量/(0.0061*病人身高+0.0128*病人体重-0.1529),2)
    '    <info name="每日量" value="60"/>
    '    <info name="每日量-按体重" value="1.25"/>
    '    <info name="每日量-按体表" value="40.87"/>
    '    <info name="给药频次" value="每日一次"/>
    '    <info name="给药途径" value="静脉输液"/>
    '  </medicine> 多个药品多个Medicine
    '</medicine_info>
    rsAdvice.Filter = "性质=0"
    strXML = "<medicine_info>"
    Do While Not rsAdvice.EOF
        strXML = strXML & _
        "  <medicine>" & vbNewLine & _
        "    <info name=""医嘱ID"" value=""" & rsAdvice!医嘱ID & """/>" & vbNewLine & _
        "    <info name=""本位码"" value=""" & rsAdvice!本位码 & """ main=""46d64420-8319-4768-9a11-f4b0f5e4ce7a""/>" & vbNewLine & _
        "    <info name=""诊疗项目ID"" value=""" & rsAdvice!诊疗项目ID & """ main=""4e19df1c-c1b9-4a43-a83d-0741a19961ab""/>" & vbNewLine & _
        "    <info name=""输液组号"" value=""" & rsAdvice!输液组号 & """/>" & vbNewLine & _
        "    <info name=""计量单位"" value=""" & rsAdvice!计量单位 & """/>" & vbNewLine & _
        "    <info name=""单次量"" value=""" & rsAdvice!单次量 & """/>" & vbNewLine & _
        "    <info name=""每日量"" value=""" & rsAdvice!每日量 & """/>" & vbNewLine & _
        "    <info name=""给药频次"" value=""" & rsAdvice!给药频次 & """/>" & vbNewLine & _
        "    <info name=""给药频次名称"" value=""" & rsAdvice!给药频次名称 & """/>" & vbNewLine & _
        "    <info name=""给药途径"" value=""" & rsAdvice!给药途径 & """/>" & vbNewLine
        If bytFunc = 3 Then
            strXML = strXML & _
                "    <info name=""给药途径名称"" value=""" & rsAdvice!给药途径名称 & """/>" & vbNewLine & _
                "    <info name=""医嘱期效"" value=""" & rsAdvice!医嘱期效 & """/>" & vbNewLine & _
                "    <info name=""开嘱时间"" value=""" & rsAdvice!开嘱时间 & """/>" & vbNewLine & _
                "    <info name=""医嘱状态"" value=""" & rsAdvice!医嘱状态 & """/>" & vbNewLine & _
                "    <info name=""紧急标志"" value=""" & rsAdvice!紧急标志 & """/>" & vbNewLine & _
                "    <info name=""开嘱医生"" value=""" & rsAdvice!开嘱医生 & """/>" & vbNewLine & _
                "    <info name=""医生职称"" value=""" & rsAdvice!医生职称 & """/>" & vbNewLine & _
                "    <info name=""医生抗菌药物等级"" value=""" & rsAdvice!医生抗菌药物等级 & """/>" & vbNewLine & _
                "    <info name=""医嘱内容"" value=""" & rsAdvice!医嘱内容 & """/>" & vbNewLine
            strXML = strXML & _
                "    <info name=""用药天数"" value=""" & rsAdvice!用药天数 & """/>" & vbNewLine & _
                "    <info name=""剂型"" value=""" & rsAdvice!剂型 & """/>" & vbNewLine & _
                "    <info name=""药品抗菌药物等级"" value=""" & rsAdvice!药品抗菌药物等级 & """/>" & vbNewLine & _
                "    <info name=""毒理分类"" value=""" & rsAdvice!毒理分类 & """/>" & vbNewLine & _
                "    <info name=""超量说明"" value=""" & rsAdvice!超量说明 & """/>" & vbNewLine & _
                "    <info name=""用药目的"" value=""" & rsAdvice!用药目的 & """/>" & vbNewLine & _
                "    <info name=""药品禁忌等级"" value=""" & rsAdvice!药品禁忌等级 & """/>" & vbNewLine & _
                "    <info name=""药品禁忌类型"" value=""" & rsAdvice!药品禁忌类型 & """/>" & vbNewLine & _
                "    <info name=""药品禁忌说明"" value=""" & rsAdvice!药品禁忌说明 & """/>" & vbNewLine & _
                "    <info name=""处方诊断"" value=""" & rsAdvice!处方诊断 & """/>" & vbNewLine
        End If
        strXML = strXML & "  </medicine>"
        rsAdvice.MoveNext
    Loop
    strXML = strXML & "</medicine_info>" & vbNewLine
    ZL_GET_Medicine = strXML
End Function

Public Function ZL_GET_Cusrules(ByVal rsAdvice As ADODB.Recordset) As String
    Dim strXML As String
 
    rsAdvice.Filter = "性质=1"
    strXML = strXML & "<cusrules><diatreat>" & vbNewLine
    Do While Not rsAdvice.EOF
        strXML = strXML & "<info name=""诊疗项目ID"" value=""" & rsAdvice!诊疗项目ID & """ main=""4e19df1c-c1b9-4a43-a83d-0741a19961ab""/>" & vbNewLine
        rsAdvice.MoveNext
    Loop
    strXML = strXML & "</diatreat></cusrules>"  '</diatreat></cusrules>替换时作为关键字
  
    ZL_GET_Cusrules = strXML
End Function

Public Function ZL_GET_Details(ByVal colPati As Collection, ByVal rsAdvice As ADODB.Recordset, _
    ByVal byt场合 As Byte, ByVal bytFunc As Byte) As String
    Dim strXML As String
'byt场合=1 门诊;2-住院
'<details_xml>
'    <patient_info>
'    </patient_info>
'    <medicine_info>
'    </medicine_info>
'</details_xml>
    strXML = "<details_xml>"
            strXML = strXML & ZL_GET_PatiInfo(colPati, byt场合)
            strXML = strXML & ZL_GET_Medicine(rsAdvice, bytFunc)
            If gstrIP <> "" And bytFunc = 0 Then strXML = strXML & ZL_GET_Cusrules(rsAdvice)
            strXML = strXML & "</details_xml>"
    ZL_GET_Details = strXML
End Function
 
Public Function ReadLobForPASS() As String
'功能：将指定的LOB字段复制为临时文件
'参数：
'返回：存放内容的文件名，失败则返回零长度""
    Dim rsLob As ADODB.Recordset
    Dim lngCount As Long
    Dim strText As String
    Dim strSQL As String
    Dim strFile As String
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select Zl_Read_中联合理用药参数([1]) as 片段 From Dual"
    lngCount = 0
    strFile = ""
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(strSQL, "ReadLobForPASS", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        strFile = strFile & strText
        lngCount = lngCount + 1
    Loop
     
    ReadLobForPASS = strFile
    Exit Function
Errhand:
    Err.Clear
End Function

Private Function Get每日量(ByVal str单次量 As String, ByVal str间隔单位 As String, ByVal int频率次数 As Integer, _
    ByVal int频率间隔 As Integer, ByVal str频率 As String) As String
'功能:
'1.每日量=单次量*日频次
'2.日频次计算：
'                    b.间隔单位=天 and 频率间隔=1，日频次=频率次数
'                    c.间隔单位=天 and 频率间隔>1 and 频率次数=1，日频次=1
'                    d.间隔单位=小时 and 频率间隔<=24,日频次=24/频率间隔*频率次数
'                    e.间隔单位=小时 and 频率间隔>24 and 频率次数=1，日频次=1
'                    f.间隔单位=周 and 频率次数=1，日频次=1
'                   str频率=一次性  日频次=1

    Dim str每日量 As String
    
    If str间隔单位 = "天" And int频率间隔 = 1 Then
        str每日量 = Val(str单次量) * int频率次数
    ElseIf str间隔单位 = "天" And int频率间隔 > 1 And int频率次数 = 1 Then
        str每日量 = Val(str单次量) * 1
    ElseIf str间隔单位 = "小时" And int频率间隔 <= 24 Then
        str每日量 = Val(str单次量) * (24 / int频率间隔 * int频率次数)
    ElseIf str间隔单位 = "小时" And int频率间隔 > 24 And int频率次数 = 1 Then
        str每日量 = Val(str单次量) * 1
    ElseIf str间隔单位 = "周" And int频率次数 = 1 Then
        str每日量 = Val(str单次量) * 1
    ElseIf str频率 = "一次性" Then
        str每日量 = Val(str单次量) * 1
    End If
    Get每日量 = FormatEx(str每日量, 2)
End Function

Private Function Get年龄周期(ByVal dbl年龄数字 As Double) As String
'功能:
'新生儿:<=28天
'婴儿:>28 天 ;<=365天
'幼儿:>365 天; <= 6岁
'儿童:>6岁;<=14岁
'少年:>14岁 and <=18
'成人:>18   and  <=60
'老人:>60岁

    Dim str年龄周期 As String
    
    If dbl年龄数字 <= 28 Then
        str年龄周期 = "新生儿"
    ElseIf dbl年龄数字 > 28 And dbl年龄数字 <= 365 Then
        str年龄周期 = "婴儿"
    ElseIf dbl年龄数字 > 365 And dbl年龄数字 <= (365 * 6) Then
        str年龄周期 = "幼儿"
    ElseIf dbl年龄数字 > (365 * 6) And dbl年龄数字 <= (365 * 14) Then
        str年龄周期 = "儿童"
    ElseIf dbl年龄数字 > (365 * 14) And dbl年龄数字 <= (365 * 18) Then
        str年龄周期 = "少年"
    ElseIf dbl年龄数字 > (365 * 18) And dbl年龄数字 <= (365 * 60) Then
        str年龄周期 = "成人"
    ElseIf dbl年龄数字 > (365 * 60) Then
        str年龄周期 = "老人"
    Else
        str年龄周期 = ""
    End If
    Get年龄周期 = str年龄周期
End Function

Private Function ZL_ParseXML(ByVal strData As String) As ADODB.Recordset
          '功能:解析XML字符串
          '返回XML:
          '1.  节点解释：
          '<order>：头尾节点
          '<orderid>：医嘱ID
          '<drugcode>:本位码
          '<type>：返回类型（文字信息），用于描述当前返回的属于哪一类问题
          '<level>：警示等级（文字信息），慎用/禁用/空
          '<describ>：内容描述，当前禁忌信息的主要描述
          '<remaks>：备注信息
          '<order><order_id>1</order_id> '相互作用、注射剂配伍、重复用药，这三类都是多个药一种提示的,故返回医嘱ID串,半角逗号分隔
          '<drugcode>86900967000160</drugcode>
          '<type>适应症</type><level></level><describ>【氯化钠注射液】只适合如下情况：静脉滴注、外用</describ>
          '<remaks>给药途径</remaks></order><order>
          '<order_id>2</order_id><drugcode>86903291000301</drugcode>
          '<type>适应症</type><level></level><describ>【康艾注射液】只适合如下情况：静脉滴注、静脉注射</describ>
          '<remaks>给药途径</remaks></order>
        Dim xmlDoc As New DOMDocument
        Dim xNode As IXMLDOMNode
        Dim xNodeList As IXMLDOMNodeList
        Dim rsRet As ADODB.Recordset
        Dim arrTemp As Variant
        Dim i As Long
          
10      On Error GoTo errH
20      Set rsRet = InitAdviceRS(FUN_审查结果_ZL)
        '读取网关响应数据（XML格式）
30      xmlDoc.loadXML (strData)
40      Set xNodeList = xmlDoc.selectNodes(".//order")
50      For Each xNode In xNodeList
60          arrTemp = Split(xNode.selectSingleNode(".//order_id").Text, ",")
70          For i = LBound(arrTemp) To UBound(arrTemp)
80              rsRet.AddNew
90              rsRet!OrderId = arrTemp(i)
100             rsRet!DrugCode = xNode.selectSingleNode(".//drugcode").Text
110             rsRet!Type = xNode.selectSingleNode(".//type").Text
120             If Mid(rsRet!Type & "", 1, Len("【问诊】")) = "【问诊】" Then
130                 rsRet!Category = 1
140             Else
150                 rsRet!Category = 0
160             End If
170             rsRet!Level = xNode.selectSingleNode(".//level").Text
180             rsRet!describ = xNode.selectSingleNode(".//describ").Text
190             rsRet!remaks = xNode.selectSingleNode(".//remaks").Text
200             If rsRet!Level = "禁止" Then
210                 rsRet!Light = 4
220                 rsRet!WarnLevel = 4   '警示级别排序
230             ElseIf rsRet!Level = "禁用" Then
240                 rsRet!Light = 2
250                 rsRet!WarnLevel = 3   '警示级别排序
260             ElseIf rsRet!Level = "注意" Then
270                 rsRet!Light = 3
280                 rsRet!WarnLevel = 1
290             ElseIf rsRet!Level = "慎用" Or rsRet!Level = "" Then
300                 rsRet!Light = 1
310                 rsRet!WarnLevel = 2
320             Else
330                 rsRet!Light = 0
340                 rsRet!WarnLevel = 0
350             End If
360             rsRet!Tag = IIf(i = LBound(arrTemp), 0, 1)    '1-标记重复内容
                '反向问诊去重
370             If rsRet!Category = 1 Then
380                 rsRet.Filter = "Category=1 And Type='" & rsRet!Type & "' And describ='" & rsRet!describ & "'"
390                 If rsRet.RecordCount > 1 Then rsRet.Delete
400             End If
410             rsRet.Update
420         Next
430     Next
440     Set ZL_ParseXML = rsRet
450     Exit Function
errH:
460     MsgBox Err.Description & vbCrLf & "ZL_ParseXML" & "行 " & Erl(), vbExclamation + vbOKOnly, gstrSysName
End Function

Private Function ZL_ParseXMLCusRules(ByVal strData As String) As ADODB.Recordset
    '功能:解析XML字符串
    '返回XML:
    '    "<cusrules>" & vbNewLine & _
    '    "  <result>" & vbNewLine & _
    '    "    <info name=""病生理情况"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""检查项目"" obsid=""fac26638-6d75"" default=""""/>" & vbNewLine & _
    '    "    <info name=""是否妊娠"" type=""平面单选"" index=""2"" value=""是|否"" class=""检查项目"" obsid=""fac26638-6d75"" default=""""/>" & vbNewLine & _
    '    "    <info name=""性别"" type=""平面单选"" index=""3"" value=""男|女"" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
    '    "    <info name=""过敏源"" type=""下拉多选"" index=""4"" value=""花粉|头孢|阿奇霉素|阿莫西林|阿司匹林"" class=""检查项目"" obsid=""fac26638-6d75""/>" & vbNewLine & _
    '    "    <info name=""年龄"" type=""下拉单项"" index=""5"" value=""婴儿|学前|儿童|少年|成年|中年|老年"" class=""检查项目"" obsid=""fac26638-6d75""/>" & vbNewLine & _
    '    "    <info name=""既往史描述"" type=""文本"" index=""475"" value="""" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
    '    "  </result>" & vbNewLine & _
    '    "</cusrules>"
        Dim xmlDoc As New DOMDocument
        Dim xNode As IXMLDOMNode
        Dim xNodeList As IXMLDOMNodeList
        Dim rsRet As ADODB.Recordset
        Dim strNodeValue As String
        
        Dim i As Long
    
        On Error GoTo errH
100     Set rsRet = InitAdviceRS(FUN_反向问诊_ZL)
        '读取网关响应数据（XML格式）
102     xmlDoc.loadXML (strData)
104     Set xNodeList = xmlDoc.selectNodes(".//cusrules/result/info")
106     For Each xNode In xNodeList
108             rsRet.AddNew
110             For i = 0 To xNode.Attributes.length - 1
112                 strNodeValue = xNode.Attributes(i).nodeValue
114                 Select Case xNode.Attributes(i).baseName
                    Case "name"
116                     rsRet!Name = strNodeValue
118                 Case "type"
120                     rsRet!Type = strNodeValue
122                 Case "index"
124                     rsRet!Index = strNodeValue
126                 Case "value"
128                     rsRet!Value = strNodeValue
130                 Case "default"
132                     rsRet!Default = strNodeValue
134                 Case "class"
136                     rsRet!Class = strNodeValue
138                 Case "obsid"
140                     rsRet!Obsid = strNodeValue
142                 Case "proid"
144                     rsRet!Proid = strNodeValue
                    End Select
                Next
146             rsRet.Filter = "Name='" & rsRet!Name & "' And Type ='" & rsRet!Type & "'"
148             If rsRet.RecordCount > 1 Then '去掉重复项目
150                 rsRet.Delete
152             Else
154                 rsRet.Update
                End If
        Next
156     rsRet.Filter = ""
         
158     Set ZL_ParseXMLCusRules = rsRet
        Exit Function
errH:
160     MsgBox Err.Description & vbCrLf & "ZL_ParseXMLCusRules" & "行 " & Erl(), vbExclamation + vbOKOnly, gstrSysName
End Function

Public Function GetTestXML(ByVal bytFunc As Byte, Optional ByRef rsAdvice As ADODB.Recordset, Optional ByRef strAsk As String) As String
    Dim strPar As String
    Dim i As Long

    If bytFunc = 0 Then
        strPar = strPar & "{""医院ID_IN"":1," & vbNewLine & _
                " ""药品医嘱XML_IN"":""<details_xml>" & vbNewLine & _
                "  <patient_info>" & vbNewLine & _
                "    <info name=\""年龄数字\"" value=\""28114.45\""/>" & vbNewLine & _
                "    <info name=\""年龄周期\"" value=\""成人\""/>" & vbNewLine & _
                "    <info name=\""性别\"" value=\""女\""/>" & vbNewLine & _
                "    <info name=\""职业\"" value=\""运动员\""/>" & vbNewLine & _
                "    <info name=\""妊娠\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""哺乳\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""肝功能不全\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""严重肝功能不全\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""肾功能不全\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""严重肾功能不全\"" value=\""1\""/>" & vbNewLine & _
                "    <info name=\""诊断\"" value=\""J18.000\""/>" & vbNewLine & _
                "  </patient_info>"
            strPar = strPar & "<medicine_info>" & vbNewLine & _
                "    <medicine>" & vbNewLine & _
                "      <info name=\""医嘱ID\"" value=\""1\""/>" & vbNewLine & _
                "      <info name=\""本位码\"" value=\""86900967000160\"" main=\""46d64420-8319-4768-9a11-f4b0f5e4ce7a\""/>" & vbNewLine & _
                "      <info name=\""诊疗项目ID\"" value=\""67232\"" main=\""4e19df1c-c1b9-4a43-a83d-0741a19961ab\""/>" & vbNewLine & _
                "      <info name=\""输液组号\"" value=\""1\""/>" & vbNewLine & _
                "      <info name=\""计量单位\"" value=\""ml\""/>" & vbNewLine & _
                "      <info name=\""单次量\"" value=\""250\""/>" & vbNewLine & _
                "      <info name=\""单次量-按体重\"" value=\""5.21\""/>" & vbNewLine & _
                "      <info name=\""单次量-按体表\"" value=\""170.3\""/>" & vbNewLine & _
                "      <info name=\""每日量\"" value=\""250\""/>" & vbNewLine & _
                "      <info name=\""每日量-按体重\"" value=\""5.21\""/>" & vbNewLine & _
                "      <info name=\""每日量-按体表\"" value=\""170.3\""/>" & vbNewLine & _
                "      <info name=\""给药频次\"" value=\""每日一次\""/>" & vbNewLine & _
                "      <info name=\""给药途径\"" value=\""静脉输液\""/>" & vbNewLine & _
                "    </medicine>"

            strPar = strPar & " <medicine>" & vbNewLine & _
        "      <info name=\""医嘱ID\"" value=\""2\""/>" & vbNewLine & _
        "      <info name=\""本位码\"" value=\""86903291000301\"" main=\""46d64420-8319-4768-9a11-f4b0f5e4ce7a\""/>" & vbNewLine & _
        "      <info name=\""诊疗项目ID\"" value=\""67231\"" main=\""4e19df1c-c1b9-4a43-a83d-0741a19961ab\""/>" & vbNewLine & _
        "      <info name=\""输液组号\"" value=\""1\""/>" & vbNewLine & _
        "      <info name=\""计量单位\"" value=\""ml\""/>" & vbNewLine & _
        "      <info name=\""单次量\"" value=\""60\""/>" & vbNewLine & _
        "      <info name=\""单次量-按体重\"" value=\""1.25\""/>" & vbNewLine & _
        "      <info name=\""单次量-按体表\"" value=\""40.87\""/>" & vbNewLine & _
        "      <info name=\""每日量\"" value=\""60\""/>" & vbNewLine & _
        "      <info name=\""每日量-按体重\"" value=\""1.25\""/>" & vbNewLine & _
        "      <info name=\""每日量-按体表\"" value=\""40.87\""/>" & vbNewLine & _
        "      <info name=\""给药频次\"" value=\""每日一次\""/>" & vbNewLine & _
        "      <info name=\""给药途径\"" value=\""静脉输液\""/>" & vbNewLine & _
        "    </medicine>" & vbNewLine & _
        "  </medicine_info>" & vbNewLine & _
        "</details_xml>""}"

    ElseIf bytFunc = 1 Then
        strPar = """[{\""通用名称\"":\""异烟肼片\"",\""商品名\"":null,\""汉语拼音\"":null,\""英文名称\"":\""\\n异烟肼片\\nIsoniazid Tablets\"",\""药物规格\"":\""0.1g\"",\""药物剂型\"":\""片剂\"",\""生产企业\"":\""杭州民生药业有限公司\""" & _
                    ",\""批准文号\"":\""国药准字H33021636\"",\""化学名称\"":\""4-吡啶甲酰肼\"",\""性状\"":\""本品为白色片或者类白色片\"",\""药理毒理\"":\""本品是一种具有杀菌作用的合成抗菌药，本品只对分枝杆菌，主要是生长繁殖期的细菌有效。其作用机制尚未阐明，可能抑制敏感细菌分枝菌酸（mycolicacid）的合成而使细胞壁破裂。\"","
        strPar = strPar & "\""药代动力学\"":\""本品口服后迅速自胃肠道吸收，并分布于全身组织和体液中，包括脑脊液、胸水、腹水、皮肤、肌肉、乳汁和干酪样组织。" & _
                "并可穿过胎盘屏障。蛋白结合率仅0~10%。口服1~2小时血药浓度可达峰值，但4~6小时后血药浓度根据患者的乙酰化快慢而不一，快乙酰化者，T1/2为0.5~1.6小时，" & _
                "慢乙酰化者为2~5小时，肝、肾功能损害者可能延长。代谢主要在肝脏中乙酰化而成无活性代谢产物，其中有的具有肝毒性。乙酰化的速率由遗传所决定。" & _
                "慢乙酰化者常有肝脏N-乙酰转移酶缺乏，未乙酰化的异烟肼可被部分结合。本品主要经肾排泄（约70%），在24小时内排出，大部分为无活性代谢物。" & _
                "快乙酰化者中93%以乙酰化型在尿液中排出，慢乙酰化者为63%。快乙酰化者尿液中7%的异烟肼呈游离或结合型，而慢乙酰化者则为37%。本品易通过血脑屏障，"

        strPar = strPar & "亦可从乳汁排出，少量可自唾液、痰液和粪便中排出。相当量的异烟肼可经血液透析与腹膜透析清除。\""," & _
                "\""适应症\"":\""1）异烟肼与其它抗结核药联合，适用于各型结核病的治疗，包括结核性脑膜炎以及其他分枝杆菌感染。" & _
                "（2）异烟肼单用适用于各型结核病的预防：①新近确诊为结核病患者的家庭成员或密切接触者；②结核菌素纯蛋白衍生物试验（PPD）" & _
                "强阳性同时胸部X射线检查符合非进行性结核病，痰菌阴性，过去未接受过正规抗结核治疗者；③正在接受免疫抑制剂或长期激素治疗的患者，" & _
                "某些血液病或网状内皮系统疾病（如白血病、霍奇金氏病）、糖尿病、尿毒症、矽肺或胃切除术等患者，其结核菌素纯蛋白衍生物试验呈阳性反应者" & _
                "；④35岁以下结核菌素纯蛋白衍生物试验阳性的患者；⑤已知或疑为HIV感染者，其结核菌素纯蛋白衍生物试验呈阳性反应者，或与活动性肺结核患者有密切接触者。\"","

                strPar = strPar & _
                "\""用法用量\"":\""口服：预防：成人一日0.3g，顿服；小儿每日按体重10mg/kg，一日总量不超过0.3g，顿服。治疗：成人与其他抗结核药合用，按体重每日口服5mg/kg，最高0.3g；或每日15mg/kg，最高900mg，每周2~3次。小儿按体重每日10~20mg/kg，每日不超过0.3g，顿服。某些严重结核病患儿（如结核性脑膜炎），每日按体重可高达30mg/kg（一日量最高500mg），但要注意肝功能损害和周围神经炎的发生。\""," & _
                "\""不良反应\"":\""发生率较多者有步态不稳或麻木针刺感、烧灼感或手指疼痛（周围神经炎）；深色尿、眼或皮肤黄染（肝毒性，35岁以上患者肝毒性发生率增高）；食欲不佳、异常乏力或软弱、恶心或呕吐（肝毒性的前驱症状）。发生率极少者有视力模糊或视力减退，合并或不合并眼痛（视神经炎）；发热、皮疹、血细胞减少及男性乳房发育等。本品偶可因神经毒性引起的抽搐。\""," & _
                "\""禁忌症\"":\""肝功能不正常者，精神病患者和癫痫病人禁用。\""," & _
                "\""注意事项\"":\""（1）交叉过敏反应，对乙硫异烟胺、吡嗪酰胺、烟酸或其他化学结构有关药物过敏者也可能对本品过敏。" & _
                "（2）对诊断的干扰：用硫酸铜法进行尿糖测定可呈假阳性反应，但不影响酶法测定的结果。异烟肼可使血清胆红素、丙氨酸氨基转移酶及门冬氨酸氨基转移酶的测定值增高。"

                strPar = strPar & "（3）有精神病、癫痫病史者、严重肾功能损害者应慎用。（4）如疗程中出现视神经炎症状，应立即进行眼部检查，并定期复查。" & _
                "（5）异烟肼中毒时可用大剂量维生素B6对抗。\""," & _
                "\""孕妇用药\"":\""（1）本品可穿过胎盘，导致胎儿血药浓度高于母血药浓度。动物实验证实异烟肼可引起死胎，但在人类中虽未证实，孕妇应用时必须充分权衡利弊。" & _
                "异烟肼与其他药物联合时对胎儿的作用尚未阐明。此外，在新生儿用药时应密切观察不良反应。（2）异烟肼在乳汁中浓度可达12mg/L,与血药浓度相近；" & _
                "虽然在人类中尚未证实有问题，哺乳期间应用仍应充分权衡利弊。如用药则宜停止哺乳。\"","

                strPar = strPar & _
                "\""儿童用药\"":\""严格按照儿童用法用量使用\""," & _
                "\""相互作用\"":\""1）服用异烟肼时每日饮酒，易引起本品诱发的肝脏毒性反应，并加速异烟肼的代谢，因此需调整异烟肼的剂量，并密切观察肝毒性征象。" & _
                "应劝告患者服药期间避免酒精饮料。（2）含铝制酸药可延缓并减少异烟肼口服后的吸收，使血药浓度减低，故应避免两者同时服用，或在口服制酸剂前至少1小时服用异烟肼。" & _
                "（3）抗凝血药（如香豆素或茚满双酮衍生物）与异烟肼同时应用时，由于抑制了抗凝药的酶代谢，使抗凝作用增强。" & _
                "（4）与环丝氨酸同服时可增加中枢神经系统不良反应（如头昏或嗜睡），需调整剂量，并密切观察中枢神经系统毒性征象，尤其对于从事需要灵敏度较高的工作的患者。"

                strPar = strPar & "（5）利福平与异烟肼合用时可增加肝毒性的危险性，尤其是已有肝功能损害者或为异烟肼快乙酰化者，因此在疗程的头3个月应密切随访有无肝毒性征象出现。" & _
                "（6）异烟肼为维生素B6的拮抗剂,可增加维生素B6经肾排出量,因而可能导致周围神经炎,服用异烟肼时维生素B6的需要量增加。" & _
                "（7）与肾上腺皮质激素(尤其泼尼松龙)合用时,可增加异烟肼在肝内的代谢及排泄,导致后者血药浓度减低而影响疗效,在快乙酰化者更为显著,应适当调整剂量。" & _
                "（8）与阿芬太尼（alfentanil）合用时，由于异烟肼为肝药酶抑制剂，可延长阿芬太尼的作用；与" & _
                "双硫仑(disulfiram)合用可增强其中枢神经系统作用，产生眩晕、动作不协调、易激惹、失眠等；与安氟醚合用可增加具有肾毒性的无机氟代谢物的形成。" & _
                "（9）与乙硫异烟胺或其他抗结核药合用，可加重后二者的不良反应。与其他肝毒性药合用可增加本品的肝毒性，因此宜尽量避免。（10）异烟肼不宜与酮康唑或咪康唑合用，" & _
                "因可使后两者的血药浓度降低。（11）与苯妥英钠或氨茶碱合用时可抑制二者在肝脏中的代谢，而导致苯妥英钠或氨茶碱血药浓度增高" & _
                "，故异烟肼与两者先后应用或合用时，苯妥英钠或氨茶碱的剂量应适当调整。（12）与对乙酰氨基酚合用时，由于异烟肼可诱导肝细胞色素P-450，" & _
                "使前者形成毒性代谢物的量增加，可增加肝毒性及肾毒性。（13）与卡马西平同时应用时，异烟肼可抑制其代谢，使卡马西平的血药浓度增高，" & _
                "而引起毒性反应；卡马西平可诱导异烟肼的微粒体代谢，形成具有肝毒性的中间代谢物增加。（14）本品不宜与其他神经毒药物合用，以免增加神经毒性。\""," & _
                "\""药物过量\"":\""未进行该项试验且无参考文献\""," & _
                "\""贮藏条件\"":\""遮光，密封，在干燥处保存。\""}]"""
    ElseIf bytFunc = 2 Then
        strPar = "<details_xml>"
        strPar = strPar & "<order><order_id>1</order_id><drugcode>86900967000160</drugcode>" & _
                            "<type>适应症</type><level>禁用</level><describ>【氯化钠注射液】只适合如下情况：静脉滴注、外用</describ>" & _
                            "<remaks>给药途径</remaks></order>" & _
                            "<order><order_id>2</order_id><drugcode>86903291000301</drugcode>" & _
                            "<type>适应症</type><level>禁用</level><describ>【康艾注射液】只适合如下情况：静脉滴注、静脉注射 【氯化钠注射液】只适合如下情况：静脉滴注、外用【氯化钠注射液】只适合如下情况：静脉滴注、外用【康艾注射液】只适合如下情况：静脉滴注、静脉注射 【氯化钠注射液】只适合如下情况：静脉滴注、外用【氯化钠注射液】只适合如下情况：静脉滴注、外用</describ>" & _
                            "<remaks>给药途径</remaks></order>"
        strPar = strPar & "<order><order_id>3</order_id><drugcode>86900967000160</drugcode>" & _
                            "<type>适应症</type><level>慎用</level><describ>【氯化钠注射液】只适合如下情况：静脉滴注、外用</describ>" & _
                            "<remaks>给药途径</remaks></order>" & _
                            "<order><order_id>4</order_id><drugcode>86903291000301</drugcode>" & _
                            "<type>适应症</type><level>慎用</level><describ>【康艾注射液】只适合如下情况：静脉滴注、静脉注射</describ>" & _
                            "<remaks>给药途径</remaks></order>"
        strPar = strPar & "<order><order_id>5</order_id><drugcode>86900967000160</drugcode>" & _
                            "<type>适应症</type><level></level><describ>【氯化钠注射液】只适合如下情况：静脉滴注、外用</describ>" & _
                            "<remaks>给药途径</remaks></order>" & _
                            "<order><order_id>6</order_id><drugcode>86903291000301</drugcode>" & _
                            "<type>适应症</type><level></level><describ>【康艾注射液】只适合如下情况：静脉滴注、静脉注射</describ>" & _
                            "<remaks>给药途径</remaks></order>"
        strPar = strPar & "<order><order_id>7,8</order_id><drugcode>86900967000160</drugcode>" & _
                    "<type>药品相互作用</type><level>禁用</level><describ>【依诺肝素钠注射液】和【维生素C注射液】有相互作用：" & vbCrLf & "大剂量维生素C可干扰抗凝药的抗凝效果。</describ>" & _
                    "<remaks>给药途径</remaks></order>" & _
                    "<order><order_id>9,10</order_id><drugcode>86903291000301</drugcode>" & _
                    "<type>药品相互作用</type><level>禁用</level><describ>【康艾注射液】只适合如下情况：静脉滴注、静脉注射</describ>" & _
                    "<remaks>给药途径</remaks></order>"
        strPar = strPar & "</details_xml>"



        strPar = " <details_xml><order><order_id>1211848</order_id><drugcode>86900002000018</drugcode><type>给药途径</type>" & vbNewLine & _
            " <level></level><describ>【阿卡波糖片】只适合如下情况：嚼服、吞服</describ>" & vbNewLine & _
            " <remaks>给药途径</remaks></order><order><order_id>1211850,1211852</order_id>" & vbNewLine & _
            " <drugcode></drugcode><type>有药品相互作用</type><level>慎用</level>" & vbNewLine & _
            " <describ>【氨茶碱注射液】和【阿司匹林肠溶片】有相互作用：包括选择性环氧化酶－2抑制剂（COX－2抑制剂）" & vbNewLine & _
            " 在内的非甾体类抗炎药（NSAIDs）会降低利尿剂或其他抗高血压药物的效果。因此，血管紧张素Ⅱ受体拮抗剂的作" & vbNewLine & _
            " 用也会被包括选择性COX－2抑制剂在内的NSAIDs类药物减弱。</describ><remaks></remaks></order><order>" & vbNewLine & _
            " <order_id>1211848,1211850</order_id><drugcode></drugcode><type>疑似重复用药</type><level>慎用</level>" & vbNewLine & _
            " <describ>【阿卡波糖片、阿司匹林肠溶片】同属于【有降低血糖作用的药品】，疑似重复用药。</describ><remaks></remaks>" & vbNewLine & _
            " </order><order><order_id>1211850,1211852,1211854</order_id><drugcode></drugcode><type>疑似重复用药</type>" & vbNewLine & _
            " <level>慎用</level><describ>【阿司匹林肠溶片、氨茶碱注射液、注射用尿激酶】同属于【心" & vbNewLine & _
            "管药】，疑似重复用药。</describ><remaks></remaks></order>"
        
        strPar = strPar & "<order><order_id>1211850,1211854</order_id><drugcode></drugcode><type>疑似重复用药</type>" & _
            "<level>慎用</level><describ>【阿司匹林肠溶片、注射用尿激酶】同属于【影响止血的药物】，疑似重复用药。</describ>" & _
            "<remaks></remaks></order><order><order_id>1211854</order_id><drugcode>86901576000121</drugcode><type>给药途径</type><level></level><describ>【注射用尿激酶】只适合如下情况：冲洗、动脉灌注、静脉滴注</describ><remaks>给药途径</remaks></order><order><order_id>1211848</order_id><drugcode>86900002000018</drugcode><type>禁忌症</type><level>慎用</level><describ>【阿卡波糖片】年龄小于18年慎用</describ><remaks></remaks></order><order><order_id>1211854,1211850</order_id><drugcode></drugcode><type>有药品相互作用</type><level>慎用</level><describ>【阿司匹林肠溶片】和【注射用尿激酶】有相互作用：NSAIDs抑制血小板聚集和损害胃肠道粘膜，可增加抗凝药物的活性，增加使用抗凝药的病人胃肠道出血的风险。除非可以进行密切的监测，醋氯芬酸应避免与香豆素类口服抗凝血药、噻氯匹定、血栓溶解" & _
            "剂及肝素合用。</describ><remaks></remaks></order><order><order_id>1211848,1211850</order_id><drugcode></drugcode><type>有药品相互作用</type><level>慎用</level><describ>【阿司匹林肠溶片】和【阿卡波糖片】有相互作用：抗糖尿病药，例如胰岛素、磺酰脲类：" & vbNewLine & _
            "高剂量阿司匹林具有降血糖作用而增强降糖效果，并且能与磺酰脲类竞争结合血浆蛋白。</describ><remaks></remaks></order><order><order_id>1211850,1211852</order_id><drugcode></drugcode><type>疑似重复用药</type><level>慎用</level><describ>【阿司匹林肠溶片、氨茶碱注射液】同属于【冠状动脉扩张药】，疑似重复用药。</describ>" & _
            "<remaks></remaks></order><order><order_id>1211850</order_id><drugcode>86979489000088</drugcode>" & _
            "<type>给药途径</type><level></level><describ>【阿司匹林肠溶片】只适合如下情况：口服给药</describ>" & _
            "<remaks>给药途径</remaks></order><order><order_id>1211848,1211850</order_id><drugcode></drugcode>" & _
            "<type>疑似重复用药</type><level>慎用</level><describ>【阿卡波糖片、阿司匹林肠溶片】同属于【与血浆蛋白高度结合的药物】，疑似重复用药。</describ><remaks></remaks></order>" & _
            "<order><order_id>1211850,1211854</order_id><drugcode></drugcode><type>疑似重复用药</type><level>慎用</level><describ>【阿司匹林肠溶片、注射用尿激酶】同属于【肝脏毒性药物】，疑似重复用药。</describ>" & _
            "<remaks></remaks></order><order><order_id>1211850,1211854</order_id><drugcode></drugcode>" & _
            "<type>疑似重复用药</type><level>慎用</level><describ>【阿司匹林肠溶片、注射用尿激酶】同属于【血液学药物】，疑似重复用药。</describ><remaks></remaks></order>" & _
            "<order><order_id>1211854,1211852</order_id><drugcode></drugcode><type>有注射剂配伍禁忌</type><level>禁用</level><describ>【注射用尿激酶】【氨茶碱注射液】忌配伍：两药混合后出现理化、药理、药动学及药效学等方面配伍禁忌。</describ><remaks></remaks></order></details_xml>"

        '用于模拟警示灯设置
        rsAdvice.Filter = ""
        For i = 1 To rsAdvice.RecordCount
            strPar = Replace(strPar, "<order_id>" & i & "</order_id>", "<order_id>" & rsAdvice!医嘱ID & "</order_id>")
            If i = 6 Then Exit For
            rsAdvice.MoveNext
        Next
    ElseIf bytFunc = 3 Then
        '反向问诊支持类型:平面多选、平面单选、文本、下拉多选、下拉单项
        strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""检查项目"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""是否妊娠adsfasdfasdf考高阿斯顿发生代发阿斯蒂芬发送打算打算地方阿斯蒂芬发送地方as地方阿斯蒂芬阿斯蒂芬阿斯蒂芬发送到法傻傻的法"" type=""平面单选"" index=""2"" value=""是|否""  class=""检查项目"" obsid=""fac26638-6d75"" default=""是""/>" & vbNewLine & _
                "    <info name=""性别"" type=""平面单选"" index=""3"" value=""男|女""  class=""检查项目"" obsid=""fac26638-6d75"" default=""男""/>" & vbNewLine & _
                "    <info name=""过敏源"" type=""下拉多选"" index=""4"" value=""花粉|头孢|阿奇霉素|阿莫西林|阿司匹林""  class=""检查项目"" obsid=""fac26638-6d75"" default=""阿奇霉素""/>" & vbNewLine & _
                "    <info name=""年龄"" type=""下拉单项"" index=""5"" value=""婴儿|学前|儿童|少年|成年|中年|老年""  class=""检查项目"" obsid=""fac26638-6d75"" default=""学前""/>" & vbNewLine & _
                "    <info name=""既往史描述"" type=""文本"" index=""475"" value="""" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""病生理情况"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全|妊娠""  class=""检查项目"" obsid=""fac26638-6d75"" default=""""/>" & vbNewLine & _
                "    <info name=""是否妊娠"" type=""平面单选"" index=""2"" value=""是|否""  class=""检查项目"" obsid=""fac26638-6d75"" default=""是""/>" & vbNewLine & _
                "    <info name=""性别"" type=""平面单选"" index=""3"" value=""男|女"" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""过敏源"" type=""下拉多选"" index=""4"" value=""花粉|头孢|阿奇霉素|阿莫西林|阿司匹林""  class=""检查项目"" obsid=""fac26638-6d75"" default=""阿奇霉素""/>" & vbNewLine & _
                "    <info name=""年龄"" type=""下拉单项"" index=""5"" value=""婴儿|学前|儿童|少年|成年|中年|老年"" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""既往史描述"" type=""文本"" index=""475"" value="""" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""病生理情况"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全|妊娠""  class=""检查项目"" obsid=""fac26638-6d75"" default=""""/>" & vbNewLine & _
                "    <info name=""是否妊娠"" type=""平面单选"" index=""2"" value=""是|否""  class=""检查项目"" obsid=""fac26638-6d75"" default=""是""/>" & vbNewLine & _
                "    <info name=""性别"" type=""平面单选"" index=""3"" value=""男|女"" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""过敏源"" type=""下拉多选"" index=""4"" value=""花粉|头孢|阿奇霉素|阿莫西林|阿司匹林""  class=""检查项目"" obsid=""fac26638-6d75"" default=""阿奇霉素""/>" & vbNewLine & _
                "    <info name=""年龄"" type=""下拉单项"" index=""5"" value=""婴儿|学前|儿童|少年|成年|中年|老年"" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""既往史描述"" type=""文本"" index=""475"" value="""" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "  </result>" & vbNewLine & _
                "</cusrules>"
                
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""是否妊娠adsfasdfasdf考高阿斯顿发生代发阿斯蒂芬发送打算打算地方阿斯蒂芬发送地方as地方阿斯蒂芬阿斯蒂芬阿斯蒂芬发送到法傻傻的法"" type=""平面单选"" index=""2"" value=""是|否""  class=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬AAAAAAAAAAA"" obsid=""fac26638-6d75"" default=""是""/>" & vbNewLine & _
                "    <info name=""性别"" type=""平面单选"" index=""3"" value=""男|女""  class=""检查项目"" obsid=""fac26638-6d75"" default=""男""/>" & vbNewLine & _
                "    <info name=""过敏源"" type=""下拉多选"" index=""4"" value=""花粉|头孢|阿奇霉素|阿莫西林|阿司匹林""  class=""检查项目"" obsid=""fac26638-6d75"" default=""阿奇霉素""/>" & vbNewLine & _
                "    <info name=""年龄"" type=""下拉单项"" index=""5"" value=""婴儿|学前|儿童|少年|成年|中年|老年""  class=""检查项目"" obsid=""fac26638-6d75"" default=""学前""/>" & vbNewLine & _
                "    <info name=""既往史描述"" type=""文本"" index=""475"" value="""" class=""检查项目"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
            '平面多选
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全|妊娠|出生日期|严重肝功不全|严重肾功不全|哺乳"" class=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬asdfasdfasdfasdfasdf"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全XXXX|妊娠|出生日期asdfasdf|严重肝功不全asdfasdf|严重肾功不全|哺乳asdfasf"" class=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""AAAAA"" obsid=""fac26638-6d75"" default=""妊娠|肾功不全""/>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""平面多选"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""AAAA"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
            '下拉单项
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬长度超过自动换行"" type=""下拉单项"" index=""1"" value=""肝功不全|肾功不全|妊娠|出生日期|严重肝功不全|严重肾功不全|哺乳"" class=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""下拉单项"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""AAAA"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""下拉单项"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""AAAA"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
            '下拉多选
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬长度超过自动换行"" type=""下拉多选"" index=""1"" value=""肝功不全|肾功不全|妊娠|出生日期|严重肝功不全|严重肾功不全|哺乳"" class=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""下拉多选"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""AAAA"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""下拉多选"" index=""1"" value=""肝功不全|肾功不全|妊娠"" class=""AAAA"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
            '文本
            strPar = "<cusrules>" & vbNewLine & _
                "  <result>" & vbNewLine & _
                "    <info name=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬长度超过自动换行"" type=""文本"" index=""1"" value="""" class=""阿斯顿发生代发的发送到发送到发送到发送到发送到发送到发送到发送到发送到发送到发是的发送到发斯蒂芬"" obsid=""fac26638-6d75"" default=""妊娠""/>" & vbNewLine & _
                "    <info name=""病生理情况"" type=""文本"" index=""1"" value="""" class=""AAAA"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "    <info name=""病生理情况"" type=""文本"" index=""1"" value="""" class=""AAAA"" obsid=""fac26638-6d75"" />" & vbNewLine & _
                "  </result>" & vbNewLine & _
                      "</cusrules>"
        strAsk = strPar
    ElseIf bytFunc = 4 Then
        '审核通过
        strPar = "{""recipes"":[]}"
        '审核未通过
        strPar = "{""recipes"":[{""ORDER_ID"":2296,""NO_PASS_REASON"":""未通过""},{""ORDER_ID"":1924,""NO_PASS_REASON"":""未通过""},{""ORDER_ID"":2202,""NO_PASS_REASON"":""未通过""}]}"
    End If
    GetTestXML = strPar
End Function

Public Sub SetFormTranslucency(hWnd As Long, crKey As Long, bAlpha As Byte, dwFlags As Long) '实现半透明窗体
'功能:设置窗体透明度
'hwnd,  窗口句柄
'crKey:指定需要透明的背景颜色值，可用RGB()宏
'bAlpha:设置透明度，0表示完全透明，255表示不透明
'dwFlags: 透明方式dwFlags参数可取以下值：
'       LWA_ALPHA=&H2时：crKey参数无效，bAlpha参数有效；
'       LWA_COLORKEY=&H1：窗体中的所有颜色为crKey的地方将变为透明，bAlpha参数无效。其常量值为1。
'       LWA_ALPHA | LWA_COLORKEY：crKey的地方将变为全透明，而其它地方根据bAlpha参数确定透明度。
   Dim lngRet As Long
   
    lngRet = GetWindowLong(hWnd, GWL_EXSTYLE)
    lngRet = lngRet Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, lngRet
    SetLayeredWindowAttributes hWnd, crKey, bAlpha, dwFlags
 End Sub

Public Sub GetDrugInstructions(objfrmMain As Object, ByRef frmDrug As frmPassDrug, ByVal bytStyle As Byte, _
        ByVal strDrugCode As String, Optional ByVal strDrugName As String, Optional ByVal blnTip As Boolean)
'功能:药品说明书
    Dim strRet As String
    
    If gblnBreak Then Exit Sub
    
    If frmDrug Is Nothing Then Set frmDrug = New frmPassDrug
    If strDrugCode <> "" Then
        If Not GetDrugInfo_ZL(strDrugCode, strRet) Then Exit Sub
    Else
        If blnTip Then Exit Sub
        strRet = """[{\""通用名称\"":\""" & strDrugName & "\"",\""商品名\"":null,\""汉语拼音\"":null,\""英文名称\"":null,\""药物规格\"":null,\""药物剂型\"":null,\""生产企业\"":null" & _
                        ",\""批准文号\"":null,\""化学名称\"":null,\""性状\"":null,\""药理毒理\"":null,\""药代动力学\"":null," & _
                    "\""适应症\"":null,\""用法用量\"":null,\""不良反应\"":null,\""禁忌症\"":null," & _
                    "\""注意事项\"":null,\""孕妇用药\"":null,\""儿童用药\"":null," & _
                    "\""相互作用\"":null,\""药物过量\"":null,\""贮藏条件\"":null}]"""
    End If
    If bytStyle = 1 Then
        Call gobjFrm.CloseGetDrugInstructions
    End If
 
    frmDrug.ShowMe objfrmMain, strRet, bytStyle, blnTip
    
End Sub

Public Function GetDrugInfo_ZL(ByVal strDrugCode As String, ByRef strDrugInfo As String) As Boolean
    Dim strUrl As String
    Dim strRet As String
    
    strUrl = "http://" & gstrDrugIP & ":" & gstrDrugPort & "/api/DrugInstructions/" & strDrugCode
    strRet = HttpGet(strUrl, responseText, , gblnBreak)
    WriteLog "" & glngModel, "GetDrugInstructions", "说明书URL:" & strUrl & ",返回值:" & strRet
    If strRet = "" Or gblnBreak Then
        Call gobjAir.OpenTransparentAirBubble(gobjFrm, "合理用药监测服务器异常断开！", 2, 3, 0, vbWhite, vbRed, , 3, , , 咳嗽, True)
        gobjFrm.SetNotifyIcon
        gsngCheckLinkTime = Timer
        Exit Function
    End If
    If InStr(strRet, "errormsg") > 0 Then
        strRet = Replace(strRet, """{", "{")
        strRet = Replace(strRet, "}""", "}")
        strRet = Replace(strRet, "\""", """")
        strUrl = JSONParse("errormsg", strRet)
        If strUrl <> "" Then
            'MsgBox "药品说明书:" & vbCrLf & strURL, vbInformation + vbOKOnly, gstrSysName
            Call gobjAir.OpenTransparentAirBubble(gobjFrm, "药品说明书:" & strUrl, 2, 3, 0, vbWhite, vbRed, , 3, , 3000, 咳嗽, True)
            Exit Function
        End If
    End If
    strDrugInfo = strRet
    GetDrugInfo_ZL = True
End Function

Private Function GetDrugPlus(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByVal strAdvice As String) As ADODB.Recordset
    Dim strSQL As String
    Dim strPati As String
    
    On Error GoTo errH
    If str挂号单 <> "" Then
        strPati = " And A.挂号单 = [3] "
    Else
        strPati = " And A.病人ID =[1] And A.主页ID = [2] "
    End If
    strSQL = "Select a.Id, a.相关id, a.处方序号, d.名称 As 给药途径名称, D.ID as 给药途径ID" & vbNewLine & _
            " , a.超量说明,E.药品剂型,E.毒理分类,E.抗生素 " & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱记录 B, 诊疗项目目录 D,药品特性 E " & vbNewLine & _
            "Where a.相关id = b.Id(+) And b.诊疗项目id = d.Id(+) And A.诊疗项目ID =E.药名ID(+) " & strPati & " And a.相关id <> 0 And Instr([4], ',' || a.Id || ',') > 0" & vbNewLine & _
            "Order By a.序号"
    Set GetDrugPlus = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_ZL", lng病人ID, lng主页ID, str挂号单, "," & strAdvice & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDoctorInfo(ByVal strDoctorName As String, ByVal byt场合 As Byte) As ADODB.Recordset
    Dim strSQL As String
    'byt场合-1住院;2-门诊
    '聘任技术职务 "1.正高", "2.副高", "3.中级", "4.助理/师级", "5.员/士", "9.待聘"
    On Error GoTo errH
    If InStr(strDoctorName, ",") > 0 Then
        strSQL = "Select a.姓名, a.专业技术职务, Decode(a.聘任技术职务, 1, '正高', 2, '副高', 3, '中级', 4, '助理/师级', 5, '员/士', 9, '待聘') As 聘任技术职务, b.级别 " & vbNewLine & _
                "From 人员表 A, 人员抗菌药物权限 B" & vbNewLine & _
                "Where a.Id = b.人员id(+) And a.姓名 In (Select /*+cardinality(C,10)*/" & vbNewLine & _
                "                                  Column_Value" & vbNewLine & _
                "                                 From Table(f_Str2list([1])) C) And b.记录状态(+) = 1 And b.场合(+)=[2]"
    Else
        strSQL = "Select a.姓名, a.专业技术职务, Decode(a.聘任技术职务, 1, '正高', 2, '副高', 3, '中级', 4, '助理/师级', 5, '员/士', 9, '待聘') As 聘任技术职务, b.级别, b.场合" & vbNewLine & _
                "From 人员表 A, 人员抗菌药物权限 B" & vbNewLine & _
                "Where a.Id = b.人员id(+) And a.姓名 =[1] And b.记录状态(+) = 1  And b.场合(+)=[2]"
    End If

    Set GetDoctorInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_ZL", strDoctorName, byt场合)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOutAdviceDiagsInfo(ByVal strAdviceIDs As String) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
     strSQL = "Select a.医嘱id, b.诊断描述" & vbNewLine & _
                "From 病人诊断医嘱 A, 病人诊断记录 B" & vbNewLine & _
                "Where a.诊断id = b.Id And a.医嘱id In (Select /*+cardinality(C,10)*/" & vbNewLine & _
                "                                    Column_Value" & vbNewLine & _
                "                                   From Table(f_Num2list([1])) C)"


    Set GetOutAdviceDiagsInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_ZL", strAdviceIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetParaURL(ByVal strSysName As String, ByVal strServiceName As String) As String
'功能:
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strUrl As String
    On Error GoTo errH
    strSQL = "Select 服务地址 From 三方服务配置目录 Where 系统标识 = [1] And 服务名称 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPassDefine_ZL", strSysName, strServiceName)
    If Not rsTmp.EOF Then strUrl = Trim(rsTmp!服务地址 & "")
    GetParaURL = strUrl
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



'------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------中联药师审方系统接口--------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetReviewResult(ByVal lngPatiID As Long, ByVal lngVisitID As Long, Optional ByRef rsRet As ADODB.Recordset, _
    Optional ByRef strAdvice As String = "0", Optional ByRef strAdviceID As String = "0") As Boolean
    '功能:调服务查询
    '     审方结果查询
    '     固定部份 http://192.168.0.231:8080//ords/zlrecipe/recipe/result
    '     参数部份 ?pid=20800808&pvid=1
    '     strURL = "http://192.168.0.231:8080//ords/zlrecipe/recipe/result?pid=20800808&pvid=1";
    '     返回值:{""recipes"":[{""ORDER_ID"":123,""ORDER_GROUP_ID"":321,""NO_PASS_REASON"":""未通过""}]}

        Dim strUrl As String
        Dim strRet As String
        Dim strUnPass As String
        Dim strID As String
        
        Dim lngLength As Long
        Dim i As Long
        
        On Error GoTo errH
100     strUrl = GetParaURL("药师处方审查", "审查结果查询")
102     If strUrl = "" Then Exit Function
104     Set rsRet = InitAdviceRS(FUN_药师审查_ZL)
106     strUrl = strUrl & "?pid=" & lngPatiID & "&pvid=" & lngVisitID
108     WriteLog "" & glngModel, "GetReviewResult", "药师审查查询URL:" & strUrl
110     strRet = HttpGet(strUrl, responseText, 1)
112     WriteLog "" & glngModel, "GetReviewResult", "药师审查查询结果:" & strRet
114     If strRet <> "" Then
116         lngLength = JSONParse("recipes.length", strRet)
118         If lngLength > 0 Then
120             For i = 0 To lngLength - 1
122                 rsRet.AddNew
124                 rsRet!医嘱ID = JSONParse("recipes[" & i & "].ORDER_ID", strRet)
126                 rsRet!相关ID = JSONParse("recipes[" & i & "].ORDER_GROUP_ID", strRet)
128                 If strAdvice <> "0" Then
130                     If InStr("," & strUnPass & ",", "," & rsRet!相关ID & ",") = 0 Then
132                         strUnPass = strUnPass & "," & rsRet!相关ID
                        End If
                    End If
134                 If strAdviceID <> "0" Then
136                     If InStr("," & strID & ",", "," & rsRet!医嘱ID & ",") = 0 Then
138                         strID = strID & "," & rsRet!医嘱ID
                        End If
                    End If
140                 rsRet!审查详情 = JSONParse("recipes[" & i & "].NO_PASS_REASON", strRet)
142                 rsRet.Update
                Next
144             If strUnPass <> "" Then strAdvice = Mid(strUnPass, 2)
146             If strID <> "" Then strAdviceID = Mid(strID, 2)
            End If
        End If
148     GetReviewResult = True
150     If rsRet.RecordCount > 0 Then rsRet.MoveFirst
        Exit Function
errH:
152     MsgBox Err.Description & vbCrLf & "GetReviewResult" & "行 " & Erl(), vbExclamation + vbOKOnly, gstrSysName
End Function

Public Function EditPatiStatus() As Boolean
'功能:编辑病人状态
'参数:
    Dim strPath         As String
    Dim strUrl          As String
    Dim strPENVR       As String
    Dim strPvid         As String
    Dim bytPType         As Byte
    Dim i As Long
    
    If gstrStatusEdit = "" Then Exit Function
    If glngModel = PM_门诊医嘱清单 Then
        strPENVR = "10"
    ElseIf glngModel = PM_住院医嘱清单 Then
        strPENVR = "11"
    End If
    
    If gobjPati.str挂号单 <> "" Then
        strPvid = gobjPati.str挂号单
        bytPType = 1
    Else
        strPvid = gobjPati.lng主页ID & ""
        bytPType = 2
    End If
    strPath = Replace(UCase(App.Path), UCase("\Public"), "") & "\ZTHL"
    If Dir(strPath & "\nw.exe") <> "" Then
        strUrl = gstrStatusEdit & "?p=113:1:::::P1_ENVR_IN,P1_PID_IN,P1_PVID_IN,P1_RECORDER_IN,P1_RECORDER_ID_IN,P1_VISIT_TYPE_IN:" & _
                strPENVR & "," & gobjPati.lng病人ID & "," & strPvid & "," & zlStr.Base64Encode(UserInfo.姓名) & "," & UserInfo.id & "," & bytPType
        Shell "cmd /c rd " & strPath & "\userData /s/q"
        WriteLog "" & glngModel, "EditPatiStatus", "病人状态编辑URL:" & strUrl & ",文件路径:" & strPath
        i = ShellExecute(0, "open", "nw.exe", strUrl, strPath, SW_SHOWMAXIMIZED)
    Else
        MsgBox "病人状态程序文件（" & strPath & "\nw.exe）不存在。" & vbCrLf & "请检查或联系医院系统管理员。", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    EditPatiStatus = True
End Function

Public Function GetXMLResult(ByVal rsRec As ADODB.Recordset)
'功能:构造反向问诊响应XML
    Dim i As Long
    Dim strXML As String
    For i = 1 To rsRec.RecordCount
        strXML = strXML & "    <info name=""" & rsRec!Name & """ type=""" & rsRec!Type & """ index=""" & _
            rsRec!Index & """ value=""" & rsRec!Default & """ obsid=""" & rsRec!Obsid & """/>" & vbNewLine
        rsRec.MoveNext
    Next
    GetXMLResult = Replace(strXML, """", "\""")
End Function

Private Sub AskPatiStatus(ByVal rsAdvice As ADODB.Recordset, ByRef strPara As String, ByVal lngPatiID As Long)
      '功能:当病人状态已经问诊过,反向问诊时需排开已经问诊过内容
      '返回值:strPara 返回反向问诊内容

          Dim strUrl      As String
          Dim strData     As String
          Dim strRet      As String
          Dim strResult   As String
          Dim rsStatus    As ADODB.Recordset
          Dim rsAsk       As ADODB.Recordset
          Dim lngLength   As Long
          Dim i           As Long
          '测试地址http://192.168.32.201:8888/bizdomain/6f73f15d-3718-4570-8cea-cf6282a6f6f6
10       On Error GoTo errH

20        strData = "{""医嘱下达XML_IN"":""" & Replace(ZL_GET_Cusrules(rsAdvice), """", "\""") & """}"
30        WriteLog "" & glngModel, "AskPatiStatus", "反向问诊URL:" & gstrIP & ",反向问诊XML:" & strData
40        strRet = HttpPost(gstrIP, strData, responseText, , "Basic " & zlStr.Base64Encode("xxx:xxx"), , gblnBreak)
50        WriteLog "" & glngModel, "AskPatiStatus", "反向问诊结果:" & strRet
60        If strRet <> "" Then
              '测试用例构建
              'Call GetTestXML(3, rsAdvice, strRet)
70            Set rsAsk = ZL_ParseXMLCusRules(strRet)
80            If gstrStatusGet <> "" Then
                  'strURL = "http://192.168.0.231:8080/ords/patstatus/pat/getpatstatus?pati_id_in=4989"
90                strUrl = gstrStatusGet & "?pati_id_in=" & lngPatiID
100               strRet = HttpGet(strUrl, responseText, 1)
110               WriteLog "" & glngModel, "AskPatiStatus", "病人状态查询URL:" & strUrl & vbCrLf & _
                                                            "病人状态查询结果:" & strRet
120               If strRet <> "" Then
130                   lngLength = JSONParse("patient_status.length", strRet)
140                   If lngLength > 0 Then
150                       Set rsStatus = InitAdviceRS(FUN_病人状态_ZL)
160                       For i = 0 To lngLength - 1
170                           rsStatus.AddNew
180                           rsStatus!STATUS_ID = JSONParse("patient_status[" & i & "].STATUS_ID", strRet)
190                           rsStatus!status_name = JSONParse("patient_status[" & i & "].STATUS_NAME", strRet)
200                           rsStatus!STATUS_SITUATION = JSONParse("patient_status[" & i & "].STATUS_SITUATION", strRet)
210                           rsStatus.Update
220                       Next
230                   End If
240               End If
250           End If
              '加载询问界面
260           If Not rsAsk Is Nothing Then
270               If rsAsk.RecordCount > 0 Then
280                   If Not rsStatus Is Nothing Then
290                       rsStatus.Filter = "": rsAsk.Filter = ""
300                       If rsStatus.RecordCount > 0 Then
310                           For i = 1 To rsAsk.RecordCount
320                               rsStatus.Filter = "STATUS_NAME='" & rsAsk!Index & "'"
330                               If rsStatus.RecordCount > 0 Then
340                                   rsAsk!Default = IIf(rsStatus!STATUS_SITUATION & "" = "3", "否", "是")
350                               End If
360                               rsAsk.MoveNext
370                           Next
380                       End If
390                   End If
400                   rsAsk.Filter = ""
410                   If rsAsk.RecordCount > 0 Then
420                       If Not frmPassAsk.ShowMe(gfrmMain, rsAsk, strResult) Then Exit Sub
430                       If strResult <> "" Then strPara = Replace(strPara, "</cusrules>", "<result>" & strResult & "</result></cusrules>")
440                   End If
450
460               End If
470           End If
480       End If

490      Exit Sub
errH:
500       MsgBox "AskPatiStatus 错误行:" & Erl() & " 错误号:" & Err.Number & "错误描述:" & Err.Description, vbExclamation + vbOKOnly, gstrSysName
End Sub

