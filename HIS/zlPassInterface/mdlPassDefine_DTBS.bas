Attribute VB_Name = "mdlPassDefine_DTBS"
Option Explicit

'大通BS版接口定义
Public Declare Function CRMS_UI Lib "CRMS_UI.dll" (ByVal lngFunc As Long, ByVal strBaseXml As String, ByVal strDetailsXml As String, ByRef strResults As String) As Long
'参数：lngFunc(功能标识)
'     strBaseXml(基本信息XMl)
'     strDetailsXml（详细信息XML）
'返回
'    strResults（his返回结果XML）

Public Function DTBS_StrToXML(ByVal strValue As String) As String
'功能:将特殊字符的替换成规定字符
    strValue = Replace(strValue, "&", "&amp;")
    strValue = Replace(strValue, ">", "&gt;")
    strValue = Replace(strValue, "<", "&lt;")
    strValue = Replace(strValue, "'", "&apos;")
    DTBS_StrToXML = Replace(strValue, """", "&quot;")
End Function

Public Function DTBS_MakeBASEXML(ByRef xmlbase As DTBS_BASE) As String
'功能：构造BASE XML字符串
    Dim strXML As String
    Dim strTab1 As String, strTab2 As String
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    
    With xmlbase
        strXML = "<base_xml>" & _
                    strTab1 & "<source>" & .strHIS & "</source>" & _
                    strTab1 & "<hosp_code>" & .str医院编码 & "</hosp_code>" & _
                    strTab1 & "<dept_code>" & .str科室代码 & "</dept_code>" & _
                    strTab1 & "<dept_name>" & .str科室名称 & "</dept_name>" & _
                    strTab1 & "<doct>" & _
                        strTab2 & "<code>" & .str医生代码 & "</code>" & _
                        strTab2 & "<name>" & .str医生名称 & "</name>" & _
                        strTab2 & "<type>" & .str医生级别代码 & "</type>" & _
                        strTab2 & "<type_name>" & .str医生级别名称 & "</type_name>" & _
                    strTab1 & "</doct>" & vbCrLf & _
                    "</base_xml>"
                
    End With
    DTBS_MakeBASEXML = strXML
End Function

Public Function DTBS_MakeDetailXML(ByVal bytFunc As DTBS_Func_NUM, Optional ByVal strDoctPWD As String) As String
'功能：构造details XML字符串
    Dim strXML As String
    
        Select Case bytFunc
        Case DTBS_登录
            strXML = "<details_xml>" & vbCrLf & _
                        "<doct_pwd>" & strDoctPWD & "</doct_pwd>" & vbCrLf & _
                    "</details_xml>"
        Case DTBS_退出
            strXML = "" & _
            "<details_xml>" & vbCrLf & _
                vbTab & "<details_info></details_info>" & vbCrLf & _
            "</details_xml>"
        Case DTBS_初始UI
            strXML = "<details_xml></details_xml>"
        End Select
    
    DTBS_MakeDetailXML = strXML
End Function

Public Function DTBS_MakeMedicXML(ByRef xmldetails As DTBS_DETAILS) As String
'功能：HIS命令 ：5调用
    Dim strXML As String
    Dim strTab1 As String, strTab2 As String
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    
    With xmldetails
        strXML = "<details_xml>" & _
                    strTab1 & "<hosp_flag>" & .str门诊住院标识 & "</hosp_flag>" & _
                    strTab1 & "<medicine>" & _
                        strTab2 & "<his_code>" & .str药品代码 & "</his_code>" & _
                        strTab2 & "<his_name>" & .str药品名称 & "</his_name>" & _
                    strTab1 & "</medicine>" & vbCrLf & _
                "</details_xml>"
    End With
'    Debug.Print strXML
    DTBS_MakeMedicXML = strXML
End Function

Public Function DTBS_MakePresXML(ByRef xmldetails As DTBS_DETAILS) As String
'功能：'HIS命令 ：6
    Dim strXML As String, strTmp As String, strSub As String, strPres As String
    Dim strTab1 As String, strTab2 As String, strTab3 As String, strTab4 As String, strTab5 As String
    Dim udt过敏源 As DTBS_ALLERGIC
    Dim udt诊断 As DTBS_DIAGNOSE
    Dim udt处方信息 As DTBS_PRESCRIPTION
    Dim udtLISFORM As DTBS_LISFORM
    Dim udtLISITEM As DTBS_LISITEM
    Dim udt药品 As DTBS_MEDICINE
    
    Dim i As Long, j As Long
    
    
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    strTab3 = vbCrLf & vbTab & vbTab & vbTab
    strTab4 = vbCrLf & vbTab & vbTab & vbTab & vbTab
    strTab5 = vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
    
    With xmldetails
        strXML = "<details_xml  is_upload =""" & .str是否上传 & """>" & _
                    strTab1 & "<his_time>" & .strHIS系统时间 & "</his_time>" & _
                    strTab1 & "<hosp_flag>" & .str门诊住院标识 & "</hosp_flag>" & _
                    strTab1 & "<treat_type>" & .str就诊类型 & "</treat_type>" & _
                    strTab1 & "<treat_code>" & .str就诊号 & "</treat_code>" & _
                    strTab1 & "<lis_adm_no>" & .str检验就诊号 & "</lis_adm_no>" & _
                    strTab1 & "<bed_no>" & .str床位号 & "</bed_no>" & _
                    strTab1 & "<area_code>" & .str病区号 & "</area_code>"
        With .udt病人信息
            strXML = strXML & _
            strTab1 & "<patient>" & _
                strTab2 & "<name>" & .str姓名 & "</name>" & _
                strTab2 & "<is_infant>" & .str是否婴儿 & "</is_infant>" & _
                strTab2 & "<birth>" & .str出生日期 & "</birth>" & _
                strTab2 & "<sex>" & .str性别 & "</sex>" & _
                strTab2 & "<weight>" & .str体重 & "</weight>" & _
                strTab2 & "<height>" & .str身高 & "</height>" & _
                strTab2 & "<id_card>" & .str身份证号 & "</id_card>" & _
                strTab2 & "<card_type>" & .str卡类型 & "</card_type>" & _
                strTab2 & "<card_code>" & .str卡号 & "</card_code>" & _
                strTab2 & "<pregnant_unit>" & .str怀孕时间单位 & "</pregnant_unit>" & _
                strTab2 & "<pregnant>" & .str怀孕时间 & "</pregnant>"
            '过敏源
            strTmp = ""
            If Not .col过敏源s Is Nothing Then
                For i = 1 To .col过敏源s.Count
                    udt过敏源 = .col过敏源s(i)
                    With udt过敏源
                        strTmp = strTmp & _
                        strTab3 & "<allergic>" & _
                            strTab4 & "<type>" & .str过敏类型 & "</type>" & _
                            strTab4 & "<name>" & .str过敏源名称 & "</name>" & _
                            strTab4 & "<code>" & .str过敏源代码 & "</code>" & _
                        strTab3 & "</allergic>"
                    End With
                Next
            End If
            strXML = strXML & strTab2 & "<allergic_data>" & strTmp & strTab2 & "</allergic_data>"
            
            '诊断
            strTmp = ""
            If Not .col诊断s Is Nothing Then
                For i = 1 To .col诊断s.Count
                    udt诊断 = .col诊断s(i)
                    With udt诊断
                        strTmp = strTmp & _
                        strTab3 & "<diagnose>" & _
                            strTab4 & "<type>" & .str诊断类型 & "</type>" & _
                            strTab4 & "<name>" & .str诊断名称 & "</name>" & _
                            strTab4 & "<code>" & .str诊断代码 & "</code>" & _
                        strTab3 & "</diagnose>"
                    End With
                Next
            End If
            strXML = strXML & strTab2 & "<diagnose_data>" & strTmp & strTab2 & "</diagnose_data>"
            '检验检测单节点
            strTmp = ""
            If Not .col检验检查 Is Nothing Then
                For i = 1 To .col检验检查.Count
                    udtLISFORM = .col检验检查(i)
                    With udtLISFORM
                        strTmp = strTmp & _
                        strTab3 & "<form>" & _
                            strTab4 & "<no>" & .str单号 & "</no>" & _
                            strTab4 & "<project_name>" & .str项目名称 & "</project_name>" & _
                            strTab4 & "<lis_flag>" & .str标记 & "</lis_flag>" & _
                            strTab4 & "<result_date>" & .str结果出具时间 & "</result_date>" & _
                            strTab4 & "<sample_code>" & .str检验样本编码 & "</sample_code>" & _
                            strTab4 & "<sample_name>" & .str检验样本名称 & "</sample_name>" & _
                            strTab4 & "<mac_flag>" & .str微生物送检标识 & "</mac_flag>" & _
                        strTab3 & "</form>"
                        
                        strSub = ""
                        If Not .col项目节点 Is Nothing Then
                            For j = 1 To .col项目节点.Count
                                udtLISITEM = .col项目节点(i)
                                With udtLISITEM
                                    strSub = strSub & _
                                    strTab4 & "<item>" & _
                                        strTab5 & "<code>" & .str编码 & "</code>" & _
                                        strTab5 & "<name>" & .str名称 & "</name>" & _
                                        strTab5 & "<value>" & .str结果 & "</value>" & _
                                        strTab5 & "<uom>" & .str结果值单位 & "</uom>" & _
                                        strTab5 & "<upper>" & .str参考范围上限 & "</upper>" & _
                                        strTab5 & "<lower>" & .str参考范围下限 & "</lower>"
                                End With
                            Next
                        End If
                        strTmp = strTmp & strSub
                    End With
                Next
            End If
            strXML = strXML & strTab2 & "<lis_data>" & strTmp & strTab2 & "</lis_data>"
        End With
        strXML = strXML & strTab1 & "</patient>"
        '处方信息
        If Not .col处方信息 Is Nothing Then
            strPres = ""
            For j = 1 To .col处方信息.Count
                udt处方信息 = .col处方信息(j)
                With udt处方信息
                    strPres = strPres & strTab2 & "<prescription>" & _
                    strTab3 & "<id>" & .str处方号 & "</id>" & _
                    strTab3 & "<reason>" & .str处方理由 & "</reason>" & _
                    strTab3 & "<is_urgent>" & .str是否紧急处方 & "</is_urgent>" & _
                    strTab3 & "<is_new>" & .str是否新开处方 & "</is_new>" & _
                    strTab3 & "<is_current>" & .str是否当前处方 & "</is_current>" & _
                    strTab3 & "<doct_code>" & .str开嘱医生代码 & "</doct_code>" & _
                    strTab3 & "<doct_name>" & .str开嘱医生姓名 & "</doct_name>" & _
                    strTab3 & "<dept_code>" & .str开嘱科室代码 & "</dept_code>" & _
                    strTab3 & "<dept_name>" & .str开嘱科室名称 & "</dept_name>" & _
                    strTab3 & "<pres_type>" & .Str医嘱类型 & "</pres_type>" & _
                    strTab3 & "<pres_time>" & .str处方时间 & "</pres_time>"
                    '药品信息
                    If .col药品信息 Is Nothing Then
                        Set .col药品信息 = New Collection
                        .col药品信息.Add udt药品, "_1"
                    End If
                    
                    strTmp = ""
                    For i = 1 To .col药品信息.Count
                        udt药品 = .col药品信息(i)
                        With udt药品
                            strTmp = strTmp & _
                            strTab4 & "<medicine>" & _
                                strTab5 & "<name>" & .str商品名 & "</name>" & _
                                strTab5 & "<his_code>" & .str医院药品代码 & "</his_code>" & _
                                strTab5 & "<insur_code>" & .str医保代码 & "</insur_code>" & _
                                strTab5 & "<pyd_code>" & .str配液单号 & "</pyd_code>" & _
                                strTab5 & "<link_group>" & .str配液单组号 & "</link_group>" & _
                                strTab5 & "<spec>" & .str规格 & "</spec>" & _
                                strTab5 & "<group>" & .str组号 & "</group>" & _
                                strTab5 & "<reason>" & .str用药理由 & "</reason>" & _
                                strTab5 & "<dose_unit>" & .str单次量单位 & "</dose_unit>" & _
                                strTab5 & "<dose>" & .str单次量 & "</dose>" & _
                                strTab5 & "<freq>" & .str频次代码 & "</freq>" & _
                                strTab5 & "<administer>" & .str给药途径代码 & "</administer>" & _
                                strTab5 & "<begin_time>" & .str用药开始时间 & "</begin_time>" & _
                                strTab5 & "<end_time>" & .str用药结束时间 & "</end_time>" & _
                                strTab5 & "<days>" & .str服药天数 & "</days>" & _
                                strTab5 & "<preventiveflag>" & .str是否预防用药 & "</preventiveflag>" & _
                                strTab5 & "<otno>" & .str手术单号 & "</otno>" & _
                                strTab5 & "<signer_code>" & .str签名医师工号 & "</signer_code>" & _
                                strTab5 & "<accredit_date>" & .str授权时间 & "</accredit_date>" & _
                                strTab5 & "<accredit_hours>" & .str允许用药时间 & "</accredit_hours>" & _
                                strTab5 & "<accredit_times>" & .str允许用药次数 & "</accredit_times>" & _
                            strTab4 & "</medicine>"
                        End With
                    Next
                    strPres = strPres & strTab3 & "<medicine_data>" & strTmp & strTab3 & "</medicine_data>" & strTab2 & "</prescription>"
                End With
            Next
            strXML = strXML & strTab1 & "<prescription_data>" & strPres & strTab1 & "</prescription_data>"
        End If
        
        strXML = strXML & vbCrLf & "</details_xml>"
    End With
'    Debug.Print strXML
    
    DTBS_MakePresXML = strXML
End Function

Public Function DTBS_GetTreatType(ByVal bytFunc As Byte, ByVal lng挂号ID As Long, Optional lng主页ID As Long) As String
'功能:获取就诊类型
'参数:bytFunc =1 门诊,bytFunc=2 住院
'     lng挂号ID =门诊 挂号ID,住院 =病人ID
'100=普通门诊
'101=专科门诊
'102=专家门诊
'200=急诊
'300=急诊观察
'400=普通住院
'401=特需住院
'500=家床
'999=其他
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strRet As String
    
    If bytFunc = 1 Then
        strSQL = "Select Nvl(a.急诊,0) as 急诊,b.号类 From 病人挂号记录 A, 挂号安排 B Where a.Id = [1] And a.号别 = b.号码"
    Else
        strSQL = "Select 病人性质, 出院病床 From 病案主页 Where 病人id = [1] And 主页id = [2] And 出院日期 Is Null"

    End If
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, G_STR_PASS, lng挂号ID, lng主页ID)
    
    If rsTmp.RecordCount > 0 Then
        If bytFunc = 1 Then
            If rsTmp!急诊 = 1 Then
                strRet = "200"
            Else
                If rsTmp!号类 & "" = "普通" Then
                    strRet = "100"
                ElseIf rsTmp!号类 & "" = "专科" Then
                    strRet = "101"
                ElseIf rsTmp!号类 & "" = "专家" Then
                    strRet = "102"
                Else
                    strRet = "999"
                End If
            End If
        ElseIf bytFunc = 2 Then
            If rsTmp!出院病床 & "" = "" Then
                strRet = "500"   '家庭病床
            ElseIf rsTmp!病人性质 = 0 Then
                strRet = "400"
            ElseIf rsTmp!病人性质 = 1 Or rsTmp!病人性质 = 2 Then
                strRet = "300"
            Else
                strRet = "999"
            End If
            
        End If
    End If
    DTBS_GetTreatType = strRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

