Attribute VB_Name = "mdlPassDefine_YWS"
Option Explicit

Public gstrBaseXml As String     '保存BaseXML


Public Function YWS_MakeBASEXML(ByRef xmlbase As YWS_BASE) As String
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
                        strTab2 & "<type_name>" & .str医生级别名称 & "</type_name >" & _
                    strTab1 & "</doct>" & vbCrLf & _
                    "</base_xml>"
                
    End With
    YWS_MakeBASEXML = strXML
End Function

Public Function YWS_MakeMedicXML(ByRef xmldetails As YWS_DETAILS) As String
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
    YWS_MakeMedicXML = strXML
End Function

Public Function YWS_MakePresXML(ByRef xmldetails As YWS_DETAILS) As String
'功能：'HIS命令 ：6、8、9
    Dim strXML As String, strTmp As String
    Dim strTab1 As String, strTab2 As String, strTab3 As String, strTab4 As String
    Dim udt过敏源 As YWS_ALLERGIC
    Dim udt诊断 As YWS_DIAGNOSE
    Dim udt药品 As YWS_MEDICINE
    
    Dim i As Long
    
    
    strTab1 = vbCrLf & vbTab
    strTab2 = vbCrLf & vbTab & vbTab
    strTab3 = vbCrLf & vbTab & vbTab & vbTab
    strTab4 = vbCrLf & vbTab & vbTab & vbTab & vbTab
    With xmldetails
        strXML = "<details_xml>" & _
                    strTab1 & "<his_time>" & .strHIS系统时间 & "</his_time>" & _
                    strTab1 & "<hosp_flag>" & .str门诊住院标识 & "</hosp_flag>" & _
                    strTab1 & "<treat_type>" & .str就诊类型 & "</treat_type>" & _
                    strTab1 & "<treat_code>" & .str就诊号 & "</treat_code>" & _
                    strTab1 & "<bed_no>" & .str床位号 & "</bed_no>"
        With .udt病人信息
            strXML = strXML & _
            strTab1 & "<patient>" & _
                strTab2 & "<name>" & .str姓名 & "</name>" & _
                strTab2 & "<birth>" & .str出生日期 & "</birth>" & _
                strTab2 & "<sex>" & .str性别 & "</sex>" & _
                strTab2 & "<weight>" & .str体重 & "</weight>" & _
                strTab2 & "<height>" & .str身高 & "</height>" & _
                strTab2 & "<id_card>" & .str身份证号 & "</id_card>" & _
                strTab2 & "<medical_record>" & .str病历卡号 & "</medical_record>" & _
                strTab2 & "<card_type>" & .str卡类型 & "</card_type>" & _
                strTab2 & "<card_code>" & .str卡号 & "</card_code>" & _
                strTab2 & "<pregnant_unit>" & .str怀孕时间单位 & "</pregnant_unit>" & _
                strTab2 & "<pregnant >" & .str怀孕时间 & "</pregnant>"
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
        End With
        strXML = strXML & strTab1 & "</patient>"
        '处方信息
        strXML = strXML & strTab1 & "<prescription_data>" & strTab2 & "<prescription>"
        With .udt处方信息
            strXML = strXML & _
            strTab3 & "<id>" & .str处方号 & "</id>" & _
            strTab3 & "<reason>" & .str处方理由 & "</reason>" & _
            strTab3 & "<is_current>" & .str是否当前处方 & "</is_current>" & _
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
                    strTab3 & "<medicine>" & _
                        strTab4 & "<zxy_type>" & .str药品类型 & "</zxy_type>" & _
                        strTab4 & "<oeridid>" & .str处方号 & "</oeridid>" & _
                        strTab4 & "<pres_type>" & .Str医嘱类型 & "</pres_type>" & _
                        strTab4 & "<pres_time>" & .str处方时间 & "</pres_time>" & _
                        strTab4 & "<name>" & .str商品名 & "</name>" & _
                        strTab4 & "<his_code>" & .str医院药品代码 & "</his_code>" & _
                        strTab4 & "<insur_code>" & .str医保代码 & "</insur_code>" & _
                        strTab4 & "<approval>" & .str批准文号 & "</approval>" & _
                        strTab4 & "<spec>" & .str规格 & "</spec>" & _
                        strTab4 & "<group>" & .str组号 & "</group>" & _
                        strTab4 & "<reason>" & .str用药理由 & "</reason>" & _
                        strTab4 & "<dose_unit>" & .str单次量单位 & "</dose_unit>" & _
                        strTab4 & "<dose>" & .str单次量 & "</dose>" & _
                        strTab4 & "<freq>" & .str频次代码 & "</freq>" & _
                        strTab4 & "<administer>" & .str给药途径代码 & "</administer>" & _
                        strTab4 & "<begin_time>" & .str用药开始时间 & "</begin_time>" & _
                        strTab4 & "<end_time>" & .str用药结束时间 & "</end_time>" & _
                        strTab4 & "<days>" & .str服药天数 & "</days>" & _
                    strTab3 & "</medicine>"
                End With
            Next
            strXML = strXML & strTab2 & "<medicine_data>" & strTmp & strTab2 & "</medicine_data>"
        End With
        strXML = strXML & strTab2 & "</prescription>" & strTab1 & "</prescription_data>"
        strXML = strXML & vbCrLf & "</details_xml>"
    End With
'    Debug.Print strXML
    
    YWS_MakePresXML = strXML
End Function

Public Function YWS_StrToXML(ByVal strValue As String) As String
'功能:将特殊字符的替换成规定字符
    strValue = Replace(strValue, "&", "&amp;")
    strValue = Replace(strValue, ">", "&gt;")
    strValue = Replace(strValue, "<", "&lt;")
    strValue = Replace(strValue, "'", "&apos;")
    YWS_StrToXML = Replace(strValue, """", "&quot;")
End Function

Public Function YWS_MakeDetailXML(ByVal bytFunc As YWS_Func_NUM) As String
'功能：构造details XML字符串
    Dim strXML As String
    
        Select Case bytFunc
        
        Case YWS_退出
            strXML = "" & _
            "<details_xml>" & vbCrLf & _
                vbTab & "<details_info></details_info>" & vbCrLf & _
            "</details_xml>"
        Case YWS_初始客户端
            strXML = "<details_xml></details_xml>"
        End Select
    
    YWS_MakeDetailXML = strXML
End Function

Public Function YWS_GetTreatType(ByVal bytFunc As Byte, ByVal lng挂号ID As Long, Optional lng主页ID As Long) As String
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
    YWS_GetTreatType = strRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function YWS_ReturnRS(ByVal strXML As String, Optional ByVal bytFunc As Byte) As ADODB.Recordset
    '功能：将药卫士审查结果集进行解析，分离出存在审核问题的药嘱行
    'bytFunc=0 解析警示灯   =1 解析审查结果
    '<?xml version='1.0' encoding='utf-8'?>
    '<ui_results_xml>
    '  <result_data>
    '    <result>
    '      <oeridid>30064、9911</oeridid>
    '      <result_type>3</result_type>
    '      <result_code>1</result_code>
    '      <result_title>配伍禁忌</result_title>
    '      <title>氨茶碱注射液≡≡ 维生素Ｃ注射液存在配伍禁忌</title>
    '      <detail>结果：注射剂配伍有禁忌，应避免！；机制：维生素C注射液的产品资料说明：维生素C不宜与碱性药物如氨茶碱溶液配伍，以免影响疗效。另有资料显示，维生素C注射液(浓度为12.5%，pH值为5.7～7.0)与氨茶碱注射液(浓
    '      度为25mg/ml，pH值为8.6～9.3)混合后，溶液呈现配伍禁忌[1]。</detail>
    '      <reference>参考文献：</reference>
    '      <mediA_hiscode>30064</mediA_hiscode>
    '      <mediA_ywscode></mediA_ywscode>
    '      <mediA_name>氨茶碱注射液</mediA_name>
    '      <mediB_hiscode>9911</mediB_hiscode>
    '      <mediB_ywscode></mediB_ywscode>
    '      <mediB_name>维生素Ｃ注射液</mediB_name>
    '    </result>
    '  </result_data>
    '</ui_results_xml>
        Dim xmlDoc As DOMDocument
        Dim xmlRoot As IXMLDOMElement
        Dim xmlNode As IXMLDOMNode
        Dim xmlNodes As IXMLDOMNodeList
        Dim rsRet As New ADODB.Recordset
        Dim arrTmp As Variant
    
        Dim str警示值 As String
        Dim str医嘱ID As String, str医嘱串 As String
    
        Dim i As Long
    
        On Error GoTo errH
    
100     Set xmlDoc = New DOMDocument
102     xmlDoc.loadXML (strXML)
104     If bytFunc = 0 Then
106         Set rsRet = InitAdviceRS(FUN_审查结果)
        Else
108         Set rsRet = InitAdviceRS(FUN_审查结果_YWS)
        End If
        '如果不包含任何元素，则退出
110     If xmlDoc.documentElement Is Nothing Then
112         Set xmlDoc = Nothing
114         Set YWS_ReturnRS = rsRet
            Exit Function
        End If
    
        '读取XML内容
116     Set xmlRoot = xmlDoc.selectSingleNode("ui_results_xml/result_data")
118     Set xmlNodes = xmlRoot.selectNodes("result")
120     If bytFunc = 0 Then
122         If Not xmlNodes Is Nothing Then
124             For Each xmlNode In xmlNodes
126                 str警示值 = xmlNode.selectSingleNode("result_type").Text
128                 If Val(str警示值) > 0 Then
130                     str医嘱串 = xmlNode.selectSingleNode("oeridid").Text   '30064、9911
132                     arrTmp = Split(str医嘱串, "、")
134                     For i = LBound(arrTmp) To UBound(arrTmp)
136                         str医嘱ID = Val(arrTmp(i))
138                         If Val(arrTmp(i)) <> 0 Then
140                             rsRet.Filter = "医嘱ID ='" & str医嘱ID & "'"
142                             If Not rsRet.EOF Then
144                                 If Val(rsRet!警示值 & "") < Val(str警示值) Then
146                                     rsRet!警示值 = Val(str警示值)
                                    End If
                                Else
148                                 rsRet.AddNew
150                                 rsRet!警示值 = Val(str警示值)
152                                 rsRet!医嘱ID = Val(str医嘱ID)
154                                 rsRet.Update
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Else
156         If Not xmlNodes Is Nothing Then
158             For Each xmlNode In xmlNodes
160                 rsRet.AddNew
162                 rsRet!Title = xmlNode.selectSingleNode("title").Text
164                 rsRet!Detail = xmlNode.selectSingleNode("detail").Text
166                 rsRet.Update
                Next
            End If
        End If
168     If rsRet.RecordCount > 0 Then rsRet.Filter = ""
    
170     Set YWS_ReturnRS = rsRet
        Exit Function
errH:
172     MsgBox "YWS_ReturnRS 错误号:" & Err.Number & "错误行:" & Erl() & " 错误描述:" & Err.Description, vbOKOnly, gstrSysName
End Function

Public Function YWS_GetDrugType(ByVal strType As String) As String
'功能:返回药品类型
'   5-西药/6-中成药/7-草药
    Dim strRet As String
    
    Select Case strType
    
    Case "5"
        strRet = "西药"
    Case "6"
        strRet = "中成药"
    Case "7"
        strRet = "草药"
    End Select
    YWS_GetDrugType = strRet
End Function
