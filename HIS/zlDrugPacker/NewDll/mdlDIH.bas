Attribute VB_Name = "mdlDIH"
Option Explicit

Private Function GetXML_RecipeDetail_DIH(ByVal LngStockID As Long, ByVal strNO As String) As String
'将处方明细组织成指定的XML格式
'适用接口：无锡碟和(DIH)
    Dim rsRecipe As Recordset   '病人和处方记录
    Dim rsDiagnosis As Recordset    '诊断记录
    Dim rsDrug As Recordset         '处方药品记录
    Dim strSql As String
    Dim strRecipe As String
    Dim strDiagnosis As String
    Dim strXML As String
    Dim strXML_Patient As String
    Dim strXML_Recipe As String
    Dim strXML_Drug As String
    Dim i As Integer
    Dim strOutput As String
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    Dim strTmp As String
    
    strOutput = strOutput & vbCrLf & "调用GetXML_RecipeDetail_DIH"
    
    On Error GoTo errHandle
    
    '判断单处方还是多处方
    If InStr(1, strNO, "|") < 1 Then
        '单处方
        strRecipe = " And a.单据=[2] And a.NO=[3] "
    Else
        '多处方
        strRecipe = " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strRecipe = strRecipe & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strRecipe = strRecipe & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strRecipe = strRecipe & ") "
    End If
    
    strOutPutExeStep = "判断和分解单处方/多处方完成"
    
    '获取病人及处方单信息
    strSql = "Select Distinct a.病人id, decode(c.姓名,null,d.姓名,c.姓名) 姓名, decode(c.性别,null,d.性别,c.性别) 性别, decode(c.年龄,null,d.年龄,c.年龄) 年龄, " & vbNewLine & _
        "       c.身份, c.医疗付款方式 医保类型, d.费别 As 收费类别, a.No As 处方号, Decode(a.处方类型, 2, 'J', 'M') As 处方类型, d.NO As 就诊编号, " & vbNewLine & _
        "       d.开单部门id As 就诊科室编码, f.名称 As 就诊科室名称, g.Id As 就诊医生编码, d.开单人 As 就诊医生姓名, d.登记时间 As 缴费时间,a.单据 " & vbNewLine & _
        " From 未发药品记录 A, 病人信息 C, 门诊费用记录 D, 药品收发记录 E, 部门表 F, 人员表 G " & vbNewLine & _
        " Where a.单据 = e.单据 And a.No = e.No And a.库房id = e.库房id And a.病人id = c.病人id(+) And e.费用id = d.Id And d.开单部门id = f.Id And " & vbNewLine & _
        "      d.开单人 = g.姓名 And a.库房id = [1] " & strRecipe
    
    strOutPutExeStep = strOutPutExeStep & vbCrLf & "查询病人及处方单信息开始：" & vbCrLf & strSql
    
    If gintMode = 0 Then
        Set rsRecipe = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsRecipe = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    End If
    
'诊断暂时不上传，所以注释
'    '获取诊断信息
'    strSql = "Select d.诊断描述, d.是否疑诊, a.单据, a.No" & vbNewLine & _
'        " From 未发药品记录 A, 门诊费用记录 B, 药品收发记录 C, 病人诊断记录 D, 病人诊断医嘱 E" & vbNewLine & _
'        " Where a.单据 = c.单据 And a.No = c.No And a.库房id = c.库房id And c.费用id = b.Id And e.医嘱id = b.医嘱序号 And d.Id = e.诊断id And" & vbNewLine & _
'        " d.取消时间 Is Null And a.库房id = [1] " & strRecipe
'
'    strOutPutExeStep = strOutPutExeStep & vbCrLf & "查询诊断信息开始：" & vbCrLf & strSql
'
'    If gintMode = 0 Then
'        Set rsDiagnosis = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
'    Else
'        Set rsDiagnosis = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
'    End If
    
    '获取处方药品信息
    strSql = "Select Distinct a.单据, a.No, b.Id 药品编码, b.名称 药品名称, b.规格 药品规格, a.产地 药品厂家, a.实际数量 / d.门诊包装 药品数量, d.门诊单位 发药单位, a.用法 As 服用方法," & vbNewLine & _
        " a.单量, g.计算单位, f.执行频次, f.医生嘱托 as 备注说明, a.库房id As 药房编码, a.序号" & vbNewLine & _
        " From 药品收发记录 A, 收费项目目录 B, 药品规格 D, 门诊费用记录 E, 病人医嘱记录 F, 诊疗项目目录 G" & vbNewLine & _
        " Where a.药品id = b.Id And a.药品id = d.药品id And a.费用id = e.Id And d.药名id = g.Id And e.医嘱序号 = f.Id(+) And a.库房id = [1] " & strRecipe
    
    strOutPutExeStep = strOutPutExeStep & vbCrLf & "查询处方药品信息开始：" & vbCrLf & strSql
    
    If gintMode = 0 Then
        Set rsDrug = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsDrug = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeDetail_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    End If
    
    strOutput = strOutput & vbCrLf & "查询处方明细完成"
    
'    1.1.1、门诊处方xml格式:
'    <outpOrder>
'        <patient>                            - 患者信息
'        <windowNo></windowNo>   - 取药窗口号（只在由 中联HIS 分配窗口号时有，且不能为 0）
'        <patientID></patientID>           - 患者唯一 ID
'        <patientName></patientName>       - 姓名
'        <patientGender></patientGender>   -性别
'        <patientAge></patientAge>          -年龄
'        <identity></identity>                 -身份
'        <insuranceType></insuranceType>    -医保类型
'        <chargeType></chargeType>         -收费类别
'        </patient>
'        <prescriptions>  - 处方清单
'            <prescription no="" type="" paymentDT="">   - 处方，no：处方唯一编号，type：M-门诊，J-急诊，O-其他，缴费时间：yyyy-MM-dd HH:mm:ss
'            <outpNo></outpNo>                 - 病历号
'            <visitNo></visitNo>               - 就诊编号
'            <deptCode></deptCode>             - 就诊科室编码
'            <deptName></deptName>             - 就诊科室名称
'            <doctCode></doctCode>             - 就诊医生编码
'            <doctName></doctName>             - 就诊医生姓名
'            <diagnosis></diagnosis>           - 临床诊断
'            <paymentDT></paymentDT>       -缴费时间：yyyy-MM-dd HH:mm:ss
'            <drugList>   -处方中药品清单
'                <drug>
'                <drugCode></drugCode>           -   药品编码
'                <drugName></drugName>       -   名称
'                <drugSpec></drugSpec>          -    规格
'                <firmName></firmName>          -    厂家
'                <amount></amount>             - 药品数量
'                <takeUnit></takeUnit>             - 发药单位
'                <takeMethod></takeMethod>      -    服用方法
'                <takeDosage></takeDosage>       - 用量
'                <takeType></takeType>           -   服用类型
'                <takeNote></takeNote>          -    备注说明
'                <pharmacyCode></pharmacyCode>   -   药房编码
'                <sortNo></sortNo>     - 在处方中的药品顺序号（整数值）
'                </drug>
'            </drugList>
'            </prescription>
'        </prescriptions>
'    </outpOrder>

    Call OutputLog(vbCrLf & strOutPutExeStep & vbCrLf & _
                    "相关参数：" & "LngStockID=" & LngStockID & " strNO=" & strNO)

    If rsRecipe.RecordCount > 0 Then
        rsRecipe.MoveFirst
    
        '病人信息
        With rsRecipe
            strOutPutExeStep = "组织病人信息XML"
            
            strXML_Patient = "<patient>"
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("windowNo", "", False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("patientID", NVL(!病人id), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("patientName", NVL(!姓名), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("patientGender", NVL(!性别), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("patientAge", NVL(!年龄), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("identity", NVL(!身份), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("insuranceType", NVL(!医保类型), False)
            strXML_Patient = strXML_Patient & vbCrLf & GetXMLFormat("chargeType", NVL(!收费类别), False)
            strXML_Patient = strXML_Patient & "</patient>"
        End With
        
        '处方信息
        With rsRecipe
            strXML_Recipe = "<prescriptions>"
            Do While Not .EOF
                strOutPutExeStep = "组织处方信息XML"
                
                strXML_Recipe = strXML_Recipe & vbCrLf & "<prescription no=""" & NVL(!处方号) & """ type=""" & NVL(!处方类型) & """ paymentDT=""" & Format(NVL(!缴费时间), "yyyy-MM-DD hh:mm:ss") & """>"
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("outpNo", "", False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("visitNo", NVL(!就诊编号), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("deptCode", NVL(!就诊科室编码), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("deptName", NVL(!就诊科室名称), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("doctCode", NVL(!就诊医生编码), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("doctName", NVL(!就诊医生姓名), False)
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("diagnosis", "", False)   '诊断
                strXML_Recipe = strXML_Recipe & vbCrLf & GetXMLFormat("paymentDT", Format(NVL(!缴费时间), "yyyy-MM-DD hh:mm:ss"), False)
                
                '药品信息
                strXML_Drug = "<drugList>"
                rsDrug.Filter = "no='" & !处方号 & "' and 单据=" & NVL(!单据)
                rsDrug.Sort = "序号"
                
                Do While Not rsDrug.EOF
                    strOutPutExeStep = "组织药品信息XML"
                    
                    strXML_Drug = strXML_Drug & vbCrLf & "<drug>"
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("drugCode", NVL(rsDrug!药品编码), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("drugName", NVL(rsDrug!药品名称), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("drugSpec", NVL(rsDrug!药品规格), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("firmName", NVL(rsDrug!药品厂家), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("amount", NVL(rsDrug!药品数量), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeUnit", NVL(rsDrug!发药单位), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeMethod", NVL(rsDrug!服用方法), False)
                    If NVL(rsDrug!单量) = "" Then
                        strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeDosage", "", False)
                    Else
                        strTmp = Format(rsDrug!单量, "#0.##########") & NVL(rsDrug!计算单位) & "；" & NVL(rsDrug!执行频次)
                        strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeDosage", strTmp, False)
                    End If
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeType", "", False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("takeNote", NVL(rsDrug!备注说明), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("pharmacyCode", NVL(rsDrug!药房编码), False)
                    strXML_Drug = strXML_Drug & vbCrLf & GetXMLFormat("sortNo", NVL(rsDrug!序号), False)
                    strXML_Drug = strXML_Drug & vbCrLf & "</drug>"
                    
                    rsDrug.MoveNext
                Loop
                
                strXML_Drug = strXML_Drug & vbCrLf & "</drugList>"
                strXML_Recipe = strXML_Recipe & vbCrLf & strXML_Drug & vbCrLf & "</prescription>"
                
                rsRecipe.MoveNext
            Loop
            
            '汇总处方药品
            strXML_Recipe = strXML_Recipe & "</prescriptions>"
        End With
        
        '汇总病人、处方和药品信息，拼凑完整的XML
        strXML = "<outpOrder>"
        strXML = strXML & vbCrLf & strXML_Patient
        strXML = strXML & vbCrLf & strXML_Recipe
        strXML = strXML & vbCrLf & "</outpOrder>"
    Else
        strOutput = strOutput & vbCrLf & "无处方数据"
    End If
    
    GetXML_RecipeDetail_DIH = strXML
    
    strOutput = strOutput & vbCrLf & "组织处方信息XML完成：" & vbCrLf & strXML
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_RecipeDetail_DIH"
    Call OutputLog(strOutput)
    
    Exit Function
errHandle:
    strOutput = strOutput & vbCrLf & "发生异常错误:" & Err.Description
    
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog

    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "相关参数：" & "LngStockID=" & LngStockID & " strNO=" & strNO
    strOutput = strOutput & vbCrLf & "相关SQL" & vbCrLf & strSql
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_RecipeDetail_DIH"
    Call OutputLog(strOutput)
End Function

Private Function GetXML_RecipeReady_DIH(ByVal LngStockID As Long, ByVal strNO As String) As Variant
'将处方单组织成指定的XML格式
'取药通知/准备发药
'适用接口：无锡碟和(DIH)
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim strSql As String
    Dim strOutput As String
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "调用函数：GetXML_RecipeReady_DIH"
   
    strSql = "Select b.编码 As 窗口号, a.病人id, a.Groupno, a.Ordertype " & _
        " From 未发药品记录 A, 发药窗口 B " & _
        " Where a.发药窗口 = b.名称(+) And a.库房id=[1] "
    
    '判断单处方还是多处方
    If InStr(1, strNO, "|") < 1 Then
        '单处方
        strSql = strSql & " And a.单据=[2] And a.NO=[3] "
    Else
        '多处方
        strSql = strSql & " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strSql = strSql & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strSql = strSql & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strSql = strSql & ") "
    End If
    
    strOutPutExeStep = "判断和分解单处方/多处方完成"
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeReady_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeReady_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    End If
    
    strOutPutExeStep = "执行SQL完成"
    
    With rsTemp
        '检查GroupNO是否填写
        If NVL(!Groupno) = "" Then
            strOutput = strOutput & vbCrLf & "GroupNO未填写，即设备未配药完成。"
            Call OutputLog(strOutput)
            Exit Function
        End If
    
'    <outpOrderTake>
'        <windowNo>2</windowNo>              --窗口号
'        <patientID>1042323</patientID>      --病人唯一ID
'        <groupNo>M15091800973</groupNo>     --组号
'        <orderType>indirect</orderType>     --处方直发或者预配发标识
'    </outpOrderTake>

        If .RecordCount > 0 Then
            strXML = "<outpOrderTake>"
            strXML = strXML & vbCrLf & GetXMLFormat("windowNo", NVL(!窗口号), False)
            strXML = strXML & vbCrLf & GetXMLFormat("patientID", NVL(!病人id), False)
            strXML = strXML & vbCrLf & GetXMLFormat("groupNo", NVL(!Groupno), False)
            strXML = strXML & vbCrLf & GetXMLFormat("orderType", NVL(!orderType), False)
            strXML = strXML & vbCrLf & "</outpOrderTake>"
        End If
        
        strOutPutExeStep = "拼装XML完成"
    End With
    
    GetXML_RecipeReady_DIH = strXML
    
    strOutput = strOutput & vbCrLf & "组织XML完成：" & vbCrLf & strXML
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_RecipeReady_DIH"
    Call OutputLog(strOutput)
    
    Exit Function
    
errHandle:
    strOutput = strOutput & vbCrLf & "发生异常错误：" & Err.Description
    
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "相关参数：" & "LngStockID=" & LngStockID & " strNO=" & strNO
    strOutput = strOutput & vbCrLf & "相关SQL" & vbCrLf & strSql
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_RecipeReady_DIH"
    Call OutputLog(strOutput)
End Function

Private Function GetXML_RecipeCompletion_DIH(ByVal LngStockID As Long, ByVal strNO As String) As String
'将处方单组织成指定的XML格式
'发药完成
'适用接口：无锡碟和(DIH)
    Dim strXML As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim strSql As String
    Dim strOutput As String
    Dim strOutPutExeStep As String    '执行步骤，用于输出日志方便查找问题
    
    On Error GoTo errHandle
    
    strOutput = strOutput & vbCrLf & "调用函数：GetXML_RecipeCompletion_DIH"
   
    strSql = "Select a.病人id, a.Groupno " & _
        " From 未发药品记录 A " & _
        " Where a.库房id=[1] "
    
    '判断单处方还是多处方
    If InStr(1, strNO, "|") < 1 Then
        '单处方
        strSql = strSql & " And a.单据=[2] And a.NO=[3] "
    Else
        '多处方
        strSql = strSql & " And ("
        For i = 0 To UBound(Split(strNO, "|"))
            If i = UBound(Split(strNO, "|")) Then
                strSql = strSql & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "')"
            Else
                strSql = strSql & "(a.单据=" & Split(Split(strNO, "|")(i), ",")(0) & " And a.NO='" & Split(Split(strNO, "|")(i), ",")(1) & "') or "
            End If
        Next
        strSql = strSql & ") "
    End If
    
    strOutPutExeStep = "判断和分解单处方/多处方完成"
    
    If gintMode = 0 Then
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetXML_RecipeCompletion_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    Else
        Set rsTemp = mdlDrugPacker.OpenSQLRecord(strSql, "GetXML_RecipeCompletion_DIH", LngStockID, CInt(Split(strNO, ",")(0)), CStr(Split(strNO, ",")(1)))
    End If
    
    strOutPutExeStep = "执行SQL完成"
    
    With rsTemp
'        <outpOrderCompletion>
'            <patientID>103278</patientID>      --患者唯一ID
'            <groupNo>M15092201309</groupNo>     --组号
'        </outpOrderCompletion>

        If .RecordCount > 0 Then
            strXML = "<outpOrderCompletion>"
            strXML = strXML & vbCrLf & GetXMLFormat("patientID", NVL(!病人id), False)
            strXML = strXML & vbCrLf & GetXMLFormat("groupNo", NVL(!Groupno), False)
            strXML = strXML & vbCrLf & "</outpOrderCompletion>"
        End If
        
        strOutPutExeStep = "拼装XML完成"
    End With
    
    GetXML_RecipeCompletion_DIH = strXML
    
    strOutput = strOutput & vbCrLf & "组织XML完成：" & vbCrLf & strXML
    strOutput = strOutput & vbCrLf & "执行成功：GetXML_RecipeCompletion_DIH"
    Call OutputLog(strOutput)
    
    Exit Function
errHandle:
    strOutput = strOutput & vbCrLf & "发生异常错误：" & Err.Description
    
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    Else
        MsgBox Err.Description, vbInformation, GSTR_SYSNAME
    End If
    
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "相关参数：" & "LngStockID=" & LngStockID & " strNO=" & strNO
    strOutput = strOutput & vbCrLf & "相关SQL" & vbCrLf & strSql
    strOutput = strOutput & vbCrLf & "执行失败：GetXML_RecipeCompletion_DIH"
    Call OutputLog(strOutput)
End Function

Public Function HisTransData_DIH(ByVal lngOper As Long, ByVal LngStockID As Long, ByVal strNO As String, ByRef strReturn As String) As Boolean
    '上传HIS数据到对方系统，或上传关键信息并接收对方返回信息
    Dim strInputXML As String
    Dim strOutput As String
    Dim strTmp As String
    Dim strOutXML As String
    Dim strOut_RETCODE As String
    Dim strOut_WinNo As String
    Dim strOut_Msg As String
    Dim strOutPutExeStep As String
    Dim objXML As clsXML
    
    strOutput = strOutput & vbCrLf & "调用函数：HisTransData_DIH"
    strOutput = strOutput & vbCrLf & "lngOper=" & lngOper
    strOutput = strOutput & vbCrLf & "strNO=" & strNO
    
    On Error GoTo errHandle
    
    '业务代码：
    Select Case lngOper
        Case gType.IntDetail
            '上传处方明细
            strInputXML = GetXML_RecipeDetail_DIH(LngStockID, strNO)
            
            If strInputXML = "" Then Exit Function
            
            '调用对方接口
            strOutPutExeStep = "调用对方接口开始：outpOrderDispense"
            strOutXML = gobjSOAP.outpOrderDispense(strInputXML)
            strOutPutExeStep = "调用对方接口完成：outpOrderDispense"
            
            strOutput = strOutput & vbCrLf & "调用对方接口完成：outpOrderDispense"
        Case gType.IntStartList
            '取药通知/准备发药
            strInputXML = GetXML_RecipeReady_DIH(LngStockID, strNO)
            
            If strInputXML = "" Then Exit Function
            
            '调用对方接口
            strOutPutExeStep = "调用对方接口开始：outpOrderTakeNotify"
            strOutXML = gobjSOAP.outpOrderTakeNotify(strInputXML)
            strOutPutExeStep = "调用对方接口完成：outpOrderTakeNotify"
            
            strOutput = strOutput & vbCrLf & "调用对方接口完成：outpOrderTakeNotify"
        Case gType.IntEndList
            HisTransData_DIH = True
            Exit Function
            
            '云南渠道要求屏蔽调用发药成功接口，因为他们有个轮循工具在做这个事。
        
            '发药完成
            strInputXML = GetXML_RecipeCompletion_DIH(LngStockID, strNO)
            
            If strInputXML = "" Then Exit Function
            
            '调用对方接口
            strOutPutExeStep = "调用对方接口开始：outpOrderCompletionNotify"
            strOutXML = gobjSOAP.outpOrderCompletionNotify(strInputXML)
            strOutPutExeStep = "调用对方接口完成：outpOrderCompletionNotify"
            
            strOutput = strOutput & vbCrLf & "调用对方接口完成：outpOrderCompletionNotify"
    End Select
        
    strOutput = strOutput & vbCrLf & "返回信息：" & vbCrLf & strOutXML
                            
    '返回信息格式
    ''上传处方明细
    '    <result>
    '    <status code="0" message="OK"/> -  code：非零为错误编号，message:     结果描述
    '    <value> - 接口返回结果内容
    '    <windowNo>2</windowNo> - DIH系统分配的取药窗口号,如果接受失败返回默认的一个窗口号
    '    </value>
    '    </result>
    
    ''取药通知/准备发药
    '<result>
    '<status code="0" message=""/>       - code：非零为错误编号，message：结果描述
    '</result>
    ''发药完成
    '返回字符0或非0
    
    '上传处方明细、取药通知/准备发药 才会返回xml结构字符串
    If lngOper = gType.IntDetail Or lngOper = gType.IntStartList Then
        '解析返回数据
        Set objXML = New clsXML
        If objXML.OpenXMLDocument(strOutXML) = False Then
            strOut_Msg = "HisTransData_DIH：创建“MSXML2.DOMDocument”失败！"
            If gblnShowMsg Then
                MsgBox strOut_Msg, vbInformation + vbOKOnly, GSTR_MESSAGE
            Else
                strReturn = strOut_Msg
            End If
            
            Call OutputLog(strOutput & vbNewLine & strOut_Msg & vbNewLine)
            Exit Function
        End If
        
        '获取code值
        strOut_RETCODE = objXML.GetXMLNodePropertyValue("status", "code")
        
        '获取message值
        strOut_Msg = objXML.GetXMLNodePropertyValue("status", "message")
        
        If lngOper = gType.IntDetail Then
            '获取返回发药窗口号
            strOut_WinNo = objXML.GetXMLNodePropertyValue("result/value", "windowNo")
        End If
        
        '释放XML对象
        If Not objXML Is Nothing Then
            objXML.CloseXMLDocument
            Set objXML = Nothing
        End If
    
        strOutPutExeStep = "解析返回参数完成"
    Else
        strOut_RETCODE = Trim(strOutXML)
    End If
           
    '返回0表示接口调用成功，其他值为不成功
    If strOut_RETCODE <> "0" Then
        If gblnShowMsg Then
            MsgBox strOut_Msg, vbInformation + vbOKOnly, GSTR_MESSAGE
        Else
            strReturn = strOut_Msg
        End If
        
        strOutput = strOutput & vbCrLf & "上传数据失败！" & vbCrLf & strInputXML
        strOutput = strOutput & vbCrLf & "错误信息：" & strOut_Msg
        strOutput = strOutput & vbCrLf & "执行失败：HisTransData_DIH"
        Call OutputLog(strOutput)

        Exit Function
    End If

    strOutput = strOutput & vbCrLf & "解析code：" & strOut_RETCODE
    strOutput = strOutput & vbCrLf & "解析message：" & strOut_Msg
    
    If lngOper = gType.IntDetail Then
        strOutput = strOutput & vbCrLf & "解析windowno：" & strOut_WinNo
        If Not SetSendWin(LngStockID, strNO, Val(strOut_WinNo)) Then
            If gblnShowMsg Then
                MsgBox "调整处方的发药窗口失败！", vbCritical, GSTR_MESSAGE
            Else
                strReturn = "调整处方的发药窗口失败！"
            End If

            strOutput = strOutput & vbCrLf & "调整处方的发药窗口失败！"
            Call OutputLog(strOutput)
            Exit Function
        End If
    End If
            
    HisTransData_DIH = True
        
    strOutput = strOutput & vbCrLf & "执行成功：HisTransData_DIH"
    Call OutputLog(strOutput)
    
    Exit Function
    
errHandle:
    strOutput = strOutput & vbCrLf & "发生异常错误：" & Err.Description
    
    If gblnShowMsg Then
        If gintMode = 0 Then
            If gobjComLib.ErrCenter = 1 Then Resume
            Call gobjComLib.SaveErrLog
        Else
            MsgBox Err.Description, vbInformation, GSTR_SYSNAME
        End If
    End If
    
    strOutput = strOutput & vbCrLf & "接口参数：lngOper=" & lngOper & " LngStockID=" & LngStockID & " strNO=" & strNO
    strOutput = strOutput & vbCrLf & "最后步骤：" & strOutPutExeStep
    strOutput = strOutput & vbCrLf & "执行失败：HisTransData_DIH"
    Call OutputLog(strOutput)
End Function


