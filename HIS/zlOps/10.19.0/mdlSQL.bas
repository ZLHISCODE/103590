Attribute VB_Name = "mdlSQL"
Option Explicit

'######################################################################################################################

Public Enum SQL

    病人基本信息
    手术部门清单
    
    疾病诊断选择
    疾病诊断过滤

    疾病编码选择
    疾病编码过滤
    
    疾病诊断对照
    
    麻醉方式选择
    麻醉方式过滤
    
    执行房间选择
    
    正在手术记录
    等待审核手术
    手术申请记录
    
    手术项目选择
    手术项目过滤
    
    治疗项目选择
    治疗项目过滤
    收费执行科室
    
    麻醉药品选择
    麻醉药品过滤
    
    药品项目选择
    药品项目过滤
    
    材料项目选择
    材料项目过滤
    
    人员信息选择
    人员信息过滤
    
    科室信息选择
    科室信息过滤
    
    人员安排选择
    人员安排过滤
    
    病人手术情况
    病人诊断记录
    
    库存数量检查
    出库检查方式
    
    方案用药参考
    方案材料参考
    方案治疗参考
    
    手术用药选择
    手术材料选择
    手术治疗选择
    手术执行科室
    
    临床部门记录
    合约单位选择
    合约单位过滤
    科室医生人员
    人员过滤选择
End Enum

'######################################################################################################################

Public Function GetPublicSQL(ByVal intMenu As SQL, Optional ByVal strParam As String) As String
    '******************************************************************************************************************
    '功能:  集中产生SQL语句
    '参数:  strMenu             要产生的SQL名称
    '       strParam            参数串,格式:"参数值1'参数值2"
    '返回:  SQL语句
    '******************************************************************************************************************
    
    Dim strSQL As String
    Dim varParam As Variant
    Dim strTmp As String
    Dim rs As New ADODB.Recordset

    Dim lng发送号 As Long
    Dim str费别 As String
            
    On Error GoTo errHand
    
    If strParam = "" Then strParam = "'"
    
    varParam = Split(strParam, "'")
    
    Select Case intMenu
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.病人基本信息
        
            strSQL = "SELECT A.ID," & _
                     "A.体检号," & _
                     "D.门诊号,D.健康号,D.就诊卡号,D.IC卡号," & _
                     "D.姓名," & _
                     "D.性别," & _
                     "D.年龄," & _
                     "D.婚姻状况," & _
                     "A.体检时间," & _
                     "C.体检病历id,C.复查时间,C.随访期限,A.体检类型,D.联系人电话,D.工作单位," & _
                     "B.名称 AS 团体名称 " & _
                "FROM 体检登记记录 A,合约单位 B,体检人员档案 C,病人信息 D " & _
                "WHERE A.ID=C.登记id AND A.合约单位ID=B.ID(+) AND D.病人id=C.病人id AND C.ID=[1] "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.手术部门清单
        
            If strParam = "所有" Then
            
                strSQL = "SELECT A.编码,A.名称,ID FROM 部门表 A,部门性质说明 B WHERE (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质='手术' ORDER BY A.编码||'-'||A.名称"
            
            Else
                strSQL = "SELECT A.编码,A.名称,ID FROM 部门表 A,部门性质说明 B WHERE (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质='手术' " & _
                            "AND A.ID IN (SELECT 部门id FROM 部门人员 WHERE 人员id=[1])  ORDER BY A.编码||'-'||A.名称"
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.疾病诊断选择
            
            strSQL = "SELECT ID," & _
                        "上级ID," & _
                        "0 AS 末级," & _
                        "编码," & _
                        "名称 " & _
                "FROM 疾病诊断分类 " & _
                "START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "B.分类id AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称 " & _
                "FROM 疾病诊断目录 A,疾病诊断属类 B " & _
                "WHERE A.ID=B.诊断ID"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.疾病诊断过滤

            If Val(varParam(0)) = 1 Then
                '是全数字，按编码查找
                
                strSQL = "SELECT A.编码,A.名称,A.ID " & _
                            "FROM 疾病诊断目录 A " & _
                            "Where A.类别 = 1 " & _
                                  "AND A.编码 LIKE [1]"
                
            ElseIf Val(varParam(0)) = 2 Then
                '是全字母，按简码查找
                
                strSQL = "SELECT A.编码,A.名称,A.ID " & _
                            "FROM 疾病诊断目录 A " & _
                            "Where A.类别 = 1 " & _
                                  "And A.id IN (SELECT B.诊断id FROM 疾病诊断别名 B WHERE 简码 LIKE [2])"
                
            Else
                
                strSQL = "SELECT A.编码,A.名称,A.ID " & _
                            "FROM 疾病诊断目录 A " & _
                            "Where A.类别 = 1 " & _
                                  "AND ((编码 LIKE [1] OR 名称 LIKE [2]) " & _
                                  "OR A.id IN (SELECT B.诊断id FROM 疾病诊断别名 B WHERE (名称 LIKE [2] OR 简码 LIKE [2])))"
                
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.疾病编码选择
            
            strSQL = "SELECT ID," & _
                        "上级ID," & _
                        "0 AS 末级," & _
                        "NULL AS 编码," & _
                        "名称," & _
                        "NULL AS 简码,Null As 附码 " & _
                "FROM 疾病编码分类 " & _
                "WHERE 类别='D' " & _
                "START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.分类id AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称, " & _
                        "A.简码,a.附码 " & _
                "FROM 疾病编码目录 A " & _
                "WHERE 类别=[1] " & _
                    "AND DECODE(性别限制,'男',1,'女',2,0) IN (0,1,2) "
                            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.疾病编码过滤
        
            strSQL = "SELECT   编码," & _
                               "名称," & _
                               "简码," & _
                               "附码," & _
                               "ID " & _
                        "FROM 疾病编码目录 " & _
                        "WHERE 类别=[3] " & _
                            "AND DECODE(性别限制,'男',1,'女',2,0) IN (0,1,2) "
                            
            If Val(varParam(0)) = 1 Then
                '是全数字，按编码查找
                
                strSQL = strSQL & " And 编码 LIKE [1] "
                
            ElseIf Val(varParam(0)) = 2 Then
                '是全字母，按简码查找
                
                strSQL = strSQL & " And 简码 LIKE [2] "
                
            Else
                
                strSQL = strSQL & "AND (编码 LIKE [1] OR 名称 LIKE [2] OR 简码 LIKE [2])"
                
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.疾病诊断对照
        
            strSQL = "SELECT A.疾病ID,A.诊断ID,B.名称 AS 疾病编码,C.名称 AS 疾病诊断 " & _
                "FROM 疾病诊断对照 A,疾病编码目录 B,疾病诊断目录 C " & _
                "WHERE A.疾病ID=B.ID AND A.诊断ID=C.ID AND (A.疾病ID=[1] OR A.诊断ID=[2])"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.正在手术记录
            
            strTmp = ""
            
            If Trim(Split(strParam, ";")(2)) <> "" Then strTmp = strTmp & " AND b.姓名 LIKE [2] "
            If Trim(Split(strParam, ";")(3)) <> "" Then strTmp = strTmp & " AND b.住院号 = [3] "
            If Trim(Split(strParam, ";")(4)) <> "" Then strTmp = strTmp & " AND b.当前床号 = [4] "
            If Trim(Split(strParam, ";")(5)) <> "" Then strTmp = strTmp & " AND b.门诊号 = [5] "
            If Val(Trim(Split(strParam, ";")(7))) > 0 Then strTmp = strTmp & " AND a.诊疗项目ID = [6] "
            
            strSQL = "Select  e.ID,Decode(e.手术状态,1,'审核',2,'安排',3,'手术',4,'完成','申请') As 图标,a.Id As 医嘱id," & vbNewLine & _
                        "       Decode(a.紧急标志,1,'紧急','') As 紧急标志," & vbNewLine & _
                        "       DECODE(a.病人来源,1,'门诊',2,'住院','外来') AS 病人来源," & vbNewLine & _
                        "       Decode(a.诊疗项目id,Null,a.医嘱内容,f.名称) As 医嘱内容," & vbNewLine & _
                        "       a.开嘱时间," & vbNewLine & _
                        "       b.姓名," & vbNewLine & _
                        "       b.门诊号," & vbNewLine & _
                        "       b.住院号,b.当前床号 As 床号," & vbNewLine & _
                        "       c.名称 As 病人科室," & vbNewLine & _
                        "       d.名称 As 开单科室," & vbNewLine & _
                        "       a.开嘱医生 As 开单人," & vbNewLine & _
                        "       a.医嘱状态," & vbNewLine & _
                        "       a.病人id," & vbNewLine & _
                        "       a.主页id," & vbNewLine & _
                        "       a.诊疗项目id," & vbNewLine & _
                        "       e.手术状态,g.发送号,g.执行状态,a.挂号单,0 As 状态,b.出院时间 As 出院日期,b.当前病区id,b.当前科室id,b.IC卡号,b.身份证号 "
            strSQL = strSQL & _
                        "From 病人医嘱记录 a,病人医嘱发送 g, " & vbNewLine & _
                        "     病人信息 b," & vbNewLine & _
                        "     部门表 c," & vbNewLine & _
                        "     部门表 d," & vbNewLine & _
                        "     病人手术记录 e,诊疗项目目录 f " & vbNewLine & _
                        "Where Nvl(a.诊疗类别,'F')='F' " & vbNewLine & _
                        "      And a.相关id Is Null" & vbNewLine & _
                        "      And a.医嘱状态<>4 " & strTmp & vbNewLine & _
                        "      And a.执行科室id+0=[1]" & vbNewLine & _
                        "      And b.病人id=a.病人id" & vbNewLine & _
                        "      And c.Id=a.病人科室id" & vbNewLine & _
                        "      And d.Id=a.开嘱科室id And f.Id(+)=a.诊疗项目id " & vbNewLine & _
                        "      And a.Id=e.医嘱id  And e.手术状态=3 And a.ID=g.医嘱id(+) And (e.手术间=[7] Or [7] Is Null) "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.手术申请记录

            strTmp = ""
            
            If Val(Trim(Split(strParam, ";")(10))) = 1 Then
                If Split(strParam, ";")(0) <> "" Then strTmp = " AND e.手术结束时间 BETWEEN [3] AND [4] "
            Else
                If Split(strParam, ";")(0) <> "" Then strTmp = " AND a.开始执行时间 BETWEEN [3] AND [4] "
            End If
            If Trim(Split(strParam, ";")(2)) <> "" Then strTmp = strTmp & " AND b.姓名 LIKE [5] "
            If Trim(Split(strParam, ";")(3)) <> "" Then strTmp = strTmp & " AND b.住院号 = [6] "
            If Trim(Split(strParam, ";")(4)) <> "" Then strTmp = strTmp & " AND b.当前床号 = [7] "
            If Trim(Split(strParam, ";")(5)) <> "" Then strTmp = strTmp & " AND b.门诊号 = [8] "
            If Val(Trim(Split(strParam, ";")(7))) > 0 Then strTmp = strTmp & " AND a.诊疗项目ID = [9] "
                        
            
            strSQL = "Select  /*+rule*/ e.ID,Decode(e.手术状态,1,'审核',2,'安排',3,'手术',4,'完成','申请') As 图标,a.Id As 医嘱id," & vbNewLine & _
                        "       Decode(a.紧急标志,1,'紧急','') As 紧急标志," & vbNewLine & _
                        "       DECODE(a.病人来源,1,'门诊',2,'住院','外来') AS 病人来源," & vbNewLine & _
                        "       Decode(a.诊疗项目id,Null,a.医嘱内容,f.名称) As 医嘱内容," & vbNewLine & _
                        "       a.开嘱时间," & vbNewLine & _
                        "       b.姓名," & vbNewLine & _
                        "       b.门诊号," & vbNewLine & _
                        "       b.住院号,b.当前床号 As 床号," & vbNewLine & _
                        "       c.名称 As 病人科室," & vbNewLine & _
                        "       d.名称 As 开单科室," & vbNewLine & _
                        "       a.开嘱医生 As 开单人," & vbNewLine & _
                        "       a.医嘱状态," & vbNewLine & _
                        "       a.病人id," & vbNewLine & _
                        "       a.主页id," & vbNewLine & _
                        "       a.诊疗项目id," & vbNewLine & _
                        "       e.手术状态,g.发送号,g.执行状态,a.挂号单,0 As 状态,b.出院时间 As 出院日期,b.当前病区id,b.当前科室id,b.IC卡号,b.身份证号 "
            strSQL = strSQL & _
                        "From 病人医嘱记录 a,病人医嘱发送 g, " & vbNewLine & _
                        "     病人信息 b," & vbNewLine & _
                        "     部门表 c," & vbNewLine & _
                        "     部门表 d," & vbNewLine & _
                        "     病人手术记录 e,诊疗项目目录 f " & vbNewLine & _
                        "Where Nvl(a.诊疗类别,'F')='F' " & vbNewLine & _
                        "      And a.相关id Is Null" & vbNewLine & _
                        "      And a.医嘱状态<>4 " & strTmp & vbNewLine & _
                        "      And a.执行科室id+0=[1]" & vbNewLine & _
                        "      And b.病人id=a.病人id" & vbNewLine & _
                        "      And c.Id=a.病人科室id" & vbNewLine & _
                        "      And d.Id=a.开嘱科室id And f.Id(+)=a.诊疗项目id " & vbNewLine & _
                        "      And a.Id=e.医嘱id(+)  And e.手术状态=[2] And a.ID=g.医嘱id "

    '------------------------------------------------------------------------------------------------------------------
    Case SQL.手术项目选择
    
        strSQL = "SELECT DISTINCT ID," & _
                        "上级ID," & _
                        "0 AS 末级," & _
                        "编码," & _
                        "名称," & _
                        "NULL AS 单位 " & _
                "FROM 诊疗分类目录 " & _
                "START WITH ID IN (SELECT DISTINCT 分类id FROM 诊疗项目目录 WHERE 类别 = 'F' AND 服务对象 IN (1, 2, 3) AND (撤档时间 = TO_DATE('30000101', 'YYYYMMDD') OR 撤档时间 IS NULL)) CONNECT BY PRIOR 上级ID=ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.分类id AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称, " & _
                        "A.计算单位 AS 单位 " & _
                "FROM 诊疗项目目录 A "
        strSQL = strSQL & _
                "WHERE (撤档时间 = TO_DATE('30000101', 'YYYYMMDD') OR 撤档时间 IS NULL) " & _
                    "AND 服务对象 IN (1, 2, 3) " & _
                    "AND 类别 = 'F'"
        strSQL = "SELECT * FROM (" & strSQL & ") ORDER BY 编码"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.手术项目过滤
        
        Select Case Val(varParam(0))
        Case 1
            '是全数字，按编码查找
            strSQL = "Select a.ID,a.编码,a.名称 " & _
                        "From 诊疗项目目录 a " & _
                        "Where a.类别 = 'F' " & _
                            "And (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 Is Null) " & _
                            "And a.编码 Like [1]"
                    
        Case 2
            '是全字母，按简码查找
            strSQL = "Select Distinct a.ID,a.编码,a.名称 " & _
                        "From    诊疗项目目录 a," & _
                                "诊疗项目别名 b " & _
                        "Where   a.类别 = 'F' " & _
                            "And (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 Is Null) " & _
                            "And a.ID=b.诊疗项目id " & _
                            "And 编码_In Is Not Null " & _
                            "And b.简码 Like [2] "
        Case Else
            strSQL = "Select Distinct a.ID,a.编码,a.名称  " & _
                        "From    诊疗项目目录 a, " & _
                                "诊疗项目别名 b " & _
                        "Where   a.类别 = 'F' " & _
                            "And (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 Is Null) " & _
                            "And a.ID=b.诊疗项目id " & _
                            "And (a.编码 Like [1] Or a.名称 Like [2] Or b.名称 Like [2] Or b.简码 Like [2])"
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.治疗项目选择
        
'        strSQL = "Select *" & vbNewLine & _
'                    "From (Select ID,上级ID,0 As 末级,编码,名称 ,'' As 单位,'' As 类别,'' 类别编码,'' As 单价,'' As 规格" & vbNewLine & _
'                    "     From 收费分类目录" & vbNewLine & _
'                    "     Start With 上级ID Is Null Connect by Prior ID = 上级ID" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select -1 As ID,Null+0 As 上级ID,0 As 末级,'-1' As 编码,'西成药' As 名称 ,'' As 单位,'' As 类别,'' 类别编码,'' As 单价,'' As 规格 From Dual" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select -2 As ID,Null+0 As 上级ID,0 as 末级,'-2' As 编码,'中成药' As 名称 ,'' As 单位,'' As 类别,'' 类别编码,'' As 单价,'' As 规格 from Dual" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select -3 As ID,Null+0 As 上级ID,0 as 末级,'-3' As 编码,'中草药' As 名称 ,'' As 单位,'' As 类别,'' 类别编码,'' As 单价,'' As 规格 from Dual" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select -7 As ID,Null+0 As 上级ID,0 as 末级,'-7' As 编码,'卫生材料' As 名称 ,'' As 单位,'' As 类别,'' 类别编码,'' As 单价,'' As 规格 from Dual" & vbNewLine & _
'                    "     Union All" & vbNewLine & _
'                    "     Select a.ID,Decode(a.类别,'5',-1,'6',-2,'7',-3,'4',-7,a.分类id) As 上级ID,1 As 末级, a.编码,a.名称,a.计算单位 As 单位,b.名称 As 类别,a.类别 As 类别编码,Trim(To_Char(c.单价,'9999999999999.00')) As 单价,a.规格" & vbNewLine & _
'                    "     From  收费项目目录 a," & vbNewLine & _
'                    "          收费项目类别 b," & vbNewLine & _
'                    "          (Select 收费细目id,Sum(现价) As 单价 From 收费价目 Where 执行日期<=Sysdate And (终止日期 Is Null Or 终止日期>Sysdate) Group by 收费细目id) c" & vbNewLine & _
'                    "     Where c.收费细目id(+)=a.ID" & vbNewLine & _
'                    "            And Nvl(a.是否变价,0)=0" & vbNewLine & _
'                    "            And a.类别=b.编码" & vbNewLine & _
'                    "            And (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 Is Null)) a" & vbNewLine & _
'                    "Order By a.末级, a.编码"

        strSQL = "Select *" & vbNewLine & _
                    "From (Select ID,上级ID,0 As 末级,编码,名称 ,'' As 单位,'' As 类别,'' 类别编码,'' As 规格" & vbNewLine & _
                    "     From 收费分类目录" & vbNewLine & _
                    "     Start With 上级ID Is Null Connect by Prior ID = 上级ID" & vbNewLine & _
                    "     Union All" & vbNewLine & _
                    "     Select a.ID,a.分类id As 上级ID,1 As 末级, a.编码,a.名称,a.计算单位 As 单位,b.名称 As 类别,a.类别 As 类别编码,a.规格" & vbNewLine & _
                    "     From  收费项目目录 a," & vbNewLine & _
                    "          收费项目类别 b " & vbNewLine & _
                    "     Where Nvl(a.是否变价,0)=0" & vbNewLine & _
                    "            And a.类别=b.编码 And a.类别 Not In ('5','6','7','4') " & vbNewLine & _
                    "            And (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 Is Null)) a" & vbNewLine & _
                    "Order By a.末级, a.编码"
                    
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.治疗项目过滤
        'And a.类别 Not In ('5','6','7','4')
        strSQL = "Select a.ID,a.编码,a.名称,a.计算单位 As 单位,b.名称 As 类别,a.规格" & vbNewLine & _
                    "From   收费项目目录 a," & vbNewLine & _
                    "   收费项目类别 b " & vbNewLine & _
                    "Where  Nvl(a.是否变价,0)=0" & vbNewLine & _
                    "   And a.类别=b.编码  And a.类别 Not In ('5','6','7','4') " & vbNewLine & _
                    "   And (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 Is Null) "
                    
        Select Case Val(varParam(0))
        Case 1                  '是全数字，按编码查找
            strSQL = strSQL & " And a.编码 Like [1] "
        Case 2                  '是全字母，按简码查找
            strSQL = strSQL & " And Exists (Select 1 From 收费项目别名 bb Where (bb.名称 Like [2] Or bb.简码 Like [2]) And a.ID=bb.收费细目id) "
        Case Else               '数字字母混合
            strSQL = strSQL & " And (a.编码 Like [1] Or a.名称 Like [2] Or Exists (Select 1 From 收费项目别名 bb Where (bb.名称 Like [2] Or bb.简码 Like [2]) And a.ID=bb.收费细目id)) "
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.麻醉药品选择
        
        strSQL = "SELECT 序号,DECODE(药品ID,0,NULL,药品ID) AS 药品ID,ID,规格,DECODE(上级ID,0,NULL,上级ID) AS 上级ID,末级,编码,名称,单位,剂型,类别 " & _
                     "FROM (SELECT DECODE(编码, '5', 'A0', '6', 'B0', 'C0') AS 序号, " & _
                                  "0 AS 药品ID, " & _
                                  "DECODE(编码, '5', -1, '6', -2, -3) AS ID, " & _
                                  "0 AS 上级ID, " & _
                                  "0 AS 末级, " & _
                                  "编码, " & _
                                  "名称, NULL AS 单位," & _
                                  "NULL AS 规格," & _
                                  "NULL AS 剂型,'' As 类别 " & _
                             "From 诊疗项目类别 WHERE 编码 IN ('5','6','7') " & _
                           "Union All " & _
                             "SELECT DISTINCT DECODE(类型, 1, 'A', 2, 'B', 'C') || 编码 AS 序号, " & _
                                    "0 AS 药品ID, " & _
                                    "ID, " & _
                                    "DECODE(上级ID,NULL,DECODE(类型, 1, -1, 2, -2, -3),上级ID) AS 上级ID, " & _
                                    "0 AS 末级, " & _
                                    "编码, " & _
                                    "名称, NULL AS 单位," & _
                                    "NULL AS 规格," & _
                                    "NULL AS 剂型,'' As 类别 " & _
                               "From 诊疗分类目录  where DECODE(类型,1,'5',2,'6','7') IN ('5','6','7') "
        strSQL = strSQL & _
                            "Start With ID IN (SELECT Y.分类id FROM 诊疗项目目录 Y,药品特性 X WHERE X.药名id=Y.ID AND X.毒理分类='麻醉药') " & _
                             "Connect by Prior 上级ID = ID " & _
                             "Union All " & _
                                "SELECT 'D1' AS 序号, " & _
                                      "B.药品ID, " & _
                                      "B.药名ID AS ID, " & _
                                      "C.分类ID AS 上级ID," & _
                                      "1 as 末级, " & _
                                      "D.编码, " & _
                                      "D.名称, " & _
                                      "d.计算单位 AS 单位, " & _
                                      "D.规格," & _
                                      "A.药品剂型 AS 剂型,D.类别 " & _
                                 "FROM 药品特性 A,药品规格 B,诊疗项目目录 C,收费项目目录 D " & _
                                "WHERE A.药名id=B.药名id " & _
                                        "AND C.ID=A.药名id " & _
                                        "AND D.ID=B.药品id " & _
                                        "AND C.类别 IN ('5','6','7') " & _
                                        "AND A.毒理分类='麻醉药'" & _
                                        "AND (D.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or D.撤档时间 is NULL) " & _
                           ") " & _
                    "ORDER BY 末级,序号"
                        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.麻醉药品过滤
        
        strSQL = "SELECT DECODE(C.类别, '5', '西成药', '6', '中成药', '中草药') AS 分类," & _
                   "D.编码," & _
                   "D.名称," & _
                   "D.规格," & _
                   "A.药品剂型 As 剂型," & _
                   "d.计算单位 As 单位," & _
                   "B.药品ID," & _
                   "B.药品ID As ID," & _
                   "A.药名ID,d.类别 " & _
             "FROM 药品特性 A, 药品规格 B, 诊疗项目目录 C, 收费项目目录 D " & _
             "WHERE A.药名ID = B.药名ID " & _
                   "AND C.ID = A.药名ID " & _
                   "AND D.ID = B.药品ID " & _
                   "AND A.毒理分类 = '麻醉药' " & _
                   "AND C.类别 IN ('5','6','7') " & _
                   "AND (D.撤档时间 IS NULL OR D.撤档时间 = TO_DATE('3000-01-01', 'yyyy-MM-dd')) "
                       
        Select Case Val(varParam(0))
        Case 1                          '是全数字，按编码查找
            
            strSQL = strSQL & _
                       "AND D.编码 LIKE [1] "
            
        Case 2                          '是全字母，按简码查找
        
            strSQL = strSQL & _
                       "AND Exists (SELECT 1 FROM 收费项目别名 bb WHERE (bb.名称 Like [2] Or bb.简码 LIKE [2]) And B.药品id=bb.收费细目ID) "
            
        Case Else
        
            strSQL = strSQL & _
                       "AND (D.编码 LIKE [1] OR D.名称 LIKE [2] OR Exists (SELECT 1 FROM 收费项目别名 bb WHERE (bb.名称 LIKE [2] OR bb.简码 LIKE [2]) And B.药品id=bb.收费细目ID )) "
            
        End Select
        strSQL = strSQL & " ORDER BY D.编码,D.名称"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.药品项目选择
        
        strSQL = "SELECT 序号,DECODE(药品ID,0,NULL,药品ID) AS 药品ID,ID,规格,DECODE(上级ID,0,NULL,上级ID) AS 上级ID,末级,编码,名称,单位,剂型,类别 " & _
                     "FROM (SELECT DECODE(编码, '5', 'A1', '6', 'A2', 'A3') AS 序号, " & _
                                  "0 AS 药品ID, " & _
                                  "DECODE(编码, '5', -1, '6', -2, -3) AS ID, " & _
                                  "0 AS 上级ID, " & _
                                  "0 AS 末级, " & _
                                  "编码, " & _
                                  "名称, NULL AS 单位," & _
                                  "NULL AS 规格," & _
                                  "NULL AS 剂型,'' As 类别 " & _
                             "From 诊疗项目类别 WHERE 编码 IN ('5','6','7') " & _
                           "Union All " & _
                             "SELECT DECODE(类型, 1, 'A1', 2, 'A2', 'A3') || TO_CHAR(ROWNUM,'0000000000') AS 序号, " & _
                                    "0 AS 药品ID, " & _
                                    "ID, " & _
                                    "DECODE(上级ID,NULL,DECODE(类型, 1, -1, 2, -2, -3),上级ID) AS 上级ID, " & _
                                    "0 AS 末级, " & _
                                    "编码, " & _
                                    "名称, NULL AS 单位," & _
                                    "NULL AS 规格," & _
                                    "NULL AS 剂型,'' As 类别 " & _
                               "From 诊疗分类目录  Where DECODE(类型,1,'5',2,'6','7') IN ('5','6','7') "
        strSQL = strSQL & _
                            "Start With 上级ID is NULL " & _
                             "Connect by Prior ID = 上级ID " & _
                             "Union All " & _
                                "SELECT 'B1' AS 序号, " & _
                                      "B.药品ID, " & _
                                      "B.药名ID AS ID, " & _
                                      "C.分类ID AS 上级ID," & _
                                      "1 as 末级, " & _
                                      "D.编码, " & _
                                      "D.名称, " & _
                                      "d.计算单位 AS 单位, " & _
                                      "D.规格," & _
                                      "A.药品剂型 AS 剂型,d.类别 " & _
                                 "FROM 药品特性 A,药品规格 B,诊疗项目目录 C,收费项目目录 D " & _
                                "WHERE A.药名id=B.药名id " & _
                                        "AND C.ID=A.药名id " & _
                                        "AND D.ID=B.药品id " & _
                                        "AND C.类别 IN ('5','6','7') " & _
                                        "AND (D.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or D.撤档时间 is NULL) " & _
                           ") " & _
                    "ORDER BY 序号"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.药品项目过滤
    
        strSQL = "SELECT Decode(C.类别, '5', '西成药', '6', '中成药', '中草药') AS 分类," & _
                       "D.编码," & _
                       "D.名称," & _
                       "D.规格," & _
                       "A.药品剂型 AS 剂型," & _
                       "d.计算单位 AS 单位," & _
                       "B.药品ID," & _
                       "B.药品ID AS ID," & _
                       "A.药名ID,d.类别 " & _
                 "FROM 药品特性 A, 药品规格 B, 诊疗项目目录 C, 收费项目目录 D " & _
                 "WHERE A.药名ID = B.药名ID " & _
                       "AND C.ID = A.药名ID " & _
                       "AND D.ID = B.药品ID " & _
                       "AND C.类别 IN ('5','6','7') " & _
                       "AND (D.撤档时间 IS NULL OR D.撤档时间 = TO_DATE('3000-01-01', 'yyyy-MM-dd')) "
                       
        Select Case Val(varParam(0))
        Case 1                          '是全数字，按编码查找
            strSQL = strSQL & _
                       "AND D.编码 LIKE [1] "
        Case 2                          '是全字母，按简码查找
            strSQL = strSQL & _
                       "AND Exists (SELECT 1 FROM 收费项目别名 bb WHERE (bb.名称 Like [2] Or bb.简码 LIKE [2]) And B.药品id=bb.收费细目ID) "
        Case Else
            strSQL = strSQL & _
                       "AND (D.编码 LIKE [1] OR D.名称 LIKE [2] OR Exists (SELECT 1 FROM 收费项目别名 bb WHERE (bb.名称 Like [2] Or bb.简码 LIKE [2]) And B.药品id=bb.收费细目ID)) "
            
        End Select
        strSQL = strSQL & " ORDER BY D.编码,D.名称"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.材料项目选择
    
        strSQL = "Select ID," & _
                    "上级ID," & _
                    "0 as 末级," & _
                    "编码," & _
                    "名称," & _
                    "'' as 规格,'' As 产地," & _
                    "'' as 单位," & _
                    "0 AS 是否变价,0 As 最低现价,0 As 最高现价,'' As 单价 " & _
              "From 诊疗分类目录 " & _
              "where 类型=7 " & _
              "Start With 上级ID is NULL " & _
                "Connect by Prior ID = 上级ID " & _
                "Union All "
        strSQL = strSQL & _
                  "Select A.材料ID AS ID, " & _
                     "C.分类id AS 上级ID, " & _
                     "1 as 末级, " & _
                     "B.编码, " & _
                     "B.名称, " & _
                     "B.规格,B.产地, " & _
                     "B.计算单位 as 单位, " & _
                     "B.是否变价,D.原价 as 最低现价,D.现价 as 最高现价,DECODE(B.是否变价,1,TRIM(TO_CHAR(D.原价,'999999990.99'))||'～'||TRIM(TO_CHAR(D.现价,'999999990.99')),TRIM(TO_CHAR(D.现价,'999999990.99'))) as 单价 " & _
                "FROM 材料特性 A,收费项目目录 B,诊疗项目目录 C,收费价目 d  " & _
               "Where A.材料id=B.ID AND (B.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or B.撤档时间 is NULL) " & _
                    "AND C.ID=A.诊疗id And d.执行日期<=SYSDATE AND (d.终止日期>=SYSDATE OR d.终止日期 IS NULL)"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.材料项目过滤
            
        strSQL = "SELECT A.是否变价,B.原价 as 最低现价,B.现价 as 最高现价,C.名称 AS 类别,A.编码,A.名称,A.规格,A.产地,A.计算单位,DECODE(A.是否变价,1,TRIM(TO_CHAR(B.原价,'999999990.99'))||'～'||TRIM(TO_CHAR(B.现价,'999999990.99')),TRIM(TO_CHAR(B.现价,'999999990.99'))) as 单价,A.ID  " & _
                    "FROM 收费项目目录 A,收费价目 B,收费项目类别 C " & _
                    "WHERE C.编码=A.类别 " & _
                            "AND A.ID=B.收费细目ID " & _
                            "AND A.类别='4' " & _
                            "AND B.执行日期<=SYSDATE AND (B.终止日期>=SYSDATE OR B.终止日期 IS NULL) " & _
                            "AND (A.撤档时间 IS NULL OR A.撤档时间=TO_DATE('3000-01-01','yyyy-MM-dd'))"
                                
        Select Case Val(varParam(0))
        Case 1                          '是全数字，按编码查找
        
            strSQL = strSQL & " AND A.编码 Like [1] "
            
        Case 2                          '是全字母，按简码查找
            
            strSQL = strSQL & " AND Exists (SELECT 1 FROM 收费项目别名 bb WHERE (bb.名称 Like [2] Or bb.简码 LIKE [2]) And a.ID=bb.收费细目ID) "
        Case Else

            strSQL = strSQL & " AND (A.编码 Like [1] or A.名称 Like [2] Or Exists (SELECT 1 FROM 收费项目别名 bb WHERE (bb.名称 Like [2] Or bb.简码 LIKE [2]) And a.ID=bb.收费细目ID))"

        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.麻醉方式选择
        
        strSQL = "SELECT 名称,编码,计算单位 AS 单位,操作类型 AS 麻醉类型,ID FROM 诊疗项目目录 a WHERE 类别='G' "
    
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.麻醉方式过滤

        strSQL = "SELECT 名称,编码,计算单位 AS 单位,操作类型 AS 麻醉类型,ID FROM 诊疗项目目录 a WHERE 类别='G' "
                                                
        Select Case Val(varParam(0))
        Case 1                          '是全数字，按编码查找
        
            strSQL = strSQL & " AND A.编码 Like [1] "
            
        Case 2                          '是全字母，按简码查找
            
            strSQL = strSQL & " AND Exists (SELECT 1 FROM 诊疗项目别名 bb WHERE (bb.名称 Like [2] Or bb.简码 LIKE  [2]) And a.ID=bb.诊疗项目id) "

        Case Else

            strSQL = strSQL & " AND (A.编码 Like [1] or A.名称 Like [2] Or Exists (SELECT 1 FROM 诊疗项目别名 bb WHERE (bb.名称 Like [2] Or bb.简码 LIKE  [2]) And a.ID=bb.诊疗项目id))"

        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.人员信息选择
                
        strSQL = "SELECT   A.编号," & _
                           "A.姓名," & _
                           "A.简码," & _
                           "C.名称 AS 科室," & _
                           "A.ID " & _
                    "FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D " & _
                    "WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID AND D.缺省=1 " & _
                        "AND B.人员性质=[1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
        
        strSQL = strSQL & " Order By Decode(c.ID,[2],1,2) "
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.人员信息过滤
    
        strSQL = "SELECT   A.编号," & _
                           "A.姓名," & _
                           "A.简码," & _
                           "C.名称 AS 科室," & _
                           "A.ID " & _
                    "FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D " & _
                    "WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID AND D.缺省=1 " & _
                        "AND B.人员性质=[1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
        
        Select Case Val(varParam(0))
        Case 1                          '是全数字，按编码查找
        
            strSQL = strSQL & " AND A.编号 Like [3] "
            
        Case 2                          '是全字母，按简码查找
            
            strSQL = strSQL & " AND A.简码 LIKE  [4]) "

        Case Else

            strSQL = strSQL & " AND (A.编号 LIKE [3] OR A.姓名 LIKE [4] OR A.简码 LIKE [4]) "

        End Select
        strSQL = strSQL & " Order By Decode(c.ID,[2],1,2) "
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.科室信息选择
                
        
        strSQL = "SELECT a.编码,a.名称,a.简码,a.ID FROM 部门表 a,部门性质说明 b Where a.ID=b.部门id And b.工作性质=[1]"
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.科室信息过滤
    
        strSQL = "SELECT a.编码,a.名称,a.简码,a.ID FROM 部门表 a,部门性质说明 b Where a.ID=b.部门id And b.工作性质=[1]"
        
        
        Select Case Val(varParam(0))
        Case 1                          '是全数字，按编码查找
        
            strSQL = strSQL & " AND A.编码 Like [2] "
            
        Case 2                          '是全字母，按简码查找
            
            strSQL = strSQL & " AND A.简码 LIKE  [3]) "

        Case Else

            strSQL = strSQL & " AND (A.编码 LIKE [2] OR A.名称 LIKE [3] OR A.简码 LIKE [3]) "

        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.人员安排选择
            
        strSQL = "SELECT   A.编号," & _
                           "A.姓名," & _
                           "A.简码," & _
                           "C.名称 As 科室," & _
                           "Decode(e.人员id,Null,'空闲',Decode(e.手术状态,2,'预订',3,'在用')) As 状态," & _
                           "A.ID " & _
                    "FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D, " & _
                                        "(SELECT AA.人员id,bb.手术状态 " & _
                                         "FROM 病人手术人员 AA," & _
                                                "病人手术记录 BB, " & _
                                                "病人手术记录 DD " & _
                                        "WHERE AA.记录ID = BB.ID " & _
                                                "AND BB.医嘱id <> [3] " & _
                                                "AND BB.手术状态 In (2,3) " & _
                                                "AND DD.医嘱id = [3] " & _
                                                "AND NOT (DD.手术开始时间 > BB.手术结束时间 OR DD.手术结束时间 < BB.手术开始时间)) e " & _
                    "WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID AND D.缺省=1 " & _
                        "AND A.ID=e.人员id(+) And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                        "AND B.人员性质=[1] And ((b.人员性质='护士' And c.ID=[2]) Or b.人员性质<>'护士') " & _
                        "ORDER BY Decode(c.ID,[2],1,2)"
                        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.人员安排过滤
        
        strSQL = "SELECT   A.编号," & _
                           "A.姓名," & _
                           "A.简码," & _
                           "C.名称 AS 科室," & _
                           "Decode(e.人员id,Null,'空闲',Decode(e.手术状态,2,'预订',3,'在用')) As 状态," & _
                           "A.ID " & _
                    "FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D, " & _
                                        "(SELECT AA.人员id,bb.手术状态 " & _
                                         "FROM 病人手术人员 AA," & _
                                                "病人手术记录 BB, " & _
                                                "病人手术记录 DD " & _
                                        "WHERE AA.记录ID = BB.ID " & _
                                                "AND BB.医嘱id <> [3] " & _
                                                "AND BB.手术状态 In (2,3) " & _
                                                "AND DD.医嘱id = [3] " & _
                                                "AND NOT (DD.手术开始时间 > BB.手术结束时间 OR DD.手术结束时间 < BB.手术开始时间)) e " & _
                    "WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID AND D.缺省=1 AND A.ID=e.人员id(+) And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                        "AND b.人员性质=[1] And ((b.人员性质='护士' And c.ID=[2]) Or b.人员性质<>'护士')  "
                        
'        Select Case Val(varParam(0))
'        Case 1                          '是全数字，按编码查找
'
'            strSQL = strSQL & " AND A.编号 Like [4] "
'
'        Case 2                          '是全字母，按简码查找
'
'            strSQL = strSQL & " AND A.简码 LIKE  [5]) "
'
'        Case Else

            strSQL = strSQL & " AND (A.编号 LIKE [4] OR A.姓名 LIKE [5] OR A.简码 LIKE [5]) "

'        End Select
        strSQL = strSQL & " Order By Decode(c.ID,[2],1,2) "
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.病人手术情况
    
        strSQL = "SELECT DECODE(A.诊疗项目ID,null,'2-疾病','1-诊疗') AS 编码方式," & _
                        "A.手术名称," & _
                        "A.缺省," & _
                        "DECODE(A.诊疗项目id,Null,A.手术操作ID,A.诊疗项目id) As ID " & _
                    "FROM 病人手术情况 A " & _
                    "WHERE A.记录id=[1] " & _
                            "AND A.性质=[2] "
                            
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.病人诊断记录
        
        strSQL = "Select 1 AS ID,诊断id," & _
                         "疾病id," & _
                         "c.编码 As 诊断编码," & _
                         "d.编码 As 疾病编码," & _
                         "诊断描述 " & _
                    "From 病人诊断记录 a,病人手术记录 b,疾病诊断目录 c,疾病编码目录 d " & _
                   "where a.医嘱id = b.医嘱id and 诊断类型 = [2] And b.ID=[1] And c.ID(+)=a.诊断id And d.ID(+)=a.疾病id"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.收费执行科室
    
        Select Case varParam(0)
        Case "5", "6", "7", "4"
            strSQL = "Select a.ID,a.编码,a.名称" & vbNewLine & _
                        "From  部门表 a" & vbNewLine & _
                        "Where (a.撤档时间 IS NULL OR a.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD'))" & vbNewLine & _
                        "    And a.ID In (   Select  b.部门id" & vbNewLine & _
                        "            From    部门性质说明 b" & vbNewLine & _
                        "            Where   b.服务对象 in (1,2,3)" & vbNewLine & _
                        "                And b.工作性质=Decode('" & varParam(0) & "','5','西药房','6','成药房','7','中药房','4','发料部门')) " & _
                        " Order By Decode(a.ID,[1],0,1) "

        Case Else
            strSQL = "Select * From (" & vbNewLine & _
                        "Select  a.ID,a.编码,a.名称,1 As 末级,b.OrderCol" & vbNewLine & _
                        "From    部门表 a," & vbNewLine & _
                        "    (" & vbNewLine & _
                        "    Select a.ID,1 As OrderCol From 部门表 a,收费项目目录 X Where X.ID=[2] And X.执行科室=1 And A.ID=[3]" & vbNewLine & _
                        "    Union All" & vbNewLine & _
                        "    Select a.ID,2 As OrderCol From 部门表 a,病区科室对应 B,收费项目目录 X Where X.ID=[2] And X.执行科室=2 And A.ID=B.病区id And B.科室ID=[3]" & vbNewLine & _
                        "    Union All" & vbNewLine & _
                        "    Select a.ID,3 As OrderCol From 部门表 a,收费项目目录 X Where X.ID=[2] And X.执行科室=3 And A.ID=[4]" & vbNewLine & _
                        "    Union All" & vbNewLine & _
                        "    Select a.ID,4 As OrderCol From 部门表 a,收费执行科室 B,收费项目目录 X Where X.ID=[2] And X.执行科室=4 And A.ID=B.执行科室id And B.病人来源=1 And B.收费细目id=X.ID" & vbNewLine & _
                        "    ) b," & vbNewLine & _
                        "    部门性质说明 c" & vbNewLine & _
                        "Where a.ID = c.部门ID" & vbNewLine & _
                        "    And c.服务对象 In (1,2,3)" & vbNewLine & _
                        "    And a.ID=b.ID(+)" & vbNewLine & _
                        "Order By Decode(a.ID,[1],0,b.OrderCol)" & vbNewLine & _
                        ") "

        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.库存数量检查
        
        strSQL = "Select Nvl(Sum(A.可用数量),0) As 库存 From 药品库存 A Where (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate)) And A.性质=1 And A.药品ID=[1] And A.库房ID=[2]"
        
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.出库检查方式
        
        If varParam(0) = "4" Then
            strSQL = "Select 检查方式 From 材料出库检查 Where 库房ID=[1]"
        Else
            strSQL = "Select 检查方式 From 药品出库检查 Where 库房ID=[1]"
        End If
    
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.方案用药参考
        
        strSQL = "SELECT D.ID,d.类别,D.规格,B.计算单位 As 单位,A.类型 As 用药类型,D.名称 AS 药品名称,A.总量,A.总量 As 准备数量 " & _
                "FROM 方案用药参考 A,诊疗项目目录 B,药品规格 C,收费项目目录 D " & _
                "WHERE A.药名id=C.药品id AND B.ID=C.药名id AND D.ID=C.药品id AND A.方案ID=[1]"
                
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.方案材料参考

        strSQL = "SELECT B.计算单位 As 单位,B.ID,B.名称,A.数次,A.数次 As 准备数量,B.规格 " & _
                "FROM 方案材料参考 A,收费项目目录 B " & _
                "WHERE A.材料id=B.ID AND A.方案ID=[1]"

                
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.方案治疗参考
    
        strSQL = "SELECT D.ID,b.名称 As 类别,d.计算单位,D.名称,A.数量,d.类别 As 类别编码 " & _
                "FROM 方案附费参考 A,收费项目目录 D,收费项目类别 B " & _
                "WHERE A.细目id=D.ID And B.编码=d.类别 And A.方案ID=[1]"
                
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.手术用药选择
        
        
        strSQL = "Select    Null AS 类型," & _
                           "ID," & _
                           "编码," & _
                           "名称," & _
                           "Null AS 单位,NULL AS 规格," & _
                           "Null AS 总量," & _
                           "0 As 上级id," & _
                           "0 As 末级 " & _
                      "From 手术方案参考 " & _
                     "WHERE ID IN (Select 方案ID From 方案适用手术 A, 病人手术记录 B, 病人医嘱记录 C WHERE A.手术项目ID = C.诊疗项目id AND B.医嘱id = C.ID AND B.ID=[1]) " & _
                    "Union All "
                            
        strSQL = strSQL & _
                "Select Null AS 类型," & _
                           "ID," & _
                           "编码," & _
                           "名称," & _
                           "Null AS 单位,NULL AS 规格," & _
                           "Null AS 总量," & _
                           "0 As 上级id," & _
                           "0 As 末级 " & _
                      "From 手术方案参考 b " & _
                     "WHERE Not Exists (Select 1 From 方案适用手术 a Where a.方案id=b.ID) " & _
                    "Union All "
                    
        strSQL = strSQL & _
                 "SELECT A.类型, " & _
                 "      D.ID, " & _
                 "      d.编码, " & _
                 "      d.名称, " & _
                 "      d.计算单位 AS 单位, " & _
                 "      D.规格, " & _
                 "      TRIM(TO_CHAR(a.总量, '9999999990.99')) AS 总量, " & _
                 "      A.方案ID AS 上级id, " & _
                 "      1 AS 末级  " & _
                 " FROM 方案用药参考 A, 诊疗项目目录 B,药品规格 C,收费项目目录 D  " & _
                 " WHERE A.药名id = C.药品id AND C.药名id=B.id AND D.ID=C.药品ID"
                 
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.手术材料选择
        
        strSQL = "Select   ID," & _
                           "编码," & _
                           "名称," & _
                           "Null AS 单位,NULL AS 规格," & _
                           "Null AS 数量," & _
                           "0 As 上级id," & _
                           "0 As 末级 " & _
                      "From 手术方案参考 " & _
                     "WHERE ID IN (Select 方案ID From 方案适用手术 A, 病人手术记录 B, 病人医嘱记录 C WHERE A.手术项目ID = C.诊疗项目id AND B.医嘱id = C.ID AND B.ID=[1]) " & _
                    "Union All "
                            
        strSQL = strSQL & _
                "Select    ID," & _
                           "编码," & _
                           "名称," & _
                           "Null AS 单位,NULL AS 规格," & _
                           "Null AS 数量," & _
                           "0 As 上级id," & _
                           "0 As 末级 " & _
                      "From 手术方案参考 b " & _
                     "WHERE Not Exists (Select 1 From 方案适用手术 a Where a.方案id=b.ID) " & _
                    "Union All "
                    
        strSQL = strSQL & _
                      "SELECT ROWNUM AS ID," & _
                             "B.编码," & _
                             "B.名称," & _
                             "B.规格,B.计算单位 AS 单位," & _
                             "TRIM(TO_CHAR(A.数次, '9999999990.99')) AS 数量," & _
                             "方案ID AS 上级id," & _
                             "1 AS 末级 " & _
                        "FROM 方案材料参考 A, 收费项目目录 B " & _
                       "WHERE A.材料id = B.id"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.手术治疗选择
    
        strSQL = "Select   ID," & _
                           "编码," & _
                           "名称," & _
                           "Null AS 单位,NULL AS 规格," & _
                           "Null AS 数量," & _
                           "0 As 上级id," & _
                           "0 As 末级 " & _
                      "From 手术方案参考 " & _
                     "WHERE ID IN (Select 方案ID From 方案适用手术 A, 病人手术记录 B, 病人医嘱记录 C WHERE A.手术项目ID = C.诊疗项目id AND B.医嘱id = C.ID AND B.ID=[1]) " & _
                    "Union All "
                    
        strSQL = strSQL & _
                "Select    ID," & _
                           "编码," & _
                           "名称," & _
                           "Null AS 单位,NULL AS 规格," & _
                           "Null AS 数量," & _
                           "0 As 上级id," & _
                           "0 As 末级 " & _
                      "From 手术方案参考 b " & _
                     "WHERE Not Exists (Select 1 From 方案适用手术 a Where a.方案id=b.ID) " & _
                    "Union All "
                    
        strSQL = strSQL & _
                      "SELECT ROWNUM AS ID," & _
                             "B.编码," & _
                             "B.名称," & _
                             "B.规格,B.计算单位 AS 单位," & _
                             "TRIM(TO_CHAR(A.数量, '9999999990.99')) AS 数量," & _
                             "方案ID AS 上级id," & _
                             "1 AS 末级 " & _
                        "FROM 方案附费参考 A, 收费项目目录 B " & _
                       "WHERE A.细目ID = B.id"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.临床部门记录
        
        If varParam(0) = "所有" Then
        
            strSQL = "SELECT A.编码||'-'||A.名称 As 名称,ID FROM 部门表 A,部门性质说明 B WHERE (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质='临床' ORDER BY A.编码||'-'||A.名称"
        
        Else
            strSQL = "SELECT A.编码||'-'||A.名称 As 名称,ID FROM 部门表 A,部门性质说明 B WHERE (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质='临床' " & _
                        "AND A.ID IN (SELECT 部门id FROM 部门人员 WHERE 人员id=" & UserInfo.ID & ")  ORDER BY A.编码||'-'||A.名称"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.合约单位选择

        strSQL = "SELECT -1 AS ID,NULL+0 AS 上级id,'0' AS 编码,'所有' AS 名称,'' as 简码,'' as 地址,0 AS 末级,'' AS 联系人,'' AS 电话,'' AS 电子邮件,'' AS 开户银行,'' AS 帐号,'' AS 地址,'' AS 说明 from dual " & _
                    "Union All " & _
                    "SELECT ID,DECODE(上级id,NULL,-1,0,-1,上级id) AS 上级id,编码,名称,简码,地址,0 AS 末级,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位  where 末级<>1 " & _
                    "Start With 上级id is null connect by prior ID=上级id " & _
                    "Union All " & _
                    "SELECT ID,DECODE(上级id,NULL,-1,0,-1,上级id) AS 上级id,编码,名称,简码,地址,1 AS 末级,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位  where 末级=1"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.合约单位过滤
        
            strSQL = "select ID,编码,名称,简码,地址,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位  where 末级=1 " & _
                " AND (编码 Like [1] or 名称 Like [1] OR 简码 Like [1])"
                
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.科室医生人员
    
        strSQL = _
            "Select Distinct A.姓名,A.ID,B.部门ID,A.编号,Upper(A.简码) as 简码," & _
            " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
            " And C.人员性质 IN('医生') And B.部门ID=[1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
        strSQL = strSQL & " Order by 简码,人员性质 Desc"
    '------------------------------------------------------------------------------------------------------------------
    Case SQL.人员过滤选择
        
            strSQL = "SELECT C.病人id AS ID,C.当前病区id," & _
                    "C.姓名," & _
                    "C.门诊号," & _
                    "C.年龄," & _
                    "C.性别," & _
                    "C.出生日期," & _
                    "C.身份证号," & _
                    "C.婚姻状况,c.职业,c.住院号,c.国籍,c.民族,c.医疗付款方式,c.费别, " & _
                    "C.合同单位id,c.工作单位,c.联系人姓名,c.联系人电话,c.家庭地址,c.家庭电话,c.户口邮编,c.单位电话,c.单位邮编,b.主页id " & _
                "FROM 病人信息 C,病案主页 b " & _
                "WHERE c.病人id=b.病人id(+) And b.出院日期 Is Null " & IIf(strParam = "'", "", strParam)

            
    End Select
    
    GetPublicSQL = strSQL
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
End Function


