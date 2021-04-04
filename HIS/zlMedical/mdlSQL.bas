Attribute VB_Name = "mdlSQL"
Option Explicit

Public Enum SQL

    病人基本信息
    单位人员选择
    体检部门清单
    
'    疾病编码选择
'    疾病编码过滤
    
    收费项目选择
    收费项目过滤
    体检项目选择
    体检项目过滤选择
    体检项目清单
    体检人员档案
    体检人员档案_单个
    体检类型分类
    体检类型分类选择
    体检类型选择
    体检类型过滤选择
    体检类型项目
    体检类型计价
    病人所有项目
    体检诊断分类
    团体过滤选择
    诊治项目选择
    诊治项目过滤选择
    人员体检项目
    人员原始项目
    团体体检项目
    体检团体选择
    检查诊治项目过滤选择
    检查诊治项目选择
    团体未结明细
    
    个人费用概况
    团体费用概况
    
    体检项目价表
    体检预约单据
    体检登记单据
    诊疗执行科室
    收费执行科室
    药品执行科室
    人员档案
    体检组别人员
    体检组别人员1
    人员过滤选择
    体检人数统计
End Enum

Public Function GetPublicSQL(ByVal intMenu As SQL, Optional ByVal strParam As String, Optional ByVal blnMoveOuted As Boolean = False) As String
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
            
    On Error GoTo errHand
    
    If strParam = "" Then strParam = "'"
    
    varParam = Split(strParam, "'")
    
    Select Case intMenu
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.人员档案
            strSQL = "select A.病人id AS ID,A.病人id,A.姓名,A.身份证号 AS 身份证,A.性别,A.年龄,TO_CHAR(A.出生日期,'yyyy-mm-dd') AS 出生日期,A.婚姻状况," & _
                    "A.民族,A.国籍,A.学历,A.职业,A.身份,A.联系人姓名,A.联系人电话,A.联系人地址,A.工作单位 " & _
                    "from 病人信息 A " & _
                    "WHERE A.病人ID=[1]"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.人员过滤选择
            
            strSQL = "SELECT C.病人id AS ID," & _
                    "C.姓名," & _
                    "C.门诊号," & _
                    "C.健康号," & _
                    "C.年龄," & _
                    "C.性别," & _
                    "TO_CHAR(C.出生日期,'yyyy-mm-dd') AS 出生日期," & _
                    "C.身份证号," & _
                    "C.婚姻状况, " & _
                    "C.合同单位id " & _
                "FROM 病人信息 C " & _
                "WHERE 1=1 " & strParam
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.单位人员选择
        
            strSQL = "Select * From (SELECT 1 As 选择,C.病人id AS ID," & _
                    "C.姓名," & _
                    "C.门诊号," & _
                    "C.健康号," & _
                    "Decode(c.出生日期,Null,Decode(c.年龄,Null,0,Decode(Trim(Substr(c.年龄,Length(c.年龄),1)),'岁',Zl_To_Number(Substr(c.年龄,1,Length(c.年龄)-1)),'月',Zl_To_Number(Substr(c.年龄,1,Length(c.年龄)-1))/12,'天',Zl_To_Number(Substr(c.年龄,1,Length(c.年龄)-1))/365,Zl_To_Number(c.年龄))),Trunc(Months_between(Sysdate,c.出生日期)/12)) As 年龄," & _
                    "C.性别," & _
                    "TO_CHAR(C.出生日期,'yyyy-mm-dd') AS 出生日期," & _
                    "C.身份证号,民族,国籍,学历,职业,身份,联系人姓名,联系人电话,联系人地址,工作单位,IC卡号,就诊卡号," & _
                    "C.婚姻状况, " & _
                    "C.合同单位id " & _
                "FROM 病人信息 c Where 合同单位id In (Select ID From 合约单位 Start With ID=[1] Connect by Prior ID=上级id)) " & _
                "WHERE (Instr(性别,[2])>0 Or [2] Is Null) And 年龄 Between [3] And [4] "

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.病人基本信息
        
            strSQL = "SELECT A.ID," & _
                     "A.体检号," & _
                     "D.门诊号,D.健康号,D.就诊卡号,D.IC卡号," & _
                     "D.姓名," & _
                     "D.性别," & _
                     "D.年龄," & _
                     "D.婚姻状况," & _
                     "C.体检时间," & _
                     "C.体检病历id,C.复查时间,C.随访期限,A.体检类型,D.联系人电话,D.工作单位," & _
                     "B.名称 AS 团体名称 " & _
                "FROM 体检登记记录 A,合约单位 B,体检人员档案 C,病人信息 D " & _
                "WHERE A.ID=C.登记id AND A.合约单位ID=B.ID(+) AND D.病人id=C.病人id AND C.ID=[1] "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检部门清单
        
            If strParam = "所有" Then
            
                strSQL = "SELECT A.编码||'-'||A.名称,ID FROM 部门表 A,部门性质说明 B WHERE (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质='体检' ORDER BY A.编码||'-'||A.名称"
            
            Else
                strSQL = "SELECT A.编码||'-'||A.名称,ID FROM 部门表 A,部门性质说明 B WHERE (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质='体检' " & _
                            "AND A.ID IN (SELECT 部门id FROM 部门人员 WHERE 人员id=[1])  ORDER BY A.编码||'-'||A.名称"
            End If
        
'        Case SQL.疾病编码选择
'
'            strSQL = "Select * " & _
'                     "from (Select 0 As 选择,ID,上级ID,0 as 末级,序号,'' As 编码,名称 ,'' as 简码,'' AS 附码 " & _
'                             "From 疾病编码分类 " & _
'                            "Where 类别 = 'D' " & _
'                            "Start With 上级id Is Null Connect by Prior ID = 上级ID " & _
'                           "Union All " & _
'                             "Select 0 As 选择,A.ID,A.分类id AS 上级ID,1 as 末级,0 As 序号, A.编码,A.名称,A.简码,A.附码 " & _
'                               "FROM 疾病编码目录 A " & _
'                              "Where A.类别='D' " & _
'                           ") A Order by A.末级,A.序号 "
'        Case SQL.疾病编码过滤
'
'            varParam(0) = "'%" & UCase(varParam(0)) & "%'"
'            strSQL = "SELECT A.ID,A.编码,A.名称,A.简码,A.附码 " & _
'                        "FROM 疾病编码目录 A " & _
'                        "WHERE A.类别 ='D' "
'
'            strSQL = strSQL & " AND (UPPER(A.编码) Like " & varParam(0) & " OR A.名称 Like " & varParam(0) & " OR A.简码 Like " & varParam(0) & ")"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.收费项目选择
    
            strSQL = "select * " & _
                     "from (Select ID,上级ID,0 as 末级,编码,名称 ,'' as 单位,'' AS 类别,'' As 单价,'' As 规格 " & _
                             "From 收费分类目录 " & _
                            "Start With 上级ID Is Null " & _
                           "Connect by Prior ID = 上级ID "
            
            strSQL = strSQL & " Union All Select -1 As ID,Null+0 As 上级ID,0 as 末级,'-1' As 编码,'西成药' As 名称 ,'' as 单位,'' AS 类别,'' As 单价,'' As 规格 from dual "
            strSQL = strSQL & " Union All Select -2 As ID,Null+0 As 上级ID,0 as 末级,'-2' As 编码,'中成药' As 名称 ,'' as 单位,'' AS 类别,'' As 单价,'' As 规格 from dual "
            strSQL = strSQL & " Union All Select -3 As ID,Null+0 As 上级ID,0 as 末级,'-3' As 编码,'中草药' As 名称 ,'' as 单位,'' AS 类别,'' As 单价,'' As 规格 from dual "
            strSQL = strSQL & " Union All Select -7 As ID,Null+0 As 上级ID,0 as 末级,'-7' As 编码,'卫生材料' As 名称 ,'' as 单位,'' AS 类别,'' As 单价,'' As 规格 from dual "
             
            strSQL = strSQL & _
                           "Union All " & _
                             "Select A.ID,Decode(A.类别,'5',-1,'6',-2,'7',-3,'4',-7,A.分类id) AS 上级ID,1 as 末级, A.编码,A.名称,A.计算单位 AS 单位,A.类别,Trim(To_Char(c.单价,'9999999999999.00000')) As 单价,a.规格 " & _
                               "FROM 收费项目目录 A,收费项目类别 B,(select 收费细目id,sum(现价) AS 单价 from 收费价目 where 执行日期<=SYSDATE and (终止日期 IS NULL OR 终止日期>SYSDATE) group by 收费细目id) C " & _
                              "Where C.收费细目id(+)=A.ID AND Nvl(a.是否变价,0)=0 And A.类别=b.编码 AND (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) " & _
                           ") A " & _
                    "ORDER BY A.末级, A.编码"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.收费项目过滤
                        
            If CheckStrType(varParam(0), 1) And Left(ParamInfo.收费诊疗项目匹配, 1) = 1 Then
                '是全数字，按编码查找
                
                strSQL = "SELECT A.ID,A.编码,A.名称,A.计算单位 AS 单位,A.类别,Trim(To_Char(c.单价,'9999999999999.00000')) As 单价,a.规格 " & _
                            "FROM 收费项目目录 a,收费项目类别 b,(select 收费细目id,sum(现价) AS 单价 from 收费价目 where 执行日期<=SYSDATE and (终止日期 IS NULL OR 终止日期>SYSDATE) group by 收费细目id) c " & _
                            "WHERE c.收费细目id(+)=a.ID and Nvl(a.是否变价,0)=0 And a.类别=b.编码 AND (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 is NULL) "
                        
                strSQL = strSQL & " AND a.编码 Like [1]"
                
            ElseIf CheckStrType(varParam(0), 2) And Left(ParamInfo.收费诊疗项目匹配, 2) = 1 Then
                '是全字母，按简码查找

                strSQL = "SELECT Distinct A.ID,A.编码,A.名称,A.计算单位 AS 单位,A.类别,Trim(To_Char(c.单价,'9999999999999.00000')) As 单价,a.规格 " & _
                            "FROM 收费项目目录 a,收费项目类别 b,收费项目别名 d,(select 收费细目id,sum(现价) AS 单价 from 收费价目 where 执行日期<=SYSDATE and (终止日期 IS NULL OR 终止日期>SYSDATE) group by 收费细目id) c " & _
                            "WHERE c.收费细目id(+)=a.ID and Nvl(a.是否变价,0)=0 And a.类别=b.编码 AND (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 is NULL) "
                        
                strSQL = strSQL & " AND a.ID=d.收费细目ID AND [1] Is Not Null And d.简码 Like [2]"
                
            Else
                strSQL = "SELECT Distinct A.ID,A.编码,A.名称,A.计算单位 AS 单位,A.类别,Trim(To_Char(c.单价,'9999999999999.00000')) As 单价,a.规格 " & _
                            "FROM 收费项目目录 a,收费项目类别 b,收费项目别名 d,(select 收费细目id,sum(现价) AS 单价 from 收费价目 where 执行日期<=SYSDATE and (终止日期 IS NULL OR 终止日期>SYSDATE) group by 收费细目id) c " & _
                            "WHERE c.收费细目id(+)=a.ID and Nvl(a.是否变价,0)=0 And a.类别=b.编码 AND (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 is NULL) "
                        
                strSQL = strSQL & " AND A.ID=d.收费细目ID AND (a.编码 Like [1] OR a.名称 Like [2] Or d.名称 Like [2] Or d.简码 Like [2])"
            End If

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检项目选择
            
            strSQL = "select * " & _
                     "from (Select DISTINCT 0 As 选择,ID,上级ID,0 as 末级,编码,名称 ,'' as 单位,'' AS 类别,'' As 标本部位, " & _
                                           "DECODE(上级ID, Null, ID * POWER(10, 20), 上级ID * POWER(10, 20) + ID) As 排序 " & _
                             "From 诊疗分类目录 " & _
                            "Where 类型 = 5 " & _
                            "Start With ID IN (SELECT DISTINCT 分类id FROM 诊疗项目目录 WHERE (撤档时间 = To_Date('30000101', 'YYYYMMDD') Or 撤档时间 is NULL) AND 类别 IN ('C','D')) " & _
                           "Connect by Prior 上级ID = ID " & _
                           "Union All " & _
                             "Select 0 As 选择,A.ID,A.分类id AS 上级ID,1 as 末级, A.编码,A.名称,A.计算单位 AS 单位,DECODE(A.类别,'C','检验','检查') AS 类别,a.标本部位, " & _
                                    "1 AS 排序 " & _
                               "FROM 诊疗项目目录 A " & _
                              "Where A.适用性别 In (0,[1],[2]) And A.类别 IN ('C','D') AND (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) " & _
                           ") A " & _
                    "ORDER BY A.末级, A.编码"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检项目过滤选择
            
            If CheckStrType(varParam(0), 1) And Left(ParamInfo.收费诊疗项目匹配, 1) = 1 Then
                '是全数字，按编码查找
                    
                strSQL = "SELECT A.ID,A.编码,A.名称,A.计算单位 AS 单位,DECODE(A.类别,'C','检验','检查') AS 类别,a.标本部位 " & _
                        "FROM 诊疗项目目录 A " & _
                        "WHERE A.适用性别 In (0,[5],[6]) And A.类别 IN ([1],[2]) AND (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) "
                strSQL = strSQL & " AND A.编码 Like [3]"
                
            ElseIf CheckStrType(varParam(0), 2) And Left(ParamInfo.收费诊疗项目匹配, 2) = 1 Then
                '是全字母，按简码查找
                
                strSQL = "SELECT Distinct A.ID,A.编码,A.名称,A.计算单位 AS 单位,DECODE(A.类别,'C','检验','检查') AS 类别,a.标本部位 " & _
                        "FROM 诊疗项目目录 A,诊疗项目别名 B " & _
                        "WHERE A.适用性别 In (0,[5],[6]) And A.类别 IN ([1],[2]) AND (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) "
                strSQL = strSQL & " AND A.ID=B.诊疗项目id AND [3] Is Not Null And b.简码 Like [4]"
                
            Else
            
                strSQL = "SELECT Distinct A.ID,A.编码,A.名称,A.计算单位 AS 单位,DECODE(A.类别,'C','检验','检查') AS 类别,a.标本部位 " & _
                        "FROM 诊疗项目目录 A,诊疗项目别名 B " & _
                        "WHERE A.适用性别 In (0,[5],[6]) And A.类别 IN ([1],[2]) AND (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) "
                strSQL = strSQL & " AND A.ID=B.诊疗项目id AND (A.编码 Like [3] OR A.名称 Like [4] Or B.名称 Like [4] Or B.简码 Like [4])"
                
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.团体未结明细
            
            Dim strSub As String
            Dim strCond As String
            Dim blnZero As Boolean
            
            blnZero = (Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "对零费用进行结帐", 1)) = 1)
    
            strSQL = "Select Nvl(B.名称, '未知') as 科室, " & _
                            "A.时间, " & _
                            "A.NO as 单据号, " & _
                            "Nvl(E.名称, C.名称) as 项目, " & _
                            "A.收据费目 as 费目, " & _
                            "A.ID, " & _
                            "A.序号, " & _
                            "A.记录性质, " & _
                            "A.记录状态, " & _
                            "A.执行状态, " & _
                            "A.主页ID, " & _
                            "A.开单部门ID, " & _
                            "A.登记时间, " & _
                            "Nvl(A.未结金额, 0) 未结金额, " & _
                            "Nvl(A.未结金额, 0) 结帐金额, " & _
                            "Nvl(A.费用类型, C.费用类型) As 类型 " & _
                    "From ( "
                    
            strSQL = strSQL & _
                            "SELECT A.ID, " & _
                                     "A.NO, " & _
                                     "A.序号, " & _
                                     "A.记录性质, " & _
                                     "A.记录状态, " & _
                                     "A.执行状态, " & _
                                     "A.主页ID, " & _
                                     "A.开单部门ID, " & _
                                     "To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') as 时间, " & _
                                     "A.登记时间, " & _
                                     "A.收费细目ID, " & _
                                     "A.收入项目ID, " & _
                                     "A.收据费目, " & _
                                     "Nvl(A.实收金额, 0) as 未结金额, " & _
                                     "费用类型 "
            
            strCond = " And A.医嘱序号 IN (SELECT A.ID FROM 病人医嘱记录 A,体检登记记录 B WHERE A.挂号单 = B.体检号 and A.病人来源 = 4 AND B.合约单位ID=[1]) "
            
            If blnZero Then
                strSQL = strSQL & _
                                    "From 病人费用记录 A " & _
                                    "Where A.记录状态 <> 0 And A.记帐费用 = 1  And A.结帐id Is Null " & strCond
            Else
                strSub = _
                        "Select A.NO,A.序号,A.记录性质,Nvl(Sum(A.实收金额),0) as 实收金额 " & _
                        "From 病人费用记录 A " & _
                        "Where A.记录状态<>0  And A.记帐费用=1 And Nvl(A.实收金额,0)<>0 And A.结帐id Is Null " & strCond & _
                        "Group by A.NO,A.序号,A.记录性质 " & _
                        "Having Nvl(Sum(A.实收金额),0)<>0 "
                
                strSQL = strSQL & _
                            "From 病人费用记录 A," & _
                                "(" & strSub & ") B " & _
                            "Where A.NO=B.NO And A.序号=B.序号 And A.记录性质=B.记录性质 " & _
                            "And A.记录状态<>0 And A.记帐费用=1 And Nvl(A.实收金额,0)<>0 And A.结帐id Is Null "
            End If
                                                
            strSQL = strSQL & " Union All " & _
                              "SELECT 0 as ID, " & _
                                     "A.NO, " & _
                                     "A.序号, " & _
                                     "Mod(A.记录性质, 10) as 记录性质, " & _
                                     "A.记录状态, " & _
                                     "A.执行状态, " & _
                                     "A.主页ID, " & _
                                     "A.开单部门ID, " & _
                                     "To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') as 时间, " & _
                                     "A.登记时间, " & _
                                     "A.收费细目ID, " & _
                                     "A.收入项目ID, " & _
                                     "A.收据费目, " & _
                                     "Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) as 未结金额, " & _
                                     "A.费用类型 " & _
                                "FROM 病人费用记录 A " & _
                                "Where A.结帐id Is Not Null And A.记录状态 <> 0 And A.记帐费用 = 1 And Nvl(A.实收金额, 0) <> Nvl(A.结帐金额, 0) " & strCond & "  " & _
                                "Having Sum (Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) <> 0 " & _
                                "Group by A.NO, A.序号, Mod(A.记录性质, 10), A.记录状态, A.执行状态 , A.主页ID,A.开单部门ID,To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS'),A.登记时间,A.收费细目ID,A.收入项目ID,A.收据费目,A.费用类型"
            
            strSQL = strSQL & ") A," & _
                          "部门表 B," & _
                          "收费项目目录 C," & _
                          "收入项目 D," & _
                          "收费项目别名 E " & _
                    "Where A.开单部门ID = B.ID(+) And A.收费细目ID = C.ID And A.收入项目ID = D.ID And A.收费细目ID = E.收费细目ID(+) And E.码类(+) = 1 And E.性质(+) = 1 " & _
                    "Order by A.时间 Desc, A.NO Desc, A.记录性质, A.序号"
            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检项目清单
        
            strSQL = "SELECT A.ID," & _
                          "DECODE(A.类别, 'C', '检验', 'D', '检查') AS 类别," & _
                          "A.名称," & _
                          "B.基本价格,"
                          
            strSQL = strSQL & _
                          "D.名称 as 执行科室," & _
                          "C.名称 as 采集方式, " & _
                          "B.体检类型, " & _
                          "B.采集方式id, " & _
                          "DECODE(B.结算途径,1,'记帐','收费') AS 结算, " & _
                          "B.执行科室id, " & _
                          "B.检查部位, " & _
                          "B.检查部位id, " & _
                          "B.检验标本, " & _
                          "B.体检价格, " & _
                          "B.组别名称 AS 组别 " & _
                     "FROM 诊疗项目目录 A, 体检项目清单 B,诊疗项目目录 C,部门表 D " & _
                    "WHERE B.病人id IS NULL AND B.执行科室id=D.ID(+) AND B.采集方式id=C.ID(+) and  A.ID = B.诊疗项目ID AND B.登记id=[1] and B.组别名称=[2]"
        
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检人员档案
                    
            strSQL = "Select A.健康号,A.IC卡号,TO_CHAR(B.体检时间,'yyyy-mm-dd') aS 体检时间,B.次数,A.病人id AS ID,A.病人id,A.姓名,A.门诊号,A.身份证号 AS 身份证,A.性别,A.年龄,TO_CHAR(A.出生日期,'yyyy-mm-dd') AS 出生日期,A.婚姻状况,C.组别名称 AS 组别," & _
                    "A.民族,A.国籍,A.学历,A.职业,A.身份,A.联系人姓名,A.联系人电话,B.电子邮件,A.联系人地址,A.工作单位,a.就诊卡号,B.登记时间 " & _
                    "from 病人信息 A,体检人员档案 B,(SELECT * FROM 体检组别 WHERE 登记id=[1]) C  " & _
                    "WHERE A.病人ID=B.病人ID AND B.组别名称=C.组别名称(+) AND B.登记id=[1] Order By C.组别名称,A.门诊号 "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检人员档案_单个
                    
            strSQL = "select A.健康号,A.IC卡号,A.病人id AS ID,A.病人id,A.姓名,A.门诊号,A.身份证号 AS 身份证,A.性别,A.年龄,TO_CHAR(A.出生日期,'yyyy-mm-dd') AS 出生日期,A.婚姻状况,C.组别名称 AS 组别," & _
                    "A.民族,A.国籍,A.学历,A.职业,A.身份,A.联系人姓名,A.联系人电话,B.电子邮件,A.联系人地址,A.工作单位,a.就诊卡号,B.登记时间 " & _
                    "from 病人信息 A,体检人员档案 B,(SELECT * FROM 体检组别 WHERE 登记id=[1]) C  " & _
                    "WHERE A.病人ID=B.病人ID AND B.组别名称=C.组别名称(+) AND B.登记id=[1] AND B.病人id=[2]"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检类型分类
        
            strSQL = "SELECT 0 AS ID,NULL+0 AS 上级id,'所有分类' AS 名称,1 AS 图标,1 AS 打开图标 FROM dual union all " & _
                    "SELECT 序号 AS ID,DECODE(上级序号,NULL,0,上级序号) AS 上级ID,'['||编码||']'||名称 AS 名称,1 AS 图标,1 AS 打开图标 " & _
                        "FROM 体检类型 " & _
                        "WHERE NVL(末级,0)=0 " & _
                        "START WITH 上级序号 IS NULL " & _
                        "CONNECT BY PRIOR 序号=上级序号 "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检类型分类选择
                        
            strSQL = "SELECT 0 As 选择,-1 AS ID,NULL+0 AS 上级id,'所有分类' AS 名称,'' AS 编码,'' AS 简码,'' AS 说明,1 AS 图标,1 AS 打开图标,0 AS 末级 FROM dual union all " & _
                    "SELECT 0 As 选择,序号 AS ID,DECODE(上级序号,NULL,-1,上级序号) AS 上级ID,名称,编码,简码,说明,1 AS 图标,1 AS 打开图标,末级 " & _
                        "FROM 体检类型 " & _
                        "WHERE 适用范围 IN (0,[1]) " & _
                        "START WITH 上级序号 IS NULL " & _
                        "CONNECT BY PRIOR 序号=上级序号 "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检诊断分类
        
            strSQL = "SELECT 0 AS ID,NULL+0 AS 上级id,'所有分类' AS 名称,'class' AS 图标,'class' AS 打开图标,1 AS 排序,'0' as 编码 FROM dual union all " & _
                    "SELECT 序号 AS ID,DECODE(上级序号,NULL,0,上级序号) AS 上级ID,'['||编码||']'||名称 AS 名称,'class' AS 图标,'class' AS 打开图标,2 AS 排序,编码 " & _
                        "FROM 体检诊断建议 " & _
                        "WHERE NVL(末级,0)=0 " & _
                        "START WITH 上级序号 IS NULL " & _
                        "CONNECT BY PRIOR 序号=上级序号  order by 排序,编码 "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检类型选择
        
            strSQL = "SELECT -1 AS ID,NULL+0 AS 上级id,'所有分类' AS 名称,'' AS 编码,'' AS 简码,'' AS 说明,1 AS 图标,1 AS 打开图标,0 AS 末级 FROM dual"
            
            strSQL = strSQL & " UNION ALL " & _
                    "SELECT 序号 AS ID,DECODE(上级序号,NULL,-1,上级序号) AS 上级ID,'['||编码||']'||名称 AS 名称,编码,简码,说明,1 AS 图标,1 AS 打开图标,末级 " & _
                        "FROM 体检类型 " & _
                        "WHERE NVL(末级,0)=0 " & _
                        "START WITH 上级序号 IS NULL " & _
                        "CONNECT BY PRIOR 序号=上级序号 "
                        
            strSQL = strSQL & " UNION ALL " & _
                    "SELECT 序号 AS ID,DECODE(上级序号,NULL,-1,上级序号) AS 上级ID,名称,编码,简码,说明,1 AS 图标,1 AS 打开图标,末级 " & _
                        "FROM 体检类型 " & _
                        "WHERE NVL(末级,0)=1 "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检类型过滤选择
            
            
            strSQL = "SELECT 序号 AS ID,名称,编码,简码,说明 " & _
                        "FROM 体检类型 " & _
                        "WHERE 末级=1 AND (编码 LIKE [1] OR 名称 LIKE [2] OR 简码 LIKE [2])"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.团体过滤选择
            
            '调用:  1.
            '       2.
            
            strSQL = "select ID,编码,名称,简码,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位 " & _
                " Where (编码 Like [1] or 名称 Like [1] OR 简码 Like [1])"

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检团体选择
            
            '调用:  1.frmSchedualEdit\cmd_Click
            '       2.
            
            strSQL = "SELECT -1 AS ID,NULL+0 AS 上级id,'0' AS 编码,'所有' AS 名称,'' as 简码,0 AS 末级,'' AS 联系人,'' AS 电话,'' AS 电子邮件,'' AS 开户银行,'' AS 帐号,'' AS 地址,'' AS 说明 from dual " & _
                        "Union All " & _
                        "SELECT ID,DECODE(上级id,NULL,-1,0,-1,上级id) AS 上级id,编码,名称,简码,0 AS 末级,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位   " & _
                        "Start With 上级id is null connect by prior ID=上级id " & _
                        "Union All " & _
                        "SELECT ID,DECODE(上级id,NULL,-1,0,-1,上级id) AS 上级id,编码,名称,简码,1 AS 末级,联系人,电话,电子邮件,开户银行,帐号,地址,说明 from 合约单位 "
                                                
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.诊治项目选择
        
            '调用:  1.
            '       2.
            
            strSQL = "SELECT * FROM (" & _
                        "(select -1 AS ID,0 AS 上级id,'0' AS 编码,'所有项目' AS 名称,'' AS 临床意义,'' AS 数值域,0 AS 末级,0 AS 排序,0 AS 类型,'' AS 单位 from dual UNION ALL " & _
                        "Select DISTINCT ID," & _
                                        "DECODE(上级ID,NULL,-1,上级ID) AS 上级ID," & _
                                        "编码," & _
                                        "名称," & _
                                        "'' as 临床意义," & _
                                        "'' as 数值域," & _
                                        "0 as 末级," & _
                                        "DECODE(上级ID,Null,ID * POWER(10, 20),上级ID * POWER(10, 20) + ID) As 排序,0 AS 类型,'' AS 单位 " & _
                                  "From 诊治所见分类 " & _
                                 "Start With ID IN " & _
                                               "( " & _
                                               "SELECT 分类id from 诊治所见项目 A " & _
                                               "where A.ID IN (SELECT A.所见项id " & _
                                                              "FROM 病历所见单 A, 病历元素目录 B " & _
                                                              "WHERE A.元素id = B.ID AND B.类型 = 2 AND B.适用 LIKE '%1') " & _
                                               "Union " & _
                                               "SELECT 分类id from 诊治所见项目 A " & _
                                               "where A.ID IN (SELECT DISTINCT 报告项目id from 检验报告项目 A) " & _
                                               ") " & _
                                "Connect by Prior 上级ID = ID) "

            strSQL = strSQL & _
                        "Union All " & _
                        "(SELECT ID, 分类id AS 上级id, 编码, 中文名 AS 名称, 临床意义, 数值域,1 AS 末级,1 AS 排序,类型,单位 " & _
                          "from 诊治所见项目 A " & _
                         "where A.ID IN " & _
                               "(SELECT A.所见项id FROM 病历所见单 A, 病历元素目录 B WHERE A.元素id = B.ID AND B.类型 = 2 AND B.适用 LIKE '%1') " & _
                        "Union " & _
                        "SELECT ID, DECODE(分类id,NULL,-1,分类id) AS 上级id, 编码, 中文名 AS 名称, 临床意义, 数值域,1 AS 末级,1 AS 排序,类型,单位 " & _
                          "from 诊治所见项目 A " & _
                         "where A.ID IN " & _
                               "(SELECT DISTINCT 报告项目id from 检验报告项目 A)) " & _
                        ") A ORDER BY A.末级,A.编码"

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.诊治项目过滤选择
            
            '调用者:1.
            '       2.
            
'            varParam(0) = "'%" & UCase(varParam(0)) & "%'"
            
            strSQL = "SELECT * FROM (" & _
                        "SELECT ID, 分类id AS 上级id, 编码, 中文名 AS 名称, 临床意义, 数值域,英文名,类型,单位 " & _
                          "from 诊治所见项目 A " & _
                         "where A.ID IN " & _
                               "(SELECT A.所见项id FROM 病历所见单 A, 病历元素目录 B WHERE A.元素id = B.ID AND B.类型 = 2 AND B.适用 LIKE '%1') " & _
                        "Union " & _
                        "SELECT ID, DECODE(分类id,NULL,-1,分类id) AS 上级id, 编码, 中文名 AS 名称, 临床意义, 数值域,英文名,类型,单位 " & _
                          "from 诊治所见项目 A " & _
                         "where A.ID IN " & _
                               "(SELECT DISTINCT 报告项目id from 检验报告项目 A) " & _
                        ") A WHERE A.编码 LIKE [1] OR A.名称 LIKE [2] OR A.英文名 LIKE [2] Or zlSpellCode(A.名称) Like [2]  ORDER BY A.编码"


        '--------------------------------------------------------------------------------------------------------------
        Case SQL.检查诊治项目选择
            
            '调用者:1.
            '       2.
            
            strSQL = "SELECT * FROM (" & _
                        "(select -1 AS ID,0 AS 上级id,'0' AS 编码,'所有项目' AS 名称,'' AS 临床意义,'' AS 数值域,0 AS 末级,0 AS 排序,0 AS 类型,'0' As 性质 from dual UNION ALL " & _
                        "Select DISTINCT ID," & _
                                        "DECODE(上级ID,NULL,-1,上级ID) AS 上级ID," & _
                                        "编码," & _
                                        "名称," & _
                                        "'' as 临床意义," & _
                                        "'' as 数值域," & _
                                        "0 as 末级," & _
                                        "DECODE(上级ID,Null,ID * POWER(10, 20),上级ID * POWER(10, 20) + ID) As 排序,0 AS 类型,性质 " & _
                                  "From 诊治所见分类 " & _
                                "Connect by Prior 上级ID = ID) "

            strSQL = strSQL & _
                        "Union All " & _
                        "(SELECT ID, 分类id AS 上级id, 编码, 中文名 AS 名称, 临床意义, 数值域,1 AS 末级,1 AS 排序,类型,'z' As 性质 " & _
                          "from 诊治所见项目 A ) " & _
                        ") A ORDER BY A.性质,A.末级,A.编码"

        '--------------------------------------------------------------------------------------------------------------
        Case SQL.检查诊治项目过滤选择
        
            '调用者:1.
            '       2.
            
            strSQL = "SELECT * FROM (" & _
                        "SELECT ID, 分类id AS 上级id, 编码, 中文名 AS 名称, 临床意义, 数值域,英文名,类型 " & _
                          "from 诊治所见项目 A " & _
                         "where 分类id IS NOT NULL " & _
                        ") A WHERE A.编码 LIKE [1] OR A.名称 LIKE [2] OR A.英文名 LIKE [2] ORDER BY A.编码"
                                 
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.人员体检项目
            
            '调用者:1.frmMedicalStation\EditData

            strSQL = "SELECT A.ID,B.ID AS 清单id,F.复查清单id," & _
                          "DECODE(A.类别, 'C', '检验', 'D', '检查') AS 类别," & _
                          "A.名称," & _
                          "Decode(B.病人id,NULL,'','附加') AS 附加," & _
                          "D.名称 as 执行科室," & _
                          "G.名称 as 采集科室," & _
                          "B.基本价格,"
                          
            strSQL = strSQL & _
                          "C.名称 as 采集方式, " & _
                          "E.组别名称, " & _
                          "B.采集方式id, " & _
                          "B.采集科室id, " & _
                          "B.执行科室id, " & _
                          "B.检查部位, " & _
                          "B.体检类型, " & _
                          "B.体检价格,Decode(b.基本价格,0,0,Null,0,10*B.体检价格/B.基本价格) As 折扣," & _
                          "DECODE(B.结算途径,1,'记帐','收费') AS 结算方式, " & _
                          "B.检查部位id, " & _
                          "Decode(B.病人id,NULL,'1','0') AS 公共, " & _
                          "B.检验标本 " & _
                     "FROM 诊疗项目目录 A, 体检项目清单 B,诊疗项目目录 C,部门表 D,体检人员档案 E,体检项目医嘱 F,部门表 G " & _
                    "WHERE B.执行科室id=D.ID(+) AND B.采集方式id=C.ID(+) AND B.采集科室id=G.ID(+) And  A.ID = B.诊疗项目ID AND E.登记id=[1] AND E.登记id=B.登记id AND E.病人id=F.病人id AND F.清单id=B.ID AND F.病人id=[2] "
            
            strSQL = strSQL & " Order By 公共 Desc,A.名称"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.人员原始项目
            strSQL = "SELECT A.ID,B.ID AS 清单id," & _
                          "DECODE(A.类别, 'C', '检验', 'D', '检查') AS 类别," & _
                          "A.名称," & _
                          "Decode(B.病人id,NULL,'','附加') AS 附加," & _
                          "D.名称 as 执行科室," & _
                          "G.名称 as 采集科室," & _
                          "B.基本价格,"
                          
            strSQL = strSQL & _
                          "C.名称 as 采集方式, " & _
                          "E.组别名称, " & _
                          "B.采集方式id, " & _
                          "B.采集科室id, " & _
                          "B.执行科室id, " & _
                          "B.检查部位, " & _
                          "B.体检类型, " & _
                          "B.体检价格,Decode(b.基本价格,0,0,Null,0,10*B.体检价格/B.基本价格) As 折扣," & _
                          "DECODE(B.结算途径,1,'记帐','收费') AS 结算方式, " & _
                          "B.检查部位id, " & _
                          "Decode(B.病人id,NULL,'1','0') AS 公共, " & _
                          "B.检验标本,0 As 选择 " & _
                     "FROM 诊疗项目目录 A, 体检项目清单 B,诊疗项目目录 C,部门表 D,体检人员档案 E,体检项目医嘱 F,部门表 G " & _
                    "WHERE B.执行科室id=D.ID(+) AND B.采集方式id=C.ID(+) AND B.采集科室id=G.ID(+) And  A.ID = B.诊疗项目ID AND E.登记id=[1] AND E.登记id=B.登记id AND E.病人id=F.病人id AND F.清单id=B.ID AND F.病人id=[2] and F.复查清单id Is Null "
            
            strSQL = strSQL & " Order By 公共 Desc,A.名称"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.病人所有项目
        
            '传入体检人员档案id

            strSQL = "SELECT A.ID," & _
                          "A.名称 AS 项目,Decode(f.复查清单id,0,0,Null,0,255) As 前景色," & _
                          "Decode(B.病人id,NULL,'','附加') AS 公共," & _
                          "D.名称 as 执行科室," & _
                          "(SELECT DECODE(病历文件id,NULL,'','单据') FROM 诊疗单据应用 WHERE 应用场合=4 AND 诊疗项目id=A.ID) AS 状态 " & _
                     "FROM 诊疗项目目录 A, 体检项目清单 B,部门表 D,体检人员档案 E,体检项目医嘱 F " & _
                    "WHERE B.执行科室id=D.ID(+)  And  A.ID = B.诊疗项目ID AND E.ID=[1] AND E.登记id=B.登记id AND E.病人id=F.病人id AND F.清单id=B.ID"
            
            strSQL = strSQL & " Order By D.名称"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.团体体检项目
        
            '调用者:1.frmMedicalStation\EditData
            '       2.frmSchedual\
                            
            strSQL = _
               "SELECT A.ID,B.ID As 清单id,0 As 复查清单id, " & _
                  "DECODE(A.类别, 'C', '检验', 'D', '检查') AS 类别," & _
                  "A.名称," & _
                  "D.名称 as 执行科室," & _
                  "E.名称 as 采集科室," & _
                  "C.名称 as 采集方式," & _
                  "B.检验标本," & _
                  "B.检查部位," & _
                  "B.采集方式id," & _
                  "B.采集科室id," & _
                  "B.检查部位id," & _
                  "B.执行科室id," & _
                  "B.体检类型," & _
                  "B.组别名称," & _
                  "B.体检价格,'1' As 公共," & _
                  "DECODE(B.结算途径,1,'记帐','收费') AS 结算方式," & _
                  "B.基本价格,Decode(B.基本价格,0,0,Null,0,10*B.体检价格/B.基本价格) As 折扣 "

            strSQL = strSQL & _
                "FROM 诊疗项目目录 A, " & _
                      "体检项目清单 B, " & _
                      "诊疗项目目录 C, " & _
                      "部门表 D, " & _
                      "部门表 E " & _
                "Where B.组别名称 Is Not Null " & _
                      "AND B.执行科室id=D.ID(+) " & _
                      "AND B.采集科室id=E.ID(+) " & _
                      "AND B.采集方式id=C.ID(+) " & _
                      "AND A.ID = B.诊疗项目ID " & _
                      "AND B.登记id=[1] " & _
                "ORDER BY B.组别名称 "
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检类型项目
            
            strSQL = "SELECT x.ID, " & _
                          "DECODE(x.类别, 'C', '检验', 'D', '检查') AS 项目类别, " & _
                          "x.名称 As 项目名称, " & _
                          "x.计算单位, " & _
                          "y.检验标本, " & _
                          "y.检查部位, " & _
                          "t.名称 As 采集方式, " & _
                          "z.基本价格,z.体检价格, " & _
                          "Decode(z.基本价格,Null,0,0,0,10*z.体检价格/z.基本价格) As 折扣," & _
                          "y.采集方式id, " & _
                          "y.检查部位id,'' As 计费明细 " & _
                     "FROM 诊疗项目目录 x, " & _
                          "体检类型目录 y, " & _
                          "(Select a.诊疗项目id,Sum(b.现价*a.数次) As 基本价格,Sum(b.现价*a.数次*Nvl(a.折扣,1)) As 体检价格 " & _
                           "From 体检类型计价 a, " & _
                                "收费价目 b " & _
                           "Where a.序号 = [1] " & _
                                 "and b.收费细目id=a.收费细目id " & _
                                 "and b.执行日期<=SYSDATE and (b.终止日期 IS NULL OR b.终止日期>SYSDATE) " & _
                           "Group by a.诊疗项目id " & _
                          ") z, " & _
                          "诊疗项目目录 t " & _
                    "Where x.ID = y.诊疗项目ID And y.序号 = [1] and x.id=z.诊疗项目id(+) " & _
                          "and t.id(+)=y.采集方式id"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检类型计价
            
            strSQL = "Select c.名称,c.计算单位,b.现价,a.数次,b.收入项目id,b.现价*a.数次 As 金额,c.id " & _
                        "from 体检类型计价 a,收费价目 b,收费项目目录 c " & _
                        "Where a.收费细目id = c.ID " & _
                              "and b.收费细目id=a.收费细目id " & _
                              "and b.执行日期<=SYSDATE and (b.终止日期 IS NULL OR b.终止日期>SYSDATE) and a.序号=[1]"
       
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.个人费用概况

            strSQL = "SELECT D.应收金额,D.实收金额,D.结帐金额,D.记帐费用 " & _
                     "FROM 病人费用记录 D, " & _
                          "(SELECT C.ID " & _
                             "FROM 体检人员档案 A, 体检登记记录 B, 病人医嘱记录 C " & _
                            "WHERE A.登记ID = B.ID AND A.病人ID = C.病人ID AND C.病人来源 = 4 AND " & _
                                  "B.体检号 = C.挂号单 AND A.ID = [1]) E " & _
                    "WHERE D.医嘱序号 = E.ID"
            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.团体费用概况

            strSQL = "SELECT D.应收金额,D.实收金额,D.结帐金额,D.记帐费用 " & _
                     "FROM 病人费用记录 D, " & _
                          "(SELECT C.ID " & _
                             "FROM 体检登记记录 B, 病人医嘱记录 C " & _
                            "WHERE C.病人来源 = 4 AND " & _
                                  "B.体检号 = C.挂号单 AND B.ID = [1]) E " & _
                    "WHERE D.医嘱序号 = E.ID"
                    
'            strSQL = "SELECT NVL(SUM(D.实收金额), 0) AS 实收金额," & _
'                        "NVL(SUM(DECODE(D.记帐费用,1,D.实收金额,0)),0) AS 记帐金额," & _
'                        "NVL(SUM(DECODE(D.记帐费用,1,0,D.实收金额)),0) AS 收费金额, " & _
'                        "NVL(SUM(DECODE(D.记帐费用,1,NVL(D.实收金额,0) - NVL(D.结帐金额,0),0)),0) AS 未结金额, " & _
'                        "NVL(SUM(DECODE(D.记帐费用,1,0,NVL(D.实收金额,0) -  NVL(D.结帐金额,0))),0) AS 未收金额, " & _
'                        "NVL(SUM(NVL(D.实收金额,0) - NVL(D.结帐金额,0)), 0) AS 未结算合计 " & _
'                     "FROM 病人费用记录 D, " & _
'                          "(SELECT C.ID " & _
'                             "FROM 体检人员档案 A, 体检登记记录 B, 病人医嘱记录 C " & _
'                            "WHERE A.登记ID = B.ID AND A.病人ID = C.病人ID AND C.病人来源 = 4 AND " & _
'                                  "C.医嘱状态 <> 4 AND B.体检号 = C.挂号单 AND A.ID = [1]) E " & _
'                    "WHERE D.记录状态 IN (0, 1) AND D.医嘱序号 = E.ID"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检项目价表
            
            If varParam(2) = "" Then
                
                strSQL = "Select y.名称,y.计算单位,z.收费数量,x.现价,y.id,Decode(x.诊疗项目id,[2],2,1) As 计价性质,y.类别 " & _
                            "From ( " & _
                              "Select a.诊疗项目id,a.收费项目id,Sum(c.现价) As 现价 " & _
                              "From 收费价目 c, " & _
                                   "诊疗收费关系 a, " & _
                                   "诊疗项目目录 b " & _
                              "Where a.收费项目id = c.收费细目id " & _
                                    "and c.执行日期<=SYSDATE and (c.终止日期 IS NULL OR c.终止日期>SYSDATE) " & _
                                    "AND b.ID=a.诊疗项目id " & _
                                    "AND NVL(b.计价性质,0)=0 " & _
                                    "and a.诊疗项目id In ([1],[2]) " & _
                              "Group by a.诊疗项目id,a.收费项目id " & _
                            ") x, " & _
                            "收费项目目录 y, " & _
                            "诊疗收费关系 z " & _
                            "Where x.收费项目id = y.ID " & _
                                  "and z.收费项目id=x.收费项目id " & _
                                  "and z.诊疗项目id=x.诊疗项目id"
                                  
            Else
            
                strTmp = Val(varParam(0)) & "," & Val(varParam(1)) & "," & varParam(2)
                If Right(strTmp, 1) = "," Then strTmp = strTmp & "0"
            
                strSQL = "Select y.名称,y.计算单位,z.收费数量,x.现价,y.id,Decode(x.诊疗项目id," & Val(varParam(1)) & ",2,1) As 计价性质,y.类别 " & _
                            "From ( " & _
                              "Select a.诊疗项目id,a.收费项目id,Sum(c.现价) As 现价 " & _
                              "From 收费价目 c, " & _
                                   "诊疗收费关系 a, " & _
                                   "诊疗项目目录 b " & _
                              "Where a.收费项目id = c.收费细目id " & _
                                    "and c.执行日期<=SYSDATE and (c.终止日期 IS NULL OR c.终止日期>SYSDATE) " & _
                                    "AND b.ID=a.诊疗项目id " & _
                                    "AND NVL(b.计价性质,0)=0 " & _
                                    "and a.诊疗项目id In (" & strTmp & ") " & _
                              "Group by a.诊疗项目id,a.收费项目id " & _
                            ") x, " & _
                            "收费项目目录 y, " & _
                            "诊疗收费关系 z " & _
                            "Where x.收费项目id = y.ID " & _
                                  "and z.收费项目id=x.收费项目id " & _
                                  "and z.诊疗项目id=x.诊疗项目id"
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检预约单据

            
            strSQL = "SELECT A.ID," & _
                             "A.体检号," & _
                             "A.联系人 AS 预约人," & _
                             "A.联系电话," & _
                             "DECODE(A.是否团体,1,'','个人') AS 性质," & _
                             "DECODE(A.体检状态,1,'新开',2,'确认',3,'取消',4,'开始',5,'完成') AS 状态," & _
                             "(SELECT COUNT(1) FROM 体检人员档案 WHERE 病人id>0 AND 登记id=A.ID) AS 人数," & _
                             "A.体检状态,A.合约单位id," & _
                             "A.联系地址, "
            
            strSQL = strSQL & _
                            "(SELECT NVL(SUM(基本价格),0) FROM 体检项目清单 X,体检人员档案 T WHERE X.登记ID = T.登记ID AND X.组别名称=T.组别名称 AND T.病人id>0 AND T.登记id=A.ID) AS 应收金额,"
                            
            strSQL = strSQL & _
                            "(SELECT NVL(SUM(体检价格),0) FROM 体检项目清单 X,体检人员档案 T WHERE X.登记ID = T.登记ID AND X.组别名称=T.组别名称 AND T.病人id>0 AND T.登记id=A.ID) AS 体检价格,"
            
            strSQL = strSQL & _
                             "A.结算折扣," & _
                             "A.体检类型," & _
                             "TO_CHAR(A.体检时间,'yyyy-MM-dd') AS 预约时间," & _
                             "to_char(A.登记时间,'yyyy-MM-dd HH:mm') AS 登记时间," & _
                             "B.名称 AS 团体," & _
                             "A.附加说明 " & _
                        "FROM 体检登记记录 A,合约单位 B " & _
                        "WHERE A.合约单位ID=B.ID(+) AND A.体检部门id=[1] " & strParam
            
            strSQL = "SELECT ID,体检号 AS No,预约人,团体,附加说明,联系电话,性质,状态,人数,体检状态,合约单位id,联系地址,DECODE(结算折扣,1,NULL,10*结算折扣) AS 折扣,应收金额,体检价格 AS 实收金额,体检类型,预约时间,登记时间 FROM (" & strSQL & ") ORDER BY 体检号 DESC"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检人数统计
            
            If Val(varParam(0)) = 2 Then
                strSQL = "Select Count(1) From 体检人员档案 A,体检登记记录 B Where A.登记id=B.ID AND A.病人id>0 And A.体检报到=[4] AND B.体检状态=[1] AND B.体检时间 BETWEEN [2] AND [3]"
            Else
                strSQL = "Select Count(1) From 体检人员档案 A,体检登记记录 B Where A.登记id=B.ID And a.病人id>0 And a.体检报到=[4] AND a.体检状态=[1] AND b.体检时间 BETWEEN [2] AND [3]"
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检组别人员1
        
            strSQL = "SELECT B.ID," & _
                            "1 AS 排序," & _
                            "B.组别名称 AS 排序2," & _
                            "C.合约单位id AS 上级id," & _
                            "A.病人id,A.门诊号,a.健康号,a.就诊卡号,b.体检编号, " & _
                            "C.体检号 AS 体检单号," & _
                            "B.次数," & _
                            "A.姓名," & _
                            "A.性别," & _
                            "A.年龄," & _
                            "A.婚姻状况,B.登记id," & _
                            "Decode(C.合约单位id,NULL,98,99) AS 标志," & _
                            "C.体检号 AS 单据号," & _
                            "'单据' AS 报告," & _
                            "DECODE(B.体检状态,1,'确认',4,'开始',5,'完成') AS 状态 " & _
                        "FROM 病人信息 A,体检人员档案 B,体检登记记录 C " & _
                        "WHERE B.体检报到=[4]  AND C.体检状态=[3] AND A.病人ID=B.病人ID AND C.ID=B.登记id " & _
                            "AND B.组别名称=[2]  "
            If Val(varParam(0)) > 0 Then
                strSQL = strSQL & " AND C.ID=[1] "
            Else
                strSQL = strSQL & "AND Nvl(C.是否团体,0)=0 AND C.体检时间 BETWEEN [5] AND [6]"
            End If
            
            If blnMoveOuted Then
                strTmp = strSQL
                strTmp = Replace(strTmp, "体检人员档案", "H体检人员档案")
                strTmp = Replace(strTmp, "体检登记记录", "H体检登记记录")
                strSQL = "Select * From (" & strSQL & " Union All " & strTmp & ") a "
            End If
            
            strSQL = strSQL & " Order By a.门诊号"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检组别人员
            
            strSQL = _
                "Select A.ID, 1 As 排序, A.组别名称 As 排序2, A.合约单位id As 上级id, A.体检号 As 体检单号, a.体检编号,A.次数, A.病人id, B.门诊号," & vbNewLine & _
                "       B.健康号, B.就诊卡号, B.姓名, B.性别, B.年龄, B.婚姻状况, A.登记id, Decode(A.合约单位id, Null, 98, 99) As 标志," & vbNewLine & _
                "       A.体检号 As 单据号, Decode(A.已填, A.总数, '报告', '单据') As 报告," & vbNewLine & _
                "       Decode(A.体检状态, 1, '确认', 4, '开始', 5, '完成') As 状态" & vbNewLine & _
                "From (Select A.登记id, A.病人id, A.体检号,a.体检编号, A.ID, A.组别名称, A.合约单位id, A.次数, A.体检状态," & vbNewLine & _
                "              Sum(Decode(B.病历文件id, Null, 0, 1)) As 总数, Sum(Decode(A.报告id, Null, 0, 1)) As 已填" & vbNewLine & _
                "       From (Select A.登记id, A.病人id, A.体检号,a.体检编号, A.ID, A.组别名称, A.合约单位id, A.次数, A.体检状态, D.诊疗类别," & vbNewLine & _
                "                     D.诊疗项目id, E.报告id, D.相关id" & vbNewLine & _
                "              From (Select A.登记id, A.病人id, B.体检号,a.体检编号, A.ID, A.组别名称, B.合约单位id, A.次数, A.体检状态" & vbNewLine & _
                "                     From 体检人员档案 A, 体检登记记录 B" & vbNewLine & _
                "                     Where A.登记id = B.ID And A.体检报到 = [4] And B.体检状态 = [3] And A.组别名称 = [2] "
        
            If Val(varParam(1)) = 1 Then
                '按报到人员查
                                        
                If Val(varParam(0)) > 0 Then
                    strSQL = strSQL & _
                            "                          And b.ID=[1] And a.体检时间 Between [5] And [6] "
                Else
                
                    strSQL = strSQL & _
                        "                          And Nvl(B.是否团体, 0) = 0 And a.体检时间 Between [5] And [6] "
                        
                End If
                
            Else
                If Val(varParam(0)) > 0 Then
    
                    strSQL = strSQL & _
                        "                          And b.ID=[1] "
                Else
                
                    strSQL = strSQL & _
                        "                          And Nvl(B.是否团体, 0) = 0 And B.体检时间 Between [5] And [6] "
                End If
            End If
            
        
                
            strSQL = strSQL & _
                ") A, 病人医嘱记录 D," & vbNewLine & _
                "                   病人医嘱发送 E" & vbNewLine & _
                "              Where D.挂号单(+) = A.体检号 And D.病人id(+) = A.病人id And D.诊疗类别(+) <> 'E' And D.病人来源(+) = 4 And" & vbNewLine & _
                "                    D.医嘱状态(+) <> 4 And D.ID = E.医嘱id(+)) A, 诊疗单据应用 B" & vbNewLine & _
                "       Where ((A.诊疗类别 = 'D' And A.相关id Is Null) Or A.诊疗类别 = 'C' Or A.诊疗类别 Is Null) And" & vbNewLine & _
                "             A.诊疗项目id = B.诊疗项目id(+) And B.应用场合(+) = 4" & vbNewLine & _
                "       Group By A.登记id, A.病人id, A.体检号,a.体检编号, A.ID, A.组别名称, A.合约单位id, A.次数, A.体检状态) A, 病人信息 B" & vbNewLine & _
                "Where A.病人id = B.病人id "

            If blnMoveOuted Then
                strTmp = strSQL
                strTmp = Replace(strTmp, "体检人员档案", "H体检人员档案")
                strTmp = Replace(strTmp, "体检登记记录", "H体检登记记录")
                strTmp = Replace(strTmp, "病人医嘱记录", "H病人医嘱记录")
                strTmp = Replace(strTmp, "病人医嘱发送", "H病人医嘱发送")
                strSQL = "Select * From (" & strSQL & " Union All " & strTmp & ") b "
            End If
            strSQL = strSQL & " Order By b.门诊号"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.体检登记单据
            '0-个人分组项;1-团体名称项;2-团体组别项;98-非团体受检人员项;99-团体受检人员
            '个人体检单头
            strSQL = "SELECT -1 AS ID," & _
                            "-1 AS 排序1," & _
                            "'' AS 排序2," & _
                            "0 AS 上级id," & _
                            "'' AS 状态," & _
                            "'' AS 报告," & _
                            "NULL+0 AS 病人id," & _
                            "'' AS 体检单号," & _
                            "0 AS 门诊号," & _
                            "'' AS 健康号,'' As 就诊卡号,'' As 体检编号," & _
                            "'<个人>' AS 姓名," & _
                            "'' AS 性别," & _
                            "'' AS 年龄," & _
                            "'' AS 婚姻状况," & _
                            "NULL+0 AS 登记id," & _
                            "0 AS 标志,Null+0 As 次数," & _
                            "'' AS 单据号 " & _
                        "FROM DUAL "

            '团体资料
            strSQL = strSQL & " UNION ALL " & _
                        "SELECT DISTINCT A.ID," & _
                                    "1 AS 排序1," & _
                                    "'' AS 排序2," & _
                                    "0 AS 上级id," & _
                                    "'' AS 状态," & _
                                    "'' AS 报告," & _
                                    "0 AS 病人id," & _
                                    "B.体检号 AS 体检单号," & _
                                    "NULL+0 AS 门诊号," & _
                                    "'' AS 健康号,'' As 就诊卡号,'' As 体检编号," & _
                                    "A.名称||'('||B.体检号||')' AS 姓名," & _
                                    "'' AS 性别," & _
                                    "'' AS 年龄," & _
                                    "'' AS 婚姻状况," & _
                                    "B.ID AS 登记id, " & _
                                    "1 AS 标志,Null+0 As 次数," & _
                                    "B.体检号 AS 单据号 " & _
                    "FROM   合约单位 A," & _
                            "体检登记记录 B " & _
                    "WHERE  A.ID=B.合约单位id " & _
                            "AND B.体检状态=[4] " & _
                            "AND B.体检部门id+0=[1] "
                            
            If Val(varParam(0)) = 1 Then
                '按报到时间查正在体检单据
                strSQL = strSQL & "AND B.ID In (Select 登记id From 体检人员档案 Where 体检报到=1 And 体检时间 BETWEEN [2] AND [3]) "
            Else
                strSQL = strSQL & "AND B.体检时间 BETWEEN [2] AND [3] "
            End If
            
            If blnMoveOuted Then
            
                strSQL = strSQL & " UNION ALL " & _
                            "SELECT DISTINCT A.ID," & _
                                        "1 AS 排序1," & _
                                        "'' AS 排序2," & _
                                        "0 AS 上级id," & _
                                        "'' AS 状态," & _
                                        "'' AS 报告," & _
                                        "0 AS 病人id," & _
                                        "B.体检号 AS 体检单号," & _
                                        "NULL+0 AS 门诊号," & _
                                        "'' AS 健康号,'' As 就诊卡号,'' As 体检编号," & _
                                        "A.名称||'('||B.体检号||')' AS 姓名," & _
                                        "'' AS 性别," & _
                                        "'' AS 年龄," & _
                                        "'' AS 婚姻状况," & _
                                        "B.ID AS 登记id, " & _
                                        "1 AS 标志,Null+0 As 次数," & _
                                        "B.体检号 AS 单据号 " & _
                        "FROM   合约单位 A," & _
                                "H体检登记记录 B " & _
                        "WHERE  A.ID=B.合约单位id " & _
                                "AND B.体检状态=[4] " & _
                                "AND B.体检部门id+0=[1] "
                                
                If Val(varParam(0)) = 1 Then
                    '按报到时间查正在体检单据
                    strSQL = strSQL & "AND B.ID In (Select 登记id From 体检人员档案 Where 体检报到=1 And 体检时间 BETWEEN [2] AND [3]) "
                Else
                    strSQL = strSQL & "AND B.体检时间 BETWEEN [2] AND [3] "
                End If
            
            End If
            
            '团体组别
            strSQL = strSQL & " UNION ALL " & _
                        "SELECT DISTINCT A.ID," & _
                                    "1 AS 排序1," & _
                                    "C.组别名称 AS 排序2," & _
                                    "A.ID AS 上级id," & _
                                    "'' AS 状态," & _
                                    "'' AS 报告," & _
                                    "NULL+0 AS 病人id," & _
                                    "B.体检号 AS 体检单号," & _
                                    "0 AS 门诊号,'' As 就诊卡号,'' As 体检编号," & _
                                    "'' AS 健康号," & _
                                    "C.组别名称 AS 姓名," & _
                                    "'' AS 性别," & _
                                    "'' AS 年龄," & _
                                    "'' AS 婚姻状况," & _
                                    "B.ID AS 登记id, " & _
                                    "2 AS 标志,Null+0 as 次数," & _
                                    "B.体检号 AS 单据号 " & _
                    "FROM   合约单位 A," & _
                            "体检登记记录 B,体检组别 C " & _
                    "WHERE  C.登记id=B.ID AND A.ID=B.合约单位id " & _
                            "AND B.体检状态=[4] " & _
                            "AND B.体检部门id+0=[1] "
                            
            If Val(varParam(0)) = 1 Then
                '按报到时间查正在体检单据
                strSQL = strSQL & "AND B.ID In (Select 登记id From 体检人员档案 Where 体检报到=1 And 体检时间 BETWEEN [2] AND [3]) "
            Else
                strSQL = strSQL & "AND B.体检时间 BETWEEN [2] AND [3] "
            End If
                
            If blnMoveOuted Then
                strSQL = strSQL & " UNION ALL " & _
                            "SELECT DISTINCT A.ID," & _
                                        "1 AS 排序1," & _
                                        "C.组别名称 AS 排序2," & _
                                        "A.ID AS 上级id," & _
                                        "'' AS 状态," & _
                                        "'' AS 报告," & _
                                        "NULL+0 AS 病人id," & _
                                        "B.体检号 AS 体检单号," & _
                                        "0 AS 门诊号,'' As 就诊卡号,'' As 体检编号," & _
                                        "'' AS 健康号," & _
                                        "C.组别名称 AS 姓名," & _
                                        "'' AS 性别," & _
                                        "'' AS 年龄," & _
                                        "'' AS 婚姻状况," & _
                                        "B.ID AS 登记id, " & _
                                        "2 AS 标志,Null+0 as 次数," & _
                                        "B.体检号 AS 单据号 " & _
                        "FROM   合约单位 A," & _
                                "H体检登记记录 B,H体检组别 C " & _
                        "WHERE  C.登记id=B.ID AND A.ID=B.合约单位id " & _
                                "AND B.体检状态=[4] " & _
                                "AND B.体检部门id+0=[1] "
                                
                If Val(varParam(0)) = 1 Then
                    '按报到时间查正在体检单据
                    strSQL = strSQL & "AND B.ID In (Select 登记id From 体检人员档案 Where 体检报到=1 And 体检时间 BETWEEN [2] AND [3]) "
                Else
                    strSQL = strSQL & "AND B.体检时间 BETWEEN [2] AND [3] "
                End If
            
            End If
            
            strSQL = "SELECT * FROM (" & strSQL & ") A ORDER BY 排序1,体检单号,上级id,排序2,门诊号"
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.诊疗执行科室
                                                                                        
            '参数:诊疗项目id'病人科室id'开单科室id'查找内容
            
            strSQL = _
                "SELECT A.ID FROM 部门表 A,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=1 AND A.ID=[2]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,床位状况记录 B,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=2 AND A.ID=B.病区id AND B.科室ID=[2]"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=3 AND A.ID=[3]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,诊疗执行科室 B,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=4 AND A.ID=B.执行科室id AND B.病人来源=1 AND B.诊疗项目id=X.ID"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,诊疗执行科室 B,诊疗项目目录 X WHERE X.ID=[1] AND X.执行科室=4 AND " & _
                            "A.ID=B.执行科室id AND B.病人来源 IS NULL AND (B.开单科室id IS NULL OR B.开单科室id=[3]) AND B.诊疗项目id=X.ID "
                            
            If Val(varParam(0)) = 0 Then
            
                strSQL = _
                    "SELECT 1 As 末级,A.编码,A.名称,A.简码,A.ID FROM 部门表 A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.编码) Like [4] OR UPPER(A.简码) Like [4] OR A.名称 Like [4])"
                    
            Else
                strSQL = _
                    "SELECT 1 As 末级,A.编码,A.名称,A.简码,A.ID FROM 部门表 A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.编码) Like [4] OR UPPER(A.简码) Like [4] OR A.名称 Like [4]) Union All " & _
                    "SELECT Distinct 1 As 末级,A.编码,A.名称,A.简码,A.ID FROM 部门表 A,部门性质说明 B WHERE A.ID=B.部门ID And B.服务对象 In (1,3) And A.ID Not IN (" & strSQL & ") AND (UPPER(A.编码) Like [4] OR UPPER(A.简码) Like [4] OR A.名称 Like [4])"
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.收费执行科室
        
            '参数:诊疗项目id'病人科室id'开单科室id'查找内容
            
            strSQL = _
                "SELECT A.ID FROM 部门表 A,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=1 AND A.ID=[2]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,床位状况记录 B,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=2 AND A.ID=B.病区id AND B.科室ID=[2]"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=3 AND A.ID=[3]"
            
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,收费执行科室 B,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=4 AND A.ID=B.执行科室id AND B.病人来源=1 AND B.收费细目id=X.ID"
                
            strSQL = strSQL & " UNION ALL " & _
                "SELECT A.ID FROM 部门表 A,收费执行科室 B,收费项目目录 X WHERE X.ID=[1] AND X.执行科室=4 AND " & _
                            "A.ID=B.执行科室id AND B.病人来源 IS NULL AND (B.开单科室id IS NULL OR B.开单科室id=[3]) AND B.收费细目id=X.ID "
            
            If Val(varParam(0)) = 0 Then
                strSQL = _
                    "SELECT 1 As 末级,A.编码,A.名称,A.ID FROM 部门表 A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.编码) Like [4] OR UPPER(A.简码) Like [4] OR A.名称 Like [4])"
                    
            Else
                strSQL = _
                    "SELECT 1 As 末级,A.编码,A.名称,A.ID FROM 部门表 A WHERE A.ID IN (" & strSQL & ") AND (UPPER(A.编码) Like [4] OR UPPER(A.简码) Like [4] OR A.名称 Like [4]) Union All " & _
                    "SELECT Distinct 1 As 末级,A.编码,A.名称,A.ID FROM 部门表 A,部门性质说明 B WHERE A.ID=B.部门ID And B.服务对象 In (1,3) And A.ID Not IN (" & strSQL & ") AND (UPPER(A.编码) Like [4] OR UPPER(A.简码) Like [4] OR A.名称 Like [4]) "
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case SQL.药品执行科室
            
            strSQL = "SELECT Distinct 1 As 末级,A.编码,A.名称,A.ID " & _
                    "from 部门表 A,部门性质说明 B " & _
                    "where (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                    "and A.ID=B.部门ID and B.服务对象 in (1,3) " & _
                    "and B.工作性质=Decode([1],'5','西药房','6','成药房','7','中药房','4','发料部门')"
                              
    End Select
    
    GetPublicSQL = strSQL
    
    Exit Function
    
errHand:
    
End Function




