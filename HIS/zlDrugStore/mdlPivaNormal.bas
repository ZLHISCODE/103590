Attribute VB_Name = "mdlPivaNormal"
Option Explicit






Public Function Piva_GetMedi(ByVal intStemp As Integer, ByVal strIDS As String, ByVal int显示所有 As Integer) As Recordset
    '取医嘱对应的药品，包括可能其他非静配的药品数据
    'intStemp：0-审核医嘱，1-已通过审核医嘱，2-未通过审核医嘱
    'strIDS：主医嘱ID
    'int显示所有：审核该药房的所有数据，0-不启用，1-启用
    '             启用时可能输液配药记录中无记录
    Dim strTmp As String
    Dim strSqlTmp As String
    
    On Error GoTo errHandle
        
    gstrSQL = "Select Distinct a.Id, a.相关id, a.病人id, a.主页id, a.开嘱医生, a.审查结果, a.药师审核原因, g.病人病区id, g.病人科室id, b.名称 科室名称, f.当前床号 床号, p.配药类型," & vbNewLine & _
        "                Decode(a.医嘱期效, 0, '长期', 1, '临时') 医嘱期效, m.名称 给药途径, g.标识号 As 住院号, a.姓名, a.性别, a.年龄, c.Id 药品id, c.名称 药品名称," & vbNewLine & _
        "                c.规格, a.单次用量, i.计算单位, i.Id 药名id, a.执行频次, Nvl(a.药师审核标志, 0) 审核标志, a.执行时间方案, a.皮试结果, a.开嘱时间," & vbNewLine & _
        "                Nvl(t.是否皮试, 0) 是否皮试, a.执行性质, a.执行标记" & vbNewLine & _
        " From 病人医嘱记录 A, 病人医嘱记录 L, 住院费用记录 G, 部门表 B, 收费项目目录 C, 病人信息 F, 诊疗项目目录 I, 药品规格 J, 药品特性 T, 输液药品属性 P, 诊疗项目目录 M," & vbNewLine & _
        "     Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) K" & vbNewLine & _
        " Where a.相关id = l.Id And a.Id = g.医嘱序号 And a.病人id = f.病人id And a.病人科室id = b.Id And g.收费细目id = c.Id And j.药品id = c.Id And" & vbNewLine & _
        "      j.药品id = p.药品id And j.药名id = i.Id And l.诊疗项目id = m.Id And j.药名id = t.药名id And l.Id = k.Column_Value"
    
    If int显示所有 = 1 Then
        '启用审核药房所有数据，则不需要区分审核状态
    Else
        '不启用审核药房所有数据，则按审核状态来划分
        If intStemp = 1 Then
            '审核医嘱标志：已通过审核
            gstrSQL = gstrSQL & " and (a.药师审核标志=1 or a.药师审核标志=3) "
        ElseIf intStemp = 2 Then
            '审核医嘱标志: 未通过审核
            gstrSQL = gstrSQL & " and (a.药师审核标志=2 or a.药师审核标志=3) "
        Else
            '审核医嘱标志: 未审核
            gstrSQL = gstrSQL & " and (Nvl(a.药师审核标志,0)=0 or a.药师审核标志=3) "
        End If
    End If

    '排除自备药，不取药，后面单独提取
    gstrSQL = gstrSQL & " And Not Exists " & _
        " (Select 1 From 病人医嘱记录 Aa Where Aa.相关id = a.相关id And Aa.执行性质 = 5 And (Aa.执行标记 = 0 Or Aa.执行标记 = 2)) "
 
    If int显示所有 = 0 Then
        '只查询输液单据
        gstrSQL = gstrSQL & " And Exists (Select 1 From 药品收发记录 Aa, 输液配药内容 Bb Where Aa.Id = Bb.收发id And Aa.费用id = g.Id) "
    End If
    
    '合并门诊费用数据
    strTmp = Replace(gstrSQL, "住院费用记录", "门诊费用记录")
    gstrSQL = gstrSQL & " Union All " & strTmp
    
    '合并自备药，不取药，不包含费用数据
    strTmp = "Select Distinct a.Id, a.相关id, a.病人id, a.主页id, a.开嘱医生, a.审查结果, a.药师审核原因, f.当前病区id As 病人病区id, f.当前科室id As 病人科室id," & vbNewLine & _
        "                b.名称 科室名称, f.当前床号 床号, p.配药类型, Decode(a.医嘱期效, 0, '长期', 1, '临时') 医嘱期效, m.名称 给药途径, f.住院号, a.姓名, a.性别, a.年龄," & vbNewLine & _
        "                c.Id 药品id, c.名称 药品名称, c.规格, a.单次用量, i.计算单位, i.Id 药名id, a.执行频次, Nvl(a.药师审核标志, 0) 审核标志, a.执行时间方案, a.皮试结果," & vbNewLine & _
        "                a.开嘱时间, Nvl(t.是否皮试, 0) 是否皮试, a.执行性质, a.执行标记" & vbNewLine & _
        " From 病人医嘱记录 A, 病人医嘱记录 L, 部门表 B, 收费项目目录 C, 病人信息 F, 诊疗项目目录 I, 药品规格 J, 药品特性 T, 输液药品属性 P, 诊疗项目目录 M," & vbNewLine & _
        "     Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) K" & vbNewLine & _
        " Where a.相关id = l.Id And a.病人id = f.病人id And a.病人科室id = b.Id And a.收费细目id = c.Id And j.药品id = c.Id And j.药品id = p.药品id And" & vbNewLine & _
        "      j.药名id = i.Id And l.诊疗项目id = m.Id And j.药名id = t.药名id And l.Id = k.Column_Value And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From 病人医嘱记录 Aa" & vbNewLine & _
        "       Where Aa.相关id = a.相关id And Aa.执行性质 = 5 And (Aa.执行标记 = 0 Or Aa.执行标记 = 2))"

    '合并所有
    gstrSQL = gstrSQL & " Union All " & strTmp
    gstrSQL = gstrSQL & " Order By 科室名称,病人id,相关Id"
    
    Set Piva_GetMedi = zlDatabase.OpenSQLRecord(gstrSQL, "Piva_GetMedi", strIDS)

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function Piva_GetTrans(ByVal strIDS As String, ByVal lng部门id As Long, ByVal strStep As String, ByVal intPack As Integer, ByVal blnShowOhters As Boolean) As ADODB.Recordset
        
    '取输液配药记录
    'lngCenterID：输液配置中心ID
    'str病区ID：病区ID串
    'dateExeStart、dateExeEnd：输液配药单据的执行时间范围
    'strStep(操作类型)：01-摆药印签(1)，02-配药核查(2)，03-发送核查(4)，04-销帐审核(9)，10-审核已通过医嘱(10)，11-审核未通过医嘱(10)，12-已发送查看(5), 13-已签收查看(6)，14-拒绝签收查(7)，15-已作废查看
    '操作类型：1、填制，2、摆药，3、校对，4、配药，5、发送，6、签收，7、拒绝签收  8，确认拒收，9，销帐申请，10，销帐审核
    'intPack传入：0-所有；1-仅配药；2-仅打包
    '是否打包：0-不打包(配液),1-病区打包,2-配置中心打包
    'intShowType:显示样式。0-普通显示；2-增加自备药的显示
    
    Dim strOhterSQL As String       '用于提取医嘱中[自备药]的相关数据
    Dim strTmp As String
    
    On Error GoTo errHandle
    
    If strStep = "15" Then
        '已摆药状态
        '1.销帐审核通过
        gstrSQL = "Select A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,K.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次, A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' NO, F.编码 As 药品编码,' ' 销帐原因, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, e.产地, e.批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, e.频次, '销帐审核通过' As 作废类型, " & _
            " 0 As 发药数量, (e.入出系数*e.实际数量 / G.住院包装) As 数量,e.入出系数*e.实际数量 As 实际数量, G.住院单位 As 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, e.用法, e.药品id,0 as 费用序号,0 As 费用id,null As 险类, A.摆药单号,L.发送时间 As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id, M.id As 对应医嘱ID,A.发送号,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1, m.执行性质, m.执行标记  " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G,输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 住院费用记录 D, 药品收发记录 E, 病人医嘱发送 L ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 k,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V,输液配药内容 Z "

        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID And A.病人科室id = C.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And g.药品id = e.药品id And T.药名id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=K.名称(+) And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID " & _
            " And m.Id = d.医嘱序号 And d.Id = e.费用id And a.医嘱id = l.医嘱id(+) And a.发送号 = l.发送号(+) And a.id = z.记录id And z.收发id = e.Id " & _
            " And a.操作状态=10 And A.id=V.Column_Value  And Exists (Select 1 From 输液配药内容 D, 药品收发记录 E Where d.收发id = e.Id And d.记录id = a.Id)"

        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
        
        '2.销账审核拒绝
        gstrSQL = gstrSQL & " Union All " & _
            "Select A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,K.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次, A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' NO, F.编码 As 药品编码,' ' 销帐原因, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, e.产地, e.批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, e.频次, '销帐审核拒绝' As 作废类型, " & _
            " 0 As 发药数量, (e.入出系数*e.实际数量 / G.住院包装) As 数量,e.入出系数*e.实际数量 As 实际数量, G.住院单位 As 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, e.用法, e.药品id,0 as 费用序号,0 As 费用id,null As 险类, A.摆药单号,L.发送时间 As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id, M.id As 对应医嘱ID,A.发送号,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1, m.执行性质, m.执行标记  " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G,输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 住院费用记录 D, 药品收发记录 E, 病人医嘱发送 L ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 k,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V,输液配药内容 Z "

        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID And A.病人科室id = C.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And g.药品id = e.药品id And T.药名id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=K.名称(+) And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID " & _
            " And m.Id = d.医嘱序号 And d.Id = e.费用id And a.医嘱id = l.医嘱id(+) And a.发送号 = l.发送号(+)  and e.实际数量>0 And a.id = z.记录id And z.收发id = e.Id " & _
            " And a.操作状态=11 And A.id=V.Column_Value And Exists (Select 1 From 输液配药内容 D, 药品收发记录 E Where d.收发id = e.Id And d.记录id = a.Id)"
            
        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
                
        '合并门诊费用
        strTmp = Replace(gstrSQL, "住院费用记录", "门诊费用记录")
        gstrSQL = gstrSQL & " Union All " & strTmp
        
        '未摆药状态
        '按规格
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,K.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' As NO, F.编码 As 药品编码,' ' 销帐原因, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, '' As 产地, '' As 批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, '' As 频次, '未摆药销帐' As 作废类型, " & _
            " 0 As 发药数量, (M.单次用量/ G.剂量系数 / G.住院包装) As 数量,M.单次用量/ G.剂量系数 As 实际数量, G.住院单位 As 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, '' As 用法, M.收费细目id As 药品id,0 as 费用序号,0 As 费用id,null As 险类, " & _
            " A.摆药单号,Null As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,M.id As 对应医嘱ID,A.发送号,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1, m.执行性质, m.执行标记  " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G,输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 k,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "
        
        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID  And A.病人科室id = C.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And M.收费细目id = F.ID And T.药名id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=K.名称(+) And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And a.操作状态=10  " & _
            " And  A.id=V.Column_Value  And Not Exists (Select 1 From 输液配药内容 D, 药品收发记录 E Where d.收发id = e.Id And d.记录id = a.Id) "
        
        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
                
        '按品种
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,K.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' As NO, J.编码 As 药品编码,' ' 销帐原因, " & _
            " J.名称 As 通用名, '' As 商品名, I.名称 As 英文名, '' as 规格, '' As 产地, '' As 批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, '' As 频次, '未摆药销帐' As 作废类型, " & _
            " 0 As 发药数量, 0 As 数量,0 As 实际数量, '' 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, '' As 用法, Decode(Nvl(m.收费细目id, 0), 0, j.Id, m.收费细目id) As 药品id,0 as 费用序号,0 As 费用id,null As 险类, " & _
            " A.摆药单号,Null As 医嘱发送时间,0 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,M.id As 对应医嘱ID,A.发送号,nvl(T.是否皮试,0) 是否皮试,'' As 配药类型1, m.执行性质, m.执行标记  " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C,诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 k,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "

        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID  And A.病人科室id = C.ID and M.收费细目id is null and M.诊疗项目id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=K.名称(+) And a.操作状态=10  " & _
            " And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And J.id = I.诊疗项目id(+) And I.性质(+) = 2 And j.Id = t.药名id " & _
            " And  A.id=V.Column_Value And Not Exists (Select 1 From 输液配药内容 D, 药品收发记录 E Where d.收发id = e.Id And d.记录id = a.Id) "

        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
        
        '提取医嘱中[自备药]的相关数据
        strOhterSQL = "Union All" & vbNewLine & _
                    "Select Distinct a.Id As 配药id, a.批次标记, a.优先级, a.是否确认调整, a.部门id, a.序号, a.配药批次, s.颜色, a.姓名, a.性别, a.年龄, a.住院号, a.床号," & vbNewLine & _
                    "                LPad(a.床号, 10, ' ') 床号排序, k.编码, m.序号 医嘱序号, m.药师审核时间, m.执行频次, a.病人病区id, a.病人科室id, a.执行时间, a.瓶签号, a.打包时间," & vbNewLine & _
                    "                m.病人id, m.主页id, a.是否调整批次, a.是否锁定, a.手工调整批次, '' 拒收原因, a.操作人员, a.操作时间, Nvl(a.打印标志, 0) As 打印标志, a.是否打包," & vbNewLine & _
                    "                b.名称 As 病人病区, c.名称 As 病人科室, 0 As 收发id, 0 As 单据, '' As NO, f.编码 As 药品编码, ' ' 销帐原因, f.名称 As 通用名," & vbNewLine & _
                    "                h.名称 As 商品名, i.名称 As 英文名, f.规格, f.产地, '' As 批号, m.单次用量 As 单量, j.计算单位 As 剂量单位, j.Id 药名id, m.执行频次 As 频次," & vbNewLine & _
                    "                '' As 作废类型, 0 As 发药数量, (m.单次用量 / g.剂量系数 / g.住院包装) As 数量, (m.单次用量 / g.剂量系数) As 实际数量, g.住院单位 As 单位," & vbNewLine & _
                    "                0 As 批次, 0 As 库存数量, Nvl(m.审查结果, -1) 审查结果, Zc.医嘱内容 As 用法, m.收费细目id As 药品id, 0 As 费用序号, 0 As 费用id, o.险类," & vbNewLine & _
                    "                a.摆药单号, r.发送时间 As 医嘱发送时间, Nvl(t.抗生素, '0') 配药类型, t.溶媒, m.皮试结果, m.开嘱时间, a.医嘱id, m.Id As 对应医嘱id, a.发送号," & vbNewLine & _
                    "                Nvl(t.是否皮试, 0) 是否皮试, x.配药类型 As 配药类型1, m.执行性质, m.执行标记 " & vbNewLine & _
                    "From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G, 输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 病案主页 O, 配药工作批次 S," & vbNewLine & _
                    "     药品特性 T, 床位状况记录 Q, 床位编制分类 K, 病人医嘱记录 Zc, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V," & vbNewLine & _
                    "     病人医嘱发送 R"

        strOhterSQL = strOhterSQL & " Where a.医嘱id = m.相关id And (m.执行性质 = 5 And m.执行标记 = 0) And a.部门id = s.配置中心id And" & vbNewLine & _
                    "      a.配药批次 = s.批次 And a.床号 = q.床号(+) And a.病人病区id = q.病区id(+) And q.床位编制 = k.名称(+) And a.病人病区id = b.Id And" & vbNewLine & _
                    "      a.病人科室id = c.Id And m.收费细目id = f.Id And f.Id = g.药品id And g.药品id = h.收费细目id(+) And h.性质(+) = 3 And" & vbNewLine & _
                    "      g.药品id = x.药品id(+) And g.药名id = i.诊疗项目id(+) And i.性质(+) = 2 And g.药名id = j.Id And j.Id = t.药名id And" & vbNewLine & _
                    "      a.医嘱id = Zc.Id And m.病人id = o.病人id(+) And m.主页id = o.主页id(+) And a.医嘱id = r.医嘱id And a.发送号 = r.发送号 And " & vbNewLine & _
                    "      a.Id = v.Column_Value "

        If intPack = 1 Then
            '不打包
            strOhterSQL = strOhterSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            strOhterSQL = strOhterSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
    ElseIf strStep = "16" Then
        '医嘱回退(按规格)
        gstrSQL = " Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,K.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' As NO, F.编码 As 药品编码,' ' 销帐原因, " & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, '' As 产地, '' As 批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, '' As 频次, '医嘱回退' As 作废类型, " & _
            " 0 As 发药数量, (M.单次用量/ G.剂量系数 / G.住院包装) As 数量,M.单次用量/ G.剂量系数 As 实际数量, G.住院单位 As 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, '' As 用法, M.收费细目id As 药品id,0 as 费用序号,0 As 费用id,null As 险类, " & _
            " A.摆药单号,Null As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,M.id As 对应医嘱ID,A.发送号,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1, m.执行性质, m.执行标记 " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G,输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 K,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "
        
        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID  And A.病人科室id = C.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And M.收费细目id = F.ID And T.药名id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=K.名称(+) And " & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And a.操作状态=12  " & _
            " And A.id=V.Column_Value "
        
        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
                
        '合并医嘱回退(按品种发送)
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次,S.颜色, A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,K.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,M.病人id,M.主页id,A.是否调整批次,A.是否锁定,A.手工调整批次,'' 拒收原因," & _
            " A.操作人员,A.操作时间, Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, 0 As 收发id, 9 As 单据, '' As NO, J.编码 As 药品编码,' ' 销帐原因, " & _
            " J.名称 As 通用名, '' As 商品名, I.名称 As 英文名, '' as 规格, '' As 产地, '' As 批号, M.单次用量 As 单量, J.计算单位 As 剂量单位,J.id 药名id, '' As 频次, '医嘱回退' As 作废类型, " & _
            " 0 As 发药数量, 0 As 数量,0 As 实际数量, '' 单位,0 As 批次, 0 As 库存数量, Nvl(M.审查结果,-1) 审查结果, '' As 用法, Decode(Nvl(m.收费细目id, 0), 0, j.Id, m.收费细目id) As 药品id,0 as 费用序号,0 As 费用id,null As 险类, " & _
            " A.摆药单号,Null As 医嘱发送时间,0 配药类型,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,M.id As 对应医嘱ID,A.发送号,nvl(T.是否皮试,0) 是否皮试,'' As 配药类型1, m.执行性质, m.执行标记 " & _
            " From 输液配药记录 A, 部门表 B, 部门表 C,诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M ,配药工作批次 S,药品特性 T,床位状况记录 O,床位编制分类 K,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "

        gstrSQL = gstrSQL & " Where A.医嘱id = M.相关id And A.病人病区id = B.ID  And A.病人科室id = C.ID and M.收费细目id is null and M.诊疗项目id=J.id And A.床号=O.床号(+) And  A.病人病区id=O.病区id(+) And A.病人科室id=O.科室id(+) and O.床位编制=K.名称(+) And a.操作状态=12  " & _
            " And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) And J.id = I.诊疗项目id(+) And I.性质(+) = 2 And j.Id = t.药名id " & _
            " And A.id=V.Column_Value "

        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
    Else
        '其他
        gstrSQL = "Select Distinct A.ID As 配药ID,A.批次标记,A.优先级,A.是否确认调整, A.部门id, A.序号, A.配药批次, S.颜色,A.姓名, A.性别, A.年龄, A.住院号,A.床号,LPad(A.床号, 10, ' ') 床号排序,K.编码,M.序号 医嘱序号,M.药师审核时间,M.执行频次,  A.病人病区id, A.病人科室id, A.执行时间, A.瓶签号,A.打包时间,A.是否调整批次,A.是否锁定,A.手工调整批次," & IIf(strStep = "13", "W.操作说明 拒收原因,", "'' 拒收原因,") & _
            "  A.操作人员,A.操作时间,Nvl(A.打印标志,0) As 打印标志, A.是否打包, B.名称 As 病人病区, C.名称 As 病人科室, D.收发id, E.单据, E.NO, F.编码 As 药品编码, " & IIf(strStep = "04", "Y.销帐原因, ", "' ' 销帐原因,") & _
            " F.名称 As 通用名, H.名称 As 商品名, I.名称 As 英文名, F.规格, E.产地, E.批号, E.单量, J.计算单位 As 剂量单位,J.id 药名id, E.频次, '' As 作废类型, " & _
            " Case Nvl(E.审核人, '未审核') When '未审核' Then E.实际数量 * Nvl(E.付数, 1) / G.住院包装 Else 0 End As 发药数量,M.病人id,M.主页id,T.溶媒,M.皮试结果,M.开嘱时间,A.医嘱id,M.id As 对应医嘱ID,A.发送号, " & _
            " (D.数量 / G.住院包装)  As 数量,D.数量 As 实际数量, G.住院单位 As 单位,Nvl(E.批次,0) As 批次, Nvl(L.实际数量, 0)/ G.住院包装 As 库存数量, Nvl(M.审查结果,-1) 审查结果, E.用法, E.药品id, n.序号 As 费用序号,E.费用id, o.险类, A.摆药单号,r.发送时间 As 医嘱发送时间,nvl(T.抗生素,'0') 配药类型,nvl(T.是否皮试,0) 是否皮试,x.配药类型 As 配药类型1, m.执行性质, m.执行标记  " & _
            " From  输液配药记录 A, 部门表 B, 部门表 C, 输液配药内容 D, 药品收发记录 E, 收费项目目录 F, 药品规格 G,输液药品属性 X,  收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 住院费用记录 N, 病案主页 O ,配药工作批次 S,药品特性 T,床位状况记录 Q,床位编制分类 K,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "
        
        '提取医嘱中[自备药]的相关数据
        strOhterSQL = " Union All "
        strOhterSQL = strOhterSQL & "Select Distinct a.Id As 配药id, a.批次标记, a.优先级, a.是否确认调整, a.部门id, a.序号, a.配药批次, s.颜色, a.姓名, a.性别, a.年龄, a.住院号, a.床号," & vbNewLine & _
                    "                LPad(a.床号, 10, ' ') 床号排序, k.编码, m.序号 医嘱序号, m.药师审核时间, m.执行频次, a.病人病区id, a.病人科室id, a.执行时间, a.瓶签号, a.打包时间," & vbNewLine & _
                    "                a.是否调整批次, a.是否锁定, a.手工调整批次, '' 拒收原因, a.操作人员, a.操作时间, Nvl(a.打印标志, 0) As 打印标志, a.是否打包, b.名称 As 病人病区," & vbNewLine & _
                    "                c.名称 As 病人科室, 0 As 收发id, 0 As 单据, '' As NO, f.编码 As 药品编码, ' ' 销帐原因, f.名称 As 通用名, h.名称 As 商品名," & vbNewLine & _
                    "                i.名称 As 英文名, f.规格, f.产地, '' As 批号, m.单次用量 As 单量, j.计算单位 As 剂量单位, j.Id 药名id, m.执行频次 As 频次, '' As 作废类型," & vbNewLine & _
                    "                0 As 发药数量, m.病人id, m.主页id, t.溶媒, m.皮试结果, m.开嘱时间, a.医嘱id, m.Id As 对应医嘱id, a.发送号," & vbNewLine & _
                    "                (m.单次用量 / g.剂量系数 / g.住院包装) As 数量, (m.单次用量 / g.剂量系数) As 实际数量, g.住院单位 As 单位, 0 As 批次, 0 As 库存数量," & vbNewLine & _
                    "                Nvl(m.审查结果, -1) 审查结果, Zc.医嘱内容 As 用法, m.收费细目id As 药品id, 0 As 费用序号, 0 As 费用id, o.险类, a.摆药单号," & vbNewLine & _
                    "                r.发送时间 As 医嘱发送时间, Nvl(t.抗生素, '0') 配药类型, Nvl(t.是否皮试, 0) 是否皮试, x.配药类型 As 配药类型1, m.执行性质, m.执行标记 " & vbNewLine & _
                    " From 输液配药记录 A, 部门表 B, 部门表 C, 收费项目目录 F, 药品规格 G, 输液药品属性 X, 收费项目别名 H, 诊疗项目别名 I, 诊疗项目目录 J, 病人医嘱记录 M, 病案主页 O, 配药工作批次 S," & vbNewLine & _
                    "     药品特性 T, 床位状况记录 Q, 床位编制分类 K, 病人医嘱记录 Zc, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V," & vbNewLine & _
                    "     病人医嘱发送 R"
        
        If strStep = "13" Then gstrSQL = gstrSQL & ",输液配药状态 W "
        
        If strStep = "04" Then gstrSQL = gstrSQL & ",病人费用销帐 Y "
        
'        If strStep = "01" And bln启用审方 Then
'            gstrSQL = gstrSQL & ",处方审查记录 Q,处方审查明细 K "
'        End If
        
        gstrSQL = gstrSQL & ",(Select 库房id, 药品id, Nvl(批次, 0) As 批次, Nvl(实际数量, 0) As 实际数量 " & _
            " From 药品库存 Where 性质 = 1 And 库房id = [2]) L, 药品收发记录 P, 病人医嘱发送 R " & IIf(strStep = "04", ", 输液配药状态 U ", "")
        
        gstrSQL = gstrSQL & " Where A.病人病区id = B.ID And A.病人科室id = C.ID And A.ID = D.记录id And D.收发id = E.ID And E.药品id = F.ID And F.ID = G.药品id And G.药品id=X.药品id(+) And E.费用id = N.ID And N.医嘱序号 = M.ID And " & IIf(strStep = "13", "W.配药id=A.id And A.操作状态=W.操作类型 And A.操作时间=W.操作时间 And ", "") & _
            " G.药品id = H.收费细目id(+) And H.性质(+) = 3 And G.药名id = I.诊疗项目id(+) And I.性质(+) = 2 And G.药名id = J.ID And T.药名id=J.ID And A.配药批次=S.批次(+) And a.部门id = s.配置中心id(+) " & IIf(strStep = "04", " And Y.费用id=N.id And y.申请时间 = u.操作时间 And u.配药id = v.Column_Value ", "") & _
            " And E.库房id = L.库房id(+) And E.药品id = L.药品id(+) And A.床号=Q.床号(+) And  A.病人病区id=Q.病区id(+) And A.病人科室id=Q.科室id(+) and Q.床位编制=K.名称(+) And Nvl(E.批次, 0) = L.批次(+) " & _
            " And n.病人id = o.病人id(+) And n.主页id = o.主页id(+) And a.医嘱id = r.医嘱id And a.发送号 = r.发送号  " & _
            " And e.单据 = p.单据 And e.No = p.No And e.库房id + 0 = p.库房id+0 And e.药品id + 0 = p.药品id+0 And e.序号 = p.序号 And (p.记录状态 = 1 Or Mod(p.记录状态, 3) = 0) And A.id=V.Column_Value  "
            
        strOhterSQL = strOhterSQL & " Where a.医嘱id = m.相关id And (m.执行性质 = 5 And m.执行标记 = 0) And a.部门id = s.配置中心id And" & vbNewLine & _
                    "      a.配药批次 = s.批次 And a.床号 = q.床号(+) And a.病人病区id = q.病区id(+) And q.床位编制 = k.名称(+) And a.病人病区id = b.Id And" & vbNewLine & _
                    "      a.病人科室id = c.Id And m.收费细目id = f.Id And f.Id = g.药品id And g.药品id = h.收费细目id(+) And h.性质(+) = 3 And" & vbNewLine & _
                    "      g.药品id = x.药品id(+) And g.药名id = i.诊疗项目id(+) And i.性质(+) = 2 And g.药名id = j.Id And j.Id = t.药名id And" & vbNewLine & _
                    "      a.医嘱id = Zc.Id And m.病人id = o.病人id(+) And m.主页id = o.主页id(+) And a.医嘱id = r.医嘱id And a.发送号 = r.发送号 And " & vbNewLine & _
                    "      a.Id = v.Column_Value"

        If strStep = "01" Then
            '待摆药
            gstrSQL = gstrSQL & " And A.操作状态=1 "
            strOhterSQL = strOhterSQL & " And A.操作状态=1 "
        ElseIf strStep = "02" Then
            '待配药
            gstrSQL = gstrSQL & " And A.操作状态=2 "
            strOhterSQL = strOhterSQL & " And A.操作状态=2 "
        ElseIf strStep = "03" Then
            '待发送
            gstrSQL = gstrSQL & " And A.操作状态=4 "
            strOhterSQL = strOhterSQL & " And A.操作状态=4 "
        ElseIf strStep = "11" Then
            '已销账审核
            gstrSQL = gstrSQL & " And A.操作状态=10 "
            strOhterSQL = strOhterSQL & " And A.操作状态=10 "
        ElseIf strStep = "12" Then
            '已发送
            gstrSQL = gstrSQL & " And A.操作状态=5 "
            strOhterSQL = strOhterSQL & " And A.操作状态=5 "
        ElseIf strStep = "13" Then
            '已签收
            gstrSQL = gstrSQL & " And A.操作状态=6 "
            strOhterSQL = strOhterSQL & " And A.操作状态=6 "
        ElseIf strStep = "14" Then
            '已拒绝签收
            gstrSQL = gstrSQL & " And A.操作状态=7 "
            strOhterSQL = strOhterSQL & " And A.操作状态=7 "
        ElseIf strStep = "04" Then
            gstrSQL = gstrSQL & " And A.操作状态=9 "
            strOhterSQL = strOhterSQL & " And A.操作状态=9 "
        End If
        
        If intPack = 1 Then
            '不打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0)=0 "
            strOhterSQL = strOhterSQL & " And Nvl(A.是否打包,0)=0 "
        ElseIf intPack = 2 Then
            '打包：包括病区打包和配置中心打包
            gstrSQL = gstrSQL & " And Nvl(A.是否打包,0) In (1,2) "
            strOhterSQL = strOhterSQL & " And Nvl(A.是否打包,0) In (1,2) "
        End If
        
        '合并门诊费用
        strTmp = Replace(gstrSQL, "住院费用记录", "门诊费用记录")
        gstrSQL = gstrSQL & " Union All " & strTmp
    End If
    
    If blnShowOhters Then
        '合并SQL
        gstrSQL = gstrSQL & strOhterSQL
        '排序
        gstrSQL = gstrSQL & " Order By 配药id "
    End If
    
    Set Piva_GetTrans = zlDatabase.OpenSQLRecord(gstrSQL, "读取输液配药记录", strIDS, lng部门id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function PIVA_GetExcStatus(ByVal str配药ids As String, ByVal intStatus As Integer) As ADODB.Recordset
    '检查不符合当前状态的输液单
    'str配药ids：输液单ID串
    'intStatus：当前应该的业务状态
    Dim i As Integer
    Dim arrExecute As Variant
    
    On Error GoTo errHandle
    arrExecute = GetArrayByStr(str配药ids, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = " Select ID, 瓶签号, 操作状态,是否打包 " & _
            " From 输液配药记录 Where (操作状态 <> [2] " & IIf(intStatus = 2, " or 是否打包<>0", "") & ") And ID In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) "
        Set PIVA_GetExcStatus = zlDatabase.OpenSQLRecord(gstrSQL, "PIVA_GetStatus", CStr(arrExecute(i)), intStatus)
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function PIVA_GetTransCount(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal dateExeEnd As Date, _
    ByVal bln审核 As Boolean, ByVal bln启用审方 As Boolean, Optional ByVal intType As Integer, _
    Optional ByVal strMsg As String, Optional ByVal lng药品id As Long, Optional ByVal str瓶签号 As String, _
    Optional ByVal lng科室ID As Long, Optional intCheck As Integer, Optional ByVal strSourceDep As String) As ADODB.Recordset
    '取病区输液单数目
    'lngCenterID：输液配置中心ID
    'dateExeStart、dateExeEnd：输液配药单据的执行时间范围
    On Error GoTo errHandle
    
    gstrSQL = "select 类型, 病区id, 病区,  id,药师审核标志,名称,编码 from " & _
        " (with W as (Select Distinct a.操作状态,c.药师审核标志,a.病人病区id As 病区id, '[' || b.编码 || ']' || b.名称 As 病区,b.名称,b.编码, c.相关id As 医嘱id,A.id,A.瓶签号 " & vbNewLine & _
        "       From 输液配药记录 A, 部门表 B, 病人医嘱记录 C" & IIf(bln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & vbNewLine & _
        "       Where a.病人病区id = b.Id And a.医嘱id = c.相关id And c.执行性质 <> 5 And a.部门id = [1] And" & IIf(bln启用审方, " c.id=k.医嘱id and Q.id=K.审方id and K.最后提交=1 and Q.审查结果=1 and", "") & vbNewLine & _
        "             a.执行时间 Between [2] And [3] And "

    If strMsg <> "" Then
        gstrSQL = gstrSQL & IIf(intType = 1, "a.姓名=[4] and ", IIf(intType = 2, "a.床号=[4] and ", "a.住院号=[4] and "))
    End If
    
    If lng药品id <> 0 Then
        gstrSQL = gstrSQL & "C.收费细目id=[6] And "
    End If
    
    If lng科室ID <> 0 Then
        gstrSQL = gstrSQL & "C.病人科室id=[7] And "
    End If
    
    gstrSQL = gstrSQL & " Exists" & vbNewLine & _
        "        (Select 1 From 输液配药内容 D Where d.记录id = a.Id))," & vbNewLine & _
        "       R as (Select Distinct a.操作状态, a.病人病区id As 病区id, '[' || b.编码 || ']' || b.名称 As 病区,b.名称,b.编码," & vbNewLine & _
        "                                 a.Id" & vbNewLine & _
        "                 From 输液配药记录 A, 部门表 B,输液配药内容 C,药品收发记录 D" & vbNewLine & _
        "                 Where a.病人病区id = b.Id  And a.部门id = [1] and A.id=C.记录id and C.收发id=D.id And" & vbNewLine & _
        "                       a.执行时间 Between [2]  and" & vbNewLine & _
        "                       [3] And "
        
    If strMsg <> "" Then
        gstrSQL = gstrSQL & IIf(intType = 1, "a.姓名=[4] and ", IIf(intType = 2, "a.床号=[4] and ", "a.住院号=[4] and "))
    End If
    
    If lng药品id <> 0 Then
        gstrSQL = gstrSQL & "D.药品id=[6] And "
    End If
    
    If lng科室ID <> 0 Then
        gstrSQL = gstrSQL & "A.病人科室id=[7] And "
    End If
    
    If str瓶签号 <> "" Then
        gstrSQL = gstrSQL & "A.瓶签号=[5] And "
    End If
    
    gstrSQL = gstrSQL & "Exists  (Select 1 From 输液配药内容 D Where d.记录id = a.Id))"
    
    If bln审核 = True And intCheck <> 0 Then
        gstrSQL = gstrSQL & ", T as (Select distinct D.病人病区id As 病区id, '[' || B.编码 || ']' || B.名称 As 病区, c.相关id As id,c.药师审核标志,B.名称,b.编码, a.库房id, d.姓名," & _
            " d.床号, d.标识号, d.收费细目id, d.病人科室id " & _
            " From 药品收发记录 A, 部门表 B,病人医嘱记录 C,住院费用记录 D " & IIf(bln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & vbNewLine & _
            " Where D.病人病区id = B.ID And D.医嘱序号=C.id And A.费用id=D.id And A.单据=9 And C.执行性质<>5  And 库房id = [1] " & _
            " And A.填制日期 Between [2] And [3] " & IIf(bln启用审方, " And c.id=k.医嘱id and Q.id=K.审方id and K.最后提交=1 and Q.审查结果=1", "") & _
            " Union all " & _
            " Select distinct D.病人病区id As 病区id, '[' || B.编码 || ']' || B.名称 As 病区, c.相关id As id,c.药师审核标志,B.名称,b.编码, a.库房id, d.姓名," & _
            " '' as 床号, d.标识号, d.收费细目id, d.病人科室id " & _
            " From 药品收发记录 A, 部门表 B,病人医嘱记录 C,门诊费用记录 D " & IIf(bln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & vbNewLine & _
            " Where D.病人病区id = B.ID And D.医嘱序号=C.id And A.费用id=D.id And A.单据=9 And C.执行性质<>5  And 库房id = [1] " & _
            " And A.填制日期 Between [2] And [3] " & IIf(bln启用审方, " And c.id=k.医嘱id and Q.id=K.审方id and K.最后提交=1 and Q.审查结果=1", "") & ")"
    End If
    
    If bln审核 = True Then
        '审核医嘱
        If intCheck = 0 Then
            gstrSQL = gstrSQL & " select Distinct '00' 类型,病区id,病区,医嘱id id,药师审核标志,名称,编码 from  W where (Nvl(药师审核标志, 0) = 0 or Nvl(药师审核标志, 0)=3) and 操作状态=1" & vbNewLine & _
            "union all"
        Else
            gstrSQL = gstrSQL & " Select Distinct '00' 类型, 病区id, 病区, ID, 药师审核标志, 名称, 编码 " & _
                " From t Where 1=1 "
            If strMsg <> "" Then
                gstrSQL = gstrSQL & IIf(intType = 1, " And t.姓名=[4] ", IIf(intType = 2, " And t.床号=[4] ", " And t.标识号=[4] "))
            End If

            If lng药品id <> 0 Then
                gstrSQL = gstrSQL & " And t.收费细目id=[6]"
            End If

            If lng科室ID <> 0 Then
                gstrSQL = gstrSQL & " And t.病人科室id=[7]"
            End If
            
            gstrSQL = gstrSQL & " union all"
        End If
            
         '摆药
        gstrSQL = gstrSQL & " select distinct '01' 类型,病区id,病区,id,1 药师审核标志,名称,编码 from  W where Nvl(药师审核标志, 0) =1 and 操作状态=1 "
        
        If str瓶签号 <> "" Then
            gstrSQL = gstrSQL & " and 瓶签号=[5]"
        End If
    Else
        '摆药
        gstrSQL = gstrSQL & " select distinct '01' 类型,病区id,病区,id,1 药师审核标志,名称,编码 from  W where 操作状态=1"
        
        If str瓶签号 <> "" Then
            gstrSQL = gstrSQL & " And 瓶签号=[5]"
        End If
    End If
    '配药
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '02' 类型,病区id,病区,id,1 药师审核标志,名称,编码 from  R where 操作状态=2"
    '发送
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '03' 类型,病区id,病区 ,id,1 药师审核标志,名称,编码 from  R where 操作状态=4"

    '销账审核
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '04' 类型,病区id,病区,id,1 药师审核标志,名称,编码 from  R where 操作状态=9"
        
    If bln审核 = True Then
        If intCheck = 0 Then
            '已审核通过医嘱查看
            gstrSQL = gstrSQL & " Union All " & _
                "select Distinct  '10' 类型,病区id,病区,医嘱id,1 药师审核标志,名称,编码 from  W where  Nvl(药师审核标志, 0) =1"
                
            '未审核通过医嘱查看
            gstrSQL = gstrSQL & " Union All " & _
                "select Distinct  '11' 类型,病区id,病区,医嘱id,1 药师审核标志,名称,编码 from  W where Nvl(药师审核标志, 0) =2"
        Else
            gstrSQL = gstrSQL & " Union All " & _
                " Select Distinct '10' 类型, 病区id, 病区, ID, 药师审核标志, 名称, 编码 " & _
                " From t Where t.药师审核标志=1 "
            If strMsg <> "" Then
                gstrSQL = gstrSQL & IIf(intType = 1, " And t.姓名=[4] ", IIf(intType = 2, " And t.床号=[4] ", " And t.标识号=[4] "))
            End If

            If lng药品id <> 0 Then
                gstrSQL = gstrSQL & " And t.收费细目id=[6]"
            End If

            If lng科室ID <> 0 Then
                gstrSQL = gstrSQL & " And t.病人科室id=[7]"
            End If

            gstrSQL = gstrSQL & " Union All " & _
                " Select Distinct '11' 类型, 病区id, 病区, ID, 药师审核标志, 名称, 编码 " & _
                " From t Where t.药师审核标志=2 "
            If strMsg <> "" Then
                gstrSQL = gstrSQL & IIf(intType = 1, " And t.姓名=[4] ", IIf(intType = 2, " And t.床号=[4] ", " And t.标识号=[4] "))
            End If

            If lng药品id <> 0 Then
                gstrSQL = gstrSQL & " And t.收费细目id=[6]"
            End If

            If lng科室ID <> 0 Then
                gstrSQL = gstrSQL & " And t.病人科室id=[7]"
            End If
        End If

    End If
    '已发送查看
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '12' 类型,病区id,病区,id,1 药师审核标志,名称,编码 from  R where 操作状态=5"

    '已签收查看
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '13' 类型,病区id,病区,id,1 药师审核标志,名称,编码 from  R where 操作状态=6"

    '拒绝签收查看
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '14' 类型,病区id,病区,id,1 药师审核标志,名称,编码 from  R where 操作状态=7"
    '已作废审核查看
    gstrSQL = gstrSQL & " Union All " & _
        "Select distinct '15' As 类型,a.病人病区id As 病区id, '[' || b.编码 || ']' || b.名称 As 病区,a.Id,1 药师审核标志,名称,b.编码" & vbNewLine & _
        "       From (Select a.ID, 病人病区id" & vbNewLine & _
        "              From 输液配药记录 A,病人医嘱记录 B" & vbNewLine & _
        "              Where A.医嘱id=B.相关id and a.部门id = [1] And a.执行时间 Between [2] And [3] And Nvl(a.操作状态, 0) In (10,11)"
    
    If strMsg <> "" Then
        gstrSQL = gstrSQL & IIf(intType = 1, " And a.姓名=[4] ", IIf(intType = 2, " And a.床号=[4] ", " And a.住院号=[4] "))
    End If
    
    If lng药品id <> 0 Then
        gstrSQL = gstrSQL & " And b.收费细目id=[6]"
    End If
    
    If lng科室ID <> 0 Then
        gstrSQL = gstrSQL & " And a.病人科室id=[7]"
    End If
            
    If str瓶签号 <> "" Then
        gstrSQL = gstrSQL & " And a.瓶签号=[5]"
    End If
        
       gstrSQL = gstrSQL & ") A, 部门表 B" & vbNewLine & _
        "       Where a.病人病区id = b.Id"
        
    '医嘱回退查看
    gstrSQL = gstrSQL & " Union All " & _
        " Select distinct '16' As 类型, A.病人病区id As 病区id, '[' || B.编码 || ']' || B.名称 As 病区, A.ID,1 药师审核标志,b.名称,b.编码 " & _
        " From 输液配药记录 A, 部门表 B ,病人医嘱记录 C" & _
        " Where A.病人病区id = B.ID and A.医嘱id=c.相关id And A.操作状态=12 And A.部门id = [1] And A.执行时间 Between [2] And [3] "
        
    If strMsg <> "" Then
        gstrSQL = gstrSQL & IIf(intType = 1, " And a.姓名=[4] ", IIf(intType = 2, " And a.床号=[4] ", " And a.住院号=[4] "))
    End If
    
    If lng药品id <> 0 Then
        gstrSQL = gstrSQL & " And c.收费细目id=[6]"
    End If
    
    If lng科室ID <> 0 Then
        gstrSQL = gstrSQL & " And a.病人科室id=[7]"
    End If
            
    If str瓶签号 <> "" Then
        gstrSQL = gstrSQL & " And a.瓶签号=[5]"
    End If
    
    gstrSQL = gstrSQL & " Order By 类型, 名称 )" & IIf(strSourceDep = "", "", " N, Table(Cast(f_Num2list([8]) As Zltools.t_Numlist)) S Where n.病区id = s.Column_Value ")
        
    Set PIVA_GetTransCount = zlDatabase.OpenSQLRecord(gstrSQL, "取病区输液单数目", lngCenterID, dateExeStart, dateExeEnd, strMsg, str瓶签号, lng药品id, lng科室ID, strSourceDep)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function





Public Sub PIVA_AnalysisTrans(ByVal lngCenterID As Long, ByVal dateStart As String, ByVal dateEnd As String)
    'PIVA后台工作：分解发药单，产生输液单
    'lngCenterID：输液配置中心ID
    'dateStart、dateEnd：发药单据的填制时间范围
    On Error GoTo ErrHand
    gstrSQL = "Zl_输液配药记录_Insert("
    '配置中心ID
    gstrSQL = gstrSQL & lngCenterID
    '开始时间
    gstrSQL = gstrSQL & "," & dateStart
    '结束时间
    gstrSQL = gstrSQL & "," & dateEnd
    gstrSQL = gstrSQL & ")"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "产生输液配药记录")
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function DeptSendWork_Get科室名称() As Recordset
'获取病人科室名称，取工作性质为临床或护理的部门
    On Error GoTo ErrHand
    
    gstrSQL = "Select distinct a.Id, a.编码, a.名称,zlSpellCode(a.名称) 简码,zlWBCode(a.名称) 五笔简码, a.撤档时间" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where a.Id = b.部门id And (b.工作性质 = '临床' Or b.工作性质 = '护理') And" & vbNewLine & _
            "      (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000/1/1', 'yyyy/mm/dd'))"
    
    
    Set DeptSendWork_Get科室名称 = zlDatabase.OpenSQLRecord(gstrSQL, "获取科室信息")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DeptSendWork_Get配药类型() As Recordset
'获取药品的配药类型
    On Error GoTo ErrHand
    gstrSQL = "select 编码,名称 from 输液配药类型"
    
    Set DeptSendWork_Get配药类型 = zlDatabase.OpenSQLRecord(gstrSQL, "获取配药类型")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get频次() As Recordset
'获取药品的配药类型
    On Error GoTo ErrHand
    gstrSQL = "select 编码,名称,英文名称 from 诊疗频率项目 where 编码 not like '-%'"
    
    Set DeptSendWork_Get频次 = zlDatabase.OpenSQLRecord(gstrSQL, "获取频次")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function PIVA_已摆药输液单(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal lng病人id As Long, ByVal lng主页ID As Long) As Recordset
    '获取该病人当天的已经摆药还未配药的输液单
    Dim strTmp As String
    
    On Error GoTo errHandle

    gstrSQL = "Select Distinct a.Id As 配药id, a.操作状态, a.配药批次, a.执行时间, a.瓶签号, a.操作人员, a.操作时间, a.是否打包, a.摆药单号, e.No, f.名称 As 通用名, f.规格," & vbNewLine & _
        "                e.单量, j.计算单位 As 剂量单位, (d.数量 / g.住院包装) As 数量, g.住院单位 As 单位, r.发送时间 As 医嘱发送时间" & vbNewLine & _
        " From 输液配药记录 A, 输液配药内容 D, 药品收发记录 E, 收费项目目录 F, 药品规格 G, 诊疗项目目录 J," & vbNewLine & _
        "     (Select ID From 病人医嘱记录 Where 病人id = [4] And 主页id = [5] And 诊疗类别 = 'E') M, 病人医嘱发送 R" & vbNewLine & _
        " Where a.Id = d.记录id And d.收发id = e.Id And e.药品id = f.Id And f.Id = g.药品id And g.药名id = j.Id And a.部门id = [1] And" & vbNewLine & _
        "      a.医嘱id = r.医嘱id And a.发送号 = r.发送号 And a.执行时间 Between [2] And [3] And" & vbNewLine & _
        "      ((a.操作状态 > 1 And a.操作状态 < 6) Or (a.操作状态 = 1 And a.是否确认调整 = 1)) And a.医嘱id = m.Id "
    
    Set PIVA_已摆药输液单 = zlDatabase.OpenSQLRecord(gstrSQL, "读取输液配药记录", lngCenterID, _
        CDate(Format(dateExeStart, "yyyy-mm-dd 00:00:00")), CDate(Format(dateExeStart, "yyyy-mm-dd 23:59:59")), _
        lng病人id, lng主页ID)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function






