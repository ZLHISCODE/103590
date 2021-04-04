Attribute VB_Name = "mdlStuffRx"
Option Explicit
'定义参数结构
Private Type TYPE_Para
    卫材单位 As Integer
    单据类型 As String
    收费单据 As Integer
    bln允许未收费单据发料 As Boolean
End Type

Private Enum mFindType
    单据号 = 0
    门诊号 = 1
    姓名 = 2
    身份证 = 3
    IC卡 = 4
    医保号 = 5
    住院号 = 6
End Enum

Private T_Para As TYPE_Para
Private mlngModule As Long

Private mstrOracleMoneyForamt As String

Public Sub GetPara(ByVal lngModule As Long)
'获取本地参数
    On Error GoTo errHandle
    With T_Para
        .卫材单位 = Val(zlDataBase.GetPara("卫材单位", glngSys, lngModule, "0"))
        .收费单据 = zlDataBase.GetPara("收费处方显示方式", glngSys, lngModule, "0")
        .单据类型 = zlDataBase.GetPara("查询业务类型", glngSys, lngModule, "0")
        .bln允许未收费单据发料 = zlDataBase.GetPara("允许未收费的门诊划价处方发料", glngSys, lngModule, "0")
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetCheckPara(ByVal lng部门ID As Long) As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取库存检查参数
    '入参:
    '出参:
    '返回:0-不检查，1-检查，不足提醒，2-不足禁止发料
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(检查方式,0) 库存检查 From 材料出库检查 Where 库房ID=[1]"

    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "GetCheckPara_库存检查参数", lng部门ID)
    With rsTemp
        If Not .EOF Then
            GetCheckPara = Nvl(!库存检查, 0)
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Stuff_GetDept(ByVal strPrivs As String) As Recordset
    On Error GoTo errHandle
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.简码 || '-' || a.名称 As 名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称 And (a.站点=[2] or a.站点 is null) " & _
        "       AND b.编码 ='W' " & _
        "       AND a.id = c.部门id " & _
        "       AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
        IIf(InStr(strPrivs, "所有部门") <> 0, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])") & _
        " Order by a.简码 || '-' || a.名称"
    
    Set Stuff_GetDept = zlDataBase.OpenSQLRecord(gstrSQL, "获取相应的库房_Stuff_GetDept", UserInfo.Id, gstrNodeNo)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Stuff_GetPrePeople(ByVal strKey As String) As Recordset
    On Error GoTo errHandle
    
    gstrSQL = "" & _
        "   Select distinct a.编号 as 编码,A.姓名 As 名称,简码" & _
        "   From 人员表 A,部门人员 B,部门性质说明 C,人员性质说明 D " & _
        "   Where A.Id=B.人员id And B.部门id=C.部门Id And D.人员id=A.Id " & _
        "       And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) AND B.部门id in (Select 部门ID From 部门人员 where 人员id=[2] ) "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "    And  ((A.姓名) like [1] or  A.编号  like [1] or  简码  like  upper([1]))  " & _
        "    "
    End If
    
    gstrSQL = "Select rownum as ID,a.* from (" & gstrSQL & ") A" & _
        "   ORDER BY 编码 "
        
    Set Stuff_GetPrePeople = zlDataBase.OpenSQLRecord(gstrSQL, "获取配料人员_Stuff_GetPrePeople", UserInfo.Id, gstrNodeNo)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Stuff_RxValied(ByVal strPrivs As String, ByVal intType As Integer, ByVal lng库房ID As Long, ByVal int单据 As Integer, ByVal strNo As String, ByVal rsData As Recordset) As Boolean
'检查当前数据是否可以进行发料操作
'返回值为Boolean类型：true-可以发料，false-不能发料
'rsTemp参数：传入本次发料的数据
'参数;strPrivs-权限字符串
'strIDS-收发id串
'intType-业务类型，1-发料，2-退料
    Dim strTemp As String
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim int库存检查 As Integer
    
    On Error GoTo errHandle
    If strNo = "" Then Exit Function
    
'1 检查单据是否存在
    gstrSQL = " Select 1 From 药品收发记录" & _
             " Where No=[1] And (库房ID=[3] Or 库房ID Is NULL) And 单据=[2]"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Stuff_RxValied_检查单据是否存在", strNo, int单据, lng库房ID)
    
    If rsTemp.EOF Then
        MsgBox "该单据不存在，请检查单据信息！"
        Stuff_RxValied = False
        Exit Function
    End If
    
'2 检查单据是否已经进行了相应的操作
    If intType = 0 Then
        gstrSQL = " Select 单据,已收费 From 未发药品记录" & _
                 " Where No=[1] And (库房ID=[3] Or 库房ID Is NULL) And 单据=[2]"
        
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Stuff_RxValied_检查单据是否存在", strNo, int单据, lng库房ID)
        
        If rsTemp.EOF Then
            MsgBox "该单据已经发药", vbInformation, gstrSysName
            Stuff_RxValied = False
            Exit Function
        End If
        
        '3 当没有勾选参数"允许对未收费的门诊划价处方发料"，检查当前单据是否已经收费
        If rsTemp!单据 = 8 And rsTemp!已收费 = 0 And T_Para.bln允许未收费单据发料 = False Then
            MsgBox "该单据还未收费，不能执行发料操作！", vbInformation, gstrSysName
            Exit Function
        End If
    End If


'4 检查单据的卫材和当前发料部门是否设置了存储库房
    
'5 根据参数"库存检查"，检查当前的库存是否满足当前单据的数量操作
    Set rsTemp = GetMatStock(lng库房ID)
    int库存检查 = GetCheckPara(lng库房ID)
    
    rsData.MoveFirst
    Do While Not rsData.EOF
        rsTemp.Filter = "收费细目id=" & rsData!材料ID
        
        If rsTemp.EOF Then MsgBox rsData!材料名称 & "该材料未设置存储库房", vbInformation, gstrSysName
        
        If intType = 0 Then
            If int库存检查 <> 0 Then
                If Not LocaleStockData(rsData!数量, lng库房ID, rsData!材料ID, rsData!批次, rsData!序号) Then
                    If int库存检查 = 1 Then
                        If MsgBox("当前库存不足是否继续发药！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    ElseIf int库存检查 = 2 Then
                        MsgBox "当前库存不足禁止继续发药！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName
                    End If
                End If
            End If
        End If
        rsData.MoveNext
    Loop
    
'6 检查处方是否已经结帐,病人是否已经出院，再对权限进行相关的检查
    rsData.MoveFirst
    Call Stuff_Check出院病人(strPrivs, int单据, strNo, rsData!记录性质, rsData!门诊标志)
    Call Stuff_Check结帐处方(strPrivs, int单据, strNo, rsData!记录性质, rsData!门诊标志)

    Stuff_RxValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function LocaleStockData(ByVal lng实际数量 As Long, _
    ByVal lng发料部门ID As Long, ByVal lng材料ID As Long, ByVal lng批次 As Long, Optional ByRef lng序号 As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查指定材料指定批的库存是否充足
    '入参:rsStock-指定检查的库存数据(可以为空记录),可以自动扩展
    '     lng发料部门ID-发料部门id
    '     lng材料id-材料id
    '     lng批次-批次
    '
    '出参:lng序号-返回库存的序号
    '返回:成功,表示找到,否则表示未找到
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim dbl库存 As Double
    LocaleStockData = False
    
    err = 0: On Error GoTo ErrHand:
    
    gstrSQL = "" & _
    " Select nvl(F.是否变价,0) 变价,nvl(A.实际数量,0) 数量" & _
    " From 材料特性 B,收费项目目录 F," & _
    "      (Select A.药品id as 材料ID,a.实际数量 From 药品库存 A Where 性质=1 And 库房ID=[1] And 药品ID=[2] And nvl(批次,0)=[3]) A" & _
    " Where B.材料ID=F.ID And B.材料ID=A.材料ID(+) And B.材料ID=[2] "
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "LocaleStockData", lng发料部门ID, lng材料ID, lng批次)
    
    dbl库存 = Val(Nvl(rsTemp!数量))
    
    If lng实际数量 > dbl库存 Then LocaleStockData = False

    LocaleStockData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMatStock(ByVal lng库房ID As Long) As Recordset
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 收费细目id From 收费执行科室 Where 执行科室id = [1] "
    Set GetMatStock = zlDataBase.OpenSQLRecord(gstrSQL, "取存储库房", lng库房ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Stuff_RxRefReturnDetail(ByVal int单据 As Integer, ByVal strNo As String, ByVal lng库房ID As Long, ByVal int记录状态 As Integer) As Recordset
'获取指定已发料单据的明细信息
'函数返回指定的明细记录
    Dim strWhere As String, strWhere1 As String, strTemp As String, strFields As String
    Dim blnHistory As Boolean, strTable As String, strTable1 As String, rsTemp As New ADODB.Recordset
    Dim str门诊 As String
    Dim str病区发料 As String
    
    On Error GoTo errHandle
    
    Select Case T_Para.卫材单位
    Case 0  '散装单位
         strFields = "X.计算单位 单位,1 as 换算系数, "
    Case Else
         strFields = "D.包装单位 单位,D.换算系数,"
    End Select
    
 
    '获取已发料或退料的金额
    strTable = " " & _
    "   Select A.ID, A.NO, A.单据, A.序号, A.药品id, A.费用id, A.批次, A.批号, A.效期, Nvl(A.扣率, 0) 扣率, " & _
    "          Nvl(A.付数, 1) 付数, A.实际数量 实际数量, (A.实际数量 - B.已发数量) 已退数量, B.已发数量 已发数量, " & _
    "          A.记录状态, A.零售价, A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门id, A.库房id, " & _
    "          A.产地, Decode(Nvl(A.领用人, ''), '', '', DECODE(mod(a.记录状态,3),2,'(退)','(领)') || A.领用人) 领料人, H.医嘱序号, " & _
    "          H.序号 As 费用序号,H.开单人 as 开单医生,H.姓名,H.病人id,H.记录性质,H.门诊标志,H.标识号,'' 床号,1 可操作" & _
    "   From 药品收发记录 A, 门诊费用记录 H,病人信息 H1, " & _
    "        (Select A.NO, A.单据, A.药品id, A.序号, Sum(Nvl(A.付数, 1) * A.实际数量) 已发数量 " & _
    "          From 药品收发记录 A " & _
    "          Where A.审核人 Is Not Null And A.库房id + 0 = [3] And A.NO=[2] " & _
    "          Group By A.NO, A.单据, A.药品id, A.序号) B " & _
    "   Where A.NO = B.NO And A.单据 = B.单据 And A.药品id + 0 = B.药品id And A.序号 = B.序号  " & _
    "         And A.审核人 Is Not Null And (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0)  " & _
    "         And A.费用id = H.ID And H.病人ID=H1.病人id(+) "
    

    '清单显示每笔操作过程
    strTable = strTable & strWhere1
    If blnHistory Then
        strTable = AnalyseHistorySQL(strTable, "1 可操作", "-99 可操作")
    End If
    
    strTable1 = " Union All " & _
    "     Select A.ID, A.NO, A.单据, A.序号, A.药品id, A.费用id, A.批次, A.批号, A.效期, Nvl(A.扣率, 0), Nvl(A.付数, 1) 付数, " & _
    "            A.实际数量 实际数量, 0 已退数, 0 已发数量, A.记录状态, A.零售价, A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, " & _
    "            A.审核日期, A.对方部门id, A.库房id, " & _
    "            A.产地, " & _
    "            Decode(Nvl(A.领用人, ''), '', '',Decode(A.记录状态, 2,'(退)', '(领)' )|| A.领用人) 领料人,H.医嘱序号, " & _
    "          H.序号 As 费用序号,H.开单人 as 开单医生,H.姓名,H.病人id,H.记录性质,H.门诊标志,H.标识号,'' 床号, Decode(A.记录状态, 1, 1,Mod(A.记录状态, 3) + 1) 可操作 " & _
    "     From 药品收发记录 A, 门诊费用记录 H ,病人信息 H1" & _
    "     Where A.费用id=H.id And H.病人id=H1.病人ID(+) and A.审核人 Is Not Null And Not (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0) And A.库房id + 0 = [3] "
    
    If blnHistory Then
        '历史数据，不能操作
        strTable1 = AnalyseHistorySQL(strTable1, "Decode(A.记录状态, 1, 1,Mod(A.记录状态, 3) + 1) 可操作", "-99 可操作")
    End If
    
    strTable = strTable & vbCrLf & strTable1
    gstrSQL = " " & _
    "     Select /*+ Rule*/ Distinct S.ID, S.单据, S.药品id 材料id, S.NO, S.序号, S.扣率, P.名称 科室, S.记录性质,S.门诊标志, S.标识号, S.病人id, S.床号, " & _
    "                     S.姓名,M.性别,M.年龄, '[' || X.编码 || ']' || X.名称 材料名称, Nvl(D.在用分批, 0) 分批, X.规格, " & strFields & _
    "                     S.付数 付, S.实际数量/" & IIf(T_Para.卫材单位 = 0, 1, "d.换算系数") & " 数量, S.已退数量/" & IIf(T_Para.卫材单位 = 0, 1, "d.换算系数") & " 已退数量, S.已发数量/" & IIf(T_Para.卫材单位 = 0, 1, "d.换算系数") & " 准退数, " & _
    "                     Decode(S.批号, Null, '', S.批号)  批号, " & _
    "                     Nvl(S.批次, 0) 批次, S.效期, S.零售价*" & IIf(T_Para.卫材单位 = 0, 1, "d.换算系数") & " 单价, S.零售金额 金额, S.单量, S.频次, S.用法, S.摘要 说明, " & _
    "                     To_Char(S.审核日期, 'YYYY-MM-DD HH24:MI:SS') 发料时间, S.审核人, S.审核日期, 可操作, S.医嘱序号, " & _
    "                     I.计算单位, Nvl(S.产地, Nvl(X.产地, '')) 产地, Nvl(M.审查结果, -1) 审查结果, " & _
    "                     Nvl(S.医嘱序号, -1) 医嘱id, S.领料人, '' 库房货位, Z.名称 As 其它名,S.记录状态,s.开单医生 " & _
    "     From (" & strTable & ") S, 部门表 P, 材料特性 D, 收费项目目录 X, " & _
    "          收费项目别名 A, 诊疗项目目录 I, 病人医嘱记录 M, 诊疗项目别名 Z " & _
    "     Where S.药品id = D.材料id And D.材料id = X.ID And S.对方部门id + 0 = P.ID And D.诊疗id = I.ID And " & _
    "           S.医嘱序号 = M.ID(+) And D.诊疗id = Z.诊疗项目id(+) And Z.性质(+) = 2 And D.材料id = A.收费细目id(+) And " & _
    "           A.性质(+) = 3 And  S.单据 =[1] and S.NO=[2] And S.审核人 Is Not Null And S.记录状态=[4] "
    
     '所有
     str门诊 = Replace(gstrSQL, "H.病人病区ID", "H.开单部门ID")
     gstrSQL = Replace(gstrSQL, "'' 床号", "H.床号")
     gstrSQL = str门诊 & " Union All " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")

    gstrSQL = gstrSQL & " Order By NO, 单据, 审核日期"
    
    Set Stuff_RxRefReturnDetail = zlDataBase.OpenSQLRecord(gstrSQL, "Stuff_RxRefReturnDetail", _
        int单据, _
        strNo, _
        lng库房ID, _
        int记录状态)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function AnalyseHistorySQL(ByVal strSQL As String, Optional str原串 As String = "", Optional str现串 As String = "") As String
    '产生历史数据的SQL语句
    Dim strTemp As String
    strTemp = Replace(strSQL, "药品收发记录", "H药品收发记录")
    strTemp = Replace(strTemp, "门诊费用记录", "H门诊费用记录")
    strTemp = Replace(strTemp, "住院费用记录", "H住院费用记录")
    If str原串 <> "" Then
        strTemp = Replace(strTemp, str原串, str现串)
    End If
    strTemp = strSQL & " Union ALL " & strTemp
    AnalyseHistorySQL = strTemp
End Function

Public Function Stuff_RxRefSendDetail(ByVal int单据 As Integer, ByVal strNo As String, ByVal lng库房ID As Long) As Recordset
'获取指定待发料单据的明细信息
'函数返回指定的明细记录
Dim lngRow As Long, strWhere As String, strFields As String
    Dim str门诊 As String
    Dim rsTemp As New ADODB.Recordset
    Dim str病区发料 As String
    Dim str住院 As String
    
    On Error GoTo errHandle

    If T_Para.卫材单位 = 0 Then
        strFields = "x.计算单位 as 单位,1 as 换算系数,"
    Else
        strFields = "d.包装单位 as 单位,d.换算系数,"
    End If
    
    gstrSQL = "" & _
        "      Select Distinct s.Id, s.药品id AS 材料ID, Nvl(n.已收费, 0) 已收费, p.名称 科室, s.配药人 AS 配料人 ,S.费用ID, c.开单人 开单医生, " & _
        "          c.操作员姓名 审核人, s.单据, Nvl(s.扣率, 0) 扣率, s.No, s.序号, nvl(c.病人id,0) as 病人ID, '' 床号, c.姓名,m.性别,m.年龄, " & _
        "          c.标识号, c.操作员姓名, '[' || x.编码 || ']' || x.名称 材料名称, s.付数 付, (s.实际数量/" & IIf(T_Para.卫材单位 = 0, 1, "d.换算系数") & ") 数量, " & _
        "          Nvl(d.在用分批, 0) 分批, x.规格, c.登记时间," & strFields & _
        "          s.零售价*" & IIf(T_Para.卫材单位 = 0, 1, "d.换算系数") & " 单价, s.零售金额 金额, s.单量, s.频次, s.用法, s.摘要 说明, " & _
        "          Decode(s.批号, Null, '', s.批号) || Decode(s.批次, Null, '', 0, '', '(' || s.批次 || ')') 批号, " & _
        "          Nvl(s.批次, 0) 批次, c.医嘱序号, i.计算单位, Nvl(s.产地, Nvl(x.产地, '')) 产地, " & _
        "          Nvl(m.审查结果, -1) 审查结果, Nvl(c.医嘱序号, -1) 医嘱id, '' 库房货位,x.是否变价, m.相关id, " & _
        "          s.对方部门id As 科室id, c.序号 费用序号, C.记录性质,C.门诊标志,0 库存下限, z.名称 As 其它名 " & _
        "       From 未发药品记录 n,药品收发记录 s, 门诊费用记录 c,病人信息 c1, 病人医嘱记录 m,   " & _
        "          部门表 p, 材料特性 d, 收费项目目录 x, 收费项目别名 e,诊疗项目目录 i, 诊疗项目别名 z " & _
        "       Where n.单据 = s.单据 And  n.No = s.No AND nvl(n.库房id,[3])+0=nvl(s.库房id,[3])  " & _
        "             And s.费用id = c.Id AND s.对方部门id + 0 = p.Id  " & _
        "             And s.药品id = d.材料id And S.药品id = x.Id  " & _
        "             And s.药品id = e.收费细目id(+)  And e.性质(+) = 3 " & _
        "             And Nvl(Ltrim(Rtrim(s.摘要)), 'NOT拒发') <> '拒发'  AND s.审核人 Is Null And Nvl(s.发药方式, 0) <> -1 " & _
        "             And Mod(s.记录状态, 3) = 1 And s.单据=[1] " & _
        "             AND d.诊疗ID=i.id  and C.病人ID=c1.病人ID(+) " & _
        "             AND D.诊疗id = z.诊疗项目id(+) And z.性质(+) = 2    " & _
        "             AND c.医嘱序号 = m.Id(+)  And Nvl(c.费用状态,0)<>1 " & _
        "             And Nvl(n.库房id, [3]) + 0 = [3] and S.单据=[1] And S.no=[2] " & _
        "             "
    
    '排除对未发药品的销帐记录
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From 病人费用销帐 X " & _
        " Where X.申请类别 = 0 And X.状态 = 0 And X.收费细目id = S.药品id And X.费用id = S.费用id) "
    
    '所有
    str门诊 = Replace(gstrSQL, "C.病人病区ID", "C.开单部门id")
    gstrSQL = Replace(gstrSQL, "'' 床号", "c.床号")
    str住院 = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    str住院 = Replace(str住院, "And Nvl(c.费用状态,0)<>1", "")
    gstrSQL = str门诊 & " Union All " & str住院

    
    gstrSQL = gstrSQL & "  Order By No, 费用序号"
    
    Set Stuff_RxRefSendDetail = zlDataBase.OpenSQLRecord(gstrSQL, "Stuff_RxRefSendDetail", _
        int单据, _
        strNo, _
        lng库房ID)
        
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Stuff_RxRefReturnNO(ByVal lng库房ID As Long, ByVal strBeginTime As String, ByVal strEndTime As String, ByVal intType As Integer, ByVal strContent As String, ByVal bln显示整个过程 As Boolean, ByVal bln过滤模式 As Boolean, ByVal int服务对象 As Integer) As Recordset
    '获取已发料的单据
    Dim strWhere As String, strWhere1 As String, strTemp As String, strFields As String
    Dim blnHistory As Boolean, strTable As String, strTable1 As String, rsTemp As New ADODB.Recordset
    Dim str门诊 As String
    Dim str病区发料 As String
    Dim str单据 As String
    Dim i As Integer
    Dim str住院 As String
    Dim strGroup As String
    Dim strSQL As String

    On Error GoTo errHandle

    Select Case T_Para.卫材单位
    Case 0  '散装单位
         strFields = "X.计算单位 单位,1 as 换算系数, "
    Case Else
         strFields = "D.包装单位 单位,D.换算系数,"
    End Select


    strWhere1 = ""
    If bln过滤模式 Then
        Select Case intType
            Case mFindType.单据号
                strWhere1 = "  AND A.NO =[4]  "
            Case mFindType.IC卡, mFindType.身份证
                strWhere1 = "  AND H.病人iD=[4]  "
            Case mFindType.住院号
                strWhere1 = "  AND H.标识号=[4] and H.门诊标志=2 "
            Case mFindType.姓名
                strWhere1 = "  AND H.姓名 like [4] "
            Case mFindType.门诊号
                strWhere1 = "  AND H.标识号=[4] and H.门诊标志=1 "
            Case mFindType.医保号
                strWhere1 = "  AND H1.就诊卡号=[4]  "
        End Select
    End If
    
    If T_Para.单据类型 = "" Then
        str单据 = " A.单据 in (24,25,26)"
    Else
        For i = 0 To UBound(Split(T_Para.单据类型, ","))
            If str单据 = "" Then
               str单据 = "(A.单据=" & Split(T_Para.单据类型, ",")(i)
            Else
               str单据 = str单据 & " or A.单据=" & Split(T_Para.单据类型, ",")(i)
            End If
        Next
        str单据 = str单据 & ")"
    End If
    
    If bln显示整个过程 = False Then
        gstrSQL = " SELECT DISTINCT '' As 颜色, A.处方类型,'' As 选择,'0' As 标志,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 24, '收费', 25, '记帐') 类型," & lng库房ID & " 库房id,A.记录状态," & _
                 "      A.单据,1 已收费,A.审核人 配药人,A.NO,H.姓名,sum(A.零售金额) AS 金额," & _
                 "      TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS') 日期,1 可操作,' ' 说明,B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,B.医保号,B.住院号,H.门诊标志, H.记录性质 " & _
                 " FROM " & _
                 "      (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                 "          NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态,A.发药窗口," & _
                 "          A.零售价,B.零售金额 零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID, A.填制人, A.处方类型 " & _
                 "      FROM" & _
                 "          (SELECT A.ID,A.NO,A.单据,A.药品ID,A.序号,A.费用ID,A.批次,A.批号,A.效期,A.付数,A.实际数量,A.记录状态,A.发药窗口,A.零售价,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID, A.填制人, Nvl(A.注册证号, 0) As 处方类型 " & _
                 "          FROM 药品收发记录 A" & _
                 "          WHERE nvl(A.发药方式,-999)<>-1 and A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                 "          AND A.库房ID+0=[1] And A.审核日期 Between [2] And [3] and " & str单据 & _
                 "          ) A," & _
                 "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量,SUM(A.零售金额) 零售金额" & _
                 "          FROM 药品收发记录 A" & _
                 "          WHERE nvl(A.发药方式,-999)<>-1 and A.审核人 IS NOT NULL and " & str单据 & _
                 "          AND A.库房ID+0=[1] And A.审核日期 Between [2] And [3]  " & _
                 "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B"
        gstrSQL = gstrSQL & _
                 "      WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 AND B.已发数量<>0" & _
                 "     ) A,门诊费用记录 H,病人信息 B" & _
                 " WHERE A.库房ID+0=[1] and H.病人id=B.病人id(+)  " & _
                  strWhere1 & _
                 " AND (A.记录状态=1 OR MOD(A.记录状态,3)=0) AND A.审核人 IS NOT NULL AND A.费用ID=H.ID AND A.实际数量<>0 "
    Else
        '清单显示每笔操作过程
         gstrSQL = " SELECT DISTINCT '' As 颜色, A.处方类型,'' As 选择,'0' As 标志,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 24, '收费', 25, '记帐') 类型," & lng库房ID & " 库房id,A.记录状态,A.单据,1 已收费,A.审核人 配药人," & _
                  "      A.NO,H.姓名,sum(A.零售金额) AS 金额,TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS') 日期,A.可操作," & _
                  "      DECODE(A.记录状态,1,'第1次发料',DECODE(MOD(A.记录状态,3),0,'第1次发料',1,'第'||(FLOOR(A.记录状态/3)+1)||'次发料',2,'第'||(FLOOR(A.记录状态/3)+1)||'次退料')) 说明,B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,B.医保号,B.住院号,H.门诊标志, H.记录性质,Zl_Get收费类别(A.单据,A.NO,[1]) As 收费类别 " & _
                  " FROM " & _
                  "      (SELECT * FROM" & _
                  "          (SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                  "              NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态,A.发药窗口," & _
                  "              A.零售价 ,A.零售金额 零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID,1 可操作, A.填制人, A.处方类型 " & _
                  "          FROM" & _
                  "              (SELECT A.ID,A.NO,A.单据,A.药品ID,A.序号,A.费用ID,A.批次,A.批号,A.效期,A.付数,A.实际数量,A.记录状态,A.发药窗口,A.零售价,A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID, A.填制人, Nvl(A.注册证号, 0) As 处方类型 " & _
                  "              FROM 药品收发记录 A" & _
                  "              WHERE nvl(a.发药方式,-999)<>-1 and A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                  "              AND A.库房ID+0=[1] And A.审核日期 Between [2] And [3] and " & str单据 & _
                  "              ) A," & _
                  "              (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                  "              FROM 药品收发记录 A" & _
                  "              WHERE nvl(a.发药方式,-999)<>-1 and A.审核人 IS NOT NULL and " & str单据 & _
                  "              AND A.库房ID+0=[1] And A.审核日期 Between [2] And [3]  " & _
                  "              GROUP BY A.NO,A.单据,A.药品ID,A.序号) B"
         gstrSQL = gstrSQL & _
                  "          WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号)" & _
                  "          UNION" & _
                  "          SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期," & _
                  "          NVL(A.付数,1) 付数,A.实际数量,0 已退数,0 已发数量,A.记录状态,A.发药窗口," & _
                  "          A.零售价 , A.零售金额 零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID," & _
                  "          DECODE(记录状态,1,1,DECODE(MOD(记录状态,3),0,1,MOD(记录状态,3)+1)) 可操作, A.填制人, Nvl(A.注册证号, 0) As 处方类型 " & _
                  "          FROM 药品收发记录 A" & _
                  "          WHERE nvl(a.发药方式,-999)<>-1 and NOT (记录状态=1 OR MOD(记录状态,3)=0) And A.审核日期 Between [2] And [3] and " & str单据
         gstrSQL = gstrSQL & _
                  "     ) A,门诊费用记录 H,病人信息 B" & _
                  " WHERE A.库房ID+0=[1] and H.病人id=B.病人id(+) " & _
                  strWhere1 & _
                  " AND A.审核人 IS NOT NULL AND A.费用ID=H.ID "
    End If
    
    If bln显示整个过程 = False Then
        strGroup = " GROUP BY A.处方类型,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 24, '收费', 25, '记帐'),A.单据,1,A.审核人,A.NO,H.姓名,A.记录状态," & _
            " TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS'),B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,B.医保号,B.住院号,H.门诊标志, H.记录性质 "
    Else
        strGroup = " GROUP BY A.处方类型,Decode(Nvl(h.记录状态, 0),  0,'(未)','') || Decode(a.单据, 24, '收费', 25, '记帐') ,A.单据,1,A.审核人,A.记录状态," & _
            " A.NO,H.姓名,TO_CHAR(A.审核日期,'YYYY-MM-DD HH24:MI:SS'),A.可操作," & _
            " DECODE(A.记录状态,1,'第1次发料',DECODE(MOD(A.记录状态,3),0,'第1次发料',1,'第'||(FLOOR(A.记录状态/3)+1)||'次发料',2,'第'||(FLOOR(A.记录状态/3)+1)||'次退料')),B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.病人ID,B.医保号,B.住院号,H.门诊标志, H.记录性质 "
    End If
    
    '区分门诊、住院
    If int服务对象 = 1 Then
        '门诊划价及门诊记帐
        gstrSQL = gstrSQL & strGroup
    Else
        If int服务对象 = 0 Then
            '门诊及住院所有单据
            str门诊 = gstrSQL
            str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
            
            str门诊 = str门诊 & strGroup
            str住院 = str住院 & strGroup
            
            gstrSQL = str门诊 & " Union All " & str住院
        Else
            '住院记帐
            str住院 = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            str住院 = str住院 & strGroup
            gstrSQL = str住院
        End If
    End If
     
    'order by
    gstrSQL = gstrSQL & " order by 类型,单据,NO "
    
    '判断从开始日期后，是否存在转出的处方数据
    blnHistory = sys.IsMovedByDate(strBeginTime)
    
    '如果存在数据转出，则需要同时从后备表中提取数据
    If blnHistory Then
        strSQL = gstrSQL
        strSQL = Replace(strSQL, "药品收发记录", "H药品收发记录")
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL = Replace(strSQL, "住院费用记录", "H住院费用记录")
        gstrSQL = gstrSQL & " UNION ALL " & strSQL
    End If

    Set Stuff_RxRefReturnNO = zlDataBase.OpenSQLRecord(gstrSQL, "获取退料单据-Stuff_RxRefReturnNO", _
        lng库房ID, _
        CDate(strBeginTime), CDate(strEndTime), _
        strContent)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Stuff_RxRefSendNO(ByVal lng库房ID As Long, ByVal strBeginTime As String, ByVal strEndTime As String, ByVal intType As Integer, ByVal strContent As String, ByVal bln退药待发 As Boolean, ByVal bln过滤模式 As Boolean, ByVal int服务对象 As Integer) As Recordset
'获取待发料的单据
    Dim strSQL As String
    Dim str单据 As String
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim strWhere As String
    Dim str门诊 As String
    Dim str住院 As String
    
    On Error GoTo errHandle
    
    
    strWhere = ""
    If bln过滤模式 Then
        Select Case intType
            Case mFindType.单据号
                strWhere = "  AND A.NO =[4]  "
            Case mFindType.IC卡, mFindType.身份证
                strWhere = "  AND d.病人iD=[4]  "
            Case mFindType.住院号
                strWhere = "  AND d.标识号=[4] and d.门诊标志=2 "
            Case mFindType.姓名
                strWhere = "  AND d.姓名 like [4] "
            Case mFindType.门诊号
                strWhere = "  AND d.标识号=[4] and d.门诊标志=1 "
        End Select
    End If
    
    If T_Para.单据类型 = "" Then
        str单据 = " a.单据 in (24,25,26)"
    Else
        For i = 0 To UBound(Split(T_Para.单据类型, ","))
            If str单据 = "" Then
               str单据 = "(a.单据=" & Split(T_Para.单据类型, ",")(i)
            Else
               str单据 = str单据 & " or a.单据=" & Split(T_Para.单据类型, ",")(i)
            End If
        Next
        str单据 = str单据 & ")"
    End If
    
    
    gstrSQL = "Select /*+ Rule*/" & vbNewLine & _
        " 单据, 已收费, No, 姓名, To_Char(Sum(Round(零售金额, 2)), '999999990.00') As 金额, 日期, 可操作, 说明, 就诊卡号, 门诊号, 身份证号, Ic卡号, 病人id, 医保号, 住院号," & vbNewLine & _
        " Sum(Round(实收金额, 2)) 实收金额, 门诊标志, 记录性质,记录状态," & lng库房ID & " 库房id" & vbNewLine & _
        "From ("

    strSQL = "Select a.单据, a.已收费, a.No, a.姓名, c.零售金额, a.日期, a.可操作, a.说明, a.就诊卡号, a.门诊号, a.身份证号, a.Ic卡号, a.病人id, a.医保号, a.住院号," & vbNewLine & _
        "              d.实收金额, Nvl(a.处方类型, Nvl(c.注册证号, 0)) As 处方类型, d.门诊标志, d.记录性质,c.记录状态, d.收费类别" & vbNewLine & _
        "" & vbNewLine & _
        "       From (Select Distinct b.就诊卡号, b.门诊号, b.身份证号, b.Ic卡号, b.医保号, b.住院号, a.优先级, a.发药窗口, a.填制日期," & vbNewLine & _
        "                              Decode(Nvl(a.已收费, 0), 1, '', '(未)') || Decode(a.单据, 8, '收费', 9, '记帐') 类型, a.单据, a.已收费, a.No," & vbNewLine & _
        "                              a.姓名, To_Char(a.填制日期, 'yyyy-MM-dd hh24:mi:ss') 日期, 1 可操作, ' ' 说明, b.病人id, a.处方类型, a.对方部门id" & vbNewLine & _
        "              From 未发药品记录 a, 病人信息 b" & vbNewLine & _
        "              Where 1 = 1 And (a.库房id =[1] Or a.库房id Is Null) And" & vbNewLine & _
        "                    a.填制日期 Between [2] And [3]" & vbNewLine & _
        "                     And a.病人id = b.病人id(+) And " & str单据 & _
        "                     )a, 药品收发记录 c, 门诊费用记录 d, 部门表 b" & vbNewLine & _
        "       Where c.费用id = d.Id And Nvl(c.发药方式, -999) <> -1 And a.单据 = c.单据 And a.No = c.No And c.审核人 Is Null And d.执行状态 <> 9 And" & vbNewLine & _
        "             (c.库房id = [1] Or c.库房id Is Null) And a.对方部门id = b.Id And " & IIf(bln退药待发, "Mod(c.记录状态, 3) = 1", "c.记录状态=1") & strWhere
        
    If int服务对象 = 0 Then
        '所有
        str门诊 = Replace(strSQL, "C.病人病区ID", "C.开单部门id")
        strSQL = Replace(strSQL, "'' 床号", "c.床号")
        str住院 = Replace(strSQL, "门诊费用记录", "住院费用记录")
        str住院 = Replace(str住院, "And Nvl(c.费用状态,0)<>1", "")
        strSQL = str门诊 & " Union All " & str住院
    ElseIf int服务对象 = 3 Then
        '住院记帐单
        strSQL = Replace(strSQL, "'' 床号", "c.床号")
        strSQL = Replace(strSQL, "门诊费用记录", "住院费用记录")
        strSQL = Replace(strSQL, "And Nvl(c.费用状态,0)<>1", "")
    End If
    
    gstrSQL = gstrSQL & strSQL
    
    gstrSQL = gstrSQL & ") a" & vbNewLine & _
        "Group By a.单据, a .已收费, a.No, a.姓名, a.日期, a.可操作, a.说明, a.就诊卡号, a.门诊号, a.身份证号, a.Ic卡号, a.病人id, a.医保号, a.住院号, a.处方类型," & vbNewLine & _
        "         a.门诊标志, a.记录性质,a.记录状态" & vbNewLine & _
        "Order By a.单据, a.No"

    Set Stuff_RxRefSendNO = zlDataBase.OpenSQLRecord(gstrSQL, "RefreshSendList", lng库房ID, CDate(strBeginTime), CDate(strEndTime), strContent)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Stuff_RxWork(ByVal intType As Integer, ByVal strPrivs As String, ByVal rsTemp As Recordset, ByVal lng库房ID As Long, ByVal int单据 As Integer, ByVal strNo As String, ByVal str退药数量 As String) As Boolean
'对单据进行发退料操作
'函数返回值：true-操作成功，false-操作失败
'intType:0-发料，1-退料
    '数据验证
    If Stuff_RxValied(strPrivs, intType, lng库房ID, int单据, strNo, rsTemp) = False Then
        Stuff_RxWork = False
        Exit Function
    End If
    
    If intType = 0 Then
        '发料处理
        Stuff_RxWork = Stuff_RxSend(strPrivs, rsTemp, lng库房ID, str退药数量)
    Else
        '退料处理
        Stuff_RxWork = Stuff_RxReturn(strPrivs, rsTemp, str退药数量)
    End If
End Function


Public Function Stuff_RxReturn(ByVal strPrivs As String, ByVal rsTemp As Recordset, ByVal str退药数量 As String) As Boolean
'对指定的已发料单据进行退料操作
'函数返回值：true-操作成功，false-操作失败
'参数：strPrivs-权限字符串
'rsTemp-发料操作的数据集
    Dim strDate As String
    Dim arrSQL As Variant
    Dim i As Integer
    Dim blnTrans As Boolean
    Dim dbl退药数量 As Double
    
    On Error GoTo errHandle
    
    arrSQL = Array()
    strDate = sys.Currentdate
    With rsTemp
        If Not rsTemp Is Nothing Then
            If .EOF Then
                Stuff_RxReturn = False
                Exit Function
            End If
            
            Do While Not .EOF
                dbl退药数量 = Val(Mid(Mid(str退药数量, InStr(1, str退药数量, "," & !Id & ",") + 2 + Len(!Id)), 1, InStr(1, Mid(str退药数量, InStr(1, str退药数量, "," & !Id & ",") + 2 + Len(!Id)), "|") - 1))
                'Zl_材料收发记录_部门退料
                gstrSQL = "Zl_材料收发记录_部门退料("
                '    收发id_In   In 药品收发记录.ID%Type,
                gstrSQL = gstrSQL & "" & Nvl(!Id) & ","
                '    审核人_In   In 药品收发记录.审核人%Type,
                gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                '    审核日期_In In 药品收发记录.审核日期%Type,
                gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
                '    批号_In     In 药品库存.上次批号%Type := Null,
                gstrSQL = gstrSQL & "'" & Nvl(!批号) & "',"
                '    效期_In     In 药品库存.效期%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(IsNull(!效期), "NULL", IIf(Nvl(!效期) = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & ","
                '    产地_In     In 药品库存.上次产地%Type := Null,
                gstrSQL = gstrSQL & "'" & Nvl(!产地) & "',"
                '    退料数量_In In 药品收发记录.实际数量%Type := Null,
                gstrSQL = gstrSQL & "" & dbl退药数量 & ","
                '    自动销帐_In Integer := 0,
                gstrSQL = gstrSQL & "0,"
                '    退料人_In   In 药品收发记录.领用人%Type := Null
                gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "')"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
                .MoveNext
            Loop
            
            gcnOracle.BeginTrans
            blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDataBase.ExecuteProcedure(CStr(arrSQL(i)), "单据发料_Stuff_RxReturn")
            Next
            gcnOracle.CommitTrans
            blnTrans = False
        End If
    End With
    Stuff_RxReturn = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Stuff_RxSend(ByVal strPrivs As String, ByVal rsTemp As Recordset, ByVal lng库房ID As Long, ByVal str配料人 As String) As Boolean
'对指定的未发料单据进行发放操作
'函数返回值：true-操作成功，false-操作失败
'参数：strPrivs-权限字符串
'rsTemp-发料操作的数据集
    Dim strDate As String
    Dim strID批次 As String
    
    On Error GoTo errHandle
    
    strDate = sys.Currentdate
    With rsTemp
        If Not rsTemp Is Nothing Then
            If .EOF Then
                Stuff_RxSend = False
                Exit Function
            End If
            
            Do While Not .EOF
                strID批次 = IIf(strID批次 = "", "", strID批次 & "|") & !Id & "," & !批次
                .MoveNext
            Loop
            'Zl_药品收发记录_批量发料
            gstrSQL = "Zl_药品收发记录_批量发料("
            '    收发id_In     In Varchar2, --格式:"id1,批次1|id2,批次2|....."
            gstrSQL = gstrSQL & "'" & strID批次 & "',"
            '    库房id_In     In 药品收发记录.库房id%Type,
            gstrSQL = gstrSQL & "" & lng库房ID & ","
            '    审核人_In     In 药品收发记录.审核人%Type,
            gstrSQL = gstrSQL & "'" & gstrUserName & "',"
            '    审核日期_In   In 药品收发记录.审核日期%Type,
            gstrSQL = gstrSQL & "To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'),"
            '    发料方式_In   In 药品收发记录.发药方式%Type := 3, --1-处方发料;2-批量发料;3-部门发料;-1 停止发料
            gstrSQL = gstrSQL & "1,"
            '    领料人_In     In 药品收发记录.领用人%Type := Null,
            gstrSQL = gstrSQL & "'',"
            '    发料标识号_In In 药品收发记录.汇总发药号%Type := Null,
            gstrSQL = gstrSQL & "Null,"
            '    配料人_In     In 药品收发记录.配药人%Type := Null
            gstrSQL = gstrSQL & "'" & str配料人 & "',"
            '    操作员编码
            gstrSQL = gstrSQL & "'" & UserInfo.编号 & "')"
            
            Call zlDataBase.ExecuteProcedure(gstrSQL, "单据发料_Stuff_RxSend")
        End If
    End With
    Stuff_RxSend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function LoadPerson(ByVal lng人员id As Long) As Recordset
    On Error GoTo errHandle
    gstrSQL = "" & _
        "   Select distinct a.id,a.编号 as 编码,A.姓名 As 名称,简码" & _
        "   From 人员表 A,部门人员 B,部门性质说明 C,人员性质说明 D " & _
        "   Where A.Id=B.人员id And B.部门id=C.部门Id And D.人员id=A.Id " & _
        "       And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) AND B.部门id in (Select 部门ID From 部门人员 where 人员id=[1] ) " & _
        "   ORDER BY 编码 "
    Set LoadPerson = zlDataBase.OpenSQLRecord(gstrSQL, "加载配料人-LoadPerson", lng人员id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlfuncCard_GetPatiID(ByVal lngCardID As Long, ByVal strCardNo As String) As Long
    '一卡通功能：通过卡号取病人ID
    Dim lng病人id As Long
    
    On Error GoTo errHandle
    If Not gobjSquareCard Is Nothing Then
        '通过卡ID和卡号查找病人ID
        gobjSquareCard.zlGetPatiID CStr(lngCardID), strCardNo, False, lng病人id
        
        If lng病人id > 0 Then
            zlfuncCard_GetPatiID = lng病人id
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlfuncCard_SetText(ByVal objTxt As TextBox, ByVal strCardProperty As String)
    '设置输入框属性
    '银行卡类别，格式：短名|全名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
    objTxt.Text = ""
    objTxt.Tag = ""
    objTxt.MaxLength = 0
    
    objTxt.Tag = strCardProperty
    objTxt.MaxLength = Val(Split(strCardProperty, "|")(gCardFormat.卡号长度))
    objTxt.PasswordChar = IIf(Trim(Split(strCardProperty, "|")(gCardFormat.卡号密文)) <> "", "*", "")
End Sub

Public Function Stuff_Check出院病人(ByVal strPrivs As String, ByVal lng单据 As Long, ByVal strNo As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer, Optional ByVal lng病人id As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查出院病人是否允许发料,需要根据权限控制(如果没有权限“发退出院病人处方”，则不允许发退料操作)
    '入参:
    '出参:
    '返回:允许,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------

    '功能说明：如果当前病人是住院病人，
    Dim str姓名 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If lng单据 = 24 Then
        Stuff_Check出院病人 = True
        Exit Function
    End If
    
    '如果未传入病人ID，则自动提取
    If lng病人id = 0 Then
        gstrSQL = "Select 病人ID From 门诊费用记录 Where ID = (Select 费用ID From 药品收发记录 Where 单据=[1] And NO=[2] And Rownum<2)"
        If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        End If
        
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取病人ID", lng单据, strNo)
        lng病人id = rsTemp!病人ID
    End If
    
    '取病人姓名
    gstrSQL = "Select 姓名 From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取病人姓名", lng病人id)
    If rsTemp.EOF Then
        MsgBox "在处方[" & strNo & "]中，病人不存，操作中止！", vbInformation, gstrSysName
        Exit Function
    End If
    str姓名 = rsTemp!姓名
    
    '如果当前病人是住院病人，如果没有权限“发退出院病人处方”，则不允许发退药操作
    If zlStr.IsHavePrivs(strPrivs, "发退出院病人处方") = False Then
        '检查病人已预出院或出院
        gstrSQL = " Select 1 From 病案主页 A,病人信息 B" & _
                  " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.病人ID=[1] " & _
                  " And (Nvl(A.状态,0)=3 Or A.出院日期 Is Not NULL)"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "判断是否已出院", lng病人id)
        
        If rsTemp.RecordCount <> 0 Then
            MsgBox "在处方[" & strNo & "]中，病人“" & str姓名 & "”已出院，你没有对已出院病人的处方进行发料、退料的权限，操作中止！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Stuff_Check出院病人 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Stuff_Check结帐处方(ByVal strPrivs As String, ByVal lng单据 As Long, ByVal strNo As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查处方是否已经结帐了,结帐的处方不能发退料操作
    '入参:  lng单据    ：当前单据类型
    '       strNO      ：当前单据号
    '       lng病人ID  ：仅对多病人单有效
    '       str序号：相关单据序号,以,分离
    '出参:
    '返回:数据合法,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If lng单据 = 24 Then
        Stuff_Check结帐处方 = True
        Exit Function
    End If
    
    '如果没有权限“发退结帐处方”，检查该处方是否已结帐，已结帐处方不允许发退料操作
    If zlStr.IsHavePrivs(strPrivs, "发退结帐处方") = 0 Then
    
        gstrSQL = "Select Nvl(Sum(Nvl(结帐金额,0)),0) AS 结帐金额   " & _
                 "  From 门诊费用记录   " & _
                 "  Where Mod(记录性质,10) = 2 and NO = [1]"
        If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        End If
        gstrSQL = gstrSQL & " Order By 结帐金额 Desc"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "判断是否已结帐", strNo)
        If Nvl(rsTemp!结帐金额, 0) <> 0 Then
            MsgBox "该处方[" & strNo & "]已结帐，你没有对已结帐处方进行发料、退料的权限，操作中止！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Stuff_Check结帐处方 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



