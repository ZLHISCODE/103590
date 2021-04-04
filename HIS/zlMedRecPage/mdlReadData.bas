Attribute VB_Name = "mdlReadData"
Option Explicit

Public Function GetBaseCode(ByVal varInput As Variant) As ADODB.Recordset
'功能：获取基础字典表数据
'参数：varInput=基础字典下拉列表Index,或字典表名
'          strDefault=0-不获取缺省标记值，<>0获取缺省标记值
'返回：过滤后的基础字典记录集

    Dim strTables As String
    Dim strSql As String
    Dim arrTables As Variant
    Dim i As Long
    Dim strFilter As String
    Dim strSort As String
    If gclsPros.BaseCode Is Nothing Then
        '具有编码、名称、简码、缺省标志的表
        '医学警示是多选，实现多选的公共方法,不能支持数据源为记录集的类型，只支持SQL,以后考虑现正记录集的多选
        '因此医学警示不缓存
        If gclsPros.FuncType = f诊断选择 Then
            strTables = "治疗结果;分化程度;最高诊断依据;住院死亡原因"
        ElseIf gclsPros.PatiType = PF_门诊 Then
            strTables = "医疗付款方式;性别;婚姻状况;职业;民族;国籍;血型;学历;身份证未录原因"
        Else
            strTables = "医疗付款方式;性别;婚姻状况;职业;民族;国籍;血型;社会关系;病情;入院方式;分化程度;最高诊断依据;出院方式;感染部位;治疗结果;诊疗麻醉类型;住院死亡原因;学历;出院转入;身份证未录原因;路径上报变异原因"
        End If
        arrTables = Split(strTables, ";")
        For i = LBound(arrTables) To UBound(arrTables)
            strSql = strSql & " Union ALL " & vbNewLine & _
                    "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '" & arrTables(i) & "' 表名 From " & arrTables(i)
        Next
        strSql = Mid(strSql, Len(" Union ALL " & vbNewLine))
        If gclsPros.FuncType <> f诊断选择 Then
            '不规则表，即不完全具有编码、名称、简码、缺省标志的表
            If gclsPros.PatiType <> PF_门诊 Then
                '1、仅具有编码，名称
                strTables = "不良事件;感染因素;手术切口愈合"
                arrTables = Split(strTables, ";")
                For i = LBound(arrTables) To UBound(arrTables)
                    strSql = strSql & " Union ALL " & vbNewLine & _
                            "Select RowNum As ID, 编码,编码 简码, 名称, 0 缺省, '" & arrTables(i) & "' 表名 From " & arrTables(i)
                Next
                
                '2、仅具有编码，简码，名称
                strTables = "临床病例分型;病原学目录;医院感染目录;器械导管目录;ICU类型;抢救病因分类"
                arrTables = Split(strTables, ";")
                For i = LBound(arrTables) To UBound(arrTables)
                    strSql = strSql & " Union ALL " & vbNewLine & _
                            "Select RowNum As ID, 编码, 简码, 名称,0 缺省, '" & arrTables(i) & "' 表名 From " & arrTables(i)
                Next
            End If
            '病案相关表，需要判断共享或者是病案系统本身
            If gclsPros.PatiType = PF_门诊 Then
                strTables = "病人去向"
                arrTables = Split(strTables, ";")
                For i = LBound(arrTables) To UBound(arrTables)
                    strSql = strSql & " Union ALL " & vbNewLine & _
                            "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '" & arrTables(i) & "' 表名 From " & arrTables(i)
                Next
            Else
                strTables = "住院死亡期间"
                arrTables = Split(strTables, ";")
                For i = LBound(arrTables) To UBound(arrTables)
                    strSql = strSql & " Union ALL " & vbNewLine & _
                            "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '" & arrTables(i) & "' 表名 From " & arrTables(i)
                Next
            End If
        Else
            '不规则表，即不完全具有编码、名称、简码、缺省标志的表
            '1、仅具有编码，名称
            strTables = "不良事件;感染因素"
            arrTables = Split(strTables, ";")
            For i = LBound(arrTables) To UBound(arrTables)
                strSql = strSql & " Union ALL " & vbNewLine & _
                        "Select RowNum As ID, 编码,编码 简码, 名称, 0 缺省, '" & arrTables(i) & "' 表名 From " & arrTables(i)
            Next
        End If
        Set gclsPros.BaseCode = zlDatabase.OpenSQLRecord(strSql, "首页读取基础字典")
    End If
    
    gclsPros.BaseCode.Filter = 0
    If gclsPros.BaseCode.RecordCount > 0 Then
        strSort = "编码,ID" '排序自动将游标移动到首行
        If TypeName(varInput) <> "String" Then
            Select Case varInput
                Case BCC_付款方式
                    strFilter = "表名='医疗付款方式'"
                Case BCC_性别
                    strFilter = "表名='性别'"
                Case BCC_婚姻
                    strFilter = "表名='婚姻状况'"
                Case BCC_职业
                    strFilter = "表名='职业'"
                Case BCC_民族
                    strFilter = "表名='民族'"
                Case BCC_国籍
                    strFilter = "表名='国籍'"
                Case BCC_血型
                    strFilter = "表名='血型'"
                Case BCC_关系
                    strFilter = "表名='社会关系'"
                Case BCC_入院情况
                    strFilter = "表名='病情'"
                Case BCC_病例分型
                    strFilter = "表名='临床病例分型'"
                Case BCC_入院途径
                    strFilter = "表名='入院方式'"
                Case BCC_分化程度
                    strFilter = "表名='分化程度'"
                Case BCC_最高诊断依据
                    strFilter = "表名='最高诊断依据'"
                Case BCC_出院方式
                    strFilter = "表名='出院方式'"
                Case BCC_死亡期间
                    strFilter = "表名='住院死亡期间'"
                Case BCC_去向
                    strFilter = "表名='病人去向'"
                Case BCC_文化程度
                    strFilter = "表名='学历'"
                    strSort = "编码 Desc,ID"
                Case BCC_身份证
                    strFilter = "表名='身份证未录原因'"
                Case BCC_变异原因
                    strFilter = "表名='路径上报变异原因'"
            End Select
        Else
            strFilter = "表名='" & varInput & " '"
        End If
        gclsPros.BaseCode.Filter = strFilter
        gclsPros.BaseCode.Sort = strSort '排序自动将游标移动到首行
        Set GetBaseCode = gclsPros.BaseCode
    End If
    Exit Function
errH:
    If ErrCenter() <> 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetManData(Optional ByVal varManPros As Variant, Optional ByVal pfManFrom As PatiFrom = 0, Optional ByVal slSignLevel As SignLevel = -1) As ADODB.Recordset
'功能：获取人员信息
'参数：varManPros=人员下拉列表的索引或人员性质，人员性质：医生,护士,病案编码员
'          pfManFrom=人员来源：0：不计算来源，1-门诊=1，2-住院=1
'          slSignLevel=4个签名的级别代码，分别对应一定的专业技术职务与管理职务
    Dim strManPros As String, strFilter As String
    Dim strSql As String, strSQLTmp As String
    Dim bln门诊医生 As Boolean, blnAdd As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim arrFileds As Variant
    Dim int管理 As Integer, int技术 As Integer
    
    On Error GoTo errH
    
    If TypeName(varManPros) = "String" Then
        strManPros = varManPros
    Else
    '人员下拉列表索引
        pfManFrom = Decode(varManPros, MC_门诊医师, PF_门诊, PF_住院)
        strManPros = Decode(varManPros, MC_责任护士, "护士", MC_质控护士, "护士", MC_编目员, "病案编码员", "医生")
        slSignLevel = Decode(varManPros, MC_科主任, SL_科主任, MC_主任或副主任, SL_主任医师, MC_主治医师, SL_主治医师, MC_住院医师, SL_住院医师, -1)
    End If

    strFilter = strManPros & "=1"
    If pfManFrom <> 0 Then
        strFilter = strFilter & " And " & IIf(pfManFrom = PF_门诊, "门诊=1", "住院=1")
    End If
    If slSignLevel >= SL_科主任 Then
        strFilter = strFilter & IIf(slSignLevel = SL_科主任, " And 管理", " And 技术") & ">=" & Decode(slSignLevel, SL_科主任, 1, SL_主任医师, 4, SL_主治医师, 3, SL_住院医师, 1)
    End If
    '判断是否已经缓存
    If strManPros <> "医生" Then
        blnAdd = InStr(gclsPros.LoadMans, strManPros) = 0
        If blnAdd Then
            gclsPros.LoadMans = gclsPros.LoadMans & "|" & strManPros
        End If
    Else
        If InStr(gclsPros.LoadMans, "医生" & pfManFrom) = 0 Then
            blnAdd = True
            gclsPros.LoadMans = gclsPros.LoadMans & "|" & "医生" & pfManFrom
        End If
    End If
    '没有过滤到数据或没缓存数据，则读取数据库
    If blnAdd Or gclsPros.ManInfo Is Nothing Then
        '组装SQL
        bln门诊医生 = pfManFrom = PF_门诊 And strManPros = "医生"
        If gclsPros.FuncType <> f病案首页 And Not bln门诊医生 And pfManFrom <> 0 Then
            If strManPros <> "护士" Then
                strSQLTmp = "And" & vbNewLine & _
                        "     C.部门id In (Select 部门id From 部门人员 Where 人员id = [2])"
            Else
                '护士不仅是和医生所属科室，而且包括所属科室下的病区
                strSQLTmp = "And" & vbNewLine & _
                        "     C.部门id In (Select 部门id From 部门人员 Where 人员id = [2]" & vbNewLine & _
                                                "Union" & vbNewLine & _
                                                "Select Distinct B.病区ID From 部门人员 a, 病区科室对应 b Where A.部门id = B.科室ID And A.人员id = [2])"
            End If
        End If
        If gclsPros.FuncType = f病案首页 Then
            strSql = "Select Distinct A.ID, A.编号 编码, A.姓名 名称, A.简码,zlwbcode(姓名) 五笔简码, A.管理职务, A.专业技术职务,0 管理,0 技术, 0 医生, 0 护士, 0 病案编码员, 0 门诊,0 住院,0 缺省" & vbNewLine & _
                        "From 人员表 A, 人员性质说明 B" & vbNewLine & _
                        "Where A.Id = B.人员id And B.人员性质 = [1]  And" & vbNewLine & _
                        "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & IIf(gstrNodeNo <> "-", " And (A.站点 = '" & gstrNodeNo & "' Or A.站点 Is Null)", "")
        Else
            strSql = "Select Distinct A.ID, A.编号 编码, A.姓名 名称, A.简码,zlwbcode(姓名) 五笔简码, A.管理职务, A.专业技术职务,0 管理,0 技术, 0 医生, 0 护士, 0 病案编码员, 0 门诊,0 住院,0 缺省" & vbNewLine & _
                        "From 人员表 A, 人员性质说明 B ,部门人员 C, 部门性质说明 D" & vbNewLine & _
                        "Where A.Id = B.人员id And B.人员性质 = [1] And A.Id = C.人员id And C.部门id = D.部门id And D.服务对象  In (" & IIf(pfManFrom = 0, "1,2", pfManFrom) & ",3) And" & vbNewLine & _
                        "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & IIf(gstrNodeNo <> "-", " And (A.站点 = '" & gstrNodeNo & "' Or A.站点 Is Null)", "") & strSQLTmp
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "首页读取人员信息", strManPros, UserInfo.ID)
        If gclsPros.ManInfo Is Nothing Then
            Set gclsPros.ManInfo = zlDatabase.CopyNewRec(rsTmp, True, "ID,编码,名称,简码,五笔简码,管理,技术,医生,护士,病案编码员,门诊,住院,缺省")
        End If
        arrFileds = Array("ID", "编码", "名称", "简码", "五笔简码", "管理", "技术")
        With rsTmp
            .Sort = "Id": gclsPros.ManInfo.Filter = "": gclsPros.ManInfo.Sort = "ID"
            Do While Not .EOF
                gclsPros.ManInfo.Filter = "ID=" & !ID
                If gclsPros.ManInfo.EOF Then
                    int管理 = Decode(!管理职务 & "", "科室主任", 2, "科室副主任", 1, 0)
                    int技术 = Decode(!专业技术职务 & "", "主任医师", 5, "副主任医师", 4, "主治医师", 3, "医师", 2, "医士", 1, 0)
                    gclsPros.ManInfo.AddNew arrFileds, Array(!ID, !编码, !名称, !简码, !五笔简码, int管理, int技术)
                End If
                Select Case strManPros
                    Case "护士"
                        gclsPros.ManInfo!护士 = 1
                    Case "医生"
                        gclsPros.ManInfo!医生 = 1
                    Case "病案编码员"
                        gclsPros.ManInfo!病案编码员 = 1
                End Select
                Select Case pfManFrom
                    Case PF_门诊
                        gclsPros.ManInfo!门诊 = 1
                    Case PF_住院
                        gclsPros.ManInfo!住院 = 1
                End Select
                Call gclsPros.ManInfo.Update
                rsTmp.MoveNext
            Loop
        End With
    End If
    gclsPros.ManInfo.Filter = strFilter
    gclsPros.ManInfo.Sort = "名称" '排序自动将游标移动到首行
    Set GetManData = gclsPros.ManInfo
    Exit Function
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SetCboFromName(ByVal strName As String, objCbo As Object, Optional ByVal strType As String, Optional ByVal blnAdd As Boolean) As Boolean
'功能：将指定姓名的人员加入到下拉框中
'blnAdd=强制增加
    Static rsTmp As ADODB.Recordset
    Dim strSql As String, intIdx As Integer
    
    On Error GoTo errH
    If strType = "人员" Then
        If rsTmp Is Nothing Then
            strSql = "Select A.ID,A.编号 编码,A.姓名 名称,Null As 简码" & _
                " From 人员表 A,人员性质说明 B" & _
                " Where A.ID=B.人员ID And B.人员性质 IN(" & IIf(gclsPros.FuncType = f病案首页, "'医生','护士','病案编码员'", "'医生','护士'") & ")" & _
                " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.姓名"
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSql, "SetCboFromName")
        End If
        
        rsTmp.Filter = "名称='" & strName & "'"
        If Not rsTmp.EOF Then
            intIdx = objCbo.ListCount
            If objCbo.ListCount > 0 Then
                If objCbo.ItemData(objCbo.ListCount - 1) = -1 Then
                    intIdx = objCbo.ListCount - 1
                End If
            End If
            
            If IsNull(rsTmp!编码) Then
                objCbo.AddItem rsTmp!姓名, intIdx
            Else
                objCbo.AddItem rsTmp!编码 & "-" & Chr(13) & rsTmp!名称
            End If
            objCbo.ItemData(objCbo.NewIndex) = Val(rsTmp!ID)
            Call zlControl.CboSetIndex(objCbo.hwnd, objCbo.NewIndex)
        ElseIf gclsPros.FuncType = f病案首页 And blnAdd Then '病案首页直接新增该人员
            objCbo.AddItem strName
            objCbo.ListIndex = objCbo.NewIndex
            objCbo.ItemData(objCbo.NewIndex) = -999
        End If
    ElseIf blnAdd Then
        objCbo.AddItem strName
        objCbo.ListIndex = objCbo.NewIndex
    End If
    
    SetCboFromName = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDeptData() As ADODB.Recordset
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select Distinct A.Id, A.编码, A.名称, A.简码, A.位置,B.工作性质" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where A.Id = B.部门id And (B.服务对象 In (2, 3) And B.工作性质 In ('临床', '手术') OR B.工作性质='ICU') And" & vbNewLine & _
            "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & vbNewLine & _
             IIf(gstrNodeNo <> "-", " And (A.站点 = '" & gstrNodeNo & "' Or A.站点 Is Null)", "") & vbNewLine & _
            "Order By A.编码"

    Set GetDeptData = zlDatabase.OpenSQLRecord(strSql, "获取临床科室")
    Exit Function
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAllerData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal blnMoved As Boolean) As ADODB.Recordset
'功能：获取过敏数据
'参数：intType=0-门诊首页 ，1-住院首页与病案首页
'返回：病人过敏数据
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = " Select Distinct a.Id, a.记录来源, a.过敏时间, a.药物id, a.药物名, a.过敏反应, a.过敏源编码, a.记录时间" & vbNewLine & _
             " From 病人过敏记录 A" & vbNewLine & _
             " Where a.结果 = 1 And a.病人id = [1] And a.主页id = [2] And a.记录来源 " & IIf(gclsPros.FuncType <> f病案首页, " = 3 ", " in (3,4) ") & vbNewLine & _
             " Order By Nvl(a.过敏时间, a.记录时间) Desc, a.药物名"

    If blnMoved Then
        strSql = Replace(strSql, "病人过敏记录", "H病人过敏记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "首页获取过敏信息", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        If gclsPros.FuncType = f病案首页 Then
            rsTmp.Filter = "记录来源=4"
            If rsTmp.EOF Then
                rsTmp.Filter = "记录来源=3"
            End If
        End If
    End If
    Set GetAllerData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiMainInfoData(ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal str挂号单 As String) As ADODB.Recordset
'功能：获取病案主页信息以及病人信息数据
'参数：lng病人ID=病人ID
'      lng主页ID=住院病人才传
'      str挂号单=门诊病人才传
'       blnMove=是否从转出数据中读取，即从历史表中读取
'返回：病案主页信息或病人信息
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '门诊首页获取本次就诊信息以及病人信息,传染病上传优先复诊读取方便判断
    If str挂号单 <> "" Then
        gclsPros.Moved = zlDatabase.NOMoved("病人挂号记录", str挂号单)
        
        strSql = "Select B.Id As 挂号id, A.病人id, A.门诊号, A.医疗付款方式, A.出生日期, A.出生地点, A.身份证号, A.其他证件, A.职业, A.民族, A.国籍, A.籍贯, A.区域, A.学历, A.婚姻状况," & vbNewLine & _
                    "       A.家庭地址, A.家庭电话, A.家庭地址邮编, A.监护人, A.户口地址, A.户口地址邮编, A.合同单位id, A.工作单位 单位地址, A.单位电话, A.单位邮编, Nvl(A.险类, 0) 险类," & vbNewLine & _
                    "       Nvl(B.姓名, A.姓名) 姓名, Nvl(B.性别, A.性别) 性别, Nvl(B.年龄, A.年龄) 年龄, B.发病时间, B.发病地址, B.传染病上传, B.复诊," & vbNewLine & _
                    "       Nvl(Nvl(B.续诊科室id, Decode(B.转诊状态, 1, B.转诊科室id, Null)), B.执行部门id) As 科室id, B.摘要, B.社区, C.社区号" & vbNewLine & _
                    "From 病人信息 A, 病人挂号记录 B, 病人社区信息 C" & vbNewLine & _
                    "Where A.病人id = B.病人id And B.病人id = C.病人id(+) And B.社区 = C.社区(+) And " & IIf(str挂号单 = "NULL", "B.ID", "B.No") & "= [1] And B.记录性质 = 1 And B.记录状态 = 1"
         If gclsPros.Moved Then
            strSql = Replace(strSql, "病人挂号记录", "H病人挂号记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取病人信息以及本次就诊信息", IIf(str挂号单 = "NULL", lng主页ID, str挂号单))
    '病案首页与住院首页获取病人信息以及病案主页信息
    Else
        If gclsPros.FuncType <> f病案首页 Then
            strSql = "Select a.病人id, a.主页id, Nvl(a.姓名, d.姓名) As 姓名, Nvl(a.性别, d.性别) As 性别, Nvl(a.年龄, d.年龄) As 年龄, a.身高, a.体重, a.血型, a.职业, a.国籍," & vbNewLine & _
                            "       a.区域, a.学历, a.婚姻状况, Nvl(a.联系人姓名, d.联系人姓名) 联系人姓名, Nvl(a.联系人关系, d.联系人关系) 联系人关系, Nvl(a.联系人地址, d.联系人地址) 联系人地址," & vbNewLine & _
                            "       Nvl(a.联系人电话, d.联系人电话) 联系人电话, Nvl(a.户口地址, d.户口地址) 户口地址, Nvl(a.户口地址邮编, d.户口地址邮编) 户口地址邮编, Nvl(a.家庭地址, d.家庭地址) 家庭地址," & vbNewLine & _
                            "       Nvl(a.家庭电话, d.家庭电话) 家庭电话, Nvl(a.家庭地址邮编, d.家庭地址邮编) 家庭地址邮编, Nvl(a.单位地址, d.工作单位) 单位地址, Nvl(a.单位电话, d.单位电话) 单位电话," & vbNewLine & _
                            "       Nvl(a.单位邮编, d.单位邮编) 单位邮编, a.住院号, a.病人性质, a.再入院, a.入院病区id, a.入院科室id, a.入院日期, a.入院病况, a.入院方式, a.入院病床, a.出院科室id," & vbNewLine & _
                            "       a.出院病床, a.出院日期, a.出院方式, a.是否确诊, a.确诊日期, a.新发肿瘤, a.抢救次数, a.成功次数, a.随诊标志, a.随诊期限, a.尸检标志, a.门诊医师, a.责任护士, a.住院医师," & vbNewLine & _
                            "       a.编目员编号, a.编目员姓名, a.编目日期, a.费用和, a.中医治疗类别, a.病案号, a.费别, a.医疗付款方式, a.当前病区id, a.险类, a.状态, b.名称 As 入院科室," & vbNewLine & _
                            "       c.名称 As 出院科室, c.编码 As 出院科室编码, d.出生日期, d.出生地点, d.身份, d.民族, d.籍贯, d.Email, d.Qq, d.合同单位id, d.住院次数, d.当前科室id," & vbNewLine & _
                            "       d.入院时间, d.出院时间, d.医保号, d.身份证号, d.其他证件, d.健康号,a.数据转出 " & vbNewLine & _
                            "From 病案主页 a, 部门表 b, 部门表 c, 病人信息 d" & vbNewLine & _
                            "Where A.入院科室id = B.Id(+) And A.出院科室id = C.Id(+) And A.病人id = D.病人id And A.病人id = [1] And A.主页id = [2]"
        Else
            strSql = "Select D.病人id, [2] 主页id, Nvl(A.姓名, D.姓名) As 姓名, Nvl(A.性别, D.性别) As 性别, Nvl(A.年龄, D.年龄) As 年龄, a.病人性质, Nvl(A.职业, D.职业) 职业," & vbNewLine & _
                            "       Nvl(A.国籍, D.国籍) 国籍, Nvl(A.区域, D.区域) 区域, Nvl(A.婚姻状况, D.婚姻状况) 婚姻状况, Nvl(A.家庭地址, D.家庭地址) 家庭地址," & vbNewLine & _
                            "       Nvl(A.家庭电话, D.家庭电话) 家庭电话, Nvl(A.家庭地址邮编, D.家庭地址邮编) 家庭地址邮编, Nvl(A.联系人姓名, D.联系人姓名) 联系人姓名," & vbNewLine & _
                            "       Nvl(A.联系人关系, D.联系人关系) 联系人关系, Nvl(A.联系人地址, D.联系人地址) 联系人地址, Nvl(A.联系人电话, D.联系人电话) 联系人电话, Nvl(A.户口地址, D.户口地址) 户口地址," & vbNewLine & _
                            "       Nvl(A.户口地址邮编, D.户口地址邮编) 户口地址邮编, Nvl(A.单位电话, D.单位电话) 单位电话, Nvl(A.单位邮编, D.单位邮编) 单位邮编, A.再入院, A.入院病区id, A.入院科室id," & vbNewLine & _
                            "       A.入院日期, A.入院病况, A.入院方式, A.入院病床, A.出院科室id, A.出院病床, A.出院日期, A.出院方式, A.是否确诊, A.确诊日期, A.新发肿瘤, A.血型, A.抢救次数," & vbNewLine & _
                            "       A.成功次数, A.随诊标志, A.随诊期限, A.尸检标志, A.门诊医师, A.责任护士, A.住院医师, A.编目员编号, A.编目员姓名, A.编目日期, A.费用和, A.身高, A.体重," & vbNewLine & _
                            "       Nvl(A.单位地址, D.工作单位) 单位地址, A.中医治疗类别, A.状态, B.名称 As 入院科室, C.名称 As 出院科室, C.编码 As 出院科室编码, D.出生日期, D.出生地点, D.身份证号," & vbNewLine & _
                            "       D.其他证件, D.民族, D.籍贯, D.Email, D.Qq, D.合同单位id, D.住院次数, D.当前科室id, D.入院时间, D.出院时间, D.健康号, E.档案号," & vbNewLine & _
                            "       Nvl(E.病案号, A.病案号) As 病案号, Nvl(G.医保号, D.医保号) As 医保号, Nvl(A.住院号, D.住院号) 住院号, A.费别, A.医疗付款方式, A.当前病区id," & vbNewLine & _
                            "       Nvl(A.险类, D.险类) 险类, F.病案号 As 最后病案号, F.档案号 最后档案号, H.编码 最后科室编码 ,a.数据转出 " & vbNewLine & _
                            "From 病案主页 A, 部门表 B, 部门表 C, 病人信息 D, 住院病案记录 E," & vbNewLine & _
                            "     (Select N.病人id, N.病案号, N.档案号, M.出院科室id" & vbNewLine & _
                            "       From 病案主页 M, 住院病案记录 N" & vbNewLine & _
                            "       Where N.病人id = M.病人id(+) And N.主页id = M.主页id(+) And N.病人id = [1] And" & vbNewLine & _
                            "             N.主页id = (Select Max(主页id) As 主页id From 住院病案记录 Where 病人id = [1])) F, 保险帐户 G, 部门表 H" & vbNewLine & _
                            "Where A.入院科室id = B.Id(+) And A.出院科室id = C.Id(+) And A.病人id(+) = D.病人id And A.病人id = E.病人id(+) And A.主页id = E.主页id(+) And" & vbNewLine & _
                            "      D.病人id = F.病人id(+) And A.病人id = G.病人id(+) And A.险类 = G.险类(+) And F.出院科室id = H.Id(+) And D.病人id = [1] And" & vbNewLine & _
                            "      A.主页id(+) = [2]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取病案主页与病人信息", lng病人ID, lng主页ID)
        
        If rsTmp.RecordCount > 0 Then
            gclsPros.Moved = Val(NVL(rsTmp!数据转出)) <> 0
        End If
    End If
    
    Set GetPatiMainInfoData = rsTmp
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetPatiAuxiInfoData(ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal str挂号单 As String, Optional ByVal bytModel As Byte = 1) As ADODB.Recordset
'功能：获取病案主页信息从表或病人信息从表数据
'参数：lng病人ID=病人ID
'      lng主页ID=住院病人才传
'      str挂号单=门诊病人才传
'      bytModel =1 病案首页,=2病人入出管理：新生儿登记
'返回：病案主页信息从表或病人信息从表数据
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strDelicery As String
    On Error GoTo errH
    If bytModel = 1 Then
        If str挂号单 <> "" Then
            strSql = "Select Upper(信息名) 信息名, 信息值,Null 编码" & vbNewLine & _
                    "From 病人信息从表" & vbNewLine & _
                    "Where 病人id = [1] And (就诊id = [2] Or 就诊id Is Null)" & vbNewLine & _
                    "Order By Nvl(就诊id, 999999999)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取病人信息从表", lng病人ID, lng主页ID)
        Else
            strSql = "Select Decode(B.编码,Null, Upper(A.信息名), A.信息名) 信息名, A.信息值, B.编码" & vbNewLine & _
                    "From 病案主页从表 A, 病案项目 B" & vbNewLine & _
                    "Where A.信息名 = B.名称(+) And A.病人id = [1] And A.主页id = [2]" & vbNewLine & _
                    "Order By A.信息名"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取病案主页从表", lng病人ID, lng主页ID)
        End If
    Else
        strDelicery = "分娩时间, 产检次数, 胎次, 胎数, 产程时间1, 产程时间2, 产程时间3,总产程时间,产后出血量,产科并发症,会阴Ⅲ度裂伤"
        strSql = "Select A.信息名, A.信息值, 0 as 类型 " & vbNewLine & _
                "From 病案主页从表 A" & vbNewLine & _
                "Where A.病人id = [1] And A.主页id = [2] And  Instr([3], 信息名) > 0" & vbNewLine & _
                "Order By A.信息名"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取病案主页从表", lng病人ID, lng主页ID, strDelicery)
    End If
    Set GetPatiAuxiInfoData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiDiagData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intType As Integer, Optional ByVal blnLast As Boolean, Optional ByVal bln编目 As Boolean, Optional ByVal blnMoved As Boolean) As ADODB.Recordset
'功能：获取病人诊断信息
'参数：intType=0-门诊首页 ，1-住院首页与病案首页
'      blnLast=True-读取本次就诊的诊断，False=读取最后一次就诊的诊断(该参数只对门诊首页有效）
'返回：病人诊断信息
    Dim strSql As String, strSQLTmp As String, strDiagType As String
    Dim int记录来源 As Integer
    Dim strSQLJudge As String '用来判断住院首页是否有诊断
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '默认直接传入主页ID参数
    strSQLTmp = "[2]"
    If intType = 0 Then
        If blnLast Then
        '最后一次的就诊ID
            strSQLTmp = "(Select Max(ID) As 主页id" & vbNewLine & _
                        "From 病人挂号记录" & vbNewLine & _
                        "Where 病人id = [1] And 记录性质 = 1 And 记录状态 = 1 And" & vbNewLine & _
                        "      登记时间 =" & vbNewLine & _
                        "      (Select Max(A.登记时间)" & vbNewLine & _
                        "       From 病人挂号记录 A" & vbNewLine & _
                        "       Where A.病人id = [1] And A.记录性质 = 1 And A.记录状态 = 1 And A.登记时间 < (Select 登记时间 From 病人挂号记录 Where ID = [2])))"
        End If
        '设置读取诊断的类别以及诊断来源
        If gclsPros.Have中医 Then
            strDiagType = " And A.记录来源 IN(1,3) And A.诊断类型 IN(1,11) "
        Else
            strDiagType = " And A.记录来源 IN(1,3) And A.诊断类型=1 "
        End If
    Else
        '判断是否有首页来源或病案来源的诊断。
        If gclsPros.FuncType <> f病案首页 Then
            int记录来源 = 3
            strSQLJudge = "Select 1 From 病人诊断记录 Where 病人id = [1] And 主页id =[2] And 记录来源 = [3] And Rownum < 2"
            If blnMoved Then
                 strSQLJudge = Replace(strSQLJudge, "病人诊断记录", "H病人诊断记录")
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQLJudge, "首页来源诊断判断", lng病人ID, lng主页ID, int记录来源)
            If rsTmp.RecordCount > 0 Then
                strDiagType = " And A.记录来源 =[3] "
            Else
                strDiagType = " And A.记录来源 IN(1,2,4) "
            End If
        Else
            int记录来源 = 4
            If Not bln编目 Then
                strSQLJudge = "Select Nvl(Max(Nvl(记录来源, 0)), 0) 记录来源" & vbNewLine & _
                                    "From 病人诊断记录" & vbNewLine & _
                                    "Where 病人id = [1] And 主页id = [2] And Nvl(记录来源, 0) <= 4"
                If blnMoved Then
                    strSQLJudge = Replace(strSQLJudge, "病人诊断记录", "H病人诊断记录")
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQLJudge, "首页来源诊断判断", lng病人ID, lng主页ID, int记录来源)
                If rsTmp.RecordCount > 0 Then
                    int记录来源 = Val(rsTmp!记录来源 & "")
                End If
            End If
            strDiagType = " And A.记录来源 =[3] "
        End If
        
        '设置读取诊断的类别
        If gclsPros.Have中医 Then
            strDiagType = strDiagType & " And A.诊断类型 IN(1,2,3,5,6,7,10,11,12,13,21) "
        Else
            strDiagType = strDiagType & " And A.诊断类型 IN(1,2,3,5,6,7,10,21) "
        End If
    End If
    If gclsPros.FuncType <> f病案首页 Then
        '组装SQL,电子病案查阅不用查询医嘱记录
        strSql = "Select A.备注, A.Id, A.病人id, A.主页id, A.医嘱id, A.记录来源, A.诊断次序, Nvl(A.编码序号,1) 编码序号, A.诊断类型, A.入院病情, A.疾病id, A.诊断id, A.证候id,B.名称 疾病名称,C.名称 诊断名称,D.名称 证候名称," & vbNewLine & _
                "       A.诊断描述, A.出院情况, A.是否未治, A.是否疑诊, A.发病时间, B.编码 As 疾病编码,B.类别 As 疾病类别, B.附码, C.编码 As 诊断编码, D.编码 As 证候编码," & vbNewLine & _
                IIf(gclsPros.FuncType = f电子病案, " Null 医嘱id", " (Select F_List2str(Cast(Collect(C.医嘱id || '') As T_Strlist)) 医嘱id" & vbNewLine & _
                "         From 病人诊断医嘱 C,病人医嘱记录 F " & vbNewLine & _
                "         Where C.医嘱ID = F.ID and C.诊断id = A.Id and nvl(F.申请序号,0) = 0) As 医嘱id") & ",B.性别限制, B.疗效限制, B.分娩, B.附码, E.Id As 大类, E.是否病人,Null 附码ID,A.记录日期,A.记录人 " & vbNewLine & _
                "From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C, 疾病编码目录 D,疾病编码分类 E" & vbNewLine & _
                "Where A.疾病id = B.Id(+) And A.诊断id = C.Id(+) And A.证候id = D.Id(+)  And  B.分类id = E.Id(+)" & strDiagType & "And A.取消时间 Is Null And A.诊断描述 Is Not Null And 病人id = [1] And 主页id =" & strSQLTmp & vbNewLine & _
                "Order By A.诊断类型, A.记录来源 Desc, A.诊断次序, Nvl(A.编码序号,1), A.Id"
    Else
        If bln编目 Then
            strSql = "Select A.备注, A.Id, A.病人id, A.主页id, A.医嘱id, A.记录来源, A.诊断次序, Decode(Nvl(A.编码序号, 0), 0, 1, A.编码序号) 编码序号, A.诊断类型, A.入院病情, A.疾病id, A.诊断id, A.证候id," & vbNewLine & _
                    "       B.名称 疾病名称, Null 诊断名称, D.名称 证候名称, A.诊断描述, A.出院情况, A.是否未治, A.是否疑诊, A.发病时间," & vbNewLine & _
                    "       B.编码 As 疾病编码,B.类别 As 疾病类别 ,  Null 诊断编码, D.编码 As 证候编码, Null 医嘱id, B.性别限制, B.疗效限制, B.分娩, B.附码, C.Id As 大类, C.是否病人,NULL 附码ID,A.记录日期,A.记录人 " & vbNewLine & _
                    "From 病人诊断记录 A, 疾病编码目录 B, 疾病编码分类 C, 疾病编码目录 D " & vbNewLine & _
                    "Where A.疾病id = B.Id(+) And A.证候id = D.Id(+) And A.病人id = [1] And A.主页id = [2] " & strDiagType & " And B.分类id = C.Id(+)  " & vbNewLine & _
                    "Order By A.诊断类型, A.诊断次序, Decode(Nvl(A.编码序号, 0), 0, 1, A.编码序号)"
        Else
            strSql = "Select a.备注, a.Id, a.病人id, a.主页id, a.记录来源, Row_Number() Over(Partition By 诊断类型 Order By 诊断次序) As 诊断次序," & vbNewLine & _
                            "       Decode(Nvl(a.编码序号, 0), 0, 1, a.编码序号) 编码序号, a.诊断类型, a.入院病情, a.疾病id, a.诊断id, a.证候id, a.疾病名称, Null 诊断名称, a.证候名称," & vbNewLine & _
                            "       a.诊断描述, a.出院情况, a.是否未治, a.是否疑诊, a.发病时间, a.疾病编码, a.疾病类别, Null 诊断编码, a.证候编码, Null 医嘱id, a.性别限制, a.疗效限制, a.分娩, a.附码," & vbNewLine & _
                            "       a.大类, a.是否病人, a.附码id, a.记录日期, a.记录人" & vbNewLine & _
                            "From (Select Distinct a.Id, a.病人id, a.主页id, a.记录来源, Nvl(a.诊断次序, 1) As 诊断次序, Decode(Nvl(a.编码序号, 0), 0, 1, a.编码序号) 编码序号," & vbNewLine & _
                            "                       a.诊断类型, a.疾病id, a.诊断id, a.证候id, '' || a.诊断描述 As 诊断描述, a.入院病情, a.出院情况, a.是否未治, a.是否疑诊, a.发病时间, a.备注," & vbNewLine & _
                            "                       b.编码 As 疾病编码, b.名称 疾病名称, b.类别 As 疾病类别, b.性别限制, b.疗效限制, b.分娩, b.附码, c.Id As 大类, c.是否病人, d.编码 As 证候编码," & vbNewLine & _
                            "                       d.名称 As 证候名称, NULL 附码id, a.记录日期, a.记录人" & vbNewLine & _
                            "       From 病人诊断记录 a, 疾病编码目录 b, 疾病编码分类 c, 疾病编码目录 d " & vbNewLine & _
                            "       Where a.疾病id = b.Id(+) And a.证候id = d.Id(+) And a.病人id = [1] And a.主页id = [2]  " & strDiagType & " And a.记录来源 = [3] And b.分类id = c.Id(+)  " & vbNewLine & _
                            "       Union All" & vbNewLine & _
                            "       Select Distinct a.Id, a.病人id, a.主页id, a.记录来源, Nvl(a.诊断次序, 1) As 诊断次序, Decode(Nvl(a.编码序号, 0), 0, 1, a.编码序号) 编码序号," & vbNewLine & _
                            "                       a.诊断类型, a.疾病id, a.诊断id, a.证候id, '' || 诊断描述 As 诊断描述, a.入院病情, a.出院情况, 是否未治, 是否疑诊, a.发病时间, a.备注," & vbNewLine & _
                            "                       '' || Null As 疾病编码, '' || Null As 疾病名称, '' || Null As 疾病类别, '' || Null As 性别限制, '' || Null As 疗效限制," & vbNewLine & _
                            "                       '' || Null As 分娩, '' || Null As 附码, 0 * Null As 大类, 0 * Null As 是否病人, '' || Null As 证候编码," & vbNewLine & _
                            "                       '' || Null As 证候名称, 0 * Null 附码id, a.记录日期, a.记录人" & vbNewLine & _
                            "       From 病人诊断记录 a" & vbNewLine & _
                            "       Where a.病人id = [1] And a.主页id = [2] " & strDiagType & " And a.记录来源 = 0 And a.疾病id Is Null And Not Exists" & vbNewLine & _
                            "        (Select 1" & vbNewLine & _
                            "              From 病人诊断记录" & vbNewLine & _
                            "              Where a.病人id = 病人id And a.主页id = 主页id And a.诊断类型 = 诊断类型 And a.诊断次序 = 诊断次序 And 记录来源 = [3] And 疾病id Is Not Null)) a" & vbNewLine & _
                            "Order By a.诊断类型, 诊断次序, a.编码序号"
        End If
    End If
    If blnMoved Then
         strSql = Replace(strSql, "病人诊断记录", "H病人诊断记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取首页诊断", lng病人ID, lng主页ID, int记录来源)
    
    Set GetPatiDiagData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetOPSData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bln编目 As Boolean, Optional ByVal blnMoved As Boolean) As ADODB.Recordset
'功能：获取病人的手麻信息
'返回：病人的手麻信息
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If gclsPros.FuncType <> f病案首页 Then
        strSql = "Select A.Id, A.病人id, A.主页id, A.手术情况, A.记录来源, A.手术日期, A.手术开始时间, A.手术结束时间, Nvl(B.编码, C.编码) As 手术编码, A.已行手术 手术名称," & vbNewLine & _
                "       Nvl(B.名称, C.名称) 手术原名, A.主刀医师, A.助产护士, A.第一助手, A.第二助手, A.麻醉医师, A.准备天数, A.抗菌用药时间, A.抗菌用药天数, A.麻醉开始时间, A.重返目的," & vbNewLine & _
                "       A.切口部位, A.麻醉类型, Decode(A.Asa分级, 'I级', 'P1', 'II级', 'P2', 'III级', 'P3', 'IV级', 'P4', 'V级', 'P5', A.Asa分级) Asa分级, A.Nnis分级, Decode(A.手术级别, 1, '一级手术', 2, '二级手术', 3, '三级手术', 4, '四级手术',9, '无', ' ') As 手术级别, A.切口," & vbNewLine & _
                "       A.愈合, A.再次手术, A.术前抗菌用药, A.非预期的二次手术, A.麻醉并发症, A.术中异物遗留, A.手术并发症, A.术后出血或血肿, A.手术伤口裂开, A.术后深静脉血栓, A.术后生理代谢紊乱," & vbNewLine & _
                "       A.术后呼吸衰竭, A.术后肺栓塞, A.术后败血症, A.术后髋关节骨折, A.重返计划, A.切口感染, A.并发症, A.手术操作id, A.诊疗项目id, A.麻醉方式 麻醉id, D.名称 麻醉方式, A.记录日期," & vbNewLine & _
                "       A.记录人, A.取消时间, A.取消人, Decode(B.手术类型, '甲', '四级手术', '乙', '三级手术', '丙', '二级手术', '丁', '一级手术', '四级', '四级手术', '三级', '三级手术', '二级', '二级手术', '一级', '一级手术', Null) 原手术级别 " & vbNewLine & _
                "From 病人手麻记录 A, 疾病编码目录 B, 诊疗项目目录 C, 诊疗项目目录 D" & vbNewLine & _
                "Where C.Id(+) = A.诊疗项目id And A.手术操作id = B.Id(+) And A.麻醉方式 = D.Id(+) And 病人id = [1] And 主页id = [2] And" & vbNewLine & _
                "      (记录来源 <> 1 Or" & vbNewLine & _
                "       (记录来源 = 1 And 取消时间 Is Null And" & vbNewLine & _
                "       记录日期 =" & vbNewLine & _
                "       (Select Max(记录日期) From 病人手麻记录 Where 病人id = 1 And 主页id = 2 And 记录来源 = 1 And 取消时间 Is Null)))" & vbNewLine & _
                "Order By Nvl(A.手术次序,999),A.ID"
    Else
        If bln编目 Then
            strSql = "Select A.Id, A.手术情况, A.病人id, A.主页id, A.记录来源, A.手术日期, A.手术开始时间, A.手术结束时间, B.编码 As 手术编码, A.已行手术 手术名称, B.名称 手术原名, A.主刀医师," & vbNewLine & _
                    "       A.助产护士, A.第一助手, A.第二助手, A.麻醉医师, A.准备天数, A.抗菌用药时间, A.抗菌用药天数, A.麻醉开始时间, A.重返目的, A.切口部位, A.麻醉类型," & vbNewLine & _
                    "       Decode(A.Asa分级, 'I级', 'P1', 'II级', 'P2', 'III级', 'P3', 'IV级', 'P4', 'V级', 'P5', A.Asa分级) Asa分级, A.Nnis分级," & vbNewLine & _
                    "       Decode(A.手术级别, 1, '一级手术', 2, '二级手术', 3, '三级手术', 4, '四级手术',9, '无', ' ') As 手术级别, A.切口, A.愈合, A.再次手术, A.术前抗菌用药, A.非预期的二次手术," & vbNewLine & _
                    "       A.麻醉并发症, A.术中异物遗留, A.手术并发症, A.术后出血或血肿, A.手术伤口裂开, A.术后深静脉血栓, A.术后生理代谢紊乱, A.术后呼吸衰竭, A.术后肺栓塞, A.术后败血症, A.术后髋关节骨折," & vbNewLine & _
                    "       A.重返计划, A.切口感染, A.并发症, A.手术操作id, A.诊疗项目id, A.麻醉方式 麻醉id,Null 麻醉方式,A.记录日期,A.记录人, A.取消时间, A.取消人, " & vbNewLine & _
                    "      Decode(B.手术类型, '甲', '四级手术', '乙', '三级手术', '丙', '二级手术', '丁', '一级手术', '四级', '四级手术', '三级', '三级手术', '二级', '二级手术', '一级', '一级手术', Null) 原手术级别 " & vbNewLine & _
                    "From 病人手麻记录 A, 疾病编码目录 B" & vbNewLine & _
                    "Where A.手术操作id = B.Id(+) And A.病人id = [1] And A.主页id = [2] And A.记录来源 = 4" & vbNewLine & _
                    "Order By Nvl(A.手术次序,999),A.Id"
        Else
            strSql = "Select A.Id, A.手术情况, A.病人id, A.主页id, A.记录来源,A.手术日期, A.手术开始时间, A.手术结束时间, B.编码 As 手术编码, A.已行手术 手术名称, B.名称 手术原名, A.主刀医师," & vbNewLine & _
                    "       A.助产护士, A.第一助手, A.第二助手, A.麻醉医师, A.准备天数, A.抗菌用药时间, A.抗菌用药天数, A.麻醉开始时间, A.重返目的, A.切口部位, A.麻醉类型," & vbNewLine & _
                    "       Decode(A.Asa分级, 'I级', 'P1', 'II级', 'P2', 'III级', 'P3', 'IV级', 'P4', 'V级', 'P5', A.Asa分级) Asa分级, A.Nnis分级," & vbNewLine & _
                    "       Decode(A.手术级别, 1, '一级手术', 2, '二级手术', 3, '三级手术', 4, '四级手术',9, '无', ' ') As 手术级别, A.切口, A.愈合, A.再次手术, A.术前抗菌用药, A.非预期的二次手术," & vbNewLine & _
                    "       A.麻醉并发症, A.术中异物遗留, A.手术并发症, A.术后出血或血肿, A.手术伤口裂开, A.术后深静脉血栓, A.术后生理代谢紊乱, A.术后呼吸衰竭, A.术后肺栓塞, A.术后败血症, A.术后髋关节骨折," & vbNewLine & _
                    "       A.重返计划, A.切口感染, A.并发症, A.手术操作id, A.诊疗项目id, A.麻醉方式 麻醉id,Null 麻醉方式,A.记录日期,A.记录人, A.取消时间, A.取消人, " & vbNewLine & _
                    "       Decode(B.手术类型, '甲', '四级手术', '乙', '三级手术', '丙', '二级手术', '丁', '一级手术', '四级', '四级手术', '三级', '三级手术', '二级', '二级手术', '一级', '一级手术', Null) 原手术级别 " & vbNewLine & _
                    "From 病人手麻记录 A, 疾病编码目录 B" & vbNewLine & _
                    "Where A.手术操作id = B.Id(+) And 病人id = [1] And 主页id = [2]  " & vbNewLine & _
                    "Order By Nvl(A.手术次序,999),A.Id"
        End If
    End If
    
    If blnMoved Then
         strSql = Replace(strSql, "病人手麻记录", "H病人手麻记录")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取病人手麻信息", lng病人ID, lng主页ID)
    '病案
    If Not bln编目 Then
        rsTmp.Filter = "记录来源=" & IIf(gclsPros.FuncType = f病案首页, 4, 3)
        If rsTmp.EOF Then rsTmp.Filter = "记录来源=" & IIf(gclsPros.FuncType = f病案首页, 3, 1)
        If rsTmp.EOF Then rsTmp.Filter = "记录来源=" & IIf(gclsPros.FuncType = f病案首页, 1, 4)
    End If
    
    Set GetOPSData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDiagMatchData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人的诊断符合情况
'返回：病人的诊断符合情况
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select 符合类型,符合情况 From 诊断符合情况 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取诊断符合情况", lng病人ID, lng主页ID)
    
    Set GetDiagMatchData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDiagExtraID(ByVal strCode As String) As ADODB.Recordset
'功能：获取疾病编码ID
'返回：编码在疾病编码目录里面对应的ID
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    strSql = "Select ID from 疾病编码目录 where 编码 = [1] and RowNum < 2 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取疾病编码ID", strCode)

    Set GetDiagExtraID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetKSSData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人的抗生素使用情况（antibiotic)
'返回：病人的抗生素使用情况
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.药名id, A.用药目的, A.使用阶段, A.使用天数, A.药品名称 名称, 一类切口预防用, Ddd数, 联合用药" & vbNewLine & _
            "From 病人抗生素记录 A" & vbNewLine & _
            "Where A.病人id = [1] And A.主页id = [2]" & vbNewLine & _
            "Order By Ddd数 Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病人抗生素记录", lng病人ID, lng主页ID)
    
    Set GetKSSData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetChemothData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人化疗记录
'返回：病人化疗记录
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.病人id, A.主页id, A.序号, A.疾病id, A.开始日期, A.结束日期, A.疗程数, A.总量, A.化疗方案, A.化疗效果, B.编码 || '-' || B.名称 As 疾病信息" & vbNewLine & _
            "From 病案化疗记录 A, 疾病编码目录 B" & vbNewLine & _
            "Where A.疾病id = B.Id And A.病人id = [1] And A.主页id = [2]" & vbNewLine & _
            "Order By 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病案化疗记录", lng病人ID, lng主页ID)
    
    Set GetChemothData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetRadiothData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人放疗情况
'返回：病人放疗情况
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.病人id, A.主页id, A.序号, A.疾病id, A.开始日期, A.结束日期, A.设野部位, A.放射剂量, A.累计量, A.放疗效果, B.编码 || '-' || B.名称 As 疾病信息" & vbNewLine & _
            "From 病案放疗记录 A, 疾病编码目录 B" & vbNewLine & _
            "Where A.疾病id = B.Id And A.病人id = [1] And A.主页id =[2]" & vbNewLine & _
            "Order By 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病人抗生素记录", lng病人ID, lng主页ID)
    
    Set GetRadiothData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSpiritData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人精神药品使用情况
'返回：病人精神药品使用情况
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select 序号, 药品id, 药物名称, 疗程, 最高日量, 特殊反应, 疗效" & vbNewLine & _
            "From 病案精神治疗" & vbNewLine & _
            "Where 病人id = [1] And 主页id = [2]" & vbNewLine & _
            "Order By 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病人抗生素记录", lng病人ID, lng主页ID)
    
    Set GetSpiritData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetICUData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人重症监护使用情况
'返回：病人重症监护使用情况
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select 序号, 监护室名称, To_Char(进入时间, 'yyyy-mm-dd HH24:mi') As 进入时间, To_Char(退出时间, 'yyyy-mm-dd HH24:mi') As 退出时间 ,人工气道脱出,重返重症医学科,重返间隔时间 ,再入住计划,再入住原因" & vbNewLine & _
            "From 病案重症监护情况" & vbNewLine & _
            "Where 病人id = [1] And 主页id = [2]" & vbNewLine & _
            "Order By 序号"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病案重症监护情况", lng病人ID, lng主页ID)
    
    Set GetICUData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetICUInstrumentsData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人重症监护使用情况
'返回：病人重症监护使用情况
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.序号, A.序号 ||'-' ||A.监护室名称 监护室名称, C.编码||'.'||C.名称 As 器械及导管, To_Char(开始使用时间, 'yyyy-mm-dd HH24:mi') As 开始使用时间," & vbNewLine & _
                "       To_Char(结束使用时间, 'yyyy-mm-dd HH24:mi') As 结束使用时间, 感染累计时间" & vbNewLine & _
                "From 器械导管使用情况 A, 器械导管目录 C" & vbNewLine & _
                "Where A.器械及导管 = C.编码(+) And 病人id = [1] And 主页id = [2]" & vbNewLine & _
                "Order By 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取器械导管使用情况", lng病人ID, lng主页ID)
    
    Set GetICUInstrumentsData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetInfectData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人重症监护使用情况
'返回：病人重症监护使用情况
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.序号,To_Char(A.确诊日期, 'yyyy-mm-dd') As 确诊日期, B.编码 || '.' || A.感染部位 感染部位, A.感染名称 医院感染编码, C.名称 医院感染名称" & vbNewLine & _
                    "From 病人感染记录 A, 感染部位 B, 医院感染目录 C" & vbNewLine & _
                    "Where A.感染部位 = B.名称(+) And A.感染名称 = C.编码(+) And A.病人id = [1] And A.主页id = [2]" & vbNewLine & _
                    "Order By A.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病人感染记录", lng病人ID, lng主页ID)
    
    Set GetInfectData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSampleData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人重症监护使用情况
'返回：病人重症监护使用情况
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.序号,A.标本, A.病原学代码 || '-' || B.名称 As 病原学代码, To_Char(A.送检日期, 'yyyy-mm-dd') As 送检日期" & vbNewLine & _
                    "From 病人病原学检查 A, 病原学目录 B" & vbNewLine & _
                    "Where A.病原学代码 = B.编码(+) And A.病人id = [1] And A.主页id = [2]" & vbNewLine & _
                    "Order By A.序号"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取标本来源", lng病人ID, lng主页ID)
    
    Set GetSampleData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckMergePath(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngDiagType As Long, ByVal lngDiag As Long) As Boolean
'功能：检查临床路径对应的诊断不能修改
'参数：lngDiagType：诊断类型,lngDiag=疾病ID
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If lngDiag = 0 Or lngDiagType = 0 Then CheckMergePath = True: Exit Function
    strSql = " Select 诊断类型,疾病ID From 病人临床路径 Where 病人ID=[1] And 主页ID=[2]" & _
             " Union " & _
             " Select 诊断类型,疾病ID From 病人合并路径 Where 病人ID=[1] And 主页ID=[2]"
             
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gstrSysName, lng病人ID, lng主页ID)
    Do While Not rsTmp.EOF
        If lngDiagType = Val(rsTmp!诊断类型 & "") And lngDiag = Val(rsTmp!疾病id & "") Then
            Exit Function
        End If
        rsTmp.MoveNext
    Loop
    CheckMergePath = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckSign(ByVal int签名场合 As Long, ByVal lng开嘱科室ID As Long, Optional ByVal lng医技科室ID As Long, Optional ByVal lng病人科室ID As Long, _
    Optional ByVal int病人范围 As Integer = 2, Optional ByVal blnCheckCert As Boolean = True) As Boolean
'功能：判断一个部门或是一组部门中是否存在启用了电子签名控制的
'参数：int病人范围=1-门诊,2-住院(缺省)
'     int签名场合:0-门诊医嘱和病历；1-住院医生医嘱和病历；2-住院护士医嘱；3-医技医嘱和报告；4-护理记录和护理病历；5-药品发药；6-LIS;7-PACS;
'     lng开嘱科室ID=如果lng开嘱科室ID=0，则需要根据传入的医技科室，病人科室ID求对应的默认开嘱科室
'                   护士站校对和确认停止时，传入的病区ID，可判断病区是否启用了电子签名
'                   传入-1（抗菌药物审核时，如果判断是否分科室启用）
'     blnCheckCert=true 检查证书是否停用，=false表示不检查
    Dim strSql As String, intTmp As Integer
    Dim rsTmp As Recordset
    
    '如果场合都未启用，则返回false
    If int签名场合 = 0 Or int签名场合 = 1 Then
        intTmp = int签名场合 + 1
    ElseIf int签名场合 > 1 And int签名场合 <= 7 Then
        intTmp = int签名场合
    End If
    If Mid(gstrESign, intTmp, 1) <> "1" Then Exit Function
    If lng开嘱科室ID = 0 And (lng病人科室ID <> 0 Or lng医技科室ID <> 0) Then
        '取开嘱科室
        lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, lng医技科室ID, lng病人科室ID, int病人范围)
        If lng开嘱科室ID = 0 Then Exit Function
    End If
    grsSign.Filter = "部门ID=" & lng开嘱科室ID & " and 场合=" & int签名场合
    If grsSign.RecordCount = 0 Then
        strSql = "Select Zl_Fun_Getsignpar([1],[2]) as 是否启用 From dual"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlAdvice", int签名场合, lng开嘱科室ID)
        If rsTmp.RecordCount > 0 Then
            CheckSign = Val(rsTmp!是否启用 & "") = 1
            grsSign.AddNew
            grsSign!部门ID = lng开嘱科室ID
            grsSign!场合 = int签名场合
            grsSign!是否启用 = Val(rsTmp!是否启用 & "")
        End If
    Else
        grsSign.MoveFirst
        CheckSign = Val(grsSign!是否启用 & "") = 1
    End If
    If CheckSign = True And blnCheckCert Then
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            Err.Clear: On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, gclsPros.SysNo)
            End If
        End If
        '检查证书是否停用
        If gobjESign.CertificateStoped(UserInfo.姓名) Then CheckSign = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get开嘱科室ID(ByVal lng医生ID As Long, ByVal lng医技科室ID As Long, ByVal lng病人科室ID As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal lng执行科室ID As Long) As Long
'功能：由医生确定开嘱科室
'参数：int范围=1-门诊,2-住院(缺省)
'说明：在医生所属科室范围内,优先顺序如下：
'      1、医技科室(医技开嘱)
'      2、病人科室
'      3、服务于门诊/住院病人的某些特殊医嘱的执行科室
'      4、服务于门诊/住院病人的科室且为默认科室
'      5、服务于门诊/住院病人的科室
'      6、默认科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Integer
    Dim arr科室ID(1 To 6) As Long
    
    '开单部门必须是临床或医技
    strSql = "Select Distinct A.部门ID,Nvl(A.缺省,0) as 缺省" & _
        " From 部门人员 A,部门性质说明 B,部门表 C" & _
        " Where A.部门ID=C.ID And A.部门ID=B.部门ID" & _
        " And B.服务对象 IN([2],3) And A.人员ID=[1]" & _
        " And B.工作性质 IN('临床','检查','检验','手术','治疗','营养')" & _
        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng医生ID, int范围)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!部门ID = lng医技科室ID Then
            arr科室ID(1) = rsTmp!部门ID
        ElseIf rsTmp!部门ID = lng病人科室ID Then
            arr科室ID(2) = rsTmp!部门ID
        ElseIf rsTmp!部门ID = lng执行科室ID Then
            arr科室ID(3) = rsTmp!部门ID
        ElseIf rsTmp!缺省 = 1 Then
            arr科室ID(4) = rsTmp!部门ID
        ElseIf arr科室ID(4) = 0 Then
            arr科室ID(5) = rsTmp!部门ID
        End If
        rsTmp.MoveNext
    Next
    arr科室ID(6) = UserInfo.DeptID
    
    For i = LBound(arr科室ID) To UBound(arr科室ID)
        If arr科室ID(i) <> 0 Then
            Get开嘱科室ID = arr科室ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckMecRed(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strfrmCation As String, Optional ByVal strOperateName As String) As Boolean
'功能：检查病案是否已经编目,病案是否在待审查或在审查中(此时首页处于锁定状态，不允许修改)
'       lng病人ID:当前病人ID
'       lng主页ID:当前病人主页ID
'       strfrmCation:调用该函数的窗体名称
'       strOperateName:调用该函数的操作名称。strOperateName为空时，不弹出提示
    Dim strSql As String, rsTmp As Recordset
    Dim int病案状态 As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    '获取病案状态
    strSql = "Select Nvl(病案状态, 0) 病案状态,编目日期 From 病案主页 Where 病人id = [1] And 主页id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, strfrmCation, lng病人ID, lng主页ID)
    rsTmp.MoveFirst
    int病案状态 = rsTmp!病案状态
    '首页锁定与否的判断
    Select Case int病案状态
        Case 1 '等待审查
            strMsg = "该病案等待审查中,不能"
        Case 3 '正在审查
            strMsg = "该病案正在审查中,不能"
        Case 5 '审查归档
            strMsg = "该病案已经审查归档,不能"
        Case 10 '接收待审
            strMsg = "该病案在接收待审中,不能"
        Case Else '2-拒绝审查4-审查反馈;6-审查整改;13-正在抽查;14-抽查反馈;16-抽查整改
            strMsg = ""
    End Select
    
    If strMsg = "" Then
        If Not IsNull(rsTmp!编目日期) Then
            strMsg = "该病人的病案已经编目，不能"
        End If
    End If
    
    If strMsg <> "" Then  '锁定首页
        If strOperateName <> "" Then
            MsgBox strMsg & strOperateName & "！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    CheckMecRed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiPathInfo() As Boolean
'功能：获取病人的临床路径相关信息
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If gclsPros.PathState <> PS_未导入 Then
        '只处理首页中输入的诊断，以前没填的，缺省当作来自于“西医入院诊断”，由于诊断类型列非空，因此取消该条逻辑
        strSql = "Select Nvl(诊断类型, 2) As 诊断类型, Nvl(疾病id, 0) As 疾病id, Nvl(诊断id, 0) As 诊断id, 状态" & vbNewLine & _
                "From 病人临床路径" & vbNewLine & _
                "Where 病人id = [1] And 主页id = [2] And (诊断来源 = 3 Or 诊断来源 Is Null)" & vbNewLine & _
                "Order By 导入时间"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, gclsPros.主页ID)
        If rsTmp.RecordCount > 0 Then
            gclsPros.InPath = rsTmp!诊断类型
            '如果有多条路径，则取第一条的状态
            If rsTmp.RecordCount >= 2 Then gclsPros.PathState = Val(rsTmp!状态 & "")
            rsTmp.MoveNext
            Do While Not rsTmp.EOF
                gclsPros.PathDiag = gclsPros.PathDiag & "," & rsTmp!诊断类型 & "|" & rsTmp!疾病id & "|" & rsTmp!诊断ID
                rsTmp.MoveNext
            Loop
            gclsPros.PathDiag = Mid(gclsPros.PathDiag, 2)
        Else
            gclsPros.InPath = 0
        End If
        '完成路径的时间是否比出院诊断记录时间大()取第一条路径
        If gclsPros.PathState = PS_正常结束 Then
            strSql = "Select Sign(Nvl(A.结束时间, Null) - Nvl(B.记录日期, Sysdate)) As 判断" & vbNewLine & _
                    "From 病人临床路径 A, (Select 病人id, 主页id, 记录日期 From 病人诊断记录 Where 记录来源 = 3 And 诊断次序 = 1 And 诊断类型 = [3]) B" & vbNewLine & _
                    "Where A.病人id = B.病人id(+) And A.主页id = B.主页id(+) And A.病人id = [1] And A.主页id = [2] And" & vbNewLine & _
                    "      A.导入时间 = (Select Min(导入时间) From 病人临床路径 Where 病人id = [1] And 主页id = [2])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, gclsPros.主页ID, IIf(gclsPros.InPath > 10, DT_出院诊断ZY, DT_出院诊断XY))
            If rsTmp.RecordCount > 0 Then
                gclsPros.PathOutTime = Val(rsTmp!判断 & "") = 1
            Else
                gclsPros.PathOutTime = False
            End If
        End If
    End If
    GetPatiPathInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStrucAddress(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strTypeName As String) As ADODB.Recordset
'功能：获得指定类型的结构化地址
'参数：strTypeName=地址类型 出生地点(类型码1),籍贯(类型码2),现住址(类型码3),户口地址(类型码4)
    Dim strSql As String, rsTmp As Recordset
    Dim lngType As Long, blnNew As Boolean
    
    lngType = Decode(strTypeName, "出生地点", 1, "籍贯", 2, "家庭地址", 3, "户口地址", 4, "联系人地址", 5, "单位地址", 6)
    
    blnNew = gclsPros.AdressInfo Is Nothing
    If blnNew Then
        strSql = "Select 病人ID,主页ID,地址类别,省,市,县,乡镇,其他 From 病人地址信息 Where 病人ID=[1] And 主页ID=[2]"
        On Error GoTo errH
        Set gclsPros.AdressInfo = zlDatabase.OpenSQLRecord(strSql, "查询结构化地址", lng病人ID, lng主页ID)
    End If
    
    gclsPros.AdressInfo.Filter = "地址类别=" & lngType
    Set GetStrucAddress = gclsPros.AdressInfo
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetKSSID(ByVal strName As String) As Long
'功能：由于现在将病案主页从表的抗生素 移到了新表 病人抗生素记录中，以前没有记录药品id，现在根据名称将id查出来
'参数：strName=药品名
    Dim rsTmp As Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select Distinct A.Id From 诊疗项目目录 A, 药品特性 C Where A.Id = C.药名id And Nvl(C.抗生素, 0) <> 0 And A.名称 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strName)
    If rsTmp.RecordCount > 0 Then
        GetKSSID = Val(rsTmp!ID)
    Else
        GetKSSID = 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get主页IDByCur(ByVal lng主页ID As Long, Optional ByVal blnNext As Boolean = True) As Long
'功能：根据当前主页ID获取指定的主页ID
'参数：lng主页ID=要进行判断的主页ID
'      blnNext=True-获取比lng主页ID大的最小主页ID,False-获取比lng主页ID小的最大主页ID
'返回：0-不存在这样的主页ID,>0:符合条件的主页ID
    Dim strSql As String, rsTmp As ADODB.Recordset
    If gclsPros.OpenMode = EM_编辑 Or gclsPros.OpenMode = EM_查阅 Then
        If blnNext Then
            strSql = "Select Min(A.主页id) As 主页id" & vbNewLine & _
                            "From 病案主页 A" & vbNewLine & _
                            "Where A.病人id = [1] And Nvl(病人性质, 0) = 0 And 编目日期 Is Not Null And 主页id > [2]"
        Else
           strSql = "Select Max(A.主页id) As 主页id" & vbNewLine & _
                            "From 病案主页 A" & vbNewLine & _
                            "Where A.病人id = [1] And Nvl(病人性质, 0) = 0 And 编目日期 Is Not Null And 主页id < [2]"
        End If
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取主页ID", gclsPros.病人ID, lng主页ID)
        If Not rsTmp.EOF Then
            Get主页IDByCur = IIf(IsNull(rsTmp!主页ID), 0, Val(rsTmp!主页ID & ""))
        Else
            Get主页IDByCur = 0
        End If
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get住院次数Or主页id(ByVal lng病人ID As Long, ByRef lng主页ID As Long, ByRef lng次数 As Long, ByVal bln获取主页id As Boolean, Optional ByVal blnVali主页 As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:获取住院次数或主页ID
    '参数:lng病人id-病人id
    '     lng主页ID-病人的主页ID
    '     lng次数=住院次数(除去留观病人)
    '     bln获取主页id-true表示获取主页id,否则获取住院次数(除去留观病人)
    '     blnVali主页=是否验证主页ID，false,不验证，Ture-验证，此时bln获取主页id 传 False
    '出参:lng次数-返回住院次数或主页id
    '返回:获取的次数成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/5/10
    '-----------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim int标识 As Integer
    
    Err = 0: On Error GoTo Errhand:
    ' Zl_获取住院次数或主页id
    '  病人id_In 病案主页.病人id%Type,
    '  次数_In   病案主页.主页id%Type,
    '  标识_In   Integer:1-返回指定次数的主页id,0-根据主页id返回住院次数(排除了留观病人)
    If Not blnVali主页 Then
        strSql = " Select Zl_获取住院次数或主页id([1],[2],[3]) As 返回 From Dual"
    Else
        strSql = "Select Zl_获取住院次数或主页id(A.病人ID, A.主页ID,[3]) As 返回" & vbNewLine & _
                "From 病案主页 A" & vbNewLine & _
                "Where A.病人id = [1] And A.编目日期 Is Not Null And A.主页id >[2]" & vbNewLine & _
                "Order By A.主页id Desc"
    End If
    If Not blnVali主页 Then int标识 = IIf(bln获取主页id, 1, 0)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取住院次数或主页id", lng病人ID, IIf(bln获取主页id, lng次数, lng主页ID), int标识)
    
    If Not rsTemp.EOF Then
        If Not bln获取主页id Or blnVali主页 Then
            lng次数 = Val(rsTemp!返回 & "")
            If blnVali主页 Then
                MsgBox "请选择第" & lng次数 & "次入院以后的病人信息！", vbInformation, gstrSysName
                Get住院次数Or主页id = False
                Exit Function
            End If
        Else
            lng主页ID = Val(rsTemp!返回 & "")
        End If
    ElseIf Not blnVali主页 Then
        If Not bln获取主页id Or blnVali主页 Then
            lng次数 = 0
        Else
            lng主页ID = 0
        End If
    End If
    
    Get住院次数Or主页id = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Get住院次数Or主页id = False
    If Not bln获取主页id Or blnVali主页 Then
        lng次数 = 0
    Else
        lng主页ID = 0
    End If
End Function

Public Function GetNextNo(ByVal int序号 As Integer, Optional ByVal int判断 As Integer = 0, Optional ByVal strCode As String = "") As Variant
'功能:根据特定规则产生新的号码,本函数规则只适于ZLHIS10，且需要Oracle 8i(8.1.5)以上版本支持
'参数：
'int序号=项目序号:
'  1   病人ID 数字
'  2   住院号 数字
'返回：最大号码
'说明：
'  编号规则：0-按年顺序编号,1-按日顺序编号,2-按执行科室分月编号(需要读取科室号码表)
'            对门诊号：0-顺序编号,1-年月日(YYMMDD)+顺序号(0000)
'            对住院号：0-顺序编号,1-年月(YYMM)+顺序号(0000),2-年(YYYY)+顺序号(00000)
'  年度位确定：以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码
'  最大号码-10存入号码控制表,用于并发情况下补缺号(取了号,但未使用)
'  For Update在并发情况下锁定行,不用Wait选项以避免向调用者返回空
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    GetNextNo = Null
    
    On Error GoTo errH
    '问题 25779 由于调整zl3_NextNO函数,增加int判断 by lesfeng 2009-10-16 b
    If int判断 = 0 Then
        strSql = "Select zl3_NextNO([1],[2],[3]) as NO From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetNextNo", int序号, 0, strCode)
    Else
        strSql = "Select zl3_NextNO([1],[2],[3]) as NO From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetNextNo", int序号, int判断, strCode)
    End If
    '问题 25779 由于调整zl3_NextNO函数,增加int判断 by lesfeng 2009-10-16 b
    If gcnOracle.Errors.Count > 0 Then 'Select中函数出错时,在VB中不自动触发错误
        Err.Raise gcnOracle.Errors(0).Number, , gcnOracle.Errors(0).Description
    End If
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!No) Then GetNextNo = rsTmp!No
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'问题26071
Public Function GetBloodValue(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据血库输血加载输血相关的信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:Lesfeng
    '日期:2009-11-18 12:11:40
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim strCombItem As String
        
    On Error GoTo Errhand
    
    With gclsPros.CurrentForm
        'Zl_Get血库输血信息
        strSql = "select Zl_Get血库输血信息([1],[2]) as Blood from dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, lng病人ID, lng主页ID)
        If Not rsTmp.EOF Then
            strTmp = IIf(IsNull(rsTmp!Blood), "0", rsTmp!Blood)
            If strTmp <> "0" Then
                arrTmp = Split(strTmp, "|")
                strCombItem = ""
                If arrTmp(1) = "未知" Then arrTmp(1) = "不详"
                Set rsTmp = GetBaseCode("血型")
                Do While Not rsTmp.EOF
                    strCombItem = strCombItem & "," & rsTmp!名称
                Loop

                If Trim(arrTmp(1)) <> "" And strCombItem <> "" And InStr(1, strCombItem, Trim(arrTmp(1))) > 0 Then
                    .cboBaseInfo(BCC_血型).Text = Trim(arrTmp(1))
                Else
                    If Trim(arrTmp(1)) <> "" Then
                        .cboBaseInfo(BCC_血型).AddItem arrTmp(1)
                        .cboBaseInfo(BCC_血型).ListIndex = .cboBaseInfo(BCC_血型).NewIndex
                    End If
                End If
                strCombItem = "未查,阴,阳,不详"
                If arrTmp(2) = "未做" Then arrTmp(2) = "未查"
                If Trim(arrTmp(2)) <> "" And InStr(1, strCombItem, Trim(arrTmp(2))) > 0 Then
                    .cboBaseInfo(BCC_RH).Text = Trim(arrTmp(2))
                Else
                    If Trim(arrTmp(2)) <> "" Then
                        .cboBaseInfo(BCC_RH).AddItem arrTmp(2)
                        .cboBaseInfo(BCC_RH).ListIndex = .cboBaseInfo(BCC_RH).NewIndex
                    End If
                End If
                If Trim(arrTmp(3)) <> "0" And IsNumeric(arrTmp(3)) Then
                    .txtSpecificInfo(SLC_输红细胞) = Trim(arrTmp(3))
                End If
                If Trim(arrTmp(4)) <> "0" And IsNumeric(arrTmp(4)) Then
                    .txtSpecificInfo(SLC_输血小板) = Trim(arrTmp(4))
                End If
                If Trim(arrTmp(5)) <> "0" And IsNumeric(arrTmp(5)) Then
                    .txtSpecificInfo(SLC_输血浆) = Trim(arrTmp(5))
                End If
                If Trim(arrTmp(6)) <> "0" And IsNumeric(arrTmp(6)) Then
                    .txtSpecificInfo(SLC_输全血) = Trim(arrTmp(6))
                End If
                If Trim(arrTmp(7)) <> "0" And IsNumeric(arrTmp(7)) Then
                    .txtInfo(GC_输其他) = Trim(arrTmp(7))
                End If
                
                strCombItem = "有,无,未输"
                If Trim(arrTmp(8)) <> "" And InStr(1, strCombItem, Trim(arrTmp(8))) > 0 Then
                    .cboBaseInfo(BCC_输血反应).Text = Trim(arrTmp(8))
                Else
                    If Trim(arrTmp(8)) <> "" Then
                        .cboBaseInfo(BCC_输血反应).AddItem arrTmp(8)
                        .cboBaseInfo(BCC_输血反应).ListIndex = .cboBaseInfo(BCC_输血反应).NewIndex
                    End If
                End If
            End If
        End If
        rsTmp.Close
    End With
    GetBloodValue = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCareValue(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据护理接口加载护理相关的信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘硕
    '日期:2013-12-26 10:06:40
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim strItems As String
    
    On Error GoTo Errhand
    strItems = "特级护理,一级护理,二级护理,三级护理,ICU,CCU"
    'Zl3_Get护理天数
    strSql = "Select Zl3_Get护理天数([1], [2], [3]) As CareData From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng病人ID, lng主页ID, strItems)
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!CareData) Then Exit Function
    strTmp = rsTmp!CareData & ""
    arrTmp = Split(strTmp, "|")
    If UBound(arrTmp) <> UBound(Split(strItems, ",")) Then Exit Function
    With gclsPros.CurrentForm
        If arrTmp(0) <> "" Then
            .txtSpecificInfo(SLC_特护).Text = Format(Val(arrTmp(0)), "###;-###;;")
        End If
        If arrTmp(1) <> "" Then
            .txtSpecificInfo(SLC_一级护理).Text = Format(Val(arrTmp(1)), "###;-###;;")
        End If
        If arrTmp(2) <> "" Then
            .txtSpecificInfo(SLC_二级护理).Text = Format(Val(arrTmp(2)), "###;-###;;")
        End If
        If arrTmp(3) <> "" Then
            .txtSpecificInfo(SLC_三级护理).Text = Format(Val(arrTmp(3)), "###;-###;;")
        End If
        If arrTmp(4) <> "" Then
            .txtSpecificInfo(SLC_ICU).Text = Format(Val(arrTmp(4)), "###;-###;;")
        End If
        If arrTmp(5) <> "" Then
            .txtSpecificInfo(SLC_CCU).Text = Format(Val(arrTmp(5)), "###;-###;;")
        End If
    End With
    GetCareValue = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCareData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取费用数据
    '入参:lng病人ID=病人ID
    '     lng主页ID=病案主页ID
    '出参:
    '返回:返回护理记录集
    '编制:刘硕
    '日期:2013-12-26 10:32:02
    '----------------------------------------------------------------------------------------------
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select b.项目单位 as 单位, b.项目名称 as 信息名, b.记录内容 as 信息值" & _
        " From 病人护理记录 A, 病人护理内容 B Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"
    Set GetCareData = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetFreeData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bln编目 As Boolean) As ADODB.Recordset
'-----------------------------------------------------------------------------------------------------------
'功能:获取费用数据
'入参:lng病人ID=病人ID
'     lng主页ID=病案主页ID
'     bln编目=是否获取编目的数据
'出参:
'返回:返回费用记录集
'编制:刘硕
'日期:2013-12-26 10:32:02
'----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    If bln编目 Then
        strSql = "Select 编码, 名称 费目名称, 上级, 上级 || Decode(Nvl(上级, ''), '', '', '_') || 编码 || '.' || 名称 名称, 末级,金额,0 婴儿费" & vbNewLine & _
                "From (Select A.编码, A.上级,A.末级,A.名称, B.金额" & vbNewLine & _
                "       From 病案费目 A," & vbNewLine & _
                "            (Select 费用名, Sum(金额) 金额" & vbNewLine & _
                "              From 病人费用" & vbNewLine & _
                "              Where 病人id = [1] And 主页id = [2] And Nvl(性质, 0) = 0" & vbNewLine & _
                "              Group By 费用名) B" & vbNewLine & _
                "       Where A.名称 = B.费用名(+))" & vbNewLine & _
                "Start With 上级 Is Null" & vbNewLine & _
                "Connect By Prior 编码 = 上级" & vbNewLine & _
                "Order By 上级 || 编码"
    Else
        strSql = "Select /*+ Rule*/" & vbNewLine & _
                " 编码, 费目名称,上级, 名称, 末级, 金额, 婴儿费" & vbNewLine & _
                "From (Select B.编码, B.名称 费目名称,B.上级, B.上级 || Decode(Nvl(B.上级, ''), '', '', '_') || B.编码 || '.' || B.名称 名称, B.末级," & vbNewLine & _
                "              Sum(Nvl(A.金额, 0)) As 金额, Nvl(A.婴儿费, 0) 婴儿费" & vbNewLine & _
                "       From (Select 编码, 上级, 名称, 末级 From 病案费目 Start With 上级 Is Null Connect By Prior 编码 = 上级) B," & vbNewLine & _
                "            (Select B.编码, A.金额, A.婴儿费" & vbNewLine & _
                "              From (Select B.病案费目, Nvl(A.婴儿费, 0) 婴儿费, Sum(Nvl(A.实收金额, 0)) As 金额" & vbNewLine & _
                "                     From 住院费用记录 A, 收费项目目录 B" & vbNewLine & _
                "                     Where A.收费细目id = B.Id And A.记录状态 <> 0 And A.病人id = [1] And A.主页id = [2]" & vbNewLine & _
                "                     Group By B.病案费目, Nvl(A.婴儿费, 0)) A, 病案费目 B" & vbNewLine & _
                "              Where A.病案费目 = B.名称) A" & vbNewLine & _
                "       Where B.编码 = A.编码(+)" & vbNewLine & _
                "       Group By B.编码, B.名称, B.末级, B.上级, Nvl(A.婴儿费, 0))" & vbNewLine & _
                "Start With 上级 Is Null" & vbNewLine & _
                "Connect By Prior 编码 = 上级" & vbNewLine & _
                "Order By 上级 || 编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng病人ID, lng主页ID)
    Set GetFreeData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBabyInfoData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取病人分娩信息（新生儿信息）
    '入参:lng病人ID=病人ID
    '     lng主页ID=病案主页ID
    '出参:
    '返回:返回新生儿信息
    '编制:刘硕
    '日期:2013-12-27 16:34:02
    '----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select 病人id, 主页id,分娩时间, 胎儿次序, 分娩方式, 出生胎位, 分娩情况, 出生缺陷, 婴儿性别, 婴儿体重, Apgar评分" & vbNewLine & _
            "From 病人分娩信息" & vbNewLine & _
            "Where 病人id = [1] And 主页id = [2]"

    Set GetBabyInfoData = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBabyDiagData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取新生儿诊断记录（新生儿疾病信息）
    '入参:lng病人ID=病人ID
    '     lng主页ID=病案主页ID
    '出参:
    '返回:返回新生儿疾病信息
    '编制:刘硕
    '日期:2013-12-27 16:34:02
    '----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select A.病人id, A.主页id, A.胎儿次序, A.诊断次序, A.疾病id, A.描述信息, B.编码" & vbNewLine & _
            "From 新生儿诊断记录 A, 疾病编码目录 B" & vbNewLine & _
            "Where A.病人id = [1] And A.主页id = [2] And A.疾病id = B.Id" & vbNewLine & _
            "Order By 诊断次序"
            
    Set GetBabyDiagData = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiTransfer(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取病人转科信息
    '入参:lng病人ID=病人ID
    '     lng主页ID=病案主页ID
    '出参:
    '返回:返回病人转科信息
    '编制:刘硕
    '日期:2013-1-2 10:20:11
    '----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = _
                " Select A.科室ID,B.名称 AS 科室名称,A.开始时间" & _
                " From 病人变动记录 A,部门表 B" & _
                " Where A.病人ID=[1] And A.主页ID=[2]" & _
                " And A.科室ID=B.ID And A.开始原因=3 And A.开始时间 is Not NULL" & _
                " Order by A.开始时间"
                   
    Set GetPatiTransfer = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAdvicePause(ByVal lng医嘱ID As Long) As String
'功能：获取指定医嘱的暂停时间段记录
'返回："暂停时间,开始时间;...."
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    strSql = "Select 操作类型,操作时间 From 病人医嘱状态" & _
        " Where 操作类型 IN(6,7) And 医嘱ID=[1]" & _
        " Order by 操作时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISWork", lng医嘱ID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!操作类型 = 6 Then
            strTmp = strTmp & ";" & Format(rsTmp!操作时间, "yyyy-MM-dd HH:mm:ss") & ","
        ElseIf rsTmp!操作类型 = 7 Then
            '启用的那一秒不在暂停的范围之内
            strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!操作时间), "yyyy-MM-dd HH:mm:ss")
        End If
        rsTmp.MoveNext
    Next
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AutoGetOPSInfo(ByVal bln手麻 As Boolean, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取手麻系统或医嘱中的手术信息，首要从手麻中读取，若手麻读取不到或者没有安装手麻，则读取医嘱
'参数：bln手麻 是否读取手麻系统
'      lng病人ID 病人ID
'      lng主页ID 主页ID
'返回：手术信息记录集
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, rsOther As ADODB.Recordset
    Dim rsReturn As ADODB.Recordset
    Dim strTmp As String
    Dim lng手麻手术id As Long
    Dim blnDefault As Boolean
    
    Dim blnReadAdvice As Boolean
    
    blnReadAdvice = Not bln手麻
    
    On Error GoTo errH
    '为了读取表结构，因此不查询数据
    strSql = "Select a.ID,a.助产护士, a.切口, a.愈合, b.编码 手术编码, a.手术日期, a.手术开始时间, a.手术结束时间, a.拟行手术, a.手术操作id, a.诊疗项目id, a.已行手术, a.主刀医师," & vbNewLine & _
            "       a.第一助手, a.第二助手, a.手术护士, a.麻醉开始时间, a.麻醉结束时间,C.名称 手术原名 , C.名称 麻醉方式 , A.麻醉方式 麻醉id,  a.麻醉类型, a.麻醉质量, a.输液总量, a.麻醉医师, a.输氧开始时间, a.输氧结束时间, a.手术情况, a.ASA分级," & vbNewLine & _
            "       a.再次手术, a.NNIS分级, '一级手术' 手术级别, a.术前抗菌用药, a.抗菌用药天数, a.非预期的二次手术, a.麻醉并发症, a.术中异物遗留, a.手术并发症, a.术后出血或血肿, a.手术伤口裂开," & vbNewLine & _
            "       a.术后深静脉血栓, a.术后生理代谢紊乱, a.术后呼吸衰竭, a.术后肺栓塞, a.术后败血症, a.术后髋关节骨折, a.准备天数, a.抗菌用药时间, a.切口部位, a.重返计划, a.重返目的, a.切口感染," & vbNewLine & _
            "       a.并发症" & vbNewLine & _
            "From 病人手麻记录 A, 疾病编码目录 B, 诊疗项目目录 C" & vbNewLine & _
            "Where c.Id = a.诊疗项目id And a.手术操作id = b.Id And 病人id = 0 And 主页id = 0 And 记录来源 = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息")
    Set rsReturn = zlDatabase.CopyNewRec(rsTmp, True)
    With rsReturn
        If bln手麻 Then
            strSql = "Select a.Id, a.手术时间 手术日期, a.开始时间 手术开始时间, a.结束时间 手术结束时间, a.麻醉开始 麻醉开始时间, a.麻醉结束 麻醉结束时间, a.麻醉质量 麻醉质量," & vbNewLine & _
                    "       a.输氧开始 输氧开始时间, a.输氧结束 输氧结束时间, a.手术规模 手术级别" & vbNewLine & _
                    "From 病人手麻主页 A" & vbNewLine & _
                    "Where 病人id = [1] And 主页id = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", lng病人ID, lng主页ID)
            If rsTmp.EOF Then
                blnReadAdvice = True
            Else
                While Not rsTmp.EOF
                    blnDefault = True
                    .AddNew
                    !ID = rsTmp!ID
                    !手术日期 = rsTmp!手术日期
                    !手术开始时间 = rsTmp!手术开始时间
                    !手术结束时间 = rsTmp!手术结束时间
                    !麻醉开始时间 = rsTmp!麻醉开始时间
                    !麻醉结束时间 = rsTmp!麻醉结束时间
                    !麻醉质量 = rsTmp!麻醉质量
                    !输氧开始时间 = rsTmp!输氧开始时间
                    !输氧结束时间 = rsTmp!输氧结束时间
                    '麻醉类型，麻醉方式
                    strSql = "Select b.名称 As 麻醉方式, b.Id, b.操作类型 麻醉类型, a.主要麻醉" & vbNewLine & _
                            "From 病人手麻麻醉 A, 诊疗项目目录 B" & vbNewLine & _
                            "Where a.诊疗项目id = b.Id And a.手麻主页id = [1]" & vbNewLine & _
                            "Order By a.序号"
                    Set rsOther = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", Val(rsTmp!ID & ""))
                    If Not rsOther.EOF Then
                        !麻醉方式 = rsOther!麻醉方式
                        !麻醉类型 = rsOther!麻醉类型
                        !麻醉ID = Val(rsOther!ID & "")
                        rsOther.Filter = "主要麻醉=1"
                        If Not rsOther.EOF Then !麻醉方式 = rsOther!麻醉方式: !麻醉类型 = rsOther!麻醉类型
                    End If
                    
                    lng手麻手术id = 0
                    '已行手术
                    strSql = "Select a.Id 手麻手术id, a.记录性质, A.手术名称, b.名称 手术原名, Nvl(D.编码,B.编码) 手术编码 , a.诊疗项目id, d.Id 手术操作id, a.主手术, a.切口分类 切口, a.愈合情况 愈合," & vbNewLine & _
                        "Decode(d.手术类型, '甲', '四级手术', '乙', '三级手术', '丙', '二级手术', '丁', '一级手术', '四级', '四级手术', '三级', '三级手术', '二级', '二级手术', '一级', '一级手术', Null) 手术级别" & vbNewLine & _
                        "From 病人手麻手术 A, 诊疗项目目录 B, 疾病诊断对照 C, 疾病编码目录 D" & vbNewLine & _
                        "Where a.诊疗项目id = b.Id And b.Id = c.手术id(+) And c.疾病id = d.Id(+) And a.记录性质 = 2 And a.手麻主页id = [1]" & vbNewLine & _
                        "Order By 手麻手术id"
                    Set rsOther = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", Val(rsTmp!ID & ""))
                    If Not rsOther.EOF Then
                        !已行手术 = rsOther!手术名称
                        !手术原名 = rsOther!手术原名
                        lng手麻手术id = Val(rsOther!手麻手术id & "")
                        !手术编码 = rsOther!手术编码
                        rsOther.Filter = "主手术=1"
                        If Not rsOther.EOF Then
                            !已行手术 = rsOther!手术名称
                            !手术原名 = rsOther!手术原名
                            lng手麻手术id = Val(rsOther!手麻手术id & "")
                            !手术编码 = rsOther!手术编码
                        End If
                        rsOther.Filter = "手麻手术id=" & lng手麻手术id
                        !手术操作ID = rsOther!手术操作ID
                        !切口 = rsOther!切口
                        !愈合 = rsOther!愈合
                        !诊疗项目id = rsOther!诊疗项目id
                        !手术级别 = rsOther!手术级别
                        
                    End If
                    '主刀医生等人员的读取
                    If lng手麻手术id <> 0 Then
                        strSql = "Select Distinct 岗位, 姓名, B.是否唯一, B.手术医生, B.麻醉医生, B.手术护士" & vbNewLine & _
                                "From 病人手麻分布 A, 手麻岗位 B" & vbNewLine & _
                                "Where a.岗位 = b.名称 And 手麻手术id = [1]"
                        Set rsOther = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", lng手麻手术id)
                        If rsOther.EOF Then
                            strSql = "Select Distinct 岗位, 姓名, B.是否唯一, B.手术医生, B.麻醉医生, B.手术护士" & vbNewLine & _
                                    "From 病人手麻人员 A, 手麻岗位 B" & vbNewLine & _
                                    "Where A.岗位 = B.名称 And 手麻主页id = [1]"
                            Set rsOther = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", Val(rsTmp!ID & ""))
                        End If
                        If Not rsOther.EOF Then
                            rsOther.Filter = " 是否唯一=1  And 手术医生=1 "
                            If Not rsOther.EOF Then !主刀医师 = rsOther!姓名
                            
                            rsOther.Filter = " 麻醉医生=1 And 岗位 like '主麻%'"
                            If Not rsOther.EOF Then !麻醉医师 = rsOther!姓名
                            If Len(!麻醉医师 & "") = 0 Then
                                rsOther.Filter = "麻醉医生=1"
                                If Not rsOther.EOF Then !麻醉医师 = rsOther!姓名
                            End If
                            rsOther.Filter = "岗位='第一助手' OR 岗位='第1助手' OR 岗位='第Ⅰ助手' OR 岗位='助手医生一' OR 岗位='助手医生1' OR 岗位='助手医生Ⅰ' OR 岗位='助手医师一' OR 岗位='助手医师1' OR 岗位='助手医师Ⅰ' "
                            If Not rsOther.EOF Then
                                blnDefault = False '获取到第一助手，则不默认读取
                                rsOther.Sort = "岗位,姓名"
                                !第一助手 = rsOther!姓名
                            End If
                            rsOther.Filter = "岗位='第二助手' OR 岗位='第2助手' OR 岗位='第Ⅱ助手' OR 岗位='助手医生二' OR 岗位='助手医生2' OR 岗位='助手医生Ⅱ' OR 岗位='助手医师二' OR 岗位='助手医师2' OR 岗位='助手医师Ⅱ'"
                            If Not rsOther.EOF Then
                                blnDefault = False '获取到第二助手，则不默认读取
                                rsOther.Sort = "岗位,姓名"
                                !第二助手 = rsOther!姓名
                            End If
                            If blnDefault Then
                                rsOther.Filter = " 是否唯一=0  And 手术医生=1  "
                                If Not rsOther.EOF Then
                                    rsOther.Sort = "岗位,姓名"
                                    !第一助手 = rsOther!姓名
                                    If rsOther.RecordCount <> 1 Then
                                        rsOther.MoveNext
                                        !第二助手 = rsOther!姓名
                                    End If
                                End If
                            End If
                        End If
                    End If
                    'ASA分级，NNIS分级
                    strSql = "Select Upper(标题内容) 项目, 内容文本" & vbNewLine & _
                            "From 手麻要素应用 A, 病人手麻附项 B ,病人手麻事件 C" & vbNewLine & _
                            "Where a.诊治项目id = b.诊治项目id And b.手麻事件id=c.ID And b.手麻主页id = [1]"
                    Set rsOther = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", Val(!ID & ""))
                    If Not rsOther.EOF Then
                        strTmp = ""
                        rsOther.Filter = " 项目='ASA分级' "
                        If Not rsOther.EOF Then strTmp = rsOther!内容文本 & ""
                        If Len(strTmp) <> 0 Then strTmp = MidB(strTmp, 1, 20)
                        !asa分级 = Decode(Trim(strTmp), "I级", "P1", "II级", "P2", "III级", "P3", "IV级", "P4", "V级", "P5", strTmp)
                        
                        strTmp = ""
                        rsOther.Filter = " 项目='NNIS分级' "
                        If Not rsOther.EOF Then strTmp = rsOther!内容文本 & ""
                        If Len(strTmp) <> 0 Then strTmp = MidB(strTmp, 1, 20)
                        !NNIS分级 = strTmp
                    End If
                    .Update
                    rsTmp.MoveNext
                Wend
            End If
        End If
        
        If blnReadAdvice Then
            strSql = "Select a.Id, NVL(Trunc(E.安排时间),NVL(Trunc(a.手术时间),Trunc(a.开始执行时间))) 手术日期, Nvl(D.编码,B.编码) 手术编码 , NVL(E.安排时间,NVL(A.手术时间,A.开始执行时间)) 手术开始时间, NVL(E.完成时间,a.停嘱时间) 手术结束时间, a.诊疗项目id, d.Id 手术操作id, b.名称 手术名称," & vbNewLine & _
                "Decode(d.手术类型, '甲', '四级手术', '乙', '三级手术', '丙', '二级手术', '丁', '一级手术', '四级', '四级手术', '三级', '三级手术', '二级', '二级手术', '一级', '一级手术',Null) 手术级别" & vbNewLine & _
                "From 病人医嘱记录 A, 诊疗项目目录 B, 疾病诊断对照 C, 疾病编码目录 D,病人医嘱发送 E" & vbNewLine & _
                "Where a.诊疗项目id = b.Id And A.id = e.医嘱id And b.Id = c.手术id(+) And c.疾病id = d.Id(+) And a.诊疗类别 = 'F' And a.病人id = [1] And 主页id = [2] And" & vbNewLine & _
                "医嘱状态 = 8"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", lng病人ID, lng主页ID)
            While Not rsTmp.EOF
                blnDefault = True
                .AddNew
                !ID = rsTmp!ID
                !手术日期 = rsTmp!手术日期
                !手术编码 = rsTmp!手术编码
                !手术开始时间 = rsTmp!手术开始时间
                !手术结束时间 = rsTmp!手术结束时间
                !手术操作ID = rsTmp!手术操作ID
                !已行手术 = rsTmp!手术名称
                !手术级别 = rsTmp!手术级别
                '麻醉信息读取
                strSql = "Select a.开始执行时间 麻醉开始时间, a.停嘱时间 麻醉结束时间, b.名称 As 麻醉方式, b.Id, b.操作类型 麻醉类型" & vbNewLine & _
                        "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
                        "Where a.诊疗项目id = b.Id And a.诊疗类别 = 'G' And a.相关id = [1]"
                Set rsOther = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", Val(rsTmp!ID & ""))
                If Not rsOther.EOF Then
                    !麻醉开始时间 = rsOther!麻醉开始时间
                    !麻醉结束时间 = rsOther!麻醉结束时间
                    !麻醉方式 = rsOther!麻醉方式
                    !麻醉类型 = rsOther!麻醉类型
                End If
                '主刀医生，助手医生读取
                strSql = "Select 项目, 内容 From 病人医嘱附件 Where 医嘱id = [1]"
                Set rsOther = zlDatabase.OpenSQLRecord(strSql, "读取手术操作信息", Val(rsTmp!ID & ""))
                If Not rsOther.EOF Then
                
                    strTmp = ""
                    rsOther.Filter = "项目='主刀医生' OR 项目='主刀医师'"
                    If Not rsOther.EOF Then strTmp = rsOther!内容 & ""
                    If Len(strTmp) <> 0 Then !主刀医师 = MidB(strTmp, 1, 20)
                    rsOther.Filter = "项目='第一助手' OR 项目='第1助手' OR 项目='第Ⅰ助手' OR 项目='助手医生一' OR 项目='助手医生1' OR 项目='助手医生Ⅰ' OR 项目='助手医师一' OR 项目='助手医师1' OR 项目='助手医师Ⅰ' "
                    If Not rsOther.EOF Then
                        blnDefault = False '获取到第一助手，则不默认读取
                        rsOther.Sort = "项目,内容"
                        !第一助手 = MidB(rsOther!内容 & "", 1, 20)
                    End If
                    rsOther.Filter = "项目='第二助手' OR 项目='第2助手' OR 项目='第Ⅱ助手' OR 项目='助手医生二' OR 项目='助手医生2' OR 项目='助手医生Ⅱ' OR 项目='助手医师二' OR 项目='助手医师2' OR 项目='助手医师Ⅱ'"
                    If Not rsOther.EOF Then
                        blnDefault = False '获取到第二助手，则不默认读取
                        rsOther.Sort = "项目,内容"
                        !第二助手 = MidB(rsOther!内容 & "", 1, 20)
                    End If
                    If blnDefault Then
                        rsOther.Filter = "项目 Like '助手*'"
                        If Not rsOther.EOF Then
                            rsOther.Sort = "项目,内容"
                            !第一助手 = MidB(rsOther!内容 & "", 1, 20)
                            rsOther.MoveNext
                            If Not rsOther.EOF Then
                                !第二助手 = MidB(rsOther!内容 & "", 1, 20)
                            End If
                        End If
                    End If
                End If
                .Update
                rsTmp.MoveNext
            Wend
        End If
    End With
    
    Set AutoGetOPSInfo = rsReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetNameByCode(ByVal str信息名 As String, ByVal str信息值 As String) As String
'功能：根据信息值与信息名获取名称

    
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    GetNameByCode = str信息值
    
    Select Case str信息名
        Case "病例分型"
            strSql = "Select 名称 From 临床病例分型 where 编码=[1]"
    End Select
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病案审查查阅", str信息值)
    If rsTmp.RecordCount <> 0 Then
        GetNameByCode = rsTmp.Fields(0).Value
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean, Optional ByVal lngSys As Long) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
'      lngSys=指定系统的内部模块权限，传0或不传是默认是当前系统
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    If lngSys = 0 Then lngSys = gclsPros.SysNo
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear
        blnLoad = True
    End If
    On Error GoTo 0
    If blnLoad Then
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    strPrivs = GetPrivFunc(lngSys, lngProg)
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function IsPageNosCodeRule(ByVal ctCode As Code_Type) As Boolean
'功能: 检查档案号是否根据科室编码编号或者检查病案号是否是顺序编号
'参数：intType=4-检查病案号是否是顺序编号，5-检查档案号是否根据科室编码编号
'53638:刘鹏飞,2013-05-10,档案号编码规则
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTrue As Boolean
    On Error GoTo Errhand
    
    strSql = " Select 编号规则 From 号码控制表 Where 项目序号 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "号码控制表", ctCode)
    If Not rsTmp.EOF Then
        If Val(rsTmp!编号规则 & "") = IIf(ctCode = CT_档案号, 3, 0) Then
            blnTrue = True
        End If
    End If
    
    IsPageNosCodeRule = blnTrue
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ValidatePageNos(Optional ByVal blnSave As Boolean) As Boolean
'功能: 验证病案首页编辑时的病案号，档案号，等是否有效
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, rsPati As ADODB.Recordset
    Dim strTmp As String, strNo As String, strOutDeptCode As String, strTmpNo As String
    Dim bln顺序 As Boolean, blnDo As Boolean
    Dim strFilter As String
    Dim lngCount As Long
    Dim blnSamePageNo As Boolean

    On Error GoTo Errhand
    '病人ID以及住院号的重复检查，以前在界面数据检查中，即在本过程调用后调用，这样会产生：
    '(1)新增病案时，该过程验证成功，但是病人ID重复，应该现验证病人ID
    '#33282# 使多台机器同时录入病案时，可能会导致病人ID 和 住院号 重复，此处新增病案时检查 病人ID 和 住院号是否重复，如果重复，则生成新的病人ID 和住院号
    If gclsPros.OpenMode = EM_新增病案 Then
        If Not gclsPros.IsExistPati Or gclsPros.OnlyPatiInfo Then
            If IsHavePageNos(CT_病人ID, True, gclsPros.病人ID) Then
                gclsPros.病人ID = GetNextNo(CT_病人ID)
            End If
            gclsPros.主页ID = 1
            strNo = Trim(gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text)
            If strNo = "" Then
                strNo = GetNextNo(CT_住院号)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text = strNo
            End If
            
            If Not gclsPros.NewInNo And gclsPros.OutFile = "" Then
                If IsHavePageNos(CT_住院号, True, strNo, gclsPros.病人ID) Then
                    strTmp = GetNextNo(CT_住院号)
                    If strNo <> "" Then
                        MsgBox "原" & strNo & "住院号已经存在,现在生成新的" & strTmp & "住院号！", vbInformation, gstrSysName
                        strNo = strTmp
                    End If
                End If
            End If
            gclsPros.InNo = strNo
            gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text = Trim(gclsPros.InNo)
        End If
    ElseIf gclsPros.OpenMode = EM_新增首页 Then
        strNo = Trim(gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text)
        If gclsPros.InNo = strNo Then
            If gclsPros.NewInNo And IsHavePageNos(CT_住院号, True, strNo, gclsPros.病人ID) Then
                strTmp = GetNextNo(CT_住院号)
                If strNo <> "" Then
                    MsgBox "原" & strNo & "住院号已经存在,现在生成新的" & strTmp & "住院号！", vbInformation, gstrSysName
                    strNo = strTmp
                End If
            End If
            gclsPros.InNo = strNo
            gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text = Trim(gclsPros.InNo)
        Else
            If MsgBox("原" & gclsPros.InNo & "住院号已改变成" & strNo & "住院号，不能保存首页，是否还原住院号？", vbYesNo, gstrSysName) = vbYes Then
                gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号).Text = gclsPros.InNo
            End If
        End If
    End If
    If gclsPros.NewInNo Then
        If Not (gclsPros.EditPageNo And gclsPros.CurrentForm.txtInfo(GC_病案号).Text <> "" And blnSave) Then
            gclsPros.CurrentForm.txtInfo(GC_病案号).Text = gclsPros.InNo
        End If
    End If

    '住院号检查
    If gclsPros.NewInNo Then
        If IsHavePageNos(CT_住院号, gclsPros.OpenMode = EM_编辑 Or gclsPros.Is编目, gclsPros.InNo, gclsPros.病人ID) Then
            Call ShowMessage(gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号), "住院号已经存在,请重新确定住院号!")
        End If
    End If
    
    '检查病案是否已经接收
    If Not gclsPros.EditUnrecive And (gclsPros.OpenMode = EM_新增病案 Or gclsPros.OpenMode = EM_新增首页) Then
        strSql = "Select ID From 病案接收记录 Where 病人id = [1] And 主页id = [2] And 接收时间 Is Not Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, gclsPros.主页ID)
        If rsTmp.RecordCount = 0 Then
            Call ShowMessage(gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号), "病人本次住院病案室还没有接收，不能进行编目操作!")
            Exit Function
        End If
    End If
    If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
    grsDeptInfo.Filter = "ID=" & gclsPros.出院科室ID: grsDeptInfo.Sort = "ID"
    If grsDeptInfo.RecordCount > 0 Then
        strOutDeptCode = grsDeptInfo!编码 & ""
    End If
                
    '53638:刘鹏飞,2013-05-10,档案号检查
    strNo = Trim(gclsPros.CurrentForm.txtInfo(GC_档案号).Text)
    If strNo <> "" Then
        If IsHavePageNos(CT_档案号, True, strNo, gclsPros.病人ID, gclsPros.主页ID) Then
            If gclsPros.UseFileRules Then
                If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
                strTmp = NVL(GetNextNo(CT_档案号, , strOutDeptCode))
                gclsPros.CurrentForm.txtInfo(GC_档案号).Text = strTmp
                MsgBox "原" & strNo & "档案号已经存在,现在使用新生成的" & strTmp & "档案号！", vbInformation, gstrSysName
            Else
                Call ShowMessage(gclsPros.CurrentForm.txtInfo(GC_档案号), "您输入的档案号已经被其他病人使用,请重新输入!")
                Exit Function
            End If
        End If
    Else
        If gclsPros.UseFileRules Then
            If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
            strTmp = NVL(GetNextNo(CT_档案号, , strOutDeptCode))
            gclsPros.CurrentForm.txtInfo(GC_档案号).Text = strTmp
        End If
    End If
    
    strNo = Trim(gclsPros.CurrentForm.txtInfo(GC_病案号).Text)
    
    '查询相同病案号或病人ID的信息
    strSql = "Select Nvl(a.病人id, [2]) 病人id, Nvl(a.主页id, [3]) 主页id, Nvl(a.病案号,  [1]) 病案号, b.姓名, b.性别, b.身份证号" & vbNewLine & _
                "From (Select 病人id, 主页id, 病案号 From 住院病案记录 Where 病案号 =  [1] Or 病人id = [2]) A" & vbNewLine & _
                "Full Join (Select 病人id, 主页id, 姓名, 性别, 身份证号 From 病人信息 Where 病人id = [2]) B" & vbNewLine & _
                "On a.病人id = b.病人id"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strNo, gclsPros.病人ID, gclsPros.主页ID)
    '病案号规则
    '1、若该病人在数据库中不存在，且未输入病案号，则新取一个病案号
    '2、若该病人在数据库中不存在，且输入病案号，则不做处理
    '3、若该病人在数据库中存在，且未输入病案号，则判断是否使用单独病案号
    '   (1)使用单独病案号，则新取一个病案号
    '   (2)未使用单独病案号，则取上一次的病案号
    '4、若该病人在数据库中存在，且输入病案号，则判断是否使用单独病案号
    '   (1)使用单独病案号，若存在相同病案号(若果是修改模式，则排除本次住院的病案号)，则新取一个病案号
    '   (2)使用单独病案号，若不存在相同病案号(若果是修改模式，则排除本次住院的病案号)，则不做处理
    '   (3)未使用单独病案号，若存在相同病案号且病人ID不相同的记录, 则判断姓名，性别、身份证号是否一致
    '         若不一致，则提示那些信息有差异，若一致，则不做处理
    '   (4)未使用单独病案号，若存在不相同病案号且病人ID相同的记录，则不作处理
    '            （这种情况在理论上不存在，是为了罗列所有情况才写出的）
    rsTmp.Filter = "病人id=" & gclsPros.病人ID
    blnDo = True: strTmpNo = ""
    If rsTmp.EOF Then
        If strNo <> "" Then blnDo = False
    Else
        If strNo = "" Then
            If Not gclsPros.SinPageNo Then
                rsTmp.Filter = "病人id=" & gclsPros.病人ID & " And 主页ID<>" & gclsPros.主页ID
                rsTmp.Sort = "主页ID"
                If Not rsTmp.EOF Then blnDo = False: strTmpNo = rsTmp!病案号 & ""
            End If
        Else
            If gclsPros.SinPageNo Then
                rsTmp.Filter = "病案号='" & strNo & "' "
                If rsTmp.EOF Then
                    blnDo = False
                ElseIf rsTmp.RecordCount = 1 Then
                    blnDo = Not (rsTmp!病人ID = gclsPros.病人ID And rsTmp!主页ID = gclsPros.主页ID)
                    blnSamePageNo = True
                Else
                    blnDo = True
                    blnSamePageNo = True
                End If
            Else
                rsTmp.Filter = "病案号='" & strNo & "' And 病人id<> " & gclsPros.病人ID
                rsTmp.Sort = "病案号,病人id,主页ID"
                If Not rsTmp.EOF Then
                    strTmp = zlStr.NeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_性别).Text)
                    If rsTmp!姓名 & "" = gclsPros.CurrentForm.txtInfo(GC_姓名).Text And rsTmp!性别 & "" = strTmp And rsTmp!身份证号 & "" <> gclsPros.CurrentForm.cboBaseInfo(BCC_身份证).Text Then
                        MsgBox "病人的病案号重复，并且同一病案号的两个病人姓名和性别相同，但身份证号不同。" & vbCrLf & _
                            "录入的病人：身份证号[" & gclsPros.CurrentForm.cboBaseInfo(BCC_身份证).Text & "]" & vbCrLf & _
                            "病案号重复的病人：身份证号[" & rsTmp!姓名 & "]", vbInformation, gstrSysName
                        Exit Function
                    ElseIf rsTmp!姓名 & "" <> gclsPros.CurrentForm.txtInfo(GC_姓名).Text Or rsTmp!性别 & "" <> strTmp Then
                        MsgBox "病人的病案号重复，并且同一病案号的两个病人，姓名或性别不同。" & vbCrLf & _
                            "录入的病人：姓名[" & gclsPros.CurrentForm.txtInfo(GC_姓名).Text & "],性别：[" & strTmp & "]" & vbCrLf & _
                            "病案号重复的病人：姓名[" & rsTmp!姓名 & "],性别：[" & rsTmp!性别 & "]", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                    blnDo = False
                End If
            End If
        End If
    End If
    '提取上一次的病案号
    If strTmpNo <> "" And strTmpNo <> strNo Then
        MsgBox "由于相同病人使用同一病案号并且病人已经存在病案号,将自动提取病人的其他病案号！", vbInformation, gstrSysName
        blnDo = False
    End If
    If blnDo Then
        If blnSamePageNo Then
            MsgBox "当前病案号已经被使用了,将自动获取病案号！", vbInformation, gstrSysName
        Else
            MsgBox "当前病案号不是有效病案号,将自动获取病案号！", vbInformation, gstrSysName
        End If
    End If
    
    bln顺序 = IsPageNosCodeRule(CT_病案号)
    Do While blnDo
        ' IIf(lngCount = 0, 0, 1)防止跳号
        strTmpNo = GetNextNo(CT_病案号, IIf(bln顺序 = True, IIf(lngCount = 0, 0, 1), 0), strOutDeptCode) & ""
        If strTmpNo = "" Then Exit Function
        blnDo = IsHavePageNos(CT_病案号, True, strTmpNo) '存在病案号则继续循环去取
        If (lngCount >= 100 Or Not bln顺序) And blnDo Then  '避免大量的循环，退出循环
            strTmpNo = ""
            If blnSave Then
                MsgBox "自动获取病案号失败,无法进行保存，请手动修改病案号或者联系管理员！", vbInformation, gstrSysName
                ValidatePageNos = False
                Exit Function
            Else
                MsgBox "自动获取病案号失败！", vbInformation, gstrSysName
                Exit Do
            End If
        End If
        lngCount = lngCount + 1
    Loop
    
    If strTmpNo <> "" Then
        gclsPros.CurrentForm.txtInfo(GC_病案号).Text = strTmpNo
        If strNo <> "" And strTmpNo <> strNo Then
            MsgBox "原" & strNo & "病案号已经存在,现在使用新生成的" & strTmpNo & "病案号！", vbInformation, gstrSysName
        End If
    End If
    ValidatePageNos = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsHavePageNos(ByVal intType As Integer, ByVal blnCurIn As Boolean, ParamArray arrInput() As Variant) As Boolean
'功能：是否存在号码
'参数：intType= 0-是否存在住院号，
'               1-其他病人是否使用了该住院号
'               2-是否存在该病人ID,
'               3-检查是除病人本次入院外还有其他地方使用了该病案号
'      blnCurIn=是否是本次住院或本人的
'返回=True-存在该号码,False-不存在该号码
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Select Case intType
        Case CT_住院号
            If Not blnCurIn Then
                strSql = "Select 1 From 病案主页 Where 住院号 = [1] and 主页ID > 0 "
            Else
                strSql = "Select 1 From 病案主页 Where 住院号 = [1] And 病人id <> [2] and 主页ID > 0 "
            End If
        Case CT_住院号ex
            strSql = "Select 病人ID  From 病人信息 Where 住院号 = [1]"
        Case CT_病人ID
            strSql = "Select 1 From 病案主页 Where 病人id = [1]"
        Case CT_档案号
            strSql = "Select  A.病人ID,A.档案号" & vbNewLine & _
                "From 住院病案记录 A" & vbNewLine & _
                "Where A.档案号 = [1] And Not Exists" & vbNewLine & _
                " (Select 1 From 住院病案记录 Where 病人id = [2] And 主页id = [3] And A.病人id = 病人id And A.主页id = 主页id) And A.病人ID<>[2]"
        Case CT_病案号
            strSql = "Select 1 From 住院病案记录 Where 病案号 = [1]"
    
    End Select
    Select Case UBound(arrInput)
        Case 0
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "判断住院病人相关号码存在", arrInput(0))
        Case 1
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "判断住院病人相关号码存在", arrInput(0), arrInput(1))
        Case 2
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "判断住院病人相关号码存在", arrInput(0), arrInput(1), arrInput(2))
    End Select
    
    If intType = CT_住院号 Then
        If gclsPros.OutFile = "" Then
            If rsTmp.EOF Then Exit Function
        Else
            If gclsPros.PatiOut.State = adStateOpen Then
                gclsPros.PatiOut.Filter = "住院号= " & IIf(Val(arrInput(0)) = 0, 0, arrInput(0))
                If gclsPros.PatiOut.EOF Then Exit Function
            End If
        End If
        IsHavePageNos = True
    Else
        IsHavePageNos = Not rsTmp.EOF
        If IsHavePageNos And intType = CT_住院号ex Then
            gclsPros.病人ID = Val(rsTmp!病人ID & "")
        End If
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiReSeeDoctor() As Boolean
'功能：判断病人本次是否复诊
    Dim rsTmp As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String
    Dim strSql As String
    Dim vsTmp As VSFlexGrid
    
    On Error GoTo errH
    
    '医生、科室与上次相同：没有转诊、续诊的
    strSQL1 = "Select 病人ID,执行人 as 医生,执行部门ID as 科室ID From 病人挂号记录 Where ID=[2] And 转诊科室ID Is Null And 续诊科室ID Is Null"
    
    strSQL2 = "Select Max(ID) as ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
            " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
    strSQL2 = "Select 病人ID,执行人 as 医生,执行部门ID as 科室ID From 病人挂号记录 Where ID=(" & strSQL2 & ") And 转诊科室ID Is Null And 续诊科室ID Is Null"
    
    strSql = "Select 1 From (" & strSQL1 & ") A,(" & strSQL2 & ") B Where A.病人ID=B.病人ID And A.医生=B.医生 And A.科室ID=B.科室ID"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "PatiReSeeDoctor", gclsPros.病人ID, gclsPros.主页ID)
    If rsTmp.EOF Then Exit Function
    
    '主要诊断与上次相同
    Set vsTmp = gclsPros.CurrentForm.vsDiagXY
    With vsTmp
        If .TextMatrix(.FixedRows, DI_诊断描述) <> "" Then
            strSql = "Select Max(ID) as 主页ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                    " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
            strSql = "Select 1 From 病人诊断记录" & _
                " Where 病人ID=[1] And 主页ID=(" & strSql & ")" & _
                " And 诊断类型=1 And 记录来源 IN(1,3) And 诊断次序=1" & _
                " And (疾病ID=[3] And 疾病ID<>0 Or 诊断ID=[4] And 诊断ID<>0 Or 诊断描述=[5])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "PatiReSeeDoctor", gclsPros.病人ID, gclsPros.主页ID, _
                Val(.TextMatrix(.FixedRows, DI_疾病ID)), Val(.TextMatrix(.FixedRows, DI_诊断ID)), .TextMatrix(.FixedRows, DI_诊断描述))
            If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
        End If
    End With
    
    If gclsPros.Have中医 Then
        Set vsTmp = gclsPros.CurrentForm.vsDiagZY
        With vsTmp
            If .TextMatrix(.FixedRows, DI_诊断描述) <> "" Then
                strSql = "Select Max(ID) as 主页ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                       " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
                strSql = "Select 1 From 病人诊断记录" & _
                    " Where 病人ID=[1] And 主页ID=(" & strSql & ")" & _
                    " And 诊断类型=11 And 记录来源 IN(1,3) And 诊断次序=1" & _
                    " And (疾病ID=[3] And 疾病ID<>0 Or 诊断ID=[4] And 诊断ID<>0 Or 诊断描述=[5])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "PatiReSeeDoctor", gclsPros.病人ID, gclsPros.主页ID, _
                    Val(.TextMatrix(.FixedRows, DI_疾病ID)), Val(.TextMatrix(.FixedRows, DI_诊断ID)), .TextMatrix(.FixedRows, DI_诊断描述))
                If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
            End If
        End With
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceIDByDiag(ByVal str医嘱IDs As String, ByVal lng诊断ID As Long) As String
'功能：根据诊断ID获取诊断相关医嘱ID
    Dim strTmp As String
    Dim lngPos As Long
    
    If str医嘱IDs <> "" And gclsPros.AdviceID <> 0 Then
        lngPos = InStr("," & str医嘱IDs & ",", "," & gclsPros.AdviceID & ",")
        If lngPos <= 0 Then
        '当前医嘱未关联当前诊断
            strTmp = str医嘱IDs
        Else
            strTmp = Replace("," & str医嘱IDs & ",", "," & gclsPros.AdviceID, "")
            If Len(strTmp) >= 2 Then
                strTmp = Mid(strTmp, 2, Len(strTmp) - 2)
            Else
                strTmp = ""
            End If
        End If
    Else
        strTmp = str医嘱IDs
    End If
    
    With gclsPros.DiagConn
        .Filter = "诊断ID=" & lng诊断ID & " And 标识ID<>" & gclsPros.AplyMark
        .Sort = "标识ID"
        Do While Not .EOF
            strTmp = strTmp & "," & !标识ID
            .MoveNext
        Loop
    End With
    
    GetAdviceIDByDiag = strTmp
End Function

Public Function GetPatiRoom(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取入出院病房
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select B.房间号 as 入院病房,c.房间号 as 出院病房  " & vbNewLine & _
            "From 病案主页 A, 床位状况记录 B,床位状况记录 C " & vbNewLine & _
            "Where A.病人id = [1] And A.主页id = [2] And A.入院病区id = B.病区id(+) And A.入院病床 = B.床号(+) And A.当前病区id = C.病区id(+)  And" & vbNewLine & _
            "      A.出院病床 = C.床号(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng病人ID, lng主页ID)
    Set GetPatiRoom = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInDeptTime(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strDefault As String) As String
'获取入科时间
'strDefault=空值时的返回值
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 开始时间 From 病人变动记录" & _
            " Where 病人ID=[1] And 主页ID=[2] And 开始原因 IN(2,1) And 开始时间 is Not Null Order by 开始原因 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng病人ID, lng主页ID)
    If rsTmp.EOF Then
        GetInDeptTime = strDefault
    Else
        If IsNull(rsTmp!开始时间) Then
            GetInDeptTime = strDefault
        Else
            GetInDeptTime = Format(rsTmp!开始时间, "yyyy-MM-dd HH:mm")
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ValiAndGet主页ID() As Boolean
'功能：验证或获取主页ID
'返回：是否成功

    Dim lngTmp As Long
    Dim lng次数 As Long
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    If gclsPros.OpenMode <> EM_新增病案 Then
        If Get住院次数Or主页id(gclsPros.病人ID, gclsPros.主页ID, lng次数, False) = False Then
            MsgBox "获取指定主页的次数失败,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        If gclsPros.OpenMode = EM_新增首页 Then
            If gclsPros.OnLine Then
                '获取当前次数的主页ID
                If Get住院次数Or主页id(gclsPros.病人ID, lngTmp, lng次数, True) = False Then
                    MsgBox "获取指定次数的主页失败,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
                If gclsPros.OnLineNew Then
                    gclsPros.主页ID = lngTmp + 1
                Else
                    '需要检查该主页是否已经建立了病案,如果建立了,则不能进行再次建立
                    If lngTmp < gclsPros.主页ID Then
                        ShowMsgbox "此病人在收费系统中不存在入院信息,不能创建病案了"
                        Exit Function
                    End If
                End If
            End If
            strSql = "Select 1 from 病案主页 where 病人id=[1] and 主页id=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取病案信息", gclsPros.病人ID, gclsPros.主页ID)
            If rsTmp.EOF Then
                lng次数 = lng次数 + 1
            End If
        End If
        gclsPros.CurrentForm.txtSpecificInfo(SLC_入院次数).Text = lng次数  '主页ID
    End If
    ValiAndGet主页ID = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReadPatPricture(ByVal lng病人ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '参数：lng病人ID=读取指定病人的照片
    '           imgPatient=照片加载位置
    '           strFile=照片的本地路径
    '74421,刘鹏飞,2014-07-04,读取病人照片信息
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = sys.Readlob(gclsPros.SysNo, 27, lng病人ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ExistInList(ByVal str住院号 As String, ByVal blnMessage As Boolean, Optional blnOnlyCheck As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能:检查住院号是否存在于住院病人清单
    '参数:
    '     blnOnlyCheck-仅为保存时检查
    '返回:存在返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strDate As String
    Dim strTmp As String
    
    If gclsPros.InputOutList = False Then ExistInList = True: Exit Function
    On Error GoTo errH
    
    strSql = "" & _
        "   Select A.姓名,A.日期,B.名称,B.ID " & _
        "   From 出院病人清单 A,部门表 B  " & _
        "   Where A.科室ID=B.ID and A.住院号= [1] " & _
        "   Order by A.日期 desc "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, str住院号)
    
    If rsTmp.RecordCount = 0 Then
        If blnMessage = True Then
            '106826:在病案接收编辑的时候，不考虑该病人是否已经产生住院日报清单
            ExistInList = True
        End If
        zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
        Exit Function
    End If

    '问题：22163和22577
    '检查方式：
    '1.如果是HIS共享模式，则不进行反写数据
    '2.如果是独立安装模式，则只有新增时，才反写数据，否则不反写数据
    '3.如果是保存检查调用，则不检查数据
    strDate = Format(rsTmp!日期 & "", "yyyy-mm-dd hh:mm")
    If strDate = "" Then strDate = "1989-01-01 " & Format(Now, "hh:mm")
    gclsPros.OutTime = strDate
    gclsPros.出院科室ID = Val(NVL(rsTmp!ID)) '设置出院科室会自动判断科室性质
    strTmp = rsTmp!名称 & ""
    If blnOnlyCheck Or gclsPros.ShareMedRec Or gclsPros.OpenMode = EM_编辑 Then
        ExistInList = True
        Exit Function
    End If
    gclsPros.CurrentForm.txtInfo(GC_姓名).Text = rsTmp!姓名 & ""
    gclsPros.CurrentForm.txtInfo(GC_出院科室).Text = rsTmp!名称 & ""
     gclsPros.CurrentForm.mskDateInfo(DC_出院时间).Text = strDate   '设置入院时间会自动重算住院天数
    ExistInList = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Select外部主页id(str住院号 As String, Optional int次数 As Integer = 0) As Integer
    Dim arrDate() As String
    Dim str提取病人 As String
    Dim vRect As RECT
    Dim rsTmp As New ADODB.Recordset
    Dim intTmp As Integer
    Dim strIfdate As String, blnALLPati As Boolean
    Dim objList As ListItem
    Dim lngTemp As Long
    Dim strKEY As String ''记录选择列的KEY
    Dim int主页id As Integer ''记录主页id
    Dim objTxt住院号 As TextBox
    
    ReDim arrDate(2)
    arrDate(0) = zlDatabase.GetPara("开始日期", gclsPros.SysNo, gclsPros.Module)
    arrDate(1) = zlDatabase.GetPara("结束日期", gclsPros.SysNo, gclsPros.Module)
    If arrDate(0) = "" Or arrDate(1) = "" Then
        arrDate(0) = "": arrDate(1) = ""
        blnALLPati = True
    Else
        arrDate(0) = Format(arrDate(0), "yyyy-mm-dd")
        arrDate(1) = Format(arrDate(1), "yyyy-mm-dd")
    End If
    If Not blnALLPati Then blnALLPati = Val(zlDatabase.GetPara("提取所有出院病人", gclsPros.SysNo, gclsPros.Module)) = 1
    
    If Val(zlDatabase.GetPara("提取24小时内出院病人", gclsPros.SysNo, gclsPros.Module)) <> 1 Then
        If gclsPros.OutFile = "" Then
            str提取病人 = " And (B.出院日期-B.入院日期)*24>=24"
        Else
            str提取病人 = "住院时间>=24"
        End If
    End If
    
    If gclsPros.EditUnrecive = False And gclsPros.OutFile = "" Then
        str提取病人 = str提取病人 & " And E.接收时间 IS NOT NULL"
    End If
    If Not blnALLPati Then
        If gclsPros.OutFile = "" Then
            strIfdate = " And B.出院日期 Between Trunc(To_Date('" & arrDate(0) & "','yyyy-mm-dd')) And Trunc(To_Date('" & arrDate(1) & "','yyyy-mm-dd'))+1-1/24/60/60"
        Else
            strIfdate = " 出院日期 >= #" & Format(arrDate(0), "yyyy-mm-dd 00:00:00") & "# and 出院日期 <= #" & Format(arrDate(1), "yyyy-mm-dd 23:59:59") & "#"
        End If
    End If
    Set objTxt住院号 = gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
    vRect = zlControl.GetControlRect(objTxt住院号.hwnd)
    If str提取病人 <> "" Then
        strIfdate = IIf(strIfdate = "", "", strIfdate & " and ") & str提取病人
        If strIfdate <> "" Then
            If IsNumeric(objTxt住院号.Text) Then
                 strIfdate = strIfdate & " and " & "住院号=" & Trim(objTxt住院号.Text) & " and 住院次数>" & int次数
            End If
        Else
            strIfdate = "住院号=" & Trim(objTxt住院号.Text) & " and 住院次数>" & int次数
        End If
    End If
    With frmPageMedRecNOSel
        .Top = vRect.Top + 300
        .Left = vRect.Left
        strKEY = .ShowMe(gclsPros.CurrentForm, gclsPros.PatiOut, strIfdate)
        objTxt住院号.Text = Split(strKEY, "_")(0)
        If Val(objTxt住院号.Text) = 0 Then
            objTxt住院号.Text = ""
            Exit Function
        Else
            gclsPros.InNo = objTxt住院号.Text
        End If
        int主页id = Split(strKEY, "_")(1)
        gclsPros.PatiOut.Filter = " 住院号=" & objTxt住院号.Text & " and 住院次数=" & int主页id
        gclsPros.主页ID = int主页id
        LoadDataFromOutFile (objTxt住院号.Text)
        Select外部主页id = int主页id
    End With

End Function

Public Function Select主页ID(lng病人ID As Long, Optional int次数 As Integer = 0) As Integer
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim objTxt住院号 As TextBox
    Dim strSql As String
    
    Set objTxt住院号 = gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
    vRect = zlControl.GetControlRect(objTxt住院号.hwnd)
    '刘兴宏:主页id与实际的住院次数不一至,因为实际的住院次数不包含留观病人
    '问题26488 by lesfeng 2010-03-18 结清
    '问题30659 by lesfeng 2010-06-04 病人余额 访问权限
    
    '56002:刘鹏飞,2012-11-23,进行编目的病人应该只包含正常入院的病人，即病人性质=0
    '39906:刘鹏飞,2013-05-07,添加病案接收标志
    If gclsPros.OnLine Then
        strSql = "" & _
            "   Select a.病人ID||'-'||a.主页id as ID,a.住院号,b.姓名,B.性别,TO_char(a.入院日期,'YYYY-MM-DD HH24:MI') as 入院日期, " & _
            "           TO_char(a.出院日期,'YYYY-MM-DD HH24:MI') as 出院日期,C.名称 出院科室," & _
            "         '第'||Zl_获取住院次数或主页id(a.病人id,a.主页id,0)||'次' as 住院次数,decode(D.费用余额,null,'是',0,'是','否') As 结清,Decode(E.接收时间, NULL, '否', '是') AS 接收" & _
            "   from 病案主页 a,病人信息 b,病人余额 D,部门表 C,病案接收记录 E " & _
            "   Where a.病人id=b.病人id And A.病人ID=E.病人ID(+) ANd A.主页ID=E.主页ID(+) and a.编目日期 is null and nvl(a.病人性质,0)=0  and a.出院日期 is not null " & _
            "           and a.病人id=[1]  and a.主页id>[2] and A.病人id = D.病人id(+)  And D.类型(+)=2 And A.出院科室ID=C.ID(+) " & _
            "   order by 入院日期 asc"
    Else
        strSql = "" & _
        "   Select a.病人ID||'-'||a.主页id as ID,a.住院号,B.姓名,B.性别,TO_char(a.入院日期,'YYYY-MM-DD HH24:MI') as 入院日期, " & _
        "           TO_char(a.出院日期,'YYYY-MM-DD HH24:MI') as 出院日期,C.名称 出院科室," & _
        "         '第'||Zl_获取住院次数或主页id(a.病人id,a.主页id,0)||'次' as 住院次数,Decode(e.接收时间, NULL, '否', '是') AS 接收" & _
        "   from 病案主页 a,病人信息 b,部门表 C,病案接收记录 E   " & _
        "   Where a.病人id=b.病人id And A.病人ID=E.病人ID(+) and A.主页ID=E.主页ID(+) and a.编目日期 is null and nvl(a.病人性质,0)=0  and a.出院日期 is not null " & _
        "           and a.病人id=[1]  and a.主页id>[2] And A.出院科室ID=C.ID(+) " & _
        "   order by 入院日期 asc"
    End If
       
   '刘兴宏:留观病人不能建立病案
   ' Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "病人住院次数", , , , , , True, lmx, lmy, 300, , , True)
    Set rsTemp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "病人住院次数", False, "", "", False, False, True, vRect.Left, vRect.Top, 300, blnCancel, False, True, lng病人ID, int次数)
    
    If blnCancel Then Select主页ID = 0: Exit Function
    If rsTemp Is Nothing Then Select主页ID = 0: Exit Function
    If rsTemp.State = 0 Then Select主页ID = 0: Exit Function
    
    If rsTemp.RecordCount > 0 Then
        Select主页ID = CInt(Mid(rsTemp!ID, InStr(rsTemp!ID, "-") + 1))
    Else
        Select主页ID = 0
    End If
End Function

Public Function GetDeptCode(ByVal lngDeptID As Long) As String
'51446,刘鹏飞,2012-08-02
'功能：根据科室ID获取科室编码
    Dim strCode As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo Errhand
    If lngDeptID <= 0 Then Exit Function
    strSql = "select 编码 From 部门表 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取科室编码", lngDeptID)
    If rsTemp.RecordCount > 0 Then strCode = NVL(rsTemp!编码)
    
    GetDeptCode = strCode
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub OpenExtraData()
    Dim cnAccess As New ADODB.Connection
    
    If gclsPros.OutFile = "" Or gclsPros.OpenMode = EM_查阅 Or gclsPros.OpenMode = EM_编辑 Then Exit Sub
        
    On Error Resume Next
    cnAccess.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gclsPros.OutFile & ";Persist Security Info=False"
    If Err <> 0 Then
        MsgBox "打不开可获取病人信息与费用情况的文件。", vbInformation, gstrSysName
        Exit Sub
    End If
    If gclsPros.PatiOut.State = 1 Then gclsPros.PatiOut.Close
    If gclsPros.FeesOut.State = 1 Then gclsPros.FeesOut.Close
    
    '打开所需的两个表
    gclsPros.PatiOut.Open "select *,clng((出院日期-入院日期)*24) as 住院时间 from 病案主页 order by 住院号,住院次数", cnAccess, adOpenStatic, adLockReadOnly
    gclsPros.FeesOut.Open "select * from 病人费用", cnAccess, adOpenStatic, adLockReadOnly
    
End Sub

Public Sub GetDaysFromLast()
'功能：获取离上次入院的天数
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim blnGet As Boolean, lngDay As Long, strSec As String
    
    If gclsPros.主页ID > 1 And gclsPros.InTime <> "" Then
        If gclsPros.FuncType = f病案首页 And gclsPros.MedPageSandard = ST_云南省标准 Then
            blnGet = gclsPros.CurrentForm.cboBaseInfo(BCC_距上次住院时间).ListIndex = -1
        ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
            blnGet = gclsPros.CurrentForm.txtSpecificInfo(SLC_距上次住院时间).Text = ""
        End If
        If blnGet Then
            strSql = "select (To_Date('" & Format(gclsPros.InTime, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')-出院日期) 时间差 from 病案主页 where 病人ID=[1] And 主页id =[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "距上一次入住本院时间", gclsPros.病人ID, gclsPros.主页ID - 1)
            If Not rsTmp.EOF Then
                lngDay = Val(NVL(rsTmp!时间差))
                If lngDay >= 2 And lngDay <= 15 Then
                    strSec = "2-15天"
                ElseIf lngDay >= 16 And lngDay <= 31 Then
                    strSec = "16-31天"
                ElseIf lngDay > 31 Then
                    strSec = "＞31天"
                Else
                    strSec = "当天"
                End If
                lngDay = IIf(lngDay < 1, 1, lngDay)
                If gclsPros.MedPageSandard = ST_四川省标准 Then
                    gclsPros.CurrentForm.txtSpecificInfo(SLC_距上次住院时间).Text = lngDay
                Else
                    Call Cbo.SeekIndex(gclsPros.CurrentForm.cboBaseInfo(BCC_距上次住院时间), strSec)
                End If
            End If
        End If
    End If
End Sub


