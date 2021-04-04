Attribute VB_Name = "mdlPatient"
Option Explicit '要求变量声明
'=======系统控制相关变量============
Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4   '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5 '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6   '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7 '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8     '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9 '结算时可使用个人帐户支付
    support未结清出院 = 10      '允许病人还有未结费用时出院
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
End Enum

Public gobjPublicPatient As Object                 '病人信息接口对象
Public gclsInsure As New clsInsure
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gobjXWHIS As Object     'RIS接口部件zl9XWInterface.clsHISInner
Public gblnXW As Boolean      '系统参数：“启用医学影像信息系统专业版接口”
Public gblnPatiByID As Boolean   '同一身份证只能对应一个建档病人

Public gobjPlugIn As Object   '插件对象

Public gstrUnitName As String
Public gstrSysName As String                '系统名称
Public gstrDBUser As String '当前用户名
Public gstrPatiTypeColor As String '病人颜色串  名称,颜色值|名称,颜色值----
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
    g本机公共模块 = 5
    g本机私有模块 = 6
End Enum

'内部应用模块号定义
Public Enum ENUM_INSIDE_PROGRAM
    P合约单位管理 = 1100
    P病人信息管理 = 1101
    P就诊卡发放管理 = 1102
    P预交款管理 = 1103
    P预交款操作日报 = 1104
    P合约单位费用 = 1105
    P病人费用审批 = 1106
End Enum

'结构化地址类型 1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址
Public Enum Enum_IX_ADDRESS
    E_IX_出生地点 = 1
    E_IX_籍贯 = 2
    E_IX_现住址 = 3
    E_IX_户口地址 = 4
    E_IX_联系人地址 = 5
End Enum

Public gint提醒剩余票据张数 As Integer      '收费时,票据在剩余X张后开始提醒收费员:-1代表不提醒
Public gobjSquare As SquareCard  '卡结算部件
'结构化地址
Public gbln启用结构化地址 As Boolean
Public gbln显示乡镇 As Boolean

Public Function InitPatiType() As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    gstrPatiTypeColor = ""
    gstrSQL = "select 名称,颜色 from 病人类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人类型")
    Do Until rsTemp.EOF
        gstrPatiTypeColor = gstrPatiTypeColor & rsTemp!名称 & "," & NVL(rsTemp!颜色, 0) & "|"
        rsTemp.MoveNext
    Loop
    If Len(gstrPatiTypeColor) > 0 Then
        gstrPatiTypeColor = Mid(gstrPatiTypeColor, 1, Len(gstrPatiTypeColor) - 1)
    Else
        gstrPatiTypeColor = "普通病人,0|医保病人,255"
    End If
    InitPatiType = True
    Exit Function
errH:
    gstrPatiTypeColor = "普通病人,0|医保病人,255"
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetPatiColor(ByVal strPatiType) As Long
Dim arrType As Variant, i As Integer
    arrType = Split(gstrPatiTypeColor, "|")
    For i = LBound(arrType) To UBound(arrType)
        If Split(arrType(i), ",")(0) = strPatiType Then
            GetPatiColor = Split(arrType(i), ",")(1)
            Exit Function
        End If
    Next
End Function

Public Function GetUnitID(bytFlag As Byte, lngID As Long) As Long
'功能：返回收费特定项目的执行科室
'参数：bytFlag=执行科室标志,lngID=收费细目ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '无明确科室
            GetUnitID = UserInfo.部门ID '取操作员所在科室
        Case 4 '指定科室
            strSQL = "Select B.执行科室ID From 收费项目目录 A,收费执行科室 B Where B.收费细目ID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                GetUnitID = rsTmp!执行科室ID '默认取第一个(如有多个)
            Else
                GetUnitID = UserInfo.部门ID '如没有指定，则取操作员所在科室
            End If
        Case 1, 2, 3 '病人科室,操作员科室
            GetUnitID = UserInfo.部门ID '都取操作员科室
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNOFromCard(strCardNo As String) As String
'功能：由就诊卡号获取就诊卡费用记录单据号
'说明：如果该单据已经作废，则病人不可能有卡号，则读不出来
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "" & _
    " Select NO " & _
    " From 住院费用记录 A,病人信息 B" & _
    " Where A.病人ID=B.病人ID And A.实际票号=B.就诊卡号" & _
    "       And A.记录性质=5 And A.记录状态=1 And A.序号=1 And B.就诊卡号=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strCardNo)

    
    If Not rsTmp.EOF Then GetNOFromCard = rsTmp!NO
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarINFO(lngID As Long) As ADODB.Recordset
'功能：获取与指定病人信息相似的病人信息
'参数：lngID=病人ID
    On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    'by lesfeng 2010-03-08 性能优化
    strTemp = " 病人ID,门诊号,住院号,姓名,性别,年龄,费别,出生日期,出生地点,身份证号,身份,国籍,民族,学历,职业," & _
              " 婚姻状况,家庭地址,家庭电话,工作单位,单位电话,住院次数,当前床号 入院时间,出院时间,当前病区ID,当前科室ID "
    strSQL = _
        "Select A.病人ID,A.门诊号,A.住院号,A.姓名,A.性别,A.年龄,A.费别," & _
        " To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,A.出生地点,A.身份证号," & _
        " A.身份,A.国籍,A.民族,A.学历,A.职业,A.婚姻状况,A.家庭地址,A.家庭电话," & _
        " A.工作单位,A.单位电话,A.住院次数,C.名称 as 病区,D.名称 as 科室," & _
        " A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间," & _
        " To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间 From " & _
        " (Select " & strTemp & " From 病人信息 Where 病人ID<>[1]) A, " & _
        " (Select " & strTemp & " From 病人信息 Where 病人ID =[1]) B,部门表 C,部门表 D " & _
        " Where A.当前病区ID=C.ID(+) And A.当前科室ID=D.ID(+)" & _
        " And Nvl(A.国籍,'X')=Nvl(B.国籍,'X') And Nvl(A.民族,'X')=Nvl(B.民族,'X')" & _
        " And Nvl(A.出生日期,Sysdate)=Nvl(B.出生日期,Sysdate)" & _
        " And Nvl(A.身份证号,'X')=Nvl(B.身份证号,'X') And Nvl(A.性别,'X')=Nvl(B.性别,'X')" & _
        " And A.姓名=B.姓名 Order by A.病人ID Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    Set SimilarINFO = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarIDs(str国籍 As String, str民族 As String, dat出生日期 As Date, str性别 As String, str姓名 As String, str身份证号 As String, ByRef rsRet As ADODB.Recordset) As String
'功能：检查病人是否存在相似信息
'返回：相似记录的病人ID串,如"234,235,236"
    On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    'by lesfeng 2010-03-08 性能优化 TO_DATE('" & Format(dat出生日期, "YYYY-MM-DD") & "','YYYY-MM-DD')
    strSQL = _
        " Select Rownum+1 ID,病人ID,门诊号,住院号,Nvl(身份证号,'未登记') 身份证号,Nvl(家庭地址,'未登记') 地址,To_Char(登记时间,'YYYY-MM-DD') 登记时间 " & _
        " From 病人信息 Where (国籍=[1] And 民族=[2] And 性别=[3] And 姓名=[4]" & _
        " And 出生日期=[6]) Or 身份证号=[5] " & _
        " Order by 病人ID Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", str国籍, str民族, str性别, str姓名, str身份证号, CDate(Format(dat出生日期, "YYYY-MM-DD")))
    For i = 1 To rsTmp.RecordCount
        SimilarIDs = SimilarIDs & "|ID:" & rsTmp!病人ID & ",门诊号:" & NVL(rsTmp!门诊号, "无") & ",住院号:" & NVL(rsTmp!住院号, "无") & ",身份证号:" & rsTmp!身份证号 & ",地址:" & rsTmp!地址 & ",登记日期:" & rsTmp!登记时间
        rsTmp.MoveNext
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    rsTmp.Filter = ""
    Set rsRet = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'这个函数没有使用
Public Function GetUnionName(lngID As Long) As String
'功能：获取合同单位名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 名称 From 合约单位 Where ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    
    If rsTmp.RecordCount <> 0 Then GetUnionName = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMax主页ID(lng病人ID As Long) As Long
'功能：获取病人的最大病案主页ID
'返回：
'     >0:成功
'      0:失败
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 主页ID From 病案主页 Where 病人ID=[1] Order by 主页ID Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng病人ID)
    
    If rsTmp.RecordCount = 0 Then
        GetMax主页ID = 1
    ElseIf IsNull(rsTmp!主页ID) Then
        GetMax主页ID = 1
    Else
        GetMax主页ID = rsTmp!主页ID + 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 一周内再次入院(lng病人ID, dat入院时间 As Date) As Boolean
'功能：判断病人是否在一周内再次入院
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 上次出院时间 From 病人信息 Where 病人ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng病人ID)
    
    If rsTmp.RecordCount = 0 Then
        Exit Function
    ElseIf IsNull(rsTmp!上次出院时间) Then
        Exit Function
    ElseIf Abs(CDate(Format(rsTmp!上次出院时间, "yyyy-MM-dd")) - CDate(Format(dat入院时间, "yyyy-MM-dd"))) <= 7 Then
        一周内再次入院 = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function isLookRoom(lngID As Long) As ADODB.Recordset
'功能：判断科室是否观察室
'返回：空=不是
    On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    'by lesfeng 2010-03-08 性能优化 Select *
    strSQL = "Select A.环境类别,A.ID,A.上级ID,A.编码,A.名称,A.简码,A.位置,A.末级,A.建档时间,A.撤档时间,A.站点" & _
             "  From 部门表 A,部门性质说明 B Where B.部门ID=A.ID And B.服务对象=1 And B.工作性质='护理' And A.ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    
    If rsTmp.RecordCount <> 0 Then
        Set isLookRoom = rsTmp
    Else
        Set isLookRoom = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set isLookRoom = New ADODB.Recordset
End Function

Public Function NextBedNo(lngUnitID As Long) As Long
'功能：获取指定病区的下一床位号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = _
        "Select A.床号 From 床位状况记录 A " & _
        "Where A.病区ID=[1] And Not Exists(Select 床号 From 床位状况记录 B Where B.床号=A.床号+1 And B.病区ID=A.病区ID) " & _
        "Order by A.床号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngUnitID)
    
    If rsTmp.RecordCount <> 0 Then
        NextBedNo = rsTmp!床号 + 1
    Else
        NextBedNo = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function isRepeat(lngUnitID As Long, strBeds As String) As String
'功能：判断在指定病区内的一系列床号是否已经存在
'参数：lngUnitID=病区ID,strBeds=床号字符串,如"12,13,15..."
'返回：空=都不存在,否则"12,13..."这些床号重复
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 床号 From 床位状况记录 Where 病区ID=[1] And instr(','||[2]||',',','||床号||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngUnitID, strBeds)
    
    If rsTmp.RecordCount <> 0 Then
        For i = 1 To rsTmp.RecordCount
            isRepeat = isRepeat & rsTmp!床号 & ","
            rsTmp.MoveNext
        Next
        isRepeat = Left(isRepeat, Len(isRepeat) - 1)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInPatiNO(str住院号 As String, Optional lng病人ID As Long) As Boolean
'功能：判断指定住院号是否已经存在于数据库中
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 1 From 病人信息 Where 住院号=[1] And 病人ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", str住院号, lng病人ID)
    
    If rsTmp.RecordCount > 0 Then ExistInPatiNO = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistClinicNO(str门诊号 As String, Optional lng病人ID As Long) As Boolean
'功能：判断指定门诊号是否已经存在于数据库中
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-08 性能优化 Select *
    strSQL = "Select 病人ID,门诊号 From 病人信息 Where 门诊号=[1] And 病人ID<>[2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", str门诊号, lng病人ID)
    
    If rsTmp.RecordCount > 0 Then ExistClinicNO = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function ExistInPatiID(lngID As Long) As Boolean
'功能：判断指定病人ID是否已经存在于数据库中
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-08 性能优化 Select *
    strSQL = "Select 病人ID From 病人信息 Where 病人ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    
    If rsTmp.RecordCount > 0 Then ExistInPatiID = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastInfo(lngID As Long) As String
'功能：获取病人最后一次预交款单位信息
'返回："缴款单位|单位开户行|单位帐号"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '大改：记录性质=1
    strSQL = "Select 缴款单位,单位开户行,单位帐号 From 病人预交记录 " & _
            " Where (缴款单位 is Not NULL Or 单位开户行 is Not NUll Or 单位帐号 is Not NULL) And 记录性质=1 And 病人ID=[1] Order by 收款时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
    
    If Not rsTmp.EOF Then
        GetLastInfo = IIf(IsNull(rsTmp!缴款单位), "", rsTmp!缴款单位) & "|" & IIf(IsNull(rsTmp!单位开户行), "", rsTmp!单位开户行) & "|" & IIf(IsNull(rsTmp!单位帐号), "", rsTmp!单位帐号)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub InitSysPar()
'功能：初始化系统参数
    Dim strValue As String
    On Error Resume Next
        
    gbln工本费 = zlDatabase.GetPara(3, glngSys) = "1"
                    
    '费用金额小数点位数
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '卡号显示方式
    gblnShowCard = Not ISPassShowCard 'zldatabase.GetPara(12, glngSys) = "0"
    
    '票据号码长度、就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    'gbyt磁卡 = Val(Split(strValue, "|")(0))
    gbyt预交 = Val(Split(strValue, "|")(1))
    gbytCardNOLen = Val(Split(strValue, "|")(4))
    'If gbyt磁卡 = 0 Then gbyt磁卡 = 7
    If gbyt预交 = 0 Then gbyt预交 = 7
    If gbytCardNOLen = 0 Then gbytCardNOLen = 7
    
    '票号严格控制
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnBill预交 = Mid(strValue, 2, 1) = "1"
    'gblnBill磁卡 = Mid(strValue, 5, 1) = "1"
    
    '就诊卡允许的字母前缀
    gstrCardMask = UCase(zlDatabase.GetPara(27, glngSys))
    
    '一卡通消费验证
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    gbyt预存款消费验卡 = Val(Split(strValue, "|")(0))
    gbln消费卡退费验卡 = zlDatabase.GetPara(282, glngSys) = "1"
    
    '入院登记时刷卡输入密码
    gblnCheckPass = Mid(zlDatabase.GetPara(46, glngSys, , "0000000000"), 5, 1) = "1"
    
    '刘兴洪 问题:????    日期:2010-12-06 23:38:53
    '费用单价保留位数
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    
    gblnXW = Val(zlDatabase.GetPara(255, glngSys)) = 1
    '同一身份证只能对应一个建档病人
    gblnPatiByID = Val(zlDatabase.GetPara(279, glngSys)) = 1
End Sub

Public Function ISPassShowCard() As Boolean
'功能：是否密文显示就诊卡号
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnPassShowCard As Boolean
    
    On Error GoTo errHandle
    strSQL = "Select 卡号密文 From 医疗卡类别 where 名称='就诊卡' and 是否固定=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡类别")
    If Not rsTemp.EOF Then
        blnPassShowCard = NVL(rsTemp!卡号密文) <> ""
    End If
    
    ISPassShowCard = blnPassShowCard
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SaveIDCard(bytStyle As Byte, strNO As String, lng病人ID As Long, lng主页ID As Long, _
        lng病人病区ID As Long, lng病人科室ID As Long, str标识号 As String, str费别 As String, _
        str原卡号 As String, str姓名 As String, str性别 As String, str年龄 As String, str卡号 As String, str密码 As String, _
        cur应收金额 As Currency, cur实收金额 As Currency, str结算方式 As String, Dat发卡时间 As Date, lng领用ID As Long, rsMoney As ADODB.Recordset, ByVal strICCard As String) As String
'功能：产生一条就诊卡费用记录SQL语句
'参数：bytStyle=0-发卡,1-补卡,2-换卡
'      cur金额=就诊卡金额
'      str结算方式=如果为空,表示记帐,不收现金
'      rsMoney:包括就诊卡收费信息的记录集
'      str原卡号=仅换卡时用
'      lng领用ID=当前可用的就诊卡领用ID
'      strICCard=IC卡号,通过读IC卡方式发卡时,同时填写病人信息的IC卡字段
    Dim lngUnitID As Long
    Dim strSQL As String
    
    '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
    Select Case rsMoney!科室标志
        Case 4 '指定科室
            lngUnitID = GetUnitID(rsMoney!科室标志, rsMoney!收费细目ID)
        Case 1, 2 '病人科室
            If lng病人科室ID <> 0 Then
                lngUnitID = lng病人科室ID
            Else
                lngUnitID = UserInfo.部门ID
            End If
        Case 0, 3, 5, 6
            lngUnitID = UserInfo.部门ID
    End Select
    
    '调用过程"zl_就诊卡记录_Insert"
    strSQL = "zl_就诊卡记录_INSERT(" & bytStyle & ",'" & strNO & "'," & lng病人ID & "," & lng主页ID & "," & _
        str标识号 & ",'" & str费别 & "','" & UCase(str原卡号) & "','" & str卡号 & "','" & str密码 & "','" & str姓名 & _
        "','" & str性别 & "','" & str年龄 & "'," & lng病人病区ID & "," & lng病人科室ID & "," & rsMoney!收费细目ID & _
        ",'" & rsMoney!收费类别 & "','" & IIf(IsNull(rsMoney!计算单位), "", rsMoney!计算单位) & "'," & _
        rsMoney!收入项目ID & ",'" & rsMoney!收据费目 & "'," & cur应收金额 & "," & lngUnitID & "," & UserInfo.部门ID & _
        ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(OverTime(Dat发卡时间), "1", "0") & _
        ",To_Date('" & Format(Dat发卡时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & str结算方式 & "'," & IIf(lng领用ID = 0, "NULL", lng领用ID) & ",'" & strICCard & "'," & cur应收金额 & "," & cur实收金额 & ")"
    
    SaveIDCard = strSQL
End Function

Public Sub InitLocPar(ByVal lngModul As Long)
'功能：初始化模块参数
    Dim strValue As String
    On Error Resume Next

    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = 0, "%", "")
    strValue = zlDatabase.GetPara("输入法")
    gstrIme = IIf(strValue = "", "不自动开启", strValue)
    gbytCode = Val(zlDatabase.GetPara("简码方式"))
    gblnMyStyle = zlDatabase.GetPara("使用个性化风格") = "1"
    
    If lngModul = P病人信息管理 Or lngModul = P就诊卡发放管理 Or lngModul = P预交款管理 Then
        gstr磁卡ID = zlDatabase.GetPara("共用就诊卡批次", glngSys, lngModul, "")
'        glng预交ID = zldatabase.GetPara("共用预交票据批次", glngSys, lngModul, 0)
        'LED语音报价
        gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
        gblnLedWelcome = Val(zlDatabase.GetPara("LED显示欢迎信息", glngSys, lngModul, 1)) <> 0
        gbln记账 = zlDatabase.GetPara("卡费记帐", glngSys, lngModul) = "1"
    End If
    
    If lngModul = P病人信息管理 Then
        gblnMustCard = zlDatabase.GetPara("建档同时必须发卡", glngSys, lngModul) = "1"
        gbln启用结构化地址 = Val(zlDatabase.GetPara("病人地址结构化录入", glngSys)) <> 0
        gbln显示乡镇 = Val(zlDatabase.GetPara("乡镇地址结构化录入", glngSys)) <> 0
    ElseIf lngModul = P就诊卡发放管理 Then
    
    ElseIf lngModul = P预交款管理 Then
        gblnShowHave = zlDatabase.GetPara("仅显有余款的缴款单", glngSys, lngModul) = "1"
        gblnAllowOut = zlDatabase.GetPara("允许出院病人缴住院预交", glngSys, lngModul) = "1"
        gblnBanIn = zlDatabase.GetPara("禁止在院病人缴门诊预交", glngSys, lngModul) = "1"
        gbln缴款科室 = zlDatabase.GetPara("允许更改缴款科室", glngSys, lngModul) = "1"
        strValue = Trim(zlDatabase.GetPara("票据剩余X张时开始提醒收费员", glngSys, lngModul, "0|10"))
        gbln分站点显示 = zlDatabase.GetPara("预交款分站点显示", glngSys, lngModul) = "1"
        '37372
        If Val(Split(strValue & "|", "|")(0)) = 0 Then
            gint提醒剩余票据张数 = -1
        Else
            gint提醒剩余票据张数 = Val(Split(strValue & "|", "|")(1))     '问题:26948
        End If
    End If
End Sub

Public Function GetArea(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'功能：获取地区列表或选择的地区
'参数：
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtInput.hWnd)
    If Not blnShowAll Then
        strSQL = " Select 编码 as ID,编码,名称,简码 From 区域" & _
                 " Where (编码 Like [1] Or upper(简码) Like '" & gstrLike & "'||[1]||'%' Or 名称 Like '" & gstrLike & "'||[1]||'%') And  NVL(级数,0)<3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSQL = "Select 编码 as ID,编码,名称,简码 From 区域 Where  NVL(级数,0)<3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnAuditedFee(lng病人ID As Long, Optional ByVal bln记账 As Boolean = True, Optional ByVal bytPrepayType As Byte = 0) As Currency
'功能：bln记账=true:获取病人未审核的划价单记帐费用
'      否则获取病人未缴划价金额
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If bytPrepayType = 0 Then
        strSQL = _
        " Select Sum(Nvl(金额,0)) as 金额 " & _
        " From (Select Sum(Nvl(实收金额,0)) as 金额 " & _
        "       From 门诊费用记录" & _
        "       Where 记帐费用=[2] And 记录状态=0 And 病人ID=[1]" & _
        "       Union ALL " & _
        "       Select Sum(Nvl(实收金额,0)) as 金额 " & _
        "       From 住院费用记录" & _
        "       Where 记帐费用=[2] And 记录状态=0 And 病人ID=[1] ) "
    ElseIf bytPrepayType = 1 Then
        strSQL = _
        " Select Sum(Nvl(实收金额,0)) as 金额 " & _
        "       From 门诊费用记录" & _
        "       Where 记帐费用=[2] And 记录状态=0 And 病人ID=[1]"
    Else
        strSQL = _
        " Select Sum(Nvl(实收金额,0)) as 金额 " & _
        "       From 住院费用记录" & _
        "       Where 记帐费用=[2] And 记录状态=0 And 病人ID=[1]"
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, IIf(bln记账, 1, 0))
    If Not rsTmp.EOF Then
        GetUnAuditedFee = Val("" & rsTmp!金额)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDictData(strDict As String) As ADODB.Recordset
'功能：从指定的字典中读取数据
'参数：strDict=字典对应的表名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If strDict = "区域" Then
        strSQL = "Select 编码,名称,0 as 缺省 From " & strDict & " Order by 编码"
    Else
        strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlPatient")
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInsure(strNO As String) As Integer
'功能：判断收费记录中是否存在指定的医保结算方式
'参数：strNO=收费单据号
'返回：如果存在,则返回该医保结算方式结算当时的保险类型
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.险类 From 病人预交记录 A,保险结算记录 B" & _
        " Where A.记录性质=1 And A.记录状态=1 And A.NO=[1]" & _
        " And A.ID=B.记录ID And B.性质=3"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strNO)
    
    If Not rsTmp.EOF Then
        ExistInsure = IIf(IsNull(rsTmp!险类), 0, rsTmp!险类)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistFeeInsurePatient(lng病人ID As Long) As Boolean
'功能：判断医保病人是否存在未结费用
'返回：
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    strSQL = "Select Nvl(sum(B.费用余额),0) 费用余额 From 病人信息 A,病人余额 B Where A.病人ID=B.病人ID And Nvl(A.险类,0)<>0 And A.病人ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng病人ID)
    
    If Not rsTmp.EOF Then ExistFeeInsurePatient = (rsTmp!费用余额 <> 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveSpare(strNO As String) As Double
'功能：根据预交单据号判断病人是否还有预交余额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(B.预交余额,0) as 余额" & _
        " From 病人预交记录 A,病人余额 B" & _
        " Where A.记录性质=1 And A.记录状态 IN(1,3) and nvl(A.预交类别,2)=b.类型(+)" & _
        " And A.NO=[1] And A.病人ID=B.病人ID"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strNO)
    
    If Not rsTmp.EOF Then HaveSpare = Val(NVL(rsTmp!余额))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveBalance(strNO As String) As Double
'功能：根据预交单据号判断该单据是否被结帐
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Sum(Nvl(冲预交,0)) as 冲预交" & _
        " From 病人预交记录 Where NO=[1] And 记录性质 IN(1,11)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", strNO)
    
    If Not rsTmp.EOF Then HaveBalance = Val(NVL(rsTmp!冲预交))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function StrToNum(ByVal strNumber As String) As Double
    '功能:将字符串转换成数据
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function

Public Function zlIsExistsSquareCard(ByRef lng预交ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查该单据是否为卡结算单据
    '入参:strNo-预交单的单据号
    '返回:存在,则返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 17:02:55
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    
    On Error GoTo errHandle
    strSQL = "Select 1 From 病人卡结算记录 A Where a.结算ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查预交单是否存在刷卡记录", lng预交ID)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNoBeginTrans:没有事务开始
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
End Sub
Public Function zl_Get医疗卡类型(lngTypeId As Long) As String()
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据医疗类型ID获取医疗类型
    '入参:lngTypeID-医疗卡类型ID
    '返回:类型对象
    '编制:王吉
    '日期:2012-07-06
    '问题号:51072
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim arr(3) As String
    
    strSQL = "" & _
    "       Select 密码长度,密码输入限制,是否缺省密码 " & _
    "       From 医疗卡类别 " & _
    "       Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医疗卡类别", lngTypeId)
    If rsTemp Is Nothing Then zl_Get医疗卡类型 = arr: Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Get医疗卡类型 = arr: Exit Function
    rsTemp.MoveFirst
    arr(0) = NVL(rsTemp!密码长度, "0")
    arr(1) = NVL(rsTemp!密码输入限制, "0")
    arr(2) = NVL(rsTemp!是否缺省密码, "0")
    zl_Get医疗卡类型 = arr
End Function

Public Function 是否已经签约(strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查需要绑定的卡号是否已经签约
    '入参:绑定卡号
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng身份证类别ID As Long
    Dim rsTemp As Recordset
    On Error GoTo Errhand:
    
    lng身份证类别ID = Get医疗卡类别ID("二代身份证")
    strSQL = "" & _
    "   Select Count(1) as 是否签约 From 病人医疗卡信息 Where 卡号=[1] And 卡类别ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡绑定", strCardNo, lng身份证类别ID)
    是否已经签约 = rsTemp!是否签约 > 0
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function


Public Sub AddSQL绑定卡(ByVal lng病人ID As Long, 卡类别ID As Long, strCard As String, strPassWord As String, ByVal dtCurdate As Date, blnICCard As Boolean, ByRef cllPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定卡处理
    '入参:lng病人ID;strCard-绑定卡号;strPassWord-加密密码
    '出参:lngCard结帐ID-卡费的结帐ID
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str变动原因 As String
    Dim strICCard As String
    
    strICCard = IIf(blnICCard, strCard, "")
    str变动原因 = "病人挂号发卡"
          'Zl_医疗卡变动_Insert
          strSQL = "Zl_医疗卡变动_Insert("
          '      变动类型_In   Number,
          '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
          strSQL = strSQL & "" & 11 & ","
          '      病人id_In     住院费用记录.病人id%Type,
          strSQL = strSQL & "" & lng病人ID & ","
          '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
          strSQL = strSQL & "" & 卡类别ID & ","
          '      原卡号_In     病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'',"
          '      医疗卡号_In   病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      变动原因_In   病人医疗卡变动.变动原因%Type,
          '      --变动原因_In:如果密码调整，变动原因为密码.加密的
          strSQL = strSQL & "'" & str变动原因 & "',"
          '      密码_In       病人信息.卡验证码%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      操作员姓名_In 住院费用记录.操作员姓名%Type,
          strSQL = strSQL & "'" & UserInfo.姓名 & "',"
          '      变动时间_In   住院费用记录.登记时间%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
          strSQL = strSQL & "NULL)"
     zlAddArray cllPro, strSQL
End Sub

Public Function Get医疗卡类别ID(strTypeName As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别ID
    '入参:strTypeName 医疗卡类别名称
    '返回:医疗卡类别ID
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    strSQL = "" & _
    "   Select ID From 医疗卡类别 Where 名称=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡类别", strTypeName)
    If rsTemp Is Nothing Then Get医疗卡类别ID = 0: Exit Function
    If rsTemp.RecordCount <= 0 Then Get医疗卡类别ID = 0: Exit Function
    Get医疗卡类别ID = rsTemp!ID
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zl当前用户身份证是否绑定(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断当前用户身份证是否已被绑定
    '入参:lng病人ID
    '返回:True 已绑定 false 未绑定
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lng身份证类别ID As Long
    Dim rsTemp As Recordset
    On Error GoTo Errhand
    lng身份证类别ID = Get医疗卡类别ID("二代身份证")
    strSQL = "" & _
    " Select count(1) as 是否绑定 From 病人信息 A,病人医疗卡信息 B Where A.身份证号 =B.卡号 And A.病人ID=B.病人ID And A.病人ID=[1] And B.卡类别ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡绑定", lng病人ID, lng身份证类别ID)
    zl当前用户身份证是否绑定 = rsTemp!是否绑定 > 0
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub
Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
End Sub

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str病人类型 As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人类型,设置不同病人类型的显示颜色
    '入参:objPatiControl-病人控件(文本框,标签)
    '    str病人类型-病人类型
    '    lngDefaultColor-缺省病人的显示颜色
    '返回:True-设置颜色成功，False-失败
    '编制:李南春
    '日期:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str病人类型 <> "" Then
        lngColor = zlDatabase.GetPatiColor(str病人类型)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthDay As String = "", Optional ByVal datCalc As Date) As String
    '功能:年龄合法性检查
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    strBirthDay = Format(strBirthDay, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthDay) Then
        If datCalc = CDate(0) Then
            strSQL = "select Zl_Age_Check([1],[2]) From dual"
        Else
            strSQL = "select Zl_Age_Check([1],[2],[3]) From dual"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge, CDate(strBirthDay), datCalc)
    Else
        strSQL = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge)
    End If
    CheckAge = NVL(rsTemp.Fields(0).Value)
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreatePublicPatient() As Boolean
    If gobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set gobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        If gobjPublicPatient Is Nothing Then
            MsgBox "创建病人信息公共部件(zlPublicPatient.clsPublicPatient)失败!", vbInformation, gstrSysName
        Else
            Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser)
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjPublicPatient Is Nothing Then CreatePublicPatient = True
End Function

Public Sub LoadStructAddressDef(ByRef strAddress() As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取区域表中的缺省地址
    '入参:PatiAddress-结构化地址控件
    '返回:
    '编制:余伟节
    '日期:2016/1/7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    strSQL = "Select 级数,名称,level From 区域 " & _
            " Start With 缺省标志=1 " & _
            " Connect by Prior 上级编码=编码 " & _
            " Order by level Desc "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "缺省区域")
    If rsTmp.RecordCount = 0 Then Exit Sub
    Do While Not rsTmp.EOF
        strAddress(Val(NVL(rsTmp!级数))) = NVL(rsTmp!名称)
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReadStructAddress(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef PatiAddress As Object)
'功能:读取结构化地址
    Dim i As Long
    Dim rsStruct As ADODB.Recordset
    Dim rsAddress As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select a.省, a.市, a.县, a.乡镇, a.其他, a.地址类别 From 病人地址信息 A Where a.病人id = [1] And NVL(a.主页id,0) = [2]"
    Set rsStruct = zlDatabase.OpenSQLRecord(strSQL, "读取病人结构化地址", lng病人ID, lng主页ID)
    
    For i = PatiAddress.LBound To PatiAddress.UBound
        rsStruct.Filter = "地址类别=" & i
        If rsStruct.RecordCount > 0 Then
            Call PatiAddress(i).LoadStructAdress(rsStruct!省 & "", rsStruct!市 & "", rsStruct!县 & "", rsStruct!乡镇 & "", rsStruct!其他 & "")
        Else
            If rsAddress Is Nothing Then
                '同一个病人只读取一次
                If lng主页ID <> 0 Then
                    strSQL = "Select c.出生地点, c.籍贯, Nvl(b.家庭地址, c.家庭地址) As 现住址, Nvl(b.户口地址, c.户口地址) As 户口地址, Nvl(b.联系人地址, c.联系人地址) As 联系人地址" & vbNewLine & _
                        "From 病案主页 B, 病人信息 C" & vbNewLine & _
                        "Where b.病人id = c.病人id And b.病人id = [1] And b.主页id = [2] "
                Else
                    strSQL = "Select c.出生地点, c.籍贯, c.家庭地址 As 现住址,  c.户口地址 As 户口地址,c.联系人地址 As 联系人地址 " & vbNewLine & _
                        "From 病人信息 C" & vbNewLine & _
                        "Where c.病人id = [1] "
                End If
                Set rsAddress = zlDatabase.OpenSQLRecord(strSQL, "读取病人结构化地址", lng病人ID, lng主页ID)
            End If
            If rsAddress.RecordCount > 0 Then
                If NVL(rsAddress.Fields(PatiAddress(i).Tag).Value, "") <> "" Then
                    PatiAddress(i).Value = NVL(rsAddress.Fields(PatiAddress(i).Tag).Value, "")    '兼容启用结构化地址之前的数据
                End If
            End If
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub CreateStructAddressSQL(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef varSQL As Variant, ByRef PatiAddress As Object, Optional ByVal bytFunc As Byte = 0)
'功能:创建结构化地址SQL
'参数:
'PatiAddress-结构化地址控件数组名
'varSQL-返回的SQL数组集合\或者是集合对象
'bytFunc 可选参数:=1 当控件值为空时,代表删除
    Dim i As Long
    Dim strSQL As String
    
    For i = PatiAddress.LBound To PatiAddress.UBound
        If PatiAddress(i).Value <> "" Then
            '新增\修改
            strSQL = "zl_病人地址信息_update(1," & lng病人ID & "," & IIf(lng主页ID = 0, "NULL", lng主页ID) & "," & i & ",'" & PatiAddress(i).value省 & "','" & PatiAddress(i).value市 & "','" & PatiAddress(i).value区县 & "','" & PatiAddress(i).value乡镇 & "','" & PatiAddress(i).value详细地址 & "','" & PatiAddress(i).Code & "')"
            If IsArray(varSQL) Then
                ReDim Preserve varSQL(UBound(varSQL) + 1)
                varSQL(UBound(varSQL)) = strSQL
            ElseIf UCase(TypeName(varSQL)) = UCase("Collection") Then
                varSQL.Add strSQL, "K" & (varSQL.Count + 1)
            End If
        Else
            '删除
            If bytFunc = 1 Then
                strSQL = "zl_病人地址信息_update(2," & lng病人ID & "," & IIf(lng主页ID = 0, "NULL", lng主页ID) & "," & i & ")"
                If IsArray(varSQL) Then
                    ReDim Preserve varSQL(UBound(varSQL) + 1)
                    varSQL(UBound(varSQL)) = strSQL
                ElseIf UCase(TypeName(varSQL)) = UCase("Collection") Then
                    varSQL.Add strSQL, "K" & (varSQL.Count + 1)
                End If
            End If
        End If
    Next
End Sub

Public Function CreateXWHIS(Optional ByVal blnMsg As Boolean) As Boolean
'功能：判断 RIS接口部件(zl9XWInterface.clsHISInner) 是否存在，并启用
'参数：blnMsg－创建失败时是否提示

    If Not gblnXW Then Exit Function
    If Not gobjXWHIS Is Nothing Then CreateXWHIS = True: Exit Function
    
    On Error Resume Next
    Set gobjXWHIS = GetObject(, "zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    On Error Resume Next
    If gobjXWHIS Is Nothing Then Set gobjXWHIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    If gobjXWHIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS接口部件(zl9XWInterface)未创建成功！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CreateXWHIS = True
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject(, "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String, Optional ByRef strErr As String = "0")
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    Dim strMsg As String
    
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        strMsg = "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description
        If strErr = "0" Then
            MsgBox strMsg, vbInformation, gstrSysName
        Else
            strErr = strMsg
        End If
    End If
End Sub
