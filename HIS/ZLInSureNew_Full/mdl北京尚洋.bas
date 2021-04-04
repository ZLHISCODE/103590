Attribute VB_Name = "mdl北京尚洋"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)
Public gcn尚洋 As New ADODB.Connection, gint适用地区_尚洋 As Integer
Public gint是否职工 As Integer
Private mcur统筹金额 As Currency, mcur个帐支付 As Currency

Public Function 医保初始化_北京尚洋() As Boolean
'功能：测试是否可以连接到前置服务器上
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs北京尚洋 As New ADODB.Recordset, str参数值 As String
'    '如果连接已经打开，那就不用再测试
'    If gcn尚洋.State = adStateOpen Then
'        医保初始化_北京尚洋 = True
'        Exit Function
'    End If
'
'    On Error GoTo ErrH
'
'    '首先读出参数，打开连接
'    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & TYPE_北京尚洋
'    Call OpenRecordset(rsTemp, gstrSysName)
'    Do Until rsTemp.EOF
'        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
'        Select Case rsTemp("参数名")
'            Case "用户名"
'                strUser = str参数值
'            Case "服务器"
'                strServer = str参数值
'            Case "用户密码"
'                strPass = str参数值
'            Case "适用地区"
'                gint适用地区_尚洋 = Val(str参数值)
'            Case "统筹区号"
'                gstr医保机构编码 = str参数值
'        End Select
'        rsTemp.MoveNext
'    Loop
'    If strUser = "" Or strServer = "" Or strPass = "" Then
'        MsgBox "参数设置不完整,请到医保参数设置中重新设置", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    On Error Resume Next
'    If gint适用地区_尚洋 = 1 Then
''        gcn尚洋.ConnectionString = "Provider=Sybase.ASEOLEDBProvider.2;口令=" & strPass & ";持续安全性信息=True;用户 ID=" & strUser & ";数据源=" & strServer
'        gcn尚洋.ConnectionString = "Provider=MSDASQL.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
'    Else
'        gcn尚洋.ConnectionString = "Provider=MSDAORA.1;Password=" & strPass & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
'    End If
'    gcn尚洋.CursorLocation = adUseClient
'    gcn尚洋.Open
'
'    If Err <> 0 Then
'        MsgBox "连接前置服务器发生错误。", vbInformation, gstrSysName
'        医保初始化_北京尚洋 = False
'        Exit Function
'    End If
    医保初始化_北京尚洋 = True
'    Exit Function
'ErrH:
'    If ErrCenter() = 1 Then Resume
'    医保初始化_北京尚洋 = False
End Function


Public Function 建立医保连接() As Boolean
'功能：测试是否可以连接到前置服务器上
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs北京尚洋 As New ADODB.Recordset, str参数值 As String

     
    On Error GoTo ErrH
    
    
'    If MsgBox("该病人是否职工医保?", vbYesNo + vbQuestion + vbDefaultButton1, "医保接口") = vbYes Then
        
        gint是否职工 = 1
        
        '首先读出参数，打开连接
        gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋)
        Do Until rsTemp.EOF
            str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Select Case rsTemp("参数名")
                Case "用户名"
                    strUser = str参数值
                Case "服务器"
                    strServer = str参数值
                Case "用户密码"
                    strPass = str参数值
                Case "适用地区"
                    gint适用地区_尚洋 = Val(str参数值)
                Case "统筹区号"
                    gstr医保机构编码 = str参数值
            End Select
            rsTemp.MoveNext
        Loop
        If strUser = "" Or strServer = "" Or strPass = "" Then
            MsgBox "参数设置不完整,请到医保参数设置中重新设置", vbInformation, gstrSysName
            Exit Function
        End If
        
'    Else
'        gint是否职工 = 0
'
'        strUser = "login_sa"
'        strServer = "sybase"
'        strPass = "passwd"
'        gint适用地区_尚洋 = 1
'        gstr医保机构编码 = "1403000000"
'    End If
'
    
    '如果连接已经打开，那就不用再测试
    If gcn尚洋.State = adStateOpen Then
        建立医保连接 = True
        Exit Function
    End If
    
    
    On Error Resume Next
    If gint适用地区_尚洋 = 1 Then
        gcn尚洋.ConnectionString = "Provider=MSDASQL.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
    Else
        gcn尚洋.ConnectionString = "Provider=MSDAORA.1;Password=" & strPass & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
    End If
    gcn尚洋.CursorLocation = adUseClient
    gcn尚洋.Open
    
    If Err <> 0 Then
        MsgBox "连接前置服务器发生错误。", vbInformation, gstrSysName
        建立医保连接 = False
        Exit Function
    End If
    建立医保连接 = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    建立医保连接 = False
End Function

Public Function 医保设置_北京尚洋() As Boolean
    医保设置_北京尚洋 = frmSet北京尚洋.参数设置()
End Function

Public Function 个人余额_北京尚洋(lng病人ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_北京尚洋)
    '因为能提取个帐余额,因此赋予足够大的数,具体支付金额由医保决定
    个人余额_北京尚洋 = Nvl(rsTemp!帐户余额, 1000000)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function 身份标识_北京尚洋(Optional bytType As Byte = 0, Optional lng病人ID As Long = 0) As String
    '北京尚洋医保没提供专门的身分验证接口
    Dim strTemp As String
    
    '韦华荣修改于2011-3-2
    If 建立医保连接 = False Then Exit Function
    
    
    strTemp = frmIdentify北京尚洋.Identify(bytType, lng病人ID)
    Unload frmIdentify北京尚洋
    If strTemp = "" Then
        MsgBox "未提取病人信息", vbInformation, gstrSysName
    Else
        身份标识_北京尚洋 = strTemp
    End If
End Function
'
'Public Function 门诊虚拟结算_北京尚洋(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
''因为北京尚洋未提供预结算接口，因此所得到的结算数据为医保结算的正式数据，即得到数据时医保已正式结算
'    Dim str流水号 As String, lng病人ID As Long, datCurr As Date, strSql As String, strTemp As String
'    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng序号 As Long, str科室 As String
'    Dim strCardNO As String, str收据项目 As String, str科室项目 As String, str会计项目 As String
''    病人ID         adBigInt, 19, adFldIsNullable
''    收费类别       adVarChar, 2, adFldIsNullable
''    收据费目       adVarChar, 20, adFldIsNullable
''    计算单位       adVarChar, 6, adFldIsNullable
''    开单人         adVarChar, 20, adFldIsNullable
''    收费细目ID     adBigInt, 19, adFldIsNullable
''    数量           adSingle, 15, adFldIsNullable
''    单价           adSingle, 15, adFldIsNullable
''    实收金额       adSingle, 15, adFldIsNullable
''    统筹金额       adSingle, 15, adFldIsNullable
''    保险支付大类ID adBigInt, 19, adFldIsNullable
''    是否医保       adBigInt, 19, adFldIsNullable
''    摘要           adVarChar, 200, adFldIsNullable
''    是否急诊       adBigInt, 19, adFldIsNullable
''    str结算方式  "报销方式;金额;是否允许修改|...."
'    On Error GoTo errHandle
'    If rs明细.RecordCount = 0 Then
'        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    datCurr = zlDatabase.Currentdate
'    lng病人ID = rs明细(0)
'    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=" & lng病人ID & " And 险类=" & TYPE_北京尚洋
'    Call OpenRecordset(rsTemp, gstrSysName)
'    If rsTemp.EOF Then
'        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
'        Exit Function
'    End If
'    strCardNO = rsTemp!卡号
'    '生成流水号
'    str流水号 = toHex(Format(datCurr, "YYMMDDHHMMSS") & Format(lng病人ID, "0######"), 35)
'
'    '判断是否有医保编码未对应
'    Do Until rs明细.EOF
'        gstrSQL = "select A.项目编码,B.名称,B.说明 from (select * from 保险支付项目 where 险类=" & TYPE_北京尚洋 & ") A, 收费细目 B where A.收费细目id(+)=B.id and B.id = " & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        If IsNull(rsTemp!项目编码) Then
'            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
'            Exit Function
'        End If
'        If IsNull(rsTemp!说明) Then
'            MsgBox "不能确定项目<" & rsTemp!名称 & ">的收据项目类别和科室核算类别", vbInformation, gstrSysName
'            Exit Function
'        ElseIf Len(rsTemp!说明) < 2 Then
'            MsgBox "不能确定项目<" & rsTemp!名称 & ">的科室核算类别", vbInformation, gstrSysName
'            Exit Function
'        End If
'        strTemp = rsTemp!项目编码
'        strSql = "Select * From PARA_CAPTURE_ITEM Where Areaid='" & gstr医保机构编码 & "' And Item_Code='" & UCase(rsTemp!项目编码) & "'"
'        Set rsTemp = gcn尚洋.Execute(strSql)
'        If rsTemp.EOF Then
'            MsgBox "在中间数据中未找到编码为[" & UCase(strTemp) & "]的项目，请核查", vbInformation, gstrSysName
'            Exit Function
'        End If
'        rs明细.MoveNext
'    Loop
'
'    '生成DBF文件
'    lng序号 = 1
'    rs明细.MoveFirst
'    While Not rs明细.EOF
'        gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        strTemp = rsTemp!说明      '说明不能为空，其中第一位存放收据项目类别，第二位存放科室核算类别
'        str收据项目 = Left(strTemp, 1)
'        str科室类别 = Mid(strTemp, 2, 1)
'        If rsTemp!类别 = 5 Or rsTemp!类别 = 6 Or rsTemp!类别 = 7 Then
'            str会计类别 = "A"       '药品
'        Else
'            str会计类别 = "B"       '医疗
'        End If
'
'        gstrSQL = "Select 项目编码 From 保险支付项目 Where 险类=" & TYPE_北京尚洋 & " And 收费细目id=" & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, gstrSysName)             '因为之前检查了是否进行对码，所以读出的记录一定不会空
'
'        strSql = "Select * From PARA_CAPTURE_ITEM Where Areaid='" & gstr医保机构编码 & "' And Item_Code='" & UCase(rsTemp!项目编码) & "'"
'        Set rsTemp = gcn尚洋.Execute(strSql)
'        '卡号、流水号、序号、内码、名称、科目编码、规格、计量单位、数量、单价、金额、自费金额
''    VISIT_NUMBER                char(18)        not null,   //处方号码
''    ITEM_NO                     numeric(6, 0)   not null,   //同一处方中项目序号
''    ITEM_CLASS                  char(1)         not null,   //收费项目类别:如A西药
''    ITEM_CODE                   char(12)        not null,   //项目编码
''    ITEM_NAME                   char(40)        not null,   //项目名称
''    SPEC                        varchar(50)     not null,   //规格
''    PRICE_UNIT                  char(8)         not null,   //计价单位
''    PRICE                       numeric(9, 4)   not null,   //单价
''    QUANTITY                    numeric(6, 2)   not null,   //数量
''    COST                        numeric(8, 2)   not null,   //金额
''    RECEIPT_CLASS               char(1)         not null,   //收据项目分类
''    COLLATE_RELATION            char(12)        null,       //与医保中心对应关系
''    OPERATOR                    char(15)        null,       //经办人
''    OPERATE_TIME                datetime        null,       //经办日期
''    CLINIC_FLAG                 numeric(1, 0)   not null,   //门诊/住院标志
''    EXE_DEPT                    char(20)        null,       //执行科室
''    APP_DOCTOR                  char(30)        null,       //开方医生
''    APP_DEPT                    char(20)        null,       //开单科室
''    TAKE_MEDICINE_FLAG          char(8)         not null,   //出院带药标志
''    ITEM_NO_DEPT_STAT           char(2)         null,       //科室核算项目类别
''    ITEM_NO_ACCOUNTANT_ITEM char(2)         null,       //会计核算项目类别
''    constraint PK_SICK_PRICE_ITEM PRIMARY KEY CLUSTERED (VISIT_NUMBER, ITEM_NO)
''A 西药，B 成药，C 草药，D 治疗，E 检查，F 放射，G 化验，H 手术费，I 输血费，J 输氧费，K CT。ECT，L 其它，M B超，N 心电图，O 脑电图，P 胃镜，Q 喉镜
'
'        gcn尚洋.Execute "Insert Into SICK_PRICE_ITEM values ('" & str流水号 & "'," & lng序号 & ",'" & _
'            Trim(rsTemp!ITEM_TYPE) & "','" & Trim(rsTemp!ITEM_CODE) & "','" & ToVarchar(Trim(rsTemp!ITEM_NAME), 40) & "','" & _
'            ToVarchar(Trim(rsTemp!ITEM_SPEC), 50) & "','" & ToVarchar(Trim(rsTemp!PRICE_UNIT), 8) & "','" & _
'            Trim(rsTemp!CUnit) & "'," & rs明细!单价 & "," & rs明细!数量 & "," & rs明细!实收金额 & ",'" & _
'            str收据项目 & "','" & trim(rstemp!ITEM_CODE) & "','" & userinfo.姓名 & "',to_Date('" & _
'            format(zldatabase.Currentdate,"yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),0,'" & _
'
'        lng序号 = lng序号 + 1
'        rs明细.MoveNext
'    Wend
'    On Error GoTo errHandle
'
'    '等待返回结算数据
'    If frm等待返回北京尚洋.waitReturn(mstrSavePath & "\SM" & str流水号) = False Then
'        MsgBox "预结算被中止", vbInformation, gstrSysName
'        Unload frm等待返回北京尚洋
'        Exit Function
'    End If
'    Unload frm等待返回北京尚洋
'
'    '返回结算结果
'    strSql = "Select * From " & mstrSavePath & "\SM" & str流水号
'    Set rsTemp = gcn尚洋.Execute(strSql)
'    mcur个帐支付 = Val(rsTemp!JkAccR)
'    mcur统筹金额 = Val(rsTemp!JkSocialR)
'    str结算方式 = "个人帐户;" & Val(rsTemp!JkAccR) & ";0"
'    str结算方式 = str结算方式 & "|统筹记帐;" & Val(rsTemp!JkSocialR) & ";0"
'    门诊虚拟结算_北京尚洋 = True
'    Exit Function
'
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Public Function 门诊结算_北京尚洋(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency, Optional ByRef strAdvance As String) As Boolean
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency
    Dim str流水号 As String, lng病人ID As Long, datCurr As Date, strSQL As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset, lng序号 As Long, str执行部门 As String, str开单部门 As String
    Dim strCardNO As String, str收据项目 As String, str科室类别 As String, str会计类别 As String
    Dim str出院带药 As String, cur基本统筹 As Currency, cur大病统筹 As Currency, rs明细 As New ADODB.Recordset
    Dim cur公务员补助 As Currency, cur补充医疗 As Currency, str结算方式 As String, lng年龄 As Integer
    Dim strTempID As String
    Dim blnOld As Boolean
    Dim strItemType As String, strItemCode As String, strItemName As String, strItemSpec As String, strPriceUnit As String
    Dim rsPati As New ADODB.Recordset
    Dim Guanwei As String
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select * From 病人费用记录 Where 记录状态<>0 And Nvl(实收金额,0)<>0 and Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 And 结帐id=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng病人ID = rs明细!病人ID
    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_北京尚洋)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "没有找到病人信息或医保选择错误"
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    '生成流水号
'    str流水号 = toHex(Format(datCurr, "YYMMDDHHMMSS") & Format(lng病人ID, "0######"), 35)
    str流水号 = Format(zlDatabase.Currentdate(), "yyyy") & rs明细!NO
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称,B.说明,B.类别 from (select * from 保险支付项目 where 险类=[1]) A, 收费细目 B where A.收费细目id(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋, CLng(rs明细!收费细目ID))
'        If IsNull(rsTemp!项目编码) Then
'            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
'            Exit Function
'        End If
        If rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7" Then
        
        Else
            If IsNull(rsTemp!说明) Then
                Err.Raise 9000, gstrSysName, "不能确定项目<" & rsTemp!名称 & ">的收据项目类别和科室核算类别"
                Exit Function
            ElseIf Len(rsTemp!说明) < 2 Then
                Err.Raise 9000, gstrSysName, "不能确定项目<" & rsTemp!名称 & ">的科室核算类别"
                Exit Function
            End If
        End If
'        strTemp = rsTemp!项目编码
'        strSql = "Select * From PARA_CAPTURE_ITEM Where AREAID='" & gstr医保机构编码 & "' And ITEM_CODE='" & UCase(rsTemp!项目编码) & "'"
'        Set rsTemp = gcn尚洋.Execute(strSql)
'        If rsTemp.EOF Then
'            MsgBox "在中间数据中未找到编码为[" & UCase(strTemp) & "]的项目，请核查", vbInformation, gstrSysName
'            Exit Function
'        End If
        rs明细.MoveNext
    Loop
    
    '传费用明细
    lng序号 = 1
    rs明细.MoveFirst
    While Not rs明细.EOF
        gstrSQL = "Select * From 部门表 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!执行部门id))
        str执行部门 = rsTemp!名称
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!开单部门ID))
        str开单部门 = rsTemp!名称
        
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        If IsNull(rsTemp!说明) Then
            If rsTemp!类别 = "5" Then
                strTemp = "AA"
            ElseIf rsTemp!类别 = "6" Then
                strTemp = "BB"
            ElseIf rsTemp!类别 = "7" Then
                strTemp = "CC"
            End If
        Else
            strTemp = rsTemp!说明      '说明不能为空，其中第一位存放收据项目类别，第二位存放科室核算类别
        End If
        str收据项目 = Left(strTemp, 1)
        str科室类别 = Mid(strTemp, 2, 1)
        If rsTemp!类别 = 5 Or rsTemp!类别 = 6 Or rsTemp!类别 = 7 Then
            str会计类别 = "A"       '药品
        Else
            str会计类别 = "B"       '医疗
        End If
        Select Case rsTemp!类别
            Case "1"
                strItemType = "O"
            Case "4"
                strItemType = "U"
            Case "5"
                strItemType = "A"
            Case "6"
                strItemType = "C"
            Case "7"
                strItemType = "C"
            Case "C"
                strItemType = "D"
            Case "D"
                strItemType = "E"
            Case "E", "L"
                strItemType = "F"
            Case "F"
                strItemType = "G"
            Case "G"
                strItemType = "H"
            Case "H"
                strItemType = "I"
            Case "I", "Z"
                strItemType = "Z"
            Case "J"
                strItemType = "J"
            Case "K"
                strItemType = "L"
            Case "M"
                strItemType = "K"
        End Select
        str出院带药 = "在院用药"           '出院带药标志（取值还不清楚，有待询问）
        
        If gint是否职工 = 0 Then
        gstrSQL = "Select 项目编码,收费细目ID From 保险支付项目 Where 险类=[1] And 收费细目id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋, CLng(rs明细!收费细目ID))           '因为之前检查了是否进行对码，所以读出的记录一定不会空
        If rsTemp.EOF Then
            strTempID = ""
        ElseIf IsNull(rsTemp!项目编码) Then
            strTempID = ""
        Else
            strTempID = rsTemp!项目编码
        End If
         
        
        '柴明磊修改：收费类别以前取得C表名称改为取编码了
        Else
        gstrSQL = "Select a.项目编码,a.收费细目ID,substr(b.附注,3,1) as 剂型,c.编码 as 收费类别 From 保险支付大类 C,保险支付项目 a ,保险项目 b" & _
          " where a.大类ID=c.id and a.险类=b.险类 and a.项目编码=b.编码 and a.险类=[1] And a.收费细目id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋, CLng(rs明细!收费细目ID))           '因为之前检查了是否进行对码，所以读出的记录一定不会空
        If rsTemp.EOF Then
            strTempID = ""
            'str剂型 = ""
        Else
            strTempID = Nvl(rsTemp!项目编码)
            'str剂型 = Nvl(rsTemp!剂型)
            str收据项目 = Nvl(rsTemp!收费类别)
        End If
        End If
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        strItemName = rsTemp!名称
        strItemCode = rsTemp!编码
        strPriceUnit = Nvl(rsTemp!计算单位)
        strItemSpec = Nvl(rsTemp!规格)
        
        
        
        '韦华荣修改于2011-3-2
        '居民医保病人明细写入使用原方式，职工医保病人明细写入新的中间表
        If gint是否职工 = 0 Then
            '向中间表写入数据,住院/门诊标志有待询问
            gcn尚洋.Execute "Insert Into SICK_PRICE_ITEM " & _
                " (VISIT_NUMBER,ITEM_NO,ITEM_CLASS,ITEM_CODE,ITEM_NAME,SPEC,PRICE_UNIT,PRICE,QUANTITY, " & _
                "  COST,RECEIPT_CLASS,COLLATE_RELATION,OPERATOR,OPERATE_TIME,CLINIC_FLAG,EXE_DEPT,APP_DOCTOR, " & _
                "  APP_DEPT,TAKE_MEDICINE_FLAG,ITEM_NO_DEPT_STAT,ITEM_NO_ACCOUNTANT_ITEM)" & _
                " values ('" & str流水号 & "'," & lng序号 & ",'" & _
                strItemType & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "','" & _
                ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "'," & _
                Format(rs明细!实收金额 / (rs明细!付数 * rs明细!数次), "0.####") & "," & Format(rs明细!付数 * rs明细!数次, "0.####") & "," & Format(rs明细!实收金额, "0.####") & ",'" & _
                str收据项目 & "','" & strTempID & "','" & rs明细!操作员姓名 & "','" & _
                Format(rs明细!登记时间, "yyyy-MM-dd HH:mm:ss") & "',1,'" & _
                str执行部门 & "','" & rs明细!开单人 & "','" & str开单部门 & "','" & str出院带药 & "','" & _
                str科室类别 & "','" & str会计类别 & "')"
                
        Else
            '职工医保明细写入新的中间表strItemType
        
           gcn尚洋.Execute "Insert Into KC28 " & _
                        " (AKB020,AKC220,CKC158,AAE011,AAE036,AKA063,AKC222,AKC223,AKC227,CKC197,CKC198," & _
                        "  CKC159,CKC160,AKA070,CKC161,CKC169,CKC170,CKC171,CKE081,CKE085,CKE086,CKE090)" & _
                        " values ('" & gstr医保机构编码 & "','" & str流水号 & "','" & lng序号 & "','" & rs明细!操作员姓名 & "'," & _
                        " to_date('" & rs明细!登记时间 & "','yyyy-MM-DD hh24:MI:SS'),'" & str收据项目 & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "'," & _
                        Format(Format(rs明细!标准单价, "0.####") * (rs明细!付数 * rs明细!数次), "0.####") & "," & Format(rs明细!标准单价, "0.####") & "," & Format(rs明细!付数 * rs明细!数次, "0.####") & ",'" & _
                        ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "','/','" & strTempID & "','" & _
                        str执行部门 & "','" & rs明细!开单人 & "','" & str开单部门 & "','" & str收据项目 & "','" & str科室类别 & "','" & str会计类别 & "','" & str出院带药 & "')"
        
         
        End If

        lng序号 = lng序号 + 1
        rs明细.MoveNext
    Wend
    On Error GoTo errHandle
    
    '等待返回结算数据
    strTemp = frm等待返回北京尚洋.waitReturn(str流水号, 0)
    If strTemp = "" Then
        Err.Raise 9000, gstrSysName, "结算过程被中止"
        gcn尚洋.Execute "Delete From SICK_PRICE_ITEM Where VISIT_NUMBER='" & str流水号 & "'"
        Unload frm等待返回北京尚洋
        Exit Function
    End If
    Unload frm等待返回北京尚洋
    
    '返回结算结果
    strSQL = "Select * From MED_RECEIPT_RECORD_MASTER Where CHARGE_NUMBER='" & strTemp & "'"
    Set rsTemp = gcn尚洋.Execute(strSQL)
    If IsDate(rsTemp!BIRTH_DATE) Then
        lng年龄 = Int(zlDatabase.Currentdate() - CDate(rsTemp!BIRTH_DATE)) / 365
    End If
    
    '先更新病人信息
    #If gverControl < 6 Then
        gstrSQL = "Select * From 病人信息 A Where A.病人ID =[1]"
    #Else
        gstrSQL = "Select A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.卡验证码, A.费别, A.医疗付款方式, A.姓名, A.性别, A.年龄, A.出生日期, A.出生地点, A.身份证号, A.其他证件, A.身份, A.职业, A.民族, A.国籍, A.区域, A.学历, A.婚姻状况, A.家庭地址," & vbNewLine & _
            "      A.家庭电话, A.家庭地址邮编 As 户口邮编, A.监护人, A.联系人姓名, A.联系人关系, A.联系人地址, A.联系人电话, A.合同单位id, A.工作单位, A.单位电话, A.单位邮编, A.单位开户行, A.单位帐号, A.担保人, A.担保额, A.担保性质, A.就诊时间, A.就诊状态," & vbNewLine & _
            "      A.就诊诊室, A.住院次数, A.当前科室id, A.当前病区id, A.当前床号, A.入院时间, A.出院时间, A.在院, A.Ic卡号, A.健康号, A.医保号, A.险类, A.查询密码, A.登记时间, A.停用时间, A.锁定" & vbNewLine & _
            "From 病人信息 A Where A.病人ID =[1]"
    #End If
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人信息", lng病人ID)
    gstrSQL = "zl_病人信息_Update(" & _
        lng病人ID & "," & IIf(IsNull(rsPati!门诊号), "NULL", rsPati!门诊号) & "," & _
        IIf(IsNull(rsPati!住院号), "NULL", rsPati!住院号) & ",'" & IIf(IsNull(rsPati!费别), "", rsPati!费别) & "'," & _
        "'" & IIf(IsNull(rsPati!医疗付款方式), "", rsPati!医疗付款方式) & "'," & _
        "'" & rsTemp!Name & "','" & IIf(Nvl(rsTemp!Sex, "0") = "1", "男", "女") & "'," & lng年龄 & "," & _
        "     To_Date('" & Format(rsTemp!BIRTH_DATE, "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
        "'" & IIf(IsNull(rsPati!出生地点), "", rsPati!出生地点) & "','" & rsTemp!PERSONAL_NUMBER & "'," & _
        "'" & IIf(IsNull(rsPati!身份), "", rsPati!身份) & "','" & IIf(IsNull(rsPati!职业), "", rsPati!职业) & "'," & _
        "'" & IIf(IsNull(rsPati!民族), "", rsPati!民族) & "','" & IIf(IsNull(rsPati!国籍), "", rsPati!国籍) & "'," & _
        "'" & IIf(IsNull(rsPati!学历), "", rsPati!学历) & "','" & IIf(IsNull(rsPati!婚姻状况), "", rsPati!婚姻状况) & "'," & _
        "'" & IIf(IsNull(rsPati!家庭地址), "", rsPati!家庭地址) & "','" & IIf(IsNull(rsPati!家庭电话), "", rsPati!家庭电话) & "'," & _
        "'" & IIf(IsNull(rsPati!户口邮编), "", rsPati!户口邮编) & "','" & IIf(IsNull(rsPati!联系人姓名), "", rsPati!联系人姓名) & "'," & _
        "'" & IIf(IsNull(rsPati!联系人关系), "", rsPati!联系人关系) & "','" & IIf(IsNull(rsPati!联系人地址), "", rsPati!联系人地址) & "'," & _
        "'" & IIf(IsNull(rsPati!联系人电话), "", rsPati!联系人电话) & "'," & IIf(IsNull(rsPati!合同单位ID), "NULL", rsPati!合同单位ID) & "," & _
        "'" & Nvl(rsPati!工作单位) & "','" & IIf(IsNull(rsPati!单位电话), "", rsPati!单位电话) & "'," & _
        "'" & IIf(IsNull(rsPati!单位邮编), "", rsPati!单位邮编) & "','" & IIf(IsNull(rsPati!单位开户行), "", rsPati!单位开户行) & "'," & _
        "'" & IIf(IsNull(rsPati!单位帐号), "", rsPati!单位帐号) & "','" & IIf(IsNull(rsPati!担保人), "", rsPati!担保人) & "'," & _
        " " & IIf(IsNull(rsPati!担保额), "NULL", rsPati!担保额) & "," & TYPE_北京尚洋 & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
'        MsgBox "修改病人信息：" & gstrSQL
    
    '获取各种医保结算方式的支付金额
   
     cur个人帐户 = rsTemp!PAY_SIDE2
    cur基本统筹 = rsTemp!PAY_SIDE3
    cur大病统筹 = rsTemp!PAY_SIDE4
    cur补充医疗 = rsTemp!PAY_SIDE5
    cur公务员补助 = rsTemp!PAY_SIDE6
    '写结算结果
  
    If cur个人帐户 <> 0 Then
        str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
    End If
    If cur基本统筹 <> 0 Then
        str结算方式 = str结算方式 & "||基本基金|" & cur基本统筹
    End If
    If cur大病统筹 <> 0 Then
        str结算方式 = str结算方式 & "||大病基金|" & cur大病统筹
    End If
    If cur补充医疗 <> 0 Then
        str结算方式 = str结算方式 & "||补充基金|" & cur补充医疗
    End If
    If cur公务员补助 <> 0 Then
        str结算方式 = str结算方式 & "||公务员津贴|" & cur公务员补助
    End If
    
    '如果存在
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        #If gverControl < 2 Then
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
        #Else
            strAdvance = str结算方式
            gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新结算结果")
    Else
        str结算方式 = "个人帐户|0"
        strAdvance = str结算方式
        gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
    End If
    #If gverControl < 2 Then
        blnOld = True
        frm结算信息.ShowME (lng结帐ID)
    #End If
    
    gstrSQL = "Select 病人ID,结帐金额 From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '帐户年度信息
    Call Get帐户信息(TYPE_北京尚洋, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_北京尚洋 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_北京尚洋 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 & ",0,0," & _
        "0," & mcur统筹金额 & ",0,0," & mcur个帐支付 & ",Null,Null,Null,Null" & IIf(blnOld, "", ",1") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    门诊结算_北京尚洋 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_北京尚洋(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, sngArrInfo(20) As Single
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        门诊结算冲销_北京尚洋 = False
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_北京尚洋, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_北京尚洋 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_北京尚洋 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        Nvl(rsTemp("个人帐户支付"), 0) * -1 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    门诊结算冲销_北京尚洋 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院虚拟结算_北京尚洋(rsDetail As ADODB.Recordset, lng病人ID As Long, str医保号 As String) As String
'因为北京尚洋未提供预结算接口，因此所得到的结算数据为医保结算的正式数据，即得到数据时医保已正式结算
    Dim str流水号 As String, datCurr As Date, strSQL As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset, lng序号 As Long, str执行部门 As String, str开单部门 As String
    Dim strCardNO As String, str收据项目 As String, str科室类别 As String, str会计类别 As String
    Dim str出院带药 As String, cur基本统筹 As Currency, cur大病统筹 As Currency, rs明细 As New ADODB.Recordset
    Dim cur公务员补助 As Currency, cur补充医疗 As Currency, str结算方式 As String, cur个人帐户 As Currency
    Dim strTempID As String
    Dim strItemType As String, strItemCode As String, strItemName As String, strItemSpec As String, strPriceUnit As String
    Dim str住院号 As String, str剂型 As String
    
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select Max(主页ID) From 住院费用记录 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    gstrSQL = " Select * From 病人费用记录" & _
              " Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 And Nvl(实收金额,0)<>0 and nvl(数次,0)*nvl(付数,0)<>0 " & _
              " And 病人id=[1] And 主页id=" & rsTemp(0)
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "", lng病人ID)
    If rs明细.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "没有病人费用，不能结算"
        Exit Function
    End If
    
    lng病人ID = rs明细!病人ID
    gstrSQL = "Select 卡号,医保号 From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_北京尚洋)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    str医保号 = Nvl(rsTemp!医保号)
    
     '2011-06-23 读取住院号小范加入
        gstrSQL = "Select * From 病人信息 Where 病人id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
        str住院号 = Nvl(rsTemp!住院号)
    
    '生成流水号
    str流水号 = ToVarchar(Format(datCurr, "YYMMDDHHMMSS") & lng病人ID, 18)
    
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称,B.说明,B.编码,B.类别 from (select * from 保险支付项目 where 险类=[1]) A, 收费细目 B where A.收费细目id(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋, CLng(rs明细!收费细目ID))
        If rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7" Then
        
        Else
            If IsNull(rsTemp!说明) Then
                Err.Raise 9000, gstrSysName, "不能确定项目<" & rsTemp!名称 & ">的收据项目类别和科室核算类别"
                Exit Function
            ElseIf Len(rsTemp!说明) < 2 Then
                Err.Raise 9000, gstrSysName, "不能确定项目<" & rsTemp!名称 & ">的科室核算类别"
                Exit Function
            End If
        End If
        rs明细.MoveNext
    Loop
    
    '检查是否已建立连接
    If 建立医保连接 = False Then Exit Function
    
    '传费用明细
    lng序号 = 1
    rs明细.MoveFirst
    While Not rs明细.EOF
        
        gstrSQL = "Select * From 部门表 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!执行部门id))
        str执行部门 = rsTemp!名称
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!开单部门ID))
        str开单部门 = rsTemp!名称
        
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        If IsNull(rsTemp!说明) Then
            If rsTemp!类别 = "5" Then
                strTemp = "AA"
            ElseIf rsTemp!类别 = "6" Then
                strTemp = "BB"
            ElseIf rsTemp!类别 = "7" Then
                strTemp = "CC"
            End If
        Else
            strTemp = rsTemp!说明      '说明不能为空，其中第一位存放收据项目类别，第二位存放科室核算类别
        End If
        str收据项目 = Left(strTemp, 1)
        str科室类别 = Mid(strTemp, 2, 1)
        
        If rsTemp!类别 = 5 Or rsTemp!类别 = 6 Or rsTemp!类别 = 7 Then
            str会计类别 = "A"       '药品
        Else
            str会计类别 = "B"       '医疗
        End If
        Select Case rsTemp!类别
            Case "1"
                strItemType = "O"
            Case "4"
                strItemType = "U"
            Case "5"
                strItemType = "A"
            Case "6"
                strItemType = "C"
            Case "7"
                strItemType = "C"
            Case "C"
                strItemType = "D"
            Case "D"
                strItemType = "E"
            Case "E", "L"
                strItemType = "F"
            Case "F"
                strItemType = "G"
            Case "G"
                strItemType = "H"
            Case "H"
                strItemType = "I"
            Case "J"
                strItemType = "J"
            Case "K"
                strItemType = "L"
            Case "M"
                strItemType = "K"
            Case Else
                strItemType = "Z"
        End Select
        gstrSQL = "Select 扣率 From 药品收发记录 Where 费用ID=[1] And NO=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!ID), CStr(rs明细!NO))
        If rsTemp.EOF Then
            str出院带药 = "在院用药"
        Else
            If Mid(CStr(Nvl(rsTemp(0), 0)), 2, 1) = "3" Then
                str出院带药 = "出院带药"
            Else
                str出院带药 = "在院用药"
            End If
        End If
        
        If gint是否职工 = 0 Then
            gstrSQL = "Select 项目编码,收费细目ID From 保险支付项目 Where 险类=[1] And 收费细目id=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋, CLng(rs明细!收费细目ID))          '因为之前检查了是否进行对码，所以读出的记录一定不会空
            If rsTemp.EOF Then
                strTempID = ""
            ElseIf IsNull(rsTemp!项目编码) Then
                strTempID = ""
            Else
                strTempID = rsTemp!项目编码
            End If
            '柴明磊修改：收费类别以前取得C表名称改为取编码了
        Else
            gstrSQL = "Select a.项目编码,a.收费细目ID,substr(b.附注,3,1) as 剂型,c.编码 as 收费类别 From 保险支付大类 C,保险支付项目 a ,保险项目 b" & _
              " where a.大类ID=c.id and a.险类=b.险类 and a.项目编码=b.编码 and a.险类=[1] And a.收费细目id=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋, CLng(rs明细!收费细目ID))            '因为之前检查了是否进行对码，所以读出的记录一定不会空
            If rsTemp.EOF Then
                strTempID = ""
                str剂型 = ""
            Else
                strTempID = Nvl(rsTemp!项目编码)
                str剂型 = Nvl(rsTemp!剂型)
                str收据项目 = Nvl(rsTemp!收费类别)
            End If
        End If
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        strItemName = rsTemp!名称
        strItemCode = rsTemp!编码
        strPriceUnit = Nvl(rsTemp!计算单位)
        strItemSpec = Nvl(rsTemp!规格)
        '---------------韦华荣修改于2011-3-2----------------------------
        '居民医保病人明细写入使用原方式，职工医保病人明细写入新的中间表
        '柴明磊修改：加了BAL_DATE is null判断，因为多次住院，住院号不同
        
        If gint是否职工 = 0 Then
            
            '向中间表写入数据,住院/门诊标志有待询问
            gcn尚洋.Execute "Insert Into SICK_PRICE_ITEM " & _
                " (VISIT_NUMBER,ITEM_NO,ITEM_CLASS,ITEM_CODE,ITEM_NAME,SPEC,PRICE_UNIT,PRICE,QUANTITY, " & _
                "  COST,RECEIPT_CLASS,COLLATE_RELATION,OPERATOR,OPERATE_TIME,CLINIC_FLAG,EXE_DEPT,APP_DOCTOR, " & _
                "  APP_DEPT,TAKE_MEDICINE_FLAG,ITEM_NO_DEPT_STAT,ITEM_NO_ACCOUNTANT_ITEM)" & _
                " values ('" & str流水号 & "'," & lng序号 & ",'" & _
                strItemType & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "','" & _
                ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "'," & _
                Format(rs明细!实收金额 / (rs明细!付数 * rs明细!数次), "0.####") & "," & Format(rs明细!付数 * rs明细!数次, "0.####") & "," & Format(rs明细!实收金额, "0.####") & ",'" & _
                str收据项目 & "','" & strTempID & "','" & rs明细!操作员姓名 & "','" & _
                Format(rs明细!发生时间, "yyyy-MM-dd HH:mm:ss") & "',1,'" & _
                str执行部门 & "','" & rs明细!开单人 & "','" & str开单部门 & "','" & str出院带药 & "','" & _
                str科室类别 & "','" & str会计类别 & "')"
                
        Else
            '职工医保明细写入新的中间表,
            '柴明磊修改：1，把AKAO63取得strItemType改为了str收据项,2，由于医保限制计算单价乘数量是否等于合计所以修改了通过单价保留4位小数乘数量得金额传给医保 & Format(rs明细!实收金额 / (rs明细!付数 * rs明细!数次), "0.####") &
           ' Set rsTemp = gcn尚洋.Execute("Select * From SICK_VISIT_INFO Where PERSONAL_NUMBER='" & str医保号 & "'  and BAL_DATE is null And HOSPITAL_NUMBER='" & gstr医院编码 & "'")
          'xiaofan 修改根据住院号提取医保病人信息
          Set rsTemp = gcn尚洋.Execute("Select * From SICK_VISIT_INFO Where  HOSPITAL_NUMBER='" & gstr医院编码 & "' and RESIDENCE_NO='" & str住院号 & "'")
        
            If rsTemp.EOF Then
                MsgBox "请确认该病人是否已在医保系统中做入院登记!", vbInformation, "医保接口"
                Exit Function
            End If
            
            str住院号 = Nvl(rsTemp!RESIDENCE_NO)
        
            gcn尚洋.Execute "Insert Into KC27 " & _
                            " (AKB020,CKC179,AKC190,AKC220,CKC158,AAE011,AAE036,AKA063,AKC222,AKC223,AKC227,CKC197,CKC198," & _
                            "  CKC159,CKC160,AKA070,CKC161,CKC169,CKC170,CKC171,CKE081,CKE085,CKE086,CKE090)" & _
                            " values ('" & gstr医保机构编码 & "','" & str住院号 & "','" & Null & "','" & str流水号 & "'," & lng序号 & ",'" & rs明细!操作员姓名 & "'," & _
                            " to_date('" & rs明细!登记时间 & "','yyyy-MM-DD hh24:MI:SS'),'" & str收据项目 & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "'," & _
                            Format(Format(rs明细!标准单价, "0.####") * (rs明细!付数 * rs明细!数次), "0.####") & "," & Format(rs明细!标准单价, "0.####") & "," & Format(rs明细!付数 * rs明细!数次, "0.####") & ",'" & _
                            ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "','" & str剂型 & "','" & strTempID & "','" & _
                            str执行部门 & "','" & rs明细!开单人 & "','" & str开单部门 & "','" & str收据项目 & "','" & str科室类别 & "','" & str会计类别 & "','" & str出院带药 & "')"
            
            
            gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细!ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            
        End If
        
        
        lng序号 = lng序号 + 1
        
        
        rs明细.MoveNext
    Wend
    On Error GoTo errHandle
    
    
    
    If gint是否职工 = 1 Then
        '职工医保以住院号提取医保结算返回值
        str流水号 = str住院号
    End If
    
    
    
    '---------------2011-3-2修改部分---------------------------------VISIT_NUMBER
    
    '等待返回结算数据
    '柴明磊修改：发现SICK_PRICE_ITEM表VISIT_NUMBER字段为空，回传取不上数，就改成RESIDENCE_NO
    Screen.MousePointer = 0
    strTemp = frm等待返回北京尚洋.waitReturn(str流水号, 1)
    If gint是否职工 = 0 Then
        If strTemp = "" Then
            Err.Raise 9000, gstrSysName, "结算过程被中止", vbInformation, gstrSysName
            gcn尚洋.Execute "Delete From SICK_PRICE_ITEM Where VISIT_NUMBER='" & str流水号 & "'"
            Unload frm等待返回北京尚洋
            Exit Function
        End If
    Else
        If strTemp = "" Then
            Err.Raise 9000, gstrSysName, "结算过程被中止", vbInformation, gstrSysName
            'XieRong 删除职工明细表为KC227
            gcn尚洋.Execute "Delete From KC27 Where  AKB020='" & gstr医保机构编码 & "' And AKC220='" & str流水号 & "'"
            Unload frm等待返回北京尚洋
            Exit Function
        End If
    End If
    Unload frm等待返回北京尚洋
    
    '返回结算结果
    strSQL = "Select * From MED_RECEIPT_RECORD_MASTER Where CHARGE_NUMBER='" & strTemp & "'"
    Set rsTemp = gcn尚洋.Execute(strSQL)
    
    mcur个帐支付 = rsTemp!PAY_SIDE2
    mcur统筹金额 = rsTemp!PAY_SIDE3 + rsTemp!PAY_SIDE4 + rsTemp!PAY_SIDE5 + rsTemp!PAY_SIDE6

    cur个人帐户 = rsTemp!PAY_SIDE2
    cur基本统筹 = rsTemp!PAY_SIDE3
    cur大病统筹 = rsTemp!PAY_SIDE4
    cur补充医疗 = rsTemp!PAY_SIDE5
    cur公务员补助 = rsTemp!PAY_SIDE6
    '写结算结果
    If cur个人帐户 <> 0 Then
        str结算方式 = str结算方式 & "|个人帐户;" & cur个人帐户 & ";0"
    End If
    If cur基本统筹 <> 0 Then
        str结算方式 = str结算方式 & "|基本基金;" & cur基本统筹 & ";0"
    End If
    If cur大病统筹 <> 0 Then
        str结算方式 = str结算方式 & "|大病基金;" & cur大病统筹 & ";0"
    End If
    If cur补充医疗 <> 0 Then
        str结算方式 = str结算方式 & "|补充基金;" & cur补充医疗 & ";0"
    End If
    If cur公务员补助 <> 0 Then
        str结算方式 = str结算方式 & "|公务员津贴;" & cur公务员补助 & ";0"
    End If
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 2)
    Else
        str结算方式 = "个人帐户;" & cur个人帐户 & ";0"
    End If
    住院虚拟结算_北京尚洋 = str结算方式
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    住院虚拟结算_北京尚洋 = ""
End Function

Public Function 住院结算_北京尚洋(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency
    Dim datCurr As Date
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 住院费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    Do Until rsTemp.EOF
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '帐户年度信息
    Call Get帐户信息(TYPE_北京尚洋, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_北京尚洋 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_北京尚洋 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 & ",0,0," & _
        "0," & mcur统筹金额 & ",0,0," & mcur个帐支付 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    住院结算_北京尚洋 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_北京尚洋(lng结帐ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, sngArrInfo(20) As Single
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, lng病人ID As Long
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    lng冲销ID = rsTemp("ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_北京尚洋, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        住院结算冲销_北京尚洋 = False
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_北京尚洋, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_北京尚洋 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_北京尚洋 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        Nvl(rsTemp("个人帐户支付"), 0) * -1 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    住院结算冲销_北京尚洋 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 出院登记_北京尚洋(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHandle
    '对HIS之中的基础数据进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_北京尚洋 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_北京尚洋 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_北京尚洋 = False
End Function

Public Function 入院登记_北京尚洋(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHandle
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_北京尚洋 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_北京尚洋 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_北京尚洋 = False
End Function

Private Function toHex(ByVal dblNum As Double, Optional ByVal dblKey As Double = 16) As String
    Dim dblTemp As Double, dblMod As Double, strTemp As String
    dblTemp = dblNum
    Do
        dblMod = dblTemp - Int(dblTemp / dblKey) * dblKey
        dblTemp = Int(dblTemp / dblKey)
        If dblMod >= 10 Then
            strTemp = Chr(dblMod + 55) & strTemp
        Else
            strTemp = dblMod & strTemp
        End If
    Loop While dblTemp >= dblKey
    dblMod = dblTemp
    If dblMod >= 10 Then
        strTemp = Chr(dblMod + 55) & strTemp
    Else
        strTemp = dblMod & strTemp
    End If
    toHex = strTemp
End Function

Public Sub WriteInfo(ByVal strInfo As String)
    Dim strFileName As String
    Dim objSystem As FileSystemObject
    Dim objStream As TextStream
    
    strFileName = "C:\信息" & Format(Date, "MMdd") & ".txt"
    Set objSystem = New FileSystemObject
    If Not objSystem.FileExists(strFileName) Then Call objSystem.CreateTextFile(strFileName, False)
    Set objStream = objSystem.OpenTextFile(strFileName, ForAppending, False, TristateMixed)
    objStream.WriteLine (strInfo)
    objStream.Close
End Sub


