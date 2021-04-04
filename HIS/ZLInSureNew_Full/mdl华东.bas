Attribute VB_Name = "mdl华东"
Option Explicit
'Modified By 朱玉宝 2005-07-21 10:19:16；修改原因：1、床位费超过限额，超出部分对应自费码，并新产生一条记录上传；2、费用明细按费用发生时间排序上传
'Modified By 朱玉宝 2005-07-25 修改原因：1、以IC卡号做为门诊文件名，住院号做为住院文件名；2、住院记帐时不实时上传明细；3、病人费用查询调预结算时，不必读取医保返回文件

Private mcur统筹金额 As Currency, mcur个帐支付 As Currency
Public gcn华东 As New ADODB.Connection, mstrSavePath As String

Public Const MAX_PATH = 260

Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        '设置浏览窗口
        .hwndOwner = lWindowHwnd
        '返回选中的目录
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "请选定开始搜索的文件夹："
        Else
            .lpszTitle = sTitle
        End If
    End With
    
    '调出浏览窗口
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '获取路径
        SHGetPathFromIDList lpIDList, sPath
        '释放内存
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
End Function


Public Function 医保初始化_华东() As Boolean
'功能：测试是否可以连接到前置服务器上
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs华东 As New ADODB.Recordset
    '如果连接已经打开，那就不用再测试
    If gcn华东.State = adStateOpen Then
        医保初始化_华东 = True
        Exit Function
    End If
     
    On Error GoTo ErrH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_华东)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        If rsTemp!参数名 = "文件存放位置" Then mstrSavePath = rsTemp!参数值
        rsTemp.MoveNext
    Loop
    If Trim(mstrSavePath) = "" Then
        MsgBox "请到医保参数设置中设置文件存放位置", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error Resume Next
    gcn华东.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=Visual FoxPro Tables;UID=;SourceDB=" & mstrSavePath & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"""
    gcn华东.CursorLocation = adUseClient
    gcn华东.Open
    
    If Err <> 0 Then
        MsgBox "文件存放位置指定错误。", vbInformation, gstrSysName
        医保初始化_华东 = False
        Exit Function
    End If
    医保初始化_华东 = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    医保初始化_华东 = False
End Function

Public Function 医保设置_华东() As Boolean
    医保设置_华东 = frmSet华东.ShowME(TYPE_华东)
End Function

Public Function 个人余额_华东(lng病人ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_华东)
    个人余额_华东 = Nvl(rsTemp!帐户余额, 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function 身份标识_华东(Optional bytType As Byte = 0, Optional lng病人ID As Long = 0) As String
    '华东医保没提供专门的身分验证接口，通过调取挂号单号来实现验证
    Dim strTemp As String
    strTemp = frmIdentify华东.Identify(bytType, lng病人ID)
    Unload frmIdentify华东
    If strTemp = "" Then
        MsgBox "未提取病人信息", vbInformation, gstrSysName
    Else
        身份标识_华东 = strTemp
    End If
End Function

Public Function 门诊虚拟结算_华东(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'因为华东未提供预结算接口，因此所得到的结算数据为医保结算的正式数据，即得到数据时医保已正式结算
    Dim str流水号 As String, lng病人ID As Long, datCurr As Date, strSQL As String
    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng序号 As Long
    Dim strCardNO As String
'    病人ID         adBigInt, 19, adFldIsNullable
'    收费类别       adVarChar, 2, adFldIsNullable
'    收据费目       adVarChar, 20, adFldIsNullable
'    计算单位       adVarChar, 6, adFldIsNullable
'    开单人         adVarChar, 20, adFldIsNullable
'    收费细目ID     adBigInt, 19, adFldIsNullable
'    数量           adSingle, 15, adFldIsNullable
'    单价           adSingle, 15, adFldIsNullable
'    实收金额       adSingle, 15, adFldIsNullable
'    统筹金额       adSingle, 15, adFldIsNullable
'    保险支付大类ID adBigInt, 19, adFldIsNullable
'    是否医保       adBigInt, 19, adFldIsNullable
'    摘要           adVarChar, 200, adFldIsNullable
'    是否急诊       adBigInt, 19, adFldIsNullable
'    str结算方式  "报销方式;金额;是否允许修改|...."
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    datCurr = zlDatabase.Currentdate
    lng病人ID = rs明细!病人ID
    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_华东)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    '生成流水号
    str流水号 = strCardNO
    
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称 from (select * from 保险支付项目 where 险类=[1]) A, 收费细目 B where A.收费细目id(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_华东, CLng(rs明细!收费细目ID))
        If IsNull((rsTemp!项目编码)) Then
            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
            Exit Function
        End If
        rs明细.MoveNext
    Loop
    
    '生成DBF文件
    On Error Resume Next
    gcn华东.Execute "Drop Table " & mstrSavePath & "\YM" & str流水号
    
    On Error GoTo errHandle
    gcn华东.Execute "Create Table " & mstrSavePath & "\YM" & str流水号 & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
        "SelfJe N(18,4))"
    lng序号 = 1
    rs明细.MoveFirst
    While Not rs明细.EOF
        gstrSQL = "Select A.项目编码,B.ID,B.名称,B.规格,B.计算单位 From 保险支付项目 A,收费细目 B Where B.ID=A.收费细目ID And A.收费细目id=[2] And A.险类=[1]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_华东, CLng(rs明细!收费细目ID))             '因为之前检查了是否进行对码，所以读出的记录一定不会空
        
        '卡号、流水号、序号、内码、名称、科目编码、规格、计量单位、数量、单价、金额、自费金额
        gcn华东.Execute "Insert Into " & mstrSavePath & "\YM" & str流水号 & " values ('" & strCardNO & "','" & str流水号 & "'," & _
            lng序号 & ",'" & Trim(rsTemp!项目编码) & "','" & Trim(rsTemp!名称) & "','','" & Trim(rsTemp!规格) & "','" & Trim(rsTemp!计算单位) & "'," & _
            Format(rs明细!数量, "0.####") & "," & Format(rs明细!单价, "0.####") & "," & Format(rs明细!实收金额, "0.####") & "," & _
            "0)"
        lng序号 = lng序号 + 1
        rs明细.MoveNext
    Wend
    On Error GoTo errHandle
    '等待返回结算数据
    If frm等待返回华东.waitReturn(mstrSavePath & "\SM" & str流水号) = False Then
        MsgBox "预结算被中止", vbInformation, gstrSysName
        On Error Resume Next
        gcn华东.Execute "Drop Table " & mstrSavePath & "\YM" & str流水号
        Unload frm等待返回华东
        Exit Function
    End If
    Unload frm等待返回华东
    
    '返回结算结果
    strSQL = "Select * From " & mstrSavePath & "\SM" & str流水号
    Set rsTemp = gcn华东.Execute(strSQL)
    mcur个帐支付 = Val(rsTemp!JkAccR)
    mcur统筹金额 = Val(rsTemp!JkSocialR)
    str结算方式 = "个人帐户;" & Val(rsTemp!JkAccR) & ";0"
    str结算方式 = str结算方式 & "|统筹记帐;" & Val(rsTemp!JkSocialR) & ";0"
    On Error Resume Next
    gcn华东.Execute "Drop Table " & mstrSavePath & "\YM" & str流水号
    gcn华东.Execute "Drop Table " & mstrSavePath & "\SM" & str流水号
    门诊虚拟结算_华东 = True
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_华东(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency
    Dim datCurr As Date, lng病人ID As Long
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '帐户年度信息
    Call Get帐户信息(TYPE_华东, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_华东 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_华东 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 & ",0,0," & _
        "0," & mcur统筹金额 & ",0,0," & mcur个帐支付 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    门诊结算_华东 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_华东(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
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
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_华东, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        门诊结算冲销_华东 = False
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_华东, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_华东 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_华东 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        Nvl(rsTemp("个人帐户支付"), 0) * -1 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    门诊结算冲销_华东 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院虚拟结算_华东(rs明细 As ADODB.Recordset, lng病人ID As Long, str医保号 As String, Optional ByVal bln查询 As Boolean = False) As String
'因为华东未提供预结算接口，因此所得到的结算数据为医保结算的正式数据，即得到数据时医保已正式结算
    Dim str流水号 As String, datCurr As Date, strSQL As String
    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng序号 As Long
    Dim strCardNO As String
'    病人ID         adBigInt, 19, adFldIsNullable
'    收费类别       adVarChar, 2, adFldIsNullable
'    收据费目       adVarChar, 20, adFldIsNullable
'    计算单位       adVarChar, 6, adFldIsNullable
'    开单人         adVarChar, 20, adFldIsNullable
'    收费细目ID     adBigInt, 19, adFldIsNullable
'    数量           adSingle, 15, adFldIsNullable
'    单价           adSingle, 15, adFldIsNullable
'    实收金额       adSingle, 15, adFldIsNullable
'    统筹金额       adSingle, 15, adFldIsNullable
'    保险支付大类ID adBigInt, 19, adFldIsNullable
'    是否医保       adBigInt, 19, adFldIsNullable
'    摘要           adVarChar, 200, adFldIsNullable
'    是否急诊       adBigInt, 19, adFldIsNullable
'    str结算方式  "报销方式;金额;是否允许修改|...."
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
        Exit Function
    End If

    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select A.卡号,B.住院号 From 保险帐户 A,病人信息 B Where A.病人ID=B.病人ID And A.病人id=[1] And A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_华东)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    str流水号 = rsTemp!住院号

    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称 from (select * from 保险支付项目 where 险类=[1]) A, 收费细目 B where A.收费细目id(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_华东, CLng(rs明细!收费细目ID))
        If IsNull((rsTemp!项目编码)) Then
            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
            Exit Function
        End If
        rs明细.MoveNext
    Loop
    记帐传输_华东 "", 0, "", lng病人ID
    
    If bln查询 Then Exit Function
    On Error GoTo errHandle
    
    '等待返回结算数据
    If frm等待返回华东.waitReturn(mstrSavePath & "\SZ" & str流水号) = False Then
        MsgBox "预结算被中止", vbInformation, gstrSysName
        Unload frm等待返回华东
        Exit Function
    End If
    Unload frm等待返回华东
    
    '返回结算结果
    strSQL = "Select Sum(JkaccR) As JkaccR,Sum(JkSocialR) As JkSocialR From " & mstrSavePath & "\SZ" & str流水号
    Set rsTemp = gcn华东.Execute(strSQL)
    mcur个帐支付 = Val(rsTemp!JkAccR)
    mcur统筹金额 = Val(rsTemp!JkSocialR)
    住院虚拟结算_华东 = "个人帐户;" & Val(rsTemp!JkAccR) & ";0"
    住院虚拟结算_华东 = 住院虚拟结算_华东 & "|统筹记帐;" & Val(rsTemp!JkSocialR) & ";0"
    On Error Resume Next
    gcn华东.Execute "Drop Table " & mstrSavePath & "\SZ" & str流水号
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_华东(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
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
    Call Get帐户信息(TYPE_华东, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_华东 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_华东 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 & ",0,0," & _
        "0," & mcur统筹金额 & ",0,0," & mcur个帐支付 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    住院结算_华东 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_华东(lng结帐ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, sngArrInfo(20) As Single
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, lng病人ID As Long
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 住院费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_华东, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        住院结算冲销_华东 = False
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_华东, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_华东 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_华东 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        Nvl(rsTemp("个人帐户支付"), 0) * -1 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    住院结算冲销_华东 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 出院登记_华东(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHandle
    '对HIS之中的基础数据进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_华东 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_华东 = False
End Function

Public Function 入院登记_华东(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHandle
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_华东 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_华东 = False
End Function

Public Function 记帐传输_华东(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
    Dim rs明细 As New ADODB.Recordset, lng主页ID As Long, rsTemp As New ADODB.Recordset
    Dim str流水号 As String, datCurr As Date, strSQL As String
    Dim strCardNO As String
    Dim int顺序号 As Integer
    
    '以下和床位费相关
    Dim dbl限额 As Double, str自费码 As String
    
    '先读取保险参数中关于床位费的设置
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=[1] And 参数名='床位费限额'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取床位费限额", TYPE_华东)
    If rsTemp.RecordCount <> 0 Then dbl限额 = Nvl(rsTemp!参数值, 0)
    If dbl限额 <> 0 Then
        '取床位费自费码
        gstrSQL = "Select 参数值 From 保险参数 Where 险类=[1] And 参数名='床位费自费码'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取床位费自费码", TYPE_华东)
        If rsTemp.RecordCount <> 0 Then str自费码 = Nvl(rsTemp!参数值)
    End If
    
    If str单据号 <> "" Then
        gstrSQL = "Select 病人id From 住院费用记录 Where NO=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, str单据号)
        lng病人ID = rsTemp(0)
    End If
    gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng主页ID = Nvl(rsTemp(0), 1)
    If str单据号 <> "" Then
        gstrSQL = " Select A.* From 住院费用记录 A,保险帐户 B " & _
                  " Where A.记录状态<>0 And Nvl(A.是否上传,0)=0 And nvl(A.附加标志,0)<>9 " & _
                  " and A.记录性质=" & int性质 & " and A.NO='" & str单据号 & "'" & _
                  " and A.病人ID=B.病人ID And B.险类=" & TYPE_华东 & _
                  " order by A.病人ID,A.发生时间,A.记录性质,A.NO,A.序号"
    Else
        gstrSQL = "Select * From 住院费用记录 Where Nvl(实收金额,0)<>0 And 记录状态<>0 And nvl(附加标志,0)<>9 and 病人id=[1] And 主页id=[2] And NVl(是否上传,0)=0 order by 发生时间,记录性质,NO,序号"
    End If
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "", lng病人ID, lng主页ID)
    
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        记帐传输_华东 = True
        Exit Function
    End If
    
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select A.卡号,A.退休证号,B.住院号 From 保险帐户 A,病人信息 B Where A.病人ID=B.病人ID And A.病人id=[1] And A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_华东)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    '生成流水号
    str流水号 = rsTemp!住院号
    int顺序号 = Nvl(rsTemp!退休证号, 0)
    '医保商的问题，序号至少从2开始
    If int顺序号 = 0 Then int顺序号 = 1
    int顺序号 = int顺序号 + 1
    
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称 from (select * from 保险支付项目 where 险类=[1]) A, 收费细目 B where A.收费细目id(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", TYPE_华东, CLng(rs明细!收费细目ID))
        If IsNull((rsTemp!项目编码)) Then
            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
            Exit Function
        End If
        rs明细.MoveNext
    Loop
    
    '生成DBF文件
    On Error Resume Next
    gcn华东.Execute "Drop Table " & mstrSavePath & "\YZ" & str流水号
    
    On Error GoTo errHandle
    gcn华东.Execute "Create Table " & mstrSavePath & "\YZ" & str流水号 & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
        "SelfJe N(18,4),Bz1 C(25),Bz2 C(5),Bz3 C(5))"
    strSQL = "Select * From " & mstrSavePath & "\YZ" & str流水号
    rs明细.MoveFirst
    While Not rs明细.EOF
        gstrSQL = "Select A.项目编码,B.ID,B.名称,B.规格,B.计算单位 From 保险支付项目 A,收费细目 B Where B.ID=A.收费细目ID And A.收费细目id=[1] And A.险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID), TYPE_华东)            '因为之前检查了是否进行对码，所以读出的记录一定不会空
        
        '卡号、流水号、序号、内码、名称、科目编码、规格、计量单位、数量、单价、金额、自费金额
        If rs明细!收费类别 <> "J" Or dbl限额 = 0 Then
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_华东 & ",'退休证号','" & int顺序号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
            gstrSQL = "zl_病人记帐记录_上传 (" & rs明细!ID & ",0,'" & int顺序号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            gcn华东.Execute "Insert Into " & mstrSavePath & "\YZ" & str流水号 & " values ('" & strCardNO & "','" & str流水号 & "'," & _
                int顺序号 & ",'" & Trim(rsTemp!项目编码) & "','" & Trim(rsTemp!名称) & "','','" & Trim(rsTemp!规格) & "','" & Trim(rsTemp!计算单位) & "'," & _
                Format(rs明细!付数 * rs明细!数次, "0.####") & "," & Format(rs明细!实收金额 / (rs明细!付数 * rs明细!数次), "0.####") & "," & Format(rs明细!实收金额, "0.####") & "," & _
                "0,'" & Format(rs明细!发生时间, "yyyy-MM-dd") & "','2','3')"
        Else
            '如果是床位费，且超过限额，则需要将超过部分以自费码上传(上传为2条明细)
            If Val(Format(rs明细!实收金额 / (rs明细!付数 * rs明细!数次), "#0.00")) > Val(Format(dbl限额, "#0.00")) Then
                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_华东 & ",'退休证号','" & int顺序号 & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
                gstrSQL = "zl_病人记帐记录_上传 (" & rs明细!ID & ",0,'" & int顺序号 & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
                gcn华东.Execute "Insert Into " & mstrSavePath & "\YZ" & str流水号 & " values ('" & strCardNO & "','" & str流水号 & "'," & _
                    int顺序号 & ",'" & Trim(rsTemp!项目编码) & "','" & Trim(rsTemp!名称) & "','','" & Trim(rsTemp!规格) & "','" & Trim(rsTemp!计算单位) & "'," & _
                    Format(rs明细!付数 * rs明细!数次, "0.####") & "," & Format(dbl限额, "0.####") & "," & Format(rs明细!付数 * rs明细!数次 * dbl限额, "0.####") & "," & _
                    "0,'" & Format(rs明细!发生时间, "yyyy-MM-dd") & "','2','3')"
                int顺序号 = int顺序号 + 1
                gcn华东.Execute "Insert Into " & mstrSavePath & "\YZ" & str流水号 & " values ('" & strCardNO & "','" & str流水号 & "'," & _
                    int顺序号 & ",'" & str自费码 & "','" & Trim(rsTemp!名称) & "','','" & Trim(rsTemp!规格) & "','" & Trim(rsTemp!计算单位) & "'," & _
                    Format(rs明细!付数 * rs明细!数次, "0.####") & "," & Format((rs明细!实收金额 - rs明细!付数 * rs明细!数次 * dbl限额) / (rs明细!付数 * rs明细!数次), "0.####") & "," & Format((rs明细!实收金额 - rs明细!付数 * rs明细!数次 * dbl限额), "0.####") & "," & _
                    "0,'" & Format(rs明细!发生时间, "yyyy-MM-dd") & "','2','3')"
            Else
                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_华东 & ",'退休证号','" & int顺序号 & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
                gstrSQL = "zl_病人记帐记录_上传 (" & rs明细!ID & ",0,'" & int顺序号 & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
                gcn华东.Execute "Insert Into " & mstrSavePath & "\YZ" & str流水号 & " values ('" & strCardNO & "','" & str流水号 & "'," & _
                    int顺序号 & ",'" & Trim(rsTemp!项目编码) & "','" & Trim(rsTemp!名称) & "','','" & Trim(rsTemp!规格) & "','" & Trim(rsTemp!计算单位) & "'," & _
                    Format(rs明细!付数 * rs明细!数次, "0.####") & "," & Format(rs明细!实收金额 / (rs明细!付数 * rs明细!数次), "0.####") & "," & Format(rs明细!实收金额, "0.####") & "," & _
                    "0,'" & Format(rs明细!发生时间, "yyyy-MM-dd") & "','2','3')"
            End If
        End If
        
        int顺序号 = int顺序号 + 1
        rs明细.MoveNext
    Wend
    
    记帐传输_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
