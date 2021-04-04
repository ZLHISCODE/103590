Attribute VB_Name = "mdl壁山"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

Public gcn壁山 As New ADODB.Connection
Private mstr门诊号 As String

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOPMOST = -2

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetprivateprofileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, _
    ByVal lpDefault As String, ByVal lpRetrm_String As String, ByVal cbReturnString As Integer, ByVal FileName As String) As Integer

'以下结构体用来纪录虚拟结算结果，用于结算时核对
Private Type typBalance
    cur医保基金 As Double
    cur个人帐户 As Double
    cur大病基金 As Double
End Type
Private pre_Balance As typBalance

Public Function 医保初始化_壁山() As Boolean
'功能：测试是否可以连接到前置服务器上
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs壁山 As New ADODB.Recordset
    '如果连接已经打开，那就不用再测试
    If gcn壁山.State = adStateOpen Then
        医保初始化_壁山 = True
        Exit Function
    End If
     
    On Error GoTo ErrH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", TYPE_重庆壁山)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "壁山服务器"
                strServer = strTemp
            Case "壁山用户名"
                strUser = strTemp
            Case "壁山用户密码"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    If Val(Get保险参数_壁山("适用地区")) = 1 Then
        gcn壁山.Open "Provider=SQLOLEDB.1;Initial Catalog=hw_interface;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
    Else
        gcn壁山.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
            strServer, strUser, strPass
    End If
    If Err <> 0 Then
        MsgBox "医保前置服务器连接失败。", vbInformation, gstrSysName
        医保初始化_壁山 = False
        Exit Function
    End If
    医保初始化_壁山 = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    医保初始化_壁山 = False
End Function

Public Function 身份标识_壁山(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify壁山
    Dim strPatiInfo As String, cur余额 As Currency
    Dim arr, datCurr As Date, str门诊号 As String
    Dim strSQL As String, str特殊病 As String
    Dim strTemp As String, errLine As Integer
    
    '判断是否保存有IC卡验证码
    strTemp = Get保险参数_壁山("卡验证码")
    If strTemp = "" Then
        MsgBox "请在医保参数中设置本地医保的IC卡验证码。", vbInformation, gstrSysName
        Exit Function
    End If
    
    frmIDentified.mstr验证码 = strTemp
    frmIDentified.Tag = bytType
    frmIDentified.Show 1
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    On Error GoTo errHandle
    strPatiInfo = frmIDentified.mstrPatiInfo: errLine = 1
    cur余额 = frmIDentified.mcur余额: errLine = 2
    Unload frmIDentified: errLine = 3
    
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)

        lng病人ID = BuildPatiInfo(bytType, strPatiInfo & ";;;;" & cur余额 & ";;;;;;;" & cur余额 & ";;;;;", lng病人ID, TYPE_重庆壁山): errLine = 4
        '返回格式:中间插入病人ID
        strPatiInfo = strPatiInfo & ";" & lng病人ID & ";;;;" & cur余额 & ";;;;;;;" & cur余额 & ";;;;;": errLine = 5
    Else
        身份标识_壁山 = "": errLine = 6
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Exit Function
    End If
    arr = Split(strPatiInfo, ";"): errLine = 12
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
        '检查是否特殊病
        str特殊病 = frmIDentified.mstr特殊病: errLine = 7
        gstr特殊病种 = str特殊病: errLine = 8
    ElseIf Val(Get保险参数_壁山("适用地区")) = 1 Then           '记录病人是否是特殊病
        If gbln特殊门诊 = True Then
            gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆壁山 & ",'灰度级','''1''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            gbln特殊门诊 = False
            str特殊病 = Get病人ID(CStr(arr(1)), CStr(TYPE_重庆壁山))
        Else
            gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆壁山 & ",'灰度级','''0''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        End If
    Else
        str特殊病 = Get病人ID(CStr(arr(1)), CStr(TYPE_重庆壁山)): errLine = 9
    End If
    If bytType <> 0 Then
        身份标识_壁山 = strPatiInfo: errLine = 10
    End If
    '如果为门诊病人，就接着进行门诊登记
    datCurr = zlDatabase.Currentdate: errLine = 11
    str门诊号 = ToVarchar(lng病人ID & Format(datCurr, "yyddhhmmss"), 16): errLine = 13
    mstr门诊号 = str门诊号: errLine = 14
    '进行门诊登记准备
    If bytType <> 0 Then
        身份标识_壁山 = strPatiInfo
    Else
        strSQL = "insert into Check_doex_interface(Bill_no,App_code" & _
                ",Doct_flag,Doex_no,Ill_type,Ic_id,Is_bala,Regi_op_id) values('" & _
                str门诊号 & "','" & Mid(gstr医院编码, 1, 4) & "','" & IIf(bytType = 1, 1, 0) & "','" & _
                Left(str门诊号, 10) & "','" & str特殊病 & _
                "','" & arr(2) & arr(0) & "','0','" & ToVarchar(UserInfo.姓名, 8) & "')": errLine = 15
        gcn壁山.Execute strSQL: errLine = 16
        '进行门诊登记请求
        strSQL = "insert into Check_bill_request(Bill_no,App_code," & _
                "Request_status) values('" & str门诊号 & "','" & _
                Mid(gstr医院编码, 1, 4) & "','0')": errLine = 17
        gcn壁山.Execute strSQL: errLine = 18
        If Checkrequest(str门诊号) = False Then
            '删除失败的门诊登记单
            strSQL = "delete from Check_bill_request where Bill_no = '" & str门诊号 & _
                    "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": errLine = 19
            gcn壁山.Execute strSQL: errLine = 10
            strSQL = "delete from Check_doex_interface where Bill_no = '" & _
                    str门诊号 & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": errLine = 21
            gcn壁山.Execute strSQL: errLine = 22
            身份标识_壁山 = ""
            Exit Function
        Else
            身份标识_壁山 = strPatiInfo
        End If
    End If
    Exit Function
errHandle:
    MsgBox "错误出现在[身份验证]模块第" & errLine & "行", vbInformation, "错误"
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_壁山 = ""
End Function

Public Function 门诊结算_壁山(lng结帐ID As Long, cur个帐支付 As Currency, str医保号 As String, cur全自付 As Currency, cur先自付 As Currency, cur医保基金 As Currency) As Boolean
'功能：对门诊费用进行明细传递并且进行结算
'如果门诊费用明细传递失败，就直接结束函数，返回函数失败
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rs壁山 As New ADODB.Recordset
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, curDate As Date
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur个人帐户 As Currency, cur起付线 As Currency, cur基本统筹限额 As Currency
    Dim cur大额统筹限额 As Currency, cur基数自付 As Currency, cur余额 As Currency
    Dim cur发生费用 As Currency, cur特病统筹 As Currency, str结算周期 As String
    Dim strInipath As String, str配置文件名 As String
    
    On Error GoTo errHandle
    '如果个人余额不足，无法进行结算
    cur余额 = 个人余额_壁山(Get病人ID(CStr(str医保号), CStr(TYPE_重庆壁山)))
    If cur个帐支付 > cur余额 Then
        Err.Raise 9000, gstrSysName, "需要的费用已经大于剩余费用", vbInformation, gstrSysName
        门诊结算_壁山 = False
        Exit Function
    End If
    If 费用明细传递(1, lng结帐ID) = False Then
        门诊结算_壁山 = False
        Exit Function
    End If
    
    WriteInfo vbCrLf & "开始门诊结算"
    '进行结算准备
    strSQL = "Update Check_doex_interface set Ps_account_pay = " & _
            CStr(cur个帐支付) & ",Bala_op_id = '" & ToVarchar(UserInfo.姓名, 8) & _
            "' where Bill_no = '" & mstr门诊号 & "' and " & _
            "App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    gcn壁山.Execute strSQL
    
    '提交结算请求
    strSQL = "update Check_bill_request set Request_status = '1',Request_Result=null where" & _
            " Bill_no ='" & mstr门诊号 & "' and " & _
            " App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    WriteInfo "发送请求:" & strSQL
    gcn壁山.Execute strSQL
    
    str配置文件名 = mstr门诊号 & ".ini"
    'Modified By 朱玉宝 下午 06:10:58
    If Val(Get保险参数_壁山("适用地区")) = 1 Then
        '发出写卡请求（如果下卡失败，将会在请求表中返回错误，下面一步就会检测出来）
        WriteInfo "下卡:" & mstr门诊号
        Call Shell("D:\hw_ic_write\hw_ic_write.exe " & mstr门诊号, vbHide)
        '读取配置文件
        strInipath = Trim(Get保险参数_壁山("配置文件位置"))
        If strInipath = "" Then strInipath = App.Path
        If Right(strInipath, 1) <> "\" Then strInipath = strInipath & "\"
        '配置文件未赋值，询问之后赋值，或注释上一语句，并在医保设置时输入完整的路径（含文件名）
        strInipath = strInipath & str配置文件名
        
        WriteInfo "配置文件名:" & strInipath
        If Not frm等待响应_黔江.ShowME(strInipath) Then
            WriteInfo "读取文件操作中断"
            Err.Raise 9000, gstrSysName, "等待文件返回操作被中断,请检查中心是否已结算,并对数据进行核对", vbInformation, gstrSysName
            Exit Function
        End If
        
        If GetIniS("Sign", "Sign", "0", strInipath) = "0" Then
            WriteInfo "返回错误:" & GetIniS("Sign", "Error_txt", "", strInipath)
            Err.Raise 9000, gstrSysName, GetIniS("Sign", "Error_txt", "", strInipath), vbInformation, "医保返回"
            Exit Function
        End If
        
        cur个帐支付 = CCur(GetIniS("Recorde", "Ps_account_pay", "0", strInipath))
        cur余额 = CCur(GetIniS("Recorde", "Ps_bala", "0", strInipath))
        WriteInfo "个帐:" & cur个帐支付 & vbCrLf & "余额:" & cur余额
        cur余额 = CCur(GetIniS("Recorde", "Ps_cost_pay", "0", strInipath))
        
    Else
        If Checkrequest(mstr门诊号) = False Then 门诊结算_壁山 = False: Exit Function
        
        '求出结算结果
        curDate = zlDatabase.Currentdate
        '获取个人帐户支付和个人现金支付
        strSQL = "select Ps_account_pay,Ps_cost_pay,Ps_bala,Plan_pay,acc_cyc from Check_doex_interface" & _
                " where Bill_no ='" & mstr门诊号 & "' and " & _
                " App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        If rs壁山.State = adStateOpen Then rs壁山.Close
        rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
        cur个帐支付 = Nvl(rs壁山("Ps_account_pay"), 0)
        cur余额 = Nvl(rs壁山("Ps_bala"), 0)
        cur全自付 = Nvl(rs壁山("Ps_cost_pay"), 0)
        str结算周期 = Nvl(rs壁山("acc_cyc"), "")
    End If
    
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
        cur特病统筹 = Nvl(rs壁山("Plan_pay"), 0)
    Else
        cur特病统筹 = 0
    End If
    cur医保基金 = cur特病统筹
    cur发生费用 = cur全自付 + cur个帐支付 + cur特病统筹
    '帐户年度信息
    Call Get帐户信息(TYPE_重庆壁山, Get病人ID(CStr(str医保号), CStr(TYPE_重庆壁山)), Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & Get病人ID(CStr(str医保号), CStr(TYPE_重庆壁山)) & _
            "," & TYPE_重庆壁山 & "," & Year(curDate) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & _
            cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_重庆壁山 & "," & _
            Get病人ID(CStr(str医保号), CStr(TYPE_重庆壁山)) & "," & Year(curDate) & "," & _
            cur余额 & "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + cur特病统筹 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur发生费用 & "," & cur全自付 & "," & cur先自付 & ",NULL," & cur特病统筹 & ",NULL,NULL," & _
            cur个帐支付 & ",NULL,NULL,NULL,'" & mstr门诊号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
'        gstrSQL = "zl_结算周期记录_insert("
        gstrSQL = "Insert into zlhis.结算周期记录 values (" & lng结帐ID & ",'" & str结算周期 & "'," & cur发生费用 & "," & cur个帐支付 & "," & cur特病统筹 & ",'L',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
        gcnOracle.Execute gstrSQL
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "西铝厂医保")
    End If

    strSQL = "delete from Check_bill_request  where" & _
            " Bill_no ='" & mstr门诊号 & "' and  App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    gcn壁山.Execute strSQL
    门诊结算_壁山 = True
    WriteInfo "结束门诊结算"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    门诊结算_壁山 = False
End Function

Public Function 门诊结算冲销_壁山(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, str结算周期 As String
    Dim cur票据总金额 As Currency
    Dim curDate As Date
    
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额  From 门诊费用记录 Where 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng结帐ID)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng结帐ID)
    
    lng冲销ID = rsTemp("结帐ID")
    
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", TYPE_重庆壁山, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
'    str流水号 = rsTemp("支付顺序号")
    
'    strInput = "99|" & str流水号 & "|" & ToVarchar(UserInfo.姓名, 20)
'    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_重庆壁山, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_重庆壁山 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_重庆壁山 & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        cur个人帐户 * -1 & ",Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
        gstrSQL = "Select * from 结算周期记录 where 结帐id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
        If Not rsTemp.EOF Then
            str结算周期 = rsTemp!结算周期
    '        gstrSQL = "zl_结算周期记录_insert("
            gstrSQL = "Insert into zlhis.结算周期记录 values (" & lng结帐ID & ",'" & str结算周期 & "'," & cur票据总金额 * -1 & "," & cur个人帐户 * -1 & "," & Nvl(rsTemp("统筹"), 0) * -1 & ",'L',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
            gcnOracle.Execute gstrSQL
        End If
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "西铝厂医保")
    End If

    门诊结算冲销_壁山 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊虚拟结算_壁山(rs费用明细 As Recordset, str结算方式 As String) As Boolean
    Dim cur个帐支付 As Currency, cur个人现金支付 As Currency, cur个人帐户支付 As Currency
    Dim cur统筹支付 As Currency, cur大额支付 As Currency, lngCount As Long
    Dim strSQL As String, rs壁山 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, strPageId As String
    Dim lng病人ID As Long, cur费用总额 As Currency
    Dim i As Integer, frm等待 As New frm等待响应壁山
    Dim datCurr As Date, cur个人帐户余额 As Currency
    If Val(Get保险参数_壁山("适用地区")) = 0 Then         '如果不是西铝厂,则不用虚拟结算
        门诊虚拟结算_壁山 = False
        Exit Function
    End If
    '判断是否已经发生费用
    If rs费用明细.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Get保险参数_壁山("适用地区")) = 3 Or gbln特殊门诊 = False Or Val(Get保险参数_壁山("适用地区")) = 1 Then
        cur费用总额 = 0
        While Not rs费用明细.EOF
            cur费用总额 = cur费用总额 + rs费用明细!实收金额
            rs费用明细.MoveNext
        Wend
        str结算方式 = "个人帐户;" & cur费用总额 & ";1"
        门诊虚拟结算_壁山 = True
        Exit Function
    End If
    
    On Error GoTo errHandle
    '求出病人的病案主页，也同时就求出结算单号
    lng病人ID = rs费用明细!病人ID
    strBillNO = mstr门诊号
    
    '求出当前需要的序号
    strSQL = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    If rs壁山.State = adStateOpen Then rs壁山.Close
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
    If rs壁山.EOF Then
        i = 1
    Else
        i = Nvl(rs壁山("Charge_item_no"), 0) + 1
    End If
    rs费用明细.MoveFirst
    lngCount = 0
    If Val(Get保险参数_壁山("适用地区")) = 2 Then Call ShowWindow(frm等待.hwnd, 9)
    SetPos frm等待.hwnd
    frm等待.Move (Screen.Width - frm等待.Width) / 2, (Screen.Height - frm等待.Height) / 2
    DoEvents
    Do While Not rs费用明细.EOF
        '求出所有的费用金额
        cur个人帐户支付 = cur个人帐户支付 + rs费用明细("实收金额")
        gstrSQL = "Select * From 收费细目 where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "西南铝医院", CLng(rs费用明细("收费细目ID")))
        If rsTmp!类别 = 5 Or rsTmp!类别 = 6 Or rsTmp!类别 = 7 Then
            strMedi = "1"
        Else
            strMedi = "2"
        End If
        
        '进行数据提交准备
        strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                " values('" & strBillNO & "','" & _
                Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','" & _
                rs费用明细("病人ID") & "','" & rs费用明细("开单人") & _
                "',to_Date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & rs费用明细("保险编码") & _
                "','预结算','" & strMedi & "','" & _
                rs费用明细("计算单位") & "'," & rs费用明细("数量") & "," & _
                CStr(rs费用明细("单价")) & "," & CStr(rs费用明细("实收金额")) & _
                ",to_date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & UserInfo.姓名 & "')"
        gcn壁山.Execute strSQL
        
        '请求提交数据
        strSQL = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
        strBillNO & "','" & Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','0')"
        gcn壁山.Execute strSQL
        lngCount = lngCount + 1
        '请求查询数据(西铝在传输过程中不等待返回状态)
'        If frm等待.Result(2, strBillNo, i) = False Then
'            门诊虚拟结算_壁山 = False
'            MsgBox "在结算的过程之中发生中断", vbInformation, gstrSysName
'            GoTo ResetTrans
'        End If
'        '查询提交结果
'        strSql = "select Request_Result,Err_Code,Err_text from " & _
'                "check_item_request where Bill_no = '" & strBillNo & _
'                 "' and App_code = '" & Mid(gstr医院编码, 1, 4) & _
'                 "' and Charge_item_no = '" & CStr(i) & "'"
'        If rs壁山.State = adStateOpen Then rs壁山.Close
'        rs壁山.Open strSql, gcn壁山, adOpenStatic, adLockReadOnly
'        If rs壁山.BOF Then
'            门诊虚拟结算_壁山 = False
'            GoTo ResetTrans
'        Else
'            If rs壁山("Request_Result") = "0" Then
'                MsgBox "发生错误[" & rs壁山("Err_Code") & "]:" & vbCrLf & String(2, "　") & rs壁山("Err_text"), vbInformation, gstrSysName
'                门诊虚拟结算_壁山 = False
'                GoTo ResetTrans
'            End If
'        End If

        '对HIS之中的基础数据进行修改
        i = i + 1
        rs费用明细.MoveNext
    Loop
    Do While True
        '查询提交结果
        strSQL = "select Request_Result,Err_Code,Err_text from " & _
                "check_item_request where Bill_no = '" & strBillNO & _
                 "' and App_code = '" & Mid(gstr医院编码, 1, 4) & _
                 "' and Request_result is Null"
        If rs壁山.State = adStateOpen Then rs壁山.Close
        rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
        If rs壁山.EOF Then Exit Do
        DoEvents
    Loop
    Unload frm等待
    cur费用总额 = cur个人帐户支付
    '进行结算准备
    strSQL = "Update Check_doex_interface set Ps_account_pay = " & _
            CStr(cur个帐支付) & ",Bala_op_id = '" & ToVarchar(UserInfo.姓名, 8) & _
            "' where Bill_no = '" & mstr门诊号 & "' and " & _
            "App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    gcn壁山.Execute strSQL
    
    '提交结算请求
    strSQL = "update Check_bill_request set Request_status = '5',Request_Result=null where" & _
            " Bill_no ='" & strBillNO & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    gcn壁山.Execute strSQL
    
    If Checkrequest(strBillNO) = False Then
        门诊虚拟结算_壁山 = False
        GoTo ResetTrans
    End If
    
    '从对方的数据库之中提取个人帐户支付、现金支付、统筹支付、大额支付
    strSQL = "select Ps_bala from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    If rs壁山.State = adStateOpen Then rs壁山.Close
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
    cur个人帐户支付 = Nvl(rs壁山("Ps_bala"), 0)
    
    strSQL = "select Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    If rs壁山.State = adStateOpen Then rs壁山.Close
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
    cur个人帐户支付 = Nvl(rs壁山("Ps_account_pay"), 0)
    cur个人现金支付 = Nvl(rs壁山("Ps_cost_pay"), 0)
    cur统筹支付 = Nvl(rs壁山("Plan_pay"), 0)
    cur大额支付 = Nvl(rs壁山("Big_pay"), 0)
    
'    '西铝计算个人帐户支付
'    cur费用总额 = cur费用总额 - cur统筹支付 - cur大额支付
'    cur个人帐户支付 = IIf(cur个人帐户支付 > cur费用总额, cur费用总额, cur个人帐户支付)
    
    str结算方式 = "个人帐户;" & cur个人帐户支付 & ";" & IIf(Val(Get保险参数_壁山("适用地区")) = 3 Or Val(Get保险参数_壁山("适用地区")) = 2, 1, 0) '允许修改个人帐户
    If cur统筹支付 <> 0 Then
        str结算方式 = str结算方式 & IIf(str结算方式 = "", "", "|") & "统筹支付;" & cur统筹支付 & ";0" '不允许修改统筹支付
    End If
    If cur大额支付 <> 0 Then
        str结算方式 = str结算方式 & IIf(str结算方式 = "", "", "|") & "大额支付;" & cur大额支付 & ";0" '不允许修改大额支付
    End If
    门诊虚拟结算_壁山 = True
ResetTrans:             '以红字单据冲掉为预结算而上传的费用明细
    '求出当前需要的序号
    rs费用明细.MoveFirst
    strSQL = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    If rs壁山.State = adStateOpen Then rs壁山.Close
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
    i = Nvl(rs壁山("Charge_item_no"), 0) + 1
    rs费用明细.MoveFirst
    If Val(Get保险参数_壁山("适用地区")) = 2 Then Call ShowWindow(frm等待.hwnd, 9)
    SetPos frm等待.hwnd
    frm等待.Move (Screen.Width - frm等待.Width) / 2, (Screen.Height - frm等待.Height) / 2
    DoEvents
    Do While Not rs费用明细.EOF And lngCount > 0
        '求出所有的费用金额
        cur个人帐户支付 = cur个人帐户支付 + rs费用明细("实收金额")
        gstrSQL = "Select * From 收费细目 where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "西南铝医院", CLng(rs费用明细("收费细目ID")))
        If rsTmp!类别 = 5 Or rsTmp!类别 = 6 Or rsTmp!类别 = 7 Then
            strMedi = "1"
        Else
            strMedi = "2"
        End If
        '进行数据提交准备
        strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                " values('" & strBillNO & "','" & _
                Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','" & _
                rs费用明细("病人ID") & "','" & rs费用明细("开单人") & _
                "',to_Date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & rs费用明细("保险编码") & _
                "','预结算','" & strMedi & "','" & _
                rs费用明细("计算单位") & "'," & 0 - rs费用明细("数量") & "," & _
                CStr(rs费用明细("单价")) & "," & CStr(0 - rs费用明细("实收金额")) & _
                ",to_date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & UserInfo.姓名 & "')"
        gcn壁山.Execute strSQL
        
        '请求提交数据
        strSQL = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
        strBillNO & "','" & Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','0')"
        gcn壁山.Execute strSQL
        lngCount = lngCount - 1
        '请求查询数据
'        If frm等待.Result(2, strBillNo, i) = False Then
'            门诊虚拟结算_壁山 = False
'            MsgBox "在结算的过程之中发生中断", vbInformation, gstrSysName
'            Exit Function
'        End If
        '查询提交结果
        
        i = i + 1
        rs费用明细.MoveNext
    Loop
    Do While True
        '查询提交结果
        strSQL = "select Request_Result,Err_Code,Err_text from " & _
                "check_item_request where Bill_no = '" & strBillNO & _
                 "' and App_code = '" & Mid(gstr医院编码, 1, 4) & _
                 "' and Request_result is Null"
        If rs壁山.State = adStateOpen Then rs壁山.Close
        rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
        If rs壁山.EOF Then Exit Do
        DoEvents
    Loop
    
    Unload frm等待
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    门诊虚拟结算_壁山 = False
End Function

Public Function 费用明细传递(lng类别 As Long, Optional lng结帐ID As Long, Optional strNO As String, Optional lng病人ID As Long, Optional int性质 As Integer, Optional int状态 As Integer) As Boolean
'功能：逐笔提交门诊费用明细
'lng类别： 1、门诊  2、住院
'lng结帐ID：用来处理门诊费用
'strNo:单据号
'int性质：
'lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rs壁山 As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, i As Integer, j As Integer, rsTemp As New ADODB.Recordset
    'i-循环累加，处方序号，在不插入明细时不累加；j-临时记录原序号，供再次申请时使用
    Dim blnInsert As Boolean
    Dim frm等待 As New frm等待响应壁山
     
    On Error GoTo errHandle
    If lng病人ID = 0 Then
        If lng类别 = 1 Then
            gstrSQL = "select 病人ID from 门诊费用记录 where 结帐ID = " & _
                    lng结帐ID & " and rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng结帐ID)
        Else
            gstrSQL = "select 病人ID from 住院费用记录 where NO = [1] and 记录性质 = [2] and 记录状态 = [3] and rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", strNO, int性质, int状态)
        End If
        
        lng病人ID = rsTmp("病人ID")
    End If
    If lng类别 = 1 Then
       strBillNO = mstr门诊号
    Else
        gstrSQL = "select max(主页ID) as 主页ID from 病案主页 where 病人ID =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng病人ID)
        strBillNO = CStr(lng病人ID) & "_" & CStr(rsTmp("主页ID"))
    End If
    If lng类别 = 1 Then
        '将以前传递的记录和检查记录进行删除:注意，收费细目如何进行传递还需要修改
        strSQL = "delete from Check_item_list_interface where Bill_no = '" & _
                mstr门诊号 & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        strSQL = "delete from Check_item_request where Bill_no = '" & _
                mstr门诊号 & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        gstrSQL = "select A.ID,A.发生时间,A.序号,A.NO,A.收费类别,A.开单人,A.登记时间," & _
                "A.收费细目ID,A.收据费目,A.记录性质,A.记录状态,D.项目编码 as 细目编码,B.名称 as 细目名称," & _
                "C.名称  as 项目种类,B.计算单位, (A.数次 * A.付数) as 数量," & _
                "A.标准单价,A.实收金额,A.操作员姓名,A.是否上传 from  " & _
                "门诊费用记录 A,收费细目 B,收入项目 C,保险支付项目 D" & _
                " where A.记录状态<>0 And Nvl(A.是否上传,0)=0 And A.收费细目ID = B.ID and A.收入项目ID = C.ID and " & _
                "A.结帐ID =[3] and A.收费细目ID = D.收费细目ID and D.险类 = [1] And a.病人ID = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", TYPE_重庆壁山, lng病人ID, lng结帐ID)
    Else
        gstrSQL = "select A.ID,A.发生时间,A.序号,A.NO,A.收费类别,A.开单人,A.登记时间," & _
                "A.收费细目ID,A.收据费目,A.记录性质,A.记录状态,D.项目编码 as 细目编码,B.名称 as 细目名称,C.名称 as " & _
                "项目种类,B.计算单位, (A.数次 * A.付数) as 数量,A.标准单价,A.实收金额," & _
                "A.操作员姓名,A.是否上传 from 住院费用记录 A,收费细目 B,收入项目 C,保险支付项目 D,保险帐户 E" & _
                " where A.记录状态<>0 And Nvl(A.是否上传,0)=0 And A.收费细目ID = B.ID and A.收入项目ID = C.ID " & _
                " and A.NO =[2] and A.记录状态 = [3] and A.记录性质 = [4] and A.收费细目ID = D.收费细目ID " & _
                " and A.病人ID=E.病人ID And E.险类=[1] And D.险类 = [1] and A.病人ID = [5]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", TYPE_重庆壁山, strNO, int状态, int性质, lng病人ID)
    End If
    
    If rsTmp.BOF Then 费用明细传递 = False: Exit Function
    '求出初始传递的号码
    strSQL = "select max(Charge_item_no) as Charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    If rs壁山.State = adStateOpen Then rs壁山.Close
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
    If rs壁山.EOF Then
        i = 1
    Else
        i = Nvl(rs壁山("Charge_item_no"), 0) + 1
    End If
    '逐步进行费用明细传递
    If Val(Get保险参数_壁山("适用地区")) = 2 Then Call ShowWindow(frm等待.hwnd, 5)
    SetPos frm等待.hwnd
    frm等待.Move (Screen.Width - frm等待.Width) / 2, (Screen.Height - frm等待.Height) / 2
    DoEvents
    Do While Not rsTmp.EOF
        '判断记录是否已经上传
        '如果根本没有处理，或已上传，则不必管这类数据，我们认为已上传
        blnInsert = True
        If Val(Get保险参数_壁山("适用地区")) = 3 And lng类别 = 2 Then
            strSQL = "Select Charge_item_no,Nvl(Flag,'0') AS Flag From Check_item_list_interface Where HIS_PK=" & rsTmp!ID
            If rs壁山.State = adStateOpen Then rs壁山.Close
            rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
            If rs壁山.RecordCount = 0 Then
                '表中无此记录，可继续上传
            Else
                If rs壁山!flag = "1" Then
                    '此ID的记录已经成功上传,跳转
                    GoTo nextRec
                Else
                    '上传失败或医保商程序未响应，不必再插入明细，直接删除上次的请求记录，避免再次上传，然后插入新的请求记录
                    blnInsert = False
                    j = i
                    i = rs壁山!Charge_item_no
                    
                    '删除请求表
                    strSQL = "Delete check_item_request where Bill_no = '" & strBillNO & _
                             "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "' And charge_Item_NO='" & CStr(i) & "'"
                    gcn壁山.Execute strSQL
                End If
            End If
        End If
        
        If blnInsert Then
            '作提交数据的准备,如果为门诊病人就传递“病人ID + 时间”，如果为住院病人，就传递病人ID和主页ID
            If InStr(1, ",5,6,7,", "," & rsTmp!收费类别 & ",") <> 0 Then
                strMedi = "1"
            Else
                strMedi = "2"
            End If
            If Val(Get保险参数_壁山("适用地区")) <> 1 Then
                strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                        "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code,App_item_name," & _
                        "Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id" & _
                        IIf(Val(Get保险参数_壁山("适用地区")) = 3, ",HIS_PK)", ")") & _
                        " values('" & strBillNO & "','" & _
                        Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','" & rsTmp("NO") & "','" & _
                        rsTmp("开单人") & "',to_date('" & Format(rsTmp("登记时间"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS'),'" & _
                        rsTmp("细目编码") & "','" & rsTmp("细目名称") & "','" & strMedi & _
                        "','" & rsTmp("计算单位") & "'," & rsTmp("数量") & "," & CStr(rsTmp("标准单价")) & "," & _
                        CStr(rsTmp("实收金额")) & ",to_date('" & Format(rsTmp("登记时间"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS'),'" & _
                        rsTmp("操作员姓名") & "'" & IIf(Val(Get保险参数_壁山("适用地区")) = 3, "," & rsTmp!ID & ")", ")")
            Else
                strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                        "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code,App_item_name," & _
                        "Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                        " values('" & strBillNO & "','" & _
                        Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','" & rsTmp("NO") & "','" & _
                        rsTmp("开单人") & "','" & rsTmp("登记时间") & "','" & _
                        rsTmp("细目编码") & "','" & rsTmp("细目名称") & "','" & strMedi & _
                        "','" & rsTmp("计算单位") & "'," & rsTmp("数量") & "," & CStr(rsTmp("标准单价")) & "," & _
                        CStr(rsTmp("实收金额")) & ",'" & rsTmp("登记时间") & "','" & _
                        rsTmp("操作员姓名") & "')"
            End If
            Call DebugTool(strSQL)
            gcn壁山.Execute strSQL
        End If
        
        '请求提交数据
        strSQL = "Insert into Check_item_request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
                strBillNO & "','" & Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','0')"
        Call DebugTool(strSQL)
        gcn壁山.Execute strSQL
        '查询提交结果
        If Val(Get保险参数_壁山("适用地区")) <> 2 Then
            If frm等待.Result(2, strBillNO, i) = False Then
                Call DebugTool("费用明细传递发生中断")
                费用明细传递 = False
                MsgBox "费用明细传递发生中断", vbInformation, gstrSysName
                Exit Function
            End If
            strSQL = "select Request_Result,Err_Code,Err_text from check_item_request" & _
                    " where Bill_no = '" & strBillNO & _
                     "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "' and Charge_item_no = '" & _
                     CStr(i) & "'"
            If rs壁山.State = adStateOpen Then rs壁山.Close
            rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
            If rs壁山.BOF Then
                Call DebugTool("Check_Item_Request记录为空")
                费用明细传递 = False
                Exit Function
            Else
                If rs壁山("Request_Result") = "1" Then
                    Call DebugTool("发生错误" & rs壁山("Err_Code") & ":" & vbCrLf & String(2, "　") & rs壁山("Err_text"))
                    MsgBox "发生错误" & rs壁山("Err_Code") & ":" & vbCrLf & String(2, "　") & rs壁山("Err_text"), vbInformation, gstrSysName
                    费用明细传递 = False
                    Exit Function
                End If
            End If
        End If
'如果该笔费用已传递，则跳转
nextRec:
        '对HIS之中的基础数据进行修改
        gstrSQL = "zl_病人记帐记录_上传 ('" & rsTmp("ID") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
        rsTmp.MoveNext
        
        If blnInsert Then
            i = i + 1       '累加
        Else
            i = j           '未插入明细，则还原当前明细累加值
        End If
    Loop
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
        Do While True
            '查询提交结果
            strSQL = "select Request_Result,Err_Code,Err_text from " & _
                    "check_item_request where Bill_no = '" & strBillNO & _
                     "' and App_code = '" & Mid(gstr医院编码, 1, 4) & _
                     "' and Request_result is Null"
            If rs壁山.State = adStateOpen Then rs壁山.Close
            rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
            If rs壁山.EOF Then Exit Do
            DoEvents
        Loop
        Unload frm等待
    End If
    '清除本次数据上传请求
    strSQL = "Delete check_item_request where Bill_no = '" & strBillNO & _
             "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    gcn壁山.Execute strSQL
    
    rs壁山.Close
    费用明细传递 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    费用明细传递 = False
End Function

Private Function Get病人ID(str医保号 As String, str医保中心编码 As String) As String
'功能：通过医保中心号码和医保号求出病人ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 病人ID from 保险帐户 where 险类 = [1] and 医保号 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", CInt(str医保中心编码), str医保号)
    If Not rsTmp.BOF Then
        Get病人ID = CStr(rsTmp("病人ID"))
    Else
        Get病人ID = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get病人ID = ""
End Function

Public Function 个人余额_壁山(str病人ID As String) As Currency
'功能：通过病人的信息求出个人余额
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strTime As String, rs壁山 As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    'Modified By 朱玉宝 下午 06:06:13
    If Val(Get保险参数_壁山("适用地区")) = 1 Then
        '如果是适用于黔江地区，直接从保险帐户中读取
        gstrSQL = "Select 帐户余额 余额 From 保险帐户 Where 病人ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户余额", CLng(Val(str病人ID)))
        个人余额_壁山 = Nvl(rsTmp!余额, 0)
    Else
        '如果虚拟结算不通过，直接返回
        gstrSQL = "select 卡号,密码 from 保险帐户 where 病人ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", CLng(Val(str病人ID)))
        If rsTmp.BOF Then 个人余额_壁山 = 0: Exit Function
        '在数据库之中获取持卡病人的验证信息
        strTime = CStr(Format(zlDatabase.Currentdate, "yyyymmddhhmmss")) & "00"
        strSQL = "insert into Check_doex_interface(Bill_no,App_code," & _
                "Ic_id,Doct_flag,Is_bala,Regi_op_id) values('" & strTime & "','" & Mid(gstr医院编码, 1, 4) & "','" & _
                rsTmp("密码") & rsTmp("卡号") & "','0','0','" & ToVarchar(UserInfo.姓名, 8) & "')"
        gcn壁山.Execute strSQL
        strSQL = "insert into Check_bill_request(Bill_no,App_code," & _
                "Request_status) values('" & strTime & "','" & Mid(gstr医院编码, 1, 4) & _
                "','2')"
        gcn壁山.Execute strSQL
        If Checkrequest(strTime) = False Then 个人余额_壁山 = 0: Exit Function
        '从信息之中提取病人的个人帐户余额
        strSQL = "select Ps_Bala from Check_Doex_Interface where Bill_no = '" & strTime & "'" & _
                " and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        If rs壁山.State = adStateOpen Then rs壁山.Close
        rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
        If Not rs壁山.BOF Then
            个人余额_壁山 = IIf(IsNull(rs壁山("Ps_Bala")), 0, rs壁山("Ps_Bala"))
        Else
            个人余额_壁山 = 0
        End If
        strSQL = "delete from Check_bill_request where Bill_no = '" & _
                strTime & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        strSQL = "delete from Check_doex_interface where Bill_no = '" & _
                strTime & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    个人余额_壁山 = 0
End Function

Public Function 入院登记_壁山(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strSQL As String, strInNote As String
    Dim int转院 As Integer              '2-转院;1-普通入院
    Dim rsTmp As New ADODB.Recordset
    
    '求出病人的相关信息
    On Error GoTo errHandle
    gstrSQL = "select A.入院日期,A.入院方式,B.住院号,D.名称 as 住院科室,A.入院病床,A.门诊医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng病人ID, lng主页ID)
    int转院 = 1
    If (rsTmp!入院方式 Like "*转入*" Or rsTmp!入院方式 Like "*转院*") Then int转院 = 2
    
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID)   '入院诊断
    If Val(Get保险参数_壁山("适用地区")) = 2 And gstr特殊病种 <> "" Then
        '检查是否特殊病
        strInNote = gstr特殊病种
    End If
    If rsTmp.BOF Then 入院登记_壁山 = False: Exit Function
    '准备进行提交
    strSQL = "Delete from Check_doex_interface where bill_no='" & lng病人ID & "_" & lng主页ID & "' and App_code='" & Mid(gstr医院编码, 1, 4) & "' and Doct_flag=1 and Hosp_No is null"
    gcn壁山.Execute strSQL
    strSQL = "Delete from Check_bill_request where bill_no='" & lng病人ID & "_" & lng主页ID & "' and App_code='" & Mid(gstr医院编码, 1, 4) & "'"
    gcn壁山.Execute strSQL
    
    If Val(Get保险参数_壁山("适用地区")) = 1 Then
        strSQL = "Insert into Check_doex_interface(Bill_no,App_code,Doct_flag," & _
                "Doex_no,In_mode,Ill_type,Ic_id,Is_bala,Regi_op_id,Sec_off,The_bunk," & _
                "In_time,Tre_dr) values('" & lng病人ID & "_" & lng主页ID & _
                "','" & Mid(gstr医院编码, 1, 4) & "','1','" & Nvl(rsTmp("住院号")) & "_" & lng主页ID & "','" & int转院 & "','" & _
                strInNote & "','" & Nvl(rsTmp("密码")) & Nvl(rsTmp("卡号")) & "','0','" & ToVarchar(UserInfo.姓名, 8) & _
                "','" & Nvl(rsTmp("住院科室")) & "','" & ToVarchar(Nvl(rsTmp("入院病床"), "0"), 24) & "'," & _
                " '" & Nvl(rsTmp("入院日期")) & "'" & _
                ",'" & Nvl(rsTmp("门诊医师"), "未知") & "')"
    Else
        strSQL = "Insert into Check_doex_interface(Bill_no,App_code,Doct_flag," & _
                "Doex_no,In_mode,Ill_type,Ic_id,Is_bala,Regi_op_id,Sec_off,The_bunk," & _
                "In_time,Tre_dr) values('" & lng病人ID & "_" & lng主页ID & _
                "','" & Mid(gstr医院编码, 1, 4) & "','1','" & Nvl(rsTmp("住院号")) & "_" & lng主页ID & "','" & int转院 & "','" & _
                strInNote & "','" & Nvl(rsTmp("密码")) & Nvl(rsTmp("卡号")) & "','0','" & ToVarchar(UserInfo.姓名, 8) & _
                "','" & Nvl(rsTmp("住院科室")) & "','" & ToVarchar(Nvl(rsTmp("入院病床"), "0"), 24) & "'," & _
                " to_date('" & Format(rsTmp("入院日期"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS')" & _
                ",'" & Nvl(rsTmp("门诊医师"), "未知") & "')"
    End If
    gcn壁山.Execute strSQL
    '进行入院请求
    strSQL = "Insert into Check_bill_request(Bill_no,App_code,Request_status)" & _
            "values('" & lng病人ID & "_" & lng主页ID & "','" & _
            Mid(gstr医院编码, 1, 4) & "','0')"
    gcn壁山.Execute strSQL
    '查询请求的结果
    If Checkrequest(lng病人ID & "_" & lng主页ID) = False Then
        入院登记_壁山 = False
        Exit Function
    End If
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆壁山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    入院登记_壁山 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_壁山 = False
End Function

Public Function 记帐传输_壁山(strNO As String, int性质 As Integer, int状态 As Integer, Optional lng病人ID As Long) As Boolean
'将住院病人的费用传递到医保服务器并且同时修改病人费用信息之中的数据
    If lng病人ID = 0 Then
        记帐传输_壁山 = 费用明细传递(2, , strNO, , int性质, int状态)
    Else
        记帐传输_壁山 = 费用明细传递(2, , strNO, lng病人ID, int性质, int状态)
    End If
End Function

Public Function 住院虚拟结算_壁山(rs费用明细 As Recordset, lng病人ID As Long, str医保号 As String, str密码 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    Dim cur个人帐户支付 As Currency, cur个人现金支付 As Currency
    Dim cur统筹支付 As Currency, cur大额支付 As Currency, cur费用总额 As Currency
    Dim strSQL As String, rs壁山 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, strPageId As String
    Dim i As Integer, j As Integer, frm等待 As New frm等待响应壁山
    Dim datCurr As Date, cur个人帐户余额 As Currency
    Dim blnInsert As Boolean
    
    '判断是否已经发生费用
    If rs费用明细.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    On Error GoTo errHandle
    '求出病人的病案主页，也同时就求出结算单号
    gstrSQL = "select max(主页ID) as 主页ID from 病案主页 where 病人ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng病人ID)
    strPageId = CStr(rsTmp("主页ID"))
    strBillNO = CStr(lng病人ID) & "_" & CStr(rsTmp("主页ID"))
    rs费用明细.Sort = "是否上传 desc"
    
    '求出当前需要的序号
    strSQL = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    If rs壁山.State = adStateOpen Then rs壁山.Close
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
    If rs壁山.EOF Then
        i = 1
    Else
        i = Nvl(rs壁山("Charge_item_no"), 0) + 1
    End If
    rs费用明细.MoveFirst
    If Val(Get保险参数_壁山("适用地区")) = 3 Then
        '清除预结算
        strSQL = "Update Check_bill_request set Request_status = '6',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        If Checkrequest(strBillNO) = False Then
            住院虚拟结算_壁山 = ""
            Exit Function
        End If
    End If
    
    If Val(Get保险参数_壁山("适用地区")) = 2 Then Call ShowWindow(frm等待.hwnd, 5)
    SetPos frm等待.hwnd

    frm等待.Move (Screen.Width - frm等待.Width) / 2, (Screen.Height - frm等待.Height) / 2
    DoEvents
    
    Do While Not rs费用明细.EOF
        '求出所有的费用金额
        cur个人帐户支付 = cur个人帐户支付 + rs费用明细("金额")
        '如果费用还没有上传，就进行上传:注意，收费细目如何进行传递还需要修改
        
        If IIf(IsNull(rs费用明细("是否上传")), "0", rs费用明细("是否上传")) = "0" Then
            gstrSQL = "select A.ID,A.发生时间,A.序号,A.收费类别,A.NO,A.开单人,A.登记时间," & _
                    "A.收费细目ID,A.收据费目,A.记录性质,A.记录状态,D.项目编码 as 细目编码,B.名称 as 细目名称,C.名称" & _
                    " as 项目种类,B.计算单位, (A.数次 * A.付数) as 数量," & _
                    "A.标准单价,A.实收金额,A.操作员姓名 from 住院费用记录 A," & _
                    "收费细目 B,收入项目 C,保险支付项目 D where A.收费细目ID = B.ID and " & _
                    "A.收入项目ID = C.ID " & " And A.病人ID=[3]" & _
                    " and A.NO =[4] and A.记录状态 = [5] and A.记录性质 = [6]" & _
                    " and (A.价格父号 = [2] or A.价格父号 Is Null And A.序号=[2])" & _
                    " and (A.是否上传 = 0 or A.是否上传 is null) and " & _
                    "A.收费细目ID = D.收费细目ID and D.险类 = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", TYPE_重庆壁山, CLng(rs费用明细("序号")), lng病人ID, CStr(rs费用明细("NO")), CInt(rs费用明细("记录状态")), CInt(rs费用明细("记录性质")))
            If Not rsTmp.BOF Then
                '判断该笔费用是否已经上传
                blnInsert = True
                If Val(Get保险参数_壁山("适用地区")) = 3 Then
                    strSQL = "Select Charge_item_no,Nvl(Flag,'0') AS Flag From Check_item_list_interface Where HIS_PK=" & rsTmp!ID
                    If rs壁山.State = adStateOpen Then rs壁山.Close
                    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
                    If rs壁山.RecordCount = 0 Then
                        '表中无此记录，可继续上传
                    Else
                        If rs壁山!flag = "1" Then
                            '此ID的记录已经成功上传,跳转
                            GoTo nextRec
                        Else
                            '上传失败或医保商程序未响应，不必再插入明细，直接删除上次的请求记录，避免再次上传，然后插入新的请求记录
                            blnInsert = False
                            j = i
                            i = rs壁山!Charge_item_no
                            
                            '删除请求表
                            strSQL = "Delete check_item_request where Bill_no = '" & strBillNO & _
                                     "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "' And charge_Item_NO='" & CStr(i) & "'"
                            gcn壁山.Execute strSQL
                        End If
                    End If
                End If
                
                If blnInsert Then
                    If InStr(1, ",5,6,7,", "," & rsTmp!收费类别 & ",") <> 0 Then
                        strMedi = "1"
                    Else
                        strMedi = "2"
                    End If
                    '进行数据提交准备
                    If Val(Get保险参数_壁山("适用地区")) = 1 Then
                        strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                                " values('" & strBillNO & "','" & _
                                Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','" & _
                                rsTmp("NO") & "','" & rsTmp("开单人") & _
                                "','" & rsTmp("登记时间") & _
                                "','" & rsTmp("细目编码") & _
                                "','" & rsTmp("细目名称") & "','" & strMedi & "','" & _
                                rsTmp("计算单位") & "'," & rsTmp("数量") & "," & _
                                CStr(rsTmp("标准单价")) & "," & CStr(rsTmp("实收金额")) & _
                                ",'" & rsTmp("登记时间") & _
                                "','" & rsTmp("操作员姓名") & "')"
                    Else
                        strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id" & _
                                IIf(Val(Get保险参数_壁山("适用地区")) = 3, ",HIS_PK)", ")") & _
                                " values('" & strBillNO & "','" & _
                                Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','" & _
                                rsTmp("NO") & "','" & rsTmp("开单人") & _
                                "',to_Date('" & Format(rsTmp("登记时间"), "yyyy-MM-dd HH:MM:SS") & _
                                "','yyyy-MM-dd HH24:MI:SS'),'" & rsTmp("细目编码") & _
                                "','" & rsTmp("细目名称") & "','" & strMedi & "','" & _
                                rsTmp("计算单位") & "'," & rsTmp("数量") & "," & _
                                CStr(rsTmp("标准单价")) & "," & CStr(rsTmp("实收金额")) & _
                                ",to_date('" & Format(rsTmp("登记时间"), "yyyy-MM-dd HH:MM:SS") & _
                                "','yyyy-MM-dd HH24:MI:SS'),'" & rsTmp("操作员姓名") & "'" & _
                                IIf(Val(Get保险参数_壁山("适用地区")) = 3, "," & rsTmp!ID & ")", ")")
                    End If
                    Call DebugTool(strSQL)
                    gcn壁山.Execute strSQL
                End If
                
                '请求提交数据
                strSQL = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
                strBillNO & "','" & Mid(gstr医院编码, 1, 4) & "','" & CStr(i) & "','0')"
                Call DebugTool(strSQL)
                gcn壁山.Execute strSQL
                '请求查询数据
                If Val(Get保险参数_壁山("适用地区")) <> 2 Then
                    If frm等待.Result(2, strBillNO, i) = False Then
                        Call DebugTool("在结算的过程之中发生中断")
                        住院虚拟结算_壁山 = ""
                        MsgBox "在结算的过程之中发生中断", vbInformation, gstrSysName
                        Exit Function
                    End If
                    '查询提交结果
                    strSQL = "select Request_Result,Err_Code,Err_text from " & _
                            "check_item_request where Bill_no = '" & strBillNO & _
                             "' and App_code = '" & Mid(gstr医院编码, 1, 4) & _
                             "' and Charge_item_no = '" & CStr(i) & "'"
                    If rs壁山.State = adStateOpen Then rs壁山.Close
                    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
                    If rs壁山.BOF Then
                        Call DebugTool("Check_Item_Request记录为空(住院虚拟结算_壁山)")
                        住院虚拟结算_壁山 = ""
                        Exit Function
                    Else
                        If rs壁山("Request_Result") = "1" Then
                            Call DebugTool("发生错误[" & rs壁山("Err_Code") & "]:" & vbCrLf & String(2, "　") & rs壁山("Err_text"))
                            MsgBox "发生错误[" & rs壁山("Err_Code") & "]:" & vbCrLf & String(2, "　") & rs壁山("Err_text"), vbInformation, gstrSysName
                            住院虚拟结算_壁山 = ""
                            Exit Function
                        End If
                    End If
                End If
'如果该记录已经上传则跳转到
nextRec:
                '对HIS之中的基础数据进行修改
                gstrSQL = "zl_病人记帐记录_上传 ('" & rsTmp("ID") & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
                
                '上传明细后才累加处方明细序号
                If blnInsert Then
                    i = i + 1       '累加
                Else
                    i = j           '未插入明细，则还原当前明细累加值
                End If
            End If
        End If
        rs费用明细.MoveNext
    Loop
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
'        Do While True
'            '查询提交结果
'            strSql = "select Request_Result,Err_Code,Err_text from " & _
'                    "check_item_request where Bill_no = '" & strBillNo & _
'                     "' and App_code = '" & Mid(gstr医院编码, 1, 4) & _
'                     "' and Request_result is Null"
'            If rs壁山.State = adStateOpen Then rs壁山.Close
'            rs壁山.Open strSql, gcn壁山, adOpenStatic, adLockReadOnly
'            If rs壁山.EOF Then Exit Do
'            DoEvents
'        Loop
        Unload frm等待
    End If
    cur费用总额 = cur个人帐户支付
    If Val(Get保险参数_壁山("适用地区")) <> 1 Then
        '作出提交准备
        datCurr = zlDatabase.Currentdate
        '取消将个人帐户支付额设置为总额的做法，因为西希预结算时仅处理统筹支付与大额支付，只有正式结算时才更新帐户支付与现金支付
        '个人帐户支付=总费用-统筹支付-大额支付;在正式结算时根据实际帐户支付额更新Ps_account_pay，然后西希会检查并更正该字段，我们再校正
        'Ps_account_pay = " & _
                cur个人帐户支付 & ",
        strSQL = "Update Check_doex_interface set Bala_op_id = '" & ToVarchar(UserInfo.姓名, 8) & _
                "',Out_time =to_date('" & Format(datCurr, "yyyy-MM-dd") & "','yyyy-MM-dd') " & _
                "where Bill_no = '" & strBillNO & "' and App_code = '" & _
                Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        '进行虚拟结算请求,目前还不知道具体的参数值,在编译之后需要进行修改
        strSQL = "Update Check_bill_request set Request_status = '2',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        If Checkrequest(strBillNO) = False Then
            住院虚拟结算_壁山 = ""
            Exit Function
        End If
        strSQL = "select Ps_bala from" & _
                " Check_doex_interface where Bill_no = '" & strBillNO & _
                "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        If rs壁山.State = adStateOpen Then rs壁山.Close
        rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
        cur个人帐户支付 = Nvl(rs壁山("Ps_bala"), 0) '此处取的是帐户余额
        
        strSQL = "Update Check_bill_request set Request_status = '5',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        If Checkrequest(strBillNO) = False Then
            住院虚拟结算_壁山 = ""
            Exit Function
        End If
    Else
        MsgBox "请进行手工结算，结算完成后点击“确定”继续......", vbInformation, "医业软件"
    End If
    
    '取当前帐户余额
    cur个人帐户余额 = 个人余额_壁山(CStr(lng病人ID))
    '从对方的数据库之中提取个人帐户支付、现金支付、统筹支付、大额支付（根本不需要提取帐户与现金支付，西希没有更新这两个字段）
    strSQL = "select Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    If rs壁山.State = adStateOpen Then rs壁山.Close
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
    cur统筹支付 = Nvl(rs壁山("Plan_pay"), 0)
    cur大额支付 = Nvl(rs壁山("Big_pay"), 0)
    If Val(Get保险参数_壁山("适用地区")) = 1 Or Val(Get保险参数_壁山("适用地区")) = 2 Then  '在这里将西铝医院也修改成与壁山相同的处理方式.HXB
        cur个人帐户支付 = Nvl(rs壁山("Ps_account_pay"), 0)            '壁山返回个人帐户支付
        cur个人现金支付 = Nvl(rs壁山("Ps_cost_pay"), 0)
    Else
        '除统筹与大额外，都可以由帐户支付
        cur个人现金支付 = cur费用总额 - cur统筹支付 - cur大额支付
        '帐户支付不能大于帐户余额
        cur个人帐户支付 = IIf(cur个人帐户支付 > cur个人帐户余额, cur个人帐户余额, cur个人帐户支付)
        '计算帐户最多支付多少
        cur个人帐户支付 = IIf(cur个人帐户支付 >= cur个人现金支付, cur个人现金支付, cur个人帐户支付)
        cur个人现金支付 = cur个人现金支付 - cur个人帐户支付
'        '更新帐户支付额（在结算处已更新）
'        strSql = "Update Check_doex_interface set Ps_account_pay = '" & cur个人帐户支付 & "'" & _
'                "where Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
'        gcn壁山.Execute strSql
    End If
    
'    '西铝计算个人帐户支付金额
'    If Val(Get保险参数_壁山("适用地区")) <> 1 Then
'        'cur费用总额 = cur费用总额 - cur统筹支付 - cur大额支付
'        cur个人帐户支付 = IIf(cur个人帐户支付 > cur费用总额, cur费用总额, cur个人帐户支付)
'    End If
'    gstrSQL = "Select Nvl(帐户余额,0) 余额 From 保险帐户 Where 病人ID=" & lng病人ID
'    Call OpenRecordset(rsTmp, "读取帐户余额")
'    cur个人帐户余额 = rsTmp!余额
    
    If (Val(Get保险参数_壁山("适用地区")) = 3 Or Val(Get保险参数_壁山("适用地区")) = 0) Then
        住院虚拟结算_壁山 = "个人帐户;" & cur个人帐户支付 & ";1"
    Else
        住院虚拟结算_壁山 = "个人帐户;" & cur个人帐户支付 & ";0" '不允许修改个人帐户
    End If
    住院虚拟结算_壁山 = 住院虚拟结算_壁山 & "|统筹支付;" & cur统筹支付 & ";0" '不允许修改统筹支付
    住院虚拟结算_壁山 = 住院虚拟结算_壁山 & "|大额统筹;" & cur大额支付 & ";0" '不允许修改大额支付
    
    '清除本次数据上传请求
    strSQL = "Delete check_item_request where Bill_no = '" & strBillNO & _
             "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    gcn壁山.Execute strSQL
    
    With pre_Balance
        .cur大病基金 = cur大额支付
        .cur个人帐户 = cur个人帐户支付
        .cur医保基金 = cur统筹支付
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Resume
    住院虚拟结算_壁山 = ""
End Function

Public Function 住院结算_壁山(lng结帐ID As Long, ByVal lng病人ID As Long, Optional ByRef strAdvance As String = "") As Boolean
'将病人的费用进行结算，由于壁山医保不需要进行出院登记，因此不进行出院登记
    Dim rsTmp As New ADODB.Recordset, cur结算金额 As Currency
    Dim strBillNO As String, strSQL As String, datCurr As Date
    Dim rs壁山 As New ADODB.Recordset, cur个人帐户支付 As Currency
    Dim cur个人现金支付 As Currency, cur统筹支付 As Currency
    Dim cur大额支付 As Currency, int住院次数累计 As Integer
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur统筹自付 As Currency, cur基数自付 As Currency
    Dim cur超限自付 As Currency, cur大病统筹 As Currency
    Dim cur大病自付 As Currency, cur起付线 As Currency
    Dim cur全自付 As Currency, cur挂钩自付 As Currency
    Dim cur中心帐户 As Currency, str结算周期 As String
    Dim bln中途结算 As Boolean
    Dim str结算方式 As String
    Dim lng主页ID As Long
    Dim blnRevise As Boolean, blnOld As Boolean
    
    On Error GoTo errHandle
    Call DebugTool("准备进行结算，当前病人ID：" & lng病人ID)
    bln中途结算 = Not IS出院(lng病人ID)
    Call DebugTool("保险帐户的当前状态表示该病人当前：" & IIf(bln中途结算, "在院，将进行中结", "出院，将进行出院结算"))
    
    '读取本次个人帐户支付额
    gstrSQL = "Select Nvl(A.冲预交,0) 个人帐户 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=" & TYPE_重庆壁山 & _
        " And A.结算方式 in ('个人帐户') And A.结帐ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取本次个人帐户支付额", lng结帐ID)
    cur个人帐户支付 = 0
    If Not rsTmp.EOF Then
        cur个人帐户支付 = rsTmp!个人帐户
        pre_Balance.cur个人帐户 = cur个人帐户支付
    End If
    
    gstrSQL = "select sum(实收金额) as 结算金额,sum(结帐金额) as 已结金额 from 住院费用记录 where " & _
            "结帐ID=[2] and 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng病人ID, lng结帐ID)
    cur结算金额 = Nvl(rsTmp("已结金额"), 0)
    gstrSQL = "select 主页ID,出院日期 from 病案主页 where 主页ID=(select max(主页ID) from " & _
            "病案主页 where 病人ID  = [1]) and 病人ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng病人ID)
    If rsTmp.BOF Then Exit Function
    strBillNO = lng病人ID & "_" & rsTmp("主页ID")
    lng主页ID = rsTmp!主页ID
    If Val(Get保险参数_壁山("适用地区")) <> 1 Then
        '作出提交准备
        strSQL = "Update Check_doex_interface set Ps_account_pay = " & cur个人帐户支付 & _
                ",Bala_op_id = '" & ToVarchar(UserInfo.姓名, 8) & "',Out_time = to_date('" & _
                Format(rsTmp("出院日期"), "yyyy-MM-dd") & "','yyyy-MM-dd') where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        '进行结算请求
        If bln中途结算 = False Then
            Call DebugTool("出院结算")
            strSQL = "Update Check_bill_request set Request_status = '1',Request_Result=null where " & _
                    "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        Else
            Call DebugTool("中途结算")
            strSQL = "Update Check_Doex_interface Set Doct_Flag='1',Doex_Type='5' Where Bill_no='" & strBillNO & "' And App_code='" & Mid(gstr医院编码, 1, 4) & "'"
            gcn壁山.Execute strSQL
            strSQL = "Update Check_bill_request set Request_status = '6',Request_Result=null where " & _
                    "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
        End If
        gcn壁山.Execute strSQL
        If Checkrequest(strBillNO) = False Then 住院结算_壁山 = False: Exit Function
    End If
    '求出数据
    'modify by ccy, add select field Ps_bala
    strSQL = "select Ps_bala,Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay,acc_cyc from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'"
    If rs壁山.State = adStateOpen Then rs壁山.Close
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly
    cur个人帐户支付 = Nvl(rs壁山("Ps_account_pay"), 0)
    cur个人现金支付 = Nvl(rs壁山("Ps_cost_pay"), 0)
    cur统筹支付 = Nvl(rs壁山("Plan_pay"), 0)
    cur大额支付 = Nvl(rs壁山("Big_pay"), 0)
    cur大病统筹 = cur大额支付
    cur全自付 = cur个人帐户支付
    cur中心帐户 = Nvl(rs壁山("Ps_bala"), 0)
    str结算周期 = Nvl(rs壁山("ACC_CYC"), "")
    
    '比较虚拟结算与正式结算结果是否一致
    If Not (cur个人帐户支付 = pre_Balance.cur个人帐户 And cur统筹支付 = pre_Balance.cur医保基金 And _
        cur大额支付 = pre_Balance.cur大病基金) Then
        blnRevise = True
        #If gverControl < 2 Then
            blnOld = True
        #End If
    End If
    
    If blnRevise Then
        str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户支付
        str结算方式 = str结算方式 & "||统筹支付|" & cur统筹支付
        str结算方式 = str结算方式 & "||大额统筹|" & cur大病统筹
        If str结算方式 <> "" Then
            str结算方式 = Mid(str结算方式, 3)
            If blnOld Then
                gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
            Else
                strAdvance = str结算方式
                gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
        End If
    End If
    
    '填写结算表
    datCurr = zlDatabase.Currentdate
    '帐户年度信息
    Call Get帐户信息(TYPE_重庆壁山, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
            
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_重庆壁山 & "," & Year(datCurr) & "," & _
            cur帐户增加累计 & "," & cur帐户支出累计 & "," & _
            cur进入统筹累计 + cur统筹支付 + cur统筹自付 + cur基数自付 + cur超限自付 + cur大病统筹 + cur大病自付 & "," & _
            cur统筹报销累计 + cur统筹支付 + cur大病统筹 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_重庆壁山 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & ",NULL," & cur基数自付 & "," & _
        cur结算金额 & "," & cur个人现金支付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & cur个人帐户支付 & _
        ",NULL," & lng主页ID & "," & IIf(bln中途结算, "1", "0") & ",'" & strBillNO & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
'        gstrSQL = "zl_结算周期记录_insert(" & lng结帐ID & ",'" & str结算周期 & "'," & cur结算金额 & "," & cur个人帐户支付 & "," & cur统筹支付 & ",'N',to_date('" & datCurr & "','yyyy-mm-dd HH:MI:SS'))"
        gstrSQL = "Insert into zlhis.结算周期记录 values (" & lng结帐ID & ",'" & str结算周期 & "'," & cur结算金额 & "," & cur个人帐户支付 & "," & cur统筹支付 + cur大额支付 & ",'N',to_date('" & Format(datCurr, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
        gcnOracle.Execute gstrSQL
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "西铝厂医保")
    End If
    '清除本次住院请求
    If bln中途结算 = False Then
        strSQL = "Delete from Check_bill_request where bill_no='" & strBillNO & "' and App_code='" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
    End If
    
    住院结算_壁山 = True
    'modify by ccy
    If Val(Get保险参数_壁山("适用地区")) = 1 Then
        Err.Raise 9000, gstrSysName, "中心个人帐户余额为[" & Format(cur中心帐户, "0.00") & "元]", vbInformation, "住院结算"
    End If
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    住院结算_壁山 = False
End Function

Public Function 住院结算冲销_壁山(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, str结算周期 As String
    Dim curDate As Date, strBillNO As String, strSQL As String
        
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select * From 住院费用记录 Where 结帐id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "原单据的费用记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    strBillNO = rsTemp!病人ID & "_" & rsTemp!主页ID
    '删除可能存在的请求
    strSQL = "Delete from Check_bill_request where bill_no='" & strBillNO & "' and App_code='" & Mid(gstr医院编码, 1, 4) & "'"
    gcn壁山.Execute strSQL
    '重新插入请求，以便以后操作
    strSQL = "Insert into Check_bill_request(Bill_no,App_code,Request_status)" & _
            "values('" & strBillNO & "','" & _
            Mid(gstr医院编码, 1, 4) & "','9')"
    gcn壁山.Execute strSQL
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", TYPE_重庆壁山, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
'    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
'    str流水号 = NVL(rsTemp("支付顺序号"), "0")
    
    '帐户年度信息
    Call Get帐户信息(TYPE_重庆壁山, rsTemp("病人ID"), Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & TYPE_重庆壁山 & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_重庆壁山 & "," & rsTemp("病人ID") & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & rsTemp("发生费用金额") * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0,0," & _
        rsTemp("个人帐户支付") * -1 & ",Null," & rsTemp("主页ID") & "," & rsTemp("中途结帐") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
        gstrSQL = "Select * from 结算周期记录 where 结帐id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
        If Not rsTemp.EOF Then
            str结算周期 = rsTemp!结算周期
    '        gstrSQL = "zl_结算周期记录_insert(" & lng结帐ID & ",'" & str结算周期 & "'," & NVL(rsTemp("发生费用金额"), 0) * -1 & "," & NVL(rsTemp("个人帐户支付"), 0) * -1 & "," & NVL(rsTemp("统筹报销金额"), 0) * -1 & ",'N',to_date('" & curDate & "','yyyy-mm-dd HH:MI:SS'))"
    '        Call zlDatabase.ExecuteProcedure(gstrSQL, "西铝厂医保")
            gstrSQL = "Insert into zlhis.结算周期记录 values (" & lng结帐ID & ",'" & str结算周期 & "'," & Nvl(rsTemp("总额"), 0) * -1 & "," & Nvl(rsTemp("个帐"), 0) * -1 & "," & Nvl(rsTemp("统筹"), 0) * -1 & ",'N',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
            gcnOracle.Execute gstrSQL
        End If
    End If

    住院结算冲销_壁山 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 出院登记_壁山(lng病人ID As Long, lng主页ID As Long, Optional ByVal bln转普通病人 As Boolean = False) As Boolean
'功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    '个人状态的修改
    Dim strSQL As String, rs壁山 As New ADODB.Recordset
    Dim strBillNO As String, rsTmp As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim bln零费用出院 As Boolean
    
    On Error GoTo errHandle
    '检查该次住院是否没有费用发生
    If bln转普通病人 = False Then
        gstrSQL = "Select sum(实收金额) as 金额  from 住院费用记录 where 病人ID=[1] and 主页ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人出院", lng病人ID, lng主页ID)
        If rsTemp.EOF = True Then
            bln零费用出院 = True
        Else
            bln零费用出院 = (Nvl(rsTemp("金额"), 0) = 0)
        End If
    Else
        '由保险帐户的撤销医保入院功能调进来，所以不必要求HIS所有费用都已冲销，只是取消医保身份
        bln零费用出院 = True
    End If
    
    If bln零费用出院 = True Then
        '对于零费用出院，就将其处理为退入院。而不用更新其住院信息
        gstrSQL = "select 入院日期 from 病案主页 where 病人ID = " & lng病人ID & _
                " and 主页ID=" & lng主页ID
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", lng病人ID, lng主页ID)
        strBillNO = lng病人ID & "_" & lng主页ID
        '进行出院请求
        strSQL = "Update Check_bill_request set Request_status= '3',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & _
                Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
        '查询请求结果
        If Checkrequest(strBillNO) = False Then 出院登记_壁山 = False: Exit Function
        
        '删除本次的入院登记信息
        strSQL = "Delete from Check_doex_interface where bill_no='" & lng病人ID & "_" & lng主页ID & "' and App_code='" & Mid(gstr医院编码, 1, 4) & "' and Doct_flag=1"
        gcn壁山.Execute strSQL
        strSQL = "Delete from Check_bill_request where bill_no='" & lng病人ID & "_" & lng主页ID & "' and App_code='" & Mid(gstr医院编码, 1, 4) & "'"
        gcn壁山.Execute strSQL
    End If
    '对HIS之中的基础数据进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆壁山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    出院登记_壁山 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    出院登记_壁山 = False
End Function

Public Function 出院登记撤销_壁山(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆壁山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "壁山医保")
    出院登记撤销_壁山 = True
End Function

Public Function Checkrequest(strBillNO As String) As Boolean
'功能：判断是否能够求出正确的病人信息
    Dim strSQL As String, rs壁山 As New ADODB.Recordset
    Dim strResult As String '请求的结果
    Dim strTmp As String, strError As String
    Dim frm等待 As New frm等待响应壁山, lngErrLine As Long
    
    On Error GoTo errHandle
    '提交请求，进行查询
    If frm等待响应壁山.Result(1, strBillNO) = False Then
        Checkrequest = False: lngErrLine = 1
        Unload frm等待响应壁山
        DoEvents
        Exit Function
    End If
    Unload frm等待响应壁山
    '根据返回的返回的值判断结果
    strSQL = "Select Request_Result,Err_text from " & _
            "Check_bill_request where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": lngErrLine = 2
    If rs壁山.State = adStateOpen Then rs壁山.Close: lngErrLine = 3
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly: lngErrLine = 4
    If Not rs壁山.BOF Then
        strTmp = Nvl(rs壁山("Request_Result"), 0): lngErrLine = 5
        strError = Nvl(rs壁山("Err_text"), ""): lngErrLine = 6
    Else
        Exit Function
    End If
    Select Case strTmp
        Case "0"
            Err.Raise 9000, gstrSysName, "没有完成数据请求，请重试", vbInformation, gstrSysName
            Checkrequest = False
            Exit Function
        Case "1"
            If strError <> "" Then
                Err.Raise 9000, gstrSysName, "医保接口调用出现下述错误：" & vbCrLf & vbCrLf & strError, vbInformation, gstrSysName
            Else
                Err.Raise 9000, gstrSysName, "医保接口调用出现错误。", vbInformation, gstrSysName
            End If
            Exit Function
        Case "9"
            Checkrequest = True
    End Select
    Checkrequest = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description & vbCrLf & "在过程[CheckRequest]中第【" & lngErrLine & "】", vbInformation, Err.Source
    Err.Clear
    Checkrequest = False
End Function

Public Function Get保险参数_壁山(ByVal str参数名 As String) As String
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.参数名=[1] and A.险类=[2] and A.中心 is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "壁山医保", str参数名, TYPE_重庆壁山)
    
    If rsTemp.EOF = False Then
        Get保险参数_壁山 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
    End If
End Function

Public Sub SetPos(lHwnd As Long, Optional TopFlag As Boolean = True)
    If TopFlag Then
        SetWindowPos lHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    Else
        SetWindowPos lHwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If
End Sub

Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String, ByVal strFileName As String) As String
    Dim ResultString As String * 144, Temp As Integer, s As String, i As Integer
    Temp% = GetprivateprofileString(SectionName, KeyWord, "", ResultString, 144, strFileName)
    '检索关键词的值
    If Temp% > 0 Then '关键词的值不为空
        s = ""
        For i = 1 To 144
            If asc(Mid$(ResultString, i, 1)) = 0 Then
                Exit For
            Else
                s = s & Mid$(ResultString, i, 1)
            End If
        Next
    Else
        s = DefString
    End If
    GetIniS = Trim(s)
End Function

Private Function IS出院(ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    '判断病人是否已出院
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then
        gstrSQL = " Select B.出院日期 From 病人信息 A,病案主页 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[1]"
    Else
        gstrSQL = " Select B.出院日期 From 病人信息 A,病案主页 B" & _
            " Where A.病人ID=B.病人ID And A.住院次数=B.主页ID And A.病人ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断病人是否已经出院", lng病人ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If IsNull(rsTemp!出院日期) Then Exit Function
    IS出院 = True
End Function
