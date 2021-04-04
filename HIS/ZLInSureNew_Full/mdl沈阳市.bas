Attribute VB_Name = "mdl沈阳市"
Option Explicit
'接口申明区
Public Declare Function CZ_FirstRow Lib "HG_interface.dll" Alias "firstrow" (ByVal pint As Long) As Long
Public Declare Function CZ_NextRow Lib "HG_interface.dll" Alias "nextrow" (ByVal pint As Long) As Long
Public Declare Function CZ_PrevRow Lib "HG_interface.dll" Alias "prevrow" (ByVal pint As Long) As Long
Public Declare Function CZ_LastRow Lib "HG_interface.dll" Alias "lastrow" (ByVal pint As Long) As Long
Public Declare Function CZ_Run Lib "HG_interface.dll" Alias "run" (ByVal pint As Long) As Long
Public Declare Function CZ_NewInterface Lib "HG_interface.dll" Alias "newinterface" () As Long
Public Declare Function CZ_Start Lib "HG_interface.dll" Alias "start" (ByVal pint As Long, ByVal ID As Long) As Long
Public Declare Function CZ_Init Lib "HG_interface.dll" Alias "init" (ByVal pint As Long, ByVal addr As String, ByVal Port As Long, ByVal servlet As String) As Long
Public Declare Function CZ_SetDebug Lib "HG_interface.dll" Alias "setdebug" (ByVal pint As Long, ByVal flag As Integer, ByVal in_direct As String) As Long
Public Declare Function CZ_DataPut Lib "HG_interface.dll" Alias "put" (ByVal pint As Long, ByVal Row As Long, ByVal pname As String, ByVal pvalue As String) As Long
Public Declare Function CZ_GetRowCount Lib "HG_interface.dll" Alias "getrowcount" (ByVal pint As Long) As Long
Public Declare Function CZ_SetRecordset Lib "HG_interface.dll" Alias "setresultset" (ByVal pint As Long, ByVal result_name As String) As Long
Public Declare Function CZ_GetRecordset Lib "HG_interface.dll" Alias "getresultnamebyindex" (ByVal pint As Long, ByVal intIndex As Integer, ByVal result_name As String) As Long
Public Declare Function CZ_GetByName Lib "HG_interface.dll" Alias "getbyname" (ByVal pint As Long, ByVal pname As String, ByVal pvalue As String) As Long
Public Declare Function CZ_GetByIndex Lib "HG_interface.dll" Alias "getbyindex" (ByVal pint As Long, ByVal pindex As Long, ByVal pvalue As String) As Long
Public Declare Function CZ_GetMessage Lib "HG_interface.dll" Alias "getmessage" (ByVal pint As Long, ByVal msg As String) As Long
Public Declare Function CZ_GetException Lib "HG_interface.dll" Alias "getexception" (ByVal pint As Long, ByVal msg As String) As Long
Public Declare Function CZ_SetICCommport Lib "HG_interface.dll" Alias "set_ic_commport" (ByVal pint As Long, ByVal iport As Integer) As Long
Public Declare Sub CZ_DestoryInterface Lib "HG_interface.dll" Alias "destoryinterface" (ByVal pint As Long)

'全局变量区
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
Enum 业务分类_沈阳市                            '用于调接口时，填业务分类
    普通门诊 = 11
    普通住院 = 12
    家庭病床 = 14
    门诊规定病 = 13
    
    门诊急救 = 15
    特治特检 = 16
    生育门诊 = 21
    生育住院 = 22
    工伤门诊 = 41
    工伤住院 = 42
End Enum
Enum 记录集_沈阳市
    MoveFirst = 1
    MoveLast = 2
    MovePrev = 3
    MoveNext = 4
End Enum
Enum Debug_沈阳市
    Normal = 0                                  '正常模式
    Record = 1                                  '调试模式
End Enum
Enum Function_沈阳市
    登录中心 = 0
    '--项目匹配(1001)
    项目匹配_取项目信息 = 100102
    项目匹配_取匹配项目信息 = 100103
    项目匹配_删除匹配信息 = 100104
    项目匹配_重置匹配信息 = 100105
    项目匹配_项目匹配 = 100106
    'Modified By 朱玉宝 地区：长沙 原因：增加功能
    项目匹配_取单个项目匹配信息 = 120507
    '--普通门诊业务费用录入(含改费) (1101)
    普通门诊_身份验证 = 110101                  '要返回个人帐户余额
    普通门诊_收费 = 110104                      '先试算后，再保存
    门诊_查明细 = 110111
    '--普通住院入院登记(1201)
    普通住院_身份验证 = 120101                  '不返回个人帐户余额，在预结算时返回
    普通住院_入院登记 = 120104
    '--普通住院取消登记(1211)
    普通住院_取消入院 = 121104
    '--住院费用(1205)
    住院费用_上传明细 = 120502                  '对应接口的保存明细
    住院费用_查明细 = 120503
    '--住院结算(1202)
    住院结算_预结算 = 120206
    住院结算_正式结算 = 120214
    住院结算_结算冲销 = 120218                  '本次住院期间结算全部回退
    出院登记 = 120204
    '--取消出院(1212)
    取消出院 = 121202
    '--住院信息修改
    住院信息_修改 = 120302
    '--门诊规定病业务费用录入(含改费) (1305)
    门诊规定病_身份验证 = 130501
    门诊规定病_收费 = 130504
    '--其他事务
    其他_读卡 = 200900
    其他_修改密码 = 200910
    其他_个人信息 = 200001
    其他_基金余额 = 200004                      '按传入的基金编号，返回基金余额（个人帐户，医保基金...）
    其他_黑名单校验 = 200009
    其他_整张单据冲销 = 112010
    其他_获取发票信息 = 200040
    '--结算单
    结算单_住院 = 200030
    结算单_门诊 = 200031
    结算单_门诊规定病 = 200032
    '--结算汇总表（对帐单）
    结算汇总表_住院 = 200035
    结算汇总表_门诊 = 200037
    结算汇总表_门诊规定病 = 200038
    '--转院申请
    转院申请_病人信息 = 121002
    转院申请_校验信息 = 121003
    转院申请_保存转院申请 = 121005
    转院申请_查询审核信息 = 121006
    获取医院信息 = 121007
End Enum

Private Type ComInfo_沈阳市
    医院编码 As String
    操作员工号 As String
    业务类型 As String
    个人编号 As String
    业务序列号 As String
    帐户余额 As Currency
    总费用 As Currency
    疾病编码 As String                      '保存身份验证后返回的疾病编码
End Type
Public gCominfo_沈阳市 As ComInfo_沈阳市

Private Const 住院次数累计 = 0
Private Const 入院科室编号 = 1
Private Const 入院科室名称 = 2
Private Const 入院病区编号 = 3
Private Const 入院病区名称 = 4
Private Const 入院病床编号 = 5
Private Const 床位类型 = 6
Private Const 住院号 = 7
Private Const 入院诊断 = 8
Private Const 出院诊断 = 9

Private Const mintTest As Integer = 0           '试算，就是预结算（住院预结算与结算是分开的，门诊使用）
Private Const mintICCard As Integer = 1         '使用IC卡
Private Const strDebug_Path As String = "C:\Log" '保存调试信息的文件夹,医保初始化的创建

Public glngInterface_沈阳市 As Long             '连接串
Public glngReturn_沈阳市 As Long                '接口返回值
Public gstrErrInfo_沈阳市 As String             '错误信息保存区域
Public gstrField_沈阳市 As String
Public gstrValue_沈阳市 As String
Private mintICPort As Integer                   'IC设备的端口号
Public mint适用地区_沈阳 As Integer             '1-长春地区;2-沈阳地区
Private mstrAddress As String, mstrPort As String, mstrServlet As String

Public Sub ErrInformation()
    If glngReturn_沈阳市 < 0 Then
        glngReturn_沈阳市 = CZ_GetMessage(glngInterface_沈阳市, gstrErrInfo_沈阳市)
        MsgBox gstrErrInfo_沈阳市, vbInformation, gstrSysName
    End If
End Sub

Public Function 医保初始化_沈阳市() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim objFileSys As FileSystemObject
    On Error GoTo errHand
    
    If glngInterface_沈阳市 = 0 Then
        '进行医保接口的初始化工作
        gstrErrInfo_沈阳市 = Space(1000)
        If Not GetServerInfo Then Exit Function
        glngInterface_沈阳市 = CZ_NewInterface()
        glngReturn_沈阳市 = CZ_Init(glngInterface_沈阳市, mstrAddress, mstrPort, mstrServlet)
        Call ErrInformation
        If glngReturn_沈阳市 = -1 Then
            '启动失败
            Call 医保终止_沈阳市
            Exit Function
        End If
        '设置IC设备的端口号
        Call CZ_SetICCommport(glngInterface_沈阳市, mintICPort)
        '设置调试目录
        Set objFileSys = New FileSystemObject
        If Not objFileSys.FolderExists(strDebug_Path) Then
            objFileSys.CreateFolder (strDebug_Path)
        End If
        Call CZ_SetDebug(glngInterface_沈阳市, Debug_沈阳市.Record, strDebug_Path)
        '登录中心(失败则断开连接并退出)
        If Not frm登录中心.LoginCenter(TYPE_沈阳市) Then
            Call 医保终止_沈阳市
            Exit Function
        End If
        gCominfo_沈阳市.操作员工号 = Right(UserInfo.编号, 5)
        '取医院编码
        gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_沈阳市)
        gCominfo_沈阳市.医院编码 = Nvl(rsTemp!医院编码)
        
        '取适用地区
        mint适用地区_沈阳 = 0
        gstrSQL = "Select 参数值 From 保险参数 Where 参数名='适用地区' And 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取适用地区", TYPE_沈阳市)
        If Not rsTemp.EOF Then
            mint适用地区_沈阳 = Nvl(rsTemp!参数值, 0)
        End If
    End If
    
    医保初始化_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保终止_沈阳市() As Boolean
    If glngInterface_沈阳市 = 0 Then
        医保终止_沈阳市 = True
        Exit Function
    End If
    
    Call CZ_DestoryInterface(glngInterface_沈阳市)
    glngInterface_沈阳市 = 0
    
    医保终止_沈阳市 = True
End Function

Public Function 医保设置_沈阳() As Boolean
    医保设置_沈阳 = frmSet沈阳.ShowME
End Function

Private Function GetServerInfo() As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '获取IC设备的端口号
    mintICPort = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", 1)
    
    '获取服务器地址、端口及入口名称('服务器地址','服务器端口号','服务器入口程序')
    gstrSQL = " Select 参数名,参数值 From 保险参数" & _
              " Where 险类=[1] And 参数名 Like '服务器%'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取服务器地址、端口及入口名称", TYPE_沈阳市)
    
    With rsTemp
        Do While Not .EOF
            Select Case !参数名
            Case "服务器地址"
                mstrAddress = Nvl(!参数值)
            Case "服务器端口号"
                mstrPort = Nvl(!参数值)
            Case "服务器入口程序"
                mstrServlet = Nvl(!参数值)
            End Select
            .MoveNext
        Loop
    End With
    
    GetServerInfo = Not (mstrAddress = "" Or mstrPort = "" Or mstrServlet = "")
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 调用接口_准备_沈阳市(ByVal lng功能 As Long) As Boolean
    On Error GoTo errHand
    
    glngReturn_沈阳市 = CZ_Start(glngInterface_沈阳市, lng功能)
    Call ErrInformation
    
    'Modified By 朱玉宝 长沙    <=0
    If glngReturn_沈阳市 < 0 Then Exit Function
    
    调用接口_准备_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 调用接口_执行_沈阳市() As Boolean
    On Error GoTo errHand
    
    glngReturn_沈阳市 = CZ_Run(glngInterface_沈阳市)
    Call ErrInformation
    'Modified By 朱玉宝 长沙    <=0
    If glngReturn_沈阳市 < 0 Then Exit Function
    
    调用接口_执行_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 调用接口_指定记录集_沈阳市(ByVal strRecordName As String) As Boolean
    On Error GoTo errHand
    
    glngReturn_沈阳市 = CZ_SetRecordset(glngInterface_沈阳市, strRecordName)
    Call ErrInformation
    'Modified By 朱玉宝 长沙    <=0
    If glngReturn_沈阳市 < 0 Then Exit Function
    
    调用接口_指定记录集_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 调用接口_记录数_沈阳市() As Boolean
    Dim lngRecord As Long
    On Error GoTo errHand
    
    lngRecord = CZ_GetRowCount(glngInterface_沈阳市)
    调用接口_记录数_沈阳市 = (lngRecord > 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 调用接口_移动记录集_沈阳市(ByVal intType As 记录集_沈阳市) As Boolean
    On Error Resume Next
    
    Err = 0
    Select Case intType
    Case 记录集_沈阳市.MoveFirst
        glngReturn_沈阳市 = CZ_FirstRow(glngInterface_沈阳市)
    Case 记录集_沈阳市.MovePrev
        glngReturn_沈阳市 = CZ_PrevRow(glngInterface_沈阳市)
    Case 记录集_沈阳市.MoveNext
        glngReturn_沈阳市 = CZ_NextRow(glngInterface_沈阳市)
    Case 记录集_沈阳市.MoveLast
        glngReturn_沈阳市 = CZ_LastRow(glngInterface_沈阳市)
    End Select
    
    If Err <> 0 Then Exit Function
    调用接口_移动记录集_沈阳市 = Not (glngReturn_沈阳市 < 0)
    Exit Function
End Function

Public Function 调用接口_读取数据_沈阳市(ByVal strField As String, strValue As String) As Boolean
    On Error GoTo errHand
    
    strValue = Space(1000)
    Call DebugTool("取字段：" & strField)
    glngReturn_沈阳市 = CZ_GetByName(glngInterface_沈阳市, strField, strValue)
    Call ErrInformation
    If glngReturn_沈阳市 <= 0 Then Exit Function
    'Modified By 朱玉宝 地区：长沙 原因：加上去掉0字符的Replace语句
    strValue = Trim(Replace(strValue, Chr(0), ""))
    
    调用接口_读取数据_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 调用接口_写入口参数_沈阳市(ByVal lngRow As Long) As Boolean
    Dim intField As Integer, intCOUNT As Integer
    Dim arrField, arrData
    Dim blnErr As Boolean
    On Error GoTo errHand
    
    arrField = Split(gstrField_沈阳市, "||")
    'Modified By 朱玉宝 地区：长沙 原因：arrData写成了arrField
    arrData = Split(gstrValue_沈阳市, "||")
    intCOUNT = UBound(arrField)
    For intField = 0 To intCOUNT
        glngReturn_沈阳市 = CZ_DataPut(glngInterface_沈阳市, lngRow, arrField(intField), arrData(intField))
        If Not blnErr Then
            blnErr = (glngReturn_沈阳市 <= 0)
        End If
    Next
    
    'Modified By 朱玉宝 地区：长沙 原因：Exit Function应该放在最后
    调用接口_写入口参数_沈阳市 = Not blnErr
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 挂号结算_沈阳市(ByVal lng结帐ID As Long) As Boolean
    Dim str结算方式 As String
    Dim arr结算方式, intStart As Integer, intTotal As Integer, lng收入项目 As Long
    Dim cur个人帐户 As Currency, cur费用总额 As Currency '上传费用总额
    Dim rsTemp As New ADODB.Recordset
    '先调门诊预结算,主要是取个人帐户支付额,再调门诊结算
    '保险参数中保存的是允许帐户支付的收入项目ID，挂号结算时判断，如果未设置，表明全自付，不需上传，否则仅上传那一笔明细
    
    On Error GoTo errHand
    
    '如果是长春地区，由于挂号全由现金结算，中心要求不上传挂号明细
    If mint适用地区_沈阳 = 1 Then '长春地区
        挂号结算_沈阳市 = True
        Exit Function
    End If
    
    '先提取保险参数
    Call DebugTool("提取收入项目")
    gstrSQL = "Select 参数值 Value From 保险参数 Where 险类=[1] And 参数名='个人帐户支出(挂号)'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取由个人帐户支付的收入项目ID", TYPE_沈阳市)
    If Not rsTemp.EOF Then
        lng收入项目 = Nvl(rsTemp!Value, 0)
        Call DebugTool("收入项目ID:" & lng收入项目)
    End If
    
    '先取本次结算的费用明细,并调用虚拟结算获取预算结果
    gstrSQL = "Select Rownum 标识号,A.ID,A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,A.登记时间,A.开单人," & _
            "   A.收费细目ID,A.收入项目ID,A.数次*A.付数 as 数量,A.计算单位,Round(A.结帐金额/(A.数次*A.付数),4) as 单价,A.实收金额," & _
            "   A.收费类别,B.编码 as 项目编码,B.名称 as 项目名称,B.规格,D.项目编码 医保编码," & _
            "   C.名称 开单部门,E.名称 受单部门" & _
            " From (Select * From 门诊费用记录 Where 结帐ID=" & lng结帐ID & ") A,收费细目 B,部门表 C,保险支付项目 D,部门表 E " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID(+) And A.执行部门ID=E.ID(+) And A.收费细目ID=D.收费细目ID And D.险类=[1]" & _
            " Order by A.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取本次结帐费用明细", TYPE_沈阳市)
    With rsTemp
        Do While Not .EOF
            cur费用总额 = cur费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    If Not 门诊虚拟结算_沈阳市(rsTemp, str结算方式, True, lng收入项目) Then Exit Function
    
    '分解结算方式,获取个人帐户支付额及基金支付额（"报销方式;金额;是否允许修改|...."）
    arr结算方式 = Split(str结算方式, "|")
    intTotal = UBound(arr结算方式)
    For intStart = 0 To intTotal
        If Split(arr结算方式(intStart), ";")(0) = "个人帐户" Then
            cur个人帐户 = Val(Split(arr结算方式(intStart), ";")(1))
            Exit For
        End If
    Next
    
    Call DebugTool("上传费用总额:" & cur费用总额 & "|返回个人帐户支付额:" & cur个人帐户)
    '再调用门诊结算（strSelfNO因为没有使用，所以传空）
    If Not 门诊结算_沈阳市(lng结帐ID, cur个人帐户, "", True, lng收入项目) Then Exit Function
    
    '修改病人预交记录（根据现金结算方式产生相应的记录）
    If lng收入项目 <> 0 Then
        '产生预交记录
        For intStart = 0 To intTotal
            If Split(arr结算方式(intStart), ";")(0) <> "现金" Then
                gstrSQL = " insert into 病人预交记录(ID,记录性质,NO,记录状态,病人ID,主页ID,科室ID,缴款单位," & _
                         " 单位开户行,单位帐号,摘要,金额,结算方式,结算号码,收款时间,操作员编号,操作员姓名,冲预交,结帐ID) " & _
                         " select 病人预交记录_ID.nextval ID,记录性质,NO,记录状态,病人ID,主页ID,科室ID, " & _
                         " 缴款单位,单位开户行,单位帐号,摘要,金额,'" & Split(arr结算方式(intStart), ";")(0) & "',结算号码,收款时间,操作员编号, " & _
                         " 操作员姓名," & Val(Split(arr结算方式(intStart), ";")(1)) & ",结帐ID " & _
                         " from 病人预交记录" & _
                         " Where 结帐ID=" & lng结帐ID & " And 结算方式='现金' And not (记录性质=1 or 记录性质=11)"
                         cur费用总额 = cur费用总额 - Val(Split(arr结算方式(intStart), ";")(1))
                gcnOracle.Execute gstrSQL
            End If
        Next
        
        '修正现金支付额
        If cur费用总额 <> 0 Then
            '修改现金支付额
            gstrSQL = " Update 病人预交记录 Set 冲预交= " & cur费用总额 & _
                      " Where 结帐ID=" & lng结帐ID & " And 结算方式='现金' And not (记录性质=1 or 记录性质=11)"
        Else
            '无现金支付额，删除该预交记录
            gstrSQL = " Delete 病人预交记录 " & _
                      " Where 结帐ID=" & lng结帐ID & " And 结算方式='现金' And not (记录性质=1 or 记录性质=11)"
        End If
        gcnOracle.Execute gstrSQL
    End If
    
    Call frm结算信息.ShowME(lng结帐ID)
    挂号结算_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 挂号冲销_沈阳市(ByVal lng结帐ID As Long) As Boolean
    Dim cur个人帐户 As Currency, lng病人ID As Long, lng记录ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '取被冲销记录的结帐ID，病人ID
    gstrSQL = "select distinct A.结帐ID,A.病人ID from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng记录ID = rsTemp!结帐ID
    lng病人ID = rsTemp!病人ID
    
    '读取原始记录的个人帐户支付额（因为门诊结算冲销时未使用参数“个人帐户”，可以不取）
    gstrSQL = "Select Nvl(A.冲预交,0) 个人帐户 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=[2]" & _
        " And A.结算方式 in ('个人帐户') And A.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取原始单据的个人帐户支付额", lng记录ID, TYPE_沈阳市)
    cur个人帐户 = 0
    If Not rsTemp.EOF Then
        cur个人帐户 = rsTemp!个人帐户
    End If
    
    If Not 门诊结算冲销_沈阳市(lng结帐ID, cur个人帐户, lng病人ID) Then Exit Function
    挂号冲销_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 门诊虚拟结算_沈阳市(rs明细 As ADODB.Recordset, str结算方式 As String, _
Optional ByVal bln挂号 As Boolean = False, Optional ByVal lng收入项目ID As Long = 0) As Boolean
    Dim lng功能 As Long, str个人编号 As String
    Dim str基金名称 As String, str基金编号 As String, str支付金额 As String
    Dim str就诊时间 As String, str发生日期 As String
    Dim str医生编号 As String, str医生姓名 As String
    Dim str规格 As String, str产地 As String, str剂型 As String, str项目编码 As String
    Dim str科室号 As String, str科室名称 As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：开单人,病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '挂号结算也使用本过程，如果是现金支付的项目，需要传特定的医院编码和医保编码 110100001 挂号费
    On Error GoTo errHand
    
    '--读IC卡
    '读出IC卡中的信息
    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_读卡) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    '取返回的记录集
    'If Not 调用接口_指定记录集_沈阳市("ICInfo") Then Exit Sub
    If Not 调用接口_读取数据_沈阳市("indi_id", str个人编号) Then Exit Function
    
    '校验个人编号是否正确
    gstrSQL = "Select 卡号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "校验该病人是否使用自己的卡进行交易", TYPE_沈阳市, CLng(rs明细!病人ID))
    If rsTemp!卡号 <> str个人编号 Then
        MsgBox "该病人不是使用自己的医保卡，结算操作中止！", vbInformation, gstrSysName
        Exit Function
    End If
    
    str就诊时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str发生日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    lng功能 = IIf(gCominfo_沈阳市.业务类型 = 业务分类_沈阳市.门诊规定病, _
                    Function_沈阳市.门诊规定病_收费, Function_沈阳市.普通门诊_收费)
    If Not 调用接口_准备_沈阳市(lng功能) Then Exit Function
    
    '取医生姓名
    Call DebugTool("取医生姓名")
    str医生编号 = "": str医生姓名 = ""
    If bln挂号 Then
        Call DebugTool("挂号-取医生姓名")
        gstrSQL = "select 医生姓名 from 挂号安排 where 号码=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生编号", CStr(rs明细!计算单位))
        str医生姓名 = Nvl(rsTemp!医生姓名)
    Else
        Call DebugTool("门诊-取医生姓名")
        str医生姓名 = Nvl(rs明细!开单人)
    End If
    
    Call DebugTool("取医生编号")
    If Trim(str医生姓名) <> "" Then
        gstrSQL = "Select 编号 From 人员表 Where 姓名=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生编号", str医生姓名)
        str医生编号 = rsTemp!编号
    End If
    
    If str医生姓名 <> "" And mint适用地区_沈阳 = 2 Then
        Call DebugTool("取科室编号与名称")
        gstrSQL = " Select 编码,名称 From 部门表 " & _
                  " Where ID in " & _
                  "     (Select 部门ID From 部门人员 " & _
                  "     Where 人员ID=" & _
                  "         (Select ID From 人员表 Where 姓名=[1]))"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取科室编号与名称", str医生姓名)
        If rsTemp.RecordCount <> 0 Then
            str科室号 = rsTemp!编码
            str科室名称 = rsTemp!名称
        End If
    End If
    
    Call DebugTool("开始上传明细")
    With rs明细
        '写入口参数
'        1   hospital_id    医疗机构编码   20   否
'        2   indi_id        个人编号        8   否
'        3   busi_type      业务类型        2   否  "11"：门诊
'        4   ic_flag        用卡标志        1   是  "0"：不使用IC卡；"1"：使用IC卡
'        5   reg_staff      登记人员工号    5   否
'        6   reg_man        登记人姓名      10  否
'        7   begin_date     就诊时间            否  格式：YYYY-MM-DD HH:MI:SS(24小时)
'        8   in_disease     登记诊断        20  否  疾病编码
'        9   calcSaveFlag   计算保存标志    1   是  "0"：试算；"1"：收费
'        10  accMoney       个人帐户支付金额18  否
'        11  recipe_no      处方号          20  是
'--------2004-01-12增加--------
'        12  doctor_no      处方医生编号    12
'        13  doctor_name    处方医生姓名    10
'------------------------------
'        14  note           备注            100 是
'以下两个参数仅沈阳地区需要------------------------------
'        15  in_dept        科室号          10
'        16  in_dept_name   科室名称        20
        gstrField_沈阳市 = "hospital_id||indi_id||busi_type||ic_flag||reg_staff||" & _
                    "reg_man||begin_date||in_disease||calcSaveFlag||accMoney||recipe_no||doctor_no||doctor_name||note"
        If mint适用地区_沈阳 = 2 Then gstrField_沈阳市 = gstrField_沈阳市 & "||in_dept||in_dept_name"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.个人编号 & "||" & _
                    gCominfo_沈阳市.业务类型 & "||" & mintICCard & "||" & gCominfo_沈阳市.操作员工号 & "||" & _
                    gstrUserName & "||" & str就诊时间 & "||" & _
                    gCominfo_沈阳市.疾病编码 & "||" & mintTest & "||0||||" & str医生编号 & "||" & str医生姓名 & "||"
        If mint适用地区_沈阳 = 2 Then gstrValue_沈阳市 = gstrValue_沈阳市 & "||" & str科室号 & "||" & str科室名称
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
        
        '设置对应的记录集，准备传明细
        If Not 调用接口_指定记录集_沈阳市("FeeInfo") Then Exit Function
        '写单据明细
        gCominfo_沈阳市.总费用 = 0
        Do While Not .EOF
'            1   medi_item_type 项目药品类型        1   否  "0"：诊疗项目；"1"：西药；"2"：中成药；"3"：中草药
'            2   his_item_code  医院药品项目编码    20  否
'            3   his_item_name  医院药品项目名称    50  否
'            4   model          剂型                30  是
'            5   factory        厂家                50  是
'            6   standard       规格                30  是
'            7   fee_date       费用发生时间            否  格式：YYYY-MM-DD
'            8   unit           计量单位            10  是
'            9   price          单价                12  否
'            10  dosage         用量                12  是
'            11  money          金额                12  否
'            12  opp_serial_fee 对应费用序列号      12  是
            
            '如果是药品，再去取剂型
            str剂型 = ""
            str项目编码 = ""
            If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                gstrSQL = " Select C.标识码,C.编码,B.名称 剂型 From 药品信息 A,药品剂型 B,药品目录 C " & _
                          " Where A.剂型=B.编码 And A.药名ID=C.药名ID And C.药品ID =[1]"
                Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "取剂型", CLng(!收费细目ID))
                str剂型 = Nvl(rsPhysic!剂型)
                str项目编码 = Nvl(rsPhysic!标识码)
                If str项目编码 = "" Then str项目编码 = Nvl(rsPhysic!编码)
            End If
            
            '取规格和产地
            gstrSQL = "Select 编码,名称,规格,标识主码 From 收费细目 Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取收费细目相关信息", CLng(!收费细目ID))
            str产地 = ""
            str规格 = Nvl(rsTemp!规格)
            If InStr(1, str规格, "┆") <> 0 Then
                str产地 = ToVarchar(Trim(Split(str规格, "┆")(1)), 50)
                str规格 = ToVarchar(Trim(Split(str规格, "┆")(0)), 30)
            Else
                str规格 = ToVarchar(Trim(str规格), 30)
            End If
            
            '如果是除药品外项目，取其编码则取标识主码
            If Not (!收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7) Then
                str项目编码 = Nvl(rsTemp!标识主码)
                If str项目编码 = "" Then str项目编码 = Nvl(rsTemp!编码)
            End If
            
            '如果是挂号，且收入项目ID<>传入收入项目ID，则传固定的医院编码和医保编码
            If bln挂号 Then
                If rs明细!收入项目id <> lng收入项目ID Then
                    str项目编码 = "110100001"
                End If
            End If
            
            gstrField_沈阳市 = "medi_item_type||his_item_code||his_item_name||model||factory||" & _
                        "standard||fee_date||unit||price||dosage||money||opp_serial_fee||hos_serial"
            gstrValue_沈阳市 = IIf(!收费类别 = 5, "1", IIf(!收费类别 = 6, "2", IIf(!收费类别 = 7, "3", "0"))) & "||" & _
                        str项目编码 & "||" & Nvl(rsTemp!名称) & "||" & str剂型 & "||" & str产地 & "||" & str规格 & "||" & _
                        str发生日期 & "||" & Nvl(!计算单位) & "||" & !单价 & "||" & !数量 & "||" & !实收金额 & "||||"
            
            If .AbsolutePosition <> 1 Then
                If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Function
            End If
            If Not 调用接口_写入口参数_沈阳市(.AbsolutePosition) Then Exit Function
            
            gCominfo_沈阳市.总费用 = gCominfo_沈阳市.总费用 + !实收金额
            .MoveNext
        Loop
    End With
    
    If Not 调用接口_执行_沈阳市() Then Exit Function
    
    '获取计算后的各项基金支付额
    If Not 调用接口_指定记录集_沈阳市("BizInfo") Then Exit Function
'    1   fund_id;    基金编码    3
'    2   fund_name   基金名称    30
'    3   real_pay    支付金额    12
'    4   serial_no   业务序列号  12
    If 调用接口_记录数_沈阳市 Then
        Do While True
            Call 调用接口_读取数据_沈阳市("fund_id", str基金编号)
            Call 调用接口_读取数据_沈阳市("fund_name", str基金名称)
            Call 调用接口_读取数据_沈阳市("real_pay", str支付金额)
            If str基金编号 <> 999 And Val(str支付金额) <> 0 Then
                str结算方式 = str结算方式 & IIf(str结算方式 = "", "", "|") & _
                            IIf(str基金编号 = "003", "个人帐户", str基金名称) & ";" & str支付金额 & ";0"
            End If
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    If str结算方式 = "" Then str结算方式 = "个人帐户;0;0"
    门诊虚拟结算_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 门诊结算_沈阳市(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, _
Optional ByVal bln挂号 As Boolean = False, Optional ByVal lng收入项目ID As Long = 0) As Boolean
    Dim lng功能 As Long, lng病人ID As Long
    Dim str基金名称 As String, str基金编号 As String, str支付金额 As String
    Dim str业务序列号 As String, str就诊时间 As String, str发生日期 As String
    Dim cur统筹基金 As Currency, cur现金 As Currency
    Dim str医生编号 As String, str医生姓名 As String, str项目编码 As String
    Dim str规格 As String, str产地 As String, str剂型 As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传，如果卡内金额不足，可以使用圈存接口，即将卡外的钱，调到卡内，以增加卡内金额
    '卡内余额需要通过卡操作函数读取，可圈存金额是接口返回，需要修改
    On Error GoTo errHand
    str就诊时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str发生日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    lng功能 = IIf(gCominfo_沈阳市.业务类型 = 业务分类_沈阳市.门诊规定病, _
                    Function_沈阳市.门诊规定病_收费, Function_沈阳市.普通门诊_收费)
    If Not 调用接口_准备_沈阳市(lng功能) Then Exit Function
    
    '取医生编号与姓名
    Call DebugTool("取医生编号与姓名")
    str医生编号 = "": str医生姓名 = ""
    
    If bln挂号 Then
        gstrSQL = "select 医生姓名 开单人 from 挂号安排 where 号码=(Select 计算单位 From 门诊费用记录 Where 结帐ID=[1] And Rownum<2)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生编号", lng结帐ID)
    Else
        gstrSQL = "Select 开单人 From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取开单医生", lng结帐ID)
    End If
    str医生姓名 = Nvl(rsTemp!开单人)

    If Trim(str医生姓名) <> "" Then
        gstrSQL = "Select 编号,姓名 From 人员表 Where 姓名=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生编号", str医生姓名)
        str医生编号 = rsTemp!编号
    End If
    
    '上传费用明细记录
    Call DebugTool("上传费用明细记录")
    gstrSQL = "Select Rownum 标识号,A.ID,A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,A.登记时间,A.开单人 as 医生," & _
            "   A.收费细目ID,A.收入项目ID,A.计算单位,A.数次*A.付数 as 数量,Round(A.结帐金额/(A.数次*A.付数),4) as 单价,A.实收金额 金额," & _
            "   A.收费类别,B.标识主码,B.编码 as 编码,B.名称 as 项目名称,B.规格,D.项目编码 医保编码," & _
            "   C.名称 开单部门,E.名称 受单部门" & _
            " From (Select * From 门诊费用记录 Where 结帐ID=[2] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0) A,收费细目 B,部门表 C,保险支付项目 D,部门表 E " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID(+) And A.执行部门ID=E.ID(+) And A.收费细目ID=D.收费细目ID And D.险类=[1]" & _
            " Order by A.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取本次结帐费用明细", TYPE_沈阳市, lng结帐ID)
    lng病人ID = rsTemp!病人ID
    With rsTemp
        '写入口参数
        Call DebugTool("上传费用明细记录-单据头")
        gstrField_沈阳市 = "hospital_id||indi_id||busi_type||ic_flag||reg_staff||" & _
                    "reg_man||begin_date||in_disease||calcSaveFlag||accMoney||recipe_no||doctor_no||doctor_name||note"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.个人编号 & "||" & _
                    gCominfo_沈阳市.业务类型 & "||" & mintICCard & "||" & gCominfo_沈阳市.操作员工号 & "||" & _
                    gstrUserName & "||" & str就诊时间 & "||" & _
                    gCominfo_沈阳市.疾病编码 & "||1||" & cur个人帐户 & "||" & !NO & "||" & str医生编号 & "||" & str医生姓名 & "||"
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
        
        '设置对应的记录集，准备传明细
        If Not 调用接口_指定记录集_沈阳市("FeeInfo") Then Exit Function
        
        Call DebugTool("上传费用明细记录-单据体")
        Do While Not .EOF
            
            '如果是药品，再去取剂型
            Call DebugTool("上传费用明细记录-单据体-取药品剂型，医保编码")
            str剂型 = ""
            str项目编码 = ""
            If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                gstrSQL = " Select B.名称 剂型,C.编码,C.标识码 From 药品信息 A,药品剂型 B,药品目录 C " & _
                          " Where A.剂型=B.编码 And A.药名ID=C.药名ID And C.药品ID =[1]"
                Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "取剂型", CLng(!收费细目ID))
                str剂型 = Nvl(rsPhysic!剂型)
                str项目编码 = Nvl(rsPhysic!标识码)
                If str项目编码 = "" Then str项目编码 = Nvl(rsPhysic!编码)
            Else
                str项目编码 = Nvl(!标识主码)
                If str项目编码 = "" Then str项目编码 = Nvl(!编码)
            End If
            
            Call DebugTool("上传费用明细记录-单据体-取规格、产地")
            str产地 = ""
            str规格 = Nvl(!规格)
            If InStr(1, str规格, "┆") <> 0 Then
                str产地 = ToVarchar(Trim(Split(str规格, "┆")(1)), 50)
                str规格 = ToVarchar(Trim(Split(str规格, "┆")(0)), 30)
            Else
                str规格 = ToVarchar(Trim(str规格), 30)
            End If
            
            '如果是挂号，且收入项目ID<>传入收入项目ID，则传固定的医院编码和医保编码
            Call DebugTool("上传费用明细记录-单据体-如果是挂号还要重新取特定的项目编码")
            If bln挂号 Then
                If rsTemp!收入项目id <> lng收入项目ID Then
                    str项目编码 = "110100001"
                End If
            End If
            
            gstrField_沈阳市 = "medi_item_type||his_item_code||his_item_name||model||factory||" & _
                        "standard||fee_date||unit||price||dosage||money||opp_serial_fee||hos_serial"
            gstrValue_沈阳市 = IIf(!收费类别 = 5, "1", IIf(!收费类别 = 6, "2", IIf(!收费类别 = 7, "3", "0"))) & "||" & _
                        str项目编码 & "||" & Nvl(!项目名称) & "||" & str剂型 & "||" & str产地 & "||" & str规格 & "||" & _
                        str发生日期 & "||" & Nvl(!计算单位) & "||" & !单价 & "||" & !数量 & "||" & !金额 & "||||" & !ID
            
            If .AbsolutePosition <> 1 Then
                Call DebugTool("上传费用明细记录-单据体-MOVENEXT")
                If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Function
            End If
            Call DebugTool("上传费用明细记录-单据体-写入口参数")
            If Not 调用接口_写入口参数_沈阳市(.AbsolutePosition) Then Exit Function
            '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
            'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
            Call DebugTool("上传费用明细记录-单据体-打上传标志")
            gstrSQL = "zl_病人费用记录_上传('" & rsTemp("NO") & "'," & rsTemp("序号") & "," & rsTemp("记录性质") & "," & rsTemp("记录状态") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            
            .MoveNext
        Loop
    End With
    
    If Not 调用接口_执行_沈阳市() Then Exit Function
    
    '获取计算后的各项基金支付额
    Call DebugTool("取结算结果")
    If Not 调用接口_指定记录集_沈阳市("BizInfo") Then Exit Function
'    1   fund_id;    基金编码    3
'    2   fund_name   基金名称    30
'    3   real_pay    支付金额    12
'    4   serial_no   业务序列号  12
    
    If 调用接口_记录数_沈阳市 Then
        cur统筹基金 = 0
        cur现金 = 0
        
        Do While True
            Call 调用接口_读取数据_沈阳市("fund_id", str基金编号)
            Call 调用接口_读取数据_沈阳市("fund_name", str基金名称)
            Call 调用接口_读取数据_沈阳市("real_pay", str支付金额)
            Call 调用接口_读取数据_沈阳市("serial_no", str业务序列号)
            
            Select Case str基金编号
            Case "003"
                cur个人帐户 = Val(str支付金额)
            Case Is >= "900"
                cur现金 = cur现金 + Val(str支付金额)
            Case Else
                cur统筹基金 = cur统筹基金 + Val(str支付金额)
            End Select
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If

    '填写结算记录
    '帐户累计增加=公务员补助基金;帐户累计支出=公务员个人补助基金
    '累计进入统筹=个人补充基金;累计统筹报销=医疗补偿金
    Call DebugTool("写保险结算记录")
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_沈阳市 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gCominfo_沈阳市.总费用 & "," & cur现金 & "," & 0 & "," & cur统筹基金 & "," & cur统筹基金 & ",0," & _
        0 & "," & cur个人帐户 & ",'" & str业务序列号 & "',null,null," & gCominfo_沈阳市.业务类型 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊收费数据")
    
    gCominfo_沈阳市.业务序列号 = str业务序列号
    门诊结算_沈阳市 = True
    
    '20031228:周韬:加入结帐ID
    Call DebugTool("取发票信息")
    Call GetBalance(lng病人ID, lng结帐ID, str业务序列号, gCominfo_沈阳市.医院编码)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_沈阳市(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, Optional ByVal bln退费 As Boolean = True) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Dim str业务序列号 As String, str业务类型 As String, str个人编号 As String
    Dim lng记录ID As Long                               '冲销记录的结帐ID
    Dim int性质 As Integer, int状态 As Integer, str单据号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '--读IC卡
    '读出IC卡中的信息
    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_读卡) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    '取返回的记录集
    'If Not 调用接口_指定记录集_沈阳市("ICInfo") Then Exit Sub
    If Not 调用接口_读取数据_沈阳市("indi_id", str个人编号) Then Exit Function
    
    '校验个人编号是否正确
    gstrSQL = "Select 卡号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "校验该病人是否使用自己的卡进行交易", TYPE_沈阳市, lng病人ID)
    If rsTemp!卡号 <> str个人编号 Then
        Err.Raise 9000, gstrSysName, "该病人不是使用自己的医保卡，结算操作中止！"
        Exit Function
    End If
    
    Call 获取病人基本信息(lng病人ID)
    '取被冲销记录的结帐ID
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng记录ID = rsTemp!结帐ID
    
    '取原结算记录的明细
    gstrSQL = "Select * From 保险结算记录 Where 险类=[1] And 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取结算记录", TYPE_沈阳市, lng结帐ID)
    str业务序列号 = rsTemp!支付顺序号
    str业务类型 = rsTemp!备注
    
    '产生冲销结算记录
    With rsTemp
        gstrSQL = "zl_保险结算记录_insert(" & !性质 & "," & lng记录ID & "," & TYPE_沈阳市 & "," & !病人ID & "," & _
            !年度 & "," & -1 * Nvl(!帐户累计增加, 0) & "," & -1 * Nvl(!帐户累计支出, 0) & "," & -1 * Nvl(!累计进入统筹, 0) & "," & -1 * Nvl(!累计统筹报销, 0) & ",NULL,0,0,0," & _
            -1 * Nvl(!发生费用金额, 0) & "," & -1 * Nvl(!全自付金额, 0) & "," & -1 * Nvl(!首先自付金额, 0) & "," & -1 * Nvl(!进入统筹金额, 0) & "," & -1 * Nvl(!统筹报销金额, 0) & ",0," & _
            0 & "," & -1 * Nvl(!个人帐户支付, 0) & ",'" & str业务序列号 & "',null,null," & str业务类型 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "产生冲销结算记录")
    End With
    
    If bln退费 Then
        Call DebugTool("门诊退费")
        '调用整张单据冲销接口
        If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_整张单据冲销) Then Exit Function
        '写入口参数
        gstrField_沈阳市 = "hospital_id||serial_no||indi_id||busi_type||ic_flag||staff_no||staff_man"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & str业务序列号 & "||" & gCominfo_沈阳市.个人编号 & "||" & gCominfo_沈阳市.业务类型 & "||" & mintICCard & "||" & Right(UserInfo.编号, 5) & "||" & gstrUserName
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
        '执行接口
        If Not 调用接口_执行_沈阳市() Then Exit Function
    Else
        '利用改费功能实现整张单据冲销，因为接口不允许直接退中间某笔业务，而改费无此限制
        gCominfo_沈阳市.个人编号 = str个人编号
        gCominfo_沈阳市.业务类型 = str业务类型
        gCominfo_沈阳市.业务序列号 = str业务序列号
        
        Call DebugTool("门诊改费——业务类型：" & gCominfo_沈阳市.业务类型 & "；业务序列号：" & gCominfo_沈阳市.业务序列号)
        gstrSQL = "Select 记录性质,NO From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取原单据相关信息", lng结帐ID)
        int性质 = rsTemp!记录性质
        int状态 = 2
        str单据号 = rsTemp!NO
        If Not 门诊改费_沈阳市(int性质, int状态, str单据号) Then Exit Function
    End If
    
    门诊结算冲销_沈阳市 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 入院登记_沈阳市(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim str病种编码 As String, str病种名称 As String
    Dim str入院经办时间 As String
    Dim arrPatient
    Dim rsTemp As New ADODB.Recordset
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
'   入院接口入口参数
'    1   hospital_id        医疗机构编码    20  否
'    2   indi_id            个人编号    8   否
'    3   busi_type          业务类型    2   否  "12"：住院
'    4   reg_staff          登记人员工号    5   否
'    5   reg_man            登记人姓名  10  否
'    6   reg_flag           入院方式    1   否  "0"：普通住院登记
'    7   rela_hospital_id   关联医疗机构编码    20  是
'    8   rela_serial_no     关联业务序列号  12  是
'    9   begin_date         入院时间        否  格式：YYYY-MM-DD
'    10  biz_times          本年住院次数    2   否
'    11  in_dept            入院科室编号    3   是
'    12  in_dept_name       入院科室名称    20  是
'    13  in_area            入院病区编号    3   是
'    14  in_area_name       入院病区名称    20  是
'    15  in_bed             入院病床编号    10  是
'    16  bed_type           床位类型    1   是  "0"：普通床位；"1"：急救；"2"：留观；"3"：高干
'    17  patient_id         住院号  20  否
'    18  foregift           预付款总额  10  是
'    19  foregift_remain    预付款余额  10  是  等于预付款总额
'    20  in_disease         入院诊断    20  否  疾病编码
'    21  ic_flag            用卡标志    1   是  "0"：不使用IC卡；"1"：使用IC卡
'    22  note               并发症      取入院诊断，并保存于保险帐户的并发症中
'   返回记录集
'    1."ResultSet"，本次住院的业务序列号，包含以下内容：
'    序号    字段    字段说明    最大长度    备注
'    1   serialno    业务序列号  12
    
    On Error GoTo errHand
    Call 获取病人基本信息(lng病人ID, False)
    arrPatient = Split(获取病人相关信息(lng病人ID, lng主页ID), "||")
    
    '获取病人入院日期
    gstrSQL = "Select 入院日期 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取入院日期", lng病人ID, lng主页ID)
    str入院经办时间 = Format(rsTemp!入院日期, "yyyy-MM-dd")
    
    '获取入院病种编码
    gstrSQL = "Select 编码,名称 From 保险病种 Where ID=(Select 病种ID From 保险帐户 Where 病人ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取疾病ID", lng病人ID)
    str病种编码 = rsTemp!编码: str病种名称 = rsTemp!名称
    
    gstrField_沈阳市 = "hospital_id||indi_id||busi_type||reg_staff||reg_man||reg_flag||" & _
                       "rela_hospital_id||rela_serial_no||begin_date||biz_times||in_dept||" & _
                       "in_dept_name||in_area||in_area_name||in_bed||bed_type||patient_id||" & _
                       "foregift||foregift_remain||in_disease||ic_flag||note"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.个人编号 & "||" & _
                       gCominfo_沈阳市.业务类型 & "||" & gCominfo_沈阳市.操作员工号 & "||" & _
                       gstrUserName & "||0||||||" & str入院经办时间 & "||" & _
                       arrPatient(住院次数累计) & "||" & arrPatient(入院科室编号) & "||" & _
                       arrPatient(入院科室名称) & "||" & arrPatient(入院病区编号) & "||" & _
                       arrPatient(入院病区名称) & "||" & arrPatient(入院病床编号) & "||" & _
                       arrPatient(床位类型) & "||" & arrPatient(住院号) & "||" & _
                       "||||" & str病种编码 & "||" & mintICCard & "||" & arrPatient(入院诊断)
    If Not 调用接口_准备_沈阳市(Function_沈阳市.普通住院_入院登记) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    If Not 调用接口_指定记录集_沈阳市("ResultSet") Then Exit Function
    If Not 调用接口_读取数据_沈阳市("serialno", gCominfo_沈阳市.业务序列号) Then Exit Function
    
    '更新个人帐户中的信息
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_沈阳市 & ",'顺序号','''" & gCominfo_沈阳市.业务序列号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存入院业务序列号")
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_沈阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    On Error Resume Next
    '更新并发症信息
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_沈阳市 & ",'并发症','''" & arrPatient(入院诊断) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新并发症信息")
    
    入院登记_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 入院登记撤销_沈阳市(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
    On Error GoTo errHand
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        MsgBox "该医保病人存在未结费用，不允许办理撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '读取病人基本信息
    Call 获取病人基本信息(lng病人ID, False)
    
'    1   hospital_id    医疗机构编码    20  否
'    2   serial_no      业务序列号      12  否
'    3   indi_id        个人编号        8   否
'    4   staff_no       操作员工号      5   否
'    5   staff_name     操作员姓名      10  否
    gstrField_沈阳市 = "hospital_id||serial_no||indi_id||staff_no||staff_name"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号 & "||" & _
                       gCominfo_沈阳市.个人编号 & "||" & gCominfo_沈阳市.操作员工号 & "||" & gstrUserName
    If Not 调用接口_准备_沈阳市(Function_沈阳市.普通住院_取消入院) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_沈阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    入院登记撤销_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_沈阳市(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim int出院情况 As Integer
    Dim str就诊日期 As String, str出院日期 As String
    Dim str入院病种 As String, str出院病种 As String, str并发症 As String, str出院方式 As String
    Dim bln医保出院 As Boolean, bln结帐 As Boolean
    Dim arrPatient
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '功能：将出院信息发送医保前置服务器确认
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
    
    'Modified By 朱玉宝 地区：长沙 原因：配合住院结算的修改，此处仅对已结清费用的病人办理医保出院
    bln医保出院 = False
    Call 获取病人基本信息(lng病人ID, False)
    arrPatient = Split(获取病人相关信息(lng病人ID, lng主页ID), "||")
    
    '读取病人的出院情况
    gstrSQL = "Select decode(出院方式,'转院',3,0) 出院情况,出院方式,入院日期,出院日期 From 病案主页 " & _
            " Where 病人ID = [1] And 主页ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院方式", lng病人ID, lng主页ID)
    int出院情况 = rsTemp!出院情况
    str出院方式 = rsTemp!出院方式
    str就诊日期 = Format(rsTemp!入院日期, "yyyy-MM-dd")
    If Not IsNull(rsTemp!出院日期) Then
        str出院日期 = Format(rsTemp!出院日期, "yyyy-MM-dd")
    Else
        str出院日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If

    '让用户选择入院病种及出院病种，以便更新入院信息
    If Not frm病种选择_沈阳.ShowSelect(TYPE_沈阳市, lng病人ID, lng主页ID, str入院病种, str出院病种, str并发症) Then Exit Function
    
    If 存在未结费用(lng病人ID, lng主页ID) Then  'HIS出院要更新入院信息及出院病种
        '先更新病人入院情况
    '    1   hospital_id    医疗机构编码20  否
    '    2   serial_no      业务序列号  12  否
    '    3   busi_type      业务类型    2   否  "12"：住院
    '    4   staff_no       操作员工号  5   否
    '    5   staff_name     操作员姓名  10  否
    '    6   begin_date     就诊时间        是  格式：YYYY-MM-DD
    '    7   in_dept        入院科室编号3   是
    '    8   in_dept_name   入院科室名称20  是
    '    9   in_area        入院病区编号3   是
    '    10  in_area_name   入院病区名称20  是
    '    11  in_bed         入院病床编号10  是
    '    12  bed_type       床位类型    1   是  "0"：普通床位；"1"：急救；"2"：留观；"3"：高干
    '    13  patient_id     住院号      20  是
    '    14  old_patient_id 原住院号    20  是
    '    15  in_disease     入院诊断    20  是  疾病编码
    '    16  note           备注        100 是
        '取并发症信息
        gstrField_沈阳市 = "hospital_id||serial_no||busi_type||staff_no||staff_name||begin_date||" & _
                        "in_dept||in_dept_name||in_area||in_area_name||in_bed||bed_type||patient_id||old_patient_id||in_disease||note||fin_disease"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号 & "||" & _
                        gCominfo_沈阳市.业务类型 & "||" & gCominfo_沈阳市.操作员工号 & "||" & _
                        gstrUserName & "||" & str就诊日期 & "||" & arrPatient(入院科室编号) & "||" & _
                        arrPatient(入院科室名称) & "||" & arrPatient(入院病区编号) & "||" & _
                        arrPatient(入院病区名称) & "||" & arrPatient(入院病床编号) & "||" & _
                        arrPatient(床位类型) & "||" & arrPatient(住院号) & "||" & _
                        arrPatient(住院号) & "||" & str入院病种 & "||" & str并发症 & "||" & str出院病种
        If Not 调用接口_准备_沈阳市(Function_沈阳市.住院信息_修改) Then Exit Function
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
        If Not 调用接口_执行_沈阳市 Then Exit Function
    Else
        bln医保出院 = True
        '出院接口入口参数
    '    1   hospital_id    医疗机构编码20  否
    '    2   serial_no      业务序列号  12  否
    '    3   indi_id        个人编号    8   否
    '    4   busi_type      业务类型    2   否  "12"：住院
    '    5   fin_disease    出院疾病    20      疾病编码
    '    6   end_date       出院日期            格式：YYYY-MM-DD
    '    7   fin_info       出院详情    10
    '    8   end_staff      终结人工号  5   否
    '    9   end_man        终结人姓名  10  否
    '    10  end_flag       终结处理    1   否  "0"：正常终结；"3"：住院转院
    '    11  begin_date     就诊时间        否  格式：YYYY-MM-DD
    '    12  ic_flag        用卡标志    1   否  "0"：不使用IC卡；"1"：使用IC卡
        '判断该病人是否结算过，没有结算过的病人费用为零，说明需要调用就诊登记撤销
        bln结帐 = False
        gstrSQL = "Select 1 From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] And Nvl(结帐ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否该调用就诊登记撤销", lng病人ID, lng主页ID)
        If Not rsTemp.EOF Then
            bln结帐 = True
        End If
        If Not bln结帐 Then
            gstrField_沈阳市 = "hospital_id||serial_no||indi_id||staff_no||staff_name"
            gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号 & "||" & _
                               gCominfo_沈阳市.个人编号 & "||" & gCominfo_沈阳市.操作员工号 & "||" & gstrUserName
            If Not 调用接口_准备_沈阳市(Function_沈阳市.普通住院_取消入院) Then Exit Function
            If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
            If Not 调用接口_执行_沈阳市 Then Exit Function
        Else
            gstrField_沈阳市 = "hospital_id||serial_no||indi_id||busi_type||fin_disease||end_date||" & _
                            "fin_info||end_staff||end_man||end_flag||begin_date||ic_flag"
            gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号 & "||" & _
                            gCominfo_沈阳市.个人编号 & "||" & gCominfo_沈阳市.业务类型 & "||" & _
                            str出院病种 & "||" & str出院日期 & "||" & str出院方式 & "||" & _
                            gCominfo_沈阳市.操作员工号 & "||" & gstrUserName & "||" & int出院情况 & "||" & _
                            str就诊日期 & "||" & mintICCard
            If Not 调用接口_准备_沈阳市(Function_沈阳市.出院登记) Then Exit Function
            If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
            If Not 调用接口_执行_沈阳市 Then Exit Function
        End If
    End If
    
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_沈阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    MsgBox IIf(bln医保出院, "医保出院", "HIS出院") & "办理成功！", vbInformation, gstrSysName
    出院登记_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记撤销_沈阳市(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim int出院情况 As Integer, str住院号 As String
    Dim rsTemp As New ADODB.Recordset
'    1   hospital_id    医疗机构编码    20  否
'    2   busi_type      业务类型        2   否  "12"：住院
'    3   indi_id        个人编号        8   否
'    4   end_flag       终结处理标志    1   否  "0"：正常终结；"3"：转院住院
'    5   serial_no      业务序列号      12  否
'    6   patient_id     住院号          20  否
    On Error GoTo errHand
        
    '由于出院后可能进行了其它业务，所以取最后一次住院结算的业务类型和业务序列号
    gstrSQL = " Select 支付顺序号,备注 From 保险结算记录 " & _
              " Where 性质=2 And 险类=" & TYPE_沈阳市 & " And 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取住院相关信息", lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
        gCominfo_沈阳市.业务序列号 = rsTemp!支付顺序号
        gCominfo_沈阳市.业务类型 = rsTemp!备注
    Else
        Call 获取病人基本信息(lng病人ID, False)
    End If
    
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '读取病人的出院情况
        gstrSQL = "Select decode(A.出院方式,'转院',3,0) 出院方式,B.住院号 " & _
                " From 病案主页 A,病人信息 B " & _
                " Where A.病人ID=B.病人ID And A.病人ID = [1] And A.主页ID = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院方式", lng病人ID, lng主页ID)
        int出院情况 = rsTemp!出院方式
        str住院号 = Nvl(rsTemp!住院号)
        
        '撤销出院接口入口参数
        gstrField_沈阳市 = "hospital_id||busi_type||indi_id||end_flag||serial_no||patient_id"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务类型 & "||" & _
                        gCominfo_沈阳市.个人编号 & "||" & int出院情况 & "||" & _
                        gCominfo_沈阳市.业务序列号 & "||" & str住院号
        If Not 调用接口_准备_沈阳市(Function_沈阳市.取消出院) Then Exit Function
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
        If Not 调用接口_执行_沈阳市 Then Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_沈阳市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    出院登记撤销_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 住院虚拟结算_沈阳市(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim bln单据头 As Boolean, blnUp As Boolean, blnMoveNext As Boolean, blnTrans As Boolean
    Dim curYB总额 As Currency
    Dim str支付金额 As String, str基金名称 As String, str基金编号 As String
    Dim str发生日期 As String, str医保序号 As String, arr医保序号
    Dim str医生编号 As String, str医生姓名 As String, str项目编码 As String
    Dim str规格 As String, str产地 As String, str剂型 As String
    Dim lngRecord As Long, lng原始费用ID As Long, lng负计数 As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    Dim gcn上传 As New ADODB.Connection
    Dim i As Long
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    '接口返回的报销额减去本次住院期间以往报销额的汇总金额后，才是本次的实际报销额
    'rsExse记录集中的字段清单
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    On Error GoTo errHand
    
    '取医生编号与姓名
    Call DebugTool("取医生姓名")
    str医生编号 = "": str医生姓名 = ""
    If Not IsNull(rsExse!医生) Then
        gstrSQL = "Select 编号 From 人员表 Where 姓名=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生编号", CStr(rsExse!医生))
        str医生编号 = rsTemp!编号
        str医生姓名 = rsExse!医生
    End If
    
    '新打开一个事务用来上传费用明细，避免重复上传
    Set gcn上传 = GetNewConnection
    Call DebugTool("获取病人基本信息")
    Call 获取病人基本信息(lng病人ID, False)
    gCominfo_沈阳市.总费用 = 0
    '将未上传的费用明细上传
'    1   hospital_id    医疗机构编码    20  否
'    2   indi_id        个人编号        8   否
'    3   busi_type      业务类型        2   否  "12"：住院
'    4   serial_no      业务序列号      12  否
'    5   ordinal_no     内部序数        3   是
'    6   input_staff    录入人工号      5   是
'    7   input_man      录入人姓名      10  是
'    8   recipe_no      处方号          20  是
'--------2004-01-12增加--------
'    9   doctor_no      处方医生编号    12
'    10  doctor_name    处方医生姓名    10
'------------------------------
'    数据集用来存放门诊费用明细信息，其名称为："FeeInfo"，包含以下内容:
'    序号    入参    入参说明    最大长度    是否可为空  备注
'    1   medi_item_type 项目药品类型    1   否  "0"：诊疗项目；"1"：西药；"2"：中成药；"3"：中草药
'    2   his_item_code  医院药品项目编码20  否
'    3   his_item_name  医院药品项目名称50  否
'    4   model          剂型            30  是
'    5   factory        厂家            50  是
'    6   standard       规格            30  是
'    7   fee_date       费用发生时间        否  格式：YYYY-MM-DD；改费时可以不记录
'    8   unit           计量单位        10  是
'    9   price          单价            12  否
'    10  dosage         用量            12  是
'    11  money          金额            12  否
'    12  usage_flag     用药标志    1   否  "0"：普通；"1"：出院带药；"2"：抢救
'    13  usage_days     出院带药天数    3   是
'    14  opp_serial_fee 对应费用序列号  12  是  退费时使用
    
    '住院预结算时，主要上传的都是床位费等自动计算产生的处方，全部产生到一张处方里（预结算上传的明细对方没保存）
    Call DebugTool("上传明细")
    For i = 1 To 2
        With rsExse
            '20031231:周韬:床位费的冲销单据不是新单据,但应该后上传冲销单据,不然取不到要冲销的opp_serial_fee
            If i = 1 Then
                .Filter = "金额>=0"
            ElseIf i = 2 Then
                .Filter = "金额<0"
            End If
            
            'todo :先取对应的医保费用序列号（存在负记录，无法上传，需要解决，可能是床位费产生的负单据，接口只能处理冲销的单据）
            str医保序号 = ""
            Do While Not .EOF
                If Nvl(!金额, 0) < 0 And Nvl(!是否上传, 0) = 0 Then
                    '20040105:周韬:对直接冲销退费和输入负数退费分开处理
                    If !记录状态 = 1 Then
                        '为新单金额又小于0,是输入负数退费,无明确对应的原始费用记录
                        If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                            gstrSQL = "Select Decode(trim(标识码),NULL,编码,'',编码,标识码) 编码 From 药品目录 Where 药品ID=[1]"
                        Else
                            gstrSQL = "Select Decode(trim(标识主码),NULL,编码,'',编码,标识主码) 编码 From 收费细目 Where ID=[2]"
                        End If
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取项目的编码或标识码", CLng(!收费细目ID))
                        str医保序号 = str医保序号 & IIf(str医保序号 = "", "", "|") & GetInsureSerial2(rsTemp!编码, Nvl(!金额, 0), IIf(str医保序号 = "", True, False))
                    Else
                        gstrSQL = " Select ID From 住院费用记录" & _
                                  " Where 记录性质=[1] And 记录状态=3 And NO=[2] And 序号=[3]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取原始记录的费用ID", CLng(!记录性质), CStr(!NO), CLng(!序号))
                        lng原始费用ID = rsTemp!ID
                        str医保序号 = str医保序号 & IIf(str医保序号 = "", "", "|") & GetInsureSerial(lng原始费用ID, IIf(str医保序号 = "", True, False))
                    End If
                End If
                .MoveNext
            Loop
            
            '上传费用明细
            lngRecord = 1: lng负计数 = 1
            
            If str医保序号 <> "" Then arr医保序号 = Split(str医保序号, "|")
            If .RecordCount <> 0 Then .MoveFirst
            
            bln单据头 = False: blnUp = False '20031231:周韬
            gcn上传.BeginTrans: blnTrans = True '20031231:周韬:确保HIS与对方上传数据一致
            
            Do While Not .EOF
                gCominfo_沈阳市.总费用 = gCominfo_沈阳市.总费用 + !金额
                
                '传单据头，保证只传一次（因为没传的都是自动计算产生的单据）
                If Nvl(!是否上传, 0) = 0 Then
                    blnUp = True
                    If Not bln单据头 Then
                        Call DebugTool("上传明细-单据头")
                        gstrField_沈阳市 = "hospital_id||indi_id||busi_type||serial_no||ordinal_no||input_staff||input_man||recipe_no||doctor_no||doctor_name"
                        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.个人编号 & "||" & _
                                        gCominfo_沈阳市.业务类型 & "||" & gCominfo_沈阳市.业务序列号 & "||||" & _
                                        gCominfo_沈阳市.操作员工号 & "||" & gstrUserName & "||||" & str医生编号 & "||" & str医生姓名
                        If Not 调用接口_准备_沈阳市(Function_沈阳市.住院费用_上传明细) Then
                            If blnTrans Then gcn上传.RollbackTrans
                            Exit Function
                        End If
                        If Not 调用接口_写入口参数_沈阳市(1) Then
                            If blnTrans Then gcn上传.RollbackTrans
                            Exit Function
                        End If
                        bln单据头 = True
                        blnMoveNext = False
                        
                        '指定记录集
                        If Not 调用接口_指定记录集_沈阳市("FeeInfo") Then
                            If blnTrans Then gcn上传.RollbackTrans
                            Exit Function
                        End If
                    End If
                    
                    '传单据明细
                    Call DebugTool("上传明细-单据明细")
                    gstrSQL = "Select 标识主码,编码,名称,规格 From 收费细目 Where ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取收费细目相关信息", CLng(!收费细目ID))
                    str规格 = Nvl(rsTemp!规格)
                    str产地 = ""
                    If InStr(1, str规格, "┆") <> 0 Then
                        str产地 = ToVarchar(Trim(Split(str规格, "┆")(1)), 50)
                        str规格 = ToVarchar(Trim(Split(str规格, "┆")(0)), 30)
                    Else
                        str规格 = ToVarchar(Trim(str规格), 30)
                    End If
            
                    '如果是药品，再去取剂型
                    Call DebugTool("上传明细-单据明细（取标识码）")
                    str剂型 = ""
                    str项目编码 = ""
                    If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                        gstrSQL = " Select C.编码,C.标识码,B.名称 剂型 From 药品信息 A,药品剂型 B,药品目录 C " & _
                                  " Where A.剂型=B.编码 And A.药名ID=C.药名ID And C.药品ID =[1]"
                        Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "取剂型", CLng(!收费细目ID))
                        str剂型 = Nvl(rsPhysic!剂型)
                        str项目编码 = Nvl(rsPhysic!标识码)
                        If str项目编码 = "" Then str项目编码 = Nvl(rsPhysic!编码)
                    Else
                        str项目编码 = Nvl(rsTemp!标识主码)
                        If str项目编码 = "" Then str项目编码 = Nvl(rsTemp!编码)
                    End If
                    
                    gstrSQL = "Select ID From 住院费用记录 Where NO=[1] And 记录性质=[2] And 记录状态=[3] And 序号=[4]"
                    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "读取费用ID", CStr(!NO), CLng(!记录性质), CLng(!记录状态), CLng(!序号))
                    
                    Call DebugTool("上传明细-单据明细（取登记时间）")
                    str发生日期 = Format(!发生时间, "yyyy-MM-dd")
                    gstrField_沈阳市 = "medi_item_type||his_item_code||his_item_name||model||factory||standard||" & _
                                "fee_date||unit||price||dosage||money||usage_flag||usage_days||opp_serial_fee||hos_serial"
                    If Nvl(!金额, 0) >= 0 Then
                        gstrValue_沈阳市 = IIf(!收费类别 = 5, "1", IIf(!收费类别 = 6, "2", IIf(!收费类别 = 7, "3", "0"))) & "||" & _
                                    str项目编码 & "||" & Nvl(rsTemp!名称) & "||" & str剂型 & "||" & str产地 & "||" & str规格 & "||" & _
                                    str发生日期 & "||" & Nvl(!计算单位) & "||" & !价格 & "||" & !数量 & "||" & Nvl(!金额, 0) & "||0||||||" & rsDetail!ID
                    Else
                        '20040105:周韬:退费时,将Hos_Serial传为HIS这边退费记录的ID,而不是被退费记录的ID,因为负数冲销无对应原始记录
                        gstrValue_沈阳市 = IIf(!收费类别 = 5, "1", IIf(!收费类别 = 6, "2", IIf(!收费类别 = 7, "3", "0"))) & "||" & _
                                    str项目编码 & "||" & Nvl(rsTemp!名称) & "||" & str剂型 & "||" & str产地 & "||" & str规格 & "||" & _
                                    str发生日期 & "||" & Nvl(!计算单位) & "||" & !价格 & "||" & !数量 & "||" & Nvl(!金额, 0) & "||0||||" & _
                                    arr医保序号(lng负计数 - 1) & "||" & rsDetail!ID
                    End If
                    Call DebugTool("上传明细-单据明细（完成）")
                    
                    If bln单据头 And blnMoveNext Then
                        If Not 调用接口_移动记录集_沈阳市(MoveNext) Then
                            If blnTrans Then gcn上传.RollbackTrans
                            Exit Function
                        End If
                    End If
                    If Not 调用接口_写入口参数_沈阳市(lngRecord) Then
                        If blnTrans Then gcn上传.RollbackTrans
                        Exit Function
                    End If
                    
                    lngRecord = lngRecord + 1
                    If !金额 < 0 Then lng负计数 = lng负计数 + 1
                    blnMoveNext = True
                    
                    '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
                    'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                    gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
                    gcn上传.Execute gstrSQL, , adCmdStoredProc
                End If
                .MoveNext
            Loop
            If blnUp Then
                If Not 调用接口_执行_沈阳市() Then
                    If blnTrans Then gcn上传.RollbackTrans
                    Exit Function
                End If
            End If
            gcn上传.CommitTrans: blnTrans = False
        End With
    Next
    
    
    '----------------------------------------------------------------------------------------------------------------
'    预结算入口参数
'    1   hospital_id     医疗机构编码    20  否
'    2   serial_no   业务序列号  12  否
    Call DebugTool("调预结算接口")
    gstrField_沈阳市 = "hospital_id||serial_no"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号
    If Not 调用接口_准备_沈阳市(Function_沈阳市.住院结算_预结算) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
'    返回记录集信息
'    1."CalcResultInfo"，病人的本次住院未收费用信息，包含以下内容：
'    序号    字段    字段说明    最大长度    备注
'    1   fund_id        基金编码    3
'    2   fund_name      基金名称    30
'    3   real_pay       支付金额    12  单位：元
'    2."lastbalance"，个人帐户余额信息，包含以下内容：
'    序号    字段    字段说明    最大长度    备注
'    1   Last_balance   个人帐户余额18  单位：元
    If Not 调用接口_指定记录集_沈阳市("CalcResultInfo") Then Exit Function
    curYB总额 = 0
    
    Call DebugTool("准备取预结算结果")
    If 调用接口_记录数_沈阳市 Then
        Do While True
            Call 调用接口_读取数据_沈阳市("fund_id", str基金编号)
            Call 调用接口_读取数据_沈阳市("fund_name", str基金名称)
            Call 调用接口_读取数据_沈阳市("real_pay", str支付金额)
            curYB总额 = curYB总额 + Val(str支付金额)
            If str基金编号 <> 999 And Val(str支付金额) <> 0 Then
                住院虚拟结算_沈阳市 = 住院虚拟结算_沈阳市 & IIf(住院虚拟结算_沈阳市 = "", "", "|") & _
                            IIf(str基金编号 = "003", "个人帐户", str基金名称) & ";" & str支付金额 & ";0" ' & IIf(str基金编号 = "003", "1", "0")
            End If
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    Call DebugTool("判断费用是否相等")
    If Format(gCominfo_沈阳市.总费用, "#####0.00") <> Format(curYB总额, "#####0.00") Then
        If MsgBox("HIS系统的总费用与医保返回的总费用不符，要继续结帐吗？" & vbCrLf & _
        "    HIS总费用：" & Format(gCominfo_沈阳市.总费用, "#####0.00") & Space(5) & "医保总费用：" & Format(curYB总额, "#####0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            住院虚拟结算_沈阳市 = ""
            Exit Function
        End If
    End If
    If 住院虚拟结算_沈阳市 = "" Then 住院虚拟结算_沈阳市 = "个人帐户;0;0"
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcn上传.RollbackTrans
End Function

Public Function 住院结算_沈阳(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim lng主页ID As Long, lngRecord As Long
    Dim bln单据头 As Boolean, blnUp As Boolean, blnMoveNext As Boolean
    Dim str基金名称 As String, str基金编号 As String, str支付金额 As String, str发生日期 As String, str个人编号 As String
    Dim cur帐户支付 As Currency
    Dim cur统筹基金 As Currency, cur现金 As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rsExse As New ADODB.Recordset
    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
        '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
    On Error GoTo errHand
    
    '--读IC卡
    '读出IC卡中的信息
    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_读卡) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    '取返回的记录集
    'If Not 调用接口_指定记录集_沈阳市("ICInfo") Then Exit Sub
    If Not 调用接口_读取数据_沈阳市("indi_id", str个人编号) Then Exit Function
    
    '校验个人编号是否正确
    gstrSQL = "Select 卡号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "校验该病人是否使用自己的卡进行交易", TYPE_沈阳市, lng病人ID)
    If rsTemp!卡号 <> str个人编号 Then
        Err.Raise 9000, gstrSysName, "该病人不是使用自己的医保卡，结算操作中止！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Modified By 朱玉宝 地区：长沙 原因：因为住院结算冲销是冲销本次住院所有结算记录，因此HIS未出院的医保病人，不允许进行出院结算
    If Not 医保病人已经出院(lng病人ID) Then
        MsgBox "医保病人必须出院后才能办理医保出院结算！", vbInformation, gstrSysName
        Exit Function
    End If
    Call 获取病人基本信息(lng病人ID, False)
    
    '预结算时已上传明细，此处不需要再次上传
    '读取主页ID
    gstrSQL = "Select 住院次数 主页ID From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取主页ID", lng病人ID)
    lng主页ID = rsTemp!主页ID
    
    '读取帐户支付额
    gstrSQL = "Select Nvl(A.冲预交,0) 个人帐户 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=" & TYPE_沈阳市 & _
        " And A.结算方式='个人帐户' And A.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户支付额", lng结帐ID)
    cur帐户支付 = 0
    If Not rsTemp.EOF Then
        cur帐户支付 = Nvl(rsTemp!个人帐户, 0)
    End If
    
'    正式结算
'    序号    入参    入参说明    最大长度    是否可为空  备注
'    1   hospital_id    医疗机构编码    20  否
'    2   serial_no      业务序列号      12  否
'    3   ic_flag        用卡标志        1   否  "0"：不使用IC卡；"1"：使用IC卡
'    4   debit_money    调整后的个人帐户支付金额    12  否  单位：元
'    5   indi_id        个人编号        8   否
    gstrField_沈阳市 = "hospital_id||serial_no||ic_flag||debit_money||indi_id"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号 & "||" & _
                    mintICCard & "||" & cur帐户支付 & "||" & gCominfo_沈阳市.个人编号
    If Not 调用接口_准备_沈阳市(Function_沈阳市.住院结算_正式结算) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市() Then Exit Function
    
    '指定记录集
    If Not 调用接口_指定记录集_沈阳市("payment") Then Exit Function
'    1   fund_id;    基金编码    3
'    2   fund_name   基金名称    30
'    3   real_pay    支付金额    12
'    4   serial_no   业务序列号  12
    If 调用接口_记录数_沈阳市 Then
        Do While True
            Call 调用接口_读取数据_沈阳市("fund_id", str基金编号)
            Call 调用接口_读取数据_沈阳市("fund_name", str基金名称)
            Call 调用接口_读取数据_沈阳市("real_pay", str支付金额)
            
            Select Case str基金编号
            Case "003"
                cur帐户支付 = Val(str支付金额)
            Case Is >= "900"
                cur现金 = cur现金 + Val(str支付金额)
            Case Else
                cur统筹基金 = cur统筹基金 + Val(str支付金额)
            End Select
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将结帐记录打上上传标志")
    
    '填写保险结算记录
    '帐户累计增加=公务员补助基金;帐户累计支出=公务员个人补助基金
    '累计进入统筹=个人补充基金;累计统筹报销=医疗补偿金
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_沈阳市 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng主页ID & "," & 0 & "," & 0 & "," & 0 & "," & _
        cur帐户支付 + cur现金 + cur统筹基金 & "," & cur现金 & "," & 0 & "," & cur统筹基金 & "," & cur统筹基金 & ",0," & _
        0 & "," & cur帐户支付 & ",'" & gCominfo_沈阳市.业务序列号 & "',null,null," & gCominfo_沈阳市.业务类型 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")
    
    住院结算_沈阳 = True
    '同时办理出院登记
    Call 出院登记_沈阳市(lng病人ID, lng主页ID)
    
    '20031228:周韬:加入结帐ID
    Call GetBalance(lng病人ID, lng结帐ID, gCominfo_沈阳市.业务序列号, gCominfo_沈阳市.医院编码)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_沈阳市(lng结帐ID As Long) As Boolean
    Dim lng记录ID As Long, lng病人ID As Long, lng主页ID As Long
    Dim str业务序列号 As String, str业务类型 As String
    Dim rsTemp As New ADODB.Recordset
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    On Error GoTo errHand
    Call 获取病人基本信息(lng病人ID, False)                '读取业务序列号
    
    '取冲销记录的结帐ID
    gstrSQL = "select distinct A.ID,A.病人ID from 病人结帐记录 A,病人结帐记录 B where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng记录ID = rsTemp!ID
    lng病人ID = rsTemp!病人ID
    
    '读取主页ID
    gstrSQL = "Select 住院次数 主页ID From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取主页ID", lng病人ID)
    lng主页ID = rsTemp!主页ID
    
    '取原结算记录的明细
    gstrSQL = "Select * From 保险结算记录 Where 险类=" & TYPE_沈阳市 & " And 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取结算记录", lng结帐ID)
    str业务序列号 = rsTemp!支付顺序号
    str业务类型 = rsTemp!备注
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng记录ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将结帐记录打上上传标志")
    
    'todo （注意，是回退本次住院期间所有结算单据，此过程需处理）
    '产生冲销结算记录
    With rsTemp
        gstrSQL = "zl_保险结算记录_insert(" & !性质 & "," & lng记录ID & "," & TYPE_沈阳市 & "," & !病人ID & "," & _
            !年度 & "," & -1 * Nvl(!帐户累计增加, 0) & "," & -1 * Nvl(!帐户累计支出, 0) & "," & -1 * Nvl(!累计进入统筹, 0) & "," & -1 * Nvl(!累计统筹报销, 0) & "," & lng主页ID & ",0,0,0," & _
            -1 * Nvl(!发生费用金额, 0) & "," & -1 * Nvl(!全自付金额, 0) & "," & -1 * Nvl(!首先自付金额, 0) & "," & -1 * Nvl(!进入统筹金额, 0) & "," & -1 * Nvl(!统筹报销金额, 0) & ",0," & _
            0 & "," & -1 * Nvl(!个人帐户支付, 0) & ",'" & str业务序列号 & "',null,null," & str业务类型 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "产生冲销结算记录")
    End With
    
    '准备调用住院结算冲销接口
'    序号    入参    入参说明    最大长度    是否可为空  备注
'    1   hospital_id     医疗机构编码    20  否
'    2   serial_no   业务序列号  12  否
'    3   ic_flag 用卡标志    1   是  "0"：不使用IC卡；"1"：使用IC卡
    gstrField_沈阳市 = "hospital_id||serial_no||ic_flag"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & str业务序列号 & "||" & mintICCard
    If Not 调用接口_准备_沈阳市(Function_沈阳市.住院结算_结算冲销) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    
    住院结算冲销_沈阳市 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 身份标识_沈阳市(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    身份标识_沈阳市 = frmIdentify沈阳.GetPatient(bytType, lng病人ID)
End Function

Public Function 个人余额_沈阳市(ByVal strSelfNo As String, Optional ByVal bln卡号 As Boolean = True) As Currency
    '功能: 返回个人帐户余额，如果调用接口函数失败，返回身份验证时的余额
    '参数: 是否读卡
    '返回: 返回个人帐户余额
    Dim lng病人ID As Long
    Dim str个人编号 As String, str参保人单位 As String, strData As String
    Dim cur帐户余额 As Currency
    Dim rsAccount As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select A.卡号,B.工作单位,Nvl(A.帐户余额,0) 帐户余额 From 保险帐户 A,病人信息 B " & _
              " Where A.病人ID=B.病人ID And A.险类=" & TYPE_沈阳市
    If bln卡号 Then
        gstrSQL = gstrSQL & " And A.医保号=[2]"
    Else
        gstrSQL = gstrSQL & " And A.病人ID=[1]"
    End If
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "获取卡号及单位编码", CLng(strSelfNo), strSelfNo)
    str个人编号 = rsAccount!卡号
    If InStr(1, rsAccount!工作单位, "]") <> 0 Then str参保人单位 = Mid(rsAccount!工作单位, 2, InStr(1, rsAccount!工作单位, "]") - 2)
    cur帐户余额 = rsAccount!帐户余额
    个人余额_沈阳市 = cur帐户余额
    If InStr(1, rsAccount!工作单位, "]") = 0 Then
        Call DebugTool("无单位编码，返回数据库中帐户余额")
        Exit Function
    End If
    
    '--读取个人帐户余额（由于住院没有返回，强制取一次）
    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_基金余额) Then Exit Function
    
    '写入口参数
'    1   fund_id    基金编号    3   否
'    2   indi_id    个人编号    8   否
'    3   corp_ID    单位编号    3
    'Modified By 朱玉宝 地区：长沙 原因：需要多传一个参数（corp_id）
    gstrField_沈阳市 = "fund_id||indi_id||corp_id"
    gstrValue_沈阳市 = "003||" & str个人编号 & "||" & str参保人单位
    Call 调用接口_写入口参数_沈阳市(1)
    If Not 调用接口_执行_沈阳市 Then Exit Function
    If Not 调用接口_指定记录集_沈阳市("PersonAccount") Then Exit Function
    Call 调用接口_读取数据_沈阳市("last_balance", strData)
    cur帐户余额 = Val(strData)
    
    '更新个人帐户并返回
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_沈阳市 & ",'帐户余额','" & cur帐户余额 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新帐户余额")
    个人余额_沈阳市 = cur帐户余额
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 获取病人基本信息(ByVal lng病人ID As Long, Optional bln门诊 As Boolean = True) As Boolean
    Dim str个人编号 As String, str业务序列号 As String, str就诊类型 As String
    Dim rsTemp As New ADODB.Recordset
    '返回病人的个人编号、业务序列号,就诊类型
    gstrSQL = " Select 卡号,顺序号,灰度级,当前状态,业务类型 From 保险帐户" & _
              " Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人基本信息", TYPE_沈阳市, lng病人ID)
    If rsTemp.EOF Then Exit Function
    
    gCominfo_沈阳市.个人编号 = rsTemp!卡号
    gCominfo_沈阳市.业务序列号 = Nvl(rsTemp!顺序号)
    If bln门诊 Then
        gCominfo_沈阳市.业务类型 = rsTemp!灰度级
    Else
        gCominfo_沈阳市.业务类型 = rsTemp!业务类型
    End If
    
    Call 个人余额_沈阳市(lng病人ID, False)
    获取病人基本信息 = True
End Function

Private Function 获取病人相关信息(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    Dim lng住院次数 As Long
    Dim str入院科室编号 As String, str入院科室名称 As String, str入院病区编号 As String, str入院病区名称 As String, str入院病床编号 As String, str床位类型 As String
    Dim str住院号 As String, str入院诊断 As String, str出院诊断 As String
    Dim rsTemp As New ADODB.Recordset
    '读取病人相关信息（本年住院次数、入院科室编号、入院科室名称、入院病区编号、入院病区名称、入院病床编号、床位类型、住院号、入院诊断、出院诊断）
'    床位类型
'    "0"：普通床位
'    "1"：急救
'    "2"：留观
'    "3"：高干
    
    gstrSQL = " Select nvl(住院次数累计,1) 住院次数 From 帐户年度信息 " & _
              " Where 病人ID=[1] And 年度=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人相关信息", lng病人ID, CStr(Format(zlDatabase.Currentdate, "yyyy")))
    If Not rsTemp.EOF Then lng住院次数 = rsTemp!住院次数
    If lng住院次数 = 0 Then lng住院次数 = 1
    
    '读取入院相关信息
    gstrSQL = "select C.编码 入院科室编号,C.名称 入院科室名称,B.编码 入院病区编号,B.名称 入院病区名称, " & _
             " A.入院病床 入院病床编号,F.床位类型,E.住院号 住院号 " & _
             " from 病案主页 A,部门表 B,部门表 C,病人信息 E, " & _
             " (Select D.名称 床位类型,F.床号,F.科室ID,F.病区ID  From 床位等级 D ,床位状况记录 F Where F.等级ID=D.序号) F " & _
             " Where A.入院病区ID=B.ID(+) And A.入院科室ID=C.ID(+) And A.病人ID=E.病人ID ANd A.病人ID=[1] And A.主页ID=[2]" & _
             " And A.入院病床=F.床号(+) And F.科室ID(+)=A.入院科室ID And F.病区ID(+)=A.入院病区ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院相关信息", lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
        str入院科室编号 = ToVarchar(Nvl(rsTemp!入院科室编号), 3)
        str入院科室名称 = ToVarchar(Nvl(rsTemp!入院科室名称), 20)
        str入院病区编号 = ToVarchar(Nvl(rsTemp!入院病区编号), 3)
        str入院病区名称 = ToVarchar(Nvl(rsTemp!入院病区名称), 20)
        str入院病床编号 = ToVarchar(Nvl(rsTemp!入院病床编号), 10)
        str床位类型 = Nvl(rsTemp!床位类型)
        Select Case str床位类型
        Case "急救"
            str床位类型 = 1
        Case "留观"
            str床位类型 = 2
        Case "高干"
            str床位类型 = 3
        Case Else
            str床位类型 = 0
        End Select
        str住院号 = Nvl(rsTemp!住院号)
    End If
    
    '读取入出院诊断（疾病编码）
    str入院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, True, False, True)
    str入院诊断 = Split(str入院诊断, "|")(1)
    str出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, False, False, True)
    str出院诊断 = Split(str出院诊断, "|")(1)
    获取病人相关信息 = lng住院次数 & "||" & str入院科室编号 & "||" & str入院科室名称 & "||" & _
                    str入院病区编号 & "||" & str入院病区名称 & "||" & str入院病床编号 & "||" & _
                    str床位类型 & "||" & str住院号 & "||" & str入院诊断 & "||" & str出院诊断
End Function

Public Function 处方登记_沈阳市(ByVal int性质 As Integer, ByVal int状态 As Integer, ByVal str单据号 As String) As Boolean
    Dim lng病人ID As Long, strNO As String, str发生日期 As String
    Dim blnUp As Boolean, blnMoveNext As Boolean, blnInsure As Boolean, blnFirstAID As Boolean, bln上传成功 As Boolean
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsInsure As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    Dim str医生编号 As String, str医生姓名 As String, str项目编码 As String
    Dim str规格 As String, str产地 As String, str剂型 As String
    Dim str医保序号 As String, arr医保序号 As Variant, lng负计数 As Long, lngRecords As Long
    '----------只要有一个病人的单据上传成功，则返回真允许单据保存，未上传成功的在住院虚拟结算处再上传----------
    
    '上传费用明细
    On Error GoTo errHand
    Call DebugTool("进入记帐上传")
    gstrSQL = " Select A.ID,A.病人ID,A.主页ID,A.NO,A.序号,A.记录性质,A.记录状态,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 登记时间,A.收费类别," & _
              " A.开单人 医生,B.名称 开单部门,A.收费细目ID,A.计算单位,C.项目编码 医保项目编码 ,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传" & _
              " From 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=" & TYPE_沈阳市 & ") C " & _
              " Where A.记录性质=[1] And A.记录状态=[2] And A.NO=[3]" & _
              " And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0" & _
              " Order by A.NO,A.病人ID"
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "读取费用明细", int性质, int状态, str单据号)
    
    '先检查是否存在未结码的明细
    With rsExse
        Do While Not .EOF
            If IsNull(!医保项目编码) Then
                MsgBox "第" & !序号 & "行的记录未进行医保项目对码！", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    '检查本单据是否是急救用药（对于金额大于零的做是否急救用药处理）
    '因为冲销时是程序自动选择具体的明细进行冲销，所以冲销记录不必管用药标志
    blnFirstAID = False
    If FirstAid Then
        blnFirstAID = (MsgBox("这张单据是急救用药吗？（如果是，则本单据所有明细都将做为急救用药上传至医保中心）", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    End If
    
    blnUp = False
    With rsExse
        '先取对应的医保费用序列号:20031231:周韬
        '--------------------------------------------------------------------------------------------
        str医保序号 = ""
        Do While Not .EOF
            If Nvl(!金额, 0) < 0 Then
                If lng病人ID <> !病人ID Then
                    '检查本次是否以医保身份入院
                    gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
                    Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", CLng(!病人ID), TYPE_沈阳市)
                    blnInsure = (rsInsure!Records = 1)
                    If blnInsure Then
                        blnInsure = 获取病人基本信息(!病人ID, False)
                        If blnInsure Then lng病人ID = !病人ID
                    End If
                End If
                
                If blnInsure Then
                    '200312:周韬:因为是新单中输入负数,因此没有原始费用记录,特殊处理
                    '退费时,将Hos_Serial传为HIS这边退费记录的ID,而不是被退费记录的ID
                    If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                        gstrSQL = "Select Decode(trim(标识码),NULL,编码,'',编码,标识码) 编码 From 药品目录 Where 药品ID=[1]"
                    Else
                        gstrSQL = "Select Decode(trim(标识主码),NULL,编码,'',编码,标识主码) 编码 From 收费细目 Where ID=[1]"
                    End If
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取原始记录的费用ID", CLng(!收费细目ID))
                    str医保序号 = str医保序号 & IIf(str医保序号 = "", "", "|") & GetInsureSerial2(rsTemp!编码, Nvl(!金额, 0), IIf(str医保序号 = "", True, False))
                End If
            End If
            .MoveNext
        Loop
        
        '再上传数据
        '------------------------------------------------------------------------------------------
        '20031231:周韬
        lng负计数 = 1: lng病人ID = 0
        If str医保序号 <> "" Then
            arr医保序号 = Split(str医保序号, "|")
        End If
        If .RecordCount <> 0 Then .MoveFirst
        
        Do While Not .EOF
            '传单据头
            If lng病人ID <> !病人ID Or strNO <> !NO Then
                Call DebugTool("传单据头")
                If lng病人ID <> 0 And blnInsure Then
                    '上传
                    blnUp = False
                    blnMoveNext = False
                    If Not 调用接口_执行_沈阳市() Then
                        处方登记_沈阳市 = bln上传成功
                        Exit Function
                    End If
                    bln上传成功 = True
                End If
                If lng病人ID <> !病人ID Then
                    '检查本次是否以医保身份入院
                    gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
                    Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", CLng(!病人ID), TYPE_沈阳市)
                    blnInsure = (rsInsure!Records = 1)
                    If blnInsure Then
                        blnInsure = 获取病人基本信息(!病人ID, False)
                        If blnInsure Then lng病人ID = !病人ID
                    End If
                End If
                If blnInsure Then
                    strNO = !NO
    
                    '取医生编号与姓名
                    str医生编号 = "": str医生姓名 = Nvl(rsExse!医生)
                    If str医生姓名 <> "" Then
                        gstrSQL = "Select 编号,姓名 From 人员表 Where 姓名=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生编号", str医生姓名)
                        str医生编号 = rsTemp!编号
                    End If
                    
                    '写入口参数
                    gstrField_沈阳市 = "hospital_id||indi_id||busi_type||serial_no||ordinal_no||input_staff||input_man||recipe_no||doctor_no||doctor_name"
                    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.个人编号 & "||" & _
                                    gCominfo_沈阳市.业务类型 & "||" & gCominfo_沈阳市.业务序列号 & "||||" & _
                                    gCominfo_沈阳市.操作员工号 & "||" & gstrUserName & "||" & strNO & "||" & str医生编号 & "||" & str医生姓名
                    If Not 调用接口_准备_沈阳市(Function_沈阳市.住院费用_上传明细) Then
                        处方登记_沈阳市 = bln上传成功
                        Exit Function
                    End If
                    If Not 调用接口_写入口参数_沈阳市(1) Then
                        处方登记_沈阳市 = bln上传成功
                        Exit Function
                    End If
                    blnUp = True
                    
                    '指定记录集
                    If Not 调用接口_指定记录集_沈阳市("FeeInfo") Then
                        处方登记_沈阳市 = bln上传成功
                        Exit Function
                    End If
                    lngRecords = 1
                End If
            End If
            
            '传单据明细
            If blnInsure Then
                Call DebugTool("传单据明细")
                gstrSQL = "Select 标识主码,编码,名称,规格 From 收费细目 Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取收费细目相关信息", CLng(!收费细目ID))
                str产地 = ""
                str规格 = Nvl(rsTemp!规格)
                If InStr(1, str规格, "┆") <> 0 Then
                    str产地 = ToVarchar(Trim(Split(str规格, "┆")(1)), 50)
                    str规格 = ToVarchar(Trim(Split(str规格, "┆")(0)), 30)
                Else
                    str规格 = ToVarchar(Trim(str规格), 30)
                End If
                
                str剂型 = ""
                str项目编码 = ""
                str发生日期 = Format(!登记时间, "yyyy-MM-dd")
                
                Call DebugTool("取标识码")
                If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                    gstrSQL = " Select C.编码,C.标识码,B.名称 剂型 From 药品信息 A,药品剂型 B,药品目录 C " & _
                              " Where A.剂型=B.编码 And A.药名ID=C.药名ID And C.药品ID = [1]"
                    Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "取剂型", CLng(!收费细目ID))
                    Call DebugTool("剂型:" & Nvl(rsPhysic!剂型) & "|标识码:" & Nvl(rsPhysic!标识码) & "|编码:" & Nvl(rsPhysic!编码))
                    str剂型 = Nvl(rsPhysic!剂型)
                    str项目编码 = Nvl(rsPhysic!标识码)
                    If str项目编码 = "" Then str项目编码 = Nvl(rsPhysic!编码)
                Else
                    str项目编码 = Nvl(rsTemp!标识主码)
                    If str项目编码 = "" Then str项目编码 = Nvl(rsTemp!编码)
                End If
                
                Call DebugTool("标识码获取成功")
                gstrField_沈阳市 = "medi_item_type||his_item_code||his_item_name||model||factory||standard||" & _
                            "fee_date||unit||price||dosage||money||usage_flag||usage_days||opp_serial_fee||hos_serial"
                
                
                If Nvl(!金额, 0) < 0 Then
                    '20031231:周韬:负数费用要上传opp_serial_fee
                    Call DebugTool("上传负费用")
                    gstrValue_沈阳市 = IIf(!收费类别 = 5, "1", IIf(!收费类别 = 6, "2", IIf(!收费类别 = 7, "3", "0"))) & "||" & _
                                str项目编码 & "||" & Nvl(rsTemp!名称) & "||" & str剂型 & "||" & str产地 & "||" & str规格 & "||" & _
                                str发生日期 & "||" & Nvl(!计算单位) & "||" & Format(!金额 / !数量, "#####0.0000;-#####0.0000;0;") & "||" & !数量 & "||" & !金额 & "||0||||" & _
                                arr医保序号(lng负计数 - 1) & "||" & !ID
                Else
                    Call DebugTool("上传正费用")
                    gstrValue_沈阳市 = IIf(!收费类别 = 5, "1", IIf(!收费类别 = 6, "2", IIf(!收费类别 = 7, "3", "0"))) & "||" & _
                                str项目编码 & "||" & Nvl(rsTemp!名称) & "||" & str剂型 & "||" & str产地 & "||" & str规格 & "||" & _
                                str发生日期 & "||" & Nvl(!计算单位) & "||" & Format(!金额 / !数量, "#####0.0000;-#####0.0000;0;") & "||" & !数量 & "||" & !金额 & "||" & IIf(blnFirstAID, "2", "0") & "||||||" & !ID
                End If
                
                If blnMoveNext Then
                    Call DebugTool("移动记录集")
                    Call 调用接口_移动记录集_沈阳市(MoveNext)
                End If
                Call DebugTool("写入口参数")
                If Not 调用接口_写入口参数_沈阳市(lngRecords) Then
                    处方登记_沈阳市 = bln上传成功
                    Exit Function
                End If
                lngRecords = lngRecords + 1
                '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
                'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                Call DebugTool("打上传标志")
                gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
                
                blnMoveNext = True
                lng负计数 = lng负计数 + 1
            End If
            .MoveNext
        Loop
        If blnUp And blnInsure Then
            If Not 调用接口_执行_沈阳市() Then
                处方登记_沈阳市 = bln上传成功
                Exit Function
            End If
            bln上传成功 = True
        End If
    End With
    
    Call DebugTool("上传成功")
    处方登记_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    处方登记_沈阳市 = bln上传成功
End Function

Public Function 处方作废_沈阳市(ByVal int性质 As Integer, ByVal int状态 As Integer, ByVal str单据号 As String) As Boolean
    '需要找到原始那笔记录的费用ID，才能够正常作废
    Dim lng病人ID As Long, strNO As String, str发生日期 As String, int原始记录状态 As Integer
    Dim str编码 As String, str名称 As String, lng原始费用ID As Long, int序号 As Integer, lng负计数 As Long
    Dim blnUp As Boolean, blnMoveNext As Boolean, blnInsure As Boolean, bln上传成功 As Boolean
    Dim str医保序号 As String, arr医保序号
    Dim str医生编号 As String, str医生姓名 As String, str项目编码 As String
    Dim str规格 As String, str产地 As String, str剂型 As String
    
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsInsure As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    Dim cnUpData As New ADODB.Connection
    '上传费用明细
    On Error GoTo errHand
    
    int原始记录状态 = 3
    gstrSQL = " Select A.ID,A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 登记时间,A.收费类别," & _
              " A.开单人 医生,B.名称 开单部门,A.收费细目ID,A.计算单位,C.项目编码 医保项目编码 ,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传" & _
              " From 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=[4]) C " & _
              " Where A.记录性质=[1] And A.记录状态=[2] And A.NO=[3]" & _
              " And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0" & _
              " Order by A.NO,A.病人ID"
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "读取费用明细", int性质, int状态, str单据号, TYPE_沈阳市)
    
    With rsExse
        '先取对应的医保费用序列号
        str医保序号 = ""
        Do While Not .EOF
            If lng病人ID <> !病人ID Then
                '检查本次是否以医保身份入院
                gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
                Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", CLng(!病人ID), TYPE_沈阳市)
                 blnInsure = (rsInsure!Records = 1)
                If blnInsure Then
                    blnInsure = 获取病人基本信息(!病人ID, False)
                    If blnInsure Then lng病人ID = !病人ID
                End If
            End If
            
            If blnInsure Then
                gstrSQL = " Select ID From 住院费用记录" & _
                          " Where 记录性质=[1] And 记录状态=[2] And NO=[3] And 序号=[4]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取原始记录的费用ID", int性质, int原始记录状态, str单据号, CStr(!序号))
                lng原始费用ID = rsTemp!ID
                
                '20040105:周韬:
                '退费时,将Hos_Serial传为HIS这边退费记录的ID,而不是被退费记录的ID,因为负数冲销时无对应原始记录
                str医保序号 = str医保序号 & IIf(str医保序号 = "", "", "|") & GetInsureSerial(lng原始费用ID, IIf(str医保序号 = "", True, False))
            End If
            .MoveNext
        Loop
        
        '再上传数据
        blnUp = False
        lng病人ID = 0: strNO = ""
        arr医保序号 = Split(str医保序号, "|")
        lng负计数 = 1
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '传单据头
            If lng病人ID <> !病人ID Or strNO <> !NO Then
                If lng病人ID <> 0 And blnInsure Then
                    '上传
                    blnUp = False
                    blnMoveNext = False
                    If Not 调用接口_执行_沈阳市() Then
                        处方作废_沈阳市 = bln上传成功
                        Exit Function
                    End If
                    bln上传成功 = True
                End If
                strNO = !NO
                If lng病人ID <> !病人ID Then
                    '检查本次是否以医保身份入院
                    gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
                    Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", CLng(!病人ID), TYPE_沈阳市)
                    blnInsure = (rsInsure!Records = 1)
                    If blnInsure Then
                        blnInsure = 获取病人基本信息(!病人ID, False)
                        If blnInsure Then lng病人ID = !病人ID
                    End If
                End If
                
                If blnInsure Then
    
                    '取医生编号与姓名
                    str医生编号 = "": str医生姓名 = Nvl(rsExse!医生)
                    If str医生姓名 <> "" Then
                        gstrSQL = "Select 编号,姓名 From 人员表 Where 姓名=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生编号", str医生姓名)
                        str医生编号 = rsTemp!编号
                    End If
                    
                    '写入口参数
                    gstrField_沈阳市 = "hospital_id||indi_id||busi_type||serial_no||ordinal_no||input_staff||input_man||recipe_no||doctor_no||doctor_name"
                    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.个人编号 & "||" & _
                                    gCominfo_沈阳市.业务类型 & "||" & gCominfo_沈阳市.业务序列号 & "||||" & _
                                    gCominfo_沈阳市.操作员工号 & "||" & gstrUserName & "||" & strNO & "||" & str医生编号 & "||" & str医生姓名
                    If Not 调用接口_准备_沈阳市(Function_沈阳市.住院费用_上传明细) Then
                        处方作废_沈阳市 = bln上传成功
                        Exit Function
                    End If
                    If Not 调用接口_写入口参数_沈阳市(1) Then
                        处方作废_沈阳市 = bln上传成功
                        Exit Function
                    End If
                    blnUp = True
                    
                    '指定记录集
                    If Not 调用接口_指定记录集_沈阳市("FeeInfo") Then
                        处方作废_沈阳市 = bln上传成功
                        Exit Function
                    End If
                End If
            End If
            
            If blnInsure Then
                '读取相关信息
                int序号 = !序号
                
                gstrSQL = "Select 标识主码,编码,名称,规格 From 收费细目 Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取收费细目相关信息", CLng(!收费细目ID))
                str编码 = rsTemp!编码
                str名称 = rsTemp!名称
                str规格 = Nvl(rsTemp!规格)
                str产地 = ""
                If InStr(1, str规格, "┆") <> 0 Then
                    str产地 = ToVarchar(Trim(Split(str规格, "┆")(1)), 50)
                    str规格 = ToVarchar(Trim(Split(str规格, "┆")(0)), 30)
                Else
                    str规格 = ToVarchar(Trim(str规格), 30)
                End If
                
                '传单据明细
                str剂型 = ""
                str项目编码 = ""
                str发生日期 = "" '改废时发生日期为空
                
                If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                    gstrSQL = " Select C.编码,C.标识码,B.名称 剂型 From 药品信息 A,药品剂型 B,药品目录 C " & _
                              " Where A.剂型=B.编码 And A.药名ID=C.药名ID And C.药品ID =[1]"
                    Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "取剂型", CLng(!收费细目ID))
                    str剂型 = Nvl(rsPhysic!剂型)
                    str项目编码 = Nvl(rsPhysic!标识码)
                    If str项目编码 = "" Then str项目编码 = Nvl(rsPhysic!编码)
                Else
                    str项目编码 = Nvl(rsTemp!标识主码)
                    If str项目编码 = "" Then str项目编码 = Nvl(rsTemp!编码)
                End If
                
                gstrField_沈阳市 = "medi_item_type||his_item_code||his_item_name||model||factory||standard||" & _
                            "fee_date||unit||price||dosage||money||usage_flag||usage_days||opp_serial_fee||hos_serial"
                gstrValue_沈阳市 = IIf(!收费类别 = 5, "1", IIf(!收费类别 = 6, "2", IIf(!收费类别 = 7, "3", "0"))) & "||" & _
                            str项目编码 & "||" & str名称 & "||" & str剂型 & "||" & str产地 & "||" & str规格 & "||" & _
                            str发生日期 & "||" & Nvl(!计算单位) & "||" & Format(!金额 / !数量, "#####0.0000;-#####0.0000;0;") & "||" & !数量 & "||" & !金额 & "||0||||" & _
                            arr医保序号(lng负计数 - 1) & "||" & !ID
                
                If blnMoveNext Then Call 调用接口_移动记录集_沈阳市(MoveNext)
                If Not 调用接口_写入口参数_沈阳市(.AbsolutePosition) Then
                    处方作废_沈阳市 = bln上传成功
                    Exit Function
                End If
                'Call ErrInformation
                '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
                'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
                
                lng负计数 = lng负计数 + 1
                blnMoveNext = True
            End If
            .MoveNext
        Loop
        If blnUp And blnInsure Then
            If Not 调用接口_执行_沈阳市() Then
                处方作废_沈阳市 = bln上传成功
                Exit Function
            End If
            bln上传成功 = True
        End If
    End With
    
    处方作废_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    处方作废_沈阳市 = bln上传成功
End Function

Public Function 门诊改费_沈阳市(ByVal int性质 As Integer, ByVal int状态 As Integer, ByVal str单据号 As String) As Boolean
    '需要找到原始那笔记录的费用ID，才能够正常作废
    Dim lng病人ID As Long, strNO As String, str发生日期 As String, int原始记录状态 As Integer
    Dim str编码 As String, str名称 As String, lng原始费用ID As Long, int序号 As Integer, lng负计数 As Long
    Dim blnUp As Boolean, blnMoveNext As Boolean, blnInsure As Boolean
    Dim str医保序号 As String, arr医保序号
    Dim str医生编号 As String, str医生姓名 As String, str项目编码 As String
    Dim str规格 As String, str产地 As String, str剂型 As String
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsInsure As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    '上传费用明细
    On Error GoTo errHand
    
    int原始记录状态 = 3
    gstrSQL = " Select A.ID,A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 登记时间,A.收费类别," & _
              " A.开单人 医生,B.名称 开单部门,A.收费细目ID,A.计算单位,C.项目编码 医保项目编码 ,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传" & _
              " From 门诊费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=[4]) C " & _
              " Where A.记录性质=[1] And A.记录状态=[2] And A.NO=[3]" & _
              " And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0" & _
              " Order by A.NO,A.病人ID"
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "读取费用明细", int性质, int状态, str单据号, TYPE_沈阳市)
    Call DebugTool("取得冲销记录数：" & rsExse.RecordCount)
    
    With rsExse
        '先取对应的医保费用序列号
        str医保序号 = ""
        Do While Not .EOF
            gstrSQL = " Select ID From 门诊费用记录" & _
                      " Where 记录性质=[1] And 记录状态=[2] And NO=[3] And 序号=[4]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取原始记录的费用ID", int性质, int原始记录状态, str单据号, CLng(!序号))
            lng原始费用ID = rsTemp!ID
            
            '20040105:周韬:
            '退费时,将Hos_Serial传为HIS这边退费记录的ID,而不是被退费记录的ID,因为负数冲销时无对应原始记录
            str医保序号 = str医保序号 & IIf(str医保序号 = "", "", "|") & GetInsureSerial_OutExse(lng原始费用ID, IIf(str医保序号 = "", True, False))
            .MoveNext
        Loop
        Call DebugTool("退费序列号：" & str医保序号)
        
        '再上传数据
        blnUp = False
        lng病人ID = 0: strNO = ""
        arr医保序号 = Split(str医保序号, "|")
        lng负计数 = 1
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '传单据头
            If lng病人ID <> !病人ID Then
                lng病人ID = !病人ID
                '取医生编号与姓名
                str医生编号 = "": str医生姓名 = Nvl(rsExse!医生)
                If str医生姓名 <> "" Then
                    gstrSQL = "Select 编号,姓名 From 人员表 Where 姓名=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生编号", str医生姓名)
                    str医生编号 = rsTemp!编号
                End If
                
                '写入口参数
'                    1   hospital_id    医疗机构编码    20  否
'                    2   indi_id        个人编号    8   否
'                    3   busi_type      业务类型    2   否  "11"：门诊
'                    4   serial_no      业务序列号  12  否
'                    5   ic_flag        用卡标志    1   是  "0"：不使用IC卡                    "1"：使用IC卡
'                    6   Reg_staff      登记人员工号    5   否
'                    7   Reg_man        登记人姓名  10  否
'                    8   begin_date     就诊时间        是  格式：YYYY-MM-DD HH:MI:SS(24小时)
'                    9   calcSaveFlag   计算保存标志    1   否  "0"：试算                    "1"：收费
'                    10  accMoney       个人帐户支付金额    18  否
'                    11  recipe_no      处方号  20  是
                gstrField_沈阳市 = "hospital_id||indi_id||busi_type||serial_no||ic_flag||Reg_staff||Reg_man||begin_date||calcSaveFlag||accMoney||recipe_no"
                gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.个人编号 & "||" & _
                                gCominfo_沈阳市.业务类型 & "||" & gCominfo_沈阳市.业务序列号 & "||1||" & _
                                gCominfo_沈阳市.操作员工号 & "||" & gstrUserName & "||" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "||" & _
                                "1||0||" & str单据号
                If Not 调用接口_准备_沈阳市(IIf(gCominfo_沈阳市.业务类型 = 业务分类_沈阳市.门诊规定病, 门诊规定病_收费, 普通门诊_收费)) Then Exit Function
                If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
                Call DebugTool("启动接口")
                
                '指定记录集
                If Not 调用接口_指定记录集_沈阳市("FeeInfo") Then Exit Function
            End If
            
            '读取相关信息
            int序号 = !序号
            
            gstrSQL = "Select 标识主码,编码,名称,规格 From 收费细目 Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取收费细目相关信息", CLng(!收费细目ID))
            str编码 = rsTemp!编码
            str名称 = rsTemp!名称
            str规格 = Nvl(rsTemp!规格)
            str产地 = ""
            If InStr(1, str规格, "┆") <> 0 Then
                str产地 = ToVarchar(Trim(Split(str规格, "┆")(1)), 50)
                str规格 = ToVarchar(Trim(Split(str规格, "┆")(0)), 30)
            Else
                str规格 = ToVarchar(Trim(str规格), 30)
            End If
            
            '传单据明细
            str剂型 = ""
            str项目编码 = ""
            str发生日期 = "" '改废时发生日期为空
            
            If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                gstrSQL = " Select C.编码,C.标识码,B.名称 剂型 From 药品信息 A,药品剂型 B,药品目录 C " & _
                          " Where A.剂型=B.编码 And A.药名ID=C.药名ID And C.药品ID =[1]"
                Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "取剂型", CLng(!收费细目ID))
                str剂型 = Nvl(rsPhysic!剂型)
                str项目编码 = Nvl(rsPhysic!标识码)
                If str项目编码 = "" Then str项目编码 = Nvl(rsPhysic!编码)
            Else
                str项目编码 = Nvl(rsTemp!标识主码)
                If str项目编码 = "" Then str项目编码 = Nvl(rsTemp!编码)
            End If
            
            Call DebugTool("准备传递明细")
            gstrField_沈阳市 = "medi_item_type||his_item_code||his_item_name||model||factory||" & _
                        "standard||fee_date||unit||price||dosage||money||opp_serial_fee||hos_serial"
            gstrValue_沈阳市 = IIf(!收费类别 = 5, "1", IIf(!收费类别 = 6, "2", IIf(!收费类别 = 7, "3", "0"))) & "||" & _
                        str项目编码 & "||" & str名称 & "||" & str剂型 & "||" & str产地 & "||" & str规格 & "||" & _
                        str发生日期 & "||" & Nvl(!计算单位) & "||" & Format(!金额 / !数量, "#####0.0000;-#####0.0000;0;") & "||" & !数量 & "||" & !金额 & "||" & arr医保序号(lng负计数 - 1) & "||" & !ID
            
            If blnMoveNext Then Call 调用接口_移动记录集_沈阳市(MoveNext)
            If Not 调用接口_写入口参数_沈阳市(.AbsolutePosition) Then Exit Function
            
            Call DebugTool("准备打上传标志")
            'Call ErrInformation
            '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
            'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
            gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            
            lng负计数 = lng负计数 + 1
            blnMoveNext = True
            .MoveNext
        Loop
        If Not 调用接口_执行_沈阳市() Then Exit Function
    End With
    
    Call DebugTool("以改费的方式完成冲销原始门诊业务的功能")
    门诊改费_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function GetInsureSerial(ByVal lng费用序列号 As Long, Optional ByVal bln执行 As Boolean = False) As Long
    Dim str序列号 As String
    '返回医保费用序列号，是医保自己的费用序列号
'    1   hospital_id    医疗机构编码   20  否
'    2   serial_no      业务序列号     12  否
'    3   calc_flag      计算标志       1
    If bln执行 Then
        gstrField_沈阳市 = "hospital_id||serial_no||calc_flag"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号 & "||1"
        If Not 调用接口_准备_沈阳市(Function_沈阳市.住院费用_查明细) Then Exit Function
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
        If Not 调用接口_执行_沈阳市 Then Exit Function
        If Not 调用接口_指定记录集_沈阳市("calc_fee_info") Then Exit Function
    End If
    
    If 调用接口_记录数_沈阳市 Then
        Call 调用接口_移动记录集_沈阳市(MoveFirst)
        Do While True
            Call 调用接口_读取数据_沈阳市("hos_serial", str序列号)
            If Val(str序列号) = lng费用序列号 Then
                Call 调用接口_读取数据_沈阳市("serial_fee", str序列号)
                GetInsureSerial = Val(str序列号)
                Exit Function
            End If
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Function
        Loop
    End If
End Function

Private Function GetInsureSerial_OutExse(ByVal lng费用序列号 As Long, Optional ByVal bln执行 As Boolean = False) As Long
    Dim str序列号 As String
    '返回医保费用序列号，是医保自己的费用序列号
'    1   hospital_id    医疗机构编码   20  否
'    2   serial_no      业务序列号     12  否
'    3   calc_flag      计算标志       1
    If bln执行 Then
        gstrField_沈阳市 = "hospital_id||serial_no"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号
        If Not 调用接口_准备_沈阳市(Function_沈阳市.门诊_查明细) Then Exit Function
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
        If Not 调用接口_执行_沈阳市 Then Exit Function
        If Not 调用接口_指定记录集_沈阳市("FeeInfo") Then Exit Function
    End If
    
    If 调用接口_记录数_沈阳市 Then
        Call 调用接口_移动记录集_沈阳市(MoveFirst)
        Do While True
            Call 调用接口_读取数据_沈阳市("hos_serial", str序列号)
            Call DebugTool("医保中保存的HIS费用ID：" & Val(str序列号) & "；被退的原始记录的费用ID：" & lng费用序列号)
            If Val(str序列号) = lng费用序列号 Then
                Call 调用接口_读取数据_沈阳市("serial_fee", str序列号)
                Call DebugTool("找到了：" & Val(str序列号))
                GetInsureSerial_OutExse = Val(str序列号)
                Exit Function
            End If
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then
                Call DebugTool("没有找到")
                Exit Function
            End If
        Loop
    End If
End Function

Private Function GetInsureSerial2(ByVal str编码 As String, cur金额 As Currency, Optional ByVal bln执行 As Boolean = False) As Long
'功能：返回医保退费用序列号，适用于直接输入负数冲销的情况
'参数：str编码=HIS端收费细目编码
'      cur金额=当前退费金额
    Dim str序列号 As String, str对应序列号 As String, curMoney As Currency
    Dim arr退费集() As Variant
    Dim rs原始集 As New ADODB.Recordset
    Dim strFields As String, strValues As String
    Dim strTemp As String, i As Long, j As Long
    
'    1   hospital_id    医疗机构编码   20  否
'    2   serial_no      业务序列号     12  否
'    3   calc_flag      计算标志       1
    If bln执行 Then
        gstrField_沈阳市 = "hospital_id||serial_no||calc_flag"
        gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号 & "||1"
        If Not 调用接口_准备_沈阳市(Function_沈阳市.住院费用_查明细) Then Exit Function
        If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
        If Not 调用接口_执行_沈阳市 Then Exit Function
        If Not 调用接口_指定记录集_沈阳市("calc_fee_info") Then Exit Function
    End If
    
    If 调用接口_记录数_沈阳市 Then
        '初始化原始记录集
        strFields = "序列号" & "," & adDouble & "," & "20" & "|" & _
                    "金额" & "," & adDouble & "," & "20"
        Call Record_Init(rs原始集, strFields)
        
        arr退费集 = Array()
        Call 调用接口_移动记录集_沈阳市(MoveFirst)
        Do While True
            '20040105:周韬:采用新方法,必须严格对应要退费的原始记录
            '找同一项目(编码相同),金额足够退的原始记录
            Call 调用接口_读取数据_沈阳市("his_item_code", strTemp)
            If strTemp = str编码 Then
                Call 调用接口_读取数据_沈阳市("serial_fee", strTemp)
                str序列号 = strTemp
                Call 调用接口_读取数据_沈阳市("opp_serial_fee", strTemp)
                str对应序列号 = strTemp
                Call 调用接口_读取数据_沈阳市("money", strTemp)
                curMoney = Val(strTemp)
                                
                '金额足够退的且不是退费记录才可能为被退的
                If curMoney >= Abs(cur金额) And Val(str对应序列号) = 0 Then
                    strFields = "序列号|金额"
                    strValues = Val(str序列号) & "|" & curMoney
                    Call Record_Add(rs原始集, strFields, strValues)
                End If
                
                '对应序列号不为空的才是退费的记录
                If Val(str对应序列号) <> 0 Then
                    ReDim Preserve arr退费集(UBound(arr退费集) + 1)
                    arr退费集(UBound(arr退费集)) = str对应序列号
                End If
                
            End If
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
        
        '只查找未退过费的记录进行退费（查找一笔金额最接近的记录作为退费原始记录）
        With rs原始集
            If .RecordCount <> 0 Then .Sort = "金额 asc"
            Do While Not .EOF
                For j = 0 To UBound(arr退费集)
                    If Val(arr退费集(j)) = !序列号 Then
                        Exit For
                    End If
                Next
                If j > UBound(arr退费集) Then
                    '没有任何一条退费记录的对应序列号和原始记录中当前记录的序列号相同
                    GetInsureSerial2 = !序列号
                    Exit Function
                End If
                .MoveNext
            Loop
        End With
    End If
End Function

Public Function GetBalance(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, ByVal str业务序列号 As String, ByVal str医院编码 As String) As Boolean
'功能：根据业务序列号获取结算结果，存于临时表中，发票打印
'    1   hospital_id 医疗机构编码               20  否
'    2   serial_no   业务序列号                 12  否
'    3   fee_flag    是否取该次业务的费用明细   1   否  0：不取明细；1：取明细
    
    Dim int基金 As Integer, cur金额 As Currency, blnDelete As Boolean, strPara As String
    Dim str编码 As String, str名称 As String, int标记 As Integer '20031228:周韬
    Dim cur自费费用 As Currency, cur自负费用 As Currency
        
    int基金 = 1
    blnDelete = True
    gstrField_沈阳市 = "hospital_id||serial_no||fee_flag"
    gstrValue_沈阳市 = str医院编码 & "||" & str业务序列号 & "||0"
    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_获取发票信息) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    'InvoiceInfo,本次业务的基金支付组成
    If Not 调用接口_指定记录集_沈阳市("InvoiceInfo") Then Exit Function
    If 调用接口_记录数_沈阳市 Then
        Call 调用接口_移动记录集_沈阳市(MoveFirst)
        Do While True
            '20031228:周韬
            '基金编码
            If Not 调用接口_读取数据_沈阳市("Fund_id", strPara) Then Exit Function
            str编码 = Trim(strPara)
            
            '基金名称
            If Not 调用接口_读取数据_沈阳市("Fund_name", strPara) Then Exit Function
            str名称 = Trim(strPara)
            
            '基金金额
            If Not 调用接口_读取数据_沈阳市("real_pay", strPara) Then Exit Function
            cur金额 = Val(strPara)
            
            '合计标记:0-不合计的,1-合计的
            If Not 调用接口_读取数据_沈阳市("sum_flag", strPara) Then Exit Function
            int标记 = Val(strPara)
            
            '插入
            gstrSQL = "ZL_发票信息_INSERT(" & lng病人ID & "," & lng结帐ID & "," & int基金 & "," & _
                "'" & str编码 & "','" & str名称 & "'," & cur金额 & "," & int标记 & "," & _
                IIf(blnDelete, "1", "0") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "插入发票信息")
            blnDelete = False
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do '20031228:周韬,改为Exit Do
        Loop
    End If
    
    '20031228:周韬:经咨询暂时无用,但程序保留
    'FeeSquareInfo,费用结算汇总信息
    int基金 = 0
    If Not 调用接口_指定记录集_沈阳市("FeeSquareInfo") Then Exit Function
    If 调用接口_记录数_沈阳市 Then
        Call 调用接口_移动记录集_沈阳市(MoveFirst)
        Do While True
            Call 调用接口_读取数据_沈阳市("stat_name", str名称)
            Call 调用接口_读取数据_沈阳市("zfy", strPara)
            cur金额 = Val(strPara)
            Call 调用接口_读取数据_沈阳市("qzf", strPara)
            cur自费费用 = Val(strPara)
            Call 调用接口_读取数据_沈阳市("blzf", strPara)
            cur自负费用 = Val(strPara)
            gstrSQL = "ZL_发票信息_INSERT(" & lng病人ID & "," & lng结帐ID & "," & int基金 & "," & _
                "'" & cur自费费用 & "|" & cur自负费用 & "','" & str名称 & "'," & cur金额 & ",NULL," & IIf(blnDelete, "1", "0") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "插入发票信息")

            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do '20031228:周韬,改为Exit Do
        Loop
    End If
    
    GetBalance = True
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function 提取结算单_沈阳市(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, int业务类型 As Integer, str业务序列号 As String) As Boolean
    Dim intType As Integer
    Dim strData As String
    Dim strTemp As String
    Dim bln门诊 As Boolean, blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    '提取病人的结算单（分门诊、门诊规定病、住院三种类型）
    On Error GoTo errHand
    
    If Not 医保初始化_沈阳市 Then Exit Function
    
    gstrSQL = " Select 性质,支付顺序号 业务序列号,备注 业务类型 From 保险结算记录 " & _
              " Where 险类=[1] And 病人ID=[2] And 记录ID=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取该病人本次业务的业务类型及业务序列号", TYPE_沈阳市, lng病人ID, lng结帐ID)
    If rsTemp.EOF Then
        MsgBox "提取该病人本次业务数据时发生错误！（未找到交易数据）", vbInformation, gstrSysName
        Exit Function
    End If
    str业务序列号 = Trim(rsTemp!业务序列号)
    int业务类型 = Val(rsTemp!业务类型)
    bln门诊 = (rsTemp!性质 = 1)
    
    If str业务序列号 = "" Then
        MsgBox "业务序列号为空，无法提取该病人本次交易的详细数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '1   hospital_id    医疗机构编码    20  否
    '2   serial_no      业务序列号      12  否
    gstrField_沈阳市 = "hospital_id||serial_no"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & str业务序列号
    
    If bln门诊 Then
        If int业务类型 = 业务分类_沈阳市.门诊规定病 Then
            intType = 2
            If Not 调用接口_准备_沈阳市(Function_沈阳市.结算单_门诊规定病) Then Exit Function
        Else
            intType = 1
            If Not 调用接口_准备_沈阳市(Function_沈阳市.结算单_门诊) Then Exit Function
        End If
    Else
        intType = 3
        If Not 调用接口_准备_沈阳市(Function_沈阳市.结算单_住院) Then Exit Function
    End If
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    
    Call DebugTool("调用相关业务请求")
    gcnOracle.BeginTrans
    '删除今天以前的结算单
    gstrSQL = "Delete From 结算单_病人信息 Where 日期<Sysdate-1"
    gcnOracle.Execute gstrSQL
    '删除今天内相同业务序列号的数据
    gstrSQL = "Delete From 结算单_病人信息 Where 业务序列号='" & str业务序列号 & "'"
    gcnOracle.Execute gstrSQL
    
    blnTrans = True
    '取病人基本信息（只有一条）
    Call DebugTool("取病人基本信息")
    If Not 调用接口_指定记录集_沈阳市("Info") Then GoTo ExitSub
    gstrSQL = " Insert Into 结算单_病人信息" & _
              "(业务序列号,姓名,性别,年龄,医保号,身份证号,人员类别,是否享受公务员补助,公务员等级名称,住院号," & _
              " 医疗机构,医疗机构等级,单位名称,临床诊断,入院日期,出院日期,住院天数,病区名称,科室名称,床号)" & _
              " Values ('" & str业务序列号 & "'"
    If Not GetPatientSQL(gstrSQL, intType) Then GoTo ExitSub
    Call DebugTool("将执行的SQL：" & gstrSQL)
    gcnOracle.Execute gstrSQL
    
    '取历史业务信息（只有一条）
    Call DebugTool("取历史业务信息")
    If Not 调用接口_指定记录集_沈阳市("His") Then GoTo ExitSub
    If int业务类型 <> 2 Then
        gstrSQL = " Insert Into 结算单_历史就诊信息" & _
                  "(业务序列号,住院标志,本年费用累计,本年次数累计,本年总费用,本次总费用,起付线,个人帐户,医保基金,补充基金,公务员补助,本年次数)" & _
                  " Values ('" & str业务序列号 & "'," & IIf(bln门诊, 1, 3)
    Else
        gstrSQL = " Insert Into 结算单_历史就诊信息_特殊" & _
                  "(业务序列号,申报费用,总费用,起付线,医保基金,全自付,部分自付,比例自付,补充基金,公务员补助,本年特殊业务次数)" & _
                  " Values ('" & str业务序列号 & "'"
    End If
    If Not GetHistorySQL(gstrSQL, intType) Then GoTo ExitSub
    Call DebugTool("将执行的SQL：" & gstrSQL)
    gcnOracle.Execute gstrSQL
    
    '取分段费用信息（多条）
    Call DebugTool("取分段费用信息")
    If Not 调用接口_指定记录集_沈阳市("Seg") Then GoTo ExitSub
    gstrSQL = " Insert Into 结算单_结算汇总" & _
              "(业务序列号,支付名称,现金,个人帐户,医保基金,补充医疗,公务员补助)" & _
              " Values ('" & str业务序列号 & "'"
    If 调用接口_记录数_沈阳市 Then
        Do While True
            strTemp = gstrSQL
            Call 调用接口_读取数据_沈阳市("policy_type", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("cash_pay", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("acct_pay", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("found_pay", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("additional_pay", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("official_pay", strData)
            Call CombinateString(strTemp, strData)
            strTemp = strTemp & ")"
            Call DebugTool("将执行的SQL：" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    '取大类汇总信息（多条）
    Call DebugTool("取大类汇总信息")
    If Not 调用接口_指定记录集_沈阳市("Fee") Then GoTo ExitSub
    gstrSQL = " Insert Into 结算单_大类汇总" & _
              "(业务序列号,收费项目类型,收费项目名称,总费用,全自费,部分自负)" & _
              " Values ('" & str业务序列号 & "'"
    If 调用接口_记录数_沈阳市 Then
        Do While True
            strTemp = gstrSQL
            Call 调用接口_读取数据_沈阳市("stat_type", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("stat_name", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("zfy", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("qzf", strData)
            Call CombinateString(strTemp, strData)
            Call 调用接口_读取数据_沈阳市("blzf", strData)
            Call CombinateString(strTemp, strData)
            strTemp = strTemp & ")"
            Call DebugTool("将执行的SQL：" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    gcnOracle.CommitTrans
    提取结算单_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
ExitSub:
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function 提取结算表_沈阳市() As Boolean
    '提取门诊结算表、门诊规定病结算表及住院结算表
    Dim strStart As String, strEnd As String
    Dim blnTrans As Boolean
    On Error GoTo errHand
    If Not 医保初始化_沈阳市 Then Exit Function
    
    If Not frm日期范围_米易.Show_ME(strStart, strEnd) Then Exit Function
    strStart = Format(strStart, "yyyy-MM-dd")
    strEnd = Format(strEnd, "yyyy-MM-dd")
'    序号    入参    入参说明    最大长度    是否可为空  备注
'    1   hospital_id    医院编号        20  否
'    2   startdate      汇总起始日期        否  格式:YYYY-MM-DD
'    3   enddate        汇总终止日期        否  格式:YYYY-MM-DD
    
    gcnOracle.BeginTrans
    blnTrans = True
    Call DebugTool("初始化——清除相关表数据")
    '删除所有结算表（一般一个月取一次）
    gstrSQL = "Delete 结算表_日期范围"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Delete 住院结算表_明细清单"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Delete 门诊结算表_明细清单"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Delete 规定病结算表_明细清单"
    gcnOracle.Execute gstrSQL
    
    gstrSQL = " Insert Into 结算表_日期范围" & _
              " (开始日期,结束日期)" & _
              " Values ('" & strStart & "','" & strEnd & "')"
    Call DebugTool("将执行的SQL：" & gstrSQL)
    gcnOracle.Execute gstrSQL
    
    '先取住院结算汇总表
    Call DebugTool("取住院结算汇总表")
    If Not 住院结算表(strStart, strEnd) Then GoTo ExitSub
    
    '取门诊结算汇总表
    Call DebugTool("取门诊结算汇总表")
    If Not 门诊结算表(strStart, strEnd) Then GoTo ExitSub
    
    '取特种病结算汇总表
    Call DebugTool("取门诊规定病结算汇总表")
    If Not 特种病结算表(strStart, strEnd) Then GoTo ExitSub
    
    gcnOracle.CommitTrans
    提取结算表_沈阳市 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
ExitSub:
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Private Function 住院结算表(ByVal strStart As String, ByVal strEnd As String) As Boolean
    Dim intColumn As Integer, intColumns As Integer
    Dim strData As String, strTemp As String
    Const strColumn As String = "name|corp_name|pers_name|official|begin_date|end_date|total_declare" & _
    "|total|allself|partself|saccount|fund_offi|fund_bs|fund_add|cash_sta|sacc_sta|offi_sta|cash_bs|sacc_bs" & _
    "|offi_bs|cash_add|sacc_add|offi_add"
    Dim arrColumn
    On Error GoTo errHand
'    1   name           姓名            10
'    2   corp_name      单位名称        50
'    3   pers_name      人员类别        20
'    4   official       是否公务员      1   "√"-是,""-否
'    5   begin_date     入院日期        格式:YYYY-MM-DD
'    6   end_date       出院日期        格式:YYYY-MM-DD
'    7   total_declare  本次进入统筹基数费用    18  单位：元
'    8   total          总费用          18  单位：元
'    9   allself        现金支付自费    18  单位：元
'    10  partself       现金支付自负    18  单位：元
'    11  saccount       中心应付费用个人帐户    18  单位：元
'    12  fund_offi      中心应付费用公务员补助  18  单位：元
'    13  fund_bs        中心应付费用统筹基金    18  单位：元
'    14  fund_add       补充保险应付    18  单位：元
'    15  cash_sta       起付线现金      18  单位：元
'    16  sacc_sta       起付线个人帐户  18  单位：元
'    17  offi_sta       起付线公务员补助18  单位：元
'    18  cash_bs        基本医疗保险现金18  单位：元
'    19  sacc_bs        基本医疗保险个人帐户    18  单位：元
'    20  offi_bs        基本医疗保险公务员补助  18  单位：元
'    21  cash_add       补充医疗保险现金        18  单位：元
'    22  sacc_add       补充医疗保险个人帐户    18  单位：元
'    23  offi_add       补充医疗保险公务员补助  18  单位：元
    gstrField_沈阳市 = "hospital_id||startdate||enddate"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & strStart & "||" & strEnd
    If Not 调用接口_准备_沈阳市(Function_沈阳市.结算汇总表_住院) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    
    arrColumn = Split(strColumn, "|")
    intColumns = UBound(arrColumn)
    
    If Not 调用接口_指定记录集_沈阳市("His") Then Exit Function
    gstrSQL = " Insert Into 住院结算表_明细清单" & _
              "(姓名,单位名称,人员类别,公务员,入院日期,出院日期,进入统筹,总费用,总自费,总自负,个人帐户," & _
              " 公务员补助,统筹基金,补充保险,起付线,起付线个人帐户,起付线公务员补助,基本医疗保险现金," & _
              " 基本医疗保险个人帐户,基本医疗保险公务员补助,补充医疗保险现金,补充医疗保险个人帐户,补充医疗保险公务员补助)" & _
              " Values ("
    If 调用接口_记录数_沈阳市 Then
        Do While True
            strTemp = gstrSQL
            For intColumn = 0 To intColumns
                Call 调用接口_读取数据_沈阳市(arrColumn(intColumn), strData)
                Call CombinateString(strTemp, strData)
            Next
            strTemp = strTemp & ")"
            Call DebugTool("将执行的SQL：" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    住院结算表 = True
errHand:
End Function

Private Function 门诊结算表(ByVal strStart As String, ByVal strEnd As String) As Boolean
    Dim intColumn As Integer, intColumns As Integer
    Dim strData As String, strTemp As String
    Const strColumn As String = "name|corp_name|pers_name|official|end_date|total|allself|partself|saccount|fund_offi|fund_bs"
    Dim arrColumn
    On Error GoTo errHand
'    1   name           姓名            10
'    2   corp_name      单位名称        50
'    3   pers_name      人员类别        20
'    4   official       是否公务员      1   "√"-是,""-否
'    5   end_date       出院日期        格式:YYYY-MM-DD
'    6   total          总费用          18  单位：元
'    7   allself        现金支付自费    18  单位：元
'    8   partself       现金支付自负    18  单位：元
'    9   saccount       中心应付费用个人帐户    18  单位：元
'    10  fund_offi      中心应付费用公务员补助  18  单位：元
'    11  fund_bs        中心应付费用统筹基金    18  单位：元
    gstrField_沈阳市 = "hospital_id||startdate||enddate"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & strStart & "||" & strEnd
    If Not 调用接口_准备_沈阳市(Function_沈阳市.结算汇总表_门诊) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    
    arrColumn = Split(strColumn, "|")
    intColumns = UBound(arrColumn)
    
    If Not 调用接口_指定记录集_沈阳市("His") Then Exit Function
    gstrSQL = " Insert Into 门诊结算表_明细清单" & _
              "(姓名,单位名称,人员类别,公务员,出院日期,总费用,总自费,总自负,个人帐户,公务员补助,统筹基金)" & _
              " Values ("
    If 调用接口_记录数_沈阳市 Then
        Do While True
            strTemp = gstrSQL
            For intColumn = 0 To intColumns
                Call 调用接口_读取数据_沈阳市(arrColumn(intColumn), strData)
                Call CombinateString(strTemp, strData)
            Next
            strTemp = strTemp & ")"
            Call DebugTool("将执行的SQL：" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    门诊结算表 = True
errHand:
End Function

Private Function 特种病结算表(ByVal strStart As String, ByVal strEnd As String) As Boolean
    Dim intColumn As Integer, intColumns As Integer
    Dim strData As String, strTemp As String
    Const strColumn As String = "name|corp_name|pers_name|official|end_date|total_declare" & _
    "|total|allself|partself|saccount|fund_offi|fund_bs|fund_add|cash_sta|sacc_sta|offi_sta|cash_bs|sacc_bs" & _
    "|offi_bs|cash_add|sacc_add|offi_add"
    Dim arrColumn
    On Error GoTo errHand
'    1   name           姓名            10
'    2   corp_name      单位名称        50
'    3   pers_name      人员类别        20
'    4   official       是否公务员      1   "√"-是,""-否
'    5   end_date       出院日期        格式:YYYY-MM-DD
'    6   total_declare  本次进入统筹基数费用    18  单位：元
'    7   total          总费用          18  单位：元
'    8   allself        现金支付自费    18  单位：元
'    9   partself       现金支付自负    18  单位：元
'    10  saccount       中心应付费用个人帐户    18  单位：元
'    11  fund_offi      中心应付费用公务员补助  18  单位：元
'    12  fund_bs        中心应付费用统筹基金    18  单位：元
'    13  fund_add       补充保险应付    18  单位：元
'    14  cash_sta       起付线现金      18  单位：元
'    15  sacc_sta       起付线个人帐户  18  单位：元
'    16  offi_sta       起付线公务员补助18  单位：元
'    17  cash_bs        基本医疗保险现金18  单位：元
'    18  sacc_bs        基本医疗保险个人帐户    18  单位：元
'    19  offi_bs        基本医疗保险公务员补助  18  单位：元
'    20  cash_add       补充医疗保险现金        18  单位：元
'    21  sacc_add       补充医疗保险个人帐户    18  单位：元
'    22  offi_add       补充医疗保险公务员补助  18  单位：元
    gstrField_沈阳市 = "hospital_id||startdate||enddate"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & strStart & "||" & strEnd
    If Not 调用接口_准备_沈阳市(Function_沈阳市.结算汇总表_门诊规定病) Then Exit Function
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Function
    If Not 调用接口_执行_沈阳市 Then Exit Function
    
    arrColumn = Split(strColumn, "|")
    intColumns = UBound(arrColumn)
    
    If Not 调用接口_指定记录集_沈阳市("His") Then Exit Function
    gstrSQL = " Insert Into 规定病结算表_明细清单" & _
              "(姓名,单位名称,人员类别,公务员,出院日期,进入统筹,总费用,总自费,总自负,个人帐户," & _
              " 公务员补助,统筹基金,补充保险,起付线,起付线个人帐户,起付线公务员补助,基本医疗保险现金," & _
              " 基本医疗保险个人帐户,基本医疗保险公务员补助,补充医疗保险现金,补充医疗保险个人帐户,补充医疗保险公务员补助)" & _
              " Values ("
    If 调用接口_记录数_沈阳市 Then
        Do While True
            strTemp = gstrSQL
            For intColumn = 0 To intColumns
                Call 调用接口_读取数据_沈阳市(arrColumn(intColumn), strData)
                Call CombinateString(strTemp, strData)
            Next
            strTemp = strTemp & ")"
            Call DebugTool("将执行的SQL：" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    特种病结算表 = True
errHand:
End Function

Private Function GetPatientSQL(strSQL As String, ByVal intType As Integer) As Boolean
    Const str门诊 As String = ",name,sex,age,insr_code,idcard,pers_name,official_name,hospital_name,corp_name,disease,begin_date,grade_name,official,"
    Const str门诊规定病 As String = ",name,sex,age,patient_id,idcard,pers_name,official_name,hospital_name,corp_name,disease,begin_date,grade_name,official,"
    Const str住院 As String = ",name,sex,age,insr_code,idcard,pers_name,official_name,official,patient_id,hospital_name,grade_name,corp_name,disease,begin_date,end_date,days,in_area_name,in_dept_name,in_bed,"
    Dim strCompare As String
    On Error GoTo errHand
    'intType：1-门诊;2-门诊规定病;3-住院
    
    Select Case intType
    Case 1
        strCompare = str门诊
    Case 2
        strCompare = str门诊规定病
    Case 3
        strCompare = str住院
    End Select
    
    Call CombinateString(gstrSQL, GetSQL(strCompare, "name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "sex"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "age"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "insr_code"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "idcard"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "pers_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "official_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "official"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "patient_id"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "hospital_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "grade_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "corp_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "disease"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "begin_date"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "end_date"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "days"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "in_area_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "in_dept_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "in_bed"))
    gstrSQL = gstrSQL & ")"
    GetPatientSQL = True
errHand:
End Function

Private Function GetSQL(ByVal strCompareSQL As String, ByVal strColumn As String) As String
    Dim strData As String
    If InStr(1, strCompareSQL, "," & strColumn & ",") <> 0 Then
        If Not 调用接口_读取数据_沈阳市(strColumn, strData) Then Exit Function
        GetSQL = strData
    End If
End Function

Private Function GetHistorySQL(strSQL As String, ByVal intType As Integer) As Boolean
    Const str门诊 As String = ",declare_fee,inhosp_count,total_fee,cur_total_fee,start_pay,self_pay,fund_pay,additional_pay,official_pay,biz_times,"
    Const str门诊规定病 As String = ",declare_fee,total_fee,start_pay,fund_pay,all_self_pay,self_pay,percent_pay,additional_pay,official_pay,biz_times,"
    Const str住院 As String = ",declare_fee,inhosp_count,total_fee,cur_total_fee,start_pay,self_pay,fund_pay,additional_pay,official_pay,biz_times,"
    Dim strCompare As String
    On Error GoTo errHand
    'intType：1-门诊;2-门诊规定病;3-住院
    
    Select Case intType
    Case 1
        strCompare = str门诊
    Case 2
        strCompare = str门诊规定病
    Case 3
        strCompare = str住院
    End Select
    
    If intType = 2 Then
        Call CombinateString(gstrSQL, GetSQL(strCompare, "declare_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "inhosp_count"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "total_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "cur_total_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "start_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "self_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "fund_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "additional_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "official_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "biz_times"))
    Else
        Call CombinateString(gstrSQL, GetSQL(strCompare, "declare_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "total_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "start_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "fund_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "all_self_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "self_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "percent_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "additional_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "official_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "biz_times"))
    End If
    gstrSQL = gstrSQL & ")"
    GetHistorySQL = True
errHand:
End Function

Private Sub CombinateString(strSQL As String, ByVal strData As String)
    '组合字符串
    'intType：1-普通;2-字符串
    strSQL = strSQL & "," & "'" & strData & "'"
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub 更新病种_沈阳市(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    Dim arrPatient
    Dim str就诊日期 As String, str入院病种 As String, str出院病种 As String, str并发症 As String
    Dim rsTemp As New ADODB.Recordset
    
    '更新在院病人的病种信息
    Call 获取病人基本信息(lng病人ID, False)
    arrPatient = Split(获取病人相关信息(lng病人ID, lng主页ID), "||")
        
    '读取病人的出院情况
    gstrSQL = "Select decode(出院方式,'转院',3,0) 出院方式,入院日期 From 病案主页 " & _
            " Where 病人ID = [1] And 主页ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "出院方式", lng病人ID, lng主页ID)
    str就诊日期 = Format(rsTemp!入院日期, "yyyy-MM-dd")
    
    If Not frm病种选择_沈阳.ShowSelect(TYPE_沈阳市, lng病人ID, lng主页ID, str入院病种, str出院病种, str并发症) Then Exit Sub
    
    '更新病人入院情况及出院病种
    '    1   hospital_id    医疗机构编码20  否
    '    2   serial_no      业务序列号  12  否
    '    3   busi_type      业务类型    2   否  "12"：住院
    '    4   staff_no       操作员工号  5   否
    '    5   staff_name     操作员姓名  10  否
    '    6   begin_date     就诊时间        是  格式：YYYY-MM-DD
    '    7   in_dept        入院科室编号3   是
    '    8   in_dept_name   入院科室名称20  是
    '    9   in_area        入院病区编号3   是
    '    10  in_area_name   入院病区名称20  是
    '    11  in_bed         入院病床编号10  是
    '    12  bed_type       床位类型    1   是  "0"：普通床位；"1"：急救；"2"：留观；"3"：高干
    '    13  patient_id     住院号      20  是
    '    14  old_patient_id 原住院号    20  是
    '    15  in_disease     入院诊断    20  是  疾病编码
    '    16  note           备注        100 是
    '    17  fin_disease    出院病种
    gstrField_沈阳市 = "hospital_id||serial_no||busi_type||staff_no||staff_name||begin_date||" & _
                    "in_dept||in_dept_name||in_area||in_area_name||in_bed||bed_type||patient_id||old_patient_id||in_disease||note||fin_disease"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.业务序列号 & "||" & _
                    gCominfo_沈阳市.业务类型 & "||" & gCominfo_沈阳市.操作员工号 & "||" & _
                    gstrUserName & "||" & str就诊日期 & "||" & arrPatient(入院科室编号) & "||" & _
                    arrPatient(入院科室名称) & "||" & arrPatient(入院病区编号) & "||" & _
                    arrPatient(入院病区名称) & "||" & arrPatient(入院病床编号) & "||" & _
                    arrPatient(床位类型) & "||" & arrPatient(住院号) & "||" & _
                    arrPatient(住院号) & "||" & str入院病种 & "||" & str并发症 & "||" & str出院病种
    If Not 调用接口_准备_沈阳市(Function_沈阳市.住院信息_修改) Then Exit Sub
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Sub
    If Not 调用接口_执行_沈阳市 Then Exit Sub
End Sub

Private Function FirstAid() As Boolean
    '检查本工作站是否是急救用药专用机
    FirstAid = (GetSetting("ZLSOFT", "沈阳医保工具包", "急救用药专用机", 0) = 1)
End Function
