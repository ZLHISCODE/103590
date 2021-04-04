Attribute VB_Name = "mdl成都内江"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

Private mblnInit As Boolean
Public Enum 业务类型_成都内江
    读病人信息_内江 = 0
    更改密码_内江
    获取帐户余额_内江
    门诊明细写入_内江
    门诊消费确认_内江
    门诊消费取消_内江
    住院登记_内江
    住院交易上传_内江
    住院产易上传取消_内江
    出院登记上传_内江
    出院登记确认_内江
    获取单位欠缴情况_内江
    初始化函数_内江
    网上对帐_内江 '20051020 陈东
    并发症申请上传_内江
End Enum

Private gInitCard As Boolean                '初始化了卡的
Private Type InitbaseInfor
    医院编码 As String                      '初始医院编码
    串号号_内江 As Integer
    读卡器_内江 As Integer                  '0-明化,1-德森公司
    
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    解析卡内数据 As Boolean
End Type
Public InitInfor_成都内江 As InitbaseInfor
Private mblnStartTran   As Boolean '启动了事务的
Private Type 病人身份
        卡号       As String
        个人编号   As String
        身份证号   As String
        姓名       As String
        性别       As String
        工况类别   As String
        出生日期   As String
        单位号码   As String
        统筹编号   As String
        制卡日期   As String
        卡有效期   As String
        补卡次数   As String
        制卡单位   As String
        年龄        As Integer
        帐户余额    As Double
        在职情况    As String
        
        住院流水号 As String
        交易类别 As String
        lng病人ID   As Long
        
        费用总额  As Double
        结算方式    As String   '结算方式串
        病种编码    As String
        病种名称    As String
        出院类别    As String
        生育总费用 As Double
End Type

Private Type 结算数据
    待遇标志 As String
    医保交易流水号 As String
    医保内费用   As Double
    医保外费用   As Double
    基本医保支付    As Double
    高额医保支付    As Double
    公务员医疗补助  As Double
    帐户可用余额  As Double
    帐户支付        As Double
    比例支付        As Double
    起付标准        As Double
    结算标志        As Byte '0-门诊,1-住院
    结帐ID          As Long
    生育支付      As Double '20051020 新增
    生育盈亏       As Double '20051118 新增
End Type

Public g病人身份_成都内江 As 病人身份
Public gcnOracle_成都内江 As ADODB.Connection     '中间库连接
Private g结算数据   As 结算数据



'****************************************************************************************************************************************************************************************************************************************
'1 相关读卡组件函数
'****************************************************************************************************************************************************************************************************************************************
'   0-读病人信息函数(明华的)
Private Declare Function GetCardInfo_MW Lib "NeijCard.dll" Alias "GetCardInfo" (ByVal lngPort As Long, ByVal strPassWord As String, str卡号 As String, _
         str个人编号 As String, str身份证号 As String, STR姓名 As String, str性别 As String, _
        str工况类别 As String, str出生日期 As String, str单位号码 As String, str统筹编号 As String, _
        str制卡日期 As String, str卡有效期 As String, str补卡次数 As String, str制卡单位 As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'明华:
'函数原型:function GetCardInfo(port: integer;UserPassword:PChar; var CardNum,PersonNum,
'                   IDNum,Name,Sex,PersonKind,Birthday,DeptNum,Zone,MAKEDATE,EXPIREDATE,REISSUE,MAKEDEPT: PChar):integer
'参数:  a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
'       b)  UserPassword：输入参数，为用户密码，要求长度为6，字符串中只能包含0到9的数字；
'       c)  CardNum：输出参数，为卡号，长度为10；
'       d)  PersonNum：输出参数，为个人编号（医保编号），长度为8；
'       e)  IDNum：输出参数，为身份证号码，长度为18；
'       f)  Name：输出参数，为姓名，长度为20；
'       g)  Sex：输出参数，为性别编码，长度为1，其中'1'为男，'2'为女；
'       h)  PersonKind：输出参数，为工况类别，长度为1；
'       i)  Birthday：输出参数，为出生日期，长度为8，例如1982年6月23日表示为'19820623'；
'       j)  DeptNum：输出参数，为单位号码，长度为6；
'       k)  Zone：输出参数，为统筹地区编码，长度为1；
'       l)  MAKEDATE：输出参数，为制卡日期，长度为8，表示方式同出生日期；
'       m)  EXPIREDATE：输出参数，为卡有效日期（卡的有效期为99年，如制卡日期为20021101，则卡有效日期为21011101），长度为8，表示方式同出生日期；
'       n)  REISSUE：输出参数，为补卡次数，长度为2，例如：首次制卡，补卡次数为'00'，第一次补卡，补卡次数为'01'，以次类推；
'       o)  MAKEDEPT：输出参数，为制卡单位，长度为1，例如：'0'表示制卡方为温州德森公司。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   1-读病人信息函数(科瑞奇)
Private Declare Function GetCardInfo_KRQ Lib "NeijCard.dll" Alias "GetCardInfo" (ByVal lngPort As Long, str卡号 As String, _
        str个人编号 As String, str身份证号 As String, STR姓名 As String, str性别 As String, _
        str工况类别 As String, str出生日期 As String, str单位号码 As String, str统筹编号 As String, _
        str制卡日期 As String, str卡有效期 As String, str补卡次数 As String, str制卡单位 As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'科瑞奇
'说明:  与明细相比,没有密码的输入
'函数原型:function GetCardInfoForKRQ(port: integer; var CardNum,PersonNum,IDNum,Name,Sex,PersonKind,Birthday,DeptNum,Zone,
'           MAKEDATE,EXPIREDATE,REISSUE,MAKEDEPT: PChar):integer;
'参数:  a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
'       b)  CardNum：输出参数，为卡号，长度为10；
'       c)  PersonNum：输出参数，为个人编号（医保编号），长度为8；
'       d)  IDNum：输出参数，为身份证号码，长度为18；
'       e)  Name：输出参数，为姓名，长度为20；
'       f)  Sex：输出参数，为性别编码，长度为1，其中'1'为男，'2'为女；
'       g)  PersonKind：输出参数，为工况类别，长度为1；
'       h)  Birthday：输出参数，为出生日期，长度为8，例如1982年6月23日表示为'19820623'；
'       i)  DeptNum：输出参数，为单位号码，长度为6；
'       j)  Zone：输出参数，为统筹地区编码，长度为1；
'       k)  MAKEDATE：输出参数，为制卡日期，长度为8，表示方式同出生日期；
'       l)  EXPIREDATE：输出参数，为卡有效日期（卡的有效期为99年，如制卡日期为20021101，则卡有效日期为21011101），长度为8，表示方式同出生日期；
'       m)  REISSUE：输出参数，为补卡次数，长度为2，例如：首次制卡，补卡次数为'00'，第一次补卡，补卡次数为'01'，以次类推；
'       n)  MAKEDEPT：输出参数，为制卡单位，长度为1，例如：'0'表示制卡方为温州德森公司。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   2-修改密码
Private Declare Function ChangePassword Lib "NeijCard.dll" (ByVal lngPort As Long, ByVal strOldPassWord As String, ByVal strNewPassWord As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'说明:  与明细相比,没有密码的输入
'函数原型:function ChangePassword(port:integer;OldPassword,NewPassword:PChar):integer;
'参数:  a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
'       b)  OldPassword：输入参数，为原密码，要求长度为6，字符串中只能包含0到9的数字；
'       c)  NewPassword：输入参数，为新密码，要求长度为6，字符串中只能包含0到9的数字。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'****************************************************************************************************************************************************************************************************************************************
'2 业务对象
'****************************************************************************************************************************************************************************************************************************************
Public gobj成都内江 As Object
'
'Public gobj成都内江 As New clsNjYh

Public Function 医保初始化_成都内江() As Boolean
    
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    If mblnInit Then
        医保初始化_成都内江 = True
        Exit Function
    End If
    
    GetRegInFor g公共全局, "医保", "读卡器", strReg
    InitInfor_成都内江.读卡器_内江 = Val(strReg)
    
    
    GetRegInFor g公共全局, "医保", "串口号", strReg
    
    InitInfor_成都内江.串号号_内江 = IIf(strReg = "", 1, Val(strReg))
        
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_成都内江.模拟数据 = True
    Else
        InitInfor_成都内江.模拟数据 = False
    End If
    
    Call GetRegInFor(g公共模块, "操作", "解析卡内数据", strReg)
    If Val(strReg) = 1 Then
        InitInfor_成都内江.解析卡内数据 = True
    Else
        InitInfor_成都内江.解析卡内数据 = False
    End If
    InitInfor_成都内江.解析卡内数据 = InitInfor_成都内江.解析卡内数据 Or InitInfor_成都内江.模拟数据
    
    
    '创建医保对象
    If gobj成都内江 Is Nothing Then
        Err = 0
        On Error Resume Next
        Set gobj成都内江 = CreateObject("SocketOcxForNC.SocketOcxForNC")
        
        If Err <> 0 Then
                ShowMsgbox "不能创建医保接口,请检查SocketOcxForNC.ocx是否正常注册!"
                Exit Function
        End If
    End If
    
    
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & TYPE_成都内江
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "读取医院编码")
    InitInfor_成都内江.医院编码 = Nvl(rsTemp!医院编码)
    If Open中间库 = False Then Exit Function
    
    mblnInit = True
    医保初始化_成都内江 = True
End Function
Private Function Open中间库() As Boolean
    '连接中间库
    '中间库连接
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strServer As String, strPass As String, strReg As String
    Dim StrInput As String, strOutput As String
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=" & TYPE_成都内江
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "获取相关参数值")
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_成都内江 = New ADODB.Connection

    If OraDataOpen(gcnOracle_成都内江, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '检查网络是否畅通无阻
          
    'GetRegInFor g公共全局, "医保", "ConfigFileName", strReg
    'StrInput = strReg
    'GetRegInFor g公共全局, "医保", "HostPort", strReg
    'StrInput = StrInput & vbTab & strReg
    'GetRegInFor g公共全局, "医保", "IPAddress", strReg
    'StrInput = StrInput & vbTab & strReg
    
    'If 业务请求_成都内江(初始化函数_内江, StrInput, strOutput) = False Then Exit Function
    
    Open中间库 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保终止_成都内江() As Boolean
    '结束读写卡组件
    Dim strReg As String
    mblnInit = False
    Err = 0
    On Error Resume Next
    
    Set gobj成都内江 = Nothing
    If gcnOracle_成都内江.State = 1 Then
        gcnOracle_成都内江.Close
    End If
    医保终止_成都内江 = True
End Function

Public Function 身份标识_成都内江(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Err = 0
    On Error GoTo errHand:
    身份标识_成都内江 = frmIdentify成都内江.GetPatient(bytType, lng病人ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_成都内江 = ""
End Function

Public Function 个人余额_成都内江(ByVal lng病人ID As Long) As Currency
    '功能: 提取参保病人个人帐户余额
    '返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID='" & lng病人ID & "' and 险类=" & TYPE_成都内江
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "读取个人帐户余额")
    
    If rsTemp.EOF Then
        个人余额_成都内江 = 0
    Else
        If rsTemp("帐户余额") > 0 Then
        个人余额_成都内江 = rsTemp("帐户余额")
        Else
        个人余额_成都内江 = 0
        End If
    End If
End Function

Public Function 门诊虚拟结算_成都内江(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
  
    Dim StrInput As String, strOutput As String
    Dim strArr
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID  As Long
    Dim str操作员编码 As String
    Dim str收费细目ID As String '保存所有收费细目的ID，用来判断是否存在重复的项目
    Err = 0: On Error GoTo errHand:
    
    Call DebugTool("进入门诊虚拟结算")
    
    With g结算数据
        .比例支付 = 0
        .待遇标志 = ""
        .高额医保支付 = 0
        .公务员医疗补助 = 0
        .起付标准 = 0
        .医保交易流水号 = ""
        .医保内费用 = 0
        .医保外费用 = 0
        .帐户可用余额 = 0
        .帐户支付 = 0
        .生育支付 = 0 '20051021 add
    End With
    
    '检查是否存在重复的项目
    With rs明细
        Do While Not .EOF
            If InStr(1, str收费细目ID & ",", "," & !收费细目ID & ",") = 0 Then
                str收费细目ID = str收费细目ID & "," & !收费细目ID
            Else
                Err.Raise 9000, gstrSysName, "存在重复的收费细目，请合并后再进行预结算！"
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With

    lng病人ID = rs明细("病人ID")
    str操作员编码 = Nvl(rs明细!开单人)
    If g病人身份_成都内江.lng病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人还没有经过身份验证，不能进行医保结算。"
        Exit Function
    End If
    
    g结算数据.结算标志 = 0
    g结算数据.结帐ID = 0
        
    '写入明细
    If 门诊明细写入(rs明细, True) = False Then Exit Function
    If 结算方式更正(1, str结算方式) = False Then Exit Function
    
    门诊虚拟结算_成都内江 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function 获取个人帐户支付() As Double
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取个人帐户值(从预交记录中获取)
    '--入参数:
    '--出参数:
    '--返  回:成功,返回本次个人帐户支付,否则返回0
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select * From 病人预交记录 where 结帐ID=[1] and  结算方式='个人帐户'"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取个人帐户支付", g结算数据.结帐ID)
    If Not rsTemp.EOF Then
        获取个人帐户支付 = Nvl(rsTemp!冲预交, 0)
    End If
    
End Function


Public Function 门诊结算_成都内江(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency, Optional ByRef strAdvance = "") As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    Dim StrInput As String, strOutput As String
    Dim strArr
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim lng病人ID  As Long
    Dim str操作员编码 As String
    Err = 0: On Error GoTo errHand:
    
    Call DebugTool("进入门诊结算")
    'Modified by ZYB 22051123 门诊虚拟结算已上传明细，结算不再进行合法性检查及明细上传
    '#################################################################################
'    gstrSQL = "Select 收费细目ID From 病人费用记录  where 结帐id=" & lng结帐ID & " group by 收费细目iD   having Count(收费细目id)>=2 "
'    Call OpenRecordset(rsTemp, "判断明细是否重复")
'    If Not rsTemp.EOF Then
'        MsgBox "存在重复的收费细目,请合并后再结算!"
'        Exit Function
'    End If
'
'
'    With g结算数据
'        .比例支付 = 0
'        .待遇标志 = ""
'        .高额医保支付 = 0
'        .公务员医疗补助 = 0
'        .起付标准 = 0
'        .医保交易流水号 = ""
'        .医保内费用 = 0
'        .医保外费用 = 0
'        .帐户可用余额 = 0
'        .帐户支付 = 0
'        .生育支付 = 0 '20051021 add
'    End With
'
    '#################################################################################
    
    gstrSQL = "" & _
    "   Select a.*,a.付数*a.数次 as 数量,a.实收金额/(nvl(a.付数,1)*nvl(a.数次,1)) as 单价 " & _
    "   From 门诊费用记录 a " & _
    "   Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Call zlDatabase.OpenRecordset(rs明细, gstrSQL, "获取明细记录")
    If rs明细.EOF = True Then
        Err.Raise 9000, gstrSysName, "没有填写收费记录"
        Exit Function
    End If

    lng病人ID = rs明细("病人ID")
    str操作员编码 = Nvl(rs明细!操作员姓名)
    If g病人身份_成都内江.lng病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人还没有经过身份 验证，不能进行医保结算。"
        Exit Function
    End If
    
    g结算数据.结算标志 = 0
    g结算数据.结帐ID = lng结帐ID
        
    '写入明细
'    If 门诊明细写入(rs明细, False) = False Then Exit Function
    If 结算方式更正(1, strAdvance) = False Then Exit Function
    
    '获取帐户支付
    strAdvance = ""         '强制赋为空，告诉调用者不需要校正
    g结算数据.帐户支付 = 获取个人帐户支付()
    '消费进行确认
    '输入参数: 个人编号    String(8)   In
    '          社保卡号码  String(10)  In
    '          医院代码    String(5)   In
    '          操作员卡号码    String(10)  In
    '          统筹地区编码    String(1)   In
    '          医保交易流水号  String(20)  In
    '          交易类别    String(1)   In
    '          个人帐户支付    String(10)  In
    With g病人身份_成都内江
        StrInput = Rpad(.个人编号, 8)
        StrInput = StrInput & vbTab & Rpad(.卡号, 10)
        StrInput = StrInput & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
        StrInput = StrInput & vbTab & Substr(Rpad(str操作员编码, 10), 1, 10)
        StrInput = StrInput & vbTab & Rpad(.统筹编号, 1)
        StrInput = StrInput & vbTab & Substr(Rpad(g结算数据.医保交易流水号, 20), 1, 20)
        StrInput = StrInput & vbTab & Rpad(.交易类别, 1)
        StrInput = StrInput & vbTab & Lpad(Round(g结算数据.帐户支付 * 100), 10, "0")
    End With
    '调用结算
    Call DebugTool("准备调用门诊消费确认")
    If 业务请求_成都内江(门诊消费确认_内江, StrInput, strOutput) = False Then Exit Function
    Call DebugTool("调用门诊消费确认结束")
    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(无),帐户累计支出_IN(无),累计进入统筹_IN(无),累计统筹报销_IN(无),住院次数_IN(无),起付线(无),封顶线_IN(帐户可用余额),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(无),超限自付金额_IN(生育支付),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN
    'Modified by ZYB 22051123 新老版本都不存在校正的问题了
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_成都内江 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "0,0,0,0,0,0," & g结算数据.帐户可用余额 & ",0," & _
            g病人身份_成都内江.费用总额 & "," & g结算数据.医保内费用 & "," & g结算数据.医保外费用 & "," & _
           "0,0,0," & g结算数据.生育支付 & "," & g结算数据.帐户支付 & ",'" & _
            g结算数据.医保交易流水号 & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
    门诊结算_成都内江 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function Get交易流水号(ByVal str发生日期 As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取交易流水号
    '--入参数:str发生日期-以YYMMDD格式传入
    '--出参数:
    '--返  回:交易流水号
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 医院交易流水号_ID.nextval as 序列 From dual"
    OpenRecordset_成都内江 rsTemp, "获取交易流水号"
    Get交易流水号 = InitInfor_成都内江.医院编码 & str发生日期 & Lpad(Nvl(rsTemp!序列), 7, "0")
End Function

Private Function 门诊明细写入(ByVal rs明细 As ADODB.Recordset, Optional ByVal bln虚拟 As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传门诊明细数据
    '--入参数:rs明细-明细记录
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    
    Dim StrInput As String, strOutput As String, str明细 As String
    Dim strInsert As String
    Dim lngSumLen As Long
    
    Dim str规格 As String
    Dim str保险编码 As String
    Dim str大类编码 As String
    Dim str交易流水号 As String
    Dim lng处方条数 As Long
    
    Dim strArr
    
    门诊明细写入 = False
    
    DebugTool "进入门诊明细上传接口"
    
    g病人身份_成都内江.费用总额 = 0
    g病人身份_成都内江.生育总费用 = 0
       
    
    Err = 0
    On Error GoTo errHand:
    str明细 = ""
    '然后插入处方明细
    str交易流水号 = Get交易流水号(Format(zlDatabase.Currentdate, "yyyymmdd"))
    
    With rs明细
        lng处方条数 = 0
        Do While Not .EOF
            
            If Val(Nvl(rs明细("实收金额"), 0)) <> 0 Then
                '处方明细
                '1   处方项目种类 Varchar2(1)    '1'：药品编码   '2'：服务项目
                '2   处方项目代码    Varchar2(20)    "药品编码"或者"诊疗项目编码"
                '3   数量    Varchar2(10)    实际数量*100上传
                '4   规格    Varchar2(10)    汉字算2字节
                '5   单项费用    Varchar2(10)    由医院上传(主要传什么?)
                
                gstrSQL = "select A.名称,A.编码,A.类别,A.规格,A.计算单位,B.项目编码,B.附注,B.是否医保,A.计算单位,E.规格,G.名称 剂型,B.大类编码 " & _
                          "from 收费细目 A," & _
                          "         (   Select a.*,b.大类编码 " & _
                          "             From 保险支付项目 a,保险项目 b" & _
                          "             where a.险类=b.险类 and a.项目编码=b.编码 and A.收费细目ID =[1] and a.险类=[2]) B,药品目录 E ,药品信息 F,药品剂型 G " & _
                          "where A.ID=[1] and A.ID=B.收费细目ID(+) " & _
                         "        AND A.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保项目", CLng(rs明细!收费细目ID), TYPE_成都内江)
                
                If rsTemp.EOF Then
                    Err.Raise 9000, gstrSysName, "存在未对码的项目,请在保险项目管理中进行对码!"
                    Exit Function
                End If
                
                str规格 = Nvl(rsTemp!规格)
                str保险编码 = Nvl(!保险编码)
                
                '如果保险编码为空，则需要用户选择编码
                If str保险编码 = "" Then
                    str保险编码 = GetItemInsure_成都内江(0, !收费细目ID, True)
                End If
                If str保险编码 = "" Then str保险编码 = Nvl(rsTemp!项目编码)
                
                '取大类编码
                gstrSQL = "Select 大类编码 From 保险项目 Where 险类=" & TYPE_成都内江 & " And 编码='" & str保险编码 & "'"
                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "取大类编码")
                str大类编码 = Nvl(rsTemp!大类编码)
                
                str明细 = str明细 & Substr(Rpad(str大类编码, 1), 1, 1)
                str明细 = str明细 & Rpad(str保险编码, 20)
                str明细 = str明细 & Lpad(Nvl(!数量) * 100, 10, "0")
                str明细 = str明细 & Rpad(Nvl(str规格), 10)
                str明细 = str明细 & Lpad(Nvl(!实收金额) * 100, 10, "0")
                                
                'Beging 20051025 add
                '门诊部分:A.不能用“生育费用”(见生育项目代码.xls后4项)，可以用“计划生育费用
                If str保险编码 = "M10000116" Or str保险编码 = "M10000117" Or str保险编码 = "M10000118" Or str保险编码 = "M10000119" Then
                    Err.Raise 9000, gstrSysName, "门诊不能使用生育费用!"
                    Exit Function
                End If
                'End
                
                '为病人费用记录打上标记，以便随时上传
                'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                '摘要值:医院交易流水号
                'Modified by ZYB 20051123 虚拟结算时未保存处方，无法更新
                If Not bln虚拟 Then
                    gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL," & Nvl(str大类编码, "NULL") & ",NUll,'" & str保险编码 & "',1,'" & str交易流水号 & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
                End If
                lng处方条数 = lng处方条数 + 1
            End If
            
            'Beging 20051027 陈东
            If Substr(Rpad(str大类编码, 1), 1, 1) <> "4" Then
                g病人身份_成都内江.费用总额 = g病人身份_成都内江.费用总额 + Nvl(rs明细!实收金额, 0)
            Else
                g病人身份_成都内江.生育总费用 = g病人身份_成都内江.生育总费用 + Nvl(rs明细!实收金额, 0)
            End If
            'End 20051027 陈东
            rs明细.MoveNext
        Loop

        If lng处方条数 > 99 Then
            Err.Raise 9000, gstrSysName, "门诊处方明细不能大于99种项目,请分成两张处方进行录入!"
            Exit Function
        End If
        
        If .RecordCount <> 0 Then
            .MoveFirst
            '输入参数：个人编号    String(8)   In
                          '       社保卡号码  String(10)  In
                          '       医院代码    String(5)   In
                          '       操作员卡号码    String(10)  In
                          '       统筹地区编码    String(1)   In
                          '       医院交易流水号  String(20)  In
                          '       交易类别    String(1)   In
                          '       处方条数    String(2)   In
                          '       处方明细    String处方条数×51  In
            StrInput = Rpad(g病人身份_成都内江.个人编号, 8)
            StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.卡号, 10)
            StrInput = StrInput & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
            If bln虚拟 Then
                StrInput = StrInput & vbTab & Rpad(UserInfo.编号, 10)
            Else
                StrInput = StrInput & vbTab & Rpad(Nvl(!操作员编号), 10)
            End If
            StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.统筹编号, 1)
            StrInput = StrInput & vbTab & Rpad(str交易流水号, 20)
            StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.交易类别, 1)
            StrInput = StrInput & vbTab & Rpad(lng处方条数, 2)
            StrInput = StrInput & vbTab & str明细
            If 业务请求_成都内江(门诊明细写入_内江, StrInput, strOutput) = False Then Exit Function
            
            '保存相关数据
            '    医院流水号_IN IN 医保消费信息.医院流水号%TYPE,
            '    病人ID_IN IN 医保消费信息.病人ID%TYPE,
            '    医保流水号_IN IN 医保消费信息.医保流水号%TYPE,
            '    医保内费用_IN IN 医保消费信息.医保内费用%TYPE,
            '    医保外费用_IN IN 医保消费信息.医保外费用%TYPE,
            '    帐户可用余额_IN IN 医保消费信息.帐户可用余额%TYPE,
            '    在职情况_IN IN 医保消费信息.在职情况%TYPE,
            '    医保项目种类_IN IN 医保消费信息.医保项目种类%TYPE,
            '    医保项目编码_IN IN 医保消费信息.医保项目编码%TYPE,
            '    医保内费用1_IN IN 医保消费信息.医保内费用1%TYPE,
            '    费用类别_IN IN 医保消费信息.费用类别%TYPE,
            '    项目费用_IN IN 医保消费信息.项目费用%TYPE
            strArr = Split(strOutput, vbTab)
            
            With g结算数据
                .医保交易流水号 = strArr(0)
                .医保内费用 = Val(strArr(1))
                .医保外费用 = Val(strArr(2))
                .帐户可用余额 = Val(strArr(3))
                .生育支付 = Val(strArr(6)) '20051020 Add
            End With
            strInsert = "ZL_医保消费信息_INSERT("
            strInsert = strInsert & "'" & str交易流水号 & "',"
            strInsert = strInsert & "" & g病人身份_成都内江.lng病人ID & ","
            strInsert = strInsert & "'" & strArr(0) & "',"
            strInsert = strInsert & "" & Val(strArr(1)) & ","
            strInsert = strInsert & "" & Val(strArr(2)) & ","
            strInsert = strInsert & "" & Val(strArr(3)) & ","
            strInsert = strInsert & "'" & strArr(5) & "',"
            
            
            '分解明细记录记录
                        
            '1   处方项目种类 Varchar2(1)    '1'：药品编码   '2'：服务项目
            '2   处方项目代码    Varchar2(20)    "药品编码"或者"诊疗项目编码"
            '3   医保内费用  Varchar2(10)    实际数量*100
            '4   费用类别 Varchar2(10)(中药?西药?化验等)
            '5   项目费用    Varchar2(10)    实际数量*100
            '6   生育支付 20051026
            str明细 = strArr(4)
            lngSumLen = zlCommFun.ActualLen(str明细)
            StrInput = ""
            Dim r As Long, i As Integer
            
            For i = 1 To lngSumLen Step 51
                r = i
                StrInput = StrInput & "'" & Substr(str明细, r, 1) & "',"
                r = r + 1
                StrInput = StrInput & "'" & Substr(str明细, r, 20) & "',"
                r = r + 20
                StrInput = StrInput & "" & Val(Substr(str明细, r, 10)) & ","
                r = r + 10
                StrInput = StrInput & "'" & Substr(str明细, r, 10) & "',"
                r = r + 10
                StrInput = StrInput & "" & Val(Substr(str明细, r, 10)) & ","
                StrInput = StrInput & "" & Val(strArr(6)) & ")"
                '组合SQL误句
                gstrSQL = strInsert & StrInput
                
                'Modified by zyb 20051123 可能会造成HIS库中存在无效的明细（点击多次预结算），中心也一样
                ExecuteProcedure_ZLNJ "插入明细数据到中间库"
                StrInput = ""
            Next
        End If
    End With
    门诊明细写入 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_成都内江(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput  As String, str流水号 As String
    Dim lng冲销ID As Long, lng病人id1 As Long
    Dim strArr
    Dim rs明细 As New ADODB.Recordset
    Dim i As Long
    Dim intMouse  As Integer
 
    On Error GoTo errHand:

    DebugTool "进入门诊结算"
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    '身份验证
    'by 20050123 gzy
    lng病人id1 = lng病人ID
    If 身份标识_成都内江(2, lng病人id1) = "" Then
        Screen.MousePointer = intMouse
        门诊结算冲销_成都内江 = False
        If lng病人id1 = 0 Then
            Exit Function
        End If
    End If
    Screen.MousePointer = intMouse
    
    DebugTool "身份验证完成"
    
    gstrSQL = "Select * From 门诊费用记录  " & _
        " Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Call zlDatabase.OpenRecordset(rs明细, gstrSQL, "获取冲销记录")
    
    
    
    g病人身份_成都内江.费用总额 = 0
    Do Until rs明细.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        str流水号 = Nvl(rs明细!摘要)
        g病人身份_成都内江.费用总额 = g病人身份_成都内江.费用总额 + Nvl(rs明细("结帐金额"), 0)
        rs明细.MoveNext
    Loop
    g病人身份_成都内江.费用总额 = Round(g病人身份_成都内江.费用总额, 2)
    
    If lng病人ID <> lng病人id1 Then
        Err.Raise 9000, gstrSysName, " 验卡病人不是当前要冲销的病人,不能冲销结算"
        Exit Function
    End If
    
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "重庆医保")
    lng冲销ID = rsTemp("结帐ID")

    

    gstrSQL = "Select * From 门诊费用记录 " & _
        " Where 结帐ID=" & lng冲销ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "获取冲销记录")
    
    DebugTool "更新摘要标志"
    Do While Not rsTemp.EOF
        '更新上传标志

        gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(rsTemp!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str流水号 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
        rsTemp.MoveNext
    Loop
    DebugTool "更新摘要标志完成"

    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & TYPE_成都内江 & " and 记录ID=" & lng结帐ID
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "获取原来的结算记录")

    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    g结算数据.医保交易流水号 = rsTemp("支付顺序号")
    
    '    个人编号    String(8)   In
    '    社保卡号码  String(10)  In
    '    医院代码    String(5)   In
    '    操作员卡号码    String(10)  In
    '    统筹地区编码    String(1)   In
    '    医保交易流水号  String(20)  In
    '    交易类别    String(1)   In
    With g病人身份_成都内江
        StrInput = Rpad(.个人编号, 8)
        StrInput = StrInput & vbTab & Rpad(.卡号, 10)
        StrInput = StrInput & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
        StrInput = StrInput & vbTab & Rpad(gstrUserName, 10)
        StrInput = StrInput & vbTab & Rpad(.统筹编号, 1)
        StrInput = StrInput & vbTab & Rpad(g结算数据.医保交易流水号, 20)
        StrInput = StrInput & vbTab & Rpad(.交易类别, 1)
    End With
    If 业务请求_成都内江(门诊消费取消_内江, StrInput, strOutput) = False Then Exit Function
    DebugTool "业务请求成功"

 '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN(帐户可用余额),实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(无),帐户累计支出_IN(无),累计进入统筹_IN(无),累计统筹报销_IN(无),住院次数_IN(无),起付线(无),封顶线_IN(帐户可用余额),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN
    DebugTool "保存结算记录"
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_成都内江 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
        "0,0,0,0,0,0,0," & -1 * Nvl(rsTemp!封顶线, 0) & "," & _
        Nvl(rsTemp!发生费用金额, 0) * -1 & "," & Nvl(rsTemp!全自付金额, 0) * -1 & "," & Nvl(rsTemp!首先自付金额, 0) * -1 & "," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0,0," & rsTemp("个人帐户支付") * -1 & ",'" & _
       g结算数据.医保交易流水号 & "',NULL,0,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新保险结算信息")
    DebugTool "门诊结算冲销完成"
    
    门诊结算冲销_成都内江 = True
    Exit Function
errHand::
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Function 病人入院登记处理(lng病人ID As Long, lng主页ID As Long) As Boolean
    '进行门诊登记
    Dim StrInput As String, strOutput As String
    Dim str交易流水号 As String
    Dim rsTemp As New ADODB.Recordset, rsRydj As New ADODB.Recordset
    Dim strArr
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "Select C.住院号,C.当前床号,to_char(A.确诊日期,'yyyy-MM-dd') as 确诊日期,A.登记人 经办人,B.位置 入院科室,A.住院医师,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') 入院经办时间," & _
        " to_char(A.入院日期,'yyyymmdd') 入院日期  ,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') 入院时间,D.入院诊断编码,D.入院诊断名称,G.确诊诊断编码,g.确诊诊断名称 " & _
        " From 病案主页 A,部门表 B,病人信息 C, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码,'')) AS 入院诊断编码,max(DECODE(a.诊断次序,1,b.名称,'')) AS 入院诊断名称 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   D," & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,2,b.编码,'')) AS 确诊诊断编码,max(DECODE(a.诊断次序,2,b.名称,'')) AS 确诊诊断名称 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   g" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        "       and A.主页id=g.主页id(+) and a.病人id=g.病人id(+) " & _
        ""

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息", lng病人ID, lng主页ID)

    With g病人身份_成都内江
        '输入参数
        '    个人编号    String(8)   In
        '    社保卡号码  String(10)  In
        '    医院代码    String(5)   In
        '    操作员卡号码    String(10)  In
        '    统筹地区编码    String(1)   In
        '    入院日期    String(8)   In
        '    入院科别    String(10)  In
        '    入院诊治医生    String(10)  In
        '    诊断编码    String(20)  In 20051026 修改为 string(200)
        
        'Beging 20051026 陈东
        Dim vat并发症 As Variant, str其他病种 As String, i As Long
        
        gstrSQL = "Select * from 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_成都内江
        Call zlDatabase.OpenRecordset(rsRydj, gstrSQL, "取其他病种")
        str其他病种 = Nvl(rsRydj!其他病种)
        If str其他病种 <> "" Then
            If InStr(str其他病种, "|") > 0 Then
                vat并发症 = Split(str其他病种, "|")
                str其他病种 = ""
                For i = 0 To UBound(vat并发症) - 1
                    str其他病种 = str其他病种 & Rpad(Substr(vat并发症(i), 1, 20), 20)
                Next
            End If
        Else
            str其他病种 = Space(180)
        End If
        'End 20051026 陈东
        StrInput = Rpad(.个人编号, 8)
        StrInput = StrInput & vbTab & Rpad(.卡号, 10)
        StrInput = StrInput & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
        StrInput = StrInput & vbTab & Rpad(gstrUserName, 10)
        StrInput = StrInput & vbTab & Rpad(.统筹编号, 1)
        StrInput = StrInput & vbTab & Rpad(rsTemp!入院日期, 8)
        StrInput = StrInput & vbTab & Rpad(Substr(Nvl(rsTemp!入院科室), 1, 10), 10)
        StrInput = StrInput & vbTab & Rpad(Substr(Nvl(rsTemp!住院医师), 1, 10), 10)
        StrInput = StrInput & vbTab & Rpad(Rpad(Substr(g病人身份_成都内江.病种编码, 1, 20), 20) & Substr(str其他病种, 1, 180), 200)
                
        If 业务请求_成都内江(住院登记_内江, StrInput, strOutput) = False Then
            Exit Function
        End If
        
        '输出参数
        '    住院流水号  String(20)  Out
        '    享受待遇标志    Small int   Out
        '    起付标准    Long    Out
        '    在职情况    String(1)   Out
        
        strArr = Split(strOutput, vbTab)

        '保存将交易流水号
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_成都内江 & ",'顺序号','''" & strArr(0) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存交易流水号")
        '保存享受待遇标志
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_成都内江 & ",'享受待遇标志','''" & Val(strArr(1)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存享受待遇标志")
        '保存起付标准
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_成都内江 & ",'起付标准','''" & Val(strArr(2)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存起付标准")
    End With

    病人入院登记处理 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_成都内江(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
    
    '获取住院号
    Err = 0
    On Error GoTo errHand:
 
    '判断单位欠费情况
    '    个人编号    String (8)  IN
    '    社保卡号码  String (10) IN
    '    统筹地区编码    String (1)  IN
    StrInput = g病人身份_成都内江.个人编号
    StrInput = StrInput & vbTab & g病人身份_成都内江.卡号
    StrInput = StrInput & vbTab & g病人身份_成都内江.统筹编号
    
    If 业务请求_成都内江(获取单位欠缴情况_内江, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    If Val(strOutput) <> 0 Then
        ShowMsgbox "注意：" & vbCrLf & "    单位已经欠费!"
        'Exit Function
    End If
    
    '先进行登记处理
    If 病人入院登记处理(lng病人ID, lng主页ID) = False Then
        'Exit Function
    End If
    


    '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_成都内江 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_成都内江 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_成都内江 = False
End Function

Public Function 入院登记撤销_成都内江(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false

    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim str医保号  As String
    Dim str出院日期 As String

    Err = 0
    On Error GoTo errHand
    ShowMsgbox "该医保接口不支持入院登记撤消,只能办理出院"
    Exit Function
   入院登记撤销_成都内江 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 出院登记_成都内江(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    '个人状态的修改
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_成都内江 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_成都内江 = True
    Exit Function
errHand::
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_成都内江 = False
End Function
Public Function 出院登记撤销_成都内江(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '出院登记撤消
     '改变病人状态
     If Not 存在未结费用(lng病人ID, lng主页ID) Then
            ShowMsgbox "该病人已经出院结算了,不能出院登记撤消!"
            Exit Function
     End If
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_成都内江 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    出院登记撤销_成都内江 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function 住院结算_成都内江(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
  '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)

    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String

    Dim str操作员 As String
    Dim lng主页ID As Long
    Dim strArr
    Dim lng入院病种 As Long, lng出院病种 As Long
    
    Dim i As Integer

    If g病人身份_成都内江.lng病人ID <> lng病人ID Then
        MsgBox "该病人没有完成医保的预结算操作，不能进行结算。", vbInformation, gstrSysName
        Exit Function
    End If
    gstrSQL = "Select Sum(nvl(结帐金额,0)) as 总额 from 住院费用记录 where 结帐id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "费用总额"
    
    If g病人身份_成都内江.费用总额 + g病人身份_成都内江.生育总费用 <> Nvl(rsTemp!总额, 0) Then
        'Modified by ZYB 20051118 提示框显示的信息不正确
        Err.Raise 9000, gstrSysName, "虚结算费用总额不等于本次结算总额:" & vbCrLf & "虚结算总额:" & Format(g病人身份_成都内江.费用总额 + g病人身份_成都内江.生育总费用, "#####0.00;-####0.00;0;0") & vbCrLf & "当前结算总额为:" & Format(Nvl(rsTemp!总额, 0), "#####0.00;-####0.00;0;0")
        Exit Function
    End If

    Err = 0: On Error GoTo errHand:
    Call DebugTool("进入住院结算")


    With g结算数据
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & lng病人ID
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "虚拟结算")
        If IsNull(rsTemp("主页ID")) = True Then
            Err.Raise 9000, gstrSysName, "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        lng主页ID = rsTemp("主页ID")
    End With
    
'   gstrSQL = "Select A.ID From 病人费用记录 a,药品收发记录 B where A.no=b.No and B.单据 in (9,10) and a.id=b.费用ID and a.结帐ID=" & lng结帐ID & " and b.扣率 like '_3%' and rownum<=2"
'
'
'    Dim bln出院带药 As Boolean
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定出院带药"
'    If rsTemp.EOF Then
'        bln出院带药 = False
'    Else
'        bln出院带药 = True
'    End If
'
'  gstrSQL = "Select c.住院号,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 入院经办时间," & _
'        " to_char(A.入院日期,'yyyyMMdd') 入院日期,J.终止时间,J.操作员,D.诊断编码,A.出院方式,to_Char(a.出院日期,'yyyyMMDD') as 出院日期,a.出院病床,H.名称 as 出院科室" & _
'        " From 病案主页 A,部门表 B,病人信息 C,部门表 H, " & _
'        "       (Select 病人id,主页id max(DECODE(a.诊断次序,2,b.编码,'')) AS 诊断编码 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =3  and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id)   D" & _
'        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
'        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID and A.出院科室ID=H.id(+) " & _
'        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
'        ""
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定入出类别"
'
'
'
'    '入参:
'    '    个人编号    String(8)   In
'    '    社保卡号码  String(10)  In
'    '    医院代码    String(5)   In
'    '    操作员卡号码    String(10)  In
'    '    统筹地区编码    String(1)   In
'    '    出院日期    String(8)   In
'    '    出院科别    String(10)  In
'    '    出院诊治医生    String(10)  In
'    '    诊断编码    String(20)  In
'    '    出院带药    String(1)   In
'    '    出院类别    String(1)   In
'    '    住院流水号  String(20)  In
'
'    With g病人身份_成都内江
'        strInput = Rpad(.个人编号, 8)
'        strInput = strInput & vbTab & Rpad(.卡号, 10)
'        strInput = strInput & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
'        strInput = strInput & vbTab & Rpad(Nvl(rsTemp!操作员), 10)
'        strInput = strInput & vbTab & Rpad(.统筹编号, 1)
'        strInput = strInput & vbTab & Rpad(Nvl(rsTemp!出院日期), 8)
'        strInput = strInput & vbTab & Substr(Rpad(Nvl(rsTemp!出院科室), 10), 1, 10)
'        strInput = strInput & vbTab & Substr(Rpad(Nvl(rsTemp!住院医师), 10), 1, 10)
'        strInput = strInput & vbTab & Substr(Rpad(Nvl(rsTemp!诊断编码), 20), 1, 20)
'        strInput = strInput & vbTab & IIf(bln出院带药, "1", "0")
'        strInput = strInput & vbTab & Substr(Nvl(rsTemp!出院方式), 1, 1)
'        strInput = strInput & vbTab & Substr(Rpad(.住院流水号, 20), 1, 20)
'        If 业务请求_成都内江(出院登记上传_内江, strInput, stroutput) = False Then Exit Function
'    End With
'    If stroutput = "" Then Exit Function
'    strArr = Split(stroutput, vbTab)
'
'   '出参
'    '    TRANSDETIAL输出 (计算费用明细)
'    '    享受待遇标志    String(1)   Out
'    '    医保内费用  String(10)  Out
'    '    医保外费用  String(10)  Out
'    '    基本医保支付
'    '    如果参加大病医保，则为大病医保支付  String(10)  Out
'    '    高额医保支付    String(10)  Out
'    '    公务员医疗补助  String(10)  Out
'    '    个人按比例支付  String(10)  Out
'    '    TRANSDETIAL结束
'    '    起付标准    String(10)  Out
'    '    个人帐户可用余额    String(10)  Out
'    stroutput = strArr(0)
'    With g结算数据
'        .结算标志 = 1
'        .待遇标志 = Substr(stroutput, 1, 1)
'        .医保内费用 = Val(Substr(stroutput, 2, 10))
'        .医保外费用 = Val(Substr(stroutput, 12, 10))
'        .基本医保支付 = Val(Substr(stroutput, 22, 10))
'        .高额医保支付 = Val(Substr(stroutput, 32, 10))
'        .公务员医疗补助 = Val(Substr(stroutput, 42, 10))
'        .比例支付 = Val(Substr(stroutput, 52, 10))
'        .起付标准 = Val(strArr(1))
'        .帐户可用余额 = Val(strArr(2))
'
'    End With
'
'    If 结算方式更正(2) = False Then Exit Function
        
     '获取帐户支付
    g结算数据.结帐ID = lng结帐ID
    g结算数据.帐户支付 = 获取个人帐户支付() * 100
    '    个人编号    String(8)   In
    '    社保卡号码  String(10)  In
    '    操作员卡号码    String(10)  In
    '    统筹地区编号    String(1)   In
    '    住院流水号  String(20)  In
    '    个人帐户支付    String(10)  In

    With g病人身份_成都内江
        StrInput = Rpad(.个人编号, 8)
        StrInput = StrInput & vbTab & Rpad(.卡号, 10)
        StrInput = StrInput & vbTab & Rpad(Substr(gstrUserName, 1, 10), 10)
        StrInput = StrInput & vbTab & Rpad(.统筹编号, 1)
        StrInput = StrInput & vbTab & Rpad(Substr(Rpad(.住院流水号, 20), 1, 20), 20)
        StrInput = StrInput & vbTab & Lpad(Substr(Lpad(g结算数据.帐户支付, 10), 1, 10), 10, 0)
    End With
    If 业务请求_成都内江(出院登记确认_内江, StrInput, strOutput) = False Then
        Exit Function
    End If


  '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(无),帐户累计支出_IN(无),累计进入统筹_IN(无),累计统筹报销_IN(生育支付),住院次数_IN(无),起付线(比例支付),封顶线_IN(帐户可用余额),实际起付线_IN(起付标准),
    '   发生费用金额_IN(费用总额),全自付金额_IN(医保内费用),首先自付金额_IN(医保外费用),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(高额医保支付),超限自付金额_IN(公务员医疗补助),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN(享受待遇标志)
    
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    'Bgeing 陈东 20050601
    gstrSQL = "Select * from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & TYPE_成都内江
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "入出院病种")
    lng入院病种 = Nvl(rsTemp!病种ID, 0)
    lng出院病种 = Nvl(rsTemp!出院病种ID, 0)
    
    'beging 20051026 陈东
    Dim str入院其他病种 As String, str出院其他病种 As String
    str入院其他病种 = Nvl(rsTemp!其他病种)
    str出院其他病种 = Nvl(rsTemp!出院其他病种)
    'end '20051026 陈东
    
    'End  陈东 20050601
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_成都内江 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "0,0,0," & g结算数据.生育支付 & ",0," & g结算数据.比例支付 & "," & g结算数据.帐户可用余额 & "," & g结算数据.起付标准 & "," & _
            g病人身份_成都内江.费用总额 & "," & g结算数据.医保内费用 & "," & g结算数据.医保外费用 & "," & _
           g结算数据.基本医保支付 & "," & g结算数据.基本医保支付 & "," & g结算数据.高额医保支付 & "," & g结算数据.公务员医疗补助 & "," & g结算数据.帐户支付 / 100 & ",'" & _
            g病人身份_成都内江.住院流水号 & "',NULL,NULL,'" & g结算数据.待遇标志 & "')"
            

    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
    'beging 陈东 20050601
    gstrSQL = "update 保险结算记录 Set 病种ID=" & lng入院病种 & ",出院病种ID=" & lng出院病种 & _
             " where 记录ID=" & lng结帐ID & " And 性质=2 and 险类=" & TYPE_成都内江
             
    gcnOracle.Execute gstrSQL
    'end
    'beging 20051026 陈东
    If str入院其他病种 <> "" Then
        gstrSQL = "Update 保险结算记录 Set 其他病种='" & str入院其他病种 & "'" & _
                " Where 记录ID=" & lng结帐ID & " And 性质=2 and 险类=" & TYPE_成都内江
        gcnOracle.Execute gstrSQL
    End If
    If str出院其他病种 <> "" Then
        gstrSQL = "Update 保险结算记录 Set 出院其他病种='" & str出院其他病种 & "'" & _
                " Where 记录ID=" & lng结帐ID & " And 性质=2 and 险类=" & TYPE_成都内江
        gcnOracle.Execute gstrSQL
    End If
    'end '20051026 陈东
    住院结算_成都内江 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function 住院结算冲销_成都内江(lng结帐ID As Long) As Boolean
     '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------

    Err = 0: On Error GoTo errHand:
    Err.Raise 9000, gstrSysName, "本医保接不支持反结帐"
    住院结算冲销_成都内江 = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function 保存明细结果到中间库(ByVal str交易流水号 As String, ByVal strOutput As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存明细结果到中间库
    '--入参数:以vbtab分离
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim StrInput As String, str明细 As String
    Dim strArr
    Dim lngSumLen As Long
    Dim r As Long, i As Integer
    'strOutPut: 医保交易流水号  String(20)  Out
    '           处方明细    String处方条数×51  Out
    '           TRANSDETIAL输出 (计算费用明细) Out
    
    'TRANSDETIAL输出
    '        享受待遇标志    String(1)   Out
    '        医保内费用  String(10)  Out
    '        医保外费用  String(10)  Out
    '        基本医保支付 如果参加大病医保，则为大病医保支付  String(10)  Out
    '        高额医保支付    String(10)  Out
    '        公务员医疗补助  String(10)  Out
    '        个人按比例支付  String(10)  Out


    保存明细结果到中间库 = False
    Err = 0: On Error GoTo errHand:
    strArr = Split(strOutput, vbTab)
    
    '过程参数
    '    医院流水号_IN IN 医保消费信息.医院流水号%TYPE,
    '    病人ID_IN IN 医保消费信息.病人ID%TYPE,
    '    医保流水号_IN IN 医保消费信息.医保流水号%TYPE,
    '    医保内费用_IN IN 医保消费信息.医保内费用%TYPE,
    '    医保外费用_IN IN 医保消费信息.医保外费用%TYPE,
    '    帐户可用余额_IN IN 医保消费信息.帐户可用余额%TYPE:=NULL,
    '    在职情况_IN IN 医保消费信息.在职情况%TYPE:=NULL,
    '    医保项目种类_IN IN 医保消费信息.医保项目种类%TYPE:=NULL,
    '    医保项目编码_IN IN 医保消费信息.医保项目编码%TYPE,
    '    医保内费用1_IN IN 医保消费信息.医保内费用1%TYPE,
    '    费用类别_IN IN 医保消费信息.费用类别%TYPE:=NULL,
    '    项目费用_IN IN 医保消费信息.项目费用%TYPE:=NULL,
    '    享受待遇标志_IN IN 医保消费信息.享受待遇标志%TYPE:=NULL,
    '    基本医保支付_IN IN 医保消费信息.基本医保支付%TYPE:=NULL,
    '    高额医保支付_IN IN 医保消费信息.高额医保支付%TYPE:=NULL,
    '    公务员补助_IN IN 医保消费信息.公务员补助%TYPE:=NULL,
    '    个人比例支付_IN IN 医保消费信息.个人比例支付%TYPE:=NULL
    '    20051021 Add
    '    生育支付_IN IN 医保消费信息.生育支付%TYPE:NULL
    strHead = "ZL_医保消费信息_INSERT("
    strHead = strHead & "'" & str交易流水号 & "',"
    strHead = strHead & "" & g病人身份_成都内江.lng病人ID & ","
    strHead = strHead & "'" & strArr(0) & "',"
    
    strHead = strHead & "" & Val(Substr(strArr(2), 2, 10)) & ","
    strHead = strHead & "" & Val(Substr(strArr(2), 12, 10)) & ","
    strHead = strHead & "" & 0 & ","
    strHead = strHead & "null,"
    
    str明细 = strArr(1)
    lngSumLen = zlCommFun.ActualLen(strArr(1))
    StrInput = ""
    For i = 1 To lngSumLen Step 51
     '    在职情况_IN IN 医保消费信息.在职情况%TYPE:=NULL,
    '    医保项目种类_IN IN 医保消费信息.医保项目种类%TYPE:=NULL,
    '    医保项目编码_IN IN 医保消费信息.医保项目编码%TYPE,
    '    医保内费用1_IN IN 医保消费信息.医保内费用1%TYPE,
    '    费用类别_IN IN 医保消费信息.费用类别%TYPE:=NULL,
    '    项目费用_IN IN 医保消费信息.项目费用%TYPE:=NULL,
    
        r = i
        StrInput = StrInput & "'" & Substr(str明细, r, 1) & "',"
        r = r + 1
        StrInput = StrInput & "'" & Substr(str明细, r, 20) & "',"
        r = r + 20
        StrInput = StrInput & "" & Val(Substr(str明细, r, 10)) & ","
        r = r + 10
        StrInput = StrInput & "'" & Substr(str明细, r, 10) & "',"
        r = r + 10
        StrInput = StrInput & "" & Val(Substr(str明细, r, 10)) & ","
        
        
        '加上
        'TRANSDETIAL输出
         '        享受待遇标志    String(1)   Out
         '        医保内费用  String(10)  Out
         '        医保外费用  String(10)  Out
         '        基本医保支付 如果参加大病医保，则为大病医保支付  String(10)  Out
         '        高额医保支付    String(10)  Out
         '        公务员医疗补助  String(10)  Out
         '        个人按比例支付  String(10)  Out

 '    享受待遇标志_IN IN 医保消费信息.享受待遇标志%TYPE:=NULL,
    '    基本医保支付_IN IN 医保消费信息.基本医保支付%TYPE:=NULL,
    '    高额医保支付_IN IN 医保消费信息.高额医保支付%TYPE:=NULL,
    '    公务员补助_IN IN 医保消费信息.公务员补助%TYPE:=NULL,
    '    个人比例支付_IN IN 医保消费信息.个人比例支付%TYPE:=NULL
 
        StrInput = StrInput & "'" & Substr(strArr(2), 1, 1) & "',"
        StrInput = StrInput & "" & Val(Substr(strArr(2), 22, 10)) & ","
        StrInput = StrInput & "" & Val(Substr(strArr(2), 32, 10)) & ","
        StrInput = StrInput & "" & Val(Substr(strArr(2), 42, 10)) & ","
         '20051021 add
        StrInput = StrInput & "" & Val(Substr(strArr(2), 52, 10)) & ","
        StrInput = StrInput & "" & Val(Substr(strArr(2), 62, 10)) & ")"
        '组合SQL误句
        
        gstrSQL = strHead & StrInput
        ExecuteProcedure_ZLNJ "插入明细数据到中间库"
        StrInput = ""
    Next
    保存明细结果到中间库 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function StartOrCommitorRollbackTransaction(ByVal bytType As Byte, Optional blnGcnoracle As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:启动、提交、回滚事务
    '--入参数:byttype-0启动,1提交,2回滚
    '         blnGcnoracle-是否存在事务(gcnoracle)
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Select Case bytType
        Case 0
            gcnOracle_成都内江.BeginTrans
            If Not blnGcnoracle Then
                gcnOracle.BeginTrans
            End If
            mblnStartTran = True
        Case 1
            gcnOracle_成都内江.CommitTrans
            If Not blnGcnoracle Then
                gcnOracle.CommitTrans
            End If
            mblnStartTran = False
        Case Else
            gcnOracle_成都内江.RollbackTrans
            If Not blnGcnoracle Then
                gcnOracle.RollbackTrans
            End If
            mblnStartTran = False
        End Select
End Function

Private Function 处方上传(ByVal lng记录性质 As Long, lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '处方明细上传
    '功能:上传新产生的记帐明细到医保中心
    '参数:  str单据号   NO
    '       int性质     记录性质
    '       lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
    '返回:
    Dim rsTemp As New ADODB.Recordset, rs明细 As New ADODB.Recordset
    Dim StrInput As String, strOutput As String, strArr As Variant
    Dim lng病人ID As Long, str明细 As String
    Dim i As Long
    Dim 取交易号_int As Integer
    
    'Beging 20051025 add
    Dim bln传明细 As Boolean
    Dim vat并发症 As Variant, str附加诊断   As String
    Dim lng医院编码长度 As Long, str入院日期 As String
    'End 20051025 add
    
    处方上传 = False
    
    Err = 0: On Error GoTo errHand:
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.主页ID,to_char(A.发生时间,'yyyymmdd') as 发生时间,to_char(A.发生时间,'yyyy-mm-dd hh24:mi:ss') as 登记时间,Round(A.实收金额,4) 实收金额 " & _
              "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,Z.位置 as 开单部门,A.保险编码,C.项目编码,C.大类编码,J.类别 as 收费类别,C.是否医保,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型,M.医保号 " & _
              "  From 住院费用记录 A,部门表 Z,收费类别 J,收费细目 B,保险帐户 M,(Select O.*,Z.大类编码 From 保险支付项目 O,保险项目 Z where O.险类=Z.险类 and O.项目编码=Z.编码 and O.险类=" & TYPE_成都内江 & ") C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
              "  where a.病人id=M.病人id and a.开单部门ID=Z.iD(+)   and M.险类=" & TYPE_成都内江 & " and A.NO='" & str单据号 & "' and A.记录性质=" & lng记录性质 & " and A.记录状态=" & lng记录状态 & "And Nvl(A.是否上传,0)=0 " & _
              "        and A.收费类别=J.编码(+)  and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=" & TYPE_成都内江 & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID(+) " & _
              "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
              "  Order by A.病人ID,A.登记时间,C.项目编码"

    Call zlDatabase.OpenRecordset(rs明细, gstrSQL, "处方明细上传")
    Dim lng冲销ID As Long

    '先检查是否存在退单情况，如果存在，看有否对应的记录单.
    With rs明细
        '上传明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '单价检查
        
            If Val(!数量) < 0 Or Val(!价格) < 0 Then
                ShowMsgbox "在单据中不能输入负单据!"
                Exit Function
            End If
            If Nvl(!项目编码) = "" Then
                 MsgBox "有项目未设置医保编码[" & Nvl(!编码) & "-" & Nvl(!名称) & "]，不能上传明细!", vbInformation, gstrSysName
                 Exit Function
            End If
            
            .MoveNext
        Loop
        
    End With
    
    If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
    
    Dim str交易流水号 As String
    Dim blnStarTran As Boolean '启动事务
    Dim str项目编码 As String, str保险编码 As String, str大类编码 As String
    
    StrInput = ""
    mblnStartTran = False
    取交易号_int = 0
    lng病人ID = 0
    str项目编码 = "@#$%^&(*)_+_)(*&^%$$#"
    '进行费用传输
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do Until .EOF
                If mblnStartTran = False Then
                    '启动事务
                    Call StartOrCommitorRollbackTransaction(0)
                End If
                
                str保险编码 = Nvl(!保险编码)
                '如果保险编码为空，则需要用户选择编码
                If str保险编码 = "" Then
                    str保险编码 = GetItemInsure_成都内江(0, !收费细目ID, False)
                End If
                '如果为空表示没有取到缺省编码（使用新程序对码以前的项目存在这种情况），所以取当前记录集中的项目编码即可
                If str保险编码 = "" Then str保险编码 = Nvl(!项目编码)
                
                'Begin 20051025
                '普通医保病人(入院诊断不是“平产”和“剖宫产”)
                'a.不能用“生育费用”(见生育项目代码.xls后4项)，可以用“计划生育费用”(见生育项目代        码.xls前15项)和普通费用混合打包上传;
                bln传明细 = True
                gstrSQL = "Select * From 保险病种 Where id=(select 病种ID from 保险帐户 Where 病人ID=" & !病人ID & ")" & _
                          " And (名称 ='平产' or 名称='剖宫产')"
                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "取入院诊断")
                If rsTemp.EOF = True Then
                    If str保险编码 = "M10000116" Or str保险编码 = "M10000117" Or str保险编码 = "M10000118" Or str保险编码 = "M10000119" Then
                        ShowMsgbox "普通医保疾病不能使用生育费用：" & rs明细!名称 & vbCrLf & _
                                   "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                        bln传明细 = False
                        'Call StartOrCommitorRollbackTransaction(2)
                        'Exit Function
                    End If
                Else
                    gstrSQL = "Select to_char(入院日期,'yyyymmdd') as 入院日期 from 病案主页 Where 病人ID=" & !病人ID & " And 主页ID=" & !主页ID
                    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "附加诊断")
                    str入院日期 = rsTemp!入院日期
                    gstrSQL = "Select * from 保险类别 where 序号=" & TYPE_成都内江
                    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "医院编号")
                    lng医院编码长度 = Len(Trim(rsTemp!医院编码))
                    
                    gstrSQL = "Select * from 保险帐户 Where 病人ID=" & !病人ID & " And 险类=" & TYPE_成都内江
                    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "附加诊断")
                    If Nvl(rsTemp!附加诊断) = "" Then
                        ' Beging 一般生育病人(入院诊断是“平产”和“剖宫产”)
                        '       a.只能上传“生育费用”(见生育项目代码.xls后4项)和“计划生育费用”(见生育项目代        码.xls前15项),可以使用一般费用但不上传；
                        '       b.4项“生育费用”中(除生育多胞胎-M10000118)每项不能重复使用，且"平产"费和"剖宫产"费只用一个；
                        'If str保险编码 = "M10000116" Or str保险编码 = "M10000117" Or str保险编码 = "M10000118" Or str保险编码 = "M10000119" Then
                        '龚智毅 20051026 可以用计划生育项目
                        If Substr(str保险编码, 1, 1) = "M" Then
                            gstrSQL = "Select 保险编码 from 住院费用记录 Where 记录状态=1 and 摘要<>'不上传'" & _
                                      " and 保险编码 IN ('M10000116','M10000117','M10000119') " & _
                                      " And 主页id=" & !主页ID & " and 病人id=" & !病人ID
                            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "是否有重复项目")
                            Do Until rsTemp.EOF
                                If str保险编码 = rsTemp!保险编码 Then
                                    ShowMsgbox "医保规定此项目：" & rs明细!名称 & " [" & str保险编码 & "] " & "不能重复使用" & vbCrLf & _
                                               "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                                    bln传明细 = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                                rsTemp.MoveNext
                            Loop
                            
                            '龚智毅 20051026  优化语句
                            If str保险编码 = "M10000116" Or str保险编码 = "M10000117" Then
                                gstrSQL = "Select 保险编码 from 住院费用记录 Where 记录状态=1 and 摘要<>'不上传'" & _
                                          " And (保险编码='M10000116' or 保险编码='M10000117')" & _
                                          " And 主页id=" & !主页ID & " and 病人id=" & !病人ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "是否有互斥项目")
                                If rsTemp.RecordCount > 0 Then
                                    ShowMsgbox "医保规定M10000116(平产)不能和M10000117(剖宫产)一同使用" & vbCrLf & _
                                               "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                                    bln传明细 = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            If str保险编码 = "M10000118" Then
                                gstrSQL = "Select 保险编码 from 住院费用记录 Where 记录状态=1 and 摘要<>'不上传'" & _
                                          " And (保险编码='M10000116' or 保险编码='M10000117')" & _
                                          " And 主页id=" & !主页ID & " and 病人id=" & !病人ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "是否有基础项目")
                                If rsTemp.EOF Then
                                    ShowMsgbox "医保规定使用M10000118(生育多胞胎),需要先使用M10000116(平产)或M10000117(剖宫产)" & vbCrLf & _
                                               "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                                    bln传明细 = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            If str保险编码 = "M10000119" Then
                                gstrSQL = "Select 保险编码 from 住院费用记录 Where 记录状态=1 and 摘要<>'不上传'" & _
                                          " And 保险编码='M10000117'" & _
                                          " And 主页id=" & !主页ID & " and 病人id=" & !病人ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "是否有剖宫产")
                                If rsTemp.EOF Then
                                    ShowMsgbox "医保规定使用M10000119(生育全身麻醉),需要先使用M10000117(剖宫产)" & vbCrLf & _
                                               "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                                    bln传明细 = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            'bln传明细 = True
                        Else
                            bln传明细 = False
                        End If
                        'End 一般生育病人
                    Else
                        'Beging 并发症生育病人(入院诊断是“平产”和“剖宫产”，且完成了并发症申请)
'                            a.可以使用所有费用，使用的“生育费用”(见生育项目代码.xls后4项)和“计划生育费用”(见生育项目代码.xls前15项)需要上传；
'
'                            c.4项“生育费用”中每项不能重复使用，且"平产"费和"剖宫产"费只用一个;
                        'If str保险编码 = "M10000116" Or str保险编码 = "M10000117" Or str保险编码 = "M10000118" Or str保险编码 = "M10000119" Then
                        '龚智毅 20051026 可以用计划生育项目
                        If Substr(str保险编码, 1, 1) = "M" Then
                            gstrSQL = "Select 保险编码 from 住院费用记录 Where 记录状态=1 and 摘要<>'不上传'" & _
                                      " And 保险编码 IN ('M10000116','M10000117','M10000119')" & _
                                      " And 主页id=" & !主页ID & " and 病人id=" & !病人ID
                            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "是否有重复项目")
                            Do Until rsTemp.EOF
                                If str保险编码 = rsTemp!保险编码 Then
                                    ShowMsgbox "医保规定此项目：" & rs明细!名称 & "[" & str保险编码 & "]" & "不能重复使用" & vbCrLf & _
                                               "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                                    bln传明细 = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                                rsTemp.MoveNext
                            Loop
                            
                            '龚智毅 20051026  优化语句
                            If str保险编码 = "M10000116" Or str保险编码 = "M10000117" Then
                                gstrSQL = "Select 保险编码 from 住院费用记录 Where 记录状态=1 and 摘要<>'不上传'" & _
                                          " And (保险编码='M10000116' or 保险编码='M10000117')" & _
                                          " And 主页id=" & !主页ID & " and 病人id=" & !病人ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "是否有互斥项目")
                                If rsTemp.RecordCount > 0 Then
                                    ShowMsgbox "医保规定M10000116(平产)不能和M10000117(剖宫产)一同使用" & vbCrLf & _
                                               "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                                    bln传明细 = False
                                   ' Call StartOrCommitorRollbackTransaction(2)
                                   ' Exit Function
                                End If
                            End If
                            
                            If str保险编码 = "M10000118" Then
                                gstrSQL = "Select 保险编码 from 住院费用记录 Where 记录状态=1 and 摘要<>'不上传'" & _
                                          " And (保险编码='M10000116' or 保险编码='M10000117')" & _
                                          " And 主页id=" & !主页ID & " and 病人id=" & !病人ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "是否有基础项目")
                                If rsTemp.EOF Then
                                    ShowMsgbox "医保规定使用M10000118(生育多胞胎),需要先使用M10000116(平产)或M10000117(剖宫产)" & vbCrLf & _
                                               "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                                    bln传明细 = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            If str保险编码 = "M10000119" Then
                                gstrSQL = "Select 保险编码 from 住院费用记录 Where 记录状态=1 and 摘要<>'不上传'" & _
                                          " And 保险编码='M10000117'" & _
                                          " And 主页id=" & !主页ID & " and 病人id=" & !病人ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "是否有剖宫产")
                                If rsTemp.EOF Then
                                    ShowMsgbox "医保规定使用M10000119(生育全身麻醉),需要先使用M10000117(剖宫产)" & vbCrLf & _
                                               "收费项目：" & rs明细!名称 & " [" & str保险编码 & "] 不上传"
                                     bln传明细 = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            'bln传明细 = True
                        Else
                            'b.使用的普通项目如果是因并发症新需要的费用项目则上传,否则不上传;
                          '龚智毅 20051026 手工选择是否上传
                            'str附加诊断 = Nvl(rsTemp!附加诊断)
                            'If InStr(str附加诊断, "|") > 0 Then
                            '    vat并发症 = Split(str附加诊断, "|")
                            '    For i = 0 To UBound(vat并发症) - 1
                            '        gstrSQL = "Select * From 保险特准项目 A,保险病种 B " & _
                            '                "Where A.病种ID=B.Id And B.编码='" & Split(vat并发症(i), ";")(0) & "'" & _
                            '                " And A.收费细目ID=" & !收费细目ID
                            '        Call OpenRecordset(rsTemp, "是否是特准项目")
                            '        If rsTemp.EOF Then
                            '            bln传明细 = False
                            '        Else
                            '            bln传明细 = True
                            '        End If
                            '    Next
                            'End If
                            If MsgBox("收费项目：" & rs明细!名称 & "    金额：" & rs明细!实收金额 & vbCrLf & _
                                     "请确认是否并发症使用项目。" & vbCrLf & _
                                     "（点 [是] 则上传到医保，点 [否] 则不上传）", vbYesNo, gstrSysName) = vbYes Then
                                bln传明细 = True
                            Else
                                bln传明细 = False
                            End If
                        End If
                        'End 并发症生育病人
                    End If

                End If
                
                'End 20051025
                
                '取大类编码
                gstrSQL = "Select 大类编码 From 保险项目 Where 险类=" & TYPE_成都内江 & " And 编码='" & str保险编码 & "'"
                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "取大类编码")
                str大类编码 = Nvl(rsTemp!大类编码)
                
               '龚智毅 20051027
            If bln传明细 Then
               If lng病人ID <> Nvl(!病人ID, 0) Or i >= 38 Or str项目编码 = str保险编码 Then
                    'i = 1
                    If 取交易号_int = 0 Then
                        str交易流水号 = Get交易流水号(Nvl(!发生时间))
                        取交易号_int = 1
                    End If
                    lng病人ID = Nvl(!病人ID, 0)
                    str项目编码 = str保险编码
                    If StrInput <> "" Then
                        StrInput = StrInput & vbTab & Rpad(i - 1, 2) & vbTab & str明细
                        '请求相关的业务数据
                        'Beging 20051025 add
                        'If bln传明细 = True Then
                            If 业务请求_成都内江(住院交易上传_内江, StrInput, strOutput) = False Then
                                '回滚事务
                                Call StartOrCommitorRollbackTransaction(2)
                                Exit Function
                            End If
                            If 保存明细结果到中间库(str交易流水号, strOutput) = False Then
                                '提交事务,中间库数据未保存起
                                Call StartOrCommitorRollbackTransaction(1)
                                Exit Function
                            End If
                        'End If
                        'End 20051025 add
                        取交易号_int = 0
                    End If
                    i = 1
                    If 取交易号_int = 0 Then
                        str交易流水号 = Get交易流水号(Nvl(!发生时间))
                        取交易号_int = 1
                    End If
                    
                    Call Get病人信息(lng病人ID)
                    StrInput = Rpad(g病人身份_成都内江.个人编号, 8)
                    
                    StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.卡号, 10)
                    StrInput = StrInput & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
                    StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.统筹编号, 1)
                    StrInput = StrInput & vbTab & Rpad(str交易流水号, 20)
                    If Nvl(str大类编码) = "1" Then
                        StrInput = StrInput & vbTab & IIf(IS出院带药(Nvl(!NO), Nvl(!ID, 0)), "1", "0")
                    Else
                        StrInput = StrInput & vbTab & "0"
                    End If
                    StrInput = StrInput & vbTab & Rpad(Substr(Nvl(!开单部门), 1, 10), 10)
                    StrInput = StrInput & vbTab & Rpad(Substr(!医生, 1, 10), 10)
                    StrInput = StrInput & vbTab & Rpad(Substr(g病人身份_成都内江.住院流水号, 1, 20), 20)
                    str明细 = ""
                    '个人编号    String(8)   In
                    '社保卡号码  String(10)  In
                    '医院代码    String(5)   In
                    '统筹地区编码    String(1)   In
                    '医院交易流水号  String(20)  In
                    '出院带药类别    String(1)   In
                    
                    '科别    String(10)  In
                    '医生    String(10)  In
                    '住院流水号  String(20)  In
                    '处方条数    String(2)   In
                    '处方明细    String处方条数×51  In
               End If
                
                lng病人ID = Nvl(!病人ID, 0)
                str项目编码 = str保险编码
                
                str明细 = str明细 & Substr(Rpad(Nvl(str大类编码), 1), 1, 1)
                str明细 = str明细 & Rpad(str保险编码, 20)
                str明细 = str明细 & Lpad(Nvl(!数量) * 100, 10, "0")
                str明细 = str明细 & Rpad(Nvl(!规格), 10)
                str明细 = str明细 & Lpad(Nvl(!实收金额) * 100, 10, "0")
            
                i = i + 1
            End If
            
            '龚智毅 20051027
            If bln传明细 Then
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL," & Nvl(str大类编码, "NULL") & ",NULL,'" & str保险编码 & "',1,'" & str交易流水号 & "')"
            Else
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL," & Nvl(str大类编码, "NULL") & ",NULL,'" & str保险编码 & "',1,'不上传')"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            .MoveNext
        Loop
    End With
    
    'Beging 20051025 add
    'If bln传明细 = True Then
        If StrInput <> "" Then
            StrInput = StrInput & vbTab & Rpad(i - 1, 2) & vbTab & str明细
            '请求相关的业务数据
            If 业务请求_成都内江(住院交易上传_内江, StrInput, strOutput) = False Then
                '提交事务,中间库数据未保存起
                Call StartOrCommitorRollbackTransaction(2)
                Exit Function
            End If
            If 保存明细结果到中间库(str交易流水号, strOutput) = False Then
                '提交事务,中间库数据未保存起
                Call StartOrCommitorRollbackTransaction(2)
                Exit Function
            End If
            '提交
            StartOrCommitorRollbackTransaction (1)
        Else
            '龚智毅 20051027
            If bln传明细 Then
              If mblnStartTran Then
                 Call StartOrCommitorRollbackTransaction(2)
              End If
            Else
              StartOrCommitorRollbackTransaction (1)
            End If
        End If
    'End If
    'End 20051025 add

    处方上传 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    If mblnStartTran Then
        '提交事务,中间库数据未保存起
        Call StartOrCommitorRollbackTransaction(2)
    End If
End Function
Private Function IS出院带药(ByVal strNO As String, lng费用ID As Long) As Boolean
    '检查是否出院带药
    'Dim rsTemp As New ADODB.Recordset
    Dim rsTemp1 As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errHand:
    Set rsTemp1 = New ADODB.Recordset
    gstrSQL = "Select ID From 药品收发记录 where NO='" & strNO & "' and 单据 IN(9,10) and 费用id=" & lng费用ID & " and 扣率 like '_3%'"
    zlDatabase.OpenRecordset rsTemp1, gstrSQL, "获取是否出院带药"
    If rsTemp1.EOF Then
        IS出院带药 = False
        Exit Function
    End If
    IS出院带药 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 处方登记_成都内江(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0
    On Error GoTo errHand:


    处方登记_成都内江 = False
    
    If lng记录状态 = 1 Then
        '正常单据
        If 处方上传(lng记录性质, lng记录状态, str单据号) = False Then
            Exit Function
        End If
    Else
        '开始事务
        Call StartOrCommitorRollbackTransaction(0)
        '冲销单据
        If 处方作废(lng记录性质, lng记录状态, str单据号) = False Then
            '提交事务,中间库数据未保存起,所以回滚
            Call StartOrCommitorRollbackTransaction(2)
            Exit Function
        End If
        '提交事务
        Call StartOrCommitorRollbackTransaction(1)
    End If
    处方登记_成都内江 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 处方作废(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:记帐处方作废,即记录状态=2的记录
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim rs明细 As New ADODB.Recordset
    Dim rs原明细 As New ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String, str交易流水号 As String
    Dim strArr
    Dim lng病人ID As Long
    
    处方作废 = False

    Err = 0: On Error GoTo errHand:

    
    gstrSQL = " Select a.摘要,A.ID,a.收费细目id,A.序号,A.数次*nvl(A.付数,1) as 数量,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4) as 单价,a.病人id " & _
              " From 住院费用记录 A,保险帐户 B " & _
              " where a.病人id=b.病人id and A.NO='" & str单据号 & "' and A.记录性质=" & lng记录性质 & " and A.记录状态=3 and   Nvl(附加标志,0)<>9  " & _
              " order by A.病人id,A.摘要"
              
    Call zlDatabase.OpenRecordset(rs原明细, gstrSQL, "处方明细上传")
    
    
    If rs原明细.EOF Then
        ShowMsgbox "该单据没有相应的明细记录,不能作废!"
        Exit Function
    End If

    gstrSQL = " Select a.* " & _
              " From 住院费用记录 A,保险帐户 b" & _
              " where a.病人id=b.病人id and A.NO='" & str单据号 & "' and A.记录性质=" & lng记录性质 & " and A.记录状态=2 and  Nvl(附加标志,0)<>9 AND nvl(a.是否上传,0)=0 " & _
              " order by A.病人ID"
              
    Call zlDatabase.OpenRecordset(rs明细, gstrSQL, "处方明细上传")

    lng病人ID = 0
    '更新原单据的值
    With rs明细
        Do While Not .EOF
            rs原明细.Filter = "序号=" & Nvl(!序号, 0) & "  and 收费细目id=" & Nvl(!收费细目ID, 0)
            If rs原明细.EOF Then
                ShowMsgbox "冲销时未找到相应的记录,冲销失败!"
                Exit Function
            End If
            str交易流水号 = Nvl(rs原明细!摘要)
            If str交易流水号 = "" Then
                ShowMsgbox "在原单中不存在交易流水号,不能继续！"
                Exit Function
            End If
            '检查消费明细中有交易流水号没有
            '龚智毅 20051027
            If str交易流水号 <> "不上传" Then
            gstrSQL = "Select 医保流水号 From 医保消费信息 where 医院流水号='" & str交易流水号 & "' and 病人id=" & Nvl(!病人ID, 0)
            OpenRecordset_成都内江 rsTemp, "获取医保消费信息"
            If rsTemp.EOF Then
                ShowMsgbox "不存在医保交易数据,请与系统管理员联系!"
                Exit Function
            End If
            If Nvl(rsTemp!医保流水号) = "" Then
                ShowMsgbox "不存在医保交易数据,请与系统管理员联系!"
                Exit Function
            End If
            End If
            '更新上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rs原明细!摘要) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            .MoveNext
        Loop
    End With
    Dim str摘要 As String
    
    gstrSQL = " Select a.摘要,A.ID,a.收费细目id,A.序号,A.数次*nvl(A.付数,1) as 数量,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4) as 单价,a.病人id " & _
              " From 住院费用记录 A,保险帐户 B " & _
              " where a.病人id=b.病人id and A.NO='" & str单据号 & "' and A.记录性质=" & lng记录性质 & " and A.记录状态=3 and   Nvl(附加标志,0)<>9  " & _
              " order by A.病人id,A.摘要"
              
    Call zlDatabase.OpenRecordset(rs原明细, gstrSQL, "处方明细上传")
    
    lng病人ID = 0
    str摘要 = ""
    With rs原明细
        .MoveFirst
        Do While Not .EOF
                'lng病人ID <> Nvl(!病人ID, 0) And
                '龚智毅 20051027
            If Nvl(!摘要) <> "不上传" Then
                If str摘要 <> Nvl(!摘要) Then
                    If lng病人ID <> Nvl(!病人ID, 0) Then
                        lng病人ID = Nvl(!病人ID, 0)
                        '需重新获取相关的病人信息
                        If Get病人信息(lng病人ID) = False Then
                            ShowMsgbox "在获取病人信息时间出了错误,请与系统员立即联系!"
                            Exit Function
                        End If
                    End If
                    str摘要 = Nvl(!摘要)
                    gstrSQL = "Select 医保流水号 From 医保消费信息 where 医院流水号='" & str摘要 & "' and 病人id=" & lng病人ID
                    OpenRecordset_成都内江 rsTemp, "获取医保消费信息"
                    
                    StrInput = Rpad(g病人身份_成都内江.个人编号, 8)
                    StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.卡号, 10)
                    StrInput = StrInput & vbTab & Rpad(Substr(gstrUserName, 1, 10), 10)
                    StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.统筹编号, 1)
                    StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.住院流水号, 20)
                    StrInput = StrInput & vbTab & Rpad(Nvl(rsTemp!医保流水号), 20)
                    
                    '取消
                    '    个人编号    String(8)   In
                    '    社保卡号码  String(10)  In
                    '    操作员卡号码    String(10)  In
                    '    统筹地区编码    String(1)   In
                    '    住院流水号  String(20)  In
                    '    医保交易流水号  String(20)  In
                    
                    If 业务请求_成都内江(住院产易上传取消_内江, StrInput, strOutput) = False Then Exit Function
                End If
            End If
            .MoveNext
        Loop
    End With
    处方作废 = True
    Exit Function
errHand:
   If ErrCenter = 1 Then
        Resume
   End If
End Function
Private Function Read模拟数据(ByVal int业务类型 As 业务类型_成都内江, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以便测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim STRNAME As String
    
    If int业务类型 = 读取卡内数据 Then
        strFile = App.Path & "\解析卡.txt"
    Else
        strFile = App.Path & "\模拟提交串.txt"
    End If
    
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    STRNAME = Get交易代码(int业务类型, True)
    
    Dim blnStart As Boolean
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If int业务类型 = 读取卡内数据 Then
                    strArr = Split(strText, vbTab & "|")
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab & "|"
                            End If
                            strArr = Split(strText, vbTab & "|")
                            
                            If Val(strArr(0)) = 1 Then
                                str = strArr(1)
                                Exit Do
                            End If
                        Else
                             If "<" & STRNAME & ">" = strText Then
                                 blnStart = True
                             End If
                        End If
                        If "</" & STRNAME & ">" = strText Then
                            Exit Do
                        End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function

Private Function Get病人信息(ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strArr

    Get病人信息 = False
    
    Err = 0
    On Error GoTo errHand:
    'COMMENT ON COLUMN 保险帐户.病人ID   is '病人ID';
    'COMMENT ON COLUMN 保险帐户.险类     is '固定值:106';
    'COMMENT ON COLUMN 保险帐户.中心     is '0';
    'COMMENT ON COLUMN 保险帐户.卡号     is '卡号';
    'COMMENT ON COLUMN 保险帐户.医保号   is '统筹编号+个人编号';
    'COMMENT ON COLUMN 保险帐户.密码     is '无';
    'COMMENT ON COLUMN 保险帐户.人员身份 is '交易类别';
    'COMMENT ON COLUMN 保险帐户.单位编码 is '单位号码';
    '
    'COMMENT ON COLUMN 保险帐户.顺序号   is '只针对住院:住院流水号';
    'COMMENT ON COLUMN 保险帐户.退休证号 is '统筹地区编码|制卡日期|卡有效日期|制卡单位|在职情况';
    'COMMENT ON COLUMN 保险帐户.帐户余额 is '帐户余额';
    'COMMENT ON COLUMN 保险帐户.当前状态 is '0-门诊,1-在院';
    'COMMENT ON COLUMN 保险帐户.病种ID   is '无';
    'COMMENT ON COLUMN 保险帐户.在职     is '目前保存的值是1，无用处';
    'COMMENT ON COLUMN 保险帐户.年龄段   is '补卡次数';
    'COMMENT ON COLUMN 保险帐户.灰度级   is '工况类别';
    'COMMENT ON COLUMN 保险帐户.就诊时间 is '当前就诊的时间';
    '
    'COMMENT ON COLUMN 保险帐户.享受待遇标志 is '只针对住院:享受待遇标志';
    'COMMENT ON COLUMN 保险帐户.起付标准 is '只针对住院:起付标准';
    
'    gstrSQL = "select a.*,b.姓名,b.性别, b.年龄, b.出生日期, b.身份证号,b.工作单位,c.编码,C.名称 " & _
'             " from 保险帐户 a,病人信息 b,保险病种 C " & _
'             " WHERE  a.病种ID=c.ID and a.病人id=" & lng病人ID & " AND a.病人id=b.病人id and a.险类=" & TYPE_成都内江
    '2006-4-10  龚智毅修改
    gstrSQL = "select a.卡号,A.医保号,A.灰度级,A.单位编码,A.退休证号,A.年龄段,A.帐户余额,A.人员身份,A.顺序号,A.病人ID" & _
            ",b.姓名,b.性别, b.年龄, b.出生日期, b.身份证号,b.工作单位,c.编码,C.名称 " & _
            " from 保险帐户 a,病人信息 b,保险病种 C " & _
            " WHERE  a.出院病种ID=c.ID(+) and a.病人id=" & lng病人ID & " AND a.病人id=b.病人id and a.险类=" & TYPE_成都内江

    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人信息"

    With g病人身份_成都内江
        .卡号 = Nvl(rsTemp!卡号)
        .个人编号 = Mid(Nvl(rsTemp!医保号), 2)
        .身份证号 = Nvl(rsTemp!身份证号)
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Decode(Nvl(rsTemp!性别), "男", 1, "女", 2, 1)
        .工况类别 = Nvl(rsTemp!灰度级)
        .出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
        .单位号码 = Nvl(rsTemp!单位编码)
        strArr = Split(Nvl(rsTemp!退休证号) & "|||||", "|")
        .统筹编号 = strArr(0)
        .制卡日期 = strArr(1)
        .卡有效期 = strArr(2)
        .补卡次数 = Nvl(rsTemp!年龄段)
        .制卡单位 = strArr(3)
        .年龄 = Nvl(rsTemp!年龄, 0)
        .帐户余额 = Nvl(rsTemp!帐户余额, 0)
        .在职情况 = strArr(4)
        .交易类别 = Nvl(rsTemp!人员身份)
        .住院流水号 = Nvl(rsTemp!顺序号)
        .lng病人ID = Nvl(rsTemp!病人ID, 0)
        .病种编码 = Nvl(rsTemp!编码)
        .病种名称 = Nvl(rsTemp!名称)
        
    End With
    Get病人信息 = True
Exit Function
errHand:
        DebugTool "获取病人信息失败" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & " 错误信息:" & Err.Description
End Function

Private Sub OpenRecordset_成都内江(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional cnOracle As ADODB.Connection)
    '功能：打开记录集
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    If cnOracle Is Nothing Then
        rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_成都内江, adOpenStatic, adLockReadOnly
    Else
        If cnOracle.State <> 1 Then
            rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_成都内江, adOpenStatic, adLockReadOnly
        Else
            rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), cnOracle, adOpenStatic, adLockReadOnly
        End If
    End If
    Call SQLTest
End Sub


Public Function 住院虚拟结算_成都内江(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    
    'rsExse:字符集
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；

    Dim cn上传 As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim lng主页ID As Long
    Dim StrInput As String, strOutput   As String
    Dim strArr As Variant
    Dim intMouse As Integer
    Dim dbl费用总额 As Double
    Dim dbl生育包干 As Double
    Dim str结算方式  As String
    
    Err = 0: On Error GoTo errHand:

    g病人身份_成都内江.lng病人ID = lng病人ID
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If

    With g结算数据
        .比例支付 = 0
        .待遇标志 = ""
        .高额医保支付 = 0
        .公务员医疗补助 = 0
        .起付标准 = 0
        .医保交易流水号 = ""
        .医保内费用 = 0
        .医保外费用 = 0
        .帐户可用余额 = 0
        .帐户支付 = 0
    End With
    
    If Get病人信息(lng病人ID) = False Then Exit Function

    If bln结帐处 Then
        Screen.MousePointer = 1
        If 身份标识_成都内江(4, lng病人ID) = "" Then
            Screen.MousePointer = intMouse
            住院虚拟结算_成都内江 = ""
            Exit Function
        End If
        If lng病人ID <> g病人身份_成都内江.lng病人ID Then
            ShowMsgbox "你的卡可能有误,不能进行结算!"
            Exit Function
        End If
        Screen.MousePointer = intMouse
        
    Else
        Call Get病人信息(lng病人ID)
    End If
    
    '判断单位欠费情况
    '    个人编号    String (8)  IN
    '    社保卡号码  String (10) IN
    '    统筹地区编码    String (1)  IN
    StrInput = g病人身份_成都内江.个人编号
    StrInput = StrInput & vbTab & g病人身份_成都内江.卡号
    StrInput = StrInput & vbTab & g病人身份_成都内江.统筹编号
    
    If 业务请求_成都内江(获取单位欠缴情况_内江, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    If Val(strOutput) <> 0 Then
        ShowMsgbox "注意：" & vbCrLf & "    单位已经欠费"
    End If

    gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & rsExse("病人ID")
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "虚拟结算")
    If IsNull(rsTemp("主页ID")) = True Then
        MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    lng主页ID = rsTemp("主页ID")
    Screen.MousePointer = vbHourglass
    
    '补传明细
    If 补传住院明细记录(lng病人ID, lng主页ID) = False Then Exit Function
        
    
    g病人身份_成都内江.费用总额 = 0
    g病人身份_成都内江.生育总费用 = 0
    With rsExse
        Do While Not .EOF
            'g病人身份_成都内江.费用总额 = g病人身份_成都内江.费用总额 + Nvl(!金额, 0)
            
            'Beging 20051027 陈东
            Dim str大类编码 As String
            If Nvl(rsExse!保险大类id, 0) = 0 Then
                gstrSQL = "Select Nvl(大类编码,0) As 大类编码 From 保险项目 Where 险类=" & TYPE_成都内江 & _
                          " And 编码=(Select nvl(保险编码,'0') From 住院费用记录 Where 收费细目ID=" & rsExse!收费细目ID & _
                          " And NO='" & rsExse!NO & "' And 记录性质=" & rsExse!记录性质 & _
                          " And 记录状态=" & rsExse!记录状态 & " And  序号=" & rsExse!序号 & ")"
                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "取大类编码")
                
                str大类编码 = "0"
                If rsTemp.RecordCount > 0 Then
                    str大类编码 = Nvl(rsTemp!大类编码, "0")
                End If
                
                If Substr(Rpad(str大类编码, 1), 1, 1) <> "4" And Nvl(rsExse!摘要, "") <> "不上传" Then
                    g病人身份_成都内江.费用总额 = g病人身份_成都内江.费用总额 + Nvl(rsExse!金额, 0)
                End If
                If Substr(Rpad(str大类编码, 1), 1, 1) = "4" Or Nvl(rsExse!摘要, "") = "不上传" Then
                    g病人身份_成都内江.生育总费用 = g病人身份_成都内江.生育总费用 + Nvl(rsExse!金额, 0)
                End If
            Else
                If Nvl(rsExse!保险大类id, 0) <> "4" And Nvl(rsExse!摘要, "") <> "不上传" Then
                    g病人身份_成都内江.费用总额 = g病人身份_成都内江.费用总额 + Nvl(rsExse!金额, 0)
                End If
                If Nvl(rsExse!保险大类id, 0) = "4" Or Nvl(rsExse!摘要, "") = "不上传" Then
                    g病人身份_成都内江.生育总费用 = g病人身份_成都内江.生育总费用 + Nvl(rsExse!金额, 0)
                End If
            End If
            'End 20051027 陈东
            
            
            .MoveNext
        Loop
    End With
        
        
        
    gstrSQL = "Select A.ID From 住院费用记录 a,药品收发记录 B where A.no=b.No and B.单据 in (9,10) and a.id=b.费用ID and a.病人id=" & lng病人ID & " and 主页id=" & lng主页ID & " and b.扣率 like '_3%' and rownum<=2"

    
    Dim bln出院带药 As Boolean
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定出院带药"
    If rsTemp.EOF Then
        bln出院带药 = False
    Else
        bln出院带药 = True
    End If
    
    gstrSQL = "Select c.住院号,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 入院经办时间," & _
        " to_char(A.入院日期,'yyyyMMdd') 入院日期,D.诊断编码,A.出院方式,to_Char(a.出院日期,'yyyyMMDD') as 出院日期,a.出院病床,H.位置 as 出院科室" & _
        " From 病案主页 A,部门表 B,病人信息 C,部门表 H, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,2,b.编码,'')) AS 诊断编码 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =3  and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id)   D" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID and A.出院科室ID=H.id(+) " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        ""
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定入出类别"
    
    
     
     '入参:
     '    个人编号    String(8)   In
     '    社保卡号码  String(10)  In
     '    医院代码    String(5)   In
     '    操作员卡号码    String(10)  In
     '    统筹地区编码    String(1)   In
     '    出院日期    String(8)   In
     '    出院科别    String(10)  In
     '    出院诊治医生    String(10)  In
     '    诊断编码    String(20)  In
     '    出院带药    String(1)   In
     '    出院类别    String(1)   In
     '    住院流水号  String(20)  In
       
     With g病人身份_成都内江
         StrInput = Rpad(.个人编号, 8)
         StrInput = StrInput & vbTab & Rpad(.卡号, 10)
         StrInput = StrInput & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
         StrInput = StrInput & vbTab & Rpad(Nvl(gstrUserName), 10)
         StrInput = StrInput & vbTab & Rpad(.统筹编号, 1)
         
         If Trim(Nvl(rsTemp!出院日期)) = "" Then
            '如果没有出院日期,则传当前的时间上去.
            StrInput = StrInput & vbTab & Format(zlDatabase.Currentdate, "yyyymmdd")
         Else
            StrInput = StrInput & vbTab & Rpad(Nvl(rsTemp!出院日期), 8)
         End If
         
         StrInput = StrInput & vbTab & Rpad(Substr(Rpad(Nvl(rsTemp!出院科室), 10), 1, 10), 10)
         StrInput = StrInput & vbTab & Rpad(Substr(Rpad(Nvl(rsTemp!住院医师), 10), 1, 10), 10)
        'Beging 20051026 陈东
        Dim vat并发症 As Variant, str出院其他病种 As String, i As Long
        Dim rsCydj As New ADODB.Recordset
        
        gstrSQL = "Select * from 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_成都内江
        Call zlDatabase.OpenRecordset(rsCydj, gstrSQL, "取出院其他病种")
        str出院其他病种 = Nvl(rsCydj!出院其他病种)
        If str出院其他病种 <> "" Then
            If InStr(str出院其他病种, "|") > 0 Then
                vat并发症 = Split(str出院其他病种, "|")
                str出院其他病种 = ""
                For i = 0 To UBound(vat并发症) - 1
                    str出院其他病种 = str出院其他病种 & Rpad(Substr(vat并发症(i), 1, 20), 20)
                Next
            End If
        Else
            str出院其他病种 = Space(180)
        End If
        'End 20051026 陈东
         StrInput = StrInput & vbTab & Rpad(Rpad(Substr(g病人身份_成都内江.病种编码, 1, 20), 20) & Substr(str出院其他病种, 1, 180), 200)
         StrInput = StrInput & vbTab & IIf(bln出院带药, "1", "0")
         StrInput = StrInput & vbTab & Rpad(g病人身份_成都内江.出院类别, 1)
         'strInput = strInput & vbTab & Rpad(Get治渝情况_内江(lng病人ID, lng主页ID), 1)
         StrInput = StrInput & vbTab & Rpad(Substr(Rpad(.住院流水号, 20), 1, 20), 20)
         'Beging 陈东 20051020
         If bln结帐处 = True Then
            If 业务请求_成都内江(出院登记上传_内江, StrInput, strOutput) = False Then Exit Function
         Else
            Exit Function
         End If
         'End 陈东 20051020
     End With
     
     If strOutput = "" Then Exit Function
     
     strArr = Split(strOutput, vbTab)
     
    '出参
     '    TRANSDETIAL输出 (计算费用明细)
     '    享受待遇标志    String(1)   Out
     '    医保内费用  String(10)  Out
     '    医保外费用  String(10)  Out
     '    基本医保支付
     '    如果参加大病医保，则为大病医保支付  String(10)  Out
     '    高额医保支付    String(10)  Out
     '    公务员医疗补助  String(10)  Out
     '    个人按比例支付  String(10)  Out
     '    TRANSDETIAL结束
     '    起付标准    String(10)  Out
     '    个人帐户可用余额    String(10)  Out
     strOutput = strArr(0)
     With g结算数据
         .结算标志 = 1
         .待遇标志 = Substr(strOutput, 1, 1)
         .医保内费用 = Val(Substr(strOutput, 2, 10)) / 100
         .医保外费用 = Val(Substr(strOutput, 12, 10)) / 100
         .基本医保支付 = Val(Substr(strOutput, 22, 10)) / 100
         .高额医保支付 = Val(Substr(strOutput, 32, 10)) / 100
         .公务员医疗补助 = Val(Substr(strOutput, 42, 10)) / 100
         .比例支付 = Val(Substr(strOutput, 52, 10)) / 100
         .起付标准 = Val(strArr(1)) / 100
         .帐户可用余额 = Val(strArr(2)) / 100
         '20051021 add
         .生育支付 = Val(Substr(strOutput, 62, 10)) / 100
     End With
     
     
    dbl费用总额 = g结算数据.医保内费用 + g结算数据.医保外费用
    '龚智毅 20060809 内江二院要求医保外费用可以用帐户支付
    'g结算数据.帐户支付 = dbl费用总额 - g结算数据.医保外费用 - g结算数据.基本医保支付 - g结算数据.高额医保支付 - g结算数据.公务员医疗补助
    g结算数据.帐户支付 = dbl费用总额 - g结算数据.基本医保支付 - g结算数据.高额医保支付 - g结算数据.公务员医疗补助
    'by 20050122 gzy
    If g结算数据.帐户可用余额 >= 0 Then
        If g结算数据.帐户可用余额 < g结算数据.帐户支付 Then
           g结算数据.帐户支付 = g结算数据.帐户可用余额
        End If
    Else
        g结算数据.帐户支付 = 0
    End If
    If g结算数据.帐户支付 <= 0 Then g结算数据.帐户支付 = 0
    
    str结算方式 = "个人帐户;" & g结算数据.帐户支付 & ";1"
    
    If g结算数据.基本医保支付 <> 0 Then
        str结算方式 = str结算方式 & "|基本统筹;" & g结算数据.基本医保支付 & ";0"
    End If
    If g结算数据.高额医保支付 <> 0 Then
        str结算方式 = str结算方式 & "|大病支付;" & g结算数据.高额医保支付 & ";0"
    End If
    If g结算数据.公务员医疗补助 <> 0 Then
        str结算方式 = str结算方式 & "|公务员补助;" & g结算数据.公务员医疗补助 & ";0"
    End If
    
    g结算数据.生育盈亏 = 0
    If g结算数据.生育支付 <> 0 Then
        str结算方式 = str结算方式 & "|生育支付;" & g结算数据.生育支付 & ";0"
        'Modified by ZYB 20051118
        '将医保中心包干支付费用与实际发生的生育费用进行比较，多余部分计入生育盈亏中
        g结算数据.生育盈亏 = g病人身份_成都内江.生育总费用 - g结算数据.生育支付
    End If
    
    '龚智毅 20051029 增加生育费用包干模式
    'Modified by ZYB 20051118
    If g结算数据.生育盈亏 <> 0 Then
        str结算方式 = str结算方式 & "|生育盈亏;" & g结算数据.生育盈亏 & ";0"
        ShowMsgbox "本院发生的生育费用为：" & g病人身份_成都内江.生育总费用 & vbCrLf & _
                  "医保中心报销生育费用：" & g结算数据.生育支付 & vbCrLf & _
                  "生育费用当前计算方式： 多退少补  " & vbCrLf & _
                  "生育盈亏： " & g结算数据.生育盈亏
    End If
    
    'If Format(g病人身份_成都内江.费用总额, "###0.00;-###0.00;0;0") <> Format(dbl费用总额, "###0.00;-###0.00;0;0") Then
    '    Dim blnYes As Boolean
    '    费用总额与医保中心返回总额不致,不能进行结算
    '    ShowMsgbox "本次结算总额(" & g病人身份_成都内江.费用总额 & ") 与" & vbCrLf & _
    '                "   中心返回的总额(" & dbl费用总额 & ")不等,不能结算?"
    '    Exit Function
    'End If
    
    'gzy 20051129 普通生育病人不判断医保金额
    gstrSQL = "Select * From 保险病种 Where id=(select 病种ID from 保险帐户 Where 病人ID=" & lng病人ID & ")" & _
              " And (名称 ='平产' or 名称='剖宫产')"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "取入院诊断")
    If rsTemp.EOF = True Then
       If Format(g病人身份_成都内江.费用总额, "###0.00;-###0.00;0;0") <> Format(dbl费用总额, "###0.00;-###0.00;0;0") Then
           Dim blnYes As Boolean
           '费用总额与医保中心返回总额不致,不能进行结算
           ShowMsgbox "本次结算非生育费用总额(" & g病人身份_成都内江.费用总额 & ") 与" & vbCrLf & _
                    "   中心返回的非生育费用总额(" & dbl费用总额 & ")不等,不能结算?"
           Exit Function
       End If
    Else
       gstrSQL = "Select * from 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_成都内江
       Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "附加诊断")
       If Nvl(rsTemp!附加诊断) <> "" Then
          If Format(g病人身份_成都内江.费用总额, "###0.00;-###0.00;0;0") <> Format(dbl费用总额, "###0.00;-###0.00;0;0") Then
              'Dim blnYes As Boolean
              '费用总额与医保中心返回总额不致,不能进行结算
              ShowMsgbox "本次结算非生育费用总额(" & g病人身份_成都内江.费用总额 & ") 与" & vbCrLf & _
                    "   中心返回的非生育费用总额(" & dbl费用总额 & ")不等,不能结算?"
              Exit Function
          End If
       Else
          If g病人身份_成都内江.费用总额 <> 0 Then
             ShowMsgbox "补传了生育费用(" & g病人身份_成都内江.费用总额 & ")，请重新进行结算。"
              Exit Function
          End If
       End If
    End If
    
    住院虚拟结算_成都内江 = str结算方式
    g病人身份_成都内江.lng病人ID = lng病人ID   '表示该病人已经进行了虚拟结算
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get治渝情况_内江(lng病人ID As Long, lng主页ID As Long) As String
    '功能:获取治渝情况标识
    '     A-治愈、B-好转、C-未愈、D-死亡、E-其他
    '??49  治愈情况标识    CHAR    439 1   1治愈、2好转、3未愈、4死亡、5其他，住院必添 院端
    'A-治愈、B-好转、C-未愈、D-死亡、E-其他
    
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.出院情况" & _
             " From 诊断情况 A,疾病编码目录 B " & _
             " Where A.病人ID=" & lng病人ID & " And A.疾病ID=B.ID(+) And A.主页ID=" & lng主页ID & _
             "       And A.诊断类型 in (2,3)" & _
             " Order by A.诊断类型 Desc"
    
    rsInNote.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rsInNote, strTmp, "医保接口")
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = Nvl(rsInNote!出院情况)
    End If
    strTmp = Decode(strTmp, "治愈", "0", "好转", "1", "未愈", "2", "死亡", "3", "自动出院", 4, "转本地统筹地区内医院", 5, "转本地统筹地区外医院", 6, "其他", "3")
    Get治渝情况_内江 = strTmp
End Function

Private Function 补传住院明细记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '补传相关明细记录
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput  As String, strOutput As String
    Dim strArr, strArr摘要
    Dim lng冲销ID As Long
    Err = 0
    On Error GoTo errHand:


    补传住院明细记录 = False

    '读出未上传明细（排序，以便先上传正明细，再上传负明细）
    gstrSQL = "" & _
        "   Select distinct A.NO,A.记录性质,A.记录状态 " & _
        "   From 住院费用记录 A " & _
        "   Where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & " and A.记帐费用=1  and nvl(A.实收金额,0)<>0 and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 " & _
        "   Order by A.NO,A.记录性质,Decode(A.记录状态,2,2,1)"
        
    
   zlDatabase.OpenRecordset rs明细, gstrSQL, "获取补传明细记录"
    '先检查是否存在退单情况，如果存在，看有否对应的记录单.
    With rs明细
'        '上传明细
'        If .RecordCount <> 0 Then .MoveFirst
'        Do While Not .EOF
'            If Nvl(!项目编码) = "" Then
'                ShowMsgbox "项目:[" & Nvl(!编码) & "] 未设置对应的医保项目,请设置对应关系!"
'                Exit Function
'            End If
'            If (Val(!数量) < 0 Or Val(!价格) < 0) And rs明细!记录状态 = 1 Then
'                ShowMsgbox "项目:[" & Nvl(!编码) & "] 不能输入负单据!"
'                Exit Function
'            End If
'            .MoveNext
'        Loop
    End With
    '先传正单据
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!记录状态, 1) = 1 Or Nvl(!记录状态, 1) = 3 Then
                '上传指定处方
                If 处方上传(Nvl(!记录性质, 0), Nvl(!记录状态, 0), Nvl(!NO)) = False Then
                    Exit Function
                End If
            Else
                '上传指定处方
                gcnOracle_成都内江.BeginTrans
                gcnOracle.BeginTrans
                If 处方作废(Nvl(!记录性质, 0), Nvl(!记录状态, 0), Nvl(!NO)) = False Then
                    gcnOracle.RollbackTrans
                    gcnOracle_成都内江.RollbackTrans
                    Exit Function
                    
                End If
                gcnOracle.CommitTrans
                gcnOracle_成都内江.CommitTrans
            End If
           .MoveNext
        Loop
    End With
    补传住院明细记录 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
'Private Function Get原单据摘要(ByVal strNO As String, ByVal int序号 As Integer, ByVal int性质 As Integer) As Variant
'    '根据指定的值，获取摘要的相关信息
'    Dim rsTemp As New ADODB.Recordset
'    Dim strTemp As String
'
'
'    gstrSQL = " Select 摘要 From 病人费用记录" & _
'              " Where NO='" & strNO & "' And 序号=" & int序号 & _
'              " And 记录性质=" & int性质 & " And 记录状态=3"
'
'    Call OpenRecordset(rsTemp, "取原始处方明细的流水号")
'
'    If Not rsTemp.EOF Then
'        strTemp = Nvl(rsTemp!摘要) & "|||||||"
'    Else
'        strTemp = "|||||||"
'    End If
'    Get原单据摘要 = Split(strTemp, "|")
'End Function

'----200410刘兴宏加入
Public Function 医保设置_成都内江() As Boolean
    医保设置_成都内江 = frmSet成都内江.参数设置
    
End Function
'
'Public Function 下载服务项目目录_成都内江(ByVal bytType As Byte, ByVal objProgss As Object) As Boolean
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:下载服务项目目录
'    '参数:bytType-1-药品,2-诊疗,3-服务,4-费用类别,5-病种目录
'    '返回:下载成功,返回true,否则返回False
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSql As String
'    Dim rsTemp As New ADODB.Recordset
'    Dim strDate As String, strInput As String, strOutput As String
'    Dim lngCount As Long
'    Dim i As Long
'    Dim strArr
'    Dim strTitle As String
'
'    下载服务项目目录_成都内江 = False
'    strTitle = Switch(bytType = 1, "药品", bytType = 2, "诊疗项目", bytType = 3, "服务设施", bytType = 4, "费用类别", True, "病种目录取")
'
'    Err = 0
'    On Error GoTo ErrHand:
'    strSql = "" & _
'        "   Select to_char(Max(变更时间),'yyyy-mm-dd hh24:mi:ss')  as 变更时间 " & _
'        "   From 医保收费目录 " & _
'        "   where 类别=" & bytType
'    zlDatabase.OpenRecordset rsTemp, strSql, "获取最大变更时间"
'
'    strDate = Nvl(rsTemp!变更时间)
'    strDate = IIf(strDate = "", "1977-01-01 00:00:00", strDate)
'
'    If Not objProgss Is Nothing Then
'    Else
'        zlCommFun.ShowFlash "正在下载《" & strTitle & "》数据,请等待..."
'    End If
'    '预处理
'    strInput = bytType & "|" & strDate
'    If 业务请求_成都内江(收费目录下载预处理, strInput, strOutput) = False Then Exit Function
'    strArr = Split(strOutput, "|")
'    lngCount = Val(strArr(1))
'
'    If Not objProgss Is Nothing Then
'        objProgss.Max = IIf(lngCount = 0, 1, lngCount)
'        objProgss.Min = 1
'        objProgss.Value = 1
'    End If
'
'   For i = 1 To lngCount
'        '正试下载
'        If 业务请求_成都内江(收费目录下载处理, strInput, strOutput) = False Then Exit Function
'        strArr = Split(strOutput, "|")
'        '更新收费目录
'
'        '过程:类别,编码,名称,英文名称,收费类别,收费等级,费用等级,拼音码,单位,单价,规格,备注,变更时间,可维护标志,支付标准
'        strSql = "ZL_医保收费目录_UPDATE("
'        strSql = strSql & bytType & ",'"
'        strSql = strSql & strArr(1) & "','" '编码
'        strSql = strSql & strArr(2) & "','" '名称
'        Select Case bytType
'        Case 1
'            strSql = strSql & strArr(3) & "','" '英文名称
'            strSql = strSql & strArr(4) & "','" '收费类别
'            strSql = strSql & strArr(5) & "','" '费用等级
'            strSql = strSql & strArr(6) & "','" '拼音码
'            strSql = strSql & strArr(7) & "','" '单位
'            strSql = strSql & strArr(8) & "','" '单价
'            strSql = strSql & strArr(9) & "','" '剂型
'            strSql = strSql & strArr(10) & "','" '规格
'            strSql = strSql & strArr(11) & "',to_date('" '备注
'            strSql = strSql & strArr(12) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & strArr(13) & "','"     '可维护标志
'            strSql = strSql & "" & "')" '支付标准
'        Case 2
'            strSql = strSql & "" & "','" '英文名称
'            strSql = strSql & strArr(3) & "','" '收费类别
'            strSql = strSql & "" & "','" '费用等级
'            strSql = strSql & strArr(4) & "','" '拼音码
'            strSql = strSql & strArr(5) & "','" '单位
'            strSql = strSql & strArr(6) & "','" '单价
'            strSql = strSql & "" & "','" '剂型
'            strSql = strSql & "" & "','" '规格
'            strSql = strSql & strArr(7) & "',to_date('" '备注
'            strSql = strSql & strArr(8) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & strArr(9) & "','"     '可维护标志
'            strSql = strSql & "" & "')" '支付标准
'        Case 3
'            strSql = strSql & "" & "','" '英文名称
'            strSql = strSql & strArr(3) & "','" '收费类别
'            strSql = strSql & "" & "','" '费用等级
'            strSql = strSql & strArr(6) & "','" '拼音码
'            strSql = strSql & "" & "','" '单位
'            strSql = strSql & strArr(4) & "','" '单价
'            strSql = strSql & "" & "','" '剂型
'            strSql = strSql & "" & "','" '规格
'            strSql = strSql & "" & "',to_date('" '备注
'            strSql = strSql & strArr(7) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & "" & "',"     '可维护标志
'            strSql = strSql & strArr(5) & "')" '支付标准
'        Case 4
'            ' 费用类别编码|费用类别名称
'            strSql = "ZL_医保收费类别_UPDATE("
'
'            strSql = strSql & strArr(1) & "','" '编码
'            strSql = strSql & strArr(2) & "')" '名称
'        Case Else
'            '病种编码|病种名称|拼音码|变更日期
'            strSql = "ZL_医保病种目录_UPDATE("
'            strSql = strSql & strArr(1) & "','" '编码
'            strSql = strSql & strArr(2) & "','" '名称
'            strSql = strSql & strArr(3) & "',to_date('" '助记码
'            strSql = strSql & strArr(4) & "','yyyy-mm-dd hh24:mi:ss')" '变更时间
'        End Select
'        gcnOracle_成都内江.Execute strSql, , adCmdStoredProc
'        If Not objProgss Is Nothing Then
'            objProgss.Value = i
'        Else
'            zlCommFun.ShowFlash "正在下载《" & strTitle & "》数据,已下载" & i & "/" & lngCount & ""
'        End If
'   Next
'   下载服务项目目录_成都内江 = True
'   Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'End Function

Public Function 获取参保人员信息_成都内江(ByVal StrInput As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------

    '获取参保人员信息
    Dim strOutput As String
    Dim strArr
    
    Dim str出生日期 As String
    获取参保人员信息_成都内江 = False
    
    Err = 0
    On Error GoTo errHand:
    
    If 业务请求_成都内江(读病人信息_内江, StrInput, strOutput) = False Then Exit Function
    '返回串是:卡号vbtab个人编号vbtab身份证号vbtab姓名vbtab性别vbtab工况类别vbtab出生日期vbtab单位号码vbtab统筹编号vbtab制卡日期vbtab卡有效期vbtab补卡次数vbtab制卡单位
    If strOutput = "" Then Exit Function
    strArr = Split(strOutput, vbTab)
    
    With g病人身份_成都内江
        .卡号 = strArr(0)
        .个人编号 = strArr(1)
        .身份证号 = strArr(2)
        .姓名 = strArr(3)
        .性别 = strArr(4)
        .工况类别 = strArr(5)
        '.出生日期 = zlCommFun.AddDate(strArr(6))
        '陈东 20050601
        If 身份证号转出生日期(strArr(2), str出生日期) = True Then
            .出生日期 = zlCommFun.AddDate(str出生日期)
        Else
            MsgBox str出生日期 & "将采用出生日期字段的值", vbInformation, gstrSysName
            .出生日期 = zlCommFun.AddDate(strArr(6))
        End If
        .单位号码 = strArr(7)
        .统筹编号 = strArr(8)
        .制卡日期 = strArr(9)
        .卡有效期 = strArr(10)
        .补卡次数 = strArr(11)
        .制卡单位 = strArr(12)
        .年龄 = Get年龄(.出生日期)
    End With
    
    '--获取个人帐户余额
    If 获取帐户余额_成都内江() = False Then Exit Function
    
    获取参保人员信息_成都内江 = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Public Function 获取帐户余额_成都内江() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取病人当前帐户余额
    '--入参数:
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, strOutput As String
    Dim strArr
    Err = 0
    On Error GoTo errHand:
    获取帐户余额_成都内江 = False
    With g病人身份_成都内江
        '    个人编号    String (8)  IN
        '    社保卡号码  String (10) IN
        '    统筹地区编码    String (1)  IN
        StrInput = .个人编号
        StrInput = StrInput & vbTab & .卡号
        StrInput = StrInput & vbTab & .统筹编号
    End With
    
    If 业务请求_成都内江(获取帐户余额_内江, StrInput, strOutput) = False Then Exit Function
    If strOutput = "" Then Exit Function
    strArr = Split(strOutput, vbTab)
    With g病人身份_成都内江
        .帐户余额 = Val(strArr(0))
        .在职情况 = strArr(1)
    End With
    获取帐户余额_成都内江 = True
    Exit Function
errHand:
        If ErrCenter = 1 Then Resume
End Function

Private Function Get年龄(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-[1])/365 as 年龄 from dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取年龄", CDate(strDate))
    If Not rsTemp.EOF Then
        Get年龄 = Int(Nvl(rsTemp!年龄, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function

Private Function GetErrInfor(ByVal strErrCode As String) As String
        Dim strErrMsg As String
        
        Select Case strErrCode
                '---读卡的相关错误
                Case "1 ": strErrMsg = " 错误号：1 " & vbCrLf & " 错误描述：检测通讯方式错误(chk_baud异常错误)"
                Case "2 ": strErrMsg = " 错误号：2 " & vbCrLf & " 错误描述：始化端口错误(auto_init)"
                Case "3 ": strErrMsg = " 错误号：3 " & vbCrLf & " 错误描述：关闭通讯口错误(ic_exit)"
                Case "4 ": strErrMsg = " 错误号：4 " & vbCrLf & " 错误描述：读写器错误"
                Case "5 ": strErrMsg = " 错误号：5 " & vbCrLf & " 错误描述：无法初始化卡密码"
                Case "10": strErrMsg = " 错误号：10" & vbCrLf & " 错误描述： 检测读写器中是否有卡错误(get_status)"
                Case "11": strErrMsg = " 错误号：11" & vbCrLf & " 错误描述： 卡型错误（非4428卡）(chk_4428)"
                Case "12": strErrMsg = " 错误号：12" & vbCrLf & " 错误描述： 卡密码错误(csc_4428)"
                Case "13": strErrMsg = " 错误号：13" & vbCrLf & " 错误描述： 修改卡密码错误"
                Case "20": strErrMsg = " 错误号：20" & vbCrLf & " 错误描述： 读卡芯数据错误(srd_4428)"
                Case "21": strErrMsg = " 错误号：21" & vbCrLf & " 错误描述： 写卡芯数据(用户数据)错误(swr_4428)"
                Case "23": strErrMsg = " 错误号：23" & vbCrLf & " 错误描述： 写卡芯数据(用户密码)错误(swr_4428)"
                Case "30": strErrMsg = " 错误号：30" & vbCrLf & " 错误描述： 用户密码错误"
                Case "31": strErrMsg = " 错误号：31" & vbCrLf & " 错误描述： 用户数据加密错误(ic_decrypt)"
                Case "32": strErrMsg = " 错误号：32" & vbCrLf & " 错误描述： 用户数据解密错误(ic_decrypt)"
                Case "33": strErrMsg = " 错误号：33" & vbCrLf & " 错误描述： 用户密码加密错误(ic_encrypt)"
                Case "34": strErrMsg = " 错误号：34" & vbCrLf & " 错误描述： 用户密码解密错误(ic_decrypt)"
                Case "35": strErrMsg = " 错误号：35" & vbCrLf & " 错误描述： 用户原密码长度为零或者大于6"
                Case "36": strErrMsg = " 错误号：36" & vbCrLf & " 错误描述： 用户新密码长度为零或者大于6"
                Case "40": strErrMsg = " 错误号：40" & vbCrLf & " 错误描述：   不能打开数据库"
                Case "41": strErrMsg = " 错误号：41" & vbCrLf & " 错误描述：   没有制卡数据"
                Case "42": strErrMsg = " 错误号：42" & vbCrLf & " 错误描述：   个人信息不完整（姓名、性别、民族等文本信息）"
                '---医保接口返回的相关错误
                Case "000": strErrMsg = "执行成功"
                Case "001": strErrMsg = " 错误号： 001" & vbCrLf & " 错误描述：读卡器无响应"
                Case "002": strErrMsg = " 错误号： 002" & vbCrLf & " 错误描述：没有社保卡"
                Case "003": strErrMsg = " 错误号： 003" & vbCrLf & " 错误描述：社保卡无响应"
                Case "004": strErrMsg = " 错误号： 004" & vbCrLf & " 错误描述：主机无响应"
                Case "051": strErrMsg = " 错误号： 051" & vbCrLf & " 错误描述：输入参数不足"
                Case "052": strErrMsg = " 错误号： 052" & vbCrLf & " 错误描述：卡号与社保号码不符"
                Case "053": strErrMsg = " 错误号： 053" & vbCrLf & " 错误描述：处方明细与实际纪录数量不符"
                Case "054": strErrMsg = " 错误号： 054" & vbCrLf & " 错误描述：没有此交易流水号"
                Case "055": strErrMsg = " 错误号： 055" & vbCrLf & " 错误描述：处方项目不符"
                Case "056": strErrMsg = " 错误号： 056" & vbCrLf & " 错误描述：没有此住院流水号（输入医保交易号+输入的住院流水号与库表中医保流水号对应的住院流水号不一致）等等"
                Case "058": strErrMsg = " 错误号： 058" & vbCrLf & " 错误描述：重复业务操作例如：已住院还进行住院操作，已出院还进行出院操作"
                Case "059": strErrMsg = " 错误号： 059" & vbCrLf & " 错误描述：输入社保号码和对应交易流水号不一致(住院时用住院流水号对应)"
                Case "060": strErrMsg = " 错误号： 060" & vbCrLf & " 错误描述：流水号不为最大撤消时"
                Case "061": strErrMsg = " 错误号： 061" & vbCrLf & " 错误描述：交易未确认上传住院交易前"
                Case "062": strErrMsg = " 错误号： 062" & vbCrLf & " 错误描述：未进行住院登记"
                Case "071": strErrMsg = " 错误号： 071" & vbCrLf & " 错误描述：医院交易流水号异常由HIS系统生成（空，长度不正常）"
                Case "072": strErrMsg = " 错误号： 072" & vbCrLf & " 错误描述：重复数据包传送"
                Case "073": strErrMsg = " 错误号： 073" & vbCrLf & " 错误描述：交叉数据包传送"
                Case "074": strErrMsg = " 错误号： 074" & vbCrLf & " 错误描述：应该上传明细而没有上传明细"
                Case "075": strErrMsg = " 错误号： 075" & vbCrLf & " 错误描述：检查项目种类异常项目种类不在[1]，[2]之内"
                Case "077": strErrMsg = " 错误号： 077" & vbCrLf & " 错误描述：上传医疗机构编码异常在KB01之中不存在"
                Case "078": strErrMsg = " 错误号： 078" & vbCrLf & " 错误描述：非定点医疗机构"
                Case "079": strErrMsg = " 错误号： 079" & vbCrLf & " 错误描述：社保卡号与医保交易号不对应门诊消费撤销时，只允许用消费的卡来撤销交易"
                Case "080": strErrMsg = " 错误号： 080" & vbCrLf & " 错误描述：医疗机构编号与交易流水号不对应(住院时用住院流水号对应)只允许交易医疗机构撤销自己的交易"
                Case "081": strErrMsg = " 错误号： 081" & vbCrLf & " 错误描述：对应流水号明细不存在Kc07,KC08K1,Kc08k2"
                Case "082": strErrMsg = " 错误号： 082" & vbCrLf & " 错误描述：没有对应的住院流水号Kc08"
                Case "083": strErrMsg = " 错误号： 083" & vbCrLf & " 错误描述：不允许撤销已出院的交易已经出院的交易补允许撤销"
                Case "084": strErrMsg = " 错误号： 084" & vbCrLf & " 错误描述：出院医院不是入院医院"
                Case "085": strErrMsg = " 错误号： 085" & vbCrLf & " 错误描述：输入输出参数长度错误输入长度与约定长度不符"
                Case "086": strErrMsg = " 错误号： 086" & vbCrLf & " 错误描述：出院人员不是原来那个住院人员防止一卡号对应多个人编号的情况"
                Case "101": strErrMsg = " 错误号： 101" & vbCrLf & " 错误描述：个人状态异常"
                Case "102": strErrMsg = " 错误号： 102" & vbCrLf & " 错误描述：社保卡为黑名单卡"
                Case "103": strErrMsg = " 错误号： 103" & vbCrLf & " 错误描述：帐户被冻结"
                Case "104": strErrMsg = " 错误号： 104" & vbCrLf & " 错误描述：不能享受统筹待遇"
                Case "106": strErrMsg = " 错误号： 106" & vbCrLf & " 错误描述：不存在此人"
                Case "107": strErrMsg = " 错误号： 107" & vbCrLf & " 错误描述：没有参加医疗保险或生育保险"
                Case "108": strErrMsg = " 错误号： 108" & vbCrLf & " 错误描述：帐户被注销只有死亡，出国定居才存在注销"
                Case "109": strErrMsg = " 错误号： 109" & vbCrLf & " 错误描述：出院已上传，但未确认出院上传时"
                Case "110": strErrMsg = " 错误号： 110" & vbCrLf & " 错误描述：出院已确认出院确认传时"
                Case "111": strErrMsg = " 错误号： 111" & vbCrLf & " 错误描述：没有此人医保卡数据T_cardinfo中"
                Case "112": strErrMsg = " 错误号： 112" & vbCrLf & " 错误描述：挂失卡已经挂失"
                Case "113": strErrMsg = " 错误号： 113" & vbCrLf & " 错误描述：卡状态异常"
                Case "114": strErrMsg = " 错误号： 114" & vbCrLf & " 错误描述：不存在个人账户Kc04无数据"
                Case "115": strErrMsg = " 错误号： 115" & vbCrLf & " 错误描述：系统没有当年起付线参数"
                Case "116": strErrMsg = " 错误号： 116" & vbCrLf & " 错误描述：没有当年基本医疗保险报销比例"
                Case "117": strErrMsg = " 错误号： 117" & vbCrLf & " 错误描述：没有当年大病医疗保险报销比例"
                Case "118": strErrMsg = " 错误号： 118" & vbCrLf & " 错误描述：没有当年高额医疗保险报销比例"
                Case "119": strErrMsg = " 错误号： 119" & vbCrLf & " 错误描述：没有当年公务员医疗保险报销比例"
                Case "120": strErrMsg = " 错误号： 120" & vbCrLf & " 错误描述：入院时间无效"
                Case "121": strErrMsg = " 错误号： 121" & vbCrLf & " 错误描述：没有单位封锁信息Kb02"
                Case "150": strErrMsg = " 错误号： 150" & vbCrLf & " 错误描述：支出金额为负数输入的支出的金额为负数"
                Case "151": strErrMsg = " 错误号： 151" & vbCrLf & " 错误描述：个人帐户超支（扣减为负）账户扣出负数来"
                Case "152": strErrMsg = " 错误号： 152" & vbCrLf & " 错误描述：基本统筹超支"
                Case "153": strErrMsg = " 错误号： 153" & vbCrLf & " 错误描述：大病统筹超支"
                Case "154": strErrMsg = " 错误号： 154" & vbCrLf & " 错误描述：公务员医疗补助超支"
                Case "155": strErrMsg = " 错误号： 155" & vbCrLf & " 错误描述：账户支付费用超过个人应当支付费用"
                Case "255": strErrMsg = " 错误号： 255" & vbCrLf & " 错误描述：服务程序出错"
                Case "998": strErrMsg = " 错误号： 998" & vbCrLf & " 错误描述：获取各类流水号失败"
                Case "999": strErrMsg = " 错误号： 999" & vbCrLf & " 错误描述：数据库sql错误，或者未找到数据"
                Case "800": strErrMsg = " 错误号： 800" & vbCrLf & " 错误描述：未办理出院手续，却试图进行出院确认操作"
                Case "801": strErrMsg = " 错误号： 801" & vbCrLf & " 错误描述：已经出院确认，再次试图办理出院确认"
                '20051020 陈东 add
                Case "122": strErrMsg = " 错误号： 122" & vbCrLf & " 错误描述：生育支付项目参数错误"
                
                Case "200": strErrMsg = " 错误号： 200" & vbCrLf & " 错误描述：对帐开始日期到终止日期没有费用明细"
                Case "201": strErrMsg = " 错误号： 201" & vbCrLf & " 错误描述：本月已有对帐信息，不能再次对帐"
                Case "202": strErrMsg = " 错误号： 202" & vbCrLf & " 错误描述：对帐类别错误"
                Case "203": strErrMsg = " 错误号： 203" & vbCrLf & " 错误描述： 没有在系统指定日期对帐"
                    
                Case "210": strErrMsg = " 错误号： 210" & vbCrLf & " 错误描述：上传诊断代码重复"
                Case "211": strErrMsg = " 错误号： 211" & vbCrLf & " 错误描述：上传诊断代码不存在"
                    
                Case "220": strErrMsg = " 错误号： 220" & vbCrLf & " 错误描述：一次住院不能多次生育"
                Case "221": strErrMsg = " 错误号： 221" & vbCrLf & " 错误描述：一次门诊不能多次生育"
                Case "222": strErrMsg = " 错误号： 222" & vbCrLf & " 错误描述：药店不能上传生育类费用"
                Case "223": strErrMsg = " 错误号： 223" & vbCrLf & " 错误描述：没有发生生育时不能上传多胎费用"
                Case "224": strErrMsg = " 错误号： 224" & vbCrLf & " 错误描述：没有剖宫产时不能上传全身麻醉费用"
                Case "225": strErrMsg = " 错误号： 225" & vbCrLf & " 错误描述：门诊不能上传剖宫产费用或全身麻醉费用"
                    
                Case "230": strErrMsg = " 错误号： 230" & vbCrLf & " 错误描述：非生育住院不能上传生育费用"
                Case "231": strErrMsg = " 错误号： 231" & vbCrLf & " 错误描述：生育住院没有申请并发症不能上传医疗费用"
                
            Case Else
                strErrMsg = "不能确定的错误代码,代码号为" & strErrCode
    End Select
    GetErrInfor = strErrMsg
End Function
Public Sub ExecuteProcedure_ZLNJ(ByVal strCaption As String)
'功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_成都内江.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Private Function 结算方式更正(类别 As Integer, Optional ByRef strAdvance = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:更正及显示结算结果
    '--入参数:
    '--出参数:str结算方式
    '--返  回:成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, str虚拟结算返回串 As String
    Dim dbl费用总额 As Double
        
    '费用总额=病人自费金额+基本统筹支付金额+大病统筹金额      此解释是由刘兴宏根据以面公式转换而来的
    
    '病人自费金额 = 总费用额 - 基本统筹支付金额 - 大病 / 高额统筹支付金额
    '自费金额＝现金支付额＋帐户支付额 (即:可选择由现金或用帐户支付)
    '大病统筹与高额统筹意义相同
    '统筹支付金额等于医保内费用根据不同的起付标准和报销比例由医保中心算
    '此说明依据北京科瑞奇技术开发股份有限公司蒋红彬负责的解释
    结算方式更正 = False
    
    Err = 0
    On Error GoTo errHand:
    DebugTool "进入(" & "Get结算方式" & ")"
    
'    If g结算数据.结算标志 = 0 Then
        dbl费用总额 = g结算数据.医保内费用 + g结算数据.医保外费用
'    End If
'    If 类别 = 1 Then
        '龚智毅 20051028
        'g结算数据.帐户支付 = dbl费用总额 - g结算数据.医保外费用 - g结算数据.基本医保支付 - g结算数据.高额医保支付 - g结算数据.公务员医疗补助 - g结算数据.生育支付
        g结算数据.帐户支付 = dbl费用总额 - g结算数据.医保外费用 - g结算数据.基本医保支付 - g结算数据.高额医保支付 - g结算数据.公务员医疗补助
'    Else
'        g结算数据.帐户支付 = dbl费用总额 - g结算数据.基本医保支付 - g结算数据.高额医保支付 - g结算数据.公务员医疗补助
'    End If
    'by 20050122 gzy
    If g结算数据.帐户可用余额 >= 0 Then
       If g结算数据.帐户可用余额 < g结算数据.帐户支付 Then
          g结算数据.帐户支付 = g结算数据.帐户可用余额
       End If
    Else
        g结算数据.帐户支付 = 0
    End If
    str结算方式 = "||个人帐户|" & g结算数据.帐户支付
    str虚拟结算返回串 = "个人帐户;" & g结算数据.帐户支付 & ";1"
    
    If g结算数据.基本医保支付 <> 0 Then
        str结算方式 = str结算方式 & "||基本统筹|" & g结算数据.基本医保支付
        str虚拟结算返回串 = str虚拟结算返回串 & "|基本统筹;" & g结算数据.基本医保支付 & ";0"
    End If
    If g结算数据.高额医保支付 <> 0 Then
        str结算方式 = str结算方式 & "||大病支付|" & g结算数据.高额医保支付
        str虚拟结算返回串 = str虚拟结算返回串 & "|大病支付;" & g结算数据.高额医保支付 & ";0"
    End If
    If g结算数据.公务员医疗补助 <> 0 Then
        str结算方式 = str结算方式 & "||公务员补助|" & g结算数据.公务员医疗补助
        str虚拟结算返回串 = str虚拟结算返回串 & "|公务员补助;" & g结算数据.公务员医疗补助 & ";0"
    End If
    '20051020 陈东
    '>Beging 生育支付
    g结算数据.生育盈亏 = 0
    If g结算数据.生育支付 <> 0 Then
        str结算方式 = str结算方式 & "||生育支付|" & g结算数据.生育支付
        str虚拟结算返回串 = str虚拟结算返回串 & "|生育支付;" & g结算数据.生育支付 & ";0"
    End If
    '>end 生育支付
    If Format(g病人身份_成都内江.费用总额, "###0.00;-###0.00;0;0") <> Format(dbl费用总额, "###0.00;-###0.00;0;0") Then
        Dim blnYes As Boolean
        '费用总额与医保中心返回总额不致,不能进行结算
        ShowMsgbox "本次结算总额(" & g病人身份_成都内江.费用总额 & ") 与" & vbCrLf & _
                    "   中心返回的总额(" & dbl费用总额 & ")不等，不能结算！"
        Exit Function
    End If
    If g结算数据.生育支付 > g病人身份_成都内江.生育总费用 Then
       ShowMsgbox "本次发生生育费用：(" & g病人身份_成都内江.生育总费用 & ") 小于" & vbCrLf & _
                    "   中心返回的生育支付：(" & g结算数据.生育支付 & ")不能结算！"
       Exit Function
    End If
    
    strAdvance = str虚拟结算返回串
    结算方式更正 = True
    Exit Function
    
    'Modified by ZYB 20051123 由于存在虚拟结算，所以不存在校正的问题了
    '如果存在,则保存冲预交记录中
'    If str结算方式 <> "" Then
'        str结算方式 = Mid(str结算方式, 3)
'        g病人身份_成都内江.结算方式 = str结算方式
'
'        If g结算数据.结算标志 = 0 Then
'            #If gverControl < 2 Then
'                gstrSQL = "zl_病人结算记录_Update(" & g结算数据.结帐ID & ",'" & str结算方式 & "', 0)"
'            #Else
'                strAdvance = str结算方式
'                gstrSQL = "zl_医保核对表_Insert(" & g结算数据.结帐ID & ",'" & str结算方式 & "')"
'            #End If
'            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
''        Else
''                gstrSQL = "zl_病人结算记录_Update(" & g结算数据.结帐ID & ",'" & str结算方式 & "',1)"
''                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
'        End If
'    End If
'
'    '显示结算信息
'    '"个人帐户:" & g结算数据.帐户支付
'    #If gverControl < 2 Then
'        If frm结算信息.ShowMe(g结算数据.结帐ID, False, , IIf(g结算数据.结算标志 = 0, 0, 1)) = False Then
'            结算方式更正 = False
'            Exit Function
'        End If
'    #End If
    结算方式更正 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function Get交易代码(ByVal intType As 业务类型_成都内江, Optional bln读名称 As Boolean = False) As String
    Select Case intType
        Case 读病人信息_内江
            Get交易代码 = IIf(bln读名称, "读病人信息", "01")
        Case 更改密码_内江
            Get交易代码 = IIf(bln读名称, "更改密码", "02")
        Case 获取帐户余额_内江
            Get交易代码 = IIf(bln读名称, "获取帐户余额", "03")
        Case 门诊明细写入_内江
            Get交易代码 = IIf(bln读名称, "门诊明细写入", "04")
        Case 门诊消费确认_内江
            Get交易代码 = IIf(bln读名称, "门诊消费确认", "05")
        Case 门诊消费取消_内江
            Get交易代码 = IIf(bln读名称, "门诊消费取消", "06")
        Case 住院登记_内江
            Get交易代码 = IIf(bln读名称, "住院登记", "07")
        Case 出院登记上传_内江
            Get交易代码 = IIf(bln读名称, "出院登记上传", "08")
        Case 住院交易上传_内江
            Get交易代码 = IIf(bln读名称, "住院交易上传", "09")
        Case 住院产易上传取消_内江
            Get交易代码 = IIf(bln读名称, "住院产易上传取消", "10")
        Case 出院登记确认_内江
            Get交易代码 = IIf(bln读名称, "出院登记确认_内江", "11")
        Case 获取单位欠缴情况_内江
            Get交易代码 = IIf(bln读名称, "获取单位欠缴情况", "12")
        Case 初始化函数_内江
            Get交易代码 = IIf(bln读名称, "初始化函数", "13")
        Case 网上对帐_内江
            Get交易代码 = IIf(bln读名称, "网上对帐", "14")
        Case 并发症申请上传_内江
            Get交易代码 = IIf(bln读名称, "并发症申请上传", "15")
        Case Else
            Get交易代码 = IIf(bln读名称, "错误的交易代码", "-1")
    End Select
End Function

Public Function 业务请求_成都内江(ByVal intType As 业务类型_成都内江, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strReturn As String
    Dim strOutput(0 To 20) As String, dblOutPut(0 To 25) As Double, intOutPut(0 To 5) As Integer, lngOutPut(0 To 5) As Long
    Dim strArr1
    Dim strArr(0 To 20) As String
    Dim str业务 As String
    Dim strReg As String
    
    Dim strNetInput As String
    Dim lng重试 As Long
    Dim i As Integer
    

    str业务 = Get交易代码(intType, True)
    
    DebugTool "进入业务请求函数(业务类型代码为:" & intType & " 业务名称：" & str业务 & ")" & vbCrLf & "        输入参数为:" & strInputString
    
    
    业务请求_成都内江 = False
    
    StrInput = strInputString
    
    If InitInfor_成都内江.模拟数据 Then
        '读取模拟数据
        Read模拟数据 intType, strInputString, strOutPutstring
         业务请求_成都内江 = True
        Exit Function
    End If
   
 'Modify 陈东 20051020
 'Beging 检查网络是否连通
    GetRegInFor g公共全局, "医保", "ConfigFileName", strReg
    strNetInput = strReg
    GetRegInFor g公共全局, "医保", "HostPort", strReg
    strNetInput = strNetInput & vbTab & strReg
    GetRegInFor g公共全局, "医保", "IPAddress", strReg
    strNetInput = strNetInput & vbTab & strReg
    
    strArr1 = Split(strNetInput, vbTab)
    For i = 0 To UBound(strArr1)
        strArr(i) = strArr1(i)
    Next
    
    lngReturn = gobj成都内江.SetCommPara(strArr(0), Val(strArr(1)), strArr(2))
    If lngReturn <> 1 Then
         '龚智毅 20051028
         'ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
         ShowMsgbox "医保网络不通"
         Exit Function
    End If
    
    lngReturn = 0
 'End检查网络是否连通
 
    strArr1 = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr1)
        strArr(i) = strArr1(i)
    Next
         
         
    For i = 0 To 20
        strOutput(i) = Space(100)
    Next
    
    Err = 0
    On Error GoTo errHand:
    
    Select Case intType
        Case 读病人信息_内江
            If InitInfor_成都内江.读卡器_内江 = 0 Then
                '输入参数:
                ''调试时要加上 gobj成都内江.
                lngReturn = GetCardInfo_MW(Val(strArr(0)), strArr(1), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), strOutput(5), strOutput(6), strOutput(7), strOutput(8), strOutput(9), strOutput(10), strOutput(11), strOutput(12))
            Else
                lngReturn = GetCardInfo_KRQ(Val(strArr(0)), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), strOutput(5), strOutput(6), strOutput(7), strOutput(8), strOutput(9), strOutput(10), strOutput(11), strOutput(12))
            End If
            For i = 0 To 20
                strOutput(i) = Trim(Split(strOutput(i), Chr(0))(0))
            Next
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(CStr(lngReturn))
                 Exit Function
            End If
           '构建返回串
           strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & strOutput(5) & vbTab & strOutput(6) & vbTab & strOutput(7) & vbTab & strOutput(8) & vbTab & strOutput(9) & vbTab & strOutput(10) & vbTab & strOutput(11) & vbTab & strOutput(12)
        Case 更改密码_内江
            lngReturn = ChangePassword(Val(strArr(0)), strArr(1), strArr(2))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(CStr(lngReturn))
                 Exit Function
            End If
        Case 获取帐户余额_内江
            '输入参数:  个人编号    String (8)  IN
            '           社保卡号码  String (10) IN
            '           统筹地区编码    String (1)  IN
            '输出参数:  帐户余额    Long    OUT
            '           在职情况    String(1)   OUT
            lngReturn = gobj成都内江.GetAccountAmountFunc(strArr(0), strArr(1), strArr(2), lngOutPut(0), strOutput(0))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = lngOutPut(0) / 100 & vbTab & strOutput(0)
        Case 门诊明细写入_内江
            '输入参数:  个人编号    String(8)   In
            '           社保卡号码  String(10)  In
            '           医院代码    String(5)   In
            '           操作员卡号码    String(10)  In
            '           统筹地区编码    String(1)   In
            '           医院交易流水号  String(20)  In
            '           交易类别    String(1)   In
            '           处方条数    String(2)   In
            '           处方明细    String处方条数×51  In

            '输出参数:  医保流水号  String(20)  Out
            '           医保内费用  String(10)  Out
            '           医保外费用  String(10)  Out
            '           个人帐户可用余额    String(10)  Out
            '           处方明细    String处方条数×51  Out
            '           在职情况    String(1)   Out
            '           '20051020 add
            '           生育支付    String(10) Out
            Do While lng重试 <= 3
                'lngReturn = gobj成都内江.DoConsumeTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strOutPut(0), strOutPut(1), strOutPut(2), strOutPut(3), strOutPut(4), strOutPut(5))
                lngReturn = gobj成都内江.DoConsumeTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), strOutput(5), strOutput(6))
                If lngReturn <> 0 Then
                    'If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "医保接口") = vbCancel Then
                    '    lng重试 = 8
                   '     Exit Function
                   ' End If
                   lng重试 = lng重试 + 1
                   If lng重试 > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "医保接口"
                        Exit Function
                   End If
                Else
                    lng重试 = 5
                End If
            Loop
            'strReturn = strOutPut(0) & vbTab & Val(strOutPut(1)) / 100 & vbTab & Val(strOutPut(2)) / 100 & vbTab & Val(strOutPut(3)) / 100 & vbTab & strOutPut(4) & vbTab & strOutPut(5)
            strReturn = strOutput(0) & vbTab & Val(strOutput(1)) / 100 & vbTab & Val(strOutput(2)) / 100 & vbTab & Val(strOutput(3)) / 100 & vbTab & strOutput(4) & vbTab & strOutput(5) & vbTab & strOutput(6) / 100
            
        Case 门诊消费确认_内江
            '输入参数: 个人编号    String(8)   In
            '          社保卡号码  String(10)  In
            '          医院代码    String(5)   In
            '          操作员卡号码    String(10)  In
            '          统筹地区编码    String(1)   In
            '          医保交易流水号  String(20)  In
            '          交易类别    String(1)   In
            '          个人帐户支付    String(10)  In
            Do While lng重试 <= 3
                lngReturn = gobj成都内江.DoConsumeAffirmFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "医保接口") = vbCancel Then
'                        lng重试 = 8
'                        Exit Function
'                    End If
                   lng重试 = lng重试 + 1
                   If lng重试 > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "医保接口"
                        Exit Function
                   End If
                Else
                    lng重试 = 5
                End If
            Loop
            strReturn = ""
        Case 门诊消费取消_内江
            '输入参数: 个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        医院代码    String(5)   In
            '        操作员卡号码    String(10)  In
            '        统筹地区编码    String(1)   In
            '        医保交易流水号  String(20)  In
            '        交易类别    String(1)   In
            Do While lng重试 <= 3
                lngReturn = gobj成都内江.DoConsumeCancelFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "医保接口") = vbCancel Then
'                        lng重试 = 8
'                        Exit Function
'                    End If
                   lng重试 = lng重试 + 1
                   If lng重试 > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "医保接口"
                        Exit Function
                   End If

                Else
                    lng重试 = 5
                End If
            Loop
            strReturn = ""
        Case 住院登记_内江
            '输入参数: 个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        医院代码    String(5)   In
            '        操作员卡号码    String(10)  In
            '        统筹地区编码    String(1)   In
            '        入院日期    String(8)   In
            '        入院科别    String(10)  In
            '        入院诊治医生    String(10)  In
            '        诊断编码    String(20)  In
            
            '输出参数:住院流水号  String(20)  Out
            '        享受待遇标志    Small int   Out
            '        起付标准    Long    Out
            '        在职情况    String(1)   Out

            Do While lng重试 <= 3
                lngReturn = gobj成都内江.DoHospInFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strOutput(0), intOutPut(0), lngOutPut(0), strOutput(1))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "医保接口") = vbCancel Then
'                        lng重试 = 8
'                        Exit Function
'                    End If
                   lng重试 = lng重试 + 1
                   If lng重试 > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "医保接口"
                        Exit Function
                   End If

                Else
                    lng重试 = 5
                End If
            Loop
            strReturn = strOutput(0) & vbTab & intOutPut(0) & vbTab & lngOutPut(0) & vbTab & strOutput(1)
        Case 住院交易上传_内江
            '输入参数: 个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        医院代码    String(5)   In
            '        统筹地区编码    String(1)   In
            '        医院交易流水号  String(20)  In
            '        出院带药类别    String(1)   In
            '        科别    String(10)  In
            '        医生    String(10)  In
            '        住院流水号  String(20)  In
            '        处方条数    String(2)   In
            '        处方明细    String处方条数×51  In
            '输出参数:  医保交易流水号  String(20)  Out
            '           处方明细    String处方条数×51  Out
            '           TRANSDETIAL输出 (计算费用明细) Out
            Do While lng重试 <= 3
                lngReturn = gobj成都内江.DoHospTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), strArr(10), strOutput(0), strOutput(1), strOutput(2))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "医保接口") = vbCancel Then
'                        lng重试 = 8
'                        Exit Function
'                    End If
                   lng重试 = lng重试 + 1
                   If lng重试 > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "医保接口"
                        Exit Function
                   End If

                Else
                    lng重试 = 5
                End If
            Loop
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2)
        Case 住院产易上传取消_内江
            '输入参数:个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        操作员卡号码    String(10)  In
            '        统筹地区编码    String(1)   In
            '        住院流水号  String(20)  In
            '        医保交易流水号  String(20)  In
            '输出参数:

'            lngReturn = gobj成都内江.DoHospCancelFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
'            If lngReturn <> 0 Then
'                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
'                 Exit Function
'            End If
            Do While lng重试 <= 3
                lngReturn = gobj成都内江.DoHospCancelFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "医保接口") = vbCancel Then
'                        lng重试 = 8
'                        Exit Function
'                    End If
                   lng重试 = lng重试 + 1
                   If lng重试 > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "医保接口"
                        Exit Function
                   End If

                Else
                    lng重试 = 5
                End If
            Loop
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2)
        Case 出院登记上传_内江
            '输入参数:个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        医院代码    String(5)   In
            '        操作员卡号码    String(10)  In
            '        统筹地区编码    String(1)   In
            '        出院日期    String(8)   In
            '        出院科别    String(10)  In
            '        出院诊治医生    String(10)  In
            '        诊断编码    String(20)  In
            '        出院带药    String(1)   In
            '        出院类别    String(1)   In
            '        住院流水号  String(20)  In
            '输出参数
            '        TRANSDETIAL输出 (计算费用明细)
            '        享受待遇标志    String(1)   Out
            '        医保内费用  String(10)  Out
            '        医保外费用  String(10)  Out
            '        基本医保支付 如果参加大病医保，则为大病医保支付  String(10)  Out
            '        高额医保支付    String(10)  Out
            '        公务员医疗补助  String(10)  Out
            '        个人按比例支付  String(10)  Out
            '        TRANSDETIAL结束
            '        起付标准    String(10)  Out
            '        个人帐户可用余额    String(10)  Out
'            lngReturn = gobj成都内江.DoHospOutTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), strArr(10), strArr(11), strOutPut(0), strOutPut(1), strOutPut(2))
'            If lngReturn <> 0 Then
'                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
'                 Exit Function
'            End If
            Do While lng重试 <= 3
                lngReturn = gobj成都内江.DoHospOutTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), strArr(10), strArr(11), strOutput(0), strOutput(1), strOutput(2))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "医保接口") = vbCancel Then
'                        lng重试 = 8
'                        Exit Function
'                    End If
                   lng重试 = lng重试 + 1
                   If lng重试 > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "医保接口"
                        Exit Function
                   End If

                Else
                    lng重试 = 5
                End If
            Loop
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2)
        Case 出院登记确认_内江
        
            '输入参数:个人编号    String(8)   In
            '    社保卡号码  String(10)  In
            '    操作员卡号码    String(10)  In
            '    统筹地区编号    String(1)   In
            '    住院流水号  String(20)  In
            '    个人帐户支付    String(10)  In
'            lngReturn = gobj成都内江.DoHospOutAffirmFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
'            If lngReturn <> 0 Then
'                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
'                 Exit Function
'            End If
            Do While lng重试 <= 3
                lngReturn = gobj成都内江.DoHospOutAffirmFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "医保接口") = vbCancel Then
'                        lng重试 = 8
'                        Exit Function
'                    End If
                   lng重试 = lng重试 + 1
                   If lng重试 > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "医保接口"
                        Exit Function
                   End If

                Else
                    lng重试 = 5
                End If
            Loop
            strReturn = ""
        
        Case 获取单位欠缴情况_内江
            '输入参数:个人编号    String (8)  IN
            '        社保卡号码  String (10) IN
            '        统筹地区编码    String (1)  IN
            '输出参数
            '        单位欠缴情况    String(1)   OUT
            lngReturn = gobj成都内江.GetArrearInfo(strArr(0), strArr(1), strArr(2), strOutput(0))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If

            strReturn = strOutput(0)
        Case 初始化函数_内江
            '输入参数:ConfigFileName
            '        HostPort
            '        IPAddress
            lngReturn = gobj成都内江.SetCommPara(strArr(0), Val(strArr(1)), strArr(2))
            If lngReturn <> 1 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = ""
        'Add 陈东 20051020
        Case 网上对帐_内江
        '输入参数内容
            'HOSPID  医院/药店编号
            'TCDQBM  统筹地区编码
            'DZLB    对帐类别(0:门诊,1:住院, 2药店)
            'KSRQ    对帐开始日期
            'ZZRQ    对帐终止日期
            'Count   上传数据条数
            'je      上传总额
        '输出参数
            'DZQK    对帐情况(0成功,1 金额相等 , 条数不等 2 金额不等 , 条数相等 3金额不等,条数不等)
            'DZCOUNT 对帐数据条数
            'DZJE    对帐金额
            lngReturn = gobj成都内江.CompareTotal(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strOutput(0), strOutput(1), strOutput(2))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2)
        Case 并发症申请上传_内江
        'TCDQBM        '统筹地区编码    String (1)  IN
        'ZYLSH        '住院流水号  String (20) IN
        'ZDBM        '诊断编码    String (200)    IN
            lngReturn = gobj成都内江.DoBFZAffirmFunc(strArr(0), strArr(1), strArr(2))
            If lngReturn <> 1 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = ""
    End Select
    strOutPutstring = strReturn
    业务请求_成都内江 = True
    DebugTool "  输出参数为:" & strReturn
     Exit Function
errHand:
    DebugTool "业务请求失败  输出参数为:" & strReturn
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetItemInsure_成都内江(lng病人ID As Long, lng收费细目ID As Long, bln门诊 As Boolean) As String
    Dim strDefault As String            '缺省医保编码
    Dim strCurrent As String            '当前医保编码，门诊取门诊编码，住院取住院编码
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "Select B.类别,A.编码,A.名称,B.说明 From 保险项目 A,医保对照明细 B" & _
        " Where B.险类=[1] And A.险类=B.险类 And A.编码=B.项目编码 And B.收费细目ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保编码", TYPE_成都内江, lng收费细目ID)
    rsTemp.Filter = "类别=" & IIf(bln门诊, 1, 2)
    Select Case rsTemp.RecordCount
    Case 0
        '没有设置对应编码，取缺省编码
        rsTemp.Filter = "类别=0"
        If rsTemp.RecordCount <> 0 Then
            GetItemInsure_成都内江 = rsTemp!编码
        End If
    Case 1
        GetItemInsure_成都内江 = rsTemp!编码
    Case Else
        '多选
        GetItemInsure_成都内江 = frm医保项目选择.ShowSelect(rsTemp, lng收费细目ID)
    End Select
    
    rsTemp.Filter = 0
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    rsTemp.Filter = 0
End Function

Public Function 并发症选择_成都内江(lng病人ID As Long, lng主页ID As Long)
    '功能 上传并发证
    '20051024 陈东
    Dim rsBfz As New ADODB.Recordset
    gstrSQL = "Select * from 病案主页 Where 出院日期 is null and 病人ID=[1] And 主页ID=[2]"
    Set rsBfz = zlDatabase.OpenSQLRecord(gstrSQL, "是否在院", lng病人ID, lng主页ID)
    If rsBfz.RecordCount > 0 Then
        gstrSQL = "Select * from 保险帐户 Where 病人ID=[1]"
        Set rsBfz = zlDatabase.OpenSQLRecord(gstrSQL, "取统筹地区编码", lng病人ID)
    Else
        MsgBox "出院病人不能执行此操作!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If 医保初始化_成都内江 = False Then Exit Function
    
    If 存在未结费用(lng病人ID, lng主页ID) = True Then
        frm病种选择_成都内江.GetCode (lng病人ID)
    Else
        MsgBox "费用已结清，不能执行此操作！", vbInformation, gstrSysName
        Exit Function
    End If
    
End Function

Public Function 判断补办效期_成都内江(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim int有效天数 As Integer
    Dim dbl相差天数 As Double

    gstrSQL = "Select nvl(参数值,0) as 参数值 From 保险参数 where 险类=" & TYPE_成都内江 & " and 参数名='允许补办天数'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取允许补办天数"
    If rsTemp.EOF Then
       判断补办效期_成都内江 = False
       MsgBox "没有发现参数：允许补充天数。", vbInformation, gstrSysName
       Exit Function
    End If
    int有效天数 = rsTemp!参数值
    
    gstrSQL = "Select trunc(sysdate-入院日期,2) as 相差天数 From 病案主页 where 病人id =" & lng病人ID & " and 主页id =" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取相差天数"
    If rsTemp.EOF Then
       判断补办效期_成都内江 = False
       MsgBox "无法获取相差天数。", vbInformation, gstrSysName
       Exit Function
    End If
    dbl相差天数 = rsTemp!相差天数
    
    If int有效天数 < dbl相差天数 And int有效天数 > 0 Then
       判断补办效期_成都内江 = False
       MsgBox "已经超过了允许的补办天数。", vbInformation, gstrSysName
       Exit Function
    End If
    判断补办效期_成都内江 = True
End Function
