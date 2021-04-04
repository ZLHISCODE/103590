Attribute VB_Name = "mdl成都德阳"
Option Explicit
Private mblnInit As Boolean

Private gobjTest As Object

Public Enum 业务类型_成都德阳
    获得社保机构 = 0
    获得参保人员资料
    入院登记
    取消入院登记
    出院登记
    取消出院登记
    增加处方单据
    增加处方明细
    删除处方单据及其明细
    单条处方传输
    出院结算
    取消出院结算
    
    打印出院结算报表函数
    打印住院人员财务结算单
    重提人员基本资料
    提取基础资料
    获取住院记录号
    检查拔号连接
    建立拔号连接
    断开拔号连接
    获取药品信息
    跨月重提人员基本资料
End Enum
Private Type InitbaseInfor
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
    机构编码 As String                      '默认的社保机构编码
    明细时实上传 As Boolean
    数据不等不可结算 As Boolean
    打印结算单      As Boolean
    自动出院审核    As Boolean
End Type
Public InitInfor_成都德阳 As InitbaseInfor

Private Type 病人身份
        记录号        As String
        保障号    As String       '即医保号
        姓名     As String
        性别     As String
        出生日期  As String
        年龄        As Integer
        医疗性质    As String
        退休管理    As String
        单位编码    As String
        单位名称    As String
        医疗标志    As String
        机构编码    As String
        
        费用总额    As Double
        病人ID      As Long
        病种编码    As String
        病种名称    As String
        中心        As Long
End Type
Private Type 结算数据
    医保基金 As Double
    补保陪付额 As Double
End Type
Private gbln未达起付线 As Boolean

Private g虚拟结算 As 结算数据
Public g病人身份_成都德阳 As 病人身份
Public gcnOracle_成都德阳 As ADODB.Connection     '中间库连接
Private gbln检查连接 As Boolean
Private gbln已经初始 As Boolean             '已经被初始化了.
'1.获得社保机构编号和名称列表
Private Declare Function GetSBJGLB Lib "CDGK_YB.dll" () As String
'===============================================================================================================
'原型:FUNCTION GETSBJGLB:PCHAR
'功能: 获得社保机构编号和名称列表
'入口参数: 无
'出口参数: 无
'返回: A社保机构编号+列间隔符+A社保机构名称+行间隔符+B社保机构编号+列间隔符+B社保机构名称+……
'===============================================================================================================

'2．获得参保人员的基本资料
Private Declare Function GETRYJBZL Lib "CDGK_YB.dll" Alias "GetRYJBZL" (ByVal str保障号 As String, ByVal str社保编号 As String) As String
'===============================================================================================================
'原型:FUNCTION GETRYJBZL(ASBBH,ABXJGBH:PCHAR):PCHAR;
'功能: 获得参保人员的基本资料
'入口参数: ASBBH   PCHAR   参保人员的社会保障号
'          ABXJGBH PCHAR   参保人员所在的保险机构编号
'出口参数: 无
'返回: A社保机构编号+列间隔符+A社保机构名称+行间隔符+B社保机构编号+列间隔符+B社保机构名称+……
'===============================================================================================================

'3.入院登记
Private Declare Function RYDJ Lib "CDGK_YB.dll" (ByVal str住院号 As String, ByVal str个人资料 As String, ByVal str机构编号 As String) As String
'===============================================================================================================
'原型:FUNCTION RYDJ(AZYH,;ARYZL,ABXJGBH:PCHAR):PCHAR;
'功能: 在社保机构医保数据库和医院本地医保数据库中对住院的医保病人进行登记。
'入口参数: str住院号   PCHAR   住院号
'          str个人资料 PCHAR   参保人员的个人资料
'          str机构编号 PCHAR 参保人员所在的社保机构编号
'出口参数: 无
'返回:返回标志@$社会保障号||个人记录号||医疗性质||退休管理||单位名称||姓名||性别||出生日期（格式：YYYY-MM-DD）||单位编号||参加基本医疗标志||入院日期（格式：YYYY-MM-DD）||病种编号||病种名称||科室
'===============================================================================================================

'4.取消住院
Private Declare Function ZYQX Lib "CDGK_YB.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION ZYQX(AZYH:PCHAR):PCHAR
'功能: 在社保机构医保数据库和医院本地医保数据库中删除医保病人住院记录。
'入口参数: str住院号   PCHAR   住院号
'出口参数: 无
'返回:返回标志
'===============================================================================================================

'5.出院登记
Private Declare Function CYCS Lib "CDGK_YB.dll" (ByVal str住院号 As String, ByVal str出院日期 As String) As String
'===============================================================================================================
'原型:FUNCTION CYCS(AZYH ,CYRQ:PCHAR):PCHAR;
'功能: 将医保病人住院过程中所有数据上传至社保机构医保数据库；对本地医保数据库中医保病人作出院处理。
'入口参数: str住院号   PCHAR   住院号
'          str出院日期 pchar 出院日期（YYYY-MM-DD）
'出口参数: 无
'返回:返回标志
'===============================================================================================================

'6.取消出院登记
Private Declare Function CYCSQX Lib "CDGK_YB.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION CYCSQX (AZYH:PCHAR):PCHAR;
'功能:取消参保病人向社保局已经传输的出院数据，以便重新传输。
'入口参数: str住院号   PCHAR   住院号
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================


'7.增加一个处方单据
Private Declare Function AddCFJL Lib "CDGK_YB.dll" (ByVal str住院号 As String, ByVal str处方日期 As String, ByVal str医生 As String, ByVal str科室 As String) As String
'===============================================================================================================
'原型:FUNCTION ADDCFJL(AZYH,ACFRQ,AYS,AKS:PCHAR):PCHAR
'功能:增加一个处方单据。。
'入口参数:
'        AZYH    PCHAR   住院号
'        ACFRQ   PCHAR   处方日期（YYYY-MM-DD）
'        AYS PCHAR   医生
'        AKS PCHAR   科室
'出口参数: 无
'返回:'OK'+行间隔符+处方记录号或错误信息
'===============================================================================================================

'7.增加处方明细
Private Declare Function AddCFMX Lib "CDGK_YB.dll" (ByVal str处方记录号 As String, ByVal str医保编码 As String, ByVal str数量 As String, ByVal str单价 As String) As String
'===============================================================================================================
'原型:FUNCTION ADDCFMX(ACFID,AYPBH,ASL,ADJ:PCHAR):PCHAR;
'功能:增加一个处方明细。
'入口参数:
'    ACFID   PCHAR   处方记录号
'    AYPBH   PCHAR   药品编号(社保药品编号)
'    ASL PCHAR   数量(可以为负数)
'    ADJ PCHAR   单价
'出口参数: 无
'返回:'OK'+行间隔符+处方明细记录号+行间隔符+自费比例+行间隔符+自费金额或错误信息
'===============================================================================================================

'8.删除处方单据及其明细
Private Declare Function DELCFJL Lib "CDGK_YB.dll" (ByVal str处方记录号 As String) As String
'===============================================================================================================
'原型:FUNCTION DELCFJL(ACFID:PCHAR):PCHAR
'功能:删除处方单据及其下属的明细记录。
'入口参数:
'    ACFID   PCHAR   处方记录号
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================


'9.单条处方传输
Private Declare Function CFCS Lib "CDGK_YB.dll" (ByVal str住院号 As String, ByVal str处方记录号 As String) As String
'===============================================================================================================
'原型:FUNCTION CFCS(AZYH:PCHAR;ACFID:PCHAR):PCHAR
'功能:将社保病人每天的处方情况向社保局中心数据库传输（同一个处方可以多次重复传输，后一次传输的数据将覆盖前一次传输的数据）
'入口参数:
'    AZYH    PCHAR   住院号
'    ACFID   PCHAR   处方记录号（通过调用ADDCFJL返回的值）
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================

'10.出院结算
Private Declare Function CYJS Lib "CDGK_YB.dll" (ByVal str住院号 As String, ByVal lng预结标志 As Long) As String
'===============================================================================================================
'原型:FNCTION CYJS(AZYH:PCHAR; ISPREV:INTEGER):PCHAR
'功能:住院参保病人出院或住院中预结算
'入口参数:
'    AZYH    PCHAR   住院号
'    ISPREV  PCHAR   预结算标志（'0'－表示预结算）
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================

'11.取消出院结算
Private Declare Function CYJSQX Lib "CDGK_YB.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION CYJSQX(AZYH:PCHAR):PCHAR
'功能:取消参保病人出院结算
'入口参数:
'    AZYH    PCHAR   住院号
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================

'12.打印出院结算报表函数
Private Declare Function JSReport Lib "CDGK_YB.dll" (ByVal str开始住院号 As String, ByVal str结束住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION JSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL
'功能:打印社保机构提供的动态报表，目前德阳地区所用动态报表："住院结算统计表（补充）"、"个人住院结算单"和"住院结算统计表"三张报表。使用"21．提取基础资料"函数，自动更新本地报表。
'入口参数:
'    ASTARTZYH   PCHAR   打印开始住院号
'    AENDZYH PCHAR      打印结束住院号
'   注意:
'    1 ?二个住院号之间所有的住院记录必须为同一个社保局?
'    2、当只打印一个住院号的报表时，二个参数值一样。
'出口参数: 无
'返回:无须注意返回值
'===============================================================================================================

'13.打印住院人员财务结算单
Private Declare Function CWJSREPORT Lib "CDGK_YB.dll" Alias "CWJSReport" (ByVal str开始住院号 As String, ByVal str结束住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION CWJSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL;
'功能:打印住院人员财务结算单。
'入口参数:
'    ASTARTZYH   PCHAR   打印开始住院号
'    AENDZYH PCHAR      打印结束住院号
'   注意:
'    1 ?二个住院号之间所有的住院记录必须为同一个社保局?
'    2、当只打印一个住院号的报表时，二个参数值一样。
'出口参数: 无
'返回:无须注意返回值
'===============================================================================================================

'14.提取基础资料
Private Declare Function GETJCXX Lib "CDGK_YB.dll" Alias "GetJCXX" (ByVal str机构编号 As String, ByVal str下载标志 As String) As String
'===============================================================================================================
'原型:GETJCXX(SBXJGBH:PCHAR;DOWNALL:INTEGER):PCHAR
'功能:向指定的社保机构提取基础资料
'入口参数:
'    SBXJGBH PCHAR   保险机构编号
'    DOWNALL PCHAR   当值为0时表示下载本地医保数据库中没有的基础资料，为其他时表示全部重新下载
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================

'15 根据住院号得到住院记录号
Private Declare Function GETZYIDBYZYBH Lib "CDGK_YB.dll" Alias "GetZYIDByZyBH" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION GETZYIDBYZYBH(AZYH:PCHAR):PCHAR
'功能:根据住院号得到住院记录号
'入口参数:
'   AZYH    PCHAR   住院号'出口参数: 无
'返回:'OK'@$住院记录号或错误信息
'===============================================================================================================

'16.检测拔号连接是否连接成功
Private Declare Function CheckCon Lib "CDGK_YB.dll" () As String
'===============================================================================================================
'原型:FUNCTION CHECKCON:PCHAR;
'功能:检测拔号连接是否连接成功
'入口参数:
'返回:OK或错误信息
'===============================================================================================================

'17.建立拔号连接
Private Declare Function RasDial Lib "CDGK_YB.dll" (ByVal str机构代码 As String) As String
'===============================================================================================================
'原型:FUNCTION RASDIAL(SBXJGBH:PCHAR):PCHAR
'功能:拔号至选择的社保局，与其建立连接
'入口参数:SBXJGBH PCHAR   保险机构编号
'返回:  成功    川大金键HIS拔号器状态栏显示"连接"
'       失败 错误信息
'===============================================================================================================

'18.断开与社保局的连接
Private Declare Function DisDial Lib "CDGK_YB.dll" () As String
'===============================================================================================================
'原型:FUNCTION DISDIAL:PCHAR
'功能:拔号至选择的社保局，与其建立连接
'入口参数:
'返回:
'===============================================================================================================


'19 根据药品编号得到药品信息
Private Declare Function GetSINYPXX Lib "CDGK_YB.dll" (ByVal str机构编码 As String, ByVal str药品编码 As String) As String
'===============================================================================================================
'原型:FUNCTION GETSINYPXX(SBXJGBH,CYPBH:PCHAR):PCHAR
'功能:根据药品编号得到药品信息
'入口参数:
'    SBXJGBH PCHAR   保险机构编号
'    CYPBH   PCHAR   药品编号
'返回:OK@$类别:药品||中文名称:阿莫西林钠（克拉维酸钾）||计量单位:支||单价上线:0||自费比例:20
'===============================================================================================================

'20.住院病人跨月重提人员基本资料
Private Declare Function GETNEWRYZL Lib "CDGK_YB.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION GETNEWRYZL(AZYH:PCHAR):PCHAR;STDCALL;
'功能:住院病人跨月重提人员基本资料。
'入口参数:
'    AZYH    PCHAR   住院号
'返回:OK@$错误信息
'===============================================================================================================






Public Function 医保初始化_成都德阳() As Boolean
    Dim strReg As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If mblnInit Then
        医保初始化_成都德阳 = True
        Exit Function
    End If
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_成都德阳.模拟数据 = True
    Else
        InitInfor_成都德阳.模拟数据 = False
    End If
    
   Call GetRegInFor(g公共模块, "医保", "社保机构代码", strReg)
   
   InitInfor_成都德阳.机构编码 = strReg
   If strReg = "" Then
        MsgBox "你未设置默认的社保机构编码，请检查参数设置!"
        Exit Function
   End If
   
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_成都德阳)
    InitInfor_成都德阳.医院编码 = Nvl(rsTemp!医院编码)
    
    
    '中间库连接
    gstrSQL = "select 参数名,参数值 from 保险参数 where  险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "成都德阳医保", TYPE_成都德阳)
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "检查拨号连接"
                 gbln检查连接 = Nvl(rsTemp("参数值"), 0) = 1
            Case "明细时实上传"
                InitInfor_成都德阳.明细时实上传 = Nvl(rsTemp("参数值"), 0) = 1
            Case "比较结算数据"
                 InitInfor_成都德阳.数据不等不可结算 = IIf(Nvl(rsTemp!参数值, 1) = 1, 1, 0)
            Case "打印结算单"
                 InitInfor_成都德阳.打印结算单 = IIf(Nvl(rsTemp!参数值, 1) = 1, 1, 0)
            Case "自动出院审核"
                 InitInfor_成都德阳.自动出院审核 = IIf(Nvl(rsTemp!参数值, 1) = 1, 1, 0)
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_成都德阳 = New ADODB.Connection

    If OraDataOpen(gcnOracle_成都德阳, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    
   '建立拔号连接
   If gbln已经初始 = False And gbln检查连接 Then
        If 业务请求_成都德阳(建立拔号连接, InitInfor_成都德阳.机构编码, strOutput) = False Then
             Exit Function
        End If
   End If
   
   If gbln检查连接 Then
        '检查拔号连接
        If 业务请求_成都德阳(检查拔号连接, "", strOutput) = False Then
             Exit Function
        End If
    End If
    mblnInit = True
    gbln已经初始 = True
    医保初始化_成都德阳 = True
End Function

Public Function 医保终止_成都德阳() As Boolean
    Dim strOutput As String
    mblnInit = False
    If gcnOracle_成都德阳.State = 1 Then
        gcnOracle_成都德阳.Close
    End If
    '建立拔号连接
   Call 业务请求_成都德阳(断开拔号连接, "", strOutput)
    Err = 0
    On Error Resume Next
    医保终止_成都德阳 = True
End Function

Public Function 身份标识_成都德阳(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Err = 0
    On Error GoTo errHand:
    If bytType = 0 Or bytType = 3 Then Exit Function
    
    身份标识_成都德阳 = frmIdentify成都德阳.GetPatient(bytType, lng病人ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_成都德阳 = ""
End Function


Public Function 个人余额_成都德阳(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_成都德阳)
    
    If rsTemp.EOF Then
        个人余额_成都德阳 = 0
    Else
        个人余额_成都德阳 = rsTemp("帐户余额")
    End If
End Function
Public Function 门诊虚拟结算_成都德阳(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    门诊虚拟结算_成都德阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_成都德阳(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    门诊结算_成都德阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_成都德阳(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    门诊结算冲销_成都德阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_成都德阳(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
    Dim strArr
    Err = 0: On Error GoTo errHand:
    If InitInfor_成都德阳.机构编码 <> g病人身份_成都德阳.机构编码 Then
        '建立拔号连接
        If gbln已经初始 = False And gbln检查连接 Then
             If 业务请求_成都德阳(建立拔号连接, g病人身份_成都德阳.机构编码, strOutput) = False Then
                  Exit Function
             End If
        End If
        
        If gbln检查连接 Then
             '检查拔号连接
             If 业务请求_成都德阳(检查拔号连接, "", strOutput) = False Then
                  Exit Function
             End If
         End If
    End If
    
'    '获取住院号
'    gstrSQL = "Select 住院号 From 病人信息 where 病人id=" & lng病人id
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取住院号"
'
    Dim str住院号 As String
    str住院号 = Get住院号(lng病人ID, lng主页ID)
    If str住院号 = "" Then Exit Function
    
    '住院号||个人资料||社保机构编号
    StrInput = str住院号
    StrInput = StrInput & "||" & Get个人资料(lng病人ID, lng主页ID)
    StrInput = StrInput & "||" & g病人身份_成都德阳.机构编码
    If 业务请求_成都德阳(入院登记, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    '社会保障号||个人记录号||医疗性质||退休管理||单位名称||姓名||性别||出生日期（格式：YYYY-MM-DD）||单位编号||参加基本医疗标志||入院日期（格式：YYYY-MM-DD）||病种编号||病种名称||科室
    strArr = Split(strOutput, "||")
    '保存相关的信息
    ''OK'+行间隔符+社保中心住院记录号
    
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_成都德阳 & ",'医保住院号','''" & str住院号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "医保住院号")
    
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_成都德阳 & ",'住院记录号','''" & Val(strArr(0)) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保险住院记录号")
'    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_成都德阳 & ",'病种名称','''" & strArr(12) & "''')"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "病种名称")
    
    '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_成都德阳 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    入院登记_成都德阳 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_成都德阳 = False
End Function
Private Function Get个人资料(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    '    社会保障号|个人记录号|医疗性质|退休管理|单位名称|姓名|性别|出生日期（格式：YYYY-MM-DD）
    '    单位编号|参加基本医疗标志|入院日期（格式：YYYY-MM-DD）|病种编号|病种名称|科室
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String
    gstrSQL = "" & _
        "   Select  to_char(a.入院日期,'yyyy-mm-dd') as 入院日期,b.名称 as 科室" & _
        "   From 病案主页 a,部门表 b " & _
        "   Where A.入院科室ID=b.id(+) and a.病人id=" & lng病人ID & " and a.主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取主页信息"
    With g病人身份_成都德阳
        StrInput = .保障号
        StrInput = StrInput & vbTab & "|" & .记录号
        StrInput = StrInput & vbTab & "|" & .医疗性质
        StrInput = StrInput & vbTab & "|" & .退休管理
        StrInput = StrInput & vbTab & "|" & .单位名称
        StrInput = StrInput & vbTab & "|" & .姓名
        StrInput = StrInput & vbTab & "|" & .性别
        StrInput = StrInput & vbTab & "|" & .出生日期
        StrInput = StrInput & vbTab & "|" & .单位编码
        StrInput = StrInput & vbTab & "|" & .医疗标志
        StrInput = StrInput & vbTab & "|" & Nvl(rsTemp!入院日期)
        StrInput = StrInput & vbTab & "|" & .病种编码
        StrInput = StrInput & vbTab & "|" & .病种名称
        StrInput = StrInput & vbTab & "|" & Nvl(rsTemp!科室)
    End With
    Get个人资料 = StrInput
    
    
End Function
Private Function Get交易代码(ByVal intType As 业务类型_成都德阳, Optional bln读名称 As Boolean = False) As String
    '代码暂没用
    Select Case intType
        Case 获得社保机构
            Get交易代码 = IIf(bln读名称, "获得社保机构", "01")
        Case 获得参保人员资料
            Get交易代码 = IIf(bln读名称, "获得参保人员资料", "02")
        Case 入院登记
            Get交易代码 = IIf(bln读名称, "入院登记", "03")
        Case 取消入院登记
            Get交易代码 = IIf(bln读名称, "取消入院登记", "04")
        Case 出院登记
            Get交易代码 = IIf(bln读名称, "出院登记", "05")
        Case 取消出院登记
            Get交易代码 = IIf(bln读名称, "取消出院登记", "06")
        Case 增加处方单据
            Get交易代码 = IIf(bln读名称, "增加处方单据", "07")
        Case 增加处方明细
            Get交易代码 = IIf(bln读名称, "增加处方明细", "08")
        Case 删除处方单据及其明细
            Get交易代码 = IIf(bln读名称, "删除处方单据及其明细", "09")
        Case 单条处方传输
            Get交易代码 = IIf(bln读名称, "单条处方传输", "10")
        Case 出院结算
            Get交易代码 = IIf(bln读名称, "出院结算", "11")
        Case 取消出院结算
            Get交易代码 = IIf(bln读名称, "取消出院结算", "12")
        Case 打印出院结算报表函数
            Get交易代码 = IIf(bln读名称, "打印出院结算报表函数", "13")
        Case 打印住院人员财务结算单
            Get交易代码 = IIf(bln读名称, "打印住院人员财务结算单", "14")
        Case 重提人员基本资料
            Get交易代码 = IIf(bln读名称, "重提人员基本资料", "15")
        Case 提取基础资料
            Get交易代码 = IIf(bln读名称, "提取基础资料", "16")
        Case 获取住院记录号
            Get交易代码 = IIf(bln读名称, "获取住院记录号", "17")
        Case 检查拔号连接
            Get交易代码 = IIf(bln读名称, "检查拔号连接", "18")
        Case 建立拔号连接
            Get交易代码 = IIf(bln读名称, "建立拔号连接", "19")
        Case 断开拔号连接
            Get交易代码 = IIf(bln读名称, "断开拔号连接", "20")
        Case 获取药品信息
            Get交易代码 = IIf(bln读名称, "获取药品信息", "21")
        Case 跨月重提人员基本资料
            Get交易代码 = IIf(bln读名称, "跨月重提人员基本资料", "22")
        Case Else
            Get交易代码 = IIf(bln读名称, "错误的交易代码", "-1")
    End Select
End Function
Public Function 业务请求_成都德阳(ByVal intType As 业务类型_成都德阳, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strOutput As String, strReturn As String
    Dim strInValue(0 To 20) As String
    
    Dim str交易代码 As String
    Dim i As Integer
    Dim strArr
    
    str交易代码 = Get交易代码(intType)
    StrInput = str交易代码 & "|" & strInputString
    DebugTool "进入业务请求函数(业务类型为:" & intType & "),输入参数为" & vbCrLf & str交易代码 & "|" & StrInput
    
    业务请求_成都德阳 = False
    If InitInfor_成都德阳.模拟数据 Then
        '读取模拟数据
        Read模拟数据 intType, StrInput, strOutPutstring
         业务请求_成都德阳 = True
        Exit Function
    End If
    strArr = Split(strInputString, "||")
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    Err = 0
    On Error GoTo errHand:
'
'    If gobjTest Is Nothing Then
'        Set gobjTest = CreateObject("cdgk_Yb.clscdgk_Yb")
'    End If
    
    Select Case intType
        Case 获得社保机构
            strOutput = GetSBJGLB()
            
            If strOutput = "" Then
                MsgBox "获取社保机构时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Mid(strOutput, 5)
        Case 获得参保人员资料
            strOutput = GETRYJBZL(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "获得参保人员资料时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case 入院登记
            '
            strOutput = RYDJ(strInValue(0), Replace(strInValue(1), vbTab & "|", "||"), strInValue(2))
            If strOutput = "" Then
                MsgBox "入院登记时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case 取消入院登记
            strOutput = ZYQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "取消入院登记时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 出院登记
            strOutput = CYCS(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "出院登记时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 取消出院登记
            strOutput = CYCSQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "取消出院登记时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 增加处方单据
            strOutput = AddCFJL(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutput = "" Then
                MsgBox "增加处方单据时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case 增加处方明细
            strOutput = AddCFMX(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutput = "" Then
                MsgBox "增加处方明细时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strOutput = ""
            For i = 1 To UBound(strArr)
                strOutput = strOutput & "||" & strArr(i)
            Next
            If strOutput <> "" Then
                strOutput = Mid(strOutput, 3)
            End If
        Case 删除处方单据及其明细
            strOutput = DELCFJL(strInValue(0))
            If strOutput = "" Then
                MsgBox "删除处方单据及其明细时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 单条处方传输
            strOutput = CFCS(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "单条处方传输时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 出院结算
            strOutput = CYJS(strInValue(0), Val(strInValue(1)))
            If strOutput = "" Then
                MsgBox "出院结算时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                If InStr(1, strArr(0), "未达到起付标准额") <> 0 Then
                    '由于他们这边要求，如果未达起付线将会抛出该错误，为了能正常通过，所以只能采取此方返回该值.
                    strArr = Split("dfds|d", "|")
                    strArr(1) = "0||0||0||" & g病人身份_成都德阳.费用总额 & "||0||" & g病人身份_成都德阳.费用总额 & "||0||0||0||0||0||0||0||" & Format(zlDatabase.Currentdate, "yyyy") & "||" & "" & "||||||||||||||0||0||0||0||0||0||0||0||0||0||0||0||0||0||0||0||0||0||1||1||0||||现金||||||||||||||||1922-02-02||0||0||0||0||0||0||0||||||||||||||0||0||0||0||0||0||0||0||||0||"
                    gbln未达起付线 = True
                Else
                    MsgBox strArr(0), vbInformation, gstrSysName
                    Exit Function
                    gbln未达起付线 = False
                End If
            Else
                gbln未达起付线 = False
            End If
            strOutput = strArr(1)
        Case 取消出院结算
            strOutput = CYJSQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "出院结算时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 打印出院结算报表函数
            strOutput = JSReport(strInValue(0), strInValue(1))
            strOutput = ""
        Case 打印住院人员财务结算单
            strOutput = CWJSREPORT(strInValue(0), strInValue(1))
            strOutput = ""
        
        Case 提取基础资料
        
            strOutput = GETJCXX(strInValue(0), strInValue(1))
              If strOutput = "" Then
                MsgBox "提取基础资料时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 获取住院记录号
            strOutput = GETZYIDBYZYBH(strInValue(0))
            If strOutput = "" Then
                MsgBox "获取住院记录号时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case 检查拔号连接
            strOutput = CheckCon()
            If strOutput = "" Then
                MsgBox "检查拔号连接时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 建立拔号连接
            strOutput = RasDial(strInValue(0))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case 断开拔号连接
            strOutput = DisDial()
            strOutput = ""
        Case 获取药品信息
             strOutput = GetSINYPXX(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "获取药品信息时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case 跨月重提人员基本资料
'            strOutput = GETNEWRYZL(strInValue(0))
'              If strOutput = "" Then
'                MsgBox "跨月重提人员基本资料时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
    End Select
    strOutPutstring = strOutput
    业务请求_成都德阳 = True
    DebugTool " 输出参数为:" & strOutput
     Exit Function
    
errHand:
    DebugTool "业务请求失败"
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_成都德阳(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
            
    '刘兴宏:20040923增加的
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0
    On Error GoTo errHand
    
    DebugTool "进入扩院登撤消接口"
    
    入院登记撤销_成都德阳 = False
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用，不能撤消入院登记"
        Exit Function
    End If
    
    
    '获取住院号
    gstrSQL = "Select 医保住院号 住院号 From 保险帐户 where 病人ID=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "入院登记撤销"
    If 业务请求_成都德阳(取消入院登记, Nvl(rsTemp!住院号), strOutput) = False Then Exit Function

    '更新医保帐户
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_成都德阳 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    DebugTool "取消成功"
    入院登记撤销_成都德阳 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_成都德阳(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0
    On Error GoTo errHand:
    
          
    '改变当前状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_成都德阳 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_成都德阳 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_成都德阳 = False
End Function
Public Function 出院登记撤销_成都德阳(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '出院登记撤消
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
     '改变病人状态
     If Not 存在未结费用(lng病人ID, lng主页ID) Then
            ShowMsgbox "该病人已经出院结算了,不能再取消出院!"
            Exit Function
     End If
    
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_成都德阳 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    出院登记撤销_成都德阳 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_成都德阳(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
  '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String
    
    Dim lng主页ID As Long, lng是否上传 As Long
    Dim dbl费用总额 As Double
    Dim strArr
    Dim str结算方式  As String, str住院号 As String
    Dim obj结算 As 结算数据
        
    Err = 0: On Error GoTo errHand:
    Call DebugTool("进入住院结算")
    
    
    If g病人身份_成都德阳.病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人没有完成医保的预结算操作，不能进行结算。", vbInformation, gstrSysName
        Exit Function
    End If
        
    gstrSQL = "Select 当前状态,医保住院号 住院号 From 保险帐户 where 病人ID=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断当前的住院状态!"
    If Nvl(rsTemp!当前状态, 0) = 1 Then
        Err.Raise 9000, gstrSysName, "当前病人还处于在院状态,请出院后再结算!"
        Exit Function
    End If
    
    str住院号 = Nvl(rsTemp!住院号)
    gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID)
    If IsNull(rsTemp("主页ID")) = True Then
        Err.Raise 9000, gstrSysName, "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    lng主页ID = rsTemp("主页ID")
        

    gstrSQL = " " & _
          " Select sum(nvl(结帐金额,0)) as 实收金额 " & _
          " From 住院费用记录 " & _
          " Where 记录状态<>0 and 结帐ID=" & lng结帐ID & " and  Nvl(附加标志,0)<>9"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取总费用"
    dbl费用总额 = Nvl(rsTemp!实收金额, 0)
    If dbl费用总额 <> g病人身份_成都德阳.费用总额 Then
        Err.Raise 9000, gstrSysName, "费用总额不等"
        Exit Function
    End If
    g病人身份_成都德阳.费用总额 = dbl费用总额
 
    
    '再次结算
  
    'AZYH    PCHAR   住院号
    'ISPREV  PCHAR   预结算标志（'0'－表示预结算）
    StrInput = str住院号
    StrInput = StrInput & "||1"
    If 业务请求_成都德阳(出院结算, StrInput, strOutput) = False Then Exit Function
    strArr = Split(strOutput, "||")
    
    '返回值
    '应支付统筹金||超封顶自付||比例自付小计||个人支付合计||统筹支付统筹金||个人应付支付||本次补保陪付额||本次补保进入基数||实际扣减基数||统筹封顶金额||
    '统筹起付金额||住院记录号||个人记录号||年月||住院号||病种编号||病种名称||科室||医疗机构号||入院日期||出院日期||已结算统筹金||发生医疗费小计||发生药品费||
    '发生检查费||发生治疗费||发生其它费||自付小计||自付药品费||自付检查费||自付治疗费||自付其它费||比例药品费||比例检查费||比例治疗费||比例其它费||统筹支付比例||
    '统筹治疗费||统筹其它费||出院标志||结算标志||传输标志||基本医疗状态||结算方式||审核方式||结算日期||单位编号||单位名称||社会保障号||姓名||性别||出生日期||预缴金额||
    '个人应补金额||个人实补金额||退款金额||个人实际支付金额||社保结算金额||财务结算日期||财务结算标志||保险机构号||操作员编号||资料提取时间||医疗保险编号||社保机构名称||
    '补充险种类型||补充享受标志||补充起付扣减标志||补充首陪金额||补充起点基数||补充陪付比例||补充已陪付额||补充陪付最大金额||补充待遇享受开始年月||陪付总额||
    With obj结算
        .医保基金 = Val(strArr(4))
        .补保陪付额 = Val(strArr(6))
    End With
    
    If InsertInto医保结算记录(strArr, lng结帐ID) = False Then Exit Function
    
    
    '办理出院登记
    gstrSQL = "" & _
          "   Select B.医保住院号 住院号,to_Char(a.出院日期,'yyyy-MM-DD') as 出院日期" & _
          "   From 病案主页 A,保险帐户 B " & _
          "   Where A.病人iD=b.病人id " & _
          "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取住院号和出院日期"
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "无对应的住院人员信息"
        Exit Function
    End If
    
    lng是否上传 = 0
    If gbln未达起付线 Then
        '未达起付线的,需将医保中心的病人入院登记撤消
        '获取住院号
        If 业务请求_成都德阳(取消入院登记, Nvl(rsTemp!住院号), strOutput) = False Then Exit Function
    '曾明春(2006-2-20):德阳市人民医院要求出院登记(出院审核)手工处理
    Else
        If InitInfor_成都德阳.自动出院审核 Then
           StrInput = Nvl(rsTemp!住院号)
           StrInput = StrInput & "||" & Nvl(rsTemp!出院日期)
           If 业务请求_成都德阳(出院登记, StrInput, strOutput) = False Then
              Exit Function
           Else
              lng是否上传 = 1
           End If
        End If
    End If
    
    '填写结算表
    Call DebugTool("填写结算记录")
    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(),住院次数_IN(主页ID),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(补保陪付额),超限自付金额_IN(无),个人帐户支付_IN(),"
    '   支付顺序号_IN(住院号),主页ID_IN(主页ID),中途结帐_IN,备注_IN(未达起付线),校正_IN(),是否上传_IN(是否上传)
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
   
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_成都德阳 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,NULL," & lng主页ID & ",0,0,0," & _
            dbl费用总额 & ",0,0," & _
            obj结算.医保基金 & "," & obj结算.医保基金 & ",0,0," & obj结算.补保陪付额 & ",'" & _
            str住院号 & "'," & lng主页ID & ",NULL," & IIf(gbln未达起付线, "'未达起付线'", "NULL") & ",NULL," & lng是否上传 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
    住院结算_成都德阳 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Function InsertInto医保结算记录(ByVal strArr As Variant, ByVal lng结帐ID As Long) As Boolean
    '功能:往中间库插入医保结算记录
    '参数:strarr以split(stroutput,"||")产生的数组
    
    Err = 0
    On Error GoTo errHand:
    InsertInto医保结算记录 = False
    
    DebugTool "进入InsertInto医保结算记录"
    'strArr:
    '应支付统筹金||超封顶自付||比例自付小计||个人支付合计||统筹支付统筹金||个人应付支付||本次补保陪付额||本次补保进入基数||实际扣减基数||统筹封顶金额||
    '统筹起付金额||住院记录号||个人记录号||年月||住院号||病种编号||病种名称||科室||医疗机构号||入院日期||出院日期||已结算统筹金||发生医疗费小计||发生药品费||
    '发生检查费||发生治疗费||发生其它费||自付小计||自付药品费||自付检查费||自付治疗费||自付其它费||比例药品费||比例检查费||比例治疗费||比例其它费||统筹支付比例||
    '统筹治疗费||统筹其它费||出院标志||结算标志||传输标志||基本医疗状态||结算方式||审核方式||结算日期||单位编号||单位名称||社会保障号||姓名||性别||出生日期||预缴金额||
    '个人应补金额||个人实补金额||退款金额||个人实际支付金额||社保结算金额||财务结算日期||财务结算标志||保险机构号||操作员编号||资料提取时间||医疗保险编号||社保机构名称||
    '补充险种类型||补充享受标志||补充起付扣减标志||补充首陪金额||补充起点基数||补充陪付比例||补充已陪付额||补充陪付最大金额||补充待遇享受开始年月||陪付总额||
    
    '过程参数
    '性质,结帐ID,
    '应支付统筹金,超封顶自付,比例自付小计,个人支付合计,统筹支付统筹金,个人应付支付,本次补保陪付额,本次补保进入基数,实际扣减基数,统筹封顶金额,统筹起付金额,住院记录号,个人记录号,
    '年月,住院号,病种编号,病种名称,科室,医疗机构号,入院日期,出院日期,已结算统筹金,发生医疗费小计,发生药品费,发生检查费,发生治疗费,发生其它费,自付小计,自付药品费,自付检查费,自付治疗费,
    '自付其它费,比例药品费,比例检查费,比例治疗费,比例其它费,统筹支付比例,统筹治疗费,统筹其它费,出院标志,结算标志,传输标志,基本医疗状态,结算方式,审核方式,结算日期,单位编号,单位名称,
    '社会保障号,姓名,性别,出生日期,预缴金额,个人应补金额,个人实补金额,退款金额,个人实际支付金额,社保结算金额,财务结算日期,财务结算标志,保险机构号,操作员编号,资料提取时间,医疗保险编号,
    '社保机构名称,补充险种类型,补充享受标志,补充起付扣减标志,补充首陪金额,补充起点基数,补充陪付比例,补充已陪付额,补充陪付最大金额,补充待遇享受开始年月,陪付总额
    
    '    性质        number(2),
    gstrSQL = "ZL_医保结算记录_INSERT(2"
    '    结帐ID      number(18),
    gstrSQL = gstrSQL & "," & lng结帐ID
    '    应支付统筹金    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(0))
    '    超封顶自付  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(1))
    '    比例自付小计    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(2))
    '    个人支付合计    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(3))
    '    统筹支付统筹金  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(4))
    '    个人应付支付    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(5))
    '    本次补保陪付额  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(6))
    '    本次补保进入基数    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(7))
    '    实际扣减基数    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(8))
    '    统筹封顶金额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(9))
    '    统筹起付金额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(10))
    
    '    住院记录号  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(11) & "'"
    '    个人记录号  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(12) & "'"
    '    年月        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(13) & "'"
    '    住院号      varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(14) & "'"
    '    病种编号        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(15) & "'"
    '    病种名称        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(16) & "'"
    '    科室        varchar2(50),
    gstrSQL = gstrSQL & ",'" & strArr(17) & "'"
    '    医疗机构号  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(18) & "'"
    '    入院日期        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(19) & "'"
    '    出院日期        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(20) & "'"
      
    '    已结算统筹金    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(21))
    '    发生医疗费小计  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(22))
    '    发生药品费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(23))
    '    发生检查费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(24))
    '    发生治疗费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(25))
    '    发生其它费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(26))
    '    自付小计        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(27))
    '    自付药品费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(28))
    '    自付检查费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(29))
    '    自付治疗费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(30))
    '    自付其它费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(31))
    '    比例药品费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(32))
    '    比例检查费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(33))
    '    比例治疗费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(34))
    '    比例其它费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(35))
    '    统筹支付比例    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(36))
    '    统筹治疗费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(37))
    '    统筹其它费  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(38))
      
    '    出院标志        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(39) & "'"
    '    结算标志        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(40) & "'"
    '    传输标志        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(41) & "'"
    '    基本医疗状态    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(42) & "'"
    '    结算方式        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(43) & "'"
    '    审核方式        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(44) & "'"
    '    结算日期        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(45) & "'"
    '    单位编号        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(46) & "'"
    '    单位名称        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(47) & "'"
    '    社会保障号  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(48) & "'"
    '    姓名        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(49) & "'"
    '    性别        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(50) & "'"
    '    出生日期        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(51) & "'"
        
    '    预缴金额        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(52))
    '    个人应补金额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(53))
    '    个人实补金额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(54))
    '    退款金额        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(55))
    '    个人实际支付金额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(56))
    '    社保结算金额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(57))
            
    '    财务结算日期    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(58) & "'"
    '    财务结算标志    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(59) & "'"
    '    保险机构号  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(60) & "'"
    '    操作员编号  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(61) & "'"
    '    资料提取时间    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(62) & "'"
    '    医疗保险编号    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(63) & "'"
    '    社保机构名称    varchar2(50),
    gstrSQL = gstrSQL & ",'" & strArr(64) & "'"
    '    补充险种类型    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(65) & "'"
    '    补充享受标志    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(66) & "'"
    '    补充起付扣减标志    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(67) & "'"
            
    '    补充首陪金额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(68))
    '    补充起点基数    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(69))
    '    补充陪付比例    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(70))
    '    补充已陪付额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(71))
    '    补充陪付最大金额    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(72))
            
    '    补充待遇享受开始年月    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(73) & "'"
    '    陪付总额        number(16,5))
    gstrSQL = gstrSQL & "," & Val(strArr(74)) & ")"
    gcnOracle_成都德阳.Execute gstrSQL, , adCmdStoredProc
    InsertInto医保结算记录 = True
    DebugTool "保存医保结算记录成功"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 住院结算冲销_成都德阳(lng结帐ID As Long) As Boolean
     '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim rs结算记录 As New ADODB.Recordset
    
    Dim StrInput As String, strOutput  As String
    Dim lng冲销ID As Long, str住院号 As String
    Dim strArr
    Dim lng病人id1 As Long, lng病人ID As Long
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo errHand:
    
    
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", TYPE_成都德阳, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    lng病人ID = Nvl(rsTemp!病人ID, 0)
    gstrSQL = "select * from 医保结算记录 where 性质=2  and 结帐ID=" & lng结帐ID
    Call OpenRecordset_成都德阳(rs结算记录, "结算冲销")
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
        
    '判断病人的住院结算数据是否允许作废。判断标准是检查病人有新的住院记录，如果有，就不能交冲销
    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If 身份标识_成都德阳(2, lng病人id1) = "" Then
      Screen.MousePointer = intMouse
      Exit Function
    End If
    Screen.MousePointer = intMouse
    If lng病人ID <> lng病人id1 Then
      Err.Raise 9000, gstrSysName, "不是当前要冲销结算的病人!"
      Exit Function
    End If
    

    If Nvl(rsTemp!备注) = "未达起付线" Then
        '未达起付线不需取消出院
        '需重新办理入院登记
        If 入院登记_成都德阳(lng病人ID, Nvl(rsTemp!主页ID, 0), g病人身份_成都德阳.保障号) = False Then Exit Function
    Else
        '取消出院
        str住院号 = rsTemp("支付顺序号")
        '曾明春(2006-2-20):判断是否作了出院登记
        If InitInfor_成都德阳.自动出院审核 = False Then
            If rsTemp!是否上传 = 1 Then
               Err.Raise 9000, gstrSysName, "该病人已经出院审核,请到医保科取消出院审核!"
               Exit Function
            End If
        Else
            If 业务请求_成都德阳(取消出院登记, str住院号, strOutput) = False Then Exit Function
        End If
        StrInput = str住院号
        If 业务请求_成都德阳(取消出院结算, StrInput, strOutput) = False Then
            Exit Function
        End If
    End If
    
    
    '应支付统筹金||超封顶自付||比例自付小计||个人支付合计||统筹支付统筹金||个人应付支付||本次补保陪付额||本次补保进入基数||实际扣减基数||统筹封顶金额||统筹起付金额||住院记录号||个人记录号||年月||住院号||病种编号||病种名称||科室||医疗机构号||入院日期||出院日期||已结算统筹金||发生医疗费小计||发生药品费||发生检查费||发生治疗费||发生其它费||自付小计||自付药品费||自付检查费||自付治疗费||自付其它费||比例药品费||比例检查费||比例治疗费||比例其它费||统筹支付比例||统筹治疗费||统筹其它费||出院标志||结算标志||传输标志||基本医疗状态||结算方式||审核方式||结算日期||单位编号||单位名称||社会保障号||姓名||性别||出生日期||预缴金额||个人应补金额||个人实补金额||退款金额||个人实际支付金额||社保结算金额||财务结算日期||财务结算标志||保险机构号||操作员编号||资料提取时间||医疗保险编号||社保机构名称||补充险种类型||补充享受标志||补充起付扣减标志||补充首陪金额||补充起点基数||补充陪付比例||补充已陪付额||补充陪付最大金额||补充待遇享受开始年月||陪付总额
    strArr = Split("应支付统筹金||超封顶自付||比例自付小计||个人支付合计||统筹支付统筹金||个人应付支付||本次补保陪付额||本次补保进入基数||实际扣减基数||统筹封顶金额||统筹起付金额||住院记录号||个人记录号||年月||住院号||病种编号||病种名称||科室||医疗机构号||入院日期||出院日期||已结算统筹金||发生医疗费小计||发生药品费||发生检查费||发生治疗费||发生其它费||自付小计||自付药品费||自付检查费||自付治疗费||自付其它费||比例药品费||比例检查费||比例治疗费||比例其它费||统筹支付比例||统筹治疗费||统筹其它费||出院标志||结算标志||传输标志||基本医疗状态||结算方式||审核方式||结算日期||单位编号||单位名称||社会保障号||姓名||性别||出生日期||预缴金额||个人应补金额||个人实补金额||退款金额||个人实际支付金额||社保结算金额||财务结算日期||财务结算标志||保险机构号||操作员编号||资料提取时间||医疗保险编号||社保机构名称||补充险种类型||补充享受标志||补充起付扣减标志||补充首陪金额||补充起点基数||补充陪付比例||补充已陪付额||补充陪付最大金额||补充待遇享受开始年月||陪付总额", "||")
    
    StrInput = ""
    Dim i As Integer
    For i = 0 To UBound(strArr)
        If rs结算记录.Fields(strArr(i)).Type = 131 Then
            StrInput = StrInput & "||" & (Val(Nvl(rs结算记录.Fields(strArr(i)))) * -1)
        Else
            StrInput = StrInput & "||" & Nvl(rs结算记录.Fields(strArr(i)))
        End If
    Next
    StrInput = Mid(StrInput, 3)
    strArr = Split(StrInput, "||")
    If InsertInto医保结算记录(strArr, lng冲销ID) = False Then Exit Function
    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(),住院次数_IN(主页ID),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(补保陪付额),超限自付金额_IN(无),个人帐户支付_IN(),"
    '   支付顺序号_IN(住院号),主页ID_IN(主页ID),中途结帐_IN,备注_IN(),校正_IN,是否上传_IN(是否上传)

    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_成都德阳 & "," & rsTemp("病人ID") & "," & Year(zlDatabase.Currentdate) & "," & _
        "NULL,NULL,NULL,NULL," & Nvl(rsTemp!主页ID, 0) & ",0,0,0," & _
        rsTemp("发生费用金额") * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & "," & Nvl(rsTemp!大病自付金额, 0) * -1 & ",0," & _
        "NULL,'" & str住院号 & "'," & Nvl(rsTemp!主页ID, 0) & ",NULL," & Nvl(rsTemp!备注) & ",NULL,0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存医保结算记录")
    
    住院结算冲销_成都德阳 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function 处方登记_成都德阳(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
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
    Dim str处方记录号 As String
    Dim strArr
    Dim str住院号 As String
    
    Err = 0
    On Error GoTo errHand:
    
    处方登记_成都德阳 = False
    
   '读出该张单据的费用明细
    gstrSQL = "Select A.ID,A.NO,a.标识号 住院号,A.病人ID,A.主页ID,to_char(A.发生时间,'yyyy-mm-dd') as 登记时间,Round(A.实收金额,4) 实收金额 " & _
              "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,M.中心,Q.名称 as 开单部门,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位 " & _
              "  From 住院费用记录 A,收费细目 B,保险帐户 M,部门表 Q,病人信息 J" & _
              "  where A.NO=[2] and A.记录性质=[3] and A.记录状态 = [4] And Nvl(A.是否上传,0)=0 " & _
              "        and A.病人ID=J.病人ID and a.病人id=m.病人id and a.开单部门id=Q.id(+) And M.险类=[1]" & _
              "        and A.收费细目ID=B.ID  " & _
              "  Order by A.病人ID,A.NO,A.发生时间"
              
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "处方明细上传", TYPE_成都德阳, str单据号, lng记录性质, lng记录状态)
    
    With rs明细
    
        If .RecordCount = 0 Then
            ShowMsgbox "没有相关的明细记录,可能有些项目未进行医保对码!"
            Exit Function
        End If
        Do While Not .EOF
            gstrSQL = "Select * From 医保支付项目_德阳 where 险类=[1] and 中心=[2] and 收费细目id=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "确定医保支付项目", TYPE_成都德阳, CLng(Nvl(!中心, 0)), CLng(Nvl(!收费细目ID, 0)))
            If rsTemp.EOF Then
                ShowMsgbox "项目[" & Nvl(!名称) & "]未进行对码!"
                Exit Function
            End If
            .MoveNext
        Loop
        If InitInfor_成都德阳.明细时实上传 = False Then
            处方登记_成都德阳 = True
            Exit Function
        End If
        lng病人ID = 0
        str处方记录号 = ""
        Dim str摘要 As String
        .MoveFirst
        Do While Not .EOF
            gstrSQL = "Select * From 医保支付项目_德阳 where 险类=[1] and 中心=[2] and 收费细目id=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "确定医保支付项目", TYPE_成都德阳, CLng(Nvl(!中心, 0)), CLng(Nvl(!收费细目ID, 0)))
                        
            If lng病人ID <> Nvl(!病人ID, 0) Then
                 '需增加一张单据
                 'AZYH    PCHAR   住院号
                 'ACFRQ   PCHAR   处方日期（YYYY-MM-DD）
                 'AYS PCHAR   医生
                 'AKS PCHAR   科室
                 lng病人ID = Nvl(!病人ID, 0)
                 lng主页ID = Nvl(!主页ID, 0)
                 str住院号 = Get住院号(lng病人ID, Nvl(!主页ID, 0), True)
                 StrInput = str住院号
                 StrInput = StrInput & "||" & Nvl(!登记时间)
                 StrInput = StrInput & "||" & Nvl(!医生)
                 StrInput = StrInput & "||" & Nvl(!开单部门)
                 If 业务请求_成都德阳(增加处方单据, StrInput, strOutput) = False Then Exit Function
                 str处方记录号 = strOutput
                 
'                 '单条处方传输
'                'AZYH    PCHAR   住院号
'                'ACFID   PCHAR   处方记录号（通过调用ADDCFJL返回的值）
'                 strInput = Lpad(Nvl(!住院号, 0), 8)
'                 strInput = strInput & "||" & str处方记录号
'                 If 业务请求_成都德阳(单条处方传输, strInput, stroutput) = False Then
'                    '需删除该张单据
'                    Call 业务请求_成都德阳(删除处方单据及其明细, str处方记录号, stroutput)
'                    Exit Function
'                 End If
            End If
            '增加处方明细
            'ACFID   PCHAR   处方记录号
            'AYPBH   PCHAR   药品编号(社保药品编号)
            'ASL PCHAR   数量(可以为负数)
            'ADJ PCHAR   单价
            StrInput = str处方记录号
            StrInput = StrInput & "||" & Nvl(rsTemp!项目编码)
            StrInput = StrInput & "||" & Nvl(!数量)
            StrInput = StrInput & "||" & Nvl(!价格)
            
            If 业务请求_成都德阳(增加处方明细, StrInput, strOutput) = False Then
                '需删除该张单据
                Call 业务请求_成都德阳(删除处方单据及其明细, str处方记录号, strOutput)
                Exit Function
            End If
           '处方明细记录号||自费比例||自费金额
           '摘要保存值:处方记录号||明细记录号||自费比例||自费金额||住院号
            str摘要 = str处方记录号 & "||" & strOutput & "||" & str住院号
            '更改上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str摘要 & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            .MoveNext
        Loop
    End With
    处方登记_成都德阳 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function Read模拟数据(ByVal int业务类型 As 业务类型_成都德阳, ByVal strInputString As String, ByRef strOutPutstring As String)
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
                    strArr = Split(strText, vbTab)
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab
                            End If
                            strArr = Split(strText, vbTab)
                            
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
    If InStr(1, strOutPutstring, "@$") <> 0 Then
        strOutPutstring = Split(strOutPutstring, "@$")(1)
    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_成都德阳(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_成都德阳, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Function 住院虚拟结算_成都德阳(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    
    'rsExse:字符集
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng主页ID As Long
    Dim StrInput As String, strOutput   As String
    Dim strArr As Variant
    Dim str住院号 As String, str结算方式 As String
    
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo errHand:
    
    g病人身份_成都德阳.病人ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
   
    gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID)
    If IsNull(rsTemp("主页ID")) = True Then
        MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    lng主页ID = rsTemp("主页ID")
    
    gstrSQL = "Select to_char(入院日期,'yyyyMM') as 入院日期 ,to_char(出院日期,'yyyyMM') as 出院日期 from 病案主页 where 病人id=[1] and 主页id =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID, lng主页ID)
    If Nvl(rsTemp!入院日期) <> Nvl(rsTemp!出院日期) And Nvl(rsTemp!出院日期) <> "" Then
        '如果不存同一月份,需重提资料
        '建立拔号连接
        gstrSQL = "Select a.编码,b.医保住院号 as 住院号 From 保险中心目录 a,保险帐户 b where a.序号=b.中心 and a.险类=140 and b.病人id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息", lng病人ID)
        g病人身份_成都德阳.机构编码 = Nvl(rsTemp!编码)
        If gbln检查连接 Then
            '检查拔号连接
            If 业务请求_成都德阳(检查拔号连接, "", strOutput) = False Then
                 Call 业务请求_成都德阳(断开拔号连接, "", strOutput)
            End If
            If 业务请求_成都德阳(建立拔号连接, g病人身份_成都德阳.机构编码, strOutput) = False Then
                 Exit Function
            End If
            If 业务请求_成都德阳(检查拔号连接, "", strOutput) = False Then
                 Exit Function
            End If
        End If
        
        If 业务请求_成都德阳(跨月重提人员基本资料, Get住院号(lng病人ID, lng主页ID, True), strOutput) = False Then Exit Function
        
    End If
    
    
    gstrSQL = "Select 医保住院号 住院号,中心 From 保险帐户 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取住院号"
    If rsTemp.EOF Then
        ShowMsgbox "该病人不是医保病人!"
        Exit Function
    End If
    
    str住院号 = Nvl(rsTemp!住院号)
    g病人身份_成都德阳.中心 = Nvl(rsTemp!中心, 0)

    Screen.MousePointer = vbHourglass
    If 补传住院明细记录(lng病人ID, lng主页ID) = False Then Exit Function
    
    With rsExse
        g病人身份_成都德阳.费用总额 = 0
        Do While Not .EOF
            g病人身份_成都德阳.费用总额 = g病人身份_成都德阳.费用总额 + Nvl(!金额, 0)
            .MoveNext
        Loop
    End With
    
    'AZYH    PCHAR   住院号
    'ISPREV  PCHAR   预结算标志（'0'－表示预结算）
    StrInput = str住院号
    StrInput = StrInput & "||0"
    If 业务请求_成都德阳(出院结算, StrInput, strOutput) = False Then Exit Function
    strArr = Split(strOutput, "||")
    
    '返回值
    '应支付统筹金||超封顶自付||比例自付小计||个人支付合计||统筹支付统筹金||个人应付支付||本次补保陪付额||本次补保进入基数||实际扣减基数||统筹封顶金额||
    '统筹起付金额||住院记录号||个人记录号||年月||住院号||病种编号||病种名称||科室||医疗机构号||入院日期||出院日期||已结算统筹金||发生医疗费小计||发生药品费||
    '发生检查费||发生治疗费||发生其它费||自付小计||自付药品费||自付检查费||自付治疗费||自付其它费||比例药品费||比例检查费||比例治疗费||比例其它费||统筹支付比例||
    '统筹治疗费||统筹其它费||出院标志||结算标志||传输标志||基本医疗状态||结算方式||审核方式||结算日期||单位编号||单位名称||社会保障号||姓名||性别||出生日期||预缴金额||
    '个人应补金额||个人实补金额||退款金额||个人实际支付金额||社保结算金额||财务结算日期||财务结算标志||保险机构号||操作员编号||资料提取时间||医疗保险编号||社保机构名称||
    '补充险种类型||补充享受标志||补充起付扣减标志||补充首陪金额||补充起点基数||补充陪付比例||补充已陪付额||补充陪付最大金额||补充待遇享受开始年月||陪付总额||
    
    
    With g虚拟结算
    
        .医保基金 = Val(strArr(4))
        .补保陪付额 = Val(strArr(6))
    End With
    If Format(strArr(22), "####0.00;-####0.00;0;0") <> Format(g病人身份_成都德阳.费用总额, "####0.00;-####0.00;0;0") Then
        ShowMsgbox "结算数据不等!" & vbCrLf & "医保中心费用总额:" & Format(strArr(22), "####0.00;-####0.00;0;0") & vbCrLf & " 医院端为:" & Format(g病人身份_成都德阳.费用总额, "####0.00;-####0.00;0;0")
        If InitInfor_成都德阳.数据不等不可结算 Then
            Exit Function
        End If
    End If
    
    
    
    str结算方式 = "医保基金;" & g虚拟结算.医保基金 & ";0"
    str结算方式 = str结算方式 & "|补保陪付额;" & g虚拟结算.补保陪付额 & ";0"
    住院虚拟结算_成都德阳 = str结算方式
    g病人身份_成都德阳.病人ID = lng病人ID   '表示该病人已经进行了虚拟结算
    
    '打印结算单:必须在结算后才能打印
'    If InitInfor_成都德阳.打印结算单 Then
'        '调打印接口
'        '    ASTARTZYH   PCHAR   打印开始住院号
'        '    AENDZYH PCHAR   打印结束住院号
'
'        StrInput = str住院号 & "||"
'        StrInput = StrInput & str住院号
'        Call 业务请求_成都德阳(打印出院结算报表函数, StrInput, strOutput)
'    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 补传住院明细记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '补传相关明细记录
    Dim cnTemp As New ADODB.Connection
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput  As String, strOutput As String
    Dim strArr
    Dim str住院号 As String, str处方记录号 As String
    
    Err = 0
    On Error GoTo errHand:
      
    
    Call DebugTool("打开新连接")
    Set cnTemp = GetNewConnection
    cnTemp.Open
    Call DebugTool("打开连接成功，开始检查明细数据的合法性。")
      
    补传住院明细记录 = False
    

    gstrSQL = "Select A.ID,A.标识号 as 住院号,A.NO,A.记录性质,A.记录状态,A.序号,A.病人ID,A.主页ID,to_char(A.发生时间,'yyyy-mm-dd')  as 登记时间,Round(A.实收金额,4) 实收金额" & _
              "         ,M.名称 as 开单部门,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,C.项目编码,C.附注,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位" & _
              "  From 住院费用记录 A,收费细目 B,(Select * From 医保支付项目_德阳 where 中心=[2] and 险类 =[1]) C,病案主页 D,部门表 M" & _
              "  where A.病人ID=[3] and A.主页ID=[4] and A.记帐费用=1 and nvl(A.实收金额,0)<>0 and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 " & _
              "        and A.开单部门id =M.id(+) and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=[1]" & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID(+) " & _
              "  Order by A.病人ID,A.记录性质,A.No,A.记录状态,A.序号"
              
              
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", TYPE_成都德阳, g病人身份_成都德阳.中心, lng病人ID, lng主页ID)
    
    
    
   With rs明细
        Do While Not .EOF
            If Nvl(!项目编码) = "" Then
                ShowMsgbox "项目[" & Nvl(!编码) & "]" & Nvl(!名称) & " 未进行医保对码,请在保险项目管理中设置对码关系!"
                Exit Function
            End If
            .MoveNext
        Loop
        
        If .RecordCount <> 0 Then .MoveFirst
        Dim strNO As String
        
        str处方记录号 = ""
        strNO = ""
        Dim str摘要 As String
        
        Do While Not .EOF
            If strNO <> Nvl(!记录性质, 0) & "_" & Nvl(!NO) & "_" & Nvl(!记录状态, 0) Then
                strNO = Nvl(!记录性质, 0) & "_" & Nvl(!NO) & "_" & Nvl(!记录状态, 0)
                 '需增加一张单据
                 'AZYH    PCHAR   住院号
                 'ACFRQ   PCHAR   处方日期（YYYY-MM-DD）
                 'AYS PCHAR   医生
                 'AKS PCHAR   科室
                 str住院号 = Get住院号(lng病人ID, lng主页ID, True)
                 StrInput = str住院号
                 StrInput = StrInput & "||" & Nvl(!登记时间)
                 StrInput = StrInput & "||" & Nvl(!医生)
                 StrInput = StrInput & "||" & Nvl(!开单部门)
                 
                 If 业务请求_成都德阳(增加处方单据, StrInput, strOutput) = False Then Exit Function
                 str处方记录号 = strOutput
            End If
            '增加处方明细
            'ACFID   PCHAR   处方记录号
            'AYPBH   PCHAR   药品编号(社保药品编号)
            'ASL PCHAR   数量(可以为负数)
            'ADJ PCHAR   单价
            StrInput = str处方记录号
            StrInput = StrInput & "||" & Nvl(!项目编码)
            StrInput = StrInput & "||" & Nvl(!数量)
            StrInput = StrInput & "||" & Nvl(!价格)
            
            If 业务请求_成都德阳(增加处方明细, StrInput, strOutput) = False Then
                '需删除该张单据
                Call 业务请求_成都德阳(删除处方单据及其明细, str处方记录号, strOutput)
                Exit Function
            End If
           '处方明细记录号||自费比例||自费金额
           '摘要保存值:处方记录号||明细记录号||自费比例||自费金额||住院号
            str摘要 = str处方记录号 & "||" & strOutput & "||" & str住院号
            '更改上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str摘要 & "')"
             cnTemp.Execute gstrSQL, , adCmdStoredProc
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
Public Function 医保设置_成都德阳() As Boolean
    医保设置_成都德阳 = frmSet成都德阳.参数设置
End Function
Public Sub ExecuteProcedure_成都德阳(ByVal strCaption As String)
'功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_成都德阳.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Function 撤消医保入院_成都德阳(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：更新病人的出院疾病。如果是肿瘤，则结算时起付线会减半
    Dim StrInput As String
    Dim strOutput As String
    Dim blnYes  As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    
    On Error GoTo errHandle
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "该病人已经出院或无费用发生，不能撤消，可以通过入院登记进行入院撤消!"
        Exit Function
    End If
    
    
    gstrSQL = "Select * From 住院费用记录 where nvl(是否上传,0)=1 and rownum<=1 and 病人id=" & lng病人ID & " and 主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断是否存在上传记录"
        
    If Not rsTemp.EOF Then
        ShowMsgbox "已经有上传到中心明细费用，是否真的要取消医保入院?", True, blnYes
        If blnYes = False Then
            Exit Function
        End If
    End If
    
    
    '获取住院号
    gstrSQL = "Select 医保住院号 住院号 From 保险帐户 where 病人ID=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "入院登记撤销"
    If 业务请求_成都德阳(取消入院登记, Nvl(rsTemp!住院号), strOutput) = False Then Exit Function

    '更新医保帐户
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_成都德阳 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    
    gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消医保入院")
    
    DebugTool "取消成功"
    
    撤消医保入院_成都德阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get住院号(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional bln帐户 As Boolean = False) As String
    '功能:获取住院号
    Dim rsTemp As New ADODB.Recordset
    
    If bln帐户 Then
        gstrSQL = "Select 医保住院号 住院号 From 保险帐户 where 病人ID=[2]"
    Else
        gstrSQL = "Select 住院号 from 病人信息 where 病人id=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取住院号", lng病人ID)
    If rsTemp.EOF Then
        ShowMsgbox "未找到相关的病人信息!"
        Exit Function
    End If
    If Nvl(rsTemp!住院号) = "" Then
        ShowMsgbox "不存在相关的住院号,请检查!"
        Exit Function
    End If
    Get住院号 = Nvl(rsTemp!住院号) & IIf(bln帐户, "", Left(Lpad(lng主页ID, 2, "0"), 2))
End Function

