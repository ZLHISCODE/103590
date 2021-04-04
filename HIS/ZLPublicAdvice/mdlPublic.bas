Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gclsInsure As Object
Public gobjLIS As Object
Public gobjKernel As zlCISKernel.clsCISKernel
Public gobjCISJob As Object
Public glngSys As Long
Public glngModule As Long
Public gMainPrivs As String
Public gstrDBUser As String
Public gstrNodeNo As String          '当前站点编号；如果未设置启用站点，则为"-"
Public gcolPrivs As Collection              '记录内部模块的权限
Public lngNumPublicAdvice As Long  '计数器，记录clsPublicAdvice类被创建的次数


Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
    专业技术职务 As String
    用药级别 As Long
End Type

Public Enum Msg_Type '消息提醒类别
    m新开 = 1
    m新停 = 2
    m新废 = 3
    m安排 = 4
    m危机值 = 5
    m输液拒绝 = 6
    m销帐申请 = 7
    mRIS预约 = 8
    mRIS预约准备 = 9
    m取血通知 = 10
    m标本拒收 = 11
    m备血完成 = 12
    m血袋回收 = 13
End Enum

Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤消出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
    
    support允许部份冲销单据 = 28
    support出院无实际交易 = 29 '出院接口中是否要与接口商进行交易
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    
    support医生确定处方类型 = 48
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support实时监控 = 60
    
    support上传门诊档案 = 70                    '在门诊医嘱发送时，是否调用TranElecDossier函数完成门诊病人电子卷宗/电子档案的上传
End Enum

Public UserInfo As TYPE_USER_INFO

Public gobjExpense As Object  '费用公共部件

'所需参数
Public gbln药品按规格下医嘱 As Boolean
Public gstr输液配置中心 As String
Public gbln执行前先结算 As Boolean '一卡通执行前先收费或记帐审核
Public gblnStock As Boolean '指定药房时是否限定输入药品的库存
Public gstrLike As String  '项目匹配方法,%或空
Public gbytCode As Byte '简码输入方式
Public gbytDecPrice As Byte '费用单价的小数点位数
Public gstrDecPrice As String '价格按小数位数计算的格式化串,如"0.0000"
Public gbln发送生成条形码 As Boolean '是否在检验医嘱发送时生成条形码
Public gbytMediOutMode As Byte '分批药品出库方式：0-按批次先进先出，1-按效期最近先出,效期相同，则再按批次先进先出
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按金额小数位数计算的格式化串,如"0.0000"
Public gstr动态费别 As String               '存放门诊当前科室可用动态费别,在公共函数中使用,使用时才赋值:CalcDrugPrice,CalcPrice
Public gbln从项汇总折扣 As Boolean '从属项目汇总计算折扣
Public gstr住院发送划价单 As String
Public gstr门诊发送划价单 As String
Public gdbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
                                                      '为负数(-N)时表示,N元内免密支付,表示病人在消费N元内必须刷卡,不必输入密码即可支付;否则必须输入密码
Public gbyt住院自动发料 As Byte  '住院记帐完成后是否自动发料 0-不自动发料，1-自动发料，2-本科室开单时自动发料
Public gbln门诊自动发料 As Boolean '门诊记帐完成后是否自动发料
Public gbln血库系统 As Boolean  '是否安装血库系统
Public gbln报警包含划价费用 As Boolean '记帐报警包含划价费用
Public gintRXCount As Integer '门诊处方限制条数

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
'参数：blnForceNum=当为Null时，是否强制表示为数字型
    '为了保持代码一致性与调用的方便性，封装gobjComlib.ZVal
    ZVal = gobjComlib.ZVal(varValue, blnForceNum)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    '该函数无法进行再次封装，请注意保持与gobjComlib.Decode的一致性
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, Optional blnShowZero As Boolean = True) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    '为了保持代码一致性与调用的方便性，封装gobjComlib.FormatEx
    FormatEx = gobjComlib.FormatEx(vNumber, intBit, blnShowZero)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    '为了保持代码一致性与调用的方便性，封装gobjComlib.Nvl
    Nvl = gobjComlib.Nvl(varValue, DefaultValue)
End Function

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim strTmp As String
    gstrLike = IIF(gobjComlib.zlDatabase.GetPara("输入匹配") = "0", "%", "")
    gbytCode = Val(gobjComlib.zlDatabase.GetPara("简码方式"))
 
    '指定药房时限制库存
    gblnStock = Val(gobjComlib.zlDatabase.GetPara(18, glngSys)) <> 0
        
    '药品按规格下医嘱
    gbln药品按规格下医嘱 = Val(gobjComlib.zlDatabase.GetPara(69, glngSys)) = 1
    
    '输液配置中心(性质为“配制中心”的药房)
    gstr输液配置中心 = Get输液配置中心

    '门诊一卡通,项目执行前必须先收费或先记帐审核
    gbln执行前先结算 = Val(gobjComlib.zlDatabase.GetPara(163, glngSys)) <> 0
    
    gbytDec = Val(gobjComlib.zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    gbytDecPrice = Val(gobjComlib.zlDatabase.GetPara(157, glngSys, , 5))
    gstrDecPrice = "0." & String(gbytDecPrice, "0")
    '检验医嘱发送时生成条形码
    gbln发送生成条形码 = Val(gobjComlib.zlDatabase.GetPara(143, glngSys)) <> 0
    
    '分批药品出库方式
    gbytMediOutMode = Val(gobjComlib.zlDatabase.GetPara(150, glngSys))
    '从属项目汇总计算折扣
    gbln从项汇总折扣 = Val(gobjComlib.zlDatabase.GetPara(93, glngSys)) <> 0
    
    '医嘱发送生成划价单的类别
    gstr住院发送划价单 = gobjComlib.zlDatabase.GetPara(80, glngSys)
    gstr门诊发送划价单 = gobjComlib.zlDatabase.GetPara(86, glngSys)
    '一卡通消费验证
    strTmp = gobjComlib.zlDatabase.GetPara(28, glngSys) & "|"
    gdbl预存款消费验卡 = Val(Split(strTmp, "|")(0))
    '住院自动发料
    gbyt住院自动发料 = Val(gobjComlib.zlDatabase.GetPara(63, glngSys))
    '门诊自动发料
    gbln门诊自动发料 = Val(gobjComlib.zlDatabase.GetPara(92, glngSys)) <> 0
    '是否安装血库系统
    gbln血库系统 = gobjComlib.Sys.IsSysSetUp(2200)
    '记帐报警包含划价费用
    gbln报警包含划价费用 = Val(gobjComlib.zlDatabase.GetPara(98, glngSys)) <> 0
    '门诊处方条数限制
    gintRXCount = Val(gobjComlib.zlDatabase.GetPara(56, glngSys))
    
    InitSysPar = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = Nvl(rsTmp!简码)
            UserInfo.姓名 = Nvl(rsTmp!姓名)
            UserInfo.部门ID = Nvl(rsTmp!部门ID, 0)
            UserInfo.部门码 = Nvl(rsTmp!部门码)
            UserInfo.部门名 = Nvl(rsTmp!部门名)
            UserInfo.性质 = Get人员性质
            UserInfo.专业技术职务 = Nvl(rsTmp!专业技术职务)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.用户名
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共部件相关对象
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-15 15:34:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    'Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.ZLCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function

Public Function Between(X, a, B) As Boolean
'功能：判断x是否在a和b之间
    '为了保持代码一致性与调用的方便性，封装gobjComlib.Between
    Between = gobjComlib.Between(X, a, B)
End Function

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    '为了保持代码一致性与调用的方便性，封装gobjComlib.IntEx
    IntEx = gobjComlib.IntEx(vNumber)
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = gobjComlib.GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIF(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, "InitObjLis"
                Set gobjLIS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub
