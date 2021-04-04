Attribute VB_Name = "mdl兴成"
Option Explicit
Private mblnInit As Boolean     '是否已经初始化
Private Const SP_STR = "|" & vbTab & "|"
Public blnmxb As Boolean      '陈宏悦20050402添加变量，为了判断慢性病以何种方式就医进行结算

Public Enum 业务类型_兴成
    兴成_政策机服务启动 = 0
    兴成_政策机服务停止
    兴成_POS机启动
    兴成_POS机停止
    兴成_获取持卡人信息
    兴成_JbylReadIC
    兴成_门诊费用预分解
    兴成_普通门诊支付确认
    兴成_门诊退费函数
    兴成_住院费用预分解
    兴成_住院支付确认
End Enum
Private Type InitbaseInfor

    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
    
    strPath_System    As String             '系统目录,
    strPath_Get       As String
    strPath_Put         As String
    strPath_In      As String
    strPath_Out     As String
    ODBC_NAME       As String
    ODBC_UserName   As String
    ODBC_PassWord   As String
    存在读卡器      As Boolean
    启用政策审核    As Boolean            '陈宏悦于20050408修改增加
    医院级别        As String
End Type

Public InitInfor_兴成 As InitbaseInfor

Private Type 病人身份
    IC卡号              As String
    社会保障号          As String
    身份证号            As String
    姓名                As String
    性别                As String
    人员类别            As String
    个人帐户余额        As Double
    保险编号            As String
    本年统筹支付        As Double         '本年统筹基本支付累计
    共付段累计          As Double       '共付段金额累计
    参保日期            As String
    统筹共付段累计      As Double       '统筹基金共付段金额累计
    慢性病费用累计      As Double       '本年慢性病费用累计
    统筹支付累计        As Double      '统筹支付累计
    
    '陈宏悦修改增加，由于医保中心增加异地就医卡
    异地卡标志          As String         '1:是异地卡 ;0: 非异地卡

    累计划入个人帐户    As Double       '累计划入个人帐户余额
    上次充卡日期        As String
    帐户支出累计        As Double       '本年个人帐户支出累计
    当前余额            As Double
    上次门诊日期        As String      '上次门诊就医日期
    本年住院次数        As Long
    住院更新日期        As String       '住院信息更新日期
    慢病标识            As String       '慢性病患者标识
    慢性病病种          As String       '慢性病病种
    慢性有效日期        As String       '慢性病有效期截止日期
    出生日期            As String
    年龄                As Long
    费用总额            As Double
    病种编码            As String
    入院类别            As String
    住院类别            As String
    出院类别            As String
    卡状态              As String
    结帐ID              As Long         '当前结帐ID值
    病人ID              As Long
    费用发生日期        As String
    异地医院            As String
    异地医院级别        As String
    
End Type


Private Type 结算数据
        文件记录条数    As Long
        交易费用总额    As Double
        个人应付总额    As Double
        个人帐户余额    As Double
        非医保范围额    As Double
        医保范围金额    As Double
        药品甲类金额    As Double
        乙类药品医保额  As Double
        乙类药品自负额  As Double
        自费药品金额    As Double
        诊疗医保金额    As Double
        特检特治金额    As Double
        特检特治自负    As Double
        诊疗非医保额    As Double
        服务医保金额    As Double
        服务非医保额    As Double
        特殊病标识      As String
        '住院多的
        统筹支付金额    As Double
        IC卡号          As String
        入院类别        As String
        医院级别        As String
        病种            As String
        起付额          As Double
        共付段金额      As Double
        封顶线以上自负金额  As Double
        本年度统筹支付累计  As Double
        本年度共付段额累计  As Double
        本年住院次数        As Long
End Type
Private g结算数据 As 结算数据

Public g病人身份_兴成 As 病人身份
Public gcnOracle_兴成 As ADODB.Connection     '中间库连接
Public gcnSQLSEVER_兴成 As ADODB.Connection     '连接到医保中心的数据库

Private gbln检查连接 As Boolean
Private gbln已经初始 As Boolean             '已经被初始化了.

'1.服务启动；用于启动政策审核系统上机登陆服务
Private Declare Function StartPolicy Lib "XCFYFJXT.DLL" ( _
        ByVal str系统目录 As String, ByVal str医院代码 As String, _
        ByVal strODBC_Name As String, ByVal ODBC用户名 As String, ByVal ODBC用户口令 As String) As Long
'===============================================================================================================
'原型:
'功能: 服务启动；用于启动政策审核系统上机登陆服务
'入口参数:
'       1.系统目录 (完整)
'       2.医院代码
'       3.ODBC数据源名字
'出口参数: 无
'返回: 0    ――启动政策机成功
'       100              ――系统目录错误
'       101              ――医疗机构代码错误
'       102              ――连接数据库错误
'       －11 ――未经正确授权
'说明: 读卡、政策审核、费用分解、费用支付之前必须启动政策机。可设定为实例变量，启动一次即可。
'===============================================================================================================

'2.服务终止；用于停止政策审核系统登陆服务
Private Declare Function StopPolicy Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'原型:
'功能: 服务停止
'入口参数:无
'出口参数: 无
'返回:  0   ――停止政策机成功
'       103         ――断开数据库连接错误
'说明: 一般地，关闭与政策机相关的窗口之前需调用此函数
'===============================================================================================================

'3. POS机服务启动；用于启动POS机（读卡器）登陆服务、POS（读卡器）初始化等
    Private Declare Function StartPos Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'原型:
'功能: ：POS机服务启动；用于启动POS机（读卡器）登陆服务、POS（读卡器）初始化等
'入口参数:无
'出口参数: 无
'返回: 0   ――启动POS机成功
'            非0 ――启动POS机错误
'说明:  调用函数 StartPolicy成功后需调用此函数
'===============================================================================================================


'4. POS服务终止；用于停止POS登记服务
Private Declare Function StopPos Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'原型:
'功能: ：POS服务终止；用于停止POS登记服务
'入口参数:无
'出口参数: 无
'返回: ：0   ――停止POS机成功
'            非0 ――停止POS机错误
'说明:  调用函数 StartPolicy成功后需调用此函数
'===============================================================================================================

'5. 本函数用于获取POS及卡片状态，并获取持卡人信息
Private Declare Function GetPersonCommInfo Lib "XCFYFJXT.DLL" (ByVal str病人信息 As String) As Long
'===============================================================================================================
'原型:
'功能: ：本函数用于获取POS及卡片状态，并获取持卡人信息
'入口参数:无
'出口参数:
'       IC卡号|公民身份号码|姓名|性别|医疗参保人员类别|个人帐户余额|保险编号|本年统筹基金支付累计|共付段金额累计
'返回: ：0――成功
'        8――已进入黑名单
'        9――划拨累计金额大于工资总额，即为大额医疗卡，需要没收交医保中心处理。
'        其他非0值:                         错误
'说明:  门诊收费、住院登记之前需要调用此函数。
'===============================================================================================================


'6. 本函数用于读取IC卡中的医疗保险信息
'修改接口函数名称，由于此函数报错，由Read_ic修改为JbylReadIC

Private Declare Function JbylReadIC Lib "XCFYFJXT.DLL" (ByVal str病人信息 As String) As Long
'===============================================================================================================
'原型:
'功能: 本函数用于读取IC卡中的医疗保险信息。
'入口参数:无
'出口参数: 社会保障号|卡号|人员类别|参保日期|统筹基金共付段金额累计|本年慢性病费用累计|统筹支付累计|累计划入个人帐户余额|上次充卡日期|本年个人帐户支出累计|当前余额|上次门诊就医日期|本年住院次数|住院信息更新日期|慢性病患者标识|慢性病病种|慢性病有效期截止日期
'返回: 0    ――成功
'      非0 ――失败
'说明:
'===============================================================================================================


'7. 门诊费用预分解
Private Declare Function Poli_Divide Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'原型:
'功能: 本函数用于对门诊详细信息进行费用预分解，并产生个人应付金额

'入口参数:无
'           位于指定的目录中的:Poli_Divide.in文件
'出口参数: 无
'           位于指定目录中的:Poli_Divide.out，Poli_Divide.out.log文件
'返回:  0――成功
'       1――Poli_Divide.in文件入口参数错误
'       104――Poli_Divide.in或Poli_Divide.out?Poli_Divide.log文件打不开
'       105――写出口参数错误
'       -11――未经正确授权
'       其他非0参数――处理失败
'说明:
'    1） 调用此函数须保证政策机已成功启动。
'    2） 调用此函数须确保在系统目录下已建立\in及\out目录。
'    医疗机构的管理系统中，调用此函数前，向\in目录下写入Poli_Divide.in文件。本函数本身不带参数，入口参数从.in文件中读取。
'    3） 调用成功后，政策审核系统在\out目录下写Poli_Divide.out、Poli_Divide.log文件，医疗机构的管理系统读取Poli_Divide.out文件用来更新上传接口表数据，并准备门诊支付函数的入口参数。
'    4） 如果预分解失败，请查看Poli_Divide.log文件。
'===============================================================================================================


'8. 普通门诊支付确认
Private Declare Function Reg_Poli Lib "XCFYFJXT.DLL" (ByVal StrInput As String, ByVal strOutput As String) As Long
'===============================================================================================================
'原型:
'功能: 本函数用于对门诊详细信息进行费用预分解，并产生个人应付金额

'入口参数:
'   交易流水号|IC卡号|费用总金额|医保内总费用|医保外总费用|个人帐户支付金额|现金支付金额
'出口参数:
'          交易流水号|IC卡号|终端机编号|交易日期/时间|费用总金额|医保内总费用|医保外总费用|个人帐户支付金额|现金支付金额|扣减后个人帐户余额|MAC1
'返回: 0――成功
'       1――Poli_Divide.in文件入口参数错误
'       －11――未经正确授权
'       其他非0参数――处理失败
'说明:
'    1） 调用此函数须保证政策机已成功启动。
'    2） 入口参数由调用预分解函数后的参数计算得出
'    3） 调用预分解函数后调用此函数
'===============================================================================================================

'9. 门诊退费函数
Private Declare Function PoliBackCost Lib "XCFYFJXT.DLL" (ByVal StrInput As String) As Long
'===============================================================================================================
'原型:
'功能: 函数用于完成门诊退费功能，并产生交易记录

'入口参数:
'   交易流水号|IC卡号|费用总金额|医保内总费用|医保外总费用|个人帐户支付金额|现金支付金额
'出口参数:
'          交易流水号|IC卡号|终端机编号|交易日期/时间|费用总金额|医保内总费用|医保外总费用|个人帐户支付金额|现金支付金额|扣减后个人帐户余额|MAC1
'返回: 0――成功
'       －11――未经正确授权
'       其他非0参数――处理失败
'说明:
'===============================================================================================================

'10. 住院费用预分解
Private Declare Function Hosp_Divide Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'原型:
'功能: 函数用于完成门诊退费功能，并产生交易记录

'入口参数:无
'           位于指定的目录中的:Hosp_Divide.in文件
'出口参数: 无
'           位于指定目录中的:Hosp_Divide.out, Hosp_Divide_out.log文件
'返回:  0   ――成功
'       1――Hosp_Divide.in文件入口参数错误
'       104――Hosp_Divide.in或Hosp_Divide.out?Hosp_Divide.log文件打不开
'       105――写出口参数错误
'       －11――未经正确授权
'       其他非0参数――处理失败,有些错误号由数据库厂商返回，一般分解错误与分解政策参数有关
'说明:
'    1)调用此函数须保证政策机已成功启动。
'    2)调用此函数须确保在系统目录下已建立\in及\out目录。
'    3)医疗机构的管理系统中，调用此函数前，向\in目录下写入Hosp_Divide.in文件。本函数本身不带参数，入口参数从.in文件中读取。
'    4)调用成功后，政策审核系统在\out目录下写Hosp_Divide.out、Hosp_Divide.log文件，医疗机构的管理系统读取Hosp_Divide.out文件用来更新上传接口表数据，并准备住院支付函数的入口参数。
'    5) 如果预分解失败，请查看Hosp_Divide.log文件
'===============================================================================================================



'11. 住院支付确认
Private Declare Function Reg_Hospital Lib "XCFYFJXT.DLL" (ByVal StrInput As String, ByVal strOutput As String) As Long
'===============================================================================================================
'原型:
'功能: 此函数用于完成住院收费，并产生交易记录

'入口参数:
'       交易流水号|IC卡号|费用总金额|医保内总费用|医保外总费用|统筹支付金额|统筹自负金额|个人帐户支付金额|现金支付金额
'出口参数:
'       交易流水号|IC卡号|终端机编号|交易日期/时间|费用总金额|医保内总费用|医保外总费用|统筹支付金额|统筹自负金额|个人帐户支付金额|现金支付金额|扣减后个人帐户余额|MAC1|
'返回:  0――成功
'       1――Poli_Divide.in文件入口参数错误
'       －11――未经正确授权
'       其他非0参数――处理失败
'说明:
'    1） 调用此函数须保证政策机已成功启动。
'    2） 入口参数由调用预分解函数后的参数计算得出。
'    3） 调用住院预分解函数后调用此函数。
'===============================================================================================================


Private Function Get年龄(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as 年龄 from dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取年龄")
    If Not rsTemp.EOF Then
        Get年龄 = Int(Nvl(rsTemp!年龄, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function


Public Function 医保初始化_兴成() As Boolean
    
    Dim strReg As String, strOutput As String, StrInput As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If mblnInit = True Then
        医保初始化_兴成 = True
        Exit Function
    End If
    
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_兴成核工业 & " and 参数名='医院级别'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医院级别"
        
    If rsTemp.EOF Then
        ShowMsgbox "未设置医院级别,请在参数中设置"
        Exit Function
    End If
    InitInfor_兴成.医院级别 = Nvl(rsTemp!参数值)
    
    If InitInfor_兴成.医院级别 = "" Then
        ShowMsgbox "未设置医院级别,请在参数中设置"
        Exit Function
    End If
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_兴成.模拟数据 = True
    Else
        InitInfor_兴成.模拟数据 = False
    End If
   
    InitInfor_兴成.医院编码 = gstr医院编码

    InitInfor_兴成.strPath_Get = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Get"), "C:\xcyb\get")
    InitInfor_兴成.strPath_Put = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Put"), "C:\xcyb\Put")
    InitInfor_兴成.strPath_In = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_In"), "C:\xcyb\In")
    InitInfor_兴成.strPath_Out = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Out"), "C:\xcyb\Out")
    InitInfor_兴成.strPath_System = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_System"), "C:\")
    
    InitInfor_兴成.ODBC_NAME = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_NAME"), "")
    InitInfor_兴成.ODBC_UserName = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_USERNAME"), "")
    InitInfor_兴成.ODBC_PassWord = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_PASSWORD"), "")
    
    
    InitInfor_兴成.存在读卡器 = Val(GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("读卡器"), "1")) = 1
    
    '陈宏悦于200500408增加
    
    InitInfor_兴成.启用政策审核 = Val(GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("启用政策审核"), "1")) = 1
        
    '启动正策服务,相关参数
    '    a. 系统目录（完整）[50] ――类型：String       值：一般指的是c:\
    '    b. 医院代码 [8]         ――类型：String       长度为8位
    '    c. ODBC数据源名字[ ]   ――类型：String        值：一般指的是ODBC的DSN
    '    d. ODBC用户名[ ]       ――类型：String
    '    e. ODBC用户口令[ ]     ――类型：String
    
    StrInput = InitInfor_兴成.strPath_System
    StrInput = StrInput & SP_STR & InitInfor_兴成.医院编码
    StrInput = StrInput & SP_STR & InitInfor_兴成.ODBC_NAME
    StrInput = StrInput & SP_STR & InitInfor_兴成.ODBC_UserName
    StrInput = StrInput & SP_STR & InitInfor_兴成.ODBC_PassWord
    
     '陈宏悦于20050408增加修改,目的:解决住院记帐是否启用政策审核动态库
      
      If InitInfor_兴成.启用政策审核 Then
      
        If 业务请求_兴成(兴成_政策机服务启动, StrInput, strOutput) = False Then
            Exit Function
        End If
        
      End If
      
    If InitInfor_兴成.存在读卡器 Then
        '需启动POS机服务
        If 业务请求_兴成(兴成_POS机启动, "", "") = False Then Exit Function
    End If
    
    If Open中间库_兴成 = False Then
        Exit Function
    End If
    mblnInit = True
    医保初始化_兴成 = True
End Function

Public Function 医保终止_兴成() As Boolean
    
    
    '将初始化标志置为false
    mblnInit = False
    
 '陈宏悦于20050408增加修改,目的:解决住院记帐的问题
 
 If InitInfor_兴成.启用政策审核 Then
 
    Call 业务请求_兴成(兴成_政策机服务停止, "", "")
        
    If gcnOracle_兴成.State = 1 Then
        gcnOracle_兴成.Close
    End If
    If gcnSQLSEVER_兴成.State = 1 Then
        gcnSQLSEVER_兴成.Close
    End If
    Call 业务请求_兴成(兴成_POS机停止, "", "")
    
    医保终止_兴成 = True
    
  End If
  
End Function

Public Function 身份标识_兴成(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Err = 0
    On Error GoTo errHand:
    'If bytType = 1 Or bytType = 3 or  Then Exit Function
    
    身份标识_兴成 = frmIdentify兴成.GetPatient(bytType, lng病人ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_兴成 = ""
End Function

Public Function 个人余额_兴成(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_兴成核工业)
    
    If rsTemp.EOF Then
        个人余额_兴成 = 0
    Else
        个人余额_兴成 = rsTemp("帐户余额")
    End If
End Function
Private Function WriteINParaFile(ByVal rs明细 As ADODB.Recordset, Optional bln门诊虚拟结算 As Boolean = False, Optional bln住院 As Boolean = False) As Boolean
    '将明细数据写入文本文件中
    Dim str处方号 As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim strFile As String, StrInput As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim str流水号 As String
    
    WriteINParaFile = False
    If objFile.FolderExists(InitInfor_兴成.strPath_In) = False Then
        ShowMsgbox "未创建该文件夹(" & InitInfor_兴成.strPath_In & "),请创建!"
        Exit Function
    End If
    If Not bln住院 Then
        strFile = InitInfor_兴成.strPath_In & "\Poli_Divide.in"
    Else
        strFile = InitInfor_兴成.strPath_In & "\Hosp_Divide.in"
    End If
    
    Err = 0: On Error GoTo errHand:
    If Not Dir(strFile) <> "" Then
        
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, ForWriting)
        
  
    If bln门诊虚拟结算 Then
        str处方号 = Int(Rnd * 1000000000000#)
    Else
         rs明细.MoveFirst
        str处方号 = Nvl(rs明细!NO, 0)
    End If
    
    If bln住院 Then
        
        '陈宏悦于20050311修改，Hosp_Divide()入口参数传入错误，由于Hosp_Divide.in文件中第一行没有数据
        
        '第一行:
        'IC卡号|文件记录条数|本次交易费用总金额|入院类别|医院级别|病种
        StrInput = g病人身份_兴成.IC卡号 & "|"
        StrInput = StrInput & rs明细.RecordCount & "|"
        StrInput = StrInput & Format(g病人身份_兴成.费用总额 * 100, "#####0.00;-#####0.00;0;0") & "|"
        StrInput = StrInput & g病人身份_兴成.入院类别 & "|"
        StrInput = StrInput & IIf(g病人身份_兴成.出院类别 = "3", g病人身份_兴成.异地医院级别, InitInfor_兴成.医院级别) & "|"
        StrInput = StrInput & g病人身份_兴成.病种编码
    Else
        '第一行:文件记录条数|本次交易费用总金额
        StrInput = rs明细.RecordCount & "|"
        StrInput = StrInput & Format(g病人身份_兴成.费用总额 * 100, "#####0.00;-#####0.00;0;0")
    End If
    objText.WriteLine StrInput
    
    With rs明细
        .MoveFirst
        Do While Not .EOF
            gstrSQL = "" & _
                "   Select a.*,b.编码,b.名称 " & _
                "   From 保险支付项目 a,收费细目 B  " & _
                "   where a.收费细目id=b.ID and  a.险类=[2]" & _
                "           and a.收费细目id=[1]"
                
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "确定医保支付项目", CLng(Nvl(!收费细目ID, 0)), TYPE_兴成核工业)
            
            '由于中心不需要判断是否医保项目,陈宏悦于20050321修改
            
'            If rsTemp.EOF Then
'                ShowMsgbox "收费项目“" & Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称) & "”还未进行医保对码，不能进行结算!"
'                Exit Function
'            End If
            
            If Nvl(!数量, 0) - Int(Nvl(!数量, 0)) > 0 Then
                ShowMsgbox "收费项目“" & Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称) & "”输入的数量为小数了，不能进行结算!"
                Exit Function
            End If
            StrInput = ""
            
            If bln门诊虚拟结算 Then
                str流水号 = Rpad(Substr(InitInfor_兴成.医院编码, 1, 8), 8, " ") & Lpad(.AbsolutePosition, 12, "0")
            Else
                str流水号 = Rpad(Substr(InitInfor_兴成.医院编码, 1, 8), 8, " ") & Lpad(.AbsolutePosition, 12, "0")
            End If
            
            StrInput = str流水号
            If bln住院 Then
                '交易流水号|项目序号|项目代码|项目名称|项目类别|单价|数量|费用总金额|费用发生日期
                StrInput = StrInput & "|" & .AbsolutePosition
            Else
                '交易流水号|处方号|项目代码|项目名称|项目类别|单价|数量|费用总金额|病种代码
                StrInput = StrInput & "|" & str处方号
            End If
            
            '陈宏悦于20050321修改,非医保项目传给中心一个0
            If rsTemp.EOF Then
                gstrSQL = "" & _
                "   Select b.类别,b.编码,b.名称 " & _
                "   From 收费细目 B  " & _
                "   where  B.id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "确定非医保项目名称", CLng(Nvl(!收费细目ID, 0)))
                StrInput = StrInput & "|0"
                StrInput = StrInput & "|" & Substr(Nvl(rsTemp!名称), 1, 20)
                StrInput = StrInput & "|" & Substr(Decode(rsTemp!类别, "5", 0, "6", 0, "7", 0, "J", 2, 1), 1, 1)
            Else
              StrInput = StrInput & "|" & Nvl(rsTemp!项目编码)
              StrInput = StrInput & "|" & Substr(Nvl(rsTemp!名称), 1, 20)
              StrInput = StrInput & "|" & Substr(Nvl(rsTemp!附注), 1, 1)
            End If
              
              StrInput = StrInput & "|" & Format(Nvl(!单价, 0) * 10000, "######;-#####;0;0")
              StrInput = StrInput & "|" & Format(Nvl(!数量, 0), "######;-#####")
              StrInput = StrInput & "|" & Format(Nvl(!实收金额, 0) * 100, "######0.00;-#####0.00")
            
            
            If bln住院 Then
                StrInput = StrInput & "|" & Format(!发生时间, "YYYYMMDD")
            Else
                '陈宏悦于20050321修改，原因：由于存在慢性病
                           
                If g病人身份_兴成.慢病标识 = "1" And blnmxb = True Then
                    StrInput = StrInput & "|" & g病人身份_兴成.病种编码  '由于当时咨询李凯，由于暂无慢病，所以传0000
                Else
                   StrInput = StrInput & "|0000"
                End If
                
            End If
            objText.WriteLine StrInput
            
            If bln门诊虚拟结算 Or bln住院 Then
            Else
                '为病人费用记录打上标记，以便随时上传
                 'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                 gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str流水号 & "')"
                 DebugTool "     打上明细标志:SQL=" & gstrSQL
                 zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
                 DebugTool " 打上明细标志:更新病人费用记录成功:SQL=" & gstrSQL
            End If
            .MoveNext
        Loop
    End With
    objText.Close
    WriteINParaFile = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    objText.Close
End Function
Private Function ReadOutParaFile(ByRef obj结算数据 As 结算数据, _
    Optional bln门诊 As Boolean = True, Optional bln汇总 As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对出参文件的内容进行分解
    '--入参数:bln汇总-只读汇总信息
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strFile As String, StrInput As String, strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim lngRow As Long
    Dim strArr As Variant
    Dim strSQL As String
    
    ReadOutParaFile = False
    If objFile.FolderExists(InitInfor_兴成.strPath_Out) = False Then
        ShowMsgbox "未创建该文件夹(" & InitInfor_兴成.strPath_Out & "),请创建!"
        Exit Function
    End If
    
    If bln门诊 Then
        strFile = InitInfor_兴成.strPath_Out & "\Poli_Divide.out"
    Else
        strFile = InitInfor_兴成.strPath_Out & "\Hosp_Divide.out"
    End If
    
    Err = 0: On Error GoTo errHand:
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "没有产生相关的出参文件：" & strFile & vbCrLf & " 请检查!"
        Exit Function
    End If
    Set objText = objFile.OpenTextFile(strFile, ForReading)
    
    lngRow = 1
    Do While Not objText.AtEndOfStream
          strText = Trim(objText.ReadLine)
          If strText = "" Then Exit Function
          
          strArr = Split(strText, "|")
          If lngRow = 1 Then
                
                '读汇总信息
                '门诊:                                 文件记录条数|本次交易费用总金额             |个人应付总金额|个人帐户余额|非医保范围金额|医保范围金额|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|特殊病标识
                '住院:   IC卡号|入院类别|医院级别|病种|文件记录条数|本次交易费用总金额|统筹支付金额|个人应付总金额|个人帐户余额|非医保范围金额|医保范围金额|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|起付额|共付段金额|封顶线以上自负金额|本年度统筹基金支付累计|本年度共付段金额累计|本年住院次数
                If bln门诊 Then
                    
                    With obj结算数据
                        .文件记录条数 = Val(strArr(0))
                        .交易费用总额 = Format(Val(strArr(1)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .个人应付总额 = Format(Val(strArr(2)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .个人帐户余额 = Format(Val(strArr(3)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .非医保范围额 = Format(Val(strArr(4)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .医保范围金额 = Format(Val(strArr(5)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .药品甲类金额 = Format(Val(strArr(6)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .乙类药品医保额 = Format(Val(strArr(7)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .乙类药品自负额 = Format(Val(strArr(8)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .自费药品金额 = Format(Val(strArr(9)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .诊疗医保金额 = Format(Val(strArr(10)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .特检特治金额 = Format(Val(strArr(11)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .特检特治自负 = Format(Val(strArr(12)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .诊疗非医保额 = Format(Val(strArr(13)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .服务医保金额 = Format(Val(strArr(14)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .服务非医保额 = Format(Val(strArr(15)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .特殊病标识 = strArr(16)
                        .IC卡号 = ""
                        .入院类别 = ""
                        .医院级别 = ""
                        .病种 = ""
                        .统筹支付金额 = 0
                        .起付额 = 0
                        .共付段金额 = 0
                        .封顶线以上自负金额 = 0
                        .本年度统筹支付累计 = 0
                        .本年度共付段额累计 = 0
                        .本年住院次数 = 0
                    End With
                Else
                    '住院
                    With obj结算数据
                        .IC卡号 = strArr(0)
                        .入院类别 = strArr(1)
                        .医院级别 = strArr(2)
                        .病种 = strArr(3)
                        .文件记录条数 = Val(strArr(4))
                        .交易费用总额 = Format(Val(strArr(5)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .统筹支付金额 = Format(Val(strArr(6)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .个人应付总额 = Format(Val(strArr(7)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .个人帐户余额 = Format(Val(strArr(8)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .非医保范围额 = Format(Val(strArr(9)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .医保范围金额 = Format(Val(strArr(10)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .药品甲类金额 = Format(Val(strArr(11)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .乙类药品医保额 = Format(Val(strArr(12)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .乙类药品自负额 = Format(Val(strArr(13)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .自费药品金额 = Format(Val(strArr(14)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .诊疗医保金额 = Format(Val(strArr(15)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .特检特治金额 = Format(Val(strArr(16)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .特检特治自负 = Format(Val(strArr(17)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .诊疗非医保额 = Format(Val(strArr(18)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .服务医保金额 = Format(Val(strArr(19)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .服务非医保额 = Format(Val(strArr(20)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .起付额 = Format(Val(strArr(21)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .共付段金额 = Format(Val(strArr(22)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .封顶线以上自负金额 = Format(Val(strArr(23)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .本年度统筹支付累计 = Format(Val(strArr(24)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .本年度共付段额累计 = Format(Val(strArr(25)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .本年住院次数 = Format(Val(strArr(26)), "#####0.00;-#####0.00;0.00;0.00")
                        .特殊病标识 = ""
                    End With
                End If
                If bln汇总 Then
                    Exit Do
                End If
        End If
        If bln门诊 Then
            '门诊需插入数据
            '保存相关的汇总和明细数据
            If InsertIntoYBK(strArr, lngRow - 1, lngRow <> 1, bln门诊) = False Then
                Exit Function
            End If
        Else
            '住院不需插入,单独插入数据.在正试结算时才传入.
        End If
        lngRow = lngRow + 1
    Loop
    ReadOutParaFile = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InsertIntoSQLServer_门诊(ByVal str确认串 As String) As Boolean
    '插入数据到SQLSErver中.
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对出参文件的内容进行分解
    '--入参数:str确认串-门诊确认支付时传出的串
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strFile As String, StrInput As String, strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim lngRow As Long
    Dim strArr As Variant
    Dim strArr1 As Variant
    Dim strSQL As String
    Dim str开单部门 As String
    Dim str开单医师 As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select b.名称 as 开单部门,开单人 from 门诊费用记录 a,部门表 b where a.开单部门id=b.id(+) and 结帐id=" & g病人身份_兴成.结帐ID & " and rownum=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取开单人"
    If rsTemp.EOF Then
        str开单部门 = ""
        str开单医师 = ""
    Else
        str开单部门 = Substr(Nvl(rsTemp!开单部门), 1, 12)
        str开单医师 = Substr(Nvl(rsTemp!开单人), 1, 4)    '由于导出的数据库的结果为4位,但文档中是12位,请到时更改/
    End If
    
    InsertIntoSQLServer_门诊 = False
    
    If objFile.FolderExists(InitInfor_兴成.strPath_Out) = False Then
        ShowMsgbox "未创建该文件夹(" & InitInfor_兴成.strPath_Out & "),请创建!"
        Exit Function
    End If
    
  '  MsgBox str确认串, vbOKOnly, "zlsoft"
    
    strArr1 = Split(str确认串, "|")
    
'    MsgBox strArr1(5), vbOKOnly, "zlsoft"
'    MsgBox strArr1(0), vbOKOnly, "zlsoft"
'
'    MsgBox strArr1(6), vbOKOnly, "zlsoft"
    
    strFile = InitInfor_兴成.strPath_Out & "\Poli_Divide.out"
    
    Err = 0: On Error GoTo errHand:
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "没有产生相关的出参文件：" & strFile & vbCrLf & " 请检查!"
        Exit Function
    End If
    Set objText = objFile.OpenTextFile(strFile, ForReading)
    
    lngRow = 1
    Do While Not objText.AtEndOfStream
          strText = Trim(objText.ReadLine)
          If strText = "" Then Exit Do
          strArr = Split(strText, "|")
          
          If lngRow = 1 Then
                
                '读汇总信息
                '门诊:文件记录条数|本次交易费用总金额|个人应付总金额|个人帐户余额|非医保范围金额|医保范围金额|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|特殊病标识
                '门诊传入: 交易流水号|IC卡号|费用总金额|医保内总费用|医保外总费用|个人帐户支付金额|现金支付金额
                '表结构:vc_jyh(门诊收费流水号),vc_date(费用发生日期),C_miid(医保编号),vc_cardid(卡号),Vc_jzh(挂号流水号),
                '    C_yljgdm(医疗机构代码),C_ksid(科室名称),C_doctorid(开单医师),Vc_bzid(病种),
                '    N_sum(总金额),N_nyb(非医保范围金额),N_yb(医保范围金额),N_grzf(个人帐户支付金额),N_xzjf(现金支付金额（帐户不足支付）),
                '    N_drug_a(药品甲类金额),N_drug_byb(乙类药品医保金额),N_drug_bzf(乙类药品自负金额),N_drug_zf(自费药品金额),
                '    N_drug_mi(诊疗医保金额),N_drug_s(特检特治医保金额),N_drug_szf(特检特治自负金额),N_zlnyb(诊疗非医保金额),N_fwyb(服务医保金额),N_fwnyb(服务非医保金额),C_s_flag(特殊病标志),
                '    C_jzys (接诊医师), C_fscfh(复式处方号), C_qzysdm(签字医师代码), C_wpcfyy(外配处方医院), C_zt(传输状态)

                strSQL = "insert into HLD_MZJYXX(vc_jyh,vc_date,C_miid,vc_cardid,Vc_jzh,C_yljgdm,C_ksid,C_doctorid,Vc_bzid,N_sum,N_nyb,N_yb,N_grzf,N_xjzf,N_drug_a,N_drug_byb,N_drug_bzf,N_drug_zf,N_drug_mi,N_drug_s,N_drug_szf,N_zlnyb,N_fwyb,N_fwnyb,C_s_flag,C_zt) values("
                strSQL = strSQL & "'" & strArr1(0) & "',"  'vc_jyh(门诊收费流水号)
                strSQL = strSQL & "'" & Format(zlDatabase.Currentdate, "yyyymmddHHMMSS") & "'," 'vc_date(费用发生日期)
                strSQL = strSQL & "'" & g病人身份_兴成.社会保障号 & "',"  'C_miid(医保编号)
                strSQL = strSQL & "'" & g病人身份_兴成.IC卡号 & "',"   'vc_cardid(卡号)
                strSQL = strSQL & "0,"   'Vc_jzh(挂号流水号)：没有实际意义，缺省为零
                strSQL = strSQL & "'" & InitInfor_兴成.医院编码 & "',"    ' C_yljgdm(医疗机构代码)
                strSQL = strSQL & "" & IIf(str开单部门 = "", "NULL", "'" & str开单部门 & "'") & ","    ' C_ksid(科室名称)
                strSQL = strSQL & "" & IIf(str开单医师 = "", "NULL", "'" & str开单医师 & "'") & ","    ' C_doctorid(开单医师)
                
                '陈宏悦于20051231修改，由于医保前置库中的字段长度不一致
                If g病人身份_兴成.慢病标识 = "1" Then
                    strSQL = strSQL & "" & IIf(g病人身份_兴成.慢性病病种 = "", "Null", "'" & g病人身份_兴成.慢性病病种 & "'") & ","    'Vc_bzid(病种)
                Else
                    strSQL = strSQL & "" & "Null" & ","    'Vc_bzid(病种)
                End If
                
                strSQL = strSQL & "" & Format(Val(strArr(1)) / 100, "####0.00;-####0.00;0;0") & "," 'N_sum(总金额)
                strSQL = strSQL & "" & Format(Val(strArr(4)) / 100, "####0.00;-####0.00;0;0") & "," 'N_nyb(非医保范围金额)
                strSQL = strSQL & "" & Format(Val(strArr(5)) / 100, "####0.00;-####0.00;0;0") & "," 'N_yb(医保范围金额)
                strSQL = strSQL & "" & Format(Val(strArr1(7)) / 100, "####0.00;-####0.00;0;0") & "," 'N_grzf(个人帐户支付金额),
                strSQL = strSQL & "" & Format(Val(strArr1(8)) / 100, "####0.00;-####0.00;0;0") & "," 'N_xzjf(现金支付金额（帐户不足支付）)
                strSQL = strSQL & "" & Format(Val(strArr(6)) / 100, "####0.00;-####0.00;0;0") & "," 'N_drug_a(药品甲类金额)
                strSQL = strSQL & "" & Format(Val(strArr(7)) / 100, "####0.00;-####0.00;0;0") & "," 'N_drug_byb(乙类药品医保金额)
                strSQL = strSQL & "" & Format(Val(strArr(8)) / 100, "####0.00;-####0.00;0;0") & "," 'N_drug_bzf(乙类药品自负金额)
                strSQL = strSQL & "" & Format(Val(strArr(9)) / 100, "####0.00;-####0.00;0;0") & "," ',N_drug_zf(自费药品金额)
                strSQL = strSQL & "" & Format(Val(strArr(10)) / 100, "####0.00;-####0.00;0;0") & "," ',N_drug_mi(诊疗医保金额)
                strSQL = strSQL & "" & Format(Val(strArr(11)) / 100, "####0.00;-####0.00;0;0") & "," ',N_drug_s(特检特治医保金额)
                strSQL = strSQL & "" & Format(Val(strArr(12)) / 100, "####0.00;-####0.00;0;0") & "," ',N_drug_szf(特检特治自负金额)
                strSQL = strSQL & "" & Format(Val(strArr(13)) / 100, "####0.00;-####0.00;0;0") & "," 'N_zlnyb(诊疗非医保金额)
                strSQL = strSQL & "" & Format(Val(strArr(14)) / 100, "####0.00;-####0.00;0;0") & "," 'N_fwyb(服务医保金额)
                strSQL = strSQL & "" & Format(Val(strArr(15)) / 100, "####0.00;-####0.00;0;0") & "," 'N_fwnyb(服务非医保金额)
                strSQL = strSQL & "'" & strArr(16) & "',"   ''C_s_flag(特殊病标志)
                strSQL = strSQL & "0)"   'C_zt 0 新数据 1 已正确传输，状态未知 2 已正确传输，并成功入库
                
                'MsgBox "进入调试设置陷阱B1！" & vbCrLf & "开单部门：" & str开单部门 & vbCrLf & "社会保障号：" & g病人身份_兴成.社会保障号 & vbCrLf & "IC卡号：" & g病人身份_兴成.IC卡号 & vbCrLf & "慢性病病种：" & g病人身份_兴成.慢性病病种, vbOKOnly, "中联软件"
                gcnSQLSEVER_兴成.Execute strSQL
        Else
        
        '明细:交易流水号|处方号|项目代码|项目名称|项目类别|单价|数量|费用总金额|医保内费用|医保外费用|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|MAC2
        '表结构:VC_RECEIPTID(门诊收费流水号),N_SFSXH(收费顺序号),VC_ITEM_ID(项目代码),VC_ITEM_NAME(项目名称),C_MIID(保险编号),VC_JZH(挂号流水号),C_SSMLJB(所属目录级别),C_ZLXMJB(诊疗项目级别),C_FWSSFW(服务设施范围),N_PRICE(单价),N_sl(数量),N_sum(费用总金额),N_DRUG_A(药品甲类金额),N_DRUG_BYB(乙类药品医保金额),N_drug_bzf(乙类药品自负金额),N_drug_zf(自费药品金额),D_DRUG_ZLYB(诊疗医保金额),N_SYB(特检特治医保金额),N_SYBZF(特检特治自负金额),N_ZLFYB(诊疗非医保金额),N_fwyb(服务医保金额),N_fwnyb(服务非医保金额),C_CF_FLAG(处方标识),C_zt(传输状态)
        strSQL = "insert into HLD_MZCFZLXX(VC_RECEIPTID,N_SFSXH,VC_ITEM_ID,VC_ITEM_NAME,C_MIID,VC_JZH,C_SSMLJB,C_ZLXMJB,C_FWSSFW,N_PRICE,N_sl,N_sum,N_DRUG_A,N_DRUG_BYB,N_drug_bzf,N_drug_zf,N_DRUG_ZLYB,N_SYB,N_SYBZF,N_ZLFYB,N_fwyb,N_fwnyb,C_CF_FLAG,C_zt) values("
        strSQL = strSQL & "'" & strArr1(0) & "',"        'VC_RECEIPTID (门诊收费流水号)
        strSQL = strSQL & "" & lngRow - 1 & ","      'N_SFSXH (收费顺序号)
        strSQL = strSQL & "'" & strArr(2) & "',"        'VC_ITEM_ID (项目代码)
        strSQL = strSQL & "'" & Substr(strArr(3), 1, 20) & "',"      'VC_ITEM_NAME (项目名称)
        strSQL = strSQL & "'" & g病人身份_兴成.保险编号 & "',"         'C_MIID (保险编号)
        
        strSQL = strSQL & "0,"        'VC_JZH (挂号流水号)      '没有实际意义
        Select Case strArr(4)
            Case "0"    '药品
                gstrSQL = "Select ssmljb as 级别 From YB_YD where xmdm='" & strArr(2) & "'"
            Case "1"    '诊疗
                gstrSQL = "Select tjtzbz as 级别 From YB_ZLML where xmdm='" & strArr(2) & "'"
            Case Else   '服务
                gstrSQL = "Select fwfw as 级别 From YB_FWSS where xmdm='" & strArr(2) & "'"
        End Select
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnSQLSEVER_兴成
        
        '陈宏悦于20050321修改,因为存在不需对码的项目
        
'        If rsTemp.EOF Then
'            ShowMsgbox "在保存结算记录时,未发现相关的医保项目代码[" & strArr(2) & "]"
'            Exit Function
'        End If
        
            Select Case strArr(4)
                 Case "0"    '药品
                    '陈宏悦于20050321修改,因为存在不需对码的项目

                     strSQL = strSQL & "'0',"           'C_SSMLJB (所属目录级别)
                     If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (诊疗项目级别)
                     Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (诊疗项目级别)
                     End If

                 Case "1"    '诊疗
                    strSQL = strSQL & "'1',"              'C_SSMLJB (所属目录级别)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"           'C_ZLXMJB (诊疗项目级别)
                    Else
                       strSQL = strSQL & "'0',"        'C_SSMLJB (所属目录级别)
                    End If

                 Case Else   '服务
                    strSQL = strSQL & "'2',"           'C_SSMLJB (所属目录级别)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (诊疗项目级别)
                    Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (诊疗项目级别)
                    End If
                    
             End Select
             
             strSQL = strSQL & "NULL,"        'C_FWSSFW (服务设施范围)
        
        
'       Select Case strArr(4)
'            Case "0"    '药品
'                '陈宏悦于20050321修改,因为存在不需对码的项目
'                 If rsTemp.EOF Then
'                    strSql = strSql & "3,"        'C_SSMLJB (所属目录级别)
'                    strSql = strSql & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                    strSql = strSql & "NULL,"        'C_FWSSFW (服务设施范围)
'                 Else
'                    strSql = strSql & "'" & Nvl(rsTemp!级别) & "',"        'C_SSMLJB (所属目录级别)
'                    strSql = strSql & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                    strSql = strSql & "NULL,"        'C_FWSSFW (服务设施范围)
'                 End If
'
'            Case "1"    '诊疗
'                 If rsTemp.EOF Then
'                    strSql = strSql & "NULL,"        'C_SSMLJB (所属目录级别)
'                    strSql = strSql & "2,"        'C_ZLXMJB (诊疗项目级别)
'                    strSql = strSql & "NULL,"        'C_FWSSFW (服务设施范围)
'                 Else
'                    strSql = strSql & "NULL,"        'C_SSMLJB (所属目录级别)
'                    strSql = strSql & "'" & Nvl(rsTemp!级别) & "',"        'C_ZLXMJB (诊疗项目级别)
'                    strSql = strSql & "NULL,"        'C_FWSSFW (服务设施范围)
'                 End If
'
'            Case Else   '服务
'                 If rsTemp.EOF Then
'                    strSql = strSql & "NULL,"        'C_SSMLJB (所属目录级别)
'                    strSql = strSql & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                    strSql = strSql & "1,"        'C_FWSSFW (服务设施范围)
'                 Else
'                    strSql = strSql & "NULL,"        'C_SSMLJB (所属目录级别)
'                    strSql = strSql & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                    strSql = strSql & "'" & Nvl(rsTemp!级别) & "',"        'C_FWSSFW (服务设施范围)
'                 End If
'
'        End Select
        
        strSQL = strSQL & "" & Format(Val(strArr(5)) / 10000, "####0.00;-####0.00;0;0") & ","         'N_PRICE (单价)
        strSQL = strSQL & "" & Format(Val(strArr(6)), "####0;-####0;0;0") & ","          'N_sl (数量)
        strSQL = strSQL & "" & Format(Val(strArr(7)) / 100, "####0.00;-####0.00;0;0") & ","        'N_sum (费用总金额)
        strSQL = strSQL & "" & Format(Val(strArr(10)) / 100, "####0.00;-####0.00;0;0") & ","        'N_DRUG_A (药品甲类金额)
        strSQL = strSQL & "" & Format(Val(strArr(11)) / 100, "####0.00;-####0.00;0;0") & ","      'N_DRUG_BYB (乙类药品医保金额)
        strSQL = strSQL & "" & Format(Val(strArr(12)) / 100, "####0.00;-####0.00;0;0") & ","        'N_drug_bzf (乙类药品自负金额)
        strSQL = strSQL & "" & Format(Val(strArr(13)) / 100, "####0.00;-####0.00;0;0") & ","       'N_drug_zf (自费药品金额)
        strSQL = strSQL & "" & Format(Val(strArr(14)) / 100, "####0.00;-####0.00;0;0") & ","       'D_DRUG_ZLYB (诊疗医保金额)
        strSQL = strSQL & "" & Format(Val(strArr(15)) / 100, "####0.00;-####0.00;0;0") & ","        'N_SYB (特检特治医保金额)
        strSQL = strSQL & "" & Format(Val(strArr(16)) / 100, "####0.00;-####0.00;0;0") & ","        'N_SYBZF (特检特治自负金额)
        strSQL = strSQL & "" & Format(Val(strArr(17)) / 100, "####0.00;-####0.00;0;0") & ","       'N_ZLFYB (诊疗非医保金额)
        strSQL = strSQL & "" & Format(Val(strArr(18)) / 100, "####0.00;-####0.00;0;0") & ","       'N_fwyb (服务医保金额)
        strSQL = strSQL & "" & Format(Val(strArr(19)) / 100, "####0.00;-####0.00;0;0") & ","       'N_fwnyb (服务非医保金额)
        strSQL = strSQL & "NULL,"        'C_CF_FLAG (处方标识)  :无实际意义
        strSQL = strSQL & "'0')"        'C_zt (传输状态)
        
 '       MsgBox strSql, vbOKOnly, "zlsoft"
        gcnSQLSEVER_兴成.Execute strSQL
                
        End If
        lngRow = lngRow + 1
    Loop
    InsertIntoSQLServer_门诊 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InsertIntoSQLServer_门诊冲销(ByVal lng冲销ID As Long) As Boolean
    '插入数据到SQLSErver中.
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:将冲销的相关数据插入SQLServer中
    '--入参数:lng冲销ID值
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim str开单部门 As String, str开单医师 As String
    Dim strSQL As String
    Dim lngRow As Long
    Dim str交易流水号 As String
    
    
    
    InsertIntoSQLServer_门诊冲销 = False
        
    Err = 0: On Error GoTo errHand:
    

    gstrSQL = "Select b.名称 as 开单部门,开单人 from 门诊费用记录 a,部门表 b where a.开单部门id=b.id(+) and 结帐id=" & g病人身份_兴成.结帐ID & " and rownum=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取开单人"
    If rsTemp.EOF Then
        str开单部门 = ""
        str开单医师 = ""
    Else
        str开单部门 = Substr(Nvl(rsTemp!开单部门), 1, 12)
        str开单医师 = Substr(Nvl(rsTemp!开单人), 1, 4)
    End If

    gstrSQL = "" & _
        "   Select  性质, 结帐id, 入院类别, 医院级别, 病种, 记录条数, 费用总额, 统筹支付金额, 个人应付总额," & _
        "           个人帐户余额, 非医保范围金额, 医保范围金额, 药品甲类金额, 乙类药品医保金额, 乙类药品自负金额, " & _
        "           自费药品金额, 诊疗医保金额, 特检特治支付额, 特检特治自负额, 诊疗非医保金额, 服务医保金额, 服务非医保金额," & _
        "           特殊病标识, 起付额, 共付段金额, 封顶线自付金额, 年度统筹基金累计, 年度共付段累计, 本年住院次数 " & _
        "   from 医保结算记录 " & _
        "   Where 性质=1 and 结帐ID =" & g病人身份_兴成.结帐ID
        
    OpenRecordset_兴成 rsData, "获取医保结算记录"
    
    gstrSQL = "Select * From 保险结算记录 where 性质=1 and 记录id=" & g病人身份_兴成.结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取结算信息"
    str交易流水号 = Nvl(rsTemp!支付顺序号)
            
    '读汇总信息
    '门诊:文件记录条数|本次交易费用总金额|个人应付总金额|个人帐户余额|非医保范围金额|医保范围金额|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|特殊病标识
    '门诊传入: 交易流水号|IC卡号|费用总金额|医保内总费用|医保外总费用|个人帐户支付金额|现金支付金额
    '表结构:vc_jyh(门诊收费流水号),vc_date(费用发生日期),C_miid(医保编号),vc_cardid(卡号),Vc_jzh(挂号流水号),
    '    C_yljgdm(医疗机构代码),C_ksid(科室名称),C_doctorid(开单医师),Vc_bzid(病种),
    '    N_sum(总金额),N_nyb(非医保范围金额),N_yb(医保范围金额),N_grzf(个人帐户支付金额),N_xzjf(现金支付金额（帐户不足支付）),
    '    N_drug_a(药品甲类金额),N_drug_byb(乙类药品医保金额),N_drug_bzf(乙类药品自负金额),N_drug_zf(自费药品金额),
    '    N_drug_mi(诊疗医保金额),N_drug_s(特检特治医保金额),N_drug_szf(特检特治自负金额),N_zlnyb(诊疗非医保金额),N_fwyb(服务医保金额),N_fwnyb(服务非医保金额),C_s_flag(特殊病标志),
    '    C_jzys (接诊医师), C_fscfh(复式处方号), C_qzysdm(签字医师代码), C_wpcfyy(外配处方医院), C_zt(传输状态)

    strSQL = "insert into HLD_MZJYXX(vc_jyh,vc_date,C_miid,vc_cardid,Vc_jzh,C_yljgdm,C_ksid,C_doctorid,Vc_bzid,N_sum,N_nyb,N_yb,N_grzf,N_xjzf,N_drug_a,N_drug_byb,N_drug_bzf,N_drug_zf,N_drug_mi,N_drug_s,N_drug_szf,N_zlnyb,N_fwyb,N_fwnyb,C_s_flag,C_zt) values("
    
    '20051128修改，陈宏悦；由于医保中心对作废单据流水号规定为 医院编码（8位）＋“－”＋流水号（原结帐的流水号）
    'strSql = strSql & "'" & Rpad(Substr(InitInfor_兴成.医院编码, 1, 8), 8, " ") & Lpad(Substr(lng冲销ID, 1, 12), 12, "0") & "',"     'vc_jyh(门诊收费流水号)
    
    strSQL = strSQL & "'" & Rpad(Substr(InitInfor_兴成.医院编码, 1, 8), 8, " ") & "-" & Lpad(Substr(str交易流水号, 10, 11), 11, "0") & "',"     'vc_jyh(门诊收费流水号)
    strSQL = strSQL & "'" & Format(zlDatabase.Currentdate, "yyyymmddHHMMSS") & "'," 'vc_date(费用发生日期)
    strSQL = strSQL & "'" & g病人身份_兴成.社会保障号 & "',"  'C_miid(医保编号)
    strSQL = strSQL & "'" & g病人身份_兴成.IC卡号 & "',"   'vc_cardid(卡号)
    strSQL = strSQL & "0,"   'Vc_jzh(挂号流水号)：没有实际意义，缺省为零
    strSQL = strSQL & "'" & InitInfor_兴成.医院编码 & "',"    ' C_yljgdm(医疗机构代码)
    strSQL = strSQL & "" & IIf(str开单部门 = "", "NULL", "'" & str开单部门 & "'") & ","    ' C_ksid(科室名称)
    strSQL = strSQL & "" & IIf(str开单医师 = "", "NULL", "'" & str开单医师 & "'") & ","    ' C_doctorid(开单医师)
    
    '陈宏悦于20051231修改，由于医保前置库中的字段长度不一致
    If g病人身份_兴成.慢病标识 = "1" Then
        strSQL = strSQL & "" & IIf(g病人身份_兴成.慢性病病种 = "", "Null", "'" & g病人身份_兴成.慢性病病种 & "'") & ","    'Vc_bzid(病种)
    Else
        strSQL = strSQL & "" & "Null" & ","    'Vc_bzid(病种)
    End If
    
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!费用总额, 0), "####0.00;-####0.00;0;0") & "," 'N_sum(总金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!非医保范围金额, 0), "####0.00;-####0.00;0;0") & "," 'N_nyb(非医保范围金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!医保范围金额, 0), "####0.00;-####0.00;0;0") & "," 'N_yb(医保范围金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsTemp!个人帐户支付, 0), "####0.00;-####0.00;0;0") & "," 'N_grzf(个人帐户支付金额),
    strSQL = strSQL & "" & Format(-1 * Nvl(rsTemp!全自付金额, 0), "####0.00;-####0.00;0;0") & "," 'N_xzjf(现金支付金额（帐户不足支付）)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!药品甲类金额, 0), "####0.00;-####0.00;0;0") & "," 'N_drug_a(药品甲类金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!乙类药品医保金额, 0), "####0.00;-####0.00;0;0") & "," 'N_drug_byb(乙类药品医保金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!乙类药品自负金额, 0), "####0.00;-####0.00;0;0") & "," 'N_drug_bzf(乙类药品自负金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!自费药品金额, 0), "####0.00;-####0.00;0;0") & "," ',N_drug_zf(自费药品金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!诊疗医保金额, 0), "####0.00;-####0.00;0;0") & "," ',N_drug_mi(诊疗医保金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!特检特治支付额, 0), "####0.00;-####0.00;0;0") & "," ',N_drug_s(特检特治医保金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!特检特治自负额, 0), "####0.00;-####0.00;0;0") & "," ',N_drug_szf(特检特治自负金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!诊疗非医保金额, 0), "####0.00;-####0.00;0;0") & "," 'N_zlnyb(诊疗非医保金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!服务医保金额, 0), "####0.00;-####0.00;0;0") & "," 'N_fwyb(服务医保金额)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!服务非医保金额, 0), "####0.00;-####0.00;0;0") & "," 'N_fwnyb(服务非医保金额)
    strSQL = strSQL & "'" & Nvl(rsData!特殊病标识, 0) & "',"  ''C_s_flag(特殊病标志)
    strSQL = strSQL & "0)"   'C_zt 0 新数据 1 已正确传输，状态未知 2 已正确传输，并成功入库
    
    gcnSQLSEVER_兴成.Execute strSQL
                    
  
    gstrSQL = "select 结帐id, 性质, 交易流水号, 处方号, 项目代码, 项目名称, 项目类别, 单价, 数量,  " & _
             "        费用总金额, 医保内费用, 医保外费用, 药品甲类金额, 乙类医保金额, 乙类自负金额,  " & _
             "        自费药品金额, 诊疗医保金额, 特检特治金额, 特检特治自负, 诊疗非医保额, 服务医保金额,  " & _
             "        服务非医保额, 费用发生日期, mac2  " & _
             " from 医保结算明细记录" & _
             " where 性质=1 and 结帐id=" & lng冲销ID   '陈宏悦于20050315添加修改;因为冲销本张单据
             
             
    
    Call OpenRecordset_兴成(rsData, "获取结算明细记录", gstrSQL)
    
    lngRow = 1
    Do While Not rsData.EOF
        
        '明细:交易流水号|处方号|项目代码|项目名称|项目类别|单价|数量|费用总金额|医保内费用|医保外费用|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|MAC2
        '表结构:VC_RECEIPTID(门诊收费流水号),N_SFSXH(收费顺序号),VC_ITEM_ID(项目代码),VC_ITEM_NAME(项目名称),C_MIID(保险编号),VC_JZH(挂号流水号),C_SSMLJB(所属目录级别),C_ZLXMJB(诊疗项目级别),C_FWSSFW(服务设施范围),N_PRICE(单价),N_sl(数量),N_sum(费用总金额),N_DRUG_A(药品甲类金额),N_DRUG_BYB(乙类药品医保金额),N_drug_bzf(乙类药品自负金额),N_drug_zf(自费药品金额),D_DRUG_ZLYB(诊疗医保金额),N_SYB(特检特治医保金额),N_SYBZF(特检特治自负金额),N_ZLFYB(诊疗非医保金额),N_fwyb(服务医保金额),N_fwnyb(服务非医保金额),C_CF_FLAG(处方标识),C_zt(传输状态)
        
        strSQL = "insert into HLD_MZCFZLXX(VC_RECEIPTID,N_SFSXH,VC_ITEM_ID,VC_ITEM_NAME,C_MIID,VC_JZH,C_SSMLJB,C_ZLXMJB,C_FWSSFW,N_PRICE,N_sl,N_sum,N_DRUG_A,N_DRUG_BYB,N_drug_bzf,N_drug_zf,N_DRUG_ZLYB,N_SYB,N_SYBZF,N_ZLFYB,N_fwyb,N_fwnyb,C_CF_FLAG,C_zt) values("
       
       '20051128修改，陈宏悦；由于医保中心对作废单据流水号规定为 医院编码（8位）＋“－”＋流水号（原结帐的流水号）
        'strSql = strSql & "'" & Rpad(Substr(InitInfor_兴成.医院编码, 1, 8), 8, " ") & Lpad(Substr(lng冲销ID, 1, 12), 12, "0") & "',"        'VC_RECEIPTID (门诊收费流水号)
        
        strSQL = strSQL & "'" & Rpad(Substr(InitInfor_兴成.医院编码, 1, 8), 8, " ") & "-" & Lpad(Substr(str交易流水号, 10, 11), 11, "0") & "',"        'VC_RECEIPTID (门诊收费流水号)
        strSQL = strSQL & "" & lngRow & ","        'N_SFSXH (收费顺序号)
        strSQL = strSQL & "'" & Nvl(rsData!项目代码) & "',"        'VC_ITEM_ID (项目代码)
        strSQL = strSQL & "'" & Nvl(rsData!项目名称) & "',"        'VC_ITEM_NAME (项目名称)
        strSQL = strSQL & "'" & g病人身份_兴成.保险编号 & "',"            'C_MIID (保险编号)
        
        strSQL = strSQL & "0,"        'VC_JZH (挂号流水号)      '没有实际意义
        Select Case Nvl(rsData!项目类别, "0")
            Case "0"    '药品
                gstrSQL = "Select ssmljb as 级别 From YB_YD where xmdm='" & Nvl(rsData!项目代码, "0") & "'"
            Case "1"    '诊疗
                gstrSQL = "Select tjtzbz as 级别 From YB_ZLML where xmdm='" & Nvl(rsData!项目代码, "0") & "'"
            Case Else   '服务
                gstrSQL = "Select fwfw as 级别 From YB_FWSS where xmdm='" & Nvl(rsData!项目代码, "0") & "'"
        End Select
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnSQLSEVER_兴成
        
'        陈宏悦于20050321修改注释掉
        
'        If rsTemp.EOF Then
'            ShowMsgbox "在保存结算记录时,未发现相关的医保项目代码[" & Nvl(rsData!项目代码) & "]"
'            Exit Function
'        End If
        
       Select Case Nvl(rsData!项目类别, "0")

             Case "0"    '药品
                    '陈宏悦于20050321修改,因为存在不需对码的项目

                     strSQL = strSQL & "'0',"           'C_SSMLJB (所属目录级别)
                     If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (诊疗项目级别)
                     Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (诊疗项目级别)
                     End If

                 Case "1"    '诊疗
                    strSQL = strSQL & "'1',"              'C_SSMLJB (所属目录级别)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"           'C_ZLXMJB (诊疗项目级别)
                    Else
                       strSQL = strSQL & "'0',"        'C_SSMLJB (所属目录级别)
                    End If

                 Case Else   '服务
                    strSQL = strSQL & "'2',"           'C_SSMLJB (所属目录级别)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (诊疗项目级别)
                    Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (诊疗项目级别)
                    End If
                    
      End Select
      
      strSQL = strSQL & "NULL,"        'C_FWSSFW (服务设施范围)
      
'            Case "0"    '药品
'
'                '陈宏悦于20050321修改,因为存在不需对码的项目
'                 If rsTemp.EOF Then
'                    strSQL = strSQL & "3,"        'C_SSMLJB (所属目录级别)
'                    strSQL = strSQL & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                    strSQL = strSQL & "NULL,"        'C_FWSSFW (服务设施范围)
'                 Else
'                    strSQL = strSQL & "'" & Nvl(rsTemp!级别) & "',"        'C_SSMLJB (所属目录级别)
'                    strSQL = strSQL & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                    strSQL = strSQL & "NULL,"        'C_FWSSFW (服务设施范围)
'                 End If
'
'            Case "1"    '诊疗
'                  If rsTemp.EOF Then
'                    strSQL = strSQL & "NULL,"        'C_SSMLJB (所属目录级别)
'                    strSQL = strSQL & "2,"        'C_ZLXMJB (诊疗项目级别)
'                    strSQL = strSQL & "NULL,"        'C_FWSSFW (服务设施范围)
'                 Else
'                    strSQL = strSQL & "NULL,"        'C_SSMLJB (所属目录级别)
'                    strSQL = strSQL & "'" & Nvl(rsTemp!级别) & "',"        'C_ZLXMJB (诊疗项目级别)
'                    strSQL = strSQL & "NULL,"        'C_FWSSFW (服务设施范围)
'                 End If
'
'            Case Else   '服务
'                 If rsTemp.EOF Then
'                    strSQL = strSQL & "NULL,"        'C_SSMLJB (所属目录级别)
'                    strSQL = strSQL & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                    strSQL = strSQL & "1,"        'C_FWSSFW (服务设施范围)
'                 Else
'                    strSQL = strSQL & "NULL,"        'C_SSMLJB (所属目录级别)
'                    strSQL = strSQL & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                    strSQL = strSQL & "'" & Nvl(rsTemp!级别) & "',"        'C_FWSSFW (服务设施范围)
'                 End If
         
        strSQL = strSQL & "" & Format(Nvl(rsData!单价, 0), "####0.0000;-####0.0000;0;0") & ","        'N_PRICE (单价)
        strSQL = strSQL & "" & Format(Nvl(rsData!数量, 0), "####0;-####0;0;0") & ","          'N_sl (数量)
        strSQL = strSQL & "" & Format(Nvl(rsData!费用总金额, 0), "####0.00;-####0.00;0;0") & ","        'N_sum (费用总金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!药品甲类金额, 0), "####0.00;-####0.00;0;0") & ","        'N_DRUG_A (药品甲类金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!乙类医保金额, 0), "####0.00;-####0.00;0;0") & ","      'N_DRUG_BYB (乙类药品医保金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!乙类自负金额, 0), "####0.00;-####0.00;0;0") & ","        'N_drug_bzf (乙类药品自负金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!自费药品金额, 0), "####0.00;-####0.00;0;0") & ","       'N_drug_zf (自费药品金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!诊疗医保金额, 0), "####0.00;-####0.00;0;0") & ","       'D_DRUG_ZLYB (诊疗医保金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!特检特治金额, 0), "####0.00;-####0.00;0;0") & ","        'N_SYB (特检特治医保金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!特检特治自负, 0), "####0.00;-####0.00;0;0") & ","        'N_SYBZF (特检特治自负金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!诊疗非医保额, 0), "####0.00;-####0.00;0;0") & ","       'N_ZLFYB (诊疗非医保金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!服务医保金额, 0), "####0.00;-####0.00;0;0") & ","       'N_fwyb (服务医保金额)
        strSQL = strSQL & "" & Format(Nvl(rsData!服务非医保额, 0), "####0.00;-####0.00;0;0") & ","       'N_fwnyb (服务非医保金额)
        strSQL = strSQL & "NULL,"        'C_CF_FLAG (处方标识)  :无实际意义
        strSQL = strSQL & "'0')"        'C_zt (传输状态)
        gcnSQLSEVER_兴成.Execute strSQL
        rsData.MoveNext
        
        '陈宏悦于20050402修改添加
        
        lngRow = lngRow + 1
        
    Loop
    InsertIntoSQLServer_门诊冲销 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function InsertIntoData_住院(ByVal str确认串 As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '插入数据到SQLSErver中.
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对出参文件的内容进行分解
    '--入参数:str确认串-门诊确认支付时传出的串
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strFile As String, StrInput As String, strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim lngRow As Long
    Dim strArr As Variant
    Dim strArr1 As Variant
    Dim strSQL As String
    Dim str开单部门 As String
    Dim str开单医师 As String
    Dim str结算时间 As String, str入院时间 As String, str出院日期 As String, str治愈情况 As String
    Dim lng住院天数  As Long
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "Select to_char(收费时间,'yyyyMMDD') as 结算时间 from 病人结帐记录 where ID=" & g病人身份_兴成.结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取结算信息"
    
    str结算时间 = Nvl(rsTemp!结算时间)
    gstrSQL = "" & _
        "   Select to_char(a.入院日期,'yyyyMMDD') as 入院日期,to_char(a.出院日期,'yyyyMMDD') as 出院日期," & _
        "           decode(trunc(a.出院日期)-trunc(a.入院日期),0,1,trunc(a.出院日期)-trunc(a.入院日期)) as 天数,a.出院方式,a.住院医师,b.名称 as 住院科室" & _
        "   from 病案主页 a,部门表 b " & _
        "   where a.入院科室ID=b.id(+) and a.病人id=" & lng病人ID & " and a.主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsData, gstrSQL, "获取结算信息"
    
    str入院时间 = Nvl(rsData!入院日期): str出院日期 = Nvl(rsData!出院日期): lng住院天数 = Nvl(rsData!天数, 0)
    '1治愈2 好转3 未愈4 转出5 死亡

    Select Case Nvl(rsData!出院方式)
      Case "治愈"
          str治愈情况 = 1
      Case "好转"
          str治愈情况 = 2
      Case "未愈"
          str治愈情况 = 3
      Case "死亡"
          str治愈情况 = 5
      Case "转院"
          str治愈情况 = 4
      Case Else
          str治愈情况 = 1
      End Select
        
    gstrSQL = "Select b.名称 as 开单部门,开单人 from 住院费用记录 a,部门表 b where a.开单部门id=b.id(+) and 结帐id=" & g病人身份_兴成.结帐ID & " and rownum=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取开单人"
    If rsTemp.EOF Then
        str开单部门 = ""
        str开单医师 = ""
    Else
        str开单部门 = Substr(Nvl(rsTemp!开单部门), 1, 12)
        str开单医师 = Substr(Nvl(rsTemp!开单人), 1, 4)    '由于导出的数据库的结果为4位,但文档中是12位,请到时更改/
    End If
    
    InsertIntoData_住院 = False
    
    If objFile.FolderExists(InitInfor_兴成.strPath_Out) = False Then
        ShowMsgbox "未创建该文件夹(" & InitInfor_兴成.strPath_Out & "),请创建!"
        Exit Function
    End If
    strArr1 = Split(str确认串, "|")
    strFile = InitInfor_兴成.strPath_Out & "\Hosp_Divide.out"
    
    Err = 0: On Error GoTo errHand:
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "没有产生相关的出参文件：" & strFile & vbCrLf & " 请检查!"
        Exit Function
    End If
    Set objText = objFile.OpenTextFile(strFile, ForReading)
    
    lngRow = 1
    Do While Not objText.AtEndOfStream
          strText = Trim(objText.ReadLine)
          If strText = "" Then Exit Do
          strArr = Split(strText, "|")
          
          If lngRow = 1 Then
                
                '读汇总信息
                '住院:   IC卡号|入院类别|医院级别|病种|文件记录条数|本次交易费用总金额|统筹支付金额|个人应付总金额|个人帐户余额|非医保范围金额|医保范围金额|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|起付额|共付段金额|封顶线以上自负金额|本年度统筹基金支付累计|本年度共付段金额累计|本年住院次数
                '住院转入:交易流水号|IC卡号|终端机编号|交易日期/时间|费用总金额|医保内总费用|医保外总费用|统筹支付金额|统筹自负金额||个人帐户支付金额|现金支付金额|扣减后个人帐户余额|MAC1
                '表结构:Field1(住院结算流水号),Field2(住院号),Field3(保险编号),Field4(医疗机构代码),Field5(住院类别),Field6(结算日期),Field7(入院日期),Field8(出院日期),Field13(人员类别),Field14(住院天数),Field15(本年住院天数),Field16(本年统筹基金支出累计),Field17(共付段个人支出累计),Field18(入院类别),Field19(出院类别),Field20(治愈情况),Field21(病种),Field22(总金额),Field23(药品甲类金额),Field24(乙类药品医保金额),Field25(乙类药品自负金额),Field26(自费药品金额),Field27(诊疗医保金额),Field28(特检特治医保金额),Field30(诊疗非医保金额),Field31(服务医保金额),Field32(服务非医保金额),Field33(起付额),Field34(非医保范围金额),Field35(封顶线以上自负金额),Field36(个人帐户支付金额),Field37(住院科室),Field38(主管医师),Field39(医院操作员),Field29(特检特治自负金额),Field40(统筹支付金额),Field41(共付段金额),C_kzt38(卡状态),C_zt39(传输状态),ydyymc(异地医院名称),ydyyjb(异地医院级别)
'
'                strSql = "insert into  HLD_HOSP_JIESUAN(Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field13,Field14,Field15,Field16,Field17,Field18,Field19,Field20,Field21,Field22,Field23,Field24,Field25,Field26,Field27,Field28,Field30,Field31,Field32,Field33,Field34,Field35,Field36,Field37,Field38,Field39,Field29,Field40,Field41,C_kzt,C_zt,ydyymc,ydyyjb) values("

                strSQL = "insert into  HLD_HOSP_JIESUAN(Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field13,Field14,Field15,Field16,Field17,Field18,Field19,Field20,Field21,Field22,Field23,Field24,Field25,Field26,Field27,Field28,Field30,Field31,Field32,Field33,Field34,Field35,Field36,Field37,Field38,Field39,Field29,Field40,Field41,C_kzt,C_zt) values("
                strSQL = strSQL & "'" & strArr1(0) & "',"   'Field1(住院结算流水号),
                strSQL = strSQL & "'" & lng病人ID & "_" & lng主页ID & "',"   'Field2(住院号),
                strSQL = strSQL & "'" & g病人身份_兴成.保险编号 & "',"   'Field3(保险编号),
                strSQL = strSQL & "'" & InitInfor_兴成.医院编码 & "',"   'Field4(医疗机构代码),
                strSQL = strSQL & "'" & g病人身份_兴成.住院类别 & "',"  'Field5(住院类别),
                strSQL = strSQL & "'" & str结算时间 & "'," 'Field6(结算日期),
                strSQL = strSQL & "'" & str入院时间 & "'," 'Field7(入院日期),
                strSQL = strSQL & "'" & str出院日期 & "'," 'Field8(出院日期),
                strSQL = strSQL & "'" & g病人身份_兴成.人员类别 & "',"  'Field13(人员类别),
                strSQL = strSQL & "" & lng住院天数 & "," 'Field14(住院天数),
                strSQL = strSQL & "'" & g病人身份_兴成.本年住院次数 + 1 & "'," 'Field15(本年住院次数),
                strSQL = strSQL & "" & Format(Val(strArr(24)) / 100, "####0.00;-####0.00;0;0") & ","  'Field16(本年统筹基金支出累计),
                strSQL = strSQL & "" & Format(Val(strArr(25)) / 100, "####0.00;-####0.00;0;0") & ","  'Field17(共付段个人支出累计),
                strSQL = strSQL & "'" & g病人身份_兴成.入院类别 & "',"  'Field18(入院类别),
                strSQL = strSQL & "'" & g病人身份_兴成.出院类别 & "',"  'Field19(出院类别),
                strSQL = strSQL & "'" & str治愈情况 & "',"   'Field20(治愈情况),
                strSQL = strSQL & "'" & g病人身份_兴成.病种编码 & "',"  'Field21(病种),
                strSQL = strSQL & "" & Format(Val(strArr(5)) / 100, "####0.00;-####0.00;0;0") & ","  'Field22(总金额),
                strSQL = strSQL & "" & Format(Val(strArr(11)) / 100, "####0.00;-####0.00;0;0") & "," 'Field23(药品甲类金额),
                strSQL = strSQL & "" & Format(Val(strArr(12)) / 100, "####0.00;-####0.00;0;0") & "," 'Field24(乙类药品医保金额),
                strSQL = strSQL & "" & Format(Val(strArr(13)) / 100, "####0.00;-####0.00;0;0") & "," 'Field25(乙类药品自负金额),
                strSQL = strSQL & "" & Format(Val(strArr(14)) / 100, "####0.00;-####0.00;0;0") & "," 'Field26(自费药品金额),
                strSQL = strSQL & "" & Format(Val(strArr(15)) / 100, "####0.00;-####0.00;0;0") & "," 'Field27(诊疗医保金额),
                strSQL = strSQL & "" & Format(Val(strArr(16)) / 100, "####0.00;-####0.00;0;0") & "," 'Field28(特检特治医保金额),
                strSQL = strSQL & "" & Format(Val(strArr(18)) / 100, "####0.00;-####0.00;0;0") & "," 'Field30(诊疗非医保金额),
                strSQL = strSQL & "" & Format(Val(strArr(19)) / 100, "####0.00;-####0.00;0;0") & "," 'Field31(服务医保金额),
                strSQL = strSQL & "" & Format(Val(strArr(20)) / 100, "####0.00;-####0.00;0;0") & "," 'Field32(服务非医保金额),
                strSQL = strSQL & "" & Format(Val(strArr(21)) / 100, "####0.00;-####0.00;0;0") & "," 'Field33(起付额),
                strSQL = strSQL & "" & Format(Val(strArr(9)) / 100, "####0.00;-####0.00;0;0") & "," 'Field34(非医保范围金额),
                strSQL = strSQL & "" & Format(Val(strArr(23)) / 100, "####0.00;-####0.00;0;0") & "," 'Field35(封顶线以上自负金额),
                strSQL = strSQL & "" & Format(Val(strArr1(9)) / 100, "####0.00;-####0.00;0;0") & "," 'Field36(个人帐户支付金额),
                strSQL = strSQL & "'" & Substr(Nvl(rsData!住院科室), 1, 20) & "'," 'Field37(住院科室),
                strSQL = strSQL & "'" & Substr(Nvl(rsData!住院医师), 1, 16) & "',"  'Field38(主管医师),
                strSQL = strSQL & "'" & Substr(gstrUserName, 1, 4) & "'," 'Field39(医院操作员),
                strSQL = strSQL & "" & Format(Val(strArr(17)) / 100, "####0.00;-####0.00;0;0") & "," 'Field29(特检特治自负金额),
                strSQL = strSQL & "" & Format(Val(strArr1(7)) / 100, "####0.00;-####0.00;0;0") & "," 'Field40(统筹支付金额),
                strSQL = strSQL & "" & Format(Val(strArr(22)) / 100, "####0.00;-####0.00;0;0") & "," 'Field41(共付段金额),
                strSQL = strSQL & "'" & g病人身份_兴成.卡状态 & "',"     'C_kzt38(卡状态),
'                strSql = strSql & "'0'" & ","  'C_zt39(传输状态)
                strSQL = strSQL & "'0'" & ")"  'C_zt39(传输状态)
'                strSql = strSql & "'" & g病人身份_兴成.异地医院 & "',"      'C_kzt38(卡状态),
'                strSql = strSql & "'" & g病人身份_兴成.异地医院级别 & "')"
                gcnSQLSEVER_兴成.Execute strSQL
                '插入中间库:
                If InsertIntoYBK(strArr, lngRow, False, False) = False Then Exit Function
        Else
        
             '住院:交易流水号|项目序号|项目代码|项目名称|项目类别|单价|数量|费用总金额|医保内费用|医保外费用|费用发生日期|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|MAC2
             '表结构:c_receipt_no(住院结算流水号),sort_it(收费顺序号),c_insu_id(保险编号),c_item_code(项目代码),c_item_name(项目名称),c_dir(所属目录级别),c_zlxmjb(诊疗项目级别),c_ser(服务设施范围),n_price(单价),n_amount(数量),n_account(金额),n_med_jia(药品甲类金额),n_med_yi(乙类药品医保金额),n_med_yi_self(乙类药品自负金额),n_med_self(自费药品金额),n_zlyb(诊疗医保金额),n_tjtzyb(特检特治医保金额),n_tjtzyb_self(特检特治自负金额),n_zlfyb(诊疗非医保金额),n_fwyb(服务医保金额),n_fwfyb(服务非医保金额),dc_fyfsrq(费用日期),c_wzzxm(外转诊项目),C_zt(传输状态)
             
             
             strSQL = "insert into HLD_HOSP(c_receipt_no,sort_ID,c_insu_id,c_item_code,c_item_name,c_dir,c_zlxmjb,c_ser,n_price,n_amount,n_account,n_med_jia,n_med_yi,n_med_yi_self,n_med_self,n_zlyb,n_tjtzyb,n_tjtzyb_self,n_zlfyb,n_fwyb,n_fwfyb,dc_fyfsrq,c_wzzxm,C_zt) values("
             
             strSQL = strSQL & "'" & strArr1(0) & "',"  '    c_receipt_no(住院结算流水号),
             strSQL = strSQL & "" & lngRow - 1 & "," '    sort_it(收费顺序号),
             strSQL = strSQL & "'" & g病人身份_兴成.保险编号 & "',"  '    c_insu_id(保险编号),
             strSQL = strSQL & "'" & strArr(2) & "',"  '    c_item_code(项目代码),
             strSQL = strSQL & "'" & strArr(3) & "',"  '    c_item_name(项目名称),
             
             '陈宏悦于200500403修改添加，因为此部分接口文档描述不太清楚；正确描述如下：
             '1：所属目录级别应为 0-药品，1-诊疗，2-服务
             '2：诊疗项目级别应为 0-医保，1-非医保
             '3：服务设施范围没有什么意义
             
             Select Case strArr(4)
                 Case "0"    '药品
                     gstrSQL = "Select ssmljb as 级别 From YB_YD where xmdm='" & strArr(2) & "'"
                 Case "1"    '诊疗
                     gstrSQL = "Select tjtzbz as 级别 From YB_ZLML where xmdm='" & strArr(2) & "'"
                 Case Else   '服务
                     gstrSQL = "Select fwfw as 级别 From YB_FWSS where xmdm='" & strArr(2) & "'"
             End Select
             If rsTemp.State = 1 Then rsTemp.Close
             rsTemp.Open gstrSQL, gcnSQLSEVER_兴成

'             '陈宏悦于20050321修改
'
'             If rsTemp.EOF Then
'                 ShowMsgbox "在保存结算记录时,未发现相关的医保项目代码[" & strArr(2) & "]"
'                 Exit Function
'             End If

              Select Case strArr(4)
                 Case "0"    '药品
                    '陈宏悦于20050321修改,因为存在不需对码的项目

                     strSQL = strSQL & "'0',"           'C_SSMLJB (所属目录级别)
                     If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (诊疗项目级别)
                     Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (诊疗项目级别)
                     End If

                 Case "1"    '诊疗
                    strSQL = strSQL & "'1',"              'C_SSMLJB (所属目录级别)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"           'C_ZLXMJB (诊疗项目级别)
                    Else
                       strSQL = strSQL & "'0',"        'C_SSMLJB (所属目录级别)
                    End If

                 Case Else   '服务
                    strSQL = strSQL & "'2',"           'C_SSMLJB (所属目录级别)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (诊疗项目级别)
                    Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (诊疗项目级别)
                    End If
                    
             End Select
             
             strSQL = strSQL & "NULL,"        'C_FWSSFW (服务设施范围)

'陈宏悦于20050403修改
'            Select Case strArr(4)
'                 Case "0"    '药品
'                    '陈宏悦于20050321修改,因为存在不需对码的项目
'                     If rsTemp.EOF Then
'                        strSql = strSql & "3,"           'C_SSMLJB (所属目录级别)
'                        strSql = strSql & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                        strSql = strSql & "NULL,"        'C_FWSSFW (服务设施范围)
'                     Else
'                        strSql = strSql & "'" & Nvl(rsTemp!级别) & "',"        'C_SSMLJB (所属目录级别)
'                        strSql = strSql & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                        strSql = strSql & "NULL,"        'C_FWSSFW (服务设施范围)
'                     End If
'
'                 Case "1"    '诊疗
'                    If rsTemp.EOF Then
'                       strSql = strSql & "NULL,"        'C_SSMLJB (所属目录级别)
'                       strSql = strSql & "2,"           'C_ZLXMJB (诊疗项目级别)
'                       strSql = strSql & "NULL,"        'C_FWSSFW (服务设施范围)
'                    Else
'                       strSql = strSql & "NULL,"        'C_SSMLJB (所属目录级别)
'                       strSql = strSql & "'" & Nvl(rsTemp!级别) & "',"        'C_ZLXMJB (诊疗项目级别)
'                       strSql = strSql & "NULL,"        'C_FWSSFW (服务设施范围)
'                    End If
'
'                 Case Else   '服务
'                    If rsTemp.EOF Then
'                       strSql = strSql & "NULL,"        'C_SSMLJB (所属目录级别)
'                       strSql = strSql & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                       strSql = strSql & "1,"           'C_FWSSFW (服务设施范围)
'                    Else
'                       strSql = strSql & "NULL,"        'C_SSMLJB (所属目录级别)
'                       strSql = strSql & "NULL,"        'C_ZLXMJB (诊疗项目级别)
'                       strSql = strSql & "'" & Nvl(rsTemp!级别) & "',"        'C_FWSSFW (服务设施范围)
'                    End If
'             End Select
                     
             strSQL = strSQL & "" & Format(Val(strArr(5)) / 10000, "####0.00;-####0.00;0;0") & ","         '    n_price(单价),
             strSQL = strSQL & "" & Format(Val(strArr(6)), "####0;-####0;0;0") & ","         '    n_amount(数量),
             strSQL = strSQL & "" & Format(Val(strArr(7)) / 100, "####0.00;-####0.00;0;0") & "," '    n_account(金额),
             strSQL = strSQL & "" & Format(Val(strArr(11)) / 100, "####0.00;-####0.00;0;0") & "," '    n_med_jia(药品甲类金额),
             strSQL = strSQL & "" & Format(Val(strArr(12)) / 100, "####0.00;-####0.00;0;0") & "," '    n_med_yi(乙类药品医保金额),
             strSQL = strSQL & "" & Format(Val(strArr(13)) / 100, "####0.00;-####0.00;0;0") & "," '    n_med_yi_self(乙类药品自负金额),
             strSQL = strSQL & "" & Format(Val(strArr(14)) / 100, "####0.00;-####0.00;0;0") & "," '    n_med_self(自费药品金额),
             strSQL = strSQL & "" & Format(Val(strArr(15)) / 100, "####0.00;-####0.00;0;0") & "," '    n_zlyb(诊疗医保金额),
             strSQL = strSQL & "" & Format(Val(strArr(16)) / 100, "####0.00;-####0.00;0;0") & ","  '    n_tjtzyb(特检特治医保金额),
             strSQL = strSQL & "" & Format(Val(strArr(17)) / 100, "####0.00;-####0.00;0;0") & "," '    n_tjtzyb_self(特检特治自负金额),
             strSQL = strSQL & "" & Format(Val(strArr(18)) / 100, "####0.00;-####0.00;0;0") & "," '    n_zlfyb(诊疗非医保金额),
             strSQL = strSQL & "" & Format(Val(strArr(19)) / 100, "####0.00;-####0.00;0;0") & "," '    n_fwyb(服务医保金额),
             strSQL = strSQL & "" & Format(Val(strArr(20)) / 100, "####0.00;-####0.00;0;0") & "," '    n_fwfyb(服务非医保金额),
             
             strSQL = strSQL & "'" & strArr(10) & "'," '    dc_fyfsrq(费用日期),
             strSQL = strSQL & "'0'," '    c_wzzxm(外转诊项目),
             strSQL = strSQL & "'0')" '    C_zt (传输状态)
             gcnSQLSEVER_兴成.Execute strSQL
            '插入中间库:
            If InsertIntoYBK(strArr, lngRow, True, False) = False Then Exit Function
        End If
        lngRow = lngRow + 1
    Loop
    InsertIntoData_住院 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function




Private Function InsertIntoData_住院登记(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '插入数据到SQLSErver中.
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:将冲销的相关数据插入SQLServer中
    '--入参数:lng冲销ID值
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim str开单部门 As String
    Dim str开单医师 As String
    
    
    
    
    
    InsertIntoData_住院登记 = False
        
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "" & _
        "   Select a.入院科室ID,b.名称 as 入院科室,to_char(a.入院日期,'yyyy-mm-dd hh24:mi:ss') as 入院日期" & _
        "   From 病案主页 a,部门表 b" & _
        "   where a.入院科室id=b.id(+) and a.病人id=" & lng病人ID & " and a.主页id =" & lng主页ID
        
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取入院信息"
    
    '表结构:xh（序号）,yljgdm（医疗机构代码）,bxbh（保险编号）,zyh（住院号）,ryrq（入院日期）,rylb（入院类别）,bz（病种）,zyks（住院科室）,C_zt（传输状态）
    '
    '注意:接口中有序号,而在库中无序号
    gstrSQL = "insert into HLD_ZYBRXX(yljgdm,bxbm,zyh,ryrq,rylb,bz,zyks,C_zt) values ("
    
    '暂无
    'gstrSQL = gstrSQL & "," 'xh（序号）
    
    gstrSQL = gstrSQL & "'" & InitInfor_兴成.医院编码 & "',"   'yljgdm（医疗机构代码）
    gstrSQL = gstrSQL & "'" & g病人身份_兴成.保险编号 & "',"  'bxbh（保险编号）,
    gstrSQL = gstrSQL & "'" & lng病人ID & "_" & lng主页ID & "',"   'zyh（住院号）,
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!入院日期) & "'," 'ryrq（入院日期）,
    gstrSQL = gstrSQL & "'" & g病人身份_兴成.入院类别 & "',"  'rylb（入院类别）,
    gstrSQL = gstrSQL & "'" & g病人身份_兴成.病种编码 & "',"  'bz（病种）,
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!入院科室) & "',"  'zyks（住院科室）,
    gstrSQL = gstrSQL & "'0')" 'C_zt（传输状态）
    
    gcnSQLSEVER_兴成.Execute gstrSQL

    InsertIntoData_住院登记 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InsertIntoYBK(ByVal strVarr As Variant, ByVal lng序号 As Long, Optional bln明细 As Boolean = True, Optional bln门诊 As Boolean = True) As Boolean
    '根据传入参数，将数据保存到中间库去
    Dim strSQL As String
    
    Dim i As Long
    Dim str发生日期 As String
    
    Err = 0: On Error GoTo errHand:
    
    InsertIntoYBK = False
    If bln明细 = False Then
        i = 0
        '过程参数
        '    性质_IN,结帐ID_IN,入院类别_IN IN,医院级别_IN,病种_IN,
        '    记录条数_IN,费用总额_IN,统筹支付金额_IN,个人应付总额_IN,
        '    个人帐户余额_IN,非医保范围金额_IN ,医保范围金额_IN,药品甲类金额_IN,乙类药品医保金额_IN,乙类药品自负金额_IN,
        '    自费药品金额_IN,诊疗医保金额_IN,特检特治支付额_IN,特检特治自负额_IN,诊疗非医保金额_IN,
        '    服务医保金额_IN,服务非医保金额_IN,特殊病标识_IN,起付额_IN,共付段金额_IN,封顶线自付金额_IN,
        '    年度统筹基金累计_IN,年度共付段累计_IN,本年住院次数_IN
        strSQL = "ZL_医保结算记录_INSERT("
        '门诊:                                 文件记录条数|本次交易费用总金额             |个人应付总金额|个人帐户余额|非医保范围金额|医保范围金额|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|特殊病标识
        '住院:   IC卡号|入院类别|医院级别|病种|文件记录条数|本次交易费用总金额|统筹支付金额|个人应付总金额|个人帐户余额|非医保范围金额|医保范围金额|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|起付额|共付段金额|封顶线以上自负金额|本年度统筹基金支付累计|本年度共付段金额累计|本年住院次数
        
        If bln门诊 Then
            strSQL = strSQL & 1 & ","
            strSQL = strSQL & g病人身份_兴成.结帐ID & ","
        Else
            strSQL = strSQL & IIf(g病人身份_兴成.结帐ID = 0, 3, 2) & ","
            strSQL = strSQL & IIf(g病人身份_兴成.结帐ID = 0, g病人身份_兴成.病人ID, g病人身份_兴成.结帐ID) & ","
        End If
        
        If Not bln门诊 Then
            i = i + 1
            '入院类别_IN IN,医院级别_IN,病种_IN
            strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
            strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
            strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
        Else
            strSQL = strSQL & "NULL" & ","
            strSQL = strSQL & "NULL" & ","
            strSQL = strSQL & "NULL" & ","
        End If
        
        strSQL = strSQL & Val(strVarr(i)) & ",": i = i + 1
        
        '陈宏悦于20050315修改,根据数据库内容进行更新;将原不/100改为/100
        
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        
        If bln门诊 Then
            '统筹支付金额
            strSQL = strSQL & 0 & ","
        Else
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        End If
        
        '陈宏悦于20050315修改，由于存入到数据库的数据与发生的费用不符
        
       ' MsgBox strSql, vbOKOnly, "zlsoft"
       
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        If bln门诊 Then
            '特殊病标识
            strSQL = strSQL & "'" & strVarr(16) & "',"
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ")"
        Else
            '起付额|共付段金额|封顶线以上自负金额|本年度统筹基金支付累计|本年度共付段金额累计|本年住院次数
            '特殊病标识_IN,起付额_IN,共付段金额_IN,封顶线自付金额_IN,年度统筹基金累计_IN,年度共付段累计_IN,本年住院次数_IN
            strSQL = strSQL & "NULL,"
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        
            '陈宏悦修改于20060109修改,保存住院次数不正确
            'MsgBox "进入医保结算记录" & i & "-" & Format(Val(strVarr(i)) + 1, "#####;-####0;0;0"), vbOKOnly, gstrSysName
            
            strSQL = strSQL & Format(Val(strVarr(i)) + 1, "#####;-####0;0;0") & ")"
        End If
        gstrSQL = strSQL
        ExecuteProcedure_兴成 "插入保险结算记录"
        InsertIntoYBK = True
        Exit Function
    End If
    '插入明细记录
    '门诊:交易流水号|  处方号|项目代码|项目名称|项目类别|单价|数量|费用总金额|医保内费用|医保外费用             |药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|MAC2
    '住院:交易流水号|项目序号|项目代码|项目名称|项目类别|单价|数量|费用总金额|医保内费用|医保外费用|费用发生日期|药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|MAC2
    '过程参数:
    '   结帐ID_IN,性质_IN,交易流水号_IN,处方号_IN,项目代码_IN,项目名称_IN,项目类别_IN，单价_IN，
    '   数量_IN,费用总金额_IN,医保内费用_IN,医保外费用_IN,药品甲类金额_IN,乙类医保金额_IN，乙类自负金额_IN，
    '   自费药品金额_IN，诊疗医保金额_IN，特检特治金额_IN，特检特治自负_IN，诊疗非医保额_IN，服务医保金额_IN，服务非医保额_IN，费用发生日期_IN，MAC2_IN，
    i = 0
    strSQL = "ZL_医保结算明细记录_INSERT("
    If bln门诊 Then
        strSQL = strSQL & g病人身份_兴成.结帐ID & ","
        strSQL = strSQL & "1" & ","
    Else
        strSQL = strSQL & IIf(g病人身份_兴成.结帐ID = 0, g病人身份_兴成.病人ID, g病人身份_兴成.结帐ID) & ","
        strSQL = strSQL & IIf(g病人身份_兴成.结帐ID = 0, 3, 2) & ","
    End If
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    
    
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 10000, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    
    '陈宏悦于20050315修改，由于更新保险结算明细记录表时“数量”字段不需要/100；将val(strVarr(i))/100修改为val(strVarr(i))
    
    strSQL = strSQL & "" & Format(Val(strVarr(i)), "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    If Not bln门诊 Then
        '住院多费用发生日期
        str发生日期 = zlCommFun.AddDate(strVarr(i)): i = i + 1
        If Not IsDate(str发生日期) Then
            str发生日期 = ""
        End If
    Else
        str发生日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    End If
    '药品甲类金额|乙类药品医保金额|乙类药品自负金额|自费药品金额|诊疗医保金额|特检特治医保金额|特检特治自负金额|诊疗非医保金额|服务医保金额|服务非医保金额|MAC2
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & IIf(str发生日期 = "", "NULL", "to_date('" & str发生日期 & "','yyyy-mm-dd')") & ","
    strSQL = strSQL & "'" & strVarr(i) & "',"
    strSQL = strSQL & "" & lng序号 & ")"
    gstrSQL = strSQL
    ExecuteProcedure_兴成 "插入保险结算记录"
    InsertIntoYBK = True
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    InsertIntoYBK = False
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
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取个人帐户支付", g病人身份_兴成.结帐ID)
    If Not rsTemp.EOF Then
        获取个人帐户支付 = Nvl(rsTemp!冲预交, 0)
    End If
    
End Function

Public Function 门诊虚拟结算_兴成(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    
    Dim str明细 As String
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
    
    g病人身份_兴成.费用总额 = 0
    str明细 = ""
    
    '第一步:汇总费用
    DebugTool "门诊虚拟结算,第一步:汇总费用"
    With rs明细
        If rs明细.RecordCount = 0 Then ShowMsgbox "未输入相关的费用记录!": Exit Function
        Do While Not .EOF
            g病人身份_兴成.费用总额 = g病人身份_兴成.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    
    '第二步:写入参明细文件
    '写入参文件
    DebugTool "门诊虚拟结算,第二步:准备写入参明细文件"
    If WriteINParaFile(rs明细, True, False) = False Then
        DebugTool "         入参文件明细写入失败"
        Exit Function
    End If
    DebugTool "         完成写入入参文件明细"
    
    
    '第三步:门诊虚拟结算分解
    DebugTool "门诊虚拟结算,第三步:门诊虚拟结算分解"
    If 业务请求_兴成(兴成_门诊费用预分解, "", "") = False Then
        DebugTool "             门诊虚拟结算分解失败"
        Exit Function
    End If
    DebugTool "             门诊虚拟结算分解成功"
    
    '第四步:分解相关出参结果
    
    DebugTool "门诊虚拟结算,第二步:分解相关出参结果"
    
    If ReadOutParaFile(g结算数据, True, True) = False Then
        DebugTool "     分解相关出参结果失败!"
        Exit Function
    End If
    DebugTool "     分解相关出参结果成功!"
    
    If Format(g病人身份_兴成.费用总额, "#####0.00;-####0.00;0;0") <> Format(g结算数据.交易费用总额, "#####0.00;-####0.00;0;0") Then
        ShowMsgbox "费用总额不等,不能结算!" & vbCrLf & _
                " HIS费用总额:" & Format(g病人身份_兴成.费用总额, "#####0.00;-####0.00; ;") & _
                " 虚拟费用总额:" & Format(g结算数据.交易费用总额, "#####0.00;-####0.00; ;")
        Exit Function
    End If
    str结算方式 = ""
    
    '陈宏悦悦于20050321修改,由于中心的结算方式:在进行门诊业务时:如果个人帐户的余额大于本次交易额,从个人帐户支付
    
    With g结算数据
      
     '陈宏悦于20050320修改,由于慢性病的结算方式不一样,则以现金方式支付
     If g病人身份_兴成.慢病标识 <> "1" Then
        
         '陈宏悦于20050403修改增加，因为考虑个人帐户的余额小于当前个人帐户支付的金额
         If (.交易费用总额 - .个人应付总额 + .乙类药品自负额 + .特检特治自负) - .个人帐户余额 <= 0 Then
             str结算方式 = str结算方式 & "个人帐户;" & Format(.交易费用总额 - .个人应付总额 + .乙类药品自负额 + .特检特治自负, "####0.00;-####0.00;0;0") & ";1"
         Else
             str结算方式 = str结算方式 & "个人帐户;" & Format(.个人帐户余额, "####0.00;-####0.00;0;0") & ";1"
         End If
     Else
      
      '陈宏悦于20050402修改,由于慢性病患者可以按非慢性病形式就医结算
        If blnmxb = True Then
         str结算方式 = str结算方式 & "个人帐户;" & Format(.交易费用总额 - .个人应付总额, "####0.00;-####0.00;0;0") & ";1"
        Else
          If (.交易费用总额 - .个人应付总额 + .乙类药品自负额 + .特检特治自负) - .个人帐户余额 <= 0 Then
             str结算方式 = str结算方式 & "个人帐户;" & Format(.交易费用总额 - .个人应付总额 + .乙类药品自负额 + .特检特治自负, "####0.00;-####0.00;0;0") & ";1"
          Else
             str结算方式 = str结算方式 & "个人帐户;" & Format(.个人帐户余额, "####0.00;-####0.00;0;0") & ";1"
          End If
        End If
        
     End If
     
    End With
    
    DebugTool "门诊虚拟结算成功,结算方式：" & str结算方式
    门诊虚拟结算_兴成 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function 门诊结算_兴成(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    Dim StrInput As String, strOutput As String
    Dim lng病人ID  As Long
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strArr As Variant
    Dim 门诊结算数据 As 结算数据
    
    
    Err = 0: On Error GoTo errHandle

    Call DebugTool("进入门诊结算")

    gstrSQL = "" & _
        "   Select a.*,a.付数*a.数次 as 数量,a.实收金额/(nvl(a.付数,1)*nvl(a.数次,1)) as 单价 " & _
        "   From 门诊费用记录 a " & _
        "   Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"

    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取明细记录", lng结帐ID)

    If rs明细.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "没有填写收费记录"
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")

    If g病人身份_兴成.病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人还没有经过身份验证，不能进行医保结算。"
        Exit Function
    End If
    g病人身份_兴成.结帐ID = lng结帐ID
    
    '陈宏悦于20050310修改，由于“g病人身份_兴成.费用总额”在上次调用后没有将其置零
    
    g病人身份_兴成.费用总额 = 0
    
    '第一步:汇总费用
    DebugTool "门诊结算,第一步:汇总费用"
    With rs明细
        If rs明细.RecordCount = 0 Then ShowMsgbox "未输入相关的费用记录!": Exit Function
        Do While Not .EOF
            g病人身份_兴成.费用总额 = g病人身份_兴成.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    gcnOracle_兴成.BeginTrans
    gcnSQLSEVER_兴成.BeginTrans
    '第二步:写入参明细文件
    '写入参文件
    DebugTool "门诊结算,第二步:准备写入参明细文件"
    
    '陈宏悦于20050315修改，这是门诊结算_兴成()：将true 修改为false
    
    If WriteINParaFile(rs明细, False, False) = False Then
        DebugTool "         入参文件明细写入失败"
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    DebugTool "         完成写入入参文件明细"
    
    
    '第三步:门诊虚拟结算分解
    DebugTool "门诊结算,第三步:门诊结算分解"
    If 业务请求_兴成(兴成_门诊费用预分解, "", "") = False Then
        DebugTool "             门诊结算分解失败"
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    DebugTool "             门诊结算分解成功"
    
    '第四步:分解相关出参结果
    
    DebugTool "门诊结算,第四步:分解相关出参结果"
    
    If ReadOutParaFile(门诊结算数据, True, False) = False Then
        DebugTool "     分解相关出参结果失败!"
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    DebugTool "     分解相关出参结果成功!"
    
    
    If Format(g结算数据.交易费用总额, "#####0.00;-####0.00;0;0") <> Format(门诊结算数据.交易费用总额, "#####0.00;-####0.00;0;0") Then
        ShowMsgbox "费用总额不等,不能结算!" & vbCrLf & _
                " 虚拟结算费用总额:" & Format(g结算数据.交易费用总额, "#####0.00;-####0.00; ;") & _
                " 正式结算费用总额:" & Format(门诊结算数据.交易费用总额, "#####0.00;-####0.00; ;")
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    
    '第五步: 门诊支付确认
    DebugTool "门诊结算,第五步: 门诊支付确认"
    '   交易流水号|IC卡号|费用总金额|医保内总费用|医保外总费用|个人帐户支付金额|现金支付金额

    StrInput = Rpad(Substr(InitInfor_兴成.医院编码, 1, 8), 8, " ") & Lpad(Substr(g病人身份_兴成.结帐ID, 1, 12), 12, "0")
    StrInput = StrInput & "|" & g病人身份_兴成.IC卡号
    StrInput = StrInput & "|" & Int(Format((门诊结算数据.交易费用总额 * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((门诊结算数据.医保范围金额 * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((门诊结算数据.非医保范围额 * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((cur个人帐户 * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format(((门诊结算数据.交易费用总额 - cur个人帐户) * 100), "####0.00;-####0.00;0;0"))
    
    g结算数据 = 门诊结算数据
    
'    MsgBox strInput, vbOKOnly, "zlsoft"
    
    If 业务请求_兴成(兴成_普通门诊支付确认, StrInput, strOutput) = False Then
        DebugTool "     普通门诊支付确认失败!"
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    DebugTool "     普通门诊支付确认成功!"
    
    
    '提交相关的交易数据
    
    '陈宏悦于20050402修改,为了调试方便添加
'    MsgBox strOutPut, vbOKOnly, "zlsoft"
       
    If InsertIntoSQLServer_门诊(strOutput) = False Then
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
        
    '填写结算表
    Call DebugTool("填写结算记录")

    
    '交易流水号|IC卡号|终端机编号|交易日期/时间|费用总金额|医保内总费用|医保外总费用|个人帐户支付金额|现金支付金额|扣减后个人帐户余额|MAC1
    strArr = Split(strOutput, "|")
    
'    MsgBox strOutPut, vbOKOnly, "zlsoft"
'    MsgBox strArr(10), vbOKOnly, "zlsoft"
     
    '陈宏悦于20050315修改填加，主要由于调用ExecuteProcedure("保存结算记录")时报错,因此截取strArr(10)5个字符串
     
     strArr(10) = Substr(strArr(10), 1, 16)
'     MsgBox strArr(10), vbOKOnly, "zlsoft"

   '插入保险结算记录
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(),住院次数_IN(住院:主页id),起付线(),封顶线_IN(),实际起付线_IN(扣减后个人帐户余额),
    '   发生费用金额_IN(费用总金额),全自付金额_IN(现金支付金额),首先自付金额_IN(医保外总费用),
    '   进入统筹金额_IN(医保内总费用),统筹报销金额_IN(门诊:住院:统筹支付金额),大病自付金额_IN(住院:统筹自负金额),超限自付金额_IN(),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(交易流水号),主页ID_IN(主页id),中途结帐_IN,备注_IN(终端机编号|交易日期/时间|MAC1)

    
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    
    '陈宏悦于20050311修改,部分数据要除以100后,在存入我方数据库中
'
'    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_兴成核工业 & "," & lng病人id & "," & Year(zlDatabase.Currentdate) & "," & _
'            "NULL,NULL,NULL,NULL,null,0,0," & Format(Val(strArr(9)), "#####0.00;-####0.00;0;0") & "," & _
'            Format(Val(strArr(4)), "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(8)), "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(6)), "#####0.00;-####0.00;0;0") & "," & _
'            Format(Val(strArr(5)), "#####0.00;-####0.00;0;0") & " ,0,0,0," & Format(Val(strArr(7)), "#####0.00;-####0.00;0;0") & ",'" & _
'             strArr(0) & "',NULL,NULL,'" & strArr(2) & "|" & strArr(3) & "|" & strArr(10) & "')"
             
       
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_兴成核工业 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "NULL,NULL,NULL,NULL,null,0,0," & Format(Val(strArr(9)) / 100, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(4)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(8)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(6)) / 100, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(5)) / 100, "#####0.00;-####0.00;0;0") & " ,0,0,0," & Format(Val(strArr(7)) / 100, "#####0.00;-####0.00;0;0") & ",'" & _
             strArr(0) & "',NULL,NULL,'" & strArr(2) & "|" & strArr(3) & "|" & strArr(10) & "')"

    'MsgBox gstrSQL, vbOKOnly, "zlsoft"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    gcnOracle_兴成.CommitTrans
    gcnSQLSEVER_兴成.CommitTrans
    门诊结算_兴成 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Public Function 门诊结算冲销_兴成(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额

    Dim intMouse As Integer
    Dim lng冲销ID  As Long
    Dim rs明细 As New ADODB.Recordset
    Dim rs原明细 As New ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    Dim lng病人id1 As Long
    
    门诊结算冲销_兴成 = False


    '身份验证
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If 身份标识_兴成(2, lng病人id1) = "" Then
        If lng病人id1 = 0 Then
            Err.Raise 9000, gstrSysName, "你不是当前持卡人!"
            Screen.MousePointer = intMouse
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo errHand:
    Screen.MousePointer = intMouse
        
    '第一步:确定冲销ID值
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")

    '第二步:确定冲销和原始单据的明细记录

    gstrSQL = "Select * From 门诊费用记录 " & _
        " Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取冲销记录", lng冲销ID)
    g病人身份_兴成.费用总额 = 0
    g病人身份_兴成.结帐ID = lng结帐ID
    
    
    gstrSQL = "Select * From 门诊费用记录 where  结帐ID =[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rs原明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取冲销记录", lng结帐ID)

    gcnOracle_兴成.BeginTrans
    gcnSQLSEVER_兴成.BeginTrans
    
    
    '第三步:将原始记录中的接要作为冲销单据的接要，并打上上传标志
    With rs明细
        Do While Not .EOF
            rs原明细.Filter = 0
            rs原明细.Filter = "no='" & Nvl(!NO) & "' and 记录性质=" & Nvl(!记录性质, 0) & " and 序号=" & Nvl(!序号, 0)         '& "' and 执行状态=" & Nvl(!执行状态, 0)
            If rs原明细.EOF Then
                Err.Raise 9000, gstrSysName, "没有找到原始记录!" & "no='" & Nvl(!NO) & "' and 记录性质=" & Nvl(!记录性质, 0) & " and 序号='" & Nvl(!序号, 0) & "' and 执行状态=" & Nvl(!执行状态, 0)
                gcnOracle_兴成.RollbackTrans
                gcnSQLSEVER_兴成.RollbackTrans
                Exit Function
            End If
            '写上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rs原明细!摘要) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上上传标志")
            g病人身份_兴成.费用总额 = g病人身份_兴成.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
        
        
    '第四步:冲销中间库的相关记录
    
    '插入冲销费用结算记录
    '过程参数:性质_IN,原结帐ID_IN,现结帐ID_IN
    gstrSQL = "ZL_医保结算_冲销("
    gstrSQL = gstrSQL & "1,"
    gstrSQL = gstrSQL & lng结帐ID & ","
    gstrSQL = gstrSQL & lng冲销ID & ")"
    
    ExecuteProcedure_兴成 "冲销费用结算记录"
    
    
    
    '第五步:读取保险结算记录中的数据，同时调用退费函数
    
    gstrSQL = "Select * from 保险结算记录 where 性质=1 and 记录id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取中心单据号"
    
    If rsTemp.EOF Then
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Err.Raise 9000, gstrSysName, "不存在结算记录,不能冲销!"
        Exit Function
    End If
    
    '交易流水号|IC卡号|退费金额
    '陈洪悦于20050315修改，主要由于PoliBackCost()函数入口参数错误
    
'     strInput = Nvl(rsTemp!支付顺序号) & "|"
'     strInput = strInput & Nvl(rsTemp!支付顺序号)
'     strInput = strInput & g病人身份_兴成.IC卡号
'     strInput = strInput & Nvl(rsTemp!个人帐户支付)
      
     StrInput = Nvl(rsTemp!支付顺序号) & "|"
'    strInput = strInput & Nvl(rsTemp!支付顺序号)
     StrInput = StrInput & g病人身份_兴成.IC卡号 & "|"
     
     StrInput = StrInput & Nvl(rsTemp!个人帐户支付) * 100
     
    If 业务请求_兴成(兴成_门诊退费函数, StrInput, strOutput) = False Then
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    
    
    '第六步:保存门诊退费的相关记录
    
    '插入保险结算记录
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(),住院次数_IN(住院:主页id),起付线(),封顶线_IN(),实际起付线_IN(扣减后个人帐户余额),
    '   发生费用金额_IN(费用总金额),全自付金额_IN(现金支付金额),首先自付金额_IN(医保外总费用),
    '   进入统筹金额_IN(医保内总费用),统筹报销金额_IN(门诊:住院:统筹支付金额),大病自付金额_IN(住院:统筹自负金额),超限自付金额_IN(),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(交易流水号),主页ID_IN(主页id),中途结帐_IN,备注_IN(终端机编号|交易日期/时间|MAC1)
    DebugTool "结算交易提交成功,并开始保存保险结算记录"

    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_兴成核工业 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,null,null,0,0," & -1 * Nvl(rsTemp!实际起付线, 0) & "," & _
           -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * -1 * Nvl(rsTemp!首先自付金额, 0) & "," & _
           -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & "," & -1 * Nvl(rsTemp!大病自付金额, 0) & ",0," & -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & _
           Nvl(rsTemp!支付顺序号, 0) & " ',NULL,NULL,'" & Nvl(rsTemp!备注) & "')"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    
    '第七步:提交相关交易
    If InsertIntoSQLServer_门诊冲销(lng冲销ID) = False Then
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    '---------------------------------------------------------------------------------------------
    gcnOracle_兴成.CommitTrans
    gcnSQLSEVER_兴成.CommitTrans
    门诊结算冲销_兴成 = True
    Exit Function
errHand:
    gcnOracle_兴成.RollbackTrans
    gcnSQLSEVER_兴成.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Private Function Get交易代码(ByVal intType As 业务类型_兴成, Optional bln读名称 As Boolean = False) As String
    '代码暂没用
    Select Case intType
        Case 兴成_政策机服务启动
            Get交易代码 = IIf(bln读名称, "兴成_政策机服务启动", "01")
        Case 兴成_政策机服务停止
            Get交易代码 = IIf(bln读名称, "兴成_政策机服务停止", "02")
        Case 兴成_POS机启动
            Get交易代码 = IIf(bln读名称, "兴成_POS机启动", "03")
        Case 兴成_POS机停止
            Get交易代码 = IIf(bln读名称, "兴成_POS机停止", "04")
        Case 兴成_获取持卡人信息
            Get交易代码 = IIf(bln读名称, "兴成_获取持卡人信息", "05")
        Case 兴成_JbylReadIC
            Get交易代码 = IIf(bln读名称, "兴成_JbylReadIC", "06")
        Case 兴成_门诊费用预分解
            Get交易代码 = IIf(bln读名称, "兴成_门诊费用预分解", "07")
        Case 兴成_普通门诊支付确认
            Get交易代码 = IIf(bln读名称, "兴成_普通门诊支付确认", "08")
        Case 兴成_门诊退费函数
            Get交易代码 = IIf(bln读名称, "兴成_门诊退费函数", "09")
        Case 兴成_住院费用预分解
            Get交易代码 = IIf(bln读名称, "兴成_住院费用预分解", "10")
        Case 兴成_住院支付确认
            Get交易代码 = IIf(bln读名称, "兴成_住院支付确认", "11")
        Case Else
            Get交易代码 = IIf(bln读名称, "错误的交易代码", "-1")
    End Select
End Function
Public Function 业务请求_兴成(ByVal intType As 业务类型_兴成, strInputString As String, strOutPutstring As String) As Boolean
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
    
    str交易代码 = Get交易代码(intType, True)
    
    StrInput = strInputString
    DebugTool "进入业务请求函数(业务类型代码为:" & intType & " 业务名称：" & str交易代码 & ")" & vbCrLf & "        输入参数为:" & strInputString
    
    
    业务请求_兴成 = False
    If InitInfor_兴成.模拟数据 Then
        '读取模拟数据
        Read模拟数据 intType, StrInput, strOutPutstring
         业务请求_兴成 = True
        Exit Function
    End If
    strArr = Split(strInputString, SP_STR)
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    
    Err = 0
    On Error GoTo errHand:
    
    Select Case intType
        Case 兴成_政策机服务启动
            '启动正策服务,相关参数
            '    a. 系统目录（完整）[50] ――类型：String    值：一般指的是c:\
            '    b. 医院代码 [8]         ――类型：String  长度为8位
            '    c. ODBC数据源名字[ ]   ――类型：String     值：一般指的是ODBC的DSN
            '    d. ODBC用户名[ ]       ――类型：String
            '    e. ODBC用户口令[ ]     ――类型：String
            lngReturn = StartPolicy(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4))
            
            Select Case lngReturn
            Case 0          '     ――启动政策机成功
            Case 100        '     ――系统目录错误
                ShowMsgbox "系统目录错误,请检查参数设置是否正确!"
                Exit Function
            Case 101        '     ――医疗机构代码错误
                ShowMsgbox "医疗机构代码错误,请检查医院编码是否设置正确!"
                Exit Function
            Case 102        '     ――连接数据库错误
                ShowMsgbox "连接数据库错误,请检查参数设置中的ODBC配置是否正确!"
                Exit Function
            Case -11        '――未经正确授权
                ShowMsgbox "连接数据库错误,请与医保开发商联系!"
                Exit Function
            Case Else
                ShowMsgbox "未知的错误,返回的序号为:" & lngReturn
                Exit Function
            End Select
            strOutput = ""
            
        Case 兴成_政策机服务停止
        
            lngReturn = StopPolicy()
            Select Case lngReturn
            Case 0          '     停止政策机成功
            Case 103        '――断开数据库连接错误
                ShowMsgbox "――断开数据库连接错误,可能已经断开连接了!"
                Exit Function
            Case Else
                ShowMsgbox "未知的错误,返回的序号为:" & lngReturn
                Exit Function
            End Select
            strOutput = ""
        Case 兴成_POS机启动
            lngReturn = StartPos()
            If lngReturn <> 0 Then
                ShowMsgbox "启动POS机错误:" & vbCrLf & "错误描述:" & lngReturn
                Exit Function
            End If
            strOutput = ""
        Case 兴成_POS机停止
            lngReturn = StopPos()
            If lngReturn <> 0 Then
                ShowMsgbox "停用POS机错误:" & vbCrLf & "错误描述:" & lngReturn
                Exit Function
            End If
            strOutput = ""
        Case 兴成_获取持卡人信息
            strOutput = Space(200)
            lngReturn = GetPersonCommInfo(strOutput)
            
         ' 陈宏悦于20050310修改，根据返回参数临时修改：将原strOutPut修改为strReturn
         
            Select Case lngReturn
            Case 0      '――成功
            Case 8      '――已进入黑名单
                ShowMsgbox "该病人已经进入黑名单!"
                Exit Function
            Case 9      '――划拨累计金额大于工资总额，即为大额医疗卡，需要没收交医保中心处理。
                ShowMsgbox "划拨累计金额大于工资总额，即为大额医疗卡，需要没收交医保中心处理!"
                Exit Function
            
            '陈宏悦于20050321修改，根据慢性病的错误返回值添加
                
            Case -90  '--该慢性病卡没进行慢性病升级，需到医保中心处理
                 ShowMsgbox "该慢性病卡没进行慢性病升级，需到医保中心处理！"
                 Exit Function
            Case -91  '当前医疗机构不是该卡的慢性病定点医疗机构
                 ShowMsgbox "当前医疗机构不是该卡的慢性病定点医疗机构！"
                 Exit Function
            Case -65536
                 ShowMsgbox "请插入医保病人的医保卡！"
                 Exit Function
            Case Else
                ShowMsgbox "其他错误," & vbCrLf & "错误描述为:" & lngReturn
                Exit Function
            End Select
            strOutput = Trim(strOutput)
        Case 兴成_JbylReadIC
            strOutput = Space(200)
           
            '提示返回参数
            'MsgBox JbylReadIC(strOutPut), vbOKOnly, "zlhis"
            lngReturn = JbylReadIC(strOutput)
        
            '陈宏悦于20050310修改，根据返回参数临时修改：将原strOutPut修改为strReturn
        
            Select Case lngReturn
            Case 0      '――成功
            Case Else
                ShowMsgbox "读取IC失败," & vbCrLf & "错误描述为:" & lngReturn
                Exit Function
            End Select
            strOutput = Trim(strOutput)
        Case 兴成_门诊费用预分解
            lngReturn = Poli_Divide()
            Select Case lngReturn
                Case 0               '――成功
                Case 1               '――Poli_Divide.in文件入口参数错误
                    ShowMsgbox "Poli_Divide.in文件入口参数错误,请与接口商联系!"
                    Exit Function
                Case 104             '――Poli_Divide.in或Poli_Divide.out?Poli_Divide.log文件打不开
                    ShowMsgbox "Poli_Divide.in或Poli_Divide.out,Poli_Divide.log文件打不开,请与接口商联系!"
                    Exit Function
                Case 105             '――写出口参数错误
                    ShowMsgbox "写出口参数错误，请与接口商联系!"
                    Exit Function
                Case -11               '――未经正确授权
                    ShowMsgbox "未经正确授权，请与接口商联系!"
                    Exit Function
                Case Else
                    ShowMsgbox "处理失败," & vbCrLf & "错误描述为:" & strReturn
                    Exit Function
            End Select
            strOutput = ""
        Case 兴成_普通门诊支付确认
        
            '由于调用Reg_Poli()函数时，程序强行退出；将strOutPut初始化试一试；陈宏悦于20050311修改
            
            strOutput = Space(200)
            lngReturn = Reg_Poli(strInValue(0), strOutput)
            Select Case lngReturn
                Case 0               '――成功
                Case 1               '――Poli_Divide.in文件入口参数错误
                    ShowMsgbox "Poli_Divide.in文件入口参数错误,请与接口商联系!"
                    Exit Function
                Case -11               '――未经正确授权
                    ShowMsgbox "未经正确授权，请与接口商联系!"
                    Exit Function
                Case Else
                    ShowMsgbox "处理失败," & vbCrLf & "错误描述为:" & strReturn
                    Exit Function
            End Select
            strOutput = Trim(strOutput)
        Case 兴成_门诊退费函数
            lngReturn = PoliBackCost(strInValue(0))
             Select Case lngReturn
                Case 0               '――成功
                Case -11               '――未经正确授权
                   ShowMsgbox "未经正确授权，请与接口商联系!"
                    Exit Function
                Case Else
                    ShowMsgbox "处理失败," & vbCrLf & "错误描述为:" & strReturn
                    Exit Function
            End Select
        Case 兴成_住院费用预分解
            lngReturn = Hosp_Divide()
            Select Case lngReturn
                Case 0               '――成功
                Case 1               '――Poli_Divide.in文件入口参数错误
                    ShowMsgbox "Hosp_Divide.in文件入口参数错误,请与接口商联系!"
                    Exit Function
                Case 104             '――Poli_Divide.in或Poli_Divide.out?Poli_Divide.log文件打不开
                    ShowMsgbox "Hosp_Divide.in 和Hosp_Divide.out,Hosp_Divide.log文件打不开,请与接口商联系!"
                    Exit Function
                Case 105             '――写出口参数错误
                    ShowMsgbox "写出口参数错误，请与接口商联系!"
                    Exit Function
                Case -11               '――未经正确授权
                    ShowMsgbox "未经正确授权，请与接口商联系!"
                    Exit Function
                Case Else
                    ShowMsgbox "处理失败," & vbCrLf & "错误描述为:" & strReturn
                    Exit Function
            End Select
            strOutput = ""
        Case 兴成_住院支付确认
             
            strOutput = Space(200)
            
       '陈宏悦于20050318修改，将原Reg_hospital()函数修改为Reg_Hospital()：由于电子版错误
       
            lngReturn = Reg_Hospital(strInValue(0), strOutput)
           
'            MsgBox strOutPut, vbOKOnly, "zlsoft"
            
            Select Case lngReturn
                Case 0               '――成功
                Case 1               '――Poli_Divide.in文件入口参数错误
                    ShowMsgbox "Hosp_Divide.in文件入口参数错误,请与接口商联系!"
                    Exit Function
                Case -11               '――未经正确授权
                    ShowMsgbox "未经正确授权，请与接口商联系!"
                    Exit Function
                Case Else
                    ShowMsgbox "处理失败," & vbCrLf & "错误描述为:" & strReturn
                    Exit Function
            End Select
            strOutput = Trim(strOutput)
    End Select
    
    strOutPutstring = strOutput
    业务请求_兴成 = True
    DebugTool "    输出参数为:" & strOutPutstring
     Exit Function
errHand:
    DebugTool "    输出参数为:" & strOutPutstring
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function 入院登记_兴成(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    Err = 0
    On Error GoTo errHand:
    gcnSQLSEVER_兴成.BeginTrans
    If InsertIntoData_住院登记(lng病人ID, lng主页ID) = False Then
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_兴成核工业 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    gcnSQLSEVER_兴成.CommitTrans
    入院登记_兴成 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle_兴成.RollbackTrans
    gcnSQLSEVER_兴成.RollbackTrans
    入院登记_兴成 = False
End Function

Public Function 入院登记撤销_兴成(lng病人ID As Long, lng主页ID As Long) As Boolean
  '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
            
    '刘兴宏:20040923增加的
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
     Err = 0
    On Error GoTo errHand
    
    DebugTool "进入扩院登撤消接口"
    
    入院登记撤销_兴成 = False
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用，不能撤消入院登记"
        Exit Function
    End If
    gstrSQL = "Select * from HLD_ZYBRXX where zyh='" & lng病人ID & "_" & lng主页ID & "' and C_zt<>'0'"
    rsTemp.Open gstrSQL, gcnSQLSEVER_兴成
    If Not rsTemp.EOF Then
        ShowMsgbox "该病人的住院信息已经上传中心, 不能再删除!"
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_兴成核工业 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
        
    
    '删除SQLServer中的相关数据
    gstrSQL = " Delete HLD_ZYBRXX where zyh='" & lng病人ID & "_" & lng主页ID & "' and C_zt='0'"
    gcnSQLSEVER_兴成.Execute gstrSQL
    
    
    
    '更新医保帐户
    DebugTool "取消成功"
    入院登记撤销_兴成 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_兴成(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0:    On Error GoTo errHand:
    出院登记_兴成 = False
    
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "当前病人不存在未结费用，请在入院撤消即可"
        Exit Function
    End If
    If frm出院设置_兴成.ShowCard(lng病人ID, lng主页ID) = False Then Exit Function
    
    '改变当前状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_兴成核工业 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_兴成 = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function 出院登记撤销_兴成(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
  '出院登记撤消
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
    出院登记撤销_兴成 = False
    
    Err = 0: On Error GoTo errHand:
     
     If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "该病人已经出院结算了,不能再取消出院!"
        Exit Function
     End If
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_兴成核工业 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    出院登记撤销_兴成 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_兴成(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
  '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)

    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String
    Dim rs明细 As New ADODB.Recordset
    
    Dim lng主页ID As Long
    Dim dbl费用总额 As Double
    Dim strArr As Variant, strTmpArr As Variant
    
    Dim str结算方式  As String, str住院号 As String
    Dim obj结算 As 结算数据
    Dim dbl个人帐户 As Double
    
    住院结算_兴成 = False


    Err = 0: On Error GoTo errHand:
    Call DebugTool("进入住院结算")


    If g病人身份_兴成.病人ID <> lng病人ID Then
        Err.Raise 9000, gstrSysName, "该病人没有完成医保的预结算操作，不能进行结算。"
        Exit Function
    End If

    gstrSQL = "Select 当前状态 From 保险帐户  where 病人ID=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断当前的住院状态!"

    If Nvl(rsTemp!当前状态, 0) = 1 Then
        Err.Raise 9000, gstrSysName, "当前病人还处于在院状态,请出院后再结算!"
        Exit Function
    End If


    With g结算数据
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID)
        If IsNull(rsTemp("主页ID")) = True Then
            Err.Raise 9000, gstrSysName, "只有住院病人才可以使用医保结算。"
            Exit Function
        End If
        lng主页ID = rsTemp("主页ID")
    End With

   '更新结算标志
    gstrSQL = "Select ID,结帐金额 as 实收金额 From 住院费用记录 where 结帐ID=" & lng结帐ID
    zlDatabase.OpenRecordset rs明细, gstrSQL, "更证结算标志"
    dbl费用总额 = 0
    With rs明细
        Do While Not .EOF
                '为病人费用记录打上标记，以便随时上传
                 'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                 gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,NULL)"
                 DebugTool "     打上明细标志:SQL=" & gstrSQL
                 zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
                 DebugTool " 打上明细标志:更新病人费用记录成功:SQL=" & gstrSQL
                 dbl费用总额 = dbl费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    
    If dbl费用总额 <> g结算数据.交易费用总额 Then
        Err.Raise 9000, gstrSysName, "虚拟结算数据的费用总额与本次结算的费用总额不等，请检查处方是否正确!"
        Exit Function
    End If
    
 
    
    g病人身份_兴成.结帐ID = lng结帐ID
    '第一步:住院结算支付确认
    '   交易流水号|IC卡号|费用总金额|医保内总费用|医保外总费用|统筹支付金额|统筹自负金额|个人帐户支付金额|现金支付金额
  '     交易流水号|IC卡号|费用总金额|医保内总费用|医保外总费用|                         |个人帐户支付金额|现金支付金额
   
    StrInput = Rpad(Substr(InitInfor_兴成.医院编码, 1, 8), 8, " ") & Lpad(Substr(g病人身份_兴成.结帐ID, 1, 12), 12, "0")
    StrInput = StrInput & "|" & g病人身份_兴成.IC卡号
    StrInput = StrInput & "|" & Int(Format((g结算数据.交易费用总额 * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((g结算数据.医保范围金额 * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((g结算数据.非医保范围额 * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((g结算数据.统筹支付金额 * 100), "####0.00;-####0.00;0;0"))
    
    '陈宏悦于20050315修改,将(g结算数据.共付段金额 - g结算数据.统筹支付金额 * 100)修改如下
    
    StrInput = StrInput & "|" & Int(Format(((g结算数据.共付段金额 - g结算数据.统筹支付金额) * 100), "####0.00;-####0.00;0;0"))
    dbl个人帐户 = 获取个人帐户支付
    StrInput = StrInput & "|" & Int(Format((dbl个人帐户 * 100), "#####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format(((g结算数据.交易费用总额 - dbl个人帐户 - g结算数据.统筹支付金额) * 100), "####0.00;-####0.00;0;0"))
    
    DebugTool "第一步:住院结算支付确认!"
    
    If 业务请求_兴成(兴成_住院支付确认, StrInput, strOutput) = False Then
        DebugTool "     住院支付确认确认失败!"
        Exit Function
    End If
    DebugTool "     住院支付确认确认成功!"
    
    Err = 0: On Error GoTo ErrHand1:
    gcnOracle_兴成.BeginTrans
    gcnSQLSEVER_兴成.BeginTrans
    
    If InsertIntoData_住院(strOutput, lng病人ID, lng主页ID) = False Then
        gcnOracle_兴成.RollbackTrans
        gcnSQLSEVER_兴成.RollbackTrans
        Exit Function
    End If
    
    '第二步:分解确认后结果
    '   交易流水号|IC卡号|终端机编号|交易日期/时间|费用总金额|医保内总费用|医保外总费用|统筹支付金额|统筹自负金额|
    '   个人帐户支付金额|现金支付金额|扣减后个人帐户余额|MAC1
    strArr = Split(strOutput, "|")
    
    
    '填写结算表
    Call DebugTool("填写结算记录")
    DebugTool "第二步:开始保存保险结算记录"
    

   '插入保险结算记录
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(),累计统筹报销_IN(),住院次数_IN(住院:主页id),起付线(),封顶线_IN(),实际起付线_IN(扣减后个人帐户余额),
    '   发生费用金额_IN(费用总金额),全自付金额_IN(现金支付金额),首先自付金额_IN(医保外总费用),
    '   进入统筹金额_IN(医保内总费用),统筹报销金额_IN(门诊:住院:统筹支付金额),大病自付金额_IN(住院:统筹自负金额),超限自付金额_IN(),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(交易流水号),主页ID_IN(主页id),中途结帐_IN,备注_IN(终端机编号|交易日期/时间|MAC1)
    
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    '陈宏悦于20050315修改,由于报错截取strArr(12)字符串
     
    strArr(12) = Substr(strArr(12), 1, 16)
    
    '陈宏悦于20050403增加修改，因为本地医保中心存在超限自付部分，需进入商业医保统筹
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_兴成核工业 & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,NULL,null,0,0," & Format(Val(strArr(11)) / 100, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(4)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(10)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(6)) / 100, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(5)) / 100, "#####0.00;-####0.00;0;0") & " ," & Format(Val(strArr(7)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(8)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(g结算数据.封顶线以上自负金额, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(9)) / 100, "#####0.00;-####0.00;0;0") & ",'" & _
            strArr(0) & " ',NULL,NULL,'" & strArr(2) & "|" & strArr(3) & "|" & strArr(12) & "')"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算记录")
    '---------------------------------------------------------------------------------------------
    '第三步:将临时数据更正过来
    '   原结帐ID_IN,现结帐ID_IN
    
    gstrSQL = "ZL_医保结算_住院确认("
    gstrSQL = gstrSQL & g病人身份_兴成.病人ID & ","
    gstrSQL = gstrSQL & lng结帐ID & ")"
    
    DebugTool "第三步:将临时数据更正过来:" & gstrSQL
    
    ExecuteProcedure_兴成 "更正临时数据!"
 
    gcnOracle_兴成.CommitTrans
    gcnSQLSEVER_兴成.CommitTrans

    住院结算_兴成 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    Exit Function
ErrHand1:
    gcnOracle_兴成.RollbackTrans
    gcnSQLSEVER_兴成.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function 住院结算冲销_兴成(lng结帐ID As Long) As Boolean
     '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    Err.Raise 9000, gstrSysName, "本医保不支持住院结算冲销,具体请咨询接口商!"
    住院结算冲销_兴成 = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function 处方登记_兴成(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
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
    Dim str处方记录号 As String, str摘要 As String
    Dim strArr


    处方登记_兴成 = False


   '读出该张单据的费用明细
  gstrSQL = "" & _
              "  Select a.收费细目ID,b.编码,b.名称" & _
              "  From 住院费用记录 A,收费细目 B,病案主页 C" & _
              "  where A.NO=[1] and A.记录性质=[2] and A.记录状态 = [3]" & _
              "        and A.收费细目ID=B.ID and A.病人ID=C.病人ID  and A.主页ID=C.主页ID And C.险类=[4]" & _
              "  Order by A.病人ID,A.NO,A.发生时间"

    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "处方明细上传", str单据号, lng记录性质, lng记录状态, TYPE_兴成核工业)
    Err = 0:    On Error GoTo errHand:
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            gstrSQL = "Select * From 保险支付项目 where 险类=[1] and 收费细目id=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "确定医保支付项目", TYPE_兴成核工业, CLng(Nvl(!收费细目ID, 0)))
            If rsTemp.EOF Then
                ShowMsgbox "注意：" & vbCrLf & "   收费细目为:[" & Nvl(!编码) & "]" & Nvl(!名称) & " 还未进行医保对码!"
            End If
            .MoveNext
        Loop
    End With
    处方登记_兴成 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    
End Function


Private Function Read模拟数据(ByVal int业务类型 As 业务类型_兴成, ByVal strInputString As String, ByRef strOutPutstring As String)
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
                                strText = "" & vbTab & "|"
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
'    If InStr(1, strOutPutstring, "@$") <> 0 Then
'        strOutPutstring = Split(strOutPutstring, "@$")(1)
'    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_兴成(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_兴成, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Function 住院虚拟结算_兴成(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    'rsExse:字符集
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    
    Dim lng主页ID As Long, StrInput As String, strOutput  As String
    Dim str住院号 As String, str结算方式 As String, strSQL As String
    Dim lng病人id1 As Long
    Dim intMouse As Integer
    
    Dim strArr As Variant

    Err = 0: On Error GoTo errHand:
    
    g病人身份_兴成.病人ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    intMouse = Screen.MousePointer


    gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID)

    If IsNull(rsTemp("主页ID")) = True Then
        MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    lng主页ID = rsTemp("主页ID")

    If bln结帐处 Then
        Screen.MousePointer = 1
        If 身份标识_兴成(4, lng病人id1) = "" Then
            Screen.MousePointer = intMouse
            住院虚拟结算_兴成 = ""
            Exit Function
        End If
        Screen.MousePointer = intMouse
        If lng病人ID <> lng病人id1 Then
            ShowMsgbox "不是当前要结算的病人!"
            Exit Function
        End If
    End If

    
    Screen.MousePointer = vbHourglass
    
    
    strSQL = "" & _
        "   Select A.收费细目ID,a.付数*a.数次 as 数量,A.实收金额/(nvl(a.付数,1)*nvl(a.数次,1)) as 单价,a.实收金额 " & _
        "   From 住院费用记录 A" & _
        "   where 病人id=" & lng病人ID & " and 主页id =" & lng主页ID & _
        "       and a.记录状态<>0 and  A.记帐费用=1 and nvl(a.实收金额,0)<>0  And nvl(A.婴儿费,0)=0"
    
   strSQL = "" & _
    "   Select '' no ,收费细目ID, sysdate as 发生时间,sum(数量) as 数量,单价,sum(实收金额) as 实收金额 " & _
    "   From (" & strSQL & " ) " & _
    "   group by 收费细目ID,单价" & _
    "   having sum(数量)<>0"
    
    zlDatabase.OpenRecordset rs明细, strSQL, "获取明细记录"
    If rs明细.RecordCount = 0 Then
        ShowMsgbox "该病人未发生任何费用,不能结算!"
        Exit Function
    End If
    
    g病人身份_兴成.结帐ID = 0
    g病人身份_兴成.病人ID = lng病人ID
    
    '第一步:汇总费用
    DebugTool "住院虚拟结算,第一步:汇总费用"
    g病人身份_兴成.费用总额 = 0
    
    With rs明细
        If rs明细.RecordCount = 0 Then ShowMsgbox "未输入相关的费用记录!": Exit Function
        Do While Not .EOF
            g病人身份_兴成.费用总额 = g病人身份_兴成.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
        
  
    '第二步:写入参明细文件
    '写入参文件
    DebugTool "住院虚拟结算,第二步:准备写入参明细文件"
    If WriteINParaFile(rs明细, False, True) = False Then
        DebugTool "         入参文件明细写入失败"
        Exit Function
    End If
    DebugTool "         完成写入入参文件明细"
    
    
    '第三步:住院虚拟结算-结算分解
    DebugTool "住院虚拟结算,第三步:住院虚拟结算分解"
    If 业务请求_兴成(兴成_住院费用预分解, "", "") = False Then
        DebugTool "             住院虚拟结算分解失败"
        Exit Function
    End If
    DebugTool "             住院虚拟结算分解成功"
    
    '第四步:分解相关出参结果
    
    DebugTool "住院虚拟结算,第二步:分解相关出参结果"
    
    If ReadOutParaFile(g结算数据, False, True) = False Then
        DebugTool "     分解相关出参结果失败!"
        Exit Function
    End If
    DebugTool "     分解相关出参结果成功!"
    
    If Format(g病人身份_兴成.费用总额, "#####0.00;-####0.00;0;0") <> Format(g结算数据.交易费用总额, "#####0.00;-####0.00;0;0") Then
        ShowMsgbox "费用总额不等,不能结算!" & vbCrLf & _
                " HIS费用总额:" & Format(g病人身份_兴成.费用总额, "#####0.00;-####0.00; ;") & _
                " 虚拟费用总额:" & Format(g结算数据.交易费用总额, "#####0.00;-####0.00; ;")
        Exit Function
    End If
    
    str结算方式 = ""
    With g结算数据
        str结算方式 = str结算方式 & "个人帐户;" & .交易费用总额 - .统筹支付金额 - .个人应付总额 & ";1"
        str结算方式 = str结算方式 & "|统筹基金;" & .统筹支付金额 & ";0"
    End With
    住院虚拟结算_兴成 = str结算方式
    g病人身份_兴成.病人ID = lng病人ID   '表示该病人已经进行了虚拟结算
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function Open中间库_兴成() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    Dim strConn As String
    
    Open中间库_兴成 = False
    Err = 0: On Error Resume Next
        
    '连接医保中间库
    strServer = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_NAME"), "")
    strUser = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_USERNAME"), "")
    strPass = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_PASSWORD"), "")
    strConn = "dsn=" & strServer & ";uid=" & strUser & ";pwd=" & strPass & ";"
    
     
    Set gcnSQLSEVER_兴成 = New ADODB.Connection
    gcnSQLSEVER_兴成.Open strConn
    If Err <> 0 Then
        MsgBox "由于用户、口令或服务器指定错误，无法注册," & vbCrLf & "请检查参数中的数据源定义是否正确。", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    Err = 0: On Error GoTo errHand:
    
    '重新建立到医保服务器的公共连接
    '中间库连接
    gstrSQL = "select 参数名,参数值 from 保险参数 where  险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "兴成核工业医保", TYPE_兴成核工业)
    
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
    Set gcnOracle_兴成 = New ADODB.Connection
    If OraDataOpen(gcnOracle_兴成, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    Open中间库_兴成 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    
End Function
Public Function 医保设置_兴成(ByVal lng险类 As Long, ByVal lng医保中心 As Integer) As Boolean
    '功能： 该方法用于供相关应用部件调用配置连接医保数据服务器的连接串
    '返回：接口配置成功，返回true；否则，返回false
    
    Dim strConn As String
    Dim blnReturn As Boolean
    
    If frmSet兴成.参数设置 = False Then
        Exit Function
    End If
  
    If gcnOracle_兴成 Is Nothing And gcnSQLSEVER_兴成 Is Nothing Then
                blnReturn = True
    Else
        If Open中间库_兴成() Then
                blnReturn = True
        End If
    End If

    InitInfor_兴成.strPath_Get = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Get"), "C:\xcyb\get")
    InitInfor_兴成.strPath_Put = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Put"), "C:\xcyb\Put")
    InitInfor_兴成.strPath_In = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_In"), "C:\xcyb\In")
    InitInfor_兴成.strPath_Out = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Out"), "C:\xcyb\Out")
    InitInfor_兴成.strPath_System = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_System"), "C:\")
        
    InitInfor_兴成.ODBC_NAME = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_NAME"), "")
    InitInfor_兴成.ODBC_UserName = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_USERNAME"), "")
    InitInfor_兴成.ODBC_PassWord = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_PASSWORD"), "")
    
    
    InitInfor_兴成.存在读卡器 = Val(GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("读卡器"), "1")) = 1
    InitInfor_兴成.启用政策审核 = Val(GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("启用政策审核"), "1")) = 1
    
        
    医保设置_兴成 = blnReturn
End Function
Public Sub ExecuteProcedure_兴成(ByVal strCaption As String)
    '功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_兴成.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub



