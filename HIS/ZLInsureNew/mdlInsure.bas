Attribute VB_Name = "mdlInsure"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
'    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;3-公共部件增加GetNextNO();
'    99-所有交易增加附加参数(最新版)
Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As String
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000  'Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000   'Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000 'Browsing for Everything
Private Const CSIDL_NETWORK As Long = &H12

Private Const MAX_PATH = 260
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2
Private Const LVM_SETCOLUMNWIDTH = &H101E

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'输入法控制API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Const KLF_REORDER = &H8

'下列语句用于检测是否合法调用
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'对文本串进行加密或解密的函数
Public Declare Function EncryptStr Lib "FTP_Trans.dll" (ByVal SourceStr As String, ByVal Key As String, ByVal IsEncrypt As Boolean) As String

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

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetprivateprofileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, _
    ByVal lpDefault As String, ByVal lpRetrm_String As String, ByVal cbReturnString As Integer, ByVal FileName As String) As Integer
    
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门 As String
    站点 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public glngSys As Long                      '系统编号参数
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSQL As String                    '用着作为所有临时SQL语句

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrDbUser As String                 '当前数据库用户
Public gstrUserName As String               '当前用户姓名
Public gstr单位名称 As String
Public gbln特殊门诊 As Boolean              '西铝专用,用于返回是否为特殊门诊
Public gstr特殊病种 As String               '特殊病种类
Public gintDebug As Integer                 '保存从注册表中读取的调试标志

Public gstrMatchMethod As String    '匹配方式:0表示双向匹配
Public gstrDec As String

Public gintInsure As Integer
Public gstrInsure As String         '记录所有已使用的医保接口
Public gstr医院编码 As String * 10               '医院编号
Public gstr医保机构编码 As String
Public glngReturn As Long           '函数操作返回标志
Public gbln批量虚拟结算 As Boolean

Public mintOrder As Integer         '当前医保接口对象的序号
Public gclsInsure As clsInsure
Public gobjInsure_Obj() As Object   '保存所有已打开的医保部件对象
Public gobjInsure_Name() As String  '保存所有已打开的医保部件名称
Public glngInstanceCount As Long    '当前实例个数,94352

Public Type T结算数据
    病人ID       As Long
    年度         As Long
    住院次数     As Long
    帐户累计增加   As Currency
    帐户累计支出   As Currency
    累计进入统筹   As Currency
    累计统筹报销   As Currency
    起付线         As Currency
    封顶线         As Currency
    实际起付线     As Currency
    发生费用金额   As Currency
    全自费金额   As Currency
    首先自付金额   As Currency
    进入统筹金额   As Currency
    优惠金额       As Currency
    统筹报销金额   As Currency
    超限自付金额   As Currency
    个人帐户支付   As Currency
    支付顺序号     As String
    主页ID         As Long
    中途结帐       As Long
    住院床日       As Long

    基本统筹自付 As Currency
    
    '曾明春(20060711):以下内容为温江区银海医保增加
    公务员统筹支付 As Currency
    公务员报销床位费 As Currency
    公务员报销GGZF As Currency
    公务员报销起付线 As Currency
    公务员报销超限额 As Currency
    人员职称       As String
    
    补充医疗统筹支付 As Currency
    补充医疗报销起付线 As Currency
    补充医疗报销统筹自付 As Currency
    补充医疗报销超限 As Currency
End Type
Public g结算数据 As T结算数据           '保存预结算之后计算的结果，可用以填写保险结算记录
Public gcol结算计算 As New Collection   '保存预结算之后计算的结果，可用以填写保险结算计算
                                        '每个成员为一个数组，依次为档次、进入统筹金额、统筹报销金额、比例


Public Enum 交易Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum

Public Enum 余额Enum
    balan门诊 = 10
    balan入院 = 20
    balan预交 = 30
    balan结算 = 40
End Enum

Public Enum 身份验证Enum
    id门诊收费 = 0
    id入院登记 = 1
    id帐户管理 = 2
    id挂号 = 3
    id结帐 = 4
    id门诊确认 = 5
End Enum

Public Enum 医院业务
    'Modified by ZYB 2005-08-08 取消以下三个参数（费用程序不再使用，但仍保存的目的是为了医保编译需要），并增加两个参数：support门诊结算作废、support住院结算作废
    '原因：结算作废与原始结算结果不一致的问题
    '新的解决办法：使用GetCapability函数进行检查是否支持结算作废，如果strAdvance不为空，则表示检查某个特定的结算方式，该医保是否支持全退，如果不支持，则表示该结算方式全退为现金
    support门诊退费 = 1
    support结帐退个人帐户 = 3
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    
    support门诊预算 = 0
    
    support预交退个人帐户 = 2
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤销出院 = 17            '允许撤消病人出院
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
    support出院无实际交易 = 29      '出院接口中是否要与接口商进行交易

    '多单据收费时需要注意的问题
    '如果我方是多单据，而医保那边保存为一张单据的，需要进行如下考虑：
    '1、如果系统参数“78-多张单据收费分别打印”为真，不允许进行多单据收费
    '2.1、退费时，如果保险结算记录的备注中含“多单据收费”，或者该单据是多单据收费（票据打印内容大于1条记录）则继续执行
    '2.2、如果退费时该单据是单张收费（票据打印内容小于等于1条记录），则提取该病人该登记时间相同的所有单据号出来，提示操作员应该同时退费后再收取新的费用
    support多单据收费 = 30          '是否支持多单据收费
    
    support门诊收费存为划价单 = 31  '将门诊收费单转为划价单保存，修改以前固定判断某个医保的方式
    support允许部分冲销明细 = 32    '允许针对住院记帐处方的每笔明细进行部分冲销
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持当做普通病人作废处理
    support住院结算作废 = 34        'HIS始终认为住院支持结算作废，如果不支持需医保接口内部处理，返回假即可；增加该参数是为了配合GetCapability交易来检查各种结算方式是否支持全退
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    support结帐_指定住院次数 = 36   '是否支持指定住院次数进行医保结算
    support结帐_指定日期范围 = 37   '是否支持指定结帐日期范围进行医保结算
    support结帐_设置婴儿费条件 = 38 '是否允许设置婴儿费条件
    Support多单据收费必须全退 = 39  '多单据收费必须全退
    '只能各医保部件内的医保初始化函数中处理，如果要集成部件处理的话，又需要单独用个集合来保存医保化的返回值了
    Support初始化失败脱机处理 = 40  '当处方保存前的医保初始化失败，以后不再进行初始化，当做脱机处理（只保存HIS）
    support门诊结帐 = 41            '是否支持门诊医保病人的记帐费用使用住院结帐来完成
    support结帐_指定科室 = 42           '是否允许在结帐设置界面中指定科室
    support结帐_指定费用项目 = 43       '是否允许在结帐设置界面中指定费用项目
    support结帐_结帐设置后调用接口 = 44 '是否在结帐设置后才调用住院虚拟结算？
    support结帐_门诊结帐设置后调用接口 = 49 '门诊结帐:是否在结帐设置后才调用门诊虚拟结算？
    support结帐_指定费用类型 = 45       '是否允许在结帐设置界面中指定费用类型

        support医保接口打印票据 = 46                    'HIS仍然是严格控制票据但不打印，由医保完成票据的打印，单张收费一次只打印一张发票，多张收费根据系统设定决定打几张发票
        support多单据一次结算 = 47                              '如果医保支持这个参数，你将所有单据返回的报销总额，依次分摊到各单据上；结算作废时也是如此。
                                                                                '建议流程：虚拟结算时在最后一张单据时汇总上传，结算时在第一张单据时结算
        support医生确定处方类型 = 48                    '来源于北京医保，在门诊医嘱发送时弹出提示询问是医保内还是医保外处方，将结果保存在费用记录的摘要中

    support挂号必须传递明细 = 61
    support连续挂号 = 62
    support门诊改费 = 63
    
    
    support实时监控 = 60                '指明是否存在实时监控相关的接口函数：CheckClinicGuideline、CheckSettleGuideline、CheckItem
    '以下两参数用于控制是否进行特准项目限制
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support不提醒缴款金额不足 = 64            '在收费时,如果收费参数的"不进行缴款输入和累计控制"为true时,同时是医保病人时没有输入缴款金额时不提醒用户
        support门诊退费后打印回单 = 65                          '门诊退费后，由收费模块调用自定义报表完成回单打印
        support结帐作废后打印回单 = 66                          '住院结算作废后，由结帐模块调用自定义报表完成回单打印

    support上传门诊档案 = 70                    '在门诊医嘱发送时，是否调用TranElecDossier函数完成门诊病人电子卷宗/电子档案的上传
        
        support门诊_不分单据结算 = 80                                   '预结算、结算都只调用一次医保交易
    
    support挂号不收取病历费 = 81    '在挂号时，不使用医保收取病历费
    
    support按单据全退 = 82 '门诊退费时，按单据进行退费，86176
    support多单据分单据结算 = 83 '多单据一次结算按单据进行医保报销，86321
End Enum

Public gblnLED As Boolean '是否使用LED语音设备
Private rsInsure As New ADODB.Recordset

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

'================================================================================================================================
'=功能说明：提示错误信息，并且在存在事务的过程中加上事务回滚。
'=入参：
'=  1.strErrMsg：错误提示信息。
'=  2.mbs：错误提示模式，默认为VbMsgBoxStyle.vbInformation。
'=  3.strTitle：错误提示标题。
'=  4.blnTran：是否执行事务回滚。
'=出参：(无)
'=返回：(VbMsgBoxResult)提示后的选择值。
'=注意：
'=  1.在HIS中没有开启事务的地方：blnTran=False。在有事务开启的地方，blnTran=True。
'=  2.门诊挂号、门诊挂号冲销、门诊结算、门诊结算冲销中必须传入参数blnTran=True。
'=  3.入院登记、入院登记撤销、出院登记、出院登记撤销中可传入参数blnTran=True。
'=  4.在住院结算、住院结算冲销中必须传入参数blnTran=True。
'=  5.在门诊虚拟结算、处方上传、住院虚拟结算等方法中，无需传入blnTran参数或blnTran=
'================================================================================================================================
Public Function ErrMsgBox(strErrMsg As String, Optional mbsStyle As VbMsgBoxStyle = vbInformation, Optional strTitle As String = "") As VbMsgBoxResult
    Dim blnTran As Boolean
On Error GoTo ErrH
    '获取事务状态
    blnTran = gclsInsure.zlTranState
    '回滚事务
    If blnTran Then gcnOracle.RollbackTrans
    '提示错误消息
    ErrMsgBox = MsgBox(strErrMsg, mbsStyle, strTitle)
    '重新开启事务
    If blnTran Then gcnOracle.BeginTrans
    '输出到调试工具中去
    DebugTool strErrMsg
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Function
End Function

Public Function GetErrInfo(strCode As String, ByVal intinsure As Integer) As String
'功能：根据错误代码返回错误信息
'参数：bytType=保险类别,strCode=错误代码
    Dim rsTmpErr As New ADODB.Recordset
    
    strCode = Trim(strCode)
End Function

Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnMessage As Boolean = True) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
    If Err <> 0 Then
        If blnMessage = True Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            Else
                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
            End If
        End If
        
        Err.Clear
        OraDataOpen = False
        Exit Function
    End If
    OraDataOpen = True
End Function

Public Sub GetUserInfo()
 '功能：获取登陆用户信息
    Dim rsUser As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsUser = New ADODB.Recordset
    rsUser.CursorLocation = adUseClient
    'rsUser.Open "Select A.ID,A.部门ID,A.编号,A.简码,A.姓名,B.用户名,C.名称 as 部门 from 人员表 A,上机人员表 B,部门表 C Where A.部门ID=C.ID And  B.人员ID=A.ID AND Upper(B.用户名)=Upper(User)", gcnOracle, adOpenKeyset
    
    strSQL = "select P.*,D.编码 as 部门编码,D.名称 as 部门名称,M.部门ID,u.用户名 " & _
                " from 上机人员表 U,人员表 P,部门表 D,部门人员 M " & _
                " Where U.人员id = P.id And P.ID=M.人员ID and  M.缺省=1 and M.部门id = D.id and U.用户名=user"
    rsUser.Open strSQL, gcnOracle, adOpenKeyset
    
    If rsUser.RecordCount <> 0 Then
        UserInfo.ID = rsUser!ID
        UserInfo.编号 = rsUser!编号
        UserInfo.部门ID = IIf(IsNull(rsUser!部门ID), 0, rsUser!部门ID)
        UserInfo.简码 = IIf(IsNull(rsUser!简码), "", rsUser!简码)
        UserInfo.姓名 = IIf(IsNull(rsUser!姓名), "", rsUser!姓名)
        UserInfo.部门 = rsUser!部门名称
        UserInfo.用户名 = rsUser!用户名
        UserInfo.站点 = rsUser!用户名
        
        '为了不改其它程序，重复增加了一个变量
        gstrUserName = UserInfo.姓名
    End If
End Sub

Public Function DateStr() As String
    Dim rsTmp As New ADODB.Recordset

    rsTmp.Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    DateStr = Format(rsTmp.Fields(0).Value, "yyyy-MM-dd HH:mm:ss")
End Function

Public Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Public Function TruncZero(ByVal StrInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(StrInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(StrInput, 1, lngPos - 1)
    Else
        TruncZero = StrInput
    End If
End Function

Public Function NextNo(intBillID As Integer) As Variant
'功能：根据特定规则产生新的号码,规则如下：
'   一、项目序号：
'   1   病人ID         数字
'   2   住院号         数字(ZLHIS9/10规则不同，暂不支持)
'   3   门诊号         数字(ZLHIS9/10规则不同，暂不支持)
'   x   其它单据号     字符,根据编号规则顺序递增编号,不自动补缺
'   二、年度位确定原则:
'       以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码

    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim vntNo As Variant, strSQL As String
    Dim intYear, strYear As String
ReStart:
    Err = 0
    On Error GoTo errHand

    If intBillID = 1 Then '病人ID
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            strSQL = "Select Nvl(Max(病人ID),0)+1 as 病人ID From 病人信息 Where 病人ID>=" & vntNo
            
            With rsTmp
                If .State = adStateOpen Then .Close
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    Else
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            intYear = Format(!Today, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIf(IsNull(!最大号码), "", !最大号码)
            
            If IIf(IsNull(!编号规则), 0, !编号规则) = 1 Then
                '按日顺序编号
                If vntNo < strYear & Format(CDate(Format(!Today, "YYYY-MM-dd")) - CDate(Format(!Today, "YYYY") & "-01-01") + 1, "000") & "0000" Then
                    vntNo = strYear & Format(CDate(Format(!Today, "YYYY-MM-dd")) - CDate(Format(!Today, "YYYY") & "-01-01") + 1, "000") & "0000"
                End If
                vntNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + 1), 4)
            Else
                '按年顺序编号
                If Left(vntNo, 1) < strYear Then
                    vntNo = strYear & "0000000"
                End If
                vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
            End If
            
            If Not (UCase(strYear) >= "A" And UCase(strYear) <= "Z") Or zlCommFun.ActualLen(vntNo) > 8 Then GoTo ReStart
            
            On Error Resume Next
            .Update "最大号码", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    NextNo = Null
End Function

Public Function GetNextNO(ByVal intBillID As Integer, Optional lng科室 As Long) As Variant
    'blnUse:以下版本可使用公共部件中的GetNextNO函数，但必须是HIS+版本
    
    If IsZLHIS10 Then
        #If gverControl >= 3 Then
            GetNextNO = zlDatabase.GetNextNO(intBillID, lng科室)
        #Else
            GetNextNO = NextNo(intBillID)
        #End If
    Else
        GetNextNO = NextNo(intBillID)
    End If
End Function

 

Public Function Get入院诊断(lng病人ID As Long, lng主页ID As Long, _
Optional ByVal bln允许空 As Boolean = True, Optional ByVal bln疾病编码 As Boolean = False) As String
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.描述信息 as 入院诊断,B.编码 疾病编码 " & _
             " From 诊断情况 A,疾病编码目录 B " & _
             " Where A.病人ID=[1] And A.疾病ID=B.ID(+) And A.主页ID=[2] And A.诊断类型=2"
    Set rsInNote = zlDatabase.OpenSQLRecord(strTmp, "Get入院诊断", lng病人ID, lng主页ID)
    
    If Not rsInNote.EOF Then
        Get入院诊断 = IIf(IsNull(rsInNote!入院诊断), "", rsInNote!入院诊断)
    End If
    If Not bln允许空 Then
        Get入院诊断 = Trim(Get入院诊断)
        If Get入院诊断 = "" Then Get入院诊断 = "无"
    End If
    If bln疾病编码 Then
        If Not rsInNote.EOF Then
            Get入院诊断 = Get入院诊断 & "|" & Nvl(rsInNote!疾病编码)
        Else
            Get入院诊断 = Get入院诊断 & "|"
        End If
    End If
End Function

Public Function BuildPatiInfo(ByVal bytType As Byte, ByVal strInfo As String, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Long
'功能：建立病人帐户信息
'参数：bytType=0-门诊,1-住院
'      strInfo='0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
'      8中心;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(1,2,3);15退休证号;16年龄段;17灰度级
'      18帐户增加累计;19帐户支出累计;20进入统筹累计;21统筹报销累计;22住院次数累计;23就诊类别
'      24本次起付线;25起付线累计;26基本统筹限额
'返回：病人ID
    Const MAX_BOUND = 26 '要求传入的信息段数
    
    Dim rsPati As New ADODB.Recordset, str单位编码 As String, lng年龄 As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim lng中心 As Long, array信息 As Variant
    Dim lngTemp As Long
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    If Len(Trim(strInfo)) <> 0 Then
        curDate = zlDatabase.Currentdate
        
        '200308z012:保证传入的信息串够用
        If UBound(Split(strInfo, ";")) < MAX_BOUND Then
            strInfo = strInfo & String(MAX_BOUND - UBound(Split(strInfo, ";")), ";")
        End If
        array信息 = Split(strInfo, ";")
        
        '从第7项内容中取出单位编码
            If array信息(7) Like "*(*" Then
                str单位编码 = Split(array信息(7), "(")(UBound(Split(array信息(7), "(")))
                str单位编码 = Mid(str单位编码, 1, Len(str单位编码) - 1)
            End If
        
        '取年龄
        If IsDate(array信息(5)) Then
            lng年龄 = Int(curDate - CDate(array信息(5))) / 365
        End If
        
        lng中心 = Val(array信息(8))
        
        '提供了病人身份绑定的功能，因此不再需要合并
'        If lng病人ID > 0 Then
'            '该病人已经存在
'            gstrSQL = "Select nvl(病人ID,0) 病人ID from 保险帐户 where 医保号='" & CStr(array信息(1)) & "' and 中心=" & lng中心 & " and 险类=" & intInsure
'            Call OpenRecordset(rsTemp, "建立帐户")
'            If rsTemp.EOF = False Then
'                If rsTemp("病人ID") <> lng病人ID Then
'                    '曾明春(2006-01-16):以下医保支持补充登记时自动登记
'                    If intInsure = TYPE_成都市 Or intInsure = TYPE_新都 Or intInsure = type_成都郊县 Or intInsure = TYPE_乐山 Or intInsure = TYPE_广元旺苍 Or intInsure = TYPE_成都德阳 Or intInsure = TYPE_南充阆中 Then
'                        If MsgBox("已经存在相同医保号的另外一位病人，您需要将这两位病人合并吗？", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then
'                            gcnOracle.RollbackTrans
'                            Exit Function
'                        End If
'                        '对这两个病人进行合并
'                        lngTemp = MergePatient(lng病人ID, rsTemp!病人ID)
'                        If lngTemp = 0 Then
'                            gcnOracle.RollbackTrans
'                            Exit Function
'                        End If
'                        lng病人ID = lngTemp
'                    Else
'                        MsgBox "已经存在相同医保号的另外一位病人，请您在病人管理中将这两位病人合并", vbInformation, gstrSysName
'                        gcnOracle.RollbackTrans
'                        Exit Function
'                    End If
'                End If
'            End If
'        End If
        
        '帐户唯一：险类,中心,医保号
        #If gverControl < 6 Then
            strSQL = "Select A.*,B.医保号 From 病人信息 A," & _
                "   (Select * From 保险帐户" & _
                "   Where 险类=[1] And 医保号=[2] And Nvl(中心,0)=[3]) B" & _
                " Where " & IIf(lng病人ID = 0, "A.病人ID=B.病人ID", "A.病人ID=B.病人ID(+) and A.病人ID=[4]") '可能病人ID已经确定
        #Else
            strSQL = "Select A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.卡验证码, A.费别, A.医疗付款方式, A.姓名, A.性别, A.年龄, A.出生日期, A.出生地点, A.身份证号, A.其他证件, A.身份, A.职业, A.民族, A.国籍, A.区域, A.学历, A.婚姻状况, A.家庭地址," & vbNewLine & _
                "      A.家庭电话, A.家庭地址邮编 As 户口邮编, A.监护人, A.联系人姓名, A.联系人关系, A.联系人地址, A.联系人电话, A.合同单位id, A.工作单位, A.单位电话, A.单位邮编, A.单位开户行, A.单位帐号, A.担保人, A.担保额, A.担保性质, A.就诊时间, A.就诊状态," & vbNewLine & _
                "      A.就诊诊室, A.住院次数, A.当前科室id, A.当前病区id, A.当前床号, A.入院时间, A.出院时间, A.在院, A.Ic卡号, A.健康号, A.医保号, A.险类, A.查询密码, A.登记时间, A.停用时间, A.锁定," & vbNewLine & _
                "      B.医保号 From 病人信息 A," & _
                "   (Select * From 保险帐户" & _
                "   Where 险类=[1] And 医保号=[2] And Nvl(中心,0)=[3]) B" & _
                " Where " & IIf(lng病人ID = 0, "A.病人ID=B.病人ID", "A.病人ID=B.病人ID(+) and A.病人ID=[4]") '可能病人ID已经确定
        #End If
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "提取病人信息", intinsure, CStr(array信息(1)), lng中心, lng病人ID)
        If rsPati.EOF Then
            '无保险帐户则认为没有病人信息
            If lng病人ID = 0 Then lng病人ID = GetNextNO(1)
            strSQL = "zl_病人信息_Insert(" & lng病人ID & ",NULL,NULL,'社会基本医疗保险'," & _
                "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "NULL,'" & array信息(6) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,'" & array信息(7) & "',NULL,NULL,NULL," & _
                "NULL,NULL,NULL," & intinsure & "," & _
                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSQL, "产生病人档案")
        Else
            '检查该病人是否已经停用
            If Not IsNull(rsPati!停用时间) Then
                gcnOracle.RollbackTrans
                MsgBox "该病人的信息已经停用。", vbInformation, gstrSysName
                Exit Function
            End If
            
            '有病人信息和保险帐户信息
            If rsPati("姓名") <> array信息(3) Then
                If MsgBox("病人原有登记的姓名是 " & rsPati("姓名") & " ，与刷卡得到的姓名 " & array信息(3) & " 不符，" & vbCrLf & _
                          "继续会更新病人原有的登记信息，是否确定？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            End If
            
            '2005-08-13 周海全
            '处理工作单位，原处理方式将工作单位保留为ID
            If lng病人ID = 0 Then lng病人ID = rsPati!病人ID
                strSQL = "zl_病人信息_Update(" & _
                    lng病人ID & "," & IIf(IsNull(rsPati!门诊号), "NULL", rsPati!门诊号) & "," & _
                    IIf(IsNull(rsPati!住院号), "NULL", rsPati!住院号) & ",'" & IIf(IsNull(rsPati!费别), "", rsPati!费别) & "'," & _
                    "'" & IIf(IsNull(rsPati!医疗付款方式), "", rsPati!医疗付款方式) & "'," & _
                    "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                    "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                    "'" & IIf(IsNull(rsPati!出生地点), "", rsPati!出生地点) & "','" & array信息(6) & "'," & _
                    "'" & IIf(IsNull(rsPati!身份), "", rsPati!身份) & "','" & IIf(IsNull(rsPati!职业), "", rsPati!职业) & "'," & _
                    "'" & IIf(IsNull(rsPati!民族), "", rsPati!民族) & "','" & IIf(IsNull(rsPati!国籍), "", rsPati!国籍) & "'," & _
                    "'" & IIf(IsNull(rsPati!学历), "", rsPati!学历) & "','" & IIf(IsNull(rsPati!婚姻状况), "", rsPati!婚姻状况) & "'," & _
                    "'" & IIf(IsNull(rsPati!家庭地址), "", rsPati!家庭地址) & "','" & IIf(IsNull(rsPati!家庭电话), "", rsPati!家庭电话) & "'," & _
                    "'" & IIf(IsNull(rsPati!户口邮编), "", rsPati!户口邮编) & "','" & IIf(IsNull(rsPati!联系人姓名), "", rsPati!联系人姓名) & "'," & _
                    "'" & IIf(IsNull(rsPati!联系人关系), "", rsPati!联系人关系) & "','" & IIf(IsNull(rsPati!联系人地址), "", rsPati!联系人地址) & "'," & _
                    "'" & IIf(IsNull(rsPati!联系人电话), "", rsPati!联系人电话) & "'," & IIf(IsNull(rsPati!合同单位ID), "NULL", rsPati!合同单位ID) & "," & _
                    " " & IIf(IsNull(rsPati!工作单位), "NULL", "'" & rsPati!工作单位 & "'") & ",'" & IIf(IsNull(rsPati!单位电话), "", rsPati!单位电话) & "'," & _
                    "'" & IIf(IsNull(rsPati!单位邮编), "", rsPati!单位邮编) & "','" & IIf(IsNull(rsPati!单位开户行), "", rsPati!单位开户行) & "'," & _
                    "'" & IIf(IsNull(rsPati!单位帐号), "", rsPati!单位帐号) & "','" & IIf(IsNull(rsPati!担保人), "", rsPati!担保人) & "'," & _
                    " " & IIf(IsNull(rsPati!担保额), "NULL", rsPati!担保额) & "," & intinsure & ")"
                Call SQLTest(App.ProductName, "医保接口", strSQL)
            Call zlDatabase.ExecuteProcedure(strSQL, "更新病人档案")
        End If
        
        '插入或更新保险帐户信息(自动)
        strSQL = "zl_保险帐户_insert(" & lng病人ID & "," & intinsure & "," & _
            lng中心 & "," & _
            "'" & IIf(array信息(0) = "-1", array信息(1), array信息(0)) & "'," & _
            "'" & array信息(1) & "'," & _
            "'" & array信息(2) & "'," & _
            "'" & array信息(9) & "'," & _
            "'" & array信息(15) & "'," & _
            "'" & array信息(10) & "'," & _
            "'" & str单位编码 & "'," & _
            Val(array信息(11)) & "," & _
            Val(array信息(12)) & "," & _
            IIf(Val(array信息(13)) = 0, "NULL", Val(array信息(13))) & "," & _
            IIf(Val(array信息(14)) = 0, 1, Val(array信息(14))) & "," & _
            IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
            "'" & array信息(17) & "'," & _
            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, "产生保险帐户")
        
        '插入或更新帐户年度信息(自动)
        '200308z012:成都:保存"24本次起付线=zyjs,25起付线累计=tcbxbl,26基本统筹限额=zyxe"
        strSQL = "zl_帐户年度信息_Insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
            Val(array信息(18)) & "," & Val(array信息(19)) & "," & _
            Val(array信息(20)) & "," & Val(array信息(21)) & "," & _
            Val(array信息(22)) & "," & Val(array信息(24)) & "," & Val(array信息(25)) & "," & Val(array信息(26)) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "更新年度信息")
    End If
    
    gcnOracle.CommitTrans
    BuildPatiInfo = lng病人ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Public Function GetTextFromCombo(cmbTemp As ComboBox, ByVal blnAfter As Boolean, Optional strSplit As String = ".") As String
'参数：cmbTemp  准备获取数据的ComboBox控件
'      blnAfter 表示在.之前或之后取值
    Dim lngPos As Long
    
    lngPos = InStr(cmbTemp.Text, strSplit)
    If lngPos = 0 Then
        '直接返回整个字符串
        GetTextFromCombo = "'" & cmbTemp.Text & "'"
    Else
        If blnAfter = False Then
            '圆点之前
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, 1, lngPos - 1) & "'"
        Else
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, lngPos + 1) & "'"
        End If
    End If
End Function

Public Sub SetComboByText(cmbTemp As ComboBox, ByVal strText As String, ByVal blnAfter As Boolean, Optional strSplit As String = ".")
'参数：cmbTemp  准备设置的ComboBox控件
'      blnAfter 表示在.之前或之后取值
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cmbTemp.ListCount - 1
        strTemp = cmbTemp.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            '直接返回整个字符串
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                '圆点之前
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '已经找到
        cmbTemp.ListIndex = lngCount
    Else
        cmbTemp.ListIndex = -1
        If blnAfter = True Then
            cmbTemp.AddItem strText
        End If
    End If
End Sub

Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'功能：按数据库规则得到字符串的子集，也就是汉字按两个字符算，而字母仍是一个
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    '去掉可能出现的半个字符
    MidUni = Replace(MidUni, Chr(0), "")
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
'功能：将文本按Varchar2的长度计算方法进行截断
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    '去掉可能出现的半个字符
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Public Function GetComputer(frmParant As Form, Optional ByVal strCaption As String = "选择计算机") As String
'功能：返回计算机名
   Dim BI As BrowseInfo
   Dim pidl As Long
   Dim sPath As String
   Dim pos As Integer
   
  'obtain the pidl to the special folder 'network'
   If SHGetSpecialFolderLocation(frmParant.hwnd, CSIDL_NETWORK, pidl) = 0 Then
     'fill in the required members, limiting the
     'Browse to the network by specifying the
     'returned pidl as pidlRoot
      With BI
         .hwndOwner = frmParant.hwnd
         .pIDLRoot = pidl
         .pszDisplayName = Space$(MAX_PATH)
         .lpszTitle = lstrcat(strCaption, "")
         .ulFlags = BIF_BROWSEFORCOMPUTER
      End With
         
     'show the browse dialog. We don't need
     'a pidl, so it can be used in the If..then directly.
      If SHBrowseForFolder(BI) <> 0 Then
               
         'a server was selected. Although a valid pidl
         'is returned, SHGetPathFromIDList only return
         'paths to valid file system objects, of which
         'a networked machine is not. However, the
         'BROWSEINFO displayname member does contain
         'the selected item, which we return
          GetComputer = TrimStr(BI.pszDisplayName)
            
      End If  'If SHBrowseForFolder
      
      Call CoTaskMemFree(pidl)
               
   End If  'If SHGetSpecialFolderLocation
   
End Function

Public Sub CenterTableCaption(mshTemp As Object)
'功能：设置表格的列头居中对齐
    With mshTemp
        .COL = 0
        .Row = .FixedRows - 1
        .ColSel = .Cols - 1
        .RowSel = .Row
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = .FixedRows: .COL = .FixedCols
    End With
End Sub

Public Function Get住院次数(lng病人ID As Long) As Integer
'功能：获取指定病人本年度住院次数
'说明：跨年住院的情况两年都各算一次住院。
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Count(*) as 次数 From 病案主页" & _
        " Where Nvl(主页ID,0)<>0 And Nvl(出院日期,Sysdate)=To_Date(To_Char(Sysdate,'YYYY')||'-01-01','YYYY-MM-DD') And 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取指定病人本年度住院次数", lng病人ID)
    
    If Not rsTmp.EOF Then Get住院次数 = IIf(IsNull(rsTmp!次数), 0, rsTmp!次数)
End Function

Public Function Get帐户信息(ByVal intinsure As Integer, ByVal lng病人ID As Long, ByVal str年度 As String, int住院次数累计 As Integer, _
    cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, _
    cur统筹报销累计 As Currency, Optional cur本次起付线 As Currency, Optional cur起付线累计 As Currency, _
    Optional cur基本统筹限额 As Currency) As Boolean
'功能：得到帐户年度信息
'200308z012:新增几个返回参数
    Dim rsTemp As New ADODB.Recordset
    
    cur帐户增加累计 = 0
    cur帐户支出累计 = 0
    cur进入统筹累计 = 0
    cur统筹报销累计 = 0
    int住院次数累计 = 0
    cur本次起付线 = 0
    cur起付线累计 = 0
    cur基本统筹限额 = 0
    
    '帐户年度信息
    gstrSQL = "Select * From 帐户年度信息 Where 病人ID=[1] And 险类=[2] And 年度=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取帐户年度信息", lng病人ID, intinsure, str年度)
    If rsTemp.EOF = False Then
        cur帐户增加累计 = IIf(IsNull(rsTemp("帐户增加累计")), 0, rsTemp("帐户增加累计"))
        cur帐户支出累计 = IIf(IsNull(rsTemp("帐户支出累计")), 0, rsTemp("帐户支出累计"))
        cur进入统筹累计 = IIf(IsNull(rsTemp("进入统筹累计")), 0, rsTemp("进入统筹累计"))
        cur统筹报销累计 = IIf(IsNull(rsTemp("统筹报销累计")), 0, rsTemp("统筹报销累计"))
        int住院次数累计 = IIf(IsNull(rsTemp("住院次数累计")), 0, rsTemp("住院次数累计"))
        cur本次起付线 = IIf(IsNull(rsTemp("本次起付线")), 0, rsTemp("本次起付线"))
        cur起付线累计 = IIf(IsNull(rsTemp("起付线累计")), 0, rsTemp("起付线累计"))
        cur基本统筹限额 = IIf(IsNull(rsTemp("基本统筹限额")), 0, rsTemp("基本统筹限额"))
    End If

End Function
Public Function Get封销信息(ByVal intinsure As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str年度 As String, Optional str封销信息 As String) As Boolean
'功能：检查特殊病人标志,成都蒲江地区使用
    Dim rsTemp As New ADODB.Recordset
    Dim str入院年度 As String
    
    str封销信息 = "0"
    '对于跨年度结算的病人必须取入院日期所在年度的信息
    gstrSQL = "Select to_char(入院日期,'YYYY') as 入院年度 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "对于跨年度结算的病人必须取入院日期所在年度的信息", lng病人ID, lng主页ID)
    str入院年度 = rsTemp("入院年度")
    If str入院年度 <> str年度 Then str年度 = str入院年度
    
    '帐户年度信息
    gstrSQL = "Select * From 帐户年度信息 Where 病人ID=[1] And 险类=[2] And 年度=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取帐户年度信息", lng病人ID, intinsure, str年度)
    If rsTemp.EOF = False Then
        str封销信息 = IIf(IsNull(rsTemp("封销信息")), "0", rsTemp("封销信息"))
    End If

End Function

Public Function 门诊虚拟结算(rs明细 As ADODB.Recordset, str结算方式 As String, ByVal intinsure As Integer) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim rs算法 As New ADODB.Recordset
    Dim cls医保 As New clsInsure
    Dim rs大类汇总 As New ADODB.Recordset
    Dim dbl全自费 As Currency, dbl首先自付 As Currency, dbl进入统筹 As Currency, dblTemp As Double
    Dim dbl最大金额 As Double
    Dim dbl个人帐户 As Double
    Dim lng病人ID As Long
    Dim rs特准项目 As New ADODB.Recordset
    Dim dblTemp1 As Double
    

    If rs明细.RecordCount > 0 Then
        rs明细.MoveFirst
        lng病人ID = rs明细("病人ID")
    End If
    
    gstrSQL = "select A.收费细目ID from 保险特准项目 A,保险帐户 B " & _
            "where A.病种ID=B.病种ID and B.病人ID=[1] and 险类=[2]"
    Set rs特准项目 = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", lng病人ID, intinsure)
    
    '2、按统筹支付项目合计发生金额和数量
    '2.1、初始化记录集
    Set rs大类汇总 = New ADODB.Recordset
    With rs大类汇总
        If .State = adStateOpen Then .Close
        .Fields.Append "保险大类ID", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adDouble, 8, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .Fields.Append "统筹金额", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    Err = 0
    On Error Resume Next
    '分类求出各大类的汇总额
    Do Until rs明细.EOF
        rs特准项目.Filter = "收费细目ID = " & rs明细("收费细目ID")
    
        If rs明细("是否医保") = 1 Or rs特准项目.EOF = False Then
            '如果是特准项目，强行进入统筹
            If rs大类汇总.RecordCount = 0 Then
                rs大类汇总.AddNew
                rs大类汇总("保险大类ID") = rs明细("保险支付大类ID")
                rs大类汇总("数量") = rs明细("数量")
                rs大类汇总("金额") = rs明细("实收金额")
            Else
                rs大类汇总.MoveFirst
                rs大类汇总.Find "保险大类ID=" & rs明细("保险支付大类ID")
                If rs大类汇总.EOF Then
                    rs大类汇总.AddNew
                    rs大类汇总("保险大类ID") = rs明细("保险支付大类ID")
                    rs大类汇总("数量") = rs明细("数量")
                    rs大类汇总("金额") = rs明细("实收金额")
                Else
                    rs大类汇总("数量") = rs大类汇总("数量") + rs明细("数量")
                    rs大类汇总("金额") = rs大类汇总("金额") + rs明细("实收金额")
                End If
            End If
            rs大类汇总.Update
        Else
            dbl全自费 = dbl全自费 + rs明细("实收金额")
        End If
        dblTemp = dblTemp + rs明细("实收金额")
        rs明细.MoveNext
    Loop
    g结算数据.发生费用金额 = dblTemp
    
    '2.2、计算进入统筹金额
    gstrSQL = "select ID,算法,统筹比额,特准定额,特准天数,是否医保 FROM 保险支付大类 where 险类=[1]"
    Set rs算法 = zlDatabase.OpenSQLRecord(gstrSQL, "计算进入统筹金额", intinsure)
    
    dblTemp = 0
    If rs大类汇总.RecordCount > 0 Then rs大类汇总.MoveFirst
    g结算数据.优惠金额 = 0
    Do Until rs大类汇总.EOF
        rs算法.Filter = "ID=" & rs大类汇总("保险大类ID")
        If rs算法.RecordCount > 0 Then
            If rs算法("是否医保") = 1 Then
                '算法:1-总额计算项目；2-住院日核定项目;3-费用档次计算法
                Select Case rs算法("算法")
                Case 1          '1-总额计算项目
                    If rs算法("统筹比额") = 0 Then
                        dbl全自费 = dbl全自费 + rs大类汇总("金额")
                    Else
                        dblTemp = dblTemp + rs大类汇总("金额") * rs算法("统筹比额") / 100
                    End If
                Case 2      '2-住院日核定项目
                    If Val(rs大类汇总("数量")) > Val(rs算法("特准天数")) Then
                        '如果住院日超过特准天数，那么最大金额就是 特准天数*特准定额 +  (数量-特准天数)*统筹比额
                        '当特准定额或特准天数任一个为0时，就相当于不要特准天数
                        dbl最大金额 = rs算法("特准定额") * rs算法("特准天数") + _
                            (rs大类汇总("数量") - IIf(rs算法("特准定额") = 0 Or rs算法("特准天数") = 0, 0, rs算法("特准天数"))) * rs算法("统筹比额")
                    Else
                        '如果住院日低于特准天数，那么最大金额就是 数量*特准定额 或者 数量*统筹比额
                        '当特准定额或特准天数任一个为0时，就相当于不要特准定额
                        If rs算法("特准定额") = 0 Or rs算法("特准天数") = 0 Then
                            dbl最大金额 = rs大类汇总("数量") * rs算法("统筹比额")
                        Else
                            dbl最大金额 = rs大类汇总("数量") * rs算法("特准定额")
                        End If
                    End If
                    
                    '总金额比最大金额小，就取全部金额；否则只最大金额
                    dblTemp = dblTemp + IIf(rs大类汇总("金额") < dbl最大金额, rs大类汇总("金额"), dbl最大金额)
                    
                    If rs大类汇总("金额") > dbl最大金额 Then
                        '全部算作全自费
                        dbl全自费 = dbl全自费 + rs大类汇总("金额") - dbl最大金额
                    End If
                Case Else   '3-费用档次计算法
                    If Nvl(rs大类汇总!金额, 0) = 0 Then
                    Else
                        dblTemp1 = 获取费用档次额_中联(Nvl(rs大类汇总!保险大类id, 0), Nvl(rs大类汇总!金额, 0))
                        dblTemp = dblTemp + dblTemp1
                        g结算数据.优惠金额 = g结算数据.优惠金额 + (Nvl(rs大类汇总!金额, 0) - dblTemp1)
                    End If
                End Select
            Else
                dbl全自费 = dbl全自费 + rs大类汇总("金额")
            End If
        Else
            dbl全自费 = dbl全自费 + rs大类汇总("金额")
        End If
        rs大类汇总.MoveNext
    Loop
    
    g结算数据.进入统筹金额 = dblTemp
    g结算数据.全自费金额 = dbl全自费
    g结算数据.首先自付金额 = g结算数据.发生费用金额 - dbl全自费 - dblTemp - g结算数据.优惠金额
   '20040617刘兴宏屏蔽
    '
    '
    '    Do Until rs明细.EOF
    '        rs特准项目.Filter = "收费细目ID = " & rs明细("收费细目ID")
    '
    '        If rs明细("是否医保") = 1 Or rs特准项目.EOF = False Then
    '            '如果是特准项目，强行进入统筹
    '            dbl进入统筹 = dbl进入统筹 + rs明细("统筹金额")
    '            dbl首先自付 = dbl首先自付 + rs明细("实收金额") - rs明细("统筹金额")
    '        Else
    '            dbl全自费 = dbl全自费 + rs明细("实收金额")
    '        End If
    '
    '        rs明细.MoveNext
    '    Loop
    
    If cls医保.GetCapability(support收费帐户全自费, 0, intinsure) = True Then
        dbl个人帐户 = dbl个人帐户 + dbl全自费
    End If
    
    If Is全额统筹(lng病人ID, intinsure) = True Then
        '首先自付也是由医保基金支付
        If g结算数据.优惠金额 = 0 Then
            str结算方式 = "个人帐户;" & dbl个人帐户 & ";0|医保基金;" & g结算数据.进入统筹金额 + g结算数据.首先自付金额 & ";0"
        Else
            str结算方式 = "个人帐户;" & dbl个人帐户 & ";0|医保基金;" & g结算数据.进入统筹金额 + g结算数据.首先自付金额 & ";0|优惠金额;" & g结算数据.优惠金额 & ";0"
        End If
        g结算数据.统筹报销金额 = g结算数据.进入统筹金额 + g结算数据.首先自付金额
    Else
        If cls医保.GetCapability(support收费帐户首先自付, 0, intinsure) = True Then
            dbl个人帐户 = dbl个人帐户 + g结算数据.首先自付金额
        End If
        If g结算数据.优惠金额 = 0 Then
            str结算方式 = "个人帐户;" & dbl个人帐户 & ";0|医保基金;" & g结算数据.进入统筹金额 & ";0"
        Else
            str结算方式 = "个人帐户;" & dbl个人帐户 & ";0|医保基金;" & g结算数据.进入统筹金额 & ";0|优惠金额;" & g结算数据.优惠金额 & ";0"
        End If
        g结算数据.统筹报销金额 = g结算数据.进入统筹金额
    End If
    门诊虚拟结算 = True
End Function

Public Function Is全额统筹(ByVal 病人ID As Long, ByVal intinsure As Integer) As Boolean
'功能：判断是否全额统筹病人(注意：传的病人ID可能非医保病人的)
    Dim rsTemp As New ADODB.Recordset
    
        gstrSQL = _
            "Select Nvl(B.全额统筹,0) as 全额统筹" & _
            " From 保险帐户 A,保险年龄段 B" & _
            " Where A.险类 = B.险类 And Nvl(A.中心, 0) = Nvl(B.中心, 0)" & _
            " And Nvl(A.在职,0)=Nvl(B.在职,0)" & _
            " And B.下限<=Nvl(A.年龄段,0) And (A.年龄段<=B.上限 Or B.上限=0)" & _
            " And A.病人ID=[1] And A.险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", 病人ID, intinsure)
        If Not rsTemp.EOF Then Is全额统筹 = (rsTemp!全额统筹 = 1)
End Function

Public Function AddDate(ByVal strOrin As String, Optional ByVal bln时 As Boolean = False) As String
'功能：为不全的日期信息补充完整
    Dim strTemp As String
    Dim intPos As Integer
    
    strTemp = Trim(strOrin)
    
    If strTemp = "" Then
        AddDate = ""
        Exit Function
    End If
    
    intPos = InStr(strTemp, "-")
    If intPos = 0 Then
        intPos = InStr(strTemp, ".")
        If intPos <> 0 Then
            '使用 . 隔
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        '没有"-",手工加上
        intPos = Len(strTemp)
        If intPos <= 8 Then
            If intPos = 8 Then
                strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            ElseIf intPos > 4 Then
                strTemp = Left(strTemp, intPos - 4) & "-" & Mid(Right(strTemp, 4), 1, 2) & "-" & Right(strTemp, 2)
            ElseIf intPos > 2 Then
                strTemp = Format(Date, "yyyy") & "-" & Left(strTemp, intPos - 2) & "-" & Right(strTemp, 2)
            Else
                strTemp = Format(Date, "yyyy") & "-" & Format(Date, "MM") & "-" & strTemp
            End If
        End If
    Else
        If bln时 = False Then
            If IsDate(strTemp) Then
                strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
            End If
        Else
            '处理小时
            If InStr(strTemp, " ") > 0 Then
                '输入了小时
                If IsDate(strTemp & ":00") Then
                    strTemp = Format(CDate(strTemp & ":00"), "yyyy-MM-dd HH:ss")
                End If
            Else
                If IsDate(strTemp) Then
                    strTemp = Format(CDate(strTemp), "yyyy-MM-dd HH:ss")
                End If
            End If
        End If
    End If
    
    AddDate = strTemp
End Function

Public Function Insert虚拟结算数据(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str结算方式 As String) As Boolean
'功能：将虚拟结算的数据保存起来
'参数：结算方式  "报销方式;金额;是否允许修改|...."
    Dim cnTemp As New ADODB.Connection
    Dim strDate As String
    Dim lngCount As Long, arr结算方式 As Variant, arr金额 As Variant
    
    cnTemp.Open gcnOracle.ConnectionString '为了防止一个连接串多次进放事务
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    cnTemp.BeginTrans
    On Error GoTo errHandle
    
    gstrSQL = "zl_保险模拟结算_Clear(" & lng病人ID & "," & lng主页ID & ")"
    cnTemp.Execute gstrSQL, , adCmdStoredProc
    
    arr结算方式 = Split(str结算方式, "|")
    For lngCount = 0 To UBound(arr结算方式)
        If arr结算方式(lngCount) <> "" Then
            arr金额 = Split(arr结算方式(lngCount), ";")
            If UBound(arr金额) > 1 Then
                If Val(arr金额(1)) <> 0 Then
                    gstrSQL = "zl_保险模拟结算_Insert(" & lng病人ID & "," & IIf(lng主页ID = 0, "null", lng主页ID) & _
                        ",'" & arr金额(0) & "'," & Val(arr金额(1)) & "," & strDate & ")"
                    cnTemp.Execute gstrSQL, , adCmdStoredProc
                End If
            End If
        End If
    Next
    
    cnTemp.CommitTrans
    Insert虚拟结算数据 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    cnTemp.RollbackTrans
End Function

Public Function Clear虚拟结算数据(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：在结帐之后，将虚拟结算的数据清除
    
    gstrSQL = "zl_保险模拟结算_Clear(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "虚拟结算")
    
    Clear虚拟结算数据 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get出生日期(ByVal str身份证 As String, ByVal lng年龄 As Long) As String
'功能：根据身份证号码或年龄得到出生日期
    Dim strDate As String
    If Len(str身份证) = 15 Then
        '老式的身份证号
        strDate = AddDate(Mid(str身份证, 7, 6))
        strDate = "19" & strDate
    ElseIf Len(str身份证) = 18 Then
        '新式的身份证号
        strDate = AddDate(Mid(str身份证, 7, 8))
    Else
        '没有身份证号
        strDate = Format(DateAdd("yyyy", lng年龄 * -1, Date), "yyyy-MM-dd")
    End If
    
    If IsDate(strDate) = True Then
        Get出生日期 = Format(CDate(strDate), "yyyy-MM-dd")
    End If
End Function

Public Function GetOracleFormat(ByVal dat日期 As Date)
    GetOracleFormat = "To_Date('" & Format(dat日期, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Sub RemoveSelect(lvw As ListView)
'功能：删除当前选中项
    Dim lngIndex  As Long
    
    With lvw
        If .SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = .SelectedItem.Index
        .ListItems.Remove lngIndex
        
        If .ListItems.Count > 0 Then
            '如果仍有列表，则进行下一个选择
            lngIndex = IIf(.ListItems.Count > lngIndex, lngIndex, .ListItems.Count)
            .ListItems(lngIndex).Selected = True
            .ListItems(lngIndex).EnsureVisible
        End If
    End With

End Sub

Public Function Can住院结算冲销(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：判断病人的住院结算数据是否允许作废。判断标准是检查病人有新的住院记录，如果有，就不能交冲销
'参数：lng病人ID     病人ID
'      lng主页ID     该结帐记录所在的住院次数
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle

    gstrSQL = "SELECT COUNT(*) as 住院次数 FROM 病案主页 WHERE 病人ID=[1] AND 主页ID>[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断病人的住院结算数据是否允许作废", lng病人ID, lng主页ID)
    If rsTemp("住院次数") > 0 Then
        MsgBox "该病人已经有新的住院记录，不能作废以前住院的结帐数据。", vbInformation, gstrSysName
        Exit Function
    End If

    Can住院结算冲销 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 医保病人已经出院(ByVal lng病人ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = " Select DECODE(出院日期,NULL,0,1) AS 出院状态 From 病案主页 " & _
              " Where (病人ID,主页ID) IN (Select 病人ID,住院次数 From 病人信息 Where 病人ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "判断医保病人是否出院", lng病人ID)
    医保病人已经出院 = (rsTmp!出院状态 = 1)
End Function

Public Function 存在未结费用(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim rs费用 As New ADODB.Recordset
    '检查该次住院是否还有费用未结算
    #If gverControl >= 5 Then
        gstrSQL = "Select nvl(费用余额,0) as 金额  from 病人余额 where 病人ID=[1] and 性质=1 And 类型=2"
    #Else
        gstrSQL = "Select nvl(费用余额,0) as 金额  from 病人余额 where 病人ID=[1] and 性质=1"
    #End If
    Set rs费用 = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在未结费用", lng病人ID)
    If rs费用.EOF = True Then
        存在未结费用 = False
    Else
        存在未结费用 = (rs费用("金额") <> 0)
    End If
End Function

Public Function 获取入出院诊断(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
Optional ByVal bln入院诊断 As Boolean = True, Optional ByVal bln允许空 As Boolean = True, _
Optional ByVal bln疾病编码 As Boolean = False) As String
    
    '1-门诊诊断;2-入院诊断;3-出院诊断
    Dim rs诊断 As New ADODB.Recordset
    If bln疾病编码 = False Then
        gstrSQL = " Select A.描述信息" & _
                  " From 诊断情况 A" & _
                  " Where A.病人ID=[1] And A.主页ID=[2]" & _
                  " And A.诊断类型=[3] And 诊断次序=1"
    Else
        gstrSQL = " Select A.描述信息,B.编码 疾病编码" & _
                  " From 诊断情况 A,疾病编码目录 B" & _
                  " Where A.病人ID=[1] And A.主页ID=[2]" & _
                  " And A.疾病ID=B.ID(+) And A.诊断类型=[3]"
    End If
    Set rs诊断 = zlDatabase.OpenSQLRecord(gstrSQL, "获取入出院诊断", lng病人ID, lng主页ID, IIf(bln入院诊断, "1", "3"))
    
    获取入出院诊断 = ""
    If Not rs诊断.EOF Then
        获取入出院诊断 = IIf(IsNull(rs诊断!描述信息), "", rs诊断!描述信息)
    End If
    
    获取入出院诊断 = Trim(获取入出院诊断)
    If Not bln允许空 And 获取入出院诊断 = "" Then
        获取入出院诊断 = "无"
    End If
    If bln疾病编码 Then
        If Not rs诊断.EOF Then
            获取入出院诊断 = 获取入出院诊断 & "|" & Nvl(rs诊断!疾病编码, " ")
        Else
            获取入出院诊断 = 获取入出院诊断 & "| "
        End If
    End If
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDO As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDO = 1 To 12
        strSource = Mid(strOld, intDO, 1)
        strTarget = Mid(strPass, intDO, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function 存在中心(ByVal int险类 As Integer) As Boolean
    Dim rs中心 As New ADODB.Recordset
    
    存在中心 = False
    gstrSQL = "Select Nvl(具有中心,0) 中心 From 保险类别 Where 序号=[1]"
    Set rs中心 = zlDatabase.OpenSQLRecord(gstrSQL, "是否有中心", int险类)
    If Not rs中心.EOF Then
        存在中心 = (rs中心!中心 = 1)
    End If
End Function

Private Function GetPatiInfo(lngID As Long) As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrH
    
'    strSql = "Select * From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID(+) And A.病人ID=" & lngID & " Order by 主页ID"
    '主页ID=0时(不是NULL)，表示预约入院
    strSQL = _
        " Select A.病人ID,Decode(B.病人ID,NULL,NULL,Nvl(B.主页ID,0)) as 主页ID," & _
        " A.姓名,A.住院号,B.入院日期,B.出院日期" & _
        " From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID(+) And A.病人ID=[1]" & _
        " Order by Nvl(B.主页ID,0)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病人的住院信息", lngID)
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
        
Private Function MergePatient(ByVal lngOld As Long, ByVal lngInsure As Long) As Long
    Dim i As Integer, j As Integer
    Dim curDate As Date, strSQL As String
    Dim rsPatiS As New ADODB.Recordset
    Dim rsPatiO As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    Set rsPatiS = GetPatiInfo(lngOld)
    Set rsPatiO = GetPatiInfo(lngInsure)
        
    'A或B有一个办理了预约入院
    If Not IsNull(rsPatiS!主页ID) And Nvl(rsPatiS!主页ID, 0) = 0 Then
        MsgBox "病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]办理了预约入院登记，请先取消该登记。", vbInformation, gstrSysName
    End If
    If Not IsNull(rsPatiO!主页ID) And Nvl(rsPatiO!主页ID, 0) = 0 Then
        MsgBox "病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]办理了预约入院登记，请先取消该登记。", vbInformation, gstrSysName
    End If
        
    'AB都住过院
    If Not IsNull(rsPatiS!主页ID) And Not IsNull(rsPatiO!主页ID) Then
        '1.先住院的在院,不允许(先后住院可以为：出院-出院,出院-在院；不允许：在院-出院,在院-在院)
        '因为除病人合并外,程序不额外处理自动出院或撤消出院
        rsPatiS.MoveLast
        rsPatiO.MoveLast
        If rsPatiS!入院日期 <= rsPatiO!入院日期 Then
            If IsNull(rsPatiS!出院日期) Then
                MsgBox "病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If IsNull(rsPatiO!出院日期) Then
                MsgBox "病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '2.时间交叉提示是否继续
        curDate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!入院日期 >= IIf(IsNull(rsPatiS!出院日期), curDate, rsPatiS!出院日期) Or _
                    IIf(IsNull(rsPatiO!出院日期), curDate, rsPatiO!出院日期) <= rsPatiS!入院日期) Then
                    If MsgBox("发现病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]第 " & rsPatiS!主页ID & " 次住院的期间" & Format(rsPatiS!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiS!出院日期), curDate, rsPatiS!出院日期), "yyyy-MM-dd") & vbCrLf & _
                        "与病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]的第 " & rsPatiO!主页ID & " 次住院的期间" & Format(rsPatiO!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiO!出院日期), curDate, rsPatiO!出院日期), "yyyy-MM-dd") & _
                        vbCrLf & "互相交叉，应该不是同一个病人，确实要合并吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
                rsPatiO.MoveNext
            Next
            rsPatiS.MoveNext
        Next
    End If
    
    '$IF HIS9
    #If gverControl = 0 Then
        strSQL = "zl_病人信息_MERGE(" & lngOld & "," & lngInsure & ")"
    #Else
    '$ELSE  HIS+
        strSQL = "zl_病人信息_MERGE(" & lngOld & "," & lngInsure & ", '医保补充登记合并','" & gstrUserName & "')"
    #End If
    
    DoEvents
    Screen.MousePointer = 11
    Call zlDatabase.ExecuteProcedure(strSQL, "病人身份合并")
    Screen.MousePointer = 0
    
    '合并后应只剩一个病人
    strSQL = "Select 病人ID From 病人信息 Where 病人ID IN([1],[2])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "合并后应只剩一个病人", lngOld, lngInsure)
    If Not rsTmp.EOF Then
        If glngSys Like "8??" Then
            MsgBox "客户合并成功,合并后的客户ID为 " & rsTmp!病人ID & "。", vbInformation, gstrSysName
        Else
            MsgBox "病人合并成功,合并后的病人ID为 " & rsTmp!病人ID & "。", vbInformation, gstrSysName
        End If
        MergePatient = rsTmp!病人ID
    End If
End Function

Public Sub DebugTool(ByVal strInfo As String)
    '如果调试=1，表示提试调试信息,2-将调式信息写入文本；其它情况不输出调试信息
    '判断是否是调试状态，是则显示提示框
    If gintDebug = -1 Then gintDebug = Val(GetSetting("ZLSOFT", "医保", "调试", 0))
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If gintDebug <> 1 Then
        If gintDebug = 2 Then
            '写文本文件
            '将调试信息写入文件中
            Dim objFile As New FileSystemObject
            Dim objText As TextStream
            Dim strFile As String
            
            Dim rsTemp As New ADODB.Recordset
            strFile = App.Path & "\调试信息.Log"
            If Not Dir(strFile) <> "" Then
                objFile.CreateTextFile strFile
            End If
            Set objText = objFile.OpenTextFile(strFile, ForAppending)
            objText.WriteLine strInfo
            objText.Close
        End If
        Exit Sub
    End If
    MsgBox strInfo
End Sub

Public Function SystemImes() As Variant
'功能：将系统中文输入法名称返回到一个字符串数组中
'返回：如果不存在中文输入法,则返回空串
    Dim arrIme(99) As Long, ARRNAME() As String
    Dim lngLen As Long, STRNAME As String * 255
    Dim lngCount As Long, i As Integer, j As Integer
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then
            ReDim Preserve ARRNAME(j)
            lngLen = ImmGetDescription(arrIme(i), STRNAME, Len(STRNAME))
            ARRNAME(j) = Mid(STRNAME, 1, InStr(STRNAME, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, ARRNAME, vbNullString)
End Function

Public Function OpenIme(Optional strIme As String) As Boolean
'功能:按名称打开中文输入法,不指定名称时关闭中文输入法。支持部分名称。
    Dim arrIme(99) As Long, lngCount As Long, STRNAME As String * 255
    
    If strIme = "不自动开启" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), STRNAME, Len(STRNAME)
            If InStr(1, Mid(STRNAME, 1, InStr(1, STRNAME, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
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

Public Function 身份证号转出生日期(ByVal str身份证号 As String, ByRef str出生日期) As Boolean
    
    Dim intI As Integer
    身份证号转出生日期 = True
    '验证传入的参数是否符合要求
    For intI = 1 To Len(str身份证号)
        If InStr("0123456789", Mid(str身份证号, intI, 1)) <= 0 Then
            If intI = 18 Then
                If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(str身份证号, intI, 1)) <= 0 Then
                    str出生日期 = "身份证号码中包含无效字符!"
                    身份证号转出生日期 = False
                End If
            Else
                str出生日期 = "身份证号码中包含无效字符!"
                身份证号转出生日期 = False
            End If
        End If
    Next
    
    If 身份证号转出生日期 = True Then
        Select Case Len(str身份证号)
            Case 15
                str出生日期 = "19" & Mid(str身份证号, 7, 6)
                If IsDate(Mid(str出生日期, 1, 4) & "-" & Mid(str出生日期, 5, 2) & "-" & Mid(str出生日期, 7, 2)) = False Then
                    str出生日期 = "身份证号码有错误!"
                    身份证号转出生日期 = False
                End If
            Case 18
                str出生日期 = Mid(str身份证号, 7, 8)
                If IsDate(Mid(str出生日期, 1, 4) & "-" & Mid(str出生日期, 5, 2) & "-" & Mid(str出生日期, 7, 2)) = False Then
                    str出生日期 = "身份证号码有错误!"
                    身份证号转出生日期 = False
                End If
            Case Else
                str出生日期 = "身份证号码位数错误!"
                身份证号转出生日期 = False
        End Select
    End If
    
End Function

Public Function IsApartComponents(ByVal intinsure As Integer) As Boolean
    On Error GoTo errHand
    '为了避免老部件的不断增加，而用户又不可能在短期内使用医保新的管理模式，因此增加此环节，如果是分离的部件，则单独处理
    '检查该部件是否是分离的部件
    If rsInsure.State = 0 Then
        gstrSQL = "Select 序号,医保部件,医保包 From 保险类别 "
        Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保部件名称")
    End If
    rsInsure.Filter = "序号=" & intinsure
    If rsInsure.RecordCount = 0 Then rsInsure.Filter = 0: Exit Function
    If Nvl(rsInsure!医保部件) = "" Then rsInsure.Filter = 0: Exit Function
    rsInsure.Filter = 0
    
    IsApartComponents = True
errHand:
End Function

Public Function CreateObject_Insure(ByVal intinsure As Integer, ByRef intOrder As Integer, Optional ByVal intCall As Integer = 0) As Boolean
    Dim blnExist As Boolean
    Dim strObject As String, strBag As String
    Dim intObject As Integer, intCOUNT As Integer
    Dim objTemp As Object
    '参数说明:
    'intCall:0-未指定医保包则创建医保部件;1-除Identify外，其它业务强制调用各自的医保部件
    
    On Error GoTo errHand
    '创建医保接口对象
    If rsInsure.State = 0 Then
        gstrSQL = " Select 序号,医保部件,医保包 From 保险类别 "
        Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保部件名称")
    End If
    rsInsure.Filter = "序号=" & intinsure
    If rsInsure.RecordCount = 0 Then
        rsInsure.Filter = 0
        MsgBox "该医保接口还未注册！序号=" & intinsure, vbInformation, gstrSysName
        Exit Function
    End If
    strBag = Nvl(UCase(rsInsure!医保包))
    strObject = UCase(Nvl(rsInsure!医保包, rsInsure!医保部件))
    If intCall = 1 Then strObject = UCase(rsInsure!医保部件)
    rsInsure.Filter = 0
    
    '检查是否存在该对象
    On Error Resume Next
    intCOUNT = UBound(gobjInsure_Name)
    If Err <> 0 Then intCOUNT = -1
    
    '应张永康要求屏蔽，因为使用了新的控件所致 2008-10-17
    'On Error GoTo errHand
    For intObject = 0 To intCOUNT
        If gobjInsure_Name(intObject) = strObject Then
            If Not gobjInsure_Obj(intObject) Is Nothing Then
                intOrder = intObject
                CreateObject_Insure = True
                Exit Function
            Else
                blnExist = True
                Exit For
            End If
        End If
    Next
    
    '去掉文件名后缀
    strObject = Mid(strObject, 1, Len(strObject) - 4)
    '创建对象
    Set objTemp = CreateObject(strObject & ".Cls" & Mid(strObject, 4))
    If objTemp Is Nothing Then Exit Function
    intObject = intCOUNT + 1
    ReDim Preserve gobjInsure_Name(intObject)
    ReDim Preserve gobjInsure_Obj(intObject)
    gobjInsure_Name(intObject) = strObject & ".DLL"
    Set gobjInsure_Obj(intObject) = objTemp
    intOrder = intObject
    
    '医保包的组件需要调用初始化函数
    If strBag <> "" Then
        If Not gobjInsure_Obj(intObject).InitInsure(gcnOracle, intinsure) Then Exit Function
    End If
    
    CreateObject_Insure = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ChooseInsure_Base(ByVal intinsure As Integer) As Integer
    '格式如下：
    '其它医保接口
    '31-沈阳铁路局医保
    '51-贵阳市医保
    '功能：如果选择“其它医保接口”，则走原来的流程；否则创建指定的分离医保部件，并调用其CodeMan()
    Dim intSelect As Integer
    Dim rsTemp As New ADODB.Recordset
    
    ChooseInsure_Base = intinsure
    '检查是否存在独立部件方式实现的医保接口
    gstrSQL = "Select count(*) AS Records From 保险类别 Where 医保部件 Is Not NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在独立部件")
    If Nvl(rsTemp!Records, 0) = 0 Then Exit Function
    
    '弹出选择器供操作员选择
    intSelect = frm选择当前医保_Base.ShowSelect
    ChooseInsure_Base = intSelect
End Function

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
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

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名,值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
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
                Case adVarChar
                    lngLength = madLongVarCharDefault
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

Public Function 多单据收费_收费分别打印() As Boolean
    #If gverControl >= 4 Then
        多单据收费_收费分别打印 = (Val(zlDatabase.GetPara(78, glngSys, , 0)) = 1)
    #Else
        多单据收费_收费分别打印 = (Val(GetPara(78, glngSys, , , 0)) = 1)
    #End If
End Function

Public Sub 多单据收费_退费(ByVal lng结帐ID As Long)
    '2.1、退费时，如果保险结算记录的备注中含“多单据收费”，并且该单据是多单据收费（票据打印内容大于1条记录）则继续执行
    '2.2、如果退费时该单据是单张收费（票据打印内容小于等于1条记录），则提取该病人该登记时间相同的所有单据号出来，提示操作员应该同时退费后再收取新的费用
    Dim strNO As String, str单据清单 As String
    Dim lng病人ID As Long
    Dim str登记时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '检查保险结算记录中是否记录的是多单据收费（注意，虽然保险结算记录中记录为多单据收费，可能由于勾选了系统参数“78-多张单据收费分别打印”，HIS并没有当做多单据来处理，退费时允许单张退，所以要判断
    gstrSQL = " Select 备注 From 保险结算记录 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否为多单据收费", lng结帐ID)
    If InStr(1, rsTemp!备注, "多单据收费") = 0 Then Exit Sub        '不是多单据收费直接退出
    '提取本次结帐的相关信息
    gstrSQL = " Select NO,病人ID,登记时间 From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取单据号与病人ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    strNO = rsTemp!NO
    str登记时间 = Format(rsTemp!登记时间, "yyyy-MM-dd HH:mm:ss")
    '根据票据打印内容判断是否多单据收费
    gstrSQL = " Select NO From 票据打印内容 Where ID=(Select ID From 票据打印内容 Where 数据性质=1 And NO=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "根据票据打印内容判断是否多单据收费", strNO)
    If rsTemp.RecordCount > 1 Then Exit Sub                       '多条记录说明未勾选系统参数，HIS认为是多单据收费
    '提取登记时间，病人ID相同的单据清单，提示操作员
    gstrSQL = " Select Distinct NO From 门诊费用记录 " & _
              " Where Mod(记录性质,10)=1 And 记录状态=1 And 病人ID=[1] And 登记时间=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取登记时间，病人ID相同的单据清单", lng病人ID, CDate(str登记时间))
    With rsTemp
        Do While Not .EOF
            If !NO <> strNO Then str单据清单 = str单据清单 & "," & !NO
            .MoveNext
        Loop
        If str单据清单 <> "" Then
            str单据清单 = Mid(str单据清单, 2)
            MsgBox "多单据收费，退费时请一并完成以下单据的退费，然后重新收费！" & vbCrLf & str单据清单, vbInformation, gstrSysName
        End If
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function
Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModul As Long, _
    Optional ByVal blnPrivate As Boolean, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的系统参数值
    '参数:varPara-参数号或参数名，以数字或字符类型传入区分
    '     lngSys-系统号(10.20.0以后版本有效)
    '     lngModul-模块号(10.20.0以后版本有效)
    '     blnPrivate-是否私有模块(10.20.0以后版本有效)
    '     strDefault-默认值
    '     blnNotCache-是否中缓存中读取(10.20.0以后版本有效)
    '返回:参数值
    '编制:
    '日期:2008/01/04
    '-------------------------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If IsToolInPara Then
        GetPara = CallByName(zlDatabase, "GetPara", VbMethod, varPara, IIf(lngSys = 0, glngSys, lngSys), lngModul, blnPrivate, strDefault, blnNotCache)
    Else
        If TypeName(varPara) = "String" Then
            gstrSQL = "Select 参数值 From 系统参数表 where 参数名=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取系统参数", CStr(varPara))
        Else
            gstrSQL = "Select 参数值 From 系统参数表 where 参数号=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取系统参数", Val(varPara))
        End If
        If rsTemp.RecordCount <> 0 Then
            GetPara = Nvl(rsTemp!参数值, strDefault)
        Else
            GetPara = strDefault
        End If
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsToolInPara() As Boolean
    '---------------------------------------------------------------------------------
    '功能:判断是否从zlTools中的zlParameters中读取参数值
    '返回:是,返回true,否则返回Fasle
    '编制:
    '日期:2008/01/04
    '---------------------------------------------------------------------------------
    Dim arrVersion
    Dim rsTemp As New ADODB.Recordset
    
    '因医保部件只有CodeMan()才能获取系统号，在读取参数时必须知道系统号，特写入注册表，如果医保读不到默认为 100
    glngSys = GetSetting("ZLSOFT", "公共全局", "系统号", 100)

    '取系统版本号
    gstrSQL = "Select 版本号 From zlSystems Where  编号 =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取系统版本号", glngSys)
    
    '判断版本号
    arrVersion = Split(rsTemp!版本号, ".")
    If arrVersion(0) = "10" Then
        '看此版本
        If Val(arrVersion(1)) < 20 Then
            '只有次版本在20以下才能读取
            IsToolInPara = False
        Else
            IsToolInPara = True
        End If
    End If
End Function

Public Function IsZLHIS10() As Boolean
    Dim arrVersion
    Dim rsTemp As New ADODB.Recordset
    
    '取系统版本号
    gstrSQL = "Select 版本号 From zlSystems Where Floor(编号/100)=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取系统版本号", 1)
   
    '判断版本号
    arrVersion = Split(rsTemp!版本号, ".")
    If arrVersion(0) = "10" Then
        IsZLHIS10 = True
    End If
End Function

Public Sub OpenRecordset_OtherBase(rsTmp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional gcnConnect As ADODB.Connection)
'功能：打开记录集
    If rsTmp.State = adStateOpen Then rsTmp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
'    rsTmp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    If gcnConnect Is Nothing Then Set gcnConnect = gcnOracle
    Set rsTmp = gcnConnect.Execute(IIf(strSQL = "", gstrSQL, strSQL))
    Call SQLTest
End Sub


Public Sub OpenRecordset(rsTmp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional gcnConnect As ADODB.Connection)
'功能：打开记录集
    If rsTmp.State = adStateOpen Then rsTmp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
'    rsTmp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    If gcnConnect Is Nothing Then Set gcnConnect = gcnOracle
    Set rsTmp = gcnConnect.Execute(IIf(strSQL = "", gstrSQL, strSQL))
    Call SQLTest
End Sub


