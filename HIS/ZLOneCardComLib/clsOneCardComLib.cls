VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOneCardComLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*********************************************************************************************************************************************
'功能:一卡通涉及的相关数据对象(包含获取一卡通的方法，数据集等)
'涉及的方法或函数:
'一.公共方法
'  1.1  zlInitComponents-初始化对象(第一步必须先初始化，否则可能使用中会发生异常)
'  1.2  ErrCenter-错误处理中心
'  1.3  SaveErrLog-保存错误日志
'  1.3  WritDebugLog-写调试日志
'  1.4  SetPara-参数设置
'  1.5  GetPara-获取参数
'  1.7  GetPrivFunc-获取权限串
'  1.6  CloseWindows-关闭windows所有窗口
 
'二.病人相关数据
'  1.病人查找及检查
'  1.1 zlGetPatiIDFromBedNumber:根据床位号，获取病人ID
'  1.2 zlGetPatiInforFromPatiID:根据病人ID获取病人信息对象
'  1.3 zlGetPatiInforFromXML:根据XML串，获取病人信息对象
'  1.4 zlCheckOutNoIsExist:检查病人的门诊号是否被使用
'  1.5 zlIsExistFeeInsurePatient:判断医保病人是否存在未结费用
'  2. 病人相关选择器
'  2.1 zlPubPatiSel-病人选择器接口
'  2.2.
'三. 卡对象相关
'  1.1 zlGetCards-获取有效的卡类别
'  1.2 zlGetOneCardTypes-重新获取医疗卡类别信息集
'  1.3 zlGetConsumerCardTypes-重新获取消费卡目录信息集
'  1.4 zlCardTypeSelect:卡类别选择器
'四、涉及的公共属性:
'  5.1 objThirdSwap-获取三方接口对象
'  5.2 objOneCardObject-一卡通对象
'  5.3 objDataBase-涉及数据库操作对象
'编制:刘兴洪
'日期:2018-12-03 11:06:14
'*********************************************************************************************************************************************
Private mcnOracle As ADODB.Connection
Private mstrDBUser As String, mlngModul As Long
Private mfrmMain As Object
Private mobjDataBase As clsDataBase
Private mobjOneCardObject As clsOneCardDataObject
Private mobjThirdSwap As clsThirdSwapInterface  '三方交易接口
Private mobjService As clsService '服务接口
Private mblnIsObjRegisterAlone As Boolean
Private mlngfrmMainHwnd As Long

 
Public Property Get objThirdSwap() As clsThirdSwapInterface
    If mobjThirdSwap Is Nothing Then
        Set mobjThirdSwap = New clsThirdSwapInterface
        Call mobjThirdSwap.InitCommon(mcnOracle)
    End If
    Set objThirdSwap = mobjThirdSwap
End Property
Public Property Get objOneCardObject() As clsOneCardDataObject
    Call zlGetOneCardDataObject(mcnOracle, mobjOneCardObject)
    Set objOneCardObject = mobjOneCardObject
End Property
Public Property Get objService() As clsService
    Set objService = mobjService
End Property
 

Public Function zlIsExistOraConnect() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否存在Oracle连接
    '编制:刘兴洪
    '日期:2018-12-06 18:17:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    zlIsExistOraConnect = mobjDataBase.zlIsExistOraConnect
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlInitComponents(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    ByVal cnOracle As ADODB.Connection, Optional blnDeviceSet As Boolean = False, Optional strExpand As String, _
    Optional ByVal blnIsObjRegisterAlone As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参: frmMain-调用的主窗体
    '        lngModule-HIS调用模块号
    '       lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       blnDeviceSet-设备设置调用初始化
    '       strExpand-扩展信息(可选传入:卡类别ID-不传时,表示全部初始化,传入时,只初始化指定的接口)
    '       blnIsObjRegisterAlone-是否使用独立的注册部件(True:使用:zlRegisterAlone.DLL,否则使用zlRegister.dll)
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 11:49:39
    '说明:如果是第三方接口主动读取数据,则在渠道更改的部件中传入连接.
    '    1.如果第三接口存在签到的情况，可以在此接口中进行签道.
    '    2.如果第三方接口读取数据,需要先初始化.
    '调用者:
    '    1.  进入门诊收费、挂号时调用本接口
    '    2.  进入医生站、护士站、医技站等时调用本接口
    '    3.  第三方接口需要主动读取数据时,必须先初始化.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    glngSys = lngSys: gstrDBUser = strDBUser: Set mfrmMain = frmMain
    mstrDBUser = strDBUser: mlngModul = lngModule: glngModul = lngModule
     
    Set mcnOracle = cnOracle
    If gcnOracle Is Nothing Then Set gcnOracle = cnOracle
    
    If mblnIsObjRegisterAlone <> blnIsObjRegisterAlone Then
        Set gobjOneDataBase = Nothing
    End If
    
    mblnIsObjRegisterAlone = blnIsObjRegisterAlone
    Call zlInitPublicVar '初始化全局变量
    If zlGetOneDataBase(cnOracle, mobjDataBase, blnIsObjRegisterAlone) = False Then Exit Function
    If zlGetOneCardDataObject(cnOracle, mobjOneCardObject) = False Then Exit Function
    
    If mobjThirdSwap Is Nothing Then Set mobjThirdSwap = New clsThirdSwapInterface
    Call mobjThirdSwap.InitCommon(cnOracle, mblnIsObjRegisterAlone)   '初始化三方接口变量
    If mobjThirdSwap.zlInitComponents(frmMain, lngModule, glngSys, strDBUser, cnOracle, blnDeviceSet, strExpand) = False Then Exit Function
    gstrUnitName = mobjDataBase.GetUnitName
    zlInitComponents = True
    
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function zlCardDevSet(ByVal frmMain As Object, Optional bytCardType As Byte = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlCardDevSet(设备参数设置接口):主要是配置相关的设备的相关参数
    '入参:frmMain Form    HIS传入 调用的主窗体
    '       bytCardType- 1-消费卡;2-医疗卡;0-不区分消费卡和医疗卡,统一设置
    '出参:
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    'HIS调用说明.
    '    1.  在门诊收费管理的参数设置'设备配置' zlCardDevSet
    '    2.  在住院结帐界面的参数设置'设备配置' zlCardDevSet
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCardDevSet = mobjThirdSwap.zlCardDevSet(frmMain, bytCardType)
End Function


Public Function zlGetIDKindStr(Optional strIDKindStr As String = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取读卡的种类
    '入参:strIDKindStr    String  IN
    '有两种格式:
    '一种是缺省的:短名1|全名1|读卡标志1;…. ;短名n|全名n|读卡标志n
    '另一种是该接口返回的格式:短名|全名|读卡标志|卡类别ID|卡号长度;…
    '出参:
    '返回: 短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '           是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
    '        其中:卡类别ID|长度是本次增加的,由调用者根据情况来确认.
    '       比如:身|身份证号|0|0|18;IC|IC卡号|1|0|8;门|门诊号|0|0|0;就|就诊卡|0|0|0;建|建行卡|0|0|10
    '      出现错误时,返回空
    '编制:刘兴洪
    '日期:2011-06-14 14:43:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetIDKindStr = objThirdSwap.zlGetIDKindStr(strIDKindStr)
End Function


Public Function zlAdviceSendBulidBarCode(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal intRecNature As Integer, ByVal strNos As String, Optional ByVal strExpand As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医嘱发送生成条码接口
    '入参:frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  医疗卡类别ID
    '    intRecNature    Integer In  记录性质
    '    strNos  String  In  单据号，多个用逗号分隔
    '    strExpand   String  In/Out  扩展参数,暂留，现为空
    '返回:只要一个接口调用返回True:否则返回False
    '编制:刘兴洪
    '日期:2016-03-22 13:35:10
    '说明：
    '１. 在医嘱发送并生成费用后打印指引单或处方单前调用此接口
    '２. 接口内部根据记录性质和单据号，记录"二维码"或"条码"信息。
    '３. 操作形式是：医嘱发送后根据接口生成的二维码打印在指引单或处方上，然后通过移动平台扫描二维码进行支付。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlAdviceSendBulidBarCode = objThirdSwap.zlAdviceSendBulidBarCode(frmMain, lngModule, lngCardTypeID, intRecNature, strNos, strExpand)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetAvailabilityCardType() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前工作站有效的支付卡,比如银行卡;消费卡等
    '返回:格式:短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式;…
    '编制:刘兴洪
    '日期:2011-06-14 15:16:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetAvailabilityCardType = objThirdSwap.zlGetAvailabilityCardType()
End Function

Public Function zlGetCards(ByVal bytType As Byte) As Cards
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的卡对象
    '入参:bytType-0-所有医疗卡;
    '             1-启用的医疗卡,
    '             2-所有存在三方账户的三方卡
    '             3-启用的三方账户的医疗卡
    '出参:
    '返回:成功返回本地支持的卡对象
    '编制:刘兴洪
    '日期:2013-10-23 17:43:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetCards = mobjThirdSwap.zlGetCards(bytType)
End Function

Public Function ZlGetPatiFamilyMember(ByVal byt查询类型 As Byte, ByVal lng病人ID As Long, _
    ByRef str家属IDs As String, Optional ByRef rs家属信息 As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID，获取该病人的家属成员信息
    '入参:
    '   byt查询类型=查询类型：0-只返回家属成员病人id；1-查询家属成员的基本信息
    '出参:
    '   str家属IDs=家属成员病人id,多个英文逗号分隔
    '   rs家属信息=家属成员的基本信息,仅 查询类型=1 是有效,字段：病人ID,关系,姓名,性别,年龄,出生日期,民族,身份证号
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZlGetPatiFamilyMember = mdlPubServerPati.Zl_Patisvr_GetPatiFamilyMember(byt查询类型, lng病人ID, str家属IDs, rs家属信息)
End Function

Public Function ZlUpdateOutpatiState(ByVal lng病人ID As Long, ByVal cllUpdateInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新病人就诊状态及部分信息
    '入参:
    '   lng病人ID=病人ID
    '   cllUpdateInfo=更新信息，成员:(手机号(N),费别(N),就诊诊室(N),就诊状态(N),就诊时间(N),门诊号(N))
    '出参:
    '返回:执行成功返回True，失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZlUpdateOutpatiState = mdlPubServerPati.Zl_Patisvr_UpdateOutpatiState(lng病人ID, cllUpdateInfo)
End Function

Public Function ZlGetPatiCardInfo(ByVal str病人ID As String, ByRef rsCardInfo As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID获取持有有效卡的卡类别信息
    '入参:
    '出参:
    '   rsCardInfo=卡信息，字段：
    '返回:获取成功返回True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllData As Collection
    Dim cllTemp As Collection
    
    On Error GoTo ErrHandler
    Set rsCardInfo = New ADODB.Recordset
    With rsCardInfo.fields
        .Append "病人ID", adBigInt, 18, adFldIsNullable
        .Append "姓名", adVarChar, 100, adFldIsNullable
        .Append "性别", adVarChar, 4, adFldIsNullable
        .Append "年龄", adVarChar, 20, adFldIsNullable
        .Append "出生日期", adVarChar, 30, adFldIsNullable
        .Append "门诊号", adVarChar, 20, adFldIsNullable
        .Append "身份证号", adVarChar, 18, adFldIsNullable
        .Append "卡类别ID", adBigInt, 18, adFldIsNullable
        .Append "卡号", adVarChar, 50, adFldIsNullable
        .Append "二维码", adVarChar, 100, adFldIsNullable
        .Append "密码", adVarChar, 50, adFldIsNullable
    End With
    rsCardInfo.CursorLocation = adUseClient
    rsCardInfo.LockType = adLockOptimistic
    rsCardInfo.CursorType = adOpenStatic
    rsCardInfo.Open
    
    If mdlPubServerPati.Zl_Patisvr_GetPatiCardInfo("", str病人ID, 2, , , cllData) = False Then Exit Function
    '        pati_id N   1   病人id
    '        pati_name   C       姓名
    '        pati_sex    C       性别
    '        pati_age    C       年龄
    '        pati_birthdate  C       出生日期：yyyy-mm-dd hh24:mi:ss
    '        outpatient_num  C       门诊号
    '        pati_idcard C       身份证号
    '        cardtype_id N   1   卡类别ID
    '        card_no C   1   卡号
    '        card_qrcode C   1   二维码
    '        card_passwod    C   1   密码
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        With rsCardInfo
            .AddNew
            !病人ID = cllTemp("_pati_id")
            !姓名 = NVL(cllTemp("_pati_name"))
            !性别 = NVL(cllTemp("_pati_sex"))
            !年龄 = NVL(cllTemp("_pati_age"))
            !出生日期 = NVL(cllTemp("_pati_birthdate"))
            !门诊号 = NVL(cllTemp("_outpatient_num"))
            !身份证号 = NVL(cllTemp("_pati_idcard"))
            !卡类别id = NVL(cllTemp("_cardtype_id"))
            !卡号 = NVL(cllTemp("_card_no"))
            !二维码 = NVL(cllTemp("_card_qrcode"))
            !密码 = NVL(cllTemp("_card_passwod"))
            .Update
        End With
    Next
    If rsCardInfo.RecordCount > 0 Then rsCardInfo.MoveFirst
    
    ZlGetPatiCardInfo = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
   Call gobjComLib.SaveErrLog
End Function

Public Function zlGetPatiID(ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False, _
    Optional ByRef intCardStatus As Integer, Optional ByRef strValidTime As String, _
    Optional ByVal bytPatiInfoShowType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的医疗类别和卡号,获取对应的病人ID
    '入参:strCardType-卡类别,如果为数字,这为卡类别ID,如果为字符,则为类别名称
    '       strCardNo-卡号
    '       blnNotShowErrMsg-不显示错误的提示信息
    '       frmMain-调用的主窗体
    '       objCtl-调用的控件
    '       blnShowMergePati-当出现多个满足条件的病人时,是否显示合并功能按钮
    '       blnOnlyContractPati-签约病人
    '       blnUserCancel-选择器中，用户选择了取消
    '       lngShowCardNoTypeID-过滤出多条病信息时，弹出选择器中显示的卡号的卡类别ID,0-表示不显示卡号；>0表示显示指定卡号类别的ID
    '       blnNotCheckValidDate-是否对卡终止使用时间进行检查,true-不检查终止使用时间,false-检查
    '       intCardStatus-返回卡状态
    '       bytPatiInfoShowType-病区选择器列显示方式：0-所有信息，1-简略信息
    '出参:strErrMsg-返回的错误信息
    '       lng病人ID-返回的病人ID
    '       strCardPass-返回卡号的密码
    '       lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    zlGetPatiID = mobjOneCardObject.zlGetPatiID(strCardType, strCardNo, blnNotShowErrMsg, lng病人ID, _
      strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, _
      blnCertificate, blnUserCancel, lngShowCardNoTypeID, blnNotCheckValidDate, , intCardStatus, strValidTime, bytPatiInfoShowType)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiInfsByPatiIds(ByVal intQueryStatus As Integer, ByVal blnExpendInfo As Boolean, _
                ByVal str病人Ids As String, ByRef cllPatiInfos_out As Collection, Optional ByVal cllFilter As Collection, _
                Optional ByVal int场合 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息集
    '入参:intQueryStatus-查询类型(0-仅门诊;1-在院 ;2-门诊及在院)
    '     str病人Ids-病人ID
    '出参:cllPatiInfos_out-返回的数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If cllFilter Is Nothing Then Set cllFilter = New Collection
    zlGetPatiInfsByPatiIds = mobjService.zlPatiSvr_GetPatiInfsByRange(intQueryStatus, _
        cllFilter, cllPatiInfos_out, str病人Ids, , blnExpendInfo)
End Function

Public Function zlGetPatiInfsByFilter(ByVal intQueryStatus As Integer, ByVal blnExpendInfo As Boolean, _
                ByVal cllFilter As Collection, ByRef cllPatiInfos_out As Collection, _
                Optional ByVal int场合 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息集
    '入参:intQueryStatus-查询类型(0-仅门诊;1-在院 ;2-门诊及在院)
    '     cllFilter-Array(查询条件,条件值1,条件值2(N))
    '出参:cllPatiInfos_out-返回的数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If cllFilter Is Nothing Then Exit Function
    If cllFilter.count = 0 Then Exit Function
    zlGetPatiInfsByFilter = mobjService.zlPatiSvr_GetPatiInfsByRange(intQueryStatus, _
        cllFilter, cllPatiInfos_out, , , blnExpendInfo)
End Function

Public Function zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNos As String, _
    Optional ByVal strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       strCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       dblMoney-支付金额(退款时为负数)
    '       strNos-本次支付所涉及的单据
   '       strXMLExpend-XML串:
    '          1.验证密码:自助机用)
    '                <Data>
    '                   <验证密码></验证密码>
    '                </Data>

    '          2.消费卡收款时，传入XML串：
    '            消费卡收款及扫码付时，传入XML串：
    '            <IN>
    '                <MZXSJE>面值销售金额</MZXSJE>
    '                <CZJKJE>充值缴款金额</CZJKJE>
    '                <QRCODE>识别的二维码代码串</QRCODE> //由zlReadQRCode.clsReadQRCode中的ReadRQCode返回,10.35.130后支持:
    '            </IN>

    '出参:
    '   strXMLExpend-(XML串:错误信息，自助机用)
    '                <Data>
    '                   <错误信息></错误信息>
    '                </Data>
    '返回:扣款合法,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 16:42:43
    '说明:
    '   在调用扣款前，由于存在Oracle事务问题， 所以再调用扣款交易前， _
    '   先进行数据的合法性检查,以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlPaymentCheck = mobjThirdSwap.zlPaymentCheck(frmMain, lngModule, strCardTypeID, bln消费卡, strCardNo, dblMoney, strNos, strXMLExpend)
End Function


Public Function zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, ByVal strPrepayNos As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strXMLExpend As String, _
    Optional ByVal blnResolveXMLToRecord As Boolean, _
    Optional ByRef rsBalance_Out As ADODB.Recordset, Optional rsExpend_Out As ADODB.Recordset, _
    Optional intStatus_Out As Integer, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '    frmMain-调用的主窗体
    '    lngModule-HIS调用模块号
    '    lngCardTypeID-卡类别ID
    '    strCardNo -卡号
    '    strBalanceIDs-结帐ID,多个用逗号分离
    '    strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离、
    '    dblMoney-支付金额
    '    blnResolveXMLToRecord-是否解析XML串给记录订(rsBalance_Out,rsExpend_Out
    '    strXMLExpend:10.35.130以后才支持并将地原参数名《strXMLExpend》调整为了《strXMLEXPEND》
    '            <INPUT>
    '                <QRCODE>识别的二维码代码串</QRCODE> //识别的二维码代码串，10.35.130后支持该节点,由zlReadQRCode.clsReadQRCode中的ReadRQCode返回
    '            </INPUT>
    '出参:strSwapGlideNO-交易流水号(XML中返回结算信息时,交易流水号，以xml为准)
    '    strSwapMemo-交易说明(XML中返回返回结算信息时,交易以xml为准)
    '    rsBalance_Out-结算信息集(blnResolveXMLToRecord=true时返回)字段（卡类别ID,消费卡,结算方式,交易金额,交易流水号,交易说明,单据号,单据性质,卡号,结算号码,结算摘要,是否普通结算)
    '    rsExpend_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
    '    intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
    '    strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
    '    strXMLExpend -交易扩展信息,有两种方式:
    '    1)  格式1为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '    2)  格式2(10.35.110以后才支持),XML格式
    '        <OUTPUT>
    '            <JYZT>交易状态</JYZT> //接口返回False时，此参数有效:0-交易调用失败;1-交易正在处理中
    '            <ERRMSG>错误信息</ERRMSG> //误信息:为空时，不提示，非空时，提示
    '            <JYLIST>//交易列表
    '                <JY>//保存到预交记录时，按交易流水号及交易说明汇总处理.
    '                    <JYFS>交易方式</JYFS>//Varchar2 20:即结算方式.名称
    '                    <JYJE>交易金额</JYJE>//Number 15 2
    '                    <JYLSH>交易流水号</JYLSH>//Varchar2 50
    '                    <JYSM>交易说明</JYSM>   Varchar2 500
    '                    <JSHM>结算号码</JSHM>   Varchar2 50
    '                    <JSZY>结算摘要</JSZY>   Varchar2 100
    '                    <KH>卡号</KH>   Varchar2 50
    '                    <DJH>单据号<DJH> //Varchar2 8,多单据收费或挂号与卡费一起收时有用 ，针对收费：存储在"医保结算明细"表中,主要是分单据保存，针对挂号费和卡费时，分别存在储病人预交记录中
    '                    <DJXZ>单据性质</DJXZ>//number 2,1多单据收费或挂号与卡费一起收时有用,1-收费单;4-挂号单;5-就诊卡费
    '                    <SFPTJS>是否普通结算</SFPTJS>//Number 1 是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
    '                </JY>
    '                ...
    '            </JYLIST>
    '            <Expends>               扩展信息列表
    '                <Expend>
    '                    <XMMC>项目名称</XMMC> //Varchar2  50
    '                    <XMNR>项目内容</XMNR> //Varchar2  100
    '                </Expend>
    '                ...
    '            </Expends>
    '        </OUTPUT>
    '返回:扣款成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:13:48
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlPaymentMoney = mobjThirdSwap.zlPaymentMoney(frmMain, lngModule, lngCardTypeID, bln消费卡, strCardNo, strBalanceIDs, strPrepayNos, dblMoney, strSwapGlideNO, strSwapMemo, strXMLExpend, blnResolveXMLToRecord, rsBalance_Out, rsExpend_Out, intStatus_Out, strErrMsg_Out)
End Function

Public Function zlBrushCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByRef dbl金额 As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varSquareBalance As Variant, _
    Optional ByVal bln转预交 As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str费用来源 As String, _
    Optional ByVal lng病人ID As Long, _
    Optional ByRef str结算方式_Out As String = "", _
    Optional ByRef dbl结算金额_Out As Double = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定支付类别,弹出刷卡窗口
    '入参:rsClassMoney:收费类别,金额
    '        lngCardTypeID-为零时,为老一卡通刷卡
    '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
    '       dblBrushTotaled-消费有效,表示已经刷消费卡总额(主要用于多次刷卡)
    '       str上次限制类别-上次刷消费时的限制类别(同次多次刷消费卡时,需要检查本次刷卡类别与上次类别是否一致,不一致不允许刷卡消费)
    '       varSquareBalance- Collection类型,当前已经刷卡的信息(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文 ))
    '       bln预交-是否转预交
    '       blnAllPay-是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
    '       strXMLExpend-三方卡调用XML入参,目前格式如下:
    '       <IN>
    '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
    '           <QRCODE></QRCODE> //识别的二维码代码串，10.35.130后支持该节点,由zlReadQRCode.clsReadQRCode中的ReadRQCode返回
    '       </IN>
    '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '       lng病人ID - 病人ID(使用消费卡支付时传入)
    '出参:str限制类别-限制类别(消费卡返回)
    '        lng消费卡ID-消费卡信息.ID(消费卡返回)
    '       strCardNO-返回刷卡的卡号
    '       strPassWord-返回刷卡所对应的密码
    '       varSquareBalance- Collection类型,返回当前刷卡数据(array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文))
    '       str结算方式_Out-返回的结算方式
    '       dbl结算金额_Out-返回的结算金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:18:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBrushCard = mobjThirdSwap.zlBrushCard(frmMain, lngModule, rsClassMoney, lngCardTypeID, bln消费卡, _
        strPatiName, strSex, strOld, dbl金额, strCardNo, strPassWord, bln退费, blnShowPatiInfor, bln退现, _
        bln余额不足禁止, varSquareBalance, bln转预交, blnAllPay, strXmlIn, str费用来源, lng病人ID, str结算方式_Out, dbl结算金额_Out)
End Function


Public Function zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户回退交易前的检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额(正数)
    '       strSwapNo-交易流水号(退款时检查)
    '       strSwapMemo-交易说明(退款时传入)
    '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
    '返回:退款合法,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:24:55
    '说明:
    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
    '    以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlReturncheck = mobjThirdSwap.zlReturncheck(frmMain, lngModule, lngCardTypeID, bln消费卡, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strXMLExpend)
End Function
Public Function zlReturnCashCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String, _
    Optional blnDelCash_Out As Boolean, Optional strDefaultDelBalance_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退现交易检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(退款时检查)多种结算方式时，本参数为第一个结算方式的交易流水号
    '       strSwapMemo-交易说明(退款时传入) 多种结算方式时，本参数为第一个结算方式的交易说明
    '       strXMLExpend    XML IN  10.35.110后才支持
    '        <INPUT>
    '            <TKLIST>    //本次退款列表
    '                <TK>
    '                    <TKFS>退款方式</TKFS>
    '                    <TKJE>退款金额</TKJE>
    '                    <JYLSH>原交易流水号</JYLSH>
    '                    <JYSM>原交易说明</JYSM>
    '                </TK>
    '                ....
    '            </TKLIST>
    '        </INPUT>
    '出参:
    '       blnDelCash_Out-是否允许退现：
    '                1. 接口返回True时：表示是否缺省退现: True-表示缺省退现;False-表示不缺省退现,缺省退回原卡，但允许操作员操作退现
    '                2.接口返回False时：表示是否允许强制退现, True-表示允许通过"强制退现"权限来控制是否允许退现;False-表示禁止通过"强制退现"权限来退现
    '       strDefaultDelBalance_Out-缺省退现方式,比如：支票，现金等
    '       strXMLExpend:10.35.110后才支持
    '        <OUTPUT>
    '            <TXZT>退现状态<TXZT>//接口返回True时：  表示是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但允许操作员操作退现;
    '                                  接口返回False时： 表示是否允许强制退现:1-表示允许通过"强制退现"权限来控制是否允许退现,0-表示禁止通过"强制退现"权限来退现
    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
    '                    1.不允许返回三卡方的结算方式
    '                    2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
    '        </OUTPUT>
    '返回:退现合法,返回true,否则返回Flase
    '编制:刘尔旋
    '日期:2017-7-3
    '说明:
    '    1.使用模块：病人结帐管理
    '    2.触发时机:在结帐列表中点“X"时，调用，如果返回true,则允许退现，否则禁止
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlReturnCashCheck = mobjThirdSwap.zlReturnCashCheck(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strXMLExpend, blnDelCash_Out, strDefaultDelBalance_Out)
End Function



Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, _
    ByVal strSwapGlideNO As String, ByVal strSwapMemo As String, _
    ByRef strSwapExpendInfor As String, Optional strExpendXML As String, _
    Optional ByVal blnResolveXMLToRecord As Boolean, _
    Optional ByRef rsBalance_Out As ADODB.Recordset, Optional ByRef rsExpend_Out As ADODB.Recordset, _
    Optional intStatus_Out As Integer, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExpendInfor-本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       strExpendXML:10.35.110后，扩展此参数,XML格式，如下:
    '            <INPUT>
    '                <TKLIST>    //本次退款列表
    '                    <TK>
    '                        <TKFS>退款方式</TKFS>// Varchar2 20 即结算方式.名称
    '                        <TKJE>退款金额</TKJE>// Number  15  2
    '                        <JYLSH>原交易流水号</JYLSH>//   Varchar2    50
    '                        <JYSM>原交易说明</JYSM> //  Varchar2    500
    '                        <DJH>单据号</DJH> //    Varchar2 8 单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '                    </TK>
    '                    ....
    '                </TKLIST>
    '            </INPUT
    '    blnResolveXMLToRecord-是否解析XML串给记录集(rsBalance_Out,rsExpend_Out）
    '出参: strSwapNo-交易流水号(退款交易流水号)
    '      strSwapMemo-交易说明(退款交易说明)
    '    intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
    '    strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
    '      strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    '       strExpendXML：10.35.110以后才有此参数：格式如下
    '            <OUTPUT>
    '                <JYZT>交易状态</JYZT> //接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '                <ERRMSG>错误信息</ERRMSG>//:为空时，不提示，非空时，提示
    '                <TKLIST> //支付列表( 表示与原退款方式不一致才返回)
    '                    <TK> // 保存到预交记录时，按交易流水号及交易说明汇总处理.
    '                        <TKFS>退款方式</TKFS> // Varchar2 20 即结算方式.名称
    '                        <TKJE>结算金额</TKJE>// Number 12 2
    '                        <JYLSH>退款交易流水号</JYLSH>//Varchar2 50
    '                        <JYSM>退款交易说明</JYSM>// Varchar2 500
    '                        <YJYLSH>原交易流水号</YJYLSH>//Varchar2 50
    '                        <YJYSM>原交易说明</YJYSM>// Varchar2 500
    '                        <DJH>单据号</DJH> //    Varchar2 8 单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '                        <SFPTJS>是否普通结算</SFPTJS>//Number 1 是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
    '                    </TK>
    '                    ...
    '                </TKLIST>
    '            </OUTPUT>
    '    rsBalance_Out-结算信息集(blnResolveXMLToRecord=true时返回)字段（卡类别ID,消费卡,结算方式,交易金额,交易流水号,交易说明,单据号,单据性质,卡号,结算号码,结算摘要,是否普通结算)
    '    rsExpend_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:31:32
    '说明:
    '       在所有需要对已经消费的地方进行退款的情况下，调用该接口,目前规划在:收费室；挂号等。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExpendInfor-本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       strExpendXML:10.35.110后，扩展此参数,XML格式，如下:
    '            <INPUT>
    '                <TKLIST>    //本次退款列表
    '                    <TK>
    '                        <TKFS>退款方式</TKFS>// Varchar2 20 即结算方式.名称
    '                        <TKJE>退款金额</TKJE>// Number  15  2
    '                        <JYLSH>原交易流水号</JYLSH>//   Varchar2    50
    '                        <JYSM>原交易说明</JYSM> //  Varchar2    500
    '                        <DJH>单据号</DJH> //    Varchar2 8 单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '                    </TK>
    '                    ....
    '                </TKLIST>
    '            </INPUT
    '    blnResolveXMLToRecord-是否解析XML串给记录集(rsBalance_Out,rsExpend_Out）
    '出参: strSwapNo-交易流水号(退款交易流水号)
    '      strSwapMemo-交易说明(退款交易说明)
    '      strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    '       strExpendXML：10.35.110以后才有此参数：格式如下
    '            <OUTPUT>
    '                <JYZT>交易状态</JYZT> //接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '                <ERRMSG>错误信息</ERRMSG>//:为空时，不提示，非空时，提示
    '                <TKLIST> //支付列表( 表示与原退款方式不一致才返回)
    '                    <TK> // 保存到预交记录时，按交易流水号及交易说明汇总处理.
    '                        <TKFS>退款方式</TKFS> // Varchar2 20 即结算方式.名称
    '                        <TKJE>结算金额</TKJE>// Number 12 2
    '                        <JYLSH>退款交易流水号</JYLSH>//Varchar2 50
    '                        <JYSM>退款交易说明</JYSM>// Varchar2 500
    '                        <DJH>单据号</DJH> //    Varchar2 8 单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '                        <SFPTJS>是否普通结算</SFPTJS>//Number 1 是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
    '                    </TK>
    '                    ...
    '                </TKLIST>
    '            </OUTPUT>
    '    rsBalance_Out-结算信息集(blnResolveXMLToRecord=true时返回)字段（卡类别ID,消费卡,结算方式,交易金额,交易流水号,交易说明,单据号,单据性质,卡号,结算号码,结算摘要,是否普通结算)
    '    rsExpend_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:31:32
    '说明:
    '       在所有需要对已经消费的地方进行退款的情况下，调用该接口,目前规划在:收费室；挂号等。
    '      １.需要注意三方是否支持部分退费，即三方中心是否可以进行部分冲销，例如MISPOS是不能进行部分冲销的。即支付冲销时必须进行全部退款。不允许部分退款需要设置【医疗卡类别.是否全退】来进行判断。目前默认为允许部分退费用。若不允许请修改相应的值。
    '      ２.其它一些退费前的检测，如检测退费时是否需要刷卡，卡号原交易卡号必须一致。避免由于点错导致退费，给医院造成损失等。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlReturnMoney = mobjThirdSwap.zlReturnMoney(frmMain, lngModule, lngCardTypeID, bln消费卡, strCardNo, strBalanceIDs, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExpendInfor, strExpendXML, blnResolveXMLToRecord, rsBalance_Out, rsExpend_Out, intStatus_Out, strErrMsg_Out)

End Function

Public Function zlReturnMultiMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strInXML As String, _
    ByVal lng冲销ID As Long, ByRef strOutXml As String, ByRef strExpend As String, _
    Optional ByVal blnResolveXMLToRecord As Boolean, _
    Optional ByRef rsBalance_Out As ADODB.Recordset, Optional ByRef rsExpend_Out As ADODB.Recordset, _
    Optional intStatus_Out As Integer, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易(多笔回退)
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strInXML-XML串:
    '            <JSLIST>
    '               <JS>
    '                    <TKFS>退款方式</TKFS>  //10.35.110后新增,退款方式,即结算方式.名称
    '                    <KH>卡号</KH>
    '                    <JYLSH>交易流水号</JYLX>
    '                    <JYSM>交易说明</JYSM>
    '                    <ZFJE>作废金额</ZFJE>
    '                    <JSLX>类型</JSLX>
    '                    <DJH>单据号</DJH>  //10.35.110新增:单据号,多单据收费时有用 ，存储在"医保结算明细"表中,主要是分单据保存
    '                    <GLJYID>关联交易ID</GLJYID>  //10.35.110新增:原始交易关联ID(病人预交记录.关联交易ID)
    '                    <ID></ID>
    '                </JS>
    '            </JSLIST>

    '       lng冲销ID-作废时的冲销ID(作废时或退费时有效，否则为0）;类型=6时，传入结算序号
    '       strExpend-无（暂留，待以后扩展)
    '    blnResolveXMLToRecord-是否解析XML串给记录集(rsBalance_Out,rsExpend_Out）
    '出参:
    '     strOutXML-返回XML串
    '     <OUTPUT>                      //10.35.110中增加此接点(接口作了兼容处理)
    '         <JYZT>交易状态</JYZT>    10.35.110中增加此接点:接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '         <ERRMSG>错误信息</ERRMSG> 10.35.110中增加此接点(接口作了兼容处理):错误信息:为空时，不提示，非空时，提示
    '         <JSLIST>
    '            <JS>
    '                <TKFS>退款方式</TKFS>   //10.35.110新增:退款方式即结算方式.名称
    '                <TKJE>退款方式</TKJE>   //10.35.110退款金额
    '                <KH>卡号</KH>
    '                <TKLSH>退款交易流水号</TKLSH>
    '                <TKSM>退款交易说明</TKSM>
    '                <ID></ID>   //类型=1时,预交ID;类型=2,6时，为原结帐ID
    '                <DJH>单据号</DJH>   //10.35.110后增加:单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '                    <GLJYID>关联交易ID</GLJYID>  //10.35.110新增:原始交易关联ID(病人预交记录.关联交易ID)
    '                <SFPTJS>是否普通结算</SFPTJS> //是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
    '            </JS>
    '            ...
    '        </JSLIST>
    '     </OUTPUT>
    '      strExpend-交易的扩展信息
    '       <Expends>
    '           <Expend>
    '               <XMMC>项目名称1</XMMC>
    '               <XMNR>项目内容2</XMNR>
    '           </Expend>
    '       </Expends>
    '    rsBalance_Out-结算信息集(blnResolveXMLToRecord=true时返回)字段（卡类别ID,消费卡,结算方式,交易金额,交易流水号,交易说明,单据号,单据性质,卡号,结算号码,结算摘要,是否普通结算)
    '    rsExpend_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
    '    intStatus_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '    strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
    '返回:函数返回    True:调用成功,False:调用失败
    '日期:2015-11-10
    '说明:
    '   目前只有结帐程序时有效（结帐退款),用于一次性处理同一卡类别的多笔三方交易作废
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    zlReturnMultiMoney = mobjThirdSwap.zlReturnMultiMoney(frmMain, lngModule, lngCardTypeID, bln消费卡, strInXML, lng冲销ID, strOutXml, strExpend, blnResolveXMLToRecord, rsBalance_Out, rsExpend_Out, intStatus_Out, strErrMsg_Out)
End Function

Public Function zlReadCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal blnOlnyCardNO As Boolean, _
    ByVal strExpand As String, _
    ByRef strOutCardNo As String, _
    ByRef strOutPatiInforXml As String, _
    Optional ByRef strPhotoFile As String = "", _
    Optional ByVal bln消费卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读卡接口
    '入参:frmMain-调用的父窗口
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       blnOlnyCardNO-仅仅读取卡号
    '       strExpand-扩展参数
    '出参:strOutCardNO-返回的卡号
    '       strOutPatiInforXML-(病人信息返回.XML串)
    '       strPhotoFile-当从卡中读取取像片后,传给HIS的像片文件(含完整路径)
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 11:56:41
    '说明:
    '   １. 在所有绑定卡时,需要建立病人信息档案时，读取该接口
    '   ２. 在所有需要刷卡进行就诊的地方，都需要调用该接口
    'strOutPatiInforXML参数说明
    '    参数: blnOlnyCardNO=true时,返回空
    '    根接点：<DATA> ,如:
    '    <DATA>
    '       <卡号></卡号>
    '        .....
    '    </DATA>
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    '    姓名    Varchar2    100
    '    性别    Varchar2    4
    '    年龄    Varchar2    10
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '    出生地点    Varchar2    50
    '    身份证号    VARCHAR2    18
    '    其他证件    Varchar2    20
    '    职业    Varchar2    80
    '    民族    Varchar2    20
    '    国籍    Varchar2    30
    '    学历    Varchar2    10
    '    婚姻状况    Varchar2    4
    '    区域    Varchar2    30
    '    家庭地址    Varchar2    50
    '    家庭电话    Varchar2    20
    '    家庭地址邮编    Varchar2    6
    '    监护人  Varchar2    64
    '    联系人姓名  Varchar2    64
    '    联系人关系  Varchar2    30
    '    联系人地址  Varchar2    50
    '    联系人电话  Varchar2    20
    '    工作单位    Varchar2    100
    '    单位电话    Varchar2    20
    '    单位邮编    Varchar2    6
    '    单位开户行  Varchar2    50
    '    单位帐号    Varchar2    20
    '   ABO血型 Varchar2    10
    '   RH  Varchar2    10
    '   哮喘标志    Number  2
    '   心脏病标志  Number  2
    '   心脑血管病标志  Number  2
    '   癫痫病标志  Number  2
    '   凝血紊乱标志    Number  2
    '   糖尿病标志  Number  2
    '   青光眼标志  Number  2
    '   透析标志    Number  2
    '   器官移植标志    Number  2
    '   器官缺失标志    Number  2
    '   可装卸义肢标志  Number  2
    '   心脏起搏器标志  Number  2
    '   其他医学警示    Varchar2    100
    '   联系信息
    '       姓名    Varchar2    30
    '       关系    Varchar2    30
    '       电话    Varchar2    30
    '       身份证号    Varchar2    30
    '   健康档案编号    Varchar2    18
    '   新农合证号  Varchar2    18
    '   其他证件
    '       信息名  Varchar2    15
    '       信息值  Varchar2    100
    '   其他信息 可以返回上面接点不存在的数据
    '       信息名  Varchar2    20      例: 工作单位, 血压等
    '       信息值  Varchar2    50      比如:信息名=血压;信息值='XXmmHg'
    '   过敏情况        30      以下信息存放于病人过敏药物表中
    '       药物名称    Varchar2    100
    '       药物反应    Varchar2    101
    '   免疫记录
    '       疫苗名称    Varchar2    100
    '       接种时间    Varchar2    8
    '   医疗卡属性
    '       信息名  Varchar2    20
    '       信息值  Varchar2    50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlReadCard = mobjThirdSwap.zlReadCard(frmMain, lngModule, lngCardTypeID, blnOlnyCardNO, strExpand, strOutCardNo, strOutPatiInforXml, strPhotoFile, bln消费卡)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlMakeCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal stInputPatiInforXml As String, _
    ByRef strPhotoFile As String, _
    ByRef strOutPatiInforXml As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:制卡接口
    '入参:
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  卡类别ID
    '    stInputPatiInforXml String(XML) In  HIS传入已经健档的病人信息
    '    strPhotoFile    String  In  表示像片文件(含完整路径)
    '    strExpend   String  In   XML,暂无,待以后扩展
    '出参:
    '    strOutPatiInforXml  String(XML) Out 调用接口后,传给HIS系统已经制卡后的病人信息,以便建立病人档案
    '    strExpend   String  Out  XML,暂无,待以后扩展
    '    strPhotoFile    String  In  表示像片文件(含完整路径)
    '返回 :True:调用成功,False:调用失败
    '调用者:
    '    目前只在医疗卡发放管理中调用此接口(主要是考虑输入的信息较多,其他窗口不适应)，相关的调用流程医疗卡发放管理的制卡
    '  strInputPatiInforXML及strOutPatiInforXML格式
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    '    姓名    Varchar2    100
    '    性别    Varchar2    4
    '    年龄    Varchar2    10
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '    出生地点    Varchar2    50
    '    身份证号    VARCHAR2    18
    '    其他证件    Varchar2    20
    '    职业    Varchar2    80
    '    民族    Varchar2    20
    '    国籍    Varchar2    30
    '    学历    Varchar2    10
    '    婚姻状况    Varchar2    4
    '    区域    Varchar2    30
    '    家庭地址    Varchar2    50
    '    家庭电话    Varchar2    20
    '    户口邮编    Varchar2    6
    '    监护人  Varchar2    64
    '    联系人姓名  Varchar2    64
    '    联系人关系  Varchar2    30
    '    联系人地址  Varchar2    50
    '    联系人电话  Varchar2    20
    '    工作单位    Varchar2    100
    '    单位电话    Varchar2    20
    '    单位邮编    Varchar2    6
    '    单位开户行  Varchar2    50
    '    单位帐号    Varchar2    20
    '    病人ID  Number  18      可以为空
    '    ABO血型 Varchar2    10      信息名='ABO'
    '    信息值=(A型;B型;O型;AB型;不详)
    '    RH  Varchar2    10      信息名='RH'
    '    信息值=(阴;阳;不详;未查)
    '    哮喘标志    Number  2       信息名='医学警示'
    '    信息值=' 哮喘;心脏病; 心脑血管病;…'
    '    心脏病标志  Number  2
    '    心脑血管病标志  Number  2
    '    癫痫病标志  Number  2
    '    凝血紊乱标志    Number  2
    '    糖尿病标志  Number  2
    '    青光眼标志  Number  2
    '    透析标志    Number  2
    '    器官移植标志    Number  2
    '    器官缺失标志    Number  2
    '    可装卸义肢标志  Number  2
    '    心脏起搏器标志  Number  2
    '    其他医学警示    Varchar2    100
    '    联系信息
    '        姓名    Varchar2    30
    '        关系    Varchar2    30
    '        电话    Varchar2    30
    '        身份证号    Varchar2    30
    '    健康档案编号    Varchar2    18
    '    新农合证号  Varchar2    18
    '    其他证件
    '        信息名  Varchar2    15
    '        信息值  Varchar2    100
    '    其他信息 可以返回上面接点不存在的数据
    '        信息名  Varchar2    20      例: 工作单位, 血压等
    '        信息值  Varchar2    50      比如:信息名=血压;信息值='XXmmHg'
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
 
    zlMakeCard = mobjThirdSwap.zlMakeCard(frmMain, lngModule, lngCardTypeID, stInputPatiInforXml, _
        strPhotoFile, strOutPatiInforXml, strExpend)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlBandCardArfter(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  卡类别ID
    '    lng病人ID   Long    IN   病人ID
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用者:
    '    1.  医疗卡发放管理中绑定卡后调用
    '    2.  病人入院登记;病人挂号管理;病人信息登记中进行卡绑定时调用
    '其他说明:
    '    相关的调用流程见后续各模块的绑定卡的详细说明
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle

    zlBandCardArfter = mobjThirdSwap.zlBandCardArfter(frmMain, lngModule, lngCardTypeID, _
     lng病人ID, strExpend)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlMzInforWriteToCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    ByVal lngBalanceID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写门诊信息接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  传入写卡类别ID:
    '           1)传入刷卡的类别ID
    '           2)传入零时,需要选择某个卡类别ID
    '    lng病人ID   Long    In  病人ID
    '    lngBalanceID    Long    In  结算序号(某次结算的序号)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用时机:
    '         医疗卡类别.是否写卡=1才调用
    '调用者:
    '    1.  目前收费管理完成后,调用该接口
    '    2.  退费完成后,调用该接口
    '说明:
    '        门诊相关信息的写入 , 主要是渠道编制的接口内部写入
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlMzInforWriteToCard = mobjThirdSwap.zlMzInforWriteToCard(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, lngBalanceID, strExpend)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlZyInforWriteToCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    ByVal lngBalanceID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写住院信息接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  传入写卡类别ID:
    '           1)传入刷卡的类别ID
    '           2)传入零时,需要选择某个卡类别ID
    '    lng病人ID   Long    In  病人ID
    '    lngBalanceID    Long    In  结帐ID(可以不传入)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用时机:
    '        医疗卡类别.是否写卡=1才调用
    '调用者:
    '    1.  出院时
    '    2.  结帐时
    '    3.  结帐作废时
    '说明:
    '       住院相关信息的写入 , 主要是渠道编制的接口内部写入
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlZyInforWriteToCard = mobjThirdSwap.zlZyInforWriteToCard(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, lngBalanceID, strExpend)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlHealthArchiveIsSHow(frmMain As Object, _
    ByVal lngModule As Long, _
    ByRef strFunName As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:居民健档档案
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    strFunName  String  Out 返回功能名, 即在医生工作站的菜单中显示的名称(菜单位置放在"工具"菜单下)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    '        Boolean 函数返回    True:表示应该加入档案显示菜单,False:不加入档案菜单
    '调用者:
    '    1.  门诊(住院)医生工作站调用,医生站
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlHealthArchiveIsSHow = mobjThirdSwap.zlHealthArchiveIsSHow(frmMain, lngModule, strFunName, strExpend)
End Function

Public Function zlHealthArchivesShow(frmMain As Object, _
    ByVal lngModule As Long, ByVal lng病人ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:居民健档档案
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lng病人ID-病人ID
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    医生工作站不做返回判断,渠道编制的接口部件中可以返回值
    ' 调用者:门诊(住院)医生工作站后调用
    '说明:     该接口分两层:
    '    1.  一层是在zl9CardSquare部件中该接口不返回值,由医生站直接调用,该接口调用渠道编制的接口界面,否则按标准的健康卡档案界面显示(见下图)
    '    2.  二层是渠道编制的接口部件:该接口允许返回值为TRUE时,表示利用渠道自己编制的界面展示,否则将按HIS固定的健康档案信息显示(见下图)
    '调用者:
    '    1.  门诊(住院)医生工作站调用,医生站
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlHealthArchivesShow = mobjThirdSwap.zlHealthArchivesShow(frmMain, lngModule, lng病人ID, strExpend)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlInitInterFacel(ByVal lngCardTypeID As Long, _
    Optional bln消费卡 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化指定卡接口
    '入参:lngCardTypeID-指定卡类别
    '       bln消费卡-是否消费卡
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 15:29:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlInitInterFacel = mobjThirdSwap.zlInitInterFacel(lngCardTypeID, bln消费卡)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, strExpand As String, dblMoney As Double, Optional bln消费卡 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取帐户余额
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '       strCardNo-卡号
    '       strExpand-预留，为空,以后扩展
    '       bln消费卡-是否为消费卡
    '出参:dblMoney-返回帐户余额
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 16:29:48
    '说明:
    '       在所有需要扣款的地方，都要检查帐户余额是否充足,帐户不充足时不允许扣款.
    '       如果某些第三方接口不存在余额接口，可以固定返回一定的金额。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    zlGetAccountMoney = mobjThirdSwap.zlGetAccountMoney(frmMain, lngModule, lngCardTypeID, strCardNo, strExpand, dblMoney, bln消费卡)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardNODencode(ByVal strCardNo As String, ByVal lng卡类别ID As Long, bln消费卡 As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡号的加密数据
    '返回:返回被加密的数据(*****123)
    '编制:刘兴洪
    '日期:2011-08-25 09:07:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlGetCardNODencode = GetCardNODencode(strCardNo, lng卡类别ID, "", bln消费卡)
End Function


Public Function zlThreeInterFaceMgrTool(ByVal frmMain As Object, ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方接口管理工具
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '编制:刘兴洪
    '日期:2012-02-13 10:17:26
    '问题:46203
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlThreeInterFaceMgrTool = mobjThirdSwap.zlThreeInterFaceMgrTool(frmMain, lngModule)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function


Public Sub SetEnabled(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置Enabled属性
    '入参:blnEnabled-true,表示自动获取;False表示不自动获取
    '编制:刘兴洪
    '日期:2012-08-22 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle
    If mobjThirdSwap Is Nothing Then Exit Sub
    Call mobjThirdSwap.SetEnabled(blnEnabled)
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Public Sub zlInitEvents(ByVal hWnd As Long, ByVal objComEvents As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置公共事件
    '入参:objComEvents-公共事件
    '编制:刘兴洪
    '日期:2012-08-22 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call mobjThirdSwap.zlInitEvents(hWnd, objComEvents)
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlGetYLCards() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别集
    '返回: 医疗卡类别的集合
    '编制:刘兴洪
    '日期:2012-08-24 14:37:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetYLCards = zlGet医疗卡类别()
End Function


Public Function zlIsMobileNo(ByVal strInput As String, Optional ByRef strRutType As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '功能:判断传入的是否为手机号
    '入参:strInput-手机号
    '出参:strRutType-查询结果:0-成功;1-不是有效号段;2-号码长度不对
    '返回:True-传入号码为手机号;False-传入号码不为手机号
    '编制:刘尔旋
    '日期:2017-1-25
    '---------------------------------------------------------------------------------------------
    zlIsMobileNo = mobjOneCardObject.zlIsMobileNo(strInput, strRutType)
End Function

Public Function zlCheckOutNoIsExist(ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查门诊号是否被他人使用
    ' 入参 : str门诊号-传入检查的门诊号
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlCheckOutNoIsExist = mobjService.zlPatiSvr_CheckOutNoIsExist(lng病人ID, str门诊号, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlCheckPhoneIsExist(ByVal lng病人ID As Long, ByVal str手机号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查手机号是否被他人使用
    ' 入参 :
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlCheckPhoneIsExist = mobjService.zlPatiSvr_PhoneNumberExist(lng病人ID, str手机号, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlCheckMCNOIsExist(ByVal str医保号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查医保号是否被他人使用
    ' 入参 :
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlCheckMCNOIsExist = mobjService.zlPatiSvr_CheckInsNoIsExist(str医保号, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function


Public Function zlUpdateOutMedRec(ByVal lng病人ID As Long, ByVal str门诊号 As String, ByVal str建立日期 As String, _
                Optional ByVal str病案类别 As String = "一般", Optional ByVal str存储状态 As String, Optional ByVal str存放位置 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 更新门诊病案，没有记录时新增记录
    ' 入参 : str病案号-门诊号
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/21 13:53
    '---------------------------------------------------------------------------------------
    Dim cllMedRec As New Collection
    
    cllMedRec.Add Array("病人ID", lng病人ID)
    cllMedRec.Add Array("门诊号", str门诊号)
    cllMedRec.Add Array("建立日期", str建立日期)
    cllMedRec.Add Array("病案类别", str病案类别)
    cllMedRec.Add Array("存储状态", str存储状态)
    cllMedRec.Add Array("存放位置", str存放位置)

    zlUpdateOutMedRec = zl_CisSvr_UpdateOutMedRecord(cllMedRec)
End Function

Public Function zlGetPatiIDFromBedNumber(ByVal lng病区ID As Long, _
    ByVal str床号 As String, Optional ByRef lng主页Id As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号获取病人ID
    '出参:lng主页ID-返回当前床号的主页ID
    '返回:成功返回病人ID,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlGetPatiIDFromBedNumber = mobjOneCardObject.zlGetPatiIDFromBedNumber(lng病区ID, str床号, lng主页Id)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function zlGetPatiIDFromObservationNo(ByVal str留观号 As String, Optional ByRef lng主页Id As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据留观号获取病人ID
    '出参:
    '   lng主页ID-返回当前床号的主页ID
    '返回:成功返回病人ID,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection
    
    On Error GoTo errHandle
    Set cllFilter = New Collection
    cllFilter.Add Array("留观号", str留观号)
    zlGetPatiIDFromObservationNo = mdlPubServerCis.zl_CisSvr_GetPatiID(cllFilter, lng主页Id)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiIdFromPatiIDCard(ByVal objControl As Object, ByVal strIDCard As String, ByVal strName As String, ByVal strSex As String, _
    ByRef lngPatiId As Long, Optional frmMain As Object, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal byt场合 As Byte, _
    Optional ByRef blnCancel As Boolean, Optional ByVal blnOnlyUnitPatient As Boolean, _
    Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人身份证号，获取病人信息
    '入参:
    '   objControl-调用的控件
    '   strName-输入的病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   byt场合-调用场合,用于控制显示列：0-门诊挂号,1-门诊收费,2-住院;3-分诊;4-医疗卡;5-入院
    '   blnAddNewPati=在按病人模糊查找时是否增加新病人
    '   blnOnlyUnitPatient-只查合约单位病人
    '出参：
    '   blnCancel=用于表示输入取消
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiIdFromPatiIDCard = mobjOneCardObject.zlGetPatiIdFromPatiIDCard(objControl, strIDCard, _
        strName, strSex, lngPatiId, frmMain, bln显示门诊预交, bln显示住院预交, _
        byt病人来源, byt场合, blnCancel, blnOnlyUnitPatient)
End Function

Public Function zlGetPatiIdFromPatiName(ByVal objControl As Object, ByVal strName As String, ByRef lngPatiId As Long, _
    Optional frmMain As Object, Optional ByVal intDefaultDays As Integer, Optional ByVal lngDefaultCardTypeID As Long, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal byt场合 As Byte, _
    Optional ByRef blnCancel As Boolean, Optional ByVal blnAddNewPati As Boolean, _
    Optional ByVal blnOnlyUnitPatient As Boolean, _
    Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息
    '入参:
    '   objControl-调用的控件
    '   strName-输入的病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   byt场合-调用场合,用于控制显示列：0-门诊挂号,1-门诊收费,2-住院
    '   blnAddNewPati=在按病人模糊查找时是否增加新病人
    '   blnOnlyUnitPatient-只查合约单位病人
    '出参：
    '   blnCancel=用于表示输入取消
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiIdFromPatiName = mobjOneCardObject.zlGetPatiIdFromPatiName(objControl, _
        strName, lngPatiId, frmMain, bln显示门诊预交, bln显示住院预交, intDefaultDays, lngDefaultCardTypeID, _
        byt病人来源, byt场合, blnCancel, blnAddNewPati, blnOnlyUnitPatient)
End Function

Public Function zlGetPatiRecordFromPatiName(ByVal strName As String, ByRef rsPati As ADODB.Recordset, _
    Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean, _
    Optional ByVal intDefaultDays As Integer, Optional ByVal lngDefaultCardTypeID As Long, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal blnOnlyUnitPatient As Boolean, _
    Optional ByVal blnLikeByName As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息记录集
    '入参:
    '   objControl-调用的控件
    '   strName-病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   blnOnlyUnitPatient-只查合约单位病人
    '   blnLikeByName-是否按姓名做匹配模糊查找
    '出参:
    '   rsPati-病人信息集,字段：排序ID,ID,病人ID,主页ID,姓名,性别,年龄,病人类型,险类,门诊号,住院号,科室,
    '                           床号,出生日期,身份证号,家庭地址,工作单位,在院标志,手机号,是否有医疗卡,就诊时间,入院日期,出院日期
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiRecordFromPatiName = mobjOneCardObject.zlGetPatiRecordFromPatiName(strName, rsPati, _
        bln显示门诊预交, bln显示住院预交, intDefaultDays, lngDefaultCardTypeID, byt病人来源, blnOnlyUnitPatient, blnLikeByName)
End Function

Public Function zlGetPatiInfoFromPatiName(ByVal objControl As Object, ByVal strName As String, ByRef rsPatiInfo_Out As ADODB.Recordset, _
    Optional frmMain As Object, Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean, _
    Optional ByVal intDefaultDays As Integer, Optional ByVal lngDefaultCardTypeID As Long, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal byt场合 As Byte, _
    Optional ByRef blnCancel As Boolean, Optional ByVal blnAddNewPati As Boolean, _
    Optional ByVal blnOnlyUnitPatient As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息详细情况
    '入参:
    '   objControl-调用的控件
    '   strName-输入的病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   byt场合-调用场合,用于控制显示列：0-门诊挂号,1-门诊收费,2-住院
    '   blnAddNewPati=在按病人模糊查找时是否增加新病人
    '   blnOnlyUnitPatient-只查合约单位病人
    '出参:rsPatiInfo_Out-返回病人信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiInfoFromPatiName = mobjOneCardObject.zlGetPatiInfoFromPatiName(objControl, _
        strName, rsPatiInfo_Out, frmMain, bln显示门诊预交, bln显示住院预交, intDefaultDays, lngDefaultCardTypeID, _
        byt病人来源, byt场合, blnCancel, blnAddNewPati, blnOnlyUnitPatient)
End Function

Public Function zlGetPatiIDFromInpatientNum(ByVal str住院号 As String, ByRef lng病人ID_Out As Long, _
    Optional ByVal blnNotShowMsg As Boolean, Optional ByRef strErrMsg_Out As String, _
    Optional lng主页ID_out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据住院号获取病人信息
    '入参:str住院号- 住院号
    '     blnNotShowMsg-是否不显示消息框
    '出参:strErrMsg_Out-返回的错误信息
    '     lng病人Id_out-返回的病人ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-29 21:12:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    zlGetPatiIDFromInpatientNum = mobjOneCardObject.zlGetPatiIDFromInpatientNum(str住院号, lng病人ID_Out, _
        blnNotShowMsg, strErrMsg_Out, lng主页ID_out)
End Function

Public Function zlGetContractUnitPatiInfor(ByRef cllPati As Collection, Optional ByVal lng合同单位ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取所有合同单位病人
    '返回:获取成功返回True,否则返回False
    '出参:
    '   cllPati=病人信息，成员：clsPatientInfo对象=cllPati(_病人ID)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetContractUnitPatiInfor = mobjOneCardObject.zlGetContractUnitPatiInfor(cllPati, lng合同单位ID)
End Function

Public Function zlGetAvailabilityWriteCardType() As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:存在写卡类别
    '出参:返回写卡类别,多个用逗号分离
    '返回:存在写卡类别的ID,如:123,232,...
    '编制:刘兴洪
    '日期:2013-06-07 10:40:59
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetAvailabilityWriteCardType = GetAvailabilityWriteCardType
End Function


Public Function zlGetCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As Card) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别给卡对象(此对象
    '入参:lngCardTypeID-卡类别ID
    '       bln消费卡-是否消费卡
    '出参:objCard-返回卡对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-25 10:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Set objCard = mobjThirdSwap.zlGetCard(lngCardTypeID, bln消费卡)
    If Not objCard Is Nothing Then zlGetCard = True: Exit Function
    zlGetCard = GetCardFromCardtypeID(lngCardTypeID, bln消费卡, objCard)
End Function
    
Public Function zlGetCardFromTypeName(ByVal strCardName As String, ByVal bln消费卡 As Boolean, ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据医疗卡名称获取卡对象
    ' 入参 : strCardName - 卡类别名称
    '        bln消费卡-是否消费卡
    ' 出参 : objCard-返回卡对象
    ' 编制 : 李南春
    ' 日期 : 2019/11/2 10:39
    '---------------------------------------------------------------------------------------
    Set objCard = mobjThirdSwap.zlGetCardFromTypeName(strCardName, bln消费卡)
    If Not objCard Is Nothing Then zlGetCardFromTypeName = True: Exit Function
End Function


Public Function zlSetBrushCardObject(ByVal lngCardTypeID As Long, _
    ByRef objBrushCard As Object, ByRef strExpend As String, Optional ByVal bln消费卡 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡对象
    '入参:lngCardTypeID-医疗卡类别ID
    '    objBrushCard -指定的刷卡对象(文本框)
    '    strExpend-扩展信息(暂无,待以后扩展用)
    '出参:strExpend-扩展信息(暂无,待以后扩展用)
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-21 14:04:13
    '说明:主要针对刷卡方式的,比如:有些银行卡加密的,需要刷卡后进行解密
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlSetBrushCardObject = mobjThirdSwap.zlSetBrushCardObject(lngCardTypeID, objBrushCard, strExpend, bln消费卡)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function zlTransferAccountsCheck(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, ByVal dblMoney As Double, ByVal strBalanceID As String, Optional ByRef strXMLExpend As String, Optional strErrMsg_Out As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:转帐检查
    '入参:
    '   frmMain-调用的主窗体
    '   lngModule-HIS调用模块号
    '   lngCardTypeID-卡类别ID
    '   strCardNo-卡号
    '   dblMoney-转帐金额(代扣时为负数)
    '   strBalanceID-本次支付结算ID，4-门诊退费业务为原结帐ID
    '   strXMLExpend-XML串:
    '       <IN>
    '           <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务
    '       </IN>
    '出参:strXMLExpend-XML串:
    '        <OUT>
    '           <ERRMSG>错误信息</ERRMSG >
    '        </OUT>
    '     strErrMsg_Out-返回的错误信息(从XML中解析出来的错误信息)
    '返回:检查的数据合法,返回True:否则返回False
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    On Error GoTo errHandle
    zlTransferAccountsCheck = mobjThirdSwap.zlTransferAccountsCheck(frmMain, lngModule, lngCardTypeID, strCardNo, dblMoney, strBalanceID, strXMLExpend, strErrMsg_Out)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlTransferAccountsMoney(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, ByVal strBalanceID As String, ByVal dblMoney As Double, Optional ByRef strSwapGlideNO As String, _
    Optional ByRef strSwapMemo As String, Optional ByRef strSwapExtendInfor As String, Optional ByRef strXMLExpend As String, _
    Optional rsExpend_Out As ADODB.Recordset, Optional strErrMsg_Out As String, Optional intStatus_Out As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方帐户转帐
    '入参:
    '   frmMain-调用的主窗体
    '   lngModule-HIS调用模块号
    '   lngCardTypeID-卡类别ID
    '   strCardNo-卡号
    '   strBalanceID-结算ID，4-门诊退费业务为原结帐ID
    '   dblMoney-转帐金额
    '    strSwapExtendInfor-退费业务时，传入本次退费的冲销ID:
    '                        格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                        收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '   strXMLExpend-XML串:
    '       <IN>
    '           <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务
    '       </IN>
    '出参:
    '   strSwapGlideNO-交易流水号
    '   strSwapMemo -交易说明
    '   strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '   strXMLExpend-XML串:
    '        <OUT>
    '           <JYZT>错误信息</JYZT>   //10.35.110增加此接点:接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '           <ERRMSG>错误信息</ERRMSG >
    '        </OUT>
    '   rsExpend_Out-返回解析后的扩展信息集(项目名称,项目内容 )
    '   strErrMsg_Out:XML返回的错误信息
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain
    zlTransferAccountsMoney = mobjThirdSwap.zlTransferAccountsMoney(frmMain, lngModule, _
     lngCardTypeID, strCardNo, strBalanceID, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend, rsExpend_Out, strErrMsg_Out, intStatus_Out)

    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function SetParentHwnd(ByVal lngfrmMainHwnd As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置父窗口句配
    '入参:lngfrmMainHwnd-父窗口句柄
    '返回:设置成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-19 14:52:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If lngfrmMainHwnd = 0 Then Exit Function
    mlngfrmMainHwnd = lngfrmMainHwnd
End Function




Public Function ZlGetParaConfig(ByVal frmMain As Object, _
    ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, ByVal intPara As Integer, _
    Optional strErrMsg As String, Optional strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取接口参数
    '入参: frmMain-调用的主窗体
    '       intPara: 包含如下值
    '                1-刷卡和支付在同一页面:true-新模式；False-旧模式
    '                2-作废时是否单独调用退费接口
    '                参数值          参数名                                 应用场合                        参数说明
    '                1               刷卡和支付在同一页面                   自助系统                        主要应用在自助系统中，可以通过此参数，调整界面的一些布局。
    '                2               按交易流水号分别调用退款接口           病人结帐,
    '                                                                       补充结算 (含自助补允结算)
    '                                                                       病人预交款管理(余额退款)        主要是在退款时，如果此参数返回True,则调用zlReturnMoney接口.否则调用returnMultMoney接口
    '
    '                3               自助系统调用刷卡接口                   中联自助系统(10.35.10后支持)    该参数主要控制中联自助系统是否在支付时，调用zlBushCard接口
    '                4               禁止门诊线下医保或多种结算方式收款     门诊部分退费,10.35.110后支持    主要是应用于一卡通在线上支付含医保接口或多种结算方式时，在线下进行部分退时，不支持全退后重收
    '                                                                                                         返回:True时，线下本一卡通不支持持医保结算或多种结算方式;False时，表示允许使用本一卡通进行线下医保结算或多种结算方式。
    '                6               是否支持押金缴款                  预交款管理(10.35.130)     如果返回true,表示在病人预交管理中的押金部分支持使用该种类别进行缴款操作，否则禁止缴款
    '       strExpend-扩展参数，暂留，现传为空
    '出参:strErrMsg-返回的错误信息(10.35.110后启用)
    '返回:函数返回True:调用成功,False:调用失败
   '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    ZlGetParaConfig = mobjThirdSwap.ZlGetParaConfig(frmMain, lng卡类别ID, bln消费卡, intPara, strErrMsg, strExpend)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlGetCardNo(ByRef strCardNo As String, Optional strExpend As String) As Boolean
    '获取刷卡卡号
    '出参:strCardNo-刷卡卡号
    '     strExpend-扩展参数，暂留，现传为空
    On Error GoTo errHandle
    ZlGetCardNo = mobjThirdSwap.ZlGetCardNo(strCardNo, strExpend)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlSwapIsSucces(ByVal frmMain As Object, ByVal lngModule As Long, ByVal intSwapType As Integer, _
    ByVal lngCardTypeID As Long, ByVal strBalanceIDs As String, ByVal strCardNo As String, ByVal dblSwapMoney As Double, _
    ByRef intSwapStatus_Out As Integer, Optional strErrMsg As String, Optional strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断交易是否成功（10.35.110）
    '入参:  frmMain-调用的主窗体
    '       lngModule-模块号
    '       intSwapType-0-扣款;1-退款；2-转帐
    '       lngCardTypeID-卡类别ID
    '       strCardNO-卡号
    '       dblSwapMoney-交易金额
    '       strBalanceIDs-本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn 收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
    '       strExpend-扩展参数:退款或作废时，才传入,格式如下 ：
    '        <INPUT>
    '            <TKLIST>
    '                    <TK>
    '                       <JYLSH>交易流水号</JYLSH>
    '                       <KH>卡号</KH>
    '                       <JE>金额</JE>
    '                    </TK>
    '            </TKLIST>
    '        </INPUT>
    '出参:intSwapStatus_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
    '     strXMLExpend-待以后扩展
    '返回：接口调用成功返回true,否则返回Flase
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle:
    Set mfrmMain = frmMain
    zlSwapIsSucces = mobjThirdSwap.zlSwapIsSucces(frmMain, lngModule, intSwapType, lngCardTypeID, strBalanceIDs, strCardNo, dblSwapMoney, intSwapStatus_Out, strErrMsg, strXMLExpend)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiInforFromPatiID(ByVal lng病人ID As Long, ByRef objPati_Out As clsPatientInfo, _
    Optional ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String, _
    Optional ByVal blnLoadFamily As Boolean, Optional ByVal blnLoadDrug As Boolean, _
    Optional ByVal blnLoadImmune As Boolean, Optional ByVal blnLoadCard As Boolean, _
    Optional cllFamily As Collection, Optional cllDrug As Collection, _
    Optional cllImmune As Collection, Optional ByVal blnNotShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID获取病人信息集
    '入参:blnLoadFamily-是否获取病人家属信息
    '     blnLoadDrug-是否获取病人过敏药物信息
    '     blnLoadImmune-是否获取病人家属信息
    '     blnLoadCard-仅占位，目前不用
    '出参:objPati_Out-返回病人信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-01 17:04:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiInforFromPatiID = mobjOneCardObject.zlGetPatiInforFromPatiID(lng病人ID, objPati_Out, strErrMsg, _
        strOtherName, strOtherValue, blnLoadFamily, blnLoadDrug, blnLoadImmune, blnLoadCard, cllFamily, cllDrug, cllImmune, blnNotShowErrMsg)
End Function



Public Function zlPatiSvr_GetPatiExtendInfo(ByVal lng病人ID As Long, ByVal str信息名集 As String, ByRef cllExtends_Out As Collection, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg As String, Optional ByVal lng就诊ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息从表信息服务接口
    '入参:str信息名集-多个用逗号分离,如：医学警示,联系人2,联系人3等
    '
    '出参:cllPatiData_Out-返回病人从表信息数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 20:10:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetPatiExtendInfo = mobjOneCardObject.zlPatiSvr_GetPatiExtendInfo(lng病人ID, str信息名集, cllExtends_Out, blnNotShowErrMsg, strErrMsg, lng就诊ID)
End Function

Public Function zlGetMultiPatiInforFromPatiID(ByVal str病人Ids As String, ByRef cllPatis_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID一次获取多个病人信息集
    '入参:
    '   str病人IDs=多个病人ID用英文逗号分隔
    '出参:
    '   cllPatis_Out-返回病人信息集,成员：clsPatinetInfo=cllPatis_Out(_病人ID)
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetMultiPatiInforFromPatiID = mobjOneCardObject.zlGetMultiPatiInforFromPatiID(str病人Ids, cllPatis_Out)
End Function

Public Function zlGetPatiOtherInforFromPatiID(ByVal lng病人ID As Long, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg As String, _
    Optional ByVal blnLoadFamily As Boolean, Optional ByVal blnLoadDrug As Boolean, _
    Optional ByVal blnLoadImmune As Boolean, Optional ByVal blnLoadCert As Boolean, _
    Optional ByVal blnLoadOther As Boolean, _
    Optional cllFamily As Collection, Optional cllDrug As Collection, _
    Optional cllImmune As Collection, Optional cllCert As Collection, _
    Optional cllOther As Collection, Optional cllContact As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据病人ID获取证件、免疫、过敏、联系人及其他信息
    ' 入参 :
    '   blnLoadFamily-是否获取家属关系信息
    '   blnLoadDrug-是否获取过敏信息
    '   blnLoadImmune-是否获取免疫记录
    '   blnLoadCert-是否获取证件信息
    '   blnLoadOther-是否获取其他从属信息，包括联系人信息及学习，RH等
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/8 14:25
    '---------------------------------------------------------------------------------------
    zlGetPatiOtherInforFromPatiID = mobjOneCardObject.zlGetPatiOtherInforFromPatiID(lng病人ID, blnNotShowErrMsg, strErrMsg, _
            blnLoadFamily, blnLoadDrug, blnLoadImmune, blnLoadCert, blnLoadOther, cllFamily, cllDrug, _
            cllImmune, cllCert, cllOther, cllContact)
End Function

Public Function zlGetPatiSurplusFromPatiID(ByVal lng病人ID As Long, Optional ByRef dbl门诊预交余额_out As Double, Optional ByRef dbl住院预交余额_Out As Double, _
    Optional ByRef dbl门诊费用余额_Out As Double, Optional ByRef dbl住院费用余额_Out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID获取病人费用及预交余额
    '入参:lng病人ID-病人ID
    '
    '出参:dbl门诊预交余额_out
    '     dbl住院预交余额_Out
    '     dbl门诊费用余额_Out
    '     dbl住院费用余额_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-05 20:57:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiSurplusFromPatiID = GetPatiSurplusFromPatiID(lng病人ID, dbl门诊预交余额_out, dbl住院预交余额_Out, dbl门诊费用余额_Out, dbl住院费用余额_Out)
End Function

Public Function zlGetPatiCertInfo(ByVal lng病人ID As Long, ByRef cllCert As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据病人ID获取病人有效的证件信息
    ' 入参 :
    ' 出参 : cllCert-证件列表
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/8 20:00
    '---------------------------------------------------------------------------------------
    zlGetPatiCertInfo = objOneCardObject.zlGetPatiCertInfo(lng病人ID, cllCert)
End Function

Public Function zlGetPatiIDByComminuty(ByVal int社区序号 As Integer, ByVal str社区号 As String, _
                ByRef lng病人ID As Long, Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据社区号查找病人ID
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/15 17:50
    '---------------------------------------------------------------------------------------
    zlGetPatiIDByComminuty = mobjService.zlPatiSvr_GetPatiIDByComminuty(int社区序号, str社区号, lng病人ID, blnNotShowErrMsg, strErrMsg)
End Function

Public Function ZlGetPatiInsureInfo(ByVal lng病人ID As Long, ByRef insure As Integer, Optional insureName As String, _
                                                       Optional str医保号 As String, Optional str医保密码 As String, Optional str卡号 As String, _
                                                       Optional str登记时间 As String, Optional lng病种id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id和险类，获取病人的保险信息
    '入参: insure :险类
    '返回:返回病人医保信息
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZlGetPatiInsureInfo = mobjOneCardObject.ZlGetPatiInsureInfo(lng病人ID, insure, insureName, str医保号, str医保密码, str卡号, str登记时间, lng病种id)
End Function

Public Function zlPatiIsInhospital(ByVal lng病人ID As Long, ByRef blnInhospital As Boolean, _
                Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查病人是否在院就诊
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 14:35
    '---------------------------------------------------------------------------------------
    zlPatiIsInhospital = mobjService.zlCisSvr_PatiIsInhospital(lng病人ID, blnInhospital, blnNotShowErrMsg, strErrMsg)
End Function


Public Function zlCheckMzLgPati(ByVal lng病人ID As Long, ByVal lng主页Id As Long, ByRef blnMzLgPati As Boolean, _
                Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查病人是不是门诊留观病人
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 19:07
    '---------------------------------------------------------------------------------------
    Dim cllPati As Collection
    On Error GoTo errHand

    If mobjService.zlCIsSvr_GetPatiPageInfo(0, lng病人ID & ":" & lng主页Id, cllPati, , , , blnNotShowErrMsg, strErrMsg_Out) = False Then Exit Function
    If cllPati Is Nothing Then zlCheckMzLgPati = True: Exit Function
    If cllPati.count = 0 Then zlCheckMzLgPati = True: Exit Function
    blnMzLgPati = Val(NVL(cllPati(1)("_pati_nature"))) = 1
    zlCheckMzLgPati = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlExistAdvice(ByVal lng病人ID As Long, ByVal str挂号单 As String, ByRef blnHavAdvice As Boolean, _
                Optional ByVal lng主页Id As Long, Optional ByVal blnOnlyValid As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查挂号单是否发生了医嘱
    ' 入参 : str挂号单-多个单据号间用逗号分隔
    '        blnOnlyValid-是否只检查有效医嘱
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 19:48
    '---------------------------------------------------------------------------------------
    zlExistAdvice = zl_Cissvr_Existadvice(lng病人ID, str挂号单, blnHavAdvice, lng主页Id, blnOnlyValid)
End Function
    


Public Function zlGetPatiInforFromXML(ByVal strPatiXml As String, ByRef strErrMsg_Out As String, ByRef objPatiInfor_Out As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从XML中获取病人信息
    '入参:strPatiXml-病人信息XML
    '
    '出参:objPatiInfor_Out-返回病人信息对象集
    '      strErrMsg_Out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 14:29:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiInforFromXML = mobjOneCardObject.zlGetPatiInforFromXML(strPatiXml, strErrMsg_Out, objPatiInfor_Out)
End Function


Public Function zlGetPatiIdsByRange(ByVal strCondition As String, ByRef strPatiIds As String, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String, _
    Optional ByVal blnFindByFilter As Boolean, Optional ByVal cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件值获取符合条件的病人ID
    '入参:
    '   strCondition=可能是就诊卡号、身份证号、IC卡号、门诊号
    '   blnFindByFilter=True:按过滤条件(cllFilter)获取;False:按strCondition获取
    '   cllFilter=过滤条件:Array(Key,Value),Key:合同单位ID
    '出参:
    '返回:执行成功返回True，失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    zlGetPatiIdsByRange = mobjService.zlPatiSvr_GetPatiIdsByRange(strCondition, strPatiIds, _
        blnNotShowErrMsg, strErrMsg_Out, blnFindByFilter, cllFilter)
End Function
    

Public Function ErrCenter() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:错误处理中心
    '编制:刘兴洪
    '日期:2018-12-05 11:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
   ErrCenter = mobjDataBase.ErrCenter
End Function

Public Sub SaveErrLog()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写日志
    '编制:刘兴洪
    '日期:2018-12-05 11:33:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Sub
    End If
    Call mobjDataBase.SaveErrLog
End Sub
Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True, Optional ByVal lngDeptID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngSys=使用该参数的系统编号，如100
    '      lngModual=使用该参数的模块号，如1230
    '      blnSetup=调用模块是否有参数设置权限
    '      lngDeptID=部门级参数
    '返回：设置是否成功
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    SetPara = mobjDataBase.SetPara(varPara, strValue, glngSys, lngModual, blnSetup, lngDeptID)
End Function
Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer, Optional ByVal lngDeptID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：读取指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      lngSys=使用该参数的系统编号，如100
    '      lngModual=使用该参数的模块号，如1230
    '      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
    '      blnNotCache=是否不从缓存中读取
    '      arrControl=控件数组，如Array(Me.Text1, Me.CheckBox1)，用于函数内部自动处理对应控件的显示颜色，是否禁止设置。
    '      blnSetup=调用模块是否有参数设置权限
    '      intType=返回参数，返回参数类型
    '      lngDeptID=部门级参数，根据部门读取
    '返回：参数值，字符串形式
    '编制:刘兴洪
    '日期:2018-12-05 14:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    GetPara = mobjDataBase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType, lngDeptID)
End Function

Public Function GetPrivFunc(lngSys As Long, lngProgId As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：返回当前用户具有的指定程序的功能串
    '参数：lngSys     如果是固定模块，则为0
    '      lngProgId  程序序号
    '返回：分号间隔的功能串,为空表示没有权限
    '编制:刘兴洪
    '日期:2018-12-05 15:01:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    GetPrivFunc = mobjDataBase.GetPrivFunc(lngSys, lngProgId)
End Function

Public Sub WritDebugLog(ByVal strDev As String, strInput As String, strOutPut As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记录日志
    '编制:刘兴洪
    '日期:2018-12-05 11:35:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Sub
    End If
    Call mobjDataBase.WritLog(strDev, strInput, strOutPut)
End Sub


Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    glngInstanceCount = glngInstanceCount + 1
    Call zlInitCommLib  '初始化CommLib对象
    Set mobjService = New clsService
End Sub

Private Sub Class_Terminate()

    Err = 0: On Error Resume Next
    Set mobjService = Nothing: Set mcnOracle = Nothing
    Set mobjOneCardObject = Nothing: Set mobjDataBase = Nothing
    Set mobjThirdSwap = Nothing
    Set mfrmMain = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    
    Call zlReleaseResources '释放资源
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
End Function

Public Function zlGetOneCardTypes(ByRef rsTypes_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新获取所有的医疗卡类别信息集
    '入参:objPati-病人信息集
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-21 15:15:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetOneCardTypes = GetOneCardTypes(rsTypes_Out)  '获取医疗卡类别信息集
End Function
 
Public Function zlGetConsumerCardTypes(ByRef rsSquareType_Out As ADODB.Recordset, Optional ByVal blnOnlyStart As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新获取消费卡类别集
    '入参:blnOnlyStart-是否只包含启用的
    '
    '出参:rsSquareType_Out-消费卡目录
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-21 15:24:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetConsumerCardTypes = GetConsumerCardTypes(rsSquareType_Out, blnOnlyStart)
End Function
Public Function zlPubPatiSel(ByVal frmMain As Object, ByRef lng病人ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人公共选择器
    '入参:frmMain-调用的主窗体
    '出参:lng病人ID_Out-返回病人ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-21 15:44:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmPatiSelNew  As New frmPatiSel
    zlPubPatiSel = frmPatiSelNew.zlShowCard(mcnOracle, frmMain, "", lng病人ID_Out)
    If Not frmPatiSelNew Is Nothing Then Unload frmPatiSelNew
    Set frmPatiSelNew = Nothing
End Function

Public Function zlCardTypeSelect(ByVal frmMain As Object, _
    ByVal strCardTypeIDs As String, ByRef lngCardTypeID As Long, Optional strFromCaption As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人公共选择器
    '入参:strCardTypeIDs-空,代表所有医疗卡;否则指定的医疗卡
    '       strFromCaption-重新命名的窗体标题
    '出参:lngCardTypeID-当前选择的医疗卡
    '返回:选择成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-21 15:44:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTypeSel   As New frmSelectType
    zlCardTypeSelect = objTypeSel.zlSelect(frmMain, strCardTypeIDs, lngCardTypeID, strFromCaption, mcnOracle)
    If Not objTypeSel Is Nothing Then Unload objTypeSel
    Set objTypeSel = Nothing
End Function


Public Function zlIsExistFeeInsurePatient(lng病人ID As Long, Optional blnExist_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医保病人是否存在未结费用
    '入参:lng病人ID
    '出参:blnExist_out-存在未结费用，返回true,否则返回False
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 16:40:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
     zlIsExistFeeInsurePatient = mobjOneCardObject.zlIsExistFeeInsurePatient(lng病人ID, blnExist_Out)
End Function

Public Function zlGetPatiInfoRecordsFromCons(ByVal lng病人ID As Long, cllFilter As Collection, _
    ByRef rsPatiInfo_Out As ADODB.Recordset, Optional blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件获取病人信息数据给记录集
    '入参:cllFilter-病人信息集(array(查询名称,查询值)),如:array("卡号","A00001")
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-21 19:12:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    zlGetPatiInfoRecordsFromCons = mobjOneCardObject.zlGetPatiInfoRecordsFromCons(lng病人ID, cllFilter, rsPatiInfo_Out, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlGetInputItemLength(ByVal strTableItems As String, ByRef cllColumn As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    '功能 : 获取病人信息中指定字段的最大长度
    '入参 :
    '   strTableItems：表1:列1,列2|表2:列1,列2,列3|..
    '出参 :
    '   cllColumn(Collect):成员(表名,列名,列长度)=cllcolumn(表名_列名)
    '返回 :
    '编制 : 李南春
    '日期 : 2019/11/21 20:16
    '---------------------------------------------------------------------------------------
    zlGetInputItemLength = mobjOneCardObject.zlGetInputItemLength(strTableItems, cllColumn)
End Function

Public Function zlGetPatiVitalSigns(ByVal lng病人ID As Long, ByVal lng挂号ID As String, _
                ByRef cllVital As Collection, Optional ByVal blnOutPati As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人生命体征信息
    ' 入参 : blnOutPati-门诊病人
    ' 出参 : cllVital:体征信息(Collect)(项目,值,单位)
    ' 返回 : 返回病人的体征信息，包括项目，数值，单位
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 19:48
    '---------------------------------------------------------------------------------------
    zlGetPatiVitalSigns = mobjOneCardObject.zlGetPatiVitalSigns(lng病人ID, lng挂号ID, cllVital, blnOutPati)
End Function

Public Function Zl_Calc_Age(ByVal lng病人ID As Long, ByVal str出生日期 As String, Optional ByVal str计算日期 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据出生日期计算病人年龄
    '入参:str计算日期-指定的计算年龄的日期
    '出参:
    '返回:计算的年龄
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Zl_Calc_Age = zl_Patisvr_Calc_Age(lng病人ID, str出生日期, str计算日期)
End Function

Public Function zlGetPatiAddrssInfo(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
                ByVal str地址类别 As String, ByRef cllAddrList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人结构化地址信息
    ' 入参 : int地址类别:查询的地址类别：1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址；为0时表示查询所有类型的地址信息
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlGetPatiAddrssInfo = mobjService.zlPatiSvr_GetPatiAddrssInfo(lng病人ID, lng主页Id, str地址类别, cllAddrList)
End Function

Public Function zlGetPatiBaseInfoByPatiId(ByVal lng病人ID As Long, Optional ByVal lng主页Id As Long = -1, _
                Optional ByRef cllPatiList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人结构化地址信息
    ' 入参 : lng病人ID；
    '        lng主页ID :-1表示返回所有的住院信息
    ' 出参 : cllPati(Collect):成员()
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlGetPatiBaseInfoByPatiId = mobjService.ZlCissvr_GetPatiBaseInfoByPatiId(lng病人ID, lng主页Id, cllPatiList)
End Function


Public Function zlGetPatiPhoto(ByVal lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人照片
    '入参:
    '出参:
    '返回:返回病人照片Base64
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPatiPhoto As String
    strPatiPhoto = zl_Patisvr_GetPatiPhoto(lng病人ID)
    If strPatiPhoto = "" Then Exit Function
    zlGetPatiPhoto = gobjComLib.zlStr.DecodeBase64_File(strPatiPhoto)
End Function

Public Function zlSavePatiPhoto(ByVal lng病人ID As Long, ByVal strPatiPhoto As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:
    '出参:
    '返回:返回病人照片Base64
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSavePatiPhoto = zl_Patisvr_SavePatiPhoto(lng病人ID, strPatiPhoto)
End Function

Public Function zlDeletePatiPhoto(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除病人照片
    '入参:
    '出参:
    '返回:
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlDeletePatiPhoto = zl_Patisvr_DeletePatiPhoto(lng病人ID)
End Function

Public Function zlGetInpatiState(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
                Optional ByVal intPatiType As Integer, _
                Optional ByRef cllPati_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人住院状态
    '入参:lng病人ID；lng主页ID
    '     intPatiType:病人性质 0-普通住院病人 1-门诊留观病人 2-住院留观病人
    '返回:cllPati_Out 病人信息
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetInpatiState = zl_Cissvr_GetInpatiState(lng病人ID, lng主页Id, intPatiType, cllPati_Out)
End Function

