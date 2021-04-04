VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThreeSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mfrmMain As Object
Private mlngModule As Long
Private mlng病人ID As Long, mstrPatiName As String, mstrSex As String, mstrOld As String
Private mobjOneCardComLib As zlOneCardComLib.clsOneCardComLib
Private mstrOnCardPrivs As String
Private mbyt刷卡缺省金额操作 As Byte, mbln结算异步控制 As Boolean

Public Function Init(objOneCardComLib As zlOneCardComLib.clsOneCardComLib, _
    frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal lng病人ID As Long, Optional ByVal strPatiName As String, _
    Optional ByVal strSex As String, Optional ByVal strOld As String) As Boolean
    '初始化类模块变量
    On Error GoTo ErrHandler
    Set mfrmMain = frmMain
    mlngModule = lngModule
      
    mlng病人ID = lng病人ID
    mstrPatiName = strPatiName: mstrSex = strSex: mstrOld = strOld
    Set mobjOneCardComLib = objOneCardComLib
    
    mbln结算异步控制 = Val(zlDatabase.GetPara("费用结算异步控制", glngSys)) = 1
    mbyt刷卡缺省金额操作 = Val(zlDatabase.GetPara("刷卡缺省金额操作", glngSys, 1151, "0"))  '86853
    
    Init = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLSetPatientInfo(ByVal lng病人ID As Long, ByVal strPatiName As String, _
    ByVal strSex As String, ByVal strOld As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化病人信息
    '入参:
    '   mlng病人ID=病人ID
    '   strPatiName、strSex、strOld=病人姓名、性别、年龄
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    mlng病人ID = lng病人ID
    mstrPatiName = strPatiName: mstrSex = strSex: mstrOld = strOld
    ZLSetPatientInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPayValid(ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str结算方式 As String, ByRef dblMoney As Double, ByVal strNOs As String, _
    Optional ByRef strCardNo As String, Optional ByRef strPassWord As String, _
    Optional ByRef dbl帐户余额 As Double, Optional ByRef str新结算方式 As String, _
    Optional ByVal rsClassMoney As ADODB.Recordset, Optional ByVal str费用来源 As String, _
    Optional ByRef cllSquareBalance As Collection, Optional ByVal strQRCode As String, _
    Optional ByVal dblCardFaceMoney As Double, Optional ByVal dblDepositMoney As Double, _
    Optional ByVal blnAllPay As Boolean = True) As Boolean
    '功能:三方支付交易检查
    '入参:
    '   str结算方式-当前结算方式
    '   dblMoney-支付金额
    '   strNos-本次支付所涉及的单据
    '   rsClassMoney-费用类别明细(使用消费卡支付时传入)
    '   str费用来源-当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '   strQRCode-二维码支付码
    '   dblCardFaceMoney 消费卡面值金额
    '   dblDepositMoney 消费卡充值金额
    '   blnAllPay 使用消费卡时是否费用全支付，true-费用未支付完不能完成结算，false-可以只支付部分并返回
    '出参：
    '   strCardNo-刷卡卡号
    '   strPassWord-卡号密码
    '   dbl帐户余额-帐户余额
    '   str新结算方式-刷卡接口返回的结算方式及金额，格式：结算方式|结算金额
    '   cllSquareBalance- 费卡支付明细
    '返回:交易合法返回true,否则返回False
    Dim strXMLExpend As String
    Dim str结算方式_Out As String, dbl结算金额_Out As Double
    Dim strExpand As String

    On Error GoTo ErrHandler
    str新结算方式 = ""
    
    strXMLExpend = ""
    strXMLExpend = strXMLExpend & "  <CZLX>" & 0 & "</CZLX>" & vbCrLf
    If strQRCode <> "" Then strXMLExpend = strXMLExpend & "  <QRCODE>" & strQRCode & "</QRCODE>" & vbCrLf
    If strXMLExpend <> "" Then strXMLExpend = "<IN>" & vbCrLf & strXMLExpend & "</IN>"
    
    'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowclsPatientInfo As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXMLIn As String = "", _
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
        '           <QRCODE></QRCODE> //识别的支付码，10.35.130后支持该节点,由zlReadQRCode.clsReadQRCode
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
         
    If mobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, rsClassMoney, lng卡类别ID, bln消费卡, _
        mstrPatiName, mstrSex, mstrOld, dblMoney, strCardNo, strPassWord, False, True, False, False, _
        cllSquareBalance, False, blnAllPay, strXMLExpend, str费用来源, mlng病人ID, _
        str结算方式_Out, dbl结算金额_Out) = False Then Exit Function
    
    If str结算方式_Out <> "" Then
        If RoundEx(dblMoney, 6) <> Round(dbl结算金额_Out, 6) Then
            MsgBox str结算方式 & " 实际刷卡支付金额(" & Format(dbl结算金额_Out, "0.00") & ")" & _
                "与应付金额(" & Format(dblMoney, "0.00") & ")不等，支付失败！", vbInformation, gstrSysName
            Exit Function
        End If
        str新结算方式 = str结算方式_Out & "|" & dbl结算金额_Out
    End If
    
    strXMLExpend = ""
    If dblCardFaceMoney <> 0 Then strXMLExpend = strXMLExpend & "  <MZXSJE>" & dblCardFaceMoney & "</MZXSJE>" & vbCrLf
    If dblDepositMoney <> 0 Then strXMLExpend = strXMLExpend & "  <CZJKJE>" & dblDepositMoney & "</CZJKJE>" & vbCrLf
    If strQRCode <> "" Then strXMLExpend = strXMLExpend & "  <QRCODE>" & strQRCode & "</QRCODE>" & vbCrLf
    If strXMLExpend <> "" Then strXMLExpend = "<IN>" & vbCrLf & strXMLExpend & "</IN>"
    
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
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
        '       strXMLExpend-如下两种格式：
        '           第一种:XML串: 验证密码: 自助机用
        '           <Data>
        '               <验证密码></验证密码>
        '           </Data>
        '           第二种:消费卡收款及扫码付时，传入XML串：
        '           <IN>
        '               <MZXSJE>面值销售金额</MZXSJE>
        '               <CZJKJE>充值缴款金额</CZJKJE>
        '               <QRCODE>识别的支付码</QRCODE> //10.35.130后支持
        '           </IN>
        '出参:
        '   strXMLExpend-(XML串:错误信息)
        '返回:扣款合法,返回true,否则返回Flase
        '编制:刘兴洪
        '日期:2011-05-26 16:42:43
        '说明:
        '   在调用扣款前，由于存在Oracle事务问题， 所以再调用扣款交易前， _
        '   先进行数据的合法性检查,以便控制死锁情况。
    If mobjOneCardComLib.zlPaymentCheck(mfrmMain, mlngModule, lng卡类别ID, bln消费卡, _
        strCardNo, dblMoney, strNOs, strXMLExpend) = False Then Exit Function
    
    If bln消费卡 = False Then
        'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
            strExpand As String, dblMoney As Double, _
            Optional bln消费卡 As Boolean = False) As Boolean
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
        If mobjOneCardComLib.zlGetAccountMoney(mfrmMain, mlngModule, _
            lng卡类别ID, strCardNo, strExpand, dbl帐户余额, bln消费卡) = False Then Exit Function
        If dbl帐户余额 <> 0 And dbl帐户余额 < dblMoney Then
            MsgBox str结算方式 & " 帐户余额不足！", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    CheckPayValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckDelValid(ByVal lng卡类别ID As Long, ByVal dblMoney As Double, ByVal str原结帐ID As String, _
    Optional ByRef strCardNo As String, Optional ByVal str交易流水号 As String, Optional ByVal str交易说明 As String, _
    Optional ByVal strXMLExpend As String, Optional ByVal bln是否退款验卡 As Boolean, _
    Optional ByVal byt收费类型 As Byte = 3) As Boolean
    '功能:三方退款交易检查,不含消费卡
    '入参:
    '   dblMoney-退款金额
    '   byt收费类型 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡管理退费业务
    '返回:交易合法返回true,否则返回False
    Dim strBalanceIDs As String, strPassWord As String
    Dim strExpend As String
    
    On Error GoTo ErrHandler
    'zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
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
        '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡管理退费业务
        '       dblMoney-退款金额
        '       strSwapNo-交易流水号(退款时检查),保险补充结算时传入空
        '       strSwapMemo-交易说明(退款时传入),保险补充结算时传入空
        '       strXMLExpend    XML IN
        '        <TFDATA>   //退费数据
        '            <YCTF>异常退费标志<YCTF> //1-异常重退;0-退费此节点可能没传入
        '            <TFLIST>  //退费列表
        '                <NO></NO>  // 退费单据
        '                <TFITEM>     //退费项
        '                    <SerialNum>序号</SerialNum>
        '                    ….
        '                </ TFITEM >
        '            </TFLIST>
        '
        '            <TKLIST>   //退款列表（35.90以前无此内容）
        '                <TK>
        '                    <TKFS>退款方式</TKFS>// Varchar2    20
        '                    <TKJE>退款金额</TKJE>//NUMBER
        '                    <JYLSH>原交易流水号</JYLSH>//   Varchar2    50
        '                    <JYSM><原交易说明</JYSM>//  Varhcar2    500
        '                    <DJH>单据号</DJH> //    Varchar2    8
        '                </TK>
        '                ....
        '            </TKLIST>
        '        </TFDATA>
        '返回:退款合法,返回true,否则返回Flase
        '说明:
        '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
        '    以便控制死锁情况。
    strBalanceIDs = byt收费类型 & "|" & str原结帐ID
    If mobjOneCardComLib.zlReturncheck(mfrmMain, mlngModule, lng卡类别ID, False, strCardNo, _
        strBalanceIDs, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then Exit Function

    If bln是否退款验卡 Then
       'zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln消费卡 As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByRef dbl金额 As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln退费 As Boolean = False, _
            Optional ByRef blnShowclsPatientInfo As Boolean = False, _
            Optional ByRef bln退现 As Boolean = False, _
            Optional ByVal bln余额不足禁止 As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal bln转预交 As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXMLIn As String = "", _
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
        strExpend = "<IN><CZLX>2</CZLX></IN>"
        If mobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, lng卡类别ID, False, _
            mstrPatiName, mstrSex, mstrOld, dblMoney, strCardNo, strPassWord, True, True, False, True, _
            Nothing, False, False, strExpend) = False Then Exit Function
    End If
    
    CheckDelValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckTransferValid(ByVal lng卡类别ID As Long, _
    ByVal str结算方式 As String, ByVal dblMoney As Double, _
    ByVal str原结帐ID As String, ByVal byt操作类型 As Byte, _
    Optional ByRef strCardNo As String, Optional ByRef strPassWord As String, _
    Optional ByRef str新结算方式 As String) As Boolean
    '功能:三方退款转账交易检查,不含消费卡
    '入参:
    '   dblMoney-转账金额
    '   byt操作类型- 0-补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务；5-消费卡管理退费业务
    '   str原结帐ID - 补结算时为结算序号，其它为结帐ID
    '返回:交易合法返回true,否则返回False
    Dim strXMLExpend As String
    Dim str结算方式_Out As String, dbl结算金额_Out As Double
    
    On Error GoTo ErrHandler
    str新结算方式 = ""
    'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowclsPatientInfo As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXMLIn As String = "", _
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
    strXMLExpend = "<IN><CZLX>1</CZLX></IN>"
    If mobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, lng卡类别ID, False, _
        mstrPatiName, mstrSex, mstrOld, dblMoney, strCardNo, strPassWord, True, True, False, True, _
        Nothing, False, False, strXMLExpend, "", 0, str结算方式_Out, dbl结算金额_Out) = False Then Exit Function

    If str结算方式_Out <> "" Then
        If RoundEx(dblMoney, 6) <> Round(dbl结算金额_Out, 6) Then
            MsgBox str结算方式 & " 实际转账金额(" & Format(dbl结算金额_Out, "0.00") & ")" & _
                "与应转金额(" & Format(dblMoney, "0.00") & ")不等，转账失败！", vbInformation, gstrSysName
            Exit Function
        End If
        str新结算方式 = str结算方式_Out & "|" & -1 * dbl结算金额_Out
    End If

    'zlTransferAccountsCheck(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
        ByVal strCardNo As String, ByVal dblMoney As Double, ByVal strBalanceID As String, _
        Optional ByRef strXMLExpend As String) As Boolean
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
        '           <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务；5-消费卡管理退费业务
        '       </IN>
        '出参:strXMLExpend-XML串:
        '        <OUT>
        '           <ERRMSG>错误信息</ERRMSG >
        '        </OUT>
        '返回:检查的数据合法,返回True:否则返回False
        '调用者:医保补充结算(结算时调用)
        '说明:
        '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
        '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    strXMLExpend = "<IN><CZLX>" & byt操作类型 & "</CZLX></IN>"
    If mobjOneCardComLib.zlTransferAccountsCheck(mfrmMain, mlngModule, lng卡类别ID, _
        strCardNo, dblMoney, str原结帐ID, strXMLExpend) = False Then
        Call ZlShowErrMsg(0, strXMLExpend)
        Exit Function
    End If
    
    CheckTransferValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ZlShowErrMsg(ByVal bytTYPE As Byte, ByVal strXMLErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方转账检查与代扣业务出错提示
    '参数:
    '   bytType:0-转账检查,1-转账交易
    '   strXMLErrMsg:格式如下
    '            <OUT>
    '               <ERRMSG>错误信息</ERRMSG >
    '            </OUT>
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    '解析错误信息
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '提示错误信息
    If Trim(strValue) = "" Then
        If bytTYPE = 0 Then
            strValue = vbCrLf & "交易检查失败！"
        Else
            strValue = vbCrLf & "交易失败！"
        End If
    End If
    MsgBox strValue, vbExclamation + vbOKOnly, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function ExecutePay(ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, ByRef strExpend As String, _
    ByRef dblOutMoney As Double, ByRef cll结算方式 As Collection, ByRef bln正在交易 As Boolean, _
    ByRef strErrMsg As String, Optional ByVal strQRCode As String, _
    Optional ByVal dblCardFaceMoney As Double, Optional ByVal dblDepositMoney As Double) As Boolean
    '一卡通支付
    '入参：
    '   strQRCode-二维码支付码
    '   dblCardFaceMoney 消费卡面值金额
    '   dblDepositMoney 消费卡充值金额
    Dim strExpendXML As String
    
    On Error GoTo ErrHandler
    Set cll结算方式 = Nothing
    
    strExpendXML = ""
    If dblCardFaceMoney <> 0 Then strExpendXML = strExpendXML & "  <MZXSJE>" & dblCardFaceMoney & "</MZXSJE>" & vbCrLf
    If dblDepositMoney <> 0 Then strExpendXML = strExpendXML & "  <CZJKJE>" & dblDepositMoney & "</CZJKJE>" & vbCrLf
    If strQRCode <> "" Then strExpendXML = strExpendXML & "  <QRCODE>" & strQRCode & "</QRCODE>" & vbCrLf
    If strExpendXML <> "" Then strExpendXML = "<IN>" & vbCrLf & strExpendXML & "</IN>"
    
    'zlPaymentMoney(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        byval  strPrepayNos as string , _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, _
        ByRef strSwapMemo As String, _
        Optional ByRef strExpendXML As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:帐户扣款交易
        '入参:frmMain-调用的主窗体
        '        lngModule-调用模块号
        '        strBalanceIDs-结帐ID,多个用逗号分离
        '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
        '       strCardNo-卡号
        '       dblMoney-支付金额
        '     strExpendXML，分两种情况或两种格式
        '       1.验证密码: 自助机用
        '           <Data>
        '               <验证密码></验证密码>
        '           </Data>
        '       2.消费卡收款及扫码付时，传入XML串：
        '       <IN>
        '           <MZXSJE>面值销售金额</MZXSJE>
        '           <CZJKJE>充值缴款金额</CZJKJE>
        '           <QRCODE>识别的支付码</QRCODE> //10.35.130后支持
        '       </IN>
        '出参:strSwapGlideNO-交易流水号
        '       strSwapMemo-交易说明
        '       strExpendXML-交易扩展信息:
        '            1)格式1为:项目名称1|项目内容2||…||项目名称n|项目内容n
        '            2)格式2(10.35.90以后才支持),XML格式:
        '            <OUTPUT>
        '                <JYZT>交易状态</JYZT>//接口返回False时，此参数有效:0-交易调用失败;1-交易正在处理中
        '                <ERRMSG>错误信息</ERRMSG> //为空时，不提示，非空时，提示
        '                <JYLIST>//交易列表
        '                    <JY>//保存到预交记录时，按交易流水号及交易说明汇总处理.
        '                        <JYFS>交易方式</JYFS>//Varchar2 20:即结算方式.名称
        '                        <JYJE>交易金额</JYJE>//Number 15 2
        '                        <JYLSH>交易流水号</JYLSH>//Varchar2 50
        '                        <JYSM>交易说明</JYSM>   Varchar2 500
        '                        <JSHM>结算号码</JSHM>   Varchar2 50
        '                        <JSZY>结算摘要</JSZY>   Varchar2 100
        '                        <KH>卡号</KH>   Varchar2 50
        '                        <DJH>单据号<DJH> //Varchar2 8,多单据收费或挂号与卡费一起收时有用 ，针对收费：存储在"医保结算明细"表中,主要是分单据保存，针对挂号费和卡费时，分别存在储病人预交记录中
        '                        <DJXZ>单据性质</DJXZ>//number 2,1多单据收费或挂号与卡费一起收时有用,1-收费单;4-挂号单;5-就诊卡费
        '                        <SFPTJS>是否普通结算</SFPTJS>//Number 1 是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
        '                    </JY>
        '                    ...
        '                </JYLIST>
        '                <Expends>               扩展信息列表
        '                    <Expend>
        '                        <XMMC>项目名称</XMMC> //Varchar2  50
        '                        <XMNR>项目内容</XMNR> //Varchar2  100
        '                    </Expend>
        '                    ...
        '                </Expends>
        '            </OUTPUT>
        '返回:扣款成功,返回true,否则返回Flase
        '说明:
        '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
        '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
        '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
        '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjOneCardComLib.zlPaymentMoney(mfrmMain, mlngModule, lng卡类别ID, bln消费卡, _
        strCardNo, strBalanceIDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strExpendXML) = False Then
        If zlXML_ExistNode(strExpendXML, "OUTPUT") Then
            bln正在交易 = True
            If ZLGetThreeSwapXMLExpend(strExpendXML, dblOutMoney, cll结算方式, _
                strExpend, bln正在交易, strErrMsg) = False Then Exit Function
        End If
        Exit Function
    End If
    
    If zlXML_ExistNode(strExpendXML, "OUTPUT") Then
        If ZLGetThreeSwapXMLExpend(strExpendXML, dblOutMoney, cll结算方式, _
            strExpend) = False Then Exit Function
    Else
        dblOutMoney = dblMoney
        strExpend = strExpendXML
    End If
    ExecutePay = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteReturn(ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal str原结帐IDs As String, ByVal str结帐IDs As String, _
    ByVal dblDelMoney As Double, ByVal strBalanceXml As String, _
    ByVal strSwapGlideNO As String, ByVal strSwapMemo As String, ByRef strExpend As String, _
    ByRef dblOutMoney As Double, ByRef cll结算方式 As Collection, _
    ByRef bln正在交易 As Boolean, ByRef strErrMsg As String) As Boolean
    '一卡通退款交易
    '入参：
    '   strBalanceXml:
    '    <TKLIST>    //本次退款列表
    '        <TK>
    '            <TKFS>退款方式</TKFS>// Varchar2 20 即结算方式.名称
    '            <TKJE>退款金额</TKJE>// Number  15  2
    '            <JYLSH>原交易流水号</JYLSH>//   Varchar2    50
    '            <JYSM>原交易说明</JYSM> //  Varchar2    500
    '            <DJH>单据号</DJH> //    Varchar2 8 单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '        </TK>
    '        ....
    '    </TKLIST>
    '   str原结帐IDs 本次支付所涉及的结算ID(这是原结帐ID):
    '                格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算，7-消费卡管理退费业务
    '   str结帐IDs 本次退费的冲销ID
    '              格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                   收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算，7-消费卡管理退费业务
    '出参：
    '   cllBalance - 结算数据，格式：Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算",原交易流水号,原交易说明,退交易流水号,退交易说明)
    Dim strExpendXML As String, strSwapExtendInfor As String
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    strSwapExtendInfor = str结帐IDs: strTemp = strSwapExtendInfor
    strExpendXML = "<INPUT>" & strBalanceXml & "</INPUT>"
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExpendInfor As String, Optional strExpendXML As String) As Boolean
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
        '       strExpendXML:10.35.90后，扩展此参数,XML格式，如下:
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
        '出参: strSwapNo-交易流水号(退款交易流水号)
        '      strSwapMemo-交易说明(退款交易说明)
        '      strSwapExtendInfor-交易的扩展信息
        '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
        '       strExpendXML：10.35.90以后才有此参数：格式如下
        '            <OUTPUT>
        '                <JYZT>交易状态</JYZT>//接口返回False时，此参数有效:0-交易调用失败;1-交易正在处理中
        '                <ERRMSG>错误信息</ERRMSG> //为空时，不提示，非空时，提示
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
        '返回:函数返回    True:调用成功,False:调用失败
        '说明:
        '       在所有需要对已经消费的地方进行退款的情况下，调用该接口,目前规划在:收费室；挂号等。
    If mobjOneCardComLib.zlReturnMoney(mfrmMain, mlngModule, _
        lng卡类别ID, bln消费卡, strCardNo, str原结帐IDs, dblDelMoney, _
        strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strExpendXML) = False Then
        
        If zlXML_ExistNode(strExpendXML, "OUTPUT") Then
            bln正在交易 = True
            If ZLGetThreeSwapDelXMLExpend(strExpendXML, dblOutMoney, cll结算方式, _
                bln正在交易, strErrMsg) = False Then Exit Function
        End If
        Exit Function
    End If
    
    If zlXML_ExistNode(strExpendXML, "OUTPUT") Then
        If ZLGetThreeSwapDelXMLExpend(strExpendXML, dblOutMoney, cll结算方式) = False Then Exit Function
    End If
    If strTemp <> strSwapExtendInfor Then strExpend = strSwapExtendInfor
    
    ExecuteReturn = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteTransfer(ByVal lng卡类别ID As Long, ByVal strCardNo As String, _
    ByVal str原结帐IDs As String, ByVal str结帐IDs As String, ByVal dblDelMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, ByRef strExpend As String, _
    ByVal byt操作类型 As Byte, ByRef bln正在交易 As Boolean, ByRef strErrMsg As String) As Boolean
    '一卡通转账交易
    '入参：
    '   dblMoney-转账金额
    '   byt操作类型 0-补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务；5-消费卡管理退费业务
    '   str原结帐IDs 结算ID
    '   str结帐IDs 退费业务时，传入本次退费的冲销ID:
    '              格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                   收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算，7-消费卡管理退费业务
    Dim strSwapExtendInfor As String, strExpendXML As String
    Dim strTemp As String, strValue As String
    
    On Error GoTo ErrHandler
    strSwapExtendInfor = str结帐IDs: strTemp = strSwapExtendInfor
    strExpendXML = "<IN><CZLX>" & byt操作类型 & "</CZLX></IN>"
    
    'zlTransferAccountsMoney
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  HIS调用模块号
    'lngCardTypeID   Long    In  卡类别ID
    'strCardNo   String  In  卡号
    'strBalanceID    String  In  结算ID
    'dblMoney    Double  In  转帐金额
    'strSwapGlideNO  String  Out 交易流水号
    'strSwapMemo String  Out 交易说明
    'strSwapExtendInfor  String
    '                       In 退费业务时，传入本次退费的冲销ID:
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '                       Out 交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    'strXMLExpend String
    '                   In   XML串:
    '                       <IN>
    '                           <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务；5-消费卡管理退费业务
    '                       </IN>
    '                   Out  XML串:
    '                       <OUT>
    '                           <JYZT>交易状态</JYZT> //接口返回False时，此参数有效:0-交易调用失败;1-交易正在处理中
    '                           <ERRMSG>错误信息</ERRMSG >
    '                       </OUT>
    '    Boolean 函数返回    True:调用成功,False:调用失败
    '说明:
    '１. 在医保补充结算时进行的三方转帐时调用。
    '２. 一般来说，成功转帐后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '３. 在转帐成功后，返回交易流水号和相关交易说明；如果存在其他交易信息，可以放在扩展信息中返回.
    If mobjOneCardComLib.zlTransferAccountsMoney(mfrmMain, mlngModule, _
        lng卡类别ID, strCardNo, str原结帐IDs, dblDelMoney, _
        strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strExpendXML) = False Then
    
        If zlXML.OpenXMLDocument(strExpendXML) = False Then Exit Function
        Call zlXML.GetSingleNodeValue("OUT/JYZT", strValue)
        bln正在交易 = Val(strValue) = 1
        Call zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue)
        strErrMsg = strValue
        Call zlXML.CloseXMLDocument
        Exit Function
    End If
    If strTemp <> strSwapExtendInfor Then strExpend = strSwapExtendInfor
    
    ExecuteTransfer = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SwapIsSuccess(ByVal intSwapType As Integer, _
    ByVal lng卡类别ID As Long, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, ByVal dblMoney As Double, _
    ByRef bln正在交易 As Boolean, ByRef strErrMsg As String, _
    Optional ByVal strXMLExpend As String) As Boolean
    '判断交易是否成功
    '入参：
    '   intSwapType-0-扣款;1-退款；2-转帐
    '   strBalanceIDs-本次支付所涉及的结算ID
    '                   格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡,6-保险补充结算，7-消费卡管理退费业务
    Dim intSwapStatus As Integer
    On Error GoTo ErrHandler
    intSwapStatus = 0: strErrMsg = ""
    'zlSwapIsSucces(ByVal frmMain As Object, ByVal lngModule As Long, ByVal intSwapType As Integer, _
        ByVal lngCardTypeID As Long, ByVal strBalanceIDs As String, ByVal strCardNo As String, ByVal dblSwapMoney As Double, _
        ByRef intSwapStatus_Out As Integer, Optional strErrMsg As String, Optional strXMLExpend As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:判断交易是否成功（10.35.90）
        '入参:  frmMain-调用的主窗体
        '       lngModule-模块号
        '       intSwapType-0-扣款;1-退款；2-转帐
        '       lngCardTypeID-卡类别ID
        '       strBalanceIDs-本次支付所涉及的结算ID
        '                   格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡,6-保险补充结算，7-消费卡管理退费业务
        '       strCardNo-卡号
        '       strExpend-扩展参数，IN:
        '            <INPUT>
        '                <TKLIST>
        '                    <TK>
        '                       <JYLSH>交易流水号</JYLSH>
        '                       <KH>卡号</KH>
        '                       <JE>金额</JE>
        '                   </TK>
        '                </TKLIST>
        '            </INPUT>
        '出参:intSwapStatus_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
        '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
        '     strXMLExpend-待以后扩展
        '返回：接口调用成功返回true,否则返回Flase
    If mobjOneCardComLib.zlSwapIsSucces(mfrmMain, mlngModule, intSwapType, lng卡类别ID, _
        strBalanceIDs, strCardNo, dblMoney, intSwapStatus, strErrMsg, strXMLExpend) = False Then
        bln正在交易 = (intSwapStatus = 1)
        Exit Function
    End If
    
    SwapIsSuccess = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ZLGetThreeSwapXMLExpend(ByVal strXMLExpend As String, _
    ByRef dblOutMoney As Double, ByRef cllBalance As Collection, ByRef strExpend As String, _
    Optional ByRef bln正在交易 As Boolean, Optional ByRef strErrMsg As String) As Boolean
    '功能：解析三方支付数据
    '入参：
    '   strXMLExpend:XML串
    '    <OUTPUT>
    '        <JYZT>交易状态</JYZT>//接口返回False时，此参数有效:0-交易调用失败;1-交易正在处理中
    '        <ERRMSG>错误信息</ERRMSG> //为空时，不提示，非空时，提示
    '        <JYLIST> //交易列表
    '            <JY>//保存到预交记录时，按交易流水号及交易说明汇总处理.
    '                <JYFS>交易方式</JYFS>//Varchar2 20:即结算方式.名称
    '                <JYJE>交易金额</JYJE>//Number 15 2
    '                <JYLSH>交易流水号</JYLSH>//Varchar2 50
    '                <JYSM>交易说明</JYSM>   Varchar2 500
    '                <JSHM>结算号码</JSHM>   Varchar2 50
    '                <JSZY>结算摘要</JSZY>   Varchar2 100
    '                <KH>卡号</KH>   Varchar2 50
    '                <DJH>单据号<DJH> //Varchar2 8,多单据收费或挂号与卡费一起收时有用 ，针对收费：存储在"医保结算明细"表中,主要是分单据保存，针对挂号费和卡费时，分别存在储病人预交记录中
    '                <DJXZ>单据性质</DJXZ>//number 2,1多单据收费或挂号与卡费一起收时有用,1-收费单;4-挂号单;5-就诊卡费
    '                <SFPTJS>是否普通结算</SFPTJS>//Number 1 是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
    '            </JY>
    '            ...
    '        </JYLIST>
    '        <Expends> //交易扩展信息
    '            <Expend> //保存到预交记录时，按交易流水号及交易说明汇总处理
    '                <XMMC>项目名称</XMMC> //交易方式:即结算方式.名称
    '                <XMNR>项目内容</XMNR>
    '            </Expend>
    '            ...
    '        </Expends>
    '    </OUTPUT>
    '出参：
    '   dblOutMoney - 实际支付金额
    '   cllBalance - 结算数据，格式：Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算|卡号",交易流水号,交易说明)
    '   strExpend - 扩展数据，格式:项目名称1|项目内容2||…||项目名称n|项目内容n
    Dim lngCount As Long, strValue As String
    Dim i As Integer, strBalance As String
    Dim str交易流水号 As String, str交易说明 As String
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    Set cllBalance = New Collection: strExpend = ""
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    
    Call zlXML_GetNodeValue("OUTPUT/JYZT", , strValue)
    bln正在交易 = Val(strValue) = 1 '交易状态
    Call zlXML_GetNodeValue("OUTPUT/ERRMSG", , strValue)
    strErrMsg = strValue '错误信息
    
    '结算信息
    Call zlXML_GetRows("JYLIST/JY", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("JYFS", i, strValue)
        strBalance = strValue   '结算方式
        Call zlXML_GetNodeValue("JYJE", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '结算金额
        dblOutMoney = dblOutMoney + Val(strValue)
        Call zlXML_GetNodeValue("JSHM", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '结算号码
        Call zlXML_GetNodeValue("JSZY", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '结算摘要
        Call zlXML_GetNodeValue("DJH", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '单据号
        Call zlXML_GetNodeValue("SFPTJS", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '是否普通结算
        Call zlXML_GetNodeValue("KH", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '卡号
        
        Call zlXML_GetNodeValue("JYLSH", i, strValue)
        str交易流水号 = strValue '交易流水号
        Call zlXML_GetNodeValue("JYSM", i, strValue)
        str交易说明 = strValue   '交易说明
        
        cllBalance.Add Array(strBalance, str交易流水号, str交易说明)
    Next
    
    '扩展信息
    Call zlXML_GetRows("Expends/Expend", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("XMMC", i, strValue)
        strExpend = strExpend & "||" & strValue '项目名称
        Call zlXML_GetNodeValue("XMNR", i, strValue)
        strExpend = strExpend & "|" & strValue '项目内容
    Next
    If strExpend <> "" Then strExpend = Mid(strExpend, 3)
    ZLGetThreeSwapXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetThreeSwapDelXMLExpend( _
    ByVal strXMLExpend As String, ByRef dblOutMoney As Double, _
    ByRef cllBalance As Collection, _
    Optional ByRef bln正在交易 As Boolean, _
    Optional ByRef strErrMsg As String) As Boolean
    '功能：解析三方退款数据
    '入参：
    '   strXMLExpend:XML串
    '    <OUTPUT>
    '        <JYZT>交易状态</JYZT>//接口返回False时，此参数有效:0-交易调用失败;1-交易正在处理中
    '        <ERRMSG>错误信息</ERRMSG> //为空时，不提示，非空时，提示
    '        <TKLIST>
    '            <TK>
    '                <TKFS>退款方式</TKFS>
    '                <TKJE>结算金额</TKJE>
    '                <YJYLSH>原交易流水号</YJYLSH>
    '                <YJYSM>原交易说明</YJYSM>
    '                <JYLSH>退款交易流水号</JYLSH>
    '                <JYSM>退款交易说明</JYSM>
    '                <DJH>单据号</DJH>
    '                <SFPTJS>是否普通结算</SFPTJS>
    '            </TK>
    '            …
    '        </TKLIST>
    '    </OUTPUT>
    '   blnDelMoney - 是否对金额取相反数
    '出参：
    '   dblOutMoney - 实际退款金额
    '   cllBalance - 结算数据，格式：Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算",原交易流水号,原交易说明,退交易流水号,退交易说明
    Dim lngCount As Long, strValue As String
    Dim i As Integer, strBalance As String
    Dim str原交易流水号 As String, str原交易说明 As String
    Dim str退交易流水号 As String, str退交易说明 As String
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    Set cllBalance = New Collection
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    
    Call zlXML_GetNodeValue("OUTPUT/JYZT", , strValue)
    bln正在交易 = Val(strValue) = 1 '交易状态
    Call zlXML_GetNodeValue("OUTPUT/ERRMSG", , strValue)
    strErrMsg = strValue '错误信息
    
    '结算信息
    Call zlXML_GetRows("TKLIST/TK", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("TKFS", i, strValue)
        strBalance = strValue '退款方式
        Call zlXML_GetNodeValue("TKJE", i, strValue)
        strBalance = strBalance & "|" & -1 * Val(strValue)    '结算金额
        dblOutMoney = dblOutMoney + -1 * Val(strValue)
        strBalance = strBalance & "|" & " " '结算号码
        strBalance = strBalance & "|" & " " '结算摘要
        Call zlXML_GetNodeValue("DJH", i, strValue)
        strBalance = strBalance & "|" & IIf(strValue = "", " ", strValue) '单据号
        Call zlXML_GetNodeValue("SFPTJS", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '是否普通结算
        
        Call zlXML_GetNodeValue("YJYLSH", i, strValue)
        str原交易流水号 = strValue '原交易流水号
        Call zlXML_GetNodeValue("YJYSM", i, strValue)
        str原交易说明 = strValue   '原交易说明
        
        Call zlXML_GetNodeValue("JYLSH", i, strValue)
        str退交易流水号 = strValue '退交易流水号
        Call zlXML_GetNodeValue("JYSM", i, strValue)
        str退交易说明 = strValue   '退交易说明
        
        cllBalance.Add Array(strBalance, str原交易流水号, str原交易说明, str退交易流水号, str退交易说明)
    Next
    ZLGetThreeSwapDelXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckDelToCash(ByVal byt收费类型 As Byte, ByVal lng原结帐ID As Long, ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, strSwapNo As String, ByVal strSwapMemo As String, _
    ByVal dblMoney As Double, ByVal strXMLExpend As String, _
    Optional blnDelDefaultCash_Out As Boolean, _
    Optional strDefaultDelBalance_Out As String) As Boolean
    '三方结算交易退现检查
    '入参：
    '   byt收费类型 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-补充结算,7-消费卡收款
    '   strXMLExpend    XML IN  10.35.90后才支持
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
    '   blnDelDefaultCash_Out-是否缺省退现：接口返回true时有效，true时：表示缺省退成现金（缺省方式为:str缺省退现方式_Out返回值),否则缺省退回原卡，但允许操作员选择退为现金
    '   strDefaultDelBalance_Out-缺省退现方式,比如：支票，现金等
    On Error GoTo ErrHandler
    
    strXMLExpend = "<INPUT>" & vbCrLf & _
                        strXMLExpend & vbCrLf & _
                    "</INPUT>"
    'zlReturnCashCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strXMLExpend As String, _
        Optional blnDelDefaultCash_Out As Boolean, Optional strDefaultDelBalance_Out As String) As Boolean
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
    '       strXMLExpend    XML IN  10.35.90后才支持
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
    '       blnDelDefaultCash_Out-是否缺省退现：接口返回true时有效，true时：表示缺省退成现金（缺省方式为:str缺省退现方式_Out返回值),否则缺省退回原卡，但允许操作员选择退为现金
    '       strDefaultDelBalance_Out-缺省退现方式,比如：支票，现金等
    '       strXMLExpend:10.35.90后才支持
    '        <OUTPUT>
    '            <SFQSTX>是否缺省退现<SFQSTX>//NUMBER 1 是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但以许操作员操作退现
    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
    '                    1.不允许返回三卡方的结算方式
    '                    2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
    '        </OUTPUT>
    '返回:退现合法,返回true,否则返回Flase
    If mobjOneCardComLib.zlReturnCashCheck(mfrmMain, mlngModule, _
        lngCardTypeID, strCardNo, byt收费类型 & "|" & lng原结帐ID, dblMoney, _
        strSwapNo, strSwapMemo, strXMLExpend, _
        blnDelDefaultCash_Out, strDefaultDelBalance_Out) = False Then Exit Function
    CheckDelToCash = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function GetRQCardTypeIDsFromCards(ByRef objPayCards As Cards) As String
    '根据有效的支付卡对象集，获取有效的扫码付的所涉及的卡类别IDs
    '入参:objPayCards-有效的支付方式
    '返回:返回支持 的扫码符的卡类别IDs
    Dim objCard As Card, strCardTypeIDs As String
    
    On Error GoTo ErrHandler
    If objPayCards Is Nothing Then Exit Function
    
    For Each objCard In objPayCards
        If objCard.是否支持扫码付 And objCard.消费卡 = False And objCard.接口序号 > 0 Then
            strCardTypeIDs = strCardTypeIDs & "," & objCard.接口序号
        End If
    Next
    If strCardTypeIDs <> "" Then strCardTypeIDs = Mid(strCardTypeIDs, 2)
    GetRQCardTypeIDsFromCards = strCardTypeIDs
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetCardFromCardType(ByVal lng卡类别ID As Long, bln消费卡 As Boolean, ByVal str结算方式 As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡类别ID获取卡对象
    '入参:lng卡类别ID-卡类别ID
    '     bln消费卡-是否消费卡
    '     str结算方式-结算方式
    '出参:
    '返回:成功卡对象
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As New Card
    On Error GoTo errHandle
    If lng卡类别ID <> 0 Then
        'zlGetCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean,  ByRef objCard As Card) As Boolean
        If mobjOneCardComLib.zlGetCard(lng卡类别ID, bln消费卡, objCard) = False Then
            Set objCard = zlGetCardFromBalanceName(str结算方式)
        End If
    Else
        Set objCard = zlGetCardFromBalanceName(str结算方式)
    End If
    Set zlGetCardFromCardType = objCard: Exit Function

    zlGetCardFromCardType = True
    Exit Function
errHandle:
    Set objCard = zlGetCardFromBalanceName(str结算方式)
    Set zlGetCardFromCardType = objCard: Exit Function
End Function

Public Function zlThird_ReturnMoney_IsValied(ByVal objCard As Card, ByVal byt操作类型 As Byte, _
    ByVal objBalanceInfor As clsBalanceInfo, ByVal objDelBalanceItems As clsBalanceItems, ByRef objItems_Out As clsBalanceItems, Optional blnCurInput As Boolean) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方卡退款合法性检查
    '入参:
    '     objCard-卡对象
    '     objDelBalanceItems-退费结算信息
    '     byt操作类型-0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
    '     objBalanceInfor-当前结帐信息
    '     blnCurInput-是否当前输入的操作
    '出参:objItems_Out-当前返顺的有效结算信息
    '返回:允许退款返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend  As String, str交易说明 As String, str交易流水号 As String, strBalanceIDs As String, str卡号 As String, str密码 As String
    Dim bln退现 As Boolean, intStatus As Integer, strErrMsg As String
    Dim dblMoney  As Double, dbl原始金额 As Double, dbl未退金额 As Double, rsSwap As ADODB.Recordset
    Dim str退费结算IDs As String, str预交退费结算IDs As String, str关联交易IDs As String
    Dim strXmlIn As String
    Dim objItem As clsBalanceItem
    Dim i As Long, blnCheck As Boolean
        
    
    On Error GoTo errHandle
    
    If objCard Is Nothing Then
        Screen.MousePointer = 0
        If InStr(";" & mstrOnCardPrivs & ";", ";三方接口消费;") = 0 Then
            MsgBox "你没有三方接口消费权限，无法调用接口部件！", vbInformation, gstrSysName
            Exit Function
        End If
        MsgBox "未找到退款接口,请检查接口部件！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then zlThird_ReturnMoney_IsValied = True: Exit Function
    
    
    If objDelBalanceItems Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If objDelBalanceItems.Count = 0 Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
        
   '先检查数据的合法性
    blnCheck = True
    If objDelBalanceItems(1).是否保存 Then
        If Not zlThird_IsSwapIsSucces(objDelBalanceItems, intStatus, strErrMsg) Then
           If intStatus = 1 Then
                If strErrMsg <> "" Then
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行扣款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                Else
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行扣款处理中，请稍后再试" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                End If
                Exit Function
           End If
        Else
             Set objItems_Out = objDelBalanceItems
            zlThird_ReturnMoney_IsValied = True: blnCheck = False: Exit Function
        End If
    End If
 
         
         
    dblMoney = -1 * objDelBalanceItems.结算金额
    dbl未退金额 = RoundEx(objBalanceInfor.未付合计 - objBalanceInfor.冲预交合计, 6)
    If dblMoney = 0 Then
        If dbl未退金额 = 0 Then zlThird_ReturnMoney_IsValied = True: Exit Function
        Screen.MousePointer = 0
        MsgBox "退款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dbl未退金额 = 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & objCard.名称 & "已经退完，不能再退！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Abs(dblMoney) > Format(Abs(dbl未退金额), "0.00") And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
       
   
    If dblMoney <> dbl未退金额 And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1 '提示
            Screen.MousePointer = 0
            If MsgBox("当前退款金额与未退金额不符,是否继续？" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(dblMoney, "0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        Case 2 '禁止
            Screen.MousePointer = 0
            Call MsgBox("当前退款金额与未退金额不符,不能继续!" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(dblMoney, "0.00"), vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End Select
    End If
    
    str关联交易IDs = ""
    For Each objItem In objDelBalanceItems
        If InStr(str关联交易IDs & ",", "," & objItem.关联交易ID & ",") = 0 And objItem.关联交易ID <> 0 Then
            str关联交易IDs = str关联交易IDs & "," & objItem.关联交易ID
        End If
    Next
    
    If str关联交易IDs <> "" Then
        str关联交易IDs = Mid(str关联交易IDs, 2)
        If zlGetThirdMoneyInforRecordFromSwapID(str关联交易IDs, rsSwap) = False Then
            str关联交易IDs = ""
        ElseIf rsSwap Is Nothing Then
             str关联交易IDs = ""
        ElseIf rsSwap.State <> 1 Then
            str关联交易IDs = ""
        ElseIf rsSwap.RecordCount = 0 Then
            str关联交易IDs = ""
        End If
    End If
    
    If str关联交易IDs = "" Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    str关联交易IDs = ""
    dbl原始金额 = 0: dbl未退金额 = 0
    For Each objItem In objDelBalanceItems
        
        If InStr(str关联交易IDs & ",", "," & objItem.关联交易ID & ",") = 0 And objItem.关联交易ID <> 0 Then
            str关联交易IDs = str关联交易IDs & "," & objItem.关联交易ID
            rsSwap.Filter = "关联交易ID=" & objItem.关联交易ID
            Do While Not rsSwap.EOF
                dbl未退金额 = dbl未退金额 + RoundEx(Val(Nvl(rsSwap!剩余未退金额)), 6)
                dbl原始金额 = dbl原始金额 + RoundEx(Val(Nvl(rsSwap!原始金额)), 6)
                rsSwap.MoveNext
            Loop
            
            If objCard.是否全退 Then
                If RoundEx(dblMoney, 6) <> RoundEx(dbl原始金额, 6) Then
                    Screen.MousePointer = 0
                    Call MsgBox(objCard.名称 & "必须全退,不能继续!" & vbCrLf & _
                         "原始金额:" & Format(dbl原始金额, "0.00") & _
                         "当前退款:" & Format(dblMoney, "0.00"), vbInformation + vbOKOnly, gstrSysName)
                    Exit Function
                End If
            End If
            If RoundEx(dblMoney, 6) > RoundEx(dbl未退金额, 6) Then
                Screen.MousePointer = 0
                Call MsgBox("当前退款金额大于了剩余未退金额,不能继续!" & vbCrLf & _
                    "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                    "当前退款:" & Format(dblMoney, "0.00"), vbInformation + vbOKOnly, gstrSysName)
                Exit Function
            End If
        End If
        If InStr(str关联交易IDs & ",", "," & objItem.关联交易ID & ",") = 0 And objItem.关联交易ID <> 0 Then
            str关联交易IDs = str关联交易IDs & "," & objItem.关联交易ID
        End If
    Next
        
    Set objItem = objDelBalanceItems(1)
    
    'Public Function zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeId As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
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
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(退款时检查),保险补充结算时传入空
    '       strSwapMemo-交易说明(退款时传入),保险补充结算时传入空
    '       strXMLExpend    XML IN
    '        <TFDATA>   //退费数据
    '            <YCTF>异常退费标志<YCTF> //1-异常重退;0-退费此节点可能没传入
    '            <TFLIST>  //退费列表
    '                <NO></NO>  // 退费单据
    '                <TFITEM>     //退费项
    '                    <SerialNum>序号</SerialNum>
    '                    ….
    '                </ TFITEM >
    '            </TFLIST>
    '
    '            <TKLIST>   //退款列表（35.90以前无此内容）
    '                <TK>
    '                    <TKFS>退款方式</TKFS>// Varchar2    20
    '                    <TKJE>退款金额</TKJE>//NUMBER
    '                    <JSLSH>原交易流水号</JSLSH>//   Varchar2    50
    '                    <JYSM><原交易说明</JYSM>//  Varhcar2    500
    '                    <DJH>单据号</DJH> //    Varchar2    8
    '                </TK>
    '                ....
    '            </TKLIST>
    '        </TFDATA>
    '返回:退款合法,返回true,否则返回Flase
    '说明:
    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
    dblMoney = 0
    strXMLExpend = ""
    str退费结算IDs = ""
    str预交退费结算IDs = ""
    For i = 1 To objDelBalanceItems.Count
        dblMoney = dblMoney + -1 * objDelBalanceItems(i).结算金额
        If objDelBalanceItems(i).是否预交 Then
            If InStr(str预交退费结算IDs & ",", "," & objDelBalanceItems(i).结算IDs & ",") = 0 Then str预交退费结算IDs = str预交退费结算IDs & "," & objDelBalanceItems(i).结算IDs
        Else
           If InStr(str退费结算IDs & ",", "," & objDelBalanceItems(i).结算IDs & ",") = 0 Then str退费结算IDs = str退费结算IDs & "," & objDelBalanceItems(i).结算IDs
        End If
        strXMLExpend = strXMLExpend & vbCrLf & _
        "              <TK>" & vbCrLf & _
        "                   <TKFS>" & objDelBalanceItems(i).结算方式 & "</TKFS>" & vbCrLf & _
        "                   <TKJE>" & -1 * objDelBalanceItems(i).结算金额 & "</TKJE>" & vbCrLf & _
        "                   <JYLSH>" & objDelBalanceItems(i).交易流水号 & "</JYLSH>" & vbCrLf & _
        "                   <JYSM>" & objDelBalanceItems(i).交易说明 & "</JYSM>" & vbCrLf & _
        "                   <DJH>" & objDelBalanceItems(i).单据号 & "</DJH>" & vbCrLf & _
        "              </TK>" & vbCrLf
        
    Next
    strBalanceIDs = ""
    If str退费结算IDs <> "" Then str退费结算IDs = Mid(str退费结算IDs, 2): strBalanceIDs = strBalanceIDs & "||2|" & str退费结算IDs
    If str预交退费结算IDs <> "" Then str预交退费结算IDs = Mid(str预交退费结算IDs, 2): strBalanceIDs = strBalanceIDs & "||1|" & str预交退费结算IDs
    
    If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 3)
    
    str交易说明 = IIf(objDelBalanceItems.Count > 1, "", objItem.交易说明)
    str交易流水号 = IIf(objDelBalanceItems.Count > 1, "", objItem.交易流水号)
       
    strXMLExpend = "" & _
    "<TFDATA>" & vbCrLf & _
    "       <YCTF>" & IIf(objDelBalanceItems.是否异常退费, "1", "0") & "</YCTF>" & vbCrLf & _
    "       <TKLIST>" & vbCrLf & _
                strXMLExpend & vbCrLf & _
    "       </TKLIST>" & vbCrLf & _
    "<TFDATA>" & vbCrLf
    
    If mobjOneCardComLib.zlReturncheck(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, objItem.卡号, _
        strBalanceIDs, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then Exit Function
    
    
    If Not objCard.是否退款验卡 Then
        Set objItems_Out = objDelBalanceItems: zlThird_ReturnMoney_IsValied = True: Exit Function
    End If
   
   'Public Function zlBrushCard(frmMain As Object, _
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
    Optional ByVal strXMLExpend As String = "", _
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
    
    strXMLExpend = "<IN><CZLX>" & byt操作类型 & "</CZLX></IN>"
    '操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    If mobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, _
            objCard.接口序号, objCard.消费卡, mstrPatiName, mstrSex, mstrOld, dblMoney, _
            str卡号, str密码, True, True, bln退现, True, Nothing, False, False, strXmlIn) = False Then Exit Function
                 
    
    Set objItems_Out = objDelBalanceItems: zlThird_ReturnMoney_IsValied = True: Exit Function
    zlThird_ReturnMoney_IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThird_ReturnMoney(objPati As clsPatientInfo, ByVal objCard As Card, objBalanceInfor As clsBalanceInfo, objCurDelBalanceItems As clsBalanceItems, _
    ByRef cllBillPro As Collection, Optional ByVal bln异常作废 As Boolean, Optional objBalanceItems_Out As clsBalanceItems, _
    Optional blnSaveed_Out As Boolean, Optional blnCurInput As Boolean, Optional blnChangeMoney_Out As Boolean, Optional blnOlnyModifyData As Boolean, _
    Optional blnStartOracleTrans As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     objBalanceInfor-原结算方式信息集
    '     objcurDelBalanceItems-当前退费的结算项
    '     bln异常作废-异常作废时调用:true-异常作废;false-正常作废操作
    '     blnCurInput-是否当前输入项
    '     blnOlnyModifyData- 仅修正数据
    '     blnStartOracleTrans-是否启用了oracle事务
    '出参:
    '     objBalanceItems_Out-返回已经退费成功的项目
    '     blnSaveed_Out-是否发生数据保存
    '     : blnStartOracleTrans :返回是进行了事务处理
    '返回:执行成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, cllBeforePro As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim rsBalance As ADODB.Recordset, rsExpend As ADODB.Recordset
    Dim blnTrans As Boolean, blnCallInterfance As Boolean
    Dim intStatus As Integer, strErrMsg As String, strBalanceIDs As String
    Dim strSwapNo As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strCardNo As String, dblMoney As Double, dblDelMoney As Double, str结算方式  As String
    Dim i As Long, lngCount As Long, lng结帐ID As Long
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strExpendXML As String, strTemp As String
   
 
    Err = 0: On Error GoTo errHand:
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    blnSaveed_Out = False
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then zlThird_ReturnMoney = True: Exit Function
    If objCurDelBalanceItems Is Nothing Then Exit Function
    If objCurDelBalanceItems.Count = 0 Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    If GetModifyReturnMoneySaveDatasFromItems(objCurDelBalanceItems, objBalanceInfor, _
        blnOlnyModifyData, True, objPati, strExpendXML, dblMoney, cllPro) = False Then Exit Function
 
    lng结帐ID = IIf(objBalanceInfor.冲销ID <> 0, objBalanceInfor.冲销ID, objBalanceInfor.结帐ID)
    dblDelMoney = dblMoney
    Set objItem = objCurDelBalanceItems(1)
     
    If dblMoney = 0 Then
        MsgBox "当前退款为零了，不允许再次退款", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If objBalanceInfor.结算类型 = 3 And mbln结算异步控制 Then
        blnTrans = False: blnSaveed_Out = True
        If blnStartOracleTrans Then gcnOracle.CommitTrans:   '转帐模式，数据已经保存校对标志为1的记录，因此，只提交
        blnTrans = True: blnStartOracleTrans = False
        zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, True
    Else
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, IIf(mbln结算异步控制, False, True), blnStartOracleTrans
        
        If mbln结算异步控制 Then         '异步处理，需要事先处理校对标志为2(减少接口调用后更新)
           blnTrans = False: blnSaveed_Out = True: blnStartOracleTrans = False
           Set cllBeforePro = New Collection
           Call zlAddUpdateSwapSQL(False, lng结帐ID, objItem.卡类别ID, False, "", "", "", cllBeforePro, 2, objItem.关联交易ID, 1)    '将其更新为2
           zlExecuteProcedureArrAy cllBeforePro, mfrmMain.Caption, True
           blnTrans = True
        End If
    End If
   
    

    
    blnCallInterfance = True '可能会出现接口调用出现异常
    'Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
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
    '                        <DJH>单据号</DJH> //    Varchar2 8 单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '                        <SFPTJS>是否普通结算</SFPTJS>//Number 1 是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
    '                    </TK>
    '                    ...
    '                </TKLIST>
    '            </OUTPUT>
    '    rsBalance_Out-结算信息集(blnResolveXMLToRecord=true时返回)字段（卡类别ID,消费卡,结算方式,交易金额,交易流水号,交易说明,单据号,单据性质,卡号,结算号码,结算摘要,是否普通结算)
    '    rsExpend_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
    '返回:函数返回    True:调用成功,False:调用失败
    '说明:
    '       在所有需要对已经消费的地方进行退款的情况下，调用该接口,目前规划在:收费室；挂号等。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If objItem.是否预交 Then
         strBalanceIDs = "1|" & objItem.结算IDs
    Else
        strBalanceIDs = "2|" & objBalanceInfor.结帐ID
    End If
    
    Set cllPro = New Collection
    strSwapExtendInfor = "2|" & IIf(objBalanceInfor.冲销ID <> 0, objBalanceInfor.冲销ID, objBalanceInfor.结帐ID)
    If mobjOneCardComLib.zlReturnMoney(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, strCardNo, _
        strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strSwapExtendInfor, strExpendXML, True, rsBalance, rsExpend, intStatus, strErrMsg) = False Then
        
        
        If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False: blnStartOracleTrans = False
        
        If strErrMsg <> "" Then
            MsgBox "注意:" & vbCrLf & "    当前退款失败，错误原因如下:" & vbCrLf & strErrMsg, vbOKOnly, gstrSysName
        End If
        
        If intStatus = 1 Or objBalanceInfor.冲销ID <> 0 Then    '冲销的，不能直接删除当前结算，可能后期还有其他操作
            If blnSaveed_Out Then
                For Each objItem In objCurDelBalanceItems
                    objItem.是否保存 = True
                    objItem.是否允许编辑 = False
                    'objItem.是否允许删除 = objItem.是否允许删除
                    objItem.结帐ID = objBalanceInfor.结帐ID
                    objItem.冲销ID = objBalanceInfor.冲销ID
                    
                    objItem.结帐时间 = objBalanceInfor.结帐时间
                    objItem.是否结算 = False
                Next
            End If
            Set objBalanceItems_Out = objCurDelBalanceItems
            Exit Function
        End If
        
        If mbln结算异步控制 And objBalanceInfor.结算类型 <> 3 Then
        
             Set cllUpdate = New Collection
             For Each objItem In objCurDelBalanceItems
                Call zlDelBalanceItemFromItemObject(objItem, cllUpdate, False)
             Next
             Call zlExecuteProcedureArrAy(cllUpdate, mfrmMain.Caption, False, blnTrans)
             blnTrans = False
             blnStartOracleTrans = False
        End If
        
        Set objBalanceItems_Out = objCurDelBalanceItems
        Exit Function
    End If
    
    
    Set objItem = objCurDelBalanceItems(1)
    
    Set objBalanceItems_Out = objCurDelBalanceItems
    '更新数据
    lngCount = 0
    If Not rsBalance Is Nothing Then
        Set objBalanceItems_Out = New clsBalanceItems
        dblMoney = 0
        If rsBalance.RecordCount <> 0 Then rsBalance.MoveFirst
        With rsBalance
            Do While Not .EOF
                Set objItemTemp = New clsBalanceItem
                With objItemTemp
                    .卡类别ID = IIf(Val(Nvl(rsBalance!是否普通结算)) = 0, Val(Nvl(rsBalance!卡类别ID)), 0)
                    .消费卡 = False
                    .消费卡ID = 0
                    .卡号 = IIf(Nvl(rsBalance!卡号) = "", objItem.卡号, Nvl(rsBalance!卡号))
                    .交易流水号 = Nvl(rsBalance!原交易流水号)
                    .交易说明 = Nvl(rsBalance!原交易说明)
                    .退款交易流水号 = Nvl(rsBalance!交易流水号)
                    .退款交易说明 = Nvl(rsBalance!交易说明)
                    .结算方式 = Nvl(rsBalance!结算方式)
                    .结算金额 = -1 * Val(Nvl(rsBalance!交易金额))
                    .密码 = objItem.密码
                    .是否密文 = objItem.是否密文
                    .限制类别 = objItem.限制类别
                    .帐户余额 = objItem.帐户余额
                    .结算号码 = Nvl(rsBalance!结算号码)
                    .结算摘要 = Nvl(rsBalance!结算摘要)
                    .结帐时间 = objBalanceInfor.结帐时间
                    .结帐ID = objBalanceInfor.结帐ID
                    .结算类型 = objItem.结算类型
                    .冲销ID = objBalanceInfor.冲销ID
                    .校对标志 = 2
                    .关联交易ID = objItem.关联交易ID
                    If .卡号 <> "" Then strCardNo = .卡号
                    If strSwapNo = "" And .交易流水号 <> "" Then
                        strSwapNo = .交易流水号
                        If strSwapMemo = "" Then strSwapNo = .交易流水号
                    End If
                    dblMoney = dblMoney + .结算金额
                End With
                objBalanceItems_Out.AddItem objItemTemp
                .MoveNext
            Loop
            '修正数据
            Call GetModifyReturnMoneySaveDatasFromItems(objBalanceItems_Out, objBalanceInfor, False, False, objPati, "", dblMoney, cllUpdate)
            objBalanceItems_Out.结算金额 = dblMoney
        End With
        blnChangeMoney_Out = dblDelMoney <> dblMoney
    Else
        For Each objItem In objBalanceItems_Out
            objItem.退款交易流水号 = strSwapNo
            objItem.退款交易说明 = strSwapMemo
        Next
        Set objItem = objCurDelBalanceItems(1)
        If objItem.是否预交 And mbln结算异步控制 Then
            Call zlGetSaveThirdSwapDelSQLFromBalanceItem(objItem, True, cllUpdate, 0)
        End If
        blnChangeMoney_Out = False
    End If
    
    'blnTrans = True
    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, , Not mbln结算异步控制 Or blnTrans
    blnTrans = False
    For Each objItem In objBalanceItems_Out
        objItem.是否保存 = True
        objItem.是否允许编辑 = False
        objItem.是否允许删除 = False
        objItem.结帐ID = objBalanceInfor.结帐ID
        objItem.结帐时间 = objBalanceInfor.结帐时间
        objItem.冲销ID = objBalanceInfor.冲销ID
        objItem.是否结算 = True
        objItem.校对标志 = 2
    Next
    
    blnSaveed_Out = True: blnStartOracleTrans = False
    
    Set cllBillPro = New Collection
    blnTrans = False
    
    'objBalanceInfor.已付合计 = objBalanceInfor.已付合计 + dblMoney
    'objBalanceInfor.未付合计 = objBalanceInfor.当前结帐 - objBalanceInfor.已付合计
    Call zlGetThreeSwapExpendSQLToCollect(objCard, IIf(objBalanceInfor.冲销ID = 0, objBalanceInfor.结帐ID, objBalanceInfor.冲销ID), strCardNo, rsExpend, cllThreeSwap, False)
    Err = 0: On Error GoTo ErrOther:
    blnTrans = True
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption & "-执行三方接口"
    
    blnTrans = False
    zlThird_ReturnMoney = True
    Exit Function
errHand:
    If blnTrans And blnCallInterfance = False Then gcnOracle.RollbackTrans: blnTrans = False: blnStartOracleTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    If blnTrans Then gcnOracle.RollbackTrans: blnStartOracleTrans = False
    Exit Function
ErrOther:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
End Function


Private Function GetModifyReturnMoneySaveDatasFromItems(ByRef objDelItems As clsBalanceItems, ByRef objBalanceInfor As clsBalanceInfo, _
    ByVal blnOlnyModifyData As Boolean, ByVal blnExcuteBefore As Boolean, objPati As clsPatientInfo, _
    ByRef strXMLRetrun_Out As String, ByRef objDelMoney_Out As Double, ByRef cllPro_out As Collection) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算项，返回修正的数据结算SQL集
    '入参:
    '    blnExcuteBefore-接口执行前
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, strExpendXML As String
    Dim str结算方式 As String, strSql As String, strTemp As String
    Dim strCardNo As String, strSwapNo As String, strSwapMemo As String, i As Long
    On Error GoTo errHandle
    
    'objBalanceInfor.结算类型:1-门诊结帐;2-住院结帐;3-门诊费用转住院
    '1.门诊费用转住院
    If cllPro_out Is Nothing Then Set cllPro_out = New Collection
    
    If objBalanceInfor.结算类型 = 3 Then
        If blnExcuteBefore Then '接口执行前，只更新校对标志
            Set objItem = objDelItems(1)
            'Zl_门诊费用转住院_Modify_s
            strSql = "Zl_门诊费用转住院_Modify_s("
            '  --操作类型_In:
            '  --   0-仅更新校对标志:只更新关联交易ID的校对标志
            '  --   1-普通退费方式:
            '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
            '  --   2.三方卡退费结算:
            '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
            '  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
            '  --     关联交易ID_IN:
            '  --     清除原交易_In:1-表示在更新数据前，清除原来的交易信息(按结帐ID+关联交易ID来清除);0-表示不清除
            '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
            '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
            '  --   4-消费卡结算:
            '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
            '操作类型_In   Number,
            strSql = strSql & "0,"
            '冲销id_In     病人预交记录.结帐id%Type,
            strSql = strSql & "" & objBalanceInfor.冲销ID & ","
            '病人id_In     病人结帐记录.病人id%Type,
            strSql = strSql & "" & mlng病人ID & ","
            '  姓名_In         病人预交记录.姓名%Type,
            strSql = strSql & "'" & objPati.姓名 & "' ,"
            '  性别_In         病人预交记录.性别%Type,
            strSql = strSql & "'" & objPati.性别 & "' ,"
            '  年龄_In         病人预交记录.年龄%Type,
            strSql = strSql & "'" & objPati.年龄 & "' ,"
            '  门诊号_In       病人预交记录.门诊号%Type,
            strSql = strSql & "'" & objPati.门诊号 & "' ,"
            '  住院号_In       病人预交记录.住院号%Type,
            strSql = strSql & "'" & objPati.住院号 & "' ,"
            '  付款方式名称_In 病人预交记录.付款方式名称%Type,
            strSql = strSql & "'" & objPati.医疗付款方式 & "' ,"
            '结算方式_In   Varchar2,
            strSql = strSql & "NULL,"
            '操作员编号_In 病人预交记录.操作员编号%Type := Null,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '操作员姓名_In 病人预交记录.操作员姓名%Type := Null,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '完成退费_In   Number := 0,
            strSql = strSql & "0,"
            '关联交易id_In 病人预交记录.Id%Type := Null,
            strSql = strSql & ZVal(objItem.关联交易ID) & ","
            '退款时间_In   病人预交记录.收款时间%Type := Null,
            strSql = strSql & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '校对标志_In   病人预交记录.校对标志%Type := Null,
            strSql = strSql & "2)"
            '误差金额_In   病人预交记录.冲预交%Type := Null,
            '卡类别id_In   病人预交记录.卡类别id%Type := Null,
            '卡号_In       病人预交记录.卡号%Type := Null,
            '交易流水号_In 病人预交记录.交易流水号%Type := Null,
            '交易说明_In   病人预交记录.交易说明%Type := Null,
            '清除原交易_In Number:=0
            zlAddArray cllPro_out, strSql
            strTemp = ""
            For Each objItem In objDelItems
                strTemp = strTemp & "        <TK>" & vbCrLf
                strTemp = strTemp & "           <TKFS>" & objItem.结算方式 & "</TKFS>" & vbCrLf
                strTemp = strTemp & "           <TKJE>" & -1 * objItem.结算金额 & "</TKJE>" & vbCrLf
                strTemp = strTemp & "           <JYLSH>" & objItem.交易流水号 & "</JYLSH>" & vbCrLf
                strTemp = strTemp & "           <JYSM>" & objItem.交易说明 & "</JYSM>" & vbCrLf
                strTemp = strTemp & "           <DJH>" & objItem.单据号 & "</DJH>" & vbCrLf
                strTemp = strTemp & "        </TK>"
                objDelMoney_Out = objDelMoney_Out + -1 * objItem.结算金额
            Next
            strExpendXML = strExpendXML & "<INPUT>" & vbCrLf
            strExpendXML = strExpendXML & "    <TKLIST>" & vbCrLf
            strExpendXML = strExpendXML & strTemp & vbCrLf
            strExpendXML = strExpendXML & "    </TKLIST>" & vbCrLf
            strExpendXML = strExpendXML & "</INPUT>" & vbCrLf
            strXMLRetrun_Out = strExpendXML
            GetModifyReturnMoneySaveDatasFromItems = True: Exit Function
        End If
        
        For Each objItem In objDelItems
            strTemp = strTemp & "        <TK>" & vbCrLf
            strTemp = strTemp & "           <TKFS>" & objItem.结算方式 & "</TKFS>" & vbCrLf
            strTemp = strTemp & "           <TKJE>" & -1 * objItem.结算金额 & "</TKJE>" & vbCrLf
            strTemp = strTemp & "           <JYLSH>" & objItem.交易流水号 & "</JYLSH>" & vbCrLf
            strTemp = strTemp & "           <JYSM>" & objItem.交易说明 & "</JYSM>" & vbCrLf
            strTemp = strTemp & "           <DJH>" & objItem.单据号 & "</DJH>" & vbCrLf
            strTemp = strTemp & "        </TK>"
            
            objDelMoney_Out = objDelMoney_Out + -1 * objItem.结算金额
            strCardNo = objItem.卡号
            strSwapNo = objItem.交易流水号
            strSwapMemo = objItem.交易说明
            
            '结算方式|结算金额|结算号码|结算摘要||..
            str结算方式 = objItem.结算方式 & "|" & objItem.结算金额
            str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算号码) = "", " ", Trim(objItem.结算号码))
            str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算摘要) = "", " ", Trim(objItem.结算摘要))
            
            'Zl_门诊费用转住院_Modify_s
            strSql = "Zl_门诊费用转住院_Modify_s("
            '  --操作类型_In:
            '  --   0-仅更新校对标志:只更新关联交易ID的校对标志
            '  --   1-普通退费方式:
            '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
            '  --   2.三方卡退费结算:
            '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
            '  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
            '  --     关联交易ID_IN:
            '  --     清除原交易_In:1-表示在更新数据前，清除原来的交易信息(按结帐ID+关联交易ID来清除);0-表示不清除
            '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
            '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
            '  --   4-消费卡结算:
            '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
            '操作类型_In   Number,
            strSql = strSql & "2,"
            '冲销id_In     病人预交记录.结帐id%Type,
            strSql = strSql & "" & objBalanceInfor.冲销ID & ","
            '病人id_In     病人结帐记录.病人id%Type,
            strSql = strSql & "" & mlng病人ID & ","
            '  姓名_In         病人预交记录.姓名%Type,
            strSql = strSql & "'" & objPati.姓名 & "' ,"
            '  性别_In         病人预交记录.性别%Type,
            strSql = strSql & "'" & objPati.性别 & "' ,"
            '  年龄_In         病人预交记录.年龄%Type,
            strSql = strSql & "'" & objPati.年龄 & "' ,"
            '  门诊号_In       病人预交记录.门诊号%Type,
            strSql = strSql & "'" & objPati.门诊号 & "' ,"
            '  住院号_In       病人预交记录.住院号%Type,
            strSql = strSql & "'" & objPati.住院号 & "' ,"
            '  付款方式名称_In 病人预交记录.付款方式名称%Type,
            strSql = strSql & "'" & objPati.医疗付款方式 & "' ,"
            '结算方式_In   Varchar2,
            strSql = strSql & "'" & str结算方式 & "',"
            '操作员编号_In 病人预交记录.操作员编号%Type := Null,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '操作员姓名_In 病人预交记录.操作员姓名%Type := Null,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '完成退费_In   Number := 0,
            strSql = strSql & "0,"
            '关联交易id_In 病人预交记录.Id%Type := Null,
            strSql = strSql & "" & IIf(Val(objItem.关联交易ID) = 0, "NULL", objItem.关联交易ID) & ","
            '退款时间_In   病人预交记录.收款时间%Type := Null,
            strSql = strSql & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '校对标志_In   病人预交记录.校对标志%Type := Null,
            strSql = strSql & 2 & ","
            '误差金额_In   病人预交记录.冲预交%Type := Null,
            strSql = strSql & "NULL,"
            '卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSql = strSql & "" & objItem.卡类别ID & ","
            '卡号_In       病人预交记录.卡号%Type := Null,
            strSql = strSql & "'" & objItem.卡号 & "',"
            '交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSql = strSql & "'" & objItem.交易流水号 & "',"
            '交易说明_In   病人预交记录.交易说明%Type := Null,
            strSql = strSql & "'" & objItem.交易说明 & "',"
            '清除原交易_In Number:=0
            strSql = strSql & IIf(i = 0, "1", "0") & ")"
            zlAddArray cllPro_out, strSql
            i = i + 1
        Next
        strExpendXML = strExpendXML & "<INPUT>" & vbCrLf
        strExpendXML = strExpendXML & "    <TKLIST>" & vbCrLf
        strExpendXML = strExpendXML & strTemp & vbCrLf
        strExpendXML = strExpendXML & "    </TKLIST>" & vbCrLf
        strExpendXML = strExpendXML & "</INPUT>" & vbCrLf
        strXMLRetrun_Out = strExpendXML
        
        GetModifyReturnMoneySaveDatasFromItems = True
        Exit Function
    End If
    
    '2.门诊及住院结帐:
    For Each objItem In objDelItems
        strTemp = strTemp & "        <TK>" & vbCrLf
        strTemp = strTemp & "           <TKFS>" & objItem.结算方式 & "</TKFS>" & vbCrLf
        strTemp = strTemp & "           <TKJE>" & -1 * objItem.结算金额 & "</TKJE>" & vbCrLf
        strTemp = strTemp & "           <JYLSH>" & objItem.交易流水号 & "</JYLSH>" & vbCrLf
        strTemp = strTemp & "           <JYSM>" & objItem.交易说明 & "</JYSM>" & vbCrLf
        strTemp = strTemp & "           <DJH>" & "" & "</DJH>" & vbCrLf
        strTemp = strTemp & "        </TK>"
        strCardNo = objItem.卡号
        strSwapNo = objItem.交易流水号
        strSwapMemo = objItem.交易说明
        
        objDelMoney_Out = objDelMoney_Out + -1 * objItem.结算金额
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = objItem.结算方式 & "|" & objItem.结算金额
        str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算号码) = "", " ", Trim(objItem.结算号码))
        str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算摘要) = "", " ", Trim(objItem.结算摘要))
        
        
        If objBalanceInfor.冲销ID = 0 Then
            '调用之前,先处理数据
            'Zl_病人结帐结算_Modify_S
            strSql = "Zl_病人结帐结算_Modify_S("
            '  操作类型_In     Number,
            '  --操作类型_In:
            '  --   0-普通收费方式:
            '  --   1.三方卡结算:
            '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
            '  --     ②退支票额_In:传入零
            '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
            strSql = strSql & "1,"
            '  病人id_In       病人结帐记录.病人id%Type,
            strSql = strSql & "" & mlng病人ID & ","
            '  结帐id_In       病人预交记录.结帐id%Type,
            strSql = strSql & "" & objBalanceInfor.结帐ID & ","
            '  结算方式_In     Varchar2,
            strSql = strSql & "'" & str结算方式 & "',"
            '  冲预交_In       病人预交记录.冲预交%Type := Null,
            strSql = strSql & "NULL,"
            '  退支票额_In     病人预交记录.冲预交%Type := Null,
            strSql = strSql & "NULL,"
            '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
            strSql = strSql & "" & objItem.卡类别ID & ","
            '  卡号_In         病人预交记录.卡号%Type := Null,
            strSql = strSql & "'" & objItem.卡号 & "',"
            '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
            strSql = strSql & "'" & objItem.交易流水号 & "',"
            '  交易说明_In     病人预交记录.交易说明%Type := Null,
            strSql = strSql & "'" & objItem.交易说明 & "',"
            '  缴款_In         病人预交记录.缴款%Type := Null,
            strSql = strSql & "" & "Null" & ","
            '  找补_In         病人预交记录.找补%Type := Null,
            strSql = strSql & "" & "Null" & ","
            '  误差金额_In     门诊费用记录.实收金额%Type := Null,
            strSql = strSql & "NULL,"
            '  结帐类型_In     Number := 2,
            strSql = strSql & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
            '  缺省结算方式_In 结算方式.名称%Type := Null,
            strSql = strSql & "NULL,"
            '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
            strSql = strSql & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '    冲预交病人ids_In Varchar2 := Null,
            strSql = strSql & "NULL,"
            '  完成结算_In Number:=0
            strSql = strSql & "0,"
            '  校对标志_In      Number := 2,
            strSql = strSql & "" & IIf(mbln结算异步控制 And blnExcuteBefore, 1, 2) & ","
            '  预交id_In        病人预交记录.Id%Type := Null,
            strSql = strSql & "NULL,"
            '  关联交易id_In    病人预交记录.Id%Type := Null,
            strSql = strSql & IIf(Val(objItem.关联交易ID) = 0, "NULL", Val(objItem.关联交易ID)) & ","
            '  清除原交易_In Number:=0
            strSql = strSql & "" & IIf(i = 0, 1, 0) & ","
            '附加标志_In    病人预交记录.Id%Type := Null
            '对于结帐业务(三方卡退款传入)：NULl or 0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
            strSql = strSql & "" & 1 & ")"
            If Not blnOlnyModifyData Then zlAddArray cllPro_out, strSql
            If objItem.是否预交 And mbln结算异步控制 And Not blnExcuteBefore Then Call zlGetSaveThirdSwapDelSQLFromBalanceItem(objItem, True, cllPro_out, 0)
        Else
            'Zl_病人结帐作废_Modify
            strSql = "Zl_病人结帐作废_Modify("
            '  操作类型_In   Number,
            '--   1-普通退费方式:
            '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
            '--   2.三方卡退费结算:
            '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
            '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
            '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
            '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
            '--   4-消费卡结算:
            '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
            strSql = strSql & "" & 2 & ","
            '  病人id_In     病人结帐记录.病人id%Type,
            strSql = strSql & "" & mlng病人ID & ","
            '  冲销id_In     病人预交记录.结帐id%Type,
            strSql = strSql & "" & objBalanceInfor.冲销ID & ","
            '  结算方式_In   Varchar2,
            strSql = strSql & "'" & str结算方式 & "',"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSql = strSql & "" & objItem.卡类别ID & ","
            '  卡号_In       病人预交记录.卡号%Type := Null,
            strSql = strSql & "'" & objItem.卡号 & "',"
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSql = strSql & "'" & objItem.交易流水号 & "',"
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            strSql = strSql & "'" & objItem.交易说明 & "',"
            '  缴款_In       病人预交记录.缴款%Type := Null,
            strSql = strSql & "NULL,"
            '  找补_In       病人预交记录.找补%Type := Null,
            strSql = strSql & "NULL,"
            '  误差金额_In   病人预交记录.冲预交%Type := Null,
            strSql = strSql & "NULL,"
            '  预交金额_In   病人预交记录.冲预交%Type := Null,
            strSql = strSql & "NULL,"
            '  操作员编号_In    病人预交记录.操作员编号%Type := Null,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  收款时间_In      病人预交记录.收款时间%Type := Null,
            strSql = strSql & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '  冲预交病人ids_In Varchar2 := Null,
            strSql = strSql & "NULL,"
            '  完成作废_In Number:=0
            strSql = strSql & "0,"
            '  校对标志_In  Number := 0,
            strSql = strSql & IIf(mbln结算异步控制 And blnExcuteBefore, 1, 2) & ","
            '  关联交易id_In 病人预交记录.Id%Type := Null,
            strSql = strSql & "" & IIf(Val(objItem.关联交易ID) = 0, "NULL", objItem.关联交易ID) & ","
            '  清除原交易_In Number:=0
            strSql = strSql & IIf(i = 0, "1", "0") & ")"
            If Not blnOlnyModifyData Or Not blnExcuteBefore Then zlAddArray cllPro_out, strSql     '异常作废已经处理
        End If
        i = i + 1
    Next
    strExpendXML = strExpendXML & "<INPUT>" & vbCrLf
    strExpendXML = strExpendXML & "    <TKLIST>" & vbCrLf
    strExpendXML = strExpendXML & strTemp & vbCrLf
    strExpendXML = strExpendXML & "    </TKLIST>" & vbCrLf
    strExpendXML = strExpendXML & "</INPUT>" & vbCrLf
    strXMLRetrun_Out = strExpendXML
    
    GetModifyReturnMoneySaveDatasFromItems = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Class_Initialize()
    mstrOnCardPrivs = ";" & GetPrivFunc(glngSys, 1151) & ";"   '一卡通相关接口权限
End Sub

Public Function zlThird_IsSwapIsSucces(ByVal objItems As clsBalanceItems, ByRef intSwapStatu_Out As Integer, ByRef strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断交易是否调用成功
    '入参: objItems-当前结算信息集
    '出参:intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
    '     strErrMsg_out-返回的错误信息
    '返回：接口调用成功返回true,否则返回Flase
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intSwapType As Integer, dblMoney As Double, strXml As String, strXMLExpend As String
    Dim objCard As Card, objItem As clsBalanceItem, objTemp As clsBalanceItem
    Dim strBalanceIDs As String
  
    If objItems.Count = 0 Then Exit Function
    
    dblMoney = 0
    For Each objItem In objItems
        dblMoney = dblMoney + objItem.结算金额
    Next
    Set objItem = objItems(1)
    Set objCard = objItem.objCard
    On Error GoTo errHandle
   '0-扣款;1-退款；2-转帐
    intSwapType = IIf(objItem.是否退款, IIf(objItem.是否转帐, 2, 1), 0)
    strBalanceIDs = "2|" & IIf(objItem.冲销ID <> 0, objItem.冲销ID, objItem.结帐ID)
    
    'zlSwapIsSucces(ByVal frmMain As Object, ByVal lngModule As Long, ByVal intSwapType As Integer, _
    '    ByVal lngCardTypeID As Long, ByVal strBalanceIDs As String, ByVal strCardNO As String, ByVal dblSwapMoney As Double, _
    '    ByRef intSwapStatus_Out As Integer, Optional strErrMsg As String, Optional strXMLExpend As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:判断交易是否成功（10.35.110）
    '    '入参:  frmMain-调用的主窗体
    '    '       lngModule-模块号
    '    '       intSwapType-0-扣款;1-退款；2-转帐
    '    '       lngCardTypeID-卡类别ID
    '    '       strCardNO-卡号
    '    '       dblSwapMoney-交易金额
    '    '       strBalanceIDs-本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn 收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
    '    '       strExpend-扩展参数:退款或作废时，才传入,格式如下 ：
    '    '        <INPUT>
    '    '            <TKLIST>
    '    '                    <TK>
    '    '                       <JYLSH>交易流水号</JYLSH>
    '    '                       <KH>卡号</KH>
    '    '                       <JE>金额</JE>
    '    '                    </TK>
    '    '            </TKLIST>
    '    '        </INPUT>
    '    '出参:intSwapStatus_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '    '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
    '    '     strXMLExpend-待以后扩展
    '    '返回：接口调用成功返回true,否则返回Flase
    '    '说明:
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = "": strXml = ""
    For Each objTemp In objItems
        strXml = strXml & "        <TK>" & vbCrLf
        strXml = strXml & "             <JYLSH>" & objTemp.交易流水号 & "</JYLSH> & vbCrLf"
        strXml = strXml & "             <KH>" & objTemp.卡号 & "</KH> & vbCrLf"
        strXml = strXml & "             <JE>" & objTemp.结算金额 & "</JE> & vbCrLf"
        strXml = strXml & "        </TK>" & vbCrLf
    Next
    
    strXml = strXml & "    <TKLIST>" & vbCrLf & strXml
    strXml = strXml & "    </TKLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "</INPUT>" & vbCrLf
    strXMLExpend = strXMLExpend & strXml
    strXMLExpend = strXMLExpend & "</INPUT>" & vbCrLf
    zlThird_IsSwapIsSucces = mobjOneCardComLib.zlSwapIsSucces(mfrmMain, mlngModule, intSwapType, objCard.接口序号, strBalanceIDs, objItem.卡号, dblMoney, intSwapStatu_Out, strErrmsg_Out, strXMLExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetThirdMoneyInforRecordFromSwapID(ByVal str关联交易IDs As String, ByRef rsSwapRecord_Out As ADODB.Recordset, _
    Optional ByVal lng冲销ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,获取相关的结算金额信息集
    '入参:str关联交易IDs-关联交易ID，多个用逗号分离
    '   lng冲销ID 本次冲销ID，结算数据不能包含
    '出参:rsSwapRecord_Out-返回关联交易ID
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim strWhere As String, lng关联交易ID As Long
    
    On Error GoTo errHandle
    
    If InStr(str关联交易IDs, ",") > 0 Then
        strWhere = "And A.关联交易ID In (Select column_value From table(f_num2List([1])) "
        
    Else
       lng关联交易ID = Val(str关联交易IDs)
       strWhere = " And  A.关联交易ID =[2]"
    End If
    If lng冲销ID <> 0 Then strWhere = strWhere & " And  A.结帐ID <> [3]"
    
    strSql = "" & _
    "   Select 关联交易ID,卡类别ID,结算方式,交易流水号,交易说明, " & vbCrLf & _
    "          decode(sign(nvl(金额,0)),1,1,0)*nvl(金额,0)+ decode(sign(nvl(冲预交,0)),1,1,0)* nvl(冲预交,0) as 原始金额, " & _
    "          decode(sign(nvl(金额,0)),-1,1,0)*nvl(金额,0)+ decode(sign(nvl(冲预交,0)),-1,1,0)* nvl(冲预交,0) as 已退金额" & _
    "   From 病人预交记录 A " & _
    "   Where 1=1 " & strWhere & _
    "   Union all " & _
    "   Select a.关联交易ID,a.卡类别ID,a.结算方式,a.交易流水号,a.交易说明, " & vbCrLf & _
    "          0 as 原始金额, " & _
    "         -1*nvl(b.金额,0) as 已退金额" & _
    "   From 病人预交记录 A,三方退款信息 B" & _
    "   Where  a.ID=b.记录ID And b.是否转帐 =1  " & strWhere

    
    strSql = "" & _
    " Select 关联交易ID,卡类别ID,a.结算方式,a.交易流水号,a.交易说明, sum(原始金额) as 原始金额, sum(已退金额) as 已退金额, sum(原始金额)-sum(已退金额) as 剩余未退金额" & _
    " From (" & strSql & ") A " & _
    " Group by a.关联交易ID,a.卡类别ID,a.结算方式,a.交易流水号,a.交易说明"
    
    Set rsSwapRecord_Out = zlDatabase.OpenSQLRecord(strSql, "获取三方交易的原始金额及未退金额", _
        str关联交易IDs, lng关联交易ID, lng冲销ID)
    zlGetThirdMoneyInforRecordFromSwapID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlDelBalanceItemFromItemObject(ByVal objItem As clsBalanceItem, Optional ByRef cllSaveSqL_Out As Collection, Optional blnExcute As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算对象,获取删除结算项的SQL
    '入参:blnExcute-是否进行执行操作
    '出参:cllSaveSqL_Out-返回保存的SQL
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans As Boolean, strSql As String
    On Error GoTo errHandle
    If cllSaveSqL_Out Is Nothing Then Set cllSaveSqL_Out = New Collection
    '明确失败
    '删除当前结算信息
    '    Zl_病人结帐结算_Delete
    strSql = "Zl_病人结帐结算_Delete("
    '      结帐id_In     病人预交记录.结帐id%Type,
    strSql = strSql & "" & IIf(objItem.冲销ID <> 0, objItem.冲销ID, objItem.结帐ID) & ","
    '      卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSql = strSql & "" & objItem.卡类别ID & ","
    '      关联交易id_In 病人预交记录.Id%Type := Null,
    strSql = strSql & "" & objItem.关联交易ID & ")"
    '      结算方式_In Varchar2:=Null
    zlAddArray cllSaveSqL_Out, strSql
    
    If Not blnExcute Then Exit Function
    blnTrans = True
    zlExecuteProcedureArrAy cllSaveSqL_Out, "执行病人结算信息删除"
     blnTrans = True = False
    zlDelBalanceItemFromItemObject = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetSaveThirdSwapDelSQLFromBalanceItem(ByVal objItem As clsBalanceItem, ByVal blnModify As Boolean, ByRef cllPro As Collection, Optional int校对标志 As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算对象获取三方交易退款保存SQL
    '入参:objItem-当前结帐对象
    '     blnModify-是否修改
    '     bln转帐-是否当前进行的转帐操作
    '     int校对标志-1-接口未成功;0-接口调用成功
    '出参:cllPro-返回的结算信息
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle
    ' Zl_三方退款信息_Insert
    strSql = "Zl_三方退款信息_Insert("
    '  结帐id_In     三方退款信息.结帐id%Type,
    strSql = strSql & "" & objItem.结帐ID & ","
    '  记录id_In     三方退款信息.记录id%Type,
    strSql = strSql & "" & objItem.预交ID & ","
    '  金额_In       三方退款信息.金额%Type,
    strSql = strSql & "" & Abs(objItem.结算金额) & ","
    '  卡号_In       三方退款信息.卡号%Type,
    strSql = strSql & "'" & objItem.卡号 & "',"
    '  交易流水号_In 三方退款信息.交易流水号%Type,
    strSql = strSql & "'" & objItem.退款交易流水号 & "',"
    '  交易说明_In   三方退款信息.交易说明%Type,
    strSql = strSql & "'" & objItem.退款交易说明 & "',"
    '  操作类型_In   Number := 0,
    strSql = strSql & "'" & IIf(blnModify, 1, 0) & "',"
    '  是否未退_In   三方退款信息.是否未退%Type := 0
    strSql = strSql & "'" & IIf(int校对标志 = 1, 1, 0) & "',"
    '  是否转帐_In   三方退款信息.是否转帐%Type := 0
    strSql = strSql & "'" & IIf(objItem.是否转帐, 1, 0) & "',"
    '  卡类别id_In   三方退款信息.卡类别id%Type := Null
    strSql = strSql & "" & IIf(objItem.卡类别ID = 0, "NULL", objItem.卡类别ID) & ","
    '  原交易流水号_In 三方退款信息.原交易流水号%Type := Null,
    strSql = strSql & "'" & objItem.交易流水号 & "',"
    '  原交易说明_In   三方退款信息.原交易说明%Type := Null
    strSql = strSql & "'" & objItem.交易说明 & "')"
  
    zlAddArray cllPro, strSql
    zlGetSaveThirdSwapDelSQLFromBalanceItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetThreeSwapExpendSQLToCollect(objCard As Card, ByVal lng结帐ID As Long, ByVal str卡号 As String, ByVal rsExpend As ADODB.Recordset, _
    ByRef cllPro As Collection, Optional bln预交 As Boolean = False, Optional lng预交ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取扩展信息保存的SQL给集合
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, str交易信息 As String, strTemp As String
    On Error GoTo errHandle
    
    If cllPro Is Nothing Then Set cllPro = New Collection
    If rsExpend Is Nothing Then zlGetThreeSwapExpendSQLToCollect = True: Exit Function
    If rsExpend.State <> 1 Then zlGetThreeSwapExpendSQLToCollect = True: Exit Function
    
    With rsExpend
        rsExpend.Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        str交易信息 = ""
        Do While Not .EOF
            If Nvl(!项目名称) <> "" Then
                strTemp = Nvl(!项目名称) & "|" & Nvl(!项目内容)
                If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                        str交易信息 = Mid(str交易信息, 3)
                        'Zl_三方结算交易_Insert
                        strSql = "Zl_三方结算交易_Insert("
                        '卡类别id_In 病人预交记录.卡类别id%Type,
                        strSql = strSql & "" & objCard.接口序号 & ","
                        '消费卡_In   Number,
                        strSql = strSql & "" & IIf(objCard.消费卡, 1, 0) & ","
                        '卡号_In     病人预交记录.卡号%Type,
                        strSql = strSql & "'" & str卡号 & "',"
                        '结帐ids_In  Varchar2,
                        strSql = strSql & "'" & lng结帐ID & "',"
                        '交易信息_In Varchar2:交易项目|交易内容||...
                        strSql = strSql & "'" & str交易信息 & "',"
                        '预交款缴款_In Number := 0
                        strSql = strSql & IIf(bln预交, "1", "0") & ",Null,"
                        strSql = strSql & IIf(lng预交ID = 0, "NULL", lng预交ID) & ")"
                        zlAddArray cllPro, strSql
                        str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
            .MoveNext
        Loop
        
    End With
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSql = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSql = strSql & "" & objCard.接口序号 & ","
        '消费卡_In   Number,
        strSql = strSql & "" & IIf(objCard.消费卡, 1, 0) & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSql = strSql & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSql = strSql & "'" & lng结帐ID & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSql = strSql & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSql = strSql & IIf(bln预交, "1", "0") & ",Null,"
        strSql = strSql & IIf(lng预交ID = 0, "NULL", lng预交ID) & ")"
        zlAddArray cllPro, strSql
    End If
    zlGetThreeSwapExpendSQLToCollect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
