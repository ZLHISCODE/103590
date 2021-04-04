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

Public Function Init(frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal lng病人ID As Long, Optional ByVal strPatiName As String, _
    Optional ByVal strSex As String, Optional ByVal strOld As String)
    '初始化类模块变量
    On Error GoTo ErrHandler
    Set mfrmMain = frmMain
    mlngModule = lngModule
      
    mlng病人ID = lng病人ID
    mstrPatiName = strPatiName: mstrSex = strSex: mstrOld = strOld
    Init = True
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
         
    If gobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, rsClassMoney, lng卡类别ID, bln消费卡, _
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
    If gobjOneCardComLib.zlPaymentCheck(mfrmMain, mlngModule, lng卡类别ID, bln消费卡, _
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
        If gobjOneCardComLib.zlGetAccountMoney(mfrmMain, mlngModule, _
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
    If gobjOneCardComLib.zlReturncheck(mfrmMain, mlngModule, lng卡类别ID, False, strCardNo, _
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
        If gobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, lng卡类别ID, False, _
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
    If gobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, lng卡类别ID, False, _
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
    If gobjOneCardComLib.zlTransferAccountsCheck(mfrmMain, mlngModule, lng卡类别ID, _
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

Private Sub ZlShowErrMsg(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
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
        If bytType = 0 Then
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
    If gobjOneCardComLib.zlPaymentMoney(mfrmMain, mlngModule, lng卡类别ID, bln消费卡, _
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
    If gobjOneCardComLib.zlReturnMoney(mfrmMain, mlngModule, _
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
    If gobjOneCardComLib.zlTransferAccountsMoney(mfrmMain, mlngModule, _
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
    If gobjOneCardComLib.zlSwapIsSucces(mfrmMain, mlngModule, intSwapType, lng卡类别ID, _
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
    If gobjOneCardComLib.zlReturnCashCheck(mfrmMain, mlngModule, _
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

