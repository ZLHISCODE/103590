VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinicThirdDelSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mfrmMain As Object
Private mlngModule As Long
Private mobjSquareCard As Object
Private mbln补充结算 As Boolean

'保持属性值的局部变量
Private mbln异步结算 As Boolean '局部复制
Private mlng病人ID As Long '局部复制
Private mstr姓名 As String '局部复制
Private mstr性别 As String '局部复制
Private mstr年龄 As String '局部复制
Private mstr门诊号 As String '局部复制
Private mstr住院号 As String '局部复制
Private mstr医疗付款方式 As String '局部复制
Private mlng卡类别ID As Long '局部复制
Private mbln退款验卡 As Boolean '局部复制
Private mbln转账及代扣 As Boolean '局部复制
Private mbln按交易退款 As Boolean '局部复制
Private mlng结帐ID As Long '局部复制
Private mlng冲销ID As Long '局部复制
Private mlng结算ID As Long '局部复制
'字段:类型,NO,结算方式,结算性质,冲预交,卡类别ID,结算卡序号,卡号,消费卡ID,交易流水号,交易说明,原结帐ID,关联交易ID
Private mrsBalance As ADODB.Recordset '局部复制
Private mrsDelFee As ADODB.Recordset '局部复制

Public Property Let 医疗付款方式(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医疗付款方式 = 5
    mstr医疗付款方式 = vData
End Property


Public Property Get 医疗付款方式() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医疗付款方式
    医疗付款方式 = mstr医疗付款方式
End Property



Public Property Let 住院号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.住院号 = 5
    mstr住院号 = vData
End Property


Public Property Get 住院号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.住院号
    住院号 = mstr住院号
End Property



Public Property Let 门诊号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊号 = 5
    mstr门诊号 = vData
End Property


Public Property Get 门诊号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊号
    门诊号 = mstr门诊号
End Property




Public Property Get 异步结算() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.异步结算
    异步结算 = mbln异步结算
End Property

Public Property Set rsDelFee(ByVal vData As ADODB.Recordset)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.rsDelFee = Form1
    Set mrsDelFee = vData
End Property

Public Property Get rsDelFee() As ADODB.Recordset
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.rsDelFee
    Set rsDelFee = mrsDelFee
End Property

Public Function InitData(frmMain As Object, ByVal lngModule As Long, ByRef objSquareCard As Object) As Boolean
    '初始化公共基础数据
    Set mfrmMain = frmMain
    mlngModule = lngModule
    Set mobjSquareCard = objSquareCard
    
    mbln异步结算 = zlDatabase.GetPara(304, glngSys) = "1"
    InitData = True
End Function

Public Property Let 补充结算(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.补充结算 = 5
    mbln补充结算 = vData
End Property

Public Property Get 补充结算() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.补充结算
    补充结算 = mbln补充结算
End Property

Public Property Set rsBalance(ByVal vData As ADODB.Recordset)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.rsBalance = Form1
    Set mrsBalance = vData
End Property

Public Property Get rsBalance() As ADODB.Recordset
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.rsBalance
    Set rsBalance = mrsBalance
End Property

Public Property Let 冲销ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.冲销ID = 5
    mlng冲销ID = vData
End Property

Public Property Get 冲销ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.冲销ID
    冲销ID = mlng冲销ID
End Property

Public Property Let 结算ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算ID = 5
    mlng结算ID = vData
End Property

Public Property Get 结算ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算ID
    结算ID = mlng结算ID
End Property

Public Property Let 结帐ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐ID = 5
    mlng结帐ID = vData
End Property

Public Property Get 结帐ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐ID
    结帐ID = mlng结帐ID
End Property

Public Property Get 转账及代扣() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.转账及代扣
    转账及代扣 = mbln转账及代扣
End Property

Public Property Let 退款验卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.退款验卡 = 5
    mbln退款验卡 = vData
End Property

Public Property Get 退款验卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.退款验卡
    退款验卡 = mbln退款验卡
End Property

Public Property Let 卡类别ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡类别ID = 5
    mlng卡类别ID = vData
    
    Dim objCard As Card
    If GetCard(objCard, False) = False Then Exit Property
    mbln退款验卡 = objCard.是否退款验卡
    mbln转账及代扣 = objCard.是否转帐及代扣
End Property

Public Property Get 卡类别ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡类别ID
    卡类别ID = mlng卡类别ID
End Property

Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mstr年龄 = vData
End Property

Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.年龄
    年龄 = mstr年龄
End Property

Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mstr性别 = vData
End Property

Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.性别
    性别 = mstr性别
End Property

Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mstr姓名 = vData
End Property

Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    姓名 = mstr姓名
End Property

Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mlng病人ID = vData
End Property

Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病人ID
    病人ID = mlng病人ID
End Property

Private Function GetDelFeeXml(Optional ByRef dblMoney As Double) As String
    '组织传入一卡通接口的费用明细XML串，格式如下：
    '    <TFLIST>  //退费列表
    '        <NO></NO>  // 退费单据
    '        <TFITEM>     //退费项
    '            <SerialNum>序号</SerialNum>
    '            …
    '        </ TFITEM >
    '    </TFLIST>
    '    ...
    '出参：
    '   dblMoney 本次费用合计
    Dim strXML As String, strPriorNO As String
    
    dblMoney = 0
    With mrsDelFee
        .Filter = ""
        .Sort = "NO,序号"
        
        strXML = "": strPriorNO = ""
        Do While Not .EOF
            If strPriorNO <> NVL(!NO) Then
                If strPriorNO <> "" Then
                    strXML = strXML & "    </TFITEM>" & vbCrLf
                    strXML = strXML & "  </TFLIST>" & vbCrLf
                End If
                strXML = strXML & "  <TFLIST>" & vbNewLine '退费列表
                strXML = strXML & "    <NO>" & NVL(!NO) & "</NO>" & vbCrLf '退费单据
                strXML = strXML & "    <TFITEM>" & vbCrLf '退费项
            End If
            
            strXML = strXML & "      <SerialNum>" & Val(NVL(!序号)) & "</SerialNum>" & vbCrLf '序号
            strPriorNO = NVL(!NO)
            dblMoney = dblMoney + Val(NVL(!实收金额))
            
            .MoveNext
        Loop
        
        strXML = strXML & "    </TFITEM>" & vbCrLf
        strXML = strXML & "  </TFLIST>"
    End With
    dblMoney = RoundEx(dblMoney, 6)
    GetDelFeeXml = strXML
End Function

Private Function GetDelBalanceXml(ByRef dblMoney As Double, Optional ByRef str原结帐IDs As String, _
    Optional ByVal lng原结帐ID As Long) As String
    '组织传入一卡通接口的退费结算数据XML串，格式如下：
    '    <TKLIST>//本次退款列表
    '        <TK>
    '            <TKFS>退款方式</TKFS>
    '            <TKJE>退款金额</TKJE>
    '            <JYLSH>原交易流水号</JYLSH>
    '            <JYSM>原交易说明</JYSM>
    '            <DJH>单据号</DJH>
    '        </TK>
    '        ....
    '    </TKLIST>
    '入参:
    '   lng原结帐ID 获取指定结算次数的数据，不传表示获取所有的
    '出参：
    '   dblMoney 本次退款合计
    '   str原结帐IDs 可能有多个原结帐ID
    Dim strXML As String
    
    dblMoney = 0: str原结帐IDs = ""
    With mrsBalance
        If lng原结帐ID = 0 Then
            .Filter = "类型=" & gBalanceType.gEM_一卡通 & " And 卡类别ID=" & mlng卡类别ID
        Else
            .Filter = "类型=" & gBalanceType.gEM_一卡通 & " And 卡类别ID=" & mlng卡类别ID & " And 原结帐ID=" & lng原结帐ID
        End If
        .Sort = "NO"
        strXML = "  <TKLIST>" & vbCrLf
        Do While Not .EOF
            strXML = strXML & "    <TK>" & vbCrLf
            strXML = strXML & "      <TKFS>" & NVL(!结算方式) & "</TKFS>" & vbCrLf
            strXML = strXML & "      <TKJE>" & Val(NVL(!冲预交)) & "</TKJE>" & vbCrLf
            strXML = strXML & "      <JYLSH>" & NVL(!交易流水号) & "</JYLSH>" & vbCrLf
            strXML = strXML & "      <JYSM>" & NVL(!交易说明) & "</JYSM>" & vbCrLf
            strXML = strXML & "      <DJH>" & NVL(!NO) & "</DJH>" & vbCrLf
            strXML = strXML & "    </TK>" & vbCrLf
            dblMoney = dblMoney + Val(NVL(!冲预交))
            
            If Val(NVL(!原结帐ID)) <> 0 Then
                If InStr("," & str原结帐IDs & ",", "," & Val(NVL(!原结帐ID)) & ",") = 0 Then
                    str原结帐IDs = str原结帐IDs & "," & Val(NVL(!原结帐ID))
                End If
            End If
            
            .MoveNext
        Loop
        strXML = strXML & "  </TKLIST>"
    End With
    dblMoney = RoundEx(dblMoney, 6)
    If str原结帐IDs <> "" Then str原结帐IDs = Mid(str原结帐IDs, 2)
    GetDelBalanceXml = strXML
End Function

Public Function ExecuteReturnCashCheck(ByRef blnDelDefaultCash_Out As Boolean, _
    Optional ByRef strDefaultDelBalance_Out As String) As Boolean
    '三方结算交易退现检查
    '出参:
    '   blnDelDefaultCash_Out-是否缺省退现：接口返回true时有效
    '   strDefaultDelBalance_Out-缺省退现方式,比如：支票，现金等
    Dim strCardNo As String, dblMoney As Double, lng原结帐ID As Long
    Dim strSwapNO As String, strSwapMemo As String, strExpendXML As String
    
    On Error GoTo ErrHandler
    blnDelDefaultCash_Out = False: strDefaultDelBalance_Out = ""
    strExpendXML = "<INPUT>" & vbCrLf & GetDelBalanceXml(dblMoney) & vbCrLf & "</INPUT>"
    mrsBalance.Filter = "类型=" & gBalanceType.gEM_一卡通 & " And 卡类别ID=" & mlng卡类别ID
    If mrsBalance.EOF Then
        MsgBox "未发现三方结算交易数据！", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNo = NVL(mrsBalance!卡号)
    strSwapNO = NVL(mrsBalance!交易流水号)
    strSwapMemo = NVL(mrsBalance!交易说明)
    lng原结帐ID = Val(NVL(mrsBalance!原结帐ID))
    
    'zlReturnCashCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strExpendXML As String, _
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
    '       strExpendXML    XML IN  10.35.90后才支持
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
    '       strExpendXML:10.35.90后才支持
    '        <OUTPUT>
    '            <TXZT>退现状态<TXZT>//NUMBER 1
    '               接口返回True时：表示是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但允许操作员操作退现
    '               接口返回False时: 表示是否允许强制退现，1-表示允许通过”强制退现”权限来控制是否允许退现，0-表示禁止通过“强制退现”权限来退现
    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
    '                    1.不允许返回三卡方的结算方式
    '                    2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
    '        </OUTPUT>
    '返回:退现合法,返回true,否则返回Flase
    If mobjSquareCard.zlReturnCashCheck(mfrmMain, mlngModule, mlng卡类别ID, strCardNo, _
        "3|" & lng原结帐ID, dblMoney, strSwapNO, strSwapMemo, strExpendXML, _
        blnDelDefaultCash_Out, strDefaultDelBalance_Out) = False Then Exit Function
    
    ExecuteReturnCashCheck = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteReturnCheck() As Boolean
    '退款检查
    Dim strExpendXML As String, dblMoney As Double
    Dim strCardNo As String, strPassWord As String
    Dim strSwapNO As String, strSwapMemo As String
    Dim str原结帐ID As String
    Dim strExpend As String
    
    On Error GoTo ErrHandler
    If mbln补充结算 Then
        mbln按交易退款 = GetParaConfig(Val("2-按交易流水号分别调用退款接口"))
        If mbln按交易退款 Then
            If ExecuteReturnCheckBySingle() = False Then Exit Function
        Else
            If ExecuteReturnCheckByMulti() = False Then Exit Function
        End If
        ExecuteReturnCheck = True: Exit Function
    End If
    
    strExpendXML = _
        "<TFDATA>" & vbCrLf & _
        "  <YCTF>0</YCTF>" & _
            GetDelFeeXml() & vbCrLf & _
            GetDelBalanceXml(dblMoney, str原结帐ID) & vbCrLf & _
        "</TFDATA>"
    mrsBalance.Filter = "类型=" & gBalanceType.gEM_一卡通 & " And 卡类别ID=" & mlng卡类别ID
    If mrsBalance.EOF Then
        MsgBox "未发现三方结算交易数据！", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNo = NVL(mrsBalance!卡号)
    strSwapNO = NVL(mrsBalance!交易流水号)
    strSwapMemo = NVL(mrsBalance!交易说明)

    'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strExpendXML As String) As Boolean
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
        '       strExpendXML    XML IN
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
    If mobjSquareCard.zlReturnCheck(mfrmMain, mlngModule, mlng卡类别ID, False, strCardNo, _
        "3|" & str原结帐ID, dblMoney, strSwapNO, strSwapMemo, strExpendXML) = False Then Exit Function

    If mbln退款验卡 Then
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
            Optional ByRef blnShowPatiInfor As Boolean = False, _
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
            '       strExpendXML-三方卡调用XML入参,目前格式如下:
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
        If mobjSquareCard.zlBrushCard(mfrmMain, mlngModule, Nothing, mlng卡类别ID, False, _
            mstr姓名, mstr性别, mstr年龄, dblMoney, strCardNo, strPassWord, True, True, False, True, _
            Nothing, False, False, strExpend) = False Then Exit Function
    End If
    ExecuteReturnCheck = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteReturnCheckBySingle() As Boolean
    '补充结算按交易退费检查
    Dim strExpendXML As String, dblMoney As Double
    Dim strCardNo As String
    Dim strSwapNO As String, strSwapMemo As String
    Dim str原结帐ID As String, strFeeXml As String
    Dim var原结帐ID As Variant, i As Long
    
    On Error GoTo ErrHandler
    mrsBalance.Filter = "类型=" & gBalanceType.gEM_一卡通 & " And 卡类别ID=" & mlng卡类别ID
    If mrsBalance.EOF Then
        MsgBox "未发现三方结算交易数据！", vbInformation, gstrSysName
        Exit Function
    End If
    strFeeXml = GetDelFeeXml()
    
    '收集原结帐ID
    Call GetDelBalanceXml(dblMoney, str原结帐ID)
    var原结帐ID = Split(str原结帐ID, ",")
    For i = 0 To UBound(var原结帐ID)
        strExpendXML = _
            "<TFDATA>" & vbCrLf & _
            "  <YCTF>0</YCTF>" & _
                strFeeXml & vbCrLf & _
                GetDelBalanceXml(dblMoney, str原结帐ID, var原结帐ID(i)) & vbCrLf & _
            "</TFDATA>"
        mrsBalance.Filter = "类型=" & gBalanceType.gEM_一卡通 & _
            " And 卡类别ID=" & mlng卡类别ID & " And 原结帐ID=" & var原结帐ID(i)
        strCardNo = NVL(mrsBalance!卡号)
        strSwapNO = NVL(mrsBalance!交易流水号)
        strSwapMemo = NVL(mrsBalance!交易说明)

        'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strExpendXML As String) As Boolean
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
        '       strExpendXML    XML IN
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
        If mobjSquareCard.zlReturnCheck(mfrmMain, mlngModule, mlng卡类别ID, False, strCardNo, _
            "6|" & str原结帐ID, dblMoney, strSwapNO, strSwapMemo, strExpendXML) = False Then Exit Function
    Next
    ExecuteReturnCheckBySingle = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteReturnCheckByMulti() As Boolean
    '补充结算多次交易一次退款检查
    Dim strExpendXML As String, dblMoney As Double
    Dim strCardNo As String
    Dim str原结帐ID As String
    
    On Error GoTo ErrHandler
    strExpendXML = _
        "<TFDATA>" & vbCrLf & _
        "  <YCTF>0</YCTF>" & _
            GetDelFeeXml() & vbCrLf & _
            GetDelBalanceXml(dblMoney, str原结帐ID) & vbCrLf & _
        "</TFDATA>"
        
    'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
        ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, ByVal strSwapNo As String, _
        ByVal strSwapMemo As String, ByRef strExpendXML As String) As Boolean
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
        '       strExpendXML    XML IN
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
    If mobjSquareCard.zlReturnCheck(mfrmMain, mlngModule, mlng卡类别ID, False, "", _
        "6|" & str原结帐ID, dblMoney, "", "", strExpendXML) = False Then Exit Function
    ExecuteReturnCheckByMulti = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteReturnMoney(cllBillPro As Collection, ByRef blnCommitData As Boolean, _
    Optional ByRef strErrMsg_out As String) As Boolean
    '调用接口进行退款操作
    '出参：
    '   blnCommitData 数据是否有提交，主要用于出现异常时判断
    '   strErrmsg_Out 错误信息，返回False有效
    Dim strExpendXML As String, dblMoney As Double
    Dim strCardNo As String, i As Integer
    Dim strSwapNO As String, strSwapMemo As String
    Dim str原结帐ID As String, lng关联交易ID As Long
    Dim blnTrans As Boolean, dblOutMoney As Double
    Dim cll结算方式 As Collection, strSwapExtendInfor As String, strTemp As String
    Dim bln正在交易 As Boolean, strErrMsg As String
    
    On Error GoTo ErrHandler
    strErrMsg_out = "": blnCommitData = False
    If mbln补充结算 Then
        If mbln按交易退款 Then
            If ExecuteReturnMoneyBySingle(cllBillPro, blnCommitData, strErrMsg_out) = False Then Exit Function
        Else
            If ExecuteReturnMoneyByMulti(cllBillPro, blnCommitData, strErrMsg_out) = False Then Exit Function
        End If
        ExecuteReturnMoney = True: Exit Function
    End If
    
    strExpendXML = "<INPUT>" & vbCrLf & GetDelBalanceXml(dblMoney, str原结帐ID) & vbCrLf & "</INPUT>"
    mrsBalance.Filter = "类型=" & gBalanceType.gEM_一卡通 & " And 卡类别ID=" & mlng卡类别ID
    If mrsBalance.EOF Then
        MsgBox "未发现三方结算交易数据！", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNo = NVL(mrsBalance!卡号)
    strSwapNO = NVL(mrsBalance!交易流水号)
    strSwapMemo = NVL(mrsBalance!交易说明)
    lng关联交易ID = Val(NVL(mrsBalance!关联交易ID))
    
    blnTrans = True
    zlExecuteProcedureArrAy cllBillPro, "保存费用单据", True

    '先按原结算方式保存结算信息
    If GetSwapReturnXMLExpend(strExpendXML, dblOutMoney, cll结算方式) = False Then Exit Function
    'Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算",原交易流水号,原交易说明,退款交易流水号,退款交易说明)
    For i = 1 To cll结算方式.Count
        Call SaveBalanceData(strCardNo, cll结算方式(i)(0), _
            lng关联交易ID, False, cll结算方式(i)(1), cll结算方式(i)(2))
    Next
            
    If mbln异步结算 Then '费用结算异步控制，先提交数据
        gcnOracle.CommitTrans: blnTrans = False
        blnCommitData = True
    End If
    
    Set cll结算方式 = Nothing
    strSwapExtendInfor = "3|" & mlng冲销ID: strTemp = strSwapExtendInfor
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
    If mobjSquareCard.zlReturnMoney(mfrmMain, mlngModule, mlng卡类别ID, False, strCardNo, _
        "3|" & str原结帐ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strExpendXML) = False Then
        If blnTrans Then gcnOracle.RollbackTrans
        
        bln正在交易 = True
        Call GetSwapReturnXMLExpend(strExpendXML, dblOutMoney, cll结算方式, bln正在交易, strErrMsg)
        If strErrMsg <> "" Then strErrMsg_out = strErrMsg
        If Not mbln异步结算 Then Exit Function
        
        '交易已确认失败，则删除病人预交记录
        If bln正在交易 = False Then
            Call DeleteBalanceRecord(mlng冲销ID, lng关联交易ID)
            blnCommitData = False
        End If
        Exit Function
    End If
    
    '更新交易结算信息
    If zlXML_ExistNode(strExpendXML, "OUTPUT") Then
        If GetSwapReturnXMLExpend(strExpendXML, dblOutMoney, cll结算方式) = False Then Exit Function
    End If
    If blnTrans = False Then gcnOracle.BeginTrans: blnTrans = True
    If cll结算方式 Is Nothing Then
        Call UpdateBalanceData(mlng冲销ID, strCardNo, strSwapNO, strSwapMemo, 2)
    Else
        'Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算",原交易流水号,原交易说明,退款交易流水号,退款交易说明)
        For i = 1 To cll结算方式.Count
            Call SaveBalanceData(strCardNo, cll结算方式(i)(0), _
                lng关联交易ID, (i = 1), cll结算方式(i)(1), cll结算方式(i)(2), 2)
        Next
    End If
    gcnOracle.CommitTrans: blnTrans = False
    blnCommitData = True
    
    '保存交易扩展信息
    If strSwapExtendInfor = strTemp Then strSwapExtendInfor = ""
    Call SaveSwapExpendInfo(mlng冲销ID, strCardNo, strSwapExtendInfor)
    
    If (RoundEx(dblOutMoney, 6) <> RoundEx(-1 * dblMoney, 6)) Then '如果实际退款金额不同，则当做异常处理
        strErrMsg_out = "本次应退金额(" & FormatEx(dblMoney, 6, , , 2) & _
            ")与实际退款金额(" & FormatEx(-1 * dblOutMoney, 6, , , 2) & "))不一致。"
        Exit Function
    End If
    ExecuteReturnMoney = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteReturnMoneyBySingle(cllBillPro As Collection, ByRef blnCommitData As Boolean, _
    Optional ByRef strErrMsg_out As String) As Boolean
    '补充结算按交易退费
    '出参：
    '   blnCommitData 数据是否有提交，主要用于出现异常时判断
    '   strErrmsg_Out 错误信息，返回False有效
    Dim strExpendXML As String, dblMoney As Double
    Dim strCardNo As String, lng原交易ID As Long
    Dim strSwapNO As String, strSwapMemo As String
    Dim str原结帐ID As String, var原结帐ID As Variant, i As Integer
    Dim blnTrans As Boolean, dblOutMoney As Double
    Dim cll结算方式 As Collection, strSwapExtendInfor As String, strTemp As String
    Dim bln正在交易 As Boolean, strErrMsg As String
    Dim lng结算序号 As Long, str结算方式 As String, str结帐IDs As String
        
    On Error GoTo ErrHandler
    lng结算序号 = -1 * mlng冲销ID
    str结帐IDs = mlng冲销ID & "," & mlng结算ID
    
    mrsBalance.Filter = "类型=" & gBalanceType.gEM_一卡通 & " And 卡类别ID=" & mlng卡类别ID
    If mrsBalance.EOF Then
        MsgBox "未发现三方结算交易数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    blnTrans = True
    zlExecuteProcedureArrAy cllBillPro, "保存费用单据", True
    
    '收集原结帐ID
    Call GetDelBalanceXml(dblMoney, str原结帐ID)
    var原结帐ID = Split(str原结帐ID, ",")
    For i = 0 To UBound(var原结帐ID)
        With mrsBalance
            strExpendXML = "<INPUT>" & vbCrLf & GetDelBalanceXml(dblMoney, str原结帐ID, var原结帐ID(i)) & vbCrLf & "</INPUT>"
            .Filter = "类型=" & gBalanceType.gEM_一卡通 & _
                " And 卡类别ID=" & mlng卡类别ID & " And 原结帐ID=" & var原结帐ID(i)
            str结算方式 = NVL(!结算方式)
            strSwapNO = NVL(!交易流水号)
            strSwapMemo = NVL(!交易说明)
            
            '结算方式|结算金额|结算号码|结算摘要
            str结算方式 = str结算方式 & "|" & -1 * dblMoney & "| | "
            Call SaveBalanceData("", str结算方式, , , , , , True)
    
            Call SaveThreeDelSwapInfo(lng结算序号, NVL(!原结帐ID), dblMoney, _
                NVL(!卡号), NVL(!交易流水号), NVL(!交易说明), False, False, _
                mlng卡类别ID, NVL(!交易流水号), NVL(!交易说明))
    
            If mbln异步结算 Then '费用结算异步控制，先提交数据
                gcnOracle.CommitTrans: blnTrans = False
                blnCommitData = True
            End If
            
            Set cll结算方式 = Nothing
            strSwapExtendInfor = "6|" & str结帐IDs: strTemp = strSwapExtendInfor
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
            If mobjSquareCard.zlReturnMoney(mfrmMain, mlngModule, mlng卡类别ID, False, NVL(!卡号), _
                "6|" & str原结帐ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strExpendXML) = False Then
                If blnTrans Then gcnOracle.RollbackTrans
                
                bln正在交易 = True
                Call GetSwapReturnXMLExpend(strExpendXML, dblOutMoney, cll结算方式, bln正在交易, strErrMsg)
                If strErrMsg <> "" Then strErrMsg_out = strErrMsg
                If Not mbln异步结算 Then Exit Function
                
                '交易已确认失败，则删除病人预交记录
                If bln正在交易 = False Then
                    Call DeleteBalanceRecord(0, 0, mlng卡类别ID, lng结算序号, True)
                    blnCommitData = False
                End If
                Exit Function
            End If
            
            '更新交易结算信息
            If zlXML_ExistNode(strExpendXML, "OUTPUT") Then
                If GetSwapReturnXMLExpend(strExpendXML, dblOutMoney, cll结算方式) = False Then Exit Function
            End If
            If blnTrans = False Then gcnOracle.BeginTrans: blnTrans = True
            If Not cll结算方式 Is Nothing Then
                If RoundEx(dblOutMoney, 6) <> RoundEx(-1 * dblMoney, 6) Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    strErrMsg_out = "应退金额(" & FormatEx(dblMoney, 6, , , 2) & ")与实际退款(" & _
                        FormatEx(-1 * dblOutMoney, 6, , , 2) & ")金额不相等，退款失败！"
                    Call DeleteBalanceRecord(0, 0, mlng卡类别ID, lng结算序号, True)
                    blnCommitData = False
                    Exit Function
                End If
                If cll结算方式.Count > 1 Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    strErrMsg_out = "退款失败，只能按原结算方式退回！"
                    Call DeleteBalanceRecord(0, 0, mlng卡类别ID, lng结算序号, True)
                    blnCommitData = False
                    Exit Function
                End If
                
                'Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算",原交易流水号,原交易说明,退款交易流水号,退款交易说明)
                Call SaveThreeDelSwapInfo(lng结算序号, NVL(!原结帐ID), dblMoney, _
                    NVL(!卡号), cll结算方式(1)(3), cll结算方式(1)(4), True, True, _
                    mlng卡类别ID, NVL(!交易流水号), NVL(!交易说明))
            Else
                Call SaveThreeDelSwapInfo(lng结算序号, NVL(!原结帐ID), dblMoney, _
                    NVL(!卡号), strSwapNO, strSwapMemo, True, True, _
                    mlng卡类别ID, NVL(!交易流水号), NVL(!交易说明))
            End If
            gcnOracle.CommitTrans: blnTrans = False
            blnCommitData = True
            
            If strSwapExtendInfor = strTemp Then strSwapExtendInfor = ""
            If strSwapExtendInfor <> "" Then
                lng原交易ID = Get原预交ID(NVL(!原结帐ID), mlng卡类别ID, NVL(!卡号), NVL(!交易流水号))
                Call SaveSwapExpendInfo(str结帐IDs, "", strSwapExtendInfor, lng原交易ID)
            End If
            gcnOracle.BeginTrans: blnTrans = True
        End With
    Next
    gcnOracle.CommitTrans: blnTrans = False
    blnCommitData = True
    
    '更新交易信息
    Call UpdateBalanceData(str结帐IDs, "", "", "", 2)
    ExecuteReturnMoneyBySingle = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get原预交ID(ByVal lng原结帐ID As Long, ByVal lng卡类别ID As Long, _
    ByVal str卡号 As String, ByVal str交易流水号 As String) As Long
    '获取原交易预交记录ID
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select ID From 病人预交记录" & _
        " Where 记录性质 = 3 And 结帐id = [1] And 卡类别id = [2]" & _
        "       And Nvl(卡号, '-') = Nvl([3], '-') And Nvl(交易流水号, '-') = Nvl([4], '-')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取原交易ID", lng原结帐ID, lng卡类别ID, str卡号, str交易流水号)
    If rsTemp.EOF Then
        MsgBox "未发现原三方结算交易数据！", vbInformation, gstrSysName
        Exit Function
    End If
    Get原预交ID = Val(NVL(rsTemp!ID))
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ExecuteReturnMoneyByMulti(cllBillPro As Collection, ByRef blnCommitData As Boolean, _
    Optional ByRef strErrMsg_out As String) As Boolean
    '补充结算多次交易一次退费
    '出参：
    '   blnCommitData 数据是否有提交，主要用于出现异常时判断
    '   strErrmsg_Out 错误信息，返回False有效
    Dim dblMoney As Double
    Dim i As Long, j As Integer, strExpend As String
    Dim blnTrans As Boolean, dblOutMoney As Double
    Dim cll结算方式 As Collection, strSwapExtendInfor As String
    Dim bln正在交易 As Boolean, strErrMsg As String
    Dim lng结算序号 As Long, cllDelDetail As Collection
    Dim str结算方式 As String, strInXML As String, strOutXML As String
    Dim strXMLExpend As String, blnCheckMoney As Boolean, blnFind As Boolean
    Dim lngRow As Long, strValue As String
    Dim str结帐IDs As String
   
    On Error GoTo ErrHandler
    lng结算序号 = -1 * mlng冲销ID
    str结帐IDs = mlng冲销ID & "," & mlng结算ID
    Call GetDelBalanceXml(dblMoney)
    
    mrsBalance.Filter = "类型=" & gBalanceType.gEM_一卡通 & " And 卡类别ID=" & mlng卡类别ID
    If mrsBalance.EOF Then
        MsgBox "未发现三方结算交易数据！", vbInformation, gstrSysName
        Exit Function
    End If
    str结算方式 = NVL(mrsBalance!结算方式)
        
    blnTrans = True
    zlExecuteProcedureArrAy cllBillPro, "保存费用单据", True
    
    '结算方式|结算金额|结算号码|结算摘要
    str结算方式 = str结算方式 & "|" & -1 * dblMoney & "| | "
    Call SaveBalanceData("", str结算方式, , , , , , True)
    
    '退款明细
    Set cllDelDetail = New Collection
    With mrsBalance
        Do While Not .EOF
            strInXML = strInXML & vbCrLf & _
                "<JS>" & vbCrLf & _
                "  <TKFS>" & NVL(!结算方式) & "</TKFS>" & vbCrLf & _
                "  <KH>" & NVL(!卡号) & "</KH>" & vbCrLf & _
                "  <JYLSH>" & NVL(!交易流水号) & "</JYLSH>" & vbCrLf & _
                "  <JYSM>" & NVL(!交易说明) & "</JYSM>" & vbCrLf & _
                "  <ZFJE>" & Val(NVL(!冲预交)) & "</ZFJE>" & vbCrLf & _
                "  <JSLX>" & 6 & "</JSLX>" & vbCrLf & _
                "  <ID>" & Val(NVL(!原结帐ID)) & "</ID>" & vbCrLf & _
                "</JS>"
            
            Call SaveThreeDelSwapInfo(lng结算序号, Val(NVL(!原结帐ID)), Val(NVL(!冲预交)), _
                NVL(!卡号), NVL(!交易流水号), NVL(!交易说明), False, False, _
                mlng卡类别ID, NVL(!交易流水号), NVL(!交易说明))
                
            cllDelDetail.Add Array(lng结算序号, Val(NVL(!原结帐ID)), Val(NVL(!冲预交)), _
                NVL(!卡号), NVL(!交易流水号), NVL(!交易说明))
        
            .MoveNext
        Loop
    End With
    
    If mbln异步结算 Then '费用结算异步控制，先提交数据
        gcnOracle.CommitTrans: blnTrans = False
        blnCommitData = True
    End If
    
    Set cll结算方式 = Nothing
    strXMLExpend = ""
    strInXML = "<JSLIST>" & strInXML & "</JSLIST>"
    'zlReturnMultiMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strInXML As String, _
    ByVal lng冲销ID As Long, ByRef strOutXml As String, ByRef strExpend As String, _
    Optional ByVal blnResolveXMLToRecord As Boolean, _
    Optional ByRef rsBalance_Out As ADODB.Recordset, Optional ByRef rsExpend_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易(多笔回退)
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strInXML-XML串:
    '            <JSLIST>
    '               <JS>
    '                    <TKFS>退款方式</TKFS>  //10.35.90后新增,退款方式,即结算方式.名称
    '                    <KH>卡号</KH>
    '                    <JYLSH>交易流水号</JYLX>
    '                    <JYSM>交易说明</JYSM>
    '                    <ZFJE>作废金额</ZFJE>
    '                    <JSLX>类型</JSLX>
    '                    <DJH>单据号</DJH>  //10.35.90新增:单据号,多单据收费时有用 ，存储在"医保结算明细"表中,主要是分单据保存
    '                    <ID></ID>
    '                </JS>
    '            </JSLIST>
    '       lng冲销ID-作废时的冲销ID(作废时或退费时有效，否则为0）;类型=6时，传入结算序号
    '       strExpend-无（暂留，待以后扩展)
    '    blnResolveXMLToRecord-是否解析XML串给记录集(rsBalance_Out,rsExpend_Out）
    '出参:
    '     strOutXML-返回XML串
    '         <JSLIST>
    '            <JS>
    '                <TKFS>退款方式</TKFS>   //10.35.90新增:退款方式即结算方式.名称
    '                <TKJE>退款金额</TKJE>   //10．35．90退款金额
    '                <KH>卡号</KH>
    '                <TKLSH>退款交易流水号</TKLSH>
    '                <TKSM>退款交易说明</TKSM>
    '                <ID></ID>   //类型=1时,预交ID;类型=2,6时，为原结帐ID
    '                <DJH>单据号</DJH>   //10．35．90后增加:单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '                <SFPTJS>是否普通结算</SFPTJS> //是否普通结算(1-普通结算;0-一卡通结算):为1时，在预交记录中不填写卡类别ID,不属于一卡通结算
    '            </JS>
    '            ...
    '        </JSLIST>
    '      strExpend-交易的扩展信息
    '       <Expends>
    '           <Expend>
    '               <XMMC>项目名称1</XMMC>
    '               <XMNR>项目内容2</XMNR>
    '           </Expend>
    '       </Expends>
    '    rsBalance_Out-结算信息集(blnResolveXMLToRecord=true时返回)字段（卡类别ID,消费卡,结算方式,交易金额,交易流水号,交易说明,单据号,单据性质,卡号,结算号码,结算摘要,是否普通结算)
    '    rsExpend_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
    '返回:函数返回    True:调用成功,False:调用失败
    '日期:2015-11-10
    '说明:
    '   目前只有结帐程序时有效（结帐退款),用于一次性处理同一卡类别的多笔三方交易作废
    If mobjSquareCard.zlReturnMultiMoney(mfrmMain, mlngModule, _
        mlng卡类别ID, False, strInXML, lng结算序号, strOutXML, strXMLExpend) = False Then
        If blnTrans Then gcnOracle.RollbackTrans
        
        If zlXML_ExistNode(strOutXML, "OUTPUT") Or zlXML_ExistNode(strOutXML, "JSLIST") Then
            bln正在交易 = True
            Call GetMultiSwapReturnXMLExpend(strOutXML, cll结算方式, dblOutMoney, bln正在交易, strErrMsg)
        End If
        If strErrMsg <> "" Then strErrMsg_out = strErrMsg
        If Not mbln异步结算 Then Exit Function
        
        '交易已确认失败，则删除病人预交记录
        If bln正在交易 = False Then
            Call DeleteBalanceRecord(0, 0, mlng卡类别ID, lng结算序号, True)
            blnCommitData = False
        End If
        Exit Function
    End If
    
    blnCheckMoney = zlXML_ExistNode(strOutXML, "OUTPUT")
    If blnCheckMoney Or zlXML_ExistNode(strOutXML, "JSLIST") Then
        If GetMultiSwapReturnXMLExpend(strOutXML, cll结算方式, dblOutMoney) = False Then Exit Function
    End If
    
    If Not cll结算方式 Is Nothing And blnCheckMoney Then
        If RoundEx(dblOutMoney, 6) <> RoundEx(dblMoney, 6) Then
            If blnTrans Then gcnOracle.RollbackTrans
            strErrMsg_out = "本次应退金额(" & FormatEx(dblMoney, 6, , , 2) & ")与实际退款金额(" & _
                FormatEx(dblOutMoney, 6, , , 2) & ")不一致，退款失败！"
            Call DeleteBalanceRecord(0, 0, mlng卡类别ID, lng结算序号, True)
            blnCommitData = False
            Exit Function
        End If
    End If
    
    If blnTrans = False Then gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To cllDelDetail.Count
        blnFind = False
        'Array(结算序号,结帐ID,金额,卡号,原交易流水号,原交易说明)
        If Not cll结算方式 Is Nothing Then
            For j = 1 To cll结算方式.Count
                'Array(结帐ID,金额,卡号,交易流水号,交易说明)
                If cllDelDetail(i)(1) = cll结算方式(j)(0) Then
                    Call SaveThreeDelSwapInfo(cllDelDetail(i)(0), cllDelDetail(i)(1), cllDelDetail(i)(2), _
                        cllDelDetail(i)(3), cll结算方式(j)(3), cll结算方式(j)(4), _
                        True, True, mlng卡类别ID, cllDelDetail(i)(4), cllDelDetail(i)(5))
                    blnFind = True: Exit For
                End If
            Next
        End If
        If blnFind = False Then
            Call SaveThreeDelSwapInfo(cllDelDetail(i)(0), cllDelDetail(i)(1), cllDelDetail(i)(2), _
                cllDelDetail(i)(3), cllDelDetail(i)(4), cllDelDetail(i)(5), _
                True, True, mlng卡类别ID, cllDelDetail(i)(4), cllDelDetail(i)(5))
        End If
    Next
    Call UpdateBalanceData(str结帐IDs, "", "", "", 2)
    gcnOracle.CommitTrans: blnTrans = False
    blnCommitData = False
    
    '扩展信息
    If strXMLExpend <> "" Then
        '兼容处理
        strXMLExpend = Replace(strXMLExpend, "EXPENDS", "Expends")
        strXMLExpend = Replace(strXMLExpend, "EXPEND", "Expend")
        
        strSwapExtendInfor = ""
        If zlXML_Init = False Then Exit Function
        If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
        Call zlXML_GetChildRows("Expends", "Expend", lngRow)
        For i = 0 To lngRow - 1
            Call zlXML_GetNodeValue("XMMC", i, strValue)
            strSwapExtendInfor = strSwapExtendInfor & "||" & strValue
            Call zlXML_GetNodeValue("XMNR", i, strValue)
            strSwapExtendInfor = strSwapExtendInfor & "|" & strValue
        Next
    End If
    If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
    Call SaveSwapExpendInfo(str结帐IDs, "", strSwapExtendInfor)
    ExecuteReturnMoneyByMulti = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSwapReturnXMLExpend(ByVal strExpendXML As String, _
    ByRef dblOutMoney As Double, ByRef cllBalance As Collection, _
    Optional ByRef bln正在交易 As Boolean, _
    Optional ByRef strErrMsg As String) As Boolean
    '功能：解析三方退款数据
    '入参：
    '   strExpendXML:XML串
    '    <OUTPUT>
    '        <JYZT>交易状态</JYZT>//接口返回False时，此参数有效:0-交易调用失败;1-交易正在处理中
    '        <ERRMSG>错误信息</ERRMSG> //为空时，不提示，非空时，提示
    '        <TKLIST>
    '            <TK>
    '                <TKFS>退款方式</TKFS>
    '                <TKJE>结算金额</TKJE>
    '                <JYLSH>退款交易流水号</JYLSH>
    '                <JYSM>退款交易说明</JYSM>
    '                <YJYLSH>原交易流水号</JYLSH>
    '                <YJYSM>原交易说明</JYSM>
    '                <DJH>单据号</DJH>
    '                <SFPTJS>是否普通结算</SFPTJS>
    '            </TK>
    '            …
    '        </TKLIST>
    '    </OUTPUT>
    '   blnDelMoney - 是否对金额取相反数
    '出参：
    '   dblOutMoney - 实际退款金额
    '   cllBalance - 结算数据，格式：Array("结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算",原交易流水号,原交易说明,退款交易流水号,退款交易说明)
    Dim lngCount As Long, strValue As String
    Dim i As Integer, strBalance As String
    Dim str原交易流水号 As String, str原交易说明 As String
    Dim str退交易流水号 As String, str退交易说明 As String
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    Set cllBalance = New Collection
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strExpendXML, False) = False Then Exit Function
    
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
        str退交易流水号 = strValue '退款交易流水号
        Call zlXML_GetNodeValue("JYSM", i, strValue)
        str退交易说明 = strValue   '退款交易说明
        
        cllBalance.Add Array(strBalance, str原交易流水号, str原交易说明, str退交易流水号, str退交易说明)
    Next
    GetSwapReturnXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function UpdateBalanceData(ByVal str结帐IDs As String, _
    str卡号 As String, str交易流水号 As String, str交易说明 As String, _
    Optional int校对标志 As Integer = 0) As Boolean
    '更新三方交易流水号和流水说明
    Dim strSQL As String
    
    strSQL = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSQL = strSQL & "" & mlng卡类别ID & ","
    '  消费卡_In     Number,
    strSQL = strSQL & "" & "0" & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSQL = strSQL & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSQL = strSQL & "'" & str结帐IDs & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type
    strSQL = strSQL & "'" & str交易说明 & "',"
    '预交款缴款_In Number := 0
    strSQL = strSQL & "" & "0" & ","
    '退费标志 :1-退费;0-付费
    strSQL = strSQL & "0,"
    '校对标志
    strSQL = strSQL & "" & IIf(int校对标志 = 0, "NULL", int校对标志) & ")"
    zlDatabase.ExecuteProcedure strSQL, "更新交易信息"
    UpdateBalanceData = True
End Function

Private Function SaveBalanceData(ByVal str卡号 As String, ByVal str结算方式 As String, _
    Optional ByVal lng关联交易ID As Long, Optional ByVal bln删除原结算 As Boolean, _
    Optional ByVal str交易流水号 As String, Optional ByVal str交易说明 As String, _
    Optional ByVal byt校对标志 As Byte = 1, Optional ByVal bln补结算 As Boolean) As String
    '保存退款结算数据
    '入参：
    '   str结算方式 格式为:"结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算"
    Dim strSQL As String
    
    If bln补结算 Then
        'Zl_费用补充结算_完成退费_S
        strSQL = "Zl_费用补充结算_完成退费_S("
        '  结算id_In     In 费用补充记录.结算id%Type,
        strSQL = strSQL & "" & mlng结算ID & ","
        '  结算方式_In   Varchar2,格式:结算方式|结算金额|结算号码|结算摘要
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & mlng卡类别ID & ","
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & str卡号 & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & str交易流水号 & "',"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & str交易说明 & "',"
        '  误差金额_In   门诊费用记录.实收金额%Type := Null
        strSQL = strSQL & "" & 0 & ","
        '  完成结算_In Number:=0:1-完成补充结算;0-未完成补充结算
        strSQL = strSQL & "" & 0 & ","
        '  冲预交_In         病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "" & 0 & ","
        '  校对标志_In       病人预交记录.校对标志%Type := 0
        strSQL = strSQL & "" & byt校对标志 & ")"
        zlDatabase.ExecuteProcedure strSQL, "保存三方卡结算数据"
        SaveBalanceData = True
        Exit Function
    End If
    
    'Zl_门诊退费结算_Modify(
    strSQL = "Zl_门诊退费结算_Modify_S("
    '  操作类型_In      Number,
    strSQL = strSQL & "" & 5 & ","
    '  病人id_In        门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mlng病人ID & ","
    '  姓名_In          病人预交记录.姓名%Type,
    strSQL = strSQL & "'" & mstr姓名 & "',"
    '  性别_In          病人预交记录.性别%Type,
    strSQL = strSQL & "'" & mstr性别 & "',"
    '  年龄_In          病人预交记录.年龄%Type,
    strSQL = strSQL & "'" & mstr年龄 & "',"
    '  门诊号_In        病人预交记录.门诊号%Type,
    strSQL = strSQL & "'" & mstr门诊号 & "',"
    '  住院号_In        病人预交记录.住院号%Type,
    strSQL = strSQL & "'" & mstr住院号 & "',"
    '  付款方式名称_In  病人预交记录.付款方式名称%Type,
    strSQL = strSQL & "'" & mstr医疗付款方式 & "',"
    '  冲销id_In        病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & mlng冲销ID & ","
    '  结算方式_In      Varchar2,
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  冲预交_In        病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  卡类别id_In      病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & mlng卡类别ID & ","
    '  卡号_In          病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & str卡号 & "',"
    '  交易流水号_In    病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '  交易说明_In      病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '  缴款_In          病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  找补_In          病人预交记录.找补%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  误差金额_In      门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  完成退费_In      Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  原结帐id_In      病人预交记录.结帐id%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  剩余转预交_In    Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  缺省结算方式_In  结算方式.名称%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  冲预交病人ids_In Varchar2 := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  关联交易id_In    病人预交记录.关联交易id%Type := Null,
    strSQL = strSQL & "" & IIf(lng关联交易ID = 0, "NULL", lng关联交易ID) & ","
    '  删除原结算_In    Number := 0,
    strSQL = strSQL & "" & IIf(bln删除原结算, "1", "0") & ","
    '  校对标志_In      病人预交记录.校对标志%Type := 0
    strSQL = strSQL & "" & byt校对标志 & ")"
    zlDatabase.ExecuteProcedure strSQL, "保存三方卡结算数据"
    SaveBalanceData = True
End Function

Public Function GetCard(ByRef objCard As Card, Optional ByVal bln仅启用 As Boolean = True) As Boolean
    '根据卡类别ID获取卡类别信息
    On Error GoTo ErrHandler
    Set objCard = Nothing
    If gobjSquare.objOneCardComLib Is Nothing Then Exit Function
    
    'zlGetCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As Card)
    If mobjSquareCard.zlGetCard(mlng卡类别ID, False, objCard) = False Then Exit Function
    If Not objCard Is Nothing Then
        If bln仅启用 And Not objCard.启用 Then Set objCard = Nothing
    End If
    GetCard = Not objCard Is Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetParaConfig(ByVal intPara As Integer) As Boolean
    Dim strErrMsg As String, strExpendXML As String
    
    On Error GoTo ErrHandler
    'ZlGetParaConfig(ByVal frmMain As Object, _
        ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, ByVal intPara As Integer, _
        Optional strErrMsg As String, Optional strExpend As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:获取接口参数
        '入参: frmMain-调用的主窗体
        '       intPara: 包含如下值
        '                1-刷卡和支付在同一页面:true-新模式；False-旧模式
        '                2-按交易流水号分别调用退款接口
        '                3-自助系统调用刷卡接口
        '                4-禁止门诊线下医保或多种结算方式收款
        '       strExpend-扩展参数，暂留，现传为空
        '出参:strErrMsg-返回的错误信息
        '返回:函数返回True:调用成功,False:调用失败
    GetParaConfig = mobjSquareCard.ZlGetParaConfig(mfrmMain, mlng卡类别ID, False, intPara, strErrMsg, strExpendXML)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveSwapExpendInfo(ByVal str结帐IDs As String, _
    ByVal str卡号 As String, strExpend As String, Optional lng原交易ID As Long) As Boolean
    '保存三方交易扩展信息
    '入参:
    '       str结帐IDs-结帐ID
    '       lng原交易ID-多笔退调用多次接口时传入
    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long
    Dim str交易信息 As String, strTemp As String
    Dim blnTrans As Boolean, cllPro As New Collection

    On Error GoTo ErrHandler
    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
    varData = Split(strExpend, "||")
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                    str交易信息 = Mid(str交易信息, 3)
                    'Zl_三方结算交易_Insert
                    strSQL = "Zl_三方结算交易_Insert("
                    '卡类别id_In 病人预交记录.卡类别id%Type,
                    strSQL = strSQL & "" & mlng卡类别ID & ","
                    '消费卡_In   Number,
                    strSQL = strSQL & "" & "0" & ","
                    '卡号_In     病人预交记录.卡号%Type,
                    strSQL = strSQL & "'" & str卡号 & "',"
                    '结帐ids_In  Varchar2,
                    strSQL = strSQL & "'" & str结帐IDs & "',"
                    '交易信息_In Varchar2:交易项目|交易内容||...
                    strSQL = strSQL & "'" & str交易信息 & "',"
                    '预交款缴款_In Number := 0
                    strSQL = strSQL & "0" & ","
                    '结算方式_In   病人预交记录.结算方式%Type := Null
                    strSQL = strSQL & "NULL" & ","
                    '原交易id_In 病人预交记录.Id%Type := Null
                    strSQL = strSQL & IIf(lng原交易ID = 0, "NULL", lng原交易ID) & ")"
                    zlAddArray cllPro, strSQL
                    str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
        End If
    Next
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSQL = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & mlng卡类别ID & ","
        '消费卡_In   Number,
        strSQL = strSQL & "" & "0" & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSQL = strSQL & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSQL = strSQL & "'" & str结帐IDs & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSQL = strSQL & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSQL = strSQL & "0" & ","
        '结算方式_In   病人预交记录.结算方式%Type := Null
        strSQL = strSQL & "NULL" & ","
        '原交易id_In 病人预交记录.Id%Type := Null
        strSQL = strSQL & IIf(lng原交易ID = 0, "NULL", lng原交易ID) & ")"
        zlAddArray cllPro, strSQL
    End If
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, "保存三方结算交易扩展信息"
    blnTrans = False
    
    SaveSwapExpendInfo = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DeleteBalanceRecord(ByVal lng结帐ID As Long, _
    ByVal lng关联交易ID As Long, Optional ByVal lng卡类别ID As Long, _
    Optional ByVal lng结算序号 As Long, Optional ByVal blnMultiDel As Boolean)
    '三方交易调用失败删除结算记录
    '入参：
    '   blnMultiDel 是否多笔退款
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_病人结算记录_Delete(
    strSQL = "Zl_病人结算记录_Delete("
    '  结帐id_In     病人预交记录.结帐id%Type := Null,
    strSQL = strSQL & "" & ZVal(lng结帐ID) & ","
    '  关联交易id_In 病人预交记录.关联交易id%Type := Null,
    strSQL = strSQL & "" & ZVal(lng关联交易ID) & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & ZVal(lng卡类别ID) & ","
    '  结算序号_In   病人预交记录.结算序号%Type := Null,
    strSQL = strSQL & "" & ZVal(lng结算序号) & ","
    '  多笔退款_In   Number := 0
    strSQL = strSQL & "" & IIf(blnMultiDel, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, "删除结算记录"
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function SaveThreeDelSwapInfo(ByVal lng结算序号 As Long, ByVal lng原结帐ID As Long, _
    ByVal dblMoney As Double, ByVal str卡号 As String, ByVal str交易流水号 As String, _
    ByVal str交易说明 As String, ByVal blnUpdate As Boolean, ByVal bln退款成功 As Boolean, _
    ByVal lng卡类别ID As Long, ByVal str原交易流水号 As String, ByVal str原交易说明 As String) As Boolean
    '保存三方退款信息
    Dim strSQL As String

    'Zl_三方退款信息_Insert(
    strSQL = "Zl_三方退款信息_Insert("
    '  结帐id_In     三方退款信息.结帐id%Type,
    strSQL = strSQL & "" & lng结算序号 & ","
    '  记录id_In     三方退款信息.记录id%Type,
    strSQL = strSQL & "" & lng原结帐ID & ","
    '  金额_In       三方退款信息.金额%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  卡号_In       三方退款信息.卡号%Type,
    strSQL = strSQL & "'" & str卡号 & "',"
    '  交易流水号_In 三方退款信息.交易流水号%Type,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '  交易说明_In   三方退款信息.交易说明%Type,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '  操作类型_In   Number := 0,--0=新增,1=更新信息
    strSQL = strSQL & "" & IIf(blnUpdate, 1, 0) & ","
    '  是否未退_In   三方退款信息.是否未退%Type := 0,
    strSQL = strSQL & "" & IIf(bln退款成功, 0, 1) & ","
    '  是否转帐_In   三方退款信息.是否转帐%Type := 0,
    strSQL = strSQL & "" & 0 & ","
    '  卡类别id_In   三方退款信息.卡类别id%Type := Null,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  原交易流水号_In 三方退款信息.原交易流水号%Type := Null,
    strSQL = strSQL & "'" & str原交易流水号 & "',"
    '  原交易说明_In   三方退款信息.原交易说明%Type := Null
    strSQL = strSQL & "'" & str原交易说明 & "')"
    zlDatabase.ExecuteProcedure strSQL, "保存三方退款信息"
End Function

Private Function GetMultiSwapReturnXMLExpend(ByVal strXMLExpend As String, ByRef cllBalance As Collection, _
    Optional ByRef dblOutMoney As Double, Optional ByRef bln正在交易 As Boolean, _
    Optional ByRef strErrMsg As String) As Boolean
    '功能：解析三方退款数据
    '入参：
    '   strXMLExpend:XML串
    '    <OUTPUT>
    '        <JYZT>交易状态</JYZT>//接口返回False时，此参数有效:0-交易调用失败;1-交易正在处理中
    '        <ERRMSG>错误信息</ERRMSG> //为空时，不提示，非空时，提示
    '        <JSLIST>
    '            <JS>
    '                <KH>卡号</KH>
    '                <TKFS>退款方式</TKFS>
    '                <TKJE>退款金额</TKJE>
    '                <TKLSH>退款交易流水号</TKLSH>
    '                <TKSM>退款交易说明</TKSM>
    '                <ID></ID>
    '                <DJH>单据号</DJH>
    '                <SFPTJS>是否普通结算</SFPTJS>
    '            </JS>
    '            …
    '        </JSLIST>
    '    </OUTPUT>
    '   blnDelMoney - 是否对金额取相反数
    '出参：
    '   cllBalance - 结算数据，格式：Array(结帐ID,金额,卡号,交易流水号,交易说明)
    Dim lngCount As Long, strValue As String, i As Integer
    Dim str交易流水号 As String, str交易说明 As String
    Dim lng结帐ID As Long, strCardNo As String, dblMoney As Double
    
    On Error GoTo ErrHandler
    dblOutMoney = 0: bln正在交易 = False: strErrMsg = ""
    Set cllBalance = New Collection
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    
    Call zlXML_GetNodeValue("OUTPUT/JYZT", , strValue)
    bln正在交易 = Val(strValue) = 1 '交易状态
    Call zlXML_GetNodeValue("OUTPUT/ERRMSG", , strValue)
    strErrMsg = strValue '错误信息
    
    '结算信息
    Call zlXML_GetRows("JSLIST/JS", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("TKJE", i, strValue)
        dblMoney = Val(strValue) '退款金额
        dblOutMoney = dblOutMoney + dblMoney
        Call zlXML_GetNodeValue("ID", i, strValue)
        lng结帐ID = Val(strValue)   '结帐ID
        Call zlXML_GetNodeValue("KH", i, strValue)
        strCardNo = strValue '卡号
        Call zlXML_GetNodeValue("TKLSH", i, strValue)
        str交易流水号 = strValue '交易流水号
        Call zlXML_GetNodeValue("TKSM", i, strValue)
        str交易说明 = strValue   '交易说明
        
        cllBalance.Add Array(lng结帐ID, dblMoney, strCardNo, str交易流水号, str交易说明)
    Next
    GetMultiSwapReturnXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

