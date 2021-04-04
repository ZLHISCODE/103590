VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThirdSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*********************************************************************************************************************************************
'功能:执行一卡通三方交易的相关接口
'函数:
'     1.公共接口(基础部分)
'     1.1. zlInitCompents:初始化部件接口
'     1.2  zlGetThreeSwapExpendSQLToCollect:获取保证扩展信息的相关SQL
'     1.3  zlGetBalanceItemsFromBalaceDatas:根据指定的卡类别ID,获取结算明细项
'     1.4  **zlRemoveBalanceItemsFromBalaceDatas:根据指定的卡类别,移除对应的结自从数据数据
'     1.5  zlGetCardFromCardType:根据卡类别,获取卡对象
'     1.6  zlRemoveBalanceDataFromItems:根据结算项信息，从结算数据根据结帐信息集合中移除
'     1.7  zlGetBalanceCancelSaveSQLFromItems:，获取对应的作废SQL集
'     1.8  zlReGetPayCards:重新加载有效的支付卡
'     2.新一卡通消费相关接口:
'     2.1   zlThird_Payment_IsValied:扣款交易检查(同步调用了:zlBrushCard-->PaymentCheck-->zlGetAccountMoney接口)
'     2.2   zlThird_Payment:执行三方接口支付(调用 ：zlPaymentMoney)
'     2.3   zlThird_ReturnMoney_IsValied：三方卡退款合法性检查
'     2.4   zlThird_TransferAccounts:执行转帐交易接口
'     2.5   zlThirdReturnCashCheck:退现合法性检查
'     2.6   zlThirdSwapIsSwapNOCall:判断是否按交易流水号调用接口
'     2.7   zlThird_IsCancelFromItems:根据项目判断当前是否允许作废或删除
'     4.医保相关处理
'     4.1  zlGetYbPara:获取医保对数
'     4.2  zlInsure_InitPara:初始化医保参数
'     4.3  zlInsure_BalanceUpdate:医保结算数据更新
'     4.4  zlInsure_Check:医保预结算和医保结帐是否匹配
'     5 消费卡处理
'     5.1 zlSquare_AddBalanceDatas-向结算数据信息中加入消费卡结算信息
'     5.2 zlSquare_ReturnMoneySQL-获取执行消费卡退款处理的SQL
'编制:刘兴洪
'日期:2018-03-26 15:06:02
'*********************************************************************************************************************************************
Private mstrOnCardPrivs As String
Private mfrmMain As Object, mlngModule As Long
Private mobjICCard As clsICCard
Private mbln结算异步控制 As Boolean
Private mbyt刷卡缺省金额操作 As Byte
Private mbln消费卡退费验卡 As Boolean '消费卡退费时是否刷卡验证
Private mrs结算方式 As ADODB.Recordset
Private mobjPayCards As Cards '当前有效的支付方式

Public Property Set objPayCards(ByVal vData As Cards)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objPayCards = Form1
    Set mobjPayCards = vData
End Property
Public Property Get objPayCards() As Cards
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objPayCards
    Set objPayCards = mobjPayCards
End Property

Public Function zlCheckIsStartFromCard(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的卡是否被启用
    '入参:objCard-指定的卡
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-24 16:59:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Card
    
    On Error GoTo errHandle
    If mobjPayCards Is Nothing Or objCard Is Nothing Then Exit Function
    For Each objTemp In mobjPayCards
        If (objTemp.接口序号 = objCard.接口序号 And objTemp.消费卡 = objCard.消费卡) _
        Or (objCard.接口序号 = 0 And objCard.结算方式 = objTemp.结算方式) Then zlCheckIsStartFromCard = True: Exit Function
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInitCompents(ByVal frmMain As Object, ByVal lngModule As Long, objICCard As clsICCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化部件
    '入参:objICCard-IC卡对象
    '出参:
    '返回:初始化成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 17:00:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain: mlngModule = lngModule
    Set mobjICCard = objICCard
    mbln结算异步控制 = Val(zlDatabase.GetPara("费用结算异步控制", glngSys)) = 1
    mbyt刷卡缺省金额操作 = Val(zlDatabase.GetPara("刷卡缺省金额操作", glngSys, 1151, "0"))  '86853
    mbln消费卡退费验卡 = zlDatabase.GetPara(282, glngSys) = "1"

    zlInitCompents = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlThird_IsCancelFromItems(ByVal objCurItem As clsBalanceItem) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据项目，判断该项目是否允许取消操作
    '入参:  frmMain-调用的主窗体
    '       lngModule-模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn 收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
    '       strExpend-扩展参数，暂留，现传为空
    '出参:intSwapStatus_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
    '     strXMLExpend-待以后扩展
    '返回：接口调用成功返回true,否则返回Flase
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intSwapStaus As Integer, objCard As Card, objItems As clsBalanceItems
    Dim strErrMsg As String
    
    On Error GoTo errHandle
    
    If objCurItem Is Nothing Then Exit Function
    
    Set objCard = objCurItem.objCard
    
    If objCard.消费卡 Or objCurItem.卡类别ID <= 0 Then Exit Function
 
  
    If zlThird_IsSwapIsSucces(objCurItem, intSwapStaus, strErrMsg) Then
        MsgBox "注意:" & vbCrLf & _
               "    " & objCard.名称 & " 已经交易成功，不允许作废或删除，你必须退款操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If intSwapStaus <> 0 Then
        If strErrMsg <> "" Then
            MsgBox "注意:" & vbCrLf & _
                   "    " & objCard.名称 & " 交易正在进行中,不能删除或作废,详细错误信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Else
            MsgBox "注意:" & vbCrLf & _
                   "    " & objCard.名称 & " 交易正在进行中，不允许作废或删除!", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    zlThird_IsCancelFromItems = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThird_IsSwapIsSucces(ByVal objCurItem As clsBalanceItem, ByRef intSwapStatu_Out As Integer, ByRef strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断交易是否调用成功
    '入参: objItems-当前结算信息集
    '出参:intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
    '     strErrMsg_out-返回的错误信息
    '返回：接口调用成功返回true,否则返回Flase
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intSwapType As Integer, dblMoney As Double, strXML As String, strXMLExpend As String
    Dim objCard As Card, objItem As clsBalanceItem, objTemp As clsBalanceItem
    Dim strBalanceIDs As String
  
  
    If objCurItem Is Nothing Then Exit Function
    
    dblMoney = objCurItem.结算金额
     
    Set objCard = objCurItem.objCard
    
    On Error GoTo errHandle
    
   '0-扣款;1-退款；2-转帐
    intSwapType = IIf(objCurItem.是否退款, IIf(objCurItem.是否转帐, 2, 1), 0)
    strBalanceIDs = "1|" & objCurItem.ID
    
    'zlSwapIsSucces(ByVal frmMain As Object, ByVal lngModule As Long, ByVal intSwapType As Integer, _
    '    ByVal lngCardTypeID As Long, ByVal strBalanceIDs As String, ByVal strCardNO As String, ByVal dblSwapMoney As Double, _
    '    ByRef intSwapStatus_Out As Integer, Optional strErrMsg As String, Optional strXMLExpend As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:判断交易是否成功（10.35.90）
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
    '    '日期:2013-06-15 20:22:51
    '    '说明:
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = "": strXML = ""
    strXML = strXML & "    <TKLIST>" & vbCrLf
    strXML = strXML & "        <TK>" & vbCrLf
    strXML = strXML & "             <JYLSH>" & objCurItem.交易流水号 & "</JYLSH>" & vbCrLf
    strXML = strXML & "             <KH>" & objCurItem.卡号 & "</KH>" & vbCrLf
    strXML = strXML & "             <JE>" & objCurItem.结算金额 & "</JE> " & vbCrLf
    strXML = strXML & "        </TK>" & vbCrLf
    strXML = strXML & "    </TKLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "</INPUT>" & vbCrLf
    strXMLExpend = strXMLExpend & strXML
    strXMLExpend = strXMLExpend & "</INPUT>" & vbCrLf
    
    zlThird_IsSwapIsSucces = gobjSquare.objSquareCard.zlSwapIsSucces(mfrmMain, mlngModule, intSwapType, objCard.接口序号, strBalanceIDs, objCurItem.卡号, dblMoney, intSwapStatu_Out, strErrMsg_out, strXMLExpend)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThird_Payment(ByVal objCard As Card, ByVal objPati As clsPatiInfo, ByRef cllBillPro As Collection, _
    ByVal objCurItem As clsBalanceItem, ByRef objCurItem_Out As clsBalanceItem, Optional blnSaveed_Out As Boolean, _
    Optional ByVal int险类 As Integer, Optional ByVal bln电子票据 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行三方接口支付
    '入参:
    '    objCard-当前支付的卡对象
    '    objPati-病人对象集
    '    lng结帐ID-结帐ID
    '    cllBillPro-执行的相关过程
    '    lng结帐ID-结帐ID
    '    objBalanceInfor-当前结帐数据
    '    objCurItem-结算数据集
    '出参:objCurItem_Out-结帐返回数据(可能支付后结算信息发生了相关变化)
    '     objBalanceInfor-结帐返回后更新的结帐数据列表
    '     blnSaveed_Out-是否发生数据保存
    '返回:三方接口交易调用成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 17:02:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim rsBalance As ADODB.Recordset, rsExpend As ADODB.Recordset
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str结帐IDs As String, i As Long, strSQL As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim str结算方式  As String, dblMoney As Double
    Dim lng预交ID As Long, lngCount As Long, lng关联交易ID As Long
    Dim str卡号 As String, intStatus_Out As Integer, strErrMsg_out As String
    Dim blnExcuteThird As Boolean
    
    On Error GoTo errHandle
    
    blnSaveed_Out = False
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then zlThird_Payment = True: Exit Function
    If objCurItem Is Nothing Then Exit Function
    
    
    str卡号 = objItem.卡号
    dblMoney = objItem.结算金额
    If dblMoney = 0 Then Exit Function
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    If zlGetSaveSQLfromItem(objPati, objCurItem, IIf(mbln结算异步控制, 1, 0), cllPro, True, , int险类, bln电子票据) = False Then Exit Function

ReDO:
    blnTrans = True: blnExcuteThird = False
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption & "-执行三方接口", Not mbln结算异步控制
    If mbln结算异步控制 Then
        blnTrans = False: blnSaveed_Out = True
        Set cllUpdate = New Collection
        If zlGetSaveSQLfromItem(objPati, objCurItem, 0, cllUpdate, False, , int险类, bln电子票据) = False Then Exit Function
        blnTrans = True
        zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, True
    End If
    
    blnExcuteThird = True
    
    'Public Function zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, ByVal strPrepayNos As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String, _
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
    '    strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '    dblMoney-支付金额
    '    blnResolveXMLToRecord-是否解析XML串给记录订(rsBalance_Out,rsExpend_Out
    '出参:strSwapGlideNO-交易流水号(XML中返回结算信息时,交易流水号，以xml为准)
    '    strSwapMemo-交易说明(XML中返回返回结算信息时,交易以xml为准)
    '    rsBalance_Out-结算信息集(blnResolveXMLToRecord=true时返回)字段（卡类别ID,消费卡,结算方式,交易金额,交易流水号,交易说明,单据号,单据性质,卡号,结算号码,结算摘要,是否普通结算)
    '    rsExpend_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
    '    intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
    '    strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
    '    strSwapExtendInfor -交易扩展信息,有两种方式:
    
    '    1)  格式1为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '    2)  格式2(10.35.90以后才支持),XML格式
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
    '                    <JYZY>结算摘要</JYZY>   Varchar2 100
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
    
    If gobjSquare.objSquareCard.zlPaymentMoney(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, objItem.卡号, _
         str结帐IDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, True, rsBalance, rsExpend, intStatus_Out, strErrMsg_out) = False Then
        
        blnExcuteThird = False
        If blnTrans Then gcnOracle.RollbackTrans
        blnTrans = False
        'intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
        'strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
        If strErrMsg_out <> "" Then
            MsgBox "交易正在进行中，原因如下:" & vbCrLf & strErrMsg_out, vbOKOnly + vbOKOnly + vbInformation, gstrSysName
        End If
        
        If blnSaveed_Out Then Set cllBillPro = New Collection
        If intStatus_Out = 1 Then   '1-交易正在处理中
            If blnSaveed_Out Then
                objCurItem.是否保存 = True
                objCurItem.是否允许编辑 = False
                objCurItem.是否允许删除 = True
                objCurItem.关联交易ID = lng关联交易ID
                Set objCurItem_Out = objCurItem
            End If
            Exit Function
        End If
        
        objCurItem.关联交易ID = lng关联交易ID
        If mbln结算异步控制 Then
            Set cllUpdate = New Collection
            Call zlGetDeleteSQL(objCurItem.单据号, 0, cllUpdate)
            blnTrans = True
            zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption
            blnTrans = False
        End If
        Set objCurItem_Out = New clsBalanceItems
        Exit Function
    End If
    
    blnExcuteThird = False
    Set objCurItem_Out = objCurItem
    
    '更新数据
    If Not rsBalance Is Nothing Then
        rsBalance.Filter = 0
        If rsBalance.RecordCount > 1 Then
            MsgBox "当前支持接口返回了多种结算信息，充值不支持多种结算方式!", vbOKOnly, gstrSysName
            Exit Function
        End If
        Set cllUpdate = New Collection
        With rsBalance
            If Not .EOF Then
                If -1 * Val(Nvl(rsBalance!交易金额)) <> objCurItem.结算金额 Then
                    MsgBox "当前支付的金额与当前传的金额不一致！", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                With objCurItem
                    .卡类别ID = IIf(Val(Nvl(rsBalance!是否普通结算)) = 0, Val(Nvl(rsBalance!卡类别ID)), 0)
                    .消费卡 = False
                    .消费卡ID = 0
                    .卡号 = IIf(Nvl(rsBalance!卡号) = "", objItem.卡号, Nvl(rsBalance!卡号))
                    If .卡号 <> "" Then str卡号 = .卡号
                    .交易流水号 = Nvl(rsBalance!交易流水号)
                    .交易说明 = Nvl(rsBalance!交易说明)
                    .结算方式 = Nvl(rsBalance!结算方式)
                    .结算金额 = Val(Nvl(rsBalance!交易金额))
                    .密码 = objItem.密码
                    .是否密文 = objItem.是否密文
                    .限制类别 = objItem.限制类别
                    .帐户余额 = objItem.帐户余额
                    .结算号码 = Nvl(rsBalance!结算号码)
                    .结算摘要 = Nvl(rsBalance!结算摘要)
                    .是否允许编辑 = False
                    .是否允许删除 = False
                    .是否保存 = True
                    .是否结算 = True
                    .关联交易ID = lng关联交易ID
                    .是否结算 = True
                    .结算类型 = 3   ''0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                    .校对标志 = 2
                End With
                Set objCurItem_Out = objCurItem
            End If
        End With
    Else
        With objCurItem
            .交易流水号 = strSwapGlideNO
            .交易说明 = strSwapMemo
            .是否结算 = True
            .是否结算 = True
            .是否允许编辑 = False
            .是否允许删除 = False
            .是否保存 = True
            .是否结算 = True
            .校对标志 = 2
        End With
        Set objCurItem_Out = objCurItem
    End If
    
    Call zlAddUpdateSwapSQL(False, objCurItem_Out.ID, objCard.接口序号, objCard.消费卡, objCurItem.卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 0, lng关联交易ID)
    blnTrans = True
    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption & "-执行三方接口", , Not mbln结算异步控制
    blnSaveed_Out = True
    
    blnTrans = False
    Set cllBillPro = New Collection
    Call zlGetThreeSwapExpendSQLToCollect(objCard, objCurItem.ID, str卡号, rsExpend, cllThreeSwap, True)
    Err = 0: On Error GoTo ErrOther:
    blnTrans = True
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption & "-执行三方接口"
    blnTrans = False
    zlThird_Payment = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        If blnExcuteThird Then GoTo ReDO:
        Resume
    End If
    Call SaveErrLog
    Exit Function
ErrOther:
    If ErrCenter() = 1 Then Resume
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
End Function

Public Function zlThirdReturnCashCheck(ByVal objCard As Card, ByRef objBalanceItems As clsBalanceItems, ByRef blnDefaultCash_Out As Boolean, ByRef strDefaultBalance_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用"zlReturnCashCheck"接口
    '入参:
    '出参:blnDefaultCash_Out-是否缺省退现
    '     strDefaultBalance_Out-缺省退现方式
    '返回:允许退现返回tru有,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-09 14:21:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String, str卡号 As String
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim strIDs As String, strTKList As String
    Dim strBalanceIDs As String, dblMoney As Double
    
    Dim i As Long
    
    
    Err = 0: On Error GoTo errHandle
    If objBalanceItems Is Nothing Then Exit Function
    If objBalanceItems.Count = 0 Then Exit Function
    
    strBalanceIDs = objBalanceItems.收费类型
    strIDs = ""
    For i = 1 To objBalanceItems.Count
        Set objItemTemp = objBalanceItems(i)
        If objItemTemp.卡号 <> "" And objItem Is Nothing Then Set objItem = objItemTemp
        strIDs = strIDs & "," & objItemTemp.预交ID
        strTKList = ""
        strTKList = strTKList & Space(8) & "<TK>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <TKFS>" & objItemTemp.结算方式 & "</TKFS>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <TKJE>" & objItemTemp.结算金额 & "</TKJE>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <JYLSH>" & objItemTemp.交易流水号 & "</JYLSH>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <JYSM>" & objItemTemp.交易说明 & "</JYSM>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <KH>" & objItemTemp.卡号 & "</KH>" & vbCrLf
        strTKList = strTKList & Space(8) & "</TK>" & vbCrLf
        strXMLExpend = strXMLExpend & strTKList
        dblMoney = dblMoney + objItemTemp.结算金额
    Next
    
    If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    strTKList = ""
    strTKList = strTKList & "<INPUT>" & vbCrLf
    strTKList = strTKList & "    <TKLIST>" & vbCrLf
    strTKList = strTKList & strXMLExpend
    strTKList = strTKList & "    </TKLIST>" & vbCrLf
    strTKList = strTKList & "</INPUT>"

    strXMLExpend = strTKList
    If objItem Is Nothing Then Set objItem = objBalanceItems(1)
    strBalanceIDs = strBalanceIDs & "|" & strIDs
    
    'Public Function zlReturnCashCheck(frmMain As Object, ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
    '    ByVal strBalanceIDs As String, _
    '    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    '    ByVal strSwapMemo As String, ByRef strXMLExpend As String, _
    '    Optional blnDelDefaultCash_Out As Boolean, Optional strDefaultDelBalance_Out As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:退现交易检查
    '    '入参:frmMain-调用的主窗体
    '    '       lngModule-调用的模块号
    '    '       lngCardTypeID-卡类别ID
    '    '       strCardNo-卡号
    '    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '    '       dblMoney-退款金额
    '    '       strSwapNo-交易流水号(退款时检查)多种结算方式时，本参数为第一个结算方式的交易流水号
    '    '       strSwapMemo-交易说明(退款时传入) 多种结算方式时，本参数为第一个结算方式的交易说明
    '    '       strXMLExpend    XML IN  10.35.90后才支持
    '    '        <INPUT>
    '    '            <TKLIST>    //本次退款列表
    '    '                <TK>
    '    '                    <TKFS>退款方式</TKFS>
    '    '                    <TKJE>退款金额</TKJE>
    '    '                    <JYLSH>原交易流水号</JYLSH>
    '    '                    <JYSM>原交易说明</JYSM>
    '    '                </TK>
    '    '                ....
    '    '            </TKLIST>
    '    '        </INPUT>
    '    '出参:
    '    '       blnDelDefaultCash_Out-是否缺省退现：接口返回true时有效，true时：表示缺省退成现金（缺省方式为:str缺省退现方式_Out返回值),否则缺省退回原卡，但允许操作员选择退为现金
    '    '       strDefaultDelBalance_Out-缺省退现方式,比如：支票，现金等
    '    '       strXMLExpend:10.35.90后才支持
    '    '        <OUTPUT>
    '    '            <SFQSTX>是否缺省退现<SFQSTX>//NUMBER 1 是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但以许操作员操作退现
    '    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
    '    '                    1.不允许返回三卡方的结算方式
    '    '                    2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
    '    '        </OUTPUT>
    '    '返回:退现合法,返回true,否则返回Flase

    zlThirdReturnCashCheck = gobjSquare.objSquareCard.zlReturnCashCheck(mfrmMain, mlngModule, objCard.接口序号, objItem.卡号, _
          strBalanceIDs, dblMoney, objItem.交易流水号, objItem.交易说明, strXMLExpend, blnDefaultCash_Out, strDefaultBalance_Out)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThird_TransferAccounts(ByVal objPati As clsPatiInfo, ByRef objTranItem As clsBalanceItem, ByRef cllBillPro As Collection, _
    Optional strErrMsg_out As String, Optional blnSaveed_Out As Boolean, Optional blnCurInput As Boolean, _
    Optional ByVal int险类 As Integer, Optional ByVal bln电子票据 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通转帐支付
    '入参:
    '    cllBillPro-执行的相关过程
    '    objBalanceInfor-当前结帐数据
    '    objTranItem-转账项
    '    blnCurInput-是否当前输入项
    '出参:objBalanceInfor-结帐返回后更新的结帐数据列表
    '     strErrMsg_Out-返回的错误信息
    '     blnSaveed_Out-是否发生数据保存
    '返回:三方接口交易调用成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 17:02:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean, dblMoney As Double
    Dim str结帐IDs As String, i As Long, lng关联交易ID As Long, lng预交ID As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset, dbl帐户余额 As Double
    Dim strCardNo As String, str结算方式   As String
    Dim strXMLExpend As String, strErrMsg As String, rsExpend As ADODB.Recordset
    Dim objItem As clsBalanceItem, objItems As clsBalanceItems, objItemsTemp As clsBalanceItems
    Dim lngCount As Long, intStatus  As Integer
    Dim objCard As Card, blnCheck As Boolean, blnCallInterface As Boolean   '开始调用接口
    Dim cllBillUpdate As Collection     '接口调用前更新记录状态
    
    Err = 0: On Error GoTo Errhand:


    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    blnSaveed_Out = False
    
    If objTranItem Is Nothing Then Exit Function
    If UCase(TypeName(objTranItem.objTag)) <> UCase("clsBalanceItems") Then Exit Function
    If objTranItem.objTag Is Nothing Then Exit Function
    
    Set objItems = objTranItem.objTag
    Set objCard = objTranItem.objCard
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then zlThird_TransferAccounts = True: Exit Function
    If Not objCard.是否转帐及代扣 Then Exit Function
    
    dblMoney = -1 * objTranItem.结算金额
    '先检查是否合法
    blnCheck = True
    If objTranItem.是否保存 Then
        If Not zlThird_IsSwapIsSucces(objItems, intStatus, strErrMsg) Then
           If intStatus = 1 Then
                If strErrMsg <> "" Then
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行转帐处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                Else
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行转帐处理，请稍后再试!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                End If
                Exit Function
           End If
        Else
            blnCheck = False
        End If
    End If
    
    If blnCheck Then
        If zlThird_TransferAccounts_IsValied(objPati, objTranItem, objItem, dbl帐户余额, strErrMsg_out, blnCurInput) = False Then Exit Function
    End If
    
    
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    
    If Not objItem Is Nothing Then
        If objItem.结算方式 <> "" Then objTranItem.结算方式 = objItem.结算方式
        objTranItem.卡号 = objItem.卡号
        objTranItem.交易流水号 = objItem.交易流水号
        objTranItem.交易说明 = objItem.交易说明
    End If
    strCardNo = objTranItem.卡号
    strSwapNO = objTranItem.交易流水号
    strSwapMemo = objTranItem.交易说明
    lng关联交易ID = objTranItem.关联交易ID
    objTranItem.关联交易ID = 0
    If zlGetSaveSQLfromItem(objPati, objTranItem, IIf(mbln结算异步控制, 1, 0), cllPro, True, , int险类, bln电子票据) = False Then
        objTranItem.关联交易ID = lng关联交易ID
        Exit Function
    End If
    objTranItem.关联交易ID = objTranItem.ID
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, Not mbln结算异步控制
    If mbln结算异步控制 Then
        blnTrans = False: blnSaveed_Out = True: objTranItem.是否保存 = True
        Set cllBillUpdate = New Collection
        If zlGetSaveSQLfromItem(objPati, objTranItem, 0, cllBillUpdate, False, True, int险类, bln电子票据) = False Then Exit Function
        blnTrans = True
        zlExecuteProcedureArrAy cllBillUpdate, mfrmMain.Caption, True
    End If
    
    blnCallInterface = True
   'zlTransferAccountsMoney(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
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
    '           <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务；4-门诊退费业务或预交款退款
    '       </IN>
    '出参:
    '   strSwapGlideNO-交易流水号
    '   strSwapMemo -交易说明
    '   strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '   strXMLExpend-XML串:
    '        <OUT>
    '           <JYZT>错误信息</JYZT>   //10.35.90增加此接点:接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '           <ERRMSG>错误信息</ERRMSG >
    '        </OUT>
    '   strErrMsg_Out-XML返回的错误信息
    '   rsExpend_Out-返回扩展信息集
    '   intStatus_Out-交易状态:0-失败;1-交易进行中
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------

    strXMLExpend = "<IN><CZLX>" & 4 & "</CZLX></IN>"
    strSwapExtendInfor = "1|" & objTranItem.ID
    
    If gobjSquare.objSquareCard.zlTransferAccountsMoney(mfrmMain, mlngModule, objCard.接口序号, _
        strCardNo, objTranItem.ID, Abs(dblMoney), strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend, rsExpend, strErrMsg_out, intStatus) = False Then
        
        If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
        If strErrMsg_out <> "" Then MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
        If blnSaveed_Out Then Set cllBillPro = New Collection
                
        If intStatus = 1 Then   '交易正在进行中
            If blnSaveed_Out Then
                objTranItem.是否保存 = True
                objTranItem.是否结算 = False
                objTranItem.是否允许编辑 = False
                objTranItem.是否允许删除 = True
            End If
            
            Exit Function
        End If
        blnCallInterface = False
        If mbln结算异步控制 Then
            Call zlDelBalanceItemFromItemObject(objTranItem)
        End If
        Exit Function
    End If
    
    '提交状态更新
    lng关联交易ID = objTranItem.关联交易ID
    Call zlAddUpdateSwapSQL(True, objTranItem.ID, objCard.接口序号, False, strCardNo, strSwapNO, strSwapMemo, cllUpdate, 2, lng关联交易ID)

    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, True, Not mbln结算异步控制 Or blnTrans
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    
    zlThird_TransferAccounts = True
    
    If Not mbln结算异步控制 Then blnSaveed_Out = True
    Set cllBillPro = New Collection
    With objTranItem
        .交易流水号 = strSwapNO
        .交易说明 = strSwapMemo
        .卡号 = strCardNo
        .是否保存 = True
        .是否结算 = True
        .是否允许编辑 = False
        .是否允许删除 = False
    End With
    
    Call zlGetThreeSwapExpendSQLToCollect(objCard, objTranItem.ID, strCardNo, rsExpend, cllThreeSwap, True)
    On Error GoTo ErrOther:
    blnTrans = True
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption
    blnTrans = False
    
    Exit Function
Errhand:
   If blnTrans And blnCallInterface = False Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
    Exit Function
ErrOther:
    If blnTrans Then gcnOracle.CommitTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Public Function zlThird_ReturnMoney_IsValied(ByVal objPati As clsPatiInfo, ByVal objCurItem As clsBalanceItem, ByVal byt操作类型 As Byte, _
    ByRef objItem_Out As clsBalanceItem, Optional blnCurInput As Boolean, Optional blnCallSwapIsSucces As Boolean, Optional ByVal bln是否异常退费 As Boolean) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方卡退款合法性检查
    '入参:objPati-病人信息对象
    '     objCard-卡对象
    '     objDelBalanceItems-退费结算信息
    '     byt操作类型-0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
    '     objBalanceInfor-当前结帐信息
    '     blnCurInput-是否当前输入的操作
    '     blnCallSwapIsSucces-是否调用了成功交易判断接口
    '    byvalbln是否异常退费-是否异常退费
    '出参:objItems_Out-当前返顺的有效结算信息
    '返回:允许退款返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-28 18:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend  As String, str交易说明 As String, str交易流水号 As String, strBalanceIDs As String, str卡号 As String, str密码 As String
    Dim bln退现 As Boolean, intStatus As Integer, strErrMsg As String
    Dim dblMoney  As Double, dbl原始金额 As Double, dbl未退金额 As Double, rsSwap As ADODB.Recordset
    Dim strXmlIn As String
    Dim objItem As clsBalanceItem
    Dim i As Long, blnCheck As Boolean
    Dim objCard As Card
    
    
    On Error GoTo errHandle
    
    If objCurItem Is Nothing Then Exit Function
    Set objCard = objCurItem.objCard
    
    
    If objCard Is Nothing Then
        Screen.MousePointer = 0
        If InStr(";" & mstrOnCardPrivs & ";", ";三方接口消费;") = 0 Then
            MsgBox "你没有三方接口消费权限，无法调用接口部件！", vbInformation, gstrSysName
            Exit Function
        End If
        MsgBox "未找到退款接口,请检查接口部件！", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
        Exit Function
    End If
        
        
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then zlThird_ReturnMoney_IsValied = True: Exit Function
    
    
    If objCurItem Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objCurItem.结算金额 < 0 Then '
        Screen.MousePointer = 0
        MsgBox "退款金额不能为负数，请检查!", vbOKOnly, gstrSysName
        Exit Function
    End If
        
   '先检查数据的合法性
    blnCheck = True
    If objCurItem.是否保存 And Not blnCallSwapIsSucces Then
        If Not zlThird_IsSwapIsSucces(objCurItem, intStatus, strErrMsg) Then
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
             Set objItem_Out = objCurItem
            zlThird_ReturnMoney_IsValied = True: blnCheck = False: Exit Function
        End If
    End If
    
    dbl未退金额 = objCurItem.未退金额
    dblMoney = objCurItem.结算金额
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
    
    If roundEx(Abs(dblMoney), 2) > roundEx(Abs(dbl未退金额), 2) And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '    Public Function zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
    '    ByVal lngCardTypeId As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
    '    ByVal strBalanceIDs As String, _
    '    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    '    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:帐户回退交易前的检查
    '    '入参:frmMain-调用的主窗体
    '    '       lngModule-调用的模块号
    '    '       lngCardTypeID-卡类别ID
    '    '       strCardNo-卡号
    '    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算
    '    '       dblMoney-退款金额
    '    '       strSwapNo-交易流水号(退款时检查),保险补充结算时传入空
    '    '       strSwapMemo-交易说明(退款时传入),保险补充结算时传入空
    '    '       strXMLExpend    XML IN
    '    '        <TFDATA>   //退费数据
    '    '            <YCTF>异常退费标志<YCTF> //1-异常重退;0-退费此节点可能没传入
    '    '            <TFLIST>  //退费列表
    '    '                <NO></NO>  // 退费单据
    '    '                <TFITEM>     //退费项
    '    '                    <SerialNum>序号</SerialNum>
    '    '                    ….
    '    '                </ TFITEM >
    '    '            </TFLIST>
    '    '
    '    '            <TKLIST>   //退款列表（35.90以前无此内容）
    '    '                <TK>
    '    '                    <TKFS>退款方式</TKFS>// Varchar2    20
    '    '                    <TKJE>退款金额</TKJE>//NUMBER
    '    '                    <JSLSH>原交易流水号</JSLSH>//   Varchar2    50
    '    '                    <JYSM><原交易说明</JYSM>//  Varhcar2    500
    '    '                    <DJH>单据号</DJH> //    Varchar2    8
    '    '                </TK>
    '    '                ....
    '    '            </TKLIST>
    '    '        </TFDATA>
    '    '返回:退款合法,返回true,否则返回Flase
    '    '编制:刘兴洪
    '    '日期:2011-05-26 17:24:55
    '    '说明:
    '    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
    
    strXMLExpend = ""
    
    dblMoney = roundEx(objCurItem.结算金额, 2)
    strXMLExpend = "" & vbCrLf & _
    "              <TK>" & vbCrLf & _
    "                   <TKFS>" & objCurItem.结算方式 & "</TKFS>" & vbCrLf & _
    "                   <TKJE>" & roundEx(objCurItem.结算金额, 6) & "</TKJE>" & vbCrLf & _
    "                   <JYLSH>" & objCurItem.交易流水号 & "</JYLSH>" & vbCrLf & _
    "                   <JYSM>" & objCurItem.交易说明 & "</JYSM>" & vbCrLf & _
    "                   <DJH>" & "" & "</DJH>" & vbCrLf & _
    "              </TK>" & vbCrLf
    
    If strBalanceIDs = "" Then
        strBalanceIDs = "1|" & objCurItem.预交ID
    Else
        strBalanceIDs = strBalanceIDs & "||1|" & objCurItem.预交ID
    End If

    str交易说明 = objCurItem.交易说明
    str交易流水号 = objCurItem.交易流水号
       
    strXMLExpend = "" & _
    "<TFDATA>" & vbCrLf & _
    "       <YCTF>" & IIf(bln是否异常退费, "1", "0") & "</YCTF>" & vbCrLf & _
    "       <TKLIST>" & vbCrLf & _
                strXMLExpend & vbCrLf & _
    "       </TKLIST>" & vbCrLf & _
    "<TFDATA>" & vbCrLf
    
    If gobjSquare.objSquareCard.zlReturnCheck(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, objCurItem.卡号, _
        strBalanceIDs, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then Exit Function
    
    
    If Not objCard.是否退款验卡 Then
        Set objItem_Out = objCurItem: zlThird_ReturnMoney_IsValied = True: Exit Function
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
    If gobjSquare.objSquareCard.zlBrushCard(mfrmMain, mlngModule, Nothing, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, True, True, bln退现, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    
    Set objItem_Out = objCurItem: zlThird_ReturnMoney_IsValied = True: Exit Function
    zlThird_ReturnMoney_IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThird_ReturnMultiMoney(ByVal objPati As clsPatiInfo, ByRef cllBillPro As Collection, objCurItem As clsBalanceItem, _
    ByRef objDelItem_Out As clsBalanceItem, Optional blnSaveed_Out As Boolean, Optional blnCurInput As Boolean, _
    Optional ByVal int险类 As Integer, Optional ByVal bln电子票据 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一次按多笔交易返回三方接口
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     objBalanceInfor-原结算方式信息集
    '     objCurItem-当前退费的结算项(objCurItem.objTag为当前多笔退款所对应的结算信息)
    '     cllBillPro-需执行的相关SQL
    '     blnCurInput-是否当前输入项
    '出参:
    '     objDelItem_Out-返回已经退费成功的项目
    '     blnSaveed_Out-是否发生数据保存
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-16 19:52:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem, objCurDelItems As clsBalanceItems, objItemsTemp As clsBalanceItems
    Dim strXML As String, strInXML As String, strOutXml As String, strExpend As String, strErrMsg As String
    Dim cllPro As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim strCardNo As String, str结算方式 As String, str关联交易ID As String, strSQL As String
    Dim dblMoney As Double, lngCount As Long, intStatus As Integer, i As Long
    Dim rsBalance As ADODB.Recordset, rsExpend As ADODB.Recordset
    Dim blnTrans As Boolean, blnCheck As Boolean, blnCallInterface As Boolean
    Dim objCard As Card
    On Error GoTo errHandle
    
    blnSaveed_Out = False
    If objCurItem Is Nothing Then Exit Function
    Set objCard = objCurItem.objCard
    If objCard Is Nothing Then Exit Function
    If objCard.接口序号 < 0 Or objCard.消费卡 Then zlThird_ReturnMultiMoney = True: Exit Function
    
    Set objCurDelItems = objCurItem.objTag
    If objCurDelItems Is Nothing Then
        MsgBox "注意:" & vbCrLf & _
        "   " & objCard.名称 & " 进行退款时，未发现需要退款的交易，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objCurDelItems.Count = 0 Then
        MsgBox "注意:" & vbCrLf & _
        "   " & objCard.名称 & " 进行退款时，未发现需要退款的交易，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '先检查数据的合法性
    blnCheck = True
    If objCurItem.是否保存 Then
        If Not zlThird_IsSwapIsSucces(objCurDelItems, intStatus, strErrMsg) Then
           If intStatus = 1 Then
                If strErrMsg <> "" Then
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行退款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                Else
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行退款处理中，请稍后再试" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                End If
                Exit Function
           End If
        Else
            blnCheck = False
        End If
    End If
    
    If blnCheck Then
        For Each objItem In objCurDelItems
            If InStr(str关联交易ID & ",", "," & objItem.关联交易ID & ",") = 0 Then
                Set objItemsTemp = New clsBalanceItems
                If zlThird_ReturnMoney_IsValied(objPati, objItem, 2, objItemsTemp, blnCurInput) = False Then Exit Function
                str关联交易ID = str关联交易ID & "," & objItem.关联交易ID
            End If
        Next
    End If
    
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    If zlGetSaveSQLfromItem(objPati, objCurItem, IIf(mbln结算异步控制, 1, 0), cllPro, True, , int险类, bln电子票据) = False Then Exit Function
    
    strInXML = "": strXML = ""
    For Each objItem In objCurDelItems
        
        strXML = strXML & "   <JS>" & vbCrLf
        strXML = strXML & "        <TKFS>" & objItem.结算方式 & "</TKFS>" & vbCrLf     '  //10.35.90后新增,退款方式,即结算方式.名称
        strXML = strXML & "        <KH>" & objItem.卡号 & "</KH>" & vbCrLf
        strXML = strXML & "        <JYLSH>" & objItem.交易流水号 & "</JYLSH>" & vbCrLf
        strXML = strXML & "        <JYSM>" & objItem.交易说明 & "</JYSM>" & vbCrLf
        strXML = strXML & "        <ZFJE>" & -1 * objItem.结算金额 & "</ZFJE>" & vbCrLf
        strXML = strXML & "        <JYLX>" & 1 & "</JYLX>" & vbCrLf   '1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        strXML = strXML & "        <GLJYID>" & objItem.关联交易ID & "</GLJYID>" & vbCrLf
        strXML = strXML & "        <DJH></DJH>" & vbCrLf   '  //10.35.90新增:单据号,多单据收费时有用 ，存储在"医保结算明细"表中,主要是分单据保存
        strXML = strXML & "        <ID>" & objItem.预交ID & "</ID>" & vbCrLf
        strXML = strXML & "   </JS>" & vbCrLf
    Next
    
    strInXML = strInXML & "<JSLIST>" & vbCrLf
    strInXML = strInXML & strXML & vbCrLf
    strInXML = strInXML & "</JSLIST>" & vbCrLf

    
    
    blnTrans = True
    Call zlExecuteProcedureArrAy(cllPro, mfrmMain.Caption, Not mbln结算异步控制)
    If mbln结算异步控制 Then
        blnTrans = False: blnSaveed_Out = True
        Set cllUpdate = New Collection
        Call zlGetSaveSQLfromItem(objPati, objCurItem, 0, cllUpdate, False, , int险类, bln电子票据)
        
        blnTrans = True
        zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, True
    End If
    
    
    blnCallInterface = True
    'Public Function zlReturnMultiMoney(frmMain As Object, ByVal lngModule As Long, _
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
    '                    <TKFS>退款方式</TKFS>  //10.35.90后新增,退款方式,即结算方式.名称
    '                    <KH>卡号</KH>
    '                    <JYLSH>交易流水号</JYLX>
    '                    <JYSM>交易说明</JYSM>
    '                    <ZFJE>作废金额</ZFJE>
    '                    <JSLX>类型</JSLX>
    '                    <DJH>单据号</DJH>  //10.35.90新增:单据号,多单据收费时有用 ，存储在"医保结算明细"表中,主要是分单据保存
    '                    <GLJYID>关联交易ID</GLJYID>  //10.35.90新增:原始交易关联ID(病人预交记录.关联交易ID)
    '                    <ID></ID>
    '                </JS>
    '            </JSLIST>

    '       lng冲销ID-作废时的冲销ID(作废时或退费时有效，否则为0）;类型=6时，传入结算序号
    '       strExpend-无（暂留，待以后扩展)
    '    blnResolveXMLToRecord-是否解析XML串给记录集(rsBalance_Out,rsExpend_Out）
    '出参:
    '     strOutXML-返回XML串
    '     <OUTPUT>                      //10.35.90中增加此接点(接口作了兼容处理)
    '         <JYZT>交易状态</JYZT>    10.35.90中增加此接点:接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '         <ERRMSG>错误信息</ERRMSG> 10.35.90中增加此接点(接口作了兼容处理):错误信息:为空时，不提示，非空时，提示
    '         <JSLIST>
    '            <JS>
    '                <TKFS>退款方式</TKFS>   //10.35.90新增:退款方式即结算方式.名称
    '                <TKJE>退款方式</TKJE>   //10．35．90退款金额
    '                <KH>卡号</KH>
    '                <TKLSH>退款交易流水号</TKLSH>
    '                <TKSM>退款交易说明</TKSM>
    '                <ID></ID>   //类型=1时,预交ID;类型=2,6时，为原结帐ID
    '                <DJH>单据号</DJH>   //10．35．90后增加:单据号,多单据收费时有用 ，存储在“医保结算明细”表中,主要是分单据保存
    '                <GLJYID>关联交易ID</GLJYID>  //10.35.90新增:原始交易关联ID(病人预交记录.关联交易ID)
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
    If gobjSquare.objSquareCard.zlReturnMultiMoney(mfrmMain, mlngModule, objCurItem.卡类别ID, objCurItem.消费卡, strInXML, objCurItem.ID, strOutXml, strExpend, True, rsBalance, rsExpend, intStatus, strErrMsg) = False Then
         If blnTrans Then gcnOracle.RollbackTrans
         If strErrMsg <> "" Then
                MsgBox "注意:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
         End If
         
         If intStatus = 1 Then '交易正在处理中
            If blnSaveed_Out Then
                objCurItem.是否保存 = True
                objCurItem.是否允许编辑 = False
                objCurItem.是否允许删除 = False
                objCurItem.是否结算 = False
            End If
            Set objDelItem_Out = objCurItem
            Exit Function
         End If
         If mbln结算异步控制 Then
            Call zlDelBalanceItemFromItemObject(objCurItem)
         End If
         Exit Function
    End If
     
    Set objDelItem_Out = objCurItem
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    Dim lngID As Long
    str关联交易ID = ""
    '更新数据
    If Not rsBalance Is Nothing Then
        Set objCurDelItems = New clsBalanceItems
        dblMoney = 0
        With rsBalance
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
            
                Set objItemTemp = Nothing
                For Each objCurItem In objCurDelItems
                    lngID = objCurItem.预交ID
                    If (lngID = Val(Nvl(!ID)) And Val(Nvl(!ID)) <> 0) Or (Val(Nvl(!ID)) = 0 And Val(Nvl(!卡类别ID)) = objCurItem.卡类别ID And Nvl(!卡号) = objCurItem.卡号) Then
                         Exit For
                    End If
                Next
                Set objItemTemp = New clsBalanceItem
                
                With objItemTemp
                    .卡类别ID = IIf(Val(Nvl(rsBalance!是否普通结算)) = 0, Val(Nvl(rsBalance!卡类别ID)), 0)
                    .消费卡 = False
                    .消费卡ID = 0
                    .卡号 = IIf(Nvl(rsBalance!卡号) = "", objItem.卡号, Nvl(rsBalance!卡号))
                    If strCardNo = "" Then strCardNo = .卡号
                    .交易流水号 = Nvl(rsBalance!原交易流水号)
                    If .交易流水号 = "" Then .交易流水号 = objCurItem.交易流水号
                    .交易说明 = Nvl(rsBalance!原交易说明)
                    If .交易说明 = "" Then .交易说明 = objCurItem.交易说明
                    .退款交易流水号 = Nvl(rsBalance!交易流水号)
                    .退款交易说明 = Nvl(rsBalance!交易说明)
                    
                    .结算方式 = Nvl(rsBalance!结算方式)
                    If .结算方式 = "" Then .结算方式 = objCurItem.结算方式
                    .结算号码 = Nvl(rsBalance!结算号码)
                    .结算摘要 = Nvl(rsBalance!结算摘要)
                    .结算金额 = IIf(Val(Nvl(rsBalance!交易金额)) = 0, objCurItem.结算金额, -1 * Val(Nvl(rsBalance!交易金额)))
                    .密码 = objItem.密码
                    .是否密文 = objCurItem.是否密文
                    .限制类别 = objCurItem.限制类别
                    .帐户余额 = objCurItem.帐户余额
                    .关联交易ID = Val(Nvl(rsBalance!关联交易ID))
                    .收款时间 = objCurItem.收款时间
                    .预交ID = Val(Nvl(rsBalance!ID))
                     If .预交ID = 0 Then .预交ID = objCurItem.预交ID
                    
                    .是否退款分交易 = False
                    .是否预交 = objCurItem.是否预交
                    .是否允许编辑 = False
                    .是否允许删除 = False
                    .是否允许退现 = False
                    .是否转帐 = objCurItem.是否转帐
                    .是否保存 = True
                    .是否结算 = True
                    Set .objCard = objCurItem.objCard
                    dblMoney = dblMoney + .结算金额
                End With
                str关联交易ID = str关联交易ID & "," & objItemTemp.关联交易ID
                .MoveNext
            Loop
            objCurDelItems.结算金额 = dblMoney
        End With
    Else
         
        objCurItem.是否退款分交易 = False
        objCurItem.是否预交 = objCurItem.是否预交
        objCurItem.是否允许编辑 = False
        objCurItem.是否允许删除 = False
        objCurItem.是否允许退现 = False
        objCurItem.是否转帐 = False
        objCurItem.是否保存 = True
        objCurItem.是否结算 = True
    End If
    
    blnTrans = False
    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, True, True
    blnSaveed_Out = True
    Set cllBillPro = New Collection
    
    blnTrans = False
    zlThird_ReturnMultiMoney = True
    
    Err = 0: On Error GoTo ErrOther:
    Call zlGetThreeSwapExpendSQLToCollect(objCard, objCurItem.ID, strCardNo, rsExpend, cllThreeSwap, True)
    
    blnTrans = True
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption & "-执行三方接口"
    blnTrans = False
    Exit Function
 
errHandle:
    If blnTrans And blnCallInterface = False Then blnTrans = False: gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then blnTrans = False: gcnOracle.RollbackTrans
    Call SaveErrLog
    Exit Function
ErrOther:
    If ErrCenter() = 1 Then Resume
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
End Function

Public Function zlGetSaveSQLfromItem(ByVal objPati As clsPatiInfo, ByRef objCurItem As clsBalanceItem, _
    ByVal int操作状态 As Integer, ByRef cllPro As Collection, Optional blnBulidID As Boolean, _
    Optional blnOlnyOverBalance As Boolean, Optional ByVal int险类 As Integer, Optional ByVal bln电子票据 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据退款信息项获取保存的SQL
    '入参:objCurItem-当前项
    '     int校对标志-1-需要校对;2或0表示成功
    '     blnBulidID-重新生成ID和单据号
    '     blnOlnyOverBalance-仅完成结算
    '     int操作状态:0-表示完成结算;1-表示未完成结算; (int操作状态=1时,生成的预交记录的校对标志为1)
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-30 15:03:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng主页ID As Long, strSQL As String, strHead As String
    Dim objItems As clsBalanceItems, str结算信息 As String '原预交ID|金额||....
    Dim objItem As clsBalanceItem
    Dim i As Long, strTemp As String
    
    On Error GoTo errHandle
    If blnBulidID Then
        With objCurItem
            .ID = zlDatabase.GetNextId("病人预交记录")
            .收款时间 = zlDatabase.Currentdate
            .单据号 = zlDatabase.GetNextNo(11)
        End With
    End If
    If cllPro Is Nothing Then Set cllPro = New Collection
        
    Set objItems = objCurItem.objTag
    str结算信息 = ""
    If objItems Is Nothing Then
        str结算信息 = objCurItem.预交ID & "|" & objCurItem.结算金额
    End If
    'Zl_病人预交记录_余额退款_S
    strHead = "Zl_病人预交记录_余额退款_S("
    '  Id_In         病人预交记录.Id%Type,
    strHead = strHead & "" & objCurItem.ID & ","
    '  单据号_In     病人预交记录.No%Type,
    strHead = strHead & "'" & objCurItem.单据号 & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strHead = strHead & "'" & objCurItem.发票号 & "',"
    '  病人id_In     病人预交记录.病人id%Type,
    strHead = strHead & "" & ZVal(objPati.病人ID) & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strHead = strHead & "" & ZVal(objPati.主页ID) & ","
    '  姓名_In         病人预交记录.姓名%Type,
    strHead = strHead & "'" & objPati.姓名 & "',"
    '  性别_In         病人预交记录.性别%Type,
    strHead = strHead & "'" & objPati.性别 & "',"
    '  年龄_In         病人预交记录.年龄%Type,
    strHead = strHead & "'" & objPati.年龄 & "',"
    '  门诊号_In       病人预交记录.门诊号%Type,
    strHead = strHead & ZVal(objPati.门诊号) & ","
    '  住院号_In       病人预交记录.住院号%Type,
    strHead = strHead & ZVal(objPati.住院号) & ","
    '  付款方式名称_In 病人预交记录.付款方式名称%Type,
    strHead = strHead & "'" & objPati.医疗付款方式 & "',"
    '  科室id_In     病人预交记录.科室id%Type,
    strHead = strHead & "" & ZVal(objCurItem.科室ID) & ","
    '  金额_In       病人预交记录.金额%Type,
    strHead = strHead & "" & roundEx(-1 * objCurItem.结算金额, 6) & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strHead = strHead & "'" & objCurItem.结算方式 & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strHead = strHead & "'" & objCurItem.结算号码 & "',"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    strHead = strHead & "'" & objCurItem.缴款单位 & "',"
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    strHead = strHead & "'" & objCurItem.开户行 & "',"
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    strHead = strHead & "'" & objCurItem.帐号 & "',"
    '  摘要_In       病人预交记录.摘要%Type,
    strHead = strHead & "'" & objCurItem.结算摘要 & "',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strHead = strHead & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strHead = strHead & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strHead = strHead & "" & ZVal(objCurItem.领用ID) & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strHead = strHead & "" & IIf(objCurItem.门诊预交, 1, 2) & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strHead = strHead & "" & IIf(objCurItem.卡类别ID = 0 Or objCurItem.消费卡, "NULL", objCurItem.卡类别ID) & ","
    '  结算卡序号_In 病人预交记录.结算卡序号%Type := Null,
    strHead = strHead & "" & IIf(objCurItem.卡类别ID = 0 Or objCurItem.消费卡 = False, "NULL", objCurItem.卡类别ID) & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strHead = strHead & "'" & objCurItem.卡号 & "',"
    '  关联交易id_In 病人预交记录.关联交易id%Type := Null,
    strHead = strHead & "" & ZVal(objCurItem.关联交易ID) & ","
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strHead = strHead & "'" & objCurItem.交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strHead = strHead & "'" & objCurItem.交易说明 & "',"
    '  合作单位_In   病人预交记录.合作单位%Type := Null,
    strHead = strHead & "" & "NULL" & ","
    '  收款时间_In   病人预交记录.收款时间%Type := Null,
    strHead = strHead & "to_date('" & Format(objCurItem.收款时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    i = 0
    If Not objItems Is Nothing Then
        str结算信息 = ""
        For Each objItem In objItems
            strTemp = "||" & objItem.预交ID & "|" & objItem.结算金额
            If zlCommFun.ActualLen(str结算信息 & strTemp) > 4000 Then
                str结算信息 = Mid(str结算信息, 3)
                '  结算信息_In   Varchar2 := Null,原预交ID|金额||....
                strSQL = strHead & "'" & str结算信息 & "',"
                '  仅更新数据_In Number := 0,
                strSQL = strSQL & IIf(i = 0, 0, 1) & ","
                '  操作状态_In Number:=0
                strSQL = strSQL & 1 & ")"
                i = i + 1
                zlAddArray cllPro, strSQL
                str结算信息 = ""
            End If
            str结算信息 = str结算信息 & strTemp
        Next
        If str结算信息 <> "" Then str结算信息 = Mid(str结算信息, 3)
    End If
    '  结算信息_In   Varchar2 := Null,原预交ID|金额||....
    strSQL = strHead & "'" & IIf(blnOlnyOverBalance, "", str结算信息) & "',"
    '  仅更新数据_In Number := 0,
    strSQL = strSQL & IIf(i > 0 Or blnOlnyOverBalance, 1, 0) & ","
    '  操作状态_In Number:=0
    strSQL = strSQL & int操作状态 & ","
    '    结帐id_In       病人预交记录.结帐id%Type := Null,
    strSQL = strSQL & "NULL,"
    '    结算序号_In     病人预交记录.结算序号%Type := Null,
    strSQL = strSQL & "NULL,"
    '    预交电子票据_In 病人预交记录.预交电子票据%Type := Null,
    strSQL = strSQL & IIf(bln电子票据, 1, 0) & ","
    '    险类_In         保险结算记录.险类%Type := Null
    strSQL = strSQL & int险类 & ")"

    zlAddArray cllPro, strSQL
    zlGetSaveSQLfromItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlThird_ReturnMoney(ByVal objPati As clsPatiInfo, objCurItem As clsBalanceItem, _
    ByRef cllBillPro As Collection, Optional objBalanceItem_Out As clsBalanceItem, _
    Optional blnSaveed_Out As Boolean, Optional blnCurInput As Boolean, Optional blnChangeMoney_Out As Boolean, _
    Optional blnStartOracleTrans As Boolean = False, Optional ByVal int险类 As Integer, Optional ByVal bln电子票据 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     objcurDelBalanceItems-当前退费的结算项
    '     blnCurInput-是否当前输入项
    '     blnStartOracleTrans-是否启用了oracle事务
    '出参:
    '     objBalanceItems_Out-返回已经退费成功的项目
    '     blnSaveed_Out-是否发生数据保存
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, cllBeforePro As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim rsBalance As ADODB.Recordset, rsExpend As ADODB.Recordset
    Dim blnTrans As Boolean, blnCallInterfance As Boolean
    Dim intStatus As Integer, strErrMsg As String, strBalanceIDs As String
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strCardNo As String, dblMoney As Double, dblDelMoney As Double, str结算方式  As String
    Dim i As Long, lngCount As Long, lng结帐ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strExpendXML As String, strTemp As String
    Dim objCard As Card
    Dim dblMoneySum As Double, dblCurMoney As Double
    Dim cllSquare As Collection
 
    Err = 0: On Error GoTo Errhand:
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    blnSaveed_Out = False
    
    If objCurItem Is Nothing Then Exit Function
    Set objCard = objCurItem.objCard
    If objCard Is Nothing Then Exit Function
    
    '非一卡通支付,直接返回Or objCard.消费卡
    If objCard.接口序号 <= 0 Or objCurItem.消费卡 Then zlThird_ReturnMoney = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    If Not objCurItem.是否保存 Then
        If zlGetSaveSQLfromItem(objPati, objCurItem, IIf(mbln结算异步控制, 1, 0), cllPro, True, , int险类, bln电子票据) = False Then Exit Function
    End If
    dblDelMoney = objCurItem.结算金额
    dblMoney = objCurItem.结算金额
    If dblMoney = 0 Then
        MsgBox "当前退款为零了，不允许再次退款", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
                
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, IIf(mbln结算异步控制, False, True)
           
    If mbln结算异步控制 Then         '异步处理，需要事先处理校对标志为2(减少接口调用后更新)
        blnTrans = False: blnSaveed_Out = True
        Set cllBeforePro = New Collection
        If zlGetSaveSQLfromItem(objPati, objCurItem, 0, cllBeforePro, False, True, int险类, bln电子票据) = False Then Exit Function
        blnTrans = True
        zlExecuteProcedureArrAy cllBeforePro, mfrmMain.Caption, True
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
    '    blnResolveXMLToRecord-是否解析XML串给记录集(rsBalance_Out,rsExpend_Out）
    '出参: strSwapNo-交易流水号(退款交易流水号)
    '      strSwapMemo-交易说明(退款交易说明)
    '    intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
    '    strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
    '      strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    '       strExpendXML：10.35.90以后才有此参数：格式如下
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strBalanceIDs = "1|" & objCurItem.预交ID
    Set cllPro = New Collection
    strSwapExtendInfor = "1|" & objCurItem.ID
    strCardNo = objCurItem.卡号
    strSwapNO = objCurItem.交易流水号
    strSwapMemo = objCurItem.交易说明
    
    If gobjSquare.objSquareCard.zlReturnMoney(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, strCardNo, _
        strBalanceIDs, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strExpendXML, True, rsBalance, rsExpend, intStatus, strErrMsg) = False Then
  
        If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
        
        If strErrMsg <> "" Then
            MsgBox "注意:" & vbCrLf & "    当前退款失败，错误原因如下:" & vbCrLf & strErrMsg, vbOKOnly, gstrSysName
        End If
        
        If intStatus = 1 Or objCurItem.ID <> 0 Then    '不能直接删除当前结算，可能后期还有其他操作
            If blnSaveed_Out Then
                objCurItem.是否保存 = True
                objCurItem.是否允许编辑 = False
            End If
            Set objBalanceItem_Out = objCurItem
            Exit Function
        End If
        
        If mbln结算异步控制 Then
             Set cllUpdate = New Collection
             Call zlGetDeleteSQL(objCurItem.单据号, 0, cllUpdate)
             Call zlExecuteProcedureArrAy(cllUpdate, mfrmMain.Caption, False, blnTrans)
             blnTrans = False
        End If
        Set objBalanceItem_Out = objCurItem
        Exit Function
    End If
    
    Set objBalanceItem_Out = objCurItem
    
    '更新数据
    lngCount = 0
    If Not rsBalance Is Nothing Then
    
        If rsBalance.RecordCount > 2 Then
            MsgBox "充值目前只支持接收一种结算信息返回，请检查一卡通接口", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        Set objBalanceItem_Out = objCurItem
        dblMoney = 0
        If rsBalance.RecordCount <> 0 Then rsBalance.MoveFirst
        With rsBalance
            If Val(Nvl(rsBalance!交易金额)) <> objCurItem.结算金额 Then
                MsgBox "返回的结算金额与当前要求退款的金额不一致，请与接口商联系!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
             
            With objBalanceItem_Out
                .卡类别ID = IIf(Val(Nvl(rsBalance!是否普通结算)) = 0, Val(Nvl(rsBalance!卡类别ID)), 0)
                .消费卡 = False
                .消费卡ID = 0
                .卡号 = IIf(Nvl(rsBalance!卡号) = "", objCurItem.卡号, Nvl(rsBalance!卡号))
                .退款交易流水号 = Nvl(rsBalance!交易流水号)
                .退款交易说明 = Nvl(rsBalance!交易说明)
                .结算方式 = Nvl(rsBalance!结算方式)
                .结算金额 = Val(Nvl(rsBalance!交易金额))
                .密码 = objCurItem.密码
                .是否密文 = objCurItem.是否密文
                .限制类别 = objCurItem.限制类别
                .帐户余额 = objCurItem.帐户余额
                .结算号码 = Nvl(rsBalance!结算号码)
                .结算摘要 = Nvl(rsBalance!结算摘要)
                .校对标志 = 0
                If .卡号 <> "" Then strCardNo = .卡号
                If strSwapNO = "" And .交易流水号 <> "" Then
                    strSwapNO = .交易流水号
                    If strSwapMemo = "" Then strSwapNO = .交易流水号
                End If
                dblMoney = dblMoney + .结算金额
            End With
        End With
        blnChangeMoney_Out = dblDelMoney <> dblMoney
    Else
        blnChangeMoney_Out = False
    End If
    
    If Not rsBalance Is Nothing Then
        If rsBalance.RecordCount > 1 Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "交易失败，不支持多种结算方式支付！", vbInformation, gstrSysName
            Exit Function
        End If
        If Val(rsBalance!交易金额) <> dblMoney Then
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "交易失败，实际退款金额与本次退款金额不符！", vbInformation, gstrSysName
            Exit Function
        End If
        '获取更新Sql
        Set cllUpdate = New Collection
        If GetThirdUpdateSQL(objCurItem.ID, "", Nvl(rsBalance!结算方式), 0, "", "", "", "", Nvl(rsBalance!是否普通结算, 0), cllUpdate, True, bln电子票据, int险类) Then
            zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, True, True
        End If
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    
    zlThird_ReturnMoney = True
    'blnTrans = True
'    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, , Not mbln结算异步控制 Or blnTrans
'    blnTrans = False

    objBalanceItem_Out.是否保存 = True
    objBalanceItem_Out.是否允许编辑 = False
    objBalanceItem_Out.是否允许删除 = False
    objBalanceItem_Out.是否结算 = True
    objBalanceItem_Out.校对标志 = 2
    blnSaveed_Out = True
    
    Set cllBillPro = New Collection
    blnTrans = False
    
    Call zlGetThreeSwapExpendSQLToCollect(objCard, objCurItem.预交ID, strCardNo, rsExpend, cllThreeSwap, True)
    
    Err = 0: On Error GoTo ErrOther:
    blnTrans = True
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption & "-执行三方接口"
    blnTrans = False
    
    Exit Function
Errhand:
    If blnTrans And blnCallInterfance = False Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
  Resume
    End If
    Call SaveErrLog
    If blnTrans Then gcnOracle.RollbackTrans
    Exit Function
ErrOther:
    If ErrCenter() = 1 Then
  Resume
    End If
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
End Function

Public Function zlGetThreeSwapExpendSQLToCollect(objCard As Card, ByVal lng结帐ID As Long, ByVal str卡号 As String, ByVal rsExpend As ADODB.Recordset, _
    ByRef cllPro As Collection, Optional bln预交 As Boolean = False, Optional lng预交ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取扩展信息保存的SQL给集合
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-27 17:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str交易信息 As String, strTemp As String
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
                        strSQL = "Zl_三方结算交易_Insert("
                        '卡类别id_In 病人预交记录.卡类别id%Type,
                        strSQL = strSQL & "" & objCard.接口序号 & ","
                        '消费卡_In   Number,
                        strSQL = strSQL & "" & IIf(objCard.消费卡, 1, 0) & ","
                        '卡号_In     病人预交记录.卡号%Type,
                        strSQL = strSQL & "'" & str卡号 & "',"
                        '结帐ids_In  Varchar2,
                        strSQL = strSQL & "'" & lng结帐ID & "',"
                        '交易信息_In Varchar2:交易项目|交易内容||...
                        strSQL = strSQL & "'" & str交易信息 & "',"
                        '预交款缴款_In Number := 0
                        strSQL = strSQL & IIf(bln预交, "1", "0") & ",Null,"
                        strSQL = strSQL & IIf(lng预交ID = 0, "NULL", lng预交ID) & ")"
                        zlAddArray cllPro, strSQL
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
        strSQL = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & objCard.接口序号 & ","
        '消费卡_In   Number,
        strSQL = strSQL & "" & IIf(objCard.消费卡, 1, 0) & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSQL = strSQL & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSQL = strSQL & "'" & lng结帐ID & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSQL = strSQL & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSQL = strSQL & IIf(bln预交, "1", "0") & ",Null,"
        strSQL = strSQL & IIf(lng预交ID = 0, "NULL", lng预交ID) & ")"
        zlAddArray cllPro, strSQL
    End If
    zlGetThreeSwapExpendSQLToCollect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlThird_Payment_IsValied(ByVal objPati As clsPatiInfo, ByRef objCurItem As clsBalanceItem, _
    ByRef objBalanceItem_Out As clsBalanceItem, ByRef dbl帐户余额_Out As Double, Optional byt操作类型 As Byte = 0, Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查三方帐户是否有效
    '入参:frmMain-调用的主窗体
    '     objPati-病人信息对象
    '     objCard-当前卡对象
    '     rsClassFee-可选参数，表示分收费类别统计的费用，主要应用于刷消费卡的处理(字段;收费类别,金额)
    '     byt操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    '     blnCurInput-是否当前输入项
    '出参:objBalanceItem_Out-返回结算数据对象集
    '     dbl帐户余额_out-返回帐户余额
    '返回:数据合法返回true(（含非三卡或消费卡或Card对象 is nothig),否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 10:30:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem, strXMLExpend As String
    Dim str卡号 As String, str密码 As String
    Dim str结算方式 As String, dbl结算金额 As Double, varData As Variant
    Dim i As Long, dbl未付金额 As Double, dblMoney As Double
    Dim objCard As Card
    On Error GoTo errHandle
    
    If objCurItem Is Nothing Then Exit Function
    
    objCard = objCurItem.objCard
    If objCard Is Nothing Then zlThird_Payment_IsValied = True: Exit Function
    If objCard.接口序号 <= 0 Or objCurItem.消费卡 Then zlThird_Payment_IsValied = True: Exit Function
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
        Exit Function
    End If
        
        
    If objBalanceItem_Out Is Nothing Then Set objBalanceItem_Out = New clsBalanceItem
    
     
    dblMoney = objCurItem.结算金额
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dblMoney <= 0 Then
        MsgBox "本次消费金额必须大于零，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
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
    If gobjSquare.objSquareCard.zlBrushCard(mfrmMain, mlngModule, Nothing, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, False, True, False, False, Nothing, False, False, strXMLExpend, _
             "", objPati.病人ID, str结算方式, dbl结算金额) = False Then Exit Function
       
    
    If str结算方式 = "" Then
        str结算方式 = objCard.结算方式
        If dbl结算金额 = 0 Then dbl结算金额 = dblMoney
    End If
    
    If dbl结算金额 = 0 Then
        MsgBox "本次消费金额不能为零(接口返回0)，请与接口开发商或系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    Set objBalanceItem_Out = objCurItem
    With objBalanceItem_Out
        Set .objCard = objCard
        .结算方式 = str结算方式
        .结算金额 = dbl结算金额
        .卡号 = str卡号
        .卡类别ID = objCard.接口序号
        .密码 = str密码
        .结算类型 = 3
        .限制类别 = ""
        .是否允许编辑 = True
        .是否允许删除 = True
        .是否结算 = False
        .是否密文 = IIf(objCard.卡号密文规则 <> "", True, False)
     End With
      
    '保存前,一些数据检查
    'Public Function zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
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
    '       strXMLExpend-(XML串:验证密码:自助机用)
    '出参:
    '   strXMLExpend-(XML串:错误信息)
    '返回:扣款合法,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 16:42:43
    '说明:
    '   在调用扣款前，由于存在Oracle事务问题， 所以再调用扣款交易前， _
    '   先进行数据的合法性检查,以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = ""
    If gobjSquare.objSquareCard.zlPaymentCheck(mfrmMain, mlngModule, objCard.接口序号, _
    objCard.消费卡, str卡号, dblMoney, objCurItem.单据号, strXMLExpend) = False Then Exit Function
    
    strXMLExpend = ""
    'Public Function zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
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
    dbl帐户余额_Out = 0

    If gobjSquare.objSquareCard.zlGetAccountMoney(mfrmMain, mlngModule, objCard.接口序号, _
     str卡号, strXMLExpend, dbl帐户余额_Out, objCard.消费卡) = False Then Exit Function
    If dbl帐户余额_Out <> 0 Then
        If dbl帐户余额_Out < dbl结算金额 Then
            MsgBox objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        For Each objItem In objBalanceItem_Out
            objItem.帐户余额 = dbl帐户余额_Out
        Next
    End If
    zlThird_Payment_IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlThird_TransferAccounts_IsValied(ByVal objPati As clsPatiInfo, objCurTranItem As clsBalanceItem, _
    ByRef objCurTranItem_out As clsBalanceItem, Optional ByRef dbl帐户余额_Out As Double, Optional ByRef strErrMsg_out As String = "", Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行转帐较交易接口的合法性检查
    '入参:objPati-病人信息
    '    objCard-当前卡对象
    '    objCurTranItem-当前转帐信息
    '    blnCurInput-当前输入项
    '出参:objCurTranItem_out-转帐信息集
    '     strErrMsg_Out-返回的错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-29 10:08:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String, bln退现 As Boolean, str结算方式 As String, dbl结算金额 As Double
    Dim str卡号 As String, str密码 As String, dblMoney As Double, lng结帐ID As Long
    Dim objDelItem As clsBalanceItem, dbl未退金额 As Double, objCurTranItems As clsBalanceItems
    Dim objCard As Card
    On Error GoTo errHandle
    
    dbl帐户余额_Out = 0
    If objCurTranItem Is Nothing Then Exit Function
    
    Set objCard = objCurTranItem.objCard
    If objCard Is Nothing Then zlThird_TransferAccounts_IsValied = True: Exit Function
    If objCard.接口序号 <= 0 Or objCard.消费卡 = True Then zlThird_TransferAccounts_IsValied = True: Exit Function
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
        Exit Function
    End If
    
    Set objCurTranItems = objCurTranItem.objTag
    
    If objCurTranItems Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If objCurTranItems.Count = 0 Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
      
    dblMoney = objCurTranItem.结算金额
    For Each objDelItem In objCurTranItems
        dbl未退金额 = roundEx(dbl未退金额 + objDelItem.未退金额, 2)
    Next
    
    If dblMoney = 0 Then
        If dbl未退金额 = 0 Then zlThird_TransferAccounts_IsValied = True: Exit Function
        Screen.MousePointer = 0
        MsgBox "退款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dbl未退金额 = 0 Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & objCard.名称 & "已经退完，不能再退！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    If Abs(dblMoney) > Format(Abs(dbl未退金额), "0.00") And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCurTranItem_out Is Nothing Then Set objCurTranItem_out = New clsBalanceItem
    
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
    
    strXMLExpend = "<IN><CZLX>1</CZLX></IN>"
    '操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    If gobjSquare.objSquareCard.zlBrushCard(mfrmMain, mlngModule, Nothing, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, True, True, bln退现, True, Nothing, False, False, strXMLExpend, str结算方式, dbl结算金额) = False Then Exit Function
            
    
    If str结算方式 <> "" Then
        If dblMoney <> dbl结算金额 Then
            strErrMsg_out = "当前转帐金额为:" & dblMoney & "与接口返回的转帐金额不符，禁止转帐操作!"
            Exit Function
        End If
    End If
    
    
    '调用转帐接口
    '    7.1.    zltransferAccountsCheck(转帐检查接口)
    'zlTransferAccountsCheck 转帐检查接口
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  HIS调用模块号
    'lngCardTypeID   Long    In  卡类别ID
    'strCardNo   String  In  卡号
    'dblMoney    Double  In  转帐金额(代扣时为负数)
    'strBalanceIDs   String  In  结帐IDs，多个用逗号分离，表示本次对哪此收费项目进行重新医保补结算
    'strXMLExpend String In   XML串:
    '                            <IN>
    '                                <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务,4-门诊退费业务或病人预交余额退款.
    '                            </IN>
    '                    Out  XML串:
    '                            <OUT>
    '                               <ERRMSG>错误信息</ERRMSG >
    '                            </OUT>
    '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
    '说明:
    '１. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '构造XML串
    strXMLExpend = "<IN><CZLX>4</CZLX></IN>"
    If gobjSquare.objSquareCard.zltransferAccountsCheck(mfrmMain, mlngModule, objCard.接口序号, _
        str卡号, dblMoney, lng结帐ID, strXMLExpend, strErrMsg_out) = False Then Exit Function
         
    
    strXMLExpend = ""
    'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    '    ByVal strCardTypeID As Long, _
    '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '       入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '       出参:dblMoney-返回帐户余额
    Call gobjSquare.objSquareCard.zlGetAccountMoney(mfrmMain, mlngModule, objCard.接口序号, _
          str卡号, strXMLExpend, dbl帐户余额_Out, objCard.消费卡)
          
    Set objDelItem = New clsBalanceItem
    With objDelItem
        .结算方式 = IIf(str结算方式 = "", objCard.结算方式, str结算方式)
        .结算金额 = -1 * dblMoney
        .卡类别ID = objCard.接口序号
        .消费卡 = objCard.消费卡
        .卡号 = str卡号
        .密码 = str密码
        .帐户余额 = dbl帐户余额_Out
        .是否密文 = objCard.卡号密文规则 <> ""
        .是否转帐 = True
    End With
    Set objCurTranItem_out = objDelItem
    zlThird_TransferAccounts_IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
  
Public Function zlInsure_Check(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    
    On Error GoTo errHandle
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then Exit Function
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    blnMedicareCheck = True
    varData = Split(str保险结算, "||"): varData1 = Split(strAdvance, "||")
    
    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsure_Check = blnMedicareCheck
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlGetMedicareStr(ByVal objYbBalanceItems As clsBalanceItems, ByRef cur个人帐户 As Currency, cur医保基金 As Currency) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回保险结算方式串,"结算方式|金额||...."
    '入参:objYbBalanceItems-医保结算信息集
    '出参:cur个人帐户-个人帐户
    '     cur医保基金-医保基金
    '返回:返回保险结算方式串,"结算方式|金额||...."
    '编制:刘兴洪
    '日期:2015-01-13 15:16:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strTemp As String
    Dim dblMoney As Double, int类型 As Integer
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle
    strTemp = ""
    cur个人帐户 = 0: cur医保基金 = 0
    For Each objItem In objYbBalanceItems
        If objItem.结算性质 = 3 Or objItem.结算性质 = 4 Then
            dblMoney = objItem.结算金额
            strTemp = strTemp & "||" & objItem.结算方式 & "|" & Format(dblMoney, gstrDec)
            If objItem.结算性质 = 3 Then cur个人帐户 = cur个人帐户 + dblMoney
            If objItem.结算性质 = 4 Then cur医保基金 = cur医保基金 + dblMoney
        End If
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 3)
    zlGetMedicareStr = strTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSquare_ReturnMoneySQL(ByVal objPati As clsPatiInfo, ByRef cllBillPro As Collection, _
    ByRef objCurItem As clsBalanceItem, Optional ByVal bln完成结算 As Boolean = False, _
    Optional ByVal int险类 As Integer, Optional ByVal bln电子票据 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行消费卡退款及保存业务
    '入参:objCard-当前支付的卡对象
    '    objPati-病人对象集
    '    cllBillPro-执行的相关过程
    '    objCurItem-当前结算项(objTag:消费卡刷卡洋细信息)
    '    blnCommit-是否提交数据
    '出参:cllBillPro-返回执行扣款的SQL
    '     objBalanceItems_Out-返回消费卡扣款集
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-23 17:02:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItems As clsBalanceItems, objItem As clsBalanceItem
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str消费卡结算 As String, strSQL As String
    Dim objCard As Card
    Dim rsTemp As ADODB.Recordset, dblMoney As Double, dblMoneySum As Double, dblCurMoney As Double
    
    
    If objCurItem Is Nothing Then Exit Function
    Set objCard = objCurItem.objCard
    If objCard Is Nothing Then Exit Function
    If Not objCard.消费卡 Then zlSquare_ReturnMoneySQL = True: Exit Function
    
    On Error GoTo errHandle
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    Set objItems = objCurItem.objTag
    If objItems Is Nothing Then
        MsgBox objCard.名称 & "无结算信息，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objItems.Count = 0 Then
        MsgBox objCard.名称 & "无结算信息，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '暂未处理
'    With objCurItem
'        If .ID = 0 Then
'            .ID = zlDatabase.GetNextId("病人预交记录")
'            .收款时间 = zlDatabase.Currentdate
'            .单据号 = zlDatabase.GetNextNo(11)
'        End If
'    End With
    
    If Not objCurItem.是否保存 Then
        If zlGetSaveSQLfromItem(objPati, objCurItem, 0, cllPro, True, , int险类, bln电子票据) = False Then Exit Function
    
    End If
'    dblDelMoney = objCurItem.结算金额
    dblMoney = objCurItem.结算金额
    If dblMoney = 0 Then
        MsgBox "当前退款为零了，不允许再次退款", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, True

    
    '冲销消费卡金额
    strSQL = _
        "Select 接口编号, 消费卡id, 卡号, -1 * Sum(应收金额) As 应收金额" & vbNewLine & _
        "From 病人卡结算记录" & vbNewLine & _
        "Where 记录性质 = 4 And 结算id = [1]" & vbNewLine & _
        "Group By 接口编号, 消费卡id, 卡号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "消费卡", objCurItem.预交ID)
    
    '可能使用了多张消费卡
    dblMoneySum = objCurItem.结算金额
    Do While Not rsTemp.EOF
         If Val(Nvl(rsTemp!应收金额)) < dblMoneySum Then
          dblCurMoney = Val(Nvl(rsTemp!应收金额))
          dblMoneySum = Round(dblMoneySum - Val(Nvl(rsTemp!应收金额)), 6)
        Else
          dblCurMoney = dblMoneySum
          dblMoneySum = 0
        End If
        
        'Zl_病人卡结算记录_退款
        strSQL = "Zl_病人卡结算记录_退款("
        '  接口编号_In   消费卡类别目录.编号%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!接口编号)) & ","
        '  卡号_In       消费卡信息.卡号%Type,
        strSQL = strSQL & "'" & Nvl(rsTemp!卡号) & "',"
        '  消费卡id_In   消费卡信息.Id%Type,
        strSQL = strSQL & "" & Val(Nvl(rsTemp!消费卡ID)) & ","
        '  结算金额_In   病人卡结算记录.应收金额%Type,
        strSQL = strSQL & "" & dblCurMoney & ","
        '  原预交id_In   病人卡结算记录.结算id%Type,
        strSQL = strSQL & "" & objCurItem.预交ID & ","
        '  新预交id_In   病人卡结算记录.结算id%Type,
        strSQL = strSQL & "" & objCurItem.ID & ","
        '  操作员编号_In 病人卡结算记录.操作员编号%Type,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In 病人卡结算记录.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  退款时间_In   病人预交记录.收款时间%Type
        strSQL = strSQL & "To_Date('" & Format(objCurItem.收款时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"

'            zlDatabase.ExecuteProcedure strSQL, ""
        zlAddArray cllBillPro, strSQL
        If dblMoneySum = 0 Then Exit Do
        rsTemp.MoveNext
    Loop
    If dblMoneySum > 0 Then
        gcnOracle.RollbackTrans
        MsgBox "剩余可退金额(" & Format(dblMoney - dblMoneySum, "0.00") & ")不足本次退款金额(" & _
            Format(dblMoney, "0.00") & ")，不能退费！", vbInformation, gstrSysName
        Exit Function
    End If
    
    zlExecuteProcedureArrAy cllBillPro, mfrmMain.Caption, True, True
'    zlAddArray cllBillPro, strSQL
    gcnOracle.CommitTrans
    
    objCurItem.是否结算 = True
    objCurItem.是否允许编辑 = False
    objCurItem.是否保存 = True
    objCurItem.是否允许删除 = False

    
    zlSquare_ReturnMoneySQL = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlSquareDel_CheckValied(ByVal objPati As clsPatiInfo, ByRef objCurItem As clsBalanceItem, ByRef objOldItems As clsBalanceItems, _
    ByVal rsMoney As ADODB.Recordset, ByRef objCurItem_Out As clsBalanceItem) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡退费检查
    '入参:objCurItem-当前退款的消费卡信息(objTag:退款明细)
    '     objOldItems-已经刷卡的结算项
    '出参:objBalanceItems_Out-返回刷卡对象
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-23 11:07:58
    '说明:同步验证了接口和刷卡接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTotal As Double, i As Integer, j As Integer, blnFind As Boolean
    Dim cllBalance As Collection, cllSquareBalance As Collection
    Dim objItem As clsBalanceItem, objItems As clsBalanceItems, dblDelMoney As Double
    Dim strXmlIn As String, objCard As Card
    
    On Error GoTo errHandle
    If objCurItem Is Nothing Then Exit Function
    Set objCard = objCurItem.objCard
    If objCard Is Nothing Then Exit Function
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then zlSquareDel_CheckValied = True: Exit Function
    
    
    Set objItems = objCurItem.objTag
    If objItems Is Nothing Then
        MsgBox objCard.名称 & "无结算信息，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objItems.Count = 0 Then
        MsgBox objCard.名称 & "无结算信息，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objOldItems Is Nothing Then Set objOldItems = New clsBalanceItems
    
    dblDelMoney = objCurItem.结算金额
    dblTotal = 0
    For Each objItem In objItems
        dblTotal = roundEx(dblTotal + objItem.结算金额, 2)
    Next
    
    If roundEx(dblTotal, 2) < roundEx(dblDelMoney, 2) Then
        MsgBox "注意:" & vbCrLf & "   输入的退款金额大于了" & objCard.结算方式 & "的未退金额,请检查!" & vbCrLf & _
               "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
               "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
 
    
    Set cllBalance = New Collection
    For Each objItem In objOldItems
        'dblBrushMoney = GetSquareBrushMoney(objCard.接口序号, Val(Nvl(!消费卡ID)), Nvl(!卡号))
        'array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
        cllSquareBalance.Add Array(objCard.接口序号, objItem.卡类别ID, 0, objItem.卡号, "", objItem.限制类别, IIf(objItem.是否密文, 1, 0), objItem.结算金额)
    Next
    
    
    If mbln消费卡退费验卡 Then
       '弹出刷卡界面
        'zlBrushCard(frmMain As Object, _
        'ByVal lngModule As Long, _
        'ByVal rsClassMoney As ADODB.Recordset, _
        'ByVal lngCardTypeID As Long, _
        'ByVal bln消费卡 As Boolean, _
        'ByVal strPatiName As String, ByVal strSex As String, _
        'ByVal strOld As String, ByVal dbl金额 As Double, _
        'Optional ByRef strCardNo As String, _
        'Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean) As Boolean
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, objCard.接口序号, _
            objCard.消费卡, objPati.姓名, objPati.性别, _
            objPati.年龄, dblDelMoney, "", "", _
            True, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
        Set cllBalance = cllSquareBalance
    End If
    
    
    Set objItems = New clsBalanceItems
    
    '卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额
    For i = 1 To cllBalance.Count
        blnFind = False
        For Each objItem In objOldItems
            If objItem.消费卡ID = Val(cllBalance(i)(1)) Then
                objItem.结算金额 = Val(cllBalance(i)(2))
                objItem.卡号 = Trim(cllBalance(i)(3))
                objItem.密码 = Trim(cllBalance(i)(4))
                objItem.是否密文 = Val(cllBalance(i)(5)) = 1
                objItem.剩余金额 = Val(cllBalance(i)(6))
                Call objItems.AddItem(objItem)
                objItems.结算金额 = objItems.结算金额 + objItem.结算金额
                blnFind = True
                Exit For
            End If
        Next
        
        If blnFind = False Then
            Set objItem = New clsBalanceItem
            With objItem
                .objCard = objCard
                .卡类别ID = objCard.接口序号
                .结算方式 = objCard.结算方式
                .结算性质 = objCard.结算性质
                .消费卡 = Val(cllBalance(i)(1))
                .结算金额 = Val(cllBalance(i)(2))
                .卡号 = Trim(cllBalance(i)(3))
                .密码 = Trim(cllBalance(i)(4))
                .是否密文 = Val(cllBalance(i)(5)) = 1
                .剩余金额 = Val(cllBalance(i)(6))
            End With
            Call objItems.AddItem(objItem)
            objItems.结算金额 = objItems.结算金额 + objItem.结算金额
        End If
    Next
    Set objCurItem.objTag = objItems
    Set objCurItem_Out = objCurItem
    zlSquareDel_CheckValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlThirdSwapIsSwapNOCall(ByVal lngCardTypeID As Long, bln消费卡 As Boolean, Optional strErrMsg As String, Optional strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否按交易号分别调用接口
    '入参:
    '出参:
    '返回:true:按交易流水号调用接口;False-一次批量调用调用口
    '编制:刘兴洪
    '日期:2018-04-16 17:04:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '   ZlGetParaConfig（ ByVal frmMain As Object, _
    '    ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, ByVal intPara As Integer, _
    '    Optional strErrMsg As String, Optional strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取接口参数
    '入参: frmMain-调用的主窗体
    '       intPara: 包含如下值
    '                1-刷卡和支付在同一页面:true-新模式；False-旧模式
    '       strExpend-扩展参数，暂留，现传为空
    '出参:strErrMsg-返回的错误信息
    '       strExpend-扩展参数，暂留，现传为空
    '返回:函数返回True:调用成功,False:调用失败
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '''参数值  参数名  参数说明    应用场合
    ''1   刷卡和支付在同一页面：主要应用在自助系统中，可以通过此参数，调用界面的一些布局。  自助系统
    ''2   按交易流水号分别调用退款接口    主要是在退款时，如果此参数返回True,则调用zlReturnMoney接口.否则调用returnMultMoney接口：病人结帐;补充结算 (含自助补允结算);病人预交款管理 (余额退款)
    zlThirdSwapIsSwapNOCall = gobjSquare.objSquareCard.ZlGetParaConfig(mfrmMain, lngCardTypeID, bln消费卡, 2, strErrMsg, strExpend)
End Function

Public Function zlDelBalanceItemFromItemObject(ByVal objItem As clsBalanceItem, Optional ByRef cllSaveSqL_Out As Collection, Optional blnExcute As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算对象,获取删除结算项的SQL
    '入参:blnExcute-是否进行执行操作
    '出参:cllSaveSqL_Out-返回保存的SQL
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-14 17:13:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans As Boolean, strSQL As String
    On Error GoTo errHandle
    
    If cllSaveSqL_Out Is Nothing Then Set cllSaveSqL_Out = New Collection
    Call zlGetDeleteSQL(objItem.单据号, 2, cllSaveSqL_Out)
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


Private Sub Class_Initialize()
    mstrOnCardPrivs = ";" & GetPrivFunc(glngSys, 1151) & ";"   '一卡通相关接口权限
End Sub

Private Sub Class_Terminate()
    Set mfrmMain = Nothing
    Set mobjICCard = Nothing
    Set mrs结算方式 = Nothing
End Sub


Public Function zlGetCardFromCardType(ByVal lng卡类别ID As Long, bln消费卡 As Boolean, ByVal str结算方式 As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡类别ID获取卡对象
    '入参:lng卡类别ID-卡类别ID
    '     bln消费卡-是否消费卡
    '     str结算方式-结算方式
    '出参:
    '返回:成功卡对象
    '编制:刘兴洪
    '日期:2018-04-02 14:29:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As New Card
    On Error GoTo errHandle
    If lng卡类别ID <> 0 Then
        'zlGetCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean,  ByRef objCard As Card) As Boolean
        If gobjSquare.objSquareCard.zlGetCard(lng卡类别ID, bln消费卡, objCard) = False Then
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

Public Function zlGetCardFromBalanceName(ByVal str结算方式 As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算方式获取卡对象
    '入参:str结算方式-结算方式名称
    '出参:
    '返回:返回卡对象
    '编制:刘兴洪
    '日期:2018-03-30 15:39:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo errHandle
    Set objCard = New Card
    Set rsTemp = zlGet结算方式
    With objCard
        .结算方式 = str结算方式
        .名称 = str结算方式
        rsTemp.Filter = "名称='" & str结算方式 & "'"
        If Not rsTemp.EOF Then
            .结算性质 = Val(Nvl(rsTemp!性质))
        End If
    End With
    Set zlGetCardFromBalanceName = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet结算方式() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算方式
    '返回:返回结算方式信息集
    '编制:刘兴洪
    '日期:2018-03-29 17:35:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    If Not mrs结算方式 Is Nothing Then
        If mrs结算方式.State = 1 Then
            mrs结算方式.Filter = 0
            Set zlGet结算方式 = mrs结算方式: Exit Function
        End If
    End If
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select a.编码,a.名称, a.性质,b.应用场合,nvl(a.应付款,0) as 应付款,nvl(a.应收款,0) as 应收款,nvl(a.缺省标志,0) as 缺省,nvl(b.缺省标志,0) as  场合缺省" & vbNewLine & _
    "   From 结算方式 a, 结算方式应用 b" & vbNewLine & _
    "   Where a.名称 = b.结算方式(+)    "
        
    Set mrs结算方式 = zlDatabase.OpenSQLRecord(strSQL, "获取结算方式")
    Set zlGet结算方式 = mrs结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
Public Function zlAddUpdateSwapSQL(ByVal bln预交 As Boolean, ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    str卡号 As String, str交易流水号 As String, str交易说明 As String, _
    ByRef cllPro As Collection, Optional int校对标志 As Integer = 0, Optional lng关联交易ID As Long, Optional byt操作方式 As Byte = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新三方交易流水号和流水说明
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    '       byt操作方式:0-全部更新,1-只更新校对标志
    '出参:cllPro-返回SQL集
    '编制:刘兴洪
    '日期:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    'Zl_三方接口更新_Update
    strSQL = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  消费卡_In     Number,
    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSQL = strSQL & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSQL = strSQL & "'" & strIDs & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '  预交款缴款_In Number := 0,
    strSQL = strSQL & "" & IIf(bln预交, 1, 0) & ","
    '  退费标志_In   Number := 0,
    '退费标志 :1-退费;0-付费
    strSQL = strSQL & "0,"
    '  校对标志_In   Number := Null,
    strSQL = strSQL & "" & IIf(int校对标志 = 0, "NULL", int校对标志) & ","
    '  发送标志_In   Number := 0,
    strSQL = strSQL & "0,"
    '  消费卡管理_In Number := 0,
    strSQL = strSQL & "0,"
    '  关联交易id_In 病人预交记录.关联交易id%Type := Null
    strSQL = strSQL & "" & lng关联交易ID & ","
    '操作方式_In   Number := 0
    strSQL = strSQL & "" & byt操作方式 & ")"
    zlAddArray cllPro, strSQL
End Function


Public Function zlGetThirdMoneyInforRecordFromSwapID(ByVal str关联交易IDs As String, ByRef rsSwapRecord_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,获取相关的结算金额信息集
    '入参:str关联交易IDs-关联交易ID，多个用逗号分离
    '出参:rsSwapRecord_Out-返回关联交易ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-27 17:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strWhere As String, lng关联交易ID As Long
    
    On Error GoTo errHandle
    
    If InStr(str关联交易IDs, ",") > 0 Then
        strWhere = "And A.关联交易ID In (Select column_value From table(f_num2List([1])) "
        
    Else
       lng关联交易ID = Val(str关联交易IDs)
       strWhere = " And  A.关联交易ID =[2]"
    End If
    
    strSQL = "" & _
    "   Select 关联交易ID,卡类别ID,结算方式,交易流水号,交易说明, " & vbCrLf & _
    "          decode(sign(nvl(金额,0)),1,1,0)*nvl(金额,0)+ decode(sign(nvl(冲预交,0)),1,1,0)* nvl(冲预交,0) as 原始金额, " & _
    "          decode(sign(nvl(金额,0)),-1,1,0)*nvl(金额,0)+ decode(sign(nvl(冲预交,0)),-1,1,0)* nvl(冲预交,0) as 已退金额" & _
    "   From 病人预交记录 A " & _
    "   Where 1=1 " & strWhere
    strSQL = "" & _
    " Select 关联交易ID,卡类别ID,a.结算方式,a.交易流水号,a.交易说明, sum(原始金额) as 原始金额, sum(已退金额) as 已退金额, sum(原始金额)-sum(已退金额) as 剩余未退金额" & _
    " From (" & strSQL & ") A " & _
    " Group by a.关联交易ID,a.卡类别ID,a.结算方式,a.交易流水号,a.交易说明"
    
    Set rsSwapRecord_Out = zlDatabase.OpenSQLRecord(strSQL, "获取三方交易的原始金额及未退金额", str关联交易IDs, lng关联交易ID)
    zlGetThirdMoneyInforRecordFromSwapID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInsureCheck(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo errHandle
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then Exit Function
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    blnMedicareCheck = True
    varData = Split(str保险结算, "||"): varData1 = Split(strAdvance, "||")

    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsureCheck = blnMedicareCheck
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCopyNewItemFromBalanceItem(ByVal objOldItem As clsBalanceItem) As clsBalanceItem
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:复制一个新的Item
    '入参:objOldItem-旧的Item对象
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-19 14:14:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    
    
    On Error GoTo errHandle
    Set objItem = New clsBalanceItem
    If objOldItem Is Nothing Then
        Set objItem.objCard = New Card
        Set zlCopyNewItemFromBalanceItem = objItem: Exit Function
    End If
    
    With objItem
        Set .objCard = zlCopyNewCardFromCard(objOldItem.objCard)
        .Key = objOldItem.Key
        .Tag = objOldItem.Tag
        .关联交易ID = objOldItem.关联交易ID
        .交易流水号 = objOldItem.交易流水号
        .交易说明 = objOldItem.交易说明
        .缴款金额 = objOldItem.缴款金额
        .结算IDs = objOldItem.结算IDs
        .ID = objOldItem.ID
        .结算方式 = objOldItem.结算方式
        .结算号码 = objOldItem.结算号码
        .结算金额 = objOldItem.结算金额
        .结算类型 = objOldItem.结算类型
        .结算性质 = objOldItem.结算性质
        .结算摘要 = objOldItem.结算摘要
        .收款时间 = objOldItem.收款时间
        .卡号 = objOldItem.卡号
        .交易流水号 = objOldItem.交易流水号
        .交易说明 = objOldItem.交易说明
        .卡类别ID = objOldItem.卡类别ID
        .密码 = objOldItem.密码
        .是否结算 = objOldItem.是否结算
        .是否密文 = objOldItem.是否密文
        .是否缺省 = objOldItem.是否缺省
        .是否退款 = objOldItem.是否退款
        .是否预交 = objOldItem.是否预交
        .是否允许编辑 = objOldItem.是否允许编辑
        .是否允许删除 = objOldItem.是否允许删除
        .是否允许退现 = objOldItem.是否允许退现
        .是否保存 = objOldItem.是否保存
        .是否退款分交易 = objOldItem.是否退款分交易
        .未退金额 = objOldItem.未退金额
        .误差费 = objOldItem.误差费
        .限制类别 = objOldItem.限制类别
        .消费卡 = objOldItem.消费卡
        .消费卡ID = objOldItem.消费卡ID
        .校对标志 = objOldItem.校对标志
        .原始金额 = objOldItem.原始金额
        .帐户余额 = objOldItem.帐户余额
        .退款交易流水号 = objOldItem.退款交易流水号
        .退款交易说明 = objOldItem.退款交易说明
        .找补 = objOldItem.找补
        .行号 = objOldItem.行号
        .预交ID = objOldItem.预交ID
        .是否脱机医保 = objOldItem.是否脱机医保
        Set .objTag = Nothing
    End With

    Set zlCopyNewItemFromBalanceItem = objItem
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     Set zlCopyNewItemFromBalanceItem = objItem
End Function
Public Function zlCopyNewCardFromCard(ByVal objOldCard As Card) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一个卡对象，复制为新的卡对象
    '入参:objOldCard-旧卡
    '返回:返回新的Card对象
    '编制:刘兴洪
    '日期:2018-04-19 14:25:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Set objCard = New Card
    If objOldCard Is Nothing Then Set zlCopyNewCardFromCard = Nothing: Exit Function
    
    On Error GoTo errHandle
    With objOldCard
        objCard.备注 = .备注
        objCard.短名 = .短名
        objCard.功能键 = .功能键
        objCard.接口编码 = .接口编码
        objCard.接口程序名 = .接口程序名
        objCard.接口序号 = .接口序号
        objCard.结算方式 = .结算方式
        objCard.结算性质 = .结算性质
        objCard.卡号密文规则 = .卡号密文规则
        objCard.卡号长度 = .卡号长度
        objCard.卡号重复使用 = .卡号重复使用
        objCard.可否设置 = .可否设置
        objCard.快键 = .快键
        objCard.密码规则 = .密码规则
        objCard.密码输入限制 = .密码输入限制
        objCard.密码长度 = .密码长度
        objCard.密码长度限制 = .密码长度限制
        objCard.名称 = .名称
        objCard.模糊查找项 = .模糊查找项
        objCard.启用 = .启用
        objCard.前缀文本 = .前缀文本
        objCard.缺省标志 = .缺省标志
        objCard.设备是否启用回车 = .设备是否启用回车
        objCard.是否持卡消费 = .是否持卡消费
        objCard.是否存在帐户 = .是否存在帐户
        objCard.是否发卡 = .是否发卡
        objCard.是否非接触式读卡 = .是否非接触式读卡
        objCard.是否接触式读卡 = .是否接触式读卡
        objCard.是否模糊查找 = .是否模糊查找
        objCard.是否全退 = .是否全退
        objCard.是否缺省密码 = .是否缺省密码
        objCard.是否扫描 = .是否扫描
        objCard.是否刷卡 = .是否刷卡
        objCard.是否退款验卡 = .是否退款验卡
        objCard.是否退现 = .是否退现
        objCard.是否写卡 = .是否写卡
        objCard.是否严格控制 = .是否严格控制
        objCard.是否证件 = .是否证件
        objCard.是否制卡 = .是否制卡
        objCard.是否转帐及代扣 = .是否转帐及代扣
        objCard.是否自动读取 = .是否自动读取
        objCard.特定项目 = .特定项目
        objCard.图像标识 = .图像标识
        objCard.系统 = .系统
        objCard.消费卡 = .消费卡
        objCard.支付启用 = .支付启用
        objCard.支付图像标识 = .支付图像标识
        objCard.自动读取间隔 = .自动读取间隔
        objCard.自制卡 = .自制卡
    End With
    Set zlCopyNewCardFromCard = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    zlCopyNewCardFromCard = objCard
End Function

Public Function zlGetDeleteSQL(ByVal strNO As String, ByVal bytOpt As Byte, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:异常记录回退或删除
    '入参:strNo-余额退款NO
    '     bytOpt-0-删除异常充值单据，1-删除异常退款单据，2-删除异常余额退款单据
    '出参:cllPro-返回异常删除SQL给集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-30 17:03:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    If cllPro Is Nothing Then Set cllPro = New Collection
    strSQL = "Zl_病人预交异常记录_Delete("
    '  单据号_In       病人预交记录.No%Type
    strSQL = strSQL & "'" & strNO & "'," & bytOpt & ")"
    zlAddArray cllPro, strSQL
    
    zlGetDeleteSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReGetPayCards(ByVal str性质 As String, Optional ByVal str结算场合 As String = "预交款", Optional ByRef objPayCards_Out As Cards) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算信息
    ' 入参:性质(1,2等)
    '      str结算场合:结帐（预存款)等
    '出参:objPayCards_Out-返回有效的卡信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-31 12:49:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, rsTemp As ADODB.Recordset
    Dim objCards As Cards, objCard As Card
    Dim objPayCards As Cards, i As Long
    
    On Error GoTo errHandle
    
    Set objCards = New Cards: Set objPayCards = New Cards
    '性质:1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项, _
    '     6-费用折扣,7-一卡通结算,8-结算卡结算
      
    Set rsTemp = Get结算方式(str结算场合, str性质)
    If rsTemp.RecordCount = 0 Then
        Set rsTemp = Nothing
        Set mobjPayCards = objPayCards
        Set objPayCards_Out = objPayCards
        zlReGetPayCards = True
        Exit Function
    End If
    
    If Not gobjSquare Is Nothing Then Set objCards = gobjSquare.objSquareCard.zlGetCards(3)    '获取有效的三方帐户支付
    
    rsTemp.Filter = "性质<7"
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
        Do While Not .EOF
              Set objCard = New Card
              objCard.接口序号 = -1 * i
              objCard.接口编码 = !编码
              objCard.名称 = !名称
              objCard.结算方式 = !名称
              objCard.结算性质 = Val(Nvl(!性质))
              objCard.启用 = True
              objCard.是否刷卡 = 1
              objCard.缺省标志 = Val(Nvl(!缺省)) = 1
              objPayCards.Add objCard
              i = i + 1
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    If InStr(";" & mstrOnCardPrivs & ";", ";三方接口消费;") = 0 Then
        Set rsTemp = Nothing
        Set mobjPayCards = objPayCards
        Set objPayCards_Out = objPayCards
        zlReGetPayCards = True
        Exit Function
    End If
    
    rsTemp.Filter = "性质>=7 and 性质<9" '一卡通结算
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            For Each objCard In objCards
                If objCard.结算方式 = Nvl(!名称) Then
                    objCard.是否刷卡 = True
                    objCard.缺省标志 = Val(Nvl(!缺省)) = 1
                    objPayCards.Add objCard
                     i = i + 1
                    Exit For
                End If
            Next
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    rsTemp.Filter = 0
    Set rsTemp = Nothing
    Set mobjPayCards = objPayCards
    Set objPayCards_Out = objPayCards
    zlReGetPayCards = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

