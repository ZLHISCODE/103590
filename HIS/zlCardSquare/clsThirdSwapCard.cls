VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThirdSwapCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'     1.6  zlRemoveBalanceDataFromItems:根据结算项信息，从结算数据中移除
'     1.7  zlGetBalanceCancelSaveSQLFromItems:根据结帐信息集合，获取对应的作废SQL集
'     1.8  zlCheckIsStartFromCard-检查指定的卡是否启用
'     1.9  zlGetBalanceModeDatas：获取有效的结帐数据，主要是返回Cards对象和记录集
'     1.10 zlGetBalaceRecordsFromCards:根据结算信息对象,返回卡结算记录集
'     1.11 zlGetBalaceCardFromCards:获取指定类型的结算对象
'     2.12 zlGetRQCardTypeIDsFromCards：根据有效的支付卡对象集，获取有效的扫码付的所涉及的卡类别IDs
'     2.新一卡通消费相关接口:
'     2.1   zlThird_Payment_IsValied:扣款交易检查(同步调用了:zlBrushCard-->PaymentCheck-->zlGetAccountMoney接口)
'     2.2   zlThird_Payment:执行三方接口支付(调用 ：zlPaymentMoney)
'     2.3   zlThird_ReturnMoney_IsValied：三方卡退款合法性检查
'     2.4   zlThird_TransferAccounts:执行转帐交易接口
'     2.5   zlThirdReturnCashCheck:退现合法性检查
'     2.6   zlThirdSwapIsSwapNOCall:判断是否按交易流水号调用接口
'     2.7   zlThird_IsCancelFromItems:根据项目判断当前是否允许作废或删除

'     3.旧一卡通消费相关接口
'     3.1   zlOldOneCard_PaymentCheckIsValied:检查一卡通支付是否合法(旧一卡通)
'     3.2   zlOldOneCard_Payment：旧一卡通接口支付(调用：ICCard.PaymentSwap）
'     3.3   zlOldOneCard_ReturnCheckIsValied:检查一卡通回退是否合法(旧一卡通)
'     3.4   zlOldOneCard_RetrunMoney:执行退款操作(旧一卡通)
'     4.医保相关处理
'     4.1  zlGetYbPara:获取医保对数
'     4.2  zlInsure_InitPara:初始化医保参数
'     4.3  zlInsure_BalanceUpdate:医保结算数据更新
'     4.4  zlInsure_Check:医保预结算和医保结帐是否匹配
'     4.9  zlInsure_GetSelfBalance:获取医保帐户余额
'     5 消费卡处理
'     5.1 zlSquare_AddBalanceDatas-向结算数据信息中加入消费卡结算信息
'     5.2 zlSquare_PaymentSQL-获取执行消费卡扣款处理的SQL
'编制:刘兴洪
'日期:2018-03-26 15:06:02
'*********************************************************************************************************************************************
Private mstrOnCardPrivs As String
Private mfrmMain As Object, mlngModule As Long
Private mbln结算异步控制 As Boolean
Private mbyt刷卡缺省金额操作 As Byte
Private mstr脱机医保 As String
Private mobjPayCards As Cards '当前有效的支付方式
Private mstrTittle As String
Private mstrErrInfor As String, mstrPara As String, mstrFunName As String
Private mblnReturnValue As Boolean


 
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
Public Property Get GetStrPrivs() As String
    GetStrPrivs = mstrOnCardPrivs
End Property

Public Function zlGetRQCardTypeIDsFromCards(ByRef objPayCards As Cards) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据有效的支付卡对象集，获取有效的扫码付的所涉及的卡类别IDs
    '入参:objPayCards-有效的支付方式
    '返回:返回支持 的扫码符的卡类别IDs
    '编制:刘兴洪
    '日期:2019-03-07 14:13:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, strCardTypeIDs As String
    On Error GoTo errHandle
    If objPayCards Is Nothing Then Exit Function
    
    For Each objCard In objPayCards
        If objCard.是否支持扫码付 And objCard.消费卡 = False And objCard.接口序号 > 0 Then
            strCardTypeIDs = strCardTypeIDs & "," & objCard.接口序号
        End If
    Next
    If strCardTypeIDs <> "" Then strCardTypeIDs = Mid(strCardTypeIDs, 2)
    zlGetRQCardTypeIDsFromCards = strCardTypeIDs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalaceCardFromCards(ByVal intType As Integer, ByVal objCards As Cards) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定类型的卡
    '入参:intType-0-缺省支付方式;1-现金支付方式
    '返回:成功返回指定的卡,否则返回Nothing
    '编制:刘兴洪
    '日期:2018-11-22 11:50:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    
    On Error GoTo errHandle
    For Each objCard In objCards
        Select Case intType
        Case 0  '缺省
            If objCard.缺省标志 = True Then Set zlGetBalaceCardFromCards = objCard: Exit Function
        Case 1
            If objCard.结算性质 = 1 Then Set zlGetBalaceCardFromCards = objCard: Exit Function
        End Select
    Next
    Set zlGetBalaceCardFromCards = Nothing
    Exit Function
errHandle:
    Set zlGetBalaceCardFromCards = Nothing
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
    Dim objCard As Card, i As Long
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle
    
    If Not mobjPayCards Is Nothing Then
        For i = 1 To mobjPayCards.Count
            Set objCard = mobjPayCards(i)
            If objCard.结算方式 = str结算方式 Then
                Set zlGetCardFromBalanceName = objCard
                Exit Function
            End If
        Next
    End If
    
    Set objCard = New Card
    Set rsTemp = zlGet结算方式
    With objCard
        .结算方式 = str结算方式
        .名称 = str结算方式
        rsTemp.Filter = "名称='" & str结算方式 & "'"
        If Not rsTemp.EOF Then
            .结算性质 = Val(Nvl(rsTemp!性质))
            .缺省标志 = Val(Nvl(rsTemp!场合缺省)) = 1
        End If
    End With
    Set zlGetCardFromBalanceName = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlGetBalaceRecordsFromCards(ByVal objCards As Cards, ByRef rsBalance_Out As ADODB.Recordset) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算信息对象,返回卡结算记录集
    '入参:objCards-结算对象
    '出参:rsBalance_Out-返回结算信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-22 11:20:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, i As Long
    On Error GoTo errHandle
    
    Set rsBalance_Out = New ADODB.Recordset
    With rsBalance_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "卡类别ID", adDouble, 18, adFldIsNullable
        .Fields.Append "是否消费卡", adBigInt, 2
        .Fields.Append "结算方式", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "卡名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "是否缺省", adBigInt, 2
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    If objCards Is Nothing Then zlGetBalaceRecordsFromCards = True: Exit Function
    i = 1
    For Each objCard In objCards
        With rsBalance_Out
            .AddNew
            !序号 = i
            !卡类别ID = objCard.接口序号
            !是否消费卡 = IIf(objCard.消费卡, 1, 0)
            !结算方式 = objCard.结算方式
            !卡名称 = objCard.名称
            !是否缺省 = IIf(objCard.缺省标志, 1, 0)
            .Update
        End With
        i = i + 1
    Next
    zlGetBalaceRecordsFromCards = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function


Public Function zlGetBalanceModeDatas(ByVal bln含脱机医保 As Boolean, ByRef objPayCards_Out As Cards, _
    Optional ByRef rsBalace_Out As ADODB.Recordset, Optional ByRef objErrCard_Out As Card, _
    Optional ByRef objOfflineYBCards_Out As Cards, Optional ByRef strQRCodeTypeIDs_Out As String, _
    Optional str场合 As String = "就诊卡", Optional ByVal str性质 As String = "1,2,8") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的结算信息
    '入参:bln含脱机医保-返回的objPayCards_Out中是包含脱机医保,True-表示包包含,False-不包含（不包含时，objOfflineYBCards_Out返回Nothing )
    '
    '出参:objPayCards_Out-有效的支付方式
    '     objErrCard_Out-误差费结算对象
    '     rsBalace_Out-结算信息集
    '     objOfflineYBCards_Out-脱机医保对象
    '     strQRCodeTypeIDs_Out-返回本机支持的扫码付的卡类别IDs,多个用逗号分离
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-22 10:58:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    Dim objCards As Cards, objCard As Card
    Dim blnOnlyDeposit As Boolean, lng序号 As Long
    Dim objTemp As Card
    
    On Error GoTo errHandle
    
    strQRCodeTypeIDs_Out = ""
    Set objCards = New Cards: Set objPayCards_Out = New Cards: Set objOfflineYBCards_Out = New Cards
    Set objErrCard_Out = New Card
     
    strTmp = "1,2,8"
    Set rsBalace_Out = Get结算方式("就诊卡", strTmp)
    If rsBalace_Out.RecordCount = 0 Then zlGetBalanceModeDatas = True: Exit Function
 
    If Not gobjOneCardComLib Is Nothing Then
        '0-所有医疗卡;1-启用的医疗卡,2-所有存在三方账户的三方卡 3-启用的三方账户的医疗卡
       Set objCards = gobjOneCardComLib.zlGetCards(3)     '获取有效的三方帐户会付
    End If
 
 
    rsBalace_Out.Filter = "性质<7"
    i = 1
    With rsBalace_Out
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If (InStr(",3,4,", "," & Val(Nvl(!性质)) & ",") = 0) And Val(Nvl(!应付款)) <> 1 Then
                Set objCard = New Card
                objCard.接口序号 = -1 * i
                objCard.接口编码 = !编码
                objCard.名称 = !名称
                objCard.结算方式 = !名称
                objCard.结算性质 = Val(Nvl(!性质))
                objCard.启用 = True
                objCard.是否刷卡 = True
                objCard.缺省标志 = Val(Nvl(!缺省)) = 1
                If InStr("|" & mstr脱机医保 & "|", "|" & !名称 & "|") > 0 And bln含脱机医保 = False Then
                    objCard.缺省标志 = False
                    Call objOfflineYBCards_Out.Add(objCard)
                Else
                    Call objPayCards_Out.Add(objCard)
                End If
               i = i + 1
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    If InStr(";" & mstrOnCardPrivs & ";", ";三方接口消费;") > 0 Then
       rsBalace_Out.Filter = "性质>=7 and 性质<9" '一卡通结算
        With rsBalace_Out
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For Each objCard In objCards
                    If objCard.结算方式 = Nvl(!名称) Then
                    
                        Set objTemp = objCard.CloneNewCard
                        objTemp.是否非接触式读卡 = False
                        objTemp.是否接触式读卡 = False
                        
                        objTemp.缺省标志 = Val(Nvl(!缺省)) = 1
                        
                        objPayCards_Out.Add objTemp
                        If objTemp.是否支持扫码付 And objTemp.消费卡 = False Then
                            strQRCodeTypeIDs_Out = strQRCodeTypeIDs_Out & "," & objTemp.接口序号
                        End If
                          i = i + 1
                        Exit For
                    End If
                Next
                .MoveNext
            Loop
            If strQRCodeTypeIDs_Out <> "" Then strQRCodeTypeIDs_Out = Mid(strQRCodeTypeIDs_Out, 2)
            .Filter = 0
        End With
    End If
    
    rsBalace_Out.Filter = "性质=9"
    If rsBalace_Out.EOF = False Then
        With rsBalace_Out
            objErrCard_Out.接口序号 = -1 * i
            objErrCard_Out.接口编码 = !编码
            objErrCard_Out.名称 = !名称
            objErrCard_Out.结算方式 = !名称
        End With
    Else
        objErrCard_Out.接口序号 = -1 * i
        objErrCard_Out.接口编码 = ""
        objErrCard_Out.名称 = "误差费"
        objErrCard_Out.结算方式 = "误差费"
        objErrCard_Out.结算性质 = 9
    End If
    objErrCard_Out.结算性质 = 9
    objErrCard_Out.启用 = True
    objErrCard_Out.是否刷卡 = True
    objErrCard_Out.缺省标志 = False
    rsBalace_Out.Filter = 0
    zlGetBalanceModeDatas = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function


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
 
Public Function zlInitCompents(ByVal frmMain As Object, ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化部件
    '入参:objICCard-IC卡对象
    '出参:
    '返回:初始化成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 17:00:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
      
    On Error GoTo errHandle
    
    mstrFunName = "zlInitCompents"
    Call WritLog(mstrFunName, "开始", lngModule)
    
    Set mfrmMain = frmMain: mlngModule = lngModule
    
    If Not mfrmMain Is Nothing Then
        mstrTittle = mfrmMain.Caption
    End If
    
    mbln结算异步控制 = Val(zlDatabase.GetPara("费用结算异步控制", glngSys)) = 1
    mbyt刷卡缺省金额操作 = Val(zlDatabase.GetPara("刷卡缺省金额操作", glngSys, 1151, "0"))    '86853
    
    zlInitCompents = True
    Call WritLog(mstrFunName, "结束", "返回=true", mbln结算异步控制, mbyt刷卡缺省金额操作, mstrTittle)
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "错误", "重试,错误信息=" & mstrErrInfor)
        Resume
    End If
    Call WritLog(mstrFunName, "结束", "返回=false", "错误信息=" & mstrErrInfor)
End Function

Public Function zlThird_IsCancelFromItems(ByVal objItems As clsBalanceItems) As Boolean
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
    Dim intSwapStaus As Integer, objCard As Card, objItem As clsBalanceItem
    Dim strErrMsg As String
    
    
    On Error GoTo errHandle
    
    mstrFunName = "zlThird_IsCancelFromItems": mstrPara = ""
    If objItems Is Nothing Then mstrPara = "objItems=nothing "
    
    Call WritLog(mstrFunName, "开始", mstrPara)

    If objItems.Count = 0 Then
        Call WritLog(mstrFunName, "结束", "返回=False,objItems.Count=0")
        Exit Function
    End If
    
    Set objItem = objItems(1): Set objCard = objItem.objCard
    
    If objCard.消费卡 Or objItem.结算ID = 0 Then
        mstrPara = "objCard.消费卡 =" & IIf(objCard.消费卡, "true", "false") & vbTab & "objItem.结算ID=" & objItem.结算ID & vbTab
        Call WritLog(mstrFunName, "结束", "返回=False", mstrPara)
        Exit Function
    End If

 
    If zlThird_IsSwapIsSucces(objItems, intSwapStaus, strErrMsg) Then
        MsgBox "注意:" & vbCrLf & _
               "    " & objCard.名称 & " 已经交易成功，不允许作废或删除，你必须完成结帐操作!", vbInformation + vbOKOnly, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 已经交易成功，不允许作废或删除，你必须完成结帐操作!")
        Exit Function
    End If
    If intSwapStaus <> 0 Then
        If strErrMsg <> "" Then
            MsgBox "注意:" & vbCrLf & _
                   "    " & objCard.名称 & " 交易正在进行中,不能删除或作废,详细错误信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
            
            Call WritLog(mstrFunName, "结束", "返回=False,", objCard.名称 & " 交易正在进行中,不能删除或作废,详细错误信息如下:" & vbCrLf & strErrMsg & "!")
        Else
            MsgBox "注意:" & vbCrLf & _
                   "    " & objCard.名称 & " 交易正在进行中，不允许作废或删除!", vbInformation + vbOKOnly, gstrSysName
            Call WritLog(mstrFunName, "结束", "返回=False,", objCard.名称 & " 交易正在进行中，不允许作废或删除!")
        End If
        Exit Function
    End If
    zlThird_IsCancelFromItems = True
    Call WritLog(mstrFunName, "结束", "返回=true")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WritLog(mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function

Public Function zlThird_IsSwapIsSucces(ByVal objItems As clsBalanceItems, ByRef intSwapStatu_Out As Integer, ByRef strErrmsg_Out As String, _
    Optional ByVal lng预交ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断交易是否调用成功
    '入参: objItems-当前结算信息集
    '      lng预交ID-发卡同时处理预交有效
    '出参:intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
    '     strErrMsg_out-返回的错误信息
    '返回：接口调用成功返回true,否则返回Flase
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intSwapType As Integer, dblMoney As Double, strXml As String, strXMLExpend As String
    Dim objCard As Card, objItem As clsBalanceItem, objTemp As clsBalanceItem
    Dim strBalanceIDs As String, intSign As Integer
    Dim cllBalance As Collection
    Dim i As Long
    
    On Error GoTo errHandle
    
    intSwapStatu_Out = 0: strErrmsg_Out = ""
    mstrFunName = "zlThird_IsSwapIsSucces"
    Call WritLog(mstrFunName, "开始", intSwapStatu_Out, lng预交ID)
    
    If objItems Is Nothing Then Call WritLog(mstrFunName, "结束", "返回=False", "objItems = Nothing"): Exit Function
    
    If objItems.Count = 0 Then
        Call WritLog(mstrFunName, "结束", "返回=False", "objItems.Count = 0")
        Exit Function
    End If
    
    dblMoney = 0
    Set cllBalance = New Collection
    For Each objItem In objItems
        strBalanceIDs = ""
        For i = 1 To cllBalance.Count
            If cllBalance(i)(0) = objItem.单据性质 Then
                strBalanceIDs = cllBalance(i)(1)
                cllBalance.Remove i: Exit For
            End If
        Next
        strBalanceIDs = strBalanceIDs & "," & IIf(objItem.冲销ID <> 0, objItem.冲销ID, objItem.结算ID)
        cllBalance.Add Array(objItem.单据性质, strBalanceIDs)
        dblMoney = dblMoney + objItem.结算金额
    Next
    strBalanceIDs = ""
    For i = 1 To cllBalance.Count
        strBalanceIDs = strBalanceIDs & "||" & cllBalance(i)(0) & "|" & Mid(cllBalance(i)(1), 2)
    Next
    If lng预交ID <> 0 Then strBalanceIDs = strBalanceIDs & "||2|" & lng预交ID
    If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 3)
    
    Set objItem = objItems(1)
    Set objCard = objItem.objCard
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If
    
    
    On Error GoTo errHandle
   '0-扣款;1-退款；2-转帐
    intSwapType = IIf(objItem.是否退款, IIf(objItem.是否转帐, 2, 1), 0)
    intSign = IIf(objItem.是否退款, -1, 1)
    dblMoney = RoundEx(intSign * dblMoney, 2)
       
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
    '    '日期:2013-06-15 20:22:51
    '    '说明:
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = "": strXml = ""
    For Each objTemp In objItems
        strXml = strXml & "        <TK>" & vbCrLf
        strXml = strXml & "             <JYLSH>" & objTemp.交易流水号 & "</JYLSH>" & vbCrLf
        strXml = strXml & "             <KH>" & objTemp.卡号 & "</KH>" & vbCrLf
        strXml = strXml & "             <JE>" & RoundEx(intSign * objTemp.结算金额, 6) & "</JE>" & vbCrLf
        strXml = strXml & "        </TK>" & vbCrLf
    Next
    
    strXml = "    <TKLIST>" & vbCrLf & strXml & vbCrLf
    strXml = strXml & "    </TKLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "<INPUT>" & vbCrLf
    strXMLExpend = strXMLExpend & strXml
    strXMLExpend = strXMLExpend & "</INPUT>" & vbCrLf
    
    
    Call WritLogCall(mstrFunName, "zlSwapIsSucces", EM_CallBegin, objItem.卡号, dblMoney, strBalanceIDs, strXMLExpend)
    
    If gobjOneCardComLib.zlSwapIsSucces(mfrmMain, mlngModule, intSwapType, objCard.接口序号, strBalanceIDs, objItem.卡号, dblMoney, intSwapStatu_Out, strErrmsg_Out, strXMLExpend) = False Then
         Call WritLogCall(mstrFunName, "zlSwapIsSucces", EM_CallEnd, "zlSwapIsSucces=False", intSwapStatu_Out, strErrmsg_Out, strXMLExpend)
         zlThird_IsSwapIsSucces = False
         Exit Function
    End If
    Call WritLogCall(mstrFunName, "zlSwapIsSucces", EM_CallEnd, "zlSwapIsSucces=true", intSwapStatu_Out, strErrmsg_Out, strXMLExpend)
    
    zlThird_IsSwapIsSucces = True
    Call WritLog(mstrFunName, "结束", "返回=true")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WritLog(mstrFunName, "结束", "返回=False", "Err=" & Err.Description)
End Function


Public Function zlThird_Payment(ByVal objCard As Card, ByVal objPati As clsPatientInfo, ByRef cllBillPro As Collection, _
    ByVal objBalanceItems As clsBalanceItems, ByRef objBalanceItems_Out As clsBalanceItems, _
    ByRef rsSwapExpend_out As ADODB.Recordset, Optional blnSaveed_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行三方接口支付
    '入参:
    '    objCard-当前支付的卡对象
    '    objPati-病人对象集
    '    cllBillPro-执行的相关过程
    '    objBalanceItems-结算数据集
    '出参:objBalanceItems_Out-结帐返回数据(可能支付后结算信息发生了相关变化)
    '     rsSwapExpend_out-三方交易返回的扩展信息
    '     blnSaveed_Out-是否发生数据保存
    '返回:三方接口交易调用成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 17:02:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结帐IDs As String, str卡号 As String, strSQL As String, strErrmsg_Out As String
    Dim strSwapGlideNO As String, strSwapMemo As String, strExpendXML As String, str结算方式 As String
    Dim lng预交ID As Long, lngCount As Long, lng关联交易ID As Long, lng结帐ID As Long, i As Long
    Dim cllPro As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim rsBalance As ADODB.Recordset
    Dim intStatus_Out As Integer
    Dim dblMoney As Double
    Dim strQRCode As String
    Dim blnExcuteThird As Boolean, blnTrans As Boolean
    Dim cllSaveErr As Collection
    Dim strJson As String, strJsonExpend As String
    
    
    On Error GoTo errHandle
    
    mstrFunName = "zlThird_Payment"
    Call WritLog(mstrFunName, "开始", "objCard=" & IIf(objCard Is Nothing, "Nothing", "Not Nothing"))
     
    blnSaveed_Out = False
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        zlThird_Payment = True
        mstrPara = "结算方式=" & objCard.结算方式 & vbTab & " objCard.接口序号 =" & objCard.接口序号 & vbTab & "objCard.消费卡=" & IIf(objCard.消费卡, "true", "false") & vbTab
        Call WritLog(mstrFunName, "结束", "返回=true", mstrPara)
        Exit Function
    End If
    
    If objBalanceItems Is Nothing Then
        Call WritLog(mstrFunName, "结束", "返回=False", "objBalanceItems Is Nothing")
        Exit Function
    End If
    If objBalanceItems.Count = 0 Then
        Call WritLog(mstrFunName, "结束", "返回=False", "objBalanceItems.Count = 0")
        Exit Function
    End If
    
    Set objItem = objBalanceItems(1)
    str卡号 = objItem.卡号
    strQRCode = objItem.QRCode  '二维码内码
    
    dblMoney = objItem.结算金额
    If dblMoney = 0 Then
        Call WritLog(mstrFunName, "结束", "返回=False", "objItem.结算金额 = 0")
        Exit Function
    End If
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    '产生关联数据
    '更新同步标志
    Set cllSaveErr = New Collection
    '入参:cllSaveData-格式为Array(保存项名称,保存项值)
    '                   保存的项目名称包含:操作状态,异常ID,操作场景,作废标志,业务id,是否病历费,病人id,主页id,姓名,性别,年龄,门诊号,住院号,预交单号,预交金额,医疗卡单号,卡费,发卡类别id,发卡类别名称,发卡卡号,同步状态,交易信息)
    '                   其中交易信息为Json串，格式如下
    '                   {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]}
    cllSaveErr.Add Array("操作状态", 1) '操作状态:0-新增记录,1-更新状态及更新交易说明
    cllSaveErr.Add Array("异常ID", objItem.异常ID)
    cllSaveErr.Add Array("同步状态", IIf(mbln结算异步控制, 1, 2)) '同步状态:0或NULL正常记录;1-未调用调接口;2-结算接口调用成功
    
    strJson = GetJsonNodeString("card_no", str卡号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cardtype_id", objItem.卡类别ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("swapno", objItem.交易流水号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("swapmemo", objItem.交易说明, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("swapmoney", dblMoney, Json_num)
    strJson = "{" & strJson & "}"
    cllSaveErr.Add Array("交易信息", strJson)
ReDO:

    blnTrans = True: blnExcuteThird = False
    
    '1.批量执行过程
    zlExecuteProcedureArrAy cllPro, mstrTittle & "-执行三方接口", True, False, "发卡", "病人发卡调试日志"
    '2.更新异步标志
    If Zl_病人结算异常记录_Modify(1, cllSaveErr) = False Then
        blnTrans = False
        gcnOracle.RollbackTrans:
        Call WritLog(mstrFunName, "结束", "返回=False", IIf(mbln结算异步控制, 1, 2), objItem.异常ID, "1")
        Exit Function
    End If
    
    If mbln结算异步控制 Then
        gcnOracle.CommitTrans
        blnTrans = False: blnSaveed_Out = True
    End If
    
    str结帐IDs = objItem.结算ID
    blnExcuteThird = True
    
    'Public Function zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, ByVal strPrepayNos As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strExpendXML As String, _
    Optional ByVal blnResolveXMLToRecord As Boolean, _
    Optional ByRef rsBalance_Out As ADODB.Recordset, Optional rsSwapExpend_out_Out As ADODB.Recordset, _
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
    '    blnResolveXMLToRecord-是否解析XML串给记录订(rsBalance_Out,rsSwapExpend_out_Out
    '   strExpendXML，分两种情况或两种格式
    '       1.验证密码: 自助机用
    '           <Data>
    '               <验证密码></验证密码>
    '           </Data>
    '       2.消费卡收款及扫码付时，传入XML串：
    '       <IN>
    '           <MZXSJE>面值销售金额</MZXSJE>
    '           <CZJKJE>充值缴款金额</CZJKJE>
    '           <QRCODE>识别的二维码代码串</QRCODE> //10.35.130后支持
    '       </IN>
    '出参:strSwapGlideNO-交易流水号(XML中返回结算信息时,交易流水号，以xml为准)
    '    strSwapMemo-交易说明(XML中返回返回结算信息时,交易以xml为准)
    '    rsBalance_Out-结算信息集(blnResolveXMLToRecord=true时返回)字段（卡类别ID,消费卡,结算方式,交易金额,交易流水号,交易说明,单据号,单据性质,卡号,结算号码,结算摘要,是否普通结算)
    '    rsSwapExpend_out_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
    '    intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
    '    strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
    '    strExpendXML -交易扩展信息,有两种方式:
    
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
    strExpendXML = ""
    If strQRCode <> "" Then
        strExpendXML = "" & _
        "<IN>" & vbCrLf & _
        "    <QRCODE>" & strQRCode & "</QRCODE>" & vbCrLf & _
        "</IN>" & vbCrLf
    End If
    
    
    Call WritLogCall(mstrFunName, "zlPaymentMoney", EM_CallBegin, objCard.接口序号, objCard.消费卡, objItem.卡号, str结帐IDs, dblMoney, strSwapGlideNO, strSwapMemo, strExpendXML, True, intStatus_Out, strErrmsg_Out)
    
    If gobjOneCardComLib.zlPaymentMoney(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, objItem.卡号, _
         str结帐IDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strExpendXML, True, rsBalance, rsSwapExpend_out, intStatus_Out, strErrmsg_Out) = False Then
        blnExcuteThird = False
        If blnTrans Then gcnOracle.RollbackTrans
        blnTrans = False
        'intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
        'strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
        If strErrmsg_Out <> "" Then
            MsgBox "交易正在进行中，原因如下:" & vbCrLf & strErrmsg_Out, vbOKOnly + vbOKOnly + vbInformation, gstrSysName
        End If
        
        If blnSaveed_Out Then Set cllBillPro = New Collection
        If intStatus_Out = 1 Then   '1-交易正在处理中
            If blnSaveed_Out Then
                For Each objItem In objBalanceItems
                    objItem.是否保存 = True
                    objItem.是否允许编辑 = False
                    objItem.是否允许删除 = True
                    objItem.QRCode = ""
                Next
                Set objBalanceItems_Out = objBalanceItems
            End If
            Call WritLogCall(mstrFunName, "zlPaymentMoney", EM_CallEnd, "返回=False", intStatus_Out, strErrmsg_Out, strExpendXML)
            Call WritLog(mstrFunName, "结束", "返回=False", "zlPaymentMoney=False")
            Exit Function
        End If
         
        objBalanceItems(1).结算ID = objItem.结算ID
        If mbln结算异步控制 Then
            '需要删除
           ' Call zlDelBalanceItemFromItemObject(objBalanceItems(1))
        End If
        Set objBalanceItems_Out = New clsBalanceItems
        
        Call WritLogCall(mstrFunName, "zlPaymentMoney", EM_CallEnd, "返回=False", " blnSaveed_Out=" & IIf(blnSaveed_Out, "true", "false"), intStatus_Out, strErrmsg_Out, strExpendXML)
        Call WritLog(mstrFunName, "结束", "返回=False", "zlPaymentMoney=False")
        Exit Function
    End If
    Call WritLogCall(mstrFunName, "zlPaymentMoney", EM_CallEnd, "zlPaymentMoney=true", strSwapGlideNO, strSwapMemo, intStatus_Out, strErrmsg_Out, strExpendXML)
    
    blnExcuteThird = False
    Set objBalanceItems_Out = New clsBalanceItems
    
    '更新数据
    If Not rsBalance Is Nothing Then
        With rsBalance
            .Filter = 0
            If .RecordCount <> 0 Then .MoveFirst: dblMoney = 0
            Do While Not .EOF
                Set objItemTemp = objItem.Clone
                With objItemTemp
                    .卡类别ID = IIf(Val(Nvl(rsBalance!是否普通结算)) = 0, Val(Nvl(rsBalance!卡类别ID)), 0)
                    .消费卡 = False
                    .消费卡ID = 0
                    .卡号 = IIf(Nvl(rsBalance!卡号) = "", objItem.卡号, Nvl(rsBalance!卡号))
                    If .卡号 <> "" Then str卡号 = .卡号
                    .交易流水号 = Nvl(rsBalance!交易流水号)
                    .交易说明 = Nvl(rsBalance!交易说明)
                    .结算方式 = Nvl(rsBalance!结算方式)
                    .结算金额 = Val(Nvl(rsBalance!交易金额))
                    .帐户余额 = objItem.帐户余额
                    .结算号码 = Nvl(rsBalance!结算号码)
                    .结算摘要 = Nvl(rsBalance!结算摘要)
                    .结算时间 = objItem.结算时间
                    .是否允许编辑 = False
                    .是否允许删除 = False
                    .是否保存 = True
                    .是否结算 = True
                    .是否结算 = True
                    .结算类型 = 3   ''0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                    .校对标志 = 2
                End With
                objBalanceItems_Out.AddItem objItemTemp
                dblMoney = RoundEx(dblMoney + objItemTemp.结算金额, 6)
                lng预交ID = 0
                lngCount = lngCount + 1
                .MoveNext
            Loop
            objBalanceItems_Out.结算金额 = dblMoney
        End With
    Else
        Set objItemTemp = objItem.Clone
        With objItemTemp
            .交易流水号 = strSwapGlideNO
            .交易说明 = strSwapMemo
            .消费卡 = False
            .消费卡ID = 0
            .是否结算 = True
            .结算类型 = 3
            .是否结算 = True
            .是否允许编辑 = False
            .是否允许删除 = False
            .是否保存 = True
            .是否结算 = True
            .校对标志 = 2
        End With
        objBalanceItems_Out.AddItem objItemTemp
        objBalanceItems_Out.结算金额 = objItem.结算金额
    End If
   
    '异常数据更新
    Set cllSaveErr = New Collection
    '入参:cllSaveData-格式为Array(保存项名称,保存项值)
    '                   保存的项目名称包含:操作状态,异常ID,操作场景,作废标志,业务id,是否病历费,病人id,主页id,姓名,性别,年龄,门诊号,住院号,预交单号,预交金额,医疗卡单号,卡费,发卡类别id,发卡类别名称,发卡卡号,同步状态,交易信息)
    '                   其中交易信息为Json串，格式如下
    '                   {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]}
    cllSaveErr.Add Array("操作状态", 1) '操作状态:0-新增记录,1-更新状态及更新交易说明
    cllSaveErr.Add Array("异常ID", objItem.异常ID)
    cllSaveErr.Add Array("同步状态", 2) '同步状态:0或NULL正常记录;1-未调用调接口;2-结算接口调用成功
    
    Set objItemTemp = objBalanceItems_Out(1)
    strJson = GetJsonNodeString("card_no", str卡号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cardtype_id", objItemTemp.卡类别ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("swapno", objItemTemp.交易流水号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("swapmemo", objItemTemp.交易说明, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("swapmoney", objBalanceItems_Out.结算金额, Json_num)
    strJson = strJson & "," & GetJsonNodeString("qrcode", strQRCode, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("blnc_mode", objItemTemp.结算方式, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("blnc_no", objItemTemp.结算号码, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("blnc_memo", objItemTemp.结算摘要, Json_Text)
    
    Call zlGetThreeSwapExpendJsonStringByRecords(rsSwapExpend_out, strJsonExpend)
    
    If strJsonExpend <> "" Then
        strJson = strJson & "," & strJsonExpend
    End If
    strJson = "{" & strJson & "}"
    
    cllSaveErr.Add Array("交易信息", strJson)
    
    Call Zl_病人结算异常记录_Modify(1, cllSaveErr)
    If Not mbln结算异步控制 Then gcnOracle.CommitTrans
    
    blnSaveed_Out = True: blnTrans = False: Set cllBillPro = New Collection
    zlThird_Payment = True
    
    Call WritLog(mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    
    If blnTrans Then gcnOracle.RollbackTrans
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        If blnExcuteThird Then GoTo ReDO:
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call SaveErrLog
    Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=" & mstrErrInfor)
    Exit Function
ErrOther:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,ErrOther", "错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
    Call WritLog(mstrFunName, "结束", "返回=False,ErrOther", "错误信息=" & mstrErrInfor)
End Function


Public Function zlThird_ReturnCashCheck(ByVal objCard As Card, ByRef objBalanceItems As clsBalanceItems, ByRef blnDefaultCash_Out As Boolean, ByRef strDefaultBalance_Out As String) As Boolean
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
    Dim str预交IDs As String
    Dim i As Long
    
    
    Err = 0: On Error GoTo errHandle

    mstrFunName = "zlThird_ReturnCashCheck"
    Call WritLog(mstrFunName, "开始")

    If objBalanceItems Is Nothing Then
        Call WritLog(mstrFunName, "结束", "返回=False", "objBalanceItems Is Nothing")
        Exit Function
    End If
    If objBalanceItems.Count = 0 Then
        Call WritLog(mstrFunName, "结束", "返回=False", "objBalanceItems.Count = 0")
        Exit Function
    End If

    strBalanceIDs = "": str预交IDs = ""
    strIDs = ""
    For i = 1 To objBalanceItems.Count
        Set objItemTemp = objBalanceItems(i)
        If objItemTemp.卡号 <> "" And objItem Is Nothing Then Set objItem = objItemTemp
        If objItemTemp.是否预交 Then
            str预交IDs = str预交IDs & "," & IIf(objItemTemp.预交ID = 0, objItemTemp.结算ID, objItemTemp.预交ID)
        Else
             strIDs = strIDs & "," & objItemTemp.结算ID
        End If
        strTKList = ""
        strTKList = strTKList & Space(8) & "<TK>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <TKFS>" & objItemTemp.结算方式 & "</TKFS>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <TKJE>" & RoundEx(objItemTemp.结算金额, 6) & "</TKJE>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <JYLSH>" & objItemTemp.交易流水号 & "</JYLSH>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <JYSM>" & objItemTemp.交易说明 & "</JYSM>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <KH>" & objItemTemp.卡号 & "</KH>" & vbCrLf
        strTKList = strTKList & Space(8) & "</TK>" & vbCrLf
        strXMLExpend = strXMLExpend & strTKList
        dblMoney = dblMoney + -1 * objItemTemp.结算金额
    Next

    If strIDs <> "" Then strIDs = "2|" & Mid(strIDs, 2)
    If str预交IDs <> "" Then str预交IDs = "1|" & Mid(str预交IDs, 2)

    strBalanceIDs = strIDs & IIf(strIDs <> "" And str预交IDs <> "", "||", "") & str预交IDs
    strTKList = ""
    strTKList = strTKList & "<INPUT>" & vbCrLf
    strTKList = strTKList & "    <TKLIST>" & vbCrLf
    strTKList = strTKList & strXMLExpend
    strTKList = strTKList & "    </TKLIST>" & vbCrLf
    strTKList = strTKList & "</INPUT>"

    strXMLExpend = strTKList
    If objItem Is Nothing Then Set objItem = objBalanceItems(1)

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
    '    '       strXMLExpend    XML IN  10.35.110后才支持
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
    '    '       strXMLExpend:10.35.110后才支持
    '    '        <OUTPUT>
    '    '            <SFQSTX>是否缺省退现<SFQSTX>//NUMBER 1 是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但以许操作员操作退现
    '    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
    '    '                    1.不允许返回三卡方的结算方式
    '    '                    2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
    '    '        </OUTPUT>
    '    '返回:退现合法,返回true,否则返回Flase

    Call WritLogCall(mstrFunName, "zlReturnCashCheck", EM_CallBegin, objCard.接口序号, objItem.卡号, strBalanceIDs, dblMoney, objItem.交易流水号, objItem.交易说明, strXMLExpend)
    If gobjOneCardComLib.zlReturnCashCheck(mfrmMain, mlngModule, objCard.接口序号, objItem.卡号, _
          strBalanceIDs, dblMoney, objItem.交易流水号, objItem.交易说明, strXMLExpend, blnDefaultCash_Out, strDefaultBalance_Out) = False Then
          
        Call WritLogCall(mstrFunName, "zlReturnCashCheck", EM_CallEnd, "zlReturnCashCheck=false", strXMLExpend, blnDefaultCash_Out, strDefaultBalance_Out)
        Call WritLog(mstrFunName, "结束", "返回=False", "zlReturnCashCheck=false")
        zlThird_ReturnCashCheck = False
        Exit Function
    End If
    Call WritLogCall(mstrFunName, "zlReturnCashCheck", EM_CallEnd, "zlReturnCashCheck=true", strXMLExpend, blnDefaultCash_Out, strDefaultBalance_Out)
    
    zlThird_ReturnCashCheck = True
    Call WritLog(mstrFunName, "结束", "返回=true")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WritLog(mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function

Public Function zlThird_ReturnMoney_IsValied(ByVal objPati As clsPatientInfo, ByVal objCard As Card, ByVal byt操作类型 As Byte, _
    ByVal objDelBalanceItems As clsBalanceItems, ByRef objItems_Out As clsBalanceItems, _
    Optional blnCurInput As Boolean, Optional blnCallSwapIsSucces As Boolean) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方卡退款合法性检查
    '入参:objPati-病人信息对象
    '     objCard-卡对象
    '     objDelBalanceItems-退费结算信息
    '     byt操作类型-0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
    '     objBalanceInfor-当前结帐信息
    '     blnCurInput-是否当前输入的操作
    '     blnCallSwapIsSucces-是否调用了成功交易判断接口
    '出参:objItems_Out-当前返顺的有效结算信息
    '返回:允许退款返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-28 18:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend  As String, str交易说明 As String, str交易流水号 As String, strBalanceIDs As String, str卡号 As String, str密码 As String
    Dim bln退现 As Boolean, intStatus As Integer, strErrMsg As String
    Dim dblMoney  As Double, dbl原始金额 As Double, dbl未退金额 As Double, rsSwap As ADODB.Recordset
    Dim str退费结算IDs As String, str预交退费结算IDs As String, str关联交易Ids As String
    Dim strXMLIn As String
    Dim objItem As clsBalanceItem
    Dim i As Long, blnCheck As Boolean
 

    On Error GoTo errHandle

    mstrFunName = "zlThird_ReturnMoney_IsValied"
    Call WritLog(mstrFunName, "开始", byt操作类型, blnCurInput, blnCallSwapIsSucces)

    If objCard Is Nothing Then
        Screen.MousePointer = 0
        If InStr(";" & mstrOnCardPrivs & ";", ";三方接口消费;") = 0 Then
            MsgBox "你没有三方接口消费权限，无法调用接口部件！", vbInformation, gstrSysName
            
            Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=你没有三方接口消费权限，无法调用接口部件！")
            Exit Function
        End If
        MsgBox "未找到退款接口,请检查接口部件！", vbInformation, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=未找到退款接口,请检查接口部件")
        Exit Function
    End If


    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=" & objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If


    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        zlThird_ReturnMoney_IsValied = True
        mstrPara = "objCard.接口序号=" & objCard.接口序号 & vbTab & "objCard.消费卡=" & IIf(objCard.消费卡, "true", "false") & vbTab & "objCard.结算方式=" & objCard.结算方式
        Call WritLog(mstrFunName, "结束", "返回=true", objCard.接口序号, "objCard.消费卡=" & IIf(objCard.消费卡, "true", "false"), objCard.结算方式)
        Exit Function
    End If

    If objDelBalanceItems Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=当前退款结算信息不存在，请检查!")
        Exit Function
    End If

    If objDelBalanceItems.Count = 0 Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=当前退款结算信息不存在，请检查（objDelBalanceItems.Count＝0）")
        Exit Function
    End If

   '先检查数据的合法性
    blnCheck = True
    If objDelBalanceItems(1).是否保存 And Not blnCallSwapIsSucces Then
        If Not zlThird_IsSwapIsSucces(objDelBalanceItems, intStatus, strErrMsg) Then
            If intStatus = 1 Then
                 If strErrMsg <> "" Then
                     MsgBox "注意:" & vbCrLf & _
                     "   " & objCard.名称 & " 正在进行扣款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName

                     Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=" & objCard.名称 & " 正在进行扣款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg)
                 Else
                     MsgBox "注意:" & vbCrLf & _
                     "   " & objCard.名称 & " 正在进行扣款处理中，请稍后再试", vbInformation + vbOKOnly, gstrSysName

                     Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=" & objCard.名称 & " 正在进行扣款处理中，请稍后再试")
                 End If
                 Exit Function
            End If
        Else
            Set objItems_Out = objDelBalanceItems
            zlThird_ReturnMoney_IsValied = True: blnCheck = False:
            Call WritLog(mstrFunName, "结束", "返回True", "调用zlThird_IsSwapIsSucces=true")
            Exit Function
        End If
    End If



    dblMoney = objDelBalanceItems.结算金额
    dbl未退金额 = objDelBalanceItems.未退金额

    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox "退款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=退款金额未输入,请检查!")
        Exit Function
    End If
    
    If RoundEx(Abs(dblMoney), 2) > RoundEx(Abs(dbl未退金额), 2) And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！", vbInformation, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！")
        Exit Function
    End If


    If Abs(dblMoney) <> Abs(dbl未退金额) And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1 '提示
            Screen.MousePointer = 0
            If MsgBox("当前退款金额与未退金额不符,是否继续？" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(Abs(dblMoney), "0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then


                Call WritLog(mstrFunName, "结束", "返回=False", "错误信息(选择否)=" & _
                "当前退款金额与未退金额不符,是否继续？" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(Abs(dblMoney), "0.00"))
                Exit Function
            End If
        Case 2 '禁止
            Screen.MousePointer = 0
            Call MsgBox("当前退款金额与未退金额不符,不能继续!" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(Abs(dblMoney), "0.00"), vbInformation + vbOKOnly, gstrSysName)

                Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=" & _
                "当前退款金额与未退金额不符,不能继续!" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(Abs(dblMoney), "0.00"))
            Exit Function
        End Select
    End If

    str关联交易Ids = ""
    For Each objItem In objDelBalanceItems
        If InStr(str关联交易Ids & ",", "," & objItem.关联交易ID & ",") = 0 And objItem.关联交易ID <> 0 Then
            str关联交易Ids = str关联交易Ids & "," & objItem.关联交易ID
        End If
    Next

    If str关联交易Ids <> "" Then
        str关联交易Ids = Mid(str关联交易Ids, 2)
        If zlGetThirdMoneyInforRecordFromSwapID(str关联交易Ids, rsSwap) = False Then
            str关联交易Ids = ""
        ElseIf rsSwap Is Nothing Then
             str关联交易Ids = ""
        ElseIf rsSwap.State <> 1 Then
            str关联交易Ids = ""
        ElseIf rsSwap.RecordCount = 0 Then
            str关联交易Ids = ""
        End If
    End If

    If str关联交易Ids = "" Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & "  未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!", vbInformation + vbOKOnly, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "无关联交易ID")
        Exit Function
    End If

    str关联交易Ids = ""
    dbl原始金额 = 0: dbl未退金额 = 0
    For Each objItem In objDelBalanceItems

        If InStr(str关联交易Ids & ",", "," & objItem.关联交易ID & ",") = 0 And objItem.关联交易ID <> 0 Then
            str关联交易Ids = str关联交易Ids & "," & objItem.关联交易ID
            rsSwap.Filter = "关联交易ID=" & objItem.关联交易ID
            Do While Not rsSwap.EOF
                dbl未退金额 = RoundEx(dbl未退金额 + Val(Nvl(rsSwap!剩余未退金额)), 6)
                dbl原始金额 = RoundEx(dbl原始金额 + Val(Nvl(rsSwap!原始金额)), 6)
                rsSwap.MoveNext
            Loop

            If objCard.是否全退 Then
                If Abs(dblMoney) <> dbl原始金额 Then
                    Screen.MousePointer = 0
                    Call MsgBox(objCard.名称 & "必须全退,不能继续!" & vbCrLf & _
                         "原始金额:" & Format(dbl原始金额, "0.00") & vbCrLf & _
                         "当前退款:" & Format(Abs(dblMoney), "0.00") & vbCrLf, vbInformation + vbOKOnly, gstrSysName)

                    Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=" & objCard.名称 & "必须全退,不能继续!" & vbCrLf & _
                         "原始金额:" & Format(dbl原始金额, "0.00") & vbCrLf & _
                         "当前退款:" & Format(Abs(dblMoney), "0.00") & vbCrLf)
                         
                         
                    Exit Function
                End If
            End If
            If Abs(dblMoney) > dbl未退金额 Then
                Screen.MousePointer = 0
                Call MsgBox("当前退款金额大于了剩余未退金额,不能继续!" & vbCrLf & _
                    "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & vbCrLf & _
                    "当前退款:" & Format(Abs(dblMoney), "0.00") & vbCrLf, vbInformation + vbOKOnly, gstrSysName)

                Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=当前退款金额大于了剩余未退金额,不能继续!" & vbCrLf & _
                    "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & vbCrLf & _
                    "当前退款:" & Format(Abs(dblMoney), "0.00") & vbCrLf)
                Exit Function
            End If
        End If
        If InStr(str关联交易Ids & ",", "," & objItem.关联交易ID & ",") = 0 And objItem.关联交易ID <> 0 Then
            str关联交易Ids = str关联交易Ids & "," & objItem.关联交易ID
        End If
    Next



    Set objItem = objDelBalanceItems(1)


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
    dblMoney = 0
    strXMLExpend = ""
    str退费结算IDs = ""
    str预交退费结算IDs = ""
    For i = 1 To objDelBalanceItems.Count
        dblMoney = RoundEx(dblMoney + objDelBalanceItems(i).结算金额, 6)
        If objDelBalanceItems(i).是否预交 Then
            If InStr(str预交退费结算IDs & ",", "," & objDelBalanceItems(i).预交ID & ",") = 0 Then str预交退费结算IDs = str预交退费结算IDs & "," & objDelBalanceItems(i).预交ID
        Else
           If InStr(str退费结算IDs & ",", "," & objDelBalanceItems(i).结算ID & ",") = 0 Then str退费结算IDs = str退费结算IDs & "," & objDelBalanceItems(i).结算ID
        End If
        strXMLExpend = strXMLExpend & vbCrLf & _
        "              <TK>" & vbCrLf & _
        "                   <TKFS>" & objDelBalanceItems(i).结算方式 & "</TKFS>" & vbCrLf & _
        "                   <TKJE>" & RoundEx(objDelBalanceItems(i).结算金额, 6) & "</TKJE>" & vbCrLf & _
        "                   <JYLSH>" & objDelBalanceItems(i).交易流水号 & "</JYLSH>" & vbCrLf & _
        "                   <JYSM>" & objDelBalanceItems(i).交易说明 & "</JYSM>" & vbCrLf & _
        "                   <DJH>" & "" & "</DJH>" & vbCrLf & _
        "              </TK>" & vbCrLf

    Next
    strBalanceIDs = ""
    If str退费结算IDs <> "" Then str退费结算IDs = Mid(str退费结算IDs, 2): strBalanceIDs = strBalanceIDs & "||5|" & str退费结算IDs
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

    If gobjOneCardComLib.zlReturncheck(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, objItem.卡号, _
        strBalanceIDs, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then
        Call WritLog(mstrFunName, "结束", "返回=False", "调用三方接口失败(zlReturnCheck=false)")
        Exit Function
    End If


    If Not objCard.是否退款验卡 Then
        Set objItems_Out = objDelBalanceItems: zlThird_ReturnMoney_IsValied = True:
        Call WritLog(mstrFunName, "结束", "返回=True", "退款验卡=false")
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
    If gobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, True, True, bln退现, True, Nothing, False, False, strXMLIn) = False Then

        Call WritLog(mstrFunName, "结束", "返回=False", "调用三方接口失败(zlBrushCard=false)")
        Exit Function
    End If

    Set objItems_Out = objDelBalanceItems:
    zlThird_ReturnMoney_IsValied = True
    Call WritLog(mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WritLog(mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function
 

Public Function zlThird_ReturnMoney(ByVal objPati As clsPatientInfo, objCurDelBalanceItems As clsBalanceItems, _
    ByRef cllBillPro As Collection, Optional ByVal bln异常作废 As Boolean, Optional objBalanceItems_Out As clsBalanceItems, _
    Optional blnSaveed_Out As Boolean, Optional blnCurInput As Boolean, Optional blnChangeMoney_Out As Boolean, _
    Optional blnOlnyModifyData As Boolean, Optional ByVal blnCallCheckedIterface As Boolean = True, _
    Optional ByRef rsExpend_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通支付(三方接口)
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     objBalanceInfor-原结算方式信息集
    '     objcurDelBalanceItems-当前退费的结算项(退款金额项，以负数体现)
    '     bln异常作废-异常作废时调用:true-异常作废;false-正常作废操作
    '     blnCurInput-是否当前输入项
    '     blnOlnyModifyData- 仅修正数据
    '     blnCallCheckedIterface-是否调用的回退检查接口:true-表示已经调用,False-表示未调用，本接口内部将调用
    '出参:
    '     objBalanceItems_Out-返回已经退费成功的项目(退款金额项，以负数体现)
    '     rsExpend_Out-扩展信息
    '     blnSaveed_Out-是否发生数据保存
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, cllBeforePro As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim objCurItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim rsBalance As ADODB.Recordset, rsExpend As ADODB.Recordset
    Dim blnTrans As Boolean, blnCallInterfance As Boolean
    Dim intStatus As Integer, strErrMsg As String, strBalanceIDs As String
    Dim strSwapNo As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strCardNo As String, dblMoney As Double, dblDelMoney As Double, str结算方式  As String
    Dim i As Long, lngCount As Long, lng结帐ID As Long
    Dim blnCheck As Boolean, str关联交易ID As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strExpendXML As String, strTemp As String
    Dim objItems As clsBalanceItems
    Dim mstrPara As String
    
    Dim objCard As Card, cllSaveErr As Collection
    Dim strJson As String, strJsonExpend As String
    
    
    On Error GoTo Errhand
    
    mstrFunName = "zlThird_ReturnMoney"
    Call WritLog(mstrFunName, "开始", bln异常作废)

    
    If objCurDelBalanceItems Is Nothing Then
        Call WritLog(mstrFunName, "结束", "返回=False", "objCurDelBalanceItems Is Nothing ")
        Exit Function
    End If
    If objCurDelBalanceItems.Count = 0 Then
        Call WritLog(mstrFunName, "结束", "返回=False", " objCurDelBalanceItems.Count = 0")
        Exit Function
    End If
    
    Set objCurItem = objCurDelBalanceItems(1)
    
    Set objCard = objCurItem.objCard
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    Set objBalanceItems_Out = Nothing

    blnSaveed_Out = False
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        zlThird_ReturnMoney = True
        Call WritLog(mstrFunName, "结束", "返回=true", "objCard.接口序号=" & objCard.接口序号, "objCard.消费卡=" & IIf(objCard.消费卡, "true", "false"))
        Exit Function
    End If


    '先检查数据的合法性
    If Not blnCallCheckedIterface Then

        blnCheck = True
        If objCurItem.是否保存 Then
            If Not zlThird_IsSwapIsSucces(objCurDelBalanceItems, intStatus, strErrMsg) Then
               If intStatus = 1 Then
                    If strErrMsg <> "" Then
                        MsgBox "注意:" & vbCrLf & _
                        "   " & objCard.名称 & " 正在进行退款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Else
                        MsgBox "注意:" & vbCrLf & _
                        "   " & objCard.名称 & " 正在进行退款处理中，请稍后再试" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    End If
                    Call WritLog(mstrFunName, "结束", "返回=False", "错误信息= " & objCard.名称 & " 正在进行退款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg)
                    Exit Function
               End If
            Else
                blnCheck = False
            End If
        End If
        If blnCheck Then
            ' 0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡4
            If zlThird_ReturnMoney_IsValied(objPati, objCard, 2, objCurDelBalanceItems, objItems, blnCurInput, True) = False Then
                Call WritLog(mstrFunName, "结束", "返回=False", "zlThird_ReturnMoney_IsValied=false")
                Exit Function
            End If
        End If
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next

    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)

    '处理退费列表数据
    i = 0
    
    lng结帐ID = IIf(objCurItem.冲销ID <> 0, objCurItem.冲销ID, objCurItem.结算ID)
    strExpendXML = "": strTemp = ""
    For Each objItemTemp In objCurDelBalanceItems

        strTemp = strTemp & "        <TK>" & vbCrLf
        strTemp = strTemp & "           <TKFS>" & objItemTemp.结算方式 & "</TKFS>" & vbCrLf
        strTemp = strTemp & "           <TKJE>" & RoundEx(objItemTemp.结算金额, 6) & "</TKJE>" & vbCrLf
        strTemp = strTemp & "           <JYLSH>" & objItemTemp.交易流水号 & "</JYLSH>" & vbCrLf
        strTemp = strTemp & "           <JYSM>" & objItemTemp.交易说明 & "</JYSM>" & vbCrLf
        strTemp = strTemp & "           <DJH>" & "" & "</DJH>" & vbCrLf
        strTemp = strTemp & "        </TK>"
        dblMoney = dblMoney + objItemTemp.结算金额
        i = i + 1
    Next
    
    
    '产生关联数据
    '更新同步标志
    Set cllSaveErr = New Collection
    '入参:cllSaveData-格式为Array(保存项名称,保存项值)
    '                   保存的项目名称包含:操作状态,异常ID,操作场景,作废标志,业务id,是否病历费,病人id,主页id,姓名,性别,年龄,门诊号,住院号,预交单号,预交金额,医疗卡单号,卡费,发卡类别id,发卡类别名称,发卡卡号,同步状态,交易信息)
    '                   其中交易信息为Json串，格式如下
    '                   {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]}
    cllSaveErr.Add Array("操作状态", 1) '操作状态:0-新增记录,1-更新状态及更新交易说明
    cllSaveErr.Add Array("异常ID", objCurItem.异常ID)
    cllSaveErr.Add Array("同步状态", IIf(mbln结算异步控制, 1, 2)) '同步状态:0或NULL正常记录;1-未调用调接口;2-结算接口调用成功
    
    strJson = GetJsonNodeString("card_no", objCurItem.卡号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cardtype_id", objCurItem.卡类别ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("swapno", objCurItem.交易流水号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("swapmemo", objCurItem.交易说明, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("swapmoney", dblMoney, Json_num)
    strJson = "{" & strJson & "}"
    cllSaveErr.Add Array("交易信息", strJson)

    dblDelMoney = dblMoney
    strExpendXML = strExpendXML & "<INPUT>" & vbCrLf
    strExpendXML = strExpendXML & "    <TKLIST>" & vbCrLf
    strExpendXML = strExpendXML & strTemp & vbCrLf
    strExpendXML = strExpendXML & "    </TKLIST>" & vbCrLf
    strExpendXML = strExpendXML & "</INPUT>" & vbCrLf

    If dblMoney = 0 Then
        MsgBox "当前退款为零了，不允许再次退款", vbInformation + vbOKOnly, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=当前退款为零了，不允许再次退款")
        Exit Function
    End If

 
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, mstrTittle, True, , "退费操作", G_STR_LOG_NAME

    '2.更新异步标志
    If Zl_病人结算异常记录_Modify(1, cllSaveErr) = False Then
        blnTrans = False
        gcnOracle.RollbackTrans
        Call WritLog(mstrFunName, "结束", "返回=False", "更新异常标志失败", IIf(mbln结算异步控制, 1, 2), objCurItem.异常ID)
        Exit Function
    End If
    
    If mbln结算异步控制 Then
        gcnOracle.CommitTrans
        blnTrans = False: blnSaveed_Out = True
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
    '编制:刘兴洪
    '日期:2011-05-26 17:31:32
    '说明:
    '       在所有需要对已经消费的地方进行退款的情况下，调用该接口,目前规划在:收费室；挂号等。
    '---------------------------------------------------------------------------------------------------------------------------------------------

    
    If objCurItem.是否预交 Then '存在预交
         strBalanceIDs = "5|" & objCurItem.结算ID & "||1|" & objCurItem.预交ID
    Else
        strBalanceIDs = "5|" & objCurItem.结算ID
    End If
    

    strSwapExtendInfor = "5|" & lng结帐ID
    
    Call WritLogCall(mstrFunName, "zlReturnMoney", EM_CallBegin, mlngModule, objCard.接口序号, objCard.消费卡, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strSwapExtendInfor, strExpendXML, True)

    If gobjOneCardComLib.zlReturnMoney(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, strCardNo, _
        strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strSwapExtendInfor, strExpendXML, True, rsBalance, rsExpend, intStatus, strErrMsg) = False Then

        If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
        
        Call WritLogCall(mstrFunName, "zlReturnMoney", EM_CallEnd, "intStatus=" & intStatus, strErrMsg, strExpendXML)

        If strErrMsg <> "" Then
            MsgBox "注意:" & vbCrLf & "    当前退款失败，错误原因如下:" & vbCrLf & strErrMsg, vbInformation, gstrSysName
        End If
        
        ' 0-交易调用失败;1-交易正在处理中
        If intStatus = 1 Or objCurItem.冲销ID <> 0 Then    '冲销的，不能直接删除当前结算，可能后期还有其他操作
            If blnSaveed_Out Then
                For Each objItemTemp In objCurDelBalanceItems
                    objItemTemp.是否保存 = True
                    objItemTemp.是否允许编辑 = False
                    objItemTemp.是否允许删除 = False
                    objItemTemp.结算ID = objCurItem.结算ID
                    objItemTemp.冲销ID = objCurItem.冲销ID
                    objItemTemp.结算时间 = objCurItem.结算时间
                    objItemTemp.是否结算 = False
                Next
            End If
            Set objBalanceItems_Out = objCurDelBalanceItems
            
            Call WritLog(mstrFunName, "结束", "返回=False", "zlReturnMoney=false", intStatus, objCurItem.冲销ID)
            Exit Function
        End If

        If mbln结算异步控制 Then
            '无必要删除
            
        End If
        Set objBalanceItems_Out = objCurDelBalanceItems
        Call WritLog(mstrFunName, "结束", "返回=False", "zlReturnMoney=false", intStatus)
        Exit Function
    End If


    Set objCurItem = objCurDelBalanceItems(1)

    Set objBalanceItems_Out = objCurDelBalanceItems
    '更新数据
    lngCount = 0
    If Not rsBalance Is Nothing Then
        Set objBalanceItems_Out = New clsBalanceItems
        dblMoney = 0
        If rsBalance.RecordCount <> 0 Then rsBalance.MoveFirst
        With rsBalance
            Do While Not .EOF
                Set objItemTemp = objCurItem.Clone
                
                With objItemTemp
                    .卡类别ID = IIf(Val(Nvl(rsBalance!是否普通结算)) = 0, Val(Nvl(rsBalance!卡类别ID)), 0)
                    .消费卡 = False
                    .消费卡ID = 0
                    .卡号 = IIf(Nvl(rsBalance!卡号) = "", .卡号, Nvl(rsBalance!卡号))
                    .交易流水号 = Nvl(rsBalance!原交易流水号)
                    .交易说明 = Nvl(rsBalance!原交易说明)
                    .退款交易流水号 = Nvl(rsBalance!交易流水号)
                    .退款交易说明 = Nvl(rsBalance!交易说明)
                    .结算方式 = Nvl(rsBalance!结算方式)
                    .结算金额 = Val(Nvl(rsBalance!交易金额))
                    .结算号码 = Nvl(rsBalance!结算号码)
                    .结算摘要 = Nvl(rsBalance!结算摘要)
                    .校对标志 = 2
                    
                    If .卡号 <> "" Then strCardNo = .卡号
                    If strSwapNo = "" And .交易流水号 <> "" Then
                        strSwapNo = .交易流水号
                        If strSwapMemo = "" Then strSwapNo = .交易流水号
                    End If

                    dblMoney = RoundEx(dblMoney + .结算金额, 6)
                End With
                objBalanceItems_Out.AddItem objItemTemp
                objBalanceItems_Out.结算金额 = objBalanceItems_Out.结算金额 + objItemTemp.结算金额
                objBalanceItems_Out.单据号 = objCurDelBalanceItems.单据号
                
                lngCount = lngCount + 1
                .MoveNext
            Loop
            objBalanceItems_Out.结算金额 = dblMoney
        End With
        blnChangeMoney_Out = RoundEx(dblDelMoney, 6) <> RoundEx(dblMoney, 6)
    Else
        For Each objItemTemp In objBalanceItems_Out
            objItemTemp.退款交易流水号 = strSwapNo
            objItemTemp.退款交易说明 = strSwapMemo
        Next
        Set objCurItem = objCurDelBalanceItems(1)
        blnChangeMoney_Out = False
    End If


   '异常数据更新
    Set cllSaveErr = New Collection
    '入参:cllSaveData-格式为Array(保存项名称,保存项值)
    '                   保存的项目名称包含:操作状态,异常ID,操作场景,作废标志,业务id,是否病历费,病人id,主页id,姓名,性别,年龄,门诊号,住院号,预交单号,预交金额,医疗卡单号,卡费,发卡类别id,发卡类别名称,发卡卡号,同步状态,交易信息)
    '                   其中交易信息为Json串，格式如下
    '                   {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]}
    cllSaveErr.Add Array("操作状态", 1) '操作状态:0-新增记录,1-更新状态及更新交易说明
    cllSaveErr.Add Array("异常ID", objCurItem.异常ID)
    cllSaveErr.Add Array("同步状态", 2) '同步状态:0或NULL正常记录;1-未调用调接口;2-结算接口调用成功
    
    Set objItemTemp = objBalanceItems_Out(1)
    strJson = GetJsonNodeString("card_no", objItemTemp.卡号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cardtype_id", objItemTemp.卡类别ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("swapno", objItemTemp.交易流水号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("swapmemo", objItemTemp.交易说明, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("swapmoney", objBalanceItems_Out.结算金额, Json_num)
    strJson = strJson & "," & GetJsonNodeString("qrcode", "", Json_Text)
    strJson = strJson & "," & GetJsonNodeString("blnc_mode", objItemTemp.结算方式, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("blnc_no", objItemTemp.结算号码, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("blnc_memo", objItemTemp.结算摘要, Json_Text)
    
    Call zlGetThreeSwapExpendJsonStringByRecords(rsExpend_Out, strJsonExpend)
    
    If strJsonExpend <> "" Then
        strJson = strJson & "," & strJsonExpend
    End If
    strJson = "{" & strJson & "}"
    
    cllSaveErr.Add Array("交易信息", strJson)
    
    Call Zl_病人结算异常记录_Modify(1, cllSaveErr)
    If Not mbln结算异步控制 Then gcnOracle.CommitTrans
    
    blnSaveed_Out = True

    Set cllBillPro = New Collection
    blnTrans = False
    
    Call WritLog(mstrFunName, "结束", "返回=true")
    zlThird_ReturnMoney = True
    Exit Function
Errhand:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If blnTrans And blnCallInterfance = False Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call SaveErrLog
    If blnTrans Then gcnOracle.RollbackTrans

    Call WritLog(mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
    Exit Function
ErrOther:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
    Call WritLog(mstrFunName, "结束", "返回=True", "(ErrOther)", "Err=" & mstrErrInfor)
End Function



Public Function zlGetThreeSwapExpendJsonStringByRecords(ByVal rsExpend As ADODB.Recordset, strExpendJson_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取三方交易的扩展信息给Json串
    '入参:
    '出参:strExpendJson_Out-返回扩展的JSon串
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-27 17:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    On Error GoTo errHandle
    
    strExpendJson_Out = ""
    If rsExpend Is Nothing Then zlGetThreeSwapExpendJsonStringByRecords = True: Exit Function
    If rsExpend.State <> 1 Then zlGetThreeSwapExpendJsonStringByRecords = True: Exit Function
    '
    '    otherswap_list[]    C   1   其他交易信息
    '        swap_name   C   1   交易名称
    '        swap_note   C   1   交易内容
    '
    With rsExpend
        rsExpend.Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!项目名称) <> "" Then
                strTemp = GetJsonNodeString("swap_name", Nvl(!项目名称), Json_Text)
                strTemp = strTemp & "," & GetJsonNodeString("swap_note", Nvl(!项目内容), Json_Text)
                strExpendJson_Out = strExpendJson_Out & ",{" & strTemp & "}"
            End If
            .MoveNext
        Loop
    End With
    If strExpendJson_Out <> "" Then
        strExpendJson_Out = Mid(strExpendJson_Out, 2)
        strExpendJson_Out = GetNodeString("otherswap_list") & ":[" & strExpendJson_Out & "]"
    End If
    
    zlGetThreeSwapExpendJsonStringByRecords = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlGetThreeSwapExpendToCollByRecords(ByVal rsExpend As ADODB.Recordset, ByRef cllExpends_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取三方交易的扩展信息给集合
    '入参:
    '出参：cllExpends,
    '          |-cllExpend:-交易名称,交易内容
    '           格式:array(名称,值),"_名称"
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-27 17:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllExpend As Collection
    
    On Error GoTo errHandle
    
    
    Set cllExpends_Out = New Collection
    If rsExpend Is Nothing Then zlGetThreeSwapExpendToCollByRecords = True: Exit Function
    If rsExpend.State <> 1 Then zlGetThreeSwapExpendToCollByRecords = True: Exit Function
 
    With rsExpend
        rsExpend.Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!项目名称) <> "" Then
                Set cllExpend = New Collection
               cllExpend.Add Array("交易名称", Nvl(!项目名称)), "_交易名称"
               cllExpend.Add Array("交易内容", Nvl(!项目内容)), "_交易内容"
               cllExpends_Out.Add cllExpend
            End If
            .MoveNext
        Loop
    End With
    zlGetThreeSwapExpendToCollByRecords = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
 
Public Function zlThird_Payment_IsValied(ByVal objPati As clsPatientInfo, ByVal objCurItem As clsBalanceItem, _
    ByRef objBalanceItems_Out As clsBalanceItems, ByRef dbl帐户余额_Out As Double, _
    Optional byt操作类型 As Byte = 0, Optional str费用性质 As String, _
    Optional strNos As String = "", Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查三方帐户是否有效
    '入参:frmMain-调用的主窗体
    '     objPati-病人信息对象
    '     objCurItem-当前结算信息
    '     dblMoney-本次消费金额
    '     byt操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    '     str费用性质-本次支付的费用类型,多个用逗号分隔(如:1,2,3,4): 1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
    '     strNOs-划价收费时，传入
    '     blnCurInput-是否当前输入项
    '     strQRCode-二维码支付代码
    '出参:objBalanceItems_Out-返回结算数据对象集
    '     dbl帐户余额_out-返回帐户余额
    '返回:数据合法返回true(（含非三卡或消费卡或Card对象 is nothig),否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 10:30:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objItem As clsBalanceItem, strXMLExpend As String
    Dim cllBalance As Collection, str卡号 As String, str密码 As String
    Dim str结算方式 As String, dbl结算金额 As Double, varData As Variant
    Dim i As Long, dbl未付金额 As Double
    Dim strQRCode As String
    Dim mstrPara As String, dblMoney As Double
    
    On Error GoTo errHandle
    
    mstrFunName = "zlThird_Payment_IsValied"
    Call WritLog(mstrFunName, "开始", "objCurItem=" & IIf(objCurItem Is Nothing, "Nothing", "Not Nothing"))

    
    If objCurItem Is Nothing Then
        zlThird_Payment_IsValied = False
        Call WritLog(mstrFunName, "结束", "返回=False", "objCurItem Is Nothing")
        Exit Function
    End If
    
    Set objCard = objCurItem.objCard
    If objCard Is Nothing Then
        zlThird_Payment_IsValied = True
         Call WritLog(mstrFunName, "结束", "返回=true,objCard Is Nothing")
        Exit Function
    End If
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        zlThird_Payment_IsValied = True
         Call WritLog(mstrFunName, "结束", "返回=true", "objCard.接口序号=" & objCard.接口序号, " objCard.结算方式" & objCard.结算方式)
        Exit Function
    End If
    
    dblMoney = objCurItem.结算金额
    strQRCode = objCurItem.QRCode
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
         Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If
        
    
    dbl未付金额 = RoundEx(objCurItem.结算金额, 6)
     
    If dblMoney = 0 Then
        If dbl未付金额 = 0 Then
            zlThird_Payment_IsValied = True:
            Call WritLog(mstrFunName, "结束", "返回=true", "dbl未付金额=0")
            Exit Function
        End If
        Screen.MousePointer = 0
        MsgBox "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        
         Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 金额未输入,请检查置!")
        Exit Function
    End If
         
    If dblMoney <= 0 Then
        MsgBox "本次消费金额必须大于零，请检查!", vbInformation + vbOKOnly, gstrSysName
         Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 本次消费金额(" & dblMoney & ") 必须大于零，请检查")
        Exit Function
    End If
    
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "刷卡金额不能大于本次未付金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        
         Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 刷卡金额(" & dblMoney & ")不能大于本次未付金额(" & Format(Abs(dbl未付金额), "0.00") & ") ！")
        
        Exit Function
    End If
    
    
    If dblMoney <> Format(dbl未付金额, "0.00") And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1  '提示
            If MsgBox("刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                
                
                 Call WritLog(mstrFunName, "结束", "返回=False", "提示选择否，信息： " & _
                "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
                Exit Function
            End If
        Case 2
            MsgBox "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！", vbInformation, gstrSysName
                
             Call WritLog(mstrFunName, "结束", "返回=False", _
                "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
                
            Exit Function
        End Select
    End If
    Set cllBalance = New Collection
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
    '           <QRCODE></QRCODE> //识别的二维码代码串，10.35.130后支持该节点,由zlReadQRCode.clsReadQRCode
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
    
    strXMLExpend = ""
    strXMLExpend = strXMLExpend & "<IN>" & vbCrLf
    strXMLExpend = strXMLExpend & "      <CZLX>" & byt操作类型 & "</CZLX>" & vbCrLf
    If strQRCode <> "" Then
        strXMLExpend = strXMLExpend & "      <QRCODE>" & strQRCode & "</QRCODE>" & vbCrLf
    End If
    strXMLExpend = strXMLExpend & "</IN>" & vbCrLf
    
    
    '操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    If gobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, False, True, False, False, cllBalance, False, False, strXMLExpend, _
              str费用性质, objPati.病人ID, str结算方式, dbl结算金额) = False Then
        
         Call WritLog(mstrFunName, "结束", "返回=False", "调用三方接口(zlBrushCard=false) ")
        Exit Function
    End If
              
    
    If str结算方式 = "" Then
        str结算方式 = objCard.结算方式
        If dbl结算金额 = 0 Then dbl结算金额 = dblMoney
    End If
    If dbl结算金额 = 0 Then
        MsgBox "本次消费金额不能为零(接口返回0)，请与接口开发商或系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
         Call WritLog(mstrFunName, "结束", "返回=False", "本次消费金额不能为零(接口返回0)，请与接口开发商或系统管理员联系!")
        Exit Function
    End If

     Set objBalanceItems_Out = New clsBalanceItems
     
     Set objItem = objCurItem.Clone
     With objItem
        Set .objCard = objCard
        .结算方式 = str结算方式
        .结算金额 = dbl结算金额
        .卡号 = str卡号
        .卡类别ID = objCard.接口序号
        .密码 = str密码
        .QRCode = strQRCode
        .结算类型 = 3
        .限制类别 = ""
        .是否允许编辑 = True
        .是否允许删除 = True
        .是否结算 = False
        .是否密文 = IIf(objCard.卡号密文规则 <> "", True, False)
        .是否退款 = False
     End With
     objBalanceItems_Out.AddItem objItem
     objBalanceItems_Out.结算金额 = RoundEx(objBalanceItems_Out.结算金额 + dbl结算金额, 6)
 
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
    '       strXMLExpend-如下两种格式：
    '           第一种:XML串: 验证密码: 自助机用
    '           <Data>
    '               <验证密码></验证密码>
    '           </Data>
    '
    '           第二种:消费卡收款及扫码付时，传入XML串：
    '           <IN>
    '               <MZXSJE>面值销售金额</MZXSJE>
    '               <CZJKJE>充值缴款金额</CZJKJE>
    '               <QRCODE>识别的二维码代码串</QRCODE> //10.35.130后支持
    '           </IN>
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
    If strQRCode <> "" Then
        strXMLExpend = ""
        strXMLExpend = strXMLExpend & "<IN>" & vbCrLf
        strXMLExpend = strXMLExpend & "      <QRCODE>" & strQRCode & "</QRCODE>" & vbCrLf
        strXMLExpend = strXMLExpend & "</IN>" & vbCrLf
    End If
    
    If gobjOneCardComLib.zlPaymentCheck(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, str卡号, dblMoney, strNos, strXMLExpend) = False Then
         Call WritLog(mstrFunName, "结束", "返回=False", "调用三方接口(zlPaymentCheck=false)")
        Exit Function
    End If
    
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

    If gobjOneCardComLib.zlGetAccountMoney(mfrmMain, mlngModule, objCard.接口序号, _
     str卡号, strXMLExpend, dbl帐户余额_Out, objCard.消费卡) = False Then Exit Function
    If dbl帐户余额_Out <> 0 Then
        If dbl帐户余额_Out < dbl结算金额 Then
            MsgBox objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            
             Call WritLog(mstrFunName, "结束", "返回=False", "" & objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额(" & dbl帐户余额_Out & ")不足!")
            Exit Function
        End If
        For Each objItem In objBalanceItems_Out
            objItem.帐户余额 = dbl帐户余额_Out
        Next
    End If
    
    zlThird_Payment_IsValied = True
    
    Call WritLog(mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WritLog(mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function
 
Public Function zlSquare_AddBalanceDatas(ByVal objCard As Card, ByVal objBalanceItems As clsBalanceItems, _
    ByRef objBalanceItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向结算数据信息中加入消费卡结算信息
    '    objCard-当前支付的卡对象
    '    objBalanceItems-消费卡刷卡信息数据集
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-10 18:43:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim mstrPara As String
    
    
    On Error GoTo errHandle
    mstrFunName = "zlSquare_AddBalanceDatas"
    
    Call WritLog(mstrFunName, "开始", "")
    Set objBalanceItems_Out = New clsBalanceItems
    
    If objBalanceItems Is Nothing Then
        zlSquare_AddBalanceDatas = True
        Call WritLog(mstrFunName, "结束", "返回=True", "objBalanceItems Is Nothing")
        Exit Function
    End If
    If objBalanceItems.Count = 0 Then
        zlSquare_AddBalanceDatas = True
        Call WritLog(mstrFunName, "结束", "返回=True", "objBalanceItems.Count = 0")
        Exit Function
    End If
    Set objBalanceItems_Out = objBalanceItems
    
    zlSquare_AddBalanceDatas = True
    Call WritLog(mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WritLog(mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
End Function
 

 
Public Function zlSquare_Payment_IsValied(ByVal objPati As clsPatientInfo, _
    ByVal objCurItem As clsBalanceItem, ByRef objBalanceItems_Out As clsBalanceItems, ByRef dbl帐户余额_Out As Double, _
    Optional byt操作类型 As Byte = 0, Optional ByVal objOldBalanceItems As clsBalanceItems, Optional str费用性质 As String, _
    Optional ByVal rsClassFee As ADODB.Recordset, Optional strNos As String = "", Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查消费卡支付是否有效
    '入参:frmMain-调用的主窗体
    '     objPati-病人信息对象
    '     objCurItem-当前结算信息
    '     rsClassFee-可选参数，表示分收费类别统计的费用，主要应用于刷消费卡的处理(字段;收费类别,金额)
    '     byt操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    '     str费用性质-本次支付的费用类型,多个用逗号分隔(如:1,2,3,4): 1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
    '     strNOs-划价收费时，传入
    '     objOldBalanceItem-消费卡时，转入历次刷卡信息
    '     blnCurInput-是否当前输入项
    '     strQRCode-二维码支付代码
    '出参:objBalanceItems_Out-返回结算数据对象集
    '     dbl帐户余额_out-返回帐户余额
    '返回:数据合法返回true(（含非三卡或消费卡或Card对象 is nothig),否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 10:30:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objItem As clsBalanceItem, strXMLExpend As String
    Dim cllBalance As Collection, str卡号 As String, str密码 As String
    Dim str结算方式 As String, dbl结算金额 As Double, varData As Variant
    Dim i As Long, dbl未付金额 As Double
    Dim strQRCode As String
    Dim mstrPara As String, dblMoney As Double
    
    On Error GoTo errHandle
    
    mstrFunName = "zlSquare_Payment_IsValied"
    Call WritLog(mstrFunName, "开始", "objCurItem=" & IIf(objCurItem Is Nothing, "Nothing", "Not Nothing"))

    
    If objCurItem Is Nothing Then
        zlSquare_Payment_IsValied = False
        Call WritLog(mstrFunName, "结束", "返回=False", "objCurItem Is Nothing")
        Exit Function
    End If
    
    Set objCard = objCurItem.objCard
    If objCard Is Nothing Then
        zlSquare_Payment_IsValied = True
         Call WritLog(mstrFunName, "结束", "返回=True", "objCard Is Nothing")
        Exit Function
    End If
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then
        zlSquare_Payment_IsValied = True
         Call WritLog(mstrFunName, "结束", "返回=True", "objCard.接口序号=" & objCard.接口序号, " objCard.结算方式=" & objCard.结算方式)
        Exit Function
    End If
    
    dblMoney = objCurItem.结算金额
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
         Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If
        
    
    dbl未付金额 = RoundEx(objCurItem.结算金额, 6)
     
    If dblMoney = 0 Then
        If dbl未付金额 = 0 Then
            zlSquare_Payment_IsValied = True:
            Call WritLog(mstrFunName, "结束", "返回=True", "dbl未付金额=0")
            Exit Function
        End If
        Screen.MousePointer = 0
        MsgBox "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        
         Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 金额未输入,请检查置!")
        Exit Function
    End If
         
    If dblMoney <= 0 Then
        MsgBox "本次消费金额必须大于零，请检查!", vbInformation + vbOKOnly, gstrSysName
         Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 本次消费金额(" & dblMoney & ") 必须大于零，请检查")
        Exit Function
    End If
    
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "刷卡金额不能大于本次未付金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        
         Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & " 刷卡金额(" & dblMoney & ")不能大于本次未付金额(" & Format(Abs(dbl未付金额), "0.00") & ") ！")
        
        Exit Function
    End If
    
    
    If dblMoney <> Format(dbl未付金额, "0.00") And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1  '提示
            If MsgBox("刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                
                
                 Call WritLog(mstrFunName, "结束", "返回=False", "提示选择否，信息： " & _
                "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
                Exit Function
            End If
        Case 2
            MsgBox "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！", vbInformation, gstrSysName
                
             Call WritLog(mstrFunName, "结束", "返回=False", _
                "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
                
            Exit Function
        End Select
    End If
    
    
    Set cllBalance = New Collection
    For Each objItem In objOldBalanceItems
        If objItem.消费卡 Then
            'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文)
            cllBalance.Add Array(objCard.接口序号, objItem.消费卡ID, objItem.结算金额, objItem.卡号, objItem.密码, objItem.限制类别, IIf(objCard.卡号密文规则 <> "", 1, 0))
        End If
    Next
    
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
    '           <QRCODE></QRCODE> //识别的二维码代码串，10.35.130后支持该节点,由zlReadQRCode.clsReadQRCode
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
    
    strXMLExpend = ""
    '操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    If gobjOneCardComLib.zlBrushCard(mfrmMain, mlngModule, rsClassFee, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, False, True, False, True, cllBalance, True, True, strXMLExpend, _
             str费用性质, objPati.病人ID, str结算方式, dbl结算金额) = False Then
        
         Call WritLog(mstrFunName, "结束", "返回=False", "调用三方接口(zlBrushCard=false) ")
        Exit Function
    End If
              
     If RoundEx(dblMoney, 6) <> RoundEx(dbl结算金额, 6) Then
        MsgBox objCard.名称 & "支付金额为" & Format(dbl结算金额, "0.00") & "元，小于本次应付金额 " & _
                Format(dblMoney, "0.00") & "元，不能继续支付。", vbInformation, gstrSysName
        Call WritLog(mstrFunName, "结束", "返回=False", "调用三方接口(zlBrushCard=false) ")
        Exit Function
    End If
    
    
    Set objBalanceItems_Out = New clsBalanceItems
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    For Each objItem In objOldBalanceItems
        '加入已经刷的消费卡
        objBalanceItems_Out.AddItem objItem
        objBalanceItems_Out.结算金额 = objBalanceItems_Out.结算金额 + objItem.结算金额
    Next
    For i = 1 To cllBalance.Count
         varData = cllBalance(i)    'array(卡类别ID,消费卡ID,刷卡金额,卡号,密码,限制类别,是否密文 )
         Set objItem = New clsBalanceItem
         With objItem
            Set .objCard = objCard
            .结算方式 = objCard.结算方式
            .消费卡 = True
            .结算金额 = Val(varData(2))
            .卡号 = varData(3)
            .卡类别ID = varData(0)
            .消费卡ID = Val(varData(1))
            .密码 = varData(4)
            .限制类别 = varData(5)
            .是否密文 = Val(varData(6)) = 1
            .结算类型 = 5
            .是否允许编辑 = False
            .是否允许删除 = True
         End With
        objBalanceItems_Out.AddItem objItem
        objBalanceItems_Out.结算金额 = RoundEx(objBalanceItems_Out.结算金额 + objItem.结算金额, 6)
    Next

     
    
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
    '       strXMLExpend-如下两种格式：
    '           第一种:XML串: 验证密码: 自助机用
    '           <Data>
    '               <验证密码></验证密码>
    '           </Data>
    '
    '           第二种:消费卡收款及扫码付时，传入XML串：
    '           <IN>
    '               <MZXSJE>面值销售金额</MZXSJE>
    '               <CZJKJE>充值缴款金额</CZJKJE>
    '               <QRCODE>识别的二维码代码串</QRCODE> //10.35.130后支持
    '           </IN>
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
    If gobjOneCardComLib.zlPaymentCheck(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, str卡号, dblMoney, strNos, strXMLExpend) = False Then
         Call WritLog(mstrFunName, "结束", "返回=False", "zlPaymentCheck=false")
        Exit Function
    End If
    
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

    If gobjOneCardComLib.zlGetAccountMoney(mfrmMain, mlngModule, objCard.接口序号, _
     str卡号, strXMLExpend, dbl帐户余额_Out, objCard.消费卡) = False Then Exit Function
    If dbl帐户余额_Out <> 0 Then
        If dbl帐户余额_Out < dbl结算金额 Then
            MsgBox objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            
             Call WritLog(mstrFunName, "结束", "返回=False", objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额(" & dbl帐户余额_Out & ")不足!")
            Exit Function
        End If
        For Each objItem In objBalanceItems_Out
            objItem.帐户余额 = dbl帐户余额_Out
        Next
    End If
    zlSquare_Payment_IsValied = True
    Call WritLog(mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=" & mstrErrInfor)
End Function

Private Function zlSquareDel_CheckValied(ByVal objCard As Card, ByVal objPati As clsPatientInfo, ByVal dblDelMoney As Double, _
    ByVal rsMoney As ADODB.Recordset, ByRef objOldItems As clsBalanceItems, ByRef objBalanceItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡退费检查
    '入参:objCard-三方卡
    '     dblDelMoney-退款金额
    '     objOldItems-原结算对象(消费卡 ）
    '     rsMoney
    '出参:objBalanceItems_Out-返回刷卡对象
    '返回:交易合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-23 11:07:58
    '说明:同步验证了接口和刷卡接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTotal As Double, i As Integer, j As Integer, blnFind As Boolean
    Dim cllBalance As Collection, cllSquareBalance As Collection
    Dim objItem As clsBalanceItem
    Dim strXMLIn As String
    Dim mstrPara As String
    
    
    On Error GoTo errHandle
    mstrFunName = "zlSquareDel_CheckValied"
    Call WritLog(mstrFunName, "开始", dblDelMoney)
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then
        zlSquareDel_CheckValied = True
         
        Call WritLog(mstrFunName, "结束", "返回=True", "objCard.结算性质=" & objCard.结算性质, "objCard.结算方式=" & objCard.结算方式, "objCard.接口序号 =" & objCard.接口序号)
        Exit Function
    End If
    
    dblTotal = objOldItems.结算金额
    If RoundEx(dblTotal, 6) < RoundEx(dblDelMoney, 6) Then
        MsgBox "注意:" & vbCrLf & "   输入的退款金额大于了" & objCard.结算方式 & "的未退金额,请检查!" & vbCrLf & _
               "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
               "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
        
        Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=输入的退款金额大于了" & objCard.结算方式 & "的未退金额,请检查!" & vbCrLf & _
               "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
               "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"))
        Exit Function
    End If
    
    If RoundEx(dblTotal, 6) <> RoundEx(dblDelMoney, 6) Then
        If objCard.是否全退 Then
            MsgBox "注意:" & vbCrLf & "   " & objCard.结算方式 & "必须全退,请检查!" & vbCrLf & _
                   "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
                   "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
                   
            Call WritLog(mstrFunName, "结束", "返回=False", "错误信息=输入的退款金额大于了" & objCard.结算方式 & "必须全退,请检查!" & vbCrLf & _
                   "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
                   "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"))
               
            Exit Function
        End If
    End If
    
    Set cllBalance = New Collection
    For Each objItem In objOldItems
        'dblBrushMoney = GetSquareBrushMoney(objCard.接口序号, Val(Nvl(!消费卡ID)), Nvl(!卡号))
        'array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
        cllSquareBalance.Add Array(objCard.接口序号, objItem.卡类别ID, 0, objItem.卡号, "", objItem.限制类别, IIf(objItem.是否密文, 1, 0), objItem.结算金额)
    Next
    If gbln消费卡退费验卡 Then
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
        strXMLIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjOneCardComLib.zlBrushCard(Me, mlngModule, rsMoney, objCard.接口序号, _
            objCard.消费卡, objPati.姓名, objPati.性别, _
            objPati.年龄, dblDelMoney, "", "", _
            True, True, False, False, cllSquareBalance, False, False, strXMLIn) = False Then
            
            
            Call WritLog(mstrFunName, "结束", "返回=False", "zlBrushCard=false")
            Exit Function
        End If
        Set cllBalance = cllSquareBalance
    End If
    Set objBalanceItems_Out = New clsBalanceItem
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
                Call objBalanceItems_Out.AddItem(objItem)
                objBalanceItems_Out.结算金额 = RoundEx(objBalanceItems_Out.结算金额 + objItem.结算金额, 6)
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
            Call objBalanceItems_Out.AddItem(objItem)
            objBalanceItems_Out.结算金额 = RoundEx(objBalanceItems_Out.结算金额 + objItem.结算金额, 6)
        End If
    Next
    
    zlSquareDel_CheckValied = True
    Call WritLog(mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WritLog(mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call SaveErrLog
    Call WritLog(mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
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
    
    
    
    mstrFunName = "zlThirdSwapIsSwapNOCall"
   Call WritLog(mstrFunName, "开始", "zlThirdSwapIsSwapNOCall(开始)")
  
    zlThirdSwapIsSwapNOCall = gobjOneCardComLib.ZlGetParaConfig(mfrmMain, lngCardTypeID, bln消费卡, 2, strErrMsg, strExpend)
    
   Call WritLog(mstrFunName, "结束", "返回=" & IIf(zlThirdSwapIsSwapNOCall, "true", "false"))
End Function

Private Sub Class_Initialize()
    mstrOnCardPrivs = ";" & GetPrivFunc(glngSys, 1151) & ";"   '一卡通相关接口权限
    mstr脱机医保 = zlDatabase.GetPara("脱机医保结算方式", glngSys)
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mfrmMain = Nothing
    Set mobjPayCards = Nothing
End Sub



Public Function zlGetThirdMoneyInforRecordFromSwapID(ByVal str关联交易Ids As String, ByRef rsSwapRecord_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID,获取相关的结算金额信息集
    '入参:str关联交易IDs-关联交易ID，多个用逗号分离
    '出参:rsSwapRecord_Out-返回关联交易ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-27 17:19:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    On Error GoTo errHandle
    If zl_ExseSvr_GetRelatedTransInfo(str关联交易Ids, rsSwapRecord_Out, True) = False Then Exit Function
  
    zlGetThirdMoneyInforRecordFromSwapID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlExcute_OverSave(ByVal objPati As clsPatientInfo, ByVal cllSendCardInfo As Collection, ByVal objBalanceItems As clsBalanceItems, _
    ByVal blnStartEinvoice As Boolean, ByVal cllExpendInfo As Collection, Optional ByVal bln仅退病历费 As Boolean = False, Optional blnDepositEInvoice As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:完成变动数据等的保存
    '入参:objCurItem-当前支付集
    '     cllSendCardInfo-发卡信息 (卡类别ID,变动类型,卡号,原卡号,IC卡号,密码,加密密码,终止使用时间,卡费,病历费,摘要,卡号重用,领用ID),格式:array(名称,值),"_名称"
    '     blnStartEInvoice-是否启用电子票据
    '     blnDepositEinvoice-预交是否启用电子票据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-14 11:49:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllUpdateFeeData As Collection, cllTemp As Collection
    Dim objCurItem As clsBalanceItem, blnTrans As Boolean
    Dim strDepositNo As String, strCardFeeNo As String, strErrMsg As String
    Dim varTemp As Variant, lng变动ID As Long
    Dim cllPro As Collection, strSQL As String
    On Error GoTo errHandle
    
    Set cllUpdateFeeData = New Collection
    Set cllTemp = New Collection


    Set objCurItem = objBalanceItems(1)
    strCardFeeNo = objCurItem.单据号
    varTemp = Split(objCurItem.Tag & "|", "|")  '预交单|变动ID
    strDepositNo = varTemp(0)
    lng变动ID = Val(varTemp(1))
    
    cllTemp.Add Array("预交单号", strDepositNo), "_" & "预交单号"
    cllTemp.Add Array("预交ID", objCurItem.预交ID), "_" & "预交ID"
    cllTemp.Add Array("收费单号", strCardFeeNo), "_" & "收费单号"
    cllTemp.Add Array("结帐ID", IIf(objCurItem.冲销ID <> 0, objCurItem.冲销ID, objCurItem.结算ID)), "_" & "结帐ID"
    
    cllTemp.Add Array("病人ID", objPati.病人ID), "_" & "病人ID"
    cllTemp.Add Array("操作员编号", UserInfo.编号), "_" & "操作员编号"
    cllTemp.Add Array("操作员姓名", UserInfo.姓名), "_" & "操作员姓名"
    cllTemp.Add Array("收款时间", Format(objCurItem.结算时间, "yyyy-mm-dd HH:MM:SS")), "_" & "收款时间"
    
    cllTemp.Add Array("是否电子票据", IIf(blnStartEinvoice, 1, 0)), "_" & "是否电子票据"
    cllTemp.Add Array("是否预交电子票据", IIf(blnDepositEInvoice, 1, 0)), "_" & "是否预交电子票据"
    cllUpdateFeeData.Add cllTemp, "_billinfo"
    
     '结算信息
    Set cllTemp = New Collection
    Set objCurItem = objBalanceItems(1)
    cllTemp.Add Array("结算方式", objCurItem.结算方式), "_" & "结算方式"
    cllTemp.Add Array("结算号码", objCurItem.结算号码), "_" & "结算号码"
    cllTemp.Add Array("卡类别ID", IIf(objCurItem.消费卡, 0, objCurItem.卡类别ID)), "_" & "卡类别ID"
    cllTemp.Add Array("结算卡序号", IIf(objCurItem.消费卡, objCurItem.卡类别ID, 0)), "_" & "结算卡序号"
    cllTemp.Add Array("卡号", objCurItem.卡号), "_" & "卡号"
    cllTemp.Add Array("交易流水号", objCurItem.交易流水号), "_" & "交易流水号"
    cllTemp.Add Array("交易说明", objCurItem.交易说明), "_" & "交易说明"
    cllTemp.Add Array("摘要", objCurItem.结算摘要), "_" & "摘要"
    cllTemp.Add Array("合作单位", ""), "_" & "合作单位"
    
    If Not cllExpendInfo Is Nothing Then
        cllTemp.Add Array("其他信息集", cllExpendInfo), "_" & "其他信息集"
    End If
    cllUpdateFeeData.Add cllTemp, "_balanceinfo"

    
    ' cllUpdateDate-修改的结算数据
    '         |--billinfo-单据信息,"_billinfo"
    '              |-预交单号,预交ID,收费单号,结帐ID,操作员编号,操作员姓名,收款时间,病人ID,是否电子票据,是否预交电子票据)
    '         |--balanceinfo-结算信息,"_balanceinfo"
    '                |--(结算方式,结算号码,卡类别id,结算卡序号,卡号,交易流水号,交易说明,摘要,合作单位)
    '                |--其他信息集,
    '                |-----其他信息:交易名称,交易内容
    
    '2.结算完成，需要删了作异常数据及处理费用数据
    Set cllPro = New Collection

    If Not bln仅退病历费 Then
        strSQL = "Zl_病人医疗卡变动_Confirm(" & lng变动ID & "," & objPati.病人ID & ")"   '完成就诊
        zlAddArray cllPro, strSQL
    End If
 
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, mstrTittle, True
    
    Set cllTemp = New Collection
    cllTemp.Add Array("异常ID", objCurItem.异常ID), "_异常ID"
    If Zl_病人结算异常记录_Modify(2, cllTemp) = False Then
        gcnOracle.RollbackTrans: blnTrans = True
        Exit Function
    End If
    If Zl_Exsesvr_UpdCardFeeBlncInfo(0, cllSendCardInfo, cllUpdateFeeData, False, strErrMsg) = False Then
       gcnOracle.RollbackTrans: blnTrans = False
       If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
       Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False

    zlExcute_OverSave = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

