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
Private mobjInsurePara As clsInsurePara
Private mbln结算异步控制 As Boolean
Private mbyt刷卡缺省金额操作 As Byte
Private mstr脱机医保 As String
Private mobjPayCards As Cards '当前有效的支付方式


'日志相关变量
Private mstrErrInfor As String, mstrFunName As String, mblnReturnValue As Boolean
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



Public Function zlGetBalaceCardFromCards(ByVal intTYPE As Integer, ByVal objCards As Cards) As Card
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
        Select Case intTYPE
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

Public Function zlGetBalaceRecordsFromCards(ByVal objCards As Cards, ByRef rsBalance_Out As ADODB.Recordset) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算信息对象,返回卡结算记录集
    '入参:objCards-结算对象
    '出参:rsBalance_Out-返回结算信息
    '返回:成功返回=true,否则返回False
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

Public Function zlGetBalanceModeDatas(ByVal strPrivs As String, ByVal bln含脱机医保 As Boolean, ByRef objPayCards_Out As Cards, _
    Optional ByRef rsBalace_Out As ADODB.Recordset, Optional ByRef objErrCard_Out As Card, _
    Optional ByRef objOfflineYBCards_Out As Cards, Optional ByRef strQRCodeTypeIDs_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的结算信息
    '入参:strPrivs-结帐的权限串
    '     bln含脱机医保-返回的objPayCards_Out中是包含脱机医保,True-表示包包含,False-不包含（不包含时，objOfflineYBCards_Out返回Nothing )
    '
    '出参:objPayCards_Out-有效的支付方式
    '     objErrCard_Out-误差费结算对象
    '     rsBalace_Out-结算信息集
    '     objOfflineYBCards_Out-脱机医保对象
    '     strQRCodeTypeIDs_Out-返回本机支持的扫码付的卡类别IDs,多个用逗号分离
    '返回:成功返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-22 10:58:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    Dim objCards As Cards, objCard As Card
    Dim blnOnlyDeposit As Boolean, lng序号 As Long
    Dim objTemp As Card, str排除结算方式 As String
    
    On Error GoTo errHandle
    
    strQRCodeTypeIDs_Out = ""
    Set objCards = New Cards: Set objPayCards_Out = New Cards: Set objOfflineYBCards_Out = New Cards
    Set objErrCard_Out = New Card
    
    
    '性质:1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项, _
    '     6-费用折扣,7-一卡通结算,8-结算卡结算

    If InStr(1, strPrivs, ";费用打折结算;") = 0 Then
        strTmp = "1,2,3,4,5,9,7,8"
    Else
        strTmp = "1,2,3,4,5,6,9,7,8"
    End If
    
    If gTy_System_Para.TY_Balance.str预交退款结算方式 <> "" Or gTy_System_Para.TY_Balance.str作废预交结算方式 <> "" Then
        str排除结算方式 = gTy_System_Para.TY_Balance.str预交退款结算方式 & "," & gTy_System_Para.TY_Balance.str作废预交结算方式
    End If
    
    Set rsBalace_Out = Get结算方式("结帐", strTmp, str排除结算方式)
    If rsBalace_Out.RecordCount = 0 Then zlGetBalanceModeDatas = True: Exit Function

    
    
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare Is Nothing Then
        '0-所有医疗卡;1-启用的医疗卡,2-所有存在三方账户的三方卡 3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objOneCardComLib.zlGetCards(3)    '获取有效的三方帐户会付
    End If

    blnOnlyDeposit = InStr(1, strPrivs, ";允许现金结帐;") = 0
    If blnOnlyDeposit Then
        rsBalace_Out.Filter = "性质=3 Or 性质=4"
    Else
        rsBalace_Out.Filter = "性质<7"
    End If
    
    i = 1
    With rsBalace_Out
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If (InStr(",3,4,", "," & Val(NVL(!性质)) & ",") = 0) And Val(NVL(!应付款)) <> 1 Then
                Set objCard = New Card
                objCard.接口序号 = -1 * i
                objCard.接口编码 = !编码
                objCard.名称 = !名称
                objCard.结算方式 = !名称
                objCard.结算性质 = Val(NVL(!性质))
                objCard.启用 = True
                objCard.是否刷卡 = True
                objCard.缺省标志 = Val(NVL(!缺省)) = 1
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
                    If objCard.结算方式 = NVL(!名称) Then
                        Set objTemp = zlCopyNewCardFromCard(objCard)
                        objTemp.是否非接触式读卡 = False
                        objTemp.是否接触式读卡 = False
                        
                        objTemp.缺省标志 = Val(NVL(!缺省)) = 1
                        
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
    '返回:成功返回=true,否则返回False
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

Public Function zlGetYbPara() As clsInsurePara
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保参数对象
    '返回:医保参数对象
    '编制:刘兴洪
    '日期:2018-03-29 17:31:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set zlGetYbPara = mobjInsurePara
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
    '返回:初始化成功返回=true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 17:00:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mstrFunName = "zlInitCompents"
    
    Call WriteLog(mlngModule, "", mstrFunName, "开始", lngModule)


    Set mfrmMain = frmMain: mlngModule = lngModule
    
    mbln结算异步控制 = Val(zlDatabase.GetPara("费用结算异步控制", glngSys)) = 1
    mbyt刷卡缺省金额操作 = Val(zlDatabase.GetPara("刷卡缺省金额操作", glngSys, 1151, "0"))  '86853
    
    zlInitCompents = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", lngModule, IIf(mbln结算异步控制, "true", "false"), mbyt刷卡缺省金额操作)

    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
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
    '返回：接口调用成功返回=true,否则返回Flase
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intSwapStaus As Integer, objCard As Card, objItem As clsBalanceItem
    Dim strErrMsg As String
    On Error GoTo errHandle
    
    
    mstrFunName = "zlThird_IsCancelFromItems"
    Call WriteLog(mlngModule, "", mstrFunName, "开始")
    If objItems.Count = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objItems.Count=0")
        Exit Function
    End If
    
    Set objItem = objItems(1): Set objCard = objItem.objCard
    
    If objCard.消费卡 Or objItem.结帐ID = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", IIf(objCard.消费卡, "true", "false"), objItem.结帐ID)
        Exit Function
    End If

 
    If zlThird_IsSwapIsSucces(objItems, intSwapStaus, strErrMsg) Then
        MsgBox "注意:" & vbCrLf & _
               "    " & objCard.名称 & " 已经交易成功，不允许作废或删除，你必须完成结帐操作!", vbInformation + vbOKOnly, gstrSysName
               
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 已经交易成功，不允许作废或删除，你必须完成结帐操作!")
        Exit Function
    End If
    If intSwapStaus <> 0 Then
        If strErrMsg <> "" Then
            MsgBox "注意:" & vbCrLf & _
                   "    " & objCard.名称 & " 交易正在进行中,不能删除或作废,详细错误信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
            
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "返回=False", objCard.名称 & " 交易正在进行中,不能删除或作废,详细错误信息如下:" & vbCrLf & strErrMsg & "!")
        Else
            MsgBox "注意:" & vbCrLf & _
                   "    " & objCard.名称 & " 交易正在进行中，不允许作废或删除!", vbInformation + vbOKOnly, gstrSysName
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "返回=False", objCard.名称 & " 交易正在进行中，不允许作废或删除!")
        End If
        Exit Function
    End If
    zlThird_IsCancelFromItems = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "返回=true")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function

Public Function zlThird_IsSwapIsSucces(ByVal objItems As clsBalanceItems, ByRef intSwapStatu_Out As Integer, ByRef strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断交易是否调用成功
    '入参: objItems-当前结算信息集
    '出参:intSwapStatu_Out-接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    '     strErrMsg- 返回的错误信息:  为空，将不提示,不为空时，界面提示该信息
    '     strErrMsg_out-返回的错误信息
    '返回：接口调用成功返回=true,否则返回Flase
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intSwapType As Integer, dblMoney As Double, strXML As String, strXMLExpend As String
    Dim objCard As Card, objItem As clsBalanceItem, objTemp As clsBalanceItem
    Dim strBalanceIDs As String, intSign As Integer
    
    On Error GoTo errHandle
    
    mstrFunName = "zlThird_IsSwapIsSucces"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "")
    
    If objItems.Count = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objItems.Count = 0")
        Exit Function
    End If
    
    dblMoney = 0
    For Each objItem In objItems
        dblMoney = dblMoney + objItem.结算金额
    Next
    Set objItem = objItems(1)
    Set objCard = objItem.objCard
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If
    
    
    On Error GoTo errHandle
   '0-扣款;1-退款；2-转帐
    intSwapType = IIf(objItem.是否退款, IIf(objItem.是否转帐, 2, 1), 0)
    intSign = IIf(objItem.是否退款, -1, 1)
    dblMoney = RoundEx(intSign * dblMoney, 2)
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
    '    '返回：接口调用成功返回=true,否则返回Flase
    '    '日期:2013-06-15 20:22:51
    '    '说明:
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = "": strXML = ""
    For Each objTemp In objItems
        strXML = strXML & "        <TK>" & vbCrLf
        strXML = strXML & "             <JYLSH>" & objTemp.交易流水号 & "</JYLSH>" & vbCrLf
        strXML = strXML & "             <KH>" & objTemp.卡号 & "</KH>" & vbCrLf
        strXML = strXML & "             <JE>" & RoundEx(intSign * objTemp.结算金额, 6) & "</JE>" & vbCrLf
        strXML = strXML & "        </TK>" & vbCrLf
    Next
    
    strXML = "    <TKLIST>" & vbCrLf & strXML & vbCrLf
    strXML = strXML & "    </TKLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "<INPUT>" & vbCrLf
    strXMLExpend = strXMLExpend & strXML
    strXMLExpend = strXMLExpend & "</INPUT>" & vbCrLf
    
    If gobjSquare.objOneCardComLib.zlSwapIsSucces(mfrmMain, mlngModule, intSwapType, objCard.接口序号, strBalanceIDs, objItem.卡号, dblMoney, intSwapStatu_Out, strErrMsg_out, strXMLExpend) = False Then
         
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", intSwapStatu_Out, strErrMsg_out)
         zlThird_IsSwapIsSucces = False
         Exit Function
    End If
    
    zlThird_IsSwapIsSucces = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", dblMoney, intSwapStatu_Out, strErrMsg_out, strXMLExpend)
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & Err.Description)
End Function


Public Function zlThird_Payment(ByVal objCard As Card, ByVal objPati As clsPatientInfo, ByRef cllBillPro As Collection, _
    ByRef objBalanceInfor As clsBalanceInfo, ByVal objBalanceItems As clsBalanceItems, ByRef objBalanceItems_out As clsBalanceItems, _
    Optional blnSaveed_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行三方接口支付
    '入参:
    '    objCard-当前支付的卡对象
    '    objPati-病人对象集
    '    lng结帐ID-结帐ID
    '    cllBillPro-执行的相关过程
    '    lng结帐ID-结帐ID
    '    objBalanceInfor-当前结帐数据
    '    objBalanceItems-结算数据集
    '出参:objBalanceItems_Out-结帐返回数据(可能支付后结算信息发生了相关变化)
    '     objBalanceInfor-结帐返回后更新的结帐数据列表
    '     blnSaveed_Out-是否发生数据保存
    '返回:三方接口交易调用成功返回=true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 17:02:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结帐IDs As String, str卡号 As String, strSQL As String, strErrMsg_out As String
    Dim strSwapGlideNO As String, strSwapMemo As String, strExpendXML As String, str结算方式 As String
    Dim lng预交ID As Long, lngCount As Long, lng关联交易ID As Long, lng结帐ID As Long, i As Long
    Dim cllPro As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim rsBalance As ADODB.Recordset, rsExpend As ADODB.Recordset
    Dim intStatus_Out As Integer
    Dim dblMoney As Double
    Dim strQRCode As String
    Dim blnExcuteThird As Boolean, blnTrans As Boolean
  
    On Error GoTo errHandle
    
    mstrFunName = "zlThird_Payment"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "objCard=" & IIf(objCard Is Nothing, "Nothing", "Not Nothing"))
     
    blnSaveed_Out = False
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        zlThird_Payment = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", objCard.结算方式, objCard.接口序号, IIf(objCard.消费卡, "true", "false"))
        Exit Function
    End If
    
    If objBalanceItems Is Nothing Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objBalanceItems Is Nothing")
        Exit Function
    End If
    If objBalanceItems.Count = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objBalanceItems.Count = 0")
        Exit Function
    End If
    
    Set objItem = objBalanceItems(1)
    str卡号 = objItem.卡号
    strQRCode = objItem.QRCode  '二维码内码
    
    dblMoney = objItem.结算金额
    If dblMoney = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objItem.结算金额 = 0")
        Exit Function
    End If
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
     
    lng结帐ID = IIf(objBalanceInfor.冲销ID <> 0, objBalanceInfor.冲销ID, objBalanceInfor.结帐ID)
    str结帐IDs = lng结帐ID
    
    If objBalanceInfor.冲销ID <> 0 Then
         If zlGetBalanceCancelSaveSQLFromItems(objPati, objBalanceInfor, objBalanceItems, cllPro, IIf(mbln结算异步控制, 1, 2)) = False Then
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlGetBalanceCancelSaveSQLFromItems=false")
            Exit Function
        End If
         lng关联交易ID = objBalanceItems(1).关联交易ID
    Else
        If objItem.是否保存 Then
            lng预交ID = objItem.预交ID
            lng关联交易ID = objItem.关联交易ID
        End If
        If lng预交ID = 0 Then lng预交ID = zlDatabase.GetNextId("病人预交记录")
        If lng关联交易ID = 0 Then lng关联交易ID = lng预交ID
        '调用之前,先处理数据
        'Zl_病人结帐结算_Modify_S
        strSQL = "Zl_病人结帐结算_Modify_S("
        '  操作类型_In      Number,
        strSQL = strSQL & "" & "1" & ","    ' 0-普通收费方式;1.三方卡结算:2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新);3-消费卡结算:
        '  病人id_In        门诊费用记Zl_病人结帐结算_Modify录.病人id%Type,
        strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
        '  结帐id_In        病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
        '  结算方式_In      Varchar2,   '格式为:"结算方式|结算金额|结算号码|结算摘要"
        strSQL = strSQL & "'" & objItem.结算方式 & "|"
        strSQL = strSQL & objItem.结算金额 & "|"
        strSQL = strSQL & Replace(objItem.结算号码, "|", "｜") & "|"
        strSQL = strSQL & Replace(objItem.结算摘要, "|", "｜") & "',"
        '  冲预交_In        病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "" & 0 & ","
        '  退支票额_In      病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "" & 0 & ","
        '  卡类别id_In      病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & objItem.卡类别ID & ","
        '  卡号_In          病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "" & IIf(objItem.卡号 = "", "NULL", "'" & objItem.卡号 & "'") & ","
        '  交易流水号_In    病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "" & IIf(objItem.交易流水号 = "", "NULL", "'" & objItem.交易流水号 & "'") & ","
        '  交易说明_In      病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "" & IIf(objItem.交易说明 = "", "NULL", "'" & objItem.交易说明 & "'") & ","
        '  缴款_In          病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "" & objItem.缴款金额 & ","
        '  找补_In          病人预交记录.找补%Type := Null,
        strSQL = strSQL & "" & objItem.找补 & ","
        '  误差金额_In      门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "" & objItem.误差费 & ","
        '  结帐类型_In      Number := 2,
        strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
        '  缺省结算方式_In  结算方式.名称%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  收款时间_In      病人预交记录.收款时间%Type := Null,
        strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '    当前科室id_In    病人预交记录.科室id%Type := Null,
        strSQL = strSQL & ZVal(objPati.当前科室id) & ","
        '  完成结算_In      Number := 0,
        strSQL = strSQL & "" & "0" & ","
        '  校对标志_In      Number := 2,
        strSQL = strSQL & "" & IIf(mbln结算异步控制, 1, 2) & ","
        '  预交id_In        病人预交记录.Id%Type := Null,
        strSQL = strSQL & "" & lng预交ID & ","
        '  关联交易id_In    病人预交记录.Id%Type := Null,
        strSQL = strSQL & "" & lng关联交易ID & ","
        '  清除原交易_In Number:=0
        strSQL = strSQL & IIf(objItem.是否保存, "1", "0") & ")"
        zlAddArray cllPro, strSQL
    End If
     
ReDO:
    blnTrans = True
    blnExcuteThird = False
     
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption & "-执行三方接口", Not mbln结算异步控制, , "结帐", "病人结帐调试日志"
    
    If mbln结算异步控制 Then blnTrans = False: blnSaveed_Out = True
    
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
    '    rsExpend_Out-扩展信息集(blnResolveXMLToRecord=true时返回),字段（卡类别ID,消费卡,项目名称,项目内容)
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
    '返回:扣款成功,返回=true,否则返回Flase
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
    
    If gobjSquare.objOneCardComLib.zlPaymentMoney(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, objItem.卡号, _
         str结帐IDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strExpendXML, True, rsBalance, rsExpend, intStatus_Out, strErrMsg_out) = False Then
        blnExcuteThird = False
        If blnTrans Then gcnOracle.RollbackTrans
        blnTrans = False
        'intStatus_Out-交易状态:接口返回False时，此参数有效: 0-交易调用失败;1-交易正在处理中
        'strErrMsg_Out-错误信息:为空时，不提示，非空时，提示
        If strErrMsg_out <> "" Then
            MsgBox "交易正在进行中，原因如下:" & vbCrLf & strErrMsg_out, vbOKOnly + vbInformation, gstrSysName
        End If
        
        If blnSaveed_Out Then Set cllBillPro = New Collection
        If intStatus_Out = 1 Then   '1-交易正在处理中
            If blnSaveed_Out Then
                For Each objItem In objBalanceItems
                    objItem.是否保存 = True
                    objItem.是否允许编辑 = False
                    objItem.是否允许删除 = True
                    objItem.结帐ID = objBalanceInfor.结帐ID
                    objItem.结帐时间 = objBalanceInfor.结帐时间
                    objItem.关联交易ID = IIf(lng关联交易ID <> 0, lng关联交易ID, objItem.关联交易ID)
                    objItem.预交ID = lng预交ID
                    objItem.QRCode = ""
                Next
                Set objBalanceItems_out = objBalanceItems
            End If
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false", blnSaveed_Out, strSwapGlideNO, strSwapMemo, strExpendXML, intStatus_Out, strErrMsg_out)
            Exit Function
        End If
        objBalanceItems(1).关联交易ID = lng关联交易ID
        objBalanceItems(1).结帐ID = objBalanceInfor.结帐ID
        If mbln结算异步控制 Then Call zlDelBalanceItemFromItemObject(objBalanceItems(1))
        Set objBalanceItems_out = New clsBalanceItems
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=false", blnSaveed_Out, strSwapGlideNO, strSwapMemo, strExpendXML, intStatus_Out, strErrMsg_out)
        Exit Function
    End If
    
      blnExcuteThird = False
    Set objBalanceItems_out = New clsBalanceItems
    '更新数据
    If Not rsBalance Is Nothing Then
        With rsBalance
            .Filter = 0
            If .RecordCount <> 0 Then .MoveFirst: dblMoney = 0
          
            Do While Not .EOF
                If objBalanceInfor.冲销ID <> 0 Then
                    lng关联交易ID = objItem.关联交易ID
                    lng预交ID = 0
                Else
                    If lng预交ID = 0 Then lng预交ID = zlDatabase.GetNextId("病人预交记录")
                    If lng关联交易ID = 0 Then lng关联交易ID = lng预交ID
                End If
                Set objItemTemp = New clsBalanceItem
                With objItemTemp
                    .卡类别ID = IIf(Val(NVL(rsBalance!是否普通结算)) = 0, Val(NVL(rsBalance!卡类别ID)), 0)
                    .消费卡 = False
                    .消费卡ID = 0
                    .卡号 = IIf(NVL(rsBalance!卡号) = "", objItem.卡号, NVL(rsBalance!卡号))
                    If .卡号 <> "" Then str卡号 = .卡号
                    .交易流水号 = NVL(rsBalance!交易流水号)
                    .交易说明 = NVL(rsBalance!交易说明)
                    .结算方式 = NVL(rsBalance!结算方式)
                    .结算金额 = Val(NVL(rsBalance!交易金额))
                    .密码 = objItem.密码
                    .是否密文 = objItem.是否密文
                    .限制类别 = objItem.限制类别
                    .帐户余额 = objItem.帐户余额
                    .结算号码 = NVL(rsBalance!结算号码)
                    .结算摘要 = NVL(rsBalance!结算摘要)
                    .结帐时间 = objBalanceInfor.结帐时间
                    .是否允许编辑 = False
                    .是否允许删除 = False
                    .是否保存 = True
                    .是否结算 = True
                    .结算IDs = objBalanceInfor.结帐ID
                    .结帐ID = objBalanceInfor.结帐ID
                    .预交ID = lng预交ID
                    .关联交易ID = lng关联交易ID
                    .是否结算 = True
                    .结算类型 = 3   ''0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                    .校对标志 = 2
                End With
                objBalanceItems_out.AddItem objItemTemp
                If objBalanceInfor.冲销ID = 0 Then
                    
                    'Zl_病人结帐结算_Modify_S
                    strSQL = "Zl_病人结帐结算_Modify_S("
                    '  操作类型_In      Number,
                    strSQL = strSQL & "" & "1" & ","    ' 0-普通收费方式;1.三方卡结算:2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新);3-消费卡结算:
                    '  病人id_In        门诊费用记录.病人id%Type,
                    strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
                    '  结帐id_In        病人预交记录.结帐id%Type,
                    strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
                    '  结算方式_In      Varchar2,   '格式为:"结算方式|结算金额|结算号码|结算摘要"
                    strSQL = strSQL & "'" & objItemTemp.结算方式 & "|"
                    strSQL = strSQL & objItemTemp.结算金额 & "|"
                    strSQL = strSQL & Replace(objItemTemp.结算号码, "|", "｜") & "|"
                    strSQL = strSQL & Replace(objItemTemp.结算摘要, "|", "｜") & "',"
                    '  冲预交_In        病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "" & 0 & ","
                    '  退支票额_In      病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "" & 0 & ","
                    '  卡类别id_In      病人预交记录.卡类别id%Type := Null,
                    strSQL = strSQL & "" & IIf(objItemTemp.卡类别ID = 0, "NULL", objItemTemp.卡类别ID) & ","
                    '  卡号_In          病人预交记录.卡号%Type := Null,
                    strSQL = strSQL & "" & IIf(objItemTemp.卡号 = "", "NULL", "'" & objItemTemp.卡号 & "'") & ","
                    '  交易流水号_In    病人预交记录.交易流水号%Type := Null,
                    strSQL = strSQL & "" & IIf(objItemTemp.交易流水号 = "", "NULL", "'" & objItemTemp.交易流水号 & "'") & ","
                    '  交易说明_In      病人预交记录.交易说明%Type := Null,
                    strSQL = strSQL & "" & IIf(objItemTemp.交易说明 = "", "NULL", "'" & objItemTemp.交易说明 & "'") & ","
                    '  缴款_In          病人预交记录.缴款%Type := Null,
                    strSQL = strSQL & "" & 0 & ","
                    '  找补_In          病人预交记录.找补%Type := Null,
                    strSQL = strSQL & "" & 0 & ","
                    '  误差金额_In      门诊费用记录.实收金额%Type := Null,
                    strSQL = strSQL & "" & 0 & ","
                    '  结帐类型_In      Number := 2,
                    strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
                    '  缺省结算方式_In  结算方式.名称%Type := Null,
                    strSQL = strSQL & "" & "NULL" & ","
                    '  操作员编号_In    病人预交记录.操作员编号%Type := Null,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '  操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '  收款时间_In      病人预交记录.收款时间%Type := Null,
                    strSQL = strSQL & "to_date('" & Format(objItemTemp.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                    '  冲预交病人ids_In Varchar2 := Null,
                    strSQL = strSQL & "" & "NULL" & ","
                    '    当前科室id_In    病人预交记录.科室id%Type := Null,
                    strSQL = strSQL & ZVal(objPati.当前科室id) & ","
                    '  完成结算_In      Number := 0,
                    strSQL = strSQL & "" & "0" & ","
                    '  校对标志_In      Number := 2,
                       strSQL = strSQL & "" & "2" & ","
                    '  预交id_In        病人预交记录.Id%Type := Null,
                    strSQL = strSQL & "" & lng预交ID & ","
                    '  关联交易id_In    病人预交记录.Id%Type := Null,
                    strSQL = strSQL & "" & lng关联交易ID & ","
                    '  清除原交易_In Number:=0
                    strSQL = strSQL & "" & IIf(lngCount = 0, 1, 0) & ")"
                    zlAddArray cllUpdate, strSQL
                End If
                dblMoney = RoundEx(dblMoney + objItemTemp.结算金额, 6)
                lng预交ID = 0
                lngCount = lngCount + 1
                .MoveNext
            Loop
            objBalanceItems_out.结算金额 = dblMoney
            If objBalanceInfor.冲销ID <> 0 Then
                 If zlGetBalanceCancelSaveSQLFromItems(objPati, objBalanceInfor, objBalanceItems_out, cllUpdate, 2) = False Then
                    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlGetBalanceCancelSaveSQLFromItems=false")
                    Exit Function
                End If
            End If
            
        End With
    Else
        Set objItemTemp = New clsBalanceItem
        With objItemTemp
             Set .objCard = objCard
            .交易流水号 = strSwapGlideNO
            .交易说明 = strSwapMemo
            .结算方式 = objItem.结算方式
            .结算金额 = objItem.结算金额
            .卡号 = objItem.卡号
            .卡类别ID = objItem.卡类别ID
            .密码 = objItem.密码
            .是否密文 = objItem.是否密文
            .限制类别 = objItem.限制类别
            .消费卡 = False
            .消费卡ID = 0
            .帐户余额 = objItem.帐户余额
            .结算IDs = objBalanceInfor.结帐ID
            .结帐ID = objBalanceInfor.结帐ID
            .预交ID = lng预交ID
            .关联交易ID = lng关联交易ID
            .是否结算 = True
            .结算类型 = 3
            .是否结算 = True
            .是否允许编辑 = False
            .是否允许删除 = False
            .是否保存 = True
            .是否结算 = True
            .校对标志 = 2
            .结算IDs = objBalanceInfor.结帐ID
        End With
        objBalanceItems_out.AddItem objItemTemp
        objBalanceItems_out.结算金额 = objItem.结算金额
        If objBalanceInfor.冲销ID <> 0 Then
            Call zlAddUpdateSwapSQL(False, objBalanceInfor.冲销ID, objCard.接口序号, objCard.消费卡, objItem.卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 0, lng关联交易ID)
        Else
            Call zlAddUpdateSwapSQL(False, str结帐IDs, objCard.接口序号, objCard.消费卡, objItem.卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 0, lng关联交易ID)
        End If
    End If
    
    blnTrans = True
    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption & "-执行三方接口", , Not mbln结算异步控制, "结帐", "病人结帐调试日志"
    blnSaveed_Out = True
    
    zlThird_Payment = True
    
    blnTrans = False
    Set cllBillPro = New Collection
    
    dblMoney = 0
    For Each objItem In objBalanceItems_out
        dblMoney = dblMoney + objItem.结算金额
    Next
    
   '增加结算数据
    objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + dblMoney, 6)
    objBalanceInfor.未付合计 = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.已付合计, 6)
    objBalanceInfor.是否保存结帐单 = True
    Call zlGetThreeSwapExpendSQLToCollect(objCard, lng结帐ID, str卡号, rsExpend, cllThreeSwap, False)
    
    Err = 0: On Error GoTo ErrOther:
    blnTrans = True
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption & "-执行三方接口", , , "结帐", "病人结帐调试日志"
    
    blnTrans = False
    
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True", strSwapGlideNO, strSwapMemo, strExpendXML, intStatus_Out, strErrMsg_out)
    Exit Function
errHandle:
    
    If blnTrans Then gcnOracle.RollbackTrans
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        If blnExcuteThird Then GoTo ReDO:
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call SaveErrLog
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "错误信息=" & mstrErrInfor)
    Exit Function
ErrOther:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "ErrOther,重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "ErrOther,错误信息=" & mstrErrInfor)
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
    Dim str预交IDs As String
    Dim i As Long
    
    Err = 0: On Error GoTo errHandle
    
    mstrFunName = "zlThirdReturnCashCheck"
    Call WriteLog(mlngModule, "", mstrFunName, "开始")

    If objBalanceItems Is Nothing Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objBalanceItems Is Nothing")
        Exit Function
    End If
    If objBalanceItems.Count = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objBalanceItems.Count = 0")
        Exit Function
    End If
    
    strBalanceIDs = "": str预交IDs = ""
    strIDs = ""
    For i = 1 To objBalanceItems.Count
        Set objItemTemp = objBalanceItems(i)
        If objItemTemp.卡号 <> "" And objItem Is Nothing Then Set objItem = objItemTemp
        If objItemTemp.是否预交 Then
            str预交IDs = str预交IDs & "," & IIf(objItemTemp.预交ID = 0, objItemTemp.结算IDs, objItemTemp.预交ID)
        Else
             strIDs = strIDs & "," & objItemTemp.结算IDs
        End If
        strTKList = ""
        strTKList = strTKList & Space(8) & "<TK>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <TKFS>" & objItemTemp.结算方式 & "</TKFS>" & vbCrLf
        strTKList = strTKList & Space(8) & "    <TKJE>" & RoundEx(-1 * objItemTemp.结算金额, 6) & "</TKJE>" & vbCrLf
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
    '    '       blnDelDefaultCash_Out-是否缺省退现：接口返回=true时有效，true时：表示缺省退成现金（缺省方式为:str缺省退现方式_Out返回值),否则缺省退回原卡，但允许操作员选择退为现金
    '    '       strDefaultDelBalance_Out-缺省退现方式,比如：支票，现金等
    '    '       strXMLExpend:10.35.110后才支持
    '    '        <OUTPUT>
    '    '            <SFQSTX>是否缺省退现<SFQSTX>//NUMBER 1 是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但以许操作员操作退现
    '    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
    '    '                    1.不允许返回三卡方的结算方式
    '    '                    2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
    '    '        </OUTPUT>
    '    '返回:退现合法,返回=true,否则返回Flase
    
    If gobjSquare.objOneCardComLib.zlReturnCashCheck(mfrmMain, mlngModule, objCard.接口序号, objItem.卡号, _
          strBalanceIDs, dblMoney, objItem.交易流水号, objItem.交易说明, strXMLExpend, blnDefaultCash_Out, strDefaultBalance_Out) = False Then
          
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlReturnCashCheck=false", strXMLExpend, blnDefaultCash_Out, strDefaultBalance_Out)
        zlThirdReturnCashCheck = False
        Exit Function
    End If
    zlThirdReturnCashCheck = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", strXMLExpend, blnDefaultCash_Out, strDefaultBalance_Out)
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function

Public Function zlThird_TransferAccounts(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef objTranItem As clsBalanceItem, ByRef cllBillPro As Collection, _
    Optional strErrMsg_out As String, Optional blnSaveed_Out As Boolean, Optional ByVal intOprType As Integer, Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通转帐支付
    '入参:
    '    cllBillPro-执行的相关过程
    '    objBalanceInfor-当前结帐数据
    '    objTranItem-转账项
    '    blnCurInput-是否当前输入项
    '    intOprType-类型:0-结帐转帐;1-异常作废;2-结帐作废转帐
    '出参:objBalanceInfor-结帐返回后更新的结帐数据列表
    '     strErrMsg_Out-返回的错误信息
    '     blnSaveed_Out-是否发生数据保存
    '     blnChangeMoney_Out-金额是否发生变化
    '返回:三方接口交易调用成功返回=true,否则返回Fale
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
    Dim objCard As Card, blnCheck As Boolean, blnCallInterFace As Boolean   '开始调用接口
    Dim lng结帐ID As Long
    Dim blnNotBegin As Boolean
    
    
    Err = 0: On Error GoTo ErrHand:
    
    mstrFunName = "zlThird_TransferAccounts"
    Call WriteLog(mlngModule, "", mstrFunName, "开始")
    
    
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    
    blnSaveed_Out = False
    
    If objTranItem Is Nothing Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objTranItem Is Nothing")
        Exit Function
    End If

    If objTranItem.objTag Is Nothing Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objTranItem.objTag Is Nothing")
        Exit Function
    End If
    
    If UCase(TypeName(objTranItem.objTag)) <> UCase("clsBalanceItems") Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objTranItem.objTag=" & UCase(TypeName(objTranItem.objTag)))
        Exit Function
    End If
        
    Set objItems = objTranItem.objTag
    Set objCard = objTranItem.objCard
    
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        zlThird_TransferAccounts = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", objCard.接口序号, IIf(objCard.消费卡, "true", "false"))
        Exit Function
    End If
    If Not objCard.是否转帐及代扣 Then
    
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "不支持转帐及代扣,结算方式=" & objCard.结算方式)
        Exit Function
    End If
 
    
    dblMoney = -1 * objTranItem.结算金额
    lng结帐ID = IIf(objBalanceInfor.冲销ID <> 0, objBalanceInfor.冲销ID, objBalanceInfor.结帐ID)
    
    '先检查是否合法
    blnCheck = True
    If objTranItem.是否保存 Then
        If Not zlThird_IsSwapIsSucces(objItems, intStatus, strErrMsg) Then
           If intStatus = 1 Then
                If strErrMsg <> "" Then
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行转帐处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 正在进行转帐处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg)
                Else
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行转帐处理，请稍后再试!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 正在进行转帐处理，请稍后再试!")
                End If
                Exit Function
           End If
        Else
            blnCheck = False
        End If
    End If
    
    If blnCheck Then
        If zlThird_TransferAccounts_IsValied(objPati, objCard, objBalanceInfor, objItems, objItemsTemp, dbl帐户余额, strErrMsg_out, blnCurInput) = False Then
        
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlThird_TransferAccounts_IsValied=false")
            Exit Function
        End If
    End If
    
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
        
    Next
    
    If Not objItemsTemp Is Nothing Then
        If objItemsTemp.Count <> 0 Then
           If objItemsTemp(1).结算方式 <> "" Then objTranItem.结算方式 = objItemsTemp(1).结算方式
           
           objTranItem.卡号 = objItemsTemp(1).卡号
           objTranItem.交易流水号 = objItemsTemp(1).交易流水号
           objTranItem.交易说明 = objItemsTemp(1).交易说明
        End If
    End If
    
    
     strCardNo = objTranItem.卡号
     strSwapNO = objTranItem.交易流水号
     strSwapMemo = objTranItem.交易说明
     
     str结算方式 = objCard.结算方式
     
     '结算方式|结算金额|结算号码|结算摘要||..
     str结算方式 = objTranItem.结算方式 & "|" & objTranItem.结算金额
     str结算方式 = str结算方式 & "|" & IIf(objTranItem.结算号码 = "", " ", objTranItem.结算号码)
     str结算方式 = str结算方式 & "|" & "转帐结算"
    
    If objBalanceInfor.冲销ID <> 0 Then
        '已经存在
        If intOprType = 2 Then '结帐作废发生的转帐
            If zlGetBalanceCancelSaveSQLFromItems(objPati, objBalanceInfor, objItems, cllPro, IIf(mbln结算异步控制, "1", "2")) = False Then
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlGetBalanceCancelSaveSQLFromItems=false")
                Exit Function
            End If
            If objItems.Count <> 0 Then objTranItem.关联交易ID = objItems(1).关联交易ID
            lng关联交易ID = objTranItem.关联交易ID
        End If
    Else
        If Not objTranItem.是否保存 Or objTranItem.预交ID = 0 Then
           lng预交ID = zlDatabase.GetNextId("病人预交记录")
           lng关联交易ID = lng预交ID
           objTranItem.预交ID = lng预交ID
           objTranItem.关联交易ID = lng关联交易ID
        End If
        
        lng预交ID = objTranItem.预交ID
        lng关联交易ID = objTranItem.关联交易ID
        
        '调用之前,先处理数据
        'Zl_病人结帐结算_Modify_S
        strSQL = "Zl_病人结帐结算_Modify_S("
        '  操作类型_In     Number,
        '  --操作类型_In:
        '  --   0-普通收费方式:
        '  --   1.三方卡结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②退支票额_In:传入零
        '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        strSQL = strSQL & "1,"
        '  病人id_In       病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
        '  结帐id_In       病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
        '  结算方式_In     Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  冲预交_In       病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  退支票额_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & objCard.接口序号 & ","
        '  卡号_In         病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & strSwapNO & "',"
        '  交易说明_In     病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & strSwapMemo & "',"
        '  缴款_In         病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "" & objBalanceInfor.缴款 & ","
        '  找补_In         病人预交记录.找补%Type := Null,
        strSQL = strSQL & "" & objBalanceInfor.找补 & ","
        '  误差金额_In     门诊费用记录.实收金额%Type := Null,
        strSQL = strSQL & "NULL,"
        '  结帐类型_In     Number := 2,
        strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
        '  缺省结算方式_In 结算方式.名称%Type := Null,
        strSQL = strSQL & "NULL,"
        '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '    当前科室id_In    病人预交记录.科室id%Type := Null,
        strSQL = strSQL & ZVal(objPati.当前科室id) & ","
        '  完成结算_In      Number := 0,
        strSQL = strSQL & "" & "0" & ","
        ' 校对标志_In      Number := 2,
        strSQL = strSQL & "" & IIf(mbln结算异步控制, 1, 2) & ","
        '  预交id_In        病人预交记录.Id%Type := Null,
        strSQL = strSQL & "" & lng预交ID & ","
        '  关联交易id_In    病人预交记录.Id%Type := Null,
        strSQL = strSQL & "" & lng关联交易ID & ","
        '  清除原交易_In Number:=0
        strSQL = strSQL & "" & IIf(objTranItem.是否保存, 1, 0) & ","
        '附加标志_In    病人预交记录.Id%Type := Null
        '对于结帐业务(三方卡退款传入)：NULl or 0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
        strSQL = strSQL & "" & 3 & ")"
        zlAddArray cllPro, strSQL
     End If
    
    
    If mbln结算异步控制 And objBalanceInfor.冲销ID = 0 Then
        If zlGetSaveThirdSwapDelSQLFromBalanceItems(objItems, True, cllUpdate, 2) = False Then
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlGetSaveThirdSwapDelSQLFromBalanceItems=false")
            Exit Function
        End If
    ElseIf mbln结算异步控制 And objBalanceInfor.冲销ID <> 0 Then
        Set cllUpdate = New Collection
        If zlGetBalanceCancelSaveSQLFromItems(objPati, objBalanceInfor, objItems, cllUpdate, 2) = False Then
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlGetBalanceCancelSaveSQLFromItems=false")
            Exit Function
        End If
                
    End If
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, Not mbln结算异步控制, , "结帐", "病人结帐调试日志"
    If mbln结算异步控制 Then blnTrans = False: blnSaveed_Out = True: objTranItem.是否保存 = True
    
    blnTrans = True
    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, True, Not blnSaveed_Out
    Set cllUpdate = New Collection
     
    blnCallInterFace = True
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

    strXMLExpend = ""
    strXMLExpend = strXMLExpend & "<IN>" & vbCrLf
    strXMLExpend = strXMLExpend & "     <CZLX>" & IIf(objBalanceInfor.冲销ID <> 0, 3, 2) & "</CZLX>" & vbCrLf
    strXMLExpend = strXMLExpend & "</IN>" & vbCrLf
    
    strSwapExtendInfor = "2|" & lng结帐ID
    
    If gobjSquare.objOneCardComLib.zlTransferAccountsMoney(mfrmMain, mlngModule, objCard.接口序号, _
        strCardNo, objBalanceInfor.结帐ID, Abs(dblMoney), strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend, rsExpend, strErrMsg_out, intStatus) = False Then
        
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
        If objTranItem.冲销ID = 0 And objBalanceInfor.冲销ID <> 0 Then objTranItem.冲销ID = objBalanceInfor.冲销ID
        If objTranItem.结帐ID = 0 And objBalanceInfor.结帐ID <> 0 Then objTranItem.结帐ID = objBalanceInfor.结帐ID
        
        blnCallInterFace = False
        If mbln结算异步控制 Then Call zlDelBalanceItemFromItemObject(objTranItem)
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlTransferAccountsMoney=false", intStatus, strErrMsg_out, strXMLExpend)
        Exit Function
    End If
    
    If mbln结算异步控制 And objTranItem.是否预交 Then
        If zlGetSaveThirdSwapDelSQLFromBalanceItems(objItems, True, cllUpdate, 2) = False Then
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlGetSaveThirdSwapDelSQLFromBalanceItems=false", intStatus, strErrMsg_out, strXMLExpend)
            Exit Function
        End If
    End If
    
    Call zlAddUpdateSwapSQL(False, lng结帐ID, objCard.接口序号, False, strCardNo, strSwapNO, strSwapMemo, cllUpdate, 2, lng关联交易ID)
    
    
    blnNotBegin = blnTrans
    blnCallInterFace = False
    
    blnTrans = True
    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, , blnNotBegin, "结帐", "病人结帐调试日志"
    blnTrans = False
    
    If Not mbln结算异步控制 Then blnSaveed_Out = True
    
    zlThird_TransferAccounts = True
    
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
    
    Call zlGetThreeSwapExpendSQLToCollect(objCard, lng结帐ID, strCardNo, rsExpend, cllThreeSwap, False)
    
    On Error GoTo ErrOther:
    blnTrans = True
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption, , , "结帐", "病人结帐调试日志"
    blnTrans = False
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", intStatus, strErrMsg_out, strXMLExpend)
    Exit Function
ErrHand:
    If blnTrans And blnCallInterFace = False Then blnNotBegin = False: gcnOracle.RollbackTrans: blnTrans = False
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
    Exit Function
ErrOther:
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "(ErrOther),Err=" & Err.Description)
    If blnTrans Then gcnOracle.CommitTrans
    Call ErrCenter
    Call SaveErrLog
    
End Function


Public Function zlThird_ReturnMoney_IsValied(ByVal objPati As clsPatientInfo, ByVal objCard As Card, ByVal byt操作类型 As Byte, _
    ByVal objBalanceInfor As clsBalanceInfo, ByVal objDelBalanceItems As clsBalanceItems, ByRef objItems_Out As clsBalanceItems, _
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
    '返回:允许退款返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-28 18:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend  As String, str交易说明 As String, str交易流水号 As String, strBalanceIDs As String, str卡号 As String, str密码 As String
    Dim bln退现 As Boolean, intStatus As Integer, strErrMsg As String
    Dim dblMoney  As Double, dbl原始金额 As Double, dbl未退金额 As Double, rsSwap As ADODB.Recordset
    Dim str退费结算IDs As String, str预交退费结算IDs As String, str关联交易IDs As String
    Dim strXmlIn As String
    Dim objItem As clsBalanceItem
    Dim i As Long, blnCheck As Boolean
    
    On Error GoTo errHandle
    
    mstrFunName = "zlThird_ReturnMoney_IsValied"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "")
    
    If objCard Is Nothing Then
        Screen.MousePointer = 0
        If InStr(";" & mstrOnCardPrivs & ";", ";三方接口消费;") = 0 Then
            MsgBox "你没有三方接口消费权限，无法调用接口部件！", vbInformation, gstrSysName
            
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "你没有三方接口消费权限，无法调用接口部件！")
            Exit Function
        End If
        MsgBox "未找到退款接口,请检查接口部件！", vbInformation, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "未找到退款接口,请检查接口部件")
        Exit Function
    End If
    
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If
        
        
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        zlThird_ReturnMoney_IsValied = True
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", objCard.接口序号, IIf(objCard.消费卡, "true", "false"), objCard.结算方式)
        Exit Function
    End If
    
    
    If objDelBalanceItems Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "当前退款结算信息不存在，请检查!")
        Exit Function
    End If
    
    If objDelBalanceItems.Count = 0 Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "当前退款结算信息不存在，请检查（objDelBalanceItems.Count＝0）")
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
                     
                     Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 正在进行扣款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg)
                 Else
                     MsgBox "注意:" & vbCrLf & _
                     "   " & objCard.名称 & " 正在进行扣款处理中，请稍后再试", vbInformation + vbOKOnly, gstrSysName
                     
                     Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 正在进行扣款处理中，请稍后再试")
                 End If
                 Exit Function
            End If
        Else
            Set objItems_Out = objDelBalanceItems
            zlThird_ReturnMoney_IsValied = True: blnCheck = False:
            
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True", "调用zlThird_IsSwapIsSucces=true")
            Exit Function
        End If
    End If
 
         
         
    dblMoney = objDelBalanceItems.结算金额
    dbl未退金额 = RoundEx(objBalanceInfor.未付合计 - objBalanceInfor.冲预交合计 - objBalanceInfor.误差费, 2)
    
   ' RoundEx(objBalanceInfor.未付合计 - objBalanceInfor.冲预交合计 + IIf(blnCurInput, 0, dblMoney) - objBalanceInfor.误差费, 2) '不是当前录入的，需要恢复本次的来判断
    
    If dblMoney = 0 Then
        If dbl未退金额 = 0 Then
            zlThird_ReturnMoney_IsValied = True
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true,dbl未退金额=0")
            Exit Function
        End If
        Screen.MousePointer = 0
        MsgBox "退款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "退款金额未输入,请检查!")
        Exit Function
    End If
    
    If dbl未退金额 = 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & objCard.名称 & "已经退完，不能再退！", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & "已经退完，不能再退！")
        Exit Function
    End If
    
    If RoundEx(Abs(dblMoney), 2) > RoundEx(Abs(dbl未退金额), 2) And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！", vbInformation, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！")
        Exit Function
    End If
       
   
    If Abs(dblMoney) <> Abs(dbl未退金额) And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1 '提示
            Screen.MousePointer = 0
            If MsgBox("当前退款金额与未退金额不符,是否继续？" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(Abs(dblMoney), "0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                
                
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "错误信息(选择否)=" & _
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
                
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "错误信息=" & _
                "当前退款金额与未退金额不符,不能继续!" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(Abs(dblMoney), "0.00"))
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
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", " 未找到原始的结算记录,不能使用" & objCard.名称 & "进行退款!")
        Exit Function
    End If
    
    str关联交易IDs = ""
    dbl原始金额 = 0: dbl未退金额 = 0
    For Each objItem In objDelBalanceItems
        
        If InStr(str关联交易IDs & ",", "," & objItem.关联交易ID & ",") = 0 And objItem.关联交易ID <> 0 Then
            str关联交易IDs = str关联交易IDs & "," & objItem.关联交易ID
            rsSwap.Filter = "关联交易ID=" & objItem.关联交易ID
            Do While Not rsSwap.EOF
                dbl未退金额 = RoundEx(dbl未退金额 + Val(NVL(rsSwap!剩余未退金额)), 6)
                dbl原始金额 = RoundEx(dbl原始金额 + Val(NVL(rsSwap!原始金额)), 6)
                rsSwap.MoveNext
            Loop
            
            If objCard.是否全退 Then
                If Abs(dblMoney) <> dbl原始金额 Then
                    Screen.MousePointer = 0
                    Call MsgBox(objCard.名称 & "必须全退,不能继续!" & vbCrLf & _
                         "原始金额:" & Format(dbl原始金额, "0.00") & vbCrLf & _
                         "当前退款:" & Format(Abs(dblMoney), "0.00") & vbCrLf, vbInformation + vbOKOnly, gstrSysName)
                         
                    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & "必须全退,不能继续!" & vbCrLf & _
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
                    
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "当前退款金额大于了剩余未退金额,不能继续!" & vbCrLf & _
                    "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & vbCrLf & _
                    "当前退款:" & Format(Abs(dblMoney), "0.00") & vbCrLf)
                Exit Function
            End If
        End If
        If InStr(str关联交易IDs & ",", "," & objItem.关联交易ID & ",") = 0 And objItem.关联交易ID <> 0 Then
            str关联交易IDs = str关联交易IDs & "," & objItem.关联交易ID
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
    '    '返回:退款合法,返回=true,否则返回Flase
    '    '编制:刘兴洪
    '    '日期:2011-05-26 17:24:55
    '    '说明:
    '    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
    dblMoney = 0
    strXMLExpend = ""
    str退费结算IDs = ""
    str预交退费结算IDs = ""
    For i = 1 To objDelBalanceItems.Count
        dblMoney = RoundEx(dblMoney + -1 * objDelBalanceItems(i).结算金额, 6)
        If objDelBalanceItems(i).是否预交 Then
            If InStr(str预交退费结算IDs & ",", "," & objDelBalanceItems(i).结算IDs & ",") = 0 Then str预交退费结算IDs = str预交退费结算IDs & "," & objDelBalanceItems(i).预交ID
        Else
           If InStr(str退费结算IDs & ",", "," & objDelBalanceItems(i).结算IDs & ",") = 0 Then str退费结算IDs = str退费结算IDs & "," & objDelBalanceItems(i).结算IDs
        End If
        strXMLExpend = strXMLExpend & vbCrLf & _
        "              <TK>" & vbCrLf & _
        "                   <TKFS>" & objDelBalanceItems(i).结算方式 & "</TKFS>" & vbCrLf & _
        "                   <TKJE>" & RoundEx(-1 * objDelBalanceItems(i).结算金额, 6) & "</TKJE>" & vbCrLf & _
        "                   <JYLSH>" & objDelBalanceItems(i).交易流水号 & "</JYLSH>" & vbCrLf & _
        "                   <JYSM>" & objDelBalanceItems(i).交易说明 & "</JYSM>" & vbCrLf & _
        "                   <DJH>" & "" & "</DJH>" & vbCrLf & _
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
    
    If gobjSquare.objOneCardComLib.zlReturnCheck(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, objItem.卡号, _
        strBalanceIDs, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlReturnCheck=false ")
        Exit Function
    End If
    
    
    If Not objCard.是否退款验卡 Then
        Set objItems_Out = objDelBalanceItems: zlThird_ReturnMoney_IsValied = True:
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True", "退款验卡=false")
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
    '返回:成功,返回=true,否则返回False
    
    strXMLExpend = "<IN><CZLX>" & byt操作类型 & "</CZLX></IN>"
    '操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    If gobjSquare.objOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, True, True, bln退现, True, Nothing, False, False, strXmlIn) = False Then
            
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlBrushCard=false")
        Exit Function
    End If
    
    Set objItems_Out = objDelBalanceItems:
    zlThird_ReturnMoney_IsValied = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True", dblMoney, str卡号, strXMLExpend)
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function
Public Function zlThird_ReturnMultiMoney(ByVal objPati As clsPatientInfo, ByVal objCard As Card, objBalanceInfor As clsBalanceInfo, _
    ByRef cllBillPro As Collection, objCurDelBalanceItems As clsBalanceItems, ByRef objBalanceItems_out As clsBalanceItems, Optional ByVal bln异常作废 As Boolean, _
    Optional blnSaveed_Out As Boolean, Optional blnCurInput As Boolean, Optional blnChangeMoney_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一次按多笔交易返回三方接口
    '入参:dblMoney-本次结算金额
    '     cllBillPro-单据过程(执行完后清空,以便调用下次接口时重复执行)
    '     objBalanceInfor-原结算方式信息集
    '     objcurDelBalanceItems-当前退费的结算项
    '     bln异常作废-异常作废时调用:true-异常作废;false-正常作废操作
    '     cllBillPro-需执行的相关SQL
    '     blnCurInput-是否当前输入项
    '出参:
    '     objBalanceItems_Out-返回已经退费成功的项目
    '     blnSaveed_Out-是否发生数据保存
    '     blnChangeMoney_out-金额发生变化时，返回=true,否则返回False
    '返回:执行成功,返回=true,否则返回False

    '编制:刘兴洪
    '日期:2018-04-16 19:52:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String, objItem As clsBalanceItem, objItemsTemp  As clsBalanceItems, objItemTemp As clsBalanceItem, objCurItem As clsBalanceItem
    Dim i As Long, strInXML As String, strOutXML As String, strExpend As String, intStatus  As Integer, strErrMsg As String
    Dim str结算方式 As String, rsBalance As ADODB.Recordset, rsExpend As ADODB.Recordset
    Dim cllPro As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim blnTrans As Boolean, blnNotBegin As Boolean, blnCallInterfaced As Boolean '已经调用的接口
    Dim strSQL As String, dblMoney As Double, strCardNo As String
      
    Dim str关联交易ID As String
    Dim blnCheck As Boolean, dblTotal As Double
    
    On Error GoTo errHandle
    mstrFunName = "zlThird_ReturnMultiMoney"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", bln异常作废, blnCurInput)
    
    blnSaveed_Out = False: blnChangeMoney_Out = False
    Set objBalanceItems_out = Nothing
    If objCurDelBalanceItems Is Nothing Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objCurDelBalanceItems Is Nothing")
        Exit Function
    End If
    
    If objCurDelBalanceItems.Count = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objCurDelBalanceItems.Count = 0")
        Exit Function
    End If

    
    '先检查数据的合法性
    blnCheck = True
    If objCurDelBalanceItems(1).是否保存 Then
        If Not zlThird_IsSwapIsSucces(objCurDelBalanceItems, intStatus, strErrMsg) Then
           If intStatus = 1 Then
                If strErrMsg <> "" Then
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行退款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    
                    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 正在进行退款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg)
                Else
                    MsgBox "注意:" & vbCrLf & _
                    "   " & objCard.名称 & " 正在进行退款处理中，请稍后再试", vbInformation + vbOKOnly, gstrSysName
                    
                    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 正在进行退款处理中，请稍后再试")
                End If
                Exit Function
           End If
        Else
            blnCheck = False
        End If
    End If
    
    If blnCheck Then
        For Each objItem In objCurDelBalanceItems
            If InStr(str关联交易ID & ",", "," & objItem.关联交易ID & ",") = 0 Then
                Set objItemsTemp = New clsBalanceItems
                For Each objItemTemp In objCurDelBalanceItems
                    If objItem.关联交易ID = objItemTemp.关联交易ID Then
                        objItemsTemp.AddItem objItemTemp
                        objItemsTemp.结算金额 = objItemTemp.结算金额
                    End If
                Next
                If zlThird_ReturnMoney_IsValied(objPati, objCard, 2, objBalanceInfor, objItemsTemp, objItemsTemp, blnCurInput, True) = False Then
                    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlThird_ReturnMoney_IsValied=false")
                    Exit Function
                End If
                str关联交易ID = str关联交易ID & "," & objItem.关联交易ID
            End If
        Next
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    strInXML = "": strXML = ""
    For Each objItem In objCurDelBalanceItems
        
        strXML = strXML & "   <JS>" & vbCrLf
        strXML = strXML & "        <TKFS>" & objItem.结算方式 & "</TKFS>" & vbCrLf     '  //10.35.110后新增,退款方式,即结算方式.名称
        strXML = strXML & "        <KH>" & objItem.卡号 & "</KH>" & vbCrLf
        strXML = strXML & "        <JYLSH>" & objItem.交易流水号 & "</JYLSH>" & vbCrLf
        strXML = strXML & "        <JYSM>" & objItem.交易说明 & "</JYSM>" & vbCrLf
        strXML = strXML & "        <ZFJE>" & RoundEx(-1 * objItem.结算金额, 6) & "</ZFJE>" & vbCrLf
        strXML = strXML & "        <JSLX>" & objCurDelBalanceItems.收费类型 & "</JSLX>" & vbCrLf   '1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        strXML = strXML & "        <GLJYID>" & objItem.关联交易ID & "</GLJYID>" & vbCrLf
        strXML = strXML & "        <DJH></DJH>" & vbCrLf   '  //10.35.110新增:单据号,多单据收费时有用 ，存储在"医保结算明细"表中,主要是分单据保存
        strXML = strXML & "        <ID>" & IIf(objItem.是否预交, objItem.预交ID, objItem.结算IDs) & "</ID>" & vbCrLf
        strXML = strXML & "   </JS>" & vbCrLf
        
        
        dblTotal = RoundEx(dblTotal + objItem.结算金额, 6)
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = objItem.结算方式 & "|" & objItem.结算金额
        str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算号码) = "", " ", Trim(objItem.结算号码))
        str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算摘要) = "", " ", Trim(objItem.结算摘要))
        
        If objBalanceInfor.冲销ID = 0 Then
            '调用之前,先处理数据
            'Zl_病人结帐结算_Modify_S
            strSQL = "Zl_病人结帐结算_Modify_S("
            '  操作类型_In     Number,
            '  --操作类型_In:
            '  --   0-普通收费方式:
            '  --   1.三方卡结算:
            '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
            '  --     ②退支票额_In:传入零
            '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
            strSQL = strSQL & "1,"
            '  病人id_In       病人结帐记录.病人id%Type,
            strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
            '  结帐id_In       病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
            '  结算方式_In     Varchar2,
            strSQL = strSQL & "'" & str结算方式 & "',"
            '  冲预交_In       病人预交记录.冲预交%Type := Null,
            strSQL = strSQL & "NULL,"
            '  退支票额_In     病人预交记录.冲预交%Type := Null,
            strSQL = strSQL & "NULL,"
            '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
            strSQL = strSQL & "" & objItem.卡类别ID & ","
            '  卡号_In         病人预交记录.卡号%Type := Null,
            strSQL = strSQL & "'" & objItem.卡号 & "',"
            '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
            strSQL = strSQL & "'" & objItem.交易流水号 & "',"
            '  交易说明_In     病人预交记录.交易说明%Type := Null,
            strSQL = strSQL & "'" & objItem.交易说明 & "',"
            '  缴款_In         病人预交记录.缴款%Type := Null,
            strSQL = strSQL & "" & "Null" & ","
            '  找补_In         病人预交记录.找补%Type := Null,
            strSQL = strSQL & "" & "Null" & ","
            '  误差金额_In     门诊费用记录.实收金额%Type := Null,
            strSQL = strSQL & "NULL,"
            '  结帐类型_In     Number := 2,
            strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
            '  缺省结算方式_In 结算方式.名称%Type := Null,
            strSQL = strSQL & "NULL,"
            '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
            strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '    冲预交病人ids_In Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '    当前科室id_In    病人预交记录.科室id%Type := Null,
            strSQL = strSQL & ZVal(objPati.当前科室id) & ","
            '  完成结算_In Number:=0
            strSQL = strSQL & "0,"
            '  校对标志_In      Number := 2,
            strSQL = strSQL & "" & IIf(mbln结算异步控制, 1, 2) & ","
            '  预交id_In        病人预交记录.Id%Type := Null,
            strSQL = strSQL & "NULL,"
            '  关联交易id_In    病人预交记录.Id%Type := Null,
            strSQL = strSQL & IIf(Val(objItem.关联交易ID) = 0, "NULL", Val(objItem.关联交易ID)) & ","
            '  清除原交易_In Number:=0
            strSQL = strSQL & "1" & ","
            '附加标志_In    病人预交记录.Id%Type := Null
            '对于结帐业务(三方卡退款传入)：NULl or 0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
            strSQL = strSQL & "" & 2 & ")"
            
        Else
            'Zl_病人结帐作废_Modify
            strSQL = "Zl_病人结帐作废_Modify("
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
            strSQL = strSQL & "" & 2 & ","
            '  病人id_In     病人结帐记录.病人id%Type,
            strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
            '  冲销id_In     病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & objBalanceInfor.冲销ID & ","
            '  结算方式_In   Varchar2,
            strSQL = strSQL & "'" & str结算方式 & "',"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSQL = strSQL & objCard.接口序号 & ","
            '  卡号_In       病人预交记录.卡号%Type := Null,
            strSQL = strSQL & "'" & objItem.卡号 & "',"
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSQL = strSQL & "'" & objItem.交易流水号 & "',"
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            strSQL = strSQL & "'" & objItem.交易说明 & "',"
            '  缴款_In       病人预交记录.缴款%Type := Null,
            strSQL = strSQL & "NULL,"
            '  找补_In       病人预交记录.找补%Type := Null,
            strSQL = strSQL & "NULL,"
            '  误差金额_In   病人预交记录.冲预交%Type := Null,
            strSQL = strSQL & "NULL,"
            '  预交金额_In   病人预交记录.冲预交%Type := Null,
            strSQL = strSQL & "NULL,"
            '  完成作废_In Number:=0
            strSQL = strSQL & "0,"
            '  医保校对标志_In  Number := 0,
            strSQL = strSQL & "0,"
            '  关联交易id_In 病人预交记录.Id%Type := Null,
            strSQL = strSQL & "" & IIf(objItem.关联交易ID = 0, "NULL", objItem.关联交易ID) & ","
            '  清除原交易_In Number:=0
            strSQL = strSQL & "1" & "）"
        End If
        If Not bln异常作废 Then zlAddArray cllPro, strSQL   '异常作废已经处理
    Next
    
    Set objItem = objCurDelBalanceItems(1)
    strInXML = strInXML & "<JSLIST>" & vbCrLf
    strInXML = strInXML & strXML & vbCrLf
    strInXML = strInXML & "</JSLIST>" & vbCrLf

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
    '                    <TKFS>退款方式</TKFS>  //10.35.110后新增,退款方式,即结算方式.名称
    '                    <KH>卡号</KH>
    '                    <JYLSH>交易流水号</JYLSH>
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
    '                <GLJYID>关联交易ID</GLJYID>  //10.35.110新增:原始交易关联ID(病人预交记录.关联交易ID)
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
    blnTrans = True
    Call zlExecuteProcedureArrAy(cllPro, mfrmMain.Caption, Not mbln结算异步控制, , "结帐", "病人结帐调试日志")
    If mbln结算异步控制 Then blnTrans = False: blnSaveed_Out = True
    
    Dim lng结帐ID As Long
    lng结帐ID = IIf(objBalanceInfor.冲销ID = 0, objBalanceInfor.结帐ID, objBalanceInfor.冲销ID)
    blnCallInterfaced = True
    If gobjSquare.objOneCardComLib.zlReturnMultiMoney(mfrmMain, mlngModule, objItem.卡类别ID, objItem.消费卡, strInXML, lng结帐ID, strOutXML, strExpend, True, rsBalance, rsExpend, intStatus, strErrMsg) = False Then
         If blnTrans Then gcnOracle.RollbackTrans
         
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlReturnMultiMoney=false", strOutXML, strExpend, intStatus, strErrMsg)
         If strErrMsg <> "" Then
            MsgBox "注意:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
         End If
         
         If intStatus = 1 Then '交易正在处理中
            If blnSaveed_Out Then
                For Each objItem In objCurDelBalanceItems
                    objItem.是否保存 = True
                    objItem.是否允许编辑 = False
                    objItem.是否允许删除 = False
                    objItem.结帐ID = objBalanceInfor.结帐ID
                    objItem.结帐时间 = objBalanceInfor.结帐时间
                    objItem.是否结算 = False
                Next
            End If
            Set objBalanceItems_out = objCurDelBalanceItems
            blnCallInterfaced = False
            Exit Function
         End If
         
         If mbln结算异步控制 Then
             For Each objItem In objCurDelBalanceItems
                   Call zlDelBalanceItemFromItemObject(objItem)
             Next
         End If
         blnCallInterfaced = False
         Exit Function
    End If
     
    Set objBalanceItems_out = objCurDelBalanceItems
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    str关联交易ID = ""
    Dim lngID As Long
    
    '更新数据
    If Not rsBalance Is Nothing Then
        Set objBalanceItems_out = New clsBalanceItems
        
        dblMoney = 0
        With rsBalance
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
            
                Set objItemTemp = Nothing
                For Each objCurItem In objCurDelBalanceItems
                    lngID = IIf(objCurItem.是否预交, objCurItem.预交ID, Val(objCurItem.结算IDs))
                    If (lngID = Val(NVL(!ID)) And Val(NVL(!ID)) <> 0) Or (Val(NVL(!ID)) = 0 And Val(NVL(!卡类别ID)) = objCurItem.卡类别ID And NVL(!卡号) = objCurItem.卡号) Then
                         Exit For
                    End If
                Next
                If objCurItem Is Nothing Then
                    Set objCurItem = objItem
                    If objCurItem Is Nothing Then
                        Set objCurItem = New clsBalanceItem
                    End If
                    
                End If
                Set objItemTemp = New clsBalanceItem
                
                With objItemTemp
                    .卡类别ID = IIf(Val(NVL(rsBalance!是否普通结算)) = 0, Val(NVL(rsBalance!卡类别ID)), 0)
                    .消费卡 = False
                    .消费卡ID = 0
                    .卡号 = IIf(NVL(rsBalance!卡号) = "", objItem.卡号, NVL(rsBalance!卡号))
                    If strCardNo = "" Then strCardNo = .卡号
                    .交易流水号 = NVL(rsBalance!原交易流水号)
                    If .交易流水号 = "" Then .交易流水号 = objCurItem.交易流水号
                    .交易说明 = NVL(rsBalance!原交易说明)
                    If .交易说明 = "" Then .交易说明 = objCurItem.交易说明
                    .退款交易流水号 = NVL(rsBalance!交易流水号)
                    .退款交易说明 = NVL(rsBalance!交易说明)
                    
                    .结算方式 = NVL(rsBalance!结算方式)
                    If .结算方式 = "" Then .结算方式 = objCurItem.结算方式
                    .结帐ID = objCurItem.结帐ID
                    .结算IDs = objCurItem.结算IDs
                    .结算号码 = NVL(rsBalance!结算号码)
                    .结算摘要 = NVL(rsBalance!结算摘要)
                    .结算金额 = IIf(Val(NVL(rsBalance!交易金额)) = 0, objCurItem.结算金额, -1 * Val(NVL(rsBalance!交易金额)))
                    .密码 = objItem.密码
                    .是否密文 = objCurItem.是否密文
                    .限制类别 = objCurItem.限制类别
                    .帐户余额 = objCurItem.帐户余额
                    .结帐时间 = objBalanceInfor.结帐时间
                    
                    .是否退款分交易 = False
                    .是否预交 = objCurItem.是否预交
                    .是否允许编辑 = False
                    .是否允许删除 = False
                    .是否允许退现 = False
                    .是否转帐 = objCurItem.是否转帐
                    .是否保存 = True
                    .是否结算 = True
                    .冲销ID = objBalanceInfor.冲销ID
                    .关联交易ID = Val(NVL(rsBalance!关联交易ID))
                    .预交ID = Val(NVL(rsBalance!ID))
                    If .预交ID = 0 Then .预交ID = objCurItem.预交ID
                    If .关联交易ID = 0 Then .关联交易ID = objCurItem.关联交易ID
                    Set .objCard = objCurItem.objCard
                    dblMoney = RoundEx(dblMoney + .结算金额, 6)
                End With
                
                '结算方式|结算金额|结算号码|结算摘要||..
                str结算方式 = objItemTemp.结算方式 & "|" & objItemTemp.结算金额
                str结算方式 = str结算方式 & "|" & IIf(Trim(objItemTemp.结算号码) = "", " ", Trim(objItemTemp.结算号码))
                str结算方式 = str结算方式 & "|" & IIf(Trim(objItemTemp.结算摘要) = "", " ", Trim(objItemTemp.结算摘要))
                
                objBalanceItems_out.AddItem objItemTemp
                If objBalanceInfor.冲销ID = 0 Then
                    '调用之前,先处理数据
                    'Zl_病人结帐结算_Modify_S
                    strSQL = "Zl_病人结帐结算_Modify_S("
                    '  操作类型_In     Number,
                    '  --操作类型_In:
                    '  --   0-普通收费方式:
                    '  --   1.三方卡结算:
                    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
                    '  --     ②退支票额_In:传入零
                    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
                    strSQL = strSQL & "1,"
                    '  病人id_In       病人结帐记录.病人id%Type,
                    strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
                    '  结帐id_In       病人预交记录.结帐id%Type,
                    strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
                    '  结算方式_In     Varchar2,
                    strSQL = strSQL & "'" & str结算方式 & "',"
                    '  冲预交_In       病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  退支票额_In     病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
                    strSQL = strSQL & "" & objItemTemp.卡类别ID & ","
                    '  卡号_In         病人预交记录.卡号%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.卡号 & "',"
                    '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.交易流水号 & "',"
                    '  交易说明_In     病人预交记录.交易说明%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.交易说明 & "',"
                    '  缴款_In         病人预交记录.缴款%Type := Null,
                    strSQL = strSQL & "" & "Null" & ","
                    '  找补_In         病人预交记录.找补%Type := Null,
                    strSQL = strSQL & "" & "Null" & ","
                    '  误差金额_In     门诊费用记录.实收金额%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  结帐类型_In     Number := 2,
                    strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
                    '  缺省结算方式_In 结算方式.名称%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
                    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                    '    冲预交病人ids_In Varchar2 := Null,
                    strSQL = strSQL & "NULL,"
                    '    当前科室id_In    病人预交记录.科室id%Type := Null,
                    strSQL = strSQL & ZVal(objPati.当前科室id) & ","
                    '  完成结算_In Number:=0
                    strSQL = strSQL & "0,"
                    ' 校对标志_In      Number := 2,
                    strSQL = strSQL & "2,"
                    '  预交id_In        病人预交记录.Id%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  关联交易id_In 病人预交记录.Id%Type := Null,
                    strSQL = strSQL & "" & IIf(objItemTemp.关联交易ID = 0, "NULL", objItemTemp.关联交易ID) & ","
                    '  清除原交易_In Number:=0
                    strSQL = strSQL & "" & "1" & ","
                    '附加标志_In    病人预交记录.Id%Type := Null
                    '对于结帐业务(三方卡退款传入)：NULl or 0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
                    strSQL = strSQL & "" & 2 & ")"
                    zlAddArray cllUpdate, strSQL
                    If objItem.是否预交 Then Call zlGetSaveThirdSwapDelSQLFromBalanceItem(objItemTemp, True, cllUpdate, 0)
                    
                Else
                    
                    'Zl_病人结帐作废_Modify
                    strSQL = "Zl_病人结帐作废_Modify("
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
                    strSQL = strSQL & "" & 2 & ","
                    '  病人id_In     病人结帐记录.病人id%Type,
                    strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
                    '  冲销id_In     病人预交记录.结帐id%Type,
                    strSQL = strSQL & "" & objBalanceInfor.冲销ID & ","
                    '  结算方式_In   Varchar2,
                    strSQL = strSQL & "'" & str结算方式 & "',"
                    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
                    strSQL = strSQL & objCard.接口序号 & ","
                    '  卡号_In       病人预交记录.卡号%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.卡号 & "',"
                    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.交易流水号 & "',"
                    '  交易说明_In   病人预交记录.交易说明%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.交易说明 & "',"
                    '  缴款_In       病人预交记录.缴款%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  找补_In       病人预交记录.找补%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  误差金额_In   病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  预交金额_In   病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  完成作废_In Number:=0
                    strSQL = strSQL & "0,"
                    '  医保校对标志_In  Number := 0,
                    strSQL = strSQL & "0,"
                    '  关联交易id_In 病人预交记录.Id%Type := Null,
                    strSQL = strSQL & "" & IIf(objItemTemp.关联交易ID = 0, "NULL", objItemTemp.关联交易ID) & ","
                    '  清除原交易_In Number:=0
                    strSQL = strSQL & "1" & ")"
                    zlAddArray cllUpdate, strSQL
                End If
                
                .MoveNext
            Loop
            objBalanceItems_out.结算金额 = dblMoney
        End With
        blnChangeMoney_Out = dblTotal <> dblMoney
    Else
         Set objBalanceItems_out = objCurDelBalanceItems
         For Each objItemTemp In objBalanceItems_out
            objItemTemp.结帐ID = objItem.结帐ID
            objItemTemp.结算IDs = objItem.结算IDs
            objItemTemp.结帐时间 = objBalanceInfor.结帐时间
            objItemTemp.是否退款分交易 = False
            objItemTemp.是否预交 = objItem.是否预交
            objItemTemp.是否允许编辑 = False
            objItemTemp.是否允许删除 = False
            objItemTemp.是否允许退现 = False
            objItemTemp.是否转帐 = False
            objItemTemp.是否保存 = True
            objItemTemp.是否结算 = True
            objItemTemp.冲销ID = objBalanceInfor.冲销ID
         Next
    End If
    
    blnNotBegin = blnTrans: blnTrans = True: blnCallInterfaced = False
    
    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, False, blnNotBegin, "结帐", "病人结帐调试日志"
    
    blnSaveed_Out = True
    Set cllBillPro = New Collection
    
    blnTrans = False
    zlThird_ReturnMultiMoney = True
    
    Err = 0: On Error GoTo ErrOther:
    
    Call zlGetThreeSwapExpendSQLToCollect(objCard, IIf(objBalanceInfor.冲销ID = 0, objBalanceInfor.结帐ID, objBalanceInfor.冲销ID), strCardNo, rsExpend, cllThreeSwap, False)
    blnTrans = True
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption & "-执行三方接口", , , "结帐", "病人结帐调试日志"
    blnTrans = False
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If blnTrans And Not blnCallInterfaced Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
    Exit Function
ErrOther:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "（ErrOther）Err=" & mstrErrInfor)
End Function


Public Function zlGetBalanceCancelSaveSQLFromItems(ByVal objPati As clsPatientInfo, objBalanceInfor As clsBalanceInfo, _
    ByVal objItems As clsBalanceItems, ByRef cllPro_Out As Collection, Optional ByVal int校对标志 As Integer = -1, Optional blnSign As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算对象集，获取对应的作废SQL更新的集合
    '入参:objItems-结算信息集
    '      int校对标志-： -1表示根据是否异步分离来处理;>=0表示，按该值填写
    '      blnSign-是否取反
    '出参:
    '返回:成功返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-30 17:17:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str结算方式 As String
    Dim objItem As clsBalanceItem, lng预交ID As Long, lng关联交易ID As Long
    Dim i As Long
    
    
    On Error GoTo errHandle
    
    If cllPro_Out Is Nothing Then Set cllPro_Out = New Collection
    If objItems Is Nothing Then Exit Function
    If objItems.Count = 0 Then Exit Function
    
    
'    If objItems(1).是否保存 Then
'        lng预交ID = objItems(1).预交ID
'    End If
'
    lng关联交易ID = objItems(1).关联交易ID
    
    If lng预交ID = 0 Then lng预交ID = zlDatabase.GetNextId("病人预交记录")
    If lng关联交易ID = 0 Then lng关联交易ID = lng预交ID
    
    i = 0
    For Each objItem In objItems
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = objItem.结算方式 & "|" & RoundEx(IIf(blnSign, -1, 1) * objItem.结算金额, 6)
        str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算号码) = "", " ", Trim(objItem.结算号码))
        str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算摘要) = "", " ", Trim(objItem.结算摘要))
        'Zl_病人结帐作废_Modify
        strSQL = "Zl_病人结帐作废_Modify("
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
        strSQL = strSQL & "" & Decode(objItem.结算类型, 3, 2, 4, 0, 0) & "," '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        '  病人id_In     病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & objBalanceInfor.冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "" & ZVal(objItem.卡类别ID) & ","
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & objItem.卡号 & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & objItem.交易流水号 & "',"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & objItem.交易说明 & "',"
        '  缴款_In       病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "NULL,"
        '  找补_In       病人预交记录.找补%Type := Null,
        strSQL = strSQL & "NULL,"
        '  误差金额_In   病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  预交金额_In   病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  收款时间_In      病人预交记录.收款时间%Type := Null,
        strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  冲预交病人ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  完成作废_In Number:=0
        strSQL = strSQL & "0,"
        '  校对标志_In  Number := 0,
        If int校对标志 < 0 Then
            strSQL = strSQL & IIf(mbln结算异步控制, 1, 2) & ","
        Else
            strSQL = strSQL & int校对标志 & ","
        End If
        '  关联交易id_In 病人预交记录.Id%Type := Null,
        strSQL = strSQL & "" & ZVal(lng关联交易ID) & ","
        '  清除原交易_In Number:=0
        strSQL = strSQL & IIf(i = 0, "1", "0") & ","
        '  预交id_In        病人预交记录.Id%Type := Null
        strSQL = strSQL & ZVal(lng预交ID) & ")"
        zlAddArray cllPro_Out, strSQL
        objItem.关联交易ID = lng关联交易ID
        lng预交ID = 0
        i = i + 1
     Next
    zlGetBalanceCancelSaveSQLFromItems = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlThird_ReturnMoney(ByVal objPati As clsPatientInfo, ByVal objCard As Card, objBalanceInfor As clsBalanceInfo, objCurDelBalanceItems As clsBalanceItems, _
    ByRef cllBillPro As Collection, Optional ByVal bln异常作废 As Boolean, Optional objBalanceItems_out As clsBalanceItems, _
    Optional blnSaveed_Out As Boolean, Optional blnCurInput As Boolean, Optional blnChangeMoney_Out As Boolean, Optional blnOlnyModifyData As Boolean, Optional ByVal blnCallCheckedIterface As Boolean = True) As Boolean
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
    '     blnSaveed_Out-是否发生数据保存
    '返回:执行成功,返回=true,否则返回False
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
    Dim blnCheck As Boolean, str关联交易ID As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strExpendXML As String, strTemp As String
    Dim objItems As clsBalanceItems
    
    On Error GoTo ErrHand
    
    
    mstrFunName = "zlThird_ReturnMoney"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", objCard.接口序号, objCard.结算方式)
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    Set objBalanceItems_out = Nothing
    
    blnSaveed_Out = False
    '非一卡通支付,直接返回
    If objCard.接口序号 <= 0 Or objCard.消费卡 Then
        zlThird_ReturnMoney = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", objCard.接口序号, IIf(objCard.消费卡, "true", "false"))
        Exit Function
    End If
    If objCurDelBalanceItems Is Nothing Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objCurDelBalanceItems Is Nothing ")
        Exit Function
    End If
    If objCurDelBalanceItems.Count = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", " objCurDelBalanceItems.Count = 0")
        Exit Function
    End If
    
    '先检查数据的合法性
    If Not blnCallCheckedIterface Then
        
        blnCheck = True
        If objCurDelBalanceItems(1).是否保存 Then
            If Not zlThird_IsSwapIsSucces(objCurDelBalanceItems, intStatus, strErrMsg) Then
               If intStatus = 1 Then
                    If strErrMsg <> "" Then
                        MsgBox "注意:" & vbCrLf & _
                        "   " & objCard.名称 & " 正在进行退款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Else
                        MsgBox "注意:" & vbCrLf & _
                        "   " & objCard.名称 & " 正在进行退款处理中，请稍后再试" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    End If
                    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 正在进行退款处理,请稍后再试,详细信息如下:" & vbCrLf & strErrMsg)
                    Exit Function
               End If
            Else
                blnCheck = False
            End If
        End If
        If blnCheck Then
            ' 0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡4
            If zlThird_ReturnMoney_IsValied(objPati, objCard, 2, objBalanceInfor, objCurDelBalanceItems, objItems, blnCurInput, True) = False Then
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlThird_ReturnMoney_IsValied=false")
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
    lng结帐ID = IIf(objBalanceInfor.冲销ID <> 0, objBalanceInfor.冲销ID, objBalanceInfor.结帐ID)
    strExpendXML = "": strTemp = ""
    For Each objItem In objCurDelBalanceItems
        
        strTemp = strTemp & "        <TK>" & vbCrLf
        strTemp = strTemp & "           <TKFS>" & objItem.结算方式 & "</TKFS>" & vbCrLf
        strTemp = strTemp & "           <TKJE>" & RoundEx(-1 * objItem.结算金额, 6) & "</TKJE>" & vbCrLf
        strTemp = strTemp & "           <JYLSH>" & objItem.交易流水号 & "</JYLSH>" & vbCrLf
        strTemp = strTemp & "           <JYSM>" & objItem.交易说明 & "</JYSM>" & vbCrLf
        strTemp = strTemp & "           <DJH>" & "" & "</DJH>" & vbCrLf
        strTemp = strTemp & "        </TK>"
        dblMoney = dblMoney + -1 * objItem.结算金额
        strCardNo = objItem.卡号
        strSwapNO = objItem.交易流水号
        strSwapMemo = objItem.交易说明
        
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = objItem.结算方式 & "|" & objItem.结算金额
        str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算号码) = "", " ", Trim(objItem.结算号码))
        str结算方式 = str结算方式 & "|" & IIf(Trim(objItem.结算摘要) = "", " ", Trim(objItem.结算摘要))
        
        If objBalanceInfor.冲销ID = 0 Then
            '调用之前,先处理数据
            'Zl_病人结帐结算_Modify_S
            strSQL = "Zl_病人结帐结算_Modify_S("
            '  操作类型_In     Number,
            '  --操作类型_In:
            '  --   0-普通收费方式:
            '  --   1.三方卡结算:
            '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
            '  --     ②退支票额_In:传入零
            '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
            strSQL = strSQL & "1,"
            '  病人id_In       病人结帐记录.病人id%Type,
            strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
            '  结帐id_In       病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
            '  结算方式_In     Varchar2,
            strSQL = strSQL & "'" & str结算方式 & "',"
            '  冲预交_In       病人预交记录.冲预交%Type := Null,
            strSQL = strSQL & "NULL,"
            '  退支票额_In     病人预交记录.冲预交%Type := Null,
            strSQL = strSQL & "NULL,"
            '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
            strSQL = strSQL & "" & objItem.卡类别ID & ","
            '  卡号_In         病人预交记录.卡号%Type := Null,
            strSQL = strSQL & "'" & objItem.卡号 & "',"
            '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
            strSQL = strSQL & "'" & objItem.交易流水号 & "',"
            '  交易说明_In     病人预交记录.交易说明%Type := Null,
            strSQL = strSQL & "'" & objItem.交易说明 & "',"
            '  缴款_In         病人预交记录.缴款%Type := Null,
            strSQL = strSQL & "" & "Null" & ","
            '  找补_In         病人预交记录.找补%Type := Null,
            strSQL = strSQL & "" & "Null" & ","
            '  误差金额_In     门诊费用记录.实收金额%Type := Null,
            strSQL = strSQL & "NULL,"
            '  结帐类型_In     Number := 2,
            strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
            '  缺省结算方式_In 结算方式.名称%Type := Null,
            strSQL = strSQL & "NULL,"
            '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
            strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '    冲预交病人ids_In Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '    当前科室id_In    病人预交记录.科室id%Type := Null,
            strSQL = strSQL & ZVal(objPati.当前科室id) & ","
            '  完成结算_In Number:=0
            strSQL = strSQL & "0,"
            '  校对标志_In      Number := 2,
            strSQL = strSQL & "" & IIf(mbln结算异步控制, 1, 2) & ","
            '  预交id_In        病人预交记录.Id%Type := Null,
            strSQL = strSQL & "NULL,"
            '  关联交易id_In    病人预交记录.Id%Type := Null,
            strSQL = strSQL & IIf(Val(objItem.关联交易ID) = 0, "NULL", Val(objItem.关联交易ID)) & ","
            '  清除原交易_In Number:=0
            strSQL = strSQL & "" & IIf(i = 0, 1, 0) & ","
            '附加标志_In    病人预交记录.Id%Type := Null
            '对于结帐业务(三方卡退款传入)：NULl or 0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
            strSQL = strSQL & "" & 1 & ")"
            If Not blnOlnyModifyData Then zlAddArray cllPro, strSQL
        Else
            'Zl_病人结帐作废_Modify
            strSQL = "Zl_病人结帐作废_Modify("
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
            strSQL = strSQL & "" & 2 & ","
            '  病人id_In     病人结帐记录.病人id%Type,
            strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
            '  冲销id_In     病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & objBalanceInfor.冲销ID & ","
            '  结算方式_In   Varchar2,
            strSQL = strSQL & "'" & str结算方式 & "',"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSQL = strSQL & objCard.接口序号 & ","
            '  卡号_In       病人预交记录.卡号%Type := Null,
            strSQL = strSQL & "'" & objItem.卡号 & "',"
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSQL = strSQL & "'" & objItem.交易流水号 & "',"
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            strSQL = strSQL & "'" & objItem.交易说明 & "',"
            '  缴款_In       病人预交记录.缴款%Type := Null,
            strSQL = strSQL & "NULL,"
            '  找补_In       病人预交记录.找补%Type := Null,
            strSQL = strSQL & "NULL,"
            '  误差金额_In   病人预交记录.冲预交%Type := Null,
            strSQL = strSQL & "NULL,"
            '  预交金额_In   病人预交记录.冲预交%Type := Null,
            strSQL = strSQL & "NULL,"
            '  操作员编号_In    病人预交记录.操作员编号%Type := Null,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  收款时间_In      病人预交记录.收款时间%Type := Null,
            strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '  冲预交病人ids_In Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '  完成作废_In Number:=0
            strSQL = strSQL & "0,"
            '  校对标志_In  Number := 0,
            strSQL = strSQL & IIf(mbln结算异步控制, 1, 2) & ","
            '  关联交易id_In 病人预交记录.Id%Type := Null,
            strSQL = strSQL & "" & IIf(Val(objItem.关联交易ID) = 0, "NULL", objItem.关联交易ID) & ","
            '  清除原交易_In Number:=0
            strSQL = strSQL & IIf(i = 0, "1", "0") & ")"
            If Not blnOlnyModifyData Then zlAddArray cllPro, strSQL    '异常作废已经处理
        End If
        i = i + 1
    Next
    
    dblDelMoney = dblMoney
    Set objItem = objCurDelBalanceItems(1)
    strExpendXML = strExpendXML & "<INPUT>" & vbCrLf
    strExpendXML = strExpendXML & "    <TKLIST>" & vbCrLf
    strExpendXML = strExpendXML & strTemp & vbCrLf
    strExpendXML = strExpendXML & "    </TKLIST>" & vbCrLf
    strExpendXML = strExpendXML & "</INPUT>" & vbCrLf
     
    If dblMoney = 0 Then
        MsgBox "当前退款为零了，不允许再次退款", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "当前退款为零了，不允许再次退款")
        Exit Function
    End If
    
    If Not mbln结算异步控制 And objItem.是否预交 Then
        '预交款退款，需要先保存三方退款信息的是否退款标志
        Call zlGetSaveThirdSwapDelSQLFromBalanceItem(objItem, True, cllPro, 0)
    End If
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, IIf(mbln结算异步控制, False, True), , "结帐", "病人结帐调试日志"
    
    If mbln结算异步控制 Then    '异步处理，需要事先处理校对标志为2(减少接口调用后更新)
        blnTrans = False: blnSaveed_Out = True
        Set cllBeforePro = New Collection
        Call zlAddUpdateSwapSQL(False, lng结帐ID, objItem.卡类别ID, False, "", "", "", cllBeforePro, 2, objItem.关联交易ID, 1)   '将其更新为2
       
        If objItem.是否预交 Then
           Call zlGetSaveThirdSwapDelSQLFromBalanceItem(objItem, True, cllBeforePro, 0)
        End If
        
       blnTrans = True
       zlExecuteProcedureArrAy cllBeforePro, mfrmMain.Caption, True, , "结帐", "病人结帐调试日志"
 
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
    
    If objItem.是否预交 Then
         strBalanceIDs = "1|" & objItem.预交ID
    Else
        strBalanceIDs = "2|" & objBalanceInfor.结帐ID
    End If
    
    strSwapExtendInfor = "2|" & IIf(objBalanceInfor.冲销ID <> 0, objBalanceInfor.冲销ID, objBalanceInfor.结帐ID)
    If gobjSquare.objOneCardComLib.zlReturnMoney(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, strCardNo, _
        strBalanceIDs, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strExpendXML, True, rsBalance, rsExpend, intStatus, strErrMsg) = False Then
        
        If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlReturnMoney=false", "intStatus=" & intStatus)
        
        If strErrMsg <> "" Then
            MsgBox "注意:" & vbCrLf & "    当前退款失败，错误原因如下:" & vbCrLf & strErrMsg, vbInformation, gstrSysName
        End If
        
        If intStatus = 1 Or objBalanceInfor.冲销ID <> 0 Then    '冲销的，不能直接删除当前结算，可能后期还有其他操作
            If blnSaveed_Out Then
                For Each objItem In objCurDelBalanceItems
                    objItem.是否保存 = True
                    objItem.是否允许编辑 = False
                    objItem.是否允许删除 = objBalanceInfor.冲销ID = 0
                    objItem.结帐ID = objBalanceInfor.结帐ID
                    objItem.冲销ID = objBalanceInfor.冲销ID
                    
                    objItem.结帐时间 = objBalanceInfor.结帐时间
                    objItem.是否结算 = False
                Next
            End If
            Set objBalanceItems_out = objCurDelBalanceItems
            Exit Function
        End If
        
        If mbln结算异步控制 Then
        
             Set cllUpdate = New Collection
             For Each objItem In objCurDelBalanceItems
                Call zlDelBalanceItemFromItemObject(objItem, cllUpdate, False)
             Next
             Call zlExecuteProcedureArrAy(cllUpdate, mfrmMain.Caption, False, blnTrans, "结帐", "病人结帐调试日志")
             blnTrans = False
        End If
        Set objBalanceItems_out = objCurDelBalanceItems
        Exit Function
    End If
    
    
    Set objItem = objCurDelBalanceItems(1)
    
    Set objBalanceItems_out = objCurDelBalanceItems
    '更新数据
    lngCount = 0
    If Not rsBalance Is Nothing Then
        Set objBalanceItems_out = New clsBalanceItems
        dblMoney = 0
        If rsBalance.RecordCount <> 0 Then rsBalance.MoveFirst
        With rsBalance
            Do While Not .EOF
                Set objItemTemp = New clsBalanceItem
                With objItemTemp
                    .行号 = objItem.行号
                    .关联交易ID = objItem.关联交易ID
                    .预交ID = objItem.预交ID
                    
                    .卡类别ID = IIf(Val(NVL(rsBalance!是否普通结算)) = 0, Val(NVL(rsBalance!卡类别ID)), 0)
                    .消费卡 = False
                    .消费卡ID = 0
                    .卡号 = IIf(NVL(rsBalance!卡号) = "", objItem.卡号, NVL(rsBalance!卡号))
                    .交易流水号 = NVL(rsBalance!原交易流水号)
                    .交易说明 = NVL(rsBalance!原交易说明)
                    .退款交易流水号 = NVL(rsBalance!交易流水号)
                    .退款交易说明 = NVL(rsBalance!交易说明)
                    .结算方式 = NVL(rsBalance!结算方式)
                    .结算金额 = -1 * Val(NVL(rsBalance!交易金额))
                    .密码 = objItem.密码
                    .是否密文 = objItem.是否密文
                    .限制类别 = objItem.限制类别
                    .帐户余额 = objItem.帐户余额
                    .结算号码 = NVL(rsBalance!结算号码)
                    .结算摘要 = NVL(rsBalance!结算摘要)
                    .结帐时间 = objBalanceInfor.结帐时间
                    .结帐ID = objBalanceInfor.结帐ID
                    .结算类型 = objItem.结算类型
                    .冲销ID = objBalanceInfor.冲销ID
                    .校对标志 = 2
                    If .卡号 <> "" Then strCardNo = .卡号
                    If strSwapNO = "" And .交易流水号 <> "" Then
                        strSwapNO = .交易流水号
                        If strSwapMemo = "" Then strSwapNO = .交易流水号
                    End If
                    
                    dblMoney = RoundEx(dblMoney + .结算金额, 6)
                End With
                '结算方式|结算金额|结算号码|结算摘要||..
                str结算方式 = objItemTemp.结算方式 & "|" & objItemTemp.结算金额
                str结算方式 = str结算方式 & "|" & IIf(Trim(objItemTemp.结算号码) = "", " ", Trim(objItemTemp.结算号码))
                str结算方式 = str结算方式 & "|" & IIf(Trim(objItemTemp.结算摘要) = "", " ", Trim(objItemTemp.结算摘要))
                objBalanceItems_out.AddItem objItemTemp
                
                If objBalanceInfor.冲销ID = 0 Then
                        
                    '调用之前,先处理数据
                    'Zl_病人结帐结算_Modify_S
                    strSQL = "Zl_病人结帐结算_Modify_S("
                    '  操作类型_In     Number,
                    '  --操作类型_In:
                    '  --   0-普通收费方式:
                    '  --   1.三方卡结算:
                    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
                    '  --     ②退支票额_In:传入零
                    '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
                    strSQL = strSQL & "1,"
                    '  病人id_In       病人结帐记录.病人id%Type,
                    strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
                    '  结帐id_In       病人预交记录.结帐id%Type,
                    strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
                    '  结算方式_In     Varchar2,
                    strSQL = strSQL & "'" & str结算方式 & "',"
                    '  冲预交_In       病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  退支票额_In     病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
                    strSQL = strSQL & "" & objItemTemp.卡类别ID & ","
                    '  卡号_In         病人预交记录.卡号%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.卡号 & "',"
                    '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.交易流水号 & "',"
                    '  交易说明_In     病人预交记录.交易说明%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.交易说明 & "',"
                    '  缴款_In         病人预交记录.缴款%Type := Null,
                    strSQL = strSQL & "" & "Null" & ","
                    '  找补_In         病人预交记录.找补%Type := Null,
                    strSQL = strSQL & "" & "Null" & ","
                    '  误差金额_In     门诊费用记录.实收金额%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  结帐类型_In     Number := 2,
                    strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
                    '  缺省结算方式_In 结算方式.名称%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
                    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                    '    冲预交病人ids_In Varchar2 := Null,
                    strSQL = strSQL & "NULL,"
                    '    当前科室id_In    病人预交记录.科室id%Type := Null,
                    strSQL = strSQL & ZVal(objPati.当前科室id) & ","
                    '  完成结算_In Number:=0
                    strSQL = strSQL & "0,"
                    '  校对标志_In      Number := 2,
                    strSQL = strSQL & "2,"
                    '  预交id_In        病人预交记录.Id%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  关联交易id_In    病人预交记录.Id%Type := Null,
                    strSQL = strSQL & IIf(Val(objItem.关联交易ID) = 0, "NULL", Val(objItem.关联交易ID)) & ","
                    '  清除原交易_In Number:=0
                    strSQL = strSQL & "" & IIf(lngCount = 0, "1", "0") & ","
                    '附加标志_In    病人预交记录.Id%Type := Null
                    '对于结帐业务(三方卡退款传入)：NULl or 0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
                    strSQL = strSQL & "" & 1 & ")"
                    zlAddArray cllUpdate, strSQL
                    If objItem.是否预交 And (mbln结算异步控制 Or objItem.退款交易流水号 <> objItemTemp.退款交易流水号 Or objItem.退款交易说明 <> objItem.退款交易说明) Then
                        Call zlGetSaveThirdSwapDelSQLFromBalanceItem(objItemTemp, True, cllUpdate, 0)
                    End If
                    
                Else
                    'Zl_病人结帐作废_Modify
                    strSQL = "Zl_病人结帐作废_Modify("
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
                    strSQL = strSQL & "" & 2 & ","
                    '  病人id_In     病人结帐记录.病人id%Type,
                    strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
                    '  冲销id_In     病人预交记录.结帐id%Type,
                    strSQL = strSQL & "" & objBalanceInfor.冲销ID & ","
                    '  结算方式_In   Varchar2,
                    strSQL = strSQL & "'" & str结算方式 & "',"
                    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
                    strSQL = strSQL & objCard.接口序号 & ","
                    '  卡号_In       病人预交记录.卡号%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.卡号 & "',"
                    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.交易流水号 & "',"
                    '  交易说明_In   病人预交记录.交易说明%Type := Null,
                    strSQL = strSQL & "'" & objItemTemp.交易说明 & "',"
                    '  缴款_In       病人预交记录.缴款%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  找补_In       病人预交记录.找补%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  误差金额_In   病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '  预交金额_In   病人预交记录.冲预交%Type := Null,
                    strSQL = strSQL & "NULL,"
                    '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
                    strSQL = strSQL & "'" & UserInfo.编号 & "',"
                    '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    
                    '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
                    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                    '    冲预交病人ids_In Varchar2 := Null,
                    strSQL = strSQL & "NULL,"
                    '  完成作废_In Number:=0
                    strSQL = strSQL & "0,"
                    '  校对标志_In  Number := 0,
                    strSQL = strSQL & "" & objItemTemp.校对标志 & ","
                    '  关联交易id_In 病人预交记录.Id%Type := Null,
                    strSQL = strSQL & "" & IIf(objItemTemp.关联交易ID = 0, "NULL", objItemTemp.关联交易ID) & ","
                    '  清除原交易_In Number:=0
                    strSQL = strSQL & "" & IIf(lngCount = 0, "1", "0") & ")"
                    zlAddArray cllUpdate, strSQL
                End If
                lngCount = lngCount + 1
                .MoveNext
            Loop
            objBalanceItems_out.结算金额 = dblMoney
        End With
        blnChangeMoney_Out = RoundEx(dblDelMoney, 6) <> RoundEx(-1 * dblMoney, 6)
    Else
        For Each objItem In objBalanceItems_out
            objItem.退款交易流水号 = strSwapNO
            objItem.退款交易说明 = strSwapMemo
        Next
        Set objItem = objCurDelBalanceItems(1)
        If objItem.是否预交 And (mbln结算异步控制 Or objItem.退款交易流水号 <> strSwapNO Or objItem.退款交易说明 <> strSwapMemo) Then
            Call zlGetSaveThirdSwapDelSQLFromBalanceItem(objItem, True, cllUpdate, 0)
        End If
        blnChangeMoney_Out = False
    End If
    
    'blnTrans = True
    zlExecuteProcedureArrAy cllUpdate, mfrmMain.Caption, , Not mbln结算异步控制 Or blnTrans, "结帐", "病人结帐调试日志"
    blnTrans = False
    For Each objItem In objBalanceItems_out
        objItem.是否保存 = True
        objItem.是否允许编辑 = False
        objItem.是否允许删除 = False
        objItem.结帐ID = objBalanceInfor.结帐ID
        objItem.结帐时间 = objBalanceInfor.结帐时间
        objItem.冲销ID = objBalanceInfor.冲销ID
        objItem.是否结算 = True
        objItem.校对标志 = 2
    Next
    
    blnSaveed_Out = True
    
    Set cllBillPro = New Collection
    blnTrans = False
    
    zlThird_ReturnMoney = True
    Call zlGetThreeSwapExpendSQLToCollect(objCard, lng结帐ID, strCardNo, rsExpend, cllThreeSwap, False)
    Err = 0: On Error GoTo ErrOther:
    
    blnTrans = True
    '更新其他结算信息
    zlExecuteProcedureArrAy cllThreeSwap, mfrmMain.Caption & "-执行三方接口", , , "结帐", "病人结帐调试日志"
    blnTrans = False
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true")
    Exit Function
ErrHand:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If blnTrans And blnCallInterfance = False Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call SaveErrLog
    If blnTrans Then gcnOracle.RollbackTrans
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
    Exit Function
ErrOther:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    If blnTrans Then gcnOracle.CommitTrans
    Call SaveErrLog
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True", "(ErrOther),Err=" & mstrErrInfor)
End Function



Public Function zlGetThreeSwapExpendSQLToCollect(objCard As Card, ByVal lng结帐ID As Long, ByVal str卡号 As String, ByVal rsExpend As ADODB.Recordset, _
    ByRef cllPro As Collection, Optional bln预交 As Boolean = False, Optional lng预交ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取扩展信息保存的SQL给集合
    '入参:
    '出参:
    '返回:成功返回=true,否则返回Fale
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
            If NVL(!项目名称) <> "" Then
                strTemp = NVL(!项目名称) & "|" & NVL(!项目内容)
                If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                        str交易信息 = Mid(str交易信息, 3)
                        'Zl_三方结算交易_Insert
                        strSQL = "Zl_三方结算交易_Insert("
                        '卡类别id_In 病人预交记录.卡类别id%Type,
                        strSQL = strSQL & objCard.接口序号 & ","
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
        strSQL = strSQL & objCard.接口序号 & ","
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
 
Public Function zlThird_Payment_IsValied(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, _
    ByVal objCurItem As clsBalanceItem, ByRef objBalanceItems_out As clsBalanceItems, ByRef dbl帐户余额_Out As Double, _
    Optional byt操作类型 As Byte = 0, Optional ByVal objOldBalanceItems As clsBalanceItems, Optional str费用性质 As String, _
    Optional ByVal rsClassFee As ADODB.Recordset, Optional strNos As String = "", Optional blnCurInput As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查三方帐户是否有效
    '入参:frmMain-调用的主窗体
    '     objPati-病人信息对象
    '     objCurItem-当前结算信息
    '     dblMoney-本次消费金额
    '     rsClassFee-可选参数，表示分收费类别统计的费用，主要应用于刷消费卡的处理(字段;收费类别,金额)
    '     byt操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    '     str费用性质-本次支付的费用类型,多个用逗号分隔(如:1,2,3,4): 1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
    '     strNOs-划价收费时，传入
    '     objOldBalanceItem-消费卡时，转入历次刷卡信息
    '     blnCurInput-是否当前输入项
    '     strQRCode-二维码支付代码
    '出参:objBalanceItems_Out-返回结算数据对象集
    '     dbl帐户余额_out-返回帐户余额
    '返回:数据合法返回=true(（含非三卡或消费卡或Card对象 is nothig),否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 10:30:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objItem As clsBalanceItem, strXMLExpend As String
    Dim cllBalance As Collection, str卡号 As String, str密码 As String
    Dim str结算方式 As String, dbl结算金额 As Double, varData As Variant
    Dim i As Long, dbl未付金额 As Double
    Dim mstrFunName As String, strQRCode As String, mstrErrInfor As String
    Dim dblMoney As Double
    
    On Error GoTo errHandle
    mstrFunName = "zlThird_Payment_IsValied"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "objCurItem=" & IIf(objCurItem Is Nothing, "Nothing", "Not Nothing"))

    
    If objCurItem Is Nothing Then
        zlThird_Payment_IsValied = False
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objCurItem Is Nothing")
        Exit Function
    End If
    
    Set objCard = objCurItem.objCard
    If objCard Is Nothing Then
        zlThird_Payment_IsValied = True
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "objCard Is Nothing")
        Exit Function
    End If
    
    If objCard.接口序号 <= 0 Then
        zlThird_Payment_IsValied = True
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", objCard.接口序号, objCard.结算方式)
        Exit Function
    End If
    
    dblMoney = objCurItem.结算金额
    strQRCode = objCurItem.QRCode
    If objBalanceItems_out Is Nothing Then Set objBalanceItems_out = New clsBalanceItems
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If
        
    
    dbl未付金额 = RoundEx(objBalanceInfor.未付合计 - objBalanceInfor.冲预交合计 - objBalanceInfor.误差费, 6)
     
    If dblMoney = 0 Then
        If dbl未付金额 = 0 Then
            zlThird_Payment_IsValied = True:
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "dbl未付金额=0")
            Exit Function
        End If
        Screen.MousePointer = 0
        MsgBox "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 金额未输入,请检查置!")
        Exit Function
    End If
         
    If dblMoney <= 0 Then
        MsgBox "本次消费金额必须大于零，请检查!", vbInformation + vbOKOnly, gstrSysName
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 本次消费金额(" & dblMoney & ") 必须大于零，请检查")
        Exit Function
    End If
    
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "刷卡金额不能大于本次未付金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 刷卡金额(" & dblMoney & ")不能大于本次未付金额(" & Format(Abs(dbl未付金额), "0.00") & ") ！")
        
        Exit Function
    End If
    
    
    If dblMoney <> Format(dbl未付金额, "0.00") And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1  '提示
            If MsgBox("刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                
                
                 Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "提示选择否，信息： " & _
                "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
                Exit Function
            End If
        Case 2
            MsgBox "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！", vbInformation, gstrSysName
                
             Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", " " & _
                "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
                
            Exit Function
        End Select
    End If
    
    If objCard.是否独立结算 And (blnCurInput Or objCurItem.QRCode <> "") Then
        '124638:独立结算时，应该自付部分全部使用该结算方式。
        dbl结算金额 = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.医保支付合计 - objBalanceInfor.冲预交合计 - objBalanceInfor.误差费, 2)
        If dbl结算金额 <> dblMoney Then
            MsgBox objCard.名称 & "只能独立结算,不能在使用了其他结算方式后,再使用本结算方式，请检查!" & vbCrLf & "本次支付:" & Format(dblMoney, "0.00") & vbCrLf & "自付金额:" & Format(dbl结算金额, "0.00"), vbInformation + vbOKOnly, gstrSysName
            
             Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & "只能独立结算,不能在使用了其他结算方式后,再使用本结算方式，请检查!" & vbCrLf & "本次支付:" & Format(dblMoney, "0.00") & vbCrLf & "自付金额:" & Format(dbl结算金额, "0.00"))
            Exit Function
        End If
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
    '返回:成功,返回=true,否则返回False
    
    strXMLExpend = ""
    strXMLExpend = strXMLExpend & "<IN>" & vbCrLf
    strXMLExpend = strXMLExpend & "      <CZLX>" & byt操作类型 & "</CZLX>" & vbCrLf
    If strQRCode <> "" Then
        strXMLExpend = strXMLExpend & "      <QRCODE>" & strQRCode & "</QRCODE>" & vbCrLf
    End If
    strXMLExpend = strXMLExpend & "</IN>" & vbCrLf
    
    
    '操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    If gobjSquare.objOneCardComLib.zlBrushCard(mfrmMain, mlngModule, rsClassFee, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, False, True, False, False, cllBalance, False, False, strXMLExpend, _
             str费用性质, objPati.病人ID, str结算方式, dbl结算金额) = False Then
        
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlBrushCard=false", strXMLExpend, str结算方式, dbl结算金额)
        Exit Function
    End If
              
    
    If str结算方式 = "" Then
        str结算方式 = objCard.结算方式
        If dbl结算金额 = 0 Then dbl结算金额 = dblMoney
    End If
    If dbl结算金额 = 0 Then
        MsgBox "本次消费金额不能为零(接口返回0)，请与接口开发商或系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "本次消费金额不能为零(接口返回0)，请与接口开发商或系统管理员联系!")
        Exit Function
    End If

     Set objBalanceItems_out = New clsBalanceItems
     Set objItem = New clsBalanceItem
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
     objBalanceItems_out.AddItem objItem
     objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + dbl结算金额, 6)
 
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
    '返回:扣款合法,返回=true,否则返回Flase
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
    
    If gobjSquare.objOneCardComLib.zlPaymentCheck(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, str卡号, dblMoney, strNos, strXMLExpend) = False Then
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlPaymentCheck=false")
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

    If gobjSquare.objOneCardComLib.zlGetAccountMoney(mfrmMain, mlngModule, objCard.接口序号, _
     str卡号, strXMLExpend, dbl帐户余额_Out, objCard.消费卡) = False Then Exit Function
    If dbl帐户余额_Out <> 0 Then
        If dbl帐户余额_Out < dbl结算金额 Then
            MsgBox objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            
             Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额(" & dbl帐户余额_Out & ")不足!")
            Exit Function
        End If
        For Each objItem In objBalanceItems_out
            objItem.帐户余额 = dbl帐户余额_Out
        Next
    End If
    zlThird_Payment_IsValied = True
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function

Public Function zlThird_TransferAccounts_IsValied(ByVal objPati As clsPatientInfo, ByVal objCard As Card, objBalanceInfor As clsBalanceInfo, objCurTransBalanceItems As clsBalanceItems, _
    ByRef objBalanceItems_out As clsBalanceItems, Optional ByRef dbl帐户余额_Out As Double, Optional ByRef strErrMsg_out As String = "", Optional blnCurInput As Boolean, _
    Optional ByVal blnNotBrushCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行转帐较交易接口的合法性检查
    '入参:objPati-病人信息
    '    objCard-当前卡对象
    '    objBalanceInfor-当前结帐信息
    '    objCurTransBalanceItems-当前转帐信息
    '    -结帐ID
    '    blnCurInput-当前输入项
    '    blnNotBrushCard-不调用刷卡接口
    '出参:objBalanceItems_Out-转帐信息集
    '     strErrMsg_Out-返回的错误信息
    '返回:成功返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-29 10:08:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String, bln退现 As Boolean, str结算方式 As String, dbl结算金额 As Double
    Dim str卡号 As String, str密码 As String, dblMoney As Double, lng结帐ID As Long
    Dim objDelItem As clsBalanceItem, dbl未退金额 As Double
    Dim objItem As clsBalanceItem
    
    
    
    On Error GoTo errHandle
    mstrFunName = "zlThird_TransferAccounts_IsValied"
    Call WriteLog(mlngModule, "", mstrFunName, "开始")
    
    dbl帐户余额_Out = 0
    
    If objCard Is Nothing Then
        zlThird_TransferAccounts_IsValied = True
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true,objCard Is Nothing")
        Exit Function
    End If
    If objCard.接口序号 <= 0 Or objCard.消费卡 = True Then
        zlThird_TransferAccounts_IsValied = True
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", objCard.接口序号, IIf(objCard.消费卡, "true", "false"), objCard.结算方式)
        Exit Function
    End If
    
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If
        
    
    If objCurTransBalanceItems Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objCurTransBalanceItems Is Nothing!")
        Exit Function
    End If
    
    If objCurTransBalanceItems.Count = 0 Then
        Screen.MousePointer = 0
        MsgBox "当前退款结算信息不存在，请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objCurTransBalanceItems.Count =0!")
        Exit Function
    End If
    
    lng结帐ID = IIf(objBalanceInfor.冲销ID <> 0, objBalanceInfor.冲销ID, objBalanceInfor.结帐ID)
    
    dblMoney = objCurTransBalanceItems.结算金额
    dbl未退金额 = RoundEx(objBalanceInfor.未付合计 - objBalanceInfor.冲预交合计 + IIf(blnCurInput, 0, dblMoney) - objBalanceInfor.误差费, 2) '不是当前录入的，需要恢复本次的来判断

    If dblMoney = 0 Then
        If dbl未退金额 = 0 Then
            zlThird_TransferAccounts_IsValied = True
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "dbl未退金额=0!")
            Exit Function
        End If
        Screen.MousePointer = 0
        MsgBox "退款金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "退款金额未输入")
        Exit Function
    End If
        
    If dbl未退金额 = 0 Then
        Screen.MousePointer = 0
        MsgBox "注意:" & vbCrLf & objCard.名称 & "已经退完，不能再退！", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & "已经退完，不能再退！")
        Exit Function
    End If
    
    If RoundEx(Abs(dblMoney), 2) > RoundEx(Abs(dbl未退金额), 2) And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！", vbInformation, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "退款金额不能大于本次未退金额:" & Format(Abs(dbl未退金额), "0.00") & " ！")
        Exit Function
    End If
    
 
    If Abs(dblMoney) <> Abs(dbl未退金额) And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1 '提示
            Screen.MousePointer = 0
            If MsgBox("当前退款金额与未退金额不符,是否继续？" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(dblMoney, "0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "选择否,提示:当前退款金额与未退金额不符,不能继续!" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(dblMoney, "0.00") & "！")
                Exit Function
            End If
        
        Case 2 '禁止
            Call MsgBox("当前退款金额与未退金额不符,不能继续!" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(dblMoney, "0.00"), vbInformation + vbOKOnly, gstrSysName)
                
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "当前退款金额与未退金额不符,不能继续!" & vbCrLf & _
                "剩余未退:" & Format(Abs(dbl未退金额), "0.00") & _
                "当前退款:" & Format(dblMoney, "0.00") & "！")
            Exit Function
        End Select
    End If
    If objBalanceItems_out Is Nothing Then Set objBalanceItems_out = New clsBalanceItems
    
    If blnNotBrushCard Then
        '不调用刷卡接口交易，主要是先检查，后再进行接口调用，减少数据写入
        For Each objItem In objCurTransBalanceItems
            objItem.是否转帐 = True
        Next
        objCurTransBalanceItems.是否转帐 = True
        Set objBalanceItems_out = objCurTransBalanceItems
        zlThird_TransferAccounts_IsValied = True
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "不调用刷卡接口交易！")
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
    '返回:成功,返回=true,否则返回False
    
    strXMLExpend = "<IN><CZLX>1</CZLX></IN>"
    '操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    If gobjSquare.objOneCardComLib.zlBrushCard(mfrmMain, mlngModule, Nothing, objCard.接口序号, objCard.消费卡, "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, -1 * dblMoney, _
            str卡号, str密码, True, True, bln退现, True, Nothing, False, False, strXMLExpend, "", objPati.病人ID, str结算方式, dbl结算金额) = False Then
        
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlBrushCard=false", str卡号, str结算方式, dbl结算金额, strXMLExpend)
        
        Exit Function
    End If
                     
    
    If str结算方式 <> "" Then
        If RoundEx(dblMoney, 6) <> RoundEx(-1 * dbl结算金额, 6) Then
            strErrMsg_out = "当前转帐金额与接口返回的转帐金额不符，禁止转帐操作!" & vbCrLf & " 转帐金额:" & Format(-1 * dblMoney, "###0.00") & "  接口返回:" & Format(dbl结算金额, "###0.00") & ""
            
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", str卡号, str结算方式, strErrMsg_out)
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
    '                                <CZLX >操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；2-结帐业务;3-结帐退费业务
    '                            </IN>
    '                    Out  XML串:
    '                            <OUT>
    '                               <ERRMSG>错误信息</ERRMSG >
    '                            </OUT>
    '    Boolean 函数返回    检查的数据合法,返回=true:否则返回False
    '说明:
    '１. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '构造XML串
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    If gobjSquare.objOneCardComLib.zltransferAccountsCheck(mfrmMain, mlngModule, objCard.接口序号, _
        str卡号, RoundEx(-1 * dblMoney, 6), lng结帐ID, strXMLExpend, strErrMsg_out) = False Then
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zltransferAccountsCheck=false", str卡号, strXMLExpend, strErrMsg_out)
        Exit Function
    End If
         
    
    strXMLExpend = ""
    'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    '    ByVal strCardTypeID As Long, _
    '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '       入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '       出参:dblMoney-返回帐户余额
    Call gobjSquare.objOneCardComLib.zlGetAccountMoney(mfrmMain, mlngModule, objCard.接口序号, _
          str卡号, strXMLExpend, dbl帐户余额_Out, objCard.消费卡)
          
    Set objDelItem = New clsBalanceItem
    With objDelItem
        .结算方式 = IIf(str结算方式 = "", objCard.结算方式, str结算方式)
        .结算金额 = dblMoney
        .卡类别ID = objCard.接口序号
        .消费卡 = objCard.消费卡
        .卡号 = str卡号
        .密码 = str密码
        .帐户余额 = dbl帐户余额_Out
        .是否密文 = objCard.卡号密文规则 <> ""
        Set .objCard = objCard
        .结算性质 = objCard.结算性质
        .是否转帐 = True
        .冲销ID = objCurTransBalanceItems(1).冲销ID
        .结帐ID = objCurTransBalanceItems(1).结帐ID
        .结帐时间 = objCurTransBalanceItems(1).结帐时间
        .结算IDs = objCurTransBalanceItems(1).结算IDs
        .关联交易ID = objCurTransBalanceItems(1).关联交易ID
        
    End With
    objBalanceItems_out.是否转帐 = True
    objBalanceItems_out.AddItem objDelItem
    objBalanceItems_out.结算金额 = objBalanceItems_out.结算金额 + objDelItem.结算金额
    zlThird_TransferAccounts_IsValied = True
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", dbl帐户余额_Out)
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor)
End Function

Public Function zlGetBalanceItemsFromBalaceDatas(ByVal objCard As Card, ByVal objBalanceInfor As clsBalanceInfo) As clsBalanceItems
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的卡类别ID,获取结算明细项
    '入参:objCard-当前的卡对象
    '     objBalanceInfor-当前的结算信息集
    '出参:
    '返回:成功返回结算信息集,否则返回Nothing
    '编制:刘兴洪
    '日期:2018-03-28 15:57:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKey As String
    Dim objBalanceItems As clsBalanceItems
    Dim objData As clsBalanceData
    Dim objItem As clsBalanceItem
    Dim objItems_Out As clsBalanceItems
    
    Set objItems_Out = New clsBalanceItems
    If objCard Is Nothing Then Set zlGetBalanceItemsFromBalaceDatas = objItems_Out: Exit Function
    Err = 0: On Error GoTo errHandle
    For Each objData In objBalanceInfor
        Set objBalanceItems = objData.objBalanceItems
        If objBalanceItems Is Nothing Then Set objBalanceItems = New clsBalanceItems
        For Each objItem In objBalanceItems
            If objItem.卡类别ID = objCard.接口序号 And objItem.消费卡 = objCard.消费卡 Then
                objItems_Out.AddItem objItem
            End If
        Next
    Next
errHandle:
    Set zlGetBalanceItemsFromBalaceDatas = objItems_Out
End Function
 
 Public Sub zlInsure_InitPara(ByVal lng病人ID As Long, ByVal bytMCMode As Byte, ByVal intInsure As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2015-01-04 13:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mobjInsurePara.分币处理 = gclsInsure.GetCapability(support分币处理, lng病人ID, intInsure)
    mobjInsurePara.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure)
    mobjInsurePara.结帐作废后打印回单 = gclsInsure.GetCapability(support结帐作废后打印回单, lng病人ID, intInsure)
     
    If bytMCMode = 1 Then
        mobjInsurePara.门诊预结算 = gclsInsure.GetCapability(support门诊预算, lng病人ID, intInsure)
        mobjInsurePara.门诊必须传递明细 = gclsInsure.GetCapability(support门诊必须传递明细, lng病人ID, intInsure)
        mobjInsurePara.门诊结算_结帐设置 = gclsInsure.GetCapability(support门诊结帐_结帐设置后调用接口, lng病人ID, intInsure)
        mobjInsurePara.门诊病人结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure)
    Else
        mobjInsurePara.未结清出院 = gclsInsure.GetCapability(support未结清出院, lng病人ID, intInsure)
        mobjInsurePara.结算使用个人帐户 = gclsInsure.GetCapability(support结算使用个人帐户, lng病人ID, intInsure)
        mobjInsurePara.出院结算必须出院 = gclsInsure.GetCapability(support出院结算必须出院, lng病人ID, intInsure)
        mobjInsurePara.中途结帐仅处理已上传部分 = gclsInsure.GetCapability(support中途结算仅处理已上传部分, lng病人ID, intInsure)
        mobjInsurePara.结帐设置后调用接口 = gclsInsure.GetCapability(support结帐_结帐设置后调用接口, lng病人ID, intInsure)
        mobjInsurePara.住院结算作废 = gclsInsure.GetCapability(support住院结算作废, lng病人ID, intInsure)
        mobjInsurePara.门诊结算_结帐设置 = False
        mobjInsurePara.出院病人结算作废 = gclsInsure.GetCapability(support出院病人结算作废, lng病人ID, intInsure)
        mobjInsurePara.允许结多次住院费用 = gclsInsure.GetCapability(support允许一次结多次住院费用, lng病人ID, intInsure)
    End If
End Sub
Public Function zlInsure_BalanceUpdate(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByVal str医保结算 As String, _
    ByRef cllPro As Collection, Optional ByVal blnExcuteSQL As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保数据校对更正
    '入参:objPati-病人信息对象
    '     str医保结算：结算方式|结算金额||..
    '     blnExcuteSQL-是否执行相关的SQL
    '     objBalanceInfor-病人结帐信息数据
    '出参:cllPro-返回更新的相关SQL
    '     objBalanceInfor-返回更新后的结帐数据集
    '返回:校对成功,返回=true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-12 17:45:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    mstrFunName = "zlInsure_BalanceUpdate"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "str医保结算=" & str医保结算)

    'Zl_病人结帐结算_Modify_S
    strSQL = "Zl_病人结帐结算_Modify_S("
    '  操作类型_In     Number,
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    strSQL = strSQL & "" & 2 & ","
    '  病人id_In       病人结帐记录.病人id%Type,
    strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
    '  结帐id_In       病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
    '  结算方式_In     Varchar2,
    strSQL = strSQL & "" & IIf(str医保结算 = "", "NULL", "'" & str医保结算 & "'") & ","
    '  冲预交_In       病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  退支票额_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In         病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In     病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '  缴款_In         病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "NULL,"
    '  找补_In         病人预交记录.找补%Type := Null,
    strSQL = strSQL & "NULL,"
    '  误差金额_In     门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "NULL,"
    '  结帐类型_In     Number := 2,(1-门诊结帐;2-住院结帐)
    strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
    '  缺省结算方式_In 结算方式.名称%Type := Null,
    strSQL = strSQL & "NULL,"
    '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '    冲预交病人ids_In Varchar2 := Null,
    strSQL = strSQL & "NULL,"
    '    当前科室id_In    病人预交记录.科室id%Type := Null,
    strSQL = strSQL & ZVal(objPati.当前科室id) & ","
    '  完成结算_In Number:=0: (1-完成收费;0-未完成收费)
    strSQL = strSQL & "0,"
    '  校对标志_In      Number := 2,
    strSQL = strSQL & "2,"
    '  预交id_In        病人预交记录.Id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  关联交易id_In    病人预交记录.Id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  清除原交易_In Number:=0
    strSQL = strSQL & 0 & ")"
     
    If blnExcuteSQL Or cllPro Is Nothing Then
        Call zlDatabase.ExecuteProcedure(strSQL, mfrmMain.Caption)
    Else
        zlAddArray cllPro, strSQL
    End If
    
    zlInsure_BalanceUpdate = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
End Function
Public Function zlInsure_Check(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回=true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    
    On Error GoTo errHandle
    
    mstrFunName = "zlInsure_Check"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "strAdvance=" & strAdvance, "str保险结算=" & str保险结算)
    
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "需校对")
        Exit Function
    End If
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
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=" & IIf(blnMedicareCheck, "true", "false"))
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & Err.Description, 2)
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
    
    mstrFunName = "zlGetMedicareStr"
    Call WriteLog(mlngModule, "", mstrFunName, "开始")
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
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回", " cur个人帐户=" & cur个人帐户, "cur医保基金=" & cur医保基金, "zlGetMedicareStr=" & strTemp)
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=,Err=" & mstrErrInfor, 2)
End Function

Public Function zlInsure_GetSelfBalance(ByVal lng病人ID As Long, ByVal intInsure As Integer, str医保号 As String, ByVal bytMCMode As Byte, _
    ByRef objInsureInfor_Out As clsInsureInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取账户余额信息
    '入参:lng病人ID-病人ID
    '     str医保号-医保号
    '     bytMode-1-门诊，2-住院
    '出参:objInsureInfor_Out-返回医保对象
    '返回:成功返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-20 14:56:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur账户余额  As Currency, cur透支额 As Currency
    
    On Error GoTo errHandle
    
    mstrFunName = "zlInsure_GetSelfBalance"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", lng病人ID, intInsure, str医保号, bytMCMode)
    
    If objInsureInfor_Out Is Nothing Then Set objInsureInfor_Out = New clsInsureInfor
    cur账户余额 = gclsInsure.SelfBalance(lng病人ID, str医保号, IIf(bytMCMode = 1, 10, 40), cur透支额, intInsure)
    objInsureInfor_Out.帐户余额 = cur账户余额
    objInsureInfor_Out.个帐透支 = cur透支额
    zlInsure_GetSelfBalance = True
     
   Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True", "cur账户余额=" & cur账户余额, "cur透支额=" & cur透支额)
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
End Function


Public Function zlInsure_SettleSwap(ByVal objPati As clsPatientInfo, ByVal objPreBalanceItems As clsBalanceItems, _
    ByVal cllSaveBill As Collection, ByRef objBalanceInfor As clsBalanceInfo, ByRef objItems_Out As clsBalanceItems, ByRef blnProof_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保结算交易
    '入参:cllSaveBill-保存单据的sql
    '     objPreBalanceItems-医保预结算集
    '出参:
    '     objItems_Out-医保结算信息(出参)
    '     blnProof_Out-预结算和医保结算不一致时返回=true,否则返回False
    '返回:执行成功,返回=true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-13 15:11:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, i As Long, str结算信息 As String
    Dim blnTrans As Boolean, intInsure As Integer, strAdvance As String
    Dim blnTransMC As Boolean, blnMark As Boolean
    Dim cur个人帐户 As Currency, cur医保基金 As Currency
    Dim blnInsureCheck As Boolean, varTemp As Variant, varData As Variant
    Dim objItem As clsBalanceItem
    Dim rs结算方式 As ADODB.Recordset
    Dim dblMoney As Double
    
    On Error GoTo errHandle
    
    mstrFunName = "zlInsure_SettleSwap"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "intInsure=" & intInsure)
    
    intInsure = objBalanceInfor.objInsure.险类
    '非医保或结自费费用,不能执行
    If intInsure = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "intInsure = 0")
        Exit Function
    End If
    
    blnProof_Out = False
    Set cllPro = New Collection
    For i = 1 To cllSaveBill.Count
        zlAddArray cllPro, cllSaveBill(i)
    Next
    
    
    str结算信息 = zlGetMedicareStr(objPreBalanceItems, cur个人帐户, cur医保基金)
    
    dblMoney = objPreBalanceItems.结算金额
     
    '按预结算修正医保结算数据
    If zlInsure_BalanceUpdate(objPati, objBalanceInfor, str结算信息, cllPro) = False Then
    
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "按预结算修正医保结算数据,zlInsure_BalanceUpdate=fase")
        Exit Function
    End If
    
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, True, , "结帐", "病人结帐调试日志"
    '调用医保接口
    blnTransMC = False
    If objBalanceInfor.objInsure.bytMCMode = 1 Then
        '门诊医保结算
        strAdvance = ""
        If cur个人帐户 <> 0 Or cur医保基金 <> 0 Or mobjInsurePara.门诊必须传递明细 Then
            If Not gclsInsure.ClinicSwap(objBalanceInfor.结帐ID, cur个人帐户, cur医保基金, 0, 0, intInsure, strAdvance) Then
                gcnOracle.RollbackTrans:
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "调用医保接口(ClinicSwap=false)")
                Exit Function
            End If
            blnTransMC = True
        End If
    Else
        '住院医保结算
        If Not gclsInsure.SettleSwap(objBalanceInfor.结帐ID, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "调用医保接口(SettleSwap=false)")
            Exit Function
        End If
        blnTransMC = True
    End If
    
    '根据返回信息进行数据更新
    blnProof_Out = False
    If strAdvance <> "" Then
        If zlInsure_Check(str结算信息, strAdvance) Then
            blnProof_Out = True
            Call zlInsure_BalanceUpdate(objPati, objBalanceInfor, strAdvance, Nothing, True)
        End If
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    Set objItems_Out = objPreBalanceItems
    
    If blnProof_Out Then
        Set rs结算方式 = zlGet结算方式
        '分解医保结算信息
        'strAdvance:结算方式|结算金额||.."
        dblMoney = 0
        Set objItems_Out = New clsBalanceItems
        varData = Split(strAdvance, "||")
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & "|", "|")
            
            Set objItem = New clsBalanceItem
            rs结算方式.Filter = "名称='" & varTemp(0) & "'"
            With objItem
                .结算方式 = varTemp(0)
                .结算金额 = Val(varTemp(1))
                If rs结算方式.EOF = False Then .结算性质 = Val(NVL(rs结算方式!性质))
                Set .objCard = zlGetCardFromBalanceName(varTemp(0))
                .结算类型 = 2
                .是否允许编辑 = False
                .是否允许删除 = False
                .是否结算 = True
                .是否保存 = True
                .校对标志 = 2
                .结帐时间 = objBalanceInfor.结帐时间
                .结帐ID = objBalanceInfor.结帐ID
                dblMoney = RoundEx(dblMoney + .结算金额, 6)
            End With
            rs结算方式.Filter = 0
            objItems_Out.AddItem objItem
            objItems_Out.结算金额 = dblMoney
        Next
        objBalanceInfor.已付合计 = dblMoney
    Else
        For Each objItem In objItems_Out
            With objItem
                .结算类型 = 2
                .是否允许编辑 = False
                .是否允许删除 = False
                .是否结算 = True
                .是否保存 = True
                .校对标志 = 2
                .结帐时间 = objBalanceInfor.结帐时间
                .结帐ID = objBalanceInfor.结帐ID
            End With
        Next
    End If
    
    objBalanceInfor.是否保存结帐单 = True
    objBalanceInfor.未付合计 = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.已付合计, 6)
    
    If blnTransMC Then
        Call gclsInsure.BusinessAffirm(IIf(objBalanceInfor.objInsure.bytMCMode = 1, 交易Enum.Busi_ClinicSwap, 交易Enum.Busi_SettleSwap), True, intInsure)
    End If
    Set cllSaveBill = New Collection
    zlInsure_SettleSwap = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    Screen.MousePointer = 0
    mstrErrInfor = Err.Number & "-" & Err.Description
    
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call SaveErrLog
    
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(objBalanceInfor.objInsure.bytMCMode = 1, 交易Enum.Busi_ClinicSwap, 交易Enum.Busi_SettleSwap), False, intInsure)
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)

End Function


Public Function zlInsure_SetleDelSwap_IsValied(ByVal objPati As clsPatientInfo, ByVal objBalanceInfor As clsBalanceInfo, ByVal objCurItems As clsBalanceItems, ByVal bln异常作废 As Boolean) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保退费相关检查
    '入参:objPati-病人信息对象
    '     objBalanceInfor-病人结算信息
    '     objCurItems-医保当前结算信息对象
    '     bln异常作废-当前操作是异常作废
    '出参:
    '返回:合法返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-25 21:08:10
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim intInsure As Integer
    Dim objItem As clsBalanceItem
    
    On Error GoTo errHandle
    
    mstrFunName = "zlInsure_SetleDelSwap_IsValied"
    Call WriteLog(mlngModule, "", mstrFunName, "开始")
    
    If objBalanceInfor Is Nothing Then
        MsgBox "当前无法确定保险相关信息，不能执行作废操作!", vbInformation + vbOKOnly, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "当前无法确定保险相关信息,objBalanceInfor Is Nothing")
        Exit Function
    End If
    
    intInsure = objBalanceInfor.objInsure.险类
    If intInsure = 0 Then
        MsgBox "当前无法确定保险相关信息，不能执行作废操作!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "当前无法确定保险相关信息,intInsure=0")
        Exit Function
    End If
    If objCurItems Is Nothing Then
        MsgBox "当前无法确定保险相关信息，不能执行作废操作!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "当前无法确定保险相关信息,objCurItems Is Nothing")
        Exit Function
    End If
    
    On Error GoTo errHandle

    For Each objItem In objCurItems
        If objItem.是否结算 And Not bln异常作废 Then
            zlInsure_SetleDelSwap_IsValied = True
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "已经结算的，直接返回=true")
            Exit Function '已经结算的，直接返回=true
        End If
    Next
        

    If objBalanceInfor.objInsure.bytMCMode = 1 Then '门诊
        If mobjInsurePara.门诊病人结算作废 = False Then  '不支持门诊结算作废,则直接返回
            If objCurItems.Count = 0 Then
                zlInsure_SetleDelSwap_IsValied = True
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "无结算数据且不支持门诊结算作废,objCurItems.Count = 0!")
                Exit Function
            End If
            
            
            MsgBox "由于该医保不支持门诊结帐作废,因此,不能执行作废操作!", vbInformation + vbOKOnly, gstrSysName
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "由于该医保不支持门诊结帐作废,因此,不能执行作废操作!")
            Exit Function
        End If
        
        If gclsInsure.CheckInsureValid(intInsure) = False Then
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "调用医保接口(CheckInsureValid=false),门诊")
            Exit Function
        End If
        zlInsure_SetleDelSwap_IsValied = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "门诊")
        Exit Function
    End If
    
    If Not mobjInsurePara.出院病人结算作废 Then
        If objPati.险类 <> 0 And objPati.在院 = False Then
            MsgBox "该参保病人已经出院，不能取消该结帐单！", vbInformation, gstrSysName: Exit Function
            
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "该参保病人已经出院，不能取消该结帐单")
        End If
    End If
    
    If mobjInsurePara.住院结算作废 = False Then  '不支持门诊结算作废,则直接返回
        If objCurItems.Count = 0 Then
            zlInsure_SetleDelSwap_IsValied = True
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "无结算数据且不支持住院结算作废,objCurItems.Count = 0!")
            Exit Function
        End If
        
        MsgBox "由于该医保不支持住院结帐作废,因此,不能执行作废操作!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "由于该医保不支持住院结帐作废")
        Exit Function
    End If
    
    If gclsInsure.CheckInsureValid(intInsure) = False Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "调用医保接口(CheckInsureValid=false)住院")
        Exit Function
    End If
    zlInsure_SetleDelSwap_IsValied = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
End Function

Public Function zlInsure_SettleDelSwap(ByVal objPati As clsPatientInfo, ByVal objBalanceInfor As clsBalanceInfo, ByRef objCurItems As clsBalanceItems, ByRef cllDelBalancePro As Collection, _
    Optional bln异常作废 As Boolean, Optional objItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保退费操作
    '入参:cllDelBalancePro-执行的退费单据
    '     strYbBalances-医保结处信息
    '     objCurItems-当前的医保结算信息集
    '     bln异常作废-是否异常作废
    '出参: objItems_Out-医保退费返回的相关信息,Nothing表示与原结算方式保持一致，不需作调整
    '返回:退费成功,返回=true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-22 16:39:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String, strYbBalance As String, objItem As clsBalanceItem
    Dim blnTransMC As Boolean, blnTrans As Boolean, i As Long
    Dim strAdvance  As String, varData As Variant, varTemp As Variant
    Dim intInsure As Integer
    
    
    
    On Error GoTo errHandle
    
    mstrFunName = "zlInsure_SettleDelSwap"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "冲销ID=" & objBalanceInfor.冲销ID)
    
    intInsure = objBalanceInfor.objInsure.险类
    If intInsure = 0 Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "intInsure=" & intInsure)
        Exit Function
    End If
    
    If objBalanceInfor.冲销ID = 0 Then
        MsgBox "未产生作废数据，请与系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "未产生作废数据,objBalanceInfor.冲销ID=0")
        Exit Function
    End If
    
    
    '获取医保结算方式
    strYbBalance = ""
    If bln异常作废 = False Then
        For Each objItem In objCurItems
            If objItem.是否结算 And bln异常作废 = False Then zlInsure_SettleDelSwap = True: Exit Function   '结算了的，不处理
             strYbBalance = strYbBalance & "||" & objItem.结算方式 & "|" & -1 * objItem.结算金额
        Next
        strYbBalance = Mid(strYbBalance, 3)
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllDelBalancePro.Count
        cllPro.Add cllDelBalancePro(i)
    Next
    If strYbBalance <> "" Then
        '    Zl_病人结帐作废_Modify
         strSQL = "Zl_病人结帐作废_Modify("
        '  操作类型_In      Number,
         '--   1-普通退费方式:
         '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
         '--   2.三方卡退费结算:
         '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
         '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
         '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
         '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
         '--   4-消费卡结算:
         '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        strSQL = strSQL & "" & 3 & ","
        '  病人id_In        病人结帐记录.病人id%Type,
        strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
        '  冲销id_In        病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & objBalanceInfor.冲销ID & ","
        '  结算方式_In      Varchar2,
        strSQL = strSQL & "'" & strYbBalance & "',"
        '  卡类别id_In      病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In          病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In    病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In      病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  缴款_In          病人预交记录.缴款%Type := Null,
        strSQL = strSQL & "NULL,"
        '  找补_In          病人预交记录.找补%Type := Null,
        strSQL = strSQL & "NULL,"
        '  误差金额_In      病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  预交金额_In      病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  操作员编号_In    病人预交记录.操作员编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  收款时间_In      病人预交记录.收款时间%Type := Null,
        strSQL = strSQL & "NULL,"
        '  冲预交病人ids_In Varchar2 := Null,
        '多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
        strSQL = strSQL & "NULL,"
        '  完成作废_In      Number := 0,
        strSQL = strSQL & "0,"
        '  校对标志_In  Number := 0,
        strSQL = strSQL & "0,"
        '  关联交易id_In    病人预交记录.Id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  清除原交易_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End If
    
    '执行医保退费
    Err = 0: On Error GoTo ErrRoll:
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, mfrmMain.Caption, True, , "结帐", "病人结帐调试日志"
    blnTransMC = False
    If objBalanceInfor.objInsure.bytMCMode = 1 Then
        If mobjInsurePara.门诊病人结算作废 Then '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
            strAdvance = objBalanceInfor.冲销ID & "|0"
            If Not gclsInsure.ClinicDelSwap(objBalanceInfor.结帐ID, , intInsure, strAdvance) Then
                gcnOracle.RollbackTrans
                Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "调用医保接口(ClinicDelSwap=false)")
                Exit Function
            End If
            blnTransMC = True
        Else
            strAdvance = strYbBalance
        End If
        
        If zlInsureCheck(strYbBalance, strAdvance) Then
            'Zl_病人结帐作废_Modify
             strSQL = "Zl_病人结帐作废_Modify("
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
             
            strSQL = strSQL & "" & 3 & ","
            '  病人id_In     病人结帐记录.病人id%Type,
            strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
            '  冲销id_In     病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & objBalanceInfor.冲销ID & ","
            '  结算方式_In   Varchar2,
            strSQL = strSQL & "'" & strAdvance & "')"
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            '  卡号_In       病人预交记录.卡号%Type := Null,
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            '  缴款_In       病人预交记录.缴款%Type := Null,
            '  找补_In       病人预交记录.找补%Type := Null,
            '  误差金额_In   病人预交记录.冲预交%Type := Null,
            '  预交金额_In   病人预交记录.冲预交%Type := Null,
            '操作员编号_In    病人预交记录.操作员编号%Type := Null,
            '操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
            '收款时间_In      病人预交记录.操作员姓名%Type := Null,
            '冲预交病人ids_In Varchar2 := Null,
            ' 多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
            '  完成作废_In Number:=0
            zlDatabase.ExecuteProcedure strSQL, mfrmMain.Caption
            Set objItems_Out = New clsBalanceItems
            varData = Split(strAdvance, "||")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||", "|")
                Set objItem = New clsBalanceItem
                Set objItem.objCard = zlGetCardFromBalanceName(varTemp(0))
                With objItem
                    .结算IDs = objBalanceInfor.结帐ID
                    .结帐ID = objBalanceInfor.结帐ID
                    .冲销ID = objBalanceInfor.冲销ID
                    .结帐时间 = objBalanceInfor.结帐时间
                    .结算类型 = 2
                    .结算金额 = RoundEx(Val(varTemp(1)), 6)
                    .结算性质 = .objCard.结算性质
                    .门诊结帐 = objBalanceInfor.结算类型 = 1 'IIf(mEditType = g_Ed_门诊结帐, 1, 2)
                    .是否允许编辑 = False
                    .是否允许删除 = False
                    .是否允许退现 = False
                    .是否保存 = True
                    .是否强制退现 = False
                    .校对标志 = 2
                End With
                objItems_Out.AddItem objItem
                objItems_Out.结算金额 = objItems_Out.结算金额 + objItem.结算金额
            Next
        Else
            For Each objItem In objCurItems
                With objItem
                    .结算IDs = objBalanceInfor.结帐ID
                    .结帐ID = objBalanceInfor.结帐ID
                    .冲销ID = objBalanceInfor.冲销ID
                    .结帐时间 = objBalanceInfor.结帐时间
                    .是否允许编辑 = False
                    .是否允许删除 = False
                    .是否允许退现 = False
                    .是否保存 = True
                    .是否强制退现 = False
                    .校对标志 = 2
                End With
            Next
        End If
    Else
        strAdvance = ""
        If Not gclsInsure.SettleDelSwap(objBalanceInfor.结帐ID, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "调用医保接口(SettleDelSwap=false)")
            Exit Function
        End If
        For Each objItem In objCurItems
            With objItem
                .结算IDs = objBalanceInfor.结帐ID
                .结帐ID = objBalanceInfor.结帐ID
                .冲销ID = objBalanceInfor.冲销ID
                .结帐时间 = objBalanceInfor.结帐时间
                .是否允许编辑 = False
                .是否允许删除 = False
                .是否允许退现 = False
                .是否保存 = True
                .是否强制退现 = False
                .校对标志 = 2
            End With
        Next
        blnTransMC = True
    End If
    gcnOracle.CommitTrans: blnTrans = False
    Set cllDelBalancePro = New Collection   '清空保存作废结帐单据数据
    
    If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(objBalanceInfor.objInsure.bytMCMode = 1, 交易Enum.Busi_ClinicDelSwap, 交易Enum.Busi_SettleDelSwap), True, intInsure)
    zlInsure_SettleDelSwap = True
    
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description

    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "错误信息=" & mstrErrInfor, 2)
    Exit Function
ErrRoll:
    mstrErrInfor = Err.Number & "-" & Err.Description

    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(objBalanceInfor.objInsure.bytMCMode = 1, 交易Enum.Busi_ClinicDelSwap, 交易Enum.Busi_SettleDelSwap), False, intInsure)
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
End Function

Public Function zlSquare_AddBalanceDatas(ByVal objCard As Card, ByVal objBalanceItems As clsBalanceItems, ByRef objBalanceInfor As clsBalanceInfo, ByRef objBalanceItems_out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向结算数据信息中加入消费卡结算信息
    '    objCard-当前支付的卡对象
    '    objBalanceInfor-当前结帐数据
    '    objBalanceItems-消费卡刷卡信息数据集
    '出参:
    '返回:成功返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-10 18:43:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objData As clsBalanceData
    
    On Error GoTo errHandle
    mstrFunName = "zlSquare_AddBalanceDatas"
    
    Call WriteLog(mlngModule, "", mstrFunName, "开始")
    Set objBalanceItems_out = New clsBalanceItems
    
    If objBalanceItems Is Nothing Then
        zlSquare_AddBalanceDatas = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "objBalanceItems Is Nothing")
        Exit Function
    End If
    If objBalanceItems.Count = 0 Then
        zlSquare_AddBalanceDatas = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "objBalanceItems.Count = 0")
        Exit Function
    End If
    objBalanceInfor.已付合计 = RoundEx(objBalanceInfor.已付合计 + objBalanceItems.结算金额, 6)
    objBalanceInfor.未付合计 = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.已付合计, 6)
    Set objBalanceItems_out = objBalanceItems
    
    zlSquare_AddBalanceDatas = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
End Function
Public Function zlSquare_PaymentSQL(ByVal objPati As clsPatientInfo, ByRef cllBillPro As Collection, _
    ByRef objBalanceInfor As clsBalanceInfo, ByRef objBalanceItems As clsBalanceItems, Optional ByVal bln完成结算 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行消费卡扣款及保存业务
    '入参:objCard-当前支付的卡对象
    '    objPati-病人对象集
    '    cllBillPro-执行的相关过程
    '    objBalanceInfor-当前结帐数据
    '    blnCommit-是否提交数据
    '出参:cllBillPro-返回执行扣款的SQL
    '     objBalanceItems_Out-返回消费卡扣款集
    '返回:获取成功,返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-03-23 17:02:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objData As clsBalanceData, objItems As clsBalanceItems, objItem As clsBalanceItem
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str消费卡结算 As String, strSQL As String
     
    On Error GoTo errHandle
    mstrFunName = "zlSquare_PaymentSQL"
    
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "")
    
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    If objBalanceItems Is Nothing Then
        zlSquare_PaymentSQL = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "objBalanceItems Is Nothing")
        Exit Function
    End If
    
    For Each objItem In objBalanceItems
        If objItem.消费卡 Then
            str消费卡结算 = str消费卡结算 & "||" & objItem.卡类别ID
            str消费卡结算 = str消费卡结算 & "|" & objItem.卡号
            str消费卡结算 = str消费卡结算 & "|" & objItem.消费卡ID
            str消费卡结算 = str消费卡结算 & "|" & RoundEx(objItem.结算金额, 6)
        End If
    Next
 
    If str消费卡结算 <> "" Then str消费卡结算 = Mid(str消费卡结算, 3)
    If str消费卡结算 = "" Then
        zlSquare_PaymentSQL = True
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "str消费卡结算=")
        Exit Function
    End If
    
    'Zl_病人结帐结算_Modify_S
    strSQL = "Zl_病人结帐结算_Modify_S("
    '  操作类型_In     Number,
    '  --操作类型_In:
    '--   3-消费卡结算:
    '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位
    strSQL = strSQL & "3,"
    '  病人id_In       病人结帐记录.病人id%Type,
    strSQL = strSQL & "" & ZVal(objPati.病人ID) & ","
    '  结帐id_In       病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & objBalanceInfor.结帐ID & ","
    '  结算方式_In     Varchar2,
    strSQL = strSQL & "'" & str消费卡结算 & "',"
    '  冲预交_In       病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  退支票额_In     病人预交记录.冲预交%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡类别id_In     病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  卡号_In         病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易流水号_In   病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL,"
    '  交易说明_In     病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL,"
    '  缴款_In         病人预交记录.缴款%Type := Null,
    strSQL = strSQL & "NULL,"
    '  找补_In         病人预交记录.找补%Type := Null,
    strSQL = strSQL & "NULL,"
    '  误差金额_In     门诊费用记录.实收金额%Type := Null,
    strSQL = strSQL & "NULL,"
    '  结帐类型_In     Number := 2,
    strSQL = strSQL & "" & IIf(objBalanceInfor.结算类型 = 1, 1, 2) & ","
    '  缺省结算方式_In 结算方式.名称%Type := Null,
    strSQL = strSQL & "NULL,"
    '    操作员编号_In    病人预交记录.操作员编号%Type := Null,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '    操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '    收款时间_In      病人预交记录.操作员姓名%Type := Null,
    strSQL = strSQL & "to_date('" & Format(objBalanceInfor.结帐时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '    冲预交病人ids_In Varchar2 := Null,
    strSQL = strSQL & "NULL,"
    '    当前科室id_In    病人预交记录.科室id%Type := Null,
        strSQL = strSQL & ZVal(objPati.当前科室id) & ","
    '  完成结算_In Number:=0
    strSQL = strSQL & "" & IIf(bln完成结算, 1, "0") & " )"
    zlAddArray cllBillPro, strSQL
    zlSquare_PaymentSQL = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True", "str消费卡结算=" & str消费卡结算)
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
     Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
End Function

 
Public Function zlSquare_Payment_IsValied(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, _
    ByVal objCurItem As clsBalanceItem, ByRef objBalanceItems_out As clsBalanceItems, ByRef dbl帐户余额_Out As Double, _
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
    '返回:数据合法返回=true(（含非三卡或消费卡或Card对象 is nothig),否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-23 10:30:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objItem As clsBalanceItem, strXMLExpend As String
    Dim cllBalance As Collection, str卡号 As String, str密码 As String
    Dim str结算方式 As String, dbl结算金额 As Double, varData As Variant
    Dim i As Long, dbl未付金额 As Double
    Dim mstrFunName As String, strQRCode As String, mstrErrInfor As String
    Dim dblMoney As Double
    
    On Error GoTo errHandle
    
    mstrFunName = "zlSquare_Payment_IsValied"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "objCurItem=" & IIf(objCurItem Is Nothing, "Nothing", "Not Nothing"))

    
    If objCurItem Is Nothing Then
        zlSquare_Payment_IsValied = False
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "objCurItem Is Nothing")
        Exit Function
    End If
    
    Set objCard = objCurItem.objCard
    If objCard Is Nothing Then
        zlSquare_Payment_IsValied = True
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "objCard Is Nothing")
        Exit Function
    End If
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then
        zlSquare_Payment_IsValied = True
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "objCard.接口序号=" & objCard.接口序号, " objCard.结算方式" & objCard.结算方式)
        Exit Function
    End If
    
    dblMoney = objCurItem.结算金额
    If objBalanceItems_out Is Nothing Then Set objBalanceItems_out = New clsBalanceItems
    
    If zlCheckIsStartFromCard(objCard) = False Then
        MsgBox objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!", vbInformation, gstrSysName
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 未启用,请在『设备配置』中设置启用或在结算方式应用场合中设置!")
        Exit Function
    End If
        
    
    dbl未付金额 = RoundEx(objBalanceInfor.未付合计 - objBalanceInfor.冲预交合计 - objBalanceInfor.误差费, 6)
     
    If dblMoney = 0 Then
        If dbl未付金额 = 0 Then
            zlSquare_Payment_IsValied = True:
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=true", "dbl未付金额=0")
            Exit Function
        End If
        Screen.MousePointer = 0
        MsgBox "金额未输入,请检查!", vbInformation + vbOKOnly, gstrSysName
        
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 金额未输入,请检查置!")
        Exit Function
    End If
         
    If dblMoney <= 0 Then
        MsgBox "本次消费金额必须大于零，请检查!", vbInformation + vbOKOnly, gstrSysName
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 本次消费金额(" & dblMoney & ") 必须大于零，请检查")
        Exit Function
    End If
    
    If Abs(dblMoney) > Format(Abs(dbl未付金额), "0.00") And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox "刷卡金额不能大于本次未付金额:" & Format(Abs(dbl未付金额), "0.00") & " ！", vbInformation, gstrSysName
        
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & " 刷卡金额(" & dblMoney & ")不能大于本次未付金额(" & Format(Abs(dbl未付金额), "0.00") & ") ！")
        
        Exit Function
    End If
    
    
    If dblMoney <> Format(dbl未付金额, "0.00") And blnCurInput Then
        Select Case mbyt刷卡缺省金额操作
        Case 1  '提示
            If MsgBox("刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                
                
                 Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "提示选择否，信息： " & _
                "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
                Exit Function
            End If
        Case 2
            MsgBox "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！", vbInformation, gstrSysName
                
             Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", " " & _
                "刷卡金额(" & Format(dblMoney, "0.00") & ")与本次未付金额(" & Format(dbl未付金额, "0.00") & _
                ")不同，不能继续！")
                
            Exit Function
        End Select
    End If
    
    If objCard.是否独立结算 And blnCurInput Then
        '124638:独立结算时，应该自付部分全部使用该结算方式。
        dbl结算金额 = RoundEx(objBalanceInfor.当前结帐 - objBalanceInfor.医保支付合计 - objBalanceInfor.冲预交合计 - objBalanceInfor.误差费, 2)
        If dbl结算金额 <> dblMoney Then
            MsgBox objCard.名称 & "只能独立结算,不能在使用了其他结算方式后,再使用本结算方式，请检查!" & vbCrLf & "本次支付:" & Format(dblMoney, "0.00") & vbCrLf & "自付金额:" & Format(dbl结算金额, "0.00"), vbInformation + vbOKOnly, gstrSysName
            
             Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", " " & _
                objCard.名称 & "只能独立结算,不能在使用了其他结算方式后,再使用本结算方式，请检查!" & vbCrLf & "本次支付:" & Format(dblMoney, "0.00") & vbCrLf & "自付金额:" & Format(dbl结算金额, "0.00"))
            Exit Function
        End If
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
    '返回:成功,返回=true,否则返回False
    
    strXMLExpend = ""
    '操作类型:0-正常调用刷卡；1-转账调用刷卡；2-退款调用刷卡
    If gobjSquare.objOneCardComLib.zlBrushCard(mfrmMain, mlngModule, rsClassFee, _
            objCard.接口序号, objCard.消费卡, _
            "" & objPati.姓名, "" & objPati.性别, "" & objPati.年龄, dblMoney, _
            str卡号, str密码, False, True, False, False, cllBalance, False, False, strXMLExpend, _
             str费用性质, objPati.病人ID, str结算方式, dbl结算金额) = False Then
        
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlBrushCard=false ")
        Exit Function
    End If
              
    Set objBalanceItems_out = New clsBalanceItems
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    For Each objItem In objOldBalanceItems
        '加入已经刷的消费卡
        objBalanceItems_out.AddItem objItem
        objBalanceItems_out.结算金额 = objBalanceItems_out.结算金额 + objItem.结算金额
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
        objBalanceItems_out.AddItem objItem
        objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
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
    '返回:扣款合法,返回=true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 16:42:43
    '说明:
    '   在调用扣款前，由于存在Oracle事务问题， 所以再调用扣款交易前， _
    '   先进行数据的合法性检查,以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = ""
    If gobjSquare.objOneCardComLib.zlPaymentCheck(mfrmMain, mlngModule, objCard.接口序号, objCard.消费卡, str卡号, dblMoney, strNos, strXMLExpend) = False Then
         Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlPaymentCheck=false ")
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

    If gobjSquare.objOneCardComLib.zlGetAccountMoney(mfrmMain, mlngModule, objCard.接口序号, _
     str卡号, strXMLExpend, dbl帐户余额_Out, objCard.消费卡) = False Then Exit Function
    If dbl帐户余额_Out <> 0 Then
        If dbl帐户余额_Out < dbl结算金额 Then
            MsgBox objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额不足!", vbInformation + vbOKOnly, gstrSysName
            
             Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", objCard.名称 & "(卡号为:" & objCard.zlCardNOEncrypt(str卡号) & ")的帐户余额(" & dbl帐户余额_Out & ")不足!")
            Exit Function
        End If
        For Each objItem In objBalanceItems_out
            objItem.帐户余额 = dbl帐户余额_Out
        Next
    End If
    zlSquare_Payment_IsValied = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "错误信息=" & mstrErrInfor)
End Function

Private Function zlSquareDel_CheckValied(ByVal objCard As Card, ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByVal dblDelMoney As Double, _
    ByVal rsMoney As ADODB.Recordset, ByRef objOldItems As clsBalanceItems, ByRef objBalanceItems_out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消费卡退费检查
    '入参:objCard-三方卡
    '     dblDelMoney-退款金额
    '     objOldItems-原结算对象(消费卡 ）
    '     rsMoney
    '出参:objBalanceItems_Out-返回刷卡对象
    '返回:交易合法返回=true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-23 11:07:58
    '说明:同步验证了接口和刷卡接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTotal As Double, i As Integer, j As Integer, blnFind As Boolean
    Dim cllBalance As Collection, cllSquareBalance As Collection
    Dim objItem As clsBalanceItem
    Dim strXmlIn As String
    
    
    
    On Error GoTo errHandle
    mstrFunName = "zlSquareDel_CheckValied"
    Call WriteLog(mlngModule, "", mstrFunName, "开始", "")
    
    If objCard.接口序号 <= 0 Or objCard.消费卡 = False Then
        zlSquareDel_CheckValied = True
         
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True", objCard.结算性质, objCard.结算方式, objCard.接口序号)
        Exit Function
    End If
    
    dblTotal = objOldItems.结算金额
    If RoundEx(dblTotal, 6) < RoundEx(dblDelMoney, 6) Then
        MsgBox "注意:" & vbCrLf & "   输入的退款金额大于了" & objCard.结算方式 & "的未退金额,请检查!" & vbCrLf & _
               "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
               "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
        
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "错误信息=输入的退款金额大于了" & objCard.结算方式 & "的未退金额,请检查!" & vbCrLf & _
               "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
               "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"))
        Exit Function
    End If
    
    If RoundEx(dblTotal, 6) <> RoundEx(dblDelMoney, 6) Then
        If objCard.是否全退 Then
            MsgBox "注意:" & vbCrLf & "   " & objCard.结算方式 & "必须全退,请检查!" & vbCrLf & _
                   "   未退金额:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
                   "   当前退款:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
                   
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "错误信息=输入的退款金额大于了" & objCard.结算方式 & "必须全退,请检查!" & vbCrLf & _
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
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objOneCardComLib.zlBrushCard(Me, mlngModule, rsMoney, objCard.接口序号, _
            objCard.消费卡, objPati.姓名, objPati.性别, _
            objPati.年龄, dblDelMoney, "", "", _
            True, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then
            
            
            Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "zlBrushCard=false")
            Exit Function
        End If
        Set cllBalance = cllSquareBalance
    End If
    Set objBalanceItems_out = New clsBalanceItem
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
                Call objBalanceItems_out.AddItem(objItem)
                objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
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
            Call objBalanceItems_out.AddItem(objItem)
            objBalanceItems_out.结算金额 = RoundEx(objBalanceItems_out.结算金额 + objItem.结算金额, 6)
        End If
    Next
    
    zlSquareDel_CheckValied = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call SaveErrLog
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)
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
    '返回:函数返回=true:调用成功,False:调用失败
    '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '''参数值  参数名  参数说明    应用场合
    ''1   刷卡和支付在同一页面：主要应用在自助系统中，可以通过此参数，调用界面的一些布局。  自助系统
    ''2   按交易流水号分别调用退款接口    主要是在退款时，如果此参数返回=true,则调用zlReturnMoney接口.否则调用returnMultMoney接口：病人结帐;补充结算 (含自助补允结算);病人预交款管理 (余额退款)
    
    Dim mstrFunName As String
    
    mstrFunName = "zlThirdSwapIsSwapNOCall"
   Call WriteLog(mlngModule, "", mstrFunName, "开始")
  
    zlThirdSwapIsSwapNOCall = gobjSquare.objOneCardComLib.ZlGetParaConfig(mfrmMain, lngCardTypeID, bln消费卡, 2, strErrMsg, strExpend)
    
   Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=" & IIf(zlThirdSwapIsSwapNOCall, "true", "false"))
End Function

Public Function zlDelBalanceItemFromItemObject(ByVal objItem As clsBalanceItem, Optional ByRef cllSaveSqL_Out As Collection, Optional blnExcute As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算对象,获取删除结算项的SQL
    '入参:blnExcute-是否进行执行操作
    '出参:cllSaveSqL_Out-返回保存的SQL
    '返回:成功返回=true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-14 17:13:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans As Boolean, strSQL As String
    
    
    
    On Error GoTo errHandle
    mstrFunName = "zlDelBalanceItemFromItemObject"
    Call WriteLog(mlngModule, "", mstrFunName, "开始")
    
    If cllSaveSqL_Out Is Nothing Then Set cllSaveSqL_Out = New Collection
    '明确失败
    '删除当前结算信息
    '    Zl_病人结帐结算_Delete
    strSQL = "Zl_病人结帐结算_Delete("
    '      结帐id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & IIf(objItem.冲销ID <> 0, objItem.冲销ID, objItem.结帐ID) & ","
    '      卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & objItem.卡类别ID & ","
    '      关联交易id_In 病人预交记录.Id%Type := Null,
    strSQL = strSQL & "" & objItem.关联交易ID & ")"
    '      结算方式_In Varchar2:=Null
    zlAddArray cllSaveSqL_Out, strSQL
    
    If Not blnExcute Then
        Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "blnExcute=false")
        Exit Function
    End If
    
    blnTrans = True
    zlExecuteProcedureArrAy cllSaveSqL_Out, "执行病人结算信息删除", , , "结帐", "病人结帐调试日志"
    blnTrans = True = False
    
    zlDelBalanceItemFromItemObject = True
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=True")
    Exit Function
errHandle:
    mstrErrInfor = Err.Number & "-" & Err.Description
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Call WriteLog(mlngModule, "", mstrFunName, "执行", "重试,错误信息=" & mstrErrInfor, 2)
        Resume
    End If
    Call WriteLog(mlngModule, "", mstrFunName, "结束", "返回=False", "Err=" & mstrErrInfor, 2)

End Function

Private Sub Class_Initialize()
    mstrOnCardPrivs = ";" & GetPrivFunc(glngSys, 1151) & ";"   '一卡通相关接口权限
    mstr脱机医保 = zlDatabase.GetPara(261, glngSys)
    Set mobjInsurePara = New clsInsurePara
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mfrmMain = Nothing
    Set mobjInsurePara = Nothing
    Set mobjPayCards = Nothing
End Sub









