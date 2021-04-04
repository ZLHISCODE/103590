Attribute VB_Name = "mdlIDKind"
Option Explicit
Public gobjCards As Cards    '所有的卡
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long, glngSys As Long
Public gstrAviPath As String, gstrVersion As String
Public gstrProductName As String
Public gstrDBUser As String   '当前数据库用户
Public gstrUnitName As String '用户单位名称
Public gobjParent As Object

'刷卡控制全局变量
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_F4 = vbKeyF4
Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20
Public glngInstanceCount As Long
Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Public gobjPubOneCard As clsPublicOneCard   '一卡通对象
Public gblnIsObjRegisterAlone As Boolean

Public Function zlGetPubOneCard(ByRef cnOracle As ADODB.Connection, ByRef objPubOneCard_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一卡通数据访问对象
    '入参:
    '出参:objOneDataObject_Out-返回一卡通数据访问对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-04 14:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    On Error GoTo errHandle
    
    If Not gobjPubOneCard Is Nothing Then Set objPubOneCard_Out = gobjPubOneCard: zlGetPubOneCard = True:  Exit Function
    Set objPubOneCard_Out = New clsPublicOneCard
    zlGetPubOneCard = objPubOneCard_Out.zlInitComponents(gobjParent, glngModul, glngSys, gstrDBUser, cnOracle, False, strExpend, gblnIsObjRegisterAlone)
    Set gobjPubOneCard = objPubOneCard_Out
    Exit Function
errHandle:
        If ErrCenter = 1 Then Resume
End Function

Public Function ErrCenter() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:错误处理中心
    '编制:刘兴洪
    '日期:2018-12-05 11:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(gcnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    ErrCenter = gobjPubOneCard.ErrCenter
End Function

Public Sub WritLog(ByVal strDev As String, strInput As String, strOutPut As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记录日志
    '编制:刘兴洪
    '日期:2018-12-05 11:35:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(gcnOracle, gobjPubOneCard) = False Then Exit Sub
    End If
    Call gobjPubOneCard.WritDebugLog(strDev, strInput, strOutPut)
End Sub

Public Function zlGetPatiIDFromBedNumber(ByVal lng病区ID As Long, ByVal str床号 As String, Optional ByRef lng主页ID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号获取病人ID
    '出参:lng主页ID-返回当前床号的主页ID
    '返回:成功返回病人ID,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(gcnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    zlGetPatiIDFromBedNumber = gobjPubOneCard.zlGetPatiIDFromBedNumber(lng病区ID, str床号, lng主页ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function zlInitComponents(Optional lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口部件
    '编制:刘兴洪
    '日期:2012-08-16 11:09:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpand As String
    strExpand = lngCardTypeID
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(gcnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    
    '初始化卡结算部件
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参: frmMain-调用的主窗体
    '        lngModule-HIS调用模块号
    '       lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       blnDeviceSet-设备设置调用初始化
    '       strExpand-扩展信息(可选:转入卡类别ID)
    zlInitComponents = gobjPubOneCard.zlInitComponents(gobjParent, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpand)
End Function


Public Function zlInitCards(ByVal cnOracle As ADODB.Connection, ByVal RegType As gRegType) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡对象
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-15 16:43:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln模糊查找 As Boolean, strValue As String, objCard As Card
    
    On Error GoTo errHandle
    
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(cnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    
    'zlGetCards(ByVal BytType As Byte) As Cards
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的卡对象
    '入参:bytType-0-所有医疗卡;
    '             1-启用的医疗卡,
    '             2-所有存在三方账户的三方卡
    '             3-启用的三方账户的医疗卡
    'Set rsTemp = gobjSquare.zlGetYLCards
    Set gobjCards = gobjPubOneCard.zlGetCards(0)
    
    bln模糊查找 = False
    For Each objCard In gobjCards
        Call GetRegInFor(RegType, "医疗卡类别\" & objCard.名称, "回车符", strValue)
        Select Case strValue
            Case "启用"
                objCard.卡号长度 = objCard.卡号长度 + IIf(objCard.设备是否启用回车, 0, 1)
            Case "禁用"
                objCard.卡号长度 = objCard.卡号长度 - IIf(objCard.设备是否启用回车, 1, 0)
        End Select
        If objCard.是否模糊查找 And objCard.启用 And Not bln模糊查找 Then bln模糊查找 = True
    Next
    gobjCards.按缺省卡查找 = Not bln模糊查找
    zlInitCards = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function
Public Function GetPatiInforFromPatiID(ByVal cnOracle As ADODB.Connection, ByVal lng病人ID As Long, ByRef objPati As clsPatiInfor, _
    ByRef strErrMsg As String, Optional strOtherName As String = "", Optional strOtherValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID,重新获取数据
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(cnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    GetPatiInforFromPatiID = gobjPubOneCard.zlGetPatiInforFromPatiID(lng病人ID, objPati, strErrMsg, strOtherName, strOtherValue)
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function
Public Function zlGetPatiInforFromXML(ByVal cnOracle As ADODB.Connection, ByVal strPatiXml As String, _
    ByRef objPatiInfor_Out As clsPatiInfor, ByRef strErrMsg_out As String) As Boolean
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
    On Error GoTo errHandle
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(cnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    zlGetPatiInforFromXML = gobjPubOneCard.zlGetPatiInforFromXML(strPatiXml, strErrMsg_out, objPatiInfor_Out)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiIDFromCardType(ByVal cnOracle As ADODB.Connection, ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean
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
    '出参:strErrMsg-返回的错误信息
    '       lng病人ID-返回的病人ID
    '       strCardPass-返回卡号的密码
    '       lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(cnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    If gobjPubOneCard.zlIsExistOraConnect = False Then Exit Function
    
    GetPatiIDFromCardType = gobjPubOneCard.zlGetPatiID(strCardType, strCardNo, blnNotShowErrMsg, lng病人ID, _
        strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, _
        blnCertificate, blnUserCancel, lngShowCardNoTypeID, blnNotCheckValidDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetKindCards(Optional strIDKindStr As String = "", Optional blnOnlyAccouct As Boolean = False, _
                                Optional NotAutoAppendKind As Boolean = False, Optional OnlyThreeCard As Boolean = False) As Cards
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的卡对象
    '返回: 成功,卡对象
    '编制:刘兴洪
    '日期:2012-08-15 16:58:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As Cards, objCard As Card
    Dim varData As Variant, i As Long, varTemp As Variant
    Dim blnFind As Boolean, j As Long
    Dim strKinds As String
    
    On Error GoTo errHandle
    If strIDKindStr = "" Then
        '缺省类别
        strIDKindStr = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0;手|手机号|0"
    End If
    Set objCards = New Cards
    varData = Split(strIDKindStr, ";")
    j = 1
    strKinds = ""
    If Not OnlyThreeCard Then
        For i = 0 To UBound(varData)
            '先找
            varTemp = Split(varData(i) & "||||||||||||", "|")
            If Trim(varTemp(1)) <> "" Then
                blnFind = False
                If Not gobjCards Is Nothing Then
                    For Each objCard In gobjCards
                        '76243,冉俊明,2014-8-5,其它开发组人员传入全角字母IC时,默认将其处理为系统中默认的IC卡类别
                        If objCard.名称 = Trim(varTemp(1)) _
                            Or (objCard.名称 Like "*IC卡*" And (varTemp(1) = "IC卡" Or varTemp(1) = "IC卡号" Or varTemp(1) Like "*ＩＣ卡*") And objCard.系统) _
                            Or (objCard.名称 Like "*身份证*" And (varTemp(1) = "二代身份证" Or varTemp(1) = "身份证" Or varTemp(1) = "身份证号") And objCard.系统) Then
                            blnFind = True
                            If InStr(strKinds & ",", "," & objCard.接口序号 & ",") = 0 Then
                                strKinds = strKinds & "," & objCard.接口序号
                                If objCard.启用 And Not objCard.消费卡 Then
                                   objCards.Add objCard, "K" & objCard.接口序号
                                End If
                            End If
                            Exit For
                        End If
                    Next
                End If
               If blnFind = False Then
                    '补充
                    Set objCard = New Card
                    '短名1|全名1|是否刷卡1|卡类别ID1|卡号长度1|缺省标志1(1-当前缺省;0-非缺省)|是否存在帐户1(1-存在帐户;0-不存在帐户)|
                    '卡号密文1(第几位至第几位加密,空为不加密)|是否扫描|是否接触式读卡|是否非接触式读卡
                    With objCard
                        .接口编码 = "-"
                        .名称 = varTemp(1)
                        .短名 = varTemp(0)
                        .是否刷卡 = Val(varTemp(2)) <> 1
                        .接口序号 = 0 ' IIf(Val(varTemp(3)) = 0, -j, Val(varTemp(3)))
                        .缺省标志 = Val(varTemp(4)) = 1
                        .是否存在帐户 = Val(varTemp(5)) = 1
                        .卡号密文规则 = Trim(varTemp(6))
                        '85565,李南春,2015/7/10:读卡性质，缺省为Fasle
                        .是否扫描 = Val(varTemp(7)) = 1
                        .是否接触式读卡 = Val(varTemp(8)) = 1
                        .是否非接触式读卡 = Val(varTemp(9)) = 1
                    End With
                    Err = 0: On Error Resume Next
                    objCards.Add objCard, "M" & objCard.名称
                    If Err <> 0 Then Err = 0: On Error GoTo 0
                    j = j + 1
               End If
            End If
        Next
    End If
    '未加入的，放入最后
    If NotAutoAppendKind = False Or OnlyThreeCard Then
        If Not gobjCards Is Nothing Then
            For Each objCard In gobjCards
                If InStr(1, strKinds & ",", "," & objCard.接口序号 & ",") = 0 And objCard.启用 And Not objCard.消费卡 Then
                    strKinds = strKinds & "," & objCard.接口序号
                    objCards.Add objCard, "K" & objCard.接口序号
                End If
            Next
        End If
    End If
    
    If Not gobjCards Is Nothing Then
        objCards.按缺省卡查找 = gobjCards.按缺省卡查找
        objCards.加密显示 = gobjCards.加密显示
    End If
    Set zlGetKindCards = objCards
    
    Err = 0: On Error Resume Next
    Erase varData '清空数组
    
    Exit Function
errHandle:
    
    MsgBox Err.Description
End Function
Public Function zlTranErrInfor(strErrMsg) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对错语信息进行格式化
    '返回: 返回被格式化的错语信息
    '编制:刘兴洪
    '日期:2012-08-22 14:47:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlTranErrInfor = strErrMsg
End Function
Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = Forms.Count = 0
End Function

Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:释放资源
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '实例数为0时，才放资源
    If glngInstanceCount > 0 Then Exit Function
    Call zlCloseWindows '释放窗体资源
    If Not gobjCards Is Nothing Then Set gobjCards = Nothing
    If Not gobjParent Is Nothing Then Set gobjParent = Nothing
    If Not gobjPubOneCard Is Nothing Then Set gobjPubOneCard = Nothing
    zlReleaseResources = True
End Function
