Attribute VB_Name = "mdlPubExpence"
Option Explicit
'*********************************************************************************************************************************************
'功能:所涉费用公共部件的方法及调用(zlPublicExpense部件的处理)
'接口说明:
'    1. zlGetPubExpenseObject-获取费用公共部件对象
'    2. zlInitPriceGrade-初始化价格等级信息
'    3. zlGetPriceGrade:获取价格等级信息
'    4. zlPatiIdentify:病人身份验证(进行刷卡验证)
'    5. zlVerifyPassWord:输入密码验证框
'编制:刘兴洪
'日期:2019-01-25 09:51:46
'*********************************************************************************************************************************************
Public gobjPubExpense As Object  '费用公共部件
Public gintPriceGradeStartType As Integer
Public gstr药品价格等级 As String
Public gstr卫材价格等级 As String
Public gstr普通价格等级 As String
Public Function zlGetPubExpenseObject(ByRef objPubExpense As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取费用公共部件对象
    '出参:objPubExpense-返回公共费用部件对象
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-01-25 09:57:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If Not gobjPubExpense Is Nothing Then Set objPubExpense = gobjPubExpense: zlGetPubExpenseObject = True: Exit Function
    
    Err = 0: On Error Resume Next
    If gobjPubExpense Is Nothing Then
        Set gobjPubExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Err.Clear: On Error GoTo 0
            Exit Function
        End If
    End If
    
    Err.Clear:  On Error GoTo errHandle
    If gobjPubExpense Is Nothing Then Exit Function
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If gobjPubExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Function
    End If
    Set objPubExpense = gobjPubExpense: zlGetPubExpenseObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function zlInitPriceGrade() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化费用的价格等级
    '入参:
    '返回:初始化成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-01-25 10:00:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    
    gintPriceGradeStartType = gobjPubExpense.zlGetPriceGradeStartType()
    If gintPriceGradeStartType = 0 Then zlInitPriceGrade = True: Set objPubExpence = Nothing: Exit Function
    '读取站点价格等级
    Call gobjPubExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
    Set objPubExpence = Nothing:
    zlInitPriceGrade = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 
Public Function zlPatiIdentify(ByVal lngModlue As Long, ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal curMoney As Currency, _
    Optional ByVal bln退费 As Boolean = False, Optional ByVal bytDepositShowMode As Byte = 0, Optional ByVal lngDefaultCardTypeID As Long = 0, _
    Optional ByVal blnFamilyMoney As Boolean, Optional ByVal blnOlnyFamilyIDs As Boolean, Optional strFamilyPatiIDs_Out As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行刷卡验证
    '入参:lngModlue-模块号
    '     dblMoney-金额
    '     lng病人ID-病人ID
    '     bln退费-当前是否退费操作
    '     bytDepositShowMode- 预交显示方式(0-余额汇总显示;1-只显示门诊余额;2-只显示住院余额)
    '     lngDefaultCardTypeID-缺省的刷卡类别
    '     blnFamilyMoney-是否读取家属预交余额
    '     blnOlnyFamilyIDs-true:不验卡，只读取家属IDs;False-需要读取卡验卡(无效参数，保持兼容不予以删除)
    '出参:strFamilyPatiIDs-病人家属ID,多个用逗号分隔，79868
    '返回:身份验证成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-10-24 14:55:59
    '说明:
    '   一、消费验卡输入密码规则（bln退费=false时):
    '       1.不进行刷卡验证,直接返回True
    '       2.门诊消费时,需要进行刷卡验证,同时需要输入密码(无密码时,光标要经过密码框)
    '       3.门诊消费时,如果设置了密码(只要存在一张卡有密码的,就代表设置了密码的)，则必须刷卡且输入密码,无密码的,则只刷卡验证
    '       4.N元内免密支付,表示病人在消费N元内只刷卡验证,不输入密码;否则必须刷卡和输入密码
    '  二、退费验卡（bln退费=true时):
    '       1.不进行刷卡控制，直接返回true
    '       2.门诊退费时,需要进行刷卡验证,同时需要输入密码(无密码时,光标要经过密码框)
    '       3.门诊退费时,如果设置了密码(只要存在一张卡有密码的,就代表设置了密码的)，则必须刷卡且输入密码,无密码的,则只刷卡验证
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    
    zlPatiIdentify = objPubExpence.zlPatiIdentify(lngModlue, frmMain, lng病人ID, curMoney, bln退费, bytDepositShowMode, lngDefaultCardTypeID, _
                                               blnFamilyMoney, blnOlnyFamilyIDs, strFamilyPatiIDs_Out)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlGetPriceGrade(ByVal str站点 As String, _
    ByVal lng病人ID As Long, ByVal lng主页id As Long, _
    Optional ByVal str医疗付款方式 As String, _
    Optional ByRef str药品价格等级_Out As String, _
    Optional ByRef str卫材价格等级_Out As String, _
    Optional ByRef str普通项目价格等级_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 根据医疗付款方式或站点，获取对应的价格等级
    '入参:str站点-登陆的站点，必须传入，传入NULL时，价格等级为返回空
    '     lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     str医疗付款方式:如果传入非空，则以传的医疗付款方式_In方式来提取价格等级;否则以病人ID_In或主页ID来获取对应的病人的医疗付款方式。
    
    '出参:str药品价格等级_out-返回药品价格等级
    '     str卫材价格等级_out-返回卫材价格等级
    '     str普通项目价格等级_out-返回普通收费项目价格等级
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    zlGetPriceGrade = objPubExpence.zlGetPriceGrade(str站点, lng病人ID, lng主页id, str医疗付款方式, str药品价格等级_Out, str卫材价格等级_Out, str普通项目价格等级_out)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlVerifyPassWord(frmParent As Object, ByVal strPass As String, _
    Optional ByVal strName As String, Optional ByVal strSex As String, _
    Optional ByVal strOld As String, Optional blnPassEncode As Boolean = True) As Boolean
    '功能：对密码进行验证
    '参数：frmParent=显示的父窗体
    '      strPass=正确的密码
    '      strName,strSex,strOld=可选参数，病人姓名、性别、年龄，当不传入时不显示这个区域。
    '      blnPassEncode-strPass是否传入的加密串
    '返回：True=密码验证通过,False=取消输入，或连续3次输入错误的密码
    Dim objPubExpence As Object
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    zlVerifyPassWord = objPubExpence.zlVerifyPassWord(frmParent, strPass, strName, strSex, strOld, blnPassEncode)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function zlGetErrSwapInfoByJsonString(ByVal strJson As String, ByRef cllSwapInfo_out As Collection, ByRef cllExpends_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据异常交易信息的Json串，获取异常信息
    '入参:
    '出参:cllSwapInfo_out-返回的交易信息:卡号,卡类别ID,交易流水号,交易说明,交易金额,二维码,支付方式,结算摘要
    '     cllExpend_out
    '          |-cllExpend:-交易名称,交易内容
    '           格式:array(名称,值),"_名称"

    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    zlGetErrSwapInfoByJsonString = objPubExpence.zlGetErrSwapInfoByJsonString(strJson, cllSwapInfo_out, cllExpends_Out)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function
Public Function zlGetErrSwapInfoByErrID(ByVal lng异常ID As String, ByRef rsErrData_Out As ADODB.Recordset, _
    ByRef cllSwapInfo_out As Collection, ByRef cllExpends_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据异常ID获取获取异常信息
    '入参:
    '出参:
    '     rsErrData_Out-异常数据集
    '     cllSwapInfo_out-卡号,卡类别ID,交易流水号,交易说明,交易金额,二维码,支付方式,结算摘要
    '     cllExpend_out
    '          |-cllExpend:-交易名称,交易内容
    '           格式:array(名称,值),"_名称"
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    zlGetErrSwapInfoByErrID = objPubExpence.zlGetErrSwapInfoByErrID(lng异常ID, rsErrData_Out, cllSwapInfo_out, cllExpends_Out)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 

