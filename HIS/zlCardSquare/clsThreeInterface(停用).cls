VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThreeInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*******************************************************************************************
'**功能说明:一卡通相关接口
'**编制:刘兴洪
'*******************************************************************************************
Private mobjCurCard  As clsCardObject '当前有效卡
Private mcnOracle As ADODB.Connection
Private mfrmMain As Object
Private mobjDatabase As clsDataBase
Private mlngModule As Long
Private mstrDBUser As String

Private mObjYLCards As clsCards
Private mObjYLCardObjs As clsCardObjects   '当前启用有效的医疗卡
Private mobjXML As clsXML
Private mint三方卡结帐退款控制 As Integer

Public Sub InitCommon(cnMain As ADODB.Connection)
    '------------------------------------------------
    '功能:初始化连接
    '参数：
    '   cnMain:主程序的数据库连接
    '------------------------------------------------
    Dim strFunName As String
    strFunName = "InitCommon"
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "cnMain is " & IIf(cnMain Is Nothing, "", "Not") & " Nothing")
    Set mcnOracle = cnMain
    Call mobjDatabase.InitCommon(cnMain)
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "")
End Sub


Public Function zlInitInterFacel(ByVal lngCardTypeID As Long, _
    Optional bln消费卡 As Boolean = False, Optional blnAdviceSend As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化指定卡接口
    '入参:lngCardTypeID-指定卡类别
    '       bln消费卡-是否消费卡
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 15:29:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard, objCardObject As Object, strKey As String, strExpand As String
    Dim blnOnlyNotObject As Boolean
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    Dim strFunName As String
    
    strFunName = "zlInitInterFacel"
    
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "lngCardTypeID=" & lngCardTypeID & ",bln消费卡=" & bln消费卡 & ",blnAdviceSend=" & blnAdviceSend)
    If Not mobjCurCard Is Nothing Then
        If mobjCurCard.接口序号 = lngCardTypeID And mobjCurCard.消费卡 = bln消费卡 Then
            If Not mobjCurCard.InitCompents Then
                If mobjCurCard.CardObject Is Nothing Then
                    blnOnlyNotObject = True: GoTo GoCreateObject:
                End If
                If Not mobjCurCard.CardObject.zlInitComponents(mfrmMain, mlngModule, glngSys, mstrDBUser, gcnOracle, False, strExpand) Then                    '初始化部件
                    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,调用初始化接口(失败):卡类别ID=" & mobjCurCard.接口序号)
                    Exit Function
                End If
                mobjCurCard.InitCompents = True
            End If
            zlInitInterFacel = True
            Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true,初始化接口(返回成功):卡类别ID=" & lngCardTypeID)
            Exit Function
        End If
    End If
    Err = 0: On Error Resume Next
GoCreateObject:
    strKey = IIf(bln消费卡, "X", "K") & lngCardTypeID
    
    '检查设备是否启用
    '59760
    If GetCards_YL(objYLCards) = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,获取一卡通卡集失败(GetCards_YL),卡类别ID=" & lngCardTypeID & "", 2)
        Exit Function
    End If
    If GetCardObjs_YL(objYlCardObjs) = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,获取一卡通卡集失败(GetCardObjs_YL),卡类别ID=" & lngCardTypeID & "", 2)
        Exit Function
    End If
    
    Set mobjCurCard = objYlCardObjs(strKey)
    If Err <> 0 Then
            Err = 0: On Error Resume Next
            Set objCard = objYLCards.Item(strKey)
            If Err <> 0 Then
                ShowMsgbox "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!"
                Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,部件(" & strKey & "):" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!", 2)
                Exit Function
            End If
            '重新创建
            If zlCreatePatiCardObject(objCard, objCardObject, blnAdviceSend) = False Then
                Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,部件(" & strKey & ")，创新新控件失败", 2)
                Exit Function
            End If
            '增加对应
           Set mobjCurCard = objYlCardObjs.Add(objCardObject, objCard.自制卡, objCard.接口序号, objCard, bln消费卡, strKey)
    End If
    
    If mobjCurCard Is Nothing Then
    
        If Not objCard Is Nothing Then
                MsgBox "注意:" & vbCrLf & "调用接口(" & objCard.接口编码 & "-" & objCard.名称 & ")调用失败,请检查!", vbInformation, gstrSysName
                Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,调用接口(" & objCard.接口编码 & "-" & objCard.名称 & ")调用失败", 2)
        Else
                MsgBox "注意:" & vbCrLf & "调用接口(" & lngCardTypeID & ")调用失败,请检查!", vbInformation, gstrSysName
                Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,调用接口(" & lngCardTypeID & ")调用失败", 2)
        End If
        Exit Function
    End If
    
    Err = 0: On Error Resume Next
    Set objCard = mobjCurCard.CardPreporty
    If Err <> 0 Then
        ShowMsgbox "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description, 2)
        Exit Function
    End If
    If Not mobjCurCard.InitCompents Then
        If Not mobjCurCard.CardObject.zlInitComponents(mfrmMain, mlngModule, glngSys, mstrDBUser, gcnOracle, False, strExpand) Then
            Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,调用初始化接口(失败):卡类别ID=" & mobjCurCard.接口序号)
            Exit Function
        End If
        mobjCurCard.InitCompents = True
    End If
    zlInitInterFacel = True
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true,卡类别ID=" & lngCardTypeID)
End Function

Public Function zlInitComponents(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    ByVal cnOracle As ADODB.Connection, _
    Optional blnDeviceSet As Boolean = False, _
    Optional strExpand As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参: frmMain-调用的主窗体
    '        lngModule-HIS调用模块号
    '       lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       blnDeviceSet-设备设置调用初始化
    '       strExpand-扩展信息(可选传入:卡类别ID-不传时,表示全部初始化,传入时,只初始化指定的接口)
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
    Dim objYLCards As clsCards
    Dim objCardObject As clsCardObject, objYlCardObjs As clsCardObjects
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlInitComponents"
    
    On Error GoTo errHandle
    Set mcnOracle = cnOracle: mlngModule = lngModule: mstrDBUser = strDBUser
    
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "cnOracle is " & IIf(cnOracle Is Nothing, "", "Not") & " Nothing,strExpand=" & strExpand & ",strDBUser=" & strDBUser)
    
    Call InitCards    '初始化卡对象
   
    If GetCards_YL(objYLCards) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,获取一卡通对象集(GetCards_YL)失败！", 2)
        Exit Function
    End If
    If GetCardObjs_YL(objYlCardObjs) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,获取一卡通对象集(GetCardObjs_YL)失败！", 2)
        Exit Function
    End If
    
    '调用其他初始化部件
    For Each objCardObject In objYlCardObjs
        If Not objCardObject.CardObject Is Nothing Then
            If objCardObject.CardObject.zlInitComponents(frmMain, lngModule, lngSys, strDBUser, gcnOracle, blnDeviceSet, strExpand) Then
                '初始化部件
                objCardObject.InitCompents = True
                Call zlWritLogEx(Me, lngModule, strFunName, "执行", "一卡通初始化接口(成功)：卡类别ID=" & objCardObject.接口序号)
            Else
                objCardObject.InitCompents = False
                Call zlWritLogEx(Me, lngModule, strFunName, "执行", "一卡通初始化接口(失败)：卡类别ID=" & objCardObject.接口序号)
            End If
        End If
    Next
    Set mfrmMain = frmMain
    zlInitComponents = True
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=true,初始化所有部件接口")
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "初始化所有部件接口(发生错误),发生重试,错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,初始化所有部件接口(发生错误),错误信息:" & strErrInfor, 2)

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
    Dim strFunName As String
    
    strFunName = "zlCardDevSet"
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "bytCardType=" & bytCardType)
    
    Call zlInitPatiCards(mcnOracle)
    frmCardSelect.SelectCard frmMain, bytCardType
    zlCardDevSet = True

    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & zlCardDevSet)

End Function

Public Sub SetEnabled(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置Enabled属性
    '入参:blnEnabled-true,表示自动获取;False表示不自动获取
    '编制:刘兴洪
    '日期:2012-08-22 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardObject As clsCardObject
    Dim objYlCardObjs As clsCardObjects
    Dim strFunName As String
     
    strFunName = "SetEnabled"
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "blnEnabled=" & blnEnabled)
    
    '59760
    If GetCardObjs_YL(objYlCardObjs) = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,GetCardObjs_YL(失败)")
        Exit Sub
    End If
    If objYlCardObjs Is Nothing Then Exit Sub
     For Each objCardObject In objYlCardObjs
        If Not objCardObject.CardObject Is Nothing And objCardObject.InitCompents Then
            Err = 0: On Error Resume Next
            objCardObject.CardObject.SetEnabled (blnEnabled)
            If Err <> 0 Then Err = 0: On Error GoTo 0
        End If
    Next
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "")
End Sub

Public Sub zlInitEvents(ByVal hWnd As Long, ByVal objComEvents As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置公共事件
    '入参:objComEvents-公共事件
    '编制:刘兴洪
    '日期:2012-08-22 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardObject As clsCardObject
    Dim objYlCardObjs As clsCardObjects
    
    Dim strFunName As String
    
    strFunName = "zlInitEvents"
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "hWnd=" & hWnd)
    
    '59760
    If GetCardObjs_YL(objYlCardObjs) = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,GetCardObjs_YL(失败)")
        Exit Sub
    End If
    
    If objYlCardObjs Is Nothing Then Exit Sub
     For Each objCardObject In objYlCardObjs
        If Not objCardObject.CardObject Is Nothing And objCardObject.InitCompents Then
            Err = 0: On Error Resume Next
            Set objCardObject.CardObject.zlInitEvents = objComEvents
            Call objCardObject.CardObject.SetParent(hWnd)
            If Err <> 0 Then Err = 0: On Error GoTo 0
        End If
    Next
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "")
End Sub

Public Function zlReadCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal blnOlnyCardNO As Boolean, _
    ByVal strExpand As String, _
    ByRef strOutCardNo As String, _
    ByRef strOutclsPatientInfoXml As String, _
    Optional ByRef strPhotoFile As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读卡接口
    '入参:frmMain-调用的父窗口
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       blnOlnyCardNO-仅仅读取卡号
    '       strExpand-扩展参数:内部暂无用，但传入接口部件的格式（见下面代码中固定传入的格式）：  用|分离,第一栏为卡类别ID:lngCardTypeID,后续待扩展
    '出参:strOutCardNO-返回的卡号
    '       strOutclsPatientInfoXML-(病人信息返回.XML串)
    '       strPhotoFile-当从卡中读取取像片后,传给HIS的像片文件(含完整路径)
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 11:56:41
    '说明:
    '   １. 在所有绑定卡时,需要建立病人信息档案时，读取该接口
    '   ２. 在所有需要刷卡进行就诊的地方，都需要调用该接口
    'strOutclsPatientInfoXML参数说明
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
    Dim blnInitFace As Boolean, blnReadCard As Boolean
    Dim strPara As String
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlReadCard"
    
    On Error GoTo errHandle
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " blnOlnyCardNO=" & IIf(blnOlnyCardNO, "True", "False")
    strPara = strPara & vbTab & "strExpand=" & strExpand
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,初始化接口(失败),卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    '传入:strExpand-用|分离,第一栏为卡类别ID:lngCardTypeID
    strExpand = lngCardTypeID
    Err = 0: On Error Resume Next
    '56615
    blnReadCard = mobjCurCard.CardObject.zlReadCard(frmMain, lngModule, blnOlnyCardNO, strExpand, strOutCardNo, strOutclsPatientInfoXml, strPhotoFile)
    If Err <> 0 Then
        If Err <> 450 Then GoTo errHandle:
        '450-错误的参数号或无效的属性赋值
        '主要是歉容以前的
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用读卡接口(兼容接口,无像片文件返回),卡类别ID=" & lngCardTypeID)
        Err = 0: On Error GoTo errHandle:
        blnReadCard = mobjCurCard.CardObject.zlReadCard(frmMain, lngModule, blnOlnyCardNO, strExpand, strOutCardNo, strOutclsPatientInfoXml)
    End If
    
    strPara = ""
    If blnReadCard Then
        strPara = "strOutCardNo=" & strOutCardNo & vbTab & "strPhotoFile=" & strPhotoFile
        strPara = strPara & vbTab & "strOutclsPatientInfoXml=" & strOutclsPatientInfoXml
    End If
    
    zlReadCard = blnReadCard
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & IIf(blnReadCard, "True", "False") & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "重试,调用调卡接口失败,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & Err.Description, 2)
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
    Dim blnSucces As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlSetBrushCardObject"
    
    On Error GoTo errHandle
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "lngCardTypeID=" & lngCardTypeID & ",strExpend=" & strExpend & ",bln消费卡=" & bln消费卡)
     
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,初始化接口部件失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    Err = 0: On Error Resume Next
    blnSucces = mobjCurCard.CardObject.zlSetBrushCardObject(lngCardTypeID, objBrushCard, strExpend)
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
    End If
    zlSetBrushCardObject = blnSucces
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & blnSucces & ",卡类别ID=" & lngCardTypeID & ",strExpend=" & strExpend)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description

    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
End Function


Public Function zlMakeCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal stInputclsPatientInfoXml As String, _
    ByRef strPhotoFile As String, _
    ByRef strOutclsPatientInfoXml As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:制卡接口
    '入参:
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  卡类别ID
    '    stInputclsPatientInfoXml String(XML) In  HIS传入已经健档的病人信息
    '    strPhotoFile    String  In  表示像片文件(含完整路径)
    '    strExpend   String  In   XML,暂无,待以后扩展
    '出参:
    '    strOutclsPatientInfoXml  String(XML) Out 调用接口后,传给HIS系统已经制卡后的病人信息,以便建立病人档案
    '    strExpend   String  Out  XML,暂无,待以后扩展
    '    strPhotoFile    String  In  表示像片文件(含完整路径)
    '返回 :True:调用成功,False:调用失败
    '调用者:
    '    目前只在医疗卡发放管理中调用此接口(主要是考虑输入的信息较多,其他窗口不适应)，相关的调用流程医疗卡发放管理的制卡
    '  strInputclsPatientInfoXML及strOutclsPatientInfoXML格式
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
    Dim blnInitFace As Boolean
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlMakeCard"
    
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID
    strPara = strPara & vbTab & "stInputclsPatientInfoXml=" & stInputclsPatientInfoXml
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败")
        Exit Function
    End If
    
    '56615
    On Error Resume Next
    blnSucceed = mobjCurCard.CardObject.zlMakeCard(frmMain, lngModule, lngCardTypeID, stInputclsPatientInfoXml, _
        strPhotoFile, strOutclsPatientInfoXml, strExpend)
        
    If Err = 438 Then
        MsgBox mobjCurCard.CardPreporty.名称 & "没有编写制卡接口(zlMakeCard),制卡失败!", vbInformation, gstrSysName
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", mobjCurCard.CardPreporty.名称 & "没有编写制卡接口(zlMakeCard),制卡失败!")
        Err.Clear
    ElseIf Err <> 0 Then
        GoTo errHandle
    End If
    
    strPara = ""
    If blnSucceed Then
        strPara = "strPhotoFile=" & strPhotoFile & vbTab & "strOutclsPatientInfoXml=" & strOutclsPatientInfoXml
        strPara = strPara & vbTab & "strOutclsPatientInfoXml=" & strOutclsPatientInfoXml
        strPara = strPara & vbCrLf & "strExpend=" & strExpend
    End If
    zlMakeCard = blnSucceed
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
    
End Function

Public Function zlAdviceSendBulidBarCode(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal intRecNature As Integer, ByVal strNOs As String, Optional ByVal strExpand As String) As Boolean
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
    Dim objCards As clsCards, objCard As clsCard
    Dim blnOK As Boolean
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlAdviceSendBulidBarCode"
    
    On Error GoTo errHandle
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " intRecNature=" & intRecNature & vbTab & "strNos=" & strNOs
    strPara = strPara & vbTab & "strExpand=" & strExpand
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    
    If zlGetCards_YL(objCards) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,获取一卡通对象集失败(zlGetCards_YL)")
        Exit Function
    End If
    
    blnOK = False
    If lngCardTypeID = 0 Then
        For Each objCard In objCards
            If objCard.接口序号 <> 0 And objCard.接口程序名 <> "" And objCard.消费卡 = False And objCard.发送调用接口 = True Then
                If zlInitInterFacel(objCard.接口序号, objCard.消费卡, True) Then
                    Err = 0: On Error Resume Next
                    If mobjCurCard.CardObject.zlAdviceSendBulidBarCode(frmMain, lngModule, objCard.接口序号, intRecNature, strNOs, strExpand) = True Then
                        blnOK = True
                        Err = 0: On Error GoTo errHandle:
                        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "接口返回=true,zlAdviceSendBulidBarCode(结束),卡类别ID=" & objCard.接口序号 & "")
                    Else
                        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "接口返回=False,zlAdviceSendBulidBarCode(结束),卡类别ID=" & objCard.接口序号 & "")
                    End If
                Else
                    Call zlWritLogEx(Me, lngModule, strFunName, "执行", "接口返回=False,调用初始化接口失败,卡类别ID=" & objCard.接口序号)
                End If
            End If
        Next
    Else
        For Each objCard In objCards
            If objCard.接口序号 = lngCardTypeID And objCard.接口程序名 <> "" And objCard.消费卡 = False And objCard.发送调用接口 = True Then
                If zlInitInterFacel(objCard.接口序号, objCard.消费卡, True) Then
                    Err = 0: On Error Resume Next
                    If mobjCurCard.CardObject.zlAdviceSendBulidBarCode(frmMain, lngModule, objCard.接口序号, intRecNature, strNOs, strExpand) = True Then
                        blnOK = True
                        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "接口返回1=true,zlAdviceSendBulidBarCode(结束),卡类别ID=" & objCard.接口序号)
                    Else
                        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "接口返回1=False,zlAdviceSendBulidBarCode(结束),卡类别ID=" & objCard.接口序号)
                    End If
                    Err = 0: On Error GoTo errHandle:
                Else
                    Call zlWritLogEx(Me, lngModule, strFunName, "执行", "接口返回1=False,调用初始化接口失败,卡类别ID=" & objCard.接口序号)
                End If
            End If
        Next
    End If
    zlAdviceSendBulidBarCode = blnOK
    '140931:李南春,2019/5/14,结束时不用传卡类别ID
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnOK & ",strExpand=" & strExpand)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
    
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
    Dim blnInitFace As Boolean
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlBandCardArfter"
    
    On Error GoTo errHandle
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " lng病人ID=" & lng病人ID
    strPara = strPara & vbTab & "strExpand=" & strExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
 
    Err = 0: On Error Resume Next
    '56615
    blnSucceed = mobjCurCard.CardObject.zlBandCardArfter(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, strExpend)
    
    If Err <> 0 Then
        If Err.Number <> 438 Then GoTo errHandle:
        '438-对象不支持该属性或方法
        '歉容以前的,直接返回true
        Err = 0
        blnSucceed = True
    Else
        strPara = ""
        If blnSucceed Then
            strPara = "strExpend=" & strExpend
        End If
    End If
    zlBandCardArfter = blnSucceed
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description

    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & Err.Description, 2)
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
    Optional ByRef blnShowclsPatientInfo As Boolean = False, _
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
    Dim frmInput As frmInputPass
    Dim blnErrCount As Boolean
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlBrushCard"
    
    On Error GoTo errHandle
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " rsClassMoney=" & IIf(rsClassMoney Is Nothing, "Nothing", "Not Nothing")
    strPara = strPara & vbTab & "bln消费卡=" & IIf(bln消费卡, "True", "False") & vbTab & "strPatiName=" & strPatiName & vbTab & "strSex=" & strSex & vbTab & "strOld=" & strOld
    strPara = strPara & vbTab & "dbl金额=" & dbl金额 & vbTab & "bln余额不足禁止=" & IIf(bln余额不足禁止, "True", "False") & "bln转预交=" & IIf(bln转预交, "True", "False") & vbTab & "blnAllPay=" & IIf(blnAllPay, "True", "False")
    strPara = strPara & vbTab & "str费用来源=" & str费用来源 & vbTab & "lng病人ID=" & lng病人ID
    strPara = strPara & vbTab & vbCrLf & "strXMLExpend=" & strXMLExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    
    dbl结算金额_Out = 0
    str结算方式_Out = ""
    Set mfrmMain = frmMain
    If bln转预交 = False And lngModule = 1103 Then bln转预交 = True
    If lngCardTypeID = 0 Then
        '暂不支持
        Set frmInput = New frmInputPass
        Set varSquareBalance = Nothing
        zlBrushCard = frmInput.zlBrushPay(frmMain, lngModule, Nothing, rsClassMoney, lngCardTypeID, _
            bln消费卡, strPatiName, strSex, strOld, dbl金额, strCardNo, strPassWord, bln退费, _
            blnShowclsPatientInfo, bln退现, bln余额不足禁止, varSquareBalance, bln转预交, blnAllPay, str费用来源, lng病人ID)
        Set varSquareBalance = Nothing
        Unload frmInput
        Set frmInput = Nothing
         Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlBrushCard(结束)_1,卡类别ID=" & lngCardTypeID & ",返回=失败")
        Exit Function
    End If
    
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    If bln消费卡 And mobjCurCard.自制卡 Then
    
        Set frmInput = New frmInputPass
        Err = 0: On Error Resume Next
        If IsEmpty(varSquareBalance) Then   '57682
            Set varSquareBalance = Nothing
        End If
        blnErrCount = varSquareBalance.Count
        If Err <> 0 Then
            Set varSquareBalance = Nothing
            Err = 0: On Error GoTo 0
        End If
        zlBrushCard = frmInput.zlBrushPay(frmMain, lngModule, mobjCurCard, rsClassMoney, _
            lngCardTypeID, bln消费卡, strPatiName, strSex, strOld, dbl金额, _
            strCardNo, strPassWord, bln退费, blnShowclsPatientInfo, bln退现, bln余额不足禁止, varSquareBalance, _
            bln转预交, blnAllPay, str费用来源, lng病人ID)
        Unload frmInput
        Set frmInput = Nothing
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlBrushCard & ",消费卡,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    '    zlBrushCard(frmMain As Object, _
    '        ByVal lngModule As Long, _
    '        ByVal lngCardTypeID As Long, _
    '        ByVal strPatiName As String, ByVal strSex As String, _
    '        ByVal strOld As String, ByVal dbl金额 As Double, _
    '        Optional ByRef strCardNo As String, _
    '        Optional ByRef strPassWord As String
    Set varSquareBalance = Nothing

    Err = 0: On Error Resume Next
    blnSucceed = mobjCurCard.CardObject.zlBrushCard(frmMain, lngModule, lngCardTypeID, _
     strPatiName, strSex, strOld, dbl金额, strCardNo, strPassWord, strXMLExpend)
     
    If Err = 450 Then
        Err = 0
        On Error GoTo errHandle
        blnSucceed = mobjCurCard.CardObject.zlBrushCard(frmMain, lngModule, lngCardTypeID, _
        strPatiName, strSex, strOld, dbl金额, strCardNo, strPassWord)
        strXMLExpend = ""
    ElseIf Err <> 0 Then
        GoTo errHandle
    End If
    
    strPara = ""
    If blnSucceed Then
        strPara = "strCardNo=" & strCardNo & vbTab & "bln退费=" & IIf(bln退费, "true", "False")
        strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
    End If
    
    If strXMLExpend <> "" And InStr(1, UCase(strXMLExpend), "<OUTPUT>") > 0 Then
        Dim rsXML As ADODB.Recordset
        '打开XML文件
        If mobjXML.OpenXMLDocument(strXMLExpend) = False Then
            MsgBox "接口返回了无效的XML格式，不允许刷卡操作 ", vbInformation + vbOKOnly, gstrSysName
            Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,出参:" & vbCrLf & "接口返回了无效的XML格式")
            zlBrushCard = False
            Exit Function
        End If
        If mobjXML.GetMultiNodeRecord("OUTPUT/JS", rsXML) Then
            If Not rsXML Is Nothing Then
                rsXML.Filter = "Parent_id=1"
                Do While Not rsXML.EOF
                    If NVL(rsXML!node_name) = "JSFS" Then str结算方式_Out = NVL(rsXML!node_value)
                    If NVL(rsXML!node_name) = "JSJE" Then dbl结算金额_Out = Val(NVL(rsXML!node_value))
                    rsXML.MoveNext
                Loop
            End If
            If str结算方式_Out <> "" Then
                zlBrushCard = blnSucceed
                Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",存在接点OUTPUT/JS,卡类别ID=" & lngCardTypeID & ",str结算方式_Out=" & str结算方式_Out & ",dbl结算金额_Out=" & dbl结算金额_Out & "," & strPara)
                Exit Function
            End If
        End If
    End If
    zlBrushCard = blnSucceed
    Dim objCard As Card
    Set objCard = zlGetCard(lngCardTypeID, bln消费卡)
    If Not objCard Is Nothing Then
        str结算方式_Out = objCard.结算方式
        dbl结算金额_Out = dbl金额
    End If
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",卡类别ID=" & lngCardTypeID & ",str结算方式_Out=" & str结算方式_Out & ",dbl结算金额_Out=" & dbl结算金额_Out & "," & strPara)
    
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
End Function

Public Function zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
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
    
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strXMLExpendTemp As String
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlPaymentCheck"
    
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False")
    strPara = strPara & vbTab & "strCardNo=" & strCardNo & vbTab & "dblMoney=" & dblMoney & vbTab & "strNos=" & strNOs
    strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    strXMLExpendTemp = strXMLExpend
    blnSucceed = mobjCurCard.CardObject.zlPaymentCheck(frmMain, lngModule, lngCardTypeID, strCardNo, dblMoney, strNOs, strXMLExpend)
    If strXMLExpendTemp = strXMLExpend Then strXMLExpend = ""
    
    strPara = ""
    If blnSucceed Then
        strPara = "strXMLExpend=" & strXMLExpend
    End If
    zlPaymentCheck = blnSucceed
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
    
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
    '        strXMLExpend:10.35.110后才支持
    '        <OUTPUT>
    '            <TXZT>退现状态<TXZT>//接口返回True时：  表示是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但允许操作员操作退现;
    '                                  接口返回False时： 表示是否允许强制退现:1-表示允许通过"强制退现"权限来控制是否允许退现,0-表示禁止通过"强制退现"权限来退现
    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
    '                                               1.不允许返回三卡方的结算方式
    '                                               2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
    '        </OUTPUT>
    '返回:退现合法,返回true,否则返回Flase
    '编制:刘尔旋
    '日期:2017-7-3
    '说明:
    '    1) 收费退费时，在进入退费界面时按原交易调用一次；在"确定"时，再调一次。
    '    2)其他:在退款信息加载前调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDefaultCash As Boolean '以医疗卡类别的缺省设置为准
    Dim strValue As String, objCard As Card
    Dim blnCash As Boolean
    Dim blnErr As Boolean
    Dim strPara As String
    Dim strFunName As String, strErrInfor As String
     
    strFunName = "zlReturnCashCheck"
    
    On Error GoTo errHandle
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " strCardNo=" & strCardNo
    strPara = strPara & vbTab & "strBalanceIDs=" & strBalanceIDs & vbTab & "dblMoney=" & dblMoney & vbTab & " strSwapNo=" & strSwapNo
    strPara = strPara & vbTab & "strSwapMemo=" & strSwapMemo
    strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    blnDelCash_Out = False: strDefaultDelBalance_Out = ""
    Set mfrmMain = frmMain
    
    blnErr = False
    Set objCard = zlGetCard(lngCardTypeID, False)
    If Not objCard Is Nothing Then
        If objCard.接口程序名 = "" Then blnErr = True
    End If
    
    If blnErr = False Then
         If zlInitInterFacel(lngCardTypeID, False) Then
            On Error Resume Next
            blnCash = mobjCurCard.CardObject.zlReturnCashCheck(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strXMLExpend)
            If Err = 0 Then
                strPara = "strXMLExpend=" & strXMLExpend
                Call zlWritLogEx(Me, lngModule, strFunName, "执行", "mobjCurCard.CardObject.zlReturnCashCheck=" & blnCash & ",出参:" & vbCrLf & strPara)
            End If
            blnErr = Err <> 0
        Else
            blnErr = True
            Call zlWritLogEx(Me, lngModule, strFunName, "执行", "调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        End If
    End If
    
    If blnErr Or InStr(strXMLExpend, "<OUTPUT>") = 0 Then
        '无此接口
        Err = 0: On Error GoTo errHandle:
        If objCard Is Nothing Then
            MsgBox "在获取退款方式时，卡类别Id=" & lngCardTypeID & "获取失败!", vbOKOnly + vbInformation, gstrSysName
            zlReturnCashCheck = False '只能原样返回
            strPara = "blnDelCash_Out=" & IIf(blnDelCash_Out, "true", "False") & vbTab & "strDefaultDelBalance_Out=" & strDefaultDelBalance_Out
            Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,在获取退款方式时，卡类别Id=" & lngCardTypeID & "," & strPara, 2)
            Exit Function
        End If
        If lngModule = 1137 Then
            If mint三方卡结帐退款控制 = -1 Then mint三方卡结帐退款控制 = Val(zlDatabase.GetPara("三方卡结帐退款控制", glngSys, lngModule))
            '针对结帐，为了兼容，因此，需直接根据参数判断:0-三方帐户结帐退款时缺省退回帐户;1-三方帐户结帐退款缺省退回现金(缺省结算方式)
            If blnErr Then blnCash = objCard.是否退现
            If blnCash Then
                blnDelCash_Out = mint三方卡结帐退款控制 = 1 Or objCard.是否缺省退现
            Else
                blnDelCash_Out = True
            End If
            
            zlReturnCashCheck = blnCash
            
            strPara = "blnDelCash_Out=" & IIf(blnDelCash_Out, "true", "False") & vbTab & "strDefaultDelBalance_Out=" & strDefaultDelBalance_Out
            Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnCash & ",出参:" & strPara)
            Exit Function
        End If
        
        If blnErr Then blnCash = objCard.是否退现
        If blnCash Then
            blnDelCash_Out = objCard.是否缺省退现
        Else
            blnDelCash_Out = True
        End If
        zlReturnCashCheck = blnCash
        Set objCard = Nothing
        
        strPara = "blnDelCash_Out=" & IIf(blnDelCash_Out, "true", "False") & vbTab & "strDefaultDelBalance_Out=" & strDefaultDelBalance_Out
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnCash & ",出参:" & strPara)
        Exit Function
    End If
    
    zlReturnCashCheck = blnCash
    
    '打开XML文件
    If mobjXML.OpenXMLDocument(strXMLExpend) = False Then
        MsgBox "接口返回（zlReturnCashCheck）的XML无效，请与接口开发商或系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
        blnDelCash_Out = True
        zlReturnCashCheck = False
        
        strPara = "blnDelCash_Out=" & IIf(blnDelCash_Out, "true", "False") & vbTab & "strDefaultDelBalance_Out=" & strDefaultDelBalance_Out
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlReturnCashCheck & ",出参:" & strPara)
        Exit Function
    End If
    
    strValue = ""
    If mobjXML.GetSingleNodeValue("OUTPUT/TXZT", strValue) Then
         blnDelCash_Out = Val(strValue) = 1
    End If
    
    strValue = ""
    Call mobjXML.GetSingleNodeValue("OUTPUT/QSTKFS", strValue) '缺省退现方式
    strDefaultDelBalance_Out = strValue
    
    strPara = "blnDelCash_Out=" & IIf(blnDelCash_Out, "true", "False") & vbTab & "strDefaultDelBalance_Out=" & strDefaultDelBalance_Out
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlReturnCashCheck & ",出参:" & strPara)
    
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlReturnCashCheck & ",卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)

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
    '    strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
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
    Dim varData As Variant
    Dim strPara As String, strXMLExpendTemp As String
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlPaymentMoney"
    
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain
    
    strXMLExpendTemp = strXMLExpend '先备份
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False")
    strPara = strPara & vbTab & "strCardNo=" & strCardNo & vbTab & "strBalanceIDs=" & strBalanceIDs & vbTab & " strPrepayNos=" & strPrepayNos
    strPara = strPara & vbTab & "dblMoney=" & dblMoney & vbTab & "strBalanceIDs=" & strBalanceIDs & vbTab & " strPrepayNos=" & strPrepayNos
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
     
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
        
    intStatus_Out = 0: strErrMsg_Out = ""
    Set rsBalance_Out = Nothing: Set rsExpend_Out = Nothing
    
    If mobjCurCard.CardObject.zlPaymentMoney(frmMain, lngModule, _
        lngCardTypeID, strCardNo, strBalanceIDs, strPrepayNos, dblMoney, strSwapGlideNO, strSwapMemo, strXMLExpend) = False Then
        If strXMLExpendTemp = strXMLExpend Then strXMLExpend = ""
        
        If InStr(strXMLExpend, "<OUTPUT>") > 0 Then
            intStatus_Out = GetStatusFromXML(strXMLExpend, strErrMsg_Out)
        End If
        
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",intStatus_Out=" & intStatus_Out & ",strXMLExpend=" & strXMLExpend)
        Exit Function
    Else
        If strXMLExpendTemp = strXMLExpend Then strXMLExpend = ""
        
        strPara = "strSwapGlideNO=" & strSwapGlideNO & vbTab & "strSwapMemo=" & strSwapMemo
        strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "出参:" & strPara)
    End If
    
    zlPaymentMoney = True
    If blnResolveXMLToRecord Then
        '需要产生相关的结算数据集
        If InStr(strXMLExpend, "<OUTPUT>") > 0 Then
            If zlGetBalaceDataRecordFromXML(lngCardTypeID, bln消费卡, strXMLExpend, rsBalance_Out, True, rsExpend_Out, , intStatus_Out, strErrMsg_Out) = False Then
                Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlPaymentMoney(结束),卡类别ID=" & lngCardTypeID & ",接口返回=成功,XML错误")
                Exit Function
            End If
        ElseIf strXMLExpend <> "" Then
            If zlGetExpendDataRecordFromXML(lngCardTypeID, bln消费卡, strXMLExpend, rsExpend_Out) = False Then
                Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlPaymentMoney(结束),卡类别ID=" & lngCardTypeID & ",接口返回=成功,扩展信息错误")
                Exit Function
            End If
        End If
    End If
    strPara = "strSwapGlideNO=" & strSwapGlideNO & vbTab & "strSwapMemo=" & strSwapMemo
    strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
    
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回＝true,卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlPaymentMoney & ",卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
    
End Function
Private Function GetStatusFromXML(ByVal strXML As String, strErrMsg_Out As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据XML获取当前操作状态
    '入参:strxml-XML串
    '     格式:...
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-07-13 18:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim strPara As String
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "GetStatusFromXML"
    
    On Error GoTo errHandle
    
    strErrMsg_Out = ""
    strPara = "strXML=" & strXML
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    '打开XML文件
    If mobjXML.OpenXMLDocument(strXML) = False Then
        MsgBox "接口返回了无效的XML格式，不允许刷卡操作 ", vbInformation + vbOKOnly, gstrSysName
        GetStatusFromXML = 0
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=0,接口返回了无效的XML格式")
        Exit Function
    End If
    Call mobjXML.GetSingleNodeValue("OUTPUT/JYZT", strValue)
    GetStatusFromXML = Val(strValue)
    Call mobjXML.GetSingleNodeValue("OUTPUT/ERRMSG", strErrMsg_Out)
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & GetStatusFromXML)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=0,错误信息=" & strErrInfor, 2)
End Function

Public Function zlGetExpendDataRecordFromXML(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strXML As String, ByRef rsExpends_Out As ADODB.Recordset, Optional blnNotOpenXML As Boolean = False) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据XML获取扩展数据集
    '入参:strXML-XML文本或字符串(格式:项目名称1|项目内容2||…||项目名称n|项目内容n)
    '     blnNotOpenXML-true:-不打开XML对象;False-需要打开XML对象
    '     <OUTPUT>
    '            ....
    '            <Expends>               扩展信息列表
    '                <Expend>
    '                    <XMMC>项目名称</XMMC> //Varchar2  50
    '                    <XMNR>项目内容</XMNR> //Varchar2  100
    '                </Expend>
    '                ...
    '            </Expends>
    '        </OUTPUT>
    '出参:rsExpends_Out-返回数据集
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-27 11:12:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, strPreParent_id As String
    Dim i As Long, rsXML As ADODB.Recordset
    Dim blnUpdate As Boolean
    Dim strPara As String
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlGetExpendDataRecordFromXML"
    
    On Error GoTo errHandle
    
    strPara = "lngCardTypeID=" & lngCardTypeID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False")
    strPara = strPara & vbTab & "strXML=" & strXML
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    
    ':项目名称1|项目内容2||…||项目名称n|项目内容n
    Set rsExpends_Out = New ADODB.Recordset
    With rsExpends_Out
        .Fields.Append "卡类别ID", adBigInt
        .Fields.Append "消费卡", adBigInt
        .Fields.Append "项目名称", adVarChar, 50
        .Fields.Append "项目内容", adVarChar, 500
        .Open
    End With
    
    If InStr(strXML, "<OUTPUT>") = 0 Then
        varData = Split(strXML, "||")
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & "|", "|")
            If Trim(varTemp(0)) <> "" Then
                rsExpends_Out.AddNew
                    rsExpends_Out!卡类别ID = lngCardTypeID
                    rsExpends_Out!消费卡 = IIf(bln消费卡, 1, 0)
                    rsExpends_Out!项目名称 = Trim(varTemp(0))
                    rsExpends_Out!项目内容 = Trim(varTemp(1))
                rsExpends_Out.Update
            End If
        Next
        
       zlGetExpendDataRecordFromXML = True
       Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true,rsExpends_Out=" & rsExpends_Out.RecordCount)
       Exit Function
    End If
    
    If Not blnNotOpenXML Then
        '打开XML文件
        If mobjXML.OpenXMLDocument(strXML) = False Then
            MsgBox "接口返回了无效的XML格式，不允许刷卡操作 ", vbInformation + vbOKOnly, gstrSysName
            zlGetExpendDataRecordFromXML = False
            Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False, 无效的XML格式, 返回= 失败")
            Exit Function
        End If
    End If
    
    If Not mobjXML.GetMultiNodeRecord("OUTPUT/Expends/Expend", rsXML) Then
        If Not mobjXML.GetMultiNodeRecord(UCase("OUTPUT/Expends/Expend"), rsXML) = False Then
            If Not blnNotOpenXML Then mobjXML.CloseXMLDocument
            Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False, 解析XML出错, 返回= 失败")
            Exit Function
        End If
    End If
    
    If rsXML Is Nothing Then
        If Not mobjXML.GetMultiNodeRecord(UCase("OUTPUT/Expends/Expend"), rsXML) = False Then
            If Not blnNotOpenXML Then mobjXML.CloseXMLDocument
            Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False, 解析扩展信息出错, 返回= 失败")
            Exit Function
        End If
    End If
    rsXML.Filter = "parent_id<>''"
    rsXML.Sort = "parent_id"
    If rsXML.RecordCount <> 0 Then rsXML.MoveFirst
    strPreParent_id = ""
    With rsExpends_Out
        Do While Not rsXML.EOF
            If strPreParent_id <> rsXML!parent_id Then
                strPreParent_id = NVL(rsXML!parent_id, "")
                .AddNew
                !卡类别ID = lngCardTypeID
                !消费卡 = IIf(bln消费卡, 1, 0)
            End If
            blnUpdate = False
            If NVL(rsXML!node_name) = "XMMC" Then !项目名称 = NVL(rsXML!node_value): blnUpdate = True
            If NVL(rsXML!node_name) = "XMNR" Then !项目内容 = NVL(rsXML!node_value): blnUpdate = True
            If blnUpdate = True Then .Update
            rsXML.MoveNext
        Loop
    End With
    If Not blnNotOpenXML Then mobjXML.CloseXMLDocument
    zlGetExpendDataRecordFromXML = True
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true")
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,错误信息=" & strErrInfor, 2)
End Function



Public Function zlGetBalaceDataRecordFromXML(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, _
    ByVal strXML As String, ByRef rsBalance_Out As ADODB.Recordset, _
    Optional blnExpend As Boolean, Optional ByRef rsExpends_Out As ADODB.Recordset, Optional bln退款 As Boolean = False, _
    Optional ByRef intStatus_Out As Integer, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据XML获取结算信息集()
    '入参:strXml-XML串,格式如下;
    '        <OUTPUT>
    '            <JYLIST>//交易列表
    '                <JY>
    '                    <JYFS>医保基金</JYFS>
    '                    <JYJE>10</JYJE>
    '                    <JYLSH>YBJY0001</JYLSH>
    '                    <JYSM>医保结算</JYSM>
    '                    <DJH>NO00012</DJH>
    '                    <SFPTJS>0</SFPTJS>
    '                    <KH>卡号</KH>
    '                </JY>
    '                ....
    '            </JYLIST>
    '            <Expends>扩展信息列表
    '                <Expend>
    '                    <XMMC>项目名称1</XMMC>//Varchar250
    '                    <XMNR>项目内容1</XMNR>//Varchar2100
    '                </Expend>
    '                ....
    '            </Expends>
    '        </OUTPUT>
    '        或退款列表(bln退款交易=true时)
    '            <OUTPUT>
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
    '     blnExpend-就否获取扩展信息集
    '     bln退款-表示按退款的XML格式进行处理
    '出参:rsBalance_Out-返回结算信息集
    '    rsExpends_Out_Out-扩展信息集(blnExpend=true时返回)
    '    intStatus_Out-交易状态
    '    strErrMsg_Out-返回的错误信息
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-26 18:38:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsXML As ADODB.Recordset
    Dim strPreParent_id As String, strValue  As String
    Dim blnUpdate As Boolean, strNodePath As String
    Dim strPara As String
    Dim strFunName As String, strErrInfor As String
     
    strFunName = "zlGetBalaceDataRecordFromXML"
    
    On Error GoTo errHandle
    
    strPara = "lngCardTypeID=" & lngCardTypeID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False")
    strPara = strPara & vbTab & "strXML=" & strXML
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    
    Set rsBalance_Out = New ADODB.Recordset
    With rsBalance_Out
        .Fields.Append "卡类别ID", adBigInt
        .Fields.Append "消费卡", adBigInt
        .Fields.Append "结算方式", adVarChar, 50
        .Fields.Append "交易金额", adDouble
        .Fields.Append "交易流水号", adVarChar, 50
        .Fields.Append "交易说明", adVarChar, 500
        .Fields.Append "原交易流水号", adVarChar, 50
        .Fields.Append "原交易说明", adVarChar, 500
        .Fields.Append "结算号码", adVarChar, 50
        .Fields.Append "结算摘要", adVarChar, 100
        .Fields.Append "卡号", adVarChar, 100
        .Fields.Append "单据号", adVarChar, 20
        .Fields.Append "单据性质", adBigInt
        .Fields.Append "是否普通结算", adBigInt
        .Fields.Append "关联交易ID", adBigInt
        .Fields.Append "ID", adBigInt
        
        .Open
    End With
    
    '打开XML文件
    If mobjXML.OpenXMLDocument(strXML) = False Then
        If bln退款 Then
            MsgBox "三方退款接口返回了无效的XML格式，可能HIS系统中的退款数据与三方的退款不一致，请与接口开发放或系统管理员联系! ", vbInformation + vbOKOnly, gstrSysName
            
            zlGetBalaceDataRecordFromXML = True
            Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true,无效的XML格式(可能HIS系统中的退款数据与三方的退款不一致),返回失败")
            Exit Function
        Else
            MsgBox "接口返回了无效的XML格式，不允许刷卡操作 ", vbInformation + vbOKOnly, gstrSysName
        End If
        zlGetBalaceDataRecordFromXML = False
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,无效的XML格式,返回失败")
        Exit Function
    End If
    
    Call mobjXML.GetSingleNodeValue("OUTPUT/JYZT", strValue)
    intStatus_Out = Val(strValue)
    Call mobjXML.GetSingleNodeValue("OUTPUT/ERRMSG", strErrMsg_Out)
    
    strNodePath = IIf(bln退款, "OUTPUT/TKLIST/TK", "OUTPUT/JYLIST/JY")
    
    If Not mobjXML.GetMultiNodeRecord(strNodePath, rsXML) Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,GetMultiNodeRecord返回False,strNodePath=" & strNodePath)
        Exit Function
    End If
    If rsXML Is Nothing Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,rsXML=nothing")
        Exit Function
    End If
    
    
    rsXML.Filter = "parent_id<>''"
    rsXML.Sort = "parent_id"
    If rsXML.RecordCount <> 0 Then rsXML.MoveFirst
    strPreParent_id = ""
    With rsBalance_Out
        Do While Not rsXML.EOF
            If strPreParent_id <> rsXML!parent_id Then
                strPreParent_id = NVL(rsXML!parent_id, "")
                .AddNew
                !卡类别ID = lngCardTypeID
                !消费卡 = IIf(bln消费卡, 1, 0)
            End If
            blnUpdate = False
            Select Case NVL(rsXML!node_name)
            Case "JYFS", "TKFS"
                !结算方式 = NVL(rsXML!node_value): blnUpdate = True
            Case "JYJE", "TKJE"
                !交易金额 = Val(NVL(rsXML!node_value)): blnUpdate = True
            Case "JYLSH"
                !交易流水号 = NVL(rsXML!node_value): blnUpdate = True
            Case "JYSM"
                !交易说明 = NVL(rsXML!node_value): blnUpdate = True
            Case "YJYLSH"
                !原交易流水号 = NVL(rsXML!node_value): blnUpdate = True
            Case "YJYSM"
                !原交易说明 = NVL(rsXML!node_value): blnUpdate = True
            Case "JSHM"
                !结算号码 = NVL(rsXML!node_value): blnUpdate = True
            Case "GLJYID"
                !关联交易id = NVL(rsXML!node_value): blnUpdate = True
            Case "JSZY"
                !结算摘要 = NVL(rsXML!node_value): blnUpdate = True
            Case "KH"
                !卡号 = NVL(rsXML!node_value): blnUpdate = True
            Case "DJH"
                !单据号 = NVL(rsXML!node_value): blnUpdate = True
            Case "DJXZ"
                !单据性质 = Val(NVL(rsXML!node_value)): blnUpdate = True
            Case "SFPTJS"
                !是否普通结算 = Val(NVL(rsXML!node_value)): blnUpdate = True
            Case "ID"
                !id = Val(NVL(rsXML!node_value)): blnUpdate = True
            End Select
            If blnUpdate = True Then .Update
            rsXML.MoveNext
        Loop
    End With
    
    Set rsExpends_Out = Nothing
    If Not blnExpend Then
        mobjXML.CloseXMLDocument: zlGetBalaceDataRecordFromXML = True
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true,卡类别ID=" & lngCardTypeID & ",blnExpend=" & blnExpend)
        Exit Function
    End If
    
    If zlGetExpendDataRecordFromXML(lngCardTypeID, bln消费卡, strXML, rsExpends_Out, True) = False Then
        mobjXML.CloseXMLDocument:
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=false,zlGetExpendDataRecordFromXML=false")
        Exit Function
    End If
    
    mobjXML.CloseXMLDocument
    zlGetBalaceDataRecordFromXML = True
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true,卡类别ID=" & lngCardTypeID)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
End Function

Public Function zlGetCard(ByVal lngCardTypeID As Long, Optional bln消费卡 As Boolean) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡对象
    '入参:lngCardTypeID-卡类别Id
    '     bln消费卡-消费卡
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-03-19 14:37:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard, strKey As String
    Dim strFunName As String
    
    strFunName = "zlGetCard"
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "lngCardTypeID=" & lngCardTypeID & ",bln消费卡=" & bln消费卡)

    If mObjYLCards Is Nothing Or gObjYLCards Is Nothing Or Not mObjYLCards Is gObjYLCards Then
        Call zlInitPatiCards(mcnOracle): Call InitCards
    End If
    Err = 0: On Error Resume Next
    
    strKey = IIf(bln消费卡, "X", "K") & lngCardTypeID
    Set objCard = mObjYLCards(strKey)
    
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
        Set zlGetCard = Nothing
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=Nothing ")
        Exit Function
    End If
    
    Set zlGetCard = FromClsCardToCard(objCard)
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "zlGetCard=not Nothing")
    
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
    '                </TFITEM >
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
    '编制:刘兴洪
    '日期:2011-05-26 17:24:55
    '说明:
    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
    '    以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlReturncheck"
    
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False")
    strPara = strPara & vbTab & "strCardNo=" & strCardNo & vbTab & "strBalanceIDs=" & strBalanceIDs & vbTab & "dblMoney=" & dblMoney
    strPara = strPara & vbTab & "strSwapNo=" & strSwapNo & vbTab & "strSwapMemo=" & strSwapMemo
    strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    blnSucceed = mobjCurCard.CardObject.zlReturncheck(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strXMLExpend)
    zlReturncheck = blnSucceed
    
    strPara = ""
    If blnSucceed Then
        strPara = "strXMLExpend=" & strXMLExpend
    End If
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlReturncheck,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
End Function


Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
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
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlReturnMoney"
    
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False")
    strPara = strPara & vbTab & "strCardNo=" & strCardNo & vbTab & "strBalanceIDs=" & strBalanceIDs & vbTab & "dblMoney=" & dblMoney & vbTab & "strSwapGlideNO=" & strSwapGlideNO
    strPara = strPara & vbTab & "strSwapMemo=" & strSwapMemo & vbTab & "strSwapExpendInfor=" & strSwapExpendInfor
    strPara = strPara & vbTab & "strExpendXML=" & strExpendXML
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If

     
    Err = 0: On Error Resume Next
    blnSucceed = mobjCurCard.CardObject.zlReturnMoney(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExpendInfor, strExpendXML)
    If Err.Number = 450 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "参数号不对,卡类别ID=" & lngCardTypeID)
        '错误的参数号
        Err = 0: On Error GoTo errHandle:
        strExpendXML = ""
        blnSucceed = mobjCurCard.CardObject.zlReturnMoney(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExpendInfor)
    End If
    zlReturnMoney = blnSucceed
    
    If Err.Number <> 0 Then GoTo errHandle:
    
    If blnResolveXMLToRecord Then
        '需要产生相关的结算数据集
        If InStr(strExpendXML, "<OUTPUT>") > 0 Then
            If zlGetBalaceDataRecordFromXML(lngCardTypeID, bln消费卡, strExpendXML, rsBalance_Out, False, rsExpend_Out, True, intStatus_Out, strErrMsg_Out) = False Then
                Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlReturnMoney & ",解析扩展信息(XML)错误")
                Exit Function
            End If
        End If
        If strSwapExpendInfor <> "" Then
            If zlGetExpendDataRecordFromXML(lngCardTypeID, bln消费卡, strSwapExpendInfor, rsExpend_Out) = False Then
                Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlReturnMoney & ",解析扩展信息错误,strSwapExpendInfor=" & strSwapExpendInfor)
                Exit Function
            End If
        End If
    End If
    
    strPara = ""
    If blnSucceed Then
        strPara = "strSwapGlideNO=" & strSwapGlideNO & vbCrLf & vbTab & "strExpendXML=" & strExpendXML
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlReturnMoney & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)

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
    '                    <JYLSH>交易流水号</JYLX> //原交易流水号
    '                    <JYSM>交易说明</JYSM>//原交易说明
    '                    <ZFJE>作废金额</ZFJE>
    '                    <JSLX>类型</JSLX>  //1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款，6-保险补充结算
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
    Dim strXML As String
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlReturnMultiMoney"
    
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain
 
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False")
    strPara = strPara & vbTab & "lng冲销ID=" & lng冲销ID & vbTab & "strExpend=" & strExpend
    strPara = strPara & vbCrLf & "strInXML=" & strInXML
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
     
    
    If Not mobjCurCard.CardObject.zlReturnMultiMoney(frmMain, lngModule, lngCardTypeID, bln消费卡, strInXML, lng冲销ID, strOutXml, strExpend) Then
        strPara = ""
        If InStr(strOutXml, "<OUTPUT>") > 0 Then
            intStatus_Out = GetStatusFromXML(strOutXml, strErrMsg_Out)
            strPara = "strOutXml=" & strOutXml
        End If

        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & "," & strPara)
        Exit Function
    End If
    
    
    zlReturnMultiMoney = True
    
    strPara = "strOutXml=" & strOutXml & vbCrLf & "strExpend=" & strExpend
    
    strErrMsg_Out = "": intStatus_Out = 0
    If blnResolveXMLToRecord And strOutXml <> "" Then
        '需要产生相关的结算数据集
        strXML = Replace(strOutXml, "<JSLIST>", "<TKLIST>")
        strXML = Replace(strXML, "</JSLIST>", "</TKLIST>")
        
        strXML = Replace(strXML, "<JS>", "<TK>")
        strXML = Replace(strXML, "</JS>", "</TK>")
        
        strXML = Replace(strXML, "<TKJE>", "<JYJE>")
        strXML = Replace(strXML, "</TKJE>", "</JYJE>")
        
        strXML = Replace(strXML, "<TKLSH>", "<JYLSH>")
        strXML = Replace(strXML, "</TKLSH>", "</JYLSH>")
        
        strXML = Replace(strXML, "<TKSM>", "<JYSM>")
        strXML = Replace(strXML, "</TKSM>", "</JYSM>")
        
        '兼容处理
        strExpend = Replace(strExpend, "EXPENDS", "Expends")
        strExpend = Replace(strExpend, "EXPEND", "Expend")
        
        If InStr(strXML, "<OUTPUT>") = 0 Then
            strXML = "<OUTPUT>" & vbCrLf & strXML & vbCrLf & strExpend & vbCrLf & "</OUTPUT>"
        Else
            If strExpend <> "" And InStr(strExpend, "<Expends>") > 0 Then
                strXML = Replace(strXML, "</OUTPUT>", "")
                strXML = strXML & vbCrLf & strExpend
                strXML = strXML & vbCrLf & "</OUTPUT>"
            End If
        End If
        
        If zlGetBalaceDataRecordFromXML(lngCardTypeID, bln消费卡, strXML, rsBalance_Out, True, rsExpend_Out, True, intStatus_Out, strErrMsg_Out) = False Then
            strPara = strPara & vbTab & "intStatus_Out=" & intStatus_Out & vbTab & "strErrMsg_Out=" & strErrMsg_Out
            Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=true,zlGetBalaceDataRecordFromXML=false,卡类别ID=" & lngCardTypeID & "," & strPara)
            Exit Function
        End If
    End If
    
    strPara = strPara & vbTab & "intStatus_Out=" & intStatus_Out & vbTab & "strErrMsg_Out=" & strErrMsg_Out
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=true,卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlReturnMultiMoney & ",卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
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
    Dim blnInitFace As Boolean
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlMzInforWriteToCard"
    
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " lng病人ID=" & lng病人ID & vbTab & " lngBalanceID=" & lngBalanceID
    strPara = strPara & vbTab & "strExpend=" & strExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    If lngCardTypeID <> 0 Then
        lngCardTypeID = Val(ZLGetPatiCardFromCards(CStr(lngCardTypeID), lng病人ID))
    End If
    
    If lngCardTypeID = 0 Then
        '需要读取相关的卡类别ID
        If zlSelectWriteCardType(frmMain, lngCardTypeID, mcnOracle, lng病人ID) = False Then
            Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlSelectWriteCardType选择返回False")
            Exit Function
        End If
        If lngCardTypeID = 0 Then
            Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlSelectWriteCardType选择返回False,卡类别ID=0")
            Exit Function
        End If
    End If
    
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    '56615
    Err = 0: On Error Resume Next
    blnSucceed = mobjCurCard.CardObject.zlMzInforWriteToCard(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, lngBalanceID, strExpend)
    If Err = 438 Then
        'VB实时错误438：对象不支持该属性或方法
        MsgBox mobjCurCard.CardPreporty.名称 & "没有编写写卡接口(zlMzInforWriteToCard)，门诊信息写卡失败！", vbInformation, gstrSysName
        Err.Clear
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", mobjCurCard.CardPreporty.名称 & "没有编写写卡接口(zlMzInforWriteToCard)，门诊信息写卡失败！")
    ElseIf Err <> 0 Then
        GoTo errHandle
    End If
    
    zlMzInforWriteToCard = blnSucceed
    strPara = ""
    If blnSucceed Then
        strPara = "strExpend=" & strExpend
    End If
          
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlMzInforWriteToCard & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
End Function

Public Function zlZyInforWriteToCard(frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, ByVal lng病人ID As Long, _
    ByVal lngBalanceID As Long, Optional ByRef strExpend As String = "") As Boolean
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
    Dim blnInitFace As Boolean
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlZyInforWriteToCard"
    
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " lng病人ID=" & lng病人ID & vbTab & " lngBalanceID=" & lngBalanceID
    strPara = strPara & vbTab & "strExpend=" & strExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    
    If lngCardTypeID <> 0 Then
        lngCardTypeID = Val(ZLGetPatiCardFromCards(CStr(lngCardTypeID), lng病人ID))
    End If
    If lngCardTypeID = 0 Then
        '需要读取相关的卡类别ID
        If zlSelectWriteCardType(frmMain, lngCardTypeID, mcnOracle, lng病人ID) = False Then
            Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlSelectWriteCardType,返回False")
            Exit Function
        End If
        If lngCardTypeID = 0 Then
            Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlZyInforWriteToCard,卡类别ID=0,返回False")
            Exit Function
        End If
    End If

    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
 
    '56615
    Err = 0: On Error Resume Next
    blnSucceed = mobjCurCard.CardObject.zlZyInforWriteToCard(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, lngBalanceID, strExpend)
    If Err = 438 Then
        'VB实时错误438：对象不支持该属性或方法
        MsgBox mobjCurCard.CardPreporty.名称 & "没有编写写卡接口(zlZyInforWriteToCard)，住院信息写卡失败！", vbInformation, gstrSysName
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", mobjCurCard.CardPreporty.名称 & "没有编写写卡接口(zlZyInforWriteToCard)，住院信息写卡失败！", 2)
        Err.Clear
    ElseIf Err <> 0 Then
        GoTo errHandle
    End If
    
    zlZyInforWriteToCard = blnSucceed
    
    strPara = ""
    If blnSucceed Then
        strPara = "strExpend=" & strExpend
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlZyInforWriteToCard & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlZyInforWriteToCard,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
End Function

Public Function zlHealthArchiveIsSHow(frmMain As Object, ByVal lngModule As Long, ByRef strFunNameOper As String, Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:居民健档档案
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    strFunName  String  Out 返回功能名, 即在医生工作站的菜单中显示的名称(菜单位置放在"工具"菜单下)
    '    strFunNameOper   String  In/Out  XML,暂无,待以后扩展
    '        Boolean 函数返回    True:表示应该加入档案显示菜单,False:不加入档案菜单
    '调用者:
    '    1.  门诊(住院)医生工作站调用,医生站
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard, lngCardTypeID As Long
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlHealthArchiveIsSHow"
    
    On Error GoTo errHandle
    
    strPara = "lngModule=" & lngModule & vbTab & "strFunName=" & strFunName
    strPara = strPara & vbTab & "strExpend=" & strExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & strPara)
    
    '59760
    If GetCards_YL(objYLCards) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,获取一卡通对象集(GetCards_YL)失败")
        Exit Function
    End If
    
    If GetCardObjs_YL(objYlCardObjs) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,获取一卡通对象集(GetCardObjs_YL)失败")
    End If
    
    Dim i As Long
    Set mfrmMain = frmMain
    strFunName = "健康卡"
    For i = 1 To objYlCardObjs.Count
        Set objCard = objYlCardObjs(i).CardPreporty
        If objCard.启用 And objCard.名称 = "居民健康卡" And objCard.消费卡 = False Then
            '只取一个
            lngCardTypeID = objCard.接口序号: Exit For
        End If
    Next
    If lngCardTypeID = 0 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,lngCardTypeID=0")
        Exit Function
    End If
    
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    '56615
    blnSucceed = mobjCurCard.CardObject.zlHealthArchiveIsSHow(frmMain, lngModule, strFunNameOper, strExpend)
    zlHealthArchiveIsSHow = blnSucceed
    
    strPara = ""
    If blnSucceed Then
        strPara = "strExpend=" & strExpend
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",zlHealthArchiveIsSHow(结束),卡类别ID=" & lngCardTypeID & "," & strPara)
    
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlHealthArchiveIsSHow & ",卡类别ID=" & lngCardTypeID & ",错误信息=" & Err.Description, 2)
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
    Dim objCard As clsCard, lngCardTypeID As Long
    Dim blnShow As Boolean, i As Long
    Dim frmShowSelect As frmHealthArchives
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    Dim strPara As String
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlHealthArchivesShow"
    On Error GoTo errHandle
    
    strPara = "lngModule=" & lngModule & vbTab & "lng病人ID=" & lng病人ID
    strPara = strPara & vbTab & "strExpend=" & strExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
 
    '59760
    If GetCards_YL(objYLCards) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,获取一卡通对象集(GetCards_YL)失败")
        Exit Function
    End If
    If GetCardObjs_YL(objYlCardObjs) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,获取一卡通对象集(GetCardObjs_YL)失败")
        Exit Function
    End If
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    For i = 1 To objYlCardObjs.Count
        Set objCard = objYlCardObjs(i).CardPreporty
        If objCard.启用 And objCard.名称 = "居民健康卡" And objCard.消费卡 = False Then
            '只取一个
            lngCardTypeID = objCard.接口序号: Exit For
        End If
    Next
    If lngCardTypeID = 0 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=0")
        Exit Function
    End If
    
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    '56615
    blnShow = mobjCurCard.CardObject.zlHealthArchivesShow(frmMain, lngModule, lng病人ID, strExpend)
    If lng病人ID = 0 Then Exit Function
    If Not blnShow Then
        
        Set frmShowSelect = New frmHealthArchives
        Call frmShowSelect.zlShowHealthArchives(frmMain, lng病人ID, mcnOracle)
        If Not frmShowSelect Is Nothing Then Unload frmShowSelect
        Set frmShowSelect = Nothing
        Call zlWritLogEx(Me, lngModule, strFunName, "执行", "zlHealthArchivesShow(显示)")
    End If
    zlHealthArchivesShow = True
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=true,卡类别ID=" & lngCardTypeID)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
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
    Dim strPara As String
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlGetAccountMoney"
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False") & vbTab & " strCardNo=" & strCardNo
    strPara = strPara & vbTab & "strExpand=" & strExpand
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & strPara)
    
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    zlGetAccountMoney = mobjCurCard.CardObject.zlGetAccountMoney(frmMain, lngModule, lngCardTypeID, strCardNo, strExpand, dblMoney)
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlGetAccountMoney & ",dblMoney=" & dblMoney)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,zlGetAccountMoney,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
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
    '       strErrMsg_Out-返回的错误信息(从XML中解析出来的错误信息)
    '返回:检查的数据合法,返回True:否则返回False
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlTransferAccountsCheck"
    
    On Error GoTo errHandle
    Set mfrmMain = frmMain: strErrMsg_Out = ""
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " strCardNo=" & strCardNo & vbTab & " dblMoney=" & dblMoney & vbTab & " strBalanceID=" & strBalanceID
    strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    
    blnSucceed = mobjCurCard.CardObject.zlTransferAccountsCheck(frmMain, lngModule, lngCardTypeID, strCardNo, dblMoney, strBalanceID, strXMLExpend)
    zlTransferAccountsCheck = blnSucceed
    
    If strXMLExpend = "" Or InStr(strXMLExpend, "<OUT>") = 0 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ", " & strXMLExpend)
        Exit Function
    End If
    
    '打开XML文件
    If mobjXML.OpenXMLDocument(strXMLExpend) = False Then
        MsgBox "接口返回（zlTransferAccountsCheck）的XML无效，请与接口开发商或系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
        zlTransferAccountsCheck = False
        
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False(XML无效),返回=失败")
        Exit Function
    End If
    Call mobjXML.GetSingleNodeValue("OUT/ERRMSG", strErrMsg_Out)  '是否缺省退现
    
    strPara = ""
    If blnSucceed Then
        strPara = "strXMLExpend=" & strXMLExpend
    End If
    
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlTransferAccountsCheck & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    zlTransferAccountsCheck = False
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
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
    '   strErrMsg_Out-XML返回的错误信息
    '   rsExpend_Out-返回扩展信息集
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPreXMLExpend As String, strValue As String
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlTransferAccountsMoney"
    On Error GoTo errHandle
    
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " strCardNo=" & strCardNo & vbTab & " strBalanceID=" & strBalanceID & vbTab & " dblMoney=" & dblMoney
    strPara = strPara & vbCrLf & "strSwapExtendInfor=" & strSwapExtendInfor & vbCrLf & "strXMLExpend=" & strXMLExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    strSwapExtendInfor = strSwapExtendInfor
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    
    blnSucceed = mobjCurCard.CardObject.zlTransferAccountsMoney(frmMain, lngModule, _
    lngCardTypeID, strCardNo, strBalanceID, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend)
    
    zlTransferAccountsMoney = blnSucceed
    
    If strSwapExtendInfor = strPreXMLExpend Then strSwapExtendInfor = ""
    Call zlGetExpendDataRecordFromXML(lngCardTypeID, False, strSwapExtendInfor, rsExpend_Out)
    
    If strXMLExpend = "" Or InStr(strXMLExpend, "<OUT>") = 0 Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & blnSucceed & ",strXMLExpend=" & strXMLExpend)
        Exit Function
    End If
    
    strPara = ""
    If blnSucceed Then
        strPara = "strSwapGlideNO=" & strSwapGlideNO & vbTab & "strSwapMemo=" & strSwapMemo
        strPara = strPara & vbTab & "strSwapExtendInfor=" & strSwapExtendInfor & vbCrLf & " strXMLExpend=" & strXMLExpend
   End If
    '打开XML文件
    If mobjXML.OpenXMLDocument(strXMLExpend) = False Then
        MsgBox "接口返回（zlTransferAccountsCheck）的XML无效，请与接口开发商或系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
        zlTransferAccountsMoney = False
        
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,XML无效,strXMLExpend=" & strXMLExpend)
        Exit Function
    End If
    
    strValue = ""
    Call mobjXML.GetSingleNodeValue("OUT/JYZT", strValue)  '交易状态   接口返回False时，此参数有效:交易状态: 0-交易调用失败;1-交易正在处理中
    intStatus_Out = Val(strValue)
    Call mobjXML.GetSingleNodeValue("OUT/ERRMSG", strErrMsg_Out)  '是否缺省退现
    
    strPara = strPara & vbTab & "intStatus_Out=" & ",strErrMsg_Out=" & strErrMsg_Out
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlTransferAccountsMoney & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlTransferAccountsMoney & ",卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
End Function


Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Set mobjDatabase = New clsDataBase
    Set mobjXML = New clsXML
    glngInstanceCount = glngInstanceCount + 1
    mint三方卡结帐退款控制 = -1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mobjDatabase = Nothing
    Set mobjCurCard = Nothing
    Set mobjXML = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

Public Function zlThreeInterFaceMgrTool(ByVal frmMain As Object, ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方接口管理工具
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '编制:刘兴洪
    '日期:2012-02-13 10:17:26
    '问题:46203
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng卡类别ID As Long, bln消费卡 As Boolean
    Dim objCardObject As clsCardObject
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlThreeInterFaceMgrTool"
    
    On Error GoTo errHandle
    
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "")
    
    If frmSquareSelect.zlShowSelect(frmMain, lng卡类别ID, bln消费卡) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,选 择一卡通时返回失败!")
        Exit Function
    End If
    
    If zlGetCardObj(Me, lng卡类别ID, bln消费卡, objCardObject, True) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,获取一卡通对象集(zlGetCardObj)失败!")
        Exit Function
    End If
    
    If objCardObject.CardObject.zlThreeInterFaceMgrTool(frmMain, lngModule, lng卡类别ID) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用三方工具接口失败!")
        Exit Function
    End If
    
    zlThreeInterFaceMgrTool = True
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=true")
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,错误信息=" & strErrInfor, 2)
End Function

'------------------------------------------------------
'本地模块相关接口
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
    Dim objTemp As clsCard, objCard As Card
    Dim objTempObj As clsCardObject
    Dim objCards As Cards
    Dim strFunName As String
    
    strFunName = "zlGetCards"
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "bytType=" & bytType)
    
    Set objCards = New Cards
    '99858:李南春,2016/9/2,三方账户医疗卡、非自制消费卡必须有接口部件
    If bytType = 0 Or bytType = 1 Then
        If mObjYLCards Is Nothing Or gObjYLCards Is Nothing Or Not mObjYLCards Is gObjYLCards Then
            Call zlInitPatiCards(mcnOracle): Call InitCards
        End If
        
        For Each objTemp In mObjYLCards
            Select Case bytType
            Case 0  '0-所有医疗卡
              If Not objTemp.消费卡 Then
                Set objCard = FromClsCardToCard(objTemp)
                objCards.Add objCard
              End If
            Case 1  '1-启用的医疗卡,
                If objTemp.启用 And Not objTemp.消费卡 Then
                    Set objCard = FromClsCardToCard(objTemp)
                    objCards.Add objCard
                End If
            End Select
        Next
    Else
        If mObjYLCardObjs Is Nothing Or gObjYLCardObjs Is Nothing Or Not mObjYLCardObjs Is gObjYLCardObjs Then
            Call zlInitPatiCards(mcnOracle): Call InitCards
        End If
        
        For Each objTempObj In mObjYLCardObjs
            Select Case bytType
            Case 2  '2-所有存在三方账户的三方卡
                If objTempObj.CardPreporty.是否存在帐户 Then
                    Set objCard = FromClsCardToCard(objTempObj.CardPreporty)
                    objCards.Add objCard
                End If
            Case 3  '3-启用的三方账户的医疗卡
                If objTempObj.CardPreporty.是否存在帐户 And objTempObj.CardPreporty.启用 Then
                    Set objCard = FromClsCardToCard(objTempObj.CardPreporty)
                    objCards.Add objCard
                End If
            End Select
        Next
    End If
    Set zlGetCards = objCards
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & IIf(objCards Is Nothing, " Nothing ", " Not Nothing "))
End Function

Private Function FromClsCardToCard(ByVal objCard As clsCard) As Card
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将ClsCard对象换成Card对象
    '返回:成功Card对象
    '编制:刘兴洪
    '日期:2013-10-23 18:03:52
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As New Card
    Dim strFunName As String
    
    strFunName = "FromClsCardToCard"
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "")
    
    '85565,李南春,2015/7/10:读卡性质
    With objCard
        objTemp.接口序号 = .接口序号
        objTemp.接口编码 = .接口编码
        objTemp.短名 = .短名
        objTemp.名称 = .名称
        objTemp.前缀文本 = .前缀文本
        objTemp.卡号长度 = .卡号长度
        objTemp.缺省标志 = .缺省标志
        objTemp.消费卡 = .消费卡
        objTemp.系统 = .系统
        objTemp.是否严格控制 = .是否严格控制
        objTemp.是否自动读取 = .是否自动读取
        objTemp.自动读取间隔 = .自动读取间隔
        objTemp.自制卡 = .自制卡
        objTemp.是否存在帐户 = .是否存在帐户
        objTemp.是否全退 = .是否全退
        objTemp.卡号重复使用 = .卡号重复使用
        objTemp.结算方式 = .结算方式
        objTemp.接口程序名 = .接口程序名
        objTemp.特定项目 = .特定项目
        objTemp.启用 = .启用
        objTemp.备注 = .备注
        objTemp.卡号密文规则 = .卡号密文规则
        objTemp.是否退现 = .是否退现
        objTemp.密码长度 = .密码长度
        objTemp.密码长度限制 = .密码长度限制
        objTemp.密码规则 = .密码规则
        objTemp.密码输入限制 = .密码输入限制
        objTemp.是否缺省密码 = .是否缺省密码
        objTemp.是否制卡 = .是否制卡
        objTemp.是否发卡 = .是否发卡
        objTemp.是否写卡 = .是否写卡
        objTemp.结算性质 = .结算性质
        objTemp.是否模糊查找 = .是否模糊查找
        objTemp.是否转帐及代扣 = .是否转帐及代扣
        objTemp.是否刷卡 = .是否刷卡
        objTemp.是否扫描 = .是否扫描
        objTemp.是否接触式读卡 = .是否接触式读卡
        objTemp.是否非接触式读卡 = .是否非接触式读卡
        objTemp.是否持卡消费 = .是否持卡消费
        objTemp.是否退款验卡 = .是否退款验卡
        objTemp.是否缺省退现 = .是否缺省退现
        objTemp.是否独立结算 = .是否独立结算
        objTemp.是否支持扫码付 = .是否支持扫码付
        
    End With
    Set FromClsCardToCard = objTemp
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "FromClsCardToCard=" & IIf(objTemp Is Nothing, "Nothing ", "Not Nothing "))
End Function

Private Function InitCards() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡对象
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-03 16:05:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSameObject As Boolean
    Dim strFunName As String
    
    strFunName = "InitCards"
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "")
    
    blnSameObject = False
    If Not mcnOracle Is Nothing Then
        If mcnOracle.State = 1 Then
            blnSameObject = mcnOracle Is gcnOracle
        End If
    End If
    If Not blnSameObject _
        Or gObjYLCardObjs Is Nothing Or gObjYLCards Is Nothing Then
        If zlInitPatiCards(mcnOracle) = False Then
            Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,初始化病人卡对象")
            Exit Function
        End If
    End If
    Set mObjYLCardObjs = gObjYLCardObjs
    Set mObjYLCards = gObjYLCards
    InitCards = True
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & InitCards)
End Function

Private Function GetCards_YL(ByRef objCards As clsCards) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别
    '出参:objCards-医疗卡类别对象
    '返回:返回成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-23 12:03:26
    '说明:59760
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "GetCards_YL"
    On Error GoTo errHandle
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "")
    If Not mObjYLCards Is Nothing And Not gObjYLCards Is Nothing And mObjYLCards Is gObjYLCards Then
        Set objCards = mObjYLCards: GetCards_YL = True
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=True")
        Exit Function
    End If
    If InitCards = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False")
        Exit Function
    End If
    Set objCards = mObjYLCards
    GetCards_YL = True
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true")
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,错误信息=" & strErrInfor, 2)
End Function
Private Function GetCardObjs_YL(ByRef objYlCardObjects As clsCardObjects) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡对象
    '出参:objYlCardObjects-返回卡对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-23 13:59:24
    '说明:59760
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "GetCardObjs_YL"
    
    On Error GoTo errHandle
    
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "")
    
    If Not mObjYLCardObjs Is Nothing And Not gObjYLCardObjs Is Nothing And mObjYLCardObjs Is gObjYLCardObjs Then
        Set objYlCardObjects = mObjYLCardObjs
        GetCardObjs_YL = True
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true")
        Exit Function
    End If
    If InitCards = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False")
        Exit Function
    End If
    Set objYlCardObjects = mObjYLCardObjs
    GetCardObjs_YL = True
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=true")
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,,错误信息=" & Err.Description, 2)
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
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "ZlGetParaConfig"
    On Error GoTo errHandle
    
    strPara = "mlngModule=" & mlngModule & vbTab & "lng卡类别ID=" & lng卡类别ID & vbTab & " bln消费卡=" & IIf(bln消费卡, "True", "False")
    strPara = strPara & vbTab & "intPara=" & intPara
    strPara = strPara & vbCrLf & "strExpend=" & strExpend
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    
    Set mfrmMain = frmMain
    If zlInitInterFacel(lng卡类别ID, bln消费卡) = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lng卡类别ID)
        Exit Function
    End If
    
    If mobjCurCard Is Nothing Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,当前卡对象不存在")
        Exit Function
    End If
    If mobjCurCard.InitCompents = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,当前卡对象未初始化")
        Exit Function
    End If
    
    On Error Resume Next
    blnSucceed = mobjCurCard.CardObject.ZlGetParaConfig(frmMain, intPara, strErrMsg, strExpend)
    If Err <> 0 Then
        ZlGetParaConfig = blnSucceed
        strErrMsg = Err.Description
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & blnSucceed & ",ZlGetParaConfig,调用接口失败,错误信息=" & Err.Description, 2)
        Err = 0: On Error GoTo 0
    End If
    
    ZlGetParaConfig = blnSucceed
    strPara = ""
    If blnSucceed Then
        strPara = "strErrMsg=" & strErrMsg & vbTab & "strExpend=" & strExpend
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & blnSucceed & "," & strPara)
    Exit Function
    
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & blnSucceed & ",卡类别ID=" & lng卡类别ID & ",错误信息=" & strErrInfor, 2)
End Function

Public Function ZlGetCardNo(ByRef strCardNo As String, Optional strExpend As String) As Boolean
    '获取刷卡卡号
    '出参:strCardNo-刷卡卡号
    '     strExpend-扩展参数，暂留，现传为空
    
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "ZlGetCardNo"
    On Error GoTo errHandle
    
    strPara = "strCardNo=" & strCardNo & vbTab & "strExpend=" & strExpend
    Call zlWritLogEx(Me, mlngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
    
    If mobjCurCard Is Nothing Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,当前卡对象不存在")
        Exit Function
    End If
    If mobjCurCard.InitCompents = False Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,当前卡对象未初始化")
        Exit Function
    End If
    
    On Error Resume Next
    blnSucceed = mobjCurCard.CardObject.ZlGetCardNo(strCardNo, strExpend)
    If Err <> 0 Then
        blnSucceed = False
        Err = 0: On Error GoTo 0
    End If
    
    strPara = ""
    If blnSucceed Then
        strPara = "strCardNo=" & strCardNo & vbTab & "strExpend=" & strExpend
    End If
    
    ZlGetCardNo = blnSucceed
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=" & blnSucceed & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, mlngModule, strFunName, "结束", "返回=False,错误信息=" & strErrInfor, 2)
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
    Dim strPara As String
    Dim blnSucceed As Boolean
    Dim strFunName As String, strErrInfor As String
    
    strFunName = "zlSwapIsSucces"
    
    Err = 0: On Error GoTo errHandle:
    
    Set mfrmMain = frmMain
     
    strPara = "lngModule=" & lngModule & vbTab & "lngCardTypeID=" & lngCardTypeID & vbTab & " intSwapType=" & intSwapType
    strPara = strPara & vbTab & "strBalanceIDs=" & strBalanceIDs & vbTab & "strCardNo=" & strCardNo & vbTab & "dblSwapMoney=" & dblSwapMoney
    strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & vbCrLf & strPara)
         
         
    If zlInitInterFacel(lngCardTypeID, False) = False Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,调用初始化接口失败,卡类别ID=" & lngCardTypeID)
        Exit Function
    End If
    
    Err = 0: On Error Resume Next
    blnSucceed = mobjCurCard.CardObject.zlSwapIsSucces(frmMain, lngModule, intSwapType, lngCardTypeID, strBalanceIDs, strCardNo, dblSwapMoney, _
         intSwapStatus_Out, strErrMsg, strXMLExpend)
         
    If Err.Number = 438 Then
        '438-对象不支持该属性或方法
        Err = 0: On Error GoTo 0
        zlSwapIsSucces = False
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=False,intSwapStatus_Out=" & intSwapStatus_Out)
        Exit Function
    End If
   
    zlSwapIsSucces = blnSucceed
    strPara = ""
    If blnSucceed Then
        strPara = "intSwapStatus_Out=" & intSwapStatus_Out & vbTab & "strErrMsg=" & strErrMsg
        strPara = strPara & vbTab & "strXMLExpend=" & strXMLExpend
    End If
    
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlSwapIsSucces & ",卡类别ID=" & lngCardTypeID & "," & strPara)
    Exit Function
errHandle:
    strErrInfor = Err.Number & "-" & Err.Description
    If mobjDatabase.ErrCenter() = 1 Then
        Call zlWritLogEx(Me, mlngModule, strFunName, "执行", "重试,卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回=" & zlSwapIsSucces & ",卡类别ID=" & lngCardTypeID & ",错误信息=" & strErrInfor, 2)
End Function
