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


Public Sub InitCommon(cnMain As ADODB.Connection)
    '------------------------------------------------
    '功能:初始化连接
    '参数：
    '   cnMain:主程序的数据库连接
    '------------------------------------------------
    Set mcnOracle = cnMain
    Call mobjDatabase.InitCommon(cnMain)
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

    If Not mobjCurCard Is Nothing Then
        If mobjCurCard.接口序号 = lngCardTypeID And mobjCurCard.消费卡 = bln消费卡 Then
            If Not mobjCurCard.InitCompents Then
                If mobjCurCard.CardObject Is Nothing Then
                    blnOnlyNotObject = True: GoTo GoCreateObject:
                End If
                If Not mobjCurCard.CardObject.zlInitComponents(mfrmMain, mlngModule, glngSys, mstrDBUser, gcnOracle, False, strExpand) Then                    '初始化部件
                    Exit Function
                End If
                mobjCurCard.InitCompents = True
            End If
            zlInitInterFacel = True
            Exit Function
        End If
    End If
    Err = 0: On Error Resume Next
GoCreateObject:
    strKey = IIf(bln消费卡, "X", "K") & lngCardTypeID
    '检查设备是否启用
    '59760
    If GetCards_YL(objYLCards) = False Then Exit Function
    If GetCardObjs_YL(objYlCardObjs) = False Then Exit Function
    
    Set mobjCurCard = objYlCardObjs(strKey)
    If Err <> 0 Then
            Err = 0: On Error Resume Next
            Set objCard = objYLCards.Item(strKey)
            If Err <> 0 Then
                ShowMsgbox "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!"
                Call WritLog("zlInitInterFacel", "", "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!")
                Exit Function
            End If
            '重新创建
            If zlCreatePatiCardObject(objCard, objCardObject, blnAdviceSend) = False Then Exit Function
            '增加对应
           Set mobjCurCard = objYlCardObjs.Add(objCardObject, objCard.自制卡, objCard.接口序号, objCard, bln消费卡, strKey)
    End If
    
    If mobjCurCard Is Nothing Then
        If Not objCard Is Nothing Then
                MsgBox "注意:" & vbCrLf & "调用接口(" & objCard.接口编码 & "-" & objCard.名称 & ")调用失败,请检查!", vbInformation, gstrSysName
        Else
                MsgBox "注意:" & vbCrLf & "调用接口(" & lngCardTypeID & ")调用失败,请检查!", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    Err = 0: On Error Resume Next
    Set objCard = mobjCurCard.CardPreporty
    If Err <> 0 Then
        ShowMsgbox "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description
        Call WritLog("clsPatiCard.zlInitInterFacel", "", "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description)
        Exit Function
    End If
    If Not mobjCurCard.InitCompents Then
        If Not mobjCurCard.CardObject.zlInitComponents(mfrmMain, mlngModule, glngSys, mstrDBUser, gcnOracle, False, strExpand) Then Exit Function
         mobjCurCard.InitCompents = True
    End If
    zlInitInterFacel = True
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
    
    On Error GoTo errHandle
    
    Set mcnOracle = cnOracle: mlngModule = lngModule: mstrDBUser = strDBUser
    Call InitCards    '初始化卡对象
   
    If GetCards_YL(objYLCards) = False Then Exit Function
    If GetCardObjs_YL(objYlCardObjs) = False Then Exit Function
    
    '调用其他初始化部件
    For Each objCardObject In objYlCardObjs
        If Not objCardObject.CardObject Is Nothing Then
            If objCardObject.CardObject.zlInitComponents(frmMain, lngModule, lngSys, strDBUser, gcnOracle, blnDeviceSet, strExpand) Then
                '初始化部件
                objCardObject.InitCompents = True
            Else
                objCardObject.InitCompents = False
            End If
        End If
    Next
    Set mfrmMain = frmMain
    zlInitComponents = True
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
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
    Call zlInitPatiCards(mcnOracle)
    frmCardSelect.SelectCard frmMain, bytCardType
'    Call zlInitPatiCards(mcnOracle)'88185,此处不要去初始化卡对象,让其它实例对象在GetCardObjs_YL()和GetCards_YL()中去重新初始化
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
    '59760
    If GetCardObjs_YL(objYlCardObjs) = False Then Exit Sub
    If objYlCardObjs Is Nothing Then Exit Sub
     For Each objCardObject In objYlCardObjs
        If Not objCardObject.CardObject Is Nothing And objCardObject.InitCompents Then
            Err = 0: On Error Resume Next
            objCardObject.CardObject.SetEnabled (blnEnabled)
            If Err <> 0 Then Err = 0: On Error GoTo 0
        End If
    Next
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
    '59760
    If GetCardObjs_YL(objYlCardObjs) = False Then Exit Sub
    If objYlCardObjs Is Nothing Then Exit Sub
     For Each objCardObject In objYlCardObjs
        If Not objCardObject.CardObject Is Nothing And objCardObject.InitCompents Then
            Err = 0: On Error Resume Next
            Set objCardObject.CardObject.zlInitEvents = objComEvents
            Call objCardObject.CardObject.SetParent(hWnd)
            If Err <> 0 Then Err = 0: On Error GoTo 0
        End If
    Next
End Sub

Public Function zlReadCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal blnOlnyCardNO As Boolean, _
    ByVal strExpand As String, _
    ByRef strOutCardNo As String, _
    ByRef strOutPatiInforXml As String, _
    Optional ByRef strPhotoFile As String = "") As Boolean
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
    Dim blnInitFace As Boolean
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    '传入:strExpand-用|分离,第一栏为卡类别ID:lngCardTypeID
    strExpand = lngCardTypeID
    Err = 0: On Error Resume Next
    '56615
    zlReadCard = mobjCurCard.CardObject.zlReadCard(frmMain, lngModule, blnOlnyCardNO, strExpand, strOutCardNo, strOutPatiInforXml, strPhotoFile)
    If Err <> 0 Then
        If Err <> 450 Then GoTo errHandle:
        '450-错误的参数号或无效的属性赋值
        '主要是歉容以前的
        Err = 0: On Error GoTo errHandle:
        zlReadCard = mobjCurCard.CardObject.zlReadCard(frmMain, lngModule, blnOlnyCardNO, strExpand, strOutCardNo, strOutPatiInforXml)
    End If
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
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
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then Exit Function
    
    Err = 0: On Error Resume Next
    zlSetBrushCardObject = mobjCurCard.CardObject.zlSetBrushCardObject(lngCardTypeID, objBrushCard, strExpend)
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
    End If
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then Resume
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
    Dim blnInitFace As Boolean
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    '56615
    On Error Resume Next
    zlMakeCard = mobjCurCard.CardObject.zlMakeCard(frmMain, lngModule, lngCardTypeID, stInputPatiInforXml, _
        strPhotoFile, strOutPatiInforXml, strExpend)
    If Err = 438 Then
        MsgBox mobjCurCard.CardPreporty.名称 & "没有编写制卡接口(zlMakeCard),制卡失败!", vbInformation, gstrSysName
        Err.Clear
    ElseIf Err <> 0 Then
        GoTo errHandle
    End If
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then Resume
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
    Dim objCards As clsCards, objCard As clsCard
    Dim blnOK As Boolean
    On Error GoTo errHandle
    If zlGetCards_YL(objCards) = False Then Exit Function
    blnOK = False
    If lngCardTypeID = 0 Then
        For Each objCard In objCards
            If objCard.接口序号 <> 0 And objCard.接口程序名 <> "" And objCard.消费卡 = False And objCard.发送调用接口 = True Then
                If zlInitInterFacel(objCard.接口序号, objCard.消费卡, True) Then
                    Err = 0: On Error Resume Next
                    If mobjCurCard.CardObject.zlAdviceSendBulidBarCode(frmMain, lngModule, objCard.接口序号, intRecNature, strNos, strExpand) = True Then blnOK = True
                    Err = 0: On Error GoTo errHandle:
                End If
            End If
        Next
    Else
        For Each objCard In objCards
            If objCard.接口序号 = lngCardTypeID And objCard.接口程序名 <> "" And objCard.消费卡 = False And objCard.发送调用接口 = True Then
                If zlInitInterFacel(objCard.接口序号, objCard.消费卡, True) Then
                    Err = 0: On Error Resume Next
                    If mobjCurCard.CardObject.zlAdviceSendBulidBarCode(frmMain, lngModule, objCard.接口序号, intRecNature, strNos, strExpand) = True Then blnOK = True
                    Err = 0: On Error GoTo errHandle:
                End If
            End If
        Next
    End If
    zlAdviceSendBulidBarCode = blnOK
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
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
    Dim blnInitFace As Boolean
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    Err = 0: On Error Resume Next
    '56615
    zlBandCardArfter = mobjCurCard.CardObject.zlBandCardArfter(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, strExpend)
    If Err <> 0 Then
        If Err.Number <> 438 Then GoTo errHandle:
        '438-对象不支持该属性或方法
        '歉容以前的,直接返回true
        Err = 0
        zlBandCardArfter = True
    End If
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then Resume
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
    Optional ByVal lng病人ID As Long) As Boolean
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
    '       strXmlIn-三方卡调用XML入参,目前格式如下:
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
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:18:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmInput As frmInputPass
    Dim blnErrCount As Boolean
         
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain
    If bln转预交 = False And lngModule = 1103 Then bln转预交 = True
    If lngCardTypeID = 0 Then
        '暂不支持
        Set frmInput = New frmInputPass
        Set varSquareBalance = Nothing
        zlBrushCard = frmInput.zlBrushPay(frmMain, lngModule, Nothing, rsClassMoney, lngCardTypeID, _
            bln消费卡, strPatiName, strSex, strOld, dbl金额, strCardNo, strPassWord, bln退费, _
            blnShowPatiInfor, bln退现, bln余额不足禁止, varSquareBalance, bln转预交, blnAllPay, str费用来源, lng病人ID)
        Set varSquareBalance = Nothing
        Unload frmInput
        Set frmInput = Nothing
        Exit Function
    End If
    
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then Exit Function
    
    If bln消费卡 And mobjCurCard.自制卡 Then
        Set frmInput = New frmInputPass
        Err = 0: On Error Resume Next
        If IsEmpty(varSquareBalance) Then   '57682
            Set varSquareBalance = Nothing
        End If
        blnErrCount = varSquareBalance.count
        If Err <> 0 Then
            Set varSquareBalance = Nothing
            Err = 0: On Error GoTo 0
        End If
        zlBrushCard = frmInput.zlBrushPay(frmMain, lngModule, mobjCurCard, rsClassMoney, _
            lngCardTypeID, bln消费卡, strPatiName, strSex, strOld, dbl金额, _
            strCardNo, strPassWord, bln退费, blnShowPatiInfor, bln退现, bln余额不足禁止, varSquareBalance, _
            bln转预交, blnAllPay, str费用来源, lng病人ID)
        Unload frmInput
         
        Set frmInput = Nothing
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
    zlBrushCard = mobjCurCard.CardObject.zlBrushCard(frmMain, lngModule, lngCardTypeID, _
     strPatiName, strSex, strOld, dbl金额, strCardNo, strPassWord, strXmlIn)
    If Err = 450 Then
        Err = 0
        On Error GoTo errHandle
        zlBrushCard = mobjCurCard.CardObject.zlBrushCard(frmMain, lngModule, lngCardTypeID, _
        strPatiName, strSex, strOld, dbl金额, strCardNo, strPassWord)
    End If
    
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then Resume
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
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(strCardTypeID, bln消费卡) = False Then Exit Function
    zlPaymentCheck = mobjCurCard.CardObject.zlPaymentCheck(frmMain, lngModule, strCardTypeID, strCardNo, dblMoney, strNos, strXMLExpend)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then Resume
End Function

Public Function zlReturnCashCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退现交易检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(退款时检查)
    '       strSwapMemo-交易说明(退款时传入)
    '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
    '返回:退现合法,返回true,否则返回Flase
    '编制:刘尔旋
    '日期:2017-7-3
    '说明:
    '    1.使用模块：病人结帐管理
    '    2.触发时机:在结帐列表中点“X"时，调用，如果返回true,则允许退现，否则禁止
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    On Error Resume Next
    zlReturnCashCheck = mobjCurCard.CardObject.zlReturnCashCheck(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strXMLExpend)
    If Err <> 0 Then
        zlReturnCashCheck = True
        Err = 0: On Error GoTo 0
    End If
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, ByVal strPrepayNos As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:13:48
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
   On Error GoTo errHandle
   Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then Exit Function
    zlPaymentMoney = mobjCurCard.CardObject.zlPaymentMoney(frmMain, lngModule, _
    lngCardTypeID, strCardNo, strBalanceIDs, strPrepayNos, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then Resume
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
    '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
    '返回:退款合法,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:24:55
    '说明:
    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
    '    以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then Exit Function
    zlReturncheck = mobjCurCard.CardObject.zlReturncheck(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapNo, strSwapMemo, strXMLExpend)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
    ByRef strSwapExtendInfor As String) As Boolean
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
    '       strSwapExtendInfor-本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '出参: strSwapNo-交易流水号(退款交易流水号)
    '      strSwapMemo-交易说明(退款交易说明)
    '      strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:31:32
    '说明:
    '       在所有需要对已经消费的地方进行退款的情况下，调用该接口,目前规划在:收费室；挂号等。
    '---------------------------------------------------------------------------------------------------------------------------------------------
     On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then Exit Function
    zlReturnMoney = mobjCurCard.CardObject.zlReturnMoney(frmMain, lngModule, lngCardTypeID, strCardNo, strBalanceIDs, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlReturnMultiMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByVal strInXML As String, _
    ByVal lng冲销ID As Long, ByRef strOutXml As String, ByRef strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易(多笔回退)
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strInXML-XML串:
    '       <JSLIST>
    '           <JS>
    '               <KH>卡号</KH>
    '               <JYLSH>交易流水号</JYLX>
    '               <JYSM>交易说明</JYSM>
    '               <ZFJE>作废金额</ZFJE>
    '               <JSLX>类型</JSLX>  //1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算
    '               <ID></ID>    //类型=1时,预交ID;类型=2,6时，为原结帐ID
    '           </JS>
    '       </JSLIST>
    '       lng冲销ID-作废时的冲销ID(作废时或退费时有效，否则为0）;类型=6时，传入结算序号
    '       strExpend-无（暂留，待以后扩展)
    '出参:
    '     strOutXML-返回XML串
    '       <JSLIST>
    '           <JS>
    '               <KH>卡号</KH>
    '               <TKLSH>退款交易流水号</TKLSH>
    '               <TKSM>退款交易说明</TKSM>
    '               <ID></ID>
    '           </JS>
    '       </JSLIST>
    '      strExpend-交易的扩展信息
    '       <Expends>
    '           <Expend>
    '               <XMMC>项目名称1</XMMC>
    '               <XMNR>项目内容2</XMNR>
    '           </Expend>
    '       </Expends>
    '返回:函数返回    True:调用成功,False:调用失败
    '日期:2015-11-10
    '说明:
    '   目前只有结帐程序时有效（结帐退款),用于一次性处理同一卡类别的多笔三方交易作废
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then Exit Function
    zlReturnMultiMoney = mobjCurCard.CardObject.zlReturnMultiMoney(frmMain, lngModule, lngCardTypeID, bln消费卡, strInXML, lng冲销ID, strOutXml, strExpend)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
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
    Dim blnInitFace As Boolean
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If lngCardTypeID <> 0 Then
        lngCardTypeID = Val(ZLGetPatiCardFromCards(CStr(lngCardTypeID), lng病人ID))
    End If
    If lngCardTypeID = 0 Then
        '需要读取相关的卡类别ID
        If zlSelectWriteCardType(frmMain, lngCardTypeID, mcnOracle, lng病人ID) = False Then Exit Function
        If lngCardTypeID = 0 Then Exit Function
    End If
    
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    '56615
    Err = 0: On Error Resume Next
    zlMzInforWriteToCard = mobjCurCard.CardObject.zlMzInforWriteToCard(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, lngBalanceID, strExpend)
    If Err = 438 Then
        'VB实时错误438：对象不支持该属性或方法
        MsgBox mobjCurCard.CardPreporty.名称 & "没有编写写卡接口(zlMzInforWriteToCard)，门诊信息写卡失败！", vbInformation, gstrSysName
        Err.Clear
    ElseIf Err <> 0 Then
        GoTo errHandle
    End If
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then Resume
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
    Dim blnInitFace As Boolean
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If lngCardTypeID <> 0 Then
        lngCardTypeID = Val(ZLGetPatiCardFromCards(CStr(lngCardTypeID), lng病人ID))
    End If
    If lngCardTypeID = 0 Then
        '需要读取相关的卡类别ID
        If zlSelectWriteCardType(frmMain, lngCardTypeID, mcnOracle, lng病人ID) = False Then Exit Function
        If lngCardTypeID = 0 Then Exit Function
    End If
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    '56615
    Err = 0: On Error Resume Next
    zlZyInforWriteToCard = mobjCurCard.CardObject.zlZyInforWriteToCard(frmMain, lngModule, lngCardTypeID, _
    lng病人ID, lngBalanceID, strExpend)
    If Err = 438 Then
        'VB实时错误438：对象不支持该属性或方法
        MsgBox mobjCurCard.CardPreporty.名称 & "没有编写写卡接口(zlZyInforWriteToCard)，住院信息写卡失败！", vbInformation, gstrSysName
        Err.Clear
    ElseIf Err <> 0 Then
        GoTo errHandle
    End If
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
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
    Dim objCard As clsCard, lngCardTypeID As Long
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    '59760
    If GetCards_YL(objYLCards) = False Then Exit Function
    If GetCardObjs_YL(objYlCardObjs) = False Then Exit Function
    
    Dim i As Long
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    strFunName = "健康卡"
    For i = 1 To objYlCardObjs.count
        Set objCard = objYlCardObjs(i).CardPreporty
        If objCard.启用 And objCard.名称 = "居民健康卡" And objCard.消费卡 = False Then
            '只取一个
            lngCardTypeID = objCard.接口序号: Exit For
        End If
    Next
    If lngCardTypeID = 0 Then Exit Function
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    '56615
    zlHealthArchiveIsSHow = mobjCurCard.CardObject.zlHealthArchiveIsSHow(frmMain, lngModule, strFunName, strExpend)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
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
    '59760
    If GetCards_YL(objYLCards) = False Then Exit Function
    If GetCardObjs_YL(objYlCardObjs) = False Then Exit Function
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    For i = 1 To objYlCardObjs.count
        Set objCard = objYlCardObjs(i).CardPreporty
        If objCard.启用 And objCard.名称 = "居民健康卡" And objCard.消费卡 = False Then
            '只取一个
            lngCardTypeID = objCard.接口序号: Exit For
        End If
    Next
    If lngCardTypeID = 0 Then Exit Function
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    '56615
    blnShow = mobjCurCard.CardObject.zlHealthArchivesShow(frmMain, lngModule, lng病人ID, strExpend)
    If lng病人ID = 0 Then Exit Function
    If Not blnShow Then
            Set frmShowSelect = New frmHealthArchives
            Call frmShowSelect.zlShowHealthArchives(frmMain, lng病人ID, mcnOracle)
            If Not frmShowSelect Is Nothing Then Unload frmShowSelect
            Set frmShowSelect = Nothing
    End If
    zlHealthArchivesShow = True
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
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
    If zlInitInterFacel(lngCardTypeID, bln消费卡) = False Then Exit Function
    zlGetAccountMoney = mobjCurCard.CardObject.zlGetAccountMoney(frmMain, lngModule, lngCardTypeID, strCardNo, strExpand, dblMoney)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlTransferAccountsCheck(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, ByVal dblMoney As Double, ByVal strBalanceID As String, Optional ByRef strXMLExpend As String) As Boolean
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
    '返回:检查的数据合法,返回True:否则返回False
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    zlTransferAccountsCheck = mobjCurCard.CardObject.zlTransferAccountsCheck(frmMain, lngModule, lngCardTypeID, strCardNo, dblMoney, strBalanceID, strXMLExpend)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlTransferAccountsMoney(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, ByVal strBalanceID As String, ByVal dblMoney As Double, Optional ByRef strSwapGlideNO As String, _
    Optional ByRef strSwapMemo As String, Optional ByRef strSwapExtendInfor As String, Optional ByRef strXMLExpend As String) As Boolean
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
    '           <ERRMSG>错误信息</ERRMSG >
    '        </OUT>
    '编制:刘兴洪
    '日期:2014-09-03 14:22:10
    '调用者:医保补充结算(结算时调用)
    '说明:
    '  １. 在医保补充结算时进行的三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
    '  ２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    If zlInitInterFacel(lngCardTypeID, False) = False Then Exit Function
    zlTransferAccountsMoney = mobjCurCard.CardObject.zlTransferAccountsMoney(frmMain, lngModule, _
    lngCardTypeID, strCardNo, strBalanceID, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend)
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Set mobjDatabase = New clsDataBase
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mobjDatabase = Nothing
    Set mobjCurCard = Nothing
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
    If frmSquareSelect.zlShowSelect(frmMain, lng卡类别ID, bln消费卡) = False Then Exit Function
    If zlGetCardObj(Me, lng卡类别ID, bln消费卡, objCardObject, True) = False Then Exit Function
    Err = 0: On Error GoTo Errhand:
    If objCardObject.CardObject.zlThreeInterFaceMgrTool(frmMain, lngModule, lng卡类别ID) = False Then Exit Function
    zlThreeInterFaceMgrTool = True
    Exit Function
Errhand:
    If mobjDatabase.ErrCenter = 1 Then Resume
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
    Set objCards = New Cards
    '99858:李南春,2016/9/2,三方账户医疗卡、非自制消费卡必须有接口部件
    If bytType = 0 Or bytType = 1 Then
        If mObjYLCards Is Nothing Then
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
        If mObjYLCardObjs Is Nothing Then
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
    End With
    Set FromClsCardToCard = objTemp
End Function

Private Function InitCards() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡对象
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-03 16:05:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSameObject As Boolean
    blnSameObject = False
    If Not mcnOracle Is Nothing Then
        If mcnOracle.State = 1 Then
            blnSameObject = mcnOracle Is gcnOracle
        End If
    End If
    If Not blnSameObject _
        Or gObjYLCardObjs Is Nothing Or gObjYLCards Is Nothing Then
        If zlInitPatiCards(mcnOracle) = False Then Exit Function
    End If
    Set mObjYLCardObjs = gObjYLCardObjs
    Set mObjYLCards = gObjYLCards
    InitCards = True
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
    On Error GoTo errHandle
    If Not mObjYLCards Is Nothing And Not gObjYLCards Is Nothing Then
        Set objCards = mObjYLCards: GetCards_YL = True: Exit Function
    End If
    If InitCards = False Then Exit Function
    Set objCards = mObjYLCards
    GetCards_YL = True
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
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
    On Error GoTo errHandle
    If Not mObjYLCardObjs Is Nothing And Not gObjYLCardObjs Is Nothing Then
        Set objYlCardObjects = mObjYLCardObjs
        GetCardObjs_YL = True
        Exit Function
    End If
    If InitCards = False Then Exit Function
    Set objYlCardObjects = mObjYLCardObjs
    GetCardObjs_YL = True
    Exit Function
errHandle:
    If mobjDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlGetParaConfig(ByVal frmMain As Object, _
    ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, ByVal intPara As Integer, _
    Optional strErrMsg As String, Optional strExpend As String) As Boolean
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
    Set mfrmMain = frmMain
    If zlInitInterFacel(lng卡类别ID, bln消费卡) = False Then Exit Function
    If mobjCurCard Is Nothing Then Exit Function
    If mobjCurCard.InitCompents = False Then Exit Function
    On Error Resume Next
    ZlGetParaConfig = mobjCurCard.CardObject.ZlGetParaConfig(frmMain, intPara, strErrMsg, strExpend)
    If Err <> 0 Then
        ZlGetParaConfig = False
        strErrMsg = Err.Description
        Err = 0: On Error GoTo 0
    End If
End Function


Public Function ZlGetCardNo(ByRef strCardNo As String, Optional strExpend As String) As Boolean
    '获取刷卡卡号
    '出参:strCardNo-刷卡卡号
    '     strExpend-扩展参数，暂留，现传为空
    If mobjCurCard Is Nothing Then Exit Function
    If mobjCurCard.InitCompents = False Then Exit Function
    On Error Resume Next
    ZlGetCardNo = mobjCurCard.CardObject.ZlGetCardNo(strCardNo, strExpend)
    If Err <> 0 Then
        ZlGetCardNo = False
        Err = 0: On Error GoTo 0
    End If
End Function
