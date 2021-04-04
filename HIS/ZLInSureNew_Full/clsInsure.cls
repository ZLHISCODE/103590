VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Public Version As String
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
'    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;3-公共部件增加GetNextNO();4-V10.24及以上版本
'    99-所有交易增加附加参数(最新版)
'ADO事务错误状态判断
Private WithEvents mobjOracle As ADODB.Connection
Attribute mobjOracle.VB_VarHelpID = -1
Private mlngADOErrNum As Long
Private mstrADOErrDesc As String
Private mblnTranState As Boolean

Private Sub mobjOracle_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    mlngADOErrNum = 0: mstrADOErrDesc = ""
    mblnTranState = True
End Sub

Private Sub mobjOracle_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    mlngADOErrNum = 0: mstrADOErrDesc = ""
    mblnTranState = False
End Sub

Private Sub mobjOracle_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If pConnection.Errors.Count > 0 Then
        mlngADOErrNum = pConnection.Errors(0).NativeError
        mstrADOErrDesc = pConnection.Errors(0).Description
    Else
        mlngADOErrNum = 0: mstrADOErrDesc = ""
    End If
    mblnTranState = False
End Sub

'记录当前是否在事务中
Friend Property Get zlTranState() As Boolean
    zlTranState = mblnTranState
End Property

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'------------------------------------------------
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngSys : 系统编号
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'返回：
'------------------------------------------------
    Dim intinsure As Integer
    Dim intSelect As Integer
    
    Set gcnOracle = cnMain
    Set mobjOracle = cnMain
    
    gstrDbUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    
    intinsure = GetSetting("ZLSOFT", "公共全局", "医保类别", 0)
    #If gverControl >= 4 Then
        gstrMatchMethod = zlDatabase.GetPara(4)
    #Else
        gstrMatchMethod = GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0)
    #End If
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstr单位名称 = GetUnitName()
    Call GetUserInfo '获取登陆用户信息
'    If Not CheckValid Then Exit Sub

    '-------------------------------------------------
    '1600  保险类别设置
    '1601  年度结算规则
    '1602  医保大类管理
    '1603  医保项目管理
    '1604  医保帐户管理
    '1605  医保结算管理
    '1606  帐户支出管理
    '1607  医保工具管理
    
    On Error Resume Next
    '让操作员选择医保接口
    intinsure = ChooseInsure_Base(intinsure)
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Sub
        Call gobjInsure_Obj(mintOrder).CodeMan(lngSys, lngModul, cnMain, frmMain, strDBUser, intinsure)
    Else
        Select Case lngModul
            Case 1600
                frm医保类别.Show , frmMain
            Case 1601
                If Not frm结算规则.CheckForm Then Exit Sub
                frm结算规则.ShowForm frmMain
    '            frm保险报销政策.ShowForm frmMain
            Case 1602
                If intinsure = TYPE_大连市 Or intinsure = TYPE_大连开发区 Then
                    If Not frm保险大类_大连.CheckForm Then Exit Sub
                    frm保险大类_大连.ShowForm frmMain
                Else
                    If Not frm保险大类.CheckForm Then Exit Sub
                    frm保险大类.ShowForm frmMain
                End If
            Case 1603
                If intinsure = TYPE_广元旺苍 Or intinsure = TYPE_成都德阳 Or intinsure = TYPE_南充阆中 Then
                    If Not frm保险项目_广元旺苍.CheckForm(intinsure) Then Exit Sub
                    frm保险项目_广元旺苍.ShowForm frmMain, intinsure
                Else
                    If Not frm保险项目.CheckForm Then Exit Sub
                    frm保险项目.ShowForm frmMain
                End If
                
            Case 1604
                If Not frm医保帐户.CheckForm Then Exit Sub
                frm医保帐户.ShowForm frmMain
            Case 1605
                If Not frm医保结算.CheckForm Then Exit Sub
                frm医保结算.ShowForm frmMain
            Case 1606
                If Not frm保险病种.CheckForm Then Exit Sub
                frm保险病种.ShowForm frmMain
            Case 1607
                Call frm医保工具.InitInsure(intinsure)
                Call frm医保工具.ShowForm(frmMain, intinsure)
'                frm医保帐户收支管理.Show , frmMain
        End Select
    End If
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前数据库登录用户名
    Call CodeMan(lngSys, lngModul, cnMain, Nothing, strDBUser)
'    Dim intinsure As Integer
'    Dim intSelect As Integer
'
'    Set gcnOracle = cnMain
'    Set mobjOracle = cnMain
'
'    gstrDbUser = strDBUser
'    glngSys = lngSys
'
'    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
'    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
'    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
'
'    intinsure = GetSetting("ZLSOFT", "公共全局", "医保类别", 0)
'    #If gverControl >= 4 Then
'        gstrMatchMethod = zlDatabase.GetPara(4)
'    #Else
'        gstrMatchMethod = GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0)
'    #End If
'
'    gstrPrivs = GetPrivFunc(lngSys, lngModul)
'    gstr单位名称 = GetUnitName()
'    Call GetUserInfo '获取登陆用户信息
''    If Not CheckValid Then Exit Sub
'
'    '-------------------------------------------------
'    '1600  保险类别设置
'    '1601  年度结算规则
'    '1602  医保大类管理
'    '1603  医保项目管理
'    '1604  医保帐户管理
'    '1605  医保结算管理
'    '1606  帐户支出管理
'    '1607  医保工具管理
'
'    On Error Resume Next
'    '让操作员选择医保接口
'    intinsure = ChooseInsure_Base(intinsure)
'
'    If IsApartComponents(intinsure) Then
'        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Sub
'        Call gobjInsure_Obj(mintOrder).BHCodeMan(lngSys, lngModul, cnMain, lngMain, strDBUser, intinsure)
'    Else
'        Select Case lngModul
'            Case 1600
'                OS.ShowChildWindow frm医保类别.hwnd, lngMain
'            Case 1601
'                If Not frm结算规则.CheckForm Then Exit Sub
'                OS.ShowChildWindow frm结算规则.hwnd, lngMain
'            Case 1602
'                If intinsure = TYPE_大连市 Or intinsure = TYPE_大连开发区 Then
'                    If Not frm保险大类_大连.CheckForm Then Exit Sub
'                    OS.ShowChildWindow frm保险大类_大连.hwnd, lngMain
'                Else
'                    If Not frm保险大类.CheckForm Then Exit Sub
'                    OS.ShowChildWindow frm保险大类.hwnd, lngMain
'                End If
'            Case 1603
'                If intinsure = TYPE_广元旺苍 Or intinsure = TYPE_成都德阳 Or intinsure = TYPE_南充阆中 Then
'                    If Not frm保险项目_广元旺苍.CheckForm(intinsure) Then Exit Sub
'                    OS.ShowChildWindow frm保险项目_广元旺苍.hwnd, lngMain
'                Else
'                    If Not frm保险项目.CheckForm Then Exit Sub
'                    OS.ShowChildWindow frm保险项目.hwnd, lngMain
'                End If
'
'            Case 1604
'                If Not frm医保帐户.CheckForm Then Exit Sub
'                OS.ShowChildWindow frm医保帐户.hwnd, lngMain
'            Case 1605
'                If Not frm医保结算.CheckForm Then Exit Sub
'                OS.ShowChildWindow frm医保结算.hwnd, lngMain
'            Case 1606
'                If Not frm保险病种.CheckForm Then Exit Sub
'                OS.ShowChildWindow frm保险病种.hwnd, lngMain
'            Case 1607
'                Call frm医保工具.InitInsure(intinsure)
'                OS.ShowChildWindow frm医保工具.hwnd, lngMain
'        End Select
'    End If
End Sub

Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
    gintDebug = -1
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Set gclsInsure = Me
End Sub

Public Function CloseWindows() As Boolean
    Dim intObject As Integer, intCOUNT As Integer
    On Error Resume Next
    
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
    
    '--------------------------------------
    '功能:关闭所有分离医保接口部件打开的窗口
    '--------------------------------------
    intCOUNT = UBound(gobjInsure_Obj)
    For intObject = 0 To intCOUNT
        If Not gobjInsure_Obj(intObject) Is Nothing Then
            Call gobjInsure_Obj(intObject).CloseWindows
        End If
    Next
    
    Call ReleaseME
End Function

Public Sub InitOracle(cnOracle As ADODB.Connection)
'功能：初始化医保部件的本地Oracle连接
'说明：用于某些不需要连接医保前置机,但又要调用相关接口的场合,如住院记帐
'      如果调用于InitInsure,则不必再调用本函数
    Set gcnOracle = cnOracle
    Set mobjOracle = cnOracle
    
End Sub

Public Function InitInsure(cnOracle As ADODB.Connection, Optional intinsure As Integer = 0) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'参数: cnOracle -应用部件中已经建立的到ORacle的连接
'      intType=保险类别,在系统参数中设置时需要传递保险类别
'返回：初始化成功，返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1103-基本,1121-基本,1131-基本,1132-基本,1133-基本,1134-基本,1135-基本,1137-基本
    Dim intTmp As Integer
    Dim rsTmp As New ADODB.Recordset
    
    gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
    
    Set gcnOracle = cnOracle
    intTmp = GetSetting("ZLSOFT", "公共全局", "医保类别", 0)
    If intinsure = 0 Then intinsure = intTmp
    If intinsure = 0 Then
        MsgBox "不能确定保险类别,无法执行医保交易！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '系统名
    Call IsToolInPara
    gstrSysName = GetSetting("ZLSOFT", "注册信息", UCase("gstrSysName"), "")
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    Call GetUserInfo
    
    '获取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "医保接口", intinsure)
    If rsTmp.EOF = True Then
        MsgBox "本地选择的医保类型已经不存在，请重新设置！", vbExclamation, gstrSysName
        Exit Function
    End If
    If IsNull(rsTmp("医院编码")) = True Then
        MsgBox "由于未设置医院编号，无法执行医保交易！", vbExclamation, gstrSysName
        Exit Function
    End If
    gstr医院编码 = rsTmp!医院编码
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder) Then Exit Function
        InitInsure = gobjInsure_Obj(mintOrder).InitInsure(gcnOracle, intinsure)
    Else
        Select Case intinsure
            Case TYPE_四川自贡
                InitInsure = 医保初始化_自贡
            Case TYPE_昭通
                InitInsure = 医保初始化_昭通()
            Case TYPE_徐州农保
                InitInsure = 医保初始化_徐州农保()
            Case TYPE_徐州
                InitInsure = 医保初始化_徐州()
            Case TYPE_徐州六院
                InitInsure = True
            Case TYPE_成都市农医
                InitInsure = True
            Case TYPE_宁海
                InitInsure = 医保初始化_宁海
            Case TYPE_余姚
                InitInsure = 医保初始化_余姚()
            Case TYPE_浙江
                InitInsure = 医保初始化_浙江()
            Case TYPE_吉林
                '刘兴宏:20040923加入
                InitInsure = 医保初始化_吉林()
            Case TYPE_北京尚洋
                InitInsure = 医保初始化_北京尚洋
            Case TYPE_华东
                InitInsure = 医保初始化_华东
            Case TYPE_南京市
                InitInsure = 医保初始化_南京市
            Case TYPE_江苏
                InitInsure = 医保初始化_江苏
            Case TYPE_重庆市
                InitInsure = 医保初始化_重庆
            Case TYPE_涪陵
                InitInsure = 医保初始化_涪陵
            Case TYPE_广元
                InitInsure = 医保初始化_广元
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                InitInsure = 医保初始化_松藻(intinsure)
            Case TYPE_重庆壁山
                InitInsure = 医保初始化_壁山
            Case TYPE_重庆中梁山
                InitInsure = 医保初始化_中梁山
            Case TYPE_昆明市, TYPE_云南省
                InitInsure = 医保初始化_云南(intinsure)
            Case TYPE_云南建水
                InitInsure = 医保初始化_云南建水
            Case TYPE_自贡市
                InitInsure = 医保初始化_中软
            Case TYPE_泸州市
                InitInsure = 医保初始化_泸州
            Case TYPE_铜仁
                InitInsure = 医保初始化_铜仁
            Case TYPE_贵阳市
                InitInsure = 医保初始化_贵阳
            Case TYPE_成都市
                InitInsure = 医保初始化_成都
            Case TYPE_成都莲合
                InitInsure = 医保初始化_莲合
            Case TYPE_开县
                '20050124
                InitInsure = 医保初始化_开县
            Case TYPE_新都
                InitInsure = 医保初始化_新都
            Case type_成都郊县
                InitInsure = 医保初始化_成都郊县
            Case TYPE_成都南充
                InitInsure = 医保初始化_成都南充
            Case TYPE_咸阳市
                InitInsure = 医保初始化_咸阳
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                InitInsure = 医保初始化_福建巨龙
            'Modified by zyb #2003-10-08
            Case type_米易
                InitInsure = 医保初始化_米易
            Case TYPE_大连市, TYPE_大连开发区
                '刘兴宏200311
                InitInsure = 医保初始化_大连(intinsure)
            Case TYPE_四川眉山
                InitInsure = 医保初始化_眉山
            Case TYPE_沈阳市
                InitInsure = 医保初始化_沈阳市
            Case TYPE_乐山
                InitInsure = 医保初始化_乐山
            Case TYPE_重大校园卡
                '刘兴宏:200403
                InitInsure = 医保初始化_重大校园卡
            Case TYPE_重庆银海版
                InitInsure = 医保初始化_重庆银海版
            Case TYPE_北京
                InitInsure = 医保初始化_北京
            Case TYPE_重庆渝北
                '刘兴宏:20040705
                InitInsure = 医保初始化_重庆渝北
            Case TYPE_兴安
                '刘兴宏:20040827
                InitInsure = 医保初始化_兴安
            Case TYPE_毕节
                InitInsure = 医保初始化_毕节
            Case TYPE_黔南
                '刘兴宏:20041022
                InitInsure = 医保初始化_黔南
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                InitInsure = 医保初始化_奉庆
    
            Case TYPE_成都德阳
                '刘兴宏:20041103
                InitInsure = 医保初始化_成都德阳
            Case TYPE_成都内江
                '刘兴宏:20041207
                InitInsure = 医保初始化_成都内江
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                InitInsure = 医保初始化_广元旺苍
            Case TYPE_南充阆中
                '刘兴宏:20050419
                InitInsure = 医保初始化_南充阆中
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                InitInsure = 医保初始化_兴成
                
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                InitInsure = 医保初始化_神木大兴
            
            
            Case TYPE_渝北农医
                InitInsure = 医保初始化_渝北农医
            Case TYPE_慈溪农医
                InitInsure = 医保初始化_慈溪农医
           '陈东 20050311
            Case TYPE_山西
                InitInsure = 医保初始化_山西
            Case TYPE_冕宁
            '陈东 20050401
                InitInsure = 医保初始化_冕宁
            Case type_铜梁合医
            '陈东 20050530
                InitInsure = 医保初始化_铜梁合医
            Case TYPE_铜山县
                InitInsure = 医保初始化_铜山县
            Case Is > 900
                '中联医保
                InitInsure = 医保初始化_中联(intinsure)
        End Select
    End If
    
    '如果初始化成功，则记录下该医保序号
    If InitInsure Then
        If InStr(1, gstrInsure & ",", "," & intinsure & ",") = 0 Then
            gstrInsure = gstrInsure & "," & intinsure
        End If
    End If
End Function

Public Function EndInsure(ByVal intinsure As Integer) As Boolean
'功能：断开与医保系统的连接
    
    Select Case intinsure
        Case TYPE_昭通
            EndInsure = 医保终止_昭通()
        Case TYPE_徐州农保
            Set gcn徐州农保 = Nothing
        Case TYPE_徐州
            Set gcn徐州 = Nothing
        Case TYPE_宁海
            Call 医保终止_宁海
        Case TYPE_余姚
            EndInsure = 医保终止_余姚()
        Case TYPE_云南省, TYPE_昆明市
            Call 医保终止_云南
        Case TYPE_云南建水
            Call yh_quit
        Case TYPE_重庆壁山
'            gcn壁山.Close
            Set gcn壁山 = Nothing
        Case TYPE_贵阳市
            '不需要完成特别的操作
        Case TYPE_自贡市, TYPE_泸州市, TYPE_铜仁
            '不需要完成特别的操作
        Case TYPE_涪陵, TYPE_南京市, TYPE_江苏
            '不需要完成特别的操作
        Case TYPE_华东
            Set gcn华东 = Nothing
        Case TYPE_北京尚洋
            Set gcn尚洋 = Nothing
        Case TYPE_广元
            '不需要完成特别的操作
        Case TYPE_成都市
            Set gcnSybase = Nothing
        Case TYPE_成都莲合
            Set gcnSybase = Nothing
        Case TYPE_开县
            '20050124
            EndInsure = 医保终止_开县
        Case TYPE_成都南充
            Set gcnInterbase = Nothing
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            EndInsure = 医保终止_福建巨龙
        Case type_米易
            EndInsure = 医保终止_米易
        Case TYPE_四川眉山
            EndInsure = 医保终止_眉山
        Case TYPE_沈阳市
            EndInsure = 医保终止_沈阳市
        Case TYPE_大连市, TYPE_大连开发区
            EndInsure = 医保终止_大连
        Case TYPE_乐山
            EndInsure = 医保终止_乐山
        Case TYPE_重大校园卡
            '刘兴宏:200403
            EndInsure = 医保终止_重大校园卡
        Case TYPE_重庆银海版
            EndInsure = 医保终止_重庆银海版
        Case TYPE_重庆渝北
            '刘兴宏:20040705
            EndInsure = 医保终止_重庆渝北
        Case TYPE_兴安
            '刘兴宏:20040827
            EndInsure = 医保终止_兴安
        Case TYPE_毕节
            EndInsure = 医保终止_毕节
        Case TYPE_黔南
            '刘兴宏:20041022
            EndInsure = 医保终止_黔南
        Case TYPE_临沧奉庆
            '刘兴宏:20041231
            EndInsure = 医保终止_奉庆
            
        Case TYPE_成都德阳
            '刘兴宏:20041103
            EndInsure = 医保终止_成都德阳
        Case TYPE_成都内江
            '刘兴宏:20041207
            EndInsure = 医保终止_成都内江
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            EndInsure = 医保终止_广元旺苍
        Case TYPE_南充阆中
            '刘兴宏:20050419
            EndInsure = 医保终止_南充阆中
            
        Case TYPE_兴成核工业
            '刘兴宏:20050226
            EndInsure = 医保终止_兴成
        Case TYPE_陕西大兴
            '刘兴宏:20050316
            EndInsure = 医保终止_神木大兴
            
        Case TYPE_渝北农医
            EndInsure = 医保终止_渝北农医
        Case TYPE_慈溪农医
            EndInsure = 医保终止_慈溪农医
        Case type_铜梁合医
            EndInsure = 医保终止_铜梁合医
            '20050530  陈东
        Case TYPE_铜山县
            EndInsure = 医保终止_铜山县
    End Select
    EndInsure = True
End Function

Public Function Identify2(strCard As String, strPass As String, Optional bytType As Byte, Optional ByRef lng病人ID As Long = 0, _
    Optional ByRef intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strCard=要验证的卡号,在调用程序中刷卡产生
'                         新导诊系统传入的是身份识别的卡号(比如:就诊卡卡号;身份证号等,可以根据strAdvance中的卡类别ID来进行认别哪种卡)
'      strPass=要验证的密码,在调用程序中输入
'      bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'      strAdvance:新导诊系统自助有效时传入身份识别刷卡的卡信息:
'                         格式为:卡类别ID||卡号
'      lng病人ID=对病人进行补充验证时使用
'========================================================================================================
'返回：空或：
'       0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8《病人ID》
'       9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
'       19帐户增加累计,20帐户支出累计,21进入统筹累计,22统筹报销累计,23住院次数累计;24就诊类型 (1、急诊门诊);25开单科室
'========================================================================================================
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空填充；
'------------------------------------------------------------------------------------------------------------------
'调用模块：1115-自助挂号
    Dim intMouse As Integer
    Dim lng病人ID_Temp As Long
    Dim str医保号 As String
    
    If intinsure = 0 Then
        intinsure = InsureChoose()
        If intinsure = 0 Then Exit Function
    End If
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then intinsure = 0: Exit Function
    
    intMouse = Screen.MousePointer
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        Identify2 = gobjInsure_Obj(mintOrder).Identify2(strCard, strPass, bytType, lng病人ID, intinsure, strAdvance)
    Else
        If strCard = "" Then Exit Function
        Select Case intinsure
            Case TYPE_吉林
                '20040923刘兴宏加入
                Identify2 = ""
            Case TYPE_江苏, TYPE_南京市, TYPE_余姚, TYPE_华东
                Identify2 = ""
            Case TYPE_重庆市
                Identify2 = ""
            Case TYPE_重庆松藻
                Identify2 = ""
            Case TYPE_自贡市
                Identify2 = ""
            Case TYPE_泸州市
                Identify2 = 身份标识_泸州2(strCard, strPass, lng病人ID)
            Case TYPE_铜仁
                Identify2 = ""
            Case TYPE_贵阳市
                Identify2 = ""
            Case TYPE_云南省, TYPE_昆明市
                Identify2 = ""
            Case TYPE_云南建水
                Identify2 = ""
            Case TYPE_成都市
                Identify2 = 身份标识_成都2(strCard, strPass, IIf(bytType = 3, 0, IIf(bytType = 4, 1, bytType)), lng病人ID)
            Case TYPE_成都莲合
                Identify2 = ""
            Case TYPE_开县
                '20050124
                Identify2 = ""
                
            Case type_成都郊县, TYPE_新都
                Identify2 = ""
            Case TYPE_成都南充
                Identify2 = ""
            Case TYPE_咸阳市
                Identify2 = ""
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                Identify2 = 身份标识_福建巨龙(intinsure, 操作方式.挂号, lng病人ID)
            Case type_米易
                Identify2 = ""
            Case TYPE_四川眉山
                Identify2 = ""
            Case TYPE_沈阳市
                Identify2 = ""
            Case TYPE_大连市, TYPE_大连开发区
                '刘兴宏(200403)
                Identify2 = 身份标识_大连2(strCard, strPass, lng病人ID, intinsure)
            Case TYPE_涪陵
                Identify2 = ""
            Case TYPE_广元
                Identify2 = ""
            Case TYPE_乐山
                Identify2 = ""
            Case TYPE_重大校园卡
                '刘兴宏(200403)
                Identify2 = 身份标识_重大校园卡2(strCard, strPass, lng病人ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                Identify2 = ""
            Case TYPE_重大校园卡
                '刘兴宏(20040705)
                Identify2 = ""
            Case TYPE_黔南
                '刘兴宏:20041022
                Identify2 = ""
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                Identify2 = ""
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                Identify2 = ""
            Case TYPE_成都内江
                '刘兴宏:20041207
                Identify2 = ""
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                Identify2 = ""
            Case TYPE_南充阆中
                '刘兴宏:20050419
                Identify2 = ""
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                Identify2 = ""
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                Identify2 = ""
                
            Case TYPE_渝北农医
                Identify2 = ""
            Case Is > 900
                Identify2 = ""
        End Select
    End If
    
    '更新病人信息中的医保号
    If Identify2 <> "" And IsZLHIS10 Then
        str医保号 = Split(Identify2, ";")(1)
        lng病人ID_Temp = Val(Split(Identify2, ";")(8))
        gstrSQL = "zl_病人信息_更新信息(" & lng病人ID_Temp & ",'医保号','" & str医保号 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病人医保号")
    End If
    
    '还原鼠标形状，因为接口很容易中间退出，此时尚未恢复
    Screen.MousePointer = intMouse
    If Identify2 = "" Then intinsure = 0
End Function

Public Function Identify(Optional bytType As Byte, Optional lng病人ID As Long = 0, _
    Optional ByRef intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:识别指定人员是否为参保病人，返回病人的信息
    '入参:bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '     strAdvance-格式用|线分离,第一个为验证类型:(1-部分结算重新收费身份验证;2-医保补结算(二次结算)身份验证;空 -非部分结算或二次结算);第二位后暂没有,待以后扩展
    '                    如:1|...
    '出参: strAdvance参数:由医保接口返回该参数,主要决定病人就诊的就诊模式及当前挂号费用的收取方式
    '   格式:结算模式|挂号费收取方式,用|分隔
    '   a)  结算模式
    '   主要决定病人就诊过程中是采取哪种就诊模式（先诊疗后结算；还是先结算后诊疗）,0-表示先结算后诊疗;1-表示先诊疗后结算
    '   b)  挂号费收取方式
    '   表示当前的挂号费采用什么方式收取(记帐或现收): 1表示挂号费采用记帐方式；0表示用现收方式

    '返回:空或：
    '       0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8《病人ID》
    '       9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
    '       19帐户增加累计,20帐户支出累计,21进入统筹累计,22统筹报销累计,23住院次数累计;24就诊类型 (1、急诊门诊);25开单科室
    '编制:刘兴洪
    '日期:2014-10-13 17:14:55
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空填充；
    '调用模块：1121-门诊收费,1131-办理登记,1137-结帐办理,1124-保险补充结算
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim intMouse As Integer
    Dim lng病人ID_Temp As Long
    Dim str医保号 As String
    Dim str原姓名 As String, str新姓名 As String
    Dim rsTemp As New ADODB.Recordset
    
    If intinsure = 0 Then
        intinsure = InsureChoose()
        If intinsure = 0 Then Exit Function
    End If
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then intinsure = 0: Exit Function
    
    '只有挂号，且非分离部分时再检查
    If bytType = 3 And Not IsApartComponents(intinsure) Then
        '挂号不能进行身份验证
        Select Case intinsure
            'Modified By 朱玉宝 2004-05-25 原因：医保接口变动
            '------------------------------------------------
            Case TYPE_成都市, TYPE_福建巨龙, TYPE_南平市, TYPE_福建省, TYPE_江苏, _
            TYPE_福州市, TYPE_四川眉山, TYPE_泸州市, TYPE_重庆壁山, TYPE_沈阳市, _
            TYPE_重大校园卡, TYPE_大连开发区, TYPE_大连市, TYPE_贵阳市, TYPE_重庆市, _
            TYPE_重庆渝北, TYPE_宁海, TYPE_渝北农医, TYPE_吉林, TYPE_广元旺苍, TYPE_南京市, TYPE_徐州六院
            '------------------------------------------------
            Case Else
                intinsure = 0
                Exit Function
        End Select
    End If
    
    '提取原姓名
    If lng病人ID <> 0 Then
        gstrSQL = " Select 姓名 From 病人信息 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取原姓名", lng病人ID)
        If rsTemp.RecordCount <> 0 Then
            str原姓名 = Nvl(rsTemp!姓名)
        End If
    End If
    
    intMouse = Screen.MousePointer
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder) Then Exit Function
        Identify = gobjInsure_Obj(mintOrder).Identify(bytType, lng病人ID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_昭通
                Identify = 身份标识_昭通(bytType, lng病人ID)
            Case TYPE_徐州农保
                Identify = 身份标识_徐州农保(bytType, lng病人ID)
            Case TYPE_徐州
                Identify = 身份标识_徐州(bytType, lng病人ID)
            Case TYPE_徐州六院
                Identify = 身份标识_中联(bytType, lng病人ID, intinsure)
            Case TYPE_宁海
                Identify = 身份标识_宁海(bytType, lng病人ID)
            Case TYPE_余姚
                Identify = 身份标识_余姚(bytType, lng病人ID)
            Case TYPE_浙江
                Identify = 身份标识_浙江(bytType, lng病人ID)
            Case TYPE_吉林
                '20040923刘兴宏加入
                Identify = 身份标识_吉林(bytType, lng病人ID)
            Case TYPE_北京尚洋
                Identify = 身份标识_北京尚洋(bytType, lng病人ID)
            Case TYPE_华东
                Identify = 身份标识_华东(bytType, lng病人ID)
            Case TYPE_南京市
                Identify = 身份标识_南京市(bytType, lng病人ID)
            Case TYPE_江苏
                Identify = 身份标识_江苏(bytType, lng病人ID)
            Case TYPE_涪陵
                Identify = 身份标识_涪陵(bytType, lng病人ID)
            Case TYPE_广元
                Identify = 身份标识_广元(bytType, lng病人ID)
            Case TYPE_重庆市
                Identify = 身份标识_重庆(bytType, lng病人ID)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                Identify = 身份标识_松藻(bytType, lng病人ID)
            Case TYPE_重庆壁山
                Identify = 身份标识_壁山(bytType, lng病人ID)
            Case TYPE_重庆中梁山
                Identify = 身份标识_中梁山(bytType, lng病人ID)
            Case TYPE_自贡市
                '直接调用子函数
                Identify = 身份标识_中软(bytType, lng病人ID)
            Case TYPE_四川自贡
                '直接调用子函数
                Identify = 身份标识_自贡(bytType, lng病人ID)
            Case TYPE_泸州市
                '直接调用子函数
                Identify = 身份标识_泸州(bytType, lng病人ID)
            Case TYPE_铜仁
                '直接调用子函数
                Identify = 身份标识_铜仁(bytType, lng病人ID)
            Case TYPE_贵阳市
                Identify = 身份标识_贵阳(bytType, lng病人ID)
            Case TYPE_云南省, TYPE_昆明市
                Identify = 身份标识_云南(bytType, lng病人ID, intinsure)
            Case TYPE_云南建水
                Identify = 身份标识_云南建水(bytType, lng病人ID)
            Case TYPE_成都市
                Identify = 身份标识_成都(IIf(bytType = 3, 0, IIf(bytType = 4, 1, bytType)), lng病人ID)
            Case TYPE_成都莲合
                Identify = 身份标识_莲合(bytType, lng病人ID)
            Case TYPE_开县
                '20050124
                Identify = 身份标识_开县(bytType, lng病人ID)
            Case TYPE_新都
                Identify = 身份标识_新都(bytType, lng病人ID)
            Case type_成都郊县
                Identify = 身份标识_成都郊县(bytType, lng病人ID)
            Case TYPE_成都南充
                Identify = 身份标识_成都南充(bytType, lng病人ID)
            Case TYPE_咸阳市
                Identify = 身份标识_咸阳(bytType, lng病人ID)
            Case type_米易
                Identify = 身份标识_米易(bytType, lng病人ID)
            Case TYPE_四川眉山
                Identify = 身份标识_眉山(bytType, lng病人ID)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                Identify = ""
                If InStr(1, "0123", bytType) = 0 Then intinsure = 0: Exit Function
                Identify = 身份标识_福建巨龙(intinsure, IIf(bytType = 0, 操作方式.收费, IIf(bytType = 1, 操作方式.入院, IIf(bytType = 3, 操作方式.挂号, 操作方式.验证))), lng病人ID)
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                Identify = 身份标识_大连(bytType, lng病人ID, intinsure)
            Case TYPE_沈阳市
                Identify = 身份标识_沈阳市(bytType, lng病人ID)
            Case TYPE_乐山
                Identify = 身份标识_乐山(bytType, lng病人ID)
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                Identify = 身份标识_重大校园卡(bytType, lng病人ID)
            Case TYPE_重庆银海版
                Identify = 身份标识_重庆银海版(bytType, lng病人ID)
            Case TYPE_北京
                Identify = 身份标识_北京(bytType, lng病人ID)
            Case TYPE_重庆渝北
                '刘兴宏(20040705)
                Identify = 身份标识_重庆渝北(bytType, lng病人ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                Identify = 身份标识_兴安(bytType, lng病人ID)
            Case TYPE_毕节
                Identify = 身份标识_毕节(bytType, lng病人ID)
            Case TYPE_黔南
                '刘兴宏:20041022
                Identify = 身份标识_黔南(bytType, lng病人ID)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                Identify = 身份标识_奉庆(bytType, lng病人ID)
            Case TYPE_成都德阳
                '刘兴宏:20041103
                Identify = 身份标识_成都德阳(bytType, lng病人ID)
            Case TYPE_成都内江
                '刘兴宏:20041207
                Identify = 身份标识_成都内江(bytType, lng病人ID)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                Identify = 身份标识_广元旺苍(bytType, lng病人ID)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                Identify = 身份标识_南充阆中(bytType, lng病人ID)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                Identify = 身份标识_兴成(bytType, lng病人ID)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                Identify = 身份标识_神木大兴(bytType, lng病人ID)
                
            Case TYPE_渝北农医
                Identify = 身份标识_渝北农医(bytType, lng病人ID)
            Case TYPE_慈溪农医
                Identify = 身份标识_慈溪农医(bytType, lng病人ID)
            Case TYPE_山西
                '陈东 20050311
                Identify = 身份标识_山西(bytType, lng病人ID)
            Case TYPE_冕宁
                '陈东 20050401
                Identify = 身份标识_冕宁(bytType, lng病人ID)
            Case type_铜梁合医
                Identify = 身份标识_铜梁合医(bytType, lng病人ID, type_铜梁合医)
                '陈东 20050530
            Case TYPE_铜山县
                Identify = 身份标识_铜山县(bytType, lng病人ID, intinsure)
            Case Is > 900 '中联医保
                Identify = 身份标识_中联(bytType, lng病人ID, intinsure)
        End Select
    End If
    
    '更新病人信息中的医保号
    If Identify <> "" And IsZLHIS10 Then
        Dim array信息
        Dim lng中心 As Long, lng年龄 As Long
        Dim strSQL As String
        Dim rsPati As New ADODB.Recordset
        
        array信息 = Split(Identify, ";")
        lng中心 = Val(array信息(9))
        str医保号 = array信息(1)
        lng病人ID_Temp = Val(array信息(8))
        
        '取年龄
        If IsDate(array信息(5)) Then
            lng年龄 = Int(zlDatabase.Currentdate() - CDate(array信息(5))) / 365
        End If
        gstrSQL = "zl_病人信息_更新信息(" & lng病人ID_Temp & ",'医保号','" & str医保号 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病人医保号")
        
        '挂号或收费，允许更新病人信息；其它则仅提示
        str新姓名 = Split(Identify, ";")(3)
        If (str原姓名 <> "" And str原姓名 <> str新姓名) Then
            If bytType = 0 Or bytType = 3 Then
                strSQL = "Select A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.卡验证码, A.费别, A.医疗付款方式, A.姓名, A.性别, A.年龄, A.出生日期, A.出生地点, A.身份证号, A.其他证件, A.身份, A.职业, A.民族, A.国籍, A.区域, A.学历, A.婚姻状况, A.家庭地址," & vbNewLine & _
                    "      A.家庭电话, A.家庭地址邮编 As 户口邮编, A.监护人, A.联系人姓名, A.联系人关系, A.联系人地址, A.联系人电话, A.合同单位id, A.工作单位, A.单位电话, A.单位邮编, A.单位开户行, A.单位帐号, A.担保人, A.担保额, A.担保性质, A.就诊时间, A.就诊状态," & vbNewLine & _
                    "      A.就诊诊室, A.住院次数, A.当前科室id, A.当前病区id, A.当前床号, A.入院时间, A.出院时间, A.在院, A.Ic卡号, A.健康号, A.医保号, A.险类, A.查询密码, A.登记时间, A.停用时间, A.锁定," & vbNewLine & _
                    "      B.医保号 From 病人信息 A," & _
                    "   (Select * From 保险帐户" & _
                    "   Where 险类=[1] And 医保号=[2] And Nvl(中心,0)=[3]) B" & _
                    " Where A.病人ID=B.病人ID(+) and A.病人ID=[4]" '可能病人ID已经确定
                Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "提取病人信息", intinsure, str医保号, lng中心, lng病人ID_Temp)
                
                strSQL = "zl_病人信息_Update(" & _
                    lng病人ID & "," & IIf(IsNull(rsPati!门诊号), "NULL", rsPati!门诊号) & "," & _
                    IIf(IsNull(rsPati!住院号), "NULL", rsPati!住院号) & ",'" & IIf(IsNull(rsPati!费别), "", rsPati!费别) & "'," & _
                    "'" & IIf(IsNull(rsPati!医疗付款方式), "", rsPati!医疗付款方式) & "'," & _
                    "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                    "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                    "'" & IIf(IsNull(rsPati!出生地点), "", rsPati!出生地点) & "','" & array信息(6) & "'," & _
                    "'" & IIf(IsNull(rsPati!身份), "", rsPati!身份) & "','" & IIf(IsNull(rsPati!职业), "", rsPati!职业) & "'," & _
                    "'" & IIf(IsNull(rsPati!民族), "", rsPati!民族) & "','" & IIf(IsNull(rsPati!国籍), "", rsPati!国籍) & "'," & _
                    "'" & IIf(IsNull(rsPati!学历), "", rsPati!学历) & "','" & IIf(IsNull(rsPati!婚姻状况), "", rsPati!婚姻状况) & "'," & _
                    "'" & IIf(IsNull(rsPati!家庭地址), "", rsPati!家庭地址) & "','" & IIf(IsNull(rsPati!家庭电话), "", rsPati!家庭电话) & "'," & _
                    "'" & IIf(IsNull(rsPati!户口邮编), "", rsPati!户口邮编) & "','" & IIf(IsNull(rsPati!联系人姓名), "", rsPati!联系人姓名) & "'," & _
                    "'" & IIf(IsNull(rsPati!联系人关系), "", rsPati!联系人关系) & "','" & IIf(IsNull(rsPati!联系人地址), "", rsPati!联系人地址) & "'," & _
                    "'" & IIf(IsNull(rsPati!联系人电话), "", rsPati!联系人电话) & "'," & IIf(IsNull(rsPati!合同单位ID), "NULL", rsPati!合同单位ID) & "," & _
                    "'" & array信息(7) & "','" & IIf(IsNull(rsPati!单位电话), "", rsPati!单位电话) & "'," & _
                    "'" & IIf(IsNull(rsPati!单位邮编), "", rsPati!单位邮编) & "','" & IIf(IsNull(rsPati!单位开户行), "", rsPati!单位开户行) & "'," & _
                    "'" & IIf(IsNull(rsPati!单位帐号), "", rsPati!单位帐号) & "','" & IIf(IsNull(rsPati!担保人), "", rsPati!担保人) & "'," & _
                    "" & IIf(IsNull(rsPati!担保额), "NULL", rsPati!担保额) & "," & intinsure & ",0,null,null,null,null,null,null,null,null,null,null,null,null,1111)"
                Call zlDatabase.ExecuteProcedure(strSQL, "更新病人信息")
            Else
                MsgBox "病人的基本资料（姓名：" & str原姓名 & "）与医保验证后得到的基本资料（姓名：" & str新姓名 & "）不符，请在病人信息模块中更新！", vbInformation, gstrSysName
            End If
        End If
    End If
    
    '还原鼠标形状，因为接口很容易中间退出，此时尚未恢复
    Screen.MousePointer = intMouse
    If Identify = "" Then intinsure = 0
End Function

Public Function SelfBalance(病人ID As Long, str医保号 As String, Optional bytFlag As Byte, Optional cur透支额 As Currency = 0, _
    Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Currency
'功能: 提取参保病人个人帐户余额
'参数: bytFlag
'          个位表示余额类型：0-所有余额,1-本年余额,2-往年余额
'          十位表示调用位置：10-门诊,20-入院,30-预交,40-结算
'返回: 返回个人帐户余额的金额
'------------------------------------------------------------------------------------------------------------------
'调用模块：1103-预交收款,1121-门诊收费,1131-办理登记,1137-结帐办理
    Dim bytYear As Byte, bytPlace As Byte
    
    Screen.MousePointer = vbHourglass
    
    bytYear = bytFlag Mod 10        '得到个位
    bytPlace = (bytFlag \ 10) * 10  '得到十位
    '问题号:116144,焦博,2017/12/18,医保病人缴预交时,鼠标的图标会变成忙碌状态
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Screen.MousePointer = vbDefault: Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Screen.MousePointer = vbDefault: Exit Function
        SelfBalance = gobjInsure_Obj(mintOrder).SelfBalance(病人ID, str医保号, bytFlag, cur透支额, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_四川自贡
                SelfBalance = 个人余额_自贡(病人ID)
            Case TYPE_昭通
                SelfBalance = 个人余额_昭通(病人ID)
            Case TYPE_徐州农保
                SelfBalance = 0
            Case TYPE_徐州
                SelfBalance = 个人余额_徐州(病人ID)
            Case TYPE_徐州六院
                SelfBalance = 个人余额_中联(病人ID, intinsure)
            Case TYPE_宁海
                SelfBalance = 个人余额_宁海(病人ID)
            Case TYPE_余姚
                SelfBalance = 个人余额_余姚(病人ID)
            Case TYPE_浙江
                SelfBalance = 个人余额_浙江(病人ID)
                cur透支额 = 50000
            Case TYPE_吉林
                '刘兴宏:20040925
                SelfBalance = 个人余额_吉林(病人ID)
            Case TYPE_北京尚洋
                SelfBalance = 个人余额_北京尚洋(病人ID)
            Case TYPE_华东
                SelfBalance = 个人余额_华东(病人ID)
            Case TYPE_江苏
                SelfBalance = 个人余额_江苏(病人ID)
            Case TYPE_南京市
                SelfBalance = 个人余额_南京市(病人ID)
            Case TYPE_涪陵
                SelfBalance = 个人余额_涪陵(病人ID)
            Case TYPE_广元
                SelfBalance = 个人余额_广元(病人ID)
            Case TYPE_重庆市
                SelfBalance = 个人余额_重庆(str医保号)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                SelfBalance = 个人余额_松藻(病人ID)
            Case TYPE_重庆壁山
                SelfBalance = 个人余额_壁山(CStr(病人ID))
            Case TYPE_重庆中梁山
                SelfBalance = 个人余额_中梁山(CStr(病人ID))
            Case TYPE_自贡市
                '读IC卡上的余额，可能需要装钱
                SelfBalance = 个人余额_中软(病人ID)
            Case TYPE_泸州市
                '读IC卡上的余额，可能需要装钱
                SelfBalance = 个人余额_泸州(病人ID)
            Case TYPE_铜仁
                '读IC卡上的余额，可能需要装钱
                SelfBalance = 个人余额_铜仁(病人ID)
            Case TYPE_贵阳市
                SelfBalance = 个人余额_贵阳(str医保号)
            Case TYPE_云南省, TYPE_昆明市
                SelfBalance = 个人余额_云南(病人ID, str医保号, bytPlace, intinsure)
            Case TYPE_云南建水
                SelfBalance = 个人余额_云南建水(str医保号, bytPlace)
            Case TYPE_成都市
                SelfBalance = 个人余额_成都(str医保号, bytYear)
            Case TYPE_成都莲合
                SelfBalance = 个人余额_莲合()
            Case TYPE_开县
                '20050124
                SelfBalance = 个人余额_开县()
                
            Case TYPE_新都
                SelfBalance = 个人余额_新都(str医保号, bytPlace)
            Case type_成都郊县
                SelfBalance = 个人余额_成都郊县(str医保号, bytPlace)
            Case TYPE_咸阳市
                SelfBalance = 个人余额_咸阳(病人ID)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                SelfBalance = 个人余额_福建巨龙(病人ID, intinsure)
            Case type_米易
                SelfBalance = 个人余额_米易(True, "", IIf(bytFlag = 10, False, True))
            Case TYPE_四川眉山
                SelfBalance = 个人余额_眉山(str医保号)
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                SelfBalance = 个人余额_大连(病人ID, intinsure)
            Case TYPE_沈阳市
                SelfBalance = 个人余额_沈阳市(str医保号)
            Case TYPE_乐山
                SelfBalance = 个人余额_乐山(str医保号)
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                SelfBalance = 个人余额_重大校园卡(病人ID, bytPlace)
            Case TYPE_重庆银海版
                SelfBalance = 个人余额_重庆银海版(str医保号)
            Case TYPE_重庆渝北
                SelfBalance = 个人余额_重庆渝北(病人ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                SelfBalance = 个人余额_兴安(病人ID)
            Case TYPE_北京
                SelfBalance = 个人余额_北京(病人ID)
            Case TYPE_毕节
                SelfBalance = 个人余额_毕节(病人ID)
            Case TYPE_黔南
                '刘兴宏:20041022
                SelfBalance = 个人余额_黔南(病人ID)
                cur透支额 = 10000
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                SelfBalance = 个人余额_奉庆(病人ID)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                SelfBalance = 个人余额_成都德阳(病人ID)
            Case TYPE_成都内江
                '刘兴宏:20041207
                SelfBalance = 个人余额_成都内江(病人ID)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                SelfBalance = 个人余额_广元旺苍(病人ID)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                SelfBalance = 个人余额_南充阆中(病人ID)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                SelfBalance = 个人余额_兴成(病人ID)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                SelfBalance = 个人余额_神木大兴(病人ID, cur透支额)
                
            Case TYPE_渝北农医
                SelfBalance = 个人余额_渝北农医(str医保号)
            Case TYPE_慈溪农医
                SelfBalance = 个人余额_慈溪农医(str医保号)
            Case TYPE_山西
                SelfBalance = 个人余额_山西(str医保号)
            Case TYPE_冕宁
            '20050401 陈东
                SelfBalance = 个人余额_冕宁(str医保号)
            Case type_铜梁合医
            '20050530 陈东
                SelfBalance = 个人余额_铜梁合医(病人ID, str医保号, type_铜梁合医)
            Case TYPE_铜山县
                SelfBalance = 个人余额_铜山县(病人ID, intinsure)
            Case Is > 900
                '中联医保
                SelfBalance = 个人余额_中联(病人ID, intinsure)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Function

Public Function ClinicPreSwap(rsDetail As ADODB.Recordset, str结算方式 As String, _
    Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将门诊收费的明细进行计算,以得出病人的报销金额
    '入参:rsDetail     费用明细(传入)（字段:单据序号,费别,NO,序号,实际票号,结算时间,病人ID,收费类别,收据费目,计算单位,开单人,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保,保险编码,摘要,是否急诊,开单部门ID,执行部门ID)
    '     strAdvance: 格式为:结算类型|...
    '                 第一位:结算类型分别用(空或0;1,2,3表示):
    '                        空或0:代表门诊收费时的预结算;
    '                        1：代表退费时，如果发生重新收费调用的预结算,表示重新收费调用
    '                        2:保险补充结算发生的预结算
    '                        3:保险补充结算发生部分退费时发生的重新预结算
    '                 第二位:待以后业务扩展
    '出参: str结算方式  "报销方式;金额;是否允许修改|...."
    '           10.34.0以上：
    '               a)一次结算或分单据结算，格式： 报销方式;金额;是否允许修改|...
    '               b)一次结算分单据退费，格式：单据序号:报销方式;金额;是否允许修改|...||单据序号:报销方式;金额;是否允许修改|...||...
    '返回:交易成功返回true；否则，返回false
    '编制:刘兴洪
    '日期:2014-10-13 17:29:01
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '调用模块：1121-门诊收费(含部分退费的重新收费);1124-保险补充结算
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        ClinicPreSwap = gobjInsure_Obj(mintOrder).ClinicPreSwap(rsDetail, str结算方式, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_徐州农保
                ClinicPreSwap = 门诊虚拟结算_徐州农保(rsDetail, str结算方式, strAdvance)
            Case TYPE_四川自贡
                ClinicPreSwap = 门诊虚拟结算_自贡(rsDetail, str结算方式)
            Case TYPE_徐州
                ClinicPreSwap = 门诊虚拟结算_徐州(rsDetail, str结算方式)
            Case TYPE_徐州六院
                ClinicPreSwap = 门诊虚拟结算_徐州六院(rsDetail, str结算方式, strAdvance)
            Case TYPE_宁海
                ClinicPreSwap = 门诊虚拟结算_宁海(rsDetail, str结算方式)
            Case TYPE_余姚
                ClinicPreSwap = 门诊虚拟结算_余姚(rsDetail, str结算方式)
            Case TYPE_浙江
                ClinicPreSwap = 门诊虚拟结算_浙江(rsDetail, str结算方式)
            Case TYPE_华东
                ClinicPreSwap = 门诊虚拟结算_华东(rsDetail, str结算方式)
            Case TYPE_南京市
                ClinicPreSwap = 门诊虚拟结算_南京市(rsDetail, str结算方式, strAdvance)
            Case TYPE_江苏
                ClinicPreSwap = 门诊虚拟结算_江苏(rsDetail, str结算方式)
            Case TYPE_涪陵
                ClinicPreSwap = 门诊虚拟结算_涪陵(rsDetail, str结算方式)
            Case TYPE_广元
                ClinicPreSwap = 门诊虚拟结算_广元(rsDetail, str结算方式)
            Case TYPE_重庆市
                ClinicPreSwap = 门诊虚拟结算_重庆(rsDetail, str结算方式, strAdvance)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                ClinicPreSwap = 门诊虚拟结算_松藻(rsDetail, str结算方式)
            Case TYPE_重庆壁山
                ClinicPreSwap = 门诊虚拟结算_壁山(rsDetail, str结算方式)
            Case TYPE_重庆中梁山
                ClinicPreSwap = False
            Case TYPE_自贡市, Is > 900
                ClinicPreSwap = 门诊虚拟结算(rsDetail, str结算方式, intinsure)
            Case TYPE_泸州市
                ClinicPreSwap = 门诊虚拟结算_泸州(rsDetail, str结算方式)
            Case TYPE_铜仁
                ClinicPreSwap = 门诊虚拟结算_铜仁(rsDetail, str结算方式)
            Case TYPE_贵阳市
                ClinicPreSwap = 门诊虚拟结算_贵阳(rsDetail, str结算方式, strAdvance)
            Case TYPE_云南省, TYPE_昆明市
                ClinicPreSwap = 门诊虚拟结算_云南(rsDetail, str结算方式, intinsure)
            Case TYPE_云南建水
                ClinicPreSwap = 门诊虚拟结算_云南建水(rsDetail, str结算方式)
            Case TYPE_成都市
                ClinicPreSwap = False
            Case TYPE_成都莲合
                ClinicPreSwap = False
            Case TYPE_开县
                '20050124
                ClinicPreSwap = False
                
            Case TYPE_新都
                ClinicPreSwap = 门诊虚拟结算_新都(rsDetail, str结算方式)
            Case type_成都郊县
                ClinicPreSwap = 门诊虚拟结算_成都郊县(rsDetail, str结算方式)
            Case TYPE_咸阳市
                ClinicPreSwap = 门诊虚拟结算_咸阳(rsDetail, str结算方式)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                ClinicPreSwap = 门诊虚拟结算_福建巨龙(rsDetail, str结算方式, intinsure)
            Case type_米易
                ClinicPreSwap = 门诊虚拟结算_米易(rsDetail, str结算方式)
            Case TYPE_四川眉山
                ClinicPreSwap = 门诊虚拟结算_眉山(rsDetail, str结算方式)
            Case TYPE_沈阳市
                ClinicPreSwap = 门诊虚拟结算_沈阳市(rsDetail, str结算方式)
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                ClinicPreSwap = 门诊虚拟结算_大连(rsDetail, str结算方式, intinsure)
            Case TYPE_乐山
                ClinicPreSwap = 门诊虚拟结算_乐山(rsDetail, str结算方式)
            Case TYPE_重大校园卡
                  '刘兴宏(20040705)
                ClinicPreSwap = 门诊虚拟结算_重大校园卡(rsDetail, str结算方式)
            Case TYPE_重庆银海版
                ClinicPreSwap = 门诊虚拟结算_重庆银海版(rsDetail, str结算方式)
            Case TYPE_重庆渝北
                ClinicPreSwap = 门诊虚拟结算_重庆渝北(rsDetail, str结算方式)
            Case TYPE_兴安
                '刘兴宏:20040827
                ClinicPreSwap = 门诊虚拟结算_兴安(rsDetail, str结算方式)
            Case TYPE_北京
                ClinicPreSwap = 门诊虚拟结算_北京(rsDetail, str结算方式)
            Case TYPE_毕节
                ClinicPreSwap = 门诊虚拟结算_毕节(rsDetail, str结算方式)
            Case TYPE_黔南
                '刘兴宏:20041022
                ClinicPreSwap = 门诊虚拟结算_黔南(rsDetail, str结算方式)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                ClinicPreSwap = 门诊虚拟结算_奉庆(rsDetail, str结算方式)
            Case TYPE_成都德阳
                '刘兴宏:20041103
                ClinicPreSwap = 门诊虚拟结算_成都德阳(rsDetail, str结算方式)
            Case TYPE_成都内江
                '刘兴宏:20041207
                ClinicPreSwap = 门诊虚拟结算_成都内江(rsDetail, str结算方式)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                ClinicPreSwap = 门诊虚拟结算_广元旺苍(rsDetail, str结算方式)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                ClinicPreSwap = 门诊虚拟结算_南充阆中(rsDetail, str结算方式)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                ClinicPreSwap = 门诊虚拟结算_兴成(rsDetail, str结算方式)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                ClinicPreSwap = 门诊虚拟结算_神木大兴(rsDetail, str结算方式)
                        
            Case TYPE_渝北农医
                ClinicPreSwap = 门诊虚拟结算_渝北农医(rsDetail, str结算方式)
            Case TYPE_慈溪农医
                ClinicPreSwap = 门诊虚拟结算_慈溪农医(rsDetail, str结算方式)
            Case TYPE_山西
                ClinicPreSwap = 门诊虚拟结算_山西(rsDetail, str结算方式)
            Case TYPE_冕宁
                ClinicPreSwap = 门诊虚拟结算_冕宁(rsDetail, str结算方式)
            '20050401 陈东
            Case type_铜梁合医
                ClinicPreSwap = 门诊虚拟结算_铜梁合医(rsDetail, str结算方式, type_铜梁合医)
                '20050530 陈东
            Case TYPE_铜山县
                ClinicPreSwap = 门诊虚拟结算_铜山县(rsDetail, str结算方式, intinsure)
        End Select
    End If
End Function

Public Function ClinicSwap(lng结帐ID As Long, cur个帐支付 As Currency, cur医保基金 As Currency, cur全自付 As Currency, _
    cur先自付 As Currency, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '入参:lng结帐ID-收费记录或补充结算的结帐ID；，从预交记录中可以检索医保号和密码
    '    strAdvance-格式:结算序号|NO|...,每位|分隔
    '       第一位:结算序号(即:病人预交记录.结算序号),以负数区分是否多单据一次结算还是分单据结算
    '       第二位:NO:当前结算的NO
    '       第三位后:待以后扩展
    '出参:strAdvance:如果为多单据一次结算的,则不分单据，全部按标准形式返回，格式为：结算方式|金额||结算方式|金额
    '                 如果为多单据产生多次结算的,则按老方式处理
    '               10.34.0以上：
    '               a)一次结算或分单据结算，格式：结算方式|金额||结算方式|金额
    '               b)一次结算分单据退费，格式：NO:结算方式,金额|结算方式,金额|...||NO:结算方式,金额|结算方式,金额|...||...
    '     注意：
    '           strAdvance在10.34.0以前(不含补允结算)
    '               多单据一次结算时,传入的是结帐IDs:结帐ID1,结帐ID2,...
    '               其他，传入格式为:单据总张数|当前第几张单据
    '     特别说明：strAdvance请不要轻易返回空，返回空表示以预结算结果为准。如果报销金额为0也请按标准形式返回，如:个人帐户|0||医保基金|0
    '返回:交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。
    '但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，
    '但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。'
    '这样才能保证数据的完全统一
    '调用模块:
    '   1121-门诊收费(含部分退费时重收),1124-医保补充结算
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int单据_MAX As Integer
    Dim str医保号 As String, str密码 As String
    Dim lng病人ID As Long, str卡号 As String
    Dim rsTmp As New ADODB.Recordset, varData As Variant
    Dim lng结算序号 As Long, bln补充结算 As Boolean
    Dim strAdvanceIn As String
    
    strAdvanceIn = strAdvance
    varData = Split(strAdvance & "||||||", "|")
    lng结算序号 = Val(varData(0)): bln补充结算 = False
    If lng结算序号 < 0 Then
        '1.负数表示多单据一次结算(即多单据收费只产生一个结帐ID)
        '10.34.0以前(不含补允结算)strAdvance传入格式不同，但通过"lng结算序号<0"仍能判断出是否为补充结算
        '  还需要检查是否补充结算数据
        gstrSQL = "Select 1 From 病人预交记录 where 结帐ID=[1] and mod(记录性质,10)=6 and rownum <2"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否补充结算", lng结帐ID)
        bln补充结算 = Not rsTmp.EOF
    End If
    
    With rsTmp
        'Modified by ZYB 2002-10-28
        '松藻医保不需要对每个项目设置对应的医保项目
        If Not GetCapability(support允许不设置医保项目, 0, intinsure) Then
            If bln补充结算 Then
                gstrSQL = "" & _
                "Select 名称 From (Select C.名称,A.TEST From" & _
                "   (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
                "   (Select distinct 收费细目ID,'Y' as TEST " & _
                "    From 门诊费用记录 " & _
                "    Where   Nvl(附加标志,0)<>9  And Nvl(实收金额,0)<>0 " & _
                "       And 结帐ID in (select distinct 收费结帐ID From 费用补充记录 where 结算ID =[2]) ) B,收费细目 C" & _
                " Where B.收费细目ID=C.ID And B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
            Else
                gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                    " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
                    " (Select 收费细目ID,'Y' as TEST From 门诊费用记录 Where 结帐ID=[2] And Nvl(附加标志,0)<>9 And Nvl(实收金额,0)<>0) B,收费细目 C" & _
                    " Where B.收费细目ID=C.ID And B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "必须设置医保项目", intinsure, lng结帐ID)
            If Not rsTmp.EOF Then
                MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行医保交易。", vbExclamation, gstrSysName
                ClinicSwap = False: Exit Function
            End If
        End If
        
        'Modified by 朱玉宝（考虑到医保某次结算，不一定存在个人帐户和医保基金的情况，为了保证仍然调用医保接口）
'        If GetCapability(support门诊必须传递明细) = False Then
'            gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
'                " From 病人预交记录 A,保险帐户 B " & _
'                " Where A.病人ID=B.病人ID And B.险类=" & intInsure & _
'                " And A.结算方式 in ('个人帐户','医保基金') And A.结帐ID=" & lng结帐ID
'        Else
            '病人该次收费可能没有发生个人帐户费用，便也要传递费用明细
            gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
                " From 病人预交记录 A,保险帐户 B " & _
                " Where A.病人ID=B.病人ID And B.险类=[1] And A.结帐ID=[2]"
'        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "病人该次收费可能没有发生个人帐户费用，便也要传递费用明细", intinsure, lng结帐ID)
        If rsTmp.RecordCount = 0 Then
            MsgBox "没有使用个人帐户或医保基金结算，可以不向医保交易", vbExclamation, gstrSysName
            ClinicSwap = True: Exit Function
        End If
    End With
    
    lng病人ID = rsTmp!病人ID
    str卡号 = TrimStr(IIf(IsNull(rsTmp!卡号), "", rsTmp!卡号))
    str医保号 = TrimStr(IIf(IsNull(rsTmp!医保号), str卡号, rsTmp!医保号))
    str密码 = TrimStr(IIf(IsNull(rsTmp!密码), "", rsTmp!密码))
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    On Error Resume Next
    Dim cnBalance As New ADODB.Connection
    '记录下本次的交易情况，以便核对：中心成功HIS失败的情况
    If bln补充结算 Then
        gstrSQL = "" & _
        " Select 1 AS 性质,A.NO,B.医保号,C.姓名,SUM(A.实收金额) AS 费用总额" & _
        " From 门诊费用记录 A,保险帐户 B,病人信息 C" & _
        " Where A.病人ID=B.病人ID ANd B.病人ID=C.病人ID And B.险类=" & intinsure & _
        "       And C.病人ID=[1]   " & _
        "       And A.结帐ID in (select distinct 收费结帐ID From 费用补充记录 where 结算ID =[2])" & _
        " Group By A.NO,B.医保号,C.姓名"
    Else
        gstrSQL = "" & _
        " Select 1 AS 性质,A.NO,B.医保号,C.姓名,SUM(A.实收金额) AS 费用总额" & _
        " From 门诊费用记录 A,保险帐户 B,病人信息 C" & _
        " Where A.病人ID=B.病人ID ANd B.病人ID=C.病人ID And B.险类=" & intinsure & _
        " And C.病人ID=[1] And A.结帐ID=[2]" & _
        " Group By A.NO,B.医保号,C.姓名"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", lng病人ID, lng结帐ID)
    
    gstrSQL = "zl_结算日志_Insert(" & rsTmp!性质 & ",'" & rsTmp!NO & "','" & rsTmp!医保号 & "','" & rsTmp!姓名 & "'," & rsTmp!费用总额 & ")"
    Set cnBalance = GetNewConnection
    On Error GoTo 0
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        ClinicSwap = gobjInsure_Obj(mintOrder).ClinicSwap(lng结帐ID, cur个帐支付, cur医保基金, cur全自付, cur先自付, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_徐州农保
                ClinicSwap = 门诊结算_徐州农保(lng结帐ID, cur个帐支付, str医保号, strAdvance)
            Case TYPE_四川自贡
                ClinicSwap = 门诊结算_自贡(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_昭通
                ClinicSwap = 门诊结算_昭通(lng结帐ID, cur个帐支付, str医保号, cur全自付, strAdvance)
            Case TYPE_徐州
                ClinicSwap = 门诊结算_徐州(lng结帐ID, cur个帐支付, str医保号, cur全自付, cur先自付, cur医保基金)
            Case TYPE_徐州六院
                ClinicSwap = 门诊结算_中联(lng结帐ID, cur个帐支付, lng病人ID, cur全自付, cur先自付, intinsure, strAdvance)
            Case TYPE_宁海
                ClinicSwap = 门诊结算_宁海(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_余姚
                ClinicSwap = 门诊结算_余姚(lng结帐ID, cur个帐支付, str医保号, cur全自付)
            Case TYPE_浙江
                ClinicSwap = 门诊结算_浙江(lng结帐ID, cur个帐支付, str医保号, cur全自付, cur先自付, cur医保基金)
            Case TYPE_吉林
                '20040923刘兴宏加入
                ClinicSwap = 门诊结算_吉林(lng结帐ID, cur个帐支付, str医保号, cur全自付, strAdvance)
            Case TYPE_北京尚洋
                ClinicSwap = 门诊结算_北京尚洋(lng结帐ID, cur个帐支付, str医保号, cur全自付, strAdvance)
            Case TYPE_华东
                ClinicSwap = 门诊结算_华东(lng结帐ID, cur个帐支付, str医保号, cur全自付)
            Case TYPE_南京市
                ClinicSwap = 门诊结算_南京市(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_江苏
                ClinicSwap = 门诊结算_江苏(lng结帐ID, cur个帐支付, str医保号, cur全自付)
            Case TYPE_涪陵
                ClinicSwap = 门诊结算_涪陵(lng结帐ID, cur个帐支付, str医保号, cur全自付, strAdvance)
            Case TYPE_广元
                ClinicSwap = 门诊结算_广元(lng结帐ID, cur个帐支付, str医保号, cur全自付)
            Case TYPE_重庆市
                ClinicSwap = 门诊结算_重庆(lng结帐ID, cur个帐支付, str医保号, strAdvance)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                ClinicSwap = 门诊结算_松藻(lng结帐ID, cur个帐支付, lng病人ID, cur全自付, cur先自付)
            Case TYPE_重庆壁山
                ClinicSwap = 门诊结算_壁山(lng结帐ID, cur个帐支付, str医保号, cur全自付, cur先自付, cur医保基金)
            Case TYPE_重庆中梁山
                ClinicSwap = 门诊结算_中梁山(lng结帐ID, cur个帐支付, lng病人ID, cur全自付, cur先自付)
            Case TYPE_自贡市
                ClinicSwap = 门诊结算_中软(lng结帐ID, cur个帐支付, cur全自付, cur先自付)
            Case TYPE_泸州市
                ClinicSwap = 门诊结算_泸州(lng结帐ID, cur个帐支付, cur全自付, cur先自付)
            Case TYPE_铜仁
                ClinicSwap = 门诊结算_铜仁(lng结帐ID, cur个帐支付, cur全自付, cur先自付)
            Case TYPE_贵阳市
                ClinicSwap = 门诊结算_贵阳(lng结帐ID, cur个帐支付, str医保号, False, strAdvance)
            Case TYPE_云南省, TYPE_昆明市
                ClinicSwap = 门诊结算_云南(lng结帐ID, cur个帐支付, str医保号, intinsure, strAdvance)
            Case TYPE_云南建水
                ClinicSwap = 门诊结算_云南建水(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_成都市
                ClinicSwap = 门诊结算_成都(lng结帐ID, lng病人ID, str医保号, str密码, str卡号, cur全自付)
            Case TYPE_成都莲合
                ClinicSwap = 门诊结算_莲合(lng结帐ID)
            Case TYPE_开县
                '20050124
                ClinicSwap = 门诊结算_开县(lng结帐ID)
            Case type_成都郊县
                ClinicSwap = 门诊结算_成都郊县(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_新都
                ClinicSwap = 门诊结算_新都(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_咸阳市
                ClinicSwap = 门诊结算_咸阳(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                ClinicSwap = 门诊结算_福建巨龙(lng病人ID, lng结帐ID, intinsure, strAdvance)
            Case type_米易
                ClinicSwap = 门诊结算_米易(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_四川眉山
                ClinicSwap = 门诊结算_眉山(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_沈阳市
                ClinicSwap = 门诊结算_沈阳市(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_大连市, TYPE_大连开发区
                '200311
                ClinicSwap = 门诊结算_大连(lng结帐ID, cur个帐支付, str医保号, intinsure)
            Case TYPE_乐山
                ClinicSwap = 门诊结算_乐山(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_重庆银海版
                ClinicSwap = 门诊结算_重庆银海版(lng结帐ID, cur个帐支付, str医保号, strAdvance)
            Case TYPE_重大校园卡
                '刘兴宏200403
                ClinicSwap = 门诊结算_重大校园卡(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_重庆渝北
                '刘兴宏20040705
                ClinicSwap = 门诊结算_重庆渝北(lng结帐ID, cur个帐支付, str医保号, strAdvance)
            Case TYPE_北京
                ClinicSwap = 门诊结算_北京(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_兴安
                '刘兴宏:20040827
                ClinicSwap = 门诊结算_兴安(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_毕节
                ClinicSwap = 门诊结算_毕节(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_黔南
                '刘兴宏:20041022
                ClinicSwap = 门诊结算_黔南(lng结帐ID, cur个帐支付, str医保号, cur全自付)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                ClinicSwap = 门诊结算_奉庆(lng结帐ID, cur个帐支付, str医保号, cur全自付)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                ClinicSwap = 门诊结算_成都德阳(lng结帐ID, cur个帐支付, str医保号, cur全自付)
            Case TYPE_成都内江
                '刘兴宏:20041207
                ClinicSwap = 门诊结算_成都内江(lng结帐ID, cur个帐支付, str医保号, cur全自付, strAdvance)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                ClinicSwap = 门诊结算_广元旺苍(lng结帐ID, cur个帐支付, str医保号, cur全自付)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                ClinicSwap = 门诊结算_南充阆中(lng结帐ID, cur个帐支付, str医保号, cur全自付)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                ClinicSwap = 门诊结算_兴成(lng结帐ID, cur个帐支付, str医保号, cur全自付)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                ClinicSwap = 门诊结算_神木大兴(lng结帐ID, cur个帐支付, str医保号, cur全自付)
                
            Case TYPE_渝北农医
                ClinicSwap = 门诊结算_渝北农医(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_慈溪农医
                ClinicSwap = 门诊结算_慈溪农医(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_山西
                ClinicSwap = 门诊结算_山西(lng结帐ID, cur个帐支付, str医保号)
            Case TYPE_冕宁
                ClinicSwap = 门诊结算_冕宁(lng结帐ID, cur个帐支付, str医保号)
            '20050401 陈东
            Case type_铜梁合医
                ClinicSwap = 门诊结算_铜梁合医(lng结帐ID, cur个帐支付, str医保号, type_铜梁合医)
                '20050530 陈东
            Case TYPE_铜山县
                ClinicSwap = 门诊结算_铜山县(lng结帐ID, cur个帐支付, str医保号, intinsure)
            Case Is > 900
                ClinicSwap = 门诊结算_中联(lng结帐ID, cur个帐支付, lng病人ID, cur全自付, cur先自付, intinsure)
        End Select
    End If
    
    If strAdvance <> "" Then
        '如果返回的不是校正串，需要清为空，因为新的HIS程序在结算时，strAdvance传入了值
        If strAdvance = strAdvanceIn Then
            strAdvance = ""
        Else
            '说明是返回的校正串,啥也不处理
        End If
    End If
End Function

Public Function ClinicDelSwap(lngStlID As Long, Optional ByVal bln退费 As Boolean = True, _
    Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将门诊退费的明细和结算数据转发送医保前置服务器确认
    '入参:lngStlID-将要退的费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '     bln退费 -表明是退费交易还是改费交易在调用本接口
    '     strAdvance:格式:冲销ID|补充结算标志|…,每位|分隔
    '           第一位:传入冲销ID,医保可以根据冲销ID来进行取数
    '           第二位:补充结算标志,1-补充结算调和;0非补充结算调用
    '           第三位:NO:当前结算的NO
    '           第四位后: 待以后扩展
    '     注意：
    '           strAdvance在10.34.0以前(不含补允结算)
    '               多单据一次结算时,传入的是原结帐IDs:结帐ID1,结帐ID2,...
    '               其他，传入格式为:退费单据总张数|当前退第几张单据
    '出参:strAdvance:1.原样退回时，返回空
    '                2.退费结算方式与收费结算方式不一致时，返回格式为：结算方式|金额||结算方式|金额||…（其中，金额为负）
    '返回：交易成功返回true；否则，返回false
    '编制:刘兴洪
    '日期:2014-10-14 09:40:40
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '调用模块：1121-门诊退费;1124-保险补充结算
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim strSelfNo As String, strSelfPwd As String, str卡号 As String
    Dim rsTmp As New ADODB.Recordset, int结算性质 As Integer
    Dim curMoney As Currency
    Dim lng病人ID As Long
    Dim lng冲销ID As Long
    Dim varData As Variant
    Dim bln补充结算 As Boolean
    Dim strNO As String
    Dim strAdvanceIn As String
    
    strAdvanceIn = strAdvance
    '取冲销ID并更新原始结算记录的冲销ID与冲销时间
    varData = Split(strAdvance & "|||||", "|")
    lng冲销ID = Val(varData(0))
    bln补充结算 = Val(varData(1)) = 1
    strNO = varData(2)
    
    If bln退费 = False Then
        If Not GetCapability(support门诊改费, 0, intinsure) Then
            MsgBox "该医保不支持改费功能！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    'Modified by 朱玉宝
'    '读出个人帐户使用金额
'    If GetCapability(support门诊必须传递明细) = False Then
'        gstrSQL = "Select A.病人ID,A.冲预交,A.结算方式 From 病人预交记录 A " & _
'            " Where A.结算方式 in ('个人帐户','医保基金') And A.结帐ID=" & lngStlID
'    Else
        '病人该次收费可能没有发生个人帐户费用，便也要传递费用明细
        gstrSQL = "Select A.病人ID,A.冲预交,A.结算方式,A.结算序号,A.结算性质 From 病人预交记录 A " & _
            " Where A.结帐ID=[1]"
    
'    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "读出个人帐户使用金额", lngStlID)
    If rsTmp.RecordCount = 0 And intinsure <> TYPE_余姚 Then
        MsgBox "没有使用个人帐户或医保基金结算，可以不向医保交易", vbExclamation, gstrSysName
        ClinicDelSwap = True
        Exit Function
    End If
    int结算性质 = Val(Nvl(rsTmp!结算性质))
    
    lng病人ID = rsTmp("病人ID")
    If Val(Nvl(rsTmp!结算序号)) >= 0 And int结算性质 <> 2 Then '10.34.0以前（不含补充结算）未传冲销ID
        lng冲销ID = 0
        bln补充结算 = False
    End If
    
    Do Until rsTmp.EOF
        If rsTmp("结算方式") = "个人帐户" Then
            curMoney = curMoney + rsTmp("冲预交")
        End If
        rsTmp.MoveNext
    Loop
    
    'Modified by ZYB 2002-10-28
    '松藻医保不需要对每个项目设置对应的医保项目
    If Not GetCapability(support允许不设置医保项目, 0, intinsure) Then
        If bln补充结算 Then
            gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
            " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
            " ( Select distinct 收费细目ID,'Y' as TEST From 门诊费用记录  " & _
            "   Where   Nvl(附加标志,0)<>9 And Nvl(实收金额,0)<>0 " & _
            "         And 结帐ID in (Select distinct 收费结帐ID From 费用补充记录 where 结算ID=[1])) B,收费细目 C" & _
            " Where B.收费细目ID=C.ID and B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
        Else
            gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
            " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
            " (Select 收费细目ID,'Y' as TEST From 门诊费用记录 Where 结帐ID=[2] And Nvl(附加标志,0)<>9 And Nvl(实收金额,0)<>0) B,收费细目 C" & _
            " Where B.收费细目ID=C.ID and B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "必须设置医保项目", intinsure, lngStlID)
        If Not rsTmp.EOF Then
            MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行医保交易。", vbExclamation, gstrSysName
            ClinicDelSwap = False: Exit Function
        End If
    End If
    
    On Error Resume Next
    If lng冲销ID = 0 Then
        gstrSQL = "select A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
                  " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]" & _
                  " order by A.登记时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取冲销ID并更新原始结算记录的冲销ID与冲销时间", lngStlID)
        lng冲销ID = rsTmp("结帐ID")
    End If
    
    gstrSQL = "zl_保险结算记录_作废(" & lngStlID & "," & lng冲销ID & ",'" & strNO & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将作废记录与原始记录关联起来")
    On Error GoTo 0
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        ClinicDelSwap = gobjInsure_Obj(mintOrder).ClinicDelSwap(lngStlID, bln退费, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_徐州农保
                ClinicDelSwap = 门诊结算冲销_徐州农保(lngStlID, curMoney, lng病人ID, strAdvance)
            Case TYPE_四川自贡
                ClinicDelSwap = 门诊结算冲销_自贡(lngStlID, curMoney, lng病人ID)
            Case TYPE_昭通
                ClinicDelSwap = 门诊结算冲销_昭通(lngStlID, curMoney, lng病人ID)
            Case TYPE_徐州
                ClinicDelSwap = 门诊结算冲销_徐州(lngStlID, curMoney, lng病人ID)
            Case TYPE_徐州六院
                ClinicDelSwap = 门诊结算冲销_中联(lngStlID, curMoney, lng病人ID, intinsure, strAdvance)
            Case TYPE_宁海
                ClinicDelSwap = 门诊结算冲销_宁海(lngStlID, curMoney, lng病人ID)
            Case TYPE_余姚
                ClinicDelSwap = 门诊结算冲销_余姚(lngStlID, curMoney, lng病人ID)
            Case TYPE_浙江
                ClinicDelSwap = 门诊结算冲销_浙江(lngStlID, curMoney, lng病人ID)
            Case TYPE_吉林
                '20040923刘兴宏加入
                ClinicDelSwap = 门诊结算冲销_吉林(lngStlID, curMoney, lng病人ID)
            Case TYPE_北京尚洋
                ClinicDelSwap = 门诊结算冲销_北京尚洋(lngStlID, curMoney, lng病人ID)
            Case TYPE_华东
                ClinicDelSwap = 门诊结算冲销_华东(lngStlID, curMoney, lng病人ID)
            Case TYPE_南京市
                ClinicDelSwap = 门诊结算冲销_南京市(lngStlID, curMoney, lng病人ID)
            Case TYPE_江苏
                ClinicDelSwap = 门诊结算冲销_江苏(lngStlID, curMoney, lng病人ID)
            Case TYPE_涪陵
                ClinicDelSwap = 门诊结算冲销_涪陵(lngStlID, curMoney, lng病人ID)
            Case TYPE_广元
                ClinicDelSwap = 门诊结算冲销_广元(lngStlID, curMoney, lng病人ID)
            Case TYPE_重庆市
                ClinicDelSwap = 门诊结算冲销_重庆(lngStlID, curMoney, lng病人ID)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                ClinicDelSwap = 门诊结算冲销_松藻(lngStlID, curMoney, lng病人ID)
            Case TYPE_重庆壁山
                ClinicDelSwap = 门诊结算冲销_壁山(lngStlID, curMoney, lng病人ID)
            Case TYPE_重庆中梁山
                ClinicDelSwap = 门诊结算冲销_中梁山(lngStlID, curMoney, lng病人ID)
            Case TYPE_自贡市
                ClinicDelSwap = 门诊结算冲销_中软(lngStlID, curMoney)
            Case TYPE_泸州市
                ClinicDelSwap = 门诊结算冲销_泸州(lngStlID, curMoney)
            Case TYPE_铜仁
                ClinicDelSwap = 门诊结算冲销_铜仁(lngStlID, curMoney)
            Case TYPE_云南省, TYPE_昆明市
                ClinicDelSwap = 门诊结算冲销_云南(lngStlID, curMoney, lng病人ID, intinsure)
            Case TYPE_云南建水
                '由于没有yh_recedefeebalance，所以不支持
                ClinicDelSwap = False
            Case TYPE_贵阳市
                ClinicDelSwap = 门诊结算冲销_贵阳(lngStlID, curMoney, lng病人ID)
            Case TYPE_成都市
                ClinicDelSwap = False
            Case TYPE_成都莲合
                ClinicDelSwap = False
            Case TYPE_开县
                '20050124
                ClinicDelSwap = False
                
            Case type_成都郊县, TYPE_新都
                ClinicDelSwap = False
            Case TYPE_咸阳市
                ClinicDelSwap = False
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                ClinicDelSwap = 门诊结算冲销_福建巨龙(lngStlID, lng病人ID, intinsure)
            Case type_米易
                ClinicDelSwap = 门诊结算冲销_米易(lngStlID, curMoney, lng病人ID)
            Case TYPE_四川眉山
                ClinicDelSwap = 门诊结算冲销_眉山(lngStlID, curMoney, lng病人ID)
            Case TYPE_沈阳市
                ClinicDelSwap = 门诊结算冲销_沈阳市(lngStlID, curMoney, lng病人ID, bln退费)
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                ClinicDelSwap = 门诊结算冲销_大连(lngStlID, curMoney, lng病人ID, intinsure)
            Case TYPE_乐山
                ClinicDelSwap = 门诊结算冲销_乐山(lngStlID, curMoney, lng病人ID)
            Case TYPE_重大校园卡
                '刘兴宏200403
                ClinicDelSwap = 门诊结算冲销_重大校园卡(lngStlID, curMoney, lng病人ID)
            Case TYPE_重庆银海版
                ClinicDelSwap = 门诊结算冲销_重庆银海版(lngStlID, curMoney, lng病人ID)
            Case TYPE_重庆渝北
                '刘兴宏20040705
                ClinicDelSwap = 门诊结算冲销_重庆渝北(lngStlID, curMoney, lng病人ID)
            Case TYPE_北京
                ClinicDelSwap = 门诊结算冲销_北京(lngStlID, curMoney, lng病人ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                ClinicDelSwap = 门诊结算冲销_兴安(lngStlID, curMoney, lng病人ID)
            Case TYPE_毕节
                ClinicDelSwap = 门诊结算冲销_毕节(lngStlID, curMoney, lng病人ID)
            Case TYPE_黔南
                '刘兴宏:20041022
                ClinicDelSwap = 门诊结算冲销_黔南(lngStlID, curMoney, lng病人ID)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                ClinicDelSwap = 门诊结算冲销_奉庆(lngStlID, curMoney, lng病人ID)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                ClinicDelSwap = 门诊结算冲销_成都德阳(lngStlID, curMoney, lng病人ID)
            Case TYPE_成都内江
                '刘兴宏:20041207
                ClinicDelSwap = 门诊结算冲销_成都内江(lngStlID, curMoney, lng病人ID)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                ClinicDelSwap = 门诊结算冲销_广元旺苍(lngStlID, curMoney, lng病人ID)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                ClinicDelSwap = 门诊结算冲销_南充阆中(lngStlID, curMoney, lng病人ID)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                ClinicDelSwap = 门诊结算冲销_兴成(lngStlID, curMoney, lng病人ID)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                ClinicDelSwap = 门诊结算冲销_神木大兴(lngStlID, curMoney, lng病人ID)
                
            Case TYPE_渝北农医
                ClinicDelSwap = 门诊结算冲销_渝北农医(lngStlID, curMoney, lng病人ID)
            Case TYPE_慈溪农医
                ClinicDelSwap = 门诊结算冲销_慈溪农医(lngStlID, curMoney, lng病人ID)
            Case TYPE_山西
                ClinicDelSwap = 门诊结算冲销_山西(lngStlID, curMoney, lng病人ID)
            Case TYPE_冕宁
                '20050401 陈东
                ClinicDelSwap = 门诊结算冲销_冕宁(lngStlID, curMoney, lng病人ID)
    
            Case type_铜梁合医
                '20050530  陈东
                ClinicDelSwap = 门诊结算冲销_铜梁合医(lngStlID, curMoney, lng病人ID, type_铜梁合医)
            Case TYPE_铜山县
                ClinicDelSwap = 门诊结算冲销_铜山县(lngStlID, curMoney, lng病人ID, intinsure)
            Case Is > 900
                ClinicDelSwap = 门诊结算冲销_中联(lngStlID, curMoney, lng病人ID, intinsure)
        End Select
    End If
    
    If ClinicDelSwap = False Then Exit Function
    '106880,strAdvance返回格式检查
    If strAdvanceIn = strAdvance Then strAdvance = ""
    ClinicDelSwap = CheckClinicDelSwapAdvanceOutFormat(strAdvance)
End Function

Private Function CheckClinicDelSwapAdvanceOutFormat(ByVal strAdvance As String) As Boolean
    '医保退费接口(ClinicDelSwap)中strAdvance参数返回值格式检查
    Dim varData As Variant, varItem As Variant
    Dim i As Integer
    
    On Error GoTo ErrHandler
    If strAdvance = "" Then
        CheckClinicDelSwapAdvanceOutFormat = True
        Exit Function
    End If
    
    '出参:strAdvance:1.原样退回时，返回空
    '                2.退费结算方式与收费结算方式不一致时，返回格式为：结算方式|金额||结算方式|金额||…（其中，金额为负）
    varData = Split(strAdvance, "||")
    For i = 0 To UBound(varData)
        varItem = Split(varData(i), "|")
        If UBound(varItem) < 1 Then
            MsgBox "医保退费接口(ClinicDelSwap)中strAdvance参数返回值格式不正确！" & vbCrLf & _
                "当前返回值：" & strAdvance & vbCrLf & _
                vbCrLf & _
                "strAdvance的返回值正确格式：" & vbCrLf & _
                "1.原样退回时，返回空" & vbCrLf & _
                "2.退费结算方式与收费结算方式不一致时，返回格式为：结算方式1|金额1||结算方式2|金额2||…（其中，金额为负）", vbExclamation, gstrSysName
            Exit Function
        Else
            If IsNumeric(varItem(1)) = False Then
                MsgBox "医保退费接口(ClinicDelSwap)中strAdvance参数返回值格式不正确！" & vbCrLf & _
                    "当前返回值：" & strAdvance & vbCrLf & _
                    vbCrLf & _
                    "strAdvance的返回值正确格式：" & vbCrLf & _
                    "1.原样退回时，返回空" & vbCrLf & _
                    "2.退费结算方式与收费结算方式不一致时，返回格式为：结算方式1|金额1||结算方式2|金额2||…（其中，金额为负）", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    Next
    
    CheckClinicDelSwapAdvanceOutFormat = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function TransferSwap(lng预交ID As Long, curMoney As Currency, _
    Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1103-预交收款,1131-办理登记
    Dim str医保号 As String
    Dim rsTmp As New ADODB.Recordset

    On Error Resume Next
    With rsTmp
        gstrSQL = "Select B.病人ID,A.主页ID,A.金额,B.卡号,B.密码,B.医保号,B.顺序号 " & _
            " From 病人预交记录 A,保险帐户 B " & _
            " Where A.病人ID=B.病人ID And A.结算方式='个人帐户'" & _
            " And A.ID=[1] And B.险类=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保报销结果", lng预交ID, intinsure)
        If rsTmp.EOF Then
            MsgBox "没有使用个人帐户或非参保病人，不能交易！", vbExclamation, gstrSysName
            TransferSwap = True: Exit Function
        End If
    End With
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        TransferSwap = gobjInsure_Obj(mintOrder).TransferSwap(lng预交ID, curMoney, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_涪陵, TYPE_华东, TYPE_北京尚洋, TYPE_余姚
                TransferSwap = False
            Case TYPE_广元
                TransferSwap = False
            Case TYPE_自贡市
                TransferSwap = 个人帐户转预交_中软(lng预交ID, curMoney)
            Case TYPE_泸州市
                TransferSwap = 个人帐户转预交_泸州(lng预交ID, curMoney)
            Case TYPE_铜仁
                TransferSwap = 个人帐户转预交_铜仁(lng预交ID, curMoney)
            Case TYPE_云南省, TYPE_昆明市
                '顺序号采用入院验证时返回的
    '            If IsNull(rsTmp!顺序号) Then
    '                MsgBox "未能发现该病人的住院交易顺序号，不能交易！", vbExclamation, gstrSysName
    '                Exit Function
    '            End If
    '            str医保号 = TrimStr(IIf(IsNull(rsTmp("医保号")), "", rsTmp("医保号")))
    '            TransferSwap = 个人帐户转预交_云南(lng预交ID, curMoney, str医保号, rsTmp("顺序号"), rsTmp("病人ID"))
                MsgBox "不支持个人帐户转预交！", vbInformation, gstrSysName
                TransferSwap = False
            Case TYPE_云南建水
                '顺序号采用入院验证时返回的
                If IsNull(rsTmp!顺序号) Then
                    MsgBox "未能发现该病人的住院交易顺序号，不能交易！", vbExclamation, gstrSysName
                    Exit Function
                End If
                str医保号 = TrimStr(IIf(IsNull(rsTmp("医保号")), "", rsTmp("医保号")))
                TransferSwap = 个人帐户转预交_云南建水(lng预交ID, curMoney, str医保号, rsTmp("顺序号"), rsTmp("病人ID"))
            Case TYPE_重庆市
                TransferSwap = False
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28 重庆松藻不支持个人帐户转预交
                TransferSwap = False
            Case TYPE_重庆壁山
                TransferSwap = False
            Case TYPE_重庆中梁山
                TransferSwap = False
            Case TYPE_贵阳市
                TransferSwap = False
            Case TYPE_成都市
                TransferSwap = 个人帐户转预交_成都(lng预交ID, curMoney, rsTmp)
            Case TYPE_成都莲合
                TransferSwap = 个人帐户转预交_莲合(lng预交ID, curMoney)
            Case TYPE_开县
                '20050124
                TransferSwap = 个人帐户转预交_开县(lng预交ID, curMoney)
                
            Case type_成都郊县
                str医保号 = TrimStr(IIf(IsNull(rsTmp("医保号")), "", rsTmp("医保号")))
                TransferSwap = 个人帐户转预交_成都郊县(lng预交ID, curMoney, str医保号, rsTmp("顺序号"), rsTmp("病人ID"))
            Case TYPE_新都
                str医保号 = TrimStr(IIf(IsNull(rsTmp("医保号")), "", rsTmp("医保号")))
                TransferSwap = 个人帐户转预交_新都(lng预交ID, curMoney, str医保号, rsTmp("顺序号"), rsTmp("病人ID"))
            Case TYPE_咸阳市
                TransferSwap = False
            Case type_米易
                TransferSwap = False
            Case TYPE_四川眉山
                TransferSwap = False
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311),无预交
                TransferSwap = False
            Case TYPE_乐山
                TransferSwap = False
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                 TransferSwap = 个人帐户转预交_重大校园卡(lng预交ID, curMoney, rsTmp)
            Case TYPE_重庆渝北
                  '刘兴宏(20040705)
                TransferSwap = False
            Case TYPE_兴安
                '刘兴宏:20040827
                TransferSwap = False
            Case TYPE_黔南
                '刘兴宏:20041022
                TransferSwap = False
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                TransferSwap = False
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                TransferSwap = False
            Case TYPE_成都内江
                '刘兴宏:20041207
                TransferSwap = False
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                TransferSwap = False
            Case TYPE_南充阆中
                '刘兴宏:20050419
                TransferSwap = False
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                TransferSwap = False
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                TransferSwap = False
                
            Case TYPE_渝北农医
                TransferSwap = False
            Case TYPE_徐州六院
                TransferSwap = 个人帐户转预交_中联(lng预交ID, curMoney, rsTmp!病人ID, intinsure)
            Case TYPE_铜山县
                TransferSwap = 帐户转预交_铜山县(lng预交ID, curMoney, intinsure)
            Case Is > 900
                TransferSwap = 个人帐户转预交_中联(lng预交ID, curMoney, rsTmp!病人ID, intinsure)
        End Select
    End If
End Function

Public Function TransferDelSwap(lng预交ID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lngPrepID-将要退的预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1103-预交退款
    
    Dim rsTmp As New ADODB.Recordset
    Dim curMoney As Currency, lng病人ID As Long
    
    On Error Resume Next
    With rsTmp
        gstrSQL = "Select B.病人ID,A.金额,B.卡号,B.密码,B.医保号,B.顺序号 " & _
            " From 病人预交记录 A,保险帐户 B " & _
            " Where A.病人ID=B.病人ID And A.结算方式='个人帐户'" & _
            " And A.ID=[1] And B.险类=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取预交记录", lng预交ID, intinsure)
        If rsTmp.EOF Then
            MsgBox "没有使用个人帐户或非参保病人，不能交易！", vbExclamation, gstrSysName
            TransferDelSwap = True: Exit Function
        End If
        
        lng病人ID = rsTmp("病人ID")
        Do Until rsTmp.EOF
            curMoney = curMoney + rsTmp("金额")
            rsTmp.MoveNext
        Loop
    End With
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        TransferDelSwap = gobjInsure_Obj(mintOrder).TransferDelSwap(lng预交ID, curMoney, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_江苏, TYPE_华东, TYPE_北京尚洋, TYPE_余姚
                TransferDelSwap = False
            Case TYPE_涪陵, TYPE_广元
                TransferDelSwap = False
            Case TYPE_自贡市
                TransferDelSwap = False
            Case TYPE_泸州市, TYPE_铜仁
                TransferDelSwap = False
            Case TYPE_云南省, TYPE_昆明市, TYPE_云南建水
                TransferDelSwap = False
            Case TYPE_重庆市
                TransferDelSwap = False
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                TransferDelSwap = False
            Case TYPE_重庆壁山
                TransferDelSwap = False
            Case TYPE_重庆中梁山
                TransferDelSwap = False
            Case TYPE_贵阳市
                TransferDelSwap = False
            Case TYPE_成都市
                TransferDelSwap = False
            Case TYPE_成都莲合
                TransferDelSwap = False
            Case TYPE_开县
                '20050124
                TransferDelSwap = False
                
            Case type_成都郊县, TYPE_新都
                TransferDelSwap = False
            Case TYPE_咸阳市
                TransferDelSwap = False
            Case type_米易
                TransferDelSwap = False
            Case TYPE_四川眉山
                TransferDelSwap = False
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)，无退款
                TransferDelSwap = False
            Case TYPE_乐山
                TransferDelSwap = False
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                TransferDelSwap = 个人帐户转预交冲销_重大校园卡(lng预交ID, curMoney, lng病人ID)
            Case TYPE_重庆渝北
                  '刘兴宏(20040705)
                TransferDelSwap = False
            Case TYPE_兴安
                '刘兴宏:20040827
                TransferDelSwap = False
            Case TYPE_黔南
                '刘兴宏:20041022
                TransferDelSwap = False
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                TransferDelSwap = False
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                TransferDelSwap = False
            Case TYPE_成都内江
                '刘兴宏:20041207
                TransferDelSwap = False
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                TransferDelSwap = False
            Case TYPE_南充阆中
                TransferDelSwap = False
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                TransferDelSwap = False
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                TransferDelSwap = False
                
            Case TYPE_渝北农医
                TransferDelSwap = False
            Case TYPE_徐州六院
                TransferDelSwap = 个人帐户转预交冲销_中联(lng预交ID, curMoney, lng病人ID, intinsure)
            Case TYPE_铜山县
                TransferDelSwap = 帐户转预交冲销_铜山县(lng预交ID, intinsure)
            Case Is > 900
                TransferDelSwap = 个人帐户转预交冲销_中联(lng预交ID, curMoney, lng病人ID, intinsure)
        End Select
    End If
End Function

Public Function WipeoffMoney(rsExse As Recordset, lng病人ID As Long, str医保号 As String, str密码 As String, _
    Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As String
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；str医保号-医保号；str密码-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    'str密码：由于各医保一直未使用，现改为：1-结帐处调用;0-其他地方调用（病人费用查询）
    '------------------------------------------------------------------------------------------------------------------
    '调用模块：1137-结帐办理
    Dim strSort As String
    If rsExse.State = adStateClosed Then Exit Function
    If rsExse.RecordCount = 0 Then
        MsgBox "没有费用明细，或者收费项目没有设置对应的医保编码。", vbInformation, gstrSysName
        Exit Function
    End If
    gbln批量虚拟结算 = (strAdvance = "1")
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        WipeoffMoney = gobjInsure_Obj(mintOrder).WipeoffMoney(rsExse, lng病人ID, str医保号, str密码, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_四川自贡
                WipeoffMoney = 住院虚拟结算_自贡(rsExse, lng病人ID)
            Case TYPE_昭通
                WipeoffMoney = 住院虚拟结算_昭通(rsExse, lng病人ID, str医保号)
            Case TYPE_徐州农保
                WipeoffMoney = 住院虚拟结算_徐州农保(rsExse, lng病人ID)
            Case TYPE_徐州
                WipeoffMoney = 住院虚拟结算_徐州(rsExse, lng病人ID, str医保号)
            Case TYPE_徐州六院
                WipeoffMoney = 住院虚拟结算_徐州六院(rsExse)
            Case TYPE_宁海
                WipeoffMoney = 住院虚拟结算_宁海(rsExse, lng病人ID)
            Case TYPE_余姚
                WipeoffMoney = 住院虚拟结算_余姚(rsExse, lng病人ID, str医保号)
            Case TYPE_浙江
                WipeoffMoney = 住院虚拟结算_浙江(rsExse, lng病人ID, str医保号)
            Case TYPE_吉林
                '20040923刘兴宏加入
                WipeoffMoney = 住院虚拟结算_吉林(rsExse, lng病人ID, IIf(str密码 = "1", True, False))
            Case TYPE_北京尚洋
                WipeoffMoney = 住院虚拟结算_北京尚洋(rsExse, lng病人ID, str医保号)
            Case TYPE_华东
                WipeoffMoney = 住院虚拟结算_华东(rsExse, lng病人ID, str医保号, IIf(str密码 = "1", False, True))
            Case TYPE_南京市
                WipeoffMoney = 住院虚拟结算_南京市(rsExse, lng病人ID, strAdvance)
            Case TYPE_江苏
                WipeoffMoney = 住院虚拟结算_江苏(rsExse, lng病人ID, str医保号)
            Case TYPE_涪陵
                WipeoffMoney = 住院虚拟结算_涪陵(rsExse, lng病人ID, str医保号)
            Case TYPE_广元
                WipeoffMoney = 住院虚拟结算_广元(rsExse, lng病人ID, str医保号)
            Case TYPE_重庆市
                WipeoffMoney = 住院虚拟结算_重庆(rsExse, lng病人ID, str医保号, IIf(str密码 = "1", False, True), strAdvance)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                WipeoffMoney = 住院虚拟结算_松藻(rsExse)
            Case TYPE_重庆壁山
                WipeoffMoney = 住院虚拟结算_壁山(rsExse, lng病人ID, str医保号, str密码)
            Case TYPE_重庆中梁山
                WipeoffMoney = 住院虚拟结算_中梁山(rsExse, lng病人ID)
            Case TYPE_贵阳市
            '    WipeoffMoney = 住院虚拟结算_贵阳(rsExse, lng病人ID, IIf(str密码 = "1", False, True))
                WipeoffMoney = 住院虚拟结算_贵阳(rsExse, lng病人ID, IIf(str密码 = "1", True, False))

            Case TYPE_自贡市
                WipeoffMoney = 住院虚拟结算_中软(rsExse)
            Case TYPE_泸州市
                WipeoffMoney = 住院虚拟结算_泸州(rsExse)
            Case TYPE_铜仁
                WipeoffMoney = 住院虚拟结算_铜仁(rsExse)
            Case TYPE_云南省, TYPE_昆明市
                WipeoffMoney = 住院虚拟结算_云南(rsExse, lng病人ID, intinsure)
            Case TYPE_云南建水
                WipeoffMoney = 住院虚拟结算_云南建水(rsExse, lng病人ID)
            Case TYPE_成都市
                WipeoffMoney = 住院虚拟结算_成都(rsExse, str医保号, str密码)
            Case TYPE_成都莲合
                WipeoffMoney = 住院虚拟结算_莲合(rsExse, str医保号, str密码)
            Case TYPE_开县
                '20050124
                WipeoffMoney = 住院虚拟结算_开县(rsExse, str医保号, str密码)
                
            Case type_成都郊县
                WipeoffMoney = 住院虚拟结算_成都郊县(rsExse, lng病人ID, str医保号)
            Case TYPE_新都
                WipeoffMoney = 住院虚拟结算_新都(rsExse, lng病人ID, str医保号)
            Case TYPE_成都南充
                WipeoffMoney = 住院虚拟结算_成都南充(lng病人ID)
            Case TYPE_咸阳市
                WipeoffMoney = 住院虚拟结算_咸阳(rsExse, lng病人ID, str医保号)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                WipeoffMoney = 住院虚拟结算_福建巨龙(lng病人ID, intinsure)
            Case type_米易
                WipeoffMoney = 住院虚拟结算_米易(rsExse, lng病人ID)
            Case TYPE_四川眉山
                WipeoffMoney = 住院虚拟结算_眉山(rsExse, lng病人ID)
            Case TYPE_沈阳市
                WipeoffMoney = 住院虚拟结算_沈阳市(rsExse, lng病人ID)
            Case TYPE_大连市, TYPE_大连开发区
                '刘兴宏(200311)
                WipeoffMoney = 住院虚拟结算_大连(rsExse, lng病人ID, intinsure)
            Case TYPE_乐山
                WipeoffMoney = 住院虚拟结算_乐山(rsExse, lng病人ID)
            Case TYPE_重大校园卡
                '刘兴宏(200403)
                WipeoffMoney = 住院虚拟结算_重大校园卡(rsExse, lng病人ID)
            Case TYPE_重庆银海版
                WipeoffMoney = 住院虚拟结算_重庆银海版(rsExse, lng病人ID)
            Case TYPE_重庆渝北
                '刘兴宏(20040705)
                WipeoffMoney = 住院虚拟结算_重庆渝北(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
            Case TYPE_北京
                WipeoffMoney = 住院虚拟结算_北京(rsExse, lng病人ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                WipeoffMoney = 住院虚拟结算_兴安(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
            Case TYPE_毕节
                WipeoffMoney = 住院虚拟结算_毕节(rsExse, lng病人ID)
            Case TYPE_黔南
                '刘兴宏:20041022
                WipeoffMoney = 住院虚拟结算_黔南(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                WipeoffMoney = 住院虚拟结算_奉庆(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                WipeoffMoney = 住院虚拟结算_成都德阳(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
            Case TYPE_成都内江
                '刘兴宏:20041207
                WipeoffMoney = 住院虚拟结算_成都内江(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                WipeoffMoney = 住院虚拟结算_广元旺苍(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
            Case TYPE_南充阆中
                '刘兴宏:20050419
                WipeoffMoney = 住院虚拟结算_南充阆中(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                WipeoffMoney = 住院虚拟结算_兴成(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                WipeoffMoney = 住院虚拟结算_神木大兴(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
                
            Case TYPE_渝北农医
                WipeoffMoney = 住院虚拟结算_渝北农医(rsExse, lng病人ID)
            Case TYPE_慈溪农医
                WipeoffMoney = 住院虚拟结算_慈溪农医(rsExse, lng病人ID)
            Case TYPE_山西
                WipeoffMoney = 住院虚拟结算_山西(rsExse, lng病人ID)
            Case TYPE_冕宁
            '20050401 陈东
                WipeoffMoney = 住院虚拟结算_冕宁(rsExse, lng病人ID, str密码)
            Case type_铜梁合医
            '20050530 陈东
                WipeoffMoney = 住院虚拟结算_铜梁合医(rsExse, lng病人ID, type_铜梁合医)
            Case TYPE_铜山县
                WipeoffMoney = 住院虚拟结算_铜山县(rsExse, lng病人ID, str医保号, str密码, intinsure)
            Case Is > 900
                WipeoffMoney = 住院虚拟结算_中联(rsExse, intinsure)
        End Select
    End If
    
    '20040107:周韬:有些接口未填写主页ID
    rsExse.Filter = 0
    strSort = rsExse.Sort
    rsExse.Sort = "主页ID"
    rsExse.MoveLast '取最大的
    g结算数据.主页ID = Nvl(rsExse!主页ID, 0)
    rsExse.Sort = strSort
    DebugTool ("当前主页ID：" & rsExse!主页ID)
    Call Insert虚拟结算数据(lng病人ID, g结算数据.主页ID, WipeoffMoney)
End Function

Public Function SettleSwap(lng结帐ID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID-病人结帐记录ID, 从预交记录中可以检索出各种结算方式及金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。
'        但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用
'        分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。
'        (由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
'------------------------------------------------------------------------------------------------------------------
'调用模块：1137-结帐办理
    
    Dim rsTmp As New ADODB.Recordset, lng病人ID As Long
    
    rsTmp.CursorLocation = adUseClient
    
    'Modified by ZYB 2002-10-28
    '松藻医保不需要对每个项目设置对应的医保项目
    'Modified by ZYB 2003-12-14
    '其它只判断未上传的费用明细
    'Modified by ZYB 2008-06-17
    '只检查正常记录
    If Not GetCapability(support允许不设置医保项目, 0, intinsure) Then
        gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
            " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
            " (Select 收费细目ID,'Y' as TEST From 住院费用记录 Where 记录状态=1 And Nvl(是否上传,0)=0 And Nvl(附加标志,0)<>9 And Nvl(实收金额,0)<>0 And 结帐ID=[2]) B,收费细目 C" & _
            " Where B.收费细目ID=C.ID and B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "support允许不设置医保项目", intinsure, lng结帐ID)
        If Not rsTmp.EOF Then
            MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行交易！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
        
    gstrSQL = " Select A.病人ID,B.卡号,B.医保号,B.密码,B.顺序号" & _
              " From 病人结帐记录 A,医保病人档案 B,医保病人关联表 C" & _
              " Where A.病人ID=C.病人ID And B.险类=C.险类 and B.医保号=C.医保号  And B.险类=[1] And A.ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人结帐记录", intinsure, lng结帐ID)
    DebugTool "ClsInsure:读取成功"
    If rsTmp.EOF Then
        MsgBox "没有发现该病人的结算信息,不能执行交易！", vbExclamation, gstrSysName
        Exit Function
    End If
    DebugTool "ClsInsure:判断成功"
    lng病人ID = rsTmp("病人ID")
    DebugTool "ClsInsure:病人赋值成功"
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        SettleSwap = gobjInsure_Obj(mintOrder).SettleSwap(lng结帐ID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_四川自贡
                SettleSwap = 住院结算_自贡(lng结帐ID, lng病人ID)
            Case TYPE_昭通
                SettleSwap = 住院结算_昭通(lng结帐ID, lng病人ID)
            Case TYPE_徐州农保
                SettleSwap = 住院结算_徐州农保(lng结帐ID, lng病人ID)
            Case TYPE_徐州
                SettleSwap = 住院结算_徐州(lng结帐ID, lng病人ID)
            Case TYPE_徐州六院
                SettleSwap = 住院结算_中联(lng结帐ID, intinsure)
            Case TYPE_宁海
                SettleSwap = 住院结算_宁海(lng结帐ID, lng病人ID)
            Case TYPE_余姚
                SettleSwap = 住院结算_余姚(lng结帐ID)
            Case TYPE_浙江
                SettleSwap = 住院结算_浙江(lng结帐ID, lng病人ID)
            Case TYPE_吉林
                '20040923刘兴宏加入
                SettleSwap = 住院结算_吉林(lng结帐ID, strAdvance)
            Case TYPE_北京尚洋
                SettleSwap = 住院结算_北京尚洋(lng结帐ID, lng病人ID)
            Case TYPE_华东
                SettleSwap = 住院结算_华东(lng结帐ID, lng病人ID)
            Case TYPE_南京市
                SettleSwap = 住院结算_南京市(lng结帐ID, lng病人ID)
            Case TYPE_江苏
                SettleSwap = 住院结算_江苏(lng结帐ID)
            Case TYPE_涪陵
                SettleSwap = 住院结算_涪陵(lng结帐ID)
            Case TYPE_广元
                SettleSwap = 住院结算_广元(lng结帐ID)
            Case TYPE_重庆市
                SettleSwap = 住院结算_重庆(lng结帐ID, lng病人ID, strAdvance)
            Case TYPE_重庆松藻
                SettleSwap = 住院结算_松藻(lng结帐ID)
            Case TYPE_重庆壁山
                SettleSwap = 住院结算_壁山(lng结帐ID, lng病人ID, strAdvance)
            Case TYPE_重庆中梁山
                SettleSwap = 住院结算_中梁山(lng结帐ID)
            Case TYPE_自贡市
                SettleSwap = 住院结算_中软(lng结帐ID)
            Case TYPE_泸州市
                SettleSwap = 住院结算_泸州(lng结帐ID)
            Case TYPE_铜仁
                SettleSwap = 住院结算_铜仁(lng结帐ID)
            Case TYPE_贵阳市
                SettleSwap = 住院结算_贵阳(lng结帐ID, lng病人ID)
            Case TYPE_云南省, TYPE_昆明市
                If IsNull(rsTmp!顺序号) Then
                    MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbExclamation, gstrSysName
                    Exit Function
                End If
                SettleSwap = 住院结算_云南(lng结帐ID, rsTmp!顺序号, lng病人ID, intinsure)
            Case TYPE_云南建水
                If IsNull(rsTmp!顺序号) Then
                    MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbExclamation, gstrSysName
                    Exit Function
                End If
                SettleSwap = 住院结算_云南建水(lng结帐ID, rsTmp!顺序号, lng病人ID)
            Case TYPE_成都市
                SettleSwap = 住院结算_成都(lng结帐ID, rsTmp)
            Case TYPE_成都莲合
                '因为在试结算时，已经保存了，如果不成功，到不到此！所以在此不作保存
                SettleSwap = 住院结算_莲合(lng结帐ID, rsTmp)
            Case TYPE_开县
                '20050124
                SettleSwap = 住院结算_开县(lng结帐ID, rsTmp)
                
            Case TYPE_成都南充
                SettleSwap = 住院结算_成都南充(lng结帐ID, rsTmp)
            Case type_成都郊县
                SettleSwap = 住院结算_成都郊县(lng结帐ID, lng病人ID)
            Case TYPE_新都
                SettleSwap = 住院结算_新都(lng结帐ID, lng病人ID)
            Case TYPE_咸阳市
                SettleSwap = 住院结算_咸阳(lng结帐ID, lng病人ID)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                SettleSwap = 住院结算_福建巨龙(lng病人ID, lng结帐ID, intinsure, strAdvance)
            Case type_米易
                SettleSwap = 住院结算_米易(lng结帐ID, lng病人ID)
            Case TYPE_四川眉山
                SettleSwap = 住院结算_眉山(lng结帐ID, lng病人ID)
            Case TYPE_沈阳市
                SettleSwap = 住院结算_沈阳(lng结帐ID, lng病人ID)
            Case TYPE_大连市, TYPE_大连开发区
                '刘兴宏(200311)
                SettleSwap = 住院结算_大连(lng结帐ID, lng病人ID, intinsure)
            Case TYPE_乐山
                SettleSwap = 住院结算_乐山(lng结帐ID, lng病人ID)
            Case TYPE_重大校园卡
                '刘兴宏(200403)
                SettleSwap = 住院结算_重大校园卡(lng结帐ID, lng病人ID)
            Case TYPE_重庆银海版
                SettleSwap = 住院结算_重庆银海版(lng结帐ID, lng病人ID, strAdvance)
            Case TYPE_重庆渝北
                '刘兴宏(20040705)
                SettleSwap = 住院结算_重庆渝北(lng结帐ID, lng病人ID, strAdvance)
            Case TYPE_北京
                SettleSwap = 住院结算_北京(lng结帐ID, lng病人ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                SettleSwap = 住院结算_兴安(lng结帐ID, lng病人ID, strAdvance)
            Case TYPE_毕节
                SettleSwap = 住院结算_毕节(lng结帐ID, lng病人ID)
            Case TYPE_黔南
                '刘兴宏:20041022
                DebugTool "ClsInsure:开始调用住院结算"
    
                SettleSwap = 住院结算_黔南(lng结帐ID, lng病人ID)
                DebugTool "ClsInsure:调用住院结算成功"
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                SettleSwap = 住院结算_奉庆(lng结帐ID, lng病人ID, strAdvance)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                SettleSwap = 住院结算_成都德阳(lng结帐ID, lng病人ID)
            Case TYPE_成都内江
                '刘兴宏:20041207
                SettleSwap = 住院结算_成都内江(lng结帐ID, lng病人ID)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                SettleSwap = 住院结算_广元旺苍(lng结帐ID, lng病人ID)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                SettleSwap = 住院结算_南充阆中(lng结帐ID, lng病人ID)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                SettleSwap = 住院结算_兴成(lng结帐ID, lng病人ID)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                SettleSwap = 住院结算_神木大兴(lng结帐ID, lng病人ID)
                
            Case TYPE_渝北农医
                SettleSwap = 住院结算_渝北农医(lng结帐ID, lng病人ID)
            Case TYPE_慈溪农医
                SettleSwap = 住院结算_慈溪农医(lng结帐ID, lng病人ID)
            Case TYPE_山西
                SettleSwap = 住院结算_山西(lng结帐ID)
            Case TYPE_冕宁
            '陈东 20050401
                SettleSwap = 住院结算_冕宁(lng结帐ID)
            Case type_铜梁合医
            '陈东 20050530
                SettleSwap = 住院结算_铜梁合医(lng结帐ID, lng病人ID, type_铜梁合医)
            Case TYPE_铜山县
                SettleSwap = 住院结算_铜山县(lng结帐ID, lng病人ID, intinsure, strAdvance)
            Case Is > 900
                '中联医保
                SettleSwap = 住院结算_中联(lng结帐ID, intinsure)
        End Select
    End If
    DebugTool "ClsInsure:结束"
    
    Call Clear虚拟结算数据(lng病人ID, g结算数据.主页ID)
End Function

Public Function SettleDelSwap(lng结帐ID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
'功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
'参数：lng结帐ID=需要作废的结帐单ID号（注意是没作废的那条记录）；
'返回：交易成功返回true；否则，返回false
'注意：1)主要使用结帐恢复交易和费用删除交易；
'      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，
'        在病人费用记录中根据结帐ID查找；
'      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；
'        因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
'------------------------------------------------------------------------------------------------------------------
'调用模块：1137-结帐作废
    Dim lng冲销ID As Long
    Dim rsTmp As New ADODB.Recordset
    
    rsTmp.CursorLocation = adUseClient
    gstrSQL = "Select A.病人ID,B.医保号,B.卡号,B.密码,B.顺序号" & _
        " From 保险结算记录 A,医保病人档案 B,医保病人关联表 C " & _
        " Where A.性质=2 And A.记录ID=[1] And A.险类=[2]" & _
        " And A.险类=B.险类 And B.险类=C.险类 And A.病人ID=C.病人ID And C.医保号=B.医保号"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", lng结帐ID, intinsure)
    If rsTmp.RecordCount = 0 Then
        MsgBox "没有使用医保结算，可以不向医保交易", vbExclamation, gstrSysName
        SettleDelSwap = False: Exit Function
    End If
    
    On Error Resume Next
    '取冲销ID并更新原始结算记录的冲销ID及冲销时间
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取冲销ID并更新原始结算记录的冲销ID及冲销时间", lng结帐ID)
    lng冲销ID = rsTmp("ID") '冲销单据的ID
    gstrSQL = "zl_保险结算记录_作废(" & lng结帐ID & "," & lng冲销ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将作废记录与原始记录关联起来")
    On Error GoTo 0
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        SettleDelSwap = gobjInsure_Obj(mintOrder).SettleDelSwap(lng结帐ID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_四川自贡
                SettleDelSwap = 住院结算冲销_自贡(lng结帐ID)
            Case TYPE_山西
               '陈东 20050304
                SettleDelSwap = 住院结算冲销_山西(lng结帐ID)
            Case TYPE_冕宁
                '陈东 20050401
                SettleDelSwap = 住院结算冲销_冕宁(lng结帐ID)
            Case type_铜梁合医
                '陈东 20050530
                SettleDelSwap = 住院结算冲销_铜梁合医(lng结帐ID, type_铜梁合医)
                
            Case TYPE_昭通
                SettleDelSwap = 住院结算冲销_昭通(lng结帐ID)
            Case TYPE_徐州农保
                SettleDelSwap = 住院结算冲销_徐州农保(lng结帐ID)
            Case TYPE_徐州
                SettleDelSwap = 住院结算冲销_徐州(lng结帐ID)
            Case TYPE_徐州六院
                SettleDelSwap = 住院结算冲销_中联(lng结帐ID, intinsure)
            Case TYPE_宁海
                SettleDelSwap = 住院结算冲销_宁海(lng结帐ID)
            Case TYPE_余姚
                SettleDelSwap = 住院结算冲销_余姚(lng结帐ID)
            Case TYPE_浙江
                SettleDelSwap = 住院结算冲销_浙江(lng结帐ID)
            Case TYPE_吉林
                '20040923刘兴宏加入
                SettleDelSwap = 住院结算冲销_吉林(lng结帐ID)
            Case TYPE_北京尚洋
                SettleDelSwap = 住院结算冲销_北京尚洋(lng结帐ID)
            Case TYPE_华东
                SettleDelSwap = 住院结算冲销_华东(lng结帐ID)
            Case TYPE_南京市
                SettleDelSwap = 住院结算冲销_南京市(lng结帐ID)
            Case TYPE_江苏
                SettleDelSwap = 住院结算冲销_江苏(lng结帐ID)
            Case TYPE_涪陵
                SettleDelSwap = 住院结算冲销_涪陵(lng结帐ID)
            Case TYPE_广元
                SettleDelSwap = 住院结算冲销_广元(lng结帐ID)
            Case TYPE_重庆市
                SettleDelSwap = 住院结算冲销_重庆(lng结帐ID)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                SettleDelSwap = 住院结算冲销_松藻(lng结帐ID)
            Case TYPE_重庆壁山
                SettleDelSwap = 住院结算冲销_壁山(lng结帐ID)
            Case TYPE_重庆中梁山
                SettleDelSwap = 住院结算冲销_中梁山(lng结帐ID)
            Case TYPE_自贡市
                SettleDelSwap = 住院结算冲销_中软(lng结帐ID)
            Case TYPE_泸州市
                SettleDelSwap = 住院结算冲销_泸州(lng结帐ID)
            Case TYPE_铜仁
                SettleDelSwap = 住院结算冲销_铜仁(lng结帐ID)
            Case TYPE_云南省, TYPE_昆明市
                SettleDelSwap = 住院结算冲销_云南(lng结帐ID, intinsure)
            Case TYPE_云南建水
                '暂时没处理取消结帐
                SettleDelSwap = False
            Case TYPE_贵阳市
                SettleDelSwap = 住院结算冲销_贵阳(lng结帐ID)
            Case TYPE_成都市
                SettleDelSwap = 住院结算冲销_成都(lng结帐ID, rsTmp)
            Case TYPE_成都莲合
                SettleDelSwap = 住院结算冲销_莲合(lng结帐ID, rsTmp)
            Case TYPE_开县
                '20050124
                SettleDelSwap = 住院结算冲销_开县(lng结帐ID, rsTmp)
                
            Case type_成都郊县, TYPE_新都
                SettleDelSwap = False
            Case TYPE_咸阳市
                SettleDelSwap = False
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                SettleDelSwap = 住院结算冲销_福建巨龙(lng结帐ID, intinsure)
            Case type_米易
                SettleDelSwap = 住院结算冲销_米易(lng结帐ID)
            Case TYPE_四川眉山
                SettleDelSwap = 住院结算冲销_眉山(lng结帐ID)
            Case TYPE_沈阳市
                SettleDelSwap = 住院结算冲销_沈阳市(lng结帐ID)
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                SettleDelSwap = 住院结算冲销_大连(lng结帐ID, intinsure)
            Case TYPE_乐山
                SettleDelSwap = 住院结算冲销_乐山(lng结帐ID)
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                SettleDelSwap = 住院结算冲销_重大校园卡(lng结帐ID)
            Case TYPE_重庆银海版
                SettleDelSwap = 住院结算冲销_重庆银海版(lng结帐ID)
            Case TYPE_重庆渝北
                  '刘兴宏(20040705)
                SettleDelSwap = 住院结算冲销_重庆渝北(lng结帐ID)
            Case TYPE_北京
                SettleDelSwap = 住院结算冲销_北京(lng结帐ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                SettleDelSwap = 住院结算冲销_兴安(lng结帐ID)
            Case TYPE_毕节
                SettleDelSwap = 住院结算冲销_毕节(lng结帐ID)
            Case TYPE_黔南
                '刘兴宏:20041022
                SettleDelSwap = 住院结算冲销_黔南(lng结帐ID)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                SettleDelSwap = 住院结算冲销_奉庆(lng结帐ID)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                SettleDelSwap = 住院结算冲销_成都德阳(lng结帐ID)
            Case TYPE_成都内江
                '刘兴宏:20041207
                SettleDelSwap = 住院结算冲销_成都内江(lng结帐ID)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                SettleDelSwap = 住院结算冲销_广元旺苍(lng结帐ID)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                SettleDelSwap = 住院结算冲销_南充阆中(lng结帐ID)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                SettleDelSwap = 住院结算冲销_兴成(lng结帐ID)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                SettleDelSwap = 住院结算冲销_神木大兴(lng结帐ID)
                
            Case TYPE_渝北农医
                SettleDelSwap = 住院结算冲销_渝北农医(lng结帐ID)
            Case TYPE_慈溪农医
                SettleDelSwap = 住院结算冲销_慈溪农医(lng结帐ID)
            Case type_铜梁合医
            '20050530 陈东
                SettleDelSwap = 住院结算冲销_铜梁合医(lng结帐ID, type_铜梁合医)
            Case TYPE_铜山县
                SettleDelSwap = 住院结算冲销_铜山县(lng结帐ID, intinsure)
            Case Is > 900
                '中联医保
                SettleDelSwap = 住院结算冲销_中联(lng结帐ID, intinsure)
        End Select
    End If
End Function

Public Function ComeInSwap(lngPatiID As Long, lngPageID As Long, ByRef str医保号 As String, _
    Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1131-办理登记,1132-撤消出院
    Dim strNO As String
    Dim bln部分冲销单据 As Boolean
    Dim bln部分冲销明细 As Boolean
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    bln部分冲销单据 = GetCapability(support允许部份冲销单据, lngPatiID, intinsure)
    bln部分冲销明细 = GetCapability(support允许部分冲销明细, lngPatiID, intinsure)
    
    '检查如果存在部分冲销单据的情况则不允许办理入院
    If bln部分冲销单据 = False Then
        '完全未冲销的单据不统计，提出来的就是存在部分冲销单据的数据
        gstrSQL = " " & _
                  " SELECT distinct A.NO,A.记录性质 " & _
                  " FROM 住院费用记录 A, " & _
                  "     (SELECT NO,记录性质,COUNT(*) " & _
                  "     FROM 住院费用记录 " & _
                  "     WHERE 病人ID=[1] AND 主页ID=[2] AND 记录状态 IN (1,3) " & _
                  "     GROUP BY NO,记录性质 " & _
                  "     MINUS  " & _
                  "     SELECT NO,记录性质,COUNT(*) " & _
                  "     FROM 住院费用记录 " & _
                  "     WHERE 病人ID=[1] AND 主页ID=[2] AND 记录状态 =2 " & _
                  "     GROUP BY NO,记录性质) B " & _
                  " WHERE A.NO=B.NO AND A.记录性质=B.记录性质 AND A.记录状态=2"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在部分冲销单据的数据", lngPatiID, lngPageID)
        If rsTmp.RecordCount <> 0 Then
            Do While Not rsTmp.EOF
                strNO = strNO & "," & rsTmp!NO
                rsTmp.MoveNext
            Loop
            strNO = Mid(strNO, 2)
            MsgBox "该医保不支持部分冲销单据，而该病人在医保入院之前存在类似数据，请将此类单据完全冲销后重新开单！" & vbCrLf & strNO, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '检查如果存在部分冲销明细的情况则不允许办理入院
    If bln部分冲销明细 = False Then
        gstrSQL = "SELECT distinct A.NO,A.记录性质 " & _
                  " FROM 住院费用记录 A, " & _
                  "     (SELECT NO,记录性质,序号,SUM(付数*数次) AS 数量 " & _
                  "     FROM 住院费用记录 " & _
                  "     WHERE 病人ID=[1] AND 主页ID=[2]" & _
                  "     HAVING SUM(付数*数次)<>0 " & _
                  "     GROUP BY NO,记录性质,序号) B " & _
                  " WHERE A.NO=B.NO AND A.记录性质=B.记录性质 AND A.序号=B.序号 AND A.记录状态=2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在部分冲销明细的数据", lngPatiID, lngPageID)
        If rsTmp.RecordCount <> 0 Then
            Do While Not rsTmp.EOF
                strNO = strNO & "," & rsTmp!NO
                rsTmp.MoveNext
            Loop
            strNO = Mid(strNO, 2)
            MsgBox "该医保不支持部分冲销明细，而该病人在医保入院之前存在类似数据，请将此类单据完全冲销后重新开单！" & vbCrLf & strNO, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        ComeInSwap = gobjInsure_Obj(mintOrder).ComeInSwap(lngPatiID, lngPageID, str医保号, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_四川自贡
                ComeInSwap = 入院登记_自贡(lngPatiID, lngPageID)
            Case TYPE_昭通
                ComeInSwap = 入院登记_昭通(lngPatiID, lngPageID, str医保号)
            Case TYPE_徐州农保
                ComeInSwap = 入院登记_徐州农保(lngPatiID, lngPageID, str医保号)
            Case TYPE_徐州
                ComeInSwap = 入院登记_徐州(lngPatiID, lngPageID, str医保号)
            Case TYPE_徐州六院
                ComeInSwap = 入院登记_中联(lngPatiID, intinsure)
            Case TYPE_宁海
                ComeInSwap = 入院登记_宁海(lngPatiID, lngPageID, str医保号)
            Case TYPE_余姚
                ComeInSwap = 入院登记_余姚(lngPatiID, lngPageID, str医保号)
            Case TYPE_浙江
                ComeInSwap = 入院登记_浙江(lngPatiID, lngPageID, str医保号)
            Case TYPE_吉林
                '20040923刘兴宏加入
                ComeInSwap = 入院登记_吉林(lngPatiID, lngPageID, str医保号)
            Case TYPE_北京尚洋
                ComeInSwap = 入院登记_北京尚洋(lngPatiID, lngPageID, str医保号)
            Case TYPE_华东
                ComeInSwap = 入院登记_华东(lngPatiID, lngPageID, str医保号)
            Case TYPE_南京市
                ComeInSwap = True
            Case TYPE_江苏
                ComeInSwap = 入院登记_江苏(lngPatiID, lngPageID, str医保号)
            Case TYPE_涪陵
                ComeInSwap = 入院登记_涪陵(lngPatiID, lngPageID, str医保号)
            Case TYPE_广元
                ComeInSwap = 入院登记_广元(lngPatiID, lngPageID, str医保号)
            Case TYPE_重庆市
                ComeInSwap = 入院登记_重庆(lngPatiID, lngPageID, str医保号)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                ComeInSwap = 入院登记_松藻(lngPatiID)
            Case TYPE_重庆壁山
                ComeInSwap = 入院登记_壁山(lngPatiID, lngPageID, str医保号)
            Case TYPE_重庆中梁山
                ComeInSwap = 入院登记_中梁山(lngPatiID, lngPageID)
            Case TYPE_自贡市
                ComeInSwap = 入院登记_中软(lngPatiID, lngPageID, str医保号)
            Case TYPE_泸州市
                ComeInSwap = 入院登记_泸州(lngPatiID, lngPageID, str医保号)
            Case TYPE_铜仁
                ComeInSwap = 入院登记_铜仁(lngPatiID, lngPageID, str医保号)
            Case TYPE_贵阳市
                ComeInSwap = 入院登记_贵阳(lngPatiID, lngPageID, str医保号)
            Case TYPE_云南省, TYPE_昆明市
                ComeInSwap = 入院登记_云南(lngPatiID, lngPageID, str医保号, intinsure)
            Case TYPE_云南建水
                ComeInSwap = 入院登记_云南建水(lngPatiID, lngPageID, str医保号)
            Case TYPE_成都市
                ComeInSwap = 入院登记_成都(lngPatiID, lngPageID, str医保号)
            Case TYPE_成都莲合
                ComeInSwap = 入院登记_莲合(lngPatiID, lngPageID)
            Case TYPE_开县
                '20050124
                ComeInSwap = 入院登记_开县(lngPatiID, lngPageID)
                
            Case type_成都郊县
                ComeInSwap = 入院登记_成都郊县(lngPatiID, lngPageID, str医保号)
            Case TYPE_新都
                ComeInSwap = 入院登记_新都(lngPatiID, lngPageID, str医保号)
            Case TYPE_成都南充
                ComeInSwap = 入院登记_成都南充(lngPatiID)
            Case TYPE_咸阳市
                ComeInSwap = 入院登记_咸阳(lngPatiID, lngPageID, str医保号)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                ComeInSwap = 入院登记_福建巨龙(lngPatiID, intinsure)
            Case type_米易
                ComeInSwap = 入院登记_米易(lngPatiID, lngPageID, str医保号)
            Case TYPE_四川眉山
                ComeInSwap = 入院登记_眉山(lngPatiID, lngPageID, str医保号)
            Case TYPE_沈阳市
                ComeInSwap = 入院登记_沈阳市(lngPatiID, lngPageID, str医保号)
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                ComeInSwap = 入院登记_大连(lngPatiID, lngPageID, str医保号, intinsure)
            Case TYPE_乐山
                ComeInSwap = 入院登记_乐山(lngPatiID, lngPageID, str医保号)
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                ComeInSwap = 入院登记_重大校园卡(lngPatiID, lngPageID, str医保号)
            Case TYPE_重庆银海版
                ComeInSwap = 入院登记_重庆银海版(lngPatiID, lngPageID, str医保号)
            Case TYPE_重庆渝北
                  '刘兴宏(20040705)
                ComeInSwap = 入院登记_重庆渝北(lngPatiID, lngPageID, str医保号)
            Case TYPE_北京
                ComeInSwap = 入院登记_北京(lngPatiID, lngPageID)
            Case TYPE_兴安
                '刘兴宏:20040827
                ComeInSwap = 入院登记_兴安(lngPatiID, lngPageID, str医保号)
            Case TYPE_毕节
                ComeInSwap = 入院登记_毕节(lngPatiID, lngPageID)
            Case TYPE_黔南
                '刘兴宏:20041022
                ComeInSwap = 入院登记_黔南(lngPatiID, lngPageID, str医保号)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                ComeInSwap = 入院登记_奉庆(lngPatiID, lngPageID, str医保号)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                ComeInSwap = 入院登记_成都德阳(lngPatiID, lngPageID, str医保号)
            Case TYPE_成都内江
                '刘兴宏:20041207
                ComeInSwap = 入院登记_成都内江(lngPatiID, lngPageID, str医保号)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                ComeInSwap = 入院登记_广元旺苍(lngPatiID, lngPageID, str医保号)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                ComeInSwap = 入院登记_南充阆中(lngPatiID, lngPageID, str医保号)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                ComeInSwap = 入院登记_兴成(lngPatiID, lngPageID, str医保号)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                ComeInSwap = 入院登记_神木大兴(lngPatiID, lngPageID, str医保号)
                
            Case TYPE_渝北农医
                ComeInSwap = 入院登记_渝北农医(lngPatiID, lngPageID, str医保号)
            Case TYPE_慈溪农医
                ComeInSwap = 入院登记_慈溪农医(lngPatiID, lngPageID, str医保号)
            Case TYPE_山西
                ComeInSwap = 入院登记_山西(lngPatiID, lngPageID, str医保号)
            Case TYPE_冕宁
                ComeInSwap = 入院登记_冕宁(lngPatiID, lngPageID, str医保号)
    '20050401  陈东
            Case type_铜梁合医
                ComeInSwap = 入院登记_铜梁合医(lngPatiID, lngPageID, type_铜梁合医)
                '20050530 陈东
            Case TYPE_铜山县
        'Modified By 陈东 2005-06-24 16:32:07
                ComeInSwap = 入院登记_铜山县(lngPatiID, lngPageID, intinsure)
            Case Is > 900
                '中联医保
                ComeInSwap = 入院登记_中联(lngPatiID, intinsure)
        End Select
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function LeaveSwap(lngPatiID As Long, lngPageID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1131-取消入院,1132-病人出院
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select A.出院日期,A.出院病床,Decode(A.出院方式,'正常',0,'死亡',1,'转院',2,9) as 出院方式,B.名称,D.住院号,Sysdate as 经办时间," & _
            " C.卡号,C.医保号,C.密码,C.顺序号 " & _
            " From 病案主页 A,部门表 B,医保病人档案 C,医保病人关联表 E,病人信息 D " & _
            " Where A.病人ID=D.病人ID And A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.入院科室ID=B.ID And A.病人ID=E.病人ID and C.险类=E.险类 and C.医保号=E.医保号 And C.险类=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", lngPatiID, lngPageID, intinsure)
    If rsTmp.EOF Then
        MsgBox "没有此病人或此病人不是医保病人！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        LeaveSwap = gobjInsure_Obj(mintOrder).LeaveSwap(lngPatiID, lngPageID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_四川自贡
                LeaveSwap = 出院登记_自贡(lngPatiID, lngPageID)
            Case TYPE_昭通
                LeaveSwap = 出院登记_昭通(lngPatiID, lngPageID)
            Case TYPE_徐州农保
                LeaveSwap = 出院登记_徐州农保(lngPatiID, lngPageID)
            Case TYPE_徐州
                LeaveSwap = 出院登记_徐州(lngPatiID, lngPageID)
            Case TYPE_徐州六院
                LeaveSwap = 出院登记_中联(lngPatiID, lngPageID, intinsure)
            Case TYPE_宁海
                LeaveSwap = 出院登记_宁海(lngPatiID, lngPageID)
            Case TYPE_余姚
                LeaveSwap = 出院登记_余姚(lngPatiID, lngPageID)
            Case TYPE_浙江
                LeaveSwap = 出院登记_浙江(lngPatiID, lngPageID)
            Case TYPE_吉林
                '20040923刘兴宏加入
                LeaveSwap = 出院登记_吉林(lngPatiID, lngPageID)
            Case TYPE_北京尚洋
                LeaveSwap = 出院登记_北京尚洋(lngPatiID, lngPageID)
            Case TYPE_华东
                LeaveSwap = 出院登记_华东(lngPatiID, lngPageID)
            Case TYPE_南京市
                LeaveSwap = True
            Case TYPE_江苏
                LeaveSwap = True
                'LeaveSwap = 出院登记_江苏(lngPatiID, lngPageID)
            Case TYPE_涪陵
                LeaveSwap = 出院登记_涪陵(lngPatiID, lngPageID)
            Case TYPE_广元
                LeaveSwap = 出院登记_广元(lngPatiID, lngPageID)
            Case TYPE_重庆市
                LeaveSwap = 出院登记_重庆(lngPatiID, lngPageID)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                LeaveSwap = 出院登记_松藻(lngPatiID, lngPageID)
            Case TYPE_重庆壁山
                LeaveSwap = 出院登记_壁山(lngPatiID, lngPageID)
            Case TYPE_重庆中梁山
                LeaveSwap = 出院登记_中梁山(lngPatiID, lngPageID)
            Case TYPE_自贡市
                LeaveSwap = 出院登记_中软(lngPatiID)
            Case TYPE_泸州市
                LeaveSwap = 出院登记_泸州(lngPatiID)
            Case TYPE_铜仁
                LeaveSwap = 出院登记_铜仁(lngPatiID)
            Case TYPE_贵阳市
                LeaveSwap = 出院登记_贵阳(lngPatiID, lngPageID)
            Case TYPE_云南省, TYPE_昆明市
                If IsNull(rsTmp!顺序号) Then
                    MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
                    Exit Function
                End If
                LeaveSwap = 出院登记_云南(lngPatiID, lngPageID, rsTmp!顺序号, intinsure)
            Case TYPE_云南建水
                If IsNull(rsTmp!顺序号) Then
                    MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
                    Exit Function
                End If
                LeaveSwap = 出院登记_云南建水(lngPatiID, lngPageID, rsTmp!顺序号)
            Case TYPE_成都市
                LeaveSwap = 出院登记_成都(lngPatiID, lngPageID, rsTmp)
            Case TYPE_成都莲合
                LeaveSwap = 出院登记_莲合(lngPatiID, lngPageID)
            Case TYPE_开县
                '20050124
                LeaveSwap = 出院登记_开县(lngPatiID, lngPageID)
                
            Case type_成都郊县
                LeaveSwap = 出院登记_成都郊县(lngPatiID, lngPageID)
            Case TYPE_新都
                LeaveSwap = 出院登记_新都(lngPatiID, lngPageID)
            Case TYPE_成都南充
                LeaveSwap = 出院登记_成都南充(lngPatiID)
            Case TYPE_咸阳市
                LeaveSwap = 出院登记_咸阳(lngPatiID)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                LeaveSwap = 出院登记_福建巨龙(intinsure, lngPatiID, lngPageID)
            Case type_米易
                LeaveSwap = 出院登记_米易(lngPatiID, lngPageID)
            Case TYPE_四川眉山
                LeaveSwap = 出院登记_眉山(lngPatiID, lngPageID)
            Case TYPE_沈阳市
                LeaveSwap = 出院登记_沈阳市(lngPatiID, lngPageID)
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                LeaveSwap = 出院登记_大连(lngPatiID, lngPageID)
            Case TYPE_乐山
                LeaveSwap = 出院登记_乐山(lngPatiID, lngPageID)
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                LeaveSwap = 出院登记_重大校园卡(lngPatiID, lngPageID)
            Case TYPE_重庆银海版
                LeaveSwap = 出院登记_重庆银海版(lngPatiID, lngPageID)
            Case TYPE_重庆渝北
                  '刘兴宏(20040705)
                LeaveSwap = 出院登记_重庆渝北(lngPatiID, lngPageID)
            Case TYPE_北京
                LeaveSwap = 出院登记_北京(lngPatiID, lngPageID)
            Case TYPE_兴安
                '刘兴宏:20040827
                LeaveSwap = 出院登记_兴安(lngPatiID, lngPageID)
            Case TYPE_毕节
                LeaveSwap = 出院登记_毕节(lngPatiID, lngPageID)
            Case TYPE_黔南
                '刘兴宏:20041022
                LeaveSwap = 出院登记_黔南(lngPatiID, lngPageID)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                LeaveSwap = 出院登记_奉庆(lngPatiID, lngPageID)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                LeaveSwap = 出院登记_成都德阳(lngPatiID, lngPageID)
            Case TYPE_成都内江
                '刘兴宏:20041207
                LeaveSwap = 出院登记_成都内江(lngPatiID, lngPageID)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                LeaveSwap = 出院登记_广元旺苍(lngPatiID, lngPageID)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                LeaveSwap = 出院登记_南充阆中(lngPatiID, lngPageID)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                LeaveSwap = 出院登记_兴成(lngPatiID, lngPageID)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                LeaveSwap = 出院登记_神木大兴(lngPatiID, lngPageID)
                
            Case TYPE_渝北农医
                LeaveSwap = 出院登记_渝北农医(lngPatiID, lngPageID)
            Case TYPE_慈溪农医
                LeaveSwap = 出院登记_慈溪农医(lngPatiID, lngPageID)
            Case TYPE_山西
                LeaveSwap = 出院登记_山西(lngPatiID, lngPageID)
            Case TYPE_冕宁
            '20050401 陈东
                LeaveSwap = 出院登记_冕宁(lngPatiID, lngPageID)
            Case type_铜梁合医
            '20050530 陈东
                LeaveSwap = 出院登记_铜梁合医(lngPatiID, lngPageID, type_铜梁合医)
            Case TYPE_铜山县
                LeaveSwap = 出院登记_铜山县(lngPatiID, lngPageID, intinsure)
            Case Is > 900
                '中联医保
                LeaveSwap = 出院登记_中联(lngPatiID, lngPageID, intinsure)
        End Select
    End If
End Function

Public Function GetCapability(ByVal cap参数 As Integer, Optional ByVal lng病人ID As Long = 0, _
    Optional ByVal intinsure As Integer, Optional ByRef strAdvance As String = "") As Boolean
'功能：判断后来补充的一些业务在不同的医保软件是否得到支持
'参数：cap业务     医院的业务代号
'      intInsure     指定险类,缺省为当前险类(如果已连接)
'返回：True        支持该业务
'      False       不支持该业务
'说明：不需要进行医保初始化就可以完成的查询


'------------------------------------------
'修改者：朱玉宝
'修改时间：2005-08-08
'修改内容：增加两个保险参数：support门诊结算作废、support住院结算作废
'功能说明：如果检查的是这两个参数，当strAdvance为空，则表示检查该医保是否支持结算作废；如果不为空，表示检查特定的结算方式是否支持全退，如果不支持全退，费用将其归入现金退给病人
'例：
'Case support门诊结算作废
'    If strAdvance = "" Then
'        '检查是否支持门诊结算作废
'        医保参数设置_Template = True
'    Else
'        '判断结算作废时，某种结算方式是否支持全退
'        Select Case strAdvance
'        Case "个人帐户"
'            医保参数设置_Template = True
'        Case "医保基金"
'            医保参数设置_Template = True
'        Case "公务员补助"
'            医保参数设置_Template = True
'        End Select
'    End If

''------------------------------------------
''修改者：朱玉宝
''修改时间：2006-10-08
''修改内容：增加保险参数：support结算_母婴费用
''功能说明：当strAdvance为空，则表示判断是否结算所有费用；当strAdvance="婴儿费"，判断是否仅结算婴儿的费用；当strAdvance="母亲费"，判断是否结算除婴儿费以外的所有费用
''例：
''Case support结算_母婴费用
''    If strAdvance = "" Then
''        医保参数设置_Template = True
''    Else
''        Select Case strAdvance
''        Case "婴儿费"
''            医保参数设置_Template = True
''        Case "母亲费"
''            医保参数设置_Template = True
''        End Select
''    End If
'

    Dim rsTmp As New ADODB.Recordset
    Dim cap业务 As 医院业务
    
    '由于本函数只对新增业务进行判断，缺省医保是不支持这些新业务
    GetCapability = False
        
    cap业务 = cap参数
    
    '以下参数缺省为支持，不支持的接口需单独处理
    If cap业务 = support负数记帐 Then GetCapability = True
'    If cap业务 = support结算_母婴费用 Then GetCapability = True
    '----------------------------------------------------------
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        Call gobjInsure_Obj(mintOrder).InitOracle(gcnOracle)
        GetCapability = gobjInsure_Obj(mintOrder).GetCapability(cap参数, lng病人ID, intinsure, strAdvance)
        Exit Function
    End If
    
    Select Case intinsure
        'support门诊预算在下面有单独的处理
        Case TYPE_奉化市农医
            Select Case cap业务
                Case support未结清出院, _
                     support允许不设置医保项目, _
                     support撤销出院, _
                     support必须录入入出诊断, _
                     support结算使用个人帐户, _
                     support门诊预算, _
                     support门诊退费, _
                     support门诊必须传递明细
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_昭通
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support出院结算必须出院, _
                     support未结清出院, _
                     support门诊退费, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤销出院, _
                     support分币处理 ', _
                     support允许不设置医保项目
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_慈溪农医
            Select Case cap业务
                Case support门诊必须传递明细, support结算使用个人帐户, support撤销出院, support出院病人结算作废, _
                    support门诊预算, support门诊退费, support未结清出院, support记帐完成后上传, support分币处理, _
                    support医嘱上传, support记帐上传, support记帐作废上传, support结帐退个人帐户, support允许部份冲销单据
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_徐州农保
            Select Case cap业务
                Case support出院结算必须出院, _
                     support未结清出院, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤销出院, _
                     support分币处理, _
                     support门诊预算, _
                     support门诊必须传递明细, _
                     support门诊退费, _
                     support结算使用个人帐户
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_徐州
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support必须录入入出诊断, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support出院结算必须出院, _
                     support未结清出院, _
                     support门诊退费, _
                     support门诊预算, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤销出院, _
                     support分币处理, _
                     support允许不设置医保项目
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_渝北农医  ', support允许不设置医保项目
            Select Case cap业务
                Case support门诊必须传递明细, support结算使用个人帐户, support撤销出院, support出院病人结算作废, _
                    support门诊预算, support门诊退费, support未结清出院, _
                    support医嘱上传, support记帐上传, support记帐作废上传, support结帐退个人帐户, support允许部份冲销单据
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_宁海
            Select Case cap业务
                Case support门诊必须传递明细, support结算使用个人帐户, _
                    support撤销出院, support允许不设置医保项目, support门诊退费, _
                    support门诊预算, support允许冲销已结帐的记帐单据, support撤销出院, support未结清出院
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_余姚
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support记帐完成后上传, _
                     support未结清出院, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊预算, _
                     support门诊退费, _
                     support结帐退个人帐户, _
                     support撤销出院, _
                     support允许不设置医保项目
'                     support出院病人结算作废, _
'                     support记帐上传, _
'                     support记帐作废上传, _
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_浙江
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support记帐完成后上传, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support未结清出院, _
                     support门诊退费, _
                     support门诊预算, _
                     support撤销出院, _
                     support允许不设置医保项目
'                     support结帐退个人帐户, _
'                     support收费帐户全自费, _
'                     support收费帐户首先自付, _
'                     support结算使用个人帐户, _
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_毕节
            Select Case cap业务
                Case support门诊必须传递明细, support结算使用个人帐户, support记帐上传, _
                    support记帐作废上传, support医嘱上传, support记帐完成后上传, support出院病人结算作废, _
                    support撤销出院, support允许不设置医保项目, support门诊退费, _
                    support门诊预算, support未结清出院, support允许冲销已结帐的记帐单据, support允许部份冲销单据
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_吉林
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤销出院
            '                     support结算使用个人帐户, _
                     '刘兴宏:20040923取消此参数:support允许不设置医保项目
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_北京尚洋
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support记帐完成后上传, _
                     support未结清出院, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤销出院, _
                     support允许不设置医保项目
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                     GetCapability = True
            End Select
        Case TYPE_华东
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support门诊预算, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤销出院
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_南京市
            Select Case cap业务
                Case support门诊预算, _
                     support结算使用个人帐户, _
                     support门诊必须传递明细, _
                     support出院结算必须出院, _
                     support未结清出院, _
                     support门诊退费, _
                     support撤销出院, _
                     support结帐_设置婴儿费条件
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废, support出院病人结算作废
                    GetCapability = True
                Case support允许部分冲销明细, support允许部份冲销单据
                    GetCapability = True
            End Select
        Case TYPE_江苏
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support必须录入入出诊断, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support记帐完成后上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support门诊预算, _
                     support撤销出院, _
                     support出院病人结算作废, _
                     support未结清出院
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_重庆市
            Select Case cap业务
                Case support门诊退费, _
                     support结算使用个人帐户, _
                     support结帐退个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support出院结算必须出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support记帐完成后上传, _
                     support出院病人结算作废, _
                     support必须录入入出诊断, _
                     support允许部份冲销单据, _
                     support挂号使用个人帐户, _
                     support撤销出院, _
                     support分币处理, _
                     support结帐_指定日期范围, _
                     support多单据收费, _
                     Support多单据收费必须全退, _
                     Support初始化失败脱机处理
                    '支持的业务
                    GetCapability = True
                Case support中途结算仅处理已上传部分
                    Dim bln中途结算只处理已上传费用 As Boolean
                     '判断中途结算是否只结算已上传费用
                    bln中途结算只处理已上传费用 = True
                    gstrSQL = " Select Nvl(参数值,1) AS 参数值 From 保险参数 Where 险类=[1] And 参数名=[2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "判断中途结算是否只结算已上传费用", TYPE_重庆市, "中途结算")
                    If rsTmp.RecordCount <> 0 Then
                        bln中途结算只处理已上传费用 = (rsTmp!参数值 = 1)
                    End If
                    GetCapability = bln中途结算只处理已上传费用
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊必须传递明细
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_重庆松藻
            Select Case cap业务
                Case support门诊退费, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support未结清出院, support门诊预算
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_涪陵
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support必须录入入出诊断, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support记帐完成后上传, _
                     support未结清出院, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support允许不设置医保项目, _
                     support撤销出院
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_广元
            '不能支持结算作废，因为医保端结算作废后，必须重新办入院，这样就没有意义了
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support必须录入入出诊断, _
                     support记帐完成后上传, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support未结清出院, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support允许不设置医保项目, _
                     support门诊预算, _
                     support出院病人结算作废
                     '支持的业务
                     GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_重庆壁山                 'add by zyc 2003-4-9,add by ccy 2003-11-28
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support必须录入入出诊断, _
                     support记帐完成后上传, _
                     support出院结算必须出院, _
                     support未结清出院, _
                     support门诊部分退现金, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support结算使用个人帐户, _
                     support记帐完成后上传, _
                     support撤销出院, _
                     support住院结算作废
                     '支持的业务
                     GetCapability = True
            End Select
            If Val(Get保险参数_壁山("适用地区")) = 2 Then
                If cap业务 = support门诊部分退现金 Then GetCapability = False
                If cap业务 = support门诊退费 Then GetCapability = True
                If cap业务 = support门诊结算作废 Then GetCapability = True
                If cap业务 = support结帐退个人帐户 Then GetCapability = True
                If cap业务 = support记帐上传 Then GetCapability = False                                 '新增部分
                If cap业务 = support记帐作废上传 Then GetCapability = False                             '新增部分
'                If cap业务 = support必须录入入出诊断 Then GetCapability = False
            End If
            If cap业务 = support门诊预算 Then
                If Val(Get保险参数_壁山("适用地区")) = 3 Or Val(Get保险参数_壁山("适用地区")) = 1 Then
                    GetCapability = True
                Else
                    GetCapability = gbln特殊门诊
                End If
            End If
        Case TYPE_重庆中梁山
            Select Case cap业务
                Case support撤销出院, support未结清出院
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = True
                Case support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_重庆银海版
            Select Case cap业务
                Case support撤销出院, support结算使用个人帐户, support门诊必须传递明细, _
                    support结帐退个人帐户, support门诊退费, support未结清出院, support门诊预算, support允许冲销已结帐的记帐单据
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_自贡市
            '自贡医保改成固定参数，参数设置时已经无效。
            Select Case cap业务
                Case support门诊退费, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support未结清出院, _
                     support必须录入入出诊断, _
                     support结算帐户首先自付, _
                     support收费帐户首先自付              '收费帐户全自费、结算帐户全自费、结算帐户超限全不支持
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_四川自贡
            Select Case cap业务
                Case support门诊退费, _
                     support门诊预算, _
                     support结算使用个人帐户, _
                     support门诊部分退现金, _
                     support未结清出院, _
                     support撤销出院, _
                     support必须录入入出诊断, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传
                    GetCapability = True
                Case support住院结算作废
                    GetCapability = True
                Case support门诊结算作废
                    If strAdvance = "" Then
                        GetCapability = True
                    Else
                        If strAdvance = "个人帐户" Then
                            GetCapability = False
                        Else
                            GetCapability = True
                        End If
                    End If
            End Select
        Case TYPE_泸州市, TYPE_铜仁
            '自贡医保改成固定参数，参数设置时已经无效。
            Select Case cap业务
                Case support门诊退费, _
                     support未结清出院, _
                     support必须录入入出诊断, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support结算帐户首先自付, _
                     support收费帐户首先自付              '收费帐户全自费、结算帐户全自费、结算帐户超限全不支持
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_贵阳市
            Select Case cap业务
                Case support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support门诊退费, _
                     support门诊结帐, _
                     support未结清出院, _
                     support撤销出院, _
                     support允许部份冲销单据, _
                     support多单据收费
                    '支持的业务
                    GetCapability = True
                Case support门诊必须传递明细
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
                Case support记帐上传, support记帐作废上传 '程池富增加 2011-05-23
                    GetCapability = (gint进行医保交易 = 1)
            End Select
        Case TYPE_云南省, TYPE_昆明市
            Select Case cap业务
                Case support门诊退费, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support门诊必须传递明细, _
                     support必须录入入出诊断, _
                     support未结清出院, _
                     support结算使用个人帐户
                    '支持的业务
                    GetCapability = True
                Case support撤销出院
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_云南建水
            '由于不支持门诊退费，所以单列
            Select Case cap业务
                Case support未结清出院
                    GetCapability = True
                Case support收费帐户全自费, support收费帐户首先自付, support必须录入入出诊断
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support结算使用个人帐户
                    GetCapability = True
                Case support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_成都市
            Select Case cap业务
                Case support收费帐户全自费, support收费帐户首先自付
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support记帐完成后上传
                    GetCapability = True
                Case support多单据收费
                    GetCapability = True
                Case support记帐上传, support记帐作废上传, support医嘱上传
                    '200308z012:不实时上传
                    GetCapability = False
                Case support住院结算作废, support允许部分冲销明细
                    GetCapability = True
            End Select
        Case TYPE_成都莲合
            Select Case cap业务
                Case support收费帐户全自费, support收费帐户首先自付, support允许部分冲销明细, support允许部份冲销单据
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support未结清出院
                    GetCapability = True
                Case support门诊收费存为划价单
                    GetCapability = (Val(GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("HIS收费"), 0)) = 0)
                Case support住院结算作废, support撤销出院
                    GetCapability = True
            End Select
        Case TYPE_开县
            '20050124
            Select Case cap业务
                Case support收费帐户全自费, support收费帐户首先自付, support允许部分冲销明细
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support未结清出院
                    GetCapability = True
                Case support门诊收费存为划价单
                    GetCapability = True
                Case support住院结算作废
                    GetCapability = True
            End Select
        Case type_成都郊县
            Select Case cap业务
                Case support收费帐户全自费, support收费帐户首先自付, support记帐上传, support记帐作废上传, support医嘱上传, support记帐完成后上传
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support门诊结算作废
                    GetCapability = False
                Case support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_新都
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support医嘱上传, _
                     support记帐完成后上传
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_成都南充
            Select Case cap业务
                Case support记帐上传, support记帐作废上传, support医嘱上传, _
                    support允许不设置医保项目, support未结清出院, support撤销出院
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_咸阳市
            Select Case cap业务
                Case support结算使用个人帐户, support收费帐户全自费, support收费帐户首先自付
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support住院结算作废, support门诊结算作废
                    GetCapability = True
                Case support允许部份冲销单据, support允许部分冲销明细
                    GetCapability = True
            End Select
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            Select Case cap业务
                Case support门诊退费, support结算使用个人帐户, _
                support门诊必须传递明细, support撤销出院, support未结清出院, support结帐退个人帐户
                    GetCapability = True
                'Modified by 朱玉宝 20031218 地区：福州
                Case support门诊预算
                    If intinsure = TYPE_福建巨龙 Then
                        GetCapability = True
                    Else
                        GetCapability = False
                    End If
                Case support必须录入入出诊断
                    If intinsure = TYPE_南平市 Then
                        GetCapability = True
                    Else
                        GetCapability = False
                    End If
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case type_米易
            Select Case cap业务
                Case support门诊预算, support记帐上传, support记帐作废上传, _
                    support结算使用个人帐户, support撤销出院, support未结清出院, support必须录入入出诊断
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_四川眉山
            Select Case cap业务
                Case support门诊预算, support结算使用个人帐户, support门诊退费, _
                    support撤销出院, support未结清出院, support必须录入入出诊断
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_沈阳市
            Select Case cap业务
                Case support门诊预算, support结算使用个人帐户, support门诊退费, support撤销出院, _
                    support记帐上传, support记帐作废上传, support医嘱上传, support允许部份冲销单据, _
                    support未结清出院, support结帐退个人帐户, support门诊必须传递明细       'support必须录入入出诊断
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废, support门诊改费
                    GetCapability = True
            End Select
        Case TYPE_大连市, TYPE_大连开发区
            Select Case cap业务
                Case support门诊预算, _
                     support门诊退费, _
                     support未结清出院, _
                     support结算使用个人帐户, _
                     support门诊必须传递明细, _
                     support出院病人结算作废, _
                     support必须录入入出诊断, _
                     support撤销出院, _
                     support出院无实际交易, _
                     support结帐退个人帐户, _
                     support多单据收费
                     
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_北京
            Select Case cap业务
                Case support门诊预算, _
                     support门诊退费, _
                     support未结清出院, _
                     support结算使用个人帐户, _
                     support必须录入入出诊断, _
                     support撤销出院, _
                     support结帐退个人帐户, _
                     support允许不设置医保项目
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_乐山
            Select Case cap业务
                'Modified by ZYB 2004-08-17
                Case support未结清出院, support挂号使用个人帐户, _
                    support结算使用个人帐户, support撤销出院, support门诊预算, _
                    support门诊退费, support记帐上传, support记帐完成后上传, support记帐作废上传, support医嘱上传
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_重大校园卡
            Select Case cap业务
                Case support门诊退费, _
                     support预交退个人帐户, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, support挂号使用个人帐户, support允许不设置医保项目, support门诊预算
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_重庆渝北
            Select Case cap业务
               Case support门诊退费, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support必须录入入出诊断, _
                     support门诊必须传递明细, _
                     support撤销出院, _
                     support多单据收费

                    '支持的业务
                    GetCapability = True
                Case support门诊结算作废
                    GetCapability = True
                Case support住院结算作废
                    GetCapability = False
            End Select
        Case TYPE_兴安
            '刘兴宏:20040827
            Select Case cap业务
                Case support门诊退费, support门诊预算, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support必须录入入出诊断, _
                     support门诊必须传递明细, _
                     support撤销出院, _
                     support允许不设置医保项目, _
                     support出院病人结算作废, _
                     support出院结算必须出院
                     
                    '支持的业务
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_黔南
             '刘兴宏:20041022
            Select Case cap业务
                Case support门诊退费, support门诊预算, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support门诊必须传递明细, _
                     support撤销出院, _
                     support允许不设置医保项目, _
                     support出院病人结算作废, _
                     support结帐退个人帐户
                     'support出院结算必须出院
                    '支持的业务
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_临沧奉庆
            '刘兴宏:20041231
            Select Case cap业务
                Case support门诊退费, support门诊预算, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support门诊必须传递明细, _
                     support撤销出院, _
                     support允许不设置医保项目, _
                     support出院病人结算作废
                     'support出院结算必须出院
                    '支持的业务
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_成都德阳
            '刘兴宏:20041103
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support记帐完成后上传, _
                     support撤销出院, _
                     support允许不设置医保项目, _
                     support出院病人结算作废, _
                     support出院结算必须出院
                    '支持的业务
                    GetCapability = True
                Case support门诊结算作废
                    GetCapability = False
                Case support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_成都内江
            '刘兴宏:20041207
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support门诊必须传递明细, _
                     support门诊退费, _
                     support门诊预算, _
                     support未结清出院, _
                     support结算使用个人帐户, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support记帐完成后上传, _
                     support撤销出院, _
                     support出院病人结算作废, _
                     support出院结算必须出院
                    '支持的业务
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                    support门诊必须传递明细, _
                    support允许不设置医保项目, _
                    support门诊退费, _
                    support未结清出院, _
                    support记帐上传, _
                    support记帐完成后上传, _
                    support记帐作废上传, _
                    support撤销出院, _
                    support出院病人结算作废, _
                    support出院结算必须出院, _
                    support门诊预算, _
                    support结算使用个人帐户
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_南充阆中
            '刘兴宏:20050419
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                    support门诊必须传递明细, _
                    support允许不设置医保项目, _
                    support门诊退费, _
                    support未结清出院, _
                    support记帐上传, _
                    support记帐完成后上传, _
                    support记帐作废上传, _
                    support撤销出院, _
                    support出院病人结算作废, _
                    support出院结算必须出院, _
                    support门诊预算
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废, support必须录入入出诊断
                    GetCapability = True
            End Select
            
        Case TYPE_兴成核工业
            '刘兴宏:20050226
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                    support结算使用个人帐户, _
                    support门诊必须传递明细, _
                    support门诊退费, _
                    support未结清出院, _
                    support记帐上传, _
                    support记帐完成后上传, _
                    support记帐作废上传, _
                    support撤销出院, _
                    support出院病人结算作废, _
                    support出院结算必须出院, _
                    support允许不设置医保项目, _
                    support门诊预算
                    
                    '支持的业务
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case TYPE_陕西大兴
            '刘兴宏:20050316
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                    support门诊必须传递明细, _
                    support门诊退费, _
                    support未结清出院, _
                    support结算使用个人帐户, _
                    support记帐上传, _
                    support记帐完成后上传, _
                    support记帐作废上传, _
                    support撤销出院, _
                    support出院病人结算作废, _
                    support出院结算必须出院, _
                    support门诊预算
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
         '20050311 陈东
        Case TYPE_山西
            Select Case cap业务
                Case support门诊退费, _
                     support门诊必须传递明细, _
                     support门诊预算, _
                     support结算使用个人帐户, _
                     support结帐退个人帐户, _
                     support出院结算必须出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support出院病人结算作废, _
                     support撤销出院, _
                     support未结清出院
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        '20050401 陈东
        Case TYPE_冕宁
            Select Case cap业务
                Case support门诊退费, _
                     support门诊必须传递明细, _
                     support必须录入入出诊断, _
                     support门诊预算, _
                     support结算使用个人帐户, _
                     support出院结算必须出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support未结清出院
                    GetCapability = True
                Case support门诊结算作废, support住院结算作废
                    GetCapability = True
            End Select
        Case type_铜梁合医
            '20050530 陈东
            '为保证兼容性,将此内容移到mdlTLHY中的 医保参数设置_铜梁合医中
            GetCapability = 医保参数设置_铜梁合医(cap业务)
        Case TYPE_铜山县
            GetCapability = 医保参数设置_铜山县(cap参数)
    'Modified By 陈东 2005-06-24 16:36:00
        Case TYPE_徐州六院
            Select Case cap业务
                Case support门诊退费, _
                     support预交退个人帐户, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support必须录入入出诊断, _
                     support未结清出院
                    
                    GetCapability = True
                Case support撤销出院
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊结算作废, support住院结算作废    '支持门诊、住院退费，且各种结算方式也支持全退
                    GetCapability = True
            End Select
        Case Is > 900
            Select Case cap业务
                Case support门诊退费, _
                     support预交退个人帐户, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support必须录入入出诊断, _
                     support未结清出院, _
                     support允许部分冲销明细, _
                     support允许部份冲销单据, _
                     support结帐_指定日期范围
                    
                    GetCapability = True
                Case support撤销出院
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊结算作废, support住院结算作废    '支持门诊、住院退费，且各种结算方式也支持全退
                    GetCapability = True
            End Select
    End Select
    
    If cap业务 = support分币处理 Then
        Select Case intinsure
        Case TYPE_浙江, TYPE_余姚, TYPE_徐州六院, TYPE_慈溪农医
            GetCapability = True
        Case TYPE_贵阳市, TYPE_南京市, TYPE_云南建水, TYPE_云南省, TYPE_昆明市
            GetCapability = True
        End Select
    End If
    
    If cap业务 = support挂号使用个人帐户 Then
        Select Case intinsure
            Case TYPE_成都市
                GetCapability = True
            Case TYPE_重庆壁山
                GetCapability = True      '2003-4-9modify zyc
        End Select
    End If
    
    If cap业务 = support门诊预算 Then
        Select Case intinsure
            Case TYPE_重庆市
                '重庆市所有的门诊病人都使用预算
                GetCapability = True
            Case type_成都郊县, TYPE_新都, TYPE_徐州六院
                GetCapability = True
            Case TYPE_贵阳市
                '贵阳的所有的门诊病人都使用预算
                GetCapability = True
            Case TYPE_云南省, TYPE_昆明市
'                gstrSQL = "select nvl(病种ID,0) 病种ID from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & intInsure
'                Call OpenRecordset(rsTmp, "医保接口")
'                If rsTmp.EOF = False Then
'                    '有特殊病的病人需要预算
'                    GetCapability = IIf(rsTmp("病种ID") > 0, True, False)
'                End If
'                '公务员预结算 gzh
'                If GetCapability = False Then GetCapability = mbln公务员
                GetCapability = True
            Case TYPE_云南建水
                GetCapability = True
            Case TYPE_自贡市
                gstrSQL = "select 在职 from 保险帐户 where 病人ID=[1] and 险类=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取在职", lng病人ID, intinsure)
                If rsTmp.EOF = False Then
                    GetCapability = (rsTmp("在职") = 3 Or rsTmp("在职") = 4)
                End If
            Case TYPE_泸州市, TYPE_铜仁
                '所有病人，因为有可能有慢病或公务员支持预结算
                GetCapability = True
            Case Is > 900
                gstrSQL = "select 参数值  from 保险参数 where 险类=[1] and 中心 is null and 参数名=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取保险参数", intinsure, "收费使用医保基金")
                If rsTmp.EOF = False Then
                    GetCapability = IIf(rsTmp("参数值") = 1, True, False)
                End If
        End Select
    End If
    
    If intinsure = TYPE_贵阳市 Then
        '针对贵阳医保的特殊处理 XQ 2003-05-12
        If cap业务 = support门诊连续收费 Then
            gstrSQL = "select 参数值 值 from 保险参数 where 险类=[1] and Nvl(中心,0)=0 and 参数名=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", intinsure, "门诊连续收费")
            If rsTmp.EOF = False Then
                GetCapability = IIf(rsTmp("值") = 1, True, False)
            End If
            Exit Function
        End If
        If cap业务 = support门诊收费完成后验证 Then
            GetCapability = True
            Exit Function
        End If
    End If
        
    If intinsure = TYPE_重庆松藻 Then
        gstrSQL = ""
        If cap业务 = support允许不设置医保项目 Then
            gstrSQL = "select 参数值 值 from 保险参数 where 险类=[1] and Nvl(中心,0)=0 and 参数名=[2]"
        End If
        If gstrSQL <> "" Then
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", intinsure, "允许不设置医保项目")
            If rsTmp.EOF = False Then
                GetCapability = IIf(rsTmp("值") = 1, True, False)
            End If
        End If
    End If
    
    If intinsure > 900 Or intinsure = TYPE_徐州六院 Then
        Dim STRNAME As String
        gstrSQL = ""
        '读保险参数表
        If cap业务 = support收费帐户全自费 Then
            STRNAME = "收费个人帐户使用范围"
            gstrSQL = "select substr(参数值,1,1) 值 from 保险参数 where 险类=[1] and 中心 is null and 参数名=[2]"
        ElseIf cap业务 = support收费帐户首先自付 Then
            STRNAME = "收费个人帐户使用范围"
            gstrSQL = "select substr(参数值,2,1) 值 from 保险参数 where 险类=[1] and 中心 is null and 参数名=[2]"
        ElseIf cap业务 = support结算帐户全自费 Then
            STRNAME = "结算个人帐户使用范围"
            gstrSQL = "select substr(参数值,1,1) 值 from 保险参数 where 险类=[1] and 中心 is null and 参数名=[2]"
        ElseIf cap业务 = support结算帐户首先自付 Then
            STRNAME = "结算个人帐户使用范围"
            gstrSQL = "select substr(参数值,2,1) 值 from 保险参数 where 险类=[1] and 中心 is null and 参数名=[2]"
        ElseIf cap业务 = support结算帐户超限 Then
            STRNAME = "结算个人帐户使用范围"
            gstrSQL = "select substr(参数值,3,1) 值 from 保险参数 where 险类=[1] and 中心 is null and 参数名=[2]"
        End If
        
        If gstrSQL <> "" Then
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取保险参数", intinsure, STRNAME)
            If rsTmp.EOF = False Then
                GetCapability = IIf(rsTmp("值") = 1, True, False)
            End If
        End If
    End If
End Function

Public Function GetItemInsure(病人ID As Long, 收费细目ID As Long, 金额 As Currency, _
    ByVal bln门诊 As Boolean, Optional ByVal intinsure As Integer, Optional ByRef strAdvance As String = "") As String
'功能：获取收费项目对应的医保信息
'参数：金额=收入项目级的实收金额。
'      intInsure=指定险类,缺省为当前险类(如果已连接)
'      strAdvance=医嘱的摘要，现修改为：摘要||开单数量
'返回："保险项目否(0/1);保险大类ID;进入统筹金额;保险项目编码;摘要;费用类型"
'调用模块：
    Dim rsTmp As New ADODB.Recordset ', bln全额统筹 As Boolean
    Dim rs特准 As New ADODB.Recordset
    Dim bln保险项目 As Boolean, lng大类ID As Long, cur统筹金额 As Currency
    Dim lng服务对象 As Long, str项目编码 As String
    
    Static lng病人ID As Long
    Static lng收费细目ID As Long
    Static str项目编码_TEMP As String
    bln保险项目 = False
    lng大类ID = 0
    cur统筹金额 = 0
    
    'bln全额统筹 = Is全额统筹(病人ID)
    
    'Modified by ZYB 2002-10-30
    '如果参数“允许不设置医保项目”为真，则返回实收金额（意思是全部进入统筹）
    If GetCapability(support允许不设置医保项目, 病人ID, intinsure) Then
        GetItemInsure = "1;0;" & 金额 & ";"
        If IsApartComponents(intinsure) Then
            If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
            If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
            GetItemInsure = gobjInsure_Obj(mintOrder).GetItemInsure(病人ID, 收费细目ID, 金额, bln门诊, intinsure, strAdvance)
        End If
        Exit Function
    End If
    If intinsure <> TYPE_自贡市 Then
        If intinsure <> TYPE_四川眉山 Then
            gstrSQL = " Select A.大类ID,A.项目编码,A.项目名称,A.是否医保 as 项目是否医保,B.*" & _
                    " From 保险支付项目 A,(Select * From 保险支付大类 Where 险类=[1]) B" & _
                    " Where A.大类ID=B.ID(+) And A.险类=[1] And A.收费细目ID=[2]"
        Else
            gstrSQL = "Select A.大类ID,A.项目编码,A.项目名称,A.是否医保 as 项目是否医保,B.*,C.服务对象 项目服务对象 " & _
                     " From 保险支付项目 A,(Select * From 保险支付大类 Where 险类=[1]) B, " & _
                     " (Select * From 保险项目 Where 险类=[1]) C " & _
                     " Where A.大类ID=B.ID(+) And A.项目编码=C.编码(+) And A.险类=[1] And A.收费细目ID=[2]"
        End If
    Else
        '2005-08-31 以附注中保存第二编码，原项目编码仅用于特殊病人住院；第二编码用于门诊与普通住院 ZYB
        '如果病种中有特准项目，以另外计算
        gstrSQL = "SELECT C.收费细目ID " & _
                 " FROM 保险帐户 A,保险特准项目 C,(Select * From ZLYB.病种信息 Where 险类=[1] And 病人ID=[2]) B " & _
                 " WHERE A.病人ID=[2] AND A.险类=[1] AND C.收费细目ID=[3] And C.病种ID=B.病种ID"
        Set rs特准 = zlDatabase.OpenSQLRecord(gstrSQL, "提取特准项目", intinsure, 病人ID, 收费细目ID)
        If bln门诊 Or rs特准.RecordCount = 0 Then
            '按照普通病人，正常计算
            gstrSQL = "SELECT B.ID 大类ID,A.附注 as 项目编码,C.名称,A.是否医保 AS 项目是否医保,B.* " & _
                " FROM 保险支付项目 A,(SELECT * FROM 保险支付大类 WHERE 险类 =[1]) B,保险项目 C " & _
                " WHERE A.险类=[1] AND A.收费细目ID=[2] AND A.附注=C.编码 AND A.险类=C.险类 AND C.大类编码=B.编码(+)"
        Else
            '使用原有大类编码
            gstrSQL = "SELECT B.ID 大类ID,A.项目编码,A.项目名称,A.是否医保 AS 项目是否医保,B.* " & _
                " FROM 保险支付项目 A,(SELECT * FROM 保险支付大类 WHERE 险类 =[1]) B,保险项目 C " & _
                " WHERE A.险类=[1] AND A.收费细目ID=[2] AND A.项目编码=C.编码 AND A.险类=C.险类 AND C.大类编码=B.编码(+)"
            
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", intinsure, 收费细目ID)
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If Not rsTmp.EOF Then
        bln保险项目 = True
        lng大类ID = IIf(IsNull(rsTmp!大类id), 0, rsTmp!大类id)
        str项目编码 = IIf(IsNull(rsTmp!项目编码), "", rsTmp!项目编码)
        If lng大类ID <> 0 Then
            'Modified by 朱玉宝 20031218 地区：福州
'            If intInsure <> TYPE_福建巨龙 Then
            If Not (intinsure = TYPE_福建巨龙 Or intinsure = TYPE_福建省 Or intinsure = TYPE_福州市 Or intinsure = TYPE_南平市) Then
                bln保险项目 = bln保险项目 And (IIf(IsNull(rsTmp!是否医保), 0, rsTmp!是否医保) = 1)
            End If
            If IIf(IsNull(rsTmp!算法), 0, rsTmp!算法) = 1 Then
                cur统筹金额 = 金额 * (IIf(IsNull(rsTmp!统筹比额), 0, rsTmp!统筹比额) / 100)
                'Modified by 朱玉宝 20031218 地区：福州
'                If intInsure <> TYPE_福建巨龙 Then
                If Not (intinsure = TYPE_福建巨龙 Or intinsure = TYPE_福建省 Or intinsure = TYPE_福州市 Or intinsure = TYPE_南平市) Then
                    bln保险项目 = bln保险项目 And (IIf(IsNull(rsTmp!统筹比额), 0, rsTmp!统筹比额) > 0)
                End If
            ElseIf IIf(IsNull(rsTmp!算法), 0, rsTmp!算法) = 2 Then
                If intinsure = TYPE_徐州六院 Then
                    cur统筹金额 = IIf(金额 > rsTmp!统筹比额, rsTmp!统筹比额, 金额)
                End If
            End If
            
            '如果服务对象不对，也认为它不是医保项目
            lng服务对象 = IIf(IsNull(rsTmp("服务对象")), 3, rsTmp("服务对象")) '缺省认为该大类可同时服务于门诊与住院病人
            If bln门诊 = True Then
                If lng服务对象 <> 1 And lng服务对象 <> 3 Then bln保险项目 = False
            Else
                If lng服务对象 <> 2 And lng服务对象 <> 3 Then bln保险项目 = False
            End If
            
            '再判断各个项目的服务对象
            If intinsure = TYPE_四川眉山 Then
                lng服务对象 = IIf(IsNull(rsTmp("项目服务对象")), 3, rsTmp("项目服务对象")) '缺省认为该项目可同时服务于门诊与住院病人
                If bln门诊 = True Then
                    If lng服务对象 <> 1 And lng服务对象 <> 3 Then bln保险项目 = False
                Else
                    If lng服务对象 <> 2 And lng服务对象 <> 3 Then bln保险项目 = False
                End If
            End If
        End If
        bln保险项目 = bln保险项目 And (IIf(IsNull(rsTmp!项目是否医保), 0, rsTmp!项目是否医保) = 1)
        
        '全额统筹金额
        'If bln全额统筹 And bln保险项目 Then cur统筹金额 = 金额
    Else
'        If intInsure = TYPE_自贡市 Then
'            MsgBox "未找到对应的保险大类，请检查保险项目是否正确设置。", vbInformation, gstrSysName
'            Exit Function
'        End If
    End If
    If intinsure = TYPE_成都内江 Or intinsure = TYPE_贵阳市 Then
        If Not (病人ID = lng病人ID And 收费细目ID = lng收费细目ID) Then
            If intinsure = TYPE_成都内江 Then
                str项目编码_TEMP = GetItemInsure_成都内江(病人ID, 收费细目ID, bln门诊)
            Else
                str项目编码_TEMP = GetItemInsure_贵阳(病人ID, 收费细目ID, bln门诊)
            End If
        End If
        If str项目编码_TEMP = "" Then str项目编码_TEMP = str项目编码
        GetItemInsure = IIf(bln保险项目, 1, 0) & ";" & lng大类ID & ";" & cur统筹金额 & ";" & str项目编码_TEMP
        
        lng病人ID = 病人ID
        lng收费细目ID = 收费细目ID
    Else
        GetItemInsure = IIf(bln保险项目, 1, 0) & ";" & lng大类ID & ";" & cur统筹金额 & ";" & str项目编码
    End If
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        GetItemInsure = gobjInsure_Obj(mintOrder).GetItemInsure(病人ID, 收费细目ID, 金额, bln门诊, intinsure, strAdvance)
    Else
        Select Case intinsure
        Case TYPE_徐州六院
            GetItemInsure = GetItemInsure & ";;" & 医保项目_徐州六院(病人ID, 收费细目ID, 金额, bln门诊, intinsure)
        Case TYPE_徐州农保
            GetItemInsure = GetItemInsure & ";;" & GetItemInfo_徐州(病人ID, 收费细目ID, , True)
        Case TYPE_昭通
            Call CheckInsureItem_昭通(收费细目ID)
        Case TYPE_铜山县
            GetItemInsure = GetItemInsure & ";;" & 医保项目_铜山县(病人ID, 收费细目ID, 金额, bln门诊, intinsure)
        Case Is > 900
            GetItemInsure = GetItemInsure & ";;" & 医保项目_中联(收费细目ID, intinsure)
        End Select
    End If
    
End Function

Public Function TranChargeDetail(ByVal int类别 As Integer, ByVal str单据号 As String, ByVal int性质 As Integer, ByVal int状态 As Integer, _
        str消息 As String, Optional ByVal lng病人ID As Long = 0, Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
'功能: 把发生的费用明细传输到医保中间数据库
'参数:  int类别    1:表示门诊收费；２：住院计帐
'       str单据号  NO
'       int性质    记录性质
'       int状态    记录状态
'       str消息    如果传输过程中有提醒，传回前台程序完成（避免长时间的死锁）
'       lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
'返回: 传输成功与否
'------------------------------------------------------------------------------------------------------------------
'调用模块：
'   1121-门诊收费:记录修改;门诊收费;门诊退费
'   1133-住院记帐:记录修改;住院划价;住院记帐;住院销帐;专项记帐
'   1134-分散记帐:记录修改;住院记帐;住院销帐;专项记帐
'   1135-医技记帐:记录修改;住院记帐;住院销帐;专项记帐
'医嘱系统 (5100)
'   1261 - 医嘱摆药生成
'   1263 - 执行非药医嘱
'
'医护系统
'   402 - 医嘱摆药处理
'   403 - 执行非药医嘱
'支持support门诊收费存为划价单参数时，门诊收费在保存单据后会调用本接口函数
'费用部分处理流程：费用明细上传时，只要有一个医保上传成功，则所有数据正常提交
    Dim intMouse As Integer
    
    '如果是渝北农医，直接退出
    If intinsure = TYPE_渝北农医 Then TranChargeDetail = True: Exit Function
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then TranChargeDetail = True: Exit Function
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        TranChargeDetail = gobjInsure_Obj(mintOrder).TranChargeDetail(int类别, str单据号, int性质, int状态, str消息, lng病人ID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_昭通
                TranChargeDetail = 记帐传输_昭通(str单据号, int性质, str消息, lng病人ID)
            Case TYPE_徐州
                TranChargeDetail = False   ' 记帐传输_徐州(str单据号, int性质, str消息, lng病人ID)
            Case TYPE_徐州六院
                TranChargeDetail = False
            Case TYPE_四川自贡
                If int类别 = 2 Then
                    TranChargeDetail = 处方上传_自贡(str单据号, int性质, int状态, lng病人ID)
                End If
            Case TYPE_余姚
                If int类别 = 1 Then
                    Screen.MousePointer = intMouse
                    str消息 = "门诊记帐数据不能上传到医保前置机，请通过门诊费用管理收取门诊费用"
                    TranChargeDetail = False
                    Exit Function
                End If
                TranChargeDetail = 记帐传输_余姚(str单据号, int性质, str消息, lng病人ID)
            Case TYPE_吉林
                '刘兴宏:20040927
                TranChargeDetail = 处方登记_吉林(int性质, int状态, str单据号)
            Case TYPE_北京尚洋
                TranChargeDetail = False
            Case TYPE_华东
                TranChargeDetail = 记帐传输_华东(str单据号, int性质, str消息, lng病人ID)
            Case TYPE_江苏
                TranChargeDetail = True
            Case TYPE_涪陵
                TranChargeDetail = 记帐传输_涪陵(str单据号, int性质, str消息, lng病人ID)
            Case TYPE_广元
                TranChargeDetail = 记帐传输_广元(str单据号, int性质, str消息, lng病人ID)
            Case TYPE_重庆市
                If int类别 = 2 Then
                    '只处理住院记帐
                    If int状态 = 1 Then
                        TranChargeDetail = 记帐传输_重庆(str单据号, int性质, str消息, lng病人ID)
                    Else
                        '作废单据
                        TranChargeDetail = 记帐作废_重庆(str单据号, int性质, str消息)
                    End If
                End If
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                TranChargeDetail = False
            Case TYPE_重庆壁山
                If int类别 = 2 Then '住院计帐
                    TranChargeDetail = 记帐传输_壁山(str单据号, int性质, int状态, lng病人ID)
                End If
            Case TYPE_重庆中梁山
                If int类别 = 2 Then '住院计帐
                    TranChargeDetail = 记帐传输_中梁山(str单据号, int性质, int状态, lng病人ID)
                End If
            Case TYPE_自贡市
                TranChargeDetail = False
            Case TYPE_泸州市
                If int类别 = 2 Then
                    '只处理住院记帐
                    If int状态 = 1 Then
                        TranChargeDetail = 记帐传输_泸州(str单据号, int性质, lng病人ID)
                    Else
                        TranChargeDetail = 记帐作废_泸州(str单据号, int性质, lng病人ID)
                    End If
                End If
            Case TYPE_铜仁
                If int类别 = 2 Then
                    '只处理住院记帐
                    If int状态 = 1 Then
                        TranChargeDetail = 记帐传输_铜仁(str单据号, int性质, lng病人ID)
                    Else
                        TranChargeDetail = 记帐作废_铜仁(str单据号, int性质, lng病人ID)
                    End If
                End If
            Case TYPE_贵阳市
                TranChargeDetail = True
            Case TYPE_云南省, TYPE_昆明市
                TranChargeDetail = False
            Case TYPE_云南建水
                TranChargeDetail = False
            Case TYPE_成都市
                '200308z012:不实时上传
                TranChargeDetail = False
    '            If int类别 = 2 Then '住院计帐
    '                TranChargeDetail = 记帐传输_成都(str单据号, int性质, int状态, lng病人ID)
    '            End If
            Case type_成都郊县
                If int类别 = 2 Then '住院计帐
                    TranChargeDetail = 记帐传输_成都郊县(str单据号, int性质, int状态, lng病人ID)
                End If
            Case TYPE_新都
                If int类别 = 2 Then '住院计帐
                    TranChargeDetail = 记帐传输_新都(str单据号, int性质, int状态, lng病人ID)
                End If
            Case TYPE_成都莲合
                TranChargeDetail = 传输明细_莲合(str单据号, int性质, int状态, int类别)
            Case TYPE_开县
                '20050124
                TranChargeDetail = 传输明细_开县(str单据号, int性质, int状态, int类别)
                
            Case TYPE_成都南充
                If int类别 = 2 Then
                    '只处理住院记帐
                    If int状态 = 1 Then
                        TranChargeDetail = 处方登记_成都南充(int性质, int状态, str单据号)
                    Else
                        '作废单据
                        TranChargeDetail = 处方删除_成都南充(int性质, int状态, str单据号)
                    End If
                End If
            Case type_米易
                If int类别 = 2 Then
                    '只处理住院记帐
                    TranChargeDetail = 处方登记_米易(int性质, int状态, str单据号)
                End If
            Case TYPE_沈阳市
                If int类别 = 2 Then
                    If int状态 = 1 Then
                        TranChargeDetail = 处方登记_沈阳市(int性质, int状态, str单据号)
                    Else
                        TranChargeDetail = 处方作废_沈阳市(int性质, int状态, str单据号)
                    End If
                End If
            Case TYPE_大连市, TYPE_大连开发区
                If int类别 = 2 Then
                  '刘兴宏(200311)
                    '只处理住院记帐
                    TranChargeDetail = 处方登记_大连(int性质, int状态, str单据号, intinsure)
                End If
            Case TYPE_乐山
                TranChargeDetail = 上传处方_乐山(int性质, int状态, str单据号)
            Case TYPE_重大校园卡
                '刘兴宏(200403)
                TranChargeDetail = False
            Case TYPE_重庆渝北
                 '刘兴宏200407
                 TranChargeDetail = 处方登记_重庆渝北(int性质, int状态, str单据号)
            Case TYPE_兴安
                '刘兴宏:20040827
                TranChargeDetail = 处方登记_兴安(int性质, int状态, str单据号)
            Case TYPE_毕节
                TranChargeDetail = 处方上传_毕节(int性质, int状态, str单据号)
            Case TYPE_黔南
                '刘兴宏:20041022
                TranChargeDetail = 处方登记_黔南(int性质, int状态, str单据号)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                TranChargeDetail = 处方登记_奉庆(int性质, int状态, str单据号)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                TranChargeDetail = 处方登记_成都德阳(int性质, int状态, str单据号)
            Case TYPE_成都内江
                '刘兴宏:20041207
                TranChargeDetail = 处方登记_成都内江(int性质, int状态, str单据号)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                TranChargeDetail = 处方登记_广元旺苍(int性质, int状态, str单据号)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                TranChargeDetail = 处方登记_南充阆中(int性质, int状态, str单据号)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                TranChargeDetail = 处方登记_兴成(int性质, int状态, str单据号)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                TranChargeDetail = 处方登记_神木大兴(int性质, int状态, str单据号)
                
            Case TYPE_渝北农医
                TranChargeDetail = 处方上传_渝北农医(int性质, int状态, str单据号)
            Case TYPE_慈溪农医
                TranChargeDetail = 处方上传_慈溪农医(int性质, int状态, str单据号)
            Case TYPE_山西
                TranChargeDetail = 记帐上传_山西(int性质, int状态, str单据号)
                
            Case TYPE_冕宁
            '20050401
                TranChargeDetail = 记帐上传_冕宁(int性质, int状态, str单据号)
            Case type_铜梁合医
            '20050530
                TranChargeDetail = 处方上传_铜梁合医(str单据号, int性质, int状态, str消息, lng病人ID, type_铜梁合医)
            Case TYPE_铜山县
                TranChargeDetail = 处方上传_铜山县(int类别, str单据号, int性质, int状态, str消息, lng病人ID, intinsure)
            Case Is > 900
                TranChargeDetail = False
        End Select
    End If
    Screen.MousePointer = intMouse
End Function

Public Function LeaveDelSwap(lngPatiID As Long, lngPageID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
'功能：将撤消病人的出院信息
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1132-撤消出院
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select A.出院日期,A.出院病床,Decode(A.出院方式,'正常',0,'死亡',1,'转院',2,9) as 出院方式,B.名称,D.住院号,Sysdate as 经办时间," & _
            " C.卡号,C.医保号,C.密码,C.顺序号 " & _
            " From 病案主页 A,部门表 B,医保病人档案 C,医保病人关联表 E,病人信息 D " & _
            " Where A.病人ID=D.病人ID And A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.入院科室ID=B.ID And A.病人ID=E.病人ID and C.险类=E.险类 and C.医保号=E.医保号 And C.险类=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", lngPatiID, lngPageID, intinsure)
    If rsTmp.EOF Then
        MsgBox "没有此病人或此病人不是医保病人！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        LeaveDelSwap = gobjInsure_Obj(mintOrder).LeaveDelSwap(lngPatiID, lngPageID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_贵阳市
                LeaveDelSwap = 出院登记撤销_贵阳(lngPatiID, lngPageID)
            Case TYPE_四川自贡
                LeaveDelSwap = 出院登记撤销_自贡(lngPatiID, lngPageID)
            Case TYPE_昭通
                LeaveDelSwap = 出院登记撤消_昭通(lngPatiID, lngPageID)
            Case TYPE_徐州农保
                LeaveDelSwap = 出院登记撤消_徐州农保(lngPatiID, lngPageID)
            Case TYPE_徐州
                LeaveDelSwap = 出院登记撤消_徐州(lngPatiID, lngPageID)
            Case TYPE_徐州六院
                LeaveDelSwap = 出院登记撤消_中联(lngPatiID, lngPageID, intinsure)
            Case TYPE_宁海
                LeaveDelSwap = 出院登记撤销_宁海(lngPatiID, lngPageID)
            Case TYPE_余姚
                LeaveDelSwap = True
            Case TYPE_浙江
                LeaveDelSwap = 出院登记撤消_浙江(lngPatiID, lngPageID)
            Case TYPE_吉林
                '20040923刘兴宏加入
                LeaveDelSwap = 出院登记撤销_吉林(lngPatiID, lngPageID)
            Case TYPE_江苏, TYPE_南京市, TYPE_华东, TYPE_北京尚洋
                LeaveDelSwap = True
            Case TYPE_涪陵, TYPE_广元
                LeaveDelSwap = False
            Case TYPE_重庆市
                LeaveDelSwap = 出院登记撤销_重庆(lngPatiID, lngPageID)
            Case TYPE_重庆壁山
                LeaveDelSwap = 出院登记撤销_壁山(lngPatiID, lngPageID)
            Case TYPE_重庆中梁山
                LeaveDelSwap = 出院登记撤消_中梁山(lngPatiID, lngPageID)
            Case TYPE_云南省, TYPE_昆明市
                LeaveDelSwap = 出院登记撤消_云南(lngPatiID, lngPageID, intinsure)
            Case TYPE_成都南充
                LeaveDelSwap = 出院登记撤销_成都南充(lngPatiID)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                LeaveDelSwap = 出院登记撤销_福建巨龙(lngPatiID, intinsure)
            Case type_米易
                LeaveDelSwap = 出院登记撤销_米易(lngPatiID, lngPageID)
            Case TYPE_四川眉山
                LeaveDelSwap = 出院登记撤销_眉山(lngPatiID, lngPageID)
            Case TYPE_沈阳市
                LeaveDelSwap = 出院登记撤销_沈阳市(lngPatiID, lngPageID)
            Case TYPE_大连市, TYPE_大连开发区
                '刘兴宏(200311)
                LeaveDelSwap = 出院登记撤销_大连(lngPatiID, lngPageID)
            Case TYPE_乐山
                LeaveDelSwap = 出院登记撤销_乐山(lngPatiID, lngPageID)
            Case TYPE_重大校园卡
                '刘兴宏(200403)
                LeaveDelSwap = 出院登记撤销_重大校园卡(lngPatiID, lngPageID)
            Case TYPE_重庆银海版
                LeaveDelSwap = 出院登记撤销_重庆银海版(lngPatiID, lngPageID)
            Case TYPE_重庆渝北
                LeaveDelSwap = 出院登记撤销_重庆渝北(lngPatiID, lngPageID)
            Case TYPE_北京
                LeaveDelSwap = 出院登记撤销_北京(lngPatiID, lngPageID)
            Case TYPE_兴安
                '刘兴宏:20040827
                LeaveDelSwap = 出院登记撤销_兴安(lngPatiID, lngPageID)
            Case TYPE_毕节
                LeaveDelSwap = 出院登记撤销_毕节(lngPatiID, lngPageID)
            Case TYPE_黔南
                '刘兴宏:20041022
                LeaveDelSwap = 出院登记撤销_黔南(lngPatiID, lngPageID)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                LeaveDelSwap = 出院登记撤销_奉庆(lngPatiID, lngPageID)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                LeaveDelSwap = 出院登记撤销_成都德阳(lngPatiID, lngPageID)
            Case TYPE_成都内江
                '刘兴宏:20041207
                LeaveDelSwap = 出院登记撤销_成都内江(lngPatiID, lngPageID)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                LeaveDelSwap = 出院登记撤销_广元旺苍(lngPatiID, lngPageID)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                LeaveDelSwap = 出院登记撤销_南充阆中(lngPatiID, lngPageID)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                LeaveDelSwap = 出院登记撤销_兴成(lngPatiID, lngPageID)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                LeaveDelSwap = 出院登记撤销_神木大兴(lngPatiID, lngPageID)
                
            Case TYPE_渝北农医
                LeaveDelSwap = 出院登记撤销_渝北农医(lngPatiID, lngPageID)
            Case TYPE_慈溪农医
                LeaveDelSwap = 出院登记撤销_慈溪农医(lngPatiID, lngPageID)
            '20050304
            Case TYPE_山西
                LeaveDelSwap = 出院登记撤销_山西(lngPatiID, lngPageID)
            Case TYPE_冕宁
            '陈东 20050401
                LeaveDelSwap = 出院登记撤销_冕宁(lngPatiID, lngPageID)
            Case type_铜梁合医
            '陈东 20050530
                LeaveDelSwap = 出院登记撤销_铜梁合医(lngPatiID, lngPageID, type_铜梁合医)
            Case TYPE_铜山县
                LeaveDelSwap = 出院登记撤销_铜山县(lngPatiID, lngPageID, intinsure)
            Case TYPE_成都莲合
                LeaveDelSwap = 入院登记_莲合(lngPatiID, lngPageID)
            Case Is > 900
                LeaveDelSwap = 出院登记撤消_中联(lngPatiID, lngPageID, intinsure)
        End Select
    End If
End Function

Public Function RegistSwap(lng结帐ID As Long, cur金额 As Currency, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊挂号结算
    '入参:lng结帐ID=挂号单保存后的结帐ID,cur金额=个人帐户支付金额
    '     strAdvance:  目的是医保接口能有效区分出是记账模式还是挂号结算模式,以便医保接口能进行锁卡操作
    '        格式为:结算模式|挂号费收取方式|挂号单号|补结算标志,用|分隔
    '        a.结算模式:主要决定病人就诊过程中是采取哪种就诊模式（先诊疗后结算；还是先结算后诊疗）,0-表示先结算后诊疗;1-表示先诊疗后结算
    '        b.挂号费收取方式:表示当前的挂号费采用什么方式收取(记帐或现收): 1表示挂号费采用记帐方式；0表示用现收方式
    '        c.挂号单号:传入具体的挂号单
    '        d.补结算标志:1-表示是补结算接口调用;0-正常挂号接口调用
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-13 17:44:05
    '说明：目前挂号交易只支持全部使用个人帐户结算,不支持其它结算方式及预结算
    '使用模块：1115-自助挂号;1111-挂号管理;1124-保险补充结算
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim arrAdvance
    Dim str医保号 As String, str密码 As String
    Dim lng病人ID As Long, str卡号 As String
    Dim bln记帐 As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim bln补充结算 As Boolean
    
    If strAdvance = "" Then strAdvance = "||"
    
    arrAdvance = Split(strAdvance & "|||||", "|")
    bln记帐 = (Val(arrAdvance(1)) = 1)
    bln补充结算 = Val(arrAdvance(3)) = 1
    
    If Not bln记帐 Then
        If GetCapability(support挂号使用个人帐户, 0, intinsure) And intinsure <> TYPE_江苏 Then
           '松藻医保不需要对每个项目设置对应的医保项目
            If Not GetCapability(support允许不设置医保项目, 0, intinsure) Then
                If bln补充结算 Then
                    gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                        " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
                        " ( Select distinct 收费细目ID,'Y' as TEST From 门诊费用记录  " & _
                        "  Where  结帐ID in (select distinct 收费结帐ID From 费用补充记录 where 结算ID =[2]) ) B,收费细目 C" & _
                        " Where B.收费细目ID=C.ID And B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
                Else
                    gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                        " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
                        " (Select 收费细目ID,'Y' as TEST From 门诊费用记录 Where 结帐ID=[2]) B,收费细目 C" & _
                        " Where B.收费细目ID=C.ID And B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
                End If
                
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", intinsure, lng结帐ID)
                If Not rsTmp.EOF Then
                    MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行医保交易。", vbExclamation, gstrSysName
                    RegistSwap = False: Exit Function
                End If
            End If
            
            If GetCapability(support门诊必须传递明细, 0, intinsure) = False Then
                gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
                    " From 病人预交记录 A,保险帐户 B " & _
                    " Where A.病人ID=B.病人ID And B.险类=[1]" & _
                    " And A.结算方式 in ('个人帐户','医保基金') And A.结帐ID=[2]"
            Else
                '病人该次挂号可能没有发生个人帐户费用，便也要传递费用明细
                gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
                    " From 病人预交记录 A,保险帐户 B " & _
                    " Where A.病人ID=B.病人ID And B.险类=[1] And A.结帐ID=[2]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", intinsure, lng结帐ID)
            If rsTmp.RecordCount = 0 Then
                If Not GetCapability(support挂号必须传递明细, 0, intinsure) Then
                    MsgBox "没有使用保险结算，可以不向医保交易", vbExclamation, gstrSysName
                    RegistSwap = True: Exit Function
                End If
            End If
        
            If rsTmp.RecordCount = 0 Then
                '当发生总金额为零时，则打开的记录数为空
                If bln补充结算 Then
                    gstrSQL = " Select 病人ID,卡号,医保号,密码 From 保险帐户 " & _
                              " Where 病人ID=(Select 病人ID From 费用补充记录 Where 结算id=[1] And Rownum<2)"
                Else
                    gstrSQL = " Select 病人ID,卡号,医保号,密码 From 保险帐户 " & _
                              " Where 病人ID=(Select 病人ID From 门诊费用记录 Where 结帐ID=[1] And Rownum<2)"
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", lng结帐ID)
            End If
            
            lng病人ID = rsTmp!病人ID
            str卡号 = TrimStr(IIf(IsNull(rsTmp!卡号), "", rsTmp!卡号))
            str医保号 = TrimStr(IIf(IsNull(rsTmp!医保号), str卡号, rsTmp!医保号))
            str密码 = TrimStr(IIf(IsNull(rsTmp!密码), "", rsTmp!密码))
        End If
    End If
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        RegistSwap = gobjInsure_Obj(mintOrder).RegistSwap(lng结帐ID, cur金额, intinsure, strAdvance)
    Else
        '如果不是分离部件且使用记帐模式，直接返回成功！
        If bln记帐 Then
            RegistSwap = True
            Exit Function
        End If
        Select Case intinsure
            Case TYPE_徐州六院
                RegistSwap = 门诊挂号_徐州六院(lng结帐ID, intinsure)
            Case TYPE_江苏
                RegistSwap = 挂号结算_江苏(lng结帐ID, cur金额)
            Case TYPE_南京市
                RegistSwap = 挂号结算_南京市(lng结帐ID, intinsure)
            Case TYPE_涪陵, TYPE_广元, TYPE_华东, TYPE_北京尚洋
                RegistSwap = False
            Case TYPE_重庆市
                RegistSwap = 挂号结算_重庆(lng结帐ID)
            Case TYPE_重庆松藻
                RegistSwap = False
            Case TYPE_重庆壁山
                RegistSwap = True
            Case TYPE_自贡市
                RegistSwap = False
            Case TYPE_泸州市
                'Modified by 朱玉宝 2004-01-07
                RegistSwap = 挂号结算_泸州(lng结帐ID)
            Case TYPE_铜仁
                RegistSwap = False
            Case TYPE_贵阳市
                'Modified By 朱玉宝 2004-05-25 原因：医保接口变动
                '------------------------------------------------
                RegistSwap = 门诊挂号_贵阳(lng结帐ID, lng病人ID)
                '------------------------------------------------
            Case TYPE_云南省, TYPE_昆明市
                RegistSwap = False
            Case TYPE_云南建水
                RegistSwap = False
            Case TYPE_成都市
                RegistSwap = 挂号结算_成都(lng结帐ID, lng病人ID, str医保号, str密码, str卡号)
            Case TYPE_成都莲合
                RegistSwap = False
            Case TYPE_开县
                '20050124
                RegistSwap = False
                
            Case type_成都郊县, TYPE_新都
                RegistSwap = False
            Case TYPE_咸阳市
                RegistSwap = False
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                RegistSwap = 挂号结算_福建巨龙(lng结帐ID, intinsure)
            Case type_米易
                RegistSwap = False
            Case TYPE_沈阳市
                RegistSwap = 挂号结算_沈阳市(lng结帐ID)
            Case TYPE_四川眉山
                RegistSwap = 挂号结算_眉山(lng结帐ID, cur金额)
            Case TYPE_大连市, TYPE_大连开发区
                '刘兴宏(200403)
                RegistSwap = 挂号结算_大连(lng结帐ID)
            Case TYPE_吉林
                '刘兴宏(200403)
                RegistSwap = 挂号结算_吉林(lng结帐ID)
            Case TYPE_乐山
                RegistSwap = False
            Case TYPE_重大校园卡
                '刘兴宏(200403)
                RegistSwap = 挂号结算_重大校园卡(lng结帐ID)
            Case TYPE_重庆渝北
                '刘兴宏20040713
                RegistSwap = 挂号结算_重庆渝北(lng结帐ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                RegistSwap = False
            Case TYPE_黔南
                '刘兴宏:20041022
                RegistSwap = False
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                RegistSwap = False
            Case TYPE_成都德阳
                '刘兴宏:20041103
                RegistSwap = False
            Case TYPE_成都内江
                '刘兴宏:20041207
                RegistSwap = False
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                'RegistSwap =False
                '曾明春:20060724
                RegistSwap = True
            Case TYPE_南充阆中
                '刘兴宏:20050419
                RegistSwap = False
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                RegistSwap = False
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                RegistSwap = False
                
            Case TYPE_渝北农医
                RegistSwap = True
            Case TYPE_铜山县
                RegistSwap = 门诊挂号_铜山县(lng结帐ID, intinsure)
            Case Is > 900
                RegistSwap = False
        End Select
    End If
End Function

Public Function RegistDelSwap(lng结帐ID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊挂号作废交易
    '入参:lng结帐ID=作废前的挂号记录的结帐ID
    '     strAdvance:格式:
    '          格式:挂号费收取方式|挂号单号|补结算标志,用|分隔
    '           a.挂号收取方式:1-表示挂号费采用记帐方式;非零表示挂号费用采用的是现收方式
    '             主要目的是医保接口能有效区分出挂号费的收取方式:记账方式还是挂号现收方式,以便医保接口能进行解卡操作.
    '           b.挂号单号: 当前挂号单
    '           c.补结算标志:1-补结算调整;0-挂号调用
    '出参:
    '返回:交易成功返回true；否则，返回false
    '编制:刘兴洪
    '日期:2014-10-13 18:05:19
    '调用模块:1111-门诊挂号管理
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim strSelfNo As String, strSelfPwd As String, str卡号 As String
    Dim rsTmp As New ADODB.Recordset
    Dim curMoney As Currency
    Dim lng病人ID As Long
    Dim lng冲销ID As Long
    Dim bln记帐 As Boolean
    Dim arrAdvance As Variant
    Dim bln补结算 As Boolean
    
    If strAdvance = "" Then strAdvance = "|"
    arrAdvance = Split(strAdvance & "|||||", "|")
    bln记帐 = (Val(arrAdvance(0)) = 1)
    bln补结算 = Val(arrAdvance(2)) = 1
    If Not bln记帐 Then
        If GetCapability(support挂号使用个人帐户, 0, intinsure) Then
            '松藻医保不需要对每个项目设置对应的医保项目
            If Not GetCapability(support允许不设置医保项目, 0, intinsure) Then
                If bln补结算 Then
                    gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                        " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
                        " (Select distinct 收费细目ID,'Y' as TEST From 门诊费用记录 " & _
                        "  Where 结帐ID in (select distinct 收费结帐ID From 费用补充记录 where 结算ID =[2])) B,收费细目 C" & _
                        " Where B.收费细目ID=C.ID and B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
                Else
                    gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                        " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=[1]) A," & _
                        " (Select 收费细目ID,'Y' as TEST From 门诊费用记录 Where 结帐ID=[2]) B,收费细目 C" & _
                        " Where B.收费细目ID=C.ID and B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查对码数据", intinsure, lng结帐ID)
                If Not rsTmp.EOF Then
                    MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行医保交易。", vbExclamation, gstrSysName
                    RegistDelSwap = False: Exit Function
                End If
            End If
            
            '读出个人帐户使用金额
            If GetCapability(support门诊必须传递明细, 0, intinsure) = False Then
                gstrSQL = "Select A.病人ID,A.冲预交,A.结算方式 From 病人预交记录 A " & _
                    " Where A.结算方式 in ('个人帐户','医保基金') And A.结帐ID=[1]"
            Else
                '病人该次收费可能没有发生个人帐户费用，便也要传递费用明细
                gstrSQL = "Select A.病人ID,A.冲预交,A.结算方式 From 病人预交记录 A " & _
                    " Where A.结帐ID=[1]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊退费", lng结帐ID)
            If rsTmp.RecordCount = 0 Then
                MsgBox "没有使用个人帐户或医保基金结算，可以不向医保交易", vbExclamation, gstrSysName
                RegistDelSwap = True
                Exit Function
            End If
            
            lng病人ID = rsTmp("病人ID")
            
            Do Until rsTmp.EOF
                If rsTmp("结算方式") = "个人帐户" Then
                    curMoney = curMoney + rsTmp("冲预交")
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        On Error Resume Next
        '取冲销ID并更新原始结算记录的冲销ID与冲销时间
        If bln补结算 Then
            gstrSQL = "Select a.结算id as 结帐ID " & _
                     " From 费用补充记录 A, 费用补充记录 B " & _
                     " Where a.No = b.No And a.记录性质 = b.记录性质 And a.记录状态 = 2 And b.结算id = [1] " & _
                     " Order By a.登记时间 Desc"
        Else
            gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
                      " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", lng结帐ID)
        lng冲销ID = rsTmp("结帐ID")
        gstrSQL = "zl_保险结算记录_作废(" & lng结帐ID & "," & lng冲销ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "将作废记录与原始记录关联起来")
    End If
    
    On Error GoTo 0
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        RegistDelSwap = gobjInsure_Obj(mintOrder).RegistDelSwap(lng结帐ID, intinsure, strAdvance)
    Else
        If bln记帐 Then
            RegistDelSwap = True
            Exit Function
        End If
        Select Case intinsure
            Case TYPE_江苏
                RegistDelSwap = 挂号结算冲销_江苏(lng结帐ID)
            Case TYPE_南京市
                RegistDelSwap = 挂号结算冲销_南京市(lng结帐ID, intinsure)
            Case TYPE_涪陵, TYPE_广元, TYPE_华东, TYPE_北京尚洋
                RegistDelSwap = False
            Case TYPE_重庆市
                RegistDelSwap = 门诊结算冲销_重庆(lng结帐ID, 0, lng病人ID)
            Case TYPE_重庆松藻
                RegistDelSwap = False
            Case TYPE_重庆壁山
                RegistDelSwap = False
            Case TYPE_自贡市
                RegistDelSwap = False
            Case TYPE_泸州市
                'Modified by 朱玉宝 2004-01-07
                RegistDelSwap = 挂号结算冲销_泸州(lng结帐ID)
            Case TYPE_铜仁
                RegistDelSwap = False
            Case TYPE_云南省, TYPE_昆明市
                RegistDelSwap = False
            Case TYPE_云南建水
                RegistDelSwap = False
            Case TYPE_贵阳市
                RegistDelSwap = 门诊结算冲销_贵阳(lng结帐ID, 0, 0)
            Case TYPE_成都市
                RegistDelSwap = False
            Case TYPE_成都莲合
                RegistDelSwap = False
            Case TYPE_开县
                '20050124
                RegistDelSwap = False
                
            Case type_成都郊县, TYPE_新都
                RegistDelSwap = False
            Case TYPE_咸阳市
                RegistDelSwap = False
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                RegistDelSwap = 挂号结算冲销_福建巨龙(lng结帐ID, intinsure)
            Case type_米易
                RegistDelSwap = False
            Case TYPE_四川眉山
                RegistDelSwap = 挂号结算冲销_眉山(lng结帐ID)
            Case TYPE_大连市, TYPE_大连开发区
                '刘兴宏(200311)
                RegistDelSwap = 挂号冲销_大连(lng结帐ID)
            Case TYPE_吉林
                RegistDelSwap = 挂号冲销_吉林(lng结帐ID)
            
            Case TYPE_沈阳市
                RegistDelSwap = 挂号冲销_沈阳市(lng结帐ID)
            Case TYPE_乐山
                RegistDelSwap = False
            Case TYPE_重大校园卡
                '刘兴宏(200403)
                RegistDelSwap = 挂号冲销_重大校园卡(lng结帐ID)
            Case TYPE_重庆渝北
                '刘兴宏(20040714)
                RegistDelSwap = 挂号冲销_重庆渝北(lng结帐ID)
            Case TYPE_兴安
                '刘兴宏:20040827
                RegistDelSwap = False
            Case TYPE_黔南
                '刘兴宏:20041022
                RegistDelSwap = False
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                RegistDelSwap = False
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                RegistDelSwap = False
            Case TYPE_成都内江
                '刘兴宏:20041207
                RegistDelSwap = False
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                RegistDelSwap = False
            Case TYPE_南充阆中
                '刘兴宏:20050419
                RegistDelSwap = False
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                RegistDelSwap = False
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                RegistDelSwap = False
                
            Case TYPE_渝北农医
                RegistDelSwap = True
            Case TYPE_徐州六院
                RegistDelSwap = 门诊结算冲销_中联(lng结帐ID, curMoney, lng病人ID, intinsure)
            Case TYPE_铜山县
                RegistDelSwap = 门诊挂号冲销_铜山县(lng结帐ID, intinsure)
            Case Is > 900
                RegistDelSwap = False
        End Select
    End If
End Function

Public Function ComeInDelSwap(lngPatiID As Long, lngPageID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
'功能：将撤销入院信息发送医保前置服务器确认；
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1131-取消入院
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select A.出院日期,A.出院病床,Decode(A.出院方式,'正常',0,'死亡',1,'转院',2,9) as 出院方式,B.名称,D.住院号,Sysdate as 经办时间," & _
            " C.卡号,C.医保号,C.密码,C.顺序号 " & _
            " From 病案主页 A,部门表 B,保险帐户 C,病人信息 D " & _
            " Where A.病人ID=D.病人ID And A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.入院科室ID=B.ID And A.病人ID=C.病人ID And C.险类=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", lngPatiID, lngPageID, intinsure)
    If rsTmp.EOF Then
        MsgBox "没有此病人或此病人不是医保病人！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        ComeInDelSwap = gobjInsure_Obj(mintOrder).ComeInDelSwap(lngPatiID, lngPageID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_南京市
                    ComeInDelSwap = True
            Case TYPE_四川自贡
                ComeInDelSwap = 入院登记撤销_自贡(lngPatiID, lngPageID)
            Case TYPE_昭通
                ComeInDelSwap = 入院登记撤消_昭通(lngPatiID, lngPageID)
            Case TYPE_徐州农保
                ComeInDelSwap = 入院登记撤消_徐州农保(lngPatiID, lngPageID)
            Case TYPE_徐州
                ComeInDelSwap = 入院登记撤消_徐州(lngPatiID, lngPageID)
            Case TYPE_徐州六院
                ComeInDelSwap = 出院登记_中联(lngPatiID, lngPageID, intinsure)
            Case TYPE_宁海
                ComeInDelSwap = 入院登记撤销_宁海(lngPatiID, lngPageID)
            Case TYPE_余姚
                ComeInDelSwap = 出院登记_余姚(lngPatiID, lngPageID)
            Case TYPE_浙江
                ComeInDelSwap = 入院登记撤消_浙江(lngPatiID, lngPageID)
            Case TYPE_吉林
              '刘兴宏(20040924)
                ComeInDelSwap = 入院登记撤销_吉林(lngPatiID, lngPageID)
            Case TYPE_北京尚洋
                ComeInDelSwap = 出院登记_北京尚洋(lngPatiID, lngPageID)
            Case TYPE_华东
                ComeInDelSwap = 出院登记_华东(lngPatiID, lngPageID)
            Case TYPE_江苏
                ComeInDelSwap = 入院登记撤消_江苏(lngPatiID, lngPageID)
            Case TYPE_重庆市
                ComeInDelSwap = 出院登记_重庆(lngPatiID, lngPageID)
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                ComeInDelSwap = 出院登记_松藻(lngPatiID, lngPageID)
            Case TYPE_重庆壁山
                ComeInDelSwap = 出院登记_壁山(lngPatiID, lngPageID)
            Case TYPE_重庆中梁山
                ComeInDelSwap = 出院登记_中梁山(lngPatiID, lngPageID)
            Case TYPE_自贡市
                ComeInDelSwap = 出院登记_中软(lngPatiID)
            Case TYPE_泸州市
                ComeInDelSwap = 出院登记_泸州(lngPatiID)
            Case TYPE_铜仁
                ComeInDelSwap = 出院登记_铜仁(lngPatiID)
            Case TYPE_贵阳市
                ComeInDelSwap = 出院登记_贵阳(lngPatiID, lngPageID, True)
            Case TYPE_云南省, TYPE_昆明市
                If IsNull(rsTmp!顺序号) Then
                    MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
                    Exit Function
                End If
                ComeInDelSwap = 出院登记_云南(lngPatiID, lngPageID, rsTmp!顺序号, intinsure, False, True)
            Case TYPE_云南建水
                If IsNull(rsTmp!顺序号) Then
                    MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
                    Exit Function
                End If
                ComeInDelSwap = 入院登记撤销_云南建水(lngPatiID, lngPageID, rsTmp!顺序号)
            Case TYPE_成都市
                ComeInDelSwap = 出院登记_成都(lngPatiID, lngPageID, rsTmp)
            Case TYPE_成都莲合
                ComeInDelSwap = 出院登记_莲合(lngPatiID, lngPageID)
            Case TYPE_开县
                '20050124
                ComeInDelSwap = 出院登记_开县(lngPatiID, lngPageID)
                
            Case type_成都郊县
                ComeInDelSwap = 出院登记_成都郊县(lngPatiID, lngPageID)
            Case TYPE_新都
                ComeInDelSwap = 出院登记_新都(lngPatiID, lngPageID)
            Case TYPE_成都南充
                ComeInDelSwap = 出院登记_成都南充(lngPatiID)
            Case TYPE_咸阳市
                ComeInDelSwap = 出院登记_咸阳(lngPatiID)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                ComeInDelSwap = 出院登记_福建巨龙(intinsure, lngPatiID, lngPageID, True)
            Case type_米易
                ComeInDelSwap = 入院登记撤销_米易(lngPatiID, lngPageID)
            Case TYPE_四川眉山
                ComeInDelSwap = 入院登记撤销_眉山(lngPatiID, lngPageID)
            Case TYPE_沈阳市
                ComeInDelSwap = 入院登记撤销_沈阳市(lngPatiID, lngPageID)
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                ComeInDelSwap = 入院登记撤销_大连(lngPatiID, lngPageID, intinsure)
            Case TYPE_乐山
                ComeInDelSwap = 入院登记撤销_乐山(lngPatiID, lngPageID)
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                ComeInDelSwap = 入院登记撤销_重大校园卡(lngPatiID, lngPageID)
            Case TYPE_重庆银海版
                ComeInDelSwap = 入院登记撤销_重庆银海版(lngPatiID, lngPageID)
            Case TYPE_重庆渝北
                ComeInDelSwap = 入院登记撤销_重庆渝北(lngPatiID, lngPageID)
            Case TYPE_北京
                ComeInDelSwap = 入院登记撤销_北京(lngPatiID, lngPageID)
            Case TYPE_兴安
                '刘兴宏:20040827
                ComeInDelSwap = 入院登记撤销_兴安(lngPatiID, lngPageID)
            Case TYPE_毕节
                ComeInDelSwap = 入院登记撤销_毕节(lngPatiID, lngPageID)
            Case TYPE_黔南
                '刘兴宏:20041022
                ComeInDelSwap = 入院登记撤销_黔南(lngPatiID, lngPageID)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                ComeInDelSwap = 入院登记撤销_奉庆(lngPatiID, lngPageID)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                ComeInDelSwap = 入院登记撤销_成都德阳(lngPatiID, lngPageID)
            Case TYPE_成都内江
                '刘兴宏:20041207
                ComeInDelSwap = 入院登记撤销_成都内江(lngPatiID, lngPageID)
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                ComeInDelSwap = 入院登记撤销_广元旺苍(lngPatiID, lngPageID)
            Case TYPE_南充阆中
                '刘兴宏:20050419
                ComeInDelSwap = 入院登记撤销_南充阆中(lngPatiID, lngPageID)
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                ComeInDelSwap = 入院登记撤销_兴成(lngPatiID, lngPageID)
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                ComeInDelSwap = 入院登记撤销_神木大兴(lngPatiID, lngPageID)
                
            Case TYPE_渝北农医
                ComeInDelSwap = 入院登记撤销_渝北农医(lngPatiID, lngPageID)
            Case TYPE_慈溪农医
                ComeInDelSwap = 入院登记撤销_慈溪农医(lngPatiID, lngPageID)
            Case TYPE_山西
                ComeInDelSwap = 撤销入院登记_山西(lngPatiID, lngPageID)
            Case TYPE_冕宁
                ComeInDelSwap = 撤销入院登记_冕宁(lngPatiID, lngPageID)
            Case type_铜梁合医
                ComeInDelSwap = 入院登记撤销_铜梁合医(lngPatiID, lngPageID, type_铜梁合医)
                '20050530 陈东
            Case TYPE_铜山县
                ComeInDelSwap = 入院登记撤销_铜山县(lngPatiID, lngPageID, intinsure)
        '20050401 陈东
            Case Is > 900
                '中联医保
                ComeInDelSwap = 出院登记_中联(lngPatiID, lngPageID, intinsure)
        End Select
    End If
End Function

Public Function ModiPatiSwap(lngPatiID As Long, lngPageID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
'功能：在院病人的床位变动,转科,医生变化及诊断情况等相关信息变化时调用此接口
'参数：lngPatiID-病人ID；lngPageID-主页ID
'      strAdvice=第1分隔：调用场合，1=入出管理、2=首页整理
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1132-病人入出管理(当病人状态发生变化时调用,如：床位变动,转科,医生变化等.)
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then ModiPatiSwap = True: Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        ModiPatiSwap = gobjInsure_Obj(mintOrder).ModiPatiSwap(lngPatiID, lngPageID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_华东
                ModiPatiSwap = True
            Case TYPE_昭通
                ModiPatiSwap = 转科转床_昭通(lngPatiID, lngPageID)
            Case TYPE_宁海
                ModiPatiSwap = 住院信息变动_宁海(lngPatiID, lngPageID)
            Case TYPE_余姚
                ModiPatiSwap = 转科转床_余姚(lngPatiID, lngPageID)
'            Case TYPE_江苏
'                ModiPatiSwap = 转科转床_江苏(lngPatiID, lngPageID)
            Case TYPE_涪陵, TYPE_广元, TYPE_徐州, TYPE_徐州农保, TYPE_徐州六院
                ModiPatiSwap = True
            Case TYPE_重庆市
                ModiPatiSwap = True
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                ModiPatiSwap = True
            Case TYPE_重庆壁山
                ModiPatiSwap = True
            Case TYPE_重庆中梁山
                ModiPatiSwap = True
            Case TYPE_自贡市
                ModiPatiSwap = True
            Case TYPE_泸州市
                ModiPatiSwap = True
            Case TYPE_铜仁
                ModiPatiSwap = True
            Case TYPE_贵阳市
                ModiPatiSwap = True
            Case TYPE_云南省, TYPE_昆明市
                ModiPatiSwap = True
            Case TYPE_云南建水
                ModiPatiSwap = True
            Case TYPE_成都市
                ModiPatiSwap = True
            Case TYPE_成都莲合
                ModiPatiSwap = True
            Case TYPE_开县
                '20050124
                ModiPatiSwap = True
            Case TYPE_南京市
                '芮奇
                ModiPatiSwap = True
            Case TYPE_江苏
                ModiPatiSwap = True
                
            Case type_成都郊县, TYPE_新都
                ModiPatiSwap = True
            Case TYPE_成都南充
                ModiPatiSwap = True
            Case TYPE_咸阳市
                ModiPatiSwap = True
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                ModiPatiSwap = True
            Case type_米易
                ModiPatiSwap = True
            Case TYPE_四川眉山
                ModiPatiSwap = True
            Case TYPE_沈阳市
                ModiPatiSwap = True
            Case TYPE_大连市, TYPE_大连开发区
                  '刘兴宏(200311)
                ModiPatiSwap = 在院病人信息_大连(lngPatiID, lngPageID, intinsure)
            Case TYPE_乐山
                ModiPatiSwap = True
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                ModiPatiSwap = True
            Case TYPE_重庆银海版
                ModiPatiSwap = True
            Case TYPE_重庆渝北
                ModiPatiSwap = True
            Case TYPE_兴安
                '刘兴宏:20040827
                ModiPatiSwap = 在院病人信息_兴安(lngPatiID, lngPageID)
            Case TYPE_毕节
                ModiPatiSwap = 病人变动记录上传_毕节(lngPatiID, lngPageID)
            Case TYPE_黔南
                '刘兴宏:20041022
                ModiPatiSwap = True
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                ModiPatiSwap = True
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                ModiPatiSwap = True
            Case TYPE_成都内江
                '刘兴宏:20041207
                ModiPatiSwap = True
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                ModiPatiSwap = True
            Case TYPE_南充阆中
                '刘兴宏:20050419
                ModiPatiSwap = True
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                ModiPatiSwap = True
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                ModiPatiSwap = True
            Case TYPE_冕宁
            ' 陈东20050401
                ModiPatiSwap = 病人变动记录上传_冕宁(lngPatiID, lngPageID)
            Case TYPE_铜山县
                ModiPatiSwap = 病人变动_铜山县(lngPatiID, lngPageID, intinsure)
            Case TYPE_北京尚洋
                ModiPatiSwap = True
            Case Is > 900
                '中联医保
                ModiPatiSwap = True
        End Select
    End If
End Function

'刘兴宏:2004/08/30
'主要是区分在输入单据时提示信息，同时可以返回摘要的相关信息.

Public Function GetItemInfo(ByVal intinsure As Integer, ByVal lngPatiID As Long, ByVal lngItemID As Long, _
    Optional ByVal str摘要 As String, Optional intType As Integer = 0, Optional ByVal str备注 As String, _
    Optional ByRef strAdvance As String = "") As String
    '功能：获取医保项目的相关信息，目前仅大连使用，返回指定项目的报销比例
    '参数：lngPatiID-病人ID；lngItemID-收费细目ID
    '      str摘要=用于修改
    '      str备注=医嘱则传递医嘱内容，以便操作员区分明细；记帐可传可不传；或者是其它信息
    '      intType-调用类型(0-医嘱,1-门诊收费,2-住院记帐)
    '返回：不为空,返回摘要中；否则，表示无摘要
    '------------------------------------------------------------------------------------------------------------------
    '调用模块：所有记费模块使用，如住院记帐，门诊收费等
    Dim blnMsg As Boolean
    
    If intinsure = 0 Then
        '        If Not (InStr(1, "," & gstrInsure & ",", ",82,") > 0 Or InStr(1, "," & gstrInsure & ",", ",83,") > 0 _
        '         Or InStr(1, ",82,83,", "," & GetSetting("ZLSOFT", "公共全局", "医保类别", 0) & ",") > 0) Then
        '            Exit Function
        '        End If
        Dim rsTmp1 As New ADODB.Recordset
        '如果保险类别的项目提示等于1，则说明现金病人也需要提示
        gstrSQL = " Select 序号,名称 From 保险类别 Where nvl(项目提示,0)=[1]"
        Set rsTmp1 = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否存在现金病人进行提示的接口", 1)
        If rsTmp1.RecordCount = 0 Then
            gstrSQL = "Select 序号,名称 From 保险类别 where 序号 =[1] or 序号=[2]"
            Set rsTmp1 = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否存在现金病人进行提示的接口", 82, 83)
            If rsTmp1.EOF Then
                gstrSQL = "Select 序号,名称 From 保险类别 where 序号 =[1]"
                Set rsTmp1 = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否存在现金病人进行提示的接口", 669)
                If rsTmp1.EOF Then
                    Exit Function
                Else
                    intinsure = TYPE_吉林省农合
                    blnMsg = True
                End If
            Else
                intinsure = TYPE_大连市
                blnMsg = True
            End If
        Else
            intinsure = rsTmp1!序号
            blnMsg = True
        End If
    End If
    
    '大连医保继续往下执行
    '其它医保则检查，如果是医嘱调用直接退出
    If blnMsg Then
        
    Else
        '应周海全要求开放 2007-10-19
'        If intType = 0 Then Exit Function
    End If
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        GetItemInfo = gobjInsure_Obj(mintOrder).GetItemInfo(intinsure, lngPatiID, lngItemID, str摘要, intType, str备注, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_昭通
                GetItemInfo = GetItemInfo_昭通(intType, lngPatiID, lngItemID, str摘要, str备注)
            Case TYPE_徐州农保, TYPE_徐州
                If intinsure = TYPE_徐州 Then
                    If intType <> 1 Then GetItemInfo = GetItemInfo_徐州(lngPatiID, lngItemID, str摘要)
                Else
                    GetItemInfo = GetItemInfo_徐州(lngPatiID, lngItemID, str摘要)
                End If
            Case TYPE_吉林
                '刘兴宏20040923加入
                GetItemInfo = ""
            Case TYPE_涪陵, TYPE_广元, TYPE_华东, TYPE_余姚, TYPE_北京尚洋
                GetItemInfo = ""
            Case TYPE_重庆市
                GetItemInfo = ""
            Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
                GetItemInfo = ""
            Case TYPE_重庆壁山
                GetItemInfo = ""
            Case TYPE_重庆中梁山
                GetItemInfo = ""
            Case TYPE_自贡市
                GetItemInfo = ""
            Case TYPE_泸州市
                GetItemInfo = ""
            Case TYPE_铜仁
                GetItemInfo = ""
            Case TYPE_贵阳市
                Call 贵阳特殊药品提示(lngItemID)
            Case TYPE_云南省, TYPE_昆明市
                GetItemInfo = ""
            Case TYPE_云南建水
                GetItemInfo = ""
            Case TYPE_成都市
                GetItemInfo = ""
            Case TYPE_成都莲合
                GetItemInfo = ""
            Case TYPE_开县
                '20050124
                GetItemInfo = ""
                
            Case type_成都郊县, TYPE_新都
                GetItemInfo = ""
            Case TYPE_成都南充
                GetItemInfo = ""
            Case TYPE_咸阳市
                GetItemInfo = ""
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                GetItemInfo = ""
            Case type_米易
                GetItemInfo = ""
            Case TYPE_四川眉山
                GetItemInfo = ""
            Case TYPE_沈阳市
                GetItemInfo = ""
            Case TYPE_大连市, TYPE_大连开发区
                If intType <> 1 Then
                    '只有门诊记帐不提示
                    GetItemInfo = GetItemInfo_大连(lngPatiID, lngItemID, intinsure)
                End If
            Case TYPE_吉林省农合
                If intType <> 1 Then
                    GetItemInfo = GetItemInfo(intinsure, lngPatiID, lngItemID, str摘要, intType, str备注, strAdvance)
                End If
            Case TYPE_乐山
                '如果该项目可能做为特殊项目才提示操作员，本明细是否做为特殊项目上传
                Dim int可做为特殊项目 As Integer
                Dim rsTmp As New ADODB.Recordset
                
                Call DebugTool("判断可否做为特殊项目")
                int可做为特殊项目 = 0
                gstrSQL = "Select 附注 From 保险支付项目 Where 险类=[1] And 收费细目ID=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "判断医保项目是否允许做为特殊项目", TYPE_乐山, lngItemID)
                If Not rsTmp.EOF Then
                    If Not IsNull(rsTmp!附注) Then
                        If UBound(Split(rsTmp!附注, "|")) >= 2 Then
                            int可做为特殊项目 = Val(Split(rsTmp!附注, "|")(2))
                        End If
                    End If
                End If
                
                GetItemInfo = ""
                If int可做为特殊项目 = 1 Then
                    If MsgBox("该项目做为特殊项目上传至医保中心吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        GetItemInfo = "1"
                    End If
                End If
            Case TYPE_重大校园卡
                  '刘兴宏(200403)
                GetItemInfo = ""
            Case TYPE_重庆银海版
                GetItemInfo = ""
            Case TYPE_重庆渝北
                GetItemInfo = ""
            Case TYPE_兴安
                'GetItemInfo = ""
                If intType = 1 Then
                    GetItemInfo = GetItemInfo_兴安(lngPatiID, lngItemID, str摘要, intType)
                End If
            Case TYPE_黔南
              '刘兴宏:20041022
                GetItemInfo = GetItemInfo_黔南(lngPatiID, lngItemID, str摘要, intType)
            Case TYPE_临沧奉庆
                '刘兴宏:20041231
                GetItemInfo = GetItemInfo_奉庆(lngPatiID, lngItemID, str摘要, intType)
                
            Case TYPE_成都德阳
                '刘兴宏:20041103
                GetItemInfo = ""
            Case TYPE_成都内江
                '刘兴宏:20041207
                GetItemInfo = ""
            Case TYPE_广元旺苍
                '刘兴宏:20041210
                GetItemInfo = ""
            Case TYPE_南充阆中
                '刘兴宏:20050419
                GetItemInfo = ""
                
            Case TYPE_兴成核工业
                '刘兴宏:20050226
                GetItemInfo = ""
            Case TYPE_陕西大兴
                '刘兴宏:20050316
                GetItemInfo = ""
                
            Case TYPE_山西
                '陈东 20050311
                GetItemInfo = ""
           Case TYPE_冕宁
                '陈东 20050401
                GetItemInfo = ""
            Case type_铜梁合医
                '陈东 20050530
                GetItemInfo = ""
            Case TYPE_铜山县
                '陈东 20051202
                If intType <> 1 Then GetItemInfo = 医保信息_铜山县(lngItemID, intType)
            Case Is > 900
                '中联医保
                GetItemInfo = ""
        End Select
    End If
End Function

Public Sub ChooseDisease(ByVal lngPatiID As Long, ByVal lngPageID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "")
'功能：在病人入院管理模块中，供选择病人的出院病种
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回真
'------------------------------------------------------------------------------------------------------------------
'调用模块：病人入出管理
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Sub
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Sub
        Call gobjInsure_Obj(mintOrder).ChooseDisease(lngPatiID, lngPageID, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_重庆市
                Call 更新出院疾病_重庆(lngPatiID, lngPageID)
            Case TYPE_毕节
                Call 更新病种_毕节(lngPatiID, lngPageID)
            Case TYPE_宁海
                Call 住院信息变动_宁海(lngPatiID, lngPageID)
            Case TYPE_慈溪农医
                Call 更新病种_慈溪农医(lngPatiID, lngPageID)
            Case TYPE_吉林
                Call 更新病种_吉林(lngPatiID, lngPageID)
            '陈东 20050304
            Case TYPE_山西
                Call 更新病种_山西(lngPatiID, lngPageID)
            '陈东 20050620
            Case TYPE_铜山县
                Call 病种选择_铜山县(lngPatiID, lngPageID, intinsure)
            Case TYPE_大连市, TYPE_大连开发区
                Call 更新病种_大连(lngPatiID, lngPageID, intinsure)
                '刘兴宏:200610
            Case TYPE_重庆渝北
                Call 病种更改_渝北(lngPatiID, lngPageID, intinsure)
            Case TYPE_贵阳市
                Call 病种选择_贵阳(lngPatiID)
        End Select
    End If
End Sub

Public Function IdentifyCancel(ByVal bytType As Byte, ByVal lng病人ID As Long, Optional ByVal intinsure As Integer = 0, _
    Optional ByRef strAdvance As String = "") As Boolean
    '说明：用于医保病人身份验证成功后，取消就诊的情况
    '例：就诊登记成功，但发现病人选择错误；已录入明细并完成虚拟结算，病人因某种原因不想继续进行结算操作
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    IdentifyCancel = True
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        IdentifyCancel = gobjInsure_Obj(mintOrder).IdentifyCancel(bytType, lng病人ID, intinsure, strAdvance)
    Else
        Select Case intinsure
        Case TYPE_重庆市
            Call 取消就诊_重庆市(bytType, lng病人ID)
        Case TYPE_渝北农医
            Call 取消就诊登记_渝北农医(bytType, lng病人ID)
        Case TYPE_毕节
            Call 取消就诊_毕节
        Case TYPE_山西
        '陈东 20050311
            Call 取消就诊_山西
        Case TYPE_冕宁
        '陈东 20050311
            Call 取消就诊_冕宁
        Case TYPE_陕西大兴
            Call 门诊虚拟结算取消_神木大兴(bytType, lng病人ID)
        Case TYPE_四川自贡
            IdentifyCancel = 就诊登记取消_自贡()
        Case TYPE_铜山县
        '陈东 20050620
            IdentifyCancel = 取消就诊_铜山县(bytType, lng病人ID, intinsure)
        Case TYPE_南京市
            IdentifyCancel = gblnCancel_南京
        End Select
    End If
End Function

Public Function TranElecDossier(ByVal bytType As Byte, ByVal lng病人ID As Long, ByVal lng就诊号 As Long, _
    Optional ByVal intinsure As Integer = 0, Optional ByVal strAdvance As String) As Boolean
    '功能：完成门诊/住院病人电子病历内容的上传
    '时机：医嘱发送时
    '参数说明：
    'lng就诊号，门诊号或主页ID
    'bytType，1-门诊医生工作站；2-住院医生工作站
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        TranElecDossier = gobjInsure_Obj(mintOrder).TranElecDossier(bytType, lng病人ID, lng就诊号, intinsure, strAdvance)
    Else
        
    End If
End Function

Public Sub BusinessAffirm(ByVal intBusiness As Integer, ByVal blnResult As Boolean, Optional ByVal intinsure As Integer = 0, _
    Optional ByVal strAdvance As String)
    '功能说明：对交易进行确认或取消
    'strBusiness：交易名称，对应于枚举变量 交易Enum
    'blnResult：TRUE表示提交成功，FALSE表示发生异常，需要取消医保交易
    
    '医保接口内增加方法：BusinessAffirm，用来确认或取消某个交易，调用流程：
    '1 ?HIS处理
    '2 ?医保交易成功
    '3 ?HIS提交
    '4 ?调用BusinessAffirm确认医保交易
    '
    '需要考虑从第3步开始，如果出现异常，就调用确认交易，传入FALSE表示需要取消医保交易
    '第3步以前出现任何异常不在HIS考虑范围内?
    '
    '本次修改仅要求：门诊结算、门诊结算作废、住院结算、住院结算作废这四个交易处，HIS进行相应修改。
    Dim strMsg As String
    On Error Resume Next            '避免分离部件没有提供此方法时调用出错
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Sub
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Sub
        strMsg = gobjInsure_Obj(mintOrder).BusinessAffirm(intBusiness, blnResult, intinsure, strAdvance)
    Else
        Select Case intinsure
            Case TYPE_贵阳市
                Call BusinessAffirm_贵阳市(intBusiness, blnResult, intinsure, strAdvance)
            Case TYPE_徐州六院
                Call BusinessAffirm_徐州六院(intBusiness, blnResult, intinsure, strAdvance)
            Case Else
            
        End Select
    End If
    
    '确认或取消失败
    If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
    
End Sub

Public Function InsureChoose() As Integer
    Dim strSelect As String
    '说明：当本地支持多种险类时，用来确定目前要使用的险类
    '如果成功选择，返回险类，否则返回零
    strSelect = GetSetting("ZLSOFT", "公共全局", "本地支持的医保", "")
    If strSelect = "" Then Exit Function
    
    If InStr(1, strSelect, ",") = 0 Then
        InsureChoose = Val(strSelect)
    Else
        InsureChoose = frm选择当前医保.ShowSelect()
    End If
End Function
Public Function GetAvailabilityInsures() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本地支持的险类
    '返回:返回本地支持的险类(多个用逗号分离):如:12,80...
    '编制:刘兴洪
    '日期:2015-03-25 15:30:19
    '问题:81661
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetAvailabilityInsures = GetSetting("ZLSOFT", "公共全局", "本地支持的医保", "")
End Function



Public Sub InsureSupport()
    frm设置本地支持的医保.Show 1
End Sub

Private Sub Class_Terminate()
    glngInstanceCount = glngInstanceCount - 1
    If glngInstanceCount > 0 Then Exit Sub
    
    Call ReleaseME
End Sub

Public Sub ReleaseME()
    Dim arrInsure
    Dim intDO As Integer, intCOUNT As Integer, intObject As Integer
    '自动关闭所有医保接口
    On Error Resume Next
    
    Call DebugTool("进入zl9Insure的Class_Terminate();gstrinsure=" & gstrInsure)
    If gstrInsure = "" Then Exit Sub
    gstrInsure = Mid(gstrInsure, 2)
    arrInsure = Split(gstrInsure, ",")
    intCOUNT = UBound(arrInsure)
    
    '循环关闭非分离的医保部件
    Call DebugTool("循环关闭非分离的医保部件")
    For intDO = 0 To intCOUNT
        If Not IsApartComponents(Val(arrInsure(intDO))) Then
            Call gclsInsure.EndInsure(arrInsure(intDO))
        End If
    Next
    gstrInsure = ""
    
    Call DebugTool("关闭分离的医保接口部件")
    '关闭分离的医保接口部件
    intCOUNT = UBound(gobjInsure_Obj)
    For intObject = 0 To intCOUNT
        Call gobjInsure_Obj(intObject).EndInsure
        Set gobjInsure_Obj(intObject) = Nothing
    Next
End Sub

Public Function ReadCard(ByVal intinsure As Integer, Optional intReturn As Integer, Optional strAdvance As String) As String
    '通用IC卡接口调用医保部件读医保IC卡接口
    'intInsure :传入的险类，如413,412等
    'intReturn :返回调用成功的标志。该值为1，表示读IC卡成功，返回的字符串为卡号；该值<>1，表示读IC卡失败，返回的字符串为出错信息。
    'strAdvance:保留扩展参数，备用。
    
    On Error GoTo errHandle
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        ReadCard = gobjInsure_Obj(mintOrder).ReadCard(intinsure, intReturn, strAdvance)
    Else
        '非分离医保
'        Select Case intInsure
'        End Select
    End If
    Exit Function
    If intReturn <> 1 Then
        intReturn = 0
        If ReadCard = "" Then ReadCard = "调用医保接口读IC卡出现未知错误！"
    End If
errHandle:
    intReturn = 0
    ReadCard = Err.Description
End Function

Public Function CheckItem(ByVal intinsure As Integer, ByVal intType As Integer, ByVal intMode As Integer, _
    ByVal rsDetail As ADODB.Recordset, Optional strAdvance As String) As Boolean
    'zl9Insure直接先赋值为真，如果存在函数则调
    '记录集格式说明：病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室
    'intType:0-门诊;1-住院
    'intMode:0-录入明细时的常规检查;1-保存单据前的汇总检查;2-检查所有项目（用于调划价单）
    '函数返回值:False表示无异常;True表示禁止录入
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        CheckItem = gobjInsure_Obj(mintOrder).CheckItem(intinsure, intType, intMode, rsDetail, strAdvance)
    Else
        
    End If
End Function

Public Sub RePrintBill(ByVal intinsure As Integer, ByVal lng结帐ID As Long, ByVal str票据号 As String, Optional strAdvance As String)
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Sub
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Sub
        Call gobjInsure_Obj(mintOrder).RePrintBill(intinsure, lng结帐ID, str票据号, strAdvance)
    Else
        
    End If
End Sub

Public Function CheckInsureValid(ByVal intinsure As Integer) As Boolean
    '如果没有安装医保部件则返回失败,不允许办理门诊退费及住院结帐作废
    Dim intOrder As Integer

    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
    End If
        CheckInsureValid = True
End Function


