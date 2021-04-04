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
    End If
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前数据库登录用户名
    Dim intinsure As Integer
    Dim intSelect As Integer
    
    Set gcnOracle = cnMain
    Set mobjOracle = cnMain
    
    gstrDbUser = strDBUser
    glngSys = lngSys
    
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
        Call gobjInsure_Obj(mintOrder).BHCodeMan(lngSys, lngModul, cnMain, lngMain, strDBUser, intinsure)
    End If
End Sub

Private Sub Class_Initialize()
    '94352
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
    
    EndInsure = True
End Function

Public Function Identify2(strCard As String, strPass As String, Optional bytType As Byte, Optional lng病人ID As Long = 0, _
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
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        SelfBalance = gobjInsure_Obj(mintOrder).SelfBalance(病人ID, str医保号, bytFlag, cur透支额, intinsure, strAdvance)
    Else
        Select Case intinsure
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
                '20050530 陈东
            Case Is > 900
                ClinicPreSwap = 门诊虚拟结算(rsDetail, str结算方式, intinsure)
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
        gstrSQL = " Select 1 AS 性质,A.NO,B.医保号,C.姓名,SUM(A.实收金额) AS 费用总额" & _
                  " From 门诊费用记录 A,保险帐户 B,病人信息 C" & _
                  " Where A.病人ID=B.病人ID ANd B.病人ID=C.病人ID And B.险类=" & intinsure & _
                  " And C.病人ID=[1] And A.结帐ID=[2]" & _
                  " Group By A.NO,B.医保号,C.姓名"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", lng病人ID, lng结帐ID)
    
    gstrSQL = "zl_结算日志_Insert(" & rsTmp!性质 & ",'" & rsTmp!NO & "','" & rsTmp!医保号 & "','" & rsTmp!姓名 & "'," & rsTmp!费用总额 & ")"
    Set cnBalance = New ADODB.Connection
    cnBalance.Open gcnOracle.ConnectionString
    cnBalance.Execute gstrSQL, , adCmdStoredProc
    On Error GoTo 0
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        ClinicSwap = gobjInsure_Obj(mintOrder).ClinicSwap(lng结帐ID, cur个帐支付, cur医保基金, cur全自付, cur先自付, intinsure, strAdvance)
    Else
        Select Case intinsure
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
    '                2.退费结算方式与收费结算方式不一致时，返回格式为：结算方式|金额||结算方式|金额
    '返回：交易成功返回true；否则，返回false
    '编制:刘兴洪
    '日期:2014-10-14 09:40:40
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '调用模块：1121-门诊退费;1124-保险补充结算;1137-门诊结帐作废
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim strSelfNo As String, strSelfPwd As String, str卡号 As String
    Dim rsTmp As New ADODB.Recordset, int结算性质 As Integer
    Dim curMoney As Currency
    Dim lng病人ID As Long
    Dim lng冲销ID As Long
    Dim varData As Variant
    Dim bln补充结算 As Boolean
    Dim strNO As String
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
    If rsTmp.RecordCount = 0 Then
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
            Case Is > 900
                ClinicDelSwap = 门诊结算冲销_中联(lngStlID, curMoney, lng病人ID, intinsure)
        End Select
    End If
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
    
    
    If cap业务 = support门诊预算 Then
        Select Case intinsure
            Case Is > 900
                gstrSQL = "select 参数值  from 保险参数 where 险类=[1] and 中心 is null and 参数名=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取保险参数", intinsure, "收费使用医保基金")
                If rsTmp.EOF = False Then
                    GetCapability = IIf(rsTmp("参数值") = 1, True, False)
                End If
        End Select
    End If
        
    
    If intinsure > 900 Then
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
    gstrSQL = " Select A.大类ID,A.项目编码,A.项目名称,A.是否医保 as 项目是否医保,B.*" & _
                    " From 保险支付项目 A,(Select * From 保险支付大类 Where 险类=[1]) B" & _
                    " Where A.大类ID=B.ID(+) And A.险类=[1] And A.收费细目ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取费用数据", intinsure, 收费细目ID)
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then Exit Function
    
    If Not rsTmp.EOF Then
        bln保险项目 = True
        lng大类ID = IIf(IsNull(rsTmp!大类id), 0, rsTmp!大类id)
        str项目编码 = IIf(IsNull(rsTmp!项目编码), "", rsTmp!项目编码)
        If lng大类ID <> 0 Then
            bln保险项目 = bln保险项目 And (IIf(IsNull(rsTmp!是否医保), 0, rsTmp!是否医保) = 1)
            
            '如果服务对象不对，也认为它不是医保项目
            lng服务对象 = IIf(IsNull(rsTmp("服务对象")), 3, rsTmp("服务对象")) '缺省认为该大类可同时服务于门诊与住院病人
            If bln门诊 = True Then
                If lng服务对象 <> 1 And lng服务对象 <> 3 Then bln保险项目 = False
            Else
                If lng服务对象 <> 2 And lng服务对象 <> 3 Then bln保险项目 = False
            End If
        End If
        bln保险项目 = bln保险项目 And (IIf(IsNull(rsTmp!项目是否医保), 0, rsTmp!项目是否医保) = 1)
    End If
    
        GetItemInsure = IIf(bln保险项目, 1, 0) & ";" & lng大类ID & ";" & cur统筹金额 & ";" & str项目编码
    
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        GetItemInsure = gobjInsure_Obj(mintOrder).GetItemInsure(病人ID, 收费细目ID, 金额, bln门诊, intinsure, strAdvance)
    Else
        Select Case intinsure
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
    
    If gclsInsure.InitInsure(gcnOracle, intinsure) = False Then TranChargeDetail = True: Exit Function
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    If IsApartComponents(intinsure) Then
        If Not CreateObject_Insure(intinsure, mintOrder, 1) Then Exit Function
        TranChargeDetail = gobjInsure_Obj(mintOrder).TranChargeDetail(int类别, str单据号, int性质, int状态, str消息, lng病人ID, intinsure, strAdvance)
    Else
        Select Case intinsure
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
        If GetCapability(support挂号使用个人帐户, 0, intinsure) Then
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
                    blnMsg = True
                End If
            Else
                blnMsg = True
            End If
        Else
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
    '94352
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


