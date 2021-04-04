Attribute VB_Name = "mdlCISJob"
Option Explicit
Public gblnShowInTaskBar As Boolean         '是否显示窗体在任务条上
Public gobjRichEPR As New cRichEPR          '病历核心部件
Public gobjKernel As New zlPublicAdvice.clsPublicAdvice       '临床核心部件
Public gobjPath As New zlPublicPath.clsPublicPath             '临床路径部件
Public gobjRegist As Object
Public gobjCommunity As Object              '社区档案接口对象
Public gclsInsure As New clsInsure          '医保变量
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                  '当前用户具有的当前模块的功能
Public gcolPrivs As Collection              '记录内部模块的权限
Public gstrSysName As String                '系统名称
Public glngSys As Long
Public glngModul As Long
Public gstrDBUser As String                 '当前数据库用户
Public gstrUnitName As String               '用户单位名称
Public gstrProductName As String            'OEM产品名称
Public gfrmMain As Object                   '导航台窗体
Public gblnOK As Boolean
Public gobjCISBase As Object                '护士站或医技站调用诊疗收费对照
Public gobjPlugIn As Object                 '外挂功能对象
Public gobjRis As Object                    'RIS接口部件
Public gblnKSSStrict As Boolean             '抗菌药物严格控制
Public gblnKSSAuditType As Boolean             '抗菌药物审核方式，参数：按医疗小组进行抗菌药物审核 0－默认，1－按医疗小组
Public gbln手术分级管理 As Boolean  '是否启用手术分级管理
Public gbln输血分级管理 As Boolean  '是否启用输血分级管理
Public gbln血库系统 As Boolean  '是否安装血库系统
Public gobjEmr  As Object                   '新版病历部件
Public gbln允许超过挂号有效天数 As Boolean   '允许处理超过挂号有效天数的病人
Public gobjLIS As Object                    'LIS申请部件
Public gobjPublicPacs As Object                  'PACS公共部件
Public gobjExchange As Object               'HL7数据交换部件
Public gobjPublicExpense As Object           '费用公共部件
Public gobjNurseIntegrate As Object        '整体护理接口部件
Public gobjPublicBlood As Object             '血库公共部件
Public gblnGetPath As Boolean

'电子签名
Public gintCA As Integer '电子签名认证中心
Public gstrESign As String '电子签名控制场合
Public grsSign As Recordset  '电子签名启用部门

Public gbln输血申请三级审核 As Boolean  '输血申请三级审核
'合理用药接口类型,0-未使用,1-美康,2-大通,3-太元通
Public gbytPass As Byte
'0-医生选择，1-按药品目录输入，2-按过敏源输入
Public gint过敏输入来源 As Integer
'太元通接口对象
Public gobjPass As Object

Public glngPreHWnd As Long '用于支持鼠标滚轮功能

Public Enum 调用场合
    E门诊调用 = 1
    E住院调用 = 2
End Enum

'系统参数
Public gstrLike As String   '如果是双向匹配，则为%
Public gint简码 As Integer  '简码匹配方式：0-拼音,1-五笔
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"

Public gbytCardLen As Byte '就诊卡号长度
Public gblnCardHide As Boolean '就诊卡号密文显示

Public gbytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
Public gint普通挂号天数 As Integer '普通挂号单有效天数
Public gint急诊挂号天数 As Integer '急诊挂号单有效天数

Public gdbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
                                                      '为负数(-N)时表示,N元内免密支付,表示病人在消费N元内必须刷卡,不必输入密码即可支付;否则必须输入密码
Public gbln执行前先结算 As Boolean '门诊一卡通,项目执行前必须先收费或先记帐审核

Public gstr医嘱核对 As String    '输血皮试医嘱需要核对 按位存取11，第一位为 输血医嘱，第二位为 皮试医嘱
Public gstr输液配置中心 As String          '空-不启用；否则启用
Public gblnDo As Boolean  '是否使用个性化设置
Public gint医嘱执行有效天数 As Integer '允许修改n天内登记的医嘱执行记录

Public gint诊断来源 As Integer '1-由医生选择输入来源,2-按照诊断标准输入,3-按照疾病编码输入
Public gstr诊断输入 As String '1门诊/2住院：1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Public gbln启用影像信息系统接口 As Boolean
Public gbln启用影像信息系统预约 As Boolean
Public gbln启用整体护理接口 As Boolean
Public gbln挂号按排 As Boolean '根据系统参数：挂号排班模式   true新版，false老版
Public gblnPatiByID As Boolean '系统参数：同一身份证只能对应一个建档病人


'内部应用模块号定义
Public Enum Enum_Inside_Program
    p电子病历管理 = 2250
    p新版住院病历 = 2252
    p新版门诊病历 = 2251
    p疾病报告填写 = 1249
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p临床路径应用 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p电子病案查阅 = 1259
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    P新版护士站 = 1265
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    p观片工具管理 = 1289
    p病人入出 = 1132
    p住院记帐 = 1133
    p费用查询 = 1139
    p门诊分诊管理 = 1113
    p排队叫号虚拟模块 = 1160
    p抗菌用药审核 = 1266
    p手术审核管理 = 1267
    p电子病案审查 = 1560
    p输血审核管理 = 1268
    p手麻接口 = 2425
    p手术授权管理 = 1080
    p输液配置中心 = 1345
    P门诊路径应用 = 1248
    P病案查阅打印 = 1566
End Enum

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门名 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    用药级别 As Long
End Type
Public UserInfo As TYPE_USER_INFO

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.编号 = rsTmp!编号
            UserInfo.部门ID = Nvl(rsTmp!部门ID, 0)
            UserInfo.部门名 = Nvl(rsTmp!部门名)
            UserInfo.简码 = Nvl(rsTmp!简码)
            UserInfo.姓名 = Nvl(rsTmp!姓名)
            UserInfo.用户名 = rsTmp!User & ""
            GetUserInfo = True
        End If
    End If
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean, Optional ByVal lngSys As Long) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
'      lngSys=指定系统的内部模块权限，传0或不传是默认是当前系统
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    If lngSys = 0 Then lngSys = glngSys
    On Error Resume Next
    strPrivs = gcolPrivs(lngSys & "_" & lngProg)
    If err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove lngSys & "_" & lngProg
        End If
    Else
        err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(lngSys, lngProg)
        gcolPrivs.Add strPrivs, lngSys & "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    '55928:刘鹏飞,2012-11-20
    gblnDo = Val(zlDatabase.GetPara("使用个性化风格")) <> 0
    
    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    gint简码 = Val(zlDatabase.GetPara("简码方式"))
    
    '费用金额小数点位数
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '就诊卡号密文显示
    strSQL = "Select 卡号长度, Nvl(卡号密文, 0) 卡号密文 From 医疗卡类别 Where 特定项目 = '就诊卡'"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "就诊卡")
    If rsTmp.RecordCount > 0 Then
        gblnCardHide = rsTmp!卡号密文 <> "0"
        gbytCardLen = Val("" & rsTmp!卡号长度)
    Else
        gblnCardHide = False
        gbytCardLen = 8
    End If
    
    
    '挂号有效天数
    strTmp = zlDatabase.GetPara(21, glngSys)
    If Len(strTmp) = 1 Then strTmp = strTmp & strTmp
    gint普通挂号天数 = Val(Mid(strTmp, 1, 1))
    gint急诊挂号天数 = Val(Mid(strTmp, 2, 1))
    
    '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    '一卡通消费验证
    strTmp = zlDatabase.GetPara(28, glngSys) & "|"
    gdbl预存款消费验卡 = Val(Split(strTmp, "|")(0))
    
    
    '门诊一卡通,项目执行前必须先收费或先记帐审核
    gbln执行前先结算 = Val(zlDatabase.GetPara(163, glngSys)) <> 0
    
    '电子签名认证中心
    gintCA = Val(zlDatabase.GetPara(25, glngSys))
    
    '电子签名控制场合
    gstrESign = zlDatabase.GetPara(26, glngSys)
    
    '读取部门启用数据
    If glngModul = p门诊医生站 Or glngModul = p住院医生站 Or glngModul = p住院护士站 Or glngModul = p医技工作站 Or _
        glngModul = P新版护士站 Or glngModul = p抗菌用药审核 Then
        '读取部门启用数据
        Set grsSign = New ADODB.Recordset
        grsSign.Fields.Append "部门ID", adBigInt
        grsSign.Fields.Append "场合", adBigInt
        grsSign.Fields.Append "是否启用", adBigInt
        grsSign.CursorLocation = adUseClient
        grsSign.LockType = adLockOptimistic
        grsSign.CursorType = adOpenStatic
        grsSign.Open
    End If
    
    '输血和皮试医嘱执行后需要核对
    gstr医嘱核对 = zlDatabase.GetPara(186, glngSys)
    
    '抗菌药物分级管理
    gblnKSSStrict = Val(zlDatabase.GetPara(187, glngSys)) <> 0
    
    '按医疗小组进行抗菌药物审核
    gblnKSSAuditType = Val(zlDatabase.GetPara(248, glngSys)) <> 0
    
    '是否启用手术分级管理
    gbln手术分级管理 = Val(zlDatabase.GetPara(209, glngSys)) <> 0
    
    '是否启用输血分级管理
    gbln输血分级管理 = Val(zlDatabase.GetPara(216, glngSys)) <> 0
    
    '是否安装血库系统
    gbln血库系统 = (IsSysSetUp(2200) And Val(zlDatabase.GetPara(236, glngSys)) <> 0)
    
    '允许处理超过挂号有效天数的病人
    gbln允许超过挂号有效天数 = Val(zlDatabase.GetPara(210, glngSys)) <> 0
    
    '61762:刘鹏飞,2012-05-20
    gstr输液配置中心 = Get输液配置中心

    '输血申请三级审核
    gbln输血申请三级审核 = Val(zlDatabase.GetPara(218, glngSys)) <> 0
    
    '允许修改n天内登记的医嘱执行记录
    gint医嘱执行有效天数 = Val(zlDatabase.GetPara(220, glngSys))
    '合理用药接口类型，0-未启用，1-四川美通，2-大通，3-太元通
    gbytPass = Val(zlDatabase.GetPara(30, glngSys))
    
    '过敏输入来源控制
    gint过敏输入来源 = Val(zlDatabase.GetPara(224, glngSys))
    
    '诊断输入来源
    gint诊断来源 = Val(zlDatabase.GetPara(55, glngSys, , 1))
    
    '诊断输入方式
    gstr诊断输入 = zlDatabase.GetPara(65, glngSys, , "11")
    
    gbln启用影像信息系统接口 = Val(zlDatabase.GetPara(255, glngSys)) = 1
    
    gblnGetPath = Val(zlDatabase.GetPara(54, glngSys, glngModul)) = 1
    
    strTmp = ""
    gbln挂号按排 = False
    strTmp = zlDatabase.GetPara(256, glngSys) & "|" '系统参数：挂号排班模式
    If 0 <> Val(Split(strTmp, "|")(0)) Then
        If Split(strTmp, "|")(1) <> "" Then
            strTmp = Format(Split(strTmp, "|")(1), "YYYY-MM-DD")
            If Format(zlDatabase.Currentdate, "YYYY-MM-DD") >= strTmp Then
                gbln挂号按排 = True
            End If
        End If
    End If
    
    '同一身份证只能对应一个建档病人
    gblnPatiByID = Val(zlDatabase.GetPara(279, glngSys)) = 1

    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get挂号ID(ByVal strNO As String) As Long
'功能：根据挂号单获取挂号ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From 病人挂号记录 Where NO=[1] And 记录性质=1 And 记录状态=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get挂号ID", strNO)
    If Not rsTmp.EOF Then Get挂号ID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicePause(ByVal lng医嘱ID As Long) As String
'功能：获取指定医嘱的暂停时间段记录
'返回："暂停时间,开始时间;...."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    strSQL = "Select 操作类型,操作时间 From 病人医嘱状态" & _
        " Where 操作类型 IN(6,7) And 医嘱ID=[1]" & _
        " Order by 操作时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!操作类型 = 6 Then
            strTmp = strTmp & ";" & Format(rsTmp!操作时间, "yyyy-MM-dd HH:mm:ss") & ","
        ElseIf rsTmp!操作类型 = 7 Then
            '启用的那一秒不在暂停的范围之内
            strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!操作时间), "yyyy-MM-dd HH:mm:ss")
        End If
        rsTmp.MoveNext
    Next
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDiagnose(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal int来源 As Integer) As String
'功能：读取病人指定次就诊的门诊诊断
'参数：lng就诊ID=挂号ID或主页ID
'      int来源=1-门诊,2-住院
'返回：用"，"号分隔的多个诊断串
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 记录来源,诊断类型,诊断次序,诊断描述,是否疑诊,Mod(诊断类型,10) as 大类 From 病人诊断记录" & _
        " Where 病人ID=[1] And 主页ID=[2] And NVL(编码序号,1) = 1 And 诊断类型 IN(" & IIf(int来源 = 1, "1,11", "1,2,3,11,12,13") & ")" & _
        " Order by 记录来源,诊断类型,诊断次序"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiDiagnose", lng病人ID, lng就诊ID)
    
    '先按来源优先顺序过滤
    rsTmp.Filter = "记录来源=3" '首页整理
    If rsTmp.EOF Then rsTmp.Filter = "记录来源=2" '入院登记
    If rsTmp.EOF Then rsTmp.Filter = "记录来源=1" '病历
    If rsTmp.EOF Then rsTmp.Filter = "记录来源=4" '病案室录入
    
    '住院再按类型优先顺序过滤
    If Not rsTmp.EOF And int来源 = 2 Then
        strSQL = rsTmp.Filter
        rsTmp.Filter = strSQL & " And 大类=3"
        If rsTmp.EOF Then rsTmp.Filter = strSQL & " And 大类=2"
        If rsTmp.EOF Then rsTmp.Filter = strSQL & " And 大类=1"
    End If
    
    strSQL = ""
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!诊断描述) Then
            strSQL = strSQL & "，" & rsTmp!诊断描述 & IIf(Nvl(rsTmp!是否疑诊, 0) = 1, "（？）", "")
        End If
        rsTmp.MoveNext
    Loop
    
    GetPatiDiagnose = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadAdviceSignSource(ByVal int操作类型 As Integer, _
    ByVal lng病人ID As Long, ByVal varTime As Variant, strIDs As String, _
    ByVal lng签名ID As Long, ByVal blnMoved As Boolean, strSource As String, _
    Optional ByVal lng前提ID As Long, Optional ByVal colSomeTime As Collection) As Integer
'功能：获取病人用于电子签名/验证的医嘱源文内容
'参数：
'  int操作类型=要签名/验证签名的医嘱状态
'  签名时传入：
'    lng病人ID
'    varTime=病人挂号单号或主页ID
'    strIDs=指定要签名的医嘱ID序列(组ID)
'    lng前提ID=新开医嘱要签名的医嘱来源(是否医技)
'    colSomeTime=某医嘱的时间数据，如停止医嘱签名时，传入包含医嘱执行终止时间的数据，校对时传入校对时间数据
'  验证签名时：
'    lng签名ID=签名记录的ID
'    blnMoved=是否医嘱数据已转出
'返回：签名/验证签名的源文生成规则
'      strIDs=签名/验证签名的医嘱ID序列(每个明细ID)
'      strSource=签名/验证签名的医嘱源文
    Dim rsTmp As New ADODB.Recordset
    Dim str组IDs As String, strSQL As String, i As Long
    Dim arrField As Variant, strField As String
    Dim strLine As String, intRule As Integer
    
    On Error GoTo errH
    
    str组IDs = strIDs
    strSource = "": strIDs = ""
    intRule = 1 '这是最新的医嘱签名源文生成规则编号
    
    If lng签名ID = 0 Then
        '签名时
        If int操作类型 = 1 Then
            '对新开的医嘱进行签名：本次就诊/住院当前医生新下达的未签名医嘱
            strSQL = _
                " Select /*+ Rule*/ A.* From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And B.签名ID is Null And B.操作类型=1" & _
                " And A.医嘱状态=1 And Nvl(A.前提ID,0)=[5]" & _
                " And Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))=[3]" & _
                " And Exists(Select M.姓名 From 人员表 M,执业类别 N" & _
                "       Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
                "         And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')" & _
                "   )" & _
                IIf(TypeName(varTime) = "String", " And A.病人ID+0=[1] And A.挂号单=[2]", " And A.病人ID=[1] And A.主页ID=[2]") & _
                IIf(str组IDs <> "", " And Nvl(A.相关ID,A.ID) IN(Select Column_Value From Table(f_Num2list([4])))", "") & _
                " Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, varTime, UserInfo.姓名, str组IDs, lng前提ID)
        Else
            '对要作废、停止、校对的医嘱进行签名：新开时签了名的指定医嘱，不一定是当前医生下达
            strSQL = _
                " Select /*+ Rule*/ A.* From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And B.签名ID is Not Null And B.操作类型=1" & _
                IIf(TypeName(varTime) = "String", " And A.病人ID+0=[1] And A.挂号单=[2]", " And A.病人ID=[1] And A.主页ID=[2]") & _
                IIf(str组IDs <> "", " And Nvl(A.相关ID,A.ID) IN(Select Column_Value From Table(f_Num2list([3])))", "") & _
                " Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, varTime, str组IDs)
        End If
    Else
        '验证签名时:先读取签名时的源文生成规则
        strSQL = "Select 签名规则 From 医嘱签名记录 Where ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "医嘱签名记录", "H医嘱签名记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng签名ID)
        If Not rsTmp.EOF Then intRule = Nvl(rsTmp!签名规则, 1)
        '--
        strSQL = _
            " Select A.* From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And B.签名ID=[1] Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng签名ID)
    End If
    
    '医嘱源文的不同生成规则
    If intRule = 1 Then
        If int操作类型 = 3 Then
            strField = "ID,相关ID,姓名,性别,年龄,婴儿,医嘱期效,开始执行时间,医嘱内容,标本部位,单次用量,总给予量," & _
                "医生嘱托,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,校对时间,执行性质,紧急标志,开嘱医生,开嘱时间"
        ElseIf int操作类型 = 8 Then
            strField = "ID,相关ID,姓名,性别,年龄,婴儿,医嘱期效,开始执行时间,医嘱内容,标本部位,单次用量,总给予量," & _
                "医生嘱托,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,执行终止时间,执行性质,紧急标志,开嘱医生,开嘱时间"
        Else
            strField = "ID,相关ID,姓名,性别,年龄,婴儿,医嘱期效,开始执行时间,医嘱内容,标本部位,单次用量,总给予量," & _
                "医生嘱托,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,执行性质,紧急标志,开嘱医生,开嘱时间"
        End If
    End If
    arrField = Split(strField, ",")
        
    '生成医嘱签名源文
    Do While Not rsTmp.EOF
        strLine = ""
        For i = 0 To UBound(arrField)
            If lng签名ID = 0 And int操作类型 = 3 And arrField(i) = "校对时间" Then
                '校对医嘱签名时,对校对时间特殊处理：由于是在执行过程之前取签名源文,这时还未写入数据库
                strLine = strLine & vbTab & colSomeTime("_" & Nvl(rsTmp!相关ID, rsTmp!ID))
            ElseIf lng签名ID = 0 And int操作类型 = 8 And arrField(i) = "执行终止时间" Then
                '停止医嘱签名时,对终止时间特殊处理：由于是在执行过程之前取签名源文,这时还未写入数据库
                strLine = strLine & vbTab & colSomeTime("_" & Nvl(rsTmp!相关ID, rsTmp!ID))
            Else
                If IsNull(rsTmp.Fields(arrField(i)).Value) Then
                    strLine = strLine & vbTab & ""
                Else
                    If Rec.IsType(rsTmp.Fields(arrField(i)).Type, adDBTimeStamp) Then
                        strLine = strLine & vbTab & Format(rsTmp.Fields(arrField(i)).Value, "yyyy-MM-dd HH:mm:ss")
                    Else
                        strLine = strLine & vbTab & rsTmp.Fields(arrField(i)).Value
                    End If
                End If
            End If
        Next
        strSource = strSource & vbCrLf & Mid(strLine, 2)
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    strSource = Mid(strSource, 3)
    strIDs = Mid(strIDs, 2)
    
    ReadAdviceSignSource = intRule
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDept(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal bytMode As Byte) As Long
'功能：获取病人当前病区和科室
'参数：bytMode=0-查科室,1=查病区
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select " & IIf(bytMode = 0, "当前科室id", "当前病区id") & " as 科室ID" & vbNewLine & _
            "From 病人信息" & vbNewLine & _
            "Where 病人id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then GetPatiDept = Val("" & rsTmp!科室ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiLog(lng病人ID As Long, lng主页ID As Long) As ADODB.Recordset
'功能：获取病人变动记录
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 终止原因,终止时间,开始原因,Decode(开始原因, 1, '入院入住', 2, '入住', 3," & _
            " Decode(开始时间, Null, '转科', '转科入住'), 4, '换床', 5, '床位等级变动', 6, '护理等级变动', 7," & vbNewLine & _
            "               '经治医师改变', 8, '责任护士改变', 9, '转为住院病人', 10, '预出院', 11, '主治医师变动'," & _
            " 12, '主任医师变动', 13, '病况变动',14,'转医疗小组',15,Decode(开始时间, Null, '转病区', '转病区入住')) 操作" & vbNewLine & _
            "From 病人变动记录" & vbNewLine & _
            "Where Nvl(附加床位, 0) = 0 And 病人id = [1] And 主页id = [2]" & vbNewLine & _
            "Order By 终止时间 Desc, 开始时间 Desc"
    Set GetPatiLog = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPati费用信息(ByVal lng病人ID As Long, lng主页ID As Long) As String
'功能：获取当前病人的费用信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lng病人性质 As Long
    
    strSQL = "Select 病人性质 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPati费用信息", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        lng病人性质 = Val(rsTmp!病人性质 & "")
    End If
    strSQL = _
        " Select 费用余额,预交余额,0 as 预结费用,0 as 担保额 From 病人余额 Where 性质=1 And 病人ID=[1] And 类型 = [3]" & _
        " Union ALL" & _
        " Select 0,0,0, Sum(担保额) as 担保额 From 病人担保记录 Where 病人id = [1] And 主页id = [2] And 删除标志 = 1 And (Sysdate <= 到期时间 Or 到期时间 Is Null)" & _
        " Union ALL" & _
        " Select 0,0,Sum(金额),0 From 保险模拟结算 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2]"
    strSQL = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额,Sum(预结费用) as 预结费用,sum(担保额) as 担保额 From (" & strSQL & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPati费用信息", lng病人ID, lng主页ID, IIf(lng病人性质 = 1, 1, 2))
    If Not rsTmp.EOF Then
        GetPati费用信息 = _
            "预交余额:" & FormatEx(Nvl(rsTmp!预交余额, 0), 2) & ",未结费用:" & FormatEx(Nvl(rsTmp!费用余额, 0), 2) & _
            IIf(Nvl(rsTmp!预结费用, 0) <> 0, ",预结费用:" & FormatEx(Nvl(rsTmp!预结费用, 0), 2), "") & _
            ",剩余款:" & FormatEx(Nvl(rsTmp!预交余额, 0) - Nvl(rsTmp!费用余额, 0) + Nvl(rsTmp!预结费用, 0), 2) & ",担保额:" & Nvl(rsTmp!担保额, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get住院费用药占比(ByVal lng病人ID As Long, lng主页ID As Long) As String
'功能：获取当前病人的住院费用药占比
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select /*+ RULE */" & vbNewLine & _
            " c.姓名 As 名称, Sum(Decode(a.收费类别, '5', a.实收金额, 0)) As 西药费, Sum(Decode(a.收费类别, '6', a.实收金额, 0)) As 成药费," & vbNewLine & _
            " Sum(Decode(a.收费类别, '7', a.实收金额, 0)) As 草药费, Sum(Decode(a.收费类别, '5', 0, '6', 0, '7', 0, a.实收金额)) As 非药费," & vbNewLine & _
            " Sum(a.实收金额) As 所有费" & vbNewLine & _
            "From 住院费用记录 A, Table(f_Num2list2([1])) B, 病人信息 C" & vbNewLine & _
            "Where a.病人id = b.C1 And a.主页id = b.C2 And b.C1 = c.病人id And a.记录状态 <> 0 Having Sum(a.实收金额) > 0" & vbNewLine & _
            "Group By c.姓名" & vbNewLine & _
            "Order By 名称"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get住院费用药占比", lng病人ID & ":" & lng主页ID)
    If Not rsTmp.EOF Then
        Get住院费用药占比 = ",药占比:" & Format((Val(rsTmp!西药费) + Val(rsTmp!成药费) + Val(rsTmp!草药费)) / Val(rsTmp!所有费) * 100, "0.0") & "%"
    Else
        Get住院费用药占比 = ",药占比:0.0%"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Get住院费用药占比 = ",药占比:0.0%"
End Function

Public Function LoadPatiAllergy(ByVal lng病人ID As Long, Optional ByRef objCbo As Object, Optional ByRef rsAller As ADODB.Recordset) As Boolean
'功能：读取病人的过敏记录到下拉框中
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
        
    strSQL = "Select Distinct B.发生时间 as 挂号时间,D.名称 as 挂号科室,C.主页ID,E.名称 as 住院科室," & _
        " A.药物名,Nvl(A.过敏时间,A.记录时间) as 过敏时间,B.NO as 挂号单,A.药物ID,A.过敏源编码,A.过敏反应,(max(lengthB(a.药物名)) over()-lengthB(a.药物名)+4) AS 空格长度" & _
        " From 病人过敏记录 A,病人挂号记录 B,病案主页 C,部门表 D,部门表 E" & _
        " Where A.病人ID=B.病人ID(+) And A.主页ID=B.ID(+) And B.记录性质(+)=1 And B.记录状态(+)=1" & _
        " And A.病人ID=C.病人ID(+) And A.主页ID=C.主页ID(+)" & _
        " And B.执行部门ID=D.ID(+) And C.出院科室ID=E.ID(+)" & _
        " And A.结果=1 And 药物名 is Not NULL And A.病人ID=[1] And Not Exists" & vbNewLine & _
        " (Select 药物id" & vbNewLine & _
        "       From 病人过敏记录" & vbNewLine & _
        "       Where (Nvl(药物id, 0) = Nvl(a.药物id, 0) Or Nvl(药物名, 'Null') = Nvl(a.药物名, 'Null')) And Nvl(结果, 0) = 0 And" & vbNewLine & _
        "             记录时间>A.记录时间 And 病人id = [1])" & _
        " Order by Nvl(A.过敏时间,A.记录时间) Desc"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LoadPatiAllergy", lng病人ID)
        
    If Not objCbo Is Nothing Then
        objCbo.Clear
        Do While Not rsTmp.EOF
'            If Not IsNull(rsTmp!挂号时间) Then
'                strTmp = Format(rsTmp!过敏时间, "yyyy-MM-dd") & "," & Nvl(rsTmp!药物名) & ",门诊就诊:" & Nvl(rsTmp!挂号科室)
'            Else
'                strTmp = Format(rsTmp!过敏时间, "yyyy-MM-dd") & "," & Nvl(rsTmp!药物名) & ",第" & rsTmp!主页ID & "次住院:" & Nvl(rsTmp!住院科室)
'            End If
            strTmp = Nvl(rsTmp!药物名) & String(Val(rsTmp!空格长度), " ") & Format(rsTmp!过敏时间, "yyyy-MM-dd") & String(4, " ")

            If Not IsNull(rsTmp!过敏反应) Then strTmp = strTmp & IIf(Nvl(rsTmp!过敏反应) = "", "", "过敏反应:" & rsTmp!过敏反应)

            objCbo.AddItem strTmp
            
            rsTmp.MoveNext
        Loop
        If objCbo.ListCount = 0 Then
            objCbo.AddItem "无记录"
        End If
        objCbo.ListIndex = 0
        objCbo.ForeColor = vbRed
    End If
    
    If Not rsAller Is Nothing Then Set rsAller = rsTmp
        
    LoadPatiAllergy = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRefuseReason(ByVal lng病人ID As Long, lng主页ID As Long) As String
'功能：获取病人的病案提交拒审理由
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '以该次住院最近一次被拒的为准
    strSQL = "Select 拒审理由 From (Select 拒审理由 From 病案提交记录 Where 病人ID=[1] And 主页ID=[2] And 记录状态=2 Order by ID Desc) Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetRefuseReason", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then GetRefuseReason = Nvl(rsTmp!拒审理由)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiMedRecHaveSubmit(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：检查指定病人的病案是否已经提交(通过提交记录)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '以该次住院最近一次被拒的为准
    strSQL = "Select 1 From 病案提交记录 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiMedRecHaveSubmit", lng病人ID, lng主页ID)
    PatiMedRecHaveSubmit = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadCaseMap(lngID As Long) As StdPicture
'功能：根据标记图ID返回图形对象
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 图形 From 病历标记图 Where 元素ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngID)
        
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!图形) Then Exit Function
    
    On Error GoTo 0
    
    intFile = FreeFile
    strFile = CurDir & "\zlNewPicture" & Timer & ".pic"
    
    Open strFile For Binary As intFile
    
    lngFileSize = rsTmp.Fields("图形").ActualSize
    ReDim arrBin(lngFileSize - 1) As Byte
    arrBin() = rsTmp.Fields("图形").GetChunk(lngFileSize)
    Put intFile, , arrBin()
    Close intFile
    
    Set ReadCaseMap = VB.LoadPicture(strFile)
    Kill strFile
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function XWWebViewerOpen(lngOrderID As Long) As Long
''--------------------------------------------
''功能： 打开RIS的WEB Viewer
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------
    Dim strIp As String
    Dim strUrl As String
    
    On Error GoTo err
    
    strIp = zlDatabase.GetPara("XWWEB服务器IP", glngSys, 1288, "")
    
    If strIp <> "" Then
        strUrl = "C:\Program Files\Internet Explorer\iexplore.exe http://" & strIp & ":8080/imageweb/imageAction.action?ColID0=22&ColValue0=" & lngOrderID
        
        Shell strUrl, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        MsgBox "WEB服务器IP地址为空，请先设置好WEB服务器。", vbOKOnly, "提示信息"
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'功能：创建本地目录
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'功能：当指定目录的大小达到一定百分比时，清空该目录
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

Public Function zlGetLocaleComputerNamePara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDeFault As String, _
        Optional strComputerName As String = "") As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定本机参数
    '入参：varPara=参数号或参数名，以数字或字符类型传入区分
    '      lngSys=使用该参数的系统编号，如100
    '      lngModual=使用该参数的模块号，如1230
    '      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
    '     strComputerName-获取指定机器名参数
    '出参：
    '返回：参数值，字符串形式
    '编制：刘兴洪
    '日期：2010-06-07 13:56:22
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Integer, rsPara As ADODB.Recordset, rsUserPara As ADODB.Recordset
    Dim blnNew As Boolean, blnEnabled As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select ID,Nvl(参数值,缺省值) as 参数值,SYS_CONTEXT('USERENV','TERMINAL') as MName From zlParameters where 模块=[1] and 系统=[2]"
    If TypeName(varPara) = "String" Then
        strSQL = strSQL & " and 参数名=[3]"
    Else
        strSQL = strSQL & " and 参数号=[4]"
    End If
    Set rsPara = zlDatabase.OpenSQLRecord(strSQL, "GetPara", lngModual, lngSys, CStr(varPara), Val(varPara))
    If rsPara.EOF = False Then
        strSQL = _
            "   Select 参数值 " & _
            "   From zlUserParas Where 参数ID=[1]  and  机器名=[2]"
        Set rsUserPara = zlDatabase.OpenSQLRecord(strSQL, "GetPara", Val(Nvl(rsPara!ID)), IIf(strComputerName = "", CStr(Nvl(rsPara!MName)), strComputerName))
         If Not rsUserPara.EOF Then
                zlGetLocaleComputerNamePara = Nvl(rsUserPara!参数值, strDeFault)
         Else
                zlGetLocaleComputerNamePara = Nvl(rsPara!参数值, strDeFault)
         End If
    Else
        zlGetLocaleComputerNamePara = strDeFault
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function CheckDoctorPatisIsValid() As Byte
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查医生就诊人数是否有效
    '返回：0-分诊台分诊呼叫;1-医生主动呼叫;2-先分诊台和善叫,再是医生呼叫
    '编制：刘兴洪
    '日期：2010-06-07 14:32:47
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean, strComputerName As String

    '刘兴洪:应用于排队叫号的呼叫人次:需要配合分诊台模块的排队叫号模式为１并且有排队呼叫站点=1时有效
     
     '需要检查是否为医生主动呼叫方式
     '排队叫号处理模式:1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
     blnValid = Val(zlDatabase.GetPara("排队叫号模式", glngSys, p门诊分诊管理)) = 1
    If blnValid Then
         '还需要检查:排队呼叫站点=1
         '排队呼叫站点: 0-代表分诊台分诊呼叫;1-代表医生主动呼叫
         strComputerName = zlDatabase.GetPara("远端呼叫站点", glngSys, p排队叫号虚拟模块)
        blnValid = Val(zlGetLocaleComputerNamePara("排队呼叫站点", glngSys, p门诊分诊管理, "0", strComputerName)) = 1
    End If
    CheckDoctorPatisIsValid = blnValid
End Function

Public Sub PrintInMedRec(ByRef objClsMedRec As zlMedRecPage.clsInOutMedRec, ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
        ByRef objReport As Object, ByVal lng科室ID As Long, ByRef objForm As Object, Optional intPage As Integer, Optional strPDFFile As String, _
          Optional ByRef objReportForm As Object)
'功能：首页打印，预览
'参数：intType=2（打印），=1（预览）0=设置,4-PDF;5-返回嵌入窗体对象
'     mobjReport-打印部件，lng科室ID-病人科室，mobjForm-主窗口
'     intPage=1-4打印的页数（格式）=5打印正面+附页1，=6打印反面+附页2
'     strPDFFile-intType=4 时 PDF输出路径; intType=2 时 为打印机
'     objReportForm-返回嵌入窗体的报表对象
'    If lng病人ID <> 0 Then
'        If objClsMedRec Is Nothing Then
'            Set objClsMedRec = New clsInOutMedRec
'            Call objClsMedRec.InitMedRec(gcnOracle, glngSys, gobjCommunity, gclsInsure)
'        End If
'        Call objClsMedRec.PrintOrPriviewInMedRec(intType, lng病人ID, lng主页ID, objReport, lng科室ID, objForm, intPage)
'    End If
'    Exit Sub
    Dim strName As String
    Dim lngPage As Long
    
    If lng病人ID <> 0 Then
        If objReport Is Nothing Then Set objReport = New clsReport
        Select Case Val(zlDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0"))
    
            Case 0 '卫生部标准
                If Sys.DeptHaveProperty(lng科室ID, "中医科") Then
                    strName = "ZL1_INSIDE_1261_4"
                Else
                    strName = "ZL1_INSIDE_1261_1"
                End If
            Case 1    '四川省标准
                If Sys.DeptHaveProperty(lng科室ID, "中医科") Then
                    strName = "ZL1_INSIDE_1261_6"
                Else
                    strName = "ZL1_INSIDE_1261_5"
                End If
            Case 2    '云南省标准
                If Sys.DeptHaveProperty(lng科室ID, "中医科") Then
                    strName = "ZL1_INSIDE_1261_8"
                Else
                    strName = "ZL1_INSIDE_1261_7"
                End If
            Case 3    '湖南省标准
                If Sys.DeptHaveProperty(lng科室ID, "中医科") Then
                    strName = "ZL1_INSIDE_1261_10"
                Else
                    strName = "ZL1_INSIDE_1261_9"
                End If
        End Select
        If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 0) = 0 And intPage = 0 Then
            Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 1)
        End If
        If intType = 0 Then
            Call ReportPrintSet(gcnOracle, glngSys, strName, objForm)
        Else
            If intPage = 5 Then
                lngPage = 1
            ElseIf intPage = 6 Then
                lngPage = 2
            Else
                lngPage = intPage
            End If
            If intType = 4 Then
                Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), "PDF=" & strPDFFile, intType)
                If intPage > 4 Then
                    Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), "PDF=" & strPDFFile, intType)
                End If
            ElseIf intType = 5 Then
                If strPDFFile <> "" Then
                    Call objReport.SetReportPrintSet(gcnOracle, glngSys, strName, "printer", strPDFFile) '设置指定打印机
                End If
                Call objReport.LoadReport(gcnOracle, glngSys, strName, objForm, objReportForm, Nothing, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), 1)
                If intPage > 4 Then
                    Call objReport.LoadReport(gcnOracle, glngSys, strName, objForm, objReportForm, Nothing, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), 1)
                End If
            Else
                If intType = 2 And strPDFFile <> "" Then
                    Call objReport.SetReportPrintSet(gcnOracle, glngSys, strName, "printer", strPDFFile) '设置指定打印机
                End If
                Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), intType)
                If intPage > 4 Then
                    Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), intType)
                End If
            End If
        End If
    End If
End Sub

Public Function CheckDiseaseFile(ByRef frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intCurDeptID As Long, _
ByVal str疾病IDs As String, ByVal str诊断IDs As String, Optional ByRef lngFileID As Long, Optional ByVal blnOnlyCheck As Boolean, Optional ByRef blnNo As Boolean, Optional ByVal lngFrom As Long = 1) As Boolean

'功能：检查病人哪些疾病证明报告没有填写并提示进行填写
'参数:frmParent    父窗体
'     lng病人ID    病人ID
'     lng主页ID    门诊传挂号ID，住院传主页ID
'     intCurDeptID 书写病历科室ID
'     lng医嘱ID    医嘱ID（用于检查报告）
'     blnOnlyCheck true-只检查未书写病历不弹出病历列表,false-如果有未书写病历则弹出列表
'     blnNo        是否要填写传染病报告卡
'     lngFrom      来源，1-门诊；2-住院
   Dim rsTmp As ADODB.Recordset
   
   On Error GoTo errH
   
    If str疾病IDs = "" And str诊断IDs = "" Then
        CheckDiseaseFile = True
        Exit Function
    End If
    Dim strSQL As String
    If str疾病IDs <> "" Then
        strSQL = strSQL & " Union Select 文件ID From 疾病报告前提 Where 疾病ID IN (Select Column_Value From Table(f_Num2list([3])))"
    End If
    If str诊断IDs <> "" Then
        strSQL = strSQL & " Union Select 文件ID From 疾病报告前提 Where 诊断ID IN (Select Column_Value From Table(f_Num2list([4])))"
    End If
    On Error GoTo errH
    strSQL = "(" & Mid(strSQL, 8) & ") Minus Select 文件ID From 电子病历记录 Where 病人ID=[1] And 主页ID=[2] And 病历种类=5"
    strSQL = "Select /*+ Rule*/" & vbNewLine & _
            " a.Id, a.种类, a.编号, a.名称, a.保留, a.说明" & vbNewLine & _
            "From 病历文件列表 A ,(" & strSQL & ") B Where A.ID=B.文件ID  And nvl(A.保留,0)=0 And " & vbNewLine & _
            "(a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 病历应用科室 C Where c.文件id = a.Id And c.科室id = [5]))" & vbNewLine & _
            "Order By a.编号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDiseaseFile", lng病人ID, lng主页ID, str疾病IDs, str诊断IDs, intCurDeptID)
    blnNo = False
    
    If lngFrom = 1 Then
        If InStr(";" & GetPrivFunc(glngSys, p门诊病历管理) & ";", ";病历书写;") <= 0 Then
            rsTmp.Filter = "保留=4"
        End If
    ElseIf lngFrom = 2 Then
        If InStr(";" & GetPrivFunc(glngSys, p住院病历管理) & ";", ";病历书写;") <= 0 Then
            rsTmp.Filter = "保留=4"
        End If
    End If
        
    If rsTmp.RecordCount = 0 Then
        CheckDiseaseFile = True
        Exit Function
    Else
        strSQL = ""
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            strSQL = strSQL & vbCrLf & "【" & rsTmp!名称 & "】"
            rsTmp.MoveNext
        Loop
    End If

    If rsTmp.RecordCount = 1 Then
        If blnOnlyCheck Then
            If MsgBox("根据病人的诊断信息，以下疾病证明报告还没有填写：" & vbCrLf & vbCrLf & Mid(strSQL, 3) & vbCrLf & vbCrLf & "要继续吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnNo = True: Exit Function
        Else
            If MsgBox("根据病人的诊断信息，以下疾病证明报告还没有填写：" & vbCrLf & vbCrLf & Mid(strSQL, 3) & vbCrLf & vbCrLf & "要继续吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                rsTmp.MoveFirst
                lngFileID = Val(rsTmp!ID & "")
            Else
                blnNo = True
            End If
        End If
    ElseIf rsTmp.RecordCount > 1 Then
        If blnOnlyCheck Then
            If MsgBox("根据病人的诊断信息，以下疾病证明报告还没有填写：" & vbCrLf & vbCrLf & Mid(strSQL, 3) & vbCrLf & vbCrLf & "要继续吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnNo = True: Exit Function
        Else
            If MsgBox("根据病人的诊断信息，以下疾病证明报告还没有填写：" & vbCrLf & vbCrLf & Mid(strSQL, 3) & vbCrLf & vbCrLf & "要继续吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                CheckDiseaseFile = True
                                Exit Function
            Else
                blnNo = True
            End If
        End If
    End If
    CheckDiseaseFile = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub Set临床自管药(objFrom As Object)
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "诊疗基础部件(ZLCISBase)没有正确安装，该功能无法执行。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.SetMedList(objFrom, gcnOracle, glngSys, gstrDBUser)
End Sub

Public Function CheckMecRed(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strfrmCation As String, Optional ByVal strOperateName As String) As Boolean
'功能：检查病案是否已经编目,病案是否在待审查或在审查中(此时首页处于锁定状态，不允许修改)
'       lng病人ID:当前病人ID
'       lng主页ID:当前病人主页ID
'       strfrmCation:调用该函数的窗体名称
'       strOperateName:调用该函数的操作名称。strOperateName为空时，不弹出提示
    Dim strSQL As String, rsTmp As Recordset
    Dim int病案状态 As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    '获取病案状态
    strSQL = "Select Nvl(病案状态, 0) 病案状态 From 病案主页 Where 病人id = [1] And 主页id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, strfrmCation, lng病人ID, lng主页ID)
    rsTmp.MoveFirst
    int病案状态 = rsTmp!病案状态
    '首页锁定与否的判断
    Select Case int病案状态
        Case 1 '等待审查
            strMsg = "该病案等待审查中,不能"
        Case 3 '正在审查
            strMsg = "该病案正在审查中,不能"
        Case 5 '审查归档
            strMsg = "该病案已经审查归档,不能"
        Case 10 '接收待审
            strMsg = "该病案在接收待审中,不能"
        Case Else '2-拒绝审查4-审查反馈;6-审查整改;13-正在抽查;14-抽查反馈;16-抽查整改
            strMsg = ""
    End Select
    
    If strMsg = "" Then
        strSQL = "Select 编目日期 from 病案主页 where 病人ID=[1] And 主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, strfrmCation, lng病人ID, lng主页ID)
        If Not IsNull(rsTmp!编目日期) Then
            strMsg = "该病人的病案已经编目，不能"
        End If
    End If
    
    If strMsg <> "" Then  '锁定首页
        If strOperateName <> "" Then
            MsgBox strMsg & strOperateName & "！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    CheckMecRed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CanUnExec(ByVal datExec As Date, Optional ByVal datNow As Date) As Boolean
'功能：根据执行记录的执行时间判断能否取消执行或取消完成
'参数：datExec=执行记录的执行时间
'      datNow =当前时间
'返回：CanUnExec=true-可以取消执行，false-不可以取消执行

    Dim lngDatDiff As Long
    If datExec <> CDate(Format("0", "yyyy-MM-dd HH:mm")) Then
        If datNow = CDate(0) Then
            datNow = zlDatabase.Currentdate
        End If
        lngDatDiff = DateDiff("D", datExec, datNow)
        CanUnExec = lngDatDiff <= gint医嘱执行有效天数
    Else
        CanUnExec = True
    End If
    
End Function

Public Function GetPatiDiagnoseByDept(ByVal lng部门ID As Long, Optional ByVal intType As Integer = 1) As ADODB.Recordset
'功能：获取指定部门在院病人的所有诊断类型
'参数：
'      lng部门id=病区id/科室id
'      intType=0-按科室显示，1-按病区显示,默认按病区显示
'返回：记录集
    Dim strSQL As String
    
    strSQL = " Select A.病人ID,A.诊断类型, A.诊断描述" & _
        " From 病人诊断记录 A,病案主页 B,病人信息 C,在院病人 R" & _
        " Where a.诊断类型 In (1, 2, 3, 11, 12, 13) And NVL(A.编码序号,1) = 1 And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.病人ID=C.病人ID And C.主页ID=B.主页ID And C.病人ID=R.病人ID And C.当前病区ID=R.病区ID " & _
        " And 诊断次序=1 And" & IIf(intType = 1, " (R.病区ID=[1] Or b.婴儿病区ID=[1])", " (r.科室id = [1] Or b.婴儿科室id = [1])") & _
        " Order by A.病人ID asc,A.记录来源 desc,A.诊断类型 desc"
    On Error GoTo errH
    Set GetPatiDiagnoseByDept = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng部门ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub

Public Function ISPassShowCard() As Boolean
'功能：是否密文显示就诊卡号
'返回:True 密文显示,False 非密文
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnPassShowCard As Boolean
    
    On Error GoTo errHandle
    strSQL = "Select 卡号密文 From 医疗卡类别 where 名称='就诊卡' and 是否固定=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗卡类别")
    If Not rsTemp.EOF Then
        blnPassShowCard = Nvl(rsTemp!卡号密文) <> ""
    End If
    
    ISPassShowCard = blnPassShowCard
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ReadPatPricture(ByVal lng病人ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '参数：lng病人ID=读取指定病人的照片
    '           imgPatient=照片加载位置
    '           strFile=照片的本地路径
    '74421,刘鹏飞,2014-07-04,读取病人照片信息
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = Sys.Readlob(glngSys, 27, lng病人ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'功能：支持滚轮的滚动
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '向下滚
            zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '向上滚
            zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
End Function

Public Function Get病人医嘱打印(ByVal lng病人ID As Long, ByVal lng主页ID) As Integer
'功能：判断某个病人的医嘱是否打印。
'返回：0-未打印、1-部分打印；2-全部打印
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lng医嘱ID As Long
    Dim dat重整 As Date
    Dim bytPrint As Byte
    Dim blnDo As Boolean
    Dim arr婴儿 As Variant
    Dim str婴儿 As String
    Dim lngPrintType As Long
    Dim blnKey As Boolean
    Dim lng序号 As Long
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = "select count(1) as 打印 from 病人医嘱打印 a where a.病人id=[1] and a.主页id=[2] and a.打印时间 is not null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        If (rsTmp!打印 & "") = 0 Then
            Get病人医嘱打印 = 0
            Exit Function
        End If
    End If
    
    strSQL = "select 1 from 病人医嘱打印 a where a.病人id=[1] and a.主页id=[2] and a.打印时间 is not null and Exists" & _
            " (select 1 from 病人医嘱打印 where 病人id=[1] and 主页id=[2] and 打印时间 is null and rownum<2) and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        Get病人医嘱打印 = 1
        Exit Function
    End If
   
    Get病人医嘱打印 = 1
    
    '判断是不是全部已经打印
    '先取出打过的医嘱的最大序号
    lngPrintType = Val(zlDatabase.GetPara("医嘱单打印模式", glngSys, p住院医嘱下达))
    dat重整 = CDate("1900-01-01")
    strSQL = "Select 医嘱重整时间 as 时间 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng病人ID, lng主页ID)
    If Not IsNull(rsTmp!时间) Then dat重整 = CDate(rsTmp!时间 & "")
    
    strSQL = "Select 序号,婴儿姓名 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng病人ID, lng主页ID)
    str婴儿 = "0"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            str婴儿 = str婴儿 & "," & rsTmp!序号
            rsTmp.MoveNext
        Next
    End If
    arr婴儿 = Split(str婴儿, ",")
    
    For i = 0 To 1 '长嘱临嘱
        For j = 0 To UBound(arr婴儿) '婴儿
            '停嘱打印，只需要判断一次
            If i = 0 Then
                strSQL = "Select 1 From 病人医嘱打印 A, 病人医嘱记录 B" & _
                    " Where A.医嘱id=B.ID And A.期效 = 0 And A.病人id=[1] And A.主页id=[2] And Nvl(A.婴儿,0)=[3] And a.打印时间 is not null And (B.确认停嘱时间 Is Not Null And" & _
                    " Not Exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 = 2) " & _
                    IIf(lngPrintType = 1, " Or B.执行终止时间 Is Not Null And Not exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 in (1,2))  or b.校对时间 is not null and not exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 in(1,2,3))", "") & ") And Rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng病人ID, lng主页ID, Val(arr婴儿(j)))
                If Not rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            End If
        
            '未打印的也没有在 病人医嘱打印 中产生
            lng医嘱ID = 0
            lng序号 = 0
            strSQL = "Select 医嘱id From (Select 医嘱id From 病人医嘱打印 Where 病人id =[1] And 主页id =[2] And Nvl(婴儿, 0)=[3] And 期效 =[4]" & _
            " And 打印时间 + 0 >= [5] Order By 页号 Desc, 行号 Desc) Where Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng病人ID, lng主页ID, Val(arr婴儿(j)), i, dat重整)
            If Not rsTmp.EOF Then lng医嘱ID = Val(rsTmp!医嘱ID & "")
            ' lng医嘱id=0 只可能是重整后一次也没打
            If lng医嘱ID <> 0 Then
                strSQL = "Select Nvl(Max(序号), 0) as 序号 From (Select 序号 From 病人医嘱记录 Where (相关id =[1] Or ID =[1])" & _
                    " Union All Select b.序号 From 病人医嘱记录 A, 病人医嘱记录 B Where a.诊疗类别 In ('5', '6') And a.Id =[1] And a.相关id = b.Id)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng医嘱ID)
                If Not rsTmp.EOF Then lng序号 = Val(rsTmp!序号 & "")
            End If
            
            If dat重整 = CDate("1900-01-01") Then
                strSQL = "Select 1 From 病人医嘱记录 A, 诊疗项目目录 B Where a.病人id =[1] And a.主页id =[2] And Nvl(a.婴儿, 0) =[3] And" & vbNewLine & _
                        " a.诊疗项目id = b.Id(+) And a.医嘱状态 Not In (-1, 2) and a.医嘱期效 =[4] And" & vbNewLine & _
                        " ([6] = 1 And a.医嘱状态 = 1 Or a.医嘱状态 <> 1) And Nvl(a.屏蔽打印, 0) = 0 And" & vbNewLine & _
                        " Not Exists (Select 1 From 病人医嘱记录 Where 诊疗类别 = 'F' And ID = a.前提id) And a.序号 >[5] And a.病人来源 = 2 and rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng病人ID, lng主页ID, Val(arr婴儿(j)), i, lng序号, lngPrintType)
                If Not rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            Else
                strSQL = "Select 1 From 病人医嘱记录 A, 诊疗项目目录 B Where a.病人id =[1] And a.主页id =[2] And Nvl(a.婴儿, 0) =[3] And" & vbNewLine & _
                        " a.诊疗项目id = b.Id(+) And a.医嘱状态 Not In (-1, 2) and a.医嘱期效 =[4] And" & vbNewLine & _
                        " ([6] = 1 And a.医嘱状态 = 1 Or a.医嘱状态 <> 1) And Nvl(a.屏蔽打印, 0) = 0 And" & vbNewLine & _
                        " Not Exists (Select 1 From 病人医嘱记录 Where 诊疗类别 = 'F' And ID = a.前提id) And a.序号 >[5] And a.病人来源 = 2 and" & _
                        " Exists (Select 1 From 病人医嘱状态 C Where a.Id = c.医嘱id And c.操作时间 >=[7]) and rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", lng病人ID, lng主页ID, Val(arr婴儿(j)), i, lng序号, lngPrintType, dat重整)
                If Not rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            End If
        Next
        If blnKey Then Exit For
    Next
    
    If Not blnKey Then Get病人医嘱打印 = 2
 
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
        Call SaveErrLog
End Function

Public Function Get病人输液量(ByVal lng病人ID As Long, ByVal lng主页ID) As String
'功能：获取指定病人今天和明天的输液量，对应反回参数  用逗号分隔 "200,300"
'说明：药品计量单位为ml的溶媒、给药途径为输液类 总量＝已发送+未发送，（未发送中含新开的）
'      对于新开的或从未发送的以下医嘱不列入计算：频率为必要时的长嘱，频率为需要时或一次性的临嘱
'      暂存，校对疑问，已做废 医嘱不列入统计；长嘱使用过暂停/启用功能的不单独考虑，一律当做没有暂停过
    Dim rsTmp As ADODB.Recordset
    Dim rs执行时间 As ADODB.Recordset
    Dim strSQL As String, str分解时间 As String, str医嘱IDs As String
    Dim dblToday As Double, dblTomorrow As Double, dblTmp As Double
    Dim datCur As Date, datBegin As Date, datEnd As Date
    Dim lng次数 As Long
    Dim i As Long, j As Long
    Dim varArr As Variant
    
    strSQL = "Select a.单次用量,a.首次用量,a.开始执行时间,a.上次执行时间,Nvl(a.执行终止时间,[4]) as 执行终止时间,a.频率间隔,a.执行时间方案,a.频率次数,a.间隔单位,a.执行频次," & vbNewLine & _
        "     a.医嘱期效,a.停嘱时间,a.天数,nvl(a.可否分零,d.住院可否分零) as 分零,a.总给予量,d.剂量系数,a.相关id" & vbNewLine & _
        "From 病人医嘱记录 A,诊疗项目目录 B,药品特性 C,药品规格 D" & vbNewLine & _
        "Where a.诊疗项目id = b.Id And b.Id = c.药名id And a.收费细目id=d.药品id(+) And a.诊疗类别 In ('5','6') and" & vbNewLine & _
        "     Upper(Nvl(b.计算单位,'NULL')) = 'ML' And c.溶媒=1 And a.病人id =[1] And a.主页id=[2] And" & vbNewLine & _
        "     a.开始执行时间 <= [4] And a.医嘱状态 Not In (-1,2,4) And" & vbNewLine & _
        "     (a.医嘱期效 = 1 And" & vbNewLine & _
        "     (a.执行频次 = '一次性' And a.开始执行时间 >= [3] Or" & vbNewLine & _
        "     a.停嘱时间 Is Null And a.执行时间方案 Is Not Null Or" & vbNewLine & _
        "     a.停嘱时间 Is Not Null And a.执行终止时间 >= [3] And (a.执行时间方案 Is Not Null Or a.执行频次 = '需要时')) Or" & vbNewLine & _
        "     a.医嘱期效 = 0 And" & vbNewLine & _
        "     (a.上次执行时间 Is Null And a.执行时间方案 Is Not Null And Nvl(a.执行终止时间,[3])>=[3] Or" & vbNewLine & _
        "     a.上次执行时间 >= [3] ))"
    '按时间要求过滤出了这7类药品医嘱：
    '1.频率为一次性的临嘱（已发送和未发送）
    '2.频率为指定方案的临嘱（未发送）
    '3.频率为指定方案的临嘱（已发送）
    '4.频率为需要时的临嘱（已发送）
    '5.频率为指定方案的长嘱（从未发送）
    '6.频率为必要时长嘱（至少发送一次）
    '7.频率为指定方案的长嘱（至少发送一次）
    '另还有两类没有被发送过的医嘱：临嘱需要时、长嘱必要时，这两种医嘱如果没有发送则不参数与计算，SQL查询中也不会被过滤出来
    
    On Error GoTo errH
    datCur = zlDatabase.Currentdate
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get病人输液量", lng病人ID, lng主页ID, CDate(Format(datCur, "YYYY-MM-DD 00:00:00")), CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")))
    If rsTmp.EOF Then Get病人输液量 = "0,0": Exit Function
    
    '从 医嘱执行时间 表中取执行时间点
    For i = 1 To rsTmp.RecordCount
        '3.频率为指定方案的临嘱（已发送）'6.频率为必要时长嘱（至少发送一次）
        If Val(rsTmp!医嘱期效 & "") = 1 And rsTmp!执行时间方案 & "" <> "" And rsTmp!停嘱时间 & "" <> "" Or _
           Val(rsTmp!医嘱期效 & "") = 0 And rsTmp!执行时间方案 & "" = "" And rsTmp!上次执行时间 & "" <> "" Then
           
            If InStr("," & str医嘱IDs & ",", "," & Val(rsTmp!相关ID & "") & ",") = 0 Then str医嘱IDs = str医嘱IDs & "," & Val(rsTmp!相关ID & "")
        End If
        rsTmp.MoveNext
    Next
    str医嘱IDs = Mid(str医嘱IDs, 2)
    If str医嘱IDs <> "" Then
        strSQL = "select a.医嘱id,a.要求时间 from 医嘱执行时间 a where a.医嘱id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) and a.要求时间<=[2]"
        Set rs执行时间 = zlDatabase.OpenSQLRecord(strSQL, "Get病人输液量", str医嘱IDs, CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")))
    End If
    rsTmp.MoveFirst
    
    '开始计算
    For i = 1 To rsTmp.RecordCount
        '1.频率为一次性的临嘱（已发送和未发送），开始时间的那一天为准只有一次，单量
        If Val(rsTmp!医嘱期效 & "") = 1 And rsTmp!执行频次 & "" = "一次性" Then
            If Format(rsTmp!开始执行时间 & "", "YYYY-MM-DD") = Format(datCur, "YYYY-MM-DD") Then
                dblToday = dblToday + Val(rsTmp!单次用量 & "")
            Else
                dblTomorrow = dblTomorrow + Val(rsTmp!单次用量 & "")
            End If
        '2.频率为指定方案的临嘱（未发送），先计算次数再分解时间点
        ElseIf Val(rsTmp!医嘱期效 & "") = 1 And rsTmp!执行时间方案 & "" <> "" And rsTmp!停嘱时间 & "" = "" Then
            If Nvl(rsTmp!天数, 0) <> 0 And Not IsNull(rsTmp!执行频次) Then
                '一个频率周期的次数
                If rsTmp!间隔单位 = "周" Then
                    lng次数 = IntEx(rsTmp!天数 * (rsTmp!频率次数 / 7))
                ElseIf rsTmp!间隔单位 = "天" Then
                    lng次数 = IntEx(rsTmp!天数 * (rsTmp!频率次数 / rsTmp!频率间隔))
                ElseIf rsTmp!间隔单位 = "小时" Then
                    lng次数 = IntEx(rsTmp!天数 * (rsTmp!频率次数 / rsTmp!频率间隔) * 24)
                ElseIf rsTmp!间隔单位 = "分钟" Then
                    lng次数 = IntEx(rsTmp!天数 * (rsTmp!频率次数 / rsTmp!频率间隔) * (24 * 60))
                End If
            Else
                '可分零药品时,按总量对单量的倍数计算给药途径的次数,不可分零与一次性使用药品时，按总量对（单量与剂量系数比值取整）的倍数计算给药途径的次数，
                '否则按一个频率周期的次数计算
                If Nvl(rsTmp!分零, 0) = 0 And Nvl(rsTmp!单次用量, 0) <> 0 Then
                    lng次数 = IntEx(rsTmp!总给予量 * rsTmp!剂量系数 / rsTmp!单次用量)
                ElseIf (Nvl(rsTmp!分零, 0) = 1 Or Nvl(rsTmp!分零, 0) = 2) And Nvl(rsTmp!单次用量, 0) <> 0 Then
                    lng次数 = IntEx(rsTmp!总给予量 / IntEx(rsTmp!单次用量 / rsTmp!剂量系数))
                Else
                    lng次数 = Nvl(rsTmp!频率次数, 0)
                End If
            End If
            If Not IsNull(rsTmp!执行时间方案) Or Nvl(rsTmp!间隔单位) = "分钟" Then
                str分解时间 = Calc次数分解时间(lng次数, rsTmp!开始执行时间, CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")), "", Nvl(rsTmp!执行时间方案), rsTmp!频率次数, rsTmp!频率间隔, rsTmp!间隔单位)
            End If
            If str分解时间 <> "" Then
                varArr = Split(str分解时间, ",")
                For j = 0 To UBound(varArr)
                    If Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + Val(rsTmp!单次用量 & "")
                    ElseIf Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + Val(rsTmp!单次用量 & "")
                    End If
                Next
            End If
        '3.频率为指定方案的临嘱（已发送），从 医嘱执行时间 表中执行时间点即可
        ElseIf Val(rsTmp!医嘱期效 & "") = 1 And rsTmp!执行时间方案 & "" <> "" And rsTmp!停嘱时间 & "" <> "" Then
            If Not rs执行时间 Is Nothing Then
                rs执行时间.Filter = "医嘱id=" & Val(rsTmp!相关ID & "")
                For j = 1 To rs执行时间.RecordCount
                    If Between(Format(rs执行时间!要求时间 & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + Val(rsTmp!单次用量 & "")
                    ElseIf Between(Format(rs执行时间!要求时间 & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + Val(rsTmp!单次用量 & "")
                    End If
                    rs执行时间.MoveNext
                Next
            End If
        '4.频率为需要时的临嘱（已发送），这类医嘱只发一次，且当天有效。直接用开始时间判断即可
        ElseIf Val(rsTmp!医嘱期效 & "") = 1 And rsTmp!执行频次 & "" = "需要时" And rsTmp!停嘱时间 & "" <> "" Then
            If Format(rsTmp!开始执行时间 & "", "YYYY-MM-DD") = Format(datCur, "YYYY-MM-DD") Then
                dblToday = dblToday + Val(rsTmp!单次用量 & "")
            Else
                dblTomorrow = dblTomorrow + Val(rsTmp!单次用量 & "")
            End If
        '6.频率为必要时长嘱（至少发送一次），从 医嘱执行时间 表中执行时间点，考虑总量，要排序取出最小时间点，即首次执行时间点
        ElseIf Val(rsTmp!医嘱期效 & "") = 0 And rsTmp!执行时间方案 & "" = "" And rsTmp!上次执行时间 & "" <> "" Then
            If Not rs执行时间 Is Nothing Then
                rs执行时间.Filter = "医嘱id=" & Val(rsTmp!相关ID & "")
                rs执行时间.Sort = "要求时间"
                For j = 1 To rs执行时间.RecordCount
                    dblTmp = Val(rsTmp!单次用量 & "")
                    If j = 1 And Val(rsTmp!首次用量 & "") <> 0 Then dblTmp = Val(rsTmp!首次用量 & "")
                    If Between(Format(rs执行时间!要求时间 & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + dblTmp
                    ElseIf Between(Format(rs执行时间!要求时间 & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + dblTmp
                    ElseIf Format(rs执行时间!要求时间 & "", "YYYY-MM-DD HH:MM:SS") > Format(datCur + 1, "YYYY-MM-DD 23:59:59") Or _
                        Format(rs执行时间!要求时间 & "", "YYYY-MM-DD HH:MM:SS") > Format(rsTmp!执行终止时间 & "", "YYYY-MM-DD HH:MM:SS") Then
                        Exit For
                    End If
                    rs执行时间.MoveNext
                Next
            End If
        '7.频率为指定方案的长嘱（至少发送一次）
        '5.频率为指定方案的长嘱（从未发送）
        '7和5用一样的处理方式：重新计算分解时间点
        ElseIf Val(rsTmp!医嘱期效 & "") = 0 And rsTmp!执行时间方案 & "" <> "" And rsTmp!上次执行时间 & "" <> "" Or _
            Val(rsTmp!医嘱期效 & "") = 0 And rsTmp!执行时间方案 & "" <> "" And rsTmp!上次执行时间 & "" = "" Then
            '如果首次用量不为0，开始时间以医嘱开始执行时间为准为了计算出首次执行时间点用于判断
            If Val(rsTmp!首次用量 & "") = 0 And Format(datCur, "YYYY-MM-DD 00:00:00") > Format(rsTmp!开始执行时间 & "", "YYYY-MM-DD HH:MM:SS") Then
                datBegin = Format(datCur, "YYYY-MM-DD 00:00:00")
            Else
                datBegin = rsTmp!开始执行时间
            End If
        
            If Format(rsTmp!执行终止时间 & "", "YYYY-MM-DD HH:MM:SS") > Format(datCur + 1, "YYYY-MM-DD 23:59:59") Then
                datEnd = CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59"))
            Else
                datEnd = CDate(Format(rsTmp!执行终止时间 & "", "YYYY-MM-DD HH:MM:SS"))
            End If
            
            str分解时间 = Calc段内分解时间(datBegin, datEnd, "", Nvl(rsTmp!执行时间方案), Nvl(rsTmp!频率次数, 0), Nvl(rsTmp!频率间隔, 0), Nvl(rsTmp!间隔单位), rsTmp!开始执行时间)
            If str分解时间 <> "" Then
                varArr = Split(str分解时间, ",")
                For j = 0 To UBound(varArr)
                    dblTmp = Val(rsTmp!单次用量 & "")
                    If j = 0 And Val(rsTmp!首次用量 & "") <> 0 Then dblTmp = Val(rsTmp!首次用量 & "")
                    If Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + dblTmp
                    ElseIf Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + dblTmp
                    End If
                Next
            End If
        End If
        rsTmp.MoveNext
    Next
    Get病人输液量 = dblToday & "," & dblTomorrow
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function HaveOperateAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intType As Integer) As Boolean
'功能：判断指定病人是否还存在可以操作的医嘱
'参数：intType 0-校对，1－确认停止
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    
    On Error GoTo errH
    If intType = 0 Then
        If gblnKSSStrict Or gbln手术分级管理 Or gbln输血分级管理 Or gbln血库系统 Then
            strWhere = strWhere & " And (Nvl(A.审核状态,0) Not in(1,3,7" & IIf(gbln血库系统 = True, "", ",4,5") & ") or a.医嘱期效=0 and a.审核状态=1 and a.紧急标志=1 and (instr(',5,6,',A.诊疗类别)>0 or A.诊疗类别='E' and B.操作类型='2'))"
        End If
        strSQL = "select 1 from 病人医嘱记录 a,诊疗项目目录 b where a.诊疗项目id=b.id(+) and A.医嘱状态=1 and a.病人id=[1] and a.主页id=[2]" & strWhere & _
                " And Exists ( Select M.姓名 From 人员表 M,执业类别 N" & _
                " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
                " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')) And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3 and Rownum<2"
    ElseIf intType = 1 Then
        strSQL = "select 1 from 病人医嘱记录 a where A.医嘱状态=8 and Nvl(a.医嘱期效,0)=0 And a.病人id=[1] and a.主页id=[2] And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3 and Rownum<2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob.HaveOperateAdvice", lng病人ID, lng主页ID)
    HaveOperateAdvice = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub PlugInInSideBar(ByRef cbsMain As Object, ByVal strFuncName As String, Optional ByVal intInSide As Integer)
'功能：设置工具栏按钮，外挂卡片窗体中的功能按钮
'intInSide 否要要添加工具栏按钮。默认是要添加
    Dim objControl As CommandBarControl
    Dim objMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim varArr As Variant
    Dim strTmp As String
    Dim lngTmp As Long
    Dim objCbs As Object
    Dim lngidx As Long
    Dim i As Long
    Dim strName As String, lngIcon As Long
    
    If strFuncName = "" Then Exit Sub
    varArr = Split(strFuncName, "|")
    
    Set objCbs = cbsMain
    
    '扩展:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objMenuBar = objCbs.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenuBar Is Nothing Then Set objMenuBar = objCbs.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
 
    Set objMenuBar = objCbs.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展(&A)", objMenuBar.Index + 1, False)
 
    With objMenuBar.CommandBar.Controls
        For i = 0 To UBound(varArr)
            strTmp = varArr(i)
            
            strName = strTmp
            lngIcon = conMenu_Tool_PlugIn
            
            If InStr(strTmp, ",") > 0 Then
                strName = Split(strTmp, ",")(0)
                lngIcon = Val(Split(strTmp, ",")(1))
            End If
 
            Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + 1 + i, strName)
                objControl.IconId = lngIcon
                objControl.ToolTipText = strName
                objControl.Style = xtpButtonIconAndCaption
        Next
    End With
    
    If intInSide = 0 Then
        '工具栏添加
        '找到要添加的位置
        For Each objControl In objCbs(2).Controls '先求出前面的最后一个Control
            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
                Set objControl = objCbs(2).Controls(objControl.Index - 1)
                lngidx = objControl.Index
                Exit For
            End If
        Next
        
        With objCbs(2).Controls
            For i = UBound(varArr) To 0 Step -1
                strTmp = varArr(i)
                
                strName = strTmp
                lngIcon = conMenu_Tool_PlugIn
                
                If InStr(strTmp, ",") > 0 Then
                    strName = Split(strTmp, ",")(0)
                    lngIcon = Val(Split(strTmp, ",")(1))
                End If
                
                Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + 1 + i, strName, lngidx + 1)
                    objControl.IconId = lngIcon
                    objControl.ToolTipText = strName
                    objControl.Style = xtpButtonIconAndCaption
            Next
        End With
    End If
    cbsMain.RecalcLayout
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function Get传染病状态(ByVal lng记录 As Long, ByVal lng填写 As Long, ByVal lng状态 As Long) As String
'功能：获取当前病人的传染病状态
    Dim strTmp As String
    If lng状态 <> 0 Then
'       -1-已拒绝 1-已接收;2-已呈报;3-审核通过;4-待医生返修；5-医生已返修完成待审核
        Select Case lng状态
        Case -1
            strTmp = "已拒绝"
        Case 1
            strTmp = "已接收"
        Case 2
            strTmp = "已呈报"
        Case 3
            strTmp = "审核通过"
        Case 4
            strTmp = "待医生修改"
        Case 5
            strTmp = "医生修完改待审核"
        End Select
    ElseIf lng填写 > 0 Then
        strTmp = "已填写"
    ElseIf lng记录 > 0 Then
        strTmp = "有阳性结果"
    End If
    Get传染病状态 = strTmp
End Function

Public Sub FuncEPRReport(frmMain As Form, ByVal lng医嘱ID As Long, ByVal str诊疗类别 As String, _
        Optional ByVal lng报告ID As Long, Optional ByVal str检查报告ID As String, _
        Optional ByVal int场合 As Integer = 1)
'功能：查阅报告
'参数: int场合:1-门诊一览,2-住院一览
    Dim strPrivs As String, strSQL As String
    Dim lngRet As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    '调用数据交换平台，向LIS,PACS查阅报告
    If gobjExchange Is Nothing Then
        On Error Resume Next
        Set gobjExchange = CreateObject("zlExchange.clsExchange")
        If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
        err.Clear: On Error GoTo 0
    End If
    If Not gobjExchange Is Nothing Then
        '检验行存的是采集方法（诊疗类别为E），所以只判断检查行
        Call gobjExchange.SendMsg(IIf(str诊疗类别 = "D", 4, 3), "医嘱ID::" & lng医嘱ID & "||操作员姓名::" & UserInfo.姓名 & "||操作员缺省部门::" & UserInfo.部门名)
        Exit Sub
    End If
    strPrivs = GetInsidePrivs(IIf(int场合 = 1, p门诊医嘱下达, p住院医嘱下达))
    '先判断是否可以继续操作
    Select Case CheckEPRReport(lng医嘱ID, lng报告ID)
    Case 0
        MsgBox "该医嘱的报告没有书写！", vbInformation, gstrSysName
        Exit Sub
    Case 2
        If InStr(strPrivs, ";查阅未完成报告;") > 0 Then
            MsgBox "注意：该医嘱的报告还没有正式签名！", vbInformation, gstrSysName
        Else
            MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，你没有权限操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End Select
    
    If str诊疗类别 = "D" Then
        If HaveRIS Then
            'RIS报告兼容
            strSQL = "select 1 from 病人医嘱报告 a,电子病历记录 b,电子病历格式 c where a.病历id=b.id and b.id=c.文件id and a.医嘱id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmMain.Caption, lng医嘱ID)
            If rsTmp.EOF Then
                lngRet = gobjRis.ShowViewReport(frmMain.hwnd, lng医嘱ID, InStr(strPrivs, ";报告打印;") > 0)
                If lngRet = 0 Then Exit Sub
            End If
        End If
    End If
    '执行操作
    '新版PACS报告，直接强制使用新版PACS报告编辑器
    If str检查报告ID <> "" Then
        If CreateObjectPacs(gobjPublicPacs) Then
            Call gobjPublicPacs.zlDocShowReport(lng医嘱ID, , False, frmMain)
        End If
    Else
        '查阅报告
        Call gobjRichEPR.ViewDocument(frmMain, lng报告ID, False)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function InitNurseIntegrate(Optional blnMsg As Boolean = False) As Boolean
'判断如果整体护理部件为空就初始化
    If gobjNurseIntegrate Is Nothing Then
        On Error Resume Next
        Set gobjNurseIntegrate = CreateObject("zlNurseIntegrate.clsNurseIntegrate")
        If Not gobjNurseIntegrate Is Nothing Then
            If gobjNurseIntegrate.zlInitCommon(gcnOracle, gstrDBUser) = False Then
                Set gobjNurseIntegrate = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    If blnMsg = True And gobjNurseIntegrate Is Nothing Then
        MsgBox "整体护理接口部件：zlNurseIntegrate  创建失败！", vbInformation, gstrSysName
    End If
    InitNurseIntegrate = Not gobjNurseIntegrate Is Nothing
End Function



Public Function CheckStructAddr(ByVal objCtl As PatiAddress, ByVal lngLen As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查结构化地址控件中的信息录入是否正确
    '入参:objCtl-结构化地址控件，lngLen-限制长度
    '返回:True-输入信息合法
    '编制:李南春
    '日期:2015-12-7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If zlCommFun.ActualLen(objCtl.Value) > lngLen Then
        MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "最多只能输入" & lngLen \ 2 & "个汉字,请检查。", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    If objCtl.CheckNullValue(, True, False) <> "" Then
        MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "的" & objCtl.CheckNullValue & "尚未输入,请检查。", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    CheckStructAddr = True
End Function

Public Function zlReadAddrInfo(ByVal objCtrl As PatiAddress, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
                               ByVal intType As Integer, Optional ByVal strAddress As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定的病人地址信息到控件中
    '入参:objCtrl-结构化地址控件,intType -地址类型1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址
    '返回:
    '编制:李南春
    '日期:2015/12/3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    strSQL = "Select 省,市,县,乡镇,其他 From 病人地址信息 Where 病人ID=[1] and Nvl(主页ID,0)=[2] and 地址类别=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询结构化地址", lng病人ID, lng主页ID, intType)
    If rsTmp.RecordCount > 0 Then
        Call objCtrl.LoadStructAdress(Nvl(rsTmp!省), Nvl(rsTmp!市), Nvl(rsTmp!县), Nvl(rsTmp!乡镇), Nvl(rsTmp!其他))
    Else
        objCtrl.Value = strAddress
    End If
    zlReadAddrInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function FuncTraReaction(ByVal lng医嘱ID As Long, ByVal lngMoudle As Long, ByVal blnMoved As Boolean, Optional ByVal lng收发id As Long) As Boolean
'功能：输血反应
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long, lng主页ID As Long, lng病人来源 As Long
    If InitObjBlood(True) = False Then Exit Function
    On Error GoTo errH
    strSQL = "Select B.病人ID,B.主页ID,B.病人来源,A.ID 就诊ID From 病人挂号记录 A,病人医嘱记录 B where B.挂号单=A.NO(+) And  B.id=[1]"
    If blnMoved = True Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "输出反应", lng医嘱ID)
    lng病人ID = Val("" & rsTemp!病人ID)
    If IsNull(rsTemp!主页ID) Then
        lng主页ID = Val("" & rsTemp!就诊ID)
    Else
        lng主页ID = Val("" & rsTemp!主页ID)
    End If
    lng病人来源 = Val("" & rsTemp!病人来源)
    Call gobjPublicBlood.zlShowBloodReaction(Nothing, glngSys, lngMoudle, 1, lng病人ID, lng主页ID, lng病人来源, 1, lng收发id)
    FuncTraReaction = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FuncTraReactionRecord(ByVal frmParent As Object, ByVal lng场合 As Long, ByVal lngMoudle As Long) As Boolean
'功能：输血反应调用接口
    On Error GoTo errH
    If InitObjBlood(True) = False Then Exit Function
    Call gobjPublicBlood.zlShowBloodReactionRecord(frmParent, glngSys, lngMoudle, lng场合)
    FuncTraReactionRecord = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckZLPass(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：判断当前病人是否有未审核通过的医嘱
'返回：true 通过，false存在未通的药品医嘱

    Dim rsTmp As ADODB.Recordset
    Dim strErr As String
    Dim blnTmp As Boolean
    
    CheckZLPass = True
    
    
    On Error Resume Next
    
    If gobjPass Is Nothing Then
        Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "合理用药监测", True)
    End If
    err.Clear
    
    On Error GoTo errH
    
    If Not gobjPass Is Nothing Then
        blnTmp = gobjPass.ZLPharmReviewResultView(lng病人ID, lng主页ID, rsTmp, strErr)
    End If
    
    If strErr = "" Then
        If blnTmp Then
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF Then
                    CheckZLPass = False
					Call gobjPass.ZLPharmReviewResultShow(frmParent, lng病人ID, lng主页ID)
                End If
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function