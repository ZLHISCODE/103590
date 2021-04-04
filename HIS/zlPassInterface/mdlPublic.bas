Attribute VB_Name = "mdlPublic"
Option Explicit

'常量---------------------
Public Const G_STR_PASS As String = "合理用药部件"
Public Const G_STR_MATCH As String = "abcdefghigklmnopkrstuvwxyzABCDEFGHIGKLMNOPKRSTUVWXYZ0123456789"" </>_="
Public Const G_INT_MODEL_0 As Integer = 0
Public Const G_INT_MODEL_1 As Integer = 1
Public Const G_STR_SPLIT As String = "&&"
Public Const SW_SHOWNORMAL = 1
'API声明
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'全局变量-------------------------------
Public gfrmMain As Object                   '父窗体
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gobjComLib As Object                    '公共部件对象ZL9ComLib
Public gcolPrivs As Collection              '记录内部模块的权限
Public gstrSysName As String                '系统名称
Public gstrDBUser As String                 '当前数据库用户
Public glngSys As Long
Public gbytUseType As Byte                  '0-医嘱下达
                                            '1-临床路径项目的医嘱生成
                                            '2-临床路径添加路径外项目的医嘱，不允许选择病人
                                            '3-医嘱顺序调整(必须显示已停止的医嘱，因为移动时经过那些医嘱，序号要一起调整)
Public glngObject As Long                   '记录对象个数
Public gobjPlugIn   As Object
Public gsngWaitTime   As Single               '访问等待最大间隔秒数
Public gsngAutoLinkTime As Single              '每隔5分钟检查连接
Public gblnBreak     As Boolean             'T-断开连接;F-已连接
Public gsngCheckLinkTime As Single            '
Public mstrLike As String                   '输入匹配方式
Public mint简码 As Integer                  '简码匹配方式：0-拼音,1-五笔
Public gstrMatchMode As String '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
'------------------------------------------------------------------
'合理用药相关启用参数
'------------------------------------------------------------------

Public gbytPass As Byte             'ZLHIS中使用PASS接口类型,0-未使用,1-美康,2-大通,3-太元通,4-药卫士,5-杭州逸曜,6-中联信息
Public gbytBlackLamp As Byte        '是否允许禁忌药品
Public gbytReason As Byte           '禁忌药品要求填写原因
Public gbytSuperVolume As Byte      '是否禁止超极量药品
Public gbytOutBlackLamp As Byte     '是否允许院外执行的禁忌药品医嘱
Public gobjPass As Object           '3-太元通接口对象,4-药卫士
Public gbytOpenLog As Byte          '开启大通接口调试日志 0-不启用，1-启用
Public gbytSysSet As Byte           '美康允许使用系统设置 1-显示，0-隐藏
Public gstrVersion As String        '标识接口版本号
Public gblnPharmReview As Boolean   '美康启用药师审方干预系统
Public gblnPrePregnancy As Boolean  '允许选择备孕项 Preparation of pregnancy
Public gblnTEST As Boolean          '启用静默式审查 T-调用MDC_DoCheck(0,1)方法，即审查有问题不弹审查结果，只进行问题数据的采集，即使静默式审查出了用药问题，也不进行医嘱行为的拦截，以便于全院初始实施阶段大量用户数据规则清洗所用，减少无效信息对医生业务的干扰。
'---------------
Public gstrIP           As String           '服务器IP
Public gstrPort         As String           '服务器端口号
Public gstrDrugIP       As String           '药品说明书IP
Public gstrDrugPort     As String           '药品说明书端口号
Public gstrUser         As String           '用户名
Public gstrPWD          As String           '用户密码
Public gstrPortPlus     As String           '服务器端口号
Public gstrHOSCODE      As String           '医院编码
Public gstrStatusEdit   As String           '编辑病人状态
Public gstrStatusGet    As String           '获取病人状态   http://192.168.0.231:8080/ords/patstatus/pat/getpatstatus
Public gstrStatusSave As String           '保存病人状态   http://192.168.0.231:8080/ords/patstatus/pat/saverecord

Public gbytType         As Byte             '杭州逸曜:0-共用;1-非共用

Public gint过敏登记有效天数 As Integer
Public gblnInitOK As Boolean         '用于标记初始化执行状态 'T-执行过初始化;F-未执行过初始化
Public gblnPassOK   As Boolean         'T-初始化成功,允许操作;F=初始化失败 禁止Pass接口调用

Public glngPatiID As Long              '记录当前病人ID
Public glng主页ID As Long
Public gblnTip As Boolean             '美康4.0用于避免重复向接口中传人相同的药品信息
Public gbytOpen As Byte   '标记悬浮窗体

'登录用户结构
Public Type TYPE_USER_INFO
    id As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
    专业技术职务 As String
    专业技术编码 As String
    用药级别 As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum DataEnum
    responseText = 1
    responseBody = 2
End Enum

Public UserInfo As TYPE_USER_INFO

Public gobjCOL As clsVSCOL           '当前医嘱列映射
Public gobjAdvice As Object         '当前医嘱列表对象 vsAdvice
Public gobjCmdAlley As Object           '当前PASS过敏史按钮

Public glngModel As Long                '当前场景gbytModel 0-门诊编辑,1-住院编辑，2-住院医嘱清单,3-护士校对,4-门诊医嘱清单
Public gobjDiags As clsDiags              '门诊
Public gint场合 As Integer              ' 调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
Public gcolPASSExe As Collection        '菜单索引映射
Public gcolPASSState As Collection      '美康菜单状态值映射


Public gobjMap As clsPassMap  '映射对象
Public gobjPati As clsPatient

Public gblnOpen As Boolean    '简要信息是否打开
Public glngDrugID As Long    '记录下上一次传人的药品ID

'美康功能号
Public Enum G_PASS_MK
    MK_检测PASS菜单状态 = 0
    MK_住院保存审查 = 1
    MK_住院提交审查 = 2
    MK_手工调用审查 = 3
    MK_单药警告 = 6
    MK_系统设置 = 11
    MK_用药研究 = 12
    MK_药品配对信息 = 13
    MK_给药途径配对信息 = 14
    MK_病生状态过敏史查看 = 21
    MK_病生状态过敏史 = 22
    MK_门诊保存审查 = 33
    MK_药物临床信息参考 = 101
    MK_药品说明书 = 102
    MK_病人用药教育 = 103
    MK_检验值 = 104
    MK_医院药品信息 = 105
    MK_医药信息中心 = 106
    MK_中国药典 = 107
    MK_药物_药物相互作用 = 201
    MK_药物_食物相互使用 = 202
    MK_国内注射剂配伍 = 203
    MK_国外注射剂配伍 = 204
    MK_禁忌症 = 205
    MK_副作用 = 206
    MK_老年人用药 = 207
    MK_儿童用药 = 208
    MK_妊娠期用药 = 209
    MK_哺乳期用药 = 210
    MK_关闭浮动窗口 = 402 '关闭当前所有浮动窗口
    MK_警示提示窗口 = 403  '显示警示提示窗口
End Enum

Public Enum G_PASS_MK4
    MK4_检测PASS菜单状态 = 0
    MK4_审查
    MK4_自动审查
    MK4_药品说明书 = 11
    MK4_药物专论 = 21
    MK4_病人用药教育 = 31
    MK4_中国药典 = 41
    MK4_药品简要信息 = 51
    MK4_药物相互作用 = 61
    MK4_药食相互作用 = 62
    MK4_体外配伍 = 63
    MK4_配伍浓度 = 64
    MK4_药物禁忌症 = 65
    MK4_药物适应症 = 66
    MK4_不良反应 = 67
    MK4_肝损害剂量 = 68
    MK4_肾损害剂量 = 69
    MK4_儿童用药 = 70
    MK4_妊娠用药 = 71
    MK4_哺乳用药 = 72
    MK4_老人用药 = 73
    MK4_成人用药 = 74
    MK4_性别用药 = 75
    MK4_细菌耐药率 = 76
End Enum

'美康功3.0菜单索引值
Public Enum G_MK_INDEX
    MK_IX_药物临床信息参考 = 0
    MK_IX_药品说明书 = 1
    MK_IX_中国药典
    MK_IX_病人用药教育
    MK_IX_检验值
    MK_IX_专项信息
    MK_IX_药物相互作用
    MK_IX_药食相互作用
    MK_IX_国内注射剂配伍
    MK_IX_国外注射剂配伍
    MK_IX_禁忌症
    MK_IX_副作用
    MK_IX_老年人用药
    MK_IX_儿童用药
    MK_IX_妊娠期用药
    MK_IX_哺乳期用药
    MK_IX_医药信息中心
    MK_IX_药品配对信息
    MK_IX_给药途径配对信息
    MK_IX_医院药品信息
    MK_IX_系统设置
    MK_IX_用药研究
    MK_IX_警告
    MK_IX_审查
End Enum
'美康功4.0菜单索引值
Public Enum G_MK4_INDEX
    MK4_IX_审查 = 0
End Enum

'太元通 功能号
Public Enum G_PASS_TYT
    TYT_用药规范 = 0
    TYT_药嘱审查 = 1
    TYT_药品提示 = 2
    TYT_医药知识库 = 3
    TYT_系统配置 = 4
    TYT_审查详情 = 5
End Enum

'杭州逸曜 功能号
Public Enum G_PASS_HZYY
    HZYY_药品说明书 = 0
    HZYY_药嘱审查 = 1
End Enum
'中联信息 功能号
Public Enum G_PASS_ZL
    ZL_药嘱审查 = 0
    ZL_病人状态
End Enum

Public Enum G_PASS_UseStation
    US_InDoctor = 0     '住院医生站
    US_InNurse = 1      '住院护士站
    US_Intech = 2       '住院医技站
End Enum

'内部应用模块号定义
Public Enum Enum_Inside_Program
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    P药品处方发药 = 1341        '1341    药品处方发药
    P药品部门发药 = 1342        '1342    药品部门发药
    PPIVA管理 = 1345        '1345    PIVA管理
End Enum

Public Enum G_TYPE_FUN
    FUN_医嘱信息 = 1
    FUN_输出内容 = 3
    FUN_医嘱信息_DTBS = 4
    FUN_审查结果 = 5
    FUN_医嘱信息_HZYY = 6
    FUN_审查结果_HZYY = 7
    FUN_医嘱信息_ZL = 8
    FUN_审查结果_ZL = 9
    FUN_审查结果_YWS = 10
    FUN_反向问诊_ZL = 11
    FUN_药师审查_ZL = 12
    FUN_病人状态_ZL = 13
End Enum

Public Enum G_TYPE_FLOATWIN
    FLOATWIN_CLOSE = 0   '关闭
    FLOATWIN_DRUG = 1    '药品信息提示窗
    FLOATWIN_WARN = 2    '警示窗体
End Enum

'获得鼠标指针在屏幕坐标上的位置
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'获得窗口在屏幕坐标中的位置
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'判断指定的点是否在指定的矩形内部
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
'准备用来使窗体始终在最前面
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'用来移动窗体
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'获取窗体状态
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'HWND hwnd, // 指定分层窗口句柄
'COLORREF crKey, // 指定需要透明的背景颜色值，可用RGB()宏
'BYTE bAlpha, // 设置透明度，0表示完全透明，255表示不透明
'DWORD dwFlags // 透明方式
'       其中，dwFlags参数可取以下值：
'       LWA_ALPHA=&H2时：crKey参数无效，bAlpha参数有效；
'       LWA_COLORKEY=&H1：窗体中的所有颜色为crKey的地方将变为透明，bAlpha参数无效。其常量值为1。
'       LWA_ALPHA | LWA_COLORKEY：crKey的地方将变为全透明，而其它地方根据bAlpha参数确定透明度。
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = -4&
Public Const WM_MOUSEWHEEL = &H20A
 
Public glngOldWindowProc As Long '用来保存系统默认的窗口消息处理函数的地址

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const SWP_NOACTIVATE = &H10 '不激活窗体
Public Const GWL_EXSTYLE  As Long = (-20)
Public Const WS_EX_TOPMOST As Long = &H8
Public Const HWND_TOPMOST As Long = -1
Public Const SW_SHOWMAXIMIZED = 3
'API:GetSystemMetrics
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim strPara As String
    Dim arrTemp As Variant
    
    gbytPass = Val(zlDatabase.GetPara(30, glngSys))  '接口类型
    If gbytPass = UNPASS Then Exit Function
    gstrVersion = zlDatabase.GetPara(228, glngSys) '标识接口版本号
    '初始成功过不再重复读取参数值（gbytPass由于模块禁用或权限禁用的原因会设为:0-UNPASS，故每次需要重新读取）
    If gbytPass = MK Or gbytPass = YWS Then
        gbytSysSet = Val(zlDatabase.GetPara(226, glngSys))
        If gbytPass = MK And gstrVersion = "4.0" Then
            If Not MK_GetPara Then
                MsgBox "合理用药监测参数配置有误,请到:" & vbCrLf & _
                    "【临床参数设置】->【业务流程控制】->【合理用药接口】->【设置】中配置。" & vbCrLf & _
                    "在正确配置之前，相应的功能不能使用。", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
    ElseIf gbytPass = DT Then  '大通
        gbytOpenLog = Val(zlDatabase.GetPara(225, glngSys))
        If gstrVersion = "4.0" Then
            gstrHOSCODE = zlDatabase.GetPara(90001, glngSys, , "1513")
        End If
    ElseIf gbytPass = HZYY Then '杭州逸曜
        Call HZYY_GetPara
    ElseIf gbytPass = ZL Then
        If Not ZL_GetPara Then
            MsgBox "合理用药监测参数配置有误,请到:" & vbCrLf & _
                "【临床参数设置】->【业务流程控制】->【合理用药接口】->【设置】中配置。" & vbCrLf & _
                "在正确配置之前，相应的功能不能使用。", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    gbytBlackLamp = Val(zlDatabase.GetPara(161, glngSys))  '是否允许禁忌药品
    gbytReason = Val(zlDatabase.GetPara(249, glngSys)) '249    禁忌药品要求填写原因
    gbytSuperVolume = Val(zlDatabase.GetPara(182, glngSys)) '是否禁止超极量药品
    
    gbytOutBlackLamp = Val(zlDatabase.GetPara(189, glngSys)) '是否允许院外执行的禁忌药品医嘱
    
    '皮试结果有效时间
    gint过敏登记有效天数 = Val(zlDatabase.GetPara(70, glngSys))

    InitSysPar = True
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
'功能：读取当前登录人员或指定人员的人员性质
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    If str姓名 <> "" Then
        strSQL = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str姓名)
    Else
        strSQL = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.id)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人过敏记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bytFunc As Byte = 1) As ADODB.Recordset
'功能：获取病人过敏记录
'参数：bytFunc=1 传入病人历次所有过敏记录;0=传入病人本次就诊过敏记录
    Dim strSQL As String
    
    On Error GoTo errH
    If bytFunc = 0 Then
        If lng主页ID = 0 Then
            strSQL = "Select Distinct 药物ID,药物名,过敏源编码,过敏反应,记录时间 From 病人过敏记录 Where 病人ID=[1] And 结果=1 And Nvl(过敏时间,记录时间)>Trunc(Sysdate-[3])"
        Else
            strSQL = "Select Distinct 药物ID,药物名,过敏源编码,过敏反应,记录时间 From 病人过敏记录 Where 病人ID=[1] And 主页ID=[2] And 结果=1"
        End If
    Else
        strSQL = "Select Distinct 药物ID,药物名,过敏源编码,过敏反应,记录时间 From 病人过敏记录 Where 病人ID=[1] And 结果=1"
    End If
    Set Get病人过敏记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng主页ID, gint过敏登记有效天数)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人诊断记录(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str类型 As String) As ADODB.Recordset
'功能：获取病人诊断记录
'参数：lng就诊ID：门诊病人传挂号ID，住院病人传主页ID
'       诊断类型-1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;
'        11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断
'       记录来源:1-病历；2-入院登记；3-首页整理(门诊医生站,诊断摘要);
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select a.ID,a.疾病id, a.诊断id, a.诊断描述, a.诊断次序, Nvl(b.编码, c.编码) As 编码, NVL(Nvl(b.名称, c.名称),a.诊断描述) 名称" & vbNewLine & _
             ",a.记录日期,a.记录人 " & vbNewLine & _
             "From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C" & vbNewLine & _
             "Where a.病人id = [1] And a.主页id = [2] And 取消时间 Is Null And 记录来源 IN (1, 3) And Instr(',' ||[3]|| ',', ',' || 诊断类型 || ',') > 0 And a.疾病id = b.Id(+) And" & vbNewLine & _
             "      a.诊断id = c.Id(+)" & vbNewLine & _
             "Order By 记录来源, 诊断类型, 诊断次序"
    Set Get病人诊断记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng就诊ID, str类型)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人病生理情况(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：根据病人ID、主页ID获取病人病生理情况
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH

    If lng主页ID = 0 Then
        lng主页ID = Val(zlDatabase.GetPara(21, glngSys))
        strSQL = "Select 病生理情况" & vbNewLine & _
                 "From 病人挂号记录" & vbNewLine & _
                 "Where 病人id = [1] And 登记时间 > Trunc(Sysdate-[2]) And 病生理情况 Is Not Null And Rownum = 1"
    Else
        strSQL = "Select 信息值 As 病生理情况" & vbNewLine & _
                 "From 病案主页从表 Where 病人id = [1] And 主页id = [2] And 信息名 = '病生理情况'"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            Get病人病生理情况 = Get病人病生理情况 & "," & rsTmp!病生理情况
            rsTmp.MoveNext
         Wend
        Get病人病生理情况 = Mid(Get病人病生理情况, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人手麻记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人过敏记录
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 手术操作ID,已行手术,手术开始时间,手术结束时间 From 病人手麻记录 Where 病人ID=[1] And 主页ID=[2] "

    Set Get病人手麻记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPatiOperation(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal str挂号单 As String) As ADODB.Recordset
'功能：获取病人过敏记录
    Dim strSQL As String
    
    On Error GoTo errH
    If str挂号单 = "" Then
        strSQL = " And a.病人id = [1] And a.主页id = [2] "
    Else
        strSQL = "  And a.挂号单 = [3] "
    End If
    strSQL = "Select a.Id, a.手术时间, c.名称, c.编码" & vbNewLine & _
               "From 病人医嘱记录 A, 疾病诊断对照 B, 疾病编码目录 C" & vbNewLine & _
               "Where a.诊疗项目id = b.手术id And b.疾病id = c.Id And a.诊疗类别 = 'F' And a.医嘱状态  In (1,2,3,5,8) " & strSQL
    Set GetPatiOperation = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng主页ID, str挂号单)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiSymptom(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：根据病人ID、主页ID获取病人症状（太元通接口使用）
'lng主页Id :门诊传挂号ID
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.编码,a.名称 From 病人症状记录 a " & vbNewLine & _
            "Where a.病人ID=[1] And a.主页ID=[2] "
    Set GetPatiSymptom = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：根据病人ID、主页ID获取病人基本信息
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select A.住院号, A.当前床号, A.出生日期, Nvl(B.姓名, A.姓名) 姓名, Nvl(B.性别, A.性别) 性别, Nvl(B.年龄, A.年龄) 年龄, A.门诊号, A.健康号,A.身份证号,B.身高,B.体重" & vbNewLine & _
            "From 病人信息 A, 病案主页 B" & vbNewLine & _
            "Where A.病人id = B.病人id And A.病人id = [1] And B.主页id = [2]"

    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get频率信息_名称(ByVal str频率 As String, int频率次数 As Integer, _
    int频率间隔 As Integer, str间隔单位 As String, str范围 As String, Optional str频率编码 As String) As Boolean
'功能：返回频率的相关信息
'参数：str频率=频率名称
'      str范围=1-西医,2-中医,-1-一次性,-2-持续性
'返回：当按名称取到时，返回True，否则返回False
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
    
    int频率次数 = 0
    int频率间隔 = 0
    str间隔单位 = ""
    
    strSQL = "Select 频率次数,频率间隔,间隔单位,编码 From 诊疗频率项目 Where 名称=[1] And Instr([2],','||适用范围||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str频率, "," & str范围 & ",")
    If Not rsTmp.EOF Then
        int频率次数 = NVL(rsTmp!频率次数, 0)
        int频率间隔 = NVL(rsTmp!频率间隔, 0)
        str间隔单位 = NVL(rsTmp!间隔单位)
        str频率编码 = "" & rsTmp!编码
        Get频率信息_名称 = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDoctorTitleType(ByVal strDoctTitle As String) As String
'功能：根据医生职称返回职称类别
'返回值：
'C --副教授；教授；副主任医师；主任医师；专家
'B—主治医师；讲师
'A—除以上的其他职称

    If InStr(";副教授;教授;副主任医师;主任医师;专家;", ";" & strDoctTitle & ";") > 0 Then
        GetDoctorTitleType = "C"
    ElseIf InStr(";主治医师;讲师;", ";" & strDoctTitle & ";") > 0 Then
        GetDoctorTitleType = "B"
    Else
        GetDoctorTitleType = "A"
    End If

End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.id = rsTmp!id
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = NVL(rsTmp!简码)
            UserInfo.姓名 = NVL(rsTmp!姓名)
            UserInfo.部门ID = NVL(rsTmp!部门ID, 0)
            UserInfo.部门码 = NVL(rsTmp!部门码)
            UserInfo.部门名 = NVL(rsTmp!部门名)
            UserInfo.性质 = Get人员性质
            UserInfo.专业技术职务 = NVL(rsTmp!专业技术职务)
            UserInfo.专业技术编码 = Sys.RowValue("专业技术职务", UserInfo.专业技术职务, "编码", "名称")
            GetUserInfo = True
        End If
    End If
    gstrDBUser = UserInfo.用户名
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function PassCheckPrivs(ByVal lngModel As Long, Optional ByVal blnInit As Byte = False) As Boolean
'功能:根据模块号获取模块具有的权限
'参数:blnInit -是否初始化(住院医生站初始化时需要判断住院医嘱下达和住院医嘱发送的合理用药检测权限)
    Dim blnDo As Boolean
    WriteLog "clsPass", "PassCheckPrivs", "PassCheckPrivs_Begin"
    Select Case lngModel
    
    Case PM_门诊编辑, PM_门诊医嘱清单
        If InStr(GetInsidePrivs(p门诊医嘱下达), "合理用药监测") > 0 Then blnDo = True
    Case PM_住院医嘱清单
        If blnInit Then
            If InStr(GetInsidePrivs(p住院医嘱下达) & GetInsidePrivs(p住院医嘱发送), "合理用药监测") > 0 Then blnDo = True
        Else
            If InStr(GetInsidePrivs(p住院医嘱下达), "合理用药监测") > 0 Then blnDo = True
        End If
    Case PM_住院编辑
        If InStr(GetInsidePrivs(p住院医嘱下达), "合理用药监测") > 0 Then blnDo = True
    Case PM_护士校对
        If InStr(GetInsidePrivs(p住院医嘱发送), "合理用药监测") > 0 Then blnDo = True
    Case PM_住院首页
        blnDo = True
    Case PM_处方发药, PM_部门发药, PM_PIVA管理
        If InStr(GetInsidePrivs(lngModel), "合理用药监测") > 0 Then blnDo = True
    End Select
    
    PassCheckPrivs = blnDo
    WriteLog "clsPass", "PassCheckPrivs", "PassCheckPrivs_End"
End Function

Public Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
'说明:与frmDockInAdvice一并给药保持一致
    Dim i As Long, blnTmp As Boolean
    With gobjAdvice
        If .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, gobjCOL.intCOL相关ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, gobjCOL.intCOL相关ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, gobjCOL.intCOL相关ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Public Function InitAdviceRS(Optional ByVal bytFunc As Byte = 1) As ADODB.Recordset
'功能:构造医嘱记录
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '字段名称|字段类型|字段长度 缺省字段类型 为adVarChar
    Select Case bytFunc
    
    Case FUN_医嘱信息
        strFields = "医嘱ID||18,相关ID||18,医嘱期效||1,医嘱序号||5,医嘱状态||3,诊疗类别||3,开嘱科室||100,开嘱科室ID||18,开嘱医生编码||10,开嘱医生||100," & _
        "药品ID||18,药品名称||100,单次用量||16,单量单位||20,频率||50,用法||100,用法ID||18,开嘱时间||20,开始时间||20,结束时间||20,总量||16,总量单位||20," & _
        "用药目的||1,医生嘱托||100,警示|adInteger|1,滴速||100,审核状态|adInteger|1,处方号||30,处方序号||18,执行科室ID||18,天数||16"  '门诊发送
    Case FUN_输出内容
        strFields = "医嘱ID|adBigInt|18,药品名称||1000,是否禁忌|adInteger|1,禁忌药品说明||100,状态|adInteger|1"
    Case FUN_医嘱信息_DTBS
        strFields = "医嘱ID||18,相关ID||18,医嘱期效||1,医嘱序号||5,医嘱状态||3,诊疗类别||3,开嘱科室||100,开嘱科室ID||18,开嘱医生编码||10,开嘱医生||100," & _
        "诊疗项目ID||18,药品ID||18,药品名称||100,单次用量||16,单量单位||20,频率||50,用法||100,用法ID||18,开嘱时间||20,开始时间||20,结束时间||20,总量||16,总量单位||20," & _
        "用药目的||1,医生嘱托||100,警示|adInteger|1,天数||16,规格||100,频率编码||5,用药理由||1000,标志||1,离院带药|adInteger|1"
    Case FUN_医嘱信息_HZYY
        strFields = "医嘱ID||18,相关ID||18,医嘱期效||1,医嘱序号||5,医嘱状态||3,诊疗类别||3,开嘱科室||100,开嘱科室ID||18,开嘱医生ID||10,开嘱医生||100," & _
        "诊疗项目ID||18,药品ID||18,药品名称||100,单次用量||16,单量单位||20,频率||50,中药煎法||100,中药煎法ID||18,用法||100,用法ID||18,开嘱时间||20,开始时间||20,结束时间||20,总量||16,总量单位||20," & _
        "用药目的||50,医生嘱托||100,警示|adInteger|1,天数||16,规格||100,频率编码||5,用药理由||1000,标志||1,离院带药|adInteger|1,处方ID|adBigInt|18,滴速||100," & _
        "专业技术职务||50,输液||3"
    Case FUN_审查结果
        strFields = "警示值||3,医嘱ID||18"
    Case FUN_审查结果_HZYY
        strFields = "DrugName||100,DrugID||18,advice||1000,source||100,GroupNo||18,Type||200,Message||1000,Severity|adInteger|2,recipeId||18"
    Case FUN_医嘱信息_ZL
        strFields = "医嘱ID||18,诊疗项目ID||18,药品ID||18,本位码||50,输液组号||18,计量单位||20,单次量||20," & _
        "每日量||20,给药频次||50,给药频次名称||50,给药途径||18,性质|adInteger|2,给药途径名称||100,医嘱期效||1,开嘱时间||20,紧急标志||3," & _
        "开嘱医生||100,医生职称||100,医生抗菌药物等级||50,医嘱内容||1000,用药天数||16,剂型||50,药品抗菌药物等级||50," & _
        "毒理分类||50,超量说明||500,用药目的||50,药品禁忌等级||50,药品禁忌类型||50,药品禁忌说明||500,处方诊断||300,医嘱状态||3"
    Case FUN_审查结果_ZL
        strFields = "OrderId||18,Type||100,Level||50,DrugCode||50,Describ||2000,Remaks||4000,Light|adInteger|1,Tag|adInteger|2,WarnLevel|adInteger|1,Category|adInteger|1 "
    Case FUN_审查结果_YWS
        strFields = "Title||200,Detail||1000"
    Case FUN_反向问诊_ZL
        strFields = "Name||500,Type||100,Index||200,Value||1000,Class||200,Obsid||50,Default||200,ControlIndex|adInteger|3,Proid||18"
    Case FUN_药师审查_ZL
        strFields = "医嘱ID|adBigInt|18,相关ID|adBigInt|18,医嘱内容||1000,审查详情||1000,理由||1000,Tag|adInteger|1"
    Case FUN_病人状态_ZL
        strFields = "STATUS_ID||50,STATUS_NAME||100,STATUS_SITUATION||5"
    End Select
    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            If UCase(arrSubFeld(1) & "") = UCase("adVarChar") Then
                FieldType = adVarChar
            ElseIf UCase(arrSubFeld(1) & "") = UCase("adBigInt") Then
                FieldType = adBigInt
            ElseIf UCase(arrSubFeld(1) & "") = UCase("adInteger") Then
                FieldType = adInteger
            Else
                FieldType = adVarChar
            End If
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitAdviceRS = rs
End Function

Public Function GetDrugID(ByVal str诊疗项目ID As String) As Variant
'功能:返回药品ID或记录
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH
    arrTmp = Split(str诊疗项目ID, ",")
    If UBound(arrTmp) = 0 Then
        strSQL = "Select 药名ID,药品ID from 药品规格 where 药名id=[1] and rownum <2"
    ElseIf UBound(arrTmp) > 0 Then
        strSQL = "Select a.药名id, Max(a.药品id) As 药品id" & vbNewLine & _
        "From 药品规格 A" & vbNewLine & _
        "Where a.药名id In (Select * From Table(f_Num2list([1])))" & vbNewLine & _
        "Group By a.药名id"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", str诊疗项目ID)
    
    If UBound(arrTmp) = 0 Then
        If Not rs.EOF Then
            GetDrugID = rs!药品ID & ""
        Else
            GetDrugID = ""
        End If
    Else
        Set GetDrugID = rs
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get中药配方(ByVal str组IDs As String) As ADODB.Recordset
'功能:返回药品ID或记录
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select a.Id, a.相关id, a.医嘱期效, a.医嘱状态,a.诊疗类别,a.诊疗项目id,a.收费细目ID as 药品ID,a.医嘱内容 As 药品名称,a.序号,a.单次用量, d.计算单位 As 单量单位,a.执行频次 as 频率, a.间隔单位, a.频率次数, a.频率间隔, a.开始执行时间 As 开始时间," & vbNewLine & _
            "       a.执行终止时间 As 终止时间, a.开嘱时间, a.停嘱时间, c.名称 As 用法, c.Id As 用法id, a.执行性质, b.执行性质 As 组执行性质,a.开嘱医生,a.用药目的,a.天数, " & vbNewLine & _
            "       a.总给予量,f.住院单位 As 总量单位,f.门诊单位, a.医生嘱托, a.开嘱科室id,a.执行科室ID " & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱记录 B, 诊疗项目目录 C, 诊疗项目目录 D, 药品规格 F" & vbNewLine & _
            "Where a.相关id = b.Id And b.诊疗项目id = c.Id And a.诊疗项目id = d.Id And a.收费细目id = f.药品id(+) And" & vbNewLine & _
            "      a.相关id in (Select * From Table(f_Num2list([1]))) And a.诊疗类别 = '7'"


    Set rs = zlDatabase.OpenSQLRecord(strSQL, "中药组ID", str组IDs)

    Set Get中药配方 = rs

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get滴速(ByVal str组IDs As String) As ADODB.Recordset
'功能:返回药品ID或记录
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select a.Id, a.医生嘱托" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B " & vbNewLine & _
            "Where A.诊疗项目ID = B.ID And A.诊疗类别 ='E' And B.操作类型 = '2' And b.执行分类 = 1 And NVL(a.医生嘱托,'空') <> '空' And " & vbNewLine & _
            "      a.ID in (Select /*+cardinality(A,10)*/ * From Table(f_Num2list([1])) A) "


    Set rs = zlDatabase.OpenSQLRecord(strSQL, "医生嘱托", str组IDs)

    Set Get滴速 = rs

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRS(ByVal strTableName As String, ByVal strFileds As String, ByVal strInput As String, _
        Optional ByVal strWhere As String = "ID", Optional ByVal bytModel As Byte = 0, Optional ByVal bytType As Byte = 0) As Variant
'功能:返回指定表指定字段的记录集
'参数：strTableName-表名
'     strFileds
'     strInput 方式1(1个过滤条件)：ID1,ID2,...
'              方式2(2个过滤条件)：名称1,范围1;名称2,范围2;...
'             strSQL = "Select 编码, 名称, 适用范围" & vbNewLine & _
'                "From 诊疗频率项目" & vbNewLine & _
'                "Where (名称, 适用范围) In (Select /*+cardinality(B,10)*/" & vbNewLine & _
'                "                      C1, C2" & vbNewLine & _
'                "                     From Table(f_Str2list2('每天二次,1|每天三次,1', ';', ',')) B)"
'    bytModel=1 过滤条件为两列
'    当bytModel=1时： bytType=0-拆分列 C1,C2 同为字符串 =1-C1(Number),C2(Number);=2-C1(char),C2(Number);=3-C1(Number),C2(Char)
'    当bytModel=0时： bytType=0-f_num2list; bytType=1 f_Str2list


    Dim strSQL As String
    Dim strSub As String
    Dim strFun As String
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    If bytModel = 1 Then
        If bytType = 0 Then
            strSub = " C1,C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 1 Then
            strSub = " C1,C2 "
            strFun = "f_num2list2"
        ElseIf bytType = 2 Then
            strSub = "C1,To_Number(C2) As C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 3 Then
            strSub = " To_Number(C1) As C1,C2 "
            strFun = "f_Str2list2"
        End If
        strSQL = " Select  " & strFileds & vbNewLine & _
                " From  " & strTableName & vbNewLine & _
                " Where (" & strWhere & ") In (Select /*+cardinality(B,10)*/" & vbNewLine & _
                "                    " & strSub & vbNewLine & _
                "                     From Table(" & strFun & "([1], ';', ',')) B)"
    Else
        If bytType = 0 Then
            strFun = "f_num2list"
        ElseIf bytType = 1 Then
            strFun = "f_Str2list"
        End If
        arrTmp = Split(strInput, ",")
        If UBound(arrTmp) = 0 Or strInput = "" Then
            strSQL = "Select " & strFileds & "  From " & strTableName & " Where " & strWhere & " = [1]"
        ElseIf UBound(arrTmp) > 0 Then
            strSQL = "Select " & strFileds & vbNewLine & _
            "From " & strTableName & vbNewLine & _
            "Where " & strWhere & " In (Select /*+cardinality(A,10)*/ * From Table(" & strFun & "([1]))A )"
        End If
    End If
    Set GetRS = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", strInput)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AddDrugReason(ByRef objMap As Object, ByRef rsOut As ADODB.Recordset) As Boolean
'------------------------------------------------------------------------
'功能:禁忌药品添加禁忌说明
'参数:
'objMap-主窗体对象
'rsOut-输出对象
'返回:True-允许医嘱保存（不存在禁忌药品,无须填写禁忌说明;存在禁忌药品且完整填写禁忌说明）,False-禁止医嘱保存（存在禁忌药品且禁忌药品说明未完整填写）
'说明:中药配方禁忌说明保存在中药
'-----------------------------------------------------------------------
    Dim i As Long
    Dim strReason As String
    
    If rsOut Is Nothing Then AddDrugReason = True: Exit Function
    
    rsOut.Filter = "是否禁忌=1"
    
    For i = 1 To rsOut.RecordCount
        strReason = rsOut!禁忌药品说明 & ""
        Call zlCommFun.ShowMsgBox("禁忌说明", "^审查发现禁忌用药:【" & rsOut!药品名称 & "】" & _
            vbCrLf & vbCrLf & "必须录入禁忌用药说明才允许保存医嘱。^", "!确定(&O),?取消(&C)", objMap.frmMain, vbInformation, , , , , , "禁忌说明：", 99, strReason)
        If strReason = "" Then
            Exit Function
        Else
            rsOut!禁忌药品说明 = strReason
        End If
        rsOut.MoveNext
    Next
    AddDrugReason = True
End Function

Public Function ReadXML(ByVal strXML As String) As ADODB.Recordset
'功能:返回单个药品最大警示值
'xml模板
'    <his_results_xml fun_id="1006">
'    <result>
'       <type>ZDXGYWSY</type>
'       <level>2</level>
'       <prescA_hiscode>669</prescA_hiscode>
'       <mediA_hiscode>14686</mediA_hiscode>
'       <mediA_name>氨茶碱片</mediA_name>
'       <groupA>669</groupA>
'       <prescB_hiscode /><mediB_hiscode />
'       <mediB_name />
'        <groupB />
'    </result>
'    <result>
'    <type>XHZYWT</type>
'    <level>2</level>
'    <prescA_hiscode>669</prescA_hiscode>
'    <mediA_hiscode>14686</mediA_hiscode>
'    <mediA_name>氨茶碱片</mediA_name>
'    <groupA>669</groupA>
'    <prescB_hiscode>671</prescB_hiscode><mediB_hiscode>14250</mediB_hiscode>
'    <mediB_name>维生素C片</mediB_name>
'    <groupB>671</groupB>
'   </result>
'   <types>;ZDXGYWSY;XHZYWT;YHGXHCGYFYLWT_PC;YHGXHCGYFYLWT_DR;</types>
'</his_results_xml>


    Dim xmlDoc As DOMDocument
    Dim xmlRoot As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodes As IXMLDOMNodeList
    Dim rsRet As ADODB.Recordset
    
    Dim str警示值 As String
    Dim str医嘱ID As String
    
    On Error GoTo errH
    
    Set xmlDoc = New DOMDocument
    xmlDoc.loadXML (strXML)
    '如果不包含任何元素，则退出
    If xmlDoc.documentElement Is Nothing Then
        Set xmlDoc = Nothing
        Exit Function
    End If
    
    Set rsRet = InitAdviceRS(FUN_审查结果)
    '读取XML内容
    Set xmlRoot = xmlDoc.selectSingleNode("his_results_xml")
    Set xmlNodes = xmlRoot.selectNodes("result")

    If Not xmlNodes Is Nothing Then
        For Each xmlNode In xmlNodes
            str警示值 = xmlNode.selectSingleNode("level").Text
            If Val(str警示值) > 0 Then
                str医嘱ID = xmlNode.selectSingleNode("prescA_hiscode").Text
                If Val(str医嘱ID) <> 0 Then
                    rsRet.Filter = "医嘱ID ='" & str医嘱ID & "'"
                    If Not rsRet.EOF Then
                        If Val(rsRet!警示值 & "") < Val(str警示值) Then
                            rsRet!警示值 = str警示值
                        End If
                    Else
                        rsRet.AddNew
                        rsRet!警示值 = str警示值
                        rsRet!医嘱ID = str医嘱ID
                        rsRet.Update
                    End If
                End If
                str医嘱ID = xmlNode.selectSingleNode("prescB_hiscode").Text
                If Val(str医嘱ID) > 0 Then
                    rsRet.Filter = "医嘱ID ='" & str医嘱ID & "'"
                    
                    If Not rsRet.EOF Then
                        If Val(rsRet!警示值 & "") < Val(str警示值) Then
                            rsRet!警示值 = str警示值
                        End If
                    Else
                        rsRet.AddNew
                        rsRet!警示值 = str警示值
                        rsRet!医嘱ID = str医嘱ID
                        rsRet.Update
                    End If
                End If
            End If
        Next
    End If
    
    If rsRet.RecordCount > 0 Then rsRet.Filter = ""
    
    Set ReadXML = rsRet
    Exit Function
errH:
    MsgBox "ReadXML 错误号:" & Err.Number & "错误描述:" & Err.Description, vbOKOnly, gstrSysName
End Function

Public Function FuncGetDripInfo(ByVal lngIndex As Long, ByVal strDrip As String, ByVal lngPharmacyCode As Long, ByVal strPharmacyName As String, ByVal strDuration As String) As String
'功能:返回指定的JSON串
'字符串含义:
'{ "type":"druginfo","index":"drug001","driprate":"60","driptime":"120","pharmacycode":"药房编码","pharmacyname":"药房名称","duration":"用药天数"}
'driprate：   60   表示  每分钟60滴
'driptime：表示静滴所需要的时间，如果没有就传空串
'如果滴速是个区间值，就传最大的。
'单位为毫升 换算单位1毫升=20滴
    Dim strRet As String
    Dim arrTmp As Variant
    
    If InStr(strDrip, "滴/分钟") > 0 Then
        strDrip = Replace(strDrip, "滴/分钟", "")
        arrTmp = Split(strDrip, "-")
        If UBound(arrTmp) = 1 Then
            strDrip = arrTmp(1)
        Else
            strDrip = arrTmp(0)
        End If
         
    ElseIf InStr(strDrip, "毫升/小时") > 0 Then
        strDrip = Replace(strDrip, "毫升/小时", "")
        arrTmp = Split(strDrip, "-")
        If UBound(arrTmp) = 1 Then
            strDrip = arrTmp(1)
        Else
            strDrip = arrTmp(0)
        End If
        strDrip = (Val(strDrip) \ 60) * 20
    Else
        strDrip = ""
    End If
    strRet = "{""type"":""druginfo"",""index"":""" & lngIndex & """,""driprate"":""" & strDrip & """,""driptime"":""""," & _
            """pharmacycode"":""" & lngPharmacyCode & """,""pharmacyname"":""" & strPharmacyName & """,""duration"":""" & _
            strDuration & """}"
    FuncGetDripInfo = strRet
End Function

Public Function FuncGetOtherRecipInfo(ByVal strAdviceID As String, ByVal strRecipNo As String, ByVal strDrugCode As String, _
    ByVal strDrugName As String, ByVal strRouteName As String, ByVal strfrequency As String, ByVal strDoseunit As String, _
    ByVal strDosepertime As String, ByVal strNum As String, ByVal strNumUnit As String, ByVal strDuration As String) As String
    '功能:补充信息历史医嘱方法,返回指定的JSON串
    Dim strRet As String
' //历史医嘱信息
'        {
'            "type":"otherrecipinfo",
'            "hiscode":"his001",//字符串类型，机构编码
'            "index":" drug001",//字符串类型医嘱唯一码
'            "recipno":"MZ12376",//字符串类型，处方号
'            "drugsource":"USER",//字符串类型，药品类型
'            "druguniquecode":"123456",//字符串类型医嘱唯一码
'            "drugname":"阿莫西林胶囊",//字符串类型，药品名称
'            "routeCode"："口服" //字符串类型，给药途径编码用给药途径名称
'            "routeName"："口服"//字符串类型，给药途径名称
'            "routesource":"USER"
'            "frequency"："bid"//字符串类型，用药频次
'            "doseunit":"g"//字符串类型，表示每次使用剂量的用药单位
'            "dosepertime":"5"//字符串类型，表示每次使用剂量的数字部分
'            "num":"2"//字符串类型，药品开出数量，门诊处方审查专用，住院传空。
'            "numunit":"片"//字符串类型，药品开出数量单位，门诊处方审查专用，住院传空。
'            "duration":"7" // 持续用药天数
'        }

     strRet = "{""type"":""otherrecipinfo"",""hiscode"":""" & gstrHOSCODE & """,""index"":""" & strAdviceID & """,""recipno"":""" & strRecipNo & """," & _
            """drugsource"":""USER"",""druguniquecode"":""" & strDrugCode & """,""drugname"":""" & strDrugName & """,""routeCode"":""" & strRouteName & """," & _
            """routeName"":""" & strRouteName & """,""routesource"":""USER"",""frequency"":""" & strfrequency & """,""doseunit"":""" & strDoseunit & """," & _
            """dosepertime"":""" & strDosepertime & """,""num"":""" & strNum & """,""duration"":""" & strDuration & """}"
    FuncGetOtherRecipInfo = strRet
End Function


Public Function StrConvToNormal(ByVal strIn As String) As String
'功能：用StrConv(str,vbFromUnicode)转换时,有时会因为汉字乱码导致转换的xml串不是一个完整有效的XML串。
    Dim strChar As String
    Dim strRet As String
    Dim i As Long
    
    For i = 1 To Len(strIn)
        strChar = Mid(strIn, i, 1)
        If InStr(G_STR_MATCH & "=", strChar) > 0 Then
            strRet = strRet & strChar
        End If
    Next
    StrConvToNormal = strRet
End Function

Public Sub WriteLog(ByVal strModule As String, ByVal strFunction As String, ByVal strLog As String)
'------------------------------------------------
'功能：写入日志
'参数：
'      strModule  ：模块名
'      strFunction：功能名
'      strLog     ：日志内容
'备注：在系统选项中开启日志后，在写入日志时创建日志文件，每个进程每个日志创建只创建一次
'------------------------------------------------
    LogWrite "合理用药接口调试日志", strModule, strFunction, strLog
End Sub


'* ************************************** *
'* 模块名称：modCharset.bas
'* 模块功能：GB2312与UTF8相互转换函数
'* 作者：lyserver
'* ************************************** *

'- ------------------------------------------- -
'  函数说明：GB2312转换为UTF8
'- ------------------------------------------- -

Public Function GB2312ToUTF8(strIn As String, Optional ByVal ReturnValueType As VbVarType = vbString) As Variant
    Dim adoStream As Object

    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 2 'adTypeText
    adoStream.Open
    adoStream.WriteText strIn
    adoStream.Position = 0
    adoStream.Type = 1 'adTypeBinary
    GB2312ToUTF8 = adoStream.Read()
    adoStream.Close

    If ReturnValueType = vbString Then GB2312ToUTF8 = Mid(GB2312ToUTF8, 1)
End Function

'- ------------------------------------------- -
'  函数说明：UTF8转换为GB2312
'- ------------------------------------------- -
Public Function UTF8ToGB2312(ByVal varIn As Variant) As String
    Dim bytesData() As Byte
    Dim adoStream As Object

    bytesData = varIn
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 1 'adTypeBinary
    adoStream.Open
    adoStream.Write bytesData
    adoStream.Position = 0
    adoStream.Type = 2 'adTypeText
    UTF8ToGB2312 = adoStream.ReadText()
    adoStream.Close
End Function

Public Function WinHttpPost(ByVal strURL As String, ByVal strData As String, ByVal DataStic As DataEnum, Optional ByVal strHeader As String, Optional ByVal strMethod As String = "POST") As Variant
'支持HTTPS访问
'参数:strHeader 传值格式：HeaderName:HeaderValue 例子:CONTENT-TYPE:application/json
    Dim XMLHTTP As WinHttp.WinHttpRequest
    Dim DataS As String
    Dim DataB() As Byte
    Dim varHeader As Variant
    Dim varHeaderItem As Variant
    Dim i As Long

    On Error GoTo errH:
       
8      Set XMLHTTP = New WinHttpRequest
9      XMLHTTP.Open strMethod, strURL
10      If strHeader <> "" Then
            varHeader = Split(strHeader, ",")
            For i = LBound(varHeader) To UBound(varHeader)
                varHeaderItem = Split(varHeader(i), ":")
                XMLHTTP.setRequestHeader varHeaderItem(0), varHeaderItem(1)
            Next
        End If

13     XMLHTTP.send strData

110     Do Until XMLHTTP.Status = 200
112         DoEvents
        Loop

    '-----------------------------函数返回
114 Select Case DataStic
    Case responseText
        '--------------------------------直接返回字符串
116     DataS = XMLHTTP.responseText
118     WinHttpPost = DataS
120 Case responseBody
        '--------------------------------直接返回二进制
122     DataB = XMLHTTP.responseBody
124     WinHttpPost = DataS
126 Case responseBody + responseText
        '---------------------------二进制转字符串[直接返回字串出现乱码时尝试]
128     DataS = BytesToStr(XMLHTTP.responseBody)
130     WinHttpPost = DataS
132 Case Else
        '--------------------------------无效的返回
134     WinHttpPost = ""
    End Select

    '------------------------------------释放空间
136     Set XMLHTTP = Nothing

    Exit Function

errH:
138     WinHttpPost = ""
140     MsgBox "WinHttpPost失败！" & vbNewLine & "错误号:" & Err.Number & vbCrLf & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "中联软件"
End Function

'==========================================================
'| 模 块 名 | XMLHTTP
'| 说    明 | 替代Inet控件，实现数据通讯
'---------------------------------------------------------------------------《《Begin》》---------------------------------------------------------------------------------------
'==========================================================
Public Function HttpGet(ByVal Url As String, ByVal DataStic As DataEnum, Optional ByVal sngWaitTime As Single, Optional ByRef blnBreak As Boolean) As Variant
    Dim XMLHTTP As Object
    Dim DataS As String
    Dim DataB() As Byte
    Dim lngTime As Long
    On Error GoTo errH:
    
100 Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
102 XMLHTTP.Open "get", Url, True
104 XMLHTTP.send
    lngTime = Timer
106 Do While XMLHTTP.readyState <> 4
        If sngWaitTime = 0 Then
            If Timer - lngTime > gsngWaitTime Then blnBreak = True: Exit Function
        Else
            If Timer - lngTime > sngWaitTime Then blnBreak = True: Exit Function
        End If
108     DoEvents
    Loop
    blnBreak = False
    '--------------------------------------函数返回
110 Select Case DataStic
    Case responseText
        '--------------------------------直接返回字符串
112     DataS = XMLHTTP.responseText
114     HttpGet = DataS
116 Case responseBody
        '--------------------------------直接返回二进制
118     DataB = XMLHTTP.responseBody
120     HttpGet = DataB
122 Case responseBody + responseText
        '------------------------------二进制转字符串[直接返回字串出现乱码时尝试]
124     DataS = BytesToStr(XMLHTTP.responseBody)
126     HttpGet = DataS
128 Case Else
        '--------------------------------无效的返回
130     HttpGet = ""
    End Select

    '--------------------------------------释放空间
132 Set XMLHTTP = Nothing

    Exit Function

errH:
134 HttpGet = ""
136 MsgBox "HttpGet失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "中联软件"
End Function

Public Function HttpPost(ByVal strURL As String, ByVal strData As String, ByVal DataStic As DataEnum, _
    Optional ByVal strCONTENTTYPE As String, Optional ByVal strAUthorization As String, Optional ByVal sngWaitTime As Single, _
    Optional ByRef blnBreak As Boolean) As Variant
    Dim XMLHTTP As Object
    Dim DataS As String
    Dim DataB() As Byte
    Dim lngTime As Long
    
    On Error GoTo errH:

100 Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
102 XMLHTTP.Open "POST", strURL, True
104 XMLHTTP.setRequestHeader "Content-Length", Len(HttpPost)
    If strCONTENTTYPE = "" Then
106     XMLHTTP.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
    Else
        XMLHTTP.setRequestHeader "CONTENT-TYPE", strCONTENTTYPE  '"application/x-www-form-urlencoded; charset=utf-8"
    End If
    If strAUthorization <> "" Then
        XMLHTTP.setRequestHeader "AUthorization", strAUthorization
    End If
108 XMLHTTP.send (strData)
    lngTime = Timer
110 Do Until XMLHTTP.readyState = 4
        If sngWaitTime = 0 Then
            If Timer - lngTime > gsngWaitTime Then blnBreak = True: Exit Function
        Else
            If Timer - lngTime > sngWaitTime Then blnBreak = True: Exit Function
        End If
112     DoEvents
    Loop
    blnBreak = False
    '-----------------------------函数返回
114 Select Case DataStic
    Case responseText
        '--------------------------------直接返回字符串
116     DataS = XMLHTTP.responseText
118     HttpPost = DataS
120 Case responseBody
        '--------------------------------直接返回二进制
122     DataB = XMLHTTP.responseBody
124     HttpPost = DataS
126 Case responseBody + responseText
        '---------------------------二进制转字符串[直接返回字串出现乱码时尝试]
128     DataS = BytesToStr(XMLHTTP.responseBody)
130     HttpPost = DataS
    Case 6
        HttpPost = XMLHTTP.responseXML
132 Case Else
        '--------------------------------无效的返回
134     HttpPost = ""
    End Select

    '------------------------------------释放空间
136     Set XMLHTTP = Nothing

    Exit Function

errH:
138     HttpPost = ""
140     MsgBox "HttpPost失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, "中联软件"
End Function

Private Function BytesToStr(ByVal vInput As Variant) As String
    
    Dim strReturn       As String
    Dim i               As Long
    Dim intPrevCharCode As Integer
    Dim intNextCharCode As Integer

    For i = 1 To LenB(vInput)
        intPrevCharCode = AscB(MidB(vInput, i, 1))
        If intPrevCharCode < &H80 Then
            strReturn = strReturn & Chr(intPrevCharCode)
        Else
            intNextCharCode = AscB(MidB(vInput, i + 1, 1))
            strReturn = strReturn & Chr(CLng(intPrevCharCode) * &H100 + CInt(intNextCharCode))
            i = i + 1
        End If
    Next

    BytesToStr = strReturn
End Function

Public Function CreatePlugInOK() As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModel)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub


'自定义的消息处理函数
Public Function NewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'功能:捕获滚轮事件进行处理,非滚轮事件调用默认窗口消息处理函数
'参数:vsc-VScrollBar 对象
'     OldWindowProc 默认窗口消息处理函数地址
    On Error Resume Next
    If msg = WM_MOUSEWHEEL Then
        '对鼠标滚轮事件进行处理
        If wParam = -7864320 Then '向下滚动
            If frmPassAsk.vsc.Value - 10 < frmPassAsk.vsc.Max Then
                frmPassAsk.vsc.Value = frmPassAsk.vsc.Max
            Else
                frmPassAsk.vsc.Value = frmPassAsk.vsc.Value - 10
            End If
        ElseIf wParam = 7864320 Then '向上滚动
            If frmPassAsk.vsc.Value + 10 > frmPassAsk.vsc.Min Then
                frmPassAsk.vsc.Value = frmPassAsk.vsc.Min
            Else
                frmPassAsk.vsc.Value = frmPassAsk.vsc.Value + 10
            End If
        End If
    Else
        '调用默认窗口消息处理函数
        NewWindowProc = CallWindowProc(glngOldWindowProc, hWnd, msg, wParam, lParam)
    End If
End Function
