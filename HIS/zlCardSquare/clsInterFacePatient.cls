VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterFacePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mobjPubPatient As Object
'保持属性值的局部变量
Private mvarblnInit As Boolean '局部复制
Private mvarblnRealName As Boolean '局部复制

Public Property Get blnRealName() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.blnRealName
    blnRealName = mvarblnRealName
End Property


Public Property Get blnInit() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.blnInit
    blnInit = mvarblnInit
End Property


Public Function Init(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
                    ByVal cnOracle As ADODB.Connection, ByVal strDBUser As String) As Boolean
    '初始化类模块变量
    On Error GoTo ErrHandler
    If CreatePublicPatient(lngSys, strDBUser, cnOracle) = False Then Exit Function
    mvarblnInit = True
    mvarblnRealName = mobjPubPatient.blnRealName
    Init = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsRealNamed(ByVal lng病人id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：检查病人是否已实名制认证
    '入参：lng病人ID 病人id
    '返回：True/False  已实名制认证返回True，否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    IsRealNamed = mobjPubPatient.IsRealNamed(lng病人id)
End Function

Public Function CheckPatiIdCard(ByVal strIdCard As String, Optional strBirthday As String, Optional strAge As String, Optional strSex As String, Optional strErrInfo As String, Optional datCalc As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：身份证号码合法性校验
    '入参：strIdCard 身份证号码
    '出参：strBirthday  函数返回True为出生日期
    '         strSex 函数返回True为性别
    '         strErrInfo 函数返回False为错误信息
    '         datCalc 计算日期 缺省则按系统时间计算
    '返回：True/False  身份证合法返回True(可从strBirthday，strSex获取出生日期和性别)，否则返回False(可从strErrInfo获取详细错误信息)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    CheckPatiIdCard = mobjPubPatient.CheckPatiIdCard(strIdCard, strBirthday, strAge, strSex, strErrInfo, datCalc)
End Function

Public Function CheckPatiAge(ByVal strAge As String, Optional ByVal strBirthday As String = "", _
    Optional ByVal bytTag As Byte = 0, Optional ByVal strCalcDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人年龄检查
    '入参:
    '   strAge:病人年龄
    '   strBirthDay:出生日期
    '   bytTag:对于zl_Age_Check函数返回的询问类型的信息，是否要强制终止，还是保持询问.0-保持询问,1-禁止
    '   strCalcDate:计算日期,缺省按照当前时间计算
    '返回：TRUE或FALSE，TRUE:继续,FALSE:终止
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    CheckPatiAge = mobjPubPatient.CheckPatiAge(strAge, strBirthday, bytTag, strCalcDate)
End Function

Public Function RecalcBirthDay(ByVal strAge As String, ByRef strDateOfBirth As String, Optional ByVal strCalcDate As String, Optional ByRef strMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人年龄获取病人出生日期
    '入参:strAge:病人年龄,如：23岁、1岁2月
    'strCalcDate-传入计算日期
    '返回:传入的病人年龄格式正确则计算返回出生日期,否则返回空
    '  strMsg-返回警告信息
    '正确年龄格式:X岁[X月]、X月[X天]、X天、X小时[X分钟]
    '    X岁:X不能大于200,X月:X不能大于12,X天:X不能大于31,X小时:X不能大于24,X分钟:X不能大于59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    RecalcBirthDay = mobjPubPatient.RecalcBirthDay(strAge, strDateOfBirth, strCalcDate, strMsg)
End Function

Public Function SavePatiBaseInfo(ByVal lng病人id As Long, ByVal lng就诊ID As Long, ByVal strName As String, ByVal strSex As String, _
    ByVal strAge As String, ByVal strBirthday As String, ByVal str模块 As String, Optional ByVal int场合 As Integer = 1, Optional strInfo As String = "", _
    Optional ByVal blnXWHIS As Boolean = True, Optional ByVal blnEMPI As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：调整病人基本信息(含业务数据的同步调整)
    '入参：lng病人ID-病人ID (不能为空/0)
    '         lng就诊ID-挂号ID或主页ID(可为0)
    '         strName-姓名 (不能为空)
    '         strSex-性别 (不能为空)
    '         strAge-年龄 (不能为空)
    '         strBirthDay-出生日期 (不能为空)
    '         str模块-调用该功能的模块描述，如"门诊挂号"，"检查报到"。
    '         int场合 1-门诊;2-住院(lng就诊ID=0,则默认为1;lng就诊ID<>0,1-lng就诊ID为挂号ID,2-lng就诊ID为主页ID)
    '         blnXWHIS-基本信息调整时是否调用RIS的接口 缺省=True调用（该参数用于避免病人信息中重复调用RIS接口）
    '         blnEMPI-T EMPI平台已经建档，F-EMPI平台未建档
    ' 出参：strInfo:更新成功-信息调整导致的变化信息(返回True); 更新失败-信息调整未成功的原因
    ' 返回：TRUE OR False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    SavePatiBaseInfo = mobjPubPatient.SavePatiBaseInfo(lng病人id, lng就诊ID, strName, strSex, strAge, strBirthday, str模块, int场合, strInfo, blnXWHIS, blnEMPI)
End Function

Public Function PatiImageGatherer(frmMain As Object, ByRef strPictureFilePath As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人图片采集
    '入参:frmMain-调用主窗体
    '出参:strPictureFilePath-返回采集图片的路径(包含文件名及后缀)
    '返回:TRUE或False
    '编制:刘鹏飞
    '日期:2014-07-16
    '调用示例:此接口返回TRue。可直接使用LoadPicture(strPictureFilePath)加载图片
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    PatiImageGatherer = mobjPubPatient.PatiImageGatherer(frmMain, strPictureFilePath)
End Function

Private Function CreatePublicPatient(ByVal lngSys As Long, ByVal strDBUser As String, _
                    ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建zlPublicPatient部件
    '返回:创建成功,返回True,否则返回False
    '编制:冉俊明
    '日期:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "病人信息公共部件（zlPublicPatient）创建失败！", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(cnOracle, lngSys, strDBUser) = False Then
            MsgBox "病人信息公共部件（zlPublicPatient）初始化失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreatePublicPatient = True
End Function

Private Sub Class_Terminate()
    mvarblnInit = False
    If Not mobjPubPatient Is Nothing Then Set mobjPubPatient = Nothing
End Sub

Public Function CheckIsRealNameColumn(ByVal lng病人id As Long, ByVal strInputItem As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据输入项目判断，是否实名认证列
    '入参:strInputItem-当前输入项目
    '返回:是实名认证返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-26 16:53:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnInit = False Then Exit Function
    CheckIsRealNameColumn = mobjPubPatient.CheckIsRealNameColumn(lng病人id, strInputItem)
End Function
