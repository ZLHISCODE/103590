Attribute VB_Name = "mdlPass"
Option Explicit
'PASS接口函数，具体说明参见PASS接口文档描述
'说明：ShellRunAs.dll需要安装在程序或系统目录
'      DIFPassDll.dll为Pass系统自动获取并注册路径
'注册服务器
Public Declare Function RegisterServer Lib "ShellRunAs.dll" () As Integer
'PASS初始化
Public Declare Function PassInit Lib "DIFPassDll.dll" ( _
    ByVal UserName As String, _
    ByVal DepartMentName As String, _
    ByVal WorkstationType As Integer) As Integer
'PASS运行模式设置
Public Declare Function PassSetControlParam Lib "DIFPassDll.dll" ( _
    ByVal SaveCheckResult As Integer, _
    ByVal AllowAllegen As Integer, _
    ByVal CheckMode As Integer, _
    ByVal DisqMode As Integer, _
    ByVal UseDiposeIdea As Integer) As Integer
'传病人基本信息
Public Declare Function PassSetPatientInfo Lib "DIFPassDll.dll" ( _
    ByVal PatientID As String, _
    ByVal VisitID As String, _
    ByVal Name As String, _
    ByVal Sex As String, _
    ByVal Birthday As String, _
    ByVal Weight As String, _
    ByVal cHeight As String, _
    ByVal DepartMentName As String, _
    ByVal Doctor As String, _
    ByVal LeaveHospitalDate As String) As Integer
'传病人药品信息
Public Declare Function PassSetRecipeInfo Lib "DIFPassDll.dll" ( _
    ByVal OrderUniqueCode As String, _
    ByVal DrugCode As String, _
    ByVal DrugName As String, _
    ByVal SingleDose As String, _
    ByVal DoseUnit As String, _
    ByVal Frequency As String, _
    ByVal StartOrderDate As String, _
    ByVal StopOrderDate As String, _
    ByVal RouteName As String, _
    ByVal GroupTag As String, _
    ByVal OrderType As String, _
    ByVal OrderDoctor As String) As Integer
'设置需要进行单药警告的药品
Public Declare Function PassSetWarnDrug Lib "DIFPassDll.dll" (ByVal DrugUniqueCode As String) As Integer
'信息查询药品传入
Public Declare Function PassSetQueryDrug Lib "DIFPassDll.dll" ( _
    ByVal DrugCode As String, _
    ByVal DrugName As String, _
    ByVal DoseUnit As String, _
    ByVal RouteName As String) As Integer
'获取右键菜单是否可用值
Public Declare Function PassGetState Lib "DIFPassDll.dll" (ByVal QueryItemNo As String) As Integer
'PASS功能调用
Public Declare Function PassDoCommand Lib "DIFPassDll.dll" (ByVal CommandNo As Integer) As Integer
'获取药品警示级别
Public Declare Function PassGetWarn Lib "DIFPassDll.dll" (ByVal DrugUniqueCode As String) As Integer
'设置药品浮动窗口位置
Public Declare Function PassSetFloatWinPos Lib "DIFPassDll.dll" ( _
    ByVal Left As Integer, ByVal Top As Integer, _
    ByVal Right As Integer, ByVal Bottom As Integer) As Integer
'PASS退出函数
Public Declare Function PassQuit Lib "DIFPassDll.dll" () As Integer
'------------------------------------------------------------------
'ZLHIS中是否使用PASS系统
Public gblnPass As Boolean

Public Function PassInitialize() As Boolean
'功能：对PASS接口进行注册和初始化，同时检查PASS接口DLL是否正确安装
    On Error GoTo errH
    
    'PASS功能函数注册(共享客户端模式)
    If RegisterServer <> 0 Then
        MsgBox "PASS客户端注册失败，当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Exit Function
    End If
    
    'PASS初始化
    If PassInit(UserInfo.编号 & "/" & UserInfo.用户名, UserInfo.部门码 & "/" & UserInfo.部门名, 10) <> 1 Then
        MsgBox "PASS系统初始化失败，当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Exit Function
    End If
            
    'PASS是否可用检测
    If PassGetState("PassEnable") = 0 Then
        MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Call PassQuit: Exit Function
    End If
    
    'PASS应用模式设置(用默认值)
    Call PassSetControlParam(1, 2, 0, 2, 1)
    
    PassInitialize = True
    Exit Function
errH:
    If Err.Number = 53 And InStr(UCase(Err.Description), UCase("ShellRunAs.dll")) > 0 Then
        MsgBox "PASS接口文件 ShellRunAs.dll 不存在,可能合理用药监测系统未正确安装或配置。" & _
            vbCrLf & "在正确安装和配置合理用药监测系统之前，相应的功能不能使用。", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("DIFPassDll.dll")) > 0 Then
        MsgBox "PASS接口文件 DIFPassDll.dll 不存在,可能是因为以下原因：" & vbCrLf & _
            vbCrLf & "1.PASS客户端是第一次登录，请退出之后再重新登录即可正常使用。" & _
            vbCrLf & "2.合理用药监测系统未正确安装或配置，请仔细检查后再登录重试。", vbInformation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function
