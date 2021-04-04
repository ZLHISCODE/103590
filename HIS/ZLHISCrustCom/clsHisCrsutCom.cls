VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHisCrsutCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mobjHisCrustCom     As Object
Private mpbSend             As PropertyBag

Public Function InitHisCrustCom(ByRef objHisCrust As Object, ByVal strAppPath As String, ByVal strCommand As String, ByVal strLogFile As String) As Boolean
'功能：初始化部件
'objHisCrust=自动升级外科的代理对象，用来调用自动升级来处理一些事情。
'strAppPath=自动升级程序路径
'strCommand=自动升级命令行
'是否初始化成功
    Set mobjHisCrustCom = objHisCrust
    If IsDesinMode Then
        '调试模式存在BUG
        gstrAppPath = gobjFSO.GetParentFolderName(strAppPath)
    Else
        gstrAppPath = strAppPath
    End If
    gstrCommand = strCommand
    gobjTrace.LogFile = strLogFile
    gobjTrace.WriteSection "开始使用ZLHISCRUSTCOM.DLL升级", SL_LevelThree
    InitHisCrustCom = InitCommon
    gobjTrace.WriteInfo "InitHisCrustCom", "END", True
End Function

Public Function UpdateBaseFile(Optional ByVal blnUpgrade As Boolean = True) As Boolean
    On Error GoTo ErrH
    If gotCurType <> OT_CheckFile Then
        If Not UpgradeBase() Then
            Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级失败 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
            Call SetOperateProcess(gotCurType, OS_Failure, SumErrMsg)          '标识升级结束
            GoTo ReCall                      '强制退出进程
        End If
    End If
    gobjTrace.WriteInfo "UpdateBaseFile", "END", True
    Call gobjTrace.CloseLog
    UpdateBaseFile = True
    Exit Function
ErrH:
    If Not gblnHelperMain Then MsgBox Err.Description, vbInformation, App.Title
    Err.Clear
ReCall:
    Call CallHISEXE
    Call gobjMe.ExitApp
    Exit Function
AutoLogin:
    Call CallHISEXE(True)
    Call gobjMe.ExitApp
    Exit Function
End Function

Public Function UpdateFile() As Boolean
    On Error GoTo ErrH
    gobjTrace.LogFile = gobjTrace.LogFile
    gobjTrace.WriteInfo "UpdateFile", "BGN", True
    If Not GetUpgradeFileList Then
        Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级失败 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        Call SetOperateProcess(gotCurType, IIf(gotCurType = OT_CheckFile, OS_Failure, OS_NotInProcessing), SumErrMsg)          '标识升级结束
        GoTo ReCall
    End If
    If grsFileUpgrade.RecordCount = 0 Then
        Call RecordErrMsg(MT_InitEnv, "文件清单获取", "没有可升级的文件，系统自动退出。")
        Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级完成 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        Call SetOperateProcess(gotCurType, OS_Completed, SumErrMsg, glngFileBatch)          '标识升级结束
        GoTo ReCall
    End If
    Call GetKILLProcess
    Call frmHisCrust.tmrStart_Timer
    gobjTrace.WriteInfo "UpdateFile", "END", True
    Call gobjTrace.CloseLog
    UpdateFile = True
    Exit Function
ErrH:
    If Not gblnHelperMain Then MsgBox Err.Description, vbInformation, App.Title
    Err.Clear
ReCall:
    If Not gblnHelperMain Then Call CallHISEXE
    Call gobjMe.ExitApp
    Exit Function
AutoLogin:
    If Not gblnHelperMain Then Call CallHISEXE(True)
    Call gobjMe.ExitApp
    Exit Function
End Function

Public Sub ExitApp()
'功能：调用升级程序来中止程序运行
    Call mobjHisCrustCom.ExitApp
End Sub

Public Function SendCommand(ByVal strCommand As String, ByVal strParas As String, Optional ByRef pbSend As PropertyBag, Optional ByRef pbReturn As PropertyBag) As Boolean
'功能：调用升级程序来执行命令,保持接口兼容
'strCommand:执行的命令
'strParas:执行命令的参数
'pbSend：附加传送数据
'pbReturn:附加返回数据
'返回：命令是否执行成功
    If pbSend Is Nothing Then
        SendCommand = mobjHisCrustCom.ReceiveCommand(strCommand, strParas, mpbSend, pbReturn)
        Set mpbSend = Nothing
    Else
        SendCommand = mobjHisCrustCom.ReceiveCommand(strCommand, strParas, pbSend, pbReturn)
    End If
    
End Function

Public Function ReceiveCommand(ByVal strCommand As String, Optional ByVal strParas As String, Optional ByRef pbSend As PropertyBag, Optional ByRef pbReturn As PropertyBag) As Boolean
'功能：升级程序的DLL执行命令，方便扩展,保持接口兼容
'strCommand:执行的命令
'strParas:执行命令的参数
'pbSend：附加传送数据
'pbReturn:附加返回数据
'返回：命令是否执行成功
'    Dim pbReturn As New PropertyBag
'
'
'    pbReturn.WriteProperty "返回的名称1", "返回的值1"
'    pbReturn.WriteProperty "返回的名称2", "返回的值2"

End Function


Public Sub AddSendValue(ByVal strKeyName As String, varValue As Variant)
'功能：将变量加入到数据传送包
'strKeyName:数据关键字
'varValue:数据内容
    If mpbSend Is Nothing Then
        Set mpbSend = New PropertyBag
    End If
    mpbSend.WriteProperty strKeyName, varValue
End Sub

Private Sub Class_Initialize()
    Set gobjMe = Me
End Sub

Private Sub Class_Terminate()
    Call gobjTrace.CloseLog
    Set gobjTrace = Nothing
    Set gobjMe = Nothing
    Set grsFileUpgrade = Nothing
    Set gcnOracle = Nothing
    Set gobjFSO = Nothing
    Set gcllSetPath = Nothing
    Set gclsRegCom = Nothing
    Set grsErrRec = Nothing
    Set gobj7zZip = Nothing
    Set gclsConnect = Nothing
End Sub

