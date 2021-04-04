Attribute VB_Name = "mdlMain"
Option Explicit

'部件注册类型
Public Enum RegFileType
    RFT_NotReg = 0                  '不注册的对象
    RFT_NormalReg = 1               '常规注册，自动识别.NET部件，.NET部件通过Regasm注册，其他通过调用DLLRegServer注册
    RFT_NETGAC = 2                  'NET程序集注册，通过gacutil注册到全局程序集缓存
    RFT_NETServer = 3               'NET服务注册，通过installUtil进行安装卸载。
    RFT_NETComReg = 4               '.NET Com部件注册，通过调用Regasm完成
    RFT_VBComReg = 5                '通过手写注册表注册
    RFT_DelphiComReg = 6            'DelphiCom注册，通过DLLRegServer注册
    RFT_PBComReg = 7                'PBCom注册，通过DLLRegServer注册
End Enum

Public gobjFSO              As New FileSystemObject     '文件操作对象
Public gobjTrace            As New clsTrace             '日志跟踪对象
Public gstrGACPath          As String                   'GACUTIL.EXE路径
Public gstr7ZPath           As String                   '7z.exe文件路径
Public gblnIs64Bits         As Boolean                  '是否是64位系统
Public gclsRegCom           As New clsRegCom            '部件注册对象
Public gstrAPPPath          As String
Public gstSysPath           As String
Public gobj7z               As New cls7zZip

Sub Main()
    Dim strErr As String
    Call InitInstall
    If Not InstallOO4O(strErr) Then
        MsgBox "OO4O组件安装失败。信息：" & strErr, vbInformation, "中联软件"
    Else
        MsgBox "OO4O组件安装成功。", vbInformation, "中联软件"
    End If
End Sub

Private Function InitInstall() As Boolean
    '安装包是否存在
    If IsDesinMode Then
        gstrAPPPath = "C:\APPSOFT"
    Else
        gstrAPPPath = gobjFSO.GetParentFolderName(App.Path)
    End If
    gstrGACPath = gstrAPPPath & "\Public\gacutil.exe"
    gblnIs64Bits = Is64bit
    gstSysPath = gobjFSO.GetSpecialFolder(SystemFolder)
    If gblnIs64Bits Then
        gstSysPath = gobjFSO.GetParentFolderName(gstSysPath) & "\SysWOW64"
    End If
    gstr7ZPath = gstSysPath & "7z.exe"
    gobj7z.Init7zZip (gstr7ZPath)
    Call gobjTrace.OpenTace("OO4O", gstrAPPPath)
End Function

