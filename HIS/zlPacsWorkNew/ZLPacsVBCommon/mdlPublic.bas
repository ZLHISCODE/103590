Attribute VB_Name = "mdlPublic"
Option Explicit

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpvalueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function OEMViewOpen Lib "InterCOM.dll" (ByVal lPlanID As Long, ByVal cpFilter As String, ByVal lFunc As Long, ByVal cpReserved As String) As Long
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Public Const WM_COPYDATA As Long = &H4A
Public Const HKEY_CURRENT_USER = &H80000001

'模块号常量定义
Public Const G_LNG_XWPACSVIEW_MODULE As Long = 1288     'XWPACS编号
Public Const G_LNG_PACSSTATION_MODULE As Long = 1290    '影像医技系统编号
Public Const G_LNG_VIDEOSTATION_MODULE As Long = 1291   '影像采集系统编号
Public Const G_LNG_PATHSTATION_MODULE As Long = 1294    '影像病理系统编号

'保存消息内容的结构
Public Type TGetImgMsg
    strSubDir As String          '图像所在的子目录
    strDestMainDir As String            '复制图像的目的目录，本机目录
    strIP As String                 '图像服务器的IP地址
    strFtpDir As String             'FTP目录
    strFTPUser As String            'FTP用户名
    strFTPPswd As String            'FTP密码
    strSDDir As String              '共享目录名称
    strSDUser As String             '共享目录用户名
    strSDPswd As String             '共享目录密码
    blnEnable As Boolean            '本消息可用
End Type

'进程间传递内存空间，可以传字符串
Public Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

' 网络资源
Public Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
Public Const RESOURCE_PUBLICNET = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const CONNECT_UPDATE_PROFILE = &H1

Public gcnOracle As New ADODB.Connection
Public Const gstrSysName = "签名"

Public gobjComlib As Object
Public zlDatabase As Object


Public Sub InitComLib(cnOracle As ADODB.Connection, strDbUser As String)
    If gobjComlib Is Nothing Then
        Set gobjComlib = GetObject("", "zl9ComLib.clsComLib")
        If gobjComlib Is Nothing Then Set gobjComlib = CreateObject("zl9ComLib.clsComLib")
        
        gobjComlib.InitCommon cnOracle
        
        Set zlDatabase = gobjComlib.zlDatabase
    End If
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function funcConnectShardDir(frmParent As Form, strShareRemoteDir As String, strUserName As String, _
    strPassWord As String) As Long
'------------------------------------------------
'功能：创建网络资源
'参数： frmParent  -- 父窗体
'       strShareRemoteDir -- 共享目录
'       strUserName -- 共享目录用户名
'       strPassWord -- 共享目录密码
'返回：无，连接共享目录
'------------------------------------------------
    
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "网络连接失败，请检查网络设置是否正确！", vbExclamation
    End If
    funcConnectShardDir = lngResult
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function

Public Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String)
'------------------------------------------------
'功能：记录通讯日志
'参数： logSubName  --  产生日志的函数名
'       logTitle   -- 日志名称
'       logDesc   --  日志内容
'返回：无
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    Dim intHour As Integer
    Dim lngAppSoft As Long
    
    On Error GoTo err
    
    '每隔4小时产生一个日志文件
    intHour = Hour(Time)
    intHour = intHour / 4
    intHour = intHour * 4
    
    strFileName = App.Path
    
    lngAppSoft = InStr(UCase(strFileName), "APPSOFT")
    If lngAppSoft > 0 Then
        strFileName = Mid(strFileName, 1, lngAppSoft + 7 - 1) & "\Log\日志跟踪\Pacs_VBCommon接口调用"
    End If
    
    '如果日志路径不存在，则创建
    If Dir(strFileName, vbDirectory) = "" Then
        Call MkLocalDir(strFileName)
    End If
    
    strFileName = strFileName & "\VBCommon接口调用_" & Format(date, "yyyymmdd") & "_" & intHour & ".log"
    
    strLog = Now() & " 标题： " & logTitle & vbCrLf & "      函数： " & logSubName & vbCrLf & "     日志内容：" & logDesc & vbCrLf
    
    Open strFileName For Append As #1
    Print #1, strLog
    Close #1
    
    Exit Sub
err:
    Close #1
End Sub

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
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

'################################################################################################################
'## 功能：  密码转换函数
'##
'## 参数：  strOld  :原密码
'##
'## 返回：  加密生成的密码
'################################################################################################################
Public Function TranPasswd(strOld As String) As String
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew
End Function

Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub
