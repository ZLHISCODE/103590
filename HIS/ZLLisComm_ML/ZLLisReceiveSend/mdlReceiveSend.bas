Attribute VB_Name = "mdlReceiveSend"
Option Explicit

Public gstrBuffFile As String  '接收数据的存放文件名
Public gstrIniFile As String   '配置文件名
Public gstrLockFile As String  '锁定文件名
Public gstrRAWDIR As String    '原始数据目录
Public gstrResultDIR As String '解析结果存放目录
Public gstrSendDir As String   '待发送指令存放目录
Public gstrGamDir As String    '图像结果存放目录

Public gFileObject As New FileSystemObject  '公共文件系统对象，用于文件目录相关操作
Public gobjLisDev As Object                 '具体的通讯程序

Public Type T仪器设置
    
    类型       As Integer  '0-COM口方式 1-IP方式
    'Com
    COM端口       As Integer
    波特率     As Long
    数据位     As String
    校验位     As String
    停止位     As String
    握手       As String
    缓冲大小   As Long
    
    'TCP/TP
    IP端口     As Long
    IP         As String
    主机       As Long
    
    '公共
    字符模式   As String
    自动应答   As String   '自动应答间隔，单位秒，为<=0时不启用。
    通讯周期   As String
    通讯程序   As String
End Type
Public g仪器设置 As T仪器设置   '保存仪器通讯设置

'读写ini 文件的API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Sub Main()
    If Init = True Then
        frmMain.Show
    End If
End Sub

Public Sub ReadSet()
    '读取INI文件，获得通讯参数
    g仪器设置.类型 = CInt(Val(ReadIni("RECEIVE_SET", "类型", gstrIniFile)))
    g仪器设置.COM端口 = CInt(Val(ReadIni("RECEIVE_SET", "COM端口", gstrIniFile)))
    g仪器设置.IP = ReadIni("RECEIVE_SET", "IP", gstrIniFile)
    g仪器设置.IP端口 = CLng(Val(ReadIni("RECEIVE_SET", "IP端口", gstrIniFile)))
    g仪器设置.波特率 = CLng(Val(ReadIni("RECEIVE_SET", "波特率", gstrIniFile)))
    g仪器设置.缓冲大小 = CLng(Val(ReadIni("RECEIVE_SET", "缓冲大小", gstrIniFile)))
    g仪器设置.数据位 = ReadIni("RECEIVE_SET", "数据位", gstrIniFile)
    g仪器设置.停止位 = ReadIni("RECEIVE_SET", "停止位", gstrIniFile)
    g仪器设置.握手 = ReadIni("RECEIVE_SET", "握手", gstrIniFile)
    g仪器设置.校验位 = ReadIni("RECEIVE_SET", "校验位", gstrIniFile)
    g仪器设置.主机 = CLng(Val(ReadIni("RECEIVE_SET", "主机", gstrIniFile)))
    g仪器设置.自动应答 = ReadIni("RECEIVE_SET", "自动应答", gstrIniFile)  '应答字符从接口中取
    g仪器设置.字符模式 = ReadIni("RECEIVE_SET", "字符模式", gstrIniFile)
    g仪器设置.通讯程序 = ReadIni("RECEIVE_SET", "通讯程序", gstrIniFile)
    
    g仪器设置.通讯周期 = Val(ReadIni("RECEIVE_SET", "通讯周期", gstrIniFile))
    If Not (Val(g仪器设置.通讯周期) > 0.1 And Val(g仪器设置.通讯周期) < 600) Then g仪器设置.通讯周期 = 0.5

End Sub

Private Function Init() As Boolean
    Dim strPath As String
    
    On Error GoTo errH
    
    gstrIniFile = App.Path & "\ReceiveSend.ini"
    If Not gFileObject.FileExists(gstrIniFile) Then
        MsgBox "无通讯配置文件“" & gstrIniFile & "”，程序不能运行！", vbQuestion, "通讯程序"
        Exit Function
    Else
        '创建通讯锁定文件，避免重复启动
        Dim TsTmp As TextStream
        
        gstrLockFile = App.Path & "\Lock.txt"

        If gFileObject.FileExists(gstrLockFile) And App.PrevInstance = True Then
            If Dir(gstrSendDir & "\CloseEnd.txt") = "" Then
                MsgBox "程序不能重复运行！", vbQuestion, "通讯程序"
                Exit Function
            End If
        Else
            Set TsTmp = gFileObject.CreateTextFile(gstrLockFile, True)
            TsTmp.WriteLine "启动：" & Format(Now, "yyyy-MM-dd HH:mm:ss")
            TsTmp.Close
            Set TsTmp = Nothing
        End If
        If gFileObject.FileExists(gstrSendDir & "\CloseEnd.txt") Then gFileObject.DeleteFile gstrSendDir & "\CloseEnd.txt"
    End If
    
    '创建相关目录
    '    RAW-原始数据,Result-解码结果,Gam-图像结果,Send-发送数据
    gstrRAWDIR = App.Path & "\Raw"
    If Not gFileObject.FolderExists(gstrRAWDIR) Then Call gFileObject.CreateFolder(gstrRAWDIR)
    
    gstrResultDIR = App.Path & "\Result"
    If Not gFileObject.FolderExists(gstrResultDIR) Then Call gFileObject.CreateFolder(gstrResultDIR)
    
    gstrGamDir = App.Path & "\Gam"
    If Not gFileObject.FolderExists(gstrGamDir) Then Call gFileObject.CreateFolder(gstrGamDir)
    
    gstrSendDir = App.Path & "\Send"
    If Not gFileObject.FolderExists(gstrSendDir) Then Call gFileObject.CreateFolder(gstrSendDir)
    
    Init = True
    Exit Function
errH:
    MsgBox "初始化程序时出现错误" & vbNewLine & Err.Description, vbQuestion, "通讯程序"
    
End Function

Public Function ReadIni(strItem As String, strKey As String, strPath As String) As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = Replace(GetStr, Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
    On Error GoTo errH
    WriteIni = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    WriteIni = False
End Function



Public Sub WriteErrLog(ByVal strFunc As String, ByVal StrInput As String, ByVal strOutput As String)
    '------------------------------------------------------
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------
    
    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFilename As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
'    If Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv", "清空接收日志", 1)) = 1 Then
'        If Dir(App.Path & "\调试.TXT") = "" Then Exit Sub
'    End If
    strFilename = App.Path & "\错误日志_" & Format(Date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strFilename) Then Call objFileSystem.CreateTextFile(strFilename)
    Set objStream = objFileSystem.OpenTextFile(strFilename, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (String(50, "≡"))
    objStream.WriteLine ("执行时间:" & strDate & "版本:" & App.Major & "." & App.Minor & "." & App.Revision)
    objStream.WriteLine ("驱动:" & strFunc)
    objStream.WriteLine ("  :" & StrInput)
    objStream.WriteLine ("  :" & strOutput)
    'objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
End Sub

