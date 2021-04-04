Attribute VB_Name = "mdlLCDShowPublic"
Option Explicit
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gobjPrintMode As Object
Public gobjReport As Object
'Public gcnOracle As New ADODB.Connection
Public gstrSysName As String                '系统名称
Public gstrSQL As String

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
Public Sub SaveErrLog(ByVal strInfo As String)
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strFile As String
    Dim rsTemp As New ADODB.Recordset
    
    strFile = "C:\zl9LCDShow_" & Format(gobjDatabase.Currentdate, "yyyyMMdd") & ".TXT"
    '检查文件是否存在，不存在则创建
    If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine strInfo & vbCrLf
    objText.Close
End Sub
'Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnMessage As Boolean = True) As Boolean
'    '------------------------------------------------
'    '功能： 打开指定的数据库
'    '参数：
'    '   strServerName：主机字符串
'    '   strUserName：用户名
'    '   strUserPwd：密码
'    '返回： 数据库打开成功，返回true；失败，返回false
'    '------------------------------------------------
'    Dim strError As String
'
'    On Error Resume Next
'    With cnOracle
'        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
'        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
'    End With
'    If err <> 0 Then
'        If blnMessage = True Then
'            '保存错误信息
'            strError = err.Description
'            If InStr(strError, "自动化错误") > 0 Then
'                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
'            ElseIf InStr(strError, "ORA-12154") > 0 Then
'                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
'            ElseIf InStr(strError, "ORA-12541") > 0 Then
'                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
'            ElseIf InStr(strError, "ORA-01033") > 0 Then
'                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
'            Else
'                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
'            End If
'        End If
'
'        err.Clear
'        OraDataOpen = False
'        Exit Function
'    End If
'    OraDataOpen = True
'End Function
'Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional cnTemp As ADODB.Connection)
''功能：打开记录集
'    If rsTemp.State = adStateOpen Then rsTemp.Close
'
'    Call gobjComLib.SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
'    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), IIf(cnTemp Is Nothing, gcnOracle, cnTemp), adOpenStatic, adLockReadOnly
'    Call gobjComLib.SQLTest
'End Sub
Public Sub SaveDebug(ByVal strInfo As String)
    '如果调试=1，表示提试调试信息,2-将调式信息写入文本；其它情况不输出调试信息
    '判断是否是调试状态，是则显示提示框

    '写文本文件
    '将调试信息写入文件中
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strFile As String, strExchange As String
    Dim rsTemp As New ADODB.Recordset
'    If gint调试 <> 0 Then
        If strExchange = "" Then strExchange = "C:\药房排队叫号"
        strFile = strExchange & "\调试信息_" & Format(Now, "yyyyMMdd") & ".TXT"
        '检查文件夹是否存在，不存在则创建
        If Not objFile.FolderExists(strExchange) Then objFile.CreateFolder strExchange
        '检查文件是否存在，不存在则创建
        If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
            
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        objText.WriteLine strInfo
        objText.Close
'    End If
End Sub
