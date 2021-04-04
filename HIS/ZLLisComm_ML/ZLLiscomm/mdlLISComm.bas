Attribute VB_Name = "mdlLISComm"
Option Explicit

'Public gcnOracle As ADODB.Connection    '公共数据库连接
Public gstrSQL As String

'Public gstrSysName As String                '系统名称

Public glngExeDeptID As Long '执行科室
Public ParentWnd As Object
Public blnDataReceived As Boolean
'------任务栏图标处理
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_ACTIVATE = &H6
Public Const WM_KEYDOWN = &H100
Public Const WM_PAINT = &HF

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'Public Const GWL_EXSTYLE = (-20)
'Public Const WinStyle = &H40000
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4

'酶标仪参数
Public glngMBDeviceID As Long, gstrMBChannel As String, glngMBNo As Long, gstrMBPosition As String

Private mItem() As Variant

Public Const LOG_错误日志 = 0
Public Const LOG_通讯日志 = 1
Public Const LOG_未知项 = 2

Public pLast错误日志 As String '上次错误信息,用于避免输出重复的日志
Public pLast通讯日志 As String
Public mMakeNoRule As String    '标本序号生成时间规则

Public gblnFromDB As Boolean ' 是否是从数据库读取参数.

Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Public Sub SavePortsSetting()
'功能：保存连接检验仪器的串口设置
    Dim i As Integer
    Dim strSet As String
    Dim aPorts As Variant
    On Error GoTo errH
    
    strSet = ""
    If gblnFromDB Then
        '清空原来的设置
        Call gobjDatabase.SetPara("本机连接仪器", "", glngSys, 1208)
        For i = LBound(g仪器) To UBound(g仪器)
            '仪器id , 类型, COM口, 波特率, 数据位, 校验位, 停止位, 握手, TCPIP端口, IP地址, 字符模式, 另存为的仪器ID, 主机,自动应答,可发已核标本,通讯目录,自动审核人,自动计算质控,另存为通道码
            If g仪器(i).ID > 0 Then
                strSet = strSet & ";" & g仪器(i).ID & "," & g仪器(i).类型 & "," & g仪器(i).COM口 & "," & g仪器(i).波特率 & _
                   "," & g仪器(i).数据位 & "," & g仪器(i).校验位 & "," & g仪器(i).停止位 & "," & g仪器(i).握手 & _
                   "," & g仪器(i).IP端口 & "," & g仪器(i).IP & "," & g仪器(i).字符模式 & "," & g仪器(i).SaveAsID & "," & g仪器(i).主机 & _
                   "," & g仪器(i).自动应答 & "," & g仪器(i).可发已核标本 & "," & g仪器(i).通讯目录 & "," & g仪器(i).自动审核人 & "," & g仪器(i).自动计算质控 & "," & g仪器(i).另存为通道码
            
            
                If Dir(g仪器(i).通讯目录 & "\ReceiveSend.ini") <> "" Then Kill g仪器(i).通讯目录 & "\ReceiveSend.ini"
            End If
        Next
        If strSet <> "" Then
            Call gobjDatabase.SetPara("本机连接仪器", strSet, glngSys, 1208)
        End If
    Else
        'DeleteSetting "ZLSOFT", "公共模块", "ZlLISSrv"
        Err = 0: On Error Resume Next
        aPorts = GetAllSettings("ZLSOFT", "公共模块\ZlLISSrv")
        On Error GoTo errH
        If IsEmpty(aPorts) Then
            ReDim aPorts(8, 0)
            For i = 0 To 7
                aPorts(i, 0) = "COM" & i + 1
            Next
        End If
        Err = 0: On Error Resume Next
        For i = LBound(aPorts) To UBound(aPorts)
            DeleteSetting "ZLSOFT", "公共模块\ZLLISSrv", aPorts(i, 0)
            DeleteSetting "ZLSOFT", "公共模块\ZLLISSrv\" & aPorts(i, 0)
        Next
        On Error GoTo errH
        For i = LBound(g仪器) To UBound(g仪器)
            If g仪器(i).类型 = 1 Then
                'TCP
                If g仪器(i).ID > 0 Then
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv", "IP" & g仪器(i).ID, "")
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Device", g仪器(i).ID)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Enabled", g仪器(i).类型)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Host", g仪器(i).主机)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "InMode", g仪器(i).字符模式)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "IP", g仪器(i).IP)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Port", g仪器(i).IP端口)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "SaveAs", g仪器(i).SaveAsID)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "Auto", g仪器(i).自动应答)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "blnSend", g仪器(i).可发已核标本)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "ReceiveDir", g仪器(i).通讯目录)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "AutoCheckMan", g仪器(i).自动审核人)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "AutoQCCalc", g仪器(i).自动计算质控)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\IP" & g仪器(i).ID, "SaveAsTonDao", g仪器(i).另存为通道码)
                End If
            Else
                If g仪器(i).COM口 > 0 And g仪器(i).ID > 0 Then
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv", "COM" & g仪器(i).COM口, "")
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "Device", g仪器(i).ID)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "Speed", g仪器(i).波特率)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "DataBit", g仪器(i).数据位)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "Parity", g仪器(i).校验位)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "StopBit", g仪器(i).停止位)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "HandShaking", g仪器(i).握手)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "InputMode", g仪器(i).字符模式)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "SaveAs", g仪器(i).SaveAsID)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "Auto", g仪器(i).自动应答)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "blnSend", g仪器(i).可发已核标本)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "ReceiveDir", g仪器(i).通讯目录)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "AutoCheckMan", g仪器(i).自动审核人)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "AutoQCCalc", g仪器(i).自动计算质控)
                    Call SaveSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & g仪器(i).COM口, "SaveAsTonDao", g仪器(i).另存为通道码)
                    
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    MsgBox Err.Description

End Sub

Public Function GetConnectDevs() As Variant
'功能：获取系统连接的检验仪器
    Dim aSettings() As Variant
    Dim aPorts As Variant, i As Integer, PortIndex As Integer
    Dim lngDeviceID As Long, rsTmp As New adodb.Recordset, rsTmp1 As New adodb.Recordset
    Dim strConnType As String  '连接类型
    Dim strIP As String, strPort As String 'ip 和 Port
    Dim varIPSet As Variant 'IP的设置
    Dim lngSaveAsID As Long '另存为的仪器ID
    Dim strSaveAsName As String
    
    aSettings = Array()
    
    Err = 0: On Error Resume Next
    aPorts = GetAllSettings("ZLSOFT", "公共模块\ZlLISSrv")
    On Error GoTo errH
    If IsEmpty(aPorts) Then
        ReDim aPorts(8, 0)
        For i = 0 To 7
            aPorts(i, 0) = "COM" & i + 1
        Next
    End If
   
    If Not IsEmpty(aPorts) Then
        
        ReDim g仪器(UBound(aPorts))
        
        For i = LBound(g仪器) To UBound(g仪器)
            g仪器(i).ID = 0
            g仪器(i).IP = "127.0.0.1"
            g仪器(i).IP端口 = 6666
            g仪器(i).SaveAsID = 0
            g仪器(i).波特率 = 9600
            g仪器(i).类型 = 1
            g仪器(i).COM口 = 0
            g仪器(i).数据位 = 8
            g仪器(i).停止位 = 1
            g仪器(i).握手 = 0
            g仪器(i).校验位 = "N"
            g仪器(i).字符模式 = 0
            g仪器(i).主机 = 0
            g仪器(i).自动应答 = "0"
            g仪器(i).可发已核标本 = 1
        Next
        
        For i = LBound(aPorts) To UBound(aPorts)
            
            strIP = "": strPort = ""
            lngSaveAsID = 0
            strSaveAsName = ""
            
            lngSaveAsID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "SaveAs", 0))
            If lngSaveAsID > 0 Then
                Set rsTmp1 = gobjDatabase.OpenSqlRecord("Select 名称 From 检验仪器 where ID=[1]", "取另存检验仪器名", lngSaveAsID)
                Do Until rsTmp1.EOF
                    strSaveAsName = "" & rsTmp1!名称
                    rsTmp1.MoveNext
                Loop
            End If
            
            strConnType = aPorts(i, 0)

            If strConnType Like "IP*" Then
                'TCPIP连接
                g仪器(i).类型 = 1
                lngDeviceID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                
                If lngDeviceID > 0 Then

                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    gstrSQL = "Select 名称 From 检验仪器 Where ID=[1]"
                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, App.ProductName, lngDeviceID)
                    If Not rsTmp.EOF Then

                        If Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Enabled", 0)) = 1 Then
                            '启用了IP方式,检查IP和端口是否合法
                            strIP = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1")
                            strPort = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Port", 6666)
                            g仪器(i).IP = strIP
                            g仪器(i).IP端口 = Val(strPort)
                            g仪器(i).主机 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Host", 0))
                            
                            g仪器(i).自动应答 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Auto", "0")
                            g仪器(i).可发已核标本 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                            If Not ValidateIP(strIP) And Not ValidatePort(strPort) Then

                                If UBound(aSettings) = -1 Then
                                    ReDim aSettings(2, 0) As Variant
                                Else
                                    ReDim Preserve aSettings(2, UBound(aSettings, 2) + 1) As Variant
                                End If

                                aSettings(0, UBound(aSettings, 2)) = strIP & ":" & strPort
                                aSettings(1, UBound(aSettings, 2)) = "IP " & strIP & " " & rsTmp("名称") & IIf(strSaveAsName = "", "", " -> " & strSaveAsName)
                                aSettings(2, UBound(aSettings, 2)) = lngDeviceID
                            End If

                        End If
                    End If
                End If
            ElseIf strConnType Like "COM*" Then
                'COM连接
                PortIndex = Val(Mid(aPorts(i, 0), 4)) - 1
                lngDeviceID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                g仪器(i).类型 = 0
                g仪器(i).COM口 = Val(PortIndex + 1)
                If lngDeviceID > 0 Then
                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    gstrSQL = "Select 名称 From 检验仪器 Where ID=[1] "
                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, App.ProductName, lngDeviceID)
                    If Not rsTmp.EOF Then
                        If UBound(aSettings) = -1 Then
                            ReDim aSettings(2, 0) As Variant
                        Else
                            ReDim Preserve aSettings(2, UBound(aSettings, 2) + 1) As Variant
                        End If
                        aSettings(0, UBound(aSettings, 2)) = PortIndex
                        aSettings(1, UBound(aSettings, 2)) = "COM" & PortIndex + 1 & " " & rsTmp("名称") & IIf(strSaveAsName = "", "", " -> " & strSaveAsName)
                        aSettings(2, UBound(aSettings, 2)) = lngDeviceID
                    End If
                
                    With g仪器(i)
                        .ID = lngDeviceID
                        .波特率 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Speed", "9600"))
                        .数据位 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "DataBit", "8"))
                        .停止位 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "StopBit", "1"))
                        .校验位 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Parity", "n")
                        .握手 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\COM" & aPorts(i, 0), "HandShaking", "0"))
                        .字符模式 = GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "InputMode", "0")
                        .SaveAsID = lngSaveAsID
                        .自动应答 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Auto", "0"))
                        .可发已核标本 = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                    End With
                End If
            End If
        Next
    End If
    
    If UBound(aSettings) > -1 Then GetConnectDevs = aSettings
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function GetDevices() As adodb.Recordset
'功能：获取所有检验仪器
    On Error GoTo DBError
    Set GetDevices = Nothing
    If gstr仪器数量 = "" Then
        gstrSQL = "Select ID,编码,名称,通讯程序名 From 检验仪器 Order by ID"
    Else
        gstrSQL = "Select * From (Select ID,编码,名称,通讯程序名 From 检验仪器 Order by ID) where Rownum<=[1]"
    End If
    Set GetDevices = gobjDatabase.OpenSqlRecord(gstrSQL, "仪器数据接收", Val(gstr仪器数量))
    Exit Function
DBError:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetComboxIndex(objCbo As ComboBox, ByVal SeekValue As Long) As Long
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If objCbo.ItemData(i) = SeekValue Then Exit For
    Next
    If i > objCbo.ListCount - 1 Then i = 0
    GetComboxIndex = i
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub WriteLog(ByVal ModuleName As String, ByVal ErrorType As Integer, ByVal ErrorNum As Long, ByVal ErrorDesc As String)
    'Module:模块或函数名称
    'ErrorType:日志类型
    'errorNum:错误号或日志编号
    'errorDesc:错误信息或日志信息
    Dim strSQL As String
    
    Call WriteTxtLog(ErrorType, ModuleName, IIf(ErrorNum = 0, "", " ") & ErrorDesc)
    
End Sub

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '功能：在任务栏上增加一个图标
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "仪器数据接收", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub ModifyIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "", Optional ByVal blnMessage As Boolean = True)
    
    '功能：在任务栏上增加一个图标
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = IIf(blnMessage, WM_MOUSEMOVE, 0)
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "仪器数据接收", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_MODIFY, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '功能：从任务栏上删除图标
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Public Sub ResultFromFile(ByVal strFile As String, ByVal lngDeviceID As Long, ByVal strSampleNO As String, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"))

        Dim rsTmp As New adodb.Recordset
        Dim strDevice As String
        Dim objDevice As Object
        Dim aRecord() As String
        Dim i As Integer
        Dim intMicrobe As Integer   '微生物 =1 表示微生物
        Dim lngExeDeptID As Long
    
100     If Len(Trim(strFile)) = 0 Then Exit Sub
    
102     gstrSQL = "Select 通讯程序名,nvl(微生物,0) as 微生物,使用小组ID From 检验仪器 Where ID=[1]"
104     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, App.ProductName, lngDeviceID)
106     If Not rsTmp.EOF Then
108     strDevice = rsTmp(0)
110         intMicrobe = Nvl(rsTmp(1), 0)
112         lngExeDeptID = Nvl(rsTmp(2), 0)
        End If
114     If intMicrobe = 1 Then
116         gstrSQL = "Select 通道编码,抗生素ID As 项目ID, 2 as 小数位数,b.编码||nvl(b.简码,b.中文名) as 名称 From 仪器细菌对照 A, 检验用抗生素 B Where a.抗生素id = b.Id And a.仪器id = [1] "
        Else
118         gstrSQL = "Select a.通道编码, a.项目id, Nvl(a.小数位数, 2) As 小数位数, b.编码 || '-' || Nvl(b.英文名, b.中文名) As 名称," & vbNewLine & _
                        "       LPad(Decode(c.排列序号, Null, b.编码, c.排列序号), 10, '0') As 排列" & vbNewLine & _
                        "From 检验项目 C, 诊治所见项目 B, 检验仪器项目 A" & vbNewLine & _
                        "Where a.项目id = b.Id And a.项目id = c.诊治项目id And a.仪器id = [1] " & vbNewLine & _
                        "Order By LPad(Decode(c.排列序号, Null, b.编码, c.排列序号), 10, '0')"

            '2011-12-07 死锁问题修改，4/5 - 指标排序
        End If
120     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, App.ProductName, lngDeviceID)
    
122     If rsTmp.EOF Then
124         ReDim mItem(1, 0) As Variant
126         mItem(1, 0) = -1
        Else
128         mItem = rsTmp.GetRows
        End If
    
        On Error Resume Next
130     Set objDevice = CreateObject(strDevice)
132     If objDevice Is Nothing Then Call WriteLog("ResultFromFile", LOG_错误日志, Err.Number, "解析程序:" & strDevice & "创建失败!" & vbNewLine & Err.Description)
    
134     Call WriteLog(strDevice & ".ResultFromFile", LOG_通讯日志, 0, "strFile:" & strFile & vbNewLine & "strSampleNO:" & strSampleNO & vbNewLine & "dtStart:" & CStr(dtStart) & vbNewLine & "dtEnd:" & CStr(dtEnd))
136     aRecord = objDevice.ResultFromFile(strFile, strSampleNO, dtStart, dtEnd)
    
        On Error GoTo errH
        'aRecord：返回的检验结果数组(各解析程序必须按以下标准组织结果)
        '   元素之间以|分隔
        '   第0个元素：检验时间
        '   第1个元素：样本序号
        '   第2个元素：检验人
        '   第3个元素：标本
        '   第4个元素：是否质控品
        '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
        '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    
        '有返回结果

138     If UBound(aRecord) > -1 Then
        
            Dim StrUnknow As String, strCaclInfo As String, lngErr As Long, strErr As String
140         For i = 0 To UBound(aRecord)
142             Call WriteLog("mdlLISComm.ResultFromFile", LOG_通讯日志, 0, "记录" & i & ":" & aRecord(i))
144             If InStr(aRecord(i), "|") > 0 Then
                    '文件返回方式，不自动进行质控计算，不自动审核
146                 Call SaveToDataBase(lngDeviceID, lngDeviceID, lngExeDeptID, intMicrobe, 0, "", aRecord(i), mItem, StrUnknow, strCaclInfo, lngErr, strErr)
                
148                 If lngErr <> 0 Then
150                     Call WriteLog("ResultFromFile", LOG_错误日志, lngErr, strErr & vbCrLf & gstrSQL)
                    End If
                End If
            Next
        End If


        Exit Sub
errH:
    If CStr(Erl()) = 138 And Err.Number = 9 Then
        Call WriteLog("ResultFromFile", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  没有返回解码结果")
    Else
152     Call WriteLog("ResultFromFile", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description)
    End If
End Sub

Private Function GetItemID(ByVal strChannel As String, ByVal vItems As Variant, Optional ByRef iDec As Integer, Optional ByRef strItemName As String) As Long
    'iDec:小数位数,strItemNmae : 项目缩写，如果为空则为中文名
    Dim i As Integer
    For i = 0 To UBound(vItems, 2)
        If Trim(Replace(Replace(UCase(strChannel), Chr(10), ""), Chr(13), "")) = _
           Replace(Replace((UCase(vItems(0, i))), Chr(10), ""), Chr(13), "") Then Exit For
    Next
    If i > UBound(vItems, 2) Then
        GetItemID = -1
        iDec = 2
        strItemName = ""
    Else
        GetItemID = CLng(vItems(1, i))
        iDec = Val(vItems(2, i))
        strItemName = vItems(3, i)
    End If
End Function

Public Function ValidateIP(ByVal strIP As String, Optional strErrInfo As String) As Boolean
    '检查IP地址的正确性。
    
    Dim varIP As Variant
    Dim IPError As Integer
    Dim IPd As Integer
    Dim i As Integer
    
    varIP = Split(strIP, ".")
    If UBound(varIP) <> 3 Then
        IPError = 0
    Else
        For i = 0 To 3
            If Not IsNumeric(varIP(i)) Then
                IPError = 1
                Exit For
            Else
                IPd = CInt(varIP(i))
                If IPd < 0 Or IPd > 255 Then
                    IPError = 2
                    Exit For
                Else
                    IPError = -1
                End If
            End If
        Next i
    End If
    
    ValidateIP = True
    Select Case IPError
        Case -1
            If strIP <> "0.0.0.0" Then
                ValidateIP = False
                strErrInfo = ""
            Else
                strErrInfo = "IP不能设为0.0.0.0。"
            End If
        Case 0
            strErrInfo = "IP格式不对，应为XXX.XXX.XXX.XXX。其中XXX为0-255的数字。"
        Case 1
            strErrInfo = "IP地址只能为0-255的数字。"
        Case 2
            strErrInfo = "IP地址的范围只能为0-255之间。"
    End Select
End Function

Public Function ValidatePort(ByVal strPort As String, Optional strErrInfo As String) As Boolean
    '检查端口号的正确性。
    ValidatePort = True
    If Not IsNumeric(Trim(strPort)) Then
        strErrInfo = "端口号只能为1-65535的数字。"
    Else
        If Val(Trim(strPort)) > 0 And Val(Trim(strPort)) <= 65535 Then
            ValidatePort = False
            strErrInfo = ""
        Else
            strErrInfo = "端口号的范围只能在1-65535之间。"
        End If
    End If
End Function

Private Sub WriteTxtLog(ByVal lng类型 As String, ByVal str项目 As String, ByVal str内容 As String)
    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim blnClearData As Boolean
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    'If Val(GetSetting("ZLSOFT", "zlLisLog", "Test", 0)) = 0 Then Exit Sub
    
    blnClearData = gblnClearData
    
    '错误日志(产生时间,错误类型,错误号,错误信息
    If str项目 <> "" Or str内容 <> "" Then
        
        If lng类型 = LOG_错误日志 Then
            '错误日志
            strFileName = App.Path & "\zlLis错误日志_" & Format(date, "yyyyMMdd") & ".LOG"
            If pLast错误日志 = str项目 & "|" & str内容 Then
                Exit Sub
            Else
                pLast错误日志 = str项目 & "|" & str内容
            End If
        ElseIf lng类型 = LOG_通讯日志 Then
            '通讯日志
            
            If blnClearData Then Exit Sub '勾了清空日志选项，则不写日志
            strFileName = App.Path & "\zlLis通讯日志_" & Format(date, "yyyyMMdd") & ".LOG"
            If pLast通讯日志 = str项目 & "|" & str内容 Then
                Exit Sub
            Else
                pLast通讯日志 = str项目 & "|" & str内容
            End If
        ElseIf lng类型 = LOG_未知项 Then
            '未知项
            If blnClearData Then Exit Sub '勾了清空日志选项，则不写日志
            strFileName = App.Path & "\zlLis未知项目_" & Format(date, "yyyyMMdd") & ".LOG"
        End If
        
        If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
        Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
        
        
        strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine ("时间:" & strDate & " 版本:" & App.major & "." & App.minor & "." & App.Revision)
        
        objStream.WriteLine (str项目)
        objStream.WriteLine (str内容)
        
        'objStream.WriteLine (String(50, "-"))
        objStream.Close
        Set objStream = Nothing
    End If
End Sub

Public Sub SaveImg(ByVal lngDevID As Long, ByVal lngID As Long, ByVal strImg As String)
        '保存图形数据到数据库中
        'lngDevID   仪器ID
        'lngID      标本ID
        'strImg     图形数据
    
        Dim aGraphItem() As String
        Dim strImageVal As String
        Dim strImageType As String
        Dim strImageData As String
        Dim intLoop As Integer
        Dim IntCount As Integer
        Dim blnDeleImg As Boolean '保存后是否删除原来的图片
        Dim strPicPath As String, strSQL() As String
        Dim intLayOut As Integer '图片的显示方式
        Dim strBMPFile As String
        Dim blnFtp As Boolean       'FTP是否可用
        Static strFtpPara As String       '保存FTP参数
        Dim strFTPuser As String, strFTPpass As String, strFTPIP As String, strFPTPath As String
        Dim strUploadOk As String, strFTPDir As String, strNewName As String
        Dim objStream As TextStream
    
        On Error GoTo ErrHandle
    
        'FTP连接检查，有效则可以按FTP方式保存图片
100     blnFtp = False
102     If strFtpPara = "" Then
104         strFtpPara = gobjDatabase.GetPara("FTP设置", glngSys, 1208, "")
        End If
106     If UBound(Split(strFtpPara, ";")) >= 3 Then
108        strFTPuser = Split(strFtpPara, ";")(0)
110        strFTPpass = Split(strFtpPara, ";")(1)
112        strFTPIP = Split(strFtpPara, ";")(2)
114        strFPTPath = Split(strFtpPara, ";")(3)
116        If TestFTP(strFTPuser, strFTPpass, strFTPIP, strFPTPath) = "" Then
118             blnFtp = True
           End If
        End If
    
120     aGraphItem = Split(strImg, "^")
    
    
122     For intLoop = 0 To UBound(aGraphItem)
124         strImageVal = Replace(aGraphItem(intLoop), vbCrLf, "")
126         strImageType = Mid(strImageVal, 1, InStr(strImageVal, ";") - 1)
128         strImageData = Mid(strImageVal, InStr(strImageVal, ";") + 1)
        
130         If Mid(strImageData, 1, InStr(strImageData, ";") - 1) >= 100 And Mid(strImageData, 1, InStr(strImageData, ";") - 1) <= 227 Then
                '组织图片数据
            
132             intLayOut = Mid(strImageData, 1, InStr(strImageData, ";") - 1)
134             strPicPath = Mid(strImageData, InStr(strImageData, ";") + 1)
            
136             If InStr(strPicPath, ";") > 0 Then
138                 If Left(strPicPath, 2) = "1;" Then
140                     blnDeleImg = True
                    End If
142                 strPicPath = Mid(strPicPath, InStr(strPicPath, ";") + 1)
                End If
            
144             If Dir(strPicPath) <> "" Then
146                 If UCase(Right(strPicPath, 4)) = ".BMP" And intLayOut >= 100 And intLayOut <= 107 Then
148                     strBMPFile = strPicPath
150                 ElseIf (UCase(Right(strPicPath, 4)) = ".JPG" Or UCase(Right(strPicPath, 4)) = ".GIF") And intLayOut >= 110 And intLayOut <= 127 Then
152                     strBMPFile = strPicPath
154                 ElseIf intLayOut >= 200 And intLayOut <= 227 Then
156                     strPicPath = UCase$(strPicPath)
158                     strBMPFile = zlFileZip(strPicPath)
                    Else
160                     frmLISSrv.picTmp.Picture = LoadPicture(strPicPath)
162                     If Dir(App.Path & "\zlLisIn.bmp") <> "" Then Kill App.Path & "\zlLisIn.bmp"
164                     SavePicture frmLISSrv.picTmp.Picture, App.Path & "\zlLisIn.bmp"
166                     strBMPFile = App.Path & "\zlLisIn.bmp"
                    End If
                
168                 If Not blnFtp Then
                        '保存到数据库
170                     If zlLisBlobSql(lngID, strImageType, strBMPFile, intLayOut, strSQL) Then
172                         WriteLog "执行 SaveImg", LOG_通讯日志, 0, "开始时间"
174                         For IntCount = LBound(strSQL) To UBound(strSQL)
176                             If strSQL(IntCount) <> "" Then
178                                 gstrSQL = strSQL(IntCount)
180                                 gobjDatabase.ExecuteProcedure Replace(strSQL(IntCount), "Call", ""), "保存图像数据"
                                End If
                            Next
182                         WriteLog "执行 SaveImg", LOG_通讯日志, 0, "结束时间"
                        End If
                    Else
                        '保存到FTP
                        '图像位置保存的数据格式为：图像格式;FTP文件路径
                        '图像格式为100-227 。
184                     strFTPDir = strFPTPath & IIf(Right(strFPTPath, 1) = "/", "", "/") & "Dev_" & lngDevID & "/" & Format(gobjDatabase.Currentdate, "yyyyMM")
186                     strNewName = lngID & "_" & strImageType & Right(strPicPath, 4)
188                     strUploadOk = UploadFile(strFTPuser, strFTPpass, strFTPIP, strFTPDir, strBMPFile, strNewName)
190                     If strUploadOk = "" Then
192                         gstrSQL = "Zl_检验图像结果_Update(" & lngID & ",'" & strImageType & "',Null,0,1,'" & _
                             intLayOut & ";" & strFTPDir & "/" & strNewName & "')"
194                         gobjDatabase.ExecuteProcedure gstrSQL, "保存检验图形数据"
                        Else
196                         WriteLog "上传图片文件到FTP", LOG_通讯日志, 0, strUploadOk
                        End If
                    End If
                
198                 If blnDeleImg Then
200                     IntCount = 0
202                     Do While Dir(strPicPath) <> "" And IntCount < 100
204                         IntCount = IntCount + 1
206                         gobjFSO.DeleteFile strPicPath, True
                        Loop
208                     If Dir(strPicPath) <> "" Then
210                         Call WriteLog("SaveImg", LOG_错误日志, 0, "文件" & strPicPath & "估计还有其他程序在用，未能删除，需事后手工删除！")
                        End If
                    End If
                End If
            Else
                '图形数据
212             If Not blnFtp Then
214                 If Len(strImageData) > 2000 Then
                        '保存大于2000以上数据
216                     For IntCount = 1 To CInt(Len(strImageData) / 1000) + 1
218                         If Len(strImageData) > 0 Then
                            
220                             gstrSQL = "Zl_检验图像结果_Update(" & lngID & ",'" & strImageType & "','" & _
                                                        Mid(strImageData, IntCount * 1000 - 999, 1000) & "'," & _
                                                        "1," & IntCount & ")"
222                             gobjDatabase.ExecuteProcedure gstrSQL, "检验图像保存"
                            End If
                        Next
                    Else
224                     gstrSQL = "Zl_检验图像结果_Update(" & lngID & ",'" & strImageType & "','" & strImageData & "',0,1)"
226                     gobjDatabase.ExecuteProcedure gstrSQL, "检验图像保存"
                    End If
                Else
                    '存为TXT文件然后上传
228                 intLayOut = Mid(strImageData, 1, InStr(strImageData, ";") - 1)
230                 strPicPath = Mid(strImageData, InStr(strImageData, ";") + 1)
                
232                 strBMPFile = App.Path & "\" & lngID & "_" & strImageType & ".txt"
234                 If gobjFSO.FileExists(strBMPFile) Then gobjFSO.DeleteFile strBMPFile
236                 Set objStream = gobjFSO.CreateTextFile(strBMPFile)
238                 objStream.Write strPicPath
240                 objStream.Close
242                 Set objStream = Nothing
                
                    '保存到FTP
                    '图像位置保存的数据格式为：图像格式;FTP文件路径
                    '图像格式为0-6
                
244                 strFTPDir = strFPTPath & IIf(Right(strFPTPath, 1) = "/", "", "/") & "Dev_" & lngDevID & "/" & Format(gobjDatabase.Currentdate, "yyyyMM")
246                 strNewName = lngID & "_" & strImageType & ".txt"
248                 strUploadOk = UploadFile(strFTPuser, strFTPpass, strFTPIP, strFTPDir, strBMPFile, strNewName)
250                 If strUploadOk = "" Then
252                     gstrSQL = "Zl_检验图像结果_Update(" & lngID & ",'" & strImageType & "',Null,0,1,'" & _
                         intLayOut & ";" & strFTPDir & "/" & strNewName & "')"
254                     gobjDatabase.ExecuteProcedure gstrSQL, "保存检验图形数据"
                    Else
256                     WriteLog "上传图像数据文件到FTP", LOG_通讯日志, 0, strUploadOk
                    End If
258                 If gobjFSO.FileExists(strBMPFile) Then gobjFSO.DeleteFile strBMPFile
                End If
            End If
        Next

        Exit Sub
ErrHandle:
260     Call WriteLog("SaveImg-" & CStr(Erl()), LOG_错误日志, Err.Number, Err.Description)

End Sub


Private Function zlLisBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, ByVal layOut As Integer, ByRef arySql() As String) As Boolean
    '生成保存图片的SQL
    'Action 检验ID
    'KeyWord 标题
    'strFile 图片文件
    'arySql 生成的SQL存放在此数组中
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Dim lngLBound As Long, lngUBound As Long    '传入数组的最小最大下标
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo 0
    
    lngFileNum = FreeFile
    WriteLog "生成BlobSQL", LOG_通讯日志, 0, "开始时间"
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    conChunkSize = 512
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        If strText <> "" Then
            If lngCount = 0 Then strText = layOut & ";" & strText
            arySql(lngUBound + lngCount + 1) = "Zl_检验图像结果_Update(" & Action & ",'" & KeyWord & "','" & strText & "',1," & IIf(lngCount = 0, 1, 0) & ")"
        End If
    Next
    Close lngFileNum
    WriteLog "生成BlobSQL", LOG_通讯日志, 0, "结束时间"
    zlLisBlobSql = True
    Exit Function

errHand:
    Close lngFileNum
    zlLisBlobSql = False
End Function
Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal BeginDate As String) As String
    '-----------------------------------------------------------------------------------------
    '功能:获取特殊时间
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    Dim dateNow As Date
    
    If BeginDate = "" Then
        dateNow = gobjDatabase.Currentdate
    Else
        dateNow = BeginDate
    End If
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(dateNow, "YYYY-MM-DD")))
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 2, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 8 - intDay, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(dateNow, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(dateNow, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(dateNow, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "不重复"
        If bytFlag = 1 Then
            GetDateTime = "2000-01-01 00:00:00"
        Else
            GetDateTime = "3000-12-31 23:59:59"
        End If
    End Select
    
End Function

Public Function CreateSample(ByVal lngDeviceID As Long, ByVal strBarcode As String, _
    ByRef strSampleNO As String, ByVal dtSampleDate As Date, ByVal intType As Integer) As Boolean
        'inttype=0
        Dim strSQL As String, rsTmp As adodb.Recordset, rs As New adodb.Recordset
        Dim lngKey As Long, strItemRecords As String
        Dim lngDeptID As Long '当前仪器科室
        Dim rsItem As New adodb.Recordset
        Dim strItem As String                           '检验项目
        Dim str姓名 As String, str性别 As String, str年龄 As String
        On Error GoTo DBErr
    
100     CreateSample = False
    
        '查找仪器科室
102     strSQL = "Select 使用小组id From 检验仪器 Where ID = [1]"
104     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "生成条码标本", lngDeviceID)
106     lngDeptID = glngExeDeptID
108     If Not rsTmp.EOF Then
110         lngDeptID = Nvl(rsTmp("使用小组id"), glngExeDeptID)
        End If
    
112     If Val(strSampleNO) <= 0 Then
114         strSampleNO = Val(CalcNextCode(lngDeviceID, 0, intType))
        End If

        '查找符合条码的项目指标
    '    strSql = "Select A.相关ID AS ID," & _
            "C.姓名||Decode(A.婴儿,0,'',Null,'','(婴儿)') As 姓名,A.性别,A.年龄,F.No," & _
            "I.诊治项目ID As 项目ID,Decode(I.结果类型,3,Nvl(I.默认值,'-'),2,I.默认值,'') As 结果,'' As 标志," & _
            "Trim(REPLACE(REPLACE(' '||zlGetReference(I.诊治项目ID,A.标本部位,DECODE(A.性别,'男',1,'女',2,0),C.出生日期,Y.仪器ID,A.年龄),' .','0.'),'～.','～0.')) AS 结果参考," & _
            "NVL(A.紧急标志,0) AS 紧急,F.采样时间,F.采样人 " & _
            "FROM 病人医嘱记录 A," & _
            "病人信息 C,病人医嘱发送 F,检验报告项目 G,检验项目 I,检验仪器项目 Y " & _
            "WHERE A.诊疗类别 = 'C' " & _
            "AND A.病人ID=C.病人ID " & _
            "AND A.相关id IS NOT NULL " & _
            "AND A.医嘱状态=8 AND A.ID=F.医嘱id " & _
            "AND A.诊疗项目id=G.诊疗项目id AND G.细菌ID Is Null AND G.报告项目id=Y.项目id(+) " & _
            "AND G.报告项目ID=I.诊治项目ID " & _
            "AND (Y.仪器ID+0=[1] Or (Y.仪器ID Is Null And F.执行部门ID=[3])) " & _
            "And F.样本条码=[2] "
    '        "AND F.执行状态=0 "
    
116     strSQL = "Select ID, 姓名, 性别, 年龄, NO, 项目id, 结果, 标志, 结果参考, 紧急, 采样时间, 采样人, Rownum As 排列序号, 诊疗项目id," & vbNewLine & _
                "       编码,标本部位,开嘱科室ID,开嘱医生,标识号,当前床号,病人科室 " & vbNewLine & _
                "From (Select A.相关id As ID, C.姓名 || Decode(A.婴儿, 0, '', Null, '', '(婴儿)') As 姓名, A.性别, A.年龄, F.NO," & vbNewLine & _
                "              I.诊治项目id As 项目id, Decode(I.结果类型, 3, Nvl(I.默认值, '-'), 2, I.默认值, '') As 结果, '' As 标志," & vbNewLine & _
                "              Trim(Replace(Replace(' ' || Zlgetreference(I.诊治项目id, A.标本部位, Decode(A.性别, '男', 1, '女', 2, 0)," & vbNewLine & _
                "                                                          C.出生日期, Y.仪器id, A.年龄), ' .', '0.'), '～.', '～0.')) As 结果参考," & vbNewLine & _
                "              Nvl(A.紧急标志, 0) As 紧急, F.采样时间, F.采样人, G.排列序号, A.诊疗项目id, M.编码, " & vbNewLine & _
                "              a.标本部位,开嘱科室ID,开嘱医生,decode(a.病人来源,2,c.住院号,c.门诊号) as 标识号,c.当前床号,l.名称 as 病人科室 " & vbNewLine & _
                "       From 病人医嘱记录 A, 病人信息 C, 病人医嘱发送 F, 检验报告项目 G, 检验项目 I, 检验仪器项目 Y, 诊疗项目目录 M ,部门表 L " & vbNewLine & _
                "       Where A.诊疗类别 = 'C' And A.病人id = C.病人id And A.相关id Is Not Null And A.医嘱状态 = 8 And A.ID = F.医嘱id And" & vbNewLine & _
                "             A.诊疗项目id = G.诊疗项目id And G.细菌id Is Null And G.报告项目id = Y.项目id(+) And" & vbNewLine & _
                "             G.报告项目id = I.诊治项目id And A.诊疗项目id = M.ID(+) And a.病人科室ID = l.ID" & vbNewLine & _
                "             and (Y.仪器id + 0 = [1] Or (Y.仪器id Is Null And F.执行部门id = [3])) And nvl(F.执行状态,0) = 0  And F.样本条码 = [2]" & vbNewLine & _
                "       Order By M.编码, G.排列序号)"

118     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "生成条码标本", lngDeviceID, strBarcode, lngDeptID)
120     If rsTmp.EOF Then Exit Function
    
122     gstrSQL = "Select B.病人id, B.主页id, B.序号, B.婴儿姓名, B.婴儿性别" & vbNewLine & _
                        "From 病人医嘱记录 A, 病人新生儿记录 B" & vbNewLine & _
                        "Where A.病人id = B.病人id And A.主页id = B.主页id And A.婴儿 = B.序号 And A.相关id = [1] And Rownum = 1"
124     Set rs = gobjDatabase.OpenSqlRecord(gstrSQL, "CreateSample", CLng(rsTmp("ID")))
126     If rs.EOF = False Then
128         str姓名 = Nvl(rs("婴儿姓名"))
130         str性别 = Nvl(rs("婴儿性别"))
132         str年龄 = "婴儿"
        Else
134         str姓名 = Nvl(rsTmp("姓名"))
136         str性别 = Nvl(rsTmp("性别"))
138         str年龄 = Nvl(rsTmp("年龄"))
        End If
    
        '读出检验项目
140     gstrSQL = "select distinct 医嘱内容 from 病人医嘱记录 a , 病人医嘱发送 b, 检验报告项目 c , 检验仪器项目 d " & vbNewLine & _
                  "  where a.id = b.医嘱ID and a.相关id is not null and a.诊疗项目ID = c.诊疗项目ID and " & vbNewLine & _
                  "  c.报告项目ID = d.项目ID(+) and  (d.仪器id + 0 = [1] Or (d.仪器id Is Null And b.执行部门id = [3])) and b.样本条码 = [2] "
142     Set rsItem = gobjDatabase.OpenSqlRecord(gstrSQL, "生成条码标本_1", lngDeviceID, strBarcode, lngDeptID)
144     Do Until rsItem.EOF
146         strItem = strItem & " " & Nvl(rsItem("医嘱内容"))
148         rsItem.MoveNext
        Loop
150     strItem = Trim(strItem) & "(" & Nvl(rsTmp("标本部位")) & ")"
        
        '产生标本记录
152     lngKey = gobjDatabase.GetNextId("检验标本记录")
154     gstrSQL = "ZL_检验标本记录_标本核收(" & lngKey & "," & _
            rsTmp("ID") & ",'" & rsTmp("ID") & "',0,'" & _
            strSampleNO & "'," & _
            IIf(IsNull(rsTmp("采样时间")), "Null", "TO_DATE('" & Format(rsTmp("采样时间"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
            IIf(IsNull(rsTmp("采样人")), "Null", "'" & rsTmp("采样人") & "'") & "," & _
            lngDeviceID & "," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
            "'" & _
            gstrUserName & "'," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0," & _
            intType & ",NULL,'" & _
            str姓名 & "','" & str性别 & "','" & str年龄 & "','" & Nvl(rsTmp("No")) & "','" & _
            Nvl(rsTmp("标本部位")) & "'," & Nvl(rsTmp("开嘱科室ID")) & ",'" & Nvl(rsTmp("开嘱医生")) & "'," & _
            Nvl(rsTmp("标识号")) & ",'" & Nvl(rsTmp("当前床号")) & "','" & Nvl(rsTmp("病人科室")) & "','" & _
            strItem & "',Null,Null," & lngDeptID & ",'" & gstrUserCode & "','" & gstrUserName & "')"
156     gobjDatabase.ExecuteProcedure gstrSQL, "生成条码标本"
                                                                
        '填写指标
158     strItemRecords = ""
160     Do While Not rsTmp.EOF
162         strItemRecords = strItemRecords & "|" & rsTmp("ID") & "^" & rsTmp("项目ID") & "^" & _
                Nvl(rsTmp("结果")) & "^" & Nvl(rsTmp("标志"), 0) & "^" & Nvl(rsTmp("结果参考")) & "^" & _
                Nvl(rsTmp("诊疗项目ID")) & "^" & Nvl(rsTmp("排列序号"))
            
164         rsTmp.MoveNext
        Loop
    
166     If Len(strItemRecords) > 0 Then
168         strItemRecords = Mid(strItemRecords, 2)
            
170         gstrSQL = "Zl_检验普通结果_Write(" & lngKey & "," & _
                lngDeviceID & ",'" & strItemRecords & "',0,0)"
172         gobjDatabase.ExecuteProcedure gstrSQL, "生成条码标本"
        End If
        Exit Function
DBErr:
174     Call WriteLog("clsLISComm.CreateSample", LOG_错误日志, Err.Number, CStr(Erl()) & "行," & Err.Description)
End Function

Private Function CalcNextCode(ByVal lngKey As Long, ByVal intRow As Integer, ByVal iType As Integer) As String
    '--------------------------------------------------------------------------------------------------------
    '功能:计算指定仪器在当天内的下一个缺省标本号
    '参数:lngKey                检验仪器ID
    '     iType                 标本类别：0=普通、1=急诊
    '返回:缺省标本号码
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New adodb.Recordset
    Dim strToday As String
    Dim strTmp As String
    Dim lng次数 As Long
    Dim strLabNo As String, strLabQCNo As String '检验标本、质控标本
    Dim mstrSQL As String, mlngLoop As Long
    Dim mlngDefaultItemID As Long
    
    '时间,仪器,标本号
    On Error GoTo errHand
    mlngDefaultItemID = 0
    strToday = Format(gobjDatabase.Currentdate, "YYYY-MM-DD")
    
    On Error GoTo point1
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(标本序号)),0) AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本id(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                        IIf(mlngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & " And 医嘱ID Is Not Null" & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = gobjDatabase.OpenSqlRecord(mstrSQL, "计算", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                           CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabNo = gobjCommFun.Nvl(rs("最大序号"))
    
    On Error GoTo errHand
    GoTo point2
    
point1:
    On Error GoTo errHand
    
    mstrSQL = "SELECT NVL(MAX(标本序号),'') AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本id(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & " And 医嘱ID Is Not Null" & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = gobjDatabase.OpenSqlRecord(mstrSQL, "计算", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabNo = gobjCommFun.Nvl(rs("最大序号"))
    
point2:
    On Error GoTo point3
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(标本序号)),0) AS 最大序号 FROM 检验标本记录 a,检验申请项目 b " & _
                "WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本ID(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id= [1] ") & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = gobjDatabase.OpenSqlRecord(mstrSQL, "计算", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabQCNo = gobjCommFun.Nvl(rs("最大序号"))
    
    On Error GoTo errHand
    GoTo point4
    
point3:
    On Error GoTo errHand
    
    mstrSQL = "SELECT NVL(MAX(标本序号),'') AS 最大序号 FROM 检验标本记录 a,检验申请项目 b" & _
                " WHERE 核收时间 BETWEEN [2] and [3] And a.id = b.标本ID(+) And nvl(a.是否质控品,0) = 0 " & _
                    IIf(lngKey = -1, " AND 仪器id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.诊疗项目id = [4] ", ""), "AND 仪器id=[1] ") & _
                    IIf(iType = 1, " And 标本类别=1", " And Nvl(标本类别,0)<>1")
    Set rs = gobjDatabase.OpenSqlRecord(mstrSQL, "计算", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabQCNo = gobjCommFun.Nvl(rs("最大序号"))
    
point4:
    If strLabNo >= strLabQCNo Then
        CalcNextCode = strLabNo
    Else
        CalcNextCode = strLabQCNo
    End If
'    If Val(strLabQCNo) > Val(strLabNo) + 100 Then CalcNextCode = strLabNo

'    For mlngLoop = 1 To vsf2.Rows - 1
'        If mlngLoop <> intRow Then
'            If Val(vsf2.RowData(mlngLoop)) = lngKey Then
'                If Val(CalcNextCode) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
'                    CalcNextCode = Val(vsf2.TextMatrix(mlngLoop, 2))
'                End If
'            End If
'        End If
'    Next
'
    If Val(CalcNextCode) <= 0 Then
        CalcNextCode = "1"
        Exit Function
    End If
'
    CalcNextCode = Val(CalcNextCode) + 1
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String, Optional ByVal strUnZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If gobjFSO.FileExists(strUnZipFile) Then gobjFSO.DeleteFile strUnZipFile
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strUnZipFile) <> "" Then
        zlFileUnzip = strUnZipFile
    Else
        zlFileUnzip = ""
    End If
End Function
'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLLIS" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function

Public Function SaveToDataBase(ByVal lngDeviceID As Long, ByVal lngMainID As Long, ByVal lngExeDeptID As Long, ByVal intMicrobe As Integer, ByVal intAutoQCCalc As Integer, ByVal strAutoCheckMan As String, ByVal strResult As String, ByVal vItems As Variant, ByRef strUnknown As String, ByRef strAutoCaleInfo As String, ByRef lngErr As Long, ByRef strErr As String, Optional ByRef strIDs As String, Optional ByRef strlogs As String) As Boolean
        '保存数据到数据库
        'lngDeviceID :仪器ID
        'lngExeDeptID : 检验小组ID
        ' intMicrobe: 是否微生物
        ' intAutoQCCalc  :是否要自动计算质控标本
        ' strAutoCheckMan: 自动审核人
        ' strResult : 检验结果串
        ' strUnknown： 返回 未知项
        ' strAutoCaleInfo : 自动计算的结果信息
        ' lngErr :错误号
        ' strErr :错误串
        ' strIDs As String 原始数据对应的检验记录ID（可能多个）,用于串口通讯中返回前台进行刷新
      '保存数据到数据库
      Dim aRecord() As String, aItem() As String
      Dim aTmp() As String
      Dim strDate As String, strSampleID As String, strBarcode As String
      Dim strName As String, strSample As String, strSex As String, strBirth As String
      Dim i As Long, j As Long
      Dim rsTmp As New adodb.Recordset, strSQL As String

      Dim lngID As Long

      Dim blnAuditing As Boolean '是否审核
      Dim intItemAuditing As Integer '指标审核，1-已审，0-未审
      Dim strItemAuditing As String '指标审核内容
      Dim strItemCode As String     '指标通道码
      Dim lngItemID As Long '项目ID
      Dim strItemRecords As String
      Dim aNos() As String, iType As Integer '标本号数组
      Dim aQC() As String                    '质控数组
  
      Dim iDec As Integer '小数位数
      Dim blnQryWithSampleNO As Boolean
      Dim str未知项 As String
      Dim strStartDate As String
      Dim strEndDate As String
    
      Dim strQCList() As String '保存需要计算的内容
      Dim strAutoCheck() As String  '保存要自动审核的内容
      Dim int2Verify As Integer
    
      Dim strBatchSQL() As String '保存要执行的SQL
      Dim str检验备注 As String
      Dim bln无主标本 As Boolean  '无主标本不能自动审核
      Dim strLog  As String '保存入库日志，写入指定目录
      Dim strItemName As String '指标名称
      Dim strItemsInfo As String '入库的指标信息
      
      On Error GoTo DBError
100   ReDim strQCList(0) As String
102   ReDim strAutoCheck(0) As String
      Dim intMicrobeDay As Integer   '微生物天数查询


    
104    SaveToDataBase = False
       intMicrobeDay = gobjDatabase.GetPara("微生物查询时间", 100, 1208, 0)
106    int2Verify = gobjDatabase.GetPara("使用二级报告审核", 100, 1208, 0)
    
108    strLog = Format(Now, "yyyy-MM-dd HH:mm:ss") & " 开始入库"
110    If Len(strResult) > 0 Then
       
112        aRecord = Split(strResult, "||")
114        For i = 0 To UBound(aRecord)
116            ReDim strBatchSQL(0) As String
118            blnAuditing = False
            
    '118            Call Return_Decode(aRecord(i))   '返回解析结果到通讯监控
120            aTmp = Split(aRecord(i), vbCrLf)
            
122            aItem = Split(aTmp(0), "|")
124            aQC = Split(aItem(4), "^")              '标记质控
126            If UBound(aItem) >= 4 Then
                  '有效的报告组
128                aNos = Split(aItem(1), "^") '标本号格式：标本号^标本类别^SampleID（0：常规，1：急诊）
130                If UBound(aNos) = 0 Then
                      '没有标本类别，则按常规标本处理
132                    strDate = Trim(aItem(0)): strSampleID = IIf(aQC(0) = "1", aNos(0), Val(aNos(0))): iType = 0: strBarcode = ""
                   Else
134                    If gblnEmerge = True Then
136                         iType = Val(aNos(1))
                       Else
                            '2011-02-21 当按条码保存时,"是否区分急诊"参数为False时,按常规标本处理.
138                         iType = 0
                       End If
140                    strDate = Trim(aItem(0)): strSampleID = IIf(aQC(0) = "1", aNos(0), Val(aNos(0))): strBarcode = ""
142                    If UBound(aNos) > 1 Then
144                        strBarcode = Trim(aNos(2))
                       End If
                   End If
                  '单独处理标本生成规则（按时间）
146                strStartDate = GetDateTime(mMakeNoRule, 1, strDate)
148                strEndDate = GetDateTime(mMakeNoRule, 2, strDate)
                
150                strName = Trim(aItem(2)): strSample = Trim(aItem(3))
152                 If Trim(strName) = "" Then strName = gstrUserName
                  '判断是否无主标本
154                If Len(Trim(strBarcode)) = 0 Then
                      '按标本号查
156                    blnQryWithSampleNO = True
                   Else
                      '按条码查询
158                    gstrSQL = "Select  a.id,a.医嘱ID,a.出生日期,a.审核人,a.标本类型, a.初审人,Decode(A.性别,Null,0,'男',1,'女',2,0) As 性别A,to_char(c.出生日期,'yyyy-mm-dd') As 出生日期A From 检验标本记录 a,病人医嘱记录 b,病人信息 c " & _
                          " Where a.医嘱id=b.id(+) And b.病人id=c.病人id(+)" & _
                          " And a.核收时间 Between [1] And [2]" & _
                          " And a.仪器ID=[3] And a.样本条码=[6]"
160                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "查询标本记录", CDate(strStartDate), _
                          CDate(strEndDate), lngDeviceID, strSampleID, iType, strBarcode)
162                    If Not rsTmp.EOF Then
164                        blnQryWithSampleNO = False
                       Else
                          '检验是否已有标本
166                        gstrSQL = "Select a.id,a.医嘱ID,a.出生日期,a.审核人,a.标本类型, a.初审人,Decode(A.性别,Null,0,'男',1,'女',2,0) As 性别A,to_char(c.出生日期,'yyyy-mm-dd') As 出生日期A From 检验标本记录 a,病人医嘱记录 b,病人信息 c " & _
                          " Where a.医嘱id=b.id(+) And b.病人id=c.病人id(+)" & _
                          " And a.核收时间 Between [1] And [2]" & _
                          " And a.仪器ID=[3] And a.标本序号=[4] " & IIf(gblnEmerge, " And Nvl(a.标本类别,0)=[5]", "")
168                        Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "查询标本记录", CDate(Format(strDate, "yyyy-MM-dd") & " 00:00:00"), _
                              CDate(Format(strDate, "yyyy-MM-dd") & " 23:59:59"), lngDeviceID, strSampleID, iType)
170                        If rsTmp.EOF = True Then
                              '根据条码生成标本
172                            Call CreateSample(lngDeviceID, strBarcode, strSampleID, CDate(strDate), iType)
174                            blnQryWithSampleNO = True
                           Else
176                            If Val(Nvl(rsTmp("医嘱id"))) = 0 Then
                                  '标本为无主时也生成
178                                Call CreateSample(lngDeviceID, strBarcode, strSampleID, CDate(strDate), iType)
180                                blnQryWithSampleNO = True
                               End If
                           End If
                       End If
                   End If
182                If blnQryWithSampleNO Then
184                    gstrSQL = "Select a.id,a.医嘱ID,a.出生日期,a.审核人,a.标本类型, a.初审人,Decode(A.性别,Null,0,'男',1,'女',2,0) As 性别A,to_char(c.出生日期,'yyyy-mm-dd') As 出生日期A From 检验标本记录 a,病人医嘱记录 b,病人信息 c " & _
                          " Where a.医嘱id=b.id(+) And b.病人id=c.病人id(+)" & _
                          " And a.核收时间 Between [1] And [2]" & _
                          " And a.仪器ID=[3] And a.标本序号=[4] " & IIf(gblnEmerge, " And Nvl(a.标本类别,0)=[5]", "")
                        '--- 2012-11-21 微生物标本，不加时间条件。
186                     If intMicrobe = 1 Then
                            If intMicrobeDay = 0 Then
                                gstrSQL = Replace(gstrSQL, "And a.核收时间 Between [1] And [2]", " ")
                            End If
                            strStartDate = Format(gobjDatabase.Currentdate - intMicrobeDay, "yyyy-mm-dd 00:00:00")
                        End If
                        Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "查询标本记录", CDate(strStartDate), _
                        CDate(strEndDate), lngDeviceID, strSampleID, iType)
                   End If
188                bln无主标本 = False
190                If rsTmp.EOF Then
                      '无主标本增加临时标本记录
192                    bln无主标本 = True
194                    strSex = 0
196                    strBirth = ""
                    
198                    lngID = gobjDatabase.GetNextId("检验标本记录")
                       
200                    gstrSQL = "ZL_检验标本记录_INSERT(" & lngID & ",NULL,'" & _
                          strSampleID & "',NULL,NULL," & lngDeviceID & ",'" & strName & "'," & _
                          "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                          "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strSample & "'," & _
                          "Null,To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strName & "','0'," & lngExeDeptID & "," & iType & "," & intMicrobe & ")"

202                    gobjDatabase.ExecuteProcedure gstrSQL, "插入检验临时记录"
204                    strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 生成无主标本" & vbNewLine & "标本ID=" & lngID & ",日期=" & strDate & ",标本号=" & strSampleID & ",条码=" & strBarcode & ".仪器id=" & lngDeviceID & ",操作员=" & strName & ",微生物标本=" & intMicrobe
                   Else
206                    If Val("" & rsTmp!医嘱ID) = 0 Then bln无主标本 = True
208                    strSex = Nvl(rsTmp("性别A"), 0)
210                    strBirth = Nvl(rsTmp("出生日期A"))
212                    If intMicrobe = 0 Then
214                        strSample = Nvl(rsTmp("标本类型"))
                       End If
216                    lngID = rsTmp("ID")
218                    blnAuditing = Not IsNull(rsTmp("初审人"))
220                    If blnAuditing = False Then
222                        blnAuditing = Not IsNull(rsTmp("审核人"))
                       End If
224                    strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 找到标本" & vbNewLine & "标本ID=" & lngID & ",日期=" & strDate & ",标本号=" & strSampleID & ",条码=" & strBarcode & ",仪器id=" & lngDeviceID & ",操作员=" & strName & ",微生物标本=" & intMicrobe
                   End If
                

226                If Not blnAuditing Then
228                    If InStr(strIDs, "," & lngID) = 0 Then strIDs = strIDs & "," & lngID
                      '处理检验项目
230                    strItemRecords = ""
232                    str未知项 = ""
234                    str检验备注 = ""
236                    strItemsInfo = ""
238                    For j = 5 To UBound(aItem) Step 2
                          '根据通道号修改相应项目结果，未找到的则直接增加（根据通道号找不到项目的暂不处理）
                          '根据通道号找项目
                            strItemAuditing = ""
                            If InStr(aItem(j), "^") > 0 Then
                                strItemCode = Split(aItem(j), "^")(0)
                                intItemAuditing = Val(Split(aItem(j), "^")(1))
                                If UBound(Split(aItem(j), "^")) = 2 Then
                                    strItemAuditing = Split(aItem(j), "^")(2)
                                End If
                            Else
                                strItemCode = aItem(j)
                                intItemAuditing = 0
                            End If
240                         lngItemID = GetItemID(strItemCode, vItems, iDec, strItemName)
242                         If lngItemID > 0 Then
                            
244                            gstrSQL = "select 项目id from 检验仪器项目 where 仪器id = [1] and 糖耐量项目 = -1 and 项目id = [2] "
246                            Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "糖耐量", lngMainID, lngItemID)
248                            If rsTmp.EOF = False Then
                                  '仪器有糖耐量项目时的处理
250                                If strBarcode <> "" Then
                                      '有条码时的处理 ,根据通道码
252                                    gstrSQL = "Select d.项目id" & vbNewLine & _
                                              "From 病人医嘱记录 A, 病人医嘱发送 B, 检验报告项目 C, 检验仪器项目 D" & vbNewLine & _
                                              "Where A.ID = B.医嘱id And B.样本条码 = [2] And A.诊疗项目id = C.诊疗项目id And C.报告项目id = D.项目id" & vbNewLine & _
                                              "      And D.仪器id = [1] And D.通道编码 =[3] And D.糖耐量项目 = -1"
254                                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "仪器糖耐量", lngMainID, strBarcode, CStr(strItemCode))
256                                    If rsTmp.EOF = False Then
258                                         strItemRecords = strItemRecords & "|" & Nvl(rsTmp("项目ID")) & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
260                                         strItemsInfo = strItemsInfo & "," & strItemName & "(糖耐量项目)" & "=" & aItem(j + 1)
                                       Else
262                                        strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
264                                        strItemsInfo = strItemsInfo & "," & strItemName & "=" & aItem(j + 1)
                                       End If
                                   Else
                                      '没有条码时的处理
266                                    gstrSQL = "Select B.项目id" & vbNewLine & _
                                              " From 检验普通结果 A, 检验仪器项目 B" & vbNewLine & _
                                              " Where A.检验项目id = B.项目id And B.仪器id = [1] And B.糖耐量项目 = -1 And B.通道编码=[3]  And A.检验标本id = [2] "
268                                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "仪器糖耐量", lngMainID, lngID, CStr(strItemCode))
270                                    If rsTmp.EOF = False And rsTmp.RecordCount = 1 Then
272                                        strItemRecords = strItemRecords & "|" & Nvl(rsTmp("项目ID")) & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
274                                        strItemsInfo = strItemsInfo & "," & strItemName & "(糖耐量项目)" & "=" & aItem(j + 1)
                                       Else
276                                        strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
278                                        strItemsInfo = strItemsInfo & "," & strItemName & "=" & aItem(j + 1)
                                       End If
                                       
                                   End If
                               Else
                                  '仪器没有糖耐量项目时的处理
280                                strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
282                                strItemsInfo = strItemsInfo & "," & strItemName & "=" & aItem(j + 1)
                               End If
                           Else
   
                            
284                            If strItemCode = "检验备注" Then
                            
286                                str检验备注 = str检验备注 & IIf(str检验备注 <> "", vbNewLine, "") & aItem(j + 1)
288                                If InStr(UCase(str检验备注), "VBNEWLINE") > 0 Then
290                                    str检验备注 = Replace(str检验备注, "vbnewline", vbNewLine, , , vbTextCompare)
                                   End If
                               Else
292                                If str未知项 = "" Then str未知项 = "标本号     项目标识     项目值" & vbNewLine
294                                str未知项 = str未知项 & strSampleID & Space(30 - Len(strSampleID)) & _
                                  strItemCode & Space(30 - Len(strItemCode)) & _
                                  aItem(j + 1) & vbNewLine
296                               strItemsInfo = strItemsInfo & "," & strItemCode & "(未对码)=" & aItem(j + 1)
                               End If
    '                            mcnAccess.Execute strSql
                           End If
                           If strItemAuditing <> "" Then
                                strItemRecords = strItemRecords & "^" & strItemAuditing
                           End If
                       Next
298                    If str未知项 <> "" Then Call WriteLog("SaveToDataBase", LOG_未知项, 0, str未知项)
300                    strUnknown = str未知项
                       
302                    If Len(strItemRecords) > 0 Then
304                        strItemRecords = Mid(strItemRecords, 2)
                           strItemRecords = ItemLocalSort(strItemRecords, vItems) '2011-12-07 死锁问题修改，1/5 - 指标排序
                            
306                        gstrSQL = "ZL_检验普通结果_BATCHUPDATE(" & lngID & "," & _
                              lngDeviceID & ",'" & strSample & "'," & strSex & "," & _
                              IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                              strItemRecords & "'," & intMicrobe & ")"
308                        gobjDatabase.ExecuteProcedure gstrSQL, "检验结果报告"
                            
                            gstrSQL = "Zl_重新计算结果_Cale(" & lngID & ")" '2011-12-07 死锁问题修改，2/5 - 加调重新计算过程
                            gobjDatabase.ExecuteProcedure gstrSQL, "检验结果报告"
                            
310                        strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 保存检验结果" & vbNewLine & Mid$(strItemsInfo, 2)
312                        If str检验备注 <> "" Then
314                            str检验备注 = Replace(str检验备注, "'", "")
316                            gstrSQL = "Zl_检验标本记录_更新备注(" & lngID & ",'" & str检验备注 & "',1)"
318                            gobjDatabase.ExecuteProcedure gstrSQL, "检验结果备注"
320                            strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 保存检验备注" & vbNewLine & str检验备注
                           End If

                        
                          '保存为质控
322                        If aQC(0) = 1 Then
                               Dim date当前日期 As Date, lngQCID As Long, str标本号 As String
                               Dim var标本号 As Variant, iCoutn As Integer
324                            lngQCID = 0
326                            date当前日期 = gobjDatabase.Currentdate
328                            gstrSQL = "Select ID,标本号 From 检验质控品 Where [2] between 开始日期 and 结束日期 And 仪器id = [1] "
330                            Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, gstrSysName, lngDeviceID, date当前日期)
                            
332                            Do Until rsTmp.EOF Or lngQCID <> 0
334                                str标本号 = "" & rsTmp.Fields("标本号")
336                                If InStr(str标本号, ",") > 0 Then
338                                    var标本号 = Split(str标本号, ",")
340                                    For iCoutn = 0 To UBound(var标本号)
342                                        If var标本号(iCoutn) Like "*-*" Then
344                                            If strSampleID >= Val(Split(var标本号(iCoutn), "-")(0)) And strSampleID <= Val(Split(var标本号(iCoutn), "-")(1)) Then
346                                                lngQCID = rsTmp.Fields("ID")
                                               End If
                                           Else
348                                            If var标本号(iCoutn) = strSampleID Then
350                                                lngQCID = rsTmp.Fields("ID")
                                               End If
                                           End If
                                       Next
352                                ElseIf str标本号 Like "*-*" Then
354                                    If strSampleID >= Val(Split(str标本号, "-")(0)) And strSampleID <= Val(Split(str标本号, "-")(1)) Then
356                                        lngQCID = rsTmp.Fields("ID")
                                       End If
                                   Else
358                                    If strSampleID = str标本号 Then
360                                        lngQCID = rsTmp.Fields("ID")
                                       End If
                                   End If
                                
362                                rsTmp.MoveNext
                               Loop
                            
364                            If lngQCID > 0 Then
366                                gstrSQL = "ZL_检验质控记录_EDIT(1," & lngID & "," & lngQCID & ")"
368                                gobjDatabase.ExecuteProcedure gstrSQL, "保存为质控品"
370                                strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 保存为质控品:成功!"
                                      '要自动进行失控计算
372                                   If intAutoQCCalc = 1 Then
374                                     If strQCList(UBound(strQCList)) <> "" Then ReDim Preserve strQCList(UBound(strQCList) + 1)
376                                       strQCList(UBound(strQCList)) = Format(CDate(strDate), "yyyy-MM-dd") & "," & CStr(lngQCID)
                                      End If
                                End If
378                        ElseIf strAutoCheckMan <> "" Then
                              '自动审核
380                            If InStr(1, gstrPrivs, "审核标本") > 0 And bln无主标本 = False Then
                                   
382                                If strAutoCheck(UBound(strAutoCheck)) <> "" Then ReDim Preserve strAutoCheck(UBound(strAutoCheck) + 1)
384                                If int2Verify = 1 Then
386                                    strAutoCheck(UBound(strAutoCheck)) = lngID & "|Zl_检验标本记录_初审报告(" & lngID & ",1,'" & gstrUserName & "')"
                                   Else
388                                    strAutoCheck(UBound(strAutoCheck)) = lngID & "|ZL_检验标本记录_报告审核(" & lngID & ",'" & strAutoCheckMan & "','" & gstrUserCode & "','" & gstrUserName & "')"
                                   End If
                               End If
                           End If
390                    ElseIf intMicrobe = 1 Then
                              '处理微生物只有细菌返回的情况
392                        gstrSQL = "ZL_检验普通结果_BATCHUPDATE(" & lngID & "," & _
                              lngDeviceID & ",'" & strSample & "'," & strSex & "," & _
                              IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                              "0^^^" & "'," & intMicrobe & ")"
394                        gobjDatabase.ExecuteProcedure gstrSQL, "检验结果报告"
396                        strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 保存检验结果:只有细菌返回,保存成功!"
                        End If
                        
                        
398                     If UBound(aTmp) > 0 Then
400                        If Trim(aTmp(1)) <> "" Then
                              '处理图形数据
                               'Call WriteLog("SaveImg", LOG_通讯日志, 0, "开始时间:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
402                            strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 开始保存图像数据" & vbNewLine & aTmp(1)
404                            Call SaveImg(lngDeviceID, lngID, aTmp(1))
                               'Call WriteLog("SaveImg", LOG_通讯日志, 0, "结束时间:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
406                            strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 结束保存图像数据"
                           End If
                        Else
                            strLog = strLog & vbNewLine & "没有图形数据"
                        End If 'End Ubound(atmp)>0
                        
                   Else
                        strLog = strLog & vbNewLine & "标本已审核，不处理!"
                   End If 'blnAuditing
        
               Else
                  strLog = strLog & vbNewLine & "解码的格式不正确，需要至少四个元素"
               End If 'If UBound(aItem) >= 4 Then
           Next
       End If
   
      '计算质控

408    SaveToDataBase = True

410    For i = LBound(strQCList) To UBound(strQCList)
412        If InStr(strQCList(i), ",") > 0 Then
414            Call AutoQCCompute(lngDeviceID, CDate(Split(strQCList(i), ",")(0)), Split(strQCList(i), ",")(1), strAutoCaleInfo)
           End If
       Next
    
      '自动审核
       Dim strInfo As String

416    For i = LBound(strAutoCheck) To UBound(strAutoCheck)
418        If InStr(strAutoCheck(i), "|") > 0 Then
420            lngID = Val(Split(strAutoCheck(i), "|")(0))
422             If CheckSample(lngID) Then
                '检查是否可以审核
424            If VerifyAuditingRule(lngID, strInfo) <> 1 Then

426                strSQL = Split(strAutoCheck(i), "|")(1)
428                gobjDatabase.ExecuteProcedure strSQL, "自动审核" & lngID
430             strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 自动审核 " & lngID
               End If
                End If
           End If
       Next
432    strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 结束入库"
434    strlogs = strLog
       Exit Function
DBError:
436   Call WriteLog("SaveToDataBase", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description)
438 strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " 入库" & CStr(Erl()) & "行出现错误" & Err.Description
End Function



Private Sub AutoQCCompute(ByVal lngDeviceID As Long, ByVal date日期 As Date, ByVal str质控品 As String, ByRef strRetuInfo As String)

        '自动计算质控标本
        ' date日期 :质控计算日期
        ' str质控品 :质控品
        Dim rsTemp As adodb.Recordset, rsTmp As adodb.Recordset, strReturn As String
        On Error GoTo errH
100     gstrSQL = "Select Distinct B.项目id, C.编码, C.中文名, C.英文名" & vbNewLine & _
                  " From 检验质控品 A, 检验质控品项目 B, 诊治所见项目 C" & vbNewLine & _
                  " Where A.ID = B.质控品id And B.项目id = C.ID And A.仪器id = [1] "
        
102     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "LisComm自动计算", lngDeviceID)
104     Do Until rsTmp.EOF
            '计算一段时间
106             gstrSQL = "Select Zl_检验质控记录_Compute(" & lngDeviceID & ", " & rsTmp("项目ID") & ", To_Date('" & Format(date日期, "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & str质控品 & "') From Dual"
108             Set rsTemp = gobjDatabase.OpenSqlRecord(gstrSQL, "LisComm自动计算")

110             If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")  计算过程调用错误！" & vbCrLf
112             If InStr(rsTemp.Fields(0).Value, "出现失控！") > 0 Then
114                 strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")" & rsTemp.Fields(0).Value & vbCrLf

116             ElseIf InStr(rsTemp.Fields(0).Value, "计算完成！") <= 0 Then
118                 If InStr(rsTemp.Fields(0).Value, "按规则未发现警告和失控！") <= 0 Then
120                 strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                    End If
                End If
122         rsTmp.MoveNext
        Loop
124     If Trim(strReturn) <> "" Then
126        strRetuInfo = strReturn
        End If
        Exit Sub
errH:
128    WriteLog "AutoQCCompute", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description
End Sub


Public Function VerifyAuditingRule(lngSampleID As Long, Optional strErrMessage As String) As Integer
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '功能                       审核时检验审核规则
        '参数                       lngSampleID 标本ID; strErrMessage 返回1时的错误提示。
        '返回                       0 正常 1 有结果超出警示值
        '
        '结果标志 3-↑、2-↓、1-正常、4-异常、5-↓↓、6-↑↑
        '
        Dim strSQL As String
        Dim rsTmp As New adodb.Recordset
        Dim int病人id As Integer '
        On Error GoTo errH
        '处理超出警示值的结果
100     strSQL = " select 结果标志 from 检验标本记录 a , 检验普通结果 b " & _
                 " Where a.ID = b.检验标本id and a.id = [1] and (b.结果标志 = 5 Or b.结果标志 = 6)"
102     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, gstrSysName, lngSampleID)
104     If rsTmp.EOF = False Then
106         VerifyAuditingRule = 1: strErrMessage = "  结果超过警示值！"
        End If
       '处理超出警示值的结果
108     strSQL = " select 结果标志 from 检验标本记录 a , 检验普通结果 b " & _
                 " Where a.ID = b.检验标本id and a.id = [1] and (b.结果标志 = 5 Or b.结果标志 = 6)"
110     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, gstrSysName, lngSampleID)
112     If rsTmp.EOF = False Then
114         VerifyAuditingRule = 1: strErrMessage = "  结果超过警示值！"
        End If
        '-- 德阳修改，检验结果全部为空，则提示。
116     strSQL = "Select Count(B.ID) - Sum(Decode(Trim(b.检验结果), Null, 1, 0)) As 结果" & vbNewLine & _
                 "From 检验标本记录 a , 检验普通结果 B Where a.id = b.检验标本ID and  a.id = [1] and nvl(a.微生物标本,0) = 0 "
118     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, gstrSysName, lngSampleID)
120     Do Until rsTmp.EOF
122         If Nvl(rsTmp("结果")) <> "" Then
124             If Val("" & rsTmp!结果) <= 0 Then
126                VerifyAuditingRule = 1: strErrMessage = "  结果全部为空！"
                End If
            End If
128         rsTmp.MoveNext
        Loop
    
    
130     int病人id = gobjDatabase.GetPara("历史病人识别", 100, 1208, 0)
    
132     If VerifyAuditingRule <> 1 And strErrMessage = "" Then
134         strSQL = "Select Zl_检验审核规则_Check(" & lngSampleID & "," & int病人id & ") as 审核结果 From Dual"
136         Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, gstrSysName)
138         If rsTmp.RecordCount <= 0 Then
140             VerifyAuditingRule = 1
142             strErrMessage = "  计算过程调用错误! "
                Exit Function
            End If
144         strErrMessage = "" & rsTmp.Fields(0).Value
146         If strErrMessage <> "" Then VerifyAuditingRule = 1
        End If
        Exit Function
errH:
148     WriteLog "VerifyAuditingRule", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description
End Function

Private Function CheckSample(ByVal lngID As Long) As Boolean
        '审核前检查
        Dim rsTmp As adodb.Recordset, strSQL As String
        On Error GoTo errH
        '11210 权限“未收费审核”，在审核单个病人时，未生效，
100     If InStr(gstrPrivs, "未收费审核") <= 0 Then
102         If CheckChargeState(lngID, False) = False Then
104             WriteLog "InDataBase", LOG_错误日志, lngID, "单据未收费，不能进行审核！"
                Exit Function
            End If
        End If
    
        '21137 已归档报告不能审核
106     strSQL = "Select Decode(病案状态, 1, '1-等待审查', 2, '2-拒绝审查', 3, '3-正在审查', 4, '4-审查反馈', 5, '5-审查归档') As 病案状态" & vbNewLine & _
                "From 检验标本记录 A, 病案主页 B ,病案提交记录 C" & vbNewLine & _
                "Where A.病人id = B.病人id And A.主页id = B.主页id And A.病人来源 = 2 And Nvl(B.病案状态, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                " And b.病人id = c.病人Id and B.主页id = C.主页ID "
108     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "CheckSample", lngID)
110     If rsTmp.EOF = False Then
112         WriteLog "InDataBase", LOG_错误日志, lngID, "病人本次住院的病案已提交审查，不能进行审核！"
            Exit Function
        End If
114     If CheckExesState(lngID) = False Then
116         WriteLog "InDataBase", LOG_错误日志, lngID, "当前住院病人还有划价单未审核，但已出院或预出院！"
            Exit Function
        End If
    
        '检验姓名
118     Call CheckPatientInfo(lngID)
120     CheckSample = True
        Exit Function
errH:
122     Call WriteLog("CheckSample", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description)
End Function


Private Function CheckChargeState(ByVal lngKey As Long, Optional ByVal blnOrder As Boolean = True, Optional ByVal DataMoved As Boolean = False) As Boolean
        '检验收费状态
        Dim strSQL As String
        Dim rs As New adodb.Recordset
        Dim strSQLbak As String
        Dim intPatientType As Integer               '病人来源
        On Error GoTo errH
    
100     CheckChargeState = False
    
102     strSQL = "select 病人来源 from 病人医嘱记录 where id = [1]"
104     Set rs = gobjDatabase.OpenSqlRecord(strSQL, "检验查费用", lngKey)
106     If rs.EOF = True Then Exit Function
108     intPatientType = rs("病人来源")
    
110     If blnOrder Then
112         strSQL = _
                "select NVL(A.记录状态,0) As 记录状态 " & _
                      "from 住院费用记录 A, " & _
                      "( " & _
                           "select No from 病人医嘱发送 where 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE [1] In (ID,相关id))  " & _
                           "Union " & _
                           "select No from 病人医嘱附费 where 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE [1] In (ID,相关id)) " & _
                      ") B " & _
                    "Where A.NO = B.NO "
114         If intPatientType <> 2 Then
116             strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
            End If
        Else
118         strSQL = _
                "select NVL(A.记录状态,0) As 记录状态 " & _
                      "from 住院费用记录 A, " & _
                      "( " & _
                           "select No,记录性质 from 病人医嘱发送 where 医嘱id IN (Select ID From 病人医嘱记录 A,(Select 医嘱id From 检验标本记录 Where ID= [1] Union Select 医嘱id From 检验项目分布 Where 标本id= [1]) B where B.医嘱id In (A.ID,A.相关id) and A.诊疗类别 = 'C' ) " & _
                           "Union " & _
                           "select No,记录性质 from 病人医嘱附费 where 医嘱id IN (Select ID From 病人医嘱记录 A,(Select 医嘱id From 检验标本记录 Where ID= [1] Union Select 医嘱id From 检验项目分布 Where 标本id= [1]) B where B.医嘱id In (A.ID,A.相关id) and A.诊疗类别 = 'C' ) " & _
                      ") B " & _
                    "Where A.NO = B.NO and a.记录性质 = b.记录性质 "
120         If intPatientType <> 2 Then
122             strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
            End If
        End If
    
124     strSQL = strSQL & " Order by 记录状态 "
126     If DataMoved Then
128         strSQL = Replace(strSQL, "住院费用记录", "H住院费用记录")
130         strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
132         strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
134         strSQL = Replace(strSQL, "检验标本记录", "H检验标本记录")
136         strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
138         strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
        End If
    
140     Set rs = gobjDatabase.OpenSqlRecord(strSQL, "mdlLisWork", lngKey)

142     If rs.BOF Then Exit Function
144     If rs("记录状态").Value = 0 Then Exit Function
    
146     CheckChargeState = True
    
        Exit Function
errH:
148     Call WriteLog("CheckChargeState", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description)
End Function

Private Function CheckExesState(lngKey As Long) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '功能:      检查住院病人出院后是否还有划价单需要进行审核
        '参数       标本ID
        '返回       有划价单未审核 = Fasle 没有则 = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim rsTmp As New adodb.Recordset
        On Error GoTo errH
100     CheckExesState = True
    
        '81号系统不生效时不检查
102     If gobjDatabase.GetPara(81, 100) <> 1 Then Exit Function
        
        '当前病人是否已出院或预出院
104     gstrSQL = "select d.no" & vbNewLine & _
                "from (select distinct d.医嘱id" & vbNewLine & _
                "       from 检验标本记录 a, 病人信息 b, 病案主页 c, 检验项目分布 d" & vbNewLine & _
                "       where a.病人id = b.病人id and a.病人id = c.病人id and a.主页id = c.主页id and" & vbNewLine & _
                "             a.id = [1] and a.病人来源 = 2 and (b.出院时间 is not null or c.状态 = 3) and" & vbNewLine & _
                "             a.id = d.标本id) a, 病人医嘱记录 b, 病人医嘱发送 c, 住院费用记录 d" & vbNewLine & _
                "where a.医嘱id in (b.相关id, b.id) and b.id = c.医嘱id and c.记录性质 = d.记录性质 and" & vbNewLine & _
                "      c.no = d.no and d.记录性质 = 2 and d.记录状态 = 0 "
106     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "检验技师工作站-费用状态检查", lngKey)
    
108     CheckExesState = rsTmp.EOF
        Exit Function
errH:
110     Call WriteLog("CheckExesState", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description)
End Function

Private Function CheckPatientInfo(lngSampleID As Long) As Boolean
        Dim rsTmp As New adodb.Recordset
        Dim int提示修正 As Integer '1-提示修正，2-不提示修正，3-不修正

        On Error GoTo errH
    
100     gstrSQL = "Select A.病人来源,A.病人id, A.性别 As 性别1, B.性别 As 性别2, A.年龄 As 年龄1, B.年龄 As 年龄2, A.姓名 As 姓名1, B.姓名 As 姓名2,nvl(a.婴儿,0) as 婴儿 " & vbNewLine & _
                            "From 检验标本记录 A, 病人信息 B" & vbNewLine & _
                            "Where A.病人id = B.病人id And A.ID = [1]"
102     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "CheCkPatientInfo", lngSampleID)
    
        '是婴儿时不进行对比
104     If rsTmp("婴儿") > 0 Then
            Exit Function
        End If
    
    
106     If Nvl(rsTmp("姓名1")) <> Nvl(rsTmp("姓名2")) Or Nvl(rsTmp("性别1")) <> Nvl(rsTmp("性别2")) Or _
            Nvl(rsTmp("年龄1")) <> Nvl(rsTmp("年龄2")) Then
        
108         int提示修正 = 1
        
110         If rsTmp("病人来源") = 4 Then
112             int提示修正 = Val(gobjDatabase.GetPara("体检病人信息不一致的处理方式", glngSys, 1208, True, 1))
114         ElseIf rsTmp("病人来源") = 3 Then
116             int提示修正 = Val(gobjDatabase.GetPara("院外病人信息不一致的处理方式", glngSys, 1208, True, 1))
118         ElseIf rsTmp("病人来源") = 2 Then
120             int提示修正 = Val(gobjDatabase.GetPara("住院病人信息不一致的处理方式", glngSys, 1208, True, 1))
122         ElseIf rsTmp("病人来源") = 1 Then
124             int提示修正 = Val(gobjDatabase.GetPara("门诊病人信息不一致的处理方式", glngSys, 1208, True, 1))
            End If
        
126         If int提示修正 = 1 Then
128             WriteLog "InDataBase", LOG_错误日志, lngSampleID, "发现检验信息中的病人信息和病人信息中病人信息不一致!"
130         ElseIf int提示修正 = 2 Then
132             gstrSQL = "zl_检验标本记录_Update(" & lngSampleID & ",'" & Nvl(rsTmp("姓名2")) & "','" & Nvl(rsTmp("性别2")) & _
                                             "','" & Nvl(rsTmp("年龄2")) & "')"
134             gobjDatabase.ExecuteProcedure gstrSQL, "CheckPatientInfo"
            End If
136         CheckPatientInfo = True
            Exit Function
        End If
138     CheckPatientInfo = False
    
        Exit Function
errH:
140     Call WriteLog("CheckPatientInfo", LOG_错误日志, Err.Number, CStr(Erl()) & "行出现错误：  " & Err.Description)
End Function


Public Function TestFTP(ByVal strUser As String, ByVal strPassWord As String, _
                            ByVal strDevAdress As String, ByVal strFtpPath As String) As String
                            
    Dim FtpNet As New clsFtp, strPath As String, strTmpPath As String           'FTP类
    Dim lngFileNo As Long
    strPath = Format(Now, "yyyymmddHHMMSS")
    strTmpPath = IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & "temp.txt"
    lngFileNo = FreeFile
    Open strTmpPath For Output As lngFileNo
    Close lngFileNo
    If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir(strFtpPath, "FTP测试" & strPath) > 0 Then
            TestFTP = "在FTP上不能创建目录！"
        Else
            If FtpNet.FuncUploadFile(strFtpPath, strTmpPath, "temp.txt") > 0 Then
                TestFTP = "上传文件失败"
            Else
                FtpNet.FuncFtpDisConnect '先断开，再删除，不然删不掉
                If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) <= 0 Then
                     TestFTP = "FTP不能连接！"
                ElseIf FtpNet.FuncFtpDelDir(strFtpPath, "FTP测试" & strPath) > 0 Then
                    TestFTP = "在FTP上不能删除目录"
                Else
                    TestFTP = ""
                End If
            End If
        End If
    Else
        TestFTP = "不能连接FTP！"
    End If
    FtpNet.FuncFtpDisConnect
    Set FtpNet = Nothing
    Kill strTmpPath
End Function

Private Function DownFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strFile As String) As String
        '从FTP服务器下载文件。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpFile :FTP上的文件。
        'strFile    :本地文件全路径。
        '返回：空串表示成功，否则为错误提示。
        Dim objFtp As New clsFtp, lngReturn As Long, strFtpFileName As String, strLocaFile As String
        Dim strFTPDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFile = "请指定要下载的文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFTPDir = Replace(strFtpFile, "/" & strFtpFileName, "")
108     strLocaFile = strFile
110     If strLocaFile = "" Then
112         DownFile = "请指定下载的文件保存到何处！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFile = "要下载的文件已存在！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFile = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFile = "不能连接服务器！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFTPDir)
130     If lngReturn <> 0 Then
132         DownFile = "不能进入指定的目录，可能是权限不足或服务器上无此目录！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownloadFile(strFTPDir, strLocaFile, strFtpFileName)
136     If lngReturn <> 0 Then
138         DownFile = "下载失败，可能是权限不足或服务器上无此文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFile = CStr(Erl()) & "行，" & Err.Description
End Function

Private Function UploadFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                            ByVal strFtpPath As String, ByVal strFile As String, Optional strNewFileName As String) As String
        '按本地文件名上传文件到FTP服务器。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpPath :FTP上的目录，无目录会自动创建。
        'strFile    :本地文件全路径。
        'strNewFileName: 传到FTP上后的文件名，为空则按本地文件名保存
        '返回：空串表示成功，否则为错误提示。
    
        Dim objFtp As New clsFtp, lngReturn As Long, strFileName As String, strLocaFile As String
        On Error GoTo errH
    
    
100     If Left(strFtpPath, 1) = "/" Then strFtpPath = Mid$(strFtpPath, 2)
    
102     If strServer = "" Then
104         UploadFile = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
106     strLocaFile = strFile
108     If Dir(strLocaFile) = "" Then
110         UploadFile = "文件" & strLocaFile & "不存在!"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        If strNewFileName = "" Then
112         strFileName = Split(strLocaFile, "\")(UBound(Split(strLocaFile, "\")))
        Else
            strFileName = strNewFileName
        End If
114     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
116     If lngReturn <> 0 Then
            '检查目录是否存在
118         lngReturn = objFtp.FuncChangeDir(strFtpPath)
120         If lngReturn <> 0 Then
122             lngReturn = objFtp.FuncFtpMkDir("/", strFtpPath)
124             If lngReturn <> 0 Then
126                 UploadFile = "创建目录失败！可能是权限不足！"
                    objFtp.FuncFtpDisConnect
                    Set objFtp = Nothing
                    Exit Function
                End If
            End If
        
128         lngReturn = objFtp.FuncUploadFile("/" & strFtpPath, strLocaFile, strFileName)
130         If lngReturn <> 0 Then
132             UploadFile = "上传文件失败，可能是权限不足！"
                objFtp.FuncFtpDisConnect
                Set objFtp = Nothing
                Exit Function

            Else
134             UploadFile = ""
            End If
        Else
136         UploadFile = "不能连接服务器！"
        End If
        objFtp.FuncFtpDisConnect
        Set objFtp = Nothing
        Exit Function
errH:
138     UploadFile = CStr(Erl()) & "行，" & Err.Description
End Function

Private Function ItemLocalSort(ByVal strItems As String, ByRef varItems As Variant) As String
    '指标排序。2011-12-07 死锁问题修改，5/5 - 指标排序
    Dim strReturn As String, strTmp As String
    
    Dim i As Integer, varTmp As Variant
    Dim x As Integer
    '只有一个结果，不用排序
    If InStr(strItems, "|") <= 0 Then
        ItemLocalSort = strItems
        Exit Function
    End If
    
    varTmp = Split(strItems, "|")
    For i = 0 To UBound(varItems, 2)
        For x = LBound(varTmp) To UBound(varTmp)
            strTmp = Split(varTmp(x), "^")(0)
            If CStr(varItems(1, i)) = strTmp Then
                strReturn = strReturn & "|" & varTmp(x)
                Exit For
            End If
        Next
    Next
    
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    ItemLocalSort = strReturn
End Function
