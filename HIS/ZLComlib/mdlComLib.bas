Attribute VB_Name = "mdlComLib"
Option Explicit
'**************************
'       OEM代号
'
'医业  D2BDD2B5
'托普  CDD0C6D5
'中软  D6D0C8ED
'创智  B4B4D6C7
'金康泰 BDF0BFB5CCA9
'宝信  B1A6D0C5
'**************************

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gcnOracleOLEDB  As ADODB.Connection  '公共数据库连接OLEDB方式，当读取LOB对象时一次读取
Public gobjComLib As clsComLib
Public gobjRegister As Object               '注册授权部件zlRegister

Public g_AutoConnect    As Boolean          '通过该变量将不同实例中gblnAutoConnect的值共享
Public g_NodeNo As String                   '通过该变量将不同实例中gstrNodeNo的值共享
Public g_NodeName As String                 '通过该变量将不同实例中gstrNodeName的值共享
Public glngSessionID As Long
Public gstrComputerName As String
Public gstrSysName As String                '系统名称
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrAppsoft As String                'APPSOFT路径
Public gstrHelpPath As String
Public gblnOK As Boolean
Public gstrDBUser As String
Public gfrmMain As Object '导航台窗体
Public gblnShow As Boolean

Public gobjLogFile As FileSystemObject
Public gobjLogText As TextStream
Public gobjPlanExFile As FileSystemObject
Public gobjPlanExText As TextStream

Public gblnSQLTest As Boolean
Public gblnSQLLog As Boolean
Public gblnSQLPlan As Boolean   '性能监控模式

Public gstrSysUser As String
Public gcnSysConn As ADODB.Connection 'sys链接
Public gcnSysOLEDB As ADODB.Connection 'sys链接,OLEDB方式
Public gblnSys As Boolean
Public gstrRecentSQL As String  '最近执行的SQL语句

Public grsDiagConn As ADODB.Recordset '存放申请单诊断关联

'系统参数
Public gblnRunLog As Boolean '是否记录使用日志
Public gblnErrLog As Boolean '是否记录运行错误

Public grsParas As ADODB.Recordset '系统参数表缓存
Public grsUserParas As ADODB.Recordset '系统参数表缓存
Public grsDeptParas As ADODB.Recordset    '系统参数部门缓存
Public grsUserInfo As ADODB.Recordset  '当前用户的人员和部门信息缓存
Public gcolMoveDate As Collection    '历史数据的转出日期

Public gclsMipClient As clsMipClient
Public gcllComlibs  As Collection       '所有的Comlib对象实例集合

Public gcolWriteLog As Collection '存储各种类型的日志的对象
Public gstrLastLogName As String              '缓存上一次使用的日志名称
Public gobjLastLog As TextStream          '缓存上一次使用的日志对象
Public gstrLastLogInfoHeader As String        '缓存上一次使用的日志头
Public gcolLastLogInfoHeader As Collection    '缓存各种类型的日志上一次存储的日志头

Public glngPatiTypeWinProc As Long               '原始消息句柄
Public gclsPDF          As clsPDF       'PDF输出类全局缓存，以便同一个进程共用一个实例

Public Const MSTR_DBLINK_KEY As String = "zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325"  '加密解密秘钥
'连接方式
Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
    OriginalConnection = 9
End Enum

Public Function FlexScroll(ByVal hWnd As Long, ByVal wMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long
'支持滚轮的滚动
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '向下滚
            gobjComLib.zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '向上滚
            gobjComLib.zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPatiTypeWinProc, hWnd, wMsg, wParam, lParam)
End Function

Public Sub ChangeAllIntanceConn(cnMain As ADODB.Connection)
'功能：同步clsComlib各个实例中的mcnoracle进行事务重试按钮的禁用。
'参数：同步各个实例的连接
    Dim objComlib As clsComLib
    If Not gcllComlibs Is Nothing Then
        For Each objComlib In gcllComlibs
            Call objComlib.ChangeIntanceConn(cnMain)
        Next
    End If
End Sub

Public Function SQLObject(ByVal strSQL As String) As String
'功能：分析SQL语句所用到的对象名
'参数：strSQL=要分析的原始SQL语句
'返回：SQL语句所访问到的对象名,如"部门表,病人费用记录,ZLHIS.人员表"
'说明：1.与Oracle SELECT语句兼容
'      2.如果SQL语句中的对象名前加有所有者前缀,则该前缀不会被截取
'      3.需要函数TrimChar;TrueObject的支持
    Dim intB As Integer, intE As Integer, intL As Integer, intR As Integer
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Integer, j As Integer
    
    On Error GoTo errh
    
    '大写化及去除多余的字符
    strAnal = UCase(TrimChar(strSQL))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
    '先分解处理嵌套子查询
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB '匹配的左右括号位置
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                If intE - intB - 1 <= 0 Then
                    '对于非子查询,将括号换成其它符号,以使循环继续
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '子查询语句
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '将该子查询部份作为为特殊对象名
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "嵌套查询")
                    '递归分析
                    strObject = strObject & "," & SQLObject(strSub)
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '无匹配右括号
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '分解分析
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '从第一个From后面部份开始
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject, "," & strTrue) = 0 And strTrue <> "嵌套查询" Then
                strObject = strObject & "," & strTrue
            End If
        Next
    Next
    '完成
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errh:
    Err.Clear
End Function

Private Function TrimChar(Str As String) As String
'功能:去除字符串中连续的空格和回车(含两头的空格,回车),不去除TAB字符,哪怕是连续的
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(Str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(Str)
    i = InStr(strTmp, "  ")
    Do While i > 0
        strTmp = Left(strTmp, i) & Mid(strTmp, i + 2)
        i = InStr(strTmp, "  ")
    Loop
    
    i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Do While i > 0
        strTmp = Left(strTmp, i + 1) & Mid(strTmp, i + 4)
        i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Loop
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Private Function TrueObject(ByVal strObject As String) As String
'功能：SQLObject函数的子函数,用于去除对象名中的无用字符
    Dim i As Integer
    '寻找第一个正常字符位置
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    '寻找后面第一个非正常字符
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function

Public Function SeekCboIndex(objCbo As Object, varData As Variant) As Long
'功能：由ItemData或Text查找ComboBox的索引值
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '先精确查找
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf gobjComLib.zlCommFun.GetNeedName(objCbo.List(i)) = varData Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '再模糊查找
            For i = 0 To objCbo.ListCount - 1
                If InStr(objCbo.List(i), varData) > 0 Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    Else
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = varData Then
                SeekCboIndex = i: Exit Function
            End If
        Next
    End If
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'功能：判断某个ADO字段数据类型是否与指定字段类型是同一类(如数字,日期,字符,二进制)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function

'--------------------------------------------------
'功能：检查是否为网络断开或ADO断开引发的错误!
'返回：True:恢复连接成功 False恢复连接失败
'--------------------------------------------------
Public Function CheckAdoConnction(ByRef blnStatus As Boolean) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnAdoErr As Boolean
    Dim strError As String
    On Error GoTo Errhand
    blnAdoErr = False
    blnStatus = False

    On Error GoTo Errhand
    Err = 0
    DoEvents
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    gcnOracle.Open
    If blnAdoErr Then
        'True '是ORA-12560不能与ORACLE连接引起
        CheckAdoConnction = True
    Else
        'False '可以正常连接
        CheckAdoConnction = False
        On Error Resume Next
        '重连后判断客户端是否被禁止使用，若被禁止，则自动断开连接
        strSQL = "Select NVL(禁止使用,0)  禁止使用 From zlClients Where 工作站=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckAdoConnction", gstrComputerName)
        If Err.Number <> 0 Then Err.Clear
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                If rsTmp!禁止使用 = 1 Then
                    If gcnOracle.State = adStateOpen Then gcnOracle.Close
                    CheckAdoConnction = True
                    Call SaveSetting("ZLSOFT", "公共全局\网络断网自动重连", "AutoConnect", 0)
                    MsgBox "当前工作站已经被管理员禁用，请联系管理员解除禁用并重新登录！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    Exit Function
Errhand:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        If InStr(Err.Description, "ORA-12560") > 0 Then
            blnAdoErr = True
            Resume Next
        ElseIf InStr(Err.Description, "ORA-12543") > 0 Then
            blnAdoErr = True
            Resume Next
        Else
            '其他错误引发的网络问题
            CheckAdoConnction = True
            blnStatus = True
        End If
    Else
        CheckAdoConnction = False
    End If
End Function

Public Function CheckErrConnectInfo(ByVal strErrNum As String, ByVal strNote As String, ByVal strErrInfo As String, ByVal intType As Integer) As Boolean
    '------------------------------------------------
    '功能： 按照类型IntType(1,2)检查vb和oralce返回的具体错误信息，来判断是否为网络断开引发的错误或者是其他的错误引发
    '参数： strNote错误信息,strErrInfo错误详细信息,intType 错误类型 1：VB错误 2:ORACLE错误
    '返回： True:网络引发的错误 False:其他错误
    '------------------------------------------------
    Dim strTemp As String
    Dim i As Integer
    If intType = 1 Then
        'VB具体错误
   
        If InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02396") > 0 Then '超出最大空闲时间, 请重新连接 IDLE_TIME profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02399") > 0 Then '超出最大连接时间, 您将被注销 connect_time profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-01012") > 0 Then '没有登录
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-00028") > 0 Then '会话被终止
            CheckErrConnectInfo = True
        Else
            If strErrNum = "3709" Then '3709描述：连接无法用于执行此操作。在此上下文中它可能已被关闭或无效。单独处理
                CheckErrConnectInfo = True
            Else
                If strNote = "不确定的错误" Then
                    CheckErrConnectInfo = True
                Else
                    CheckErrConnectInfo = False
                End If
            End If
        End If
    Else
        'ORACLE具体错误
        If InStr(strErrInfo, "SQLSetConnectAttr") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12543") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02396") > 0 Then '超出最大空闲时间, 请重新连接 IDLE_TIME profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02399") > 0 Then '超出最大连接时间, 您将被注销 connect_time profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-01012") > 0 Then '没有登录
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-00028") > 0 Then '会话被终止
            CheckErrConnectInfo = True
        Else
            CheckErrConnectInfo = False
        End If
    End If
End Function

Public Function GetGUID() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim udtGUID As GUID
    
    On Error GoTo Errhand
    
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
                String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
                String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
                IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
                IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
                IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
                IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
                IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
                IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
                IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
                IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
    
    Exit Function
Errhand:
    'MsgBox Err.Description
End Function

Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, ByVal dwData As Long) As Long
     Dim monitorInf As MONITORINFO
     Dim R As RECT
     
     ReDim Preserve gMonitors(UBound(gMonitors) + 1)
     


     'initialize   the   MONITORINFO   structure
     monitorInf.cbSize = Len(monitorInf)
     'Get   the   monitor   information   of   the   specified   monitor
     GetMonitorInfo hMonitor, monitorInf
     'write   some   information   on   teh   debug   window

    
     gMonitors(UBound(gMonitors) - 1).monitorHandle = hMonitor
     gMonitors(UBound(gMonitors) - 1).monitorInf = monitorInf
     
     '这里必须返回1，以便后续执行
     MonitorEnumProc = 1
  End Function

Public Function GetMonitorIndex(ByVal windowHandle As Long) As Long
'    '******************************************************************************************************************
'    '功能：获得监视器ID
'    '参数：windowHandle
'    '返回：监视器ID
'    '******************************************************************************************************************

    Dim i As Integer

    Dim monitorCount As Integer
    monitorCount = 0

    On Error GoTo GetMonitorInf
      monitorCount = UBound(gMonitors)
GetMonitorInf:
      If monitorCount <= 1 Then
        ReDim Preserve gMonitors(1)
        gMonitors(1).monitorHandle = -1

        EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
      End If


    For i = 1 To UBound(gMonitors)
      If MonitorFromWindow(windowHandle, MONITOR_DEFAULTTONEAREST) = gMonitors(i).monitorHandle Then
        GetMonitorIndex = i - 1
        Exit Function
      End If
    Next i

    GetMonitorIndex = -1

End Function

'解密函数
Public Function Decipher(ByVal password As String, ByVal from_text As String) As String
    '解密
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    password = Base64Encode(password) & "WIZARDPAGE"
    
    Dim offset As Long
    Dim str_len As Integer
    Dim i As Integer
    Dim ch As Integer
    offset = NumericPassword(password)
    Rnd -1
    Randomize offset

    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            Decipher = Decipher & Chr$(ch)
        End If
    Next i
End Function


'加解密字符串函数,不支持中文
Private Function Base64Encode(InStr1 As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim mInByte(3)     As Byte, mOutByte(4)       As Byte
    Dim myByte     As Byte
    Dim i     As Integer, LenArray       As Integer, j       As Integer
    Dim myBArray()     As Byte
    Dim OutStr1     As String
    myBArray() = StrConv(InStr1, vbFromUnicode)
    LenArray = UBound(myBArray) + 1
    For i = 0 To LenArray Step 3
      If LenArray - i = 0 Then
        Exit For
      End If
      If LenArray - i = 2 Then
        mInByte(0) = myBArray(i)
        mInByte(1) = myBArray(i + 1)
        Base64EncodeByte mInByte, mOutByte, 2
      ElseIf LenArray - i = 1 Then
        mInByte(0) = myBArray(i)
        Base64EncodeByte mInByte, mOutByte, 1
      Else
        mInByte(0) = myBArray(i)
        mInByte(1) = myBArray(i + 1)
        mInByte(2) = myBArray(i + 2)
        Base64EncodeByte mInByte, mOutByte, 3
      End If
      For j = 0 To 3
        OutStr1 = OutStr1 & Chr(mOutByte(j))
      Next j
    Next i
    Base64Encode = OutStr1
    
End Function

Private Sub Base64EncodeByte(mInByte() As Byte, mOutByte() As Byte, Num As Integer)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim tByte     As Byte
    Dim i     As Integer
    If Num = 1 Then
      mInByte(1) = 0
      mInByte(2) = 0
    ElseIf Num = 2 Then
      mInByte(2) = 0
    End If
    tByte = mInByte(0) And &HFC
    mOutByte(0) = tByte / 4
    tByte = ((mInByte(0) And &H3) * 16) + (mInByte(1) And &HF0) / 16
    mOutByte(1) = tByte
    tByte = ((mInByte(1) And &HF) * 4) + ((mInByte(2) And &HC0) / 64)
    mOutByte(2) = tByte
    tByte = (mInByte(2) And &H3F)
    mOutByte(3) = tByte
    For i = 0 To 3
      If mOutByte(i) >= 0 And mOutByte(i) <= 25 Then
        mOutByte(i) = mOutByte(i) + Asc("A")
      ElseIf mOutByte(i) >= 26 And mOutByte(i) <= 51 Then
        mOutByte(i) = mOutByte(i) - 26 + Asc("a")
      ElseIf mOutByte(i) >= 52 And mOutByte(i) <= 61 Then
        mOutByte(i) = mOutByte(i) - 52 + Asc("0")
      ElseIf mOutByte(i) = 62 Then
        mOutByte(i) = Asc("+")
      Else
        mOutByte(i) = Asc("/")
      End If
    Next i
    If Num = 1 Then
      mOutByte(2) = Asc("=")
      mOutByte(3) = Asc("=")
    ElseIf Num = 2 Then
      mOutByte(3) = Asc("=")
    End If
End Sub

Private Function NumericPassword(ByVal password As String) As Long
    Dim value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    str_len = Len(password)
    For i = 1 To str_len
        ch = Asc(Mid$(password, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = value
End Function

Public Function IsOLEDBConnection(ByVal cnMain As ADODB.Connection) As Boolean
'功能：判断当前连接是否是OraOLEDB连接
'根据Provider来判断，存在两种方式
'方式一：'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
'方式二：
'.Provider = "OraOLEDB.Oracle"
'.Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUserName, strPassWord
'这两种方式均会自动设置.Provider属性
    '使用Like是因为可能后面增加版本如OraOLEDB.Oracle.1
    If UCase(cnMain.Provider) Like "ORAOLEDB.ORACLE*" Then
        IsOLEDBConnection = True
    End If
End Function

