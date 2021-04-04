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

Public Const FSHIFT = 4
Public Const FCONTROL = 8

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gobjComLib As clsComLib
Public gobjRegister As Object               '注册授权部件zlRegister

Public g_NodeNo As String
Public gstrSysName As String                '系统名称
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrHelpPath As String
Public gblnOK As Boolean
Public gstrDBUser As String
Public gstrComputerName As String
Public gfrmMain As Object '导航台窗体
Public gblnShow As Boolean

Public gobjLogFile As FileSystemObject
Public gobjLogText As TextStream

'系统参数
Public gblnRunLog As Boolean '是否记录使用日志
Public gblnErrLog As Boolean '是否记录运行错误

Public grsParas As ADODB.Recordset '系统参数表缓存
Public grsUserParas As ADODB.Recordset '系统参数表缓存
Public grsUserInfo As ADODB.Recordset  '当前用户的人员和部门信息缓存
Public gcolMoveDate As Collection    '历史数据的转出日期

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
    
    On Error GoTo errH
    
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
errH:
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

Public Function AdjustStr(Str As String) As String
'功能：将含有"'"符号的字符串调整为Oracle所能识别的字符常量
'说明：自动(必须)在两边加"'"界定符。

    Dim i As Long, strTmp As String
    
    If InStr(1, Str, "'") = 0 Then AdjustStr = "'" & Str & "'": Exit Function
    
    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "'" Then
            If i = 1 Then
                strTmp = "CHR(39)||'"
            ElseIf i = Len(Str) Then
                strTmp = strTmp & "'||CHR(39)"
            Else
                strTmp = strTmp & "'||CHR(39)||'"
            End If
        Else
            If i = 1 Then
                strTmp = "'" & Mid(Str, i, 1)
            ElseIf i = Len(Str) Then
                strTmp = strTmp & Mid(Str, i, 1) & "'"
            Else
                strTmp = strTmp & Mid(Str, i, 1)
            End If
        End If
    Next
    AdjustStr = strTmp
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
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strcnOracle As String
    Dim blnAdoErr As Boolean
    On Error GoTo errHand
    blnAdoErr = False
    blnStatus = False

    strcnOracle = CStr(gcnOracle) '原有连接串需要保存,否则CLOSE后,原有连接串会被改变
    On Error GoTo errHand
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Open
    End With

    If blnAdoErr Then
        'True '是ORA-12560不能与ORACLE连接引起
        CheckAdoConnction = True
    Else
        'False '可以正常连接
        CheckAdoConnction = False
    End If
    Exit Function
errHand:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        If InStr(Err.Description, "ORA-12560") > 0 Then
            gcnOracle = strcnOracle
            blnAdoErr = True
            Resume Next
        ElseIf InStr(Err.Description, "ORA-12543") > 0 Then
            gcnOracle = strcnOracle
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

'--------------------------------------------------
'功能：关闭ADO连接
'返回：True:关闭连接成功 False关闭连接失败
'--------------------------------------------------
Public Function CloseAdoConnction() As Boolean
    '------------------------------------------------
    '功能： 关闭数据库
    '参数：
    '返回： 关闭数据库，返回True；失败，返回False
    '------------------------------------------------
    Err = 0
    On Error Resume Next
    gcnOracle.Close
    CloseAdoConnction = True
    Err = 0
    
End Function

Private Function GetActiveConnectionInfo(ByVal strcnOracle As String, ByRef strServerName As String, ByRef strUserName As String, ByRef strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 分析ORACLE串中的 服务器，用户名，密码
    '参数：
    '返回： 成功失败，返回True；失败，返回False
    '------------------------------------------------
    Dim i As Integer
    Dim strTemp As String
    If strcnOracle = "" Then Exit Function
    strServerName = ""
    strUserName = ""
    strUserPwd = ""
    '获取 strServerName
    i = InStrRev(strcnOracle, "Server=", -1)
    If i > 0 Then
        strTemp = Right(strcnOracle, Len(strcnOracle) - i - 6)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strServerName = Left(strTemp, i - 2)
        End If
    End If
    
    '获取 strUserName
    i = InStrRev(strcnOracle, "User ID=", -1)
    If i > 0 Then
        strTemp = Right(strcnOracle, Len(strcnOracle) - i - 7)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strUserName = Left(strTemp, i - 1)
        End If
    End If
    
    '获取 strUserPwd
    i = InStrRev(strcnOracle, "Password=", -1)
    If i > 0 Then
        strTemp = Right(strcnOracle, Len(strcnOracle) - i - 8)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strUserPwd = Left(strTemp, i - 1)
        End If
    End If
    
    If strServerName <> "" And strUserName <> "" And strUserPwd <> "" Then
        GetActiveConnectionInfo = True
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
        Else
            CheckErrConnectInfo = False
        End If
    End If
End Function
