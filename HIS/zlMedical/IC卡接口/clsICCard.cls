VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'变量、API、函数申明
'
'--------------------------------------------------------------------------
'1、接口说明
'
'   此类是用于IC卡操作，为广西玉柴医院职工医疗卡接口部件

'2、接口函数过程清单：

'   1)OpenParamDialog       打开/显示IC卡参数对话框
'   2)InitCard              初始化IC卡
'   3)CloseCard             关闭IC卡
'   4)GetCounter            获取IC卡计数器值
'   5)HaveCard              检查卡座是否有IC卡
'   6)CardTypeOK            检查当前卡是否为SLE4442IC卡
'   7)SetPatient            写病人信息到IC卡
'   8)GetPatient            从IC卡读病人信息
'
'3、附带文件
'
'   1)mwic_32.dll           MW-ET-G读卡器和SLE4442IC卡驱动
'
'4、编写记录
'
'   1)2006年03月29日 By 陈福容  首次开发
'   2)
'
'5、接口项目约定
'    门诊号
'    住院号
'    姓名                   本卡支持，8位
'    性别                   本卡支持，2位
'    年龄
'    出生日期               本卡支持，8位
'    出生地点
'    身份证号               本卡支持，18位
'    身份
'    职业
'    民族
'    国籍
'    学历
'    婚姻状况
'    家庭地址
'    家庭电话
'    户口邮编
'    联系人姓名
'    联系人关系
'    联系人地址
'    联系人电话
'    工作单位
'    单位电话
'    人员编码               本卡支持，8位
'    人员性质               本卡支持，1位
'    用工性质               本卡支持，1位
'    工作岗位               本卡支持，3位
'    单位编码               本卡支持，3位
'    血型                   本卡支持，2位
'    工作日期               本卡支持，8位
'
'6、接口方法
'
'   采用数组传入传出，调用时数组
'
'--------------------------------------------------------------------------
'
'1、局部变量申明
Private mlngICDec As Long



'######################################################################################################################
'变量、API、函数申明
'2.公共变量申明

'======================================================================================================================
'接口函数/过程

Public Function OpenParamDialog(Optional ByVal frmParent As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：打开/显示IC卡参数对话框
    '参数：
    '返回：初始成功返回TRUE；否则返回FALSE
    '------------------------------------------------------------------------------------------------------------------
    
    OpenParamDialog = frmParam.ShowMe(frmParent)
    
End Function

Public Function InitCard() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化IC卡
    '参数：
    '返回：初始成功返回TRUE；否则返回FALSE
    '------------------------------------------------------------------------------------------------------------------
    Dim intComPort As Integer
    Dim intBaud As Integer
    
    If mlngICDec > 0 Then
        InitCard = True
        Exit Function
    End If
    
    '通讯端口
    Select Case UCase(Trim(GetSetting(AppName:="ZLSOFT", Section:="IC卡参数", key:="端口", Default:="COM1")))
    Case "COM1"
        intComPort = 1
    Case "COM2"
        intComPort = 2
    Case "COM3"
        intComPort = 3
    Case "COM4"
        intComPort = 4
    Case Else
        intComPort = 0
    End Select
    
    '波特率
    intBaud = Val(GetSetting(AppName:="ZLSOFT", Section:="IC卡参数", key:="波特率", Default:="9600"))
    
    mlngICDec = ic_init(intComPort - 1, intBaud)

    If mlngICDec <= 0 Then
        Call ErrorCenter(100)
        Exit Function
    End If
    
    If CheckPsw = False Then Exit Function
    
    InitCard = True
End Function

Public Function CloseCard() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：关闭IC卡
    '参数：
    '返回：成功返回TRUE；否则返回FALSE
    '------------------------------------------------------------------------------------------------------------------
    
    'IC卡卡机没准备好或没初始化 ,不能关闭IC卡机！
    If mlngICDec <= 0 Then
        CloseCard = True
        Exit Function
    End If
    
    Call ic_exit(mlngICDec)
    
    mlngICDec = 0
    
    CloseCard = True
    
End Function

Public Function GetCounter() As Integer
    '------------------------------------------------------------------------------------------------------------------
    '功能：获取IC卡计数器值
    '参数：
    '返回：计数器值
    '------------------------------------------------------------------------------------------------------------------
    Dim intCounter As Integer
    
    If mlngICDec < 0 Then Exit Function
    
    Call rsct_4442(mlngICDec, intCounter)
    
    GetCounter = intCounter

End Function

Public Function GetCardNo() As String
    '------------------------------------------------------------------------------------------------------------------
    '功能：获取IC卡卡号
    '参数：
    '返回：IC卡卡号
    '------------------------------------------------------------------------------------------------------------------
    Dim intRet As Integer
    Dim strResult As String * 50
    Dim bytResult(4) As Byte
    
    '检查是否已初始化了卡，如没有则初始化
    If mlngICDec = 0 Then
        If InitCard = False Then Exit Function
    End If
    
    '检验卡（类型、是否插入）
    If HaveCard = False Then Exit Function
    If CardTypeOK = False Then Exit Function
    
    intRet = srd_4442_hex(mlngICDec, &H1B, 4, bytResult(0))
    If intRet < 0 Then GoTo errHand
        
    intRet = hex_asc(bytResult(0), strResult, 8)
    If intRet < 0 Then GoTo errHand
    
    GetCardNo = GetSubStr(strResult, 1, 8)
    If UCase(GetCardNo) = "FFFFFF" Then GetCardNo = ""
    
    Exit Function
    
errHand:
    ShowSimpleMsg "获取IC卡的卡号失败！"
End Function
        
Public Function SetCardNo(ByVal strNo As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：设置IC卡卡号
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim bytResult(4) As Byte
    Dim intRet As Integer
    
    '加密卡号
    strNo = Lpad(strNo, 8)
    
    intRet = asc_hex(strNo, bytResult(0), 4)
    If intRet < 0 Then GoTo errHand
        
    '写卡操作
    intRet = swr_4442_hex(mlngICDec, &H1B, 4, bytResult(0))
    If intRet < 0 Then GoTo errHand
    
    SetCardNo = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg "设置IC卡的卡号失败！"
End Function

Public Function HaveCard() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：检查卡座是否有卡
    '参数：
    '返回：有返回TRUE；否则返回FALSE
    '------------------------------------------------------------------------------------------------------------------
    Dim intRet As Integer
    Dim intStatus As Integer
    
    On Error GoTo errHand
    
    If mlngICDec = 0 Then Exit Function
    
    intStatus = 0
    intRet = get_status(mlngICDec, intStatus)
    If intRet <> 0 Then
        Call ErrorCenter(300)
        Exit Function
    End If
        
    HaveCard = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg "检查卡座是否插有IC卡失败！"
End Function

Public Function CardTypeOK() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：检查当前卡是否为SLE4442IC卡
    '参数：
    '返回：是返回TRUE；否则返回FALSE
    '------------------------------------------------------------------------------------------------------------------
    Dim intRet As Integer
        
    On Error GoTo errHand
    
    If mlngICDec = 0 Then Exit Function
    
    
    intRet = chk_4442(mlngICDec)
    If intRet <> 0 Then
        Call ErrorCenter(400)
        Exit Function
    End If
    
    CardTypeOK = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg "检查IC卡类型是否正确失败！"
End Function

Public Function SetPatient(strInfo() As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：写病人信息到IC卡(病人ID,门诊号,住院号,姓名,性别,年龄,出生日期,出生地点,身份证号,身份,职业,民族,国籍,学历,
    '                       婚姻状况 家庭地址,家庭电话,户口邮编,联系人姓名,联系人关系,联系人地址,联系人电话,工作单位,单位电话)
    '参数：格式："门诊号=123344","住院号=34567"
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim lngPos As Long
    Dim strItem As String
    Dim strValue As String
    Dim intRet As Integer
    Dim intLen As Integer
    Dim intOffset As Integer
    Dim strResult As String
    Dim strAry(1 To 11) As String        '存储顺序为：m_xm, m_xb, m_blood, m_csny, m_gzny, m_gzdw, m_gzgw, m_ryxz, m_ygxz,人员编码,身份证
    Dim strFunction As String
    
    Dim bytResult1(12) As Byte
    Dim bytResult2(13) As Byte
    
    Dim strInfo_Old() As String
    
    On Error GoTo errHand
    
    '先读出原来的信息，再更新信息，这样便于本次没有更新的信息被写为了空值
    If GetPatient(strInfo_Old) = False Then Exit Function
    
    strAry(1) = Lpad(GetAryValue(strInfo_Old, "姓名"), 8)
    strAry(2) = Lpad(GetAryValue(strInfo_Old, "性别"), 2)
    strAry(3) = Lpad(GetAryValue(strInfo_Old, "血型"), 2)
    strAry(4) = DString(GetAryValue(strInfo_Old, "出生日期"))
    strAry(5) = DString(GetAryValue(strInfo_Old, "工作日期"))
    strAry(6) = Lpad(GetAryValue(strInfo_Old, "单位编码"), 3)
    strAry(7) = Lpad(GetAryValue(strInfo_Old, "工作岗位"), 3)
    strAry(8) = Lpad(GetAryValue(strInfo_Old, "人员性质"), 1)
    strAry(9) = Lpad(GetAryValue(strInfo_Old, "用工性质"), 1)
    strAry(10) = Lpad(GetAryValue(strInfo_Old, "人员编码"), 8)
    strAry(11) = Lpad(GetAryValue(strInfo_Old, "身份证号"), 18)
    
    For lngLoop = LBound(strInfo) To UBound(strInfo)
        
        lngPos = InStr(strInfo(lngLoop), "=")
        If lngPos > 0 Then
            strItem = Trim(Mid(strInfo(lngLoop), 1, lngPos - 1))
            strValue = Trim(Mid(strInfo(lngLoop), lngPos + 1))
        End If
        
        Select Case strItem
        Case "门诊号"
        Case "住院号"
        Case "姓名"
            strAry(1) = Lpad(strValue, 8)
        Case "性别"
            strAry(2) = Lpad(strValue, 2)
        Case "年龄"
        Case "出生日期"
            strAry(4) = DString(strValue)
        Case "出生地点"
        Case "身份证号"
            strAry(11) = Lpad(strValue, 18)
        Case "身份"
        Case "职业"
        Case "民族"
        Case "国籍"
        Case "学历"
        Case "婚姻状况"
        Case "家庭地址"
        Case "家庭电话"
        Case "户口邮编"
        Case "联系人姓名"
        Case "联系人关系"
        Case "联系人地址"
        Case "联系人电话"
        Case "工作单位"
        Case "单位电话"
        Case "人员编码"
            strAry(10) = Lpad(strValue, 8)
        Case "人员性质"
            strAry(8) = Lpad(strValue, 1)
        Case "用工性质"
            strAry(9) = Lpad(strValue, 1)
        Case "工作岗位"
            strAry(7) = Lpad(strValue, 3)
        Case "单位编码"
            strAry(6) = Lpad(strValue, 3)
        Case "血型"
            strAry(3) = Lpad(strValue, 2)
        Case "工作日期"
            strAry(5) = DString(strValue)
        End Select
    Next
            
    '写卡操作
    intRet = swr_4442(mlngICDec, &H20, 12, strAry(1) & strAry(2) & strAry(3))
    If intRet < 0 Then
        strFunction = "swr_4442"
        GoTo errHand
    End If
        
    '压缩
    intRet = asc_hex(strAry(4) & strAry(5) & strAry(6) & strAry(7) & strAry(8) & strAry(9), bytResult1(0), 12)
    If intRet < 0 Then
        strFunction = "asc_hex"
        GoTo errHand
    End If
        
    intRet = swr_4442_hex(mlngICDec, &H2C, 12, bytResult1(0))
    If intRet < 0 Then
        strFunction = "swr_4442"
        GoTo errHand
    End If
    
    '写个人编码身份证号
    '加密
    intRet = asc_hex(strAry(10) & strAry(11), bytResult2(0), 13)
    If intRet < 0 Then
        strFunction = "asc_hex"
        GoTo errHand
    End If

    '写卡操作
    intRet = swr_4442_hex(mlngICDec, &H8, 13, bytResult2(0))
    If intRet < 0 Then
        strFunction = "swr_4442"
        GoTo errHand
    End If
    
    SetPatient = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg "调用函数[" & strFunction & "]时发生错误：" & vbCrLf & "错误号:" & Err.Number & "|错误信息:" & Err.Description
End Function

Public Function GetPatient(strInfo() As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：从IC卡读病人信息(姓名 性别 血型 出生年月 工作年月 工作单位 工作岗位 人员性质 用工性质 人员编码 身份证
    '           8+2+2+4+4+1.5+1.5+0.5+0.5 4+9)
    '参数：
    '返回：格式："门诊号=123344","住院号=34567"
    '------------------------------------------------------------------------------------------------------------------
    Dim intRet As Integer
    Dim strResult1 As String * 50
    Dim strResult2 As String * 50
    Dim strResult3 As String * 50
    Dim bytResult1(12) As Byte
    Dim bytResult2(13) As Byte
    Dim strFunction As String
    
    ReDim strInfo(1 To 11)
    
    On Error GoTo errHand
    
    '检查是否已初始化了卡，如没有则初始化
    If mlngICDec = 0 Then
        If InitCard = False Then GoTo EndPoint
    End If
    
    '检验卡（类型、是否插入）
    If HaveCard = False Then Exit Function
    If CardTypeOK = False Then Exit Function
    
    '1.读姓名、性别、血型没压缩的信息
    intRet = srd_4442(mlngICDec, &H20, 12, strResult1)
    If intRet < 0 Then
        strFunction = "srd_4442"
        GoTo errHand
    End If
    
    strInfo(1) = "姓名=" & GetSubStr(strResult1, 1, 8)
    strInfo(2) = "性别=" & GetSubStr(strResult1, 9, 2)
    strInfo(3) = "血型=" & GetSubStr(strResult1, 11, 2)
    
    '2.读取后面压缩信息：出生日期、工作日期、单位编码、工作岗位、人员性质、用工性质
    intRet = srd_4442_hex(mlngICDec, &H2C, 12, bytResult1(0))
    If intRet < 0 Then
        strFunction = "srd_4442_hex"
        GoTo errHand
    End If
    
    intRet = hex_asc(bytResult1(0), strResult2, 24)
    If intRet < 0 Then
        strFunction = "hex_asc"
        GoTo errHand
    End If
    
    strInfo(4) = "出生日期=" & SDate(GetSubStr(strResult2, 1, 8))
    strInfo(5) = "工作日期=" & SDate(GetSubStr(strResult2, 9, 8))
    strInfo(6) = "单位编码=" & GetSubStr(strResult2, 17, 3)
    strInfo(7) = "工作岗位=" & GetSubStr(strResult2, 20, 3)
    strInfo(8) = "人员性质=" & GetSubStr(strResult2, 23, 1)
    strInfo(9) = "用工性质=" & GetSubStr(strResult2, 24, 1)
    
    '3.读人员编码、身份证号
    intRet = srd_4442_hex(mlngICDec, &H8, 13, bytResult2(0))
    If intRet < 0 Then
        strFunction = "srd_4442_hex"
        GoTo errHand
    End If
    
    intRet = hex_asc(bytResult2(0), strResult3, 26)
    If intRet < 0 Then
        strFunction = "hex_asc"
        GoTo errHand
    End If
    
    strInfo(10) = "人员编码=" & GetSubStr(strResult3, 1, 8)
    strInfo(11) = "身份证号=" & GetSubStr(strResult3, 9, 18)
        
    GetPatient = True
    
    GoTo EndPoint
    
errHand:

    ShowSimpleMsg "调用函数[" & strFunction & "]时发生错误：" & vbCrLf & "错误号:" & Err.Number & "|错误信息:" & Err.Description
    
EndPoint:
    
End Function

'######################################################################################################################
'3.非接口函数、过程
Private Function CheckPsw() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：校验IC卡密码
    '参数：
    '返回：成功返回TRUE；否则返回FALSE
    '------------------------------------------------------------------------------------------------------------------
    Dim intRet As Integer
    Dim strPsw0(5) As Byte              '初始密码
    Dim strPsw1(5) As Byte              '本系统密码
    
    On Error GoTo errHand
    
    '为了不使本张卡被锁死，校验前先检查卡的计数器，若为1则本系统中不允许操作，但可以用IC卡自带的程序恢复
    If GetCounter = 1 Then
        Call ErrorCenter(101)
        Exit Function
    End If
    
    '校对初始密码（SLE4442的初始密码为FFFFFF）
    intRet = asc_hex("FFFFFF", strPsw0(0), 3)
    intRet = asc_hex("995188", strPsw1(0), 3)
    
    '先按本系统密码进行验证，若出错则计数器减1（总值为3，计数器为0时将被锁死）；校对一次计数器将恢复为3
    intRet = csc_4442(mlngICDec, 3, strPsw1(0))
    If intRet <> 0 Then
        
        If GetCounter = 1 Then
            Call ErrorCenter(101)
            Exit Function
        End If
        
        '出错，再按原始密码进行验证
        intRet = csc_4442(mlngICDec, 3, strPsw0(0))
        If intRet = 0 Then
            '原始密码验证通过，并写入本系统密码
            intRet = wsc_4442(mlngICDec, 3, strPsw1(0))
            If intRet < 0 Then
                Call ErrorCenter(102)
                Exit Function
            End If
        Else
            Call ErrorCenter(103)
            Exit Function
        End If
    End If
    
    CheckPsw = True
    
    Exit Function
errHand:
    ShowSimpleMsg "检验IC卡密码出错！"
End Function

Private Sub Class_Initialize()
    
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", key:=UCase("gstrSysName"), Default:="")
    
End Sub

Private Sub Class_Terminate()
    Call CloseCard
End Sub
