Attribute VB_Name = "mdlOneCardComLib"
Option Explicit
'--------------------------------------------------------------------------------------------------
'--系统
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long, glngSys As Long
Public gstrAviPath As String, gstrVersion As String
Public gstrMatchMethod As String
Public gstrProductName As String
Public gstrComputerName As String
Public gstrHelpPath As String
Public gstrDBUser As String   '当前数据库用户
Public gstrUnitName As String '用户单位名称
Public gcnOracle As ADODB.Connection
Public gstrNodeNo As String
Public gblnAutoGetOracleConnect As Boolean   '是否自动获取Oracle连接
Public glngInstanceCount As Long    '实例数

'-----------------------------------------------------------------------------------------------------
'所涉及对象集
Public grs医疗卡类别 As ADODB.Recordset
Public Type Ty_UserInfor
    id As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门名称 As String
End Type
Public UserInfo As Ty_UserInfor
'-----------------------------------------------------------------------------------------------------
'小数格式化串
Public Enum g小数类型
    g_数量 = 0
    g_成本价
    g_售价
    g_金额
    g_折扣率
End Enum
Private Type m_小数位
    数量小数 As Integer
    成本价小数 As Integer
    零售价小数 As Integer
    金额小数 As Integer
    折扣率 As Integer
End Type

Public g_小数位数 As m_小数位
Public Type g_FmtString
    FM_数量 As String
    FM_成本价 As String
    FM_零售价 As String
    FM_金额 As String
    FM_折扣率 As String
End Type
Public gVbFmtString As g_FmtString
Public gOraFmtString As g_FmtString
'-----------------------------------------------------------------------------------------------------
'颜色相关设置
Public Type Ty_Color
     lngGridColorSel As OLE_COLOR     '选择颜色
     lngGridColorLost As OLE_COLOR   '离开颜色
End Type
Public gSysColor As Ty_Color
'-----------------------------------------------------------------------------------------------------
'公共部件(zl9ComLib)
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object

Public gobjOneDataBase As clsDataBase      '一卡通数据联接对象
Public gobjOneDataObject As clsOneCardDataObject   '一卡通数据对象
'------------------------------------------------------------------------------------------------------------------------------------
'Api声明.
'电脑名称(ComputerName)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'------------------------------------------------------------------------------------------------------------------------------------
'调试
Private Type Ty_TestDebug
    blndebug As Boolean
    objSquareCard As clsCard
    bytType  As Byte  '1-随机产生卡号,2-读取卡号
    strStartNo As String    '开始卡号
    bln补调交易 As Boolean
End Type
Public gTy_TestBug As Ty_TestDebug
Public gbln自动读取 As Boolean '当前是否为射频卡



Public Sub 初始小数位数()
    '------------------------------------------------------------------------------------------------------
    '功能:初始小数位数
    '入参:
    '出参:
    '返回:7
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    With g_小数位数
        .成本价小数 = 7
        .零售价小数 = 7
        .金额小数 = 2
        .数量小数 = 3
        .折扣率 = 2
    End With
    With gVbFmtString
        .FM_成本价 = GetFmtString(g_成本价, False)
        .FM_金额 = GetFmtString(g_金额, False)
        .FM_零售价 = GetFmtString(g_售价, False)
        .FM_数量 = GetFmtString(g_数量, False)
        .FM_折扣率 = GetFmtString(g_折扣率, False)
    End With
    With gOraFmtString
        .FM_成本价 = GetFmtString(g_成本价, True)
        .FM_金额 = GetFmtString(g_金额, True)
        .FM_零售价 = GetFmtString(g_售价, True)
        .FM_数量 = GetFmtString(g_数量, True)
        .FM_折扣率 = GetFmtString(g_折扣率, True)
    End With
End Sub

Public Function GetFmtString(ByVal 小数类型 As g小数类型, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '功能:返回指定的小数格式串
    '入参: lng小数位数-小数位数
    '     blnOracle-返回是oracle的格式串还是Vb的格式串
    '出参:
    '返回:返回指定的格式串
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim int位数 As Integer
    Select Case 小数类型
    Case g_数量
         int位数 = g_小数位数.数量小数
    Case g_金额
         int位数 = g_小数位数.金额小数
    Case g_成本价
         int位数 = g_小数位数.成本价小数
    Case g_售价
         int位数 = g_小数位数.零售价小数
    Case Else
        int位数 = 0
    End Select
    If blnOracle Then
       GetFmtString = "'999999999990." & String(int位数, "9") & "'"
    Else
       GetFmtString = "#0." & String(int位数, "0") & ";-#0." & String(int位数, "0") & "; ;"
    End If
End Function

Public Function zlCheckTableIsExsit(ByVal strTableName As String, Optional cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查表是否存在
    '入参:strTableName-表名
    '返回:成存返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-04 10:48:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objDatabase As clsDataBase
    
    On Error GoTo errHandle
    If zlGetOneDataBase(cnOracle, objDatabase) = False Then Exit Function
    strSQL = "Select 1 From All_tables where table_name=[1]"
    Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "检查表是否存在", strTableName)
    zlCheckTableIsExsit = Not rsTemp.EOF
    Set objDatabase = Nothing
    Exit Function
errHandle:
    If objDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetOneDataBase(ByRef cnOracle As ADODB.Connection, ByRef objDataBase_Out As Object, Optional ByVal blnIsObjRegisterAlone As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一卡通连接对象
    '入参:cnOracle-数据库连接
    '出参:objDataBase_Out-返回数据操作对象(接口返回true时返回)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-03 13:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjOneDataBase Is Nothing Then Set objDataBase_Out = gobjOneDataBase: zlGetOneDataBase = True: Exit Function
    
    On Error GoTo errHandle
    Set gobjOneDataBase = New clsDataBase
    gobjOneDataBase.InitCommon cnOracle, blnIsObjRegisterAlone
    Set objDataBase_Out = gobjOneDataBase
    zlGetOneDataBase = True
    Exit Function
errHandle:
    Exit Function
End Function
Public Function zlGetOneCardDataObject(ByRef cnOracle As ADODB.Connection, ByRef objOneDataObject_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一卡通数据访问对象
    '入参:
    '出参:objOneDataObject_Out-返回一卡通数据访问对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-04 14:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
   On Error GoTo errHandle
    If Not gobjOneDataObject Is Nothing Then Set objOneDataObject_Out = gobjOneDataObject: zlGetOneCardDataObject = True: Exit Function
    
    Set gobjOneDataObject = New clsOneCardDataObject
    gobjOneDataObject.InitCommon cnOracle
    Set objOneDataObject_Out = gobjOneDataObject
    zlGetOneCardDataObject = True
    Exit Function
errHandle:
    Exit Function
End Function
 
Public Sub zlInitPublicVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化公共变量
    '编制:刘兴洪
    '日期:2018-12-03 13:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrAviPath = GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrHelpPath = gstrAviPath & "\help"
    gstrComputerName = zlGetComputerName
    With gSysColor
        .lngGridColorLost = &HE0E0E0   '离开颜色
        .lngGridColorSel = &HFFEBD7       '选择颜色
    End With
    Call 初始小数位数
    
    '取站点
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing And gstrNodeNo = "" Then
        gstrNodeNo = gobjComLib.gstrNodeNo
    End If
    
End Sub
Public Function zlGetComputerName() As String
    '------------------------------------------------------------------------------------------------------------------
    '功能：获取电脑名称
    '参数：
    '说明：
    '------------------------------------------------------------------------------------------------------------------
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    strComputer = strComputer
    zlGetComputerName = Trim(Replace(strComputer, Chr(0), ""))
End Function

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub
Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    'clsCommFun存在该函数
    '功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Function GetCardNODencodeRule(ByVal lng卡类别ID As Long, _
    Optional bln消费卡 As Boolean = False, Optional cnOracle As ADODB.Connection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡类别ID的规则
    '入参:lng卡类别ID-卡类别ID
    '        bln消费卡-是否消费卡
    '返回:卡类别的卡号编码规则
    '编制:刘兴洪
    '日期:2011-06-22 11:01:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    If bln消费卡 Then
        Set rsTemp = zlGet消费卡接口()
        rsTemp.Filter = "ID=" & lng卡类别ID
        If rsTemp.EOF Then GoTo GoEnd:
        GetCardNODencodeRule = NVL(rsTemp!是否密文)
        GoTo GoEnd:
    End If
    Set rsTemp = zlGet医疗卡类别()
    rsTemp.Filter = "ID=" & lng卡类别ID
    If rsTemp.EOF Then GoTo GoEnd:
    GetCardNODencodeRule = NVL(rsTemp!卡号密文)
GoEnd:
    rsTemp.Filter = 0
End Function

Public Function GetCardNODencode(ByVal strCardNo As String, _
    Optional lng卡类别ID As Long = 0, _
    Optional strRule As String = "", Optional bln消费卡 As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡编码
    '入参:lng卡类别ID-卡类别ID或消费卡序号,如果传入,将以医疗卡类别或消费卡别中的"卡号密文"或是否密文进行加密
    '       strRule-规则:2-4表示从2位到4位用*代替,如无-号,则表示从最后几位显示为*
    '       strCardNo-卡号
    '出参:
    '返回:带**的卡号,如果错误,返回空
    '编制:刘兴洪
    '日期:2011-06-21 14:21:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPass As Variant
    Dim strCardPassText As String, i As Long, J As Long
    If bln消费卡 Then
        If Val(strRule) = 1 Then GetCardNODencode = String(Len(strCardNo), "*"): Exit Function
        If lng卡类别ID = 0 Then GetCardNODencode = strCardNo: Exit Function
        If Val(GetCardNODencodeRule(lng卡类别ID, True)) = 1 Then
            GetCardNODencode = String(Len(strCardNo), "*"): Exit Function
        Else
            GetCardNODencode = strCardNo: Exit Function
        End If
    End If
    If lng卡类别ID <> 0 And strRule = "" Then
        strCardPassText = GetCardNODencodeRule(lng卡类别ID)
    Else
        '取号规则
        strCardPassText = strRule
    End If
    If strCardPassText = "" Then
       GetCardNODencode = strCardNo
    End If
    varPass = Split(strCardPassText & "-", "-")
    If Val(varPass(0)) = 0 Or Val(varPass(1)) = 0 Then
        '最后几位显示*
        i = IIf(Val(varPass(0)) = 0, Val(varPass(1)), Val(varPass(0)))
        If i = 0 Then GetCardNODencode = strCardNo: Exit Function
        J = Len(strCardNo) - i: J = IIf(J < 0, 0, J)
        GetCardNODencode = Mid(strCardNo, 1, J) & String(i, "*")
        Exit Function
    End If
    i = Val(varPass(0)): J = Val(varPass(1))
    If i > Len(strCardNo) Then GetCardNODencode = strCardNo: Exit Function
    If J > Len(strCardNo) Then J = Len(strCardNo)
    If J < i Then J = i
   GetCardNODencode = Mid(strCardNo, 1, i - 1) & String(J - i + 1, "*") & Mid(strCardNo, J + 1)
End Function
Public Function GetAvailabilityWriteCardType() As String
       '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:存在写卡类别
        '出参:返回写卡类别,多个用逗号分离
        '返回:存在写卡类别的ID,如:123,232,...
        '编制:刘兴洪
        '日期:2013-06-07 10:40:59
        '说明:
        '---------------------------------------------------------------------------------------------------------------------------------------------
        Dim rsTemp As ADODB.Recordset, strWriteCardIDs As String
        Dim intAutoRead As Integer, intAutoSplitTime As Integer, blnStartCardType As Boolean, str部件 As String
        On Error GoTo errHandle
        
        Set rsTemp = zlGet医疗卡类别
        rsTemp.Filter = "是否写卡=1 And 是否启用=1"
        If rsTemp.EOF Then GetAvailabilityWriteCardType = "": Exit Function
        strWriteCardIDs = ""
         With rsTemp
            '自制卡(即消费卡)
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
                intAutoRead = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & NVL(!编码), "自动读取", "0"))
                intAutoSplitTime = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & NVL(!编码), "自动读取间隔", "300"))
                If Val(NVL(rsTemp!是否自制)) = 1 Then   '自制卡,都启用
                    blnStartCardType = True
                Else
                    blnStartCardType = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\医疗卡\" & NVL(!编码), "启用", "0")) = 1
                End If
                If blnStartCardType Then
                    strWriteCardIDs = strWriteCardIDs & "," & Val(NVL(rsTemp!id))
                End If
                .MoveNext
            Loop
         End With
         Set rsTemp = Nothing
        If strWriteCardIDs <> "" Then strWriteCardIDs = Mid(strWriteCardIDs, 2)
        GetAvailabilityWriteCardType = strWriteCardIDs
        Exit Function
errHandle:
        If gobjComLib.ErrCenter() = 1 Then
            Resume
        End If
End Function


Public Function GetCardFromCardtypeID(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As Card) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别给卡对象(此对象
    '入参:lngCardTypeID-卡类别ID
    '       bln消费卡-是否消费卡
    '出参:objCard-返回卡对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-25 10:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, str部件 As String
    Dim int自动读取 As Integer, int自动读取间隔 As Integer, bln启用 As Boolean
    Dim objDatabase As New clsDataBase, str读卡性质 As String
    
    On Error GoTo errHandle
    Set objCard = New Card
    If Not bln消费卡 Then
        Set rsTemp = zlGet医疗卡类别
        rsTemp.Filter = "id=" & lngCardTypeID
        If rsTemp.EOF Then rsTemp.Filter = 0: Exit Function
        If Val(NVL(rsTemp!是否启用)) = 1 Then
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & NVL(rsTemp!编码), "自动读取", "0"))
            int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & NVL(rsTemp!编码), "自动读取间隔", "300"))
            If Val(NVL(rsTemp!是否自制)) = 1 Then   '自制卡,都启用
                bln启用 = True
            Else
                '问题号:54098
                If (NVL(rsTemp!名称) Like "*身份证*" Or NVL(rsTemp!名称) Like "*IC卡*") And Val(NVL(rsTemp!是否固定)) = 1 And NVL(rsTemp!部件) = "" Then
                    bln启用 = True
                Else
                    bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\医疗卡\" & NVL(rsTemp!编码), "启用", "0")) = 1
                End If
            End If
        Else
            bln启用 = False
        End If
        str部件 = Trim(NVL(rsTemp!部件))
        'ID,编码,名称,短名,前缀文本,卡号长度,缺省标志,是否固定,是否严格控制,是否刷卡,是否自制,是否存在帐户,是否全退,部件,备注,特定项目,结算方式,是否启用
        Set objCard = New Card
        With objCard
            .接口序号 = NVL(rsTemp!id)
            .接口编码 = NVL(rsTemp!编码)
            .短名 = NVL(rsTemp!短名)
            .名称 = NVL(rsTemp!名称)
            .前缀文本 = NVL(rsTemp!前缀文本)
            .卡号长度 = Val(NVL(rsTemp!卡号长度)) + Val(NVL(rsTemp!设备是否启用回车))
            .缺省标志 = Val(NVL(rsTemp!缺省标志)) = 1
            .系统 = Val(NVL(rsTemp!是否固定)) = 1
            .是否严格控制 = Val(NVL(rsTemp!是否严格控制)) = 1
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动读取间隔
            .自制卡 = Val(NVL(rsTemp!是否自制)) = 1
            .是否存在帐户 = Val(NVL(rsTemp!是否存在帐户)) = 1
            .是否全退 = Val(NVL(rsTemp!是否全退)) = 1
            .卡号重复使用 = Val(NVL(rsTemp!是否重复使用)) = 1
            .结算方式 = NVL(rsTemp!结算方式)
            .接口程序名 = NVL(rsTemp!部件)
            .特定项目 = NVL(rsTemp!特定项目)
            .启用 = bln启用
            .备注 = NVL(rsTemp!备注)
            .卡号密文规则 = NVL(rsTemp!卡号密文)
            .是否退现 = Val(NVL(rsTemp!是否退现)) = 1
            .密码长度 = Val(NVL(rsTemp!密码长度))
            .密码长度限制 = Val(NVL(rsTemp!密码长度限制))
            .密码规则 = Val(NVL(rsTemp!密码规则))
            .密码输入限制 = Val(NVL(rsTemp!密码输入限制))
            .是否缺省密码 = Val(NVL(rsTemp!是否缺省密码)) = 1
            .是否制卡 = Val(NVL(rsTemp!是否制卡)) = 1   '56615
            .是否发卡 = Val(NVL(rsTemp!是否发卡)) = 1 Or .自制卡
            .是否写卡 = Val(NVL(rsTemp!是否写卡)) = 1
            .结算性质 = Val(NVL(rsTemp!结算性质))
            .是否转帐及代扣 = Val(NVL(rsTemp!是否转帐及代扣)) = 1
            str读卡性质 = NVL(rsTemp!读卡性质, "1000")
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            .是否持卡消费 = Val(NVL(rsTemp!是否持卡消费)) = 1
            .是否退款验卡 = Val(NVL(rsTemp!是否退款验卡)) = 1
            .是否证件 = Val(NVL(rsTemp!是否证件)) = 1
            .设备是否启用回车 = Val(NVL(rsTemp!设备是否启用回车)) = 1
            .是否缺省退现 = Val(NVL(rsTemp!是否缺省退现)) = 1
            .是否独立结算 = Val(NVL(rsTemp!是否独立结算)) = 1
            .是否支持扫码付 = Val(NVL(rsTemp!是否支持扫码付)) = 1
        End With
        rsTemp.Filter = 0:
       GetCardFromCardtypeID = True
        Exit Function
    End If
    
    
    Set rsTemp = zlGet消费卡接口
    rsTemp.Filter = "ID=" & lngCardTypeID
    If rsTemp.EOF Then Set rsTemp.Filter = 0: Exit Function
    
    With rsTemp
        '自制卡(即消费卡)
        If .RecordCount <> 0 Then .MoveFirst
        bln启用 = Val(NVL(!启用)) = 1
        If bln启用 Then
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & NVL(!编号), "自动读取", "0"))
            int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & NVL(!编号), "自动读取间隔", "300"))
            bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & NVL(!编号), "启用", "0")) = 1
        End If
        '编号,名称,结算方式,nvl(自制卡,0)  as 自制卡,前缀文本,卡号长度,部件,系统,是否密文
        str部件 = Trim(NVL(rsTemp!部件))
        Set objCard = New Card
        With objCard
            .接口序号 = NVL(rsTemp!编号)
            .接口编码 = NVL(rsTemp!编号)
            .短名 = Left(NVL(rsTemp!名称), 1)   '默认取第一个
            .名称 = NVL(rsTemp!名称)
            .前缀文本 = NVL(rsTemp!前缀文本)
            .卡号长度 = Val(NVL(rsTemp!卡号长度))
            .系统 = Val(NVL(rsTemp!系统)) = 1
            .是否严格控制 = False
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动读取间隔
            .自制卡 = Val(NVL(rsTemp!自制卡)) = 1
            .是否存在帐户 = True 'Not (Val(Nvl(rsTemp!自制卡)) = 1)
            .是否全退 = Val(NVL(rsTemp!是否全退)) = 1
            .结算方式 = NVL(rsTemp!结算方式)
            .接口程序名 = NVL(rsTemp!部件)
            .特定项目 = ""
            .启用 = bln启用
            .卡号重复使用 = True
            .备注 = ""
            .卡号密文规则 = NVL(rsTemp!是否密文)
            .消费卡 = True
            .是否退现 = Val(NVL(rsTemp!是否退现)) = 1
            .密码长度 = Val(NVL(rsTemp!密码长度))
            .密码长度限制 = Val(NVL(rsTemp!密码长度限制))
            .密码规则 = Val(NVL(rsTemp!密码规则))
            .密码输入限制 = Val(NVL(rsTemp!密码输入限制))
            .是否缺省密码 = Val(NVL(rsTemp!是否缺省密码)) = 1
            .是否制卡 = Val(NVL(rsTemp!是否制卡)) = 1   '56615
            .是否发卡 = Val(NVL(rsTemp!是否发卡)) = 1 Or .自制卡
            .是否写卡 = Val(NVL(rsTemp!是否写卡)) = 1
            .结算性质 = Val(NVL(rsTemp!结算性质))
            str读卡性质 = NVL(rsTemp!读卡性质, "1000")
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
        End With
    End With
    Set rsTemp.Filter = 0:
    GetCardFromCardtypeID = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = Forms.count = 0
End Function

Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:释放资源
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '实例数为0时，才放资源
    If glngInstanceCount > 0 Then Exit Function
    Call zlCloseWindows '释放窗体资源
    Set gobjComLib = Nothing: Set gobjCommFun = Nothing
    Set gobjDatabase = Nothing: Set gobjControl = Nothing
    Set gobjLog = Nothing: Set gobjOneDataBase = Nothing
    Set grs消费卡接口 = Nothing: Set grs医疗卡类别 = Nothing
    Set gcnOracle = Nothing
    Set gobjOneDataObject = Nothing
    zlReleaseResources = True
End Function

Public Sub zlInitCommLib()
   '初始化公共部件
    If Not gobjComLib Is Nothing Then Exit Sub

    Err = 0: On Error Resume Next
    Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    Err = 0: On Error GoTo 0
 End Sub
 
 Public Function zlStringEncode(ByVal strPutString As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:字符串加密
    '入参:strPutString-需要加密的串
    '出参:
    '返回:加密串
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strPutString = "" Then Exit Function
    zlStringEncode = Md5_String_Calc(strPutString)
End Function
Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function


Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'功能:读取指定字串的值,字串中可以包含汉字
 '入参:strInfor-原串
 '         lngStart-直始位置
'         lngLen-长度
'返回:子串
    Err = 0: On Error GoTo ErrH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
ErrH:
    Err.Clear
    SubB = ""
End Function
Public Function zlGetCardTypeRecStru(ByRef rsCardType As ADODB.Recordset) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡类别结构
     '出参:rsCardType-返回的记录集结构
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-24 18:10:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsCardType = New ADODB.Recordset
    With rsCardType
        If .State = 1 Then .Close
        'adBigInt
        .fields.Append "ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "编码", adLongVarChar, 50, adFldIsNullable
        .fields.Append "名称", adLongVarChar, 200, adFldIsNullable
        .fields.Append "短名", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "前缀文本", adLongVarChar, 30, adFldIsNullable
        .fields.Append "卡号长度", adSmallInt, 20, adFldIsNullable
        .fields.Append "缺省标志", adSmallInt, , adFldIsNullable
        .fields.Append "是否固定", adSmallInt, , adFldIsNullable
        .fields.Append "是否严格控制", adSmallInt, , adFldIsNullable
        .fields.Append "是否自制", adSmallInt, , adFldIsNullable
        .fields.Append "是否存在帐户", adSmallInt, , adFldIsNullable
        
        .fields.Append "是否退现", adSmallInt, , adFldIsNullable
        .fields.Append "是否缺省退现", adSmallInt, , adFldIsNullable
        .fields.Append "是否全退", adSmallInt, , adFldIsNullable
        .fields.Append "是否重复使用", adSmallInt, , adFldIsNullable
        .fields.Append "发卡性质", adSmallInt, , adFldIsNullable
        .fields.Append "密码长度", adSmallInt, , adFldIsNullable
        .fields.Append "密码长度限制", adSmallInt, , adFldIsNullable
        .fields.Append "密码规则", adSmallInt, , adFldIsNullable
        .fields.Append "部件", adLongVarChar, 100, adFldIsNullable
        .fields.Append "备注", adLongVarChar, 300, adFldIsNullable
        .fields.Append "特定项目", adLongVarChar, 100, adFldIsNullable
        .fields.Append "结算方式", adLongVarChar, 50, adFldIsNullable
        .fields.Append "是否启用", adSmallInt, , adFldIsNullable
        .fields.Append "卡号密文", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "密码输入限制", adSmallInt, , adFldIsNullable
        .fields.Append "是否缺省密码", adSmallInt, , adFldIsNullable
        .fields.Append "是否模糊查找", adSmallInt, , adFldIsNullable
        .fields.Append "是否制卡", adSmallInt, , adFldIsNullable
        .fields.Append "是否写卡", adSmallInt, , adFldIsNullable
        .fields.Append "是否发卡", adSmallInt, , adFldIsNullable
        .fields.Append "发卡控制", adSmallInt, , adFldIsNullable
        
        
        .fields.Append "结算性质", adSmallInt, , adFldIsNullable
        .fields.Append "是否证件", adSmallInt, , adFldIsNullable
        
        
        .fields.Append "是否转帐及代扣", adSmallInt, , adFldIsNullable
        .fields.Append "读卡性质", adLongVarChar, 20, adFldIsNullable
        .fields.Append "是否持卡消费", adSmallInt, , adFldIsNullable
        .fields.Append "发送调用接口", adSmallInt, , adFldIsNullable
        .fields.Append "是否退款验卡", adSmallInt, , adFldIsNullable
        .fields.Append "是否独立结算", adSmallInt, , adFldIsNullable
        .fields.Append "缺省有效时间", adLongVarChar, 50, adFldIsNullable
        .fields.Append "卡号识别规则", adSmallInt, , adFldIsNullable
        .fields.Append "是否支持扫码付", adSmallInt, , adFldIsNullable
        .fields.Append "险类", adSmallInt, , adFldIsNullable
        .fields.Append "键盘控制方式", adSmallInt, , adFldIsNullable
        .fields.Append "是否启用回车", adSmallInt, , adFldIsNullable
    
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    zlGetCardTypeRecStru = True
End Function

Public Function zlGetQueryPatiInforStru(ByRef rsPati As ADODB.Recordset, _
    Optional ByVal bytPatiInfoShowType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取查询的病人信息数据集
    '入参：
    '       bytPatiInfoShowType-病区选择器列显示方式：0-所有信息，1-简略信息
    '出参:rsPati-返回的记录集结构
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsPati = New ADODB.Recordset
    With rsPati
        .fields.Append "ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "在院", adLongVarChar, 4, adFldIsNullable
        .fields.Append "病人ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "名称", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .fields.Append "身份证号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "IC卡号", adLongVarChar, 50, adFldIsNullable
        .fields.Append "门诊号", adLongVarChar, 18, adFldIsNullable
        .fields.Append "住院号", adLongVarChar, 18, adFldIsNullable
        .fields.Append "手机号", adLongVarChar, 40, adFldIsNullable
        
        If bytPatiInfoShowType = 0 Then
            .fields.Append "出生日期", adLongVarChar, 30, adFldIsNullable
            .fields.Append "出生地点", adLongVarChar, 200, adFldIsNullable
            .fields.Append "费别", adLongVarChar, 50, adFldIsNullable
            .fields.Append "医疗付款方式", adLongVarChar, 100, adFldIsNullable
            .fields.Append "民族", adLongVarChar, 30, adFldIsNullable
            .fields.Append "家庭地址", adLongVarChar, 200, adFldIsNullable
            .fields.Append "家庭电话", adLongVarChar, 50, adFldIsNullable
            .fields.Append "联系人姓名", adLongVarChar, 100, adFldIsNullable
            .fields.Append "联系人关系", adLongVarChar, 50, adFldIsNullable
            .fields.Append "联系人电话", adLongVarChar, 100, adFldIsNullable
            .fields.Append "门诊预交余额", adDouble, , adFldIsNullable
            .fields.Append "住院预交余额", adDouble, , adFldIsNullable
            .fields.Append "卡号", adLongVarChar, 200, adFldIsNullable
            .fields.Append "密码ID", adLongVarChar, 100, adFldIsNullable
         End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    zlGetQueryPatiInforStru = True
End Function

Public Function GetOneCardTypes(ByRef rsTypes_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新获取医疗卡类别数据集
    '功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-21 15:08:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllData As Collection, cllTemp As Variant
 
    On Error GoTo errHandle
    

    If zlGetCardTypeRecStru(rsTypes_Out) = False Then Exit Function
    If zl_PatiSvr_GetCardTypes(cllData) = False Then Exit Function
    
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    'output
    '    cardtype_id N   1   ID
    '    cardtype_code   C   1   编码
    '    cardtype_name   C   1   名称
    '    cardtype_stname C   1   短名
    '    prefix_text C   1   前缀文本
    '    cardno_len  N   1   卡号长度
    '    default    N   1   缺省标志
    '    fixed N   1   是否固定:1-是系统固定;0-不是系统固定
    '    strict   N   1   是否严格控制:1-是严格控制;0-不是严格控制
    '    self_make N   1   是否自制:1-是的;0-不是
    '    exist_account  N   1   是否存在帐户:1-存在帐户;0-不存在账户
    '    allow_return_cash    N   1   是否退现:1-允许;0-不允许
    '    must_all_return   N   1   是否全退:1-必需全退;0-允许部分退
    '    component   C   1   部件
    '    memo    C   1   备注
    '    spec_item   C   1   特定项目
    '    blnc_mode   C   1   结算方式
    '    blnc_nature N   1   结算性质
    '    cardno_pwdtxt   C   1   卡号密文:卡号从第几位至第几位显示密文,格式为:S-N:S表示从第几位开始,至第几位结束.比如:3-10,表示从3位到10位用密文*表示:12********3323主要是适应不同类别的医疗卡
    '    allow_repeat_use N   1   是否重复使用:1-允许;0-不允许
    '    enabled    N   1   是否启用:1-已启用;0-未启用
    '    pwd_len N   1   密码长度
    '    pwd_len_limit   N   1   密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
    '    pwd_rule    N   1   密码规则:０-数字和字符组成;1-仅为数字组成
    '    allow_vaguefind    N   1   是否模糊查找:1-支持模糊查找;0-不支持
    '    pwd_require    N   1   密码输入限制:0-不限制;1-不输入,提醒;2-不输入禁止;缺省为不限制
    '    default_pwd  N   1   是否缺省密码:1-以身份证后N(以密码长度为准)位作为缺省密码;0-无缺省密码
    '    allow_makecard N   1   是否制卡:1-是;0-否
    '    allow_sendcard N   1   是否发卡:1-是;0-否
    '    allow_writecard    N   1   是否写卡:1-是;0-否
    '    insurance_type  N   1   险类
    '    sendcard_nature N   1   发卡性质:0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0
    '    allow_transfer N   1   是否转帐及代扣:1-支持转帐及代扣;0-不支持
    '    readcard_nature C   1   读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
    '    keyboard_mode   N   1   键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
    '    advsend_buildqrcode N   1   是否医嘱发送调用条码生成接口:1-发送调用生成二维码接口;0-不调用
    '    holding_pay   N   1   是否持卡消费:1-是;0-否
    '    cert_cardtype    N   1   是否证件类型的医疗卡:0-不是；1-是
    '    verfycard    N   1   是否退款验卡
    '    sendcard_sign   N   1   发卡控制:0或NULL-发卡时，卡号必须达到卡号长度;1-发卡时，允许卡号小于等于卡号长度,发卡时，小于卡号长度时，不提示操作员;2-发卡时，允许卡号小于等于卡号长度,小于时，提示操作员。
    '    enterkey_enabled N   1   设备是否启用回车:医疗卡对应的刷卡设备是否启用了回车，如果启用了回车，则卡号长度默认增加一位来屏蔽回车
    '    def_return_cash N   1   是否缺省退现:允许退现时,默认是否退现
    '    balalone N   1   是否独立结算:1-独立结算;0-非独立结算
    '    discern_rule    N   1   卡号识别规则:1-全部转换为大写;0-不区分大小写
    '    def_valid_time  C   1   缺省有效时间:NULL时，表示不限制;非空时，格式为:时间+单位(天，月),比如：3天,3月
    '    scanpay  N   1   是否支持扫码付:是否支持扫码付,支持时，会调用“zlReadQRCode部件”

    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        With rsTypes_Out
            .AddNew
                !id = cllTemp("_cardtype_id")
                !编码 = cllTemp("_cardtype_code")
                !名称 = cllTemp("_cardtype_name")
                !短名 = cllTemp("_cardtype_stname")
                
                !前缀文本 = cllTemp("_prefix_text")
                !卡号长度 = cllTemp("_cardno_len")
                !缺省标志 = cllTemp("_default")
                !是否固定 = cllTemp("_fixed")
                
                !是否严格控制 = cllTemp("_strict")
                !是否自制 = cllTemp("_self_make")
                !是否存在帐户 = cllTemp("_exist_account")
                !是否退现 = cllTemp("_allow_return_cash")
                !是否缺省退现 = cllTemp("_def_return_cash")
                !是否全退 = cllTemp("_must_all_return")
                !部件 = cllTemp("_component")
                !备注 = cllTemp("_memo")
                
                !特定项目 = cllTemp("_spec_item")
                !结算方式 = cllTemp("_blnc_mode")
                !卡号密文 = cllTemp("_cardno_pwdtxt")
                  
                !是否重复使用 = cllTemp("_allow_repeat_use")
                !是否启用 = cllTemp("_enabled")
                
                !密码长度 = cllTemp("_pwd_len")
                !密码长度限制 = cllTemp("_pwd_len_limit")
                !密码规则 = cllTemp("_pwd_rule")
                !密码输入限制 = cllTemp("_pwd_require")
                
                !是否模糊查找 = cllTemp("_allow_vaguefind")
                !是否缺省密码 = cllTemp("_default_pwd")
                
                
                
                !是否制卡 = cllTemp("_allow_makecard")
                !是否发卡 = cllTemp("_allow_sendcard")
                !是否写卡 = cllTemp("_allow_writecard")
                !发卡控制 = cllTemp("_sendcard_sign")
                
                !结算性质 = cllTemp("_blnc_nature")
                !险类 = cllTemp("_insurance_type")
                !发卡性质 = cllTemp("_sendcard_nature")
                !是否转帐及代扣 = cllTemp("_allow_transfer")
                !读卡性质 = cllTemp("_readcard_nature")
                !键盘控制方式 = cllTemp("_keyboard_mode")
                
                !是否持卡消费 = cllTemp("_holding_pay")
                !是否证件 = cllTemp("_cert_cardtype")
                
                !发送调用接口 = cllTemp("_advsend_buildqrcode")
                !是否退款验卡 = cllTemp("_verfycard")
                                
                !是否独立结算 = cllTemp("_balalone")
                !缺省有效时间 = cllTemp("_def_valid_time")
                !卡号识别规则 = cllTemp("_discern_rule")
                !是否支持扫码付 = cllTemp("_scanpay")
                !是否启用回车 = cllTemp("_enterkey_enabled")
                
            .Update
       End With
    Next
    If rsTypes_Out.RecordCount <> 0 Then rsTypes_Out.MoveFirst
    GetOneCardTypes = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet医疗卡类别() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别
    '返回:返回医疗卡类别的记录集
    '编制:刘兴洪
    '日期:2011-05-23 17:25:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllData As Collection, cllTemp As Variant
    
    On Error GoTo errHandle
    
    If Not grs医疗卡类别 Is Nothing Then
        If grs医疗卡类别.State = 1 Then
            grs医疗卡类别.Filter = 0
            If grs医疗卡类别.RecordCount <> 0 Then grs医疗卡类别.MoveFirst
            Set zlGet医疗卡类别 = grs医疗卡类别
            Exit Function
        End If
    End If
    If GetOneCardTypes(grs医疗卡类别) = False Then Set grs医疗卡类别 = Nothing: Exit Function
    Set zlGet医疗卡类别 = grs医疗卡类别
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Set grs医疗卡类别 = Nothing
End Function
Public Function GetPatiSurplusFromPatiID(ByVal lng病人ID As Long, ByRef dbl门诊预交余额_out As Double, ByRef dbl住院预交余额_Out As Double, _
    ByRef dbl门诊费用余额_Out As Double, ByRef dbl住院费用余额_Out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID获取病人费用及预交余额
    '入参:lng病人ID-病人ID
    '
    '出参:dbl门诊预交余额_out
    '     dbl住院预交余额_Out
    '     dbl门诊费用余额_Out
    '     dbl住院费用余额_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-05 20:57:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllData As Collection, cllTemp As Collection
    
    dbl门诊预交余额_out = 0
    dbl住院预交余额_Out = 0
    
    dbl门诊费用余额_Out = 0
    dbl住院费用余额_Out = 0
    If zl_ExseSvr_GetPatiSurplusInfo(lng病人ID, cllData) = False Then Exit Function
    Set cllTemp = zlGetNodeObjectFromCollect(cllData, "_" & lng病人ID)
    If cllTemp Is Nothing Then GetPatiSurplusFromPatiID = True: Exit Function
    dbl门诊预交余额_out = Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_outdpst_surplus", "C")))
    dbl住院预交余额_Out = Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_indpst_surplus", "C")))
    
    dbl门诊费用余额_Out = Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_outfee_surplus", "C")))
    dbl住院费用余额_Out = Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_infee_surplus", "C")))
    GetPatiSurplusFromPatiID = True
End Function
