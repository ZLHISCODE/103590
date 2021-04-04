Attribute VB_Name = "mdlPublicOneCard"
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
Public glngInstanceCount As Long    '实例数

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
    Set gobjComLib = Nothing: Set gobjCommFun = Nothing: Set gobjDatabase = Nothing
    Set gobjControl = Nothing: Set gobjOneDataBase = Nothing: Set gobjLog = Nothing
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
    Err = 0: On Error GoTo errH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
errH:
    Err.Clear
    SubB = ""
End Function
