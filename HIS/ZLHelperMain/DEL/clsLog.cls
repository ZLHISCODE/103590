VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/5/23
'模块           clsLog
'说明           日志模块
'==================================================================================================
'===========================================================================
'                0、API和常量生命
'===========================================================================
Private Declare Function GetTickCount Lib "kernel32" () As Long
'功能：检索自系统启动以来已运行的毫秒数，最多可达49.7天。
'返回：返回值是自系统启动以来运行的毫秒数。
'注意事项：GetTickCount函数的解析仅限于系统计时器的精度，通常在10毫秒到16毫秒之间。
'        GetTickCount函数的解析不会受到getsystemtime调适函数的调整的影响?
'        经过的时间存储为DWORD的值?
'        因此，如果系统连续运行49.7天，那么时间将会是零。
'        为了避免这个问题，请使用GetTickCount64函数。
'        否则，在比较时检查溢出条件。
'        如果你需要一个更高的分辨率计时器，可以使用多媒体定时器或高分辨率计时器。
'        为了获得计算机启动后的时间，在注册表关键hkeyperformance cedata的性能数据中检索系统的时间计数器。
'        返回的值是一个8字节的值?
'        要了解更多信息，请参见性能计数器。
'        Note：获得时间系统在工作状态自启动 , 使用QueryUnbiasedInterruptTime函数?
'        调试注意QueryUnbiasedInterruptTime函数产生不同的结果(“checked”)构建Windows,因为中断时间计数是由大约49天。
'        这有助于识别在系统运行很长时间之前可能不会发生的错误?
'        通过Microsoft Developer Network(MSDN)Web站点可以对MSDN的用户进行检查。
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'@原型
'    DWORD GetCurrentProcessId(
'
'    );
'@功能
'    检索调用进程的进程标识符。
'@参数
'    这个函数没有参数?
'@返回值
'    返回值是调用进程的进程标识符?
'@备注
'    在进程终止之前，进程标识符在整个系统中唯一地标识进程。
'@Requirements
'Minimum supported client       Windows XP [desktop apps | UWP apps]
'Minimum supported server       Windows Server 2003 [desktop apps | UWP apps]
'Header                         processthreadsapi.h (include Windows Server 2003, Windows Vista, Windows 7, Windows Server 2008 Windows Server 2008 R2, Windows.h)
'Library                        Kernel32.lib
'dll                            Kernel32.dll
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long

Private Const M_OFFSET_4        As Double = 4294967296#         '无符号整形的最大值
Private Const M_DAY_MS          As Double = 86400000            '一天的毫秒数
Private Const M_HOUR_MS         As Double = 3600000             '一小时的毫秒数
Private Const M_MIN_MS          As Double = 60000               '一分钟的毫秒数
Public Enum RunLogLevel
    RLL_NoneLog = 0             '不记录日志
    RLL_RunError = 4            '只记录错误
    RLL_LogInfo = 8             '只记录错误和LogInfo日志
    RLL_AllLog = 16             '记录所有错误
    RLL_ALLDetail = 32          '记录所有的细节
End Enum
'===========================================================================
'                1、常规变量
'===========================================================================
Private mobjLog                 As TextStream
Private mobjFSO                 As New FileSystemObject
Private mlngStart               As Long                         '日志生成时间
Private mlngCur                 As Long                         '当前时间
Private mblnNormalTime          As Boolean                      '常规的时间串
'日志跟踪
Private Const mlngStackLen      As Integer = 40                     '调用堆栈的长度
Private mcllMethodStack         As New Collection                   '调用堆栈集合
Private mstrText                As String
Private mlngIndex               As Long
'临时变量
Private mstrLogTiltle           As String
Private mstrLogTilEx            As String
Private mstrLogTime             As String
Private mstrLogDistance         As String
Private mdblTmp                 As Double
Private mlngTmp                 As Long
'===========================================================================
'                2、属性变量与定义
'===========================================================================
'IsAppend是否添加方式(R)
Private mblnAppend              As Boolean
'CurrentLogLevel当前日志级别(R/W)
Private mrllCurrentLogLevel     As RunLogLevel
'LogFile日志文件文件路径(R)
Private mstrLogFile             As String
'MaxLogCount日志文件文件最大行数(R/W)
Private mlngMaxLogCount         As Long
'CurrentLogCount当前已经记录的行数(R)
Private mlngCurLogCount         As Long

'IsAppend日志是否是追加方式R/W)
Public Property Get IsAppend() As Boolean
    IsAppend = mblnAppend
End Property
'CurrentLogLevel当前日志级别(R/W)
Public Property Get CurrentLogLevel() As RunLogLevel
    CurrentLogLevel = mrllCurrentLogLevel
End Property

Public Property Let CurrentLogLevel(ByVal rllNewLogLevel As RunLogLevel)
    If mrllCurrentLogLevel = RLL_NoneLog And rllNewLogLevel > RLL_NoneLog And mobjLog Is Nothing Then
        Call LogOpen(mstrLogFile, mblnAppend, mlngMaxLogCount, rllNewLogLevel, mblnNormalTime)
    End If
    mrllCurrentLogLevel = rllNewLogLevel
End Property
'IsNormalTime当前日志时间是否是自然时间(R/W)
Public Property Get IsNormalTime() As Boolean
    IsNormalTime = mblnNormalTime
End Property

Public Property Let IsNormalTime(ByVal blnNormalTime As Boolean)
    mblnNormalTime = blnNormalTime
End Property
'LogFile日志文件文件名称(R)
Public Property Get LogFile() As String
    LogFile = mstrLogFile
End Property
'MaxLogCount日志文件文件最大行数(R/W)
Public Property Get MaxLogCount() As Long
    MaxLogCount = mlngMaxLogCount
End Property

Public Property Let MaxLogCount(ByVal lngMaxLogCount As Long)
    mlngMaxLogCount = lngMaxLogCount
    If mlngMaxLogCount < 512 Then
        mlngMaxLogCount = 512
    End If
End Property
'CurrentLogCount当前已经记录的行数(R)
Public Property Get CurrentLogCount() As Long
    CurrentLogCount = mlngCurLogCount
End Property
'===========================================================================
'                3、公共方法
'===========================================================================
'--------------------------------------------------------------------------------------------------
'方法           LogOpen
'功能           打开运行日志
'返回值         Boolean                 是否成功打开日志
'入参列表:
'参数名         类型                    说明
'strLogName     String                  日志名称，可以使全路径或者只有名称
'blnAppend      Boolean                 是否是追加方式写日志
'blnLogLevel    RunLogLevel             记录的日志级别
'-------------------------------------------------------------------------------------------------
Public Function LogOpen(Optional ByVal strLogName As String, Optional ByVal blnAppend As Boolean, Optional ByVal lngMaxLogCount As Long = 2 ^ 30, Optional ByVal rllLogLevel As RunLogLevel = RLL_NoneLog, Optional ByVal blnNormalTime As Boolean) As Boolean
    Dim strLogFile  As String
    Dim arrTmp      As Variant, i       As Long
    Dim strTmp      As String
    
    On Error GoTo ErrH
    '只有些日志模式下才生成日志
'    If rllLogLevel > RLL_NoneLog Then
        '全路径模式
        If InStr(strLogName, "\") > 0 Then
            strLogFile = strLogName
        Else
            If Len(strLogName) <> 0 Then
                strLogFile = GetAppsoftPath & "\Log\日志跟踪\" & strLogName & ".log"
            Else
                strLogFile = GetAppsoftPath & "\Log\日志跟踪\" & App.EXEName & ".log"
            End If
        End If
        arrTmp = Split(strLogFile, "\")
        strTmp = arrTmp(0) & "\"
        For i = 1 To UBound(arrTmp) - 1
            strTmp = strTmp & arrTmp(i) & "\"
            If Not mobjFSO.FolderExists(strTmp) Then
                Call mobjFSO.CreateFolder(strTmp)
            End If
        Next
        If Not mobjLog Is Nothing And (mstrLogFile <> strLogFile) Then
            mlngCurLogCount = 0
            Call LogClose
        End If
        If blnAppend Then
            Set mobjLog = mobjFSO.OpenTextFile(strLogFile, ForAppending, True)
        Else
            Set mobjLog = mobjFSO.OpenTextFile(strLogFile, ForWriting, True)
        End If
        mstrText = ""
        If mlngCur = 0 Or Not blnAppend Then
            mlngStart = GetTickCount
        End If
'    Else
'        strLogFile = strLogName
'    End If
    mblnNormalTime = blnNormalTime
    mstrLogFile = strLogFile
    mblnAppend = blnAppend
    mrllCurrentLogLevel = rllLogLevel
    mlngMaxLogCount = lngMaxLogCount
    LogOpen = True
    Exit Function
ErrH:
    Err.Clear
    LogOpen = False
End Function
'--------------------------------------------------------------------------------------------------
'方法           LogClose
'功能           关闭日志
'返回值
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Public Function LogClose() As Boolean
    On Error Resume Next
    If Not mobjLog Is Nothing Then
        mobjLog.Close
        Set mobjLog = Nothing
    End If
    mstrText = ""
    mstrLogTiltle = ""
    mstrLogTilEx = ""
    mstrLogDistance = ""
    If Err.Number <> 0 Then Err.Clear
    LogClose = True
End Function
'--------------------------------------------------------------------------------------------------
'方法           PushMethod
'功能           将调用方法推入堆栈
'返回值
'入参列表:
'参数名         类型                    说明
'strMethod      String                  方法名
'arrPars        String                  参数列表
'-------------------------------------------------------------------------------------------------
Public Sub PushMethod(ByVal rllLogLevel As RunLogLevel, ByRef strMethod As String, ParamArray arrPars() As Variant)
    If mrllCurrentLogLevel > RLL_NoneLog Then
        If mrllCurrentLogLevel >= rllLogLevel Then
            mstrText = ""
            For mlngIndex = LBound(arrPars) To UBound(arrPars)
                mstrText = mstrText & "," & DisPlayOneValue(arrPars(mlngIndex))
            Next
            mstrText = Mid(mstrText, 2)
        End If
    End If
    With mcllMethodStack
        If .Count = 0 Then
            If Len(mstrText) = 0 Then
                .Add strMethod
            Else
                .Add strMethod & "(" & mstrText & ")"
            End If
        Else
            If Len(mstrText) = 0 Then
                .Add strMethod, , 1
            Else
                .Add strMethod & "(" & mstrText & ")", , 1
            End If
        End If
        If .Count > mlngStackLen Then .Remove .Count
    End With
    If mrllCurrentLogLevel > RLL_NoneLog Then
        If mrllCurrentLogLevel >= rllLogLevel Then
            If Not mobjLog Is Nothing Then
                Call GetLogTiltle
                If Len(mstrText) = 0 Then
                    mobjLog.WriteLine mstrLogTiltle & "┏" & strMethod
                Else
                    mobjLog.WriteLine mstrLogTiltle & "┏" & strMethod & "(" & mstrText & ")"
                End If
            End If
        End If
    End If
End Sub
'--------------------------------------------------------------------------------------------------
'方法           LogInfo
'功能           警告处理。该种错误可能只生成警告，并不进行错误捕获。
'返回值
'入参列表:
'参数名         类型                    说明
'strInfo        String                  警告信息
'strMethod      String                  错误发生的过程
'-------------------------------------------------------------------------------------------------
Public Sub LogInfo(ByVal rllLogLevel As RunLogLevel, ByRef strInfo As String, ParamArray arrPars() As Variant)
    If mrllCurrentLogLevel >= rllLogLevel Then
        If Not mobjLog Is Nothing Then
            mstrText = ""
            For mlngIndex = LBound(arrPars) To UBound(arrPars) Step 2
                If mlngIndex = UBound(arrPars) Then
                    mstrText = mstrText & arrPars(mlngIndex)
                Else
                    mstrText = mstrText & arrPars(mlngIndex) & "=" & arrPars(mlngIndex + 1) & "    "
                End If
            Next
            If Len(mstrText) <> 0 Then
                mstrText = strInfo & ":  " & mstrText
            Else
                mstrText = strInfo
            End If
            Call GetLogTiltle
            mobjLog.WriteLine mstrLogTiltle & "┣" & mstrText
        End If
    End If
End Sub
'--------------------------------------------------------------------------------------------------
'方法           ErrorCenter
'功能           错误处理中心
'返回值         Integer                 0-忽略继续执行，1-重试(Resume),2-中止程序
'入参列表:
'参数名         类型                    说明
'strMethod      String                  错误发生的过程
'-------------------------------------------------------------------------------------------------
Public Function ErrCenter(Optional ByVal rllLogLevel As RunLogLevel = RLL_RunError, Optional ByRef strMethod As String) As Integer
    If mrllCurrentLogLevel >= rllLogLevel Then
        If Not mobjLog Is Nothing Then
            Call GetLogTiltle
            mobjLog.WriteLine mstrLogTiltle & "┣" & strMethod & "  " & Err.Number & "-" & Err.Description
            mobjLog.WriteLine mstrLogTilEx & "┣" & "调用堆栈："
            For mlngIndex = 1 To mcllMethodStack.Count
                mobjLog.WriteLine mstrLogTilEx & "┣" & mcllMethodStack(mlngIndex)
            Next
        End If
    End If
    Call PopMethod(rllLogLevel, strMethod)
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'方法           PopMethod
'功能           将最近的入栈的方法移除，或者将指定方法之前入堆栈的方法移除（包含指定的方法）
'返回值
'入参列表:
'参数名         类型                    说明
'strMethod      String                  方法名称，不传时弹出最近入堆栈的方法
'-------------------------------------------------------------------------------------------------
Public Sub PopMethod(ByVal rllLogLevel As RunLogLevel, ByRef strMethod As String, ParamArray arrPars() As Variant)
    If mrllCurrentLogLevel > RLL_NoneLog Then
        If mrllCurrentLogLevel >= rllLogLevel Then
            mstrText = ""
            For mlngIndex = LBound(arrPars) To UBound(arrPars)
                mstrText = mstrText & "," & DisPlayOneValue(arrPars(mlngIndex))
            Next
            mstrText = Mid(mstrText, 2)
            If Not mobjLog Is Nothing Then
                Call GetLogTiltle
                If Len(mstrText) = 0 Then
                    mobjLog.WriteLine mstrLogTiltle & "┗" & strMethod
                Else
                    mobjLog.WriteLine mstrLogTiltle & "┗" & strMethod & "(" & mstrText & ")"
                End If
            End If
        End If
    End If
    With mcllMethodStack
        If Len(strMethod) <> 0 Then
            For mlngIndex = 1 To .Count
                If mcllMethodStack(mlngIndex) Like strMethod & "*" Then
                    Exit For
                End If
            Next
            If mlngIndex > .Count Then
                If .Count > 0 Then  '没有找到任何匹配，则删除一个即可
                    mlngIndex = 1
                Else                '没有数据则不删除
                    mlngIndex = 0
                End If
            End If
        Else
            mlngIndex = 1  '传空则只删除一个
        End If
        
        Do While mlngIndex > 0
            .Remove 1
            mlngIndex = mlngIndex - 1
        Loop
        mlngIndex = 1
    End With
End Sub
'===========================================================================
'                4、私有方法
'===========================================================================

'--------------------------------------------------------------------------------------------------
'方法           GetAppsoftPath
'功能           获取APPSOFT路径
'返回值         String
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function GetAppsoftPath() As String
    If IsDesinMode Then
        GetAppsoftPath = "C:\APPSOFT"
    Else
        GetAppsoftPath = Mid(App.Path & "\", 1, InStr(5, App.Path & "\", "\"))
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           IsDesinMode
'功能           当前是否是源码环境
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
End Function
'--------------------------------------------------------------------------------------------------
'方法           LogWriteHeader
'功能           写入日志头。
'返回值
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function LogWriteHeader() As Boolean
    Dim rllCurLogLevel      As RunLogLevel
    
    If mrllCurrentLogLevel > RLL_NoneLog Then
        rllCurLogLevel = mrllCurrentLogLevel
        If Not mobjLog Is Nothing Then
            If mlngCurLogCount <> 0 Then
                Me.CurrentLogLevel = RLL_NoneLog
                Call mobjFSO.DeleteFile(mstrLogFile, True)
                Set mobjLog = Nothing
                Me.CurrentLogLevel = rllCurLogLevel
            End If
            mlngCurLogCount = 0
            mobjLog.WriteLine mstrLogFile & "-" & Format(Now, "yyyy-mm-dd  hh:mm:ss") & " 进程ID=" & GetCurrentProcessId()
        End If
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           GetLogTiltle
'功能           由于GetTickCount可能归零以及返回负值，因此需要格式化处理。天:小时:分钟:秒.微妙
'返回值
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Sub GetLogTiltle()
    If mlngCurLogCount >= mlngMaxLogCount Or mlngCurLogCount = 0 Then
        Call LogWriteHeader
    End If
    mlngCurLogCount = mlngCurLogCount + 1
    If mblnNormalTime Then
        mstrLogTime = Format(Now, "YYYY-MM-DD hh:mm:ss")
        mstrLogDistance = GetLogDistance
        mstrLogTiltle = mstrLogTime & mstrLogDistance
        mstrLogTilEx = "                   " & mstrLogDistance
    Else
        mlngCur = GetTickCount
        If mlngCur < mlngStart Then
            mdblTmp = M_OFFSET_4 - LongToUnsigned(mlngStart) + LongToUnsigned(mlngCur)
        Else
            mdblTmp = mlngCur - mlngStart
        End If
        mstrLogTime = ""
        '天数计算
        mdblTmp = mdblTmp / M_DAY_MS
        mlngTmp = Int(mdblTmp)
        mstrLogTime = Format(mlngTmp, "00")
        '小时计算
        mdblTmp = (mdblTmp - mlngTmp) * M_DAY_MS / M_HOUR_MS
        mlngTmp = Int(mdblTmp)
        mstrLogTime = mstrLogTime & " " & Format(mlngTmp, "00")
        '分钟计算
        mdblTmp = (mdblTmp - mlngTmp) * M_HOUR_MS / M_MIN_MS
        mlngTmp = Int(mdblTmp)
        mstrLogTime = mstrLogTime & ":" & Format(mlngTmp, "00")
        '秒计算
        mdblTmp = (mdblTmp - mlngTmp) * M_MIN_MS / 1000
        mlngTmp = Int(mdblTmp)
        mstrLogTime = mstrLogTime & ":" & Format(mlngTmp, "00") & "." & Format(Int((mdblTmp - mlngTmp) * 1000), "000")
        mstrLogDistance = GetLogDistance
        mstrLogTiltle = mstrLogTime & mstrLogDistance
        mstrLogTilEx = "               " & mstrLogDistance
    End If
End Sub

'--------------------------------------------------------------------------------------------------
'方法           LongToUnsigned
'功能           无符号整形处理
'返回值         Double
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function LongToUnsigned(value As Long) As Double
    If value < 0 Then LongToUnsigned = value + M_OFFSET_4 Else LongToUnsigned = value
End Function
'--------------------------------------------------------------------------------------------------
'方法           GetLogDistance
'功能           获取日志距离时间标志的长度
'返回值         String
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Function GetLogDistance() As String
    If mcllMethodStack.Count = 0 Then
        GetLogDistance = " "
    Else
        GetLogDistance = String((mcllMethodStack.Count - 1) * 2, " ")
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           DisPlayOneValue
'功能           展示对象
'返回值         String
'入参列表:
'参数名         类型                    说明
'valValue       Variant                 传入的对象
'-------------------------------------------------------------------------------------------------
Private Function DisPlayOneValue(valValue As Variant) As String
    Dim strTmp  As String
    
    If IsArray(valValue) Then
        Dim i    As Long
        strTmp = "["
        For i = LBound(valValue) To UBound(valValue)
            strTmp = strTmp & DisPlayOneValue(valValue(i)) & ","
        Next
        If Len(strTmp) = 1 Then
            strTmp = strTmp & "]"
        Else
            strTmp = Mid(strTmp, 1, Len(strTmp) - 1) & "]"
        End If
    ElseIf IsNull(valValue) Then
        strTmp = "{NULL}"
    ElseIf IsEmpty(valValue) Then
        strTmp = "{EMPTY}"
    ElseIf IsObject(valValue) Then
        If valValue Is Nothing Then
            strTmp = "{NOTHING}"
        Else
            If mrllCurrentLogLevel >= RLL_ALLDetail Then
                strTmp = "{OBJECT(" + TypeName(valValue) + ")=" & Serialize(valValue) & "}"
            Else
                strTmp = "{OBJECT(" + TypeName(valValue) + ")}"
            End If
        End If
    Else
        If VarType(valValue) = vbString Then
            strTmp = """" & valValue & """"
        Else
            strTmp = CStr(valValue)
        End If
    End If
    DisPlayOneValue = strTmp
End Function
'--------------------------------------------------------------------------------------------------
'方法           Serialize               将对象或值序列化为字符串
'返回值         String                  序列化的字符串
'入参列表:
'参数名         类型                    说明
'objInfo        Variant                 对象或值
'strKeyName     String                  序列化的关键字
'-------------------------------------------------------------------------------------------------
Private Function Serialize(ByVal objInfo As Variant, Optional ByVal strKeyName As String = "K_Default") As String
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    
    On Error Resume Next

    objBag.WriteProperty strKeyName, objInfo
    If Err.Number = 330 Then
        '非法参数。  因为不支持持久性不能写对象。
        Serialize = "{NotPersistable}"
        Err.Clear
    Else
        bytData = objBag.Contents
        Serialize = EncodeBase64(bytData())
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           UnSerialize             将字符串反序列化为对象或具体的值
'返回值         Variant                 序列化字符串对应的对象或具体的值
'入参列表:
'参数名         类型                    说明
'strSource      String                  序列化字符串
'strKeyName     String                  序列化的关键字
'-------------------------------------------------------------------------------------------------
Private Function UnSerialize(ByVal strSource As String, Optional ByVal strKeyName As String = "K_Default") As Variant
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    
    On Error Resume Next
    If Len(strSource) = 0 Then Exit Function
    If strSource = "{NotPersistable}" Then
         Set UnSerialize = Nothing
    Else
        bytData = DecodeBase64(strSource, True)
        objBag.Contents = bytData
        If Not IsObject(objBag.ReadProperty(strKeyName)) Then
            UnSerialize = objBag.ReadProperty(strKeyName)
        Else
            Set UnSerialize = objBag.ReadProperty(strKeyName)
        End If
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           SerializeMulti          按顺序序列化多个信息
'返回值         String                  序列化的字符串
'入参列表:
'参数名         类型                    说明
'arrInfo        Variant                 多个序列化的对象
'[      ]       long                    按0开始索引，索引作为序列化的关键字
'-------------------------------------------------------------------------------------------------
Private Function SerializeMulti(ParamArray arrInfo() As Variant) As String
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    Dim i           As Long
    On Error Resume Next
    If UBound(arrInfo) < 0 Then Exit Function
    If UBound(arrInfo) = 0 And IsArray(arrInfo(0)) Then
        objBag.WriteProperty "KL", UBound(arrInfo(0))
        For i = LBound(arrInfo(0)) To UBound(arrInfo(0))
            If IsArray(arrInfo(0)(i)) Then
                objBag.WriteProperty "KD" & i, 1
                objBag.WriteProperty "K" & i, SerializeMulti(arrInfo(0)(i))
            Else
                objBag.WriteProperty "K" & i, arrInfo(0)(i)
            End If
            If Err.Number = 330 Then
                '非法参数。  因为不支持持久性不能写对象。
                Err.Clear
                objBag.WriteProperty "K" & i, Nothing
            End If
        Next
    Else
        objBag.WriteProperty "KL", UBound(arrInfo)
        For i = 0 To UBound(arrInfo)
            If IsArray(arrInfo(i)) Then
                objBag.WriteProperty "KD" & i, 1
                objBag.WriteProperty "K" & i, SerializeMulti(arrInfo(i))
            Else
                objBag.WriteProperty "K" & i, arrInfo(i)
            End If
            If Err.Number = 330 Then
                '非法参数。  因为不支持持久性不能写对象。
                Err.Clear
                objBag.WriteProperty "K" & i, Nothing
            End If
        Next
    End If
    bytData = objBag.Contents
    SerializeMulti = EncodeBase64(bytData())
End Function

'--------------------------------------------------------------------------------------------------
'方法           UnSerializeMulti        获取序列的对象
'返回值         Variant                 序列化的对象数组
'入参列表:
'参数名         类型                    说明
'strSource      String                  序列化字符串
'[      ]       long                    按0开始索引，索引作为序列化的关键字
'-------------------------------------------------------------------------------------------------
Private Function UnSerializeMulti(ByVal strSource As String) As Variant
    Dim objBag      As New PropertyBag
    Dim bytData()   As Byte
    Dim i           As Long, lngLen     As Long
    Dim arrVar()    As Variant
    
    On Error Resume Next
    If Len(strSource) = 0 Then Exit Function
    bytData = DecodeBase64(strSource, True)
    objBag.Contents = bytData
    lngLen = objBag.ReadProperty("KL")
    If lngLen > -1 Then
        ReDim Preserve arrVar(lngLen)
        For i = 0 To lngLen
            If Not IsObject(objBag.ReadProperty("K" & i)) Then
                If objBag.ReadProperty("KD" & i, 0) = 1 Then
                    arrVar(i) = UnSerializeMulti(arrVar(i))
                Else
                    arrVar(i) = objBag.ReadProperty("K" & i)
                End If
            Else
                Set arrVar(i) = objBag.ReadProperty("K" & i)
            End If
        Next
    End If
    UnSerializeMulti = arrVar()
End Function
'--------------------------------------------------------------------------------------------------
'方法           StringToUTF8Bytes       将字符串转换为UTF-8编码的字节数组
'返回值         Byte()                  16进制字符串转换的字节组
'入参列表:
'参数名         类型                    说明
'strInput      String                  16进制字符串
'-------------------------------------------------------------------------------------------------
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '先计算需求字节数
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    '然后转换
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function

'--------------------------------------------------------------------------------------------------
'方法           UTF8BytesToString       将UTF-8编码的字节数组转换为字符串
'返回值         String                  转换后的字符串
'入参列表:
'参数名         类型                    说明
'bytInpu        Byte(）                 字节数组
'-------------------------------------------------------------------------------------------------
Public Function UTF8BytesToString(bytInpu() As Byte) As String
    Dim lngBytesRequired As Long

    '先计算需求字节数
    lngBytesRequired = MultiByteToWideChar(CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal 0, 0)
     
    '然后转换
    UTF8BytesToString = String(lngBytesRequired, 0)
    MultiByteToWideChar CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal StrPtr(UTF8BytesToString), lngBytesRequired
End Function

'-------------------------------------------------------------------------------------------------
'方法           EncBase64Char           将6-bit字节转换为Base64字符
'返回值         Byte                    字符数值
'入参列表:
'参数名         类型                    说明
'bytValue       Byte                    转换的字节
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Private Function EncBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue < 26 Then '26个大写英文字母
        EncBase64Char = bytValue + &H41
    ElseIf bytValue < 52 Then '26个小写英文字母
        EncBase64Char = bytValue + &H61 - 26
    ElseIf bytValue < 62 Then '10个数字
        EncBase64Char = bytValue + &H30 - 52
    ElseIf bytValue = 62 Then
        EncBase64Char = &H2B '+
    Else
        EncBase64Char = &H2F '/
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           DecBase64Char           将Base64字符转换为6 bit字节
'返回值         Byte                    字符数值
'入参列表:
'参数名         类型                    说明
'bytValue       Byte                    待解码的字节
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Private Function DecBase64Char(ByVal bytValue As Byte) As Byte
    If bytValue >= &H41 And bytValue <= &H5A Then
        DecBase64Char = bytValue - &H41
    ElseIf bytValue >= &H61 And bytValue <= &H7A Then
        DecBase64Char = bytValue - &H61 + 26
    ElseIf bytValue >= &H30 And bytValue <= &H39 Then
        DecBase64Char = bytValue - &H30 + 52
    ElseIf bytValue = &H2B Then
        DecBase64Char = 62
    ElseIf bytValue = &H2F Then
        DecBase64Char = 63
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           EncodeBase64            进行Base64编码，返回Base64的字符串
'返回值         String                  Base64编码结果
'入参列表:
'参数名         类型                    说明
'varInput       Variant                 需要进行Base64编码的字符串或者字节数组，字符串采取UTF-8编码。Byte()类型前面的数组，元素个数传3的倍数，最后一次传递所有剩下的即可。
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Private Function EncodeBase64(varInput As Variant) As String
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, J              As Long, lngBit     As Long
    
    On Error GoTo ErrH
    
    If VarType(varInput) = vbString Then
        If Len(varInput) = 0 Then Exit Function
        '原始内容,先将原文以UTF-8的方式编码
        bytInput = StringToUTF8Bytes(CStr(varInput))
    ElseIf VarType(varInput) = vbArray + vbByte Then
        If UBound(varInput) < 0 Then Exit Function
        bytInput = varInput
    Else
        Exit Function
    End If
    lngInputLen = UBound(bytInput) + 1
 
    lngOutLen = lngInputLen + (lngInputLen - 1) \ 3 + 1
    ReDim bytOut(lngOutLen - 1)
    '将8-bit字节数组转换为6-bit字节数组
    For i = 0 To lngInputLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(J) = (bytInput(i) And &HFC) \ &H4
            J = J + 1
            bytOut(J) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两位
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &HF0) \ &H10)
            J = J + 1
            bytOut(J) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四位
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &HC0) / &H40)
            J = J + 1
            bytOut(J) = bytInput(i) And &H3F
            J = J + 1
            lngBit = 0 '67PPPP 'P:Prev byte '012345
        End If
    Next

    For i = 0 To lngOutLen - 1
        bytOut(i) = EncBase64Char(bytOut(i)) '转换为Base64字符
    Next
    EncodeBase64 = StrConv(bytOut, vbUnicode) & String(2 - (lngInputLen - 1) Mod 3, "=") '原文剩余内容不足3个字节需要补齐
    Exit Function
ErrH:
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           DecodeBase64            将Base64的字符串解码为原文。
'返回值         Variant                 原始字符或者原始的字节组
'入参列表:
'参数名         类型                    说明
'strInput       String                  Base64编码字符串
'blnByteArray   Boolean                 True:返回Byte(),False-返回string
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Private Function DecodeBase64(strInput As String, Optional ByVal blnByteArray As Boolean) As Variant
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, J              As Long, lngBit     As Long
    Dim lngModLen       As Long
    On Error GoTo ErrH
    If Len(strInput) = 0 Then Exit Function
    lngModLen = InStr(strInput, "=")
    If lngModLen > 0 Then
        '编码后的内容
        lngModLen = Len(strInput) - lngModLen + 1
        bytInput = StrConv(strInput, vbFromUnicode)
    Else
        lngModLen = 0
        '编码后的内容
        bytInput = StrConv(strInput, vbFromUnicode)
    End If
    lngInputLen = UBound(bytInput) + 1
 
    '原始内容
    lngOutLen = lngInputLen - lngInputLen \ 4
    lngOutLen = lngOutLen - lngModLen
    ReDim bytOut(lngOutLen - 1)
 
    For J = 0 To lngInputLen - 1
        bytInput(J) = DecBase64Char(bytInput(J)) '从Base64字符转换为6-bit字节
    Next
    '将6-bit字节数组转换为8-bit字节数组
    For J = 0 To lngOutLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(J) = bytInput(i) * &H4
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &H30) \ &H10)
            lngBit = 2
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两字节
            bytOut(J) = (bytInput(i) And &HF) * &H10
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(J) = bytOut(J) Or ((bytInput(i) And &H3C) \ &H4)
            lngBit = 4
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四字节
            bytOut(J) = (bytInput(i) And &H3) * &H40
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(J) = bytOut(J) Or bytInput(i)
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            lngBit = 0
        End If
    Next
    If blnByteArray Then
        DecodeBase64 = bytOut
    Else
        '最后将转换得到的UTF-8字符串转换为VB支持的Unicode字符串以便于显示。
        DecodeBase64 = UTF8BytesToString(bytOut)
    End If
    Exit Function
ErrH:
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'方法           EncodeBase64_file       对文件进行Base64编码，返回Base64的字符串
'返回值         String                  Base64编码结果
'入参列表:
'参数名         类型                    说明
'strFile        String                  需要进行Base64编码的文件
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Private Function EncodeBase64_File(ByVal strFile As String) As String
    Dim lngFileNum  As Long, lngFileSize    As Long, lngModSize As Long, lngBlocks As Long
    Dim lngCount    As Long, lngCurSize     As Long
    Dim strReturn   As String
    Dim aryChunk()    As Byte
    
    Const conChunkSize      As Long = 3000
    
    On Error GoTo ErrH
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    If lngFileSize <> 0 Then
        lngModSize = lngFileSize Mod conChunkSize
        lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
        For lngCount = 0 To lngBlocks
            If lngCount = lngFileSize \ conChunkSize Then
                lngCurSize = lngModSize
                ReDim aryChunk(lngCurSize - 1) As Byte
            Else
                lngCurSize = conChunkSize
                If lngCount = 0 Then '防止不停分配内存
                    ReDim aryChunk(lngCurSize - 1) As Byte
                End If
            End If
            Get lngFileNum, , aryChunk()
            strReturn = strReturn & EncodeBase64(aryChunk)
        Next
        Close lngFileNum
        EncodeBase64_File = strReturn
    End If
    Exit Function
ErrH:
    Close lngFileNum
    Err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           DecodeBase64_File       将Base64的字符串解码为原文。
'返回值         String                  生成的文件名
'入参列表:
'参数名         类型                    说明
'strInput       String                  Base64编码字符串
'strFile        String                  指定文件名
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'-------------------------------------------------------------------------------------------------
Private Function DecodeBase64_File(strInput As String, Optional ByVal strFile As String) As String
    Dim lngFileNum  As Long, lngFileSize    As Long
    Dim lngCount    As Long, lngCurSize     As Long
    Dim strTmp      As String
    Dim aryChunk()    As Byte
    Const conChunkSize      As Long = 4000
    
    On Error GoTo ErrH
    If strFile = "" Then
        strFile = mobjFSO.GetSpecialFolder(TemporaryFolder) & "\" & mobjFSO.GetTempName
    Else
        If mobjFSO.FileExists(strFile) Then Kill strFile
    End If
    lngFileNum = FreeFile
    Open strFile For Binary As lngFileNum
    lngCount = 0
    lngCurSize = 0
    lngFileSize = Len(strInput)
    If lngFileSize <> 0 Then
        For lngCount = 1 To lngFileSize Step conChunkSize
            strTmp = Mid(strInput, lngCount, conChunkSize)
            aryChunk = DecodeBase64(strTmp, True)
            Put lngFileNum, , aryChunk()
        Next
        Close lngFileNum
    End If
    DecodeBase64_File = strFile
    Exit Function
ErrH:
    Close lngFileNum
    Err.Clear
End Function
Private Sub Class_Terminate()
    Call LogClose
    Set mobjFSO = Nothing
End Sub
