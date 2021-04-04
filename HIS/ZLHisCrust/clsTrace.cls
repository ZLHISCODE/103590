VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************
'**模 块 名：clsTrace
'**创 建 人：lshuo
'**日    期：2016/10/11
'**描    述：本类模块提供了日志公共类，用来支持日志跟踪
'**版    本：V1.0
'*************************************************************************
'===========================================================================
'                0、模块变量与公共类型
'===========================================================================
Public Enum LogTimeType                                 '日志时间类型
    LTT_None = 0                                        '不添加时间
    LTT_FullDate = 1                                    '全日期时间格式
    LTT_OnlyTime = 2                                    '只有时间
End Enum
Public Enum SectionLevel                                '段级别
    SL_LevelOne = 0                                     '一级段
    SL_LevelTwo = 1                                     '二级段
    SL_LevelThree = 2                                   '三级段
End Enum
Private mobjLog                 As TextStream
Private Const INFO_LENGTH       As Long = 15
Private mlngLoop                As Long
'===========================================================================
'                1、保存属性的变量
'===========================================================================
'LogFile日志文件文件名称(R)
Private mstrLogFile             As String
'===========================================================================
'                2、属性定义
'===========================================================================
'LogFile日志文件文件名称(R)
Public Property Get LogFile() As String
    LogFile = mstrLogFile
End Property
'===========================================================================
'                3、公共方法
'===========================================================================
Public Function OpenTace(ByVal strTaceName As String, Optional ByVal strSetupPath As String) As Boolean
    Dim objFSO      As New FileSystemObject, strFolder      As String
    
    On Error GoTo ErrH
    If Not objFSO.FolderExists(strSetupPath) Then
        Call objFSO.CreateFolder(strSetupPath)
    End If
    strFolder = strSetupPath & "\Log"
    If Not objFSO.FolderExists(strFolder) Then
        Call objFSO.CreateFolder(strFolder)
    End If
    strFolder = strFolder & "\日志跟踪"
    If Not objFSO.FolderExists(strFolder) Then
        Call objFSO.CreateFolder(strFolder)
    End If
    mstrLogFile = strFolder & "\" & strTaceName & "_" & Format(Now, "YYMMDDHHmmss") & ".log"
    Set mobjLog = objFSO.CreateTextFile(mstrLogFile, True)
    OpenTace = True
    Exit Function
ErrH:
    mstrLogFile = ""
    MsgBox "打开日志跟踪失败！错误信息：" & vbNewLine & Err.Description, vbInformation, App.Title
    Err.Clear
End Function

Public Sub WriteSection(ByVal strSectionName As String, Optional ByVal slLevel As SectionLevel = SL_LevelOne)
'功能：日志中写入一个阶段
'参数：strSectionName=段名
    If Not mobjLog Is Nothing Then
        If slLevel = SL_LevelOne Then
            mobjLog.WriteLine String(120, "*")
            mobjLog.WriteLine String(4, "*") & LPAD(strSectionName, (120 - 4 * 2 + Len(strSectionName)) \ 2)
            mobjLog.WriteLine String(120, "*")
        ElseIf slLevel = SL_LevelTwo Then
            mobjLog.WriteLine String(100, "=")
            mobjLog.WriteLine String(4, "=") & LPAD(strSectionName, (100 - 4 * 2 + Len(strSectionName)) \ 2)
            mobjLog.WriteLine String(100, "=")
        ElseIf slLevel = SL_LevelThree Then
            mobjLog.WriteLine RPAD(LPAD(strSectionName, (50 + ActualLen(strSectionName) \ 2), "-"), 100, "-")
        End If
    End If
End Sub

Public Sub WriteInfo(ByVal strInfoName As String, ParamArray arrInfo() As Variant)
'功能：日志中一行信息，该信息分为多个子项
    mobjLog.Write "[" & RPAD(strInfoName, INFO_LENGTH) & "]"
    For mlngLoop = LBound(arrInfo) To UBound(arrInfo) Step 2
        mobjLog.Write arrInfo(mlngLoop) & ":" & arrInfo(mlngLoop + 1) & "    "
    Next
    mobjLog.WriteLine
End Sub

'日志方法
Public Sub WriteLog(Optional ByVal strText As String, Optional ByVal lttAddTime As LogTimeType = LTT_None, Optional ByVal lngReturnLines As Long)
'参数:strText       =要写入的一行日志文本,如果为空表示写一行换行符
'     bytAddTime    >0时，在日志文本之前加入日志时间，1=日期和时间完整格式,2-仅时间,0-不加入
'     lngReturnLines=在日志文本之后写多少行换行符,0-不写换行符
    If Not mobjLog Is Nothing Then
        '添加日志时间
        If lttAddTime <> LTT_None Then strText = LogTime(lttAddTime = LTT_OnlyTime) & strText
        '截取超长内容
        If Len(strText) > 500 Then strText = Mid(strText, 1, 500)
        '写日志文本
        If strText = "" Then
            mobjLog.WriteBlankLines 1
        Else
            mobjLog.WriteLine strText
            If lngReturnLines > 0 Then mobjLog.WriteBlankLines lngReturnLines
        End If
    End If
End Sub

Public Sub CloseLog()
    '关闭日志文件
    mobjLog.Close
    Set mobjLog = Nothing
End Sub
'===========================================================================
'                4、私有方法
'===========================================================================
Private Function LogTime(Optional blnOnlyTime As Boolean) As String
    If blnOnlyTime Then
        LogTime = Format(Now, "HH:mm:ss")
    Else
        LogTime = Format(Now, "yyyy-MM-dd HH:mm:ss")
    End If
End Function

Private Function LPAD(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = ActualLen(strCode)
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = strCode
    End If
    LPAD = Replace(strTmp, Chr(0), strChar)
End Function

Private Function RPAD(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = ActualLen(strCode)
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = strCode
    End If
    '取掉最后半个字符
    RPAD = Replace(strTmp, Chr(0), strChar)
End Function
