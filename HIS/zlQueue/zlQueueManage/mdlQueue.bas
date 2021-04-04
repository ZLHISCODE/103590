Attribute VB_Name = "mdlQueue"
Option Explicit

Public Const MS_SOUND_TYPE = "微软语音"
Public Const DEF_SOUND_TYPE = "系统默认"
Private Const ISDEBUG = False 'True
Public Const G_LNG_QUEUEMANAGE_MODULENUM = 1160

'呼叫列宽初始化
Public Const C_STR_QUEUECALL = "0,0,0,0,50,0,90,0,60,0,0,60,60,0,0,60,0,0,125"
'排队列宽初始化
Public Const C_STR_QUEUEQUEUE = "0,0,0,30,50,0,90,40,60,60,0,60,60,50,125,0,120,60,0"

'LED显示相关变量
Public plngLEDModal As Long                'LED模块代码
Public prsLEDComponent As New ADODB.Recordset  'LED部件的数据库信息
Public pobjLEDShow As Object               'LED部件
Public glngSys As Long
Public glngModul As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'输出调试字符串
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'语音播放的函数
Public Declare Function StartTextPlay Lib "StrSound.dll" (ByVal PlayText As String, ByVal intxx As Integer) As Long
Public Declare Function StopPlayStr Lib "StrSound" () As Long
Public Declare Function InitHtTextSound Lib "StrSound.dll" () As Boolean
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long


'判断数组是否为空
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long


Public Sub DebugMsg(ByVal strClassName, strMethod, strParameter As String, Optional strExceptionMsg As String = "")
    If ISDEBUG Then
        Call OutputDebugString(Now & ">> [调用过程：" & strClassName & "." & strMethod & "]  [参数内容：" & strParameter & "]  " & _
                                IIf(Trim(strExceptionMsg) <> "", "[异常信息：" & strExceptionMsg & "]", ""))
    End If
End Sub


Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function StrNvl(ByVal varValue As String, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    StrNvl = IIf(Trim(varValue) = "", DefaultValue, varValue)
End Function

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function

Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取指定字串的值,字串中可以包含汉字
    '--入参数:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '--出参数:
    '--返  回:子串
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    err = 0
    On Error GoTo errHand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
errHand:
    Substr = ""
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function
