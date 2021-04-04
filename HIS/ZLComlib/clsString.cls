VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'字符串用UTF-8编码
Private Const CP_UTF8 = 65001
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
Private mobjFSO         As New FileSystemObject
'##############################################################################
'#功能：字符串相关的处理、判断等公共函数
'命名规则：函数过程名称为通俗易懂的英文即可，不用加Str关键字
'上级类：clsComLib
'##############################################################################

Public Function CheckCharScope(ByVal strText As String, ByVal strScopeChars As String) As Boolean
'功能：检查字符串是否只包含指定的字符
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strScopeChars, Mid(strText, i, 1)) = 0 Then Exit Function
    Next
    CheckCharScope = True
End Function

Public Function ListMinus(ByVal strLeft As String, ByVal strExists As String, Optional ByVal strSplit As String = ",") As String
'功能：返回"传入项目串"不在"已有项目串"中不重复的项目串
    Dim arrIDs As Variant, strResult As String, i As Long
    
    arrIDs = Split(strLeft, strSplit)
    For i = 0 To UBound(arrIDs)
        If InStr(strSplit & strExists & strSplit, strSplit & arrIDs(i) & strSplit) = 0 _
            And InStr(strResult & strSplit, strSplit & arrIDs(i) & strSplit) = 0 Then
            strResult = strResult & strSplit & arrIDs(i)
        End If
    Next
    
    ListMinus = Mid(strResult, Len(strSplit) + 1)
End Function

Public Function TrimEx(ByVal strTrim As String, Optional ByVal strTrmChar As String = " ") As String
'功能：去除strTrim两边的strTrmChar,功能类似Trim
'         不传strTrmChar或者传空格时，相当Trim
    Dim i As Integer, intB As Integer, intE As Integer
    
    If strTrim = "" Or strTrmChar = "" Then TrimEx = strTrim: Exit Function
    If strTrmChar = " " Then TrimEx = Trim(strTrim): Exit Function
    
    intB = 1
    For i = 1 To Len(strTrim)
        If Mid(strTrim, i, 1) <> strTrmChar Then intB = i: Exit For
    Next
    intE = Len(strTrim)
    For i = Len(strTrim) To 1 Step -1
        If Mid(strTrim, i, 1) <> strTrmChar Then intE = i: Exit For
    Next
    TrimEx = Mid(strTrim, intB, intE - intB + 1)
End Function

Public Function Encode(ByVal strPutString As String) As String
'功能:字符串加密
'入参:strPutString-需要加密的串
'返回:加密串
    Encode = gobjComLib.zlCommFun.zlStringEncode(strPutString)
End Function

Public Function NeedCode(strList As String, Optional ByVal strSplit As String) As String
'功能：从编码名称组合串中分离出编码
'参数：strList=编码名称组合串,如"012-内科","(012)内科","[012]内科"
'          strSplit=指定的编码名称分割符，没有指定，则按默认优先级进行解析,解析符只能在如下或者其他单个中间分割符
'说明:1-strList以()或[]分割编码与名称时，必须以[编码]或(编码)开头,编码必须为数字或字母
'     2-分隔符有优先级：回车符(Chr(13)）>换行(Chr(10))> - > [] > ()
    Dim intType As Integer
    
    If strList = "" Then Exit Function
    intType = gobjComLib.Decode(strSplit, "", 0, Chr(13), 1, Chr(10), 2, "-", 3, "[]", 4, "()", 5, 6)
    If intType = 0 Or intType = 1 Then
        '优先判断以回车符分割
        If InStr(strList, Chr(13)) > 0 Then
            NeedCode = LTrim(Mid(strList, 1, InStr(strList, Chr(13)) - 1))
            Exit Function
        End If
    End If
    
    If intType = 0 Or intType = 2 Then
        '以换行符分割
        If InStr(strList, Chr(10)) > 0 Then
            NeedCode = Trim(Mid(strList, InStr(strList, Chr(10)) - 1))
            Exit Function
        End If
    End If
    
    If intType = 0 Or intType = 4 Then
        '以[]分割
        If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
            If IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
                NeedCode = Trim(Mid(strList, 2, InStr(strList, "]") - 2))
                Exit Function
            End If
        End If
    End If
    
    If intType = 0 Or intType = 5 Then
        '以()分割
        If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
            If IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
                NeedCode = Trim(Mid(strList, 2, InStr(strList, ")") - 2))
                Exit Function
            End If
        End If
    End If
    If intType = 0 Or intType = 3 Then
        '以-分割
        If InStr(strList, "-") > 0 Then
            NeedCode = Trim(Mid(strList, 1, InStr(strList, "-") - 1))
        End If
    Else
        If InStr(strList, strSplit) > 0 And strSplit <> "" Then
            NeedCode = Trim(Mid(strList, 1, InStr(strList, strSplit) - 1))
        End If
    End If
End Function

Public Function NeedName(strList As String, Optional ByVal strSplit As String) As String
'功能：从编码名称组合串中分离出名称
'参数：strList=编码名称组合串,如"012-内科","(012)内科","[012]内科"
'          strSplit=指定的编码名称分割符，没有指定，则按默认优先级进行解析,解析符只能在如下或者其他单个中间分割符
'说明:1-strList以()或[]分割编码与名称时，必须以[编码]或(编码)开头,编码必须为数字或字母
'     2-分隔符有优先级：回车符(Chr(13)）>换行(Chr(10))> - > [] > ()
    NeedName = gobjComLib.zlCommFun.GetNeedName(strList, strSplit)
End Function

Public Function To_Date(ByVal strDate As String, Optional ByVal strType As String = "YMDHMS") As String
'功能：获取ORACLE Date类型串
'参数：strDate=时间字符串
'         strType=格式字符串类型，ymd-年月日（yyyy-mm-dd)，ymdhm-（yyyy-mm-dd hh:mm),ymdhms-（yyyy-mm-dd hh:mm:ss)
'返回：ORACLE Date类型串
    If Not IsDate(strDate) Then To_Date = "Null": Exit Function
    Select Case UCase(strType)
        Case "YMD"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "YMDHM"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Case "YMDHMS"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
           To_Date = "Null"
    End Select
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'功能:读取指定字串的值,字串中可以包含汉字
 '入参:strInfor-原串
 '         lngStart-直始位置
'         lngLen-长度
'返回:子串
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo ErrH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
ErrH:
    Err.Clear
    SubB = ""
End Function

Public Function SetBit(ByVal StrBit As String, ByVal intBit As Integer, Optional ByVal intVal As Integer = -1) As String
'功能：将指定位字符串strBit中的第intBit位设置为0或1
'参数：intVal=设置值,0或1,不传表示反转
    If intVal = -1 Then intVal = IIf(Val(Mid(StrBit, intBit, 1)) = 0, 1, 0)
    SetBit = Left(StrBit, intBit - 1) & intVal & Mid(StrBit, intBit + 1)
End Function

Public Function PrefixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：根据ZLHIS的单据号前缀规则,返回大写的单据号年前缀
'参数：curDate=单据号前缀基准时间,不传时基准时间为服务器时间
'应用：该函数主要是配合GetFullNO函数应用

    If curDate = #1/1/1900# Then
        PrefixNO = CStr(CInt(Format(gobjComLib.zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PrefixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PrefixNO = IIf(CInt(PrefixNO) < 10, PrefixNO, Chr(55 + CInt(PrefixNO)))
End Function

Public Function CPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'功能：使字符串居中，在两侧填充字符串
'参数：
'       strText:填充字符串
'       intCount:填充后的长度
'       StrPAD:填充的字符
'       blnAutoSub:字符串超长后自动截取
'返回：处理后的字符串
'说明：一个汉字当作两个字符长度处理
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill * 2 Then
        CPAD = String((intCount - lngTmp) \ (lngFill * 2), StrPAD) & strText & String((intCount - lngTmp) \ (lngFill * 2), StrPAD)
    ElseIf lngTmp > intCount And blnAutoSub Then
        CPAD = SubB(strText, 1, intCount)
    Else
        CPAD = strText
    End If
End Function

Public Function RPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'功能：等同Oracle的RPAD函数
'功能:按指定长度填制空格
 '参数：
 '       strText:填充字符串
 '       intCount:填充后的长度
 '       StrPAD:填充的字符
 '       blnAutoSub:字符串超长后自动截取
'返回:返回字串
   
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill Then
        RPAD = strText & String((intCount - lngTmp) \ lngFill, StrPAD)
    ElseIf lngTmp > intCount And blnAutoSub Then
        RPAD = SubB(strText, 1, intCount)
    Else
        RPAD = strText
    End If
End Function

Public Function LPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'功能：等同Oracle的LPAD函数
 '功能:按指定长度填制空格
 '参数：
 '  strText:填充字符串
 '  intCount:填充后的长度
 '  StrPAD:填充的字符
 '  blnAutoSub:字符串超长后自动截取
 '返回:返回字串
 
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill Then
        LPAD = String((intCount - lngTmp) \ lngFill, StrPAD) & strText
    ElseIf lngTmp > intCount And blnAutoSub Then
        LPAD = SubB(strText, 1, intCount)
    Else
        LPAD = strText
    End If
End Function

Public Function IsNumOrChar(ByVal strAsk As String) As Boolean
'功能：判断指定字符串是否全部由数字和英文字母构成，用于允许数字
'       和字母但不允许特殊字符的情况下的检测，isnumberic只能判断数字。
'参数：（SSC编制）strAsk
'返回：
    IsNumOrChar = gobjComLib.zlCommFun.IsNumOrChar(strAsk)
End Function

Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
'功能:检查指定的权限是否存在
'参数:strPrivs-权限串
'     strMyPriv-具体权限
'返回,存在权限,返回true,否则返回False
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
'功能：判断指定字符串是否含有汉字
'参数：strAsk
'返回：
    IsCharChinese = gobjComLib.zlCommFun.IsCharChinese(strAsk)
End Function

Public Function IsCharAlpha(ByVal strAsk As String) As Boolean
'功能：判断指定字符串是否全部由英文字母构成    '
'参数： strAsk
'返回：
    IsCharAlpha = gobjComLib.zlCommFun.IsCharAlpha(strAsk)
End Function

Public Function Increase(ByVal strVal As String, Optional ByVal blnDown As Boolean, Optional ByRef strErr As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
'参数：strVal=要加1的字符串
'         blnDown=True ,减一，False,加一
    Increase = gobjComLib.zlCommFun.IncStr(strVal, blnDown, strErr)
End Function

Public Function FullDate(ByVal strText As String, Optional blnTime As Boolean = True, Optional ByVal strMintime As String, Optional strMaxtTime As String) As String
'功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd[ HH:mm])
'参数：blnTime=是否处理时间部份
'参数：strMintime=生成时间的下限
'          strOutTime=生成时间的上限
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = gobjComLib.zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '输入串中包含日期分隔符
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                '只输入了日期部份
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                '只输入了时间部份
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '输入非法日期,返回原内容
            strTmp = strText
        End If
    Else
        '不包含日期分隔符
        If Len(strTmp) <= 2 Then
            '当作输入dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '当作输入MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '当作输入yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '当作输入MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '当作输入yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '当作输入yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) Then
        If strMintime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(strMintime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMintime
            End If
        End If
        If strMaxtTime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") > Format(strMaxtTime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMaxtTime
            End If
        End If
        If Not blnTime Then
            strTmp = Format(strTmp, "yyyy-MM-dd")
        End If
    End If
    FullDate = strTmp
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
'       实际数据存储长度
    ActualLen = gobjComLib.zlCommFun.ActualLen(strAsk)
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'参数：strText=处理字符
'         blnCrlf=是否去掉换行符
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    TruncZero = gobjComLib.zlCommFun.TruncZero(strInput)
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long, Optional strDefault As String = "") As String
'功能：将文本按Varchar2的长度计算方法进行截断
    ToVarchar = gobjComLib.zlCommFun.ToVarchar(varText, lngLength, strDefault)
End Function

Public Function FormatString(ByVal strFormat As String, ParamArray arrInput() As Variant) As String
'功能：用参数数组来填充strFormat中的[x]占位符，X>=1
'参数：
'  strFormat：表达式；[1..n]为参数号关键字；例子："测试值为：[1]"
'  arrInput：表达式的参数，对应strFormat中的参数号关键字
'返回：格式化后的字符串

    Dim strPar As String, arrPar As Variant, arrTmp As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intPar As Integer, i As Integer
    Dim strReturn As String, strTmp As String
    
    FormatString = strFormat
    If Len(strFormat) > 60000 Then Exit Function
    If UBound(arrInput) < 0 Then Exit Function
    On Error GoTo ErrH
    '分析自定的[x]参数
    lngLeft = InStr(1, strFormat, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strFormat, "]")
        If lngRight = 0 Then Exit Do
        strSeq = Mid(strFormat, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            If i > 0 And i - 1 <= UBound(arrInput) Then '有效参数
                strPar = strPar & ";" & i & "," & lngLeft & "," & lngRight
                If i > intPar Then intPar = i
            End If
        End If
        lngLeft = InStr(lngRight + 1, strFormat, "[")
    Loop
    '组装字符串
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ";")
    strReturn = ""
    For i = LBound(arrPar) To UBound(arrPar)
        arrTmp = Split(arrPar(i), ","): intPar = Val(arrTmp(0)): lngLeft = Val(arrTmp(1))
        strReturn = strReturn & Mid(strFormat, lngRight + 1, lngLeft - lngRight - 1) & arrInput(intPar - 1)
        lngRight = Val(arrTmp(2))
        If i = UBound(arrPar) Then strReturn = strReturn & Mid(strFormat, lngRight + 1)
    Next
    FormatString = strReturn
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean, Optional ByVal varDefault As Variant = 0) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
'参数：blnForceNum=当为Null时，是否强制表示为数字型
    Dim varTmp As Variant
    varTmp = IIf(Val(varValue & "") = 0, varDefault, varValue) & ""
    ZVal = IIf(Val(varTmp) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varTmp))
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'功能：四舍五入方式格式化数字
'参数：intBit=最大小数位数
'问题号：94552
'说明：VB自带的Round是银行家舍入法,与实际不一致。如Round(57.575,2)=57.58,Round(57.565,2)=57.56
    If intBit >= 0 Then
        RoundEx = Val(Format(dblNumber, "0" & IIf(intBit = 0, "", ".") & String(intBit, "0")))
    Else
        RoundEx = dblNumber
    End If
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun存在该函数
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = gobjComLib.NVL(varValue, DefaultValue)
End Function

Public Function VerCompare(ByVal strVerCur As String, Optional ByVal strVerCom As String) As Integer
'功能：比较两个版本号,比当前版本号小，返回1，相等返回0，比当前版本号大返回-1
'参数：strVerCur=当前版本号
'         strVerCom=对比的版本号
' 返回：对比版本号比当前版本号小，返回1，相等返回0，比当前版本号大返回-1
    If VerFull(strVerCur) < VerFull(strVerCom) Then
        VerCompare = -1
    ElseIf VerFull(strVerCur) > VerFull(strVerCom) Then
        VerCompare = 1
    Else
        VerCompare = 0
    End If
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'功能：返回VB最大支持的版本号形式:9999.9999.9999,最小版本号0000.0000.0000
'参数：strVer=当前版本号
'           blnMax=True,若果为空，则返回最大支持版本，False=若果为空，则返回最小支持版本
    Dim arrVer As Variant
    If Not IsVersion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999", "0000.0000.0000")
        Exit Function
    End If
    arrVer = Split(strVer, ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000")
End Function

Public Function IsVersion(ByVal strVer As String) As Boolean
'功能：判断字符串是否是版本号
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) <> 2 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If Val(arrVer(i)) & "" <> arrVer(i) Then Exit Function
    Next
    
    IsVersion = True
End Function

Public Function PinYinCode(ByVal strAsk As String) As String
'功能：返回指定字符串的拼音简码
'参数：strAsk  待处理的字符串
    Dim aryStard As Variant
    Dim intBit As Integer, iCount As Integer
    Dim strCode As String, StrBit As String

'    aryStard = Split("八;擦;哒;讹;发;噶;哈;击;;咔;垃;妈;拿;噢;啪;七;热;撒;他;挖;;挖;西;丫;匝", ";")
    aryStard = Split("八;擦;咑;屙;发;旮;铪;讥;;咔;垃;妈;拿;噢;啪;七;呥;撒;他;挖;;;西;丫;匝", ";")
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '将全角转换为半角，小写转换为大写
    
    strCode = ""
    For intBit = 1 To Len(strAsk)
        StrBit = Mid(strAsk, intBit, 1)
        If InStr(1, "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩαβγ蔫趴属哇娃夕汐仨兮拚嚓饧澶赕膪欹焘恁砉铛疋覃瞿她", StrBit) > 0 Then
            '特殊字的处理
            strCode = strCode & Switch(StrBit = "Ⅰ", "1", StrBit = "Ⅱ", "2", StrBit = "Ⅲ", "3", StrBit = "Ⅳ", "4", StrBit = "Ⅴ", "5" _
                            , StrBit = "Ⅵ", "6", StrBit = "Ⅶ", "7", StrBit = "Ⅷ", "8", StrBit = "Ⅸ", "9" _
                            , StrBit = "α", "A", StrBit = "β", "B", StrBit = "γ", "G" _
                            , StrBit = "蔫", "N", StrBit = "趴", "P", StrBit = "属", "S", StrBit = "哇", "W" _
                            , StrBit = "娃", "W", StrBit = "夕", "X", StrBit = "汐", "X", StrBit = "仨", "S" _
                            , StrBit = "兮", "X", StrBit = "拚", "P", StrBit = "嚓", "C", StrBit = "饧", "X" _
                            , StrBit = "澶", "C", StrBit = "赕", "D", StrBit = "膪", "C", StrBit = "欹", "Q" _
                            , StrBit = "焘", "T", StrBit = "恁", "N", StrBit = "砉", "H", StrBit = "铛", "D" _
                            , StrBit = "疋", "P", StrBit = "覃", "Q", StrBit = "瞿", "Q", StrBit = "她", "T")
        ElseIf Asc(StrBit) < 0 Then
            For iCount = 0 To UBound(aryStard)
                If Len(aryStard(iCount)) <> 0 Then
                    If StrComp(StrBit, aryStard(iCount), vbTextCompare) = -1 Then
                        strCode = strCode & Chr(65 + iCount)
                        Exit For
                    ElseIf iCount = UBound(aryStard) Then
                        strCode = strCode & "Z"
                    End If
                End If
            Next
        Else
            If StrBit >= "A" And StrBit <= "Z" Then
                strCode = strCode & StrBit
            End If
        End If
        '部门表修改名称允许录入100个字节长度，考虑录入数据名称与简码一致，此句不在控制
'        If Len(strCode) >= 10 Then Exit For
    Next
    PinYinCode = strCode

End Function

Public Function WBCode(ByVal strAsk As String, ByVal lng方式 As Long) As String
'功能：返回指定字符串的五笔型简码
'参数：strAsk  待处理的字符串
'      lng方式 1-取首字母，2-按五笔规则
    Static blnNotFound As Boolean
    Dim lngFile As Long, strFile As String, strReturn As String
    Dim str编码表 As String, str汉字 As String, bln前字母 As Boolean, str编码 As String
    Dim intBit As Integer, StrBit As String
    
    If blnNotFound = True Then
        'wbx.txt文件未找到，不能进行编码查询
        Exit Function
    End If
    
    '打开文件
    strFile = gstrAviPath
    If Right(strFile, 1) <> "\" Then strFile = strFile & "\"
'    strFile = "C:\AppSoft\"
    strFile = strFile & "wbx.txt"
    
    On Error Resume Next
    lngFile = FreeFile
    Open strFile For Input Access Read As lngFile
    If Err <> 0 Then
        blnNotFound = True
        MsgBox "未发现" & strFile & "文件。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '找到每一个字对应编码
    Do Until EOF(lngFile)
        Line Input #lngFile, strReturn
        If InStr(strAsk, Left(strReturn, 1)) > 0 Then
            '把这个判断放在内部，主要是为了加快速度，因为只有字经过第一个判断
            If InStr(strReturn, " ") > 0 Then
                str编码表 = str编码表 & strReturn & "|"
            End If
        End If
    Loop
    Close #lngFile
    str编码表 = UCase(str编码表)
    
    '得到字符串的中汉字
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '将全角转换为半角，将字符串文字转成小写
    If lng方式 = 1 Then
        '按首字母
        For intBit = 1 To Len(strAsk)
            StrBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(StrBit, vbFromUnicode)) = 2 Then
                '汉字
                str编码 = str编码 & mGet编码By汉字(str编码表, StrBit, 1)
                bln前字母 = False
            ElseIf InStr(" ,.;:", StrBit) > 0 Then
                '空格
                bln前字母 = False
            Else
                If bln前字母 = False And StrBit >= "A" And StrBit <= "Z" Then
                    '只取一个字符串的首字母
                    str编码 = str编码 & StrBit
                End If
                bln前字母 = True
            End If
        Next
    Else
        '按五笔规则
        For intBit = 1 To Len(strAsk)
            StrBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(StrBit, vbFromUnicode)) = 2 Then
                '汉字
                str汉字 = str汉字 & StrBit
            End If
        Next
        
        Select Case Len(str汉字)
            Case 0
            Case 1
               str编码 = mGet编码By汉字(str编码表, str汉字, 4)
            Case 2
               str编码 = mGet编码By汉字(str编码表, Mid(str汉字, 1, 1), 2) & mGet编码By汉字(str编码表, Mid(str汉字, 2, 1), 2)
            Case 3
               str编码 = mGet编码By汉字(str编码表, Mid(str汉字, 1, 1), 1) & mGet编码By汉字(str编码表, Mid(str汉字, 2, 1), 1) & mGet编码By汉字(str编码表, Mid(str汉字, 3, 1), 2)
            Case Else
               str编码 = mGet编码By汉字(str编码表, Mid(str汉字, 1, 1), 1) & mGet编码By汉字(str编码表, Mid(str汉字, 2, 1), 1) & _
                         mGet编码By汉字(str编码表, Mid(str汉字, 3, 1), 1) & mGet编码By汉字(str编码表, Right(str汉字, 1), 1)
        End Select
    End If
    
    WBCode = str编码
    If Err.Number <> 0 Then Err.Clear
End Function

Private Function mGet编码By汉字(ByVal str编码表 As String, ByVal str汉字 As String, ByVal lngLen As Long) As String
'功能：根据汉字得到其编码
    Dim lngStart As Long, lngEnd As Long
    Dim str编码 As String
    
    lngStart = InStr(str编码表, str汉字)
    If lngStart = 0 Then
        '未在编码表找到该字编码
        mGet编码By汉字 = "Z"
        Exit Function
    End If
    
    lngEnd = InStr(lngStart, str编码表, "|")
    str编码 = Mid(str编码表, lngStart, lngEnd - lngStart)
    mGet编码By汉字 = Mid(Split(str编码, " ")(1), 1, lngLen)
End Function

Public Function GetCodeByVB(ByVal strAsk As String) As String
'-------------------------------------------------------------
'功能：返回指定字符串的简码
'说明：根据指定字符串生成简码，可以生成三种类型的简码
'        0、拼音，取每字的首字母构成简码
'        1、五笔，取每字的首字母构成简码
'        2、五笔，按五笔规则构成简码
'      在传入的参数中未发现※符号，就按用户在系统选项中设置的方式生成简码；
'        否则就按在※符号后的数字指定的方式强制生成简码，如※1表示按五笔首字母生成
'-------------------------------------------------------------
    GetCodeByVB = gobjComLib.zlCommFun.SpellCode(strAsk)
End Function

Public Function GetCodeByORCL(strInput As String, Optional ByVal blnWB As Boolean, Optional lngLen As Long = 10) As String
'功能：生成字符串的简码
'入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '          lngLen=<1或>40返回的字符串长度为10，否则返回指定长度字符串
'出参：正确返回字符串；错误返回"-"
     GetCodeByORCL = gobjComLib.zlCommFun.zlGetSymbol(strInput, IIf(blnWB, 1, 0), lngLen)
End Function

Public Function FullPinYin(strText As String, Optional intCapital As Integer = 0, Optional blnUseSpliter As Boolean = True) As String
'功能：根据GBK编码提取汉字的全拼音
'参数： strText         要取得全拼的字符串
'       intCapital      大小写标记：0-大写；1-小写；2-首字母大写
'       blnUseSpliter   拼音之间是否返回空格：True-带空格；False-不带空格
'返回：             返回取得的全拼如果没有则为空
    FullPinYin = gobjComLib.zlCommFun.mGetFullPY(strText, intCapital, blnUseSpliter)
End Function

Public Function ChineseMoney(curMoney) As String
'功能：将指定的金额数值转换为大写金额显示
'参数：
'       curMoney:需要转换的金额数值
'返回：
    ChineseMoney = gobjComLib.zlCommFun.UppeMoney(curMoney)
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = True, _
    Optional ByVal blnAddZero As Boolean, _
    Optional intMinBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,
'          intBit=最大小数位数
'         blnShowZero=vNumber为0时是否显示0值
'         blnAddZero=小数位不足是否补零
'       intMinBit=最少保留小数位数,如:FormatEx(1,6,True,False,2) 结果为：1.00
'返回：格式化后的字符串
    Dim strNumber As String
    Dim intDot As Integer
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
    
    If Not blnAddZero Then '小数位最后不显示零。如1.0100 就为1.01
        If intBit < intMinBit Then intMinBit = intBit
        If vNumber = 0 Then
            If blnShowZero Then
                If intMinBit > 0 Then
                    strNumber = "0." & String(intMinBit, "0")
                Else
                    strNumber = "0"
                End If
            Else
                strNumber = ""
            End If
        ElseIf Int(vNumber) = vNumber Then
            If intMinBit > 0 Then
                strNumber = vNumber & "." & String(intMinBit, "0")
            Else
                strNumber = vNumber
            End If
        Else
            intDot = intBit - intMinBit
            strNumber = Format(vNumber, "0." & String(intBit, "0"))
            If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
            If InStr(strNumber, ".") > 0 Then
                Do While Right(strNumber, 1) = "0" And intDot > 0
                    strNumber = Left(strNumber, Len(strNumber) - 1)
                    intDot = intDot - 1
                Loop
                If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
        End If
    Else '小数位数不足补零.如3位小数，1.1就为1.100
        strNumber = Format(vNumber, "#0." & String(intBit, "0"))
    End If
    FormatEx = strNumber
End Function

Public Function ExpressValue(strExpress As String, Optional ByRef strErr As String) As Variant
    '功能               计算表达式
    '参数               strExpress = 计算表达式
    '返回               计算结果
    On Error GoTo ErrH
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    sc.Language = "VBScript"
    ExpressValue = sc.Eval(Trim(strExpress))
    Exit Function
ErrH:
    strErr = "出错函数(ExpressValue),出错信息:" & Err.Number & " " & Err.Description
    Err.Clear
End Function

Public Function SQLAdjust(Str As String) As String
'功能：将含有"'"符号的字符串调整为Oracle所能识别的字符常量
'说明：自动(必须)在两边加"'"界定符。

    Dim i As Long, strTmp As String
    
    If InStr(1, Str, "'") = 0 Then SQLAdjust = "'" & Str & "'": Exit Function
    
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
    SQLAdjust = strTmp
End Function

Public Function Cipher(ByVal password As String, ByVal from_text As String) As String
    
    Dim strTotext As String
    
    password = Base64Encode(password) & "WIZARDPAGE"
    
    '加密
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
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
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            Cipher = Cipher & Chr$(ch)
        End If
    Next i
End Function

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

'加解密字符串函数,不支持中文
Public Function Base64Encode(InStr1 As String) As String
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

Public Function JSONParse(ByVal JSONPath As String, ByVal JSONString As String) As Variant
'功能：对JSON字符串的解析
'参数：JSONPath 结点名称，JSONString 完整的JSON字符串
    Dim JSON As Object
    Set JSON = CreateObject("MSScriptControl.ScriptControl")
    JSON.Language = "JScript"
    JSONParse = JSON.Eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
    Set JSON = Nothing
End Function

'======================================================================================================================
'方法           StringToUTF8Bytes       将字符串转换为UTF-8编码的字节数组
'返回值         Byte()                  16进制字符串转换的字节组
'入参列表:
'参数名         类型                    说明
'strInput      String                  16进制字符串
'======================================================================================================================
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

'======================================================================================================================
'方法           UTF8BytesToString       将UTF-8编码的字节数组转换为字符串
'返回值         String                  转换后的字符串
'入参列表:
'参数名         类型                    说明
'bytInpu        Byte(）                 字节数组
'======================================================================================================================
Public Function UTF8BytesToString(bytInpu() As Byte) As String
    Dim lngBytesRequired As Long

    '先计算需求字节数
    lngBytesRequired = MultiByteToWideChar(CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal 0, 0)
     
    '然后转换
    UTF8BytesToString = String(lngBytesRequired, 0)
    MultiByteToWideChar CP_UTF8, 0, bytInpu(0), UBound(bytInpu) + 1, ByVal StrPtr(UTF8BytesToString), lngBytesRequired
End Function
'======================================================================================================================
'方法           EncBase64Char           将6-bit字节转换为Base64字符
'返回值         Byte                    字符数值
'入参列表:
'参数名         类型                    说明
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'======================================================================================================================
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
'======================================================================================================================
'方法           DecBase64Char           将Base64字符转换为6 bit字节
'返回值         Byte                    字符数值
'入参列表:
'参数名         类型                    说明
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'======================================================================================================================
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
'======================================================================================================================
'方法           EncodeBase64            进行Base64编码，返回Base64的字符串
'返回值         String                  Base64编码结果
'入参列表:
'参数名         类型                    说明
'varInput       Variant                 需要进行Base64编码的字符串或者字节数组，字符串采取UTF-8编码。Byte()类型前面的数组，元素个数传3的倍数，最后一次传递所有剩下的即可。
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'======================================================================================================================
Public Function EncodeBase64(varInput As Variant) As String
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    
    On Error GoTo ErrH
    
    If varType(varInput) = vbString Then
        If Len(varInput) = 0 Then Exit Function
        '原始内容,先将原文以UTF-8的方式编码
        bytInput = StringToUTF8Bytes(CStr(varInput))
    ElseIf varType(varInput) = vbArray + vbByte Then
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
            bytOut(j) = (bytInput(i) And &HFC) \ &H4
            j = j + 1
            bytOut(j) = (bytInput(i) And &H3) * &H10
            lngBit = 2 '234567 'NNNN01 'N:Next byte
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两位
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HF0) \ &H10)
            j = j + 1
            bytOut(j) = (bytInput(i) And &HF) * &H4
            lngBit = 4 '4567PP 'P:Prev byte 'NN0123 'N:Next byte
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四位
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &HC0) / &H40)
            j = j + 1
            bytOut(j) = bytInput(i) And &H3F
            j = j + 1
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
'======================================================================================================================
'方法           DecodeBase64            将Base64的字符串解码为原文。
'返回值         Variant                 原始字符或者原始的字节组
'入参列表:
'参数名         类型                    说明
'strInput       String                  Base64编码字符串
'blnByteArray   Boolean                 True:返回Byte(),False-返回string
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'======================================================================================================================
Public Function DecodeBase64(strInput As String, Optional ByVal blnByteArray As Boolean) As Variant
    Dim bytInput()  As Byte, lngInputLen    As Long
    Dim bytOut()    As Byte, lngOutLen      As Long
    Dim i           As Long, j              As Long, lngBit     As Long
    Dim lngLen      As Long
    Dim lngPos      As Long, lngModLen      As Long
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
 
    For j = 0 To lngInputLen - 1
        bytInput(j) = DecBase64Char(bytInput(j)) '从Base64字符转换为6-bit字节
    Next
    '将6-bit字节数组转换为8-bit字节数组
    For j = 0 To lngOutLen - 1
        If lngBit = 0 Then 'bytOut(J)未被写入
            bytOut(j) = bytInput(i) * &H4
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H30) \ &H10)
            lngBit = 2
        ElseIf lngBit = 2 Then 'bytOut(J)已被写入两字节
            bytOut(j) = (bytInput(i) And &HF) * &H10
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or ((bytInput(i) And &H3C) \ &H4)
            lngBit = 4
        ElseIf lngBit = 4 Then 'bytOut(J)已被写入四字节
            bytOut(j) = (bytInput(i) And &H3) * &H40
            i = i + 1
            If i > UBound(bytInput) Then Exit For
            bytOut(j) = bytOut(j) Or bytInput(i)
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
'======================================================================================================================
'方法           EncodeBase64_file       对文件进行Base64编码，返回Base64的字符串
'返回值         String                  Base64编码结果
'入参列表:
'参数名         类型                    说明
'strFile        String                  需要进行Base64编码的文件
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'======================================================================================================================
Public Function EncodeBase64_File(ByVal strFile As String) As String
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
'======================================================================================================================
'方法           DecodeBase64_File       将Base64的字符串解码为原文。
'返回值         String                  生成的文件名
'入参列表:
'参数名         类型                    说明
'strInput       String                  Base64编码字符串
'strFile        String                  指定文件名
'方法说明，Base64是将三个字节，每6位分割为四个字节处理的
'======================================================================================================================
Public Function DecodeBase64_File(strInput As String, Optional ByVal strFile As String) As String
    Dim lngFileNum  As Long, lngFileSize    As Long, lngModSize As Long, lngBlocks As Long
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
'======================================================================================================================
'方法           StringToUnicode         将字符串转换为Unicode编码
'返回值         String                  Unicode标识的字符串
'入参列表:
'参数名         类型                    说明
'strInput      String                   原始字符
'======================================================================================================================
Public Function StringToUnicode(ByVal strInput As String) As String
    Dim lngLen      As Long
    Dim i           As Long
    Dim strFirstChr As String, strUniChar As String
    Dim strReturn   As String
    
    lngLen = Len(strInput)
    For i = 1 To lngLen
        strFirstChr = Mid(strInput, i, 1)
        If Asc(strFirstChr) < 0 Then
            strUniChar = Hex(AscW(strFirstChr))
            strReturn = strReturn & "\u" & String(4 - Len(strUniChar), "0") & strUniChar
        Else
            Select Case strFirstChr
                Case "\", """", "/"
                    strReturn = strReturn & "\" & strFirstChr
                Case Chr(8), Chr(12), Chr(10), Chr(13), Chr(9)
                    strReturn = strReturn & Decode(Asc(strFirstChr), 8, "\b", 12, "\f", 10, "\n", 13, "\r", 9, "\t")
                Case Else
                    strReturn = strReturn & strFirstChr
            End Select
        End If
    Next
    StringToUnicode = strReturn
End Function
'======================================================================================================================
'方法           UnicodeToString         将Unicode编码转换为字符串
'返回值         String                  原始字符串
'入参列表:
'参数名         类型                    说明
'strUnistr      String                  Unicode编码字符
'======================================================================================================================
Public Function UnicodeToString(ByVal strUnistr As String) As String
    Dim lngLen      As Long
    Dim i           As Long
    Dim strFirstChr As String, strSecChar       As String
    Dim strReturn   As String
    
    lngLen = Len(strUnistr)
    i = 1
    Do While i <= lngLen
        strFirstChr = Mid(strUnistr, i, 1)
        If strFirstChr = "\" Then
            strSecChar = Mid(strUnistr, i + 1, 1)
            Select Case strSecChar
                Case "\", """", "/"
                    strReturn = strReturn & strSecChar
                    i = i + 2
                Case "'"
                    strReturn = strReturn & strSecChar
                    i = i + 2
                    Debug.Print "Error"
                Case "b", "f", "n", "r", "t"
                    strReturn = strReturn & Chr(Decode(strSecChar, "b", 8, "f", 12, "n", 10, "r", 13, "t", 9))
                    i = i + 2
                Case "u"
                    strSecChar = Mid(strUnistr, i + 2, 4)
                    If strSecChar Like "[0-9,A-E,a-e][0-9,A-E,a-e][0-9,A-E,a-e][0-9,A-E,a-e]" Then
                        strReturn = strReturn & ChrW("&H" & strSecChar)
                        i = i + 6
                    End If
            End Select
        Else
            strReturn = strReturn & strFirstChr
            i = i + 1
        End If
    Loop
    UnicodeToString = strReturn
End Function

Private Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Private Sub Class_Terminate()
    Set mobjFSO = Nothing
End Sub

