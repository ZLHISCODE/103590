Attribute VB_Name = "mdlPlugIn"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gblnInited As Boolean
Public gcolPlugIn As Collection '扩展部件，用集合方式暂存扩展部件类的实例


Public Enum Enum_Modue '模块号
    m门诊医嘱模块 = 1252
    m住院医嘱模块 = 1253
    m住院护士站模块 = 1254
    m临床路径模块 = 1256
    m病历模块 = 1070
    m人员管理模块 = 1002
    m医嘱附费模块 = 1257
    
    m门诊医生工作站 = 1260
    m住院医生工作站 = 1261
    m住院护士工作站 = 1262
    m医技工作站 = 1263
    
    m门诊医嘱下达 = 1252
    m住院医嘱下达 = 1253
    m新版护士站 = 1265
    
    '血库相关模块
    m科室配血管理 = 1935
    '科室配血管理页签没有模块号，增加序号表示对应页签
    m科室配血管理_常规复查 = 193501
    m科室配血管理_配血记录 = 193502
    m血液目录管理 = 1900
    m血液输血反应 = 1938
    m科室发血管理 = 1936
    m血液供应入库 = 1915
    m血液报废出库 = 1922
    
    'LIS相关模块
    m临床实验室管理 = 2500
    
    m病人信息管理 = 1101
    m电子病案打印 = 1566
    
    '药品卫材模块
    m卫材外购入库管理 = 1712
    m药品外购入库管理 = 1300
    m药品计划管理 = 1330
    m药品处方发药 = 1341
    m药品部门发药 = 1342
    m输液配置中心 = 1345
    
    m病人挂号模块 = 1111
    m病人收费模块 = 1121
    m病人结帐处理 = 1137
    
    m体检中心管理 = 2121
    m体检总检登记 = 2125
    m体检分科执行 = 2122
    m体检结果登记 = 2123
End Enum

'以下的所有代码是用于支持多插件同时挂接
'CkeckUseable 方法用于限制使用，编写扩展插件时直接传当前 单位名称即可。使用该方法时要引用 zl9ComLib.dll
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const ERROR_NO_MORE_ITEMS = 259&
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_LOCAL_MACHINE = &H80000002

'支持滑轮的常量
Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROC = -4

' 注册表数据类型...
Public Enum ValueType
    REG_SZ = 1                         ' 字符串值
    REG_EXPAND_SZ = 2                  ' 可扩充字符串值
    REG_BINARY = 3                     ' 二进制值
    REG_DWORD = 4                      ' DWORD值
    REG_MULTI_SZ = 7                   ' 多字符串值
End Enum

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private marrName As Variant
Public glngPreHWnd As Long '用于支持鼠标滚轮功能
Public gobjMec As Object '首页部件对象


'记录集相关变量
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Private Sub GetPathNames()
'功能：获取注册表CLSID下级目录

    Dim hKey As Long, Cnt As Long, sName As String, sData As String, Ret As Long, RetData As Long
    Const BUFFER_SIZE As Long = 255
    marrName = Array()
    Ret = BUFFER_SIZE
    If RegOpenKey(HKEY_CLASSES_ROOT, "CLSID", hKey) = 0 Then
        sName = Space(BUFFER_SIZE)
        While RegEnumKeyEx(hKey, Cnt, sName, Ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve marrName(UBound(marrName) + 1)
            marrName(UBound(marrName)) = "CLSID\" & Left$(sName, Ret)
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
        Wend
        RegCloseKey hKey
    End If
    Cnt = 0
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, ValueName As String, Optional ValueType As Long) As String
'功能：获得已存在的注册表关键字的值
'参数：ValueName="" 则返回 KeyName 项的默认值
'      如果指定的注册表关键字不存在, 则返回空串
'      KeyRoot--根类型, KeyName--子项名称, ValueName--值项名称, ValueType--值项类型
    Dim i As Integer
    Dim hKey As Long
    Dim TempValue As String                             ' 注册表关键字的临时值
    Dim Value As String                                 ' 注册表关键字的值
    Dim ValueSize As Long                               ' 注册表关键字的值的实际长度
    TempValue = Space(1024)                             ' 存储注册表关键字的临时值的缓冲区
    ValueSize = 1024                                    ' 设置注册表关键字的值的默认长度
    
    ' 打开一个已存在的注册表关键字...
    RegOpenKeyEx KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey
    
    ' 获得已打开的注册表关键字的值...
    RegQueryValueEx hKey, ValueName, 0, ValueType, ByVal TempValue, ValueSize
    
    ' 返回注册表关键字的的值...
    Select Case ValueType                                                        ' 通过判断关键字的类型, 进行处理
        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            TempValue = Left$(TempValue, ValueSize - 1)                          ' 去掉TempValue尾部空格
            Value = TempValue
        Case REG_DWORD
            ReDim dValue(3) As Byte
            RegQueryValueEx hKey, ValueName, 0, REG_DWORD, dValue(0), ValueSize
            For i = 3 To 0 Step -1
                Value = Value + String(2 - Len(Hex(dValue(i))), "0") + Hex(dValue(i))   ' 生成长度为8的十六进制字符串
            Next i
            If CDbl("&H" & Value) < 0 Then                                              ' 将十六进制的 Value 转换为十进制
                Value = 2 ^ 32 + CDbl("&H" & Value)
            Else
                Value = CDbl("&H" & Value)
            End If
        Case REG_BINARY
            If ValueSize > 0 Then
                ReDim bValue(ValueSize - 1) As Byte                                     ' 存储 REG_BINARY 值的临时数组
                RegQueryValueEx hKey, ValueName, 0, REG_BINARY, bValue(0), ValueSize
                For i = 0 To ValueSize - 1
                    Value = Value + String(2 - Len(Hex(bValue(i))), "0") + Hex(bValue(i)) + " "  ' 将数组转换成字符串
                Next i
            End If
    End Select
    
    ' 关闭注册表关键字...
    RegCloseKey hKey
    GetKeyValue = Trim(Value)                                                    ' 返回函数值
End Function

Private Function GetAllPlugIns() As String
'功能：获取扩展插件的部件名称，逗号割。
    Dim strTmp As String
    Dim strName As String
    Dim strResult As String
    Dim i As Integer
    
    Call GetPathNames
    
    For i = 1 To UBound(marrName)
        strResult = GetKeyValue(HKEY_CLASSES_ROOT, CStr(marrName(i)), strTmp, REG_SZ)
        '以ZLPLUGIN开头
        If UCase(Left(strResult, 8)) = "ZLPLUGIN" Then
            If InStr(strResult, ".") > 0 Then
                If Len(Split(strResult, ".")(0)) > 8 And InStr(strName, Split(strResult, ".")(0)) = 0 Then
                    strName = IIf(strName = "", "", strName & ",") & Split(strResult, ".")(0)
                End If
            End If
        End If
    Next
    GetAllPlugIns = strName
End Function

Public Function HandlePlugIn(ByVal bytType As Byte, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal cnOracle As ADODB.Connection, _
        Optional ByVal int场合 As Integer = -1, Optional strReserve As String, Optional strFuncName As String, Optional ByVal lngPatiID As Long, _
        Optional ByVal varRecId As Variant, Optional ByVal varKeyId As Variant)
'功能：扩展插件功能支持相关处理
'参数：bytType 操作类型 1=初始化，2=获取功能名，3=执行功能，4=终止。当bytType=2时 strFunName作为出参
'      cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的上级系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=保留参数,用于扩展使用
'      strFunName 出参和入参 当bytType=2时出参，当bytType=3时入参
'      lngPatiID=当前病人ID
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如医嘱ID
    Dim strTmp As String
    Dim strFuncNameTmp As String
    Dim strUserName As String
    Dim objTmp As Object
    Dim varArr As Variant
    Dim i As Integer
    Dim strTmpReserve As String
    Dim strReserveOther As String
    
    On Error Resume Next
    
    If bytType = 1 Then
        strTmp = GetAllPlugIns
        If strTmp = "" Then Exit Function
        varArr = Split(strTmp, ",")
        Set gcolPlugIn = New Collection
        For i = 0 To UBound(varArr)
            Set objTmp = CreateObject(varArr(i) & ".clsPlugIn")
            If Not objTmp Is Nothing Then
                Call objTmp.Initialize(cnOracle, lngSys, lngModual, int场合)
                '部件使用限制，用户名空时表示不限制
                strUserName = objTmp.GetUserName '医院用户--单位名称
                
                If strUserName <> "" Then
                    If CkeckUseable(strUserName) Then
                        gcolPlugIn.Add objTmp, "_" & varArr(i)
                    End If
                Else
                    gcolPlugIn.Add objTmp, "_" & varArr(i)
                End If
            End If
            Set objTmp = Nothing
        Next i
    End If
    
    If gcolPlugIn Is Nothing Then Exit Function
    
    If bytType = 2 Then
        For i = 1 To gcolPlugIn.Count
            Set objTmp = gcolPlugIn.Item(i)
            strTmp = ""
            strTmpReserve = ""
            strTmp = objTmp.GetFuncNames(lngSys, lngModual, int场合, strTmpReserve)
            strFuncNameTmp = IIf(strFuncNameTmp = "", "", strFuncNameTmp & ",") & strTmp
            strReserveOther = IIf(strReserveOther = "", "", strReserveOther & ",") & strTmpReserve
        Next i
        strFuncName = strFuncNameTmp
        strReserve = strReserveOther
    ElseIf bytType = 3 Then
        For i = 1 To gcolPlugIn.Count
            Set objTmp = gcolPlugIn.Item(i)
            Call objTmp.ExecuteFunc(lngSys, lngModual, strFuncName, lngPatiID, varRecId, varKeyId, strReserve, int场合)
        Next i
    ElseIf bytType = 4 Then
        For i = 1 To gcolPlugIn.Count
            Set objTmp = gcolPlugIn.Item(i)
            Call objTmp.Terminate(lngSys, lngModual, int场合)
        Next i
    End If
    Err.Clear: On Error GoTo 0
End Function

Public Function GetFormCaptionEx(ByVal lngSys As Long, ByVal lngModual As Long) As String
'获取扩展部件中的卡片名称，要求每个扩展部件以及主部件之间卡片不能重名
    Dim i As Integer
    Dim objTmp As Object
    Dim strTmp As String
    Dim strCaption As String
    
    If gcolPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    For i = 1 To gcolPlugIn.Count
        Set objTmp = gcolPlugIn.Item(i)
        strTmp = ""
        strTmp = objTmp.GetFormCaption(lngSys, lngModual)
        strCaption = IIf(strCaption = "", "", strCaption & ",") & strTmp
    Next i
    GetFormCaptionEx = strCaption
    Err.Clear: On Error GoTo 0
End Function

Public Function GetFormEx(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'获取扩展部件中的卡片对象。因为卡片名称不重名，则每次调用只会返回一个对象，或者不返回
    Dim i As Integer
    Dim objTmp As Object
    Dim objForm As Object
    Dim strTmp As String
    Dim strCaption As String
    
    If gcolPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    For i = 1 To gcolPlugIn.Count
        Set objTmp = gcolPlugIn.Item(i)
        Set objForm = objTmp.GetForm(lngSys, lngModual, strName)
        If Not objForm Is Nothing Then Exit For
    Next i
    Err.Clear: On Error GoTo 0
    
    Set GetFormEx = objForm
End Function

Private Function CkeckUseable(ByVal str单位名称 As String) As Boolean
'功能：扩展插件使用限制示例代码
'参数：使用单位的全名
    Dim strTmp As String
    
    strTmp = zlRegInfo("单位名称", , 0)
    If strTmp = "" Then Exit Function
    If InStr("," & strTmp & ",", "," & str单位名称 & ",") > 0 Then CkeckUseable = True
End Function



'----------------------------------------------------------------------------------------------------------------------------
'记录集操作相关
'----------------------------------------------------------------------------------------------------------------------------
Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名|值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID|5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    Dim strValues As String
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strValues = ""
            intCols = .Fields.Count - 1
            For intCol = 0 To intCols
                strValues = strValues & "," & .Fields(intCol).Name & ":" & .Fields(intCol).Value
            Next
            Debug.Print strValues
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub
'----------------------------------------------------------------------------------------------------------------------------

Public Function MecFlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'支持外挂附页窗体滚轮的滚动，35以上版本启用
    On Error GoTo errH
    If Not gobjMec Is Nothing Then
        Call gobjMec.PlugWndProc(wMsg, wParam, lParam, 0)
    End If
    MecFlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFlexCol(ByVal objFlex As Object, ByVal strCaption As String) As Long
'功能：根据列头显示文字获取vsFlexGrid的列号
'返回：无对应的列时，返回-1
    Dim i As Long
    
    GetFlexCol = -1
    
    For i = 1 To objFlex.Cols - 1
        If UCase(objFlex.TextMatrix(0, i)) = UCase(strCaption) Then
            GetFlexCol = i: Exit Function
        End If
    Next
End Function
