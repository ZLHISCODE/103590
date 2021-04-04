Attribute VB_Name = "mdlPubFunc"
Option Explicit
'包含函数类型枚举
'====剪贴板操作=============
'clipClear：清空粘贴板
'clipCopyFiles:文件放入粘贴板
'====进程操作==============
'fun_KillProcess:杀掉指定名称进程
'====自定义存储过程==========
'GetBlankProcedure:获取存储过程参数默认值
'IsSpaceProcedure:存储过程名是否被占用
'====静态记录集操作==========
'CopyNewRec:直接生成或者从原始记录集生成一个本地的静态记录集（可以修改记录集的内容）
'RecDelete:删除满足指定条件的静态记录集的数据行
'RecUpdate:跟新满足指定条件的静态记录集的某些字段
'RecDataAppend:将一个静态记录集的数据附加到另一个静态记录集上
'====其他公共函数===========
'ActualLen: 获取字符串的字节长度
'ActualStr:字符串截取指定字节长度
'CancelNetServer:断开服务器连接
'Decode:模拟Oracle的Decode函数
'IsNetServer:服务器连接是否正常
'OpenFolder:选择文件夹
'SetCtrlPosOnLine:设置一组控件的对齐方式以及控件间距
'CboSetWidth:设置cbo控件下拉列表宽度
'GetControlRect获取控件在屏幕中的位置
'CboSetIndex：为一个Combo控件选择列表项，但又不触发其Click事件
'GetClientPoint：获取当前指针对应在控件中的位置

'压缩解压常量
Public Const PROAPPCTION = "7z.exe" '执行程序
Public Const COMPRESSIONRATE = 5 '标准压缩
'''压缩等级 压缩算法 字典大小 快速字节 匹配器 过滤器 描述
'''0 Copy 无压缩
'''1 LZMA 64 KB 32 HC4 BCJ 最快压缩
'''3 LZMA 1 MB 32 HC4 BCJ 快速压缩
'''5 LZMA 16 MB 32 BT4 BCJ 正常压缩
'''7 LZMA 32 MB 64 BT4 BCJ 最大压缩
'''9 LZMA 64 MB 64 BT4 BCJ2 极限压缩

'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'切换到指定的输入法。
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal Flags As Long) As Long
Private Const NORMAL_PRIORITY_CLASS             As Long = &H20&
Private Const STARTF_USESTDHANDLES              As Long = &H100&
Private Const STARTF_USESHOWWINDOW              As Long = &H1
Private Const SW_HIDE                           As Integer = 0 '隐藏窗口，激活另一个窗口
Public Const INFINITE                           As Long = &HFFFF&

Private mrsProgFuncs As ADODB.Recordset   '用于缓存一个模块所拥有的授权功能

'1-变动过程;2-空白过程;3-用户过程
Public Enum ProcType
    变动过程 = 1
    空白过程 = 2
    用户过程 = 3
End Enum

Public Enum ProcState
    待检查 = 0
    待调整 = 1
    调整中 = 2
    已调整 = 3
    无变化 = 4
End Enum

'1-上次自定过程;2-上次标准过程;3-本次自定过程;4-本次标准过程
Public Enum ProcTextType
    上次自定过程 = 1
    上次标准过程 = 2
    本次自定过程 = 3
    本次标准过程 = 4
End Enum
'自定义存储过程管理
Public Enum Color
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    公共模块色 = &HC00000
    默认前景色 = &H80000008
    锁色 = &HF5F5F5
    启用色 = 0
    停用色 = 255
End Enum

Public Type AbortInfo
    AbortSys As Long
    AbortFile As String
    AbortLine As Long
    AbortInfo As String
    IsHistory As Boolean
End Type

Private mlngPid As Long '进程操作
Public gHwnd As Long '进程操作
Public gstrSplite As String

Public Enum LogType
    LT_安装 = 0
    LT_常规升迁 = 1
    LT_提前升迁 = 2
    LT_历史库升迁 = 3
    LT_系统控制 = 4 '升迁界面系统控制，以及参数导入
    LT_自定义 = -1 '自定义文件夹以及文件名
    LT_跟踪日志 = 999
End Enum
Public gobjRIS As Object '新网接口对象
Public gstrSTOwner As String '标准版100所有者
Public gblnRIS As Boolean '可以创建RIS接口
Public gblnMustRIS As Boolean

'OpenFolder初始路径设置
Public gstrAPIPath As String
Public Declare Function GetTickCount Lib "kernel32" () As Long  '获取当前时间

Public Function clipClear() As Boolean
'清空当前剪贴板
    Call EmptyClipboard
End Function

Public Function clipCopyFiles(File() As String) As Boolean
'       模块：剪贴板操作
'       功能：剪贴板操作,复制文件目录到剪贴板
'       编写：祝庆
'       日期：2011年1月3日
'复制多个文件到剪贴板
   On Error Resume Next
   Dim strData As String
   Dim df As DROPFILES
   Dim hGlobal As Long
   Dim lpGlobal As Long
   Dim i As Long
   strData = ""

   
   '清除剪贴版中现存的数据
   If OpenClipboard(0&) Then
        '清空当前剪贴板
        Call EmptyClipboard
        
        '判断文件数组是否为空
        If SafeArrayGetDim(File) = 0 Then Exit Function
        For i = LBound(File) To UBound(File)
            strData = strData & File(i) & vbNullChar
        Next
        
        hGlobal = GlobalAlloc(GHND, Len(df) + LenB(strData))
        
        If hGlobal Then
            lpGlobal = GlobalLock(hGlobal)
         
            df.pFiles = Len(df)
            Call CopyMemory(ByVal lpGlobal, df, Len(df))
            Call CopyMemory(ByVal (lpGlobal + Len(df)), ByVal strData, LenB(strData))
   
            Call GlobalUnlock(hGlobal)
         
            If SetClipboardData(CF_HDROP, hGlobal) Then
                clipCopyFiles = True
            End If

        End If
        
        Call CloseClipboard
    End If
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
'=功能： 通过PID枚举所属的句柄,查找需要的窗口
    Dim Pid1 As Long
    Dim wText As String * 255
    GetWindowThreadProcessId hwnd, Pid1
    If mlngPid = Pid1 Then
        GetWindowText hwnd, wText, 100
        If InStrRev(wText, "%", -1) > 0 Then
            gHwnd = hwnd
        End If
    End If
    EnumWindowsProc = True
End Function

Private Sub Find_Window(ByVal lngPid As Long)
'       模块：进程句柄操作
'       功能：进程句柄操作获得指定进程的Hwnd
'       编写：祝庆
'       日期：2010年11月24日
    mlngPid = lngPid
    gHwnd = 0
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub

'查找进程的函数
Public Sub fun_KillProcess(ByVal ProcessName As String)
    Dim strData As String
    Dim my As PROCESSENTRY32
    Dim l As Long
    Dim l1 As Long
    Dim mName As String
    Dim i As Integer, Pid As Long
    Dim mProcID As Long
    l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If l Then
        my.dwSize = 1060
        If (Process32First(l, my)) Then
            Do
                i = InStr(1, my.szExeFile, Chr(0))
                mName = LCase(Left(my.szExeFile, i - 1))
                If mName = LCase(ProcessName) Then
                    Pid = my.th32ProcessID
                    mProcID = OpenProcess(1&, -1&, Pid)

                    TerminateProcess mProcID, 0&
                End If
            Loop Until (Process32Next(l, my) < 1)
        End If
        l1 = CloseHandle(l)
    End If
End Sub


Public Function IsSpaceProcedure(ByVal strOwner As String, ByVal strProcName As String) As Boolean
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "Select 1 From zlProcedure Where 名称=[1] And 类型=2"
    Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSql, "", UCase(strProcName))
    IsSpaceProcedure = (rsData.BOF = False)
    
End Function

Public Function GetBlankProcedure(ByVal strProc As String) As String
    Dim lngCount As Long
'    Dim blnTitleFlag As Boolean
'    Dim strEnd As String
    Dim strSql As String
'    Dim lngInstr As Long
    Dim strArr() As String
    
    Dim strLine As String
    Dim lngPostion As Long
    
    Dim strReturnType As String
    
    strArr = Split(strProc, vbCrLf)
    strSql = ""
    strReturnType = ""
    
    For lngCount = 0 To UBound(strArr)
        
        strLine = Replace(Trim(strArr(lngCount)), Chr(10), "")
        strLine = UCase(Replace(strLine, Chr(13), ""))

        '取掉--注释
        lngPostion = InStr(strLine, "--")
        If lngPostion > 0 Then strLine = Mid(strLine, 1, lngPostion - 1)
        
        lngPostion = InStr(strLine, "RETURN ")
        If lngPostion > 0 Then
            
            If InStr(strLine, " NUMBER") > 0 Then
                strReturnType = "NUMBER"
            ElseIf InStr(strLine, " VARCHAR") > 0 Then
                strReturnType = "VARCHAR"
            ElseIf InStr(strLine, " DATE") > 0 Then
                strReturnType = "DATE"
            End If
        End If
        
        Select Case strLine
        Case "AS", "IS"
            strSql = strSql & strArr(lngCount) & vbCrLf
            Exit For
        Case Else
            If Right(strLine, 3) = " AS" Then
                strSql = strSql & strArr(lngCount) & vbCrLf
                Exit For
            ElseIf Right(strLine, 3) = " IS" Then
                strSql = strSql & strArr(lngCount) & vbCrLf
                Exit For
            Else
                strSql = strSql & strArr(lngCount) & vbCrLf
            End If
        
        End Select
    Next
    strSql = strSql & "Begin" & vbCrLf
    strSql = strSql & " " & vbCrLf
    Select Case strReturnType
    Case "NUMBER"
        strSql = strSql & vbTab & "return 0;" & vbCrLf
    Case "VARCHAR"
        strSql = strSql & vbTab & "return '';" & vbCrLf
    Case "DATE"
        strSql = strSql & vbTab & "return sysdate;" & vbCrLf
    End Select
    strSql = strSql & "End;"
    GetBlankProcedure = strSql
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'功能：取指定字符串按字节算的长度
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function ActualStr(ByVal strAsk As String, ByVal lngLen As Long) As String
'功能：取指定字符串左边指定字节长度的内容
    Dim strTemp As String, i As Long
    
    strTemp = StrConv(LeftB(StrConv(strAsk, vbFromUnicode), lngLen), vbUnicode)
    If InStr(strTemp, Chr(0)) > 0 Then
        strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
    End If
    ActualStr = strTemp
End Function

Public Function HScrollVisible(vsInput As Object) As Boolean
'判断水平滚动条的可见性
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    HScrollVisible = False
    i = GetScrollRange(vsInput.hwnd, SB_HORZ, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos And Not (lpMaxPos = 100 And lpMinPos = 0) Then HScrollVisible = True
End Function

Public Function VScrollVisible(vsInput As Object) As Boolean
'判断垂直滚动条的可见性
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    VScrollVisible = False
    i = GetScrollRange(vsInput.hwnd, SB_VERT, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos And Not (lpMaxPos = 100 And lpMinPos = 0) Then VScrollVisible = True
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
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

Public Function OpenFolder(ByVal frmodtvOwner As Form, Optional strTitle As String, Optional ByVal strInitDir As String) As String
'    '----------------------------------------------------------------------------------------------------
'    '功能:选择文件夹
'    '参数:frmodtvOwner-选择文件夹的父窗体
'    '       strFolderName-指定的文件夹
'    '       strTitle-标题
'    '       strInitDir-默认打开路径
'    '返回:strFolderName-返回选择的文件夹
'    '----------------------------------------------------------------------------------------------------
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    gstrAPIPath = strInitDir & Chr(0)
    With tBrowseInfo
        .hwndOwner = frmodtvOwner.hwnd
        .lpszTitle = lstrcat(strTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        .lpfnCallback = AddressOfFunction(AddressOf OpenDirCallbackProc)
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH * 2)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
       OpenFolder = sBuffer
    End If
End Function

Public Function OpenDirCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
 '功能：OpenFolder回调函数，用来设置打开的文件的初始路径
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
  
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hwnd, BFFM_SETSELECTION, 1, ByVal gstrAPIPath)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH * 2)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
            End If
    End Select
    
    OpenDirCallbackProc = 0
End Function

Private Function AddressOfFunction(Address As Long) As Long
'功能：OpenFolder子函数
    AddressOfFunction = Address
End Function

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'编制人:朱玉宝
'修改人：刘硕
'修改日期：2014-1-6
'修改点：增加复制记录集的部分字段功能
'编制日期:2000-11-02
'复制记录集
'参数：strFields=需要复制的记录集的字段的列顺序或字段名组成的字符串
'          如：1 别名1,3 别名2,7 别名3...表示复制记录集的第1,3,7..字段组成记录集并返回
'              ID 别名1,姓名 别名2,....表示复制记录集的ID,姓名...字段组成记录集返回
'              别名*为新的记录集的列名
'              两中类型混搭容易出现列名相同的问题，请注意
'               *,在表示复制原记录集的所有字段，可能需要将原来的字段重新产生新列
'           arrAppFields=追加的字段信息：列名,类型,长度,默认值,没有默认值传Empty,没有指定长度传Empty
'      blnOnlyStructure=是否只复制结构
'在程序中，经常会涉及到相互传递记录集，而使用ADO的Clone复制产生的记录集，当其中一个记录集的数据发生变化的时候，所有副本都将发生相同的变化（通常指修改或删除），而我们往往希望这些记录集相互间保持独立
  
    Dim rsClone As ADODB.Recordset
    Dim rsTarget As ADODB.Recordset
    Dim intFields As Integer, blnALlFileds As Boolean
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant, arrFieldsTmp As Variant
    Dim i As Long
    
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '产生记录集结构
        If strFields = "" Then
            strFields = "*"
        End If
        arrFieldsTmp = Split(strFields, ",")
        arrFieldsName = Array()
        For intFields = LBound(arrFieldsTmp) To UBound(arrFieldsTmp)
            If Trim(arrFieldsTmp(intFields)) = "*" Then '标识此处将增加原记录集的所有列
                If Not rsClone Is Nothing Then
                    For i = 0 To rsClone.Fields.Count - 1
                        ReDim Preserve arrFieldsName(UBound(arrFieldsName) + 1)
                        arrFieldsName(UBound(arrFieldsName)) = rsClone.Fields(i).name & ""
                        .Fields.Append rsClone.Fields(i).name, IIf(rsClone.Fields(i).Type = adNumeric, adDouble, rsClone.Fields(i).Type), rsClone.Fields(i).DefinedSize, adFldIsNullable    '0:表示新增
                    Next
                End If
            Else
                ReDim Preserve arrFieldsName(UBound(arrFieldsName) + 1)
                '列包含别名
                arrTmp = Split(arrFieldsTmp(intFields) & " ", " ")
                strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).name & ""
                '获取字段原名，存入数组
                arrFieldsName(UBound(arrFieldsName)) = strFieldName
                '添加字段,若果存在别名，则新增列的列名为别名
                .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:表示新增
            End If
        Next
        
        '追加字段添加
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '复制数据
        If Not blnOnlyStructure Then
            If rsClone Is Nothing Then Set CopyNewRec = rsTarget: Exit Function
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '新记录集的列按顺序添加，因此可以这样
                    .Fields(intFields).value = rsClone.Fields(arrFieldsName(intFields)).value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Function RecDelete(ByRef rsInput As ADODB.Recordset, Optional ByVal strFilter As String) As Boolean
'功能：删除指定条件的记录集的记录
'参数：rsInput=记录集
'      strFilter=条件
'返回：是否成功
'      rsInput=经过删除后的记录集
    rsInput.Filter = strFilter
    If rsInput.RecordCount > 0 Then
        rsInput.MoveFirst
        Do While Not rsInput.EOF
            Call rsInput.Delete
            rsInput.MoveNext
        Loop
        Call rsInput.UpdateBatch
    End If
    RecDelete = True
End Function

Public Function RecUpdate(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
'功能：更新指定条件的记录集的记录
'参数：rsInput=记录集
'      strFilter=条件
'      arrInput=输入的字段名以及值，格式：字段名1,值1, 字段名2,值2,....
'返回：是否成功
'      rsInput=经过更新后的记录集
'说明：arrInput的字段值可以用记录集中的其他字段来更新该字段，此时格式为：!字段名 处理函数(暂时支持Val)
    Dim strFiledName As String, strFileValue As String, strFun As String, strFindFiled As String
    Dim blnFiled As Boolean, i As Long
    Dim arrTmp As Variant
    
    If rsInput Is Nothing Then Exit Function
    On Error GoTo errH
    With rsInput
        .Filter = strFilter
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            For i = LBound(arrInput) To UBound(arrInput) Step 2
                strFiledName = arrInput(i)
                If arrInput(i + 1) & "" = "" Then
                    rsInput(strFiledName).value = Null
                Else
                    strFun = ""
                    strFindFiled = arrInput(i + 1)
                    If arrInput(i + 1) Like "!?*" Then
                        blnFiled = True
                        On Error Resume Next
                        strFindFiled = Mid(arrInput(i + 1), 2)
                        arrTmp = Split(strFindFiled & " ", " ")
                        strFindFiled = Trim(arrTmp(0))
                        strFun = Trim(arrTmp(1))
                        strFileValue = rsInput(strFindFiled).value & ""
                        If err.Number <> 0 Then err.Clear: blnFiled = False
                        On Error GoTo errH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).value = arrInput(i + 1)
                    Else
                        If strFun = "" Then
                            rsInput(strFiledName).value = rsInput(strFindFiled).value
                        ElseIf strFun = "Val" Then
                            rsInput(strFiledName).value = Val(rsInput(strFindFiled).value & "")
                        ElseIf strFun = "Trim" Then
                            rsInput(strFiledName).value = Trim(rsInput(strFindFiled).value & "")
                            If rsInput(strFiledName).value & "" = "" Then
                                rsInput(strFiledName).value = Null
                            End If
                        Else
                            rsInput(strFiledName).value = rsInput(strFindFiled).value
                        End If
                    End If
                End If
                blnFiled = False
            Next
            .MoveNext
        Loop
        Call rsInput.UpdateBatch
    End With
    RecUpdate = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function RecDataAppend(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
'功能：将指定记录集的数据添加到另一个记录集上
'参数：rsSource=目标记录集
'      rsAppend=数据记录集
'      arrInput=字段对应规则，该参数不传时，默认两记录集结构相同，格式：arrInput(0):[记录集1].字段1,字段2...；arrInput(1)：[记录集2].字段1,字段2...
'返回：是否成功
'      rsSource=添加数据后的记录集
    Dim arrSource As Variant, arrAppend As Variant
    Dim i As Long, arrValues() As Variant
    Dim strTmp As String
    
    If rsAppend Is Nothing Then RecDataAppend = True: Exit Function
    If rsAppend.RecordCount = 0 Then RecDataAppend = True: Exit Function
    If rsSource Is Nothing Then Set rsSource = rsAppend: RecDataAppend = True: Exit Function
    On Error GoTo errH
    If LBound(arrInput) = 2 Then
        '此段代码需要经过仔细测试
        arrSource = Split(arrInput(LBound(arrInput)), ",")
        arrAppend = Split(arrInput(UBound(arrInput)), ",")
        If UBound(arrSource) <> UBound(arrAppend) Then Exit Function
        ReDim arrValues(UBound(arrAppend)): rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            For i = LBound(arrAppend) To UBound(arrAppend)
                arrValues(i) = rsAppend(arrAppend(i)).value
            Next
            rsSource.AddNew arrSource, arrValues
            Erase arrValues
            rsAppend.MoveNext
        Loop
    ElseIf LBound(arrInput) = 0 Then
        strTmp = ""
        For i = 0 To rsSource.Fields.Count - 1
            strTmp = strTmp & "," & rsSource.Fields(i).name
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        arrSource = Split(strTmp, ",")
        On Error Resume Next
        If rsAppend.RecordCount <> 0 Then rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            rsSource.AddNew
            For i = LBound(arrSource) To UBound(arrSource)
                rsSource.Fields(arrSource(i)).value = rsAppend.Fields(arrSource(i)).value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
        If err.Number <> 0 Then err.Clear
        On Error GoTo errH
    End If
    
    RecDataAppend = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function RecDistinct(ByVal rsSource As ADODB.Recordset, Optional ByVal strDisFieldsName As String, Optional ByVal strFieldsName As String) As ADODB.Recordset
'功能：记录集去重复
'参数：rsSource=要去重复的记录集
'strDisFieldsName=去重复的字段,为空，则对所有字段去重
'strFieldsName=返回结果集字段，为空，则返回去重复的字段
'返回：操作后的记录集
    Dim rsReturn As ADODB.Recordset
    Dim arrFilds As Variant, arrValues As Variant
    Dim i As Long, j As Long
    Dim strTmp As String, strOldRow As String

    '读取默认字段名
    If strDisFieldsName = "" Then
        For i = 0 To rsSource.Fields.Count - 1
            strTmp = strTmp & "," & rsSource.Fields(i).name
        Next
        strTmp = Mid(strTmp, 2)
        If strDisFieldsName = "" Then strDisFieldsName = strTmp
    End If
    If strFieldsName = "" Then strFieldsName = strDisFieldsName
    
    Set rsReturn = CopyNewRec(rsSource, , strFieldsName)
    If rsSource.RecordCount = 0 Then Set RecDistinct = rsReturn: Exit Function
    
    rsReturn.Sort = strDisFieldsName '排序，自动将光标移动到开头
    Do While Not rsReturn.EOF
        strTmp = rsReturn.GetString(, 1, "[ColumnSpliter]", , "[NULLEXP]") '自动移动光标
        rsReturn.MovePrevious
        If strTmp = strOldRow Then  '删除重复行
            Call rsReturn.Delete: Call rsReturn.Update
        Else
            strOldRow = strTmp
        End If
        rsReturn.MoveNext
    Loop
    rsReturn.Sort = strDisFieldsName
    Set RecDistinct = rsReturn
End Function

Public Function GetALLPars(Optional ByVal lngSys As Long = -1, Optional ByVal blnDetails As Boolean = True, Optional ByVal blnAddSets As Boolean) As ADODB.Recordset
'获取所有参数
'参数：blnDetails=True ,获取部门参数设置详情，本机私有参数设置详情,false-只获取参数列表
'         lngSys=-1-获取所有系统，<>0获取某一个系统,=-9仅获取系统信息
'         blnAddSets=是否增加配置信息
' 返回：获取的参数记录集

    Dim strSql As String
    Dim rsParas As ADODB.Recordset
    If lngSys <> -9 Then
        '所有参数信息
        strSql = "Select 0 类型, Nvl(a.系统, 0) 系统," & vbNewLine & _
                        "       Nvl(a.模块, 0) 模块, Nvl(a.私有, 0) 私有, Nvl(a.本机, 0) 本机,Nvl(a.部门, 0) 部门, Nvl(a.性质, 0) 性质,  Nvl(a.授权, 0) 授权, Nvl(a.固定, 0) 固定, a.参数号, a.参数名, a.参数值, a.缺省值, a.影响控制说明, a.参数值含义, a.关联说明, a.适用说明," & vbNewLine & _
                        "       a.警告说明, 0 部门id, Null 用户名, Null 机器名," & vbNewLine & _
                        "       Null 详细参数值" & vbNewLine & _
                        "From zlParameters A" & IIf(lngSys = -1, "", " Where Nvl(a.系统, 0)=[1]")
        If blnDetails Then
            '部门参数详情
            strSql = strSql & vbNewLine & _
                            "Union All " & vbNewLine & _
                            "Select 1 类型, Nvl(a.系统, 0) 系统, Nvl(a.模块, 0) 模块, Null  私有,Null  本机, Null 部门, Null 性质,Null  授权, Null 固定, a.参数号, a.参数名, Null 参数值, Null 缺省值," & vbNewLine & _
                            "       Null 影响控制说明, Null 参数值含义, Null 关联说明, Null 适用说明,Null 警告说明," & vbNewLine & _
                            "       b.部门id 部门id, Null 用户名, Null 机器名, b.参数值  详细参数值" & vbNewLine & _
                            "From zlParameters A, Zldeptparas B" & vbNewLine & _
                            "Where a.Id = b.参数id And Nvl(a.部门, 0) = 1" & IIf(lngSys = -1, "", " And Nvl(a.系统, 0)=[1]")
            '私有本机参数详情
            strSql = strSql & vbNewLine & _
                            "Union All " & vbNewLine & _
                            "Select 2 类型, Nvl(a.系统, 0) 系统, Nvl(a.模块, 0) 模块, Null  私有,  Null 本机,  Null 部门,Null  性质,Null  授权, Null  固定,a.参数号, a.参数名, Null 参数值,Null 缺省值," & vbNewLine & _
                            "       Null 影响控制说明, Null 参数值含义, Null 关联说明, Null 适用说明, Null 警告说明," & vbNewLine & _
                            "       Null 部门id, c.用户名 用户名, c.机器名 机器名, c.参数值 详细参数值" & vbNewLine & _
                            "From zlParameters A, zlUserParas C" & vbNewLine & _
                            "Where a.Id = c.参数id And Nvl(a.部门, 0) = 0 And (Nvl(a.私有, 0) = 1 Or Nvl(a.本机, 0) = 1)" & IIf(lngSys = -1, "", " And Nvl(a.系统, 0)=[1]")
            
        End If
    End If
    '系统信息，导出信息、参数统计信息
    strSql = IIf(lngSys <> -9, strSql & vbNewLine & _
                    "Union All " & vbNewLine, "") & _
                    "Select -9 类型, Nvl(编号, 0) 系统, Null 模块, Null 私有, Null 本机,  Null 部门,  Null  性质, Null 授权,Null 固定,Null 参数号, 名称 参数名, 版本号 参数值, b.计数 || ''   缺省值," & vbNewLine & _
                    "       Null 影响控制说明, Null  参数值含义, Null 关联说明, Null 适用说明, Null 警告说明," & vbNewLine & _
                    "       Null 部门id, Null 用户名, Null 机器名, Null 详细参数值" & vbNewLine & _
                    "From (Select 编号, 名称, 版本号" & vbNewLine & _
                    "       From zlSystems" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select 0, '服务器管理工具', 内容" & vbNewLine & _
                    "       From zlRegInfo" & vbNewLine & _
                    "       Where 项目 = '版本号') A, (Select Nvl(系统, 0) 系统, Count(1) 计数 From zlParameters Group By Nvl(系统, 0)) B" & vbNewLine & _
                    "Where a.编号 = b.系统" & IIf(lngSys = -1 Or lngSys = -9, "", " And Nvl(a.编号, 0)=[1]")
    If blnAddSets And lngSys <> -9 Then
        '配置信息
        strSql = strSql & vbNewLine & _
                        "Union All " & vbNewLine & _
                        "Select -99 类型, Null 系统, Null 模块, " & IIf(blnDetails, 1, 0) & " 私有, Null 本机, Null 部门, Null  性质, Null 授权,Null 固定,  " & lngSys & " 参数号, Null 参数名, To_Char(Sysdate, 'yyyy-mm-dd HH24:mi:ss')  参数值, User 缺省值," & vbNewLine & _
                        "       Null 影响控制说明, Null 参数值含义, Null 关联说明, Null 适用说明, Null 警告说明," & vbNewLine & _
                        "       Null 部门id, Null 用户名, Null 机器名, Null 详细参数值" & vbNewLine & _
                        "From Dual"
    End If
    '嵌套并排序
    strSql = "Select D.系统||'#'||D.模块||'#'||D.参数名 MainKey,D.*" & vbNewLine & _
                    "From (" & strSql & ") D" & vbNewLine & _
                    "Order By 类型, 系统, 模块, 参数名"
    '增加排序关键字
    strSql = "Select RowNum SortKey,E.*" & vbNewLine & _
                "From (" & strSql & ") E" & vbNewLine & _
                "Order By 类型, 系统, 模块, 参数名"
    Set rsParas = gclsBase.OpenSQLRecord(gcnOracle, strSql, "获取所有参数", lngSys)
    Set GetALLPars = rsParas
End Function

Public Function GetCompareRec(ByVal rsSouce As ADODB.Recordset, ByVal rsCompare As ADODB.Recordset, ByVal strKeyFields As String, Optional ByVal strComPareFileds As String, Optional ByVal strAddtionFileds As String, Optional arrAppFields As Variant) As ADODB.Recordset
'功能：获取记录集比较结果记录集
'参数：rsSouce=比较记录集
'         rsCompare=对比记录集
'         strComPareFileds=进行对比的字段,字段名字之间以逗号分割，为空值表示以rsSouce的字段作为对比字段,"-字段串"，标识该字段串中字段不参与比较
'         strAddtionFileds=添加入比较记录集但是不进行比较判断，这些字段为了方便比较记录集的使用
'         strKeyFields：字段名间以逗号分割。主键字段,格式为:Nvl(字段1,0)_Nvl(字段2,0)_Nvl(字段3,0)...
'         arrAppFields=追加的字段信息：列名,类型,长度,默认值,没有默认值传Empty,没有指定长度传Empty
    Dim i As Long
    Dim strFileds As String
    Dim varKey As Variant, varCom As Variant, varAddtion As Variant
    Dim rsReturn As ADODB.Recordset, rsSort As ADODB.Recordset
    Dim strTmpKey As String, strPreKey As String
    Dim blnNew As Boolean, strDifCols As String
    Dim strNotCom As String
    Dim cllNumCol As Collection '数值列，这种列空值等于0
    Dim intState As Integer
    
    On Error GoTo errH
    If strKeyFields = "" Then Exit Function
    Set rsReturn = New ADODB.Recordset
    Set cllNumCol = New Collection
    
    With rsReturn
        .Fields.Append "MainKey", adVarChar, 200, adFldIsNullable
        .Fields.Append "State", adInteger '-1-删除，0-不变，1-新增，2-更新
        .Fields.Append "DifInfo", adVarChar, 2000, adFldIsNullable
        .Fields.Append "Sort", adInteger, Empty, Empty '更新排在删除与新增之间，-1删除，0-不变,1-更新,2-新增
        varKey = Split(strKeyFields, ",") '主键字段
        varCom = Split(strComPareFileds & "-", "-")
        strComPareFileds = UCase(Trim(varCom(0))) '比较字段
        strNotCom = UCase(Trim(varCom(1))) '不比较字段
        varAddtion = Split(strAddtionFileds, ",") '附加字段
        If strComPareFileds = "" Then
            For i = 0 To rsSouce.Fields.Count - 1
                '比较字段不包括主键,不比较字段与附加字段
                If InStr("," & strKeyFields & ",", "," & rsSouce.Fields(i).name & ",") = 0 And InStr("," & strNotCom & ",", "," & rsSouce.Fields(i).name & ",") = 0 And InStr("," & strAddtionFileds & ",", "," & rsSouce.Fields(i).name & ",") = 0 Then
                    strComPareFileds = strComPareFileds & IIf(strComPareFileds = "", "", ",") & rsSouce.Fields(i).name
                End If
            Next
        Else
            strComPareFileds = "," & strComPareFileds & ","
            If strNotCom <> "" Then
                varCom = Split(strNotCom, ",")
                For i = LBound(varCom) To UBound(varCom)
                    strComPareFileds = Replace(strComPareFileds, "," & varCom(i) & ",", ",")
                Next
            End If
            For i = LBound(varKey) To UBound(varKey)
                strComPareFileds = Replace(strComPareFileds, "," & varKey(i) & ",", ",")
            Next
            For i = LBound(varAddtion) To UBound(varAddtion)
                strComPareFileds = Replace(strComPareFileds, "," & varAddtion(i) & ",", ",")
            Next
            If strComPareFileds = "," Then
                strComPareFileds = ""
            Else
                strComPareFileds = Mid(strComPareFileds, 2, Len(strComPareFileds) - 2)
            End If
        End If
        If strComPareFileds = "" Then Exit Function '无比较字段，则不能进行比较
        varCom = Split(strComPareFileds, ",")
        For i = LBound(varCom) To UBound(varCom)
            If IsType(rsSouce.Fields(varCom(i)).Type, adNumeric) Then
                cllNumCol.Add 1, varCom(i)
            Else
                cllNumCol.Add 0, varCom(i)
            End If
            '原始字段
            .Fields.Append varCom(i), IIf(rsSouce.Fields(varCom(i)).Type = adNumeric, adDouble, rsSouce.Fields(varCom(i)).Type), rsSouce.Fields(varCom(i)).DefinedSize, adFldIsNullable
            '新字段
            .Fields.Append varCom(i) & "_New", IIf(rsSouce.Fields(varCom(i)).Type = adNumeric, adDouble, rsSouce.Fields(varCom(i)).Type), rsSouce.Fields(varCom(i)).DefinedSize, adFldIsNullable
        Next
        '数据源中字段，仅作为附件数据添加到记录集，不进行记录集比对
        For i = LBound(varAddtion) To UBound(varAddtion)
            '原始字段
            .Fields.Append varAddtion(i), IIf(rsSouce.Fields(varAddtion(i)).Type = adNumeric, adDouble, rsSouce.Fields(varAddtion(i)).Type), rsSouce.Fields(varAddtion(i)).DefinedSize, adFldIsNullable
            '新字段
            .Fields.Append varAddtion(i) & "_New", IIf(rsSouce.Fields(varAddtion(i)).Type = adNumeric, adDouble, rsSouce.Fields(varAddtion(i)).Type), rsSouce.Fields(varAddtion(i)).DefinedSize, adFldIsNullable
        Next
        '追加字段添加
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '排序复制,可能会  导致内存不足，所以慎用多字段排序
'        rsSouce.Sort = strKeyFields
'        rsCompare.Sort = strKeyFields
        Set rsSort = CopyNewRec(Nothing, , , Array("MainKey", adVarChar, 200, Empty, "类型", adInteger, 1, 0, "BookMark", adDouble, Empty, Empty))
        '生成主键
        If rsSouce.RecordCount <> 0 Then rsSouce.MoveFirst
        Do While Not rsSouce.EOF
            strTmpKey = ""
            For i = LBound(varKey) To UBound(varKey)
                strTmpKey = strTmpKey & IIf(strTmpKey = "", "", "#") & Nvl(rsSouce.Fields(varKey(i)).value, 0)
            Next
            rsSort.AddNew Array("MainKey", "类型", "BookMark"), Array(strTmpKey, 0, rsSouce.Bookmark)
            rsSouce.MoveNext
        Loop
        If rsSouce.RecordCount <> 0 Then rsSouce.MoveFirst
        '生成主键
        If rsCompare.RecordCount <> 0 Then rsCompare.MoveFirst
        Do While Not rsCompare.EOF
            strTmpKey = ""
            For i = LBound(varKey) To UBound(varKey)
                strTmpKey = strTmpKey & IIf(strTmpKey = "", "", "#") & Nvl(rsCompare.Fields(varKey(i)).value, 0)
            Next
            rsSort.AddNew Array("MainKey", "类型", "BookMark"), Array(strTmpKey, 1, rsCompare.Bookmark)
            rsCompare.MoveNext
        Loop
        If rsCompare.RecordCount <> 0 Then rsCompare.MoveFirst
        rsSort.Sort = "MainKey"
        Do While Not rsSort.EOF
            strTmpKey = rsSort!MainKey
            blnNew = rsSort!类型 = 1
            If blnNew Then
                rsCompare.Bookmark = CDbl(rsSort!Bookmark)
            Else
                rsSouce.Bookmark = CDbl(rsSort!Bookmark)
            End If
            If strPreKey <> strTmpKey Then
                .AddNew Array("MainKey", "State", "Sort"), Array(strTmpKey, 0, 0) '主键变化，则新增一行
                strPreKey = strTmpKey
            End If
            intState = Val(!State) + IIf(blnNew, 1, -1)
            .Update Array("State", "Sort"), Array(intState, IIf(intState = 1, 2, intState)) '用来区分新增与删除。新旧两个都有的后续判断是改变
            On Error Resume Next
            '比较数据填充
            For i = LBound(varCom) To UBound(varCom)
                If blnNew Then
                    .Update varCom(i) & "_New", rsCompare.Fields(varCom(i)).value
                Else
                    .Update varCom(i), rsSouce.Fields(varCom(i)).value
                End If
            Next
            '附加数据填充
            For i = LBound(varAddtion) To UBound(varAddtion)
                If blnNew Then
                    .Update varAddtion(i) & "_New", rsCompare.Fields(varAddtion(i)).value
                Else
                    .Update varAddtion(i), rsSouce.Fields(varAddtion(i)).value
                End If
            Next
            If err.Number <> 0 Then err.Clear
            On Error GoTo errH
            rsSort.MoveNext
        Loop
        '比较细微差异
        .Filter = "State=0": .Sort = "MainKey"
        Do While Not .EOF
            strDifCols = ""
            For i = LBound(varCom) To UBound(varCom)
                If cllNumCol(varCom(i)) = 1 Then
                    If Val(.Fields(varCom(i) & "_New").value & "") <> Val(.Fields(varCom(i)).value & "") Then  '获取差异列
                        strDifCols = strDifCols & IIf(strDifCols = "", "", ",") & varCom(i)
                    End If
                Else
                    If .Fields(varCom(i) & "_New").value & "" <> .Fields(varCom(i)).value & "" Then '获取差异列
                        strDifCols = strDifCols & IIf(strDifCols = "", "", ",") & varCom(i)
                    End If
                End If
            Next
            If strDifCols <> "" Then
                .Update Array("State", "DifInfo", "Sort"), Array(2, strDifCols, 1)
            End If
            .MoveNext
        Loop
        .Filter = ""
    End With
    Set GetCompareRec = rsReturn
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
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

Public Function SQLAdjust(ByVal varInput As Variant) As String
'功能：将含有"'"符号的字符串调整为Oracle所能识别的字符常量,并将空串转化为Null
'说明：自动(必须)在两边加"'"界定符。

    Dim i As Long, strTmp As String, strOneChar As String
    Dim strReturn As String
    Dim lngLine As Long
    
    strReturn = varInput & ""
    If strReturn & "" = "" Then SQLAdjust = "Null": Exit Function
    If InStr(1, strReturn, "'") = 0 And InStr(1, strReturn, Chr(10)) = 0 And InStr(1, strReturn, Chr(13)) = 0 Then SQLAdjust = "'" & strReturn & "'": Exit Function
    
    For i = 1 To Len(strReturn)
        strOneChar = Mid(strReturn, i, 1)
        Select Case strOneChar
            Case "'"
                If i = 1 Then
                    strTmp = "CHR(39)||'"
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & "'||CHR(39)"
                Else
                    strTmp = strTmp & "'||CHR(39)||'"
                End If
                lngLine = lngLine + 1 '标识有非换行字符
            Case Chr(10), Chr(13)
                If i = 1 Then
                    strTmp = "CHR(13)||'"
                ElseIf lngLine = 0 Then '连着多个换行，保留一个
                    If i = Len(strReturn) Then '最后一个是换行
                        strTmp = strTmp & "'"
                    End If
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & "'||CHR(13)"
                Else
                    strTmp = strTmp & "'||CHR(13)||'"
                End If
                lngLine = 0 '标识已经有换行
            Case Else
                If i = 1 Then
                    strTmp = "'" & Mid(strReturn, i, 1)
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & Mid(strReturn, i, 1) & "'"
                Else
                    strTmp = strTmp & Mid(strReturn, i, 1)
                End If
                lngLine = lngLine + 1 '标识有非换行字符
        End Select
    Next
    SQLAdjust = strTmp
End Function

Public Sub SetCtrlPosOnLine(ByVal blnvertical As Boolean, ByVal intAligType As Integer, ParamArray arrControls() As Variant)
'功能:对同一行的控件进行位置设置
'参数：
'blnvertical  true ,垂直方向设置控件位置，false,水平方向设置控件位置
'blnvertical=false :intAligType=-1,顶端对齐，0-中间对齐，1-底端对齐,blnvertical=true,intAligType=-1,左对齐，0-水平中心对齐，1-右对齐
'   arrControls格式为控件1,间距1,控件2,间距2,控件3,...
    Dim i As Long
    Dim lngPos As Long '第一个控件的某一位置
    Dim dblRate As Double
    If UBound(arrControls) = -1 Then Exit Sub
    If blnvertical Then
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Left
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Left + 0.5 * arrControls(0).Width
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Left + arrControls(0).Width
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Top = arrControls(i - 2).Top + arrControls(i - 2).Height + arrControls(i - 1)
                arrControls(i).Left = lngPos - arrControls(i).Width * dblRate
            End If
        Next
    Else
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Top
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Top + 0.5 * arrControls(0).Height
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Top + arrControls(0).Height
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Left = arrControls(i - 2).Left + arrControls(i - 2).Width + arrControls(i - 1)
                arrControls(i).Top = lngPos - arrControls(i).Height * dblRate
            End If
        Next
    End If
End Sub

Public Sub SetCtrlSameDistance(ByVal blnvertical As Boolean, ByVal intSameType As Integer, ByVal intAligType As Integer, ParamArray arrControls() As Variant)
'功能:讲一组控件设置为相同的间距
'参数：
'blnvertical  true ,垂直方向设置控件位置，false,水平方向设置控件位置
'intAligType=2不设置
'blnvertical=false :intAligType=-1,顶端对齐，0-中间对齐，1-底端对齐,blnvertical=true,intAligType=-1,左对齐，0-水平中心对齐，1-右对齐
'intSameType=0:边界间距相同，1-中心间距相同,
'arrControls格式为控件1,控件2,控件3,...
'说明：以首位两个控件作为基准，自动设置中间控件间距
    Dim i As Long, lngSart As Long, lngEnd As Long
    Dim lngSum As Long, lngDistance As Long
    Dim lngPos As Long
    Dim dblSameRate As Double, dblAligRate As Double
    
    If UBound(arrControls) < 2 Then Exit Sub '低于三个控件不处理
    '获取计算率
    dblSameRate = IIf(intSameType = 1, 0.5, 1)
    dblAligRate = intAligType / 2
    '计算起始位置
    If blnvertical Then
        lngSart = arrControls(0).Top + dblSameRate * arrControls(0).Height
        lngEnd = arrControls(UBound(arrControls)).Top + (1 - dblSameRate) * arrControls(UBound(arrControls)).Height
    Else
        lngSart = arrControls(0).Left + dblSameRate * arrControls(0).Width
        lngEnd = arrControls(UBound(arrControls)).Left + (1 - dblSameRate) * arrControls(UBound(arrControls)).Width
    End If
    '获取需要剔除的无效间距
    If intSameType = 0 Then '控件间边界间距相同
        For i = 1 To UBound(arrControls) - 1
            lngSum = lngSum + IIf(blnvertical, arrControls(i).Height, arrControls(i).Width)
        Next
    Else
        lngSum = 0
    End If
    '获取对齐位置
    If intAligType <> 2 Then
        If blnvertical Then
            lngPos = arrControls(0).Left + (dblAligRate + 0.5) * arrControls(0).Width
        Else
            lngPos = arrControls(0).Top + (dblAligRate + 0.5) * arrControls(0).Height
        End If
    End If
    '获取平均间距
    lngDistance = (lngEnd - lngSart - lngSum) / UBound(arrControls)
    '设置控件位置
    For i = 1 To UBound(arrControls)
        If blnvertical Then
            arrControls(i).Top = lngSart + lngDistance - (1 - dblSameRate) * arrControls(i).Height
            lngSart = arrControls(i).Top + arrControls(i).Height - (1 - dblSameRate) * arrControls(i).Height
            If intAligType <> 2 Then arrControls(i).Left = lngPos - (dblAligRate + 0.5) * arrControls(0).Width
        Else
            arrControls(i).Left = lngSart + lngDistance - (1 - dblSameRate) * arrControls(i).Width
            lngSart = arrControls(i).Left + arrControls(i).Width - (1 - dblSameRate) * arrControls(i).Width
            If intAligType <> 2 Then arrControls(i).Top = lngPos - (dblAligRate + 0.5) * arrControls(0).Height
        End If
    Next
End Sub

Public Sub SetCtrlEnabled(ByVal blnEnabled As Boolean, ParamArray arrControls() As Variant)
'功能:对一批控件的Enabled属性进行设置
'参数：
'blnEnabled  true ,空间可用，false,空间不可用
'arrControls格式为控件1,控件2,控件3,...

    Dim i As Long
    
    For i = LBound(arrControls) To UBound(arrControls)
        arrControls(i).Enabled = blnEnabled
    Next
End Sub


Public Function CancelNetServer(ByVal strPath As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '功能:断开服务器连接
    '参数:
    '返回:断找成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------
    err = 0
    On Error Resume Next
    If WNetCancelConnection2(strPath, CONNECT_UPDATE_PROFILE, True) = 0 Then
        CancelNetServer = True
    Else
        CancelNetServer = False
    End If
    err = 0
End Function

Public Function IsNetServer(ByVal strPath As String, ByVal strUser As String, ByVal strPassword As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '--功能:检查服务器是否正常并连接
    '--参数:strPath -访问路径
    '       strUser-用户名
    '       strPassWord -访问密码
    '返回:连接顺畅,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/06
    '----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
      
    '刘兴洪:可能存在windows资源管理器已经有访问的了
    '
'    If objFile.FolderExists(strPath) Then
'        IsNetServer = True: Exit Function
'    End If
    
    Dim NetR As NETRESOURCE
    With NetR
        .dwScope = RESOURCE_GLOBALNET
        .dwType = RESOURCETYPE_DISK
        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        .dwUsage = RESOURCEUSAGE_CONNECTABLE
        .lpLocalName = "" '映射的驱动器
        .lpRemoteName = strPath  '服务器路径
    End With
    
    err = 0
    On Error GoTo ErrHand:
    If WNetAddConnection2(NetR, strPassword, strUser, CONNECT_UPDATE_PROFILE) = NO_ERROR Then
       IsNetServer = True
    Else
       IsNetServer = False
    End If
    Exit Function
ErrHand:
       IsNetServer = False
End Function

Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single = -1, Optional frmParent As Object, Optional blnPer As Boolean)
'功能：显示或隐藏等待或进度窗体(strInfo)
'参数:strInfo=等待或进度提示信息
'     sngPer=进度
    Static blnShow As Boolean
    
    If sngPer > 1 Then sngPer = 1
    
    If strInfo = "" Then
        frmFlash.avi.Close
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            If sngPer = -1 Then
                '显示等待
                frmFlash.avi.Open GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "") & "\" & "Findfile.avi"
                If err.Number <> 0 Then
                    err.Clear
                End If
                frmFlash.lbl.Caption = strInfo
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    err.Clear
                    frmFlash.Show , frmParent
                    If err.Number <> 0 Then
                        err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.avi.Play
                frmFlash.Refresh
            Else
                '显示进度
                frmFlash.avi.Visible = False
                frmFlash.picDo.Visible = True
                frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
                frmFlash.lbl.Left = frmFlash.picDo.Left
                frmFlash.lblPer.Top = frmFlash.lbl.Top
                frmFlash.lbl.Caption = strInfo
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If blnPer Then
                    If sngPer > 0 Then
                        frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                    Else
                        frmFlash.lblPer.Caption = ""
                    End If
                    frmFlash.lblPer.Visible = True
                End If
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    err.Clear
                    frmFlash.Show , frmParent
                    If err.Number <> 0 Then
                        err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.Refresh
            End If
            blnShow = True
        Else
            frmFlash.lbl.Caption = strInfo
            If sngPer >= 0 Then
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

Public Function GetLogPath(ByVal ltLogType As LogType, Optional ByVal strSysCodes As String, Optional ByVal strBakUser As String, Optional ByVal strFolder As String, Optional ByVal strName As String) As String
'功能：获取日志目录
'ltLogType=日志类型，0-系统安装，1-正常升迁日志，2-提前升迁日志，3-历史库单独升迁日志,4系统控制
'strSysCodes=所要操作的系统的系统编码,多个系统以逗号分割
'strBakUser=历史库单独升级时所需历史库名
'strFolder=自定类型日志的文件夹
'strName=自定义类型的日志名
'返回：日志文件名以及路径,strFileName需要返回，主要是由于
    Dim strFileName As String
    Dim arrTmp  As Variant, strSys As String
    Dim i As Long
    Dim strTime As String
    
    On Error GoTo errH
    If gblnInIDE Then
        strFolder = GetSetting("ZLSOFT", "公共全局", "程序路径")
        strFolder = "C:\Appsoft\Log"
    Else
        strFolder = App.Path & "\Log"
    End If
    If Not gobjFile.FolderExists(strFolder) Then
        Call gobjFile.CreateFolder(strFolder)
    End If
    strTime = Format(Now, "YYMMDDHHmm")
    Select Case ltLogType
        Case LT_系统控制
            strFolder = strFolder & "\系统控制"
            strFileName = Mid(strTime, 1, 6) & ".Log"
        Case LT_跟踪日志
            strFolder = strFolder & "\日志跟踪"
            strFileName = strTime & ".Log"
        Case LT_安装, LT_常规升迁, LT_历史库升迁, LT_提前升迁
            strFolder = strFolder & "\安装升迁"
            arrTmp = Split(strSysCodes, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                strSys = strSys & Format(Val(arrTmp(i)) \ 100, "00")
            Next
            strFileName = Mid(strTime, 1, 6) & "_" & strSys & Decode(ltLogType, LT_安装, "_Install", LT_常规升迁, "", LT_提前升迁, "_BEF", LT_历史库升迁, "_" & strBakUser) & "_" & Mid(strTime, 7, 4) & ".log"
        Case LT_自定义
            strFileName = strName & strTime & ".log"
    End Select
    If Not gobjFile.FolderExists(strFolder) Then
        Call gobjFile.CreateFolder(strFolder)
    End If
    GetLogPath = strFolder & "\" & strFileName
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Function Identity(ByRef lngCount As Long) As Long
'功能：模拟主键自增
'参数：lngCount=自增变量
    lngCount = lngCount + 1
    Identity = lngCount
End Function

Public Function GetOracleVersion(Optional ByVal blnGetVerNum As Boolean = True, Optional ByVal blnGetBigVer As Boolean) As Variant
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim arrTmp As Variant
    
    If gstrOracleVer = "" Then
        'CORE    10.2.0.3.0  Production
        strSql = "Select Banner From V$version Where Banner Like  'CORE%'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.Title)
        If rsTmp.RecordCount > 0 Then
            arrTmp = Split(TrimEx(rsTmp!Banner & ""), " ")
            If UBound(arrTmp) = 2 Then
                gstrOracleVer = arrTmp(1)
            End If
        End If
    End If
    
    If gstrOracleVer <> "" Then
        If Not blnGetVerNum Then
            GetOracleVersion = gstrOracleVer
        Else
            If blnGetBigVer Then
                arrTmp = Split(gstrOracleVer, ".")
                GetOracleVersion = Val(arrTmp(0))
            Else
                GetOracleVersion = Val(Replace(Mid(gstrOracleVer, 4), ".", ""))
            End If
        End If
    Else
        GetOracleVersion = IIf(blnGetVerNum, 0, "获取失败")
    End If
End Function

Public Function ReadFileToString(ByVal strFile As String) As String
    Dim strBuffer As String
    Dim lngHwnd As Long
    Dim lngFileLen As Long

    lngHwnd = FreeFile

    On Error Resume Next
    Open strFile For Binary Shared As lngHwnd
    If err.Number <> 0 Then
        MsgBox "Error " & err.Number & vbCrLf & err.Description & vbCrLf & "Error in ReadFileToString, File='" & strFile & "'", vbCritical
        GoTo Proc_Exit
    End If
    On Error GoTo 0
    
    lngFileLen = LOF(lngHwnd)
    strBuffer = Space(lngFileLen)
    Get lngHwnd, , strBuffer
    
    Close lngHwnd
    
Proc_Exit:
    ReadFileToString = strBuffer
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub CboSetWidth(ByVal hWnd_combo As Long, ByVal lngWidth As Long)
'功能：设置Combo控件下拉列表的宽度
'此处的宽度是批下拉列表的宽度，并且是以TWIP为单位
    Const CB_SETDROPPEDWIDTH As Long = &H160

    SendMessage hWnd_combo, CB_SETDROPPEDWIDTH, lngWidth / Screen.TwipsPerPixelX, 0
End Sub

Public Sub CboSetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'功能：设置Combo控件的Index值
'为一个Combo控件选择列表项，但又不触发其Click事件
    Const CB_SETCURSEL = &H14E
    
    SendMessage hWnd_combo, CB_SETCURSEL, lngIndex, 0
End Sub

Public Sub WriteTraceLog(Optional ByVal strLog As String)
    If Not gblnTrace Then Exit Sub
    gobjLog.WriteLine strLog
End Sub

Public Function RestoreVsGridWidth(ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String) As Boolean
    '------------------------------------------------------------------------------
    '功能:从数据库中恢复网格的宽度等信息
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '     blnSaveToDataBase-是否是往数据库中保存参数(如果是往数据库中保存,则强制保存为true,否则根据是否使用个性化风格来确定)
    '     bln强制恢复保存-决定是否将保存注册表的参数值,进行强制恢复
    '返回:恢复成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    strParaValue = Trim((GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & strCaption, strKey)))
    If strParaValue = "" Then Exit Function
    RestoreVsGridWidth = False
    
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    err = 0: On Error GoTo ErrHand:
    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrTemp(1))
                If Val(arrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    RestoreVsGridWidth = True
    Exit Function
ErrHand:
End Function

Public Function SaveVsGridWidth(ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String) As Boolean
    '------------------------------------------------------------------------------
    '功能:保存vsFlex的宽度到注册表
    '参数:vsGrid-对应的网络控件
    '     strKey-主建
    '返回:保存成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & strCaption, strKey, strCol)
    SaveVsGridWidth = True
End Function

Public Function GetClientPoint(ByVal lngHwnd As Long) As POINTAPI
'获取当前指针对应在控件中的位置
    Dim pRet As POINTAPI
    Dim lngReturn As Long
    
    pRet = GetCursorPosition()
    lngReturn = ScreenToClient(lngHwnd, pRet)
    pRet.x = pRet.x * Screen.TwipsPerPixelX
    pRet.y = pRet.y * Screen.TwipsPerPixelY
    GetClientPoint = pRet
End Function

Public Function GetCursorPosition() As POINTAPI
'获取鼠标位置
    Dim pRet As POINTAPI
    Dim lngReturn As Long
    lngReturn = GetCursorPos(pRet)
    GetCursorPosition = pRet
End Function
Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, LngStyle As Long
    
    Call GetWindowRect(objForm.hwnd, vRect)
    LngStyle = GetWindowLong(objForm.hwnd, GWL_STYLE)
    If blnCaption Then
        LngStyle = LngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then LngStyle = LngStyle Or WS_SYSMENU
        If objForm.MaxButton Then LngStyle = LngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then LngStyle = LngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            LngStyle = LngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            LngStyle = LngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hwnd, GWL_STYLE, LngStyle
    SetWindowPos objForm.hwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function CompareFolder(ByVal strPath1 As String, ByVal strPath2 As String, ByVal strReports As String) As Boolean
'功能：对比两文件夹中文件，对相同文件名文件进行差异对比，有差异则生成报告。
    Dim strCommand As String
    Dim lngProcess As Long
    Dim lngTemp As Long
    
    err.Clear
    strCommand = GetWinSystemPath & "\wincmp3.exe " & strPath1 & "\" & " " & strPath2 & "\" & " /G:HNISE " & strReports
    lngTemp = Shell(strCommand, vbHide)
    DoEvents
    If err <> 0 Then
        err.Clear
         MsgBox "文件比较失败，请检查" & GetWinSystemPath & "\wincmp3.exe文件是否存在", vbExclamation, "中联软件"
        Exit Function
    End If
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
    Loop While lngTemp = Still_Active
    CompareFolder = True
    err.Clear
    DoEvents
End Function

Public Function CollectionHave(ByVal Coll As Collection, ByVal strKey As String) As Boolean
    On Error GoTo ErrHand
    
    Dim Item As Variant
    Set Item = Coll.Item(strKey)
    CollectionHave = True
    Set Item = Nothing
    Exit Function
ErrHand:
    '不存在返回False
    If err.Number = 5 Then CollectionHave = False
    err.Clear
End Function

Public Function GetRIS() As Boolean
'创建新网接口
    If Not gblnCreate Then Exit Function
    If Not gobjRIS Is Nothing Then GetRIS = True: Exit Function
    On Error Resume Next
    Set gobjRIS = CreateObject("zl9XWInterface.clsSvrTools")
    err.Clear: On Error GoTo 0
    GetRIS = Not gobjRIS Is Nothing
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function


Public Function OpenIme(Optional blnOpen As Boolean = False, Optional strImeName As String) As Boolean
'功能:打开中文输入法，或关闭输入法
'参数：strImeName-打开指定的输入法
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    
 
    '用户没进行设置，就不处理
    If blnOpen Then
        If strImeName <> "" Then
            strIme = strImeName
        End If
        If strIme = "" Then Exit Function                  '要求打开输入法，但是又没有设置
    End If
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))

    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否指定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then
                        OpenIme = True
                        Exit Function
                    End If
                End If
            End If
        ElseIf blnOpen = False Then
            '不是中文输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnOpen = False Then
        '由于windows Vista系统的英文输入法用ImmIsIME测试出是1的输入法,因此,需要单独处理.
        '刘兴宏:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function


Public Function SetSQLTrace(ByVal strServerName As String, ByVal strUserName As String, ByRef cnOracle As ADODB.Connection) As String
'功能:调用100046事件启动SQL Trace功能
'返回:Trc文件名
    Dim strSql As String, strLevel As String, strFile As String
    Dim rsTmp As ADODB.Recordset
    
    strServerName = UCase(strServerName)
    
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSql = "alter session set timed_statistics=true"
        cnOracle.ExeCute strSql
        strSql = "alter session set max_dump_file_size='100M'"
        cnOracle.ExeCute strSql
        If err.Number <> 0 Then err.Clear
        
        '下面这一条语句在8.1.7及以后才支持
        strFile = "ZL_" & strUserName
        strSql = "alter session set tracefile_identifier='" & strFile & "'"
        cnOracle.ExeCute strSql
        If err.Number <> 0 Then strFile = "*.trc": err.Clear
        
        strLevel = "12"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSql = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        cnOracle.ExeCute strSql
        If err.Number = 0 Then
            SetSQLTrace = strFile
            
            If CheckAndAdjustMustTable("ZLREGINFO", , True) Then    '先检测zlreginfo表是否存在
                strSql = "Select 1 From zlreginfo Where 项目='TRACE文件'"
                Set rsTmp = cnOracle.ExeCute(strSql)
                
                If rsTmp.RecordCount > 0 Then
                    strSql = "Update zlreginfo Set 内容 ='TRACE文件' Where 项目='" & strFile & ".trc'"
                Else
                    strSql = "Insert Into zlreginfo (项目,内容) Values ('TRACE文件','" & strFile & ".trc')"
                End If
                cnOracle.ExeCute strSql
            
                If err.Number <> 0 Then
                    MsgBox err.Description
                End If
            End If
        End If
    End If
End Function

Public Function RunCommand(ByVal strCommand As String, Optional ByRef strErr As String, Optional ByVal blnCiper As Boolean, Optional ByVal lngWait As Long = INFINITE) As String
'功能：执行命令行，并获取命令行输出
'新增逻辑:如果lngWait为0, 主程序不等待
    Dim piProc          As PROCESS_INFORMATION '进程信息
    Dim stStart         As STARTUPINFO '启动信息
    Dim saSecAttr       As SECURITY_ATTRIBUTES '安全属性
    Dim lnghReadPipe    As Long '读取管道句柄
    Dim lnghWritePipe   As Long '写入管道句柄
    Dim lngBytesRead    As Long '读出数据的字节数
    Dim strBuffer       As String * 256 '读取管道的字符串buffer
    Dim lngRet          As Long 'API函数返回值
    Dim lngRetPro       As Long
    Dim strlpOutputs    As String '读出的最终结果
    
    DoEvents
    On Error Resume Next
    '设置安全属性
    With saSecAttr
        .nLength = LenB(saSecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With
    
    '创建管道
    lngRet = CreatePipe(lnghReadPipe, lnghWritePipe, saSecAttr, 0)
    If lngRet = 0 Then
        strErr = "无法创建管道。" & GetLastDllErr()
        Exit Function
    End If
    '设置进程启动前的信息
    With stStart
        .Cb = LenB(stStart)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = lnghWritePipe '设置输出管道
        .hStdError = lnghWritePipe '设置错误管道
    End With
    '启动进程
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS进程以ipconfig.exe为例
    lngRetPro = CreateProcess(vbNullString, strCommand & vbNullChar, saSecAttr, saSecAttr, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, stStart, piProc)
    If lngRetPro = 0 Then
        strErr = "无法启动进程。" & GetLastDllErr()
        lngRet = CloseHandle(lnghWritePipe)
        lngRet = CloseHandle(lnghReadPipe)
        Exit Function
    Else
        '因为无需写入数据，所以先关闭写入管道。而且这里必须关闭此管道，否则将无法读取数据
        lngRet = CloseHandle(lnghWritePipe)
        WaitForSingleObject piProc.hProcess, lngWait
        Do
            If lngWait <> 0 Then
                lngRet = ReadFile(lnghReadPipe, strBuffer, 256, lngBytesRead, ByVal 0)
            End If
            If lngRet <> 0 Then
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            Else
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            End If
            DoEvents
        Loop While (lngRet <> 0) '当ret=0时说明ReadFile执行失败，已经没有数据可读了
        '读取操作完成，关闭各句柄
        lngRet = CloseHandle(lngRetPro)
        lngRet = CloseHandle(piProc.hProcess)
        lngRet = CloseHandle(piProc.hThread)
        lngRet = CloseHandle(lnghReadPipe)
    End If
    RunCommand = Replace(strlpOutputs, vbNullChar, "")
End Function

Public Function GetProgFuncs(ByVal strProg As String, Optional ByVal blnInitData As Boolean = False) As String
'功能：获取一个模块对应的功能权限或对模块权限字符串进行初始化操作
'参数：
'      strProg：模块号或模块权限字符串
'      blnInitData = True，此时strProg即为权限字符串，返回值为初始化后的权限字符串
'      blnInitData = False，此时strProg即为模块号，返回值为该模块拥有的功能权限
    Dim arrProg() As String
    Dim arrFunc() As String
    Dim strProgNo As String
    Dim i, j As Long
    
    If blnInitData Then
        '初始化权限字符串
    
        arrProg = Split(strProg, ",")
        Set mrsProgFuncs = New ADODB.Recordset
        
        '建立字段
        Call mrsProgFuncs.Fields.Append("模块号", adVarChar, 6)
        Call mrsProgFuncs.Fields.Append("功能名称", adVarChar, 30)

        '填充记录集
        mrsProgFuncs.Open
        With mrsProgFuncs
            For i = 0 To UBound(arrProg)
                strProgNo = Split(arrProg(i), ":")(0)
                arrFunc = Split(Split(arrProg(i) & ":", ":")(1), "|")
                If UBound(arrFunc) >= 0 Then
                    For j = 0 To UBound(arrFunc)
                        .AddNew
                        .Fields("模块号").value = strProgNo
                        .Fields("功能名称").value = arrFunc(j)
                    Next
                Else
                    .AddNew
                    .Fields("模块号").value = strProgNo
                    .Fields("功能名称").value = ""
                End If
                GetProgFuncs = GetProgFuncs & "," & strProgNo
            Next
        End With
    Else
        '根据模块号返回功能名称
        If Not mrsProgFuncs Is Nothing Then   'mrsProgFuncs为nothing说明并没有进行初始化操作，即当前登录用户拥有所有权限
            mrsProgFuncs.Filter = "模块号 = '" & strProg & "'"
            With mrsProgFuncs
                Do While Not .EOF
                    GetProgFuncs = GetProgFuncs & "|" & !功能名称
                    .MoveNext
                Loop
            End With
        Else
            GetProgFuncs = ""
        End If
    End If
    
    GetProgFuncs = Mid(GetProgFuncs, 2)
End Function
