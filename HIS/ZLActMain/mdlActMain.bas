Attribute VB_Name = "mdlActmain"
Option Explicit
#Const SYS_TRYUSE = "正式" '正式/试用

'注意：
'当类模块使用SingleUse时，每个对象实例将使用单独的进程，不同类的对象可以共享最近的进程。
'在单独的对象进程中，公共模块、窗体等的内容也是按进程单独存在的，在不同进程之间互不影响
'另外如果对象实例卸载了(Terminate)，但还存在未卸载的窗体，则该进程不会结束，进程中的公共内容此时仍可有效访问
Public gcnOracle As New ADODB.Connection
Public gstrSysName As String                '当前系统名称
Public gstrClass() As String
Public gobjClass() As Object
Public gobjRegister As Object               '注册授权部件zlRegister
Public gclsLogin As New clsRelogin
Public Enum Register
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum

'------------------------------------------------------------------------------------------
'功能:设置相关进程处理的API声明:2008-10-30 11:34:11:刘兴宏
'------------------------------------------------------------------------------------------
Private gcll_His_PId As Collection        '存储相关的进程信息:array(进程名称,PID,窗口个数),"K"+进程数
Private Const PROCESS_TERMINATE = &H1
Private Type PROCESSENTRY32
      lSize             As Long
      lUsage            As Long
      lProcessId        As Long
      lDefaultHeapId    As Long
      lModuleId         As Long
      lThreads          As Long
      lParentProcessId  As Long
      lPriClassBase     As Long
      lFlags            As Long
      sExeFile          As String * 1024
End Type
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private mstrModalWindows As String
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public glngMain     As Long                 'BH主窗体
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
'在z序中的位于被置位的窗口前的窗口句柄
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public gstrCommand      As String

Public Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '创建模块所需对象的同义词(如果已创建则不会再创建)
    On Error Resume Next
    strSQL = "Zl_Createsynonyms(" & lngSys & ")"
    zlDatabase.ExecuteProcedure strSQL, "创建同义词"
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim StrPass As String, strReturn As String, strSource As String, strTarget As String
    
    StrPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(StrPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

'**********************************************************************************************************************
'功能:以下处理相关进程的函数
'编制:刘兴洪
'日期:2008-10-30 11:38:58
Public Function zlKillHISPID() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:杀死所有HIS启动程序的移常进程(杀的条件是:所有ZLHIS+.exe的进程中无任何窗口)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long, i As Long
    
    zlKillHISPID = False
    Err = 0: On Error GoTo Errhand:
    '第一步:需要处理相关的ZLHIS的相关进程
    Set gcll_His_PId = New Collection
    If zlHISPidToCollect(gcll_His_PId) = False Then zlKillHISPID = True: Exit Function  '如果存在相关的错误，就直接返回了
    If gcll_His_PId Is Nothing Then zlKillHISPID = True: Exit Function
    If gcll_His_PId.Count = 0 Then zlKillHISPID = True: Exit Function
    
    '第二步:需要处理相关ZLHIS的相关进程的相关窗口个数,这样才好判断出相关的进程是否存在异常,出现异常的，就得杀掉
    Call EnumWindows(AddressOf EnumWindowsProc, 0&)
    For i = 1 To gcll_His_PId.Count
        If Val(gcll_His_PId(i)(2)) <= 1 Then
            '肯定窗口数小于1或零,那么肯定有异常，需要杀死他
            If Val(gcll_His_PId(i)(1)) <> 0 Then
                '可能未成功，暂无处理此种情况
                Call TerminatePID(Val(gcll_His_PId(i)(1)))
            End If
        End If
    Next
    zlKillHISPID = True
Errhand:
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取所有窗口符合HIS的进程的窗口
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 10:26:02
    '-----------------------------------------------------------------------------------------------------------
    Dim strTittle As String, lngPID As Long, strName As String
    Dim lngCount As Long
    
    If GetParent(hwnd) = 0 Then
        '读取 hWnd 的视窗标题
        strTittle = String(80, 0)
        Call GetWindowText(hwnd, strTittle, 80)
        strTittle = Left(strTittle, InStr(strTittle, Chr(0)) - 1)
        If Trim(strTittle) <> "" Then
            Call GetWindowThreadProcessId(hwnd, lngPID)
            If IsWindowVisible(hwnd) Then
                Err = 0: On Error Resume Next
                strName = gcll_His_PId("K" & lngPID)(0)
                If Err = 0 Then
                    lngCount = Val(gcll_His_PId("K" & lngPID)(2)) + 1
                    gcll_His_PId.Remove "K" & lngPID
                    gcll_His_PId.Add Array(strName, lngPID, lngCount), "K" & lngPID
                End If
                Err.Clear: On Error GoTo 0
            End If
        End If
    End If
    EnumWindowsProc = True ' 表示继续列举 hWnd
    Exit Function
End Function

Private Function TerminatePID(ByVal lngPID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:结束指定的进程
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long
    TerminatePID = False
    
    Err = 0: On Error GoTo Errhand:
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPID)
    Call TerminateProcess(lngProcess, 1&)
    
    TerminatePID = True
Errhand:
End Function

Private Function zlHISPidToCollect(ByRef cll_His_Pid As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取ZLHIS的进程给相关的集合(gcll_HIS_Pid)
    '入参:
    '出参:cll_His_Pid-将符合HIS.exe的程序，装载该集合中
    '返回:
    '编制:刘兴洪
    '日期:2008-10-30 10:07:38
    '-----------------------------------------------------------------------------------------------------------
    Dim strExeName  As String, lngSnapShot As Long, lngProcess As Long, lngCount  As Long
    Dim strCurExeName As String, lngCurPid As Long
    Dim uProcess   As PROCESSENTRY32
    Const TH32CS_SNAPPROCESS = &H2
    
    Err = 0: On Error GoTo Errhand:
    strCurExeName = "*" & UCase(App.EXEName) & "*"
    
    lngCurPid = GetCurrentProcessId '获取当前应用程序进程
    lngSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If lngSnapShot <> 0 Then
        uProcess.lSize = Len(uProcess)
        lngProcess = ProcessFirst(lngSnapShot, uProcess)
        lngCount = 0
        Do While lngProcess
            '不等于当前进程的才处理
            If lngCurPid <> uProcess.lProcessId Then
                strExeName = UCase(Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1))
                If strExeName Like strCurExeName Then
                    cll_His_Pid.Add Array(strExeName, uProcess.lProcessId, 0), "K" & uProcess.lProcessId
                End If
            End If
            lngProcess = ProcessNext(lngSnapShot, uProcess)
        Loop
        CloseHandle (lngSnapShot)
    End If
    zlHISPidToCollect = True
    Exit Function
Errhand:
End Function
'**********************************************************************************************************************

Public Function ExistModalWindows(ByVal lngHwnd As Long) As String
'功能：判断指定的父窗体中是否存在模态子窗体，如有则返回模态子窗体标题
    mstrModalWindows = ""
    
    EnumChildWindows lngHwnd, AddressOf EnumChildProc, ByVal 0&
    
    If mstrModalWindows <> "" Then
        ExistModalWindows = Mid(mstrModalWindows, 2)
    End If
End Function

Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim strCaption As String
    Dim strCalss As String
    
    strCalss = String(200, " ")
    GetClassName hwnd, strCalss, 200
    strCalss = Trim(Left(strCalss, InStr(strCalss, Chr(0)) - 1))
    
    If strCalss = "ThunderRT6FormDC" Or strCalss = "ThunderFormDC" Then
        If IsWindowEnabled(hwnd) = 0 Then
            strCaption = String(200, " ")
            GetWindowText hwnd, strCaption, 200
            strCaption = Trim(Left(strCaption, InStr(strCaption, Chr(0)) - 1))
            
            mstrModalWindows = mstrModalWindows & "," & strCaption
        End If
    End If
    
    EnumChildProc = 1
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant, i As Long
    arrPars = arrInput
    If gblnTimer Then
        Set OpenSQLRecord = zlDatabase.OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    Else
        Set OpenSQLRecord = OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    End If
End Function

Private Function OpenSQLRecordByArray(ByVal strSQL As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'      cnOracle=当不使用公共连接时传入
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim lngErrNum As Long, strErrInfo As String
    
    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLTmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(zlStr.FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL语句绑定变量不全，调用来源：" & strTitle
    End If

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next
'    If gblnSys = True Then
'        Set cmdData.ActiveConnection = gcnSysConn
'    Else
    Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
'    End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'      cnOracle=当不使用公共连接时传入
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    Dim lngErrNum As Long, strErrInfo As String
    
    If Right(Trim(strSQL), 1) = ")" Then
        '执行的过程名
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符串
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '双"''"的绑定变量处理
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时超过2000个字符要用adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else '可能是其他复杂的表达式，无法处理
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '程序员调用过程时书写错误
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "调用 Oracle 过程""" & strProc & """时，引号或括号书写不匹配。原始语句如下：" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '补充?号
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
        cmdData.CommandType = adCmdText
        cmdData.CommandText = strProc
        
'        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
'        Call gobjComLib.SQLTest
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnOracle.Execute strSQL, , adCmdText
'    Call gobjComLib.SQLTest
End Sub


Public Function IP(Optional ByVal strErr As String) As String
    '******************************************************************************************************************
    '功能:通过oracle获取的计算机的IP地址
    '入参:strDefaultIp_Address-缺省IP地址
    '出参:
    '返回:返回IP地址
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strIp_Address As String
    Dim strSQL As String
        
    On Error GoTo Errhand
    
    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "获取IP地址")
    If rsTmp.EOF = False Then
        strIp_Address = NVL(rsTmp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = OS.IP(strErr)
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    IP = strIp_Address
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strErr & IIf(strErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngErrNum As Long, strErrInfo As String
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    Currentdate = 0
    Err = 0
End Function
