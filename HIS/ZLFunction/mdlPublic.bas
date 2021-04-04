Attribute VB_Name = "mdlPublic"
Option Explicit
'**************************
'       OEM代号
'
'医业  D2BDD2B5
'托普  CDD0C6D5
'**************************
Public Type CustomPar
    格式 As Byte
    值列表 As String
    分类SQL As String
    明细SQL As String
    分类字段 As String
    明细字段 As String
    对象 As String
End Type
Public gfrmMain As Object
Public gblnOK As Boolean, gblnModi As Boolean

'数据库相关定义
Public gcnOracle As ADODB.Connection
Public gstrDBUser As String '用户名
Public gblnDBA As Boolean '是否DBA用户
Public gstrUserName As String '用户姓名
Public gstrUserNO As String '用户编号
Public grsObject As ADODB.Recordset '当前用户所具有Select权限的对象集
'错误日志处理相关变量
Private lngErrNum As Long, strErrInfo As String, bytErrType As Byte

'API相关
Public glngOldProc As Long, glngSelProc As Long
Public glngMinW As Long, glngMaxW As Long, glngMinH As Long, glngMaxH As Long
Public lngTXTProc As Long '保存默认的消息函数的地址

Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Const HKEY_CURRENT_USER = &H80000001
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = &H14F

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = &H101E

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000

'控制TAB键的函数
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const WH_KEYBOARD = 2
Public Const HC_ACTION = 0
Public Const HC_NOREMOVE = 3

Public glngKeyHook As Long
Public gobjTab As clsTabInput
'Html Help
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Const HH_DISPLAY_TOPIC = &H0

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

Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional StrName As String = "")
'功能：将PictureBox模拟成3D平面按钮
'参数：intStyle:0=平面,-1=凹下,1=凸起
    
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If StrName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(StrName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(StrName)) / 2
            picBox.Print StrName
        End If
    End With
End Sub

Public Function CustomHook(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'说明：
'   Code=Hook Code(HC_ACTION、HC_NOREMOVE)
'   wParam=Virtual-Key Code
'   lParam=0-15位(按键的重复次数)
'          16-23位(OEM Scan Code)
'          24位(是否扩展键,如Fx,小键盘键)
'          25-28位(保留)
'          29(ALT是否按下)
'          30(发送消息之前键是否按下)
'          31(0-正在按下,1-正在松开)
    Static blnShift As Boolean
    
    If wParam = vbKeyShift Then
        If lParam > 0 Then
            blnShift = True
        ElseIf lParam < 0 Then
            blnShift = False
        End If
    End If
    If wParam = vbKeyTab Then
        CustomHook = 1
        If blnShift Then
            If lParam > 0 Then
                gobjTab.ACT_sTabKeyDown
            ElseIf lParam < 0 Then
                gobjTab.ACT_sTabKeyUp
            End If
        Else
            If lParam > 0 Then
                gobjTab.ACT_TabKeyDown
            ElseIf lParam < 0 Then
                gobjTab.ACT_TabKeyUp
            End If
        End If
    Else
        CallNextHookEx glngKeyHook, Code, wParam, lParam
    End If
End Function

Public Sub RegFuncFile()
'功能：注册中联函数文件
    Dim strSys As String * 255
    
    GetSystemDirectory strSys, 255
    
    RegSetValue HKEY_CLASSES_ROOT, ".zlf", REG_SZ, "zlFunction", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction", REG_SZ, "中联函数文件", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction\DefaultIcon", REG_SZ, Left(strSys, InStr(strSys, Chr(0)) - 1) & "\zl9Function.dll,0", 24
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction\Shell", REG_SZ, "Read", 4
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction\Shell\Read", REG_SZ, "打开中联函数文件(&1)", 12
    RegSetValue HKEY_CLASSES_ROOT, "zlFunction\Shell\Read\Command", REG_SZ, "NotePad.exe ""%1""", 22
End Sub

Public Function CustomMessage(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    If msg = WM_GETMINMAXINFO Then

        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngMinW
        MinMax.ptMinTrackSize.Y = glngMinH
        MinMax.ptMaxTrackSize.X = glngMaxW
        MinMax.ptMaxTrackSize.Y = glngMaxH
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        CustomMessage = 1
        Exit Function
    End If
    CustomMessage = CallWindowProc(glngOldProc, hwnd, msg, wp, lp)
End Function

Public Function SelMessage(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    If msg = WM_GETMINMAXINFO Then

        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = 400
        MinMax.ptMinTrackSize.Y = 300
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SelMessage = 1
        Exit Function
    End If
    SelMessage = CallWindowProc(glngSelProc, hwnd, msg, wp, lp)
End Function

Public Sub ShowPercent(sngPercent As Single, objPanel As Object)
'功能:在状态条上根据百分比显示当前处理进度(█)
    Dim intAll As Integer
    intAll = objPanel.Width / frmAbout.TextWidth("█") - 4
    objPanel.Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
End Sub

Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    Else
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    End If
End Sub

Public Function CheckLen(txt As Object, intLen As Integer, strInfo As String, Optional AllowNULL As Boolean = True) As Boolean
'功能：检查工本框的真实长度是否在指定限制长度内
    If txt.Text = "" And Not AllowNULL Then
        MsgBox "请输入" & strInfo & "！", vbInformation, App.Title
        txt.SetFocus: Exit Function
    End If
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox "[" & strInfo & "]的长度不能大于 " & intLen & " ！", vbInformation, App.Title
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function

Public Function TLen(str As String) As Integer
'功能：返回字符串的真实长度
    TLen = LenB(StrConv(str, vbFromUnicode))
End Function

Public Function TrimChar(str As String) As String
'功能:去除字符串中连续的空格和回车(含两头的空格,回车),不去除TAB字符,哪怕是连续的
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(str)
    i = InStr(strTmp, "  ")
    Do While i > 0
        strTmp = Left(strTmp, i) & Mid(strTmp, i + 2)
        i = InStr(strTmp, "  ")
    Loop
    
    i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Do While i > 0
        strTmp = Left(strTmp, i + 1) & Mid(strTmp, i + 4)
        i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Loop
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Public Sub CopyPars(ByVal objSPars As FuncPars, ByRef objOPars As FuncPars)
'功能：拷贝参数集对象
    Dim objPar As FuncPar
    
    Set objOPars = New FuncPars
    For Each objPar In objSPars
        With objPar
            objOPars.Add .组名, .序号, .名称, .中文名, .类型, .缺省值, .格式, .值列表, .分类SQL, .明细SQL, .分类字段, .明细字段, .对象, "_" & .Key, .Reserve
        End With
    Next
End Sub

Public Function GetCboIndex(cbo As ComboBox, strFind As String) As Long
'功能：由字任串查找ComboBox的索引值
'参数：cbo=ComboBox,strFind=查找字符串
    Dim i As Integer
    If strFind = "" Then GetCboIndex = -1: Exit Function
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = strFind Then
            GetCboIndex = i
            Exit Function
        End If
    Next
    GetCboIndex = -1
End Function

Public Function CheckSQL(ByVal strSQL As String, strErr As String) As String
'功能：根据SQL语句书写是否正确
'返回：
'     成功=SQL的字段串,包含了各个字段的名称及类型,格式如"姓名,111|年龄,111|奖金,123",类型值以ADO.Field.Type为准
'     失败=空
    Dim rsTmp As New ADODB.Recordset, tmpFld As Field
    Dim strCheck As String, i As Integer
    
    strCheck = strSQL
    
    If InStr(UCase(strCheck), "WHERE") > 0 Then
        strCheck = Replace(UCase(strCheck), "WHERE", "Where Rownum<1 And ")
    End If
    
    Err.Clear
    On Error Resume Next

    Set rsTmp = zlDatabase.OpenSQLRecord(strCheck, "检查SQL")
    If Err.Number = 0 Then
        strErr = ""
        For Each tmpFld In rsTmp.Fields
            If InStr(tmpFld.Name, "|") > 0 Then
                strErr = "字段""" & tmpFld.Name & """没有别名！"
                CheckSQL = "": Exit Function
            Else
                If InStr(CheckSQL & "|", "|" & tmpFld.Name & "," & tmpFld.Type & "|") = 0 Then
                    CheckSQL = CheckSQL & "|" & tmpFld.Name & "," & tmpFld.Type
                Else
                    strErr = "在数据源中发现相同的字段项目！"
                    CheckSQL = "": Exit Function
                End If
            End If
        Next
        CheckSQL = Mid(CheckSQL, 2)
    Else
        strErr = Err.Number & ":" & vbCrLf & Err.Description
        Err.Clear
    End If
End Function

Public Function AdjustStr(str As String) As String
'功能：将含有"'"符号的字符串调整为Oracle所能识别的字符常量
'说明：自动(必须)在两边加"'"界定符。

    Dim i As Long, strTmp As String
    
    If InStr(1, str, "'") = 0 Then AdjustStr = "'" & str & "'": Exit Function
    
    For i = 1 To Len(str)
        If Mid(str, i, 1) = "'" Then
            If i = 1 Then
                strTmp = "CHR(39)||'"
            ElseIf i = Len(str) Then
                strTmp = strTmp & "'||CHR(39)"
            Else
                strTmp = strTmp & "'||CHR(39)||'"
            End If
        Else
            If i = 1 Then
                strTmp = "'" & Mid(str, i, 1)
            ElseIf i = Len(str) Then
                strTmp = strTmp & Mid(str, i, 1) & "'"
            Else
                strTmp = strTmp & Mid(str, i, 1)
            End If
        End If
    Next
    AdjustStr = strTmp
End Function

Public Function MakeFile(strID As String, Optional strFormat As String = "CUSTOM") As String
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, strFormat)
    intFile = FreeFile
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".AVI"
    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    MakeFile = strR
End Function

Public Function Currentdate() As Date
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "SELECT SYSDATE FROM DUAL"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取当前时间")
    Currentdate = rsTmp.Fields(0).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowAbout(Optional frmParent As Object)
    Dim frmShow As New frmAbout
    If frmParent Is Nothing Then
        frmShow.Show 1
    Else
        Load frmShow
        Err.Clear
        On Error Resume Next
        frmShow.Show 1, frmParent
        If Err.Number <> 0 Then
            Err.Clear
            frmShow.Show 1
        End If
    End If
End Sub

Public Function UserObject() As ADODB.Recordset
'功能：获取当前用户所具有查询权限的所有表、视图、函数名(包含用户自身对象及被授权对象)
'返回：成功=对象名称列表(以中英顺序排序),失败=空
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    '所有者."表、视图、函数"
    strSQL = _
        "Select Upper(USER) as OWNER,Upper(OBJECT_NAME) as OBJECT_NAME,OBJECT_TYPE" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW','FUNCTION')" & _
        " Union" & _
        " Select Upper(OWNER) as OWNER,Upper(OBJECT_NAME) as OBJECT_NAME,OBJECT_TYPE" & _
        " From All_Objects O," & _
        " (Select TABLE_NAME From All_Tab_Privs Where Privilege In('SELECT','EXECUTE')) G" & _
        " Where O.Object_Type in('TABLE','VIEW','FUNCTION')" & _
        " and O.OBJECT_NAME=G.TABLE_NAME" & _
        " Order by OWNER,OBJECT_TYPE,OBJECT_NAME"
    
    'ALL_Object是视图,只包含当前用户有权限访问的对象
    strSQL = _
        "Select Upper(USER) as OWNER,Upper(OBJECT_NAME) as OBJECT_NAME,OBJECT_TYPE" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW','FUNCTION')" & _
        " Union" & _
        " Select Upper(OWNER) as OWNER,Upper(OBJECT_NAME) as OBJECT_NAME,OBJECT_TYPE" & _
        " From All_Objects" & _
        " Where Object_Type in('TABLE','VIEW','FUNCTION')" & _
        " Order by OWNER,OBJECT_TYPE,OBJECT_NAME"
 
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "UserObject")
    Set UserObject = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function TrueObject(ByVal strObject As String) As String
'功能：SQLObject函数的子函数,用于去除对象名中的无用字符
    Dim i As Integer
    '寻找第一个正常字符位置
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    '寻找后面第一个非正常字符
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function

Public Function SQLObject(ByVal strSQL As String) As String
'功能：分析SQL语句所用到的对象名
'参数：strSQL=要分析的原始SQL语句
'返回：SQL语句所访问到的对象名,如"部门表,病人费用记录,ZLHIS.人员表"
'说明：1.与Oracle SELECT语句兼容
'      2.如果SQL语句中的对象名前加有所有者前缀,则该前缀不会被截取
'      3.需要函数TrimChar;TrueObject的支持
    Dim intB As Integer, intE As Integer, intL As Integer, intR As Integer
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Integer, j As Integer
    
    On Error GoTo errH
    
    '大写化及去除多余的字符
    strAnal = UCase(TrimChar(strSQL))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
    '先分解处理嵌套子查询
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB '匹配的左右括号位置
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                If intE - intB - 1 <= 0 Then
                    '对于非子查询,将括号换成其它符号,以使循环继续
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '子查询语句
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '将该子查询部份作为为特殊对象名
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "嵌套查询")
                    '递归分析
                    strObject = strObject & "," & SQLObject(strSub)
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '无匹配右括号
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '分解分析(此时strAnal为简单查询,可能带Union等连接)
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '从第一个From后面部份开始
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "START WITH") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "START WITH") - 1)
        ElseIf InStr(strCur, "CONNECT BY") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "CONNECT BY") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        ElseIf InStr(strCur, "MINUS") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "MINUS") - 1)
        ElseIf InStr(strCur, "INTERSECT") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "INTERSECT") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject & ",", "," & strTrue & ",") = 0 And strTrue <> "嵌套查询" Then
                If InStr(strTrue, "'") = 0 And InStr(strTrue, "@") = 0 Then
                    strObject = strObject & "," & strTrue
                End If
            End If
        Next
    Next
    '完成
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    Err.Clear
End Function

Public Function CheckObjectPriv(strObject As String, strOwner As String) As String
'功能：检查当前用户对指定对象是否完全有权限访问
'参数：strObject=对象名串,如"部门表,病人费用记录"
'      strOwner=检查依据的所有者
'返回：完全=空,不完全=不能访问的对象名,如"部门表,病人费用记录"
'说明：用于在校验数据源之前检查是否有权限查询SQL语句中的对象
'参考：grsObject
    Dim i As Integer
    For i = 0 To UBound(Split(strObject, ","))
        If Split(strObject, ",")(i) <> "DUAL" Then
            If InStr(Split(strObject, ",")(i), ".") = 0 Then
                If gblnDBA Then
                    grsObject.Filter = "OBJECT_TYPE<>'FUNCTION' And OBJECT_NAME='" & UCase(Split(strObject, ",")(i)) & "'"
                Else
                    grsObject.Filter = "OBJECT_TYPE<>'FUNCTION' And OBJECT_NAME='" & UCase(Split(strObject, ",")(i)) & "' And OWNER='" & UCase(strOwner) & "'"
                End If
            Else
                '如果本身就加了所有者前缀,则检查该所有者对象权限
'                If gblnDBA Then
'                    grsObject.Filter = "OBJECT_NAME='" & UCase(Split(Split(strObject, ",")(i), ".")(1)) & "'" & _
'                        " And OBJECT_TYPE<>'FUNCTION'"
'                Else
'                    grsObject.Filter = "OWNER='" & UCase(Split(Split(strObject, ",")(i), ".")(0)) & _
'                        "' And OBJECT_NAME='" & UCase(Split(Split(strObject, ",")(i), ".")(1)) & "'" & _
'                        " And OBJECT_TYPE<>'FUNCTION'"
'                End If
                grsObject.Filter = "OBJECT_NAME='" & UCase(Split(Split(strObject, ",")(i), ".")(1)) & "'" & _
                    " And OBJECT_TYPE<>'FUNCTION'"
            End If
            If grsObject.EOF Then
                If InStr(CheckObjectPriv & ",", "," & Split(strObject, ",")(i) & ",") = 0 Then
                    CheckObjectPriv = CheckObjectPriv & "," & Split(strObject, ",")(i)
                End If
            End If
        End If
    Next
    If CheckObjectPriv <> "" Then CheckObjectPriv = Mid(CheckObjectPriv, 2)
End Function

Public Function ObjectOwner(strObject As String, strOwner As String, Optional frmParent As Object) As String
'功能：根据对象名加上当前用户所能访问的所有者前缀(包括对同一对象名有多个所有者要求选其中之一)
'参数：strObject=对象名串,如"部门表,病人费用记录"
'返回：正常=加了所有者前缀的对象串,如"ZLPER.部门表,ZLHIS.病人费用记录",取消="取消"
'参考：grsObject
    Dim i As Integer, j As Integer
    
    For i = 0 To UBound(Split(strObject, ","))
        If Split(strObject, ",")(i) <> "DUAL" Then
            If InStr(Split(strObject, ",")(i), ".") > 0 Then
                '如果本身就加了所有者前缀,则使用其本身不变
                If InStr(ObjectOwner, "," & Split(strObject, ",")(i)) = 0 Then
                    ObjectOwner = ObjectOwner & "," & Split(strObject, ",")(i)
                End If
            Else
                If gblnDBA Then
                    grsObject.Filter = "OBJECT_TYPE<>'FUNCTION' And OBJECT_NAME='" & UCase(Split(strObject, ",")(i)) & "'"
                Else
                    grsObject.Filter = "OBJECT_TYPE<>'FUNCTION' And OBJECT_NAME='" & UCase(Split(strObject, ",")(i)) & "' And OWNER='" & UCase(strOwner) & "'"
                End If
                If grsObject.RecordCount = 1 Then
                    If InStr(ObjectOwner & ",", "," & grsObject!OWNER & "." & Split(strObject, ",")(i) & ",") = 0 Then
                        ObjectOwner = ObjectOwner & "," & grsObject!OWNER & "." & Split(strObject, ",")(i)
                    End If
                ElseIf grsObject.RecordCount > 1 Then
                    '同一对象有多个所有者,则要求选择
                    Set frmSelOwner.rsObject = grsObject
                    If frmParent Is Nothing Then
                        frmSelOwner.Show 1
                    Else
                        frmSelOwner.Show 1, frmParent
                    End If
                    If gblnOK Then
                        With frmSelOwner.lvw.SelectedItem
                            If InStr(ObjectOwner & ",", "," & .Text & "." & Split(strObject, ",")(i) & ",") = 0 Then
                                ObjectOwner = ObjectOwner & "," & .Text & "." & Split(strObject, ",")(i)
                            End If
                        End With
                        Unload frmSelOwner
                    Else
                        '取消选择,也就是取消操作(调用程序),返回空
                        ObjectOwner = "取消": Exit Function
                    End If
                End If
            End If
        End If
    Next
    If ObjectOwner <> "" Then ObjectOwner = Mid(ObjectOwner, 2)
End Function

Public Function SQLOwner(ByVal strSQL As String, strOwner As String) As String
'功能：将SQL语句替换成带对象所有者的形式
'参数：strSQL=原始SQL语句,strOwner=对象所有者串,如"ZLPER.部门表,ZLHIS.病人费用记录"
'返回：访问对象加了所有者前缀的SQL语句
'说明：1.本函数用于直接执行用户SQL语句,而不需要授权对象的私有同义词。
'      2.对表名与字段名相同且字段名没有带表别名,则会出错
    Dim i As Integer, j As Integer
    Dim intLoc As Integer, blnDo As Boolean
    
    '处理成只用空格间隔
    strSQL = UCase(SpaceSQL(strSQL))
    
    For i = 0 To UBound(Split(strOwner, ","))
        '采用循环确认方式,确保替换的是表名,而不是其它语句部份或被包含在其它表名中的部份
        j = 0 '当前开始查找位置
        Do
            j = j + 1
            intLoc = InStr(j, strSQL, Split(Split(strOwner, ",")(i), ".")(1))
            If intLoc > 12 Then '至少有"SELECT FROM "
                '本身就有所有者前缀的不替换
                blnDo = True
                '右边以空格、","号、右括号结束
                blnDo = blnDo And (InStr(",) ", Mid(strSQL, intLoc + Len(Split(Split(strOwner, ",")(i), ".")(1)), 1)) > 0)
                '左边则为","号或"FROM "
                blnDo = blnDo And (Mid(strSQL, intLoc - 1, 1) = "," Or Mid(strSQL, intLoc - 5, 5) = "FROM ")
                If blnDo Then
                    strSQL = Left(strSQL, intLoc - 1) & _
                        Replace(strSQL, Split(Split(strOwner, ",")(i), ".")(1), Split(strOwner, ",")(i), intLoc, 1)
                    j = intLoc + Len(Split(strOwner, ",")(i))
                End If
            End If
        Loop Until j >= Len(strSQL)
    Next
    SQLOwner = strSQL
End Function

Public Function InDesign() As Boolean
'功能：判断当前运行程序是否在VB的工程环境中
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function GetDBUser() As String
'功能：获取当前登录数据库用户名
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
        
    On Error GoTo errH
        
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State = adStateClosed Then Exit Function
    If InStr(UCase(gcnOracle.ConnectionString), "USER ID=") > 0 Then
        For i = 0 To UBound(Split(UCase(gcnOracle.ConnectionString), ";"))
            If Split(UCase(gcnOracle.ConnectionString), ";")(i) Like "USER ID=*" Then
                GetDBUser = Trim(Split(Split(UCase(gcnOracle.ConnectionString), ";")(i), "=")(1))
                Exit For
            End If
        Next
    Else
        strSQL = "Select User From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "当前登录数据库用户名")
        If Not rsTmp.EOF Then GetDBUser = rsTmp!USER
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AutoSizeCol(lvw As Object)
'功能：根据自动ListView当前内容自动调整各列宽度
'参数：blnByHead=是否按列头文本调整,Col=指定列还是所有列(1-N)
    Dim i As Integer, lngW As Long
    For i = 1 To lvw.ColumnHeaders.Count
        SendMessage lvw.hwnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If lvw.ColumnHeaders(i).Width < 200 Then lvw.ColumnHeaders(i).Width = 0
        If lvw.ColumnHeaders(i).Width < (TLen(lvw.ColumnHeaders(i).Text) + 2) * 90 And lvw.ColumnHeaders(i).Width <> 0 Then lvw.ColumnHeaders(i).Width = (TLen(lvw.ColumnHeaders(i).Text) + 2) * 90
    Next
End Sub

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：保存窗体及其中各种控件的状态
'参数：objForm:要保存的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    
    Dim objThis As Object
    Dim strTmp As String
    Dim i As Integer, blnDo As Boolean
    
    On Error Resume Next
    If Not gfrmMain Is Nothing Then Call gfrmMain.Shut任务(objForm)
    On Error GoTo 0
    
    If GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "1") = "0" Then
        Call DelWinState(objForm, strProjectName, strUserDef)
        SaveWinState = True: Exit Function
    End If
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '保存窗体状态、位置、大小
    With objForm
        Select Case .WindowState
            Case 0
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", objForm.WindowState & "," & .Left & "," & .Top & "," & .Width & "," & .Height
            Case 1
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", 0
            Case 2
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", objForm.WindowState
        End Select
    End With
    
    '保存各种控件的各种状态
    For Each objThis In objForm.Controls
        strTmp = ""
        On Error Resume Next
        If UCase(TypeName(objThis)) = UCase("Menu") Then
            If objThis.Caption Like "标准按钮*" Or _
                objThis.Caption Like "文本标签*" Or _
                objThis.Caption Like "状态栏*" Or _
                UCase(objThis.Name) Like UCase("mnuViewTool*") Then
                '特殊菜单的复选
                strTmp = objThis.Checked & "," & objThis.Enabled
            Else
                strTmp = ""
            End If
        ElseIf (UCase(objThis.Tag) = "SAVE" Or UCase(objThis.Name) Like "*_S" Or _
            UCase(TypeName(objThis)) = UCase("StatusBar") Or _
            UCase(TypeName(objThis)) = UCase("Toolbar") Or _
            UCase(TypeName(objThis)) = UCase("Coolbar")) And objForm.Visible Then

            blnDo = True
            If UCase(TypeName(objThis)) = UCase("Toolbar") Or UCase(objThis.Tag) = "SAVE" Or UCase(objThis.Name) Like "*_S" Then
                If TypeName(objThis.Container) = "PictureBox" Then blnDo = False
            End If
            'Left,Top,Width、Height,Visible
            strTmp = strTmp & "," & objThis.Left
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Top
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Width
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Height
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            If blnDo Then
                strTmp = strTmp & "," & objThis.Visible
                If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            Else
                strTmp = strTmp & ",-32767"
            End If
            strTmp = Mid(strTmp, 2)
        End If
        If strTmp <> "" Then
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "状态", strTmp
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("Toolbar")
                If objThis.Buttons.Count > 0 Then
                    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "文本", IIf(objThis.Buttons(1).Caption <> "", 1, objThis.ButtonHeight)
                End If
            Case UCase("ListView")
                SaveListViewState objThis, strProjectName & objForm.Name & strUserDef
            Case UCase("CoolBar")
                strTmp = ""
                For i = 1 To objThis.Bands.Count
                    strTmp = strTmp & "," & objThis.Bands(i).NewRow
                Next
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "行序", Mid(strTmp, 2)
                
                strTmp = ""
                For i = 1 To objThis.Bands.Count
                    strTmp = strTmp & "," & objThis.Bands(i).Visible
                Next
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "可见栏", Mid(strTmp, 2)
        End Select
    Next
    SaveWinState = True
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：恢复窗体的状态，当左顶边界超出时，则自动设置为0
'参数：objForm:要恢复的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
   
    Dim aryInfo() As String
    Dim strTmp As String, i As Integer
    Dim objThis As Object
    Dim blnDo As Boolean
    Dim strSave As String
    Dim strOEM As String
    
    On Error Resume Next
    
    If Not gfrmMain Is Nothing Then Call gfrmMain.Show任务(objForm)
    
    blnDo = (GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "0") = "1")
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '恢复窗体的状态、位置、大小
    If UCase(objForm.Name) = UCase("frmReport") _
        Or UCase(objForm.Name) = UCase("frmPreview") _
            Or UCase(objForm.Name) = UCase("frmDesign") Then
        strTmp = "2" '特殊窗体初始最大化
    Else
        strTmp = "0," & (Screen.Width - objForm.Width) / 2 & "," & (Screen.Height - objForm.Height) / 2 & "," & objForm.Width & "," & objForm.Height
    End If
    If blnDo Then
        strSave = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", "")
        RestoreWinState = (strSave <> "")
        If strSave = "" Then strSave = strTmp
        aryInfo = Split(strSave, ",")
    Else
        aryInfo = Split(strTmp, ",")
    End If
    With objForm
        .WindowState = aryInfo(0)
        If UBound(aryInfo) = 4 Then
            .Left = IIf(aryInfo(1) < 0, 0, aryInfo(1))
            .Top = IIf(aryInfo(2) < 0, 0, aryInfo(2))
            .Width = IIf(aryInfo(3) > Screen.Width, Screen.Width, aryInfo(3))
            .Height = IIf(aryInfo(4) > Screen.Height, Screen.Height, aryInfo(4))
        Else
            .Left = (Screen.Width - objForm.Width) / 2
            .Top = (Screen.Height - objForm.Height) / 2
        End If
    End With

    '恢复窗体中各种控件的各种状态
    For Each objThis In objForm.Controls
        
        On Error Resume Next
        
        If blnDo Then
            strTmp = ""
            If UCase(TypeName(objThis)) = UCase("Menu") Then
                '特殊菜单的复选
                If objThis.Caption Like "标准按钮*" Or _
                    objThis.Caption Like "文本标签*" Or _
                    objThis.Caption Like "状态栏*" Or _
                    UCase(objThis.Name) Like UCase("mnuViewTool*") Then
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "状态", "")
                    If UBound(Split(strTmp, ",")) = 1 Then
                        objThis.Checked = Split(strTmp, ",")(0)
                        objThis.Enabled = Split(strTmp, ",")(1)
                    End If
                End If
            ElseIf UCase(objThis.Tag) = "SAVE" Or UCase(objThis.Name) Like "*_S" Or _
                UCase(TypeName(objThis)) = UCase("StatusBar") Or _
                UCase(TypeName(objThis)) = UCase("Toolbar") Or _
                UCase(TypeName(objThis)) = UCase("Coolbar") Then
                
                strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "状态", "")
                If strTmp <> "" Then
                    'Left,Top,Width、Height,Visible
                    If UBound(Split(strTmp, ",")) = 4 Then
                        If Split(strTmp, ",")(0) <> "-32767" Then objThis.Left = Split(strTmp, ",")(0)
                        If Split(strTmp, ",")(1) <> "-32767" Then objThis.Top = Split(strTmp, ",")(1)
                        If Split(strTmp, ",")(2) <> "-32767" Then objThis.Width = Split(strTmp, ",")(2)
                        If Split(strTmp, ",")(3) <> "-32767" Then objThis.Height = Split(strTmp, ",")(3)
                        If Split(strTmp, ",")(4) <> "-32767" Then objThis.Visible = Split(strTmp, ",")(4)
                    End If
                End If
            End If
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("StatusBar")
                '状态条试用标志
'                If zlRegInfo("授权性质") <> "1" Then
'                    If objThis.Panels(1).Bevel = sbrRaised Then
'                        objThis.Panels(1).Text = ""
'                        Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
'                        objThis.Panels(1).ToolTipText = ""
'                        objThis.Height = 360
'                    End If
'                Else
                    If objThis.Panels(1).Bevel = sbrRaised Then
                        strTmp = zlRegInfo("产品简名")
                        If strTmp <> "-" Then
                            objThis.Panels(1).Text = strTmp & "软件"
                            '处理状态栏图标的OEM策略
                            If strTmp = "中联" Then
                                If zlRegInfo("授权性质") <> "1" Then
                                    Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
                                    objThis.Panels(1).Text = ""
                                Else
                                    Set objThis.Panels(1).Picture = LoadCustomPicture("Logo")
                                End If
                            Else
                                strOEM = GetOEM(strTmp)
                                Set objThis.Panels(1).Picture = LoadCustomPicture(strOEM)
                                If Err <> 0 Then
                                    Err.Clear
                                Set objThis.Panels(1).Picture = LoadCustomPicture("Logo")
                                End If
                                If zlRegInfo("授权性质") <> "1" Then objThis.Panels(1).Text = strTmp & "(试用)"
                            End If
                            objThis.Panels(1).ToolTipText = ""
                            objThis.Height = 360
                        End If
                    End If
'                End If
            Case UCase("Menu")
                If UCase(objThis.Name) = UCase("mnuHelpWeb") Then
                    'WEB上的中联
                    strTmp = zlRegInfo("支持商简名")
                    If strTmp <> "-" Then
                        objThis.Caption = "&WEB上的" & strTmp
                    End If
                ElseIf UCase(objThis.Name) = UCase("mnuHelpWebHome") Then
                    '中联主页
                    strTmp = zlRegInfo("支持商简名")
                    If strTmp <> "-" Then
                        objThis.Caption = strTmp & "主页(&H)"
                    End If
                End If
            Case UCase("Toolbar")
                If blnDo Then
                    If objThis.Buttons.Count > 0 Then
                        strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "文本", 1)
                        For i = 1 To objThis.Buttons.Count
                            objThis.Buttons(i).Caption = IIf(strTmp = 1, objThis.Buttons(i).Tag, "")
                        Next
                    End If
                End If
            Case UCase("ListView")
                If blnDo Then
                    RestoreListViewState objThis, strProjectName & objForm.Name & strUserDef
                End If
            Case UCase("CoolBar")
                If blnDo Then
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "行序", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).NewRow = Split(strTmp, ",")(i)
                        Next
                    End If
            
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "可见栏", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).Visible = Split(strTmp, ",")(i)
                        Next
                    End If
                End If
        End Select
    Next
End Function

Public Function RestoreFlexState(objThis As Object, strForm As String) As Boolean
    Dim i As Integer, strTmp As String
        
    On Error Resume Next
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & "宽度", "")
    If UBound(Split(strTmp, ",")) >= 0 Then
        For i = 0 To objThis.Cols - 1
            If objThis.ColWidth(i) > 0 Then
                objThis.ColWidth(i) = Split(strTmp, ",")(i)
            End If
        Next
        RestoreFlexState = True
    End If
End Function

Public Sub SaveFlexState(objThis As Object, strForm As String)
    Dim strTmp As String, i As Integer
        
    On Error Resume Next
    
    strTmp = ""
    For i = 0 To objThis.Cols - 1
        strTmp = strTmp & "," & objThis.ColWidth(i)
    Next
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & "宽度", Mid(strTmp, 2)
End Sub

Public Sub SaveListViewState(objLvw As Object, ByVal strForm As String)
'功能：保存ListView的各种特性
'参数：objLvw=ListView对象,strForm=窗体关键字
'说明：视图方式、列宽、列位置、列标题、列对齐、排序
    Dim lngCol As Long
    Dim strWidth As String
    Dim strPosition As String
    Dim strText As String
    Dim strAlign As String
    
    For lngCol = 1 To objLvw.ColumnHeaders.Count
        strWidth = strWidth & "," & objLvw.ColumnHeaders(lngCol).Width
        strPosition = strPosition & "," & objLvw.ColumnHeaders(lngCol).Position
        strText = strText & "," & objLvw.ColumnHeaders(lngCol).Text
        strAlign = strAlign & "," & objLvw.ColumnHeaders(lngCol).Alignment
    Next
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "视图", objLvw.View
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "宽度", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "位置", Mid(strPosition, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "名称", Mid(strText, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "对齐", Mid(strAlign, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "排序", objLvw.SortKey & "," & objLvw.SortOrder & "," & objLvw.Sorted
End Sub

Public Sub RestoreListViewState(objLvw As Object, ByVal strForm As String)
'功能：恢复ListView的各种特性
'参数：objLvw=ListView对象,strForm=窗体关键字
'说明：视图方式、列宽、列位置、列标题、列对齐、排序
    Dim lngCol As Long
    Dim strWidth As String
    Dim strPosition As String
    Dim strText As String, varText As Variant
    Dim strAlign As String
    Dim strSort As String
    
    On Error Resume Next
    
    '视图缺省保持初始值
    lngCol = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "视图", -1)
    If lngCol <> -1 Then objLvw.View = lngCol
    
    strWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "宽度")
    strPosition = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "位置")
    strAlign = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "对齐")
    For lngCol = 1 To objLvw.ColumnHeaders.Count
        '列缺省关键字为"_" & 列标题
        objLvw.ColumnHeaders(lngCol).Key = "_" & objLvw.ColumnHeaders(lngCol).Text
        If strWidth <> "" Then objLvw.ColumnHeaders(lngCol).Width = Split(strWidth, ",")(lngCol - 1)
        If strPosition <> "" Then objLvw.ColumnHeaders(lngCol).Position = Split(strPosition, ",")(lngCol - 1)
        If strAlign <> "" Then objLvw.ColumnHeaders(lngCol).Alignment = Split(strAlign, ",")(lngCol - 1)
    Next
    
    '排序特性
    strSort = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & "排序")
    If strSort <> "" Then
        objLvw.SortKey = Split(strSort, ",")(0)
        objLvw.SortOrder = Split(strSort, ",")(1)
        objLvw.Sorted = Split(strSort, ",")(2)
    End If
End Sub

Public Function DelWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：删除窗体个性化设置值
'参数：objForm:要恢复的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    Dim strProject As String
    Dim lngR As Long
    Dim objThis As Object
    
    strProject = strProjectName
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    For Each objThis In objForm.Controls
        lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis) & Chr(0))
        If lngR <> 0 And lngR <> 2 Then Exit Function
    Next
    
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form" & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    
    DelWinState = True
End Function

Public Function LoadCustomPicture(strID As String, Optional strFormat As String = "GIF") As StdPicture
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, strFormat)
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Function MatchIndex(ByVal cbo As Object, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：cbo.Hwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：-2=未加处理,其它=匹配的索引(含不匹配的索引)
'说明：请将该函数在KeyPress事件中调用。

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> cbo.hwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = cbo.hwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        MatchIndex = SendMessage(cbo.hwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then
            cbo.Text = strFind
            cbo.SelStart = Len(cbo.Text)
        End If
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
End Function

Public Function RemoveOrderBy(ByVal str As String) As String
'功能：将SQL语句中最后的Order by 语句去除
    Dim i As Integer, intMax As Integer
    Dim strTmp As String
    
    strTmp = UCase(str): intMax = -1
    Do While strTmp Like UCase("*ORDER BY*")
        i = InStr(UCase(strTmp), "ORDER BY")
        If i > intMax Then intMax = i
        strTmp = Left(strTmp, i - 1) & "12345678" & Mid(strTmp, i + 8)
    Loop
    If intMax <> -1 Then
        RemoveOrderBy = Left(str, intMax - 1)
    Else
        RemoveOrderBy = str
    End If
End Function

Public Function GetDefaultValue(strSQL As String, strFld As String) As String
'功能：根据参数选择器SQL定义，返回显示字段及绑定字段的值
'返回：显示值|绑定值|记录数
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim strShow As String, strBand As String
    Dim strSQLT As String
    
    On Error GoTo errH
    
    strSQLT = Replace(RemoveNote(strSQL), "[*]", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQLT, "选择器SQL定义")
    If Not rsTmp.EOF Then
        For i = 0 To UBound(Split(strFld, "|"))
            strTmp = Split(strFld, "|")(i)
            If Split(strTmp, ",")(2) Like "*&D*" Then
                strShow = IIf(IsNull(rsTmp.Fields(CStr(Split(strTmp, ",")(0))).Value), "", rsTmp.Fields(CStr(Split(strTmp, ",")(0))).Value)
            End If
            If Split(strTmp, ",")(2) Like "*&B*" Then
                strBand = IIf(IsNull(rsTmp.Fields(CStr(Split(strTmp, ",")(0))).Value), "", rsTmp.Fields(CStr(Split(strTmp, ",")(0))).Value)
            End If
        Next
    End If
    If strShow <> "" Or strBand <> "" Then
        GetDefaultValue = strShow & "|" & strBand & "|" & rsTmp.RecordCount
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hwnd, msg, wp, lp)
End Function

Private Function GetOEM(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Function RemoveNote(ByVal strSQL As String) As String
'功能：移除SQL语句中的注释
    Dim strTmp As String, i As Integer
    
    strSQL = Replace(strSQL, vbTab, " ")
    strSQL = Replace(strSQL, vbCrLf, vbLf)
    strSQL = Replace(strSQL, vbCr, vbLf)
    strSQL = Replace(strSQL, vbLf & vbLf, vbLf)
    
    For i = 0 To UBound(Split(strSQL, vbLf))
        If Not Trim(Split(strSQL, vbLf)(i)) Like "--*" Then
            RemoveNote = RemoveNote & vbCrLf & Split(strSQL, vbLf)(i)
        End If
    Next
    RemoveNote = Mid(RemoveNote, 3)
End Function

Public Function ShowHelpFunc(SHwnd As Long, ByVal htmName As String, Optional Sys As Integer = 1) As Boolean
'显示帮助窗体
'SHwnd:传入窗口句柄(作为宿主窗口)
'htmName:射映在CHM中的htm文件名称
'Sys:系统,0:函数工具;1:zlhis
    Dim Path As String
    Dim strSave As String
    
    On Error GoTo ShowHelpErr
    
    ShowHelpFunc = False
    strSave = String(200, Chr$(0))
    If Sys = 0 Then
        Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9server.chm"
        If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
        Call Htmlhelp(SHwnd, Path, &H0, "zlreport\" & htmName & ".htm")
    Else
        If Mid(UCase(htmName), 5, 6) = "INSIDE" Then
            Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9server.chm"
            If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
            Call Htmlhelp(SHwnd, Path, &H0, "zlreport\report.htm")
        Else
            Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9app" & Trim(Format(Sys)) & ".chm"
            If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
            strSave = "zl9app" & Trim(Format(Sys)) & "rpt\" & htmName & ".htm"
            Call Htmlhelp(SHwnd, Path, &H0, strSave)
        End If
    End If
    ShowHelpFunc = True
    Exit Function
ShowHelpErr:
    Err.Clear
End Function

Public Sub SetColWidth(msh As Control, objForm As Object)
'功能：自动调整表格列宽,以最小适合为准
    Dim arrWidth() As Long
    Dim i As Integer, j As Integer
    
    ReDim arrWidth(msh.Cols - 1)
    
    msh.Redraw = False
    Set objForm.Font = msh.Font
    
    For i = 0 To msh.Cols - 1
        If msh.ColWidth(i) <> 0 Then
            For j = IIf(msh.FixedRows = 0, 0, msh.FixedRows - 1) To msh.Rows - 1
                If objForm.TextWidth(msh.TextMatrix(j, i) & "ab") + 45 > arrWidth(i) Then
                    arrWidth(i) = objForm.TextWidth(msh.TextMatrix(j, i) & "AB") + 45
                End If
            Next
        End If
    Next
    
    For i = 0 To msh.Cols - 1
        If msh.ColWidth(i) <> 0 Then msh.ColWidth(i) = arrWidth(i)
    Next
    msh.Redraw = True
End Sub

Public Function HaveDBA() As Boolean
'功能：判断当前用户是否具有DBA权限
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select * From Session_Roles Where Role='DBA'"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "判断当前用户DBA权限")
    HaveDBA = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFunSource(strOwner As String, strFunc As String) As String
'功能：获取指定函数的源代码
'说明：1.返回的函数文本肯定以"FUNCTION xxxxx"开头
'      2.函数代码行在数据库中以vbLf结束,vbcf则转换成了空格
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strText As String, strTmp As String
    
    On Error GoTo errH
    
    strSQL = "Select * From All_Source Where TYPE='FUNCTION' And Upper(Owner)=Upper('" & strOwner & "') And Upper(Name)=Upper('" & strFunc & "') Order by Line"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取指定函数的源代码")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strTmp = IIf(IsNull(rsTmp!Text), "", rsTmp!Text)
            strText = strText & strTmp
            rsTmp.MoveNext
        Next
    End If
    GetFunSource = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReplaceName(ByVal strCode As String, strOld As String, strNew As String) As String
'功能：在函数代码中将函数名替换成新的函数名
'说明：为了不改变函数代码的大小写,所以有此函数
'说明：函数代码行在数据库中以vbLf结束,vbcf则转换成了空格
    Dim i As Integer, strText As String
    Dim arrText() As String, strTmp As String
    
    arrText = Split(strCode, vbLf)
    For i = 0 To UBound(arrText)
        strTmp = arrText(i)
        If UCase(strTmp) Like UCase("*" & strOld & "*") Then
            strTmp = Replace(UCase(strTmp), UCase(strOld), UCase(strNew))
        End If
        strText = strText & vbCrLf & strTmp
    Next
    ReplaceName = Mid(strText, 3)
End Function

Public Function CheckParPrivs(系统 As Long, 所有者 As String, 函数号 As Integer) As String
'功能：检查是否具有指定函数中的选择器对象的查询权限
'参数：所有者=用于检查在该所有者下是否具有权限
'返回：该所有者下不具有权限的对象,如"部门表,人员表"
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim arrObj() As String, i As Integer, j As Integer
    Dim strOwner As String, strObj As String
    
    On Error GoTo errH
    
    strSQL = "Select * From zlFuncpars Where 对象 is Not NULL And 系统=" & 系统 & " And 函数号=" & 函数号
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询权限")
    
    For i = 1 To rsTmp.RecordCount
        strObj = Replace(rsTmp!对象, "|", ",")
        If Left(strObj, 1) = "," Then strObj = Mid(strObj, 2)
        If Right(strObj, 1) = "," Then strObj = Mid(strObj, 1, Len(strObj) - 1)
        arrObj = Split(strObj, ",")
        For j = 0 To UBound(arrObj)
            strOwner = Split(arrObj(j), ".")(0)
            strObj = Split(arrObj(j), ".")(1)
            
            If gblnDBA Then
                grsObject.Filter = "Object_Type<>'FUNCTION' And Object_Name='" & UCase(strObj) & "'"
            Else
                grsObject.Filter = "OWNER='" & UCase(所有者) & "' And Object_Type<>'FUNCTION' And Object_Name='" & UCase(strObj) & "'"
            End If
            If grsObject.EOF Then
                CheckParPrivs = CheckParPrivs & "," & strObj
            End If
        Next
        rsTmp.MoveNext
    Next
    CheckParPrivs = Mid(CheckParPrivs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadFuncPars(系统 As Long, 函数号 As Integer) As FuncPars
'功能：读取指定函数的参数集
    Dim strSQL As String, objPars As New FuncPars
    Dim rsTmp As New ADODB.Recordset, i As Integer
    
    Set ReadFuncPars = New FuncPars
    
    On Error GoTo errH
    
    strSQL = "Select * From zlFuncPars Where 系统=" & 系统 & " And 函数号=" & 函数号 & " Order by 参数号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取指定函数的参数集")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With rsTmp
                objPars.Add IIf(IsNull(!组名), "", !组名), !参数号, !参数名, _
                    IIf(IsNull(!中文名), "", !中文名), !类型, IIf(IsNull(!缺省值), "", !缺省值), _
                    IIf(IsNull(!格式), 0, !格式), IIf(IsNull(!值列表), "", !值列表), _
                    IIf(IsNull(!分类SQL), "", !分类SQL), IIf(IsNull(!明细SQL), "", !明细SQL), _
                    IIf(IsNull(!分类字段), "", !分类字段), IIf(IsNull(!明细字段), "", !明细字段), _
                    IIf(IsNull(!对象), "", !对象), "_" & !参数名
            End With
            rsTmp.MoveNext
        Next
    End If
    Set ReadFuncPars = objPars
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFuncPars(ByVal strCode As String) As String
'功能：从函数代码中获取参数定义
'参数：strCode=函数代码,至少从参数部份"("开始。
'返回："参数名,参数类型;...",如"NO_IN,VarChar2;ID_IN,Number;..."
'说明：1.因为函数定义允许,参数类型也可能是"部门表.名称%Type"的形式
'      2.返回值中不区分参数的IN/OUT类型。
    Dim strTmp As String, i As Integer, j As Integer
    Dim blnStart As Boolean, arrPars() As String
    Dim strPar As String, arrOne() As String
    
    '移除注释
    strCode = RemoveNote(strCode)
    
    '仅用空格间隔
    strCode = Replace(strCode, vbTab, " ")
    strCode = Replace(strCode, vbCr, " ")
    strCode = Replace(strCode, vbLf, " ")
    
    '求出Begin关键字的开始位置:Begin不会单独作为参数名,函数名
    strTmp = "": blnStart = False: j = 0
    For i = 1 To Len(strCode)
        If Mid(strCode, i, 1) <> " " Then
            blnStart = True
            strTmp = strTmp & Mid(strCode, i, 1)
        ElseIf blnStart Then
            blnStart = False
            If UCase(strTmp) = "BEGIN" Then
                j = i - Len("Begin")
                Exit For
            End If
            strTmp = ""
        End If
    Next
    If j = 0 Then Exit Function
    
    'Begin前面的代码
    strCode = Trim(Left(strCode, j - 1))
    If InStr(strCode, "(") = 0 Then Exit Function
    '代码中()之间的参数定义
    strCode = Trim(Mid(strCode, InStr(strCode, "(") + 1))
    strCode = Trim(Left(strCode, InStr(strCode, ")") - 1))
    
    '函数局部变量使用"(x),(x,y)"
    If IsNumeric(Trim(strCode)) Then Exit Function
    arrPars = Split(strCode, ",") '参数以","号间隔
    For i = 0 To UBound(arrPars)
        If IsNumeric(Trim(arrPars(i))) Then Exit Function
    Next
    
    '分解参数:忽略参数IN.OUT
    For i = 0 To UBound(arrPars)
        arrOne = Split(Trim(arrPars(i)), " ")
        For j = 0 To UBound(arrOne)
            If j = 0 Then
                strPar = strPar & ";" & Trim(arrOne(j))
            ElseIf InStr(UCase(Trim(arrOne(j))), "CHAR") > 0 Or _
                InStr(UCase(Trim(arrOne(j))), "DATE") > 0 Or _
                InStr(UCase(Trim(arrOne(j))), "NUMBER") > 0 Or _
                InStr(UCase(Trim(arrOne(j))), "%TYPE") > 0 Then
                strPar = strPar & "," & Trim(arrOne(j))
                Exit For
            End If
        Next
    Next
    GetFuncPars = Mid(strPar, 2)
End Function

Public Function GetLenStr(str As String, lngW As Long, objBase As Object) As String
'功能：根据指定的长度截取字符串
    Dim lngTmp As Long, i As Integer
    
    For i = 1 To Len(str)
        lngTmp = lngTmp + objBase.TextWidth(Mid(str, i, 1))
        If lngTmp <= lngW Then
            GetLenStr = GetLenStr & Mid(str, i, 1)
        Else
            Exit For
        End If
    Next
    If GetLenStr <> str Then
        GetLenStr = Left(GetLenStr, Len(GetLenStr) - 1) & ".."
    End If
End Function

Public Function GetParVBMacro(str As String) As String
'功能:分析报表参数宏,并返回转换后的VB可用值
    Dim curDate As Date
    
    If InStr(str, "&") = 0 Then GetParVBMacro = str: Exit Function
    
    curDate = Currentdate
    Select Case str
        Case "&当前日期"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd")
        Case "&当前日期时间"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd HH:mm:ss")
        Case "&前一周日期"
            GetParVBMacro = Format(curDate - 7, "yyyy-MM-dd")
        Case "&前一月日期"
            GetParVBMacro = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd")
        Case "&前一季日期"
            GetParVBMacro = Format(DateAdd("m", -3, curDate), "yyyy-MM-dd")
        Case "&前一年日期"
            GetParVBMacro = Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd")
        Case "&下一周日期"
            GetParVBMacro = Format(curDate + 7, "yyyy-MM-dd")
        Case "&下一月日期"
            GetParVBMacro = Format(DateAdd("m", 1, curDate), "yyyy-MM-dd")
        Case "&下一季日期"
            GetParVBMacro = Format(DateAdd("m", 3, curDate), "yyyy-MM-dd")
        Case "&下一年日期"
            GetParVBMacro = Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd")
        Case "&当天开始时间"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 00:00:00")
        Case "&当天结束时间"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&前一天同时间"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd HH:mm:ss")
        Case "&后一天同时间"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd HH:mm:ss")
        Case "&本月初时间"
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&本月末时间"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&上月初时间"
            curDate = DateAdd("m", -1, curDate)
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&上月末时间"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&本年初时间"
            GetParVBMacro = Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&本年末时间"
            GetParVBMacro = Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59")
        Case "&上年初时间"
            GetParVBMacro = Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&上年末时间"
            GetParVBMacro = Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59")
    End Select
End Function

Public Function ExeFunction(strOwner As String, strFunc As String, strPars As String, objPars As FuncPars) As String
'功能：执行一个函数
'参数：strOwner=函数的所有者
'      strFunc=函数名
'      strPars=函数的参数描述串,如"NO_IN,Varchar;...",顺序与函数定义一致
'      objPars=函数的参数值对象,当前值存放在"缺省值"中
'返回：正确=函数值；出错=错误信息,以"ERROR"开头
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim arrPar() As String, i As Integer, j As Integer
    Dim StrName As String, strType As String
    On Error GoTo errH
    
    If strPars = "" Then
        strSQL = "Select " & strOwner & "." & strFunc & " as 函数值 From Dual"
    Else
        arrPar = Split(strPars, ";")
        For i = 0 To UBound(arrPar)
            StrName = Split(arrPar(i), ",")(0)
            strType = Split(arrPar(i), ",")(1)
            For j = 1 To objPars.Count
                If UCase(objPars(j).名称) = UCase(StrName) Then Exit For
            Next
            If j <= objPars.Count Then
                '取定义的值
                If UCase(strType) Like "*NUMBER*" Then
                    If objPars(j).缺省值 = "" Then
                        strSQL = strSQL & ",NULL"
                    Else
                        strSQL = strSQL & "," & Val(objPars(j).缺省值)
                    End If
                ElseIf UCase(strType) Like "*CHAR*" Then
                    strSQL = strSQL & ",'" & objPars(j).缺省值 & "'"
                ElseIf UCase(strType) Like "*DATE*" Then
                    If IsDate(objPars(j).缺省值) Then
                        If Not (#1/1/3000# - CDate(objPars(j).缺省值)) Like "*.*" Then
                            strSQL = strSQL & ",To_DATE('" & Format(objPars(j).缺省值, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                        Else
                            strSQL = strSQL & ",To_DATE('" & Format(objPars(j).缺省值, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                    Else
                        strSQL = strSQL & ",NULL"
                    End If
                Else
                    strSQL = strSQL & ",NULL"
                End If
            Else
                '按类型赋缺省值
                If UCase(strType) Like "*NUMBER*" Then
                    strSQL = strSQL & ",1"
                ElseIf UCase(strType) Like "*CHAR*" Then
                    strSQL = strSQL & ",'A'"
                ElseIf UCase(strType) Like "*DATE*" Then
                    strSQL = strSQL & ",To_DATE('" & Format(Currentdate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                Else
                    strSQL = strSQL & ",NULL"
                End If
            End If
        Next
        strSQL = "Select " & strOwner & "." & strFunc & "(" & Mid(strSQL, 2) & ") as 函数值 From Dual"
    End If
    
    On Error Resume Next
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "执行一个函数")
    
    If Err.Number = 0 Then
        ExeFunction = IIf(IsNull(rsTmp!函数值), "", rsTmp!函数值)
    Else
        ExeFunction = "ERROR" & Err.Description
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetFunctionExp(strOwner As String, strFunc As String, strPars As String, objPars As FuncPars) As String
'功能：返回函数的执行公式
'参数：strOwner=函数的所有者
'      strFunc=函数名
'      strPars=函数的参数描述串,如"NO_IN,Varchar;...",顺序与函数定义一致
'      objPars=函数的参数值对象,当前值存放在"缺省值"中
'说明：对于动态时间参数,返回参数格式为"[zlBeginTime]","[zlEndTime]"
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim arrPar() As String, i As Integer, j As Integer
    Dim StrName As String, strType As String
    
    If strPars = "" Then
        strSQL = strOwner & "." & strFunc
    Else
        arrPar = Split(strPars, ";")
        For i = 0 To UBound(arrPar)
            StrName = Split(arrPar(i), ",")(0)
            strType = Split(arrPar(i), ",")(1)
            
            For j = 1 To objPars.Count
                If UCase(objPars(j).名称) = UCase(StrName) Then Exit For
            Next
            
            If j <= objPars.Count Then
                '取定义的值
                If UCase(strType) Like "*NUMBER*" Then
                    If objPars(j).缺省值 = "" Then
                        strSQL = strSQL & ",NULL"
                    Else
                        strSQL = strSQL & "," & Val(objPars(j).缺省值)
                    End If
                ElseIf UCase(strType) Like "*CHAR*" Then
                    strSQL = strSQL & ",'" & objPars(j).缺省值 & "'"
                ElseIf UCase(strType) Like "*DATE*" Then
                    If UCase(StrName) = "ZLBEGINTIME" Or UCase(StrName) = "ZLENDTIME" Then
                        If IsDate(objPars(j).缺省值) Then
                            If Not (#1/1/3000# - CDate(objPars(j).缺省值)) Like "*.*" Then
                                strSQL = strSQL & ",To_DATE('" & Format(objPars(j).缺省值, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            Else
                                strSQL = strSQL & ",To_DATE('" & Format(objPars(j).缺省值, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                        Else
                            strSQL = strSQL & ",[" & StrName & "]"
                        End If
                    Else
                        If IsDate(objPars(j).缺省值) Then
                            If Not (#1/1/3000# - CDate(objPars(j).缺省值)) Like "*.*" Then
                                strSQL = strSQL & ",To_DATE('" & Format(objPars(j).缺省值, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            Else
                                strSQL = strSQL & ",To_DATE('" & Format(objPars(j).缺省值, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                        Else
                            strSQL = strSQL & ",NULL"
                        End If
                    End If
                Else
                    strSQL = strSQL & ",NULL"
                End If
            Else
                '按类型赋缺省值
                If UCase(strType) Like "*NUMBER*" Then
                    strSQL = strSQL & ",1"
                ElseIf UCase(strType) Like "*CHAR*" Then
                    strSQL = strSQL & ",'A'"
                ElseIf UCase(strType) Like "*DATE*" Then
                    If UCase(StrName) = "ZLBEGINTIME" Or UCase(StrName) = "ZLENDTIME" Then
                        strSQL = strSQL & ",[" & StrName & "]"
                    Else
                        strSQL = strSQL & ",To_DATE('" & Format(Currentdate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                    End If
                Else
                    strSQL = strSQL & ",NULL"
                End If
            End If
        Next
        strSQL = strOwner & "." & strFunc & "(" & Mid(strSQL, 2) & ")"
    End If
    GetFunctionExp = strSQL
End Function

Public Function GetFuncName(ByVal strCode As String) As String
'功能：根据函数代码获取函数名
    Dim strTmp As String, blnStart As Boolean
    Dim i As Integer, j As Integer
    
    '移除注释
    strCode = RemoveNote(strCode)
    
    '仅用空格间隔
    strCode = Replace(strCode, vbTab, " ")
    strCode = Replace(strCode, vbCr, " ")
    strCode = Replace(strCode, vbLf, " ")
    
    '求出Begin关键字的开始位置:Begin不会单独作为参数名,函数名
    strTmp = "": blnStart = False: j = 0
    For i = 1 To Len(strCode)
        If Mid(strCode, i, 1) <> " " Then
            blnStart = True
            strTmp = strTmp & Mid(strCode, i, 1)
        ElseIf blnStart Then
            blnStart = False
            If UCase(strTmp) = "BEGIN" Then
                j = i - Len("Begin")
                Exit For
            End If
            strTmp = ""
        End If
    Next
    
    '没有Begin,函数代码错
    If j = 0 Then Exit Function
    
    'Begin前面的代码
    strCode = Trim(Left(strCode, j - 1))
    
    '没有Function,函数代码错
    j = InStr(UCase(strCode), "FUNCTION")
    If j = 0 Then Exit Function
    j = j + Len("FUNCTION")
    
    '函数名开始
    For i = j To Len(strCode)
        If Mid(strCode, i, 1) <> " " Then
            j = i: Exit For
        End If
    Next
    If i > Len(strCode) Then Exit Function
    
    '取函数名
    strTmp = ""
    For i = j To Len(strCode)
        '可以参数的"("紧跟函数名
        If Mid(strCode, i, 1) = " " Or Mid(strCode, i, 1) = "(" Then Exit For
        strTmp = strTmp & Mid(strCode, i, 1)
    Next
    GetFuncName = strTmp
End Function

Public Function FuncOwnerName(ByVal strCode As String, StrName As String, strOwner As String) As String
'功能：在函数代码中函数名前加上所有者
'参数：strName=函数名
    Dim i As Integer, strTmp As String
    
    i = InStr(UCase(strCode), UCase(StrName))
    If i = 0 Then
        strTmp = strCode
    Else
        strTmp = Left(strCode, i - 1) & strOwner & "." & Mid(strCode, i)
    End If
    FuncOwnerName = strTmp
End Function

Public Function GetBalndValue(strSQL As String, strFld As String, strVal As String) As String
'功能：根据参数选择器SQL定义，返回指定绑定值的显示字段及绑定字段的值
'返回：显示值|绑定值
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim strShowFld As String, strBandFld As String
    Dim strShow As String, strBand As String
    Dim strSQLT As String
    
    On Error GoTo errH
    
    For i = 0 To UBound(Split(strFld, "|"))
        strTmp = Split(strFld, "|")(i)
        If Split(strTmp, ",")(2) Like "*&D*" Then
            strShowFld = CStr(Split(strTmp, ",")(0))
        End If
        If Split(strTmp, ",")(2) Like "*&B*" Then
            strBandFld = CStr(Split(strTmp, ",")(0))
        End If
    Next
    strSQLT = Replace(RemoveNote(strSQL), "[*]", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQLT, "根据参数选择器SQL定义")
    
    Do While Not rsTmp.EOF
        strShow = IIf(IsNull(rsTmp.Fields(strShowFld).Value), "", rsTmp.Fields(strShowFld).Value)
        strBand = IIf(IsNull(rsTmp.Fields(strBandFld).Value), "", rsTmp.Fields(strBandFld).Value)
        If strBand = strVal Then Exit Do
        rsTmp.MoveNext
    Loop
    If strShow <> "" Or strBand <> "" Then GetBalndValue = strShow & "|" & strBand
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetDate(strVal As String) As Date
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select " & strVal & " as 日期 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取日期")
    GetDate = rsTmp!日期
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SplitFunc(ByVal strExp As String, strOwner As String, strFunc As String, strPars As String)
'功能：根据函数表达式分解参数
'参数：strExp="ZLCOST.ZL_FUN_GATHER([ZLBEGINTIME],[ZLENDTIME],102,2558)"
'返回：strOwner=函数所有者,strFunc=函数名,strPars=用"|"间隔的原参数,如"[ZLBEGINTIME]|[ZLENDTIME]|102|'张三'"
    Dim i As Integer, intA As Integer, intB As Integer, intSign As Single

    If Trim(strExp) = "" Then Exit Sub
    If Not UCase(strExp) Like "*.ZL_FUN_*" Then Exit Sub
    
    strOwner = UCase(Left(strExp, InStr(strExp, ".") - 1))
    strExp = Mid(strExp, InStr(strExp, ".") + 1)
    If InStr(strExp, "(") = 0 Then
        strFunc = strExp
        strPars = ""
    Else
        strFunc = Left(strExp, InStr(strExp, "(") - 1)
        strExp = Mid(strExp, InStr(strExp, "(") + 1)
        strExp = Left(strExp, Len(strExp) - 1)
        
        intA = 0: intB = 0: intSign = 1
        For i = 1 To Len(strExp)
            
            If Mid(strExp, i, 1) = "(" Then
                intA = intA + 1
            ElseIf Mid(strExp, i, 1) = ")" Then
                intA = intA - 1
            ElseIf Mid(strExp, i, 1) = "'" Then
                intB = intB + intSign
                intSign = -1 * intSign
            End If
            
            If Mid(strExp, i, 1) = "," And intA = 0 And intB = 0 Then
                strPars = strPars & "|"
            Else
                strPars = strPars & Mid(strExp, i, 1)
            End If
        Next
    End If
End Sub

Public Function GetFuncSys(strOwner As String, strFunc As String) As Long
'功能：获取数据库函数所属系统
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 系统 From zlFunctions Where Upper(函数名)='" & UCase(strFunc) & "' And 系统 IN(Select 编号 From zlSystems Where Upper(所有者)='" & UCase(strOwner) & "')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取数据库函数所属系统")
    If Not rsTmp.EOF Then GetFuncSys = rsTmp!系统
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SpaceSQL(ByVal strSQL As String) As String
'功能：将SQL语句变换为只为空格间隔的形式,以便于分析
    Dim i As Long, j As Long, lngB As Long, lngE As Long
    Dim arrSeg() As Variant
                
    strSQL = Replace(strSQL, vbCr, " ")
    strSQL = Replace(strSQL, vbLf, " ")
    strSQL = Replace(strSQL, vbTab, " ")
    
    lngB = -1
    arrSeg = Array()
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "'" Then
            If lngB = -1 Then
                lngB = i
            Else
                ReDim Preserve arrSeg(UBound(arrSeg) + 1)
                arrSeg(UBound(arrSeg)) = lngB & "," & i
                lngB = -1
            End If
        End If
    Next
    If lngB = -1 Then
        For i = 0 To UBound(arrSeg)
            lngB = CLng(Split(arrSeg(i), ",")(0)) + 1
            lngE = CLng(Split(arrSeg(i), ",")(1)) - 1
            For j = lngB To lngE
                If Mid(strSQL, j, 1) = " " Then
                    strSQL = Left(strSQL, j - 1) & Chr(250) & Mid(strSQL, j + 1)
                End If
            Next
        Next
    End If
    
    Do While InStr(strSQL, "  ") > 0
        strSQL = Replace(strSQL, "  ", " ")
    Loop
    
    strSQL = Replace(strSQL, Chr(250), " ")
    
    strSQL = Replace(strSQL, " ,", ",")
    strSQL = Replace(strSQL, ", ", ",")
    SpaceSQL = strSQL
End Function

Public Function zlHomePage(hwnd As Long) As Boolean
'功能：根据产品发行码，联结主页
    Dim strCode As String
    
    strCode = zlRegInfo("支持商URL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlHomePage = True
    End If
End Function

Public Function zlWebForum(hwnd As Long) As Boolean
'功能：根据产品发行码，联结论坛
    Dim strCode As String
    
    'strCode = zlRegInfo("支持商BBS")
    strCode = "www.zlsoft.com/techbbs/index.asp"
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlWebForum = True
    End If
End Function

Public Function zlMailTo(hwnd As Long) As Boolean
'功能：根据产品发行码发送电子邮件
    Dim strCode As String
    strCode = zlRegInfo("支持商MAIL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "mailto:" & strCode, "", "", 1
        zlMailTo = True
    End If
End Function
