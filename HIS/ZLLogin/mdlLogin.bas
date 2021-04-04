Attribute VB_Name = "mdlLogin"
Option Explicit

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000 'Forces a top-level window onto the taskbar when the window is visible.强制一个可见的顶级视窗到工具栏上
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

Public gstrSysName As String
Public gdtStart As Long
Public gobjRegister As Object               '注册授权部件zlRegister
Public gcnOracle As ADODB.Connection     '公共数据库连接
Public gstrCommand As String '命令行

Public gobjFile As New FileSystemObject
Public gclsLogin As clsLogin '登录对象
Public gintCallType As Integer '0-不展示修改密码与服务器配置,1-显示修改密码,2-现实服务器配置
Public gblnExitApp  As Boolean '是否因为重复运行，需要退出整个程序

'clsLogin属性缓存
Public gobjEmr             As Object   'EMR新版电子病历
Public gstrUserName        As String   'InputUser属性
Public gstrInputPwd        As String   'InputPwd属性
Public gstrServerName      As String   'ServerName属性
Public gstrDBUser          As String   'DBUser属性
Public gblnTransPwd        As Boolean  'blnTransPwd属性
Public gblnSysOwner        As Boolean  '是否系统所有者
Public gstrConnString      As String   '连接字符串
Public gstrSystems         As String   '多帐套选择的系统
Public gblnCancel          As Boolean  '是否取消退出
Public gstrMenuGroup       As String   '菜单组名称
Public gstrDeptName        As String   '用户登录部门名称
Public gstrStation         As String   '用户登录工作站名称
Public gstrNodeNo          As String   '站点编号
Public gstrNodeName        As String   '站点名称
Public gblnEMRProxy         As Boolean
Public gstrEMRPwd           As String
Public gstrEMRUser          As String

Public gblnTimer            As Boolean  '是否定时器触发的客户端更新检查
Public glngInstanceCount    As Long     '实例计数

Public Sub SetAppBusyState()
'当其他进程对象未创建完成时，替换在执行主进程功能时弹出的“部件被挂起”对话框
    On Error Resume Next
    App.OleServerBusyMsgTitle = App.ProductName
    App.OleRequestPendingMsgTitle = App.ProductName
    
    App.OleServerBusyMsgText = "相关组件正在创建，请耐心等待。"
    App.OleRequestPendingMsgText = "相关组件正创建，请耐心等待。"
    
    App.OleServerBusyTimeout = 3000
    App.OleRequestPendingTimeout = 10000
    Err.Clear
End Sub

Public Function ShowSplash(Optional ByVal blnRefresh As Boolean) As Boolean
    Dim strUnitName As String, intCount As Integer
    Dim objPic As IPictureDisp
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    strUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    If blnRefresh Then
        With frmSplash
            .lblGrant = Replace(strUnitName, ";", vbCrLf)
            .lbl技术支持商.Caption = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
            
            .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
            .lbltag = GetSetting("ZLSOFT", "注册信息", "产品系列", "")
            strUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
            .lbl开发商.Caption = ""
            For intCount = 0 To UBound(Split(strUnitName, ";"))
                .lbl开发商.Caption = .lbl开发商.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
            Next
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            If gobjFile.FileExists(gstrSetupPath & "\附加文件\logo_login.jpg") Then
                Set objPic = LoadPicture(gstrSetupPath & "\附加文件\logo_login.jpg")
                .picHos.Visible = True
                .picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183像素
                .picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323像素
                .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
            Else
                .picHos.Visible = False
            End If
            If InStr(gstrCommand, "=") <= 0 Then .Show
            ShowSplash = True
        End With
    Else
        If strUnitName <> "" And strUnitName <> "-" Then
            gdtStart = Timer
            With frmSplash
                '有两处需要处理
                '此时就开始创建clsComLib类实例
                Call ApplyOEM_Picture(.ImgIndicate, "Picture")
                Call ApplyOEM_Picture(.imgPic, "PictureB")
                If gobjFile.FileExists(gstrSetupPath & "\附加文件\logo_login.jpg") Then
                    Set objPic = LoadPicture(gstrSetupPath & "\附加文件\logo_login.jpg")
                    .picHos.Visible = True
                    .picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183像素
                    .picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323像素
                    .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
                Else
                    .picHos.Visible = False
                End If
                If InStr(gstrCommand, "=") <= 0 Then .Show
                
                .lblGrant = Replace(strUnitName, ";", vbCrLf)
                strUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
                If Trim(strUnitName) = "" Then
                    .Label3.Visible = False
                    .lbl开发商.Visible = False
                Else
                    .Label3.Visible = True
                    .lbl开发商.Visible = True
                    .lbl开发商.Caption = ""
                    For intCount = 0 To UBound(Split(strUnitName, ";"))
                        .lbl开发商.Caption = .lbl开发商.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
                    Next
                End If
                .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
                If Len(.LblProductName) > 10 Then
                    .LblProductName.FontSize = 15.75 '三号
                Else
                    .LblProductName.FontSize = 21.75 '二号
                End If
                .lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
                .lbltag = GetSetting("ZLSOFT", "注册信息", "产品系列", "")
                
                If Trim$(.lbl技术支持商.Caption) = "" Then
                    .Label1.Visible = False
                    .lbl技术支持商.Visible = False
                Else
                    .Label1.Visible = True
                    .lbl技术支持商.Visible = True
                End If
            End With
            Do
                If (Timer - gdtStart) > 1 Then Exit Do
                DoEvents
            Loop
            
            ShowSplash = True
        End If
    End If
End Function

Public Function SaveRegInfo() As Boolean
    Dim strTag As String, strTitle As String
    
    Select Case zlRegInfo("授权性质")
        Case "1"
            '正式
            SaveSetting "ZLSOFT", "注册信息", "Kind", ""
        Case "2"
            '试用
            SaveSetting "ZLSOFT", "注册信息", "Kind", "试用"
        Case "3"
            '测试
            SaveSetting "ZLSOFT", "注册信息", "Kind", "测试"
        Case Else
            '不对
            MsgBox "授权性质不正确，程序被迫退出！", vbInformation, gstrSysName
            Exit Function
    End Select
    
    gstrSysName = zlRegInfo("产品简名") & "软件"
    SaveSetting "ZLSOFT", "注册信息", "提示", gstrSysName
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    strTag = ""
    strTitle = zlRegInfo("产品标题")
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "旗舰版"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "专业版"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    '将用户注册相关信息写入注册表,供下次启动时显示
    SaveSetting "ZLSOFT", "注册信息", "单位名称", zlRegInfo("单位名称", , -1)
    SaveSetting "ZLSOFT", "注册信息", "产品全称", strTitle
    SaveSetting "ZLSOFT", "注册信息", "产品名称", zlRegInfo("产品简名")
    SaveSetting "ZLSOFT", "注册信息", "技术支持商", zlRegInfo("技术支持商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "开发商", zlRegInfo("产品开发商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", zlRegInfo("支持商简名")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", zlRegInfo("支持商MAIL")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", zlRegInfo("支持商URL")
    SaveSetting "ZLSOFT", "注册信息", "产品系列", strTag
    SaveRegInfo = True
End Function

Public Function TestComponent() As Boolean
    '如果没有任何部件可使用，则返回假
    TestComponent = False
    
    Dim strObjs As String, strCodes As String, strSQL As String
    Dim objComponent As Object
    Dim resComponent As New ADODB.Recordset
    
    On Error GoTo errH
    '--由注册表获取授权部件--
    strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")
    If strObjs <> "" Then
        If InStr(strObjs, "'ZL9REPORT'") = 0 Then
            If CreateComponent("ZL9REPORT.ClsREPORT") Then
                strObjs = strObjs & ",'ZL9REPORT'"
                SaveSetting "ZLSOFT", "注册信息", "本机部件", strObjs
            End If
        End If
        TestComponent = True
        Exit Function
    End If
    '--分析授权安装部件--
    strSQL = "Select Distinct 部件 From (" & _
                " Select Upper(g.部件) As 部件" & _
                " From zlPrograms g, zlRegFunc r" & _
                " Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统" & _
                " Union " & _
                " Select Upper(部件) as 部件 From zlPrograms Where 序号 Between 10000 And 19999)"
    Set resComponent = zlDatabase.OpenSQLRecord(strSQL, "")
    With resComponent
        Do While Not .EOF
            If CreateComponent(!部件 & ".Cls" & Mid(!部件, 4)) Then
                strObjs = strObjs & IIf(strObjs = "", "", ",") & "'" & !部件 & "'"
            End If
            .MoveNext
        Loop
    End With
    If strObjs = "" Then Exit Function
    TestComponent = True
    SaveSetting "ZLSOFT", "注册信息", "本机部件", strObjs
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CreateComponent(StrComponent) As Boolean
    Dim objComponent        As Object
On Error GoTo errH
    Set objComponent = CreateObject(StrComponent)
    CreateComponent = True
    Exit Function
errH:
    Err.Clear
    CreateComponent = False
    Exit Function
End Function

Public Function ValEx(ByVal varInput As Variant) As Variant
'功能：由于Val只能以数字开头识别，ValEx以第一个数字进行识别
    Dim arrTmp As Variant, lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
End Function

Public Function CreateRegister() As Boolean
    '创建注册部件(用于登录时获取连接对象)
    On Error Resume Next
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    If gobjRegister Is Nothing Then
        Err.Clear
        MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
        Exit Function
    End If
    CreateRegister = True
End Function

Public Function CheckPWDComplex(ByRef cnInput As ADODB.Connection, ByVal strChcekPWD As String, Optional ByRef strToolTip As String) As Boolean
'功能：检查密码复杂度
'参数：cnInput=传入的连接
'          strChcekPWD=等待检查的密码
'          strToolTip=鼠标提示生成
'返回：True-检查成功；False-检查失败
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim lngLen As Long, i As Integer, intChr As Integer
    
    On Error GoTo errH
    strToolTip = ""
    strSQL = "Select 参数号,Nvl(参数值,缺省值) 参数值 From zlOptions Where 参数号 in (20,21,22,23)"
    rsData.Open strSQL, cnInput
    blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
    blnComplex = False: strOterChrs = ""
    Do While Not rsData.EOF
        Select Case rsData!参数号
            Case 20 '是否控制密码长度
                blnPwdLen = Val(rsData!参数值 & "") = 1
            Case 21 '密码长度下限
                intPwdMin = Val(rsData!参数值 & "")
            Case 22 '密码长度上限
                intPwdMax = Val(rsData!参数值 & "")
            Case 23 '是否控制密码复杂度
                blnComplex = Val(rsData!参数值 & "") = 1
        End Select
        rsData.MoveNext
    Loop
    '生成悬浮提示
    If blnPwdLen Then
        If intPwdMin = intPwdMax Then
            strToolTip = "密码必须为" & intPwdMax & " 位字符。"
        Else
            strToolTip = "密码必须为" & intPwdMin & "至" & intPwdMax & " 位字符。"
        End If
     End If
     If blnComplex Then
        If strToolTip <> "" Then
            strToolTip = strToolTip & vbNewLine & "至少包含一个数字、一个字母与一个特殊字符组成。"
        Else
            strToolTip = "至少由一个数字、一个字母与一个特殊字符组成。"
        End If
     End If
    '长度检查
    lngLen = zlStr.ActualLen(strChcekPWD)
    If lngLen <> Len(strChcekPWD) Then
        MsgBox "新密码包含双字节字符，请检查！", vbInformation, gstrSysName
        Exit Function
    End If
    If blnPwdLen Then
        If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
            If intPwdMin = intPwdMax Then
                MsgBox "密码必须为" & intPwdMax & " 位字符！", vbInformation, gstrSysName
                Exit Function
            Else
                MsgBox "密码必须为" & intPwdMin & "至" & intPwdMax & " 位字符！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    For i = 1 To Len(strChcekPWD)
        intChr = Asc(UCase(Mid(strChcekPWD, i, 1)))
        If intChr >= 32 And intChr < 127 Then
            'Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
            Select Case intChr
                Case 48 To 57 '数字
                    blnHaveNum = True
                Case 65 To 90 '字母
                    blnAlpha = True
                Case 32, 34, 47, 64  '空格,双引号,/,@
                    strOterChrs = strOterChrs & Chr(intChr)
                Case Is < 48, 58 To 64, 91 To 96, Is > 122
                    blnChar = True
            End Select
        Else
            strOterChrs = strOterChrs & Chr(intChr)
        End If
    Next
    If strOterChrs <> "" Then
        MsgBox "密码不容许有以下字符：" & strOterChrs, vbInformation, gstrSysName
        Exit Function
    ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
        MsgBox "密码至少由一个数字、一个字母与一个特殊字符组成。", vbInformation, gstrSysName
        Exit Function
    End If
    CheckPWDComplex = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function CheckSysState() As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnHaveTools As Boolean, blnDBA As Boolean
    
    On Error Resume Next
    strSQL = "SELECT 1 FROM ZLTOOLS.ZLSYSTEMS WHERE 所有者=USER"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取所有者")
    
    If Err.Number <> 0 Then
        blnHaveTools = False
        gclsLogin.IsSysOwner = False
        Err.Clear
    Else
        blnHaveTools = True
        gclsLogin.IsSysOwner = rsTmp.EOF
    End If

    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "判断DBA")
    blnDBA = Not rsTmp.EOF

    If Not (blnDBA) And Not (blnHaveTools) Then
        CheckSysState = False
        MsgBox "尚创建服务器管理工具，请先进行创建！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If Not (blnDBA) And Not (gclsLogin.IsSysOwner) Then
        CheckSysState = False
        MsgBox "不是数据库DBA或应用系统的所有者，不能使用本工具。", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not blnHaveTools Then
        CheckSysState = False
        MsgBox "尚创建服务器管理工具，请先进行创建！", vbExclamation, gstrSysName
        Exit Function
    End If
    CheckSysState = True
End Function

Public Function GetMenuGroup(ByVal strCommand As String) As String
    Dim ArrCommand As Variant
    '--分析权限菜单--
    If strCommand = "" Then
        GetMenuGroup = "缺省"
    Else
        ArrCommand = Split(gstrCommand, " ")
        If UBound(ArrCommand) = 0 Then
            '仅仅包含菜单组别（如果含有/，表示是用户加密码的格式，如：zlhis/his）
            If InStr(1, ArrCommand(0), "/") = 0 And InStr(ArrCommand(0), ",") = 0 Then
                GetMenuGroup = ArrCommand(0)
            Else
                GetMenuGroup = "缺省"
            End If
        Else
            '用户名、密码及菜单组别
            If UBound(ArrCommand) = 2 And InStr(ArrCommand(0), "=") <= 0 Then
                GetMenuGroup = ArrCommand(2)
            Else
                GetMenuGroup = "缺省"
            End If
        End If
    End If
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
        
    On Error GoTo errHand
    
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
errHand:
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
