VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'仅用来代替ZL9Rigister部件
Private mcnOracle           As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Private mcolPrivs           As Collection              '当前用户具备的所有程序的功能权限
Private mlngCount           As Long
Private mstrUserName        As String               '用户名
Private mstrPassword        As String               '密码
Private mstrServer          As String               '服务器
Private mobjRegister        As Object               '10.35.10以上的ZLRigster部件对象，由于该类在管理工具只有一个实例，因此放在类内，若移植到其他不见，请将该定义放在模块
Private mblnEnableTrace     As Boolean              '是否启用跟踪

Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
    OriginalConnection = 9
End Enum

'zlRegCheck,zlRegInfo使用旧的
Public Function zlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
    If mobjRegister Is Nothing Then
        Static srsInfo As New ADODB.Recordset
        Static sblnTemp As Boolean
        Dim strInfo As String, aryInfo() As String
        Dim strSQL As String
        
        err = 0: On Error GoTo errHand
        If blnTemp Or sblnTemp <> blnTemp Or (srsInfo.State <> adStateOpen) Then
            sblnTemp = blnTemp
            strSQL = "Select Item, Text From Table(Cast(zltools.f_Reg_Info(" & IIf(blnTemp, 1, 0) & ") As zlTools.t_Reg_Rowset))"
            If srsInfo.State = adStateOpen Then srsInfo.Close
            srsInfo.Open strSQL, mcnOracle, adOpenKeyset, adLockReadOnly
        End If
        
        srsInfo.Filter = "Item='" & strItem & "'"
        If srsInfo.RecordCount <> 1 Then zlRegInfo = "": Exit Function
        strInfo = "" & srsInfo!Text
        If (strItem = "单位名称" Or strItem = "产品开发商" Or strItem = "技术支持商") And intBits <> -1 Then
            aryInfo = Split(strInfo, ";")
            If intBits > UBound(aryInfo) Then
                strInfo = ""
            Else
                strInfo = aryInfo(intBits)
            End If
        End If
        zlRegInfo = strInfo
        Exit Function
errHand:
        zlRegInfo = ""
    Else
        zlRegInfo = mobjRegister.zlRegInfo(strItem, blnTemp, intBits)
    End If
End Function

'--------------------------------------------------
'功能：验证系统注册授权的正确性
'参数：blnTemp-是否从未保存的临时注册信息验证
'blnPLBUpgrade-仅升级时使用，当低于没有ZLReigster时生效。用来判断PLB是否需要替换
'返回：正确返回"";错误返回错误信息
'--------------------------------------------------
Public Function zlRegCheck(Optional blnTemp As Boolean, Optional ByVal blnPLBUpgrade As Boolean) As String
    If mobjRegister Is Nothing Then
        If blnPLBUpgrade Then
            Dim rsTemp As ADODB.Recordset
            Dim strSQL As String
            Dim arrMd5_1(5) As String
            Dim arrMd5_2(5) As String
            Dim arrMd5_3(5) As String
            Dim arrMd5_4(5) As String
            Dim arrMd5_5(5) As String
            Dim strMD5 As String
            Dim intLine As Integer
            On Error GoTo errHand
            
            '---------------------------------Beging 验证 F_Reg_Audit 是否被替换
            '-- 管理工具 9.25 HIS 10.15
            arrMd5_1(0) = "6746B20191FD2AA9B0E08AFB44E80D4B"
            arrMd5_1(1) = "93C94497A547C10EC3B5C95F5188BA5D"
            arrMd5_1(2) = "A5596EA1AB4F6D4939CBD9599CBFBA0F"
            arrMd5_1(3) = "07069FF5FF76C204EEFCC88366F6A495"
            arrMd5_1(4) = "73C7DB3F742EBC654FAC289B4D37A7B0"
                
            '-- 管理工具 9.35 HIS 10.24
            arrMd5_2(0) = "10E1A9794EF861981C7F53D887990B1F"
            arrMd5_2(1) = "C4A92BE1F6882A57564206E9B391A600"
            arrMd5_2(2) = "F4878F9061BFC4357DC4545EAC326CD2"
            arrMd5_2(3) = "4BBF3E2A0D667A50B8CBC443A1110EA2"
            arrMd5_2(4) = "07BC27215593F6ED86C9905C0D215BD9"
                
            '-- 管理工具 9.37 HIS 10.26
            arrMd5_3(0) = "4D1B31CCB39BDCCE4EE61357555DAD9D"
            arrMd5_3(1) = "F544A3A12A833F6EE10CEA514D65782C"
            arrMd5_3(2) = "5CEF0276B15026C1D5546A85F9A3BE1F"
            arrMd5_3(3) = "487CC8AD6D5F2E0DC337677D02EA702F"
            arrMd5_3(4) = "20AD16738F21A228D962E59DAECB0D84"
            
            '-- 管理工具 9.41 HIS 10.30
            arrMd5_4(0) = "01322819F7B38E12BCAA8525895EF288"
            arrMd5_4(1) = "75E62456DB5F6742B9140DFB73D094FE"
            arrMd5_4(2) = "4270A613EA65B66BF4200BA42F205319"
            arrMd5_4(3) = "64FD2D54E72F9F647DD01D14116988AE"
            arrMd5_4(4) = "D7A22AF77FAC34E04086B800570BCB37"
                
            '-- 管理工具 9.45 HIS 10.34
            arrMd5_5(0) = "01322819F7B38E12BCAA8525895EF288"
            arrMd5_5(1) = "02AC74A017BEE67D26051B4BA5DA98E8"
            arrMd5_5(2) = "9D1143BA317F835426BB8ED2F319A8CA"
            arrMd5_5(3) = "E2718B7863EB402205FAC8CDD348D649"
            arrMd5_5(4) = "39A9E549EAB1EDD396230AD61DC559B0"
            '数据启动，第一次执行，RowNum不是和Line排序对应的，第二次执行以后均正常，因此增加子查询
            strSQL = "Select 源码, Rownum As Line" & vbNewLine & _
                    "From (Select Substr(Text, 1, 512) As 源码" & vbNewLine & _
                    "       From All_Source" & vbNewLine & _
                    "       Where Owner = 'ZLTOOLS' And Name = 'F_REG_AUDIT' And Line In (3, 5, 7, 9, 11)" & vbNewLine & _
                    "       Order By Line)"
        
            Set rsTemp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "zlRegCheck")
            Do Until rsTemp.EOF
                strMD5 = Md5_String_Calc("" & rsTemp!源码)
                intLine = val("" & rsTemp!Line)
                If Not (arrMd5_1(intLine - 1) = strMD5 Or arrMd5_2(intLine - 1) = strMD5 _
                    Or arrMd5_3(intLine - 1) = strMD5 Or arrMd5_4(intLine - 1) = strMD5 _
                    Or arrMd5_5(intLine - 1) = strMD5) Then
                    zlRegCheck = "注册验证程序不正确，请恢复正确的注册函数！"
                    Exit Do
                End If
                rsTemp.MoveNext
            Loop
            If zlRegCheck <> "" Then Exit Function
        End If
        
        '---------------------------------          End  验证 F_Reg_Audit 是否被替换
        If blnTemp Then
            strSQL = "Select zltools.f_Reg_Audit(1) As Stamp From zlRegFile r Where 项目 = '授权证章'"
        Else
            strSQL = "Select zltools.f_Reg_Audit(0) As Stamp From zlRegInfo r Where 项目 = '授权证章'"
        End If
        Set rsTemp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "zlRegCheck", IIf(blnTemp, 1, 0))
        If rsTemp.RecordCount > 0 Then
            If Left(rsTemp.Fields(0).value, 6) <> "ERROR-" Then
                zlRegCheck = ""
            Else
                zlRegCheck = rsTemp.Fields(0).value
            End If
        Else
            zlRegCheck = "无授权信息!"
        End If
        Exit Function
errHand:
        zlRegCheck = err.Description
    Else
        zlRegCheck = mobjRegister.zlRegCheck(blnTemp)
    End If
End Function

Public Function zlRegTool(Optional ByVal blnTemp As Boolean) As Long
'功能：获得授权工具信息
'参数：blnTemp  :是否从未保存的临时注册信息验证
'返回：按2的工具末位次方返回工具许可
    If mobjRegister Is Nothing Then
        Dim rsTool As ADODB.Recordset
        Dim strSQL As String, lngRetu As Long
        
        On Error GoTo errHand
        strSQL = "Select Prog From Table(Cast(zltools.f_Reg_Tool([1]) As zlTools.t_Reg_Rowset))"
        Set rsTool = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "zlRegTool", IIf(blnTemp, 1, 0))
        lngRetu = 0
        Do While Not rsTool.EOF
            lngRetu = lngRetu + 2 ^ ((val("" & rsTool.Fields(0).value) Mod 10) - 1)
            rsTool.MoveNext
        Loop
        zlRegTool = lngRetu
        Exit Function
errHand:
        MsgBox "读取注册工具出错:" & vbNewLine & err.Description, vbExclamation, "zlRegister"
        zlRegTool = 0
    Else
        zlRegTool = mobjRegister.zlRegTool(blnTemp)
    End If
End Function

Public Function zlRegFunc(lngSys As Long, lngProgId As Long) As String
'功能：返回当前用户具有的指定程序的功能串
'参数：lngSys     :如果是固定模块，则为0
'      lngProgId  :程序序号(模块号)
'返回：分号间隔的功能串,为空表示没有权限
    If mobjRegister Is Nothing Then
        Dim rsTmp As ADODB.Recordset, blnNew As Boolean
        Dim strSQL As String, strPrivs As String
        Dim blnRegCheck As Boolean
            
        If mcolPrivs Is Nothing Then
            Set mcolPrivs = New Collection
            blnNew = True
        Else
            On Error Resume Next
            strPrivs = mcolPrivs("_" & lngSys & "_" & lngProgId)
            If err.Number > 0 Then
                blnNew = True
                err.Clear
                On Error GoTo 0
            End If
        End If
        
        If blnNew Then
            strSQL = "Select Text as 功能 From Table(Cast(zltools.f_Reg_Func([1],[2]) as zlTools.t_Reg_Rowset))"
            On Error GoTo errHand
    
            Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "GetPrivFunc", lngSys, lngProgId)
                    
            Do While Not rsTmp.EOF
                strPrivs = strPrivs & ";" & rsTmp!功能
                rsTmp.MoveNext
            Loop
            strPrivs = Mid(strPrivs, 2)
            mcolPrivs.Add strPrivs, "_" & lngSys & "_" & lngProgId
        End If
            
        zlRegFunc = strPrivs
        Exit Function
errHand:
        MsgBox "读取注册功能出错:" & vbNewLine & err.Description, vbExclamation, "zlRegister"
        zlRegFunc = ""
    Else
        zlRegFunc = mobjRegister.zlRegFunc(lngSys, lngProgId)
    End If
End Function

Public Function zlRegBuild(ByVal strFile As String, ByRef pgbState As Object) As Boolean
'功能：将授权注册文件信息写入临时表
'参数：strFile  :注册码文件路径（含文件名）
'      pgbState :进度条控件(ProgressBar)
    If mobjRegister Is Nothing Then
        Dim rsFile As New ADODB.Recordset
        Dim i As Long, strSQL As String
        
        On Error GoTo errHand
        
        With rsFile
            .Open strFile, , adOpenKeyset, adLockReadOnly
            
            pgbState.Visible = True
            pgbState.Max = .RecordCount
                    
            mcnOracle.Execute "Delete zlRegFile"
                    
            strSQL = ""
            i = 1
            Do While Not .EOF
                If i = 1 Then
                    strSQL = "Select '" & !项目 & "'," & !系统 & "," & !序号 & ", '" & !功能 & "' From Dual"
                Else
                    strSQL = strSQL & " Union All Select '" & !项目 & "'," & !系统 & "," & !序号 & ", '" & !功能 & "' From Dual"
                End If
                
                If i = 100 Then
                    strSQL = "Insert Into zlRegFile(项目, 系统, 序号, 功能) " & strSQL
                    mcnOracle.Execute strSQL
                    i = 1
                    strSQL = ""
                Else
                    i = i + 1
                End If
                
                pgbState.value = .Bookmark
                .MoveNext
            Loop
            If strSQL <> "" Then
                strSQL = "Insert Into zlRegFile(项目, 系统, 序号, 功能) " & strSQL
                mcnOracle.Execute strSQL
            End If
        End With
        pgbState.Visible = False
        
        zlRegBuild = True
        Exit Function
errHand:
        MsgBox "注册授权文件写入数据库时出现错误，请检查！" & vbNewLine & err.Description, vbExclamation, "zlRegister"
        pgbState.Visible = False
    Else
        zlRegBuild = mobjRegister.zlRegBuild(strFile, pgbState)
    End If
End Function

Private Sub Class_Initialize()
    mlngCount = mlngCount + 1
    On Error Resume Next
    If mobjRegister Is Nothing Then
        Set mobjRegister = CreateObject("zlRegister.clsRegister")
        err.Clear: On Error GoTo 0
        '当前部件环境是10.35.10以下。
    End If
End Sub

Private Sub Class_Terminate()
    '同一个进程中，在不同部件中,对不同的对象变量进行set xx=nothing，每次都会触发此事件
    mlngCount = mlngCount - 1
    If mlngCount = 0 Then
        Set mcolPrivs = Nothing
        Set mobjRegister = Nothing
        Set mcnOracle = Nothing
    End If
End Sub

Public Sub zlRegInit(cnMain As ADODB.Connection)
'功能：初始化注册部件内的连接对象(一般用于登录后调用，仅仅获取一个连接对象时不必初始化)
    If mobjRegister Is Nothing Then
        Set mcnOracle = cnMain
        
        If mstrUserName = "" Then
            '如果之前没有调用过GetConnection方法，则用户名为空。从连接对象的ConnectionString中解析出连接串
            '这种情况一般用于医保工具等不从导航台登录的程序，在调用zlComlib部件时初始化连接
            Call GetConnectionInfo(cnMain.ConnectionString, mstrServer, mstrUserName, mstrPassword)
        End If
    Else
        mobjRegister.zlRegInit cnMain
    End If
End Sub

Public Function GetConnection(ByVal strServer As String, ByVal strUsername As String, ByVal strPassword As String, ByVal blnTransPassword As Boolean, _
     Optional ByVal bytProvider As enuProvider = MSODBC, Optional ByRef strError As String = "无须返回错误信息", Optional ByVal blnSaveAccount As Boolean = True) As ADODB.Connection
'功能： 打开指定的数据库，并返回已实例化的ADO连接对象(如果是10.35.10以前的密码，则按新的转换规则更新密码)
'参数： strServer       :服务器名，或者可以直接指定IP:Port/SID
'       strUserName     :用户名
'       strPassword     :密码
'       blnTransPassword:是否进行密码转换
'       bytProvider     :打开数据库连接的两种方式,0-msODBC方式,1-OraOLEDB方式
'       strError        :连接失败后，如果指定了此参数，则返回错误信息，未指定时直接弹出提示信息。
'       blnSaveAccount  :保存用户名、密码、服务器名到全局变量（一般，仅在登录调用时保存，供接口ReGetConnection，GetUserName，GetServerName，GetPassword，LoginValidate使用）
'返回： 数据库打开成功，连接对象的状态属性返回adStateOpen(1),失败则返回AdStateClosed(0)
    If mobjRegister Is Nothing Then
        Dim cnOracle As ADODB.Connection
        Dim strPWD_New As String, strPWD_Old As String
        Dim blnPrompt As Boolean
        
        blnPrompt = (strError = "无须返回错误信息")
        strError = ""
        
        If blnTransPassword Then
            strPWD_Old = TransPassword(strPassword)
        Else
            strPWD_Old = strPassword
        End If
        Set cnOracle = OraDBOpen(strServer, strUsername, strPWD_Old, bytProvider, strError, blnSaveAccount)
        
        If blnPrompt And strError <> "" Then
            MsgBox strError, vbExclamation, "提示"
        End If
        
        mstrPassword = strPWD_Old
        Set GetConnection = cnOracle
    Else
        If blnSaveAccount Then
            mstrServer = strServer
            mstrUserName = strUsername
            If blnTransPassword Then
                mstrPassword = TransPassword(strPassword)
            Else
                mstrPassword = strPassword
            End If
        End If
        Set GetConnection = mobjRegister.GetConnection(strServer, strUsername, strPassword, blnTransPassword, bytProvider, strError, blnSaveAccount)
    End If
    If mblnEnableTrace And GetConnection.State = adStateOpen Then
        Call SetSQLTrace(Me.GetServerName, Me.GetUserName, GetConnection)
    End If
End Function


Public Function ReGetConnection(ByVal bytProvider As enuProvider, ByRef strError As String) As ADODB.Connection
'功能：返回登录导航台时的连接对象，或者根据之前打开的数据库连接对象，重新获取一个OLEDB或MSODBC方式打开的连接对象
'参数：bytProvider  :打开数据库连接的两种方式,0-msODBC方式,1-OraOLEDB方式,9-登录导航台时的连接对象
'      strError     :返回打开连接失败后的错误信息
'返回： 数据库打开成功，连接对象的状态属性返回adStateOpen(1),失败则返回AdStateClosed(0)
    If mobjRegister Is Nothing Then
        If bytProvider = OriginalConnection Then
            Set ReGetConnection = mcnOracle
        Else
            Set ReGetConnection = OraDBOpen(mstrServer, mstrUserName, mstrPassword, bytProvider, strError, False)
        End If
    Else
        Set ReGetConnection = mobjRegister.ReGetConnection(bytProvider, strError)
    End If
    If mblnEnableTrace And ReGetConnection.State = adStateOpen Then
        Call SetSQLTrace(Me.GetServerName, Me.GetUserName, ReGetConnection)
    End If
End Function

Public Function LoginValidate(ByVal strServer As String, ByVal strUsername As String, ByRef strPassword As String, ByRef strError As String, _
    Optional lngInstance As Long) As Boolean
'功能：根据服务器名、用户名、密码验证用户登录（如果是10.35.10以前的密码，则自动按新的转换规则更新密码）
'参数：strServer    :服务器名，或者可以直接指定IP:Port/SID,如果传入空值，则取登录系统(调用GetConnection函数时)使用的服务器名
'      strUserName  :用户名
'      strPassword  :返回转换后的密码(指定的程序和窗体才返回转换后的，未指定的则返回错误提示信息)
'      strError     :验证失败时返回错误信息
'      lngInstance  :当前应用程序实例的句柄（如果需要返回转换后的密码，当前没有窗体名，或窗体名不固定时才需要传入）
'返回：验证登录是否成功
    If mobjRegister Is Nothing Then
        Dim cnOracle As ADODB.Connection
        Dim strPWD_New As String, strPWD_Old As String
            
        If strServer = "" Then strServer = mstrServer
        strPWD_New = TransPassword(strPassword)
            
        Set cnOracle = OraDBOpen(strServer, strUsername, strPWD_New, MSODBC, strError, False)
            
        LoginValidate = (cnOracle.State = adStateOpen)
        If LoginValidate Then
            strPassword = strPWD_New
            cnOracle.Close
        End If
    Else
        LoginValidate = mobjRegister.LoginValidate(strServer, strUsername, strPassword, strError, lngInstance)
    End If
End Function


Public Function CreateUser(ByVal cnOracle As ADODB.Connection, ByVal strUsername As String, ByVal strPassword As String, ByRef strError As String) As Boolean
'功能：创建用户并设置密码
'参数： cnOracle        :ADO连接对象
'       strUserName     :用户名
'       strPassword     :密码
'       strError        :创建失败后的错误信息
'返回：是否创建成功
    If mobjRegister Is Nothing Then
        Dim strPWD_New As String, strSQL As String
        
        strPWD_New = TransPassword(strPassword)
        strSQL = "Create User " & strUsername & " identified by """ & strPWD_New & """"
        On Error Resume Next
        cnOracle.Execute strSQL
        
        '用户已存在、密码不符合复杂度要求等错误
        If err.Number <> 0 Then
            strError = err.Description
        End If
        
        CreateUser = err.Number = 0
    Else
        CreateUser = mobjRegister.CreateUser(cnOracle, strUsername, strPassword, strError)
    End If
End Function

Public Function UpdateUserPassword(ByVal cnOracle As ADODB.Connection, ByVal strUsername As String, ByVal strPassword As String, Optional ByVal blnTransPassword As Boolean = True, Optional ByRef strError As String) As Boolean
'功能：修改用户密码(如果是修改自己的密码，则更新之前缓存的密码)
'参数： cnOracle        :ADO连接对象
'       strUserName     :用户名
'       strPassword     :新的密码
'       blnTransPassword:是否转换密码
'       strError        :修改失败后的错误信息
'返回：修改是否成功
    If mobjRegister Is Nothing Then
        strError = UpdatePassword(cnOracle, strUsername, strPassword, blnTransPassword)
            
        If strError = "" And mstrUserName = strUsername Then mstrPassword = strPassword
        UpdateUserPassword = strError = ""
    Else
        UpdateUserPassword = mobjRegister.UpdateUserPassword(cnOracle, strUsername, strPassword, blnTransPassword, strError)
    End If
End Function

Public Function GetUserName(Optional ByVal blnUcase As Boolean = True) As String
'功能：获取当前登录的用户名
'参数：blnUcase     :是否转换为大写字母
    If mobjRegister Is Nothing Then
        GetUserName = IIf(blnUcase, UCase(mstrUserName), mstrUserName)
    Else
        GetUserName = mobjRegister.GetUserName(blnUcase)
    End If
End Function

Public Function GetServerName(Optional ByVal blnUcase As Boolean = True) As String
'功能：获取当前登录的服务器名
'参数：blnUcase     :是否转换为大写字母
    If mobjRegister Is Nothing Then
        GetServerName = IIf(blnUcase, UCase(mstrServer), mstrServer)
    Else
        GetServerName = mobjRegister.GetServerName(blnUcase)
    End If
End Function


Public Function GetPassword(Optional lngInstance As Long) As String
'功能：获取当前登录(调用GetConnection之后)用户加密后的密码
'参数：lngInstance-当前应用程序实例的句柄（没有窗体名，或窗体名不固定时才需要传入）
'说明：仅允许指定的程序和窗口访问本接口(未允许的情况下，返回提示文字，例如：未授权的程序、未授权的窗体)
    If mobjRegister Is Nothing Then
        GetPassword = mstrPassword
    Else
        GetPassword = mobjRegister.GetPassword(lngInstance)
    End If
    If GetPassword Like "未授权的程序:*" And gblnInIDE Then
        GetPassword = mstrPassword
    End If
End Function

Public Sub EnableTrace(ByVal blnEnable As Boolean)
'功能：是否启用跟踪
    mblnEnableTrace = blnEnable
    If Not mcnOracle Is Nothing And blnEnable Then
        If mcnOracle.State = adStateOpen Then
            Call SetSQLTrace(Me.GetServerName, Me.GetUserName, mcnOracle)
        End If
    End If
End Sub

'======================================================================
'私有方法
'======================================================================
Private Function TransPassword(ByVal strPassword) As String
'功能：转换用户密码,密码和用户按大写转换
'参数：
    Dim intBit As Integer, intLen As Integer
    Dim strPWDNew As String, strBit As String
    
    strPassword = UCase(Trim(strPassword))

     If Len(strPassword) = 0 Then TransPassword = "": Exit Function
     
     strPWDNew = ""
     For intBit = 1 To Len(strPassword)
         strBit = Mid(strPassword, intBit, 1)
         Select Case (intBit Mod 3)
         Case 1
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                    strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                    strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                    strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
         Case 2
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                    strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                    strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                    strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
         Case 0
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                    strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                    strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                    strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
         End Select
     Next
    
    TransPassword = strPWDNew
End Function

Private Function OraDBOpen(ByVal strServer As String, ByVal strUsername As String, ByVal strPassword As String, _
    ByVal bytProvider As enuProvider, ByRef strError As String, ByVal blnSaveAccount As Boolean) As ADODB.Connection
'功能： 打开指定的数据库，并返回ADO连接对象
'参数： strServer：服务器名，或者可以直接指定IP:Port/SID
'       strUserName：用户名
'       strUserPwd：密码
'       bytProvider：打开数据库连接的两种方式,0-msODBC方式,1-OraOLEDB方式
'       blnSaveAccount：保存用户名、密码、服务器名到全局变量（一般，登录时调用则保存，验证用户密码时不保存）
'返回： 数据库打开成功，返回true；失败，返回false
    Dim strPersist_Security_Info As String
    Dim arrTmp As Variant, strIp As String, strPort As String, strSID As String
    
    On Error Resume Next

    Set OraDBOpen = New ADODB.Connection
        
    With OraDBOpen
        If InStr(strServer, "/") > 0 Then
            arrTmp = Split(strServer, "/")
            strSID = arrTmp(1)
            If InStr(arrTmp(0), ":") > 0 Then
                arrTmp = Split(arrTmp(0), ":")
                strIp = arrTmp(0)
                strPort = arrTmp(1)
            Else
                strIp = arrTmp(0)
                strPort = "1521"
            End If
            strServer = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIp & ")(PORT=" & strPort & "))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
            
            '下面这种加了ADDRESS_LIST的写法，在ODBC下，只支持SID，不支持SERVICE_NAME;OLEDB则两种都支持
            'If bytProvider = enuProvider.MSODBC Then
            'strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIP & ")(PORT=" & strPort & ")))(CONNECT_DATA=(SID=" & strSID & ")))"
        End If
        
        '当Persist Security Info为false时，连接对象的ConnectionString属性中，不包含密码，MSDataShape方式下甚至不包含服务器名,所以，用模块变量存储，以便获得另一种连接方式时使用
        If blnSaveAccount Then
            mstrServer = strServer
            mstrUserName = strUsername
            mstrPassword = strPassword
        End If
        
        strPersist_Security_Info = ";Persist Security Info=False" '避免调用者从返回的连接对象中获得用户密码，不指定该属性的话，缺省是false
        
        '缺省为adUseServer，如果不指定本句，对于用OLEDB打开的连接，设置Command对象Execute方法返回的Recordset对象的ActiveConnection = Nothing会报错:对象打开时不允许操作(MSODBC方式打开的连接不会报错)
        .CursorLocation = adUseClient
        
        If bytProvider = enuProvider.MSODBC Then
            .Provider = "MSDataShape"
            .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer & strPersist_Security_Info, strUsername, strPassword
        Else
            .Provider = "OraOLEDB.Oracle"
            .Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUsername, strPassword
            'DistribTX=1,允许分布事务(缺省);DistribTx=0:屏蔽分布事务。oracle8.1.7版本有BUG，所以10.35.10之前的管理工具登录时是禁用的。
            'PLSQLRSet=1 用于操作返回游标参数的存储过程，也可写成Extended Properties=PLSQLRSet=1
        End If
    End With
    
    If err = 0 Then
        strError = ""
    Else
        strError = err.Description
        On Error GoTo 0
        
        If InStr(strError, "自动化错误") > 0 Then
            If bytProvider = enuProvider.MSODBC Then
                strError = "msoracl32.dll"
            Else
                strError = "OraOLEDB.dll"
            End If
            strError = "无法创建连接对象，请检查数据访问部件(" & strError & ")是否正常安装并注册。"
        ElseIf InStr(strError, "ORA-12505") > 0 Then
            strError = "ORA-12505,监听程序当前无法识别连接描述符中所给出的 SID,请检查服务名中配置的实例名称。"
            
        ElseIf InStr(strError, "ORA-12170") > 0 Then
            strError = "ORA-12170,连接超时，请检查服务器名是否正确，网络是否可访问，以及是否被服务器防火墙阻止。"
            
        ElseIf InStr(strError, "ORA-12154") > 0 Then
            strError = "ORA-12154,无法分析服务器名，" & vbCrLf & "请检查本机的Oracle配置文件(tnsnames.ora)中是否存在当前使用的服务名。"
            
        ElseIf InStr(strError, "ORA-12541") > 0 Then
            strError = "ORA-12541,无法连接服务器，请检查服务器上的Oracle监听器服务是否启动。"
            
        ElseIf InStr(strError, "ORA-01033") > 0 Then
            strError = "ORA-01033,ORACLE正在初始化或在关闭，请稍候再试。"
            
        ElseIf InStr(strError, "ORA-01034") > 0 Then
            strError = "ORA-01034,ORACLE不可用，请检查数据库实例是否启动。"
            
        ElseIf InStr(strError, "ORA-02391") > 0 Then
            strError = "ORA-02391,用户" & strUsername & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。"
            
        ElseIf InStr(strError, "ORA-01017") > 0 Then
            strError = "ORA-01017,无效的用户名或密码，登录被拒绝。"
        
        ElseIf InStr(strError, "ORA-28000") > 0 Then
            strError = "ORA-28000,该用户已经被禁用，不允许登录。"
        End If
    End If
End Function

Private Function UpdatePassword(ByVal cnOracle As ADODB.Connection, ByVal strUsername As String, ByVal strPassword As String, ByVal blnTransPassword As Boolean) As String
'功能：修改用户密码
'返回：修改出错后的错误描述
    Dim strPWD_New As String
    
    If blnTransPassword Then
        strPWD_New = TransPassword(strPassword)
    Else
        strPWD_New = strPassword
    End If
    
    '当启用密码复杂度要求时，可能有逗号等特殊字符，所以需要加双引号
    '违反密码安全策略等情况会出错
    On Error Resume Next
    cnOracle.Execute "Alter user " & strUsername & " identified by """ & strPWD_New & """"
    
    If err.Number <> 0 Then
       UpdatePassword = err.Description
       err.Clear
    End If
End Function

Private Sub GetConnectionInfo(ByVal strConect As String, ByRef strServerName As String, ByRef strUsername As String, ByRef strPassword As String)
'功能： 分析MSODBC连接对象中的ORACLE串中的 服务器，用户名，密码
'返回： 成功失败，返回True；失败，返回False

    Dim i As Integer
    Dim strTemp As String
    If strConect = "" Then Exit Sub
            
    strServerName = ""
    strUsername = ""
    strPassword = ""
    strConect = Replace(strConect, """", "")
    
    If InStr(strConect, "ODBC") > 0 Then
        'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=DYYY";Persist Security Info=True;User ID=zlhis;Password=his;Data Provider=MSDASQL"
        'Provider=MSDataShape.1;Persist Security Info=False;User ID=ZLHIS;Data Provider=MSDASQL;
        '获取 strServerName(Security为false时，无法获得)
        i = InStrRev(strConect, "Server=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 6)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            End If
        End If
    Else
        'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
        'Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=ZLHIS;Data Source="DYYY"
        i = InStrRev(strConect, "Data Source=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 11)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            Else    'Security为false时，没有;号
                strServerName = strTemp
            End If
        End If
    End If
    
    '获取 strUserName
    i = InStrRev(strConect, "User ID=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 7)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strUsername = Left(strTemp, i - 1)
        End If
    End If
    
    '获取 strPassword
    i = InStrRev(strConect, "Password=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 8)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strPassword = Left(strTemp, i - 1)
        End If
    End If
End Sub

Private Function SetSQLTrace(ByVal strServerName As String, ByVal strUsername As String, ByRef cnOracle As ADODB.Connection) As String
'功能:调用100046事件启动SQL Trace功能
'返回:Trc文件名
    Dim strSQL As String, strLevel As String, strFile As String
    Dim rsTmp As ADODB.Recordset
    
    strServerName = UCase(strServerName)
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSQL = "alter session set timed_statistics=true"
        cnOracle.Execute strSQL
        strSQL = "alter session set max_dump_file_size='100M'"
        cnOracle.Execute strSQL
        If err.Number <> 0 Then err.Clear
        
        '设置Trc文件别名
        strFile = GetTrcFile(strUsername, cnOracle)
        strLevel = "12"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSQL = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        cnOracle.Execute strSQL
        If err.Number = 0 Then
            SetSQLTrace = strFile
            If CheckAndAdjustMustTable("ZLREGINFO", , True) Then    '先检测zlreginfo表是否存在
                strSQL = "Select 1 From zlreginfo Where 项目='TRACE文件'"
                Set rsTmp = cnOracle.Execute(strSQL)
                
                If rsTmp.RecordCount > 0 Then
                    strSQL = "Update zlreginfo Set 内容 ='" & strFile & "' Where 项目='TRACE文件'"
                Else
                    strSQL = "Insert Into zlreginfo (项目,内容) Values ('TRACE文件','" & strFile & "')"
                End If
                cnOracle.Execute strSQL
            
                If err.Number <> 0 Then
                    MsgBox err.Description
                End If
            End If
        End If
    End If
End Function

Private Function GetTrcFile(ByVal strUsername As String, ByRef cnOracle As ADODB.Connection) As String
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strFile As String
        
    On Error Resume Next
    strFile = "ZL_" & strUsername
    strSQL = "alter session set tracefile_identifier='" & strFile & "'"
    cnOracle.Execute strSQL
    If err.Number <> 0 Then     '发生错误,说明设置traceid失败,保存默认的traceFile名称
        strFile = ""
    Else
        strFile = "_" & strFile
    End If
    
    strSQL = "Select Lower(Sys_Context('userenv', 'instance_name')) || '_ora_' || p.Spid || '" & "_" & strFile & ".trc' As Trace_File" & vbNewLine & _
                    "From V$session S, V$process P" & vbNewLine & _
                    "Where s.Paddr = p.Addr And s.Sid = Userenv('sid') And s.Audsid = Userenv('Sessionid')"
    Set rsTmp = gclsBase.OpenSQLRecord(cnOracle, strSQL, "获取TraceFile名称")
    
    If rsTmp.RecordCount > 0 Then
        GetTrcFile = rsTmp!Trace_File
    End If
    
End Function
