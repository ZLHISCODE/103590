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
'==================================================================================================
'编写           lshuo
'日期           2018/12/25
'模块           clsRegister
'说明
'==================================================================================================
Private Const mstrCurModule     As String = "clsRegister"           '当前模块名称
'仅用来代替ZL9Rigister部件
Private mcnOracle           As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Private mstrUserName        As String               '用户名
Private mstrPassword        As String               '密码
Private mstrServer          As String               '服务器
Private mobjRegister        As Object               '10.35.10以上的ZLRigster部件对象，由于该类在管理工具只有一个实例，因此放在类内，若移植到其他不见，请将该定义放在模块

Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
    OriginalConnection = 9
End Enum


Private Sub Class_Initialize()
    On Error Resume Next
    Set mobjRegister = CreateObject("zlRegisterAlone.clsRegister")
    Err.Clear
    If mobjRegister Is Nothing Then
        Set mobjRegister = CreateObject("zlRegister.clsRegister")
        Err.Clear: On Error GoTo 0
        '当前部件环境是10.35.10以下。
    End If
End Sub

Private Sub Class_Terminate()
    Set mobjRegister = Nothing
    Set mcnOracle = Nothing
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
    If GetPassword Like "未授权的程序:*" And IsDesinMode Then
        GetPassword = mstrPassword
    End If
End Function

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
    
    If Err = 0 Then
        strError = ""
    Else
        strError = Err.Description
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
    
    If Err.Number <> 0 Then
       UpdatePassword = Err.Description
       Err.Clear
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

Public Function OpenSQLRecord(ByVal cnOracle As ADODB.Connection, ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant
    arrPars = arrInput
    Set OpenSQLRecord = OpenSQLRecordByArray(cnOracle, strSQL, strTitle, arrPars)
End Function

Public Function OpenSQLRecordByArray(ByVal cnOracle As ADODB.Connection, ByVal strSQL As String, ByVal strTitle As String, arrInput() As Variant, Optional intLobOprate As Integer = 0) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'      intLobOprate=0:普通SQL,1:LOB类型读取SQL,2:LOB保存SQL
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
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
    
    '执行返回记录集
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = cnOracle '这句比较慢(这句执行1000次约0.5x秒)
    'End If

    cmdData.CommandText = strSQL
    If intLobOprate > 0 Then '保存LOB,读取LOB也要使用该参数，否则很慢，约10倍差距
        Set OpenSQLRecordByArray = New ADODB.Recordset
        OpenSQLRecordByArray.Open cmdData, , adOpenStatic, adLockOptimistic
    Else
        Set OpenSQLRecordByArray = cmdData.Execute
    End If
    If intLobOprate < 2 Then
        Set OpenSQLRecordByArray.ActiveConnection = Nothing
    End If
End Function
