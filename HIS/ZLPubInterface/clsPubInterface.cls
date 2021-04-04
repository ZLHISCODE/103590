VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPubInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
End Enum
'======================================================================================================================
'接口           GetConnection           验证三方接口的APPID与授权码，成功，则返回的连接为Open状态，否则为Closed状态
'返回值         Connection对象
'入参列表:
'参数名         类型                    说明
'Key            String                  三方接口授权码的明文
'Server         String                  三方接口要连接的服务器，或者可以直接指定IP:Port/SID
'bytProvider    Byte（Optional）        打开数据库连接的两种方式,0-msODBC方式,1-OraOLEDB方式，缺省为0-msODBC方式。(C#版本没有该参数)
'strErrNote     String(Out,Optional)    返回错误信息
'======================================================================================================================
Public Function GetConnection(ByVal strKey As String, ByVal strServer As String, Optional ByVal bytProvider As Byte = MSODBC, Optional ByRef strErrNote As String) As ADODB.Connection
    Dim conZLUA     As ADODB.Connection
    Dim conReturn   As New ADODB.Connection
    Dim strIntPWD   As String
    
    strErrNote = ""
    On Error GoTo errh
    '公共连接创建
    Set conZLUA = OraDBOpen(strServer, "ZLUA", Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)), MSODBC, strErrNote)
    If conZLUA.State = adStateClosed Then
        Set GetConnection = conReturn
        Exit Function
    End If
    '校验授权码
    If Not CheckGrantKey(conZLUA, strKey, strErrNote) Then
        Set GetConnection = conReturn
        Exit Function
    End If
    '获取ZLInterface密码
    strIntPWD = GetZLInterfacePWD(conZLUA, strErrNote)
    If strErrNote <> "" Then
        Set GetConnection = conReturn
        Exit Function
    End If
    '创建ZLInterface连接
    Set conReturn = OraDBOpen(strServer, "ZLINTERFACE", strIntPWD, bytProvider, strErrNote)
    Set GetConnection = conReturn
    Exit Function
errh:
    Set GetConnection = conReturn
    strErrNote = "(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function
'======================================================================================================================
'接口           GetUnitName             验证三方接口的授权码，成功返回单位名称，否则返回空串
'返回值         String
'入参列表:
'参数名         类型                    说明
'Key            String                  三方接口授权码的明文
'varServerOrConnection   String         三方接口要连接的服务器（可以直接指定IP:Port/SID）或者已经打开的连接对象
'strErrNote     String(Out,Optional)    返回错误信息
'======================================================================================================================
Public Function GetUnitName(ByVal strKey As String, ByVal varServerOrConnection As Variant, Optional ByRef strErrNote As String) As String
    Dim connTmp     As ADODB.Connection
    Dim strSQL      As String, rstmp    As ADODB.Recordset
    
    On Error GoTo errh
    strErrNote = ""
    '根据传入的参数类型初始化连接
    If TypeName(varServerOrConnection) = "String" Then
        Set connTmp = GetConnection(strKey, varServerOrConnection, MSODBC, strErrNote)
        If connTmp.State = adStateClosed Then
            Exit Function
        End If
    ElseIf TypeName(varServerOrConnection) = "Connection" Then
        If varServerOrConnection Is Nothing Then
            strErrNote = "传入的连接未实例化。"
            Exit Function
        ElseIf varServerOrConnection.State = adStateClosed Then
            strErrNote = "传入的连接是未打开。"
            Exit Function
        Else
            Set connTmp = varServerOrConnection
            If Not CheckGrantKey(connTmp, strKey, strErrNote) Then
                strErrNote = "接口授权码无效。"
                Exit Function
            End If
        End If
    Else
        strErrNote = "传入的不是String或ADODB.Connection类型。"
        Exit Function
    End If
    strSQL = "Select Max(内容) 内容 From zlRegInfo A Where a.项目  = [1]"
    Set rstmp = OpenSQLRecord(connTmp, strSQL, "GetUnitName", "单位名称")
    If IsNull(rstmp!内容) Then
        strErrNote = "缺失单位名称信息（登录服务器管理工具会自动重建）。"
    Else
        GetUnitName = rstmp!内容 & ""
    End If
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    strErrNote = "(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function
'======================================================================================================================
'接口           GetPassword             验证三方接口的授权码，成功返回用户的登录密码，否则返回空串
'返回值         String
'入参列表:
'参数名         类型                    说明
'Key            String                  三方接口授权码的明文
'varServerOrConnection   String         三方接口要连接的服务器（可以直接指定IP:Port/SID）或者已经打开的连接对象
'strUserName    String                  用户名称
'lngSys         Long                    系统编号，-1缺省从标准版读取。没有读到则从其他系统读取，优先系统编号比较小的
'strErrNote     String(Out,Optional)    返回错误信息
'======================================================================================================================
Public Function GetPassword(ByVal strKey As String, ByVal varServerOrConnection As Variant, ByVal strUserName As String, Optional ByVal lngSys As Long = -1, Optional ByRef strErrNote As String) As String
    Dim connTmp     As ADODB.Connection
    Dim strSQL      As String, rstmp    As ADODB.Recordset
    Dim strOwner    As String
    
    On Error GoTo errh
    strErrNote = ""
    '根据传入的参数类型初始化连接
    If TypeName(varServerOrConnection) = "String" Then
        '公共连接创建
        Set connTmp = OraDBOpen(varServerOrConnection, "ZLUA", Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)), MSODBC, strErrNote)
        If connTmp.State <> adStateOpen Then
            Exit Function
        End If
    ElseIf TypeName(varServerOrConnection) = "Connection" Then
        If varServerOrConnection Is Nothing Then
            strErrNote = "传入的连接未实例化。"
            Exit Function
        ElseIf varServerOrConnection.State <> adStateOpen Then
            strErrNote = "传入的连接是未打开。"
            Exit Function
        Else
            Set connTmp = varServerOrConnection
        End If
    Else
        strErrNote = "传入的不是String或ADODB.Connection类型。"
        Exit Function
    End If
    If Not CheckGrantKey(connTmp, strKey, strErrNote) Then
        Exit Function
    End If
    If lngSys = -1 Then
        strSQL = "Select a.所有者 From zlSystems A Where a.编号 In (100, 300, 400, 600, 2500) Order By a.编号"
    Else
        strSQL = "Select a.所有者 From zlSystems A Where a.编号 =[1]"
    End If
    Set rstmp = OpenSQLRecord(connTmp, strSQL, "获取上机人员表所有者", lngSys)
    If Not rstmp.EOF Then
        strOwner = rstmp!所有者 & ""
        strSQL = "Select Max(登录密码) 登录密码 From " & strOwner & ".上机人员表 A Where a.用户名  = [1]"
        Set rstmp = OpenSQLRecord(connTmp, strSQL, "GetPassword", UCase(strUserName))
        If Not IsNull(rstmp!登录密码) Then
            GetPassword = Sm4DecryptEcb(rstmp!登录密码 & "", GetGeneralAccountKey(G_PASSWORD_KEY))
        End If
    Else
        strErrNote = IIf(lngSys = -1, "当前数据库不存在上机人员表", "当前系统不存在上机人员表")
    End If
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    strErrNote = "(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function
