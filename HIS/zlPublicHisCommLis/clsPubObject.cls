VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPubObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'创    建:王振涛
'创建时间:2018/9/27
'模块功能:对ZLLIS引用的其他公共部件进行加载 包括  zlreport,
'-------------------1、zl9report 相关的程序----------------------------------------------
'-------------------2、zl9register相关的程序---------------------------------------------
'-------------------3、zl9LisComLib相关的程序--------------------------------------------
'...后续同步扩充
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Option Explicit

'-------------------1、zl9report 相关的程序----------------------------------------------
'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:报表打印打印设置功能，提供给外部调用，提供给外部调用
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function ReportPrintSet(ByVal cnMain As ADODB.Connection, ByVal lngSys _
    As Long, ByVal varReport As Variant, Optional frmParent As Object) As Boolean
1         On Error GoTo ReportPrintSet_Error

2            ReportPrintSet = FunReportPrintSet(cnMain, lngSys, varReport, frmParent)


3         Exit Function
ReportPrintSet_Error:
4         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(ReportPrintSet)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
5         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:报表功能中，打开报表功能，提供给外部调用
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function ReportOpen(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, ParamArray arrPar() As Variant) As Boolean
1         On Error GoTo ReportOpen_Error

          Dim lngCount As Long
          Dim var(30) As Variant
          
2         lngCount = UBound(arrPar)
3         If lngCount > 30 Then
4             Err.Raise -2147483645, , "不支持超过30个参数的报表！"
5             Exit Function
6         End If
7         For lngCount = LBound(arrPar) To UBound(arrPar) - 1
8             var(lngCount) = arrPar(lngCount)
9         Next
10        If UBound(arrPar) > 0 Then
11            var(29) = arrPar(UBound(arrPar))
12        End If
13        If initReport = True Then
14            ReportOpen = FunReportOpen(cnMain, lngSys, varReport, frmParent, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))
15        End If

16        Exit Function
ReportOpen_Error:
17        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(ReportOpen)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
18        Err.Clear
End Function

'编    码:王振涛
'编码时间:2018/9/27
'功    能:保存自定义报表工具的打印设置信息，提供给外部调用
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function SetReportPrintSet( _
    ByVal cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal strReportCode As String, ByVal strKey As String, _
    ByVal strValue As String, Optional ByVal bytType As Byte = 1, Optional ByVal intFormat As Integer = 0) As Boolean

1         On Error GoTo PubSetReportPrintSet_Error

2         SetReportPrintSet = FunSetReportPrintSet(cnOracle, lngSysNo, strReportCode, strKey, strValue, bytType, intFormat)


3         Exit Function
PubSetReportPrintSet_Error:
4         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(PubSetReportPrintSet)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
5         Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/10/10
'功    能:对自定义报表的数据库连接进行初始化
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Sub InitOracle(ByVal cnMain As ADODB.Connection)

    If initReport = True Then
          zlReport.InitOracle cnMain
    End If

End Sub

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/10/10
'功能：根据报表的编号及名称返回合法的密码
'说明：调用该接口前一定注意先调用InitOracle
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function GenReportPass(ByVal strCode As String, ByVal strName As String) As String

    If initReport = True Then
         GenReportPass = zlReport.GenReportPass(strCode, strName)
    End If
End Function


'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/10/10
'功    能:删除一张报表
'入    参:nMain=数据库链接,strNum=报表编号
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function ReportDel(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strNum As String) As Boolean

    If initReport = True Then
         ReportDel = zlReport.ReportDel(cnMain, lngSys, strNum)
    End If

End Function

'---------------------------------------------------------------------------------------
'编    码:李小东
'编码时间:2018/10/31
'功    能:关闭所有子窗口
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    If initReport = True Then
         CloseWindows = zlReport.CloseWindows
    End If
End Function

'---------------------------------------------------------------------------------------
'编    码:李小东
'编码时间:2018/10/31
'功    能:开始打印
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Sub ReportTaskBegin()
    '功能：开始打印，用于多报表在单一任务中打印
    '说明：需要最先调用该函数，再用ReportOpen，且调用ReportOpen时一定要以打印方式
    If initReport = True Then
         zlReport.ReportTaskBegin
    End If
End Sub

'---------------------------------------------------------------------------------------
'编    码:李小东
'编码时间:2018/10/31
'功    能:结束打印
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Sub ReportTaskEnd()
    '功能：结束打印，用于多报表在单一任务中打印
    '说明：调用完ReportOpen之后，最后才调用该函数
    If initReport = True Then
         zlReport.ReportTaskEnd
    End If
End Sub

'-------------------zl9report 相关的程序END----------------------------------------------
'-------------------2、zl9register相关的程序---------------------------------------------
'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:打开指定的数据库,调用register部件
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function GetConnection(ByVal strServer As String, ByVal strUserName As String, ByVal strPassWord As String, ByVal blnTransPassword As Boolean, _
     Optional ByVal bytProvider As Byte = 1, Optional ByRef strError As String = "无须返回错误信息", Optional ByVal blnSaveAccount As Boolean = True) As ADODB.Connection
'功能： 打开指定的数据库，并返回已实例化的ADO连接对象(如果是10.35.10以前的密码，则按新的转换规则更新密码),保存服务器名、用户名和密码到变量gstrServer，gstrUserName，gstrPassword
'参数： strServer       :服务器名，或者可以直接指定IP:Port/SID
'       strUserName     :用户名
'       strPassword     :密码
'       blnTransPassword:是否进行密码转换
'       bytProvider     :打开数据库连接的两种方式,0-msODBC方式,1-OraOLEDB方式
'       strError        :连接失败后，如果指定了此参数，则返回错误信息，未指定时直接弹出提示信息。
'       blnSaveAccount  :保存用户名、密码、服务器名到全局变量（一般，仅在登录调用时保存，供接口ReGetConnection，GetUserName，GetServerName，GetPassword，LoginValidate使用）
'返回： 数据库打开成功，连接对象的状态属性返回adStateOpen(1),失败则返回AdStateClosed(0)

    On Error GoTo GetConnection_Error
    
    Set GetConnection = FunGetConnection(strServer, strUserName, strPassWord, blnTransPassword, bytProvider, strError, blnSaveAccount)

    Exit Function
GetConnection_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "执行(GetConnection)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear
End Function


'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:返回登录导航台时的连接对象，使用rsgister部件
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function ReGetConnection(ByVal bytProvider As Byte, ByRef strError As String, Optional ByRef cnThis As ADODB.Connection) As ADODB.Connection
'功能：返回登录导航台时的同一会话的连接对象，或者重新获取一个指定方式（OLEDB或MSODBC）打开的新会话的连接对象
'参数：bytProvider  :9-登录导航台时的同一会话的连接对象，否则为一个新会话的连接对象：0-msODBC方式,1-OraOLEDB方式,
'      cnThis       :传入该参数时，根据打开该连接对象时缓存的帐号信息，返回一个新会话的连接对象，不传入该参数时，则用登录导航台时的帐号信息返回一个新会话的连接对象
'      strError     :返回打开连接失败后的错误信息,如果传入关键字“无须返回错误信息”，则不弹出错误提示
'返回： 数据库打开成功，连接对象的状态属性返回adStateOpen(1),失败则返回AdStateClosed(0)
    
    On Error GoTo ReGetConnection_Error
    
    Set ReGetConnection = FunReGetConnection(bytProvider, strError, cnThis)

    Exit Function
ReGetConnection_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "执行(ReGetConnection)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:根据服务器名、用户名、密码验证用户登录，使用regitser功能
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function LoginValidate(ByVal strServer As String, ByVal strUserName As String, ByRef strPassWord As String, ByRef strError As String, _
    Optional lngInstance As Long) As Boolean
'功能：根据服务器名、用户名、密码验证用户登录（如果是10.35.10以前的密码，则自动按新的转换规则更新密码）
'参数：strServer    :服务器名，或者可以直接指定IP:Port/SID,如果传入空值，则取登录系统(调用GetConnection函数时)使用的服务器名
'      strUserName  :用户名
'      strPassword  :返回转换后的密码(指定的程序和窗体才返回转换后的，未指定的则返回错误提示信息)
'      strError     :验证失败时返回错误信息
'      lngInstance  :当前应用程序实例的句柄（例如：app.hInstance，如果需要返回转换后的密码，当前没有窗体名，或窗体名不固定时才需要传入）
'返回：验证登录是否成功
    On Error GoTo PubLoginValidate_Error

    LoginValidate = FunLoginValidate(strServer, strUserName, strPassWord, strError, lngInstance)

    Exit Function
PubLoginValidate_Error:
    Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(PubLoginValidate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear
End Function

'--------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功能：获得指定的产品发行或注册授权信息
'参数： strItem-指定的授权项目
'       blnTemp-是否从未保存的临时注册信息验证
'       intBits-对于同时有多项信息的单位名称、产品开发商等指定获得第几个信息,0-N,为-1时表示返回";"间隔的多个
'       cnOracle:用传入的连接来查询
'返回：正确时返回指定的信息；错误返回""
'--------------------------------------------------
Public Function zlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer, Optional ByVal cnOracle As ADODB.Connection) As String
1         On Error GoTo zlRegInfo_Error

2         zlRegInfo = FunzlRegInfo(strItem, blnTemp, intBits, cnOracle)
3         Exit Function

zlRegInfo_Error:
4         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(zlRegInfo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
5         Err.Clear
End Function


'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/10/31
'功    能:初始化注册部件内的连接对象(一般用于登录后调用，仅仅获取一个连接对象时不必初始化)
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Sub zlRegInit(cnMain As ADODB.Connection)
    If initRegister = True Then
        zlRegister.zlRegInit cnMain
    End If
End Sub

'--------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功能：验证系统注册授权的正确性
'参数：blnTemp-是否从未保存的临时注册信息验证
'返回：正确返回"";错误返回错误信息
'--------------------------------------------------
Public Function zlRegCheck(Optional ByVal blnTemp As Boolean, Optional ByVal cnOracle As ADODB.Connection, Optional ByVal blnInit As Boolean) As String
'功能：验证系统注册授权的正确性，并且对当前会话进行认证。（登录时必须调用）
'参数：blnTemp  :是否从未保存的临时注册信息验证（仅用于注册码导入功能）
'      cnOracle :根据传入的连接进行会话认证，否则以部件初始化zlRegInit的连接进行会话认证
'      blnInit  :是否将传入的连接cnOracle用来进行部件初始化zlRegInit
'返回：正确返回"";错误返回错误信息
    
   zlRegCheck = FunzlRegCheck(blnTemp, cnOracle, blnInit)
End Function

'--------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功能：获得授权工具信息
'返回：按2的工具末位次方返回工具许可
'--------------------------------------------------
Public Function ZlRegTool(Optional blnTemp As Boolean, Optional ByVal cnOracle As ADODB.Connection) As Long
1         On Error GoTo ZlRegTool_Error

2         If initRegister = True Then
3             On Error GoTo agin
4             ZlRegTool = zlRegister.ZlRegTool(blnTemp, cnOracle)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo ZlRegTool_Error
7             ZlRegTool = zlRegister.ZlRegTool(blnTemp)
8         End If


9         Exit Function
ZlRegTool_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(ZlRegTool)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
End Function

'--------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功能：返回当前用户具有的指定程序的功能串
'参数：lngSys     如果是固定模块，则为0
'      lngProgId  程序序号
'返回：分号间隔的功能串,为空表示没有权限
'--------------------------------------------------
Public Function GetPrivFunc(lngSys As Long, lngProgId As Long, Optional ByVal cnOracle As ADODB.Connection) As String
1         On Error GoTo GetPrivFunc_Error

2         If initRegister = True Then
3             On Error GoTo agin
4             GetPrivFunc = zlRegister.zlRegFunc(lngSys, lngProgId, cnOracle)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo GetPrivFunc_Error
7             GetPrivFunc = zlRegister.zlRegFunc(lngSys, lngProgId)
8         End If


9         Exit Function
GetPrivFunc_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(GetPrivFunc)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
End Function

Public Function GetUserName(Optional ByVal blnUcase As Boolean = True) As String
'功能：获取当前登录的用户名
'参数：blnUcase     :是否转换为大写字母
       If initRegister = True Then
        GetUserName = zlRegister.GetUserName(blnUcase)
    End If
    
End Function


Public Function GetServerName(Optional ByVal blnUcase As Boolean = True) As String
'功能：获取当前登录的服务器名
'参数：blnUcase     :是否转换为大写字母
       If initRegister = True Then
        GetServerName = zlRegister.GetServerName(blnUcase)
    End If
    
End Function

Public Function GetPassword(Optional lngInstance As Long) As String
'功能：获取当前登录(调用GetConnection之后)用户加密后的密码
'参数：lngInstance-当前应用程序实例的句柄（例如：app.hInstance，没有窗体名，或窗体名不固定时才需要传入）
'说明：仅允许指定的程序和窗口访问本接口(未允许的情况下，返回提示文字，例如：未授权的程序、未授权的窗体)
       If initRegister = True Then
        GetPassword = zlRegister.GetPassword(lngInstance)
    End If
    
End Function

'-------------------zl9register相关的程序END---------------------------------------------

