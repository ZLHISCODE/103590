Attribute VB_Name = "mdlPubObjcet"
'---------------------------------------------------------------------------------------
'创    建:王振涛
'创建时间:2018/9/27
'模块功能:对ZLLIS引用的其他公共部件进行加载 包括  zlreport,
'-------------------1、zl9report 相关的程序----------------------------------------------
'-------------------2、zl9register相关的程序---------------------------------------------
'-------------------3、zl9LisComLib相关的程序--------------------------------------------
'...后续同步扩充
'---------------------------------------------------------------------------------------
Option Explicit

Public zlReport As Object                                           '报表部件
Public zlRegister  As Object                                        '注册部件zlRegister
Public gobjSample As Object                                         'LIS公共部件中处理报告相关的内容


Public gobjEmr As Object                                            '病历部件
Public gobjEmrInterface As Object                                   '新版电子病历部件
Public gobjPublicLIS As Object                                      'LIS公共接口部件（）

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:
'入    参:
'出    参:
'返    回:
'调整影响:报表 打印打印设置功能
'---------------------------------------------------------------------------------------
Public Function FunReportPrintSetHis(ByVal cnMain As ADODB.Connection, ByVal lngSys _
    As Long, ByVal varReport As Variant, Optional frmParent As Object) As Boolean
1         On Error GoTo ReportPrintSet_Error

2         If initReport = True Then
3            FunReportPrintSetHis = zlReport.ReportPrintSet(cnMain, lngSys, varReport, _
                 frmParent)
4         End If


5         Exit Function
ReportPrintSet_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(ReportPrintSet)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
7         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:报表功能中，打开报表功能
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function FunReportOpenHis(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, ParamArray arrPar() As Variant) As Boolean
1         On Error GoTo ReportOpen_Error

          Dim lngCount As Long
          Dim var(30) As Variant
2          initReport
          
3         lngCount = UBound(arrPar)
4         If lngCount > 30 Then
5             Err.Raise -2147483645, , "不支持超过30个参数的报表！"
6             Exit Function
7         End If
8         For lngCount = LBound(arrPar) To UBound(arrPar) - 1
9             var(lngCount) = arrPar(lngCount)
10        Next
11        If UBound(arrPar) > 0 Then
12            var(29) = arrPar(UBound(arrPar))
13        End If
          
14        FunReportOpenHis = zlReport.ReportOpen(cnMain, lngSys, varReport, frmParent, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))


15        Exit Function
ReportOpen_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(ReportOpen)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
17        Err.Clear
End Function


'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:保存自定义报表工具的打印设置信息
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function FunSetReportPrintSetHis( _
    ByVal cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal strReportCode As String, ByVal strKey As String, _
    ByVal strValue As String, Optional ByVal bytType As Byte = 1, Optional ByVal intFormat As Integer = 0) As Boolean

    If initReport = True Then
        FunSetReportPrintSetHis = zlReport.SetReportPrintSet(cnOracle, lngSysNo, strReportCode, strKey, strValue, bytType, intFormat)
    End If

End Function


'-------------------1、zl9report 相关的程序----------------------------------------------
'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:
'入    参:
'出    参:
'返    回:
'调整影响:初始化 zl9report部件
'---------------------------------------------------------------------------------------
Public Function initReport() As Boolean

1         On Error GoTo initReport_Error

2         If zlReport Is Nothing Then
3             Set zlReport = CreateObject("zl9Report.clsReport")
4         End If
5         initReport = True
6         Exit Function
initReport_Error:
7         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(initReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
8         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:
'入    参:
'出    参:
'返    回:
'调整影响:报表 打印打印设置功能
'---------------------------------------------------------------------------------------
Public Function FunReportPrintSet(ByVal cnMain As ADODB.Connection, ByVal lngSys _
    As Long, ByVal varReport As Variant, Optional frmParent As Object) As Boolean
1         On Error GoTo ReportPrintSet_Error

2         If initReport = True Then
3            FunReportPrintSet = zlReport.ReportPrintSet(cnMain, lngSys, varReport, _
                 frmParent)
4         End If


5         Exit Function
ReportPrintSet_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(ReportPrintSet)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
7         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:报表功能中，打开报表功能
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function FunReportOpen(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, ParamArray arrPar() As Variant) As Boolean
1         On Error GoTo ReportOpen_Error

          Dim lngCount As Long
          Dim var(30) As Variant
2          initReport
          
3         lngCount = UBound(arrPar)
4         If lngCount > 30 Then
5             Err.Raise -2147483645, , "不支持超过30个参数的报表！"
6             Exit Function
7         End If
8         For lngCount = LBound(arrPar) To UBound(arrPar) - 1
9             var(lngCount) = arrPar(lngCount)
10        Next
11        var(29) = arrPar(UBound(arrPar))
          
12        FunReportOpen = zlReport.ReportOpen(cnMain, lngSys, varReport, frmParent, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))


13        Exit Function
ReportOpen_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "执行(ReportOpen)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
15        Err.Clear
End Function


'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:保存自定义报表工具的打印设置信息
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function FunSetReportPrintSet( _
    ByVal cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal strReportCode As String, ByVal strKey As String, _
    ByVal strValue As String, Optional ByVal bytType As Byte = 1, Optional ByVal intFormat As Integer = 0) As Boolean

    If initReport = True Then
        FunSetReportPrintSet = zlReport.SetReportPrintSet(cnOracle, lngSysNo, strReportCode, strKey, strValue, bytType, intFormat)
    End If

End Function
'-------------------1、zl9report相关的程序END----------------------------------------------
'-------------------2、zl9register相关的程序-----------------------------------------------
'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:初始化zlRegister部件
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function initRegister() As Boolean
1         On Error GoTo initRegister_Error

2         If zlRegister Is Nothing Then
3             Set zlRegister = CreateObject("zlRegister.clsRegister")
4         End If
5         initRegister = True
6         Exit Function
initRegister_Error:
7         Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "执行(initRegister)时发生错误,错误号:" & Err.Number & " 出错原因:" & "没有zlRegister部件！" & " 错误行：" & Erl, True)
End Function


'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:打开指定的数据库,调用register部件
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function FunGetConnection(ByVal strServer As String, ByVal strUserName As String, ByVal strPassWord As String, ByVal blnTransPassword As Boolean, _
                                 Optional ByVal bytProvider As Byte = 0, Optional ByRef strError As String = "无须返回错误信息", Optional ByVal blnSaveAccount As Boolean = True) As ADODB.Connection
      '功能： 打开指定的数据库，并返回已实例化的ADO连接对象(如果是10.35.10以前的密码，则按新的转换规则更新密码),保存服务器名、用户名和密码到变量gstrServer，gstrUserName，gstrPassword
      '参数： strServer       :服务器名，或者可以直接指定IP:Port/SID
      '       strUserName     :用户名
      '       strPassword     :密码
      '       blnTransPassword:是否进行密码转换
      '       bytProvider     :打开数据库连接的两种方式,0-msODBC方式,1-OraOLEDB方式
      '       strError        :连接失败后，如果指定了此参数，则返回错误信息，未指定时直接弹出提示信息。
      '       blnSaveAccount  :保存用户名、密码、服务器名到全局变量（一般，仅在登录调用时保存，供接口ReGetConnection，GetUserName，GetServerName，GetPassword，LoginValidate使用）
      '返回： 数据库打开成功，连接对象的状态属性返回adStateOpen(1),失败则返回AdStateClosed(0)

1         On Error GoTo GetConnection_Error
2         If initRegister = True Then
3             If zlRegister.GetUserName = "" And blnSaveAccount = False Then
4                 blnSaveAccount = True
5             End If
6             On Error GoTo agin
7             Set FunGetConnection = zlRegister.GetConnection(strServer, strUserName, strPassWord, blnTransPassword, , strError, blnSaveAccount)
8             Exit Function
agin:
9             Err.Clear: On Error GoTo GetConnection_Error
10            Set FunGetConnection = zlRegister.GetConnection(strServer, strUserName, strPassWord, blnTransPassword, , strError)
11        End If
12        Exit Function
GetConnection_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "执行(GetConnection)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
14        Err.Clear
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
Public Function FunReGetConnection(ByVal bytProvider As Byte, ByRef strError As String, Optional ByRef cnThis As ADODB.Connection) As ADODB.Connection
      '功能：返回登录导航台时的连接对象，或者根据之前打开的数据库连接对象，重新获取一个OLEDB或MSODBC方式打开的连接对象
      '参数：bytProvider  :打开数据库连接的两种方式,0-msODBC方式,1-OraOLEDB方式,9-登录导航台时的连接对象(相同的会话)
      '      strError     :返回打开连接失败后的错误信息
      '     cnThis       :传入该参数时，根据打开该连接对象时缓存的帐号信息，返回一个新会话的连接对象，不传入该参数时，则用登录导航台时的帐号信息返回一个新会话的连接对象
      '返回： 数据库打开成功，连接对象的状态属性返回adStateOpen(1),失败则返回AdStateClosed(0)

1         On Error GoTo ReGetConnection_Error
2         If initRegister = True Then
3             On Error GoTo agin
4             Set FunReGetConnection = zlRegister.ReGetConnection(bytProvider, strError, cnThis)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo ReGetConnection_Error
7             Set FunReGetConnection = zlRegister.ReGetConnection(bytProvider, strError)
8         End If
9         Exit Function
ReGetConnection_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "执行(ReGetConnection)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
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
Public Function FunLoginValidate(ByVal strServer As String, ByVal strUserName As String, ByRef strPassWord As String, ByRef strError As String, _
                                 Optional lngInstance As Long) As Boolean
      '功能：根据服务器名、用户名、密码验证用户登录（如果是10.35.10以前的密码，则自动按新的转换规则更新密码）
      '参数：strServer    :服务器名，或者可以直接指定IP:Port/SID,如果传入空值，则取登录系统(调用GetConnection函数时)使用的服务器名
      '      strUserName  :用户名
      '      strPassword  :返回转换后的密码(指定的程序和窗体才返回转换后的，未指定的则返回错误提示信息)
      '      strError     :验证失败时返回错误信息
      '      lngInstance  :当前应用程序实例的句柄（例如：app.hInstance，如果需要返回转换后的密码，当前没有窗体名，或窗体名不固定时才需要传入）
      '返回：验证登录是否成功
1         On Error GoTo FunLoginValidate_Error
2         If initRegister = True Then
3             On Error GoTo agin
4             FunLoginValidate = zlRegister.LoginValidate(strServer, strUserName, strPassWord, strError, lngInstance)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo FunLoginValidate_Error
7             FunLoginValidate = zlRegister.LoginValidate(strServer, strUserName, strPassWord, strError)
8         End If
9         Exit Function
FunLoginValidate_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "执行(FunLoginValidate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
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
Public Function FunzlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer, Optional ByVal cnOracle As ADODB.Connection) As String
1         On Error GoTo zlRegInfo_Error


2         If initRegister = True Then
3             On Error GoTo agin
4             FunzlRegInfo = zlRegister.zlRegInfo(strItem, blnTemp, intBits, cnOracle)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo zlRegInfo_Error
7             FunzlRegInfo = zlRegister.zlRegInfo(strItem, blnTemp, intBits)
8         End If
9         Exit Function
zlRegInfo_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "执行(FunzlRegInfo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
End Function

'--------------------------------------------------
'功能：获得指定的产品发行或注册授权信息
'参数： strItem-指定的授权项目
'       blnTemp-是否从未保存的临时注册信息验证
'       intBits-对于同时有多项信息的单位名称、产品开发商等指定获得第几个信息,0-N,为-1时表示返回";"间隔的多个
'返回：正确时返回指定的信息；错误返回""
'--------------------------------------------------
Public Function HisZlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
    Static srsInfo As New ADODB.Recordset
    Static sblnTemp As Boolean
    Dim strInfo As String, aryInfo() As String
    Dim strSQL As String
    
    On Error GoTo Errhand
    If blnTemp Or sblnTemp <> blnTemp Or (srsInfo.State <> adStateOpen) Then
        sblnTemp = blnTemp
        strSQL = "Select Item,Text From Table(Cast(zltools.f_Reg_Info([1]) As zlTools.t_Reg_Rowset))"
        Set srsInfo = OpenSQLRecord(Sel_Lis_DB, strSQL, "zlRegInfo", IIf(blnTemp, 1, 0))
    End If
    
    srsInfo.Filter = "Item='" & strItem & "'"
    If srsInfo.RecordCount <> 1 Then HisZlRegInfo = "": Exit Function
    strInfo = "" & srsInfo!Text
    If (strItem = "单位名称" Or strItem = "产品开发商" Or strItem = "技术支持商") And intBits <> -1 Then
        aryInfo = Split(strInfo, ";")
        If intBits > UBound(aryInfo) Then
            strInfo = ""
        Else
            strInfo = aryInfo(intBits)
        End If
    End If
    HisZlRegInfo = strInfo
    Exit Function
Errhand:
    HisZlRegInfo = ""
End Function

'--------------------------------------------------
'功能：获得授权工具信息
'返回：按2的工具末位次方返回工具许可
'--------------------------------------------------
Public Function HisZlRegTool(Optional blnTemp As Boolean) As Long
    Dim rsTool As ADODB.Recordset
    Dim strSQL As String, lngRetu As Long

    On Error GoTo Errhand
    strSQL = "Select Prog From Table(Cast(zltools.f_Reg_Tool([1]) As zlTools.t_Reg_Rowset))"
    Set rsTool = OpenSQLRecord(Sel_Lis_DB, strSQL, "zlRegTool", IIf(blnTemp, 1, 0))
    lngRetu = 0
    Do While Not rsTool.EOF
        lngRetu = lngRetu + 2 ^ ((Val("" & rsTool.Fields(0).value) Mod 10) - 1)
        rsTool.MoveNext
    Loop
    HisZlRegTool = lngRetu
    Exit Function
Errhand:
    HisZlRegTool = 0
End Function

'--------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'返回：正确返回"";错误返回错误信息
'--------------------------------------------------
Public Function FunzlRegCheck(Optional ByVal blnTemp As Boolean, Optional ByVal cnOracle As ADODB.Connection, Optional ByVal blnInit As Boolean) As String
      '功能：验证系统注册授权的正确性，并且对当前会话进行认证。（登录时必须调用）
      '参数：blnTemp  :是否从未保存的临时注册信息验证（仅用于注册码导入功能）
      '      cnOracle :根据传入的连接进行会话认证，否则以部件初始化zlRegInit的连接进行会话认证
      '      blnInit  :是否将传入的连接cnOracle用来进行部件初始化zlRegInit
1         On Error GoTo FunzlRegCheck_Error

2         If initRegister = True Then
3             On Error GoTo agin
4             FunzlRegCheck = zlRegister.zlRegCheck(blnTemp, cnOracle, blnInit)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo FunzlRegCheck_Error
7             FunzlRegCheck = zlRegister.zlRegCheck(blnTemp)
8         End If

9         Exit Function
FunzlRegCheck_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "执行(FunzlRegCheck)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/29
'功    能:通过选择器，从记录集中选择内容并以字符串的方式返回
'集成老的选择器接口。新的选择器如果需要使用，则使用SeletItemFromRsnew方法
'入    参:
'           objfrm              调用选择器的父级对象
'           rsTmpIn             选择器的数据来源
'           strFind             默认过滤条件
'           lngID               默认过滤ID，如果记录集中包含ID字段的话
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function SeletItemFromRsOld(objFrm As Object, ByVal rsTmpIn As Recordset, ByVal strFind As String) As String
    SeletItemFromRsOld = frmPubDicSelOld.ShowMe(objFrm, rsTmpIn, strFind)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/8/16
'功    能:显示公共选择器
'入    参:
'           objfrm      调用来源
'           rsTmp       需要展示的数据来源
'           strFilter   需要过滤的内容
'           lngID       默认过滤ID，如果记录集中包含ID字段的话
'           intShowCol  需要展示多少列数据，从第0列开始依次往后数
'           strHiddenID 需要隐藏的行,多个ID使用","分割
'           blnShowCheckBox     是否显示复选框，若显示复选框，则表示可以多选

'出    参:
'返    回:  选择的内容，每列之间使用“;”分隔
'调整影响:
'---------------------------------------------------------------------------------------
Public Function SeletItemFromRs(objFrm As Object, ByVal rsTmp As ADODB.Recordset, Optional ByVal strFilter As String, _
                               Optional ByVal lngID As String, Optional ByVal intShowCol As Integer = 3, _
                               Optional ByVal strHiddenID As String, Optional ByVal blnShowCheckBox As Boolean) As String
    '文本框在输入中文时（使用输入法打出一个中文串）会多次触发keypress事件，若次方法在该事件中，则会被反复调用会报错，加个判断，如果窗体已经显示，则不再显示
    If frmPubDicSel.Visible = False Then
        SeletItemFromRs = frmPubDicSel.ShowMe(objFrm, rsTmp, strFilter, lngID, intShowCol, strHiddenID, blnShowCheckBox)
    End If
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/29
'功    能:设置vsf控件列头的显示顺序,及显示隐藏列,调用此功能时,必须设置一个参数来保存这些设置
'入    参:
'           VSFlexGrid                     被操作的VSF
'           X                              弹出窗体的X坐标
'           Y                              弹出窗体的Y坐标
'           strPara                        参数名
'           lngSysNo                       系统号
'           lngModlNo                      模块号
'           [strHiddenCols]                固定永远都不显示的列,比如ID,,这些
'           [strShwoCols]                  固定永远都显示的列
'出    参:
'返    回:
'           返回设置之后的列头顺序 , 保存的参数也是这个格式
'           格式:列的key值1,宽度,是否显示(1=显示,0=不显示);列的key值2,宽度,是否显示(1=显示,0=不显示),,,,,,,,
'调整影响:
'---------------------------------------------------------------------------------------
Public Function SetVsfColHiden(objFrm As Object, objVSF As Object, ByVal X As Long, ByVal Y As Long, _
                    ByVal strPara As String, ByVal lngSysNo As Long, ByVal lngModlNo As Long, _
                    Optional ByVal strHiddenCols As String, Optional ByVal strShwoCols As String) As String
                    
        SetVsfColHiden = frmPubColShow.ShowMe(objFrm, objVSF, X, Y, strPara, lngSysNo, lngModlNo, strHiddenCols, strShwoCols)
        '在公共部件中保存了一次，在接口部件中，需要再次保存。
        Call ComSetPara(Sel_Lis_DB, strPara, SetVsfColHiden, lngSysNo, lngModlNo)

End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/29
'功    能:  返回指定字符串的简码
'           根据指定字符串生成简码，可以生成三种类型的简码
'           0、拼音，取每字的首字母构成简码
'           1、五笔，取每字的首字母构成简码
'           2、五笔，按五笔规则构成简码
'           在传入的参数中未发现※符号，就按用户在系统选项中设置的方式生成简码；
'           否则就按在※符号后的数字指定的方式强制生成简码，如※1表示按五笔首字母生成
'入    参:
'           strAsk      序号生成简码的字符串
'出    参:
'返    回:  简码
'调整影响:
'---------------------------------------------------------------------------------------
Public Function SpellCode(ByVal strAsk As String) As String
    SpellCode = gobjHisComLib.zlCommFun.SpellCode(strAsk)
End Function

Public Sub PressKey(bytKey As Byte)
    gobjHisComLib.zlCommFun.PressKey (bytKey)
End Sub
