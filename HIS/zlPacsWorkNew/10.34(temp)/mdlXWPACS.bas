Attribute VB_Name = "mdlXWPACS"
Option Explicit

Public gcnXWDBServer As New ADODB.Connection         '公共数据库连接
Public gfrmPacsMain As frmPacsMain                  '用来接收报告图消息的窗体指针
Public glngXWDeptID As Long                         '当前科室ID
Public gblnXWMoved As Boolean                       '是否转储
Public gblnXWLog As Boolean                         '是否记录通讯日志
Public gstrOracleOwner As String                    'oracle包的拥有者
Public gblnUseXinWangPacs As Boolean                '判断是否配置了新网观片
Public gstrImageShareDir As String                  '老版的影像共享存储目录

'调用新网提供的 InterCOM.dll，启动ADViewer查看图像

'函数声明
Public Declare Function OEMViewStart Lib "InterCOM.dll" (ByVal cpReserved1 As String, ByVal cpReserved2 As String, ByVal cpReserved3 As String) As Long
Public Declare Function OEMViewExit Lib "InterCOM.dll" (ByVal cpReserved1 As String, ByVal cpReserved2 As String, ByVal cpReserved3 As String) As Long
Public Declare Function OEMViewOpen Lib "InterCOM.dll" (ByVal lPlanID As Long, ByVal cpFilter As String, ByVal lFunc As Long, ByVal cpReserved As String) As Long
Public Declare Function OEMViewClose Lib "InterCOM.dll" (ByVal cpReserved As String) As Long
 
'新网提供的函数
'启动ADViewer： long OEMViewStart ( LPCTSTR cpReserved1, LPCTSTR cpReserved2, LPCTSTR cpReserved3 );
'退出ADViewer： long OEMViewExit ( LPCTSTR cpReserved1, LPCTSTR cpReserved2, LPCTSTR cpReserved3 );
'打开指定图像： long OEMViewOpen ( long lPlanID, LPCTSTR cpFilter, long lFunc, LPCTSTR cpReserved );
'关闭图像： long OEMViewClose ( LPCTSTR cpReserved );


'接收报告图像消息的API
Public Const WM_XWReportImage As Long = 5120
'消息Hook变量
Public plngXWPreWndProc As Long       '原来的消息处理程序


'-----------------------------------------------------------------------------------------------------
'ADViewer函数调用
'-----------------------------------------------------------------------------------------------------

Function XWADViewerStart() As Long
'--------------------------------------------
'功能： 启动ADViewer
'       该函数通常只需要运行一次。虽然打开图像时如果ADViewer 会自动启动，
'       但是为了加快执行速度，建议第三方软件在启动时执行此函数，以同时启动ADViewer
'参数：无
'返回：
'--------------------------------------------
    'OEMViewStart 的三个参数 cpReserved1、cpReserved2、cpReserved3：均为保留，固定为NULL
    
    On Error GoTo err
    
    XWADViewerStart = OEMViewStart("", "", "")
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Function XWADViewerExit() As Long
'--------------------------------------------
'功能： 退出ADViewer
'
'参数：无
'返回：
'--------------------------------------------
    'XWViewerExit 的三个参数 cpReserved1、cpReserved2、cpReserved3：均为保留，固定为NULL
    
    On Error GoTo err
    
    XWADViewerExit = OEMViewExit("", "", "")
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Function XWADViewerOpen(ByVal strFilter As String, ByVal lngPlanId As Long) As Long
'--------------------------------------------
'功能： 打开指定图像。条件由参数指定，且必须与配置文件中的配置相符。
'       图像打开时与ADViewer 当前模式有关，如果是单记录模式，则软件会自动关闭原来的图像；如果是对比
'       模式，则会添加到ADViewer中。
'参数：
'       lngOrderID -- 医嘱ID
'返回：
'--------------------------------------------
    Dim strRev As String
    Dim lngFunction As Long
    Dim strXwPrivs As String
    
    'XWViewerOpen 参数说明：
    'lPlanID：  方案ID。该ID 必须与INI 文件中一致，在简单网络的情况下，通常该值为1，建议把该ID 做为一配置项，调用时读取该项并传入。
    'cpFilter： 该代表要打开图像的条件值。例如检查号、申请号等，可以传入多个值，
    '           不同值之间用分隔符[;]隔开，该参数的意义及顺序在INI 文件中配置，并且与lPlanID对应。
    'lFunc：    功能权限。每一位代表一项功能，如果具有多项权限，按位“或”即可，具体功能意义:
    '           0x00000002： 重建图像保存，例如：减影后图像、拼接图像等
    '           0 x00000200: 胶片打印
    '           0 x00040000: 图像导出?另存为其他格式
    '           0 x00080000: GSPS 保存
    'cpReserved：   保留，设为NULL
    
    On Error GoTo err
    
    '记录接口日志
    If gblnXWLog = True Then
        Call WriteCommLog("XWADViewerOpen", "XW接口", "打开ADViewer，并显示图像，医嘱ID= " & strFilter)
    End If
    
    '根据RIS中的权限，组织权限串
    lngFunction = 0
    strXwPrivs = GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    If InStr(strXwPrivs, "PACS保存重建图像") <> 0 Then
        lngFunction = lngFunction Or &H2
    End If
    
    If InStr(strXwPrivs, "PACS胶片打印") <> 0 Then
        lngFunction = lngFunction Or &H200
    End If
    
    If InStr(strXwPrivs, "PACS图像导出") <> 0 Then
        lngFunction = lngFunction Or &H40000
    End If
    
    If InStr(strXwPrivs, "PACS GSPS保存") <> 0 Then
        lngFunction = lngFunction Or &H80000
    End If
        
    XWADViewerOpen = OEMViewOpen(lngPlanId, strFilter, lngFunction, "")
    
    If XWADViewerOpen <> 0 Then
        MsgBox "ADViewer打开错误，返回的信息是：" & XWADViewerOpen
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Function XWADViewerClose() As Long
'--------------------------------------------
'功能： 关闭图像，不退出ADViewer
'
'参数：无
'返回：
'--------------------------------------------
    'XWViewerClose 的参数 cpReserved1为保留，固定为NULL
    On Error GoTo err
    
    XWADViewerClose = OEMViewClose("")
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As ADODB.Connection
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSql As String
    Dim strError As String
    Dim cnOra As New ADODB.Connection
    
    On Error Resume Next
    err = 0
    
    DoEvents
    
    With cnOra
        If .State = adStateOpen Then .Close
        
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        
        If err <> 0 Then
            '保存错误信息
            strError = err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            'OraDataOpen = Nothing
            Exit Function
        End If
    End With
    
    err = 0
    On Error GoTo errHand
    
    Set OraDataOpen = cnOra
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function


Sub XWTestDBConnection(ByVal strServerName As String, ByVal strUser As String, ByVal strPwd As String)
'功能： 测试新网SQLServer数据库连接
'参数：
'返回：成功返回空字符
'--------------------------------------------
    Dim cnTest As New ADODB.Connection

    If strServerName = "" Then
        MsgBox "未找到数据库服务器配置信息，请设置。"
        Exit Sub
    End If
    
    On Error Resume Next
    err = 0
    
    If cnTest.State = adStateOpen Then cnTest.Close
    
    Set cnTest = OraDataOpen(strServerName, strUser, strPwd)
    
    If err <> 0 Or cnTest Is Nothing Then
        '数据库连接错误
        MsgBox "数据库连接失败。" & vbCrLf & vbCrLf & "错误代码是：" & err.Number & "；错误描述是： " & err.Description
        Exit Sub
    End If
    
    MsgBox "数据库连接成功。"
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


'----------------------------------------------------------------------------------------------
'新网SQLSERVER数据库连接和关闭

'----------------------------------------------------------------------------------------------

Public Function XWDBServerOpen() As Long
'--------------------------------------------
'功能： 连接新网SQLServer数据库
'参数：无
'返回：0-成功
'--------------------------------------------
    Dim strSqlUser As String
    Dim strSqlPWD As String
    Dim strDataSource As String

    gblnUseXinWangPacs = False
    
    If InStr(GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE), "基本") <= 0 Then Exit Function
    
    '从中联ORACLE 模块参数中获取新网的数据库服务器IP地址，用户名和密码
    strDataSource = zlDatabase.GetPara("XW数据库服务器IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    strSqlUser = zlDatabase.GetPara("XW数据库服务器用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    strSqlPWD = zlDatabase.GetPara("XW数据库服务器密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    gstrImageShareDir = zlDatabase.GetPara("XW历史图像共享目录", glngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    
    If strDataSource = "" Then
        MsgBox "未找到SQLSERVER数据库服务器，请在“影像RIS工作站”的PACS参数中设置。"
        XWDBServerOpen = 1
        Exit Function
    End If

    On Error Resume Next
    err = 0
    If gcnXWDBServer.State = adStateOpen Then gcnXWDBServer.Close
    
    Set gcnXWDBServer = OraDataOpen(strDataSource, strSqlUser, strSqlPWD)
    
    If err <> 0 Or gcnXWDBServer Is Nothing Then
        '数据库连接错误
        MsgBox "DBServer数据库连接错误，可能会导致部分图像无法查看。" & vbCrLf & vbCrLf & "错误代码是：" & err.Number & "；错误描述是： " & err.Description
    End If
    
    gblnUseXinWangPacs = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Function XWDBServerClose() As Long
'--------------------------------------------
'功能： 关闭新网SQLServer数据库连接
'参数：无
'返回：0-成功
'--------------------------------------------
    On Error GoTo err
    
    If gcnXWDBServer Is Nothing Then Exit Function
    
    If gcnXWDBServer.State = adStateOpen Then gcnXWDBServer.Close
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


'-------------------------------------------------------------------------------------------------------
'ADViewer查看图像应用函数
'-------------------------------------------------------------------------------------------------------

Public Function XWShowImage(ByVal lngViewerType As Long, ByVal strFilter As String, Optional ByVal lngPlanId As Long = 1) As Long
''--------------------------------------------
''功能： 打开新网的ADViewer或者WEB Viewer
''参数：    lngViewerType -- 打开Viewer的方式；1-放射科ADViewer；2-临床WEB Viewer
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------
    On Error GoTo err
    
    '记录接口日志
    If gblnXWLog = True Then
        Call WriteCommLog("XWShowImage", "XW接口", "调用ADViewer或者WEB观片，观片方式是： " & IIf(lngViewerType = 1, "ADViewer", "WEB"))
    End If
    
    If lngViewerType = 1 Then
        Call XWADViewerOpen(strFilter, lngPlanId)
    ElseIf lngViewerType = 2 Then
        Call XWWebViewerOpen(strFilter)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function XWWebViewerOpen(ByVal lngOrderID As Long) As Long
''--------------------------------------------
''功能： 打开新网的WEB Viewer
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------
    Dim strIP As String
    Dim strURL As String
    
    On Error GoTo err
    
    strIP = zlDatabase.GetPara("XWWEB服务器IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    If strIP <> "" Then
        strURL = "C:\Program Files\Internet Explorer\iexplore.exe http://" & strIP & ":8080/imageweb/imageAction.action?ColID0=22&ColValue0=" & lngOrderID
        
        '记录接口日志
        If gblnXWLog = True Then
            Call WriteCommLog("XWWebViewerOpen", "XW接口", "通过WEB方式观片： " & strURL)
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        '记录接口日志
        If gblnXWLog = True Then
            Call WriteCommLog("XWWebViewerOpen", "XW接口", "通过WEB方式观片：WEB服务器IP地址为空。")
        End If
        
        MsgBox "WEB服务器IP地址为空，请先设置好WEB服务器。", vbOKOnly, "提示信息"
        
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function XWShowUnMatched(frmParent As Form, lngOrderID As Long, strModality As String) As Long
''--------------------------------------------
''功能： 连接新网SQLServer数据库，调用V_OEM_STUDY_UNMATCHED视图，显示未匹配的记录
''参数：    frmParent -- 父窗体
'           lngOrderID -- 需要关联检查的医嘱ID
'           strModality --- 影像类别
''返回：0-成功;1-未关联;2-出错
''--------------------------------------------
    Dim lngXWStudyID As Long    '从新网数据库中读取出来的检查主键
    Dim strSql As String
    Dim rsOrderInfo As ADODB.Recordset
    Dim strStudyDate As String
    Dim blnOpenDB As Boolean
    
    On Error GoTo err
    
    '判断数据库是否已经连接，如果没有连接，则打开连接
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            blnOpenDB = True
        End If
    End If
    
    XWShowUnMatched = 2
    
    '显示未匹配的记录
    lngXWStudyID = frmXWRelateImage.zlShowMe(frmParent, lngOrderID, True, strModality)
    If lngXWStudyID > 0 Then
        strStudyDate = frmXWRelateImage.pstrStudyDate
        '使用这个医嘱ID进行图像关联
        
        strSql = "Select b.病人ID,b.门诊号,b.住院号,b.健康号 as 体检号,b.姓名,b.性别,b.年龄,To_char(b.出生日期,'yyyymmdd') As 出生日期, " _
                    & " c.英文名 as 拼音名,c.影像类别,c.检查号,a.病人来源,a.执行科室ID,d.名称 As 执行科室,a.开嘱时间,a.开始执行时间 " _
                    & " From 病人医嘱记录 a,病人信息 b,影像检查记录 c,部门表 d  " _
                    & " Where a.病人Id = b.病人ID And a.Id = c.医嘱ID And a.执行科室ID =d.Id  and a.Id = [1]"
        Set rsOrderInfo = zlDatabase.OpenSQLRecord(strSql, "查询检查信息", lngOrderID)
        
        If rsOrderInfo.RecordCount <> 0 Then
            '调用新网存储过程“P_OEM_MATCHING_RIS”，关联图像
                    
            strSql = "P_OEM_MATCHING_RIS(" & lngXWStudyID & ",'" & lngOrderID & "','" & rsOrderInfo!病人ID & "','" & Nvl(rsOrderInfo!门诊号, 0) _
                    & "','" & Nvl(rsOrderInfo!住院号, 0) & "','" & Nvl(rsOrderInfo!体检号, 0) & "','" & Nvl(rsOrderInfo!姓名) & "','" _
                    & Nvl(rsOrderInfo!性别) & "','" & Nvl(rsOrderInfo!年龄, 0) & "','" & Nvl(rsOrderInfo!出生日期) & "','" & Nvl(rsOrderInfo!拼音名) _
                    & "','" & Nvl(rsOrderInfo!影像类别) & "'," & rsOrderInfo!检查号 & "," & Nvl(rsOrderInfo!病人来源, 3) & "," & Nvl(rsOrderInfo!执行科室ID) _
                    & ",'" & Nvl(rsOrderInfo!执行科室) & "','','')"
                    
            gcnXWDBServer.Execute strSql
            
            '调用中联存储过程"b_XINWANGInterface.PacsStatusChange"，关联图像
            strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsStatusChange(1," & lngOrderID & ",'" & Nvl(rsOrderInfo!影像类别) & "'," & rsOrderInfo!检查号 & ",to_date('" _
                        & Trim(strStudyDate) & "','YYYY.MM.DD'),null,null)"
            zlDatabase.ExecuteProcedure strSql, "关联图像"
        End If
        XWShowUnMatched = 0
        
    ElseIf lngXWStudyID = -1 Then
        '图像关联修复
        XWShowUnMatched = 0
        
    Else
        XWShowUnMatched = 1
    End If
    
    '如果是在过程中打开的数据库连接，则退出时关闭连接
    If blnOpenDB = True Then
        Call XWDBServerClose
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function XWShowMatched(frmParent As Form, lngOrderID As Long) As Long
''--------------------------------------------
''功能： 连接新网SQLServer数据库，调V_OEM_SERIES视图，显示已匹配的记录
''参数：    frmParent -- 父窗体
'           lngOrderID -- 需要关联检查的医嘱ID
''返回：0-成功;1-未取消关联；2-出错
''--------------------------------------------
    Dim lngXWStudyID As Long    '从新网数据库中读取出来的检查主键
    
    On Error GoTo err
    
    XWShowMatched = 2
    
    '显示已匹配的记录
    lngXWStudyID = frmXWRelateImage.zlShowMe(frmParent, lngOrderID, False, "")
    If lngXWStudyID <> 0 Then
        '使用这个医嘱ID取消图像关联
        '使用这个lngXWStudyID在新网数据库中取消关联
        
        If MsgBoxD(frmParent, "是否确认取消图像和检查信息的关联？", vbOKCancel, "提示信息") = vbCancel Then
            Exit Function
        End If
        
        XWShowMatched = XWUnmatchImage(lngOrderID, lngXWStudyID)
    Else
        XWShowMatched = 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function XWUnmatchImage(lngOrderID As Long, lngXWStudyID As Long) As Long
''--------------------------------------------
''功能： 连接新网SQLServer数据库，调用P_OEM_UNMATCHING_RIS过程取消指定记录的关联
''参数：    lngOrderID -- 需要取消关联检查的医嘱ID
''          lngXWStudyID -- 需要取消关联检查的新网检查号，0表示删除医嘱ID下的所有检查
''返回：0-成功;1-未取消关联；2-出错
''--------------------------------------------
    Dim blnOpenDB As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset

    XWUnmatchImage = 1
    
    '判断数据库是否已经连接，如果没有连接，则打开连接
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            blnOpenDB = True
        End If
    End If
    
    '如果lngXWStudyID=0，则需要到新网数据库中查找医嘱ID对应的检查号
    If lngXWStudyID = 0 Then
        strSql = "select distinct F_STU_ID as Study主键 from V_OEM_SERIES where F_STU_NO ='" & lngOrderID & "'"
        Set rsTemp = gcnXWDBServer.Execute(strSql)
        If Not rsTemp.EOF Then
            lngXWStudyID = rsTemp!Study主键
            rsTemp.MoveNext
        End If
    End If

    While lngXWStudyID <> 0
    
        '调用新网存储过程“P_OEM_UNMATCHING_RIS”，取消关联
        strSql = "P_OEM_UNMATCHING_RIS(" & lngXWStudyID & ")"
        gcnXWDBServer.Execute strSql
        
        If rsTemp Is Nothing Then
            lngXWStudyID = 0
        Else
            If Not rsTemp.EOF Then
                lngXWStudyID = rsTemp!Study主键
                rsTemp.MoveNext
            Else
                lngXWStudyID = 0
            End If
        End If
    Wend
    
    '调用中联存储过程，取消关联
    strSql = "select F_SER_ID as SERIES主键,F_STU_ID as Study主键,F_SER_UID as 序列UID,F_SER_DATE as 序列日期,F_SER_TIME as 序列时间, " _
            & " F_SER_CONTEXT as 序列描述,F_MODALITY as 影像类型,F_STU_NO as 医嘱ID from V_OEM_SERIES where F_STU_NO ='" & lngOrderID _
            & "' order by F_STU_ID ,F_SER_ID"
    Set rsTemp = gcnXWDBServer.Execute(strSql)
    If rsTemp.EOF = True Then
        strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsUnmatchImage(" & lngOrderID & ")"
        zlDatabase.ExecuteProcedure strSql, "取消关联"
    End If
    
    '如果是在过程中打开的数据库连接，则退出时关闭连接
    If blnOpenDB = True Then
        Call XWDBServerClose
    End If
    
    XWUnmatchImage = 0
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    XWUnmatchImage = 2
End Function

'---------------------------------------------------------------------------------------
'接收报告的Windows消息
'---------------------------------------------------------------------------------------

Public Function XWHook(ByVal hWnd As Long) As Long
    '指定自定义的窗口过程
    '返回并保存原来默认的窗口过程指针
    If App.LogMode <> 0 Then
        XWHook = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf XWWindowProc)
        Debug.Print "Hooked"
    End If
End Function

Public Sub XWUnhook(ByVal hWnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
    If App.LogMode <> 0 Then
        temp = SetWindowLong(hWnd, GWL_WNDPROC, lpWndProc)
    End If
End Sub

Function XWWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'消息处理程序，专门处理特定的 WM_XWReportImage 消息
    Dim strLog As String
    
    If uMsg = WM_XWReportImage Then
        strLog = Now & " umsg = " & uMsg & ";wparam = " & wParam & ";lparam = " & lParam & vbCrLf
        
        If gblnXWLog = True Then
            Call WriteCommLog("XWWindowProc", "XW接口", strLog)
        End If
        '接收新网发送到系统剪贴板的报告图像
        If lParam <> 0 Then
            Call XWSaveReportImages(lParam)
        End If
    End If
  
    '调用原来的窗口过程
    XWWindowProc = CallWindowProc(plngXWPreWndProc, hw, uMsg, wParam, lParam)
End Function

Public Sub XWSaveReportImages(lngOrderID As Long)
'------------------------------------------------
'功能：将图像从剪贴板保存成报告图
'参数： lngOrderID -- 医嘱ID
'返回：
'------------------------------------------------
    Dim dcmImage As New DicomImage
    Dim strFileName As String
    Dim strLocalPath As String
    Dim dcmG As New DicomGlobal
    Dim strTempPath As String, lngBuffSize As Long
    Dim strStudyUID As String
    Dim strDeviceNO As String
    Dim strFtpIp As String
    Dim strFtpUrl As String
    Dim strFtpVirtualPath As String
    Dim strFTPUser As String
    Dim strFTPPwd As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim Inet As New clsFtp
    Dim lngResult As Long

    On Error GoTo err

    If gfrmPacsMain Is Nothing Then Exit Sub

    '从剪贴板获取报告图
    dcmImage.Paste
    '根据规则产生报告图名称
    dcmG.RegString("UIDRoot") = "1"
    strFileName = dcmG.NewUID & ".jpg"
    
    '获取存储设备并建立FTP中的各级目录
    strSql = "select 检查UID FROM 影像检查记录 where 医嘱ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询检查UID", lngOrderID)
    
    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW接口", "查询检查UID，查询的SQL是：" & strSql & vbCrLf & "查询到的记录数为： " & rsTemp.RecordCount)
    End If
    
    If rsTemp.EOF = True Then Exit Sub
    
    strStudyUID = rsTemp!检查UID
    Call GetDeptStorageDevice(gfrmPacsMain, lngOrderID, strStudyUID, glngXWDeptID, G_LNG_PACSSTATION_MODULE, gblnXWMoved, strDeviceNO, _
                strFtpIp, strFtpUrl, strFtpVirtualPath, strFTPUser, strFTPPwd)

    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW接口", "提取存储设备，检查UID = " & strStudyUID & "，医嘱ID = " & lngOrderID & "，执行科室ID= " & glngXWDeptID & "，存储设备号= " & strDeviceNO)
    End If
    
    '获取本地临时文件名
    strFtpVirtualPath = Replace(strFtpVirtualPath, strFtpUrl, "")
    strLocalPath = App.Path & "\TmpImage\" & Replace(strFtpVirtualPath, "/", "\")
    '创建本地目录
    Call MkLocalDir(strLocalPath)

    '将报告图保存成文件
    dcmImage.FileExport strLocalPath & "\" & strFileName, "JPG"
    
    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW接口", "保存本地图像，图像文件名为：" & strLocalPath & "\" & strFileName)
    End If

    '将报告图上传到FTP目录,并保存到数据库
    lngResult = Inet.FuncFtpConnect(strFtpIp, strFTPUser, strFTPPwd)
    If lngResult <> 0 Then
        lngResult = Inet.FuncUploadFile(strFtpUrl & strFtpVirtualPath, strLocalPath & "\" & strFileName, strFileName)
        Inet.FuncFtpDisConnect
        
        If gblnXWLog = True Then
            Call WriteCommLog("XWSaveReportImages", "XW接口", "上传FTP图像，FTP IP地址= " & strFtpIp & "，FTP子目录=" & strFtpUrl & strFtpVirtualPath & "，图像文件名为：" & strFileName)
        End If
        
         '修改数据库，增加报告图
        If lngResult = 0 Then
            strSql = "ZL_影像检查报告_ADD('" & strStudyUID & "','" & strFileName & "')"
            zlDatabase.ExecuteProcedure strSql, "保存报告图像"
            
            If gblnXWLog = True Then
                Call WriteCommLog("XWSaveReportImages", "XW接口", "数据库保存报告图，执行存储过程：" & strSql)
            End If
        
            strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsSetFTPDeviceNo(" & lngOrderID & ",'" & strDeviceNO & "')"
            zlDatabase.ExecuteProcedure strSql, "保存报告图像的FTP设备号"
            
            If gblnXWLog = True Then
                Call WriteCommLog("XWSaveReportImages", "XW接口", "保存报告图的设备号，执行存储过程：" & strSql)
            End If
        End If
    End If

    Exit Sub
err:
    '不处理，不提示
    Debug.Print Now & "出错:" & err.Description & vbCrLf
    Inet.FuncFtpDisConnect
    
    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW接口", "保存报告图出错，错误描述是：" & err.Description)
    End If
        
End Sub

Public Sub subXWShowArchiveManager(intType As Integer)
'------------------------------------------------
'功能：调用新网ArchiveManager实现额外的功能
'参数： intType = 1---删除图像；2---发送图像；3---光盘刻录
'返回：无，直接打开ArchiveManager
'------------------------------------------------
    Dim strCommand As String
    Dim strUser As String
    Dim strPswd As String
    
    On Error GoTo err
    
    If intType = 1 Then     '删除图像
        strUser = zlDatabase.GetPara("XW删除图像用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
        strPswd = zlDatabase.GetPara("XW删除图像密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    ElseIf intType = 2 Then     '发送图像
        strUser = zlDatabase.GetPara("XW发送图像用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
        strPswd = zlDatabase.GetPara("XW发送图像密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    Else    '光盘刻录
        strUser = zlDatabase.GetPara("XW光盘刻录用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
        strPswd = zlDatabase.GetPara("XW光盘刻录密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    End If
    
    
    If strUser <> "" And strPswd <> "" Then
        strCommand = "C:\PACS\ArchiveManager.exe " & strUser & "^" & strPswd
        
        If gblnXWLog = True Then
            Call WriteCommLog("subXWShowArchiveManager", "XW接口", "打开ArchiveManager，进行" & IIf(intType = 1, "删除图像", IIf(intType = 2, "发送图像", "光盘刻录")) & "的操作，命令是：" & strCommand)
        End If
            
        Shell strCommand, vbMaximizedFocus
    Else
        If gblnXWLog = True Then
            Call WriteCommLog("subXWShowArchiveManager", "XW接口", "进行" & IIf(intType = 1, "删除图像", IIf(intType = 2, "发送图像", "光盘刻录")) & "的操作时，用户名和密码不能为空。 " & vbCrLf _
                & "用户名是：" & strUser & "，密码是：" & strPswd)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub
