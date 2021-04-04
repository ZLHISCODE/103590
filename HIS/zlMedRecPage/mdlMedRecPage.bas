Attribute VB_Name = "mdlMedRecPage"
Option Explicit
'-----------------------------------------------------------
'标准版与病案系统共用项目
'-----------------------------------------------------------
'1------接口变量
Public gcnOracle As ADODB.Connection    '公共数据库连接，特别注意：不能设置为新的实例
Public gclsMipModule As zl9ComLib.clsMipModule
'2------全局变量
Public gclsMain As Object                         '当前的类实例
Public gclsPros As clsProperty                   '属性类
Public gstrSysName As String                       '系统名称
Public gstrProductName As String               'OEM产品名称
Public gstrUnitName As String                     '用户单位名称
Public gobjReport As clsReport                    '报表打印部件，用于首页打印预览
Public gcolPrivs As Collection                      '记录内部模块的权限
Public UserInfo As TYPE_USER_INFO
Public gobjPatient As Object                        '病人信息接口
'5------全局变量
Public grsDeptInfo As ADODB.Recordset         '临床科室缓存记录集
Public gintCA As Integer                           '电子签名认证中心
Public gstrESign As String                         '电子签名控制场合
Public grsSign As Recordset                      '电子签名启用部门
Public gobjRis As Object                            '新网RIS接口
Public gblnSet  As Boolean                         '中间变量防止事件重复调用
Public gColErr As New Collection
Public gColWarn As New Collection
Public gColCtl As New Collection                    '首页控件集合

Public gBlnNew As Boolean                     '是否开启首页外挂附页
Public gfrmMecCol As Collection                     '外挂附页对象
Public gPic外挂附页 As Integer             '外挂附页picturebox的index
Public colErrTmp As Collection                 '外挂部件提示信息集合
Public gIntPic As Integer                  '源程序的picturebox的总量
'-----------------------------------------------------------
'标准版(独有)
'-----------------------------------------------------------
'1------接口变量
Public gblnHaveOPS As Boolean               '是否安装手麻系统，系统号= 2400
Public gobjCommunity As Object              '社区档案接口对象
Public gobjPass As Object                          '合理用药接口对象
Public gobjESign As Object                        '签名部件对象
Public gclsInsure As New clsInsure           '医保变量
Public gobjPlugIn As Object                      '外挂功能对象
'-----------------------------------------------------------
''病案系统(独有)
'-----------------------------------------------------------
'5------全局变量
Public grsBabyInfo As ADODB.Recordset         '孕妇分娩的新生儿信息
Public grsBabyDiag As ADODB.Recordset        '孕妇分娩的新生儿的诊断信息
Public grsDeliceryInfo As ADODB.Recordset    '孕妇分娩的相关信息
'日志跟踪对象
Public gobjLog As TextStream
Public gobjFSO As New FileSystemObject
Public gblnUnload As Boolean '用于记录首页窗体是否已经卸载

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = Nvl(rsTmp!简码)
            UserInfo.姓名 = Nvl(rsTmp!姓名)
            UserInfo.DeptID = Nvl(rsTmp!部门ID, 0)
            UserInfo.DeptNo = rsTmp!部门码 & ""
            UserInfo.DeptName = rsTmp!部门名 & ""
            UserInfo.DBUser = rsTmp!用户名 & ""
            GetUserInfo = True
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckShare(ByVal lngSysShareNO As Long, Optional ByVal lngSysMainNO As Long = 100) As Boolean
'功能：标准系统和其他系统是否是共享安装
'参数：lngSysShareNO= 共享安装的系统
'           lngSysMainNO=主系统
    Dim lngShareNum As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
'Select * From (Select * From zlSystems Start With 编号 = 100 Connect By Prior 编号 = 共享号) Where 编号 = 300
'Select * From (Select * From zlSystems Start With 编号 = 300 Connect By Prior 编号 = 共享号) Where 编号 = 100
    strSQL = "Select s.编号" & vbNewLine & _
            "From zlSystems S" & vbNewLine & _
            "Where s.正常安装 = 1 And s.编号  = [1] And s.共享号 = [2]"
    On Error GoTo errH
    '由于存在多帐套情况若标准版多帐套编号100，101，。。。。，199，因此若此判断
    '多张套不能共享安装
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, (lngSysShareNO \ 100) * 100, (lngSysMainNO \ 100) * 100)
    CheckShare = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckHaveSys(ByVal lngSysNO As Long) As Boolean
'功能：判断是否安装了某个系统
'参数：lngSysNO 判定的系统
    Dim lngShareNum As Long
    Dim strSQL As String
    Dim rsTmp As Recordset

    strSQL = "Select s.编号" & vbNewLine & _
            "From zlSystems S" & vbNewLine & _
            "Where s.正常安装 = 1 And Floor(s.编号 / 100) = [1] "

    On Error GoTo errH
    '由于存在多帐套情况若标准版多帐套编号100，101，。。。。，199，因此若此判断
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngSysNO \ 100)

    If rsTmp.RecordCount > 0 Then CheckHaveSys = True: Exit Function
    CheckHaveSys = False
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function TranNumToDate(ByVal strNum As String, Optional ByVal blnDec As Boolean = False) As String
'功能：转换数值为日期
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String

    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)

    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"

    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay

    If Not IsDate(strDate) Then Exit Function

    strDate = Format(strDate, "yyyy-mm-dd")
    If blnDec Then strDate = DateAdd("d", -1, Format(strDate, "yyyy-mm-dd"))
    TranNumToDate = strDate
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function HaveRIS(Optional ByVal blnMsg As Boolean) As Boolean
'功能：判断 新网接口部件 是否存在
'参数：blnMsg－创建失败时是否提示
    If gobjRis Is Nothing Then
        On Error Resume Next
        Set gobjRis = CreateObject("zl9XWInterface.clsHISInner")
        Err.Clear: On Error GoTo 0
    End If
    If gobjRis Is Nothing Then
        If blnMsg Then
            MsgBox "新网接口部件(zl9XWInterface)未创建成功！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    HaveRIS = True
End Function
