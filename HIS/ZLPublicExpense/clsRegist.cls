VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'接口说明:
'1.通用接口
'1.1 zlInitCommon：初始化相关的系统号及相关连接(本工程使用时，可以不初始化)
'1.2 zlGetRegisterMaxLen_Tradition:获取挂号安排的最大长度
'1.3 zlGetDoctorFromDeptID:根据部门ID来获取所涉及的医生集
'1.4 zlPatiEdit:编辑病人信息
'2.挂号处理相关接口
'2.1 zlBulidingPriceDataFromRegistNo：根据挂号单构键收费划价单
'2.2 zlCheckRegisterNoIsCharge:检查指定挂号单是否已经收费
'2.3 zlRegisterPriceDeleteFromNO：根据挂号单据，删除对应的划价单（必须满足病人挂号记录．收费单不为空）
'3.根据指定号码或日期获取挂号相关数据
'3.0   zlGetRegisterMode:根据指定日期获取排班模式
'3.1   zlGetRegistTimes_Tradition：根据指定日期和号码，获取传统排班模式下的的时间段集
'3.2   zlGetRegistSNStartus_Tradition：根据号别及日期，获取传统排班模式下的序号状态数据集
'3.3   zlGetRegistUnitCtrol_Tradition：根据号别及指定挂号时间获取合约单位挂号控制数据集
'3.4   zlRegisterIsStartTimeSlice_Tradition：:获取指定号码及指定时间是否启用时间时段
'3.5   zlGetRegisterData_Tradition:获取指定日期的挂号数据（传统模式下)
'3.6   zlGetRegisterData_Visits:获取指定日期的挂号数据（临床出诊模式下)
'3.7   zlGetRegRoomsFromPlanID:根据安排ID来获取所涉及的诊室集
'3.8   zlRegisterCheckValied:检查挂号有效性检查(不区分模式)
'3.9   zlRegisterCheckValied_Visits:检查挂号有效性检查(临床出诊表模式)
'3.10  zlRegisterCheckValied_Tradition:检查挂号有效性检查(传统模式)
'3.11  zlGetRegisterPlanID_Tradition:根据号码和日期 获取有效的安排ID或计划ID
'3.12  zlGetRegRoomsFromDeptid:根据部门ID或医生相关信息,获取诊室信息
'3.13  zlGetRegisterMaxDaysFromDeptAndDoctor:获取最大的预约天数(不区分模式)
'3.13.01  zlGetRegisterMaxDaysFromDeptAndDoctor_Tradition:获取最大的预约天数(传统模式)
'3.13.02  zlGetRegisterMaxDaysFromDeptAndDoctor_Visits:获取最大的预约天数(临床出诊表模式)
'3.14   zlGetRegisterNextSn:取下一个序号(不分模式)
'3.14.01  zlGetRegisterNextSn__Tradition:取一下个序号(传统模式)
'3.14.02  zlGetRegisterNextSn__Visits:取一下个序号(出诊表模式)
'3.15  zlReserveRegNo-挂号序号锁号
'3.16  CancelRegNo-挂号解锁

'4.合作单位控制相关
'4.1  zlRegisterIsStartUnitCtrol_Tradition：判断挂号安排是否启用了合约单位的相关控制
'6. 获取上班时间段相关信息
'6.1. zlGetRegisterWorkTime_Record：获取上班时间段(即”时间段”表的所有数据）
'6.2. zlGetWeekNameFromDate：根据指定日期，获取星期几名称(比如：周一，周二...)
'6.3. zlGetRegisterWorkTime：根据时间段名称及站点等，获取缺省预约时间，开始时间和终止时间)
'6.4. zlGetRegisterTimeWorksFromDate:获取指定日期所涉及的时间段名称,: 比如：全日,上午等...
'6.5  zlCheckIsValiedWorkTimeFromWorkTimeName:根据时间段名称及日期来判断当前时间是否在上班时间范围内
'*********************************************************************************************************************************************
'内部方法及过程说明
'1.InitParaSysPara:初始化系统参数
'*********************************************************************************************************************************************

Private mfrmRegist As frmRegist
Attribute mfrmRegist.VB_VarHelpID = -1
Private mlngModul  As Long
Private mintUnitStart As Integer    '0-未初始化合并作单数据 ;-1未启用合作单位;1-启用了合作单位
Private mrsWorkTime As ADODB.Recordset  '时间段
Private mintStartTimeSlice As Integer   '0-未初始化是否启用时间段;-1未启用时间段 ;1-启用了时间段

Private mbytMaxLen  As Integer
Private Type ty_SysPara '系统参数类型
    bln启用排班模式 As Boolean  'true-表示启用了排班模式的
    bln已启用出诊表  As Boolean '是否已经启用了出诊表模式
    str启用日期     As String   '格式为:yyyy-mm-dd
    
    int允许预约天数 As Integer
    bln预约接收确定挂号费  As Boolean
    int预约失效次数    As Integer
     
    byt普通有效天数 As Byte    '普通挂号有效天数
    byt急诊有效天数 As Byte '急诊挂号有效天数
    bln退号重用  As Boolean
End Type
Private mSysPara As ty_SysPara
Private Sub InitParaSysPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化系统参数
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-03 14:19:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strParValue As String, varPara As Variant
    
    '排班模式判断
    strParValue = gobjDatabase.GetPara("挂号排班模式", glngSys, , "0")
    varPara = Split(strParValue & "|", "|")
    mSysPara.str启用日期 = CStr(varPara(1))
    mSysPara.bln启用排班模式 = Val(varPara(0)) = 1 And IsDate(mSysPara.str启用日期)   '是否已启用了排班模式
    mSysPara.bln已启用出诊表 = False
    If mSysPara.bln启用排班模式 Then
        If CDate(mSysPara.str启用日期) <= gobjDatabase.CurrentDate Then mSysPara.bln已启用出诊表 = True
    End If
    On Error GoTo errHandle
    With mSysPara
        .int允许预约天数 = Val(gobjDatabase.GetPara("挂号允许预约天数", glngSys))
        If .int允许预约天数 = 0 Then .int允许预约天数 = 15 '缺省为15天
        
        .bln预约接收确定挂号费 = gobjDatabase.GetPara("预约接收确定挂号费", glngSys, 1111, 0) = "1"   '暂以挂号窗口为准
        .int预约失效次数 = Val(gobjDatabase.GetPara("预约失约次数", glngSys, 1111, 0))         '暂以挂号窗口为准
        
        
        '两位:前一位普能挂号;后一位急诊挂号
        strParValue = gobjDatabase.GetPara(21, glngSys, , "01") & "1"
        .byt普通有效天数 = Val(Left(strParValue, 1))
        .byt急诊有效天数 = Val(Mid(strParValue, 2, 1))
        If .byt普通有效天数 = 0 Then .byt普通有效天数 = 1
        If .byt急诊有效天数 = 0 Then .byt急诊有效天数 = 1
        .bln退号重用 = Val(gobjDatabase.GetPara("已退序号允许挂号", glngSys, 1111)) = 1
    End With

    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    Call gclsInsure.InitOracle(gcnOracle)
    If zlGetComLib = False Then Exit Function
    
    Call GetUserInfo
    Call InitVar
    Call InitParaSysPara
    
    Set mfrmRegist = New frmRegist
    If mfrmRegist.zlInitVar(Me) = False Then Exit Function
    
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInitData(ByVal bytState As Byte, _
     Optional ByVal strDocs As String, Optional ByVal strDeptIDs As String, Optional ByVal strRooms As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的限制条件
    '入参:bytState-应用场合：目前支持 0-分诊台调用，1-门诊医生站调用，2-住院医生站调用, 3-患者服务中心
    '     strDocs-医生ID或者医生姓名,用逗号分隔，格式为 147,235,313…或者张三，李四，王五…，传入时显示传入医生的挂号安排
    '     strDeptIDs-科室ID,用逗号分隔，格式为 147,235,313…，传入时显示该科室的挂号安排
    '     strRooms-医生站调用时，可传入诊室，挂号时自动确定诊室。
    '返回:初始化成功,返回true,否则返回False
    '编制:刘尔旋
    '日期:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    gbytState = bytState
    Select Case bytState
        Case 0
            glngModul = 1113
        Case 1
            glngModul = 1260
        Case 2
            glngModul = 1261
        Case 3
            glngModul = 1115
    End Select
    mlngModul = glngModul
    gstrDocs = strDocs
    gstrDeptIDs = strDeptIDs
    gstrRooms = strRooms
    gstrPrivs = ";" & gobjComlib.GetPrivFunc(glngSys, 9000) & ";"
    zlInitData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlDefCommandBars(ByVal frmMain As Object, ByVal CommandBars As Object, _
                            Optional ByVal blnAddInTool As Boolean, Optional ByVal MenuControlBefore As Object, _
                            Optional ByVal ToolControlBefore As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建窗体菜单和工具条
    '入参:frmMain-调用的主窗体
    '     CommandBars-调用窗体的CommandBars对象
    '     blnAddInTool-是否将功能按钮中添加进工具栏
    '     MenuControlBefore-菜单栏中要添加功能对象的后一个Control对象，用以控制添加的功能对象的位置，不传时，根据默认规则指定按钮位置。
    '     ToolControlBefore-工具栏中要添加功能对象的后一个Control对象，用以控制添加的功能对象的位置，不传时，根据默认规则指定按钮位置。
    '编制:刘尔旋
    '日期:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glngModul = mlngModul
    Call mfrmRegist.zlDefCommandBars(frmMain, CommandBars, blnAddInTool, MenuControlBefore, ToolControlBefore)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体的菜单和工具条的各项功能的Visible属性或Eanbeld属性
    '入参:Control-菜单控件
    '编制:刘尔旋
    '日期:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glngModul = mlngModul
    Call mfrmRegist.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal frmMain As Object, ByVal Control As Object, ByRef strOutNO As String, Optional ByVal lngPatiID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行挂号的相关功能
    '入参:frmMain-调用的主窗体
    '     Control-菜单控件
    '出参:strOutNo-成功保存的挂号单据号
    '编制:刘兴洪
    '日期:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glngModul = mlngModul
    Call mfrmRegist.zlExecuteCommandBars(frmMain, Control, strOutNO, lngPatiID)
End Sub

Public Function zlOpenStopedPlanBySN(ByVal frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal lng记录ID As Long, _
    Optional ByVal lngDeptID As Long, Optional ByVal lngDoctorID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：对启用了序号控制分时段的已停诊安排按序号开放挂号
    '入参：
    '   frmMain 调用的主窗体
    '   lngModule 调用模块号
    '   lng记录ID 记录ID,1114模块调用时传入
    '   lngDeptID 科室ID
    '   lngDoctorID 医生ID
    '返回：成功返回True，失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmOpenStopedPlanBySN As New frmOpenStopedPlanBySN
    zlOpenStopedPlanBySN = frmOpenStopedPlanBySN.ShowMe(frmMain, lngModule, lng记录ID, lngDeptID, lngDoctorID)
End Function
Public Function zlGetWeekNameFromDate(ByVal dtDate As Date) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回指定日期的星期
    '入参:dtDate-指定日期
    '出参:
    '返回:返回星期几（如：周一，周二...周日)
    '编制:刘兴洪
    '日期:2017-12-18 15:14:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWeekName As String, strSQL As String

    On Error GoTo errHandle
    Select Case Weekday(dtDate)
    Case vbSunday    '1 星期日（缺省值）
        strWeekName = "周日"
    Case vbMonday    '2 星期一
        strWeekName = "周一"
    Case vbTuesday   '3 星期二
        strWeekName = "周二"
    Case vbWednesday '4 星期三
        strWeekName = "周三"
    Case vbThursday  '5 星期四
        strWeekName = "周四"
    Case vbFriday    '6 星期五
        strWeekName = "周五"
    Case vbSaturday  '7 星期六
        strWeekName = "周六"
    Case Else
        strWeekName = ""
    End Select
    zlGetWeekNameFromDate = strWeekName
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegisterMode(ByVal strDate As String, ByRef bytRegisterMode As Byte, Optional strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定日期获取排班模式
    '入参:strDate-当前日期(yyyy-mm-dd hh:mm:ss),为空时，以当前日期判断
    '出参:bytRegisterMode-获取成功返回排班模式(0-传统排班模式;1-临床出诊表排班模式)
    '    strErrMsg-出错时，返回错误信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-09-27 19:52:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mSysPara.bln启用排班模式 = False Then bytRegisterMode = 0: zlGetRegisterMode = True: Exit Function
    If mSysPara.bln已启用出诊表 Then bytRegisterMode = 1: zlGetRegisterMode = True: Exit Function
    
    If IsDate(mSysPara.str启用日期) = False Then bytRegisterMode = 0: zlGetRegisterMode = True: Exit Function
    
    '检查指定日期是否启用了出诊表模式
    If strDate = "" Then strDate = Format(gobjDatabase.CurrentDate, "YYYY-MM-DD HH:MM:SS")
    If CDate(mSysPara.str启用日期) <= CDate(strDate) Then
        bytRegisterMode = 1
    Else
        bytRegisterMode = 0
    End If
    zlGetRegisterMode = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlGetRegistTimes_Tradition(ByVal str号码 As Long, ByVal dt挂号时间 As Date, ByRef rs时间段_Out As ADODB.Recordset, _
    Optional bln仅读安排 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：根据指定日期和号码，获取传统排班模式下的的时间段集
    '入参:str号码-号码
    '     dt挂号时间-指定挂号或预约的日期
    '     bln仅读安排-仅读安排
    '出参:rs时间段-返回时间段数据集
    '返回:获取成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-15 11:46:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWeekName As String, strSQL As String
    On Error GoTo errHandle
    strWeekName = zlGetWeekNameFromDate(dt挂号时间)
    Set rs时间段_Out = Nothing
    
    If bln仅读安排 Then
        strSQL = "" & _
        " Select A.安排id, A.序号,to_char(a.开始时间,'hh24')||':00' as 时间点, " & _
        "           to_char(A.开始时间,'hh24:mi') as 开始时间, to_char(A.结束时间,'hh24:mi') as 结束时间, " & _
        "           to_Char([2],'yyyy-mm-dd')||' '||to_char(A.开始时间,'hh24:mi') as 实际开始时间, " & _
        "           to_Char([2],'yyyy-mm-dd')||' '||to_char(A.结束时间,'hh24:mi') as 实际结束时间, " & _
        "           A.限制数量,A.是否预约 " & vbNewLine & _
        " From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
        " Where A.安排id = B.ID  And B.号码=[1] And a.星期=[3] " & vbNewLine & _
        " Order By A.开始时间 "
        Set rs时间段_Out = gobjDatabase.OpenSQLRecord(strSQL, "根据号码及时间，获取有效的时间段(zlPublicExpense.clsRegist)", str号码, dt挂号时间, strWeekName)
        zlGetRegistTimes_Tradition = True
        Exit Function
    End If
    

    '普通号分时段的情况 仅对预约号限制
    strSQL = "" & _
    " Select Distinct A.序号, To_Char(A.开始时间, 'hh24') || ':00' As 时间点, " & vbNewLine & _
    "        To_Char(A.开始时间, 'hh24:mi') As 开始时间,To_Char(A.结束时间, 'hh24:mi') As 结束时间, " & vbNewLine & _
    "           to_Char([2],'yyyy-mm-dd')||' '||to_char(A.开始时间,'hh24:mi') as 实际开始时间," & vbNewLine & _
    "           to_Char([2],'yyyy-mm-dd')||' '||to_char(A.结束时间,'hh24:mi') as 实际结束时间, " & vbNewLine & _
    "        A.限制数量, A.是否预约 " & vbNewLine & _
    " From 挂号安排时段 A, 挂号安排 B " & vbNewLine & _
    " Where A.安排id = B.ID And B.号码 =[1] And A.星期=[3] " & vbNewLine & _
    "       And Not Exists ( Select 1 From 挂号安排计划 E Where E.安排id = B.ID And E.审核时间 Is Not Null" & vbNewLine & _
    "                        And [2] Between Nvl(E.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And E.失效时间" & vbNewLine & _
    "                       )" & vbNewLine
    
    strSQL = strSQL & _
    " Union All " & vbNewLine & _
    " Select Distinct A.序号, To_Char(A.开始时间, 'hh24') || ':00' As 时间点," & vbNewLine & _
    "        To_Char(A.开始时间, 'hh24:mi') As 开始时间, To_Char(A.结束时间, 'hh24:mi') As 结束时间," & vbNewLine & _
    "           to_Char([2],'yyyy-mm-dd')||' '||to_char(A.开始时间,'hh24:mi') as 实际开始时间," & vbNewLine & _
    "           to_Char([2],'yyyy-mm-dd')||' '||to_char(A.结束时间,'hh24:mi') as 实际结束时间, " & vbNewLine & _
    "        A.限制数量, A.是否预约" & vbNewLine & _
    " From 挂号计划时段 A,挂号安排计划 B," & vbNewLine & _
    "       (   Select Max(A.号码) as 号码,Max(A.生效时间) 生效  " & vbNewLine & _
    "           From 挂号安排计划 A " & _
    "           Where A.号码 = [1] And A.审核时间 Is Not Null " & vbNewLine & _
    "              And  [2] Between Nvl(A.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And  A.失效时间 " & vbNewLine & _
    "        ) D" & vbNewLine & _
    " Where A.计划id = B.ID And B.号码 = D.号码 And B.生效时间 = D.生效  And A.星期=[3] And B.审核时间 Is Not Null " & vbNewLine & _
    " Order By 开始时间　"
    Set rs时间段_Out = gobjDatabase.OpenSQLRecord(strSQL, "根据号码及时间，获取有效的时间段(zlPublicExpense.clsRegist)", str号码, dt挂号时间, strWeekName)
    zlGetRegistTimes_Tradition = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlGetRegistSNStartus_Tradition(ByVal str号别 As String, ByVal dtRegistDate As Date, _
    Optional ByRef rsSN As ADODB.Recordset, Optional lngSN As Long, Optional bln专家 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据号别及日期，获取传统排班模式下的序号状态数据集
    '入参:str号别-号别
    '     dtRegistDate-挂号日期
    '     bln专家-True:专家号,False-表示普通分时段
    '出参:rsSN-序号集(字段:序号,状态,操作员姓名,预约,日期)
    '返回:获取成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-15 17:05:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, datStart  As Date, datEnd  As Date
    Dim strWhere As String
    
    On Error GoTo errHandle
    
    Set rsSN = Nothing
    
    If Not bln专家 Then  '普通分时段
        strSQL = "" & _
        "   Select Count(1) As 预约数量,To_Char(日期,'HH24:MI') AS 日期,日期 as 实际日期 " & _
        "   From 挂号序号状态" & _
        "   Where 号码=[1] And trunc(日期)=[2]" & vbNewLine & _
        " Group By 日期 "
        Set rsSN = gobjDatabase.OpenSQLRecord(strSQL, "根据号别及日期，获取传统排班模式下的序号状态数据集(zlPublicExpense.clsRegist)", str号别, CDate(Format(dtRegistDate, "yyyy-mm-dd")))
        zlGetRegistSNStartus_Tradition = True
        Exit Function
    End If
    
    If dtRegistDate = CDate(0) Then
        strWhere = " And 日期 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 "
    Else
        strWhere = " And 日期 Between  [2] And [3] "
    End If
    If lngSN > 0 Then strWhere = strWhere & " And 序号=[4] "
    
    datStart = CDate(Format(dtRegistDate, "yyyy-MM-dd"))
    datEnd = DateAdd("s", -1, DateAdd("d", 1, datStart))
    
    strSQL = "" & _
    " Select 序号,状态,操作员姓名,Nvl(预约,0) as 预约,TO_Char(日期,'hh24:mi:ss') as 日期  " & vbNewLine & _
    " From 挂号序号状态 " & vbNewLine & _
    " Where 号码=[1]" & strWhere
    Set rsSN = gobjDatabase.OpenSQLRecord(strSQL, "根据号别及日期，获取传统排班模式下的序号状态数据集(zlPublicExpense.clsRegist)", str号别, datStart, datEnd, lngSN)
    zlGetRegistSNStartus_Tradition = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetRegistUnitCtrol_Tradition(ByVal str号码 As String, ByRef dtRegDate As Date, ByRef rsUnitReg As ADODB.Recordset, Optional bln专家 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据号别及指定挂号时间获取合约单位挂号控制数据集
    '入参:str号码-号别
    '     dtRegDate-挂号日期
    '     bln专家-true:专家号号源;False-普通号
    '出参:dtRegDate-挂号数据集
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-18 11:42:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWeekName As String, strSQL As String
    
    On Error GoTo errHandle
    strWeekName = zlGetWeekNameFromDate(dtRegDate)
    
    Set rsUnitReg = Nothing
    If bln专家 Then
        strSQL = "" & _
        " Select a.合作单位, a.限制项目, a.序号, a.数量" & vbCrLf & _
        " From 合作单位安排控制 a, 挂号安排 b" & vbCrLf & _
        " Where a.安排id = b.Id And b.号码 =[1]   And a.限制项目(+)=[3] And  a.数量 <> 0 " & vbCrLf & _
        "     And Not Exists (  Select 1 rom 挂号安排计划 e where e.安排id = b.Id And e.审核时间 Is Not Null " & vbCrLf & _
        "                       And [2] Between Nvl(e.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And e.失效时间)" & vbCrLf & _
        " Union All" & vbCrLf & _
        " Select a.合作单位, a.限制项目, a.序号, a.数量" & vbCrLf & _
        " From 合作单位计划控制 a,  挂号安排计划 b," & vbCrLf & _
        "          (   Select Max(A.安排ID) 安排ID,Max(a.生效时间) 生效" & vbCrLf & _
        "              From 挂号安排计划 a, 挂号安排 b" & vbCrLf & _
        "              Where a.安排id = b.Id And b.号码 =[1] And a.审核时间 Is Not Null " & vbCrLf & _
        "                    And [2] Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And a.失效时间) d" & vbCrLf & _
        " Where a.计划id = b.Id And b.安排id = D.安排id And b.生效时间 = d.生效 And  a.数量 <> 0  And b.审核时间 Is Not Null" & vbCrLf & _
        "       And a.限制项目(+)=[3] "
        Set rsUnitReg = gobjDatabase.OpenSQLRecord(strSQL, "根据号码及时间，获取有效的时间段(zlPublicExpense.clsRegist)", str号码, dtRegDate, strWeekName)
        zlGetRegistUnitCtrol_Tradition = True
        Exit Function
    End If
    
  
    strSQL = "" & _
    " Select a.序号, Sum(Nvl(a.数量, 0)) As 数量" & vbCrLf & _
    " From 合作单位安排控制 a, 挂号安排 b" & vbCrLf & _
    " Where a.安排id = b.Id And b.号码 =[1] And a.限制项目=[3] And a.数量 <> 0  " & vbCrLf & _
    "       And Not Exists ( Select 1 From 挂号安排计划 e Where e.安排id = b.Id And e.审核时间 Is Not Null  " & vbCrLf & _
    "                        And [2] Between Nvl(e.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And e.失效时间)" & vbCrLf & _
    " Group By a.序号" & vbCrLf & _
    " Union All" & vbCrLf & _
    " Select a.序号, Sum(Nvl(a.数量, 0)) As 数量" & vbCrLf & _
    " From 合作单位计划控制 a, 挂号安排计划 b," & vbCrLf & _
    "          (Select Max(A.安排ID) as 安排ID, Max(a.生效时间) 生效" & vbCrLf & _
    "              From 挂号安排计划 a, 挂号安排 b" & vbCrLf & _
    "              Where a.安排id = b.Id And b.号码 = [1] And a.审核时间 Is Not Null " & vbCrLf & _
    "                    And [2] Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And a.失效时间) d" & vbCrLf & _
    " Where a.计划id = b.Id And b.安排id = D.安排Id  And b.生效时间 = d.生效 And a.数量 <> 0 " & _
    "             And a.限制项目(+)=[3] " & vbCrLf & _
    " Group By a.序号" & vbCrLf & _
    " Order By 序号"
    Set rsUnitReg = gobjDatabase.OpenSQLRecord(strSQL, "根据号码及时间，获取有效的时间段(zlPublicExpense.clsRegist)", str号码, dtRegDate, strWeekName)
    zlGetRegistUnitCtrol_Tradition = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlRegisterIsStartUnitCtrol_Tradition() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断挂号安排是否启用了合约单位的相关控制
    '返回:启用了合约单位返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-18 14:09:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp  As ADODB.Recordset
        
    On Error GoTo errHandle
    
    If mintUnitStart <> 0 Then zlRegisterIsStartUnitCtrol_Tradition = mintUnitStart = 1: Exit Function
    strSQL = "" & _
    " select 1 as 数据  From 合作单位安排控制 where Rownum<2  Union ALL " & _
    " Select 1 as 数据 from 合作单位计划控制 Where Rownum<2"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "判断挂号安排中是否启用了合作单位控制(zlPublicExpense.clsRegist)")
    mintUnitStart = IIf(rsTmp.EOF, -1, 1)
    zlRegisterIsStartUnitCtrol_Tradition = mintUnitStart = 1
    Set rsTmp = Nothing
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlGetRegisterWorkTime_Record(ByRef rsWorkTime As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号安排的工作时间段(时间段表中取数)
    '入参:
    '出参:rsWorkTime-返回所有的上班的工作时间段
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-18 14:55:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    If Not mrsWorkTime Is Nothing Then
        If mrsWorkTime.State = 1 Then Set rsWorkTime = mrsWorkTime: zlGetRegisterWorkTime_Record = True: Exit Function
    End If
    strSQL = "Select 时间段,开始时间,终止时间,缺省时间, 站点,号类,出诊预留时间,休息时段  From 时间段"
    Set mrsWorkTime = gobjDatabase.OpenSQLRecord(strSQL, "获取上班时间段(zlPublicExpense.clsRegist)")
    Set rsWorkTime = mrsWorkTime: zlGetRegisterWorkTime_Record = True: Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlCheckIsValiedWorkTimeFromWorkTimeName(ByVal strDate As String, ByVal strWorkTimeName As String, ByVal str站点 As String, _
    ByVal str号类 As String, Optional blnValiedWorkTime_Out As Boolean, Optional dtDefaultTime_Out As Date, Optional dtStartDate_out As Date, Optional dtEndDate_Out As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据工作时间段名称，检查该时间段名称是否在有效时间范围内
    '入参:strDate-指定日期(yyyy-mm-dd HH:MM:SS)
    '出参:blnValiedWorkTime_Out-False:在有效的时间范围内，True:不在有效的时间范围内
    '     dtDefaultTime_Out-缺省预约时间
    '     dtStartDate_out-开始时间
    '     dtEndDate_Out-开始时间
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-31 15:24:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsWorkTime As ADODB.Recordset
    Dim dtDate As Date, dtStartDate As Date, dtEndDate As Date, dtDefault As Date
    
    Dim strTime(0 To 1) As String, strStartTime(0 To 1) As String, strEndTime(0 To 1) As String
    
    On Error GoTo errHandle
    
    blnValiedWorkTime_Out = True
    If zlGetRegisterWorkTime_Record(rsWorkTime) = False Then Exit Function
    rsWorkTime.Filter = "时间段='" & strWorkTimeName & "'"
    '无效的时间段，返回False
    If rsWorkTime.EOF Then rsWorkTime.Filter = 0: Exit Function
    If strDate <> "" Then
        dtDate = CDate(strDate)
    Else
        dtDate = gobjDatabase.CurrentDate
    End If
    
    With rsWorkTime
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dtStartDate = Format(dtDate, "yyyy-mm-dd") & " " & Format(!开始时间, "HH:MM:SS")
            dtEndDate = Format(dtDate, "yyyy-mm-dd") & " " & Format(!终止时间, "HH:MM:SS")
            If Format(!缺省时间, "HH:MM:SS") <> "" Then
                dtDefault = Format(dtDate, "yyyy-mm-dd") & " " & Format(!缺省时间, "HH:MM:SS")
            Else
                dtDefault = dtDate
            End If
                    
            If dtStartDate >= dtEndDate Then dtEndDate = dtEndDate + 1   '加一天
            
            If Nvl(!站点) = str站点 And Nvl(!号类) = str号类 Then '老版暂不支持分站点及号类来设置设计
                blnValiedWorkTime_Out = False: rsWorkTime.Filter = 0:
                If dtDate >= dtStartDate And dtDate <= dtEndDate Then blnValiedWorkTime_Out = False
                dtDefaultTime_Out = dtDefault
                dtStartDate_out = dtStartDate
                dtEndDate_Out = dtEndDate
                zlCheckIsValiedWorkTimeFromWorkTimeName = True
                 Exit Function
            End If
            If Nvl(!站点) = "" And Nvl(!号类) = "" Then
                strTime(0) = Format(dtDefault, "yyyy-mm-dd HH:MM:SS")
                strStartTime(0) = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS")
                strEndTime(0) = Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
            End If
            
            If Nvl(!站点) = str站点 And Nvl(!号类) = "" Then
                strTime(1) = Format(dtDefault, "yyyy-mm-dd HH:MM:SS")
                strStartTime(1) = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS")
                strEndTime(1) = Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
            End If
            .MoveNext
        Loop
    End With
 
    rsWorkTime.Filter = 0
    If strStartTime(1) <> "" Then
        dtDefault = CDate(strTime(1))
        dtStartDate = CDate(strStartTime(1))
        dtEndDate = CDate(strEndTime(1))
        If dtDate >= dtStartDate And dtDate <= dtEndDate Then blnValiedWorkTime_Out = False
        dtDefaultTime_Out = dtDefault
        dtStartDate_out = dtStartDate
        dtEndDate_Out = dtEndDate
        zlCheckIsValiedWorkTimeFromWorkTimeName = True
    End If
    
    If strStartTime(0) <> "" Then
        dtDefault = CDate(strTime(0))
        dtStartDate = CDate(strStartTime(0))
        dtEndDate = CDate(strEndTime(0))
        
        If dtDate >= dtStartDate And dtDate <= dtEndDate Then blnValiedWorkTime_Out = False
        dtDefaultTime_Out = dtDefault
        dtStartDate_out = dtStartDate
        dtEndDate_Out = dtEndDate
        zlCheckIsValiedWorkTimeFromWorkTimeName = True
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function zlGetRegisterWorkTime(ByVal strWorkTimeName As String, ByVal str站点 As String, ByVal str号类 As String, _
    Optional strDefaultTime_Out As String, Optional strStartTime_Out As String, Optional strEndTime_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据时间段名称及站点等，获取缺省预约时间，开始时间和终止时间)
    '入参:strWorkTimeName-工作时间段(比如：上午，下午，白天等)
    '     str站点-工作站点名
    '     str号类-号类
    '出参:strDefaultTime_Out-缺省的预约时间(HH:MM:SS")
    '     strStartTime_Out-开始时间(HH:MM:SS")
    '     strEndTime_Out-终止时间(HH:MM:SS")
    '返回:获取成功，返回true,否则返回false
    '编制:刘兴洪
    '日期:2017-12-18 15:55:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsWorkTime As ADODB.Recordset
    Dim strTime(0 To 1) As String, strStartTime(0 To 1) As String, strEndTime(0 To 1) As String
    Dim dtDate As Date
    
    
    On Error GoTo errHandle
    If zlGetRegisterWorkTime_Record(rsWorkTime) = False Then Exit Function
    
    rsWorkTime.Filter = "时间段='" & strWorkTimeName & "'"
    If rsWorkTime.EOF Then
        '缺省为当前时间
        dtDate = gobjDatabase.CurrentDate
        strDefaultTime_Out = Format(dtDate, "HH:MM:SS"):
        strStartTime_Out = "00:00:00"
        strEndTime_Out = "23:59:59"
        rsWorkTime.Filter = 0
        zlGetRegisterWorkTime = True: Exit Function
    End If
    
    With rsWorkTime
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!站点) = str站点 And Nvl(!号类) = str号类 Then '老版暂不支持分站点及号类来设置设计
                strDefaultTime_Out = Format(!缺省时间, "HH:MM:SS")
                If Format(!终止时间, "HH:MM:SS") < Format(!开始时间, "HH:MM:SS") Then
                    strEndTime_Out = "23:59:59"
                Else
                    strEndTime_Out = Format(!终止时间, "HH:MM:SS")
                End If
                strStartTime_Out = Format(!开始时间, "HH:MM:SS")
                zlGetRegisterWorkTime = True
                rsWorkTime.Filter = 0: Exit Function
            End If
            If Nvl(!站点) = "" And Nvl(!号类) = "" Then
                strTime(0) = Format(!缺省时间, "HH:MM:SS")
                strStartTime(0) = Format(!开始时间, "HH:MM:SS")
                strEndTime(0) = Format(!终止时间, "HH:MM:SS")
            End If
            If Nvl(!站点) = str站点 And Nvl(!号类) = "" Then
                strTime(1) = Format(!缺省时间, "HH:MM:SS")
                strStartTime(1) = Format(!开始时间, "HH:MM:SS")
                strEndTime(1) = Format(!终止时间, "HH:MM:SS")
            End If
            .MoveNext
        Loop
    End With
    rsWorkTime.Filter = 0
    If strTime(1) <> "" Then
        
        strDefaultTime_Out = strTime(1)
        If strEndTime(1) < strStartTime(1) Then
            strEndTime_Out = "23:59:59"
        Else
            strEndTime_Out = strEndTime(1)
        End If
        strStartTime_Out = strStartTime(1)
        zlGetRegisterWorkTime = True: Exit Function
    End If
    If strTime(0) <> "" Then
        strDefaultTime_Out = strTime(0)
        If strEndTime(0) < strStartTime(0) Then
            strEndTime_Out = "23:59:59"
        Else
            strEndTime_Out = strEndTime(0)
        End If
        strStartTime_Out = strStartTime(0)
        zlGetRegisterWorkTime = True: Exit Function
    End If
    dtDate = gobjDatabase.CurrentDate
    strDefaultTime_Out = Format(dtDate, "HH:MM:SS"):
    strStartTime_Out = "00:00:00"
    strEndTime_Out = "23:59:59"
    zlGetRegisterWorkTime = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetIsStartTimeSlice_Tradition() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前是否启用了时间段
    '返回:启用了时间段返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-19 16:39:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    
    On Error GoTo errHandle
        
    If mintStartTimeSlice <> 0 Then GetIsStartTimeSlice_Tradition = IIf(mintStartTimeSlice = -1, False, True): Exit Function
    
    strSQL = _
    "    Select 1  dt From  挂号安排时段 Where Rownum<2" & vbNewLine & _
    "    Union All " & vbNewLine & _
    "    Select 1  dt From 挂号计划时段  Where Rownum<2 "
 
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查是否启用了时间段(zlPublicExpense.clsRegist.GetIsStartTimeSlice)")
    mintStartTimeSlice = IIf(rsTemp.EOF, -1, 1)
    GetIsStartTimeSlice_Tradition = IIf(mintStartTimeSlice = -1, False, True)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlRegisterIsStartTimeSlice_Tradition(ByVal str号码 As String, ByVal dtRegDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定号码及指定时间是否启用时间时段
    '入参:str号码-号码
    '     dtRegDate-指定的挂号日期
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-19 16:33:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strWeekName As String, strSQL As String
    
    
    On Error GoTo errHandle
    If GetIsStartTimeSlice_Tradition = False Then zlRegisterIsStartTimeSlice_Tradition = False: Exit Function
    strWeekName = zlGetWeekNameFromDate(dtRegDate)
    strSQL = _
    " Select   Havedata, 安排id" & vbNewLine & _
    " From ( Select 1 As Havedata, b.Id As 安排id " & vbNewLine & _
    "              From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
    "              Where B.号码=[1] And A.安排id = b.ID And a.星期=[3]  And Rownum<2 " & vbNewLine & _
    "                    And Not Exists (   Select 1 From 挂号安排计划 C Where c.安排id = b.Id And c.审核时间 Is Not Null " & _
    "                                       And [2] Between Nvl(c.生效时间, [2]) And c.失效时间) " & vbNewLine & _
    "        Union All " & vbNewLine & _
    "        Select 1 As Havedata, c.Id As 安排id" & vbNewLine & _
    "        From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C," & vbNewLine & _
    "             (  SELECT Max(A.安排ID) as 安排ID,MAX(a.生效时间 ) 生效 " & vbNewLine & _
    "                FROM 挂号安排计划 a,挂号安排 B   " & vbNewLine & _
    "                WHERE a.安排Id=b.ID AND b.号码=[1] AND a.审核时间 IS NOT NULL" & vbNewLine & _
    "                      And [2] Between nvl(a.生效时间,to_date('1900-01-01','yyyy-mm-dd')) And a.失效时间" & vbNewLine & _
    "              ) D  " & vbNewLine & _
    "         Where  a.计划id=b.ID and b.安排ID=d.安排ID And b.生效时间=d.生效 " & vbNewLine & _
    "                      And   a.星期=[3]  And Rownum<2  " & vbNewLine & _
    "   ) "
    
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "根据号码及时间，获取有效的时间段(zlPublicExpense.clsRegist)", str号码, dtRegDate, strWeekName)
    zlRegisterIsStartTimeSlice_Tradition = rsTemp.EOF = False
    Set rsTemp = Nothing
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetRegisterMaxLen_Tradition() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号安排的号码的最大长度
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-19 18:47:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If mbytMaxLen > 0 Then zlGetRegisterMaxLen_Tradition = mbytMaxLen: Exit Function
    
    strSQL = "Select Nvl(Max(Length(号码)),5) as 长度 From 挂号安排"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "根据号码及时间，获取有效的时间段(zlPublicExpense.clsRegist)")
    mbytMaxLen = Val(Nvl(rsTemp!长度))
    zlGetRegisterMaxLen_Tradition = mbytMaxLen
    Set rsTemp = Nothing
    Exit Function
errH:
    zlGetRegisterMaxLen_Tradition = 5
End Function
Public Function zlGetRegisterTimeWorksFromDate(Optional ByVal strDate As String = "", Optional str站点 As String, Optional str号类 As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定日期所涉及的时间段名称
    '入参:strDate-指定日期,为空时，表示以当前时间为准,格式:\yyyy-mm-dd hh24:mi:ss
    '     str站点-站点编号
    '     str号类=""
    '出参:
    '返回:返回时间段名称，多个用逗号分隔，比如：全日,上午等...
    '编制:刘兴洪
    '日期:2018-01-03 16:13:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTimeWorks As String, rsWorkTime As ADODB.Recordset
    Dim strTimeWorks1 As String '区分站点且号类的,以站点且号类的为准
    Dim strTimeWorks2 As String '区分站点或号类的,以站点或号类的为准
    Dim dtStart As Date, dtEnd As Date
    Dim dtDate As Date
    
    On Error GoTo errHandle
    
    If strDate = "" Then
        dtDate = gobjDatabase.CurrentDate
    Else
        dtDate = CDate(strDate)
    End If
    strTimeWorks = "": strTimeWorks1 = ""
    If zlGetRegisterWorkTime_Record(rsWorkTime) = False Then Exit Function
    With rsWorkTime
        .Filter = 0: If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' 时间段,开始时间,终止时间,nvl(nvl(缺省时间,终止时间),sysdate) as 缺省时间, 站点,号类,出诊预留时间,休息时段
            dtStart = CDate(Format(dtDate, "yyyy-mm-dd") & " " & Format(!开始时间, "HH:MM:SS"))
            dtEnd = CDate(Format(dtDate, "yyyy-mm-dd") & " " & Format(!终止时间, "HH:MM:SS"))
            
            If dtStart >= dtEnd Then dtEnd = dtEnd + 1 - 1 / 24 / 60 / 60
            If dtDate >= dtStart And dtDate <= dtEnd Then
                If Nvl(!站点) <> str站点 And str站点 <> "" And Nvl(!号类) <> str号类 And str号类 <> "" Then
                 strTimeWorks1 = strTimeWorks1 & "," & Nvl(!时间段)
                ElseIf Nvl(!站点) <> str站点 And str站点 <> "" Or Nvl(!号类) <> str号类 And str号类 <> "" Then
                 strTimeWorks2 = strTimeWorks2 & "," & Nvl(!时间段)
                ElseIf Nvl(!站点) = "" And Nvl(!号类) = "" Then
                 strTimeWorks = strTimeWorks & "," & Nvl(!时间段)
                End If
            End If
            .MoveNext
        Loop
    End With
    If strTimeWorks <> "" Then strTimeWorks = Mid(strTimeWorks, 2)
    zlGetRegisterTimeWorksFromDate = strTimeWorks
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    
End Function
 
 
Public Function zlGetRegisterData(ByRef rsRegisterData_Out As ADODB.Recordset, Optional ByVal str科室IDs As String = "", _
    Optional ByVal strRegDate As String = "", Optional bln预约 As Boolean = False, Optional bytRegMode_Out As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取安排数据
    '入参:str科室IDs-只取所涉及的科室IDs(传入空表示,所有科室)
    '     strRegDate-预约日期(为空时，表示取当前有效挂号数据）,格式:yyyy-mm-dd hh24:mi:ss
    '出参:rsRegisterData_Out-返回挂号安排的相关数据集
    '     bytRegMode_Out:当前挂号模式:1-临床出诊表模式;0-传统模式
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-03 11:51:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    bytRegMode_Out = 0
    If strRegDate = "" Then    '主要是避免子过程重复调用gobjDatabase.Currentdate,
        strRegDate = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd HH:MM:SS")
    End If
    
    Set rsRegisterData_Out = Nothing
    If zlGetRegisterMode(strRegDate, bytRegMode_Out) = False Then Exit Function
    If bytRegMode_Out = 1 Then
       zlGetRegisterData = zlGetRegisterData_Visits(rsRegisterData_Out, str科室IDs, strRegDate, bln预约)
    Else
       zlGetRegisterData = zlGetRegisterData_Tradition(rsRegisterData_Out, str科室IDs, strRegDate, bln预约)
    End If
    
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Public Function zlGetRegisterData_Tradition(ByRef rsRegisterData_Out As ADODB.Recordset, Optional ByVal str科室IDs As String = "", Optional ByVal strRegDate As String = "", Optional bln预约 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号安排数据
    '入参:str科室IDs-只取所涉及的科室IDs(传入空表示,所有科室)
    '     strRegDate-预约日期(为空时，表示取当前有效挂号数据）,格式:yyyy-mm-dd hh24:mi:ss
    '     bln预约-当前是否预约
    '出参:rsRegisterData_Out-返回挂号安排的相关数据集
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-03 11:51:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strSQL As String
    Dim str星期 As String, dtRegDate As Date, dtNextDayDate As Date
    Dim strWorkTimes As String
    
    
    If strRegDate = "" Then
        dtRegDate = gobjDatabase.CurrentDate
        dtNextDayDate = dtRegDate
    Else
        dtRegDate = CDate(strRegDate): dtNextDayDate = CDate(Format(dtRegDate, "yyyy-mm-dd")) + 1 - 1 / 24 / 60 / 60
    End If
    
    str星期 = "a." & zlGetWeekNameFromDate(dtRegDate)
    
    On Error GoTo errHandle
    
    strWorkTimes = zlGetRegisterTimeWorksFromDate(dtRegDate)
    
    
    strWhere = " And " & str星期 & " Is not NULL And " & "B.限制项目(+)='" & Replace(str星期, "a.", "") & "' And  instr([1],','||" & str星期 & "||',')>0 "
    
    If strRegDate <> "" Then
        strWhere = strWhere & " And [2] Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And  Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))"
    Else
        strWhere = strWhere & " And Sysdate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And  Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))"
    End If
    
    If str科室IDs <> "" Then strWhere = strWhere & " And A.科室ID In (Select  /*+cardinality(J,10) */ Column_Value From Table(f_num2list([4]))) "
    If strRegDate <> "" And bln预约 Then strWhere = strWhere & " And [3] <=Trunc(Sysdate)+Nvl(A.预约天数," & mSysPara.int允许预约天数 & ")" '预约限制
    
    strWhere = strWhere & " And (Nvl(A.医生ID,0)=0 Or Exists(Select 1 From 人员表 Q Where A.医生ID=Q.ID And nvl(Q.撤档时间,Sysdate+1)>Sysdate)) "
    
    
    strSQL = "" & _
    "  Select A.ID, 0 as 计划ID,A.号码, A.号类, A.科室id,J.简码 as 科室简码,J.名称 as 科室名称,j.编码 as 科室编码 , A.项目id,L.编码 as 项目编码,L.名称 as 项目名称, A.病案必须, " & _
    "         A.医生id, M.编号 as 医生编号,M.简码 as 医生简码,A.医生姓名,  " & _
    "         A.周日, A.周一, A.周二, A.周三,  A.周四 , A.周五, A.周六, A.分诊方式,a.开始时间,a.终止时间, A.序号控制, B.限号数, B.限约数,a.停用日期 " & vbNewLine & _
    "  From 挂号安排 A, 挂号安排限制 B,部门表 J,人员表 M,收费项目目录 L" & vbNewLine & _
    " Where a.停用日期 Is Null And a.ID = B.安排id(+) And Nvl(a.是否删除, 0) = 0 " & _
    "       And A.项目ID=L.ID  And Nvl(L.撤档时间,sysdate+1)>sysdate " & _
    "       and A.科室ID=J.ID And   Nvl(J.撤档时间,sysdate+1)>sysdate  And (J.站点='" & gstrNodeNo & "' Or J.站点 is Null)  " & _
    "       and A.医生ID=M.ID(+)  " & _
    "       " & strWhere
    
    If strRegDate <> "" And bln预约 Then
        strSQL = strSQL & _
        "  And   Not Exists(Select 1 From 挂号安排计划 where 安排ID=A.id And ([2] BETWEEN 生效时间 + 0 and 失效时间)  And 审核时间 is not NULL  )"
 
       strWhere = Replace(strWhere, "a.开始时间", "a.生效时间")
       strWhere = Replace(strWhere, "Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))", "a.失效时间")
       strWhere = Replace(strWhere, "a.预约天数", "c.预约天数")
       strWhere = Replace(strWhere, "A.科室ID", "c.科室ID")
       
        strSQL = strSQL & " Union  ALL " & _
        "  Select C.ID,A.ID as 计划ID, A.号码, C.号类, C.科室id,J.简码 as 科室简码,J.名称 as 科室名称,j.编码 as 科室编码 , A.项目id,L.编码 as 项目编码,L.名称 as 项目名称,  C.病案必须, " & _
        "         A.医生id,M.编号 as 医生编号,M.简码 as 医生简码,A.医生姓名," & _
        "         A. 周日, A.周一, A.周二, A.周三,A.周四 , A.周五, A.周六, A.分诊方式,a.生效时间,a.失效时间, A.序号控制, B.限号数, B.限约数,C.停用日期 " & vbNewLine & _
        "  From 挂号安排计划 A, 挂号计划限制 B,挂号安排 C,部门表 J,人员表 M,收费项目目录 L " & vbNewLine & _
        "  Where c.停用日期 Is Null  And Nvl(c.是否删除, 0) = 0 And a.审核时间 Is Not Null  " & _
        "       And a.生效时间 = (  Select Max(生效时间) From 挂号安排计划  " & _
        "                       Where 安排id = a.安排id And [2] Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And 失效时间 And 审核时间 Is Not Null)" & _
        "       And a.ID = B.计划id(+) And a.安排id = c.Id " & _
        "       And A.项目ID=L.ID and Nvl(L.撤档时间,sysdate+1)>sysdate " & _
        "       and C.科室=J.ID And and Nvl(J.撤档时间,sysdate+1)>sysdate  And (J.站点='" & gstrNodeNo & "' Or J.站点 is Null)  " & _
        "       and A.医生ID=M.ID(+)   " & strWhere
    End If
    strSQL = strSQL & " Order by 科室编码"
    
    
    Set rsRegisterData_Out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist", strWorkTimes, dtRegDate, dtNextDayDate, str科室IDs)
    zlGetRegisterData_Tradition = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
Public Function zlGetRegisterData_Visits(ByRef rsRegisterData_Out As ADODB.Recordset, Optional ByVal str科室IDs As String = "", Optional ByVal strRegDate As String = "", Optional bln预约 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取临床出诊安排数据
    '入参:str科室IDs-只取所涉及的科室IDs(传入空表示,所有科室)
    '     strRegDate-预约日期(为空时，表示取当前有效挂号数据）,格式:yyyy-mm-dd hh24:mi:ss
    '出参:rsRegisterData_Out-返回挂号安排的相关数据集
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-03 11:51:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strSQL As String
    Dim dtRegDateTime  As Date, dtDayDate As Date, dtPreDayDate As Date
    
    If strRegDate = "" Then
        dtRegDateTime = gobjDatabase.CurrentDate
    Else
        dtRegDateTime = CDate(strRegDate)
    End If
    
    '无时间:当前挂号日期及昨天至今的挂号日期
    dtDayDate = CDate(Format(dtRegDateTime, "yyyy-mm-dd")): dtPreDayDate = CDate(Format(dtDayDate - 1, "yyyy-mm-dd"))
    
    On Error GoTo errHandle
    If bln预约 Then
        '预约处理:'0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约;3-按比例控制预约;4-按总量控制预约;5-按序号控制预约
        strWhere = strWhere & " And  Nvl(a.预约控制,0) <> 1 "
        strWhere = strWhere & " And Sysdate + zl_Fun_GetAppointmentDays + Decode(Nvl(B.预约天数,0),0," & mSysPara.int允许预约天数 & ",B.预约天数) > [3]"
    End If
    
     If str科室IDs <> "" Then strWhere = strWhere & " And Instr([5],','||B.科室ID||',')>0 "
     strWhere = strWhere & " And (Nvl(A.医生ID,0)=0 Or Exists(Select 1 From 人员表 Q Where A.医生ID=Q.ID And nvl(Q.撤档时间,Sysdate+1)>Sysdate)) "
     
    strSQL = "" & _
    " Select a.Id, b.号码, b.号类, b.科室id, c.简码 as 科室简码,c.名称 As 科室名称, c.编码 As 科室编码, a.项目id, d.编码 as 项目编码,d.名称 As 项目名称, Nvl(b.是否建病案, 0) As 病案, " & _
    "       m.简码 as 替诊医生简码,a.替诊医生id, a.替诊医生姓名, a.替诊开始时间, a.替诊终止时间,a.医生id, E.编号 as 医生编号,E.简码 as 医生简码,a.医生姓名, Nvl(a.已挂数, 0) As 已挂," & vbNewLine & _
    "       Nvl(a.已约数, 0) As 已约, a.限号数 As 限号, a.限约数 As 限约, Nvl(d.项目特性, 0) As 急诊, Decode(a.分诊方式,1,'指定',2,'动态',3,'平均',NULL) As 分诊," & vbNewLine & _
    "       a.是否序号控制 As 序号控制, a.是否分时段, a.是否独占, a.上班时段 As 排班, a.号源id, a.缺省预约时间, a.开始时间, a.终止时间, a.出诊日期 " & vbNewLine & _
    " From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D, 人员表 E,人员表 M" & vbNewLine & _
    " Where (a.出诊日期 = [1] Or a.出诊日期 = [2]) And [3] Between Nvl(a.提前挂号时间,a.开始时间) And  a.终止时间 And Nvl(a.是否锁定, 0) = 0  And Nvl(a.是否发布,0) = 1  " & _
    "         And ( a.开始时间 < Nvl(a.停诊开始时间,a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间,a.开始时间)  " & _
    "               or   Exists (Select 1 From 临床出诊序号控制 C Where C.记录ID=A.ID And Nvl(C.是否停诊,0) = 0 And A.是否序号控制 =1 And A.是否分时段 = 1 ) " & _
    "              ) And [3] Not Between  Nvl(a.停诊开始时间,a.终止时间) And Nvl(a.停诊终止时间,a.开始时间)" & _
    "       And a.号源id = B.Id  and nvl(b.是否删除,0)=0 and  nvl(B.撤档时间,sysdate+1)>sysdate  " & _
    "       And b.科室id = c.Id And Nvl(C.撤档时间,sysdate+1) > Sysdate  And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null) " & _
    "       And a.项目id = d.Id and Nvl(d.撤档时间,sysdate+1) > Sysdate " & vbNewLine & _
    "       And a.医生id = e.Id(+)  And A.替诊医生ID=M.ID(+)   " & strWhere & _
     " Order By 科室编码 "
                
    Set rsRegisterData_Out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist)", dtDayDate, dtPreDayDate, dtRegDateTime, gdatRegistTime, "," & str科室IDs & ",")
    zlGetRegisterData_Visits = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
Public Function zlGetDoctorFromDeptID(ByVal lngDeptID As Long, ByRef rsDoctor_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据部门ID来获取所涉及的医生集
    '入参:lngDeptID-科室ID
    '出参:rsDoctor_Out-返回医集信息集
    '返回:获取成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 11:46:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String

    On Error GoTo errHandle
    strSQL = _
    " Select  distinct C.ID,c.编号,c.姓名,C.简码,d.编码 as 科室编码,d.名称 as 科室名称,d.简码 as 科室简码,d.ID as 科室ID" & _
    " From 人员性质说明 a, 部门人员 b ,人员表 c,部门表 D" & vbCrLf & _
    " Where b.人员id=c.id And b.人员id=a.人员id And a.人员性质='医生' " & vbCrLf & _
    "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) And nvl(c.撤档时间,sysdate+1)>sysdate" & vbNewLine & _
    "       And  b.部门ID=d.ID And b.部门id=[1]  "
    
    Set rsDoctor_Out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(获取指定部门所涉及医生)", lngDeptID)
    zlGetDoctorFromDeptID = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
Public Function zlGetRegisterPlanID_Tradition(ByVal str号码 As String, _
    ByRef lng安排ID_out As Long, ByRef lng计划ID_Out As Long, Optional ByVal strDate As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据号别获取传统模式下的按排ID
    '入参:str号码-号码
    '     strDate-日期(格式为:yyyy-mm-dd hh:mm:ss),为空以当前时间判断
    '出参:lng安排ID_out-返回安排ID(安排有效时,无效返回0)
    '     lng计划ID_Out-返回计划ID(计划有效时,无效返回0)
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-16 17:24:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtDate As Date
    On Error GoTo errHandle
    
    lng安排ID_out = 0: lng计划ID_Out = 0
    
    strSQL = "" & _
    "   Select ID as 安排ID,0 as 计划ID From 挂号安排 a where 号码=[1] And a.停用日期 Is Null  And Nvl(a.是否删除, 0) = 0  "
    If strDate <> "" Then
        dtDate = CDate(strDate)
        strSQL = strSQL & " UNION ALL " & _
        " Select 0 as 安排ID,B.ID as 计划ID  " & _
        " From 挂号安排 A,挂号安排计划 B " & _
        " where A.号码=[1] And  A.id=B.安排ID And a.停用日期 Is Null And  Nvl(a.是否删除, 0) = 0  And B.审核时间 Is Not Null" & _
        "       And B.生效时间 = (  Select Max(生效时间) From 挂号安排计划  " & _
        "                       Where 安排id = b.安排id And [2] Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And 失效时间 And 审核时间 Is Not Null)" & _
        ""
        strSQL = "Select Max(安排ID) as 安排ID,Max(计划ID) as 计划ID From (" & strSQL & ")"
    End If
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist", str号码, dtDate)
    If rsTemp.EOF Then Exit Function
    lng安排ID_out = Val(Nvl(rsTemp!安排ID)): lng计划ID_Out = Val(Nvl(rsTemp!计划ID))
    If lng计划ID_Out <> 0 Then lng安排ID_out = 0
    
    zlGetRegisterPlanID_Tradition = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function
Public Function zlGetRegRoomsFromDeptid(ByVal lng科室ID As Long, ByVal lng项目id As Long, ByVal str医生姓名 As String, ByRef rsRooms_out As ADODB.Recordset, Optional strRegData As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据部门ID或医生相关信息,获取诊室信息
    '入参:strRegData-指定日期
    '     lng项目ID-项目ID
    '出参:rsRooms_out-返回诊室集
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-16 18:08:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim bytMode As Byte, dtDate As Date
    
    On Error GoTo errHandle
    
    If zlGetRegisterMode(strRegData, bytMode) = False Then Exit Function

    
    If bytMode = 0 Then
        strSQL = "" & _
        "   Select C.ID,C.ID, c.编码, c.名称,c.简码, c.位置  " & _
        "   From 挂号安排 A,挂号安排诊室 B,门诊诊室 C   " & _
        "   Where A.ID=b.号表ID and b.门诊诊室=c.名称 and A.科室ID=[1] " & IIf(lng项目id <> 0, " And A.项目ID=[3]", "") & _
        "       And (A.医生姓名=[2] Or A.医生姓名 Is Null )  " & _
        "   Order by 名称"
        Set rsRooms_out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist", lng科室ID, str医生姓名, lng项目id)
        zlGetRegRoomsFromDeptid = True
        Exit Function
    End If
    
    If strRegData = "" Then
        dtDate = Format(CDate(strRegData), "yyyy-mm-dd")
    Else
        dtDate = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd")
    End If
    
    strSQL = "" & _
    "   Select C.ID,C.ID, c.编码, c.名称,c.简码, c.位置  " & _
    "   From 临床出诊记录 A,临床出诊诊室记录 B,门诊诊室 C  " & _
    "   Where  A.ID=b.记录ID and A.诊室ID=C.ID  and a.出诊日期=[3]   " & _
    "       And  A.科室ID=[1] And (A.医生姓名=[2] Or A.医生姓名 Is Null)" & IIf(lng项目id <> 0, " And A.项目ID=[4]", "") & _
    "   Order by C.名称"
    
    Set rsRooms_out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist", lng科室ID, str医生姓名, dtDate, lng项目id)
    If rsRooms_out.RecordCount = 0 Then
        strSQL = "" & _
        "   Select C.ID,C.ID, c.编码, c.名称,c.简码, c.位置  " & _
        "   From 门诊诊室适用科室 A,门诊诊室 C  " & _
        "   Where  A.诊室iD=c.id and  A.科室ID=[1]" & _
        "   Order by C.名称"
        Set rsRooms_out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist", lng科室ID)
    End If
    
    zlGetRegRoomsFromDeptid = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegRoomsFromPlanID(ByVal lng安排ID As Long, ByRef rsRooms_out As ADODB.Recordset, Optional ByVal blnBusy As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据安排ID来获取所涉及的诊室集
    '入参:lng安排ID-安排ID(如果是新版本，则是记录ID,否则是安排ID)
    '     blnBusy-诊室忙时，允许分诊
    '出参:rsDoctor_Out-返回诊室信息集
    '返回:获取成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 13:39:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytMode As Byte, strWhere As String, rsTemp As ADODB.Recordset
    Dim byt分诊方式 As Byte, lngDeptID As Long, strSQL As String
    
    On Error GoTo errHandle
    
    Call zlGetRegisterMode("", bytMode)
    
    If bytMode = 0 Then '传统模式
        strWhere = IIf(blnBusy, " ", " And b.缺省标志=0 ")
        strSQL = _
        " Select b.ID, b.编码, b.名称,b.简码, b.位置" & vbNewLine & _
        " From 挂号安排诊室 a, 门诊诊室 b" & vbNewLine & _
        " Where a.门诊诊室 = b.名称 And a.号表id=[1]" & strWhere & _
        " Order By B.编码"
        Set rsRooms_out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(根据安排ID来获取所涉及的诊室集)", lng安排ID)
        zlGetRegRoomsFromPlanID = True
        Exit Function
    End If
    
    '0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
    strWhere = IIf(blnBusy, " ", " And b.缺省标志=0 ")
    
    byt分诊方式 = 0: lngDeptID = 0
    strSQL = "Select 分诊方式,科室ID From 临床出诊记录 A where ID=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(根据安排ID来获取所涉及的诊室集)", lng安排ID)
    If Not rsTemp.EOF Then
        byt分诊方式 = Val(Nvl(rsTemp!分诊方式))
        lngDeptID = Val(Nvl(rsTemp!科室ID))
    End If
    
    If byt分诊方式 = 0 Then
    
        strSQL = _
         " Select b.ID, b.编码, b.名称,b.简码, b.位置" & vbNewLine & _
         " From 门诊诊室适用科室 a, 门诊诊室 b" & vbNewLine & _
         " Where a.诊室id = b.id And a.科室id=[1]" & strWhere & _
         " Order By B.编码"
        Set rsRooms_out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(根据安排ID来获取所涉及的诊室集)", lngDeptID)
        zlGetRegRoomsFromPlanID = True
        Exit Function
    End If
    
    
    strSQL = _
    " Select b.ID, b.编码, b.名称,b.简码, b.位置" & vbNewLine & _
    " From 临床出诊诊室记录 a, 门诊诊室 b" & vbNewLine & _
    " Where   a.诊室id = b.id And a.记录id = [1] " & strWhere & _
    " Order By B.编码 "
    
    Set rsRooms_out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(根据安排ID来获取所涉及的诊室集)", lng安排ID)
    zlGetRegRoomsFromPlanID = True

    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
Public Function zlRegisterCheckValied_Tradition(ByVal lng病人ID As Long, ByVal lng安排ID As Long, ByVal lng计划ID As Long, ByRef strErrMsg_Out As String, Optional ByVal intRegNum As Integer = 1, _
    Optional strRegDate As String = "", Optional ByVal bln预约 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查挂号有效性检查(传统模式)
    '入参:lng安排ID-安排ID
    '     lng计划ID-计划ID(如果不为零，以计划为准)
    '     inRegNum-本次挂号数量
    '     strRegDate-挂号日期(yyyy-mm-dd HH:MM:SS)
    '     bln预约-是否预约
    '出参:strErrMsg_Out-返回错误信息
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 15:09:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsRegPlan As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strWeekName As String, dtDate As Date, dbl限号数 As Double, bytMode As Byte
    Dim dtTruncDate As Date   'yyyy-mm-dd
    Dim lng科室ID As Long, lng项目id As Long, str医生姓名 As String, lng医生ID As Long, str号码 As String
    
    
    On Error GoTo errHandle
    If strRegDate = "" Then
        dtDate = gobjDatabase.CurrentDate
    Else
        dtDate = CDate(strRegDate)
    End If
    
    dtTruncDate = Format(dtDate, "yyyy-mm-dd")
    strWeekName = zlGetWeekNameFromDate(dtDate)
    
    If lng计划ID <> 0 Then
    
        strSQL = "" & _
        " Select decode( nvl(c.预约天数,0),0," & mSysPara.int允许预约天数 & ",c.预约天数) as 预约天数,A.限约数,A.限号数,C.号码,C.科室ID,B.项目ID,B.医生ID,B.医生姓名  " & _
        " From 挂号安排计划 B,挂号计划限制 A,挂号安排 c " & _
        " Where B.安排ID=C.ID and  B.ID=a.计划ID(+) And  B.ID=[1] And A.限制项目(+)=[2] "
        
        Set rsRegPlan = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号有效性检查(传统模式))", lng计划ID, strWeekName)
    
    Else
        strSQL = "" & _
        " Select decode( nvl(B.预约天数,0),0," & mSysPara.int允许预约天数 & ",B.预约天数) as 预约天数,A.限约数,A.限号数,B.号码,B.科室ID,B.项目ID,B.医生ID,B.医生姓名 " & _
        " From 挂号安排限制 A,挂号安排 B " & _
        " Where  B.ID=[1] And  B.ID=A.安排ID(+) And A.限制项目(+)=[2] "
        Set rsRegPlan = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号有效性检查(传统模式))", lng安排ID, strWeekName)
    End If
    If rsRegPlan.EOF Then
        Set rsRegPlan = Nothing
        strErrMsg_Out = strWeekName & "不存安排"
        Exit Function
    End If
    lng科室ID = Val(Nvl(rsRegPlan!科室ID)): lng项目id = Val(Nvl(rsRegPlan!项目ID))
    lng医生ID = Val(Nvl(rsRegPlan!医生ID)): str号码 = Nvl(rsRegPlan!号码)
    str医生姓名 = Nvl(rsRegPlan!医生姓名)
    If str医生姓名 = "" Then str医生姓名 = "-"
  
    If bln预约 Then
        dbl限号数 = Val(Nvl(rsRegPlan!限约数))
        If dbl限号数 = 0 Then dbl限号数 = Val(Nvl(rsRegPlan!限号数))
        If dbl限号数 <> 0 Then
           '检查当前日期已经预约出去的号
           strSQL = "Select sum(已约数) as 已约数 From 病人挂号汇总 Where 日期=[1] and 科室ID=[2] and 项目ID=[3] and nvl(医生姓名,'-')=[4] and nvl(医生ID,0)=[5] And 号码=[6]"
           Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号有效性检查(传统模式))", dtTruncDate, lng科室ID, lng项目id, str医生姓名, lng医生ID, str号码)
           If dbl限号数 - Val(Nvl(rsTemp!已约数)) - intRegNum < 0 Then
                 strErrMsg_Out = Format(dtTruncDate, "yyyy-mm-dd") & "已无余号"
                 Exit Function
           End If
        End If
        If dtTruncDate - gobjDatabase.CurrentDate > Nvl(rsRegPlan!预约天数, 0) Then
            strErrMsg_Out = Format(dtTruncDate, "yyyy-mm-dd") & "超过了预约天数!"
            Exit Function
        End If
    Else
        dbl限号数 = Val(Nvl(rsRegPlan!限号数))
        If dbl限号数 <> 0 Then
           '检查当前日期已经预约出去的号
           strSQL = "Select sum(已挂数) as 已挂数 From 病人挂号汇总 Where 日期=[1] and 科室ID=[2] and 项目ID=[3] and nvl(医生姓名,'-')=[4] and nvl(医生ID,0)=[5] And 号码=[6]"
           Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号有效性检查(传统模式))", dtTruncDate, lng科室ID, lng项目id, str医生姓名, lng医生ID, str号码)
           If dbl限号数 - Val(Nvl(rsTemp!已挂数)) - intRegNum < 0 Then
                 strErrMsg_Out = Format(dtTruncDate, "yyyy-mm-dd") & "已无余号"
                 Exit Function
           End If
        End If
    End If
    
    'Zl_Fun_病人挂号记录_Check
    '--功能：挂号有效性检查(包含预约;预约挂号不扣款;预约挂号扣款)
    '--入参:操作方式_IN:0-挂号(包含收款预约),1-预约,2-预约接收
    '--     是否加号_In:是否加号调用，0-非加号调用，1-加号调用
    '--返回:0-检查通过
    '--     1-特定检查项检查失败，同时返回错误提示文本
    '--     2-其他错误导致的检查失败，同时返回错误提示文本
    '--说明:
    '--    当出诊记录id_In=0是,为不当班医生挂号,不需要进行挂号有效性检查,直接退出
    strSQL = "Select Zl_Fun_病人挂号记录_Check([1],[2],[3],NULL,[4],[5]) As 检查结果 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号有效性检查(传统模式))", bytMode, lng病人ID, str号码, dtDate, IIf(str医生姓名 <> "-", 1, 0))
    If rsTemp.EOF Then strErrMsg_Out = "有效性检查失败,无法继续！": Exit Function
    strSQL = Nvl(rsTemp!检查结果)
    If Val(Mid(strSQL, 1, 1)) <> 0 Then strErrMsg_Out = Mid(strSQL, 3): Exit Function
    zlRegisterCheckValied_Tradition = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlRegisterCheckValied_Visits(ByVal lng病人ID As Long, ByVal lng记录ID As Long, ByRef strErrMsg_Out As String, Optional ByVal intRegNum As Integer = 1, _
    Optional strRegDate As String = "", Optional ByVal bln预约 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查挂号有效性检查(临床出诊表模式)
    '入参:lng记录ID-记录ID
    '     inRegNum-本次挂号数量
    '     strRegDate-挂号日期(yyyy-mm-dd HH:MM:SS)
    '     bln预约-是否预约
    '出参:strErrMsg_Out-返回错误信息
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 15:09:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsRegPlan As ADODB.Recordset, rsTemp As ADODB.Recordset, bytMode As Byte
    Dim strWeekName As String, dtDate As Date, dbl限号数 As Double, str号码 As String, str医生姓名 As String
    Dim dtTruncDate As Date   'yyyy-mm-dd
     
     On Error GoTo errHandle
    If strRegDate = "" Then
        dtDate = gobjDatabase.CurrentDate
    Else
        dtDate = CDate(strRegDate)
    End If
    
    dtTruncDate = Format(dtDate, "yyyy-mm-dd")
 
    strSQL = "" & _
    "   Select decode( nvl(B.预约天数,0),0," & mSysPara.int允许预约天数 & ",B.预约天数)  as 预约天数 ,A.限约数,A.限号数,A.已挂数,A.已约数,B.号码,A.医生姓名 " & _
    "   From 临床出诊记录 A,临床出诊号源 B " & _
    "   Where A.号源ID=B.ID and A.ID=[1]"
    Set rsRegPlan = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号有效性检查(临床出诊表))", lng记录ID)
    
    If rsRegPlan.EOF Then
        Set rsRegPlan = Nothing
        strErrMsg_Out = Format(dtTruncDate, "yyyy-mm-dd") & "不存安排"
        Exit Function
    End If
     str号码 = Nvl(rsRegPlan!号码)
     str医生姓名 = Nvl(rsRegPlan!医生姓名)
    If bln预约 Then
        dbl限号数 = Val(Nvl(rsRegPlan!限约数))
        If dbl限号数 = 0 Then dbl限号数 = Val(Nvl(rsRegPlan!限号数))
        If dbl限号数 - Val(Nvl(rsRegPlan!已约数)) - intRegNum < 0 Then
            strErrMsg_Out = Format(dtTruncDate, "yyyy-mm-dd") & "已无余号"
            Exit Function
        End If
        If dtTruncDate - gobjDatabase.CurrentDate > Nvl(rsRegPlan!预约天数, 0) And dbl限号数 <> 0 Then
            strErrMsg_Out = Format(dtTruncDate, "yyyy-mm-dd") & "超过了预约天数!"
            Exit Function
        End If
    Else
        dbl限号数 = Val(Nvl(rsRegPlan!限号数))
        If dbl限号数 - Val(Nvl(rsRegPlan!已挂数)) - intRegNum < 0 And dbl限号数 <> 0 Then
            strErrMsg_Out = Format(dtTruncDate, "yyyy-mm-dd") & "已无余号"
            Exit Function
        End If
    End If
    bytMode = 0
    If bln预约 Then bytMode = 2
    
    strSQL = "Select Zl_Fun_病人挂号记录_Check([1],[2],[3],[4],[5],[6]) As 检查结果 From Dual"
    Set rsRegPlan = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号有效性检查(临床出诊表))", bytMode, lng病人ID, str号码, lng记录ID, dtDate, IIf(str医生姓名 <> "", 1, 0))
    If rsRegPlan.EOF Then strErrMsg_Out = "有效性检查失败,无法继续！": Exit Function
    
    strSQL = Nvl(rsRegPlan!检查结果)
    If Val(Mid(strSQL, 1, 1)) <> 0 Then strErrMsg_Out = Mid(strSQL, 3): Exit Function
    zlRegisterCheckValied_Visits = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
 
Public Function zlRegisterCheckValied(ByVal lng病人ID As Long, ByVal lng安排ID As Long, ByRef strErrMsg_Out As String, Optional ByVal intRegNum As Integer = 1, Optional ByVal lng计划ID As Long = 0, _
    Optional strRegDate As String = "", Optional ByVal bln预约 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查挂号有效性检查
    '入参:inRegNum-本次挂号数量
    '     strRegDate-挂号日期(yyyy-mm-dd HH:MM:SS)
    '     bln预约-是否预约
    '     lng计划ID-传统模式下，如果传入非零，则以计划来检查
    '出参:strErrMsg_Out-返回错误信息
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 15:09:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytMode As Byte
    Call zlGetRegisterMode(strRegDate, bytMode)
     
    On Error GoTo errHandle
    If bytMode = 0 Then
        zlRegisterCheckValied = zlRegisterCheckValied_Tradition(lng病人ID, lng安排ID, lng计划ID, strErrMsg_Out, intRegNum, strRegDate, bln预约)
    Else
        zlRegisterCheckValied = zlRegisterCheckValied_Visits(lng病人ID, lng安排ID, strErrMsg_Out, intRegNum, strRegDate, bln预约)
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlCheckRegisterNoIsCharge(ByVal strNO As String, ByRef intChargeed_Out As Integer, Optional blnCheckPriceNoIsChargeed As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定挂号单是否已经收费
    '入参:strNO-挂号单号
    '     blnCheckPriceNoIsChargeed-检查划价单是否收费：如果不检查划价单，生成划价单的挂号单，统一返回为未收费
    '出参:intChargeed_Out:-1-未找到对应的挂号单据,0-未收费;1-挂号单已收;2-免挂号模式下，还未产生划价记录; 3-挂号单对应的收费划价单已全收费(存在多张划价单时，必须全收的);4-挂号单对应的划价单存在部分收费)
    '返回:检查成功返回true,检查失败返回Fale
    '编制:刘兴洪
    '日期:2018-01-12 10:53:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str收费单 As String, blnAll As Boolean, blnNotCharge As Boolean
    On Error GoTo errHandle
    strSQL = " " & _
    "   Select a.No, Max(a.收费单) As 收费单, Max(a.取号标志) As 取号标志, Max(a.预约) As 是否预约, Max(b.结帐id) As 结帐id, Max(a.记录状态) As 记录状态, " & _
    "          Max(a.记录性质) As 记录性质 " & _
    "   From 病人挂号记录 A, 门诊费用记录 B " & _
    "   Where a.No = [1] And a.No = b.No And b.记录性质 = 4 And b.记录状态 In (0, 1, 3) " & _
    "   Group By a.No " & _
    "  "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号单是否收费）", strNO)
    
    intChargeed_Out = 0
    
    If rsTemp.EOF Then intChargeed_Out = -1: zlCheckRegisterNoIsCharge = True: Exit Function '未找到挂号单
    
    str收费单 = Nvl(rsTemp!收费单)
 
    If str收费单 = "" Then '未生成划价单
        '1.免挂号模式下，还未产生划价记录
        If Val(Nvl(rsTemp!取号标志)) = 1 Then
            intChargeed_Out = 2: zlCheckRegisterNoIsCharge = True: Exit Function
        End If
        
        If Val(Nvl(rsTemp!结帐ID)) = 0 Or Nvl(rsTemp!记录性质) = 2 Then   '未结帐的或预约单据，肯定是未收费的单据
            intChargeed_Out = 0: zlCheckRegisterNoIsCharge = True: Exit Function   '未收费
        Else
              intChargeed_Out = 1: zlCheckRegisterNoIsCharge = True: Exit Function    '已收费
        End If
    End If
    
    If Not blnCheckPriceNoIsChargeed Then zlCheckRegisterNoIsCharge = True: Exit Function    '有划价单且不检查划价单是否收费的，统一为未收费明:
    
    '/*+cardinality(b,10)*/
    
    '是否判断划价单是否收费
    strSQL = " " & _
    "   Select /*+cardinality(b,10)*/  a.No, Max(a.结帐id) as 结帐id " & _
    "   From 门诊费用记录 A, Table(f_Str2list([1])) B " & _
    "   Where a.记录性质 = 1 And a.记录状态 In (0, 1, 3) and a.No = b.column_value " & _
    "   group by a.No"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号单是否收费）", str收费单)
    If rsTemp.EOF Then intChargeed_Out = 2: zlCheckRegisterNoIsCharge = True: Exit Function '未找到划价单
    With rsTemp
        blnAll = True: blnNotCharge = True
        Do While Not .EOF
            If Val(Nvl(!结帐ID)) = 0 Then
                blnAll = False
            Else
                blnNotCharge = False
            End If
            If blnAll = False And blnNotCharge = False Then Exit Do
            .MoveNext
        Loop
        If blnNotCharge Then
            intChargeed_Out = 0
        Else
            intChargeed_Out = IIf(blnAll, 3, 4)
        End If
    End With
    zlCheckRegisterNoIsCharge = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
Public Function zlBulidingPriceDataFromRegistNo(ByVal strNO As String, ByVal lng科室ID As Long, str医生姓名 As String, _
    Optional str划价单_Out As String, Optional blnNotExcuteSaveData As Boolean = False, _
    Optional ByRef cllProSQL As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据挂号单构键收费划价单
    '入参:strNO-挂号单号
    '     lng科室ID-传入科室ID
    '     str医生姓名-传入医生姓名(传入医生和科室ID时）将会将划价单的执行人，执行科室及执行状态及执行时间填上)
    '     blnNotExcuteSaveData-true:不执行数据保存，直接返回执行的SQL集,False-执行数据保存
    '出参:cllProSQL:blnNotExcuteSaveData=true时，则返回需要执行的SQL集，否则返回Nothing
    '     str划价单_Out-返回当前生成的划价单
    '返回:执行成功或获取保存的SQL数据成功,返回true,检查失败返回Fale
    '编制:刘兴洪
    '日期:2018-01-12 10:53:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str划价单_NO As String
    
    On Error GoTo errHandle
    
    Set cllProSQL = Nothing
    str划价单_NO = gobjDatabase.GetNextNo(13)
    If str划价单_NO = "" Then
        MsgBox "未能正常获取划价单据号，请与系统管理员联系", vbInformation + vbOKOnly
        Exit Function
    End If
    
    
    'Zl_门诊划价记录_Buliding
    strSQL = "Zl_门诊划价记录_Buliding("
    '  单据号_In     病人挂号记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  划价单号_In   门诊费用记录.No%Type := Null,
    strSQL = strSQL & "'" & str划价单_NO & "',"
    '  科室id_In     门诊费用记录.执行部门id%Type := Null,
    strSQL = strSQL & IIf(lng科室ID = 0, "NULL", lng科室ID) & ","
    '  医生姓名_In   门诊费用记录.执行人%Type := Null,
    strSQL = strSQL & IIf(str医生姓名 = "", "NULL", "'" & str医生姓名 & "'") & ","
    '  操作员姓名_In 门诊费用记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  站点_In Varchar2:=Null
    strSQL = strSQL & "'" & gstrNodeNo & "')"
    If blnNotExcuteSaveData Then
        Set cllProSQL = New Collection
        zlAddArray cllProSQL, strSQL
        zlBulidingPriceDataFromRegistNo = True
        Exit Function
    End If
    
    gobjDatabase.ExecuteProcedure strSQL

    zlBulidingPriceDataFromRegistNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
Public Function zlRegisterPriceDeleteFromNO(ByVal strNO As String, ByVal lng科室ID As Long, ByVal str医生姓名 As String, _
    Optional blnNotExcuteSaveData As Boolean = False, Optional ByRef bytErrType_Out As Byte, _
    Optional ByRef cllProSQL As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据挂号单据，删除对应的划价单（必须满足病人挂号记录．收费单不为空）
    '入参:strNO-挂号单号
    '     lng科室ID-科室ID,如果为０，则不检查此条件
    '     str医生姓名-医生姓名：为空串时，表示不检查此条件
    '     blnNotExcuteSaveData-true:不执行数据保存，直接返回执行的SQL集,False-执行数据保存
    '出参:cllProSQL:blnNotExcuteSaveData=true时，则返回需要执行的SQL集，否则返回Nothing
    '     bytErrType_Out-错误类型:0-正常执行;1-未找到挂号单;2-未生成划价单;3-未找到符合条件的划价单;4-存在已经收费的单据
    '返回:执行成功或获取保存的SQL数据成功,返回true,检查或保存失败返回Fale
    '编制:刘兴洪
    '日期:2018-01-12 10:53:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, str收费单 As String
    Dim rsTemp As ADODB.Recordset
    Dim strWhere As String
    Dim blnTrans As Boolean
    
    On Error GoTo errHandle
    If cllProSQL Is Nothing Then
       Set cllProSQL = New Collection
    End If
    
    bytErrType_Out = 0
    strSQL = "Select 收费单 From 病人挂号记录 A  Where A.NO=[1] and a.记录状态 in (0,1,3) "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(删除挂号单对应的划价单）", strNO)
    If rsTemp.EOF Then
        bytErrType_Out = 1
        zlRegisterPriceDeleteFromNO = False: Exit Function   '未找到划价单
    End If
    
    If Nvl(rsTemp!收费单) = "" Then
        bytErrType_Out = 2
        Exit Function
    End If
    
    str收费单 = Nvl(rsTemp!收费单)
    strWhere = ""
    
    If lng科室ID <> 0 Then strWhere = strWhere & " and A.执行部门ID=[2] "
    If str医生姓名 <> "" Then strWhere = strWhere & "  and 执行人=[3] "
     
     
     strSQL = " " & _
    "   Select /*+cardinality(b,10)*/  a.No, Max(a.结帐id) as 结帐ID " & _
    "   From 门诊费用记录 A, Table(f_Str2list([1])) B " & _
    "   Where a.记录性质 = 1 And a.记录状态 In (0, 1, 3) and a.No = b.column_value " & strWhere & _
    "   group by a.No"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(检查挂号单是否收费）", str收费单, lng科室ID, str医生姓名)
    If rsTemp.EOF Then
        bytErrType_Out = 3
        Exit Function
    End If
    With rsTemp
        Do While Not .EOF
            If Val(Nvl(rsTemp!结帐ID)) <> 0 Then
                bytErrType_Out = 4: Exit Function
            End If
            ' Zl_门诊划价记录_Delete
            strSQL = "Zl_门诊划价记录_Delete("
            '  No_In   门诊费用记录.No%Type,
            strSQL = strSQL & "'" & Nvl(rsTemp!NO) & "')"
            '  序号_In Varchar2:=Null - -主要用于门诊医生站作废单个药品
            zlAddArray cllProSQL, strSQL
            rsTemp.MoveNext
        Loop
    End With
            
    If Not blnNotExcuteSaveData Then
        blnTrans = True
        Call zlExecuteProcedureArrAy(cllProSQL, "zlPublicExpense.clsRegist(删除划价单）")
        blnTrans = False
    End If
    zlRegisterPriceDeleteFromNO = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function


Public Function zlGetRegisterBookData(ByVal lng病人ID As Long, ByRef rsBookData_Out As ADODB.Recordset, _
    Optional strBookNo As String = "", Optional str科室IDs As String = "", Optional bytType As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID或预约单获取预约挂号单数据
    '入参:lng病人ID-病人ID
    '     strBookNo-预约单号
    '     str科室IDs：不为空时，只读取指定科室的预约单或挂号单
    '     bytType-0-仅包含预约单据;1-仅包含已经支付但没在签道的预约单;2-仅包含回诊病人;3-包含(0,1,2)
    '出参:rsBookData_Out-预约单数据
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-15 15:53:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errHandle
    Set rsBookData_Out = Nothing
'
'    If mSysPara.int预约失效次数 > 0 Then
'       strWhere = "  And A.发生时间 between trunc(sysdate) and  trunc(sysdate)+1-1/24/60/60 "
'    Else
       strWhere = " And ((nvl(A.加班标志,0) =0 And A.发生时间 > Trunc(Sysdate) - [3]) or  (nvl(A.加班标志,0) =1 And A.发生时间 > Trunc(Sysdate) - [4])) "
'    End If
    
    If str科室IDs <> "" Then strWhere = strWhere & " And A.执行部门ID In (Select  /*+cardinality(J,10) */ Column_Value From Table(f_num2list([5]))) "

    If lng病人ID <> 0 Then strWhere = strWhere & "  And A.病人ID = [1]"
    If strBookNo <> "" Then strWhere = strWhere & "   And A.NO = [2]"
    
    
    Select Case bytType
    Case 3  '所有:3-包含(0,1,2)
        strWhere = strWhere & " And (A.记录状态=0 or (A.记录状态=1 And nvl(M.记录标志,0) in (0,2) ))"
    Case 2  '仅包含回诊病人
        strWhere = strWhere & " And   nvl(M.记录标志,0) =2"
    Case 1  '仅包含已经支付但没在签道的预约单
        strWhere = strWhere & " And  A.记录状态=1 And nvl(M.记录标志,0)=0"
    Case Else  '0 Then
        strWhere = strWhere & " And  A.记录状态=0"
    End Select
    
    strSQL = "" & _
    " Select distinct A.病人ID, A.NO,nvl(M.号别,A.计算单位) as 号别, " & _
    "           A.执行部门id as 科室ID,C.编码 as 科室编码,C.名称 as  科室名称,C.简码 as 科室简码, " & _
    "           J.ID as 医生ID,J.编号 as 医生编号,J.姓名 as  医生姓名,J.简码 as 医生简码,M.诊室," & _
    "           A.收费细目ID as 项目ID,D.编码 as 项目编码,D.名称 as 挂号项目,   " & _
    "           to_char(nvl(M.预约时间,A.发生时间),'yyyy-mm-dd hh24:mi:ss') as 预约时间,B.身份证号,M.出诊记录ID,A.记录状态,M.ID as 挂号ID,M.记录标志 " & _
    " From 门诊费用记录 A,病人挂号记录 M, 病人信息 B,部门表 C,收费项目目录 D,人员表 J" & _
    " Where A.NO=M.NO And A.记录性质 = 4  and A.序号=1 and nvl(M.执行状态,0)<>1 and nvl(M.执行状态,0)<>-1" & _
    "       And A.执行部门ID=C.ID  And (C.站点='" & gstrNodeNo & "' Or C.站点 is NULL) " & _
    "       And A.执行人=J.姓名(+) And (A.执行人 is NULL or exists(Select 1 From 人员表 Where A.执行人=姓名 And ( 站点='" & gstrNodeNo & "' or 站点 is null)))" & _
    "       And A.病人id = B.病人id(+) And  a.收费细目Id=d.ID(+) " & strWhere
      
    '没有设置黑名单,保留以前的处理方式,否则只能取当天的预约单(如果失了约的,则以红色字体显示)
    Set rsBookData_Out = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(获取预约单据）", lng病人ID, strBookNo, mSysPara.byt普通有效天数, mSysPara.byt急诊有效天数, str科室IDs)
    zlGetRegisterBookData = True
    
   Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
Public Function zlPatiEdit(ByVal frmMain As Object, ByVal lng病人ID As Long, ByRef lng病人ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:新建病信息
    '入参:frmMain-调用的主窗体
    '     lng病人ID-传入更改的病人ID
    '出参:lng病人ID_Out-返回病人信息
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-17 14:58:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmTemp As frmPatiInfo
    Dim bytFun As Byte
    On Error GoTo errHandle
    
    
    Set frmTemp = New frmPatiInfo
     '0-编辑或查看病人信息;2-新建档病人
    bytFun = IIf(lng病人ID = 0, 2, 0)
    zlPatiEdit = frmTemp.ShowMe(frmMain, bytFun, lng病人ID, lng病人ID_Out)
    If Not frmTemp Is Nothing Then Unload frmTemp
    Set frmTemp = Nothing
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function
Public Function zlGetRegisterMaxDaysFromDeptAndDoctor_Tradition(ByVal str科室IDs As String, ByVal str医生姓名 As String, ByVal bln含本科号 As Boolean, _
    ByRef intMaxDays_Out As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取最大的预约天数(传统模式)
    '入参:str科室IDs-科室ID,多个用逗号分离
    '     str医生姓名-医生姓名,为空时，按科室确定
    '     bln含本科号-包含本科的号，即医生ID is null 或医生 is NULL
    '出参:intMaxDays_Out-最大的预约天数
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-30 14:31:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String
    intMaxDays_Out = mSysPara.int允许预约天数
    On Error GoTo errHandle
    
    strWhere = ""
    If str医生姓名 <> "" Then
        If Not bln含本科号 Then
            strWhere = strWhere & " And a.医生姓名=[2]"
        Else
            strWhere = strWhere & " And ( a.医生姓名=[2] Or (A.医生ID is null and A.医生姓名 is null) )"
        End If
    End If
    
    strSQL = "" & _
    " Select /*+cardinality(B,10) */ Max(A.预约天数) as 预约天数   " & _
    " From 挂号安排 A,table(f_num2List([1])) B " & _
    " where a.科室ID=b.column_value" & strWhere
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(获取最大的预约天数）", str科室IDs, str医生姓名)
    
    If Val(Nvl(rsTemp!预约天数)) <> 0 Then
        intMaxDays_Out = Val(Nvl(rsTemp!预约天数))
    End If
    
    zlGetRegisterMaxDaysFromDeptAndDoctor_Tradition = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    zlGetRegisterMaxDaysFromDeptAndDoctor_Tradition = True
End Function
Public Function zlGetRegisterMaxDaysFromDeptAndDoctor_Visits(ByVal str科室IDs As String, ByVal str医生姓名 As String, ByVal bln含本科号 As Boolean, _
    ByRef intMaxDays_Out As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取最大的预约天数(出诊表模式)
    '入参:str科室IDs-科室ID,多个用逗号分离
    '     str医生姓名-医生姓名,为空时，按科室确定
    '     bln含本科号-包含本科的号，即医生ID is null 或医生 is NULL
    '出参:intMaxDays_Out-最大的预约天数
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-30 14:31:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String
    intMaxDays_Out = mSysPara.int允许预约天数
    On Error GoTo errHandle
    
    strWhere = ""
    If str医生姓名 <> "" Then
        If Not bln含本科号 Then
            strWhere = strWhere & " And a.医生姓名=[2]"
        Else
            strWhere = strWhere & " And (a.医生姓名=[2] Or (A.医生ID is null and A.医生姓名 is null) )"
        End If
    End If
    
    strSQL = "" & _
    " Select /*+cardinality(B,10) */ Max(A.预约天数) as 预约天数   " & _
    " From 临床出诊号源 A,table(f_num2List([1])) B " & _
    " where a.科室ID=b.column_value"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(获取最大的预约天数）", str科室IDs, str医生姓名)
    
    If Val(Nvl(rsTemp!预约天数)) <> 0 Then
        intMaxDays_Out = Nvl(rsTemp!预约天数)
    End If
    zlGetRegisterMaxDaysFromDeptAndDoctor_Visits = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    zlGetRegisterMaxDaysFromDeptAndDoctor_Visits = True
End Function
Public Function zlGetRegisterMaxDaysFromDeptAndDoctor(ByVal str科室IDs As String, ByVal str医生姓名 As String, ByVal bln含本科号 As Boolean, _
    ByRef intMaxDays_Out As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取最大的预约天数
    '入参:str科室IDs-科室ID,多个用逗号分离
    '     str医生姓名-医生姓名,为空时，按科室确定
    '     bln含本科号-包含本科的号，即医生ID is null 或医生 is NULL
    '出参:intMaxDays_Out-最大的预约天数
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-30 14:31:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytRegMode As Byte, strErrMsg As String
    On Error GoTo errHandle
    intMaxDays_Out = mSysPara.int允许预约天数
    If zlGetRegisterMode("", bytRegMode, strErrMsg) = False Then zlGetRegisterMaxDaysFromDeptAndDoctor = True '以缺省预约天数为准
    
    If bytRegMode = 0 Then
          zlGetRegisterMaxDaysFromDeptAndDoctor = zlGetRegisterMaxDaysFromDeptAndDoctor_Tradition(str科室IDs, str医生姓名, bln含本科号, intMaxDays_Out)
          Exit Function
    End If
    zlGetRegisterMaxDaysFromDeptAndDoctor = zlGetRegisterMaxDaysFromDeptAndDoctor_Visits(str科室IDs, str医生姓名, bln含本科号, intMaxDays_Out)
    Exit Function
errHandle:
    zlGetRegisterMaxDaysFromDeptAndDoctor = True
 End Function
 
Public Function zlGetRegisterNextSn(ByVal lng安排ID As Long, ByVal strRegDate As String, _
    ByVal bln加号 As Boolean, ByVal bln预约 As Boolean, ByVal bln锁号 As Boolean, ByRef lngSn_Out As Long, ByRef str开始时间_Out As String, _
    Optional str备注 As String, Optional str机器名 As String, Optional ByVal str合作单位 As String, Optional ByVal lng计划ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一下个序号
    '入参:lng安排ID -安排ID(出表模式时，传入记录ID,否则传入安排ID)
    '     lng计划ID-计划ID（只有传统模式时，才传入)
    '出参:lngSn_Out-返回序号
    '     str开始时间_Out-返回分时段的开始时间
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-04 14:35:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strWhere As String, strValue As String
    Dim dtDate As Date, bytMode As Byte
    

    On Error GoTo errHandle
    
    lngSn_Out = 0: str开始时间_Out = "":
    If zlGetRegisterMode(strRegDate, bytMode) = False Then Exit Function
    
    If bytMode = 0 Then
        zlGetRegisterNextSn = zlGetRegisterNextSn__Tradition(lng安排ID, lng计划ID, strRegDate, bln加号, bln预约, bln锁号, lngSn_Out, str开始时间_Out, str备注, str机器名, str合作单位)
        Exit Function
    End If
    zlGetRegisterNextSn = zlGetRegisterNextSn__Visits(lng安排ID, strRegDate, bln加号, bln预约, bln锁号, lngSn_Out, str开始时间_Out, str备注, str机器名, str合作单位)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
 End Function
Public Function zlGetRegisterNextSn__Tradition(ByVal lng安排ID As Long, ByVal lng计划ID As Long, ByVal strRegDate As String, _
    ByVal bln加号 As Boolean, ByVal bln预约 As Boolean, ByVal bln锁号 As Boolean, ByRef lngSn_Out As Long, ByRef str开始时间_Out As String, _
    Optional str备注 As String, Optional str机器名 As String, Optional ByVal str合作单位 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一下个序号
    '入参:lng安排ID -安排ID
    '     lng计划ID-计划ID
    '出参:lngSn_Out-返回序号
    '     str开始时间_Out-返回分时段的开始时间
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-04 14:35:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strWhere As String, strValue As String
    Dim dtDate As Date

    On Error GoTo errHandle
    
    lngSn_Out = 0: str开始时间_Out = "":
    
    If strRegDate = "" Then
        strRegDate = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd HH:MM:SS")
    End If
    dtDate = CDate(strRegDate)
    
    strSQL = "Select Zl_Fun_挂号安排_传统_Nextsn([1],[2],[3],[4],NULL,[5],[6],[7],[8],[9],[10]) as 序号信息  From dual "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(获取一下个序号）", dtDate, lng安排ID, lng计划ID, UserInfo.姓名, str备注, _
                                    str机器名, str合作单位, IIf(bln加号, 1, 0), IIf(bln预约, 1, 0), IIf(bln锁号, 1, 0))
    'Zl_Fun_挂号安排_传统_Nextsn
    '(
    '  日期_In       In Date,
    '  安排id_In     In 挂号安排.Id%Type,
    '  计划id_In     In 挂号安排计划.Id%Type,
    '  操作员姓名_In 人员表.姓名%Type,
    '  星期_In       In 挂号安排限制.限制项目%Type := Null,
    '  备注_In       In 挂号序号状态.备注%Type := Null,
    '  机器名_In     In 挂号序号状态.机器名%Type := Null,
    '  合作单位_In   In Varchar2 := Null,
    '  是否加号_In   In Number := 0,
    '  是否预约_In   In Number := 0,
    '  是否锁号_In   In Number := 1
    strValue = rsTemp!序号信息
    lngSn_Out = Split(strValue & "||", "|")(0)
    str开始时间_Out = Split(strValue & "||", "|")(1)
    If str开始时间_Out = "" Then str开始时间_Out = strRegDate
    zlGetRegisterNextSn__Tradition = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
 End Function
Public Function zlGetRegisterNextSn__Visits(ByVal lng记录ID As Long, ByVal strRegDate As String, _
    ByVal bln加号 As Boolean, ByVal bln预约 As Boolean, ByVal bln锁号 As Boolean, ByRef lngSn_Out As Long, ByRef str开始时间_Out As String, _
    Optional str备注 As String, Optional str机器名 As String, Optional ByVal str合作单位 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一下个序号
    '入参:lng记录ID
    '出参:lngSn_Out-返回序号
    '     str开始时间_Out-返回分时段的开始时间
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-04 14:35:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strWhere As String, strValue As String
    Dim dtDate As Date

    On Error GoTo errHandle
    
    lngSn_Out = 0: str开始时间_Out = "":
    
    If strRegDate = "" Then
        strRegDate = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd HH:MM:SS")
    End If
    dtDate = CDate(strRegDate)
    
    strSQL = "Select Zl_Fun_挂号安排_出诊表_Nextsn([1],[2],[3],[4],[5],[6],[7],[8],[9]) as 序号信息  From dual "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "zlPublicExpense.clsRegist(获取一下个序号）", dtDate, lng记录ID, UserInfo.姓名, IIf(bln加号, 1, 0), IIf(bln预约, 1, 0), IIf(bln锁号, 1, 0), str备注, _
                                    str机器名, str合作单位)
    ' Zl_Fun_挂号安排_出诊表_Nextsn
    '(
    '  日期_In       In Date,
    '  记录id_In     In 临床出诊记录.Id%Type,
    '  操作员姓名_In 人员表.姓名%Type,
    '  是否加号_In   In Number := 0,
    '  是否预约_In   In Number := 0,
    '  是否锁号_In   In Number := 1,
    '  备注_In       In 挂号序号状态.备注%Type := Null,
    '  机器名_In     In 挂号序号状态.机器名%Type := Null,
    '  合作单位_In   In Varchar2 := Null
    strValue = rsTemp!序号信息
    lngSn_Out = Split(strValue & "||", "|")(0)
    str开始时间_Out = Split(strValue & "||", "|")(1)
    If str开始时间_Out = "" Then str开始时间_Out = strRegDate
    zlGetRegisterNextSn__Visits = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlReserveRegNo(ByVal str号码 As String, ByVal bln严格控制 As Boolean, ByVal bln分时段 As Boolean, _
                            ByVal strTime As String, ByRef lng序号 As Long, ByVal str备注 As String, _
                            Optional ByVal lng记录ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 对当前序号锁号
    ' 入参 :str号码-号别
    '       bln预约-是否预约锁号
    '       bln严格控制-是否严格控制
    '       bln分时段-是否分时段
    '       lng序号需要锁号的序号
    '       lng记录ID - 出诊排班模式需要传入记录id
    ' 出参 :lng序号_Out如果lng序号被锁定，则取新的序号锁号，并返回新的序号
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/2/18 15:34
    '---------------------------------------------------------------------------------------
    zlReserveRegNo = ReserveRegNo(str号码, bln严格控制, bln分时段, strTime, lng序号, str备注, lng记录ID)
End Function

Public Sub zlCancelRegNo(Optional ByVal lng记录ID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:取消挂号时删除锁定挂号序号
    '入参 :lng记录ID-出诊记录ID
    '编制:李南春
    '日期:2019/2/18 15:34
    '-----------------------------------------------------------------------------------------------------------------------
    Call CancelRegNo(lng记录ID)
End Sub

Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()

    Err = 0: On Error Resume Next
    If Not mfrmRegist Is Nothing Then Unload mfrmRegist
    Set mfrmRegist = Nothing
    If mrsWorkTime Is Nothing Then Set mrsWorkTime = Nothing

    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
