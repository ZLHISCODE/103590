VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VERSION As String

Private WithEvents mfrmQueueStation As frmQueueStation
Attribute mfrmQueueStation.VB_VarHelpID = -1
Private WithEvents mfrmUpdateInfo As frmUpdateInfo
Attribute mfrmUpdateInfo.VB_VarHelpID = -1

Public Event OnQueueRoomLoad(ByVal str业务ID As String, rsRoomData As ADODB.Recordset, rsDoctorData As ADODB.Recordset)
Public Event OnQueueChange(ByVal lngQueueId As Long, ByVal strQueue As string, ByVal strPatient As string, ByVal strRoom As String, ByVal strDoctor As String, ByRef blnIsAllowChange As Boolean, ByRef blnIsAlreadyProcess As Boolean)
Public Event OnQueueExecuteBefore(ByVal str业务ID As String, ByVal byt操作类型 As Byte, blnCancel As Boolean, strNewQueueName As String)
Public Event OnQueueExecuteAfter(ByVal str业务ID As String, ByVal byt操作类型 As Byte)
Public Event OnRecevieDiagnose(ByVal str业务ID As String, ByVal lng业务类型 As Long)
Public Event OnSelectionChanged(ByVal blnIsCallingList As Boolean, objReportRow As Object, cbrMain As Object)


Property Get Enable() As Boolean
    Enable = mfrmQueueStation.Enabled
End Property

Property Let Enable(value As Boolean)
    mfrmQueueStation.Enabled = value
End Property

'是否显示查找工具栏
Public Property Get IsShowFindTools() As Boolean
    IsShowFindTools = mfrmQueueStation.mblnIsShowFindTools
End Property

Public Property Let IsShowFindTools(ByVal vNewValue As Boolean)
    mfrmQueueStation.mblnIsShowFindTools = vNewValue
End Property

Public Function FindQueue(ByVal strLocateType As String, ByVal strLocateValue As String) As Boolean
    On Error GoTo errHandle
    
    FindQueue = mfrmQueueStation.FindQueue(strLocateType, strLocateValue)
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

'Public Sub ShwoMeTest(cnOracle As ADODB.Connection)
'
'    Dim strQueue(2) As String
'
'    '队列的下标从1开始
'    strQueue(1) = "队列1"
'    strQueue(2) = "队列2"
'
'    mfrmQueueStation.zlShowMe cnOracle, strQueue, "队列1", 0
'End Sub


Public Sub zlInitVar(cnOracle As ADODB.Connection, Optional lngSys As Long = 100, _
    Optional int业务类型 As Integer = 0, Optional intValidDays As Integer = 1, _
    Optional strPrivs As String = "", Optional strOption As String = "", Optional blnIsGroup As Boolean = True)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化系统参数
    '入参：strOption-暂留,以后扩展
    '编制：刘兴洪
    '日期：2010-06-11 11:01:09
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If Not mfrmQueueStation Is Nothing Then
        Call mfrmQueueStation.zlInitVar(cnOracle, lngSys, int业务类型, intValidDays, strPrivs, strOption, blnIsGroup)
    End If
End Sub

'设置焦点
Public Sub SetFocus()
    On Error Resume Next
    
    If Not mfrmQueueStation Is Nothing Then
        If mfrmQueueStation.Visible Then
            If mfrmQueueStation.rptQueueList.Visible Then Call mfrmQueueStation.rptQueueList.SetFocus
        End If
    End If
    
    On Error GoTo 0
End Sub


Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function


Public Sub zlCommandBarSet(ByVal intFuncType As Integer, ByVal blnUseState As Boolean)
'************************************************************************************
'
'设置功能状态
'
'intFuncType：功能类型 0-复诊，1-直呼/顺呼，2-弃号 ，3-暂停，4-完成就诊，5,-广播
'blnUseState：是否启用
'
'************************************************************************************
    If Not mfrmQueueStation Is Nothing Then
        Call mfrmQueueStation.zlCommandBarSet(intFuncType, blnUseState)
    End If
End Sub


Private Sub Class_Initialize()
    VERSION = App.Major & "." & App.Minor & "." & App.Revision
    Set mfrmQueueStation = New frmQueueStation
    Set mfrmUpdateInfo = frmUpdateInfo
    mfrmQueueStation.mblnIsShowFindTools = True
End Sub


Private Sub Class_Terminate()
    Unload mfrmQueueStation
    Set mfrmQueueStation = Nothing
    Set mfrmUpdateInfo = Nothing
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = mfrmQueueStation
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object)
    Call mfrmQueueStation.zlDefCommandBars(CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmQueueStation.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmQueueStation.zlExecuteCommandBars(Control)
End Sub

Public Sub zlSetToolIcon(ByVal intIconSize As Integer, ByVal blnIsDisplayText As Boolean)
  Call mfrmQueueStation.zlSetToolIcon(intIconSize, blnIsDisplayText)
End Sub

Public Function zlRefresh(str队列名称() As String, ByVal strCur队列名称 As String, ByVal strCur业务ID As String, _
    Optional str诊室 As String = "", Optional str医生 As String = "", Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '功能：调用刷新指定医嘱id的报告内容，并根据情况提供编辑功能
    '入参：str队列名称():传入的指定队列数组(从1开始)
    '         strCur队列名称-当前队列名称
    '         lngCur业务ID-业务ID
    '         str诊室-限制为指定的诊室,可以为多个诊室:如"一诊室,二诊室,..."
    '         str医生-限制为制定的医生,可以传多个医生,用逗号分隔,如"张三,李四,..."
    '         strExcludeData-排队的指定业务ID
    '         intViewDataType数据显示类型，0显示当前科室下的所有数据，
    '                                      1显示诊室为当前诊室且医生姓名为空，或者医生姓名等于当前医生，或者诊室为空和医生为空的数据
    '                                      2显示诊室为当前诊室和医生姓名为空或医生姓名等于当前医生的数据
    '                                      3显示当前医生的数据
    '编制：刘兴洪
    '日期：2010-06-11 20:54:55
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If Not mfrmQueueStation Is Nothing Then
        'Call DebugMsg("ZLQueueInterface", "ZLRefresh", "str诊室=" & str诊室 & ", str医生=" & str医生 & ", strExcludeData=" & strExcludeData & ", intViewDataType=" & intViewDataType)
        zlRefresh = mfrmQueueStation.zlRefresh(str队列名称, strCur队列名称, strCur业务ID, str诊室, str医生, strExcludeData, intViewDataType)
    End If
End Function

Private Function To_Date(ByVal dat日期 As Date) As String
'功能:将入参中的日期传换成ORACLE需要的日期格式串
    To_Date = "To_Date('" & Format(dat日期, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Function zlInQueue(ByVal str队列名称 As String, ByVal lng业务类型 As Long, ByVal str业务ID As String, ByVal lng科室ID As Long, _
        ByVal str患者姓名 As String, ByVal lng病人id As Long, ByVal str诊室 As String, ByVal str医生姓名 As String, _
        Optional str排队标记 As String = "", Optional str排队号码 As String = "", Optional dat排队日期 As Date) As Boolean
'*************************************************************************************
'
'将需要进行叫号操作的数据入队
'
'str队列名称：
'
'lng业务类型：
'
'lng业务ID：
'
'lng科室ID：
'
'str患者姓名：
'
'str诊室：
'
'str医生姓名：
'
'str排队标记：
'
'lng排队号码：
'
'dat排队日期：
'
'*************************************************************************************
    
    Dim strSql As String
    Dim datTest As Date
    
    On Error GoTo err
    
    If dat排队日期 = datTest Then
      dat排队日期 = zlDatabase.Currentdate
    End If
        
    strSql = "ZL_排队叫号队列_INSERT('" & str队列名称 & "', " & lng业务类型 & ",'" & str业务ID & "'," & lng科室ID & ",'" & str排队号码 & "','" & str排队标记 & "','" _
             & str患者姓名 & "'," & lng病人id & ",'" & str诊室 & "','" & str医生姓名 & "',to_date('" & Format(dat排队日期, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"

    zlDatabase.ExecuteProcedure strSql, "插入队列"
    zlInQueue = True
    Exit Function
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub zlQueueExec(ByVal str当前队列名 As String, ByVal lng业务类型 As Long, ByVal str业务ID As String, ByVal byt操作类型 As Byte)
'************************************************************************************
'顺呼，对排队的呼叫状态进行更新

'str当前队列名：入参，字符类型。需要操作的队列名称。
'lng业务类型：入参，数字类型。目前门诊排队为0，PACS排队为1，LIS排队为2。
'lng业务ID：入参，字符类型。标记排队记录跟业务数据的关系，比如门诊中填写门诊号，PACS中填写医嘱ID。
'byt操作类型：入参，数字类型。叫号操作的类型 0-复诊，1-直呼/顺呼（Lng业务ID=0为顺呼），2-弃号 ，3-暂停，4-完成就诊，5,-广播, 6,回诊。
    

'*************************************************************************************
        Call mfrmQueueStation.zlQueueExec(str当前队列名, lng业务类型, str业务ID, byt操作类型)
End Sub








Public Function zlLISQueueJoin(ByVal str队列名称 As String, ByVal lng业务类型 As Long) As Long
'************************************************************************************
'将需要进行叫号操作的数据入队，并返回最大排队号码

'str队列名称：入参，字符类型。需要操作的队列名称。
'lng业务类型：入参，数字类型。目前门诊排队为0，PACS排队为1，LIS排队为2。
'lng科室ID：入参，数字类型，科室编号
'dat排队日期：可选入参，数据入队时间。

    
'编制：罗冠骁
'日期：2012-6-26
'说明：这里使用了lngWorkOrderID来代替业务ID。如果没有查找到排队号码数据，那么接口返回的排队号码值为0
'*************************************************************************************

    Dim strSql As String
    Dim datTest As Date
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
        
    strSql = "ZL_排队叫号队列_INSERT('" & str队列名称 & "', " & lng业务类型 & "," & Val(Format(zlDatabase.Currentdate, "HHMMSS")) & " ,null,'','','',null,'','',to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSql, "插入队列"
    
    
    strSql = "select max(排队号码) as 排队号码 from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and 排队状态=0 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "得到当前病人的排队最大号码", str队列名称, lng业务类型)
    
    zlLISQueueJoin = Nvl(rsTemp!排队号码, 0)
    
    Exit Function
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlGetQueuePatientCount(ByVal str队列名称 As String, ByVal lng业务类型 As Long) As Long
'************************************************************************************
'得到指定队列中等待病人数量

'str队列名称：入参，字符类型。需要操作的队列名称。
'lng业务类型：入参，数字类型。目前门诊排队为0，PACS排队为1，LIS排队为2。
    
'编制：罗冠骁
'日期：2012-6-26
'说明：如果没有查找到数据，那么接口返回的等待病人数量为0
'*************************************************************************************

 Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSql = "select count(*) as 等待病人数量 from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and 排队状态=0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "得到等待病人数量", str队列名称, lng业务类型)
    
    zlGetQueuePatientCount = Nvl(CLng(rsTemp!等待病人数量) - 1, 0)
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub zlQueueParameterSetup(frm父窗体 As Object, lng系统号 As Long)
'************************************************
'打开排队叫号参数配置界面

'frm父窗体：入参，窗体对象。打开配置界面需要的父窗体对象。
'lng系统号：入参，数字类型。当前系统号。
    
'编制：罗冠骁
'日期：2012-6-26
'说明：直接调用参数配置界面
'************************************************
    
    Call mfrmQueueStation.QueueParameterSetup(frm父窗体, lng系统号)

End Sub


Public Function zlQueueOrderCall(ByVal str当前队列名 As String, ByVal lng业务类型 As Long, ByVal byt操作类型 As Byte) As Long
'************************************************************************************
'顺呼，对排队的呼叫状态进行更新

'str当前队列名：入参，字符类型。需要操作的队列名称。
'lng业务类型：入参，数字类型。目前门诊排队为0，PACS排队为1，LIS排队为2。
'byt操作类型：入参，数字类型。叫号操作的类型 0-复诊，1-直呼/顺呼（Lng业务ID=0为顺呼），2-弃号 ，3-暂停，4-完成就诊，5,-广播, 6,回诊。
    
'编制：罗冠骁
'日期：2012-6-26
'说明：
'*************************************************************************************
    
    Dim strSql As String
    Dim rsWorkID As ADODB.Recordset
    Dim rsQueueNo As ADODB.Recordset
    
    On Error GoTo err
    
    strSql = "select 业务ID from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and 排队状态=0 and rownum=1 order by 排队号码 "
    Set rsWorkID = zlDatabase.OpenSQLRecord(strSql, "得到等待病人数量", str当前队列名, lng业务类型)
    
    If rsWorkID.RecordCount < 1 Then
        zlQueueOrderCall = 0
        Exit Function
    End If
    
    '查找当前操作的排队号码
    strSql = "select min(排队号码) as 排队号码 from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and 排队状态=0 "
    Set rsQueueNo = zlDatabase.OpenSQLRecord(strSql, "得到当前病人的排队号码", str当前队列名, lng业务类型)
    
     Call mfrmQueueStation.zlQueueExec(str当前队列名, lng业务类型, rsWorkID!业务ID, byt操作类型)
    
    zlQueueOrderCall = Nvl(rsQueueNo!排队号码, 0)

    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlQueueBroadcastCall(ByVal str呼叫文本 As String) As Boolean
'*************************************************************************************
'对呼叫文本进行广播呼叫

'str队列名称：入参，字符类型。需要操作的队列名称。
'lng业务ID：入参，字符类型。标记排队记录跟业务数据的关系，比如门诊中填写门诊号，PACS中填写医嘱ID。

'修改：罗冠骁
'日期：2012-6-27
'说明：如果呼叫成功返回True 失败返回false
'*************************************************************************************
    
    '调用广播呼叫函数
    zlQueueBroadcastCall = mfrmQueueStation.QueueBroadcastCall(str呼叫文本)

End Function


Public Function zlDelQueue(ByVal str队列名称 As String, Optional ByVal str业务ID As String = "") As Boolean
'*************************************************************************************
'从队列中删除排队信息

'str队列名称：入参，字符类型。需要操作的队列名称。
'lng业务ID：入参，字符类型。标记排队记录跟业务数据的关系，比如门诊中填写门诊号，PACS中填写医嘱ID。

'修改：罗冠骁
'日期：2012-6-26
'说明：如果没有查找到数据，那么接口返回的排队号码为0
'*************************************************************************************
    
    Dim strSql As String
    
    On Error GoTo err
    
    strSql = "ZL_排队叫号队列_DELETE('" & str队列名称 & "','" & str业务ID & "')"
    zlDatabase.ExecuteProcedure strSql, "删除排队信息"
    zlDelQueue = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function






Public Function zlUpdateQueue(ByVal str队列名称 As String, ByVal str业务ID As String, ByVal lng科室ID As Long, _
        ByVal str患者姓名 As String, ByVal str诊室 As String, ByVal str医生姓名 As String) As Boolean
'*************************************************************************************
'
'更新队列中的信息
'
'str队列名称：
'
'lng业务ID：
'
'lng科室ID：
'
'str患者姓名：
'
'str诊室：
'
'str医生姓名：
'
'*************************************************************************************
    
    Dim strSql As String
    
    On Error GoTo err
  
    strSql = "ZL_排队叫号队列_UPDATE('" & str队列名称 & "','" & str业务ID & "'," & lng科室ID & ",'" _
            & str患者姓名 & "','" & str诊室 & "','" & str医生姓名 & "')"
    zlDatabase.ExecuteProcedure strSql, "修改排队信息"
    zlUpdateQueue = True
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetQueueBusinessDataIDs() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取业务IDs
    '入参:bytType-0-挂号;1...
    '出参:
    '返回:成功返回业务IDs,多个用逗号分离,如:22,33,44
    '编制:刘兴洪
    '日期:2014-03-11 16:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
 

    GetQueueBusinessDataIDs = mfrmQueueStation.GetQueueBusinessDataIDs()

End Function


Private Sub mfrmQueueStation_OnQueueExecuteAfter(ByVal str业务ID As String, ByVal byt操作类型 As Byte)
    RaiseEvent OnQueueExecuteAfter(str业务ID, byt操作类型)
End Sub

Private Sub mfrmQueueStation_OnQueueExecuteBefore(ByVal str业务ID As String, ByVal byt操作类型 As Byte, blnCancel As Boolean, strNewQueueName As String)
    RaiseEvent OnQueueExecuteBefore(str业务ID, byt操作类型, blnCancel, strNewQueueName)
End Sub

Private Sub mfrmQueueStation_OnQueueRoomLoad(ByVal str业务ID As String, rsRoomData As ADODB.Recordset, rsDoctorData As ADODB.Recordset)
    RaiseEvent OnQueueRoomLoad(str业务ID, rsRoomData, rsDoctorData)
End Sub

Private Sub mfrmQueueStation_OnRecevieDiagnose(ByVal str业务ID As String, ByVal lng业务类型 As Long)
    RaiseEvent OnRecevieDiagnose(str业务ID, lng业务类型)
End Sub

Private Sub mfrmQueueStation_OnSelectionChanged(ByVal blnIsCallingList As Boolean, objDataRow As XtremeReportControl.ReportRow, cbrMain As XtremeCommandBars.CommandBars)
    RaiseEvent OnSelectionChanged(blnIsCallingList, objDataRow, cbrMain)
End Sub

Private Sub mfrmUpdateInfo_OnQueueChange(ByVal lngQueueId As Long, ByVal strQueue As string, ByVal strPatient As string, ByVal strRoom As String, ByVal strDoctor As String, ByRef blnIsAllowChange As Boolean, ByRef blnIsAlreadyProcess As Boolean)
    RaiseEvent OnQueueChange(lngQueueId, strQueue, strPatient, strRoom, strDoctor, blnIsAllowChange, blnIsAlreadyProcess)
End Sub
