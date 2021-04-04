VERSION 5.00
Object = "*\A..\queueOper\zlQueueOper.vbp"
Begin VB.Form frmWork_Queue 
   BorderStyle     =   0  'None
   Caption         =   "排队叫号管理"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11595
   Icon            =   "frmWork_Queue.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin zlQueueOper.UcQueue ucPacsQueue 
      Height          =   5085
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   8969
      Interval        =   30000
      ValidDays       =   0
   End
End
Attribute VB_Name = "frmWork_Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const M_LNG_PACS_BUSINESS_IMG_TYPE As Long = 1                  'pacs影像医技业务类型定义
Private Const M_LNG_PACS_BUSINESS_CAP_TYPE As Long = 1                  'pacs视频采集业务类型定义

Private Const M_STR_NOT_ALLOT_TECHNIC As String = "科室队列"      '未分配队列名称定义
Private Const M_STR_FINDWAY_EX As String = "门诊号,住院号,就诊卡,医保号"


Private mrsPacsQueueGroupConfig As ADODB.Recordset
Private mrsPacsQueueTechnicConfig As ADODB.Recordset


Private mlngCurDeptId As Long                       '当前科室ID
Private mstrCurDeptName As String                   '当前科室名称
Private mstrQueryTechnicQueueNames  As String       'pacs排队叫号查询队列名称
Private mlngQueueNoWay As Long                      '排队号码生成方式
Private mlngValidDays As Long
Private mstrReportNum As String
Private mlngPrintWay As Long
Private mblnUseQueueMsg As Boolean

Private mstrQueueCols As String
Private mstrCalledCols As String


Private mstrCurTechnicRoomName As String            '当前执行间名称
Private mstrCurTechnicDevice As String              '对应当前执行间设备
Private mstrCurTechnicGroupName As String           '当前执行间所属分组
Private mstrTurnPage As String                      '接诊后跳转页面

Private mlngModule As Long
Private mstrPrivs As String

Public Event OnQueueQuick(blnOpenQuick As Boolean)
Public Event OnCallAboutLock(ByVal lngType As Long, ByRef strLockedName As String, ByVal blnLockPara As Boolean)
'104686相关，呼叫后锁定检查，
'lngType类型  1:判断是否启用了参数并且是否已经有被锁定的检查        2:更新参数
'strLockedName   若="" 对流程没有影响，否则说明已经启用参数并且返回之前锁定的检查患者名称
'blnLockPara   用于更新PacsMain中的参数

'重排事件
Public Event OnResotre(ByVal lngAdviceID As Long, ByVal strExeRoom As String)
'完成事件
Public Event OnCompleted(ByVal lngAdviceID As Long, ByVal strExeRoom As String)
'接诊事件
Public Event OnDiagnose(ByVal lngAdviceID As Long, ByVal strExeRoom As String, ByVal strTurnPage As String)
'呼叫事件， 呼叫后不需要更改检查诊室，只有在接诊和完成时才进行更改
Public Event OnCalled(ByVal lngAdviceID As Long, ByVal strRoom As String, ByVal TCallWay As zlQueueOper.TCallWay)

'排队叫号的选择改变事件
Public Event OnSelChange(ByVal lngAdviceID As Long)
'分组内容提示文本事件
Public Event OnGroupHint(ByVal strHint As String)



Public Sub zlInitPacsQueueCfg(ByVal lngModule As Long, _
                            ByVal lngCurDeptId As Long, _
                            ByVal strCurDeptName As String, _
                            ByVal strPrivs As String)
'初始化pacs排队叫号队列配置
    Dim lngCurWorkType As Long
    Dim strQueuePrivs As String
    
    mlngModule = lngModule
    mlngCurDeptId = lngCurDeptId
    mstrCurDeptName = strCurDeptName
    mstrPrivs = strPrivs
    
    strQueuePrivs = ";" & GetPrivFunc(glngSys, 1160) & ";"
    
    lngCurWorkType = IIf(mlngModule = 1290, M_LNG_PACS_BUSINESS_IMG_TYPE, M_LNG_PACS_BUSINESS_CAP_TYPE)
    
    '读取排队叫号参数配置
    Call ReadQueueParameters(lngCurDeptId)
    
    
    ucPacsQueue.ValidDays = mlngValidDays
    ucPacsQueue.ReportNum = mstrReportNum
    ucPacsQueue.GroupField = "队列名称"
    ucPacsQueue.IsReleationQueueTag = True
    
    ucPacsQueue.FindWayEx = M_STR_FINDWAY_EX
    
    '需要使用与业务有关的查询时，需要对DataFields属性进行设置
    ucPacsQueue.DataFields = "ID,业务类型,队列名称,科室ID,病人ID,业务ID,排队序号,排队号码,诊室,患者姓名,性别,年龄,检查项目,医嘱内容,排队状态,排队时间,呼叫医生,呼叫时间,备注"
    ucPacsQueue.DisplayQueueFields = mstrQueueCols '& ",排队序号"
    ucPacsQueue.DisplayCallFields = mstrCalledCols '& ",排队序号"
    
    ucPacsQueue.CalledTarget = mstrCurTechnicRoomName       '设置呼叫所在目的地
    
    
    If mblnUseQueueMsg = True Then
        '启用排队消息处理
        Call ucPacsQueue.UseMsgCenter(glngSys, lngModule)
    End If
    
    Call ucPacsQueue.InitQueue(gcnOracle, _
                                lngCurWorkType, _
                                Me, _
                                App.ProductName, _
                                UserInfo.姓名, _
                                strQueuePrivs)
                                                                
    '清除已经过期的排队叫号业务
    Call ucPacsQueue.QueueOper.CustomClearData("科室ID=" & lngCurDeptId)
    
    '应用呼叫配置，并启动语音呼叫
    Call ucPacsQueue.ApplyVoiceConfig
End Sub


Public Sub zlRefreshQueueData(ByVal strTechnics As String)
'刷新排队数据
    Dim i As Integer
    Dim strTmp As String
    Dim strTechnicGroupNames As String
    
    '配置需要读取的执行间数据（即指定的排队队列数据）
    mstrQueryTechnicQueueNames = ""
    
    If strTechnics <> "" Then
        '0-按默认规则分组，1-按分组设置分组
        If mlngQueueNoWay = 1 Then
            '获取所有选择的执行间对应的分组,问题号:80403
            If UBound(Split(strTechnics, ",")) > 0 Then
                For i = 0 To UBound(Split(strTechnics, ","))
                    strTmp = GetTechnicRoomGrounName(mlngCurDeptId, Split(Split(strTechnics, ",")(i), "-")(1))
                    If strTmp <> "" Then strTechnicGroupNames = strTechnicGroupNames & "," & strTmp
                Next
                
                strTechnicGroupNames = Mid(strTechnicGroupNames, 2)
            Else
                strTmp = GetTechnicRoomGrounName(mlngCurDeptId, Split(strTechnics, "-")(1))
                If strTmp <> "" Then strTechnicGroupNames = strTmp
            End If
            
            mstrQueryTechnicQueueNames = strTechnics & "," & strTechnicGroupNames
        Else
            mstrQueryTechnicQueueNames = strTechnics & "," & mstrCurDeptName & "-" & M_STR_NOT_ALLOT_TECHNIC
        End If
    End If
    
    ucPacsQueue.QueryQueueNames = mstrQueryTechnicQueueNames
    
    If mlngQueueNoWay = 0 Then
        ucPacsQueue.LastFixedQueue = M_STR_NOT_ALLOT_TECHNIC
    Else
        ucPacsQueue.LastFixedQueue = mstrCurTechnicGroupName
    End If
    
    Call ucPacsQueue.RefreshQueueData
End Sub


Private Sub ReadQueueParameters(ByVal lngCurDeptId As Long)
'读取排队叫号参数
    Dim strDeptId As String
    Dim strRoomName As String
    
    '读取当前执行间名称
    strDeptId = Val(zlDatabase.GetPara("本机执行间科室", glngSys, mlngModule, ""))
    strRoomName = zlDatabase.GetPara("本机执行间名称", glngSys, mlngModule, "")
    mstrTurnPage = zlDatabase.GetPara("接诊后跳转页面", glngSys, mlngModule, "")
    mlngValidDays = Val(GetDeptPara(lngCurDeptId, "排队数据保存天数", 1))
    mstrReportNum = GetDeptPara(lngCurDeptId, "排队单报表编号", "")
    mlngPrintWay = Val(GetDeptPara(lngCurDeptId, "排队单打印方式", 0))
    mblnUseQueueMsg = Val(GetDeptPara(lngCurDeptId, "启用排队消息处理", 1))
    
    mstrCurTechnicRoomName = Trim(zlStr.NeedCode(strRoomName))
    mstrCurTechnicDevice = Trim(zlStr.NeedName(strRoomName))
    
    mstrQueueCols = zlDatabase.GetPara("排队队列信息定义", glngSys, mlngModule, "排队号码,患者姓名") 'GetDeptPara(lngCurDeptId, "排队队列信息定义", "")
    mstrCalledCols = zlDatabase.GetPara("呼叫队列信息定义", glngSys, mlngModule, "排队号码,患者姓名") 'GetDeptPara(lngCurDeptId, "呼叫队列信息定义", "")
    
    mlngQueueNoWay = Val(GetDeptPara(lngCurDeptId, "排队叫号编码规则", 0))
    
    mstrCurTechnicGroupName = GetTechnicRoomGrounName(strDeptId, mstrCurTechnicRoomName)   '获取当前执行间分组
End Sub

Private Sub ReadQueueRuleConfig()
'读取排队规则配置
    Dim strSql As String
    
    strSql = "select id,科室ID,组名,分组前缀 from 影像执行分组"
    Set mrsPacsQueueGroupConfig = zlDatabase.OpenSQLRecord(strSql, "查询排队分组信息")
    
    strSql = "select 科室ID,执行间,简码,当前分配,检查设备,号码前缀,分组ID from 医技执行房间"
    Set mrsPacsQueueTechnicConfig = zlDatabase.OpenSQLRecord(strSql, "查询执行间信息")
End Sub


Public Sub zlGetInQueueInf(ByVal lngAdviceID As Long, ByVal lngExecuteDeptId As Long, _
    ByRef strQueueName As String, ByRef strCodeNo As String)
'获取入队相关信息
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strQueueName = ""
    strCodeNo = ""
    
    If mlngQueueNoWay = 0 Then
        '按科室排队
        strSql = "select 名称 from 部门表 where id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询科室名称", lngExecuteDeptId)
        
        If rsData.RecordCount <= 0 Then Exit Sub
                
        strQueueName = Nvl(rsData!名称) & "-" & M_STR_NOT_ALLOT_TECHNIC
        strCodeNo = ""
    Else
        '按分组排队
        strSql = "select a.组名,a.分组前缀,b.名称 from 影像执行分组 a, 部门表 b " & _
                " where a.科室Id=b.Id and a.id=(select a.分组ID " & _
                        " from 影像分组关联 a, 病人医嘱记录 b " & _
                        " where a.诊疗项目id = b.诊疗项目id and a.科室ID=[1] and b.id=[2] and b.相关ID is null)"
                        
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询检查分组", lngExecuteDeptId, lngAdviceID)
        
        If rsData.RecordCount <= 0 Then Exit Sub
        
        strQueueName = Nvl(rsData!名称) & "-" & Nvl(rsData!组名)
        strCodeNo = Nvl(rsData!分组前缀)
    End If
End Sub

Public Function zlGetTechnicRoomCodeNo(ByVal strTechnicRoom As String, ByVal lngDeptID As Long) As String
'查询执行间的排队号码标记
    mrsPacsQueueTechnicConfig.Filter = "执行间='" & strTechnicRoom & "' and 科室ID=" & lngDeptID
    
    zlGetTechnicRoomCodeNo = ""
    
    If mrsPacsQueueTechnicConfig.RecordCount <= 0 Then
        mrsPacsQueueGroupConfig.Filter = ""
        Exit Function
    End If
    
    zlGetTechnicRoomCodeNo = Nvl(mrsPacsQueueTechnicConfig!号码前缀)
    mrsPacsQueueTechnicConfig.Filter = ""
End Function


Private Function GetTechnicRoomGrounName(ByVal lngDeptID As Long, ByVal strTechnicRoom As String) As String
'获取执行间分组名
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetTechnicRoomGrounName = ""
    strSql = "select c.名称,组名 from 影像执行分组 a, 医技执行房间 b, 部门表 c where a.id=b.分组ID  and b.科室ID=c.Id and b.科室Id=[1] and b.执行间=[2]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询医技分组", lngDeptID, strTechnicRoom)
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetTechnicRoomGrounName = Nvl(rsData!名称) & "-" & Nvl(rsData!组名)
End Function

Public Function zlInQueue(ByVal lngAdviceID As Long, _
                                ByVal strName As String, _
                                ByVal lngDeptID As Long, _
                                ByVal strQueueName As String, _
                                ByVal strTarget As String, _
                                ByVal strNoTag As String) As Boolean
        
    Dim rsData As ADODB.Recordset
    Dim lngTimePoint As Long
    Dim lngTimeInterval As Long
On Error GoTo errHandle
    
    zlInQueue = False

    Set rsData = ucPacsQueue.QueueOper.FindQueueInf(lngAdviceID)
    
    If rsData.RecordCount > 0 Then  '存在排队数据
        lngTimePoint = Val(Format(time, "h"))
        If lngTimePoint <= 4 Then
            lngTimeInterval = DateDiff("s", Nvl(rsData!排队时间), Format(zlDatabase.Currentdate - 1, "YYYY-MM-DD 20:00:00"))
        Else
            lngTimeInterval = DateDiff("s", Nvl(rsData!排队时间), Format(zlDatabase.Currentdate, "YYYY-MM-DD 00:00:00"))
        End If
        
        If lngTimeInterval > 0 Then
            '若是今天以前的数据，则直接更新排队
            Call zlUpdatePacsQueue(lngAdviceID, strName, lngDeptID, strQueueName, strTarget, strNoTag)
        Else
            If MsgBoxD(Me, "此病人已在排队叫号队列中，是否重新排队？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                Call zlUpdatePacsQueue(lngAdviceID, strName, lngDeptID, strQueueName, strTarget, strNoTag)
            End If
        End If
    Else
        Call zlInPacsQueue(lngAdviceID, strName, lngDeptID, strQueueName, strTarget, strNoTag)
    End If
    
    zlInQueue = True
Exit Function
errHandle:
    zlInQueue = False
End Function

Public Function zlInPacsQueue(ByVal lngAdviceID As Long, _
                                ByVal strName As String, _
                                ByVal lngDeptID As Long, _
                                ByVal strQueueName As String, _
                                ByVal strTarget As String, _
                                ByVal strNoTag As String) As Boolean
'插入pacs排队队列

On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim strExpandData As String
    Dim strNewQueueNo As String
    
    zlInPacsQueue = False
    
    strExpandData = "科室Id=" & lngDeptID & ",排队标记='" & strNoTag & "'"
    '插入队列数据
    lngQueueId = ucPacsQueue.QueueOper.InsertQueue(strQueueName, , lngAdviceID, strName, strTarget, , strExpandData)
    If lngQueueId <= 0 Then Exit Function
    
    '开始排队
    Call ucPacsQueue.QueueOper.LineQueue(lngQueueId, strNewQueueNo)
    
    '刷新列表数据显示
    Call ucPacsQueue.RefreshQueueRowState(lngQueueId, TQueueState.qsQueueing)
    
    Call AutoPrintQueueInf(lngQueueId)
    
    zlInPacsQueue = True
Exit Function
errHandle:
    zlInPacsQueue = False
End Function

Private Sub AutoPrintQueueInf(ByVal lngQueueId As Long)
'自动打印队列信息
On Error GoTo errHandle
    If mlngPrintWay = 1 Then
        '自动打印
        Call ucPacsQueue.QueueOper.PrintQueueNo(lngQueueId, True, Me)
    ElseIf mlngPrintWay = 2 Then
        '提示打印
        If MsgBoxD(Me, "是否打印当前排号信息？", vbYesNo, gstrSysName) = vbYes Then
            Call ucPacsQueue.QueueOper.PrintQueueNo(lngQueueId, True, Me)
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function zlUpdatePacsQueue(ByVal lngAdviceID As Long, _
                                ByVal strPatientName As String, _
                                ByVal lngDeptID As Long, _
                                Optional ByVal strQueueName As String = "", _
                                Optional ByVal strTarget As String = " ", _
                                Optional ByVal strNoTag As String = " ") As Boolean
'更新排队队列信息
    Dim lngQueueId As Long
    Dim strExpandData As String
    
    zlUpdatePacsQueue = False
    
    If strQueueName = "" Then Exit Function
    
    lngQueueId = ucPacsQueue.QueueOper.FindQueueId(lngAdviceID)

    
    If strPatientName <> "" Then
        Call ucPacsQueue.QueueOper.DeleteQueue(lngQueueId)
        zlUpdatePacsQueue = zlInPacsQueue(lngAdviceID, strPatientName, lngDeptID, strQueueName, strTarget, strNoTag)
    Else
    
        strExpandData = ""
        If strPatientName <> "" Then
            strExpandData = strExpandData & "患者姓名=''" & strPatientName & "''"
        End If
    
        If strTarget <> " " Then
            If strExpandData <> "" Then strExpandData = strExpandData & ","
            strExpandData = strExpandData & "诊室=''" & strTarget & "''"
        End If
        
        If strNoTag <> " " Then
            If strExpandData <> "" Then strExpandData = strExpandData & ","
            strExpandData = strExpandData & "排队标记=''" & strNoTag & "''"
        End If
    
        Call ucPacsQueue.QueueOper.UpdateQueue(lngQueueId, strExpandData)
        Call ucPacsQueue.RefreshQueueData
    
        zlUpdatePacsQueue = True
    End If
End Function


Public Function zlCancelPacsQueue(ByVal lngAdviceID As Long) As Boolean
'撤销pacs排队
    Dim lngQueueId As Long
    
    zlCancelPacsQueue = False
    lngQueueId = ucPacsQueue.QueueOper.FindQueueId(lngAdviceID)
    
    '执行数据删除操作
    Call ucPacsQueue.QueueOper.DeleteQueue(lngQueueId)
    
    zlCancelPacsQueue = True
    
    '刷新列表数据显示
    Call ucPacsQueue.RefreshQueueRowState(lngQueueId, TQueueState.qsAbstain)
End Function


Public Function zlCompletePacsQueue(ByVal lngAdviceID As Long) As Boolean
'完成pacs排队
    Dim lngQueueId As Long
    
    lngQueueId = ucPacsQueue.QueueOper.FindQueueId(lngAdviceID)
    
    '执行完成排队操作
    zlCompletePacsQueue = ucPacsQueue.QueueOper.CompleteQueue(lngQueueId)
    
    '刷新列表数据显示
    Call ucPacsQueue.RefreshQueueRowState(lngQueueId, TQueueState.qsComplete)
End Function

Public Sub zlExecuteCommandbar(control As CommandBarControl)
'执行菜单事件
    Call ucPacsQueue.zlExecuteCommandBars(control)
End Sub


Private Sub Form_Load()
'    'Debug Code...
'        Call InitDebugObject(1290, Me, "zlhis", "HIS")
'        Call InitPacsQueueCfg("测试队列,050204-超声波室", "超声波室", "排队号码,患者姓名,医嘱内容", "排队号码,患者姓名")
'    'Debug End

    Call ReadQueueRuleConfig
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ucPacsQueue.Left = 0
    ucPacsQueue.Top = 0
    ucPacsQueue.Width = Me.ScaleWidth
    ucPacsQueue.Height = Me.ScaleHeight
    
    If Me.ScaleWidth < 12900 Then
        ucPacsQueue.IsIconLarge = False
        ucPacsQueue.IsShowToolText = IIf(Me.ScaleWidth < 8000, False, True)
    Else
        ucPacsQueue.IsIconLarge = True
        ucPacsQueue.IsShowToolText = True
    End If
err.Clear
End Sub

Private Function GetRoom(ByVal lngQueueId As Long) As String
'获取排队叫号对应的诊室（执行间）
    GetRoom = Nvl(ucPacsQueue.QueueOper.GetQueueInf(lngQueueId, "诊室")!诊室)
End Function


Private Function GetAdviceId(ByVal lngQueueId As Long) As Long
'获取排队叫号对应的医嘱ID
    GetAdviceId = Val(Nvl(ucPacsQueue.QueueOper.GetQueueInf(lngQueueId, "业务ID")!业务ID))
End Function


Private Sub ucPacsQueue_OnCallPreAfter(ByVal lngQueueId As Long, ByVal lngCallWay As zlQueueOper.TCallWay)
'呼叫后触发事件
    Dim lngAdviceID As Long
    Dim strRoom As String
    
    lngAdviceID = GetAdviceId(lngQueueId)
    strRoom = GetRoom(lngQueueId)
    
    RaiseEvent OnCalled(lngAdviceID, strRoom, lngCallWay)

End Sub

Private Sub UcPacsQueue_OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As zlQueueOper.TCallWay, strCallContext As String, blnCancel As Boolean)
'Pacs排队叫号呼叫事件处理
    Dim strOldTechnicRoomName As String
    Dim lngResult As Long
    Dim lngRowIndex As Long
    Dim strSql As String
    Dim lngAdviceID As Long
    Dim strName As String
    Dim blnTmp As Boolean
        
    '对于未分配执行间的检查，则需要插入当前所在的执行间到排队叫号的诊室中
    If lngCallWay = cwOrder Or lngCallWay = cwSpecify Or lngCallWay = cwWaitRoom Then
        '已经有锁定的检查，解锁然后锁定成现在的
        RaiseEvent OnCallAboutLock(1, strName, blnTmp)
                
        '判断当前队列是否已经分执行间，如果已经分配但与当前执行间不同，则需要进行提醒
        strOldTechnicRoomName = Trim(Nvl(ucPacsQueue.QueueOper.GetQueueInf(lngQueueId, "诊室")!诊室))
        
        If strOldTechnicRoomName <> "" And strOldTechnicRoomName <> mstrCurTechnicRoomName Then
            lngResult = MsgBoxD(Me, "当前检查已被分配到 【" & strOldTechnicRoomName & "】 执行间，是否需要更改到本执行间执行？" & vbCrLf & _
                                    "选择“是”表示更改到本执行间后呼叫；" & vbCrLf & _
                                    "选择“否”表示不更改执行间直接呼叫；" & vbCrLf & _
                                    "选择“取消”表示不进行呼叫；", vbYesNoCancel, "提示")
            
            If lngResult = vbCancel Then
                blnCancel = True
                Exit Sub
            End If
        End If
          
        '调整呼叫目的地
        If lngResult = vbYes Or strOldTechnicRoomName = "" Then
            Call ucPacsQueue.QueueOper.WriteTarget(lngQueueId)
            '需要同步更新医嘱发送的执行间
            
            lngAdviceID = GetAdviceId(lngQueueId)
            
            strSql = "zl_影像检查_更新执行间(" & lngAdviceID & ",'" & mstrCurTechnicRoomName & "','" & mstrCurTechnicDevice & "')"
            
            Call zlDatabase.ExecuteProcedure(strSql, "更新检查执行间")
        
            '更新排队列表上的诊室显示
            lngRowIndex = ucPacsQueue.GetRowIndex(qftWaitQueue, "ID", lngQueueId)
            If lngRowIndex >= 0 Then
                Call ucPacsQueue.SetListValue(qftWaitQueue, lngRowIndex, "诊室", mstrCurTechnicRoomName)
                Call ucPacsQueue.Populate(qftWaitQueue)
            End If
        End If
    End If
End Sub

Private Sub UcPacsQueue_OnCmdBarUpdate(objComandBarControl As Object)
'屏蔽接诊按钮
'    If objComandBarControl.ID = TMenuId.mi接诊 Then
'        objComandBarControl.Visible = False
'    End If
End Sub


Private Sub ucPacsQueue_OnConfigEvent(blnUseCustom As Boolean)
'Pacs排队叫号配置事件
On Error GoTo errHandle
    Dim objCfgWindow As frmWork_QueueCfg
    Dim blnLock As Boolean
    Dim blnQuick As Boolean
    Dim strTmp As String
    
    blnUseCustom = True
    
    Set objCfgWindow = New frmWork_QueueCfg
    
    If objCfgWindow.ShowQueueConfig(ucPacsQueue, mlngModule, mstrPrivs, Me, blnLock, blnQuick) Then
        '重新读取和应用配置
        RaiseEvent OnCallAboutLock(2, strTmp, blnLock)
        RaiseEvent OnQueueQuick(blnQuick)
        Call zlInitPacsQueueCfg(mlngModule, mlngCurDeptId, mstrCurDeptName, mstrPrivs)
        Call ucPacsQueue.RefreshQueueData
    End If
    
Exit Sub
errHandle:
    Unload objCfgWindow
    Set objCfgWindow = Nothing
    
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucPacsQueue_OnFindData(ByVal strFindWay As String, ByVal strFindValue As String, txtFind As Object, rsData As ADODB.Recordset, blnUseCustom As Boolean)
'自定义查找
    Dim strSql As String
    Dim strCurQueryQueueNames As String
    Dim strQueryCols As String
    Dim blnQueryProject As Boolean
    Dim strTempCols As String '这些列的数据类型为number类型，因此需用val处理
    
    blnUseCustom = True
    strCurQueryQueueNames = Replace(mstrQueryTechnicQueueNames, ",", "','")
    
    If strFindWay = "排队号" Or strFindWay = "排队号码" Then strFindWay = "排队标记||a.排队号码"
    If strFindWay = "姓名" Then strFindWay = "患者姓名"
    strTempCols = "门诊号,住院号,ID,业务类型,科室ID,病人ID,业务ID,排队状态"
        
    '"ID,业务类型,队列名称,科室ID,病人ID,业务ID,排队序号,排队号码,诊室,患者姓名,性别,年龄,检查项目,医嘱内容,排队状态,排队时间,呼叫医生,呼叫时间,备注"
    
    '获取需要从数据库中查询的字段
    strQueryCols = ucPacsQueue.GetValidCols("a.ID,a.业务类型,a.队列名称,a.科室ID,a.病人ID,a.业务ID,a.排队号码,a.排队标记,a.排队序号,a.诊室," & _
                                            "a.患者姓名,b.性别,b.年龄,c.名称 as 检查项目,b.医嘱内容,a.排队状态," & _
                                            "a.排队时间,a.呼叫医生,a.呼叫时间,a.备注", "a")
    
    blnQueryProject = IIf(InStr(strQueryCols, "检查项目") > 0, True, False)
    
    strSql = "select " & strQueryCols & _
            " from 排队叫号队列 a, 病人医嘱记录 b,病人信息 d " & IIf(blnQueryProject, ", 诊疗项目目录 c ", "") & _
            " where a.业务ID=b.Id and b.病人ID=d.病人ID " & _
                    IIf(blnQueryProject, " and b.诊疗项目ID=c.ID and c.类别='D'", "") & _
            "       and b.相关ID is null and a.业务类型=1 " & _
            "       and a.科室ID=[1] " & IIf(strCurQueryQueueNames = "", "", "and 队列名称 in ('" & strCurQueryQueueNames & "') ") & _
            IIf(InStr(M_STR_FINDWAY_EX, strFindWay) > 0, " and upper(d.", " and upper(a.") & IIf(strFindWay = "就诊卡", "就诊卡号", strFindWay) & ")=upper([2]) " & _
            IIf(ucPacsQueue.QueueOper.CustomOrder = "", "", " order by " & ucPacsQueue.QueueOper.CustomOrder)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询PACS排队队列", mlngCurDeptId, IIf(InStr(strTempCols, strFindWay) > 0, Val(strFindValue), Trim(strFindValue)))
End Sub

Private Sub ucPacsQueue_OnLocateData(ByVal strLocateWay As String, ByVal strLocateValue As String, txtFind As Object, lngQueueId As Long, blnUseCustom As Boolean)
'排队数据定位事件
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strTempCols As String '这些列的数据类型为number类型，因此需用val处理
    
    blnUseCustom = True
    
    If strLocateWay = "排队号" Or strLocateWay = "排队号码" Then strLocateWay = "排队标记||a.排队号码"
    If strLocateWay = "姓名" Then strLocateWay = "患者姓名"
    strTempCols = "门诊号,住院号,ID,业务类型,科室ID,病人ID,业务ID,排队状态"
    
    strSql = "select a.ID from 排队叫号队列 a, 病人医嘱记录 b, 病人信息 d" & _
            " where a.业务ID=b.ID and b.病人ID=d.病人ID and b.相关ID is null and upper(" & _
            IIf(InStr(M_STR_FINDWAY_EX, strLocateWay) > 0, " d.", " a.") & IIf(strLocateWay = "就诊卡", "就诊卡号", strLocateWay) & ")=upper([1])"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "定位排队数据", IIf(InStr(strTempCols, strLocateWay) > 0, Val(strLocateValue), Trim(strLocateValue)))
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    lngQueueId = Val(Nvl(rsData!ID))
End Sub

Private Sub ucPacsQueue_OnGroupHint(ByVal strHintContext As String)
On Error Resume Next
    RaiseEvent OnGroupHint(strHintContext)
err.Clear
End Sub

Private Sub ucPacsQueue_OnModifyBefore(ByVal lngListType As zlQueueOper.TQueueFromType, ByVal lngQueueId As Long, objInputCfg As Dictionary, blnCancel As Boolean, blnUseCustom As Boolean)
    '查询当前科室的诊室信息
    Dim strRooms As String
    
    If mrsPacsQueueTechnicConfig Is Nothing Then Exit Sub
    
    mrsPacsQueueTechnicConfig.Filter = "科室ID=" & mlngCurDeptId
    If mrsPacsQueueTechnicConfig.RecordCount <= 0 Then Exit Sub
    
    While Not mrsPacsQueueTechnicConfig.EOF
        If strRooms <> "" Then strRooms = strRooms & ","
        strRooms = strRooms & Nvl(mrsPacsQueueTechnicConfig!执行间)
        
        Call mrsPacsQueueTechnicConfig.MoveNext
    Wend
    
    mrsPacsQueueTechnicConfig.Filter = ""
    
    objInputCfg.Item("诊室") = objInputCfg.Item("诊室") & ":" & strRooms
err.Clear
End Sub

Private Sub ucPacsQueue_OnModifyAfter(ByVal lngQueueId As Long, objUpdateValue As Dictionary)
    objUpdateValue.Item("排队号码") = objUpdateValue.Item("排队标记") & objUpdateValue.Item("排队号码")
End Sub

Private Sub UcPacsQueue_OnQueryQueueData(rsData As ADODB.Recordset, blnUseCustom As Boolean)
'查询pacs排队队列数据
'由于涉及到查询pacs检查相关的数据信息，因此需要使用该事件进行自定义查询
'查询当天的排队情况
    Dim strSql As String
    Dim strCurQueryQueueNames As String
    Dim lngTimePoint As Long
    Dim strStartTime As String
    Dim strEndTime As String
    Dim strQueryCols As String
    Dim blnQueryProject As Boolean
    Dim dtNow As Date
    blnUseCustom = True
    
    strCurQueryQueueNames = Replace(mstrQueryTechnicQueueNames, ",", "','")
    dtNow = zlDatabase.Currentdate
    
    lngTimePoint = Val(Format(time, "h"))
    If lngTimePoint <= 4 Then
        strStartTime = zlStr.To_Date(Format(dtNow - 1, "yy-mm-dd 20:00:00"))
        strEndTime = zlStr.To_Date(Format(dtNow, "yy-mm-dd 08:00:00"))
    Else
        strStartTime = zlStr.To_Date(Format(dtNow, "yy-mm-dd 00:00:00"))
        strEndTime = zlStr.To_Date(Format(dtNow, "yy-mm-dd 23:59:59"))
    End If
    
    '"ID,业务类型,队列名称,科室ID,病人ID,业务ID,排队序号,排队号码,诊室,患者姓名,性别,年龄,检查项目,医嘱内容,排队状态,排队时间,呼叫医生,呼叫时间,备注"
    
    '获取需要从数据库中查询的字段
    strQueryCols = ucPacsQueue.GetValidCols("a.ID,a.业务类型,a.队列名称,a.科室ID,a.病人ID,a.业务ID,a.排队标记,a.排队号码,a.排队序号,a.诊室," & _
                                            "a.患者姓名,b.性别,b.年龄,c.名称 as 检查项目,b.医嘱内容,a.排队状态," & _
                                            "a.排队时间,a.呼叫医生,a.呼叫时间,a.备注", "a")
    
    'strQueryCols = Replace(strQueryCols, "A.排队号码", "A.排队标记 || A.排队号码 as 排队号码")
    
    blnQueryProject = IIf(InStr(strQueryCols, "检查项目") > 0, True, False)
    
    strSql = "select " & strQueryCols & _
            " from 排队叫号队列 a, 病人医嘱记录 b" & IIf(blnQueryProject, ", 诊疗项目目录 c ", "") & _
            " where a.业务ID=b.Id " & _
                    IIf(blnQueryProject, " and b.诊疗项目ID=c.ID and c.类别='D'", "") & " and b.相关ID is null and a.业务类型=1 and a.排队时间 between " & strStartTime & " and " & strEndTime & _
            "       and a.科室ID=[1] " & IIf(strCurQueryQueueNames = "", "", "and 队列名称 in ('" & strCurQueryQueueNames & "') ") & IIf(ucPacsQueue.QueueOper.CustomOrder = "", "", " order by " & ucPacsQueue.QueueOper.CustomOrder)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询PACS排队队列", mlngCurDeptId)
End Sub

Private Sub UcPacsQueue_OnSelectionChanged(ByVal lngListType As zlQueueOper.TQueueFromType, ByVal lngQueueId As Long, objQueueList As Object, objReportRow As Object)
'排队叫号选择行改变事件
    Dim lngAdviceID As Long
    Dim lngColIndex As Long
    
    If objReportRow Is Nothing Then Exit Sub
    If objReportRow.Record Is Nothing Then Exit Sub
    
    lngColIndex = ucPacsQueue.GetColumnIndex(lngListType, "业务ID")
    
    lngAdviceID = Val(objReportRow.Record(lngColIndex).value)
    
    RaiseEvent OnSelChange(lngAdviceID)
End Sub

Private Sub ucPacsQueue_OnWorkAfter(ByVal lngQueueId As Long, ByVal strCurQueueName As String, ByVal lngOperationType As zlQueueOper.TOperationType)
'如果进行接诊操作，则需要更新检查的“执行间”数据
    Dim lngAdviceID As Long
    Dim strSql As String
    Dim strRoom As String
    Dim lngRowIndex As Long
    Dim strCodeTag As String
    Dim rsData As ADODB.Recordset
    
    If lngOperationType = otComplete Then
        lngAdviceID = GetAdviceId(lngQueueId)
        
        '完成时，需要更新最终的执行间
        strSql = "zl_影像检查_更新执行间(" & lngAdviceID & ",'" & mstrCurTechnicRoomName & "','" & mstrCurTechnicDevice & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        RaiseEvent OnCompleted(lngAdviceID, mstrCurTechnicRoomName)
        
    ElseIf lngOperationType = otDiagnose Then
        lngAdviceID = GetAdviceId(lngQueueId)
        
        '接诊时，更新当前所在执行间
        strSql = "zl_影像检查_更新执行间(" & lngAdviceID & ",'" & mstrCurTechnicRoomName & "','" & mstrCurTechnicDevice & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                        
        RaiseEvent OnDiagnose(lngAdviceID, mstrCurTechnicRoomName, mstrTurnPage)
    ElseIf lngOperationType = otRestore Then
        strRoom = ""
        
        lngRowIndex = ucPacsQueue.GetRowIndex(qftWaitQueue, "ID", lngQueueId)
            
        '如果是重排，需要判断是否调整了排队队列，如果队列进行了调整，则需要对诊室或执行间进行对应的更新
        If strCurQueueName <> mstrCurTechnicGroupName And strCurQueueName <> mstrCurDeptName & "-" & M_STR_NOT_ALLOT_TECHNIC Then
            '获取当前队列对应的诊室名称
            strRoom = Replace(strCurQueueName, mstrCurDeptName & "-", "")
            
            '更新排队诊室
            Call ucPacsQueue.QueueOper.WriteTarget(lngQueueId, strRoom)
        End If
        
        '更新医嘱执行间
        lngAdviceID = GetAdviceId(lngQueueId)
        
        strSql = "zl_影像检查_更新执行间(" & lngAdviceID & ",'" & strRoom & "','" & mstrCurTechnicDevice & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "更新检查执行间")

        '刷新界面排队号码及排队诊室的显示
        If lngRowIndex >= 0 Then
            Call ucPacsQueue.SetListValue(qftWaitQueue, lngRowIndex, "诊室", strRoom)
            Call ucPacsQueue.Populate(qftWaitQueue)
        End If
        
        RaiseEvent OnResotre(lngAdviceID, strRoom)
    End If
End Sub


Private Sub ucPacsQueue_OnCreateQueueNo(ByVal lngQueueId As Long, ByVal strQueueName As String, strQueueNo As String)
'排队号码生成事件
    Dim strRoom As String
    Dim strCodeTag As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If strQueueNo = "" Then Exit Sub
    
    strCodeTag = ""
    If strQueueName <> mstrCurTechnicGroupName And strQueueName <> mstrCurDeptName & "-" & M_STR_NOT_ALLOT_TECHNIC Then
        '获取执行间前缀
        strRoom = Replace(strQueueName, mstrCurDeptName & "-", "")
        strCodeTag = zlGetTechnicRoomCodeNo(strRoom, mlngCurDeptId)
    Else
        
        '如果是按科室排队，则没有排队标记
        If mlngQueueNoWay = 1 Then
            '获取分组的排队标记
            strSql = "select a.组名,a.分组前缀 from 影像执行分组 a " & _
                    " where a.科室ID=[1] and a.组名=[2]"
            Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询重排分组前缀", mlngCurDeptId, Replace(mstrCurTechnicGroupName, mstrCurDeptName & "-", ""))
                    
            If rsData.RecordCount > 0 Then
                strCodeTag = Nvl(rsData!分组前缀)
            End If
        End If
    End If
    
    '需要更新排队标记
    Call ucPacsQueue.QueueOper.UpdateQueue(lngQueueId, "排队标记=''" & strCodeTag & "''")
        
    strQueueNo = strCodeTag & strQueueNo
End Sub

Public Sub CloseQueueQuick()
    If Not ucPacsQueue Is Nothing Then
        ucPacsQueue.CloseQueueQuick
    End If
End Sub

Public Sub OpenQueueQuick(ByVal strTechnics As String, objOwer As Object)
    Call zlRefreshQueueData(strTechnics)
    
    If Not ucPacsQueue Is Nothing Then
        ucPacsQueue.OpenQueueQuick objOwer
    End If
End Sub
