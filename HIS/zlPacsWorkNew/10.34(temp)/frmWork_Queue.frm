VERSION 5.00
Object = "{6856A5DD-B624-47EE-85F4-F9812BFD363A}#1.0#0"; "UcQueueManage.ocx"
Begin VB.Form frmWork_Queue 
   BorderStyle     =   0  'None
   Caption         =   "排队叫号管理"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   Icon            =   "frmWork_Queue.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin UcQueueManage.UcQueue ucPacsQueue 
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


Private Const M_LNG_PACS_BUSINESS_TYPE As Long = 1                  'pacs业务类型定义
Private Const M_STR_NOT_ALLOT_TECHNIC As String = "未分配队列"      '未分配队列名称定义


Private mrsPacsQueueGroupConfig As ADODB.Recordset
Private mrsPacsQueueTechnicConfig As ADODB.Recordset


Private mlngCurDeptId As Long                       '当前科室ID
Private mstrQueryTechnicQueueNames  As String       'pacs排队叫号查询队列名称

Private mstrQueueCols As String
Private mstrCalledCols As String

Private mstrCurTechnicRoomName As String            '当前执行间名称
Private mstrCurTechnicGroupName As String           '当前执行间所属分组

Private mlngModule As Long


'完成事件
Public Event OnCompleted(ByVal lngAdviceID As Long)
'接诊事件
Public Event OnDiagnose(ByVal lngAdviceID As Long)
'呼叫事件
Public Event OnCalled(ByVal lngAdviceID As Long)

'排队叫号的选择改变事件
Public Event OnSelChange(ByVal lngAdviceID As Long)




Property Get Queue() As clsQueueOperation
'队列操作对象
    Set Queue = ucPacsQueue.QueueOper
End Property






Public Sub zlInitPacsQueueCfg(ByVal lngModule As Long, ByVal lngCurDeptId As Long)
'初始化pacs排队叫号队列配置

    
    mlngModule = lngModule
    mlngCurDeptId = lngCurDeptId

    '读取排队叫号参数配置
    Call ReadQueueParameters(lngCurDeptId)
    
    
    ucPacsQueue.GroupField = "队列名称"
    
    ucPacsQueue.FindWayEx = "门诊号,住院号,就诊号,医保号"
    ucPacsQueue.DisplayQueueFields = mstrQueueCols
    ucPacsQueue.DisplayCallFields = mstrCalledCols
    
    
    Call ucPacsQueue.InitQueue(gcnOracle, _
                                M_LNG_PACS_BUSINESS_TYPE, _
                                Me, _
                                App.ProductName, _
                                ",打号,顺呼,直呼,广播,优先,插队,重排,接诊,暂停,弃号,恢复,完成,刷新,查找,修改,设置,")
                                                                
    
End Sub


Public Sub zlRefreshQueueData(ByVal strTechnics As String)
'刷新排队数据
    
    '配置需要读取的执行间数据（即指定的排队队列数据）
    mstrQueryTechnicQueueNames = strTechnics & "," & mstrCurTechnicGroupName & "," & M_STR_NOT_ALLOT_TECHNIC
    
    ucPacsQueue.QueryQueueNames = mstrQueryTechnicQueueNames
    
    Call ucPacsQueue.RefreshQueueData
End Sub


Private Sub ReadQueueParameters(ByVal lngCurDeptId As Long)
'读取排队叫号参数
    '读取当前执行间名称
    mstrCurTechnicRoomName = zlDatabase.GetPara("本机执行间名称", glngSys, mlngModule, "")
    
    mstrQueueCols = GetDeptPara(lngCurDeptId, "排队队列信息定义", "")
    mstrCalledCols = GetDeptPara(lngCurDeptId, "呼叫队列信息定义", "")
    
    mstrCurTechnicGroupName = GetTechnicRoomGrounName(NeedNo(mstrCurTechnicRoomName))   '获取当前执行间分组
End Sub


Private Sub ReadQueueRuleConfig()
'读取排队规则配置
    Dim strSql As String
    
    strSql = "select id,科室ID,组名,分组前缀,当前序号 from 影像执行分组"
    Set mrsPacsQueueGroupConfig = zlDatabase.OpenSQLRecord(strSql, "查询排队分组信息")
    
    strSql = "select 科室ID,执行间,简码,当前分配,检查设备,号码前缀,分组ID,当前序号 from 医技执行房间"
    Set mrsPacsQueueTechnicConfig = zlDatabase.OpenSQLRecord(strSql, "查询执行间信息")
End Sub


Public Function zlGetStudyGroupName(ByVal lngAdviceID As Long) As String
'获取检查项目分组名称
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    zlGetStudyGroupName = ""
    strSql = "select 组名 from 影像执行分组 " & _
            " where id=(select 分组ID " & _
                    " from 影像检查项目 a, 病人医嘱记录 b " & _
                    " where a.诊疗项目id = b.诊疗项目id and b.id=[1] and b.相关ID is null)"
    
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询检查分组", lngAdviceID)
    If rsData.RecordCount <= 0 Then Exit Function
    
    zlGetStudyGroupName = Nvl(rsData!组名)
End Function

Public Function zlGetGroupCodeNo(ByVal strGroupName As String) As String
'查询分组的排队号码标记
    mrsPacsQueueGroupConfig.Filter = "组名='" & strGroupName & "'"
    
    zlGetGroupCodeNo = ""
    
    If mrsPacsQueueGroupConfig.RecordCount <= 0 Then
        mrsPacsQueueGroupConfig.Filter = ""
        Exit Function
    End If
    
    zlGetGroupCodeNo = Nvl(mrsPacsQueueGroupConfig!分组前缀)
    mrsPacsQueueGroupConfig.Filter = ""
End Function

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


Private Function GetTechnicRoomGrounName(ByVal strTechnicRoom As String) As String
'获取执行间分组名
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetTechnicRoomGrounName = ""
    strSql = "select 组名 from 影像执行分组 a, 医技执行房间 b where a.id=b.分组ID and b.科室Id=[1] and b.执行间=[2]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询医技分组", mlngCurDeptId, strTechnicRoom)
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetTechnicRoomGrounName = Nvl(rsData!组名)
End Function


Public Function zlInPacsQueue(ByVal lngAdviceID As Long, _
                                ByVal strName As String, _
                                ByVal strQueueName As String, _
                                ByVal strTarget As String, _
                                ByVal strNoTag As String) As Boolean
'插入pacs排队队列
On Error GoTo ErrHandle
    Dim lngQueueId As Long
    
    zlInPacsQueue = False
    
    '插入队列数据
    lngQueueId = ucPacsQueue.QueueOper.InsertQueue(strQueueName, , lngAdviceID, strName, strTarget, , "排队标记='" & strNoTag & "'")
    If lngQueueId <= 0 Then Exit Function
    
    '开始排队
    Call ucPacsQueue.QueueOper.StartQueue(lngQueueId)
    
    Call ucPacsQueue.RefreshQueueData
    
    zlInPacsQueue = True
Exit Function
ErrHandle:
    zlInPacsQueue = False
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
err.Clear
End Sub

Private Function GetQueryCol(ByVal strCols As String) As String
'必须具备的查询列字段：ID,队列名称,业务ID,患者姓名,排队状态,排队号码,排队时间,业务类型,排队序号
    Dim strResult As String
    
    strResult = UCase(strCols)
    
    strResult = Replace(strResult, "ID,", "")
    strResult = Replace(strResult, "队列名称,", "")
    strResult = Replace(strResult, "业务ID,", "")
    strResult = Replace(strResult, "患者姓名,", "")
    strResult = Replace(strResult, "排队状态,", "")
    strResult = Replace(strResult, "排队号码,", "")
    strResult = Replace(strResult, "排队时间,", "")
    strResult = Replace(strResult, "业务类型,", "")
    strResult = Replace(strResult, "排队序号,", "")
    
    strResult = "a.ID,a.队列名称,a.业务ID,a.患者姓名,a.排队状态, a.排队标记 || a.排队号码 as 排队号码,a.排队时间,a.业务类型,a.排队序号 " & IIf(strResult = "", "", "," & strResult)
    
    GetQueryCol = strResult
End Function


Private Function GetAdviceId(ByVal lngQueueId As Long) As Long
'获取排队叫号对应的医嘱ID
    GetAdviceId = Val(Nvl(ucPacsQueue.QueueOper.GetQueueInf(lngQueueId, "业务ID")!业务ID))
End Function


Private Sub UcPacsQueue_OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As UcQueueManage.TCallWay, strCallContext As String, blnCancel As Boolean)
'Pacs排队叫号呼叫事件处理
    Dim lngAdviceID As Long
    
    
    '对于未分配执行间的检查，则需要插入当前所在的执行间到排队叫号的诊室中
    If lngCallWay = cwOrder Or lngCallWay = cwSpecify Or lngCallWay = cwWaitRoom Then
        lngAdviceID = GetAdviceId(lngQueueId)
        
        Call ucPacsQueue.QueueOper.WriteTarget(lngQueueId, mstrCurTechnicRoomName)
        RaiseEvent OnCalled(lngAdviceID)
    End If
End Sub

Private Sub UcPacsQueue_OnCmdBarUpdate(objComandBarControl As Object)
'屏蔽接诊按钮
'    If objComandBarControl.ID = TMenuId.mi接诊 Then
'        objComandBarControl.Visible = False
'    End If
End Sub

Private Sub UcPacsQueue_OnQueryCallData(rsData As ADODB.Recordset, blnUseCustom As Boolean)
'查询pacs排队已呼叫数据
'由于涉及到查询pacs检查相关的数据信息，因此需要使用该事件进行自定义查询
    Dim strSql As String
    Dim strCurQueryQueueNames As String
    
    blnUseCustom = True
    
    strCurQueryQueueNames = Replace(mstrQueryTechnicQueueNames, ",", "','")
    
    strSql = "select " & GetQueryCol(mstrCalledCols) & " from 排队叫号队列 a, 病人医嘱记录 b where a.业务ID=b.Id and b.相关ID is null and a.业务类型=1 and a.排队状态 in(1,7) " & IIf(strCurQueryQueueNames = "", "", "and 队列名称 in ('" & strCurQueryQueueNames & "') ")
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询PACS已呼叫队列")
End Sub

Private Sub UcPacsQueue_OnQueryQueueData(ByVal lngQueueLoadModle As UcQueueManage.TQueueSelState, rsData As ADODB.Recordset, blnUseCustom As Boolean)
'查询pacs排队队列数据
'由于涉及到查询pacs检查相关的数据信息，因此需要使用该事件进行自定义查询
    Dim strSql As String
    Dim strCurQueryQueueNames As String
    
    blnUseCustom = True
    
    strCurQueryQueueNames = Replace(mstrQueryTechnicQueueNames, ",", "','")
    
    strSql = "select " & GetQueryCol(mstrQueueCols) & " from 排队叫号队列 a, 病人医嘱记录 b where a.业务ID=b.Id  and b.相关ID is null and a.业务类型=1 and a.排队状态=[1] " & IIf(strCurQueryQueueNames = "", "", "and 队列名称 in ('" & strCurQueryQueueNames & "') ")
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询PACS排队队列", lngQueueLoadModle)
End Sub

Private Sub UcPacsQueue_OnSelectionChanged(ByVal lngListType As UcQueueManage.TQueueFromType, ByVal lngQueueId As Long, objQueueList As Object, objReportRow As Object)
'排队叫号选择行改变事件
    Dim lngAdviceID As Long
    Dim lngColIndex As Long
    
    If objReportRow Is Nothing Then Exit Sub
    If objReportRow.Record Is Nothing Then Exit Sub
    
    lngColIndex = ucPacsQueue.GetColumnIndex(lngListType, "业务ID")
    
    lngAdviceID = Val(objReportRow.Record(lngColIndex).value)
    
    RaiseEvent OnSelChange(lngAdviceID)
End Sub

Private Sub UcPacsQueue_OnWorkBefore(ByVal lngQueueId As Long, ByVal lngOperationType As UcQueueManage.TOperationType, blnCancel As Boolean)
'如果进行接诊操作，则需要更新检查的“执行间”数据
    Dim lngAdviceID As Long
    
    '接诊操作需要触发接诊事件
    If lngOperationType = otComplete Then
        lngAdviceID = GetAdviceId(lngQueueId)
        RaiseEvent OnCompleted(lngAdviceID)
        
    ElseIf lngOperationType = otDiagnose Then
        lngAdviceID = GetAdviceId(lngQueueId)
        RaiseEvent OnDiagnose(lngAdviceID)
        
    End If
End Sub
