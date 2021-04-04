VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mintValidDays           As Integer '数据有效天数
Private mintCallWay             As Integer '呼叫方式 0-系统默认，1-微软语言
Private mintCallCount           As Integer '呼叫次数 默认为1
Private mintCallSpeed           As Integer '呼叫速度
Private mintCallLen             As Integer '呼叫时间长度，默认10秒
Private mstrCallStation         As String  '待呼叫站点名,默认当前站点
Private mstrCallTarget          As String  '呼叫前设置病人的检查目的地
Private mstrLocalStation        As String  '本机站点名
Private mintWorkType            As String  '业务类型
Private mcnOracle               As ADODB.Connection
Private mlngSysNo               As Long
Private mstrCustomOrder         As String  '自定义排序
Private mstrDefQueryCols        As String  '默认查询列，默认为排队叫号队列表中的所有列
Private mstrCallScript          As String
Private mblnUseVbsCall          As Boolean

Public Event OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As Long, ByRef strCallContext As String, blnCancel As Boolean)
Public Event OnCallPreAfter(ByVal lngQueueId As Long, ByVal lngCallWay As Long)
Public Event OnWorkBefore(ByVal lngQueueId As Long, ByVal lngOperationType As Long, blnCancel As Boolean)
Public Event OnWorkAfter(ByVal lngQueueId As Long, ByVal lngOperationType As Long)
Public Event OnCallingBefore(ByVal lngCallId As Long, ByVal strCallContext As String, blnCancel As Boolean)
Public Event OnCallingAfter(ByVal lngCallId As Long, ByVal strCallContext As String)

'语音播放的函数
Private Declare Function StartTextPlay Lib "StrSound.dll" (ByVal PlayText As String, ByVal intxx As Integer) As Long
Private Declare Function StopPlayStr Lib "StrSound" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long



'操作类型
Public Enum OperationType
    otStart = 0       '排队
    otPause = 1       '暂停
    otAbstain = 2     '弃号
    otRestore = 3     '重排
    otComplete = 4    '完成
    otPriority = 5    '优先
    otInsertQueue = 6 '插队
    otDiagnose = 7   '接诊
End Enum

Public Sub InitQueue(cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal intWorkType As Integer)
'功能：初始化队列
    Set mcnOracle = cnOracle
    
    mlngSysNo = lngSysNo
    mintWorkType = intWorkType
    
    mintValidDays = 0
    mintCallWay = 0
    mintCallCount = 1
    mintCallSpeed = 10
    mintCallLen = 10
    mstrCallScript = ""
    mblnUseVbsCall = False
    
    mstrCustomOrder = GetCustomOrderWhereStr
    mstrLocalStation = AnalyseComputer
    mstrCallStation = AnalyseComputer
    mstrDefQueryCols = GetAllCols

End Sub

Public Function InsertQueue(ByVal strQueueName As String, _
                            Optional ByVal strQueueNo As String, _
                            Optional ByVal strBusinessId As String, _
                            Optional ByVal strName As String, _
                            Optional ByVal strTarget As String, _
                            Optional ByVal strCustomOrder As String, _
                            Optional ByVal strExpandData As String) As Long
'strQueueName: 要插入的队列名称
'strQueueNo: 排队号码
'strBusinessId: 业务ID
'strName: 姓名
'strTarget: 目的地
'strCustomOrder: 自定义排队序号
'strExpandData:扩展数据 格式为/* "姓名='张四', 性别='女',年龄=10"*/
'功能：返回插入后的排队ID，失败返回-1
    Dim strSql          As String
    Dim rsQueue         As ADODB.Recordset
    
    On Error GoTo ErrorHand
    
    InsertQueue = -1
    
    strSql = "select zl_排队叫号队列_插入([1],[2],[3],[4],[5],[6],[7],[8]) as 排队id from dual"
    
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "返回插入后的排队ID", strQueueName, mintWorkType, _
                                        strQueueNo, strBusinessId, strName, strTarget, strCustomOrder, strExpandData)
                                        
    If rsQueue.RecordCount > 0 Then InsertQueue = Nvl(rsQueue!排队id)
    
    Exit Function
ErrorHand:
    InsertQueue = -1
End Function

Public Sub DeleteQueue(ByVal lngQueueId As Long)
'lngQueueID: 排队ID
'功能：删除队列数据
    Dim strSql          As String
    
    strSql = "ZL_排队叫号队列_删除(" & lngQueueId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "删除队列数据")

End Sub

Public Sub ClearQueueforQueueName(ByVal strQueueName As String)
'strQueueName: 队列名称
'功能：清除有效天数外指定队列的队列数据
    Dim strSql          As String

    strSql = "ZL_排队叫号队列_清除('" & strQueueName & "'," & ValidDays & "," & mintWorkType & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "清除队列数据")

End Sub

Public Sub ClearQueueData()
'intWorkType: 业务类型
'intValidDays: 有效天数
'功能：根据数据有效天数清除排队数据

    Dim strSql As String
    On Error GoTo errHandle

    strSql = "ZL_排队清除(" & mintWorkType & "," & ValidDays & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "清除排队数据")
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Public Sub UpdateQueue(ByVal lngQueueId As Long, ByVal strExpandData As String)
'lngQueueID: 排队ID
'arrUpdateColData: 的格式为/* "姓名='张四', 性别='女',年龄=10"*/
'功能：更新队列数据
    Dim strSql          As String
    
    strSql = "ZL_排队叫号队列_更新(" & lngQueueId & ",'" & strExpandData & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "更新队列数据")
    
End Sub

Public Function GetQueueInf(ByVal lngQueueId As Long, Optional strQueueFields As String) As ADODB.Recordset
'lngQueueId: 排队ID
'strQueueFields:需要查询的列，格式：'姓名，性别，年龄，...'
'功能：获取当前排队数据
    Dim rsRecord        As ADODB.Recordset
    
    On Error GoTo ErrorHand:
    Call SerachQueue("id=" & lngQueueId, rsRecord, strQueueFields)
    Set GetQueueInf = rsRecord
    
    Exit Function
ErrorHand:
    Set GetQueueInf = Nothing
End Function

Public Sub WriteTarget(ByVal lngQueueId As Long)
'lngQueueId: 排队ID
'功能：写入检查目的地
    Dim strSql          As String
    
    strSql = "ZL_排队叫号队列_插入目的地(" & lngQueueId & ",'" & CallTarget & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "写入检查目的地")

End Sub


Private Function GetAllCols() As String
'获取排队叫号表中的所有字段列
On Error GoTo errHandle
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strCols As String
    
    GetAllCols = "*"
    
    strSql = "select Column_Name as col from User_Tab_Columns where Table_Name='排队叫号队列'"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "查询排队叫号列")
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    strCols = ""
    rsTmp.MoveFirst
    
    While Not rsTmp.EOF
        If strCols <> "" Then strCols = strCols & ","
        strCols = strCols & Nvl(rsTmp!Col)
        
        rsTmp.MoveNext
    Wend
    
    GetAllCols = strCols
Exit Function
errHandle:
    GetAllCols = "*"
End Function


Public Sub OrderCall(ByVal strQueueName As String)
'strQueueName: 队列名称
'功能：顺序呼叫
    Dim strSql                  As String
    Dim blnCancel               As Boolean
    Dim rsQueue                 As ADODB.Recordset
    Dim rsTemp                  As ADODB.Recordset
    Dim strCallContext          As String
    
    blnCancel = False
    strCallContext = ""
    
    strSql = "select " & mstrDefQueryCols & _
             " from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and 排队状态=0 and rownum=1" & CustomOrder
             
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "顺序呼叫", strQueueName, mintWorkType)
    
    If rsQueue.RecordCount <= 0 Then Exit Sub
    
    rsQueue.MoveFirst
        
    RaiseEvent OnCallPreBefore(Nvl(rsQueue!Id), CallWay, strCallContext, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    If strCallContext = "" Then
        '获取呼叫内容
        strSql = "select zl_排队叫号队列_获取呼叫内容([1]) as 呼叫内容 from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取呼叫内容", Val(rsQueue!Id))
        strCallContext = Nvl(rsTemp!呼叫内容)
    End If
    
    strSql = "ZL_排队叫号队列_Call(" & Nvl(rsQueue!Id) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "根据呼叫修改排队状态")
    
    strSql = "Zl_排队语音呼叫_插入(" & Nvl(rsQueue!Id) & ",'" & LocalStation & "','" & strCallContext & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "写入呼叫内容")
    
    RaiseEvent OnCallPreAfter(Nvl(rsQueue!Id), CallWay)

End Sub

Public Sub SpecifiedCall(ByVal lngQueueId As Long)
'lngQueueId: 排队ID
'功能：直接呼叫
    Dim strSql              As String
    Dim blnCancel           As Boolean
    Dim strCallContext      As String
    Dim rsTemp              As ADODB.Recordset

    blnCancel = False
    strCallContext = ""
    
    RaiseEvent OnCallPreBefore(lngQueueId, CallWay, strCallContext, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    If strCallContext = "" Then
        '获取呼叫内容
        strSql = "select zl_排队叫号队列_获取呼叫内容([1]) as 呼叫内容 from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取呼叫内容", lngQueueId)
        strCallContext = Nvl(rsTemp!呼叫内容)
    End If
    
    strSql = "ZL_排队叫号队列_Call(" & lngQueueId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "根据呼叫修改排队状态")
    
    strSql = "Zl_排队语音呼叫_插入(" & lngQueueId & ",'" & LocalStation & "','" & strCallContext & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "写入呼叫内容")
    
    RaiseEvent OnCallPreAfter(lngQueueId, CallWay)
    
End Sub

Public Sub BroadcastCall(ByVal lngQueueId As Long)
'lngQueueId: 排队ID
'功能：广播呼叫
    Dim strSql              As String
    Dim blnCancel           As Boolean
    Dim strCallContext      As String
    Dim rsTemp              As ADODB.Recordset

    blnCancel = False
    strCallContext = ""
    
    RaiseEvent OnCallPreBefore(lngQueueId, CallWay, strCallContext, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    If strCallContext = "" Then
        '获取呼叫内容
        strSql = "select zl_排队叫号队列_获取呼叫内容([1]) as 呼叫内容 from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取呼叫内容", lngQueueId)
        strCallContext = Nvl(rsTemp!呼叫内容)
    End If
    
    strSql = "Zl_排队语音呼叫_插入(" & lngQueueId & ",'" & LocalStation & "','" & strCallContext & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "写入呼叫内容")
    
    RaiseEvent OnCallPreAfter(lngQueueId, CallWay)
    
End Sub


Private Function StartCall(ByVal strCallingContext As String) As Boolean
'lngQueueId：呼叫id
'strCallingContext:呼叫内容
'功能：完成呼叫
    Dim start                   As Date
    Dim intCurCallCount         As Long
    
    StartCall = False
    If strCallingContext = "" Then Exit Function
                        
    intCurCallCount = 0
    While (intCurCallCount < CallCount)
        
        start = Timer
        If CallWay = 1 Then
            Call MSSoundPlay(strCallingContext, CallSpeed)
        Else
            Call StartTextPlay(strCallingContext, CallSpeed * 10)
        End If
        
        intCurCallCount = intCurCallCount + 1
                                        
        Do While Timer < start + CallLen
            Call Sleep(5)
            
            DoEvents
        Loop
    Wend
    
    StartCall = True
End Function

Private Sub MSSoundPlay(ByVal strConnetxt As String, ByVal lngSoundSpeed As Long)
    Dim objVoice        As Object
    
    On Error Resume Next
    
    If objVoice Is Nothing Then
        Set objVoice = CreateObject("SAPI.SpVoice")
    End If
    
    objVoice.Rate = lngSoundSpeed   '速度:-10,10  0
    objVoice.Volume = 100 '声音:0,100   100
    objVoice.Speak strConnetxt, 1

End Sub

Public Sub CustomCall(ByVal strContext As String)
'strContext: 语音内容 '
'功能：自定呼叫
    Dim strSql          As String
    Dim blnCancel       As Boolean
    
'    id(序列生成),内容，站点，时间(系统时间)
    strSql = "Zl_排队语音呼叫_自定义插入('" & strContext & "','" & CallStation & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "自定呼叫")
        
End Sub

Public Sub StartQueue(ByVal lngQueueId As Long)
'lngQueueId: 排队ID
'功能：开始排队
    Dim strSql                  As String
    Dim blnCancel               As Boolean
    
    blnCancel = False
    
    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otStart, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    strSql = "zl_排队叫号队列_排队状态更新(" & lngQueueId & ",0)"
    Call zlDatabase.ExecuteProcedure(strSql, "开始排队")
    
    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otStart)

End Sub

Public Sub PauseQueue(ByVal lngQueueId As Long)
'lngQueueID: 排队ID
'功能：暂停排队
    Dim strSql                  As String
    Dim blnCancel               As Boolean
    
    blnCancel = False

    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otPause, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    strSql = "zl_排队叫号队列_排队状态更新(" & lngQueueId & ",3)"
    Call zlDatabase.ExecuteProcedure(strSql, "暂停排队")
    
    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otPause)

End Sub

Public Sub AbstainQueue(ByVal lngQueueId As Long)
'lngQueueID: 排队ID
'功能：放弃排队
    Dim strSql                      As String
    Dim blnCancel                   As Boolean
    
    blnCancel = False
    
    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otAbstain, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    strSql = "zl_排队叫号队列_排队状态更新(" & lngQueueId & ",2)"
    Call zlDatabase.ExecuteProcedure(strSql, "放弃排队")
    
    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otAbstain)

End Sub

Public Sub RestoreQueue(ByVal lngQueueId As Long, Optional ByVal strCustomOrder As String)
'lngQueueID: 排队ID
'功能：重新排队（排到最后位置）
    Dim strSql                  As String
    Dim blnCancel               As Boolean
    
    blnCancel = False
    
    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otRestore, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    '重新计算排队序号
    strSql = "zl_排队叫号队列_排队序号重置(" & lngQueueId & ",'" & strCustomOrder & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "重置计算排队序号")
    
    '恢复排队状态
    strSql = "zl_排队叫号队列_排队状态更新(" & lngQueueId & ",0)"
    Call zlDatabase.ExecuteProcedure(strSql, "重新排队")


    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otRestore)
    
End Sub

Public Sub CompleteQueue(ByVal lngQueueId As Long)
'lngQueueID: 排队ID
'功能：完成排队
    Dim strSql                      As String
    Dim blnCancel                   As Boolean
    
    blnCancel = False

    strSql = "zl_排队叫号队列_排队状态更新(" & lngQueueId & ",4)"
    
    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otComplete, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    Call zlDatabase.ExecuteProcedure(strSql, "完成排队")
    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otComplete)
    
End Sub

Public Sub DoCustomWork(ByVal lngQueueId As Long, ByVal lngOperationType As Long)
'lngQueueId: 排队ID
'lngOperationType：操作类型，为5及以后的数字
'功能：执行自定义业务操作，如接诊、回诊、转诊等

    Dim blnCancel       As Boolean

    blnCancel = False

    RaiseEvent OnWorkBefore(lngQueueId, lngOperationType, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    RaiseEvent OnWorkAfter(lngQueueId, lngOperationType)
End Sub

Public Function GetQueueState(ByVal lngQueueId As Long) As Long
'lngQueueID: 排队ID
'-1-占位中，0-排队中，1-呼叫中，2-已弃号，3-已暂停，4-完成就诊,5-弃用，6-弃用，7-已呼叫，
'功能：获取排队状态，失败返回-2
    Dim strSql              As String
    Dim rsQueue             As ADODB.Recordset
    
    On Error GoTo ErrorHand
    GetQueueState = -2
    strSql = "select 排队状态 from 排队叫号队列 where id=[1]"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "排队状态", lngQueueId)
    If rsQueue.RecordCount > 0 Then GetQueueState = Nvl(rsQueue!排队状态)
    
    Exit Function
ErrorHand:
    GetQueueState = -2
End Function

Public Sub DoLoopCall()
'功能：检索属于本站点的语音呼叫内容，并进行语音播放
    Dim strSql              As String
    Dim rsQueue             As ADODB.Recordset
    Dim blnCancel           As Boolean
    
    blnCancel = False
    
    strSql = "select id,呼叫内容 from 排队语音呼叫  where 站点=[1] order by 生成时间"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "语音呼叫内容", CallStation)

    Do While Not rsQueue.EOF
        RaiseEvent OnCallingBefore(Nvl(rsQueue!Id), Nvl(rsQueue!呼叫内容), blnCancel)
        
        If blnCancel = True Then Exit Sub
        
        '启用自定义呼叫后就不用语音呼叫
        If UseVbsCall = True Then
            Call DoScriptCall(Nvl(rsQueue!Id), Nvl(rsQueue!呼叫内容))
        Else
            If StartCall(Nvl(rsQueue!呼叫内容)) = False Then Exit Sub
        End If
        
        RaiseEvent OnCallingAfter(Nvl(rsQueue!Id), Nvl(rsQueue!呼叫内容))
        
        '呼叫成功后删除呼叫过的内容
        strSql = "ZL_排队语音呼叫_DELETE(" & Nvl(rsQueue!Id) & ")"
        zlDatabase.ExecuteProcedure strSql, "语音呼叫完成"
        
        rsQueue.MoveNext
    Loop
End Sub

Private Sub DoScriptCall(ByVal lngCallId As Long, ByVal strCallContext As String)
'调用vbs脚本实现自定义呼叫
'lngCallId:待呼叫id
'strCallContext: 呼叫内容
    Dim strMethod As String
    Dim lngIndex As Long
    Dim objCall
   
    If strCallContext = "" Then Exit Sub
    
    Set objCall = CreateObject("ScriptControl")
   
    strMethod = CallScript
    lngIndex = InStr(UCase(strMethod), "FUNCTION")
    
    If lngIndex <= 0 Then
        lngIndex = InStr(UCase(strMethod), "SUB")
        '设置过程名的开始位置
        lngIndex = lngIndex + 4
    Else
        lngIndex = lngIndex + 9
    End If
    
    strMethod = Mid(strMethod, lngIndex, InStr(strMethod, "(") - lngIndex)
    
    objCall.Language = "vbscript"
   
    Call objCall.AddCode(CallScript)
    Call objCall.Run(Trim(strMethod), lngCallId, strCallContext)
End Sub

 
Public Sub ChangeOrder(ByVal lngQueueId As Long, ByVal strNewOrder As String)
'lngQueueID: 排队ID
'strNewOrder: 新的排队序号
'功能：调整排队顺序
    Dim strSql As String
    
    strSql = "zl_排队叫号队列_排队序号更新(" & lngQueueId & ",'" & strNewOrder & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "更新排队序号")
    
End Sub

Public Function GetOrder(ByVal lngQueueId As Long) As String
'lngQueueID: 排队ID
'功能：返回对应的排队序号，失败返回-1
    Dim strSql          As String
    Dim rsReocrd        As ADODB.Recordset
    
    On Error GoTo ErrorHand
    GetOrder = ""
    
    strSql = "select 排队序号 from 排队叫号队列 where id=[1]"
    Set rsReocrd = zlDatabase.OpenSQLRecord(strSql, "获取排队序号", lngQueueId)
    
    If rsReocrd.RecordCount > 0 Then GetOrder = Nvl(rsReocrd!排队序号)
    
    Exit Function
ErrorHand:
    GetOrder = ""
End Function

Public Function GetCustomOrderWhereStr() As String
'功能：获取自定义排序条件
    Dim strSql          As String
    Dim rsQueue         As ADODB.Recordset
    
    On Error GoTo ErrorHand
    
    GetCustomOrderWhereStr = ""
    
    strSql = "select zl_排队叫号队列_获取排序方式([1]) as 排序条件 from dual"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "获取自定义排序条件", mintWorkType)
    
    If rsQueue.RecordCount > 0 Then GetCustomOrderWhereStr = Nvl(rsQueue!排序条件)
    
    Exit Function
ErrorHand:
    GetCustomOrderWhereStr = ""
End Function

Public Function GetInsertOrder(ByVal lngInsertQueueId As Long, _
                               ByVal strOrderNo1 As String, _
                               ByVal strOrderNo2 As String) As String
'功能：获取插队序号
    Dim strSql          As String
    Dim rsQueue         As ADODB.Recordset
    
On Error GoTo ErrorHand
    GetInsertOrder = ""
    
    strSql = "select zl_排队叫号队列_获取插队序号([1],[2],[3]) as 插队序号 from dual"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "获取插队序号", lngInsertQueueId, strOrderNo1, strOrderNo2)
    
    If rsQueue.RecordCount > 0 Then GetInsertOrder = Nvl(rsQueue!插队序号)
        
    Exit Function
ErrorHand:
    GetInsertOrder = ""
End Function


Public Function GetMinOrder(ByVal strQueueName As String) As String
'strQueueName: 队列名称
'功能：获取当前队列最小排队号码
    Dim strSql          As String
    Dim rsReocrd        As ADODB.Recordset
    
On Error GoTo ErrorHand
    GetMinOrder = ""
    
    
    strSql = "select 排队序号 from (select 排队序号 from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] " & CustomOrder & " ) a where rownum =1"
    
    Set rsReocrd = zlDatabase.OpenSQLRecord(strSql, "获取当前队列最小排队号码", strQueueName, mintWorkType)
    
    If rsReocrd.RecordCount > 0 Then GetMinOrder = Nvl(rsReocrd!排队序号)
    
    Exit Function
ErrorHand:
    GetMinOrder = ""
End Function

Public Function GetMaxQueueNo(Optional ByVal strFilter As String) As String
'strFilter:获取最大号码的自定义条件，默认为空表示当前业务下的最大排队号码,格式为/*"性别='男',年龄=10,..."*/
'如: '队列名称=超声科'，则获取队列名称为超声科的当前最大号码；‘排队标记=黑白超'，则获取排队标记为黑白超的当前最大号码
'功能：获取指定条件下的最大排队号码（排队号）
    Dim strSql          As String
    Dim strNewFilter    As String
    Dim varData(20)     As Variant
    Dim rsQueue         As ADODB.Recordset
   
On Error GoTo ErrorHand
    GetMaxQueueNo = ""
    
    If strFilter = "" Then
        strFilter = "业务类型=[1]"
        varData(0) = mintWorkType
    Else
        Call GetFilterCfg(strFilter, strNewFilter, varData)
    End If
    
    strSql = "select max(to_number(排队号码)) as 排队号码 from 排队叫号队列 where " & strFilter
        
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "获取指定条件下的最大排队号码", varData(0), varData(1), varData(2), varData(3), _
                                                    varData(4), varData(5), varData(6), varData(7), varData(8), varData(9), _
                                                    varData(10), varData(11), varData(12), varData(13), varData(14), varData(15), _
                                                    varData(16), varData(17), varData(18), varData(19), varData(20))
    
    If rsQueue.RecordCount > 0 Then GetMaxQueueNo = Nvl(rsQueue!排队号码)
    
    Exit Function
ErrorHand:
    GetMaxQueueNo = ""
End Function

Public Function SerachQueue(ByVal strFilter As String, rsRecord As ADODB.Recordset, Optional strQueueFields As String) As Long
'strFilter: 过滤条件,格式为/*"性别='男',年龄=10,..."*/
'返回检索到的排队ID
'功能：根据条件检索，并返回排队ID,失败返回-1
    Dim strSql          As String
    Dim strNewFilter As String
    Dim varData(20)     As Variant

On Error GoTo ErrorHand
    
    SerachQueue = -1
    Set rsRecord = Nothing
    
    If strFilter = "" Then Exit Function
    
    Call GetFilterCfg(strFilter, strNewFilter, varData)
    
    If strQueueFields <> "" Then
        strSql = "select id," & strQueueFields & " from 排队叫号队列 where " & strNewFilter
    Else
        strSql = "select " & mstrDefQueryCols & " from 排队叫号队列 where " & strFilter
    End If

    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "返回检索到的排队ID", varData(0), varData(1), varData(2), varData(3), _
                                                    varData(4), varData(5), varData(6), varData(7), varData(8), varData(9), _
                                                    varData(10), varData(11), varData(12), varData(13), varData(14), varData(15), _
                                                    varData(16), varData(17), varData(18), varData(19), varData(20))
    If rsRecord.RecordCount > 0 Then SerachQueue = Nvl(rsRecord!Id)
    
    Exit Function
ErrorHand:
    SerachQueue = -1
End Function

Private Sub GetFilterCfg(ByVal strFilter As String, ByRef strNewFilter As String, ByRef varParament() As Variant)
    Dim i               As Integer
    Dim varPara(20)     As Variant
    Dim varRela(20)     As Variant
    Dim strResult       As String
    Dim strValue        As String
    
    strFilter = Trim(strFilter) & " "
    strResult = strFilter
    strFilter = ""
    
    Do While InStr(strResult, "=") > 0
        varPara(i) = Mid(strResult, 1, InStr(strResult, "=") - 1)
        
        strValue = Trim(Mid(strResult, InStr(strResult, "=") + 1, InStr(strResult, " ") - InStr(strResult, "=") - 1))
            
        If IsNumeric(strValue) = True Then
            varParament(i) = Val(strValue)
        ElseIf IsDate(strValue) = True Then
            varParament(i) = CDate(strValue)
        Else
            If InStr(strValue, "(") > 0 Then
                varParament(i) = Mid(strValue, 3, Len(strValue) - 4)
            Else
                varParament(i) = Mid(strValue, 2, Len(strValue) - 2)
            End If
        End If
        
        strResult = LTrim(Mid(strResult, InStr(strResult, "=") + Len(strValue) + 1))
        
        If strResult <> "" Then
            varRela(i) = Mid(strResult, 1, InStr(strResult, " ") - 1)
            strResult = Mid(strResult, InStr(strResult, " ") + 1)
        End If
        
        strFilter = strFilter & varPara(i) & "=[" & i + 1 & "] " & varRela(i) & " "
        i = i + 1
    Loop
    
    strNewFilter = Trim(strFilter)
End Sub

Private Function AnalyseComputer() As String
'获取本地计算机名称
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


'初始化各属性值
Private Sub Class_Initialize()
'    mintCallCount = 1
'    mintCallLen = 10
End Sub

'清除指定天数前的数据,即数据有效天数
Public Property Get ValidDays() As Integer
    ValidDays = mintValidDays
End Property

Public Property Let ValidDays(ByVal vNewValue As Integer)
    mintValidDays = vNewValue
End Property

'呼叫方式
Public Property Get CallWay() As Integer
    CallWay = mintCallWay
End Property

Public Property Let CallWay(ByVal vNewValue As Integer)
    mintCallWay = vNewValue
End Property

'呼叫次数
Public Property Get CallCount() As Integer
    CallCount = mintCallCount
End Property

Public Property Let CallCount(ByVal vNewValue As Integer)
    mintCallCount = vNewValue
End Property

'呼叫速度
Public Property Get CallSpeed() As Integer
    CallSpeed = mintCallSpeed
End Property

Public Property Let CallSpeed(ByVal vNewValue As Integer)
    mintCallSpeed = vNewValue
End Property

'呼叫时间长度
Public Property Get CallLen() As Integer
    CallLen = mintCallLen
End Property

Public Property Let CallLen(ByVal vNewValue As Integer)
    mintCallLen = vNewValue
End Property

'本机站点名
Public Property Get LocalStation() As Variant
    LocalStation = mstrLocalStation
End Property

Public Property Let LocalStation(ByVal vNewValue As Variant)
    mstrLocalStation = vNewValue
End Property

'待呼叫站点名
Public Property Get CallStation() As String
    CallStation = mstrCallStation
End Property

Public Property Let CallStation(ByVal vNewValue As String)
    mstrCallStation = vNewValue
End Property

'呼叫前设置病人的检查目的地
Public Property Get CallTarget() As String
    CallTarget = mstrCallTarget
End Property

Public Property Let CallTarget(ByVal vNewValue As String)
    mstrCallTarget = vNewValue
End Property

'自定义排序
Public Property Get CustomOrder() As String
    CustomOrder = mstrCustomOrder
End Property

Public Property Let CustomOrder(ByVal vNewValue As String)
    mstrCustomOrder = vNewValue
End Property

'默认查询列
Public Property Get DefQueryCols() As String
    DefQueryCols = mstrDefQueryCols
End Property

Public Property Get CallScript() As String
    CallScript = mstrCallScript
End Property

Public Property Let CallScript(ByVal vNewValue As String)
    mstrCallScript = vNewValue
End Property

Public Property Get UseVbsCall() As Boolean
    UseVbsCall = mblnUseVbsCall
End Property

Public Property Let UseVbsCall(ByVal vNewValue As Boolean)
    UseVbsCall = vNewValue
End Property




