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
Attribute VB_Exposed = True
Option Explicit

'呼叫方式
Public Enum tcallway
    cwOrder = 0         '顺乎
    cwSpecify = 1       '直呼
    cwBroadcast = 2     '广播
    cwWaitRoom = 3      '候诊
End Enum

'操作类型
Public Enum TOperationType
    otStart = 0         '排队
    otPause = 1         '暂停
    otAbstain = 2       '弃号
    otRestore = 3       '重排
    otComplete = 4      '完成
    otPriority = 5      '优先
    otInsertQueue = 6   '插队
    otDiagnose = 7      '接诊
    otUpdateMemo = 8    '更新备注信息
    otPrintNo = 9       '打号
End Enum


'队列状态
Public Enum TQueueState
    qsPlace = -1        '占位中，当进入队列时，为占位状态
    qsQueueing = 0      '排队中，执行开始排队后，进入该状态
    qsCalling = 1       '呼叫中，当前正在呼叫的数据
    qsAbstain = 2       '已弃号，放弃呼叫的数据
    qsPause = 3         '已暂停，暂停呼叫的数据
    qsComplete = 4      '已完成，完成呼叫的数据
    qsBeforeWaitCall = 5     '待预呼
    qsAfterWaitCall = 6    '已预呼
    qsCalled = 7        '已呼叫，已执行呼叫但并非当前呼叫的数据
    qsDiagnose = 8      '接诊中
    qsWaitCall = 9      '待呼叫
End Enum


Private mobjVoice As Object                  '语音呼叫对象
Private mobjLCDShow As Object               'LCDShow对象
Private WithEvents mobjMsg As clsQueueMsgCenter
Attribute mobjMsg.VB_VarHelpID = -1

Private mstrLoginUser           As String
Private mintValidDays           As Integer  '数据有效天数
Private mstrVoiceType           As String   '微软语音类型名称
Private mblnIsPlayHintSound     As Boolean  '呼叫前播放提示音
Private mlngPlayCount           As Long     '呼叫次数 默认为1
Private mlngPlaySpeed           As Long     '语音播放速度
Private mlngPlayTimeLength      As Long     '呼叫时间长度，默认10秒
Private mstrPlayStation         As String   '播放语音的站点名,默认当前站点
Private mstrCallTarget          As String   '呼叫前设置病人的检查目的地
Private mstrLocalStation        As String   '本机站点名
Private mlngWorkType            As Long     '业务类型
Private mcnOracle               As ADODB.Connection
Private mstrCustomOrder         As String   '自定义排序
Private mstrDefQueryCols        As String   '默认查询列，默认为排队叫号队列表中的所有列
Private mstrCusVoiceScript      As String   '自定义语音播放脚本
Private mblnUseVbsPlay          As Boolean
Private mblnAbortState          As Boolean
Private mstrReportNum           As String   '报表编号

'语音播放的函数
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Sub UseMsgCenter(ByVal lngSys As Long, ByVal lngModule As Long, Optional ByVal strPrivs As String = "")
'启用消息中心
    If gobjMsgCenter Is Nothing Then
        Set gobjMsgCenter = New clsQueueMsgCenter
        Call gobjMsgCenter.OpenMsgCenter(lngSys, lngModule, mlngWorkType, strPrivs)
    End If
    
    Set mobjMsg = gobjMsgCenter
        
End Sub

Public Sub InitQueue(cnOracle As ADODB.Connection, ByVal lngWorkType As Long, Optional ByVal strLoginUser As String = "system")
'cnOracle:数据库连接对象
'lngSysNo
'功能：初始化队列
    
    Set mcnOracle = cnOracle
'    InitCommon cnOracle
    
    mlngWorkType = lngWorkType
    
    mstrLoginUser = strLoginUser
    mstrCustomOrder = GetCustomOrderWhere
    mstrLocalStation = AnalyseComputer
    mstrPlayStation = AnalyseComputer
    mstrDefQueryCols = GetAllQueueTabCols

    Set mobjVoice = Nothing
    
    '更新业务类型
    If Not mobjMsg Is Nothing Then
        Call mobjMsg.ConfigMsgBusinessType(mlngWorkType)
    End If
End Sub


Public Function InsertQueue(ByVal strQueueName As String, _
                            Optional ByVal strQueueNo As String = "", _
                            Optional ByVal lngBusinessId As Long = 0, _
                            Optional ByVal strName As String = "", _
                            Optional ByVal strTarget As String = "", _
                            Optional ByVal strCustomOrder As String = "", _
                            Optional ByVal strExpandData As String = "") As Long
'strQueueName: 要插入的队列名称
'strQueueNo: 排队号码
'lngBusinessId: 业务ID
'strName: 姓名
'strTarget: 目的地
'strCustomOrder: 自定义排队序号
'strExpandData:扩展数据 格式为/* "姓名='张四', 性别='女',年龄=10"*/
'功能：返回插入后的排队ID，失败返回-1
    Dim strSql          As String
    Dim rsQueue         As ADODB.Recordset
    
    On Error GoTo ErrorHand
    
    InsertQueue = -1
    
    strSql = "select zl_排队叫号队列_插入([1],[2],[3], " & IIf(lngBusinessId = 0, "null", "[4]") & ",[5],[6],[7],[8]) as 排队id from dual"
    
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "返回插入后的排队ID", strQueueName, mlngWorkType, _
                                        strQueueNo, lngBusinessId, strName, strTarget, strCustomOrder, strExpandData)
                                        
    If rsQueue.RecordCount > 0 Then InsertQueue = Nvl(rsQueue!排队id)
    
    Exit Function
ErrorHand:
    InsertQueue = -1
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub DeleteQueue(ByVal lngQueueId As Long)
'lngQueueID: 排队ID

'功能：删除队列数据
    Dim strSql          As String
    
    strSql = "ZL_排队叫号队列_删除(" & lngQueueId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "删除队列数据")

End Sub

Public Sub ClearQueueData(ByVal strQueueName As String, Optional ByVal blnDelAll As Boolean = False)
'strQueueName: 队列名称
'blnDelAll:是否删除所有属于该队列的数据，如果参数传入true，则排除有效天数，删除所有数据

'功能：清除指定队列下的排队数据

    Dim strSql          As String
    Dim lngValidDays As Long
    
    lngValidDays = mintValidDays
    If blnDelAll Then lngValidDays = 0

    strSql = "ZL_排队叫号队列_清除队列('" & strQueueName & "'," & lngValidDays & "," & mlngWorkType & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "清除队列数据")

End Sub

Public Sub ClearBusinessData(Optional ByVal blnDelAll As Boolean = False)
'blnDelAll:是否删除该业务下的所有数据，如果为true，则清除所有数据

'功能：清除业务类型下的排队数据

    Dim strSql As String
    Dim lngValidDays As Long
    
    lngValidDays = mintValidDays
    If blnDelAll Then lngValidDays = 0

    strSql = "ZL_排队叫号队列_清除业务(" & mlngWorkType & "," & lngValidDays & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "清除排队数据")
End Sub

Public Sub CustomClearData(ByVal strFilter As String, Optional ByVal blnDelAll As Boolean = False)
'自定义清除数据
    Dim strSql As String
    Dim lngValidDays As Long
    
    lngValidDays = mintValidDays
    If blnDelAll Then lngValidDays = 0

    strSql = "ZL_排队叫号队列_自定义清除(" & mlngWorkType & ",'" & strFilter & "'," & lngValidDays & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "清除排队数据")
End Sub

Public Sub ClearVoiceData(Optional ByVal lngQueueId As Long = 0, Optional ByVal blnDelAll As Boolean = False)
'lngQueueId:需要删除的呼叫队列数据，如果未指定，则删除所有

'清除排队预约呼叫中的数据
    Dim strSql As String
    Dim lngValidDays As Long
    
    lngValidDays = mintValidDays
    If blnDelAll Then lngValidDays = 0
    
    strSql = "ZL_排队叫号队列_清除呼叫(" & lngQueueId & "," & mlngWorkType & "," & lngValidDays & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "清除语音呼叫数据")
End Sub


Public Sub UpdateQueue(ByVal lngQueueId As Long, ByVal strExpandData As String)
'lngQueueID: 排队ID
'arrUpdateColData: 的格式为/* "姓名='张四', 性别='女',年龄=10"*/
'功能：更新队列数据
    Dim strSql          As String
    
    strSql = "ZL_排队叫号队列_更新(" & lngQueueId & ",'" & strExpandData & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "更新队列数据")
    
End Sub

Public Function GetQueueInf(ByVal lngQueueId As Long, Optional strReturnFields As String = "") As ADODB.Recordset
'lngQueueId: 排队ID
'strReturnFields:需要查询的列，格式：'姓名，性别，年龄，...'

'功能：获取当前排队数据
On Error GoTo ErrorHand:
    Dim rsRecord        As ADODB.Recordset
    
    Call SerachQueue("id=" & lngQueueId, rsRecord, strReturnFields)
    Set GetQueueInf = rsRecord
    
    Exit Function
ErrorHand:
    Set GetQueueInf = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub WriteTarget(ByVal lngQueueId As Long, Optional ByVal strSpecifyTarget As String = "")
'lngQueueId: 排队ID
'strSpecifyTarget:指定排队叫号的检查目的地，如果为空，则使用CallTarget属性进行设置

'功能：写入检查目的地
    Dim strSql As String
    Dim strTarget As String
    
    strTarget = mstrCallTarget
    If strSpecifyTarget <> "" Then strTarget = strSpecifyTarget
    
    strSql = "ZL_排队叫号队列_插入目的地(" & lngQueueId & ",'" & strTarget & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "写入检查目的地")

End Sub


Public Function GetAllQueueTabCols(Optional ByVal strTabPrefix As String = "") As String
'获取排队叫号表中的所有字段列
On Error GoTo errHandle
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strCols As String
    
    GetAllQueueTabCols = "*"
    
    strSql = "select Column_Name as col from all_tab_columns  where Table_Name='排队叫号队列' "

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "查询排队叫号列")
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    strCols = ""
    rsTmp.MoveFirst
    
    While Not rsTmp.EOF
        If strCols <> "" Then strCols = strCols & ","
        strCols = strCols & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & Nvl(rsTmp!Col)
        
        rsTmp.MoveNext
    Wend
    
    GetAllQueueTabCols = strCols
Exit Function
errHandle:
    GetAllQueueTabCols = "*"
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function GetQueueIdWithOrderCall(ByVal strQueueName As String) As Long
'功能:获取顺序呼叫的队列ID
'strQueueName: 队列名称

    Dim strSql                  As String
    Dim rsQueue                 As ADODB.Recordset
    
    GetQueueIdWithOrderCall = -1
    
    '获取当前队列中，排队顺序处于第一位的排队数据
    strSql = "select ID from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] and 排队状态=0 and rownum=1" & _
             IIf(Trim(mstrCustomOrder) <> "", " order by " & mstrCustomOrder, "")
             
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "顺序呼叫", strQueueName, mlngWorkType)
    
    If rsQueue.RecordCount <= 0 Then Exit Function
    
    rsQueue.MoveFirst
    
    GetQueueIdWithOrderCall = Nvl(rsQueue!Id)
End Function

Public Function WaitRoomCall(ByVal lngQueueId As Long, Optional ByVal strCallContext As String = "") As Boolean
'lngQueueId: 排队ID

'功能：候诊呼叫准备（呼叫进入排队队列的患者，到某个诊室等待就诊）
'开始排队后，才能进行候诊呼叫
    Dim strSql              As String
    Dim rsTemp              As ADODB.Recordset

    WaitRoomCall = False
    
    If strCallContext = "" Then
        '获取呼叫内容
        strSql = "select zl_排队叫号队列_获取呼叫内容([1],[2]) as 呼叫内容 from dual"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取呼叫内容", lngQueueId, tcallway.cwWaitRoom)
        
        strCallContext = Nvl(rsTemp!呼叫内容)
    End If
    
    strSql = "Zl_排队叫号队列_状态更新(" & lngQueueId & "," & qsBeforeWaitCall & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "更新呼叫状态")
    
    strSql = "select Zl_排队语音呼叫_插入([1],[2],[3],[4]) as 语音ID from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "写入呼叫内容", lngQueueId, mlngWorkType, mstrPlayStation, strCallContext)
    
    WaitRoomCall = True
End Function

Public Function SpecifiedCall(ByVal lngQueueId As Long, Optional ByVal strCallContext As String = "", Optional ByVal TypeCallWay As tcallway) As Long
'功能：直接呼叫准备
'lngQueueId: 排队ID
'返回生成的语音ID
    Dim strSql              As String
    Dim rsTemp              As ADODB.Recordset

    SpecifiedCall = -1
    
    '判断如果排队语音呼叫中存在该队列的呼叫内容，则不进行呼叫
    strSql = "select ID as 语音ID,队列ID from 排队语音呼叫 where rownum <= 1 and 队列Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询排队语音呼叫记录", lngQueueId)
    
    If rsTemp.RecordCount <= 0 Then
    
        If strCallContext = "" Then
            '获取呼叫内容
            strSql = "select zl_排队叫号队列_获取呼叫内容([1],[2]) as 呼叫内容 from dual"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取呼叫内容", lngQueueId, TypeCallWay)
            
            strCallContext = Nvl(rsTemp!呼叫内容)
        End If
        
        strSql = "ZL_排队叫号队列_Call(" & lngQueueId & ",'" & mstrLoginUser & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "根据呼叫修改排队状态")
        
        strSql = "select Zl_排队语音呼叫_插入([1],[2],[3],[4]) as 语音ID from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "写入呼叫内容", lngQueueId, mlngWorkType, mstrPlayStation, strCallContext)
    Else
        '如果已经有呼叫内容，则只更新呼叫状态
        strSql = "ZL_排队叫号队列_Call(" & lngQueueId & ",'" & mstrLoginUser & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "根据呼叫修改排队状态")
    End If
    
    If rsTemp.RecordCount > 0 Then
        SpecifiedCall = Val(Nvl(rsTemp!语音ID))
    End If
    
    If Not mobjMsg Is Nothing Then
        '发送状态改变消息
        Call mobjMsg.SendMsg_StateSync(lngQueueId)
        
        '发送语音播放消息
        Call mobjMsg.SendMsg_VoicePlay(lngQueueId)
    End If
End Function

Public Function BroadcastCall(ByVal lngQueueId As Long, Optional ByVal strCallContext As String = "") As Long
'功能：广播呼叫准备，在对排队数据广播时，更新状态为待呼叫
'lngQueueId: 排队ID
'返回生成的语音Id

    Dim strSql              As String
    Dim rsTemp              As ADODB.Recordset

    BroadcastCall = -1
    
    If strCallContext = "" Then
        '获取呼叫内容
        strSql = "select zl_排队叫号队列_获取呼叫内容([1],[2]) as 呼叫内容 from dual"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取呼叫内容", lngQueueId, tcallway.cwBroadcast)
        
        strCallContext = Nvl(rsTemp!呼叫内容)
    End If
    
    strSql = "select Zl_排队语音呼叫_插入([1],[2],[3],[4]) as 语音ID from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "写入呼叫内容", lngQueueId, mlngWorkType, mstrPlayStation, strCallContext)
    
    If rsTemp.RecordCount > 0 Then
        BroadcastCall = Val(Nvl(rsTemp!语音ID))
        
        '更新状态为待呼叫
        strSql = "ZL_排队叫号队列_Call(" & lngQueueId & ",'" & mstrLoginUser & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "根据呼叫修改排队状态")
    End If
    
    If Not mobjMsg Is Nothing Then
        '发送语音播放消息
        Call mobjMsg.SendMsg_VoicePlay(lngQueueId)
    End If
End Function


Private Function StartCall(ByVal strCallingContext As String) As Boolean
'lngQueueId：呼叫id
'strCallingContext:呼叫内容

'功能：完成呼叫
    Dim dtStart                   As Date
    Dim intCurCallCount         As Long
    
    StartCall = False
    If strCallingContext = "" Then Exit Function
                        
    intCurCallCount = 0
    While (intCurCallCount < mlngPlayCount)
        '如果被强行结束，这里则直接退出
        If mblnAbortState Then
            StartCall = True
            Exit Function
        End If
            
        dtStart = Timer
        
        '微软语音播放
        Call MSSoundPlay(strCallingContext, mlngPlaySpeed)
        
        intCurCallCount = intCurCallCount + 1
                                        
        Do While Timer < dtStart + mlngPlayTimeLength
            Call Sleep(15)
            
            If mblnAbortState Then
                StartCall = True
                Exit Function
            End If
            
            DoEvents
        Loop
    Wend
    
    StartCall = True
End Function

Private Sub PlayHintSound()
'播放提示声音
    On Error Resume Next
    
    '该方法主要是利用微软语音功能进行播放
    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
    End If
    
    Set mobjVoice.Voice = mobjVoice.GetVoices("Name=" & mstrVoiceType).Item(0)
    mobjVoice.Speak "DingDong.wav", 4
End Sub


Private Sub MSSoundPlay(ByVal strConnetxt As String, ByVal lngSoundSpeed As Long)
'strConnetxt:播放为语音的字符串
'lngSoundSpeed:播放速度

'功能：播放文本语音
    On Error Resume Next
    
    '该方法主要是利用微软语音功能进行播放
    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
    End If
    
    Set mobjVoice.Voice = mobjVoice.GetVoices("Name=" & mstrVoiceType).Item(0)
    mobjVoice.Rate = lngSoundSpeed   '速度:-10,10  0
    mobjVoice.Volume = 100 '声音:0,100   100
    mobjVoice.Speak strConnetxt, 1

End Sub

Public Function CustomCall(ByVal strContext As String) As Long
'strContext: 语音内容 '
'功能：自定呼叫准备
'返回生成的语音Id
    Dim strSql          As String
    Dim rsTemp As ADODB.Recordset
    
    CustomCall = -1
    
'    id(序列生成),内容，站点，时间(系统时间)
    strSql = "select Zl_排队语音呼叫_自定义插入([1],[2]) as 语音ID from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "写入自定义呼叫内容", strContext, mstrPlayStation)
        
    If rsTemp.RecordCount > 0 Then
        CustomCall = Val(Nvl(rsTemp!语音ID))
    End If
End Function


Public Function FindQueueId(ByVal lngBusinessId As Long) As Long
'根据业务ID查找队列ID
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Id from 排队叫号队列 where 业务类型=[1] and 业务ID=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询排队队列ID", mlngWorkType, lngBusinessId)
    
    FindQueueId = -1
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    FindQueueId = Val(Nvl(rsData!Id))
    
End Function

Public Function FindQueueInf(ByVal lngBusinessId As Long) As ADODB.Recordset
'根据业务ID查找队列信息
    Dim strSql As String
    
    strSql = "select ID,病人id,队列名称,科室id,排队号码,排队标记,诊室,排队时间,排队状态,呼叫时间 " & _
             "from 排队叫号队列 where 业务类型=[1] and 业务ID=[2]"
             
    Set FindQueueInf = zlDatabase.OpenSQLRecord(strSql, "查询排队队列信息", mlngWorkType, lngBusinessId)
End Function

Public Function LineQueue(ByVal lngQueueId As Long, ByRef strNewQueueNo As String, _
    Optional ByVal blnVerifyQueueNo As Boolean = True) As Boolean
'开始排队或恢复排队，改变排队状态为0
'blnVerifyQueueNo:排队时，是否对排队号码进行验证，如果进行验证，则可保证当天的排队号码唯一
'lngQueueId: 排队ID
'功能：开始排队
    Dim strSql                  As String
    Dim rsData As ADODB.Recordset
    Dim lngMsgResult As Long
    Dim curDate As Date
    
    LineQueue = False
    strNewQueueNo = ""
    
    '需要判断排队时间是否为当前时间，如果不是则需要重新排队
    If blnVerifyQueueNo = True Then
        strSql = "select 排队时间,队列名称 from 排队叫号队列 where Id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询排队时间", lngQueueId)
        If rsData.RecordCount > 0 Then
            curDate = zlDatabase.Currentdate
            
            '如果排队时间不是当天时间，则需要重新产生排队号码
            If Format(Nvl(rsData!排队时间, curDate), "yyyy-mm-dd") <> Format(curDate, "yyyy-mm-dd") Then
                lngMsgResult = MsgBox("当前数据的排队日期为 [" & Format(Nvl(rsData!排队时间, curDate), "yyyy-mm-dd") & "],与当前日期不一致，将重新进行排队，是否继续？", _
                                        vbYesNo, "提示")
                If lngMsgResult = vbNo Then Exit Function
                
                strNewQueueNo = GetQueueMaxNo(Nvl(rsData!队列名称))
                
                Call UpdateQueue(lngQueueId, "排队号码=''" & strNewQueueNo & "'', 排队时间=sysdate")
                
                '重新排序
                Call RestoreQueue(lngQueueId)
                
                LineQueue = True
                
                Exit Function
            End If
        End If
    End If
    
    strSql = "zl_排队叫号队列_状态更新(" & lngQueueId & "," & TQueueState.qsQueueing & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "开始排队")
    
    If Not mobjMsg Is Nothing Then
        '发送排队消息
        Call mobjMsg.SendMsg_LineQueue(lngQueueId)
    End If
    
    LineQueue = True

End Function

Public Function PauseQueue(ByVal lngQueueId As Long) As Boolean
'lngQueueID: 排队ID
'功能：暂停排队
    Dim strSql                  As String
    
    PauseQueue = False
    
    strSql = "zl_排队叫号队列_状态更新(" & lngQueueId & "," & TQueueState.qsPause & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "暂停排队")
    
    If Not mobjMsg Is Nothing Then
        '发送状态同步消息
        Call mobjMsg.SendMsg_StateSync(lngQueueId)
    End If
    
    PauseQueue = True
End Function

Public Function AbstainQueue(ByVal lngQueueId As Long) As Boolean
'lngQueueID: 排队ID
'功能：放弃排队
    Dim strSql                      As String
    
    AbstainQueue = False
    
    strSql = "zl_排队叫号队列_状态更新(" & lngQueueId & "," & TQueueState.qsAbstain & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "放弃排队")

    If Not mobjMsg Is Nothing Then
        '发送状态同步消息
        Call mobjMsg.SendMsg_StateSync(lngQueueId)
    End If
    
    AbstainQueue = True
End Function

Public Function RestoreQueue(ByVal lngQueueId As Long, Optional ByVal strCustomOrder As String) As String
'lngQueueID: 排队ID
'功能：重新排队（排到最后位置）
    Dim strSql                  As String
    Dim rsData As ADODB.Recordset

    RestoreQueue = ""
    
    '重新计算排队序号
    strSql = "select zl_排队叫号队列_序号重置([1], [2]) as 排队序号 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "重置计算排队序号", lngQueueId, strCustomOrder)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    '恢复排队状态
    strSql = "zl_排队叫号队列_状态更新(" & lngQueueId & "," & TQueueState.qsQueueing & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "重新排队")
    
    If Not mobjMsg Is Nothing Then
            '发送排队消息
        Call mobjMsg.SendMsg_LineQueue(lngQueueId)
    End If
    
    RestoreQueue = Nvl(rsData!排队序号)
End Function

Public Function CompleteQueue(ByVal lngQueueId As Long) As Boolean
'lngQueueID: 排队ID
'功能：完成排队
    Dim strSql                      As String
    
    CompleteQueue = False
    
    strSql = "zl_排队叫号队列_状态更新(" & lngQueueId & "," & TQueueState.qsComplete & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "完成排队")
    
    If Not mobjMsg Is Nothing Then
            '发送队列完成消息
        Call mobjMsg.SendMsg_CompleteQueue(lngQueueId)
    End If
    
    CompleteQueue = True
End Function

Public Function DiagnoseQueue(ByVal lngQueueId As Long) As Boolean
'lngQueueID: 排队ID
'功能：接诊队列
    Dim strSql                      As String

    DiagnoseQueue = False
    
    strSql = "zl_排队叫号队列_状态更新(" & lngQueueId & "," & TQueueState.qsDiagnose & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "接诊排队")

    If Not mobjMsg Is Nothing Then
        '发送状态同步消息
        Call mobjMsg.SendMsg_StateSync(lngQueueId)
    End If
    
    DiagnoseQueue = True
End Function


'Public Function DoCustomWork(ByVal lngQueueId As Long, ByVal lngOperationType As Long) As Boolean
''lngQueueId: 排队ID
''lngOperationType：操作类型，为5及以后的数字
''功能：执行自定义业务操作，如接诊、回诊、转诊等
'
'    Dim blnCancel       As Boolean
'
'    blnCancel = False
'    DoCustomWork = False
'
'    RaiseEvent OnWorkBefore(lngQueueId, lngOperationType, blnCancel)
'
'    If blnCancel = True Then Exit Function
'
'    RaiseEvent OnWorkAfter(lngQueueId, lngOperationType)
'
'    DoCustomWork = True
'End Function


Public Function PrintQueueNo(ByVal lngQueueId As Long, Optional ByVal blnPrint As Boolean = False, _
    Optional objOwner As Object = Nothing) As Boolean
'打印排队号码
    Dim blnCancel As Boolean
    Dim objReportPrint As Object

    PrintQueueNo = False
    
    If Trim(mstrReportNum) = "" Then
        Err.Raise -1, "PrintQueueNo", "尚未设置打印所用的报表编号。"
        Exit Function
    End If

    Set objReportPrint = CreateObject("zl9Report.clsReport")
    If objReportPrint Is Nothing Then Exit Function
    
On Error GoTo errHandle

    '1:预览,2:打印
    Call objReportPrint.ReportOpen(mcnOracle, 0, mstrReportNum, objOwner, "ID=" & lngQueueId, "排队ID=" & lngQueueId, "队列ID=" & lngQueueId, "QUEUEID=" & lngQueueId, , IIf(blnPrint, 2, 1))
    
    PrintQueueNo = True
    
    Set objReportPrint = Nothing
Exit Function
errHandle:
    Set objReportPrint = Nothing
    Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function


Public Sub PrintSet(Optional objOwner As Object = Nothing)
'打印机设置
    Dim objReportPrint As Object
    
    If Trim(mstrReportNum) = "" Then
        Err.Raise -1, "PrintQueueNo", "尚未设置打印所用的报表编号。"
        Exit Sub
    End If

    Set objReportPrint = CreateObject("zl9Report.clsReport")
    If objReportPrint Is Nothing Then Exit Sub
    
On Error GoTo errHandle

    Call objReportPrint.ReportPrintSet(mcnOracle, 0, mstrReportNum, objOwner)
    
    Set objReportPrint = Nothing
Exit Sub
errHandle:
    Set objReportPrint = Nothing
    Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

Public Function GetQueueState(ByVal lngQueueId As Long) As Long
'lngQueueID: 排队ID
'-1-占位中，0-排队中，1-呼叫中，2-已弃号，3-已暂停，4-完成就诊,5-弃用，6-弃用，7-已呼叫，
'功能：获取排队状态，失败返回-2
On Error GoTo ErrorHand
    Dim strSql              As String
    Dim rsQueue             As ADODB.Recordset
        
    GetQueueState = -2
    
    strSql = "select 排队状态 from 排队叫号队列 where id=[1] and 业务类型=[2]"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "排队状态", lngQueueId, mlngWorkType)
    
    If rsQueue.RecordCount > 0 Then GetQueueState = Nvl(rsQueue!排队状态)
    
    Exit Function
ErrorHand:
    GetQueueState = -2
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function GetStateCount(ByVal strQueueName As String, ByVal lngQueueState As TQueueState) As Long
'strQueueName:队列名称
'lngQueueState:排队状态

'获取队列下指定排队状态的队列数
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    GetStateCount = 0
    
    strSql = "select count(*) as 数量 from 排队叫号队列 where 业务类型=[1] and 队列名称=[2] and 排队状态=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", mlngWorkType, strQueueName, lngQueueState)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    GetStateCount = Val(Nvl(rsTemp!数量))
End Function

Public Sub StopVoice()
'blnWaitFinish：是否等待终止完成
'功能：终止播放
    If mblnAbortState = True Then Exit Sub
    
    mblnAbortState = True
    
    '停止微软语音呼叫
    If Not (mobjVoice Is Nothing) Then
        Call mobjVoice.Speak("", 2)
    End If
End Sub

'Public Sub OpenLCDShow(objOwner As Object, ByVal strLCDComponentFile As String, _
'    ByVal strQueueNames As String, Optional varParEx As Variant)
''打开LCD显示
'
'    If mobjLCDShow Is Nothing Then
'        Set mobjLCDShow = CreateLCDObj(strLCDComponentFile)
'    End If
'
'    Call mobjLCDShow.zlShow(objOwner, mlngWorkType, strQueueNames, mcnOracle, varParEx)
'End Sub
'
'
'Public Sub ConfigLCDPar(objOwner As Object, ByVal strLCDComponentFile As String)
''配置LCD显示参数
'    If mobjLCDShow Is Nothing Then
'        Set mobjLCDShow = CreateLCDObj(strLCDComponentFile)
'    End If
'
'    Call mobjLCDShow.zlSetup(objOwner)
'End Sub
'
'Private Function CreateLCDObj(ByVal strLCDComponentFile As String) As Object
''创建lcdshow显示部件
'    Dim strComponent As String
'    Dim lngIndex As Long
'
'    Set CreateLCDObj = Nothing
'    strComponent = strLCDComponentFile
'
'    lngIndex = InStr(strComponent, ".")
'    If lngIndex > 0 Then
'        strComponent = Mid(strComponent, 1, lngIndex - 1)
'    End If
'
'    strComponent = strComponent & ".clsLCDShow"
'
'    Set CreateLCDObj = CreateObject(strComponent)
'End Function
'
'
'Public Sub CloseLCDShow()
''关闭LCD显示
'    If mobjLCDShow Is Nothing Then Exit Sub
'
'    Call mobjLCDShow.zlClose
'    Set mobjLCDShow = Nothing
'End Sub


Public Function GetVoiceIdWithOrderPlay() As Long
'按顺序获取需要播放的语音Id
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetVoiceIdWithOrderPlay = -1
    
    strSql = "select ID from 排队叫号队列 where 站点=[1] and rownum=1 order by 生成时间"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询待呼叫语音ID", mstrLocalStation)
    
    
    If rsData.RecordCount > 0 Then
        GetVoiceIdWithOrderPlay = Val(Nvl(rsData!Id))
    End If
End Function



Public Function PlayCustomVoice(ByVal lngVoiceId As Long, ByVal blnAutoDelVoice As Boolean, _
                                Optional ByVal strVoiceContext As String = "") As Boolean
'播放自定义的语音内容
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strCurVoice As String
    
    PlayCustomVoice = False
    
    strCurVoice = strVoiceContext
    If Trim(strVoiceContext) = "" Then
        strSql = "select 呼叫内容 from 排队语音呼叫  where ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询呼叫内容", lngVoiceId)
        
        If rsData.RecordCount <= 0 Then Exit Function
        
        strCurVoice = Nvl(rsData!呼叫内容)
    End If
    
    '启用自定义呼叫后就不用语音呼叫
    If mblnUseVbsPlay = True Then
        If Trim(mstrCusVoiceScript) = "" Then
            Call DebugMsg("clsQueueOperation", "PlayCustomVoice", "mstrCusVoiceScript", "自定义呼叫脚本为空。")
            Exit Function
        End If
        
        Call DoScriptCall(lngVoiceId, strCurVoice)
    Else
        If StartCall(strCurVoice) = False Then
            Exit Function
        End If
    End If
    
    '呼叫成功后删除呼叫过的内容
    If blnAutoDelVoice = True Then
        Call DelVoiceData(lngVoiceId)
    End If
End Function


Public Function PlayQueueVoice(ByVal lngVoiceId As Long, ByVal lngQueueId As Long, ByVal blnAutoDelVoice As Boolean, _
                            Optional ByVal strVoiceContext As String = "") As Boolean
'播放指定队列的语音内容
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strCurVoice As String
    
    PlayQueueVoice = False
    
    strCurVoice = strVoiceContext
    If Trim(strVoiceContext) = "" Then
        strSql = "select 呼叫内容 from 排队语音呼叫  where 队列ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询呼叫内容", lngQueueId)
        
        If rsData.RecordCount <= 0 Then Exit Function
        
        strCurVoice = Nvl(rsData!呼叫内容)
    End If

    If GetQueueState(lngQueueId) = qsWaitCall Then
        '原来的流程:排队列表 执行状态更新
        strSql = "Zl_排队叫号队列_状态更新(" & lngQueueId & "," & qsCalling & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "更新呼叫状态")
    ElseIf GetQueueState(lngQueueId) = qsBeforeWaitCall Then
        '预呼的排队信息，不更新状态，继续后面的流程
    Else
        Exit Function
    End If
    
    '播放呼叫前的提示音
    If mblnIsPlayHintSound Then Call PlayHintSound
    
    '启用自定义呼叫后就不用语音呼叫
    If mblnUseVbsPlay = True Then
        If Trim(mstrCusVoiceScript) = "" Then
            Call DebugMsg("clsQueueOperation", "PlayQueueVoice", "mstrCusVoiceScript", "自定义呼叫脚本为空。")
            Exit Function
        End If
        
        Call DoScriptCall(lngVoiceId, strCurVoice)
    Else
        If StartCall(strCurVoice) = False Then
            Exit Function
        End If
    End If
    
    '更新呼叫状态为已呼叫,需要判断状态是否已经发生改变，如果改变，则不进行状态更新
    If GetQueueState(lngQueueId) = qsCalling Then
        strSql = "Zl_排队叫号队列_状态更新(" & lngQueueId & "," & qsCalled & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "更新呼叫状态")
    ElseIf GetQueueState(lngQueueId) = qsBeforeWaitCall Then
        strSql = "Zl_排队叫号队列_状态更新(" & lngQueueId & "," & qsAfterWaitCall & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "更新呼叫状态")
    End If
    
    '呼叫成功后删除呼叫过的内容
    If blnAutoDelVoice = True Then
        Call DelVoiceData(lngVoiceId)
    End If
    
    PlayQueueVoice = True
End Function

Public Sub DelVoiceData(ByVal lngVoiceId As Long)
'删除语音数据
    Dim strSql As String
    
    strSql = "ZL_排队语音呼叫_DELETE(" & lngVoiceId & ")"
    zlDatabase.ExecuteProcedure strSql, "语音呼叫完成"
End Sub

Public Sub ShellEx(ByVal strPathName As String, Optional ByVal lngWindowStyle As Long = 2)
    Shell strPathName, lngWindowStyle
End Sub

Public Function SaveFileEx(ByVal strContext As String, ByVal strFile As String) As String
    Dim lngFileNum As Long
    
    SaveFileEx = ""
    If Trim(strFile) = "" Then Exit Function
    
    lngFileNum = FreeFile()
    
    Open strFile For Output As lngFileNum
    
    Print #lngFileNum, strContext
    
    Close #lngFileNum

End Function


Private Sub DoScriptCall(ByVal lngCallId As Long, ByVal strCallContext As String)
'调用vbs脚本实现自定义呼叫
'lngCallId:待呼叫id
'strCallContext: 呼叫内容

    Dim lngIndex As Long
    Dim objCall As Object
    Dim strMethod As String
    Dim start As Long

    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
    End If

    '创建脚本执行对象
    Set objCall = CreateObject("ScriptControl")
    objCall.timeout = 60000

    strMethod = mstrCusVoiceScript
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

    Call objCall.AddObject("Me", Me, True)
    Call objCall.AddObject("SpVoice", mobjVoice, True)
    Call objCall.AddObject("zlDatabase", zlDatabase, True)
    Call objCall.AddObject("zlCommFun", zlCommFun, True)
    Call objCall.AddObject("zlControl", zlControl, True)

    Call objCall.AddCode(mstrCusVoiceScript)

    start = Timer
    
    Call objCall.Run(Trim(strMethod), lngCallId, strCallContext)
    

    Do While Timer < start + mlngPlayTimeLength
        Call Sleep(5)

        If mblnAbortState Then
            Exit Sub
        End If

        DoEvents
    Loop
End Sub

 
Public Function ChangeOrder(ByVal lngQueueId As Long, ByVal strNewOrder As String, _
    ByVal strMemo As String) As Boolean
'lngQueueID: 排队ID
'strNewOrder: 新的排队序号
'strMemo:备注，可保存为插队原因
'功能：调整排队顺序
    Dim strSql As String

    ChangeOrder = False
    
    strSql = "zl_排队叫号队列_序号更新(" & lngQueueId & ",'" & strNewOrder & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "更新排队序号")
    
    strSql = "zl_排队叫号队列_备注更新(" & lngQueueId & ",'" & strMemo & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "更新备注信息")

    ChangeOrder = True
End Function

Public Function ModifyMemo(ByVal lngQueueId As Long, ByVal strMemo As String) As String
'修改备注
    Dim strSql As String

    ModifyMemo = False
    
    strSql = "zl_排队叫号队列_备注更新(" & lngQueueId & ",'" & strMemo & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "更新备注信息")

    ModifyMemo = True
End Function

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
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function GetCustomOrderWhere() As String
'功能：获取自定义排序条件
    Dim strSql          As String
    Dim rsQueue         As ADODB.Recordset
    
    On Error GoTo ErrorHand
    
    GetCustomOrderWhere = ""
    
    strSql = "select zl_排队叫号队列_获取排序方式([1]) as 排序条件 from dual"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "获取自定义排序条件", mlngWorkType)
    
    If rsQueue.RecordCount > 0 Then GetCustomOrderWhere = Nvl(rsQueue!排序条件)
    
    Exit Function
ErrorHand:
    GetCustomOrderWhere = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function GetMinOrder(ByVal strQueueName As String) As String
'strQueueName: 队列名称
'功能：获取当前队列最小排队号码
    Dim strSql          As String
    Dim rsReocrd        As ADODB.Recordset
    
On Error GoTo ErrorHand
    GetMinOrder = ""
    
    
    strSql = "select 排队序号 from (select 排队序号 from 排队叫号队列 where 队列名称=[1] and 业务类型=[2] " & _
            IIf(Trim(mstrCustomOrder) <> "", " order by " & mstrCustomOrder, "") & " ) a where rownum =1"
    
    Set rsReocrd = zlDatabase.OpenSQLRecord(strSql, "获取当前队列最小排队号码", strQueueName, mlngWorkType)
    
    If rsReocrd.RecordCount > 0 Then GetMinOrder = Nvl(rsReocrd!排队序号)
    
    Exit Function
ErrorHand:
    GetMinOrder = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function GetQueueMaxNo(ByVal strQueueName As String, Optional ByVal dtQueueDate As Date) As String
'获取队列中的最多号码
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetQueueMaxNo = ""
    
    strSql = "select zl_排队叫号队列_获取排队号([1], [2],[3]) as 排队号码 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询排队号码", mlngWorkType, strQueueName, IIf(dtQueueDate = 0, Now, dtQueueDate))
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetQueueMaxNo = Nvl(rsData!排队号码)
End Function


Public Function GetQueueNoWithFilter(Optional ByVal strFilter As String) As String
'strFilter:根据自定义条件获取最大号码，默认为空表示当前业务下的最大排队号码,格式为/*"性别='男',年龄=10,..."*/
'如: '队列名称=超声科'，则获取队列名称为超声科的当前最大号码；‘排队标记=黑白超'，则获取排队标记为黑白超的当前最大号码
'功能：获取指定条件下的最大排队号码（排队号）
    Dim strSql          As String
    Dim strNewFilter    As String
    Dim varData(20)     As Variant
    Dim rsQueue         As ADODB.Recordset
   
On Error GoTo ErrorHand
    GetQueueNoWithFilter = ""
    
    If strFilter = "" Then
        strFilter = "业务类型=[1]"
        varData(0) = mlngWorkType
    Else
        Call GetFilterCfg(strFilter, strNewFilter, varData)
    End If
    
    strSql = "select max(to_number(排队号码)) as 排队号码 from 排队叫号队列 where " & strFilter
        
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "获取指定条件下的最大排队号码", varData(0), varData(1), varData(2), varData(3), _
                                                    varData(4), varData(5), varData(6), varData(7), varData(8), varData(9), _
                                                    varData(10), varData(11), varData(12), varData(13), varData(14), varData(15), _
                                                    varData(16), varData(17), varData(18), varData(19), varData(20))
    
    If rsQueue.RecordCount > 0 Then GetQueueNoWithFilter = Nvl(rsQueue!排队号码)
    
    Exit Function
ErrorHand:
    GetQueueNoWithFilter = ""
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Sub GetFilterCfg(ByVal strFilter As String, ByRef strNewFilter As String, ByRef varParament() As Variant)
'获取过滤条件配置
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
    
    '设置默认属性
    mintValidDays = 1
    mstrVoiceType = ""
    mblnIsPlayHintSound = False
    mlngPlayCount = 1
    mlngPlaySpeed = 10
    mlngPlayTimeLength = 10
    mstrCusVoiceScript = ""
    mblnUseVbsPlay = False
    mstrReportNum = ""
    mblnAbortState = False
End Sub


Private Sub Class_Terminate()
On Error Resume Next
    Call StopVoice
    Set mobjVoice = Nothing
    
    If Not gobjMsgCenter Is Nothing Then
        Call gobjMsgCenter.CloseMsgCenter
        
        Set gobjMsgCenter = Nothing
    End If
Err.Clear
End Sub


'清除指定天数前的数据,即数据有效天数
Public Property Get ValidDays() As Integer
    ValidDays = mintValidDays
End Property

Public Property Let ValidDays(ByVal value As Integer)
    mintValidDays = value
End Property

'呼叫方式
Public Property Get VoiceType() As String
    VoiceType = mstrVoiceType
End Property

Public Property Let VoiceType(ByVal value As String)
    mstrVoiceType = value
End Property

'呼叫次数
Public Property Get PlayCount() As Integer
    PlayCount = mlngPlayCount
End Property

Public Property Let PlayCount(ByVal value As Integer)
    mlngPlayCount = value
End Property

'语音播放速度
Public Property Get PlaySpeed() As Long
    PlaySpeed = mlngPlaySpeed
End Property

Public Property Let PlaySpeed(ByVal value As Long)
    mlngPlaySpeed = value
End Property

'呼叫播放时间长度
Public Property Get PlayTimeLength() As Long
    PlayTimeLength = mlngPlayTimeLength
End Property

Public Property Let PlayTimeLength(ByVal value As Long)
    mlngPlayTimeLength = value
End Property

'本机站点名
Public Property Get LocalStation() As String
    LocalStation = mstrLocalStation
End Property

Public Property Let LocalStation(ByVal value As String)
    mstrLocalStation = value
End Property

'待呼叫站点名
Public Property Get PlayStation() As String
    PlayStation = mstrPlayStation
End Property

Public Property Let PlayStation(ByVal value As String)
    mstrPlayStation = value
End Property

'呼叫前设置病人的检查目的地
Public Property Get CallTarget() As String
    CallTarget = mstrCallTarget
End Property

Public Property Let CallTarget(ByVal value As String)
    mstrCallTarget = value
End Property

'自定义排序
Public Property Get CustomOrder() As String
    CustomOrder = mstrCustomOrder
End Property

Public Property Let CustomOrder(ByVal value As String)
    mstrCustomOrder = value
End Property

'默认查询列
Public Property Get DefQueryCols() As String
    DefQueryCols = mstrDefQueryCols
End Property

'自定义语音播放脚本
Public Property Get CusVoiceScript() As String
    CusVoiceScript = mstrCusVoiceScript
End Property

Public Property Let CusVoiceScript(ByVal value As String)
    mstrCusVoiceScript = value
End Property

'使用vbScript脚本进行语音播放
Public Property Get UseVbsPlay() As Boolean
    UseVbsPlay = mblnUseVbsPlay
End Property

Public Property Let UseVbsPlay(ByVal value As Boolean)
    mblnUseVbsPlay = value
End Property


'报表编号
Public Property Get ReportNum() As String
    ReportNum = mstrReportNum
End Property

Public Property Let ReportNum(value As String)
    mstrReportNum = value
End Property

'呼叫前播放提示音
Public Property Get IsPlayHintSound() As Boolean
    IsPlayHintSound = mblnIsPlayHintSound
End Property

Public Property Let IsPlayHintSound(value As Boolean)
    mblnIsPlayHintSound = value
End Property







