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

'队列状态
Public Enum TQueueState
    qsPlace = -1        '占位中，当进入队列时，为占位状态
    qsQueueing = 0      '排队中，执行开始排队后，进入该状态
    qsCalling = 1       '呼叫中，当前正在呼叫的数据
    qsAbstain = 2       '已弃号，放弃呼叫的数据
    qsPause = 3         '已暂停，暂停呼叫的数据
    qsComplete = 4      '已完成，完成呼叫的数据
    qs_BroadCast = 5    '保留，广播
    qs_ReDiagnose = 6   '保留，回诊
    qsCalled = 7        '已呼叫，已执行呼叫但并非当前呼叫的数据
    qsDiagnose = 8      '接诊中
    qsWaitCall = 9      '待呼叫
End Enum

Private mobjComLib As Object                'zl9ComLib.clsComLib
Private mobjVoice As Object                 '语音呼叫对象

Private mstrVoiceType           As String   '微软语音类型名称
Private mblnIsPlayHintSound     As Boolean  '呼叫前播放提示音
Private mlngPlayCount           As Long     '呼叫次数 默认为1
Private mlngPlaySpeed           As Long     '语音播放速度
Private mlngPlayTimeLength      As Long     '呼叫时间长度，默认10秒
Private mstrPlayStation         As String   '播放语音的站点名,默认当前站点
Private mstrLocalStation        As String   '本机站点名
Private mlngWorkType            As Long     '业务类型
Private mstrCusVoiceScript      As String   '自定义语音播放脚本
Private mblnUseVbsPlay          As Boolean
Private mblnAbortState          As Boolean

'语音播放的函数
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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

Public Function GetQueueState(ByVal lngQueueId As Long) As Long
'lngQueueID: 排队ID
'-1-占位中，0-排队中，1-呼叫中，2-已弃号，3-已暂停，4-完成就诊,5-弃用，6-弃用，7-已呼叫，
'功能：获取排队状态，失败返回-2
On Error GoTo ErrorHand
    Dim strSql              As String
    Dim rsQueue             As ADODB.Recordset
        
    GetQueueState = -2
    
    strSql = "select 排队状态 from 排队叫号队列 where id=[1] and 业务类型=[2]"
    Set rsQueue = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "排队状态", lngQueueId, mlngWorkType)
    
    If rsQueue.RecordCount > 0 Then GetQueueState = Nvl(rsQueue!排队状态)
    
    Exit Function
ErrorHand:
    GetQueueState = -2
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub DebugMsg(ByVal strClassName, strMethod, strParameter As String, Optional strExceptionMsg As String = "")
    If False Then
        Call OutputDebugString(Now & ">> [调用过程：" & strClassName & "." & strMethod & "]  [参数内容：" & strParameter & "]  " & _
                                IIf(Trim(strExceptionMsg) <> "", "[异常信息：" & strExceptionMsg & "]", ""))
    End If
End Sub

Public Sub DelVoiceData(ByVal lngVoiceId As Long)
'删除语音数据
    Dim strSql As String
    
    strSql = "ZL_排队语音呼叫_DELETE(" & lngVoiceId & ")"
    mobjComLib.zlDatabase.ExecuteProcedure strSql, "语音呼叫完成"
End Sub

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
        Set rsData = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "查询呼叫内容", lngVoiceId)
        
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

Public Function PlayQueueVoice(ByVal objMsgCenter As Object, ByVal lngVoiceId As Long, ByVal lngQueueId As Long, ByVal blnAutoDelVoice As Boolean, _
                            Optional ByVal strVoiceContext As String = "") As Boolean
'播放指定队列的语音内容
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strCurVoice As String
    
    PlayQueueVoice = False
    
    strCurVoice = strVoiceContext
    If Trim(strVoiceContext) = "" Then
        strSql = "select 呼叫内容 from 排队语音呼叫  where 队列ID=[1]"
        Set rsData = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "查询呼叫内容", lngQueueId)
        
        If rsData.RecordCount <= 0 Then Exit Function
        
        strCurVoice = Nvl(rsData!呼叫内容)
    End If
        
    '如果该队列不处于待呼叫状态，则不进行呼叫处理
    If lngQueueId <= 0 Then Exit Function
    
    '更新呼叫状态为呼叫中
    If GetQueueState(lngQueueId) = qsWaitCall Then
        strSql = "Zl_排队叫号队列_状态更新(" & lngQueueId & "," & qsCalling & ")"
        Call mobjComLib.zlDatabase.ExecuteProcedure(strSql, "更新呼叫状态")
        
        If Not objMsgCenter Is Nothing Then Call objMsgCenter.SendMsg_StateSync(lngQueueId)
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
        Call mobjComLib.zlDatabase.ExecuteProcedure(strSql, "更新呼叫状态")
        
        If Not objMsgCenter Is Nothing Then Call objMsgCenter.SendMsg_StateSync(lngQueueId)
    End If
    
    '呼叫成功后删除呼叫过的内容
    If blnAutoDelVoice = True Then
        Call DelVoiceData(lngVoiceId)
    End If
    
    PlayQueueVoice = True
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
'    Call objCall.AddObject("zlDatabase", zlDatabase, True)
'    Call objCall.AddObject("zlCommFun", zlCommFun, True)
'    Call objCall.AddObject("zlControl", zlControl, True)

    Call objCall.AddCode(mstrCusVoiceScript)

    start = Timer
    
    Call objCall.Run(Trim(strMethod), lngCallId, strCallContext)

    Do While Timer < start + mlngPlayTimeLength
        Call Sleep(5)

        If mblnAbortState Then Exit Sub

        DoEvents
    Loop
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
    mstrVoiceType = ""
    mblnIsPlayHintSound = False
    mlngPlayCount = 1
    mlngPlaySpeed = 10
    mlngPlayTimeLength = 10
    mstrCusVoiceScript = ""
    mblnUseVbsPlay = False
    mblnAbortState = False
End Sub

Public Sub setComLib(objComLib As Object)
    Set mobjComLib = objComLib
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Call StopVoice
    Set mobjVoice = Nothing
    Set mobjComLib = Nothing
Err.Clear
End Sub

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

'呼叫前播放提示音
Public Property Get IsPlayHintSound() As Boolean
    IsPlayHintSound = mblnIsPlayHintSound
End Property

Public Property Let IsPlayHintSound(value As Boolean)
    mblnIsPlayHintSound = value
End Property

'业务类型
Public Property Get BusinessType() As Long
    BusinessType = mlngWorkType
End Property

Public Property Let BusinessType(value As Long)
    mlngWorkType = value
End Property
