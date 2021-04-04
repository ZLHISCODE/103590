VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mint场合 As Integer '0-门诊医生工作站，1－住院医生工作站，2－住院护士工作站，3－老版医技工作站
Private mstrCurDate As String '当前时刻
 
Private mobjVBA As Object
Private mobjScript As zlPublicAdvice.clsScript
Private mobjVoice As Object                  '语音呼叫对象
Private mrsPars As ADODB.Recordset '参数信息
Private mrsMsgNo As ADODB.Recordset '消息类型和编号对应

Public Sub InitCISMsg(ByVal intType As Integer)
'功能：初始化参数，
    Dim strTmp As String
    Dim varTmp As Variant
    
    Dim strNos As String
    Dim varNo As Variant
    
    Dim i As Long, lng模块 As Long
    
    mint场合 = intType
    
    Call CreateScript(mobjVBA, mobjScript)
    
    Call InitRsMsgPar(mrsPars)
    
    Set mrsMsgNo = New ADODB.Recordset
    mrsMsgNo.Fields.Append "声音类型", adVarChar, 800
    mrsMsgNo.Fields.Append "消息编号", adVarChar, 60
    mrsMsgNo.CursorLocation = adUseClient
    mrsMsgNo.LockType = adLockOptimistic
    mrsMsgNo.CursorType = adOpenStatic
    mrsMsgNo.Open
    
    If mint场合 = 0 Then
        lng模块 = p门诊医生站
    ElseIf mint场合 = 1 Then
        lng模块 = p住院医生站
    ElseIf mint场合 = 2 Then
        lng模块 = p住院护士站
    ElseIf mint场合 = 3 Then
        lng模块 = p医技工作站
    End If
     
    strTmp = Get消息类别(intType)
    varTmp = Split(strTmp, ",")
    strNos = Get消息编号(intType)
    varNo = Split(strNos, "|")
    For i = 0 To UBound(varTmp)
        Call AddDataToRsMsgPar(mrsPars, lng模块, i + 1, varTmp(i) & "语音配置", varTmp(i))
        Call AppendNoPar(varTmp(i), varNo(i))
    Next
    
    mstrCurDate = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
End Sub
 
Private Sub AppendNoPar(ByVal str消息名 As String, ByVal strNos As String)
'功能：消息与消息编号关联
    Dim varTmp As Variant
    Dim i As Long
    varTmp = Split(strNos, ",")
    For i = 0 To UBound(varTmp)
        mrsMsgNo.AddNew
        mrsMsgNo!声音类型 = str消息名
        mrsMsgNo!消息编号 = varTmp(i)
        mrsMsgNo.Update
    Next
End Sub

Public Sub PlayMsgSound(ByRef rsMsg As ADODB.Recordset)
'播报消息
    Dim strBegin As String, strEnd As String
    Dim rs播放 As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    strBegin = mstrCurDate
    mstrCurDate = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    strEnd = mstrCurDate
    
    If rsMsg.RecordCount > 0 Then
        Set rs播放 = zlDatabase.CopyNewRec(rsMsg, , , Array("是否播放", adInteger, 10, 0))
        For i = 1 To rs播放.RecordCount
            strTmp = Format(rs播放!登记时间, "yyyy-MM-dd HH:mm:ss")
            If Between(strTmp, strBegin, strEnd) Then
                rs播放!是否播放 = 1
            End If
            rs播放.MoveNext
        Next
        rs播放.Filter = "是否播放 = 1"
        If Not rs播放.EOF Then
            For i = 1 To rs播放.RecordCount
                Call PlaySound(rs播放)
                rs播放.MoveNext
            Next
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PlaySound(ByRef rsMsg As ADODB.Recordset)
'功能：护士站按指定配置播消息
    Dim strText As String, strFile As String
    Dim int提示方式 As Integer, i As Long
    Dim start As Date
    Dim str内容 As String
    
    On Error Resume Next
 
    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
    End If
   
    mrsMsgNo.Filter = "消息编号='" & rsMsg!类型编码 & "'"
    If mrsMsgNo.EOF Then Exit Sub
    mrsPars.Filter = "声音类型='" & mrsMsgNo!声音类型 & "'"
    
    If Val(mrsPars!状态 & "") = 0 Then Exit Sub
    
    str内容 = mrsPars!内容 & ""
    int提示方式 = Val(mrsPars!提示方式 & "")
    If int提示方式 = 0 Then
        strText = Get播放文本(rsMsg, str内容)  '文本
    Else
        strFile = str内容 '文件
    End If
     
    For i = 1 To Val(mrsPars!次数 & "")
        If int提示方式 = 0 Then
            mobjVoice.Speak strText, 1
        Else
            Call sndPlaySound(strFile, 1) '文件丢失损坏不会报错不用特殊检查
        End If
        start = Timer
        Do While Timer < start + 5
            Call OS.Wait(5)
            DoEvents
        Loop
    Next
    err.Clear
End Sub

Private Function Get播放文本(ByRef rsMsg As ADODB.Recordset, ByVal strText As String) As String
'功能：获取播入的文本
    Dim str床号 As String, str住院号 As String
    Dim strVal As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str项目内容 As String
    
    On Error GoTo errH
    str床号 = IIf("" = rsMsg!床号 & "", "", rsMsg!床号 & "床")
    str住院号 = IIf("" = rsMsg!住院号 & "", "", rsMsg!住院号 & "号")
    strVal = strText
    strVal = Replace(strVal, "[床号]", """" & str床号 & """")
    strVal = Replace(strVal, "[住院号]", """" & str住院号 & """")
    
    If InStr("ZLHIS_PACS_006,ZLHIS_PACS_007", rsMsg!类型编码 & "") > 0 Then
        strSQL = "select 医嘱内容 as 内容 from 病人医嘱记录 where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "消息语音", Val(rsMsg!业务标识 & ""))
        If Not rsTmp.EOF Then str项目内容 = rsTmp!内容 & ""
    End If
    strVal = Replace(strVal, "[项目内容]", """" & str项目内容 & """")
    
    strVal = mobjVBA.Eval(strVal)
    If mobjVBA.Error.Number <> 0 Then
        err.Clear
        strVal = IIf(str床号 = "", "家庭床", str床号) & rsMsg!消息内容
    End If
    Get播放文本 = strVal
    Exit Function
errH:
    err.Clear
    Get播放文本 = "病人有  " & rsMsg!消息内容
End Function

Private Sub Class_Terminate()
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mobjVoice = Nothing
    Set mrsPars = Nothing
    Set mrsMsgNo = Nothing
End Sub

'-------------------公用方法
Public Sub InitRsMsgPar(ByRef rsPar As ADODB.Recordset)
'功能：初始化参数记录集
    Set rsPar = New ADODB.Recordset
    
    rsPar.Fields.Append "序号", adBigInt
    rsPar.Fields.Append "模块", adBigInt
    rsPar.Fields.Append "参数名", adVarChar, 1000
    rsPar.Fields.Append "声音类型", adVarChar, 800
    rsPar.Fields.Append "状态", adInteger '1-开启，0－关闭
    rsPar.Fields.Append "提示方式", adInteger '1－音频文件，0－朗读文本
    rsPar.Fields.Append "内容", adVarChar, 3800
    rsPar.Fields.Append "次数", adInteger
    rsPar.Fields.Append "修改", adInteger '1-被修改，0－未修改
    
    rsPar.Fields.Append "原参数值", adVarChar, 4000
    rsPar.Fields.Append "现参数值", adVarChar, 4000
    
    rsPar.CursorLocation = adUseClient
    rsPar.LockType = adLockOptimistic
    rsPar.CursorType = adOpenStatic
    rsPar.Open
End Sub

Public Sub AddDataToRsMsgPar(ByRef rsPar As ADODB.Recordset, ByVal lngModual As Long, ByVal lng序号 As Long, ByVal str参数名 As String, ByVal str消息名 As String)
'功能：向本地记录集中添加参数数据
    Dim varTmp As Variant
    Dim strTmp As String
    
    On Error GoTo errH
    strTmp = zlDatabase.GetPara(str参数名, glngSys, lngModual, "1<sTab>0<sTab>iif([床号]<>"""",[床号],""家庭床"")+""有" & str消息名 & "。""<sTab>2")
    varTmp = Split(strTmp, "<sTab>")
    rsPar.AddNew
    rsPar!序号 = lng序号
    rsPar!模块 = lngModual
    rsPar!参数名 = str参数名
    rsPar!声音类型 = str消息名
    rsPar!状态 = Val(varTmp(0))
    rsPar!提示方式 = Val(varTmp(1))
    rsPar!内容 = varTmp(2)
    rsPar!次数 = Val(varTmp(3))
    rsPar!原参数值 = strTmp
    rsPar!现参数值 = strTmp
    rsPar!修改 = 0
    rsPar.Update
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CreateScript(Optional ByRef objVBA As Object, Optional ByRef objScript As clsScript) As Boolean
'功能：创建Script和VBA对象
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    err.Clear: On Error GoTo 0
    If Not objVBA Is Nothing Then
        objVBA.Language = "VBScript"
        Set objScript = New clsScript
        objVBA.addObject "clsScript", objScript, True
        CreateScript = True
    End If
End Function

Public Function Get消息编号(ByVal intType As Integer) As String
'功能：获取每类消息对应的编号
'参数：intType '0-门诊医生工作站，1－住院医生工作站，2－住院护士工作站，3－老版医技工作站
    If intType = 0 Then
        Get消息编号 = "ZLHIS_LIS_003,ZLHIS_PACS_005|ZLHIS_OPER_001,ZLHIS_CIS_015,ZLHIS_CIS_005|ZLHIS_RECIPEAUDIT_001|ZLHIS_CIS_032,ZLHIS_CIS_033"
    ElseIf intType = 1 Then
        Get消息编号 = "ZLHIS_EMR_021|ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015,ZLHIS_CIS_020|ZLHIS_LIS_003,ZLHIS_PACS_005|ZLHIS_LIS_002,ZLHIS_PACS_003|ZLHIS_CIS_026,ZLHIS_CIS_027,ZLHIS_CIS_028,ZLHIS_CIS_029,ZLHIS_CIS_030|ZLHIS_RECIPEAUDIT_002|ZLHIS_CIS_032,ZLHIS_CIS_033|ZLHIS_EMR_025"
    ElseIf intType = 2 Then
        Get消息编号 = "ZLHIS_CIS_001|ZLHIS_CIS_002|ZLHIS_CIS_003|ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015|ZLHIS_LIS_003,ZLHIS_PACS_005|ZLHIS_TRANSFUSION_001|ZLHIS_CHARGE_001|ZLHIS_PACS_006|ZLHIS_PACS_007"
    ElseIf intType = 3 Then
        Get消息编号 = "ZLHIS_CHARGE_001|ZLHIS_CIS_004"
    End If
End Function

Public Function Get消息类别(ByVal intType As Integer) As String
'功能：获取工作站的消息类别
'参数：intType  0-门诊医生工作站，1－住院医生工作站，2－住院护士工作站，3－老版医技工作站
    If intType = 0 Then
        Get消息类别 = "危机值消息,安排消息,处方审查消息,传染病消息"
    ElseIf intType = 1 Then
        Get消息类别 = "病历审阅消息,安排消息,危机值消息,报告撤消消息,医嘱审核消息,处方审查消息,传染病消息,病历质控消息"
    ElseIf intType = 2 Then
        Get消息类别 = "新开消息,新停消息,新废消息,安排消息,危机值消息,输液拒绝消息,销帐申请消息,RIS预约消息,RIS预约准备消息"
    ElseIf intType = 3 Then
        Get消息类别 = "销帐申请消息,待安排消息"
    End If
End Function
