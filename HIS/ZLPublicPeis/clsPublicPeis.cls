VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicPeis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'用途：体检外部接口
'方法：Initialize,WriteLisResult,GetLastError
'######################################################################################################################

Private mstrError As String
Private mobjPeisPDFReport As clsPeisPDFReport
Private mobjTaskPersonReport As frmTaskPersonReport

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'属性：PeisPDFReport
Public Property Get PeisPDFReport() As clsPeisPDFReport
    If mobjPeisPDFReport Is Nothing Then
        Set mobjPeisPDFReport = New clsPeisPDFReport
        Call mobjPeisPDFReport.Initialize(2100, 2121)
    End If
    Set PeisPDFReport = mobjPeisPDFReport
End Property

Public Function Initialize(ByVal cnOracle As ADODB.Connection) As Boolean
    '******************************************************************************************************************
    '功能：部件的初始化
    '参数：cnOracle:主程序的数据库连接
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    ParamInfo.系统名称 = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    ParamInfo.产品名称 = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    mstrError = ""
    Set gobjComLib = Nothing
    Set gobjDatabase = Nothing
    
    Set gcnOracle = cnOracle
    
    On Error Resume Next
    Err = 0
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    On Error GoTo errHand
    
    If Err <> 0 Then
        mstrError = "创建公共部件(zl9ComLib.clsComLib)失败。"
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    Set gobjComFun = CreateObject("zl9ComLib.clsCommFun")
    On Error GoTo errHand
    If Err <> 0 Then
        mstrError = "创建公共部件(zl9ComLib.clsCommFun)失败。"
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    Set gobjDatabase = CreateObject("zl9ComLib.clsDatabase")
    On Error GoTo errHand
    
    If Err <> 0 Then
        mstrError = "创建公共部件(zl9ComLib.clsDatabase)失败。"
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    Set gobjReport = CreateObject("zl9Report.clsReport")
    On Error GoTo errHand
    
    If Err <> 0 Then
        mstrError = "创建公共部件(zl9Report.clsReport)失败。"
        Exit Function
    End If
    
    
    Call gobjComLib.InitCommon(cnOracle)
    
    Initialize = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

'******************************************************************************************************************
'功能：传递一个病人id给接口方法,根据这个病人进行判断是否有报告
'参数：lngPatientKey 病人id
'返回：true--该病人有报告  false--该病人没有报告
'******************************************************************************************************************
Public Function HasExaminationReport(ByVal lngPatientKey As Long) As Boolean
    Dim rsData As ADODB.Recordset
    Dim strPrivs As String
    
    On Error GoTo errHandle
    
    '判断当前操作员是否有报告查阅权限
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "体检报告查阅") Then
       HasExaminationReport = False
       Exit Function
    End If
    Set rsData = gclsPackage.Get_PeisPersonReportBill(lngPatientKey)
    If rsData.BOF = False Then
        HasExaminationReport = True
        Exit Function
    End If
    HasExaminationReport = False
    
    Exit Function
errHandle:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

'******************************************************************************************************************
'功能：病人报告查阅
'参数：frmParent 窗体对象
'      lngPatientKey 病人id
'返回：
'******************************************************************************************************************
Public Function OpenExaminationReport(ByVal frmParent As Object, ByVal lngPatientKey As Long) As Boolean
    Dim objPersonReport As Object
    Dim strPrivs As String
    
    '判断当前操作员是否有报告查阅权限
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "体检报告查阅") Then
        
        Call ShowSimpleMsg("您没有体检报告查阅的权限,不能进行报告查阅。")
        OpenExaminationReport = False
        Exit Function
    End If
    
    Set objPersonReport = New frmPersonReport
    
    OpenExaminationReport = objPersonReport.ShowReportForm(frmParent, lngPatientKey)
End Function

'******************************************************************************************************************
'功能：获取病人任务记录集
'参数：lngPatientKey 病人id
'返回：
'******************************************************************************************************************
Public Function GetTaskExaminationReportRecord(ByVal lngPatientKey As Long) As ADODB.Recordset
    Dim objPersonReport As Object
    Dim strPrivs As String
    
    '判断当前操作员是否有报告查阅权限
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "体检报告查阅") Then
        
        Call ShowSimpleMsg("您没有体检报告查阅的权限,不能进行报告查阅。")
        Set GetTaskExaminationReportRecord = Nothing
        Exit Function
    End If
    Set GetTaskExaminationReportRecord = gclsPackage.Get_PeisPersonReportBill(lngPatientKey)
    
End Function


'******************************************************************************************************************
'功能：获取任务人员报告窗体对象
'参数：
'返回：窗体对象
'******************************************************************************************************************
Public Function GetTaskPersonExaminationReportForm() As Object
    Dim strPrivs As String
    
    '判断当前操作员是否有报告查阅权限
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "体检报告查阅") Then
        
        Call ShowSimpleMsg("您没有体检报告查阅的权限,不能进行报告查阅。")
        Set GetTaskPersonExaminationReportForm = Nothing
        Exit Function
    End If
    
    If mobjTaskPersonReport Is Nothing Then
       Set mobjTaskPersonReport = New frmTaskPersonReport
       Call mobjTaskPersonReport.InitData
    End If
    Set GetTaskPersonExaminationReportForm = mobjTaskPersonReport
    
End Function

'******************************************************************************************************************
'功能：加载任务人员报告信息
'参数：lngTaskKey 任务ID
'      lngPatientKey 病人id
'返回：
'******************************************************************************************************************
Public Function LoadTaskPersonExaminationReportDetail(ByVal lngTaskKey As Long, ByVal lngPatientKey As Long) As Boolean
    Dim strPrivs As String
    
    '判断当前操作员是否有报告查阅权限
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "体检报告查阅") Then
        
        Call ShowSimpleMsg("您没有体检报告查阅的权限,不能进行报告查阅。")
        LoadTaskPersonExaminationReportDetail = False
        Exit Function
    End If
    
    If mobjTaskPersonReport Is Nothing Then
       Set mobjTaskPersonReport = New frmTaskPersonReport
       Call mobjTaskPersonReport.InitData
    End If
    
    LoadTaskPersonExaminationReportDetail = mobjTaskPersonReport.LoadData(lngTaskKey, lngPatientKey)
    
End Function
    
Public Function WriteLisResult(ByVal strVal As String, ByVal intType As Integer, Optional ByRef strErr As String) As Boolean
    '******************************************************************************************************************
    '功能：把检验结果回写入体检中
    '参数           strVal 专入的标本结果内容
    '               类型(1=普通)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2>
    '                   指标1<split4>检验结果1<split4>单位1<split4>结果标志1<split4>结果参数1<split4>排列序号1<split4>隐私项目1<split4>指标代码1<split3>
    '                   指标2<split4>检验结果2<split4>单位2<split4>结果标志2<split4>结果参数2<split4>排列序号2<split4>隐私项目2<split4>指标代码2<split3>
    '                   指标3<split4>检验结果3<split4>单位3<split4>结果标志3<split4>结果参数3<split4>排列序号3<split4>隐私项目3<split4>指标代码3<split1>
    '
    '               类型(2=微生物)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2>
    '               细菌名1<split3>描述1<split3>耐药机制1<split3>
    '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
    '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split2>
    '               细菌名2<split3>描述2<split3>耐药机制2<split3>
    '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
    '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split1>
    '
    '               intType 0=审核 1=取消审核
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Dim intSample As Integer                    '循环变量
    Dim astrSample() As String                  '循环标本数组
    Dim strSQL As String
    Dim lngState As Long '获取提醒状态
    Dim lng医嘱id As Long
    Dim strSQLtmp As String
    Dim rs As ADODB.Recordset
    Dim varMsg As VbMsgBoxResult
    Dim blnBegin As Boolean

    If intType = 1 Then
        lngState = Val(gobjDatabase.GetPara("LIS取消审核提醒", 2100, 0, "0"))
    End If
    strSQLtmp = "Select Sum(个数) As 个数" & vbNewLine & _
                "From (Select Count(1) As 个数" & vbNewLine & _
                "       From 体检任务人员 A, 体检任务发送 B" & vbNewLine & _
                "       Where a.任务id = b.任务id And a.病人id = b.病人id And a.体检状态 = 3 And b.采集医嘱id = [1] And Rownum < 2" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Count(1) As 个数" & vbNewLine & _
                "       From 体检任务结论 A, 体检任务发送 B" & vbNewLine & _
                "       Where a.任务id = b.任务id And a.病人id = b.病人id And a.清单id = b.清单id And a.体检项目id = b.体检项目id And a.结论描述 Is Not Null And" & vbNewLine & _
                "             b.采集医嘱id = [1] And Rownum < 2" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Count(1) As 个数" & vbNewLine & _
                "       From 体检任务结论 A, 体检任务发送 B" & vbNewLine & _
                "       Where a.Id = b.任务id And b.采集医嘱id = [1] And Rownum < 2)"
    astrSample = Split(strVal, "<split1>")
    For intSample = 0 To UBound(astrSample)
        If Trim(astrSample(intSample)) <> "" Then
            If intType = 1 Then
                lng医嘱id = Val(Split(astrSample(intSample), "<split2>")(1))
                Set rs = gobjDatabase.OpenSQLRecord(strSQLtmp, "提取病人审核信息", lng医嘱id)
                If rs.BOF = False And blnBegin = False Then
                    If IIf(IsNull(rs("个数").Value), 0, rs("个数").Value) Then
                        Select Case lngState
                            Case 0 '提示
                                varMsg = MsgBox("当前人员处于正在总检或者已总检状态，或者当前人员已经体检完成，也有可能当前人员所对应的检验项目已填写了项目小结，你确定要取消报告审核？" & vbCrLf & _
                                              "(注:取消报告审核，则可能导致项目小结与项目结果不匹配！也有可能导致已填写的总检结论不准确！)", vbYesNo + vbQuestion + vbDefaultButton2)
                                If varMsg = vbNo Then
                                    Exit Function
                                Else
                                    blnBegin = True
                                End If
                            Case 1 '禁止
                                MsgBox ("当前人员正在总检，或者已总检，或者当前人员已经体检完成，或者该项目已经作了小结。无法取消报告审核！" & vbCrLf & _
                                        "(提示：确实需要取消报告审核，则需要进行如下操作:" & vbCrLf & _
                                        "       请检查当前人员的体检状态是否为体检完成,如果是,则取消体检完成;" & vbCrLf & _
                                        "       请检查当前人员是否正在总检或者已总检,如果是,则需要将当前人员恢复到待总检或体检中;" & vbCrLf & _
                                        "       请检查当前人员所对应的检验项目是否填写了相应的项目小结,如果有,则进行清空)")
                                Exit Function
                        End Select
                    End If
                End If
            End If
            strSQL = "zl_体检任务结果_FROMLISTOTJ('" & astrSample(intSample) & "'," & intType & ")"
           Call gobjDatabase.ExecuteProcedure(strSQL, "保存")
        End If
    Next
    WriteLisResult = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    strErr = Err.Description
    mstrError = Err.Description
End Function

Public Function GetLastError() As String
    '******************************************************************************************************************
    '功能：获取错误信息
    '参数：
    '返回：
    '******************************************************************************************************************
        
    GetLastError = mstrError

End Function
'
Private Sub Class_Initialize()
    mstrError = ""
End Sub

Private Sub Class_Terminate()
     If Not (gcnOracle Is Nothing) Then Set gcnOracle = Nothing
     If Not (gobjComLib Is Nothing) Then Set gobjComLib = Nothing
     If Not (gobjComFun Is Nothing) Then Set gobjComFun = Nothing
     If Not (gobjDatabase Is Nothing) Then Set gobjDatabase = Nothing
     If Not (gclsPackage Is Nothing) Then Set gclsPackage = Nothing
     If Not (gobjReport Is Nothing) Then Set gobjReport = Nothing
     
     Set mobjPeisPDFReport = Nothing
     Set mobjTaskPersonReport = Nothing
End Sub


