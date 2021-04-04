VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicHisCommLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Version As String
Private gobjFile As New FileSystemObject '调试日志(问题号99437)-XT 2016-08-17 10:18
Private objTxt As TextStream
Private mobjITranLib As Object
Private mblnITranLib As Boolean

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    mblnITranLib = False
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
    If Not gobjHisComLib Is Nothing Then
        Set gobjHisComLib = Nothing
        Set gobjHisDatabase = Nothing
    End If
    
    CloseWindows = (Forms.Count = 0)

End Function
Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
      '功能： 根据主程序指定功能，调用执行相关程序
      '参数：
      '   lngModul:需要执行的功能序号
      '   cnMain:主程序的数据库连接
      '   frmMain:主窗体
      '   strDBUser:当前登录数据库用户名

          '------------------------------------------------
          Dim strErr As String
          Dim objLabWork As Object
          
1         On Error GoTo CodeMan_Error
          
2         If InitComponentsHIS(gSysInfo.SysNo, gSysInfo.ModlNo, cnMain, strErr) = False Then
3             MsgBox "初始LIS接口失败!" & vbCrLf & strErr, vbInformation, "LIS接口初始化"
4             Exit Sub
5         End If
          
6         Set gobjEmr = frmMain.mobjEMR
          
7         Select Case lngModul
              Case 1285   '临川调用阳性报告打印模块
8                 Call PrintDiseaseReport(strDBUser)
9             Case 1727   '卫材科调用检验试剂管理
10                Set objLabWork = CreateObject("zl9LabWork.clsLabWork")
11                Call objLabWork.CodeMan(2500, 1014, gcnLisOracle, frmMain, strDBUser)
12        End Select
          
13        Exit Sub
CodeMan_Error:
14        strErr = Err.Number & " " & Err.Description
15        Call WriteErrLog("zlPublicHisCommLis", "clsLisInsideComm", "执行(CodeMan)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
16        Err.Clear
End Sub


Public Function GetHISConn() As Connection
    Set GetHISConn = gcnHisOracle
End Function

Public Function CheckLisSate() As Boolean
    '功能                   检查LIS是否在使用状态
    '返回                   True=使用   False=未使用
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CheckLisSate = IIf(intLis_Setup = 1, True, False)
End Function

Public Function InitComponentsHIS(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnMainHIS As ADODB.Connection, Optional strErr As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能   初始化接口部件(HIS专用)
    '参数   lngSys          系统号
    '       lngModul        模块号
    '       cnMainHIS       HIS用户传入的连接
    '返回                   True=成功 False=失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If InitComponents(lngSys, lngModul, cnMainHIS, strErr) = False Then
        InitComponentsHIS = False
        Exit Function
    End If
    InitComponentsHIS = True
End Function
Public Function InitComponentsTJ(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnMainTJ As ADODB.Connection, Optional strErr As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能   初始化接口部件(体检专用)
    '参数   lngSys          系统号
    '       lngModul        模块号
    '       cnMainHIS       体检用户传入的连接
    '返回                   True=成功 False=失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If InitComponents(lngSys, lngModul, cnMainTJ, strErr) = False Then
        InitComponentsTJ = False
        Exit Function
    End If
    InitComponentsTJ = True
End Function
Public Function InitComponentsXK(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnMainXK As ADODB.Connection, Optional strErr As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能   初始化接口部件(血库专用)
    '参数   lngSys          系统号
    '       lngModul        模块号
    '       cnMainHIS       血库用户传入的连接
    '返回                   True=成功 False=失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If InitComponents(lngSys, lngModul, cnMainXK, strErr) = False Then
        InitComponentsXK = False
        Exit Function
    End If
    InitComponentsXK = True
End Function

Public Function InitComponentsEMR(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnMainEMR As ADODB.Connection, Optional strErr As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能   初始化接口部件(新版电子病历专用)
    '参数   lngSys          系统号
    '       lngModul        模块号
    '       cnMainEMR       新版电子病历用户传入的连接
    '返回                   True=成功 False=失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If InitComponents(lngSys, lngModul, cnMainEMR, strErr) = False Then
        InitComponentsEMR = False
        Exit Function
    End If
    InitComponentsEMR = True
End Function


Public Function InitComponents(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnMainHIS As ADODB.Connection, Optional strErr As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能   初始化接口部件
    '参数   lngSys          系统号
    '       lngModul        模块号
    '       intType         1=lis 2=his 3=体验 4=血库 5=新版电子病历
    '       cnMainLIS       LIS系统传入的连接
    '       cnMainHIS       HIS用户传入的连接
    '       cnMaintj        体检用户传入的连接
    '       cnMainxk        血库用户传入的连接
    '       cnMainEMR       新版电子病历传人的连接
    '       strErr          失败时返回的复杂信息
    '返回                   True=成功 False=失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '初始化ComLib部件
    If Not cnMainHIS Is Nothing Then Set gcnHisOracle = cnMainHIS
    
    '初始化ComLib
    If Not ComInitComLib(strErr) Then
        MsgBox strErr, vbQuestion, "错误"
        Exit Function
    End If
    
    '连接数据库
    If InitDBConn(gcnHisOracle, strErr) = False Then
        If strErr <> "" Then
            MsgBox strErr
        End If
        Exit Function
    End If
        
     '取用户信息
    If Not ComGetUserInfo(strErr) Then
        Exit Function
    End If
    
    '读取系统参数
    If ComGetSysParameter(strErr) = False Then
        Exit Function
    End If
    
    '初始化系统号与模块号
    gSysInfo.SysNo = 2500
    gSysInfo.ModlNo = 2500
    
    '模板权限
    gSysParameter.Privs = ComGetPrivs(Sel_His_DB, lngSys, lngModul)
    
    InitComponents = True

End Function


Public Function SendLisApplicationForm(strAdvices As String, strDiagnose As String, Optional strErr As String) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   发送检验申请单到LIS系统中
    '参数                   strAdvices 医嘱内容,格式
    '                       格式：<检验医嘱1,采集医嘱1,执行科室编码1,标本1;检验医嘱2,采集医嘱2,执行科室编码2,标本2;.....>
    '                       strDiagnose 诊断信息
    '                       strErr 有错误信息时返回错误信息
    '返回                   TRUE=发送成功 FALSE=发送失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SendLisApplicationForm = SendLisApplication(strAdvices, strDiagnose, strErr)
End Function
Public Function DelLisApplicationForm(strAdvices As String, Optional strErr As String) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   删除LIS中对应的请单
    '参数                   strAdvices 医嘱内容,格式
    '                       格式：<采集医嘱1,采集医嘱2,.....>
    '                       strErr 有错误信息时返回错误信息
    '返回                   TRUE=发送成功 FALSE=发送失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    DelLisApplicationForm = DelLisApplication(strAdvices, strErr)
End Function

Public Function SampleBarcodeWrite(strAdvices As String, strBarCode As String, Optional strSamplingName As String, Optional strErr As String, Optional ByVal intContinue As Integer) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   在采样时完成时写入样本条码到申请记录中，取消写入空的条码信息
    '
    '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
    '                       格式:“医嘱1,医嘱2,医嘱3,..
    '                       strBarCode   要写入的条码
    '                       strSamplingName      采样人
    '                       intContinue 1=让步检验，2=取消检验，3=重采样本
    '返回                   如果返回False时可在strErr中显示具体错误
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '回写取消审核标志
    If SampleBarcodeUpdate(strAdvices, strBarCode, strSamplingName, strErr, intContinue) = True Then
        SampleBarcodeWrite = True
    End If
    
    
End Function


Public Function GetPatientSampleValue(lngPatientID As Long) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   根据病人ID提取检验标本结果
    '参数                   lngPatientID   病人ID
    '返回
    '               类型(1=普通)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2> 婴儿序号 <split2>
    '                   指标1<split4>检验结果1<split4>单位1<split4>结果标志1<split4>结果参数1<split4>排列序号1<split4>隐私项目1<split4>指标代码1<split3>
    '                   指标2<split4>检验结果2<split4>单位2<split4>结果标志2<split4>结果参数2<split4>排列序号2<split4>隐私项目2<split4>指标代码2<split3>
    '                   指标3<split4>检验结果3<split4>单位3<split4>结果标志3<split4>结果参数3<split4>排列序号3<split4>隐私项目3<split4>指标代码3<split1>
    '
    '               类型(2=微生物)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2> 婴儿序号 <split2>
    '               细菌名1<split3>描述1<split3>耐药机制1<split3>
    '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
    '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split2>
    '               细菌名2<split3>描述2<split3>耐药机制2<split3>
    '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
    '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split1>
    '
    '               查询:       Set rsTmp = ComOpenSQL(Sel_XK_DB, strSQL, "读取结果")
    '               执行过程：  ComExecuteProc Sel_XK_DB, strSQL, "保存"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GetPatientSampleValue = CreateLisValueStr(0, lngPatientID)
End Function

Public Function BeforGetSampleValue(ByVal strAdvices As String, Optional ByVal DateE As Date, Optional strErr As String) As Boolean
    '传入医嘱ID和时间,判断传入时间之后是否存在医嘱ID对应的已审核记录
    'stradvices     '医嘱ID,多个医嘱使用","号分隔
    'DateE          '获取在该时间之后是否存在医嘱ID对应的记录
    
    '返回       True=存在记录,False=不存在记录
    
    BeforGetSampleValue = BeforCreateLisValueStr(strAdvices, DateE, strErr)
End Function

Public Function GetSampleValue(strAdvices As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   提取指定标本的结果
    '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
    '返回
    '               1;8453863;2;2013-03-26 09:57:38;管理员;管理员;2013-03-26 09:58:26;输血常规;枸橼酸钠抗凝;
    '                   RH血型鉴定(RH)^+^^^^21^0^RH,
    '                   ABO血型鉴定(ABO)^A^^^^19^0^ABO
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
    '               查询:       Set rsTmp = ComOpenSQL(Sel_XK_DB, strSQL, "读取结果")
    '               执行过程：  ComExecuteProc Sel_XK_DB, strSQL, "保存"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    GetSampleValue = CreateLisValueStr(strAdvices)
    Call CreateLogFile("GetSampleValue=" & GetSampleValue)
End Function

Public Function GetSampleValueForTJ(strAdvices As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   提取指定标本的结果(为体检专门提供)
    '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
    '返回
    '               1;8453863;2;2013-03-26 09:57:38;管理员;管理员;2013-03-26 09:58:26;输血常规;枸橼酸钠抗凝;
    '                   RH血型鉴定(RH)^+^^^^21^0^RH,
    '                   ABO血型鉴定(ABO)^A^^^^19^0^ABO
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
    '               查询:       Set rsTmp = ComOpenSQL(Sel_XK_DB, strSQL, "读取结果")
    '               执行过程：  ComExecuteProc Sel_XK_DB, strSQL, "保存"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    GetSampleValueForTJ = CreateLisValueStrForTJ(strAdvices)
    Call CreateLogFile("GetSampleValueForTJ=" & GetSampleValueForTJ)
End Function

Public Function SampleCheckinInfoWrite(strAdvices As String, strName As String, strBatchNO As Long, Optional strErr As String, Optional ByVal strSentName As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   签收时把接收人接收时间写入到LIS中
    '
    '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
    '                       strName      签收人
    '                       strBatchNO   批号
    '返回                   如果返回False时可在strErr中显示具体错误
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SampleCheckinInfoWrite = funSampleCheckinInfoWrite(strAdvices, strName, strBatchNO, strErr, strSentName)
End Function
Public Function SampleSendInfo(strAdvices As String, intType As Integer, ByVal strUser As String, Optional strErr As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   标本发送人发送时间写入LIS中
    '
    '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
    '                       intType      --0为送检 1为取消送检
    '返回                   如果返回False时可在strErr中显示具体错误
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SampleSendInfo = funSampleSendInfo(strAdvices, intType, strUser, strErr)
End Function

Public Function GetApplicationFormShowType() As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   用于取得检验申请单中的申请的显示类型
    '                       现有两种类型
    '                       1.只显示组合项目
    '                       2.显示组合项目中的明细项目
    '返回                   True = 显示明细项目 False = 只显示组合项目
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    GetApplicationFormShowType = ComGetPara(Sel_Lis_DB, "申请单显示组合明细", gSysInfo.SysNo, gSysInfo.ModlNo, 0)
End Function
Public Function WriteAdvicesLookState(strAdvices As String, intType As Integer) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   写入医嘱的查阅状态
    '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
    '                       intType      1=已查阅 0=未查阅
    '返回                   True=成功   False=失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    WriteAdvicesLookState = funWriteAdvicesLookState(strAdvices, intType)
    
End Function

Public Function ModifyApplyItemStateYJ(strAdvices As String, intType As Integer) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   改变医技执行时，申请项目状态
    '参数
    '                       lngAdviceID         医嘱id
    '                       intType      0=未核收 1=已核收 2=已审核 3=已执行（医技站直接执行）
    '返回                   True=成功   False=失败
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ModifyApplyItemStateYJ = funModifyApplyItemStateYJ(strAdvices, intType)
End Function

Public Function ModifyPathState(lngPartentID As Long, lngMainID As Long, lngPathSatae As Long, Optional strErr As String) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   写入临床路径的路径状态
    '参数                   lngPartentID   病人id
    '                       lngMainID      主页id
    '                       lngPathSatae   路径状态
    '                       写入
    '返回                   True=成功   False=失败
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ModifyPathState = funModifyPathState(lngPartentID, lngMainID, lngPathSatae)
End Function

Public Function GetLabNewReportList(ByVal lngPatientID As Long, ByVal lngMainID As Long, ByRef strXMLNewLIS As String, Optional lngApplyID As Long) As Boolean
    '功能               LIS的公共部件，调用生成XML格式的病人的检验报告列表
    '参数
    '                   lngPatientID            消息头
    '                   lngPatientID            消息内容
    '                   strXMLNewLIS            返回的字串
    '                   lngApplyID              申请id
    '返回               True=成功   False=失败
    GetLabNewReportList = funGetLabNewReportList(lngPatientID, lngMainID, strXMLNewLIS, lngApplyID)
End Function

Public Function GetLabNewReportResultList(ByVal lngRepottID As Long, ByRef strXMLNewLIS As String) As Boolean
    '功能：提取病人的检验报告结果
    '参数
    'lngRepottID            报告id
    '返回                   XML格式的字串
    GetLabNewReportResultList = funGetLabNewReportResultList(lngRepottID, strXMLNewLIS)
End Function

Public Function GetNewBloodBankItem(ByVal lngApplyID As Long, ByRef strXMLNewLIS As String) As Boolean
    '功能：提取病人的检验报告结果
    '参数
    'lngApplyID            相关id ,医嘱id
    '返回                   XML格式的字串
    GetNewBloodBankItem = funGetNewBloodBankItem(lngApplyID, strXMLNewLIS)
End Function

Public Function GetNewTransFusionApplyFor(strItemCodeing As String, lngPatientID As Long, intPatientType As Integer, lngHomePageID As Long, Optional strRegistrationBill As String, _
                                        Optional intBaby As Integer, Optional intType As Integer, Optional ByVal intDay As Integer) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                  根据传入医嘱ID返回结果
    '参数
    '                       strItemCodeing 诊疗项目编码（可传入多个，使用逗号分隔）
    '                       lngPatientID 病人ID
    '                       intPatientType 病人来源 1-门诊，2-住院
    '                       lngHomePageID 主页ID （病人来源=2时查询)
    '                       lngRegistrationBill 挂号单NO（病人来源<>2时查询本次就诊）
    '                       intBaby           是否婴儿
    '                       intType           那种方式，1=再此查7天内的。0 = 不查询 2=指定查询天数（intDay参数）其他= 暂定
    '                       intDay            当intType=2时，此参数才有效，指定要查询多少天的数据

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GetNewTransFusionApplyFor = funGetNewTransFusionApplyFor(strItemCodeing, lngPatientID, intPatientType, lngHomePageID, strRegistrationBill, intBaby, intType, intDay)
End Function



'*************************************************************************
'**函 数 名：GetMobileImages  返回图片名
'**输    入：ByVal lngApplyID(Long)     -   申请ID
'**        ：ByRef ObjDrawGraph(Object) -  LISDev对象
'**        ：ByVal strDir(String)       - 保存路径
'**        ：ByVal strErr(String)       - 错误信息
'**输    出：(String) -返回图像名称
'**功能描述：根据申请ID获取标本ID，并查询标本图像数据
'**日    期：2014-11-13 10:21:58
'**版    本：V10.35.00
'*************************************************************************
Public Function GetMobileImages(ByVal lngApplyID As Long, ByRef ObjDrawGraph As Object, ByVal strDir As String, ByRef strErr As String) As String
    Dim rsTmp  As ADODB.Recordset
    Dim lngSampleID   As Long
    Dim strReturn As String
    Dim strSQL As String
    '先初始化LISDev的对象
    Call ObjDrawGraph.GetSampleImgInit(gSysInfo.SysNo, gcnLisOracle, strErr)
    If strErr <> "" Then
        '初始化LISDEv对象错误，退出
        Exit Function
    End If
    strSQL = "select  标本id  from  检验申请组合 where  申请id=[1]"
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "查询标本ID", lngApplyID)
    If Not rsTmp.EOF And rsTmp.RecordCount = 1 Then
        lngSampleID = Val(Trim(rsTmp!标本ID & ""))
        strReturn = ObjDrawGraph.GetSampleImages(lngSampleID, strDir, False, strErr, 1, 1)
    End If
    If rsTmp.EOF Then
        strErr = "申请ID（" & lngApplyID & "）未找到标本ID"
    End If
    GetMobileImages = strReturn
End Function

Public Function openSqlOtherDB(ByVal selDB As Integer, ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset

    '功能，查询其他数据库，返回记录集
    On Error GoTo openSqlOtherDB_Error

    Dim lngCount As Long
    Dim var(30) As Variant
    
    lngCount = UBound(arrInput)
    If lngCount > 30 Then
        Err.Raise -2147483645, , "不支持超过30个参数的SQL！"
        Exit Function
    End If
    For lngCount = LBound(arrInput) To UBound(arrInput)
        var(lngCount) = arrInput(lngCount)
    Next
    
    Set openSqlOtherDB = OpenSQLRecord(selDB, strSQL, strTitle, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))

    Exit Function
openSqlOtherDB_Error:
    Call WriteErrLog("zlPublicHisCommLis", "clsLisInsideComm", "执行(openSqlOtherDB)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
    Err.Clear
End Function

Public Function ComExecuteOtherProc(ByVal selDB As Integer, strSQL As String, ByVal strFormCaption As String) As String
    '功能：执行过程语句,并自动对过程参数进行绑定变量处理
    '返回：无错误返回空串，否则返回错误提示
    

    Call ExecuteProcedure(selDB, strSQL, strFormCaption)
    
End Function

Public Function GetNewDatatoXK(ByVal lngPatientID As Long, ByVal strItemCode As String) As String
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                  根据传入病人ID，指标代码 返回当前病人最近一次的检验结果
    '参数
    '                       lngPatientID 病人id
    '                       strItemCode 指标代码串，使用 ，号分隔
    '                      返回格式：   病人id<A>指标代码1<S>检验结果1<B>指标代码2<S>检验结果2<B>指标代码3<S>检验结果3

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GetNewDatatoXK = funGetNewDataToXK(lngPatientID, strItemCode)
End Function



Public Function GetReadNotify(objFrm As Object, strAdvice As String, ByVal strDicName As String, Optional ByRef strReturn As String) As Boolean
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   医生工作站反馈危急值处理措施到LIS
    '参数
    '       入参
    '                       objfrm          调用窗体
    '                       strAdvice       医嘱id串
    '                       strDicName      医生姓名
    '       出参
    '                       strReturn       返回医生填写的处理措施
    '返回                   True=成功,False=失败
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GetReadNotify = funGetReadNotify(objFrm, strAdvice, strDicName, strReturn)
End Function

Public Function CanUseLISApp(ByVal strNO As String, ByRef strEx As String) As Boolean
          '功能：判断当前项目是否可以使用LIS申请单方式下达
          '参数：strNo=诊疗项目编码，strEx 扩展参数
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strErr As String

1         On Error GoTo CanUseLISApp_Error
2         If gUserInfo.NodeNo <> "-" Then
3             strSQL = "select  id from 检验组合项目 where  诊疗编码 =[1] and (站点=[2] or 站点 is null)"
4         Else
5             strSQL = "select  id from 检验组合项目 where  诊疗编码 =[1]"
6         End If
7         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "组合项目查询", strNO, gUserInfo.NodeNo)
8         If rsTmp.RecordCount > 0 Then
9             CanUseLISApp = True
10        Else
11            CanUseLISApp = False
12        End If

13        Exit Function
CanUseLISApp_Error:
14        strErr = Err.Number & " " & Err.Description
15        Call WriteErrLog("zlPublicHisCommLis", "clsLisInsideComm", "执行(CanUseLISApp)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
16        Err.Clear
End Function



Public Function GetTatTimeShow(ByVal intType As Integer, ByVal strItem As String, _
                                ByVal strDept As String, ByVal strGroup As String, _
                                ByVal strMachine As String, ByVal strSex As String, _
                                intMsg As Integer, Optional strShowBef As String, _
                                Optional lngTATTime As Long, Optional strUser As String) As String
    '功能       检查TAT是否超时
    
    '入参
    'intType            '1=送检,2=签收,3=核收,4=审核
    'strItem            '申请项目ID  项目ID1,项目名称1,上个时间节点1,急诊1;项目ID2,项目名称2,上个时间节点2,急诊2;
    'strDept            '申请科室ID
    'strGroup           '申请小组ID
    'strMachine         '申请仪器ID
    'strSex             '病人性别
    
    '出参
    'intMsg             '限制       1=只提示,2=提示并限制
    'strMsgShow         '提示信息
    
    'GetTatTimeShow=true表示超时,=false表示未超时
    GetTatTimeShow = getTATTime(intType, strItem, strDept, strGroup, strMachine, strSex, intMsg, strShowBef, lngTATTime, strUser)
    
End Function

Public Function GetLISPatientRecord(ByVal lngPatientID As Long, ByVal lngRecordID As Long, ByVal intType As Integer, Optional strErr As String) As ADODB.Recordset
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                       查询LIS病人信息
    '
    '入参
    '                           lngPatientID    病人ID
    '                           lngRecordID     intType=1 挂号单ID intType=2 主页ID intType=3 体检
    '                           intType         1=门诊 2=住院 3=体验
    '                           strErr          返回的错误信息
    '返回
    '                           病人信息记录集,记录集中包含以下病人信息
    '                           医嘱ID,相关ID,申请时间,病人来源,婴儿,姓名,性别,年龄,年龄数字,年龄单位,申请人, 申请科室,床号,病人科室,紧急,挂号单,门诊号,住院号,主页ID,标本类型,病区 from 检验申请组合
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set GetLISPatientRecord = funGetLISPatientRecord(lngPatientID, lngRecordID, intType, strErr)
End Function

Public Function ModifyPatientBaseintoLIS(ByVal lng病人ID As Long, ByVal str就诊ID As String, ByVal int场合 As Integer, ByVal strName As String, _
                                        ByVal strSex As String, ByVal lngAgeNum As Long, ByVal strAgeUnit As String, ByVal strEditMode As String, _
                                        ByVal strEditUser As String, Optional strErr As String) As Boolean
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                       同步修改到ZLLIS病人信息中
    '
    '入参
    '                           lng病人ID
    '                           str就诊ID       场合=1 挂号单 场合=2 主页ID 场合=3 体检
    '                           int场合         1=门诊 2=住院 3=体验
    '                           strName         要修改的病人姓名
    '                           strSex          要修改的病人性别
    '                           strAgeNum       要修改的病人年龄数字
    '                           strAgeUnit      要修改的病人年龄单位
    '
    '                           strEditMode     修改源于哪个模块
    '                           strEditUser     修改人
    '                           strErr          返回的错误信息
    '返回
    '                           True=保存成功 False=保存失败
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ModifyPatientBaseintoLIS = funModifyPatientBaseIntoLIS(lng病人ID, str就诊ID, int场合, strName, strSex, lngAgeNum, strAgeUnit, strEditMode, strEditUser, strErr)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/12/12
'功    能:获取医嘱TAT倒计时
'入    参:
'           strSQID     医嘱ID1,申请ID1;医嘱ID2,申请ID2......
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function GetTATAllTime(ByVal strSQIDs As String) As Long
          Dim lngTATTime As Long
          Dim lngTATTimeBefor As Long
          Dim strSQID() As String
          Dim i As Integer
          
1         On Error GoTo GetTATAllTime_Error

2         strSQID = Split(strSQIDs, ";")
3         For i = 0 To UBound(strSQID)
4             lngTATTimeBefor = GetTATAllTimefun(Split(strSQID(i), ",")(0), Split(strSQID(i), ",")(1))
5             If lngTATTime = 0 And lngTATTimeBefor <> 0 Then
6                 lngTATTime = lngTATTimeBefor
7             ElseIf lngTATTime > lngTATTimeBefor Then
8                 lngTATTime = lngTATTimeBefor
9             End If
10        Next
11        GetTATAllTime = lngTATTime


12        Exit Function
GetTATAllTime_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "clsLisInsideComm", "执行(GetTATAllTime)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
14        Err.Clear
          
End Function

Public Function CheckAcceptance(ByVal strAdviceID As String, Optional strErr As String) As Boolean
          '功能：门诊(医生)工作站 判断标本是否未核收
          
          '参数：strAdviceID--相关ID(多个相关ID以","隔开)
          
          '返回值：CheckAcceptance--True表示标本未核收
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo CheckAcceptance_Error

2         CheckAcceptance = True
          
3         strSQL = "Select /*+cardinality(b,10)*/ Nvl(a.申请状态, 0) 申请状态" & vbNewLine & _
                  "From 检验申请组合 A, Table(f_Str2list([1])) B" & vbNewLine & _
                  "Where a.申请id = b.Column_Value"
          
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "判断标本核收", strAdviceID)
          
5         Do While Not rsTmp.EOF
6             If rsTmp!申请状态 <> 0 Then CheckAcceptance = False: Exit Function
7             rsTmp.MoveNext
8         Loop
          
9         Exit Function
CheckAcceptance_Error:
10        strErr = "出错函数(CheckAcceptance),错误信息:" & Err.Number & " " & Err.Description
11        Call WriteErrLog("zlPublicHisCommLis", "clsLisInsideComm", "执行(CheckAcceptance)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
12        Err.Clear
End Function

Public Function ModifyBabyInfo(ByVal lngBID As Long, ByVal lngZID As Long, ByVal strNO As String, ByVal lngBabyID As Long, ByVal strBabyName As String, ByVal strBabySex As String, Optional strErr As String) As Boolean
    '功能        修改新生儿基本信息
    'lngBID      病人id
    'lngZID      主页id
    'strNO       挂号单
    'lngBabyID   婴儿序号
    'strBabyName 婴儿姓名
    'strBabyAge  婴儿性别
    '返回        True=修改成功,Flase=修改失败
    
    ModifyBabyInfo = funModifyBabyInfo(lngBID, lngZID, strNO, lngBabyID, strBabyName, strBabySex, strErr)
End Function

'---------------------------------------------------------------------------------------
' 编    码:蔡青松
' 编码日期:2017-3-23
' 功    能:写入检验通知记录
' 入    参:
'           intType                 1=检验科拒收标本，2=临床传染病处理反馈
'           strAdviceID             检验医嘱ID，多个使用","分割
'           [可选][intSampleType    标本状态    采集工作站调用时传入 1=让步检验，2=取消检验，3=样本重采]
' 出    参:
'          strErr                   错误信息
' 返    回:True=成功，False=失败
' 修 改 人:
' 修改日期:
'---------------------------------------------------------------------------------------
Public Function WriteInLisNotify(ByVal intType As Integer, ByVal strAdviceID As String, _
                                    Optional ByVal intSampleType As Integer, Optional strErr As String) As Boolean
    WriteInLisNotify = funWriteInLisNotify(intType, strAdviceID, intSampleType, strErr)
End Function

Public Function PrintReport(objFrm As Object, lngAdive As Long, lngPaint As Long, Optional byRunMode As Byte = 2, Optional strErr As String) As Boolean
'---------------------------------------------------------------------------------------
' 功    能:打印新版LIS报告
' 入    参:
'           objfrm               打印的主窗体
'           lngAdive             检验医嘱ID。
'           lngPaint             病人id
'           byRunMode            打印方式1-预览 ，2-打印
'           strErr               错误信息
'
' 出    参:
'          strErr                   错误信息
' 返    回:True=成功，False=失败
' 修 改 人:
' 修改日期:
'---------------------------------------------------------------------------------------
    PrintReport = PrintReportNew(objFrm, lngAdive, lngPaint, byRunMode, strErr)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/8/25
'功    能:  查询病人住院期间所有已出报告的申请ID，并且通过标本进行分组
'入    参:
'           lngPatientID    HIS病人ID
'           intPage         主页ID
'出    参:
'           strErr          错误信息
'返    回:  返回按照标本进行分组的医嘱ID串。医嘱之间用","分割，标本之间用";"分割
'---------------------------------------------------------------------------------------
Public Function GetPatientAdvice(ByVal lngPatientID As Long, ByVal intPage As Integer, Optional ByRef strErr As String) As String
    GetPatientAdvice = funGetPatientAdvice(lngPatientID, intPage, strErr)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/11/22
'功    能:回写危急值处理措施到LIS系统中
'入    参:
'           lngSampleID     标本ID
'           strUserName     处理人员
'           strNotify       处理措施
'出    参:
'           [strErr         错误信息]
'返    回:  回写成功返回True,否则返回False
'---------------------------------------------------------------------------------------
Public Function WriteNotifyToLis(ByVal lngSampleID As Long, ByVal strUserName As String, ByVal strNotify As String, Optional ByRef strErr As String) As Boolean
    WriteNotifyToLis = funWriteNotifyToLis(lngSampleID, strUserName, strNotify, strErr)
End Function

'---------------------------------------------------------------------------------------
'体检跟踪日志编写
'---------------------------------------------------------------------------------------

Private Function CreateLogFile(ByVal strValue As String)
    '生成目录,每天只生成一个目录
    Dim strFolderTime As String
    Dim strGetTime As String
    Dim strTime As String
    Dim strExePath As String
    Dim strLogPath As String
    
    
    '保存路径
    strExePath = Mid(App.Path, 1, InStr(5, App.Path, "\"))
    strLogPath = strExePath & "Log\日志跟踪\zl9LisInsideComm"
    
     '判断日志路径是否存在，若不存在，则创建
    If gobjFile.FolderExists(strExePath & "\Log") = False Then
        gobjFile.CreateFolder strExePath & "\Log"
    End If
    If gobjFile.FolderExists(strExePath & "\Log\日志跟踪") = False Then
        gobjFile.CreateFolder strExePath & "\Log\日志跟踪"
    End If
    
    If gobjFile.FolderExists(strLogPath) = False Then
        gobjFile.CreateFolder strLogPath
    End If
    
    strFolderTime = Format(Now, "yyyyMMdd")
    strGetTime = GetSetting("WriteLogTime", "value", "SaveTime", strFolderTime)
    
    strTime = Format(Now, "yyyyMMddHH")
    
    Set objTxt = gobjFile.OpenTextFile(strLogPath & "\WriteValToTJ" & strTime & ".log", ForAppending, True)
    objTxt.WriteLine "*" & DateTime.Now & "--------------------------------------------------"
    objTxt.WriteLine (strValue)
    objTxt.Close
End Function





'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/2
'功    能:提供给外部及程序来通过zl9LisInsideComm打开记录集
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function ComOutOpenSQL(ByVal selDB As Integer, ByVal strSQL As String, ByVal strTitle As String, _
    ParamArray arrInput() As Variant) As ADODB.Recordset
    
    Dim lngCount As Long
    Dim var(30) As Variant
    
    lngCount = UBound(arrInput)
    If lngCount > 30 Then
        Err.Raise -2147483645, , "不支持超过30个参数的SQL！"
        Exit Function
    End If
    For lngCount = LBound(arrInput) To UBound(arrInput)
        var(lngCount) = arrInput(lngCount)
    Next
    
    Set ComOutOpenSQL = ComOpenSQL(selDB, strSQL, strTitle, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/14
'功    能:  判断是否允许核收新门诊发送的医嘱
'入    参:
'           strAdvicIDs     医嘱ID，多个医嘱使用“,”分割
'出    参:
'返    回:  True=新门诊病人，False=非新门诊病人
'---------------------------------------------------------------------------------------
Public Function NewSystemSvr(ByVal strAdvicIDs As String) As Boolean
    NewSystemSvr = funNewSystemSvr(strAdvicIDs)
End Function

'---------------------------------------------------------------------------------------
'编    码:李小东
'编码时间:2018/03/21
'功    能:  获取检验报告记录审核时间
'入    参:
'           lngId     申请ID
'出    参:
'返    回:  审核时间
'---------------------------------------------------------------------------------------
Public Function GetNewTime(ByVal lngID As Long) As String
          Dim rsTime As Recordset
          
1     On Error GoTo GetNewTime_Error

2         gstrSQL = "Select Distinct b.审核时间 From 检验申请组合 A, 检验报告记录 B Where a.标本id = b.Id And a.申请id = [1]"
3         Set rsTime = ComOpenSQL(Sel_Lis_DB, gstrSQL, "查询审核时间", lngID)
4         If Not rsTime.EOF Then
5             GetNewTime = NVL(rsTime("审核时间"))
6         End If


7         Exit Function
GetNewTime_Error:
8         Call WriteErrLog("zlPublicHisCommLis", "clsLisInsideComm", "执行(GetNewTime)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
9         Err.Clear
End Function

Public Function GetSampleType(ByVal strAdvice As String, Optional strErr As String) As ADODB.Recordset
    '功能       获取标本状态
    
    '参数
    '           strAdvice       医嘱ID,多个医嘱用","分割
    '           strErr          返回错误信息
    
    '返回记录集
    '记录集字段: "医嘱ID", adBigInt
    '           "医嘱状态", adVarChar, 20
    '           "操作员", adVarChar, 20
    '           "操作时间", adDate
    
    Set GetSampleType = funGetSampleType(strAdvice, strErr)
    
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/18
'功    能:检查当前时间是否是业务高峰期，并且业务指定的查询范围是否在许可范围内(之所以要放在zl9LisInsideComm里，是方便LIS的其他部件进行调用)
'入    参:
'      lngSysNo=系统号
'      lngModuleNo=模块号
'      strFuncName=功能名称
'      datBegin=功能进行查询数据范围的开始时间，当类型为数值类型时
'      datEnd=功能进行查询数据范围的截至时间
'      lngDays=查询的时间范围,当为0时，通过datBegin与datEnd反算，不为0时忽略datBegin与datEnd
'出    参:
'返    回:是否可以进行操作
'---------------------------------------------------------------------------------------
Public Function CheckRushHours(ByVal lngSysNo As Long, ByVal lngModuleNo As Long, ByVal strFuncName As String, _
                                Optional ByVal datBegin As Date, Optional ByVal datEnd As Date, Optional ByVal lngDays As Long) As Boolean
    CheckRushHours = funCheckRushHours(lngSysNo, lngModuleNo, strFuncName, datBegin, datEnd, lngDays)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/25
'功    能:将拒收信息写入新版LIS
'入    参:
'           strAdviceID         申请ID 多个申请ID使用","分割
'           strUser             拒收人
'           strRefuseInfo       拒收理由
'           strRegName          拒收接收人
'           strRegTime          拒收接收时间
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function RefuseSampleInNew(ByVal strAdviceID As String, ByVal strUser As String, ByVal strRefuseInfo As String, _
                                Optional ByVal strRegName As String, Optional ByVal strRegTime As String, Optional ByRef strErr As String)
    RefuseSampleInNew = funRefuseSampleInNew(strAdviceID, strUser, strRefuseInfo, strRegName, strRegTime, strErr)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/7/5
'功    能:获取新版LIS中的标本类型
'入    参:
'           strInfo     不查询重复的标本类型
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function GetSampleTypeNew() As ADODB.Recordset
    Set GetSampleTypeNew = funGetSampleTypeNew
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/7/6
'功    能:通过医嘱ID打印LIS报告
'入    参:
'           objFrm          调用窗体
'           lngAdvice       医嘱ID（采集医嘱ID)
'           byRunMode       1=打印预览，2=打印，3=打印设置，4=打印PDF
'           BlnlimitPrint   打印新版报告时，是否受到打印次数参数的限制（超出新版LIS中打印次数的报告无法打印）
'           strPDF          需要打印的PDF文件的文件路径
'           strPrinter      指定打印机的名称，若果指定了打印机名称，则默认在指定的打印机上打印
'出    参:
'           strErr          返回错误或打印失败原因
'返    回:  是否打印成功    True=成功，False=失败
'---------------------------------------------------------------------------------------
Public Function PrintLisReport(ByVal objFrm As Object, ByVal lngAdvice As Long, ByVal byRunMode As Byte, _
                                    Optional ByVal BlnLimitPrint As Boolean, Optional ByVal strPDF As String, _
                                    Optional ByVal strPrinter As String, Optional ByRef strErr As String) As Boolean
    PrintLisReport = funPrintLisReport(objFrm, lngAdvice, byRunMode, BlnLimitPrint, strPDF, strPrinter, strErr)
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/10/24
'功    能：比较两个版本号,比VersionHIS版本号小，返回1，相等返回0，比VersionHIS版本号大返回-1
'参    数：strVerCur=当前版本号
'         strVerCom=对比的版本号
'返    回：对比版本号比当前版本号小，返回1，相等返回0，比当前版本号大返回-1
'调整影响:
'---------------------------------------------------------------------------------------
Public Function VerCompareVersionHIS(ByVal strVerCom As String) As Integer
      VerCompareVersionHIS = VerCompare(gSysInfo.VersionHIS, strVerCom)
End Function

'-----------------------------提供给外部调用的窗体接口BEGIN------------------------------
'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/2
'功    能:调用检验传染病报告列表
'入    参:
'           strUserID       用户ID
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub PrintDiseaseReport(ByVal strDBUser As String)
    Call frmDiseaseReportPrint.ShowMe(strDBUser)
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/5/3
'功    能:打开手工登记申请单
'入    参:
'           objFrm          调用来源
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub ShowLisHandApplicationForm(ByVal objFrm As Object)
    frmCollectionApply.Show 1, objFrm
End Sub

'---------------------------------------------------------------------------------------
'编    码:李小东
'编码时间:2017/8/14
'功    能:加载送检核对窗体
'入    参:
'           intType      1-检验采集工作站，2-标本签收
'           intDays      允许查询天数
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function ShowFrmSampleSendCheck(objFrm As Object, intType As Integer, intDays As Integer) As Boolean
    ShowFrmSampleSendCheck = frmSampleSendCheck.ShowMe(objFrm, intType, intDays)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/9
'功    能:浏览检验报告
'入    参:
'           objFrm              调用窗体
'           strAdviceID         主医嘱ID
'           intType             调用类型，0=单报告查阅，1=打印预览检验报告
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function LisSingleReprotBrowse(objFrm As Object, ByVal lngAdviceID As Long, ByVal intType As Integer, Optional strErr As String) As Boolean
    If lngAdviceID > 0 Then
        LisSingleReprotBrowse = frmAdviceReprotBrowse.ShowMe(objFrm, lngAdviceID, intType)
    Else
        strErr = "传入医嘱ID错误"
        Exit Function
    End If
End Function

Public Function ShowPatientSampleBrowse(objFrm As Object, lngPatientID As Long, strPrivs As String, lngDept As Long, lngDeptDistrict As Long, _
                intPatientType As Integer, Optional strErr As String) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           HIS对LIS报告单的浏览
    '参数           objfrm              父对象
    '               lngPatientID        病人ID
    '               strPrivs            父权限
    '               lngDept             科室ID
    '               lngDeptDistrict     病区ID
    '               strErr              有错误号后显示具体的错误信息
    '返回           成功返回 true 失败返回 false
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ShowPatientSampleBrowse = frmPatientReprotBrowse.ShowMe(objFrm, lngPatientID, strPrivs, lngDept, lngDeptDistrict, intPatientType, 0, strErr)
End Function

Public Function SystemCommSetup(objFrm As Object) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能            用于设置内部接口间的连接设置
    '参数            父窗体对象
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    frmCommSetup.Show vbModal, objfrm
'    Call frmPatientReprotBrowse.ShowMe(objfrm, 0, "所有科室", 0, 0)
End Function

Public Function PatientSampleBrowse(objFrm As Object, lngPatientID As Long, strPrivs As String, lngDept As Long, lngDeptDistrict As Long, _
                intPatientType As Integer, Optional lngPatientPage As Long, Optional strErr As String, Optional ByVal blnShowBorder As Boolean = True, _
                Optional ByRef objOutFrm As Object) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           打开LIS检验浏览器
    '参数           objfrm              父窗体对象
    '               lngPatientID        病人ID
    '               strPrivs            父窗体的权限
    '               lngDept             科室ID
    '               lngDeptDistrict     病区ID
    '               intPatientType      病人来源
    '               strErr              如果有错误时返回的错误信息
    '               blnShowBorder       是否显示窗体的边框，若不显示则表示该窗体为嵌入式调用，不需要调用窗体的show方法
    '               lngHwnd             嵌入式调用时被嵌入窗体的句柄
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    PatientSampleBrowse = frmPatientReprotBrowse.ShowMe(objFrm, lngPatientID, strPrivs, lngDept, lngDeptDistrict, intPatientType, lngPatientPage, strErr, blnShowBorder, objOutFrm)
End Function


Public Function ShowLisApplicationForm(frmMain As Object, lngModifyAppForNO As Long, lngPatientID As Long, intBaby As Byte, _
                                        lngPatientPage As Long, _
                                        strPatientName As String, strPatientSex As String, _
                                        strPatientAge As String, intPatientType As Integer, strOutPatientsNO As String, _
                                        strInPatientsNO As String, lngPhysicalExamination As Double, strDiagnose As String, _
                                        strAppForMan As String, lngAppForDeptID As Long, strAppForDeptName As String, _
                                        lngPatientDeptID As Long, strPatientDeptName As String, _
                                        Optional blnCancel As Boolean, Optional strErr As String, Optional strInData As String, _
                                        Optional strItemCode As String) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能           医生站调用检验申请单
          '参数           frmMain                         父窗体
          '               lngModifyAppForNO               传入申请单号,0=新增，>0修改
          '               strAdvice                       修改申请时传入申请单对应的医嘱ID串 格式:'采集医嘱1,检验医嘱1,采集医嘱2,检验医嘱2,...
          '               lngPatientPage                  病人ID
          '               intBaby                         婴儿
          '               intPatientPage                  病案主页
          '               strPatientName                  病人姓名
          '               strPatientSex                   病人性别
          '               strPatientAge                   病人年龄
          '               intPatientType                  病人来源
          '               strOutPatientsNO                门诊号
          '               strInPatientsNO                 住院号
          '               lngPhysicalExamination          健康号(体检号)
          '               strDiagnose                     诊断(最后一次诊断)
          '               strAppForMan                    申请人
          '               lngAppForDeptID                 申请科室ID
          '               strAppForDeptName               申请科室名
          '               lngPatientDeptID                病人科室ID
          '               strPatientDeptName              病人科室名
          '               blnCancel                       返回True=按下了取消按钮，False=按下了申请 如果申请了返回结果为空时也返回为=True
          '               strErr                          有错误信息时返回错误信息
          '返回           选择的申请信息
          '               格式:<采诊科室ID1,执行科室ID1,申请时间1,诊疗项目ID1,标本1;采诊科室ID2,执行科室ID2,申请时间2,诊疗项目ID2,标本2;.....>
          '
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '====================================================================================================
          '测试时使用
      '    lngPatientID = 1
      '    lngPatientPage = 1
      '    intBaby = 1
      '    strPatientName = "测试1"
      '    strPatientSex = "男"
      '    strPatientAge = "33岁"
      '    intPatientType = 2
      '    strOutPatientsNO = 0
      '    strInPatientsNO = "323333"
      '    lngPhysicalExamination = 0
      '    strDiagnose = "HIV阳性"
      '    strAppForMan = "张医师"
      '    lngAppForDeptID = 41
      '    strAppForDeptName = "二病区"
      '    lngPatientDeptID = 23
      '    strPatientDeptName = "内科一室"
          '====================================================================================================
1         On Error GoTo ShowLisApplicationForm_Error

2         ShowLisApplicationForm = frmAppforBill.ShowMe(frmMain, lngModifyAppForNO, lngPatientID, intBaby, lngPatientPage, strPatientName, strPatientSex, strPatientAge, _
                                      intPatientType, strOutPatientsNO, strInPatientsNO, lngPhysicalExamination, strDiagnose, strAppForMan, lngAppForDeptID, _
                                      strAppForDeptName, lngPatientDeptID, strPatientDeptName, blnCancel, strErr, strInData, strItemCode)


3         Exit Function
ShowLisApplicationForm_Error:
4         strErr = Err.Description
5         Call WriteErrLog("zlPublicHisCommLis", "clsLisInsideComm", "执行(ShowLisApplicationForm)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
6         Err.Clear
          
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-03-13
'功    能:  通过诊疗编码获取当前项目是否为耐受实验项目
'入    参:
'           strItemID       诊疗项目ID
'出    参:
'           strErr          错误或者提示信息
'返    回:  True=当前项目是耐受项目，False=当前项目不是耐受项目
'调整影响:
'---------------------------------------------------------------------------------------
Public Function IsToleranceItem(ByVal strItemID As Long, ByRef strErr As String) As Boolean
    IsToleranceItem = funIsToleranceItem(strItemID, strErr)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-04-15
'功    能:  通过传入医嘱ID串，返回哪些医嘱是属于同一个标本
'入    参:
'           strAdvice       医嘱ID，多个医嘱ID使用英文逗号分割
'出    参:
'返    回:  医嘱ID串，不同标本之间的医嘱ID使用英文分号分割，相同标本的医嘱ID使用逗号分割
'调整影响:
'---------------------------------------------------------------------------------------
Public Function GetSampleAdvice(ByVal strAdvice As String) As String
    GetSampleAdvice = funGetSampleAdvice(strAdvice)
End Function

'---------------------------------------------------------------------------------------
'编    码:李小东
'编码时间:2019-05-06
'功    能:打开让步检验理由填写窗体
'入    参:
'         objForm       父窗体
'         strIDs        医嘱ID，多个医嘱ID使用英文逗号分割
'         strUser       操作人
'出    参:
'返    回:是否保存成功
'调整影响:
'---------------------------------------------------------------------------------------
Public Function ShowTestReason(ByVal objForm As Object, ByVal strIDs As String, ByVal strUser As String) As Boolean
    Dim intValue As Integer
    
    If VerCompare(gSysInfo.VersionLIS, "10.35.140") = -1 Then
        ShowTestReason = True
    End If
    
    intValue = Val(ComGetPara(Sel_Lis_DB, "让步检验时填写让步理由", 2500, 2500, 0))
    If intValue = 1 Then
        ShowTestReason = frmTestReason.ShowMe(objForm, strIDs, strUser)
    Else
        ShowTestReason = True
    End If
End Function

'----------------------------提供给外部调用的窗体接口END-----------------------------------





'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-09-26
'功    能:  通过诊疗ID获取组合项目明细
'入    参:
'           strInfo         intType=0,组合项目对应的诊疗项目的ID，多个使用“,”分割;intType=1,诊疗编码，多个使用逗号分割
'出    参:
'返    回:  组成组合项目的指标记录集
'调整影响:
'调用注意:  intType=0时，返回的记录集有4个字段:指标ID，指标名称，耐受方案ID，耐受时间
'           intType=1时，返回的记录集有8个字段:组合名称，组合编码，诊疗编码，标本类型，指标ID，指标名称，耐受方案ID，耐受时间
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Public Function GetGroupItemInfo(ByVal strInfo As String, Optional ByVal intType As Integer) As ADODB.Recordset
    Set GetGroupItemInfo = funGetGroupItemInfo(strInfo, intType)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-10-29
'功    能:  通过医嘱ID病人报告（新版LIS和老版LIS）
'入    参:
'           strAdviceID     医嘱ID串，多个使用逗号分割
'出    参:
'           strErr          错误信息或者提示信息
'返    回:  包含传入医嘱ID对应的所有检验指标及结果，记录集内容：中文名,检验结果,单位

'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function GetPatientReport(ByVal strAdviceID As String, Optional ByRef strErr As String) As ADODB.Recordset
    Set GetPatientReport = funGetPatientReport(strAdviceID, strErr)
End Function

Public Function GetPersonTwo(ByVal strCode As String) As ADODB.Recordset
      '获取病人信息
          Dim rsTemp As Recordset
          Dim strSQL As String

1         On Error GoTo GetPersonTwo_Error

2         strSQL = "Select a.id, a.姓名, b.用户名" & vbNewLine & _
                   "From 人员表 A, 上机人员表 B" & vbNewLine & _
                   "Where a.Id = b.人员id   And ( a.姓名 like [1]  or a.编号 like [1] or a.简码 like [1])  "

3         Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "人员信息", UCase(strCode) & "%")
4         Set GetPersonTwo = rsTemp


5         Exit Function
GetPersonTwo_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "clsLisInsideComm", "执行(GetPersonTwo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear

End Function

Private Sub Class_Terminate()
    If Not (mobjITranLib Is Nothing) Then Set mobjITranLib = Nothing
End Sub


