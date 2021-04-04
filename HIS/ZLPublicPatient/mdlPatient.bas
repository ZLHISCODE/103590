Attribute VB_Name = "mdlPatient"
Option Explicit

Public gobjSquare As SquareCard  '卡结算部件

Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
End Sub

Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

Public Function GetPatiColor(ByVal strPatiType As String, Optional ByVal lngColor As Long = 0) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人颜色
    '入参:strPatiType:病人类型
    '返回:病人颜色
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strPatiType <> "" Then
        GetPatiColor = gobjDatabase.GetPatiColor(strPatiType)
    Else
        GetPatiColor = lngColor
    End If
End Function

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthday As String = "", Optional ByVal bytTag As Byte = 0, Optional ByVal strCalcDate As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人年龄检查
    '入参:strAge:病人年龄
    '     strBirthDay:出生日期
    '     bytTag:对于zl_Age_Check函数返回的询问类型的信息，是否要强制终止，还是保持询问.0-保持询问,1-禁止
    '     strCalcDate:计算日期,默认计算日期为系统日期.
    '返回：TRUE或FALSE，TRUE:继续,FALSE:终止
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfo As String, lngTmp As Long
    
    On Error GoTo ErrHand
    strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthday) Then
        If strCalcDate = "" Then
            strSQL = "select Zl_Age_Check([1],[2]) From dual"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge, CDate(strBirthday))
        Else
            strCalcDate = Format(strCalcDate, "YYYY-MM-DD HH:mm")
            strSQL = "select Zl_Age_Check([1],[2],[3]) From dual"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge, CDate(strBirthday), CDate(strCalcDate))
        End If
    Else
        strSQL = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge)
    End If
    strInfo = Nvl(rsTemp.Fields(0).Value)
    If InStr(1, strInfo, "|") > 0 Then
        lngTmp = Val(Split(strInfo, "|")(0)) '1禁止,0提示
        strInfo = Split(strInfo, "|")(1)
        If lngTmp = 1 Or (lngTmp = 0 And bytTag = 1) Then
            MsgBox strInfo & vbCrLf & vbCrLf & "请重新输入年龄!", vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox(strInfo & vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    CheckAge = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckIdcard(ByVal strIdcard As String, Optional strBirthday As String, Optional strAge As String, Optional strSex As String, _
    Optional strErrInfo As String, Optional datCalc As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：身份证号码合法性校验
    '入参：strIdCard 身份证号码
    '出参：strBirthday  函数返回True为出生日期
    '         strSex 函数返回True为性别
    '         strErrInfo 函数返回False为错误信息
    '         datCalc 计算日期 缺省则按系统时间计算
    '返回：True/False  身份证合法返回True(可从strBirthday，strSex获取出生日期和性别)，否则返回False(可从strErrInfo获取详细错误信息)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim xmlDoc As DOMDocument
    Dim xmlRoot As IXMLDOMElement
    
    On Error GoTo errH
     '检查身份证号是否合法
    '--<OUTPUT>
    '--       <BIRTHDAY></BIRTHDAY> //出生日期
    '--       <SEX></SEX>           //性别
    '--       <AGE></AGE>          //年龄
    '--     <MSG></MSG>         //身份证合法返回空(可从身份证号中获取出生日期和性别)，否则返回错误信息
    '--</OUTPUT>
    If datCalc = CDate(0) Then
        strSQL = "Select Zl_Fun_Checkidcard([1]) As Info From Dual"
    Else
        strSQL = "Select Zl_Fun_Checkidcard([1],[2]) As Info From Dual"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Fun_Checkidcard", strIdcard, datCalc)
    strXML = Trim(Nvl(rsTmp!Info))
    If strXML = "" Then Exit Function
    
    Set xmlDoc = New DOMDocument
    xmlDoc.loadXML (strXML)
    '读取XML内容
    Set xmlRoot = xmlDoc.selectSingleNode("OUTPUT")
    strErrInfo = xmlRoot.selectSingleNode("MSG").Text
    If strErrInfo <> "" Then Exit Function
    
    strBirthday = xmlRoot.selectSingleNode("BIRTHDAY").Text
    strSex = xmlRoot.selectSingleNode("SEX").Text
    strAge = xmlRoot.selectSingleNode("AGE").Text
    
    CheckIdcard = True
    Exit Function
errH:
 If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function SaveBaseInfo(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal strName As String, ByVal strSex As String, _
    ByVal strAge As String, ByVal strBirthday As String, ByVal str模块 As String, Optional ByVal int场合 As Integer = 1, Optional strInfo As String = "", _
    Optional ByVal blnXWHIS As Boolean, Optional ByVal blnEMPI As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：调整病人基本信息(含业务数据的同步调整)
    '入参：lng病人ID-病人ID (不能为空/0)
    '         lng就诊ID-挂号ID或主页ID(可为0)
    '         strName-姓名 (不能为空)
    '         strSex-性别 (不能为空)
    '         strAge-年龄 (不能为空)
    '         strBirthDay-出生日期 (不能为空)
    '         str模块-调用该功能的模块描述，如"门诊挂号"，"检查报到"。
    '         int场合 1-门诊;2-住院(lng就诊ID=0,则默认为1;lng就诊ID<>0,1-lng就诊ID为挂号ID,2-lng就诊ID为主页ID)
    '         strInfo-病人基本信息调整传入修改原因
    '         blnXWHIS-基本信息调整时是否调用RIS的接口 =True调用（该参数用于避免病人信息中重复调用RIS接口）
    '         blnEMPI-T EMPI平台已经建档，F-EMPI平台未建档
    ' 出参：strInfo:更新成功-信息调整导致的变化信息(返回True); 更新失败-信息调整未成功的原因
    ' 返回：TRUE OR False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    Dim strSQL As String, strSQLProc As String
    Dim blnTrans As Boolean
    Dim lngAgeNum As Long, strAgeUnit As String
    Dim str挂号NO As String
    Dim strErr As String, strTip As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRet As Long
    
    strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
    Set cmdTmp = New ADODB.Command
    strSQLProc = "Zl_病人信息_基本信息调整("
'   病人id_In 病人信息变动.病人id%Type,
    strSQLProc = strSQLProc & "" & lng病人ID & ","
    Set cmdPara = cmdTmp.CreateParameter("病人ID", adVarNumeric, adParamInput, 18, lng病人ID)
    cmdTmp.Parameters.Append cmdPara
'   就诊id_In Number := Null,
    strSQLProc = strSQLProc & "" & lng就诊ID & ","
    Set cmdPara = cmdTmp.CreateParameter("就诊ID", adVarNumeric, adParamInput, 18, lng就诊ID)
    cmdTmp.Parameters.Append cmdPara
'   模块_In   病人信息变动.变动模块%Type,
    strSQLProc = strSQLProc & "'" & str模块 & "',"
    Set cmdPara = cmdTmp.CreateParameter("变动模块", adVarChar, adParamInput, 100, str模块)
    cmdTmp.Parameters.Append cmdPara
'   姓名_In   病人信息.姓名%Type,
    strSQLProc = strSQLProc & "'" & strName & "',"
    Set cmdPara = cmdTmp.CreateParameter("姓名", adVarChar, adParamInput, 100, strName)
    cmdTmp.Parameters.Append cmdPara
'   性别_In   病人信息.性别%Type,
    strSQLProc = strSQLProc & "'" & strSex & "',"
    Set cmdPara = cmdTmp.CreateParameter("性别", adVarChar, adParamInput, 100, strSex)
    cmdTmp.Parameters.Append cmdPara
'   年龄_In   病人信息.年龄%Type
    strSQLProc = strSQLProc & "'" & strAge & "',"
    Set cmdPara = cmdTmp.CreateParameter("年龄", adVarChar, adParamInput, 100, strAge)
    cmdTmp.Parameters.Append cmdPara
'   出生日期_In 病人信息.出生日期%Type,
    strSQLProc = strSQLProc & "" & "TO_Date('" & strBirthday & "','YYYY-MM-DD HH24:mi')" & ","
    If Not IsDate(strBirthday) Then
        Set cmdPara = cmdTmp.CreateParameter("出生日期", adVarChar, adParamInput, 18, strBirthday)
    Else
        Set cmdPara = cmdTmp.CreateParameter("出生日期", adDBTimeStamp, adParamInput, , CDate(strBirthday))
    End If
    cmdTmp.Parameters.Append cmdPara
'   场合_In   number(1)  --1-门诊;2-住院
    strSQLProc = strSQLProc & "" & int场合 & ","
    Set cmdPara = cmdTmp.CreateParameter("场合", adVarNumeric, adParamInput, 1, int场合)
    cmdTmp.Parameters.Append cmdPara
    '修改原因_IN    varchar2
    strSQLProc = strSQLProc & "'" & strInfo & "',"
    Set cmdPara = cmdTmp.CreateParameter("修改原因", adVarChar, adParamInput, 100, strInfo)
    cmdTmp.Parameters.Append cmdPara
'   说明_Out    Out 病人信息变动.说明%Type --出参
    strSQLProc = strSQLProc & "" & "" & ")"
    Set cmdPara = cmdTmp.CreateParameter("说明", adLongVarChar, adParamOutput, 4000)
    cmdTmp.Parameters.Append cmdPara
    cmdTmp.ActiveConnection = gcnOracle
    cmdTmp.CommandType = adCmdStoredProc
    cmdTmp.CommandText = "Zl_病人信息_基本信息调整"
    
    strInfo = ""
    On Error GoTo errH
    
    'LIS 部件之所以在这初始化，是因为创建LIS部件后，要调用LIS部件初始化接口，里面有可能会弹出消息框，所以放在事务之前。
    Call InitObjLis
    If Not gobjLIS Is Nothing Then
        lngAgeNum = CLng(Val(Trim(strAge)))
        strAgeUnit = Mid(strAge, InStr(strAge, CStr(lngAgeNum)) + Len(CStr(lngAgeNum)))
        If lng就诊ID <> 0 And int场合 = 1 Then
            strSQL = "select NO from 病人挂号记录 where 病人ID = [1] and ID = [2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "读取病人挂号单NO", lng病人ID, lng就诊ID)
            If rsTmp.RecordCount > 0 Then
                str挂号NO = rsTmp!NO & ""
            End If
        End If
    End If
    Call CreatePlugInOK(glngModule)
    If (Not gobjLIS Is Nothing) Or (Not gobjPlugIn Is Nothing) Or blnXWHIS Then gcnOracle.BeginTrans: blnTrans = True
    
    Call gobjComlib.SQLTest(App.ProductName, "Zl_病人信息_基本信息调整", strSQLProc)
    cmdTmp.Execute
    Call gobjComlib.SQLTest
    '90816,修改新版LIS数据
    If Not gobjLIS Is Nothing Then
        If lng就诊ID <> 0 And int场合 = 1 Then
            If Not gobjLIS.ModifyPatientBaseintoLIS(lng病人ID, str挂号NO, int场合, strName, strSex, lngAgeNum, strAgeUnit, str模块, UserInfo.姓名, strInfo) Then
                gcnOracle.RollbackTrans: blnTrans = False
                strInfo = "LIS 系统病人信息修改失败，错误：" & strInfo
                Exit Function
            Else
                If strInfo <> "" Then strTip = strInfo  '成功后的提示
            End If
        Else
            If Not gobjLIS.ModifyPatientBaseintoLIS(lng病人ID, CStr(lng就诊ID), int场合, strName, strSex, lngAgeNum, strAgeUnit, str模块, UserInfo.姓名, strInfo) Then
                gcnOracle.RollbackTrans: blnTrans = False
                strInfo = "LIS 系统病人信息修改失败，错误：" & strInfo
                Exit Function
            Else
                If strInfo <> "" Then strTip = strInfo
            End If
        End If
    End If
    'EMPI
    If Not gobjPlugIn Is Nothing Then
        If blnEMPI Then
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModule, lng病人ID, IIf(int场合 = 2, lng就诊ID, 0), IIf(int场合 = 1, lng就诊ID, 0), strInfo)  '1=成功;0-失败
            Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo", strErr)
            If Err.Number = 438 Then lngRet = 1
            Err.Clear: On Error GoTo 0
        Else
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModule, lng病人ID, IIf(int场合 = 2, lng就诊ID, 0), IIf(int场合 = 1, lng就诊ID, 0), strInfo)  '1=成功;0-失败
            Call zlPlugInErrH(Err, "EMPI_AddPatiInfo", strErr)
            If Err.Number = 438 Then lngRet = 1
            Err.Clear: On Error GoTo 0
        End If
        If strErr <> "" Or lngRet = 0 Then
            gcnOracle.RollbackTrans
            strInfo = IIf(blnEMPI, "向EMPI平台更新病人信息失败！", "向EMPI平台新增病人信息失败！") & vbCrLf & IIf(strErr <> "", strErr, strInfo)
            Exit Function
        End If
    End If
    If blnTrans Then gcnOracle.CommitTrans: blnTrans = False
    
    If blnXWHIS Then
        'RIS 118004
        If CreateXWHIS() Then
            If gobjXWHIS.HISModPati(int场合, lng病人ID, lng就诊ID) <> 1 Then
                strTip = strTip & "当前启用了影像信息系统接口，但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。"
            End If
        ElseIf gblnXW = True Then
            strTip = strTip & "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。"
        End If
    End If
    
    strInfo = Trim(Nvl(cmdTmp.Parameters("说明"), "")) & IIf(strTip <> "", vbCrLf & strTip, "")
    SaveBaseInfo = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function Get病人信息从表(ByVal lng病人ID As Long, Optional ByVal str信息串 As String = "") As ADODB.Recordset
'功能：
'    获取病人信息从表项
'参数:
    Dim strSQL As String
    Dim intRet As Integer
    
    intRet = UBound(Split(str信息串, ","))
    If intRet = -1 Then '读取病人所有从表信息
        strSQL = "Select 信息名,信息值 From 病人信息从表 Where 病人ID =[1] And 信息值 is Not Null"
    ElseIf intRet = 0 Then '读取指定某个从表信息
        strSQL = "Select 信息名,信息值 From 病人信息从表 Where 病人ID =[1] And 信息名='" & Split(str信息串, ",")(0) & "'" & " And 信息值 is Not Null "
    ElseIf intRet > 0 Then '读取指定的多个从表信息值
        strSQL = "Select 信息名, 信息值" & vbNewLine & _
            "From 病人信息从表" & vbNewLine & _
            "Where 病人id = [1] And" & vbNewLine & _
            "      信息名 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) And 信息值 is Not Null "
    End If
    
    On Error GoTo errH
    Set Get病人信息从表 = gobjDatabase.OpenSQLRecord(strSQL, "读取病人从表", lng病人ID, str信息串)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function CreateXWHIS(Optional ByVal blnMsg As Boolean) As Boolean
'功能：判断 RIS接口部件(zl9XWInterface.clsHISInner) 是否存在，并启用
'参数：blnMsg－创建失败时是否提示

    If Not gblnXW Then Exit Function
    If Not gobjXWHIS Is Nothing Then CreateXWHIS = True: Exit Function
    
    On Error Resume Next
    Set gobjXWHIS = GetObject(, "zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    On Error Resume Next
    If gobjXWHIS Is Nothing Then Set gobjXWHIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    If gobjXWHIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS接口部件(zl9XWInterface)未创建成功！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CreateXWHIS = True
End Function

Public Function InitObjLis(Optional ByVal blnMsg As Boolean) As Boolean
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = GetObject(, "zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
    
        On Error Resume Next
        If gobjLIS Is Nothing Then Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
        
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If blnMsg Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
                Exit Function
            End If
        End If
    End If
    InitObjLis = True
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject(, "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
    
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String, Optional ByRef strErr As String = "0")
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    Dim strMsg As String
    
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        strMsg = "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description
        If strErr = "0" Then
            MsgBox strMsg, vbInformation, gstrSysName
        Else
            strErr = strMsg
        End If
    End If
End Sub
